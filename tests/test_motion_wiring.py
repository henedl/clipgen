"""Static wiring checks for the shared ClipgenMotion animation module (motion.js).

Locks the contract that makes the exit/entry animations work across Studio and
Screenspace: the module exists and exports its API, both pages load it after
utils.js, every mutation site calls it, and the old CSS landing animation stays
retired (so a refactor can't silently drop the script tag or resurrect two
parallel animation systems).
"""

from pathlib import Path

_WEB = Path(__file__).resolve().parent.parent / "assets" / "web"
MOTION_JS = _WEB / "motion.js"


def test_motion_module_exists_and_exports_api():
    assert MOTION_JS.is_file(), "assets/web/motion.js is missing"
    src = MOTION_JS.read_text(encoding="utf-8")
    assert "global.ClipgenMotion = {" in src
    for fn in ("animateOut", "animateOutAll", "animateIn", "flyTo"):
        assert "function " + fn + "(" in src, f"motion.js should define {fn}()"
    assert "var PARAMS = {" in src, "motion.js should expose the tweakable PARAMS"


def test_motion_is_size_aware():
    # Motion scales by element size so big cards calm down while small pills stay
    # lively — one animation definition, differentiated by size.
    src = MOTION_JS.read_text(encoding="utf-8")
    assert "function measureSizeScale(" in src
    assert "getBoundingClientRect" in src
    assert "sizeAware: true" in src  # stash opts in
    assert "size: {" in src  # the size-awareness config block


def test_motion_wiggle_is_tunable():
    # The chime is tunable: wiggleSpan sets how much of the duration it occupies,
    # and wiggleCycles: 0 genuinely disables it (no forced minimum of 1 swing).
    src = MOTION_JS.read_text(encoding="utf-8")
    assert "wiggleSpan" in src
    assert "Math.max(0, p.wiggleCycles" in src


def test_motion_loaded_after_utils_on_both_pages():
    for page in ("screenspace.html", "studio.html"):
        html = (_WEB / page).read_text(encoding="utf-8")
        assert '<script src="motion.js"></script>' in html, (
            f"{page} must load motion.js"
        )
        # motion.js provides window.ClipgenMotion; it must load after utils.js and
        # before the page hub/satellite scripts that call it.
        assert html.index('src="utils.js"') < html.index('src="motion.js"'), (
            f"{page}: motion.js must load after utils.js"
        )


def test_motion_wired_at_mutation_sites():
    studio = (_WEB / "studio.js").read_text(encoding="utf-8")
    ss_overlay = (_WEB / "screenspace-overlay-interaction.js").read_text(
        encoding="utf-8"
    )
    ss_hub = (_WEB / "screenspace.js").read_text(encoding="utf-8")

    # Exit animations (stash + delete) are wired in both tools.
    assert 'ClipgenMotion.animateOut(card, "delete")' in studio  # remove one card
    assert 'ClipgenMotion.animateOutAll(cards, "delete")' in studio  # clear queue
    assert 'ClipgenMotion.animateOutAll(cards, "stash")' in studio  # stash queue
    assert 'ClipgenMotion.animateOut(chip, "delete")' in ss_overlay  # delete region
    assert 'ClipgenMotion.animateOutAll(chips, "delete")' in ss_overlay  # delete all
    assert 'ClipgenMotion.animateOutAll(chips, "stash")' in ss_hub  # stash regions

    # Stash-card landing goes through the shared system on both tools; region
    # pills reuse the same entry animation.
    assert 'ClipgenMotion.animateIn(card, "stashLand")' in studio
    assert 'ClipgenMotion.animateIn(card, "stashLand")' in ss_hub
    assert 'ClipgenMotion.animateIn(chip, "stashLand")' in ss_overlay


def test_old_css_stash_landing_removed():
    # The CSS keyframe entry animation was migrated onto ClipgenMotion.animateIn;
    # leaving it behind would resurrect a parallel second animation system.
    css = (_WEB / "studio.css").read_text(encoding="utf-8")
    assert "stash-card-land" not in css
    assert "stash-card-landed" not in css
