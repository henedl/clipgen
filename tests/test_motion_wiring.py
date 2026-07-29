"""Static wiring checks for the shared ClipgenMotion animation module (motion.js).

Locks the contract that makes the exit/entry animations work: the module exists
and exports its API, every page loads it after utils.js, every mutation site
calls it, and the CSS animations it replaced stay retired — so a refactor can't
silently drop the script tag or resurrect two parallel animation systems.

The last group covers the other half of that boundary. Continuous loops (the
spinner rotation, the "working" opacity breathe) deliberately stay in CSS, so
they get deduped by hoisting into tokens.css rather than by moving onto the
engine; these tests keep both halves from regrowing per-page copies.
"""

import re

from _frontend_source import WEB as _WEB
from _frontend_source import concat_js

MOTION_JS = _WEB / "motion.js"

# Same list as tests/test_hotkeys_frontend_source.py — every page ships the shared
# engine, and the two exportable templates get it inlined by source/viewer.py.
ALL_TEMPLATES = (
    "studio.html",
    "screenspace.html",
    "transcripts.html",
    "composer.html",
    "overview.html",
    "workflows.html",
    "viewer.html",
    "gallery.html",
    "timeline-viewer.html",
)


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


def test_motion_loaded_after_utils_on_every_page():
    # Every template loads the engine, not just the two that animate cards: the
    # shared showToast fade lives in utils.js and is guarded on
    # window.ClipgenMotion, so a page missing the script snaps its toasts.
    for page in ALL_TEMPLATES:
        html = (_WEB / page).read_text(encoding="utf-8")
        assert '<script src="motion.js" defer></script>' in html, (
            f"{page} must load motion.js"
        )
        # motion.js provides window.ClipgenMotion; it must load after utils.js and
        # before the page hub/satellite scripts that call it.
        assert html.index('src="utils.js"') < html.index('src="motion.js"'), (
            f"{page}: motion.js must load after utils.js"
        )


def test_motion_wired_at_mutation_sites():
    # Studio's mutation sites are spread across the hub + its satellites (the
    # stash sites live in studio-stash.js), so read the whole studio*.js group —
    # same convention as test_studio_frontend_source.py.
    studio = concat_js("studio")
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


def test_motion_has_generic_fade_and_pop_kinds():
    # Beyond stash/delete/stashLand, the engine exposes two reusable kinds:
    # `fade` (opacity only) and `pop` (translateY + scale + opacity, mirroring the
    # studio.css cg-overlay-pop entrance). Both are registered in PARAMS and
    # dispatched by animateIn (entry) and animateOut (exit).
    src = MOTION_JS.read_text(encoding="utf-8")
    assert "function buildFadeKeyframes(" in src
    assert "function buildPopKeyframes(" in src
    assert "pop: {" in src
    assert "fade: {" in src
    assert 'kind === "pop"' in src
    assert 'kind === "fade"' in src


def test_toast_uses_shared_fade():
    # showToast is the first reused (hide/re-show) surface to adopt the engine: it
    # fades in on show and out on dismiss. Every page loads motion.js now, but the
    # window.ClipgenMotion guard stays: viewer.py's inline steps each swallow
    # OSError, so a partially-failed export must still hide its toast.
    utils = (_WEB / "utils.js").read_text(encoding="utf-8")
    assert 'ClipgenMotion.animateIn(toastEl, "fade")' in utils
    assert 'ClipgenMotion.animateOut(toastEl, "fade")' in utils
    assert "window.ClipgenMotion" in utils


def test_animate_in_flushes_layout_before_animating():
    # A reveal (display:none → shown) followed by animate() in the same tick skips
    # the entry's opening frames — the element snaps in. animateIn forces a reflow
    # (offsetWidth read) first so toasts/overlay cards actually ease in.
    src = MOTION_JS.read_text(encoding="utf-8")
    assert "offsetWidth" in src


def test_studio_overlays_pop_in_via_motion():
    # The Studio overlay cards' entrance was migrated off the CSS cg-overlay-pop
    # keyframe onto ClipgenMotion.animateIn(card, "pop"); leaving the keyframe behind
    # would resurrect a parallel second animation system (cf. the stash landing).
    css = (_WEB / "studio.css").read_text(encoding="utf-8")
    # The keyframe and its animation rule are gone (a history note in a comment is
    # fine); an orphaned `animation: cg-overlay-pop` would be a parallel system.
    assert "@keyframes cg-overlay-pop" not in css
    assert "animation: cg-overlay-pop" not in css
    studio = (_WEB / "studio.js").read_text(encoding="utf-8")
    assert "function popOverlayIn(" in studio
    assert 'ClipgenMotion.animateIn(cardEl, "pop")' in studio


def test_studio_overlays_animate_out_on_dismiss():
    # The reused overlay surfaces get a symmetric exit: the card pops out while
    # the container carrying the backdrop veil fades, so the veil cannot snap out
    # behind the card. Both directions must be WAAPI or a forwards-filled exit
    # strands the reused card invisible on the next open (see motion.js).
    studio = (_WEB / "studio.js").read_text(encoding="utf-8")
    assert "function popOverlayOut(" in studio
    assert 'ClipgenMotion.animateOut(cardEl, "pop")' in studio
    assert 'ClipgenMotion.animateOut(overlayEl, "fade")' in studio
    assert 'ClipgenMotion.animateIn(overlayEl, "fade")' in studio
    # Generation guard: a stale exit must not hide a freshly-reopened overlay,
    # and a cancelled exit must be re-entered so neither element stays filled
    # invisible (mirrors showToast's _toastGen).
    assert "_popGen" in studio
    assert "_popExiting" in studio
    # Every dismiss path routes through the helper rather than snapping .hidden on.
    # (#logOverlay is deliberately excluded: it owns its own --veil-alpha timing.)
    for card in (
        ".status-card",
        ".confirm-card",
        ".gallery-card",
        ".build-status-card",
    ):
        assert re.search(
            r"popOverlayOut\([^,]+, qs\(\"" + re.escape(card) + r'"\)', studio
        ), f"{card} should be dismissed through popOverlayOut()"
