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


# Continuous CSS loops are the one animation family motion.js deliberately does
# NOT own — a one-shot promise engine is the wrong tool for an infinite loop, so
# they stay declarative. The dedup is therefore a CSS hoist: one definition in
# tokens.css that every page references by name, parameterized by a custom
# property where the values differed. Redefining either per page is how five
# near-identical opacity breathes accumulated in the first place.
_HOISTED_LOOPS = ("spin", "cg-pulse", "cg-shimmer-sweep")
_RETIRED_LOOPS = (
    "studio-tab-pulse",
    "status-pulse",
    "streaming-pulse",
    "pill-dot-pulse",
    "so-screenspace-pulse",
    "so-spinner-rotate",
)


def test_shared_css_loops_live_only_in_tokens():
    for name in _HOISTED_LOOPS:
        pattern = re.compile(r"@keyframes\s+" + re.escape(name) + r"\s*\{")
        owners = [
            path.name
            for path in sorted(_WEB.glob("*.css"))
            if pattern.search(path.read_text(encoding="utf-8"))
        ]
        assert owners == ["tokens.css"], (
            f"@keyframes {name} must be defined once, in tokens.css; found in {owners}"
        )


def test_retired_pulse_keyframes_are_gone():
    for path in sorted(_WEB.glob("*.css")):
        text = path.read_text(encoding="utf-8")
        for name in _RETIRED_LOOPS:
            assert name not in text, (
                f"{path.name} still references the retired {name}; "
                "use the shared cg-pulse / spin from tokens.css instead"
            )


def test_shimmer_is_parameterized_and_opts_out_of_reduced_motion():
    # The sweep is a continuous loop like the two above, so it lives in
    # tokens.css and is themed through --shimmer-base/--shimmer-peak rather than
    # per-page greys. It also *replaces* the text colour, so losing the
    # reduced-motion branch would leave those users staring at a moving band
    # they asked not to see — and, worse, no fallback fill at all.
    tokens = (_WEB / "tokens.css").read_text(encoding="utf-8")
    assert "var(--shimmer-base)" in tokens and "var(--shimmer-peak)" in tokens
    for block in ("  --shimmer-base:", "  --shimmer-peak:"):
        assert tokens.count(block) == 2, (
            f"{block.strip()} must be declared in both the dark :root and the "
            'html[data-theme="light"] block'
        )
    reduced = re.search(
        r"@media \(prefers-reduced-motion: reduce\) \{\s*\n"
        r"(?:\s*/\*.*?\*/\s*\n)?"
        r"\s*\.cg-shimmer \{(.*?)\}",
        tokens,
        re.DOTALL,
    )
    assert reduced, "tokens.css needs a prefers-reduced-motion opt-out for .cg-shimmer"
    body = reduced.group(1)
    assert "animation: none" in body
    assert "-webkit-text-fill-color:" in body, (
        "the opt-out must restore a visible fill, not just stop the animation"
    )


def test_pulse_trough_is_parameterized():
    # The shared breathe reads a custom property for its trough so one keyframe
    # covers 0.3 / 0.35 / 0.55 dots. Losing the var() would silently flatten
    # every consumer onto the default.
    tokens = (_WEB / "tokens.css").read_text(encoding="utf-8")
    assert "var(--pulse-trough, 0.35)" in tokens


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


def test_old_css_overlay_pop_removed():
    # The Studio overlay cards' entrance was migrated off the CSS cg-overlay-pop
    # keyframe onto ClipgenMotion.animateIn(card, "pop"); leaving the keyframe behind
    # would resurrect a parallel second animation system (cf. the stash landing).
    # A history note in a comment is fine; an orphaned `animation:` rule is not.
    css = (_WEB / "studio.css").read_text(encoding="utf-8")
    assert "@keyframes cg-overlay-pop" not in css
    assert "animation: cg-overlay-pop" not in css


# Every blocking modal shares ONE reveal/dismiss path (utils.js popModalIn /
# popModalOut) and ONE backdrop (.cg-modal-veil in tokens.css). Before that,
# Studio and Composer each carried their own artifact-log copy with different
# durations and no generation guard, which is why spam-toggling Composer's popped
# it in and out while the fade played only intermittently.

_MODAL_SURFACES = (
    # (page script, overlay id, card class)
    ("studio.js", "#statusOverlay", ".status-card"),
    ("studio.js", "#confirmOverlay", ".confirm-card"),
    ("studio.js", "#galleryOverlay", ".gallery-card"),
    ("studio.js", "#buildStatus", ".build-status-card"),
    ("studio.js", "#logOverlay", ".log-panel"),
    ("composer.js", "#logOverlay", ".log-panel"),
)


def test_modal_animation_helpers_are_shared_in_utils():
    utils = (_WEB / "utils.js").read_text(encoding="utf-8")
    assert "var popModalIn = function (overlayEl, cardEl)" in utils
    assert "var popModalOut = function (overlayEl, cardEl, commit)" in utils
    assert 'ClipgenMotion.animateIn(cardEl, "pop")' in utils
    assert 'ClipgenMotion.animateOut(cardEl, "pop")' in utils
    # Generation guard: a stale exit must not hide a freshly-reopened modal, and a
    # cancelled exit must re-run its entrance or the card stays filled invisible.
    assert "_cgModalGen" in utils
    assert "_cgModalExiting" in utils
    # The veil is a CSS transition, so its duration is read back off the element
    # instead of being duplicated here as a number that could drift.
    assert "--duration-veil" in utils
    # The overlay's own opacity must never be animated — that would establish a
    # backdrop root and kill the frost (see the tokens.css comment).
    assert "animateOut(overlayEl" not in utils
    assert "animateIn(overlayEl" not in utils


def test_every_modal_uses_the_shared_helpers():
    # openPopModal / closePopModal wrap the pair; pages may call either layer.
    utils = (_WEB / "utils.js").read_text(encoding="utf-8")
    assert "var openPopModal = function (overlayEl, cardEl, opts)" in utils
    assert "popModalIn(overlayEl, cardEl);" in utils
    assert "popModalOut(overlayEl, cardEl, function () {" in utils
    for page, overlay, card in _MODAL_SURFACES:
        src = (_WEB / page).read_text(encoding="utf-8")
        assert re.search(
            r"(popModalIn|openPopModal)\([^,]+, qs\(\"" + re.escape(card) + r'"\)', src
        ), f"{page} {overlay} should reveal through popModalIn()"
        assert re.search(
            r"(popModalOut|closePopModal)\([^,]+, qs\(\"" + re.escape(card) + r'"\)',
            src,
        ), f"{page} {overlay} should dismiss through popModalOut()"


def test_modal_veil_is_defined_once_in_tokens():
    # Five copies of this declaration is what the hoist removed (Studio's log,
    # Composer's log, Settings, Start, plus the flat backdrop the dialogs used).
    owners = [
        path.name
        for path in sorted(_WEB.glob("*.css"))
        if "backdrop-filter: blur(var(--host-blur))" in path.read_text(encoding="utf-8")
    ]
    assert owners == ["tokens.css"], (
        f"the frosted modal veil should live only in tokens.css, not {owners}"
    )


def test_veiled_surfaces_opt_in_via_class():
    # The class is what popModalIn/Out branch on, so a surface that wants the
    # frost has to carry it in markup (or, for JS-built overlays, in the el()
    # call). `cg-modal-veil` generates the layer as a ::before; `cg-veil` +
    # `cg-veil-layer` is for a surface whose backdrop must be a real element.
    for page in ("studio.html", "composer.html"):
        html = (_WEB / page).read_text(encoding="utf-8")
        assert "cg-modal-veil" in html, f"{page} should opt its overlays into the veil"
    settings = (_WEB / "settings-modal.js").read_text(encoding="utf-8")
    assert "cg-modal-veil" in settings
    assert "is-veiled" in settings
    # Start overlay: host on the root, layer on the real backdrop element.
    start = (_WEB / "start-overlay.html").read_text(encoding="utf-8")
    assert "start-overlay cg-veil hidden" in start
    assert 'class="cg-veil-layer" data-role="backdrop"' in start
    assert "is-veiled" in (_WEB / "start-overlay.js").read_text(encoding="utf-8")


def test_start_overlay_backdrop_stays_an_element():
    # It is the click-to-dismiss target, and a pseudo-element cannot be an event
    # target — so this surface must NOT be switched to cg-modal-veil's ::before.
    js = (_WEB / "start-overlay.js").read_text(encoding="utf-8")
    assert 'on(els.backdrop, "click", close)' in js
    start = (_WEB / "start-overlay.html").read_text(encoding="utf-8")
    assert "cg-modal-veil" not in start


def test_veil_layer_paint_is_shared_by_both_carriers():
    # One paint definition, two selectors. Splitting them would let the real
    # element and the pseudo drift apart again.
    tokens = (_WEB / "tokens.css").read_text(encoding="utf-8")
    assert ".cg-veil-layer,\n.cg-modal-veil::before {" in tokens
    # Only the pseudo may ignore pointer events; a real layer is clickable.
    assert '.cg-modal-veil::before {\n  content: "";\n  pointer-events: none;' in tokens


def test_toast_element_exists_wherever_something_toasts():
    # showToast no-ops without #toast, so a page that loads a toasting script
    # needs the element. start-overlay.js and export-actions.js both toast.
    toasting = ("start-overlay.js", "export-actions.js")
    for page in ALL_TEMPLATES:
        html = (_WEB / page).read_text(encoding="utf-8")
        loads_toaster = any(f'src="{name}"' in html for name in toasting)
        if not loads_toaster:
            continue
        assert 'id="toast"' in html, (
            f"{page} loads a script that calls showToast but has no #toast element"
        )


def test_composer_log_traps_focus_like_studio():
    # It was the one modal without a trap: Tab walked out into the timeline
    # behind the veil and Escape fell through to the page's back-out cascade.
    composer = (_WEB / "composer.js").read_text(encoding="utf-8")
    utils = (_WEB / "utils.js").read_text(encoding="utf-8")
    # Trap and topnav gate come from openPopModal's opts (utils.js).
    assert (
        'openPopModal(qs("#logOverlay"), qs(".log-panel"), { modalOpen: true, onEscape: closeLog })'
        in composer
    )
    assert (
        'closePopModal(qs("#logOverlay"), qs(".log-panel"), { modalOpen: true })'
        in composer
    )
    assert "openBlockingModal(overlayEl, {" in utils
    assert "closeBlockingModal(overlayEl)" in utils
    assert 'if (opts.modalOpen) document.body.classList.add("modal-open");' in utils
    # With the trap owning Escape, a cascade branch for the log would be dead code.
    assert "logOverlayVisible" not in composer


def test_composer_log_is_open_only():
    # A toggle on the topnav button made it a spam surface that fought its own
    # animation; dismissal is the X, the backdrop, or Escape, as in Studio.
    composer = (_WEB / "composer.js").read_text(encoding="utf-8")
    assert "toggleLogPanel" not in composer
    assert 'logBtn.addEventListener("click", openLog)' in composer


def test_motion_exposes_reduced_motion_check():
    # utils.js pairs a WAAPI exit with the veil's CSS transition and must not
    # wait on that transition when reduced motion has disabled it. One JS-side
    # source for the query, rather than a second matchMedia call.
    src = MOTION_JS.read_text(encoding="utf-8")
    assert "function isReduced(" in src
    assert "isReduced: isReduced" in src
    assert "ClipgenMotion.isReduced()" in (_WEB / "utils.js").read_text(
        encoding="utf-8"
    )
