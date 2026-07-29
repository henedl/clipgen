# Motion unification — migrating frontend animations onto `motion.js`

**Status: done.** Every step is either landed or explicitly descoped with its reasoning below.
Kept as the record of where the boundary between `motion.js` and CSS was drawn, and why.

## Background

`motion.js` is a one-shot, promise-based micro-animation engine — the codebase's only
`element.animate()` usage. Keyframes touch **only `transform` + `opacity`** (compositor-safe),
tunables live in `ClipgenMotion.PARAMS` (console-retunable), it respects `prefers-reduced-motion`,
is size-aware, and **always resolves** so callers can "animate, then commit." It loads after
`utils.js` on **every page**, and is inlined into exported viewers by `viewer.py`. Guarded by
`tests/test_motion_wiring.py`.

Public API: `animateOut(el, kind)` (`stash`/`delete`/`pop`/`fade`), `animateOutAll(els, kind)`
(staggered exit), `animateIn(el, kind)` (`stashLand`/`pop`/`fade`), `flyTo` (stubbed FLIP seam).

## Done

### Baseline
- ✅ **`fade` + `pop` kinds** added; `animateIn`/`animateOut` dispatch generalized;
  `reducedFade` reuses `buildFadeKeyframes`.
- ✅ **Toast** (`utils.js showToast`) fades in/out, guarded on `window.ClipgenMotion`, with a
  generation token so a stale fade-out can't hide a re-shown toast.
- ✅ **Studio overlay entrances** — `.status-card`/`.confirm-card`/`.gallery-card`/`.log-panel`/
  `.build-status-card` pop in via the shared motion engine; the duplicate CSS
  `@keyframes cg-overlay-pop` was retired.
- ✅ **Reveal-reflow fix** — `animateIn` forces a layout flush (`void el.offsetWidth`) so entrances
  on elements revealed from `display:none` in the same tick actually ease in (see Gotchas).

### 1. Dismiss / exit animations for reused surfaces ✅
`popOverlayCardIn(cardEl, wasHidden)` became a symmetric `popOverlayIn` / `popOverlayOut` pair in
`studio.js` that owns the reveal/hide itself, so the call sites stopped hand-rolling the
capture-`wasHidden` dance. Applied to `#statusOverlay`, `#confirmOverlay`, `#galleryOverlay`, and
`#buildStatus`.

- The **container** carries the backdrop veil, so it fades (`fade`) while the **card** pops
  (`pop`). Animating only the card left the veil snapping out ~150 ms behind it. Where the card is
  the container's only child (`#buildStatus`, no veil) the two opacities multiply into a slightly
  steeper fade — intentional; one helper for all four surfaces is worth more than the micro-detail.
- Both directions are WAAPI, which is what stops a `fill:"forwards"` exit from stranding a reused
  card invisible on the next open. `_popGen` makes a stale commit a no-op after a re-open;
  `_popExiting` re-enters a *cancelled* exit so neither element is left filled invisible.
- Logical cleanup (focus trap, listeners, confirm state) stays **synchronous**; only the visual
  hide waits on the fade, so no dismiss is ever swallowed. `handleYes()`'s `onYes()` — which
  usually opens another overlay — therefore still sequences correctly.
- **`#logOverlay` deliberately does not use the pair.** It owns its own veil timing
  (`--veil-alpha` / `--host-blur` over `LOG_EXIT_MS`), and a container fade would fight that. It
  already had a card-only exit; this pass additionally fixed a latent bug where re-opening during
  the 360 ms close read `wasHidden === false`, skipped the entrance, and left the panel holding its
  exit's end state.

### 2. motion.js on every page ✅
The script tag went onto all nine templates (it was on studio + screenspace only), and
`viewer.py::_generate_viewer_html` inlines it for exports, mirroring the `hotkeys.js` /
`card-scrubber.js` blocks. Placement among those blocks is immaterial: `utils.js` is prepended last
so it always ends up first, and every consumer reads `window.ClipgenMotion` lazily inside a
function, never at load time.

**Honest accounting of the payoff.** The engine's only cross-page consumer is `showToast`, and a
page without a `#toast` element never shows a toast at all:

| Template | `#toast` | Effect of this step |
| --- | --- | --- |
| studio, screenspace | yes | already animated |
| transcripts, workflows, composer | yes | **toasts now fade instead of snapping** |
| overview, viewer, gallery, timeline-viewer | **no** | groundwork only — nothing animates there yet |

Those four were wired anyway for uniform availability (the cost is one deferred, cached ~13 KB
script) so the next surface added to them doesn't silently hit the `window.ClipgenMotion` guard —
the exact trap this step exists to remove. The guards themselves stay: `viewer.py`'s inline steps
each swallow `OSError`, so a partially-failed export must still be able to hide its toast.

**Adjacent gap found, not fixed:** `overview.html` has no `#toast` element, but it loads
`start-overlay.js`, which toasts on folder errors (`start-overlay.js:1200/1216/1553`). Those
messages silently vanish on Overview. Pre-existing and unrelated to motion; worth its own fix.

### 4. Continuous-loop dedup — resolved as **CSS-hoist** ✅
The fork was (a) a JS `loop(el, kind)` handle vs (b) collapsing each family into one shared
`@keyframes`. **(b) won**: a one-shot promise engine is the wrong tool for an infinite loop, and
CSS loops are already trivial and performant. The precedent was set earlier when the four spinner
keyframes collapsed into `@keyframes spin` in `tokens.css`.

So the remaining work was the opacity family. Five near-identical breathes across four stylesheets
differed only in trough (0.3 / 0.35 / 0.55) and duration; they became one `cg-pulse` in
`tokens.css`, parameterized by a `--pulse-trough` custom property. Retired:
`studio-tab-pulse`, `status-pulse`, `streaming-pulse`, `pill-dot-pulse`, `so-screenspace-pulse`,
and `so-spinner-rotate` (a byte-identical clone of `spin`).

- The two **phase-inverted** consumers (`.streaming-dot`, `.art-screenspace__solid`) keep their
  exact prior look via a negative delay of half the duration, placed in the `animation` shorthand's
  4th slot — a trailing `animation-delay` would be reset by the shorthand.
- `pulse-dot` **stays its own keyframe**: it translates rather than fades, so it is a different
  animation, not a duplicate one.
- `.streaming-dot` and `pulse-dot`'s three circles were the only animated indicators with no
  `prefers-reduced-motion` override. They got one.

## Explicitly out of scope (considered, excluded)

- **Step 3, staggered enter (`animateInAll`)** — descoped. Its only candidate adopter was the
  start-overlay cascade (`start-overlay.js runIntro` + `.cascade-in.is-in`), and that is a tuned CSS
  *transition* system (420–460 ms `cubic-bezier(.22, 1, .36, 1)`) whose end state must **persist** on
  a long-lived interactive panel. A WAAPI port means either keeping `.is-in` anyway (two systems for
  one effect) or holding a `fill:"both"` forever, which then blocks every later CSS
  transform/opacity change on those elements. Adding the API with no adopter is speculative.
  Revisit only if a genuine staggered-entrance surface appears.
- **Box-shadow glows** (`wf-run-pulse`, `timeline-pulse-glow`) & **shimmers** (`skeleton-shimmer`,
  `filmstrip-shimmer`) — animate `box-shadow`/`background-position`, not `transform`/`opacity`; leave
  in CSS (shimmer dedup, if wanted, is a CSS-hoist).
- **SVG stroke-draws** (`brand-mark-draw`, `so-screenspace-march`) — `stroke-dashoffset`, SVG-specific.
- **Start-overlay `so-*` decorative artworks** — documented standing exception.
- **Panel collapse/expand** (studio `maxHeight`, screenspace `height`) — layout properties; a
  FLIP-height reimplementation trades content-squish artifacts. Not worth it.
- **~35 RAF render/throttle loops** (playheads, drag indicators, canvas/minimap draws, tooltips,
  scroll coalescing) — event-coalesced rendering, not declarative animation. Wrong tool.

## Gotchas / constraints (don't re-derive)

- **Reveal-reflow:** calling `element.animate()` in the SAME tick an element is revealed from
  `display:none` skips the entry's opening frames (it snaps to the end). `animateIn` already forces a
  reflow to fix this; any new reveal path inherits it. Diagnose with `el.getAnimations().length`.
- **Reused-DOM exits:** a `fill:"forwards"` exit on DOM that is hidden-and-reshown (rather than
  rebuilt) strands the element invisible on the next open **unless** the entrance is also WAAPI —
  and the hide commit must be generation-guarded so a stale exit can't hide a freshly-reopened
  surface. A *cancelled* exit needs the entrance re-run too, or it stays filled invisible. See
  `popOverlayIn`/`popOverlayOut` and `showToast`'s `_toastGen`.
- **Keep it `transform`+`opacity` only** (compositor thread) — the reason "clear 40 cards" is cheap.
- New durations/easings belong in `PARAMS` (no Python↔JS constant duplication); `tokens.css` has
  `--duration-*` if a bridge is ever wanted.
- **Continuous loops belong in CSS, deduped by hoisting into `tokens.css`** (`spin`, `cg-pulse`),
  not by moving onto the engine. Parameterize with a custom property rather than forking the
  keyframe — `var()` inside `@keyframes` resolves against the animated element, verified in-browser.

## Verification (what was run)

- `tests/test_motion_wiring.py` covers the API, the per-page load order across all nine templates,
  the mutation sites, the overlay exit helpers + generation guard, and the CSS-loop hoist (shared
  keyframes live only in `tokens.css`; the six retired names appear nowhere).
  `tests/test_viewer_inline.py` covers the export inlining.
- `/check` green; `/ui-check` clean on all six live pages (no page errors, no `motion.js` 404).
- In-browser probes via `tests/ui/shot.py --eval` confirmed `var(--pulse-trough)` resolves inside
  `@keyframes` (0.3 / 0.35 default / 0.55) and that the negative-delay phase shift preserves the
  old look exactly: `.streaming-dot` reads 0.3 at t=0 and 1.0 at mid-cycle.
- Exports generated from all three templates: engine inlined, no dangling `<script src>`; the
  offline file boots with `window.ClipgenMotion` present.
- **Left to a human eye** (a screenshot cannot judge these): the overlay exits' feel, re-opening a
  surface mid-fade, the confirm → status hand-off, and everything once more under
  `prefers-reduced-motion: reduce`.
