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
(staggered exit), `animateIn(el, kind)` (`stashLand`/`pop`/`fade`), `isReduced()` (the single
JS-side reduced-motion check), `flyTo` (stubbed FLIP seam).

Built on top of it, in `utils.js`: `popModalIn` / `popModalOut`, the shared reveal/dismiss path for
every blocking modal, paired with `openBlockingModal`'s focus-trap half and with the
`.cg-modal-veil` backdrop in `tokens.css`.

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
`popOverlayCardIn(cardEl, wasHidden)` became a symmetric **`popModalIn` / `popModalOut`** pair that
owns the reveal/hide itself, so the call sites stopped hand-rolling the capture-`wasHidden` dance.
It lives in `utils.js` beside `openBlockingModal`, whose logical half it pairs with, and **every**
blocking modal in the app uses it: Studio's `#statusOverlay` / `#confirmOverlay` /
`#galleryOverlay` / `#buildStatus` / `#logOverlay` and Composer's `#logOverlay`.

- Both directions are WAAPI, which is what stops a `fill:"forwards"` exit from stranding a reused
  card invisible on the next open. `_cgModalGen` makes a stale commit a no-op after a re-open;
  `_cgModalExiting` re-runs the entrance after a *cancelled* exit so nothing is left filled
  invisible.
- Logical cleanup (focus trap, listeners, confirm state) stays **synchronous**; only the visual
  hide waits on the fade, so no dismiss is ever swallowed. `handleYes()`'s `onYes()` — which
  usually opens another overlay — therefore still sequences correctly. Studio's log is the one
  surface that defers its trap release too: it is a panel, and focus escaping to the trigger while
  the veil is still up reads as the modal already being gone.
- **Backdrop: one shared `.cg-modal-veil`** (tokens.css) for all of them, replacing three
  byte-identical `::before` copies (Studio log, Composer log, Settings) *and* the flat
  `--color-backdrop` the three Studio dialogs used. Per-surface intensity via
  `--veil-blur-open` / `--veil-alpha-open`.
- The veil is ramped by toggling `.is-veiled`, **not** by fading the overlay's opacity. That is not
  a style preference: opacity on an ancestor establishes a backdrop root, so the `backdrop-filter`
  samples nothing. Measured on sharp stripes, even `opacity: .999` takes the frost from stdev 0.34
  to 69.88 against a 127 unblurred reference. The first cut of this step *did* fade the container,
  which is why the log had to keep a bespoke veil; unifying the backdrop is what removed that split.
- **One duration**, `--duration-veil` (360 ms), which `utils.js` reads back off the element rather
  than duplicating as a JS number. Before this the log raced three (150 ms card, 360 ms hide timer,
  460 ms veil) and the hide cut the veil at ~78 %.
- `#buildStatus` is the sole exception and carries no veil: a non-blocking corner card that must
  never cover the page, so only its card animates.

**What this fixed.** Composer's artifact log was a second, divergent copy of Studio's — its own
`LOG_*` constants and close timer, no card animation, no generation guard, and a *toggling* topnav
button where Studio's only opens. Spam-clicking it popped the panel in and out with the fade playing
only intermittently. Composer's button is now open-only too. A latent Studio bug went with it:
re-opening the log during its close read `wasHidden === false`, skipped the entrance, and left the
panel holding its exit's end state.

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
  `popModalIn`/`popModalOut` and `showToast`'s `_toastGen`.
- **A `backdrop-filter` cannot be faded via an ancestor's opacity.** Opacity on an ancestor
  establishes a backdrop root and the filter then samples nothing — measured, even `opacity: .999`
  removes the frost entirely. Ramp the custom properties the blur reads instead (`.cg-modal-veil` /
  `.is-veiled`). This is why the modal veil is a CSS transition rather than a WAAPI fade, and why
  its duration has to be published to JS (`--duration-veil`) instead of guessed.
- **Two copies of a surface will diverge.** The artifact log existed twice (Studio + Composer) and
  drifted into different constants, a missing card animation, a missing generation guard, and an
  extra toggle — which is what produced the spam-popping bug. Shared surfaces belong in `utils.js` /
  `tokens.css`, not copy-pasted per page.
- **Keep it `transform`+`opacity` only** (compositor thread) — the reason "clear 40 cards" is cheap.
- New durations/easings belong in `PARAMS` (no Python↔JS constant duplication); `tokens.css` has
  `--duration-*` if a bridge is ever wanted.
- **Continuous loops belong in CSS, deduped by hoisting into `tokens.css`** (`spin`, `cg-pulse`),
  not by moving onto the engine. Parameterize with a custom property rather than forking the
  keyframe — `var()` inside `@keyframes` resolves against the animated element, verified in-browser.

## Verification (what was run)

- `tests/test_motion_wiring.py` covers the API, the per-page load order across all nine templates,
  the mutation sites, the shared modal helpers + generation guard, that all six modal surfaces route
  through them, that the veil is defined once in `tokens.css`, that Composer's log is open-only, and
  the CSS-loop hoist (shared keyframes only in `tokens.css`; the six retired names appear nowhere).
  `tests/test_viewer_inline.py` covers the export inlining.
- `/check` green; `/ui-check` clean on all six live pages (no page errors, no `motion.js` 404).
- In-browser probes via `tests/ui/shot.py --eval`, since all three risky mechanisms here are
  engine-dependent:
  - `var(--pulse-trough)` resolves inside `@keyframes` (0.3 / 0.35 default / 0.55), and the
    negative-delay phase shift preserves the old look exactly — `.streaming-dot` reads 0.3 at t=0
    and 1.0 at mid-cycle.
  - the backdrop-root measurement above (stripe stdev 0.34 → 69.88 at `opacity: .999`).
  - the spam race, driven for real: close-then-instant-reopen settles **open** with the card at
    opacity 1 (previously stranded invisible); five rapid clicks on Composer's log button can no
    longer close it; Escape still dismisses it; and the newly-veiled dialogs render the frost
    (screenshot-confirmed against Settings).
- Exports generated from all three templates: engine inlined, no dangling `<script src>`; the
  offline file boots with `window.ClipgenMotion` present.
- **Left to a human eye** (a screenshot cannot judge these): the exits' *feel* at 150 ms card /
  360 ms veil, the confirm → status hand-off, and everything once more under
  `prefers-reduced-motion: reduce`.

## Known gaps left open (deliberately)

- `overview.html` has no `#toast` element but loads `start-overlay.js`, which toasts on folder
  errors — those messages vanish silently. Pre-existing, unrelated to motion.
- Composer's artifact log has no focus trap, where Studio's does. Now that both share one open/close
  path, adding `openBlockingModal` there is a small follow-up.
- `start-overlay` keeps its own veil (600 ms, on a real element rather than `::before`) because its
  intro sequencing is bound up with the launcher's cascade. It is the last copy.
