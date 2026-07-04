# Motion unification — migrating frontend animations onto `motion.js`

Rolling plan for consolidating the frontend's ad-hoc animations onto the shared
`assets/web/motion.js` (`window.ClipgenMotion`) WAAPI engine. Update the status markers as
work lands.

## Background

`motion.js` is a one-shot, promise-based micro-animation engine — the codebase's only
`element.animate()` usage. Keyframes touch **only `transform` + `opacity`** (compositor-safe),
tunables live in `ClipgenMotion.PARAMS` (console-retunable), it respects `prefers-reduced-motion`,
is size-aware, and **always resolves** so callers can "animate, then commit." It loads after
`utils.js` on **studio.html + screenspace.html only**. Guarded by `tests/test_motion_wiring.py`.

Public API: `animateOut(el, kind)` (`stash`/`delete`/`pop`/`fade`), `animateOutAll(els, kind)`
(staggered exit), `animateIn(el, kind)` (`stashLand`/`pop`/`fade`), `flyTo` (stubbed FLIP seam).

## Done (baseline)

- ✅ **`fade` + `pop` kinds** added; `animateIn`/`animateOut` dispatch generalized;
  `reducedFade` reuses `buildFadeKeyframes`.
- ✅ **Toast** (`utils.js showToast`) fades in/out, guarded on `window.ClipgenMotion`, with a
  generation token so a stale fade-out can't hide a re-shown toast.
- ✅ **Studio overlay entrances** — `.status-card`/`.confirm-card`/`.gallery-card`/`.log-panel`/
  `.build-status-card` pop in via `animateIn(card,"pop")` through a shared `popOverlayCardIn(cardEl,
  wasHidden)` helper (studio.js); the duplicate CSS `@keyframes cg-overlay-pop` was retired.
- ✅ **Reveal-reflow fix** — `animateIn` forces a layout flush (`void el.offsetWidth`) so entrances
  on elements revealed from `display:none` in the same tick actually ease in (see Gotchas).
- All browser-verified; `test_motion_wiring.py` covers the above.

## Open steps

### 1. Dismiss / exit animations for reused surfaces  ⏳
Give the toast and the Studio overlays a graceful **exit** (they currently snap out).

- **Constraint (reused-DOM hazard):** these surfaces toggle a container's `.hidden` and **reuse the
  same card DOM** (cards are never rebuilt). A `fill:"forwards"` exit strands the card invisible on
  the next open **unless** the entrance is also WAAPI (newest animation wins — which is why the
  toast's `animateIn`-on-show already makes its guarded `animateOut` safe). The overlay entrances are
  now WAAPI too, so exits are unblocked — but each exit **commit must be generation-guarded** so a
  stale fade-out can't hide a freshly-reopened overlay (mirror the toast's `_toastGen` pattern).
- **Order of risk (do safest first, browser-verify each):**
  1. `#buildStatus` corner card — non-blocking, no modal trap, no `onYes` callback. Safest.
  2. `#logOverlay`, `#statusOverlay` — blocking modals dismissed by user action; release the focus
     trap (`closeModalTrap`) **immediately**, delay only the visual `.hidden` until the exit settles.
  3. `#confirmOverlay` — trickiest: `handleYes()` runs `cleanup()` then `onYes()`. Keep the logical
     cleanup (listeners, trap, state) synchronous; only defer the visual hide. Verify `onYes` side
     effects (which often open another overlay) still sequence correctly against the ~150 ms fade.
- Likely adds a `popOut`/reuse of the `pop` exit; extend `test_motion_wiring.py`.

### 2. Toast fade on the remaining pages  ⏳
Today the toast animates only on studio + screenspace (elsewhere the guarded call no-ops → snaps).
To fade everywhere, add `<script src="motion.js"></script>` (after `utils.js`) to `viewer.html`,
`transcripts.html`, `workflows.html`, `gallery.html`, `start-overlay`. For the **timeline viewer**,
`motion.js` must be **inlined via `viewer.py`** (exported/offline viewers have no asset routes).
Extend the load-order assertions in `test_motion_wiring.py` to each newly-wired page.

### 3. Staggered enter (`animateInAll`)  ⏳
Add the symmetric mirror of `animateOutAll` (staggered entrance). First adopter: the start-overlay
cascade-in (`start-overlay.js` `runIntro`, `setTimeout(BASE + idx*STEP)` → `.is-in`) + the
`start-overlay-wordmark` keyframe.

### 4. Continuous-loop dedup (spinners / pulses) — DESIGN FORK  ⏳
Biggest duplication, worst fit for a one-shot/promise engine. Duplicated **spinner** ×4 (`spin`,
`task-card-spin`, `wf-spin`, `pill-spin` — identical `rotate(360deg)`) and **opacity pulse** ×5
(`studio-tab-pulse`, `status-pulse`, `streaming-pulse`, `pill-dot-pulse`, `so-screenspace-pulse`),
plus `pulse-dot`. Decide before building:
- **(a) motion.js `loop(el, kind)`** returning a handle with `.stop()` — JS-controlled, console-tunable.
- **(b) CSS-hoist** — collapse each into ONE shared `@keyframes` in `tokens.css`; stays declarative
  and simple. Often the better call for these — CSS loops are trivial and already performant.

## Explicitly out of scope (considered, excluded)

- **Box-shadow glows** (`wf-run-pulse`, `timeline-pulse-glow`) & **shimmers** (`skeleton-shimmer`,
  `filmstrip-shimmer`) — animate `box-shadow`/`background-position`, not `transform`/`opacity`; leave
  in CSS (shimmer dedup, if wanted, is a CSS-hoist).
- **SVG stroke-draws** (`brand-mark-draw`, `so-screenspace-march`) — `stroke-dashoffset`, SVG-specific.
- **Start-overlay `so-*` decorative artworks** (11 keyframes) — documented standing exception.
- **Panel collapse/expand** (studio `maxHeight`, screenspace `height`) — layout properties; a
  FLIP-height reimplementation trades content-squish artifacts. Not worth it.
- **~35 RAF render/throttle loops** (playheads, drag indicators, canvas/minimap draws, tooltips,
  scroll coalescing) — event-coalesced rendering, not declarative animation. Wrong tool.

## Gotchas / constraints (don't re-derive)

- **Reveal-reflow:** calling `element.animate()` in the SAME tick an element is revealed from
  `display:none` skips the entry's opening frames (it snaps to the end). `animateIn` already forces a
  reflow to fix this; any new reveal path inherits it. Diagnose with `el.getAnimations().length`.
- **Reused-DOM exits:** see step 1 — forwards-fill exit + reused card needs a WAAPI entrance + a
  generation-guarded hide commit.
- **Page reach is per-candidate/incremental**, not a blanket rollout; the `window.ClipgenMotion`
  guard degrades gracefully where the script isn't loaded.
- **Keep it `transform`+`opacity` only** (compositor thread) — the reason "clear 40 cards" is cheap.
- New durations/easings belong in `PARAMS` (no Python↔JS constant duplication); `tokens.css` has
  `--duration-*` if a bridge is ever wanted.

## Verification

- Extend `tests/test_motion_wiring.py` for every new kind/API and every new `<script src="motion.js">`
  include; keep `tests/test_frontend_satellite_wiring.py` green.
- Manual browser check per touched surface (repo rule: no headless-browser installs) — verify the
  animation, the reused-surface re-open, and `prefers-reduced-motion`. A DevTools
  `el.getAnimations()` snippet is the fastest bisect for "animation created but not painted."
