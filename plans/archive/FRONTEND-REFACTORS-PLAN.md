# Frontend refactors & carveouts — plan

> **Status: closed 2026-07-25.** 13 of 17 items landed (commit notes inline).
> The four unchecked items are deliberately not-doing rather than pending: three
> were logged as _optional later_ / _investigate only_ / _opportunistic_, and the
> fourth — the **build** carve (`studio.js 4136–4615`) — is blocked on a separate
> reel-vs-viewer split and should be re-opened as its own plan if wanted.

Originally proposed (2026-06-30). Tracks a sweep of refactor/carveout opportunities
across `assets/web/`. Each numbered item is sized to be one focused `refactor:` commit/PR.
Check items off and add a "Done" note as they land (per AGENTS.md plan-maintenance rule).

## Why

Two page hubs are large outliers and are the explicit, already-in-progress
architectural direction of the repo:

| File | Lines | Satellites today | Notes |
|------|-------|------------------|-------|
| `screenspace.js` | 5835 | 5 (tasks, results, multitool-params, calibration, color) + utils leaf | still the largest JS file *after* carving 5 |
| `studio.js` | 5278 | 1 (intake) | AGENTS.md: "mid-carve… generate/trim/stash/scrubber not yet carved" |

`transcripts.js` and `workflows.js` are already cleanly carved and are **not** targets.
The carve pattern, its load-order contract, and the `tests/test_frontend_satellite_wiring.py`
guard are all established — so these carves are low-novelty, incremental, and individually shippable.

The remaining themes (cross-cutting JS dedup, CSS token/duplication wins) are smaller,
opportunistic, and independent of the carves.

## Verified-false leads (do NOT action)

Two "dead code" findings from analysis were checked and are **live** — left here so they
aren't re-flagged:

- `gallery.js:createGalleryLoopVideo` — used at `gallery.js:116`.
- `screenspace.js` `_pendingFrameTs` / `_loadedFrameTs` — used in the stale-frame guard
  (`screenspace.js:1491–1541`); the lines 1019–1020 occurrences are resets, not dead reads.

---

## Theme A — JS hub→satellite carveouts (headline)

Procedure for every item: `agents/skills/carve-satellite/SKILL.md`. After each: `node --check`
the touched files, run `tests/test_frontend_satellite_wiring.py` + `tests/test_studio_frontend_source.py`,
then `/check`. Add each new satellite to the page's `<script>` load order respecting the
load-order contract (a satellite that *destructures* another's published fn loads after it;
otherwise late-bind `NS.fn(...)`).

### Studio (`window.ClipgenStudio` / STUDIO)

- [x] **A1. `studio-scrubber.js`** — Done (`5ca7440`). Carved the sprite-prefetch + per-card
  scrubber wiring (the true cluster was `studio.js:2278–2396`, ~120 lines; the rest of the
  agent's 2276–2553 range was unrelated drag/queue code). Hub keeps `attachQueueScrubbers` +
  a new `resetScrubberPrefetch` delegator (the latter replaces a bare `_spritePrefetchQueue = []`
  reset so the queue stays satellite-private). Loads before `studio-intake.js`. studio.js −121/+8.
- [x] **A2. `studio-trim.js`** — Done. Carved the duration-badge trim pop-over cluster
  (the post-A1 range was `studio.js:2616–2968`, ~353 lines: 5 `TRIM_*`/`activeTrim` vars +
  `closeTrimPopover`/`positionTrimPopover`/`bindTrimDrag`/`makeTrimButton`/`openTrimPopover`/
  `appendDurationBadge`/`buildCellOverrides`). Self-contained — no `state.*` access; mutates
  the passed-in `item` + a satellite-local `activeTrim`. Hub keeps `appendDurationBadge`/
  `buildCellOverrides` delegators (called at `studio.js:2268` and the generate/reel builders);
  the satellite reaches the hub's `saveQueues`/`isIntakeSource` through STUDIO (no `renderQueue`
  dep — the rerender flows via the passed-in `renderFn` callback). Loads after the hub, before
  `studio-intake.js` (order vs. the other satellites is free — no cross-destructuring). studio.js
  5164 → 4824 (−340); `studio-trim.js` 383 lines. **Must load before A3** (generate needs
  `buildCellOverrides`).
- [x] **A3. `studio-generate.js`** — Done. Carved the streaming `api/generate` +
  `api/generate-intake` flow: `onGenerate`, `onCancelGenerate`, `buildGenerateCardIndex`,
  `isGenerateFetchAborted` (post-A2 the move-set was 3 non-contiguous ranges — `3290–3303`,
  `3375–3662`, `3669–3679` — with the ETA section and `onCancelReel` left in the hub between
  them). studio.js 4824 → 4531 (−293); satellite 365 lines.
  - **Hub keeps** (shared with the deferred reel/build path + job-status polling, so they did
    **not** move): the card painters `setCardQueued`/`clearCardStatus`/`setCardResult`,
    `readNDJSONStream`, `updateGenerateProgress`, the `_paint*`/`_tick*` painters, and the
    `_generateEtaTracker`/`_studioEtaTicker` objects. The satellite reaches all of these +
    `setArtifactGenerating`/`showResult`/`revealStatusOverlay`/`stampLog`/`isIntakeSource`/
    `buildCellOverrides` (trim) through STUDIO (11 new publications).
  - **Hub delegators added**: `onGenerate`/`onCancelGenerate` (button wiring at `studio.js`
    init still calls them).
  - **Bare-var gotcha**: `onGenerate` drives the hub-owned `_generateEtaTracker`/
    `_studioEtaTicker` — published as object refs on STUDIO (like `state`), not left bare.
  - **`setButtonProgress`** is a `ClipgenPrimitives` namespace fn the hub aliases locally;
    re-aliased the same way in the satellite (the wiring guard caught the missing alias).
  - **Load order**: after `studio-trim.js` (uses its `STUDIO.buildCellOverrides`), before
    `studio-intake.js`.
- [x] **stash** — Done (2026-07-23) as `studio-stash.js` (~377 lines; hub 5086 → 4778). The
  `renderQueue` coupling resolved cleanly: the stash configs hold the hub's already-published
  `STUDIO.renderReelQueue`/`renderArtifactQueue` refs, and the queue rerender flows via
  `cfg.renderQueue()`. The two hub drop-target callbacks (which wrote `_justStashedId`) moved
  into the satellite as late-bound `STUDIO.stashDropReel`/`.stashDropArtifacts`; hub keeps 6
  same-named delegators (loadStashes, loadArtifactStashes, stashCurrentReel,
  stashCurrentArtifacts, revealEmptyStashAreas, hideEmptyStashAreas) and newly publishes
  `isReelQueueLocked`/`isArtifactQueueLocked`/`cellKey`/`updateSingleCellClass`/
  `ssEnqueueThumbCustom`.
- [ ] _Defer:_ **build** (`4136–4615`, reel/timeline/gallery interleaved — needs a separate
  reel-vs-viewer split first). Re-evaluate.

### Screenspace (`window.ClipgenScreenspace` / SS)

- [x] **A4. `screenspace-overlay.js`** — Done. Carved `renderOverlay` (`screenspace.js:2763–2979`,
  217 lines). It's a pure read of `state` + 7 hub helpers (reached via SS; `hexToRgba`/`qs` are
  ambient utils.js globals). Hub keeps a `renderOverlay` delegator for its **33** call sites;
  published the 5 helpers the satellite needs that weren't already on SS (`regionColorForIndex`,
  `computeLabelRect`, `getThemeColors`, `templateOverlayBounds`, `_overlayEligibleForActiveTool`).
  Loads after the hub, **before** `screenspace-tasks.js`/`-results.js` (they destructure
  `SS.renderOverlay`). screenspace.js −216/+11.
- [x] **A5. `screenspace-timeline.js`** — Done. Carved the timeline cluster (post-A4 range was
  two non-contiguous blocks: `screenspace.js:2769–3411` + `renderTimelineLegend` at `3484–3507`,
  with the unrelated "Tool info tooltip" block kept in the hub between them) plus the
  `getTimelineRect` helper + `_timelineHitRects`/`_cachedTimelineRect`/`TIMELINE_CANVAS_HEIGHT`.
  15 functions, ~660 lines. screenspace.js 5629 → 4968 (−661); satellite 714 lines.
  - **state**: `timelineZoom/Offset/Dragging`, `inMarker/outMarker`, `hoveredBoundaryTs`,
    `amplitudeGraphEnabled` were already on `state` (read by the run button etc.) — left there.
  - **Hub delegators added** (`initTimeline`/`renderTimeline`/`renderPlayhead` — all called by
    hub init/seekPlayhead/the video RAF loop). `showSsTooltip`/`hideSsTooltip`/`sizeTimelineCanvas`
    etc. are cluster-internal — no delegators.
  - **Satellite→hub via SS**: `loadFrame`/`seekPlayhead` (frame viewer, stay in hub — `seekPlayhead`
    newly published), `taskTypeColor`/`getThemeColors`/`buildTypeIcon`/`iconSpan`. `findTask`/
    `focusedTaskId` are **late-bound `SS.fn(...)`** (owned by tasks.js, which loads *after* timeline).
  - **Load order**: timeline loads after the hub/overlay but **before** tasks/results, which
    destructure `SS.renderTimeline`/`SS.updateMarkerInfo` at load.
  - **Bare-var gotcha caught**: the resize/scroll handlers cleared the hub-owned `_cachedOverlayRect`
    (a strict-mode `ReferenceError` the wiring test can't see) — added `SS.invalidateOverlayRect`
    and route through it.
- [x] **A6. `screenspace-overlay-interaction.js`** — Done. Carved the region draw/drag/resize
  state machine + region toolbar (post-A4/A5 the contiguous range was `screenspace.js:1735–2319`,
  ~585 lines: `initRegionDrawing` + the overlay-canvas mousedown/move/up handlers, `canvasCoords`/
  `findHitRegion`/`saveRegionUpdate`/`scheduleOverlayRender`/`flushOverlayRender`/`computeLabelRect`,
  the region-name modal, `renderRegionChips`/`updateRegionButtons`/`updateRegionChipsOverflow`).
  `initRegionDrag` + `templateOverlayBounds` (after the stashing gap) and the stashing section stay
  in the hub. screenspace.js 4462 → 3885 (−577); satellite 662 lines.
  - **Key gotcha (resolved):** moved `_overlayRaf` + `_cachedOverlayRect` + `invalidateOverlayRect`
    into the satellite (the cluster owns ~all 15 uses). The hub's Escape handler now calls the
    `invalidateOverlayRect()` delegator instead of touching the var; the timeline satellite drops
    the rect via `SS.invalidateOverlayRect` (A6 loads before timeline). `_playheadRaf` (interleaved
    with `_overlayRaf` in the hub var block) was **left in the hub** — it's the playhead's RAF.
  - **Publishes** `initRegionDrawing`/`renderRegionChips`/`updateRegionButtons`/`computeLabelRect`/
    `invalidateOverlayRect`/`hideRegionNameModal`; hub keeps same-named delegators for all but
    `computeLabelRect` (publish-only — read by the overlay satellite, no hub caller). Removed those
    SS publications from the hub block so the delegators don't republish onto themselves and recurse.
  - **New hub publications for the satellite to read:** `pauseVideo` + the stash/pin toolbar
    callbacks `stashRegions`/`pinCurrentFrame`/`togglePinTrayVisibility`/`clearAllPins`/
    `updatePinButtons` (referenced as button-handler values inside `initRegionDrawing`, so a bare
    cross-file reference that the wiring test can't see — caught by a free-identifier scan).
    `renderOverlay`/`setTargetColor` are late-bound local wrappers (overlay/color load after A6).
  - **Load order:** loads after model-view, **before** overlay (`SS.computeLabelRect`)/timeline
    (`SS.invalidateOverlayRect`)/tasks (`SS.renderRegionChips`/`SS.updateRegionButtons`).
- [x] **A7. `screenspace-model-view.js`** — Done. Carved the live preprocessed-frame preview +
  overlay-layer UI cluster (post-A4/A5 the range was `screenspace.js:3474–3986`, ~513 lines:
  `initModelView`/`toggleModelView`/`refreshModelView`/`_doRefreshModelView`, the overlay-layer
  helpers, the preview-region resolvers, and `_updateMinAreaReadout`/`_collectPreviewParams` +
  the `_modelView*`/`MODEL_VIEW_META` vars). Self-contained sessionStorage (`ss_overlayEnabled`/
  `ss_overlayLayer`). screenspace.js 4968 → 4462 (−506); satellite 559 lines.
  - **Plan correction:** `_collectPreviewParams` is **not** called by `initRunButton` (that uses
    `gatherWorkflowParams`) — it's internal to `_doRefreshModelView`, so it stays private, no SS
    publication.
  - **Reaches via SS:** `state`, `normalizeRegionRef`/`activeRegionRef` (destructured); `SS.renderOverlay(...)`
    (3 sites) and `SS.getColorHiddenInputs()` late-bound (overlay/color load after).
  - **Publishes** `initModelView`/`refreshModelView`/`_updateOverlayUi`/`_overlayEligibleForActiveTool`/
    `_updateMinAreaReadout`/`_previewRegionRef`; hub keeps same-named delegators for the first five.
    These three were moved off the hub's SS block (else the delegators would publish onto themselves
    and recurse).
  - **Load order:** loads first after the hub, **before** overlay/tasks/multitool-params/calibration,
    which destructure `_overlayEligibleForActiveTool`/`_updateMinAreaReadout`/`_previewRegionRef`.
- [ ] _Optional later:_ single-tool param builders (`3771–4346`, ~576) and run button
  (`4954–5296`, ~343). Lower ROI; param rehydration (`applyColorMode`/`applyNormalizeMode`)
  must stay hub-side for task restore.

**Stays in both hubs** (do not move): the shared `state` object, request-version/cache-
invalidation guards, participant routing, frame loading, the `DOMContentLoaded` init loop,
settings-sync, and STUDIO/SS namespace publication.

Outcome (A1–A3, A4–A7 all landed): studio 5278 → 4531; screenspace 5835 → 3885 (~1950 lines
moved into 4 new screenspace satellites: overlay, timeline, model-view, overlay-interaction).

---

## Theme B — cross-cutting JS consolidation

- [x] **B1. `createSSEStream()` in `utils.js`** — Done in **#493** (landed on `master`
  independently as `createSSEStream(url, opts)`; migrated `screenspace-tasks.js`,
  `workflows-runs.js`, and the studio sites). A duplicate B1 was developed on this branch and
  **dropped** when rebasing onto `master` to avoid the collision — #493's version stands.
- [x] **B2. Generalize the modal focus trap** — Done in **#493** (`openBlockingModal` in
  `utils.js`; `studio.js`/`transcripts.js` migrated off the hand-rolled traps).
- [ ] _Investigate only (do not blind-merge):_ color conversion exists twice with **incompatible**
  HSV ranges — `color-picker.js:36–77` (h∈[0,1]) vs `screenspace-utils.js:14–51` (OpenCV h 0–180,
  s/v 0–255). Merging risks silent corruption. At most: clarify names/comments
  (`rgbToHsvStandard` vs `rgbToHsvCv`). Likely skip.

---

## Theme C — CSS token & duplication wins (low risk, mechanical)

These follow the "convert touched values to tokens when editing" rule. Done as their own small
commits **or** folded into whichever page CSS a Theme-A carve touches. **Not** a blanket sweep
(AGENTS.md: "don't bulk-rewrite either system" / page-CSS spacing left alone unless touched).

- [x] **C1. `border-radius: 50%` → `var(--radius-full)`** — Done. 21 occurrences across 7 CSS
  files; all verified square (circles/dots/spinners) so `50%`≡`999px`. `viewer.css` already uses
  radius tokens, confirming they reach the offline export.
- [x] **C2. Backdrop literal `rgba(0,0,0,0.4)`** → Done. Added `--color-backdrop` token
  (`tokens.css`) and applied to the 7 true modal overlays (screenspace/studio×4/transcripts×2).
  Left `settings-modal.css .card-tile-text-overlay` (a card text-legibility scrim, different
  semantic) and the two box/drop-shadow literals as-is.
- [x] **C3. Consolidate the duplicated theme-toggle.** Done — but narrower than first framed: the
  *button base* is entangled with topnav's intentional `.topnav-right #themeToggle` override
  (30px, 5px radius, transparent), so it stays per-page. The genuinely-duplicated **sun/moon icon
  visuals** (byte-identical across all 4 topnav pages) moved into `topnav.css` once; removed from
  studio/screenspace/transcripts/workflows CSS. `gallery.css`/`viewer.css` keep their copies (no
  topnav.css). Updated `tests/test_workflows_frontend_source.py::test_theme_toggle_icons_styled`
  to the new location. Net −94 lines.
- [ ] _Opportunistic, not standalone:_ raw-px spacing/font-size/transition → tokens **only inside
  files a carve already touches**. There are ~200 such values; a dedicated mass-conversion is out
  of scope and against the "don't bulk-rewrite" guidance.

---

## Descoped (considered, not recommended now)

- **Splitting large CSS files** (`screenspace.css` 4041, `studio.css` 3004, …) into multiple
  linked stylesheets. No build step; `viewer.css` is inlined into offline exports by `viewer.py`;
  multi-`<link>` splitting cuts against the current single-file-per-page convention. Low payoff,
  real churn. Skip unless a concrete maintenance pain emerges.
- **Unifying the two color systems** — see Theme B note. Risk > reward.

## Suggested execution order

1. C1 (trivial warm-up, isolated). 2. A1 → A2 → A4 (lowest-risk carves). 3. B1.
4. A5, A3. 5. A6, A7, B2, C2, C3. Each lands green via `/check` before the next.
