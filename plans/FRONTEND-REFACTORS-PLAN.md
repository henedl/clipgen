# Frontend refactors & carveouts — plan

Status: **proposed** (2026-06-30). Tracks a sweep of refactor/carveout opportunities
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

- [ ] **A1. `studio-scrubber.js`** (~277 lines, `studio.js:2276–2553`). Card hover-scrub
  (sprite sheet + audio). Lowest risk: feature-gated on `state.cardScrubberEnabled`,
  read-only state, entry point `attachQueueScrubbers` is **already published** to STUDIO
  (`studio.js:5265`). Keep `buildQueueCardThumb`/`ssClearPending` access via STUDIO.
- [ ] **A2. `studio-trim.js`** (~354 lines, `studio.js:2729–3083`). Duration-badge trim
  popover + `buildCellOverrides()`. Self-contained; add hub delegators for `saveQueues()`/
  `renderQueue()`. **Must load before A3** (generate needs `buildCellOverrides`).
- [ ] **A3. `studio-generate.js`** (~393 lines, `studio.js:3743–4135`). Streaming artifact
  generation. Medium risk: needs delegators for `setArtifactGenerating`/`showResult`/
  `revealStatusOverlay`; ETA trackers + card-state painters **stay in hub** (shared with build).
- [ ] _Defer:_ **build** (`4136–4615`, reel/timeline/gallery interleaved — needs a separate
  reel-vs-viewer split first) and **stash** (`3240–3523`, coupled to `renderQueue`). Re-evaluate
  after A1–A3.

### Screenspace (`window.ClipgenScreenspace` / SS)

- [ ] **A4. `screenspace-overlay.js`** (~218 lines, `screenspace.js:2763–2980`). Pure canvas
  draw of regions/pending/template/heatmap/scene markers. Zero blockers — no writes; only
  reads state + pure helpers (`regionToPixels`, `getThemeColors`, `hexToRgba`). `renderOverlay`
  stays callable via a hub delegator (called from 6+ sites).
- [ ] **A5. `screenspace-timeline.js`** (~643 lines, `screenspace.js:2981–3623`). Canvas ruler,
  zoom/pan, scrubbing, markers, playhead, tooltips. Route `showSsTooltip`/`hideSsTooltip` and
  frame-load callbacks (`loadFrame`, `seekPlayhead`) through SS. Move `timelineZoom/Offset/
  Dragging`, `inMarker/outMarker`, `hoveredBoundaryTs` onto `state` if any hub site still reads them.
- [ ] **A6. `screenspace-overlay-interaction.js`** (~586 lines, `screenspace.js:1818–2323`,
  region draw/drag/resize state machine). Medium: must route `setTargetColor`/pipette
  activation from the color satellite (already deferred via SS) and keep `renderOverlay`/
  `renderRegionChips` as delegators.
- [ ] **A7. `screenspace-model-view.js`** (~514 lines, `screenspace.js:4347–4860`). Live
  preview overlay + layer UI. Publish `_collectPreviewParams` via SS (hub's `initRunButton`
  calls it). Self-contained sessionStorage.
- [ ] _Optional later:_ single-tool param builders (`3771–4346`, ~576) and run button
  (`4954–5296`, ~343). Lower ROI; param rehydration (`applyColorMode`/`applyNormalizeMode`)
  must stay hub-side for task restore.

**Stays in both hubs** (do not move): the shared `state` object, request-version/cache-
invalidation guards, participant routing, frame loading, the `DOMContentLoaded` init loop,
settings-sync, and STUDIO/SS namespace publication.

Expected outcome: studio ~5278 → ~4250 (A1–A3); screenspace ~5835 → ~4460 (A4–A7, ~1375 lines moved).

---

## Theme B — cross-cutting JS consolidation

- [ ] **B1. `createSSEStream()` in `utils.js`.** The EventSource→`onerror`→`createPoller`
  fallback is duplicated across 4 sites: `screenspace-tasks.js:1064` and `workflows-runs.js`
  (run `:77`, batch `:142`, discover `:1019`). Add `createSSEStream(url, {onMessage, onError, poll})`
  returning `{ close() }`; migrate all four. Modest, mechanical.
- [ ] **B2. Generalize the modal focus trap.** `studio.js` `openModalTrap`/`closeModalTrap`
  (`4627–4687`) is a reusable blocking-dialog trap; `transcripts.js confirmModelInstall`
  (`2180–2264`) hand-rolls the same backdrop/escape lifecycle. Promote a
  `openBlockingModal({onEscape, onBackdropClick})` to `utils.js`; migrate both. Leave singleton
  pickers (`color-picker.js`, `settings-modal.js`) owning their own lifecycle — they are not
  blocking dialogs.
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
