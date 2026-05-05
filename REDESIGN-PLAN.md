# clipgen frontend redesign — running plan

This document is the living plan for the clipgen frontend redesign that started in May 2026. It covers all three primary surfaces (Studio, Screenspace, Transcripts) plus shared infrastructure (tokens, TopNav, primitives).

The handoff bundle from the design lead lives in `redesign/` (gitignored — local-only working copy):

- `redesign/README.md` — handoff overview, locked decisions, design tokens, component anatomy, state model
- `redesign/AGENTS.md` — coding rules for AI agents working on the port
- `redesign/source/Studio v2.html` — interactive React+Babel prototype (single file, ~3300 lines)
- `redesign/screenshots/` — reference renders for Studio, Screenspace, Transcripts

The prototype is a **design reference**, not production code. Production stack stays vanilla JS / CSS / HTML — no React, no build tools.

## Locked design decisions (from `redesign/README.md`)

| Surface     | Setting                | Locked value     |
|-------------|------------------------|------------------|
| Studio      | Density                | Standard (`rowPadV` 7, `cellW` 56, font 12.5, `headPadV` 8) |
| Studio      | Sheet read mode        | Cells (participant grid with timestamp chips) |
| Studio      | Severity rendering     | Pill (colored pill with dot) |
| Studio      | Convergence viz        | Expressive (opacity + saturated dots) |
| Studio      | Screenspace card size  | Large (lg) |
| Screenspace | Layout                 | Safe (viewer-left + detector/results-right) |
| Transcripts | Layout                 | Safe (centered single column ~720 px) |
| Transcripts | Summary mode           | Editable |
| Transcripts | PiP mode               | Sticky (sticks to top on scroll) |
| Theme       | Default                | Dark (light is opt-in via toggle) |
| Avatar slot | Replaces with          | Theme toggle |

**Do not build the alternative variants** (`bold` Screenspace, `detached` PiP, `quiet` convergence, `static` summary, etc.) — only the locked variants ship.

## Multi-session roadmap

| # | Session | Scope | Status |
|---|---|---|---|
| 1 | Foundation | Tokens (dark default + light alt), unified TopNav across all three surfaces, Studio Sheet "cells" rendering with severity pills, sub-tab bar with slide-fade transitions, bottom artifacts strip drag-resize | **Done** |
| 2 | Studio primitives + sidebar + intake refresh | Extract `FilterChip` / `ParticipantPill` / `DensityTimeline` / `SparkBars` / `ClipCard` / `TranscriptCard` / `Btn` into shared `assets/web/primitives.{js,css}`. Add Sheet sidebar (VIEWS / CATEGORIES / PARTICIPANTS). Tighten sheet-cell padding + restyle hover-expand overlay. Refresh Screenspace Intake + Transcript Intake interiors. Restructure bottom strip into ARTIFACTS / REEL columns. Drop Studio-internal centering for full-bleed layout. | **Done** |
| 3 | Convergence + Metadata + dead-CSS sweep | Apply primitives + new visualisations (`SwimLane`, `KpiCard`, `CoverageMatrix`) inside Studio's Convergence and Metadata sub-tabs; trim Metadata to default 5 KPIs + activity histogram + coverage matrix with 7 detail sections behind a "Show details" expander; convert Convergence detail panel to popover anchored to clicked cluster; final dead-CSS sweep. **Bottom-strip cards intentionally deferred to pass 6.** | **Done** |
| 4 | Screenspace | Safe layout: viewer + scrub timeline + detector panel + results panel | Pending |
| 5 | Transcripts | Safe layout: editable summary + sticky PiP video + annotation popover; cross-ref with sheet rows / screenspace events | Pending |
| 6 | Polish | Animation finalization with design lead; viewer/gallery surfaces audit; light-theme parity check across all surfaces; **unify card-size tokens across `.clip-card` / `.transcript-card` / `.queue-card` (intake, artifact, reel, stash thumbnails) so the bottom strip and intake cards share one width scale**; **layout restructure for floating-nav scroll-under** (the visual glass treatment landed in pass 3, but Studio / Screenspace / Transcripts use `body { height: 100vh; overflow: hidden }` with internal flex layouts so no content actually scrolls under the topnav today — restructure scrollable panels so their scroll containers extend to viewport top with internal padding pushing visible content below the chrome, and bump sticky offsets on sheet headers + intake density host + any other sticky-inside-scroll element to clear the 48 + 44 px chrome strip; reference: frameset.app/search) | Pending |

## Implementation rules (carry-overs from `redesign/AGENTS.md`)

- Vanilla JS only — no React/JSX, no `useState`/`useRef`, no inline `style={{}}` objects, no Babel
- Match the prototype's component names (`TopNav`, `FilterChip`, `ParticipantPill`, `DensityTimeline`, `ClipCard`, `TranscriptCard`, `SwimLane`, `AnnoPopover`, `SsViewer`, `SsScrubTimeline`, `SsDetectorPanel`, `SsResultsPanel`) so design and engineering can talk
- Tokens in `:root` + utility classes; component-specific styles alongside the component when they exceed utilities
- Use `oklch()` for category hues at runtime (chip rings, dots, density bars), not pre-flattened hex
- Tabular numerals on every timestamp via `font-feature-settings: 'tnum'` on `.cg-mono`
- No `scrollIntoView` — use `element.scrollTo(...)` or `scrollTop` directly (existing convention)
- Hit targets ≥ 28 px on clickable elements
- Top-tab routing uses real URLs (`/studio/`, `/screenspace/`, `/transcripts/`) — not in-page state

## Session log

### Session 1 — foundation pass (2026-05-04)

**Branch:** `henedl/redesign-pass-1`

**Shipped:**

- **Tokens** — `assets/web/tokens.css` carries the prototype's dark-default palette (`--bg`, `--fg`, `--accent`, `--severity-*`, `--cell-data*`, `--shadow-pop`, etc.) plus Inter / JetBrains Mono fonts. Legacy `--color-*` and `--sev-*` names alias through `var()` so existing pages keep rendering. `@media (prefers-color-scheme: dark)` and `html[data-theme="dark"]` blocks deleted; light is now opt-in via `html[data-theme="light"]` / `.theme-light`. `utils.js` theme toggle simplified to dark-default; `applyStoredThemePreference` always sets `data-theme="dark"|"light"`.
- **TopNav** — new shared module `assets/web/topnav.{js,css}` mounted on Studio, Screenspace, and Transcripts via `<topnav-mount data-frontend="...">`. Cluster: logo + 3 center tabs + Quick Actions menu (with outside-click + Esc) + Log icon + tooltip toggle (Studio/Transcripts only) + Settings + theme toggle (in place of avatar) + version pill. Per-page `frontend-switcher` markup removed. Studio populates Quick Actions in `initTopNavActions()` (Open Timeline / Open Gallery / Refresh sheet / Filter rows). Screenspace and Transcripts have empty Quick Actions for now (the trigger auto-hides).
- **Sub-tab bar** — Studio's five preview tabs moved out of `.sheet-preview-header` into a new `#studioSubheader` (left: study + version meta; center: tabs; right: Timeline Viewer / Gallery / Filter / Refresh). Underline-on-active styling matches the prototype. Tab-content slide-fade transition (180ms ease-out, translateX 12 → 0, opacity 0 → 1) wired in `syncPreviewTab()` via `.tab-slide-enter` class + double-rAF.
- **Sheet "cells" rendering** — `renderDataRow()` in `studio.js` now wraps timestamps in `<span class="ts-chip cg-mono">` chips on `--cell-data` background (uniform across severities, per locked variant). Severity rendered as `<span class="sev-pill"><dot/><label/></span>` — pill with severity-tinted background. Severity-color tinting on cells removed. Long observations get CSS ellipsis (no JS truncate) with full text in `title` attr.
- **Bottom strip** — drag-resize handle bounded to [60, 560] px (was unbounded above); `state.dividerOffset` and `state.bottomCollapsed` persist to `localStorage["clipgen-studio-bottom-h"]`. Existing double-click collapse retained. The existing visual handle (4px pill at top-center) already matches the prototype.
- **Category hue helper** — new `categoryHue(label)` and `categoryColor(label, alpha)` in `utils.js` returning the prototype's hue table via `oklch()` at runtime. Unblocks session 2 primitives.
- **Persistence** — bottom-strip height + collapse state survive reload. Theme preference unchanged (still `clipgen-theme`).

**Verified:**

- `uv run clipgen.py --screenspace -i /tmp/empty -o /tmp/empty` serves the Screenspace page; TopNav, tokens, and topnav.css load cleanly. All icon and asset routes return 200.
- `uvx ruff check` and `uvx ruff format --check` both clean (50 files already formatted).
- `uv run --extra dev pytest -c tests/pytest.ini` — 733 tests pass, 0 regressions.
- `uvx ty check` shows 13 diagnostics, all pre-existing "Cannot resolve imported module pytest" environment issues unrelated to this PR (verified via `git stash`).

**Deferred to later sessions:**

- Studio sub-tab interiors (Screenspace Intake, Transcript Intake, Convergence, Metadata) — tab bar wired with slide-fade, but interior content keeps current rendering. Session 2.
- Screenspace and Transcripts surface redesigns — both pages got the new TopNav, but layouts and primitives unchanged. Sessions 3 and 4.
- Sticky PiP, editable summary, annotation popover, expressive SwimLane. Session 4.
- Detector iconography revision (handoff flagged as "not final"). Session 5.
- Shared primitives extraction (decided to inline within Studio first; primitives extracted in session 2 once we know the real shape from real data).
- `cellColorCoding` setting toggle now no-ops (the per-severity cell-tinting it drove was removed for the locked "cells" variant). Either remove the setting in session 2 primitives pass or repurpose for a different cell-mode.
- Dead CSS to clean up: `#studioHeaderTop`, `#headerMeta`, `#headerTop`, `.sheet-preview-header` — applied to nothing post-redesign, can be deleted in session 2.

**Surprises / lessons:**

- The existing studio.js drag-resize divider was already most of the way to the redesign spec — just needed the 60/560 bound and localStorage persistence. The redesign README's "drag handle, range 60–560, double-click collapse" reads as a new feature but ~95% of the work was already in place from prior commits.
- `prefers-color-scheme: dark` handling complicated the theme flip more than expected. Cleanest fix was to drop OS-preference detection entirely (matching the prototype's strict dark-default) rather than try to honor it as a fallback. Users who liked auto-light on light-OS machines will need to manually toggle once.
- Page-specific `#themeToggle` styling (`width: 32px; border: 1px solid; background: surface-alt`) had to be overridden by `.topnav-right #themeToggle` (specificity 0,1,1 beats 0,1,0). The sun/moon glyph CSS lives in studio/screenspace/transcripts CSS — those still apply correctly inside the TopNav cluster because the icon class names match.
- Quick Actions menu items kept simple this session. Many would-be items (Open Timeline / Gallery) duplicate visible buttons in the Studio sub-header, which is fine for v1 — duplication beats refactoring the gallery drawer's two-state click flow this session.

### Session 2 — primitives, Sheet sidebar, intake refresh, bottom-strip + full-bleed (2026-05-04)

**Branch:** `henedl/redesign-pass-2`

**Shipped:**

- **Primitives** — new `assets/web/primitives.{js,css}` exports `FilterChip`, `ParticipantPill`, `DensityTimeline`, `SparkBars`, `ClipCard`, `TranscriptCard`, `Btn` as DOM-element factory functions on `window.ClipgenPrimitives`. Hue resolution falls back through `categoryHue(label)` from `utils.js`; active state uses inline `oklch()` style for chip fg/bg/border. Loaded by `studio.html` (`primitives.css` between `topnav.css` and `studio.css`; `primitives.js` after `utils.js`, before `topnav.js`). Generic blueprint static-route handles `/screenspace/primitives.{css,js}` and `/transcripts/primitives.{css,js}` 200s for sessions 4+5 with no server.py changes.
- **Sheet sidebar** — collapsible `<aside id="studioSidebar">` with three sections (VIEWS / CATEGORIES / PARTICIPANTS). VIEWS chips drive `state.filters.sevMin/sevMax` via a small `SIDEBAR_VIEWS` map (All / Highlights / Positive / Medium / High). CATEGORIES toggle into `state.filters.categories`. PARTICIPANTS toggle into `state.sidebarParticipants`, which `renderGrid()` consumes by filtering `d.participants` into a local `visibleParticipants` array used for `<col>`, `<th>`, and `<td>` iteration. Sidebar collapse persists to `localStorage["clipgen-studio-sidebar-open"]`; `body[data-active-tab]` toggles the sidebar visibility per sub-tab.
- **Sheet-cell conform** — `.ts-cell` padding tightened to `0 4px` with `height: 30px` so chips fit cells without floating gaps. `.ts-chip` font-family promoted to `var(--font-mono)` (was missing). Floating expand-overlay `.ts-cell-float` rewritten to mirror chip visuals (same `--cell-data` bg, `--cell-data-fg`, `cg-mono`/tnum, 11px, radius 3, 1px `--border-strong`); `showFloat()` no longer overrides `backgroundColor` from cell's computed style; toggles a `.has-text` modifier on the float for invalid-timestamp cells.
- **Full-bleed layout** — removed the centering `max-width: var(--layout-max-width); margin: 0 auto;` on `#sheetPreview`, `#panelDivider`, `#dropAreas`, `#reelArea`, `#stashedReelsArea`, `#stashedArtifactsArea`. Studio surface and bottom strip now run edge-to-edge.
- **Sheet sub-header layout** — `#sheetPreview` reflowed to flex-row (`#studioSidebar` + `#sheetMain`); `#sheetMain` carries the existing filter-bar + sheet-grid (and intake/convergence/metadata panels) and absorbs the inner padding.
- **Screenspace Intake refresh** — replaced static detector pill markup with dynamic `createFilterChip` row (driven by an `INTAKE_DETECTORS` constant + per-detector counts), participant pill row with `createParticipantPill`, and the `<canvas id="intakeTimeline">` with a div-based `createDensityTimeline` host. Cards switched to `createClipCard` with `size: 'lg'`. Lazy-loaded thumbnails (`ssObserveThumb`) and cross-ref badges (`buildXrefBadges`) preserved via post-create DOM injection on `.clip-card-thumb`. Drag/click/hover handlers and `.queue-card.intake-queue-card` class kept for downstream selectors.
- **Transcript Intake refresh** — same anatomy with `createTranscriptCard` (timeRange = `mm:ss–mm:ss`, text from cluster). New `buildTrIntakeCategoryPills()` consumes the `TR_INTAKE_CATEGORIES` map. `buildTrIntakeParticipantPills` rewritten on top of `createParticipantPill`. Tooltip flow (`#trIntakeTooltip`) untouched.
- **Bottom-strip refresh** — wrapped `#dropAreas` + `#stashedArtifactsArea` in `#artifactColumn`, `#reelArea` + `#stashedReelsArea` in `#reelColumn`, both inside a new flex-row `#bottomColumns` with a `--hairline` divider between columns. ARTIFACTS now renders as a CSS grid (`auto-fill minmax(160px, 1fr)`); REEL stays a horizontal scroll. Section H3 labels restyled to uppercase 12.5px with mono count badges (matches prototype's "ARTIFACTS"/"REEL" labels).
- **Cleanup** — deleted dead CSS blocks (`#studioHeader`, `#studioHeaderTop`, `#headerMeta`, the `.sheet-preview-header` carry-over comment) and the no-longer-used `.intake-filter-pills` / `.intake-filter-det` / `.intake-det-icon` / `.intake-filter-participant-pills` / `.intake-filter-participant` rules. Removed canvas helpers `sizeIntakeCanvas`, `renderIntakeTimeline`, `intakeHitTest`, `_intakeHitRects`, `sizeTrIntakeCanvas`, `renderTrIntakeTimeline`, `trIntakeHitTest`, `_trIntakeHitRects`, `intakeComputeTickInterval`, and `hexToRgba` (canvas-only). Dropped the no-op `cellColorCoding` setting from `state`, `studio.js` settings sync, `server.py` `/api/sheet` payload, and `config.py` (constant + description + Studio settings registration).

**Verified:**

- `node -c` syntax-checks clean for `studio.js` and `primitives.js` after each edit.
- Quality gates pending in task 12 (run before commit).

**Deferred to later sessions:**

- Studio Convergence + Metadata sub-tab interiors — explicitly punted to session 3 per user direction. Will introduce `SwimLane`, `KpiCard`, `CoverageMatrix` primitives.
- Bottom-strip cards still use the legacy `.queue-card` shape for queued/generating/done state. Unification with `.clip-card`/`.transcript-card` deferred (separate state model).
- Screenspace surface (Safe layout) — session 4.
- Transcripts surface (Safe + Editable summary + Sticky PiP) — session 5.
- Detector iconography revision (handoff flagged as "not final") — session 6 polish.

**Surprises / lessons:**

- Generic blueprint route `/<path:filename>` already serves any file in `assets/web/`, so `primitives.{js,css}` is reachable from `/studio/`, `/screenspace/`, and `/transcripts/` without explicit per-blueprint registration. The plan's "mirror topnav routes" line was redundant.
- `.ts-cell-float` JS was overriding the cell's computed `backgroundColor` inline at every show — that line silently fought every cell-conform CSS attempt in session 1. Removing the inline override was the actual fix; the new chip-styled CSS already had the right paint.
- The intake render had a long-tail of canvas-coupled state (`_intakeHitRects`, `intakeHitTest`, hover-driven `renderIntakeTimeline()` repaints) that all dissolved with the SVG/div primitive. The same code shape repeated for transcript intake — both became ~120 lines lighter.
- Removing `state.cellColorCoding` left the function `_findSetting("STUDIO_SHEET_CELL_COLOR_CODING")` dangling in `applySettingsFromAPI`; deleted the lookup in the same pass and dropped the constant from `config.py` so settings UI no longer surfaces a knob with no effect.

**Iteration lessons (from the post-pass review cycle):**

These came out of repeated visual-comparison passes against the prototype after the initial port. Recording them so future sessions don't repeat the loop.

- **`min-width: 0` / `min-height: 0` on every flex/grid container in the chain.** A child's intrinsic content size (`min-*: auto`) blocks the parent from shrinking below it, which silently breaks flex layouts the moment a sibling expands. We hit this 3× in this session:
  - `#sheetMain` → `#sheetGrid` had a fixed-layout table forcing the parent past its flex allocation when the sidebar opened.
  - `#bottomColumns` → columns squished cards once the artifact list exceeded one row.
  - `#sheetPreview` itself with the new `flex: 1 1 auto` needed `min-height: 0` to share viewport with the bottom panel.
  - **Rule of thumb**: when adding a flex/grid layer, immediately add `min-width: 0; min-height: 0` to every descendant in the same flex chain unless one specifically needs to grow.
- **Avoid `width: auto` on `<col>` under `table-layout: fixed`.** When other tracks sum past the container width, the auto track collapses to 0 and content visibly bleeds into adjacent cells. Always give every `<col>` an explicit width and let `#sheetGrid { overflow: auto }` handle horizontal scroll.
- **Overflow signals move down with the chip.** When you clamp inline-block content with `max-width: 100%; overflow: hidden; text-overflow: ellipsis`, the parent `td.scrollWidth` equals `td.clientWidth`. Hover-overflow detection has to read the chip's own `scrollWidth`, not the cell's. Same applies to any "is this clipped?" UI.
- **Reset flex/grid defaults explicitly.** `align-items: stretch` (flex) and `align-content: stretch` (grid, when there's free space) are invisible stretchers. Reel cards growing to fill the column was `align-items: stretch`; squished artifact card rows were Chrome distributing free space across grid rows. Set `align-items: flex-start` / `start` and `align-content: start` plus `grid-auto-rows: max-content` whenever the layout should be content-sized.
- **Bottom-panel drag = explicit pixel height + flex:1 upper pane.** The first attempt drove an offset on the upper pane via `max-height`; that fought the `flex: 1 1 auto` upper pane and left the bottom panel un-resizable. Rewriting to set `#bottomPanel { height: ${state.bottomH}px }` directly and letting the upper pane absorb the remainder via flex was simpler and just worked.
- **Primitive interactivity surface = options + imperative methods.** `createDensityTimeline` returns the element AND exposes `el.update(events, marker)` + `el.setHovered(idx)`. Callers wire bidirectional hover (cards ↔ bars) through `setHovered`, and accept callbacks (`onBarMouseEnter / Leave / Click`) via opts. Either alone is insufficient: callbacks let the bar drive card state, methods let the cards drive bar state.
- **Static `[data-icon]` markup needs a sweep.** `createBtn` sets `mask-image` inline on the icon span. For static HTML buttons (the bottom-strip toolbar) we added `applyDataIconMasks()` at `DOMContentLoaded` that walks every `[data-icon]` and applies the same mask-image. New static markup that uses `data-icon` should rely on this sweep.
- **Two-step "open form / confirm" buttons → modal.** The old Gallery button toggled a slide-down drawer and changed its own label to "Confirm" mid-flow — that pattern can't be triggered from a TopNav Quick Actions menu. The replacement is a centered modal (`#galleryOverlay`) with explicit Cancel + Confirm buttons. Same pattern is the right answer if more sub-header buttons get demoted to Quick Actions.
- **`.queue-card` width still wins over `.clip-card` in the bottom strip** (140 px vs 180 px) because both classes apply at equal specificity and `studio.css` loads after `primitives.css`. Cards in the bottom strip therefore render at the queue-card width even though they include the `.clip-card` class. Logged as a session-6 polish item ("unify card-size tokens").
- **Don't override CSS-driven background inline.** `showFloat()` was setting `cellFloat.style.backgroundColor = getComputedStyle(td).backgroundColor` every show, which silently overrode every CSS attempt to give the float its own chip styling. Remove inline overrides as soon as the CSS rule has the right paint.
- **Reference screenshot coverage matters.** The first comparison pass missed Tr-Intake / Convergence / Metadata gaps because the design handoff only included Studio-Sheet / Screenspace / Transcripts main screenshots. The remaining sub-tab states had to be reverse-engineered from `redesign/source/Studio v2.html`. For future handoffs, request a screenshot per locked sub-tab/state up front; otherwise budget time to study the prototype source.
- **Lazy thumbnail observers must include all card-thumb selectors.** `ssGetObserver` originally hardcoded `.queue-card-thumb`. New `.clip-card-thumb` / `.transcript-card-thumb` cards never had their IntersectionObserver fire and the gradient sat there forever. Centralized into a shared `SS_THUMB_SELECTOR` constant so the next card primitive can be added once.

### Session 3 — Convergence + Metadata + dead-CSS sweep (2026-05-05)

**Branch:** `henedl/redesign-pass-3`

**Shipped:**

- **Three new primitives** in `assets/web/primitives.{js,css}`: `createSwimLane(opts)`, `createKpiCard(opts)`, `createCoverageMatrix(opts)`. SwimLane renders a 22 px tick-axis + per-participant lane rows + dashed cluster bands + 10×12 px event markers tinted via `categoryHue(label)`; exposes `.update()`, `.setHovered(idx)`, `.setSelectedCluster(idx)` so callers can wire bidirectional hover + selection. KpiCard exposes `label / value / sub / accent / spark` slots; CoverageMatrix accepts `[{p, sheet, screenspace, transcript}]` and renders the prototype's hue-coded heat cells (Sheet=280, Screenspace=220, Transcript=145) plus a 6 px stacked distribution bar.
- **Convergence sub-tab** (`assets/web/convergence.{js,css}`) rebuilt around `createSwimLane`. Controls bar restyled to the prototype's compact mono inputs (Min participants / Window±s / Cluster s) plus a sort dropdown. Stream filter row swapped from per-stream-color pills to the prototype's inverted-pill "All / Sheet / Screenspace / Transcript" buttons followed by a 1 px divider and a `createFilterChip` row for event types. Render path reduced to a single `createSwimLane(...)` call seeded from `cvState.filteredEvents` + `cvState.convergenceZones`; cluster-callout cards render below the swim lane (`{n} participants converged within ~{X}s · {Y}h` + Pin button). Detail panel converted to a popover anchored at the clicked cluster's `(t0+t1)/2` mid-point with a small caret and outside-click + Esc dismiss; popover content uses `createParticipantPill` + `createBtn` for actions. Drag-to-select removed (cluster-only selection now). All canvas paths (`renderCanvases`, `renderSummaryLane`, `renderAllRowShading`, hit-rect arrays) deleted.
- **Metadata sub-tab** (`assets/web/metadata.{js,css}`) trimmed to the prototype's default view: 5 `KpiCard`s (Participants / Sheet observations / Screenspace events / Transcript moments / Project duration) each fed a `createSparkBars` driven from `cache.coverage` + `cache.histogramData.maxTime`; activity histogram block (card-wrapped `createSparkBars` with 7 mono tick labels below) replaces the canvas-based histogram; coverage matrix section uses `createCoverageMatrix`. The 7 lower-density sections (Per Event Type / Tr Category / Observation / Severity / Category Breakdown / Collisions / Sessions) move behind a single "Show details" expander (`#mdDetailsExpander`) with a chevron + count badge + persisted open/closed state in `localStorage["clipgen-studio-md-details"]`. Header bar Refresh / JSON / CSV swapped to `createBtn`; participant pills swapped to `createParticipantPill`. `renderHistogram`, `initHistogramHover`, `_histogramHitRects`, `MD_HISTOGRAM_HEIGHT` deleted; `metadata.css` dropped `.md-pill`, `.md-icon-refresh/-export`, `.md-coverage-table`, `.md-coverage-cell/-zero/-participant/-row` rules.
- **Dead-CSS sweep**: convergence.css fully rewritten — all canvas-coupled classes (`.cv-summary-canvas`, `.cv-row-canvas`, `.cv-summary-lane-wrap`, `.cv-axis-track`, `.cv-tick`, `.cv-axis-spacer`, `.cv-summary-track`, `.cv-event-marker`, `.cv-tracks-container`, `.cv-sub-track`, `.cv-time-axis`, `.cv-participant-row(s)`, `.cv-stream-toggle`, `.cv-event-type-pill`, `.cv-zone-tooltip`, `.cv-selection-overlay`, `.cv-drag-preview`, `.cv-sort-toggle`, `.cv-detail-btn`, `.cv-markers-dimmed`) gone. `studio.css` searched for any leftover `cellColorCoding` styling — none remained from pass 2.

**Verified:**

- `node -c` syntax-checks clean for `primitives.js`, `convergence.js`, `metadata.js`.
- Quality gates pending in task 8 (run before commit).

**Deferred to later sessions:**

- Bottom-strip card unification (visuals + width tokens) — pass 6, unchanged from earlier deferral but the user explicitly pulled this back into pass 6 mid-pass-3 to keep this session focused.
- Screenspace surface (Safe layout) — session 4.
- Transcripts surface (Safe layout + editable summary + sticky PiP) — session 5.
- Detector iconography revision — session 6 polish.
- Animation polish (TabSlide etc), light-theme parity check — session 6.

**Surprises / lessons:**

- **`#convergenceDetail` recycled as the popover host.** The existing `<div id="convergenceDetail" class="convergence-detail hidden">` slot in `studio.html` could be reused as-is; `ensurePopover()` finds it and rewrites its className to `cv-detail-popover hidden`. No HTML change needed.
- **`.area-controls` is shared across intake + transcript intake + convergence.** Initial convergence.css had a `.area-controls { padding: 10px var(--space-3); border-bottom: ... }` rule that clobbered the intake controls' padding. Scoped to `#convergenceControls` instead — the prototype's compact-input look only applies to convergence.
- **Dropping drag-to-select cuts ~150 lines.** The prototype only has cluster callouts as a selection mechanism; arbitrary time-range selection wasn't being used in the new design. Removing `_drag`, `onDragMousedown/move/up`, `timeFromMouseX`, the drag preview, and the summary-lane click handler simplified the file substantially. The keyboard shortcuts (Esc to dismiss, ←/→ to step between zones) keep working since they only depend on `cvState.selection.zone`.
- **SwimLane `.update()` does a full re-render.** Lane count is dynamic (per participant) and event positions depend on lane height, so partial updates aren't safe — wipe and rebuild on each `.update()`. `setHovered(idx)` and `setSelectedCluster(idx)` toggle classes on cached DOM refs, so per-frame hover stays cheap.
- **CoverageMatrix needs an adapter from `cache.coverage`.** The legacy metadata code stored coverage as `{pid -> {sheet, screenspace, transcript}}` but the primitive expects `[{p, sheet, screenspace, transcript}]`. Trivial adapter in `renderCoverageBody` — kept the legacy shape so other sections (collisions, session summary) keep working unchanged.
- **`metadataResize` now triggers a full re-render.** The old version called `renderHistogram()` which only repainted the canvas. With SparkBars + KpiCards, the cleanest re-flow on viewport change is to re-run `renderAll(mdState.cache)` since the primitives are pure DOM and the cost of re-rendering a few hundred elements is negligible.
- **Plan course-correction mid-flight.** Started planning bottom-strip card unification (visuals + width) pulled forward into pass 3, then user reversed: defer everything bottom-strip-related to pass 6. The plan file was updated before any bottom-strip code was written, so no rollback needed — but it reinforces the value of writing the plan early and getting buy-in before each scope commitment.
