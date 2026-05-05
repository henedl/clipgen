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
| 3 | Convergence + Metadata + bottom-strip polish | Apply primitives + new visualisations (`SwimLane`, `KpiCard`, `CoverageMatrix`) inside Studio's Convergence and Metadata sub-tabs; unify bottom-strip cards with the new primitives; final dead-CSS sweep | Pending |
| 4 | Screenspace | Safe layout: viewer + scrub timeline + detector panel + results panel; refresh detector iconography with design lead | Pending |
| 5 | Transcripts | Safe layout: editable summary + sticky PiP video + annotation popover; cross-ref with sheet rows / screenspace events | Pending |
| 6 | Polish | Animation finalization with design lead; viewer/gallery surfaces audit; light-theme parity check across all surfaces; **unify card-size tokens across `.clip-card` / `.transcript-card` / `.queue-card` (intake, artifact, reel, stash thumbnails) so the bottom strip and intake cards share one width scale** | Pending |

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
