# Frontend refactor — plan

Planning doc. No implementation in this file. Scope is `assets/web/` and how Flask / `viewer.py` deliver it.

**Context:** The May 2026 redesign ([REDESIGN-PLAN.md](../REDESIGN-PLAN.md)) shipped tokens, TopNav, start overlay, settings, and Studio primitives. Remaining maintenance cost concentrates in three page monoliths and repeated cross-page patterns.

**Related plans:**

- [PERFORMANCE-PLAN.md](PERFORMANCE-PLAN.md) — Phase 4 (Studio grid, NDJSON intake) overlaps structural/perf work here
- [MULTIPLAYER-PLAN.md](MULTIPLAYER-PLAN.md) — future presence layer; plan module boundaries before adding shared state
- [FRONTEND-PLAN.md](FRONTEND-PLAN.md) — superseded preparatory research (pre-redesign); kept for history

---

## Summary

The frontend is a **vanilla JS/CSS stack** (~35k lines in 34 files under `assets/web/`) with a clear shared foundation (`tokens.css`, `utils.js`, TopNav, start overlay, settings). **Three page bundles** (`screenspace.js`, `studio.js`, `transcripts.js`) hold ~60% of all JavaScript (~13,900 lines) and drive most refactor value.

**Already working well:**

- Thin server / thick client for rendering
- `utils.get_frontend_config()` → `"config"` in API responses; `tests/test_shared_constants.py`
- ES5 + `.then()` everywhere (no `async/await`)
- Shared chrome: `ClipgenTopNav`, `ClipgenStartOverlay`, `openSettingsModal`
- Studio sub-tabs split (`metadata.js`, `convergence.js`) but coupled via `window._studioState`

**Main pain points:**

1. Copy-pasted **export quick action**, **toast**, **video speed controls**, **polling**, and **timeline rendering** across pages
2. **Four separate `renderTimeline()` implementations** (only amplitude bands shared)
3. **Convention debt:** inline SVG in HTML, raw `px`/`rem` in page CSS, dual button systems (`.cg-btn` vs `.btn`)
4. **Implicit globals** between Studio and sub-tabs (`window._studioCluster*`)
5. **Dead / partial assets:** `card-scrubber.js` parked; `primitives.css` on Screenspace/Transcripts without `primitives.js`

**Stack constraints (do not violate):**

- No React, TypeScript, bundler, or build step
- No backwards-compat shims for client persisted state
- No headless browser CI for UI — human verification + existing pytest for inlined viewers / shared constants

---

## Inventory

### JavaScript (approx. lines)

| File | Lines | Role |
|------|------:|------|
| `screenspace.js` | 6,173 | Canvas, tasks, timeline, regions |
| `studio.js` | 4,392 | Sheet, queues, intake |
| `transcripts.js` | 3,317 | Editor, search, video, agents |
| `viewer.js` | 1,890 | Exported timeline viewer |
| `metadata.js` | 1,542 | Studio Metadata tab |
| `start-overlay.js` | 1,102 | Folder / spreadsheet picker |
| `convergence.js` | 1,020 | Studio Convergence tab |
| `utils.js` | 965 | Shared globals |
| `primitives.js` | 754 | DOM factories (Studio only) |
| Other chrome / parked | ~1,000 | topnav, settings, card-scrubber, gallery |

### CSS (approx. lines)

| File | Lines | Notes |
|------|------:|-------|
| `screenspace.css` | 2,750 | Many raw px/rem |
| `studio.css` | 2,408 | Many raw px/rem |
| `transcripts.css` | 1,385 | |
| `start-overlay.css` | 1,346 | |
| `viewer.css` | 1,230 | |
| `tokens.css` | 576 | Canonical tokens |
| `primitives.css` | 669 | Some raw values despite token header |

### Delivery paths

| Path | Mechanism |
|------|-----------|
| Live Studio / Screenspace / Transcripts | `register_static_routes()` + ordered `<script>` tags |
| Exported viewer / gallery | `viewer.py` inlines CSS/JS + `window.CLIPGEN_DATA` |

---

## Duplication map

| Pattern | Locations | Notes |
|---------|-----------|--------|
| Export quick action | `studio.js`, `screenspace.js`, `transcripts.js` | ~40 lines × 3; raw `fetch` not `apiPost` |
| Toast markup + CSS | All three page HTML + CSS | Screenspace: opacity; others: `display: none` |
| Video speed cycle | `screenspace.js`, `transcripts.js` | Same handler; speeds differ by design |
| `renderTimeline()` | screenspace, transcripts, viewer, convergence | Four implementations |
| Icon masks | `maskIconStyle`, `svgMask`, page `iconHTML` | Same technique, three APIs |
| Canvas theme colors | `refreshThemeColors()` vs inline `getCSSVar` | Unify read/cache |
| Polling | Studio intake, Transcripts agents, Screenspace tasks | `POLL_INTERVAL` in utils underused |
| Card scrubber | `card-scrubber.js` + inline in `viewer.js` | Parked per ARCHITECTURE.md |
| HTML `<head>` | 7 HTML files | Favicon + fonts repeated |

### Studio coupling

- `window._studioState` — contract for `metadata.js` / `convergence.js`
- `window._studioClusterIntakeEvents` etc. — clustering should move to shared module

---

## Convention gaps (AGENTS.md)

| Rule | Gap |
|------|-----|
| Icons via `mask-image` from `assets/icons/` | Inline `<svg>` in several HTML files; data-URI SVG in `viewer.css` |
| Tokens for spacing/type in new/touched CSS | Hundreds of raw values in page stylesheets |
| No duplicate Python/JS constants | Intentional mirrors in `utils.js` + tests — maintenance surface only |
| `primitives.js` only where factories needed | Screenspace/Transcripts load `primitives.css` without JS |

Decorative SVG in `start-overlay.html` may remain an explicit exception after icon pass.

---

## Refactor tiers

### Tier A — Quick wins (low risk)

| ID | Work | Impact |
|----|------|--------|
| A1 | `export-actions.js` — shared export + TopNav; use `apiPost`/`apiGet` | ~120 lines DRY |
| A2 | Toast in shared CSS (`primitives.css` or `topnav.css`); align show/hide | Visual consistency |
| A3 | `createPoller(fn, ms)` in `utils.js` | Fewer interval bugs |
| A4 | `getCanvasThemeColors()` in `utils.js` | One theme path for canvases |
| A5 | Drop unused `primitives.css` from Transcripts; audit Screenspace | Smaller load |
| A6 | Wire `card-scrubber.js` into viewer **or** delete module + inline dup | Remove dead code |

**First PR suggestion:** A1 + A2.

### Tier B — Medium extractions

| ID | Work | Impact |
|----|------|--------|
| B1 | `video-controls.js` with per-page `speeds` config | DRY two players |
| B2 | `intake-cluster.js` — drop `window._studioCluster*` | Decouple Studio sub-tabs |
| B3 | Single `iconMask(name, basePath?)` | Fewer icon URL bugs |
| B4 | Shared timeline core (ruler/ticks only) | Partial DRY; markers stay per-surface |
| B5 | `.btn` → `.cg-btn` or document dual system | Design consistency — **defer** |

### Tier C — Structural

| ID | Work | Impact |
|----|------|--------|
| C1 | Split page monoliths into feature scripts (ordered tags, no bundler) | Maintainability |
| C2 | Studio grid perf ([PERFORMANCE-PLAN.md](PERFORMANCE-PLAN.md) §4.1) | Large-sheet UX — **profile first** |
| C3 | NDJSON intake streaming (PERFORMANCE-PLAN §4.2) | Per-item progress |
| C4 | Inline SVG → mask-image (viewer + Screenspace toolbars first) | Convention compliance |
| C5 | Token migration on touched CSS files | Incremental, not big-bang |
| C6 | Flask-injected HTML head partial | DRY favicon/fonts |

#### Suggested split order for C1

1. **`screenspace.js`** (largest) → `state.js`, `canvas.js`, `timeline.js`, `tasks.js`, `workflow.js`
2. **`studio.js`** → sheet, queues, intake
3. **`transcripts.js`** → segments, search, agents

Globals stay on `window`; script order in HTML documents dependencies.

---

## Recommended waves

| Wave | Items | Outcome |
|------|-------|---------|
| **1** | A1, A2, A3, A4 | Chrome DRY, no feature change |
| **2** | A5, A6, B1, B2 | Small modules, clearer boundaries |
| **3** | B3, C4 (incremental), C5 (opportunistic) | Convention cleanup |
| **4** | C1 (screenspace first), C2 after profiling | Structural + perf |

---

## Explicitly out of scope

- React / TypeScript / bundler
- Backwards-compat for old exported HTML viewers
- Headless browser CI
- Merging Quick Actions with Studio sub-header (gallery drawer flow — see REDESIGN-PLAN session 1)
- Screenspace flat-scroll, Transcripts sticky PiP (redesign deferred)
- SSE task progress ([PERFORMANCE-PLAN.md](PERFORMANCE-PLAN.md) dropped for now)
- Full timeline abstraction (B4) unless timeline bugs force it

---

## Verification

| Change | Checks |
|--------|--------|
| Shared JS extract | `tests/test_studio_frontend_source.py`, `test_viewer_inline.py`, `test_shared_constants.py` |
| Export → `apiPost` | Export/start API tests if applicable |
| CSS-only | Human: `/studio/`, `/screenspace/`, `/transcripts/` — toast, theme, export menu |
| Monolith split | Above + HTML script order review |
| Viewer paths | `test_viewer_inline.py` |

Pre-commit: [agents/skills/check/SKILL.md](../agents/skills/check/SKILL.md). Docs-only PRs do not bump `build/VERSION`.

---

## Checklist (implementation tracking)

Use this when executing waves; check items in PR descriptions.

### Wave 1

- [x] A1 `export-actions.js` + TopNav wiring on all three surfaces
- [x] A2 Shared toast CSS + consistent hide behavior
- [x] A3 `createPoller` in `utils.js`; adopt in hot paths
- [x] A4 `getCanvasThemeColors` in `utils.js`; Screenspace + Transcripts canvases

### Wave 2

- [x] A5 CSS include audit (Transcripts/Screenspace) — toast moved to `topnav.css`; `primitives.css` dropped from both
- [ ] A6 Card-scrubber: integrate or remove — deferred, left parked (unused; re-hook when filmstrip thumbnails return)
- [x] B1 `video-controls.js` — shared `nextSpeed`/`applyPlaybackRate`
- [x] B2 `intake-cluster.js` — `window.ClipgenIntakeCluster`; drops `window._studioCluster*`

### Wave 3

- [ ] B3 Unified icon helper
- [ ] C4 SVG → mask (viewer, Screenspace toolbar)
- [ ] C5 Token sweep on files touched in same PRs

### Wave 4

- [ ] C1 `screenspace.js` module split
- [ ] C1 `studio.js` module split
- [ ] C1 `transcripts.js` module split
- [ ] C2 Studio grid: profile → incremental filter → chunk render → virtualize only if needed
- [ ] C3 NDJSON intake (coordinate with server.py)

---

## Document history

| Date | Change |
|------|--------|
| 2026-05-19 | Initial plan from frontend refactor investigation (plan-only session) |
