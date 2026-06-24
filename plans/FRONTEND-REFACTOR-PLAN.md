# Frontend refactor — plan

Planning doc. No implementation in this file. Scope is `assets/web/` and how Flask / `viewer.py` deliver it.

**Context:** The May 2026 redesign ([REDESIGN-PLAN.md](archive/REDESIGN-PLAN.md), now archived) shipped tokens, TopNav, start overlay, settings, and Studio primitives. Remaining maintenance cost concentrates in three page monoliths and repeated cross-page patterns.

**Related plans:**

- [MULTIPLAYER-PLAN.md](MULTIPLAYER-PLAN.md) — future presence layer; plan module boundaries before adding shared state
- [archive/PERFORMANCE-PLAN.md](archive/PERFORMANCE-PLAN.md) + [archive/PERFORMANCE-PLAN-2.md](archive/PERFORMANCE-PLAN-2.md) — Studio grid / NDJSON intake overlaps structural/perf work here (**archived** — verify what shipped before treating C2/C3 as pending)
- [archive/FRONTEND-PLAN.md](archive/FRONTEND-PLAN.md) — superseded preparatory research (pre-redesign); **archived** for history

---

## Summary

The frontend is a **vanilla JS/CSS stack** (~50k lines in 43 files under `assets/web/`: 22 JS / 14 CSS / 7 HTML) with a clear shared foundation (`tokens.css`, `utils.js`, TopNav, start overlay, settings). **Three page bundles** (`screenspace.js`, `studio.js`, `transcripts.js`) hold ~57% of all JavaScript (~18,500 lines) and drive most refactor value.

**Already working well:**

- Thin server / thick client for rendering
- `utils.get_frontend_config()` → `"config"` in API responses; `tests/test_shared_constants.py`
- ES5 + `.then()` everywhere (no `async/await`)
- Shared chrome: `ClipgenTopNav`, `ClipgenStartOverlay`, `openSettingsModal` (now `settings-modal.js`)
- **Wave 1–3 DRY shipped:** `export-actions.js`, `video-controls.js`, `intake-cluster.js` (`_studioCluster*` globals gone), shared toast (`topnav.css`), `getCanvasThemeColors`, unified icon-mask helpers in `utils.js`
- Studio sub-tabs split (`metadata.js`, `convergence.js`, each now with its own CSS) but coupled via `window._studioState`

**Main pain points (still open):**

1. ~~**Four separate `renderTimeline()` implementations**~~ 🟡 Ruler core shared (B4, 2026-06-24): the two canvas surfaces (screenspace / transcripts) share `niceTimeInterval`/`drawTimelineRuler` in `utils.js`; markers/bands/playhead stay per-surface, and the two DOM rulers (viewer / convergence) remain a separate model
2. ~~**Polling only half-unified**~~ ✅ Resolved: `createPoller` now drives every periodic poller (Studio, Screenspace, Transcripts); only the Ollama model-pull loop (Promise-based, custom cancel/miss-count) stays a raw `setInterval`
3. **Convention debt:** inline SVG ✅ closed (documented intentional exceptions) and raw `px` ✅ swept to tokens (2026-06-23); dual button systems (`.cg-btn` vs `.btn`) ✅ de-duplicated + documented as intentional (B5, 2026-06-24 — shared `.btn` base now in `tokens.css`; full merge onto `.cg-btn` is incremental)
4. **Page monoliths persist:** `screenspace.js` 7.5k (partially carved), `studio.js` 5.8k (state hub); `transcripts.js` ✅ carved (2026-06-24) into a ~2.4k hub + 5 `transcripts-*.js` satellites behind `window.ClipgenTranscripts`
5. **Dead / partial assets:** `card-scrubber.js` parked **and** duplicated inline in `viewer.js` (~~`<head>` favicon/fonts copy-paste~~ ✅ resolved 2026-06-23 — live pages now inject a shared `_head.html` partial)

**Stack constraints (do not violate):**

- No React, TypeScript, bundler, or build step
- No backwards-compat shims for client persisted state
- No headless browser CI for UI — human verification + existing pytest for inlined viewers / shared constants

---

## Inventory

### JavaScript (approx. lines, 22 files)

| File | Lines | Role |
|------|------:|------|
| `screenspace.js` | 7,498 | Hub: canvas, tasks, timeline, regions (+ satellites below) |
| `studio.js` | 5,813 | Sheet, queues, intake; state hub (`window._studioState`) for sub-tabs |
| `transcripts.js` | ~2,396 | Hub: `state`, xref index, `selectParticipant`, segments+marks editor, task poller, model-install (+ satellites below; `window.ClipgenTranscripts`/TS) |
| `transcripts-agents.js` | ~1,188 | Satellite: summary + citations + friction + panel tabs + heatmap toggle |
| `transcripts-video.js` | ~760 | Satellite: player + timeline canvas + seek + sync + PiP + `_drawFrictionBand` |
| `transcripts-pills.js` | ~660 | Satellite: participant pills + options popover + transcribe POST flow |
| `transcripts-search.js` | ~230 | Satellite: full-text search + results + mark-all + jump |
| `transcripts-corrections.js` | ~120 | Satellite: global find→replace corrections modal |
| `viewer.js` | 2,142 | Exported timeline viewer (+ inline card-scrubber dup) |
| `convergence.js` | 1,754 | Studio Convergence tab |
| `metadata.js` | 1,714 | Studio Metadata tab |
| `utils.js` | 1,499 | Shared globals + helpers (poller, icon masks, canvas colors) |
| `start-overlay.js` | 1,211 | Folder / spreadsheet picker |
| `settings-modal.js` | 966 | Shared tabbed settings modal (`openSettingsModal`) |
| `screenspace-multitool-params.js` | 882 | Screenspace satellite: multitool step/param UI |
| `primitives.js` | 763 | DOM factories (Studio only) |
| `dev-token-tweak.js` | 748 | Dev-only live token-tweak widget (gated; stripped from exports) |
| `screenspace-calibration.js` | 688 | Screenspace satellite: calibration strip |
| `topnav.js` | 343 | Shared top nav |
| `card-scrubber.js` | 298 | **Parked** (unloaded; dup'd inline in `viewer.js`) |
| `color-picker.js` | 266 | Reusable popover color picker (titlecards) |
| `gallery.js` | 214 | Gallery viewer |
| `screenspace-color.js` | 179 | Screenspace satellite: HSV color picker |
| `intake-cluster.js` | 106 | `window.ClipgenIntakeCluster` — Studio intake clustering |
| `screenspace-utils.js` | 100 | Screenspace satellite: pure, state-free helpers |
| `export-actions.js` | 63 | `window.ClipgenExportActions` — shared export quick action |
| `video-controls.js` | 36 | `window.ClipgenVideoControls` — shared speed cycle |

### CSS (approx. lines, 14 files)

| File | Lines | Notes |
|------|------:|-------|
| `screenspace.css` | 4,004 | ~273 raw px |
| `studio.css` | 2,969 | ~327 raw px |
| `transcripts.css` | 1,926 | ~170 raw px |
| `start-overlay.css` | 1,560 | |
| `viewer.css` | 1,450 | Inline data-URI SVG masks (intentional — exported/offline) |
| `metadata.css` | 718 | Studio Metadata tab |
| `tokens.css` | 657 | Canonical tokens |
| `primitives.css` | 613 | DOM-factory styles — **now linked by `studio.html` only** |
| `convergence.css` | 599 | Studio Convergence tab |
| `settings-modal.css` | 553 | Shared settings modal |
| `topnav.css` | 327 | Top nav + **shared toast** |
| `gallery.css` | 304 | |
| `color-picker.css` | 118 | |
| `card-scrubber.css` | 24 | Parked (unloaded) |

### Delivery paths

| Path | Mechanism |
|------|-----------|
| Live Studio / Screenspace / Transcripts | `register_static_routes()` + ordered `<script>` tags |
| Exported viewer / gallery | `viewer.py` inlines CSS/JS + `window.CLIPGEN_DATA` |
| Multi-participant export | `timeline-viewer.html` — swimlane variant of `viewer.html`, reuses `viewer.js`/`viewer.css` |

---

## Duplication map

| Pattern | Status | Notes |
|---------|--------|-------|
| Export quick action | ✅ RESOLVED | `export-actions.js` (`window.ClipgenExportActions`), used by all three pages |
| Toast markup + CSS | ✅ RESOLVED | Shared in `topnav.css`; all three pages aligned |
| Video speed cycle | ✅ RESOLVED | `video-controls.js` (`nextSpeed`/`applyPlaybackRate`) |
| Icon masks | ✅ RESOLVED | Unified `iconMask*` helpers in `utils.js`; old APIs collapsed onto them |
| Canvas theme colors | ✅ RESOLVED | `getCanvasThemeColors()` in `utils.js`; Screenspace + Transcripts |
| Intake clustering | ✅ RESOLVED | `intake-cluster.js` (`window.ClipgenIntakeCluster`); `_studioCluster*` gone |
| Polling | ✅ RESOLVED | `createPoller` adopted across Studio, Screenspace, Transcripts (8 pollers converted 2026-06-23); only the Ollama model-pull loop stays a raw `setInterval` (Promise-based, custom cancel/miss-count) |
| `renderTimeline()` | 🟡 PARTIAL | Ruler core (`niceTimeInterval`/`drawTimelineRuler` in `utils.js`) now drives the two **canvas** surfaces (screenspace + transcripts). Markers/bands/playhead stay per-surface; the two **DOM** rulers (viewer fixed-8, convergence fixed-11) are a different model and untouched. |
| Card scrubber | ❌ STILL PRESENT | `card-scrubber.js` parked (unloaded) **and** dup'd inline in `viewer.js` |
| HTML `<head>` | ✅ RESOLVED | Live pages (Studio/Screenspace/Transcripts) embed `<!-- CLIPGEN_HEAD_HERE -->`, expanded server-side from `assets/web/_head.html` by `utils.render_index_html()`. Exported viewers keep their self-contained inline `data:` favicons. |

### Studio coupling

- `window._studioState` — **still the contract** for `metadata.js` / `convergence.js` (lazy-read on activation)
- `window._studioCluster*` — ✅ gone; clustering now lives in `intake-cluster.js`

---

## Convention gaps (AGENTS.md)

| Rule | Gap |
|------|-----|
| Icons via `mask-image` from `assets/icons/` | ✅ Closed (with documented exceptions): all functional icons use `mask-image`. Remaining inline `<svg>` are intentional exceptions — loading/pulse animations (`studio.html` spinners, `studio.js` `createPulserOverlay`), brand/file-type glyphs (start-overlay Google/Excel tabs), start-overlay decorative artworks, and `viewer.css` data-URIs (exported/offline). Each carries an inline comment; see AGENTS.md "Standing exceptions". |
| Tokens for spacing/type in new/touched CSS | ✅ Swept (2026-06-23): 181 raw-`px` → tokens across all five page CSS files (`font-size`→`--text-*`, `margin`/`padding`/`gap`→`--space-*`, `border-radius`→`--radius-*`). Added 4 tokens: `--text-2xs` (11px), `--radius-xs` (2px), `--space-1-5` (6px), `--space-2-5` (10px). Intentionally left raw: `1px`/hairline borders, computed constants (`92px` chrome), widths/heights/shadows, the root `html { font-size }` (the rem anchor — must stay concrete px), and one-off magic numbers (fractional fonts, transform nudges, slider tracks). |
| No duplicate Python/JS constants | Intentional mirrors in `utils.js` + tests — maintenance surface only |
| `primitives.js` only where factories needed | ✅ Resolved: `primitives.css` now linked by `studio.html` only (dropped from Screenspace/Transcripts) |

Decorative SVG in `start-overlay.html` remains an explicit, documented exception (the tool-tile artworks + brand/file-type tab glyphs); see AGENTS.md "Standing exceptions".

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
| C2 | Studio grid perf ([archive/PERFORMANCE-PLAN.md](archive/PERFORMANCE-PLAN.md) §4.1) | Large-sheet UX — **profile first** |
| C3 | NDJSON intake streaming (archive/PERFORMANCE-PLAN §4.2) | Per-item progress |
| C4 | Inline SVG → mask-image (viewer + Screenspace toolbars first) | Convention compliance |
| C5 | Token migration on touched CSS files | Incremental, not big-bang |
| C6 | Flask-injected HTML head partial | DRY favicon/fonts |

#### Split order for C1 — follow the established hub+satellite pattern

The earlier proposal (`state/canvas/timeline/tasks/workflow`) is **OBE**. The codebase has
since settled on a **hub + feature-satellite** convention (documented in AGENTS.md): a hub
file keeps shared mutable state under a `window.Clipgen*` namespace (aliased e.g. `SS`), and
satellites are carved by **feature/UI surface**, loaded as ordered `<script>` tags after the
hub. New C1 work should match this, not invent a new axis.

Current state and next targets:

1. ~~**`transcripts.js`**~~ ✅ **Done (2026-06-24):** carved into a ~2.4k hub + 5 satellites
   (`transcripts-{corrections,search,video,pills,agents}.js`) behind `window.ClipgenTranscripts`
   (TS), one PR. Hub keeps `state`, the xref index, `selectParticipant`, the segments+marks
   editor, the task poller, and model-install; satellite functions reached via same-named guarded
   delegators. Notable: **no `transcripts-utils.js`** (no shared pure-helper cluster — `showToast`
   stayed hub-local to preserve its 2500 ms override); routed cross-file mutable state through
   `state` (`participantReqVer`, `cachedSegmentRows`, `frictionTooltipShown`) and accessors
   (`TS.isSummaryPolling`/`TS.hasTimelineHover`); fixed a latent `accent` ReferenceError in
   `renderPlayhead` while moving it.
2. **`screenspace.js`** (7.5k hub; satellites `multitool-params`/`calibration`/`color`/`utils`
   already carved via `window.ClipgenScreenspace`/`SS`) → continue extracting cohesive feature
   surfaces (tasks/queue UI, timeline, region CRUD) off the hub. Audit hub refs to satellite
   `var`s when carving (see AGENTS.md gotcha — `node --check` won't catch `ReferenceError`s).
3. **`studio.js`** (5.8k state hub) → `intake-cluster.js` + `metadata.js`/`convergence.js`
   already split out; remaining win is reducing the `window._studioState` surface, not more
   file splits.

Globals stay on `window` namespaces; script order in HTML documents dependencies.

---

## Waves — status

| Wave | Items | Status |
|------|-------|--------|
| **1** | A1, A2, A3, A4 | ✅ Shipped (PR #366); A3 polling adoption completed 2026-06-23 (Screenspace + Transcripts) |
| **2** | A5, A6, B1, B2 | ✅ Shipped (PR #385) — except A6 card-scrubber (still parked) |
| **3** | B3, C4 (incremental), C5 (opportunistic) | ✅ B3 shipped (PR #388); C4 closed (documented exceptions); C5 full sweep done (2026-06-23) |
| **4** | C1, C2, C3 | ⬜ Not started (Screenspace partially carved on a different axis) |

## Remaining work — re-prioritized (as of 2026-06-23)

1. **Finish half-done (now, low risk):** ✅ Cleared
   - ~~**A3** — adopt `createPoller` in Screenspace + Transcripts~~ ✅ Done (2026-06-23): 8 pollers converted (Screenspace ×1, Transcripts ×6, Studio job-status ×1); Ollama model-pull loop left as a raw `setInterval` (Promise-based, custom cancel/miss-count).
   - ~~**C6** — server-injected `<head>` partial (favicon + fonts) for the three live pages~~ ✅ Done (2026-06-23): shared `assets/web/_head.html` expanded into live index pages by `utils.render_index_html()`; exported viewers left self-contained.
2. **Decide & close:**
   - **A6** — card-scrubber: **delete** the parked module + the inline `viewer.js` dup, or keep parked with a one-line pointer. Parked across two ARCHITECTURE.md notes — make the call.
3. **Opportunistic (touched files only):**
   - ~~**C4** — remaining inline SVG → `mask-image`~~ ✅ Closed: audited `studio.js`/`studio.html`/`start-overlay.html`; remaining inline `<svg>` are documented intentional exceptions, `viewer.css` data-URIs kept.
   - ~~**C5** — token sweep on touched CSS~~ ✅ Done: full 181-`px`→token sweep across all five page CSS files + 4 new tokens.
   - ~~**B5** — `.btn` vs `.cg-btn`~~ ✅ Done (2026-06-24): shared `.btn` base consolidated into `tokens.css`; dual system documented as intentional in AGENTS.md; full merge onto `.cg-btn` left to per-button commits.
4. **Structural (largest remaining value, when there's appetite):**
   - **C1** — per the split-order section above: **Transcripts first**, then continue Screenspace/Studio carve-outs on the hub+satellite axis.
   - ~~**B4** — shared timeline ruler core~~ ✅ Done (2026-06-24): `niceTimeInterval`/`drawTimelineRuler` in `utils.js` shared by the two canvas surfaces; DOM rulers (viewer/convergence) left as a separate model. Full marker-level abstraction still deferred.
5. **Verify / relink before scheduling:**
   - **C2 / C3** — reconcile against `archive/PERFORMANCE-PLAN.md` + `archive/PERFORMANCE-PLAN-2.md`; confirm what perf work already shipped before treating Studio-grid / NDJSON-intake as pending.

---

## Explicitly out of scope

- React / TypeScript / bundler
- Backwards-compat for old exported HTML viewers
- Headless browser CI
- Merging Quick Actions with Studio sub-header (gallery drawer flow — see REDESIGN-PLAN session 1)
- Screenspace flat-scroll, Transcripts sticky PiP (redesign deferred)
- SSE task progress ([archive/PERFORMANCE-PLAN.md](archive/PERFORMANCE-PLAN.md) dropped for now)
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
- [x] A3 `createPoller` in `utils.js` — adopted across Studio, Screenspace, Transcripts (8 pollers, 2026-06-23); Ollama model-pull loop intentionally left as raw `setInterval`
- [x] A4 `getCanvasThemeColors` in `utils.js`; Screenspace + Transcripts canvases

### Wave 2

- [x] A5 CSS include audit (Transcripts/Screenspace) — toast moved to `topnav.css`; `primitives.css` dropped from both
- [x] A6 Card-scrubber: **integrated** (opt-in). Studio settings toggle (`STUDIO_CARD_SCRUBBER`) + `/api/sprite`/`/api/clip-audio` routes; timeline-viewer header toggle layering audio/waveform onto the existing `<video>`-seek scrub (inlined by `viewer.py`)
- [x] B1 `video-controls.js` — shared `nextSpeed`/`applyPlaybackRate`
- [x] B2 `intake-cluster.js` — `window.ClipgenIntakeCluster`; drops `window._studioCluster*`

### Wave 3

- [x] B3 Unified icon helper — `iconMaskUrl`/`iconMaskStyle`/`applyIconMask`/`iconMaskSpan`/`applyIconMasksIn` in `utils.js`; `svgMask`, `iconSpan`, `xrefBadgeIcon`, `applyDataIconMasks`, `applyIcons`, and inline calls collapse onto them
- [x] C4 SVG → mask — **Closed (with documented exceptions).** Screenspace toolbar converted (11 `.wf-tab` + info + Run reuse the `.ss-task-icon` family); audited `studio.js`/`studio.html`/`start-overlay.html` — no plain functional icons remain. Remaining inline `<svg>` are intentional exceptions (animations, brand/file-type glyphs, decorative artworks) with inline comments; `viewer.css` data-URIs stay (offline exports). See AGENTS.md "Standing exceptions".
- [x] C5 Token sweep — full pass: 181 raw `px` → tokens across all five page CSS files; 4 new tokens added to `tokens.css` (`--text-2xs`, `--radius-xs`, `--space-1-5`, `--space-2-5`)

### Remaining (re-prioritized)

- [x] A3 finish — `createPoller` adopted in Screenspace + Transcripts (and Studio job-status); only the Ollama model-pull loop stays raw `setInterval`
- [x] C6 server-injected `<head>` partial — `assets/web/_head.html` expanded into the three live index pages via `<!-- CLIPGEN_HEAD_HERE -->` + `utils.render_index_html()`; exported viewers stay self-contained
- [x] A6 card-scrubber — **integrated** (opt-in, default off) on Studio + the timeline viewer; the viewer keeps its `<video>`-seek visual scrub and adds the module's audio + waveform. Note: the inline `viewer.js` scrubber and `card-scrubber.js` were never true duplicates (video-seek vs sprite-sheet), so the viewer-dup line in the duplication map is moot
- [x] C4 SVG → mask — closed: `studio.js`/`studio.html`/`start-overlay.html` audited; remaining inline `<svg>` are documented intentional exceptions (animations, brand/file-type glyphs, decorative artworks), `viewer.css` data-URIs kept
- [x] C5 token sweep — done: full pass across all five page CSS files; 4 new tokens (`--text-2xs`, `--radius-xs`, `--space-1-5`, `--space-2-5`)
- [x] B5 `.btn` base consolidated into `tokens.css` (2026-06-24) — was copy-pasted verbatim across studio/screenspace/transcripts CSS; zero visual change. Dual `.btn`/`.cg-btn` system documented as intentional (AGENTS.md); full merge onto `.cg-btn` left to per-button commits
- [x] C1 `transcripts.js` split first (2026-06-24) — hub + 5 `transcripts-*.js` satellites behind `window.ClipgenTranscripts`; one PR. No `transcripts-utils.js` (no shared pure-helper cluster); cross-file state routed through `state` + `TS.*` accessors; fixed a latent `accent` ReferenceError in `renderPlayhead`
- [ ] C1 `screenspace.js` — continue hub carve-outs (satellites already on `window.ClipgenScreenspace`)
- [ ] C1 `studio.js` — shrink `window._studioState` surface (intake/metadata/convergence already split)
- [x] B4 shared timeline **ruler** core (2026-06-24) — `niceTimeInterval`/`drawTimelineRuler` in `utils.js`, adopted by the two canvas surfaces (screenspace + transcripts); markers/bands/playhead stay per-surface. The two DOM rulers (viewer fixed-8, convergence fixed-11) left as-is — different model. Full marker-level abstraction still out of scope
- [ ] C2 / C3 — **first** reconcile against `archive/PERFORMANCE-PLAN*.md` (confirm what shipped), then Studio grid profile + NDJSON intake if still pending

---

## Document history

| Date | Change |
|------|--------|
| 2026-05-19 | Initial plan from frontend refactor investigation (plan-only session) |
| 2026-06-23 | Refreshed inventory/counts to current reality (43 files); repointed archived cross-links; reconciled Screenspace C1 to the actual hub+satellite axis; marked Waves 1–3 shipped + A3 partial; re-prioritized remaining work (Transcripts split first) |
| 2026-06-23 | Closed C4 (icon gap — documented intentional inline-SVG exceptions), C5 (full 181-`px`→token sweep + 4 new tokens), and A3 (`createPoller` across all 8 remaining pollers). Verified A1/A2/A4/A5 already shipped. |
| 2026-06-23 | Closed C6: shared `assets/web/_head.html` favicon/fonts partial injected into the three live index pages via `<!-- CLIPGEN_HEAD_HERE -->` + `utils.render_index_html()`; exported viewers left self-contained. "Finish half-done" bucket now empty — next up is C1 (Transcripts split). |
| 2026-06-24 | B4 (ruler core): extracted `niceTimeInterval`/`drawTimelineRuler` into `utils.js`, adopted by the two canvas surfaces (screenspace + transcripts); DOM rulers (viewer/convergence) left as a separate model; markers/bands/playhead stay per-surface. B5 (`.btn` de-dup): consolidated the verbatim-triplicated `.btn` base into `tokens.css` (zero visual change; also gave Studio the missing `.btn:disabled:hover` guard), documented the intentional `.btn`/`.cg-btn` duality in AGENTS.md, and deferred the full merge onto `.cg-btn` to per-button commits. |
| 2026-06-24 | C1 (Transcripts split): carved `transcripts.js` (5.2k monolith) into a ~2.4k hub + 5 satellites (`transcripts-{corrections,search,video,pills,agents}.js`) behind `window.ClipgenTranscripts` (TS), one PR. Hub keeps `state`, xref index, `selectParticipant`, segments+marks editor, the task poller, and model-install; satellite fns reached via same-named guarded delegators. No `transcripts-utils.js` (no shared pure-helper cluster; `showToast` stayed hub-local). Routed cross-file mutable state through `state` (`participantReqVer`/`cachedSegmentRows`/`frictionTooltipShown`) + accessors (`TS.isSummaryPolling`/`TS.hasTimelineHover`); fixed a latent `accent` ReferenceError in `renderPlayhead`. `node --check` + a vm load/boot smoke harness + dom-wiring test (now globs `transcripts*.js`); browser pass pending. |
