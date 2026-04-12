# Convergence Browser

## Studio Tab 4

---

## Why We're Building This

Clipgen's Studio currently operates per-stream. Researchers work through spreadsheet data, ScreenSpace events, and transcript tags in separate intake tabs, manually cross-referencing by timestamp. This works for single-session analysis, but breaks down when working across multiple participants.

The gap: **cross-participant signal is invisible unless actively sought.** A researcher focused on participant 5 may not notice that participants 1, 3, and 4 all flagged the same moment. The data exists in memory — Studio already holds all session JSONs simultaneously for cross-referencing — but there is no surface that makes recurrence visible passively.

The Convergence Browser closes this gap. It answers two questions researchers ask constantly but currently answer by hand:

- **Did this happen for multiple participants?** (convergence)
- **When did it happen, and how much variance is there?** (distribution)

A tight temporal distribution (everyone hit this at minute 4) is a different finding than a wide one (spread across the session). Both are signal. Neither is currently legible without manual comparison.

---

## What It Is

An interactive, multi-participant timeline view — a 4th tab in Studio. It borrows the visual structure of the Timeline Viewer (already a generated artifact) but makes it interactive and adds filter and selection controls. The output of working in the Convergence Browser feeds directly into the existing artifact and reel fields.

It is not a replacement for the three intake tabs. Researchers who want to work stream-by-stream should continue to do so. The Convergence Browser is a higher-altitude entry point: use it when you want the data to surface candidates rather than hunting manually.

---

## How It Works

### Data Source

All session JSONs are already in memory when Studio is active. No new loading infrastructure is required. The Convergence Browser reads from the same in-memory data the cross-reference system already uses: `state.sheetData`, `state.intakeClusters`/`state.intakeEvents`, and `state.trIntakeClusters`/`state.trIntakeMarks`.

#### Baseline Time Normalization (prerequisite)

ScreenSpace events and transcript marks are already video-relative (seconds from video start). Spreadsheet timestamps may be absolute clock times (e.g. `09:12:00`) when a baseline row exists in the study spreadsheet. Currently, baseline-to-relative conversion only happens server-side in `files.prepare_clip()`. The client's `parseClipTimestamps()` does not apply baseline correction — it parses raw `MM:SS` / `HH:MM:SS` strings to seconds without offset.

The Convergence Browser must baseline-adjust sheet timestamps client-side before plotting them alongside ScreenSpace and transcript events. The baseline offset per participant is derivable from the sheet data (the baseline row's value for each participant column). Without this, sheet timestamps and detector timestamps will be misaligned for studies that use clock-style entry.

### Layout

- **Horizontal axis**: session-relative time (normalized across participants by study design — all participants start at minute 0 and proceed through the same tasks)
- **Rows**: one row per participant, each with sub-tracks per stream (spreadsheet, ScreenSpace, transcript)
- **Convergence overlay**: a dual-layer density visualization:
  - **Summary lane** at the top of the participant rows — a single horizontal band spanning the full timeline width. Color intensity encodes convergence strength (participant count × temporal tightness). The researcher scans this lane for hot spots, then looks down at participant rows to see who contributed. Rendered as a canvas layer for smooth continuous gradients.
  - **Per-participant row shading** — subtle background gradient on each participant's row showing that participant's contribution to active convergence zones. Secondary to the summary lane; helps answer "which participants drove this convergence?" at a glance.
  - Clustered flags (tight distribution) are visually distinguishable from spread ones (wide distribution) through the summary lane's intensity profile.

### Filters

Populated dynamically from whatever event type strings exist in the loaded JSONs. No prepackaged taxonomy — the researcher named their detectors when configuring ScreenSpace, and those names are the vocabulary for this study.

Filters operate as **prerequisites for convergence calculation**, not post-filters. The researcher selects an event type first; convergence is then calculated for that subset. This is necessary because ScreenSpace event density varies wildly by detector type (a health bar detector may fire hundreds of times per minute; a loading screen detector may fire three times per session). Mixing them in a single density calculation produces noise.

Controls:

- Filter by stream (spreadsheet / ScreenSpace / transcript) — or **"all streams"** toggle to treat events from all streams as candidates for convergence, enabling cross-stream questions ("where do spreadsheet pain-point flags AND ScreenSpace change events co-occur across participants?")
- Filter by event type (populated from loaded JSONs) — when a single stream is selected, shows that stream's types; when "all streams" is active, shows types from all streams grouped by source
- Convergence threshold (minimum number of participants)
- Convergence window (±W seconds — how close in time events must be to count as convergent; default from existing cluster threshold)
- Time range

### Convergence Algorithm

Convergence zones are identified using a sweep-line approach — not fixed time bins, which break down given ScreenSpace density variance.

1. **Filter**: collect the union of all events matching the active stream/type filters, across all participants.
2. **Sort** the merged event list by start time.
3. **Sweep**: for each event, count how many distinct participants have at least one event within ±W seconds (where W is the user-configurable convergence window).
4. **Threshold**: regions where the distinct-participant count ≥ the convergence threshold become convergence zones.
5. **Merge**: adjacent or overlapping convergence zones are merged into contiguous regions.

This is O(n log n) and efficient for client-side computation. The configurable window reuses the threshold-slider pattern already present in the intake tabs (the same interaction model as `clusterIntakeEvents` in the existing codebase, extended to work across participants rather than within a single participant's stream).

The key difference from per-participant clustering: the existing intake clustering groups events within one participant by temporal proximity. The convergence algorithm groups events across participants — asking not "which of P01's events belong together?" but "at this moment, how many different participants had something happen?"

### Display Normalisation

Per-track density is normalised for display so sparse tracks don't visually disappear alongside dense ones. Absolute event counts are preserved in the underlying data; this is a display concern only.

Method: per-participant min-max scaling within the filtered event type. Because filters operate as prerequisites (a single event type or "all streams" is selected before convergence runs), normalisation only needs to handle density variation across participants within that subset — not across wildly different detector types simultaneously.

### Selection

Two selection modes:

- **Click** a convergence zone in the summary lane to select that zone's time range. The zone is already algorithm-identified — clicking it is a single action.
- **Drag-to-select** on the timeline: mousedown → mousemove → mouseup creates a custom time range selection. This lets the researcher define an arbitrary region, not just one the algorithm identified — useful for exploring near-misses or partial convergence.

Selection highlights the time range across all participant rows and opens a detail panel.

The Convergence Browser does not generate its own output format — it is a curation surface that feeds the existing generation pipeline.

### Detail Panel

When a convergence zone or custom range is selected, a detail panel appears showing per-participant event breakdowns within that range. This is the drill-down path from "4 participants converged here" to "here is exactly what each participant's data looks like at this moment."

The detail panel reuses the cross-reference join pattern already in Studio (`findOverlappingData` — given a participant and time range, returns matching events from all three streams). Each participant within the selection gets a section showing their events, with cross-reference badges indicating cross-stream overlap.

From the detail panel, the researcher can:

- Add individual events to the artifact or reel queue (one per participant, consistent with the existing per-participant dispatch in `api/generate` and `api/generate-intake` — no new queue item shape or backend changes needed)
- Add all events in the selection as a batch (N individual items, one per participant)
- Dismiss the selection to continue exploring

### Video Preview

Hovering over event markers in participant rows shows a video frame preview, consistent with all other interactive surfaces in Studio. This reuses the existing video frame endpoint (`../screenspace/api/video/frame/{participant}/{time}`) with a 60ms hover debounce.

---

## Key Design Decisions

**Query interface, not dashboard.** The browser's value is in answering specific cross-participant questions about specific event types — not displaying everything at once. The holistic view is useful for orientation, but the real work happens when the researcher narrows to a type and asks "did this happen for everyone, and when?"

**No interpretation imposed.** The browser surfaces convergence and distribution as evidence. What that evidence means is the researcher's call. This is consistent with clipgen's overall architecture.

**Study-relative taxonomy.** Event type labels are read from the loaded JSONs. The browser cannot and should not make assumptions about what those labels mean across studies. Clipgen must run on thousands of different games in various stages of development — no prepackaged definitions are possible.

**Additive, not disruptive.** The three existing intake tabs are unchanged. The Convergence Browser is an additional entry point for a different mode of working.

**Complements, not replaces, cross-reference badges.** The existing cross-reference badge system on intake cards answers a per-event, per-participant question: "what else happened at this moment for this participant across streams?" The Convergence Browser answers a different question: "across participants, where did this type of event cluster in time?" Both are useful; neither subsumes the other.

---

## Integration Points

- Reads from: all session JSONs currently in Studio memory (`state.sheetData`, `state.intakeEvents`/`state.intakeClusters`, `state.trIntakeMarks`/`state.trIntakeClusters`)
- Writes to: existing artifact queue and reel queue — N individual items (one per participant), no new queue shape or backend endpoints required
- Relationship to Timeline Viewer: borrows its visual structure (percentage-based positioning, participant rows, track expansion); the Timeline Viewer remains a separate static generation artifact for stakeholder output
- Relationship to Metadata Overview (Tab 5): the cross-stream collision data visible in the Metadata Overview provides useful context before entering the Convergence Browser
- Relationship to cross-reference badges: the Convergence Browser's detail panel reuses the `findOverlappingData()` join and `buildXrefBadges()` rendering for per-event cross-stream context

### Tab Visibility

The Convergence tab is hidden by default, like the existing intake tabs. It appears only when multiple participants' data is loaded — convergence across a single participant is meaningless. The visibility condition is: `state.sheetData.participants.length > 1` OR multiple distinct participants exist across intake/transcript data. This follows the existing `checkNavLinks()` gating pattern.

### Data Freshness

When the Convergence tab is active and new ScreenSpace or transcript events arrive via the existing polling cycle, the browser does **not** auto-recalculate. Instead, it shows a "new data available" refresh indicator. The researcher clicks to recalculate when ready. Auto-recalculation during active exploration or with an active selection would be disorienting — the view would shift under the researcher's cursor.

---

## Rendering & Performance

**Hybrid rendering.** The density overlay (summary lane, per-row shading) is rendered on a canvas layer for smooth continuous gradients. Individual event markers on participant rows are DOM elements with percentage-based CSS positioning, enabling native hover, click, and drag interactions without manual hit-testing. This matches the Timeline Viewer's approach (DOM markers, canvas optional for screenspace overlay).

**Debounce filter changes.** Filter and threshold adjustments trigger convergence recalculation after a 200–300ms debounce, not on every keystroke or slider tick. The convergence result is cached and only recomputed when filters actually change.

**Participant row ordering.** Default order matches the spreadsheet column order (consistent with how participants appear in other tabs). An optional sort-by-convergence-density reordering surfaces the most convergent participants at the top.

**Scale.** At a fixed row height (~60px with sub-tracks), more than ~8 participants requires vertical scrolling. The time axis and participant labels should be sticky (CSS `position: sticky`) so orientation is maintained during scroll. The convergence summary lane at the top should also be sticky.

---

## Key Findings From Design Discussion

- Studio already holds all session JSONs in memory simultaneously — no new aggregation infrastructure is required as a prerequisite
- Session-relative timestamps are effectively normalised by study design (all participants start at minute 0)
- Fixed time windows for convergence calculation break down given ScreenSpace density variance; event-type-aware calculation is required
- The most common researcher query is temporal distribution ("when did this happen across participants, and is that spread tight or wide?") — threshold filtering alone doesn't capture this
- Task region as a concept is too rigid for freeform playtests; temporal distribution is the more honest and useful framing
- The Convergence Browser makes the intermediate curation step explicit: curate → query/filter → generate, rather than curate → generate

---

## Implementation Plan

### Architecture

The Convergence Browser is implemented as a **separate JS/CSS file pair** (`convergence.js` + `convergence.css`), integrated into Studio via minimal HTML additions and a bridge API exposed from `studio.js`. One new server endpoint provides baseline timestamp data. Studio.js is already 4,500 lines — keeping convergence logic in its own file prevents further bloat while maintaining the same vanilla JS patterns.

**Files to create:**
| File | Purpose | Est. size |
|------|---------|-----------|
| `assets/web/convergence.js` | All Convergence Browser logic (IIFE, ES5) | ~1,000 lines |
| `assets/web/convergence.css` | All Convergence Browser styles | ~350 lines |

**Files to modify:**
| File | Changes |
|------|---------|
| `assets/web/studio.html` | Tab button, panel container, script/css includes (~15 lines) |
| `assets/web/studio.js` | State fields, tab switching, bridge exports (~30 lines) |
| `server.py` | `/api/sheet/baseline` endpoint (~25 lines) |
| `tests/test_studio_api.py` | Baseline endpoint tests (~40 lines) |

**Bridge pattern:** `studio.js` exposes select internal functions on `window` for `convergence.js` to consume:
- `window._studioState` — reference to the shared state object (same in-memory data, no copies)
- `window._studioFindOverlappingData` — cross-reference join (studio.js:227)
- `window._studioBuildXrefBadges` — badge rendering (studio.js:3361)
- `window._studioParseClipTimestamps` — timestamp parsing (studio.js:166)
- `window._studioHexToRgba` — color utility
- `window._studioIntakeComputeTickInterval` — time axis tick calculation
- `window._studioRenderArtifactQueue` / `_studioRenderReelQueue` / `_studioSaveQueues` — queue re-rendering after dispatch

`convergence.js` also uses globals from `utils.js` directly: `qs`, `qsa`, `el`, `formatTime`, `MARK_CATEGORIES`, `XREF_BADGES`, `DETECTOR_COLORS`.

---

### Phase 0: Baseline Time Normalization ✓

**Goal:** Expose per-participant baseline timestamps to the client so sheet timestamps can be converted from absolute clock time to video-relative seconds.

**What becomes functional:** Nothing user-visible. Infrastructure for Phase 2.

#### 0.1 — New server endpoint

**File:** `server.py` (add near existing `/api/sheet` at line 162)

```
GET /api/sheet/baseline → {"ok": true, "baselines": {"P01": "09:12:00", "P02": "", ...}}
```

Reads `_sheet_context.baseline_row_idx` and iterates participant columns — the same access pattern already in `spreadsheet.py:582-587` (`prepare_clip`). Empty string = relative timestamps (no adjustment needed). Non-empty = clock-format baseline.

Returns `{"ok": true, "baselines": {}}` when no baseline row exists.

**Why a separate endpoint:** Avoids changing the `/api/sheet` response shape that the sheet grid depends on. Baseline data is only needed by the Convergence Browser.

#### 0.2 — Client-side baseline parsing

**File:** `convergence.js`

- `parseTimestampToSeconds(str)` — same logic as studio.js:154, but accessible outside the IIFE. Move to `utils.js` as a shared global (or duplicate in convergence.js — small function, either is acceptable).
- `getBaselineAdjustedSheetEvents(sheetData, baselines)` — iterates `state.sheetData.rows`, parses each valid cell, subtracts `baselines[participant]` from `startSeconds` when the participant has a non-empty baseline.

#### 0.3 — Tests

**File:** `tests/test_studio_api.py`

- `test_api_sheet_baseline_no_context` — returns 500
- `test_api_sheet_baseline_no_baseline_row` — returns `{"ok": true, "baselines": {}}`
- `test_api_sheet_baseline_with_values` — returns per-participant baseline strings
- `test_api_sheet_baseline_partial` — some participants have baselines, some don't

---

### Phase 1: Tab Shell and Panel Skeleton

**Goal:** The Convergence tab appears when multiple participants are loaded. Clicking it shows an empty panel with filter controls.

**What becomes functional:** Tab navigation works. Filter UI renders. Panel is structurally complete but shows no data.

*Can start in parallel with Phase 0.*

#### 1.1 — HTML additions

**File:** `studio.html`

- Add `<link rel="stylesheet" href="convergence.css">` after existing CSS includes (line 8)
- Add tab button after transcript-intake tab (line 83): `<button class="preview-tab hidden" data-tab="convergence">Convergence</button>`
- Add panel div after `#trIntakePanel` (line 152):
  ```html
  <div id="convergencePanel" class="hidden">
    <div id="convergenceControls" class="area-controls"></div>
    <div id="convergenceFilters" class="convergence-filters"></div>
    <div id="convergenceTimeline" class="convergence-timeline-container"></div>
    <div id="convergenceDetail" class="convergence-detail hidden"></div>
  </div>
  ```
- Add `<script src="convergence.js"></script>` after `studio.js` (line 317)

#### 1.2 — State additions in studio.js

**File:** `studio.js` (lines 8-48, state object)

```javascript
convergenceBaselines: {},
convergenceDataVersion: 0,
convergenceStale: false,
```

#### 1.3 — Tab visibility gating

**File:** `studio.js`

New function `checkConvergenceTabVisibility()` — called after `loadSheetData()` resolves (line 683) and after `pollIntakeEvents`/`pollTranscriptIntakeMarks` update data. Shows the tab when:

```javascript
(state.sheetData && state.sheetData.participants.length > 1)
|| (uniqueParticipants(state.intakeEvents) > 1)
|| (uniqueParticipants(state.trIntakeMarks) > 1)
```

This differs from the existing `checkNavLinks()` pattern (which queries `/api/status` for server capability flags) because the convergence gate is about loaded data, not server features.

#### 1.4 — Tab switching integration

**File:** `studio.js`, `syncPreviewTab()` (line 605)

- Add `#convergencePanel` to the "hide everything" block (line 614-619)
- Add `else if (state.activePreviewTab === "convergence")` branch that shows the panel and calls `window.convergenceActivate()` — the entry point from `convergence.js`

#### 1.5 — Convergence.js skeleton

**File:** `convergence.js` (new, IIFE)

Internal state:
```javascript
var cvState = {
  active: false,
  baselines: null,           // {P01: 33120, ...} in seconds, null = not fetched
  events: [],                // unified events from all sources
  filteredEvents: [],
  convergenceZones: [],
  selection: null,           // {start, end} or null
  filters: {
    streams: [],             // subset of ["sheet", "screenspace", "transcript"]
    eventTypes: [],
    minParticipants: 2,
    windowSec: 5,
    timeRange: null,
  },
  dataVersion: 0,
  duration: 0,
  participants: [],
};
```

Exposes `window.convergenceActivate`, `window.convergenceDeactivate`, `window.convergenceInit`.

#### 1.6 — Filter controls

**File:** `convergence.js`

`buildFilterControls()` renders into `#convergenceControls` and `#convergenceFilters`:

- **Stream toggles** (sheet / screenspace / transcript / all) — styled like `.intake-filter-det` pills, colored with `XREF_BADGES` values
- **Event type pills** — dynamically populated from loaded data. When a single stream is selected, show that stream's types. When "all" is active, show types grouped by source
- **Min participants** — `<input type="number" min="2">`, default 2, `.intake-cluster-input` pattern (studio.html:110)
- **Window** — `<input type="number" min="1" max="60" value="5">`, label "±s"
- All filter changes debounced 250ms before triggering recalculation

#### 1.7 — CSS skeleton

**File:** `convergence.css` (new)

Initial styles using `tokens.css` variables throughout:
- `#convergencePanel` — flex column layout matching `#intakePanel` pattern
- `.convergence-filters` — horizontal flex wrap
- `.convergence-timeline-container` — flex: 1, min-height: 0, overflow-y: auto
- Filter pill styles reusing `.intake-filter-det` visual pattern

#### Verification

- [x] Load multi-participant study → Convergence tab appears
- [x] Load single-participant study → tab hidden
- [x] Click Convergence tab → empty panel with filter controls
- [x] Switch between all four tabs → correct panel visibility, no poll timer leaks

---

### Phase 2: Data Collection, Algorithm, and Timeline Rendering

**Goal:** The Convergence tab loads data from all three sources, runs the sweep-line convergence algorithm, and renders the multi-participant timeline with convergence summary lane.

**What becomes functional:** Researcher can see all participants on a shared timeline, see convergence zones highlighted, and use filters to explore event types.

**Depends on:** Phase 0 + Phase 1.

#### 2.1 — Data collection

**File:** `convergence.js`

`collectAllEvents()` gathers events from all three sources into a unified shape:

```javascript
{ participant, start, end, source, eventType, label, id, rawData }
```

- **Sheet:** iterate `state.sheetData.rows`, parse each valid cell via `parseClipTimestamps()`, apply baseline offset from `cvState.baselines[participant]`. Source = "sheet", eventType = row.category
- **Screenspace:** iterate `state.intakeEvents` (raw events, not clusters — finer granularity). Source = "screenspace", eventType = `event.event_type`, start = `time_in`, end = `time_out`
- **Transcript:** iterate `state.trIntakeMarks`. Source = "transcript", eventType = `mark.category`, start/end from mark

`cvState.duration` = `max(all event end times) * 1.05` (matching studio.js:3451 pattern).

`cvState.participants` = union of all participants, sorted by spreadsheet column order (from `state.sheetData.participants`), with non-sheet participants appended alphabetically.

#### 2.2 — Baseline fetch

`convergenceActivate()` fetches baseline data on first activation:

```javascript
if (!cvState.baselines) {
  fetch("api/sheet/baseline").then(...).then(function (data) {
    // parse each baseline string to seconds via parseTimestampToSeconds
    cvState.baselines = parsed;
    recalculate();
  });
} else {
  checkStaleness();
}
```

#### 2.3 — Convergence algorithm

`computeConvergenceZones(events, windowSec, minParticipants)`:

1. **Sort** events by start time
2. **Sweep:** for each event, count distinct participants with events within ±windowSec
3. **Threshold:** regions where distinct-participant count ≥ minParticipants become zones
4. **Merge:** overlapping zones into contiguous regions
5. **Enrich:** per zone — participant count, contributing events, tightness (stddev of start times), strength score (`participantCount/total × 1/(1 + tightness/window)`)

O(n × k) effective complexity where k = avg events per window. Sub-millisecond for typical study sizes.

#### 2.4 — Filter pipeline

`applyFilters()`:
1. Filter `cvState.events` by active streams → event types → time range
2. Store in `cvState.filteredEvents`
3. Run `computeConvergenceZones()` on filtered set
4. Store in `cvState.convergenceZones`

`recalculate()` = `collectAllEvents()` → `applyFilters()` → `render()`. Filter changes (debounced) skip collection and only re-run `applyFilters()` + `render()`.

#### 2.5 — Timeline layout

`renderTimeline()` builds inside `#convergenceTimeline`:

```
[sticky] time axis          — tick marks via intakeComputeTickInterval pattern
[sticky] summary lane       — canvas, 32px, convergence density gradient
[scroll] participant rows
  per participant:
    [sticky-left] label     — participant ID, 52px wide, monospace
    tracks container        — relative positioned
      sheet sub-track       — 14px, DOM markers
      screenspace sub-track — 14px, DOM markers
      transcript sub-track  — 14px, DOM markers
      row shading canvas    — behind markers, convergence contribution
```

**Event markers:** DOM elements with percentage-based positioning (matching viewer.js:1041-1049):
```javascript
marker.style.left = ((event.start / cvState.duration) * 100) + "%";
marker.style.width = Math.max(((event.end - event.start) / cvState.duration * 100), 0.3) + "%";
```

Color-coded by source: sheet = `XREF_BADGES.sheet.color`, screenspace = `DETECTOR_COLORS[eventType]`, transcript = `MARK_CATEGORIES[eventType].color`.

**Display normalization:** per participant, per sub-track — normalize marker opacity between 0.3 and 1.0 so sparse tracks don't vanish next to dense ones.

#### 2.6 — Summary lane (canvas)

`renderSummaryLane()` following studio.js:3404-3496 pattern:

- DPR-aware canvas sizing
- For each pixel column, sample convergence strength at that time position
- Draw with `--color-accent` at variable alpha proportional to strength
- Store hit rects in `_summaryHitRects` for Phase 3 click interaction

#### 2.7 — Per-participant row shading (canvas)

Small canvas behind DOM markers per participant row. For each convergence zone, shade proportionally to how many events that participant contributed. Subtle (max alpha 0.15) — secondary to the summary lane.

#### 2.8 — Data freshness indicator

`checkStaleness()` compares current data lengths/IDs against what was collected at `cvState.dataVersion`. If changed, show a "New data available — Refresh" banner. No auto-recalculation. Check triggered on `visibilitychange` when convergence tab is active.

#### Verification

- [ ] All participants appear as rows with correct sub-tracks
- [ ] Sheet events are baseline-adjusted (verify against raw sheet values)
- [ ] Single-stream filter shows only that stream's markers
- [ ] Event type filter recalculates convergence for that type
- [ ] Window slider widens/narrows convergence zones
- [ ] Min participants threshold shows/hides zones
- [ ] Summary lane gradient reflects convergence density
- [ ] "New data available" banner appears after external changes, refresh incorporates them

---

### Phase 3: Selection, Detail Panel, and Queue Dispatch

**Goal:** Researcher can select convergence zones or arbitrary time ranges, see per-participant event breakdowns, and send items to the artifact/reel queues.

**What becomes functional:** Complete curation workflow — discovery → inspection → queue dispatch.

**Depends on:** Phase 2.

#### 3.1 — Click-to-select convergence zone

Click handler on summary lane canvas using `_summaryHitRects`. Sets `cvState.selection`. Draws semi-transparent overlay across all participant rows (`<div class="convergence-selection-overlay">` with percentage positioning, pointer-events: none).

#### 3.2 — Drag-to-select arbitrary range

Mousedown/mousemove/mouseup on participant rows container. Draws selection preview during drag. On mouseup, if range > 1s, set `cvState.selection = {start, end, zone: null}` and render detail panel. Minimum drag threshold prevents accidental selections.

#### 3.3 — Detail panel

`renderDetailPanel()` populates `#convergenceDetail`:

```
Header: "2:15 – 2:45 · 4 participants" [Close ×]
Per participant section:
  P01 heading
    event list: time range | source badge | type | description
    cross-reference badges via findOverlappingData + buildXrefBadges
    [Add to Artifacts] [Add to Reel] per event
Actions bar:
  [Add All to Artifacts] [Add All to Reel]
```

Uses `findOverlappingData(participant, start, end)` (studio.js:227) and `buildXrefBadges()` (studio.js:3361) via the bridge API.

#### 3.4 — Queue dispatch

Push items to `state.artifactQueue` / `state.reelQueue` using existing item shapes:

- **Screenspace events:** `{participant, segStart, segDuration, desc, source: "screenspace", event_type, event_ids: [id]}`
- **Transcript marks:** `{participant, segStart, segDuration, desc, source: "transcript", mark_ids: [id]}`
- **Sheet events:** `{participant, segStart, segDuration, desc, source: "screenspace", row}` — dispatched through the intake generation path (`/api/generate-intake`)

After pushing, call `renderArtifactQueue()` / `renderReelQueue()` / `saveQueues()` via bridge.

"Add All" iterates all events in the selection across all participants, pushing each as an individual queue item (one per participant — consistent with existing per-participant dispatch).

#### 3.5 — Selection dismissal

- Close button on detail panel
- Click empty space in timeline
- Escape key

#### Verification

- [ ] Click convergence zone → detail panel opens with correct events
- [ ] Cross-reference badges appear on events with overlapping data
- [ ] "Add to Artifacts" adds individual event to queue
- [ ] "Add All to Reel" adds all events from selection
- [ ] Drag-to-select works for arbitrary ranges
- [ ] Dismiss clears overlays and hides detail panel
- [ ] Generated artifacts from convergence queue items succeed

---

### Phase 4: Video Preview and Interactive Polish

**Goal:** Hover interactions, video frame previews, keyboard shortcuts, visual refinements.

**What becomes functional:** Full interactive experience matching the polish level of existing Studio tabs.

**Depends on:** Phase 3.

#### 4.1 — Video frame preview

60ms debounced hover on event markers (matching viewer.js pattern). Floating `<img>` near marker:
- Screenspace: `../screenspace/api/video/frame/{participant}/{timestamp}`
- Transcript/Sheet: `api/thumbnail/{participant}/{start_seconds}`

#### 4.2 — Marker hover highlighting

On hover, dim all other markers (opacity 0.15 — matching studio.js:3484-3485). Highlight corresponding summary lane region.

#### 4.3 — Tooltips on convergence zones

Hover summary lane → tooltip: "4 participants · 2:15–2:45 · 12 events"

#### 4.4 — Keyboard navigation

- **Escape:** dismiss selection
- **←/→:** navigate between convergence zones when one is selected

#### 4.5 — Participant row reordering

Sort toggle: "Sort by convergence density" reorders rows so most-convergent participants appear at top. Default = spreadsheet column order.

#### 4.6 — CSS polish

- Hover transitions (use `--duration-fast` token)
- Cursor: crosshair on timeline, pointer on markers
- Frame preview: `--shadow-md`, `--radius-sm`, 240px max-width
- Dark theme: all colors via CSS custom properties from `tokens.css`
- `@media (prefers-reduced-motion: reduce)` — disable transitions

#### Verification

- [ ] Frame preview appears on hover with correct image
- [ ] Convergence zone tooltip shows on hover
- [ ] Escape dismisses, arrows navigate
- [ ] Participant sort reorders rows
- [ ] Dark mode works
- [ ] Reduced motion preference respected

---

### Phase 5: Edge Cases and Robustness

**Goal:** Handle all edge cases, performance at scale, integration testing.

**Depends on:** Phase 4.

#### 5.1 — Sticky scroll

With >8 participants: time axis and summary lane sticky at top, participant labels sticky at left. Verify no z-index conflicts.

#### 5.2 — Dense data performance

If >500 markers per sub-track, cluster adjacent markers within 1px of each other (display optimization only — algorithm data unchanged). Canvas rendering already handles density efficiently.

#### 5.3 — Empty states

- No matching events: "No events match the current filters"
- No convergence zones: "No convergence detected. Try widening the window or lowering the threshold."
- Single participant (defensive): tab hidden by Phase 1 gating

#### 5.4 — Baseline edge cases

- No baseline row: `baselines = {}`, no adjustment
- Partial baselines: apply offset only to participants with non-empty values
- Malformed baseline: treat as 0, console warning

#### 5.5 — Resize handling

Debounced 200ms re-render. Re-size all canvases. Percentage-based DOM markers auto-adjust.

#### 5.6 — Tests

**File:** `tests/test_studio_api.py` — baseline endpoint tests (Phase 0).

**Manual integration tests:**
- [ ] Study with clock-format timestamps: sheet events align with screenspace events in the Convergence Browser
- [ ] Study with no baseline row: sheet timestamps display correctly as-is
- [ ] Generate artifacts from detail panel → appear in queue → generate successfully
- [ ] Build reel from convergence-selected events across participants → generates correctly

---

### Dependency Graph

```
Phase 0 (Baseline endpoint)──────┐
                                  ├──→ Phase 2 (Algorithm + Rendering)
Phase 1 (Tab shell + filters)────┘         │
                                      Phase 3 (Selection + Queue)
                                            │
                                      Phase 4 (Preview + Polish)
                                            │
                                      Phase 5 (Edge cases + Tests)
```

Phases 0 and 1 can be developed in parallel.
