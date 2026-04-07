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
