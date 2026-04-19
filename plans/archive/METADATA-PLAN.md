# Metadata Overview & Export

## Studio Tab 5

---

## Why We're Building This

Clipgen produces structured artifacts, but currently has no surface for summarising the data that underlies them. Researchers can visually eyeball event distributions in the intake tabs, but there is no way to get quick aggregate statistics — how often did an event type occur, when did it first appear, how many participants were affected — without manually reviewing individual artifacts or session JSONs.

This creates two practical problems:

**1. No fast sanity check.** A misconfigured ScreenSpace detector — wrong threshold, wrong reference frame — shows up as anomalous event counts. Without a summary view, this isn't visible until the researcher is deep into artifact curation and wonders why one participant has 10x the events of everyone else.

**2. No metadata export.** Researchers working in different reporting contexts — stakeholder decks, research repositories, internal documentation — frequently need summary statistics alongside clip artifacts. Currently these must be compiled manually.

The Metadata Overview solves both. It is a computation and display problem, not an infrastructure problem — all the data required is already in memory when Studio is active.

---

## What It Is

A 5th tab in Studio that presents aggregate statistics across all loaded sessions and streams, with export to JSON and CSV. It is a read-only summary surface — no curation or selection happens here. Its output is statistics, not artifacts.

It also serves as a natural starting point before entering the Convergence Browser: a researcher can orient themselves to the overall shape of the data before drilling into cross-participant queries.

---

## Data Sources

Three data streams are already in Studio memory when the tab is active. No new loading infrastructure is required.

**Spreadsheet data** (`state.sheetData`)
- `participants[]` — participant IDs (P01, P02, G01, etc.)
- `rows[]` — each row has: `observation` (description text), `category`, `severity`, and `cells` keyed by participant ID. Each cell has `value` (timestamp string), `valid` (boolean), `hasText` (boolean)
- Timestamps parsed via `parseClipTimestamps(cell.value)` which returns `[{ startSeconds, duration }]`

**Screenspace events** (`state.intakeEvents`)
- Each event: `id`, `participant`, `event_type` (researcher-defined label), `detector` (tool that found it: color, change, similarity, text, numbers, template, flow, scene, inactivity, multitool), `time_in`, `time_out` (seconds), `confidence` (0.0–1.0), `excluded` (boolean)
- Clusters available in `state.intakeClusters` — grouped by participant + event_type + temporal proximity

**Transcript marks** (`state.trIntakeMarks`)
- Each mark: `id`, `participant`, `text`, `category` (pain_point, delight, quote, insight, task, bookmark), `time_in`, `time_out` (seconds), `label`
- Clusters available in `state.trIntakeClusters` — grouped by participant + category + temporal proximity

---

## What It Shows

### 1. Data Coverage Matrix

A participant × stream grid that provides the fastest possible QA scan. This goes first because it immediately answers: "which participants am I missing data for in which streams?"

| | Sheet | Screenspace | Transcript |
|---|---|---|---|
| P01 | 12 timestamps | 47 events | 8 marks |
| P02 | 12 timestamps | 0 events | 15 marks |
| P03 | 10 timestamps | 23 events | 0 marks |

- **Sheet column**: count of valid timestamp cells per participant (cells where `cell.valid === true`, summing `parseClipTimestamps(cell.value).length` across all rows)
- **Screenspace column**: count of non-excluded events per participant
- **Transcript column**: count of marks per participant
- Zero counts highlighted to draw attention to gaps
- **Visualization**: heatmap coloring — each cell's background opacity is proportional to its count (normalized per-column so streams with different scales are independently readable). Follows the existing `rgba(hm, t * 0.45)` heatmap pattern from studio.js's spreadsheet grid. Zero counts get a distinct muted/warning treatment rather than just zero opacity

### 2. Per Event Type — Screenspace

Aggregate statistics grouped by `event_type` (the researcher-defined label, e.g. "health_bar_change", "loading_screen"). The `detector` (tool that produced the event) is shown as secondary metadata since multiple event types can share a detector.

Per event type:
- Total occurrence count across all participants
- First occurrence (earliest `time_in`)
- Last occurrence (latest `time_out`)
- Mean time of occurrence (mean of `time_in` values)
- Participant coverage: appeared in X of N loaded participants
- Mean confidence (mean of `confidence` values — useful for QA; low-confidence event types may indicate detector misconfiguration or a threshold that needs tuning)
- Mean duration (mean of `time_out - time_in`)
- Per-participant count breakdown (enables spotting within-type outliers: "P03 has 147 events of this type while others average 20")

Computed from raw events (`state.intakeEvents` where `excluded !== true`), not clusters. Raw events are the ground truth for counts and means.

Table is sortable by any column. Each row includes an **inline horizontal bar** (CSS `width` percentage) showing relative count, so the researcher can scan the magnitude column visually without reading numbers.

### 3. Per Category — Transcript

Aggregate statistics grouped by mark `category`. The six categories are: pain_point, delight, quote, insight, task, bookmark.

Per category:
- Total mark count across all participants
- Participant coverage: appeared in X of N participants
- First occurrence (earliest `time_in`)
- Last occurrence (latest `time_out`)

Additionally, a sentiment ratio: total negative marks (pain_point) vs. positive marks (delight) as a quick pulse check on the overall session tone.

**Visualization**: a stacked horizontal bar showing all six categories in their mark-category colors, with pain_point and delight at opposite ends so the sentiment balance is immediately readable. Each category segment is proportional to its count.

Computed from `state.trIntakeMarks`.

### 4. Per Observation — Spreadsheet

Aggregate statistics per spreadsheet row (each row represents one observation/finding).

Per observation:
- Total timestamp count across all participants (sum of parsed timestamp segments per cell). Reuses the existing `ROW_FUNCTIONS.Count` pattern from studio.js
- Unique participant count (how many participants have a valid cell for this row). Reuses the existing `ROW_FUNCTIONS.Unique` pattern
- Earliest timestamp (min `startSeconds` across all participants)
- Latest timestamp (max `startSeconds + duration` across all participants)
- Category and severity labels

Computed from `state.sheetData.rows` via `parseClipTimestamps()`.

Table is sortable by any column. Each row includes a small **inline coverage bar** showing participant coverage (X/N) as a filled proportion, so the researcher can scan which observations are widespread vs. isolated.

### 5. Severity Distribution

Count of observations per severity level across the study: Critical, High, Medium, Low, N/A, Positive, Very Positive.

**Visualization**: a horizontal stacked bar where each segment is colored with the corresponding `--sev-*` token from tokens.css and sized proportionally. This is the primary visual for this section — the chart *is* the content, with exact counts shown as labels on or below each segment. Gives an immediate sense of the study's overall severity profile (e.g. "mostly medium with a few criticals" is instantly readable as a shape).

### 6. Category Breakdown — Spreadsheet

Count of observations per spreadsheet `category` value, with participant coverage per category (how many participants have at least one timestamp in that category).

Useful for understanding the shape of the observation taxonomy — whether one category dominates or observations are evenly spread.

**Visualization**: horizontal bars (CSS width percentages) showing relative count per category, sorted by count descending. Each bar includes a subtle coverage indicator showing participant spread.

### 7. Cross-Stream Collisions

How often events from different streams co-occur within a configurable time window. This is a lightweight version of the convergence calculation — useful to see in aggregate before going into the Convergence Browser for deeper exploration.

Three pairwise combinations:

- **Screenspace ↔ Spreadsheet**: how often a screenspace event cluster overlaps with a parsed spreadsheet timestamp (within ±W seconds) for the same participant
- **Screenspace ↔ Transcript**: how often a screenspace event cluster overlaps with a transcript mark for the same participant
- **Transcript ↔ Spreadsheet**: how often a transcript mark overlaps with a parsed spreadsheet timestamp for the same participant

Per pair:
- Total collision count
- Number of participants with at least one collision
- Percentage of stream-A items that have a collision in stream B

The time window W is researcher-adjustable via a numeric input. The right window varies by study — a fast UI interaction has a different meaningful window than a slow narrative moment.

Collision detection uses **clusters** (not raw events) to avoid counting the same temporal moment multiple times. Efficient computation: pre-sort by participant + time, sweep-line merge per participant.

### 8. Session-Level Summary

Per participant, per stream: total counts. This is the primary outlier detection surface.

- **Spreadsheet**: count of valid cells, total timestamp segments
- **Screenspace**: count of events, count of distinct event types
- **Transcript**: count of marks, count per mark category

**Outlier detection**: a participant whose event count in any stream exceeds 3× the median is flagged. The ratio-based threshold (rather than standard deviation) is robust for the small sample sizes typical of UX studies (3–8 participants). Outlier cells receive a visual highlight using the `--sev-high` color token and a warning icon from `assets/icons/`.

**Visualization**: small multiples — one mini bar chart per participant, stacked vertically. Each mini chart has one bar per stream (sheet, screenspace, transcript), colored by stream. All mini charts share the same scale so outliers are immediately obvious as disproportionately long bars. This is more readable than a dense grouped bar chart when there are 6+ participants, and each participant's profile is self-contained.

With a single participant, the summary table still shows data but outlier detection is disabled with a note: "Outlier detection requires multiple participants."

### 9. Temporal Density Histogram

A canvas-rendered histogram showing where activity concentrates across the session timeline. All events from all participants and all streams are binned into time buckets and plotted as vertical bars, giving a bird's-eye view of the study's temporal shape.

**Rendering**: canvas-based (follows existing `renderTimeline()` pattern from screenspace.js). Horizontal axis = session time. Vertical axis = event count per bin. Three layered series — one per stream — stacked or overlaid with transparency so the researcher can see which stream contributes to each peak.

- **Bin width**: auto-calculated from total session duration (aim for ~40–60 bins to fill the canvas width comfortably). Adjustable if needed, but a sensible default avoids an extra control
- **Per-stream coloring**: each stream gets a distinct color with alpha blending. Stacked rendering so peaks show total activity, with color breakdown showing which stream dominates
- **Aggregate across participants**: all participants' events are merged into the same timeline. This is the key difference from the Convergence Browser, which shows per-participant rows
- **Hover interaction**: hovering over a bin shows a tooltip with the exact counts per stream for that time window

This is the visualization that adds the most insight no table can convey. A cluster at minute 5 across all streams says something fundamentally different from an even spread — but a table of bin counts would be unreadable. It also serves as the strongest bridge to the Convergence Browser: "I see a peak at minute 5 — let me go to the CB to see which participants contributed."

Computed from the union of: parsed spreadsheet timestamps (`startSeconds`), screenspace events (`time_in`), and transcript marks (`time_in`).

### Examples of Researcher Queries This Answers

- Which participants am I missing ScreenSpace data for?
- When was a loading screen first detected across all participants?
- How often did the health bar change event fire per session on average?
- What's the average confidence for my text detector events? (QA: low confidence → threshold may need adjustment)
- How many participants had a spreadsheet flag within 5 seconds of a ScreenSpace event of type X?
- How many pain points vs. delights were marked across all sessions?
- Which participants had no transcript tags at all?
- Is one participant's event count anomalously high compared to others?
- Is there a hot spot at minute 5 where all streams show activity?
- Does activity concentrate early in the session or spread evenly?

---

## UI Layout

Two tiers: summary charts at the top for immediate visual orientation, then detailed sections with inline visualizations below.

Top-to-bottom scrolling within the tab panel:

1. **Header bar** — tab title, Refresh button, Export JSON button, Export CSV button, collision window numeric input (W seconds)
2. **Summary charts strip** — a compact row of the highest-signal visualizations, always visible:
   - Temporal density histogram (canvas, full width) — the bird's-eye view of session activity
   - Severity distribution stacked bar — immediate severity profile
   - Session-level small multiples — per-participant bar charts showing outliers at a glance
3. **Data Coverage Matrix** — compact heatmap table, always visible (not collapsible), zero counts highlighted
4. **Per Event Type (Screenspace)** — collapsible section, sortable table with inline count bars
5. **Per Category (Transcript)** — collapsible section with stacked sentiment bar + table
6. **Per Observation (Spreadsheet)** — collapsible section, sortable table with inline coverage bars
7. **Severity Distribution** — collapsible section (detail view of the summary chart, with exact counts)
8. **Category Breakdown** — collapsible section with horizontal bars
9. **Cross-Stream Collisions** — collapsible section, per-pair stat rows
10. **Session-Level Summary** — collapsible section (detail view of the summary small multiples, with full table + outlier flags)

Collapsible sections: click the section header to toggle content visibility. All sections default to expanded on first visit.

The summary charts strip and the corresponding detail sections show the same data at different fidelity — the strip is for scanning, the sections are for reading. This avoids forcing the researcher to scroll past charts to find numbers or vice versa.

CSS uses existing design tokens throughout: `--color-panel-bg` and `--color-border` for section containers, `--text-xs` for table cell text, `--sev-*` tokens for severity colors, `--color-text-dim` for muted empty-state messages. No new tokens needed.

---

## Visualization Approach

Two rendering techniques, chosen by what each chart needs:

### CSS-based charts (simple distributions and proportions)

Used for: severity stacked bar, category bars, inline count bars in tables, coverage heatmap cells, transcript sentiment bar, session small multiples.

Implementation: HTML elements with `width` set to percentage values, `background-color` from design tokens, and flexbox for layout. These auto-inherit theme colors, respond to dark/light mode without extra logic, and require no canvas boilerplate.

**Inline table bars**: within sortable tables (Per Event Type, Per Observation), a narrow horizontal bar is rendered in each count/coverage cell. The bar's width is proportional to the cell's value relative to the column maximum. This provides a scannable magnitude column without a separate chart. Uses a `<span>` with `display: inline-block`, `height: 4px`, `background: var(--color-accent)`, and `width: {pct}%`.

**Stacked bars**: the severity distribution and transcript sentiment bars are single-row flex containers where each segment is a `<div>` with `flex-basis` proportional to its count and `background` from the corresponding color token (`--sev-*` or mark category color). Segments below a minimum width threshold show a tooltip on hover rather than a label, to avoid visual clutter.

**Small multiples**: for the session-level summary, each participant gets a compact row containing one horizontal bar per stream. All rows share the same max scale so relative magnitudes are comparable. Outlier bars receive a `--sev-high` border or background tint. Each row is ~24px tall — the full set of participants stays compact even with 10+ participants.

**Heatmap cells**: coverage matrix cells get `background-color: rgba(accent, t * 0.45)` where `t` is the normalized value within that column. Follows the existing heatmap pattern from studio.js's spreadsheet grid (`line 1239`). Zero-count cells use a distinct muted treatment (e.g. dashed border, dim text) rather than zero opacity.

### Canvas-based chart (temporal data)

Used for: the temporal density histogram only.

Implementation: a single `<canvas>` element rendered using the 2D Canvas API. Follows the existing `renderTimeline()` pattern from screenspace.js: device pixel ratio handling (`canvas.width = w * dpr`), theme color retrieval via `getThemeColors()`, and `ctx.fillRect()` for bar rendering.

**Histogram specifics**:
- Horizontal axis: session time (0 to max timestamp across all data). Time labels at regular intervals (MM:SS format)
- Vertical axis: event count per bin. No explicit Y-axis labels — the shape is the insight, and hover tooltips provide exact values
- Bin count: auto-calculated to produce ~40–60 bins across the canvas width. For a 30-minute session, this means ~30–45 second bins
- Three stacked series (one per stream), each with a distinct color and slight transparency. Stacking order: spreadsheet on bottom (typically sparsest), screenspace in middle, transcript on top
- Stream colors: reuse `--color-accent` variants or introduce three muted stream-identity colors (not task-specific colors, which would conflict with per-event-type coloring elsewhere)
- Hover: show a vertical highlight line and tooltip with per-stream counts for the hovered bin
- Resize: redraw on container resize (follow existing `ResizeObserver` or `getBoundingClientRect` patterns)
- Empty state: if no temporal data is available, show "No events to plot" centered in the canvas area

---

## Interaction

**Compute on demand.** Statistics are computed when the tab is activated and on explicit Refresh. No continuous polling — this avoids the cost of recomputation every 10 seconds and is consistent with the Convergence Browser's planned approach (show a refresh indicator, let the researcher decide when to recalculate).

**Data fetch on activation.** When the tab is selected, trigger a one-time fetch from the screenspace and transcript APIs if `state.intakeEvents` or `state.trIntakeMarks` are stale or empty. Reuse existing state data if already populated from a recent intake tab visit.

**Refresh button.** Re-fetches all three streams and recomputes all statistics.

**Data freshness banner.** When screenspace tasks are still running (detected by checking `/screenspace/api/tasks` for non-completed tasks on tab activation), display a warning banner at the top: "Screenspace analysis is still running — statistics may be incomplete." The banner includes its own Refresh button.

**Participant filter pills.** Reuse the existing pill-based participant filter pattern from the intake tabs (`buildParticipantPills`). Toggling participants updates all statistics sections to include/exclude those participants. Useful for focusing on a subset or excluding a known-bad session.

**Collision window control.** A numeric input for W (seconds) in the header bar. Changing it triggers recomputation of the Cross-Stream Collisions section only — other sections are unaffected.

**Drill-down clicks.** Statistic rows are clickable and navigate to the relevant intake tab with pre-applied filters:
- Event type row → Screenspace Intake tab, filtered to that `event_type` (sets `state.intakeFilterDetector`)
- Transcript category row → Transcript Intake tab, filtered to that `category` (sets `state.trIntakeFilterCategory`)
- Participant name (in session summary or coverage matrix) → Screenspace or Transcript Intake tab, filtered to that participant

Implementation: set the target tab's filter state variables, set `state.activePreviewTab`, and call `syncPreviewTab()`. Future: when the Convergence Browser (Tab 4) is implemented, cross-stream collision rows could additionally navigate there.

---

## Empty States

| Scenario | Behavior |
|---|---|
| No data at all (no sheet, no screenspace, no transcripts) | Single centered message: "No data loaded. Statistics will appear once data is available from at least one stream." |
| Single stream only (e.g. sheet data but no screenspace or transcripts) | Show that stream's sections normally. Other sections show muted text: "No screenspace events available" / "No transcript marks available." Cross-stream collisions section: "Cross-stream collisions require data from at least two streams." |
| Single participant | Show all per-stream statistics normally. Session-level summary visible but outlier detection disabled: "Outlier detection requires multiple participants." Coverage matrix still useful (shows which streams have data). |
| Screenspace tasks still running | Warning banner at top of tab. Statistics show current data with the understanding that more may arrive. |

Use existing `drop-target-empty` CSS patterns from the intake tabs for muted empty-state styling.

---

## Export

Two formats, both generated entirely client-side via Blob download. All data needed is already in JS `state` — no server endpoint required. Consistent with the thin-server principle.

Given that clipgen must fit into many different reporting contexts across many different organisations, minimal friction on export format is a design requirement, not a nice-to-have.

### JSON

A single file named `{study}_metadata.json`:

```json
{
  "study": "Study Name",
  "exported_at": "2025-01-15T14:30:00Z",
  "participants": ["P01", "P02", "P03"],
  "coverage_matrix": {
    "P01": { "spreadsheet": 12, "screenspace": 47, "transcript": 8 },
    "P02": { "spreadsheet": 12, "screenspace": 0, "transcript": 15 },
    "P03": { "spreadsheet": 10, "screenspace": 23, "transcript": 0 }
  },
  "event_type_stats": [
    {
      "event_type": "health_bar_change",
      "detector": "multitool",
      "total_count": 123,
      "participant_coverage": 5,
      "participant_total": 6,
      "first_occurrence_sec": 45.2,
      "last_occurrence_sec": 1834.5,
      "mean_time_sec": 890.3,
      "mean_confidence": 0.87,
      "mean_duration_sec": 3.2,
      "per_participant": {
        "P01": { "count": 20, "mean_time_sec": 850.0 },
        "P02": { "count": 35, "mean_time_sec": 920.1 }
      }
    }
  ],
  "transcript_category_stats": [
    {
      "category": "pain_point",
      "total_count": 15,
      "participant_coverage": 4,
      "participant_total": 6,
      "first_occurrence_sec": 120.0,
      "last_occurrence_sec": 1650.3,
      "per_participant": { "P01": 3, "P02": 5 }
    }
  ],
  "observation_stats": [
    {
      "observation": "Loading screen detected",
      "category": "Performance",
      "severity": "High",
      "total_timestamps": 8,
      "unique_participants": 4,
      "first_occurrence_sec": 120.0,
      "last_occurrence_sec": 1500.0
    }
  ],
  "severity_distribution": {
    "Critical": 3, "High": 12, "Medium": 8, "Low": 5,
    "N/A": 2, "Positive": 4, "Very Positive": 1
  },
  "category_breakdown": {
    "Gameplay": 15, "UI": 8, "Performance": 5
  },
  "cross_stream_collisions": {
    "window_seconds": 5,
    "screenspace_spreadsheet": { "collision_count": 23, "participants_with_collisions": 4 },
    "screenspace_transcript": { "collision_count": 17, "participants_with_collisions": 5 },
    "transcript_spreadsheet": { "collision_count": 8, "participants_with_collisions": 3 }
  },
  "session_summary": [
    {
      "participant": "P01",
      "spreadsheet_valid_cells": 12,
      "spreadsheet_total_timestamps": 15,
      "screenspace_events": 47,
      "screenspace_distinct_event_types": 3,
      "transcript_marks": 8,
      "transcript_by_category": { "pain_point": 3, "delight": 2, "quote": 3 },
      "outlier_flags": ["screenspace_events"]
    }
  ]
}
```

### CSV

Multiple files offered as separate downloads (the hierarchical data does not flatten cleanly into a single table):

**`{study}_metadata_events.csv`** — one row per screenspace event type:
```
event_type,detector,total_count,participant_coverage,participant_total,first_occurrence_sec,last_occurrence_sec,mean_time_sec,mean_confidence,mean_duration_sec
health_bar_change,multitool,123,5,6,45.2,1834.5,890.3,0.87,3.2
loading_screen,scene,12,4,6,180.0,1200.5,650.2,0.93,1.5
```

**`{study}_metadata_sessions.csv`** — one row per participant:
```
participant,spreadsheet_valid_cells,spreadsheet_timestamps,screenspace_events,screenspace_event_types,transcript_marks,transcript_pain_points,transcript_delights,transcript_quotes,transcript_insights,transcript_tasks,transcript_bookmarks,outlier
P01,12,15,47,3,8,3,2,3,0,0,0,false
P02,12,14,0,0,15,5,4,3,1,1,1,false
```

**`{study}_metadata_collisions.csv`** — one row per collision pair:
```
pair,window_seconds,collision_count,participants_with_collisions
screenspace_spreadsheet,5,23,4
screenspace_transcript,5,17,5
transcript_spreadsheet,5,8,3
```

**`{study}_metadata_observations.csv`** — one row per spreadsheet observation:
```
observation,category,severity,total_timestamps,unique_participants,first_occurrence_sec,last_occurrence_sec
Loading screen detected,Performance,High,8,4,120.0,1500.0
Health bar drops to zero,Gameplay,Critical,3,2,450.0,890.0
```

---

## Computation Notes

**Raw events for statistics, clusters for collisions.** Count, mean, and coverage statistics are computed from raw events (`state.intakeEvents` where `excluded !== true`) because raw events are the ground truth. Cross-stream collision detection uses clusters (`state.intakeClusters`, `state.trIntakeClusters`) because consolidated time ranges avoid double-counting the same temporal moment.

**Pre-sort once, compute once, cache.** On tab activation, pre-sort each stream's data by participant + time. Compute all statistics and cache results in a `state.metadataCache` object. Invalidate on Refresh or when participant filters change.

**Performance.** Even with aggressive detector settings (20 participants × 5000 events each = 100K events), all group-by and aggregation operations are sub-50ms. Cross-stream collision with sorted-merge is O(n log n) per participant. No virtualization or lazy rendering needed — the tab renders tables, not hundreds of cards.

**Excluded events.** Screenspace events with `excluded: true` are filtered out, consistent with how the Screenspace Intake tab operates.

**Spreadsheet timestamp parsing.** Uses the existing `parseClipTimestamps()` function which returns `[{ startSeconds, duration }]`. Note: baseline time normalization (converting absolute clock times to video-relative seconds) is currently server-side only (`files.prepare_clip()`). If the Convergence Browser prerequisite work moves baseline correction client-side, this tab should adopt the same correction. Until then, sheet timestamps are displayed as-parsed.

**Temporal histogram binning.** Bin width = `totalDuration / numBins` where `numBins` is chosen to produce ~40–60 bins across the canvas width. Each bin accumulates counts from all three streams independently (for the stacked rendering). The binning pass is a single O(n) iteration over the merged, sorted event list.

---

## Key Design Decisions

**Read-only.** No curation happens in this tab. It is a summary of what has been loaded, not a workspace. Keeping this boundary clean prevents the tab from becoming a parallel intake surface.

**Study-relative.** Like the Convergence Browser, all statistics are derived from whatever event types and labels exist in the loaded JSONs. No assumptions about taxonomy.

**QA function is first-class.** The session-level outlier view is explicitly useful for catching misconfigured detectors before committing to full artifact generation. This should be surfaced clearly, not buried. The data coverage matrix and outlier detection are the two most direct QA tools.

**Configurable collision window.** The time window for cross-stream collision detection should be researcher-adjustable. The right window varies by what is being studied — a fast UI interaction has a different meaningful window than a slow narrative moment.

**Compute on demand, not live.** Statistics are computed when the tab is activated and on manual refresh. No polling. This avoids the cost of continuous recomputation and matches the Convergence Browser's planned approach.

**Client-side computation and export.** All data is already in JS state. No new server endpoints. Export via Blob download. Consistent with thin-server principle.

**Raw events for stats, clusters for collisions.** Different computation units for different questions. Raw events give accurate counts; clusters give meaningful temporal overlap without double-counting.

**Drill-down, not duplication.** The tab presents aggregate statistics that are invisible in the per-event intake views. It does not list individual events — that is what the intake tabs do. Clicking a stat row navigates to the relevant intake tab with pre-applied filters, bridging the two views.

**Charts for shape, tables for precision.** Visualizations are used where the insight is a pattern (distribution shape, proportion, temporal clustering, outlier magnitude) that numbers alone cannot convey. Tables remain for exact values, sortable detail, and export-oriented data. The summary charts strip at the top provides the at-a-glance view; the collapsible detail sections below provide the numbers. Both show the same data at different fidelity.

**CSS charts where possible, canvas only where necessary.** Simple proportional charts (bars, stacked bars, heatmap cells, small multiples) use CSS-based rendering — HTML elements with percentage widths and design token colors. This auto-inherits theming, requires no canvas boilerplate, and is accessible. Canvas is reserved for the temporal density histogram where a continuous axis and smooth rendering are needed.

---

## Integration Points

- Reads from: `state.sheetData`, `state.intakeEvents`, `state.trIntakeMarks` (all already in Studio memory)
- Writes to: JSON and CSV export files via client-side Blob download
- Relationship to Convergence Browser (Tab 4): provides orienting context before the researcher enters the Convergence Browser; the cross-stream collision data here is a lighter-weight preview of what the browser explores interactively. Future: drill-down links from collision stats into the CB with pre-applied filters
- Relationship to intake tabs: drill-down navigation injects filter state and switches to the target tab via `state.activePreviewTab` + `syncPreviewTab()`
- Reuses existing studio.js functions and patterns: `parseClipTimestamps()`, `ROW_FUNCTIONS.Count`, `ROW_FUNCTIONS.Unique`, `buildParticipantPills()`, `SEVERITY_ORDER`, `severityClass()`, `findOverlappingData()`, `hexToRgba()`
- Canvas rendering follows existing patterns from screenspace.js: `renderTimeline()` for the histogram, `getThemeColors()` for adaptive colors, device pixel ratio handling, `ResizeObserver` for responsive sizing
- Relationship to existing artifacts: the metadata export is a companion to clip artifacts, not a replacement — intended to travel alongside them in reporting

---

## Key Findings From Design Discussion

- All data required for this tab is already in Studio memory — this is a computation and display problem only
- Researchers currently have no fast path to aggregate statistics; this must be compiled manually from individual session reviews
- The data coverage matrix (participant × stream) is the single fastest QA check and should be the first thing visible
- The session outlier view (anomalous event counts per participant) doubles as a detector QA tool, catching misconfiguration early in the workflow
- Transcript marks are a full data stream that deserves aggregate coverage alongside screenspace events and spreadsheet observations
- Cross-stream collisions across all three pairwise combinations (SS↔Sheet, SS↔Transcript, TR↔Sheet) give a complete picture of where streams agree
- Export format flexibility is a hard requirement given clipgen's need to fit into diverse organisational reporting contexts
- This tab is lower implementation complexity than the Convergence Browser and could be built first, with its collision data informing what queries to prioritise in the browser
