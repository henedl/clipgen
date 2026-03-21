# Insights Feature Plan

## Problem

clipgen serves UX researchers well as a workbench for generating artifacts from research sessions. However, there is no structured path from raw artifacts to a deliverable that stakeholders can consume. Researchers currently assemble findings outside of clipgen (slide decks, reports), losing the connection between the narrative and the underlying evidence.

## Design Principle

The value of research is the interpretation, not the data. Stakeholders should experience findings through the researcher's curated narrative, not by browsing raw clips. The tool should make the researcher's argument legible, not replace their analytical role.

The unit of persuasion is a **narrative arc**: cause → behavior → impact, supported by 1-3 examples per bucket. If a reader can see what caused a behavior, witness the behavior, and understand its impact — with clear connections between them — they are typically convinced enough to act.

**"The software works at the speed of the researcher."** Artifact generation and insight authoring are architecturally separate, but the boundary must be invisible to the researcher. The builder has full, instant access to all available artifacts — browsable, filterable, searchable, draggable — so the researcher never leaves the builder to find evidence. Two tools, one fluid experience.

## Concept: Insights

An **Insight** is a researcher-authored finding composed of:

- **Title** — a concise claim (e.g. "Players don't engage with social features")
- **Summary** — a researcher-written description of why this matters, summarizing key points (including suggestions and references to previous findings)
- **Severity** — how critical the finding is (using the existing severity system: Critical/High/Medium/Low/N/A/Positive/Very Positive)
- **Status** — draft or final
- **Three evidence buckets:**
  - **Causes** — what led to the observed behavior (artifacts + researcher narrative)
  - **Behaviors** — what users actually did (artifacts + researcher narrative)
  - **Impacts** — consequences of the behavior (artifacts + researcher narrative)
- **Narrative annotations** — researcher-written text in each bucket connecting the evidence and telling the story. One narrative block per bucket. Causes and impacts are where the researcher's analysis matters most.
- **Timeline context** — where the finding occurred in the general session flow (text input for now)

Insights reference artifacts by ID from the existing manifest. Artifacts don't need to know they're part of an insight — the insight layer is additive, not a modification of the existing system.

Insights can span multiple participants and sessions. For now, they are limited to a single study.

## Three-Tier Deliverable

The reader experience is structured as three tiers of increasing detail:

1. **Insights index** — the landing page. Lists all findings with title, severity, and summary. Answers: "what did you find?"
2. **Insight detail** — drill into one finding. Cause → behavior → impact structure with inline artifacts and researcher narrative. Answers: "why should I believe this and why does it matter?"
3. **Full artifact view** — the existing timeline/gallery viewer, accessible via tab or link. Answers: "show me everything."

Most stakeholders live in tiers 1 and 2. Curious readers can "back out" to tier 3 for the complete evidence set. This respects the researcher's interpretive authority while giving readers an escape valve for deeper exploration.

## Workflow

1. Researcher generates artifacts across sessions as they do today (clipgen always runs "after" observations are collected in the spreadsheet)
2. Researcher opens the insight builder, which **immediately loads all artifacts from the manifest** — browsable, filterable, searchable, and ready to drag into buckets. The researcher arranges evidence into cause → behavior → impact buckets and adds narrative annotations without leaving the builder.
3. Researcher saves their work and exports/generates the Insight view
4. Reader receives the deliverable and navigates the three tiers

## Architecture

- **`insights_manifest.json`** — persistence layer alongside `clipgen_manifest.json`. Each insight has an ID, title, summary, severity, status (draft/final), three buckets with artifact ID references and narrative text, timeline context, and timestamps (creation date, update date).
- **Insight builder interface** — web view served by `insights_server.py` (Flask). Opens with full awareness of the artifact manifest — all artifacts are immediately browsable, filterable, searchable, and draggable. The builder is a **consumer** of artifacts, not a producer; artifact generation stays in clipgen's existing CLI/batch flow, keeping both sides simple. Supports multiple insights in parallel and across sessions. Must be faster than making a slide deck.
- **Insight viewer** — standalone HTML deliverable (inlined CSS/JS, embedded data) following the same pattern as the timeline and gallery viewers, using the shared `_generate_viewer_html()` function in `viewer.py`.
- **Deliverable format** — a folder/bundle (HTML files + media assets) that can be shared via zip, shared drive, or eventually a hosted URL.

## Thumbnail Sprite Sheet Infrastructure

Sprite sheets enable hover-to-scrub previews in the media library and inline thumbnails in bucket artifacts. They are generated via ffmpeg and stored as single PNG files alongside clip artifacts.

### How it works

1. **Generation**: `video.generate_sprite_sheet(input_file, output_file, ...)` extracts frames at intervals and tiles them into a single PNG image.

2. **ffmpeg command pattern**:

   ```bash
   ffmpeg -i input.mp4 -vf "fps=1/{interval},scale={thumb_width}:-1,tile={cols}x{rows}" -frames:v 1 output_sprite.png
   ```

3. **Interval calculation**: `interval = max(1, clip_duration // target_frame_count)` where `target_frame_count` defaults to 20. For very short clips (under 5s), use 1 frame per second.

4. **Grid layout**: `cols` and `rows` computed from frame count (e.g. 15 frames → 5x3 grid).

5. **Storage**: sprite sheet PNGs saved alongside clip artifacts in the output directory. Naming: `{clip_basename}_sprite.png`.

6. **Metadata**: the existing `thumbnail` field in artifact records (currently always `""`) is repurposed for the sprite sheet filename. A `spriteData` object provides grid dimensions:

   ```json
   {
     "thumbnail": "clip_name_sprite.png",
     "spriteData": {
       "cols": 5,
       "rows": 3,
       "frameCount": 15,
       "frameWidth": 160,
       "frameHeight": 90,
       "interval": 2
     }
   }
   ```

7. **When generated**: on builder startup (blocking), before opening the browser. The server iterates clip artifacts and generates sprite sheets for any that lack them. Once a sprite sheet exists on disk, it is served directly without regeneration.

8. **Frontend hover-to-scrub**: mouse X position over the thumbnail maps to a frame index. The sprite sheet is set as CSS `background-image`, and `background-position` is updated on `mousemove`:

   ```css
   frameIndex = floor((mouseX / cardWidth) * frameCount)
   col = frameIndex % cols
   row = floor(frameIndex / cols)
   backgroundPosition = -(col * frameWidth) + "px " + -(row * frameHeight) + "px"
   ```

### Config constants

```python
SPRITE_SHEET_FRAME_COUNT = 20      # Target frames per sprite sheet
SPRITE_SHEET_THUMB_WIDTH = 160     # Pixel width per frame
SPRITE_SHEET_MIN_INTERVAL = 1      # Minimum seconds between frames
```

## Insights Builder Details

### Overall Layout

```ascii
+------------------------------------------------------------------+
| HEADER: "clipgen Insight Builder"   [Save] [Generate Viewer] [O] |
| N artifact(s) | M insight(s) | [unsaved indicator]               |
+--------------------+---------------------------------------------+
|                    |                                              |
| MEDIA LIBRARY      | INSIGHT EDITOR                              |
| (resizable)        |                                              |
|                    |                                              |
| [Filters panel]    | [+ New Insight]                             |
| [Search box]       |                                              |
|                    | [Insight Card 1 - collapsed/expanded]        |
| [Artifact grid]    | [Insight Card 2 - collapsed/expanded]        |
|                    | [Insight Card N]                             |
|                    |                                              |
+--------------------+---------------------------------------------+
```

Two-column layout: resizable media library sidebar on the left, insight editor main area on the right.

Each insight card has three horizontal areas -  Cause, Behavior, Imapct -  with an editable summary row above each.

### Media Library Panel (left sidebar)

The media library takes inspiration from video editing suites. It is a visual-first browsing experience for all artifacts in the manifest.

**Resizable behavior:**

- Default width: ~420px
- Minimum width: 280px
- Maximum width: 50% of viewport
- Draggable resize handle on the right border (visible on hover)
- Collapse button that shrinks to a narrow icon strip (~48px)
- Width persisted to `localStorage`

**Filter panel** (top of sidebar, collapsible):

- Participant dropdown — "All participants" default
- Category dropdown — "All categories" default
- Severity dropdown — "All severities" default, hidden if no severity data exists
- Type checkboxes: Clip, Screenshot, GIF — hidden if only one type exists
- Text search input below the structured filters
- Filters and search combine: an artifact must match all active filters AND the search query
- Filter implementation ports the proven pattern from the timeline viewer (`viewer.js`): `populateFilters()` extracts unique values, `fillSelect()` builds option elements, `applyFilters()` reads all filter values and filters the array, `bindFilterEvents()` attaches change listeners. Filter state is purely client-side.

**Artifact card grid** (scrollable area below filters):

- Cards in a responsive layout (single column in narrow sidebar, 2-column grid when widened)
- Each card shows:
  - **Media area**: For clips, a div with sprite sheet as `background-image`. On hover, mouse X position maps to frame index and updates `background-position` for scrub preview. On mouse leave, resets to first frame. For screenshots and GIFs, a standard `<img>` element.
  - **Meta badges**: participant badge, type badge (colored by type), category badge, severity badge (colored by severity)
  - **Description**: truncated observation text (~60 chars)
  - **Time range**: start–end timestamps
- Cards are `draggable="true"` for drag-and-drop into insight buckets
- Clicking a card shows a popover to choose which bucket (and which insight, when multiple are expanded) to add to

### Insight Editor (main area)

**Toolbar**: "+ New Insight" button at top.

**Insight cards** stack vertically with spacing. New insights are added below existing ones. Artifacts can be dragged from the media library into any insight's buckets — the researcher doesn't need to target the "active" insight.

**Collapsed state**: single row showing chevron, title, severity pill, status badge, artifact count.

**Expanded state**:

1. **Field row** (three columns): title text input, severity dropdown, status dropdown (draft/final).

2. **Summary textarea**: multi-line text area for the researcher's overall description of the insight. Placeholder: "Describe the key finding and why it matters..."

3. **Three bucket sections** (causes, behaviors, impacts), each containing:

   **Bucket label**: colored heading with a border-bottom in the bucket's color (orange for causes, blue for behaviors, red for impacts).

   **Narrative textarea**: researcher's written explanation for this bucket. Placeholder text guides the researcher.

   **Artifact drop zone with inline preview cards**: this is where evidence lives. Each artifact in the bucket renders as a card (~200px wide, arranged in a wrapping flex row) showing:
   - Thumbnail image (first frame from sprite sheet for clips, scaled image for screenshots, the GIF itself for GIFs)
   - Meta badges (participant, type, category)
   - Full description text (not truncated)
   - Time range (start–end)
   - Remove button (×) in top-right corner

   Artifacts within a bucket are **drag-reorderable** — the order in the array determines display order in both the builder and the viewer.

   The drop zone accepts drags from the media library sidebar. Visual feedback (dashed border highlight) indicates valid drop targets during drag operations.

4. **Timeline context**: text input field. Placeholder: "e.g. During onboarding (first 5 minutes)".

5. **Footer**: delete button with confirmation dialog.

### Save Behavior

Manual save only. No auto-save, no debounce timers.

- **Global Save button** in the header (primary style). Saves all insights with unsaved changes.
- **Per-card save button** on each insight card for granular control.
- **Dirty tracking**: a `dirtyIds` set in client state. Any edit adds the insight's ID to the set. The header shows an "Unsaved changes" indicator when the set is non-empty. Saving clears the set.
- **Keyboard shortcut**: `Ctrl/Cmd+S` triggers global save.
- **Navigation guard**: `window.onbeforeunload` returns a prompt when `dirtyIds` is non-empty, preventing accidental data loss.

### Click-to-Add Popover

Clicking an artifact card in the media library shows a popover with bucket buttons (Causes, Behaviors, Impacts). When exactly one insight is expanded, the popover targets that insight directly. When multiple insights are expanded, the popover includes an insight selector (showing titles) so the researcher can choose which insight to add to.

### Theme Toggle

Light/dark theme toggle, preference persisted to `localStorage`. Shared pattern across all clipgen viewers.

## REST API

The builder is served by `insights_server.py` (Flask).

| Method | Path | Description |
| -------- | ------ | ------------- |
| GET | `/` | Serve `insights-builder.html` |
| GET | `/<filename>` | Serve static assets (CSS, JS) |
| GET | `/media/<filename>` | Serve video/image artifacts and sprite sheets from output directory |
| GET | `/api/artifacts` | List all artifacts from manifest, enriched with `thumbnail` and `spriteData` |
| GET | `/api/insights` | List all insights from insights manifest |
| GET | `/api/insights/<id>` | Fetch single insight |
| POST | `/api/insights` | Create new insight (accepts title, summary, severity, etc.) |
| PUT | `/api/insights/<id>` | Update insight fields |
| DELETE | `/api/insights/<id>` | Delete insight |
| POST | `/api/sprites/generate` | Generate missing sprite sheets for clip artifacts |
| POST | `/api/generate-viewer` | Finalize insights data and generate standalone viewer HTML |

## Insights Viewer Details

- The Insights Viewer is a standalone HTML deliverable, a single HTML file with inlined CSS/JS and embedded data.
- Two-tier navigation: insights index (card grid) → insight detail (full page with buckets and artifacts).
- The index shows title, severity pill, summary preview (~120 chars truncated), and artifact count per insight.
- The detail view shows title, severity, full summary, and three bucket sections (only shown if they have content) with narrative and artifact grid including inline media.
- Footer links to the timeline viewer HTML if available.
- Only "final" insights are shown in the viewer if any exist; otherwise all drafts are shown.

## Implementation Sequence

1. Build the thumbnail sprite sheet infrastructure (`config.py`, `video.py`, `insights_server.py`)
2. Update the insight data model — add summary field (`insights.py`, `insights_server.py`)
3. Build the builder frontend — media library panel with resizable sidebar, filtering, sprite-based hover-to-scrub (`insights-builder.html/js/css`)
4. Build the builder frontend — insight editor with summary field, inline preview cards in buckets, drag-reorder, manual save (`insights-builder.html/js/css`)
5. Update the viewer to render the summary field (`insights-viewer.html/js/css`)
6. Wire "back out to full view" links between insight detail and existing timeline viewer

## Learnings From First Build

These inform the design decisions above:

1. **320px sidebar was too narrow** — artifact cards need visual space to be useful. The resizable sidebar with a ~420px default fixes this.
2. **Auto-save debounce timers were fragile** — per-insight timers, save status indicators, and race conditions between rapid edits added hidden complexity. Manual save with dirty tracking is simpler and gives researchers explicit control.
3. **Chips were too abstract** — showing `P01 | clip | First time onb...` as a tiny badge told the researcher almost nothing. They need to see the actual media in context. Inline preview cards solve this.
4. **The Flask server pattern works well** — REST API structure, module-level artifact state, and separation between insights persistence and server routing are clean. No architecture changes needed.
5. **The viewer.js filtering pattern is proven** — `populateFilters` / `applyFilters` / `initTypeFilters` / `bindFilterEvents` handles all edge cases. Port it rather than reinventing.
6. **Sprite sheets are the right primitive for hover previews** — loading a full `<video>` element for hover-to-play in a list of 50+ artifacts is expensive. A single sprite sheet PNG per clip is lightweight, cacheable, and gives instant scrub preview via CSS `background-position`.
7. **The `thumbnail` field already exists** in artifact records (always `""`) — repurposing it for sprite sheet filenames is zero-cost migration.
8. **Builder is a consumer, not a producer** — the builder reads from `clipgen_manifest.json` and writes to `insights_manifest.json`. These are separate persistence layers. The builder never modifies the artifact manifest.

## Deferred Features

Explicitly out of scope for now:

- **Auto-save** — removed from scope. Revisit later if manual save proves tedious.
- **Connection drawing between artifacts** — visual lines linking cause artifacts to behavior artifacts across buckets. Interesting idea but complex.
- **Timeline context painting** — interactive timeline where researchers paint affected time ranges. Currently a text field.
- **Clip request from builder** — shell out to clipgen's generation pipeline without leaving the builder. Deferred until we validate whether full manifest access is sufficient.
- **Multi-study insights** — insights currently scoped to single study/manifest.
- **Viewer redesign** — the existing viewer works. Only adding summary field rendering for now.

## Assumptions to Monitor

- **Researcher adoption**: the builder must feel like a single workspace, not a second tool requiring setup or context-switching. If researchers feel they're "switching tools," adoption will suffer regardless of feature quality. The builder must be faster than the slide-deck alternative.
- **Artifact availability gap**: if researchers frequently realize they need clips that don't exist yet while authoring insights, the "consumer only" builder design may need a lightweight clip-request mechanism. Monitor how often this happens before building it.
- **Three-bucket generality**: cause/behavior/impact fits usability and playtest research well. May need configurable buckets for broader research types (attitudinal, competitive, survey).
- **Sharing friction**: the deliverable is a folder of files, not a single file. Sharing needs to be easy (zip, shared drive, or hosted).
- **Sprite sheet size**: for studies with many clips, sprite sheet generation on startup may take noticeable time. Monitor whether this needs to become background/async.
