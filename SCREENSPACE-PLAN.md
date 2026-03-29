# Screenspace Plan

## Overview

Screenspace is a post-processing feature set that performs image analysis on source video files to find and label moments matching specific visual conditions. It produces a duct-taped form of telemetry for software that lacks traditional telemetry — searching for visual changes, color states, text appearances, and similarity matches across long video recordings.

Results are displayed on a timeline in the Screenspace UI. Completed analysis produces **events** — timestamped detections with metadata — that flow via `screenspace_manifest.json` to Studio (for clip generation) and the Timeline Viewer (for visualization).

### Use cases driving this work

1. **State monitoring** — detect when a UI element enters a specific visual state (e.g. a health bar turning red, a meter emptying)
2. **Change tracking** — find every moment a screen region changes (e.g. a scoreboard updating, a notification appearing)
3. **Visual search** — find frames that look similar to a reference moment (e.g. "show me every time the inventory screen looked like this")
4. **Text detection** — find when specific text appears on screen (e.g. a chat message, an error dialog, a score value)
5. **Region timelapse** — generate a sped-up view of a screen region over time (e.g. a minimap timelapse, a resource counter progression)
6. **Numeric monitoring** — find when a numeric value in a screen region meets a condition (e.g. a score exceeding 1000, health dropping below 20)

### Key architectural decisions

- **Combined server model.** Screenspace runs within the existing combined Flask server (alongside Studio and Insights), registered as its own blueprint at `/screenspace/`. A background worker thread handles compute-intensive analysis tasks.
- **Per-participant source videos.** Screenspace operates on `{study}_{participant}.mp4` files, consistent with clipgen's participant model. The user selects a participant from a dropdown.
- **Named region definitions.** Rectangular screen regions can be named, saved, and reused across participants and studies. Game UIs are typically positionally consistent across sessions, so a region like "healthbar" defined once applies to all participants.
- **Task queue with single worker.** Analysis tasks are queued and processed sequentially by a background worker thread. The queue is visible in the frontend with reorder and cancel controls.
- **Separate manifest.** Screenspace persists tasks, region definitions, and results to `screenspace_manifest.json`, independent of `clipgen_manifest.json`.
- **OpenCV as the analysis engine.** `opencv-python-headless` (~40MB, pip-installable, no system dependencies) handles color analysis, frame differencing, similarity comparison, and frame extraction. ffmpeg handles timelapse generation via crop filters.
- **EasyOCR for text and numeric detection.** Lazy-imported so users who don't use OCR features never need PyTorch installed. Models downloaded on first use (~100MB for English). Used by both the Text and Numbers workflows.
- **Compression-aware thresholds.** All comparisons apply Gaussian blur and noise thresholds to handle H.264/H.265 compression artifacts. Default pixel-diff threshold of 30, SSIM threshold of 0.90.
- **Screenspace UI has its own timeline.** Results are displayed on a built-in timeline with zoom, pan, and click-to-seek. In/out markers restrict analysis to sub-ranges.
- **Event data model.** Task results are converted to `ScreenspaceEvent` records — raw point detections stored in the manifest. Clustering into clip-worthy spans happens downstream at Studio ingest time, keeping the manifest as ground truth.

---

## Phase 1: Infrastructure

Build the analysis engine, server integration, and task queue.

### Analysis module (`screenspace.py`)

Core image analysis functions, all operating on NumPy arrays (OpenCV frames):

- [x] `extract_region(frame, region)` — crop a rectangular region from a frame; region is `{x, y, w, h}` in pixels
- [x] `average_color_hsv(region_pixels)` — compute mean HSV color of a cropped region; return `{h, s, v}` values
- [x] `color_matches(region_pixels, target_color, tolerance)` — check if region's average color is within tolerance of a target; comparison in HSV space (hue-aware, handles red wraparound)
- [x] `compute_frame_diff(region_a, region_b, noise_threshold=30)` — absolute pixel difference between two same-sized regions; apply Gaussian blur first; return change ratio (0.0–1.0)
- [x] `regions_are_similar(region_a, region_b, threshold=0.90)` — SSIM-based similarity check with blur preprocessing; return boolean + score
- [x] `compute_phash(region_pixels)` — perceptual hash of a region for fast similarity scanning; return hash object
- [x] `scan_video_frames(video_path, region, interval_seconds, callback)` — iterate through video at interval, extract region, call callback per frame; yield `(timestamp, frame_data)` pairs; use OpenCV `VideoCapture` in sequential mode
- [x] `build_timelapse_command(video_path, region, speedup_factor, output_path, output_format)` — construct ffmpeg argv for cropped timelapse (video or GIF); use `crop=w:h:x:y` + `setpts` filters

### Analysis workflows

Each workflow wraps the primitives above into a complete scan-and-report operation:

- [x] **Color scan** (`scan_color`): iterate frames at interval, check `color_matches()` on region, collect timestamps where condition is met; return list of `{start, end, duration}` spans (merge consecutive matches into spans)
- [x] **Change scan** (`scan_changes`): compare region across consecutive sampled frames, flag timestamps where `compute_frame_diff()` exceeds threshold; return list of change-point timestamps with magnitude
- [x] **Similarity scan** (`scan_similarity`): given a reference frame's region, scan video for frames where `regions_are_similar()` or phash distance is below threshold; return ranked list of `{timestamp, score}` matches
- [x] **Text scan** (`scan_text`): iterate frames at interval, run EasyOCR on region, match extracted text against search string using fuzzy matching (`difflib.SequenceMatcher`, ratio > 0.8); return list of `{timestamp, text_found, confidence}` matches; EasyOCR lazy-imported with clear error if missing
- [x] **Numbers scan** (`scan_numbers`): iterate frames at interval, run EasyOCR on region, parse numeric values, compare against target with operator (eq/gt/lt/gte/lte/range); return list of `{timestamp, value}` matches; shares EasyOCR dependency with text scan
- [x] **Timelapse generation** (`generate_timelapse`): run ffmpeg crop+speed command via `video.run_ffmpeg_process()`; return output file path

### Task queue

- [x] `ScreenspaceWorker` — background thread that pulls tasks from a `queue.PriorityQueue` and executes them; single worker by default
- [x] Task lifecycle: `queued` → `running` → `completed` / `failed` / `cancelled` / `paused`
- [x] Each task is a dict: `{id, type, participant, video_path, region, parameters, status, progress, result, created_at}`
- [x] Progress reporting: worker updates task `progress` (0.0–1.0) as frames are processed; frontend polls for updates
- [x] Cancellation: set a `cancelled` flag on the task; worker checks flag between frame iterations and aborts early
- [x] Queue reordering: tasks can be reordered by adjusting priority values while status is `queued`
- [x] Pause/resume support: worker thread can be paused and resumed; partial results are preserved and scan resumes from last position
- [x] Task dismissal: `remove_task()` for permanently removing completed/failed/cancelled tasks from the queue
- [x] Partial result accumulation: results collected before a pause are preserved; scan resumes from where it left off

### Manifest (`screenspace_manifest.json`)

- [x] Schema:
  ```json
  {
    "regions": {
      "healthbar": {"x": 100, "y": 20, "w": 300, "h": 30, "description": "Player health bar"},
      "scoreboard": {"x": 500, "y": 0, "w": 200, "h": 50, "description": "Score display"}
    },
    "tasks": [
      {
        "id": "ss_<8hex>",
        "type": "color|change|similarity|text|numbers|timelapse",
        "participant": "P01",
        "source_video": "study_P01.mp4",
        "region": "healthbar",
        "parameters": {},
        "status": "completed",
        "created_at": "iso8601",
        "completed_at": "iso8601",
        "results": []
      }
    ],
    "events": []
  }
  ```
- [x] Regions are top-level and shared across tasks (referenced by name)
- [x] `load_screenspace_manifest()` / `save_screenspace_manifest()` — read/write with merge semantics (match existing manifest patterns)
- [x] Manifest lives in the output directory alongside `clipgen_manifest.json`

### Server integration

- [x] Create `screenspace_server.py` with `screenspace_bp = Blueprint("screenspace", __name__)`
- [x] `_init_screenspace_state()` — initialize worker thread, load manifest, resolve participant video paths
- [x] Register blueprint in `start_combined_server()` at `/screenspace/` prefix; blueprint is always registered (not conditional on `--screenspace`)
- [x] Add `--screenspace` CLI flag in `cli.py`; can combine with `-s`, `-i/-o`, `-v`; works with or without spreadsheet (auto-discovers participant videos); cannot combine with mode flags
- [x] Add `SCREENSPACE_MANIFEST_FILENAME` to `config.py`
- [x] Static file serving: serve `screenspace.html` at `/screenspace/`
- [x] API endpoints:
  - `GET /api/participants` — list participants with source video availability
  - `GET /api/video/frame/<participant>/<timestamp>` — extract and return a single frame as JPEG at given timestamp
  - `GET /api/video/info/<participant>` — return video metadata (duration, resolution, fps)
  - `GET /api/regions` — list saved region definitions
  - `POST /api/regions` — create/update a named region
  - `DELETE /api/regions/<name>` — delete a region definition
  - `POST /api/tasks` — enqueue a new analysis task
  - `GET /api/tasks` — list all tasks with status and progress
  - `GET /api/tasks/<id>` — get task detail including results
  - `DELETE /api/tasks/<id>` — cancel a queued/running task; `?dismiss=true` for full removal
  - `PUT /api/tasks/reorder` — reorder queued tasks
  - `POST /api/tasks/pause` — pause the task queue
  - `POST /api/tasks/resume` — resume the task queue
  - `GET /api/tasks/<id>/results` — get task results (timestamps, artifacts)
  - `GET /media/<filename>` — serve generated timelapse files

---

## Phase 2: Frontend

Build the Screenspace web interface.

### Video display and region selection

- [x] Video frame viewer: display frames from the selected participant's source video; frame fetched via `/api/video/frame/<participant>/<timestamp>`
- [x] Frame navigation: scrub through video by timestamp using a slider or direct input; display current frame
- [x] Participant selector: dropdown listing participants with available source videos
- [x] Rectangular region selection tool: click-and-drag overlay on the video frame to define a region; show coordinates and dimensions
- [x] Named regions: save a selected region with a name; load/apply saved regions from the dropdown; edit and delete existing regions
- [x] Region visualization: draw saved regions as labeled overlays on the current frame, each in a distinct color

### Toolbar and workflow configuration

- [x] Workflow selector: toolbar below the video with buttons/tabs for each workflow type (Color, Change, Similarity, Text, Numbers, Timelapse)
- [x] Per-workflow parameter panels:
  - **Color**: target color picker (HSV), tolerance slider, sample interval; includes pipette tool (eyedropper from frame) and "From Region" button (samples average color of active region)
  - **Change**: sensitivity threshold slider, noise threshold slider, sample interval
  - **Similarity**: reference frame selector (current frame becomes reference), similarity threshold, sample interval
  - **Text**: search string input, fuzzy match threshold, sample interval, language selector
  - **Numbers**: operator dropdown (eq/gt/lt/gte/lte/range), target value or min/max range inputs, sample interval
  - **Timelapse**: speedup factor, output format (video/GIF)
- [x] Multi-participant picker: checkbox dropdown with "Select all / Deselect all" toggle for running tasks across participants
- [x] Multi-region picker: checkbox dropdown with color-coded dots, auto-selects active region
- [x] Time range selection: optional in/out markers on the timeline to restrict analysis to a sub-range; default is full video duration
- [x] "Run" button to enqueue the configured task

### Timeline

- [x] Horizontal timeline bar showing the full duration of the selected participant's video
- [x] Result markers: completed tasks' result timestamps shown as markers/spans on the timeline; color-coded by task type
- [x] In/out markers: draggable markers to set analysis time range
- [x] Click-to-seek: clicking a point on the timeline updates the video frame viewer to that timestamp
- [x] Zoom and pan: ability to zoom into a time range for dense results

### Task queue panel

- [x] Queue display: list of all tasks (queued, running, completed, failed, paused) with status indicators
- [x] Progress bar for the currently running task
- [x] Spinner/indeterminate progress when frame count is unknown (shows spinner + elapsed duration)
- [x] Drag-to-reorder queued tasks (queued tasks reorderable among themselves, finished tasks reorderable in their zone)
- [x] Cancel button for queued and running tasks
- [x] Click a completed task to load its results onto the timeline
- [x] Task detail view: show parameters, region used, result count, timestamps
- [x] Pause/resume queue controls
- [x] Task filter buttons (completed/failed visibility toggle)
- [x] Edit button to restore workflow parameters from a completed task
- [x] Retry button for failed tasks

### Results display

- [x] Results list: scrollable list of timestamps/spans returned by a completed task
- [x] Click a result to seek the video frame viewer to that timestamp
- [x] For color/change/similarity tasks: show the match score or change magnitude alongside each timestamp
- [x] For text tasks: show the OCR'd text and confidence alongside each timestamp
- [x] For numbers tasks: show the detected numeric value alongside each timestamp
- [x] For timelapse tasks: inline video/GIF player showing the generated timelapse
- [x] Export results: button to download results as JSON or CSV

---

## Phase 3: Integration and Export

Connect Screenspace detections with Studio and the Timeline Viewer. Screenspace acts as a **detection layer** that emits events — timestamped observations stored as raw points in the manifest. Studio reads these events, clusters them into clip-worthy spans, and feeds them into the generation pipeline.

### Event data model

Events are the bridge between Screenspace detection and the rest of clipgen. Each event represents a single detection at a point in time.

```python
ScreenspaceEvent:
    id: str                      # stable uuid (e.g. "ev_<8hex>")
    source_video: str            # e.g. "study_P01.mp4"
    participant: str             # e.g. "P01"
    detector: str                # analysis tool: "color", "change", "similarity", "text", "numbers"
    event_type: str              # user-defined semantic label: "low_health", "death_screen", etc.
    time_in: float               # seconds (== time_out for point events)
    time_out: float              # seconds (== time_in for point events)
    confidence: float            # normalized 0.0–1.0
    metadata: dict               # detector-specific payload (magnitude, ocr_value, region rect, etc.)
    excluded: bool               # soft-delete; excluded events hidden from Studio
    task_id: str                 # originating task ID for provenance
    region: str                  # region name used in detection
```

Key design decisions:
- **Events are raw point events** (`time_in == time_out`). Clustering into spans is a post-processing step at Studio ingest, not baked into the manifest. The manifest is ground truth.
- **`detector` and `event_type` are separate fields.** The same detector can produce multiple event types; the same event type could come from different detectors. This enables flexible filtering in Studio.
- **`excluded` is a soft delete.** Users can remove events from the active set in Screenspace without destroying data. Studio shows only non-excluded events.
- **Events are per-source-video.** `source_video` and `participant` are on each event.

Manifest `events` key (added to `screenspace_manifest.json`):
```json
{
  "events": [
    {
      "id": "ev_a1b2c3d4",
      "source_video": "study_P01.mp4",
      "participant": "P01",
      "detector": "change",
      "event_type": "hud_update",
      "time_in": 50.0,
      "time_out": 50.0,
      "confidence": 0.85,
      "metadata": {"magnitude": 0.12},
      "excluded": false,
      "task_id": "ss_12345678",
      "region": "healthbar"
    }
  ]
}
```

### Event generation from task results

- [ ] Events generated automatically when a task completes; each matched frame becomes one event
- [ ] **Color scan**: preserve raw per-frame match timestamps before span merging; each matched frame → one event with `confidence` from color match score
- [ ] **Change scan**: each change-point → one event with `confidence` from magnitude
- [ ] **Similarity scan**: each matching frame → one event with `confidence` from SSIM score
- [ ] **Text scan**: each OCR match → one event with `confidence` from fuzzy match ratio, `metadata.text_found`
- [ ] **Numbers scan**: each numeric match → one event with `confidence` from OCR confidence, `metadata.value`
- [ ] **Timelapse**: no events (produces a video file, not detections)
- [ ] `create_event()` helper in `screenspace.py`; event generation in worker's task completion path
- [ ] `confidence` normalization: each detector maps its native score to 0–1

### `event_type` in workflow configuration

- [ ] New optional "Event label" text input in each workflow's parameter panel (placeholder: e.g. "low_health", "death_screen")
- [ ] Maps to `event_type` on generated events; if blank, defaults to `"{detector}: {region}"` (e.g. "change: healthbar")
- [ ] Passed through task parameters in `screenspace_server.py` to `screenspace.py`

### Event exclusion in Screenspace

- [ ] After a task completes, results panel shows events with per-row exclude/include toggle
- [ ] Excluded events are dimmed but still visible; "Show excluded" toggle in results header
- [ ] Exclusion changes persisted to manifest immediately
- [ ] API endpoints:
  - `PUT /api/events/<event_id>/exclude` — set `excluded: true`
  - `PUT /api/events/<event_id>/include` — set `excluded: false`
  - `GET /api/events` — list all events, with optional `?excluded=false` filter
  - `PUT /api/events/bulk-exclude` — body: `{"ids": [...]}`
  - `PUT /api/events/bulk-include` — body: `{"ids": [...]}`

### Studio Intake section

Studio reads non-excluded events from `screenspace_manifest.json` and displays them in a new **"Screenspace Intake"** section in the bottom panel.

- [ ] **Clustering at Studio ingest**: raw point events grouped into spans using a configurable proximity threshold (default 5s); each cluster becomes one intake card showing the span's time range
- [ ] Clustering threshold configurable via a small control in the Intake section header; clusters recalculate on change
- [ ] Single isolated events get ±padding (default 5s) to form a clip-worthy span; clamped to `[0, videoDuration]`
- [ ] **Intake card rendering**: participant badge, time range, event type label, detector badge (color-coded), region name, event count, confidence indicator, "Add to Artifacts" and "Add to Reel" buttons, dismiss button
- [ ] **"New events" indicator**: events not previously seen by Studio are highlighted with a subtle badge; clears on interaction
- [ ] **Bulk actions**: "Add All to Artifacts", "Add All to Reel" buttons in section header
- [ ] **Data flow**: Studio fetches `GET /screenspace/api/events?excluded=false` on load and periodically (every 10s when Screenspace is active); clustering happens client-side
- [ ] Section hidden when no events exist
- [ ] New config constant: `SCREENSPACE_INTAKE_CLUSTER_SECONDS = 5`

### Studio clip generation from intake

- [ ] New endpoint `POST /studio/api/generate-intake` generates clips directly from intake spans; resolves source video via participant, calls `video.cut_clip()` directly (bypasses spreadsheet pipeline)
- [ ] Artifact IDs use `intake_{hash}_s{seg}` format
- [ ] Reel integration: intake items can be added to the Reel area; `POST /studio/api/reel` accepts optional `intake_items` array
- [x] Quick link from Studio to open Screenspace for the selected participant (already implemented as nav link)

### Timeline Viewer Screenspace track

Screenspace events rendered as a separate track in both the standalone Timeline Viewer HTML export and Studio-generated viewers.

- [ ] Data contract extension — new optional `screenspaceEvents` key in `window.CLIPGEN_DATA`:
  ```json
  {
    "meta": {"screenspaceEnabled": true},
    "screenspaceEvents": [
      {"id": "ev_...", "type": "change", "eventType": "hud_update", "participant": "P01",
       "timeIn": 50.0, "timeOut": 50.0, "confidence": 0.85, "region": "healthbar", "metadata": {}}
    ]
  }
  ```
- [ ] Second track row ("Screenspace") below the main artifact track
- [ ] Client-side clustering of raw events into visual markers (same threshold logic as Studio)
- [ ] Markers color-coded by detector type (same palette as Screenspace UI)
- [ ] Tooltips showing event type, region, participant, time range, confidence
- [ ] Legend extended with Screenspace detector type swatches
- [ ] Track hidden when no events present
- [ ] `finalize_timeline_data()` in `viewer.py` gets optional `screenspace_events` parameter
- [ ] `load_screenspace_events_for_viewer()` helper reads non-excluded events from screenspace manifest

### Artifact attribution

- [ ] Clips generated from Screenspace intake carry `source: "screenspace"` and `event_ids: [...]` in artifact records
- [ ] `intake_label` field from the event type
- [ ] Visible in Timeline Viewer tooltips and Insights Builder artifact browser

---

## Phase 4: Additional Detection Modes

Expand the analysis engine with new detection capabilities.

- [ ] **Template/Object detection** — find instances of a reference image or template anywhere in the frame, not constrained to a fixed region; useful for detecting specific UI elements, icons, or visual patterns regardless of position
- [ ] **Optical flow** — detect and quantify motion in a region (magnitude, direction); useful for finding moments of high activity vs stillness, tracking movement patterns
- [ ] **Scene type classification** — categorize frames by visual content (menu screen, gameplay, cutscene, loading screen); useful for segmenting long recordings by activity type

---

## Future Considerations

### Multi-region analysis
- Run the same analysis across multiple regions simultaneously (e.g. check three UI elements at once)
- Composite conditions: "healthbar is red AND scoreboard changed"

### Batch participant analysis
- Run a task across all participants automatically (same region, same parameters)
- Cross-participant comparison: "show me when P01 and P03 had the same screen state"

### Video playback
- Full video playback in the Screenspace UI (not just frame-by-frame viewing)
- Playback integration with result timestamps: auto-jump between detected moments

### Smart interval selection
- Adaptive frame sampling: start coarse (every 5s), then refine around detected events (every 0.5s) for precise timestamps
- Keyframe-aware sampling: prefer I-frames for cleaner comparisons

### Template library
- Pre-built analysis templates for common game UI patterns (healthbar, minimap, chat, scoreboard)
- Community-shareable region + parameter presets

---

## Technical Notes

### Dependencies

| Package | Purpose | Size | Install |
|---------|---------|------|---------|
| `opencv-python-headless` | Frame extraction, color analysis, differencing, SSIM | ~40MB | `uv add opencv-python-headless` |
| `easyocr` | OCR for text and numbers detection workflows | ~100MB models (downloaded on first use) | `uv add easyocr` (pulls PyTorch) |
| `imagehash` | Perceptual hashing for fast similarity scans | ~200KB | `uv add imagehash` |
| `scikit-image` | SSIM computation (convenient API) | ~30MB | `uv add scikit-image` |

OpenCV and imagehash are always required. EasyOCR is lazy-imported — only needed when running text or numbers scan tasks. If missing, those workflows show an install instruction rather than crashing.

### Compression artifact handling

All frame comparisons apply these defenses against H.264/H.265 noise:

1. **Gaussian blur** (5x5 kernel) before any pixel comparison — eliminates block artifacts
2. **Pixel diff threshold of 30** — values below this are treated as compression noise
3. **Morphological opening** (3x3 kernel) on diff masks — removes isolated noise pixels
4. **SSIM threshold of 0.90** — below this indicates real visual change; above is noise
5. **Change ratio threshold of 3–5%** — require a minimum percentage of pixels to change before flagging

These defaults are configurable per-task in the UI.

### Performance considerations

- **Frame extraction speed**: OpenCV `VideoCapture` in sequential mode processes ~100–200 frames/sec for region extraction (decode + crop). For a 1-hour video sampled every 1 second, that's 3,600 frames — roughly 20–35 seconds of processing.
- **OCR is the bottleneck**: EasyOCR on CPU runs ~200–500ms per crop. Scanning a 1-hour video every second would take 12–30 minutes. Recommend sampling every 2–5 seconds for text search, with adaptive refinement around hits.
- **Color and change scans are fast**: HSV averaging and pixel differencing are sub-millisecond per frame. The bottleneck is frame decoding, not analysis.
- **Timelapse generation is delegated to ffmpeg**: single-pass crop+speed filter, performance is limited by disk I/O and encode speed rather than Python.

### Assumptions to monitor

- **OpenCV seek accuracy** — `VideoCapture.set(CAP_PROP_POS_MSEC)` can be imprecise on some codecs; may need to fall back to sequential reads with frame counting for exact timestamps
- **EasyOCR on game text** — game fonts (pixel fonts, stylized fonts) may need preprocessing (upscaling, thresholding) before OCR; test with real captures early
- **Region consistency across participants** — assumes game UI is in the same screen position for all participants; may not hold if participants play at different resolutions or with different HUD settings
- **Manifest size** — tasks with dense results (change detection returning hundreds of timestamps) could produce large manifests; may need result summarization or pagination
