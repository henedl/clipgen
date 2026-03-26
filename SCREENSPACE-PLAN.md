# Screenspace Plan

## Overview

Screenspace is a post-processing feature set that performs image analysis on source video files to find and label moments matching specific visual conditions. It produces a duct-taped form of telemetry for software that lacks traditional telemetry — searching for visual changes, color states, text appearances, and similarity matches across long video recordings.

Results are displayed on a timeline in the Screenspace UI, and can be exported as timestamp markers for use in other clipgen viewers (e.g. Timeline Viewer).

### Use cases driving this work

1. **State monitoring** — detect when a UI element enters a specific visual state (e.g. a health bar turning red, a meter emptying)
2. **Change tracking** — find every moment a screen region changes (e.g. a scoreboard updating, a notification appearing)
3. **Visual search** — find frames that look similar to a reference moment (e.g. "show me every time the inventory screen looked like this")
4. **Text detection** — find when specific text appears on screen (e.g. a chat message, an error dialog, a score value)
5. **Region timelapse** — generate a sped-up view of a screen region over time (e.g. a minimap timelapse, a resource counter progression)

### Key architectural decisions

- **Combined server model.** Screenspace runs within the existing combined Flask server (alongside Studio and Insights), registered as its own blueprint at `/screenspace/`. A background worker thread handles compute-intensive analysis tasks.
- **Per-participant source videos.** Screenspace operates on `{study}_{participant}.mp4` files, consistent with clipgen's participant model. The user selects a participant from a dropdown.
- **Named region definitions.** Rectangular screen regions can be named, saved, and reused across participants and studies. Game UIs are typically positionally consistent across sessions, so a region like "healthbar" defined once applies to all participants.
- **Task queue with single worker.** Analysis tasks are queued and processed sequentially by a background worker thread. The queue is visible in the frontend with reorder and cancel controls.
- **Separate manifest.** Screenspace persists tasks, region definitions, and results to `screenspace_manifest.json`, independent of `clipgen_manifest.json`.
- **OpenCV as the analysis engine.** `opencv-python-headless` (~40MB, pip-installable, no system dependencies) handles color analysis, frame differencing, similarity comparison, and frame extraction. ffmpeg handles timelapse generation via crop filters.
- **EasyOCR for text detection.** Lazy-imported so users who don't use OCR features never need PyTorch installed. Models downloaded on first use (~100MB for English).
- **Compression-aware thresholds.** All comparisons apply Gaussian blur and noise thresholds to handle H.264/H.265 compression artifacts. Default pixel-diff threshold of 30, SSIM threshold of 0.90.
- **Screenspace UI has its own timeline.** Results are displayed on a built-in timeline. Export to Timeline Viewer markers is a future integration point.

---

## Phase 1: Infrastructure

Build the analysis engine, server integration, and task queue.

### Analysis module (`screenspace.py`)

Core image analysis functions, all operating on NumPy arrays (OpenCV frames):

- [ ] `extract_region(frame, region)` — crop a rectangular region from a frame; region is `{x, y, w, h}` in pixels
- [ ] `average_color_hsv(region_pixels)` — compute mean HSV color of a cropped region; return `{h, s, v}` values
- [ ] `color_matches(region_pixels, target_color, tolerance)` — check if region's average color is within tolerance of a target; comparison in HSV space (hue-aware, handles red wraparound)
- [ ] `compute_frame_diff(region_a, region_b, noise_threshold=30)` — absolute pixel difference between two same-sized regions; apply Gaussian blur first; return change ratio (0.0–1.0)
- [ ] `regions_are_similar(region_a, region_b, threshold=0.90)` — SSIM-based similarity check with blur preprocessing; return boolean + score
- [ ] `compute_phash(region_pixels)` — perceptual hash of a region for fast similarity scanning; return hash object
- [ ] `scan_video_frames(video_path, region, interval_seconds, callback)` — iterate through video at interval, extract region, call callback per frame; yield `(timestamp, frame_data)` pairs; use OpenCV `VideoCapture` in sequential mode
- [ ] `build_timelapse_command(video_path, region, speedup_factor, output_path, output_format)` — construct ffmpeg argv for cropped timelapse (video or GIF); use `crop=w:h:x:y` + `setpts` filters

### Analysis workflows

Each workflow wraps the primitives above into a complete scan-and-report operation:

- [ ] **Color scan** (`scan_color`): iterate frames at interval, check `color_matches()` on region, collect timestamps where condition is met; return list of `{start, end, duration}` spans (merge consecutive matches into spans)
- [ ] **Change scan** (`scan_changes`): compare region across consecutive sampled frames, flag timestamps where `compute_frame_diff()` exceeds threshold; return list of change-point timestamps with magnitude
- [ ] **Similarity scan** (`scan_similarity`): given a reference frame's region, scan video for frames where `regions_are_similar()` or phash distance is below threshold; return ranked list of `{timestamp, score}` matches
- [ ] **Text scan** (`scan_text`): iterate frames at interval, run EasyOCR on region, match extracted text against search string using fuzzy matching (`difflib.SequenceMatcher`, ratio > 0.8); return list of `{timestamp, text_found, confidence}` matches; EasyOCR lazy-imported with clear error if missing
- [ ] **Timelapse generation** (`generate_timelapse`): run ffmpeg crop+speed command via `video.run_ffmpeg_process()`; return output file path

### Task queue

- [ ] `ScreenspaceWorker` — background thread that pulls tasks from a `queue.PriorityQueue` and executes them; single worker by default
- [ ] Task lifecycle: `queued` → `running` → `completed` / `failed` / `cancelled`
- [ ] Each task is a dict: `{id, type, participant, video_path, region, parameters, status, progress, result, created_at}`
- [ ] Progress reporting: worker updates task `progress` (0.0–1.0) as frames are processed; frontend polls for updates
- [ ] Cancellation: set a `cancelled` flag on the task; worker checks flag between frame iterations and aborts early
- [ ] Queue reordering: tasks can be reordered by adjusting priority values while status is `queued`

### Manifest (`screenspace_manifest.json`)

- [ ] Schema:
  ```json
  {
    "regions": {
      "healthbar": {"x": 100, "y": 20, "w": 300, "h": 30, "description": "Player health bar"},
      "scoreboard": {"x": 500, "y": 0, "w": 200, "h": 50, "description": "Score display"}
    },
    "tasks": [
      {
        "id": "ss_<8hex>",
        "type": "color|change|similarity|text|timelapse",
        "participant": "P01",
        "source_video": "study_P01.mp4",
        "region": "healthbar",
        "parameters": {},
        "status": "completed",
        "created_at": "iso8601",
        "completed_at": "iso8601",
        "results": []
      }
    ]
  }
  ```
- [ ] Regions are top-level and shared across tasks (referenced by name)
- [ ] `load_screenspace_manifest()` / `save_screenspace_manifest()` — read/write with merge semantics (match existing manifest patterns)
- [ ] Manifest lives in the output directory alongside `clipgen_manifest.json`

### Server integration

- [ ] Create `screenspace_server.py` with `screenspace_bp = Blueprint("screenspace", __name__)`
- [ ] `_init_screenspace_state()` — initialize worker thread, load manifest, resolve participant video paths
- [ ] Register blueprint in `start_combined_server()` at `/screenspace/` prefix
- [ ] Add `--screenspace` CLI flag in `cli.py`; can combine with `-s`, `-i/-o`, `-v` (requires spreadsheet for participant resolution); cannot combine with mode flags
- [ ] Add `SCREENSPACE_MANIFEST_FILENAME` to `config.py`
- [ ] Static file serving: serve `screenspace.html` at `/screenspace/`
- [ ] API endpoints:
  - `GET /api/participants` — list participants with source video availability
  - `GET /api/video/frame/<participant>/<timestamp>` — extract and return a single frame as JPEG at given timestamp
  - `GET /api/video/info/<participant>` — return video metadata (duration, resolution, fps)
  - `GET /api/regions` — list saved region definitions
  - `POST /api/regions` — create/update a named region
  - `DELETE /api/regions/<name>` — delete a region definition
  - `POST /api/tasks` — enqueue a new analysis task
  - `GET /api/tasks` — list all tasks with status and progress
  - `GET /api/tasks/<id>` — get task detail including results
  - `DELETE /api/tasks/<id>` — cancel a queued/running task
  - `PUT /api/tasks/reorder` — reorder queued tasks
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

- [x] Workflow selector: toolbar below the video with buttons/tabs for each workflow type (Color, Change, Similarity, Text, Timelapse)
- [x] Per-workflow parameter panels:
  - **Color**: target color picker (HSV), tolerance slider, sample interval
  - **Change**: sensitivity threshold slider, sample interval
  - **Similarity**: reference frame selector (current frame becomes reference), similarity threshold, sample interval
  - **Text**: search string input, fuzzy match threshold, sample interval, language selector
  - **Timelapse**: speedup factor, output format (video/GIF), output resolution
- [x] Time range selection: optional in/out markers on the timeline to restrict analysis to a sub-range; default is full video duration
- [x] "Run" button to enqueue the configured task

### Timeline

- [x] Horizontal timeline bar showing the full duration of the selected participant's video
- [x] Result markers: completed tasks' result timestamps shown as markers/spans on the timeline; color-coded by task type
- [x] In/out markers: draggable markers to set analysis time range
- [x] Click-to-seek: clicking a point on the timeline updates the video frame viewer to that timestamp
- [x] Zoom and pan: ability to zoom into a time range for dense results

### Task queue panel

- [x] Queue display: list of all tasks (queued, running, completed, failed) with status indicators
- [x] Progress bar for the currently running task
- [ ] Spinner/indeterminate progress when frame count is unknown
- [ ] Drag-to-reorder queued tasks
- [x] Cancel button for queued and running tasks
- [x] Click a completed task to load its results onto the timeline
- [x] Task detail view: show parameters, region used, result count, timestamps

### Results display

- [x] Results list: scrollable list of timestamps/spans returned by a completed task
- [x] Click a result to seek the video frame viewer to that timestamp
- [x] For color/change/similarity tasks: show the match score or change magnitude alongside each timestamp
- [x] For text tasks: show the OCR'd text and confidence alongside each timestamp
- [x] For timelapse tasks: inline video/GIF player showing the generated timelapse
- [x] Export results: button to download results as JSON or CSV

---

## Phase 3: Integration and Export

Connect Screenspace results with the rest of clipgen.

### Timeline Viewer integration

- [ ] Export Screenspace result timestamps as a marker format compatible with Timeline Viewer
- [ ] Timeline Viewer: render Screenspace markers as a separate track/layer alongside clip artifacts
- [ ] Marker metadata: task type, region name, match score, linked to the originating Screenspace task

### Studio integration

- [ ] Surface Screenspace markers in Studio's spreadsheet grid (e.g. highlight cells whose timestamps overlap with Screenspace results)
- [ ] Quick link from Studio to open Screenspace for the selected participant

### Artifact generation from results

- [ ] Allow generating clips/screenshots from Screenspace result timestamps (reuse existing `video.py` pipeline)
- [ ] Generated artifacts written to `clipgen_manifest.json` alongside manually selected artifacts
- [ ] Attribution: artifacts generated from Screenspace results carry a `source: "screenspace"` field and link back to the task ID

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
| `easyocr` | OCR for text detection workflow | ~100MB models (downloaded on first use) | `uv add easyocr` (pulls PyTorch) |
| `imagehash` | Perceptual hashing for fast similarity scans | ~200KB | `uv add imagehash` |
| `scikit-image` | SSIM computation (convenient API) | ~30MB | `uv add scikit-image` |

OpenCV and imagehash are always required. EasyOCR is lazy-imported — only needed when running text scan tasks. If missing, the text workflow shows an install instruction rather than crashing.

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
