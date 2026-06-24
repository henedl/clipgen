# Architecture

| File | Role |
| ------ | ------ |
| [clipgen.py](clipgen.py) | Entry point (`uv run clipgen.py`), spreadsheet opening helpers, interactive mode dispatch; delegates to pipeline.py for processing |
| [pipeline.py](pipeline.py) | Clip processing pipeline: process_clips, process_reel, compute_reel_id, regenerate_from_manifest, is_excel_worksheet |
| [viewer.py](viewer.py) | Timeline viewer: artifact record building, data finalization, HTML generation with inlined CSS/JS |
| [cli.py](cli.py) | CLI argument parsing, CLI mode detection, setup, Google auth, worksheet selection, CLI mode dispatch, `main()` |
| [spreadsheet.py](spreadsheet.py) | Spreadsheet parsing, header validation, selector parsing (`reel` input), pure timestamp generation for all modes (no prompts) |
| [interactive.py](interactive.py) | Interactive prompt helpers for all modes (line/range/cell/category/participant selection, browse mode); keeps generation functions pure |
| [video.py](video.py) | ffmpeg/ffprobe operations: cut clips, screenshots, GIFs, concatenate reels, optional filesize compression |
| [transcripts.py](transcripts.py) | Transcription via faster-whisper: `transcribe_video()`, segment filtering, write/read transcript files (Markdown/SRT/VTT) |
| [transcripts_server.py](transcripts_server.py) | Transcripts Flask blueprint and thinking-agent orchestrator: `_next_eligible_agent()`, `_run_agent()`, `_run_agent_chain()`, `_agent_in_flight` tracking, REST endpoints |
| [ollama_client.py](ollama_client.py) | Ollama HTTP transport: `is_available()`, `list_models()`, `generate()`, auto-start of `ollama serve`. Pure transport — no prompt or response-parsing logic lives here. |
| [thinking_agents.py](thinking_agents.py) | Registry of Ollama-powered "thinking agents" that reason over transcripts (summary, citations, friction). Owns prompts, model selection, response parsing, and the `AGENTS` list. New agents are added by appending an `Agent` entry — no orchestrator edits needed. |
| [friction.py](friction.py) | Pure deterministic friction scorer: compiled phrase patterns across six categories, per-segment score/markers, candidate selection, session stats, rolling-window smoothing. No Ollama/I/O — the LLM refinement (`find_friction_moments`, `_run_friction`) lives in thinking_agents.py. |
| [titlecards.py](titlecards.py) | Titlecard/endcard generation: `build_titlecard_frame()`, `build_endcard_frame()`, `wrap_clip_with_cards()` — prepends a title card and appends an endcard in a single FFmpeg encode pass |
| [files.py](files.py) | Filename handling (unique names, truncation), `prepare_clip()` (parse timestamps + annotations, sanitize desc/category), clip discovery for reel-late |
| [utils.py](utils.py) | Timestamp parsing, cell/header annotation parsing, rich/plain output helpers, progress bar utilities, keyword-aware input helpers |
| [config.py](config.py) | Global constants and settings (version, headers, limits, commands) |
| [google_api.py](google_api.py) | Google Sheets auth, worksheet selection by priority, spreadsheet listing/search |
| [excel_io.py](excel_io.py) | Excel adapter: `ExcelSheetAdapter` mimics gspread Worksheet interface for local .xlsx |
| [server.py](server.py) | Combined Flask server for Studio + Screenspace + Transcripts; registers blueprints per active mode, `start_combined_server()` handles all three on one port |
| [screenspace.py](screenspace.py) | Screenspace re-export facade: a thin module that re-exports every public name (and the test-touched private names) from the nine `screenspace_*` modules below, so `import screenspace; screenspace.NAME` keeps working unchanged. Holds the engine's 11-tool overview docstring; no implementation |
| [screenspace_primitives.py](screenspace_primitives.py) | Pure cv2/numpy/PIL image-analysis primitives: region crop/denormalize/resolve, HSV color math, frame-diff, SSIM, perceptual hash, template matching, optical flow, scene fingerprinting, plus scan-support helpers (`_morph_kernel`, `_ConsecutiveBuffer`, `_is_static_skip`, `_merge_timestamp_spans`). No I/O or ffmpeg |
| [screenspace_ocr.py](screenspace_ocr.py) | Cached EasyOCR readers, region preprocessing, glyph confusion-folding, numeric comparison helpers/allowlists, region reading/scoring, and `run_calibration_ocr`. Imports primitives |
| [screenspace_frames.py](screenspace_frames.py) | ffmpeg-pipe frame extraction (`scan_video_frames`/`scan_video_full_frames`, `_ffmpeg_pipe_frames`), timelapse command builder, ffprobe metadata. Imports primitives |
| [screenspace_scans.py](screenspace_scans.py) | The eleven per-tool scan workflows (`scan_color`…`scan_boundaries`, `generate_timelapse`); scans never call each other. Imports primitives, ocr, frames |
| [screenspace_heatmap.py](screenspace_heatmap.py) | Template/flow/change heatmap PNG + animated-GIF generation (cumulative + rolling-window). Pure cv2/PIL leaf, no sibling-module deps |
| [screenspace_tools.py](screenspace_tools.py) | `AnalysisTool` base + twelve tool subclasses, the `TOOLS` registry, and per-frame dispatch (`check_frame_for_tool`, `score_frame_for_tool`). Imports primitives, ocr, scans; one function-local import of `scan_multitool` in `MultitoolTool.scan` breaks the tools↔multitool cycle |
| [screenspace_multitool.py](screenspace_multitool.py) | Multitool chaining + time-offset window joining (`scan_multitool`, `score_multitool_frame`). Imports tools + frames |
| [screenspace_manifest.py](screenspace_manifest.py) | Task-status constants, `create_task`, manifest load/save, result-time offsetting, and event generation. Imports tools (`_extract_confidence`) |
| [screenspace_worker.py](screenspace_worker.py) | `ScreenspaceWorker` background task queue (threading, pause/resume/cancel, multi-video, heatmaps, events). Imports tools, heatmap, manifest, multitool, frames |
| [screenspace_server.py](screenspace_server.py) | Screenspace Flask REST API: region CRUD, video frame extraction, task queue management, results retrieval |
| [data_export.py](data_export.py) | Analysis-ready JSON+CSV export from Screenspace and Transcripts manifests; powers `--export` CLI flag and `/screenspace/api/export/events` endpoint |
| [assets/web/](assets/web/) | Static HTML/JS/CSS templates: timeline viewer (`viewer.html/js/css`), gallery (`gallery.html/js/css`), studio (`studio.html/js/css`), screenspace (`screenspace.html/js/css`), transcripts (`transcripts.html/js/css`). Shared utilities and constants live in `utils.js` (loaded before each page's main script). The Screenspace page is split across a hub (`screenspace.js`) plus state-free leaf helpers (`screenspace-utils.js`) and five satellite files — `screenspace-tasks.js` (task queue + right-pane tabs + SSE/polling + ETA), `screenspace-results.js` (results panel: confidence histogram, rows, heatmap overlay), `screenspace-multitool-params.js`, `screenspace-calibration.js`, `screenspace-color.js` — that share the hub's state and helpers through a `window.ClipgenScreenspace` namespace — the same hub-plus-satellite pattern studio.js uses with `convergence.js`/`metadata.js`. The Transcripts page is split the same way: a `transcripts.js` hub plus `transcripts-{corrections,search,video,pills,agents}.js` satellites sharing state through `window.ClipgenTranscripts` (TS) — the hub keeps `state`, the cross-reference index, `selectParticipant`, the segments+marks editor, the transcription/agent task poller, and model-install, reaching satellite functions through same-named guarded delegators. The Studio page is mid-carve onto the same axis: a `studio.js` hub plus a `studio-intake.js` satellite (Screenspace + Transcript intake panels) sharing state through `window.ClipgenStudio` (STUDIO) with same-named guarded delegators — separate from the older `window._studio*` legacy globals that `convergence.js`/`metadata.js` still lazy-read on tab activation. `card-scrubber.{js,css}` provides reusable hover-to-scrub: sprite-sheet frame scrubbing + synced audio snippets + a waveform/playhead overlay (exposes `window.clipgenCardScrubber`). Opt-in on two surfaces: Studio (settings toggle `STUDIO_CARD_SCRUBBER`, backed by the `/api/sprite` + `/api/clip-audio` routes) loads it directly; the timeline viewer (header toggle, localStorage `clipgen-viewer-scrubaudio`) layers its audio/waveform primitives onto the existing `<video>`-seek scrub and inlines the module via `viewer.py`. |

> Subsystem internals (data contracts, key function signatures, merge/dedup rules) live in each module's docstring. The sections below describe how each subsystem is launched and how it fits into the overall flow; for internals, open the referenced module.

## Timeline HTML Viewer ([viewer.py](viewer.py))

Opt-in via `--viewer` or interactive `viewer`. Injects `window.CLIPGEN_DATA` into `assets/web/viewer.html` replacing `<!-- CLIPGEN_DATA_HERE -->`.

## Gallery HTML Viewer ([viewer.py](viewer.py))

Opt-in via `--gallery [VIDEO]` or interactive `gv`/`gallery`. Uses the same injection pattern as the timeline viewer but with `assets/web/gallery.html`.

## Studio Web Interface ([server.py](server.py))

Opt-in via `--studio`; requires a spreadsheet. Starts a Flask app via `start_combined_server()` at `config.SERVER_PORT` (8089), serving `assets/web/studio.html`.

## Screenspace ([screenspace.py](screenspace.py) facade over `screenspace_primitives`/`_ocr`/`_frames`/`_scans`/`_heatmap`/`_tools`/`_multitool`/`_manifest`/`_worker`, [screenspace_server.py](screenspace_server.py))

Opt-in via `--screenspace` or interactive `ss`/`screenspace`; no spreadsheet required. Served at `/screenspace/` by the combined Flask server. The engine ships eleven scan workflows and twelve `AnalysisTool` classes (the ten region detectors plus `timelapse` and `multitool`); `screenspace.py` is a thin re-export facade and the implementation lives in the cohesive `screenspace_*` sibling modules listed above, wired in an acyclic deepest-first DAG (`primitives`/`heatmap` → `ocr` → `frames` → `scans` → `tools` → `multitool`/`manifest` → `worker`). The browser UI is split the same way: a `screenspace.js` hub plus `screenspace-{multitool-params,calibration,color}.js` satellites sharing state through `window.ClipgenScreenspace` (see the `assets/web/` row).

## Artifact Manifest ([viewer.py](viewer.py))

Opt-in via `--manifest` or `config.MANIFEST_ENABLED`. Writes `clipgen_manifest.json` alongside artifacts. Key functions: `save_manifest`, `load_manifest_artifacts`, `load_manifest_reels`. Consumed by `--regenerate` and standalone `--viewer`.

## Transcription ([transcripts.py](transcripts.py))

Opt-in via `--transcribe` or `config.TRANSCRIBE_ENABLED`. Uses faster-whisper; model is lazy-loaded and cached per session.

## Local AI: Ollama thinking-agent pipeline ([ollama_client.py](ollama_client.py), [thinking_agents.py](thinking_agents.py), [transcripts_server.py](transcripts_server.py))

Two layers, kept strictly separate:

1. **Transport** — [ollama_client.py](ollama_client.py) is a thin HTTP wrapper around the Ollama REST API. Every LLM call in the project routes through `ollama_client.generate()`. It knows nothing about prompts, transcripts, or parsing. On connection-refused it auto-starts `ollama serve` and retries once.
2. **Agents** — [thinking_agents.py](thinking_agents.py) defines the `Agent` shape (prompt building + model selection + response parsing + `depends_on` metadata + target `manifest_field`) and exports an ordered `AGENTS` list. The built-in agents are `summary` (Pass 1: paragraph + bullets), `citations` (Pass 2: supporting segments per claim, depends on `summary`), and `friction` (Pass 3: deterministic scorer in [friction.py](friction.py) + LLM-refined moments, depends on `summary` — runs even when citations is disabled).

**Orchestration** lives in [transcripts_server.py](transcripts_server.py) as three small helpers — `_next_eligible_agent()`, `_run_agent()`, `_run_agent_chain()`. When a transcript completes (or a user hits Regenerate), the chain is re-entered and the first eligible agent is spawned in a daemon thread. An agent is eligible when it is enabled in config, its `manifest_field` is empty on the transcript entry, its `depends_on` agents' fields are populated, and it is not already running for that participant. After each agent finishes, the chain advances.

Results are written into `source_transcripts[participant][manifest_field]` in `transcripts_manifest.json`. In-flight tracking is keyed per-agent via `_agent_in_flight[agent_key]`, surfaced to the frontend through generic `_is_generating(participant, agent_key)` checks at the API endpoints.

**Adding a new agent**: see [agents/skills/new-thinking-agent/SKILL.md](skills/new-thinking-agent/SKILL.md). The orchestrator picks up new agents from the `AGENTS` list automatically — no edits to [transcripts_server.py](transcripts_server.py) or [ollama_client.py](ollama_client.py) are needed.

## Titlecards ([titlecards.py](titlecards.py))

Opt-in via `config.TITLECARDS_ENABLED` or `--titlecards` / `--no-titlecards`. Prepends a title card (first source frame + text overlay) to each clip.
