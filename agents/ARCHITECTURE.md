# Architecture

| File | Role |
| ------ | ------ |
| [clipgen.py](clipgen.py) | Entry point (`python clipgen.py`), spreadsheet opening helpers, interactive mode dispatch; delegates to pipeline.py for processing |
| [pipeline.py](pipeline.py) | Clip processing pipeline: process_clips, process_reel, compute_reel_id, regenerate_from_manifest, is_excel_worksheet |
| [viewer.py](viewer.py) | Timeline viewer: artifact record building, data finalization, HTML generation with inlined CSS/JS |
| [cli.py](cli.py) | CLI argument parsing, CLI mode detection, setup, Google auth, worksheet selection, CLI mode dispatch, `main()` |
| [spreadsheet.py](spreadsheet.py) | Spreadsheet parsing, header validation, selector parsing (`reel` input), pure timestamp generation for all modes (no prompts) |
| [interactive.py](interactive.py) | Interactive prompt helpers for all modes (line/range/cell/category/participant selection, browse mode); keeps generation functions pure |
| [video.py](video.py) | ffmpeg/ffprobe operations: cut clips, screenshots, GIFs, concatenate reels, optional filesize compression |
| [transcripts.py](transcripts.py) | Transcription via faster-whisper: `transcribe_video()`, segment filtering, write/read transcript files (Markdown/SRT/VTT) |
| [ollama_client.py](ollama_client.py) | Ollama HTTP transport: `is_available()`, `list_models()`, `generate()`, auto-start of `ollama serve`. Pure transport — no prompt or response-parsing logic lives here. |
| [thinking_agents.py](thinking_agents.py) | Registry of Ollama-powered "thinking agents" that reason over transcripts (summary, citations). Owns prompts, model selection, response parsing, and the `AGENTS` list. New agents are added by appending an `Agent` entry — no orchestrator edits needed. |
| [titlecards.py](titlecards.py) | Titlecard/endcard generation: `build_titlecard_frame()`, `build_endcard_frame()`, `wrap_clip_with_cards()` — prepends a title card and appends an endcard in a single FFmpeg encode pass |
| [files.py](files.py) | Filename handling (unique names, truncation), `prepare_clip()` (parse timestamps + annotations, sanitize desc/category), clip discovery for reel-late |
| [utils.py](utils.py) | Timestamp parsing, cell/header annotation parsing, rich/plain output helpers, progress bar utilities, keyword-aware input helpers |
| [config.py](config.py) | Global constants and settings (version, headers, limits, commands) |
| [google_api.py](google_api.py) | Google Sheets auth, worksheet selection by priority, spreadsheet listing/search |
| [excel_io.py](excel_io.py) | Excel adapter: `ExcelSheetAdapter` mimics gspread Worksheet interface for local .xlsx |
| [server.py](server.py) | Combined Flask server for Studio + Insights + Screenspace; registers blueprints per active mode, `start_combined_server()` handles all three on one port |
| [screenspace.py](screenspace.py) | Screenspace analysis engine: image analysis primitives, eleven analysis tools (color, change, similarity, text, numbers, timelapse, template, flow, scene, inactivity + multitool chaining), background task queue worker (`ScreenspaceWorker`), manifest persistence |
| [screenspace_server.py](screenspace_server.py) | Screenspace Flask REST API: region CRUD, video frame extraction, task queue management, results retrieval |
| [insights.py](insights.py) | Insights data model: CRUD operations for insight records, insights manifest read/write |
| [insights_server.py](insights_server.py) | Insights Flask REST API: insight CRUD, artifact browsing, sprite sheet generation, viewer export |
| [data_export.py](data_export.py) | Analysis-ready JSON+CSV export from Screenspace, Insights, and Transcripts manifests; powers `--export` CLI flag and `/screenspace/api/export/events` endpoint |
| [assets/web/](assets/web/) | Static HTML/JS/CSS templates: timeline viewer (`viewer.html/js/css`), gallery (`gallery.html/js/css`), studio (`studio.html/js/css`), insights (`insights-builder.html/js/css`), insights viewer (`insights-viewer.html/js/css`), screenspace (`screenspace.html/js/css`). Shared utilities and constants live in `utils.js` (loaded before each page's main script) |

## Timeline HTML Viewer

Opt-in via `--viewer` or interactive `viewer`. Injects `window.CLIPGEN_DATA` into `assets/web/viewer.html` replacing `<!-- CLIPGEN_DATA_HERE -->`. Data contract, key functions, filmstrip mode, and gallery variant are documented in the [viewer.py](viewer.py) module docstring.

## Gallery HTML Viewer

Opt-in via `--gallery [VIDEO]` or interactive `gv`/`gallery`. Uses the same injection pattern as the timeline viewer but with `assets/web/gallery.html`. Key functions and data contract are documented in the [viewer.py](viewer.py) module docstring.

## Studio Web Interface

Opt-in via `--studio`; requires a spreadsheet. Starts a Flask app via `start_combined_server()` in [server.py](server.py) at `config.SERVER_PORT` (8089), serving `assets/web/studio.html`. API endpoints, module-level state, and key function signatures are documented in the [server.py](server.py) module docstring.

## Insights ([insights.py](insights.py), [insights_server.py](insights_server.py))

Opt-in via `--insights`; no spreadsheet required — reads from `clipgen_manifest.json`. Served at `/insights/` by the same Flask server as Studio. Insight record shape and CRUD function signatures are in the [insights.py](insights.py) module docstring. Flask API endpoints are in the [insights_server.py](insights_server.py) module docstring. The exported `insights_viewer.html` is generated by `finalize_insights_viewer_data()` / `generate_insights_viewer()` in [viewer.py](viewer.py). The frontend source file is `assets/web/insights-builder.{html,js,css}`.

## Screenspace ([screenspace.py](screenspace.py), [screenspace_server.py](screenspace_server.py))

Opt-in via `--screenspace` or interactive `ss`/`screenspace`; no spreadsheet required. Served at `/screenspace/` by the combined Flask server. Analysis tool descriptions (color/change/similarity/text/numbers/timelapse/template/flow/scene/inactivity) and API endpoints are documented in the [screenspace.py](screenspace.py) and [screenspace_server.py](screenspace_server.py) module docstrings.

## Artifact Manifest

Opt-in via `--manifest` or `config.MANIFEST_ENABLED`. Writes `clipgen_manifest.json` alongside artifacts. Key functions (`save_manifest`, `load_manifest_artifacts`, `load_manifest_reels`) and merge/dedup behavior are documented in the [viewer.py](viewer.py) module docstring. Consumed by Insights, `--regenerate`, and standalone `--viewer`.

### Transcription ([transcripts.py](transcripts.py))

Opt-in via `--transcribe` or `config.TRANSCRIBE_ENABLED`. Uses faster-whisper; model is lazy-loaded and cached per session. Data types, key function signatures, and pipeline integration details are documented in the [transcripts.py](transcripts.py) module docstring.

## Local AI: Ollama thinking-agent pipeline

Two layers, kept strictly separate:

1. **Transport** — [ollama_client.py](ollama_client.py) is a thin HTTP wrapper around the Ollama REST API. Every LLM call in the project routes through `ollama_client.generate()`. It knows nothing about prompts, transcripts, or parsing. On connection-refused it auto-starts `ollama serve` and retries once.
2. **Agents** — [thinking_agents.py](thinking_agents.py) defines the `Agent` shape (prompt building + model selection + response parsing + `depends_on` metadata + target `manifest_field`) and exports an ordered `AGENTS` list. The two built-in agents are `summary` (Pass 1: paragraph + bullets) and `citations` (Pass 2: supporting segments per claim, depends on `summary`).

**Orchestration** lives in [transcripts_server.py](transcripts_server.py) as three small helpers — `_next_eligible_agent()`, `_run_agent()`, `_run_agent_chain()`. When a transcript completes (or a user hits Regenerate), the chain is re-entered and the first eligible agent is spawned in a daemon thread. An agent is eligible when it is enabled in config, its `manifest_field` is empty on the transcript entry, its `depends_on` agents' fields are populated, and it is not already running for that participant. After each agent finishes, the chain advances.

Results are written into `source_transcripts[participant][manifest_field]` in `transcripts_manifest.json`. In-flight tracking is keyed per-agent via `_agent_in_flight[agent_key]`, surfaced to the frontend through generic `_is_generating(participant, agent_key)` checks at the API endpoints.

**Adding a new agent** (e.g. keyword extractor, pain-point tagger):

1. Add an `OLLAMA_<NAME>_ENABLED` toggle (and any `_MODEL` keys) to [config.py](config.py).
2. Write a `run` callable in [thinking_agents.py](thinking_agents.py) that takes the transcript entry dict and returns the value to store (or `None` to skip).
3. Append an `Agent` entry to `AGENTS` (respect topological order — dependencies before dependents).

No edits to [transcripts_server.py](transcripts_server.py) or [ollama_client.py](ollama_client.py) are needed — the orchestrator picks up the new agent automatically. If the UI should surface the new result, add one endpoint or extend the existing response shape.

## Titlecards ([titlecards.py](titlecards.py))

Opt-in via `config.TITLECARDS_ENABLED` or `--titlecards` / `--no-titlecards`. Prepends a title card (first source frame + text overlay) to each clip. Key functions and config knobs are documented in the [titlecards.py](titlecards.py) module docstring.
