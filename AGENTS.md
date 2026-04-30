# clipgen – Project context for AI assistants

This document contains stable project facts — architecture, data structures, genuine gotchas, conventions (the things that don't change run-to-run).

## Other files

@CLOUD.md contains instructions specifically for cloud agents.

## Project overview

clipgen is a tool for user researchers that:

1. Generates clips from timestamps stored in a Google Sheet or a local Excel file. It uses **gspread** for Google Sheets access, **openpyxl** for Excel, and **ffmpeg/ffprobe** for media processing.
2. Allows video and audio analysis of local video files, in the **Screenspace** and **Transcript** tools.

**Data flow:** Timestamps in spreadsheet → clipgen reads records (description, study, participant ID, category) → timestamp parsing/annotation filtering → ffmpeg → video clips, screenshots, GIFs, or a single reel. Optionally, generated artifacts can be transcribed via faster-whisper to produce timestamped transcript files.

## Architecture

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

### Timeline HTML Viewer

Opt-in via `--viewer` or interactive `viewer`. Injects `window.CLIPGEN_DATA` into `assets/web/viewer.html` replacing `<!-- CLIPGEN_DATA_HERE -->`. Data contract, key functions, filmstrip mode, and gallery variant are documented in the [viewer.py](viewer.py) module docstring.

### Gallery HTML Viewer

Opt-in via `--gallery [VIDEO]` or interactive `gv`/`gallery`. Uses the same injection pattern as the timeline viewer but with `assets/web/gallery.html`. Key functions and data contract are documented in the [viewer.py](viewer.py) module docstring.

### Studio Web Interface

Opt-in via `--studio`; requires a spreadsheet. Starts a Flask app via `start_combined_server()` in [server.py](server.py) at `config.SERVER_PORT` (8089), serving `assets/web/studio.html`. API endpoints, module-level state, and key function signatures are documented in the [server.py](server.py) module docstring.

### Insights ([insights.py](insights.py), [insights_server.py](insights_server.py))

Opt-in via `--insights`; no spreadsheet required — reads from `clipgen_manifest.json`. Served at `/insights/` by the same Flask server as Studio. Insight record shape and CRUD function signatures are in the [insights.py](insights.py) module docstring. Flask API endpoints are in the [insights_server.py](insights_server.py) module docstring. The exported `insights_viewer.html` is generated by `finalize_insights_viewer_data()` / `generate_insights_viewer()` in [viewer.py](viewer.py). The frontend source file is `assets/web/insights-builder.{html,js,css}`.

### Screenspace ([screenspace.py](screenspace.py), [screenspace_server.py](screenspace_server.py))

Opt-in via `--screenspace` or interactive `ss`/`screenspace`; no spreadsheet required. Served at `/screenspace/` by the combined Flask server. Analysis tool descriptions (color/change/similarity/text/numbers/timelapse/template/flow/scene/inactivity) and API endpoints are documented in the [screenspace.py](screenspace.py) and [screenspace_server.py](screenspace_server.py) module docstrings.

### Artifact Manifest

Opt-in via `--manifest` or `config.MANIFEST_ENABLED`. Writes `clipgen_manifest.json` alongside artifacts. Key functions (`save_manifest`, `load_manifest_artifacts`, `load_manifest_reels`) and merge/dedup behavior are documented in the [viewer.py](viewer.py) module docstring. Consumed by Insights, `--regenerate`, and standalone `--viewer`.

### Transcription ([transcripts.py](transcripts.py))

Opt-in via `--transcribe` or `config.TRANSCRIBE_ENABLED`. Uses faster-whisper; model is lazy-loaded and cached per session. Data types, key function signatures, and pipeline integration details are documented in the [transcripts.py](transcripts.py) module docstring.

### Local AI: Ollama thinking-agent pipeline

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

### Titlecards ([titlecards.py](titlecards.py))

Opt-in via `config.TITLECARDS_ENABLED` or `--titlecards` / `--no-titlecards`. Prepends a title card (first source frame + text overlay) to each clip. Key functions and config knobs are documented in the [titlecards.py](titlecards.py) module docstring.

## Key data structures

**Clip record** (built in spreadsheet layer, enriched in files):

```python
{
    'cell': gspread.Cell,      # Cell with timestamp value (1-based row/col)
    'desc': str,               # Observation text from Observation column
    'study': str,              # Normalized study name (filesystem-safe)
    'participant': str,        # e.g. 'P01', 'G02' from header
    'category': str,           # Row category (sanitized; empty → 'uncategorized')
    'severity': str,           # Row severity (normalized label; empty if no Severity column)
    'times': [(start, end)]    # Added by files.prepare_clip() – list of (start_time, end_time) strings
}
```

Source video filenames follow `{study}_{participant}.mp4` (e.g. `mystudy_P01.mp4`).

## Development tools

- **uv** – Use `uv run` instead of `python` to run scripts (e.g. `uv run clipgen.py`). Use `uv add` to add dependencies.
- **Ruff** – Linting and formatting. A `PostToolUse` hook in `.claude/settings.json` automatically runs `uv run ruff check --fix` and `uv run ruff format` on every edited/written file. You can also run these manually: `uv run ruff check --fix` and `uv run ruff format`.
- **ty** – Use `uv run ty check` for type checking.

## SVG icons

316 Heroicons (outline, 24×24) live in [assets/icons/](assets/icons/) with kebab-case filenames. Use these for all web UI icons rather than writing new inline SVG paths. The canonical pattern is CSS `mask-image` referencing `.svg` files — see `XREF_BADGES` in `utils.js` and `.xref-badge-icon` in `tokens.css`. Icon routes already exist per blueprint (`/screenspace/icons/`, `/transcripts/icons/`, etc.).

## Conventions and patterns

- **Shared Python/JS config:** values mirrored on the frontend (severity labels, `DEFAULT_DURATION_SECONDS`, `ANNOTATION_KEYPHRASES`, `IGNORED_TIMESTAMP_TOKENS`) flow through `utils.get_frontend_config()`. Server routes (`server.py`, `insights_server.py`) embed it as `"config":` in their JSON; `viewer.py` finalize_* functions embed it in `window.CLIPGEN_DATA`. JS overlays the payload onto `CLIPGEN_CONFIG` in `assets/web/utils.js` via `clipgenApplyConfig()`; the `CLIPGEN_CONFIG` defaults are the offline-fallback for re-opened exported viewers. `tests/test_shared_constants.py` asserts the JS defaults match Python.
- **Coordinates:** gspread uses **1-based** row/col. `sheet.get_all_values()` is a list of lists with **0-based** indices: `sheet_data[row_idx][col_idx]`. Conversions: sheet row = `row_idx + 1`, sheet col = `col_idx + 1`.
- **Timestamps:** Formats `MM:SS` or `HH:MM:SS`. Ranges with `-` (e.g. `1:23-1:45`). Multiple pairs separated by `,`, `;`, `+`, or space. Single time gets end = start + `DEFAULT_DURATION_SECONDS`.
- **Annotations:** `utils.parse_cell_annotations()` strips supported keyphrases (configured in `ANNOTATION_KEYPHRASES`, currently `!key`) before timestamp parsing. Ignored tokens (configured in `IGNORED_TIMESTAMP_TOKENS`, currently `x`) are skipped.
- **Participant IDs:** Headers must start with `P` (individual) or `G` (group); see `config.PARTICIPANT_PREFIXES`.
- **User feedback:** Use `utils.error_print()`, `utils.warning_print()`, `utils.verbose_print()`, `utils.info_print()`. Prefer these over direct `print()` for user-facing messages.
- **Debug:** Set `config.DEBUGGING = True` to enable icecream output, skip ffmpeg execution paths in [video.py](video.py), and return stub transcript results in [transcripts.py](transcripts.py) without loading a Whisper model.
- **Interactive keywords:** All interactive prompts go through `utils.read_user_input()`, which treats first-token commands as:
  - `quit` / `exit` → exit clipgen
  - `top` → return to spreadsheet selection
  - `back` → return to mode selection (or spreadsheet selection if already at mode selection)

## Modes

For the full mode and flag reference, run `uv run clipgen.py --help`. What follows are only gotchas not visible there.

`-H/--highlights` generates a highlights reel scored by severity, uniqueness, and keyword annotations within a configurable time budget (default 180s); optionally pass a duration in seconds (e.g. `-H 120`).

`--studio`, `--screenspace`, and `--insights` all launches a Flask-based web interface for interactive artifact generation. These run off the same Flask, but their flags are mutually exclusive.

## Spreadsheet layout

Required columns: **ID**, **Observation**, **Category**. Participant columns follow ID with headers starting `P` or `G`. An optional `Baseline time` marker row enables clock/absolute timestamps per participant column (converted to relative offsets in `files.prepare_clip()`). Full layout details and baseline semantics are documented in the [spreadsheet.py](spreadsheet.py) module docstring and in [README.md](README.md).

## Version

- The version is stored as `VERSIONNUM` in [config.py](config.py).
- **When making substantive code changes** (bug fixes or features), increment the **last segment only** (patch) in `config.py`, e.g. `0.9.0` → `0.9.1`. Do not bump for docs-only, comment-only, or refactor-only changes unless they affect user-visible behavior.

## Testing notes

- Run the test suite from the project root with `uv run --extra dev pytest -c tests/pytest.ini`. The `dev` optional extra (tests/CI only) supplies pytest; default `uv sync` does not install it.
- Tests cover: CLI argument parsing, CLI mode dispatch, clip pipeline, file/artifact handling, Google/Excel adapters, insights data model, insights API, manifest operations, selectors, spreadsheet generation, studio API, titlecards, transcripts, timestamp utilities, video commands, viewer data, and viewer inlining.
- Every new CLI mode, flag, or selector should include at least one smoke test in the same PR.
- With `config.DEBUGGING = True`, icecream is enabled, [video.py](video.py) does not invoke ffmpeg, and [transcripts.py](transcripts.py) returns stub results without loading a Whisper model.

## Project learnings for agents

### Learned User Preferences

- When the user attaches an implementation plan that already has created todos, do not edit the plan file; mark those todos in_progress as you work and complete them without recreating the list.
- Don't extract helpers unless it is called more than once.
- Prefer placing generic index/letter conversion utilities (e.g. index_to_letter, letter_to_index) in utils.py rather than in domain-specific modules like files.py.
- Prefer naming new helpers to match existing method naming patterns in the same module.
- Never write a class when a function will do.
- Treat spreadsheet layout and timestamp semantics as domain rules; if tests conflict with these, reconsider or adjust the tests rather than changing core semantics to satisfy them.
- All web UIs use vanilla JavaScript (ES5-style `.then()` chaining, not async/await), hand-written CSS with CSS variables for theming, and plain HTML. No React, TypeScript, CSS frameworks, or build tools.
- **CSS design tokens**: `assets/web/tokens.css` is the single source of truth for shared design values. Never redefine these in page CSS; page-specific variables (e.g. `--color-cell-text` in studio, `--color-overlay-bg` in gallery) stay in their own files.
  - **Layout**: spacing (`--space-N`), font sizes (`--text-N`), border radius (`--radius-N`), shadows (`--shadow-N`), transitions (`--duration-N`), z-index (`--z-N`). Never write raw `rem`/`px` for these in new code; convert touched values to tokens when editing.
  - **Core theme**: `--color-bg`, `--color-surface`, `--color-surface-alt`, `--color-text`, `--color-text-dim`, `--color-accent`, `--color-border`, `--color-selected`, `--color-panel-bg/border/shadow`, `--color-grid`, `--font-mono` — light/dark variants included.
  - **Severity**: `--sev-critical`, `--sev-high`, `--sev-medium`, `--sev-low`, `--sev-na`, `--sev-positive`, `--sev-very-positive`, `--sev-unknown`.
  - **Content types**: `--color-clip`, `--color-screen`, `--color-gif`.
  - **Screenspace tasks**: `--color-task-{tool}` for multitool, color, change, similarity, text, numbers, timelapse, template, flow, scene. In JS, read via `getComputedStyle(document.documentElement).getPropertyValue("--color-task-" + type)` — do not hardcode hex values.
  - **Insight categories**: `--color-causes`, `--color-behaviors`, `--color-impacts` (+ `-bg` variants).
- Thin server, thick client: keep the Flask server focused on data/media endpoints; UI logic, state management, and rendering happen client-side.
- Plan-driven development: detailed implementation plans with specific files, line numbers, code structure, and verification steps are written before coding begins. Follow attached plans closely.
- Features are often built incrementally across multiple sessions. Check for existing groundwork before starting from scratch.
- Manifest-driven state persistence: JSON manifest files (clipgen, insights, screenspace, stashes, settings) follow the pattern of load-on-startup, save-after-mutations.
- No hardcoded version numbers in evergreen docs (CLAUDE.md, README.md). Reference `VERSIONNUM` in `config.py` instead.
- **Icons**: Prefer SVG icons from `assets/icons/` (316 Heroicons outline, kebab-case names like `pencil-square.svg`) over crafting new inline SVG paths or using text/emoji glyphs in web UIs. **Never define SVG path data in JavaScript.** Use the CSS `mask-image` pattern to reference `.svg` files: create a `<span>` with `mask-image: url("icons/icon-name.svg")` and `background-color: currentColor`. See `XREF_BADGES` in `utils.js` and `.xref-badge-icon` in `tokens.css` for the canonical example. Icon routes already exist per blueprint (`/screenspace/icons/`, `/transcripts/icons/`, etc.).
- **Linting/formatting**: Run `uv run ruff check --fix && uv run ruff format` after editing Python files. Run `uv run ty check` for type checking.
- **Pre-commit format gate**: Before every `git commit`, run `uv run ruff format --check` on all modified `.py` files. If any would be reformatted, run `uv run ruff format` on them before committing. This catches files missed by the per-file PostToolUse hook (e.g. in worktrees where the hook is absent). The most common CI lint failure by far is unformatted Python code.
- Commit early and commit often, so we can roll back changes more easily.
- If a problem is reoccurring and survives fix attempts, check git logs for clues.
- Never edit .gitignore automatically, always confirm changes to this file with the user.
- When working through a plan file, e.g. FEATURE-PLAN.md, always make sure to check off items after they are completed.
- **No backwards-compatibility layers.** Do not add migration shims, schema-version checks for hard-break detection, legacy-format readers, or "warn and ignore" branches for old persisted state (manifests, stashes, sessionStorage queues, settings files, etc.). The user base is small and the work is ephemeral; just change the shape and let users re-run. Tests covering persisted shapes get updated to the new shape, not gated behind a version flag.

### Learned Workspace Facts

- Baseline time row placement in the spreadsheet layer is tied to header/`id_cell` row math (e.g. offsets from `id_cell.row`); changing that offset without aligning tests and sheet layout has broken baseline timestamp handling before.
- When making substantive code changes (fixes or features), increment the patch (last number) of VERSIONNUM in config.py.
- Interactive prompts use a keyword-aware helper: `quit`/`exit` exit the program, `top` returns to spreadsheet selection, and `back` returns to mode selection (or spreadsheet selection when already at mode selection).
- Textual-based TUI support (tui.py, TEXTUAL_TUI) has been removed; prefer CLI prompts and the HTML timeline viewer for interactive features.
- Browse mode scrolling is controlled via `BROWSE_LINES_TO_SCROLL` in `config.py`, with a default of 5 rows per up/down step.
- Always use `uv run` to execute Python commands (e.g. `uv run clipgen.py`). This ensures the correct venv is used, even in worktrees where no `.venv` exists yet.
- **Running tests:** from the project root, `uv run --extra dev pytest -c tests/pytest.ini`. Pytest is only in the optional `dev` extra (`pyproject.toml`); `uv sync` alone does not install it, intentionally, so runtime installs stay lean. Do not use ad-hoc `pip install pytest` — use this command so versions match `uv.lock`.
- Be careful about using the `generate_list()`, `sheet.find()`, `sheet.get_all_values()` methods as they are API calls to Google Sheets and are heavily rate-limited. Repeatedly calling the Google API will lead to rate-limiting without warnings, which can appear as bugs (e.g. silently skipping timestamps) and make development difficult.
- CI uses `uv pip install --torch-backend cpu` to avoid downloading ~2.5GB of CUDA/nvidia packages (tests never use CUDA). This override is CI-only (in `tests.yml`), not in `pyproject.toml`, so end-user installs still get GPU-capable torch. If Linux end users emerge and report CUDA issues, check that the CI-only approach hasn't leaked into project config.
- Timelapse produces a single output file, not per-frame timeline events. It does not need icon/color entries in Viewer's `SS_DETECTOR_COLORS`/`SS_DETECTOR_ICON_PATHS` maps.

### Performance Principles

Patterns that should be applied from the start when writing new features, so that dedicated optimization passes are not needed later.

#### Avoid redundant I/O and API calls

- **Never re-fetch what you already have.** If a function needs data that a caller already holds (e.g. `SheetContext`, parsed manifest), accept it as an optional parameter rather than re-reading from disk or network. `generate_list()` now takes `ctx: Optional[SheetContext]`; follow this pattern for any function that calls `build_sheet_context`, `get_all_values`, or reads a manifest file.
- **Read a file once, extract multiple keys.** When you need both artifacts and reels (or any two keys) from the same JSON file, use a single read/parse. See `viewer._load_manifest_both()`. Never call two separate load functions that each read the same file.
- **Google Sheets API calls are precious.** Every `sheet.get_all_values()` / `sheet.find()` / `generate_list()` is a network round-trip subject to rate limits. In server routes, always reuse the cached `_sheet_context` rather than rebuilding it.

#### Design for parallelism from the start

- **Batch first, iterate second.** When processing N independent items (clips, screenshots, reel segments), collect them into a list and process with `ThreadPoolExecutor`, not a sequential for-loop. Use `_resolve_clip_workers()` for the worker count and gate on `len(items) >= 2`.
- **Return results, don't mutate shared state.** Functions that run inside a thread pool must return their output rather than appending to a closure list. Assemble ordered results from the return values after the pool completes (use a pre-allocated results list indexed by future). See `process_reel_clip` returning `(segment_paths, component_dicts)` instead of appending to a shared `components` list.
- **Streaming + parallelism can coexist.** For ndjson-streaming routes, split into two passes: (1) yield cached/skipped items immediately, (2) submit remaining work to a thread pool and yield per-future results via `as_completed()`. This preserves the per-item streaming contract while enabling parallel execution. See `/api/generate` in `server.py`.

#### Pre-compute outside hot loops

- **Normalize comparison data once.** If a loop compares against a set of strings (e.g. filename matching), lowercase / normalize the set once before the loop, not inside each iteration. Sorting callbacks (`key=lambda`) are called O(n log n) times — avoid per-call work that can be hoisted.
- **Use `DocumentFragment` for DOM batching.** When rendering lists of cards/rows, build all elements in a fragment and append once. Never append per-item inside a loop. Viewer and Screenspace already do this; apply the same pattern in any new UI list.

### Code Review Checklist

Patterns distilled from recurring post-review and post-merge fixes across the project's history.

#### Frontend (JS/CSS)

- **CSS toggle completeness**: Every JS class toggle (`.hidden`, `.active`, `.disabled`, etc.) must have a corresponding CSS rule. Verify the rule exists in the stylesheet, not just the JS call.
- **Falsy-safe DOM helpers**: Use `!== undefined` or `!= null` instead of `if (x)` when `0`, `""`, or `false` are valid values.
- **Event listener cleanup**: When rebuilding UI or re-initializing components (e.g. color pickers, modals), remove previous `document`-level listeners before adding new ones. Store references for cleanup.
- **UI state after DOM rebuilds**: If a function rebuilds a container's `innerHTML`, re-apply transient UI state (filmstrip mode, toggle states, scroll position) after the rebuild.
- **Async race conditions**: For video seeking, image loading, or any async chain that can be re-triggered before completion, use a generation counter to reject stale callbacks. Coalesce rapid-fire requests with `requestAnimationFrame`.
- **Canvas/rendering performance**: RAF-throttle canvas draws and mouse-tracking renders. Cache `getBoundingClientRect()` results instead of calling in loops. Pause polling when the tab is hidden (`document.hidden`).
- **Flex layout**: Elements inside flex containers need explicit `flex: 1` or `min-width: 0` to avoid zero-width collapse. Verify new elements are visible after adding them to flex parents.
- **Autocomplete off on text inputs**: Every `<input type="text">` (static or dynamic) must have `autocomplete="off"` to prevent browser autofill (e.g. contact names). For static HTML use the attribute directly; for JS-created inputs set `.autocomplete = "off"` after creation.

#### Backend (Python / Flask)

- **Route parameter types**: Prefer string route parameters with manual `float()`/`int()` parsing over Flask's `<float:x>` converter — JS may send integers where Flask expects floats, causing silent 404s.
- **JSON serialization safety**: Filter `math.isfinite()` on any float derived from OpenCV or numpy before including in JSON responses. Non-finite floats produce invalid JSON that `JSON.parse` silently drops.
- **numpy/ndarray in JSON**: Exclude numpy arrays and other non-serializable objects from manifest saves and API responses. Convert to lists or omit.
- **Dependency manifests**: When importing a new package, immediately add it to `pyproject.toml`. Missing dependencies surface as silent task failures.
- **All call sites**: When modifying a shared function's signature or adding a new parameter, grep for every call site — not just the one you're working on. Functions like `finalize_timeline_data()` have 5+ callers across CLI, Studio, Viewer, and Screenspace.

#### Type Checking (ty)

`ty` is a blocking CI gate. These rules prevent the most common typecheck failures.

- **Narrow Optional before use**: When a variable can be `None` (e.g. `cap: Optional[cv2.VideoCapture]`, `proc.stdout`, a lookup return), add `assert x is not None` before the first use — with a comment explaining the invariant (e.g. `# guaranteed by stdout=PIPE`). Do not use `# type: ignore` instead.
- **JSON dicts need `cast`**: Iterating over dicts from JSON, `isinstance(item, dict)` narrows to `dict[Unknown, Unknown]`, not `Dict[str, Any]`. After the isinstance guard, use `cast(Dict[str, Any], item)`. Annotate the source list explicitly: `steps: list[dict[str, Any]] = data.get("steps", [])`.
- **Avoid None-initialized result lists**: `[None] * n` forces `List[Optional[T]]` and requires narrowing at every use site. Define a typed empty sentinel (e.g. `_EMPTY: T = (0, [])`) and pre-fill with that.
- **cv2 output parameters**: cv2 type stubs reject `None` for output-array parameters (e.g. `calcOpticalFlowFarneback`). Pass a pre-allocated `np.zeros(...)` array instead.
- **Hoist annotations above branches**: Annotating a variable inside one branch of an if/else does not carry to the other. Declare the annotation before the if (`region: Dict[str, Any]`), then assign in each branch.
- **`list[T] | None` vs `list[T | None]`**: For optional list parameters, write `details: list[str] | None = None`. The form `list[str | None] = None` declares a non-optional list of nullable elements — a different type.
- **Narrow properly, don't suppress**: Replace `# type: ignore[union-attr]` and similar with proper narrowing (`assert`, `isinstance`, `if is not None`). Suppressions hide real bugs.

#### Integration

- **Data contract completeness**: When creating records consumed by the frontend (artifacts, events, tasks), include all fields the renderer expects — even optional ones. Missing fields cause empty/broken cards.
- **New flags in mode detection**: When adding a CLI flag, verify it appears in the mode-detection logic (`cli.py`), not just in argparse definition.
- **Bundled/frozen paths**: Use `utils.get_bundled_assets_root()` for asset resolution, never raw `Path(__file__).parent`. Test that asset paths resolve in both source and PyInstaller environments.
- **No duplicated constants between Python and JS.** Any value that lives in `config.py` (or a Python helper) and that the frontend also needs — severity labels, default clip duration, annotation keyphrases (`!key`), ignored timestamp tokens (`x`) — must flow through `utils.get_frontend_config()`, not be hardcoded in JS. When adding a new one:
  1. Extend `get_frontend_config()` in `utils.py`.
  2. Confirm every consumer already embeds `"config": utils.get_frontend_config()` (server.py `/api/sheet`, insights_server.py `/api/artifacts`, viewer.py `finalize_timeline_data` / gallery / insights viewer). Add it if missing.
  3. Add the default to `CLIPGEN_CONFIG` in `assets/web/utils.js` and extend `clipgenApplyConfig` to copy the new key.
  4. Add an assertion in `tests/test_shared_constants.py` that the JS default matches the Python value.

  Why: severity tables, `!key`, `x`, and the 60s default each previously drifted across 3–5 JS files (`studio.js`, `viewer.js`, `insights-builder.js`, `metadata.js`, `transcripts.js`) because they were independently hardcoded. Renaming `!key` in `config.py` would update the backend silently while the frontend kept stripping the old token. Same risk class as `MARK_CATEGORIES` (already guarded by `test_shared_constants.py`).
