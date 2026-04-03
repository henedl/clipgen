# clipgen – Project context for AI assistants

This document contains stable project facts — architecture, data structures, genuine gotchas, conventions (the things that don't change run-to-run).

@AGENTS.md contains learned behavioral preferences — the things that evolved from past mistakes.

## Project overview

clipgen is a Python CLI tool that generates clips from timestamps stored in a Google Sheet or a local Excel file. It uses **gspread** for Google Sheets access, **openpyxl** for Excel, and **ffmpeg/ffprobe** for media processing. The target audience is UX Researchers and professionals who manage playtest videos locally.

**Data flow:** Timestamps in spreadsheet → clipgen reads records (description, study, participant ID, category) → timestamp parsing/annotation filtering → ffmpeg → video clips, screenshots, GIFs, or a single reel. Optionally, generated artifacts can be transcribed via faster-whisper to produce timestamped transcript files.

## Architecture

| File | Role |
| ------ | ------ |
| [clipgen.py](clipgen.py) | Entry point (`python clipgen.py`), spreadsheet opening helpers, interactive mode dispatch, clip/reel processing |
| [viewer.py](viewer.py) | Timeline viewer: artifact record building, data finalization, HTML generation with inlined CSS/JS |
| [cli.py](cli.py) | CLI argument parsing, CLI mode detection, setup, Google auth, worksheet selection, CLI mode dispatch, `main()` |
| [spreadsheet.py](spreadsheet.py) | Spreadsheet parsing, header validation, selector parsing (`reel` input), pure timestamp generation for all modes (no prompts) |
| [interactive.py](interactive.py) | Interactive prompt helpers for all modes (line/range/cell/category/participant selection, browse mode); keeps generation functions pure |
| [video.py](video.py) | ffmpeg/ffprobe operations: cut clips, screenshots, GIFs, concatenate reels, optional filesize compression |
| [transcripts.py](transcripts.py) | Transcription via faster-whisper: `transcribe_video()`, segment filtering, write/read transcript files (Markdown/SRT/VTT) |
| [titlecards.py](titlecards.py) | Titlecard/endcard generation: `build_titlecard_frame()`, `build_endcard_frame()`, `prepend_titlecard_to_clip()`, `append_endcard_to_clip()` — prepends/appends short FFmpeg video cards with text overlays |
| [files.py](files.py) | Filename handling (unique names, truncation), `prepare_clip()` (parse timestamps + annotations, sanitize desc/category), clip discovery for reel-late |
| [utils.py](utils.py) | Timestamp parsing, cell/header annotation parsing, rich/plain output helpers, progress bar utilities, keyword-aware input helpers |
| [config.py](config.py) | Global constants and settings (version, headers, limits, commands) |
| [google_api.py](google_api.py) | Google Sheets auth, worksheet selection by priority, spreadsheet listing/search |
| [excel_io.py](excel_io.py) | Excel adapter: `ExcelSheetAdapter` mimics gspread Worksheet interface for local .xlsx |
| [server.py](server.py) | Combined Flask server for Studio + Insights + Screenspace; registers blueprints per active mode, `start_combined_server()` handles all three on one port |
| [screenspace.py](screenspace.py) | Screenspace analysis engine: image analysis primitives, eleven analysis tools (color, change, similarity, text, numbers, timelapse, template, flow, scene, inactivity + multitool chaining), background task queue worker (`ScreenspaceWorker`), manifest persistence |
| [screenspace_server.py](screenspace_server.py) | Screenspace Flask REST API: region CRUD, video frame extraction, task queue management, results retrieval |
| [insights.py](insights.py) | Insights data model: CRUD operations for insight records, insights manifest read/write |
| [insights_server.py](insights_server.py) | Insights Builder Flask REST API: insight CRUD, artifact browsing, sprite sheet generation, viewer export |
| [assets/web/](assets/web/) | Static HTML/JS/CSS templates: timeline viewer (`viewer.html/js/css`), gallery (`gallery.html/js/css`), studio (`studio.html/js/css`), insights builder (`insights-builder.html/js/css`), insights viewer (`insights-viewer.html/js/css`), screenspace (`screenspace.html/js/css`) |

### Timeline HTML Viewer

Opt-in via `--viewer` or interactive `viewer`. Injects `window.CLIPGEN_DATA` into `assets/web/viewer.html` replacing `<!-- CLIPGEN_DATA_HERE -->`. Data contract, key functions, filmstrip mode, and gallery variant are documented in the [viewer.py](viewer.py) module docstring.

### Gallery HTML Viewer

Opt-in via `--gallery [VIDEO]` or interactive `gv`/`gallery`. Uses the same injection pattern as the timeline viewer but with `assets/web/gallery.html`. Key functions and data contract are documented in the [viewer.py](viewer.py) module docstring.

### Studio Web Interface

Opt-in via `--studio`; requires a spreadsheet. Starts a Flask app via `start_combined_server()` in [server.py](server.py) at `config.SERVER_PORT` (8089), serving `assets/web/studio.html`. API endpoints, module-level state, and key function signatures are documented in the [server.py](server.py) module docstring.

### Insights Builder ([insights.py](insights.py), [insights_server.py](insights_server.py))

Opt-in via `--insights`; no spreadsheet required — reads from `clipgen_manifest.json`. Served at `/insights/` by the same Flask server as Studio. Insight record shape and CRUD function signatures are in the [insights.py](insights.py) module docstring. Flask API endpoints are in the [insights_server.py](insights_server.py) module docstring. The exported `insights_viewer.html` is generated by `finalize_insights_viewer_data()` / `generate_insights_viewer()` in [viewer.py](viewer.py).

### Screenspace ([screenspace.py](screenspace.py), [screenspace_server.py](screenspace_server.py))

Opt-in via `--screenspace` or interactive `ss`/`screenspace`; no spreadsheet required. Served at `/screenspace/` by the combined Flask server. Analysis tool descriptions (color/change/similarity/text/numbers/timelapse/template/flow/scene/inactivity) and API endpoints are documented in the [screenspace.py](screenspace.py) and [screenspace_server.py](screenspace_server.py) module docstrings.

### Artifact Manifest

Opt-in via `--manifest` or `config.MANIFEST_ENABLED`. Writes `clipgen_manifest.json` alongside artifacts. Key functions (`save_manifest`, `load_manifest_artifacts`, `load_manifest_reels`) and merge/dedup behavior are documented in the [viewer.py](viewer.py) module docstring. Consumed by Insights Builder, `--regenerate`, and standalone `--viewer`.

### Transcription ([transcripts.py](transcripts.py))

Opt-in via `--transcribe` or `config.TRANSCRIBE_ENABLED`. Uses faster-whisper; model is lazy-loaded and cached per session. Data types, key function signatures, and pipeline integration details are documented in the [transcripts.py](transcripts.py) module docstring.

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

316 Heroicons (outline, 24×24) live in [assets/icons/](assets/icons/) with kebab-case filenames. Use these for all web UI icons rather than writing new inline SVG paths. See the comment near `svgEditIcon()` in `screenspace.js` for the `createElementNS()` pattern.

## Conventions and patterns

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

- Run the test suite with `uv run pytest -c tests/pytest.ini` from the project root.
- Tests cover: CLI argument parsing, CLI mode dispatch, clip pipeline, file/artifact handling, Google/Excel adapters, insights data model, insights API, manifest operations, selectors, spreadsheet generation, studio API, titlecards, transcripts, timestamp utilities, video commands, viewer data, and viewer inlining.
- Every new CLI mode, flag, or selector should include at least one smoke test in the same PR.
- With `config.DEBUGGING = True`, icecream is enabled, [video.py](video.py) does not invoke ffmpeg, and [transcripts.py](transcripts.py) returns stub results without loading a Whisper model.
