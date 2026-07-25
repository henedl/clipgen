# clipgen – Project context for AI assistants

This document contains stable project facts: architecture, data structures, genuine gotchas, conventions (the things that don't change run-to-run).

## Project overview

clipgen is a tool for user researchers that:

1. Generates clips from timestamps stored in a Google Sheet or a local Excel file. It uses **gspread** for Google Sheets access, **openpyxl** for Excel, and **ffmpeg/ffprobe** for media processing.
2. Allows video and audio analysis of local video files, in the **Screenspace** and **Transcript** tools.

**Data flow:** Timestamps in spreadsheet → clipgen reads records (description, study, participant ID, category) → timestamp parsing/annotation filtering → ffmpeg → video clips, screenshots, GIFs, or a single reel. Optionally, generated artifacts can be transcribed via faster-whisper to produce timestamped transcript files.

## Quick commands

```bash
uv run clipgen.py --help                                    # full mode and flag reference
uv run --extra dev pytest -c tests/pytest.ini               # tests (see agents/skills/test/SKILL.md)
uv run ruff format --check && uv run ruff check --fix && uv run ty check  # pre-commit (see agents/skills/check/SKILL.md)

uv run clipgen.py --studio                                  # spreadsheet UI (requires -s)
uv run clipgen.py --screenspace -i INPUT_DIR -o OUTPUT_DIR  # http://127.0.0.1:8089/screenspace/
uv run clipgen.py --transcripts -i INPUT_DIR -o OUTPUT_DIR  # http://127.0.0.1:8089/transcripts/
uv run clipgen.py --composer -i INPUT_DIR -o OUTPUT_DIR     # http://127.0.0.1:8089/composer/
```

Always use `uv run` instead of `python`. Use `uv add` to add dependencies.

## Other files

@agents/CLOUD.md — cloud agent environment and git attribution.
@agents/ARCHITECTURE.md — module roles and subsystem launch notes.
@agents/PERFORMANCE.md — I/O, parallelism, and hot-loop patterns.
@agents/CODE-REVIEW.md — recurring frontend, backend, ty, and integration checks.
@agents/CONTRIBUTING.md — PR/commit conventions and the pre-merge checklist.
@agents/skills/README.md — development and CLI skill procedures.

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
    'times': [(start, end)]    # Added by files.prepare_clip(): list of (start_time, end_time) strings
}
```

Source video filenames follow `{study}_{participant}.mp4` (e.g. `mystudy_P01.mp4`).

## Development tools

- **uv** – Always use `uv run` instead of `python` to run scripts (e.g. `uv run clipgen.py`). This ensures the correct venv is used, including in worktrees that don't yet have a `.venv`. Use `uv add` to add dependencies.
- **Ruff** – Linting and formatting: `uv run ruff check --fix` and `uv run ruff format`. Run these yourself. A per-file `PostToolUse` hook exists in the maintainer's main checkout, but `.gitignore` excludes `.claude/settings.json`, so it is **absent in worktrees** — unformatted Python is the most common CI failure precisely because agents assume the hook ran.
- **ty** – Use `uv run ty check` for type checking.

## SVG icons

316 Heroicons (outline, 24×24) live in [assets/icons/](assets/icons/) with kebab-case filenames. Use these for all web UI icons. Never define SVG path data in JavaScript. The canonical pattern is CSS `mask-image` referencing `.svg` files: create a `<span>` with `mask-image: url("icons/icon-name.svg")` and `background-color: currentColor`. See `XREF_BADGES` in `utils.js` and `.xref-badge-icon` in `tokens.css`. Icon routes already exist per blueprint (`/screenspace/icons/`, `/transcripts/icons/`, etc.).

**Standing exceptions** (intentional inline `<svg>` / data-URI, do not "fix" to mask-image): loading/pulse **animations** that need `<animateTransform>` or animated child elements (studio.html `.title-spinner`, studio.js `createPulserOverlay`); **brand/file-type glyphs** with no Heroicon equivalent (start-overlay Google Sheets / Excel tab marks); the start-overlay **decorative tool-tile artworks**; and `viewer.css` **data-URI masks**, which must stay self-contained because exported/offline viewers have no `/icons/` route. Each site carries an inline comment saying so.

## Conventions and patterns

- **Shared Python/JS config:** values mirrored on the frontend (severity labels, `DEFAULT_DURATION_SECONDS`, `ANNOTATION_KEYPHRASES`, `IGNORED_TIMESTAMP_TOKENS`) flow through `utils.get_frontend_config()`. Server routes embed it as `"config":` in their JSON; `viewer.py` finalize_* functions embed it in `window.CLIPGEN_DATA`. JS overlays the payload onto `CLIPGEN_CONFIG` in `assets/web/utils.js` via `clipgenApplyConfig()`; the `CLIPGEN_CONFIG` defaults are the offline-fallback for re-opened exported viewers. `tests/test_shared_constants.py` asserts the JS defaults match Python.
- **Coordinates:** gspread uses **1-based** row/col. `sheet.get_all_values()` is a list of lists with **0-based** indices: `sheet_data[row_idx][col_idx]`. Conversions: sheet row = `row_idx + 1`, sheet col = `col_idx + 1`.
- **Timestamps:** Formats `MM:SS` or `HH:MM:SS`. Ranges with `-` (e.g. `1:23-1:45`). Multiple pairs separated by `,`, `;`, `+`, or space. Single time gets end = start + `DEFAULT_DURATION_SECONDS`.
- **Annotations:** `utils.parse_cell_annotations()` strips supported keyphrases (configured in `ANNOTATION_KEYPHRASES`, currently `!key`) before timestamp parsing. Ignored tokens (configured in `IGNORED_TIMESTAMP_TOKENS`, currently `x`) are skipped.
- **Participant IDs:** Headers must start with `P` (individual) or `G` (group); see `config.PARTICIPANT_PREFIXES`.
- **User feedback:** Use `utils.error_print()`, `utils.warning_print()`, `utils.verbose_print()`, `utils.info_print()`. Prefer these over direct `print()` for user-facing messages.
- **Debug:** Set `config.DEBUGGING = True` to enable icecream output, skip ffmpeg execution paths in [video.py](video.py), and return stub transcript results in [transcripts.py](transcripts.py) without loading a Whisper model.
- **Keyboard shortcuts:** all web-frontend hotkeys go through the shared registry in `assets/web/hotkeys.js` (loaded right after utils.js on every page, inlined into exports). Pages call `ClipgenHotkeys.register([{id, handler, when, onRelease, repeat}])` against the `HOTKEY_CATALOG` data literal (ids, labels, default combos) and hand their Escape cascade to `registerEscape(fn)`; never add a bare `document.addEventListener("keydown", ...)` — `tests/test_hotkeys_frontend_source.py` enforces an allowlist. Shared ids (`transport.*`, `nav.*`, `edit.*`, `global.*`) keep behavior uniform across pages; the `?` cheatsheet is auto-generated from what a page registers. User rebinds persist as the `HOTKEY_OVERRIDES` dict setting (Settings → Hotkeys), flow via `get_frontend_config()` → `CLIPGEN_CONFIG.hotkeyOverrides`, and are stripped from exported viewers (`viewer._export_config`) so exports always run the defaults. Escape/Tab are reserved, never rebindable. **Numeral convention:** `Shift+numeral` targets a panel/region for keyboard focus, bare `numeral` is a tool/action within the current context (Studio preview tabs, Screenspace tools) — page code routes the two by giving the transport/nav handlers mutually-exclusive `when` gates on a `focusRegion`-style state field. Arrow keys drive the focused surface; `Escape` returns focus to the page's primary surface (the video player) before running the rest of the Escape cascade. Because `Shift+digit` yields layout-dependent symbols, the registry maps those combos by physical `e.code` — so `Shift+1` works across layouts, but only the plain arrows/`Enter` (not Tab) can be caught for in-panel navigation.
- **Interactive keywords:** All interactive prompts go through `utils.read_user_input()`, which treats first-token commands as:
  - `quit` / `exit` → exit clipgen
  - `top` → return to spreadsheet selection
  - `back` → return to mode selection (or spreadsheet selection if already at mode selection)

## Version

The version lives in [build/VERSION](build/VERSION). Agents bump the patch number as part of any `feat:` PR; `fix:`, `refactor:`, `docs:`, `chore:`, `test:`, `build:`, and `ci:` PRs do not bump. The human may also bump manually at any time. There is no CI auto-bump. See [agents/skills/bump/SKILL.md](agents/skills/bump/SKILL.md).

## Hard rules

- **No backwards-compatibility layers.** Do not add migration shims, schema-version checks for hard-break detection, legacy-format readers, or "warn and ignore" branches for old persisted state (manifests, stashes, sessionStorage queues, settings files, etc.). The user base is small and the work is ephemeral; just change the shape and let users re-run. Tests covering persisted shapes get updated to the new shape, not gated behind a version flag.
- **No duplicated constants between Python and JS.** Any value that lives in `config.py` (or a Python helper) and that the frontend also needs must flow through `utils.get_frontend_config()`, not be hardcoded in JS. Procedure: [agents/skills/sync-constants/SKILL.md](agents/skills/sync-constants/SKILL.md).
- **Don't install heavy software to verify UI changes.** No headless Chromium, Playwright, Puppeteer, or similar pulled in unilaterally. For Studio / Screenspace / Transcripts / Viewer changes, ask the human to test in their browser and report back; if you need in-browser diagnostics, give them a small DevTools console snippet to paste.
- **Pre-commit check:** Before every `git commit`, run [agents/skills/check/SKILL.md](agents/skills/check/SKILL.md).
- **Never edit .gitignore automatically.** Always confirm changes with the user.

## Soft preferences

- When the user attaches an implementation plan that already has created todos, do not edit the plan file; mark those todos in_progress as you work and complete them without recreating the list.
- Don't extract helpers unless it is called more than once.
- Prefer placing generic index/letter conversion utilities (e.g. index_to_letter, letter_to_index) in utils.py rather than in domain-specific modules like files.py.
- Prefer naming new helpers to match existing method naming patterns in the same module.
- Never write a class when a function will do.
- Treat spreadsheet layout and timestamp semantics as domain rules; if tests conflict with these, reconsider or adjust the tests rather than changing core semantics to satisfy them.
- All web UIs use vanilla JavaScript (ES5-style `.then()` chaining, not async/await), hand-written CSS with CSS variables for theming, and plain HTML. No React, TypeScript, CSS frameworks, or build tools.
- **CSS design tokens:** `assets/web/tokens.css` is the single source of truth for shared design values. See the file header for the full token list. Never redefine shared tokens in page CSS; page-specific variables (e.g. `--color-cell-text` in studio) stay in their own files. Never write raw `rem`/`px` for layout spacing, font sizes, radius, shadows, transitions, or z-index in new code; convert touched values to tokens when editing. In JS, read Screenspace task colors via `getComputedStyle(document.documentElement).getPropertyValue("--color-task-" + type)`. Do not hardcode hex values.
- **Buttons: two intentional systems.** `.btn` (base + `.btn-small`/`.btn-icon`/`.btn-primary` in `tokens.css`, used cross-page; page-local `.btn-*` extras stay in their own page CSS, not `tokens.css`) is the roomy classic button. `.cg-btn` (`primitives.css`, **Studio-only**, plus the `createBtn()` factory in `primitives.js`) is the compact, richer system (ghost/solid/bare × sm/md/lg + progress). They coexist on purpose; the long-term direction is to migrate `.btn` onto `.cg-btn` **button-by-button**, not in one sweep, so don't re-duplicate the `.btn` base into page CSS, and don't bulk-rewrite either system.
- Thin server, thick client: keep the Flask server focused on data/media endpoints; UI logic, state management, and rendering happen client-side.
- Plan-driven development: detailed implementation plans with specific files, line numbers, code structure, and verification steps are written before coding begins. Follow attached plans closely.
- Features are often built incrementally across multiple sessions. Check for existing groundwork before starting from scratch.
- Manifest-driven state persistence: JSON manifest files (clipgen, screenspace, transcripts, stashes, settings) follow the pattern of load-on-startup, save-after-mutations.
- No hardcoded version numbers in evergreen docs (AGENTS.md, README.md). Reference `build/VERSION` or call `utils.get_version()` instead.
- Commit early and commit often, so we can roll back changes more easily.
- If a problem is reoccurring and survives fix attempts, check git logs for clues.
- When working through a plan file, e.g. `FEATURE-PLAN.md` or anything under `plans/`, keep its status current **as each unit of work lands, not just at the very end**: check off completed items (or update the status table / set a "Done" marker for prose plans that have no checkboxes) and note anything resolved or descoped. This is a hard expectation, not a nicety. Agents are habitually sloppy here, and the next session (or a reviewer) relies on the plan's recorded state to tell what's actually built vs. still open. If you finish work a plan describes, update that plan in the same change.
- Keep AGENTS.md concise with cross-cutting bird's-eye context; per-module detail belongs in inline comments/docstrings, not duplicated in agent docs.

## Workspace facts

- Baseline time row placement in the spreadsheet layer is tied to header/`id_cell` row math (e.g. offsets from `id_cell.row`); changing that offset without aligning tests and sheet layout has broken baseline timestamp handling before.
- Textual-based TUI support (tui.py, TEXTUAL_TUI) has been removed; prefer CLI prompts and the HTML timeline viewer for interactive features.
- Browse mode scrolling is controlled via `BROWSE_LINES_TO_SCROLL` in `config.py`, with a default of 5 rows per up/down step.
- Be careful about using the `generate_list()`, `sheet.find()`, `sheet.get_all_values()` methods as they are API calls to Google Sheets and are heavily rate-limited. Repeatedly calling the Google API will lead to rate-limiting without warnings, which can appear as bugs (e.g. silently skipping timestamps) and make development difficult.
- CI uses `uv pip install --torch-backend cpu` to avoid downloading ~2.5GB of CUDA/nvidia packages (tests never use CUDA). This override is CI-only (in `tests.yml`), not in `pyproject.toml`, so end-user installs still get GPU-capable torch. If Linux end users emerge and report CUDA issues, check that the CI-only approach hasn't leaked into project config.
- Timelapse produces a single output file, not per-frame timeline events. It does not need icon/color entries in Viewer's `SS_DETECTOR_COLORS`/`SS_DETECTOR_ICON_PATHS` maps.
- When a start time uses `H:MM:SS`, computed end times must preserve the hours component (`seconds_to_timestamp(..., force_hours=True)` or match the start format). Emitting `M:SS` for the end breaks mixed-format pairs and duration parsing.
- Homebrew's default ffmpeg 8.x may be built without libfreetype, so the `drawtext` filter is missing and titlecard encoding fails unless ffmpeg is rebuilt with drawtext support.
- New top-level Python modules must be listed in `pyproject.toml` `[tool.setuptools] py-modules` (this is a flat, package-less layout). Source-tree `pytest` imports them fine and masks the omission; only `uv pip install .` breaks with `ModuleNotFoundError`. Splitting a god-file into new modules without listing them ships an incomplete wheel. `tests/test_packaging.py` guards this (it would have caught both the `screenspace_*` split and a long-missing `friction`). Procedure for a god-file split (py-modules + facade re-exports + patch targets): [agents/skills/split-module/SKILL.md](agents/skills/split-module/SKILL.md).
- **Browser UIs are hub + satellite.** Six pages (Screenspace, Transcripts, Studio, Overview, Workflows, Composer) are a hub script plus feature satellites that share state through a `window.Clipgen*` namespace (`SS`/`TS`/`STUDIO`/`OV`/`WF`/`CO`), with same-named guarded delegators in the hub (`function f(){ return SS.f && SS.f.apply(null, arguments); }`). Two things bite every time: a bare cross-file reference to a moved `var` throws a runtime `ReferenceError` that `node --check` cannot see (route shared mutable state through `state.`/the namespace, never a bare read *or write*), and **load order is a contract** (a satellite may only *destructure* a fn published by an earlier-loading file; otherwise late-bind as `SS.fn(...)` at the call site). Procedure, per-page load orders, and worked state-routing examples: [agents/skills/carve-satellite/SKILL.md](agents/skills/carve-satellite/SKILL.md). File-by-file inventory: the `assets/web/` row in [agents/ARCHITECTURE.md](agents/ARCHITECTURE.md). `tests/test_frontend_satellite_wiring.py` fails CI on any bare cross-file call with no delegator/import.

## Pointers

- **CLI modes and flags** — `uv run clipgen.py --help`. Gotchas: `-H/--highlights` scores a reel by severity/uniqueness/annotations (default 180s budget; pass seconds e.g. `-H 120`). `--studio`, `--screenspace`, and `--transcripts` share one Flask app but are mutually exclusive.
- **Spreadsheet layout** — required columns: **ID**, **Observation**, **Category**; participant columns start with `P` or `G`; optional `Baseline time` row. Details in [spreadsheet.py](spreadsheet.py) module docstring and [README.md](README.md).
- **Tests** — [agents/skills/test/SKILL.md](agents/skills/test/SKILL.md). Every new CLI mode, flag, or selector needs at least one smoke test.
- **Performance** — [agents/PERFORMANCE.md](agents/PERFORMANCE.md).
- **Code review** — [agents/CODE-REVIEW.md](agents/CODE-REVIEW.md).
- **CLI command recipes** — [agents/skills/generate/SKILL.md](agents/skills/generate/SKILL.md).
- **Diagnostics** — [agents/skills/debug/SKILL.md](agents/skills/debug/SKILL.md).
