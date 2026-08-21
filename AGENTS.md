# clipgen – Project context for AI assistants

This document contains project facts: architecture, data structures, gotchas, conventions.

## Project overview

clipgen is a tool for user researchers that generates clips from timestamps. It uses **gspread** for Google Sheets access, **openpyxl** for Excel, and **ffmpeg/ffprobe** for media processing. It also provides video and audio analysis in the **Screenspace** and **Transcript** tools.

## Quick commands

```bash
uv run clipgen.py --help                                    # full mode and flag reference
uv run --extra dev pytest -c tests/pytest.ini               # tests (see agents/skills/test/SKILL.md)
uv run ruff format --check && uv run ruff check --fix && uv run ty check  # pre-commit (see agents/skills/check/SKILL.md)

uv run clipgen.py --studio                                  # spreadsheet UI (requires -s)
uv run clipgen.py --screenspace -i INPUT_DIR -o OUTPUT_DIR  # http://127.0.0.1:8089/screenspace/
uv run clipgen.py --transcripts -i INPUT_DIR -o OUTPUT_DIR  # http://127.0.0.1:8089/transcripts/
uv run clipgen.py --composer -i INPUT_DIR -o OUTPUT_DIR     # http://127.0.0.1:8089/composer/

CLIPGEN_UI_CHECK=1 uv run --extra dev --extra ui pytest -c tests/pytest.ini tests/ui  # headless 6-page smoke
uv run --extra ui python tests/ui/shot.py studio --eval "return CLIPGEN_CONFIG"       # one page: shot + probe
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
    "cell": gspread.Cell,  # Cell with timestamp value (1-based row/col)
    "desc": str,  # Observation text from Observation column
    "study": str,  # Normalized study name (filesystem-safe)
    "participant": str,  # e.g. 'P01', 'G02' from header
    "category": str,  # Row category (sanitized; empty → 'uncategorized')
    "severity": str,  # Row severity (normalized label; empty if no Severity column)
    "times": [(start, end)],  # Set by files.prepare_clip(): (start, end) strings
}
```

Source video filenames follow `{study}_{participant}.mp4` (e.g. `mystudy_P01.mp4`).

## Development tools

- **uv** – Always use `uv run` instead of `python` to run scripts (e.g. `uv run clipgen.py`). This ensures the correct venv is used, including in worktrees that don't yet have a `.venv`. Use `uv add` to add dependencies.
- **Ruff** – Linting and formatting: `uv run ruff check --fix` and `uv run ruff format`. A per-file `PostToolUse` hook runs both on every edited `.py`. It lives in `.claude/settings.json`, which `.gitignore` excludes, so it reaches worktrees only by being copied — [.worktreeinclude](.worktreeinclude) lists it for exactly that reason. Still run the commands yourself before committing: the hook is per-file and absent from any checkout that did not get the copy, and unformatted Python is the most common CI failure.
- **ty** – Use `uv run ty check` for type checking.

## SVG icons

316 Heroicons (16×16 viewBox) live in [assets/icons/](assets/icons/) with kebab-case filenames, and the 380 16px [Octicons](assets/icons/octicon/README.md) in [assets/icons/octicon/](assets/icons/octicon/) alongside them (upstream names, size suffix kept — `icons/octicon/dependabot-16.svg`). Use these for all web UI icons. Never define SVG path data in JavaScript. The canonical pattern is CSS `mask-image` referencing `.svg` files: create a `<span>` with `mask-image: url("icons/icon-name.svg")` and `background-color: currentColor`. See `XREF_BADGES` in `utils.js` and `.xref-badge-icon` in `tokens.css`. Icon routes already exist per blueprint (`/screenspace/icons/`, `/transcripts/icons/`, etc.) and take a `<path:filename>`, so the nested subdirectory needs no server change.

**Standing exceptions** (intentional inline `<svg>` / data-URI, do not "fix" to mask-image): loading/pulse **animations** that need `<animateTransform>` or animated child elements (studio.html `.title-spinner`, studio.js `createPulserOverlay`); **brand/file-type glyphs** with no Heroicon equivalent (start-overlay Google Sheets / Excel tab marks); the start-overlay **decorative tool-tile artworks**; `viewer.css` **data-URI masks**, which must stay self-contained because exported/offline viewers have no `/icons/` route; and `tokens.css`'s **`--select-caret`**, which duplicates `chevron-down.svg` because a `<select>` can carry neither a mask (it would clip its own option text) nor a `::after` in WebKit — the caret has to be a `background-image`, so its color is baked per theme rather than `currentColor`. Each site carries an inline comment saying so.

## Conventions and patterns

- **Shared Python/JS config:** values mirrored on the frontend (severity labels, `DEFAULT_DURATION_SECONDS`, `ANNOTATION_KEYPHRASES`, `IGNORED_TIMESTAMP_TOKENS`) flow through `utils.get_frontend_config()`. Server routes embed it as `"config":` in their JSON; `viewer.py` finalize_* functions embed it in `window.CLIPGEN_DATA`. JS overlays the payload onto `CLIPGEN_CONFIG` in `assets/web/utils.js` via `clipgenApplyConfig()`; the `CLIPGEN_CONFIG` defaults are the offline-fallback for re-opened exported viewers. `tests/test_shared_constants.py` asserts the JS defaults match Python.
- **Coordinates:** gspread uses **1-based** row/col. `sheet.get_all_values()` is a list of lists with **0-based** indices: `sheet_data[row_idx][col_idx]`. Conversions: sheet row = `row_idx + 1`, sheet col = `col_idx + 1`.
- **Timestamps:** Formats `MM:SS` or `HH:MM:SS`. Ranges with `-` (e.g. `1:23-1:45`). Multiple pairs separated by `,`, `;`, `+`, or space. Single time gets end = start + `DEFAULT_DURATION_SECONDS`.
- **Annotations:** `utils.parse_cell_annotations()` strips supported keyphrases (configured in `ANNOTATION_KEYPHRASES`, currently `!key`) before timestamp parsing. Ignored tokens (configured in `IGNORED_TIMESTAMP_TOKENS`, currently `x`) are skipped.
- **Participant IDs:** Headers must start with `P` (individual) or `G` (group); see `config.PARTICIPANT_PREFIXES`.
- **User feedback:** Use `utils.error_print()`, `utils.warning_print()`, `utils.verbose_print()`, `utils.info_print()`. Prefer these over direct `print()` for user-facing messages.
- **Debug:** Set `config.DEBUGGING = True` to enable icecream output, skip ffmpeg execution paths in [video.py](source/video.py), and return stub transcript results in [transcripts.py](source/transcripts.py) without loading a Whisper model.
- **Keyboard shortcuts:** all web-frontend hotkeys go through the shared registry in `assets/web/hotkeys.js` (loaded right after utils.js on every page, inlined into exports). Pages call `ClipgenHotkeys.register([{id, handler, when, onRelease, repeat}])` against the `HOTKEY_CATALOG` data literal (ids, labels, default combos) and hand their Escape cascade to `registerEscape(fn)`; never add a bare `document.addEventListener("keydown", ...)` — `tests/test_hotkeys_frontend_source.py` enforces an allowlist. Shared ids (`transport.*`, `nav.*`, `edit.*`, `global.*`) keep behavior uniform across pages; the `?` cheatsheet is auto-generated from what a page registers. User rebinds persist as the `HOTKEY_OVERRIDES` dict setting (Settings → Hotkeys), flow via `get_frontend_config()` → `CLIPGEN_CONFIG.hotkeyOverrides`, and are stripped from exported viewers (`viewer._export_config`) so exports always run the defaults. Escape/Tab are reserved, never rebindable. **Numeral convention:** `Shift+numeral` targets a panel/region for keyboard focus, bare `numeral` is a tool/action within the current context (Studio preview tabs, Screenspace tools) — page code routes the two by giving the transport/nav handlers mutually-exclusive `when` gates on a `focusRegion`-style state field. Arrow keys drive the focused surface; `Escape` returns focus to the page's primary surface (the video player) before running the rest of the Escape cascade. Because `Shift+digit` yields layout-dependent symbols, the registry maps those combos by physical `e.code` — so `Shift+1` works across layouts, but only the plain arrows/`Enter` (not Tab) can be caught for in-panel navigation.
- **Interactive keywords:** All interactive prompts go through `utils.read_user_input()`, which treats first-token commands as:
  - `quit` / `exit` → exit clipgen
  - `top` → return to spreadsheet selection
  - `back` → return to mode selection (or spreadsheet selection if already at mode selection)

## Version

The version lives in [build/VERSION](build/VERSION). Agents bump the patch number as part of any `feat:` PR; `fix:`, `refactor:`, `docs:`, `chore:`, `test:`, `build:`, and `ci:` PRs do not bump. The human may also bump manually at any time. There is no CI auto-bump. See [agents/skills/bump/SKILL.md](agents/skills/bump/SKILL.md).

**Do not write `CHANGELOG.md` by hand.** A scheduled cloud agent fills it in on its own cadence, so a hand-added entry only risks duplicating or conflicting with that pass. Bump `build/VERSION` and leave the changelog alone unless the human explicitly asks for an entry. The curated file is also the Highlights section of every GitHub Release ([build/release_notes.py](build/release_notes.py) selects the versions in the tag's range and groups them by tool), so its wording is user-facing twice over.

## Hard rules

- **No backwards-compatibility layers.** Do not add migration shims, schema-version checks for hard-break detection, legacy-format readers, or "warn and ignore" branches for old persisted state (manifests, stashes, sessionStorage queues, settings files, etc.). The user base is small and the work is ephemeral; just change the shape and let users re-run. Tests covering persisted shapes get updated to the new shape, not gated behind a version flag.
- **No duplicated constants between Python and JS.** Any value that lives in `config.py` (or a Python helper) and that the frontend also needs must flow through `utils.get_frontend_config()`, not be hardcoded in JS. Procedure: [agents/skills/sync-constants/SKILL.md](agents/skills/sync-constants/SKILL.md).
- **Never install heavy software without asking; do use what is already installed.** Browsers, torch, and similar large downloads are never pulled in unilaterally — ask first, every time. But verifying UI changes in a browser is now expected rather than forbidden: run [/ui-check](agents/skills/ui-check/SKILL.md), which loads all six pages headless, fails on any uncaught error, and writes screenshots you should **actually look at** with `Read`. For in-browser diagnostics, run the snippet yourself via `tests/ui/shot.py --eval` instead of handing the human a DevTools paste. Ask the human only for what the smoke genuinely cannot see: interaction feel, motion, drag behaviour, and real-media playback.
- **Pre-commit check:** Before every `git commit`, run [agents/skills/check/SKILL.md](agents/skills/check/SKILL.md).
- **Never edit .gitignore automatically.** Always confirm changes with the user.
- **Be concise.** Comment blocks are <= 7 words, function names <= 4 words. User-facing message strings should be <= 10 words. Use an active voice, no stage performances, and pick the most common word when choosing among alternatives.

## Soft preferences

- When the user attaches an implementation plan that already has created todos, do not edit the plan file; mark those todos in_progress as you work and complete them without recreating the list.
- Don't extract helpers unless it is called more than once.
- Prefer placing generic index/letter conversion utilities (e.g. index_to_letter, letter_to_index) in utils.py rather than in domain-specific modules like files.py.
- Prefer naming new helpers to match existing method naming patterns in the same module.
- Never write a class when a function will do.
- Treat spreadsheet layout and timestamp semantics as domain rules; if tests conflict with these, reconsider or adjust the tests rather than changing core semantics to satisfy them.
- All web UIs use vanilla JavaScript (ES5-style `.then()` chaining, not async/await), hand-written CSS with CSS variables for theming, and plain HTML. No React, TypeScript, CSS frameworks, or build tools.
- **CSS design tokens:** `assets/web/tokens.css` is the single source of truth for shared design values. Its header maps the token families; the declarations themselves are the authoritative list. Never redefine shared tokens in page CSS; page-specific variables (e.g. `--color-cell-text` in studio) stay in their own files. Never write raw `rem`/`px` for layout spacing, font sizes, radius, shadows, transitions, or z-index in new code; convert touched values to tokens when editing. In JS, read Screenspace task colors via `getComputedStyle(document.documentElement).getPropertyValue("--color-task-" + type)`. Do not hardcode hex values.
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
- Timelapse produces a single output file, not per-frame timeline events. It does not need icon/color entries in Viewer's `SS_DETECTOR_COLORS`/`SS_DETECTOR_ICON_PATHS` maps.
- When a start time uses `H:MM:SS`, computed end times must preserve the hours component (`seconds_to_timestamp(..., force_hours=True)` or match the start format). Emitting `M:SS` for the end breaks mixed-format pairs and duration parsing.
- Homebrew's default ffmpeg 8.x may be built without libfreetype, so the `drawtext` filter is missing and titlecard encoding fails unless ffmpeg is rebuilt with drawtext support.
- **A fragmented MP4 looks perfectly healthy server-side and is broken in the browser.** OBS's "fragmented recording" writes thousands of `moof`/`mdat` pairs, `mvhd` duration `0` and no sample index; ffmpeg reads it via the `mfra` tail box, so every probe, scan and clip cut is correct and fast. Browsers do not read `mfra`: `video.duration` is `Infinity`, `seekable` grows only as bytes arrive, and a seek past the buffered range silently lands somewhere else — a multi-GB recording is unusable until the whole file has downloaded. Detected by `video.probe_container_seekability()` (bounded pure-Python box walk, no ffprobe) and surfaced as `browser_seekable` on every `/api/participants` entry; the fix is `video.remux_to_faststart()`, a stream copy. Do not try to diagnose this class with ffprobe — it cannot tell the two containers apart.
- **All Python modules live in [source/](source/); the repo root holds only the `clipgen.py` launcher.** `source/` is *not* a package — no `__init__.py` — it is simply put on `sys.path`, so every module is still imported by bare name (`import config`). Three places do that insert: the launcher (skipped when frozen), `tests/conftest.py`, and `tests/ui/shot.py`. A new module goes in `source/` **and** in `pyproject.toml` `[tool.setuptools] py-modules`; source-tree `pytest` imports it fine either way and masks the omission, but `uv pip install .` ships only listed modules, so installed/frozen environments die with `ModuleNotFoundError`. `tests/test_packaging.py` guards all of it — py-modules coverage, no phantom entries, `package-dir`, no stray root `.py`, no `source/__init__.py`, and that the launcher bootstraps from a foreign cwd. Repo-root paths (`assets/`, `build/VERSION`, `CHANGELOG.md`) resolve through `utils.get_bundled_assets_root()`, which is `source/`'s parent — never hand-roll `Path(__file__).parent`. Procedure for a god-file split (py-modules + facade re-exports + patch targets): [agents/skills/split-module/SKILL.md](agents/skills/split-module/SKILL.md).
- **Native window chrome cannot be reasoned about, only measured.** [desktop_chrome.py](source/desktop_chrome.py) reaches into private AppKit view trees (`NSTitlebarContainerView`, `_NSTheme*Widget`, `NSWindowSharingSessionRecipientIndicator`), and none of it is reachable from `/ui-check`, pytest or a browser. Three consecutive fixes were written from plausible models of how those views nest and when AppKit resets them; the first real log disproved all three. Run `uv run clipgen.py --studio --desktop -v` — it prints a deduped inventory of the titlebar's subview shapes on every layout pass — and read the measurements *before* writing the fix. Two things it settled that no amount of reading would: Sequoia lays its screen-sharing pill **over** the traffic lights rather than replacing them (all three stay `shown`, so an `if buttons are gone` branch is dead code), and AppKit's **last** move of the slot posts no notification at all, so observers alone can never get the last word — hence the module's bounded deferred re-check chain. Corollary for any retry loop: its "already correct" predicate must test every value the writer writes, or a partially-scrambled state reads as settled and the loop stops early.
- **Browser UIs are hub + satellite.** Six pages (Screenspace, Transcripts, Studio, Overview, Workflows, Composer) are a hub script plus feature satellites that share state through a `window.Clipgen*` namespace (`SS`/`TS`/`STUDIO`/`OV`/`WF`/`CO`), with same-named guarded delegators in the hub (`function f(){ return SS.f && SS.f.apply(null, arguments); }`). Two things bite every time: a bare cross-file reference to a moved `var` throws a runtime `ReferenceError` that `node --check` cannot see (route shared mutable state through `state.`/the namespace, never a bare read *or write*), and **load order is a contract** (a satellite may only *destructure* a fn published by an earlier-loading file; otherwise late-bind as `SS.fn(...)` at the call site). Procedure, per-page load orders, and worked state-routing examples: [agents/skills/carve-satellite/SKILL.md](agents/skills/carve-satellite/SKILL.md). File-by-file inventory: the `assets/web/` row in [agents/ARCHITECTURE.md](agents/ARCHITECTURE.md). `tests/test_frontend_satellite_wiring.py` fails CI on any bare cross-file call with no delegator/import.
- Studio Build Reel uses two paths. Spreadsheet-only queues call `/studio/api/reel`: Reel panel order and per-segment card removals are ignored; unique cell refs are resolved from the sheet and sorted by row then column, including every timestamp in those cells. Intake or mixed queues call `/studio/api/reel-direct`, which concatenates in panel order from explicit start/end; titlecards and highlights are not applied on that path.

## Pointers

- **CLI modes and flags** — `uv run clipgen.py --help`. Gotchas: `-H/--highlights` scores a reel by severity/uniqueness/annotations (default 180s budget; pass seconds e.g. `-H 120`). `--studio`, `--screenspace`, and `--transcripts` share one Flask app but are mutually exclusive.
- **Spreadsheet layout** — required columns: **ID**, **Observation**, **Category**; participant columns start with `P` or `G`; optional `Baseline time` row. Details in [spreadsheet.py](source/spreadsheet.py) module docstring and [README.md](README.md).
- **Tests** — [agents/skills/test/SKILL.md](agents/skills/test/SKILL.md). Every new CLI mode, flag, or selector needs at least one smoke test. A test that lands in `--durations=20` gets [agents/skills/test-perf/SKILL.md](agents/skills/test-perf/SKILL.md) before it ships.
- **Performance** — [agents/PERFORMANCE.md](agents/PERFORMANCE.md). Measure before optimizing: [agents/skills/profile/SKILL.md](agents/skills/profile/SKILL.md) (`--profile`, `/api/profile`, `shot.py --perf`).
- **Code review** — [agents/CODE-REVIEW.md](agents/CODE-REVIEW.md).
- **CLI command recipes** — [agents/skills/generate/SKILL.md](agents/skills/generate/SKILL.md).
- **Diagnostics** — [agents/skills/debug/SKILL.md](agents/skills/debug/SKILL.md).
