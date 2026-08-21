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

@agents/ARCHITECTURE.md — module roles and subsystem launch notes.

Not auto-loaded; read when the trigger applies:

- [agents/CLOUD.md](agents/CLOUD.md) — running in a cloud agent VM? Read before installing deps or committing (git attribution rules).
- [agents/PERFORMANCE.md](agents/PERFORMANCE.md) — read before writing I/O-heavy, parallel, or hot-loop code. Holds the tuning-knob table.
- [agents/CODE-REVIEW.md](agents/CODE-REVIEW.md) — read when reviewing or self-reviewing a diff (`/check` points here).
- [agents/CONTRIBUTING.md](agents/CONTRIBUTING.md) — read before committing or opening a PR.
- [agents/skills/README.md](agents/skills/README.md) — index of skill procedures; each is a slash command.

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

316 Heroicons in [assets/icons/](assets/icons/) (kebab-case names) and 380 16px [Octicons](assets/icons/octicon/README.md) in [assets/icons/octicon/](assets/icons/octicon/) cover all web UI icons; every blueprint serves them at `<prefix>/icons/<path:filename>`. Render an icon as a `<span>` with `mask-image: url("icons/name.svg")` and `background-color: currentColor` (see `XREF_BADGES` in `utils.js`). Never define SVG path data in JavaScript. The standing exceptions (animations, brand glyphs, self-contained exports, `--select-caret`) are the frozen allowlists in `tests/test_icon_conventions.py`; each site carries an inline comment saying why.

## Conventions and patterns

- **Shared Python/JS config:** values mirrored on the frontend (severity labels, `DEFAULT_DURATION_SECONDS`, `ANNOTATION_KEYPHRASES`, `IGNORED_TIMESTAMP_TOKENS`) flow through `utils.get_frontend_config()`. Server routes embed it as `"config":` in their JSON; `viewer.py` finalize_* functions embed it in `window.CLIPGEN_DATA`. JS overlays the payload onto `CLIPGEN_CONFIG` in `assets/web/utils.js` via `clipgenApplyConfig()`; the `CLIPGEN_CONFIG` defaults are the offline-fallback for re-opened exported viewers. `tests/test_shared_constants.py` asserts the JS defaults match Python.
- **Coordinates:** gspread uses **1-based** row/col. `sheet.get_all_values()` is a list of lists with **0-based** indices: `sheet_data[row_idx][col_idx]`. Conversions: sheet row = `row_idx + 1`, sheet col = `col_idx + 1`.
- **Timestamps:** Formats `MM:SS` or `HH:MM:SS`. Ranges with `-` (e.g. `1:23-1:45`). Multiple pairs separated by `,`, `;`, `+`, or space. Single time gets end = start + `DEFAULT_DURATION_SECONDS`.
- **Annotations:** `utils.parse_cell_annotations()` strips supported keyphrases (configured in `ANNOTATION_KEYPHRASES`, currently `!key`) before timestamp parsing. Ignored tokens (configured in `IGNORED_TIMESTAMP_TOKENS`, currently `x`) are skipped.
- **Participant IDs:** Headers must start with `P` (individual) or `G` (group); see `config.PARTICIPANT_PREFIXES`.
- **User feedback:** Use `utils.error_print()`, `utils.warning_print()`, `utils.verbose_print()`, `utils.info_print()`. Prefer these over direct `print()` for user-facing messages.
- **Debug:** Set `config.DEBUGGING = True` to enable icecream output, skip ffmpeg execution paths in [video.py](source/video.py), and return stub transcript results in [transcripts.py](source/transcripts.py) without loading a Whisper model.
- **Keyboard shortcuts:** all web-frontend hotkeys go through the shared registry in `assets/web/hotkeys.js`: pages call `ClipgenHotkeys.register(...)` against the `HOTKEY_CATALOG` literal and hand Escape handling to `registerEscape(fn)`. Never add a bare `document.addEventListener("keydown", ...)` — `tests/test_hotkeys_frontend_source.py` enforces an allowlist. The full conventions (numeral routing, `e.code` mapping, rebind persistence, reserved keys) are in the `hotkeys.js` header comment.
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
- **Be concise.** Comment blocks are <= 15 words, function names <= 4 words. User-facing message strings should be <= 10 words. Use an active voice, no stage performances, and pick the most common word when choosing among alternatives. Legacy code that breaks these limits gets redone as encountered: delete the old comment or name and write a new one from scratch, never trim word-by-word.

## Soft preferences

- When the user attaches an implementation plan that already has created todos, do not edit the plan file; mark those todos in_progress as you work and complete them without recreating the list.
- Code style: don't extract a helper until it has a second caller; never write a class when a function will do; name new helpers after existing patterns in the same module; generic conversion utilities (index_to_letter etc.) go in utils.py, not domain modules.
- Treat spreadsheet layout and timestamp semantics as domain rules; if tests conflict with these, reconsider or adjust the tests rather than changing core semantics to satisfy them.
- Frontend stack is vanilla ES5 JavaScript (`.then()` chaining, no async/await), hand-written CSS with variables for theming, plain HTML. No frameworks, TypeScript, or build tools. Enforced by `test_frontend_syntax.py::test_js_is_es5`.
- **CSS design tokens:** shared design values come from `assets/web/tokens.css` (its header maps the families). No new raw `px`/`rem`/hex/durations/z-index in page CSS and no redefining shared tokens; convert touched values when editing — `tests/test_css_token_discipline.py` ratchets the per-file counts. In JS, read colors via `getComputedStyle(...).getPropertyValue(...)`, never hardcoded hex.
- Buttons: `.btn` (`tokens.css`, cross-page) and `.cg-btn` (`primitives.css`, Studio-only) coexist on purpose; migrate button-by-button, never in one sweep. See the header comments above each.
- Thin server, thick client: keep the Flask server focused on data/media endpoints; UI logic, state management, and rendering happen client-side.
- Plan-driven development: detailed implementation plans with specific files, line numbers, code structure, and verification steps are written before coding begins. Follow attached plans closely.
- Features are often built incrementally across multiple sessions. Check for existing groundwork before starting from scratch.
- Manifest-driven state persistence: JSON manifest files (clipgen, screenspace, transcripts, stashes, settings) follow the pattern of load-on-startup, save-after-mutations.
- No hardcoded version numbers in evergreen docs (AGENTS.md, README.md). Reference `build/VERSION` or call `utils.get_version()` instead.
- When working through a plan file (anything under `plans/` or a `*-PLAN.md`), keep its status current **as each unit of work lands, not just at the very end**: check off completed items, update status tables, note anything resolved or descoped. This is a hard expectation — the next session and reviewers rely on the recorded state to tell what's built vs. still open.
- Keep AGENTS.md concise with cross-cutting bird's-eye context; per-module detail belongs in inline comments/docstrings, not duplicated in agent docs.

## Workspace facts

- Textual-based TUI support (tui.py, TEXTUAL_TUI) has been removed; prefer CLI prompts and the HTML timeline viewer for interactive features.
- `generate_list()`, `sheet.find()`, and `sheet.get_all_values()` are Google Sheets API calls and heavily rate-limited. Repeated calls get throttled **without warnings**, which masquerades as bugs (e.g. silently skipped timestamps) and makes development difficult.
- **A fragmented MP4 (OBS "fragmented recording") looks healthy server-side and is broken in the browser.** Detect with `video.probe_container_seekability()` — never ffprobe, which cannot tell the containers apart — and fix with `video.remux_to_faststart()`. Full story: the probe's docstring in [video.py](source/video.py).
- **All Python modules live in [source/](source/); the repo root holds only the `clipgen.py` launcher.** `source/` is not a package — it goes on `sys.path` and modules import by bare name (`import config`). A new module must also be listed in `pyproject.toml` `[tool.setuptools] py-modules`, or installed/frozen builds die with `ModuleNotFoundError` (source-tree pytest masks the omission; `tests/test_packaging.py` guards it). Repo-root paths resolve through `utils.get_bundled_assets_root()`, never a hand-rolled `Path(__file__).parent`. God-file splits: [agents/skills/split-module/SKILL.md](agents/skills/split-module/SKILL.md).
- **Native window chrome cannot be reasoned about, only measured.** Before touching [desktop_chrome.py](source/desktop_chrome.py), run `uv run clipgen.py --studio --desktop -v` and read the printed titlebar view shapes; three fixes written from plausible models were all wrong. The module docstring holds the full story (Sequoia sharing-pill behavior, the notification gap, the deferred re-check chain).
- **Browser UIs are hub + satellite.** Six pages share state through a `window.Clipgen*` namespace (`SS`/`TS`/`STUDIO`/`OV`/`WF`/`CO`) with same-named guarded delegators in the hub. A bare cross-file reference to a moved `var` throws a runtime `ReferenceError` that `node --check` cannot see, and load order is a contract. Procedure and worked examples: [agents/skills/carve-satellite/SKILL.md](agents/skills/carve-satellite/SKILL.md); file inventory: the `assets/web/` row in [agents/ARCHITECTURE.md](agents/ARCHITECTURE.md); guard: `tests/test_frontend_satellite_wiring.py`.
- Studio Build Reel uses two paths: `/studio/api/reel` (spreadsheet-only queues; re-resolves cells from the sheet, ignores panel order) and `/studio/api/reel-direct` (intake/mixed queues; panel order, no titlecards/highlights). Details in the route docstrings in [server.py](source/server.py).

## Pointers

- **CLI modes and flags** — `uv run clipgen.py --help`. Gotchas: `-H/--highlights` scores a reel by severity/uniqueness/annotations (default 180s budget; pass seconds e.g. `-H 120`). `--studio`, `--screenspace`, and `--transcripts` share one Flask app but are mutually exclusive.
- **Spreadsheet layout** — required columns: **ID**, **Observation**, **Category**; participant columns start with `P` or `G`; optional `Baseline time` row. Details in [spreadsheet.py](source/spreadsheet.py) module docstring and [README.md](README.md).
- **Tests** — [agents/skills/test/SKILL.md](agents/skills/test/SKILL.md). Every new CLI mode, flag, or selector needs at least one smoke test. A test that lands in `--durations=20` gets [agents/skills/test-perf/SKILL.md](agents/skills/test-perf/SKILL.md) before it ships.
- **Performance** — [agents/PERFORMANCE.md](agents/PERFORMANCE.md). Measure before optimizing: [agents/skills/profile/SKILL.md](agents/skills/profile/SKILL.md) (`--profile`, `/api/profile`, `shot.py --perf`).
- **Code review** — [agents/CODE-REVIEW.md](agents/CODE-REVIEW.md).
- **CLI command recipes** — [agents/skills/generate/SKILL.md](agents/skills/generate/SKILL.md).
- **Diagnostics** — [agents/skills/debug/SKILL.md](agents/skills/debug/SKILL.md).
