# Project learnings for agents

## Learned User Preferences

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
- **Linting/formatting**: Run `uv run ruff check --fix && uv run ruff format` after editing Python files. Run `uv run ty check` for type checking. After editing frontend JavaScript under `assets/web/`, run `bun run lint:js`; use `bun run lint:js:fix` for safe auto-fixes. A pre-commit hook installed by `scripts/install-git-hooks.sh` also runs `bun run lint:js` on staged JS files automatically.
- **Pre-commit format gate**: Before every `git commit`, run `uv run ruff format --check` on all modified `.py` files. If any would be reformatted, run `uv run ruff format` on them before committing. This catches files missed by the per-file PostToolUse hook (e.g. in worktrees where the hook is absent). The most common CI lint failure by far is unformatted Python code.
- Commit early and commit often, so we can roll back changes more easily.
- If a problem is reoccurring and survives fix attempts, check git logs for clues.
- Never edit .gitignore automatically, always confirm changes to this file with the user.
- When working through a plan file, e.g. FEATURE-PLAN.md, always make sure to check off items after they are completed.
- **No backwards-compatibility layers.** Do not add migration shims, schema-version checks for hard-break detection, legacy-format readers, or "warn and ignore" branches for old persisted state (manifests, stashes, sessionStorage queues, settings files, etc.). The user base is small and the work is ephemeral; just change the shape and let users re-run. Tests covering persisted shapes get updated to the new shape, not gated behind a version flag.

## Learned Workspace Facts

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

## Performance Principles

Patterns that should be applied from the start when writing new features, so that dedicated optimization passes are not needed later.

### Avoid redundant I/O and API calls

- **Never re-fetch what you already have.** If a function needs data that a caller already holds (e.g. `SheetContext`, parsed manifest), accept it as an optional parameter rather than re-reading from disk or network. `generate_list()` now takes `ctx: Optional[SheetContext]`; follow this pattern for any function that calls `build_sheet_context`, `get_all_values`, or reads a manifest file.
- **Read a file once, extract multiple keys.** When you need both artifacts and reels (or any two keys) from the same JSON file, use a single read/parse. See `viewer._load_manifest_both()`. Never call two separate load functions that each read the same file.
- **Google Sheets API calls are precious.** Every `sheet.get_all_values()` / `sheet.find()` / `generate_list()` is a network round-trip subject to rate limits. In server routes, always reuse the cached `_sheet_context` rather than rebuilding it.

### Design for parallelism from the start

- **Batch first, iterate second.** When processing N independent items (clips, screenshots, reel segments), collect them into a list and process with `ThreadPoolExecutor`, not a sequential for-loop. Use `_resolve_clip_workers()` for the worker count and gate on `len(items) >= 2`.
- **Return results, don't mutate shared state.** Functions that run inside a thread pool must return their output rather than appending to a closure list. Assemble ordered results from the return values after the pool completes (use a pre-allocated results list indexed by future). See `process_reel_clip` returning `(segment_paths, component_dicts)` instead of appending to a shared `components` list.
- **Streaming + parallelism can coexist.** For ndjson-streaming routes, split into two passes: (1) yield cached/skipped items immediately, (2) submit remaining work to a thread pool and yield per-future results via `as_completed()`. This preserves the per-item streaming contract while enabling parallel execution. See `/api/generate` in `server.py`.

### Pre-compute outside hot loops

- **Normalize comparison data once.** If a loop compares against a set of strings (e.g. filename matching), lowercase / normalize the set once before the loop, not inside each iteration. Sorting callbacks (`key=lambda`) are called O(n log n) times — avoid per-call work that can be hoisted.
- **Use `DocumentFragment` for DOM batching.** When rendering lists of cards/rows, build all elements in a fragment and append once. Never append per-item inside a loop. Viewer and Screenspace already do this; apply the same pattern in any new UI list.

## Code Review Checklist

Patterns distilled from recurring post-review and post-merge fixes across the project's history.

### Frontend (JS/CSS)

- **CSS toggle completeness**: Every JS class toggle (`.hidden`, `.active`, `.disabled`, etc.) must have a corresponding CSS rule. Verify the rule exists in the stylesheet, not just the JS call.
- **Falsy-safe DOM helpers**: Use `!== undefined` or `!= null` instead of `if (x)` when `0`, `""`, or `false` are valid values.
- **Event listener cleanup**: When rebuilding UI or re-initializing components (e.g. color pickers, modals), remove previous `document`-level listeners before adding new ones. Store references for cleanup.
- **UI state after DOM rebuilds**: If a function rebuilds a container's `innerHTML`, re-apply transient UI state (filmstrip mode, toggle states, scroll position) after the rebuild.
- **Async race conditions**: For video seeking, image loading, or any async chain that can be re-triggered before completion, use a generation counter to reject stale callbacks. Coalesce rapid-fire requests with `requestAnimationFrame`.
- **Canvas/rendering performance**: RAF-throttle canvas draws and mouse-tracking renders. Cache `getBoundingClientRect()` results instead of calling in loops. Pause polling when the tab is hidden (`document.hidden`).
- **Flex layout**: Elements inside flex containers need explicit `flex: 1` or `min-width: 0` to avoid zero-width collapse. Verify new elements are visible after adding them to flex parents.
- **Autocomplete off on text inputs**: Every `<input type="text">` (static or dynamic) must have `autocomplete="off"` to prevent browser autofill (e.g. contact names). For static HTML use the attribute directly; for JS-created inputs set `.autocomplete = "off"` after creation.

### Backend (Python / Flask)

- **Route parameter types**: Prefer string route parameters with manual `float()`/`int()` parsing over Flask's `<float:x>` converter — JS may send integers where Flask expects floats, causing silent 404s.
- **JSON serialization safety**: Filter `math.isfinite()` on any float derived from OpenCV or numpy before including in JSON responses. Non-finite floats produce invalid JSON that `JSON.parse` silently drops.
- **numpy/ndarray in JSON**: Exclude numpy arrays and other non-serializable objects from manifest saves and API responses. Convert to lists or omit.
- **Dependency manifests**: When importing a new package, immediately add it to `pyproject.toml`. Missing dependencies surface as silent task failures.
- **All call sites**: When modifying a shared function's signature or adding a new parameter, grep for every call site — not just the one you're working on. Functions like `finalize_timeline_data()` have 5+ callers across CLI, Studio, Viewer, and Screenspace.

### Type Checking (ty)

`ty` is a blocking CI gate. These rules prevent the most common typecheck failures.

- **Narrow Optional before use**: When a variable can be `None` (e.g. `cap: Optional[cv2.VideoCapture]`, `proc.stdout`, a lookup return), add `assert x is not None` before the first use — with a comment explaining the invariant (e.g. `# guaranteed by stdout=PIPE`). Do not use `# type: ignore` instead.
- **JSON dicts need `cast`**: Iterating over dicts from JSON, `isinstance(item, dict)` narrows to `dict[Unknown, Unknown]`, not `Dict[str, Any]`. After the isinstance guard, use `cast(Dict[str, Any], item)`. Annotate the source list explicitly: `steps: list[dict[str, Any]] = data.get("steps", [])`.
- **Avoid None-initialized result lists**: `[None] * n` forces `List[Optional[T]]` and requires narrowing at every use site. Define a typed empty sentinel (e.g. `_EMPTY: T = (0, [])`) and pre-fill with that.
- **cv2 output parameters**: cv2 type stubs reject `None` for output-array parameters (e.g. `calcOpticalFlowFarneback`). Pass a pre-allocated `np.zeros(...)` array instead.
- **Hoist annotations above branches**: Annotating a variable inside one branch of an if/else does not carry to the other. Declare the annotation before the if (`region: Dict[str, Any]`), then assign in each branch.
- **`list[T] | None` vs `list[T | None]`**: For optional list parameters, write `details: list[str] | None = None`. The form `list[str | None] = None` declares a non-optional list of nullable elements — a different type.
- **Narrow properly, don't suppress**: Replace `# type: ignore[union-attr]` and similar with proper narrowing (`assert`, `isinstance`, `if is not None`). Suppressions hide real bugs.

### Integration

- **Data contract completeness**: When creating records consumed by the frontend (artifacts, events, tasks), include all fields the renderer expects — even optional ones. Missing fields cause empty/broken cards.
- **New flags in mode detection**: When adding a CLI flag, verify it appears in the mode-detection logic (`cli.py`), not just in argparse definition.
- **Bundled/frozen paths**: Use `utils.get_bundled_assets_root()` for asset resolution, never raw `Path(__file__).parent`. Test that asset paths resolve in both source and PyInstaller environments.
- **No duplicated constants between Python and JS.** Any value that lives in `config.py` (or a Python helper) and that the frontend also needs — severity labels, default clip duration, annotation keyphrases (`!key`), ignored timestamp tokens (`x`) — must flow through `utils.get_frontend_config()`, not be hardcoded in JS. When adding a new one:
  1. Extend `get_frontend_config()` in `utils.py`.
  2. Confirm every consumer already embeds `"config": utils.get_frontend_config()` (server.py `/api/sheet`, insights_server.py `/api/artifacts`, viewer.py `finalize_timeline_data` / gallery / insights viewer). Add it if missing.
  3. Add the default to `CLIPGEN_CONFIG` in `assets/web/utils.js` and extend `clipgenApplyConfig` to copy the new key.
  4. Add an assertion in `tests/test_shared_constants.py` that the JS default matches the Python value.

  Why: severity tables, `!key`, `x`, and the 60s default each previously drifted across 3–5 JS files (`studio.js`, `viewer.js`, `insights-builder.js`, `metadata.js`, `transcripts.js`) because they were independently hardcoded. Renaming `!key` in `config.py` would update the backend silently while the frontend kept stripping the old token. Same risk class as `MARK_CATEGORIES` (already guarded by `test_shared_constants.py`).
