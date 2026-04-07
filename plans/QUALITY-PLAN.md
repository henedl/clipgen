# Quality Plan

Addressing organic growth, duplication, and overly defensive patterns identified in the vibe-code audit. Grouped by theme, sorted by impact within each group.

---

## 1. Deduplicate Python Manifest I/O

The same load/save JSON pattern is repeated in `viewer.py`, `insights.py`, `screenspace.py`, `transcripts.py`, and `server.py` (stashes, settings).

- [x] Add `utils.load_json_manifest(path, default)` and `utils.save_json_manifest(path, data)` helpers
- [x] Replace all 5+ load sites and 5+ save sites with calls to the shared helpers
- [x] Ensure consistent error handling (log on write failure, return default on read failure)

---

## 2. Consolidate Interactive Prompt Functions

`prompt_category_selection`, `prompt_severity_selection`, and `prompt_keyword_selection` in `interactive.py` are near-identical 70-line functions differing only in source list, display formatter, and fuzzy-match strategy.

- [x] Extract a generic `prompt_multi_selection()` that accepts: item list, display callback, fuzzy-match callback, and prompt text
- [x] Rewrite the three functions as thin wrappers calling the generic helper

---

## 3. Reduce Frontend JS Duplication

`formatTime()`, `qs()`/`qsa()`/`el()`, `severityClass()`, and `apiGet()`/`apiPost()` are cloned across 4-7 standalone JS files.

- [ ] Evaluate whether a shared `utils.js` can be injected at build/inline time alongside each HTML file
- [ ] If not feasible, document the canonical version of each utility and add a comment pointing to it in each copy, so future changes are propagated deliberately
- [ ] Standardize `apiGet`/`apiPost` error handling (screenspace.js checks `!r.ok`, transcripts.js doesn't)

---

## 4. Address Frontend/Backend Constant Mirroring

`transcripts.js` manually mirrors `MARK_CATEGORIES`, `SS_DETECTOR_COLORS`, and badge SVGs from Python/other JS with "mirrored from" comments.

- [ ] Evaluate serving these values from a `/api/constants` endpoint or injecting them into `window.CLIPGEN_DATA` at page load
- [ ] If kept as manual mirrors, add a test or CI check that compares the JS and Python values

---

## 5. Narrow Broad Exception Handling

9 `except Exception:` blocks across `clipgen.py`, `viewer.py`, `screenspace.py`, `transcripts.py`, and `video.py` swallow errors too broadly.

- [ ] Replace each with the narrowest applicable exception type (e.g., `ValueError` for cell A1 conversion, `OSError` for file I/O, `FileNotFoundError` for ffmpeg)
- [ ] Extract the duplicated cell_a1 conversion pattern (`clipgen.py:894`, `viewer.py:89`) into a shared `utils.safe_cell_a1(row, col)` helper

---

## 6. Break Up God Functions

Six functions exceed 150 lines with mixed concerns.

- [ ] `cli.py main()` (~388 lines): extract config-override application, mode detection, and dispatch into separate functions
- [ ] `interactive.py browse_spreadsheet()` (~358 lines): extract display, search, and nested helper functions to module level
- [ ] `screenspace_server.py api_tasks_create()` (~313 lines): separate validation, parameter normalization, and task creation
- [ ] `screenspace.js renderWorkflowParams()` (~536 lines): split into per-workflow-type render functions
- [ ] `studio.js renderIntake()` (~437 lines): separate clustering, filtering, and rendering logic

---

## 7. Unify Static File Serving Across Blueprints

`server.py`, `insights_server.py`, `screenspace_server.py`, and `transcripts_server.py` each independently implement `serve_index()` and `serve_static()` routes with their own module-level `_assets_dir`/`_output_dir` state.

- [ ] Create a shared factory or helper that registers standard static-serving routes on a Blueprint given an assets directory
- [ ] Replace the 4 independent implementations

---

## 8. Untangle the server.py -> clipgen.py Dependency

`server.py` imports `clipgen` (the script-style entry point) to call `process_clips` and related functions. The web layer should not depend on the CLI entry point.

- [ ] Identify which functions in `clipgen.py` the server actually needs (clip processing pipeline)
- [ ] Evaluate whether those can be accessed without importing the full entry-point module (e.g., moving them to a `pipeline.py` or keeping them but making `clipgen.py` a thin wrapper)

---

## 9. Merge or Clarify Timeline Viewer Templates

Both `viewer.html` and `timeline-viewer.html` exist. Studio/CLI per-participant export uses `timeline-viewer.html`; the default viewer uses `viewer.html`.

- [ ] Determine if these can be merged into a single template with a conditional block
- [ ] If they must remain separate, document the distinction clearly and align naming

---

## 10. Tidy Config Access

`config.py` is 301 lines of flat globals mutated from `cli.py`. `server.py` works around this with `_settings_defaults` snapshots and an `_override_config()` context manager.

- [ ] Group related constants (screenspace thresholds, ffmpeg params, file format settings) into named sections or lightweight dataclass-like groupings
- [ ] Extract the CLI config-override block (`cli.py:1300-1312`) into a dedicated `apply_cli_overrides(args)` function

---

## 11. Consolidate Lazy Imports

8 function-local imports of `transcripts`, `screenspace`, `screenspace_server`, and `transcripts_server` are scattered across the codebase as circular-dependency workarounds.

- [ ] Map the actual dependency graph to determine which cycles are real vs. vestigial
- [ ] Where possible, resolve by moving the needed function to a non-cyclic location or by restructuring imports
- [ ] For genuinely optional heavy deps (easyocr), create a shared `utils.try_import(name)` pattern

---

## 12. Move Hardcoded Frontend Values to Tokens

Scattered `rgba()` values, canvas dimensions, poll intervals, and font-size scaling factors are hardcoded in JS and CSS.

- [ ] Add opacity-variant tokens to `tokens.css` (e.g., `--alpha-subtle`, `--alpha-overlay`)
- [ ] Replace hardcoded accent-with-opacity values in `studio.css` with token-derived values
- [ ] Define `POLL_INTERVAL` in one place (currently duplicated in `screenspace.js` and `transcripts.js`)
