# Performance principles

Patterns to apply from the start when writing new features, so dedicated optimization passes are not needed later.

## Avoid redundant I/O and API calls

- **Never re-fetch what you already have.** If a function needs data that a caller already holds (e.g. `SheetContext`, parsed manifest), accept it as an optional parameter rather than re-reading from disk or network. `generate_list()` now takes `ctx: Optional[SheetContext]`; follow this pattern for any function that calls `build_sheet_context`, `get_all_values`, or reads a manifest file.
- **Read a file once, extract multiple keys.** When you need both artifacts and reels (or any two keys) from the same JSON file, use a single read/parse. See `viewer._load_manifest_both()`. Never call two separate load functions that each read the same file.
- **Google Sheets API calls are precious.** Every `sheet.get_all_values()` / `sheet.find()` / `generate_list()` is a network round-trip subject to rate limits. In server routes, always reuse the cached `_sheet_context` rather than rebuilding it.

## Design for parallelism from the start

- **Batch first, iterate second.** When processing N independent items (clips, screenshots, reel segments), collect them into a list and process with `ThreadPoolExecutor`, not a sequential for-loop. Use `_resolve_clip_workers()` for the worker count and gate on `len(items) >= 2`.
- **Return results, don't mutate shared state.** Functions that run inside a thread pool must return their output rather than appending to a closure list. Assemble ordered results from the return values after the pool completes (use a pre-allocated results list indexed by future). See `process_reel_clip` returning `(segment_paths, component_dicts)` instead of appending to a shared `components` list.
- **Streaming + parallelism can coexist.** For ndjson-streaming routes, split into two passes: (1) yield cached/skipped items immediately, (2) submit remaining work to a thread pool and yield per-future results via `as_completed()`. This preserves the per-item streaming contract while enabling parallel execution. See `/api/generate` in `server.py`.

## Pre-compute outside hot loops

- **Normalize comparison data once.** If a loop compares against a set of strings (e.g. filename matching), lowercase / normalize the set once before the loop, not inside each iteration. Sorting callbacks (`key=lambda`) are called O(n log n) times — avoid per-call work that can be hoisted.
- **Use `DocumentFragment` for DOM batching.** When rendering lists of cards/rows, build all elements in a fragment and append once. Never append per-item inside a loop. Viewer and Screenspace already do this; apply the same pattern in any new UI list.
