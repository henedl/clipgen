# clipgen-carve-satellite — Carve a JS hub into hub + satellite

Splitting a large page script (`screenspace.js`, `transcripts.js`, `studio.js`, `workflows.js`) into a hub plus feature satellites has shipped runtime `ReferenceError`s **at least 3×** (`e4f67b2` screenspace tasks/results, `8c7f347` transcripts, plus the historical `_calibrationGen` / multitool-drag-state bugs). Each file is its own IIFE scope, so a function or `var` in one file is **invisible** to a sibling unless it is routed through the `window.Clipgen*` namespace. `node --check` and the linter cannot see a bare cross-file reference. It is valid syntax that throws at runtime, often aborting page init. Follow this procedure; it encodes the audit the fix commits said was missing ("every satellite-defined function vs. bare hub references").

## Procedure

1. **Move the code, then route the state.** For every mutable `var` the moved code shares with code left behind, do **not** leave a bare cross-file read. Route it through `state.` / `SS.state` (the namespace's shared state object) or a guarded accessor (`if (SS.fn) SS.fn()`). This is the bare-*variable* class the automated test cannot catch (`_segTooltipRaf` in `8c7f347`); audit it by hand.

2. **Add a delegator for every cross-file function call.** For each function the hub still calls but that now lives in a satellite, add a same-named guarded delegator in the hub and publish it from the satellite:
   - Hub: `function findTask() { return SS.findTask && SS.findTask.apply(null, arguments); }`
   - Satellite: `SS.findTask = findTask;`
   The reverse direction matters too: a satellite that calls a hub-owned helper must reach it via the namespace, not a bare name (`8c7f347` `_currentParticipantHasTranscript` was never published → the agents satellite threw).

3. **Respect the load-order contract.** A satellite that **destructures** another file's published fn at load time (`var findTask = SS.findTask;`) must load *after* that file. When the owner loads *later*, late-bind at the call site instead (`SS.findTask(...)`), never destructure. Update the `<script>` order in the page HTML to match.

4. **Watch poll/render gates.** A control gated on poll-driven state must re-render when *any* gate input changes. Read both the immediately-rendered source and the poll-lagged source: `5683a96`'s friction Re-run gate read only `state.participants[].agents.summary` (lags the poll) and stayed stuck until reload; it also had to watch `state.summaryText` (set on render).

## Verify

1. `node --check` each touched `.js` (catches syntax only).
2. **Run the static wiring guard.** It fails on any bare cross-file call with no delegator/import:
   `uv run --extra dev pytest -c tests/pytest.ini tests/test_frontend_satellite_wiring.py`
3. Run `/check`.
4. Ask the human to load the page in their browser and confirm init completes and the carved feature works (per the "no heavy software" rule, do not pull in headless Chromium).

## Reference

The *why* lives in the dense AGENTS.md "Workspace facts" entries for Screenspace / Transcripts / Studio. `tests/test_frontend_satellite_wiring.py` is the automated guard (the JS analogue of `tests/test_packaging.py`); `tests/test_studio_frontend_source.py` / `tests/test_workflows_frontend_source.py` glob `*.js` so source assertions stay valid wherever a function lands.
