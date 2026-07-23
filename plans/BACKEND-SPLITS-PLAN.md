# Backend god-file splits — deferred plan

> **Status: Split 1 (workflows) done, 2026-07-23.** Designed during the 2026-07 backend refactor
> pass; the hygiene/consolidation stages of that pass shipped separately (dead config constants,
> truncation consolidation, `server_utils.make_debounced_persist`). This file holds the three
> fully-designed god-file splits that were deliberately deferred, plus the items surveyed and
> skipped. Line numbers reference the tree as of that pass — re-verify before executing, but the
> seams and traps were verified against real code and tests, not just outlines. Splits 2 (cli) and
> 3 (utils) remain open.

## Context

cli.py (~3.8k lines), workflows.py (~3.7k), and utils.py (~2.1k) each contain self-contained
subsystems with verified seams. The splits below follow the repo's established pattern
([agents/skills/split-module/SKILL.md](../agents/skills/split-module/SKILL.md)): new modules go in
`pyproject.toml [tool.setuptools] py-modules` (guarded by `tests/test_packaging.py`); facades
re-export every public **and test-touched private** name; test patch targets move to the owning
sibling (re-export only rebinds). Each split should land as its own `refactor:` commit with the full
gate green: `uv run ruff format --check && uv run ruff check --fix && uv run ty check` + full suite.

Recommended order (lowest risk first): workflows → cli → utils.

## Split 1: workflows.py → workflows_catalog.py + workflows_runner.py (facade) — **Done (2026-07-23)**

> Landed as designed, with two drift adjustments: the runner sibling also took the
> post-design additions (`compute_resume_plan`, `TRIGGER_TYPES`/`TRIGGER_TYPE_IDS`,
> `inspectable_sidecar_view`, `NOTE_NODE_TYPE`), and `workflows_server`'s attribute
> surface had grown to 22 names — all covered by the facade. Zero workflows test
> edits, as predicted. Result: workflows.py 4338 → ~2030 (executors + wiring +
> facade), catalog ~1600, runner ~830.

Import DAG: `workflows_catalog` → `workflows_runner` → `workflows` (facade + executors + wiring).

**workflows_catalog.py (~1,300 lines)** — declarative registry; imports `config`/`utils` only,
late-imports `files`/`video` inside adapters:
- `Port`/`ParamSpec`/`NodeType` TypedDicts (wf:173–230), `NodeContext` (127–172), the `NODE_TYPES`
  literal (231–~710), `_SS_DETECTOR_LABELS`/`_SS_DETECTOR_DESCRIPTIONS`/`_INTERVAL_PARAM`/
  `_SS_DETECTOR_SPECS` + the `ss_*` generation loop + `detect` (711–1076), `serialize_catalog`,
  `serialize_adapters`, `BUILTIN_STASHES` (1119–1352), source-descriptor helpers
  (`_study_from_filename`, `_source_descriptor`, `_source_from_events`, `_clip_source_filename`,
  `_DEFAULT_EVENT_CLUSTER_GAP`, 1352–1401 — the adapters need them), the seven `_adapt_*` functions
  (2873–2983), `ADAPTERS` (2984), `_ADAPTER_DESCRIPTIONS`.
- Docstring must state: collection node types and every `execute` key are wired by workflows.py at
  import — import `workflows`, not this module, before running graphs.

**workflows_runner.py (~600 lines)** — wf:3148–3747: `RUN_STATUS_*`/`NODE_STATUS_*`,
`WorkflowCycleError`, `topo_order`, `blueprint_participant_nodes`, `bind_participant`,
`_summarize_value`/`_node_result_summary`, `run_results_dir`, `_inspectable_result`,
`write_node_sidecar`, `WorkflowRunner`. Imports
`from workflows_catalog import NODE_TYPES, ADAPTERS, NodeContext` — safe because the runner reads
`["execute"]` / `ADAPTERS.get(...)` only at runtime, after workflows.py's wiring has run.

**workflows.py keeps (~1,750 lines):** manifest load/save, all `_exec_*` executors + helpers
(1403–2872 minus moved adapters), gate/collection machinery, `_EXECUTORS` + the generated collection
NODE_TYPES loops + the final wiring loop (3011–3147, mutating the **imported shared** `NODE_TYPES`
dict). Facade re-exports every moved public + test-touched name (incl. `_SS_REFERENCE_DETECTORS`,
`_SS_DETECTOR_SPECS`, `WorkflowRunner`, `write_node_sidecar`, `_inspectable_result`, all statuses).

**Verified facts:**
- Tests patch only via `monkeypatch.setitem(workflows.NODE_TYPES[...], "execute", ...)` and
  `setitem(workflows.ADAPTERS, ...)` — dict mutation, identity shared through the facade re-export
  ⇒ **zero workflows test edits**. There are no `setattr(workflows, ...)` patches anywhere in tests.
  Corollary: **move, never copy** the dicts (`dict(NODE_TYPES)` anywhere would break this silently).
- workflows_server's ~19 distinct `workflows.` attribute reads are all covered by the facade.
- The status constants stay duplicated on purpose — see *Skipped items* below; update the "mirrors
  screenspace_manifest's TASK_STATUS_*" comment (moving to workflows_runner.py) to say the
  duplication is intentional and why.

pyproject: add both modules. ARCHITECTURE.md: two new rows + amend the workflows.py row
("executors + wiring + facade over workflows_catalog/workflows_runner") + note the import DAG.
Sanity check: `uv run python -c "import workflows; assert workflows.NODE_TYPES['dedup_events']['execute']"`.

Failure modes: importing catalog/runner directly and executing nodes before `workflows` wires
`execute` (docstring + facade discipline; all current importers go through the facade); a missed
facade re-export (suite catches); copying instead of sharing `NODE_TYPES`.

## Split 2: cli.py → cli_screenspace.py + cli_event_clips.py

Both seams verified self-contained: neither region calls any cli-level function defined outside
itself; `screenspace`/`pipeline` imports are function-local — **keep them function-local** (cv2 must
stay off the CLI startup path; `uv run clipgen.py --help` should stay fast).

**cli_screenspace.py (~860 lines)** — cli:1296–2152: `_SS_VALID_TASK_TYPES`,
`_ss_resolve_videos_for_participant`, `_ss_frame_extractor`, `_ss_hex_to_hsv`,
`_ss_parse_tolerance`, `_ss_parse_scene_ref`, `_ss_build_params`, `_print_ss_table`,
`_run_ss_list_{regions,stashes,tasks}`, `_run_ss_task`, `_ss_run_and_persist_task`,
`_ss_extract_scene_frames`, `_ss_rehydrate_task_media`, `_run_ss_rerun_task`.

**cli_event_clips.py (~610 lines)** — cli:2146–2752: `_SS_CLIPS_CELL_COL`,
`_TRANSCRIPT_CLIPS_CELL_COL`, `_split_csv_set`, `_split_study_participant`,
`_filter_screenspace_events`, `_filter_transcript_segments`, `_cluster_groups`,
`_build_clusters_from_{ss_events,transcript_segments}`, `_truncate_for_filename`, `_run_ss_clips`,
`_run_transcript_clips`, `_post_marks_to_running_server`, `_run_transcript_mark`.

cli.py (→ ~2,380 lines): `import cli_screenspace` + `import cli_event_clips`;
`_dispatch_standalone_mode` calls become module-attribute calls (`cli_screenspace._run_ss_task(args)`,
…). **No facade re-exports** — these are private names with no non-test external consumers;
module-attribute dispatch keeps future monkeypatches on the owning module visible.

**Test churn (the only split with test edits; verified line refs):**
- `tests/test_cli_screenspace_args.py`: import cli_screenspace; re-point every `cli._ss_*` /
  `cli._run_ss_*` call and the two `monkeypatch.setattr(cli, "_ss_resolve_videos_for_participant", ...)`
  (156, 509). Leave `cli.parse_arguments` / `cli._validate_mode_conflicts` untouched.
- `tests/test_multi_video.py`:1003–1010 → `cli_screenspace._ss_resolve_videos_for_participant`.
- `tests/test_cli_event_clip_args.py`: import cli_event_clips; re-point `_filter_*`,
  `_build_clusters_*`, `_run_ss_clips`, `_run_transcript_clips`, `_run_transcript_mark`, and
  `setattr(cli, "_post_marks_to_running_server", ...)` (573, 757).
- Unaffected (verified): test_cli_modes.py (`_generate_cli_clips`, `_dispatch_standalone_mode` stay
  in cli.py), test_cli_args.py, test_cli_thinking_agents_args.py (`_run_summarize`/`_run_citations`
  stay).

pyproject: add both. ARCHITECTURE.md: two rows + amend the cli.py row. Review the diff for import
placement — no lazy import promoted to top-level.

## Split 3: utils.py → utils_output.py + utils_timestamps.py + utils_artifacts.py (utils stays facade + core)

Riskiest split — do last. Import DAG deepest-first: `utils_output` (config + rich only) ←
`utils_timestamps` ← `utils_artifacts` ← `utils` (core + facade). All ~30 importers keep
`import utils` unchanged — **no mass import rewrite**; only the dual-lookup sites below change.

**utils_output.py (~370 lines)** — utils:~130–492: the rich try-import block, `_CLIPGEN_THEME`,
`console`, `RICH_AVAILABLE`, `BrowseRow` (create_browse_table needs it; re-export from utils for
interactive.py), `_use_rich`, `_use_panels`, `use_progress`, `debug_print`, `verbose_print`,
`standard_print`, `_styled_print`, `error_print`, `warning_print`, `info_print`,
`create_browse_table`, `format_browse_rows_plain`, `create_progress_bar`, `run_with_spinner`,
`print_mode_heading`.

**utils_timestamps.py (~530 lines)** — utils:1145–1673: token split/clean helpers,
`get_ignored_timestamp_tokens`, `add_duration`, `parse_cell_annotations`, `timestamp_to_seconds`,
`parse_timestamps`, `_clock_to_seconds`, `seconds_to_timestamp`, `map_global_to_segment`,
`resolve_timeline_segment`, `map_global_range_to_segments`, `convert_clock_pairs_to_relative`,
`cluster_spans` — plus `get_known_annotation_map` (utils:755; annotation-domain, config-only; moves
along because `parse_cell_annotations` calls it). Imports config +
`from utils_output import warning_print, debug_print`.

**utils_artifacts.py (~290 lines)** — utils:953–1144 + 1874–1963: `_resolve_segment_source_fields`,
`_clip_metadata_fields`, `build_artifact_record`, `build_reel_component`,
`participant_id_from_source_name`, `discover_participant_videos`. Needs core path helpers
(`resolve_input_path`, `get_effective_input_dir`) which **stay in utils** → function-local
`import utils` (skill-sanctioned cycle break; precedent: `MultitoolTool.scan`). Do not move the path
helpers: utils-core functions call them bare-name, and `setattr(utils, "resolve_output_path")`
patches (test_studio_api:1525, 1555) must keep hitting the caller's lookup point.

**Traps (each verified against real call sites):**
- **`NO_INPUT_MODE` is a mutable module scalar** — written at runtime by cli.py:3639 and
  server.py:3777, read bare-name inside `read_user_input` (utils.py:1739). A facade re-export of a
  rebindable scalar silently desyncs. It stays in utils core, along with `read_user_input`, the
  navigation exceptions (`QuitProgram`/`TopToSpreadsheet`/`BackToModeSelection`),
  `check_navigation_keywords`, `suggest_close_match`, `set_program_settings`.
- **Dual-lookup sites**: `utils._use_rich` / `utils.use_progress` / `utils.console` are read both
  externally via attribute and bare-name inside the print helpers. The external read sites —
  interactive.py:158-159, 447, 527-530, 553-554; clipgen.py:760; pipeline.py:1561; cli.py:1569, 1584
  (in `_print_ss_table`; lands in cli_screenspace.py if split 2 ran first) — must switch to
  `utils_output.` so a single patch point exists.
- **Test edits**: tests/test_interactive_prompts.py `setattr(utils, "_use_rich"/"use_progress", ...)`
  (152, 160, 293, 307, 325-326 + `_browse_env`) → re-point to `utils_output`. **All other utils
  patches stay on the facade** (verified: their callers look up via `utils.` attribute —
  `discover_participant_videos` (test_workflows_api ×4), `create_progress_bar` (×3 files),
  `parse_cell_annotations`/`has_non_ignored_timestamp_content` (test_files_and_artifacts,
  test_spreadsheet_generation), `resolve_output_path`, `NO_INPUT_MODE`). tests/test_utils_timestamps.py
  needs no changes.
- Pre-commit audit for this split: `grep -oE "utils\.[A-Za-z_]+" *.py tests/*.py | sort -u` diffed
  against the moved-name list to catch facade re-export misses; run test_interactive_prompts.py
  explicitly and confirm the plain-text (non-rich) paths are exercised.

pyproject: add all three. ARCHITECTURE.md: three rows + amend the utils.py row ("core
file/path/naming/input helpers + facade over utils_output/utils_timestamps/utils_artifacts").

## End-of-branch checks (whichever splits land)

- `uv pip install .` into a throwaway venv, then import every new module (SKILL.md install check;
  catches py-modules omissions that source-tree pytest masks).
- `uv run clipgen.py --help` smoke — confirms no lazy import got promoted (should stay fast, no
  cv2/torch load).

## Skipped items (surveyed during the same pass — do not re-litigate without new evidence)

| Item | Why skipped |
|---|---|
| server.py → studio split | Deferred by maintainer decision. ~90% of server.py is Studio logic with 80+ module-level state vars and many test patch targets (`server._worksheet`, `_sheet_context`, `_generated_artifacts`, …); `build_combined_app` blueprint order and the cli.py↔server.py bidirectional import need their own design round. |
| Task-status constant consolidation | Three *intentionally divergent* copies (screenspace_manifest.py:23 with `PAUSED`, transcripts.py:810 without, workflows.py:~3148 `RUN_STATUS_*`/`NODE_STATUS_*` with `SKIPPED`). The only viable import direction (workflows → screenspace_manifest) drags `screenspace_tools`'s top-level `cv2` into the workflows/transcripts import chains; a new shared module for five strings fails the repo's minimalism bar. |
| SSE / NDJSON response helpers | 3–4 lines per site with domain-specific payload builders; below the "helpers used once stay inline" threshold. All NDJSON sites are in server.py (deferred anyway). |
| Global bare-`except Exception` sweep (~14 sites) | Low payoff; annotate opportunistically when touching those files. |
| Flask envelopes / ffmpeg invocation / manifest I/O dedup | Already centralized (server_utils `err`/`ok`, `video.run_ffmpeg_process` + `probe_video_properties`, `utils.load/save_json_manifest`). Survey confirmed no action needed. |
| screenspace.py facade drift | Audited: complete, no drift. |
