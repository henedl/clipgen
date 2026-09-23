# Workflows bug hunt

Status: fixed (2026-09-18). Regressions in `tests/test_workflows_runner.py` and
`tests/test_workflows_api.py`.

Scope: workflow execution, resume planning, failure status, batches, and arrival
triggers. No implementation changes. This extends the separate Composer hunt.

## 1. Resume reuses results after execution settings change

Priority: medium. Locations:

- `source/workflows_runner.py:324`, `compute_resume_plan`.
- `source/workflows_runner.py:669`, `WorkflowRunner._run`, especially the seed branch before the disabled check.
- `source/workflows_server.py:538`, `api_run_create` resume handling.

The resume planner compares node IDs, completion status, output type, and output
port coverage. It does not compare parameters, disabled state, or incoming wiring
against the previous run. Neither node states nor sidecars retain the execution
definition needed for those comparisons. Results can therefore be reused even
though they no longer represent the current blueprint.

Verified reproduction:

1. Run a blueprint containing node `range`, type `time_range`, with
   `params: {"ranges":"0:01-0:02"}`.
2. Change that same node's parameter to `"0:10-0:20"`.
3. Call `compute_resume_plan` with the current blueprint, prior node states, and
   the prior run's real sidecar loader; pass its seeds into a new runner.
4. Actual resumed output: `{"timeRange":{"ranges":[[1.0,2.0]],"source":{}}}`.
   Expected output uses `[10.0,20.0]`.
5. Set `disabled:true` on the same node and resume from the original run again.
   Actual status: `completed`, with the original result still available.
   Expected: the node is skipped and downstream behavior respects that state.

The UI-facing scenario is editing a completed upstream node before resuming a
workflow whose later node failed. The API also permits resuming completed runs.
Changed wiring has the same missing-comparison problem; it was inspected in code,
not independently reproduced in this hunt.

Implementation plan:

- [x] Persist the execution definition needed to validate cached results.
- [x] Compare parameters, disabled state, incoming edges, and relevant run options on resume.
- [x] Invalidate changed nodes and their descendants while retaining safe reuse elsewhere.
- [x] Ensure disabled nodes cannot execute through the seed shortcut.
- [x] Add regressions for changed ranges, newly disabled nodes, rewired inputs, and unchanged graphs.
- [x] Cover a failed downstream node resumed after an upstream edit through the API.

Use the new persisted shape directly; do not add schema migrations or legacy readers.

Landed as: sidecars carry an `__exec__` stamp (params, incoming edges, sample
window) written by `write_node_sidecar`; `compute_resume_plan` re-runs any node
whose stamp differs or that is now muted, and the runner checks `disabled` and
gates before accepting a seed.

## 2. Operational failures become successful, reusable results

Priority: medium. Locations:

- `source/workflows.py:479`, `_exec_summarize`, with similar failure returns in other executors.
- `source/workflows_runner.py:768–801`, executor result and status handling.
- `source/workflows_runner.py:324`, completed-node eligibility for resume.

When the AI server cannot start, `_exec_summarize` returns an empty summary and
`__note__` explaining that summary generation was skipped. The runner preserves
the note but calculates degraded status only from adapter and sidecar failures.
It marks both the node and run `completed`, persists the empty result, and makes
that result eligible for reuse.

Verified reproduction:

1. Run a blueprint with a `summarize` node while mocking
   `llm_client.ensure_server` to return `False`.
2. Actual node and run status: `completed`; note:
   `AI server would not start. Summary skipped`; result: `{"summary":""}`.
3. Compute resume seeds from this run's states and real sidecar.
4. Restore successful server startup and execute the resumed runner.
5. Actual: the empty summary is reused, with zero calls to `ensure_server`.
   The failed operation never retries.

Expected: an operational failure is reported as failed or degraded and remains
eligible for execution on resume. An empty successful result must be distinguishable
from work that could not execute. A later failed branch makes this especially
misleading: resuming the run can retry that branch while preserving the empty summary.

Implementation plan:

- [x] Separate informational executor notes from operational failure outcomes.
- [x] Mark AI startup failure as failed or degraded, and propagate that status to the run.
- [x] Audit existing failure returns for the same outcome ambiguity, especially failed exports.
- [x] Keep genuinely informational notes from becoming false failure reports.
- [x] Add a regression using the real summary executor with mocked startup failure.
- [x] Verify resume retries the operation after recovery instead of reusing its empty output.

`tests/test_workflows_runner.py::test_executor_note_surfaces_and_is_stripped`
currently expects a generic note to remain completed. Preserve that distinction
or deliberately refine the outcome contract; treating every note as failure would
also misclassify informational results such as an already-seekable video.

Landed as: a second reserved key, `__degraded__`, marks work that could not run;
`__note__` stays informational. The runner folds `__degraded__` into the node and
run status, and only `completed` nodes seed a resume.

## 3. Batch precomputation bypasses closed gates

Priority: medium. Locations:

- `source/workflows_server.py:777`, `_precompute_shared_nodes`.
- `source/workflows_server.py:854`, `_run_batch` seed preparation.
- `source/workflows_runner.py:701–748`, seed handling before `_should_skip`.

Batch precomputation runs every enabled Sheet Selection node without considering
incoming control edges. Each child receives the result as a seed. The runner
accepts that seed before checking gates, so a closed gate cannot skip the node.
The same graph therefore behaves differently in normal and batch execution.

Verified reproduction:

1. Create a Threshold Gate (`gate_collection`) with count `>= 1` and no collection
   wired, so its real executor returns `{"pass":false}`.
2. Connect `gate.pass` to a Sheet Selection node's universal `__gate__` input.
3. Run normally: Sheet Selection status is `skipped`.
4. Call `_precompute_shared_nodes`, then run a child with those seeds, as the
   batch coordinator does: the gate still returns false, but Sheet Selection is
   `completed` and its output is present.

The reproduction used the actual gate and Sheet Selection executors with no sheet
context; the latter safely returns an empty selection. With a real sheet context,
precomputation also performs the Sheets lookup before any gate can reject it.
Real spreadsheet access and downstream media generation were not exercised.

Expected: precomputation may cache work but must preserve gate-controlled execution
and output visibility for every participant.

Implementation plan:

- [x] Exclude control-dependent Sheet Selection nodes from unconditional precomputation.
- [x] Apply each child's skip and mute decisions before accepting batch seeds.
- [x] Preserve shared caching for independent Sheet Selection sources.
- [x] Add closed/open gate regressions comparing normal and batch child runs.
- [x] Verify a participant-dependent gate can allow one child and block another.
- [x] Verify blocked outputs cannot reach downstream consumers through cached seeds.

Coordinate the seed-order change with finding 1; this bug also occurs in a fresh
batch with an unchanged graph, without any resume operation.

Landed as: `_precompute_shared_nodes` skips any cacheable node with an incoming
edge, and the runner's seed check now follows the mute and gate checks.

## 4. Arrival triggers ignore recording parts still being copied

Priority: medium. Locations:

- `source/workflows_server.py:1088`, `_stat_first_video`.
- `source/workflows_server.py:1134`, `_poll_new_videos`.

The new-video trigger promises two stable polls before launching. Its fingerprint
contains only the first video path's size and modification time. For multipart
participants, changes in subsequent parts are invisible. The watcher fires early
and marks the participant seen, so it does not launch again after those parts
finish copying. Media-dependent work can consequently fail on incomplete inputs
or use an incomplete recording without an automatic retry.

Verified reproduction:

1. Discover P01 with two paths, `study_P01-1.mp4` and `study_P01-2.mp4`.
2. Keep part 1 unchanged; poll once to establish its fingerprint.
3. Append bytes to part 2, then poll again.
4. Actual: `_maybe_fire_trigger("P01", "new_video")` fires despite part 2 changing.
5. Append more bytes to part 2 and then poll twice with both files stable.
6. Actual: no further trigger; P01 is already in `_watch_seen`.

The reproduction used two temporary byte files and real filesystem stat calls,
mocked discovery, and a trigger-call collector. It tested arrival detection only;
the files were not media and no workflow or media process was launched.

Expected: every currently discovered part must remain stable across consecutive
polls before the participant is marked seen and launched.

Implementation plan:

- [x] Fingerprint all ordered video paths, including path identity, size, and modification time.
- [x] Restart stability tracking when any part changes, appears, disappears, or cannot be read.
- [x] Add multipart regressions for a growing later part and changing path membership.
- [x] Verify exactly one launch after all discovered parts stabilize.
- [x] Retain single-file behavior and the no-retroactive-launch rule when arming.

This fix can cover known parts; identifying a future part that has not appeared
yet would require a separate arrival/completion contract and is outside this finding.

Landed as: `_stat_videos` fingerprints every part as `(path, size, mtime)`.

## Verification and handoff

Reproduced on 2026-09-18 using `uv run --no-sync python`, actual `WorkflowRunner`
instances, actual `compute_resume_plan`, and real sidecars inside an automatically
removed temporary directory. The time-range executor ran unchanged. AI startup
was mocked; no AI server, media processing, downloads, or project-state writes
were needed. Batch and watcher reproductions used the same temporary-directory
approach; their specific mocks are recorded above. These were engine/helper-level
reproductions, not browser journeys or live spreadsheet tests.

After implementation, run focused regressions and the existing workflow suites:

```sh
uv run --extra dev pytest -c tests/pytest.ini tests/test_workflows_runner.py tests/test_workflows_executors.py tests/test_workflows_api.py
```

Follow `agents/skills/check/SKILL.md` before committing fixes. The hunt did not run
the full test suite or modify production code or tests.
