# More reliable performance diagnostics

Status: Planned; implementation has not started.

## Summary

Keep the existing opt-in profiler and extend its reporting and benchmark tools.
Prioritize trustworthy comparisons, then realistic browser scenarios, then
concurrent-job attribution.

The current coverage is strong: backend spans, cache counters, startup milestones,
streaming latency, browser timings, deep profiling, and scan/clip benchmarks. The
main weaknesses are how measurements are retained, compared, and explained.

This review used source inspection; it does not establish new performance baselines.

## Handoff and status

| Phase | Status | Completion criterion |
| --- | --- | --- |
| 1. Trustworthy measurements | Not started | Reliable structured exports and benchmark comparisons |
| 2. Browser workloads | Not started | Verified workloads with load, interaction, and idle reports |
| 3. Concurrent work | Not started | Bounded operation records and inspectable deep profiles |

Implement phases in order and ship them independently. Mark checklist items as
they land, update this table, and record validation and unresolved issues below.
Do not mark implementation complete merely because this plan is committed.

Read `agents/PERFORMANCE.md` and `agents/skills/profile/SKILL.md` before implementation.
Use the repository's test, UI-check, and pre-commit procedures as applicable.

## Phase 1 — Make measurements trustworthy

Primary areas: `source/profiling.py`, the profile route in `source/server.py`,
CLI argument handling, and `tests/perf/`.

### Export structured results

- [ ] Add `--profile-output PATH`, implying `--profile`, to write full-precision JSON alongside the existing console report.
- [ ] Include metrics, startup marks, profiling mode, elapsed time, dropped-label counts, and environment metadata.
- [ ] Change benchmark runners to consume JSON directly; remove their console-parsing dependency.
- [ ] Record commit and dirty state, Python/platform, relevant dependency and binary versions, effective tuning settings, and fixture identity.
- [ ] Treat deep-profile runs as diagnostic artifacts; reject them from ordinary timing comparisons.

### Make benchmark success explicit

- [ ] Require successful process exit, required measurements, expected work counts, and completed outputs.
- [ ] Invalidate the scenario when any repetition fails; a successful repetition must not hide failure.
- [ ] Distinguish invalid runs, incomparable workloads, and performance regressions in reports and exit status.
- [ ] Validate fixture parameters and contents, including dimensions, frame rate, sampling interval, regions, and spreadsheet geometry. Duration alone is insufficient.
- [ ] Reject workload mismatches. Report environment differences and require an explicit benchmark override to compare across them.
- [ ] Update persisted result shapes directly, without legacy readers or migrations.

### Retain repeated measurements

- [ ] Default to three measured runs and retain every sample. Report median, minimum, maximum, and median absolute deviation.
- [ ] Compare medians; keep the minimum as supplementary information.
- [ ] Measure subprocess elapsed time as the primary end-to-end metric. Preserve callback, pool, encoding, and post-processing timings as diagnostic breakdowns.
- [ ] Retain `--fail-on` as the percentage threshold; add metric selection and an optional absolute threshold. When both thresholds are supplied, require both to be exceeded.
- [ ] Describe runs as fresh-process runs; do not imply that operating-system caches are cold.

### Correct live collection

- [ ] Make snapshot-and-reset one atomic operation. Define windows by measurement completion; report active spans crossing the boundary.
- [ ] Exclude profile inspection requests from route totals.
- [ ] Separate aggregate reset from deep-profiler lifecycle; never discard an active profiler during reset.
- [ ] Count encoded bytes for text streams, record completion/error/disconnect outcomes, and preserve original exceptions during cleanup.
- [ ] Report capacity overflow instead of silently dropping labels.
- [ ] Rename the displayed “cold hit” to “first observation”: resetting counters does not clear caches.

Acceptance: failed or incomplete work cannot appear as a speedup, comparisons
preserve precision, and concurrent recording cannot disappear between snapshot
and reset.

## Phase 2 — Benchmark realistic browser workloads

Extend `tests/ui/shot.py` and reuse its browser, fixture, and page-readiness helpers.

### Reproducible scenarios

| Scenario | Workload and actions |
| --- | --- |
| Studio | 200 observations × 12 participants; load, filter, restore, queue selection |
| Transcripts | 2,400 segments; load, search, switch participant, restore |
| Screenspace | 2,000 synthetic events with matching media; load, filter results, switch views |
| Idle behavior | Sample each of the six app pages while visible, then exercise its existing hidden-page lifecycle |

- [ ] Generate fixtures centrally and assert loaded data counts before measuring. Validate rendered or visible counts according to each page's rendering strategy.
- [ ] Use readiness and action-completion conditions instead of fixed sleeps to declare success.
- [ ] Capture initial load separately from repeated interactions.
- [ ] Save frontend measurements, backend window snapshots, workload metadata, screenshots, and errors in one run artifact.
- [ ] Add `--perf-output PATH` to save the browser report without scraping stdout.
- [ ] Replace the two-point soak comparison with samples every five seconds for a default 60-second observation window. Report trajectories and rates for heap, DOM nodes, listeners, polling, and transfer volume.
- [ ] Treat sustained growth as a diagnostic signal, not proof of a leak.
- [ ] Record unsupported browser metrics explicitly rather than presenting them as zero.
- [ ] Make frontend timing safe for nested same-label spans and keep recording bounded and disabled when profiling is off.

Acceptance: each scenario proves that its intended workload and actions occurred;
reports distinguish load cost, interaction cost, and idle churn.

## Phase 3 — Explain overlapping work

- [ ] Add bounded operation records for clip batches, Screenspace tasks, transcription tasks, and workflow runs.
- [ ] Record operation ID, parent ID, kind, queued/start/end times, outcome, work count, and accumulated measurements.
- [ ] Pass operation context explicitly into worker submissions; do not assume request context follows threads.
- [ ] Keep existing aggregate labels stable. Store task IDs as operation fields rather than creating a label per task.
- [ ] Separate queue wait, execution duration, and overall elapsed time. Clearly identify inclusive timings that overlap.
- [ ] Retain the latest 100 completed operations plus active operations; expose eviction counts.
- [ ] Include operation records in JSON and `/api/profile`, allowing inspection of running jobs without waiting for process exit.
- [ ] Add downloadable deep-profile files for completed profiles so investigation is not limited to the printed top 15 functions.

Defer continuous CPU/process-tree memory sampling and a dedicated dashboard. Both
can build on these artifacts later; neither is needed for the first diagnostic
improvements.

## Validation and rollout

- [ ] Test atomic reset under concurrent recording, active deep profiles, Unicode streams, disconnects, capacity limits, and disabled-mode behavior.
- [ ] Test failed subprocesses with partial metrics, missing outputs, mismatched fixtures, missing baseline scenarios, repeated-run aggregation, and threshold boundaries.
- [ ] Test nested browser spans, rejected promises, unsupported observers, scenario readiness failures, and bounded collection.
- [ ] Run existing profiling and benchmark tests, relevant frontend checks, and six-page UI smoke; inspect screenshots.
- [ ] Measure instrumentation overhead using identical workloads with profiling off, ordinary profiling, and deep profiling separately.
- [ ] Keep timing regression gates opt-in; ordinary CI should enforce correctness and workload validity.
- [ ] Update the profiling skill and performance guidance with canonical commands, interpretation rules, and current labels, including `llm.generate`.

Defaults: developer-facing, local artifacts, no telemetry, no new heavy dependencies
or automatic browser/model downloads. Preserve opt-in instrumentation and avoid
per-frame recording overhead.

## Implementation notes

No implementation or runtime validation has been performed for this plan.
Record completed work, verification results, deviations, and unresolved issues
here as each phase lands.
