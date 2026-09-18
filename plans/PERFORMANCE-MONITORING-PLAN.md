# More reliable performance diagnostics

Status: In progress; Phases 1-2 landed, Phase 3 pending.

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
| 1. Trustworthy measurements | Done | Reliable structured exports and benchmark comparisons |
| 2. Browser workloads | Done | Verified workloads with load, interaction, and idle reports |
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

- [x] Add `--profile-output PATH`, implying `--profile`, to write full-precision JSON alongside the existing console report.
- [x] Include metrics, startup marks, profiling mode, elapsed time, dropped-label counts, and environment metadata.
- [x] Change benchmark runners to consume JSON directly; remove their console-parsing dependency.
- [x] Record commit and dirty state, Python/platform, relevant dependency and binary versions, effective tuning settings, and fixture identity.
- [x] Treat deep-profile runs as diagnostic artifacts; reject them from ordinary timing comparisons.

### Make benchmark success explicit

- [x] Require successful process exit, required measurements, expected work counts, and completed outputs.
- [x] Invalidate the scenario when any repetition fails; a successful repetition must not hide failure.
- [x] Distinguish invalid runs, incomparable workloads, and performance regressions in reports and exit status.
- [x] Validate fixture parameters and contents, including dimensions, frame rate, sampling interval, regions, and spreadsheet geometry. Duration alone is insufficient.
- [x] Reject workload mismatches. Report environment differences and require an explicit benchmark override to compare across them.
- [x] Update persisted result shapes directly, without legacy readers or migrations.

### Retain repeated measurements

- [x] Default to three measured runs and retain every sample. Report median, minimum, maximum, and median absolute deviation.
- [x] Compare medians; keep the minimum as supplementary information.
- [x] Measure subprocess elapsed time as the primary end-to-end metric. Preserve callback, pool, encoding, and post-processing timings as diagnostic breakdowns.
- [x] Retain `--fail-on` as the percentage threshold; add metric selection and an optional absolute threshold. When both thresholds are supplied, require both to be exceeded.
- [x] Describe runs as fresh-process runs; do not imply that operating-system caches are cold.

### Correct live collection

- [x] Make snapshot-and-reset one atomic operation. Define windows by measurement completion; report active spans crossing the boundary.
- [x] Exclude profile inspection requests from route totals.
- [x] Separate aggregate reset from deep-profiler lifecycle; never discard an active profiler during reset.
- [x] Count encoded bytes for text streams, record completion/error/disconnect outcomes, and preserve original exceptions during cleanup.
- [x] Report capacity overflow instead of silently dropping labels.
- [x] Rename the displayed “cold hit” to “first observation”: resetting counters does not clear caches.

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

- [x] Generate fixtures centrally and assert loaded data counts before measuring. Validate rendered or visible counts according to each page's rendering strategy.
- [x] Use readiness and action-completion conditions instead of fixed sleeps to declare success.
- [x] Capture initial load separately from repeated interactions.
- [x] Save frontend measurements, backend window snapshots, workload metadata, screenshots, and errors in one run artifact.
- [x] Add `--perf-output PATH` to save the browser report without scraping stdout.
- [x] Replace the two-point soak comparison with samples every five seconds for a default 60-second observation window. Report trajectories and rates for heap, DOM nodes, listeners, polling, and transfer volume.
- [x] Treat sustained growth as a diagnostic signal, not proof of a leak.
- [x] Record unsupported browser metrics explicitly rather than presenting them as zero.
- [x] Make frontend timing safe for nested same-label spans and keep recording bounded and disabled when profiling is off.

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

### Phase 1 (landed)

- `--profile-output PATH` writes `profiling.export()`: labels, startup marks,
  mode, window seconds, dropped-label count, active spans, and `env` (commit +
  dirty, Python, platform, `hardware.profile()`, ffmpeg version line, package
  versions, the PERFORMANCE.md tuning knobs). `/api/profile` returns the same
  document; `?reset=1` is one atomic snapshot-and-clear, and the label map moved
  from `profile` to `labels`.
- `reset()` no longer discards deep profilers (`deep_reset()` exists for tests).
  `stream_span(body, rule=)` counts UTF-8 bytes and records
  `stream.complete|error|disconnect <rule>`; a close-time error never masks the
  body's own exception. Profile inspection requests are excluded from `route`
  totals. Labels refused at the cap are counted and printed as `labels_dropped`.
- `tests/perf/bench_common.py` holds the shared run/aggregate/compare logic;
  both benches read `profile.json`, default to 3 fresh-process runs, keep every
  sample, compare medians, and exit 3 (invalid) / 4 (incomparable) / 1
  (regression). Fixture identity covers duration, size, fps and audio; workload
  and environment mismatches void a compare (`--allow-env-mismatch` for the
  latter). Deep runs refuse `--save`/`--compare`.
- Verified: full pytest, ruff, ty; `scan_bench.py --tools color,change
  --duration 15 --runs 2 --save` and `clip_bench.py --duration 15 --runs 2
  --save` both produce valid rows and JSON; `--compare` against those baselines
  passes.
- Found while validating: on this machine's ffmpeg (no `drawtext`) clipgen
  disables titlecards for the run, so `clips-cards` and `reel` had been
  reporting card-free timings as "carded". The bench now marks those rows
  `invalid` (`0 clips carded, expected 20`) instead of comparing them.
- Deferred to Phase 3: per-operation records in the export.

### Phase 2 (landed)

- `clipgenPerf` (`assets/web/utils.js`) times spans from a per-label stack, so
  nested same-label spans each record; labels cap at 512 with a `dropped`
  count; `record` is gated like the rest; `snapshot().supported.longtask`
  is `false` on a browser without the observer (printed as `unsupported`,
  never 0). Node behaviour test: `tests/test_clipgen_perf_js.py`.
- `tests/perf/bench_fixtures.py` is the single source of benchmark inputs
  (sheet geometry, testsrc video, synthetic transcripts/screenspace
  manifests, count helpers); `clip_bench`, `scan_bench`, `ui_bench` and
  `test_shot` all build from it.
- `shot.py`: `--perf-output PATH`; `--soak` samples every `--soak-interval`
  (5 s) and reports per-minute slopes for heap, nodes, listeners, transfer,
  longtasks and each poll label, split visible/hidden with `--soak-hidden`
  (`_ui_pages.set_hidden` emulates a background tab); the resource-timing
  buffer is raised to 10,000 so transfer volume is not truncated.
- `tests/ui/ui_bench.py`: `studio` / `transcripts` / `screenspace` / `idle`
  scenarios with asserted workload counts, condition-based actions, per-run
  artifacts (load, actions with before/after counts and a server profile
  window each, errors, screenshot), and the shared compare/exit-code rules.
- Verified (this machine, headless shell, one run each): studio 200 rows
  loaded, filter 15 ms → 50 rows, restore 33 ms, queue 5 ms; transcripts
  2400 rows, search 354 ms (300 ms debounce inside), switch 65 ms → 100
  rows, restore 82 ms; screenspace open-results 130 ms (120-row first
  chunk), drain 869 ms → 2000 rows, filter 139 ms → 1142 rows. Idle (12 s
  window, 3 s interval, hidden second half): every poller reads 0/min while
  hidden (`studio.*Intake`, `workflows.discover` at 10/min visible);
  transcripts listeners grew ~6/min visible — a signal to look at, not a
  finding of this plan.
- Not done: a Composer poller check (Composer has no poller); browser
  scenarios are opt-in (`CLIPGEN_UI_CHECK=1`) and never run in CI.
