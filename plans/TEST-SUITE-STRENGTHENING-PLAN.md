# Strengthen test behavior, output verification, and concurrency checks

## Summary

Address all five review findings through five independently assignable work packages. Preserve the existing pytest, Node, and Playwright harnesses.

Baseline: **4,328 passed, 1 skipped in 46.06 seconds**, without coverage or browser tests. Re-confirmed 2026-09-23 with `uv run --extra dev pytest -c tests/pytest.ini -p no:randomly -q`; the slowest test is 1.66 s.

## Findings addressed

The plan answers a five-point review of the suite (the review itself is not committed):

1. Studio cancellation and streaming are asserted as source strings, which pass while behavior breaks → package 1.
2. The Studio generate journey checks "Done" and success classes but never the produced file → package 2.
3. `test_pipeline_io.py` accepts 1–2 s for a 1 s cut, so an unchanged source passes; screenshots are never decoded → package 3.
4. The endcard test sleeps to arrange contention and accepts one collected result; two other races are sleep-arranged → package 4.
5. CI relies on runner-provided Node and ffmpeg while the dependent tests skip silently when they are absent → package 5.

Verified against the tree on 2026-09-23: every file, function, and route named below exists with the described shape.

The work changes tests and CI only. No public APIs, persisted schemas, production behavior, dependencies, version bump, or changelog changes are planned. If a stronger test exposes a product defect, report it separately rather than weakening the assertion.

## Status

| Package | Owner | Status |
| --- | --- | --- |
| Frontend streaming and cancellation | A | Not started |
| Studio artifact journey | B | Not started |
| Real-media assertions | C | Not started |
| Concurrency tests | D | Not started |
| CI prerequisites | E | Not started |
| Integration and validation | Integrating agent | Not started |

Update each package's status as it lands, including resolved issues or changed scope.

## Follow-up candidates (not scheduled)

Recorded so the review's "apply the same pattern" remark is not lost. Henrik decides whether any becomes a package.

- **Sleep-arranged coordination elsewhere.** Non-UI tests hold 49 `time.sleep` calls across 15 files; package 4 touches three. The concentrations are `tests/screenspace/test_worker.py` (16), `tests/test_profiling.py` (7), `tests/test_workflows_api.py` (6), and `tests/test_transcripts_api.py` (4). Many are legitimate pacing, so this is an audit, not a sweep.
- **Other source-string frontend tests.** `tests/test_studio_frontend_source.py` has around twenty more string-assertion tests beyond the two package 1 executes. Each is a candidate for the same `node_eval` treatment where the function is pure enough to slice.

## 1. Execute frontend streaming and cancellation behavior

**Owner A**

**Files:** `tests/test_js_units.py`, `tests/test_studio_frontend_source.py`, `tests/_frontend_source.py`.

Use the existing `slice_between()` and `node_eval()` approach to execute shipped JavaScript. Do not duplicate its implementation or introduce another JavaScript runner.

### Streaming cases

Extract `readNDJSONStream` and `apiPostNDJSON` from `assets/web/utils.js` with `slice_between(src, "// ---- NDJSON streaming reader ----", "// ---- File downloads ----")`; both helpers sit in that one block.

Build a fake response reader that yields byte chunks through promises. Use Node's real `TextEncoder` and `TextDecoder`. `fetch` is stubbed in the prelude; `AbortController`, `TextDecoder`, and `Promise` are Node globals.

`node_eval()` today runs with `check=True`, no timeout, and returns stdout lines. A rejected promise exits Node nonzero (a clear failure), but an *unresolved* promise exits zero with empty stdout, so the caller's `[0]` raises `IndexError`. That is the gap the harness items below close.

Assert:

- Several records in one chunk produce separate callbacks in order.
- A record split across chunks produces exactly one complete callback.
- A multibyte Unicode character split across chunks survives unchanged.
- Empty lines are ignored; CRLF and surrounding whitespace follow the current trimming contract.
- A final record without a newline is delivered.
- An empty stream resolves without callbacks.
- A missing streaming body rejects.
- A rejected `read()` propagates its error.
- An exception from `onLine` rejects the returned promise.
- The reader delivers strings; it does not parse JSON itself.

For `apiPostNDJSON`, assert:

- URL, POST method, JSON body, content type, and supplied abort signal reach `fetch`.
- Successful responses flow through the real streaming reader.
- HTTP errors preserve `status` and `bodyText`.
- A failed error-body read still rejects with the HTTP status and empty body text.
- Fetch and reader `AbortError` objects propagate unchanged.

### Studio cancellation cases

Execute the real `onCancelGenerate` and `isGenerateFetchAborted` functions from `studio-generate.js`, with minimal state, DOM, and API stubs. The two functions are not adjacent (`isGenerateFetchAborted` near line 76, `onCancelGenerate` near line 362), so slice each one separately; the file is one IIFE, and the slices need `state`, `qs`, and `apiPost` stubbed in the prelude.

Assert:

- Every active controller receives `abort()`.
- One throwing controller does not prevent later controllers from aborting.
- Cancellation sets the user-cancelled flag, hides the button, clears the controller list, and posts both cancellation endpoints.
- Empty or absent controller lists are safe.
- Cancellation-endpoint failures are consumed.
- An `AbortError`, explicit user cancellation, and an ordinary error are classified correctly.

Use this coverage to remove redundant source-string assertions in `test_on_cancel_generate_aborts_and_posts_intake_cancel`. Retain architectural checks and assertions about branches not actually executed by these new tests.

### Harness and acceptance

- Bound Node subprocesses with a timeout.
- Make asynchronous probes explicitly report failures and exit unsuccessfully; unresolved promises must not produce a passing result.
- Batch related cases into a small number of Node processes.
- Preserve local skip behavior when Node is absent.
- Demonstrate that removing a cancellation endpoint or dropping the final buffered line makes the relevant test fail. Restore temporary mutations before handoff.

## 2. Verify the Studio journey's generated artifact

**Owner B**

**File:** `tests/ui/test_ui_journeys.py`, existing `test_studio_generate_produces_an_artifact`.

Keep the existing click sequence, DOM checks, screenshot capture, and fatal-browser-error assertion. Extend this journey rather than adding a seventh journey.

### Implementation

1. Before selecting the cell, GET `/studio/api/manifest` and record existing artifact IDs.
2. Run the existing row-6, P01 generation interaction.
3. After the completion UI appears, poll the manifest within the existing generation timeout.
4. Require exactly one new clip artifact for P01 and the selected spreadsheet row. Match on the record's own keys: `type == "clip"`, `participant == "P01"`, `cellRow == 6` (the grid's `data-row` is the 1-based sheet row, and the artifact id is `a6c<col>s0`).
5. Resolve its `file` beneath `_ui_fixtures.OUTPUT_DIR`; assert the path remains inside that directory.
6. Assert the file exists and contains bytes.
7. GET `/studio/media/<URL-encoded file>` through the live server.
8. Require HTTP 200 and response bytes identical to the file.

Use the returned artifact filename, not an independently reconstructed naming convention. Read the manifest endpoint without posting an explicit save: generation itself must persist the artifact. Today it does: the generate stream calls `_save_manifest_quiet()` after every batch, unconditionally, and `GET /studio/api/manifest` returns `{ok, artifacts, reels}` from `viewer.load_manifest_both()`. The live server's `config.OUTPUT_DIR` is `_ui_fixtures.OUTPUT_DIR`, and `/studio/media/` serves it through `send_from_directory`.

Use bounded condition polling, not a fixed sleep. Do not add decoding here; package 3 owns media correctness.

### Acceptance

- The journey fails when success UI appears without a new manifest artifact.
- It fails for missing, empty, or unservable files.
- It passes independently and within the complete UI suite.
- Inspect `.context/ui-check/screenshots/journey-generate.png`.
- Keep all browser opt-in gates and the current six-journey count.

## 3. Make real-media assertions reject incorrect output

**Owner C**

**File:** `tests/test_pipeline_io.py`.

Strengthen the existing four tests without increasing their count.

### Deterministic source

Replace the shared two-second `testsrc` fixture with a three-second video containing:

- One second of solid red, followed by two seconds of solid blue.
- 160×120 resolution, 15 fps, H.264, `yuv420p`.
- A keyframe at each second and no B-frames: `-g 15 -bf 0 -x264-params keyint=15:min-keyint=15:scenecut=0`.
- Built from two lavfi `color=` inputs joined with the `concat` filter, or an equivalent single-command form.

Why three seconds, not two: screenshots key off the start time only, so a `0:01` screenshot samples the frame at exactly t=1.0, the first blue frame. The cut is deterministic but has zero margin. With blue running to 3 s the screenshot can sample `0:02`, a full second from either color boundary, while the clip and reel windows below stay within the first two seconds. The extra second of 160×120 encode is negligible.

Explicitly pin, via `monkeypatch` on `config`: `REENCODING = False` (stream copy), `AUDIO_NORMALIZE = False`, `DEBUGGING = False`, `CLIP_PARALLEL_WORKERS = 1` (already set), and pass `titlecards_enabled=False` as the tests already do.

Leave the separate five-second frame-seek fixture unchanged.

### Clip and screenshot assertions

- Cut `0:01–0:02`; the expected content is blue.
- Read precise video-stream duration using ffprobe JSON, not the rounded production duration helper.
- Require one second within one frame's tolerance.
- Decode the complete clip through ffmpeg with decoding errors treated as failures.
- Require decoded frames and verify a central sample is blue.
- For the screenshot, use `0:02` as the start and call `image.load()` before checking dimensions and blue content.

Use RGB channel dominance to classify colors: the expected channel must exceed 150, and each other channel must remain below 100. Avoid encoded-byte comparisons and exact pixel equality.

### Reel assertions

Keep the existing failed-second-cut branch and its zero-artifact/cleanup assertions.

For the successful branch:

- Request the blue window `0:01–0:02` first, then the red window `0:00–0:01`.
- Require approximately two seconds, allowing at most two frames of tolerance for the combined output. The reel path is the concat demuxer over stream-copied parts (`video.concatenate_clips`), so keyframe-aligned parts make this tolerance realistic.
- Decode the complete reel successfully.
- Verify a middle frame from the first second is blue and one from the second second is red.

This must reject an unchanged source, reversed ordering, duplicated segments, and a missing segment.

### Acceptance

- Keep ffmpeg/ffprobe skips for local environments without those tools.
- Bound every new subprocess.
- Share probe/decode helpers only where they have multiple callers.
- Measure the changed tests twice following `agents/skills/test-perf/SKILL.md`.
- Verify that substituting the source for the cut, or reversing the reel inputs, fails the assertions.
- Keep temporary outputs beneath pytest's temporary directories.

## 4. Make concurrency windows explicit and cleanup reliable

**Owner D**

**Files:** `tests/test_titlecards.py`, `tests/test_video_commands.py`, `tests/test_clip_pipeline.py`.

Limit this package to the three identified tests. Do not sweep unrelated sleeps or redesign production synchronization.

### Endcard cache contention

Update `test_get_or_build_endcard_builds_once_under_parallel_workers`.

- Hold the first builder on an event.
- Instrument the per-key flight-lock entry with a test-only context manager that delegates to a real lock.
- Signal immediately before each competing acquisition.
- Release the builder only after all four callers have reached the acquisition boundary.
- Collect exceptions explicitly.
- Require one build, four results, identical nonempty paths, and no live workers.

Supply the instrumented lock through the existing flight lookup without patching `threading.Lock` globally. The lookup is `_endcard_flights.setdefault(cache_key, threading.Lock())` under `_endcard_lock`, so the seam is `monkeypatch.setattr(titlecards, "_endcard_flights", <dict subclass whose setdefault returns the instrumented lock>)`. Do not reconstruct `cache_key` in the test; it folds in config values.

### Audio extraction contention

Update `test_extract_audio_track_concurrent_calls_run_ffmpeg_once`.

- Wrap the existing `_audio_track_lock` accessor (a function returning the per-key lock, used as `with _audio_track_lock(resolved, idx):`) with a context manager that records acquisition attempts and delegates to its original lock.
- Hold the first mocked ffmpeg call until the second caller reaches lock acquisition.
- Require one ffmpeg call, two results, identical paths, expected bytes, and no worker errors.
- Preserve the existing temporary-filename assertions.

### Pipeline cancellation

Update `test_run_clip_pipeline_cancel_captures_started_clip_results`.

The production loop (`_parallel_map_ordered`) checks cancellation once per completed future, then cancels pending futures and breaks; the pool's `__exit__` waits for running futures and `_drain_ran_futures` collects them. The serial branch's per-clip check is not on this path. Arrange the test accordingly:

- Use two workers and two clips.
- Both workers signal that they started.
- The first waits until the second has started, then completes.
- The second remains blocked.
- The cancellation callback asserts that the second worker is still running, records that cancellation was observed, releases that worker, and returns `True`. Make it safe to call more than once; only the first call asserts.
- Assert cancellation was observed and both started clips' exact result paths were recovered.

Run the pipeline in a controlled thread so the test can release gates and perform cleanup if orchestration fails.

### Cleanup and acceptance

For all three tests:

- Put gate release and worker joins in `finally`.
- Assert waits succeeded and every worker exited.
- Surface worker exceptions on the test thread.
- Use bounded waits as failure limits, never as normal coordination.
- Do not leave workers running after fixture teardown.

Run the affected tests eight times and the shuffled parallel suite three times, following the performance skill. Confirm that bypassing duplicate-work protection or removing post-cancellation result draining makes the relevant test fail. Record before/after timings.

## 5. Fail CI when required test tools are missing

**Owner E**

**File:** `.github/workflows/tests.yml`.

Add a prerequisite step in the `tests` job between "Install dependencies" and "Run tests with coverage gate" (no such step exists today; `ubuntu-latest` ships `node`, `ffmpeg`, and `ffprobe`):

- Require `node`, `ffmpeg`, and `ffprobe` on PATH.
- Execute each binary's version command so unusable executables also fail.
- Print concise version information for diagnosis.
- Fail immediately with the missing tool's name.

Use the runner's installed tools. Do not add downloads, package installation, browser setup, or a version matrix in this package.

Add `-ra` to the existing pytest command so skipped-test reasons appear in CI. It composes with `-n auto`; the non-UI suite carries 17 skip markers, so the report stays short.

Preserve:

- The full-suite run, coverage configuration, fixed random seed, and xdist execution.
- Local skip behavior.
- Optional model-test skips.
- Browser exclusion from CI.

### Acceptance

Run the exact prerequisite script in the normal environment. Then execute it with a temporary PATH containing tool stubs, omitting each required tool in turn. Each omission must return nonzero and identify the tool.

Do not add a source-string test that merely checks the workflow's spelling.

## Integration, validation, and handoff

The five packages can proceed independently with the ownership above. Owner A exclusively owns changes to shared Node helpers. Other owners should avoid editing those files.

Each handoff must include changed files, focused test results, temporary mutation checks performed, timings where required, and remaining limitations.

The integrating agent must:

1. Review all changes against `agents/CODE-REVIEW.md`.
2. Run the format, lint, type, and full-test stages from `agents/skills/check/SKILL.md`.
3. Complete the repeated timing and concurrency checks required by the performance skill.
4. Run the full opt-in UI suite and inspect the Studio journey screenshot.
5. Confirm default pytest collection still excludes browser tests, the UI journey count remains six, and the pipeline I/O count remains four.
6. Report unexpected skips separately from passes.

Use existing installed tools. If Playwright or Chromium is missing, request installation permission as required by the UI skill; a skipped browser run does not satisfy acceptance.

Update this plan's status as each package lands, including resolved issues or changed scope.
