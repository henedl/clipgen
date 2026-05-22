# Streaming and Concurrency Fixes

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix the remaining streaming, cancellation, and concurrency bugs found in Studio generation plus shared media/manifest paths.

**Architecture:** Keep the Flask server thin but make shared mutation boundaries explicit: reserve output filenames atomically, serialize manifest read-modify-write cycles, and propagate cancellation through every long-running ffmpeg/model path. Keep Studio's independent artifact/reel UX, but make each stream branch report failures and cancellation consistently.

**Tech Stack:** Python Flask, `threading.Lock`, `ThreadPoolExecutor`, ffmpeg subprocess helpers, vanilla JavaScript `fetch`/NDJSON readers, existing `uv run --extra dev pytest -c tests/pytest.ini` test stack.

---

## Scope

This plan covers issues found in the recent streaming/concurrency commits plus a second pass through older shared paths. It intentionally avoids visual redesign and broad frontend refactors, except for a small NDJSON helper if it keeps the error/cancel paths consistent.

Primary files:

- Modify: `files.py`
- Modify: `viewer.py`
- Modify: `server.py`
- Modify: `pipeline.py`
- Modify: `titlecards.py`
- Modify: `assets/web/studio.js`
- Modify: `screenspace_server.py`
- Modify: `transcripts.py`
- Test: `tests/test_clip_pipeline.py`
- Test: `tests/test_manifest.py`
- Test: `tests/test_studio_api.py`
- Test: `tests/test_studio_frontend_source.py`
- Test: `tests/test_screenspace_api.py`
- Test: `tests/test_transcripts_api.py`

## Findings

### P0: Output Path Reservation Race

`files.get_unique_filename()` is check-then-use. Parallel clip, reel, and gallery workers can select the same path before either ffmpeg process creates it.

Impact:

- Parallel Studio `/api/generate` can overwrite artifacts when two cells share the same filename template.
- Parallel reel part generation can clobber `_reel_part_...` outputs.
- Parallel gallery GIF extraction has the same risk.

Fix:

- Replace check-only uniqueness with an atomic reservation helper, using `os.open(..., O_CREAT | O_EXCL)` for placeholder files.
- Keep the returned path reserved until ffmpeg overwrites it.
- Ensure callers that fail before ffmpeg starts remove the placeholder.

Tests:

- Add a threaded test where many workers request the same filename and assert all returned paths differ and exist as reservations.
- Add a process failure test that verifies an unused reservation is cleaned up.

### P0: Manifest Save Race

> **Status:** Done — branch `claude/streaming-fixes-plan-6ConI`. Added
> `_MANIFEST_WRITE_LOCK` in `viewer.py` wrapping the full load-merge-write
> cycle of `save_manifest()`; regression test in `tests/test_manifest.py`.

`server._save_manifest_quiet()` snapshots `_generated_artifacts` / `_generated_reels` under `_generated_output_lock`, then releases it before `viewer.save_manifest()` performs load-merge-save. Concurrent Studio artifact/reel/intake completions can each read the same old manifest and last-writer-wins a partial merge.

Fix:

- Add a manifest write lock in `viewer.py` that wraps the full `save_manifest()` load-merge-write operation.
- Keep `_MANIFEST_CACHE_LOCK` focused on cache internals; avoid holding it while writing unless needed.
- Call `_reset_manifest_cache()` after successful writes as today.

Tests:

- Add a concurrent `viewer.save_manifest()` test with artifacts and reels written from multiple threads.
- Assert the final manifest includes every id.

### P0: Endcard Cache Deleted During Parallel Generate

`pipeline.process_clips()` clears the global titlecard endcard cache at the end of each call. Studio now calls `process_clips([clip])` concurrently per cell, so one worker can delete cached endcard temp files while another worker is still wrapping a clip.

Fix:

- Add a `clear_titlecard_cache: bool = True` keyword to `process_clips()`.
- Pass `clear_titlecard_cache=False` from Studio per-cell workers.
- Clear once in the request generator `finally` after all cell futures have completed or been cancelled.
- Keep CLI/batch behavior clearing once per top-level `process_clips()` call.

Tests:

- Add a regression test that runs two `process_clips()` calls in parallel with titlecards enabled and asserts `clear_endcard_cache()` is not called from worker-scoped calls when disabled.
- Add a Studio API test that consumes `/api/generate` and asserts the cache clear happens once after the stream.

### P1: Reel Per-Request Titlecards and Cancellation Are Not Wired to Segment Workers

`pipeline.process_reel()` accepts `cancel_flag`, `titlecards_enabled`, and `titlecard_duration_seconds`, but `process_reel_clip()` does not pass those values into `_process_single_clip_segments()`.

Impact:

- `POST /studio/api/reel/cancel` does not promptly terminate in-flight ffmpeg part generation.
- Reel transcript offsets can be computed with per-request titlecard settings while the actual reel parts use global config.

Fix:

- Pass `cancel_flag`, `titlecards_enabled`, and `titlecard_duration_seconds` into `_process_single_clip_segments()` inside `process_reel_clip()`.
- Keep `_build_reel_transcript()` using the same resolved options.

Tests:

- Add `test_process_reel_forwards_cancel_and_titlecard_options_to_segments`.
- Add a cancellation test where a fake segment function sees the same cancel callable supplied to `process_reel()`.

### P1: Direct Reel Route Bypasses Titlecard Settings

Mixed reels that include any intake item route through `server.api_reel_direct()` instead of the spreadsheet-backed `server.api_reel()` path. The UI sends `titlecards_enabled` and `titlecard_duration` to both endpoints, but `api_reel_direct()` currently cuts raw segments and concatenates them without parsing or applying those per-request titlecard options.

Impact:

- Intake-only and mixed intake/sheet reels silently omit titlecards even when the Studio checkbox is on.
- The two reel paths no longer share the same visible media contract.

Fix:

- Parse titlecard options in `api_reel_direct()` with the same `_parse_titlecard_request()` helper used by `api_reel()`.
- Apply titlecard wrapping to direct reel parts, or explicitly route direct reel segment generation through the same segment helper used by `pipeline.process_reel()` once it accepts already-expanded start/end seconds.
- Keep baseline handling client-side for `api_reel_direct()`; do not re-run spreadsheet parsing for intake-derived items.

Tests:

- Add a Studio API test that posts to `/studio/api/reel-direct` with titlecards enabled and asserts the segment builder/wrapper receives `titlecards_enabled=True` and the requested duration.
- Add a negative test that titlecards disabled on `/studio/api/reel-direct` produces raw segment behavior.

### P1: Intake Generation Has No Cancellation Contract

Studio mixed artifact generation runs sheet cells through `/api/generate` and intake cards through `/api/generate-intake`. The Cancel button only sets `_generate_cancel_event`; intake fetches and ffmpeg calls continue.

Fix:

- In `assets/web/studio.js`, create an `AbortController` for each branch of `onGenerate()`.
- Make `onCancelGenerate()` abort the active intake request and mark pending intake cards as cancelled/cleared.
- In `server.py`, thread `_generate_cancel_event.is_set` into `_process_intake_item()` and then into `video.run_ffmpeg()`.
- Consider claiming the generate busy slot for intake-only requests, or add an `_intake_in_progress` slot if sheet and intake must remain independently cancellable.

Tests:

- Add a Studio frontend source test that asserts `AbortController` is used for generate intake fetches.
- Add an API test where `_generate_cancel_event` is set before an intake item and no ffmpeg call is made.
- Add an API test where cancellation during `video.run_ffmpeg()` yields an `ok: false, cancelled: true` NDJSON line.

### P1: Generate Cancel Does Not Stop Running ffmpeg Workers

`server.api_generate()` breaks out of the `as_completed()` loop on cancellation and calls `future.cancel()`. That only cancels queued futures; already-running `pipeline.process_clips()` calls continue their ffmpeg subprocesses, can leave media files on disk, and are not yielded or persisted.

Fix:

- Pass `_generate_cancel_event.is_set` into every `pipeline.process_clips()` call from the Studio generate worker path.
- Ensure `_process_single_clip_segments()` forwards that cancel flag into `video.run_ffmpeg_process()` / `video.run_ffmpeg()` calls for extraction, titlecard wrapping, and post-processing.
- On cancel, continue draining completed futures long enough to terminate running subprocesses cleanly, but do not append new artifacts after cancellation.
- Clean up incomplete reserved output paths when cancellation prevents manifest append.

Tests:

- Add a Studio API test where cancellation during a fake long-running ffmpeg call triggers the cancel flag in the worker.
- Add a pipeline test that `process_clips(cancel_flag=...)` passes cancellation to the video command layer and does not return artifacts for cancelled clips.

### P1: Streaming Branch Errors Leave Bad UI State

Sheet generation fetch failures only call `finishBranch()`, leaving cards visually queued and often reporting "Generated 0 artifacts". Card cleanup also queries wrong class names: `.queue-card.queued` and `.card-result-badge`, while rendered classes are `.queue-card-queued` and `.card-gen-badge`.

Fix:

- Extract a small Studio helper to read NDJSON streams with a body guard and shared error handling.
- On sheet fetch failure, mark captured sheet cards failed and increment `totalFail`.
- Treat `totalSuccess === 0 && totalFail === 0 && !cancelled` as an error.
- Fix selectors in `clearCardStatus()` and cancellation cleanup to use `.card-gen-badge` and `.queue-card-queued`.
- Handle HTTP 409 in reel builds as a JSON error response rather than feeding it through the stream reader.

Tests:

- Add frontend source tests for the corrected selectors.
- Add a source test that the sheet branch catch marks failures.
- Add API/front-end source coverage for 409 handling in reel streaming.

### P1: Streamed Work Persists Only at Stream End

`/api/generate`, `/api/generate-intake`, and reel streams append completed records in memory, but persist the manifest only after the generator reaches its end. Client disconnects, server restarts, or generator errors can leave generated files on disk without manifest records.

Fix:

- Call `_save_manifest_quiet()` from generator `finally` blocks whenever any record was appended.
- For long batches, save after each completed cell/item or after a small batch interval.
- Combine this with the manifest write lock so increased save frequency is safe.

Tests:

- Add generator-close tests that consume only the first NDJSON line, close the response iterator, and assert `_save_manifest_quiet()` was still called.
- Add a test that a successful early item is persisted even when a later item raises.

### P2: Reel Prep Shares Mutable Fuzzy-Match State Across Threads

`process_reel()` uses `_run_clip_pipeline(parallel=True)`, and each worker calls `_prepare_and_check_clip()` with shared `missing_videos` and `fuzzy_matches`. Those structures are mutated without synchronization.

Fix:

- Move reel preparation into a sequential phase, mirroring `process_clips()`, then parallelize only ffmpeg segment generation.
- Alternatively, protect `fuzzy_matches` and `missing_videos` with a lock and keep all user-input fuzzy prompts out of worker threads.
- For server-side paths, ensure `utils.NO_INPUT_MODE` behavior prevents worker prompts.

Tests:

- Add a test where multiple reel clips miss the same video and assert only one missing-video record is produced.
- Add a test where parallel reel generation does not call `utils.read_user_input()`.

### P2: Studio Thumbnail Cache Is Not Locked

Studio `_thumbnail_cache` is an `OrderedDict` with `move_to_end()` and eviction but no lock. Screenspace already locks its frame cache.

Fix:

- Add `_thumbnail_cache_lock`.
- Guard get, `move_to_end`, insert, and eviction in `api_thumbnail()` / `_thumbnail_cache_put()`.

Tests:

- Add a threaded thumbnail cache test similar to the existing cache tests.

### P2: Intake Artifact IDs Collide on Identical Spans

`_process_intake_item()` hashes only `{participant}_{start}_{end}` for filename/id. Two intake events covering the same span produce identical artifact ids; manifest dedupe silently replaces one.

Fix:

- Include stable source metadata in the id hash: `source`, `event_ids`, `mark_ids`, and the item index when available.
- Thread the request index into `_process_intake_item()`.

Tests:

- Add an intake stream test with two same-span items and distinct event ids; assert artifact ids differ.

### P2: Sheet Switch Can Rebind Generated Lists During Active Work

`_init_studio_state()` and sheet-open state swaps replace `_generated_artifacts` / `_generated_reels` while generation streams may append under `_generated_output_lock`.

Fix:

- Reject spreadsheet switching while `_generate_in_progress`, `_reel_in_progress`, or intake generation is active.
- Hold `_generated_output_lock` when rebinding generated lists.

Tests:

- Add an API test that sheet switch/open returns 409 while generation is active.
- Add a unit test that `_init_studio_state()` rebinds generated lists under the lock.

### P2: Older Shared Read Paths Still Bypass Manifest Locks

Second-pass review found several read endpoints returning or filtering `_manifest` without `_manifest_lock`, while write endpoints mutate it under lock. Examples include Screenspace region/stash/event list routes and participant notes reads.

Fix:

- Snapshot manifest-derived data under `_manifest_lock` before filtering or returning JSON.
- Prefer shallow/deep copies depending on whether nested dicts are returned to Flask.

Tests:

- Add concurrency tests for Screenspace event list while bulk-excluding events.
- Add tests that returned region/stash payloads are stable snapshots, not live aliases.

### P2: Screenspace SSE Can Leave Clients Stale After Queue Overflow

`screenspace_server._notify_sse_clients()` drops notifications when a per-client queue is full. The keepalive path only sends a comment, not the current task state, so a client can remain stale until another notification fits the queue or the page falls back to polling.

Fix:

- Make overflow coalescing stateful: on `queue.Full`, drain one or all pending notifications and enqueue a single `"update"` marker.
- Alternatively, store a per-client dirty flag and have the keepalive tick send `_sse_task_payload()` whenever dirty.
- Keep the bounded queue to protect the server from slow clients.

Tests:

- Add a Screenspace API test with a full SSE queue and assert a later generator tick yields fresh task state.
- Add a unit test for `_notify_sse_clients()` overflow behavior that proves state updates are coalesced, not lost.

### P3: Transcription Cancellation Is Segment-Boundary Only

`TranscriptWorker.cancel()` sets `_cancelled`, but `transcribe_video()` only observes cancellation through the `on_segment` callback. A long model startup or long first segment can keep running after the user cancels.

Fix:

- Add an optional `cancel_flag` parameter to `transcripts.transcribe_video()`.
- Check it before model load, after model load, before iterating segments, and between yielded segments.
- If faster-whisper exposes no safe hard abort, document cancellation as cooperative and update UI copy to "Stopping..." until the worker transitions.

Tests:

- Add a worker test where cancellation before model load prevents `model.transcribe()`.
- Add a worker test where cancellation between fake segments marks the task cancelled and clears partial segments.

## Implementation Order

### Task 1: Safe File and Manifest Foundations

- [ ] Add atomic output reservation to `files.py`.
- [ ] Update ffmpeg call sites that use reserved paths to clean up unused placeholders on early failure.
- [x] Add a manifest write lock to `viewer.save_manifest()`.
- [ ] Run:

```bash
uv run --extra dev pytest -c tests/pytest.ini tests/test_manifest.py tests/test_clip_pipeline.py
```

### Task 2: Studio Generate and Reel Correctness

- [ ] Fix titlecard cache lifecycle for parallel Studio generation.
- [ ] Forward reel cancellation/titlecard kwargs into `_process_single_clip_segments()`.
- [ ] Apply titlecard settings to `/studio/api/reel-direct`.
- [ ] Pass generate cancellation into running ffmpeg workers and clean up cancelled outputs.
- [ ] Add intake cancellation support and a busy-slot decision.
- [ ] Persist streamed successes in generator `finally` blocks.
- [ ] Run:

```bash
uv run --extra dev pytest -c tests/pytest.ini tests/test_studio_api.py tests/test_clip_pipeline.py tests/test_studio_frontend_source.py
```

### Task 3: Studio Frontend Stream Handling

- [ ] Add or inline a shared NDJSON reader helper with `response.body` checks.
- [ ] Fix sheet branch failure accounting and zero-result messaging.
- [ ] Fix queued/result badge selectors.
- [ ] Normalize 409/error handling for reel streams.
- [ ] Run:

```bash
uv run --extra dev pytest -c tests/pytest.ini tests/test_studio_frontend_source.py
```

### Task 4: Shared Backend Hardening

- [ ] Lock Studio thumbnail cache operations.
- [ ] Snapshot Screenspace manifest reads under `_manifest_lock`.
- [ ] Coalesce Screenspace SSE queue overflow into a later full state update.
- [ ] Improve transcript cancellation checks.
- [ ] Run:

```bash
uv run --extra dev pytest -c tests/pytest.ini tests/test_screenspace_api.py tests/test_transcripts_api.py tests/test_transcripts.py
```

### Task 5: Full Verification

- [ ] Run focused tests from Tasks 1-4.
- [ ] Run full checks:

```bash
uv run --extra dev pytest -c tests/pytest.ini
uv run ruff format --check
uv run ruff check
uv run ty check
```

## Notes for Reviewers

- Keep the independent artifact/reel UX from `#348`; do not restore a single global Studio lock.
- Prefer one well-scoped lock per shared resource: output filename reservation, manifest write, thumbnail cache, generated-output list, Screenspace manifest.
- Flask NDJSON generators currently parse request JSON before yielding; if future changes read `request` inside generators, wrap them with `stream_with_context()`.
- Avoid backwards-compatibility shims for old manifest shapes; update tests to the new behavior.
- For UI verification, do not install browser automation. Ask for a manual Studio/Screenspace/Transcripts browser pass after tests are green.
