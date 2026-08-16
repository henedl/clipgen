# Performance Review & Plan — clipgen (PERFORMANCE-PLAN-3)

> **Status: closed 2026-07-30.** Seven of eight items shipped; the last one (**8b**, the Drive
> listing TTL cache) landed with the archive commit. **Item 4 (Ollama prewarm) is deferred
> indefinitely** on RAM grounds — see the note in its section; it is a decision, not a backlog
> entry, so a future perf pass should not re-propose it. Items **8c** and **8d** shipped as
> partial-by-design: the halves that were dropped are recorded inline with the reasoning.

## Context

This is a thorough review of where clipgen can gain performance — real speedups *and* perceived-responsiveness improvements (warm-ups, preloads, wait coverage). Two prior passes (`plans/archive/PERFORMANCE-PLAN.md`, `PERFORMANCE-PLAN-2.md`) landed nearly everything they identified, so this review focused on (a) explicitly deferred wins, (b) gaps those passes didn't cover, and (c) perceived-latency polish. Every finding below was **verified against the current code** (many candidate findings from exploration were rejected as already-done or wrong — see "Rejected" at the end, kept so future passes don't re-propose them).

**Execution model (per the maintainer):** the workspace that authored this plan produces the plan document only, so other workspaces can each pick up one item. Each item below is a self-contained work packet: one small PR, `feat:` bumps `build/VERSION` patch, run the `/check` skill before commit, update this plan's status column as items land (repo hard expectation).

**Decision:** hardware encoding defaults to **`"auto"`** (speed-first; runtime fallback to libx264).

## Status

Status audited against the tree on 2026-07-29 (the table had drifted — every item still read
☐ although #3 and three of #8's sub-items had landed), then closed out on 2026-07-30 with 8b's
landing and item 4's deferral. Evidence column records what the audit actually found, so a future
pass can re-verify cheaply rather than re-grepping from scratch.

| # | Item | PR type | Effort | Status | Evidence |
|---|------|---------|--------|--------|----------|
| 1 | Parallel ffprobe loops | feat | S | ☑ | `video.py:1381` `_parallel_probe`; wired at `_detect_clip_mismatches` (`video.py:2201`) and `build_source_timeline` (`video.py:1424`). Measured 4.2× on a 20-clip probe (852 ms → 202 ms), identical results. Helper name and shape differ slightly from the item text — see below |
| 2 | SSE-primary / poller-fallback fix | fix | S | ☑ | `screenspace-tasks.js:1226` — `onUnsupported: startPolling` plus `stopPolling()` in `onOpen`; guarded by `test_task_stream_retires_the_poller_when_it_reconnects`. Landed in the frontend packet (see below), not as its own `fix:` PR |
| 3 | Titlecard concat-demuxer stream copy | feat | L | ☑ | `video.py:2480` `concat_copy`; `titlecards.py:72` `_body_is_copy_safe` + dispatch at `:640`; `agents/PERFORMANCE.md:36` rewritten. Landed with deviations — see "Item 3 follow-ups" below |
| 4 | Ollama model prewarm during transcription | feat | S–M | ⊘ | **Deferred indefinitely (2026-07-30) — RAM.** Never built; see the note in the Item 4 section |
| 5 | Bounded frame-0 preload (Screenspace) | feat | S | ☑ | `screenspace.js:288` `queueFrameZeroPreload` (concurrency 2); `pickId`/`selectParticipant` now resolve above it. Landed with deviations — see "Item 5 follow-ups" below |
| 6 | Hardware video encoding (VideoToolbox) | feat | M | ☑ | `video.py` `check_videotoolbox_support` / `resolve_video_encoder` / `video_encoder_args` / `note_hw_encode_failure` / `run_ffmpeg_encode`; five call sites converted. Measured 1080p/60s re-encode 48.4s → 12.5s wall (524s → 20.6s CPU); reel concat end-to-end 35.5s → 8.8s. Landed with deviations — see "Item 6 follow-ups" below |
| 7 | Skeleton coverage for initial loads | feat | S | ☑ | `.pill-skeleton` (transcripts) and `.wf-row-skeleton` (workflows palette + stashes); `#transcriptEmpty` ships `.hidden`. Landed with deviations — see "Item 7 follow-ups" below |
| 8 | Small backlog (each optional, S) | mixed | S | ☑ | 8a/8c/8e landed via their own PRs; 8c's scroll half and 8d's xref half landed in the frontend packet; 8b landed 2026-07-30 — see the per-bullet markers in the Item 8 section |

Items #2, #5, #7 and the open halves of #8c/#8d landed together as one frontend packet
(2026-07-30) — one PR, one `/ui-check` pass, since they touch four overlapping files and share
their browser verification. **Nothing remains:** #8b landed the same day and #4 is deferred
indefinitely.

---

## Item 1 — Parallelize sequential ffprobe loops (`video.py`) — real speedup

Reel validation probes every clip sequentially; multi-file participants probe each source file sequentially. For a 20–50 clip reel that's seconds of serial ffprobe on every reel generation.

- **`_detect_clip_mismatches()`** (~video.py:1709): sequential `probe_video_properties` list comprehension → ThreadPoolExecutor, ordered results, gate on `len >= 2`, workers `min(4, len)`.
- **`build_source_timeline()`** (~video.py:1238): parallel `get_file_duration` over paths, then sequential cumulative fold; any `None` → return `None` (contract preserved). Single path keeps today's loop.
- Add a private `_parallel_map_ordered(fn, items, max_workers)` used by both sites (pre-allocated results list indexed by future — the `agents/PERFORMANCE.md` pattern).
- **No locks needed**: the probe caches are plain dicts keyed by `(resolved_path, mtime_ns)`; paths within each loop are distinct, and cross-call duplicate probes are idempotent. Document with a comment.
- Workers must resolve `probe_video_properties` via module attribute at call time so monkeypatched tests keep working.
- **Tests** (`tests/test_video_commands.py`): ordering + cumulative math with a recorder probe; `None` propagation; single-item takes sequential path; existing concat tests stay green.
- No config knob (ffprobe too cheap to justify one).

### Item 1 notes (as shipped)

- The helper is `video._parallel_probe(items, probe_fn)` (`video.py:1381`), not
  `_parallel_map_ordered(fn, items, max_workers)`: no `cancel_flag` / `on_error` / caller-owned
  results list. Both call sites already treat a `None` probe as "skip", and ffprobe is short enough
  that mid-probe cancellation isn't worth the surface area. `pipeline._parallel_map_ordered`
  (`pipeline.py:711`) keeps the full signature for clip generation; the two are deliberately not
  merged (`pipeline` imports `video`, so the dependency can only point one way).
- Ordering is asserted rather than assumed: `test_build_source_timeline_order_survives_parallel_probes`
  (`tests/test_multi_video.py:92`) makes the probes complete in reverse order and checks both the
  timeline and the completion order.

## Item 2 — SSE-primary / poller-fallback fix (`screenspace-tasks.js`) — latency + redundant requests

Verified behavior: SSE and the 3s poller do **not** run concurrently at boot. The real gap: after an SSE drop the poller starts, and a later `startSSE()` (from queueing, results, or the `visibilitychange` handler at screenspace-tasks.js:1143) opens a new stream while the poller keeps running → both transports run concurrently from then on. Also `startSSE()` passes no `onUnsupported`, so no-EventSource environments get zero task updates.

- In `startSSE()` (screenspace-tasks.js:1062–1079): add `onUnsupported: startPolling` (`createSSEStream` supports it, utils.js:931) and call `stopPolling()` in `onOpen` next to the existing `state.sseFellBack = false`. workflows-runs.js already implements this pattern — mirror it.
- **As shipped:** exactly that. Worth recording for a future pass: `onUnsupported` is parity-only (no browser this ships in lacks `EventSource`); the `stopPolling()` in `onOpen` is the fix that matters. Safe because `make_sse_channel`'s streamer yields the current payload on connect (`server_utils.py:303`), so retiring the poller there cannot open a data gap.
- **Transcripts 3s task poll: leave as-is** (wontfix). No SSE endpoint exists in transcripts_server.py; tasks run minutes, ≤3s latency is invisible; the poller is entangled with `POST_COMPLETION_GRACE_CYCLES` and agent rearming and already self-gates. Record in the PR description.
- **Test**: source-level assertion (pattern of `tests/test_studio_frontend_source.py`) that the `createSSEStream("api/tasks/stream"` options include `onUnsupported` and `onOpen` contains `stopPolling()`.
- `fix:` — no VERSION bump.

## Item 3 — Titlecard concat-demuxer stream copy (`titlecards.py`, `video.py`) — highest-value real speedup

The acknowledged future win in `agents/PERFORMANCE.md`: `wrap_clip_with_cards()` (titlecards.py:453) currently re-encodes the **entire clip body** through libx264 filter_complex just to prepend/append ~2s cards. Fast path: encode only the cards to exactly match the clip's stream parameters, then concat-demux all three with `-c copy`. Expected 3–10× on titlecard wrap for typical clips.

- **`video.probe_video_properties()`** (video.py:1258): extend `-show_entries stream=` with `pix_fmt,profile,level,sample_rate,channels`; add keys to result dict + DEBUGGING stub. Additive, safe.
- **New `video.concat_stream_copy(paths, output_file, *, cancel_flag)`**: extract list-file + `-f concat -safe 0 -i list -c copy` from `_concatenate_demuxer` (1929–1962, keep the quote-escaping); `_concatenate_demuxer` calls it for its first attempt (helper reused twice → satisfies the extraction preference).
- **`titlecards._build_card_frame()`** (79): optional `stream_spec` dict → matched output args (`-pix_fmt`, `-profile:v`, `-level:v`, `-r`; keep preset/CRF from config) and, when the clip has audio, `anullsrc` encoded `-c:a aac -ar {rate} -ac {channels} -shortest` so cards carry the same stream layout.
- **Endcard cache key** (titlecards.py:326): append the stream signature so cached endcards never cross mismatched clips; plain cards keep a distinct key.
- **`wrap_clip_with_cards()` dispatch**: eligibility gates — `config.TITLECARD_FAST_WRAP` on, probe ok, `h264` + `yuv420p`, even dimensions, `0 < fps <= 120`, profile in `baseline/main/high`, audio in `(None, aac)` with sane rate/channels. Eligible → spec cards + `concat_stream_copy` → verify (`verify_output_file` + fresh probe duration within `1.0 + 0.05×expected` of `clip + n_cards×card_duration`) → `os.replace`. Any failure → delete fast artifacts, rebuild **plain** cards, run existing filter_complex path (refactor into `_wrap_reencode(...)`; the existing filter graph assumes video-only cards, so spec cards must not be reused in fallback).
- **Config**: `TITLECARD_FAST_WRAP: bool = True` near `TITLECARD_ENCODE_PRESET` + `SETTINGS_DESCRIPTIONS` entry; *not* in `STUDIO_SETTINGS` (debugging escape hatch). Update the `agents/PERFORMANCE.md` knob table and its "(A future win …)" sentence in the same PR.
- **Tests** (`tests/test_titlecards.py`, `tests/test_video_commands.py`): probe extension; fast path asserts final command uses `-f concat`+`-c copy` with no `-filter_complex` and card commands carry the matched flags; fallback triggers (concat rc≠0, duration mismatch, hevc ineligible); endcard cache keys per spec; `concat_stream_copy` unit test.
- **Risks**: mismatched SPS/PPS gives garbled playback without an ffmpeg error — mitigated by matching at encode time (gates), not just the duration check. AAC priming tick at joins — same artifact reels already accept. Homebrew drawtext caveat unchanged (`check_drawtext_support()` stays authoritative).

### Item 3 follow-ups (shipped, but not as written above)

The speedup works and is covered by five tests (`tests/test_titlecards.py:307–460`), so these are
**optional hardening, not regressions**. Recorded because the item text above no longer describes
the code — read this before building on it.

- Helper is `video.concat_copy` (`video.py:2480`), not `concat_stream_copy`, and the planned reuse
  did not happen: `_concatenate_demuxer` (`video.py:2388`) still builds its own concat-copy
  command. The two share only `_concat_list_file` (`video.py:481`).
- **No `TITLECARD_FAST_WRAP` knob.** The fast path is unconditional whenever the body is
  copy-safe, so there is no escape hatch if some source shape corrupts. The most worthwhile
  follow-up of the five.
- **No post-concat duration verification.** `concat_copy` calls `verify_output_file` only
  (`video.py:2526`); `expected_duration_sec` feeds progress reporting, not validation.
- Eligibility gates are narrower than specified. `_body_is_copy_safe` (`titlecards.py:72`) checks
  h264 + yuv420p + fps > 0 + aac but **not** even dimensions, `fps <= 120`, or
  `profile in baseline/main/high`; `profile`/`level` were never added to the probe
  (`video.py:1472`), and pix_fmt is hardcoded in `_x264_video_args` (`titlecards.py:49`) rather
  than matched. Combined with the missing duration check, the "garbled playback with no ffmpeg
  error" risk noted above is guarded at encode time only.
- No standalone `concat_copy` unit test — it is exercised only indirectly through the titlecards
  tests.

## Item 4 — Ollama model prewarm during transcription — perceived win

> **Deferred indefinitely (2026-07-30). Do not build; do not re-propose as an oversight.**
> The whole point of the item is to hold an Ollama model resident *while Whisper is still
> loaded and decoding* — on the limited-RAM machines clipgen runs on, that is two multi-GB
> models at once. Swapping (or an OOM) mid-transcription loses minutes of real work and can
> take the transcript with it; the thing it buys is a one-time, already-visible 5–30 s wait
> behind a "generating…" label. Wrong side of that trade at any RAM level we can assume.
> An `OLLAMA_PREWARM_ENABLED` knob doesn't rescue it either: a default-on knob ships the
> hazard, and a default-off knob is a feature nobody finds. Revisit only if clipgen ever
> gains a reliable free-RAM probe *and* a reason to trust it. The design below is kept
> intact for whoever revisits it.

Today the summary agent's first `generate()` after transcription pays the 5–30s Ollama model load while the user watches "generating…". Page the model in while Whisper is still running.

- **`ollama_client.prewarm_model(model, keep_alive="10m") -> bool`**: same request shape as `unload_model` (ollama_client.py:129) but positive keep_alive, ~120s timeout, never raises, does **not** auto-start `ollama serve`.
- **`transcripts.TranscriptWorker.on_task_start`** callback (symmetric to existing `on_task_complete`, fired at the RUNNING transition, transcripts.py:953–963, exception-swallowed).
- **`transcripts_server._maybe_prewarm_agent_models()`** wired in `_init_transcripts_state`: gates in order — `config.OLLAMA_PREWARM_ENABLED`; first enabled agent via `_agent_enabled` (1603); model via `_agent_model` (134); TTL debounce (~240s, under keep_alive/2); then daemon thread: `is_available()` (never start the server for a prewarm) → model installed (never trigger a download) → `_cancel_pending_unload(model)` → `prewarm_model`. Mirrors the `api_transcribe_warmup` guarded-thread pattern (1216–1227). Server-side only, no frontend work.
- **Config**: `OLLAMA_PREWARM_ENABLED: bool = True` + `SETTINGS_DESCRIPTIONS` (call out RAM: LLM resident while Whisper runs; never downloads) + `STUDIO_SETTINGS` (tab Summaries).
- **Tests**: `test_ollama_client.py` request shape/failure; `test_transcripts.py` callback fires on RUNNING; `test_transcripts_api.py` prewarm called once on POST /api/transcribe, suppressed by knob-off / agents-disabled / not-installed / TTL.

## Item 5 — Bounded frame-0 preload (Screenspace boot) — perceived win

`screenspace.js` boot (~3745–3758) fires an **unbounded** fetch of frame 0 for every participant simultaneously, competing with the selected participant's own first-frame render (server has a 3-slot VideoCapture pool).

- Reorder the `apiGet("api/participants")` handler: compute `pickId` (localStorage-restore block currently *below* the loop) **first**, call `selectParticipant(pickId, initialTs)` so its request wins, then pump a bounded queue (selected first, then the rest) at concurrency 2 — the `studio-scrubber.js` prefetch-queue pattern (lines ~27–52). Keep `_videoVersions` seeding above everything (`?v=` cache-bust contract) and the existing `TODO: apiGetBlob` comment. ES5 note: release the slot in a final `.then()` after `.catch()` so it always decrements.
- No config change. Verify with `/ui-check` — `tests/ui/shot.py screenspace --eval` can read back `performance.getEntriesByType("resource")` to confirm ≤2 concurrent `frame/*/0` requests with the selected participant first, rather than reading DevTools Network by hand.

### Item 5 follow-ups (shipped, but not as written above)

- **The selected participant is enqueued last, not first.** The item text above is wrong on this
  point: `_fetchFrame` consults `_preloadedFrames` only at request time (`screenspace.js:1571`),
  so a blob that lands after `selectParticipant` already issued its `<img>` request buys nothing
  at boot. The queue exists for the participants the user might switch *to*.
- Two hazards the item text didn't anticipate, both fixed in the same change because spreading
  the preload over seconds turns them from theoretical into likely:
  - **Stale-version writes.** `selectParticipant`'s `api/video/info` handler drops
    `_preloadedFrames[pid]` when the source mtime changed (`:1096-1101`); the queue captures
    `_videoVersions[pid]` at enqueue and discards a result whose version moved.
  - **`pagehide` leak.** The handler now sets `_preloadStopped` and clears the queue, and an
    in-flight `.then` revokes its own blob instead of storing it.
- **Not done:** capping the queue to the first N participants. Concurrency 2 already removes the
  boot burst; a cap would silently drop the instant-switch payoff for later participants.
- Measured in the `/ui-check` fixture (2 participants): three `frame/*/0` requests, the selected
  participant's own request issued in the same tick as the queue's first item, and the queue's
  own selected-participant entry served from the memory cache (identical `responseEnd`) rather
  than costing a second extraction.

## Item 6 — Opt-in hardware encoding, default `auto` (VideoToolbox) — real speedup on Apple Silicon

No `-hwaccel`/`h264_videotoolbox` anywhere today (verified). Clips default to stream copy (`config.REENCODING` off), so this targets the re-encode paths: reel filter_complex (video.py:1895), concat re-encode fallback (1977–1992), cut re-encode branch (478–480), `compress_to_size` (1503), timelapse mp4.

- **`check_videotoolbox_support()`**: `sys.platform == "darwin"` + parse `ffmpeg -encoders` (mirror the `check_webp_support` pattern + module cache, video.py:141–145).
- **`resolve_video_encoder()`**: `FFMPEG_VIDEO_ENCODER` knob → `h264_videotoolbox` when explicitly selected, or when `auto` + supported + session `_hw_encode_failed` flag clear; else `libx264` (warn once on explicit-but-unsupported).
- **`video_encoder_args(crf, preset)`**: libx264 → exactly today's flags per site (byte-identical when hw unavailable/off); VT → `-c:v h264_videotoolbox -q:v {clamp(100-2*crf, 30, 80)} -allow_sw 1` (`-q:v` needs Apple Silicon; Intel rejection is caught by runtime fallback).
- **`run_ffmpeg_encode(build_command, ...)`** wrapper: on rc≠0 with hw args → set `_hw_encode_failed`, warn once, rerun with libx264 args.
- `compress_to_size` with VT: **single-pass** `-b:v {target}k` (VT has no `-pass` — skip pass 1, an extra 2× win); keep the size check and fall back to the libx264 two-pass body on overshoot.
- **Not converted**: titlecard card generation and Item 3's fallback re-encode (cards must parameter-match; fallbacks stay maximally compatible).
- **Optional second commit**: `SCREENSPACE_HWACCEL_DECODE: bool = False` (config-only) — insert `-hwaccel videotoolbox` before `-i` in `_ffmpeg_pipe_frames` (screenspace_frames.py:201–204); containment in `_scan_via_ffmpeg_pipe` (292): hw enabled + zero frames consumed + not cancelled → flag, log, rerun once without hwaccel. Worthwhile mainly for decode-bound scans (4K/HEVC + cheap CV).
- **Config**: `FFMPEG_VIDEO_ENCODER: str = "auto"` (options auto/libx264/h264_videotoolbox) + `SETTINGS_DESCRIPTIONS` (state the size/quality-per-bit trade-off) + `STUDIO_SETTINGS` select. Document both knobs in the `agents/PERFORMANCE.md` table.
- **Tests** (`tests/test_video_commands.py`, timelapse tests): capability parse; argv byte-identical with encoder forced to libx264; VT args (`-q:v`, no `-crf`, no `-pass`); fallback wrapper rc=1→rc=0 rewrites to libx264 and the flag sticks; timelapse honors the knob. Since default is `auto`, tests that pin argv must force `libx264` via config (and one test covers auto-detection separately with `check_videotoolbox_support` monkeypatched both ways).
- **Risk**: encoder listed but broken (VMs) → one-shot runtime fallback covers it. Larger files per quality — documented; users can pick libx264 in Settings.

### Item 6 follow-ups (shipped, but not as written above)

Measured on an M-series Mac with `h264_videotoolbox` present. Read this before building on it.

- **`compress_to_size` was deliberately NOT converted** — the item text above is wrong on this
  point. Measured: asking for 105 kbps video, `h264_videotoolbox` delivered **246 kbps** (2.3× over
  target) where x264 two-pass delivered 127 kbps. The planned "single pass + fall back to two-pass
  on overshoot" therefore overshoots essentially always, so it pays for the hardware attempt *and*
  the two-pass — measured 2.4× slower for byte-identical output. Size capping stays libx264
  two-pass, locked in by
  `test_compress_to_size_stays_on_libx264_with_hardware_available`. The planned
  `_compress_two_pass` / `_compress_single_pass` split was reverted with it (one caller each).
- **The win needs a real workload.** On a short 720p clip the filter graph plus encoder setup
  dominate and VideoToolbox came out *slower* (2.67s vs 1.31s). The 4× numbers in the status table
  are 1080p/30s+ sources. Do not quote a hardware speedup for small artifacts.
- `note_hw_encode_failure(encoder)` is public because the timelapse runner
  (`screenspace_scans.generate_timelapse`) needs the sticky flag but cannot use `run_ffmpeg_encode`
  — its progress parsing owns its own `Popen` loop. Its retry mirrors the wrapper by hand.
- `video_encoder_args(encoder, crf=None, preset=None)` omits `-crf`/`-preset` when unset, so the two
  concat sites (which passed neither and relied on libx264's own crf 23 / preset medium) keep
  **byte-identical** software argv. Same reason `build_ffmpeg_cut_command` adds no `-c:v` at all
  unless a non-libx264 encoder is passed — `tests/test_video_commands.py`'s
  `assert "-c:v" not in cmd_reencode` still holds unmodified.
- Composer's annotation burn (`composer_server._build_overlay_command`) was added to the site list
  (it post-dates the item text); gif output never resolves an encoder.
- `tests/conftest.py` gained an autouse `_force_software_encoder` fixture. Without it the `"auto"`
  default makes every argv assertion depend on the host's ffmpeg build **and** every test shells out
  to `ffmpeg -encoders`.
- Drive-by: `settings-modal.js`'s label humanizer rendered `FFMPEG_*` as "Ffmpeg"; now "FFmpeg".
- **Deferred, not missing:** the optional `SCREENSPACE_HWACCEL_DECODE` second commit (hw *decode* in
  `_ffmpeg_pipe_frames`) was consciously left out — separate risk, needs its own zero-frames
  containment state machine, and only helps decode-bound 4K/HEVC scans. Do not re-propose it as an
  oversight.
- Not verified: the runtime fallback on hardware that lists the encoder but cannot run it (no Intel
  Mac / VM available here). Unit-tested only (`test_run_ffmpeg_encode_falls_back_to_libx264_once`).

## Item 7 — Skeleton coverage for initial loads — perceived polish

Studio sheet (`populateSheetSkeleton`, studio.js:1082) and the Screenspace frame (`.skeleton-frame`) already have skeletons. Verified gaps show *misleading* empty states during the first fetch:

- **transcripts.html:40**: `#participantPills` ships `No participants` before the fetch resolves → replace with 3–4 `<span class="skeleton pill-skeleton"></span>` placeholders (`.pill-skeleton` in transcripts.css, sized off tokens); `renderPills()` (transcripts-pills.js:97) rewrites innerHTML on first data so the real empty state still appears — post-fetch only.
- **transcripts.html:159–162**: `#transcriptEmpty` ships "No transcript available" pre-data → add `hidden` in HTML, unhide in the hub's no-participants path.
- **workflows.html:31/35**: `#wfPalette` and `#wfStashList` are blank until catalog/blueprints resolve → seed ~6 / 2 skeleton rows (page CSS); `renderPalette()` and the stash renderer already replace container contents. Zero JS changes. Leave `#wfRuns`/`#wfCanvasEmpty` (intentional empty states).
- Reuse only the `.skeleton` shimmer token (tokens.css:399); no `buildSkeletonGrid` (these are pills/rows, not tables).
- **Tests**: DOM-wiring/source assertions (`tests/test_transcripts_dom_wiring.py`, `tests/test_workflows_frontend_source.py`) that the skeleton classes exist and `#transcriptEmpty` carries `hidden`.
- **Risk**: a failed fetch leaves shimmer instead of a message — acceptable; the fetch `.catch` already toasts.

### Item 7 follow-ups (shipped, but not as written above)

- **That last risk bullet was wrong, and the change would have shipped a bug on it.**
  `loadParticipants` (`transcripts.js:605`) had *no* `.catch` at all and returned early on
  `!data.ok`, so a failed boot fetch would have left the pills shimmering forever **and** the
  transcript pane permanently blank — a truthful terminal state converted into a permanent lie.
  Both exits now call `_clearBootPlaceholders()`, which falls back to the real empty states and
  no-ops when a list is already rendered (so a *refresh* failing stays harmless). Workflows had
  the same hole: `setCanvasState("error")` now clears both sidebar containers, because a
  catalog/blueprint rejection never reaches `renderPalette`/`renderStashPalette`.
- `#transcriptEmpty` ships the `hidden` **class**, not the attribute — every JS site toggles
  `.hidden` (rule at `transcripts.css:39`), so the attribute would never have been removed.
- Skeleton geometry classes are `.pill-skeleton` (transcripts.css) and `.wf-row-skeleton`
  (workflows.css, one class for both the palette and the stash list). Both are geometry-only over
  the shared `.skeleton` shimmer, following `.skeleton-cell`'s raw-px precedent.
- Containers carry `aria-busy="true"` in the HTML, cleared by `renderPills` / `renderPalette` /
  `renderStashPalette`; the skeleton spans themselves are `aria-hidden`.
- The plan's line numbers for workflows were wrong (`31/35`); the containers are at
  `workflows.html:138/142`.

## Item 8 — Small backlog (optional; verified but low individual impact)

- **8a ☑ Lazy `import gspread`** in clipgen.py:21 → first Google-Sheets use. Measured: gspread is ~200ms of the ~260ms module import (`-X importtime`); Excel-only and `--help` flows never need it. `fix:`/`perf:`-class change; check `ty` on the deferred-import pattern. **Landed** via `2f84d24b` (#563) rather than an Item-8 packet: `clipgen.py` imports only `sys`/`pathlib`, and `cli.py` defers gspread into function bodies (`cli.py:884/911/934/956`).
- **8b ☑ TTL cache for the Drive spreadsheet listing** (`server._cached_spreadsheet_meta`, 300 s): module-level `(monotonic stamp, metas)` + a dedicated lock. **Landed 2026-07-30** with three deviations from the item text worth recording:
  - **Three call sites, not two.** The item counted the picker route (`server.py:3497`) and open-by-name (`:3239`); the worksheet dropdown is a third, one level down — `app.list_worksheet_titles` (`app.py:272`) listed Drive again. It now takes an optional `doc_list` (the `PERFORMANCE.md` "accept what the caller already holds" pattern), so one pick-a-sheet flow went from **three** identical `files.list` calls to one. The URL branch still skips the listing entirely (`35a7a606`).
  - **Staleness needed two escape hatches, not one.** `?refresh=true` (wired to a new Refresh button in the Google picker) covers the *visible* list; `_spreadsheet_names_for()` re-lists once when a name isn't in the cache, so a spreadsheet created mid-session is still openable by paste. The cache is also dropped when a new client authenticates, or a second account would inherit the first's listing.
  - **The lock is held across the fetch on purpose** (single-flight: concurrent misses queue behind one Drive call, waiters re-check freshness after acquiring). Only successful fetches are stored, so a 429 can't poison it. Cached in `server.py`, deliberately *not* in `google_api.py` — a process-lifetime TTL would be wrong for interactive CLI sessions, which already have a per-invocation cache (`cli.py:1033`).
  - Gotcha for anyone touching `refresh()` in `start-overlay.js`: `loadGoogleSheets` is chained as `.then(function () { return loadGoogleSheets(); })`, not by reference — passed by reference it receives the previous link's resolved value as its `force` argument and re-lists on every overlay open.
- **8c ◑ Transcripts segment list**: `renderSegments()` (transcripts.js:812–896) rebuilds one big innerHTML string per render. Cheap wins: `content-visibility: auto` (+ `contain-intrinsic-size`) on segment rows in transcripts.css; preserve scroll position across rebuilds. Virtualize only if profiling shows >2000-segment sessions hurting. **CSS half landed** via `0c9716bd` (#560): `transcripts.css:1233` carries both properties, with a comment on the `auto` fallback for `scrollToSegment()`'s rect math. The scroll-preservation half **landed** in the frontend packet: `renderSegments` captures and restores `#trMain`'s `scrollTop` (scroll lives there, not on `#segmentList` — same probe `renderPartialSegments` uses), gated on a module-local `_renderedSegmentsPid` so a participant switch resets to 0 and a same-transcript rebuild (heatmap toggle, tooltip toggle, streaming→final swap) holds position. Two things the item text didn't mention and that a re-implementation would get wrong: the restore has to be marked programmatic via the new `TS.ignoreNextScroll()` or `initAutoFollowScrollPause` (`transcripts-video.js:1161`) reads it as a reader scroll and kills playhead follow for 3 s; and it has to be written in the same task as the `innerHTML` wipe so `initPipScroll`'s rAF-coalesced listener never sees the transient top. Verified in-browser (300 filler segments): same-participant re-render held 900 px, participant switch reset to 0. **Virtualization half closed 2026-08-16: measured, not worth it.** `transcripts.renderSegments` is now spanned; on a synthesized 2400-segment manifest (the gate's own threshold) it renders in **30.1 / 32.4 ms across two runs, 0 longtasks**. Pushed to 10 000 segments it is 115.8 ms / 1 longtask — linear at ~12 us/segment, so the 50 ms longtask line sits near **~4300 segments**, i.e. a ~3.5-hour session at this segment density. Real 30-90 min sessions land at ~600-1800 segments (~7-22 ms). The `content-visibility` half already landed is carrying it. Revisit only if sessions routinely exceed ~4000 segments; synthesize the manifest rather than transcribing (recipe in [profile/SKILL.md](../../agents/skills/profile/SKILL.md) Step 2) — 2000 real Whisper segments is hours of audio.
- **8d ◑ Gate always-on pollers**: transcripts xref poller (transcripts.js:119, 30s) could start lazily on first xref use — **retargeted and landed** in the frontend packet. "Lazily on first xref use" was dropped: the indexes also feed cross-transcript search, so there is no clean first-use trigger. What shipped instead: the poller's `/studio/api/sheet` leg goes idle (`_sheetXrefIdle`) once a response reports `sheet_loaded: false`, since that route can only keep answering with no rows, and re-arms in `startXrefPolling` (which also runs on tab focus — a spreadsheet opened from another tab reloads only *that* document). The first tick always runs: that handler is this page's only caller of `clipgenApplyConfig`, so `CLIPGEN_CONFIG` on Transcripts now refreshes per focus rather than every 30 s. Small win, recorded honestly: two requests a minute to a route in the same Flask process. Studio's intake pollers — **superseded, treat as wontfix**: there are now three (endpoints were combined per-domain), they carry `{ maxIntervalMs: 30000 }` idle backoff which removes most of the cost, and `studio.js:4786-4791` documents a deliberate decision to poll regardless of visible sub-tab so the start-overlay pills and sub-tab badges stay fresh.
- **8e ☑ `decoding="async"`** on dynamically created thumbnails/imgs (queue cards, results) — trivial, prevents main-thread decode jank. **Landed** across ~18 sites including both the plan named (`screenspace-results.js:652/764`, `studio.js:2591`). The remaining misses are `new Image()` preload/measure objects that never enter the DOM, where the attribute is a no-op.
- **Studio grid filter re-render** (studio.js:1432 region): deferred item from PERF-PLAN-2 §4.1 — **closed 2026-08-16: measured, not worth it.** Baseline on 200 rows x 12 participants: filter toggles cost a full re-render at ~27 ms (avg 22.8-31.3 ms over 24 toggles across two runs), 0-2 longtasks. The incremental-`.hidden` idea is confirmed in shape but not worth building at this size. Full numbers and the revisit threshold are in PERF-PLAN-2 §4.1.

## Rejected findings (verified against code — do not re-propose)

- Ollama model prewarm during transcription (**Item 4**): deferred indefinitely on RAM grounds, not an oversight. Reasoning in that item's banner.

- `defer` on `<script>` tags: all scripts already sit at end of `<body>`, pages are served from localhost, and satellite load order is a documented contract. No meaningful FCP win.
- Screenspace boot fetch "waterfall": the four boot fetches (participants/regions/stashes/tasks) are already parallel independent chains.
- Thumbnail/sprite/audio cache LRU refresh "missing": `move_to_end` present on all three (server.py:535/631/681).
- Studio `/api/generate` batching "deferred": landed — ThreadPoolExecutor + `as_completed` inside the NDJSON streaming route (server.py:1567–1572).
- Screenspace `api_video_frame` "uncached": 256-entry LRU frame cache with mtime keys exists (screenspace_server.py:158, 1006–1037).
- Heatmap generation "blocking request threads": wrong — heatmaps are produced in `ScreenspaceWorker`, not in routes.
- Whisper prewarm: already shipped (`TRANSCRIBE_PREWARM="queue_open"`, `/api/transcribe/warmup`).
- Screenspace task cards "no progress/ETA": progress bar + per-second ETA ticker already rendered (screenspace-tasks.js:879–888).
- Model-view preview "no feedback": "Loading preview…" label + debounce option exist (screenspace-model-view.js:207, 432).
- Fast Scan mode "not activated": fully wired (tool class attrs + worker fast_opts) with a scan-mode picker in the UI.
- Startup import cost: measured — full `import server` ≈0.7s; cv2/easyocr/torch/whisper all lazy already. Only gspread remains (Item 8).
- List virtualization: screenspace-results uses IntersectionObserver for 500+ rows; Studio thumbs and viewer thumbs lazy-load the same way.

## Verification

For each item: the per-item test plans above, plus `/check` (ruff format + lint, ty, full suite) before every commit; browser checks (Items 5, 7, and the Settings UI of 4/6) run through `/ui-check` (page smoke + screenshots + `shot.py --eval` probes), with the maintainer asked only for interaction feel; benchmark notes worth capturing in PR bodies: titlecard wrap before/after on a ~60s clip (Item 3), reel generation on a 20+ clip reel (Item 1), `ffmpeg` wall time for a reel re-encode with `FFMPEG_VIDEO_ENCODER=auto` vs `libx264` (Item 6).
