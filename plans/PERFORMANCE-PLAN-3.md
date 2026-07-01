# Performance Review & Plan — clipgen (PERFORMANCE-PLAN-3)

## Context

This is a thorough review of where clipgen can gain performance — real speedups *and* perceived-responsiveness improvements (warm-ups, preloads, wait coverage). Two prior passes (`plans/archive/PERFORMANCE-PLAN.md`, `PERFORMANCE-PLAN-2.md`) landed nearly everything they identified, so this review focused on (a) explicitly deferred wins, (b) gaps those passes didn't cover, and (c) perceived-latency polish. Every finding below was **verified against the current code** (many candidate findings from exploration were rejected as already-done or wrong — see "Rejected" at the end, kept so future passes don't re-propose them).

**Execution model (per the maintainer):** the workspace that authored this plan produces the plan document only, so other workspaces can each pick up one item. Each item below is a self-contained work packet: one small PR, `feat:` bumps `build/VERSION` patch, run the `/check` skill before commit, update this plan's status column as items land (repo hard expectation).

**Decision:** hardware encoding defaults to **`"auto"`** (speed-first; runtime fallback to libx264).

## Status

| # | Item | PR type | Effort | Status |
|---|------|---------|--------|--------|
| 1 | Parallel ffprobe loops | feat | S | ☐ |
| 2 | SSE-primary / poller-fallback fix | fix | S | ☐ |
| 3 | Titlecard concat-demuxer stream copy | feat | L | ☐ |
| 4 | Ollama model prewarm during transcription | feat | S–M | ☐ |
| 5 | Bounded frame-0 preload (Screenspace) | feat | S | ☐ |
| 6 | Hardware video encoding (VideoToolbox) | feat | M | ☐ |
| 7 | Skeleton coverage for initial loads | feat | S | ☐ |
| 8 | Small backlog (each optional, S) | mixed | S | ☐ |

Recommended landing order = table order: #1 lands `video.py` groundwork before #3 touches the same file; #6 lands after #3 so the encode call-site list is stable.

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

## Item 2 — SSE-primary / poller-fallback fix (`screenspace-tasks.js`) — latency + redundant requests

Verified behavior: SSE and the 3s poller do **not** run concurrently at boot. The real gap: after an SSE drop the poller starts, and a later `startSSE()` (from queueing, results, or the `visibilitychange` handler at screenspace-tasks.js:1143) opens a new stream while the poller keeps running → both transports run concurrently from then on. Also `startSSE()` passes no `onUnsupported`, so no-EventSource environments get zero task updates.

- In `startSSE()` (screenspace-tasks.js:1062–1079): add `onUnsupported: startPolling` (`createSSEStream` supports it, utils.js:931) and call `stopPolling()` in `onOpen` next to the existing `state.sseFellBack = false`. workflows-runs.js already implements this pattern — mirror it.
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

## Item 4 — Ollama model prewarm during transcription — perceived win

Today the summary agent's first `generate()` after transcription pays the 5–30s Ollama model load while the user watches "generating…". Page the model in while Whisper is still running.

- **`ollama_client.prewarm_model(model, keep_alive="10m") -> bool`**: same request shape as `unload_model` (ollama_client.py:129) but positive keep_alive, ~120s timeout, never raises, does **not** auto-start `ollama serve`.
- **`transcripts.TranscriptWorker.on_task_start`** callback (symmetric to existing `on_task_complete`, fired at the RUNNING transition, transcripts.py:953–963, exception-swallowed).
- **`transcripts_server._maybe_prewarm_agent_models()`** wired in `_init_transcripts_state`: gates in order — `config.OLLAMA_PREWARM_ENABLED`; first enabled agent via `_agent_enabled` (1603); model via `_agent_model` (134); TTL debounce (~240s, under keep_alive/2); then daemon thread: `is_available()` (never start the server for a prewarm) → model installed (never trigger a download) → `_cancel_pending_unload(model)` → `prewarm_model`. Mirrors the `api_transcribe_warmup` guarded-thread pattern (1216–1227). Server-side only, no frontend work.
- **Config**: `OLLAMA_PREWARM_ENABLED: bool = True` + `SETTINGS_DESCRIPTIONS` (call out RAM: LLM resident while Whisper runs; never downloads) + `STUDIO_SETTINGS` (tab Summaries).
- **Tests**: `test_ollama_client.py` request shape/failure; `test_transcripts.py` callback fires on RUNNING; `test_transcripts_api.py` prewarm called once on POST /api/transcribe, suppressed by knob-off / agents-disabled / not-installed / TTL.

## Item 5 — Bounded frame-0 preload (Screenspace boot) — perceived win

`screenspace.js` boot (~3745–3758) fires an **unbounded** fetch of frame 0 for every participant simultaneously, competing with the selected participant's own first-frame render (server has a 3-slot VideoCapture pool).

- Reorder the `apiGet("api/participants")` handler: compute `pickId` (localStorage-restore block currently *below* the loop) **first**, call `selectParticipant(pickId, initialTs)` so its request wins, then pump a bounded queue (selected first, then the rest) at concurrency 2 — the `studio-scrubber.js` prefetch-queue pattern (lines ~27–52). Keep `_videoVersions` seeding above everything (`?v=` cache-bust contract) and the existing `TODO: apiGetBlob` comment. ES5 note: release the slot in a final `.then()` after `.catch()` so it always decrements.
- No config, no automated tests (no JS runtime harness) — manual browser verification: DevTools Network shows ≤2 concurrent `frame/*/0` requests, selected participant first. Per AGENTS.md, ask the maintainer to browser-check.

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

## Item 7 — Skeleton coverage for initial loads — perceived polish

Studio sheet (`populateSheetSkeleton`, studio.js:1082) and the Screenspace frame (`.skeleton-frame`) already have skeletons. Verified gaps show *misleading* empty states during the first fetch:

- **transcripts.html:40**: `#participantPills` ships `No participants` before the fetch resolves → replace with 3–4 `<span class="skeleton pill-skeleton"></span>` placeholders (`.pill-skeleton` in transcripts.css, sized off tokens); `renderPills()` (transcripts-pills.js:97) rewrites innerHTML on first data so the real empty state still appears — post-fetch only.
- **transcripts.html:159–162**: `#transcriptEmpty` ships "No transcript available" pre-data → add `hidden` in HTML, unhide in the hub's no-participants path.
- **workflows.html:31/35**: `#wfPalette` and `#wfStashList` are blank until catalog/blueprints resolve → seed ~6 / 2 skeleton rows (page CSS); `renderPalette()` and the stash renderer already replace container contents. Zero JS changes. Leave `#wfRuns`/`#wfCanvasEmpty` (intentional empty states).
- Reuse only the `.skeleton` shimmer token (tokens.css:399); no `buildSkeletonGrid` (these are pills/rows, not tables).
- **Tests**: DOM-wiring/source assertions (`tests/test_transcripts_dom_wiring.py`, `tests/test_workflows_frontend_source.py`) that the skeleton classes exist and `#transcriptEmpty` carries `hidden`.
- **Risk**: a failed fetch leaves shimmer instead of a message — acceptable; the fetch `.catch` already toasts.

## Item 8 — Small backlog (optional; verified but low individual impact)

- **Lazy `import gspread`** in clipgen.py:21 → first Google-Sheets use. Measured: gspread is ~200ms of the ~260ms module import (`-X importtime`); Excel-only and `--help` flows never need it. `fix:`/`perf:`-class change; check `ty` on the deferred-import pattern.
- **TTL cache for `google_api.get_all_spreadsheets`** (server.py:3411, 3498): module-level list + timestamp, ~300s TTL, refresh param — settings opens stop paying a Sheets round-trip (and its 429 backoff risk).
- **Transcripts segment list**: `renderSegments()` (transcripts.js:812–896) rebuilds one big innerHTML string per render. Cheap wins: `content-visibility: auto` (+ `contain-intrinsic-size`) on segment rows in transcripts.css; preserve scroll position across rebuilds. Virtualize only if profiling shows >2000-segment sessions hurting.
- **Gate always-on pollers**: transcripts xref poller (transcripts.js:119, 30s) could start lazily on first xref use; Studio's four intake pollers (studio.js:4409–4413, 5–10s) could pause while the intake panel isn't the active tab (`createPoller` only gates on `document.hidden`).
- **`decoding="async"`** on dynamically created thumbnails/imgs (queue cards, results) — trivial, prevents main-thread decode jank.
- **Studio grid filter re-render** (studio.js:1432 region): known deferred item (PERF-PLAN-2 §4.1), still gated on profiling — profile before building incremental filtering.

## Rejected findings (verified against code — do not re-propose)

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

For each item: the per-item test plans above, plus `/check` (ruff format + lint, ty, full suite) before every commit; browser checks (Items 5, 7, and the Settings UI of 4/6) are manual — ask the maintainer to verify in their browser per the no-headless-browser rule; benchmark notes worth capturing in PR bodies: titlecard wrap before/after on a ~60s clip (Item 3), reel generation on a 20+ clip reel (Item 1), `ffmpeg` wall time for a reel re-encode with `FFMPEG_VIDEO_ENCODER=auto` vs `libx264` (Item 6).
