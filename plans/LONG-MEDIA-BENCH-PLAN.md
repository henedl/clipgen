# Long-media benchmarks: findings and tool design

## Status

| Unit | Status |
| --- | --- |
| 1. Long fixtures (`make_long_video`, speech seed, `clip_cells`) | Done: 1 h 1080p OBS-like fixture builds in ~3 s (48 MB, screen content) |
| 2. Harness: identity, `--source`, tiers, `xrt`, `--scale`, timeout, overrides | Done. Non-probed keys (`content`, `audio`) live in a hidden `.<video>.spec.json` sidecar; `xrt` is named `media_s`-based so clip_bench can normalize by clip seconds |
| 3a. `scan_bench` session tier + decode-mode column | Done. Label is `scan.decode_mode.{keyframe,full}` |
| 3b. `clip_bench` `--clip-seconds`, re-encode / compress scenarios | Done. `reel-long` dropped: `reel` under `--clip-seconds`/session already cuts multi-minute bodies. Compress cap is 0.1 MB so screen-content clips exceed it |
| 3c. `transcribe_bench.py` | Done (`whisper`, `speakers` scenarios; needs `build/vendor/speakers/speaker_embed.onnx`) |
| 3d. `scale_bench.py` | Done (`cluster`, `redact-map`, `friction`; `--project N`) |
| 3e. `ui_bench` session tier, `composer` + `studio-scrub` scenarios | Done. Both always use a 3 h 1080p long project; `studio-scrub` times the routes via in-page `fetch` (no hover synthesis) |
| 4. Measurement campaign | Done on synthetic fixtures (results below); a `--source` run on a real OBS session is still open |
| Docs | Done: profile SKILL Step 2/3, PERFORMANCE.md, bench docstrings |

Update each row as it lands, with any descoped item or changed decision.

## Context

The perf tooling (`tests/perf/{scan,clip}_bench.py`, `tests/ui/ui_bench.py`, `bench_fixtures.py`) was built and validated against **≤120 s synthetic sources with 3–7 s clips**. Real sessions are 60–180+ min OBS recordings (1080p, keyint ≈250 frames) with multi-minute clips. The suspicion is confirmed: the benches cannot see the costs that dominate real use, and several product paths scale with source or clip duration in ways no bench exercises.

### Findings: what the benches assume vs. reality

| Bench | Assumes | Real sessions | Consequence |
|---|---|---|---|
| `scan_bench` | 120 s, 720p, `-g 30` (1 s GOP), `--interval 0.1` | 1–3 h, 1080p, ~8.3 s GOP, default interval `1.0` | Measures neither real decode volume (×90 frames, ×2.25 pixels) nor the real decode path (below). 1800 s subprocess timeout (`bench_common.py:58`) would kill a 3 h run. |
| `clip_bench` | Clips written as `0:SS`, capped at 0:57 regardless of `--duration` (`clip_bench.py:119-129`); 3–7 s each | Clips anywhere in 3 h, often 1–5 min | Deep seeks, multi-minute copy cuts, card wraps of long bodies, re-encode/compress of long clips are all unmeasured. `--duration 3600` only makes the fixture slower to build. |
| `ui_bench` | Screenspace: 2000 events on the **20 s** 320×240 UI video; Transcripts: 2400 segments (≈2 h at 3 s); no Composer scenario | 3 h ≈ 3600–6000 segments (speakers.py docstring: ~2000/h) | PERFORMANCE-PLAN-3 put the `renderSegments` longtask line at ~4300 segments and closed virtualization on a "30–90 min" assumption — a 3 h session crosses it. Composer's 3 h timeline is never loaded. |
| *(none)* | — | Whisper, speaker clustering, redact run on every long session | No bench at all. `xrt` exists in the profiler but nothing records it. |
| All | Raw seconds only, no normalization | — | A 120 s number cannot be extrapolated; nothing reports seconds-per-source-hour or checks linearity. |

### Product hot paths the benches would expose (measure first, then fix)

1. **Screenspace never uses keyframe-only decode on OBS footage.** Fast skip needs `max_gap ≤ interval × 1.0` (`screenspace_frames.py:316-325`); 8.3 s GOP > 1.0 s interval → every scan full-decodes every 1080p frame of the session. The bench's 1 s GOP + 0.1 s interval also never takes the fast path. Candidate: seek-based sampling (`-ss` per sample or `select` on keyframes + nearest-frame) when interval is long relative to fps.
2. **Speaker clustering is O(U²) memory / O(U³) time** (`speakers.py:259-293`, full-matrix `argmin` per merge). 3 h of fast turn-taking → U≈6000 → 144 MB matrix, ~2×10¹¹ ops.
3. **Whisper/speakers PCM decode holds the whole track twice** (`video.py:2147-2188`): ~1.4 GB peak for 180 min; 900 s ffmpeg timeout.
4. **Studio `/api/sprite` decodes the whole clip span** (`server.py:565-573`, no `seek_frames=True`, which Composer already passes at `composer_server.py:648`); 20 s timeout → multi-minute clips 404.
5. **Studio `/api/clip-audio` is uncapped** (`server.py:587-593`) unlike Composer's `COMPOSER_SCRUB_MAX_AUDIO_SECONDS=180`.
6. **Gallery batch screenshots decode to EOF** (`video.py:3455-3465`, `fps=1/interval`, no `-t`).
7. **Redact span→segment mapping is O(spans × segments)** (`redact.py:616-632`).
8. Smaller: `extract_audio_track` re-encode fallback 300 s timeout (`video.py:2035-2117`); `probe_max_keyframe_gap` only samples the first 20 s.

Clip cutting itself is fine: `-ss` before `-i` + stream copy is O(clip bytes), not O(offset).

## Design

All bench work stays in `tests/perf/` and `tests/ui/` (no product code changes in Phases 1–3). Benches keep their current defaults so existing `--save` baselines remain comparable; long-media runs are opt-in tiers.

### Phase 1: Realistic fixtures (`tests/perf/bench_fixtures.py`)

- **`make_long_video(path, *, duration, size="1920x1080", rate=30, gop=250, audio=True, content="testsrc"|"screen")`**: encode a short seed (a whole number of GOPs, e.g. 6×250 frames = 50 s) with `-g {gop} -keyint_min {gop} -sc_threshold 0` (OBS fixed keyint), then `ffmpeg -stream_loop N -i seed -c copy -t duration`. A 3 h fixture builds in seconds instead of ~30+ min of encoding.
  - `content="screen"`: `testsrc=rate=0.2` upsampled with `fps={rate}`, giving static frames with a step change every 5 s, like a screen recording. This exercises phash-skip realistically. `testsrc` stays the worst case.
  - `audio=True` mixes the sine track; optional **`speech=True`** (macOS `say -o` → AIFF of a fixed paragraph, looped) so Whisper/VAD/speakers produce real segments. On non-darwin, fall back to requiring `--source`.
- Cache long fixtures under `/tmp/clipgen-bench/` keyed by spec; print size on build (a 3 h 1080p fixture is multi-GB; use `-crf 28` on the seed to bound it).
- **`clip_cells(rows, *, duration, clip_seconds)`**: evenly spread `H:MM:SS-H:MM:SS` ranges across the whole source with a fixed clip length. Bench fixtures import no clipgen modules, so it gets its own formatter.

### Phase 2: Bench harness (`tests/perf/bench_common.py` + both benches)

- **Fixture identity**: extend `probe_fixture` / `fixture_matches` with `max_keyframe_gap` (ffprobe packet flags) and `content`. A 1 s-GOP leftover must not pass for an OBS-like one.
- **`--source PATH` (bring your own recording)** on `scan_bench`, `clip_bench`, and the new `transcribe_bench`. The bench symlinks the file as `{study}_P01.mp4` into the work dir, writes a matching sheet, and records identity as `{name, bytes, mtime, probed spec}`. `--compare` refuses a different source. It never copies, uploads or commits the media; JSON stores only the basename.
- **Tiers**: `--tier smoke` (today's defaults, the regression gate) and `--tier session` (3600 s, 1080p, gop 250, `screen` content, interval `1.0`, `--runs 1`). Explicit flags override tier values. The tier lands in `meta.workload` so compare refuses cross-tier baselines.
- **Normalized metrics**: every saved row gains `source_s` and derived `xrt = source_s / elapsed_s` and `s_per_source_hour`, so numbers from different durations can be read side by side.
- **`--scale 300,1800,3600`**: run the same workload at several durations, then `bc.fit_scaling(points)` prints intercept, slope (s per source-hour), and the log–log exponent. Exponent > 1.15 prints `superlinear`, the signal for issues like #2 and #7. Lives in `bench_common` so every bench gets it.
- **Timeout** scales: `max(1800, 3 × source_s)`, override with `--timeout`.
- **Config overrides** for scenarios with no CLI flag (`REENCODING`, `MAX_FILESIZE_MB`, `TRANSCRIBE_SPEAKERS`): `run_clipgen(..., overrides={})` launches `uv run python -c` that sets `config.X` and then calls `cli.main()`. Overrides are recorded in `meta.workload`.

### Phase 3: Coverage additions

- **`scan_bench`**: under `--tier session`, the default tool set and interval match what users run. Add a `fast_path` column (did keyframe skip engage? requires a new `scan.nonkey_skip` count label in `screenspace_frames.py`, the one product touch, a profiling counter only). Print frames decoded vs. sampled so finding #1 is visible as a number.
- **`clip_bench`**: `--clip-seconds N` (spread across the source via `clip_cells`) plus new scenarios `clips-reencode` and `clips-compress` (via overrides), and `reel-long` (multi-minute bodies). The existing `0:SS` generator stays the `smoke` default; its pinned tests stay.
- **New `tests/perf/transcribe_bench.py`**: drives `--pre-transcribe P01 [--speakers]` on the speech fixture or `--source`. Reports `transcribe.*`, the whisper `xrt`, `peak_rss_mb` (finding #3), and the speakers pass time. Same `--save/--compare/--scale` treatment.
- **New `tests/perf/scale_bench.py`**: in-process micro-bench of pure functions over synthetic N, no media: `speakers.cluster` (U = 500…6000 random unit embeddings), redact span mapping, `friction` scoring. Prints `fit_scaling` per function. Fast and good for catching superlinear regressions.
- **`ui_bench`**:
  - Add `--tier session`: 6000 segments at real density, and Screenspace events on a long 320×240 `make_long_video` fixture (event density matched to the real duration).
  - New **`composer`** scenario: 3 h participant with cuts plus the three marker lanes. Time `composer.renderTimeline`, zoom, and the thumbnail-strip/scrub fetches.
  - New **`studio-scrub`** scenario: multi-minute clips with `STUDIO_CARD_SCRUBBER` on. Time `/api/sprite` and `/api/clip-audio` (findings #4, #5) through the existing `/api/profile?reset=1` per-action window.

### Phase 4: Measurement campaign (run, record, decide)

Run the `session` tier, `--scale`, and at least one `--source` real OBS session. Record the numbers in the plan doc. Each product finding (#1–#8) then gets a measured verdict: fix now, fix later, or wontfix with numbers. Fixes are separate PRs, each proven with `--compare`. Pre-scoped likely fixes:
- #4: pass `seek_frames=True`.
- #5: cap with the Composer constant.
- #6: add `-t` or per-frame seek.
- #2: condensed-matrix or nearest-neighbour-chain linkage, or a pre-cap on U.
- #3: stream PCM to a numpy buffer once.
- #1: seek-sampling path when `interval ≫ 1/fps` and the GOP is long.

#### Results (2026-10-10, Apple M3 Pro, 11 cores, 18 GB; synthetic fixtures)

Raw JSON/console output was kept under `.context/bench/` (not committed). Each row is a single run; treat differences under ~10% as noise.

| Measurement | Number | Reading |
| --- | --- | --- |
| `scan_bench --tier session` (1 h, 1080p, 8.3 s GOP, 1.0 s interval), all 9 non-OCR tools | 24–35 s each, 104–152× realtime, **every tool `full` decode**; decode is 50–80% of wall | #1 confirmed: OBS keyframes never qualify for keyframe-only decode. A 3 h session costs ~80 s per tool on screen content |
| `scan_bench --scale 600,1800,3600` (`change`, `color`) | ~27 s per source-hour, exponent 0.95 / 0.99 | Linear; no hidden superlinear step in the scan path |
| `clip_bench --tier session` `clips` (20 × 180 s spread over 1 h) | 0.93 s total | Stream-copy cutting is free at any length |
| `clip_bench --scale 600,3600,10800 --clip-seconds 60` | 0.76 → 0.72 → 0.86 s (exponent 0.16) | Flat: seek depth does not matter (`-ss` before `-i`) |
| `clips-reencode` (60 min of clip content) | **509 s**, 7.1× realtime, parallelism 4.0 | Re-encode is the long-clip cost centre; 20 three-minute clips ≈ 8.5 min |
| `clips-compress` (two-pass, libx264) | **253 s**, 14.2× realtime | Second-largest; always software by design |
| `reel` | invalid locally: ffmpeg without `drawtext` | Needs a run on a `drawtext` build |
| `transcribe_bench` smoke (5 min speech, `base`) | 24 s, 12.3× realtime, **805 MB** peak RSS | |
| `transcribe_bench --tier session` (1 h speech) | 286 s, 12.6× realtime, **2550 MB** peak RSS; speakers pass +15 s (cluster 0.17 s at 757 segments) | Time is linear. **Memory grows ~32 MB per audio-minute**, ~4× the PCM buffer (#3). See the 3 h row |
| `transcribe_bench --duration 10800` (3 h speech, `whisper` only) | 870 s, 12.4× realtime, 2287 segments, **7794 MB** peak RSS | #3 confirmed at full length: 805 MB → 2.55 GB → 7.8 GB for 5 min → 1 h → 3 h. That is ~44 MB/min over the last two hours, above the linear extrapolation (6.4 GB). A 3 h session can exhaust an 8 GB machine |
| `scale_bench` `cluster` | 0.37 s @1000, 3.0 s @2000, 24 s @4000, **81.8 s @6000** (projected 81.4 s); exponent 3.01 | #2 confirmed cubic. Real speech density (~2000 segments/h) puts 3 h of fast turn-taking at ~80 s and a 144 MB matrix. The speech fixture yields 757 segments/h, so `transcribe_bench` under-states this |
| `scale_bench` `redact-map` | 0.07 s @2000, projected 0.66 s @6000; exponent 2.0 | #7 quadratic but small: wontfix until sessions exceed ~20k segments |
| `scale_bench` `friction` | exponent 1.04 | Linear control |
| `ui_bench --tier session` (3 h project) | transcripts 6000 segments: search 364 ms, restore 155 ms; screenspace drain 876 ms; composer zoom-in 8 ms, reset 48 ms; **sprite-30s 307 ms vs sprite-180s 1187 ms**; audio 95 / 144 ms | #4 confirmed O(clip) on Studio's sprite route (Composer's seek path is O(frames)); #5 audio is cheap on this fixture. No UI page is in trouble at 3 h |

**Verdicts**

- **#3 Whisper memory — fix first.** Peak RSS reaches 7.8 GB on a 3 h session, ~32–44 MB per audio-minute, far above the 7.7 MB/min the PCM buffer and its copy explain. The likely cause is faster-whisper computing the log-mel STFT over the whole array up front; check that with `tracemalloc` before designing. Candidate fix: transcribe in bounded windows (e.g. 20 min with overlap) and stitch.
- **#2 speaker clustering — fixed.** Nearest-neighbour-chain average linkage, cut to the same merge count as the greedy loop's stop rule. Identical labels on 4800 randomized cases against the old loop (kept as the test oracle). `scale_bench --only cluster`: 6000 units 81.8 s → 0.48 s, exponent 3.01 → 2.19. The O(n²) distance matrix (144 MB at 6000) remains.
- **#1 Screenspace full decode — fix later, measured smaller than feared** on screen content (~27 s per source-hour). Re-measure with `--source` on a real OBS session (webcam overlays and motion decode slower). Candidate fix: a seek-sampling path, or relax the keyframe gate when `interval ≥ 1 s` and accept snapping to the nearest keyframe.
- **#4 Studio sprite — fixed.** A blanket `seek_frames=True` would have regressed short clips: on 8.3 s GOPs each seek decodes up to a GOP, so the seek path has a flat ~1.0 s floor, while single-pass costs 159 ms (10 s) → 1.09 s (180 s) → 3.5 s (600 s). Studio now seeks only at `STUDIO_SCRUBBER_SEEK_SECONDS` (120 s) and up. `ui_bench --scenarios studio-scrub`: sprite-30s 313 ms (single-pass), sprite-600s 1088 ms (was 3.5 s).
- **#5 clip-audio cap — low priority.** 144 ms for 180 s; the cost is browser-side decode of a 4 MB WAV. Cap it for parity with Composer.
- **#6 gallery batch to EOF**, **#8 timeouts** — not measured here. Add a `gallery` scenario to `clip_bench` when someone touches them.
- **#7 redact mapping** — wontfix (numbers above).

### Docs

- `agents/skills/profile/SKILL.md`: Step 2 gets the long-fixture recipe, tiers, `--source`, `--scale`, and how to read `xrt` / `s_per_source_hour`. Fix the docstring claim at `clip_bench.py:18` ("3-9 s" → actual).
- `agents/PERFORMANCE.md` line 5: one sentence that the smoke tier is a regression gate, not a realism claim.
- Save this plan as `plans/LONG-MEDIA-BENCH-PLAN.md` with a status table, kept current per unit, and archived when done.

## Critical files

- `tests/perf/bench_fixtures.py`, `bench_common.py`, `scan_bench.py`, `clip_bench.py`; new `transcribe_bench.py`, `scale_bench.py`
- `tests/ui/ui_bench.py` (new tiers and scenarios), reusing `_ui_fixtures` / `bench_fixtures`
- `source/screenspace_frames.py` (one `profiling.add("scan.nonkey_skip", …)` counter only)
- Unit tests: `tests/perf/test_{scan,clip}_bench_parse.py`, `tests/test_perf_bench.py` (extend for `fit_scaling`, the new identity fields, `clip_cells`, tier/compare refusal, timeout scaling); new `tests/perf/test_transcribe_bench_parse.py`

Reuse: `bc.run_clipgen` / `aggregate` / `compare` / `meta_diff`, `bf.write_sheet(cell=)`, `bf.write_manifest`, `transcripts_section`, `screenspace_section`, and the profiler's existing `xrt` and `peak_rss_mb`.

## Verification

- `uv run --extra dev pytest -c tests/pytest.ini tests/perf tests/test_perf_bench.py`: parse/compare/identity/fit tests pass (mocked ffmpeg, fast).
- `uv run python tests/perf/scan_bench.py --save /tmp/a.json && … --compare /tmp/a.json`: smoke tier unchanged and still compares against a pre-change baseline.
- `scan_bench.py --tier session --tools color,change` builds a 1 h OBS-like fixture in under 1 min, reports `xrt`, `fast_path=no` (confirming #1), and finishes without timeout.
- `clip_bench.py --tier session --clip-seconds 180`: 20 clips spread across the hour, all valid.
- `scale_bench.py`: `speakers.cluster` reports an exponent of about 2–3 (confirming #2).
- `transcribe_bench.py --source <real OBS session>`: valid row with `xrt` and `peak_rss_mb`.
- `CLIPGEN_UI_CHECK=1 uv run --extra ui python tests/ui/ui_bench.py --tier session --scenarios transcripts,composer`: DOM counts match the fixture. Read the screenshots.
- `/check` before commit. This is a `test:`/`chore:` PR, so no version bump.
