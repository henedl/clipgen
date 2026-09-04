# clipgen-profile — Measure performance before optimizing

Perf fixes written from plausible models get disproved by the first real number —
desktop_chrome burned three consecutive fixes that way, and every claim in
[PERFORMANCE.md](../../PERFORMANCE.md) needed external measurement. clipgen now has an
opt-in stopwatch on both halves of the stack. Use it **before** proposing a fix and
**after** to prove it. A perf claim with no number is not reviewable.

## Step 1 — Enable

`--profile` combines with every CLI mode and web launch (like `-v`). It flips
`config.PROFILING`, prints one grep-able line per label at process exit, and is a
strict no-op when off — hot loops pay one boolean check. Unlike `DEBUGGING` it never
changes what work runs.

```
profile | scan.callback                  1.339s  n=962  avg=1.4ms
```

Label glossary — backend: `scan.decode_wait` / `scan.fast_filter` / `scan.callback.<tool>`
(the per-frame split for every Screenspace tool, plus a per-scan summary line
`profile | scan <tool> <file>:`). `scan.callback` without a suffix is only the
no-kind fallback; each `scan_*` / multitool pass sets the tool name so a
workflow of mixed detectors does not lump analysis into one bucket. Callback
flushes pass `peak=` (largest single frame). `ffmpeg.run` / `ffmpeg.bytes`
(every encode/extract subprocess) and `ffprobe.run` (duration / props / keyframe
probes — the label `_parallel_probe` was measured without). `media_cache.*`,
`video.*_cache.*`, and `screenspace.*_cache.*` (hit/miss counters; the last
covers the frame-JPEG, decoded-frame, and pin-OCR LRUs in
`screenspace_server.py`), `worker.progress_lock_wait`, `route <rule>`
(per-route totals; polls aggregate instead of spamming), `stream <rule>` and
`sse.open <rule>` (streaming responses — see below), `transcribe.*`, `sheets.*`
(Google via `_call_with_api_retry`; local `.xlsx` is `sheets.excel_load`),
`pipeline.clip` / `pipeline.pool_wall`, `ocr.pool_wait` / `ocr.reader_build`,
`heatmap.gif` / `heatmap.rolling` / `heatmap.gifs` (pair wall) /
`heatmap.grid_layers` (shared grid accumulation — see below),
`manifest.load <section>` / `manifest.save <section>` (one JSON parse or
dump+write of that `clipgen.json` section, with its size as `bytes=`) and
`manifest.save.unchanged <section>` (saves skipped as identical — a debounced
persister that mostly skips is healthy, a poll path that loads a multi-MB
section every tick is not),
`ollama.generate`, `titlecard.wrap` plus
`titlecard.copy` / `titlecard.reencode` counts (the concat-demuxer vs filter
fallback), `workflows.run` / `workflows.node <type>` / `workflows.batch_child` /
`workflows.batch_wall` (`WORKFLOWS_BATCH_WORKERS` effective parallelism, same
ratio as the clip pools).
Frontend (via `CLIPGEN_CONFIG.profiling`): `poll.<page>.<name>` per poller tick,
`studio.renderGrid`/`renderIntakeCards`/`renderQueue`, `transcripts.renderSegments`/`renderPartialSegments`,
`screenspace.renderResults`/`renderChunk`/`renderTimeline`, `composer.renderTimeline`,
`viewer.renderList`, `workflows.renderAllNodes`, `gallery.renderGrid`,
`overview.renderMetadata`/`computeMetadata`/`renderConvergence`/`computeConvergence`,
and a `longtask` observer for
main-thread stalls >50 ms. To add a span, follow the hooks' pattern: accumulate
into locals and flush once per scan/tick — **never** call `profiling.add` /
`performance.mark` per frame.

**Drilling into one label — `--profile-deep LABEL`.** The stopwatch names the
hot bucket but not the functions inside it, and whole-process cProfile cannot
see into worker threads (a scan's callback runs in `ScreenspaceWorker`, so
`cProfile.run("cli.main()")` shows only lock waits). `--profile-deep
scan.callback.template` (any substring of a label; implies `--profile`)
attaches a per-thread cProfile to exactly the matching spans — every
`profiling.span()` label plus the per-frame scan callback — and appends a
`profile-deep | <label>` pstats block (top functions by tottime) to the exit
report. Two rules: the stopwatch totals of a deep run include cProfile's own
overhead, so never compare them against a plain run; and match narrowly —
matching many labels at once (`--profile-deep scan`) profiles them all into
separate blocks but slows everything that matches. When a match hits both an
outer span and work nested inside it on one thread (`heatmap.gifs` wraps an
inline `heatmap.gif` encode), the outermost profiler wins and absorbs the
nested work — cProfile cannot nest, and the run must never break over it.

**Startup marks — `startup | ` lines.** Launch-to-window attribution for the
desktop/web startup path. `clipgen.py` anchors T0 before any clipgen import;
`profiling.mark()` records milestones *unconditionally* (marks land before
`--profile` is even parsed) but they are only reported under `--profile`: as a
`startup |` section at build-ready and in the exit report, and as `startup` on
`/api/profile` (like `peak_rss`, untouched by `?reset=1`). Key labels:
`startup.imports_utils_cli`, `startup.webview_start` (last main-thread stop
before the window), `startup.window_shown` (first-paint proxy),
`startup.boot_page_alive` (boot page JS executing = content painted),
`startup.silent_google_auth` / `startup.phase_*` (boot-build thread). Marks
from concurrent threads interleave chronologically, so a `+delta` spanning a
thread boundary attributes wall time, not per-thread work. To measure a frozen
launch, read `/api/profile` from the running app — the frozen bootloader
ignores `PYTHONUNBUFFERED`, so grepping a killed process's stdout loses the
buffered `startup |` lines.

Four report tokens are easy to misread:

- **`bytes=`** is a payload sum, present only on labels that carry one:
  `route <rule>` (the `Content-Length` of every non-streamed response),
  `stream <rule>` (characters drained), and `manifest.*` (that section's
  JSON text). It is the second axis a poll goes wrong on — `route` time alone
  hides an endpoint that answers in 2 ms with a 2 MB body every tick. Divide
  by `n=` for the per-response size; `/api/profile` exposes it as `bytes`.
- **`max=`** is the largest single occurrence. It is absent on labels fed only by
  batched flushes, because a batch's `seconds` is a sum with no per-item max. A
  flusher that tracked its own maximum passes `add(..., peak=)` to populate it —
  `scan.callback.<tool>` and `transcribe.decode` do.
- **`first=`** is the cold hit: the label's first single-occurrence
  contribution, printed only when it is at least 5 ms *and* double the average
  of the calls after it. A route whose first call pays a lazy import or a
  cache fill (`/api/models` imports `huggingface_hub`: measured 80-120 ms
  cold, 1.5 ms warm) otherwise reads as a modest `avg` with a large `max`,
  indistinguishable from an occasional slow request. Absent on batched labels,
  which have no per-item first. `/api/profile` exposes it as `first`.
- **`peak_rss`** is process-global, monotonic and POSIX-only (omitted on Windows;
  no psutil dependency). `?reset=1` does not and cannot clear it. `RUSAGE_SELF`
  excludes ffmpeg subprocesses — for clipgen the memory that hurts (Whisper
  weights, OCR engines, decoded frames) is all in-process. It also appears on
  `/api/profile` as `peak_rss_mb`, since the knobs it exists for
  (`SCREENSPACE_OCR_POOL_SIZE`, `WORKFLOWS_BATCH_WORKERS`) are live-server knobs.

## Step 2 — Build a benchmark input

No repo fixture is big enough to measure. Generate one (deterministic; testsrc's
motion defeats phash-skip so every frame reaches the callback), named
`{study}_{participant}.mp4` so participant resolution works:

```bash
mkdir -p /tmp/ssbench
ffmpeg -y -f lavfi -i "testsrc=duration=120:size=1280x720:rate=30" \
    -pix_fmt yuv420p -c:v libx264 -g 30 /tmp/ssbench/bench_P01.mp4
```

Add `-f lavfi -i "sine=frequency=220:duration=120" -c:a aac -shortest` when you
need a Whisper input — the video-only file above has no audio stream and
transcription refuses it before any `transcribe.*` label is recorded.

**The UI fixture is not a benchmark.** `tests/ui/_ui_fixtures.py` builds 6 rows ×
2 participants, so `shot.py studio --perf` reports a `studio.renderGrid` of a few
milliseconds and tells you nothing about the 200×12 case
[PERFORMANCE-PLAN-2](../../../plans/archive/PERFORMANCE-PLAN-2.md) §4.1 named.
For grid / Sheets work, generate a real one — geometry mirrors
`_ui_fixtures._make_workbook`, which is the authoritative layout (`ID` at F2 with
participant columns to its right *on row 2*, `Observation`/`Category` on row 5,
data from row 6):

```python
# /tmp/gridbench.xlsx — 200 rows x 12 participants
import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Observations"
ws["A1"] = "gridbench"
ws["F2"] = "ID"
for i in range(12):
    ws.cell(2, 7 + i, f"P{i + 1:02d}")
for col, h in enumerate(
    ("Count", "Reported", "Severity", "Category", "Observation", "Summary"), 1
):
    ws.cell(5, col, h)
sevs = ("Critical", "Serious", "Moderate", "Minor")
for r in range(200):
    ws.cell(6 + r, 3, sevs[r % 4])  # renderGrid paints .sev-* classes, so an
    ws.cell(6 + r, 4, "Onboarding")  # empty Severity column under-measures it
    ws.cell(6 + r, 5, f"Observation {r}")
    for i in range(12):
        ws.cell(6 + r, 7 + i, "0:01-0:04" if i % 3 == 0 else "")
wb.save("/tmp/gridbench.xlsx")
```

Sanity-check it before trusting any number — a drifted layout yields a silently
small grid, not an error:

```bash
uv run python -c "import sys; sys.path.insert(0,'source'); import excel_io, spreadsheet; \
  print(spreadsheet.build_sheet_context(excel_io.open_excel_workbook('/tmp/gridbench.xlsx')) is not None)"
```

For `transcripts.renderSegments` ([PERFORMANCE-PLAN-3](../../../plans/archive/PERFORMANCE-PLAN-3.md)
§8c gates virtualization on ">2000-segment sessions"), **synthesize the manifest**
— 2000 real Whisper segments is hours of audio:

```python
import json, pathlib

segs = [
    {
        "id": f"P01:{i}",
        "start": i * 3.0,
        "end": i * 3.0 + 2.8,
        "text": f"Synthetic segment {i} for render benchmarking.",
    }
    for i in range(2400)
]
out = pathlib.Path("/tmp/tsbench/clipgen.json")
out.parent.mkdir(parents=True, exist_ok=True)
out.write_text(
    json.dumps(
        {
            "transcripts": {
                "source_transcripts": {
                    "P01": {"segments": segs, "language": "en", "model": "synthetic"}
                },
                "corrections": [],
                "marks": [],
            }
        }
    )
)
```

## Step 3 — Capture a baseline

Run twice; trust only what reproduces (same discipline as
[test-perf](../test-perf/SKILL.md)).

```bash
uv run clipgen.py --ss-task change P01 --ss-threshold 0.05 --ss-interval 0.1 \
    -i /tmp/ssbench -o /tmp/ssbench/out --profile 2>&1 | grep "profile |"
```

`--ss-threshold` (and the tool's other required flags) is load-bearing: without
it `change` / `similarity` / `inactivity` / `flow` / `template` refuse to
build a task, and the report is just `ffprobe.run` — which looks like a
fast-filter skip. **For the whole-sweep comparison, don't hand-roll the
per-tool loop** — `tests/perf/scan_bench.py` owns the canonical flag set per
tool, builds the fixture on demand, wipes each tool's output dir (a cached
manifest hides the callback), and diffs runs as JSON:

```bash
uv run python tests/perf/scan_bench.py --save /tmp/base.json     # baseline
# ...make the change...
uv run python tests/perf/scan_bench.py --compare /tmp/base.json --fail-on 10
```

`--duration` rebuilds when the existing fixture's probed length does not match.
A leftover 120 s `/tmp/ssbench/bench_P01.mp4` would otherwise ignore
`--duration 15` and poison `--compare`. The harness prints `fixture … Ns` at
the start of a run so a wrong-length file is visible before the table.

Template and Shape use a seeded top-left 20%-frame region. This keeps their
reference non-degenerate while measuring bounded correlation instead of a
full-frame reference. Shape is part of the default non-OCR sweep. Drill into
its scale correlations with the same workload:

```bash
uv run python tests/perf/scan_bench.py --tools shape --deep --runs 2
```

`--tools color,text` narrows the sweep (`text` is off by default: OCR is 10×
slower and pins ~0.8 GB RSS per pooled engine — 3.3 GB at the default auto
pool of 4); `--runs 2` keeps the fastest run per tool. For a
single tool the direct CLI form is still useful (unique `-o` dir per run):

```bash
uv run clipgen.py --ss-task color P01 --ss-target-color '#FF0000' \
    --ss-tolerance 20,30,30 --ss-interval 0.1 -i /tmp/ssbench -o /tmp/ssbench/cb-color --profile
```

**The clip pipeline has the same harness** — `tests/perf/clip_bench.py` runs
three scenarios (plain clips, carded clips, a carded reel) against a
generated video+sheet fixture and reports `pipeline.clip`, the pool
parallelism ratio, and the titlecard copy/reencode split, with the same
`--save`/`--compare`/`--runs` treatment:

```bash
uv run python tests/perf/clip_bench.py --save /tmp/clipbase.json
uv run python tests/perf/clip_bench.py --compare /tmp/clipbase.json --fail-on 10
```

`--duration` rebuilds the video *and* the sheet when the existing file's length
does not match, and timestamps stay inside that length (an MM:SS with SS > 59
is dropped at parse). Same `fixture … Ns` line as scan_bench.

Live server: launch with `--profile`, then `curl http://127.0.0.1:8089/api/profile`
(404 without the flag; `?reset=1` snapshots then clears, bracketing a window).

Browser: `uv sync --extra dev --extra ui` (~1 s from cache; `/check` uninstalls the
ui extra), then

```bash
CLIPGEN_UI_CHECK=1 uv run --extra ui python tests/ui/shot.py studio --perf --wait 5000
```

The UI fixture is 6 rows × 2 participants — `studio.renderGrid` will be a few
milliseconds and tell you nothing. Point `--sheet` / `--output` at the
benchmark inputs from Step 2:

```bash
CLIPGEN_UI_CHECK=1 uv run --extra ui python tests/ui/shot.py studio \
    --perf --sheet /tmp/gridbench.xlsx --wait 2000 \
    --eval "return document.querySelectorAll('#sheetGrid tbody tr').length"
CLIPGEN_UI_CHECK=1 uv run --extra ui python tests/ui/shot.py transcripts \
    --perf --output /tmp/tsbench --wait 2000 \
    --eval "return document.querySelectorAll('.segment-row').length"
CLIPGEN_UI_CHECK=1 uv run --extra ui python tests/ui/shot.py screenspace \
    --perf --input /tmp/ssbench --output /tmp/ssbench/out --wait 2000
```

Sanity-check the `--eval` counts before trusting `perf | studio.renderGrid`
(200 rows) or `transcripts.renderSegments` (2400 rows). A drifted sheet
layout yields a silently small grid, not an error.

Each `--perf` run prints `perf | ` lines (CDP layout/script/heap metrics,
navigation/resource timing, the clipgenPerf measures) plus one `perf-json:`
line for parsing; the server's `profile | ` route report follows at exit.
The `perf | api <path> start= dur= bytes=` lines are the **boot waterfall**:
every API fetch in start order. Read them for the shape, not the totals — a
fetch whose `start=` equals another's `start=` + `dur=` is a serial chain
(one round trip the page waits out before issuing the next), and a `dur=`
far above the route's `avg=` in the exit report is a cold hit (see `first=`
above). The static files are deliberately left out; `resources.count` and
`resources.transferSize` cover them, and `--soak` covers what happens after.

**`--input` needs a sheet that matches it.** Participants come from the
sheet's columns first, then disk. With the 6-row UI fixture workbook still
loaded, its `P01` column resolves to the *fixture's* study name and wins over
`bench_P01.mp4` on disk (`has_video: false`), so Screenspace and Composer
never reach their ready selector and `--perf` reports a page that showed no
media. Pass a sheet whose study name matches the videos (`--sheet` with the
clip_bench workbook, or a gridbench workbook renamed to the study), or read
the Transcripts and Overview numbers, which do not need a video.
`--trace /tmp/page.trace.json` writes a Chrome trace for Perfetto — the
"open DevTools" of last resort. The headless shell's paint metrics are only
indicative; add `--full-chromium` when paint fidelity matters.

**Leaks and idle churn — `--soak SECONDS`.** Load-time numbers say nothing
about a page that merely sits open: pollers re-rendering into the DOM,
listeners re-bound per tick, payloads retained per poll. `--perf --soak 20`
re-samples after the page has idled that long and prints `perf | soak.*`
deltas — `soak.Nodes` / `soak.JSEventListeners` / `soak.JSHeapUsedSize`
(growth with nothing happening is a leak; a few nodes for a toast or a clock
is not), `soak.poll.<page>.<name> n=+N` (ticks during the window — check
against the poller's interval, and against 0 for a page that should be
paused), and `soak.longtasks`. Read the server's exit report alongside it:
`route` lines with a large `bytes=` ÷ `n=` are the polls to slim.

```bash
CLIPGEN_UI_CHECK=1 uv run --extra ui python tests/ui/shot.py transcripts \
    --perf --soak 20 --output /tmp/tsbench
```

## Step 4 — Interpret

- `scan.decode_wait` dominates → the scan is decode-bound: look at the fast-scan
  knobs (`SCREENSPACE_FAST_SCAN_SKIP_NONKEY`, GOP/interval) in
  [PERFORMANCE.md](../../PERFORMANCE.md), not the callback.
- `scan.callback` dominates → analysis-bound: per-frame CV work is the target.
  The proven shape is redundant work across consecutive frames — frame N processed
  again at step N+1 (`blur_gray`, `flow_downscale` carries), or the same conversion
  twice in one frame (`compute_phash(gray=)`).
- `*_cache.miss` climbing on repeat requests → a cache key or invalidation bug.
- **Whisper** prints a per-run `profile | whisper <file>:` line —
  `prepare` / `decode` / `callback` plus `audio` / `file` / `vad` / `xrt`.
  `prepare` is *not* overhead: faster-whisper loads audio, extracts features and
  runs VAD + language detection eagerly inside `model.transcribe()` before
  yielding anything, so that is where the `TRANSCRIBE_VAD_*` cost lands. **The
  realtime factor divides audio by `prepare + decode`, never decode alone** —
  measured on a 30 s file, enabling VAD moved `prepare` 0.185 s → 0.641 s while
  `decode` collapsed 1.138 s → 0.000 s, so a decode-only ratio reports *infinity*
  and would endorse the knob no matter what it did. `audio` << `file` means a
  truncated or cancelled run; `vad/file` is the VAD win. A cold HF cache puts the
  model download inside `transcribe.model_load`, so a first run is orders of
  magnitude larger and is not a regression.
- `sheets.*` — the **count** is the invariant, the duration is only the symptom.
  `sheets.get_all_values n=1` per sheet load is `build_sheet_context`'s documented
  "exactly one API call"; anything higher is the redundant-fetch regression
  [PERFORMANCE.md](../../PERFORMANCE.md) opens with, and AGENTS.md warns
  rate-limiting surfaces as silently skipped timestamps rather than an error. A
  large `sheets.backoff_sleep` means throttling, not a slow sheet. A local
  `.xlsx` emits `sheets.excel_load` for the openpyxl read; the adapter's
  in-memory `get_all_values` is unlabelled (it is not an API call).
- `pipeline.clip ÷ pipeline.pool_wall` is **effective parallelism**, the number
  `CLIP_PARALLEL_WORKERS` is otherwise tuned blind against (`ffmpeg.run` times
  each encode but knows nothing about overlap). Both labels are shared by all five
  clip pools — CLI, reel regeneration, and Studio's three — because the knob is
  global. A ratio near 1.0 with several clips queued means the pool is serializing.
- `workflows.batch_child ÷ workflows.batch_wall` is the same ratio for
  `WORKFLOWS_BATCH_WORKERS`. `workflows.node <type>` splits a graph so a slow
  Transcribe node is not mistaken for canvas overhead (`workflows.renderAllNodes`).
- `ffprobe.run` is probe I/O (duration / props / keyframe gap). It is *not*
  folded into `ffmpeg.run` — a reel-validation storm is a probe problem, and
  `_parallel_probe` cannot be proven if the label does not exist.
- `titlecard.copy` vs `titlecard.reencode` counts (not durations) tell you which
  wrap path ran; `titlecard.wrap` is the wall including card encodes. A generate
  that is all `reencode` is the concat-demuxer missing its copy-safe gate.
- `heatmap.gif` / `heatmap.rolling` is post-scan work, not `scan.callback`. A
  drop in callback with an unchanged heatmap total means the CV win did not
  touch GIF encode. The encode itself frame-differences in numpy
  (`_delta_frames`) before handing PIL palette frames with `optimize=False`;
  if `--profile-deep heatmap` ever shows `putdata` / `get_flattened_data`
  again, PIL's own per-frame delta (a Python tuple of every pixel) is back —
  it was 8-10 ms of a 1280×720 frame's ~15 ms. `heatmap.gifs` is the pair wall (same ratio as
  `pipeline.clip ÷ pipeline.pool_wall`): near 2.0 means the cumulative and
  rolling encodes overlapped; near 1.0 means they ran back-to-back.
- `heatmap.grid_layers` is the per-cell circle drawing for a grid tool
  (flow/change/attention), hoisted out of the GIF pair so the PNG and both GIFs
  share one pass instead of replaying the results four times. It is *outside*
  `heatmap.gifs`, so a grid tool's post-scan cost is `grid_layers + gifs` — read
  them together or the total looks like it halved when the work only moved.
  Expect the pair wall to sit near 2.0 once it is hoisted: what was left in the
  threads (numpy folds + PIL's encoder) parallelizes, whereas the circle loop
  it replaced measured **0.82×** — slower in two threads than in one.
- `ocr.pool_wait` is idle time blocked on a busy OCR engine — raise
  `SCREENSPACE_OCR_POOL_SIZE` only when this is large *and* `peak_rss` leaves
  headroom, since each Reader holds its own model copy.
- `route <rule>` totals expose which endpoints actually cost; `poll.*` tick counts
  expose pollers that fail to pause when hidden (a twice-recurring bug class).
- **`route` never includes a streamed body — read `stream <rule>` for those.**
  Flask's `after_request` runs on the `Response` object before the WSGI server
  iterates the generator, so `route /studio/api/generate` reports the time to
  *build* the generator (~0.1 ms) no matter how long the clips take. Verified on
  a real one-clip run: `route` 0.000 s vs `stream` 0.077 s, with `ffmpeg.run`
  0.023 s nested inside. The two are separate families on purpose — a drain's
  cost is server-side job execution, not request handling. Persistent SSE
  channels get `sse.open <rule>` as a **count only**: an `EventSource`'s lifetime
  measures how long a tab stayed open, and it auto-reconnects, so a duration
  would shrink the more broken the stream is. The bounded summary-token stream is
  the exception and *is* timed, because it ends when the run ends.
- **`stream.first <rule>` is the perceived latency of a drain** — time to the
  first chunk, which the `stream` total cannot show: a generate that streams
  its first cached clip at 10 ms and finishes at 40 s and one that sits silent
  for 20 s then floods have the same total. Optimize `stream.first` for
  time-to-first-response (see the two-pass pattern in
  [PERFORMANCE.md](../../PERFORMANCE.md)); `stream` for throughput.
- `longtasks` / `cdp.LayoutCount` → render work; check DocumentFragment batching and
  rAF-throttling per [CODE-REVIEW.md](../../CODE-REVIEW.md).

## Step 5 — Prove the fix

Same input, same command, main vs. branch, both runs reproduced. The optimized
bucket must drop; an untouched bucket (usually `decode_wait`) is your control and
must not move. Any hot-path change needs a bit-identical-output test alongside it
(exact equality, not tolerance — see `test_mean_gray_diff_equals_numpy_mean_absdiff`).
Paste the before/after `profile |` lines into the commit body and PR.
