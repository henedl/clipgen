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

Label glossary — backend: `scan.decode_wait` / `scan.fast_filter` / `scan.callback`
(the per-frame split for every Screenspace tool, plus a per-scan summary line),
`ffmpeg.run` / `ffmpeg.bytes` (every subprocess), `media_cache.*` and
`video.*_cache.*` (hit/miss counters), `worker.progress_lock_wait`, `route <rule>`
(per-route totals; polls aggregate instead of spamming). Frontend (via
`CLIPGEN_CONFIG.profiling`): `poll.<page>.<name>` per poller tick, `studio.renderGrid`,
`screenspace.renderResults`/`renderChunk`, and a `longtask` observer for main-thread
stalls >50 ms. To add a span, follow the hooks' pattern: accumulate into locals and
flush once per scan/tick — **never** call `profiling.add`/`performance.mark` per frame.

## Step 2 — Build a benchmark input

No repo fixture is big enough to measure. Generate one (deterministic; testsrc's
motion defeats phash-skip so every frame reaches the callback), named
`{study}_{participant}.mp4` so participant resolution works:

```bash
mkdir -p /tmp/ssbench
ffmpeg -y -f lavfi -i "testsrc=duration=120:size=1280x720:rate=30" \
    -pix_fmt yuv420p -c:v libx264 -g 30 /tmp/ssbench/bench_P01.mp4
```

## Step 3 — Capture a baseline

Run twice; trust only what reproduces (same discipline as
[test-perf](../test-perf/SKILL.md)).

```bash
uv run clipgen.py --ss-task change P01 --ss-threshold 0.05 --ss-interval 0.1 \
    -i /tmp/ssbench -o /tmp/ssbench/out --profile 2>&1 | grep "profile |"
```

Live server: launch with `--profile`, then `curl http://127.0.0.1:8089/api/profile`
(404 without the flag; `?reset=1` snapshots then clears, bracketing a window).

Browser: `uv sync --extra dev --extra ui` (~1 s from cache; `/check` uninstalls the
ui extra), then

```bash
CLIPGEN_UI_CHECK=1 uv run --extra ui python tests/ui/shot.py studio --perf --wait 5000
```

prints `perf | ` lines (CDP layout/script/heap metrics, navigation/resource timing,
the clipgenPerf measures) plus one `perf-json:` line for parsing; the server's
`profile | ` route report follows at exit. `--trace /tmp/page.trace.json` writes a
Chrome trace for Perfetto — the "open DevTools" of last resort. The headless shell's
paint metrics are only indicative; add `--full-chromium` when paint fidelity matters.

## Step 4 — Interpret

- `scan.decode_wait` dominates → the scan is decode-bound: look at the fast-scan
  knobs (`SCREENSPACE_FAST_SCAN_SKIP_NONKEY`, GOP/interval) in
  [PERFORMANCE.md](../../PERFORMANCE.md), not the callback.
- `scan.callback` dominates → analysis-bound: per-frame CV work is the target.
  The proven shape is redundant work across consecutive frames — frame N processed
  again at step N+1 (`blur_gray`, `flow_downscale` carries), or the same conversion
  twice in one frame (`compute_phash(gray=)`).
- `*_cache.miss` climbing on repeat requests → a cache key or invalidation bug.
- `route <rule>` totals expose which endpoints actually cost; `poll.*` tick counts
  expose pollers that fail to pause when hidden (a twice-recurring bug class).
- `longtasks` / `cdp.LayoutCount` → render work; check DocumentFragment batching and
  rAF-throttling per [CODE-REVIEW.md](../../CODE-REVIEW.md).

## Step 5 — Prove the fix

Same input, same command, main vs. branch, both runs reproduced. The optimized
bucket must drop; an untouched bucket (usually `decode_wait`) is your control and
must not move. Any hot-path change needs a bit-identical-output test alongside it
(exact equality, not tolerance — see `test_mean_gray_diff_equals_numpy_mean_absdiff`).
Paste the before/after `profile |` lines into the commit body and PR.
