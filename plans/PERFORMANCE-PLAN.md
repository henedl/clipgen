# Performance Plan

Strategies and experiments to improve clipgen's real and perceived performance across the CLI, web UIs, and CI pipeline.

Topics are sorted roughly by expected user impact within each category.

---

## Prior Work & Proven Directions

16 commits in the git history already address performance. The table below maps each to our planned experiments — showing what's landed, what has further runway, and patterns worth replicating.

### What's already landed

| Area | Commit | What was done | Related experiment |
|------|--------|--------------|-------------------|
| **Frontend rendering** | `89bd136` (#105) | RAF-throttle canvas ops, cache `getBoundingClientRect()` and `getComputedStyle()`, split playhead to separate canvas layer, DocumentFragment batching, skip re-render on unchanged poll data, debounce intake search (250ms), pause polling when tab hidden | 1E (DOM batching) — Studio grid; Viewer list + Screenspace lists also use fragments now; Insights Builder still a candidate (see §8) |
| **Frontend rendering** | `18607ae` (#75) | Decouple timeline playhead from frame loading, coalesce frame requests, RAF-throttle frame requests, cache VideoCapture and computed styles | 1A, 1E — foundational work |
| **Video preview** | `8206470` (#103) | Preload clips as muted paused `<video>`, RAF-throttled scrub, 60ms debounce on hover | 1B (preloading) — done for viewer clips, not Screenspace frames |
| **Sprite scrubbing** | `444d129` (#106) | Replaced canvas sprite system with direct `<video>` seek — simpler and faster, net -65 lines | Shows simplification > optimization |
| **Screenspace analysis** | `253d592` (#67) | Sequential grab/retrieve for intervals ≤3s, pHash pre-filter for similarity, skip OCR on unchanged frames, downsample regions (color→64px, SSIM→256px), cache `probe_video_properties()` | 2A (adaptive resolution) — partially done; 2C (universal phash) — similarity-only so far |
| **Screenspace streaming** | `910df4a` (#71) | Stream results during execution, cache EasyOCR readers, pre-resize similarity refs, cache preprocessed frames, static-frame skip in similarity | 1A (streaming) — results stream exists but frontend still polls |
| **Startup** | `8e85ac0` (#100) | Reduce `build_sheet_context()` from 7→1 API calls (saved 3-6s), defer cv2/screenspace/gspread imports | 4A (lazy imports) — partially done |
| **Sprites** | `c4f2541` (#13) | On-demand sprite generation, in-memory cache | Already optimized |
| **Gallery/filmstrip** | `c5b604a` (#76) | Batch screenshots in single ffmpeg pass (4-6x faster), parallel GIF extraction with ThreadPoolExecutor (4 workers, 3-4x faster), increase filmstrip concurrency 1→3 | 1F — later raised to 4 concurrent loads in viewer.js; 3A (parallel cutting) — proven pattern; Studio multi-cell path still sequential (see §8.2) |
| **Artifact reuse** | `cbee9eb` (#26) | Skip regeneration when output file exists, seed session from manifest | Reduces redundant work — same principle as 3D (transcript caching) |
| **Region reuse** | `2f70455` (#52), `0e7bb96` (#90) | Normalized coordinates (0-1 fractions), region stashing and restore | Enables cross-resolution analysis — relevant to 2A |
| **CI** | `b253685` (#79) | Path filters to skip tests on non-Python changes, scope push trigger to master | 5C (path-scoped tests) — partially done at workflow level |
| **CI** | `0f3da15` (#108) | CPU-only torch in CI, saves ~2.5GB download | Already done |
| **Thread safety** | `370b3e4` | Threading lock around VideoCapture seek+read | Prerequisite for 2F (parallel regions) |

### Proven patterns to extend

**1. RAF-throttling and cache-on-interaction** (3 commits, all kept)
The pattern of caching expensive DOM/style lookups at interaction start and clearing on interaction end has been applied to Studio grid and Screenspace canvas. **Insights Builder** is still a good place to extend the same pattern where heavy lists repaint often.

**2. Sequential frame reading** (`253d592`)
The switch from random-access seeking to sequential `grab()/retrieve()` for intervals ≤3s (`SCREENSPACE_SEQUENTIAL_READ_MAX_INTERVAL`) was a significant win because H.264 keyframe decoding is expensive. Our experiment 2B (variable intervals) should preserve this: the coarse pass at 3-5s intervals still falls within sequential-read range, so it gets both the content-skip benefit *and* the decode-path benefit.

**3. pHash pre-filter** (`253d592`, `910df4a`)
Already proven for similarity scans. Experiment 2C (universal phash) extends this to all tool types. The infrastructure (phash computation, threshold constant) is already in place — this is primarily a wiring change.

**4. Downsampling before analysis** (`253d592`)
Color analysis already downsamples to 64px, SSIM to 256px. Experiment 2A extends this by halving these targets in fast scan mode. The downsampling code paths exist; the change is parameterizing them by scan mode.

**5. ThreadPoolExecutor for ffmpeg** (`c5b604a`)
Parallel GIF extraction with 4 workers achieved 3-4x speedup. Experiment 3A (parallel clip cutting) replicates this exact pattern for `cut_clip()`. The `concurrent.futures` import and worker config (`GALLERY_PARALLEL_WORKERS`) are already in video.py.

**6. Skip-if-exists** (`cbee9eb`)
Studio already skips artifact regeneration by checking output files. Experiment 3D (transcript caching) applies the same principle to transcription — check for a sidecar file before re-running faster-whisper.

**7. Result streaming exists but is under-leveraged** (`910df4a`)
The backend already streams results during analysis, but the frontend polls at 3s intervals to pick them up. Experiment 1A (SSE) would close this gap — the backend already produces events incrementally, so the main work is the transport change.

### Directions that didn't pan out

**Sprite sheet scrubbing** (`0d13472` → `444d129`): Canvas-based sprite scrubbing was built and then replaced one commit later with direct `<video>` element seeking. Lesson: the browser's native video decoder is hard to beat for scrubbing UX. Keep this in mind for any future thumbnail/preview work — use `<video>` elements over custom sprite rendering when possible.

---

## 1. Perceptual Performance

How fast clipgen *feels* to the user, independent of actual processing time.

### - [x] 1A. Streaming progress for Screenspace analysis tasks

**Prior art:** Backend already streams results during analysis (`910df4a`). Frontend polls every 3s (increased from 2s in `89bd136`) with skip-if-unchanged fingerprinting and tab-hidden pause.

**Current:** The frontend polls `/api/tasks` every 3 seconds. Results appear in a batch when a poll lands after new events are written. Long analyses (e.g. 10-minute video at 1s interval = 600 frames) can feel stalled between polls.

**Experiment:** Switch from polling to Server-Sent Events (SSE). The worker thread would push each new event as it's produced, giving the UI a live heartbeat. The poll interval becomes irrelevant; results trickle in continuously. The backend already produces events incrementally — main work is the transport change.

**Trade-offs:** SSE adds a persistent connection per client. Flask's dev server handles this fine for single-user; production would need `gunicorn --worker-class gevent` or similar. Fallback: shorten poll interval to 500ms (cheap since `/api/tasks` is a dict read).

### - [x] 1B. Preload first frame when a video is selected in Screenspace

**Prior art:** Viewer already preloads clips as muted paused `<video>` elements (`8206470`). Same principle, different context.

**Current:** Selecting a participant fetches `/api/video/frame/<participant>/0` on demand. There's a visible blank canvas until the JPEG arrives.

**Experiment:** When the participant list renders, fire a background `fetch()` for frame 0 of each participant and cache the blob. When the user clicks, the image is already in memory. For 5 participants this is ~5 small JPEGs (~50-100 KB total).

**Trade-offs:** Negligible. Tiny network cost, large perceptual win.

### - [x] 1C. Optimistic UI updates in Studio

**Current:** Generate actions in Studio stream progress via ndjson (`/api/generate`), but the clip list only refreshes after the stream closes. The user waits for the full batch to finish before seeing new artifacts.

**Experiment:** As each ndjson line reports a completed clip, append it to the artifact list immediately. If the batch is later cancelled, remove the optimistically-added items.

**Trade-offs:** Requires handling partial state on cancel/error. Low risk since Studio already parses the ndjson stream line-by-line.

### - [x] 1D. Skeleton loading states for web UIs

**Current:** Panels render empty, then fill. In Studio and Insights Builder, the clip/insight lists pop in after the initial API fetch completes.

**Experiment:** Show skeleton placeholder rows (grey pulsing bars) until data arrives. This is a standard perceptual trick that makes sub-second loads feel instant.

**Trade-offs:** Pure CSS/HTML, no backend change. Small effort.

### - [x] 1E. Frontend DOM batching for large artifact lists

**Prior art:** DocumentFragment batching already applied to Studio grid rendering (`89bd136`).

**Status (audit):** `renderList()` in [assets/web/viewer.js](../assets/web/viewer.js) and results rendering in [assets/web/screenspace.js](../assets/web/screenspace.js) now build a `DocumentFragment` and append once. **Remaining candidate:** [assets/web/insights-builder.js](../assets/web/insights-builder.js) artifact grid and insight cards still append per item in a loop — apply the same fragment pattern (or virtual scrolling for very large lists; pairs with 6A).

**Experiment:** Build cards inside a `DocumentFragment`, then append the fragment in a single DOM write. For 200+ artifacts this eliminates hundreds of reflows.

**Trade-offs:** None. Strictly better for the Insights Builder lists still on the per-append path.

### - [x] 1F. Filmstrip thumbnail concurrency

**Prior art:** Filmstrip concurrency was increased from 1→3 in `c5b604a`, then raised further in viewer.js.

**Status (audit):** `FILMSTRIP_CONCURRENCY = 4` in [assets/web/viewer.js](../assets/web/viewer.js). Optional next step: adaptive ramp (2→6) or `meta` override for slow links.

**Experiment:** If needed, increase to 5–6 concurrent loads or use an adaptive strategy (start at 2, ramp up if loads complete quickly).

**Trade-offs:** Marginal risk of saturating a slow connection. Could make configurable via `window.CLIPGEN_DATA.meta`.

---

## 2. Screenspace Tool Performance

The heaviest computation in clipgen. Several of the experiments below trade result precision for speed — these are best offered as an explicit **"Fast Scan" mode** that the user opts into, rather than silently degrading quality. The UX contract: fast scan gives you quick, approximate results to orient yourself; if something looks interesting, re-run at full fidelity.

### - [x] 2.0. Fast Scan mode — design concept

A toggle (button or dropdown next to the Run button, or a config flag `SCREENSPACE_FAST_SCAN`) that bundles several of the strategies below into a single user-facing choice:

| Lever | Normal | Fast Scan |
|-------|--------|-----------|
| Frame interval | 1.0s (uniform) | Coarse-to-fine: 3s initial, refine to 1s on change |
| Analysis resolution | Per-tool defaults (64-256px) | Halved across the board |
| Static frame skip | Similarity-only phash | Universal phash pre-filter |
| Template matching | Full resolution | 2x downscale |

When fast scan completes, the results panel could show a subtle label ("Fast scan — re-run for full detail") so the user always knows what they're looking at. A "Re-run full" button next to it makes the upgrade path obvious.

The individual levers are described below. Any of them could also be adopted independently (always-on) if testing shows the quality trade-off is negligible.

### - [x] 2A. Adaptive resolution: low-res scan, high-res confirm

**Prior art:** Downsampling already exists per-tool (`253d592`): color→64px, SSIM→256px, optical flow→256px, scene fingerprint→128px. These were introduced as fixed targets.

**Current:** Each analysis type has its own fixed downscale target (color: 64px, similarity/optical-flow: 256px, template matching: full resolution). These apply uniformly to every frame.

**Experiment:** In fast scan mode, halve each tool's downscale target (color: 32px, similarity/flow: 128px, template: half-frame). When a frame exceeds a threshold at low-res, optionally re-read and re-evaluate at full resolution to confirm.

Concrete example for change detection:
- Normal: blur + grayscale + absdiff at full region size every frame.
- Fast scan: downsample region to 128px wide, run the same pipeline. If change ratio > threshold * 0.8 (slightly below the real threshold), re-evaluate at full resolution.

**Trade-offs:** Adds ~10% overhead for frames that trigger (double read + re-evaluate). Net win depends on what fraction of frames are uneventful — for most videos, >90%. In fast scan mode without re-confirmation, some borderline events may be missed, but the user has opted into that trade-off.

### - [x] 2B. Variable frame intervals (coarse-to-fine)

**Current:** Fixed interval (default 1.0s, config.py:127). Every frame is sampled uniformly regardless of content.

**Experiment:** In fast scan mode, start with a coarse interval (e.g. 3-5s). For each interval where the analysis detects a significant change, subdivide that interval and re-scan at finer granularity (e.g. 1s, then 0.5s). Static periods are skipped quickly; transitions are resolved precisely.

This is particularly valuable for long videos (30+ minutes) where large stretches may be idle (e.g. loading screens, static menus).

**UX consideration:** Results arrive in two phases — coarse pass first, then refinements. The progress bar and results list need to reflect this (e.g. "Pass 1/2: coarse scan... Pass 2/2: refining 12 regions"). This is actually a perceptual *win* — the user sees approximate results quickly and watches them sharpen, rather than staring at a linear progress bar.

**Trade-offs:** More complex scan logic. The two-phase UX needs thoughtful design. In normal mode, the current uniform interval is kept unchanged.

### - [x] 2C. Universal phash pre-filter for static frames

**Prior art:** pHash pre-filter introduced for similarity scans in `253d592`. Static-frame skipping added for similarity in `910df4a`. OCR skip-on-unchanged also uses a fast mean pixel diff (~0.1ms vs ~500-1000ms OCR). Infrastructure exists — threshold constant `SCREENSPACE_PHASH_THRESHOLD`, phash computation, hamming distance check.

**Current:** phash-based static frame skipping exists for similarity analysis (screenspace.py:825-843, threshold: `SCREENSPACE_PHASH_THRESHOLD = 15`) but is not used by other tool types.

**Experiment:** Apply the phash check as a universal pre-filter across all tool types. Before running any analysis on a frame, compare its phash to the previous frame. If the hamming distance is below threshold, skip the frame entirely. This avoids the cost of the actual CV operation (blur, SSIM, optical flow, etc.) for frames that are perceptually identical.

This could be always-on (not just fast scan), since phash is very cheap (~0.1ms per frame) and skipping identical frames is almost always correct. The pattern is already proven for similarity and OCR — this is primarily a wiring change to apply it universally.

**Trade-offs:** Risk: aggressive skipping could miss subtle changes that specific tools are designed to detect (e.g. a small counter incrementing in one corner while the rest of the frame is static). Mitigate by computing phash only over the analysis region, not the full frame. Use tool-specific thresholds if needed (tighter for change detection, looser for color analysis).

### - [x] 2D. Template matching at reduced resolution

**Current:** Template matching runs at full frame resolution with a blur kernel of 3. `cv2.matchTemplate()` is O(W*H) per template — expensive for full-HD frames.

**Experiment:** In fast scan mode, downscale both frame and template by 2x before matching. The confidence threshold may need slight adjustment (lower, since features blur), but localization accuracy at 960x540 is still sub-pixel when mapped back to the original.

**Trade-offs:** 4x speedup for the most expensive per-frame operation. Slightly reduced precision for small template features (<20px). Could also be adopted as the default for the initial pass in normal mode, with full-res confirmation for frames that match.

### - [x] 2E. Batch frame extraction via ffmpeg instead of cv2

**Current:** Frames are extracted one-at-a-time via `cv2.VideoCapture.grab()/retrieve()` or seek-based access (screenspace.py:457-496).

**Experiment:** For analysis tasks with known intervals, use a single ffmpeg command to extract all frames as a batch (similar to the gallery's `_batch_extract_screenshots()`). Write frames to a temp directory or pipe, then process them. ffmpeg's decoder is often faster than OpenCV's for H.264, especially with hardware acceleration.

**Trade-offs:** Requires temp disk space for frames (or memory for pipe). Loses the ability to skip frames mid-scan (e.g. phash skip). Best combined with the coarse-to-fine approach: batch extract at coarse intervals, then targeted extraction for refinements. This optimization is independent of fast scan mode — it's a backend improvement that could apply always.

### - [x] 2F. Parallel analysis of independent regions

**Prior art:** Threading lock around VideoCapture seek+read was added in `370b3e4` to fix concurrent access crashes. This confirms that concurrent VideoCapture usage needs careful synchronization — each thread needs its own capture object.

**Current:** The `ScreenspaceWorker` processes tasks sequentially in a single thread (screenspace.py:2098-2232).

**Experiment:** When multiple regions are defined, analyze them in parallel using a thread pool (or process pool for CPU-bound CV work). Each region's frames are independent. This scales linearly with CPU cores.

**Trade-offs:** Memory usage scales with parallelism (each thread holds a frame buffer). `cv2.VideoCapture` is not thread-safe (learned in `370b3e4`) — each thread needs its own capture object. Process pool avoids GIL but adds serialization cost. Independent of fast scan mode — this is a structural improvement.

---

## 3. Artifact Output Performance

Speed of generating clips, screenshots, GIFs, and reels.

### - [x] 3A. Parallel clip cutting

**Prior art:** ThreadPoolExecutor with 4 workers already used for gallery GIF extraction (`c5b604a`), achieving 3-4x speedup. Same pattern, same module (video.py), different function.

**Status (audit):** [clipgen.py](../clipgen.py) `process_clips()` uses a thread pool when `len(prepared) >= 2` and workers ≥ 2. **Gap:** Studio [`/api/generate`](../server.py) calls `process_clips([clip], …)` once per cell inside the ndjson stream, so multi-cell batches stay **sequential** and do not use parallel cutting. See **§8.2** for follow-up.

**Experiment (original):** Use `ThreadPoolExecutor` to run N clip cuts in parallel. Default to `min(4, cpu_count)` workers, configurable via config — **landed for multi-clip `process_clips` calls**.

**Trade-offs:** Disk I/O may become the bottleneck on HDDs. On SSDs, 3-4x speedup is realistic. Source video reads are sequential (same file), but ffmpeg handles concurrent reads from the same input well since it seeks independently.

### - [x] 3B. Titlecard batching

**Current:** Each clip gets its own titlecard via a separate ffmpeg subprocess (titlecards.py:53-101). For 20 clips, that's 20 extra ffmpeg invocations.

**Experiment:** Generate all titlecards in a single ffmpeg command using filter_complex with multiple outputs. Or, generate the titlecard frame once per unique source video (since the first frame is often the same) and reuse it.

**Trade-offs:** Complex filter_complex commands are harder to debug. The "generate once, reuse" approach is simpler and covers the common case where clips from the same video share a titlecard frame.

### - [ ] 3C. Stream-copy reel detection improvements

**Current:** Reel concatenation already uses the fast concat demuxer (stream copy) when all clips share the same codec/resolution (video.py:1115-1209), falling back to filter_complex re-encoding otherwise.

**Experiment:** Log when the slow path is taken and surface a warning to the user explaining *why* (e.g. "Clip X has resolution 1280x720, others are 1920x1080"). This lets users fix the root cause rather than silently eating a 10x slowdown.

**Trade-offs:** Diagnostic-only. No risk, helps users help themselves.

### - [x] 3D. Transcript caching across sessions

**Prior art:** Skip-if-exists pattern already proven for artifact regeneration in Studio (`cbee9eb`). Same principle: check for cached output before doing expensive work.

**Current:** Transcripts are cached per-session (in-memory dict keyed by source video, clipgen.py:709). If the user runs clipgen again on the same video, the model reloads and re-transcribes.

**Experiment:** Cache the full transcript to a sidecar file (e.g. `mystudy_P01.transcript.json`) alongside the source video. On subsequent runs, load from disk instead of re-transcribing. Invalidate when the source video's mtime changes.

**Trade-offs:** Disk space is trivial (transcript JSON is a few KB). First-run cost is unchanged, but repeat runs skip the most expensive step entirely.

---

## 4. Load Times

Startup, mode-switching, and media loading.

### - [x] 4A. Lazy-import EasyOCR and scikit-image in Screenspace

**Prior art:** Deferred imports for cv2, screenspace, gspread, and openpyxl already done in `8e85ac0`, saving 3-6s on startup. Same pattern to extend.

**Current:** Heavy libraries are generally lazy-loaded well (faster-whisper, etc.). Verify that EasyOCR and `skimage` (used for SSIM in screenspace.py) are also fully lazy — they pull in torch and numpy submodules that add seconds to import time.

**Experiment:** Audit and ensure these imports happen inside the function that first needs them, not at module import time. If already lazy, no change needed.

**Trade-offs:** None. Pure improvement.

### - [ ] 4B. Defer Screenspace manifest loading

**Current:** `load_screenspace_manifest()` reads the full JSON on startup (screenspace.py:2614-2639). For large manifests (many events from long analysis sessions), this can take noticeable time.

**Experiment:** Load the manifest lazily — only when the user actually opens Screenspace. Since Studio, Insights, and Screenspace share the Flask server, a Screenspace manifest shouldn't be parsed when the user only wants Studio.

**Trade-offs:** First interaction in Screenspace may be slightly delayed. Acceptable since it only happens once per session.

### - [ ] 4C. HTTP cache headers for static assets

**Current:** CSS/JS served with `max-age=3600` (1 hour), SVGs with `max-age=86400` (24 hours), HTML with `no-cache` (server.py:1361-1373).

**Experiment:** Add content-hash or ETag-based caching for CSS/JS files. Since clipgen's web assets change only on version updates, a longer max-age with cache-busting query params (e.g. `studio.css?v=0.10.11`) would eliminate redundant fetches across page navigations within a session.

**Trade-offs:** Need to append version param to asset URLs. Small change.

### - [x] 4D. VideoCapture pool size for Screenspace

**Prior art:** LRU VideoCapture cache was expanded from 1→3 in `89bd136`. Same direction, further increment.

**Current:** LRU cache holds max 3 open `VideoCapture` objects (`_VIDEO_CAP_MAX = 3`, screenspace_server.py:67).

**Experiment:** If users commonly work with more than 3 participants, increase the pool to 5-6. Each open capture holds ~10-20 MB of state. Re-opening a capture requires re-seeking, which adds latency on frame requests.

**Trade-offs:** Memory cost is modest (50-120 MB for 6 captures). Worth it for multi-participant workflows.

---

## 5. CI Workflow Performance

Faster feedback loops for development.

### - [x] 5A. Cache uv dependencies between CI runs

**Current:** Each CI run does a fresh `uv venv && uv pip install . --torch-backend cpu` with no caching (tests.yml:34-36). The install step downloads all dependencies every time.

**Experiment:** Add `cache: true` to the `astral-sh/setup-uv@v7` step, which caches the uv global cache directory. This makes repeat installs near-instant when dependencies haven't changed.

```yaml
- name: Setup uv
  uses: astral-sh/setup-uv@v7
  with:
    python-version: "3.12"
    cache: true
```

**Trade-offs:** None. The `setup-uv` action has built-in cache support. Cache invalidation is automatic when `uv.lock` changes.

### - [x] 5B. Run lint and typecheck without installing dependencies — skipped

**Current:** The lint job uses `uvx ruff check` (no install needed, good). The typecheck job installs all dependencies including CPU-only torch (tests.yml:69-72) just to run `uvx ty check`. Total typecheck job is ~25 seconds.

**Experiment:** Check if `ty check` can run without the full dependency install.

**Result:** `ty check` needs installed packages to resolve third-party imports. Without them, every import of `gspread`, `flask`, `cv2`, `torch`, etc. produces `Unresolved import` errors that drown out real type errors. The caching from 5A makes the install fast on cache hits (~5-10s vs ~90s), so this is a non-issue.

**Trade-offs:** The job is already fast (~25s). Savings would be marginal — low priority.

### - [x] 5C. Path-scoped test runs

**Prior art:** Workflow-level path filters already skip CI entirely for non-Python changes (`b253685`). This experiment goes further — scoping *within* Python changes.

**Current:** All tests run on every PR that touches any `.py` file (tests.yml:8). A change to `screenspace.py` runs tests for CLI parsing, viewer, insights, etc.

**Experiment:** Use `pytest` markers or directory-based selection to run only relevant tests when changes are scoped. For example:
- Changes in `screenspace*.py` → run `tests/test_screenspace*`
- Changes in `viewer.py` → run `tests/test_viewer*`
- Changes in `cli.py` or `config.py` → run all tests (broad impact)

Implement via a matrix job or a script that maps changed files to test paths.

**Trade-offs:** Adds CI complexity. Risk of missing cross-cutting regressions. Mitigate by always running the full suite on pushes to master, but scoped tests on PR checks.

### - [x] 5D. Parallel test execution with pytest-xdist — skipped

**Current:** Tests run sequentially in a single pytest process (tests.yml:39).

**Experiment:** Add `pytest-xdist` and run tests with `-n auto` to parallelize across CPU cores. The test suite (19 files, mock-heavy, no shared state) is a good candidate for parallelization.

**Result:** The test suite runs in 0.62s (224 tests). pytest-xdist adds ~2-3s startup overhead per worker. On `ubuntu-latest` (2 vCPUs), the best case with xdist would be: 0.31s (tests) + 3s (startup) = ~3.3s — a 5x regression. Not worth implementing until the test suite grows significantly.

**Trade-offs:** Slight overhead for test discovery. Need to verify no tests rely on shared module-level state (e.g. config mutations). Quick win if tests are already isolated.

---

## 6. Memory and Data Efficiency

Additional performance area: how efficiently clipgen handles data in memory.

### - [ ] 6A. Paginate or stream large Screenspace event lists

**Current:** All Screenspace events are loaded into memory and sent to the frontend as a single JSON payload. A 30-minute video at 1s intervals with 3 regions generates ~5,400 events; each event has timestamps, scores, and metadata.

**Experiment:** Paginate the `/api/results/<task_id>` endpoint. Send the first 100 events immediately, then load more on scroll. Or filter server-side by time range based on the visible timeline viewport.

**Trade-offs:** Adds pagination state. Alternatively, use virtual scrolling on the frontend (complements 1E; strong pairing with Insights Builder large grids) to handle large DOM lists without reducing the data transfer.

### - [ ] 6B. Compress manifest files

**Current:** Manifest JSON is written with default formatting. For projects with hundreds of artifacts, manifests can grow to several MB.

**Experiment:** Use compact JSON separators (`(",", ":")`) when writing manifests, and optionally gzip large manifests on disk (read with `gzip.open`).

**Trade-offs:** Compact JSON is a one-liner change. Gzip adds complexity for marginal gain unless manifests exceed 10+ MB.

**Related (audit):** [viewer.py](../viewer.py) `save_manifest()` calls `load_manifest_artifacts()` and `load_manifest_reels()` separately — each reads and `json.loads` the **same** manifest file. A single read/parse that extracts both keys would cut duplicate I/O on every save (small, safe win).

---

## 7. Video Decode and I/O Efficiency

Additional performance area: the raw speed of reading video frames and writing output.

### - [ ] 7A. Hardware-accelerated video decoding

**Current:** Frame extraction uses software decoding via OpenCV's default backend (FFmpeg's libavcodec) and ffmpeg CLI.

**Experiment:** On macOS, use VideoToolbox hardware decoding (`-hwaccel videotoolbox` for ffmpeg, or `cv2.CAP_PROP_HW_ACCELERATION` for OpenCV). This offloads H.264/HEVC decoding to the GPU, freeing CPU for analysis.

**Trade-offs:** Platform-specific (macOS only for VideoToolbox; NVDEC on Linux). Needs fallback to software decode. Only helps when decode is the bottleneck (true for high-res sources at fast intervals).

### - [ ] 7B. Memory-mapped I/O for large manifests

**Current:** Manifests are read with `json.load(open(...))`, buffering the entire file in memory.

**Experiment:** For manifests >1 MB, use `mmap` or streaming JSON parsing (`ijson`) to avoid loading the full file. In practice, clipgen manifests rarely exceed a few MB, so this is low priority unless usage patterns change.

**Trade-offs:** Added dependency (`ijson`) or complexity (`mmap`). Only valuable for extreme cases.

---

## 8. Follow-up findings (codebase audit)

Additional opportunities identified in a fresh pass over the codebase (April 2026). **Transcript-related work is excluded** here by request. Line references drift over time — verify in source.

### 8.1. Reuse `SheetContext` in `generate_list` (Studio / API latency + quota)

**Issue:** [`spreadsheet.generate_list()`](../spreadsheet.py) always calls [`build_sheet_context(sheet)`](../spreadsheet.py), which performs `sheet.get_all_values()` (one full sheet fetch for Google Sheets). The combined server already holds [`_sheet_context`](../server.py) from `_init_studio_state` and refreshes it only via `POST /api/sheet/refresh`.

**Impact:** Routes such as `/api/generate`, `/api/reel`, and `/api/highlights-preview` pay an **extra full-sheet download** on every call, even when the user has not refreshed. This hurts latency and risks Google API rate limits (see AGENTS.md guidance on `get_all_values()`).

**Direction:** Optional `ctx: SheetContext | None` on `generate_list` (and callers); when provided and valid, skip `build_sheet_context`. Invalidate when the sheet is explicitly refreshed.

### 8.2. Studio multi-cell generate vs parallel clip cutting (3A gap)

**Issue:** Parallel ffmpeg in `process_clips` only activates when **multiple clips** are prepared in **one** call. Studio streams one `process_clips([clip], …)` per selected cell.

**Direction:** Batch all clips for one generate request into a single `process_clips` (then map results back to ndjson lines), or parallelize at the server layer with careful handling of `_generated_artifacts`, manifest saves, and thread safety.

### 8.3. Reel generation: sequential intermediate segments

**Issue:** [`process_reel()`](../clipgen.py) uses [`_run_clip_pipeline()`](../clipgen.py), which runs `per_clip_fn` in a simple `for` loop — no thread pool. Each segment spawns ffmpeg separately, unlike `process_clips` phase 2.

**Direction:** Parallelize segment generation after preparation (same worker discipline as 3A), or share executor logic between reel and batch clip paths.

### 8.4. Screenshot artifact pipeline vs gallery batch extract

**Issue:** Gallery uses [`_batch_extract_screenshots()`](../video.py) (one ffmpeg pass for many timestamps). The clip/screenshot path uses [`extract_screenshot()`](../video.py) per segment in [`_process_single_clip_segments()`](../clipgen.py).

**Direction:** For many screenshot outputs from the **same source file**, consider grouping timestamps into a batch pass (then applying per-clip filenames/metadata) — same spirit as 2E but scoped to artifact screenshots.

### 8.5. Highlights scoring vs output directory size

**Issue:** [`_clip_highlight_score()`](../spreadsheet.py) runs an `any(… for f in existing_filenames)` over `discover_clips()` **per clip**. Large output directories make this **O(clips × files)**.

**Direction:** One pass over `existing_filenames` to build a normalized lookup (or substring index) reused for all clips.

### 8.6. Plan document drift (corrected above)

- **1E:** Viewer + Screenspace use `DocumentFragment`; Insights Builder remains.
- **1F:** Filmstrip concurrency is 4 in viewer.js, not 2.
- **3A:** Parallel cutting is implemented for multi-clip `process_clips`; Studio single-clip loop is the remaining gap (**§8.2**).

---

## Summary: Estimated Impact Ranking

| Status | # | Experiment | Expected Impact |
|--------|---|-----------|----------------|
| [x] | 1 | 2.0 — **Fast Scan mode** (bundles 2A-2D) | High — 3-5x faster Screenspace with explicit quality trade-off |
| [x] | 2 | 2C — Universal phash pre-filter (always-on candidate) | High — cheap filter, broad applicability, minimal quality loss |
| [x] | 3 | 5A — Cache uv in CI | High — near-instant CI installs |
| [x] | 4 | 3A — Parallel clip cutting | High — 3-4x faster batch generation |
| [x] | 5 | 1A — SSE for Screenspace progress | Medium — eliminates perceived stalls |
| [x] | 6 | 1B — Preload first frames | Medium — instant first interaction |
| [x] | 7 | 3D — Transcript caching to disk | Medium — eliminates repeat transcription |
| [x] | 8 | 1E — DOM batching with fragments | Medium — smoother large lists |
| [x] | 9 | 5D — Parallel tests with xdist — skipped (suite runs in 0.62s, xdist overhead would regress) | Medium — faster CI feedback |
| [x] | 10 | 2F — Parallel region analysis | Medium — scales with CPU cores |
| [x] | 11 | 2E — Batch frame extraction via ffmpeg | Low-Medium — depends on video codec |
| [x] | 12 | 4A — Lazy-import heavy libs | Low-Medium — saves seconds on startup |
| [x] | 13 | 1C — Optimistic UI in Studio | Low-Medium — better feel during generation |
| [x] | 14 | 5B — Skip torch for typecheck — skipped (ty needs installed deps) | Low — job already ~25s, marginal gain |
| [x] | 15 | 3B — Titlecard batching | Low — small per-clip overhead |
| [ ] | 16 | 4C — Cache-busted static assets | Low — marginal for local tool |
| [ ] | 17 | 7A — Hardware decode | Low — platform-specific, complex |
| [ ] | 18 | **8.1** — Reuse `SheetContext` in `generate_list` / Studio | High — latency + Sheets quota |
| [ ] | 19 | **8.2** — Studio batch `process_clips` (unlock 3A) | High — wall time multi-cell |
| [ ] | 20 | **8.3** — Parallel reel segment ffmpeg | Medium — reel wall time |
| [ ] | 21 | **8.4** — Batch screenshot extraction (clip pipeline) | Medium — many screens / same video |
| [ ] | 22 | **8.5** — Highlights scoring index + manifest single-read | Low — scales with artifacts |
| [ ] | 23 | **1E (remainder)** — Insights Builder DOM batching / virtual scroll | Low–Medium — perceived smoothness |
