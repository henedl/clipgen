# Performance — clipgen

## Context

Three independent audits flagged overlapping throughput and responsiveness issues in clipgen: an external first pass (Studio generate path), a verification review of that pass, and a deeper follow-up across Screenspace, Transcripts, and CLI/pipeline. The repo already follows good performance patterns (parallel ffmpeg, mtime-cached manifests, lazy thumbnails, NDJSON streaming) — what remains is a small set of *inconsistent* parallel paths, a few hot-loop allocations and missed caches, one cache LRU bug, and a couple of full DOM rebuilds in Studio.

Code-level claims were spot-verified against the current source. One first-pass claim (a cross-subsystem cache double-extraction) and one deep-pass claim (`pipeline.py:238` probe-per-segment) did not survive verification and are dropped. The remainder is accurate and actionable.

This plan organizes the work into five phases plus an appendix so individual phases can ship independently.

Source review: `~/.claude/plans/system-instruction-you-are-working-shimmying-puddle.md`.

---

## ✅ Phase 1 — Quick wins (single PR)

Seven small, independent diffs. None changes a public contract; all are safe to bundle.

### ✅ 1.1 Studio uses `_resolve_clip_workers()`

- **File:** `server.py:757`
- **Today:** `workers = min(4, os.cpu_count() or 1)` (hardcoded).
- **Change:** `workers = pipeline._resolve_clip_workers()`. Move `_resolve_clip_workers` to `utils.py` if `server.py` should not import from `pipeline.py` (check existing import direction first).
- **Why:** `config.CLIP_PARALLEL_WORKERS` already exists and is honored by CLI; Studio silently ignores it. Aligns the two paths and unlocks tuning.

### ✅ 1.2 LRU refresh on Studio thumbnail cache hit

- **File:** `server.py:243-249`
- **Today:** `cached = _thumbnail_cache.get(cache_key)` — no `move_to_end()` on hit.
- **Change:** After the `get`, before returning, `_thumbnail_cache.move_to_end(cache_key)`.
- **Why:** Mirrors `screenspace_server.py:328-331`. Without it, hot thumbnails are evicted before cold ones under load, causing repeat ffmpeg extractions that *look* like server slowness.

### ✅ 1.3 Targeted cell-class updates after queue mutations

- **File:** `assets/web/studio.js`
- **Today:** `updateCellClasses()` (lines 1483-1496) is called from 12+ sites including baselines, sidebar filters, and batch ops. It iterates *every* `.ts-cell` and runs `findInQueue` per cell.
- **Change:** Audit the 12 call sites. Where the mutation is known to affect a small set of cells (toggle, single add/remove), call `updateSingleCellClass` per affected cell instead. Where a sweep *is* required (e.g. participant column visibility), keep `updateCellClasses` but cache `findInQueue` lookups by `cellKey` for the duration of the call.
- **Why:** O(rows × participants × queueLen) per sidebar toggle is the largest source of "Studio feels janky" complaints on large sheets.

### ✅ 1.4 Pre-compile correction patterns in `apply_corrections`

- **File:** `transcripts.py:346-364`
- **Today:** `re.compile(re.escape(from_text), re.IGNORECASE)` runs **inside** the per-segment loop. With 1000 segments × 20 corrections that is 20k recompiles per transcript.
- **Change:** Build `pairs = [(re.compile(re.escape(c["from"]), re.IGNORECASE), c["to"]) for c in corrections ...]` once at function entry; iterate that list inside the segment loop.
- **Verified:** TRUE — re-reading the source confirms the recompile is in the inner loop.

### ✅ 1.5 Cache `get_known_annotation_map()`

- **File:** `utils.py:633-639`
- **Today:** Function rebuilds the dict from `config.ANNOTATION_KEYPHRASES` on every call. No decorator, even though `@functools.cache` is already used in the same file (line 572).
- **Change:** Add `@functools.cache` above the `def`. The map is keyed by configured tokens; it never changes during a process lifetime.
- **Verified:** TRUE — `Grep` confirmed no cache decorator on this function.

### ✅ 1.6 Faster `timestamp_to_seconds` dispatch

- **File:** `utils.py:1050-1065`
- **Today:** Tries `datetime.strptime(ts, "%M:%S")`, catches `ValueError`, then tries `"%H:%M:%S"`. Every HH:MM:SS timestamp pays the exception cost.
- **Change:** Pre-dispatch by counting colons in `ts` before calling `strptime`:
  ```python
  fmt = "%H:%M:%S" if ts.count(":") == 2 else "%M:%S"
  try:
      parsed = datetime.strptime(ts, fmt)
  except ValueError:
      return None
  ```
  Same semantics, no exception-driven control flow on the hot path.
- **Verified:** TRUE — re-reading confirmed the two-format try/except chain.

### ✅ 1.7 Tuning documentation

- **File:** `AGENTS.md` (or a new section in `agents/PERFORMANCE.md`).
- **Add:** A short table listing the existing knobs and when to turn them: `CLIP_PARALLEL_WORKERS`, `SCREENSPACE_PARALLEL_WORKERS`, `SCREENSPACE_CV_RESOLUTION_SCALE`, `SCREENSPACE_BATCH_EXTRACT`, `GALLERY_BUNDLE_ENABLED`. Note that `GALLERY_BUNDLE_ENABLED` should not be used on large galleries.

**Verification (Phase 1):**
- `uv run --extra dev pytest -c tests/pytest.ini` — no regressions. Pay particular attention to `tests/test_transcripts*.py` and any `parse_timestamps` / annotation-map tests.
- Manual: open Studio on a real sheet, toggle a participant column off/on, confirm the grid still updates correctly. Open DevTools Performance tab, record a toggle, confirm scripting time drops vs baseline.

---

## Phase 2 — Backend throughput

Three medium-sized improvements that each unlock a parallel path.

### ✅ 2.1 Parallelize `_generate_intake_clips`

- **File:** `server.py:475-505`
- **Today:** Bare `for item in items:` calling `video.run_ffmpeg()`. No executor.
- **Change:** Mirror the gallery pattern — collect work, dispatch to `ThreadPoolExecutor(max_workers=pipeline._resolve_clip_workers())`, return ordered results via index. Each task returns the existing `_ok` / `_error` shape.
- **Constraint:** Output filename collisions — verify each item already gets a unique output path before parallelizing (or serialize the naming step).

### ✅ 2.2 Parallel reel regeneration

- **File:** `pipeline.py:1173-1182` (and the follow-up loops at 1202-1204, 1209-1211).
- **Today:** `for reel in reel_list: _regenerate_reel(reel)` — sequential despite the surrounding code parallelizing media artifacts.
- **Change:** Wrap the reel loop in a `ThreadPoolExecutor` using the same worker count. Reels are independent encode jobs.
- **Constraint:** Confirm `_regenerate_reel` is reentrant (no shared mutable state, no progress-bar contention). The existing media-artifact executor at 1138 is a working precedent.

### ✅ 2.3 mtime cache for `load_screenspace_events_for_viewer`

- **Files:** `viewer.py:137-156`, `screenspace.py:load_screenspace_manifest`, `utils.py:515-526`.
- **Today:** Every call re-reads and re-parses the full `screenspace_manifest.json`.
- **Change:** Add an mtime cache to `load_screenspace_events_for_viewer` mirroring `viewer._load_manifest_both` (`viewer.py:381-421`). Key by `(path_str, st_mtime_ns)`. Bound at 1 entry (current path) — this is a single-process viewer export, not multi-tenant.
- **Why:** Low absolute cost per call, but the function is hit on every viewer export, every regenerate, and any future feature that loops over participants for export. Cheap insurance.

**Verification (Phase 2):**
- Add or extend a smoke test in `tests/` for parallel intake (assert ordering + per-item error propagation).
- Manual: trigger Studio intake on a batch of ≥4 items, confirm wall-clock drops roughly proportional to `CLIP_PARALLEL_WORKERS`.
- Manual: `--regenerate` on a manifest with multiple reels; confirm reel encode times overlap in logs.

---

## Phase 3 — `_find_existing_artifacts` skip-pass

- **File:** `server.py:653-664`.
- **Today:** Linear scan over `_generated_artifacts` + per-candidate `Path(...).is_file()`. Called once per cell during Phase 1 of `/api/generate`.
- **Cost:** O(N_cells × N_artifacts) plus N_cells × disk syscalls.
- **Change (in order of effort):**
  1. **First:** mtime-cache the existence check. Wrap `is_file()` in a tiny LRU keyed by absolute path string + a session-scoped invalidation hook (clear on new generate run). The syscall is the dominant cost; the linear scan is secondary.
  2. **Then if still slow:** Build a `defaultdict[(cellRow, cellCol, type), list[artifact]]` index on `_generated_artifacts`. Rebuild on append; invalidate on session reset.
- **Why first the cache, then the index:** The review noted that the agent's "index it" recommendation likely overstates the linear-scan cost relative to syscalls. Doing both is fine; do them in the order above so each step's win is measurable.

**Verification (Phase 3):**
- Manual: run a generate over a session with 1000+ existing artifacts and ≥100 cells, log Phase 1 duration before/after each step.

---

## Phase 4 — Frontend perceived perf

Two larger frontend changes. Each is its own session.

### 4.1 Studio grid: incremental filter, then virtualization (gated on profiling)

- **File:** `assets/web/studio.js`, `renderGrid()` around line 1024.
- **Today:** `grid.innerHTML = ""` + full rebuild on filter, baseline, participant visibility, sheet load.
- **Approach:**
  1. **First:** Profile a real large sheet (200+ rows × 12+ participants) with the Performance tab. Record `renderGrid` duration on filter toggles. **Do not start work until this baseline exists** — virtualization is a real cost to maintain and we should know the savings.
  2. **If filter toggles dominate:** Convert participant-column visibility and sidebar filter changes to incremental updates — toggle `.hidden` on existing nodes instead of rebuilding.
  3. **If sheet load / baseline arrival dominates:** Render in chunks via `requestAnimationFrame`, yielding to paint between row batches.
  4. **Only if both still feel slow:** Virtualize rows (windowed render + spacer divs). This is the largest invasive change in the plan — keep it last.

### ✅ 4.2 NDJSON streaming for `/api/generate-intake`

- **Files:** `server.py:1474-1497` (endpoint), `assets/web/studio.js:2946` (client).
- **Today:** Blocking POST, single `jsonify()` response after the full batch.
- **Change:** Convert to NDJSON streaming mirroring `/api/generate` (server.py ~730-800). Per-item yields, client appends to UI as each item completes. Pair with Phase 2.1 (parallel intake) so the stream feels live.
- **Why:** Today a 10-item intake batch shows a spinner for ~30s with no feedback; after this, items pop in as they finish.

**Verification (Phase 4):**
- Manual: Studio on a large sheet — toggle filters, confirm subjective improvement and no flicker.
- Manual: Intake on ≥10 items — confirm per-item feedback appears as each finishes.
- Confirm no regressions in existing studio smoke tests.

---

## Phase 5 — Verify-then-fix (deeper audit follow-ups)

Items surfaced by the deep audit that look real but were not exhaustively re-read by the reviewer. Each starts with a short verification step before the diff. None is a single-line fix; each is its own small PR.

### 5.1 Debounce Screenspace manifest writes

- **File:** `screenspace_server.py` (calls to `_persist_manifest(drain_events=False)` around 1419, 1433, 1437, 1452, 1463, 1474 — verify exact lines before editing).
- **Today:** Every task enqueue / cancel / reorder / pause / resume calls `_persist_manifest()`, which serializes and disk-writes the full JSON.
- **Verify first:** Re-read the call sites and confirm they are all on hot mutation paths (not e.g. user-explicit "save" actions). Confirm that delaying a write by 2–3 seconds is acceptable across power-loss / kill scenarios.
- **Change:** Introduce a `_manifest_dirty` flag and a `threading.Timer` (or a worker-thread tick) that flushes after a short debounce. Force a flush on shutdown and on explicit user save actions.

### ✅ 5.2 Hoist morphology kernel allocations in Screenspace

- **File:** `screenspace.py` (~lines 197, 996 per the audit — verify).
- **Today:** `np.ones((mk, mk), np.uint8)` allocated inside per-frame loops in `extract_inactivity()` and `scan_changes()`.
- **Verify first:** Confirm kernel size (`mk`) is a config constant, not derived from per-frame data. If it is config-static, the kernel is safe to share.
- **Change:** Hoist to a module-level lazy-initialized `_MORPH_KERNEL` keyed by size if multiple sizes are used; otherwise a single module constant.

### ✅ 5.3 Re-read assets from disk on every viewer export

- **File:** `viewer.py` (HTML / CSS / JS template reads in `finalize_*` / export helpers around 176-217 per audit — verify exact functions).
- **Today:** Each export call re-reads the bundled CSS and JS template files from disk via `Path.read_text()`.
- **Verify first:** Confirm the file paths are bundled assets (via `utils.get_bundled_assets_root()` per `agents/CODE-REVIEW.md`), not user-supplied templates that could change between exports.
- **Change:** Module-level lazy cache (similar to the existing `_MANIFEST_CACHE` pattern but keyed by absolute path; no mtime invalidation needed for bundled assets).

### 5.4 `requests.Session()` reuse in Ollama client

- **File:** `ollama_client.py` (the deep audit pointed at lines 238-250; verify whether the project already uses `requests` or only `urllib`).
- **Today:** Each `generate()` call opens a fresh HTTP connection — no keep-alive, no connection pool.
- **Verify first:** Confirm the current transport (`urllib.request` vs `requests`). If `urllib`, decide whether to introduce `requests` as a dependency or switch to `urllib3.PoolManager`.
- **Change:** Module-level `_session` (or `PoolManager`) reused by all `generate()` / `list_models()` / `is_available()` calls. Set a sane connect/read timeout.

### ✅ 5.5 mtime cache for `/api/participants` artifact lookup

- **File:** `transcripts_server.py:~199` (verify the exact function name and that `viewer.load_manifest_artifacts()` is indeed called per request).
- **Today:** Each poll of `/api/participants` re-reads and re-parses the artifact manifest, then does a per-participant linear scan over all artifacts.
- **Verify first:** Confirm the frontend's poll interval and whether the endpoint returns enough state that a per-participant ETag would be useful.
- **Change:** mtime-cache the artifact list mirroring `viewer._load_manifest_both`. If the linear scan is also hot, index by participant at cache-build time.

### ✅ 5.6 Binary-search citation parsing

- **File:** `thinking_agents.py:~224` (`_find_closest_segment`).
- **Today:** O(claims × segments) linear scans during citation parsing.
- **Verify first:** Confirm `seg_starts` is sorted by start time (Whisper output is sorted, but the parser may receive a transformed list). Add an assertion if needed.
- **Change:** Replace the per-timestamp linear scan with `bisect.bisect_left` on `seg_starts`, checking adjacent indices for the closer match.

**Verification (Phase 5, common):** Each sub-item ships with its own targeted check. No shared verification step.

---

## Out of scope (explicitly dropped or deferred)

- **Batch `process_clips` in Studio /api/generate (the original P0).** Real lever, but it competes with the existing NDJSON-per-cell streaming contract and requires a callback-based progress shape from inside `process_clips`. Tracked separately — needs its own design doc covering fuzzy-match cache sharing and per-cell progress emission. Not in this plan.
- **Shared frame cache Studio ↔ Screenspace (original P2).** The premise (intake thumbs hitting the Screenspace endpoint) is not present in current code. Revisit only if a future feature actually wires Studio thumbs through Screenspace.
- **SSE / WebSocket for task progress.** Polling is fine; effort vs benefit is too low for now.
- **`pipeline.py:238` probe-per-segment claim.** Verified false — the probe is hoisted outside the time loop with an explicit comment saying so. Dropped.

---

## Appendix — Correctness bug to investigate separately

Not a perf item, but surfaced by the audit and worth a separate look:

- **`video.py:876-880`** — In `get_file_duration`, after a successful `probe_video_properties(filepath)` call, the function returns `_file_duration_cache.get(resolved)`. But `probe_video_properties` populates `_video_properties_cache`, not `_file_duration_cache`. If the cache hasn't been seeded elsewhere, this branch returns `None` despite a valid duration being available in `probed["duration"]`. Trace whether `probe_video_properties` has a side effect that populates `_file_duration_cache`; if not, this is a real bug (silently returns `None` and falls through to the `_probe_duration_seconds_ffprobe_format` fallback, paying for an extra ffprobe). File a separate bug-fix PR rather than bundling into the perf work.

---

## Critical files

- `server.py` — Phases 1.1, 1.2, 2.1, 3, 4.2
- `assets/web/studio.js` — Phases 1.3, 4.1, 4.2
- `transcripts.py` — Phase 1.4
- `utils.py` — Phases 1.5, 1.6, 2.3
- `pipeline.py` — Phase 2.2
- `viewer.py` / `screenspace.py` — Phases 2.3, 5.2, 5.3
- `screenspace_server.py` — Phase 5.1
- `ollama_client.py` — Phase 5.4
- `transcripts_server.py` — Phase 5.5
- `thinking_agents.py` — Phase 5.6
- `AGENTS.md` or `agents/PERFORMANCE.md` — Phase 1.7

## Reuse

- `pipeline._resolve_clip_workers()` — `pipeline.py:51-56`.
- `viewer._load_manifest_both()` mtime-cache pattern — `viewer.py:381-421`.
- `screenspace_server._frame_cache.move_to_end()` LRU pattern — `screenspace_server.py:328-331`.
- Existing NDJSON-streaming generate route — `server.py` ~730-800, as the template for 4.2.
- `@functools.cache` precedent — `utils.py:572`, applied in 1.5.

## Suggested commit ordering

1. Phase 1 (single PR, seven diffs). Land first — independent quick wins.
2. Phase 2.3 (mtime cache) — small, no-risk; can ship with Phase 1 or alone.
3. Phase 5.2 / 5.3 (kernel hoist + viewer template cache) — tiny, no-risk; can also ride with Phase 1 if scope allows.
4. Phase 2.1 (parallel intake) — separate PR; touches an endpoint.
5. Phase 2.2 (parallel reel regen) — separate PR; touches the regenerate path.
6. Phase 5.1 (manifest debounce) — separate PR; touches mutation hot paths.
7. Phase 5.4 (Ollama session reuse) — separate PR; small but introduces a dependency consideration.
8. Phase 5.5 / 5.6 (transcripts caching + bisect) — bundle into a single "transcripts perf" PR.
9. Phase 3 — depends on at least one of the above being deployed so we can measure real Phase-1 durations.
10. Phase 4 — last, gated on real profiling.
11. Appendix correctness bug (`video.py:876-880`) — own PR, independent of perf work.
