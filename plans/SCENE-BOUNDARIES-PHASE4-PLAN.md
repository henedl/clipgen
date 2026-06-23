# Scene Boundaries — Phase 4: Scene-aware period segmentation (detailed plan)

> Detail for the Phase 4 stub in [SCENE-BOUNDARIES-PLAN.md](SCENE-BOUNDARIES-PLAN.md).
> **Status: ✅ implemented.** Default metric `hybrid`; period-reference model + confirm window;
> post-run merge/round-trip/relative-prune pass; settled-frame retention; period-span metadata;
> metric selector UI; two tunable knobs in the Screenspace settings tab.

## Context / problem

The v1 boundary detector (`scan_boundaries`) compares **consecutive sampled frames'**
perceptual hashes and fires on any Hamming spike ≥ threshold. Real-footage testing showed this is
noisy even with the sensitivity slider maxed: camera pans, fast action, and transient overlays each
produce a large per-sample jump, fragmenting one continuous "period" (menu / gameplay / loading)
into many ticks. Phase 4 makes boundaries mark **coherent period transitions**, not every visual
jolt, by (a) measuring against a *period reference* fingerprint rather than the previous frame, (b)
requiring a shift to **hold** across a short confirmation window, and (c) using the Scene tool's
content fingerprint (robust to motion) instead of raw phash.

This is a backend-weighted change. The detector already emits ordinary navigational
`ScreenspaceEvent`s, so everything Phase 3 built (intake hiding, thin ticks, metadata count)
continues to work unchanged — Phase 4 only changes *when/where* boundaries fire and adds optional
period-span metadata.

## Building blocks to reuse (verified in code)

- **`scan_boundaries()`** — `screenspace_scans.py:1082`. Current phash loop with `on_result`
  streaming, progress, cancel, ffmpeg-pipe full-frame iteration via `scan_video_full_frames`. This
  is the function we extend.
- **`compute_scene_fingerprint()` / `compare_scene_fingerprints()`** — `screenspace_primitives.py:686/735`.
  HSV 3D histogram + edge density + color stats → similarity 0.0–1.0 (higher = more similar). The
  Scene tool fires when similarity ≥ `SCREENSPACE_SCENE_SIMILARITY_THRESHOLD` (0.75); we use
  **distance = 1 − similarity** and fire when distance ≥ a scene threshold.
- **`compute_phash()`** — `screenspace_primitives.py:406`. Works on any input size (internally
  resized), so a 128 px pipe frame feeds both phash and the fingerprint.
- **`friction.smooth_scores()`** — `friction.py:200`. Rolling-window pattern referenced by the plan;
  we mirror its "shift must persist" idea via an explicit confirmation counter (a centered mean
  isn't needed for a causal streaming scan).
- **`BoundaryTool.scan()`** — `screenspace_tools.py:781`. Passes `params.get(...)` straight to
  `scan_boundaries`; adding a `metric` arg is a one-line passthrough.
- **Param flow** — JS `gatherWorkflowParams` (`screenspace.js:5302`) sets `threshold/min_gap/interval`;
  `_validate_task_request` passes `parameters` through untouched; the server already forces a
  full-frame region for boundary. So a new `metric` param needs only: JS gather + restore, and
  `scan_boundaries` reading it. No new validation branch required.

## Algorithm (v2)

`scan_boundaries` gains a `metric` argument: `"phash"` (current v1), `"scene"`, or `"hybrid"`.

### Period-reference scene model (used by `scene` and `hybrid`)

```
ref_fp      = fingerprint of the first sampled frame   # current period reference
pending     = 0                                         # consecutive samples above threshold
pending_start = None                                    # (ts, fp, phash_corroborated) of run start
for each sample (ts, pixels):
    fp   = compute_scene_fingerprint(pixels)
    dist = 1 - compare_scene_fingerprints(fp, ref_fp)
    if dist >= scene_threshold:
        if pending == 0: pending_start = (ts, fp, <phash spike seen at this sample?>)
        pending += 1
        if pending >= confirm_window and not within min_gap:
            fire boundary at pending_start.ts          # transition START, not the confirm frame
            ref_fp = pending_start.fp                   # seed next period's reference
            last_boundary_ts = pending_start.ts
            pending = 0
    else:
        pending = 0                                     # blip didn't hold → discard, ref unchanged
```

- **Why measure against `ref_fp`, not the previous frame:** during a sustained new scene every frame
  stays far from the *old* reference (so `pending` accrues and fires once), then the reference jumps
  to the new scene and subsequent frames read as "same period" — no re-fire. Intra-period motion
  stays *similar to the period reference*, so it never accrues `pending`.
- **Confirmation window** suppresses a single-sample blip (a one-frame overlay/explosion): the shift
  must persist `confirm_window` samples to count.
- **Boundary timestamp** is the run's first exceeding sample (the real transition start); it's just
  *emitted* `confirm_window` samples later in wall-clock — the timestamp itself is correct.
- **`min_gap`** is kept as a secondary debounce.

### `hybrid` = scene-confirmed AND phash-corroborated

Run the scene model, but additionally track the consecutive-frame phash distance each sample (we
already have `prev_hash`). Emit a confirmed boundary **only if** a phash spike
(`phash_dist ≥ phash_threshold`) occurred within the pending run. This catches hard cuts (phash
spike that *also* sustains a fingerprint shift) and rejects both motion (phash spikes that don't
sustain) and slow fades (fingerprint drift with no phash spike).

### Per-period representative for the post-run pass

On each fire, seed the next period's reference from the **settled** frame (the current sample at
confirm time, which has held for `confirm_window` samples) rather than the transition-start frame —
this is both a better streaming reference and the representative the post-run pass needs. Retain,
per period, the **downscaled settled frame pixels** (≈64×64 BGR, ~12 KB) — *not* the full HSV
histogram (~1 MB each). Even hundreds of over-fired periods cost only a few MB, freed after the
scan, and the post-run pass recomputes fingerprints on demand via the same
`compute_scene_fingerprint` / `compare_scene_fingerprints`.

### Confidence

`scene`/`hybrid`: `conf = clamp((dist − scene_threshold) / scene_threshold, eps, 1.0)` — same shape
as the phash path, reusing `SCREENSPACE_BOUNDARY_CONFIDENCE_EPSILON`.

### Frame size

For `scene`/`hybrid`, set the pipe downscale to `SCREENSPACE_BOUNDARY_SCENE_HASH_DIM` (128) so the
HSV histogram is meaningful (the existing 64 px phash downscale is too coarse for a 64³-bin
histogram). `phash` metric keeps the cheap 64 px path. `compute_phash` is size-agnostic, so hybrid's
phash component is unaffected at 128 px.

## Post-run sanity pass (`scene` / `hybrid` only)

A forward-only scan can't know that "the scene 4 s ahead is identical to the one before this
boundary." After the scan loop, a global pass over the finished period list cleans up what the
causal scan can't. It runs only for `scene`/`hybrid` (which retain settled frames); `phash` keeps v1
behavior. **Principle: merge periods that aren't really different, and dissolve transient
interruptions.**

### Step 1 — fixed-point merge loop (repeat until no change; tens of periods, cheap)

- **Adjacent-merge (rule B):** for a boundary separating periods *Pᵢ* / *Pᵢ₊₁*, if
  `scene_distance(rep[i], rep[i+1]) < merge_threshold` the two sides are the same scene → drop the
  boundary, merge. (Catches a single spurious boundary inside one continuous scene — a long pan the
  confirm window couldn't reject.)
- **Round-trip / transient dissolve (rule A):** for a *short* period *Pᵢ*
  (duration `< short_period_seconds`), if `scene_distance(rep[i-1], rep[i+1]) < merge_threshold` the
  surrounding periods are the same scene → *Pᵢ* was a transient (popup/flash) → drop **both**
  bracketing boundaries, dissolving *Pᵢ*.
- **Safety:** a short-but-*distinct* period (a 2 s loading screen differing from both neighbors)
  matches neither rule and is **kept** — this is why we don't use a blanket minimum-duration filter.
- On any drop: extend the predecessor's `period_end` to the next survivor and keep the
  **higher-confidence** entry distance.

### Step 2 — relative-confidence prune (gated, default ON)

When `SCREENSPACE_BOUNDARY_RELATIVE_PRUNE_ENABLED` and ≥ a small minimum count of boundaries survive
(guard so the median is stable, e.g. ≥ 4): compute the **median** of surviving boundaries' entry
distances and drop any boundary whose entry distance `< relative_prune_factor × median`. This is a
*session-relative* cutoff — the strongest mitigation for the "threshold portability" open question,
since it adapts to each game's typical inter-scene distance instead of a fixed absolute number. It
runs after merging so it ranks only real period transitions.

> The prune is intentionally the *last* and *toggleable* step: merging is the safe, high-value core;
> the prune trades a little recall for cross-footage robustness, so it's a default-on switch the
> researcher can flip off in settings.

### Streaming contract (correctness)

The worker generates authoritative events from the **`on_result` stream** (`_raw_results`), not the
returned list (`screenspace_worker.py:535`). Since the post-run pass *revises* the boundary set,
`scene`/`hybrid` must **not** stream provisional boundaries during the scan; instead accumulate
internally, consolidate, then emit `on_result` for each *final* boundary at scan end (and return the
same list). `phash` keeps streaming live as today (no consolidation). Net effect for scene/hybrid:
the progress bar advances live; ticks appear together at completion — fine for a coarse scan.

## Period spans (metadata, emit-only)

After the scan loop, attach to each result `metadata`:
`period_start = this boundary ts`, `period_end = next boundary ts` (last → `end_seconds`). These
flow through existing `generate_events_from_results` (which already copies the result's `distance`
into `metadata`) into `event["metadata"]`. **Rendering** spans as segments in Studio/Viewer is *not*
in Phase 4 — only the data is emitted, as the bridge to the future "Segment labeling" item.

## Config (config.py)

```python
SCREENSPACE_BOUNDARY_METRIC: str = "hybrid"               # default algorithm
SCREENSPACE_BOUNDARY_SCENE_THRESHOLD: float = 0.25        # fingerprint distance (1 − similarity); mirrors scene's 0.75 sim
SCREENSPACE_BOUNDARY_CONFIRM_WINDOW: int = 2              # samples a shift must persist
SCREENSPACE_BOUNDARY_SCENE_HASH_DIM: int = 128           # pipe downscale for fingerprinting
SCREENSPACE_BOUNDARY_MERGE_THRESHOLD: float = 0.15        # post-run: merge periods this similar (exposed in settings)
SCREENSPACE_BOUNDARY_SHORT_PERIOD_SECONDS: float = 3.0    # post-run: transient-dissolve candidate ceiling
SCREENSPACE_BOUNDARY_RELATIVE_PRUNE_ENABLED: bool = True  # post-run: session-relative weak-boundary prune (exposed in settings)
SCREENSPACE_BOUNDARY_RELATIVE_PRUNE_FACTOR: float = 0.5   # drop boundaries below factor × median entry distance
```

### Exposed in the settings modal (per decisions 4 & 5)

Register the two user-tunable knobs in `config.STUDIO_SETTINGS` (tab **"Screenspace"**) +
`config.SETTINGS_DESCRIPTIONS`. They auto-apply to `config` globals via the existing
`setattr(config, name, value)` path (server.py settings load/save), so `scan_boundaries` reading
`config.SCREENSPACE_BOUNDARY_*` picks up tuned values with **no per-task param and no JS** — the
settings modal is data-driven from `/api/settings`:

```python
"SCREENSPACE_BOUNDARY_MERGE_THRESHOLD":       {"tab": "Screenspace", "group": "Boundaries", "type": "float", "min": 0.0, "max": 1.0, "step": 0.01},
"SCREENSPACE_BOUNDARY_RELATIVE_PRUNE_ENABLED":{"tab": "Screenspace", "group": "Boundaries", "type": "bool"},
```

The remaining boundary knobs (scene threshold, confirm window, short-period ceiling, prune factor)
stay config-only to keep the surface small. No new Python↔JS duplicated constants: the metric
selector's **"Auto"** default sends *no* `metric` (server applies `SCREENSPACE_BOUNDARY_METRIC`), and
every post-run knob is a server-side scan parameter — nothing needs to flow through
`get_frontend_config()`.

## UI (screenspace.js / screenspace.html / screenspace.css)

- Add a small **Metric** selector to the Boundary param panel (`renderWorkflowParams`, near
  `screenspace.js:4296`): `Auto` (default) · `Scene` · `Phash` · `Hybrid`. `Auto` omits the param;
  others send `metric`. Wire into `gatherWorkflowParams` (~5302) and the task-edit restore (~5964).
- The existing **Sensitivity** slider keeps driving the **phash** threshold (hard-cut sensitivity,
  used by `phash` and `hybrid`). Scene threshold + confirm window stay config-only to keep the tool
  parameter-light ("one click" still works because `Auto` + defaults need no setup). *(See Q2: if
  you'd rather the slider drive the scene-distance threshold in scene mode, that's a small remap.)*
- Update the boundary tool info text (`screenspace.js:3654`) to describe period-based detection.

## Tests (tests/screenspace/)

Mirror the existing boundary scan tests (synthetic frame sequences via the `_ss_helpers` patterns):

- High intra-period motion but a stable scene fingerprint → **no** spurious boundaries (the
  regression the slider can't fix).
- A genuine scene change → **exactly one** boundary, timestamped at the transition start.
- Confirm-window suppresses a single-sample blip.
- `metric="phash"` reproduces current v1 behavior byte-for-byte (regression guard).
- `hybrid` rejects a fingerprint shift with no phash spike, and a phash spike with no sustained
  shift; fires when both coincide.
- Period-span metadata: `period_end` of boundary *i* equals the timestamp of boundary *i+1*; last
  equals `end_seconds`.
- **Post-run merge (B):** a spurious boundary between two near-identical periods is removed.
- **Post-run round-trip (A):** a short transient period bracketed by identical scenes is dissolved
  (both boundaries dropped); a short *distinct* period (loading screen) is **kept**.
- **Relative prune:** with the toggle on, a weak boundary far below the session median is dropped;
  with it off, it survives; the min-count guard prevents pruning when too few boundaries exist.
- The post-run pass is best tested directly as a pure helper over a synthetic period list (settled
  frames + entry distances), independent of the video scan, plus one end-to-end scan assertion.
- Server smoke: a boundary task with `parameters.metric="scene"` round-trips and runs (extend
  `tests/test_screenspace_api.py`); `tests/test_studio_api.py` covers the two new settings being
  listed/saved (mirrors existing `SCREENSPACE_*` settings tests).

## Files touched

- `config.py` — 8 new constants; register 2 in `STUDIO_SETTINGS` + `SETTINGS_DESCRIPTIONS` (Screenspace tab).
- `screenspace_scans.py` — `scan_boundaries` period model + `metric` arg + settled-frame retention +
  period spans + the post-run sanity pass (factor the pass into a pure, separately-testable helper).
- `screenspace_tools.py` — `BoundaryTool.scan` passes `metric`.
- `screenspace_primitives.py` — reuse only (no change expected).
- `assets/web/screenspace.{js,html,css}` — metric selector + info text. (Settings need no JS — the
  modal is data-driven from `/api/settings`.)
- `tests/screenspace/…`, `tests/test_screenspace_api.py`, `tests/test_studio_api.py` — new coverage.
- `build/VERSION` — patch bump (this is a `feat:`).

## Verification

1. `/check` green (ruff + ty + pytest), `node --check` on touched JS.
2. Local headless: `uv run clipgen.py --screenspace -i IN -o OUT`, run a boundary task with each
   metric on busy footage; confirm `hybrid`/`scene` produce far fewer, period-aligned boundaries
   than `phash`, and `period_start/period_end` appear in `screenspace_manifest.json`.
3. Browser (you): the Metric selector appears, defaults to Auto, and the thin ticks land at real
   period transitions; "Detect boundaries" from Studio (Phase 3) now yields a clean skeleton.

## Resolved decisions

1. **Default metric = `hybrid`** — scene-confirmed AND phash-corroborated. This is the core intent.
2. **Sensitivity slider stays on the phash threshold only**; scene threshold + confirm window are
   config-only knobs (tool stays parameter-light).
3. **Expose** the Auto/Scene/Phash/Hybrid metric selector in the Boundary param panel.
