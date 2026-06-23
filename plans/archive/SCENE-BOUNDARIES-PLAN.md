# Automated Scene Boundaries Plan

## Overview

Every signal Screenspace produces today is hypothesis-conditioned: the researcher defines a region and criteria, then gets events. There is no unconditioned signal for "I don't know what I'm looking for yet." This plan adds the first one, per CV-PLAN.md's near-term roadmap: **automated scene boundary detection** — a region-less, parameter-light, full-frame pass that marks where the video's visual content changes substantially (menu → gameplay, level transitions, app/screen switches, loading screens ending).

Boundaries give researchers a skeleton of the session: entry points into long recordings without linear scrubbing. They are sparse point detections, which is exactly what the existing event model, Studio intake clustering, and Timeline Viewer track were designed for — so emitting them as ordinary `ScreenspaceEvent` records gets the entire downstream pipeline for free.

The detector itself is ~90% built. The inactivity tool's perceptual-hash machinery (`compute_phash`, hamming distance) already measures frame-to-frame visual distance; inactivity looks for plateaus, boundaries look for spikes. One new tool class in the registry, no new dependencies.

### Use cases driving this work

1. **Orientation** — "where does the tutorial end and the first level begin, across all eight sessions?"
2. **Segmentation by activity** — jump between menu time, gameplay time, and loading time without defining scene references first
3. **Calibration support** — boundary frames are maximally distinctive frames, i.e. good candidates for pinning (see PINNED-FRAMES-PLAN.md) and for scene-tool reference capture
4. **Coverage sanity** — a session whose boundary density wildly differs from its peers is a quick outlier flag in Studio's Metadata tab

## Key architectural decisions

- **Implemented as a regular `AnalysisTool` ("boundary") in the `TOOLS` registry.** This buys the task queue, progress, cancel, pause/resume, SSE updates, and event generation with zero new infrastructure ([`screenspace.py`](screenspace.py)). Region is always full frame; the run-region picker is hidden for this tool, as for template.
- **Boundaries are events, not a derived series.** They are sparse and point-like; the manifest stays ground truth and the Timeline Viewer / intake consume them unchanged. (Contrast with the session activity curve from CV-PLAN.md, which is continuous and must be a derived overlay — that is a separate, later piece.)
- **`detector: "boundary"` with a `navigational` flag.** Boundary events are for orientation, not clip candidacy. A long session can produce dozens to hundreds; flooding Studio's intake with them would degrade the curation surface. Events gain an optional `navigational: true` field; Studio intake hides navigational detectors behind a default-off toggle. Timeline tracks (Screenspace, Studio, Viewer) always show them.
- **phash distance is the v1 metric.** Hamming distance between consecutive sampled frames' perceptual hashes, computed at a downscaled resolution. Histogram or fingerprint deltas (the scene tool's machinery) are deliberately out of v1 — phash is the cheapest credible signal, and the plan's value test is whether researchers use boundaries at all, not boundary precision.
- **On-demand, not automatic.** Per CV-PLAN.md's constraint, no background pre-analysis on load. Boundary detection runs when asked — but because it is parameter-light, "asking" can be one click.
- **Studio-first surface.** Orientation is a batch problem ("which of eight sessions, and where"), and Studio already aggregates per-participant streams. Studio gains a one-click "Detect boundaries" action that enqueues a boundary task per participant through the existing Screenspace task API. Screenspace gets the tool tab in the same change at near-zero cost.

## Detection algorithm (v1)

1. Iterate full frames via the existing ffmpeg pipe at `interval` seconds (default 1.0 s), downscaled via `max_region_dim`-style scaling to ~64 px for hashing.
2. Compute phash per frame; distance `d` to the previous sampled frame.
3. A boundary fires when `d ≥ threshold` (config `SCREENSPACE_BOUNDARY_PHASH_THRESHOLD`, default ~14 — to be tuned against real footage; deliberately above the inactivity/static-skip thresholds since boundaries want large jumps only).
4. **Debounce:** after a boundary fires, suppress further boundaries for `min_gap` seconds (default 3 s). Camera-heavy gameplay produces sustained high inter-frame distance; without a gap, action sequences become boundary storms. The suppressed run's *first* spike is the boundary.
5. Confidence: `min(1.0, (d − threshold) / threshold)` floored at a small epsilon — larger jumps are more certainly boundaries.
6. Result dicts `{timestamp, distance}`; standard event generation maps these to `ScreenspaceEvent`s with `event_type` defaulting to `"boundary"` and `metadata: {distance}`.

Parameters exposed in the workflow panel: interval, sensitivity (threshold), min gap. Nothing else — the tool's entire value proposition is that it needs no setup.

## Status (updated 2026-06-18)

- **Phase 1 (Detector)** — ✅ Done. Commits `92669eb` (feat), `9e02a7f` (post-review hardening).
- **Phase 2 (Screenspace UI)** — ✅ Done. Commits `92669eb`, `12d8129` (slider widened to 0–64), `57216a6` (tab color).
- **Phase 3 (Studio + Viewer)** — ✅ Done. Studio intake hides navigational events behind a default-off "Show navigational" toggle; a one-click "Detect boundaries" action (Intake header) enqueues a full-frame boundary task per video-bearing participant; the Convergence swimlane and Timeline Viewer render boundary events as thin/lighter ticks (excluded from convergence-zone math), with a distance tooltip; Metadata gained a per-participant boundary count. `navigational` now flows through the viewer payload and `--export`.
- **Phase 4 (Scene-aware period segmentation)** — ✅ Done. Default metric `hybrid` (scene fingerprint vs. a period reference + confirm window, corroborated by a phash spike); a post-run pass merges near-identical periods, dissolves transient round-trips, and prunes session-relative weak boundaries; emits period-span metadata; adds a metric selector and two tunable knobs in the Screenspace settings tab. Detail: [SCENE-BOUNDARIES-PHASE4-PLAN.md](SCENE-BOUNDARIES-PHASE4-PLAN.md).

Post-landing fix outside the original scope: dismissing a *running* boundary task didn't cancel its worker thread (pinned CPU + SSE/icon spam) — fixed in `a7b289e` (`remove_task` keeps a running task alive until its cancel lands).

## Phase 1: Detector — ✅ Done

- [x] `scan_boundaries()` in [`screenspace.py`](screenspace.py): ffmpeg-pipe full-frame iteration, phash distance, threshold + min-gap debounce, incremental `on_result`, progress, cancel
- [x] `BoundaryTool(AnalysisTool)` registered in `TOOLS`; `supports_fast_scan = False` (the tool is already coarse; fast scan's phash-skip would fight the detector's own phash logic)
- [x] Config: `SCREENSPACE_BOUNDARY_PHASH_THRESHOLD`, `SCREENSPACE_BOUNDARY_MIN_GAP_SECONDS`, `SCREENSPACE_BOUNDARY_INTERVAL` — plus `SCREENSPACE_BOUNDARY_HASH_DIM` (pipe-level downscale) and `SCREENSPACE_BOUNDARY_CONFIDENCE_EPSILON`
- [x] Event generation: `_extract_confidence` and `generate_events_from_results` branches for `"boundary"`; events carry `navigational: true`
- [x] Tests: hard cuts → boundaries at cuts only; gradual drift below threshold → **no boundary** (documented: the detector compares consecutive samples, not cumulative drift); min-gap suppression; cancel mid-scan
- [x] Server: `"boundary"` added to `_VALID_TASK_TYPES`; the tool is genuinely region-less — `api_tasks_create` forces a `full_frame` region so events are never mislabeled
- [x] Decision: boundary is **not calibratable in v1** (no `score_key`); `/api/calibrate` rejects it. Calibration is future work (see Open questions).

## Phase 2: Screenspace UI — ✅ Done

- [x] "Boundary" workflow tab (flag icon + `--color-task-boundary` fuchsia token, dark+light); region picker **and** fast-scan toggle hidden (full-frame, coarse); params: sensitivity / min gap / interval / event label. Sensitivity slider spans **0–64** (full 8×8 phash Hamming range) after real-footage feedback that 30 was too noisy.
- [x] Timeline rendering: boundary results as thin (1px) full-height ticks, lighter (≈0.55 alpha) than detector markers — scaffolding, not findings
- [x] Results list rows: timestamp + distance bar (reuses `buildConfBar`) + `d:<distance>`; certainty-cutoff filter and task-edit param restore wired for boundary
- [x] Tool info text for the info tooltip; shared detector maps updated (`utils.js` `_DETECTOR_TYPES`/`_DETECTOR_FALLBACK`, `CATEGORY_HUES`, `TOOL_INFO`, `TOOL_LABELS`, icon-name map)
- [ ] Open follow-up: bake the validated sensitivity default into `SCREENSPACE_BOUNDARY_PHASH_THRESHOLD` + the slider's starting value once the researcher confirms the sweet spot (still 14)

## Phase 3: Studio and Viewer integration — ✅ Done

- [x] Event model: `navigational` (absent = false) round-trips through manifest save/load and the events API (already true); added to the viewer payload (`viewer.load_screenspace_events_for_viewer`) and `--export` (`data_export.build_screenspace_events` + `SCREENSPACE_EVENT_COLUMNS`)
- [x] Studio intake: navigational events excluded from clustering and "Add all" by default (`intakeClusterSource()`); a "Show navigational" toggle in the intake header includes them
- [x] Studio "Detect boundaries" action — placed in the **Intake header** (decision); enqueues one full-frame boundary task per video-bearing participant via `GET /screenspace/api/participants` → `POST /screenspace/api/tasks`; progress via existing task/event polling
- [x] Studio timeline (Convergence swimlane) + Timeline Viewer: boundary ticks render thin + lighter (~0.55 alpha) on the Screenspace track; legend entry auto-included; tooltip shows distance. **Navigational events are excluded from convergence-zone computation** (orientation scaffolding, not findings) but still rendered
- [x] Viewer data contract: `navigational` flows through `screenspaceEvents`; the viewer styles boundary ticks via `.screenspace-marker--navigational`
- [x] **Metadata: per-participant boundary count** (new, per decision) — boundaries are tallied separately and excluded from the `ss_events` findings count so coverage/outlier stats stay meaningful; surfaced in the session-summary rows and the sessions CSV (`screenspace_boundaries`)

## Phase 4: Scene-aware period segmentation — ✅ Done

> Implemented per [SCENE-BOUNDARIES-PHASE4-PLAN.md](SCENE-BOUNDARIES-PHASE4-PLAN.md). Checklist below
> reflects the shipped design (the post-run consolidation pass was added during design review).

**Motivation.** Real-footage testing showed the v1 phash-spike detector is noisy even with the sensitivity slider pushed high: it compares *consecutive samples*, so any large per-sample jump fires — camera pans, fast action, transient overlays — fragmenting what a researcher reads as one continuous "period." The goal of this phase is to make boundaries delineate **coherent periods** (menu, gameplay, loading) rather than every visual jolt, by borrowing the Scene tool's content-fingerprint sampling.

**Approach (reuse, don't reinvent).** The Scene tool already fingerprints frames robustly: `compute_scene_fingerprint()` / `compare_scene_fingerprints()` in [`screenspace.py`](screenspace.py) (color histogram, `SCREENSPACE_SCENE_HISTOGRAM_BINS`, compared against `SCREENSPACE_SCENE_SIMILARITY_THRESHOLD`). A histogram fingerprint is far less twitchy than phash under motion/HUD churn. Combine that with a trailing-window model so a boundary marks a *sustained* shift, not a one-frame spike — [`friction.py`](friction.py)'s `smooth_scores()` rolling window is the pattern to mirror.

- [x] **Period model in `scan_boundaries()`.** Each sample's fingerprint distance is measured against the *current period's* reference (seeded from the settled frame after the last boundary), not the previous frame; a boundary fires only when the distance crosses `SCREENSPACE_BOUNDARY_SCENE_THRESHOLD` and holds across `SCREENSPACE_BOUNDARY_CONFIRM_WINDOW` samples; the new boundary reseeds the reference.
- [x] **Metric selection.** `metric` param: `"phash"` (v1), `"scene"`, `"hybrid"`. Policy default `hybrid` lives in `BoundaryTool.scan` (UI "Auto" sends nothing); the scan primitive's own default stays `phash` for direct callers/tests. Sensitivity slider drives the phash threshold (hard-cut sensitivity); scene threshold is a config knob (param-light).
- [x] **Config:** `SCREENSPACE_BOUNDARY_METRIC`, `_SCENE_THRESHOLD`, `_CONFIRM_WINDOW`, `_SCENE_HASH_DIM`, plus the post-run knobs (`_MERGE_THRESHOLD`, `_SHORT_PERIOD_SECONDS`, `_RELATIVE_PRUNE_ENABLED`, `_RELATIVE_PRUNE_FACTOR`). The two researcher-facing knobs are exposed in the Screenspace settings tab; no Python↔JS duplication (server-side scan params).
- [x] **Period spans.** Each boundary result carries `period_start`/`period_end`, flowed into event metadata via `generate_events_from_results`.
- [x] **UI:** Auto/Scene/pHash/Hybrid selector in the Boundary param panel (Auto omits the param → server default); tool info text updated.
- [x] **Post-run consolidation** (added in design review): merge near-identical adjacent periods, dissolve short transient round-trips, prune session-relative weak boundaries (gated, default on).
- [x] **Tests:** consolidation helper (merge/round-trip/keep-distinct/prune/guard); scene & hybrid scan behavior (motion ignored, confirm-window blip suppression, phash corroboration); `metric="phash"` v1 regression preserved; settings listed + persisted.

> Sequencing: land and validate Phase 1+2 first (a researcher may find the phash slider sufficient). Phase 4 is the answer if retuning the slider keeps failing on busy footage — it directly subsumes the "Threshold portability" and "Letterboxed / overlaid footage" open questions below.

## Open questions (not blockers)

- **Threshold portability.** A phash threshold tuned on one game's footage may over/under-fire on another's. v1 ships a default plus the sensitivity slider (now 0–64). Early real-footage feedback already showed retuning isn't enough on busy footage → **Phase 4** (scene-aware segmentation) is the primary mitigation. Boundary is deliberately **not** calibratable in v1 (no `score_key`); if a calibration strip is wanted later it'd need `score_key` + a per-frame `check_frame` + a `needs_prev` entry + full-frame handling in `/api/calibrate` (see the `BoundaryTool` comment).
- **Interval vs localization.** At 1 s sampling, a boundary timestamp is accurate to ±1 s. Likely fine for orientation. If precision matters later, a refine pass (binary search between the two frames straddling the spike) is cheap and local — noted as future work, not v1.
- **Letterboxed / overlaid footage.** Persistent HUD overlays dampen full-frame phash distance. If this proves problematic, hash a center crop instead of the full frame (one config constant).

## Future work (revisit)

- **Session activity curve.** The continuous companion to boundaries (CV-PLAN.md): change-ratio magnitude over time as a derived series rendered in the existing amplitude band. Separate plan; boundaries should land and be judged first.
- **Boundary-guided coarse-to-fine scanning.** Boundaries partition a session into visually stable segments; subsequent detector tasks could skip segments whose boundary frames score zero, generalizing PERFORMANCE-PLAN's coarse-to-fine idea across tools.
- **Segment labeling.** Boundaries plus the scene tool compose naturally: classify the first frame after each boundary against scene references to label segments ("menu", "gameplay"). Builds directly on Phase 4's period spans. Interpretation-adjacent — keep researcher-initiated.
