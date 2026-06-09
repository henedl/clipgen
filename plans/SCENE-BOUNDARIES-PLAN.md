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

## Phase 1: Detector

- [ ] `scan_boundaries()` in [`screenspace.py`](screenspace.py): ffmpeg-pipe full-frame iteration, phash distance, threshold + min-gap debounce, incremental `on_result`, progress, cancel
- [ ] `BoundaryTool(AnalysisTool)` registered in `TOOLS`; `supports_fast_scan = False` (the tool is already coarse; fast scan's phash-skip would fight the detector's own phash logic)
- [ ] Config: `SCREENSPACE_BOUNDARY_PHASH_THRESHOLD`, `SCREENSPACE_BOUNDARY_MIN_GAP_SECONDS`, `SCREENSPACE_BOUNDARY_INTERVAL`
- [ ] Event generation: `_extract_confidence` and `generate_events_from_results` branches for `"boundary"`; events carry `navigational: true`
- [ ] Tests: synthetic video with hard cuts → boundaries at cuts only; gradual fade → behavior documented (expected: fires once when cumulative drift crosses threshold, or not at all — assert and document whichever, don't leave it undefined); min-gap suppression; cancel mid-scan

## Phase 2: Screenspace UI

- [ ] "Boundary" workflow tab (icon + `--color-task-boundary` token); region picker hidden, params: interval / sensitivity / min gap / event label
- [ ] Timeline rendering: boundary results as thin full-height ticks in the result band, visually lighter than detector markers (they are scaffolding, not findings)
- [ ] Results list rows: timestamp + distance bar (reuse `buildConfBar`)
- [ ] Tool info text for the info tooltip

## Phase 3: Studio and Viewer integration

- [ ] Event model: `navigational` field (absent = false) flows through manifest save/load and the events API
- [ ] Studio intake: navigational events excluded from clustering and "Add all" by default; "Show navigational" toggle in the intake header includes them (a researcher *can* clip around a boundary deliberately)
- [ ] Studio "Detect boundaries" action (participant list or Metadata tab header): enqueues one boundary task per participant with defaults via `POST /screenspace/api/tasks`; progress visible through existing task polling
- [ ] Studio timeline + Timeline Viewer: boundary ticks on the Screenspace track with distinct rendering; legend entry; tooltip shows distance
- [ ] Viewer data contract: events pass through `screenspaceEvents` unchanged; `navigational` included so the viewer can style ticks differently

## Open questions (not blockers)

- **Threshold portability.** A phash threshold tuned on one game's footage may over/under-fire on another's. v1 ships a default plus the sensitivity slider; if real use shows constant retuning, the pinned-frame calibration strip applies here too (pin a known boundary pair as positive, a known continuous pair as negative).
- **Interval vs localization.** At 1 s sampling, a boundary timestamp is accurate to ±1 s. Likely fine for orientation. If precision matters later, a refine pass (binary search between the two frames straddling the spike) is cheap and local — noted as future work, not v1.
- **Letterboxed / overlaid footage.** Persistent HUD overlays dampen full-frame phash distance. If this proves problematic, hash a center crop instead of the full frame (one config constant).

## Future work (revisit)

- **Session activity curve.** The continuous companion to boundaries (CV-PLAN.md): change-ratio magnitude over time as a derived series rendered in the existing amplitude band. Separate plan; boundaries should land and be judged first.
- **Boundary-guided coarse-to-fine scanning.** Boundaries partition a session into visually stable segments; subsequent detector tasks could skip segments whose boundary frames score zero, generalizing PERFORMANCE-PLAN's coarse-to-fine idea across tools.
- **Segment labeling.** Boundaries plus the scene tool compose naturally: classify the first frame after each boundary against scene references to label segments ("menu", "gameplay"). Interpretation-adjacent — keep researcher-initiated.
