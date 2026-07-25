# Attention — computational saliency Screenspace tool

Rolling plan for the Attention tool: predicted visual-attention heatmaps from screen
recordings with no eye-tracking hardware, plus timeline events at attention-shift
moments. Update the status markers as work lands.

## Background

Revives the "Attention-Guided Scrubbing" idea from `plans/archive/CV-PLAN.md`.

Design decisions (settled 2026-07):

- **Engine: classic CV composite, zero new dependencies.** `cv2.saliency` is
  unavailable (opencv-python-headless, non-contrib — see `plans/LICENSE-PLAN.md`);
  pysaliency is a benchmark harness and deepgaze needs torch weights. The composite:
  spectral residual (Hou & Zhang 2007, numpy FFT) + Lab center-surround contrast +
  frame-diff motion (+ opt-in Haar faces, `SCREENSPACE_ATTENTION_FACE_CHANNEL`,
  default off) × a center-weighted prior, EMA-smoothed across samples. All pure
  cv2/numpy in `screenspace_primitives.py`.
- **Output: heatmaps + shift events.** Every sampled frame contributes a
  `saliency_grid` (flow_grid-shaped) to the heatmap pipeline — static PNG,
  accumulation GIF, and the rolling-window GIF (the eye-tracking-style gaze replay).
  Timeline events fire only at **attention shifts**: the smoothed saliency peak
  jumped ≥ `shift_threshold` (normalized) and persisted `SHIFT_CONFIRM` samples.
- **Full-frame only**, like Boundary (forced `{"source": "full_frame"}` region_ref
  server-side and in the CLI; region picker hidden in the UI).
- **No phash-skip / static-skip**: dwell weighting requires static frames to keep
  accumulating heat, and the EMA + shift-confirm counters assume uniform sampling.
  `supports_fast_scan = False`.

Key mechanism: the scan **returns** all samples (heatmap input, `t["result"]`) but
**streams** only confirmed shifts via `on_result` (`t["_raw_results"]` → events).
Post-heatmap the worker replaces the task's visible results with the shift subset,
so the timeline/results panel show one tick per shift, not per sample.

## Done

- ✅ **Primitives + config** — `compute_spectral_residual` / `compute_color_contrast`
  / `compute_motion_saliency` / `compute_face_saliency` / `compute_saliency_map` /
  `saliency_grid_from_map` / `saliency_peak` in `screenspace_primitives.py`;
  `SCREENSPACE_ATTENTION_*` constants + `SCREENSPACE_GENERATE_ATTENTION_HEATMAP`
  (Settings → Screenspace → Heatmaps); attention in `SCREENSPACE_MASK_FALLBACK_TOOLS`.
- ✅ **Scan + tool** — `scan_attention` (shift state machine, boundary-style
  full-frame opts) in `screenspace_scans.py`; `AttentionTool` registered in `TOOLS`
  (scan-only: no multitool step, not calibratable); `_extract_confidence` branch.
- ✅ **Heatmaps + worker/events** — `_GRID_KEYS` generalization + `generate_attention_heatmap`
  in `screenspace_heatmap.py`; worker branch probes video dims, writes rolling GIF,
  filters completed results to shifts, strips `saliency_grid` from reads/manifest;
  event metadata (`shift_distance`, `from/to`, `peak_value`) in `screenspace_manifest.py`.
- ✅ **Server + CLI** — `_VALID_TASK_TYPES` + forced full-frame rewrite widened to
  attention; `--ss-task attention P01` (region ignored with warning; `--ss-threshold`
  → `shift_threshold`).
- ✅ **Frontend** — Attention tab (eye icon, lime `--color-task-attention`), params
  Sensitivity/Smoothing/Interval, full-frame Run gating, shift result rows (Δ),
  heatmap section + full-canvas overlay, timeline tooltip, detector palette mirrors,
  viewer.js 16px eye icon for exports.
- ✅ **Model view** — `_preview_attention` panel strip (input/spectral/contrast/motion/
  combined) + frame-scoped `saliency_map` overlay layer in `screenspace_preview.py`.
  The preview route decodes a companion frame at the attention interval (the motion
  panel was dead until `screenspace_server.py` added attention to the prev-frame
  tuple) and both preview paths honor per-task weight overrides.
- ✅ **Weight tuning interface** — per-task sliders (Spectral/Contrast/Motion/Faces
  weights + Center bias; Faces at 0 disables the channel) flow as `weight_*` /
  `center_bias` params through `saliency_kwargs_from_params()` into both the scan
  and the live Model-view preview, so tuning is visual before committing a run.
  Config defaults still apply when the params are absent (CLI/API callers).
- ✅ **Tests** — `tests/screenspace/test_attention.py` (primitives, scan state machine,
  heatmaps, events, worker end-to-end) + API/CLI/preview suite extensions.

## Open steps

> **Archived 2026-07-25.** All implementation landed; the two items below were
> closed as accepted-as-is rather than completed — browser verification and the
> defaults tuning pass were never formally run. Re-open if the Attention tool's
> defaults turn out to need work on real footage.

### 1. Browser verification  ⏳
Manual check (no headless browsers — repo rule): tab renders with lime hue + eye icon
and no region picker; a run on a real recording streams shift ticks live; on completion
the results list shows only shifts with Δ scores, the heatmap section shows
Static/Accumulation/Rolling, and "Overlay on Frame" covers the whole frame; model view
shows the channel strip; light theme; an exported events viewer renders the markers.

### 2. Tuning pass on real footage  ⏳
Defaults (`shift_threshold` 0.15, EMA 0.6, weights 1.0/0.7/1.2, center bias 0.25) were
chosen analytically. Validate on real usability footage: shift-event density (a
10-minute session should mark dozens of shifts, not hundreds), rolling-GIF legibility,
and whether the motion weight over-dominates on high-motion game footage.

## Ideas (not committed)

- Per-event thumbnail crops around the from→to focus points in the results panel.
- Workflows node for Attention (chain: new video → attention scan → heatmap export).
- Optional DeepGaze-style learned backend behind the same result contract (would need
  a Whisper-style confirm-before-download flow and a direct torch dependency).
- Attention-vs-Screenspace-event cross-correlation on the Overview page (does
  predicted attention coincide with detected UI events?).

## Gotchas

- Pause/resume: `_partial_results` re-seeds from the shift-filtered `t["result"]`,
  so a resumed scan's heatmap under-accumulates the pre-pause segment slightly.
  Accepted (boundary has analogous behavior).
- The saliency map is deliberately **not** re-normalized per frame (peak_value stays
  comparable across frames); only `saliency_grid_from_map` normalizes per frame
  (dwell weighting: one unit of attention per sample).
- `_backfill_missing_events` regenerates from `task["result"]` — the defensive
  shift filter in `generate_events_from_results` keeps that path event-correct.
- opencv-python-headless 5.x wheels removed the legacy `cv2.CascadeClassifier`
  (CI resolves 5.x; local lockfile has 4.x). The face channel feature-detects it
  (`face_detection_available()`): degrades to a zeros map, keeps the face weight
  out of the mix denominator, and `scan_attention` warns once when the channel
  is requested but unsupported. Never name `cv2.CascadeClassifier` in
  annotations or call it unguarded.
