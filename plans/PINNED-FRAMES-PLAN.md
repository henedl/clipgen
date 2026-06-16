# Pinned Frames & Detector Calibration Plan

## Overview

Threshold tuning is currently the slowest loop in Screenspace. Researchers guess a threshold, run a full-video scan, inspect results, and repeat. Fast scan reduces the cost of each iteration but the loop remains O(video duration) per parameter guess.

**Pinned frames** invert this: the researcher pins a small set of frames they already understand — frames where the condition *is* true (positives) and frames where it *must not* fire (negatives) — and tunes parameters against those frames with instant feedback. Calibration is O(pinned frames), i.e. milliseconds. The full-video scan becomes a confirmation step, not a search procedure.

The core UI is not "iterate until green." Every per-frame tool already reduces to a monotonic scalar (SSIM score, change magnitude, fuzzy ratio, color distance). The calibration strip shows each pin's **score as a dot positioned against the threshold control**, color-coded positive/negative. The researcher reads the gap between the two populations and places the cutoff inside it in one glance. Overlapping populations are equally informative: they reveal that the region or tool choice is wrong — something no amount of threshold tweaking would surface.

### Use cases driving this work

1. **Threshold placement** — "the health bar is red here and not here; what tolerance separates them?"
2. **Region validation** — overlapping positive/negative score populations expose contaminated or badly-drawn regions before any scan is run
3. **Regression anchoring** — re-running a tuned task on re-encoded or replacement footage; pins verify the parameters still hold before a full scan
4. **OCR sanity-checking** — see what the text/numbers tools actually read on a known frame while adjusting Enhance ROI / Normalize / Min OCR conf.

## Key architectural decisions

- **Pins are first-class manifest objects.** New `pins` key in `screenspace_manifest.json`, per participant. A pin is a timestamp plus a polarity label — *not* a stored image. Frames are fetched on demand through the existing frame API, so re-encoded source videos invalidate naturally via the existing `?v=` mtime versioning.
- **Pins are tool-agnostic.** A pin marks "this frame matters," not "this frame matters for the color tool." The same pin set calibrates any tool the researcher switches to. Polarity (`positive` / `negative`) is the only semantic attached.
- **Evaluation reuses the single-frame path.** [`screenspace.py`](screenspace.py) `check_frame_for_tool` / `AnalysisTool.check_frame` already evaluates one frame against one tool's parameters. Calibration is a loop over pins calling into this machinery. The only detector-level change is exposing the per-tool scalar on both branches so a threshold-independent score is available (see the Score decision below); no *new* metrics are computed.
- **Score, not pass/fail, is the primary output.** `check_frame` returns a boolean at the current threshold; the calibration strip needs the underlying scalar *independent of threshold* so dots stay stable while the slider moves. Today each tool's `check_frame` returns `(False, None)` on a miss — the scalar lives in the detail dict *only on pass* — so the scalar cannot be read by simply calling `check_frame`. **This is real detector work, not a free loop:** each tool's `check_frame` is refactored to populate the scalar in the detail dict on **both** the pass and fail branches, and a new `score_frame` method reads it (see Phase 2). The boolean return is preserved, so existing `scan_*` / `scan_multitool` callers are unaffected.
- **Per-frame criteria only.** Temporal parameters (`require_consecutive`, interval, detect-first) cannot be validated on single frames and are explicitly out of calibration scope. The UI greys them out in the strip's coverage note. This is acceptable — per-frame criteria are where the rerun pain lives.
- **Synchronous, debounced evaluation.** Calibration requests do not go through the task queue. A dedicated endpoint evaluates ≤ N pins synchronously (mirroring the model-view preview endpoint's request pattern), with the frontend debouncing on parameter input exactly as `refreshModelView({debounce: true})` does today.

## Score definition per tool

The calibration scalar must be monotonic in "matchiness" and independent of the threshold being tuned. Tolerance-relative confidences (color) are recomputed against the *current* tolerance, so those dots legitimately move when tolerance changes — that is correct behavior, since the question is "does this frame pass at these settings."

**Every tool renders on a normalized 0–1 "matchiness" axis** (not the raw threshold-control range). For unit-matched tools the scalar and the threshold share units and are both normalized via the frontend slider's min/max. For tools whose scalar is not in threshold units, the axis still applies but the cutoff is tool-specific (see the table notes): **color** draws the cutoff where the tool's boolean `matched` flips (the strip's real value there is population *separation*, use case #2); **numbers** plots the best OCR confidence of any reading and treats the `operator`/`target_value` comparison as the polarity *expectation* (recorded in `passed`, not as the axis); **scene** uses the winning reference's per-scene threshold.

| Tool | Calibration scalar | Threshold control it maps to |
|---|---|---|
| color | normalized HSV distance → confidence (existing `color_matches` conf) | Tolerance |
| change | change ratio (`magnitude`) vs companion frame | Threshold |
| similarity | raw SSIM score vs reference | Threshold |
| text | best fuzzy ratio among OCR readings ≥ min conf | Fuzzy Thr. |
| numbers | best OCR conf among readings (operator/target_value → polarity *expectation*, recorded in `passed`) | Min OCR conf. |
| template | best match score | Threshold |
| flow | flow magnitude vs companion frame | Magnitude |
| scene | best fingerprint similarity (with winning scene name) | per-scene threshold |
| inactivity | phash distance vs companion frame (inverted) | Sensitivity |
| multitool | per-step scalars; strip renders one row per step | per-step thresholds |
| timelapse | n/a — excluded from calibration | — |

**Companion frames.** change / flow / inactivity compare against a prior frame. A pin stores only its own timestamp; the companion is fetched at `ts − interval` (the workflow's current interval param) at evaluation time. Consequence: changing the interval changes these tools' scores. The strip surfaces this with the interval value in the row label rather than hiding it. Pins where `ts < interval` have no companion frame and return `not_evaluable` (hollow dot).

## Phase 1: Pin storage and management

- [x] Manifest: `pins` key — `{participant: [{id, timestamp, polarity, label?, created_at}]}`; `id` format `pin_<8hex>`
- [x] API: `GET/POST /screenspace/api/pins/<participant>`, `DELETE /screenspace/api/pins/<pin_id>`, `PUT /screenspace/api/pins/<pin_id>` (polarity toggle, label edit)
- [x] Viewer UI: pin button pair (✓ / ✗) in the region bar pins the current frame as positive/negative
- [x] Pin tray: thumbnail strip (existing frame API at small size) below the viewer; click seeks, hover highlights the timeline tick, × removes
- [x] Timeline: pin ticks rendered above the result band (distinct glyph; green/red by polarity)
- [x] Soft cap (config `SCREENSPACE_MAX_PINS`, default 12) — calibration must stay interactive
- [x] Manifest round-trip: pins survive server restart; pins referencing timestamps beyond a replaced video's duration are flagged stale in the tray

## Phase 2: Backend evaluation

- [ ] Refactor each calibratable tool's `check_frame` (`ColorTool`…`InactivityTool`) to populate the scalar in the detail dict on **both** branches (today it's `(False, None)` on a miss); keep the boolean return so `scan_*`/`scan_multitool` callers are unaffected. Re-verify `_extract_confidence` and `on_result` callers still read the same keys
- [ ] `AnalysisTool.score_frame(frame, prev_frame, region, params) -> {score: float, passed: bool, detail: dict}` on the base class, reading the scalar `check_frame` now always returns; `passed` = the boolean at current params. Filter `math.isfinite` on `score` (numpy/OpenCV floats) before returning
- [ ] Multitool calibration evaluates **every** step unconditionally (a separate path from `scan_multitool`, which short-circuits the AND chain) so every step row gets a score; return per-step scores plus an overall chain `passed`
- [ ] `POST /screenspace/api/calibrate` — body: `{participant, tool, parameters, region_ref, pin_ids?}`; returns per-pin `{pin_id, score, passed, detail}`; evaluates synchronously
- [ ] Frame fetches go through the existing frame-extraction path; add a small per-participant LRU of decoded pin frames keyed on `(timestamp, video_version)` so slider drags don't re-decode
- [ ] OCR tools: cache raw `reader.readtext` output per `(pin, region, preprocess-flags, video_version)` — fuzzy/conf threshold changes then re-score from cached readings without re-running EasyOCR
- [ ] Binary-parameter tools (similarity reference, template image, scene refs) resolve references exactly as task creation does
- [ ] Degenerate inputs (zero-size region, missing reference) return a structured `not_evaluable` status per pin, not a 500

## Phase 3: Calibration strip UI

- [ ] New collapsible "Calibration" panel in the workflow body, beside Model view; visible whenever ≥ 1 pin exists for the selected participant
- [ ] Score strip: normalized 0–1 horizontal axis (scalar and threshold normalized via the tool's slider min/max); one dot per pin (green = positive, red = negative, hollow = `not_evaluable`); threshold cutoff drawn as a vertical line that tracks the slider live
- [ ] Dot hover: tooltip with timestamp, score, and tool `detail` (e.g. OCR text found); click seeks the viewer
- [ ] Pass/fail summary chip: "5/5 positives pass · 0/4 negatives pass" — turns green only when all positives pass *and* all negatives fail
- [ ] Debounced re-evaluation on any parameter input (reuse the model-view debounce pattern, 150 ms)
- [ ] Multitool: one strip row per step; summary chip reflects chain logic including NOT steps
- [ ] Coverage note when temporal params are set: "Consecutive/interval settings are not validated by calibration"

## Phase 4: Workflow integration

- [ ] "Run" button affordance: when calibration is green, subtle indicator on Run ("calibrated"); when red, no blocking — researcher agency preserved
- [ ] Restore-task-to-workflow (`restoreTaskToWorkflow`) re-triggers calibration so edited tasks immediately show whether saved params still satisfy pins
- [ ] Export: pins included in the events JSON/CSV export as a separate section (provenance for "why this threshold")

## Open questions (not blockers)

- **Region edits during calibration.** Dragging a region should re-trigger evaluation like any parameter change; needs a hook from `saveRegionUpdate` into the strip's debounce.
- **EasyOCR cold start.** First text/numbers calibration loads the OCR model (~seconds). Show a one-time loading state in the strip rather than blocking the panel.
- **Cross-participant pins.** Game UIs are positionally consistent, so pins from P01 could plausibly calibrate a region for P02 — deferred until single-participant calibration proves out.

## Future work (revisit)

- **Unify reference frames under pins.** Similarity references, template captures, and scene references are all ad-hoc "frames that matter." A shared pinned-frame object could back all of them; deferred until the pin schema has stabilized.
- **Suggested threshold.** With clearly separated populations, the midpoint of the gap is a trivially computable suggestion. Held back from v1 deliberately: the tool surfaces signal, the researcher places the cutoff.
- **Boundary frames as pin candidates.** Once automated scene boundaries land (see SCENE-BOUNDARIES-PLAN.md), boundary events give researchers a candidate set of distinctive frames to pin.
