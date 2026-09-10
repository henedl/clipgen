# Drop OpenCV Plan

**Status: unscheduled, optional future work package.** Investigated 2026-08-16; recorded so a
future session inherits the evidence instead of re-deriving it. The near-term route is the
self-compiled `WITH_FFMPEG=OFF` wheel — researched and de-risked the same day, see
[OPENCV-SELF-COMPILE-PLAN.md](OPENCV-SELF-COMPILE-PLAN.md) — which fixes the license (and −76 MB
of the bundle) for ~1/8 the effort but keeps cv2's remaining size, cold import, and CI guard
burden.

**2026-08-17 update: the premise below no longer holds.** The maintainer decided to keep cv2
indefinitely, and scikit-image, scipy, PyWavelets, and imagehash have since been *removed* from
the dependency tree (replaced by in-tree `screenspace_primitives.structural_similarity` and
`PHash`, built on cv2; guarded by `tests/test_dropped_dependencies.py` and a build-binaries.yml
bundle check). Executing this plan now would require re-adding scipy + scikit-image and reverting
those guards, on top of the effort estimated below. Plan stays for the evidence, not as a route.

## Context

The macOS arm64 `opencv-python-headless` wheel bundles GPL-3 ffmpeg dylibs
(opencv/opencv-python#1260), forcing the DMG to be conveyed GPL-3.0-or-later; the dylibs cannot
be stripped post-hoc ([LICENSE-PLAN.md](LICENSE-PLAN.md) has the three reasons). This plan is the
other escape route: remove the cv2 dependency entirely by reimplementing its used surface on
packages already in the tree — numpy, Pillow, scipy + scikit-image (SSIM already comes from
skimage), shapely + pyclipper (already rapidocr deps). **No new dependencies.**

Verdict from the investigation: **feasible, ~13–16 focused sessions across 6 PRs.** Payoff beyond
the license: −120 MB unpacked bundle (`cv2/` is 120 MB, 75 MB of it dylibs), the ~10 s cv2 cold
import gone (`utils.preload_vision_libs_quietly` deleted), and the build-binaries.yml license
guard + upstream-bug exposure eliminated permanently on all platforms.

### Findings the plan rests on

- **clipgen's own cv2 use is Screenspace-only**: 12 modules, 66 distinct symbols; hot spots
  `screenspace_primitives.py` (82 refs) and `screenspace_preview.py` (105 refs, mostly cosmetic
  drawing/colormap/resize). No `cv2.VideoCapture` anywhere — decoding is ffmpeg-piped, so the
  bundled ffmpeg is 100% dead weight. Only three genuinely hard algorithms: Farneback optical
  flow, masked `TM_CCOEFF_NORMED` template matching, and Canny — plus the Haar face cascade,
  which is default-off and already degrades to zeros on OpenCV 5 wheels.
- **RapidOCR is the hidden blocker**: it imports cv2 at module scope in 10 modules (~28 symbols
  incl. `findContours`, `minAreaRect`/`boxPoints`, `warpPerspective`). Verified in its installed
  sources (3.9.2): the DB postprocess consumes only the *derived box* through shapely/pyclipper it
  already ships, and `score_mode="fast"` (the default) scores the box, not the raw contour — so
  `findContours` can return **unordered per-blob boundary pixels** (`skimage.measure.label` +
  label-minus-erosion) and `minAreaRect` is `shapely.oriented_envelope`. No Suzuki–Abe contour
  tracer needed. This is the load-bearing simplification of the whole plan.
- **Perf-critical sites with documented cv2 wins** that must be benchmarked, not assumed:
  `mean_gray_diff` (`cv2.norm NORM_L1` fused, exact-equality test), `color_present`'s `inRange`
  band path (~12× note vs float32 numpy), `_area_resize` (~14× note on the phash path).
- **Tests holding verbatim cv2 oracles** that must be rewritten in lockstep:
  `tests/screenspace/test_static_gating.py`, `test_color.py`, `test_change_similarity.py`,
  `test_attention.py`, `test_template.py`, plus PNG-decode helpers in `tests/test_screenspace_api.py`
  and `tests/test_screenspace_preview.py`.

## Architecture

Two new modules (both added to `pyproject.toml [tool.setuptools] py-modules`):

1. **`source/imageops.py`** — clipgen's native primitive layer, BGR uint8 ndarrays in/out (no
   call-site data-model churn):
   - Conversions `bgr_to_gray/hsv/lab`, `hsv_to_bgr` as **cv2-bit-exact uint8 fixed-point
     formulas** (H∈[0,180], cv2 Lab constants) — HSV targets persist in user manifests, so
     bit-exactness is non-negotiable. Frozen golden tests from literals captured off cv2 4.13
     while it is still installed; the tests never import cv2.
   - `resize` via PIL (AREA→`Image.BOX`, LINEAR/NEAREST/CUBIC→BILINEAR/NEAREST/BICUBIC);
     `area_resize_fast` via the numpy reshape-mean trick for integer ratios (expected to *beat*
     cv2's two-pass trick on the phash path).
   - `gaussian_blur` with cv2's sigma-from-ksize derivation (`0.3*((k-1)*0.5-1)+0.8`) on
     `scipy.ndimage.gaussian_filter1d` separable passes, `mode="mirror"` (= BORDER_REFLECT_101);
     `box_blur` via `uniform_filter`.
   - Pixel ops: `absdiff` via uint8 `np.maximum(a-b, b-a)` (no dtype promotion), `mean_abs_diff`
     (integer sum ÷ size in double, preserving the exact-equality contract), `in_range`,
     threshold, bitwise, normalize, min_max_loc, cart_to_polar/magnitude, copy_make_border.
   - Morphology via `scipy.ndimage` min/max filters; `fill_poly`/line/rectangle/ellipse/
     arrowed_line/put_text via PIL.ImageDraw; `disk_splat` with a cached boolean disk mask and
     **set-not-accumulate assignment** (the heatmap's "last draw wins" replay semantics are
     load-bearing, radius is a constant `acc_size//16`).
   - Frozen bit-identical `JET_LUT` (256×3 uint8 literal captured from `cv2.applyColorMap`);
     PNG/JPEG codecs via PIL (keep `_write_png`'s OSError semantics); 3-D HSV histogram
     replicating `calcHist` bin edges + Pearson correlation (= HISTCMP_CORREL);
     `dct2` = `scipy.fft.dctn(type=2, norm="ortho")` (≡ `cv2.dct`).
   - Algorithms: `canny` (skimage.feature.canny), `clahe` (`skimage.exposure.equalize_adapthist`,
     8×8-tile-equivalent kernel), `match_template_masked` (Padfield masked ZNCC on
     `scipy.fft.rfft2`; `mask=None` reduces to standard ZNCC ≡ TM_CCOEFF_NORMED),
     `tile_flow` (see hard algorithms), spectral-residual internals on scipy `fft2`/complex64.

2. **`source/cv2_shim.py`** — cv2-namespace adapter exclusively for rapidocr (~28 symbols),
   delegating to `imageops` wherever the computation exists ("one computation, one
   implementation"); owns the perspective/contour/minAreaRect pieces clipgen itself never needs
   (`getPerspectiveTransform` = 8-unknown `np.linalg.solve`; `warpPerspective` via
   `scipy.ndimage.map_coordinates` `mode="nearest"` = BORDER_REPLICATE, more faithful than PIL's
   black fill for rotated text-crop rectification; `findContours` returning `(N,1,2)` int32
   boundary-pixel arrays in cv2's `(x,y)` order; `minAreaRect`/`boxPoints` via shapely, converted
   to cv2's `((cx,cy),(w,h),angle)` tuple — note rapidocr's `get_mini_boxes` reads
   `min(rect[1])`; mask-aware 4-tuple `mean`; saturating `add`; rotate, polylines, imwrite).

   **Injection**: `screenspace_ocr._ensure_cv2_stub()` — direct analogue of
   `transcripts._ensure_av_stub` — does `sys.modules.setdefault("cv2", cv2_shim)` immediately
   before the lazy `from rapidocr import ...` in `_build_ocr_reader`. `setdefault` means a dev
   venv with a real cv2 keeps working during the transition; once opencv leaves `uv.lock` the
   shim always wins. No file is ever named `cv2.py`, so nothing shadows a real install. The shim's
   module `__getattr__` raises a named error (`"cv2 shim: rapidocr needs '<name>'"`) on rapidocr
   drift. The `[tool.uv]` override `"opencv-python ; sys_platform == 'never'"` **stays** —
   rapidocr still declares it; only `opencv-python-headless` leaves `[project] dependencies`.

## Phases (each a separate green PR; cv2 stays installed until Phase 6)

1. **`imageops.py` core + golden tests + baseline benchmarks** (2–3 sessions). No call-site
   changes. Record cv2 perf baselines (profile skill) while cv2 is still importable.
2. **Mechanical migrations** (2): `screenspace_server.py` (imdecode/imencode),
   `screenspace_frames.py` (resize clamp), `screenspace_scans.py`, `screenspace_tools.py`,
   `screenspace_heatmap.py` entirely, `screenspace_preview.py`'s ~100 cosmetic refs (the
   Farneback/Canny preview sites migrate with the primitives in Phase 4), `cli.py`
   `_ss_hex_to_hsv`. Retarget test monkeypatches (`screenspace_heatmap.cv2.imencode` →
   `imageops.encode_png`) and PNG-decode helpers to PIL.
3. **Primitives hot paths** (2–3): `mean_gray_diff`, `color_present`/`_hsv_match_bands` (band
   architecture kept verbatim), blurs, `_frame_diff_mask`, phash (`area_resize_fast` + `dct2`),
   scene fingerprint (hist + Pearson + canny; numbers change — safe, references are recomputed
   per scan, never persisted), `region_mask_for`. Rewrite the cv2 oracles in lockstep —
   exact-equality contracts in `test_static_gating`/`test_color` are *preserved* (integer
   arithmetic on both sides). Benchmark loop against Phase 1 baselines.
4. **Hard algorithms + attention** (3):
   - **Template**: Padfield masked ZNCC, mask support kept (alpha-masked templates are a real
     product feature with dedicated degenerate handling). Scores drift in the third decimal,
     ranking preserved; NMS/candidate-cap/degeneracy handling untouched.
   - **Flow**: replace Farneback with `tile_flow` — per-tile phase correlation (windowed
     cross-power spectrum via one batched 3-D FFT over the ~16×16 grid tiles, argmax + parabolic
     refinement). The dense Farneback field is only ever reduced to per-tile grid aggregates
     (`compute_optical_flow` → mean magnitude, dominant angle, `flow_grid`), so this is
     deterministic, ~80 lines, and likely *faster*. skimage's `optical_flow_ilk`/`tvl1` were
     evaluated and rejected: hundreds of ms to seconds per pair at 256², per-frame-scan poison.
   - **Attention**: delete the cv2 dft branch (+`_dft_friendly`); the existing numpy
     spectral-residual branch is the surviving implementation, moved to scipy fft2/complex64.
     Lab center-surround on `imageops.bgr_to_lab`. **Drop the face channel end-to-end**: code
     (`_get_face_cascade`, `compute_face_saliency`, weights), config keys
     (`SCREENSPACE_ATTENTION_WEIGHT_FACE`/`_FACE_CHANNEL`), the `paramAttnWFace` slider
     (screenspace.html/js + model-view map), and the CascadeClassifier tests.
5. **OCR + shim** (2–3): CLAHE swap in `_preprocess_for_ocr`; `cv2_shim.py` with per-symbol unit
   tests; a guard test importing all 10 cv2-importing rapidocr modules under the shim (fails
   loudly on an unshimmed import-time symbol); an integration test instantiating rapidocr's
   `DBPostProcess` directly on a synthetic probability map (no ONNX) asserting box corners within
   1–2 px. **Manual gate before Phase 6**: run the Text tool on real fixture footage with cv2
   uninstalled, diff readings against a cv2 baseline run.
6. **Cut-over + build/CI/license** (1–2): drop the dep from `pyproject.toml`; delete
   `utils.preload_vision_libs_quietly` + its server call + boot-page "cv2" phase; `clipgen.spec`
   `excludes += ["cv2"]` (a stray dev-machine opencv can never leak into the bundle); invert the
   build-binaries.yml guard — fail if any `cv2/` directory or non-`bin/` `libav*` dylib appears
   in the bundle (same shape as the existing PyAV guard); remove the OpenCV/FFmpeg-dylib GPL
   sections from `build/THIRD-PARTY-LICENSES` (the pinned ffmpeg/ffprobe *executables* section
   stays); DMG licensing text back to MIT-app + GPL-aggregated-executables; update
   [LICENSE-PLAN.md](LICENSE-PLAN.md) (supersede the OpenCV decision, close the #1260-tracking
   and self-compile items as moot); add `tests/test_no_cv2.py` (greps `source/` for cv2 imports
   allowlisting the shim, asserts `import screenspace` leaves no real `"cv2"` in `sys.modules`);
   sweep docs (AGENTS/ARCHITECTURE/PERFORMANCE cv2 mentions, `screenspace.py` docstring).

## User-visible behavior changes (accepted; no compat layers per the hard rule)

- **Flow tool**: magnitude scale changes (per-tile displacement vs Farneback per-pixel mean) —
  saved Flow thresholds need re-calibration via the existing calibration UI.
- **Template tool**: scores drift in the third decimal; ranking preserved.
- **Attention tool**: loses the optional face channel (default-off, already broken on OpenCV 5).
- **Scene/OCR**: fingerprint numbers change (self-healing, recomputed per scan); CLAHE output
  visually differs slightly.

## Performance strategy

Baseline before Phase 1 with the profile skill; re-measure per phase. Acceptance gate:
end-to-end scan wall-clock per tool on the fixture within ~+15% of the cv2 baseline (keyframe
decode dominates all tools, bounding the blast radius). Expectations: `mean_gray_diff` 2–5×
slower in isolation but sub-ms at working sizes; `color_present` target ≤3× with an integer
channel-decomposed HSV-test reformulation held in reserve (never materializes the HSV image);
phash resize expected to win; tile_flow expected to win; template FFT ≈ parity; OCR shim cost is
noise next to ONNX inference. Side benefits to record: ~10 s cold start drop, one fewer native
thread pool under `SCREENSPACE_PARALLEL_WORKERS`.

## Top risks

1. **Per-frame perf regression on the scan core** → baselines first, +15% gate, uint8
   reformulations in reserve.
2. **OCR box-quality regression through the shim** → boundary-pixels + `oriented_envelope` is
   ±0.5 px of cv2 geometrically; synthetic DBPostProcess test + mandatory real-footage diff.
3. **Saved thresholds invalidated by number drift** (flow, template) → accepted; users
   re-calibrate; scene references self-heal.
4. **rapidocr drift outgrowing the shim** → existing `<4` pin, import-coverage guard, named
   `__getattr__` errors.
5. **Convention mismatch leaking into persisted state** (HSV in manifests, region masks) →
   bit-exact conversions with frozen goldens; fill_poly tolerance confined to ±1 px.

## Recommended shape if picked up

Land Phases 1–3 unconditionally (pure wins, low risk, exact-equality preserved), then a go/no-go
on Phases 4–6 with the Phase 3 benchmarks in hand. Phases 1–3 also stand alone as prep that
makes a later full drop cheap, even if the self-compile route lands first for the license.

## Verification

- Per phase: the /check pipeline (ruff format + lint, ty, full test suite).
- Phase 1/3: benchmark script vs recorded cv2 baselines (profile skill).
- Phase 4: synthetic-translation flow-recovery tests; planted-icon template property tests.
- Phase 5: DBPostProcess synthetic-map test; manual Text-tool diff on real footage, cv2
  uninstalled.
- Phase 6: `tests/test_no_cv2.py`; /ui-check across all six pages; a full desktop build
  confirming no `cv2/` in the bundle and the inverted CI guard passing.
