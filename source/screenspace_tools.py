"""Screenspace analysis tools (strategy registry) + per-frame dispatch.

Each tool is a small class wrapping the module-level ``scan_*`` function (kept
standalone so tests can monkeypatch them). Two dispatch points look a tool up in
``TOOLS`` and delegate: :func:`check_frame_for_tool` (single-frame, used by the
multitool chainer and pin calibration) and ``ScreenspaceWorker._dispatch``
(full-video scan).

The one tools->multitool edge (``MultitoolTool.scan`` -> ``scan_multitool``) is a
function-local import, keeping the module graph acyclic.
"""

import functools
import math
from collections.abc import Callable
from pathlib import Path
from typing import Any, ClassVar

import cv2
import numpy as np

import config
import utils
from screenspace_primitives import (
    _frame_edge_map,
    _mask_corr_outside_window,
    _prepare_shape_reference,
    _prepare_template,
    _scale_template,
    _template_correlation_map,
    blur_gray,
    color_matches,
    color_present,
    compare_scene_fingerprints,
    compute_frame_diff_gray,
    compute_optical_flow,
    compute_phash,
    compute_scene_fingerprint,
    extract_region,
    filter_matches_by_region_mask,
    mask_points_key,
    match_shape,
    match_template,
    region_mask_for,
    region_search_window,
    regions_are_similar,
    saliency_kwargs_from_params,
)
from screenspace_ocr import (
    _ocr_region_readings,
    _score_numbers_readings,
    _score_text_readings,
)
from screenspace_scans import (
    generate_timelapse,
    scan_attention,
    scan_boundaries,
    scan_changes,
    scan_color,
    scan_flow,
    scan_inactivity,
    scan_numbers,
    scan_scene,
    scan_shape,
    scan_similarity,
    scan_template,
    scan_text,
)


def _extract_confidence(tool_type: str, result: dict[str, Any]) -> float:
    """Extract a normalized [0, 1] confidence from a tool-specific result dict."""
    if tool_type == "color":
        return result.get("_confidence", 1.0)
    elif tool_type == "change":
        return min(result.get("magnitude", 0.0), 1.0)
    elif tool_type == "similarity":
        return result.get("score", 0.0)
    elif tool_type == "text":
        return result.get("confidence", 0.0)
    elif tool_type == "numbers":
        return result.get("confidence", 1.0)
    elif tool_type == "template" or tool_type == "shape":
        return result.get("best_score", 0.0)
    elif tool_type == "flow":
        return min(result.get("magnitude", 0.0) / 10.0, 1.0)
    elif tool_type == "scene":
        return result.get("score", 0.0)
    elif tool_type == "multitool":
        return result.get("min_confidence", 0.0)
    elif tool_type == "inactivity":
        if "_confidence" in result:
            return float(result["_confidence"])
        return min(result.get("duration", 0.0) / 30.0, 1.0)
    elif tool_type == "boundary":
        if "_confidence" in result:
            return float(result["_confidence"])
        thr = config.SCREENSPACE_BOUNDARY_PHASH_THRESHOLD
        dist = result.get("distance", 0.0)
        if thr <= 0:
            return 1.0
        return max(0.0, min((dist - thr) / float(thr), 1.0))
    elif tool_type == "attention":
        # Shift events carry _confidence; backfilled/raw samples fall back to
        # the peak strength.
        return float(result.get("_confidence", result.get("peak_value", 1.0)))
    return 1.0


# ---------------------------------------------------------------------------
# Per-frame memoization (multitool chains)
# ---------------------------------------------------------------------------
#
# scan_multitool passes per-frame cache/prev_cache; None disables memoization.

_MISSING = object()


def _region_key(prefix: str, region: dict[str, Any]) -> tuple[Any, ...]:
    """Cache key scoped to a region's pixel rect (steps may use distinct regions).

    Includes the shaped-region polygon so two steps sharing a bbox but not a
    shape never collide on masked values (mask, fingerprint, OCR readings).
    """
    return (
        prefix,
        region.get("x"),
        region.get("y"),
        region.get("w"),
        region.get("h"),
        mask_points_key(region.get("mask_points")),
    )


def _memo(
    cache: dict[Any, Any] | None, key: tuple[Any, ...], compute: Callable[[], Any]
) -> Any:
    """Return ``cache[key]``, computing and storing it on miss. ``None`` cache = no memo."""
    if cache is None:
        return compute()
    value = cache.get(key, _MISSING)
    if value is _MISSING:
        value = compute()
        cache[key] = value
    return value


def _cached_crop(
    cache: dict[Any, Any] | None, frame: np.ndarray, region: dict[str, int]
) -> np.ndarray:
    """Region crop of *frame*, memoized on *cache* (shared across chain steps)."""
    if cache is None:
        return extract_region(frame, region)
    key = _region_key("crop", region)
    value = cache.get(key, _MISSING)
    if value is _MISSING:
        value = cache[key] = extract_region(frame, region)
    return value


@functools.lru_cache(maxsize=64)
def _mask_raster_cached(
    points_key: tuple[Any, ...], h: int, w: int
) -> np.ndarray | None:
    """Rasterized shaped-region mask, cached across frames (shape + size constant)."""
    return region_mask_for({"mask_points": points_key}, h, w)


def _cached_mask(
    cache: dict[Any, Any] | None, frame: np.ndarray, region: dict[str, Any]
) -> np.ndarray | None:
    """Shaped-region mask at the cached crop's size, cached across frames.

    ``None`` for rect regions, so masked tools can pass the value straight to
    the primitives' optional ``mask`` params. The per-frame memo dict is fresh
    each frame, so the raster lives in :func:`_mask_raster_cached` instead —
    the polygon and crop size never change mid-scan. Callers must not mutate.
    """
    points = region.get("mask_points")
    if not points:
        return None
    crop = _cached_crop(cache, frame, region)
    return _mask_raster_cached(mask_points_key(points), *crop.shape[:2])


def _cached_gray(
    cache: dict[Any, Any] | None, frame: np.ndarray, region: dict[str, int]
) -> np.ndarray:
    """Grayscale of the region crop, memoized on *cache* (reuses the cached crop)."""
    if cache is None:
        return cv2.cvtColor(_cached_crop(cache, frame, region), cv2.COLOR_BGR2GRAY)
    key = _region_key("gray", region)
    value = cache.get(key, _MISSING)
    if value is _MISSING:
        value = cache[key] = cv2.cvtColor(
            _cached_crop(cache, frame, region), cv2.COLOR_BGR2GRAY
        )
    return value


def _cached_blur_gray(
    cache: dict[Any, Any] | None, frame: np.ndarray, region: dict[str, int]
) -> np.ndarray:
    """``blur_gray`` of the region crop, memoized on *cache* (reuses the crop).

    The multitool dispatcher rolls this frame's memo dict forward as the next
    frame's ``prev_cache``, so ChangeTool's previous-side blur+grayscale is a
    dict hit rather than a recompute.
    """
    if cache is None:
        return blur_gray(_cached_crop(cache, frame, region))
    key = _region_key("blurgray", region)
    value = cache.get(key, _MISSING)
    if value is _MISSING:
        value = cache[key] = blur_gray(_cached_crop(cache, frame, region))
    return value


def _cached_phash(
    cache: dict[Any, Any] | None, frame: np.ndarray, region: dict[str, int]
) -> "Any":
    """Perceptual hash of the region crop, memoized on *cache* (reuses the crop)."""
    if cache is None:
        return compute_phash(_cached_crop(cache, frame, region))
    key = _region_key("phash", region)
    value = cache.get(key, _MISSING)
    if value is _MISSING:
        value = cache[key] = compute_phash(_cached_crop(cache, frame, region))
    return value


def _cached_ocr(
    cache: dict[Any, Any] | None,
    frame: np.ndarray,
    region: dict[str, Any],
    *,
    languages: list[str],
    preprocess: bool,
) -> list[Any]:
    """OCR readings for the region crop, memoized on *cache*.

    Keyed by ``(region, languages, preprocess)`` so identical OCR calls dedup.
    Text and Numbers steps on one region deliberately share readings: the
    engine has no per-tool recognition mode, and every tool difference
    (fuzzy match, numeric operators, integers_only) is applied at scoring time.
    """
    langs = tuple(languages)
    if cache is None:
        return _ocr_region_readings(
            _cached_crop(cache, frame, region),
            languages=list(langs),
            preprocess=preprocess,
            mask_points=region.get("mask_points"),
        )
    key = _region_key("ocr", region) + (langs, preprocess)
    value = cache.get(key, _MISSING)
    if value is _MISSING:
        value = cache[key] = _ocr_region_readings(
            _cached_crop(cache, frame, region),
            languages=list(langs),
            preprocess=preprocess,
            mask_points=region.get("mask_points"),
        )
    return value


def check_frame_for_tool(
    frame: np.ndarray,
    prev_frame: np.ndarray | None,
    region: dict[str, int],
    tool_type: str,
    parameters: dict[str, Any],
    cache: dict[Any, Any] | None = None,
    prev_cache: dict[Any, Any] | None = None,
) -> tuple[bool, dict[str, Any] | None]:
    """Evaluate whether a single frame passes a tool's criteria.

    Used by :func:`scan_multitool` for steps 1+ in the chain. Returns
    ``(passed, result_dict)``, the dict carrying tool-specific metadata on a pass
    and ``None`` otherwise.

    **change** and **flow** require *prev_frame* (the frame immediately before
    the candidate timestamp); a ``None`` there skips the check as ``(False, None)``.

    ``cache`` / ``prev_cache`` are optional per-frame memo dicts letting chained
    steps that share a region reuse crop/gray/phash/OCR work; ``None`` (the
    calibration path) disables memoization.

    Degenerate regions (width or height ≤ 0, e.g. a 1-px user draw rounded
    to zero pixels on a small preview) are treated as a non-match here so
    downstream cv2 ops never see empty arrays.  Template and shape are exempt:
    they match against the full frame and ignore ``region``, so an uploaded
    reference with no region (zero-size region_coords) is valid — mirrors
    :func:`score_frame_for_tool`.
    """
    if tool_type not in ("template", "shape") and (
        region.get("w", 0) <= 0 or region.get("h", 0) <= 0
    ):
        return False, None
    tool = TOOLS.get(tool_type)
    if tool is None:
        return False, None
    return tool.check_frame(frame, prev_frame, region, parameters, cache, prev_cache)


def _score_result(
    score_key: str, passed: bool, detail: dict[str, Any] | None
) -> dict[str, Any]:
    """Wrap a ``check_frame``-style ``(passed, detail)`` into a calibration score.

    Reads the threshold-independent scalar at ``detail[score_key]``. Returns a
    ``not_evaluable`` status when the frame could not be scored (missing detail,
    empty/missing score key, or a non-finite scalar from numpy/OpenCV).
    """
    if detail is None or not score_key:
        return {"status": "not_evaluable"}
    raw = detail.get(score_key)
    if raw is None:
        return {"status": "not_evaluable"}
    score = float(raw)
    if not math.isfinite(score):
        return {"status": "not_evaluable"}
    return {"status": "ok", "score": score, "passed": bool(passed), "detail": detail}


def score_frame_for_tool(
    tool_type: str,
    frame: np.ndarray,
    prev_frame: np.ndarray | None,
    region: dict[str, int],
    params: dict[str, Any],
    *,
    ocr_reader: "Callable[[str, dict[str, int], dict[str, Any]], list[Any]] | None" = None,
) -> dict[str, Any]:
    """Score a single frame against one tool's parameters for pin calibration.

    Mirrors :func:`check_frame_for_tool` but returns the threshold-independent
    scalar (``{status, score, passed, detail}``) instead of a boolean. Degenerate
    regions and unknown / non-scorable tools (timelapse, multitool) return
    ``not_evaluable``.

    ``ocr_reader`` (text/numbers only) supplies cached OCR readings keyed per
    pin so fuzzy/confidence changes re-score without re-running OCR; when absent
    the tool runs OCR live through its ``check_frame``.
    """
    # Template and shape ignore region; every other tool crops, and empty crops break cv2.
    if tool_type not in ("template", "shape") and (
        region.get("w", 0) <= 0 or region.get("h", 0) <= 0
    ):
        return {"status": "not_evaluable"}
    tool = TOOLS.get(tool_type)
    if tool is None or not tool.score_key:
        return {"status": "not_evaluable"}
    if ocr_reader is not None and tool_type in ("text", "numbers"):
        if tool_type == "text" and not params.get("search_string", ""):
            return {"status": "not_evaluable"}
        readings = ocr_reader(tool_type, region, params)
        if tool_type == "text":
            passed, detail = _score_text_readings(readings, params)
        else:
            passed, detail = _score_numbers_readings(readings, params)
        return _score_result(tool.score_key, passed, detail)
    return tool.score_frame(frame, prev_frame, region, params)


class AnalysisTool:
    """Base class for screenspace analysis tools.

    Subclasses set ``name`` and override ``check_frame`` (for multitool
    chaining) and/or ``scan`` (for the full-video sweep). The dispatch
    layer reads ``fast_scan_region_dim`` / ``supports_fast_scan`` /
    ``fast_scan_extra_opts`` to build the shared ``fast_opts`` payload.
    """

    name: ClassVar[str] = ""
    # Fast-scan max region dimension (max_region_dim to the frame extractor); 0 = no downscale.
    fast_scan_region_dim: ClassVar[int] = 0
    # False opts out of fast scan (timelapse has its own sample_interval).
    supports_fast_scan: ClassVar[bool] = True
    # Extra keys merged into ``fast_opts`` (e.g. ``{"template_downscale": True}``).
    fast_scan_extra_opts: ClassVar[dict[str, Any]] = {}
    # Detail key of the threshold-independent calibration scalar; empty means not calibratable.
    score_key: ClassVar[str] = ""
    # scan_<x> name, resolved via globals() per call so test monkeypatches apply. Empty: subclass overrides scan.
    scan_fn_name: ClassVar[str] = ""
    # Tool-specific scan kwargs, forwarded as ``params.get(key, default)``.
    scan_defaults: ClassVar[dict[str, Any]] = {}
    # Sampling interval when ``params`` carries none (the OCR tools use 2.0).
    scan_interval_default: ClassVar[float] = 0

    def check_frame(
        self,
        frame: np.ndarray,
        prev_frame: np.ndarray | None,
        region: dict[str, int],
        params: dict[str, Any],
        cache: dict[Any, Any] | None = None,
        prev_cache: dict[Any, Any] | None = None,
    ) -> tuple[bool, dict[str, Any] | None]:
        """Evaluate a single frame for use in a multitool chain step.

        ``cache`` / ``prev_cache`` are optional per-frame memo dicts (see
        :func:`check_frame_for_tool`); subclasses that crop a region route through
        the ``_cached_*`` helpers so chained steps sharing a region reuse the work.

        Default: tool not supported as a multitool step.
        """
        return False, None

    def score_frame(
        self,
        frame: np.ndarray,
        prev_frame: np.ndarray | None,
        region: dict[str, int],
        params: dict[str, Any],
    ) -> dict[str, Any]:
        """Return the threshold-independent calibration score for one frame.

        Reads the scalar ``check_frame`` now populates on both branches (keyed by
        ``score_key``). Returns ``{"status": "not_evaluable"}`` when the frame
        cannot be scored (missing companion/reference, degenerate input, etc.).
        """
        passed, detail = self.check_frame(frame, prev_frame, region, params)
        return _score_result(self.score_key, passed, detail)

    def _scan_kwargs(self, params: dict[str, Any]) -> dict[str, Any]:
        """Tool-specific kwargs for the scan function, from ``scan_defaults``.

        Subclasses override to post-process (or-expression defaults, config
        reads, saliency merges). Dict-valued defaults are copied so a callee
        can never mutate the shared ClassVar.
        """
        kwargs: dict[str, Any] = {}
        for key, default in self.scan_defaults.items():
            value = params.get(key, default)
            if value is default and isinstance(default, dict):
                value = dict(default)
            kwargs[key] = value
        return kwargs

    def scan(
        self,
        video_path: str,
        region: dict[str, int],
        params: dict[str, Any],
        *,
        task_id: str,
        scan_mode: str,
        on_progress: Callable[[float], None],
        cancel_flag: Callable[[], bool],
        on_result: Callable[[dict[str, Any]], None] | None,
        fast_opts: dict[str, Any] | None,
    ) -> Any:
        """Run a full-video scan.

        Declarative dispatch: forwards the common tail plus ``_scan_kwargs``
        to the module-global named by ``scan_fn_name``. Tools whose scan
        differs structurally (template's downscale, timelapse's output file,
        multitool's chaining) override this outright.
        """
        if not self.scan_fn_name:
            raise NotImplementedError
        scan_fn = globals()[self.scan_fn_name]
        return scan_fn(
            video_path,
            region,
            interval_seconds=params.get("interval", self.scan_interval_default),
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
            **self._scan_kwargs(params),
        )


class ColorTool(AnalysisTool):
    name = "color"
    fast_scan_region_dim = 32
    score_key = "_confidence"
    scan_fn_name = "scan_color"
    scan_defaults: ClassVar[dict[str, Any]] = {
        "target_color": {"h": 0, "s": 0, "v": 0},
        "tolerance": {"h": 10, "s": 50, "v": 50},
        "color_mode": "average",
        "min_coverage": 0.0,
    }

    def check_frame(
        self, frame, prev_frame, region, params, cache=None, prev_cache=None
    ):
        pixels = _cached_crop(cache, frame, region)
        mask = _cached_mask(cache, frame, region)
        target = params.get("target_color", {"h": 0, "s": 0, "v": 0})
        tol = params.get("tolerance", {"h": 10, "s": 50, "v": 50})
        if params.get("color_mode") == "presence":
            matched, conf = color_present(
                pixels, target, tol, params.get("min_coverage", 0.0), mask=mask
            )
        else:
            matched, conf = color_matches(pixels, target, tol, mask=mask)
        return matched, {"_confidence": conf}

    def scan(
        self,
        video_path,
        region,
        params,
        *,
        task_id,
        scan_mode,
        on_progress,
        cancel_flag,
        on_result,
        fast_opts,
    ):
        # Presence needs full resolution: INTER_AREA downscaling erases small color patches.
        if params.get("color_mode", "average") == "presence":
            fast_opts = None
        return super().scan(
            video_path,
            region,
            params,
            task_id=task_id,
            scan_mode=scan_mode,
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class ChangeTool(AnalysisTool):
    name = "change"
    fast_scan_region_dim = 128
    score_key = "magnitude"
    scan_fn_name = "scan_changes"
    scan_defaults: ClassVar[dict[str, Any]] = {
        "threshold": 0,
        "noise_threshold": 0,
        "require_consecutive": 1,
    }

    def check_frame(
        self, frame, prev_frame, region, params, cache=None, prev_cache=None
    ):
        if prev_frame is None:
            return False, None
        threshold = params.get("threshold", config.SCREENSPACE_CHANGE_RATIO_THRESHOLD)
        noise_threshold = params.get(
            "noise_threshold", config.SCREENSPACE_NOISE_THRESHOLD
        )
        mag = compute_frame_diff_gray(
            _cached_blur_gray(prev_cache, prev_frame, region),
            _cached_blur_gray(cache, frame, region),
            noise_threshold,
            mask=_cached_mask(cache, frame, region),
        )
        return mag >= threshold, {"magnitude": round(mag, 4)}


class SimilarityTool(AnalysisTool):
    # Scalar SSIM only; a spatial heatmap needs per-frame ssim_diff_map, so gate it behind phash.
    name = "similarity"
    fast_scan_region_dim = 128
    score_key = "score"
    scan_fn_name = "scan_similarity"
    scan_defaults: ClassVar[dict[str, Any]] = {
        "reference_frame": None,
        "threshold": 0,
    }

    def check_frame(
        self, frame, prev_frame, region, params, cache=None, prev_cache=None
    ):
        ref = params.get("reference_frame")
        if ref is None:
            return False, None
        pixels = _cached_crop(cache, frame, region)
        threshold = params.get("threshold", config.SCREENSPACE_SSIM_THRESHOLD)
        is_sim, score = regions_are_similar(pixels, ref, threshold)
        return is_sim, {"score": round(score, 4)}

    def scan(
        self,
        video_path,
        region,
        params,
        *,
        task_id,
        scan_mode,
        on_progress,
        cancel_flag,
        on_result,
        fast_opts,
    ):
        # `is None`, not falsy: the reference is an ndarray with ambiguous truthiness.
        if params.get("reference_frame") is None:
            raise ValueError("Similarity scan requires a reference_frame parameter")
        return super().scan(
            video_path,
            region,
            params,
            task_id=task_id,
            scan_mode=scan_mode,
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class TextTool(AnalysisTool):
    name = "text"
    # Fuzzy match quality, a different axis from _extract_confidence's OCR "confidence"; keep them apart.
    score_key = "fuzzy_ratio"
    scan_fn_name = "scan_text"
    scan_interval_default = 2.0
    scan_defaults: ClassVar[dict[str, Any]] = {
        "search_string": "",
        "fuzzy_threshold": 0,
        "ocr_confidence_threshold": None,
        "ocr_preprocess": False,
        "require_consecutive": 1,
        "languages": None,
    }

    def _scan_kwargs(self, params):
        kwargs = super()._scan_kwargs(params)
        # or-expression, not a get-default: an explicit "" must coerce to "off".
        kwargs["ocr_normalize"] = params.get("ocr_normalize") or "off"
        return kwargs

    def check_frame(
        self, frame, prev_frame, region, params, cache=None, prev_cache=None
    ):
        if not params.get("search_string", ""):
            return False, None
        readings = _cached_ocr(
            cache,
            frame,
            region,
            languages=params.get("languages") or ["en"],
            preprocess=params.get("ocr_preprocess", False),
        )
        return _score_text_readings(readings, params)


class NumbersTool(AnalysisTool):
    name = "numbers"
    score_key = "confidence"
    scan_fn_name = "scan_numbers"
    scan_interval_default = 2.0
    scan_defaults: ClassVar[dict[str, Any]] = {
        "operator": "gt",
        "target_value": 0,
        "range_min": None,
        "range_max": None,
        "ocr_confidence_threshold": None,
        "ocr_preprocess": False,
        "integers_only": False,
        "require_consecutive": 1,
        "languages": None,
    }

    def check_frame(
        self, frame, prev_frame, region, params, cache=None, prev_cache=None
    ):
        readings = _cached_ocr(
            cache,
            frame,
            region,
            languages=params.get("languages") or ["en"],
            preprocess=params.get("ocr_preprocess", False),
        )
        return _score_numbers_readings(readings, params)


class TemplateTool(AnalysisTool):
    name = "template"
    fast_scan_extra_opts: ClassVar[dict[str, Any]] = {"template_downscale": True}
    score_key = "best_score"

    def check_frame(
        self, frame, prev_frame, region, params, cache=None, prev_cache=None
    ):
        # Nothing region-scoped to memoize: template searches the whole frame and caches its prep on params.
        template_img = params.get("template_image")
        if template_img is None:
            return False, None
        threshold = params.get("threshold", config.SCREENSPACE_TEMPLATE_MATCH_THRESHOLD)
        # Cache the scaled template prep on params so multitool amortizes it; template_scale mirrors scan_template.
        cached = params.get("_prepared_template")
        if cached is None:
            scaled_img, scaled_mask = _scale_template(
                template_img,
                params.get("template_mask"),
                float(params.get("template_scale", 1.0)),
            )
            cached = (
                scaled_img,
                scaled_mask,
                _prepare_template(scaled_img, scaled_mask),
            )
            params["_prepared_template"] = cached
        scaled_img, scaled_mask, prepared = cached
        # Peak correlation is threshold-independent and available on a miss, unlike match_template.
        corr = _template_correlation_map(frame, prepared)
        if corr is None:
            return False, None
        # The run region scopes search and peak (zero-size = anywhere), so calibration scores the target.
        window = region_search_window(region)
        if window is not None:
            tpl_h, tpl_w = prepared[0].shape[:2]
            corr = _mask_corr_outside_window(corr, tpl_w, tpl_h, window)
            if corr is None:
                return False, {"best_score": -1.0, "match_count": 0}
        peak = float(corr.max())
        if peak < threshold:
            return False, {"best_score": round(peak, 4), "match_count": 0}
        matches = match_template(
            frame,
            scaled_img,
            threshold=threshold,
            mask=scaled_mask,
            prepared=prepared,
            corr=corr,
        )
        # Region polygon filters detections; passing needs a surviving match. best_score stays the window peak.
        matches = filter_matches_by_region_mask(matches, region)
        if region.get("mask_points") and not matches:
            return False, {"best_score": round(peak, 4), "match_count": 0}
        return True, {"best_score": round(peak, 4), "match_count": len(matches)}

    def scan(
        self,
        video_path,
        region,
        params,
        *,
        task_id,
        scan_mode,
        on_progress,
        cancel_flag,
        on_result,
        fast_opts,
    ):
        template_img = params.get("template_image")
        if template_img is None:
            raise ValueError("Template scan requires a template_image parameter")
        tmpl_mask = params.get("template_mask")
        # Fast scan halves template + mask; the template_downscale fast_opts flag halves frames.
        if scan_mode == "fast":
            th, tw = template_img.shape[:2]
            ntw, nth = tw // 2, th // 2
            if ntw > 0 and nth > 0:
                template_img = cv2.resize(
                    template_img, (ntw, nth), interpolation=cv2.INTER_AREA
                )
                if tmpl_mask is not None:
                    tmpl_mask = cv2.resize(
                        tmpl_mask, (ntw, nth), interpolation=cv2.INTER_AREA
                    )
        return scan_template(
            video_path,
            region,
            template_image=template_img,
            threshold=params.get("threshold", 0),
            interval_seconds=params.get("interval", 0),
            template_mask=tmpl_mask,
            template_scale=float(params.get("template_scale", 1.0)),
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class ShapeTool(AnalysisTool):
    name = "shape"
    fast_scan_extra_opts: ClassVar[dict[str, Any]] = {"template_downscale": True}
    score_key = "best_score"

    def check_frame(
        self, frame, prev_frame, region, params, cache=None, prev_cache=None
    ):
        # Shape ignores region; the per-scale edge prep is task-constant, so cache it on params.
        shape_img = params.get("shape_image")
        if shape_img is None:
            return False, None
        threshold = params.get("threshold", config.SCREENSPACE_SHAPE_MATCH_THRESHOLD)
        prepared = params.get("_prepared_shape")
        if prepared is None:
            prepared = _prepare_shape_reference(
                shape_img,
                params.get("shape_mask"),
                float(params.get("scale_min", 0.0)),
                float(params.get("scale_max", 0.0)),
                int(params.get("scale_steps", 0)),
                float(params.get("scale_y_min", 0.0)),
                float(params.get("scale_y_max", 0.0)),
                int(params.get("scale_y_steps", 0)),
            )
            params["_prepared_shape"] = prepared
        if not prepared:
            # Degenerate reference: no scale kept enough edge pixels.
            return False, None
        # Cross-scale peak: threshold-independent, available on a miss. The run region scopes search and peak.
        window = region_search_window(region)
        matches, peak = match_shape(
            _frame_edge_map(frame), prepared, threshold, window=window
        )
        if not matches:
            return False, {"best_score": round(peak, 4), "match_count": 0}
        # Shaped region: the polygon refines the window as a detection filter.
        matches = filter_matches_by_region_mask(matches, region)
        if region.get("mask_points") and not matches:
            return False, {"best_score": round(peak, 4), "match_count": 0}
        return True, {"best_score": round(peak, 4), "match_count": len(matches)}

    def scan(
        self,
        video_path,
        region,
        params,
        *,
        task_id,
        scan_mode,
        on_progress,
        cancel_flag,
        on_result,
        fast_opts,
    ):
        shape_img = params.get("shape_image")
        if shape_img is None:
            raise ValueError("Shape scan requires a shape_image parameter")
        shape_mask = params.get("shape_mask")
        # Fast scan halves reference + mask like TemplateTool; template_downscale halves frames.
        if scan_mode == "fast":
            sh, sw = shape_img.shape[:2]
            nsw, nsh = sw // 2, sh // 2
            if nsw > 0 and nsh > 0:
                shape_img = cv2.resize(
                    shape_img, (nsw, nsh), interpolation=cv2.INTER_AREA
                )
                if shape_mask is not None:
                    shape_mask = cv2.resize(
                        shape_mask, (nsw, nsh), interpolation=cv2.INTER_AREA
                    )
        return scan_shape(
            video_path,
            region,
            shape_image=shape_img,
            threshold=params.get("threshold", 0),
            interval_seconds=params.get("interval", 0),
            shape_mask=shape_mask,
            scale_min=float(params.get("scale_min", 0.0)),
            scale_max=float(params.get("scale_max", 0.0)),
            scale_steps=int(params.get("scale_steps", 0)),
            scale_y_min=float(params.get("scale_y_min", 0.0)),
            scale_y_max=float(params.get("scale_y_max", 0.0)),
            scale_y_steps=int(params.get("scale_y_steps", 0)),
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class FlowTool(AnalysisTool):
    name = "flow"
    fast_scan_region_dim = 128
    score_key = "magnitude"
    scan_fn_name = "scan_flow"
    scan_defaults: ClassVar[dict[str, Any]] = {
        "magnitude_threshold": 0,
        "require_consecutive": 1,
    }

    def check_frame(
        self, frame, prev_frame, region, params, cache=None, prev_cache=None
    ):
        if prev_frame is None:
            return False, None
        curr_gray = _cached_gray(cache, frame, region)
        prev_gray_f = _cached_gray(prev_cache, prev_frame, region)
        magnitude_threshold = params.get(
            "magnitude_threshold", config.SCREENSPACE_FLOW_MAGNITUDE_THRESHOLD
        )
        flow_result = compute_optical_flow(
            prev_gray_f, curr_gray, mask=_cached_mask(cache, frame, region)
        )
        return flow_result["magnitude"] >= magnitude_threshold, {
            "magnitude": flow_result["magnitude"],
            "angle": flow_result["angle"],
        }


class SceneTool(AnalysisTool):
    name = "scene"
    fast_scan_region_dim = 64
    score_key = "score"
    scan_fn_name = "scan_scene"
    scan_defaults: ClassVar[dict[str, Any]] = {
        "reference_scenes": None,
        "threshold": 0,
    }

    def check_frame(
        self, frame, prev_frame, region, params, cache=None, prev_cache=None
    ):
        ref_scenes = params.get("reference_scenes")
        if not ref_scenes:
            return False, None
        threshold = params.get(
            "threshold", config.SCREENSPACE_SCENE_SIMILARITY_THRESHOLD
        )
        fp = _memo(
            cache,
            _region_key("fingerprint", region),
            lambda: compute_scene_fingerprint(
                _cached_crop(cache, frame, region),
                mask=_cached_mask(cache, frame, region),
            ),
        )
        # Ref dicts are shared across steps with different regions; fingerprints only compare under one mask.
        mask_key = mask_points_key(region.get("mask_points"))
        best_name = ""
        best_score = 0.0
        for ref in ref_scenes:
            # Cache fingerprint on the reference dict to avoid recomputing per frame
            ref_fp = ref.get("_cached_fingerprint")
            if ref_fp is None or ref.get("_cached_fingerprint_mask") != mask_key:
                ref_fp = compute_scene_fingerprint(
                    ref["frame"],
                    mask=region_mask_for(region, *ref["frame"].shape[:2]),
                )
                ref["_cached_fingerprint"] = ref_fp
                ref["_cached_fingerprint_mask"] = mask_key
            score = compare_scene_fingerprints(fp, ref_fp)
            if score > best_score:
                best_score = score
                best_name = ref["name"]
        return best_score >= threshold, {
            "scene_name": best_name,
            "score": round(best_score, 4),
        }

    def scan(
        self,
        video_path,
        region,
        params,
        *,
        task_id,
        scan_mode,
        on_progress,
        cancel_flag,
        on_result,
        fast_opts,
    ):
        # Falsy on purpose: an empty list is as unusable as None. Contrast Similarity's `is None`.
        if not params.get("reference_scenes"):
            raise ValueError("Scene scan requires reference_scenes parameter")
        return super().scan(
            video_path,
            region,
            params,
            task_id=task_id,
            scan_mode=scan_mode,
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class InactivityTool(AnalysisTool):
    name = "inactivity"
    fast_scan_region_dim = 64
    # Raw phash distance in Sensitivity-slider units; the strip inverts it (lower = more inactive).
    score_key = "distance"
    scan_fn_name = "scan_inactivity"
    scan_defaults: ClassVar[dict[str, Any]] = {
        "threshold": 0,
        "min_duration": 0.0,
    }

    def check_frame(
        self, frame, prev_frame, region, params, cache=None, prev_cache=None
    ):
        if prev_frame is None:
            return False, None
        thresh = params.get("threshold", config.SCREENSPACE_INACTIVITY_PHASH_THRESHOLD)
        curr_h = _cached_phash(cache, frame, region)
        prev_h = _cached_phash(prev_cache, prev_frame, region)
        dist = int(curr_h - prev_h)
        if thresh > 0:
            conf = max(0.0, min((thresh - dist) / float(thresh), 1.0))
        else:
            conf = 1.0 if dist <= thresh else 0.0
        return dist <= thresh, {"distance": dist, "_confidence": conf}


class BoundaryTool(AnalysisTool):
    name = "boundary"
    # The scanner runs its own phash per sample; the fast-scan phash-skip would fight it.
    supports_fast_scan = False
    # Scan-only: no check_frame or score_key, so the pinned strip and /api/calibrate skip it.
    scan_fn_name = "scan_boundaries"
    scan_defaults: ClassVar[dict[str, Any]] = {
        "threshold": 0,
        "min_gap": 0.0,
    }

    def _scan_kwargs(self, params):
        kwargs = super()._scan_kwargs(params)
        # Policy default lives here; the primitive keeps "phash" for direct callers and tests.
        kwargs["metric"] = params.get("metric") or config.SCREENSPACE_BOUNDARY_METRIC
        return kwargs


class AttentionTool(AnalysisTool):
    name = "attention"
    # Dwell weighting needs every sampled frame; the fast-scan phash-skip would drop the static ones.
    supports_fast_scan = False
    # Scan-only (full-frame, temporal state): no check_frame or score_key, so calibration skips it.
    scan_fn_name = "scan_attention"
    scan_defaults: ClassVar[dict[str, Any]] = {
        "shift_threshold": 0.0,
        "ema_alpha": 0.0,
    }

    def _scan_kwargs(self, params):
        kwargs = super()._scan_kwargs(params)
        # Channel weights, center bias, face toggle; absent keys use SCREENSPACE_ATTENTION_* defaults.
        kwargs.update(saliency_kwargs_from_params(params))
        return kwargs


class TimelapseTool(AnalysisTool):
    name = "timelapse"
    # Own sample_interval and a media-file output; fast scan and Viewer detector entries don't apply.
    supports_fast_scan = False

    def scan(
        self,
        video_path,
        region,
        params,
        *,
        task_id,
        scan_mode,
        on_progress,
        cancel_flag,
        on_result,
        fast_opts,
    ):
        output_path = params.get("output_path", "")
        if not output_path:
            ext = "gif" if params.get("output_format") == "gif" else "mp4"
            output_path = str(
                Path(utils.get_effective_output_dir()) / f"timelapse_{task_id}.{ext}"
            )
        return generate_timelapse(
            video_path,
            region,
            speedup_factor=params.get("speedup_factor", 10.0),
            output_path=output_path,
            output_format=params.get("output_format", "mp4"),
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            sample_interval=params.get("sample_interval", 0.0),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
        )


class MultitoolTool(AnalysisTool):
    name = "multitool"

    def scan(
        self,
        video_path,
        region,
        params,
        *,
        task_id,
        scan_mode,
        on_progress,
        cancel_flag,
        on_result,
        fast_opts,
    ):
        # Local import breaks the tools<->multitool cycle; runs at task time, not import time.
        from screenspace_multitool import scan_multitool

        steps = [dict(s) for s in params.get("steps", [])]
        if len(steps) < 2:
            raise ValueError("Multitool requires at least 2 steps")
        # Fast scan: multiply the interval used by scan_multitool
        # (reads from steps[0]["interval"] with fallback to default).
        if scan_mode == "fast" and steps:
            mt_interval = steps[0].get("interval", config.SCREENSPACE_DEFAULT_INTERVAL)
            steps[0]["interval"] = (
                mt_interval * config.SCREENSPACE_FAST_SCAN_INTERVAL_MULTIPLIER
            )
        return scan_multitool(
            video_path,
            region,
            steps=steps,
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


TOOLS: dict[str, AnalysisTool] = {
    t.name: t
    for t in (
        ColorTool(),
        ChangeTool(),
        SimilarityTool(),
        TextTool(),
        NumbersTool(),
        TemplateTool(),
        ShapeTool(),
        FlowTool(),
        SceneTool(),
        InactivityTool(),
        BoundaryTool(),
        AttentionTool(),
        TimelapseTool(),
        MultitoolTool(),
    )
}

# Keeps the by-name scan_<x> imports referenced; a typo'd scan_fn_name fails at import, not mid-scan.
_DISPATCHABLE_SCAN_FNS = {
    scan_attention,
    scan_boundaries,
    scan_changes,
    scan_color,
    scan_flow,
    scan_inactivity,
    scan_numbers,
    scan_scene,
    scan_similarity,
    scan_text,
}
for _tool in TOOLS.values():
    if (
        _tool.scan_fn_name
        and globals()[_tool.scan_fn_name] not in _DISPATCHABLE_SCAN_FNS
    ):
        raise AssertionError(
            f"{_tool.name}: unknown scan_fn_name {_tool.scan_fn_name!r}"
        )
