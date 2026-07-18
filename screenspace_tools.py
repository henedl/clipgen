# -*- coding: utf-8 -*-
"""Screenspace analysis tools (strategy registry) + per-frame dispatch.

The ``AnalysisTool`` base class, one subclass per tool, the ``TOOLS`` registry,
and the single-frame evaluation/scoring entry points used by the multitool
chainer and pin calibration. Imports primitives, OCR helpers, and scan
workflows from sibling modules; the one tools->multitool edge
(``MultitoolTool.scan`` -> ``scan_multitool``) is a function-local import to
keep the module graph acyclic.
"""

import math
from pathlib import Path
from typing import Any, Callable, ClassVar

import cv2
import numpy as np

import config
import utils
from screenspace_primitives import (
    _prepare_template,
    _scale_template,
    _template_correlation_map,
    color_matches,
    color_present,
    compare_scene_fingerprints,
    compute_frame_diff,
    compute_optical_flow,
    compute_phash,
    compute_scene_fingerprint,
    extract_region,
    filter_matches_by_region_mask,
    mask_points_key,
    match_template,
    region_mask_for,
    regions_are_similar,
)
from screenspace_ocr import (
    _numbers_ocr_allowlist,
    _ocr_region_readings,
    _score_numbers_readings,
    _score_text_readings,
)
from screenspace_scans import (
    generate_timelapse,
    scan_boundaries,
    scan_changes,
    scan_color,
    scan_flow,
    scan_inactivity,
    scan_numbers,
    scan_scene,
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
    elif tool_type == "template":
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
    return 1.0


# ---------------------------------------------------------------------------
# Per-frame memoization (multitool chains)
# ---------------------------------------------------------------------------
#
# Within a multitool chain every step re-derives crop / gray / phash / OCR from
# the same ``(frame, region)``. ``scan_multitool`` passes a fresh dict per frame
# as ``cache`` (and the previous frame's dict as ``prev_cache``, rolled forward)
# so steps sharing a region compute each derived value once, and temporal tools
# (change / flow / inactivity) reuse the previous frame's already-computed crop /
# gray instead of recomputing it. Single-frame callers (pin calibration) pass no
# cache — ``None`` disables memoization and keeps results byte-identical.

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
    return _memo(
        cache, _region_key("crop", region), lambda: extract_region(frame, region)
    )


def _cached_mask(
    cache: dict[Any, Any] | None, frame: np.ndarray, region: dict[str, Any]
) -> np.ndarray | None:
    """Shaped-region mask at the cached crop's size, memoized on *cache*.

    ``None`` for rect regions, so masked tools can pass the value straight to
    the primitives' optional ``mask`` params.
    """
    return _memo(
        cache,
        _region_key("mask", region),
        lambda: region_mask_for(region, *_cached_crop(cache, frame, region).shape[:2]),
    )


def _cached_gray(
    cache: dict[Any, Any] | None, frame: np.ndarray, region: dict[str, int]
) -> np.ndarray:
    """Grayscale of the region crop, memoized on *cache* (reuses the cached crop)."""
    return _memo(
        cache,
        _region_key("gray", region),
        lambda: cv2.cvtColor(_cached_crop(cache, frame, region), cv2.COLOR_BGR2GRAY),
    )


def _cached_phash(
    cache: dict[Any, Any] | None, frame: np.ndarray, region: dict[str, int]
) -> "Any":
    """Perceptual hash of the region crop, memoized on *cache* (reuses the crop)."""
    return _memo(
        cache,
        _region_key("phash", region),
        lambda: compute_phash(_cached_crop(cache, frame, region)),
    )


def _cached_ocr(
    cache: dict[Any, Any] | None,
    frame: np.ndarray,
    region: dict[str, Any],
    *,
    languages: list[str],
    allowlist: str | None,
    preprocess: bool,
) -> list[Any]:
    """OCR readings for the region crop, memoized on *cache*.

    Keyed by ``(region, languages, allowlist, preprocess)`` so only *identical*
    OCR calls dedup (e.g. two Text steps on one region). Text (no allowlist) and
    Numbers (digit allowlist) key differently and are intentionally not shared.
    """
    langs = tuple(languages)
    key = _region_key("ocr", region) + (langs, allowlist, preprocess)
    return _memo(
        cache,
        key,
        lambda: _ocr_region_readings(
            _cached_crop(cache, frame, region),
            languages=list(langs),
            allowlist=allowlist,
            preprocess=preprocess,
            mask_points=region.get("mask_points"),
        ),
    )


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

    Used by :func:`scan_multitool` for steps 1+ in the chain.  Returns
    ``(passed, result_dict)`` where *result_dict* contains tool-specific
    metadata when the check passes, or ``None`` when it does not.

    For **change** and **flow** tools *prev_frame* is required (the frame
    immediately before the candidate timestamp).  If it is ``None`` the
    check is skipped (returns ``(False, None)``).

    ``cache`` / ``prev_cache`` are optional per-frame memo dicts (this frame's and
    the previous frame's) that let chained steps sharing a region reuse crop /
    gray / phash / OCR work; ``None`` (the calibration path) disables memoization.

    Degenerate regions (width or height ≤ 0, e.g. a 1-px user draw rounded
    to zero pixels on a small preview) are treated as a non-match here so
    downstream cv2 ops never see empty arrays.  Template is exempt: it matches
    against the full frame and ignores ``region``, so an uploaded template step
    with no region (zero-size region_coords) is valid — mirrors
    :func:`score_frame_for_tool`.
    """
    if tool_type != "template" and (region.get("w", 0) <= 0 or region.get("h", 0) <= 0):
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

    ``ocr_reader`` (text/numbers only) supplies cached EasyOCR readings keyed per
    pin so fuzzy/confidence changes re-score without re-running OCR; when absent
    the tool runs OCR live through its ``check_frame``.
    """
    # Template matches against the full frame and ignores ``region``, so a
    # zero-size region — the case when an uploaded template scans the whole
    # frame with no region_ref — is valid for it. Every other tool crops the
    # region, where an empty crop would break downstream cv2 ops.
    if tool_type != "template" and (region.get("w", 0) <= 0 or region.get("h", 0) <= 0):
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
    # Max region dimension when running in "fast" scan mode (passed to the
    # generic frame extractor as ``max_region_dim``). 0 means no downscale.
    fast_scan_region_dim: ClassVar[int] = 0
    # Whether the tool participates in the fast-scan optimization at all.
    # Timelapse opts out because it has its own ``sample_interval``.
    supports_fast_scan: ClassVar[bool] = True
    # Extra keys merged into ``fast_opts`` (e.g. ``{"template_downscale": True}``).
    fast_scan_extra_opts: ClassVar[dict[str, Any]] = {}
    # Detail-dict key holding the threshold-independent calibration scalar that
    # ``check_frame`` populates on both branches. Empty ⇒ tool not calibratable.
    score_key: ClassVar[str] = ""

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
        """Run a full-video scan. Subclasses must override."""
        raise NotImplementedError


class ColorTool(AnalysisTool):
    name = "color"
    fast_scan_region_dim = 32
    score_key = "_confidence"

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
        color_mode = params.get("color_mode", "average")
        # Presence scans must stay full-resolution: the fast-scan max_region_dim
        # downscale uses INTER_AREA averaging, which erases small color patches.
        if color_mode == "presence":
            fast_opts = None
        return scan_color(
            video_path,
            region,
            target_color=params.get("target_color", {"h": 0, "s": 0, "v": 0}),
            tolerance=params.get("tolerance", {"h": 10, "s": 50, "v": 50}),
            interval_seconds=params.get("interval", 0),
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            color_mode=color_mode,
            min_coverage=params.get("min_coverage", 0.0),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class ChangeTool(AnalysisTool):
    name = "change"
    fast_scan_region_dim = 128
    score_key = "magnitude"

    def check_frame(
        self, frame, prev_frame, region, params, cache=None, prev_cache=None
    ):
        if prev_frame is None:
            return False, None
        pixels = _cached_crop(cache, frame, region)
        prev_pixels = _cached_crop(prev_cache, prev_frame, region)
        threshold = params.get("threshold", config.SCREENSPACE_CHANGE_RATIO_THRESHOLD)
        noise_threshold = params.get(
            "noise_threshold", config.SCREENSPACE_NOISE_THRESHOLD
        )
        mag = compute_frame_diff(
            prev_pixels,
            pixels,
            noise_threshold,
            mask=_cached_mask(cache, frame, region),
        )
        return mag >= threshold, {"magnitude": round(mag, 4)}

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
        return scan_changes(
            video_path,
            region,
            threshold=params.get("threshold", 0),
            interval_seconds=params.get("interval", 0),
            noise_threshold=params.get("noise_threshold", 0),
            require_consecutive=params.get("require_consecutive", 1),
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class SimilarityTool(AnalysisTool):
    # A spatial Similarity heatmap is feasible but deferred: this tool computes
    # only a scalar SSIM via ``regions_are_similar``. An accumulated heatmap would
    # need the per-pixel map (``ssim_diff_map`` /
    # ``structural_similarity(..., full=True)``, already used by the Model view
    # preview off the hot path), which adds per-frame CPU/memory; if added later,
    # gate it behind the phash pre-filter so the full map is only computed on
    # candidate frames. (See Change/Template heatmaps in screenspace_heatmap.py
    # for the accumulation pattern.)
    name = "similarity"
    fast_scan_region_dim = 128
    score_key = "score"

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
        ref_frame = params.get("reference_frame")
        if ref_frame is None:
            raise ValueError("Similarity scan requires a reference_frame parameter")
        return scan_similarity(
            video_path,
            region,
            reference_frame=ref_frame,
            threshold=params.get("threshold", 0),
            interval_seconds=params.get("interval", 0),
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class TextTool(AnalysisTool):
    name = "text"
    # Calibration scalar is fuzzy match quality. Deliberately distinct from
    # ``_extract_confidence``'s "confidence" key (OCR reading confidence) — two
    # different axes; do not unify them.
    score_key = "fuzzy_ratio"

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
            allowlist=None,
            preprocess=params.get("ocr_preprocess", False),
        )
        return _score_text_readings(readings, params)

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
        return scan_text(
            video_path,
            region,
            search_string=params.get("search_string", ""),
            interval_seconds=params.get("interval", 2.0),
            fuzzy_threshold=params.get("fuzzy_threshold", 0),
            ocr_confidence_threshold=params.get("ocr_confidence_threshold"),
            ocr_preprocess=params.get("ocr_preprocess", False),
            ocr_normalize=params.get("ocr_normalize") or "off",
            require_consecutive=params.get("require_consecutive", 1),
            languages=params.get("languages"),
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class NumbersTool(AnalysisTool):
    name = "numbers"
    score_key = "confidence"

    def check_frame(
        self, frame, prev_frame, region, params, cache=None, prev_cache=None
    ):
        languages = params.get("languages") or ["en"]
        readings = _cached_ocr(
            cache,
            frame,
            region,
            languages=languages,
            allowlist=_numbers_ocr_allowlist(
                languages, params.get("integers_only", False)
            ),
            preprocess=params.get("ocr_preprocess", False),
        )
        return _score_numbers_readings(readings, params)

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
        return scan_numbers(
            video_path,
            region,
            operator=params.get("operator", "gt"),
            target_value=params.get("target_value", 0),
            interval_seconds=params.get("interval", 2.0),
            range_min=params.get("range_min"),
            range_max=params.get("range_max"),
            ocr_confidence_threshold=params.get("ocr_confidence_threshold"),
            ocr_preprocess=params.get("ocr_preprocess", False),
            integers_only=params.get("integers_only", False),
            require_consecutive=params.get("require_consecutive", 1),
            languages=params.get("languages"),
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class TemplateTool(AnalysisTool):
    name = "template"
    fast_scan_extra_opts: ClassVar[dict[str, Any]] = {"template_downscale": True}
    score_key = "best_score"

    def check_frame(
        self, frame, prev_frame, region, params, cache=None, prev_cache=None
    ):
        # Template matches the full frame (ignores region); nothing region-scoped
        # to memoize, and it already caches its scaled template on ``params``.
        template_img = params.get("template_image")
        if template_img is None:
            return False, None
        threshold = params.get("threshold", config.SCREENSPACE_TEMPLATE_MATCH_THRESHOLD)
        # Cache the per-task-constant scaled template/mask + grayscale prep on the
        # parameters dict so multitool scans amortize it across frames. The
        # template_scale slider resizes the (often uploaded) template to its
        # in-video pixel size before matching, mirroring scan_template.
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
        # Peak correlation is the threshold-independent scalar; available even on
        # a miss (unlike match_template, which only returns above-threshold hits).
        corr = _template_correlation_map(frame, prepared)
        if corr is None:
            return False, None
        peak = float(corr.max())
        if peak < threshold:
            return False, {"best_score": round(peak, 4), "match_count": 0}
        matches = match_template(
            frame,
            scaled_img,
            threshold=threshold,
            mask=scaled_mask,
            prepared=prepared,
        )
        # Shaped region: the match still runs full-frame (template ignores the
        # rect for rect regions too), but the polygon acts as a detection
        # filter — passing requires at least one surviving match. best_score
        # stays the frame peak (the threshold-independent calibration scalar).
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
        # Fast scan: downscale template + mask by 2x before passing to the
        # scan function (which separately downscales the frame via the
        # ``template_downscale`` fast_opts flag).
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


class FlowTool(AnalysisTool):
    name = "flow"
    fast_scan_region_dim = 128
    score_key = "magnitude"

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
        return scan_flow(
            video_path,
            region,
            magnitude_threshold=params.get("magnitude_threshold", 0),
            interval_seconds=params.get("interval", 0),
            require_consecutive=params.get("require_consecutive", 1),
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class SceneTool(AnalysisTool):
    name = "scene"
    fast_scan_region_dim = 64
    score_key = "score"

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
        # Fingerprints are only comparable under the same mask, so the per-ref
        # cache is keyed by the region's polygon (refs are cached on the ref
        # dict, which multitool steps with different regions could share).
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
        ref_scenes = params.get("reference_scenes")
        if not ref_scenes:
            raise ValueError("Scene scan requires reference_scenes parameter")
        return scan_scene(
            video_path,
            region,
            reference_scenes=ref_scenes,
            threshold=params.get("threshold", 0),
            interval_seconds=params.get("interval", 0),
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class InactivityTool(AnalysisTool):
    name = "inactivity"
    fast_scan_region_dim = 64
    # Calibration scalar is the raw phash distance (Sensitivity-slider units);
    # the strip inverts the axis for display (lower distance = more inactive).
    score_key = "distance"

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
        return scan_inactivity(
            video_path,
            region,
            threshold=params.get("threshold", 0),
            min_duration=params.get("min_duration", 0.0),
            interval_seconds=params.get("interval", 0),
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class BoundaryTool(AnalysisTool):
    name = "boundary"
    # The scanner is already coarse and runs its own phash on every sample;
    # the generic fast-scan phash-skip would fight that logic, so opt out.
    supports_fast_scan = False
    # Scan-only in v1: not a multitool step and not calibratable (no score_key),
    # so the pinned-frame strip and /api/calibrate correctly skip it. Wiring
    # calibration later means adding score_key + a per-frame check_frame here,
    # plus a `needs_prev` entry and full-frame handling in the calibrate endpoint
    # (boundary needs the previous sampled frame and always scans the full frame).

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
        return scan_boundaries(
            video_path,
            region,
            threshold=params.get("threshold", 0),
            min_gap=params.get("min_gap", 0.0),
            interval_seconds=params.get("interval", 0),
            # Policy default lives here, not in the scan primitive: a task with
            # no metric (UI "Auto") gets the configured default; the primitive's
            # own default stays "phash" for direct callers/tests.
            metric=params.get("metric") or config.SCREENSPACE_BOUNDARY_METRIC,
            start_seconds=params.get("start_seconds", 0.0),
            end_seconds=params.get("end_seconds"),
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


class TimelapseTool(AnalysisTool):
    name = "timelapse"
    # Has its own ``sample_interval`` and produces a media file rather than
    # per-frame events, so the generic fast-scan path does not apply.
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
        # Function-local import breaks the tools<->multitool cycle: scan_multitool
        # needs check_frame_for_tool/score_frame_for_tool from this module, while
        # only this one method needs scan_multitool. Imported at task-execution
        # time, never at module import, so there is no hot-path cost.
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
        FlowTool(),
        SceneTool(),
        InactivityTool(),
        BoundaryTool(),
        TimelapseTool(),
        MultitoolTool(),
    )
}
