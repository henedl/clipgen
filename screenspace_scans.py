# -*- coding: utf-8 -*-
"""Screenspace scan workflows (one per analysis tool).

Each scan sweeps a video via the frame-extraction drivers and applies one
analysis primitive (color, change, similarity, text, numbers, timelapse,
template, flow, scene, inactivity, boundary, attention). Scans never call
each other.
Imports primitives, OCR helpers, and frame extractors from sibling modules.
"""

import subprocess
from typing import TYPE_CHECKING, Any, Callable

import cv2
import numpy as np

if TYPE_CHECKING:
    import imagehash

import config
import utils
from screenspace_primitives import (
    _ConsecutiveBuffer,
    _frame_diff_mask,
    _frame_is_static,
    _is_static_skip,
    _match_template_prepared,
    _merge_timestamp_spans,
    _prepare_template,
    _scale_template,
    color_matches,
    color_present,
    compare_scene_fingerprints,
    compute_optical_flow,
    compute_phash,
    compute_saliency_map,
    compute_scene_fingerprint,
    filter_matches_by_region_mask,
    region_masker,
    saliency_grid_from_map,
    saliency_peak,
)
from screenspace_ocr import (
    _VALID_OPERATORS,
    _effective_ocr_confidence_threshold,
    _numbers_ocr_allowlist,
    _ocr_region_readings,
    _score_numbers_readings,
    _score_text_readings,
)
from screenspace_frames import (
    _probe_video_meta,
    _resolve_scan_window,
    build_timelapse_command,
    scan_video_frames,
    scan_video_full_frames,
)


def scan_color(
    video_path: str,
    region: dict[str, int],
    target_color: dict[str, float],
    tolerance: dict[str, float],
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    color_mode: str = "average",
    min_coverage: float = 0.0,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan video for frames where the region matches a target color.

    With ``color_mode="average"`` (default) a frame matches when the region's
    *average* color is within tolerance. With ``color_mode="presence"`` a frame
    matches when the target color appears anywhere in the region (per-pixel),
    covering at least ``min_coverage`` of it.

    Returns list of ``{start, end, duration}`` spans (consecutive matches
    merged).
    """
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    window = _resolve_scan_window(video_path, start_seconds, end_seconds)
    if window is None:
        return []
    vid_fps, vid_duration, end_seconds, total_range = window

    matches: list[float] = []
    mask_for = region_masker(region)

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        mask = mask_for(pixels)
        if color_mode == "presence":
            matched, conf = color_present(
                pixels, target_color, tolerance, min_coverage, mask=mask
            )
        else:
            matched, conf = color_matches(pixels, target_color, tolerance, mask=mask)
        if matched:
            matches.append(ts)
            if on_result:
                on_result({"timestamp": ts, "_confidence": conf})
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_frames(
        video_path,
        region,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return _merge_timestamp_spans(matches, interval_seconds)


def scan_changes(
    video_path: str,
    region: dict[str, int],
    threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    noise_threshold: int = 0,
    require_consecutive: int = 1,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan video for change points in a region.

    Returns list of ``{timestamp, magnitude}`` dicts.
    """
    if threshold <= 0:
        threshold = config.SCREENSPACE_CHANGE_RATIO_THRESHOLD
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL
    if noise_threshold <= 0:
        noise_threshold = config.SCREENSPACE_NOISE_THRESHOLD

    window = _resolve_scan_window(video_path, start_seconds, end_seconds)
    if window is None:
        return []
    vid_fps, vid_duration, end_seconds, total_range = window

    results: list[dict[str, Any]] = []
    prev_pixels: list[np.ndarray | None] = [None]
    buf = _ConsecutiveBuffer(require_consecutive)
    # change_grid feeds only the Change heatmap; skip the per-frame downsample
    # entirely when heatmaps are disabled (the data would just be discarded).
    build_grid = config.SCREENSPACE_GENERATE_CHANGE_HEATMAP
    grid = config.SCREENSPACE_CHANGE_HEATMAP_GRID
    min_frac = config.SCREENSPACE_CHANGE_HEATMAP_MIN_FRAC

    mask_for = region_masker(region)

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        if prev_pixels[0] is not None:
            mask = _frame_diff_mask(prev_pixels[0], pixels, noise_threshold)
            region_mask = mask_for(pixels)
            if region_mask is not None:
                # Shaped region: only changes inside the polygon count, and the
                # magnitude is relative to the polygon's area. The ANDed mask
                # also feeds change_grid below, so heatmap cells outside the
                # polygon are suppressed for free.
                mask = cv2.bitwise_and(mask, region_mask)
                denom = float(np.count_nonzero(region_mask))
            else:
                denom = float(mask.size)
            mag = float(np.count_nonzero(mask)) / denom if denom else 0.0
            if mag >= threshold:
                rd: dict[str, Any] = {"timestamp": ts, "magnitude": round(mag, 4)}
                if build_grid:
                    # Downsample the change mask to a small, thresholded grid
                    # (mirrors flow_grid) recording the fraction of pixels changed
                    # per cell, so the Change heatmap can show where pixels move
                    # without bloating the per-frame results.
                    cells = (
                        cv2.resize(
                            mask, (grid, grid), interpolation=cv2.INTER_AREA
                        ).astype(np.float32)
                        / 255.0
                    )
                    ys, xs = np.nonzero(cells >= min_frac)
                    rd["change_grid"] = [
                        {
                            "x": round((int(x) + 0.5) / grid, 3),
                            "y": round((int(y) + 0.5) / grid, 3),
                            "mag": round(float(cells[y, x]), 3),
                        }
                        for y, x in zip(ys, xs)
                    ]
                emitted = buf.push(ts, rd)
                if emitted is not None:
                    results.append(emitted)
                    if on_result:
                        on_result(emitted)
            else:
                buf.reset()
        prev_pixels[0] = pixels
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_frames(
        video_path,
        region,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


def scan_similarity(
    video_path: str,
    region: dict[str, int],
    reference_frame: np.ndarray,
    threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Find frames where region is similar to a reference.

    Returns list of ``{timestamp, score}`` dicts, sorted by score
    descending.
    """
    from skimage.metrics import structural_similarity as ssim

    if threshold <= 0:
        threshold = config.SCREENSPACE_SSIM_THRESHOLD
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    window = _resolve_scan_window(video_path, start_seconds, end_seconds)
    if window is None:
        return []
    vid_fps, vid_duration, end_seconds, total_range = window

    results: list[dict[str, Any]] = []
    ref_phash = compute_phash(reference_frame)
    phash_threshold = config.SCREENSPACE_PHASH_THRESHOLD

    # Pre-resize and preprocess reference frame once
    max_dim = 256
    rh, rw = reference_frame.shape[:2]
    if rh > max_dim or rw > max_dim:
        scale = max_dim / max(rh, rw)
        cmp_w, cmp_h = int(rw * scale), int(rh * scale)
        ref_resized = cv2.resize(
            reference_frame, (cmp_w, cmp_h), interpolation=cv2.INTER_AREA
        )
    else:
        cmp_w, cmp_h = rw, rh
        ref_resized = reference_frame
    bk = config.SCREENSPACE_BLUR_KERNEL
    ref_gray = cv2.cvtColor(
        cv2.GaussianBlur(ref_resized, (bk, bk), 0), cv2.COLOR_BGR2GRAY
    )

    prev_skip_gray: list[np.ndarray | None] = [None]

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        # Static-frame skip
        gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
        if _frame_is_static(prev_skip_gray[0], gray):
            if on_progress and total_range > 0:
                on_progress((ts - start_seconds) / total_range)
            return None
        prev_skip_gray[0] = gray

        frame_phash = compute_phash(pixels)
        if ref_phash - frame_phash <= phash_threshold:
            # Always resize candidate to match reference dimensions for SSIM
            ph, pw = pixels.shape[:2]
            if pw != cmp_w or ph != cmp_h:
                cand = cv2.resize(pixels, (cmp_w, cmp_h), interpolation=cv2.INTER_AREA)
            else:
                cand = pixels
            cand_gray = cv2.cvtColor(
                cv2.GaussianBlur(cand, (bk, bk), 0), cv2.COLOR_BGR2GRAY
            )
            score = float(ssim(ref_gray, cand_gray))
            if score >= threshold:
                rd = {"timestamp": ts, "score": round(score, 4)}
                results.append(rd)
                if on_result:
                    on_result(rd)
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_frames(
        video_path,
        region,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    results.sort(key=lambda r: r["score"], reverse=True)
    return results


def scan_text(
    video_path: str,
    region: dict[str, Any],
    search_string: str,
    interval_seconds: float = 2.0,
    *,
    fuzzy_threshold: float = 0.0,
    ocr_confidence_threshold: float | None = None,
    ocr_preprocess: bool = False,
    ocr_normalize: str = "off",
    require_consecutive: int = 1,
    languages: list[str] | None = None,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan for text appearances in a region using EasyOCR.

    EasyOCR is lazy-imported. Raises ``ImportError`` with install
    instructions if missing.
    """
    if fuzzy_threshold <= 0:
        fuzzy_threshold = config.SCREENSPACE_OCR_FUZZY_THRESHOLD
    ocr_confidence_threshold = _effective_ocr_confidence_threshold(
        ocr_confidence_threshold
    )
    utils.require_optional("easyocr", "text scan")
    if languages is None:
        languages = ["en"]

    window = _resolve_scan_window(video_path, start_seconds, end_seconds)
    if window is None:
        return []
    vid_fps, vid_duration, end_seconds, total_range = window

    results: list[dict[str, Any]] = []
    text_params: dict[str, Any] = {
        "search_string": search_string,
        "fuzzy_threshold": fuzzy_threshold,
        "ocr_confidence_threshold": ocr_confidence_threshold,
        "ocr_normalize": ocr_normalize,
    }
    prev_gray: list[np.ndarray | None] = [None]
    buf = _ConsecutiveBuffer(require_consecutive)
    mask_points = region.get("mask_points")

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        if _is_static_skip(
            ts,
            pixels,
            prev_gray,
            buf,
            results,
            on_result,
            on_progress,
            start_seconds,
            total_range,
        ):
            return None
        readings = _ocr_region_readings(
            pixels,
            languages=languages,
            preprocess=ocr_preprocess,
            mask_points=mask_points,
        )
        passed, detail = _score_text_readings(readings, text_params)
        matched_rd = {"timestamp": ts, **detail} if passed else None
        if matched_rd is not None:
            emitted = buf.push(ts, matched_rd)
            if emitted is not None:
                results.append(emitted)
                if on_result:
                    on_result(emitted)
        else:
            buf.reset()
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_frames(
        video_path,
        region,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


def scan_numbers(
    video_path: str,
    region: dict[str, Any],
    operator: str,
    target_value: float = 0,
    interval_seconds: float = 2.0,
    *,
    range_min: float | None = None,
    range_max: float | None = None,
    ocr_confidence_threshold: float | None = None,
    ocr_preprocess: bool = False,
    integers_only: bool = False,
    require_consecutive: int = 1,
    languages: list[str] | None = None,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan for numeric values in a region and apply a comparison.

    Uses EasyOCR to detect text, parses numbers from it, and returns
    timestamps where the detected number satisfies the comparison.
    """
    if operator not in _VALID_OPERATORS:
        raise ValueError(
            f"Unknown operator '{operator}'. Must be one of: {', '.join(_VALID_OPERATORS)}"
        )

    ocr_confidence_threshold = _effective_ocr_confidence_threshold(
        ocr_confidence_threshold
    )
    utils.require_optional("easyocr", "numbers scan")
    if languages is None:
        languages = ["en"]

    window = _resolve_scan_window(video_path, start_seconds, end_seconds)
    if window is None:
        return []
    vid_fps, vid_duration, end_seconds, total_range = window

    results: list[dict[str, Any]] = []
    prev_gray: list[np.ndarray | None] = [None]
    numbers_params: dict[str, Any] = {
        "operator": operator,
        "target_value": target_value,
        "range_min": range_min,
        "range_max": range_max,
        "ocr_confidence_threshold": ocr_confidence_threshold,
    }
    # Hoisted out of the per-frame callback: constrain English OCR to digits.
    numbers_allowlist = _numbers_ocr_allowlist(languages, integers_only)
    buf = _ConsecutiveBuffer(require_consecutive)
    mask_points = region.get("mask_points")

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        if _is_static_skip(
            ts,
            pixels,
            prev_gray,
            buf,
            results,
            on_result,
            on_progress,
            start_seconds,
            total_range,
        ):
            return None
        readings = _ocr_region_readings(
            pixels,
            languages=languages,
            allowlist=numbers_allowlist,
            preprocess=ocr_preprocess,
            mask_points=mask_points,
        )
        passed, detail = _score_numbers_readings(readings, numbers_params)
        matched_rd = {"timestamp": ts, **detail} if passed else None
        if matched_rd is not None:
            emitted = buf.push(ts, matched_rd)
            if emitted is not None:
                results.append(emitted)
                if on_result:
                    on_result(emitted)
        else:
            buf.reset()
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_frames(
        video_path,
        region,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


def generate_timelapse(
    video_path: str,
    region: dict[str, int],
    speedup_factor: float,
    output_path: str,
    output_format: str = "mp4",
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    sample_interval: float = 0.0,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
) -> str | None:
    """Generate a cropped timelapse via ffmpeg.

    *sample_interval* controls frame sampling (seconds between captured
    frames).  0 means every frame is used.

    Reports encoding progress through *on_progress* by parsing ffmpeg's
    ``-progress`` output.  Checks *cancel_flag* periodically to allow
    early termination.

    Returns output file path on success, ``None`` on failure.
    """
    cmd = build_timelapse_command(
        video_path,
        region,
        speedup_factor,
        output_path,
        output_format,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        sample_interval=sample_interval,
    )

    # Estimate output duration for progress calculation.
    input_duration = 0.0
    if end_seconds is not None and end_seconds > start_seconds:
        input_duration = end_seconds - start_seconds
    else:
        _, vid_dur = _probe_video_meta(video_path)
        if vid_dur > 0:
            input_duration = max(vid_dur - start_seconds, 0.0)

    expected_out_us = (
        (input_duration / speedup_factor * 1_000_000) if input_duration > 0 else 0
    )

    # Use Popen with -progress to get real-time encoding updates
    cmd += ["-progress", "pipe:1"]
    try:
        # stderr → DEVNULL: only stdout (progress lines) is read, so a PIPE'd
        # stderr could fill its 64 KB OS buffer and deadlock ffmpeg.
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
        assert proc.stdout is not None  # guaranteed by stdout=PIPE
    except (FileNotFoundError, OSError):
        return None

    if on_progress:
        on_progress(0.0)

    try:
        for line in proc.stdout:
            text = line.decode("utf-8", errors="replace").strip()
            if text.startswith("out_time_us=") and expected_out_us > 0 and on_progress:
                try:
                    us = int(text.split("=", 1)[1])
                    on_progress(min(us / expected_out_us, 0.99))
                except (ValueError, ZeroDivisionError):
                    pass
            if cancel_flag and cancel_flag():
                utils.terminate_subprocess(proc)
                return None
    finally:
        if proc.stdout:
            proc.stdout.close()
        proc.wait()

    if on_progress:
        on_progress(1.0)
    if proc.returncode != 0:
        return None
    return output_path


def scan_template(
    video_path: str,
    region: dict[str, int],
    template_image: np.ndarray,
    threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    template_mask: np.ndarray | None = None,
    template_scale: float = 1.0,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan video for frames containing the template image.

    *template_scale* resizes the uploaded template before matching
    (e.g. ``0.5`` for a template captured at 2x the in-video scale).  An
    optional *template_mask* restricts matching to non-transparent regions
    of an uploaded PNG.

    Returns list of ``{timestamp, matches, best_score, match_count}`` dicts.
    """
    if threshold <= 0:
        threshold = config.SCREENSPACE_TEMPLATE_MATCH_THRESHOLD
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    window = _resolve_scan_window(video_path, start_seconds, end_seconds)
    if window is None:
        return []
    vid_fps, vid_duration, end_seconds, total_range = window

    results: list[dict[str, Any]] = []

    _tmpl_downscale = bool(fast_opts and fast_opts.get("template_downscale"))

    scaled_template, scaled_mask = _scale_template(
        template_image, template_mask, template_scale
    )

    # Hoist the constant template prep (blur + grayscale + variance check)
    # out of the per-frame callback. A degenerate template yields [] every
    # frame, so bail early before even opening the ffmpeg pipe.
    _prepared = _prepare_template(scaled_template, scaled_mask)
    if _prepared[2]:  # degenerate template
        if on_progress:
            on_progress(1.0)
        return results

    _nms_overlap = config.SCREENSPACE_TEMPLATE_NMS_OVERLAP
    # Frame is already at cv_scale * original due to ffmpeg scaling. We need
    # to undo both the user-set cv_scale and the fast-scan internal 2x
    # downscale to report match boxes in original-frame pixels.
    _cv_scale = (
        config.SCREENSPACE_CV_RESOLUTION_SCALE
        if config.SCREENSPACE_CV_RESOLUTION_SCALE > 0
        else 1.0
    )

    # Static-frame carry: template results are per-frame rows with no consecutive
    # buffer, so a naive skip would drop a persistent match's rows and make the
    # detection flicker. Cache the last processed frame's result and, on a static
    # frame, re-emit it (re-stamped) instead of re-running the expensive match.
    prev_skip_gray: list[np.ndarray | None] = [None]
    last_rd: list[dict[str, Any] | None] = [None]

    def _cb(ts: float, frame: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        curr_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        if _frame_is_static(prev_skip_gray[0], curr_gray):
            # Near-duplicate of the last matched frame — matches (already in
            # original-frame coords) still hold. Carry the row forward; keep
            # prev_skip_gray as the baseline so drift out of the run recomputes.
            if last_rd[0] is not None:
                carried = dict(last_rd[0])
                carried["timestamp"] = ts
                results.append(carried)
                if on_result:
                    on_result(
                        {
                            "timestamp": ts,
                            "best_score": carried["best_score"],
                            "match_count": carried["match_count"],
                        }
                    )
            if on_progress and total_range > 0:
                on_progress((ts - start_seconds) / total_range)
            return None
        prev_skip_gray[0] = curr_gray
        work_frame = frame
        scale_back = 1
        if _tmpl_downscale:
            fh, fw = work_frame.shape[:2]
            nw, nh = fw // 2, fh // 2
            if nw > 0 and nh > 0:
                work_frame = cv2.resize(
                    work_frame, (nw, nh), interpolation=cv2.INTER_AREA
                )
                scale_back = 2
        matches = _match_template_prepared(
            work_frame, _prepared, threshold, _nms_overlap
        )
        if matches:
            # Undo fast-scan downscale, then undo cv_scale, so reported
            # coords are in the original (un-scaled) frame coordinate space.
            inv = scale_back / _cv_scale
            if abs(inv - 1.0) > 1e-6:
                for m in matches:
                    m["x"] = int(round(m["x"] * inv))
                    m["y"] = int(round(m["y"] * inv))
                    m["w"] = int(round(m["w"] * inv))
                    m["h"] = int(round(m["h"] * inv))
            # Shaped region: the match itself runs full-frame (as for rects,
            # which don't restrict template search either), but detections
            # whose center falls outside the polygon are dropped. Runs before
            # the static-frame carry caches last_rd, so carried rows are
            # already filtered. This mask is the *region's* shape — distinct
            # from template_mask, the template's own alpha channel.
            matches = filter_matches_by_region_mask(matches, region)
        if matches:
            best = max(m["score"] for m in matches)
            rd = {
                "timestamp": ts,
                "matches": matches,
                "best_score": round(best, 4),
                "match_count": len(matches),
            }
            results.append(rd)
            last_rd[0] = rd
            if on_result:
                on_result(
                    {
                        "timestamp": ts,
                        "best_score": rd["best_score"],
                        "match_count": rd["match_count"],
                    }
                )
        else:
            # No match this frame — a subsequent static frame has nothing to carry.
            last_rd[0] = None
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_full_frames(
        video_path,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


def scan_flow(
    video_path: str,
    region: dict[str, int],
    magnitude_threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    require_consecutive: int = 1,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan video for motion in a region using dense optical flow.

    Returns list of ``{timestamp, magnitude, angle}`` dicts where
    magnitude exceeds *magnitude_threshold*.
    """
    if magnitude_threshold <= 0:
        magnitude_threshold = config.SCREENSPACE_FLOW_MAGNITUDE_THRESHOLD
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    window = _resolve_scan_window(video_path, start_seconds, end_seconds)
    if window is None:
        return []
    vid_fps, vid_duration, end_seconds, total_range = window

    results: list[dict[str, Any]] = []
    prev_gray: list[np.ndarray | None] = [None]
    buf = _ConsecutiveBuffer(require_consecutive)
    mask_for = region_masker(region)

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        curr_gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
        # Static-frame skip: a near-duplicate of the previous frame produces
        # ~zero optical flow, so the expensive Farneback pass would only confirm
        # magnitude < threshold and reset the run anyway. Short-circuit to that
        # outcome, still advancing prev_gray (flow is measured frame-to-frame).
        if _frame_is_static(prev_gray[0], curr_gray):
            buf.reset()
            prev_gray[0] = curr_gray
            if on_progress and total_range > 0:
                on_progress((ts - start_seconds) / total_range)
            return None
        if prev_gray[0] is not None:
            flow_result = compute_optical_flow(
                prev_gray[0], curr_gray, return_grid=True, mask=mask_for(pixels)
            )
            if flow_result["magnitude"] >= magnitude_threshold:
                rd: dict[str, Any] = {
                    "timestamp": ts,
                    "magnitude": flow_result["magnitude"],
                    "angle": flow_result["angle"],
                    "flow_grid": flow_result.get("flow_grid", []),
                }
                emitted = buf.push(ts, rd)
                if emitted is not None:
                    results.append(emitted)
                    if on_result:
                        # Keep incremental updates lightweight (no grid)
                        on_result(
                            {
                                "timestamp": emitted["timestamp"],
                                "magnitude": emitted["magnitude"],
                                "angle": emitted["angle"],
                            }
                        )
            else:
                buf.reset()
        prev_gray[0] = curr_gray
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_frames(
        video_path,
        region,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


def scan_scene(
    video_path: str,
    region: dict[str, int],
    reference_scenes: list[dict[str, Any]],
    threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Classify frames by similarity to reference scene fingerprints.

    *reference_scenes* is a list of ``{name: str, frame: np.ndarray}``
    dicts, optionally with a per-scene ``threshold`` float.  Each
    frame's region is fingerprinted and compared against all references;
    the best match above its threshold is reported.

    Returns list of ``{timestamp, scene_name, score}`` dicts.
    """
    default_threshold = (
        threshold if threshold > 0 else config.SCREENSPACE_SCENE_SIMILARITY_THRESHOLD
    )
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    # Pre-compute fingerprints for reference scenes (with per-scene thresholds).
    # Shaped regions: fingerprints are only comparable when both sides use the
    # mask, so each reference crop gets the mask rasterized at its own size
    # (references are source-resolution; scan crops may be rescaled).
    mask_for = region_masker(region)
    ref_fps: list[tuple[str, dict[str, Any], float]] = []
    for ref in reference_scenes:
        fp = compute_scene_fingerprint(ref["frame"], mask=mask_for(ref["frame"]))
        ref_thresh = float(ref.get("threshold", default_threshold))
        ref_fps.append((ref["name"], fp, ref_thresh))

    if not ref_fps:
        return []

    window = _resolve_scan_window(video_path, start_seconds, end_seconds)
    if window is None:
        return []
    vid_fps, vid_duration, end_seconds, total_range = window

    results: list[dict[str, Any]] = []
    prev_skip_gray: list[np.ndarray | None] = [None]

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False

        # Static-frame skip
        curr_gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
        if _frame_is_static(prev_skip_gray[0], curr_gray):
            if on_progress and total_range > 0:
                on_progress((ts - start_seconds) / total_range)
            return None
        prev_skip_gray[0] = curr_gray

        fp = compute_scene_fingerprint(pixels, mask=mask_for(pixels))
        best_name = ""
        best_score = 0.0
        best_thresh = default_threshold
        for ref_name, ref_fp, ref_thresh in ref_fps:
            score = compare_scene_fingerprints(fp, ref_fp)
            if score > best_score:
                best_score = score
                best_name = ref_name
                best_thresh = ref_thresh

        if best_score >= best_thresh:
            rd = {
                "timestamp": ts,
                "scene_name": best_name,
                "score": round(best_score, 4),
            }
            results.append(rd)
            if on_result:
                on_result(rd)
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_frames(
        video_path,
        region,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


def scan_inactivity(
    video_path: str,
    region: dict[str, int],
    threshold: int = 0,
    min_duration: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan video for inactivity spans (near-duplicate consecutive frames).

    Uses perceptual hashing to compare consecutive frames.  When frames
    remain similar (Hamming distance <= *threshold*) for at least
    *min_duration* seconds, a span result is emitted.

    Returns list of ``{start, end, duration, avg_distance}`` dicts.
    """
    if threshold <= 0:
        threshold = config.SCREENSPACE_INACTIVITY_PHASH_THRESHOLD
    if min_duration <= 0:
        min_duration = config.SCREENSPACE_INACTIVITY_MIN_DURATION
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    window = _resolve_scan_window(video_path, start_seconds, end_seconds)
    if window is None:
        return []
    vid_fps, vid_duration, end_seconds, total_range = window

    results: list[dict[str, Any]] = []
    prev_hash: list["imagehash.ImageHash | None"] = [None]
    prev_skip_gray: list[np.ndarray | None] = [None]
    span_start: list[float | None] = [None]
    span_distances: list[list[int]] = [[]]
    last_ts: list[float] = [start_seconds]

    def _extend_span(ts: float, dist: int) -> None:
        # Frame is similar — extend or start span.
        if span_start[0] is None:
            # Clamp to the scan start so a match early in the video
            # (ts < interval_seconds, or start_seconds > 0) can't begin
            # the span before 0:00 / the requested start.
            span_start[0] = max(start_seconds, ts - interval_seconds)
        span_distances[0].append(dist)

    def _flush_span(span_end: float) -> None:
        if span_start[0] is not None:
            dur = span_end - span_start[0]
            if dur >= min_duration:
                dists = span_distances[0]
                avg_dist = sum(dists) / len(dists) if dists else 0.0
                rd = {
                    "start": round(span_start[0], 2),
                    "end": round(span_end, 2),
                    "duration": round(dur, 2),
                    "avg_distance": round(avg_dist, 2),
                }
                results.append(rd)
                if on_result:
                    on_result(rd)
        span_start[0] = None
        span_distances[0] = []

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False

        last_ts[0] = ts

        # Static-frame fast-path: a gray mean-diff below the static threshold
        # implies a phash Hamming distance of ~0 — well within the (always ≥1)
        # inactivity threshold — so the frame is provably inactive. Extend the
        # span with a nominal 0 distance and skip the much heavier compute_phash.
        # prev_hash/prev_skip_gray are left as the run baseline so slow drift is
        # still measured against the start of the frozen run.
        curr_gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
        if _frame_is_static(prev_skip_gray[0], curr_gray):
            _extend_span(ts, 0)
            if on_progress and total_range > 0:
                on_progress((ts - start_seconds) / total_range)
            return None
        prev_skip_gray[0] = curr_gray

        curr_hash = compute_phash(pixels)

        if prev_hash[0] is not None:
            dist = int(curr_hash - prev_hash[0])
            if dist <= threshold:
                _extend_span(ts, dist)
            else:
                # Frame changed — flush any active span
                _flush_span(ts - interval_seconds)
        prev_hash[0] = curr_hash

        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_frames(
        video_path,
        region,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=fast_opts,
    )

    # Flush final span if video ended during an inactive period
    _flush_span(last_ts[0])

    if on_progress:
        on_progress(1.0)
    return results


def _boundary_rep_pixels(pixels: np.ndarray) -> np.ndarray:
    """Downscale a settled frame to a small period representative (max 64 px).

    Retaining the frame (≈12 KB) rather than its HSV histogram (≈1 MB) keeps the
    per-period memory tiny; the post-run pass recomputes fingerprints on demand.
    """
    h, w = pixels.shape[:2]
    longest = max(h, w)
    if longest > 64:
        scale = 64.0 / longest
        return cv2.resize(
            pixels,
            (max(1, int(w * scale)), max(1, int(h * scale))),
            interpolation=cv2.INTER_AREA,
        )
    return pixels.copy()


def _scene_label(index: int) -> str:
    """Human-readable scene label for a period's cluster index (0 → 'Scene A')."""
    return f"Scene {utils.index_to_letter(index)}"


def _boundaries_from_periods(
    periods: list[dict[str, Any]], end_seconds: float
) -> list[dict[str, Any]]:
    """Turn a consolidated period list into boundary result dicts.

    ``periods[0]`` is the initial period (video start, no entering boundary); a
    boundary exists at the start of every period after it. Each result carries
    the span it opens as ``period_start``/``period_end`` and the label of the
    scene it enters as ``scene_label`` (set by the labeler; falls back to a
    sequential label by position when absent).
    """
    results: list[dict[str, Any]] = []
    for k in range(1, len(periods)):
        p = periods[k]
        period_end = (
            periods[k + 1]["start_ts"]
            if k + 1 < len(periods)
            else round(end_seconds, 2)
        )
        results.append(
            {
                "timestamp": p["start_ts"],
                "distance": p["entry_dist"],
                "_confidence": round(p["entry_conf"], 4),
                "period_start": p["start_ts"],
                "period_end": period_end,
                "scene_label": p.get("scene_label") or _scene_label(k),
            }
        )
    return results


def _consolidate_boundary_periods(
    periods: list[dict[str, Any]],
    *,
    end_seconds: float,
    merge_threshold: float,
    short_period_seconds: float,
    relative_prune_enabled: bool,
    relative_prune_factor: float,
    type_threshold: float = 1.0,
    min_boundaries_for_prune: int = 4,
) -> list[dict[str, Any]]:
    """Post-run sanity pass over the scene/hybrid period list.

    Principle: *merge periods that aren't really different, and dissolve
    transient interruptions.* A forward-only scan can't know that the scene a
    few seconds ahead is identical to the one before a boundary; this global
    pass can. Each period carries a settled-frame ``pixels`` representative;
    fingerprints are computed lazily here and cached on the period dict.

    Steps: (1) a fixed-point merge loop — drop a boundary when the two periods
    it separates are the same scene (rule B), or dissolve a *short* transient
    period bracketed by identical scenes (rule A); (2) a session-relative prune
    that drops boundaries far weaker than the median (gated).
    """
    periods = [dict(p) for p in periods]
    if len(periods) <= 1:
        return _boundaries_from_periods(periods, end_seconds)

    def _fp(p: dict[str, Any]) -> dict[str, Any]:
        fp = p.get("_fp")
        if fp is None:
            fp = compute_scene_fingerprint(p["pixels"])
            p["_fp"] = fp
        return fp

    def _scene_dist(a: dict[str, Any], b: dict[str, Any]) -> float:
        return 1.0 - compare_scene_fingerprints(_fp(a), _fp(b))

    def _run_merge_passes(plist: list[dict[str, Any]]) -> list[dict[str, Any]]:
        changed = True
        while changed and len(plist) > 1:
            changed = False
            # Rule B: adjacent-merge — same scene on both sides → spurious boundary.
            i = 1
            while i < len(plist):
                if _scene_dist(plist[i - 1], plist[i]) < merge_threshold:
                    del plist[i]  # boundary i removed; period i absorbed into i-1
                    changed = True
                else:
                    i += 1
            # Rule A: round-trip — a short transient bracketed by the same scene.
            i = 1
            while i < len(plist) - 1:
                duration = plist[i + 1]["start_ts"] - plist[i]["start_ts"]
                if (
                    duration < short_period_seconds
                    and _scene_dist(plist[i - 1], plist[i + 1]) < merge_threshold
                ):
                    del plist[i + 1]  # drop both boundaries bracketing the transient
                    del plist[i]
                    changed = True
                else:
                    i += 1
        return plist

    periods = _run_merge_passes(periods)

    # Session-relative prune: drop boundaries far below the median strength.
    if relative_prune_enabled and (len(periods) - 1) >= min_boundaries_for_prune:
        dists = sorted(float(p["entry_dist"]) for p in periods[1:])
        n = len(dists)
        median = dists[n // 2] if n % 2 else (dists[n // 2 - 1] + dists[n // 2]) / 2.0
        cutoff = relative_prune_factor * median
        periods = [periods[0]] + [
            p for p in periods[1:] if float(p["entry_dist"]) >= cutoff
        ]
        # Dropping a weak middle period can leave two same-scene neighbors
        # adjacent; re-run the merge passes to collapse the duplicate boundary.
        periods = _run_merge_passes(periods)

    # Hierarchical scene labels (Scene A1, A2, B1, …): the letter is the *type*
    # (similar scenes grouped at the looser type_threshold), the number is the
    # distinct *scene* within that type (exact recurrence at merge_threshold). A
    # revisited scene reuses its full label; a similar-but-distinct scene shares
    # the letter with a new number.
    #
    # 1. Tight clustering → a distinct scene id per period (exact recurrence).
    scene_reps: list[dict[str, Any]] = []
    for p in periods:
        scene_id = None
        for ci, rep in enumerate(scene_reps):
            if 1.0 - compare_scene_fingerprints(_fp(p), rep) < merge_threshold:
                scene_id = ci
                break
        if scene_id is None:
            scene_id = len(scene_reps)
            scene_reps.append(_fp(p))
        p["_scene_id"] = scene_id

    # 2. Loose clustering over the scene representatives → a type id per scene,
    #    so every period of one scene shares a type.
    type_of_scene: list[int] = []
    type_reps: list[dict[str, Any]] = []
    for rep in scene_reps:
        type_id = None
        for ci, trep in enumerate(type_reps):
            if 1.0 - compare_scene_fingerprints(rep, trep) < type_threshold:
                type_id = ci
                break
        if type_id is None:
            type_id = len(type_reps)
            type_reps.append(rep)
        type_of_scene.append(type_id)

    # 3. Number the distinct scenes within each type, in first-appearance order.
    type_scenes: dict[int, list[int]] = {}
    for sid in range(len(scene_reps)):
        type_scenes.setdefault(type_of_scene[sid], []).append(sid)

    # 4. Label: Scene <type letter><scene number within type>.
    for p in periods:
        sid = p["_scene_id"]
        tid = type_of_scene[sid]
        scene_num = type_scenes[tid].index(sid) + 1
        p["scene_label"] = f"Scene {utils.index_to_letter(tid)}{scene_num}"

    return _boundaries_from_periods(periods, end_seconds)


def scan_boundaries(
    video_path: str,
    region: dict[str, int] | None = None,
    threshold: int = 0,
    min_gap: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    metric: str = "phash",
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan the full frame for scene boundaries.

    Three metrics (``metric``):

    - ``"phash"`` (v1): a boundary fires when the perceptual-hash Hamming
      distance to the *previous* sampled frame is ≥ *threshold*, debounced by
      *min_gap* seconds. Streams each boundary live via *on_result*.
    - ``"scene"``: each sample's content fingerprint is measured against the
      *current period's* reference (not the previous frame); a boundary fires
      only when the fingerprint distance crosses
      ``SCREENSPACE_BOUNDARY_SCENE_THRESHOLD`` and *holds* for
      ``SCREENSPACE_BOUNDARY_CONFIRM_WINDOW`` samples. Robust to motion.
    - ``"hybrid"``: a confirmed scene shift that is *also* corroborated by a
      phash spike — catches hard cuts and rejects both motion (phash spikes
      that don't sustain) and slow fades (drift with no spike).

    ``scene``/``hybrid`` run a post-run consolidation pass
    (:func:`_consolidate_boundary_periods`) and emit ``on_result`` only for the
    *final* boundaries (the worker derives events from the stream), so the
    progress bar advances live but ticks appear together at completion. The
    function param defaults to ``"phash"``; the tool layer applies the policy
    default (``config.SCREENSPACE_BOUNDARY_METRIC``).

    Region is ignored (full-frame only); the parameter exists for signature
    parity. Returns ``{timestamp, distance, _confidence}`` dicts (scene/hybrid
    also carry ``period_start``/``period_end``).
    """
    if threshold <= 0:
        threshold = config.SCREENSPACE_BOUNDARY_PHASH_THRESHOLD
    if min_gap <= 0:
        min_gap = config.SCREENSPACE_BOUNDARY_MIN_GAP_SECONDS
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_BOUNDARY_INTERVAL
    metric = (metric or "phash").strip().lower()
    if metric not in ("phash", "scene", "hybrid"):
        metric = "phash"
    use_scene = metric in ("scene", "hybrid")
    is_hybrid = metric == "hybrid"

    window = _resolve_scan_window(video_path, start_seconds, end_seconds)
    if window is None:
        return []
    vid_fps, vid_duration, end_seconds, total_range = window

    # Downscale frames at the ffmpeg pipe. Scene/hybrid fingerprinting needs more
    # detail than the coarse phash dim (the HSV histogram is too sparse at 64 px).
    # We pass only ``max_region_dim`` (no ``phash_skip``) so the pipe downsizes
    # without dropping frames — this scanner samples every interval itself.
    boundary_opts = dict(fast_opts or {})
    boundary_opts.setdefault(
        "max_region_dim",
        config.SCREENSPACE_BOUNDARY_SCENE_HASH_DIM
        if use_scene
        else config.SCREENSPACE_BOUNDARY_HASH_DIM,
    )
    boundary_opts.pop("phash_skip", None)

    results: list[dict[str, Any]] = []
    prev_hash: list["imagehash.ImageHash | None"] = [None]
    last_boundary_ts: list[float | None] = [None]
    eps = config.SCREENSPACE_BOUNDARY_CONFIDENCE_EPSILON

    if not use_scene:
        # ---- v1 phash path: consecutive-frame spike, streamed live ----
        prev_skip_gray: list[np.ndarray | None] = [None]

        def _cb_phash(ts: float, pixels: np.ndarray) -> bool | None:
            if cancel_flag and cancel_flag():
                return False
            # Static-frame skip: a near-duplicate frame yields a phash distance
            # ~0, far below the boundary threshold, so it can never be a scene
            # boundary. Skip the heavy compute_phash and keep prev_hash as the
            # run baseline (a real spike is never gray-static, so none is missed).
            curr_gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
            if _frame_is_static(prev_skip_gray[0], curr_gray):
                if on_progress and total_range > 0:
                    on_progress((ts - start_seconds) / total_range)
                return None
            prev_skip_gray[0] = curr_gray
            curr_hash = compute_phash(pixels)
            if prev_hash[0] is not None:
                dist = int(curr_hash - prev_hash[0])
                within_gap = (
                    last_boundary_ts[0] is not None
                    and ts - last_boundary_ts[0] < min_gap
                )
                if dist >= threshold and not within_gap:
                    conf = (
                        1.0 if threshold <= 0 else (dist - threshold) / float(threshold)
                    )
                    conf = max(eps, min(conf, 1.0))
                    rd = {
                        "timestamp": round(ts, 2),
                        "distance": dist,
                        "_confidence": round(conf, 4),
                        # phash has no fingerprints to cluster, so labels are
                        # sequential: the Nth boundary opens the (N+1)th segment.
                        "scene_label": _scene_label(len(results) + 1),
                    }
                    results.append(rd)
                    if on_result:
                        on_result(rd)
                    last_boundary_ts[0] = ts
            prev_hash[0] = curr_hash
            if on_progress and total_range > 0:
                on_progress((ts - start_seconds) / total_range)
            return None

        scan_video_full_frames(
            video_path,
            interval_seconds,
            _cb_phash,
            start_seconds=start_seconds,
            end_seconds=end_seconds,
            fps=vid_fps,
            duration=vid_duration,
            fast_opts=boundary_opts,
        )
        if on_progress:
            on_progress(1.0)
        return results

    # ---- scene / hybrid path: period-reference model + post-run pass ----
    scene_threshold = config.SCREENSPACE_BOUNDARY_SCENE_THRESHOLD
    confirm_window = max(1, config.SCREENSPACE_BOUNDARY_CONFIRM_WINDOW)
    ref_fp: list[dict[str, Any] | None] = [None]
    periods: list[dict[str, Any]] = []
    # Pending run of consecutive samples above threshold vs the period reference.
    pending: dict[str, Any] = {"count": 0, "start_ts": 0.0, "phash_seen": False}

    def _reset_pending() -> None:
        pending["count"] = 0
        pending["start_ts"] = 0.0
        pending["phash_seen"] = False

    def _cb_scene(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False

        fp = compute_scene_fingerprint(pixels)
        phash_spike = False
        if is_hybrid:
            curr_hash = compute_phash(pixels)
            if prev_hash[0] is not None:
                phash_spike = int(curr_hash - prev_hash[0]) >= threshold
            prev_hash[0] = curr_hash

        ref = ref_fp[0]
        if ref is None:
            # First sample seeds the initial period (video start, no boundary).
            ref_fp[0] = fp
            periods.append(
                {
                    "start_ts": round(start_seconds, 2),
                    "pixels": _boundary_rep_pixels(pixels),
                    "entry_dist": 0,
                    "entry_conf": 0.0,
                }
            )
        else:
            dist_raw = 1.0 - compare_scene_fingerprints(fp, ref)
            if dist_raw >= scene_threshold:
                if pending["count"] == 0:
                    pending["start_ts"] = ts
                    pending["phash_seen"] = False
                pending["count"] += 1
                if phash_spike:
                    pending["phash_seen"] = True
                if pending["count"] >= confirm_window:
                    # A sustained shift — we have entered a new scene. Advance the
                    # period reference NOW, regardless of whether we emit a
                    # boundary: min_gap (and hybrid's phash gate) only suppress the
                    # boundary *event*, not the fact that the content moved on.
                    # Tying reference advancement to emission would leave ref_fp
                    # stuck on the old scene for the rest of the clip when a
                    # transition lands within min_gap, silently dropping every
                    # later boundary.
                    within_gap = (
                        last_boundary_ts[0] is not None
                        and pending["start_ts"] - last_boundary_ts[0] < min_gap
                    )
                    hybrid_ok = (not is_hybrid) or pending["phash_seen"]
                    if not within_gap and hybrid_ok:
                        conf = (
                            1.0
                            if scene_threshold <= 0
                            else (dist_raw - scene_threshold) / scene_threshold
                        )
                        conf = max(eps, min(conf, 1.0))
                        periods.append(
                            {
                                "start_ts": round(pending["start_ts"], 2),
                                "pixels": _boundary_rep_pixels(pixels),
                                "entry_dist": int(round(dist_raw * 100)),
                                "entry_conf": conf,
                            }
                        )
                        last_boundary_ts[0] = pending["start_ts"]
                    ref_fp[0] = fp  # settled reference = the confirming frame
                    _reset_pending()
            else:
                _reset_pending()

        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_full_frames(
        video_path,
        interval_seconds,
        _cb_scene,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=boundary_opts,
    )

    final = _consolidate_boundary_periods(
        periods,
        end_seconds=end_seconds,
        merge_threshold=config.SCREENSPACE_BOUNDARY_MERGE_THRESHOLD,
        short_period_seconds=config.SCREENSPACE_BOUNDARY_SHORT_PERIOD_SECONDS,
        relative_prune_enabled=config.SCREENSPACE_BOUNDARY_RELATIVE_PRUNE_ENABLED,
        relative_prune_factor=config.SCREENSPACE_BOUNDARY_RELATIVE_PRUNE_FACTOR,
        type_threshold=config.SCREENSPACE_BOUNDARY_TYPE_THRESHOLD,
    )
    for rd in final:
        results.append(rd)
        if on_result:
            on_result(rd)
    if on_progress:
        on_progress(1.0)
    return results


def scan_attention(
    video_path: str,
    region: dict[str, int] | None = None,
    shift_threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    ema_alpha: float = 0.0,
    weights: dict[str, float] | None = None,
    center_bias: float | None = None,
    include_face: bool | None = None,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan the full frame for predicted visual attention (saliency).

    Every sampled frame lands in the *returned* list carrying a
    ``saliency_grid`` — that list feeds heatmap generation, where each frame
    contributes one unit of dwell. ``on_result`` streams only *confirmed
    attention shifts* (the EMA-smoothed saliency peak jumped at least
    *shift_threshold* in normalized distance and persisted for
    ``SCREENSPACE_ATTENTION_SHIFT_CONFIRM`` samples); the worker derives
    events from that stream, so the timeline shows shifts, not every sample.

    No phash-skip and no static-frame short-circuit, on purpose: a static
    screen stared at for 30 s must accumulate 30 s of heat (dwell weighting),
    and the EMA + shift-confirm counters assume uniform sampling.

    Region is ignored (full-frame only); the parameter exists for signature
    parity. Returns ``{timestamp, saliency_grid, peak_x, peak_y, peak_value}``
    dicts; shift frames additionally carry ``shift``, ``shift_distance``,
    ``_confidence``, and ``from_x/from_y/to_x/to_y``.
    """
    if shift_threshold <= 0:
        shift_threshold = config.SCREENSPACE_ATTENTION_SHIFT_THRESHOLD
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_ATTENTION_INTERVAL
    if ema_alpha <= 0:
        ema_alpha = config.SCREENSPACE_ATTENTION_EMA_ALPHA
    ema_alpha = min(1.0, ema_alpha)
    shift_confirm = max(1, config.SCREENSPACE_ATTENTION_SHIFT_CONFIRM)

    window = _resolve_scan_window(video_path, start_seconds, end_seconds)
    if window is None:
        return []
    vid_fps, vid_duration, end_seconds, total_range = window

    # Downsize at the ffmpeg pipe without dropping frames (see docstring for
    # why phash-skip is disabled); the saliency math runs at ≤ WORKING_DIM.
    attention_opts = dict(fast_opts or {})
    attention_opts.setdefault(
        "max_region_dim", config.SCREENSPACE_ATTENTION_WORKING_DIM
    )
    attention_opts.pop("phash_skip", None)

    results: list[dict[str, Any]] = []
    prev_gray: list[np.ndarray | None] = [None]
    smoothed: list[np.ndarray | None] = [None]
    # Shift state machine: the last *emitted* focus plus a pending candidate
    # that must persist near its anchor for shift_confirm samples.
    last_emitted: list[tuple[float, float] | None] = [None]
    pending: list[dict[str, Any] | None] = [None]

    def _dist(a: tuple[float, float], b: tuple[float, float]) -> float:
        return float(np.hypot(a[0] - b[0], a[1] - b[1]))

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        combined, curr_gray = compute_saliency_map(
            pixels,
            prev_gray[0],
            weights=weights,
            center_bias=center_bias,
            include_face=include_face,
        )
        prev_gray[0] = curr_gray
        prev_smoothed = smoothed[0]
        if prev_smoothed is not None and prev_smoothed.shape == combined.shape:
            combined = ema_alpha * combined + (1.0 - ema_alpha) * prev_smoothed
        smoothed[0] = combined

        peak_x, peak_y, peak_value = saliency_peak(combined)
        rd: dict[str, Any] = {
            "timestamp": round(ts, 2),
            "saliency_grid": saliency_grid_from_map(
                combined,
                config.SCREENSPACE_ATTENTION_GRID,
                config.SCREENSPACE_ATTENTION_GRID_MIN_MAG,
            ),
            "peak_x": peak_x,
            "peak_y": peak_y,
            "peak_value": peak_value,
        }
        results.append(rd)

        peak = (peak_x, peak_y)
        prev_focus = last_emitted[0]
        if prev_focus is None:
            last_emitted[0] = peak  # seed on the first sample; never an event
        elif _dist(peak, prev_focus) < shift_threshold:
            pending[0] = None
        else:
            cand = pending[0]
            anchor: tuple[float, float] = cand["anchor"] if cand else peak
            if cand is None or _dist(peak, anchor) > 0.5 * shift_threshold:
                # New (or wandered-off) jump target: restart confirmation at
                # this frame, remembering its result dict to stamp on emit.
                new_cand: dict[str, Any] = {
                    "count": 1,
                    "anchor": peak,
                    "peak_value": peak_value,
                    "result": rd,
                }
                cand = new_cand
                anchor = peak
                pending[0] = cand
            else:
                cand["count"] += 1
            if int(cand["count"]) >= shift_confirm:
                shift_distance = _dist(anchor, prev_focus)
                confidence = float(cand["peak_value"]) * min(
                    1.0, shift_distance / (2.0 * shift_threshold)
                )
                confidence = round(max(0.05, min(1.0, confidence)), 4)
                shift_rd: dict[str, Any] = cand["result"]
                shift_rd.update(
                    {
                        "shift": True,
                        "shift_distance": round(shift_distance, 4),
                        "_confidence": confidence,
                        "from_x": prev_focus[0],
                        "from_y": prev_focus[1],
                        "to_x": anchor[0],
                        "to_y": anchor[1],
                    }
                )
                if on_result:
                    stream_rd = dict(shift_rd)
                    stream_rd.pop("saliency_grid", None)
                    on_result(stream_rd)
                last_emitted[0] = anchor
                pending[0] = None

        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_full_frames(
        video_path,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=attention_opts,
    )
    if on_progress:
        on_progress(1.0)
    return results


# ---------------------------------------------------------------------------
# Multitool: per-frame evaluation and multi-factor scan
# ---------------------------------------------------------------------------
