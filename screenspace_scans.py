# -*- coding: utf-8 -*-
"""Screenspace scan workflows (one per analysis tool).

Each scan sweeps a video via the frame-extraction drivers and applies one
analysis primitive (color, change, similarity, text, numbers, timelapse,
template, flow, scene, inactivity, boundary). Scans never call each other.
Imports primitives, OCR helpers, and frame extractors from sibling modules.
"""

import difflib
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
    _is_static_skip,
    _match_template_prepared,
    _merge_timestamp_spans,
    _morph_kernel,
    _prepare_template,
    _scale_template,
    color_matches,
    color_present,
    compare_scene_fingerprints,
    compute_optical_flow,
    compute_phash,
    compute_scene_fingerprint,
)
from screenspace_ocr import (
    _NUMBERS_RE,
    _OCR_DIGITS_ONLY_ALLOWLIST,
    _OCR_NUMBER_ALLOWLIST,
    _VALID_OPERATORS,
    _effective_ocr_confidence_threshold,
    _get_ocr_reader,
    _normalize_ocr_text,
    _number_matches,
    _preprocess_for_ocr,
)
from screenspace_frames import (
    _probe_video_meta,
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

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    matches: list[float] = []

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        if color_mode == "presence":
            matched, conf = color_present(pixels, target_color, tolerance, min_coverage)
        else:
            matched, conf = color_matches(pixels, target_color, tolerance)
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

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []
    prev_gray: list[np.ndarray | None] = [None]
    k = config.SCREENSPACE_BLUR_KERNEL
    morph_kernel = _morph_kernel(config.SCREENSPACE_MORPH_KERNEL)
    buf = _ConsecutiveBuffer(require_consecutive)
    # change_grid feeds only the Change heatmap; skip the per-frame downsample
    # entirely when heatmaps are disabled (the data would just be discarded).
    build_grid = config.SCREENSPACE_GENERATE_CHANGE_HEATMAP
    grid = config.SCREENSPACE_CHANGE_HEATMAP_GRID
    min_frac = config.SCREENSPACE_CHANGE_HEATMAP_MIN_FRAC

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        curr_gray = cv2.cvtColor(
            cv2.GaussianBlur(pixels, (k, k), 0), cv2.COLOR_BGR2GRAY
        )
        if prev_gray[0] is not None:
            diff = cv2.absdiff(prev_gray[0], curr_gray)
            _, mask = cv2.threshold(diff, noise_threshold, 255, cv2.THRESH_BINARY)
            mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, morph_kernel)
            mag = float(np.count_nonzero(mask)) / float(mask.size) if mask.size else 0.0
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

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

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
        if prev_skip_gray[0] is not None:
            if (
                float(np.mean(cv2.absdiff(prev_skip_gray[0], gray)))
                < config.SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD
            ):
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
    region: dict[str, int],
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

    reader = _get_ocr_reader(languages)

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []
    search_cmp = _normalize_ocr_text(search_string, ocr_normalize)
    prev_gray: list[np.ndarray | None] = [None]
    buf = _ConsecutiveBuffer(require_consecutive)

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
        ocr_input = _preprocess_for_ocr(pixels) if ocr_preprocess else pixels
        ocr_results = reader.readtext(ocr_input, detail=1)
        matched_rd: dict[str, Any] | None = None
        for _, text, conf in ocr_results:
            if conf < ocr_confidence_threshold:
                continue
            ocr_cmp = _normalize_ocr_text(text, ocr_normalize)
            ratio = difflib.SequenceMatcher(None, search_cmp, ocr_cmp).ratio()
            if ratio >= fuzzy_threshold:
                matched_rd = {
                    "timestamp": ts,
                    "text_found": text,
                    "confidence": round(conf, 4),
                }
                break
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
    region: dict[str, int],
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

    reader = _get_ocr_reader(languages)

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []
    prev_gray: list[np.ndarray | None] = [None]
    # Hoisted out of the per-frame callback: constrain English OCR to digits.
    ocr_kwargs: dict[str, Any] = {"detail": 1}
    if languages == ["en"]:
        ocr_kwargs["allowlist"] = (
            _OCR_DIGITS_ONLY_ALLOWLIST if integers_only else _OCR_NUMBER_ALLOWLIST
        )
    buf = _ConsecutiveBuffer(require_consecutive)

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
        ocr_input = _preprocess_for_ocr(pixels) if ocr_preprocess else pixels
        ocr_results = reader.readtext(ocr_input, **ocr_kwargs)
        matched_rd: dict[str, Any] | None = None
        for _, text, conf in ocr_results:
            if conf < ocr_confidence_threshold:
                continue
            cleaned = text.replace(",", "")
            for match in _NUMBERS_RE.findall(cleaned):
                num = float(match)
                if _number_matches(num, operator, target_value, range_min, range_max):
                    matched_rd = {
                        "timestamp": ts,
                        "number_found": num,
                        "confidence": round(conf, 4),
                    }
                    break
            if matched_rd is not None:
                break
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

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

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

    def _cb(ts: float, frame: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
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
            best = max(m["score"] for m in matches)
            rd = {
                "timestamp": ts,
                "matches": matches,
                "best_score": round(best, 4),
                "match_count": len(matches),
            }
            results.append(rd)
            if on_result:
                on_result(
                    {
                        "timestamp": ts,
                        "best_score": rd["best_score"],
                        "match_count": rd["match_count"],
                    }
                )
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

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []
    prev_gray: list[np.ndarray | None] = [None]
    buf = _ConsecutiveBuffer(require_consecutive)

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        curr_gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
        if prev_gray[0] is not None:
            flow_result = compute_optical_flow(
                prev_gray[0], curr_gray, return_grid=True
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

    # Pre-compute fingerprints for reference scenes (with per-scene thresholds)
    ref_fps: list[tuple[str, dict[str, Any], float]] = []
    for ref in reference_scenes:
        fp = compute_scene_fingerprint(ref["frame"])
        ref_thresh = float(ref.get("threshold", default_threshold))
        ref_fps.append((ref["name"], fp, ref_thresh))

    if not ref_fps:
        return []

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []
    prev_skip_gray: list[np.ndarray | None] = [None]

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False

        # Static-frame skip (same pattern as similarity scan)
        curr_gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY).astype(np.float32)
        if prev_skip_gray[0] is not None:
            if (
                abs(float(np.mean(curr_gray)) - float(np.mean(prev_skip_gray[0])))
                < config.SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD
            ):
                if on_progress and total_range > 0:
                    on_progress((ts - start_seconds) / total_range)
                return None
        prev_skip_gray[0] = curr_gray

        fp = compute_scene_fingerprint(pixels)
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

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []
    prev_hash: list["imagehash.ImageHash | None"] = [None]
    span_start: list[float | None] = [None]
    span_distances: list[list[int]] = [[]]
    last_ts: list[float] = [start_seconds]

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

        curr_hash = compute_phash(pixels)
        last_ts[0] = ts

        if prev_hash[0] is not None:
            dist = int(curr_hash - prev_hash[0])
            if dist <= threshold:
                # Frame is similar — extend or start span
                if span_start[0] is None:
                    span_start[0] = ts - interval_seconds
                span_distances[0].append(dist)
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


def scan_boundaries(
    video_path: str,
    region: dict[str, int] | None = None,
    threshold: int = 0,
    min_gap: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan the full frame for scene boundaries (large frame-to-frame jumps).

    The inverse of :func:`scan_inactivity`: where inactivity looks for phash
    *plateaus*, boundaries look for *spikes*.  A boundary fires when the
    perceptual-hash Hamming distance to the previous sampled frame is at least
    *threshold*, debounced so that no two boundaries land within *min_gap*
    seconds (the first spike of a sustained high-distance run is the boundary,
    which keeps camera-heavy action sequences from becoming boundary storms).

    Region is ignored (full-frame only); the parameter exists for signature
    parity with the other scanners.  Returns a list of
    ``{timestamp, distance, _confidence}`` dicts.
    """
    if threshold <= 0:
        threshold = config.SCREENSPACE_BOUNDARY_PHASH_THRESHOLD
    if min_gap <= 0:
        min_gap = config.SCREENSPACE_BOUNDARY_MIN_GAP_SECONDS
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_BOUNDARY_INTERVAL

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    # Downscale frames at the ffmpeg pipe for cheap hashing. We pass only
    # ``max_region_dim`` (no ``phash_skip``) so the pipe downsizes without
    # dropping frames — this scanner runs its own phash on every sample.
    boundary_opts = dict(fast_opts or {})
    boundary_opts.setdefault("max_region_dim", config.SCREENSPACE_BOUNDARY_HASH_DIM)
    boundary_opts.pop("phash_skip", None)

    results: list[dict[str, Any]] = []
    prev_hash: list["imagehash.ImageHash | None"] = [None]
    last_boundary_ts: list[float | None] = [None]
    eps = config.SCREENSPACE_BOUNDARY_CONFIDENCE_EPSILON

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False

        curr_hash = compute_phash(pixels)
        if prev_hash[0] is not None:
            dist = int(curr_hash - prev_hash[0])
            within_gap = (
                last_boundary_ts[0] is not None and ts - last_boundary_ts[0] < min_gap
            )
            if dist >= threshold and not within_gap:
                conf = 1.0 if threshold <= 0 else (dist - threshold) / float(threshold)
                conf = max(eps, min(conf, 1.0))
                rd = {
                    "timestamp": round(ts, 2),
                    "distance": dist,
                    "_confidence": round(conf, 4),
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
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=boundary_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


# ---------------------------------------------------------------------------
# Multitool: per-frame evaluation and multi-factor scan
# ---------------------------------------------------------------------------
