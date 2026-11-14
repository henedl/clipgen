"""Screenspace frame extraction (ffmpeg pipe + ffprobe)."""

import queue
import re
import shutil
import subprocess
import threading
import time
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import numpy as np

import config
import profiling
import utils
import video
from screenspace_primitives import PHash, ScanCallback, compute_phash, fit_within

# Long-GOP codecs where keyframe-only decode pays off; intra-only formats gain nothing.
_NONKEY_SKIP_CODECS = frozenset({"h264", "hevc"})
# Below this span the reference-gap probe costs more than the decode it saves.
_NONREF_MIN_SPAN_SECONDS = 120.0
# Frames read ahead of the consumer, so ffmpeg decodes while a callback runs.
_READ_AHEAD_FRAMES = 2


def scan_video_frames(
    video_path: str,
    region: dict[str, int] | None,
    interval_seconds: float,
    callback: ScanCallback,
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    duration: float = 0.0,
    fast_opts: dict[str, Any] | None = None,
    cv_scale: float | None = None,
    profile_kind: str = "",
) -> None:
    """Iterate through video at interval, extract region, call callback.

    The *callback* receives ``(timestamp_seconds, region_pixels)`` and may
    return ``False`` to stop early. ``region=None`` passes the full frame
    (used by template detection). *duration*, when given, skips an internal
    metadata probe.

    *fast_opts* enables fast-scan optimizations:
    - ``phash_skip``: skip frames whose perceptual hash is unchanged
    - ``max_region_dim``: downscale extracted region to this max dimension

    *profile_kind* names the tool (``change``, ``flow``, …) so a multi-tool
    process attributes ``scan.callback.<kind>`` instead of lumping every
    analysis into one ``scan.callback`` bucket. Decode/filter stay shared.
    """
    # Zero-size regions make ffmpeg emit empty frames that crash cv2 downstream.
    if region is not None and (region.get("w", 0) <= 0 or region.get("h", 0) <= 0):
        utils.warning_print(
            f"Skipping scan: region has zero width or height ({region})"
        )
        return
    if cv_scale is None:
        cv_scale = config.SCREENSPACE_CV_RESOLUTION_SCALE

    if not _scan_via_ffmpeg_pipe(
        video_path,
        region,
        interval_seconds,
        callback,
        start_seconds=start_seconds,
        end_seconds=end_seconds if end_seconds is not None else 0.0,
        duration=duration,
        fast_opts=fast_opts,
        cv_scale=cv_scale,
        profile_kind=profile_kind,
    ):
        # Raise: zero frames examined must not read as "detector found nothing".
        raise RuntimeError(
            f"Could not extract frames from {Path(video_path).name} — "
            "check that ffmpeg is installed and the file is readable."
        )


def scan_video_full_frames(
    video_path: str,
    interval_seconds: float,
    callback: ScanCallback,
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    duration: float = 0.0,
    fast_opts: dict[str, Any] | None = None,
    cv_scale: float | None = None,
    profile_kind: str = "",
) -> None:
    """Like :func:`scan_video_frames` but passes the full frame (no region crop)."""
    scan_video_frames(
        video_path,
        None,
        interval_seconds,
        callback,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        duration=duration,
        fast_opts=fast_opts,
        cv_scale=cv_scale,
        profile_kind=profile_kind,
    )


# ---------------------------------------------------------------------------
# Batch frame extraction via ffmpeg pipe
# ---------------------------------------------------------------------------


def _pipe_frames_command(
    video_path: str,
    interval_seconds: float,
    *,
    start_seconds: float,
    end_seconds: float,
    region: dict[str, int] | None,
    frame_width: int,
    frame_height: int,
    max_dim: int,
    cv_scale: float,
    skip_non_keyframes: bool,
    skip_non_reference: bool,
) -> tuple[list[str], int, int] | None:
    """The scan pipe's ffmpeg argv and output size; ``None`` for a zero-area frame."""

    # `select` keeps source PTS; `fps` rewrites it and drifts preview seeks.
    # Needs `-fps_mode vfr`.
    if skip_non_reference:
        cell = f"floor((t+0.000001)/{interval_seconds})"
        prev_cell = f"floor((prev_selected_t+0.000001)/{interval_seconds})"
        filters = [f"select='isnan(prev_selected_t)+gt({cell},{prev_cell})'"]
    else:
        filters = [
            f"select='isnan(prev_selected_t)+gte(t-prev_selected_t,{interval_seconds})'"
        ]

    if region:
        filters.append(f"crop={region['w']}:{region['h']}:{region['x']}:{region['y']}")
        out_w, out_h = region["w"], region["h"]
    else:
        out_w, out_h = frame_width, frame_height

    if out_w <= 0 or out_h <= 0:
        return None

    # Global CV scale after crop, before max_dim; skipped at 1.0 for identical argv.
    if cv_scale > 0 and abs(cv_scale - 1.0) > 1e-6:
        scaled_w = max(2, round(out_w * cv_scale))
        scaled_h = max(2, round(out_h * cv_scale))
        scaled_w += scaled_w % 2
        scaled_h += scaled_h % 2
        filters.append(f"scale={scaled_w}:{scaled_h}:flags=lanczos")
        out_w, out_h = scaled_w, scaled_h

    if max_dim > 0 and (out_w > max_dim or out_h > max_dim):
        scale = max_dim / max(out_w, out_h)
        out_w = int(out_w * scale)
        out_h = int(out_h * scale)
        # Ensure even dimensions for rawvideo
        out_w += out_w % 2
        out_h += out_h % 2
        filters.append(f"scale={out_w}:{out_h}")

    # showinfo last so its `pts_time:` lines map 1-to-1 onto stdout frames.
    filters.append("showinfo")

    cmd: list[str] = ["ffmpeg"]
    if start_seconds > 0:
        cmd += ["-ss", str(start_seconds)]
    if skip_non_keyframes:
        # Must precede -i. First frame may land a GOP late; showinfo still
        # reports real PTS.
        cmd += ["-skip_frame", "nokey"]
    elif skip_non_reference:
        cmd += ["-skip_frame", "noref"]
    cmd += ["-i", video_path]
    if end_seconds > start_seconds:
        cmd += ["-t", str(end_seconds - start_seconds)]
    cmd += [
        "-vf",
        ",".join(filters),
        # Without `vfr` ffmpeg duplicates each kept frame back to the source rate.
        "-fps_mode",
        "vfr",
        "-pix_fmt",
        "bgr24",
        "-f",
        "rawvideo",
        # showinfo only emits at `info` level.
        "-loglevel",
        "info",
        "pipe:1",
    ]
    return cmd, out_w, out_h


def _ffmpeg_pipe_frames(
    video_path: str,
    interval_seconds: float,
    *,
    start_seconds: float = 0.0,
    end_seconds: float = 0.0,
    region: dict[str, int] | None = None,
    frame_width: int = 0,
    frame_height: int = 0,
    max_dim: int = 0,
    cv_scale: float = 1.0,
    skip_non_keyframes: bool = False,
    skip_non_reference: bool = False,
) -> Iterator[tuple[float, np.ndarray]]:
    """Yield ``(timestamp, frame)`` tuples from one ``-f rawvideo`` ffmpeg pipe.

    *region* crops in ffmpeg so only ROI pixels are decoded and transferred;
    *max_dim* caps the largest output dimension. Stopping iteration early (e.g.
    on cancel) is safe — the ``finally`` tears the subprocess down.

    *skip_non_keyframes* decodes keyframes only (GOP-sized savings). Timestamps
    stay accurate since ``showinfo`` reports true PTS; only temporal resolution
    drops to keyframe granularity, so the caller must gate this on keyframes
    being frequent enough (see ``_scan_via_ffmpeg_pipe``).

    *skip_non_reference* drops frames nothing else is predicted from before
    decode. The frames that remain are bit-identical to a full decode; each
    sample is the first reference frame in its interval, so it lands at most
    one reference gap after the grid point. Selection is then per grid cell,
    not relative to the previous sample, so the late landings never add up
    into fewer samples. The caller gates it on the probed gap.
    """
    if not shutil.which("ffmpeg"):
        return
    built = _pipe_frames_command(
        video_path,
        interval_seconds,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        region=region,
        frame_width=frame_width,
        frame_height=frame_height,
        max_dim=max_dim,
        cv_scale=cv_scale,
        skip_non_keyframes=skip_non_keyframes,
        skip_non_reference=skip_non_reference,
    )
    if built is None:
        return
    cmd, out_w, out_h = built

    frame_size = out_w * out_h * 3
    pts_re = re.compile(rb"pts_time:(\S+)")
    pts_q: queue.Queue[float] = queue.Queue(maxsize=256)

    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    assert proc.stdout is not None  # guaranteed by stdout=PIPE
    assert proc.stderr is not None  # guaranteed by stderr=PIPE

    frames_q: queue.Queue[tuple[float, np.ndarray] | None] = queue.Queue(
        maxsize=_READ_AHEAD_FRAMES
    )
    stop_drain = threading.Event()

    def _offer(target: queue.Queue[Any], item: Any) -> bool:
        # Bounded put that gives up once the consumer has left.
        while not stop_drain.is_set():
            try:
                target.put(item, timeout=0.2)
                return True
            except queue.Full:
                continue
        return False

    # Drain stderr so its buffer never blocks ffmpeg; forward each `pts_time:`
    # as seconds.
    def _drain_stderr() -> None:
        assert proc.stderr is not None
        for line in proc.stderr:
            if stop_drain.is_set():
                break
            m = pts_re.search(line)
            if not m:
                continue
            try:
                value = float(m.group(1))
            except ValueError:
                continue
            _offer(pts_q, value)

    # Own thread: a reader waiting on the callback would stall ffmpeg.
    def _read_frames() -> None:
        assert proc.stdout is not None
        try:
            while not stop_drain.is_set():
                raw = proc.stdout.read(frame_size)
                if len(raw) < frame_size:
                    break
                # Read-only view of pipe bytes; callers must .copy() to mutate.
                frame = np.frombuffer(raw, dtype=np.uint8).reshape((out_h, out_w, 3))
                try:
                    relative_pts = pts_q.get(timeout=5.0)
                except queue.Empty:
                    # Stderr stalled or showinfo missing; a misaligned timestamp
                    # is worse than a short scan.
                    break
                actual_ts = start_seconds + relative_pts
                if end_seconds > 0 and actual_ts > end_seconds:
                    break
                if not _offer(frames_q, (actual_ts, frame)):
                    break
        except (OSError, ValueError):
            pass  # pipe closed under the reader during teardown
        finally:
            _offer(frames_q, None)

    drain_thread = threading.Thread(target=_drain_stderr, daemon=True)
    drain_thread.start()
    reader_thread = threading.Thread(target=_read_frames, daemon=True)
    reader_thread.start()

    try:
        while True:
            item = frames_q.get()
            if item is None:
                break
            yield item
    finally:
        stop_drain.set()
        # Terminate first: it unblocks a reader parked in read().
        utils.terminate_subprocess(proc)
        reader_thread.join(timeout=2.0)
        proc.stdout.close()
        drain_thread.join(timeout=1.0)


def _reference_skip_allowed(
    video_path: str, props: dict[str, Any], interval_seconds: float, span: float
) -> bool:
    """Whether this scan may decode reference frames only.

    Needs a long-GOP codec, a span long enough to repay the probe, at least
    one droppable frame between references (otherwise there is nothing to
    save), and a reference gap small against the interval, so samples stay
    within ``SCREENSPACE_REFERENCE_GAP_MARGIN`` of their grid point.
    """
    if not config.SCREENSPACE_REFERENCE_FRAMES_ONLY:
        return False
    if props.get("video_codec") not in _NONKEY_SKIP_CODECS:
        return False
    fps = props.get("fps") or 0.0
    if fps <= 0 or span < _NONREF_MIN_SPAN_SECONDS:
        return False
    gap = video.probe_max_reference_gap(video_path)
    if gap is None:
        return False
    return 1.5 / fps < gap <= interval_seconds * config.SCREENSPACE_REFERENCE_GAP_MARGIN


def _decode_plan(
    video_path: str,
    props: dict[str, Any],
    interval_seconds: float,
    span: float,
    fast_opts: dict[str, Any] | None,
) -> tuple[bool, bool, float]:
    """``(keyframes only, references only, select interval)`` for one scan.

    Keyframe-only decode is the cheapest and only for fast scans; it needs
    keyframes at least as dense as the interval. Reference-only decode is
    the fallback for every scan mode, gated on the probed reference gap.
    """
    # Keyframe-only decode: gated on fast_opts, codec allowlist, master switch,
    # worst-case keyframe gap probe.
    skip_non_keyframes = False
    select_interval = interval_seconds
    if (
        fast_opts
        and config.SCREENSPACE_FAST_SCAN_SKIP_NONKEY
        and props.get("video_codec") in _NONKEY_SKIP_CODECS
    ):
        max_gap = video.probe_max_keyframe_gap(video_path)
        if (
            max_gap is not None
            and max_gap <= interval_seconds * config.SCREENSPACE_KEYFRAME_SKIP_MARGIN
        ):
            skip_non_keyframes = True
            # `select` snaps to the next keyframe; shrink by the worst-case gap
            # to keep coverage.
            select_interval = max(0.0, interval_seconds - max_gap)

    # Reference-only decode: every scan mode, gated on the probed reference gap.
    skip_non_reference = (
        not skip_non_keyframes
        and interval_seconds > 0
        and _reference_skip_allowed(video_path, props, interval_seconds, span)
    )
    return skip_non_keyframes, skip_non_reference, select_interval


def _scan_via_ffmpeg_pipe(
    video_path: str,
    region: dict[str, int] | None,
    interval_seconds: float,
    callback: ScanCallback,
    *,
    start_seconds: float = 0.0,
    end_seconds: float = 0.0,
    duration: float = 0.0,
    fast_opts: dict[str, Any] | None = None,
    cv_scale: float = 1.0,
    profile_kind: str = "",
) -> bool:
    """Try to scan frames via ffmpeg pipe, calling *callback* for each.

    ``True`` if the ffmpeg path ran, ``False`` if it could not start at all (no
    ffmpeg on PATH, unprobeable video, zero dimensions). There is no cv2
    fallback for ``False`` — ``scan_video_frames`` raises, since a scan that
    quietly yields zero frames looks like one that legitimately found nothing.
    """
    if not shutil.which("ffmpeg"):
        return False

    props = video.probe_video_properties(video_path)
    if not props:
        return False
    frame_width = props.get("width", 0)
    frame_height = props.get("height", 0)
    if frame_width <= 0 or frame_height <= 0:
        return False

    if end_seconds <= 0:
        end_seconds = duration

    # Fast-scan state
    _phash_skip = bool(fast_opts and fast_opts.get("phash_skip"))
    _max_dim = (fast_opts or {}).get("max_region_dim", 0)
    _phash_thresh = (fast_opts or {}).get(
        "phash_threshold", config.SCREENSPACE_FAST_SCAN_PHASH_THRESHOLD
    )
    prev_phash: PHash | None = None

    # Hashing needs the un-downscaled frame, so phash_skip downscales in Python.
    pipe_max_dim = _max_dim if (not _phash_skip and _max_dim > 0) else 0

    skip_non_keyframes, skip_non_reference, select_interval = _decode_plan(
        video_path, props, interval_seconds, end_seconds - start_seconds, fast_opts
    )

    # Profiling accumulates in locals and flushes after the loop (see profiling.py).
    _prof = config.PROFILING
    if _prof:
        # A count, not a time: which decode path this scan took.
        mode = "full"
        if skip_non_keyframes:
            mode = "keyframe"
        elif skip_non_reference:
            mode = "reference"
        profiling.add(f"scan.decode_mode.{mode}", 0.0, 1)
    _cb_label = f"scan.callback.{profile_kind}" if profile_kind else "scan.callback"
    _deep = profiling.deep_profiler(_cb_label) if _prof else None
    _decode_s = _filter_s = _cb_s = 0.0
    _cb_max = 0.0
    _n_frames = _n_skipped = 0
    _t_last = time.perf_counter() if _prof else 0.0
    _t_dec = _t_cb = 0.0

    try:
        for ts, frame in _ffmpeg_pipe_frames(
            video_path,
            select_interval,
            start_seconds=start_seconds,
            end_seconds=end_seconds,
            region=region,
            frame_width=frame_width,
            frame_height=frame_height,
            max_dim=pipe_max_dim,
            cv_scale=cv_scale,
            skip_non_keyframes=skip_non_keyframes,
            skip_non_reference=skip_non_reference,
        ):
            if _prof:
                _t_dec = time.perf_counter()
                _decode_s += _t_dec - _t_last
            if _phash_skip:
                fh = compute_phash(frame)
                if prev_phash is not None and fh - prev_phash <= _phash_thresh:
                    if _prof:
                        _t_last = time.perf_counter()
                        _filter_s += _t_last - _t_dec
                        _n_skipped += 1
                    continue
                prev_phash = fh
                if _max_dim > 0:
                    frame = fit_within(frame, _max_dim)

            if _prof:
                _t_cb = time.perf_counter()
                _filter_s += _t_cb - _t_dec
            if _deep is not None and profiling.deep_enable(_deep):
                # finally: a raising callback must not leave cProfile attached to
                # the reused worker thread.
                try:
                    result = callback(ts, frame)
                finally:
                    _deep.disable()
            else:
                result = callback(ts, frame)
            if _prof:
                _t_last = time.perf_counter()
                _cb_dt = _t_last - _t_cb
                _cb_s += _cb_dt
                _cb_max = max(_cb_max, _cb_dt)
                _n_frames += 1
            if result is False:
                break

        if _prof:
            _seen = _n_frames + _n_skipped
            profiling.add("scan.decode_wait", _decode_s, _seen)
            profiling.add("scan.fast_filter", _filter_s, _seen)
            profiling.add(_cb_label, _cb_s, _n_frames, peak=_cb_max)
            profiling.scan_summary(
                Path(video_path).name,
                [
                    ("decode_wait", _decode_s, _seen),
                    ("fast_filter", _filter_s, _seen),
                    ("callback", _cb_s, _n_frames),
                ],
                kind=f"scan {profile_kind}" if profile_kind else "scan",
            )
        return True  # ffmpeg pipe succeeded (even if video had 0 frames)
    except Exception as exc:
        utils.warning_print(f"ffmpeg pipe scan failed: {exc}")
        return False


def build_timelapse_command(
    video_path: str,
    region: dict[str, int],
    speedup_factor: float,
    output_path: str,
    output_format: str = "mp4",
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    sample_interval: float = 0.0,
    encoder: str = "libx264",
) -> list[str]:
    """Construct ffmpeg argv for a cropped timelapse.

    *sample_interval* (seconds) keeps one frame per interval before crop and
    speed-up; 0 (default) uses every frame. *encoder* picks the H.264 encoder
    for mp4 (see ``video.resolve_video_encoder``); gif has none to pick.
    """
    x, y, w, h = region["x"], region["y"], region["w"], region["h"]
    filters: list[str] = []
    if sample_interval > 0:
        filters.append(f"fps=1/{sample_interval}")
    filters.append(f"crop={w}:{h}:{x}:{y}")
    filters.append(f"setpts=PTS/{speedup_factor}")
    vf = ",".join(filters)

    cmd = video.ffmpeg_cmd()

    if start_seconds > 0:
        cmd += ["-ss", str(start_seconds)]

    cmd += ["-i", video_path]

    if end_seconds is not None and end_seconds > start_seconds:
        cmd += ["-t", str(end_seconds - start_seconds)]

    cmd += ["-vf", vf, "-an"]

    if output_format == "gif":
        cmd.extend(["-loop", "0"])
    else:
        cmd.extend(video.video_encoder_args(encoder, crf=23, preset="fast"))

    cmd.append(output_path)
    return cmd


def _probe_video_meta(video_path: str) -> tuple[float, float]:
    """Return ``(fps, duration)`` via ffprobe."""
    props = video.probe_video_properties(video_path)
    if props and props.get("fps", 0) > 0 and props.get("duration", 0) > 0:
        return (props["fps"], props["duration"])
    return (0.0, 0.0)


def _resolve_scan_window(
    video_path: str, start_seconds: float, end_seconds: float | None
) -> tuple[float, float, float, float] | None:
    """Probe fps/duration and clamp the scan window.

    Returns ``(vid_fps, vid_duration, end_seconds, total_range)`` or ``None`` when
    the video is unreadable (fps <= 0), so callers early-return ``[]``. Owns the
    probe/clamp prologue shared by every scan workflow in screenspace_scans.
    """
    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return None
    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds
    return vid_fps, vid_duration, end_seconds, total_range
