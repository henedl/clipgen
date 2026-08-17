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

import cv2
import numpy as np

import config
import profiling
import utils
import video
from screenspace_primitives import PHash, ScanCallback, compute_phash

# Codecs where keyframe-only decode (`-skip_frame nokey`) pays off: long-GOP
# inter-coded formats. Intra-only formats (every frame is a keyframe) gain
# nothing, so they are left on the full-decode path.
_NONKEY_SKIP_CODECS = frozenset({"h264", "hevc"})


def scan_video_frames(
    video_path: str,
    region: dict[str, int] | None,
    interval_seconds: float,
    callback: ScanCallback,
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    fps: float = 0.0,
    duration: float = 0.0,
    fast_opts: dict[str, Any] | None = None,
    cv_scale: float | None = None,
    profile_kind: str = "",
) -> None:
    """Iterate through video at interval, extract region, call callback.

    The *callback* receives ``(timestamp_seconds, region_pixels)`` and may
    return ``False`` to stop early. ``region=None`` passes the full frame
    (used by template detection). *fps* / *duration*, when given, skip an
    internal metadata probe.

    *fast_opts* enables fast-scan optimizations:
    - ``phash_skip``: skip frames whose perceptual hash is unchanged
    - ``max_region_dim``: downscale extracted region to this max dimension

    *profile_kind* names the tool (``change``, ``flow``, …) so a multi-tool
    process attributes ``scan.callback.<kind>`` instead of lumping every
    analysis into one ``scan.callback`` bucket. Decode/filter stay shared.
    """
    full_frame = region is None
    # Reject zero-dimension regions early: ffmpeg's crop would emit empty frames,
    # crashing cv2.cvtColor/GaussianBlur downstream with a shape mismatch.
    if region is not None and (region.get("w", 0) <= 0 or region.get("h", 0) <= 0):
        utils.warning_print(
            f"Skipping scan: region has zero width or height ({region})"
        )
        return
    if cv_scale is None:
        cv_scale = config.SCREENSPACE_CV_RESOLUTION_SCALE

    if not _scan_via_ffmpeg_pipe(
        video_path,
        None if full_frame else region,
        interval_seconds,
        callback,
        start_seconds=start_seconds,
        end_seconds=end_seconds if end_seconds is not None else 0.0,
        fps=fps,
        duration=duration,
        fast_opts=fast_opts,
        full_frame=full_frame,
        cv_scale=cv_scale,
        profile_kind=profile_kind,
    ):
        # Raise, don't warn: a scan that examined zero frames must not return an
        # empty result list, which reads as "the detector found nothing".
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
    fps: float = 0.0,
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
        fps=fps,
        duration=duration,
        fast_opts=fast_opts,
        cv_scale=cv_scale,
        profile_kind=profile_kind,
    )


# ---------------------------------------------------------------------------
# Batch frame extraction via ffmpeg pipe
# ---------------------------------------------------------------------------


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
) -> Iterator[tuple[float, np.ndarray]]:
    """Yield ``(timestamp, frame)`` tuples from one ``-f rawvideo`` ffmpeg pipe.

    *region* crops in ffmpeg so only ROI pixels are decoded and transferred;
    *max_dim* caps the largest output dimension. Stopping iteration early (e.g.
    on cancel) is safe — the ``finally`` tears the subprocess down.

    *skip_non_keyframes* decodes keyframes only (GOP-sized savings). Timestamps
    stay accurate since ``showinfo`` reports true PTS; only temporal resolution
    drops to keyframe granularity, so the caller must gate this on keyframes
    being frequent enough (see ``_scan_via_ffmpeg_pipe``).
    """
    if not shutil.which("ffmpeg"):
        return

    # `select`, not `fps`: fps=1/N rewrites each kept frame's PTS to the output
    # slot, so the preview seeking to that PTS lands on a different source frame
    # (the long-running click-vs-preview drift). `select` preserves the original
    # PTS, which showinfo reports verbatim. Needs `-fps_mode vfr` below.
    filters = [
        f"select='isnan(prev_selected_t)+gte(t-prev_selected_t,{interval_seconds})'"
    ]

    if region:
        filters.append(f"crop={region['w']}:{region['h']}:{region['x']}:{region['y']}")
        out_w, out_h = region["w"], region["h"]
    else:
        out_w, out_h = frame_width, frame_height

    if out_w <= 0 or out_h <= 0:
        return

    # Global CV resolution scale: after the region crop, before any max_dim cap.
    # Skipped at 1.0 so default-config runs emit byte-identical argv.
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

    # showinfo last so its `pts_time:` lines map 1-to-1 onto the rawvideo frames
    # on stdout, tagging each yielded frame with its real source PTS rather than
    # a synthetic frame_idx*interval.
    filters.append("showinfo")

    cmd: list[str] = ["ffmpeg"]
    if start_seconds > 0:
        cmd += ["-ss", str(start_seconds)]
    if skip_non_keyframes:
        # Must precede -i: the decoder drops non-keyframe packets before decode.
        # The first frame can land up to one GOP past start_seconds — benign,
        # since showinfo still reports its real PTS.
        cmd += ["-skip_frame", "nokey"]
    cmd += ["-i", video_path]
    if end_seconds > start_seconds:
        cmd += ["-t", str(end_seconds - start_seconds)]
    cmd += [
        "-vf",
        ",".join(filters),
        # `select` keeps source PTS, so without `vfr` ffmpeg pads back up to the
        # source rate by duplicating each kept frame (~30x at 30 fps).
        "-fps_mode",
        "vfr",
        "-pix_fmt",
        "bgr24",
        "-f",
        "rawvideo",
        # showinfo emits at `info`; at `error` its lines never reach stderr and
        # there is no PTS data to read.
        "-loglevel",
        "info",
        "pipe:1",
    ]

    frame_size = out_w * out_h * 3
    pts_re = re.compile(rb"pts_time:(\S+)")
    pts_q: queue.Queue[float] = queue.Queue(maxsize=256)

    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    assert proc.stdout is not None  # guaranteed by stdout=PIPE
    assert proc.stderr is not None  # guaranteed by stderr=PIPE

    # Daemon thread drains stderr so its OS buffer never blocks ffmpeg, forwarding
    # each `pts_time:` as a float (seconds since the seek point).
    stop_drain = threading.Event()

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
            # Bounded put: once the consumer breaks early the queue fills, and an
            # unbounded put() would wedge this thread. Time out and re-check
            # stop_drain so it exits at teardown instead of leaking.
            while not stop_drain.is_set():
                try:
                    pts_q.put(value, timeout=0.2)
                    break
                except queue.Full:
                    continue

    drain_thread = threading.Thread(target=_drain_stderr, daemon=True)
    drain_thread.start()

    try:
        while True:
            raw = proc.stdout.read(frame_size)
            if len(raw) < frame_size:
                break
            # Read-only view on the ffmpeg pipe bytes; callers treat frames as
            # read-only and must .copy() if they retain past the next yield.
            frame = np.frombuffer(raw, dtype=np.uint8).reshape((out_h, out_w, 3))
            try:
                relative_pts = pts_q.get(timeout=5.0)
            except queue.Empty:
                # Stderr stalled or showinfo missing — bail out rather than
                # yield a frame with a misaligned synthetic timestamp.
                break
            actual_ts = start_seconds + relative_pts
            if end_seconds > 0 and actual_ts > end_seconds:
                break
            yield (actual_ts, frame)
    finally:
        stop_drain.set()
        if proc.stdout:
            proc.stdout.close()
        utils.terminate_subprocess(proc)
        drain_thread.join(timeout=1.0)


def _scan_via_ffmpeg_pipe(
    video_path: str,
    region: dict[str, int] | None,
    interval_seconds: float,
    callback: ScanCallback,
    *,
    start_seconds: float = 0.0,
    end_seconds: float = 0.0,
    fps: float = 0.0,
    duration: float = 0.0,
    fast_opts: dict[str, Any] | None = None,
    full_frame: bool = False,
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
    _prev_phash: list[PHash | None] = [None]

    pipe_region = None if full_frame else region
    # Push max_dim downscaling into ffmpeg only when phash_skip is off; hashing
    # needs the un-downscaled frame, so that path downscales in Python instead.
    pipe_max_dim = _max_dim if (not _phash_skip and _max_dim > 0) else 0

    # Keyframe-only decode, gated on fast_opts (precise/boundary scans stay on
    # full decode), a codec allowlist, the master switch, and a probe of the
    # *worst-case* keyframe gap. Probe uncertainty (None) means full decode.
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
            # `select` snaps each sample up to the next keyframe, overshooting the
            # grid when the GOP doesn't divide the interval (2s GOP, 3s interval →
            # 0,4,8 not 0,3,6). Shrinking by the worst-case gap keeps consecutive
            # samples < interval apart, so coverage is never coarser than asked —
            # at a bounded oversample that phash-skip largely absorbs.
            select_interval = max(0.0, interval_seconds - max_gap)

    # Profiling accumulates into locals and flushes once after the loop, so the
    # off-path per-frame cost is a single boolean check (see profiling.py).
    _prof = config.PROFILING
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
            region=pipe_region,
            frame_width=frame_width,
            frame_height=frame_height,
            max_dim=pipe_max_dim,
            cv_scale=cv_scale,
            skip_non_keyframes=skip_non_keyframes,
        ):
            if _prof:
                _t_dec = time.perf_counter()
                _decode_s += _t_dec - _t_last
            if _phash_skip:
                fh = compute_phash(frame)
                if _prev_phash[0] is not None and fh - _prev_phash[0] <= _phash_thresh:
                    if _prof:
                        _t_last = time.perf_counter()
                        _filter_s += _t_last - _t_dec
                        _n_skipped += 1
                    continue
                _prev_phash[0] = fh
                if _max_dim > 0:
                    rh, rw = frame.shape[:2]
                    if rh > _max_dim or rw > _max_dim:
                        sc = _max_dim / max(rh, rw)
                        frame = cv2.resize(
                            frame,
                            (int(rw * sc), int(rh * sc)),
                            interpolation=cv2.INTER_AREA,
                        )

            if _prof:
                _t_cb = time.perf_counter()
                _filter_s += _t_cb - _t_dec
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
            cb_label = (
                f"scan.callback.{profile_kind}" if profile_kind else "scan.callback"
            )
            profiling.add("scan.decode_wait", _decode_s, _seen)
            profiling.add("scan.fast_filter", _filter_s, _seen)
            profiling.add(cb_label, _cb_s, _n_frames, peak=_cb_max)
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

    cmd = [
        "ffmpeg",
        "-y",
        "-loglevel",
        config.FFMPEG_LOGLEVEL,
    ]

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


# ---------------------------------------------------------------------------
# Analysis workflows
# ---------------------------------------------------------------------------
