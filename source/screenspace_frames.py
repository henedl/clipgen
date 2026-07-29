"""Screenspace frame extraction (ffmpeg pipe + ffprobe).

The per-frame scan drivers (``scan_video_frames`` / ``scan_video_full_frames``),
the ffmpeg-pipe extractor, the timelapse command builder, and a small ffprobe
metadata helper. Imports perceptual hashing from screenspace_primitives.
"""

import queue
import re
import shutil
import subprocess
import threading
from collections.abc import Iterator
from pathlib import Path
from typing import TYPE_CHECKING, Any

import cv2
import numpy as np

if TYPE_CHECKING:
    import imagehash

import config
import utils
import video
from screenspace_primitives import ScanCallback, compute_phash

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
) -> None:
    """Iterate through video at interval, extract region, call callback.

    The *callback* receives ``(timestamp_seconds, region_pixels)`` and may
    return ``False`` to stop iteration early.  When *region* is ``None``
    the full frame is passed (used by template detection).

    When *fps* and *duration* are provided, skips internal metadata reads.
    Uses sequential frame reading (grab/retrieve) for small intervals
    to avoid expensive H.264 seeking.

    *fast_opts* enables fast-scan optimizations when provided:
    - ``phash_skip``: skip frames whose perceptual hash is unchanged
    - ``max_region_dim``: downscale extracted region to this max dimension
    """
    full_frame = region is None
    # Reject zero-dimension regions early so ffmpeg's crop filter doesn't
    # produce empty frames downstream (which would crash cv2.cvtColor /
    # cv2.GaussianBlur with an opaque shape mismatch).
    if region is not None and (region.get("w", 0) <= 0 or region.get("h", 0) <= 0):
        utils.warning_print(
            f"Skipping scan: region has zero width or height ({region})"
        )
        return
    if cv_scale is None:
        cv_scale = config.SCREENSPACE_CV_RESOLUTION_SCALE

    # Extract frames via ffmpeg pipe (faster H.264 decoding, no cv2.VideoCapture)
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
    ):
        utils.warning_print(
            f"ffmpeg pipe extraction failed for {Path(video_path).name}"
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
) -> None:
    """Like :func:`scan_video_frames` but passes the full frame (no region crop).

    Thin wrapper that calls ``scan_video_frames`` with ``region=None``.
    """
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
    )


# ---------------------------------------------------------------------------
# Batch frame extraction via ffmpeg pipe (experiment 2E)
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
    """Yield ``(timestamp, frame)`` tuples extracted via an ffmpeg pipe.

    Uses a single ffmpeg process with ``-f rawvideo`` piped to stdout,
    which is typically faster than per-frame ``cv2.VideoCapture`` seeking
    for H.264 content.

    *region* applies an ffmpeg ``crop`` filter so only the ROI pixels are
    decoded and transferred.  *max_dim* adds a ``scale`` filter to cap the
    largest output dimension (useful for fast-scan downscaling).

    *skip_non_keyframes* adds the input-level ``-skip_frame nokey`` decoder
    option so only keyframes are decoded (GOP-sized savings). The ``select``
    filter still thins them to ``interval_seconds`` apart, and ``showinfo``
    still reports each keyframe's true PTS, so timestamps stay accurate — only
    temporal resolution drops to keyframe granularity. The caller must gate this
    so it is used only when keyframes are frequent enough (see
    ``_scan_via_ffmpeg_pipe``).

    The caller can stop iteration at any time (e.g. on cancel); the
    ``finally`` block ensures the subprocess is cleaned up.
    """
    if not shutil.which("ffmpeg"):
        return

    # Determine output dimensions.
    #
    # `select` instead of `fps`: at interval N, fps=1/N picks the *last*
    # source frame whose PTS rounds to each output slot — for 30 fps source
    # that's ~0.5 s past the slot — and assigns it the slot's PTS in the
    # output. The preview path seeks to that slot PTS and lands on a
    # different source frame, hence the long-running click-vs-preview drift.
    # `select` passes the chosen frame through with its original PTS, which
    # showinfo (below) reports verbatim. Paired with `-fps_mode vfr` below to
    # stop ffmpeg from duplicating frames to fill the source rate.
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

    # Global CV resolution scale: applied after the region crop, before any
    # max_dim cap. Skipped at 1.0 so default-config runs are byte-identical
    # to pre-feature behavior.
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

    # showinfo last so its `pts_time:` lines correspond 1-to-1 with the
    # rawvideo frames pushed to stdout. The drain thread below parses these
    # and queues them so each yielded frame is tagged with its actual source
    # PTS instead of a synthetic frame_idx*interval.
    filters.append("showinfo")

    cmd: list[str] = ["ffmpeg"]
    if start_seconds > 0:
        cmd += ["-ss", str(start_seconds)]
    if skip_non_keyframes:
        # Input-level decoder option (must precede -i): the H.264/HEVC decoder
        # drops non-keyframe packets before decode, so only keyframes reach the
        # filter graph. The first yielded frame may land up to one GOP past
        # start_seconds, which is benign — showinfo still reports its real PTS.
        cmd += ["-skip_frame", "nokey"]
    cmd += ["-i", video_path]
    if end_seconds > start_seconds:
        cmd += ["-t", str(end_seconds - start_seconds)]
    cmd += [
        "-vf",
        ",".join(filters),
        # `select` keeps source PTS, so without `vfr` ffmpeg pads the output
        # back up to the source frame rate by duplicating each kept frame
        # (~30 dupes per kept frame at 30 fps). vfr emits only the actual
        # kept frames, one-to-one with the showinfo log lines.
        "-fps_mode",
        "vfr",
        "-pix_fmt",
        "bgr24",
        "-f",
        "rawvideo",
        # showinfo emits at the `info` level; with `error` its lines never
        # reach stderr and we'd have no PTS data to read.
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

    # Daemon thread drains stderr so the OS buffer never blocks ffmpeg, and
    # forwards every `pts_time:` line to the read loop as a float (seconds
    # since the seek point).
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
            # Bounded put: when the consumer stops early (break below), the
            # queue fills and an unbounded put() would wedge this thread
            # forever. Time out and re-check stop_drain so the thread exits
            # instead of leaking once we tear the subprocess down.
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
) -> bool:
    """Try to scan frames via ffmpeg pipe, calling *callback* for each.

    Returns ``True`` if the ffmpeg path succeeded (caller should skip the
    cv2 fallback), ``False`` if it failed and the caller should fall back
    to cv2-based extraction.
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
    _prev_phash: list[imagehash.ImageHash | None] = [None]

    pipe_region = None if full_frame else region
    # For the pipe, push max_dim downscaling into ffmpeg when phash_skip is off.
    # When phash_skip is on, we need the un-downscaled frame for hashing, so
    # we downscale in Python after the hash check.
    pipe_max_dim = _max_dim if (not _phash_skip and _max_dim > 0) else 0

    # Keyframe-only decode (fast-scan only): decode just keyframes when the
    # source is an inter-coded codec whose GOP is short enough. Gated on
    # fast_opts (so precise/boundary scans are untouched), a codec allowlist, the
    # master switch, and a per-video probe of the *worst-case* keyframe gap. Any
    # probe uncertainty (None) leaves it off — full decode.
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
            # `select` snaps each sample up to the next keyframe, overshooting
            # the interval grid whenever the GOP doesn't divide the interval
            # (e.g. 2s GOP, 3s interval → samples at 0,4,8 not 0,3,6). Shrinking
            # the select interval by the worst-case gap guarantees consecutive
            # samples stay < interval apart (in [interval-max_gap, interval)), so
            # coverage is never coarser than requested — at the cost of a bounded
            # (<2x for uniform GOP) oversample that phash-skip largely absorbs.
            select_interval = max(0.0, interval_seconds - max_gap)

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
            if _phash_skip:
                fh = compute_phash(frame)
                if _prev_phash[0] is not None and fh - _prev_phash[0] <= _phash_thresh:
                    continue
                _prev_phash[0] = fh
                # Downscale after phash if needed
                if _max_dim > 0:
                    rh, rw = frame.shape[:2]
                    if rh > _max_dim or rw > _max_dim:
                        sc = _max_dim / max(rh, rw)
                        frame = cv2.resize(
                            frame,
                            (int(rw * sc), int(rh * sc)),
                            interpolation=cv2.INTER_AREA,
                        )

            result = callback(ts, frame)
            if result is False:
                break

        return True  # ffmpeg pipe succeeded (even if video had 0 frames)
    except Exception as exc:
        utils.warning_print(f"ffmpeg pipe scan failed, falling back to cv2: {exc}")
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

    *sample_interval* (seconds) controls frame sampling: when > 0, only one
    frame per interval is kept before cropping and speed-up.  0 means every
    frame is used (default).

    *encoder* selects the H.264 encoder for mp4 output (see
    ``video.resolve_video_encoder``); gif output has no video encoder to pick.
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
