# -*- coding: utf-8 -*-
"""Video processing operations for clipgen."""

import concurrent.futures
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
from collections import Counter
from collections.abc import Callable
from datetime import datetime
from pathlib import Path
from typing import Any

from icecream import ic

import config
import files
import utils

INVALID_END_TIMESTAMP = None

# Two-stage ffmpeg seek window. We pre-seek (fast, key-frame-aligned) to
# `target - FFMPEG_PRESEEK_SECONDS` and then seek the rest accurately after
# `-i`. This keeps long-video performance while landing on the exact frame
# the caller asked for, instead of the nearest preceding key-frame.
FFMPEG_PRESEEK_SECONDS = 2.0

# Caches are keyed on (resolved_path, mtime_ns) so a re-encoded or replaced
# source file naturally yields a fresh entry instead of stale data. Mirrors
# the pattern in viewer.py and pipeline.py.
_file_duration_cache: dict[tuple[str, int], int] = {}
_video_properties_cache: dict[tuple[str, int], dict[str, Any]] = {}


def _resolved_path_and_mtime(filepath: str) -> tuple[str, int] | None:
    """Return ``(resolved_path_str, mtime_ns)`` or ``None`` if the file is missing.

    Used as a cache-version source: callers key per-file caches on the tuple
    so the cache invalidates automatically when the file changes on disk.
    """
    path = Path(filepath)
    try:
        st = path.stat()
    except OSError:
        return None
    return str(path.resolve()), st.st_mtime_ns


def accurate_seek_args(timestamp_seconds: float) -> tuple[list[str], list[str]]:
    """Return ``(pre_input_args, post_input_args)`` for a frame-accurate seek.

    Splits a seek into a fast pre-input ``-ss`` near the target plus a small
    accurate post-input ``-ss`` for the residual. Callers splice the lists
    around ``-i <video>``. For ``timestamp <= FFMPEG_PRESEEK_SECONDS`` the
    pre-input list is empty and the full seek is post-input.
    """
    if timestamp_seconds <= 0:
        return [], []
    if timestamp_seconds <= FFMPEG_PRESEEK_SECONDS:
        return [], ["-ss", str(timestamp_seconds)]
    pre = timestamp_seconds - FFMPEG_PRESEEK_SECONDS
    return ["-ss", str(pre)], ["-ss", str(FFMPEG_PRESEEK_SECONDS)]


def _ffmpeg_install_guidance_lines() -> list[str]:
    """Return actionable install guidance based on the current platform."""
    platform_specific = []
    if sys.platform == "darwin":
        platform_specific = [
            "macOS: install with Homebrew: brew install ffmpeg",
        ]
    elif sys.platform.startswith("linux"):
        platform_specific = [
            "Linux (Debian/Ubuntu): sudo apt update && sudo apt install ffmpeg",
            "Linux (Fedora): sudo dnf install ffmpeg",
        ]
    elif sys.platform.startswith("win"):
        platform_specific = [
            "Windows (winget): winget install Gyan.FFmpeg",
            "Windows (chocolatey): choco install ffmpeg",
        ]
    else:
        platform_specific = [
            "Install from: https://www.ffmpeg.org/download.html",
        ]

    return platform_specific + [
        "Then verify in a new terminal:",
        "  ffmpeg -version",
        "  ffprobe -version",
    ]


def check_ffmpeg_tools_available() -> bool:
    """Verify ffmpeg and ffprobe are available in PATH at startup."""
    missing_tools = [
        tool for tool in ("ffmpeg", "ffprobe") if shutil.which(tool) is None
    ]
    if not missing_tools:
        return True

    details = [
        f"Missing command(s): {', '.join(missing_tools)}",
        "clipgen requires both ffmpeg and ffprobe to cut and inspect videos.",
    ]
    details.extend(_ffmpeg_install_guidance_lines())
    utils.error_print("Required video tools are missing from PATH.", details)
    return False


def _probe_ffmpeg_listing(listing_arg: str, target_tokens: set[str]) -> bool:
    """Run `ffmpeg -hide_banner <listing_arg>` and report whether any line's
    second whitespace-separated token is in *target_tokens*.

    Used to detect optional ffmpeg features (encoders, filters). Only the
    listing arg matters — `-encoders`, `-filters`, etc. each emit a tabular
    listing whose second column is the encoder/filter name.
    """
    try:
        result = subprocess.run(
            ["ffmpeg", "-hide_banner", listing_arg],
            capture_output=True,
            text=True,
            timeout=10,
        )
    except (OSError, subprocess.SubprocessError):
        return False
    for line in (result.stdout or "").splitlines():
        tokens = line.strip().split()
        if len(tokens) >= 2 and tokens[1] in target_tokens:
            return True
    return False


_webp_support_cache: bool | None = None
_webp_missing_warned: bool = False
_drawtext_support_cache: bool | None = None
_vp9_support_cache: bool | None = None
_vp9_missing_warned: bool = False


def check_webp_support() -> bool:
    """Return True when ffmpeg has a libwebp encoder available.

    Queries `ffmpeg -encoders` (not `-codecs`) — only the encoders listing is
    authoritative for "can ffmpeg write this format". The codecs listing
    includes the webp muxer/decoder even on builds without libwebp. Looks for
    a line starting with `libwebp` or `libwebp_anim`. Result is cached.
    """
    global _webp_support_cache
    if _webp_support_cache is None:
        _webp_support_cache = _probe_ffmpeg_listing(
            "-encoders", {"libwebp", "libwebp_anim"}
        )
    return _webp_support_cache


def check_drawtext_support() -> bool:
    """Return True when ffmpeg has the `drawtext` filter available.

    `drawtext` requires libfreetype, which is omitted from Homebrew's default
    ffmpeg 8.x build. Without it, titlecard encoding fails. Result is cached.
    """
    global _drawtext_support_cache
    if _drawtext_support_cache is None:
        _drawtext_support_cache = _probe_ffmpeg_listing("-filters", {"drawtext"})
    return _drawtext_support_cache


def check_vp9_support() -> bool:
    """Return True when ffmpeg has a libvpx-vp9 encoder available.

    Needed when GIF_FORMAT is ".webm". Same caveat as libwebp — only the
    `-encoders` listing is authoritative. Result is cached.
    """
    global _vp9_support_cache
    if _vp9_support_cache is None:
        _vp9_support_cache = _probe_ffmpeg_listing("-encoders", {"libvpx-vp9"})
    return _vp9_support_cache


def _warn_webp_unavailable_once(output_file: str) -> None:
    """Print a single clear error per session when WebP is requested but unsupported."""
    global _webp_missing_warned
    if _webp_missing_warned:
        return
    _webp_missing_warned = True
    utils.error_print(
        "WebP output requested but ffmpeg has no libwebp encoder.",
        [
            f"Tried to write: '{output_file}'",
            "Install an ffmpeg build with libwebp, or change SCREENSHOT_FORMAT/GIF_FORMAT back to .png/.jpg/.gif.",
            "Skipping all WebP outputs for this run.",
        ],
    )


def _warn_vp9_unavailable_once(output_file: str) -> None:
    """Print a single clear error per session when WebM/VP9 is requested but unsupported."""
    global _vp9_missing_warned
    if _vp9_missing_warned:
        return
    _vp9_missing_warned = True
    utils.error_print(
        "WebM output requested but ffmpeg has no libvpx-vp9 encoder.",
        [
            f"Tried to write: '{output_file}'",
            "Install an ffmpeg build with libvpx, or change GIF_FORMAT back to .gif/.webp.",
            "Skipping all WebM outputs for this run.",
        ],
    )


def run_ffmpeg_process(
    ffmpeg_command: list[str],
    *,
    input_file: str,
    output_file: str,
    os_error_message: str,
    cancel_flag: Callable[[], bool] | None = None,
    on_progress: Callable[[float], None] | None = None,
    expected_duration_sec: float | None = None,
) -> subprocess.CompletedProcess[str] | None:
    """Run an ffmpeg subprocess and wrap common OS-level failures.

    When *cancel_flag* is supplied and returns ``True`` during execution,
    the ffmpeg process is terminated and ``None`` is returned.

    When both *on_progress* and *expected_duration_sec* are supplied, ffmpeg
    is invoked with ``-progress pipe:1`` and the callback is fed fractional
    progress (0.0–1.0) as encoding advances. *expected_duration_sec* is the
    **output** (not input) duration in seconds. See screenspace.py
    ``generate_timelapse`` for the canonical pattern.
    """
    if (
        on_progress is not None
        and expected_duration_sec is not None
        and expected_duration_sec > 0
    ):
        return _run_ffmpeg_with_progress(
            ffmpeg_command,
            input_file=input_file,
            output_file=output_file,
            os_error_message=os_error_message,
            on_progress=on_progress,
            expected_duration_sec=expected_duration_sec,
            cancel_flag=cancel_flag,
        )
    try:
        if cancel_flag is not None:
            proc = subprocess.Popen(
                ffmpeg_command,
                encoding="utf-8",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            # Drain pipes via communicate() in a polling loop. Reading them
            # only after proc.poll() returns can deadlock once ffmpeg fills
            # the OS pipe buffer (~64 KB). Per Python docs, retrying after
            # TimeoutExpired does not lose output.
            while True:
                if cancel_flag():
                    utils.terminate_subprocess(proc)
                    return None
                try:
                    out, err = proc.communicate(timeout=0.5)
                except subprocess.TimeoutExpired:
                    continue
                return subprocess.CompletedProcess(
                    ffmpeg_command, proc.returncode, out, err
                )
        return subprocess.run(ffmpeg_command, encoding="utf-8", capture_output=True)
    except FileNotFoundError:
        utils.error_print(
            "ffmpeg is not installed or not found in system PATH.",
            [
                "Please install ffmpeg and ensure it's in your PATH.",
                "Download from: https://www.ffmpeg.org/download.html",
            ],
        )
        return None
    except OSError as error:
        utils.error_print(
            os_error_message,
            [
                f"Error: {error}",
                f"Working directory: '{os.getcwd()}'",
                f"Input file: '{input_file}'",
                f"Output file: '{output_file}'",
            ],
        )
        return None


def _run_ffmpeg_with_progress(
    ffmpeg_command: list[str],
    *,
    input_file: str,
    output_file: str,
    os_error_message: str,
    on_progress: Callable[[float], None],
    expected_duration_sec: float,
    cancel_flag: Callable[[], bool] | None,
) -> subprocess.CompletedProcess[str] | None:
    """Run ffmpeg with ``-progress pipe:1`` and stream fractional progress.

    stdout carries the progress key/value stream; stderr is drained on a
    background thread so the OS pipe buffer can't deadlock ffmpeg while we
    block reading stdout. Returns a CompletedProcess whose stderr field
    holds the collected ffmpeg stderr output (for error reporting on
    failure), or None on cancellation / OS-level failure.
    """
    cmd = list(ffmpeg_command) + ["-progress", "pipe:1"]
    try:
        proc = subprocess.Popen(
            cmd,
            encoding="utf-8",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except FileNotFoundError:
        utils.error_print(
            "ffmpeg is not installed or not found in system PATH.",
            [
                "Please install ffmpeg and ensure it's in your PATH.",
                "Download from: https://www.ffmpeg.org/download.html",
            ],
        )
        return None
    except OSError as error:
        utils.error_print(
            os_error_message,
            [
                f"Error: {error}",
                f"Working directory: '{os.getcwd()}'",
                f"Input file: '{input_file}'",
                f"Output file: '{output_file}'",
            ],
        )
        return None

    assert proc.stdout is not None  # guaranteed by stdout=PIPE
    assert proc.stderr is not None  # guaranteed by stderr=PIPE

    stderr_chunks: list[str] = []

    def _drain_stderr() -> None:
        # Read until EOF; without this the stderr pipe can fill its 64 KB
        # OS buffer and deadlock ffmpeg while we're blocked on stdout.
        assert proc.stderr is not None
        for chunk in proc.stderr:
            stderr_chunks.append(chunk)

    stderr_thread = threading.Thread(target=_drain_stderr, daemon=True)
    stderr_thread.start()

    expected_us = expected_duration_sec * 1_000_000.0
    on_progress(0.0)
    cancelled = False
    try:
        for line in proc.stdout:
            text = line.strip()
            if text.startswith("out_time_us="):
                try:
                    us = int(text.split("=", 1)[1])
                    on_progress(min(us / expected_us, 0.99))
                except (ValueError, ZeroDivisionError):
                    pass
            if cancel_flag is not None and cancel_flag():
                utils.terminate_subprocess(proc)
                cancelled = True
                break
    finally:
        try:
            proc.stdout.close()
        except OSError:
            pass
        proc.wait()
        stderr_thread.join(timeout=1.0)
        try:
            proc.stderr.close()
        except OSError:
            pass

    if cancelled:
        return None

    on_progress(1.0)
    return subprocess.CompletedProcess(
        ffmpeg_command, proc.returncode, "", "".join(stderr_chunks)
    )


def _add_ffmpeg_stderr(
    error_details: list[str], ffmpeg_result: subprocess.CompletedProcess[str]
) -> list[str]:
    """Append trimmed ffmpeg stderr output to an error details list when available."""
    if ffmpeg_result.stderr:
        error_details.append(f"ffmpeg error: {ffmpeg_result.stderr.strip()}")
    return error_details


def verify_output_file(output_file: str, operation_label: str) -> bool:
    """Return True when an expected ffmpeg output file exists, otherwise log an error."""
    if Path(output_file).is_file():
        return True
    utils.error_print(
        f"{operation_label} completed but output file was not created: '{output_file}'"
    )
    return False


def build_ffmpeg_cut_command(
    input_file: str,
    output_file: str,
    start_pos: str,
    duration_seconds: int,
    reencode: bool,
    audio_normalize: bool,
) -> list[str]:
    """Build ffmpeg argv for cutting a clip. Caller runs subprocess.

    Args:
        input_file: Input video path
        output_file: Output video path
        start_pos: Start timestamp
        duration_seconds: Clip duration in seconds
        reencode: If True, re-encode; if False, stream copy
        audio_normalize: If True, apply loudnorm
    Returns:
        argv list for subprocess (e.g. ['ffmpeg', '-y', ...])
    """
    base = [
        "ffmpeg",
        "-y",
        "-loglevel",
        config.FFMPEG_LOGLEVEL,
        "-ss",
        start_pos,
        "-i",
        input_file,
        "-t",
        str(duration_seconds),
    ]
    if not reencode:
        if audio_normalize:
            # loudnorm: I=-16 (target LUFS), TP=-1.5 (true peak dB), LRA=11 (loudness range)
            # -avoid_negative_ts 1: shift timestamps so output starts at 0 (avoids glitches after cut)
            return base + [
                "-c:v",
                "copy",
                "-c:a",
                "aac",
                "-af",
                "loudnorm=I=-16:TP=-1.5:LRA=11",
                "-avoid_negative_ts",
                "1",
                output_file,
            ]
        # Stream copy; -avoid_negative_ts 1 fixes timestamp issues when cutting
        return base + ["-c", "copy", "-avoid_negative_ts", "1", output_file]
    if audio_normalize:
        return base + ["-af", "loudnorm=I=-16:TP=-1.5:LRA=11", output_file]
    return base + [output_file]


_SUBTITLE_CODEC_BY_CONTAINER = {
    ".mp4": "mov_text",
    ".m4v": "mov_text",
    ".mov": "mov_text",
    ".mkv": "srt",
    ".webm": "webvtt",
}


def mux_subtitles(
    input_video: str,
    srt_path: str,
    output_video: str,
    *,
    track_title: str = "Transcript",
    track_language: str = "und",
) -> bool:
    """Stream-copy *input_video* and add *srt_path* as a soft subtitle stream.

    The subtitle codec is chosen from the output container: ``mov_text`` for
    .mp4/.mov/.m4v, ``srt`` (subrip) for .mkv, ``webvtt`` for .webm.
    Video and audio streams are stream-copied (no re-encode). Returns True
    on success, False on any validation or ffmpeg failure.
    """
    if not Path(input_video).is_file():
        utils.error_print(
            f"Input video file not found: '{input_video}'",
            [f"Expected location: {Path(input_video).resolve()}"],
        )
        return False
    if not Path(srt_path).is_file():
        utils.error_print(
            f"Subtitle file not found: '{srt_path}'",
            [f"Expected location: {Path(srt_path).resolve()}"],
        )
        return False

    suffix = Path(output_video).suffix.lower()
    codec = _SUBTITLE_CODEC_BY_CONTAINER.get(suffix)
    if codec is None:
        utils.error_print(
            f"Unsupported output container '{suffix}' for subtitle muxing.",
            [
                f"Output: '{output_video}'",
                "Supported: " + ", ".join(sorted(_SUBTITLE_CODEC_BY_CONTAINER)),
            ],
        )
        return False

    ffmpeg_command = [
        "ffmpeg",
        "-y",
        "-loglevel",
        config.FFMPEG_LOGLEVEL,
        "-i",
        input_video,
        "-i",
        srt_path,
        "-map",
        "0",
        "-map",
        "1:0",
        "-c",
        "copy",
        "-c:s",
        codec,
        "-metadata:s:s:0",
        f"language={track_language}",
        "-metadata:s:s:0",
        f"title={track_title}",
        "-disposition:s:0",
        "default",
        output_video,
    ]

    utils.verbose_print(f"Muxing subtitles into {Path(output_video).name} ({codec}).")
    if config.DEBUGGING:
        ic(ffmpeg_command)
        return False

    result = run_ffmpeg_process(
        ffmpeg_command,
        input_file=input_video,
        output_file=output_video,
        os_error_message="Failed to mux subtitles into video.",
    )
    if result is None:
        return False
    if result.returncode != 0:
        details = _add_ffmpeg_stderr(
            [
                f"Input video: '{input_video}'",
                f"Subtitle: '{srt_path}'",
                f"Output: '{output_video}'",
            ],
            result,
        )
        utils.error_print("ffmpeg returned a non-zero exit status.", details)
        return False
    return verify_output_file(output_video, "Subtitle mux")


def run_ffmpeg(
    input_file: str,
    output_file: str,
    start_pos: str,
    end_pos: str,
    reencode: bool,
    *,
    cancel_flag: Callable[[], bool] | None = None,
) -> bool:
    """Calls ffmpeg to cut a video clip. Requires ffmpeg in system PATH.

    Args:
        input_file: Path to input video file
        output_file: Path for output video file
        start_pos: Start timestamp (format: HH:MM:SS or MM:SS)
        end_pos: End timestamp (format: HH:MM:SS or MM:SS)
        reencode: If True, re-encode video; if False, use stream copy
        cancel_flag: Optional callable; when it returns True the in-flight
            ffmpeg subprocess is terminated and the function returns False.

    Returns:
        True if video was generated successfully, False otherwise.
    """
    if config.DEBUGGING:
        ic(input_file, output_file, start_pos, end_pos)
    if not Path(input_file).is_file():
        utils.error_print(
            f"Input video file not found: '{input_file}'",
            [f"Expected location: {Path(input_file).resolve()}", "Skipping this clip."],
        )
        return False

    duration = get_duration(start_pos, end_pos)
    if duration is None:
        # Error already printed by get_duration
        return False

    duration_seconds = get_file_duration(input_file)
    if duration_seconds is None:
        # Error already printed by get_file_duration
        return False

    start_seconds = utils.timestamp_to_seconds(start_pos)
    if start_seconds is not None and start_seconds >= duration_seconds:
        utils.error_print(
            f"Start timestamp ({start_pos}) is beyond video duration ({duration_seconds}s). Skipping.",
            [f"Video file: '{input_file}'"],
        )
        return False
    if start_seconds is not None and start_seconds + duration > duration_seconds:
        utils.error_print(
            f"Clip range ({start_pos} to {end_pos}) extends beyond video duration ({duration_seconds}s). Skipping.",
            [f"Video file: '{input_file}'"],
        )
        return False

    if duration < 0:
        utils.error_print(
            "Negative duration calculated for video clip. Skipping.",
            [
                f"Start: {start_pos}, End: {end_pos}, Duration: {duration}s",
                "The end timestamp must be after the start timestamp.",
            ],
        )
        return False
    if duration > duration_seconds:
        utils.error_print(
            f"Timestamp duration ({duration}s) exceeds video file length ({duration_seconds}s). Skipping.",
            [f"Start: {start_pos}, End: {end_pos}", f"Video file: '{input_file}'"],
        )
        return False
    if config.DEBUGGING:
        ic(duration, duration_seconds)
    if duration > config.MAX_CLIP_DURATION_SECONDS:
        if utils.NO_INPUT_MODE:
            utils.warning_print(
                f"Generating long clip ({duration}s, > {config.MAX_CLIP_DURATION_SECONDS}s) in non-interactive mode."
            )
        else:
            yn = utils.read_user_input(
                f"The generated video will be {duration}s ({duration // 60}m {duration % 60}s), over 10 minutes long. Generate anyway? [y/n]\n>> "
            )
            if yn != "y":
                return False

    utils.verbose_print(f"Cutting {input_file} from {start_pos} to {end_pos}.")
    if config.DEBUGGING:
        utils.debug_print(
            f"Debugging enabled, not calling ffmpeg.\n  input_file: {input_file},\n  output_file: {output_file}"
        )
        return False

    ffmpeg_command = build_ffmpeg_cut_command(
        input_file, output_file, start_pos, duration, reencode, config.AUDIO_NORMALIZE
    )
    utils.debug_print(f"ffmpeg_command is '{' '.join(ffmpeg_command)}'")
    ffmpeg_result = run_ffmpeg_process(
        ffmpeg_command,
        input_file=input_file,
        output_file=output_file,
        os_error_message="ffmpeg could not successfully run.",
        cancel_flag=cancel_flag,
    )
    if ffmpeg_result is None:
        return False

    if ffmpeg_result.returncode != 0:
        error_details = [
            f"Input: '{input_file}', Output: '{output_file}'",
            f"Timestamps: {start_pos} to {end_pos}",
        ]
        utils.error_print(
            f"ffmpeg failed with exit code {ffmpeg_result.returncode}",
            _add_ffmpeg_stderr(error_details, ffmpeg_result),
        )
        return False

    if not verify_output_file(output_file, "ffmpeg"):
        return False

    if config.MAX_FILESIZE_MB and config.MAX_FILESIZE_MB > 0:
        if not compress_to_size(
            output_file, config.MAX_FILESIZE_MB, cancel_flag=cancel_flag
        ):
            utils.warning_print(f"Could not compress '{output_file}' to target size")

    utils.verbose_print(
        f"+ Generated video '{output_file}' successfully.\n File size: {utils.format_filesize(Path(output_file).stat().st_size)}\n Expected duration: {duration} s\n"
    )
    return True


def extract_screenshot(
    input_file: str,
    output_file: str,
    timestamp: str,
    *,
    cancel_flag: Callable[[], bool] | None = None,
) -> bool:
    """Extract a single screenshot frame at the given timestamp.

    Args:
        input_file: Path to input video file
        output_file: Path for output screenshot file (.png)
        timestamp: Timestamp to capture (format: HH:MM:SS or MM:SS)
        cancel_flag: Optional callable; when it returns True the in-flight
            ffmpeg subprocess is terminated and the function returns False.

    Returns:
        True if screenshot was generated successfully, False otherwise.
    """
    if config.DEBUGGING:
        ic(input_file, output_file, timestamp)
    if output_file.lower().endswith(".webp") and not check_webp_support():
        _warn_webp_unavailable_once(output_file)
        return False
    if not Path(input_file).is_file():
        utils.error_print(
            f"Input video file not found: '{input_file}'",
            [
                f"Expected location: {Path(input_file).resolve()}",
                "Skipping this screenshot.",
            ],
        )
        return False

    file_duration = get_file_duration(input_file)
    if file_duration is not None:
        start_seconds = utils.timestamp_to_seconds(timestamp)
        if start_seconds is not None and start_seconds >= file_duration:
            utils.error_print(
                f"Screenshot timestamp ({timestamp}) is beyond video duration ({file_duration}s). Skipping.",
                [f"Video file: '{input_file}'"],
            )
            return False

    utils.verbose_print(f"Extracting screenshot from {input_file} at {timestamp}.")
    if config.DEBUGGING:
        utils.debug_print(
            f"Debugging enabled, not calling ffmpeg.\n  input_file: {input_file},\n  output_file: {output_file}"
        )
        return False

    ffmpeg_command = [
        "ffmpeg",
        "-y",
        "-loglevel",
        config.FFMPEG_LOGLEVEL,
        "-ss",
        timestamp,
        "-i",
        input_file,
        "-vframes",
        "1",
        "-q:v",
        config.FFMPEG_SCREENSHOT_QUALITY,
        output_file,
    ]
    utils.debug_print(f"ffmpeg screenshot command: {' '.join(ffmpeg_command)}")

    ffmpeg_result = run_ffmpeg_process(
        ffmpeg_command,
        input_file=input_file,
        output_file=output_file,
        os_error_message="ffmpeg could not successfully run for screenshot extraction.",
        cancel_flag=cancel_flag,
    )
    if ffmpeg_result is None:
        return False

    if ffmpeg_result.returncode != 0:
        error_details = [
            f"Input: '{input_file}', Output: '{output_file}'",
            f"Timestamp: {timestamp}",
        ]
        utils.error_print(
            f"ffmpeg screenshot failed with exit code {ffmpeg_result.returncode}",
            _add_ffmpeg_stderr(error_details, ffmpeg_result),
        )
        return False
    if not verify_output_file(output_file, "ffmpeg screenshot"):
        return False
    utils.verbose_print(
        f"+ Generated screenshot '{output_file}' successfully.\n File size: {utils.format_filesize(Path(output_file).stat().st_size)}\n"
    )
    return True


def extract_thumbnail_bytes(
    input_file: str,
    start_seconds: float,
    *,
    width: int = 200,
) -> bytes | None:
    """Extract a small JPEG thumbnail frame from a video at *start_seconds*.

    Uses two-stage seeking (fast pre-input ``-ss`` near the target, then a
    small accurate ``-ss`` after ``-i``) so the returned thumbnail matches
    the requested timestamp instead of snapping to the nearest preceding
    key-frame. Returns raw JPEG bytes on success or ``None`` on any
    failure.
    """
    if config.DEBUGGING:
        ic(input_file, start_seconds, width)
        return None

    if not Path(input_file).is_file():
        return None

    pre_seek, post_seek = accurate_seek_args(max(0.0, start_seconds))
    cmd = [
        "ffmpeg",
        "-y",
        "-loglevel",
        config.FFMPEG_LOGLEVEL,
        *pre_seek,
        "-i",
        input_file,
        *post_seek,
        "-vframes",
        "1",
        "-vf",
        f"scale={width}:-1",
        "-f",
        "image2pipe",
        "-vcodec",
        "mjpeg",
        "-q:v",
        "5",
        "pipe:1",
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, timeout=15)
    except (FileNotFoundError, OSError, subprocess.TimeoutExpired):
        return None

    if result.returncode != 0 or not result.stdout:
        return None
    return result.stdout


def extract_gif(
    input_file: str,
    output_file: str,
    timestamp: str,
    duration_seconds: int,
    *,
    cancel_flag: Callable[[], bool] | None = None,
) -> bool:
    """Extract a GIF segment starting at timestamp.

    Args:
        input_file: Path to input video file
        output_file: Path for output GIF file (.gif)
        timestamp: Start timestamp (format: HH:MM:SS or MM:SS)
        duration_seconds: GIF duration in seconds
        cancel_flag: Optional callable; when it returns True the in-flight
            ffmpeg subprocess is terminated and the function returns False.

    Returns:
        True if GIF was generated successfully, False otherwise.
    """
    if config.DEBUGGING:
        ic(input_file, output_file, timestamp, duration_seconds)
    if output_file.lower().endswith(".webp") and not check_webp_support():
        _warn_webp_unavailable_once(output_file)
        return False
    if output_file.lower().endswith(".webm") and not check_vp9_support():
        _warn_vp9_unavailable_once(output_file)
        return False
    if not Path(input_file).is_file():
        utils.error_print(
            f"Input video file not found: '{input_file}'",
            [f"Expected location: {Path(input_file).resolve()}", "Skipping this GIF."],
        )
        return False
    if duration_seconds <= 0:
        utils.error_print(
            f"Invalid GIF duration: {duration_seconds}",
            ["Duration must be greater than 0 seconds."],
        )
        return False

    file_duration = get_file_duration(input_file)
    if file_duration is not None:
        start_seconds = utils.timestamp_to_seconds(timestamp)
        if start_seconds is not None and start_seconds >= file_duration:
            utils.error_print(
                f"GIF start timestamp ({timestamp}) is beyond video duration ({file_duration}s). Skipping.",
                [f"Video file: '{input_file}'"],
            )
            return False
        if (
            start_seconds is not None
            and start_seconds + duration_seconds > file_duration
        ):
            utils.error_print(
                f"GIF range ({timestamp} + {duration_seconds}s) extends beyond video duration ({file_duration}s). Skipping.",
                [f"Video file: '{input_file}'"],
            )
            return False

    utils.verbose_print(
        f"Extracting GIF from {input_file} at {timestamp} ({duration_seconds}s)."
    )
    if config.DEBUGGING:
        utils.debug_print(
            f"Debugging enabled, not calling ffmpeg.\n  input_file: {input_file},\n  output_file: {output_file}"
        )
        return False

    out_lower = output_file.lower()
    is_webm = out_lower.endswith(".webm")
    is_webp = out_lower.endswith(".webp")

    ffmpeg_command = [
        "ffmpeg",
        "-y",
        "-loglevel",
        config.FFMPEG_LOGLEVEL,
        "-ss",
        timestamp,
        "-t",
        str(duration_seconds),
        "-i",
        input_file,
        "-vf",
        f"fps={config.GIF_FPS},scale={config.GIF_SCALE_WIDTH}:-1:flags=lanczos",
    ]
    if is_webm:
        # Silent VP9 loop; the loop is controlled by the <video loop> attribute
        # in the viewer, not by the container. -an strips audio.
        ffmpeg_command += [
            "-c:v",
            "libvpx-vp9",
            "-b:v",
            "0",
            "-crf",
            "32",
            "-row-mt",
            "1",
            "-an",
        ]
    else:
        ffmpeg_command += ["-loop", "0"]
        if is_webp:
            ffmpeg_command += ["-quality", str(config.WEBP_QUALITY)]
    ffmpeg_command.append(output_file)
    utils.debug_print(f"ffmpeg gif command: {' '.join(ffmpeg_command)}")

    ffmpeg_result = run_ffmpeg_process(
        ffmpeg_command,
        input_file=input_file,
        output_file=output_file,
        os_error_message="ffmpeg could not successfully run for GIF extraction.",
        cancel_flag=cancel_flag,
    )
    if ffmpeg_result is None:
        return False

    if ffmpeg_result.returncode != 0:
        error_details = [
            f"Input: '{input_file}', Output: '{output_file}'",
            f"Timestamp: {timestamp}, Duration: {duration_seconds}s",
        ]
        utils.error_print(
            f"ffmpeg GIF extraction failed with exit code {ffmpeg_result.returncode}",
            _add_ffmpeg_stderr(error_details, ffmpeg_result),
        )
        return False
    if not verify_output_file(output_file, "ffmpeg GIF extraction"):
        return False
    utils.verbose_print(
        f"+ Generated GIF '{output_file}' successfully.\n File size: {utils.format_filesize(Path(output_file).stat().st_size)}\n"
    )
    return True


def _probe_duration_seconds_ffprobe_format(filepath: str) -> int | None:
    """Duration-only ffprobe read for fallback paths; emits detailed errors."""
    probe_command = [
        "ffprobe",
        "-v",
        "error",
        "-show_entries",
        "format=duration",
        "-of",
        "default=noprint_wrappers=1:nokey=1",
        filepath,
    ]
    utils.debug_print(f"probe_command is {' '.join(probe_command)}")

    try:
        duration_seconds = float(
            subprocess.check_output(probe_command, encoding="utf-8")
        )
        return round(duration_seconds)
    except FileNotFoundError:
        utils.error_print(
            "ffprobe is not installed or not found in system PATH.",
            [
                "Please install ffmpeg (which includes ffprobe) and ensure it's in your PATH.",
                "Download from: https://www.ffmpeg.org/download.html",
            ],
        )
        return None
    except subprocess.CalledProcessError as e:
        utils.error_print(
            f"ffprobe failed to read video file: '{filepath}'",
            [
                f"ffprobe exit code: {e.returncode}",
                "The file may be corrupted, not a valid video, or in an unsupported format.",
            ],
        )
        return None
    except ValueError as e:
        utils.error_print(
            f"Could not parse duration from video file: '{filepath}'",
            [f"ffprobe returned unexpected output. Error: {e}"],
        )
        return None


def get_file_duration(filepath: str) -> int | None:
    """Calls ffprobe to get duration of video container.

    Reuses ``probe_video_properties`` when possible so one file is not probed
    twice for duration vs stream metadata.

    Args:
        filepath: Path to video file

    Returns:
        The duration in seconds, or None if the file cannot be probed.
    """
    key = _resolved_path_and_mtime(filepath)
    if key is None:
        utils.error_print(
            f"Video file not found: '{filepath}'",
            [
                f"Expected location: {Path(filepath).resolve()}",
                "Please ensure the video file exists in the configured input directory or working directory.",
            ],
        )
        return None
    cached_dur = _file_duration_cache.get(key)
    if cached_dur is not None:
        return cached_dur

    cached_props = _video_properties_cache.get(key)
    if cached_props is not None:
        dur_f = float(cached_props.get("duration") or 0)
        if dur_f > 0:
            rounded = round(dur_f)
            _file_duration_cache[key] = rounded
            return rounded

    probed = probe_video_properties(filepath)
    if probed is not None:
        dur_f = float(probed.get("duration") or 0)
        if dur_f > 0:
            rounded = round(dur_f)
            _file_duration_cache[key] = rounded
            return rounded

    dur = _probe_duration_seconds_ffprobe_format(filepath)
    if dur is not None:
        _file_duration_cache[key] = dur
    return dur


def build_source_timeline(paths: list[str]) -> list[tuple[str, int, int]] | None:
    """Build a concatenated-timeline view of multiple source videos.

    Given an ordered list of source-video paths that form one continuous
    recording (a participant whose session spans several files), probe each
    duration and return ``[(path, duration, cumulative_start), ...]`` where
    ``cumulative_start`` is the sum of all preceding durations (the first entry
    starts at 0). A global timestamp is mapped into a sub-video by walking these
    ranges (see ``utils.map_global_to_segment``).

    Returns ``None`` if any file's duration cannot be probed — the caller should
    skip the clip rather than guess offsets. Only called for 2+ paths; the
    single-video path never probes durations.
    """
    timeline: list[tuple[str, int, int]] = []
    cumulative = 0
    for path in paths:
        duration = get_file_duration(path)
        if duration is None:
            return None
        timeline.append((path, duration, cumulative))
        cumulative += duration
    return timeline


def timeline_or_none(paths: list[str]) -> list[tuple[str, int, int]] | None:
    """Build a source timeline for 2+ paths, else None (single-video fast path).

    The single guard that preserves the no-extra-ffprobe contract: callers pass a
    participant's ordered paths and only multi-video participants get a probed
    timeline; a single-video participant returns None and keeps the original
    single-file code path.
    """
    return build_source_timeline(paths) if len(paths) >= 2 else None


def probe_video_properties(filepath: str) -> dict[str, Any] | None:
    """Probe video file for stream properties (resolution, codecs, timing).

    Returns:
        Dict with 'width' (int), 'height' (int), 'video_codec' (str),
        'audio_codec' (str or None if no audio stream),
        'fps' (float, 0.0 if unknown), 'duration' (float seconds, 0.0 if unknown),
        'nb_frames' (int, 0 if unknown),
        or None if probe fails.
    """
    if config.DEBUGGING:
        result = {
            "width": 1920,
            "height": 1080,
            "video_codec": "h264",
            "audio_codec": "aac",
            "fps": 30.0,
            "duration": 300.0,
            "nb_frames": 9000,
        }
        # In DEBUGGING mode the file may not exist on disk; fall back to a
        # synthetic key so callers still get a cached result.
        key = _resolved_path_and_mtime(filepath) or (str(Path(filepath).resolve()), 0)
        _video_properties_cache[key] = result
        _file_duration_cache[key] = round(result["duration"])
        return result

    key = _resolved_path_and_mtime(filepath)
    if key is None:
        return None
    if key in _video_properties_cache:
        return _video_properties_cache[key]

    probe_command = [
        "ffprobe",
        "-v",
        "error",
        "-show_entries",
        "stream=width,height,codec_name,codec_type,r_frame_rate,nb_frames",
        "-show_entries",
        "format=duration",
        "-of",
        "json",
        filepath,
    ]
    try:
        raw = subprocess.check_output(probe_command, encoding="utf-8")
    except (subprocess.CalledProcessError, FileNotFoundError, OSError):
        return None

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return None

    streams = data.get("streams", [])
    width = height = 0
    video_codec: str | None = None
    audio_codec: str | None = None
    fps = 0.0
    nb_frames = 0
    for stream in streams:
        codec_type = stream.get("codec_type", "")
        if codec_type == "video" and video_codec is None:
            width = int(stream.get("width", 0))
            height = int(stream.get("height", 0))
            video_codec = stream.get("codec_name")
            # Parse r_frame_rate (e.g. "30/1", "30000/1001")
            rfr = stream.get("r_frame_rate", "")
            if "/" in rfr:
                parts = rfr.split("/")
                try:
                    num, den = float(parts[0]), float(parts[1])
                    fps = num / den if den > 0 else 0.0
                except (ValueError, IndexError):
                    pass
            try:
                nb_frames = int(stream.get("nb_frames") or 0)
            except (ValueError, TypeError):
                nb_frames = 0
        elif codec_type == "audio" and audio_codec is None:
            audio_codec = stream.get("codec_name")

    if not video_codec or width <= 0 or height <= 0:
        return None

    # Duration from format-level metadata (more reliable than stream-level)
    fmt_duration = 0.0
    fmt = data.get("format", {})
    try:
        fmt_duration = float(fmt.get("duration", 0))
    except (ValueError, TypeError):
        pass
    # Fallback: compute from frame count and fps
    if fmt_duration <= 0 and nb_frames > 0 and fps > 0:
        fmt_duration = nb_frames / fps

    result = {
        "width": width,
        "height": height,
        "video_codec": video_codec,
        "audio_codec": audio_codec,
        "fps": fps,
        "duration": fmt_duration,
        "nb_frames": nb_frames,
    }
    _video_properties_cache[key] = result
    if fmt_duration > 0:
        _file_duration_cache[key] = round(fmt_duration)
    return result


def extract_frame_at_timestamp(
    video_path: str,
    timestamp_seconds: float,
) -> Any | None:
    """Extract a single video frame at the given timestamp via ffmpeg.

    Uses two-stage seeking (fast pre-input ``-ss`` near the target, then a
    small accurate ``-ss`` after ``-i``) so the returned frame is the one
    at ``timestamp_seconds`` rather than the nearest preceding key-frame.
    Returns a BGR numpy array (H x W x 3) or None if extraction fails.
    Requires ffprobe to determine resolution and ffmpeg to decode the frame.
    """
    if config.DEBUGGING:
        import numpy as np

        return np.zeros((1080, 1920, 3), dtype=np.uint8)

    props = probe_video_properties(video_path)
    if props is None or props.get("width", 0) <= 0 or props.get("height", 0) <= 0:
        return None

    width, height = props["width"], props["height"]
    pre_seek, post_seek = accurate_seek_args(max(0.0, timestamp_seconds))
    cmd = [
        "ffmpeg",
        *pre_seek,
        "-i",
        video_path,
        *post_seek,
        "-frames:v",
        "1",
        "-pix_fmt",
        "bgr24",
        "-f",
        "rawvideo",
        "-loglevel",
        "error",
        "pipe:1",
    ]
    try:
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=10,
        )
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        return None

    frame_size = width * height * 3
    if len(result.stdout) < frame_size:
        return None

    import numpy as np

    return (
        np.frombuffer(result.stdout[:frame_size], dtype=np.uint8)
        .reshape((height, width, 3))
        .copy()
    )


def get_duration(start_time: str, end_time: str | None) -> int | None:
    """Calculate the duration between two timestamps.

    Args:
        start_time: Start timestamp (format: HH:MM:SS or MM:SS)
        end_time: End timestamp (format: HH:MM:SS or MM:SS)

    Returns:
        Duration in seconds, or None if timestamps are invalid.
    """
    if config.DEBUGGING:
        ic(start_time, end_time)
    utils.debug_print(
        f"start_time is {start_time} with length {len(start_time)}, end_time is {end_time}"
    )

    if end_time is INVALID_END_TIMESTAMP:
        utils.error_print(
            f"Invalid end timestamp (derived from start: '{start_time}')",
            ["Could not calculate end time. Check the timestamp format."],
        )
        return None

    formats = (
        ["%M:%S", "%H:%M:%S"]
        if len(str(start_time)) <= config.MAX_MMSS_LENGTH
        else ["%H:%M:%S", "%M:%S"]
    )

    for time_format in formats:
        try:
            start_datetime = datetime.strptime(str(start_time), time_format)
            end_datetime = datetime.strptime(str(end_time), time_format)
            duration = int((end_datetime - start_datetime).total_seconds())
            if config.DEBUGGING:
                ic(duration)
            return duration
        except ValueError:
            continue

    utils.error_print(
        "Timestamp formatting error in get_duration().",
        [
            f"Start time: '{start_time}', End time: '{end_time}'",
            "Accepted formats: HH:MM:SS, MM:SS, or M:SS (e.g., 1:23:45, 12:34, 1:23)",
        ],
    )
    return None


def calculate_target_bitrate(
    target_size_mb: float,
    duration_seconds: int,
    audio_bitrate_kbps: int = config.AUDIO_BITRATE_KBPS,
) -> int:
    """Calculate video bitrate needed to achieve target filesize.

    Args:
        target_size_mb: Target file size in megabytes
        duration_seconds: Video duration in seconds
        audio_bitrate_kbps: Audio bitrate in kbps

    Returns:
        Target video bitrate in kbps, minimum MIN_VIDEO_BITRATE_KBPS
    """
    if duration_seconds <= 0:
        return config.MIN_VIDEO_BITRATE_KBPS

    target_size_bytes = target_size_mb * 1024 * 1024
    total_bitrate_kbps = (target_size_bytes * 8) / duration_seconds / 1000

    video_bitrate_kbps = int(total_bitrate_kbps - audio_bitrate_kbps)
    return max(video_bitrate_kbps, config.MIN_VIDEO_BITRATE_KBPS)


def compress_to_size(
    filepath: str,
    target_size_mb: float,
    *,
    cancel_flag: Callable[[], bool] | None = None,
    on_progress: Callable[[float], None] | None = None,
) -> bool:
    """Recompress video to fit within target filesize using two-pass encoding.

    Args:
        filepath: Path to the video file to compress
        target_size_mb: Maximum file size in megabytes
        cancel_flag: Optional callable; when it returns True either pass of the
            in-flight ffmpeg subprocess is terminated and the function returns False.

    Returns:
        True if compression succeeded or was unnecessary, False on error
    """
    current_size_bytes = Path(filepath).stat().st_size
    target_size_bytes = target_size_mb * 1024 * 1024
    if current_size_bytes <= target_size_bytes:
        utils.debug_print(
            f"File already within size limit: {utils.format_filesize(current_size_bytes)}"
        )
        return True

    duration = get_file_duration(filepath)
    if duration is None or duration <= 0:
        utils.error_print(
            f"Cannot compress: unable to determine duration of '{filepath}'"
        )
        return False

    target_bitrate = calculate_target_bitrate(
        target_size_mb * config.COMPRESSION_SIZE_FACTOR, duration
    )
    if target_bitrate <= config.MIN_VIDEO_BITRATE_KBPS:
        utils.warning_print(
            f"Target bitrate very low ({target_bitrate} kbps) for {duration}s video.",
            [
                f"Target size: {target_size_mb}MB, Duration: {duration}s",
                "Quality may be significantly reduced.",
            ],
        )

    utils.verbose_print(f"Compressing video to fit within {target_size_mb}MB...")
    utils.verbose_print(f"  Current size: {utils.format_filesize(current_size_bytes)}")
    utils.verbose_print(
        f"  Target bitrate: {target_bitrate} kbps (video) + {config.AUDIO_BITRATE_KBPS} kbps (audio)"
    )

    compressed_temp_path = filepath + ".temp.mp4"
    passlog_base = filepath + ".passlog"

    try:
        null_output = "/dev/null" if os.name != "nt" else "NUL"
        pass1_command = [
            "ffmpeg",
            "-y",
            "-loglevel",
            config.FFMPEG_LOGLEVEL,
            "-i",
            filepath,
            "-c:v",
            "libx264",
            "-b:v",
            f"{target_bitrate}k",
            "-pass",
            "1",
            "-passlogfile",
            passlog_base,
            "-an",
            "-f",
            "null",
            null_output,
        ]

        utils.debug_print(f"Pass 1 command: {' '.join(pass1_command)}")
        # Split the progress bar 50/50 between the two passes so the UI shows a
        # single monotonic 0→1 fill across both ffmpeg invocations.
        pass1_progress = (
            (lambda f: on_progress(f * 0.5)) if on_progress is not None else None
        )
        pass1_result = run_ffmpeg_process(
            pass1_command,
            input_file=filepath,
            output_file=null_output,
            os_error_message="ffmpeg could not successfully run during compression pass 1.",
            cancel_flag=cancel_flag,
            on_progress=pass1_progress,
            expected_duration_sec=float(duration) if duration else None,
        )
        if pass1_result is None:
            return False

        if pass1_result.returncode != 0:
            utils.error_print(
                "Compression pass 1 failed",
                [
                    pass1_result.stderr.strip()
                    if pass1_result.stderr
                    else "Unknown error"
                ],
            )
            return False

        pass2_command = [
            "ffmpeg",
            "-y",
            "-loglevel",
            config.FFMPEG_LOGLEVEL,
            "-i",
            filepath,
            "-c:v",
            "libx264",
            "-b:v",
            f"{target_bitrate}k",
            "-pass",
            "2",
            "-passlogfile",
            passlog_base,
            "-c:a",
            "aac",
            "-b:a",
            f"{config.AUDIO_BITRATE_KBPS}k",
            compressed_temp_path,
        ]

        utils.debug_print(f"Pass 2 command: {' '.join(pass2_command)}")
        pass2_progress = (
            (lambda f: on_progress(0.5 + f * 0.5)) if on_progress is not None else None
        )
        pass2_result = run_ffmpeg_process(
            pass2_command,
            input_file=filepath,
            output_file=compressed_temp_path,
            os_error_message="ffmpeg could not successfully run during compression pass 2.",
            cancel_flag=cancel_flag,
            on_progress=pass2_progress,
            expected_duration_sec=float(duration) if duration else None,
        )
        if pass2_result is None:
            return False

        if pass2_result.returncode != 0:
            utils.error_print(
                "Compression pass 2 failed",
                [
                    pass2_result.stderr.strip()
                    if pass2_result.stderr
                    else "Unknown error"
                ],
            )
            return False

        if not verify_output_file(compressed_temp_path, "Compression"):
            return False

        new_size = Path(compressed_temp_path).stat().st_size

        os.replace(compressed_temp_path, filepath)

        utils.verbose_print(
            f"  Compressed: {utils.format_filesize(current_size_bytes)} -> {utils.format_filesize(new_size)}"
        )

        if new_size > target_size_bytes:
            utils.warning_print(
                f"Compressed file still exceeds target ({utils.format_filesize(new_size)} > {target_size_mb}MB)",
                ["The video may need a higher size limit or shorter duration."],
            )

        return True

    except OSError as e:
        utils.error_print(f"Compression failed: {e}")
        return False
    finally:
        for ext in ["-0.log", "-0.log.mbtree", ""]:
            log_path = Path(passlog_base + ext)
            if log_path.exists():
                try:
                    log_path.unlink()
                except OSError as e:
                    utils.debug_print(
                        f"Could not remove passlog file '{log_path}': {e}"
                    )
        temp_path = Path(compressed_temp_path)
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError as e:
                utils.warning_print(
                    f"Could not remove temp file: {compressed_temp_path}", [str(e)]
                )


def _detect_clip_mismatches(
    clip_paths: list[str],
) -> tuple[list[dict[str, Any] | None], bool, bool]:
    """Probe clips and detect property mismatches.

    Returns:
        (properties_list, has_resolution_mismatch, has_audio_presence_mismatch)
        properties_list parallels clip_paths (None entries for failed probes).
    """
    props_list: list[dict[str, Any] | None] = [
        probe_video_properties(p) for p in clip_paths
    ]
    probed = [p for p in props_list if p is not None]
    if len(probed) < 2:
        return (props_list, False, False)

    resolutions = Counter((p["width"], p["height"]) for p in probed)
    video_codecs = Counter(p["video_codec"] for p in probed)
    has_audio = [p["audio_codec"] is not None for p in probed]

    has_resolution_mismatch = len(resolutions) > 1
    has_audio_presence_mismatch = len(set(has_audio)) > 1

    if has_resolution_mismatch:
        detail = ", ".join(
            f"{w}x{h} ({n} clip{'s' if n > 1 else ''})"
            for (w, h), n in resolutions.most_common()
        )
        utils.warning_print(f"Resolution mismatch across reel clips: {detail}.")

    if len(video_codecs) > 1:
        detail = ", ".join(
            f"{c} ({n} clip{'s' if n > 1 else ''})"
            for c, n in video_codecs.most_common()
        )
        utils.warning_print(f"Video codec mismatch across reel clips: {detail}.")

    if has_audio_presence_mismatch:
        utils.warning_print("Audio stream mismatch: some clips have no audio track.")

    return (props_list, has_resolution_mismatch, has_audio_presence_mismatch)


def _pick_target_resolution(
    props_list: list[dict[str, Any] | None],
) -> tuple[int, int]:
    """Choose target resolution from probed properties (most common, ties broken by largest)."""
    resolutions = Counter(
        (p["width"], p["height"]) for p in props_list if p is not None
    )
    max_count = max(resolutions.values())
    candidates = [res for res, cnt in resolutions.items() if cnt == max_count]
    w, h = max(candidates, key=lambda r: r[0] * r[1])
    # libx264 requires even dimensions
    return (w if w % 2 == 0 else w + 1, h if h % 2 == 0 else h + 1)


def _build_filter_complex_concat(
    clip_paths: list[str],
    props_list: list[dict[str, Any] | None],
    target_w: int,
    target_h: int,
) -> tuple[str, bool]:
    """Build a filter_complex string that scales/pads all inputs to target resolution.

    Returns:
        (filter_complex_string, has_any_audio)
    """
    filter_parts: list[str] = []
    has_any_audio = any(
        p is not None and p.get("audio_codec") is not None for p in props_list
    )

    for i, props in enumerate(props_list):
        filter_parts.append(
            f"[{i}:v]scale={target_w}:{target_h}"
            f":force_original_aspect_ratio=decrease,"
            f"pad={target_w}:{target_h}:(ow-iw)/2:(oh-ih)/2,"
            f"setsar=1[v{i}]"
        )
        if has_any_audio:
            if props is not None and props.get("audio_codec") is not None:
                filter_parts.append(f"[{i}:a]aresample=44100[a{i}]")
            else:
                dur = get_file_duration(clip_paths[i]) or 1
                filter_parts.append(
                    f"anullsrc=r=44100:cl=stereo[sil{i}];"
                    f"[sil{i}]atrim=duration={dur}[a{i}]"
                )

    n = len(props_list)
    if has_any_audio:
        concat_inputs = "".join(f"[v{i}][a{i}]" for i in range(n))
        filter_parts.append(f"{concat_inputs}concat=n={n}:v=1:a=1[outv][outa]")
    else:
        concat_inputs = "".join(f"[v{i}]" for i in range(n))
        filter_parts.append(f"{concat_inputs}concat=n={n}:v=1:a=0[outv]")

    return (";".join(filter_parts), has_any_audio)


def concatenate_clips(
    clip_paths: list[str],
    output_file: str,
    reencode_on_fail: bool = True,
    cancel_flag: Callable[[], bool] | None = None,
    on_progress: Callable[[float], None] | None = None,
) -> bool:
    """Concatenate multiple video clips into a single file.

    Probes all clips for property mismatches. When resolutions or audio presence
    differ, uses ffmpeg filter_complex to scale/pad inputs to a common resolution.
    Otherwise uses the fast concat demuxer with stream copy, falling back to
    re-encoding if stream copy fails.

    Args:
        clip_paths: List of paths to clip files (order preserved)
        output_file: Path for the concatenated output file
        reencode_on_fail: If True, retry with re-encoding when stream copy fails

    Returns:
        True if concatenation succeeded, False otherwise.
    """
    if not clip_paths:
        utils.error_print("No clips to concatenate.", ["clip_paths must not be empty."])
        return False

    for path in clip_paths:
        if not Path(path).is_file():
            utils.error_print(
                f"Clip file not found: '{path}'",
                ["Ensure all clips were generated successfully before concatenating."],
            )
            return False

    utils.standard_print(f"Concatenating {len(clip_paths)} clips into {output_file}.")
    if config.DEBUGGING:
        utils.debug_print("Debugging enabled, not calling ffmpeg for concat.")
        return False

    props_list, res_mismatch, audio_mismatch = _detect_clip_mismatches(clip_paths)
    total_duration = sum(
        float(p["duration"]) for p in props_list if p is not None and p.get("duration")
    )

    if res_mismatch or audio_mismatch:
        utils.warning_print(
            "Re-encoding all clips to produce a compatible reel (this may take longer)."
        )
        return _concatenate_filter_complex(
            clip_paths,
            props_list,
            output_file,
            cancel_flag=cancel_flag,
            on_progress=on_progress,
            expected_duration_sec=total_duration,
        )

    return _concatenate_demuxer(
        clip_paths,
        output_file,
        reencode_on_fail,
        cancel_flag=cancel_flag,
        on_progress=on_progress,
        expected_duration_sec=total_duration,
    )


def _concatenate_filter_complex(
    clip_paths: list[str],
    props_list: list[dict[str, Any] | None],
    output_file: str,
    cancel_flag: Callable[[], bool] | None = None,
    on_progress: Callable[[float], None] | None = None,
    expected_duration_sec: float | None = None,
) -> bool:
    """Concatenate clips using filter_complex (handles resolution/audio mismatches)."""
    target_w, target_h = _pick_target_resolution(props_list)
    filter_str, has_audio = _build_filter_complex_concat(
        clip_paths, props_list, target_w, target_h
    )

    ffmpeg_command = ["ffmpeg", "-y", "-loglevel", config.FFMPEG_LOGLEVEL]
    for path in clip_paths:
        ffmpeg_command.extend(["-i", str(Path(path).resolve())])
    ffmpeg_command.extend(["-filter_complex", filter_str])
    ffmpeg_command.extend(["-map", "[outv]"])
    if has_audio:
        ffmpeg_command.extend(["-map", "[outa]"])
    ffmpeg_command.extend(["-c:v", "libx264", "-c:a", "aac", output_file])

    utils.debug_print(f"ffmpeg filter_complex concat: {' '.join(ffmpeg_command)}")
    try:
        result = run_ffmpeg_process(
            ffmpeg_command,
            input_file=clip_paths[0],
            output_file=output_file,
            os_error_message="Filter-complex concatenation failed.",
            cancel_flag=cancel_flag,
            on_progress=on_progress,
            expected_duration_sec=expected_duration_sec,
        )
        if result is None or result.returncode != 0:
            error_details = [
                f"Output: '{output_file}'",
                f"Clips: {len(clip_paths)} files",
                f"Target resolution: {target_w}x{target_h}",
            ]
            if result is not None:
                error_details = _add_ffmpeg_stderr(error_details, result)
            utils.error_print("ffmpeg filter_complex concat failed.", error_details)
            return False

        if not verify_output_file(output_file, "Concat"):
            return False

        utils.standard_print(f"+ Generated reel '{output_file}' successfully.")
        return True
    except OSError as e:
        utils.error_print(f"Concatenation failed: {e}")
        return False


def _concatenate_demuxer(
    clip_paths: list[str],
    output_file: str,
    reencode_on_fail: bool,
    cancel_flag: Callable[[], bool] | None = None,
    on_progress: Callable[[float], None] | None = None,
    expected_duration_sec: float | None = None,
) -> bool:
    """Concatenate clips using concat demuxer (fast path for matching properties)."""
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".txt", delete=False, encoding="utf-8"
    ) as file_handle:
        concat_list_file = file_handle.name
        for path in clip_paths:
            abs_path = str(Path(path).resolve())
            escaped_path = abs_path.replace("'", "'\\''")
            file_handle.write(f"file '{escaped_path}'\n")

    try:
        ffmpeg_command = [
            "ffmpeg",
            "-y",
            "-loglevel",
            config.FFMPEG_LOGLEVEL,
            "-f",
            "concat",
            "-safe",
            "0",
            "-i",
            concat_list_file,
            "-c",
            "copy",
            output_file,
        ]
        utils.debug_print(f"ffmpeg concat command: {' '.join(ffmpeg_command)}")

        ffmpeg_result = run_ffmpeg_process(
            ffmpeg_command,
            input_file=concat_list_file,
            output_file=output_file,
            os_error_message="Concatenation failed.",
            cancel_flag=cancel_flag,
        )
        if ffmpeg_result is None:
            return False

        if ffmpeg_result.returncode != 0 and reencode_on_fail:
            utils.warning_print(
                "Stream copy concat failed (e.g. codec mismatch), retrying with re-encoding."
            )
            ffmpeg_command_reencode = [
                "ffmpeg",
                "-y",
                "-loglevel",
                config.FFMPEG_LOGLEVEL,
                "-f",
                "concat",
                "-safe",
                "0",
                "-i",
                concat_list_file,
                "-c:v",
                "libx264",
                "-c:a",
                "aac",
                output_file,
            ]
            ffmpeg_result = run_ffmpeg_process(
                ffmpeg_command_reencode,
                input_file=concat_list_file,
                output_file=output_file,
                os_error_message="Concatenation failed during re-encoding fallback.",
                cancel_flag=cancel_flag,
                on_progress=on_progress,
                expected_duration_sec=expected_duration_sec,
            )
            if ffmpeg_result is None:
                return False

        if ffmpeg_result.returncode != 0:
            error_details = [
                f"Output: '{output_file}'",
                f"Clips: {len(clip_paths)} files",
            ]
            utils.error_print(
                "ffmpeg concat failed.",
                _add_ffmpeg_stderr(error_details, ffmpeg_result),
            )
            return False

        if not verify_output_file(output_file, "Concat"):
            return False

        utils.standard_print(f"+ Generated reel '{output_file}' successfully.")
        return True
    except OSError as e:
        utils.error_print(f"Concatenation failed: {e}")
        return False
    finally:
        concat_path = Path(concat_list_file)
        if concat_path.exists():
            try:
                concat_path.unlink()
            except OSError as e:
                utils.debug_print(
                    f"Could not remove concat list file '{concat_list_file}': {e}"
                )


def _batch_extract_screenshots(
    input_file: str,
    timestamps: list[int],
    interval_seconds: int,
    *,
    cancel_flag: Callable[[], bool] | None = None,
) -> list[dict[str, Any]] | None:
    """Extract all gallery screenshots in a single ffmpeg pass.

    Uses fps=1/interval filter to avoid spawning one process per frame.
    Returns artifact list on success, or None to signal fallback.
    """
    ext = config.SCREENSHOT_FORMAT
    if ext.lower() == ".webp" and not check_webp_support():
        _warn_webp_unavailable_once(f"frame_*{ext}")
        return None
    tmpdir = tempfile.mkdtemp(prefix="clipgen_gallery_")
    try:
        # The fps filter samples from t=0, so for an offset grid (a multi-video
        # part aligned to the global interval) seek the input first; frame index
        # i then still maps to timestamps[i] because the grid is evenly spaced.
        start_offset = timestamps[0] if timestamps else 0
        seek_args = ["-ss", str(start_offset)] if start_offset > 0 else []
        ffmpeg_command = [
            "ffmpeg",
            "-y",
            "-loglevel",
            config.FFMPEG_LOGLEVEL,
            *seek_args,
            "-i",
            input_file,
            "-vf",
            f"fps=1/{interval_seconds}",
            "-q:v",
            config.FFMPEG_SCREENSHOT_QUALITY,
            "-f",
            "image2",
        ]
        if ext.lower() == ".webp":
            ffmpeg_command += ["-c:v", "libwebp", "-quality", str(config.WEBP_QUALITY)]
        ffmpeg_command.append(os.path.join(tmpdir, f"frame_%04d{ext}"))
        utils.debug_print(
            f"ffmpeg batch screenshot command: {' '.join(ffmpeg_command)}"
        )

        ffmpeg_result = run_ffmpeg_process(
            ffmpeg_command,
            input_file=input_file,
            output_file=os.path.join(tmpdir, f"frame_*{ext}"),
            os_error_message="ffmpeg could not run for batch screenshot extraction.",
            cancel_flag=cancel_flag,
        )
        if ffmpeg_result is None or ffmpeg_result.returncode != 0:
            return None

        artifacts: list[dict[str, Any]] = []
        for i, ts in enumerate(timestamps):
            frame_file = os.path.join(tmpdir, f"frame_{i + 1:04d}{ext}")
            if not os.path.isfile(frame_file):
                continue
            ts_str = utils.seconds_to_timestamp(ts)
            ts_safe = ts_str.replace(":", "_")
            filename = f"gallery_{ts_safe}{ext}"
            output_path = files.get_unique_filename(filename, file_format=ext)
            shutil.move(frame_file, output_path)
            artifacts.append(
                {
                    "file": Path(output_path).name,
                    "timestamp": float(ts),
                    "timestamp_formatted": ts_str,
                    "type": "screen",
                    "duration": None,
                }
            )
        return artifacts if artifacts else None
    except (OSError, ValueError, KeyError) as exc:
        utils.debug_print(f"Batch screenshot extraction failed: {exc}")
        return None
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _parallel_extract_gifs(
    input_file: str,
    timestamps: list[int],
    interval_seconds: int,
    gif_duration_seconds: int,
    duration: int,
    *,
    cancel_flag: Callable[[], bool] | None = None,
) -> list[dict[str, Any]] | None:
    """Extract gallery GIFs using parallel ffmpeg processes.

    Returns artifact list on success, or None to signal fallback.
    """
    ext = config.GIF_FORMAT
    tasks: list[tuple[str, str, str, int, float]] = []
    for ts in timestamps:
        ts_str = utils.seconds_to_timestamp(ts)
        ts_safe = ts_str.replace(":", "_")
        filename = f"gallery_{ts_safe}{ext}"
        output_path = files.get_unique_filename(filename, file_format=ext)
        gif_dur = min(gif_duration_seconds, duration - ts)
        if gif_dur <= 0:
            break
        tasks.append((input_file, output_path, ts_str, gif_dur, float(ts)))

    if not tasks:
        return None

    total = len(tasks)
    artifacts: list[dict[str, Any]] = []
    completed = 0

    try:
        with concurrent.futures.ThreadPoolExecutor(
            max_workers=config.GALLERY_PARALLEL_WORKERS,
        ) as pool:
            future_to_task = {
                pool.submit(
                    extract_gif, t[0], t[1], t[2], t[3], cancel_flag=cancel_flag
                ): t
                for t in tasks
            }
            for future in concurrent.futures.as_completed(future_to_task):
                task = future_to_task[future]
                _, output_path, ts_str, gif_dur, ts_float = task
                completed += 1
                try:
                    ok = future.result()
                except (OSError, RuntimeError) as exc:
                    utils.debug_print(f"GIF extraction task failed: {exc}")
                    ok = False
                if ok:
                    artifacts.append(
                        {
                            "file": Path(output_path).name,
                            "timestamp": ts_float,
                            "timestamp_formatted": ts_str,
                            "type": "gif",
                            "duration": gif_dur,
                        }
                    )
                else:
                    files.release_reservation(output_path)
                utils.standard_print(f"  Captured {completed}/{total} at {ts_str}")
    except (OSError, RuntimeError) as exc:
        utils.debug_print(f"Parallel GIF extraction failed: {exc}")
        return None

    artifacts.sort(key=lambda a: a["timestamp"])
    return artifacts if artifacts else None


def generate_interval_captures(
    input_file: str,
    *,
    interval_seconds: int = 10,
    output_format: str = "screen",
    gif_duration_seconds: int = 3,
    timestamps: list[int] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
) -> list[dict[str, Any]]:
    """Generate screenshots or GIFs at regular intervals throughout a video.

    Args:
        input_file: Path to the source video file
        interval_seconds: Seconds between each capture
        output_format: 'screen' for PNG screenshots or 'gif' for animated GIFs
        gif_duration_seconds: Duration of each GIF in seconds (ignored for screenshots)
        timestamps: Explicit (local) capture times in seconds. When omitted,
            defaults to ``range(0, duration, interval_seconds)``. Multi-video
            galleries pass a grid aligned to the global interval so spacing stays
            even across part boundaries; values must still be evenly spaced by
            ``interval_seconds`` (the batch screenshot path relies on it).
        cancel_flag: Optional callable; when it returns True the build stops and
            the in-flight ffmpeg encode is terminated (used by Studio's gallery
            Cancel button).

    Returns:
        List of artifact metadata dicts with file, timestamp, type, etc.
        Returns empty list on failure.
    """
    if not Path(input_file).is_file():
        utils.error_print(
            f"Video file not found: '{input_file}'",
            [f"Expected location: {Path(input_file).resolve()}"],
        )
        return []

    duration = get_file_duration(input_file)
    if duration is None or duration <= 0:
        return []

    if interval_seconds <= 0:
        utils.error_print("Interval must be a positive number of seconds.")
        return []

    ext = config.SCREENSHOT_FORMAT if output_format == "screen" else config.GIF_FORMAT
    if timestamps is None:
        timestamps = list(range(0, duration, interval_seconds))
    total = len(timestamps)
    artifacts: list[dict[str, Any]] = []

    utils.standard_print(
        f"Generating {total} {output_format}{'s' if total != 1 else ''} "
        f"at {interval_seconds}s intervals from '{Path(input_file).name}'."
    )

    if output_format == "screen" and not config.DEBUGGING and total > 1:
        batch_artifacts = _batch_extract_screenshots(
            input_file, timestamps, interval_seconds, cancel_flag=cancel_flag
        )
        if batch_artifacts:
            utils.standard_print(
                f"  Captured {len(batch_artifacts)}/{total} screenshots (batch)"
            )
            utils.info_print(
                f"Gallery complete: {len(batch_artifacts)} of {total} captures succeeded."
            )
            return batch_artifacts

    if output_format != "screen" and not config.DEBUGGING and total > 1:
        parallel_artifacts = _parallel_extract_gifs(
            input_file,
            timestamps,
            interval_seconds,
            gif_duration_seconds,
            duration,
            cancel_flag=cancel_flag,
        )
        if parallel_artifacts is not None:
            utils.info_print(
                f"Gallery complete: {len(parallel_artifacts)} of {total} captures succeeded."
            )
            return parallel_artifacts

    for i, ts in enumerate(timestamps):
        if cancel_flag and cancel_flag():
            break
        ts_str = utils.seconds_to_timestamp(ts)
        ts_safe = ts_str.replace(":", "_")
        filename = f"gallery_{ts_safe}{ext}"
        output_path = files.get_unique_filename(filename, file_format=ext)

        if output_format == "screen":
            ok = extract_screenshot(
                input_file, output_path, ts_str, cancel_flag=cancel_flag
            )
            gif_dur = None
        else:
            gif_dur = min(gif_duration_seconds, duration - ts)
            if gif_dur <= 0:
                break
            ok = extract_gif(
                input_file, output_path, ts_str, gif_dur, cancel_flag=cancel_flag
            )

        if ok:
            artifacts.append(
                {
                    "file": Path(output_path).name,
                    "timestamp": float(ts),
                    "timestamp_formatted": ts_str,
                    "type": output_format,
                    "duration": gif_dur,
                }
            )
        else:
            files.release_reservation(output_path)
        utils.standard_print(f"  Captured {i + 1}/{total} at {ts_str}")

    utils.info_print(
        f"Gallery complete: {len(artifacts)} of {total} captures succeeded."
    )
    return artifacts
