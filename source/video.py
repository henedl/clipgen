"""Video processing operations for clipgen."""

import concurrent.futures
import contextlib
import hashlib
import json
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import threading
from collections import Counter
from collections.abc import Callable, Iterator
from pathlib import Path
from typing import Any

import config
import files
import utils
import itertools

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
# Max keyframe gap (seconds) per file; None means "unknown / too sparse to
# confirm" and callers must treat that as "do not enable keyframe-only decode".
_keyframe_gap_cache: dict[tuple[str, int], float | None] = {}
# Container seekability per file; None means "shape not determined" (not an
# MP4, truncated, unreadable) and callers must stay silent rather than warn.
_container_seekability_cache: dict[tuple[str, int], dict[str, Any] | None] = {}

# Generic audio-stream handler names muxers emit by default — treated as "no
# useful name" when labelling audio tracks (fall back to language / ordinal).
_GENERIC_AUDIO_HANDLERS = frozenset(
    {"soundhandler", "core media audio", "isom", "audio"}
)

# Track-name hints for pick_speech_audio_track(). Screen recorders name their
# streams ("Participant Mic", "System Audio"), and that name is the only signal
# available without decoding, so the heuristic is purely lexical.
_SPEECH_TRACK_HINTS = [
    "participant", "participants", "interview", "interviewee", "interviewer",
    "meeting", "mic", "mics", "microphone", "mikrofon", "voice", "voices",
    "speech", "talk", "moderator", "respondent", "facilitator", "guest", "host",
    "presenter", "narration", "commentary", "headset", "lav", "lavalier",
    "boom", "call", "zoom", "teams", "webex",
    "deltagare", "intervju", "röst", "samtal",
]  # fmt: skip
# "speaker"/"speakers" is deliberately NEGATIVE: on macOS/Windows a track named
# "Speakers" is the *output* capture (an aggregate device), not a person. Reading
# it as speech would silently transcribe system audio — the exact failure this
# whole feature exists to prevent.
_NON_SPEECH_TRACK_HINTS = [
    "system", "screen", "desktop", "display", "music", "game", "output",
    "loopback", "playback", "monitor", "soundtrack", "background", "ambience",
    "effects", "sfx", "application", "app", "browser", "tab", "share", "shared",
    "speaker", "speakers", "blackhole", "soundflower", "aggregate", "mix",
    "skärm", "skarm", "musik", "dator",
]  # fmt: skip


def _hint_pattern(words: list[str]) -> re.Pattern[str]:
    """Word-boundary alternation over *words* (so "mic" misses "dynamic")."""
    return re.compile(
        r"\b(?:" + "|".join(sorted(words, key=len, reverse=True)) + r")\b"
    )


_SPEECH_TRACK_RE = _hint_pattern(_SPEECH_TRACK_HINTS)
_NON_SPEECH_TRACK_RE = _hint_pattern(_NON_SPEECH_TRACK_HINTS)


def pick_speech_audio_track(audio_tracks: list[dict[str, Any]]) -> int:
    """Index of the track most likely to carry participant speech (0 if unknown).

    Scores each track's title/handler/label against the speech vs. system-audio
    name hints above and returns the argmax, with a lowest-index tiebreak. An
    all-zero field (no track named anything meaningful) yields 0 — which is also
    faster-whisper's own default stream, so an unnamed multitrack file behaves
    exactly as it did before this function existed.

    Takes the track dicts ``probe_video_properties`` already built rather than a
    path: this is *policy*, and baking it into the ``(path, mtime_ns)``-keyed
    probe cache would serve stale picks if the hints ever become tunable.
    """
    best_index = 0
    # None (not 0) so a field of only *negatively* scored tracks still moves off
    # track 0 — "System Audio" then an unnamed track should pick the unnamed one.
    best_score: int | None = None
    for position, track in enumerate(audio_tracks):
        haystack = " ".join(
            str(track.get(key) or "") for key in ("title", "handler", "label")
        ).lower()
        score = 2 * len(set(_SPEECH_TRACK_RE.findall(haystack))) - 3 * len(
            set(_NON_SPEECH_TRACK_RE.findall(haystack))
        )
        if best_score is None or score > best_score:
            best_score = score
            best_index = int(track.get("index", position))
    return best_index


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
    return utils.install_guidance_lines(
        brew_command="brew install ffmpeg",
        linux=[
            "Linux (Debian/Ubuntu): sudo apt update && sudo apt install ffmpeg",
            "Linux (Fedora): sudo dnf install ffmpeg",
        ],
        windows=[
            "Windows (winget): winget install Gyan.FFmpeg",
            "Windows (chocolatey): choco install ffmpeg",
        ],
        download_url="https://www.ffmpeg.org/download.html",
        verify_commands=["ffmpeg -version", "ffprobe -version"],
    )


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
    if not getattr(sys, "frozen", False):
        # Source checkouts ship a script that does the whole job; a frozen
        # bundle has no repo to run it from, so only mention it when there is.
        details.append("Or, from this checkout: scripts/install-ffmpeg-ollama.sh")
    # Not error_print: this aborts startup, and a windowed launch has no console
    # to read the guidance above — the app would just quit with nothing on screen.
    utils.fatal_startup_error("Required video tools are missing from PATH.", details)
    return False


def _probe_ffmpeg_listing(listing_arg: str, target_tokens: set[str]) -> bool:
    """Run `ffmpeg -hide_banner <listing_arg>` and report whether any line's
    second whitespace-separated token is in *target_tokens*.

    Used to detect optional ffmpeg features (encoders, filters). Only the
    listing arg matters — `-encoders`, `-filters`, etc. each emit a tabular
    listing whose second column is the encoder/filter name.
    """
    try:
        # check=False throughout this module: every ffmpeg/ffprobe call inspects
        # returncode itself and turns a failure into a warning or a None return.
        # check=True would raise past that handling and lose the diagnostics.
        result = subprocess.run(
            ["ffmpeg", "-hide_banner", listing_arg],
            capture_output=True,
            text=True,
            timeout=10,
            check=False,
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
_videotoolbox_support_cache: bool | None = None
_hw_encoder_warned: bool = False
# Session-sticky: one hardware-encode failure disables the hardware encoder for
# the rest of the run, so a broken media engine costs one wasted encode, not one
# per clip. Reset only by restarting clipgen (or by tests).
_hw_encode_failed: bool = False


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


def check_videotoolbox_support() -> bool:
    """Return True when ffmpeg can encode H.264 via Apple's VideoToolbox.

    macOS only. Same caveat as libwebp — only the `-encoders` listing is
    authoritative for "can ffmpeg write this"; the codecs listing says nothing
    about whether the hardware encoder was compiled in. Result is cached.
    """
    global _videotoolbox_support_cache
    if _videotoolbox_support_cache is None:
        _videotoolbox_support_cache = (
            sys.platform == "darwin"
            and _probe_ffmpeg_listing("-encoders", {"h264_videotoolbox"})
        )
    return _videotoolbox_support_cache


def resolve_video_encoder() -> str:
    """Return the H.264 encoder to use for this re-encode: hardware or libx264.

    Honors ``config.FFMPEG_VIDEO_ENCODER``: ``"auto"`` prefers VideoToolbox when
    it is available, an explicit ``"h264_videotoolbox"`` warns once and degrades
    if unsupported, and anything else (including ``"libx264"`` and a junk value)
    means libx264. The junk case is real: ``select`` settings are coerced with
    ``str()`` and not validated against their options list
    (``server._coerce_studio_setting``), so an unknown value must degrade quietly
    rather than raise mid-encode.

    A runtime failure (``_hw_encode_failed``) is session-sticky for **both**
    hardware modes, not just ``"auto"``. On hardware that lists the encoder but
    cannot run it, honoring the explicit choice again would spend a doomed
    hardware attempt on every single encode instead of one per session.
    """
    global _hw_encoder_warned
    choice = str(getattr(config, "FFMPEG_VIDEO_ENCODER", "auto")).strip().lower()

    if choice == "h264_videotoolbox":
        if _hw_encode_failed:
            # note_hw_encode_failure already warned when the flag was set.
            return "libx264"
        if check_videotoolbox_support():
            return "h264_videotoolbox"
        if not _hw_encoder_warned:
            _hw_encoder_warned = True
            utils.warning_print(
                "Hardware encoding requested but ffmpeg has no h264_videotoolbox encoder.",
                ["Falling back to libx264 for this session."],
            )
        return "libx264"

    if choice == "auto" and not _hw_encode_failed and check_videotoolbox_support():
        return "h264_videotoolbox"
    return "libx264"


def video_encoder_args(
    encoder: str, *, crf: int | None = None, preset: str | None = None
) -> list[str]:
    """Return the ffmpeg output args selecting *encoder* at a comparable quality.

    ``crf``/``preset`` are omitted from the libx264 branch when not given, so
    every call site reproduces its existing flags exactly and forcing
    ``FFMPEG_VIDEO_ENCODER="libx264"`` yields byte-identical argv (the sites that
    passed neither were relying on libx264's own defaults, crf 23 / preset
    medium, and still are).

    VideoToolbox has no CRF; ``-q:v`` is its constant-quality control on the
    inverted 0-100 scale, so a CRF is mapped onto roughly the same visual target
    and clamped to a sane band. ``-q:v`` requires Apple Silicon (Intel Macs
    reject it — caught by ``run_ffmpeg_encode``'s runtime fallback), and
    ``-allow_sw 1`` lets the encoder drop to software internally instead of
    hard-failing when the media engine is unavailable or saturated.
    """
    if encoder == "h264_videotoolbox":
        # No CRF given means the site relied on libx264's default of 23.
        quality = max(30, min(80, 100 - 2 * (crf if crf is not None else 23)))
        return [
            "-c:v",
            "h264_videotoolbox",
            "-q:v",
            str(quality),
            "-allow_sw",
            "1",
        ]

    args = ["-c:v", "libx264"]
    if preset is not None:
        args += ["-preset", preset]
    if crf is not None:
        args += ["-crf", str(crf)]
    return args


def note_hw_encode_failure(encoder: str) -> None:
    """Record that *encoder* failed at runtime and warn about it once.

    Makes the fallback one-shot: an encoder that ffmpeg lists but cannot actually
    run (VMs, older Intel Macs) costs one wasted encode per session, not one per
    clip. Callers do the actual libx264 retry.
    """
    global _hw_encode_failed
    if _hw_encode_failed:
        return
    _hw_encode_failed = True
    utils.warning_print(
        f"Hardware encoder '{encoder}' failed; retrying with libx264.",
        [
            "Remaining encodes this session use libx264.",
            "Set FFMPEG_VIDEO_ENCODER to 'libx264' in Settings to skip this attempt.",
        ],
    )


def run_ffmpeg_encode(
    build_command: Callable[[str], list[str]],
    *,
    encoder: str,
    **kwargs: Any,
) -> subprocess.CompletedProcess[str] | None:
    """Run an encode, retrying once with libx264 if the hardware encoder fails.

    *build_command* takes an encoder name and returns the full argv, so the
    retry re-builds rather than patching args. ``**kwargs`` pass straight through
    to ``run_ffmpeg_process``. A ``None`` result (cancellation or an OS-level
    failure) is returned as-is: a cancel is not an encoder failure and must not
    burn the retry or set the session flag (see ``note_hw_encode_failure``).
    """
    result = run_ffmpeg_process(build_command(encoder), **kwargs)
    if result is None or result.returncode == 0 or encoder == "libx264":
        return result

    note_hw_encode_failure(encoder)
    return run_ffmpeg_process(build_command("libx264"), **kwargs)


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
        return subprocess.run(
            ffmpeg_command, encoding="utf-8", capture_output=True, check=False
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
        stderr_chunks.extend(proc.stderr)

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
    """Return True when an ffmpeg output file exists and is non-empty.

    Callers reserve the output path up front via ``files.get_unique_filename()``,
    which pre-creates a zero-byte placeholder. An existence-only check therefore
    always passed once reservations landed — so an ffmpeg run that exited 0 but
    wrote nothing (e.g. a degenerate span) left the empty placeholder on disk,
    counted as a successful artifact and never released. Require a non-empty file
    so callers route the empty-output case into their normal failure path (which
    releases the reservation / unlinks the placeholder).
    """
    try:
        if Path(output_file).stat().st_size > 0:
            return True
    except OSError:
        pass  # missing (FileNotFoundError) or unstattable → treat as failure
    utils.error_print(
        f"{operation_label} completed but produced an empty or missing output "
        f"file: '{output_file}'"
    )
    return False


def _finalize_ffmpeg_output(
    ffmpeg_result: subprocess.CompletedProcess[str] | None,
    output_file: str,
    *,
    error_message: str,
    error_details: list[str],
    verify_label: str,
    success_noun: str | None = None,
    success_extra: str = "",
) -> bool:
    """Shared post-run tail for single-output ffmpeg ops.

    None → OS/cancel failure (already reported). Non-zero rc → report with
    trimmed stderr. Then require a non-empty output. When *success_noun* is
    given, emit the standard success line (with filesize + optional extra).
    """
    if ffmpeg_result is None:
        return False
    if ffmpeg_result.returncode != 0:
        utils.error_print(
            f"{error_message} with exit code {ffmpeg_result.returncode}",
            _add_ffmpeg_stderr(error_details, ffmpeg_result),
        )
        return False
    if not verify_output_file(output_file, verify_label):
        return False
    if success_noun is not None:
        size = utils.format_filesize(Path(output_file).stat().st_size)
        utils.verbose_print(
            f"+ Generated {success_noun} '{output_file}' successfully.\n"
            f" File size: {size}\n{success_extra}"
        )
    return True


@contextlib.contextmanager
def _concat_list_file(clip_paths: list[str]) -> Iterator[str]:
    """Write an ffmpeg concat-demuxer list file; unlink it on exit.

    Yields the temp path; each clip is a `file '...'` line with single
    quotes escaped. Always cleaned up on block exit (best-effort).
    """
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".txt", delete=False, encoding="utf-8"
    ) as file_handle:
        concat_list_file = file_handle.name
        for path in clip_paths:
            abs_path = str(Path(path).resolve())
            escaped_path = abs_path.replace("'", "'\\''")
            file_handle.write(f"file '{escaped_path}'\n")
    try:
        yield concat_list_file
    finally:
        concat_path = Path(concat_list_file)
        if concat_path.exists():
            try:
                concat_path.unlink()
            except OSError as e:
                utils.debug_print(
                    f"Could not remove concat list file '{concat_list_file}': {e}"
                )


def build_ffmpeg_cut_command(
    input_file: str,
    output_file: str,
    start_pos: str,
    duration_seconds: int,
    reencode: bool,
    audio_normalize: bool,
    encoder: str | None = None,
) -> list[str]:
    """Build ffmpeg argv for cutting a clip. Caller runs subprocess.

    Args:
        input_file: Input video path
        output_file: Output video path
        start_pos: Start timestamp
        duration_seconds: Clip duration in seconds
        reencode: If True, re-encode; if False, stream copy
        audio_normalize: If True, apply loudnorm
        encoder: Video encoder for the re-encode branch. ``None`` or
            ``"libx264"`` adds no ``-c:v`` at all, leaving ffmpeg on its default
            (libx264 for mp4) exactly as before hardware encoding existed.
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
    encoder_args = (
        video_encoder_args(encoder)
        if encoder is not None and encoder != "libx264"
        else []
    )
    if audio_normalize:
        return (
            base
            + ["-af", "loudnorm=I=-16:TP=-1.5:LRA=11"]
            + encoder_args
            + [output_file]
        )
    return base + encoder_args + [output_file]


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
        config.debug_ic(ffmpeg_command)
        return False

    result = run_ffmpeg_process(
        ffmpeg_command,
        input_file=input_video,
        output_file=output_video,
        os_error_message="Failed to mux subtitles into video.",
    )
    return _finalize_ffmpeg_output(
        result,
        output_video,
        error_message="ffmpeg subtitle mux failed",
        error_details=[
            f"Input video: '{input_video}'",
            f"Subtitle: '{srt_path}'",
            f"Output: '{output_video}'",
        ],
        verify_label="Subtitle mux",
    )


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
        config.debug_ic(input_file, output_file, start_pos, end_pos)
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
        config.debug_ic(duration, duration_seconds)
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

    def build_command(encoder: str) -> list[str]:
        return build_ffmpeg_cut_command(
            input_file,
            output_file,
            start_pos,
            duration,
            reencode,
            config.AUDIO_NORMALIZE,
            encoder=encoder if reencode else None,
        )

    # A stream copy has no encoder to fail over, so it stays on libx264's
    # "add nothing" branch and never spends a probe on the hardware listing.
    encoder = resolve_video_encoder() if reencode else "libx264"
    utils.debug_print(f"ffmpeg_command is '{' '.join(build_command(encoder))}'")
    ffmpeg_result = run_ffmpeg_encode(
        build_command,
        encoder=encoder,
        input_file=input_file,
        output_file=output_file,
        os_error_message="ffmpeg could not successfully run.",
        cancel_flag=cancel_flag,
    )
    # NB: file-size enforcement is deliberately NOT done here. A cut is often
    # followed by a titlecard wrap or a concat that re-encodes the body, which
    # would discard any bitrate targeting applied at cut time (and waste two
    # passes). Callers apply enforce_filesize_limit() to the *final* artifact
    # after all wrapping/concat instead.
    return _finalize_ffmpeg_output(
        ffmpeg_result,
        output_file,
        error_message="ffmpeg failed",
        error_details=[
            f"Input: '{input_file}', Output: '{output_file}'",
            f"Timestamps: {start_pos} to {end_pos}",
        ],
        verify_label="ffmpeg",
        success_noun="video",
        success_extra=f" Expected duration: {duration} s\n",
    )


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
        config.debug_ic(input_file, output_file, timestamp)
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
    return _finalize_ffmpeg_output(
        ffmpeg_result,
        output_file,
        error_message="ffmpeg screenshot failed",
        error_details=[
            f"Input: '{input_file}', Output: '{output_file}'",
            f"Timestamp: {timestamp}",
        ],
        verify_label="ffmpeg screenshot",
        success_noun="screenshot",
    )


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
        config.debug_ic(input_file, start_seconds, width)
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
        result = subprocess.run(cmd, capture_output=True, timeout=15, check=False)
    except (FileNotFoundError, OSError, subprocess.TimeoutExpired):
        return None

    if result.returncode != 0 or not result.stdout:
        return None
    return result.stdout


def extract_sprite_sheet_bytes(
    input_file: str,
    start_seconds: float,
    duration_seconds: float,
    cols: int,
    rows: int,
    *,
    frame_width: int = 160,
    seek_frames: bool = False,
) -> bytes | None:
    """Extract a single tiled JPEG sprite sheet for hover scrubbing.

    Samples ``cols * rows`` frames evenly across ``[start, start + duration]``
    and lays them out left-to-right, top-to-bottom into one image. Frame ``i``
    corresponds to source time ``start + i * (duration / (cols * rows))`` — the
    frontend card scrubber maps cursor position to a frame and shifts
    ``background-position`` accordingly. Returns JPEG bytes or ``None`` on
    failure.

    ``seek_frames=True`` grabs each frame with its own fast input-seek in a
    thread pool and composites the grid with PIL, making the cost
    O(frame_count) instead of O(duration): the default single-pass ``fps``
    filter decodes the *entire* span to emit its frames, which takes seconds
    on the minutes-long spans Composer's timeline tiles cover. Studio's short
    clip sprites keep the single-pass default.
    """
    if config.DEBUGGING:
        config.debug_ic(input_file, start_seconds, duration_seconds, cols, rows)
        return None

    if not Path(input_file).is_file():
        return None

    if seek_frames:
        return _extract_sprite_sheet_seek(
            input_file, start_seconds, duration_seconds, cols, rows, frame_width
        )

    frame_count = max(1, cols * rows)
    duration = max(0.1, duration_seconds)
    cmd = [
        "ffmpeg",
        "-y",
        "-loglevel",
        config.FFMPEG_LOGLEVEL,
        "-ss",
        str(max(0.0, start_seconds)),
        "-t",
        str(duration),
        "-i",
        input_file,
        "-frames:v",
        "1",
        "-vf",
        f"fps={frame_count}/{duration},scale={frame_width}:-1,tile={cols}x{rows}",
        "-f",
        "image2pipe",
        "-vcodec",
        "mjpeg",
        "-q:v",
        "5",
        "pipe:1",
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, timeout=20, check=False)
    except (FileNotFoundError, OSError, subprocess.TimeoutExpired):
        return None

    if result.returncode != 0 or not result.stdout:
        return None
    return result.stdout


def _extract_sprite_sheet_seek(
    input_file: str,
    start_seconds: float,
    duration_seconds: float,
    cols: int,
    rows: int,
    frame_width: int,
) -> bytes | None:
    """Seek-based sprite sheet: one fast ``-ss`` grab per frame, PIL composite.

    Each frame decodes at most one GOP (keyframe seek + roll-forward), so a
    ten-minute span costs the same as a ten-second one. Frame ``i`` samples the
    *center* of its slot (``start + (i + 0.5) * step``) — indistinguishable
    from the single-pass grid for the scrubber's purposes. Frames past EOF (or
    individually failed grabs) reuse the previous frame so the grid stays
    aligned; returns ``None`` only when every grab fails.
    """
    frame_count = max(1, cols * rows)
    duration = max(0.1, duration_seconds)
    step = duration / frame_count
    start = max(0.0, start_seconds)
    times = [start + (i + 0.5) * step for i in range(frame_count)]

    def grab(ts: float) -> bytes | None:
        cmd = [
            "ffmpeg",
            "-y",
            "-loglevel",
            config.FFMPEG_LOGLEVEL,
            "-ss",
            str(ts),
            "-i",
            input_file,
            "-frames:v",
            "1",
            "-vf",
            f"scale={frame_width}:-1",
            "-f",
            "image2pipe",
            "-vcodec",
            "mjpeg",
            "-q:v",
            "5",
            "pipe:1",
        ]
        try:
            result = subprocess.run(cmd, capture_output=True, timeout=15, check=False)
        except (FileNotFoundError, OSError, subprocess.TimeoutExpired):
            return None
        if result.returncode != 0 or not result.stdout:
            return None
        return result.stdout

    workers = min(8, os.cpu_count() or 4, frame_count)
    if frame_count >= 2:
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
            grabs = list(pool.map(grab, times))
    else:
        grabs = [grab(times[0])]

    # Deferred: keep PIL off the CLI's hot import path (video.py loads at startup).
    from io import BytesIO

    from PIL import Image

    frames: list[Any] = []
    fw = fh = 0
    for data in grabs:
        if data is None:
            frames.append(None)
            continue
        try:
            img = Image.open(BytesIO(data))
            img.load()
        except OSError:
            frames.append(None)
            continue
        frames.append(img)
        fw = max(fw, img.width)
        fh = max(fh, img.height)
    if not fw or not fh:
        return None

    sheet = Image.new("RGB", (cols * fw, rows * fh), (16, 16, 16))
    last = None
    for i in range(frame_count):
        img = frames[i] if frames[i] is not None else last
        if img is None:
            continue
        last = img
        sheet.paste(img, ((i % cols) * fw, (i // cols) * fh))
    out = BytesIO()
    sheet.save(out, format="JPEG", quality=82)
    return out.getvalue()


def extract_audio_segment_bytes(
    input_file: str,
    start_seconds: float,
    duration_seconds: float,
    *,
    sample_rate: int = 22050,
) -> bytes | None:
    """Extract ``[start, start + duration]`` as mono 16-bit PCM WAV bytes.

    Downsampled to *sample_rate* and written to a temp file (not a pipe) so the
    WAV header carries a correct data-chunk size — a non-seekable pipe leaves
    that size unwritten, which some browsers' WebAudio ``decodeAudioData``
    reject. PCM WAV decodes reliably across browsers, unlike compressed audio in
    a video container. Returns WAV bytes or ``None`` on failure.
    """
    if config.DEBUGGING:
        config.debug_ic(input_file, start_seconds, duration_seconds, sample_rate)
        return None

    if not Path(input_file).is_file():
        return None

    duration = max(0.05, duration_seconds)
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".wav")
    os.close(tmp_fd)
    cmd = [
        "ffmpeg",
        "-y",
        "-loglevel",
        config.FFMPEG_LOGLEVEL,
        "-ss",
        str(max(0.0, start_seconds)),
        "-t",
        str(duration),
        "-i",
        input_file,
        "-vn",
        "-ac",
        "1",
        "-ar",
        str(sample_rate),
        "-acodec",
        "pcm_s16le",
        "-f",
        "wav",
        tmp_path,
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, timeout=20, check=False)
        if result.returncode != 0:
            return None
        data = Path(tmp_path).read_bytes()
        return data or None
    except (FileNotFoundError, OSError, subprocess.TimeoutExpired):
        return None
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


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
        config.debug_ic(input_file, output_file, timestamp, duration_seconds)
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
    return _finalize_ffmpeg_output(
        ffmpeg_result,
        output_file,
        error_message="ffmpeg GIF extraction failed",
        error_details=[
            f"Input: '{input_file}', Output: '{output_file}'",
            f"Timestamp: {timestamp}, Duration: {duration_seconds}s",
        ],
        verify_label="ffmpeg GIF extraction",
        success_noun="GIF",
    )


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
        # -1 is a sentinel recording a prior probe that couldn't determine the
        # duration, so repeat calls skip re-running the full probe chain.
        return cached_dur if cached_dur >= 0 else None

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
    _file_duration_cache[key] = dur if dur is not None else -1
    return dur


def _parallel_probe(items: list[str], probe_fn: Callable[[str], Any]) -> list[Any]:
    """Map *probe_fn* over *items* in a small thread pool, preserving input order.

    ffprobe is a subprocess, so probing is pure I/O wait and releases the GIL —
    a reel's worth of clips probes in roughly the time of the slowest single
    probe instead of the sum. Results land at their original index, which every
    caller relies on (props lists run parallel to their path list).

    No lock is needed: the probe caches (``_video_properties_cache`` /
    ``_file_duration_cache``) are plain dicts keyed by
    ``(resolved_path, mtime_ns)``, the paths in one call are distinct, and a
    duplicate concurrent probe of the same file is idempotent — both threads
    write the same value under the same key. Fewer than 2 items skips the pool.
    """
    if len(items) < 2:
        return [probe_fn(item) for item in items]

    results: list[Any] = [None] * len(items)
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(4, len(items))) as pool:
        future_to_idx = {
            pool.submit(probe_fn, item): idx for idx, item in enumerate(items)
        }
        for future in concurrent.futures.as_completed(future_to_idx):
            results[future_to_idx[future]] = future.result()
    return results


def build_source_timeline(paths: list[str]) -> list[tuple[str, int, int]] | None:
    """Build a concatenated-timeline view of multiple source videos.

    Given an ordered list of source-video paths that form one continuous
    recording (a participant whose session spans several files), probe each
    duration and return ``[(path, duration, cumulative_start), ...]`` where
    ``cumulative_start`` is the sum of all preceding durations (the first entry
    starts at 0). A global timestamp is mapped into a sub-video by walking these
    ranges (see ``utils.map_global_to_segment``).

    Durations are probed in parallel (``_parallel_probe``); only the cumulative
    fold is sequential. Returns ``None`` if any file's duration cannot be probed
    — the caller should skip the clip rather than guess offsets. Normally called
    for 2+ paths (``timeline_or_none`` keeps single-video participants from
    probing at all); a single path still takes the sequential path.
    """
    durations = _parallel_probe(paths, get_file_duration)
    if any(duration is None for duration in durations):
        return None

    timeline: list[tuple[str, int, int]] = []
    cumulative = 0
    for path, duration in zip(paths, durations):
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
        'audio_tracks' (list of per-audio-stream dicts: index/codec/channels/
        title/language/handler/label), 'audio_track_count' (int),
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
            "pix_fmt": "yuv420p",
            "audio_sample_rate": 48000,
            "audio_channels": 2,
            "audio_channel_layout": "stereo",
            "audio_tracks": [
                {
                    "index": 0,
                    "codec": "aac",
                    "channels": 2,
                    "title": "",
                    "language": "",
                    "handler": "",
                    "label": "Track 1",
                }
            ],
            "audio_track_count": 1,
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
        (
            "stream=width,height,codec_name,codec_type,r_frame_rate,nb_frames,"
            "pix_fmt,sample_rate,channels,channel_layout"
        ),
        "-show_entries",
        "stream_tags=title,language,handler_name",
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
    pix_fmt: str | None = None
    audio_sample_rate = 0
    audio_channels = 0
    audio_channel_layout: str | None = None
    audio_tracks: list[dict[str, Any]] = []
    fps = 0.0
    nb_frames = 0
    for stream in streams:
        codec_type = stream.get("codec_type", "")
        if codec_type == "video" and video_codec is None:
            width = int(stream.get("width", 0))
            height = int(stream.get("height", 0))
            video_codec = stream.get("codec_name")
            pix_fmt = stream.get("pix_fmt")
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
        elif codec_type == "audio":
            try:
                track_channels = int(stream.get("channels") or 0)
            except (ValueError, TypeError):
                track_channels = 0
            tags = stream.get("tags") or {}
            title = (tags.get("title") or "").strip()
            language = (tags.get("language") or "").strip()
            handler = (tags.get("handler_name") or "").strip()
            # Audio-relative index (0, 1, 2…) — the `a:N` selector ffmpeg needs
            # for per-track extraction (multitrack mixing lands in a follow-up).
            track_index = len(audio_tracks)
            # Prefer an explicit title, then a meaningful handler name (skip the
            # generic muxer defaults), then the language code, then an ordinal.
            meaningful_handler = (
                handler
                if handler and handler.lower() not in _GENERIC_AUDIO_HANDLERS
                else ""
            )
            # "und" (undefined) is MP4's default language tag, not a real name.
            lang_label = (
                language.upper() if language and language.lower() != "und" else ""
            )
            label = (
                title or meaningful_handler or lang_label or f"Track {track_index + 1}"
            )
            audio_tracks.append(
                {
                    "index": track_index,
                    "codec": stream.get("codec_name"),
                    "channels": track_channels,
                    "title": title,
                    "language": language,
                    "handler": handler,
                    "label": label,
                }
            )
            # Retain the first audio stream's details as the flat top-level fields
            # (backward-compatible with the ~20 existing callers).
            if audio_codec is None:
                audio_codec = stream.get("codec_name")
                try:
                    audio_sample_rate = int(stream.get("sample_rate") or 0)
                except (ValueError, TypeError):
                    audio_sample_rate = 0
                audio_channels = track_channels
                audio_channel_layout = stream.get("channel_layout")

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
        "pix_fmt": pix_fmt,
        "audio_sample_rate": audio_sample_rate,
        "audio_channels": audio_channels,
        "audio_channel_layout": audio_channel_layout,
        "audio_tracks": audio_tracks,
        "audio_track_count": len(audio_tracks),
        "fps": fps,
        "duration": fmt_duration,
        "nb_frames": nb_frames,
    }
    _video_properties_cache[key] = result
    if fmt_duration > 0:
        _file_duration_cache[key] = round(fmt_duration)
    return result


# Per-(file, index) locks for extract_audio_track. Two requests for the same
# track are routine (a re-select tears the <audio> element down and re-requests
# it; range requests can arrive on a second connection), and the Flask server is
# threaded, so without this two ffmpeg processes would race on one output path.
_audio_track_locks: dict[tuple[str, int], threading.Lock] = {}
_audio_track_locks_guard = threading.Lock()


def _audio_track_lock(resolved: str, audio_index: int) -> threading.Lock:
    key = (resolved, audio_index)
    with _audio_track_locks_guard:
        lock = _audio_track_locks.get(key)
        if lock is None:
            lock = threading.Lock()
            _audio_track_locks[key] = lock
        return lock


def _prune_stale_audio_tracks(cache_dir: Path, digest: str, keep_mtime_ns: int) -> None:
    """Drop this source's extractions from earlier mtimes.

    Every re-encode of a source yields a fresh ``mtime_ns`` filename, so without
    this the superseded tracks accumulate forever — and the directory does not
    use ``config.TEMP_ARTIFACT_PREFIX``, so the normal temp cleanup never reaps
    it either.
    """
    prefix = f"{digest}_{keep_mtime_ns}_"
    try:
        entries = list(cache_dir.glob(f"{digest}_*"))
    except OSError:
        return
    for entry in entries:
        if not entry.name.startswith(prefix):
            _delete_quietly(entry)


def extract_audio_track(filepath: str, audio_index: int) -> Path | None:
    """Demux one audio stream (``0:a:<audio_index>``) to a standalone, seekable
    ``.m4a`` for the browser's per-track mixer.

    HTML5 ``<video>`` plays only the container's default audio track, so
    independent per-track volume needs each track as its own media source.
    Extraction is demux-only (no video decode): a stream copy when the source is
    already AAC (near-instant), falling back to an AAC re-encode otherwise. The
    result is cached on disk under the OS temp dir keyed by (resolved path,
    mtime_ns, index) so a re-encoded/replaced source re-extracts and range
    requests (seeking) are served from a real file. Returns the path or ``None``.
    """
    key = _resolved_path_and_mtime(filepath)
    if key is None:
        return None
    resolved, mtime_ns = key
    cache_dir = Path(tempfile.gettempdir()) / "clipgen_audio_tracks"
    # The scheme tag (_v2) invalidates caches from before +faststart was added.
    digest = hashlib.sha1(resolved.encode("utf-8")).hexdigest()[:16]
    out_path = cache_dir / f"{digest}_{mtime_ns}_a{audio_index}_v2.m4a"
    if out_path.is_file() and out_path.stat().st_size > 0:
        return out_path
    if config.DEBUGGING:
        return None

    try:
        cache_dir.mkdir(parents=True, exist_ok=True)
    except OSError:
        return None

    with _audio_track_lock(resolved, audio_index):
        # Re-check: another thread may have finished while we waited on the lock.
        if out_path.is_file() and out_path.stat().st_size > 0:
            return out_path
        # Unique per caller so a concurrent extraction of a *different* track of
        # the same file can never share a scratch path either.
        tmp_path = (
            cache_dir
            / f"{out_path.stem}.partial.{os.getpid()}.{threading.get_ident()}.m4a"
        )
        # +faststart relocates the moov atom to the front so the browser can
        # stream and seek the track smoothly (a tail moov forces buffering
        # stalls → pops).
        base = [
            "ffmpeg",
            "-y",
            "-i",
            filepath,
            "-map",
            f"0:a:{audio_index}",
            "-vn",
            "-movflags",
            "+faststart",
        ]
        # Try a stream copy first (instant for AAC), then an AAC re-encode for
        # codecs that can't be copied into an MP4/M4A container (Opus, PCM, …).
        for codec_args in (["-c:a", "copy"], ["-c:a", "aac", "-b:a", "160k"]):
            try:
                subprocess.run(
                    base + codec_args + [str(tmp_path)],
                    check=True,
                    capture_output=True,
                    # Generous: the copy path is near-instant but the AAC
                    # re-encode fallback runs the length of a long session
                    # recording. Bounded all the same, so a wedged ffmpeg can't
                    # pin this request thread (and every waiter on the lock).
                    timeout=300,
                )
            except (
                subprocess.CalledProcessError,
                subprocess.TimeoutExpired,
                FileNotFoundError,
                OSError,
            ):
                _delete_quietly(tmp_path)
                continue
            try:
                tmp_path.replace(out_path)
            except OSError:
                _delete_quietly(tmp_path)
                return None
            _prune_stale_audio_tracks(cache_dir, digest, mtime_ns)
            return out_path
        return None


def _delete_quietly(path: Path) -> None:
    """Best-effort unlink; ignore a missing file or OS error."""
    try:
        path.unlink()
    except OSError:
        pass


def probe_max_keyframe_gap(filepath: str) -> float | None:
    """Return the largest gap (seconds) between consecutive keyframes near the start.

    Reads packet flags — no decoding — over the first
    ``config.SCREENSPACE_KEYFRAME_PROBE_SECONDS`` of video and returns the
    **maximum** interval between consecutive keyframes. The max (not the median)
    is the safe measure for gating keyframe-only decode: a single long-GOP
    stretch inside the window is a real coverage hole and must not be masked by
    surrounding short gaps.

    Returns ``None`` when the cadence can't be confirmed: fewer than two keyframes
    in the probe window (GOP longer than the window), ffprobe/parse failure, or a
    missing file. Callers treat ``None`` as "keyframes too sparse — do not enable
    keyframe-only decode". Result is cached per ``(resolved_path, mtime)``.
    """
    if config.DEBUGGING:
        # Synthetic short GOP so DEBUGGING/dev runs don't shell out to ffprobe;
        # mirrors probe_video_properties' DEBUGGING branch.
        return 1.0

    key = _resolved_path_and_mtime(filepath)
    if key is None:
        return None
    if key in _keyframe_gap_cache:
        return _keyframe_gap_cache[key]

    window = config.SCREENSPACE_KEYFRAME_PROBE_SECONDS
    probe_command = [
        "ffprobe",
        "-v",
        "error",
        "-select_streams",
        "v:0",
        "-read_intervals",
        f"%{window}",
        "-show_entries",
        "packet=pts_time,flags",
        "-of",
        "csv=print_section=0",
        filepath,
    ]
    result: float | None = None
    try:
        raw = subprocess.check_output(probe_command, encoding="utf-8")
    except (subprocess.CalledProcessError, FileNotFoundError, OSError):
        _keyframe_gap_cache[key] = None
        return None

    # Each CSV row is "pts_time,flags" (e.g. "1.000000,K__"). A keyframe packet
    # carries 'K' as the first flag char. Collect keyframe PTS in order.
    keyframe_times: list[float] = []
    for line in raw.splitlines():
        parts = line.split(",")
        if len(parts) < 2:
            continue
        pts_str, flags = parts[0], parts[1]
        if not flags or flags[0] != "K":
            continue
        try:
            keyframe_times.append(float(pts_str))
        except ValueError:
            continue

    if len(keyframe_times) >= 2:
        keyframe_times.sort()
        gaps = [b - a for a, b in itertools.pairwise(keyframe_times) if b - a > 0]
        if gaps:
            result = max(gaps)

    _keyframe_gap_cache[key] = result
    return result


# Top-level boxes scanned before giving up on finding `moov`. A fragmented MP4
# always writes `moov` before its first `moof`, and a normal MP4 writes at most
# a handful of boxes (`ftyp`/`free`/`mdat`) before or after it, so the real
# files need 2-4 iterations. The bound only exists so a corrupt file whose box
# sizes walk us through garbage can't loop for long.
_MAX_TOPLEVEL_BOXES = 64


def _read_mvhd_duration(body: bytes) -> float | None:
    """Return the movie duration (seconds) from an ``mvhd`` box body, or None."""
    if len(body) < 20:
        return None
    version = body[0]
    try:
        if version == 0:
            timescale, duration = struct.unpack(">II", body[12:20])
        else:
            if len(body) < 32:
                return None
            timescale, duration = struct.unpack(">IQ", body[20:32])
    except struct.error:
        return None
    if not timescale:
        return None
    return duration / timescale


def probe_container_seekability(filepath: str) -> dict[str, Any] | None:
    """Report whether a browser can seek this container without downloading it.

    OBS's "fragmented recording" mode writes a fragmented MP4: thousands of
    ``moof``/``mdat`` pairs, an ``mvhd`` duration of 0, and no sample index in
    ``moov``. ffmpeg reads such a file happily because it picks up the ``mfra``
    index box at the tail, so every server-side path (probing, scanning, clip
    cutting, frame extraction) is unaffected. **Browsers do not read ``mfra``.**
    They see a stream of unknown length: ``video.duration`` is ``Infinity``,
    ``seekable`` grows only as bytes arrive, and a seek past the buffered range
    silently lands somewhere else entirely. On a multi-GB recording that makes
    the page unusable until the whole file has been pulled over the wire.

    Detection is a bounded read of box headers — no ffprobe, no decoding, sub-
    millisecond even on multi-GB files — because ``moov`` always precedes the
    fragments and box bodies are skipped by seeking, never read.

    Returns ``{"fragmented", "header_duration", "browser_seekable"}``, or
    ``None`` when the shape can't be determined: a non-MP4 container (Matroska,
    WebM, raw QuickTime oddities), a truncated or unreadable file, or a box walk
    that runs off the rails. ``None`` means *unknown*, not *broken* — callers
    must stay silent rather than warn, since every non-MP4 source would
    otherwise be reported as a problem. Cached per ``(resolved_path, mtime)``.
    """
    if config.DEBUGGING:
        # Synthetic "fine" answer so DEBUGGING runs never touch the disk;
        # mirrors the probe_max_keyframe_gap / probe_video_properties branches.
        return {"fragmented": False, "header_duration": 300.0, "browser_seekable": True}

    key = _resolved_path_and_mtime(filepath)
    if key is None:
        return None
    if key in _container_seekability_cache:
        return _container_seekability_cache[key]

    result = _walk_mp4_for_seekability(filepath)
    _container_seekability_cache[key] = result
    return result


def _walk_mp4_for_seekability(filepath: str) -> dict[str, Any] | None:
    """Box-walk implementation behind probe_container_seekability()."""
    try:
        total = os.path.getsize(filepath)
        with open(filepath, "rb") as handle:
            offset = 0
            for _ in range(_MAX_TOPLEVEL_BOXES):
                if offset >= total:
                    break
                handle.seek(offset)
                header = handle.read(16)
                if len(header) < 8:
                    break
                size = struct.unpack(">I", header[0:4])[0]
                box_type = header[4:8]
                header_size = 8
                if size == 1:
                    if len(header) < 16:
                        break
                    size = struct.unpack(">Q", header[8:16])[0]
                    header_size = 16
                elif size == 0:
                    # Only legal on the final box: "extends to end of file".
                    size = total - offset
                if size < header_size:
                    break
                if box_type == b"moov":
                    return _scan_moov(handle, offset + header_size, offset + size)
                offset += size
    except (OSError, struct.error):
        return None
    return None


def _scan_moov(handle: Any, start: int, end: int) -> dict[str, Any] | None:
    """Scan a ``moov`` box's direct children for ``mvex`` and ``mvhd``."""
    fragmented = False
    header_duration: float | None = None
    offset = start
    while offset < end - 8:
        handle.seek(offset)
        header = handle.read(8)
        if len(header) < 8:
            break
        size = struct.unpack(">I", header[0:4])[0]
        box_type = header[4:8]
        if size < 8:
            break
        if box_type == b"mvex":
            # A movie-extends box is the spec's definition of "fragmented":
            # sample data lives in later moof boxes, not in moov's tables.
            fragmented = True
        elif box_type == b"mvhd":
            header_duration = _read_mvhd_duration(handle.read(min(size - 8, 32)))
        offset += size
    if not fragmented and header_duration is None:
        # Found moov but neither marker — not a shape we understand well enough
        # to make a claim about.
        return None
    return {
        "fragmented": fragmented,
        "header_duration": header_duration,
        "browser_seekable": not fragmented and bool(header_duration),
    }


# Kept beside the remuxed source until the user discards it. The suffix lands
# *after* the extension (``study_P15.mp4.orig``) on purpose: the participant
# scan globs ``*.mp4``, so anything still ending in .mp4 would come back as a
# second, phantom participant.
REMUX_ORIGINAL_SUFFIX = ".orig"

_remux_locks: dict[str, threading.Lock] = {}
_remux_locks_guard = threading.Lock()


def _remux_lock(resolved: str) -> threading.Lock:
    with _remux_locks_guard:
        lock = _remux_locks.get(resolved)
        if lock is None:
            lock = threading.Lock()
            _remux_locks[resolved] = lock
        return lock


def original_backup_path(filepath: str) -> Path:
    """Return where :func:`remux_to_faststart` parks this file's original."""
    return Path(str(filepath) + REMUX_ORIGINAL_SUFFIX)


def remux_to_faststart(
    filepath: str,
    *,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
) -> tuple[bool, str]:
    """Rewrite a fragmented recording into a normal, browser-seekable MP4.

    A stream copy — no re-encode, so the picture and audio are bit-identical and
    the cost is pure I/O (measured ~330 MB/s, i.e. ~11 s for a 3.7 GB capture).
    All it changes is the container: one contiguous ``mdat`` with a real sample
    index, and ``+faststart`` to put ``moov`` at the front. See
    :func:`probe_container_seekability` for why this matters.

    The source is replaced in place and the original parked at
    ``<name>.mp4.orig`` for the user to discard or restore. Replacing in place is
    what keeps participant ids stable — writing ``study_P15.remuxed.mp4`` beside
    the source would make it a second participant.

    Returns ``(ok, message)``. On any failure the source is left exactly as it
    was: the new file is only swapped in after it has been re-probed and matched
    against the original's duration, dimensions and stream count, because a
    silently truncated or half-muxed replacement is far worse than no remux.
    """
    src = Path(filepath)
    backup = original_backup_path(filepath)
    progress = on_progress or (lambda _fraction: None)

    if not src.is_file():
        return False, "Source file is missing."
    if backup.exists():
        return False, (
            f"An earlier original is still kept at '{backup.name}'. "
            "Discard or restore it before remuxing again."
        )

    before = probe_video_properties(str(src))
    if before is None:
        return False, "Could not probe the source file."

    resolved = str(src.resolve())
    with _remux_lock(resolved):
        # Re-check under the lock: a racing job may have finished the swap.
        if backup.exists():
            return False, "Another remux of this file just completed."
        # No .mp4 extension — pathlib.glob('*.mp4') matches dotfiles too, so a
        # hidden name alone would not keep the scratch file out of the
        # participant list. -f mp4 supplies the format the name no longer does.
        tmp = src.parent / f".{src.stem}.remux.{os.getpid()}.{threading.get_ident()}"
        command = [
            "ffmpeg",
            "-y",
            "-loglevel",
            config.FFMPEG_LOGLEVEL,
            "-i",
            str(src),
            # Every stream: these recordings routinely carry two audio tracks
            # (mic + system) and ffmpeg's default mapping would keep only one.
            "-map",
            "0",
            "-c",
            "copy",
            "-movflags",
            "+faststart",
            "-f",
            "mp4",
            str(tmp),
        ]
        try:
            result = _run_ffmpeg_with_progress(
                command,
                input_file=str(src),
                output_file=str(tmp),
                os_error_message="Failed to remux the source video.",
                on_progress=progress,
                expected_duration_sec=float(before.get("duration") or 0.0) or 1.0,
                cancel_flag=cancel_flag,
            )
            if result is None:
                return False, "Remux was cancelled or ffmpeg could not be started."
            if result.returncode != 0:
                return False, f"ffmpeg failed: {(result.stderr or '').strip()[:400]}"

            problem = _remux_output_mismatch(tmp, before)
            if problem is not None:
                return False, problem

            # Atomic within the directory: both names are on the same filesystem.
            src.rename(backup)
            try:
                tmp.rename(src)
            except OSError as error:
                backup.rename(src)  # put the original back, leave no gap
                return False, f"Could not swap in the remuxed file: {error}"
        finally:
            _delete_quietly(tmp)

    return True, f"Remuxed. Original kept as '{backup.name}'."


def _remux_output_mismatch(tmp: Path, before: dict[str, Any]) -> str | None:
    """Return why a remux output must not be swapped in, or None if it is sound."""
    if not tmp.is_file() or tmp.stat().st_size <= 0:
        return "Remux produced no output."
    seekability = probe_container_seekability(str(tmp))
    if seekability is None or not seekability["browser_seekable"]:
        return "The remuxed file is still not browser-seekable; keeping the original."
    after = probe_video_properties(str(tmp))
    if after is None:
        return "Could not probe the remuxed file; keeping the original."

    source_duration = float(before.get("duration") or 0.0)
    new_duration = float(after.get("duration") or 0.0)
    if source_duration > 0:
        tolerance = max(1.0, source_duration * 0.01)
        if abs(new_duration - source_duration) > tolerance:
            return (
                f"Remuxed duration ({new_duration:.0f}s) does not match the source "
                f"({source_duration:.0f}s); keeping the original."
            )
    if after.get("audio_track_count") != before.get("audio_track_count"):
        return (
            f"Remux kept {after.get('audio_track_count')} audio track(s) but the "
            f"source has {before.get('audio_track_count')}; keeping the original."
        )
    if (after.get("width"), after.get("height")) != (
        before.get("width"),
        before.get("height"),
    ):
        return "Remuxed dimensions do not match the source; keeping the original."
    return None


def restore_remux_original(filepath: str) -> tuple[bool, str]:
    """Put a kept ``.orig`` back, discarding the remuxed file."""
    src = Path(filepath)
    backup = original_backup_path(filepath)
    if not backup.is_file():
        return False, "No kept original to restore."
    with _remux_lock(str(src.resolve())):
        try:
            backup.replace(src)
        except OSError as error:
            return False, f"Could not restore the original: {error}"
    return True, "Original restored."


def discard_remux_original(filepath: str) -> tuple[bool, str]:
    """Delete a kept ``.orig`` once the user is happy with the remux."""
    backup = original_backup_path(filepath)
    if not backup.is_file():
        return False, "No kept original to discard."
    with _remux_lock(str(Path(filepath).resolve())):
        try:
            backup.unlink()
        except OSError as error:
            return False, f"Could not delete the original: {error}"
    return True, "Original deleted."


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
            capture_output=True,
            timeout=10,
            check=False,
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
        config.debug_ic(start_time, end_time)
    utils.debug_print(f"start_time is {start_time}, end_time is {end_time}")

    if end_time is INVALID_END_TIMESTAMP:
        utils.error_print(
            f"Invalid end timestamp (derived from start: '{start_time}')",
            ["Could not calculate end time. Check the timestamp format."],
        )
        return None

    # Parse each end independently with the canonical timestamp parser rather
    # than picking one strptime format for both ends: the two ends can
    # legitimately need different formats (a clip 59:50 -> 1:00:10 straddling
    # the hour, or a single timestamp whose default-duration end crosses it),
    # and MM:SS minutes may exceed 59 ("75:00", a long session written without
    # an hours component). A shared-format strptime rejected all of those and
    # silently dropped the clip.
    start_seconds = utils.timestamp_to_seconds(start_time)
    end_seconds = utils.timestamp_to_seconds(end_time)
    if start_seconds is None or end_seconds is None:
        utils.error_print(
            "Timestamp formatting error in get_duration().",
            [
                f"Start time: '{start_time}', End time: '{end_time}'",
                "Accepted formats: HH:MM:SS, MM:SS, or M:SS (e.g., 1:23:45, 12:34, 1:23)",
            ],
        )
        return None

    duration = int(end_seconds - start_seconds)
    if config.DEBUGGING:
        config.debug_ic(duration)
    return duration


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


def enforce_filesize_limit(
    path: str, *, cancel_flag: Callable[[], bool] | None = None
) -> None:
    """Compress *path* down to ``config.MAX_FILESIZE_MB`` when the limit is set.

    No-op when the limit is disabled (``0``). Apply this to a *final* clip
    artifact after any titlecard wrap or concat — never to an intermediate cut
    that will be re-encoded downstream (see the note in :func:`run_ffmpeg`).
    """
    if (
        config.MAX_FILESIZE_MB
        and config.MAX_FILESIZE_MB > 0
        and not compress_to_size(path, config.MAX_FILESIZE_MB, cancel_flag=cancel_flag)
    ):
        utils.warning_print(f"Could not compress '{path}' to target size")


def compress_to_size(
    filepath: str,
    target_size_mb: float,
    *,
    cancel_flag: Callable[[], bool] | None = None,
    on_progress: Callable[[float], None] | None = None,
) -> bool:
    """Recompress video to fit within target filesize using two-pass encoding.

    Always libx264, deliberately ignoring ``config.FFMPEG_VIDEO_ENCODER``: this is
    the one bitrate-*targeting* path, and hardware encoders have no ``-pass`` and
    much looser ABR. Measured on a 720p clip asking for 105 kbps video,
    h264_videotoolbox delivered 246 kbps (2.3x over) where x264 two-pass
    delivered 127 kbps — so a hardware attempt would overshoot the cap, get
    discarded, and pay for the two-pass anyway.

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

    Probing is parallel; the mismatch detection and its warnings stay on the
    calling thread so the output order is stable.
    """
    props_list: list[dict[str, Any] | None] = _parallel_probe(
        clip_paths, probe_video_properties
    )
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
                # Reuse the duration already probed into props_list (by
                # _detect_clip_mismatches) rather than a redundant lookup; fall
                # back to a fresh probe only for an unprobed clip.
                if props is not None and props.get("duration"):
                    dur = props["duration"]
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

    def build_command(encoder: str) -> list[str]:
        ffmpeg_command = ["ffmpeg", "-y", "-loglevel", config.FFMPEG_LOGLEVEL]
        for path in clip_paths:
            ffmpeg_command.extend(["-i", str(Path(path).resolve())])
        ffmpeg_command.extend(["-filter_complex", filter_str])
        ffmpeg_command.extend(["-map", "[outv]"])
        if has_audio:
            ffmpeg_command.extend(["-map", "[outa]"])
        ffmpeg_command.extend(video_encoder_args(encoder))
        ffmpeg_command.extend(["-c:a", "aac", output_file])
        return ffmpeg_command

    encoder = resolve_video_encoder()
    utils.debug_print(
        f"ffmpeg filter_complex concat: {' '.join(build_command(encoder))}"
    )
    try:
        result = run_ffmpeg_encode(
            build_command,
            encoder=encoder,
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
    with _concat_list_file(clip_paths) as concat_list_file:
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

                def build_reencode(encoder: str) -> list[str]:
                    return [
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
                        *video_encoder_args(encoder),
                        "-c:a",
                        "aac",
                        output_file,
                    ]

                ffmpeg_result = run_ffmpeg_encode(
                    build_reencode,
                    encoder=resolve_video_encoder(),
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


def concat_copy(
    clip_paths: list[str],
    output_file: str,
    *,
    cancel_flag: Callable[[], bool] | None = None,
    on_progress: Callable[[float], None] | None = None,
    expected_duration_sec: float | None = None,
) -> bool:
    """Concat clips with the demuxer + stream copy (no re-encode), or return False.

    A focused, quiet variant of _concatenate_demuxer for callers that build their
    own copy-safe inputs and own the fallback decision: no re-encode retry and no
    user-facing success message. Returns True only when the copy succeeds and the
    output verifies. Callers MUST guarantee the inputs are stream-copy compatible
    (matching codec/pix_fmt/SAR/timebase and audio params) — a mismatch produces a
    silently corrupt file with a zero return code, which this cannot detect.
    """
    with _concat_list_file(clip_paths) as concat_list_file:
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
            utils.debug_print(f"ffmpeg concat-copy command: {' '.join(ffmpeg_command)}")
            ffmpeg_result = run_ffmpeg_process(
                ffmpeg_command,
                input_file=concat_list_file,
                output_file=output_file,
                os_error_message="Stream-copy concat failed.",
                cancel_flag=cancel_flag,
                on_progress=on_progress,
                expected_duration_sec=expected_duration_sec,
            )
            if ffmpeg_result is None or ffmpeg_result.returncode != 0:
                return False
            return verify_output_file(output_file, "Concat")
        except OSError as e:
            utils.debug_print(f"Stream-copy concat failed: {e}")
            return False


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
