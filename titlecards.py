# -*- coding: utf-8 -*-
"""Titlecard and endcard generation for clipgen.

Prepends a short title card (first frame of source video + text overlay) and optionally
appends an endcard (last frame, no text) to each generated clip via FFmpeg. Falls back
to a color fill when no source frame is available. Duration is set by
config.TITLECARD_DURATION_SECONDS (default 2s). Enabled via config.TITLECARDS_ENABLED
or the --titlecards / --no-titlecards CLI flags.

Key functions:
  build_titlecard_frame(clip, resolution) – build the title card FFmpeg segment
  build_endcard_frame(resolution)         – build the endcard FFmpeg segment
  prepend_titlecard_to_clip(clip, clip_path) – prepend title card to an existing clip file
  append_endcard_to_clip(clip_path)          – append endcard to an existing clip file
"""

import os
import tempfile
import threading
from pathlib import Path
from typing import Dict, Optional

import config
import utils
import video
from utils import ClipRecord

# ---------------------------------------------------------------------------
# Endcard cache — keyed by resolution, shared across threads
# ---------------------------------------------------------------------------

_endcard_cache: Dict[str, str] = {}
_endcard_lock = threading.Lock()


def _get_video_resolution(filepath: str) -> Optional[str]:
    """Return 'WIDTHxHEIGHT' resolution string for the first video stream."""
    props = video.probe_video_properties(filepath)
    if props is None:
        return None
    return f"{props['width']}x{props['height']}"


def _build_drawtext_filter(text: str) -> str:
    safe_text = (text or "").strip()
    # The description has already been sanitized for filenames, but escape colons and backslashes just in case.
    safe_text = (
        safe_text.replace("\\", "\\\\").replace(":", "\\:").replace("'", "'\\''")
    )
    return (
        "drawtext=text='{}'"
        ":font=monospace"
        ":fontcolor=white"
        ":fontsize=min(w\\,h)/16"
        ":x=(w-text_w)/2"
        ":y=(h-text_h)/2"
        ":box=1:boxcolor=black@0.4:boxborderw=10"
    ).format(safe_text)


def _build_card_frame(
    *,
    resolution: str,
    background_path: Path,
    label: str,
    drawtext_filter: Optional[str] = None,
    allow_color_fallback: bool = False,
) -> Optional[str]:
    """Generate a short title/end card video segment.

    Shared implementation for titlecard and endcard frame generation.
    Returns the path to the generated card video, or None on failure.
    """
    if not resolution:
        return None

    use_image_background = background_path.is_file()
    if not use_image_background and not allow_color_fallback:
        return None

    try:
        with tempfile.NamedTemporaryFile(suffix=config.FILEFORMAT, delete=False) as tmp:
            card_path = tmp.name
    except OSError as error:
        utils.warning_print(
            f"Could not create temporary file for {label}.",
            [str(error)],
        )
        return None

    if config.DEBUGGING:
        utils.debug_print(
            f"Debugging enabled, would generate {label} '{card_path}' "
            f"at resolution {resolution}."
        )
        return None

    if "x" not in resolution:
        utils.warning_print(
            f"Invalid resolution string '{resolution}' for {label}.",
            ["Expected format 'WIDTHxHEIGHT' (e.g. '1280x720')."],
        )
        return None

    width_str, height_str = resolution.split("x", 1)
    width_str = width_str.strip()
    height_str = height_str.strip()

    vf_with_scale = (
        f"scale={resolution}:force_original_aspect_ratio=decrease,"
        f"pad={width_str}:{height_str}:(ow-iw)/2:(oh-ih)/2"
    )
    if drawtext_filter:
        vf_with_scale += f",{drawtext_filter}"

    if use_image_background:
        ffmpeg_command = [
            "ffmpeg",
            "-y",
            "-loglevel",
            config.FFMPEG_LOGLEVEL,
            "-loop",
            "1",
            "-t",
            str(config.TITLECARD_DURATION_SECONDS),
            "-i",
            str(background_path),
            "-vf",
            vf_with_scale,
            "-c:v",
            "libx264",
            "-pix_fmt",
            "yuv420p",
            card_path,
        ]
        input_label = str(background_path)
    else:
        ffmpeg_command = [
            "ffmpeg",
            "-y",
            "-loglevel",
            config.FFMPEG_LOGLEVEL,
            "-f",
            "lavfi",
            "-i",
            f"color=c=black:s={resolution}:d={config.TITLECARD_DURATION_SECONDS}",
            "-vf",
            vf_with_scale,
            "-c:v",
            "libx264",
            "-pix_fmt",
            "yuv420p",
            card_path,
        ]
        input_label = "lavfi:color"

    utils.debug_print(f"ffmpeg {label} command: {' '.join(ffmpeg_command)}")
    ffmpeg_result = video.run_ffmpeg_process(
        ffmpeg_command,
        input_file=input_label,
        output_file=card_path,
        os_error_message=f"ffmpeg could not successfully run for {label} generation.",
    )
    if ffmpeg_result is None or ffmpeg_result.returncode != 0:
        utils.warning_print(
            f"{label.capitalize()} generation failed; clip will be used without a {label}.",
            [ffmpeg_result.stderr.strip()]
            if ffmpeg_result and ffmpeg_result.stderr
            else None,
        )
        try:
            Path(card_path).unlink(missing_ok=True)
        except TypeError:
            if Path(card_path).exists():
                try:
                    Path(card_path).unlink()
                except OSError:
                    pass
        return None

    if not video.verify_output_file(card_path, f"{label.capitalize()} generation"):
        return None
    return card_path


def build_titlecard_frame(clip: ClipRecord, resolution: str) -> Optional[str]:
    """Generate a short titlecard video segment for a clip."""
    return _build_card_frame(
        resolution=resolution,
        background_path=Path("assets") / "titlecard.png",
        label="titlecard",
        drawtext_filter=_build_drawtext_filter(str(clip.get("desc", ""))),
        allow_color_fallback=True,
    )


def build_endcard_frame(resolution: str) -> Optional[str]:
    """Generate a short endcard video segment if assets/endcard.png exists."""
    return _build_card_frame(
        resolution=resolution,
        background_path=Path("assets") / "endcard.png",
        label="endcard",
    )


def get_or_build_endcard(resolution: str) -> Optional[str]:
    """Return a cached endcard path for the given resolution, building if needed."""
    with _endcard_lock:
        cached = _endcard_cache.get(resolution)
        if cached and Path(cached).is_file():
            return cached
    path = build_endcard_frame(resolution)
    if path:
        with _endcard_lock:
            existing = _endcard_cache.get(resolution)
            if existing and Path(existing).is_file():
                try:
                    Path(path).unlink()
                except OSError:
                    pass
                return existing
            _endcard_cache[resolution] = path
    return path


def clear_endcard_cache() -> None:
    """Remove all cached endcard temp files."""
    with _endcard_lock:
        for path in _endcard_cache.values():
            try:
                Path(path).unlink(missing_ok=True)
            except (OSError, TypeError):
                pass
        _endcard_cache.clear()


def prepend_titlecard_to_clip(
    clip: ClipRecord, clip_path: str, resolution: Optional[str] = None
) -> bool:
    """Prepend a generated titlecard to an existing clip file in-place.

    Returns True when the clip is usable (with or without titlecard), False on hard failure.
    """
    if not config.TITLECARDS_ENABLED:
        return True

    if config.DEBUGGING:
        utils.debug_print(
            f"Debugging enabled, skipping titlecard prepend for '{clip_path}'."
        )
        return True

    clip_file = Path(clip_path)
    if not clip_file.is_file():
        utils.warning_print(
            f"Cannot prepend titlecard; clip file not found: '{clip_path}'"
        )
        return False

    if not resolution:
        resolution = _get_video_resolution(clip_path)
    if not resolution:
        utils.warning_print(
            f"Could not determine video resolution for '{clip_path}'. "
            "Skipping titlecard for this clip."
        )
        return True

    titlecard_path = build_titlecard_frame(clip, resolution)
    if not titlecard_path:
        return True

    output_temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            suffix=config.FILEFORMAT, delete=False
        ) as out_tmp:
            output_temp_path = out_tmp.name

        ffmpeg_command = [
            "ffmpeg",
            "-y",
            "-loglevel",
            config.FFMPEG_LOGLEVEL,
            "-i",
            titlecard_path,
            "-i",
            clip_path,
            "-filter_complex",
            "[0:v][1:v]concat=n=2:v=1:a=0[v]",
            "-map",
            "[v]",
            "-map",
            "1:a?",
            "-c:v",
            "libx264",
            "-c:a",
            "aac",
            output_temp_path,
        ]
        utils.debug_print(
            "ffmpeg prepend titlecard filter concat command: "
            + " ".join(ffmpeg_command)
        )
        ffmpeg_result = video.run_ffmpeg_process(
            ffmpeg_command,
            input_file=clip_path,
            output_file=output_temp_path,
            os_error_message="Filter-based concat failed while prepending titlecard.",
        )
        if ffmpeg_result is None or ffmpeg_result.returncode != 0:
            utils.warning_print(
                f"Could not prepend titlecard to '{clip_path}'. Original clip will be kept.",
                [ffmpeg_result.stderr.strip()]
                if ffmpeg_result and ffmpeg_result.stderr
                else None,
            )
            return True

        if not video.verify_output_file(output_temp_path, "Prepend titlecard"):
            return True

        os.replace(output_temp_path, clip_path)
        return True
    finally:
        for temp in (titlecard_path, output_temp_path):
            if not temp:
                continue
            try:
                Path(temp).unlink()
            except OSError:
                pass


def append_endcard_to_clip(clip_path: str, resolution: Optional[str] = None) -> bool:
    """Append an endcard segment to an existing clip file in-place."""
    if not config.TITLECARDS_ENABLED:
        return True

    if config.DEBUGGING:
        utils.debug_print(
            f"Debugging enabled, skipping endcard append for '{clip_path}'."
        )
        return True

    clip_file = Path(clip_path)
    if not clip_file.is_file():
        utils.warning_print(
            f"Cannot append endcard; clip file not found: '{clip_path}'"
        )
        return False

    if not resolution:
        resolution = _get_video_resolution(clip_path)
    if not resolution:
        utils.warning_print(
            f"Could not determine video resolution for '{clip_path}'. "
            "Skipping endcard for this clip."
        )
        return True

    endcard_path = get_or_build_endcard(resolution)
    if not endcard_path:
        return True

    output_temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            suffix=config.FILEFORMAT, delete=False
        ) as out_tmp:
            output_temp_path = out_tmp.name

        ffmpeg_command = [
            "ffmpeg",
            "-y",
            "-loglevel",
            config.FFMPEG_LOGLEVEL,
            "-i",
            clip_path,
            "-i",
            endcard_path,
            "-filter_complex",
            "[0:v][1:v]concat=n=2:v=1:a=0[v]",
            "-map",
            "[v]",
            "-map",
            "0:a?",
            "-c:v",
            "libx264",
            "-c:a",
            "aac",
            output_temp_path,
        ]
        utils.debug_print(
            "ffmpeg append endcard filter concat command: " + " ".join(ffmpeg_command)
        )
        ffmpeg_result = video.run_ffmpeg_process(
            ffmpeg_command,
            input_file=clip_path,
            output_file=output_temp_path,
            os_error_message="Filter-based concat failed while appending endcard.",
        )
        if ffmpeg_result is None or ffmpeg_result.returncode != 0:
            utils.warning_print(
                f"Could not append endcard to '{clip_path}'. Original clip will be kept.",
                [ffmpeg_result.stderr.strip()]
                if ffmpeg_result and ffmpeg_result.stderr
                else None,
            )
            return True

        if not video.verify_output_file(output_temp_path, "Append endcard"):
            return True

        os.replace(output_temp_path, clip_path)
        return True
    finally:
        # Only clean up the concat output temp — endcard is managed by the cache
        if output_temp_path:
            try:
                Path(output_temp_path).unlink()
            except OSError:
                pass
