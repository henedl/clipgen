import os
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

import config
import utils
import video
from utils import ClipRecord


def _get_video_resolution(filepath: str) -> Optional[str]:
    """Return 'WIDTHxHEIGHT' resolution string for the first video stream."""
    if not Path(filepath).is_file():
        return None

    probe_command = [
        "ffprobe",
        "-v",
        "error",
        "-select_streams",
        "v:0",
        "-show_entries",
        "stream=width,height",
        "-of",
        "csv=s=x:p=0",
        filepath,
    ]
    utils.debug_print(f"ffprobe resolution command: {' '.join(probe_command)}")
    try:
        output = subprocess.check_output(probe_command, encoding="utf-8").strip()
    except (subprocess.CalledProcessError, FileNotFoundError, OSError) as error:
        utils.warning_print(
            f"Could not probe video resolution for '{filepath}'.",
            [str(error)],
        )
        return None

    if not output or "x" not in output:
        return None
    return output


def _build_drawtext_filter(text: str) -> str:
    safe_text = (text or "").strip()
    # The description has already been sanitized for filenames, but escape colons and backslashes just in case.
    safe_text = safe_text.replace("\\", "\\\\").replace(":", "\\:")
    return (
        "drawtext=text='{}'"
        ":font=monospace"
        ":fontcolor=white"
        ":fontsize=min(w\\,h)/16"
        ":x=(w-text_w)/2"
        ":y=(h-text_h)/2"
        ":box=1:boxcolor=black@0.4:boxborderw=10"
    ).format(safe_text)


def build_titlecard_frame(clip: ClipRecord, resolution: str) -> Optional[str]:
    """Generate a short titlecard video segment for a clip.

    Returns the path to the generated titlecard video, or None on failure.
    """
    if not resolution:
        return None

    background_path = Path("assets") / "titlecard.png"
    use_image_background = background_path.is_file()

    try:
        with tempfile.NamedTemporaryFile(
            suffix=config.FILEFORMAT, delete=False
        ) as tmp:
            titlecard_path = tmp.name
    except OSError as error:
        utils.warning_print(
            "Could not create temporary file for titlecard.",
            [str(error)],
        )
        return None

    if config.DEBUGGING:
        utils.debug_print(
            f"Debugging enabled, would generate titlecard '{titlecard_path}' "
            f"for clip '{clip.get('desc', '')}' at resolution {resolution}."
        )
        return None

    drawtext_filter = _build_drawtext_filter(str(clip.get("desc", "")))

    if "x" not in resolution:
        utils.warning_print(
            f"Invalid resolution string '{resolution}' for titlecard.",
            ["Expected format 'WIDTHxHEIGHT' (e.g. '1280x720')."],
        )
        return None

    width_str, height_str = resolution.split("x", 1)
    width_str = width_str.strip()
    height_str = height_str.strip()

    vf_with_scale = (
        f"scale={resolution}:force_original_aspect_ratio=decrease,"
        f"pad={width_str}:{height_str}:(ow-iw)/2:(oh-ih)/2,{drawtext_filter}"
    )

    if use_image_background:
        # Loop the PNG as video only for the configured duration; concat will
        # handle audio alignment and re-encoding if needed.
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
            titlecard_path,
        ]
        input_label = str(background_path)
    else:
        # Pure video color source with explicit duration; concat will sort out
        # audio stream differences via its re-encode fallback.
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
            titlecard_path,
        ]
        input_label = "lavfi:color"

    utils.debug_print(f"ffmpeg titlecard command: {' '.join(ffmpeg_command)}")
    ffmpeg_result = video._run_ffmpeg_process(
        ffmpeg_command,
        input_file=input_label,
        output_file=titlecard_path,
        os_error_message="ffmpeg could not successfully run for titlecard generation.",
    )
    if ffmpeg_result is None or ffmpeg_result.returncode != 0:
        utils.warning_print(
            "Titlecard generation failed; clip will be used without a titlecard.",
            [ffmpeg_result.stderr.strip()] if ffmpeg_result and ffmpeg_result.stderr else None,
        )
        try:
            Path(titlecard_path).unlink(missing_ok=True)
        except TypeError:
            if Path(titlecard_path).exists():
                try:
                    Path(titlecard_path).unlink()
                except OSError:
                    pass
        return None

    if not video._verify_output_file(titlecard_path, "Titlecard generation"):
        return None
    return titlecard_path


def prepend_titlecard_to_clip(clip: ClipRecord, clip_path: str) -> bool:
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
        ffmpeg_result = video._run_ffmpeg_process(
            ffmpeg_command,
            input_file=clip_path,
            output_file=output_temp_path,
            os_error_message="Filter-based concat failed while prepending titlecard.",
        )
        if ffmpeg_result is None or ffmpeg_result.returncode != 0:
            utils.warning_print(
                f"Could not prepend titlecard to '{clip_path}'. Original clip will be kept.",
                [ffmpeg_result.stderr.strip()] if ffmpeg_result and ffmpeg_result.stderr else None,
            )
            return True

        if not video._verify_output_file(output_temp_path, "Prepend titlecard"):
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

