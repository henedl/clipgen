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
  wrap_clip_with_cards(clip, clip_path)   – single-pass prepend+append via one ffmpeg encode
"""

import os
import tempfile
import threading
from collections.abc import Callable
from pathlib import Path

import config
import utils
import video
from utils import ClipRecord

# ---------------------------------------------------------------------------
# Endcard cache — keyed by resolution, shared across threads
# ---------------------------------------------------------------------------

_endcard_cache: dict[str, str] = {}
_endcard_lock = threading.Lock()


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
    drawtext_filter: str | None = None,
    allow_color_fallback: bool = False,
    cancel_flag: Callable[[], bool] | None = None,
    card_duration_seconds: int | None = None,
) -> str | None:
    """Generate a short title/end card video segment.

    Shared implementation for titlecard and endcard frame generation.
    Returns the path to the generated card video, or None on failure.
    """
    if not resolution:
        return None

    duration = (
        config.TITLECARD_DURATION_SECONDS
        if card_duration_seconds is None
        else card_duration_seconds
    )

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
            str(duration),
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
            f"color=c=black:s={resolution}:d={duration}",
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
        cancel_flag=cancel_flag,
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


def build_titlecard_frame(
    clip: ClipRecord,
    resolution: str,
    *,
    cancel_flag: Callable[[], bool] | None = None,
    card_duration_seconds: int | None = None,
) -> str | None:
    """Generate a short titlecard video segment for a clip."""
    return _build_card_frame(
        resolution=resolution,
        background_path=utils.get_bundled_assets_root() / "assets" / "titlecard.png",
        label="titlecard",
        drawtext_filter=_build_drawtext_filter(str(clip.get("desc", ""))),
        allow_color_fallback=True,
        cancel_flag=cancel_flag,
        card_duration_seconds=card_duration_seconds,
    )


def build_endcard_frame(
    resolution: str,
    *,
    cancel_flag: Callable[[], bool] | None = None,
    card_duration_seconds: int | None = None,
) -> str | None:
    """Generate a short endcard video segment if assets/endcard.png exists."""
    return _build_card_frame(
        resolution=resolution,
        background_path=utils.get_bundled_assets_root() / "assets" / "endcard.png",
        label="endcard",
        cancel_flag=cancel_flag,
        card_duration_seconds=card_duration_seconds,
    )


def get_or_build_endcard(
    resolution: str,
    *,
    cancel_flag: Callable[[], bool] | None = None,
    card_duration_seconds: int | None = None,
) -> str | None:
    """Return a cached endcard path for the given resolution, building if needed."""
    duration = (
        config.TITLECARD_DURATION_SECONDS
        if card_duration_seconds is None
        else card_duration_seconds
    )
    cache_key = f"{resolution}:{duration}"
    with _endcard_lock:
        cached = _endcard_cache.get(cache_key)
        if cached and Path(cached).is_file():
            return cached
    path = build_endcard_frame(
        resolution, cancel_flag=cancel_flag, card_duration_seconds=duration
    )
    if path:
        with _endcard_lock:
            existing = _endcard_cache.get(cache_key)
            if existing and Path(existing).is_file():
                try:
                    Path(path).unlink()
                except OSError:
                    pass
                return existing
            _endcard_cache[cache_key] = path
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


def _input_count(input_args: list[str]) -> int:
    """Return the number of ffmpeg -i inputs present in *input_args*."""
    return sum(1 for tok in input_args if tok == "-i")


def _build_wrap_filter_and_inputs(
    *,
    titlecard_path: str | None,
    clip_path: str,
    endcard_path: str | None,
    has_clip_audio: bool,
    card_duration: int,
) -> tuple[list[str], str, list[str]]:
    """Construct the ffmpeg input args, filter_complex string, and map args for wrapping.

    Returns a tuple of (input_args, filter_complex, map_args). Caller appends codec
    options and the output path. The clip itself is always one of the inputs; title
    and end cards are optional. Audio is preserved (with silence padding aligned to
    card durations) when the clip has an audio stream, otherwise the output is
    video-only so we don't depend on clip duration probing.
    """
    input_args: list[str] = []
    video_labels: list[str] = []
    audio_labels: list[str] = []

    def add_input(args: list[str]) -> int:
        idx = _input_count(input_args)
        input_args.extend(args)
        return idx

    def anullsrc_input(duration: int) -> list[str]:
        return [
            "-f",
            "lavfi",
            "-t",
            str(duration),
            "-i",
            "anullsrc=channel_layout=stereo:sample_rate=48000",
        ]

    # Video inputs, in playback order: title, clip, end.
    title_idx: int | None = None
    end_idx: int | None = None

    if titlecard_path:
        title_idx = add_input(["-i", titlecard_path])
        video_labels.append(f"[{title_idx}:v]")

    clip_idx = add_input(["-i", clip_path])
    video_labels.append(f"[{clip_idx}:v]")

    if endcard_path:
        end_idx = add_input(["-i", endcard_path])
        video_labels.append(f"[{end_idx}:v]")

    filter_parts: list[str] = []

    if has_clip_audio:
        # Matching silent audio tracks for each card so concat sees a 1:1 v/a pairing.
        title_a_label: str | None = None
        end_a_label: str | None = None
        if title_idx is not None:
            idx = add_input(anullsrc_input(card_duration))
            title_a_label = f"[{idx}:a]"
        if end_idx is not None:
            idx = add_input(anullsrc_input(card_duration))
            end_a_label = f"[{idx}:a]"

        # Normalize clip audio to a consistent rate/layout so concat doesn't reject it.
        filter_parts.append(
            f"[{clip_idx}:a]aresample=48000,"
            "aformat=channel_layouts=stereo:sample_rates=48000[aclip]"
        )

        if title_a_label:
            audio_labels.append(title_a_label)
        audio_labels.append("[aclip]")
        if end_a_label:
            audio_labels.append(end_a_label)

        interleaved: list[str] = []
        for v_label, a_label in zip(video_labels, audio_labels):
            interleaved.append(v_label)
            interleaved.append(a_label)
        n = len(video_labels)
        filter_parts.append(f"{''.join(interleaved)}concat=n={n}:v=1:a=1[v][a]")
        map_args = ["-map", "[v]", "-map", "[a]"]
    else:
        n = len(video_labels)
        filter_parts.append(f"{''.join(video_labels)}concat=n={n}:v=1:a=0[v]")
        map_args = ["-map", "[v]"]

    return (input_args, ";".join(filter_parts), map_args)


def wrap_clip_with_cards(
    clip: ClipRecord,
    clip_path: str,
    resolution: str | None = None,
    *,
    cancel_flag: Callable[[], bool] | None = None,
    on_progress: Callable[[float], None] | None = None,
    titlecards_enabled: bool | None = None,
    titlecard_duration_seconds: int | None = None,
) -> bool:
    """Prepend a titlecard and append an endcard to a clip in a single ffmpeg encode.

    Replaces the previous two-invocation (prepend then append) flow so the clip body
    is decoded and re-encoded once instead of twice. Returns True when the clip file
    is usable afterwards (either wrapped or left untouched on soft failure), and
    False only on a hard failure (e.g. the clip file is missing). When *cancel_flag*
    is supplied and returns True during a card or wrap encode, the in-flight ffmpeg
    is terminated and the original clip file is left untouched.
    """
    cards_enabled = (
        config.TITLECARDS_ENABLED
        if titlecards_enabled is None
        else bool(titlecards_enabled)
    )
    if not cards_enabled:
        return True

    card_duration = (
        config.TITLECARD_DURATION_SECONDS
        if titlecard_duration_seconds is None
        else titlecard_duration_seconds
    )

    if config.DEBUGGING:
        utils.debug_print(
            f"Debugging enabled, skipping titlecard/endcard wrap for '{clip_path}'."
        )
        return True

    clip_file = Path(clip_path)
    if not clip_file.is_file():
        utils.warning_print(
            f"Cannot wrap clip with titlecards; clip file not found: '{clip_path}'"
        )
        return False

    # One probe gives us both audio presence and (when needed) resolution; it's
    # cached by resolved path in video.probe_video_properties.
    probed = video.probe_video_properties(clip_path)
    if not resolution and probed:
        resolution = f"{probed['width']}x{probed['height']}"
    if not resolution:
        utils.warning_print(
            f"Could not determine video resolution for '{clip_path}'. "
            "Skipping title/endcard for this clip."
        )
        return True

    has_clip_audio = bool(probed and probed.get("audio_codec"))
    clip_duration = (
        float(probed["duration"]) if probed and probed.get("duration") else 0.0
    )
    expected_wrap_duration = (
        clip_duration + 2 * card_duration if clip_duration > 0 else None
    )

    titlecard_path = build_titlecard_frame(
        clip,
        resolution,
        cancel_flag=cancel_flag,
        card_duration_seconds=card_duration,
    )
    endcard_path = get_or_build_endcard(
        resolution,
        cancel_flag=cancel_flag,
        card_duration_seconds=card_duration,
    )

    if not titlecard_path and not endcard_path:
        # Both cards failed to build; nothing to do, keep clip as-is.
        return True

    input_args, filter_complex, map_args = _build_wrap_filter_and_inputs(
        titlecard_path=titlecard_path,
        clip_path=clip_path,
        endcard_path=endcard_path,
        has_clip_audio=has_clip_audio,
        card_duration=card_duration,
    )

    output_temp_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            suffix=config.FILEFORMAT, delete=False
        ) as out_tmp:
            output_temp_path = out_tmp.name

        ffmpeg_command: list[str] = [
            "ffmpeg",
            "-y",
            "-loglevel",
            config.FFMPEG_LOGLEVEL,
            *input_args,
            "-filter_complex",
            filter_complex,
            *map_args,
            "-c:v",
            "libx264",
            "-pix_fmt",
            "yuv420p",
        ]
        if has_clip_audio:
            ffmpeg_command.extend(["-c:a", "aac"])
        ffmpeg_command.append(output_temp_path)

        utils.debug_print(
            "ffmpeg wrap clip with cards command: " + " ".join(ffmpeg_command)
        )
        ffmpeg_result = video.run_ffmpeg_process(
            ffmpeg_command,
            input_file=clip_path,
            output_file=output_temp_path,
            os_error_message="Filter-based concat failed while wrapping clip with cards.",
            cancel_flag=cancel_flag,
            on_progress=on_progress,
            expected_duration_sec=expected_wrap_duration,
        )
        if ffmpeg_result is None or ffmpeg_result.returncode != 0:
            utils.warning_print(
                f"Could not wrap '{clip_path}' with title/endcard. Original clip will be kept.",
                [ffmpeg_result.stderr.strip()]
                if ffmpeg_result and ffmpeg_result.stderr
                else None,
            )
            return True

        if not video.verify_output_file(output_temp_path, "Wrap clip with cards"):
            return True

        os.replace(output_temp_path, clip_path)
        output_temp_path = None
        return True
    finally:
        # Titlecards are per-clip temps; endcards are managed by _endcard_cache.
        if titlecard_path:
            try:
                Path(titlecard_path).unlink()
            except OSError:
                pass
        if output_temp_path:
            try:
                Path(output_temp_path).unlink()
            except OSError:
                pass
