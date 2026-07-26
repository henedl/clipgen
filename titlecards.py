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


def _x264_video_args() -> list[str]:
    """Shared libx264 output args for card generation and the card wrap.

    A fast preset + explicit CRF: cards are always encoded with these args, and
    the fallback wrap path (non-copy-safe body, or a failed copy concat)
    re-encodes the whole clip with them too, so the preset is the main lever on
    titlecard generation time there. Tuned via config.TITLECARD_ENCODE_PRESET /
    TITLECARD_ENCODE_CRF.
    """
    return [
        "-c:v",
        "libx264",
        "-preset",
        config.TITLECARD_ENCODE_PRESET,
        "-crf",
        str(config.TITLECARD_ENCODE_CRF),
        "-pix_fmt",
        "yuv420p",
    ]


def _resolve_channel_layout(probed: dict) -> str | None:
    """Resolve a silent-audio channel layout matching the probed body audio.

    Prefers the probed channel_layout; falls back to mono/stereo derived from the
    channel count. Returns None when the layout can't be resolved (unsupported).
    """
    layout = probed.get("audio_channel_layout")
    if isinstance(layout, str) and layout.strip():
        return layout.strip()
    channels = probed.get("audio_channels") or 0
    if channels == 1:
        return "mono"
    if channels == 2:
        return "stereo"
    return None


def _body_is_copy_safe(probed: dict | None) -> bool:
    """Whether the clip body can be stream-copy-concatenated with freshly built cards.

    Conservative gate: only h264 / yuv420p bodies with a finite fps qualify, and if
    the body has audio it must be aac with a usable sample rate and a resolvable
    channel layout (so we can build a matching silent card audio track). Anything
    else routes to the filter_complex re-encode path. This gate — not a return-code
    check — is the sole guard against silent stream-copy corruption.
    """
    if not probed:
        return False
    if probed.get("video_codec") != "h264":
        return False
    if probed.get("pix_fmt") != "yuv420p":
        return False
    fps = probed.get("fps") or 0.0
    if not (isinstance(fps, (int, float)) and fps > 0):
        return False
    if probed.get("audio_codec"):
        if probed.get("audio_codec") != "aac":
            return False
        if not (probed.get("audio_sample_rate") or 0) > 0:
            return False
        if _resolve_channel_layout(probed) is None:
            return False
    return True


def _build_drawtext_filter(text: str) -> str:
    safe_text = (text or "").strip()
    # The description has already been sanitized for filenames, but escape colons and backslashes just in case.
    safe_text = (
        safe_text.replace("\\", "\\\\").replace(":", "\\:").replace("'", "'\\''")
    )
    return (
        f"drawtext=text='{safe_text}'"
        # The text is static — disable drawtext's %{...} expansion so a
        # description containing "%" (which sanitize_filename keeps) renders
        # literally instead of erroring the encode.
        ":expansion=none"
        ":font=monospace"
        ":fontcolor=white"
        ":fontsize=min(w\\,h)/16"
        ":x=(w-text_w)/2"
        ":y=(h-text_h)/2"
        ":box=1:boxcolor=black@0.4:boxborderw=10"
    )


def _ffmpeg_color(value: str) -> str:
    """Convert a #rrggbb value to ffmpeg's 0xRRGGBB; pass named colors through."""
    v = (value or "").strip()
    if v.startswith("#") and len(v) == 7:
        return "0x" + v[1:]
    return v or "black"


def _build_card_frame(
    *,
    resolution: str,
    background_path: Path | None,
    label: str,
    drawtext_filter: str | None = None,
    allow_color_fallback: bool = False,
    fill_color: str = "black",
    cancel_flag: Callable[[], bool] | None = None,
    card_duration_seconds: int | None = None,
    match_fps: float | None = None,
    audio_match: dict | None = None,
) -> str | None:
    """Generate a short title/end card video segment.

    Shared implementation for titlecard and endcard frame generation.
    Returns the path to the generated card video, or None on failure. A None
    *background_path* (or a path that doesn't exist) renders a solid-color card
    when *allow_color_fallback* is set, otherwise returns None.

    When *match_fps* / *audio_match* are supplied (the stream-copy wrap path), the
    card is encoded to match the clip body: a fixed framerate + timebase and a
    self-contained silent AAC track at the body's sample rate / channel layout, so
    the card can be concat-demuxed alongside the body with ``-c copy``. Without
    them the card is video-only (the filter_complex wrap injects silence itself).
    """
    if not resolution:
        return None

    duration = (
        config.TITLECARD_DURATION_SECONDS
        if card_duration_seconds is None
        else card_duration_seconds
    )

    use_image_background = background_path is not None and background_path.is_file()
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
        f"pad={width_str}:{height_str}:(ow-iw)/2:(oh-ih)/2,setsar=1"
    )
    if drawtext_filter:
        vf_with_scale += f",{drawtext_filter}"

    if use_image_background:
        assert background_path is not None  # guaranteed by use_image_background
        video_input = [
            "-loop",
            "1",
            "-t",
            str(duration),
            "-i",
            str(background_path),
        ]
        input_label = str(background_path)
    else:
        video_input = [
            "-f",
            "lavfi",
            "-i",
            f"color=c={_ffmpeg_color(fill_color)}:s={resolution}:d={duration}",
        ]
        input_label = "lavfi:color"

    # Silent AAC track matching the body so the card can be stream-copied alongside
    # it (stream-copy wrap path only). Video is always input 0, silence input 1.
    audio_input: list[str] = []
    audio_out_args: list[str] = []
    map_args: list[str] = []
    if audio_match:
        layout = audio_match["channel_layout"]
        rate = int(audio_match["sample_rate"])
        channels = int(audio_match.get("channels") or (1 if layout == "mono" else 2))
        audio_input = [
            "-f",
            "lavfi",
            "-t",
            str(duration),
            "-i",
            f"anullsrc=channel_layout={layout}:sample_rate={rate}",
        ]
        audio_out_args = ["-c:a", "aac", "-ar", str(rate), "-ac", str(channels)]
        map_args = ["-map", "0:v", "-map", "1:a"]

    video_out_args = list(_x264_video_args())
    if match_fps:
        # Fixed framerate + timebase so concat-demuxer copy sees consistent cards.
        video_out_args += [
            "-r",
            f"{match_fps:g}",
            "-video_track_timescale",
            "90000",
        ]

    ffmpeg_command = [
        "ffmpeg",
        "-y",
        "-loglevel",
        config.FFMPEG_LOGLEVEL,
        *video_input,
        *audio_input,
        "-vf",
        vf_with_scale,
        *map_args,
        *video_out_args,
        *audio_out_args,
        card_path,
    ]

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
        Path(card_path).unlink(missing_ok=True)
        return None

    if not video.verify_output_file(card_path, f"{label.capitalize()} generation"):
        return None
    return card_path


def resolve_card_background(kind: str) -> tuple[Path | None, bool, bool, str]:
    """Resolve the configured background for a card from config.

    *kind* is "title" or "end". Returns (background_path, allow_color, skip, fill_color):
      - background_path: image to use, or None for a solid-color card.
      - allow_color: whether to fall back to a color fill when no image is present.
      - skip: True when no card should be produced at all (endcard "none").
      - fill_color: the color used for the fill (the configured solid color when
        the card is set to a solid color, otherwise "black" for the
        missing-image fallback).

    Selection ids (config.TITLECARD_IMAGE / config.ENDCARD_IMAGE): empty = bundled
    default asset; CARD_IMAGE_COLOR = solid color; CARD_IMAGE_NONE = no endcard;
    any other value is an uploaded filename under TITLECARD_IMAGES_DIRNAME (falling
    back to the bundled default when the file is missing). Title cards always render
    (their text is the point), so they never skip.
    """
    if kind == "end":
        value = config.ENDCARD_IMAGE
        default_asset = "endcard.png"
        # Endcards historically only render when an image exists (no color fill).
        default_allow_color = False
        solid_color = config.ENDCARD_COLOR
    else:
        value = config.TITLECARD_IMAGE
        default_asset = "titlecard.png"
        default_allow_color = True
        solid_color = config.TITLECARD_COLOR

    if kind == "end" and value == config.CARD_IMAGE_NONE:
        return (None, False, True, "black")
    if value == config.CARD_IMAGE_COLOR:
        return (None, True, False, solid_color)

    default_path = utils.get_bundled_assets_root() / "assets" / default_asset
    if not value:
        return (default_path, default_allow_color, False, "black")

    # Only a bare filename inside the upload pool is a valid selection. Reject
    # path separators / traversal so a stray config value can't resolve outside
    # TITLECARD_IMAGES_DIRNAME. The Studio settings route validates this already;
    # this guards the CLI / persisted-settings path too.
    if Path(value).name != value:
        return (default_path, default_allow_color, False, "black")

    upload_path = (
        utils.get_effective_output_dir() / config.TITLECARD_IMAGES_DIRNAME / value
    )
    if upload_path.is_file():
        return (upload_path, default_allow_color, False, "black")
    # Selected upload is missing — fall back to the bundled default.
    return (default_path, default_allow_color, False, "black")


def build_titlecard_frame(
    clip: ClipRecord,
    resolution: str,
    *,
    cancel_flag: Callable[[], bool] | None = None,
    card_duration_seconds: int | None = None,
    match_fps: float | None = None,
    audio_match: dict | None = None,
) -> str | None:
    """Generate a short titlecard video segment for a clip."""
    background_path, allow_color, skip, fill_color = resolve_card_background("title")
    if skip:
        return None
    return _build_card_frame(
        resolution=resolution,
        background_path=background_path,
        label="titlecard",
        drawtext_filter=_build_drawtext_filter(str(clip.get("desc", ""))),
        allow_color_fallback=allow_color,
        fill_color=fill_color,
        cancel_flag=cancel_flag,
        card_duration_seconds=card_duration_seconds,
        match_fps=match_fps,
        audio_match=audio_match,
    )


def build_endcard_frame(
    resolution: str,
    *,
    cancel_flag: Callable[[], bool] | None = None,
    card_duration_seconds: int | None = None,
    match_fps: float | None = None,
    audio_match: dict | None = None,
) -> str | None:
    """Generate a short endcard video segment for the configured background."""
    background_path, allow_color, skip, fill_color = resolve_card_background("end")
    if skip:
        return None
    return _build_card_frame(
        resolution=resolution,
        background_path=background_path,
        label="endcard",
        allow_color_fallback=allow_color,
        fill_color=fill_color,
        cancel_flag=cancel_flag,
        card_duration_seconds=card_duration_seconds,
        match_fps=match_fps,
        audio_match=audio_match,
    )


def _audio_match_signature(audio_match: dict | None) -> str:
    """Cache-key fragment identifying an endcard's silent-audio params (or none)."""
    if not audio_match:
        return "noaudio"
    return f"{audio_match.get('sample_rate')}:{audio_match.get('channel_layout')}"


def get_or_build_endcard(
    resolution: str,
    *,
    cancel_flag: Callable[[], bool] | None = None,
    card_duration_seconds: int | None = None,
    match_fps: float | None = None,
    audio_match: dict | None = None,
) -> str | None:
    """Return a cached endcard path for the given resolution, building if needed."""
    duration = (
        config.TITLECARD_DURATION_SECONDS
        if card_duration_seconds is None
        else card_duration_seconds
    )
    # Key by the selected endcard so switching the background (or, for a solid
    # color, the color itself) doesn't reuse a stale cached file. The fps + audio
    # signature must be part of the key: a copy-path endcard is encoded to match a
    # specific body's framerate/audio, so reusing it for a body with different
    # params would silently desync the stream-copy concat.
    endcard_id = config.ENDCARD_IMAGE or "__default__"
    if config.ENDCARD_IMAGE == config.CARD_IMAGE_COLOR:
        endcard_id = endcard_id + config.ENDCARD_COLOR
    audio_sig = _audio_match_signature(audio_match)
    fps_sig = f"{match_fps:g}" if match_fps else "nofps"
    cache_key = f"{resolution}:{duration}:{endcard_id}:{fps_sig}:{audio_sig}"
    with _endcard_lock:
        cached = _endcard_cache.get(cache_key)
        if cached and Path(cached).is_file():
            return cached
    path = build_endcard_frame(
        resolution,
        cancel_flag=cancel_flag,
        card_duration_seconds=duration,
        match_fps=match_fps,
        audio_match=audio_match,
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
            except OSError:
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
) -> tuple[bool, bool]:
    """Prepend a titlecard and append an endcard to a clip.

    Fast path: when the clip body is a copy-safe shape (see _body_is_copy_safe), only
    the two cards are encoded — matched to the body's fps/pixel-format and (if present)
    a silent AAC track at the body's audio params — then all three are joined with the
    concat demuxer + ``-c copy``, so the already-cut body is not re-encoded. Any
    non-copy-safe body, or a failed copy concat, falls back to a single filter_complex
    encode that re-encodes the whole clip.

    Returns ``(clip_ok, cards_applied)``:

    - ``clip_ok`` is True when the clip file is usable afterwards (either wrapped or
      left untouched on soft failure), and False only on a hard failure (e.g. the clip
      file is missing).
    - ``cards_applied`` is True only when the cards are actually in the output file.

    Both values matter to the caller. A soft failure leaves a perfectly good *unwrapped*
    clip, so it must still be recorded — but recording it as ``titlecards: true`` makes
    the manifest lie, and the generate-cache check (``server.py`` Phase 1) then skips
    the clip forever, so the card can never be applied. Callers must persist
    ``cards_applied``, not the requested flag.

    When *cancel_flag* is supplied and returns True during a card or wrap encode, the
    in-flight ffmpeg is terminated and the original clip is left untouched.
    """
    cards_enabled = (
        config.TITLECARDS_ENABLED
        if titlecards_enabled is None
        else bool(titlecards_enabled)
    )
    if not cards_enabled:
        return (True, False)

    card_duration = (
        config.TITLECARD_DURATION_SECONDS
        if titlecard_duration_seconds is None
        else titlecard_duration_seconds
    )

    if config.DEBUGGING:
        utils.debug_print(
            f"Debugging enabled, skipping titlecard/endcard wrap for '{clip_path}'."
        )
        return (True, False)

    clip_file = Path(clip_path)
    if not clip_file.is_file():
        utils.warning_print(
            f"Cannot wrap clip with titlecards; clip file not found: '{clip_path}'"
        )
        return (False, False)

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
        return (True, False)

    has_clip_audio = bool(probed and probed.get("audio_codec"))
    clip_duration = (
        float(probed["duration"]) if probed and probed.get("duration") else 0.0
    )
    expected_wrap_duration = (
        clip_duration + 2 * card_duration if clip_duration > 0 else None
    )

    # When the body is a known copy-safe shape (h264 / yuv420p / aac), encode only
    # the cards to match it and concat-demux with -c copy, avoiding a full body
    # re-encode. Otherwise (or if the copy concat fails) fall back to the
    # filter_complex re-encode path below.
    copy_safe = _body_is_copy_safe(probed)
    match_fps: float | None = None
    audio_match: dict | None = None
    if copy_safe and probed:
        match_fps = float(probed["fps"])
        if has_clip_audio:
            audio_match = {
                "sample_rate": probed["audio_sample_rate"],
                "channel_layout": _resolve_channel_layout(probed),
                "channels": probed.get("audio_channels") or 0,
            }

    titlecard_temps: list[str] = []
    output_temp_path: str | None = None
    try:
        if copy_safe:
            titlecard_path = build_titlecard_frame(
                clip,
                resolution,
                cancel_flag=cancel_flag,
                card_duration_seconds=card_duration,
                match_fps=match_fps,
                audio_match=audio_match,
            )
            if titlecard_path:
                titlecard_temps.append(titlecard_path)
            endcard_path = get_or_build_endcard(
                resolution,
                cancel_flag=cancel_flag,
                card_duration_seconds=card_duration,
                match_fps=match_fps,
                audio_match=audio_match,
            )
            if not titlecard_path and not endcard_path:
                # Both cards failed to build; nothing to do, keep clip as-is.
                return (True, False)

            segments = [p for p in (titlecard_path, clip_path, endcard_path) if p]
            with tempfile.NamedTemporaryFile(
                suffix=config.FILEFORMAT, delete=False
            ) as out_tmp:
                output_temp_path = out_tmp.name
            if video.concat_copy(
                segments,
                output_temp_path,
                cancel_flag=cancel_flag,
                on_progress=on_progress,
                expected_duration_sec=expected_wrap_duration,
            ):
                os.replace(output_temp_path, clip_path)
                output_temp_path = None
                return (True, True)
            # Copy concat failed — discard the temp output and fall through to the
            # re-encode path (which rebuilds video-only cards).
            if output_temp_path:
                try:
                    Path(output_temp_path).unlink()
                except OSError:
                    pass
                output_temp_path = None
            utils.debug_print(
                f"Stream-copy card wrap failed for '{clip_path}'; re-encoding instead."
            )

        # Fallback: single filter_complex encode that re-encodes the whole clip.
        titlecard_path = build_titlecard_frame(
            clip,
            resolution,
            cancel_flag=cancel_flag,
            card_duration_seconds=card_duration,
        )
        if titlecard_path:
            titlecard_temps.append(titlecard_path)
        endcard_path = get_or_build_endcard(
            resolution,
            cancel_flag=cancel_flag,
            card_duration_seconds=card_duration,
        )
        if not titlecard_path and not endcard_path:
            # Both cards failed to build; nothing to do, keep clip as-is.
            return (True, False)

        input_args, filter_complex, map_args = _build_wrap_filter_and_inputs(
            titlecard_path=titlecard_path,
            clip_path=clip_path,
            endcard_path=endcard_path,
            has_clip_audio=has_clip_audio,
            card_duration=card_duration,
        )

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
            *_x264_video_args(),
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
            return (True, False)

        if not video.verify_output_file(output_temp_path, "Wrap clip with cards"):
            return (True, False)

        os.replace(output_temp_path, clip_path)
        output_temp_path = None
        return (True, True)
    finally:
        # Titlecards are per-clip temps; endcards are managed by _endcard_cache.
        for titlecard_temp in titlecard_temps:
            try:
                Path(titlecard_temp).unlink()
            except OSError:
                pass
        if output_temp_path:
            try:
                Path(output_temp_path).unlink()
            except OSError:
                pass
