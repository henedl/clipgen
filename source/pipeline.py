"""Clip processing pipeline: clips, reels, screenshots, GIFs, and manifest
regeneration.

Deliberately separate from the CLI entry-point module so the web layer can drive
the pipeline without importing it.
"""

import concurrent.futures
import difflib
import hashlib
import os
import threading
import time
from collections.abc import Callable
from pathlib import Path
from typing import Any

import config
import files
import titlecards
import transcripts
import profiling
import utils
import video
import viewer
from utils import ClipRecord

# Active progress bar reference, set during clip pipeline so nested functions
# (e.g. fuzzy match prompts) can pause/resume the live display.
_active_progress = None
_active_secondary_task = None


def _resolve_titlecard_options(
    titlecards_enabled: bool | None,
    titlecard_duration_seconds: int | None,
) -> tuple[bool, int]:
    """Resolve per-request titlecard settings, falling back to config defaults."""
    enabled = (
        config.TITLECARDS_ENABLED
        if titlecards_enabled is None
        else bool(titlecards_enabled)
    )
    duration = (
        config.TITLECARD_DURATION_SECONDS
        if titlecard_duration_seconds is None
        else int(titlecard_duration_seconds)
    )
    return enabled, duration


def _card_image_identity(kind: str) -> str:
    """Cache identity for one card kind, matching what will actually be rendered.

    Derived from titlecards.resolve_card_background() — the same resolution the
    encoder uses — so the recorded identity never drifts from the rendered card:
      - solid color: the color is baked in (e.g. "__color__#ff0000") so changing
        it invalidates cached clips;
      - "none" endcard: the no-card sentinel;
      - a resolved upload: its filename;
      - the bundled default — including the silent fallback when a selected upload
        is missing on disk: the empty-string default. Collapsing the missing-upload
        case to "" (rather than recording the absent filename) means a clip rendered
        with the fallback regenerates once the file later appears, instead of cache-
        matching on a filename it was never actually rendered with.
    """
    background_path, _allow_color, skip, fill_color = (
        titlecards.resolve_card_background(kind)
    )
    if skip:
        return config.CARD_IMAGE_NONE
    if background_path is None:
        return config.CARD_IMAGE_COLOR + fill_color
    images_dir = utils.get_effective_output_dir() / config.TITLECARD_IMAGES_DIRNAME
    if background_path.parent == images_dir:
        return background_path.name
    return ""


def _resolve_titlecard_images(cards_enabled: bool) -> tuple[str, str]:
    """Return the selected (titlecard, endcard) image identities when enabled.

    Image selection is config-global (no per-request override). Returns empty
    strings when cards are disabled so artifact records and the Studio cache-skip
    comparison agree on a single canonical value. Each identity reflects the
    background that will actually be rendered (see _card_image_identity).
    """
    if not cards_enabled:
        return "", ""
    return _card_image_identity("title"), _card_image_identity("end")


def is_excel_worksheet(worksheet: Any) -> bool:
    """Return True if worksheet is the Excel adapter (local file, no URL)."""
    spread = getattr(worksheet, "spreadsheet", None)
    return spread is not None and getattr(spread, "url", None) is None


# ---- Clip processing pipeline ----


def _resolve_clip_workers() -> int:
    """Return the effective parallel worker count for clip generation."""
    workers = config.CLIP_PARALLEL_WORKERS
    if workers <= 0:
        workers = min(4, os.cpu_count() or 1)
    return workers


# Large .mp4 files in the input dir, keyed by path + mtime_ns (one glob/stat pass per run).
_fuzzy_input_videos_cache: dict[str, tuple[int | None, list[tuple[int, Path]]]] = {}

# Serializes the missing-video branch of _check_source_video: reel preparation
# runs per-clip in worker threads, so the shared missing_videos / fuzzy_matches
# structures and any fuzzy-match prompt must not be touched concurrently.
_fuzzy_match_lock = threading.Lock()


def _large_input_videos(input_dir: Path) -> list[tuple[int, Path]]:
    """Return (size_bytes, path) for source-video candidates under *input_dir*."""
    dir_str = str(input_dir)
    try:
        mtime_ns: int | None = (
            input_dir.stat().st_mtime_ns if input_dir.is_dir() else None
        )
    except OSError:
        mtime_ns = None

    cached = _fuzzy_input_videos_cache.get(dir_str)
    if cached is not None and cached[0] == mtime_ns:
        return cached[1]

    size_threshold = config.MIN_SOURCE_VIDEO_SIZE_MB * 1_000_000
    entries: list[tuple[int, Path]] = []
    for p in input_dir.glob(f"*{config.FILEFORMAT}"):
        try:
            size = p.stat().st_size
        except OSError:
            continue
        if size >= size_threshold:
            entries.append((size, p))

    _fuzzy_input_videos_cache[dir_str] = (mtime_ns, entries)
    return entries


def _resolve_one_source(
    base_name: str,
    missing_videos: set[str],
    skip_detail: str,
    fuzzy_matches: dict[str, str | None],
) -> str | None:
    """Resolve one expected source-video filename to an on-disk path, or None.

    Exact match in the input directory wins. Otherwise scans for large .mp4 files
    and offers the closest fuzzy match for user confirmation. Paths already seen
    in *missing_videos* are not reported again. The missing/fuzzy branch is
    serialized by ``_fuzzy_match_lock`` so parallel reel workers neither race nor
    interleave prompts.
    """
    full_path = utils.resolve_input_path(base_name)
    if full_path.is_file():
        return str(full_path)

    full_path_str = str(full_path)

    with _fuzzy_match_lock:
        # Check fuzzy match cache (value may be None = user rejected or no candidate)
        if full_path_str in fuzzy_matches:
            return fuzzy_matches[full_path_str]

        input_dir = utils.get_effective_input_dir()
        candidates: list[tuple[float, int, Path]] = []
        for size, p in _large_input_videos(input_dir):
            ratio = difflib.SequenceMatcher(
                None, base_name.lower(), p.name.lower()
            ).ratio()
            candidates.append((ratio, size, p))

        # Sort by similarity descending, then file size descending as tiebreaker
        candidates.sort(key=lambda c: (c[0], c[1]), reverse=True)

        if candidates and candidates[0][0] >= 0.7 and not utils.NO_INPUT_MODE:
            _best_ratio, best_size, best_path = candidates[0]
            size_gb = best_size / 1_000_000_000
            # Pause progress bar so the prompt is visible and input is rendered
            paused = False
            if _active_progress is not None:
                _active_progress.stop()
                paused = True
            utils.info_print(f"Source video '{base_name}' not found.")
            utils.info_print(
                f"Closest match found: '{best_path.name}' ({size_gb:.1f} GB)"
            )
            answer = utils.read_user_input("Use this file instead? [y/n]\n>> ")
            if paused and _active_progress is not None:
                _active_progress.start()
            if answer.strip().lower() == "y":
                resolved = str(best_path)
                fuzzy_matches[full_path_str] = resolved
                return resolved

        # No match or user rejected — cache and report error
        fuzzy_matches[full_path_str] = None
        if full_path_str not in missing_videos:
            missing_videos.add(full_path_str)
            utils.error_print(
                f"Source video file not found: '{base_name}'",
                [
                    f"Expected location: {full_path_str}",
                    f"Expected format: {{study}}_{{participant}}{config.FILEFORMAT}",
                    skip_detail,
                ],
            )
        return None


def _check_source_video(
    clip: ClipRecord,
    missing_videos: set[str],
    skip_detail: str,
    fuzzy_matches: dict[str, str | None],
) -> str | None:
    """Resolve a clip's source video(s) and return the first source path, or None.

    A participant's session may span several source videos that form one
    continuous timeline (a recording that broke off, or a diary study). They are
    declared either by the spreadsheet ``Filename`` row (plus-separated, e.g.
    ``"morning.mp4 + afternoon.mp4"``) or auto-detected on disk by numbered
    suffix (``study_P01-1.mp4``, ``study_P01-2.mp4``).

    Resolution order:
    - Override present → resolve each plus-separated part; any missing → skip clip.
    - No override → the plain patterned name (``config.SOURCE_FILENAME_PATTERN``,
      default ``{study}_{participant}.mp4``) wins when present (single video).
      Only when it is absent do we auto-detect numbered parts; if none exist,
      fall back to the fuzzy-match prompt on the plain name.

    When 2+ videos resolve, the duration timeline is built and stored on
    ``clip['source_timeline']`` so the cut/artifact stages can map global
    timestamps into the right sub-video. The single-video path never probes
    durations and leaves ``source_timeline`` unset.
    """
    override = clip.get("source_filename")
    names = files.get_source_video_filenames(
        clip["study"], clip["participant"], override
    )

    resolved: list[str] = []
    if override:
        for name in names:
            one = _resolve_one_source(name, missing_videos, skip_detail, fuzzy_matches)
            if one is None:
                return None
            resolved.append(one)
    else:
        plain = names[0]
        plain_path = utils.resolve_input_path(plain)
        if plain_path.is_file():
            resolved = [str(plain_path)]
        else:
            numbered = files.discover_numbered_source_videos(
                utils.get_effective_input_dir(), clip["study"], clip["participant"]
            )
            if numbered:
                resolved = [str(p) for p in numbered]
            else:
                one = _resolve_one_source(
                    plain, missing_videos, skip_detail, fuzzy_matches
                )
                if one is None:
                    return None
                resolved = [one]

    if len(resolved) >= 2:
        timeline = video.build_source_timeline(resolved)
        if timeline is None:
            utils.error_print(
                "Could not read durations for all source videos; clip will be skipped.",
                [
                    f"Participant '{clip['participant']}' in study '{clip['study']}'",
                    "Source files: " + ", ".join(Path(p).name for p in resolved),
                ],
            )
            return None
        clip["source_timeline"] = timeline

    return resolved[0]


def _prepare_and_check_clip(
    clip: ClipRecord,
    missing_videos: set[str],
    fuzzy_matches: dict[str, str | None],
) -> tuple[ClipRecord, str | None]:
    """Prepare one clip and validate that its source video exists.

    Returns:
        Tuple of (prepared clip dict, source video path or None).
        When None is returned for source video, the clip should be skipped.
    """
    clip = files.prepare_clip(clip)
    if not clip["times"]:
        return (clip, None)

    base_video = _check_source_video(
        clip,
        missing_videos,
        f"Clips for participant '{clip['participant']}' in study '{clip['study']}' will be skipped.",
        fuzzy_matches,
    )
    return (clip, base_video)


def _local_timestamp(seconds: float) -> str:
    """Format local *seconds* as ``H:MM:SS`` for ffmpeg.

    Always emits the hours component so every ffmpeg-bound timestamp shares one
    uniform shape. This is a consistency/readability invariant, not a
    correctness requirement: ``video.get_duration`` parses each end
    independently (via ``utils.timestamp_to_seconds``) and tolerates mixed
    ``M:SS`` / ``H:MM:SS`` pairs. These strings feed ffmpeg, not the
    spreadsheet, so the explicit hours are harmless.
    """
    return utils.seconds_to_timestamp(round(seconds), force_hours=True)


def _point_source(
    timeline: list[tuple[str, int, int]], start_time: str
) -> tuple[str, str, int] | None:
    """Map a single global timestamp to ``(source_path, local_start_ts, remaining)``.

    For screenshot/GIF cuts, which key off the start time only. ``remaining`` is
    how many seconds of the owning sub-video follow the start (used to clamp GIF
    length so it never reads past the sub-video). Returns None if out of range.
    """
    global_start = utils.timestamp_to_seconds(start_time) or 0.0
    mapped = utils.map_global_to_segment(timeline, global_start)
    if mapped is None:
        return None
    index, local_start = mapped
    seg_duration = timeline[index][1]
    remaining = max(1, seg_duration - round(local_start))
    return (timeline[index][0], _local_timestamp(local_start), remaining)


def _stitch_clip_pieces(
    timeline: list[tuple[str, int, int]],
    pieces: list[tuple[int, float, float]],
    out_name: str,
    file_extension: str,
    cancel_flag: Callable[[], bool] | None,
) -> bool:
    """Cut each boundary-spanning piece to a temp file and concatenate into *out_name*.

    Used when a clip range straddles the boundary between two source videos: each
    piece is cut from its sub-video at local offsets, then stitched into one clip
    (mirrors the reel temp-file + concat pattern). Returns False (cleaning up) if
    any piece fails.
    """
    temp_paths: list[str] = []
    for n, (index, local_start, local_end) in enumerate(pieces):
        tmp = files.get_unique_filename(
            f"_multipart_{n + 1}{file_extension}", file_format=file_extension
        )
        if video.run_ffmpeg(
            input_file=timeline[index][0],
            output_file=tmp,
            start_pos=_local_timestamp(local_start),
            end_pos=_local_timestamp(local_end),
            reencode=config.REENCODING,
            cancel_flag=cancel_flag,
        ):
            temp_paths.append(tmp)
        else:
            files.release_reservation(tmp)
            for p in temp_paths:
                Path(p).unlink(missing_ok=True)
            return False
    ok = video.concatenate_clips(
        temp_paths, out_name, reencode_on_fail=True, cancel_flag=cancel_flag
    )
    for p in temp_paths:
        try:
            Path(p).unlink(missing_ok=True)
        except OSError:
            pass
    return ok


def cut_global_range(
    timeline: list[tuple[str, int, int]] | None,
    base_video: str,
    start_seconds: float,
    end_seconds: float,
    out_path: str,
    *,
    reencode: bool,
    cancel_flag: Callable[[], bool] | None = None,
) -> dict[str, Any] | None:
    """Cut a GLOBAL ``[start, end]`` span into *out_path*, mapping into sub-videos.

    For single-video participants (*timeline* is None) this is a plain cut from
    *base_video* at the global times. For multi-video participants the global span
    is mapped onto the timeline: a within-segment span is cut from the owning
    sub-video at local offsets; a boundary-spanning span is stitched from each
    piece (via :func:`_stitch_clip_pieces`).

    Returns the source-reference fields for the artifact —
    ``{sourceVideo (basename), localStart, localEnd}`` plus ``parts`` for a
    stitched span — or ``None`` on failure (caller releases the reservation).
    Used by the Studio intake (single clip + intake reel); mirrors the artifact
    contract the main clip pipeline writes.
    """
    if timeline is None:
        ok = video.run_ffmpeg(
            base_video,
            out_path,
            _local_timestamp(start_seconds),
            _local_timestamp(end_seconds),
            reencode,
            cancel_flag=cancel_flag,
        )
        if not ok:
            return None
        return {
            "sourceVideo": Path(base_video).name,
            "localStart": start_seconds,
            "localEnd": end_seconds,
        }

    pieces = utils.map_global_range_to_segments(timeline, start_seconds, end_seconds)
    if not pieces:
        return None
    if len(pieces) == 1:
        seg_index, local_start, local_end = pieces[0]
        ok = video.run_ffmpeg(
            timeline[seg_index][0],
            out_path,
            _local_timestamp(local_start),
            _local_timestamp(local_end),
            reencode,
            cancel_flag=cancel_flag,
        )
        if not ok:
            return None
        return {
            "sourceVideo": Path(timeline[seg_index][0]).name,
            "localStart": local_start,
            "localEnd": local_end,
        }

    extension = Path(out_path).suffix or config.FILEFORMAT
    if not _stitch_clip_pieces(timeline, pieces, out_path, extension, cancel_flag):
        return None
    parts = [
        {
            "sourceVideo": Path(timeline[index][0]).name,
            "localStart": local_start,
            "localEnd": local_end,
        }
        for index, local_start, local_end in pieces
    ]
    return {
        "sourceVideo": parts[0]["sourceVideo"],
        "localStart": parts[0]["localStart"],
        "localEnd": parts[0]["localEnd"],
        "parts": parts,
    }


def _process_single_clip_segments(
    clip: ClipRecord,
    base_video: str,
    missing_videos: set[str],
    *,
    filename_prefix: str = "",
    output_format: str = "clip",
    collect_paths: bool = False,
    include_severity: bool = False,
    enforce_size: bool = True,
    cancel_flag: Callable[[], bool] | None = None,
    titlecards_enabled: bool | None = None,
    titlecard_duration_seconds: int | None = None,
    pad_pre: float = 0.0,
    pad_post: float = 0.0,
    max_duration: float = 0.0,
) -> tuple[int, list[tuple[str, int]], bool]:
    """Process one clip's segments: run ffmpeg for each (start, end), optionally collect output paths.

    Returns ``(generated, output_paths, cards_applied)``. ``cards_applied`` is True only
    when cards were requested *and* every generated segment actually got them. A wrap
    that soft-fails leaves a usable but *unwrapped* clip; recording that as carded makes
    the generate-cache (``server.py`` Phase 1) skip it forever, so the card could never
    be applied. See ``titlecards.wrap_clip_with_cards``.

    Caller must have already called prepare_clip(clip). Does not add to missing_videos; caller handles that.

    Args:
        clip: Prepared clip dict with 'times', 'category', 'study', 'participant', 'desc'
        base_video: Path to source video file
        missing_videos: Set of already-reported missing paths (read-only here)
        filename_prefix: Prefix for output filename (e.g. '_reel_part_' for reel)
        collect_paths: If True, return list of (output_path, time_index) pairs; otherwise return empty list
        enforce_size: If True (default), compress each final clip to config.MAX_FILESIZE_MB
            after the titlecard wrap. Pass False for intermediate cuts (e.g. reel parts)
            that get concatenated/re-encoded downstream, so the size cap is applied once
            to the final artifact rather than wasted on throwaway pieces.
        include_severity: If True and clip has severity, include [Severity] in filename
        cancel_flag: Optional callable; checked before each segment and forwarded to
            ffmpeg helpers so an in-flight encode can be terminated. Already-finished
            segments are kept; the partial output of the killed segment is unlinked.
        pad_pre: Seconds to extend each clip's start earlier (negative trims inward).
        pad_post: Seconds to extend each clip's end later (negative trims inward).
            No effect on gifs/screenshots (they key off the start time only).
        max_duration: When > 0, cap each (padded) clip's length to this many seconds
            by pulling the end in; also caps gif duration.

    Returns:
        (number of segments successfully generated, list of (out_path, time_index) pairs
        if collect_paths else []). The ``time_index`` indexes into ``clip['times']``;
        downstream callers look up start/end strings there rather than carrying duplicates.
    """
    generated = 0
    output_paths: list[tuple[str, int]] = []
    # Cards are per-segment; a single soft failure makes the whole clip uncarded
    # as far as the manifest is concerned, so regeneration retries it.
    all_cards_applied = True
    extension_map = {
        "clip": config.FILEFORMAT,
        "screen": config.SCREENSHOT_FORMAT,
        "gif": config.GIF_FORMAT,
    }
    file_extension = extension_map.get(output_format)
    if not file_extension:
        utils.error_print(f"Unsupported output format: '{output_format}'")
        return (generated, output_paths, False)

    severity_tag = (
        f"[{clip['severity']}]" if include_severity and clip.get("severity") else ""
    )
    template = f"{filename_prefix}[{clip['category']}]{severity_tag} {clip['study']} {clip['participant']} {clip['desc']}{file_extension}"

    cards_enabled, card_duration = _resolve_titlecard_options(
        titlecards_enabled, titlecard_duration_seconds
    )

    # Multi-video participants carry a duration timeline; global timestamps are
    # mapped into the owning sub-video at cut time. Absent = single-video fast
    # path (unchanged behavior, no probing).
    timeline = clip.get("source_timeline")

    # Padding / max-duration (Workflows artifact nodes); the default path is a
    # no-op and never probes. The EOF limit is only needed when *extending* the
    # end — trimming or capping can never push the end out. run_ffmpeg now
    # clamps an over-running span itself, so this is about keeping the padded
    # end *honest* in the artifact record rather than about salvaging the cut.
    padding_active = pad_pre != 0.0 or pad_post != 0.0 or max_duration > 0.0
    span_limit: float | None = None
    if pad_post > 0.0:
        if timeline:
            span_limit = float(sum(seg[1] for seg in timeline))
        else:
            span_limit = video.get_file_duration(base_video)

    for time_idx, (start_time, end_time) in enumerate(clip["times"]):
        if cancel_flag and cancel_flag():
            break
        # Screenshots are a single frame with no duration — leave them untouched.
        if padding_active and output_format != "screen":
            s0 = utils.timestamp_to_seconds(start_time) or 0.0
            e0 = utils.timestamp_to_seconds(end_time) or 0.0
            s0, e0 = utils.apply_span_padding(
                s0,
                e0,
                pad_pre=pad_pre,
                pad_post=pad_post,
                max_duration=max_duration,
                limit=span_limit,
            )
            start_time = utils.seconds_to_timestamp(s0, force_hours=True)
            end_time = utils.seconds_to_timestamp(e0, force_hours=True)
        try:
            out_name = files.get_unique_filename(template, file_format=file_extension)
            if config.DEBUGGING:
                config.debug_ic(out_name)
        except (TypeError, UnicodeEncodeError, UnicodeDecodeError) as e:
            if config.DEBUGGING:
                config.debug_ic(e, clip)
            utils.error_print(
                f"Character encoding issue occurred: {e}",
                [
                    f"Category: '{clip['category']}', Study: '{clip['study']}', Participant: '{clip['participant']}'",
                    "Try simplifying the description or category names to use only ASCII characters.",
                ],
            )
            return (generated, output_paths, False)
        if output_format == "clip":
            if timeline:
                global_start = utils.timestamp_to_seconds(start_time) or 0.0
                global_end = utils.timestamp_to_seconds(end_time) or 0.0
                pieces = utils.map_global_range_to_segments(
                    timeline, global_start, global_end
                )
                if not pieces:
                    ok = False
                elif len(pieces) == 1:
                    seg_index, local_start, local_end = pieces[0]
                    ok = video.run_ffmpeg(
                        input_file=timeline[seg_index][0],
                        output_file=out_name,
                        start_pos=_local_timestamp(local_start),
                        end_pos=_local_timestamp(local_end),
                        reencode=config.REENCODING,
                        cancel_flag=cancel_flag,
                    )
                else:
                    ok = _stitch_clip_pieces(
                        timeline,
                        pieces,
                        out_name,
                        file_extension,
                        cancel_flag,
                    )
            else:
                ok = video.run_ffmpeg(
                    input_file=base_video,
                    output_file=out_name,
                    start_pos=start_time,
                    end_pos=end_time,
                    reencode=config.REENCODING,
                    cancel_flag=cancel_flag,
                )
            if ok and cards_enabled:
                # Wrap at the clip's own resolution: for multi-video participants
                # a clip may come from a later part whose resolution differs from
                # the first source, so trusting the clip avoids a concat mismatch.
                ok, cards_applied = titlecards.wrap_clip_with_cards(
                    clip,
                    out_name,
                    cancel_flag=cancel_flag,
                    titlecards_enabled=cards_enabled,
                    titlecard_duration_seconds=card_duration,
                )
                # A soft wrap failure keeps a usable, unwrapped clip — record it,
                # but not as carded, so the generate-cache retries it later.
                if ok and not cards_applied:
                    all_cards_applied = False
            # Enforce the size cap on the finished clip (after any wrap re-encode),
            # not on intermediate reel parts (enforce_size=False).
            if ok and enforce_size:
                video.enforce_filesize_limit(out_name, cancel_flag=cancel_flag)
        else:  # output_format == 'screen' or 'gif' — keys off the start time only
            src_path: str | None = base_video
            cut_ts = start_time
            remaining: int | None = None
            if timeline:
                point = _point_source(timeline, start_time)
                if point is None:
                    src_path = None
                else:
                    src_path, cut_ts, remaining = point
            if src_path is None:
                ok = False
            elif output_format == "screen":
                ok = video.extract_screenshot(
                    input_file=src_path,
                    output_file=out_name,
                    timestamp=cut_ts,
                    cancel_flag=cancel_flag,
                )
            else:  # output_format == 'gif'
                # pad_pre already shifted cut_ts above and pad_post is moot for a
                # fixed-length gif. max_duration caps the length, floored at 1s so
                # a fractional cap never yields a 0s gif.
                gif_duration = config.DEFAULT_GIF_DURATION_SECONDS
                if remaining is not None:
                    gif_duration = min(gif_duration, remaining)
                if max_duration > 0:
                    gif_duration = max(1, min(gif_duration, int(max_duration)))
                ok = video.extract_gif(
                    input_file=src_path,
                    output_file=out_name,
                    timestamp=cut_ts,
                    duration_seconds=gif_duration,
                    cancel_flag=cancel_flag,
                )
        if ok:
            generated += 1
            if collect_paths:
                output_paths.append((out_name, time_idx))
        elif cancel_flag and cancel_flag():
            # Killed mid-encode: the partial output is corrupt; remove it.
            try:
                Path(out_name).unlink(missing_ok=True)
            except OSError:
                pass
            break
        else:
            files.release_reservation(out_name)
    return (generated, output_paths, all_cards_applied and generated > 0)


def _parallel_map_ordered(
    items: list[Any],
    worker_fn: Callable[[Any], Any],
    *,
    workers: int,
    results: list[Any],
    on_error: Callable[[int, Exception], Any],
    cancel_flag: Callable[[], bool] | None = None,
    on_done: Callable[[], None] | None = None,
) -> dict[concurrent.futures.Future[Any], int]:
    """Submit ``worker_fn`` over ``items`` to a ThreadPoolExecutor, collecting
    each result into ``results`` at its original index.

    The caller owns ``results`` (pre-allocated to ``len(items)``), so it sets the
    sentinel/shape for slots that never complete. On cancel, pending futures are
    cancelled and the collection loop breaks, leaving un-collected slots at their
    pre-filled value. ``on_error`` receives ``(idx, exc)`` for a per-item failure
    and returns the value to store. ``on_done`` fires once per collected future
    (e.g. to advance a progress bar). Returns the future->index map so callers
    that need post-cancel draining (see ``_run_clip_pipeline``) can use it.
    """
    # pipeline.clip (summed item time) over pipeline.pool_wall (elapsed) is
    # effective parallelism — the number CLIP_PARALLEL_WORKERS is tuned against.
    # ffmpeg.run times each encode but knows nothing about overlap, so it cannot
    # tell a saturated pool from a serialized one. Both labels are shared by
    # every clip pool (here, _regenerate_reel, and Studio's three) because the
    # knob is global; the ratio is meaningful in aggregate.
    worker_fn = profiling.timed("pipeline.clip")(worker_fn)
    with (
        profiling.span("pipeline.pool_wall"),
        concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool,
    ):
        future_to_idx = {
            pool.submit(worker_fn, item): idx for idx, item in enumerate(items)
        }
        for future in concurrent.futures.as_completed(future_to_idx):
            if cancel_flag and cancel_flag():
                for f in future_to_idx:
                    f.cancel()
                break
            idx = future_to_idx[future]
            try:
                results[idx] = future.result()
            except Exception as exc:
                results[idx] = on_error(idx, exc)
            if on_done is not None:
                on_done()
    return future_to_idx


def _run_clip_pipeline(
    clips_list: list[Any],
    *,
    empty_warning: str,
    intro_message: str,
    task_label: str,
    per_clip_fn: Callable[[Any, set[str]], Any],
    show_fallback_counter: bool = False,
    secondary_task_label: str | None = None,
    parallel: bool = False,
    cancel_flag: Callable[[], bool] | None = None,
    on_clip_complete: Callable[[int, int], None] | None = None,
) -> tuple[list[Any], set[str]]:
    """Run shared clip-processing pipeline and return per-clip results.

    When *parallel* is True and there are at least 2 clips, a ThreadPoolExecutor
    is used to run ``per_clip_fn`` concurrently (same worker count as
    ``_resolve_clip_workers()``).  Results are collected in original order.

    On cancellation, any per-clip slots that did not complete are dropped from
    the returned list — callers receive only successful results and can iterate
    them without defensive ``None`` checks.
    """
    if not clips_list:
        utils.warning_print(empty_warning)
        return ([], set())

    utils.standard_print(intro_message)
    missing_videos: set[str] = set()

    def wrapped_process(clip: Any) -> Any:
        return per_clip_fn(clip, missing_videos)

    total_clips = len(clips_list)
    workers = _resolve_clip_workers() if parallel else 0
    use_parallel = parallel and workers >= 2 and total_clips >= 2

    completed_count = 0

    def _notify_clip_done() -> None:
        nonlocal completed_count
        completed_count += 1
        if on_clip_complete is not None:
            on_clip_complete(completed_count, total_clips)

    def _drain_ran_futures(
        future_to_idx: dict[concurrent.futures.Future[Any], int],
        results: list[Any],
    ) -> None:
        """Capture results from futures that ran to completion but were not
        collected because the as_completed loop broke on cancel.

        The pool's ``__exit__`` waits for already-running futures, so by the
        time this runs every started future is done. Collecting their results
        keeps any files they produced tracked (e.g. reel ``_reel_part_*``
        segments) so the caller's cancel cleanup can remove them instead of
        orphaning them. Futures cancelled before starting stay ``None`` and are
        filtered out by the caller.
        """
        for fut, idx in future_to_idx.items():
            if results[idx] is None and fut.done() and not fut.cancelled():
                try:
                    results[idx] = fut.result()
                except Exception:
                    results[idx] = None

    def _clip_error(idx: int, exc: Exception) -> Any:
        clip = clips_list[idx]
        desc = (clip.get("desc") or "")[: config.PROGRESS_DESCRIPTION_LENGTH]
        utils.error_print(
            f"Clip failed: [{clip.get('participant', '')}] {desc}",
            [str(exc)],
        )
        return []

    if use_parallel:
        results: list[Any] = [None] * total_clips
        progress = utils.create_progress_bar()
        if progress:
            global _active_progress, _active_secondary_task
            _active_progress = progress
            with progress:
                task = progress.add_task(task_label, total=total_clips)

                def _on_done() -> None:
                    progress.update(task, advance=1)
                    _notify_clip_done()

                future_to_idx = _parallel_map_ordered(
                    clips_list,
                    wrapped_process,
                    workers=workers,
                    results=results,
                    on_error=_clip_error,
                    cancel_flag=cancel_flag,
                    on_done=_on_done,
                )
                _drain_ran_futures(future_to_idx, results)
            _active_progress = None
            _active_secondary_task = None
        else:
            future_to_idx = _parallel_map_ordered(
                clips_list,
                wrapped_process,
                workers=workers,
                results=results,
                on_error=_clip_error,
                cancel_flag=cancel_flag,
                on_done=_notify_clip_done,
            )
            _drain_ran_futures(future_to_idx, results)
    else:
        results = []
        progress = utils.create_progress_bar()
        if progress:
            _active_progress = progress
            with progress:
                task = progress.add_task(task_label, total=total_clips)
                if secondary_task_label:
                    _active_secondary_task = progress.add_task(
                        secondary_task_label, total=total_clips
                    )
                for clip in clips_list:
                    if cancel_flag and cancel_flag():
                        break
                    desc_preview = (clip.get("desc") or "")[
                        : config.PROGRESS_DESCRIPTION_LENGTH
                    ]
                    participant = clip.get("participant", "")
                    progress.update(
                        task, description=f"[{participant}] {desc_preview}..."
                    )
                    results.append(wrapped_process(clip))
                    progress.update(task, advance=1)
                    _notify_clip_done()
            _active_progress = None
            _active_secondary_task = None
        else:
            for index, clip in enumerate(clips_list, start=1):
                if cancel_flag and cancel_flag():
                    break
                if (
                    show_fallback_counter
                    and getattr(config, "VERBOSITY", config.STANDARD) >= config.VERBOSE
                    and total_clips > 1
                ):
                    utils.verbose_print(f"Processing clip {index} of {total_clips}...")
                results.append(wrapped_process(clip))
                _notify_clip_done()

    if missing_videos:
        utils.standard_print(f"* Missing source video files: {len(missing_videos)}")
    # Drop slots from cancelled futures so callers get completed results only —
    # the sequential path produces the same shape via early-break + append.
    results = [r for r in results if r is not None]
    return (results, missing_videos)


def _embed_transcript_on_artifacts(
    clip: Any,
    base_video: str,
    artifacts: list[dict[str, Any]],
    segment_details: list[tuple[str, int]],
    transcript_cache: dict[str, Any],
    transcripts_manifest: dict[str, Any] | None = None,
) -> None:
    """Embed transcript segments on clip artifact records.

    Source priority: transcripts manifest -> in-memory cache. Modifies artifacts in-place.
    ``segment_details`` is a list of (out_path, time_index) pairs aligned to ``artifacts``.
    """

    participant = clip.get("participant", "")
    manifest = transcripts_manifest or transcripts.load_transcripts_manifest()
    source_transcripts = manifest.get("source_transcripts", {})
    corrections = manifest.get("corrections", [])

    full_transcript = None
    transcript_version = ""
    if participant and participant in source_transcripts:
        entry = source_transcripts[participant]
        raw_segments = entry.get("segments", [])
        corrected = transcripts.apply_corrections(raw_segments, corrections)
        full_transcript = transcripts.TranscriptResult(
            segments=corrected,
            language=entry.get("language", ""),
            source_file=entry.get("source_file", str(base_video)),
            model=entry.get("model", ""),
        )
        transcript_version = entry.get("transcribed_at", "")
    elif transcript_cache.get(base_video):
        full_transcript = transcript_cache[base_video]

    if not full_transcript:
        return

    times = clip.get("times", [])
    for art_idx, (_out_path, time_idx) in enumerate(segment_details):
        if art_idx >= len(artifacts):
            break
        start_str, end_str = times[time_idx]
        start_sec = utils.timestamp_to_seconds(start_str) or 0.0
        end_sec = utils.timestamp_to_seconds(end_str) or 0.0
        clipped = transcripts.filter_segments(
            full_transcript, start_sec, end_sec, offset_to_zero=True
        )
        transcript_segments = []
        for seg_idx, seg in enumerate(clipped["segments"]):
            transcript_segments.append(
                {
                    "id": f"{participant}:{seg_idx}",
                    "start": seg["start"],
                    "end": seg["end"],
                    "text": seg["text"],
                }
            )
        if transcript_segments:
            artifacts[art_idx]["transcript"] = transcript_segments
            if transcript_version:
                artifacts[art_idx]["transcript_version"] = transcript_version


def _transcribe_segments(
    clip: Any,
    base_video: str,
    segment_details: list[tuple[str, int]],
    all_artifacts: list[dict[str, Any]],
    transcript_cache: dict[str, Any],
    transcripts_manifest: dict[str, Any] | None = None,
) -> None:
    """Transcribe segments of a clip and write transcript files.

    ``segment_details`` is a list of (out_path, time_index) pairs;
    start/end strings are looked up via ``clip['times'][time_index]``.
    """
    if base_video not in transcript_cache:
        participant = clip.get("participant", "")
        manifest = transcripts_manifest or transcripts.load_transcripts_manifest()
        source_transcripts = manifest.get("source_transcripts", {})
        corrections = manifest.get("corrections", [])

        if participant and participant in source_transcripts:
            entry = source_transcripts[participant]
            raw_segments = entry.get("segments", [])
            corrected = transcripts.apply_corrections(raw_segments, corrections)
            transcript_cache[base_video] = transcripts.TranscriptResult(
                segments=corrected,
                language=entry.get("language", ""),
                source_file=entry.get("source_file", str(base_video)),
                model=entry.get("model", ""),
            )
        else:
            context_keywords = transcripts.get_corrections_keywords(corrections) or None
            known_terms = transcripts.get_known_terms(manifest) or None
            timeline = clip.get("source_timeline")
            if timeline:
                # Multi-video participant: transcribe all parts as one global
                # timeline so segment times match the clip artifacts.
                transcript_cache[base_video] = transcripts.transcribe_timeline(
                    timeline,
                    context_keywords=context_keywords,
                    known_terms=known_terms,
                )
            else:
                resolved = str(utils.resolve_input_path(base_video))
                transcript_cache[base_video] = transcripts.transcribe_video(
                    resolved,
                    context_keywords=context_keywords,
                    known_terms=known_terms,
                )
    full_transcript = transcript_cache[base_video]
    if not full_transcript:
        return

    ext = transcripts.get_transcript_extension()
    times = clip.get("times", [])

    for seg_idx, (out_path, time_idx) in enumerate(segment_details):
        start_str, end_str = times[time_idx]
        start_sec = utils.timestamp_to_seconds(start_str) or 0.0
        end_sec = utils.timestamp_to_seconds(end_str) or 0.0
        clipped = transcripts.filter_segments(
            full_transcript, start_sec, end_sec, offset_to_zero=True
        )
        t_path = files.get_unique_filename(Path(out_path).stem + ext, file_format=ext)
        if transcripts.write_transcript(clipped, t_path):
            artifact = utils.build_artifact_record(
                clip,
                base_video,
                t_path,
                start_str,
                end_str,
                artifact_type="transcript",
                seg_idx=seg_idx,
            )
            artifact["id"] += "_transcript"
            artifact["transcriptFormat"] = config.TRANSCRIBE_FORMAT
            all_artifacts.append(artifact)
        else:
            files.release_reservation(t_path)


def process_clips(
    clips_list: list[ClipRecord],
    output_format: str = "clip",
    include_severity: bool = False,
    *,
    cancel_flag: Callable[[], bool] | None = None,
    titlecards_enabled: bool | None = None,
    titlecard_duration_seconds: int | None = None,
    clear_titlecard_cache: bool = True,
    pad_pre: float = 0.0,
    pad_post: float = 0.0,
    max_duration: float = 0.0,
) -> tuple[int, list[dict[str, Any]]]:
    """Process and generate outputs from the clips list.

    Uses a three-phase approach:
    1. Sequential preparation (user prompts, fuzzy video matching)
    2. Parallel ffmpeg execution via ThreadPoolExecutor (when workers >= 2)
    3. Sequential post-processing (artifact records, transcription)

    When *cancel_flag* is supplied and returns True, the pipeline short-circuits at
    the next safe boundary. Already-completed segments are kept (each clip is a
    standalone deliverable); in-flight ffmpeg encodes are terminated and their
    partial outputs unlinked.

    *clear_titlecard_cache* controls whether the shared endcard cache is purged
    when this call finishes. Callers that fan out concurrent ``process_clips``
    invocations (e.g. Studio per-cell generation) must pass ``False`` so one
    worker does not delete endcard temp files still in use by another, then
    clear the cache once themselves after all workers complete.

    Returns:
        Tuple of (count of files generated, list of artifact records).
    """
    if config.DEBUGGING:
        config.debug_ic(len(clips_list))

    if not clips_list:
        utils.warning_print(
            "No clips to process. No timestamps were found or selected."
        )
        return (0, [])

    utils.standard_print(
        "\n* ffmpeg is set to never prompt for input and will always overwrite.\n"
        "  Only warns if close to crashing.\n"
    )

    all_artifacts: list[dict[str, Any]] = []
    fuzzy_matches: dict[str, str | None] = {}
    transcript_cache: dict[str, Any] = {}
    transcripts_manifest: dict[str, Any] | None = None  # lazy-loaded
    missing_videos: set[str] = set()

    # -- Phase 1: Sequential preparation (handles user prompts) ---------------
    prepared: list[tuple[ClipRecord, str]] = []
    skipped_no_times = 0
    skipped_no_video = 0

    global _active_progress
    progress = utils.create_progress_bar()
    if progress:
        _active_progress = progress
        with progress:
            prep_task = progress.add_task("Preparing clips", total=len(clips_list))
            for clip in clips_list:
                if cancel_flag and cancel_flag():
                    break
                clip, base_video = _prepare_and_check_clip(
                    clip, missing_videos, fuzzy_matches
                )
                if not clip["times"]:
                    skipped_no_times += 1
                elif base_video is None:
                    skipped_no_video += len(clip["times"])
                else:
                    prepared.append((clip, base_video))
                progress.update(prep_task, advance=1)
        _active_progress = None
    else:
        for clip in clips_list:
            if cancel_flag and cancel_flag():
                break
            clip, base_video = _prepare_and_check_clip(
                clip, missing_videos, fuzzy_matches
            )
            if not clip["times"]:
                skipped_no_times += 1
            elif base_video is None:
                skipped_no_video += len(clip["times"])
            else:
                prepared.append((clip, base_video))

    if not prepared:
        utils.warning_print(
            "No clips to process. No timestamps were found or selected."
        )
        if missing_videos:
            utils.standard_print(f"* Missing source video files: {len(missing_videos)}")
        return (0, [])

    # -- Phase 2: Execute ffmpeg work ------------------------------------------
    workers = _resolve_clip_workers()
    use_parallel = workers >= 2 and len(prepared) >= 2
    _EMPTY_RESULT: tuple[int, list[tuple[str, int]], bool] = (0, [], False)
    # Pre-allocate results in original order for deterministic artifact output
    results: list[tuple[int, list[tuple[str, int]], bool]] = [_EMPTY_RESULT] * len(
        prepared
    )

    if use_parallel:

        def _cut(
            pair: tuple[ClipRecord, str],
        ) -> tuple[int, list[tuple[str, int]], bool]:
            clip, base_video = pair
            return _process_single_clip_segments(
                clip,
                base_video,
                missing_videos,
                output_format=output_format,
                collect_paths=True,
                include_severity=include_severity,
                cancel_flag=cancel_flag,
                titlecards_enabled=titlecards_enabled,
                titlecard_duration_seconds=titlecard_duration_seconds,
                pad_pre=pad_pre,
                pad_post=pad_post,
                max_duration=max_duration,
            )

        def _cut_error(idx: int, exc: Exception) -> tuple[int, list[tuple[str, int]]]:
            clip, _ = prepared[idx]
            desc = (clip.get("desc") or "")[: config.PROGRESS_DESCRIPTION_LENGTH]
            utils.error_print(
                f"Clip failed: [{clip.get('participant', '')}] {desc}",
                [str(exc)],
            )
            return (0, [])

        progress = utils.create_progress_bar()
        if progress:
            with progress:
                cut_task = progress.add_task("Processing clips", total=len(prepared))
                _parallel_map_ordered(
                    prepared,
                    _cut,
                    workers=workers,
                    results=results,
                    on_error=_cut_error,
                    cancel_flag=cancel_flag,
                    on_done=lambda: progress.update(cut_task, advance=1),
                )
        else:
            _parallel_map_ordered(
                prepared,
                _cut,
                workers=workers,
                results=results,
                on_error=_cut_error,
                cancel_flag=cancel_flag,
            )
    else:
        # Sequential execution (workers=1 or single clip)
        progress = utils.create_progress_bar()
        if progress:
            with progress:
                cut_task = progress.add_task("Processing clips", total=len(prepared))
                for idx, (clip, base_video) in enumerate(prepared):
                    if cancel_flag and cancel_flag():
                        break
                    desc_preview = (clip.get("desc") or "")[
                        : config.PROGRESS_DESCRIPTION_LENGTH
                    ]
                    participant = clip.get("participant", "")
                    progress.update(
                        cut_task,
                        description=f"[{participant}] {desc_preview}...",
                    )
                    results[idx] = _process_single_clip_segments(
                        clip,
                        base_video,
                        missing_videos,
                        output_format=output_format,
                        collect_paths=True,
                        include_severity=include_severity,
                        cancel_flag=cancel_flag,
                        titlecards_enabled=titlecards_enabled,
                        titlecard_duration_seconds=titlecard_duration_seconds,
                        pad_pre=pad_pre,
                        pad_post=pad_post,
                        max_duration=max_duration,
                    )
                    progress.update(cut_task, advance=1)
        else:
            for idx, (clip, base_video) in enumerate(prepared):
                if cancel_flag and cancel_flag():
                    break
                if (
                    getattr(config, "VERBOSITY", config.STANDARD) >= config.VERBOSE
                    and len(prepared) > 1
                ):
                    utils.verbose_print(
                        f"Processing clip {idx + 1} of {len(prepared)}..."
                    )
                results[idx] = _process_single_clip_segments(
                    clip,
                    base_video,
                    missing_videos,
                    output_format=output_format,
                    collect_paths=True,
                    include_severity=include_severity,
                    cancel_flag=cancel_flag,
                    titlecards_enabled=titlecards_enabled,
                    titlecard_duration_seconds=titlecard_duration_seconds,
                    pad_pre=pad_pre,
                    pad_post=pad_post,
                    max_duration=max_duration,
                )

    # -- Phase 3: Build artifacts and transcribe (sequential) ------------------
    outputs_generated = 0
    outputs_skipped = skipped_no_times + skipped_no_video

    # Load transcripts manifest once for embedding + transcription
    transcripts_manifest = transcripts.load_transcripts_manifest()

    cards_enabled, card_duration = _resolve_titlecard_options(
        titlecards_enabled, titlecard_duration_seconds
    )

    for idx, (clip, base_video) in enumerate(prepared):
        generated_count, segment_details, cards_applied = results[idx]
        outputs_generated += generated_count
        if generated_count < len(clip["times"]):
            outputs_skipped += len(clip["times"]) - generated_count
        if segment_details:
            # Record what actually landed on disk, not what was asked for: a clip
            # whose wrap soft-failed is usable but uncarded, and claiming otherwise
            # makes the Phase-1 generate cache skip it forever.
            carded = cards_enabled and cards_applied
            title_img, end_img = _resolve_titlecard_images(carded)
            clip_artifacts = viewer.build_artifact_records_for_clip(
                clip,
                base_video,
                segment_details,
                output_format,
                titlecards=carded,
                titlecard_duration=card_duration if carded else 0,
                titlecard_image=title_img,
                endcard_image=end_img,
            )
            _embed_transcript_on_artifacts(
                clip,
                base_video,
                clip_artifacts,
                segment_details,
                transcript_cache,
                transcripts_manifest,
            )
            all_artifacts.extend(clip_artifacts)

    if config.TRANSCRIBE_ENABLED and not (cancel_flag and cancel_flag()):
        transcribe_items = [
            (clip, base_video, results[idx][1])
            for idx, (clip, base_video) in enumerate(prepared)
            if results[idx][1]
        ]
        if transcribe_items:
            progress = utils.create_progress_bar()
            if progress:
                with progress:
                    t_task = progress.add_task(
                        "Transcribing", total=len(transcribe_items)
                    )
                    for clip, base_video, segment_details in transcribe_items:
                        desc_preview = (clip.get("desc") or "")[
                            : config.PROGRESS_DESCRIPTION_LENGTH
                        ]
                        participant = clip.get("participant", "")
                        progress.update(
                            t_task,
                            description=f"[{participant}] {desc_preview}...",
                        )
                        _transcribe_segments(
                            clip,
                            base_video,
                            segment_details,
                            all_artifacts,
                            transcript_cache,
                            transcripts_manifest,
                        )
                        progress.update(t_task, advance=1)
            else:
                for clip, base_video, segment_details in transcribe_items:
                    _transcribe_segments(
                        clip,
                        base_video,
                        segment_details,
                        all_artifacts,
                        transcript_cache,
                        transcripts_manifest,
                    )

    if missing_videos:
        utils.standard_print(f"* Missing source video files: {len(missing_videos)}")

    item_name = {
        "clip": "video(s)",
        "screen": "screenshot(s)",
        "gif": "GIF(s)",
    }.get(output_format, "file(s)")
    cards_enabled, _card_duration = _resolve_titlecard_options(
        titlecards_enabled, titlecard_duration_seconds
    )
    if cards_enabled and clear_titlecard_cache:
        titlecards.clear_endcard_cache()

    if outputs_skipped > 0:
        utils.standard_print(
            f"* Summary: {outputs_generated} {item_name} generated, {outputs_skipped} skipped due to errors."
        )
    return (outputs_generated, all_artifacts)


# ---- Reel processing ----


def compute_reel_id(components: list[dict[str, Any]]) -> str:
    """Compute a deterministic reel ID from its component metadata."""
    parts = sorted(
        f"{c['cellRow']}:{c['cellCol']}:{c['start']}:{c['end']}" for c in components
    )
    return "reel_" + hashlib.sha256("|".join(parts).encode()).hexdigest()[:8]


def _build_reel_transcript(
    components: list[dict[str, Any]],
    *,
    titlecards_enabled: bool | None = None,
    titlecard_duration_seconds: int | None = None,
) -> list[dict[str, Any]]:
    """Assemble merged transcript for a reel from its components.

    Segments are derived from each component's source transcript,
    with timestamps offset by cumulative component durations + titlecard durations.
    """
    manifest = transcripts.load_transcripts_manifest()
    source_transcripts = manifest.get("source_transcripts", {})
    corrections = manifest.get("corrections", [])

    merged_segments: list[dict[str, Any]] = []
    cumulative_offset = 0.0
    seg_counter = 0
    cards_enabled, titlecard_duration = _resolve_titlecard_options(
        titlecards_enabled, titlecard_duration_seconds
    )
    # Each wrapped clip is titlecard + clip body + endcard (wrap_clip_with_cards),
    # so the per-component span the next component starts after is
    # titlecard + clip + endcard. The endcard is skipped only when ENDCARD_IMAGE
    # is the "none" sentinel; mirror that decision via resolve_card_background.
    endcard_duration = 0
    if not cards_enabled:
        titlecard_duration = 0
    else:
        _bg, _allow, skip_end, _color = titlecards.resolve_card_background("end")
        endcard_duration = 0 if skip_end else titlecard_duration

    for comp in components:
        participant = comp.get("participant", "")
        comp_start = comp.get("start", 0.0)
        comp_end = comp.get("end", 0.0)
        comp_duration = comp_end - comp_start

        full_transcript = None
        if participant and participant in source_transcripts:
            entry = source_transcripts[participant]
            raw_segments = entry.get("segments", [])
            corrected = transcripts.apply_corrections(raw_segments, corrections)
            full_transcript = transcripts.TranscriptResult(
                segments=corrected,
                language=entry.get("language", ""),
                source_file=entry.get("source_file", ""),
                model=entry.get("model", ""),
            )

        if full_transcript:
            clipped = transcripts.filter_segments(
                full_transcript, comp_start, comp_end, offset_to_zero=True
            )
            for seg in clipped["segments"]:
                merged_segments.append(
                    {
                        "id": f"reel:{seg_counter}",
                        "start": seg["start"] + cumulative_offset + titlecard_duration,
                        "end": seg["end"] + cumulative_offset + titlecard_duration,
                        "text": seg["text"],
                    }
                )
                seg_counter += 1

        cumulative_offset += comp_duration + titlecard_duration + endcard_duration

    return merged_segments


def process_reel(
    clips_list: list[ClipRecord],
    output_file: str | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    progress_cb: Callable[[dict[str, Any]], None] | None = None,
    titlecards_enabled: bool | None = None,
    titlecard_duration_seconds: int | None = None,
    pad_pre: float = 0.0,
    pad_post: float = 0.0,
    max_duration: float = 0.0,
) -> tuple[int, list[dict[str, Any]]]:
    """Process clips for reel mode: generate individual clips, concatenate into one video, clean up.

    When *progress_cb* is supplied, fires dict-shaped progress events as work
    advances:

      ``{"phase": "start", "total_clips": N}``
      ``{"phase": "clip_done", "clip_index": i, "total_clips": N}``
      ``{"phase": "concat", "progress": 0.0..1.0}`` (throttled)
      ``{"phase": "done"}``

    Callers (the Flask reel routes) yield these as NDJSON to the frontend.

    Returns:
        Tuple of (1 if reel generated successfully else 0, reel records list).
        Each reel record contains an ``id``, ``file``, ``study``, ``description``,
        and an ordered ``components`` list with per-segment metadata for regeneration.
    """
    try:
        return _process_reel(
            clips_list,
            output_file,
            cancel_flag=cancel_flag,
            progress_cb=progress_cb,
            titlecards_enabled=titlecards_enabled,
            titlecard_duration_seconds=titlecard_duration_seconds,
            pad_pre=pad_pre,
            pad_post=pad_post,
            max_duration=max_duration,
        )
    finally:
        # Endcard temp files are cached per-process across every wrap call;
        # purge them so per-request cards don't leak between reel builds. The
        # CLI, interactive, and Studio /api/reel paths all route through here,
        # mirroring the cleanup process_clips and /api/reel-direct already do.
        titlecards.clear_endcard_cache()


def _process_reel(
    clips_list: list[ClipRecord],
    output_file: str | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    progress_cb: Callable[[dict[str, Any]], None] | None = None,
    titlecards_enabled: bool | None = None,
    titlecard_duration_seconds: int | None = None,
    pad_pre: float = 0.0,
    pad_post: float = 0.0,
    max_duration: float = 0.0,
) -> tuple[int, list[dict[str, Any]]]:
    """Reel build implementation; see process_reel for the public contract."""
    if not clips_list:
        utils.warning_print(
            "No clips to process for reel. No timestamps were found or selected."
        )
        return (0, [])

    def _emit(event: dict[str, Any]) -> None:
        if progress_cb is not None:
            try:
                progress_cb(event)
            except Exception as exc:
                utils.debug_print(f"reel progress_cb raised: {exc}")

    _emit({"phase": "start", "total_clips": len(clips_list)})

    study_name = ""
    for clip in clips_list:
        s = (clip.get("study") or "").strip()
        if s:
            study_name = s
            break

    fuzzy_matches: dict[str, str | None] = {}

    def process_reel_clip(
        clip: Any, missing_videos: set[str]
    ) -> tuple[list[tuple[str, int]], list[dict[str, Any]], list[str], bool]:
        """Process one clip for reel mode.

        Returns ``(segment_paths, component_dicts, failures, cards_applied)``.
        *failures* carries a human-readable line per segment that could not be
        produced, so the caller can abort rather than silently concatenating the
        survivors into a short reel. *cards_applied* reports whether every part this
        clip contributed actually got its cards — the reel record must persist that,
        not the requested flag, or the generate cache skips retrying a reel whose
        parts are missing their cards.
        """
        clip, base_video = _prepare_and_check_clip(clip, missing_videos, fuzzy_matches)
        # `times` is populated by prepare_clip, so the expected segment count is
        # only knowable after the call above.
        expected = len(clip.get("times") or [])
        label = (
            " ".join(
                part
                for part in (
                    (clip.get("participant") or "").strip(),
                    (clip.get("desc") or "").strip()[
                        : config.PROGRESS_DESCRIPTION_LENGTH
                    ],
                )
                if part
            )
            or "(unnamed clip)"
        )
        if base_video is None:
            # A row with no timestamps is not a failure — nothing was requested of
            # it. A missing source video when timestamps *were* requested is.
            if not expected:
                return ([], [], [], True)
            return ([], [], [f"{label} — source video not found"], False)
        _, segment_paths, cards_applied = _process_single_clip_segments(
            clip,
            base_video,
            missing_videos,
            filename_prefix="_reel_part_",
            collect_paths=True,
            enforce_size=False,  # parts are concatenated into an uncapped reel
            cancel_flag=cancel_flag,
            titlecards_enabled=titlecards_enabled,
            titlecard_duration_seconds=titlecard_duration_seconds,
            pad_pre=pad_pre,
            pad_post=pad_post,
            max_duration=max_duration,
        )
        times = clip.get("times", [])
        clip_components = [
            utils.build_reel_component(clip, base_video, *times[time_idx])
            for _out_path, time_idx in segment_paths
        ]
        failures: list[str] = []
        if len(segment_paths) < expected:
            failures.append(
                f"{label} — {expected - len(segment_paths)} of {expected} "
                "segment(s) failed to cut"
            )
        return (segment_paths, clip_components, failures, cards_applied)

    def _on_clip_complete(done: int, total: int) -> None:
        _emit({"phase": "clip_done", "clip_index": done - 1, "total_clips": total})

    all_results, _ = _run_clip_pipeline(
        clips_list,
        empty_warning="No clips to process for reel. No timestamps were found or selected.",
        intro_message="* Reel mode: generating individual clips, then concatenating into one file.",
        task_label="Generating reel clips",
        per_clip_fn=process_reel_clip,
        parallel=True,
        cancel_flag=cancel_flag,
        on_clip_complete=_on_clip_complete,
    )

    # If cancelled, clean up any partial clip files and bail out
    if cancel_flag and cancel_flag():
        for segment_paths, _, _, _ in all_results:
            for entry in segment_paths:
                try:
                    Path(entry[0]).unlink(missing_ok=True)
                except OSError:
                    pass
        return (0, [])

    # Assemble ordered paths and components from per-clip results
    components: list[dict[str, Any]] = []
    clip_paths = []
    reel_failures: list[str] = []
    # One part whose wrap soft-failed makes the whole reel not-carded: the
    # concatenated output really is missing that card, so recording it as carded
    # would make the generate cache skip the rebuild (see process_clips).
    all_parts_carded = True
    for segment_paths, clip_components, clip_failures, clip_carded in all_results:
        for entry in segment_paths:
            clip_paths.append(entry[0])
        components.extend(clip_components)
        reel_failures.extend(clip_failures)
        if not clip_carded:
            all_parts_carded = False

    # A reel is a single deliverable: concatenating only the clips that happened to
    # succeed produces a silently short video that looks complete, and its reel id
    # is hashed from the truncated component list, so the cache would then serve
    # that truncated reel even after the problem is fixed. Abort instead.
    if reel_failures:
        for path in clip_paths:
            try:
                Path(path).unlink(missing_ok=True)
            except OSError:
                pass
        files.release_reservation(output_file)
        utils.error_print(
            f"Reel aborted: {len(reel_failures)} clip(s) could not be generated.",
            reel_failures
            + [
                (
                    "No reel was written — a partial reel would look complete but "
                    "silently omit these moments."
                ),
                (
                    "Fix the sources (or deselect those clips) and run again; the "
                    "clips that did cut are not reused, so nothing is left stale."
                ),
            ],
        )
        return (0, [])

    if not clip_paths:
        utils.warning_print("No clips were generated for the reel.")
        # Reclaim a caller-supplied output reservation we'll never fill.
        files.release_reservation(output_file)
        return (0, [])

    if output_file is None and study_name:
        output_file = files.get_unique_filename(f"{study_name}_reel{config.FILEFORMAT}")
    elif output_file is None:
        output_file = files.get_unique_filename(f"reel{config.FILEFORMAT}")

    # Check cancel flag before starting concatenation
    if cancel_flag and cancel_flag():
        for path in clip_paths:
            try:
                Path(path).unlink(missing_ok=True)
            except OSError:
                pass
        files.release_reservation(output_file)
        return (0, [])

    # Throttle concat progress events to ~5 Hz; ffmpeg's default -progress
    # cadence (1/sec) is already low, but the throttle keeps things bounded
    # if ffmpeg emits faster on short clips.
    last_emit_ts: list[float] = [0.0]

    def _on_concat_progress(fraction: float) -> None:
        now = time.monotonic()
        # Always emit the first and the final (>=0.99) update; throttle the rest.
        if last_emit_ts[0] != 0.0 and fraction < 0.99 and now - last_emit_ts[0] < 0.2:
            return
        last_emit_ts[0] = now
        _emit({"phase": "concat", "progress": fraction})

    def _concat() -> bool:
        return video.concatenate_clips(
            clip_paths,
            output_file,
            reencode_on_fail=True,
            cancel_flag=cancel_flag,
            on_progress=_on_concat_progress if progress_cb is not None else None,
        )

    ok = (
        utils.run_with_spinner("Concatenating clips into final reel...", _concat)
        if utils.use_progress()
        else _concat()
    )

    # If cancelled during concatenation, clean up output and temp clips
    if cancel_flag and cancel_flag():
        try:
            Path(output_file).unlink(missing_ok=True)
        except OSError:
            pass
        for path in clip_paths:
            try:
                Path(path).unlink(missing_ok=True)
            except OSError:
                pass
        return (0, [])

    for path in clip_paths:
        try:
            clip_path = Path(path)
            if clip_path.is_file():
                clip_path.unlink()
        except OSError as e:
            utils.warning_print(
                f"Could not remove temporary reel clip: {path}", [str(e)]
            )

    if not ok:
        # Concatenation failed; the output is empty or partial and useless —
        # drop it whether we or the caller reserved the name.
        files.release_reservation(output_file)
        return (0, [])

    cards_enabled, card_duration = _resolve_titlecard_options(
        titlecards_enabled, titlecard_duration_seconds
    )
    # Persist the cards that actually landed on the parts, not the ones asked
    # for: a reel whose part wraps soft-failed is genuinely missing those cards,
    # and claiming otherwise makes the generate cache skip rebuilding it.
    reel_carded = cards_enabled and all_parts_carded
    reel_id = compute_reel_id(components)
    title_img, end_img = _resolve_titlecard_images(reel_carded)
    reel_record: dict[str, Any] = {
        "id": reel_id,
        "file": Path(output_file).name,
        "study": study_name,
        "description": f"Reel: {len(components)} segments",
        "components": components,
        "titlecards": reel_carded,
        "titlecardDuration": card_duration if reel_carded else 0,
        "titlecardImage": title_img,
        "endcardImage": end_img,
    }

    reel_transcript = _build_reel_transcript(
        components,
        titlecards_enabled=titlecards_enabled,
        titlecard_duration_seconds=titlecard_duration_seconds,
    )
    if reel_transcript:
        reel_record["transcript"] = reel_transcript

    _emit({"phase": "done"})
    return (1, [reel_record])


# ---- Manifest regeneration ----


def _regenerate_batch(
    items: list[dict[str, Any]],
    worker_fn: Callable[[dict[str, Any]], bool],
    *,
    describe: Callable[[dict[str, Any]], str],
    error_prefix: str,
    workers: int,
    progress: Any = None,
    task: Any = None,
) -> int:
    """Regenerate a batch of items (media artifacts or reels).

    Runs in parallel when ``workers >= 2`` and there are at least 2 items, else
    sequentially. Advances ``task`` per item when a progress bar is active.
    Returns the count of truthy ``worker_fn`` results.
    """
    if not items:
        return 0

    if workers >= 2 and len(items) >= 2:
        results: list[bool] = [False] * len(items)

        def _on_error(idx: int, exc: Exception) -> bool:
            utils.error_print(f"{error_prefix}: {describe(items[idx])}", [str(exc)])
            return False

        _parallel_map_ordered(
            items,
            worker_fn,
            workers=workers,
            results=results,
            on_error=_on_error,
            on_done=(
                (lambda: progress.update(task, advance=1))
                if progress is not None
                else None
            ),
        )
        return sum(1 for ok in results if ok)

    count = 0
    for item in items:
        if progress is not None:
            progress.update(task, description=describe(item))
        if worker_fn(item):
            count += 1
        if progress is not None:
            progress.update(task, advance=1)
    return count


def regenerate_from_manifest(
    artifacts: list[dict[str, Any]],
    reels: list[dict[str, Any]] | None = None,
) -> int:
    """Regenerate media artifacts and reels from manifest entries.

    Skips transcript-type artifacts. For each clip/screen/gif artifact,
    resolves the source video, converts start/end seconds to timestamps,
    and invokes the appropriate ffmpeg operation. For each reel, regenerates
    component clips then concatenates them.

    Returns the number of successfully regenerated items.
    """
    media = [a for a in artifacts if a.get("type") != "transcript"]
    reel_list = reels or []
    total = len(media) + len(reel_list)
    if total == 0:
        utils.warning_print("No media artifacts or reels to regenerate.")
        return 0

    utils.print_mode_heading("Regenerating artifacts", "mode.regenerate")
    missing_videos: set[str] = set()
    missing_lock = threading.Lock()
    workers = _resolve_clip_workers()
    parallel_reels = workers >= 2 and len(reel_list) >= 2

    def _regen_artifact(artifact: dict[str, Any]) -> bool:
        local_missing: set[str] = set()
        ok = _regenerate_single_artifact(artifact, local_missing)
        if local_missing:
            with missing_lock:
                missing_videos.update(local_missing)
        return ok

    def _regen_reel(reel: dict[str, Any]) -> bool:
        local_missing: set[str] = set()
        # When the reel batch itself runs concurrently (parallel_reels), cut each
        # reel's components sequentially to avoid nested CPU oversubscription;
        # otherwise let a lone/large reel parallelize its own segment cuts.
        ok = _regenerate_reel(reel, local_missing, parallel=not parallel_reels)
        if local_missing:
            with missing_lock:
                missing_videos.update(local_missing)
        return ok

    def _describe_artifact(a: dict[str, Any]) -> str:
        desc = (a.get("description") or "")[: config.PROGRESS_DESCRIPTION_LENGTH]
        return f"[{a.get('participant', '')}] {desc}..."

    def _describe_reel(r: dict[str, Any]) -> str:
        return r.get("description", "Reel")[: config.PROGRESS_DESCRIPTION_LENGTH]

    def _run(progress: Any, task: Any) -> int:
        return _regenerate_batch(
            media,
            _regen_artifact,
            describe=_describe_artifact,
            error_prefix="Regenerate failed",
            workers=workers,
            progress=progress,
            task=task,
        ) + _regenerate_batch(
            reel_list,
            _regen_reel,
            describe=_describe_reel,
            error_prefix="Reel regen failed",
            workers=workers,
            progress=progress,
            task=task,
        )

    progress = utils.create_progress_bar()
    if progress:
        with progress:
            task = progress.add_task("Regenerating", total=total)
            generated = _run(progress, task)
    else:
        generated = _run(None, None)

    if missing_videos:
        utils.standard_print(f"* Missing source video files: {len(missing_videos)}")
    return generated


def _reapply_titlecards(artifact: dict[str, Any], output_path: str) -> bool:
    """Reapply titlecards to a regenerated clip when the manifest entry used them.

    Returns False unless the cards actually landed. The manifest entry claims this
    artifact is carded, so a soft wrap failure means the regenerated file no longer
    matches its own record — report that as a failed regeneration rather than
    silently writing an uncarded clip under a carded manifest entry.
    """
    clip: ClipRecord = {"desc": artifact.get("description", "")}
    clip_ok, cards_applied = titlecards.wrap_clip_with_cards(
        clip,
        output_path,
        titlecards_enabled=True,
        titlecard_duration_seconds=(
            artifact.get("titlecardDuration") or config.TITLECARD_DURATION_SECONDS
        ),
    )
    return clip_ok and cards_applied


def _regenerate_single_artifact(
    artifact: dict[str, Any], missing_videos: set[str]
) -> bool:
    """Regenerate one artifact from its manifest entry. Returns True on success.

    Cuts from the artifact's mapped ``sourceVideo`` using ``localStart``/
    ``localEnd`` (the offsets local to that sub-video; equal to the global
    ``start``/``end`` for single-video participants). A boundary-spanning clip
    carries a ``parts`` list — each part is re-cut from its sub-video and
    stitched back together.
    """
    output_path = str(utils.resolve_output_path(artifact.get("file", "")))
    artifact_type = artifact.get("type", "clip")

    parts = artifact.get("parts")
    if artifact_type == "clip" and parts:
        temp_paths: list[str] = []
        for n, part in enumerate(parts):
            part_source = part.get("sourceVideo", "")
            part_path = str(utils.resolve_input_path(part_source))
            if not Path(part_path).is_file():
                if part_path not in missing_videos:
                    missing_videos.add(part_path)
                    utils.warning_print(f"Source video not found: '{part_source}'")
                for p in temp_paths:
                    Path(p).unlink(missing_ok=True)
                return False
            tmp = files.get_unique_filename(f"_multipart_{n + 1}{config.FILEFORMAT}")
            if video.run_ffmpeg(
                input_file=part_path,
                output_file=tmp,
                start_pos=utils.seconds_to_timestamp(int(part.get("localStart", 0))),
                end_pos=utils.seconds_to_timestamp(int(part.get("localEnd", 0))),
                reencode=config.REENCODING,
            ):
                temp_paths.append(tmp)
            else:
                files.release_reservation(tmp)
                for p in temp_paths:
                    Path(p).unlink(missing_ok=True)
                return False
        ok = video.concatenate_clips(temp_paths, output_path, reencode_on_fail=True)
        for p in temp_paths:
            Path(p).unlink(missing_ok=True)
        if ok and artifact.get("titlecards"):
            ok = _reapply_titlecards(artifact, output_path)
        if ok:
            video.enforce_filesize_limit(output_path)
        return ok

    source_name = artifact.get("sourceVideo", "")
    if not source_name:
        utils.warning_print(
            f"Artifact '{artifact.get('file', '?')}' has no sourceVideo, skipping."
        )
        return False

    source_path = str(utils.resolve_input_path(source_name))
    if not Path(source_path).is_file():
        if source_path not in missing_videos:
            missing_videos.add(source_path)
            utils.warning_print(f"Source video not found: '{source_name}'")
        return False

    local_start = artifact.get("localStart", artifact.get("start", 0))
    local_end = artifact.get("localEnd", artifact.get("end", 0))
    start_ts = utils.seconds_to_timestamp(int(local_start))
    end_ts = utils.seconds_to_timestamp(int(local_end))

    if artifact_type == "clip":
        ok = video.run_ffmpeg(
            input_file=source_path,
            output_file=output_path,
            start_pos=start_ts,
            end_pos=end_ts,
            reencode=config.REENCODING,
        )
        # Reapply titlecards if the manifest entry was generated with them.
        if ok and artifact.get("titlecards"):
            ok = _reapply_titlecards(artifact, output_path)
        if ok:
            video.enforce_filesize_limit(output_path)
        return ok
    elif artifact_type == "screen":
        return video.extract_screenshot(
            input_file=source_path,
            output_file=output_path,
            timestamp=start_ts,
        )
    elif artifact_type == "gif":
        duration = max(
            int(local_end - local_start), config.DEFAULT_GIF_DURATION_SECONDS
        )
        return video.extract_gif(
            input_file=source_path,
            output_file=output_path,
            timestamp=start_ts,
            duration_seconds=duration,
        )
    else:
        utils.warning_print(
            f"Unknown artifact type '{artifact_type}' for '{artifact.get('file', '?')}', skipping."
        )
        return False


def _regenerate_reel(
    reel: dict[str, Any], missing_videos: set[str], parallel: bool = False
) -> bool:
    """Regenerate a reel from its manifest entry by cutting components then concatenating.

    A boundary-spanning component carries a ``parts`` list; each part is a
    separate cut and the parts simply become consecutive entries. Single-segment
    components cut directly from their mapped sub-video using local offsets. Any
    missing source or failed cut aborts the whole reel (matching the
    multipart-clip path) rather than silently producing a shorter reel.

    When *parallel* is True (a single large reel being regenerated on its own),
    the independent segment cuts run through a ThreadPoolExecutor, mirroring the
    fresh-build reel path. Callers that already parallelize *across* reels pass
    ``parallel=False`` so the two levels don't oversubscribe the CPU.
    """
    components = reel.get("components", [])
    if not components:
        return False

    # Pre-flight: flatten segments into an ordered task list and validate every
    # source up front (cheap stat calls, single-threaded) before spending any
    # ffmpeg time. A missing source aborts the whole reel without reserving or
    # cutting anything.
    tasks: list[dict[str, Any]] = []
    for comp in components:
        for segment in comp.get("parts") or [comp]:
            source = segment.get("sourceVideo", "")
            source_path = str(utils.resolve_input_path(source))
            if not Path(source_path).is_file():
                if source_path not in missing_videos:
                    missing_videos.add(source_path)
                    utils.warning_print(f"Source video not found: '{source}'")
                return False
            local_start = segment.get("localStart", segment.get("start", 0))
            local_end = segment.get("localEnd", segment.get("end", 0))
            tasks.append(
                {
                    "idx": len(tasks),
                    "source_path": source_path,
                    # Match _local_timestamp's rounding + force_hours so a
                    # regenerated segment reproduces the original cut instead
                    # of truncating up to ~1s off it.
                    "start_ts": utils.seconds_to_timestamp(
                        round(local_start), force_hours=True
                    ),
                    "end_ts": utils.seconds_to_timestamp(
                        round(local_end), force_hours=True
                    ),
                }
            )

    if not tasks:
        return False

    def _cut_segment(task: dict[str, Any]) -> tuple[int, str, bool]:
        """Reserve a part name and cut one segment; returns (idx, path, ok)."""
        out_name = files.get_unique_filename(
            f"_reel_part_{task['idx'] + 1}{config.FILEFORMAT}"
        )
        ok = video.run_ffmpeg(
            input_file=task["source_path"],
            output_file=out_name,
            start_pos=task["start_ts"],
            end_pos=task["end_ts"],
            reencode=config.REENCODING,
        )
        return (task["idx"], out_name, ok)

    workers = _resolve_clip_workers()
    cut_results: dict[int, tuple[str, bool]] = {}
    # Second, separate pool from _parallel_map_ordered; same two labels (see there).
    _cut_timed = profiling.timed("pipeline.clip")(_cut_segment)
    if parallel and len(tasks) >= 2 and workers >= 2:
        with (
            profiling.span("pipeline.pool_wall"),
            concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool,
        ):
            for future in concurrent.futures.as_completed(
                [pool.submit(_cut_timed, t) for t in tasks]
            ):
                idx, out_name, ok = future.result()
                cut_results[idx] = (out_name, ok)
    else:
        for t in tasks:
            idx, out_name, ok = _cut_segment(t)
            cut_results[idx] = (out_name, ok)

    # Reassemble in concat order regardless of completion order.
    ordered = [cut_results[t["idx"]] for t in tasks]
    all_names = [name for name, _ok in ordered]

    def _cleanup() -> None:
        for p in all_names:
            try:
                Path(p).unlink(missing_ok=True)
            except OSError:
                pass

    if not all(ok for _name, ok in ordered):
        _cleanup()  # any failed cut aborts the whole reel
        return False

    output_file = str(utils.resolve_output_path(reel.get("file", "reel.mp4")))
    ok = video.concatenate_clips(all_names, output_file, reencode_on_fail=True)

    _cleanup()
    return ok
