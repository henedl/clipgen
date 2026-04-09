# -*- coding: utf-8 -*-
"""Clip processing pipeline for clipgen.

Reusable functions for generating clips, reels, screenshots, GIFs, and
regenerating artifacts from manifests.  Extracted from clipgen.py so that
the web layer (server.py) can use the pipeline without importing the CLI
entry-point module.

Public API:
    is_excel_worksheet(worksheet) -> bool
    process_clips(clips_list, output_format, include_severity) -> (count, artifacts)
    process_reel(clips_list, output_file) -> (count, reel_records)
    compute_reel_id(components) -> str
    regenerate_from_manifest(artifacts, reels) -> int
"""

import concurrent.futures
import difflib
import hashlib
import os
from pathlib import Path
from typing import Any, Callable

from icecream import ic

import config
import files
import titlecards
import transcripts
import utils
import video
import viewer
from utils import ClipRecord

# Active progress bar reference, set during clip pipeline so nested functions
# (e.g. fuzzy match prompts) can pause/resume the live display.
_active_progress = None
_active_secondary_task = None


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


def _check_source_video(
    clip: ClipRecord,
    missing_videos: set[str],
    skip_detail: str,
    fuzzy_matches: dict[str, str | None],
) -> str | None:
    """Return the expected source video path if it exists; log a detailed error once per missing file.

    The expected filename is derived from clip['study'] and clip['participant'] by default,
    but can be overridden per-participant via an optional source_filename field.
    When no exact match is found, scans the input directory for large .mp4 files and
    offers the closest fuzzy match for user confirmation.
    Paths already seen in missing_videos are not reported again.
    """
    override = clip.get("source_filename")
    base_name = files.get_source_video_filename(
        clip["study"], clip["participant"], override
    )
    full_path = utils.resolve_input_path(base_name)
    if full_path.is_file():
        return str(full_path)

    full_path_str = str(full_path)

    # Check fuzzy match cache (value may be None = user rejected or no candidate)
    if full_path_str in fuzzy_matches:
        return fuzzy_matches[full_path_str]

    # Scan input directory for large .mp4 files as fuzzy candidates
    input_dir = utils.get_effective_input_dir()
    size_threshold = config.MIN_SOURCE_VIDEO_SIZE_MB * 1_000_000
    candidates = []
    for p in input_dir.glob(f"*{config.FILEFORMAT}"):
        try:
            size = p.stat().st_size
        except OSError:
            continue
        if size >= size_threshold:
            ratio = difflib.SequenceMatcher(
                None, base_name.lower(), p.name.lower()
            ).ratio()
            candidates.append((ratio, size, p))

    # Sort by similarity descending, then file size descending as tiebreaker
    candidates.sort(key=lambda c: (c[0], c[1]), reverse=True)

    if candidates and candidates[0][0] >= 0.7:
        best_ratio, best_size, best_path = candidates[0]
        size_gb = best_size / 1_000_000_000
        # Pause progress bar so the prompt is visible and input is rendered
        global _active_progress
        paused = False
        if _active_progress is not None:
            _active_progress.stop()
            paused = True
        utils.info_print(f"Source video '{base_name}' not found.")
        utils.info_print(f"Closest match found: '{best_path.name}' ({size_gb:.1f} GB)")
        answer = utils.read_user_input("Use this file instead? [y/n]\n>> ")
        if paused:
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


def _process_single_clip_segments(
    clip: ClipRecord,
    base_video: str,
    missing_videos: set[str],
    *,
    filename_prefix: str = "",
    output_format: str = "clip",
    collect_paths: bool = False,
    include_severity: bool = False,
) -> tuple[int, list[tuple[str, str, str]]]:
    """Process one clip's segments: run ffmpeg for each (start, end), optionally collect output paths.

    Caller must have already called prepare_clip(clip). Does not add to missing_videos; caller handles that.

    Args:
        clip: Prepared clip dict with 'times', 'category', 'study', 'participant', 'desc'
        base_video: Path to source video file
        missing_videos: Set of already-reported missing paths (read-only here)
        filename_prefix: Prefix for output filename (e.g. '_reel_part_' for reel)
        collect_paths: If True, return list of output paths; otherwise return empty list
        include_severity: If True and clip has severity, include [Severity] in filename

    Returns:
        (number of segments successfully generated, list of output paths if collect_paths else [])
    """
    generated = 0
    output_paths: list[tuple[str, str, str]] = []
    extension_map = {
        "clip": config.FILEFORMAT,
        "screen": ".png",
        "gif": ".gif",
    }
    file_extension = extension_map.get(output_format)
    if not file_extension:
        utils.error_print(f"Unsupported output format: '{output_format}'")
        return (generated, output_paths)

    severity_tag = (
        f"[{clip['severity']}]" if include_severity and clip.get("severity") else ""
    )
    template = f"{filename_prefix}[{clip['category']}]{severity_tag} {clip['study']} {clip['participant']} {clip['desc']}{file_extension}"
    for start_time, end_time in clip["times"]:
        try:
            out_name = files.get_unique_filename(template, file_format=file_extension)
            if config.DEBUGGING:
                ic(out_name)
        except (TypeError, UnicodeEncodeError, UnicodeDecodeError) as e:
            if config.DEBUGGING:
                ic(e, clip)
            utils.error_print(
                f"Character encoding issue occurred: {e}",
                [
                    f"Category: '{clip['category']}', Study: '{clip['study']}', Participant: '{clip['participant']}'",
                    "Try simplifying the description or category names to use only ASCII characters.",
                ],
            )
            return (generated, output_paths)
        if output_format == "clip":
            ok = video.run_ffmpeg(
                input_file=base_video,
                output_file=out_name,
                start_pos=start_time,
                end_pos=end_time,
                reencode=config.REENCODING,
            )
            clip_resolution = None
            if ok and config.TITLECARDS_ENABLED:
                props = video.probe_video_properties(out_name)
                if props:
                    clip_resolution = f"{props['width']}x{props['height']}"
                ok = titlecards.prepend_titlecard_to_clip(
                    clip, out_name, resolution=clip_resolution
                )
            if ok:
                ok = titlecards.append_endcard_to_clip(
                    out_name, resolution=clip_resolution
                )
        elif output_format == "screen":
            ok = video.extract_screenshot(
                input_file=base_video,
                output_file=out_name,
                timestamp=start_time,
            )
        else:  # output_format == 'gif'
            ok = video.extract_gif(
                input_file=base_video,
                output_file=out_name,
                timestamp=start_time,
                duration_seconds=config.DEFAULT_GIF_DURATION_SECONDS,
            )
        if ok:
            generated += 1
            if collect_paths:
                output_paths.append((out_name, start_time, end_time))
    return (generated, output_paths)


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
) -> tuple[list[Any], set[str]]:
    """Run shared clip-processing pipeline and return per-clip results.

    When *parallel* is True and there are at least 2 clips, a ThreadPoolExecutor
    is used to run ``per_clip_fn`` concurrently (same worker count as
    ``_resolve_clip_workers()``).  Results are collected in original order.
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

    if use_parallel:
        results: list[Any] = [None] * total_clips
        progress = utils.create_progress_bar()
        if progress:
            global _active_progress, _active_secondary_task
            _active_progress = progress
            with progress:
                task = progress.add_task(task_label, total=total_clips)
                with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
                    future_to_idx = {
                        pool.submit(wrapped_process, clip): idx
                        for idx, clip in enumerate(clips_list)
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
                            clip = clips_list[idx]
                            desc = (clip.get("desc") or "")[
                                : config.PROGRESS_DESCRIPTION_LENGTH
                            ]
                            utils.error_print(
                                f"Clip failed: [{clip.get('participant', '')}] {desc}",
                                [str(exc)],
                            )
                            results[idx] = []
                        progress.update(task, advance=1)
            _active_progress = None
            _active_secondary_task = None
        else:
            with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
                future_to_idx = {
                    pool.submit(wrapped_process, clip): idx
                    for idx, clip in enumerate(clips_list)
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
                        clip = clips_list[idx]
                        desc = (clip.get("desc") or "")[
                            : config.PROGRESS_DESCRIPTION_LENGTH
                        ]
                        utils.error_print(
                            f"Clip failed: [{clip.get('participant', '')}] {desc}",
                            [str(exc)],
                        )
                        results[idx] = []
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

    if missing_videos:
        utils.standard_print(f"* Missing source video files: {len(missing_videos)}")
    return (results, missing_videos)


def _embed_transcript_on_artifacts(
    clip: Any,
    base_video: str,
    artifacts: list[dict[str, Any]],
    segment_details: list[tuple[str, str, str]],
    transcript_cache: dict[str, Any],
    transcripts_manifest: dict[str, Any] | None = None,
) -> None:
    """Embed transcript segments on clip artifact records.

    Source priority: transcripts manifest -> in-memory cache. Modifies artifacts in-place.
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
    elif base_video in transcript_cache and transcript_cache[base_video]:
        full_transcript = transcript_cache[base_video]

    if not full_transcript:
        return

    for art_idx, (_out_path, start_str, end_str) in enumerate(segment_details):
        if art_idx >= len(artifacts):
            break
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
    segment_details: list[tuple[str, str, str]],
    all_artifacts: list[dict[str, Any]],
    transcript_cache: dict[str, Any],
    transcripts_manifest: dict[str, Any] | None = None,
) -> None:
    """Transcribe segments of a clip and write transcript files."""
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
            resolved = str(utils.resolve_input_path(base_video))
            context_keywords = transcripts.get_corrections_keywords(corrections) or None
            result = transcripts.transcribe_video(
                resolved, context_keywords=context_keywords
            )
            transcript_cache[base_video] = result
    full_transcript = transcript_cache[base_video]
    if not full_transcript:
        return

    ext = transcripts.get_transcript_extension()
    cell = clip.get("cell")
    cell_row = getattr(cell, "row", None)
    cell_col = getattr(cell, "col", None)
    cell_a1 = utils.safe_cell_a1(cell_row, cell_col)
    annotations = list(clip.get("cell_annotations", []))

    for seg_idx, (out_path, start_str, end_str) in enumerate(segment_details):
        start_sec = utils.timestamp_to_seconds(start_str) or 0.0
        end_sec = utils.timestamp_to_seconds(end_str) or 0.0
        clipped = transcripts.filter_segments(
            full_transcript, start_sec, end_sec, offset_to_zero=True
        )
        t_path = files.get_unique_filename(Path(out_path).stem + ext, file_format=ext)
        if transcripts.write_transcript(clipped, t_path):
            all_artifacts.append(
                {
                    "id": f"a{cell_row}c{cell_col}s{seg_idx}_transcript",
                    "type": "transcript",
                    "file": Path(t_path).name,
                    "start": start_sec,
                    "end": end_sec,
                    "thumbnail": "",
                    "study": clip.get("study", ""),
                    "participant": clip.get("participant", ""),
                    "category": clip.get("category", ""),
                    "severity": clip.get("severity", ""),
                    "description": clip.get("desc", ""),
                    "cellRow": cell_row,
                    "cellCol": cell_col,
                    "cellA1": cell_a1,
                    "annotations": annotations,
                    "sourceVideo": base_video,
                    "transcriptFormat": config.TRANSCRIBE_FORMAT,
                }
            )


def process_clips(
    clips_list: list[ClipRecord],
    output_format: str = "clip",
    include_severity: bool = False,
) -> tuple[int, list[dict[str, Any]]]:
    """Process and generate outputs from the clips list.

    Uses a three-phase approach:
    1. Sequential preparation (user prompts, fuzzy video matching)
    2. Parallel ffmpeg execution via ThreadPoolExecutor (when workers >= 2)
    3. Sequential post-processing (artifact records, transcription)

    Returns:
        Tuple of (count of files generated, list of artifact records).
    """
    if config.DEBUGGING:
        ic(len(clips_list))

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

    global _active_progress, _active_secondary_task
    progress = utils.create_progress_bar()
    if progress:
        _active_progress = progress
        with progress:
            prep_task = progress.add_task("Preparing clips", total=len(clips_list))
            for clip in clips_list:
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
    _EMPTY_RESULT: tuple[int, list[tuple[str, str, str]]] = (0, [])
    # Pre-allocate results in original order for deterministic artifact output
    results: list[tuple[int, list[tuple[str, str, str]]]] = [_EMPTY_RESULT] * len(
        prepared
    )

    if use_parallel:
        progress = utils.create_progress_bar()
        if progress:
            with progress:
                cut_task = progress.add_task("Processing clips", total=len(prepared))
                with concurrent.futures.ThreadPoolExecutor(
                    max_workers=workers,
                ) as pool:
                    future_to_idx = {
                        pool.submit(
                            _process_single_clip_segments,
                            clip,
                            base_video,
                            missing_videos,
                            output_format=output_format,
                            collect_paths=True,
                            include_severity=include_severity,
                        ): idx
                        for idx, (clip, base_video) in enumerate(prepared)
                    }
                    for future in concurrent.futures.as_completed(future_to_idx):
                        idx = future_to_idx[future]
                        try:
                            results[idx] = future.result()
                        except Exception as exc:
                            clip, _ = prepared[idx]
                            desc = (clip.get("desc") or "")[
                                : config.PROGRESS_DESCRIPTION_LENGTH
                            ]
                            utils.error_print(
                                f"Clip failed: [{clip.get('participant', '')}] {desc}",
                                [str(exc)],
                            )
                            results[idx] = (0, [])
                        progress.update(cut_task, advance=1)
        else:
            with concurrent.futures.ThreadPoolExecutor(
                max_workers=workers,
            ) as pool:
                future_to_idx = {
                    pool.submit(
                        _process_single_clip_segments,
                        clip,
                        base_video,
                        missing_videos,
                        output_format=output_format,
                        collect_paths=True,
                        include_severity=include_severity,
                    ): idx
                    for idx, (clip, base_video) in enumerate(prepared)
                }
                for future in concurrent.futures.as_completed(future_to_idx):
                    idx = future_to_idx[future]
                    try:
                        results[idx] = future.result()
                    except Exception as exc:
                        clip, _ = prepared[idx]
                        desc = (clip.get("desc") or "")[
                            : config.PROGRESS_DESCRIPTION_LENGTH
                        ]
                        utils.error_print(
                            f"Clip failed: [{clip.get('participant', '')}] {desc}",
                            [str(exc)],
                        )
                        results[idx] = (0, [])
    else:
        # Sequential execution (workers=1 or single clip)
        progress = utils.create_progress_bar()
        if progress:
            with progress:
                cut_task = progress.add_task("Processing clips", total=len(prepared))
                for idx, (clip, base_video) in enumerate(prepared):
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
                    )
                    progress.update(cut_task, advance=1)
        else:
            for idx, (clip, base_video) in enumerate(prepared):
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
                )

    # -- Phase 3: Build artifacts and transcribe (sequential) ------------------
    outputs_generated = 0
    outputs_skipped = skipped_no_times + skipped_no_video

    # Load transcripts manifest once for embedding + transcription
    transcripts_manifest = transcripts.load_transcripts_manifest()

    for idx, (clip, base_video) in enumerate(prepared):
        generated_count, segment_details = results[idx]
        outputs_generated += generated_count
        if generated_count < len(clip["times"]):
            outputs_skipped += len(clip["times"]) - generated_count
        if segment_details:
            clip_artifacts = viewer.build_artifact_records_for_clip(
                clip, base_video, segment_details, output_format
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

    if config.TRANSCRIBE_ENABLED:
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
    if config.TITLECARDS_ENABLED:
        titlecards.clear_endcard_cache()

    if outputs_skipped > 0:
        utils.standard_print(
            f"* Summary: {outputs_generated} {item_name} generated, {outputs_skipped} skipped due to errors."
        )
    return (outputs_generated, all_artifacts)


def compute_reel_id(components: list[dict[str, Any]]) -> str:
    """Compute a deterministic reel ID from its component metadata."""
    parts = sorted(
        f"{c['cellRow']}:{c['cellCol']}:{c['start']}:{c['end']}" for c in components
    )
    return "reel_" + hashlib.sha256("|".join(parts).encode()).hexdigest()[:8]


def _build_reel_transcript(
    components: list[dict[str, Any]],
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
    titlecard_duration = (
        config.TITLECARD_DURATION_SECONDS if config.TITLECARDS_ENABLED else 0
    )

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

        cumulative_offset += comp_duration + titlecard_duration

    return merged_segments


def process_reel(
    clips_list: list[ClipRecord],
    output_file: str | None = None,
    cancel_flag: Callable[[], bool] | None = None,
) -> tuple[int, list[dict[str, Any]]]:
    """Process clips for reel mode: generate individual clips, concatenate into one video, clean up.

    Returns:
        Tuple of (1 if reel generated successfully else 0, reel records list).
        Each reel record contains an ``id``, ``file``, ``study``, ``description``,
        and an ordered ``components`` list with per-segment metadata for regeneration.
    """
    if not clips_list:
        utils.warning_print(
            "No clips to process for reel. No timestamps were found or selected."
        )
        return (0, [])

    study_name = ""
    for clip in clips_list:
        s = (clip.get("study") or "").strip()
        if s:
            study_name = s
            break

    fuzzy_matches: dict[str, str | None] = {}

    def process_reel_clip(
        clip: Any, missing_videos: set[str]
    ) -> tuple[list[tuple[str, str, str]], list[dict[str, Any]]]:
        """Process one clip for reel mode and return (segment_paths, component_dicts)."""
        clip, base_video = _prepare_and_check_clip(clip, missing_videos, fuzzy_matches)
        if base_video is None:
            return ([], [])
        _, segment_paths = _process_single_clip_segments(
            clip,
            base_video,
            missing_videos,
            filename_prefix="_reel_part_",
            collect_paths=True,
        )
        clip_components = []
        for _out_path, start_str, end_str in segment_paths:
            clip_components.append(
                {
                    "cellRow": getattr(clip.get("cell"), "row", None),
                    "cellCol": getattr(clip.get("cell"), "col", None),
                    "participant": clip.get("participant", ""),
                    "sourceVideo": base_video,
                    "start": utils.timestamp_to_seconds(start_str),
                    "end": utils.timestamp_to_seconds(end_str),
                    "category": clip.get("category", ""),
                    "description": clip.get("desc", ""),
                    "severity": clip.get("severity", ""),
                }
            )
        return (segment_paths, clip_components)

    all_results, _ = _run_clip_pipeline(
        clips_list,
        empty_warning="No clips to process for reel. No timestamps were found or selected.",
        intro_message="* Reel mode: generating individual clips, then concatenating into one file.",
        task_label="Generating reel clips",
        per_clip_fn=process_reel_clip,
        parallel=True,
        cancel_flag=cancel_flag,
    )

    # If cancelled, clean up any partial clip files and bail out
    if cancel_flag and cancel_flag():
        for item in all_results:
            if item is None:
                continue
            segment_paths, _ = item
            for entry in segment_paths:
                try:
                    Path(entry[0]).unlink(missing_ok=True)
                except OSError:
                    pass
        return (0, [])

    # Assemble ordered paths and components from per-clip results
    components: list[dict[str, Any]] = []
    clip_paths = []
    for segment_paths, clip_components in all_results:
        for entry in segment_paths:
            clip_paths.append(entry[0])
        components.extend(clip_components)
    if not clip_paths:
        utils.warning_print("No clips were generated for the reel.")
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
        return (0, [])

    def _concat() -> bool:
        return video.concatenate_clips(
            clip_paths, output_file, reencode_on_fail=True, cancel_flag=cancel_flag
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
        return (0, [])

    reel_id = compute_reel_id(components)
    reel_record: dict[str, Any] = {
        "id": reel_id,
        "file": Path(output_file).name,
        "study": study_name,
        "description": f"Reel: {len(components)} segments",
        "components": components,
    }

    reel_transcript = _build_reel_transcript(components)
    if reel_transcript:
        reel_record["transcript"] = reel_transcript

    return (1, [reel_record])


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
    total = len(media) + len(reels or [])
    if total == 0:
        utils.warning_print("No media artifacts or reels to regenerate.")
        return 0

    utils.print_mode_heading("Regenerating artifacts", "mode.regenerate")
    missing_videos: set[str] = set()
    generated = 0

    progress = utils.create_progress_bar()
    if progress:
        with progress:
            task = progress.add_task("Regenerating", total=total)
            for artifact in media:
                desc_preview = (artifact.get("description") or "")[
                    : config.PROGRESS_DESCRIPTION_LENGTH
                ]
                progress.update(
                    task,
                    description=f"[{artifact.get('participant', '')}] {desc_preview}...",
                )
                if _regenerate_single_artifact(artifact, missing_videos):
                    generated += 1
                progress.update(task, advance=1)
            for reel in reels or []:
                progress.update(
                    task,
                    description=reel.get("description", "Reel")[
                        : config.PROGRESS_DESCRIPTION_LENGTH
                    ],
                )
                if _regenerate_reel(reel, missing_videos):
                    generated += 1
                progress.update(task, advance=1)
    else:
        for artifact in media:
            if _regenerate_single_artifact(artifact, missing_videos):
                generated += 1
        for reel in reels or []:
            if _regenerate_reel(reel, missing_videos):
                generated += 1

    if missing_videos:
        utils.standard_print(f"* Missing source video files: {len(missing_videos)}")
    return generated


def _regenerate_single_artifact(
    artifact: dict[str, Any], missing_videos: set[str]
) -> bool:
    """Regenerate one artifact from its manifest entry. Returns True on success."""
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

    output_path = str(utils.resolve_output_path(artifact.get("file", "")))
    start_sec = artifact.get("start", 0)
    end_sec = artifact.get("end", 0)
    start_ts = utils.seconds_to_timestamp(int(start_sec))
    end_ts = utils.seconds_to_timestamp(int(end_sec))
    artifact_type = artifact.get("type", "clip")

    if artifact_type == "clip":
        return video.run_ffmpeg(
            input_file=source_path,
            output_file=output_path,
            start_pos=start_ts,
            end_pos=end_ts,
            reencode=config.REENCODING,
        )
    elif artifact_type == "screen":
        return video.extract_screenshot(
            input_file=source_path,
            output_file=output_path,
            timestamp=start_ts,
        )
    elif artifact_type == "gif":
        duration = max(int(end_sec - start_sec), config.DEFAULT_GIF_DURATION_SECONDS)
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


def _regenerate_reel(reel: dict[str, Any], missing_videos: set[str]) -> bool:
    """Regenerate a reel from its manifest entry by cutting components then concatenating."""
    components = reel.get("components", [])
    if not components:
        return False

    temp_paths: list[str] = []
    for comp in components:
        source = comp.get("sourceVideo", "")
        source_path = str(utils.resolve_input_path(source))
        if not Path(source_path).is_file():
            if source_path not in missing_videos:
                missing_videos.add(source_path)
                utils.warning_print(f"Source video not found: '{source}'")
            continue

        start_ts = utils.seconds_to_timestamp(int(comp["start"]))
        end_ts = utils.seconds_to_timestamp(int(comp["end"]))
        out_name = files.get_unique_filename(
            f"_reel_part_{len(temp_paths) + 1}{config.FILEFORMAT}"
        )
        if video.run_ffmpeg(
            input_file=source_path,
            output_file=out_name,
            start_pos=start_ts,
            end_pos=end_ts,
            reencode=config.REENCODING,
        ):
            temp_paths.append(out_name)

    if not temp_paths:
        return False

    output_file = str(utils.resolve_output_path(reel.get("file", "reel.mp4")))
    ok = video.concatenate_clips(temp_paths, output_file, reencode_on_fail=True)

    for p in temp_paths:
        try:
            Path(p).unlink(missing_ok=True)
        except OSError:
            pass
    return ok
