# -*- coding: utf-8 -*-
"""clipgen - Video clip generator from Google Sheets timestamps.

This program will help quickly create video snippets from longer video files, based on timestamps in a spreadsheet!
Check out README.md for more detailed information about setting up and using clipgen.

Data flow: Spreadsheet -> Worksheet -> clip records (cell, desc, study, participant, category; 'times' added by prepare_clip) -> ffmpeg clips or reel.

This script supports full unicode/UTF-8 for international characters in:
- Study names
- Participant IDs
- Category names
- Descriptions
- File paths
"""

import concurrent.futures
import difflib
import hashlib
import os
import sys
from pathlib import Path
from typing import Any, Callable

import gspread
from icecream import ic

import config
import files
import google_api
import interactive
import spreadsheet
import utils
import video
import viewer
from utils import ClipRecord

# Active progress bar reference, set during clip pipeline so nested functions
# (e.g. fuzzy match prompts) can pause/resume the live display.
_active_progress = None
_active_secondary_task = None

# ---- Mode configuration ----

MODE_ALIASES = {
    "b": "batch",
    "batch": "batch",
    "l": "line",
    "line": "line",
    "r": "range",
    "range": "range",
    "c": "category",
    "cat": "category",
    "category": "category",
    "ce": "cell",
    "cell": "cell",
    "p": "participant",
    "participant": "participant",
    "k": "keyword",
    "keyword": "keyword",
    "sv": "severity",
    "severity": "severity",
    "s": "screen",
    "screen": "screen",
    "g": "gif",
    "gif": "gif",
    "re": "reel",
    "reel": "reel",
    "rl": "reellate",
    "reellate": "reellate",
    "br": "browse",
    "browse": "browse",
    "v": "viewer",
    "viewer": "viewer",
    "tv": "timeline-viewer",
    "timeline-viewer": "timeline-viewer",
    "rg": "regenerate",
    "regenerate": "regenerate",
    "gv": "gallery",
    "gallery": "gallery",
    "in": "insights",
    "insights": "insights",
    "st": "studio",
    "studio": "studio",
    "ss": "screenspace",
    "screenspace": "screenspace",
    "tr": "transcripts",
    "transcripts": "transcripts",
    "se": "settings",
    "settings": "settings",
}

FORMAT_MODE_ALIASES = {
    alias: mode
    for alias, mode in MODE_ALIASES.items()
    if mode
    in {
        "batch",
        "line",
        "range",
        "category",
        "cell",
        "participant",
        "keyword",
        "severity",
    }
}

# Dispatch table for standard interactive modes: mode -> (prompt_fn, generate_fn).
# prompt_fn(ctx) returns a result or None/False to cancel.
# generate_fn(ctx, result) returns a list of clip records.
_STANDARD_MODES = {
    "batch": (
        lambda ctx: interactive.prompt_batch_confirm(ctx),
        lambda ctx, _: spreadsheet.generate_batch_timestamps(ctx),
    ),
    "line": (
        lambda ctx: interactive.prompt_line_selection(ctx),
        lambda ctx, v: spreadsheet.generate_line_timestamps(ctx, v),
    ),
    "range": (
        lambda ctx: interactive.prompt_range_selection(ctx),
        lambda ctx, v: spreadsheet.generate_range_timestamps(ctx, *v),
    ),
    "category": (
        lambda ctx: interactive.prompt_category_selection(ctx),
        lambda ctx, v: spreadsheet.generate_category_timestamps(ctx, v),
    ),
    "cell": (
        lambda ctx: interactive.prompt_cell_selection(ctx),
        lambda ctx, v: spreadsheet.generate_cell_timestamps(ctx, v),
    ),
    "participant": (
        lambda ctx: interactive.prompt_participant_selection(ctx),
        lambda ctx, pids: [
            c
            for pid in pids
            for c in spreadsheet.generate_participant_timestamps(ctx, pid)
        ],
    ),
    "keyword": (
        lambda ctx: interactive.prompt_keyword_selection(ctx),
        lambda ctx, v: spreadsheet.generate_keyword_timestamps(ctx, annotation_ids=v),
    ),
    "severity": (
        lambda ctx: interactive.prompt_severity_selection(ctx),
        lambda ctx, v: spreadsheet.generate_severity_timestamps(ctx, v),
    ),
}


# ---- Spreadsheet opening and selection ----


def _open_worksheet(open_callable: Callable[[], Any], error_context: str) -> Any | None:
    """Try to open a worksheet via a callable; catch gspread errors and print a consistent message."""
    try:
        return google_api.get_worksheet(open_callable())
    except (
        gspread.SpreadsheetNotFound,
        gspread.exceptions.APIError,
        gspread.exceptions.GSpreadException,
    ) as e:
        utils.error_print(f"Could not open spreadsheet {error_context}: {e}")
        return None


def open_spreadsheet_by_url(
    gspread_client: Any, url: str, *, use_spinner: bool = False
) -> Any | None:
    """Open a spreadsheet by URL."""

    def open_fn() -> Any | None:
        return _open_worksheet(lambda: gspread_client.open_by_url(url), "by URL")

    if use_spinner:
        return utils.run_with_spinner("Opening document by URL...", open_fn)
    return open_fn()


def open_spreadsheet_by_index(
    gspread_client: Any, doc_list: list[str], index: int, *, use_spinner: bool = False
) -> Any | None:
    """Open a spreadsheet by 1-based index number from the document list."""
    if index < 1 or index > len(doc_list):
        utils.error_print(
            f"Invalid index {index}. Must be between 1 and {len(doc_list)}"
        )
        return None
    doc_name = doc_list[index - 1].strip()
    if not use_spinner:
        utils.standard_print(f"Opening document: {doc_name}")

    def open_fn() -> Any | None:
        return _open_worksheet(
            lambda: gspread_client.open(doc_name), f"at index {index}"
        )

    if use_spinner:
        return utils.run_with_spinner(f"Opening document: {doc_name}...", open_fn)
    return open_fn()


def open_spreadsheet_by_name(
    gspread_client: Any,
    doc_list: list[str],
    name: str,
    *,
    use_spinner: bool = False,
    prompt_prefix: str = "No exact match found. Did you mean",
) -> Any | None:
    """Open a spreadsheet by name search against the document list."""
    chosen_index = google_api.find_spreadsheet_by_name(name, doc_list)
    if chosen_index < 0:
        suggestion = utils.suggest_close_match(
            name,
            doc_list,
            prompt_prefix=prompt_prefix,
        )
        if suggestion is None:
            return None
        chosen_index = doc_list.index(suggestion)
    matched_name = doc_list[chosen_index].strip()
    if not use_spinner:
        utils.standard_print(f"Opening document: {matched_name}")

    def open_fn() -> Any | None:
        return _open_worksheet(lambda: gspread_client.open(matched_name), f"'{name}'")

    if use_spinner:
        return utils.run_with_spinner(f"Opening document: {matched_name}...", open_fn)
    return open_fn()


def _handle_spreadsheet_command(
    gspread_client: Any, doc_list: list[str], input_name: str
) -> Any | None:
    """Handle one spreadsheet selection command. Returns worksheet when one was opened, None to show prompt again."""
    if not input_name:
        return None
    # Handle 'excel' for local .xlsx
    if input_name.strip().lower() == config.COMMAND_EXCEL:
        import excel_io

        return excel_io.select_excel_file()
    # Handle URL
    if input_name.startswith(config.COMMAND_HTTP_PREFIX):
        return open_spreadsheet_by_url(gspread_client, input_name, use_spinner=True)
    # Handle 'all' command
    if input_name.startswith(config.COMMAND_LIST_ALL):
        utils.info_print("Available documents:")
        for i, doc in enumerate(doc_list):
            utils.info_print(f"{i + 1}. {doc.strip()}")
        return None
    # Handle 'new' command
    if input_name.startswith(config.COMMAND_LIST_NEW):
        utils.info_print("Newest documents: (modified or opened most recently)")
        for i in range(min(config.NUM_NEWEST_DOCS_TO_SHOW, len(doc_list))):
            utils.info_print(f"{i + 1}. {doc_list[i].strip()}")
        return None
    # Handle 'last' command
    if input_name.startswith(config.COMMAND_OPEN_LAST):
        latest_spreadsheet_name = google_api.get_all_spreadsheets(gspread_client)[0]
        return open_spreadsheet_by_name(
            gspread_client, doc_list, latest_spreadsheet_name, use_spinner=True
        )
    # Handle numeric index
    if input_name[0].isdigit():
        return open_spreadsheet_by_index(
            gspread_client, doc_list, int(input_name), use_spinner=True
        )
    # Handle 'settings' command
    if input_name.startswith(config.COMMAND_SETTINGS):
        utils.set_program_settings()
        return None
    # Handle name search
    return open_spreadsheet_by_name(
        gspread_client, doc_list, input_name, use_spinner=True
    )


def select_spreadsheet(gspread_client: Any, doc_list: list[str]) -> Any:
    """Interactive spreadsheet selection. Returns the selected worksheet."""
    consecutive_open_failures = 0
    utils.print_mode_heading("Spreadsheet selection", "mode.spreadsheet")

    while True:
        try:
            input_name = utils.read_user_input(
                f"\nPlease enter the index, name, URL, or '{config.COMMAND_EXCEL}' for local file "
                f"('{config.COMMAND_LIST_ALL}' for list, '{config.COMMAND_LIST_NEW}' for list of newest, "
                f"'{config.COMMAND_OPEN_LAST}' to immediately open latest, '{config.COMMAND_SETTINGS}' to change settings):\n>> "
            )
        except utils.TopToSpreadsheet:
            # Already at spreadsheet selection; bubble up so main can restart selection loop.
            raise
        except utils.BackToModeSelection:
            # From this context, going "back" is equivalent to going to spreadsheet selection.
            raise utils.TopToSpreadsheet()
        try:
            worksheet = _handle_spreadsheet_command(
                gspread_client, doc_list, input_name
            )
            if worksheet is not None:
                return worksheet
        except (
            gspread.SpreadsheetNotFound,
            gspread.exceptions.APIError,
            gspread.exceptions.GSpreadException,
        ) as e:
            consecutive_open_failures += 1
            if consecutive_open_failures == 1:
                utils.error_print(
                    f"Could not access spreadsheet: {e}",
                    [
                        f"Please try again. Type '{config.COMMAND_LIST_ALL}' to see available documents."
                    ],
                )
            elif consecutive_open_failures == 2:
                utils.error_print(
                    "Spreadsheet not found or not accessible.",
                    [
                        "Common causes:",
                        "  - The spreadsheet name is misspelled",
                        "  - The spreadsheet hasn't been shared with your service account",
                        "    (Share it with the email in credentials.json 'client_email' field)",
                        "  - The spreadsheet doesn't contain any worksheets",
                        "",
                        f"  Type '{config.COMMAND_LIST_ALL}' to see accessible documents, or '{config.COMMAND_LIST_NEW}' for recent ones.",
                    ],
                )
            else:
                utils.error_print(
                    str(e),
                    [
                        f"Tip: Use the document index number (1, 2, 3...) from the '{config.COMMAND_LIST_ALL}' list."
                    ],
                )
        except (utils.QuitProgram, utils.TopToSpreadsheet, utils.BackToModeSelection):
            raise
        except Exception as e:
            utils.error_print(f"Could not open document: {e}")


# ---- Shared helpers ----

_SELECTION_MODE_HELP = [
    "    b or batch   - Generate all clips in the spreadsheet",
    "    r or range   - Generate clips from a range of rows",
    "    c or category - Generate clips by category",
    "    l or line    - Generate clips from specific line(s)",
    "    ce or cell   - Generate clips from specific cell(s) (e.g., P01.11)",
    "    p or participant - Generate all clips for one participant",
    "    k or keyword - Generate only annotated clips/timestamps (e.g., !key)",
    "    sv or severity - Generate clips by severity level",
]

_ALL_MODE_HELP = _SELECTION_MODE_HELP + [
    "    s or screen  - Generate screenshots (.png)",
    "    g or gif     - Generate GIFs (.gif)",
    "    re or reel   - Combine selectors into one highlight reel video",
    "    rl or reellate - Combine existing clips into a highlight reel",
    "    gv or gallery - Generate gallery from interval screenshots/GIFs of a video",
    "    br or browse - Browse spreadsheet rows interactively",
]


def _resolve_unrecognized_input(
    worksheet: Any, user_input: str, *, help_lines: list[str]
) -> list[ClipRecord] | None:
    """Try auto-detection and mixed-selector parsing for input that didn't match a mode alias.

    Attempts, in order: single-type auto-detection (line/range/cell/participant)
    then mixed-selector parsing. Prints appropriate messages.
    Returns clip list on success, None to signal the caller should re-prompt.
    """
    detected_mode, detected_kwargs = spreadsheet.detect_mode_from_input(user_input)
    if detected_mode:
        utils.standard_print(f"  {detected_mode.capitalize()} mode detected.")
        return spreadsheet.generate_list(worksheet, detected_mode, **detected_kwargs)

    parsed = spreadsheet.parse_reel_input(user_input)
    selector_types = [
        ("batch", bool(parsed.get("batch"))),
        ("keyword", bool(parsed.get("keyword"))),
        ("lines", len(parsed["lines"]) > 0),
        ("ranges", len(parsed["ranges"]) > 0),
        ("cells", len(parsed["cells"]) > 0),
        ("participants", len(parsed["participants"]) > 0),
        ("categories", len(parsed["categories"]) > 0),
    ]
    non_empty_types = [name for name, present in selector_types if present]
    has_chronologic = bool(parsed.get("chronologic"))

    if not non_empty_types:
        utils.info_print(f"  Unknown mode or input '{user_input}'. Available modes:")
        for line in help_lines:
            utils.info_print(line)
        return None
    if has_chronologic:
        utils.info_print(
            "  Chronologic selector is only supported for reel/chronologic modes."
        )
        utils.info_print(
            "  Use 're' or 'reel' for a combined reel video, or -T on the command line."
        )
        return None

    selector_summary = ", ".join(non_empty_types)
    utils.standard_print(
        f"  Mixed selectors detected ({selector_summary}). Generating from combined selectors."
    )
    return spreadsheet.generate_list(worksheet, "reel", reel_input=user_input)


def _run_standard_mode(mode: str, worksheet: Any) -> list[ClipRecord] | None:
    """Run a standard interactive mode (batch/line/range/category/cell/participant/keyword).

    Prompts the user for mode-specific input, then generates clips.
    Returns clip list on success, None if the user cancels or context fails.
    """
    utils.print_mode_heading(f"{mode.capitalize()} mode", f"mode.{mode}")
    ctx = spreadsheet.build_sheet_context(worksheet)
    if ctx is None:
        return None
    prompt_fn, gen_fn = _STANDARD_MODES[mode]
    result = prompt_fn(ctx)
    if result is None or result is False:
        return None
    return gen_fn(ctx, result)


def _print_run_summary(message: str) -> None:
    """Print a run summary block with newlines and a Summary header."""
    utils.info_print("")
    utils.print_mode_heading("Summary", "mode.selection")
    utils.info_print(message)
    utils.info_print("")


def _is_excel_worksheet(worksheet: Any) -> bool:
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
                import titlecards

                props = video.probe_video_properties(out_name)
                if props:
                    clip_resolution = f"{props['width']}x{props['height']}"
                ok = titlecards.prepend_titlecard_to_clip(
                    clip, out_name, resolution=clip_resolution
                )
            if ok:
                import titlecards

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
    import transcripts

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
    import transcripts

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
    try:
        cell_a1 = (
            gspread.utils.rowcol_to_a1(cell_row, cell_col)
            if cell_row and cell_col
            else ""
        )
    except Exception:
        cell_a1 = ""
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

    # Lazy-load transcripts manifest once for embedding + transcription
    import transcripts as _tr_mod

    transcripts_manifest = _tr_mod.load_transcripts_manifest()

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
        import titlecards

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
    import transcripts

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
    )
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

    def _concat() -> bool:
        return video.concatenate_clips(clip_paths, output_file, reencode_on_fail=True)

    ok = (
        utils.run_with_spinner("Concatenating clips into final reel...", _concat)
        if utils.use_progress()
        else _concat()
    )
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


# ---- Interactive mode flows ----


def _print_reencoding_warning(printer: Callable[[str], None]) -> None:
    """Print a reminder about trade-offs when ffmpeg runs without re-encoding."""
    printer(
        "* No re-encoding done, expect:\n- inaccurate start and end timings\n- lossy frames until first keyframe\n- bad timecodes at the end\n"
    )


def _print_completion_message(
    outputs_generated: int, output_format: str, is_reel: bool
) -> None:
    """Print a summary of generated outputs tailored to format and reel mode."""
    output_dir = utils.get_effective_output_dir()
    if is_reel:
        _print_run_summary(f"All done, created 1 reel!\nFiles are in {output_dir}")
        return
    noun = {"screen": "screenshots", "gif": "GIFs"}.get(output_format, "videos")
    _print_run_summary(
        f"All done, created {outputs_generated} {noun}!\nFiles are in {output_dir}"
    )


def _prompt_chronologic_participant_selection(worksheet: Any) -> str | None:
    """Prompt user to pick exactly one participant for chronologic reels."""
    ctx = spreadsheet.build_sheet_context(worksheet)
    if ctx is None:
        return None

    available_list = spreadsheet.get_participant_list(
        ctx.header_row, ctx.id_cell, ctx.num_participants
    )
    if not available_list:
        utils.info_print("No participants found in the spreadsheet.")
        return None

    utils.print_mode_heading("Chronologic participant", "mode.chronologic")
    utils.info_print("Chronologic mode requires exactly one participant.")
    utils.info_print("Available participants:")
    for i, pid in enumerate(available_list, 1):
        utils.info_print(f"  {i}. {pid}")

    while True:
        selection = utils.read_user_input(
            "\nEnter one participant number or ID (e.g., 1 or P01):\n>> "
        )
        if not selection:
            utils.info_print("Please enter one participant.")
            continue
        tokens = spreadsheet.parse_participant_selection(selection)
        if len(tokens) != 1:
            utils.info_print("Please provide exactly one participant.")
            continue

        token = tokens[0]
        if token.isdigit():
            idx = int(token)
            if 1 <= idx <= len(available_list):
                return available_list[idx - 1]
            utils.info_print(
                f"Not found: {token}. Available: {', '.join(available_list)}"
            )
            continue

        col_idx = spreadsheet.find_participant_column(
            ctx.header_row, ctx.id_cell, token
        )
        if col_idx is None:
            utils.info_print(
                f"Not found: {token}. Available: {', '.join(available_list)}"
            )
            continue
        if col_idx < len(ctx.header_row):
            return utils.normalize_participant_id(ctx.header_row[col_idx])
        return token


def _run_reel_mode_interactive(
    worksheet: Any,
) -> tuple[list[ClipRecord], bool, str | None]:
    """Run reel mode UI: instructions, input, generate_list, preview, confirm, output filename.
    Returns (clips_list, True, reel_output_file or None) when user confirms; caller may loop on continue.
    """
    utils.print_mode_heading("Reel mode", "mode.reel")

    utils.info_print("Combine selectors into one video. Syntax:")
    utils.info_print("  batch                    - all clips")
    utils.info_print("  keyword                  - annotated clips only")
    utils.info_print(
        "  chronologic              - chronological reel (requires exactly one participant)"
    )
    utils.info_print(
        "  severity                 - order reel by severity (most severe first)"
    )
    utils.info_print(
        "  highlights               - auto-select best clips within time budget"
    )
    utils.info_print("  11, 12, 13-16, 18        - lines and ranges")
    utils.info_print('  "Observations", "Onboarding" - categories (quoted)')
    utils.info_print("  P01.11, P02.15           - cells (participant.row)")
    utils.info_print("  P01, P02                 - participants (all their clips)")
    utils.info_print('  Example: chronologic, P01, 11, 13-16, "Observations"')
    reel_input = utils.read_user_input(
        "\nEnter reel selectors (combine any of the above, comma-separated):\n>> "
    )
    if not reel_input:
        utils.info_print("No input. Skipping reel.")
        return ([], False, None)

    parsed_reel = spreadsheet.parse_reel_input(reel_input)
    if parsed_reel.get("highlights") and (
        parsed_reel["severity"] or parsed_reel["chronologic"]
    ):
        utils.error_print(
            "Cannot combine highlights with severity or chronologic ordering.",
            ["Use highlights on its own for auto-ranked highlight reel."],
        )
        return ([], False, None)
    if parsed_reel["severity"] and parsed_reel["chronologic"]:
        utils.error_print(
            "Cannot combine severity and chronologic ordering.",
            [
                "Use one ordering at a time: either chronologic (chronological) or severity."
            ],
        )
        return ([], False, None)
    if parsed_reel["chronologic"]:
        if len(parsed_reel["participants"]) > 1:
            utils.error_print(
                "Chronologic selector supports only one participant.",
                ["Please provide exactly one participant (e.g., chronologic, P01)."],
            )
            return ([], False, None)
        if len(parsed_reel["participants"]) == 0:
            selected_pid = _prompt_chronologic_participant_selection(worksheet)
            if not selected_pid:
                return ([], False, None)
            reel_input = f"{reel_input}, {selected_pid}"
            parsed_reel["participants"] = [selected_pid]

    clips_list = spreadsheet.generate_list(worksheet, "reel", reel_input=reel_input)
    if not clips_list:
        utils.info_print("No clips matched. Try different selectors.")
        return ([], False, None)
    utils.info_print(
        f"Preview: {len(clips_list)} clip(s) will be included (deduplicated by cell)."
    )
    for i, clip in enumerate(clips_list[: config.REEL_PREVIEW_CLIP_COUNT]):
        desc = (clip.get("desc") or "")[: config.DESCRIPTION_PREVIEW_LENGTH]
        utils.info_print(
            f"  {i + 1}. [{clip.get('category', '')}] {clip.get('participant', '')} row {clip['cell'].row}: {desc}..."
        )
    if len(clips_list) > config.REEL_PREVIEW_CLIP_COUNT:
        utils.info_print(
            f"  ... and {len(clips_list) - config.REEL_PREVIEW_CLIP_COUNT} more"
        )
    yn = utils.read_user_input("\nGenerate reel? [y/n]\n>> ")
    if yn != "y":
        return ([], False, None)

    study_name = clips_list[0].get("study", "").strip() if clips_list else ""
    default_filename = (
        f"{study_name}_reel{config.FILEFORMAT}"
        if study_name
        else f"reel{config.FILEFORMAT}"
    )
    if parsed_reel["chronologic"] and parsed_reel["participants"]:
        chronologic_pid = utils.normalize_participant_id(
            parsed_reel["participants"][0]
        ).strip()
        if study_name and chronologic_pid:
            default_filename = (
                f"{study_name}_{chronologic_pid}_chronologic{config.FILEFORMAT}"
            )
        elif chronologic_pid:
            default_filename = f"{chronologic_pid}_chronologic{config.FILEFORMAT}"
        else:
            default_filename = f"chronologic{config.FILEFORMAT}"
    elif parsed_reel.get("highlights"):
        default_filename = (
            f"{study_name}_highlights{config.FILEFORMAT}"
            if study_name
            else f"highlights{config.FILEFORMAT}"
        )
    elif parsed_reel["severity"]:
        default_filename = (
            f"{study_name}_severity_reel{config.FILEFORMAT}"
            if study_name
            else f"severity_reel{config.FILEFORMAT}"
        )

    output_file = utils.read_user_input(
        f"\nOutput filename (Enter for default {default_filename}):\n>> "
    )
    reel_output_file = None
    if output_file:
        reel_output_file = (
            output_file
            if output_file.endswith(config.FILEFORMAT)
            else output_file + config.FILEFORMAT
        )
    else:
        reel_output_file = files.get_unique_filename(default_filename)
    return (clips_list, True, reel_output_file)


def _parse_clip_selection(selection_input: str, num_clips: int) -> list[int]:
    """Parse user selection input into list of clip indices.

    Supports formats: "A + B + C", "A, B, C", "A B C", or mixed.

    Args:
        selection_input: User's selection string (e.g., "A + B" or "A, C")
        num_clips: Total number of available clips (for validation)

    Returns:
        List of valid 0-based indices, deduplicated and in selection order
    """
    # Normalize separators to spaces
    normalized = selection_input.replace("+", " ").replace(",", " ")
    tokens = normalized.split()

    indices = []
    seen = set()
    for token in tokens:
        idx = utils.letter_to_index(token)
        if idx >= 0 and idx < num_clips and idx not in seen:
            indices.append(idx)
            seen.add(idx)
    return indices


def _run_reellate_mode_interactive() -> tuple[bool, str | None]:
    """Run reel-late mode UI: discover clips, display list, select, concatenate.

    Returns (True, output_file) when reel was generated; (False, None) otherwise.
    """
    clips = files.discover_clips()
    output_dir = utils.get_effective_output_dir()
    utils.print_mode_heading("Reel-late mode", "mode.reellate")

    if not clips:
        utils.info_print("No clips found in the output directory.")
        utils.info_print("  Source videos (like study_P01.mp4) are excluded.")
        utils.info_print(
            "  Generate some clips first, then use this mode to combine them."
        )
        return (False, None)

    utils.info_print("Combine existing clips into a highlight reel.")
    utils.info_print(f"Found {len(clips)} clip(s) in {output_dir}:")

    # Display indexed list
    for i, clip in enumerate(clips):
        letter = utils.index_to_letter(i)
        utils.info_print(f'  {letter}. "{clip}"')

    utils.info_print("Select clips to include (order preserved). Syntax:")
    utils.info_print("  A + B + C    - combine clips A, B, and C")
    utils.info_print("  A, B, C      - same as above")
    utils.info_print("  A B C        - same as above")

    selection_input = utils.read_user_input("\nEnter clip selection:\n>> ")
    if not selection_input:
        utils.info_print("No selection. Skipping reel.")
        return (False, None)

    indices = _parse_clip_selection(selection_input, len(clips))
    if not indices:
        utils.warning_print(
            "No valid clips selected.",
            ["Use letters from the list above (e.g., A + B + C)"],
        )
        return (False, None)

    selected_clips = [clips[i] for i in indices]

    # Preview selection
    utils.info_print(f"Selected {len(selected_clips)} clip(s):")
    for i, clip in enumerate(selected_clips):
        utils.info_print(f'  {i + 1}. "{clip}"')

    yn = utils.read_user_input("\nGenerate reel from these clips? [y/n]\n>> ").lower()
    if yn != "y":
        utils.info_print("Cancelled.")
        return (False, None)

    output_file = utils.read_user_input(
        '\nOutput filename (Enter for default "reel.mp4"):\n>> '
    )
    if not output_file:
        output_file = files.get_unique_filename(f"reel{config.FILEFORMAT}")
    elif not output_file.endswith(config.FILEFORMAT):
        output_file = output_file + config.FILEFORMAT

    resolved_clips = [str(utils.resolve_output_path(name)) for name in selected_clips]

    def _concat_reellate() -> bool:
        return video.concatenate_clips(
            resolved_clips, output_file, reencode_on_fail=True
        )

    ok = (
        utils.run_with_spinner(
            f"Concatenating {len(selected_clips)} clips into {output_file}...",
            _concat_reellate,
        )
        if utils.use_progress()
        else _concat_reellate()
    )

    if ok:
        return (True, output_file)
    return (False, None)


def _run_format_mode_interactive(worksheet: Any, output_format: str) -> None:
    """Run interactive flow for screen/gif output formats.

    Prompts the user for a timestamp-selection mode, generates a clip list, then renders
    screenshots/GIFs using the regular clip processing pipeline.
    """
    format_display_name = "Screenshot" if output_format == "screen" else "GIF"
    output_label = "screenshots" if output_format == "screen" else "GIFs"
    utils.print_mode_heading(f"{format_display_name} mode", "mode.format")

    utils.info_print("Choose how to select timestamps.")
    while True:
        selection = utils.read_user_input(
            "\nSelect source rows for this output:\n"
            "  Modes: (b)atch, (r)ange, (c)ategory, (l)ine, (ce)ll, (p)articipant, (k)eyword, (sv) severity\n"
            '  Or enter mixed selectors directly: e.g. 5, P01.11, 13-16, "Observations"\n>> '
        )
        if not selection:
            utils.info_print(
                "  Please enter a mode or direct input (e.g. P01.11, 5, 7, 13-16, P01)."
            )
            continue

        mode = FORMAT_MODE_ALIASES.get(selection.lower())
        if mode is None and len(selection.split()) == 1 and selection.strip():
            full_names = [k for k in FORMAT_MODE_ALIASES if len(k) > 2]
            suggestion = utils.suggest_close_match(selection.strip(), full_names)
            if suggestion is not None:
                mode = FORMAT_MODE_ALIASES[suggestion]
        if mode:
            clips_list = spreadsheet.generate_list(worksheet, mode)
            break

        clips_list = _resolve_unrecognized_input(
            worksheet, selection, help_lines=_SELECTION_MODE_HELP
        )
        if clips_list is not None:
            break

    outputs_generated, artifacts = process_clips(
        clips_list, output_format=output_format, include_severity=(mode == "severity")
    )
    if artifacts:
        viewer.INTERACTIVE_ARTIFACTS.extend(artifacts)
        if config.MANIFEST_ENABLED:
            viewer.save_manifest(
                viewer.INTERACTIVE_ARTIFACTS,
                new_reels=viewer.INTERACTIVE_REELS,
                study=artifacts[0].get("study", ""),
                worksheet_title=getattr(worksheet, "title", ""),
                is_excel=_is_excel_worksheet(worksheet),
                mode="interactive",
            )
    _print_run_summary(
        f"All done, created {outputs_generated} {output_label}!\nFiles are in {utils.get_effective_output_dir()}"
    )


def _run_viewer_mode(worksheet: Any) -> None:
    """Generate the timeline viewer from session artifacts, falling back to manifest."""
    artifacts = list(viewer.INTERACTIVE_ARTIFACTS)
    mode_label = "interactive"

    if not artifacts:
        manifest_artifacts = viewer.load_manifest_artifacts()
        if not manifest_artifacts:
            utils.info_print(
                "No artifacts in this session and no manifest file found.\n"
                "Generate clips first, or run a prior session with --manifest to save one."
            )
            return
        count = len(manifest_artifacts)
        yn = utils.read_user_input(
            f"No artifacts in this session, but found {count} artifact(s) in manifest.\n"
            "Generate viewer from manifest? [y/n]\n>> "
        )
        if yn.strip().lower() != "y":
            return
        artifacts = manifest_artifacts
        mode_label = "manifest"

    study = artifacts[0].get("study", "")
    participant = artifacts[0].get("participant", "")
    ss_events = viewer.load_screenspace_events_for_viewer()
    data = viewer.finalize_timeline_data(
        artifacts,
        study=study,
        participant=participant,
        worksheet_title=""
        if mode_label == "manifest"
        else getattr(worksheet, "title", ""),
        is_excel=False if mode_label == "manifest" else _is_excel_worksheet(worksheet),
        mode=mode_label,
        output_format="clip",
        screenspace_events=ss_events or None,
    )
    viewer_path = viewer.generate_timeline_viewer(data)
    if viewer_path:
        utils.info_print(f"Timeline viewer created: {viewer_path}")


def _run_regenerate_mode() -> None:
    """Regenerate all media artifacts and reels from saved manifest."""
    existing_artifacts, existing_reels = viewer._load_manifest_both()
    if not existing_artifacts and not existing_reels:
        utils.info_print(
            "No manifest file found.\nGenerate clips first with --manifest to save one."
        )
        return
    media_count = sum(1 for a in existing_artifacts if a.get("type") != "transcript")
    reel_count = len(existing_reels)
    total = media_count + reel_count
    if total == 0:
        utils.info_print(
            "Manifest contains only transcript artifacts; nothing to regenerate."
        )
        return
    yn = utils.read_user_input(
        f"Found {media_count} media artifact(s) and {reel_count} reel(s) in manifest.\n"
        "Regenerate all? [y/n]\n>> "
    )
    if yn.strip().lower() != "y":
        return
    regenerated = regenerate_from_manifest(existing_artifacts, reels=existing_reels)
    utils.info_print(f"Regenerated {regenerated} of {total} item(s).")


def _run_gallery_mode_interactive() -> None:
    """Interactive gallery mode: select a video, generate interval captures, build gallery viewer."""
    utils.print_mode_heading("Gallery mode", "mode.gallery")

    input_dir = utils.get_effective_input_dir()
    videos = sorted(p for p in input_dir.glob(f"*{config.FILEFORMAT}") if p.is_file())
    if not videos:
        utils.error_print(
            f"No {config.FILEFORMAT} files found in {input_dir}.",
            ["Place a video file in the input directory and try again."],
        )
        return

    utils.info_print("Available videos:")
    for i, v in enumerate(videos, 1):
        size_mb = v.stat().st_size / 1_000_000
        utils.info_print(f"  {i}. {v.name}  ({size_mb:.0f} MB)")

    selection = utils.read_user_input("Select a video (number or filename):\n>> ")
    video_path: Path | None = None
    try:
        idx = int(selection) - 1
        if 0 <= idx < len(videos):
            video_path = videos[idx]
    except ValueError:
        for v in videos:
            if v.name == selection.strip() or v.stem == selection.strip():
                video_path = v
                break

    if video_path is None:
        utils.error_print(f"Could not find video matching '{selection}'.")
        return

    fmt_input = (
        utils.read_user_input("Generate (s)creenshots or (g)ifs? [s]\n>> ")
        .strip()
        .lower()
    )
    output_format = "gif" if fmt_input in ("g", "gif") else "screen"

    interval_input = utils.read_user_input(
        f"Capture interval in seconds [{config.GALLERY_INTERVAL_SECONDS}]:\n>> "
    ).strip()
    try:
        interval = (
            int(interval_input) if interval_input else config.GALLERY_INTERVAL_SECONDS
        )
    except ValueError:
        interval = config.GALLERY_INTERVAL_SECONDS
    if interval <= 0:
        interval = config.GALLERY_INTERVAL_SECONDS

    gif_duration = config.GALLERY_GIF_DURATION_SECONDS
    if output_format == "gif":
        dur_input = utils.read_user_input(
            f"GIF duration in seconds [{config.GALLERY_GIF_DURATION_SECONDS}]:\n>> "
        ).strip()
        try:
            gif_duration = (
                int(dur_input) if dur_input else config.GALLERY_GIF_DURATION_SECONDS
            )
        except ValueError:
            gif_duration = config.GALLERY_GIF_DURATION_SECONDS
        if gif_duration <= 0:
            gif_duration = config.GALLERY_GIF_DURATION_SECONDS

    artifacts = video.generate_interval_captures(
        str(video_path),
        interval_seconds=interval,
        output_format=output_format,
        gif_duration_seconds=gif_duration,
    )
    if not artifacts:
        return

    duration = video.get_file_duration(str(video_path)) or 0
    data = viewer.finalize_gallery_data(
        artifacts,
        source_video=video_path.name,
        video_duration=duration,
        output_format=output_format,
        interval=interval,
    )
    gallery_path = viewer.generate_gallery_viewer(data)
    if gallery_path:
        utils.info_print(f"Gallery viewer created: {gallery_path}")


def _dispatch_interactive_mode(
    mode: str | None, worksheet: Any, raw_input: str
) -> tuple[list[ClipRecord], bool, str | None] | None:
    """Dispatch a resolved mode or raw input. Returns result tuple, or None to re-prompt."""
    # Special modes with their own interactive flows
    if mode == "browse":

        def _browse_process_fn(clips_list, output_format):
            outputs_generated, artifacts = process_clips(
                clips_list, output_format=output_format
            )
            if artifacts:
                viewer.INTERACTIVE_ARTIFACTS.extend(artifacts)
                if config.MANIFEST_ENABLED:
                    viewer.save_manifest(
                        viewer.INTERACTIVE_ARTIFACTS,
                        new_reels=viewer.INTERACTIVE_REELS,
                        study=artifacts[0].get("study", ""),
                        worksheet_title=getattr(worksheet, "title", ""),
                        is_excel=_is_excel_worksheet(worksheet),
                        mode="interactive",
                    )
            if not config.REENCODING:
                _print_reencoding_warning(utils.info_print)
            return (outputs_generated, artifacts)

        interactive.browse_spreadsheet(worksheet, process_fn=_browse_process_fn)
        return ([], False, None)
    if mode == "viewer":
        _run_viewer_mode(worksheet)
        return None
    if mode == "regenerate":
        _run_regenerate_mode()
        return None
    if mode == "gallery":
        _run_gallery_mode_interactive()
        return None
    if mode == "insights":
        import server

        server.start_combined_server(worksheet=worksheet, default_page="insights")
        return None
    if mode == "studio":
        import server

        server.start_combined_server(worksheet=worksheet, default_page="studio")
        return None
    if mode == "screenspace":
        import server

        server.start_combined_server(worksheet=worksheet, default_page="screenspace")
        return None
    if mode == "transcripts":
        import server

        server.start_combined_server(worksheet=worksheet, default_page="transcripts")
        return None
    if mode == "timeline-viewer":
        clips_list = spreadsheet.generate_list(worksheet, "batch", skip_prompts=True)
        outputs_generated, artifacts = process_clips(clips_list, output_format="clip")
        if not config.REENCODING:
            _print_reencoding_warning(utils.info_print)
        _print_completion_message(outputs_generated, "clip", is_reel=False)
        if artifacts:
            viewer.INTERACTIVE_ARTIFACTS.extend(artifacts)
            if config.MANIFEST_ENABLED:
                viewer.save_manifest(
                    viewer.INTERACTIVE_ARTIFACTS,
                    new_reels=viewer.INTERACTIVE_REELS,
                    study=artifacts[0].get("study", ""),
                    worksheet_title=getattr(worksheet, "title", ""),
                    is_excel=_is_excel_worksheet(worksheet),
                    mode="timeline-viewer",
                )
            study = artifacts[0].get("study", "")
            ss_events = viewer.load_screenspace_events_for_viewer()
            data = viewer.finalize_timeline_data(
                artifacts,
                study=study,
                worksheet_title=getattr(worksheet, "title", ""),
                is_excel=_is_excel_worksheet(worksheet),
                mode="timeline-viewer",
                output_format="clip",
                screenspace_events=ss_events or None,
            )
            viewer_path = viewer.generate_timeline_viewer(
                data,
                template_name="timeline-viewer.html",
                output_basename="timeline_viewer.html",
            )
            if viewer_path:
                utils.info_print(f"Participant timeline viewer created: {viewer_path}")
        else:
            utils.warning_print(
                "No artifacts were generated; skipping timeline viewer."
            )
        return None
    if mode == "settings":
        utils.set_program_settings()
        return None
    if mode == "reel":
        clips, confirmed, reel_file = _run_reel_mode_interactive(worksheet)
        return (clips, True, reel_file) if confirmed else None
    if mode == "reellate":
        success, output_file = _run_reellate_mode_interactive()
        if success:
            _print_run_summary(
                f"Reel created: {output_file}\nFiles are in {utils.get_effective_output_dir()}"
            )
            return ([], False, None)
        return None
    if mode in ("screen", "gif"):
        _run_format_mode_interactive(worksheet, mode)
        return ([], False, None)

    # Standard modes with interactive prompts
    if mode in _STANDARD_MODES:
        clips = _run_standard_mode(mode, worksheet)
        return (clips or [], False, None)
    if mode:
        return (spreadsheet.generate_list(worksheet, mode), False, None)

    # No alias match -- try auto-detection and mixed selectors
    clips = _resolve_unrecognized_input(worksheet, raw_input, help_lines=_ALL_MODE_HELP)
    if clips is None:
        return None
    return (clips, False, None)


def run_interactive_mode(worksheet: Any) -> None:
    """Execute interactive mode - main processing loop."""
    viewer.INTERACTIVE_ARTIFACTS.clear()
    viewer.INTERACTIVE_REELS.clear()

    while True:
        try:
            # Mode selection (inlined from select_mode_and_generate)
            utils.print_mode_heading("Mode selection", "mode.selection")
            input_mode = utils.read_user_input(
                "\nEnter mode or input directly:\n"
                "  Tools: (s)creen, (g)if, (re)el, (rl) reel-late, (rg) regenerate, (se)ttings \n"
                "  Front: (st) studio, (in) insights, (ss) screenspace, (tr)anscripts, (br)owse \n"
                "  Packs: (v)iewer, (tv) timeline-viewer, (gv) gallery \n"
                "  Modes: (b)atch, (r)ange, (c)ategory, (l)ine, (ce)ll, (p)articipant, (k)eyword, (sv) severity \n"
                '  Or enter mixed selectors directly: e.g. 5, P01.11, 13-16, "Observations"\n>> '
            )
            if not input_mode:
                utils.info_print(
                    "  Please enter a mode or direct input (e.g. P01.11, 5, 7, 13-16, P01)."
                )
                continue

            resolved_mode = MODE_ALIASES.get(input_mode.strip().lower())
            if (
                resolved_mode is None
                and len(input_mode.split()) == 1
                and input_mode.strip()
            ):
                full_mode_names = [k for k in MODE_ALIASES if len(k) > 2]
                suggestion = utils.suggest_close_match(
                    input_mode.strip(), full_mode_names
                )
                if suggestion is not None:
                    resolved_mode = MODE_ALIASES[suggestion]
            result = _dispatch_interactive_mode(resolved_mode, worksheet, input_mode)
            if result is None:
                continue
            clips_list, is_reel, reel_output_file = result

            if not clips_list and not is_reel:
                yn = utils.read_user_input(
                    "Continue working (y) or quit the program (n)? [y/n]\n>> "
                )
                if yn == "n":
                    break
                continue
            if is_reel:
                outputs_generated, reel_records = process_reel(
                    clips_list,
                    output_file=reel_output_file,
                )
                if reel_records:
                    viewer.INTERACTIVE_REELS.extend(reel_records)
                    if config.MANIFEST_ENABLED:
                        viewer.save_manifest(
                            viewer.INTERACTIVE_ARTIFACTS,
                            new_reels=viewer.INTERACTIVE_REELS,
                            study=reel_records[0].get("study", ""),
                            worksheet_title=getattr(worksheet, "title", ""),
                            is_excel=_is_excel_worksheet(worksheet),
                            mode="interactive",
                        )
            else:
                outputs_generated, artifacts = process_clips(
                    clips_list, include_severity=(resolved_mode == "severity")
                )
                if artifacts:
                    viewer.INTERACTIVE_ARTIFACTS.extend(artifacts)
                    if config.MANIFEST_ENABLED:
                        viewer.save_manifest(
                            viewer.INTERACTIVE_ARTIFACTS,
                            new_reels=viewer.INTERACTIVE_REELS,
                            study=artifacts[0].get("study", ""),
                            worksheet_title=getattr(worksheet, "title", ""),
                            is_excel=_is_excel_worksheet(worksheet),
                            mode="interactive",
                        )

            if not config.REENCODING:
                _print_reencoding_warning(utils.info_print)
            _print_completion_message(
                outputs_generated, output_format="clip", is_reel=is_reel
            )

            yn = utils.read_user_input(
                "Continue working (y) or quit the program (n)? [y/n]\n>> "
            )
            if yn == "n":
                break
        except gspread.exceptions.GSpreadException as e:
            utils.error_print(f"Google Sheets API error: {e}")
            utils.debug_print(f"ERROR Message '{e}', Attempting reconnect")
        except utils.TopToSpreadsheet:
            # Escalate to main to trigger spreadsheet reselection.
            raise
        except utils.BackToModeSelection:
            # Return to main mode selection prompt.
            continue
        except utils.QuitProgram:
            break


if __name__ == "__main__":
    from cli import main

    try:
        main()
    except KeyboardInterrupt:
        utils.info_print("Interrupted by user")
        sys.exit(0)
