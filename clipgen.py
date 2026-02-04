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
import io
import os
import sys
from typing import Any, List, Optional, Tuple

import gspread
from icecream import ic

import config
import excel_io
import files
import google_api
import spreadsheet
import utils
import video


def _open_worksheet(gc: Any, open_fn, error_context: str):
    """Try to open a worksheet via a callable; catch gspread errors and print a consistent message.
    
    Args:
        gc: Google client connection
        open_fn: Callable that takes no args and returns a gspread Spreadsheet (e.g. lambda: gc.open_by_url(url))
        error_context: Short description for error message (e.g. "by URL", "at index 3")
        
    Returns:
        Worksheet object or None if error
    """
    try:
        return google_api.get_worksheet(open_fn())
    except (gspread.SpreadsheetNotFound, gspread.exceptions.APIError, gspread.exceptions.GSpreadException) as e:
        utils.error_print(f"Could not open spreadsheet {error_context}: {e}")
        return None


def open_spreadsheet_by_url(gc: Any, url: str) -> Optional[Any]:
    """Open a spreadsheet by URL.
    
    Args:
        gc: Google client connection
        url: Spreadsheet URL
        
    Returns:
        Worksheet object or None if error
    """
    return _open_worksheet(gc, lambda: gc.open_by_url(url), "by URL")


def open_spreadsheet_by_index(gc: Any, doc_list: List[str], index: int) -> Optional[Any]:
    """Open a spreadsheet by index number.
    
    Args:
        gc: Google client connection
        doc_list: List of spreadsheet names
        index: Index number (1-based)
        
    Returns:
        Worksheet object or None if error
    """
    if index < 1 or index > len(doc_list):
        utils.error_print(f"Invalid index {index}. Must be between 1 and {len(doc_list)}")
        return None
    chosen_index = index - 1
    doc_name = doc_list[chosen_index].strip()
    utils.verbose_print(f'Opening document: {doc_name}')
    return _open_worksheet(gc, lambda: gc.open(doc_name), f"at index {index}")


def open_spreadsheet_by_name(gc: Any, doc_list: List[str], name: str) -> Optional[Any]:
    """Open a spreadsheet by name.
    
    Args:
        gc: Google client connection
        doc_list: List of spreadsheet names
        name: Spreadsheet name to search for
        
    Returns:
        Worksheet object or None if not found
    """
    chosen_index = google_api.find_spreadsheet_by_name(name, doc_list)
    if chosen_index >= 0:
        matched_name = doc_list[chosen_index].strip()
        utils.verbose_print(f'Opening document: {matched_name}')
        return _open_worksheet(gc, lambda: gc.open(matched_name), f"'{name}'")
    return None


def handle_list_all_command(doc_list: List[str]) -> None:
    """Handle 'all' command - list all available documents."""
    utils.info_print('\nAvailable documents:')
    for i, doc in enumerate(doc_list):
        utils.info_print(f'{i+1}. {doc.strip()}')

def handle_list_new_command(doc_list: List[str]) -> None:
    """Handle 'new' command - list newest documents."""
    utils.info_print('\nNewest documents: (modified or opened most recently)')
    for i in range(min(config.NUM_NEWEST_DOCS_TO_SHOW, len(doc_list))):
        utils.info_print(f'{i+1}. {doc_list[i].strip()}')

def handle_error_message(consecutive_open_failures: int, e: Exception) -> None:
    """Handle error messages with progressive detail based on failure count."""
    if consecutive_open_failures == 1:
        utils.error_print(f"Could not access spreadsheet: {e}", 
            [f"Please try again. Type '{config.COMMAND_LIST_ALL}' to see available documents."])
    elif consecutive_open_failures == 2:
        utils.error_print("Spreadsheet not found or not accessible.",
            ["Common causes:",
             "  - The spreadsheet name is misspelled",
             "  - The spreadsheet hasn't been shared with your service account",
             "    (Share it with the email in credentials.json 'client_email' field)",
             "  - The spreadsheet doesn't contain any worksheets",
             "",
             f"  Type '{config.COMMAND_LIST_ALL}' to see accessible documents, or '{config.COMMAND_LIST_NEW}' for recent ones."])
    else:
        utils.error_print(str(e), [f"Tip: Use the document index number (1, 2, 3...) from the '{config.COMMAND_LIST_ALL}' list."])

def _handle_spreadsheet_command(gc: Any, doc_list: List[str], input_name: str) -> Optional[Any]:
    """Handle one spreadsheet selection command. Returns worksheet when one was opened, None to show prompt again."""
    # Handle 'excel' for local .xlsx
    if input_name.strip().lower() == config.COMMAND_EXCEL:
        return excel_io.select_excel_file()
    # Handle URL
    if input_name.startswith(config.COMMAND_HTTP_PREFIX):
        return open_spreadsheet_by_url(gc, input_name)
    # Handle 'all' command
    if input_name.startswith(config.COMMAND_LIST_ALL):
        handle_list_all_command(doc_list)
        return None
    # Handle 'new' command
    if input_name.startswith(config.COMMAND_LIST_NEW):
        handle_list_new_command(doc_list)
        return None
    # Handle 'last' command
    if input_name.startswith(config.COMMAND_OPEN_LAST):
        latest = google_api.get_all_spreadsheets(gc).split(',')[0]
        return open_spreadsheet_by_name(gc, doc_list, latest)
    # Handle numeric index
    if input_name[0].isdigit():
        return open_spreadsheet_by_index(gc, doc_list, int(input_name))
    # Handle 'settings' command
    if input_name.startswith(config.COMMAND_SETTINGS):
        utils.set_program_settings()
        return None
    # Handle name search
    return open_spreadsheet_by_name(gc, doc_list, input_name)


def select_spreadsheet(gc: Any, doc_list: List[str]) -> Any:
    """Interactive spreadsheet selection. Returns the selected worksheet."""
    consecutive_open_failures = 0
    while True:
        input_name = input(f"\nPlease enter the index, name, URL, or '{config.COMMAND_EXCEL}' for local file ('{config.COMMAND_LIST_ALL}' for list, '{config.COMMAND_LIST_NEW}' for list of newest, '{config.COMMAND_OPEN_LAST}' to immediately open latest, '{config.COMMAND_SETTINGS}' to change settings):\n>> ")
        try:
            worksheet = _handle_spreadsheet_command(gc, doc_list, input_name)
            if worksheet is not None:
                return worksheet
        except (gspread.SpreadsheetNotFound, gspread.exceptions.APIError, gspread.exceptions.GSpreadException) as e:
            consecutive_open_failures += 1
            handle_error_message(consecutive_open_failures, e)
        except Exception as e:
            utils.error_print(f"Could not open document: {e}")

def _run_reel_mode_interactive(worksheet: Any) -> Tuple[List[Any], bool, Optional[str]]:
    """Run reel mode UI: instructions, input, generate_list, preview, confirm, output filename.
    Returns (clips_list, True, reel_output_file or None) when user confirms; caller may loop on continue.
    """
    utils.info_print('\nReel mode: combine selectors into one video. Syntax:')
    utils.info_print('  batch                    - all clips')
    utils.info_print('  11, 12, 13-16, 18        - lines and ranges')
    utils.info_print('  "Observations", "Onboarding" - categories (quoted)')
    utils.info_print('  P01.11, P02.15           - cells (participant.row)')
    utils.info_print('  P01, P02                 - participants (all their clips)')
    utils.info_print('  Example: 11, 13-16, P01, "Observations"')
    reel_input = input('\nEnter reel selectors (combine any of the above, comma-separated):\n>> ').strip()
    if not reel_input:
        utils.info_print('No input. Skipping reel.')
        return ([], False, None)
    clips_list = spreadsheet.generate_list(worksheet, 'reel', reel_input=reel_input)
    if not clips_list:
        utils.info_print('No clips matched. Try different selectors.')
        return ([], False, None)
    utils.info_print(f'\nPreview: {len(clips_list)} clip(s) will be included (deduplicated by cell).')
    for i, clip in enumerate(clips_list[:config.REEL_PREVIEW_CLIP_COUNT]):
        desc = (clip.get('desc') or '')[:config.DESCRIPTION_PREVIEW_LENGTH]
        utils.info_print(f"  {i+1}. [{clip.get('category', '')}] {clip.get('participant', '')} row {clip['cell'].row}: {desc}...")
    if len(clips_list) > config.REEL_PREVIEW_CLIP_COUNT:
        utils.info_print(f"  ... and {len(clips_list) - config.REEL_PREVIEW_CLIP_COUNT} more")
    yn = input('\nGenerate reel? y/n\n>> ')
    if yn != 'y':
        return ([], False, None)
    output_file = input('\nOutput filename (Enter for default {study}_reel.mp4):\n>> ').strip()
    reel_output_file = None
    if output_file:
        reel_output_file = output_file if output_file.endswith(config.FILEFORMAT) else output_file + config.FILEFORMAT
    return (clips_list, True, reel_output_file)


def select_mode_and_generate(worksheet: Any) -> Tuple[List[Any], bool, Optional[str]]:
    """Interactive mode selection. Returns (clips list, is_reel_mode, reel_output_file or None)."""
    mode_map = {
        'b': 'batch', 'batch': 'batch',
        'l': 'line', 'line': 'line',
        'r': 'range', 'range': 'range',
        'c': 'category', 'cat': 'category', 'category': 'category',
        'ce': 'cell', 'cell': 'cell',
        'p': 'participant', 'participant': 'participant',
        're': 'reel', 'reel': 'reel',
        'br': 'browse', 'browse': 'browse',
        'test': 'test'
    }

    while True:
        input_mode = input(
            '\nEnter mode or input directly:\n'
            '  Modes: (b)atch, (r)ange, (c)ategory, (l)ine, (ce)ll, (p)articipant, (re)el, (br)owse\n'
            '  Or enter directly: line numbers (5, 7), ranges (13-16), cells (P01.11), participants (P01)\n>> '
        ).strip()
        if not input_mode:
            utils.info_print("  Please enter a mode or direct input (e.g. P01.11, 5, 7, 13-16, P01).")
            continue
        input_lower = input_mode.strip().lower()
        try:
            # Only treat as explicit mode when input exactly matches a mode shortcut or name
            mode = mode_map.get(input_lower)
            if mode == 'browse':
                spreadsheet.browse_spreadsheet(worksheet)
                return ([], False, None)
            if mode == 'reel':
                result = _run_reel_mode_interactive(worksheet)
                if result[1]:
                    return result
                continue
            if mode:
                return (spreadsheet.generate_list(worksheet, mode), False, None)

            # Try implicit mode detection from input syntax
            detected_mode, detected_kwargs = spreadsheet.detect_mode_from_input(input_mode)
            if detected_mode:
                utils.verbose_print(f"  {detected_mode.capitalize()} mode detected.")
                return (spreadsheet.generate_list(worksheet, detected_mode, **detected_kwargs), False, None)

            # No mode and no detection: check for mixed input to give a helpful message
            parsed = spreadsheet.parse_reel_input(input_mode)
            types_present = [
                ('lines', parsed['lines']),
                ('ranges', parsed['ranges']),
                ('cells', parsed['cells']),
                ('participants', parsed['participants']),
            ]
            non_empty = [name for name, vals in types_present if vals]
            if len(non_empty) > 1:
                utils.info_print("  Mixed input types detected (" + ", ".join(non_empty) + "). Use 're' or 'reel' mode to combine selectors.")
            else:
                utils.info_print(f"  Unknown mode or input '{input_mode}'. Available modes:")
                utils.info_print("    b or batch   - Generate all clips in the spreadsheet")
                utils.info_print("    r or range   - Generate clips from a range of rows")
                utils.info_print("    c or category - Generate clips by category")
                utils.info_print("    l or line    - Generate clips from specific line(s)")
                utils.info_print("    ce or cell   - Generate clips from specific cell(s) (e.g., P01.11)")
                utils.info_print("    p or participant - Generate all clips for one participant")
                utils.info_print("    re or reel   - Combine selectors into one highlight reel video")
                utils.info_print("    br or browse - Browse spreadsheet rows interactively")
        except gspread.exceptions.GSpreadException as e:
            utils.error_print(f"Google Sheets API error: {e}")
            utils.debug_print(f"ERROR Message '{e}', Attempting reconnect")

def _process_single_clip_segments(
    clip: Any,
    base_video: str,
    missing_videos: set,
    *,
    filename_prefix: str = "",
    collect_paths: bool = False,
) -> Tuple[int, List[str]]:
    """Process one clip's segments: run ffmpeg for each (start, end), optionally collect output paths.

    Caller must have already called prepare_clip(clip). Does not add to missing_videos; caller handles that.

    Args:
        clip: Prepared clip dict with 'times', 'category', 'study', 'participant', 'desc'
        base_video: Path to source video file
        missing_videos: Set of already-reported missing paths (read-only here)
        filename_prefix: Prefix for output filename (e.g. '_reel_part_' for reel)
        collect_paths: If True, return list of output paths; otherwise return empty list

    Returns:
        (number of segments successfully generated, list of output paths if collect_paths else [])
    """
    generated = 0
    paths: List[str] = []
    template = (
        f"{filename_prefix}[{clip['category']}] {clip['study']} {clip['participant']} {clip['desc']}{config.FILEFORMAT}"
    )
    for start_time, end_time in clip['times']:
        try:
            out_name = files.get_unique_filename(template)
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
            return (generated, paths)
        ok = video.run_ffmpeg(
            input_file=base_video,
            output_file=out_name,
            start_pos=start_time,
            end_pos=end_time,
            reencode=config.REENCODING,
        )
        if ok:
            generated += 1
            if collect_paths:
                paths.append(out_name)
    return (generated, paths)


def process_clips(clips_list: List[Any]) -> int:
    """Process and generate video clips from the clips list. Returns count of videos generated."""
    if config.DEBUGGING:
        ic(len(clips_list))
    if not clips_list:
        utils.warning_print("No clips to process. No timestamps were found or selected.")
        return 0

    utils.verbose_print('\n* ffmpeg is set to never prompt for input and will always overwrite.\n  Only warns if close to crashing.\n')
    videos_generated = 0
    videos_skipped = 0
    missing_videos: set = set()
    total_clips = len(clips_list)
    progress = utils.create_progress_bar()

    def process_single_clip(clip, progress_task=None):
        """Process a single clip, updating progress if available. Returns (generated, skipped)."""
        nonlocal videos_generated, videos_skipped, missing_videos

        if config.DEBUGGING:
            ic(clip)
        clip = files.prepare_clip(clip)
        if not clip['times']:
            return (0, 1)

        base_video = f"{clip['study']}_{clip['participant']}{config.FILEFORMAT}"
        if not os.path.isfile(base_video):
            if base_video not in missing_videos:
                missing_videos.add(base_video)
                utils.error_print(
                    f"Source video file not found: '{base_video}'",
                    [
                        f"Expected location: {os.path.join(os.getcwd(), base_video)}",
                        f"Clips for participant '{clip['participant']}' in study '{clip['study']}' will be skipped.",
                    ],
                )
            return (0, len(clip['times']))

        count, _ = _process_single_clip_segments(clip, base_video, missing_videos, collect_paths=False)
        skipped = len(clip['times']) - count if count < len(clip['times']) else 0
        return (count, skipped)

    if progress:
        with progress:
            task = progress.add_task("Processing clips", total=total_clips)
            for clip in clips_list:
                # Update description to show current clip
                desc_preview = (clip.get('desc') or '')[:30]
                participant = clip.get('participant', '')
                progress.update(task, description=f"[{participant}] {desc_preview}...")

                generated, skipped = process_single_clip(clip, task)
                videos_generated += generated
                videos_skipped += skipped
                progress.update(task, advance=1)
    else:
        # Fallback: no progress bar
        for i, clip in enumerate(clips_list):
            if config.VERBOSE and total_clips > 1:
                utils.verbose_print(f"Processing clip {i+1} of {total_clips}...")

            generated, skipped = process_single_clip(clip)
            videos_generated += generated
            videos_skipped += skipped

    if videos_skipped > 0:
        utils.verbose_print(f"\n* Summary: {videos_generated} video(s) generated, {videos_skipped} skipped due to errors.")
    if missing_videos:
        utils.verbose_print(f"* Missing source video files: {len(missing_videos)}")
    return videos_generated


def process_reel(clips_list: List[Any], output_file: Optional[str] = None) -> int:
    """Process clips for reel mode: generate individual clips, concatenate into one video, clean up.

    Returns 1 if the reel was generated successfully, 0 otherwise.
    """
    if not clips_list:
        utils.warning_print("No clips to process for reel. No timestamps were found or selected.")
        return 0

    utils.verbose_print('\n* Reel mode: generating individual clips, then concatenating into one file.\n')
    clip_paths: List[str] = []
    missing_videos: set = set()
    study_name: Optional[str] = None
    total_clips = len(clips_list)
    progress = utils.create_progress_bar()

    def process_reel_clip(clip):
        """Process a single clip for reel mode. Returns list of generated clip paths."""
        nonlocal study_name, missing_videos

        clip = files.prepare_clip(clip)
        if not clip['times']:
            return []
        if study_name is None:
            study_name = clip['study']
        base_video = f"{clip['study']}_{clip['participant']}{config.FILEFORMAT}"
        if not os.path.isfile(base_video):
            if base_video not in missing_videos:
                missing_videos.add(base_video)
                utils.error_print(
                    f"Source video file not found: '{base_video}'",
                    [
                        f"Expected location: {os.path.join(os.getcwd(), base_video)}",
                        f"Clips for participant '{clip['participant']}' will be skipped.",
                    ],
                )
            return []
        _, paths = _process_single_clip_segments(
            clip, base_video, missing_videos, filename_prefix="_reel_part_", collect_paths=True
        )
        return paths

    if progress:
        with progress:
            task = progress.add_task("Generating reel clips", total=total_clips)
            for clip in clips_list:
                desc_preview = (clip.get('desc') or '')[:30]
                participant = clip.get('participant', '')
                progress.update(task, description=f"[{participant}] {desc_preview}...")

                paths = process_reel_clip(clip)
                clip_paths.extend(paths)
                progress.update(task, advance=1)
    else:
        # Fallback: no progress bar
        for clip in clips_list:
            paths = process_reel_clip(clip)
            clip_paths.extend(paths)

    if missing_videos:
        utils.verbose_print(f"* Missing source video files: {list(missing_videos)}")
    if not clip_paths:
        utils.warning_print("No clips were generated for the reel.")
        return 0

    if output_file is None and study_name:
        output_file = files.get_unique_filename(f"{study_name}_reel{config.FILEFORMAT}")
    elif output_file is None:
        output_file = files.get_unique_filename(f"reel{config.FILEFORMAT}")

    utils.verbose_print("Concatenating clips into final reel...")
    ok = video.concatenate_clips(clip_paths, output_file, reencode_on_fail=True)
    for path in clip_paths:
        try:
            if os.path.isfile(path):
                os.remove(path)
        except OSError:
            pass
    return 1 if ok else 0


def setup_encoding() -> None:
    """Ensure UTF-8 encoding for stdout/stderr to handle unicode properly."""
    if sys.stdout.encoding.lower() != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

def parse_cli_mode_args(args: Any) -> Tuple[Optional[List[int]], Optional[int], Optional[int], Optional[List[Tuple[str, int]]]]:
    """Parse CLI arguments for line, range, and cell modes.
    
    Args:
        args: Parsed command-line arguments
        
    Returns:
        tuple: (cli_line_numbers, cli_range_start, cli_range_end, cli_cell_specs)
    """
    cli_line_numbers = None
    cli_range_start = None
    cli_range_end = None
    cli_cell_specs = None
    
    if args.lines:
        try:
            # Support both + and , as separators
            line_str = args.lines.replace(',', '+')
            cli_line_numbers = [int(num.strip()) for num in line_str.split('+')]
        except ValueError:
            utils.error_print(f'Invalid line numbers "{args.lines}". Use format: 1+4+5 or 1,4,5')
            sys.exit(1)
    
    if args.range:
        try:
            parts = args.range.split('-')
            if len(parts) != 2:
                raise ValueError('Range must have exactly two parts')
            cli_range_start = int(parts[0].strip())
            cli_range_end = int(parts[1].strip())
            if cli_range_start > cli_range_end:
                utils.error_print(f'Range start ({cli_range_start}) must be less than or equal to end ({cli_range_end})')
                sys.exit(1)
        except ValueError as e:
            utils.error_print(f'Invalid range "{args.range}". Use format: 1-10')
            sys.exit(1)
    
    if args.cell:
        try:
            cli_cell_specs = spreadsheet.parse_cell_specifications(args.cell)
        except ValueError as e:
            utils.error_print(f'Invalid cell specification: {e}')
            sys.exit(1)
    
    return (cli_line_numbers, cli_range_start, cli_range_end, cli_cell_specs)

def authenticate_google() -> Any:
    """Authenticate with Google Sheets API.
    
    Returns:
        Google client connection object
    """
    try:
        utils.debug_print('Attempting login...')
        gc = gspread.oauth(credentials_filename='credentials.json')
        utils.debug_print('Login successful!')
        return gc
    except gspread.exceptions.GSpreadException as e:
        utils.error_print("Could not authenticate with Google.",
            [f"Error details: {e}",
             f"Credentials file location: {os.path.join(os.getcwd(), 'credentials.json')}",
             "",
             "Troubleshooting steps:",
             "  1. Ensure 'credentials.json' exists in the working directory",
             "  2. Verify the credentials file is valid JSON",
             "  3. Check that the service account has access to Google Sheets API",
             "  4. For OAuth flow, delete any existing token files and re-authenticate"])
        sys.exit(1)

def _is_excel_worksheet(worksheet: Any) -> bool:
    """Return True if worksheet is the Excel adapter (local file, no URL)."""
    spread = getattr(worksheet, 'spreadsheet', None)
    return spread is not None and getattr(spread, 'url', None) is None


def select_worksheet(gc: Any, doc_list: List[str], args: Any, cli_mode: bool) -> Any:
    """Select worksheet based on command-line arguments or interactive selection.
    
    Args:
        gc: Google client connection
        doc_list: List of available spreadsheet names
        args: Parsed command-line arguments
        cli_mode: Whether running in CLI mode
        
    Returns:
        Worksheet object
    """
    worksheet = None
    if args.spreadsheet:
        # CLI-specified spreadsheet
        raw = args.spreadsheet.strip()
        raw_lower = raw.lower()
        if raw_lower == config.COMMAND_EXCEL:
            # -s excel: use single .xlsx in cwd, else error
            paths = excel_io.list_excel_in_cwd()
            if not paths:
                utils.error_print('No .xlsx files found in the current directory.',
                    ['Place an Excel file (.xlsx) in the working directory or use -s path/to/file.xlsx'])
                sys.exit(1)
            if len(paths) > 1:
                utils.error_print(f'Multiple .xlsx files found ({len(paths)}). Specify one with -s path/to/file.xlsx',
                    [os.path.basename(p) for p in paths])
                sys.exit(1)
            worksheet = excel_io.open_excel_workbook(paths[0])
            if not worksheet:
                sys.exit(1)
        elif raw_lower.endswith('.xlsx'):
            # -s path/to/file.xlsx
            path = os.path.join(os.getcwd(), raw) if not os.path.isabs(raw) else raw
            worksheet = excel_io.open_excel_workbook(path)
            if not worksheet:
                utils.error_print(f'Could not open Excel file "{args.spreadsheet}"')
                sys.exit(1)
        elif args.spreadsheet.startswith(config.COMMAND_HTTP_PREFIX):
            worksheet = open_spreadsheet_by_url(gc, args.spreadsheet)
        elif args.spreadsheet.isdigit():
            worksheet = open_spreadsheet_by_index(gc, doc_list, int(args.spreadsheet))
        else:
            worksheet = open_spreadsheet_by_name(gc, doc_list, args.spreadsheet)
        
        if not worksheet:
            utils.error_print(f'Could not find or open spreadsheet "{args.spreadsheet}"')
            sys.exit(1)
    else:
        # Auto-connect if working directory name matches a spreadsheet
        cwd_name = os.path.basename(os.getcwd())
        worksheet = open_spreadsheet_by_name(gc, doc_list, cwd_name)
        if worksheet:
            utils.verbose_print(f'\nAuto-connecting to spreadsheet: {worksheet.spreadsheet.title}')
        elif cli_mode:
            # CLI mode requires a spreadsheet - can't prompt interactively
            utils.error_print('No spreadsheet found matching working directory name.',
                ['Use -s to specify a spreadsheet name, URL, or index.'])
            sys.exit(1)
        else:
            worksheet = select_spreadsheet(gc, doc_list)
    
    if worksheet and config.DEBUGGING:
        ic(worksheet.title)
    if _is_excel_worksheet(worksheet):
        utils.verbose_print('\nUsing local Excel file.')
    else:
        utils.verbose_print('\nConnected to Google Drive!')
    return worksheet

def run_cli_mode(worksheet: Any, args: Any, cli_line_numbers: Optional[List[int]], cli_range_start: Optional[int], cli_range_end: Optional[int], cli_cell_specs: Optional[List[Tuple[str, int]]]) -> None:
    """Execute CLI mode - run once and exit.

    Args:
        worksheet: Selected worksheet
        args: Parsed command-line arguments
        cli_line_numbers: Parsed line numbers (if line mode)
        cli_range_start: Range start (if range mode)
        cli_range_end: Range end (if range mode)
        cli_cell_specs: Parsed cell specifications (if cell mode)
    """
    skip_prompts = args.yes

    if args.batch:
        clips_list = spreadsheet.generate_list(worksheet, 'batch', skip_prompts=skip_prompts)
    elif args.lines:
        clips_list = spreadsheet.generate_list(worksheet, 'line', line_numbers=cli_line_numbers, skip_prompts=skip_prompts)
    elif args.range:
        clips_list = spreadsheet.generate_list(worksheet, 'range', range_start=cli_range_start, range_end=cli_range_end, skip_prompts=skip_prompts)
    elif args.cell:
        clips_list = spreadsheet.generate_list(worksheet, 'cell', cell_specs=cli_cell_specs, skip_prompts=skip_prompts)
    elif args.participant:
        clips_list = spreadsheet.generate_list(worksheet, 'participant', participant_id=args.participant, skip_prompts=skip_prompts)
    elif args.reel:
        clips_list = spreadsheet.generate_list(worksheet, 'reel', reel_input=args.reel, skip_prompts=skip_prompts)
    else:
        clips_list = []

    if args.reel:
        videos_generated = process_reel(clips_list)
    else:
        videos_generated = process_clips(clips_list)
    
    if not config.REENCODING:
        utils.verbose_print('* No re-encoding done, expect:\n- inaccurate start and end timings\n- lossy frames until first keyframe\n- bad timecodes at the end\n')
    if args.reel:
        utils.info_print(f'All done, created 1 reel!\nFiles are in {os.getcwd()}\n')
    else:
        utils.info_print(f'All done, created {videos_generated} videos!\nFiles are in {os.getcwd()}\n')

def run_interactive_mode(worksheet: Any) -> None:
    """Execute interactive mode - main processing loop.

    Args:
        worksheet: Selected worksheet
    """
    while True:
        clips_list, is_reel, reel_output_file = select_mode_and_generate(worksheet)
        if is_reel:
            videos_generated = process_reel(clips_list, output_file=reel_output_file)
        else:
            videos_generated = process_clips(clips_list)

        if not config.REENCODING:
            utils.info_print('* No re-encoding done, expect:\n- inaccurate start and end timings\n- lossy frames until first keyframe\n- bad timecodes at the end\n')
        if is_reel:
            utils.info_print(f'All done, created 1 reel!\nFiles are in {os.getcwd()}\n')
        else:
            utils.info_print(f'All done, created {videos_generated} videos!\nFiles are in {os.getcwd()}\n')

        yn = input('Continue working (y) or quit the program (n)? y/n\n>> ')
        if yn == 'n':
            break

def main() -> None:
    """Main entry point for clipgen."""
    setup_encoding()
    
    # Parse command-line arguments
    args = utils.parse_arguments()
    if config.DEBUGGING:
        ic(args)
    
    # Determine if running in CLI mode (any mode argument provided)
    cli_mode = args.batch or args.lines or args.range or args.cell or args.participant or args.reel
    
    # Set verbose mode: silent by default in CLI mode, verbose in interactive mode
    config.VERBOSE = not cli_mode or args.verbose
    
    # Parse CLI arguments for line, range, and cell modes
    cli_line_numbers, cli_range_start, cli_range_end, cli_cell_specs = parse_cli_mode_args(args)
    
    # Change working directory to place of python script
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    utils.verbose_print('-------------------------------------------------------------------------------')
    utils.verbose_print(f'Welcome to clipgen v{config.VERSIONNUM}\nWorking directory: {os.getcwd()}\nPlace video files and the credentials.json file in this directory.')
    utils.debug_print('Debug mode is ON. Several limitations apply and more things will be printed.')
    
    # Authenticate with Google
    gc = authenticate_google()

    # Get document list and select spreadsheet
    doc_list = google_api.get_all_spreadsheets(gc).split(',')
    worksheet = select_worksheet(gc, doc_list, args, cli_mode)

    # Execute based on mode
    if cli_mode:
        run_cli_mode(worksheet, args, cli_line_numbers, cli_range_start, cli_range_end, cli_cell_specs)
    else:
        run_interactive_mode(worksheet)

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        utils.info_print('\nInterrupted by user')
        sys.exit(0)
