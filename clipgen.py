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
from typing import Any, Callable, List, NamedTuple, Optional, Set, Tuple

import gspread
from icecream import ic

import config
import excel_io
import files
import google_api
import spreadsheet
import utils
import video

# ---- Mode configuration ----

MODE_ALIASES = {
    'b': 'batch', 'batch': 'batch',
    'l': 'line', 'line': 'line',
    'r': 'range', 'range': 'range',
    'c': 'category', 'cat': 'category', 'category': 'category',
    'ce': 'cell', 'cell': 'cell',
    'p': 'participant', 'participant': 'participant',
    'f': 'filter', 'filter': 'filter',
    's': 'screen', 'screen': 'screen',
    'g': 'gif', 'gif': 'gif',
    're': 'reel', 'reel': 'reel',
    'rl': 'reellate', 'reellate': 'reellate',
    'br': 'browse', 'browse': 'browse',
}

FORMAT_MODE_ALIASES = {
    alias: mode for alias, mode in MODE_ALIASES.items()
    if mode in {'batch', 'line', 'range', 'category', 'cell', 'participant', 'filter'}
}


# ---- CLI data structures and runtime utilities ----

class CliModeArgs(NamedTuple):
    line_numbers: Optional[List[int]]
    range_start: Optional[int]
    range_end: Optional[int]
    cell_specs: Optional[List[Tuple[str, int]]]


def parse_cli_mode_args(args: Any) -> CliModeArgs:
    """Parse CLI arguments for line, range, and cell modes.
    
    Args:
        args: Parsed command-line arguments
        
    Returns:
        Parsed mode argument values as CliModeArgs
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
        except ValueError:
            utils.error_print(f'Invalid range "{args.range}". Use format: 1-10')
            sys.exit(1)
    
    if args.cell:
        try:
            cli_cell_specs = spreadsheet.parse_cell_specifications(args.cell)
        except ValueError as e:
            utils.error_print(f'Invalid cell specification: {e}')
            sys.exit(1)
    
    return CliModeArgs(cli_line_numbers, cli_range_start, cli_range_end, cli_cell_specs)


def setup_encoding() -> None:
    """Ensure UTF-8 encoding for stdout/stderr to handle unicode properly."""
    encoding = sys.stdout.encoding
    if not encoding or encoding.lower() != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


def get_runtime_working_dir() -> str:
    """Return the runtime working directory.

    Source runs use the script directory; frozen one-file builds use the
    executable directory so local assets resolve from where the binary lives.
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


# ---- Spreadsheet opening and selection utilities ----

def _open_worksheet(gspread_client: Any, open_callable: Callable[[], Any], error_context: str):
    """Try to open a worksheet via a callable; catch gspread errors and print a consistent message.
    
    Args:
        gspread_client: Google client connection
        open_callable: Callable that takes no args and returns a gspread Spreadsheet
        error_context: Short description for error message (e.g. "by URL", "at index 3")
        
    Returns:
        Worksheet object or None if error
    """
    try:
        return google_api.get_worksheet(open_callable())
    except (gspread.SpreadsheetNotFound, gspread.exceptions.APIError, gspread.exceptions.GSpreadException) as e:
        utils.error_print(f"Could not open spreadsheet {error_context}: {e}")
        return None


def open_spreadsheet_by_url(gspread_client: Any, url: str) -> Optional[Any]:
    """Open a spreadsheet by URL.
    
    Args:
        gspread_client: Google client connection
        url: Spreadsheet URL
        
    Returns:
        Worksheet object or None if error
    """
    return _open_worksheet(gspread_client, lambda: gspread_client.open_by_url(url), "by URL")


def open_spreadsheet_by_index(gspread_client: Any, doc_list: List[str], index: int) -> Optional[Any]:
    """Open a spreadsheet by index number.
    
    Args:
        gspread_client: Google client connection
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
    return _open_worksheet(gspread_client, lambda: gspread_client.open(doc_name), f"at index {index}")


def open_spreadsheet_by_name(gspread_client: Any, doc_list: List[str], name: str) -> Optional[Any]:
    """Open a spreadsheet by name.
    
    Args:
        gspread_client: Google client connection
        doc_list: List of spreadsheet names
        name: Spreadsheet name to search for
        
    Returns:
        Worksheet object or None if not found
    """
    chosen_index = google_api.find_spreadsheet_by_name(name, doc_list)
    if chosen_index >= 0:
        matched_name = doc_list[chosen_index].strip()
        utils.verbose_print(f'Opening document: {matched_name}')
        return _open_worksheet(gspread_client, lambda: gspread_client.open(matched_name), f"'{name}'")
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

def _handle_spreadsheet_command(gspread_client: Any, doc_list: List[str], input_name: str) -> Optional[Any]:
    """Handle one spreadsheet selection command. Returns worksheet when one was opened, None to show prompt again."""
    # Handle 'excel' for local .xlsx
    if input_name.strip().lower() == config.COMMAND_EXCEL:
        return excel_io.select_excel_file()
    # Handle URL
    if input_name.startswith(config.COMMAND_HTTP_PREFIX):
        return open_spreadsheet_by_url(gspread_client, input_name)
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
        latest_spreadsheet_name = google_api.get_all_spreadsheets(gspread_client).split(',')[0]
        return open_spreadsheet_by_name(gspread_client, doc_list, latest_spreadsheet_name)
    # Handle numeric index
    if input_name[0].isdigit():
        return open_spreadsheet_by_index(gspread_client, doc_list, int(input_name))
    # Handle 'settings' command
    if input_name.startswith(config.COMMAND_SETTINGS):
        utils.set_program_settings()
        return None
    # Handle name search
    return open_spreadsheet_by_name(gspread_client, doc_list, input_name)


def select_spreadsheet(gspread_client: Any, doc_list: List[str]) -> Any:
    """Interactive spreadsheet selection. Returns the selected worksheet."""
    consecutive_open_failures = 0
    while True:
        try:
            input_name = utils.read_user_input(
                f"\nPlease enter the index, name, URL, or '{config.COMMAND_EXCEL}' for local file "
                f"('{config.COMMAND_LIST_ALL}' for list, '{config.COMMAND_LIST_NEW}' for list of newest, "
                f"'{config.COMMAND_OPEN_LAST}' to immediately open latest, '{config.COMMAND_SETTINGS}' to change settings):\n>> "
            )
        except utils.BackToTop:
            # Treat 'top'/'back' at spreadsheet selection as a request to exit.
            raise utils.QuitProgram()
        try:
            worksheet = _handle_spreadsheet_command(gspread_client, doc_list, input_name)
            if worksheet is not None:
                return worksheet
        except (gspread.SpreadsheetNotFound, gspread.exceptions.APIError, gspread.exceptions.GSpreadException) as e:
            consecutive_open_failures += 1
            handle_error_message(consecutive_open_failures, e)
        except Exception as e:
            utils.error_print(f"Could not open document: {e}")


# ---- Interactive selection helpers (reel, reel-late, screen/gif, browse) ----

def _prompt_timeline_participant_selection(worksheet: Any) -> Optional[str]:
    """Prompt user to pick exactly one participant for timeline reels."""
    header_result = spreadsheet.validate_spreadsheet_headers(worksheet)
    if header_result is None:
        return None

    id_cell, _, _ = header_result
    sheet_data = worksheet.get_all_values()
    header_row = sheet_data[id_cell.row - 1] if id_cell.row > 0 and sheet_data else []
    num_participants = spreadsheet.get_num_participants(header_row, id_cell, worksheet.col_count)
    available_list = spreadsheet.get_participant_list(header_row, id_cell, num_participants)
    if not available_list:
        utils.info_print('\nNo participants found in the spreadsheet.')
        return None

    utils.info_print('\nTimeline requires exactly one participant.')
    utils.info_print('Available participants:')
    for i, pid in enumerate(available_list, 1):
        utils.info_print(f'  {i}. {pid}')

    while True:
        selection = utils.read_user_input('\nEnter one participant number or ID (e.g., 1 or P01):\n>> ')
        if not selection:
            utils.info_print('Please enter one participant.')
            continue
        tokens = spreadsheet.parse_participant_selection(selection)
        if len(tokens) != 1:
            utils.info_print('Please provide exactly one participant.')
            continue

        token = tokens[0]
        if token.isdigit():
            idx = int(token)
            if 1 <= idx <= len(available_list):
                return available_list[idx - 1]
            utils.info_print(f'Not found: {token}. Available: {", ".join(available_list)}')
            continue

        col_idx = spreadsheet.find_participant_column(header_row, id_cell, token)
        if col_idx is None:
            utils.info_print(f'Not found: {token}. Available: {", ".join(available_list)}')
            continue
        if col_idx < len(header_row):
            return utils.normalize_participant_id(header_row[col_idx])
        return token


def _run_reel_mode_interactive(worksheet: Any) -> Tuple[List[Any], bool, Optional[str]]:
    """Run reel mode UI: instructions, input, generate_list, preview, confirm, output filename.
    Returns (clips_list, True, reel_output_file or None) when user confirms; caller may loop on continue.
    """
    utils.info_print('\nReel mode: combine selectors into one video. Syntax:')
    utils.info_print('  batch                    - all clips')
    utils.info_print('  filter                   - key-marked clips only')
    utils.info_print('  timeline                 - chronological reel (requires exactly one participant)')
    utils.info_print('  11, 12, 13-16, 18        - lines and ranges')
    utils.info_print('  "Observations", "Onboarding" - categories (quoted)')
    utils.info_print('  P01.11, P02.15           - cells (participant.row)')
    utils.info_print('  P01, P02                 - participants (all their clips)')
    utils.info_print('  Example: timeline, P01, 11, 13-16, "Observations"')
    reel_input = utils.read_user_input('\nEnter reel selectors (combine any of the above, comma-separated):\n>> ')
    if not reel_input:
        utils.info_print('No input. Skipping reel.')
        return ([], False, None)

    parsed_reel = spreadsheet.parse_reel_input(reel_input)
    if parsed_reel['timeline']:
        if len(parsed_reel['participants']) > 1:
            utils.error_print(
                "Timeline selector supports only one participant.",
                ["Please provide exactly one participant (e.g., timeline, P01)."],
            )
            return ([], False, None)
        if len(parsed_reel['participants']) == 0:
            selected_pid = _prompt_timeline_participant_selection(worksheet)
            if not selected_pid:
                return ([], False, None)
            reel_input = f'{reel_input}, {selected_pid}'
            parsed_reel['participants'] = [selected_pid]

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
    yn = utils.read_user_input('\nGenerate reel? y/n\n>> ')
    if yn != 'y':
        return ([], False, None)

    study_name = clips_list[0].get('study', '').strip() if clips_list else ''
    default_filename = f'{study_name}_reel{config.FILEFORMAT}' if study_name else f'reel{config.FILEFORMAT}'
    if parsed_reel['timeline'] and parsed_reel['participants']:
        timeline_pid = utils.normalize_participant_id(parsed_reel['participants'][0]).strip()
        if study_name and timeline_pid:
            default_filename = f'{study_name}_{timeline_pid}_timeline{config.FILEFORMAT}'
        elif timeline_pid:
            default_filename = f'{timeline_pid}_timeline{config.FILEFORMAT}'
        else:
            default_filename = f'timeline{config.FILEFORMAT}'

    output_file = utils.read_user_input(f'\nOutput filename (Enter for default {default_filename}):\n>> ')
    reel_output_file = None
    if output_file:
        reel_output_file = output_file if output_file.endswith(config.FILEFORMAT) else output_file + config.FILEFORMAT
    else:
        reel_output_file = files.get_unique_filename(default_filename)
    return (clips_list, True, reel_output_file)


def _parse_clip_selection(selection_input: str, num_clips: int) -> List[int]:
    """Parse user selection input into list of clip indices.
    
    Supports formats: "A + B + C", "A, B, C", "A B C", or mixed.
    
    Args:
        selection_input: User's selection string (e.g., "A + B" or "A, C")
        num_clips: Total number of available clips (for validation)
        
    Returns:
        List of valid 0-based indices, deduplicated and in selection order
    """
    # Normalize separators to spaces
    normalized = selection_input.replace('+', ' ').replace(',', ' ')
    tokens = normalized.split()
    
    indices = []
    seen = set()
    for token in tokens:
        idx = utils.letter_to_index(token)
        if idx >= 0 and idx < num_clips and idx not in seen:
            indices.append(idx)
            seen.add(idx)
    return indices


def _run_reellate_mode_interactive() -> Tuple[bool, Optional[str]]:
    """Run reel-late mode UI: discover clips, display list, select, concatenate.
    
    Returns (True, output_file) when reel was generated; (False, None) otherwise.
    """
    clips = files.discover_clips()
    
    if not clips:
        utils.info_print('\nNo clips found in the working directory.')
        utils.info_print('  Source videos (like study_P01.mp4) are excluded.')
        utils.info_print('  Generate some clips first, then use this mode to combine them.')
        return (False, None)
    
    utils.info_print('\nReel-late mode: combine existing clips into a highlight reel.')
    utils.info_print(f'\nFound {len(clips)} clip(s) in {os.getcwd()}:\n')
    
    # Display indexed list
    for i, clip in enumerate(clips):
        letter = utils.index_to_letter(i)
        utils.info_print(f'  {letter}. "{clip}"')
    
    utils.info_print('\nSelect clips to include (order preserved). Syntax:')
    utils.info_print('  A + B + C    - combine clips A, B, and C')
    utils.info_print('  A, B, C      - same as above')
    utils.info_print('  A B C        - same as above')
    
    selection_input = utils.read_user_input('\nEnter clip selection:\n>> ')
    if not selection_input:
        utils.info_print('No selection. Skipping reel.')
        return (False, None)
    
    indices = _parse_clip_selection(selection_input, len(clips))
    if not indices:
        utils.warning_print('No valid clips selected.',
            ['Use letters from the list above (e.g., A + B + C)'])
        return (False, None)
    
    selected_clips = [clips[i] for i in indices]
    
    # Preview selection
    utils.info_print(f'\nSelected {len(selected_clips)} clip(s):')
    for i, clip in enumerate(selected_clips):
        utils.info_print(f'  {i+1}. "{clip}"')
    
    yn = utils.read_user_input('\nGenerate reel from these clips? y/n\n>> ').lower()
    if yn != 'y':
        utils.info_print('Cancelled.')
        return (False, None)
    
    output_file = utils.read_user_input('\nOutput filename (Enter for default "reel.mp4"):\n>> ')
    if not output_file:
        output_file = files.get_unique_filename(f'reel{config.FILEFORMAT}')
    elif not output_file.endswith(config.FILEFORMAT):
        output_file = output_file + config.FILEFORMAT
    
    utils.verbose_print(f'\nConcatenating {len(selected_clips)} clips into {output_file}...')
    ok = video.concatenate_clips(selected_clips, output_file, reencode_on_fail=True)
    
    if ok:
        return (True, output_file)
    return (False, None)


def _run_format_mode_interactive(worksheet: Any, output_format: str) -> None:
    """Run interactive flow for screen/gif output formats.

    Prompts the user for a timestamp-selection mode, generates a clip list, then renders
    screenshots/GIFs using the regular clip processing pipeline.
    """
    format_display_name = 'Screenshot' if output_format == 'screen' else 'GIF'
    output_label = 'screenshots' if output_format == 'screen' else 'GIFs'

    utils.info_print(f'\n{format_display_name} mode: choose how to select timestamps.')
    while True:
        selection = utils.read_user_input(
            '\nSelect source rows for this output:\n'
            '  Modes: (b)atch, (r)ange, (c)ategory, (l)ine, (ce)ll, (p)articipant, (f)ilter\n'
            '  Or enter mixed selectors directly: e.g. 5, P01.11, 13-16, "Observations"\n>> '
        )
        if not selection:
            utils.info_print("  Please enter a mode or direct input (e.g. P01.11, 5, 7, 13-16, P01).")
            continue
        selection_lower = selection.lower()
        mode = FORMAT_MODE_ALIASES.get(selection_lower)
        if mode:
            clips_list = spreadsheet.generate_list(worksheet, mode)
            break

        detected_mode, detected_kwargs = spreadsheet.detect_mode_from_input(selection)
        if detected_mode in ('batch', 'line', 'range', 'category', 'cell', 'participant', 'filter'):
            utils.verbose_print(f"  {detected_mode.capitalize()} mode detected.")
            clips_list = spreadsheet.generate_list(worksheet, detected_mode, **detected_kwargs)
            break
        if detected_mode in ('reel', 'browse'):
            utils.info_print('  This mode is not available for screen/gif output. Choose batch/line/range/category/cell/participant/filter.')
            continue

        parsed = spreadsheet.parse_reel_input(selection)
        has_timeline = bool(parsed.get('timeline'))
        has_batch = bool(parsed.get('batch'))
        has_filter = bool(parsed.get('filter'))
        has_lines = len(parsed['lines']) > 0
        has_ranges = len(parsed['ranges']) > 0
        has_cells = len(parsed['cells']) > 0
        has_participants = len(parsed['participants']) > 0
        has_categories = len(parsed['categories']) > 0

        selector_types = [
            ('batch', has_batch),
            ('filter', has_filter),
            ('lines', has_lines),
            ('ranges', has_ranges),
            ('cells', has_cells),
            ('participants', has_participants),
            ('categories', has_categories),
        ]
        non_empty_types = [name for name, present in selector_types if present]

        if not non_empty_types:
            utils.info_print(f"  Unknown mode or input '{selection}'. Available modes:")
            utils.info_print("    b or batch   - Generate from all clips in the spreadsheet")
            utils.info_print("    r or range   - Generate from a range of rows")
            utils.info_print("    c or category - Generate by category")
            utils.info_print("    l or line    - Generate from specific line(s)")
            utils.info_print("    ce or cell   - Generate from specific cell(s) (e.g., P01.11)")
            utils.info_print("    p or participant - Generate all outputs for one participant")
            utils.info_print("    f or filter  - Generate only key-marked outputs")
            continue

        if has_timeline:
            utils.info_print('  Timeline selector is only available for reel/timeline modes.')
            utils.info_print('  Use reel/timeline for a single .mp4 output, not for screenshots/GIFs.')
            continue

        utils.verbose_print(
            "  Mixed selectors detected ("
            + ", ".join(non_empty_types)
            + "). Generating individual outputs from combined selectors.",
        )
        clips_list = spreadsheet.generate_list(worksheet, 'reel', reel_input=selection)
        break

    outputs_generated = process_clips(clips_list, output_format=output_format)
    utils.info_print(f'All done, created {outputs_generated} {output_label}!\nFiles are in {os.getcwd()}\n')


def select_mode_and_generate(worksheet: Any) -> Tuple[List[Any], bool, Optional[str]]:
    """Interactive mode selection. Returns (clips list, is_reel_mode, reel_output_file or None)."""
    while True:
        input_mode = utils.read_user_input(
            '\nEnter mode or input directly:\n'
            '  Modes: (b)atch, (r)ange, (c)ategory, (l)ine, (ce)ll, (p)articipant, (f)ilter, (s)creen, (g)if, (re)el, (rl) reel-late, (br)owse\n'
            '  Or enter mixed selectors directly: e.g. 5, P01.11, 13-16, "Observations"\n>> '
        )
        if not input_mode:
            utils.info_print("  Please enter a mode or direct input (e.g. P01.11, 5, 7, 13-16, P01).")
            continue
        input_lower = input_mode.strip().lower()
        try:
            # Only treat as explicit mode when input exactly matches a mode shortcut or name
            mode = MODE_ALIASES.get(input_lower)
            if mode == 'browse':
                spreadsheet.browse_spreadsheet(worksheet)
                return ([], False, None)
            if mode == 'reel':
                reel_selection_result = _run_reel_mode_interactive(worksheet)
                if reel_selection_result[1]:
                    return reel_selection_result
                continue
            if mode == 'reellate':
                success, output_file = _run_reellate_mode_interactive()
                if success:
                    utils.info_print(f'\nReel created: {output_file}')
                    utils.info_print(f'Files are in {os.getcwd()}\n')
                    return ([], False, None)
                continue
            if mode in ('screen', 'gif'):
                _run_format_mode_interactive(worksheet, mode)
                return ([], False, None)
            if mode:
                return (spreadsheet.generate_list(worksheet, mode), False, None)

            # Try implicit mode detection from input syntax for single-type inputs
            detected_mode, detected_kwargs = spreadsheet.detect_mode_from_input(input_mode)
            if detected_mode:
                utils.verbose_print(f"  {detected_mode.capitalize()} mode detected.")
                return (spreadsheet.generate_list(worksheet, detected_mode, **detected_kwargs), False, None)

            # Fallback: treat as mixed selector input using reel parsing, but keep outputs as individual clips
            parsed = spreadsheet.parse_reel_input(input_mode)
            has_batch = bool(parsed.get('batch'))
            has_filter = bool(parsed.get('filter'))
            has_timeline = bool(parsed.get('timeline'))
            has_lines = len(parsed['lines']) > 0
            has_ranges = len(parsed['ranges']) > 0
            has_cells = len(parsed['cells']) > 0
            has_participants = len(parsed['participants']) > 0
            has_categories = len(parsed['categories']) > 0

            selector_types = [
                ('batch', has_batch),
                ('filter', has_filter),
                ('lines', has_lines),
                ('ranges', has_ranges),
                ('cells', has_cells),
                ('participants', has_participants),
                ('categories', has_categories),
            ]
            non_empty_types = [name for name, present in selector_types if present]

            if not non_empty_types:
                utils.info_print(f"  Unknown mode or input '{input_mode}'. Available modes:")
                utils.info_print("    b or batch   - Generate all clips in the spreadsheet")
                utils.info_print("    r or range   - Generate clips from a range of rows")
                utils.info_print("    c or category - Generate clips by category")
                utils.info_print("    l or line    - Generate clips from specific line(s)")
                utils.info_print("    ce or cell   - Generate clips from specific cell(s) (e.g., P01.11)")
                utils.info_print("    p or participant - Generate all clips for one participant")
                utils.info_print("    f or filter  - Generate only key-marked clips/timestamps")
                utils.info_print("    s or screen  - Generate screenshots (.png)")
                utils.info_print("    g or gif     - Generate GIFs (.gif)")
                utils.info_print("    re or reel   - Combine selectors into one highlight reel video")
                utils.info_print("    rl or reellate - Combine existing clips into a highlight reel")
                utils.info_print("    br or browse - Browse spreadsheet rows interactively")
                continue

            if has_timeline:
                utils.info_print(
                    "  Timeline selector is only supported for reel/timeline modes.",
                )
                utils.info_print("  Use 're' or 'reel' for a combined reel video, or -T on the command line.")
                continue

            utils.verbose_print(
                "  Mixed selectors detected ("
                + ", ".join(non_empty_types)
                + "). Generating individual clips from combined selectors.",
            )
            clips_list = spreadsheet.generate_list(worksheet, 'reel', reel_input=input_mode)
            return (clips_list, False, None)
        except gspread.exceptions.GSpreadException as e:
            utils.error_print(f"Google Sheets API error: {e}")
            utils.debug_print(f"ERROR Message '{e}', Attempting reconnect")

# ---- Clip processing core pipeline ----

def _process_single_clip_segments(
    clip: Any,
    base_video: str,
    missing_videos: set,
    *,
    filename_prefix: str = "",
    output_format: str = "clip",
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
    output_paths: List[str] = []
    extension_map = {
        'clip': config.FILEFORMAT,
        'screen': '.png',
        'gif': '.gif',
    }
    file_extension = extension_map.get(output_format)
    if not file_extension:
        utils.error_print(f"Unsupported output format: '{output_format}'")
        return (generated, output_paths)

    template = f"{filename_prefix}[{clip['category']}] {clip['study']} {clip['participant']} {clip['desc']}{file_extension}"
    for start_time, end_time in clip['times']:
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
        if output_format == 'clip':
            ok = video.run_ffmpeg(
                input_file=base_video,
                output_file=out_name,
                start_pos=start_time,
                end_pos=end_time,
                reencode=config.REENCODING,
            )
        elif output_format == 'screen':
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
                output_paths.append(out_name)
    return (generated, output_paths)


def _print_reencoding_warning(printer: Callable[[str], None]) -> None:
    printer('* No re-encoding done, expect:\n- inaccurate start and end timings\n- lossy frames until first keyframe\n- bad timecodes at the end\n')


def _print_completion_message(outputs_generated: int, output_format: str, is_reel: bool) -> None:
    if is_reel:
        utils.info_print(f'All done, created 1 reel!\nFiles are in {os.getcwd()}\n')
        return

    if output_format == 'screen':
        utils.info_print(f'All done, created {outputs_generated} screenshots!\nFiles are in {os.getcwd()}\n')
    elif output_format == 'gif':
        utils.info_print(f'All done, created {outputs_generated} GIFs!\nFiles are in {os.getcwd()}\n')
    else:
        utils.info_print(f'All done, created {outputs_generated} videos!\nFiles are in {os.getcwd()}\n')


def _check_source_video(clip: Any, missing_videos: set, skip_detail: str) -> Optional[str]:
    base_video = f"{clip['study']}_{clip['participant']}{config.FILEFORMAT}"
    if os.path.isfile(base_video):
        return base_video

    if base_video not in missing_videos:
        missing_videos.add(base_video)
        utils.error_print(
            f"Source video file not found: '{base_video}'",
            [
                f"Expected location: {os.path.join(os.getcwd(), base_video)}",
                skip_detail,
            ],
        )
    return None


def _prepare_and_check_clip(clip: Any, missing_videos: Set[str]) -> Tuple[Any, Optional[str]]:
    """Prepare one clip and validate that its source video exists.

    Returns:
        Tuple of (prepared clip dict, source video path or None).
        When None is returned for source video, the clip should be skipped.
    """
    clip = files.prepare_clip(clip)
    if not clip['times']:
        return (clip, None)

    base_video = _check_source_video(
        clip,
        missing_videos,
        f"Clips for participant '{clip['participant']}' in study '{clip['study']}' will be skipped.",
    )
    return (clip, base_video)


def _run_clip_pipeline(
    clips_list: List[Any],
    *,
    empty_warning: str,
    intro_message: str,
    task_label: str,
    per_clip_fn: Callable[[Any, Set[str]], Any],
    show_fallback_counter: bool = False,
) -> Tuple[List[Any], Set[str]]:
    """Run shared clip-processing pipeline and return per-clip results."""
    if not clips_list:
        utils.warning_print(empty_warning)
        return ([], set())

    utils.verbose_print(intro_message)
    missing_videos: Set[str] = set()

    def wrapped_process(clip: Any) -> Any:
        return per_clip_fn(clip, missing_videos)

    results = _process_with_progress(
        clips_list,
        task_label,
        wrapped_process,
        show_fallback_counter=show_fallback_counter,
    )
    if missing_videos:
        utils.verbose_print(f"* Missing source video files: {len(missing_videos)}")
    return (results, missing_videos)


def _process_with_progress(
    clips_list: List[Any],
    task_label: str,
    process_fn: Callable[[Any], Any],
    *,
    show_fallback_counter: bool = False,
) -> List[Any]:
    total_clips = len(clips_list)
    progress = utils.create_progress_bar()
    results: List[Any] = []
    
    if progress:
        with progress:
            task = progress.add_task(task_label, total=total_clips)
            for clip in clips_list:
                desc_preview = (clip.get('desc') or '')[:30]
                participant = clip.get('participant', '')
                progress.update(task, description=f"[{participant}] {desc_preview}...")
                results.append(process_fn(clip))
                progress.update(task, advance=1)
        return results
    
    for index, clip in enumerate(clips_list, start=1):
        if show_fallback_counter and config.VERBOSE and total_clips > 1:
            utils.verbose_print(f"Processing clip {index} of {total_clips}...")
        results.append(process_fn(clip))
    return results


def process_clips(clips_list: List[Any], output_format: str = "clip") -> int:
    """Process and generate outputs from the clips list. Returns count of files generated."""
    if config.DEBUGGING:
        ic(len(clips_list))
    
    def process_single_clip(clip: Any, missing_videos: Set[str]) -> Tuple[int, int]:
        """Process a single clip and return (generated, skipped)."""
        clip, base_video = _prepare_and_check_clip(clip, missing_videos)
        if not clip['times']:
            return (0, 1)
        if base_video is None:
            return (0, len(clip['times']))
    
        generated_count, _ = _process_single_clip_segments(
            clip,
            base_video,
            missing_videos,
            output_format=output_format,
            collect_paths=False,
        )
        skipped_count = len(clip['times']) - generated_count if generated_count < len(clip['times']) else 0
        return (generated_count, skipped_count)
    
    results, _ = _run_clip_pipeline(
        clips_list,
        empty_warning="No clips to process. No timestamps were found or selected.",
        intro_message='\n* ffmpeg is set to never prompt for input and will always overwrite.\n  Only warns if close to crashing.\n',
        task_label="Processing clips",
        per_clip_fn=process_single_clip,
        show_fallback_counter=True,
    )
    outputs_generated = sum(generated_count for generated_count, _ in results)
    outputs_skipped = sum(skipped_count for _, skipped_count in results)
    
    item_name = {
        'clip': 'video(s)',
        'screen': 'screenshot(s)',
        'gif': 'GIF(s)',
    }.get(output_format, 'file(s)')
    if outputs_skipped > 0:
        utils.verbose_print(f"\n* Summary: {outputs_generated} {item_name} generated, {outputs_skipped} skipped due to errors.")
    return outputs_generated


def process_reel(clips_list: List[Any], output_file: Optional[str] = None) -> int:
    """Process clips for reel mode: generate individual clips, concatenate into one video, clean up.

    Returns 1 if the reel was generated successfully, 0 otherwise.
    """
    if not clips_list:
        utils.warning_print("No clips to process for reel. No timestamps were found or selected.")
        return 0

    study_name = next(
        (
            (clip.get('study') or '').strip()
            for clip in clips_list
            if (clip.get('study') or '').strip()
        ),
        None,
    )

    def process_reel_clip(clip: Any, missing_videos: Set[str]) -> List[str]:
        """Process one clip for reel mode and return generated segment paths."""
        clip, base_video = _prepare_and_check_clip(clip, missing_videos)
        if base_video is None:
            return []
        _, segment_paths = _process_single_clip_segments(
            clip,
            base_video,
            missing_videos,
            filename_prefix="_reel_part_",
            collect_paths=True,
        )
        return segment_paths

    all_segment_paths, _ = _run_clip_pipeline(
        clips_list,
        empty_warning="No clips to process for reel. No timestamps were found or selected.",
        intro_message='\n* Reel mode: generating individual clips, then concatenating into one file.\n',
        task_label="Generating reel clips",
        per_clip_fn=process_reel_clip,
    )
    clip_paths = [segment_path for segment_paths in all_segment_paths for segment_path in segment_paths]
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



# ---- Google authentication and worksheet selection ----

def authenticate_google() -> Any:
    """Authenticate with Google Sheets API.
    
    Returns:
        Google client connection object
    """
    try:
        utils.debug_print('Attempting login...')
        gspread_client = gspread.oauth(credentials_filename='credentials.json')
        utils.debug_print('Login successful!')
        return gspread_client
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


def select_worksheet(gspread_client: Any, doc_list: List[str], args: Any, cli_mode: bool) -> Any:
    """Select worksheet based on command-line arguments or interactive selection.
    
    Args:
        gspread_client: Google client connection
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
            worksheet = open_spreadsheet_by_url(gspread_client, args.spreadsheet)
        elif args.spreadsheet.isdigit():
            worksheet = open_spreadsheet_by_index(gspread_client, doc_list, int(args.spreadsheet))
        else:
            worksheet = open_spreadsheet_by_name(gspread_client, doc_list, args.spreadsheet)
        
        if not worksheet:
            utils.error_print(f'Could not find or open spreadsheet "{args.spreadsheet}"')
            sys.exit(1)
    else:
        # Auto-connect if working directory name matches a spreadsheet
        cwd_name = os.path.basename(os.getcwd())
        worksheet = open_spreadsheet_by_name(gspread_client, doc_list, cwd_name)
        if worksheet:
            utils.verbose_print(f'\nAuto-connecting to spreadsheet: {worksheet.spreadsheet.title}')
        elif cli_mode:
            # CLI mode requires a spreadsheet - can't prompt interactively
            utils.error_print('No spreadsheet found matching working directory name.',
                ['Use -s to specify a spreadsheet name, URL, or index.'])
            sys.exit(1)
        else:
            worksheet = select_spreadsheet(gspread_client, doc_list)
    
    if worksheet and config.DEBUGGING:
        ic(worksheet.title)
    if _is_excel_worksheet(worksheet):
        utils.verbose_print('\nUsing local Excel file.')
    else:
        utils.verbose_print('\nConnected to Google Drive!')
    return worksheet

# ---- Mode runners and entry point ----

def run_cli_mode(worksheet: Any, args: Any, cli_mode_args: CliModeArgs) -> None:
    """Execute CLI mode - run once and exit.

    Args:
        worksheet: Selected worksheet
        args: Parsed command-line arguments
        cli_mode_args: Parsed line/range/cell arguments
    """
    skip_prompts = args.yes
    output_format = 'screen' if args.screen else 'gif' if args.gif else 'clip'
    mixed_selectors = getattr(args, 'mixed', None)

    if (args.reel or args.timeline) and output_format != 'clip':
        utils.error_print("Reel/timeline mode cannot be combined with --screen or --gif.",
            ["Use reel/timeline mode for a single .mp4 output, or use screen/gif with batch/line/range/category/cell/participant/filter selection."])
        sys.exit(1)

    if mixed_selectors:
        parsed_mixed = spreadsheet.parse_reel_input(mixed_selectors)
        if parsed_mixed.get('timeline'):
            utils.error_print(
                "Timeline selector is not supported in mixed mode.",
                [
                    "Use -T PARTICIPANT for a chronological reel,",
                    "or use -R with timeline selectors to create a single reel video.",
                ],
            )
            sys.exit(1)

    selection_mode_set = bool(
        args.batch
        or args.lines
        or args.range
        or args.cell
        or args.participant
        or args.filter
        or mixed_selectors
        or args.reel
        or args.timeline
    )

    if args.batch or (output_format != 'clip' and not selection_mode_set):
        clips_list = spreadsheet.generate_list(worksheet, 'batch', skip_prompts=skip_prompts)
    elif args.lines:
        clips_list = spreadsheet.generate_list(worksheet, 'line', line_numbers=cli_mode_args.line_numbers, skip_prompts=skip_prompts)
    elif args.range:
        clips_list = spreadsheet.generate_list(worksheet, 'range', range_start=cli_mode_args.range_start, range_end=cli_mode_args.range_end, skip_prompts=skip_prompts)
    elif args.cell:
        clips_list = spreadsheet.generate_list(worksheet, 'cell', cell_specs=cli_mode_args.cell_specs, skip_prompts=skip_prompts)
    elif args.participant:
        clips_list = spreadsheet.generate_list(worksheet, 'participant', participant_id=args.participant, skip_prompts=skip_prompts)
    elif args.filter:
        clips_list = spreadsheet.generate_list(worksheet, 'filter', skip_prompts=skip_prompts)
    elif mixed_selectors:
        clips_list = spreadsheet.generate_list(
            worksheet,
            'reel',
            reel_input=mixed_selectors,
            skip_prompts=skip_prompts,
        )
    elif args.reel:
        clips_list = spreadsheet.generate_list(worksheet, 'reel', reel_input=args.reel, skip_prompts=skip_prompts)
    elif args.timeline:
        clips_list = spreadsheet.generate_list(
            worksheet,
            'reel',
            reel_input=f'timeline, {args.timeline}',
            skip_prompts=skip_prompts,
        )
    else:
        clips_list = []

    if args.reel or args.timeline:
        reel_output_file = None
        if args.timeline:
            participant_id = utils.normalize_participant_id(args.timeline).strip()
            study_name = clips_list[0].get('study', '').strip() if clips_list else ''
            if study_name and participant_id:
                reel_output_file = files.get_unique_filename(f'{study_name}_{participant_id}_timeline{config.FILEFORMAT}')
            elif participant_id:
                reel_output_file = files.get_unique_filename(f'{participant_id}_timeline{config.FILEFORMAT}')
            else:
                reel_output_file = files.get_unique_filename(f'timeline{config.FILEFORMAT}')
        outputs_generated = process_reel(clips_list, output_file=reel_output_file)
    else:
        outputs_generated = process_clips(clips_list, output_format=output_format)

    if not config.REENCODING:
        _print_reencoding_warning(utils.verbose_print)
    _print_completion_message(outputs_generated, output_format, is_reel=bool(args.reel or args.timeline))

def run_interactive_mode(worksheet: Any) -> None:
    """Execute interactive mode - main processing loop.

    Args:
        worksheet: Selected worksheet
    """
    while True:
        try:
            clips_list, is_reel, reel_output_file = select_mode_and_generate(worksheet)
            if not clips_list and not is_reel:
                yn = utils.read_user_input('Continue working (y) or quit the program (n)? y/n\n>> ')
                if yn == 'n':
                    break
                continue
            if is_reel:
                outputs_generated = process_reel(clips_list, output_file=reel_output_file)
            else:
                outputs_generated = process_clips(clips_list)

            if not config.REENCODING:
                _print_reencoding_warning(utils.info_print)
            _print_completion_message(outputs_generated, output_format='clip', is_reel=is_reel)

            yn = utils.read_user_input('Continue working (y) or quit the program (n)? y/n\n>> ')
            if yn == 'n':
                break
        except utils.BackToTop:
            # Return to main mode selection prompt.
            continue
        except utils.QuitProgram:
            break

def main() -> None:
    """Main entry point for clipgen."""
    setup_encoding()
    
    # Parse command-line arguments
    args = utils.parse_arguments()
    if config.DEBUGGING:
        ic(args)
    
    # Determine if running in CLI mode (any mode argument provided)
    mixed_selectors = getattr(args, 'mixed', None)
    cli_mode = (
        args.batch
        or args.lines
        or args.range
        or args.cell
        or args.participant
        or args.filter
        or mixed_selectors
        or args.reel
        or args.timeline
        or args.screen
        or args.gif
    )
    
    # Set verbose mode: silent by default in CLI mode, verbose in interactive mode
    config.VERBOSE = not cli_mode or args.verbose
    
    # Parse CLI arguments for line, range, and cell modes
    cli_mode_args = parse_cli_mode_args(args)
    
    # Change working directory to runtime location (script/executable)
    os.chdir(get_runtime_working_dir())
    utils.verbose_print('-------------------------------------------------------------------------------')
    utils.verbose_print(f'Welcome to clipgen v{config.VERSIONNUM}\nWorking directory: {os.getcwd()}\nPlace video files and the credentials.json file in this directory.')
    utils.debug_print('Debug mode is ON. Several limitations apply and more things will be printed.')

    if not video.check_ffmpeg_tools_available():
        sys.exit(1)
    
    # Authenticate with Google
    gspread_client = authenticate_google()

    # Get document list and select spreadsheet
    doc_list = google_api.get_all_spreadsheets(gspread_client).split(',')
    worksheet = select_worksheet(gspread_client, doc_list, args, cli_mode)

    # Execute based on mode
    if cli_mode:
        run_cli_mode(worksheet, args, cli_mode_args)
    else:
        try:
            run_interactive_mode(worksheet)
        except utils.QuitProgram:
            utils.info_print('Exiting on user request.')

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        utils.info_print('\nInterrupted by user')
        sys.exit(0)
