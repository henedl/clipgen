# -*- coding: utf-8 -*-
"""clipgen - Video clip generator from Google Sheets timestamps.

This program will help quickly create video snippets from longer video files, based on timestamps in a spreadsheet!
Check out README.md for more detailed information about setting up and using clipgen.

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
import files
import google_api
import spreadsheet
import utils
import video


def open_spreadsheet_by_url(gc: Any, url: str) -> Optional[Any]:
    """Open a spreadsheet by URL.
    
    Args:
        gc: Google client connection
        url: Spreadsheet URL
        
    Returns:
        Worksheet object or None if error
    """
    try:
        return google_api.get_worksheet(gc.open_by_url(url))
    except (gspread.SpreadsheetNotFound, gspread.exceptions.APIError, gspread.exceptions.GSpreadException) as e:
        utils.error_print(f"Could not open spreadsheet by URL: {e}")
        return None


def open_spreadsheet_by_index(gc: Any, doc_list: List[str], index: int) -> Optional[Any]:
    """Open a spreadsheet by index number.
    
    Args:
        gc: Google client connection
        doc_list: List of spreadsheet names
        index: Index number (1-based)
        
    Returns:
        Worksheet object or None if error
    """
    try:
        if index < 1 or index > len(doc_list):
            utils.error_print(f"Invalid index {index}. Must be between 1 and {len(doc_list)}")
            return None
        chosen_index = index - 1
        doc_name = doc_list[chosen_index].strip()
        utils.verbose_print(f'Opening document: {doc_name}')
        return google_api.get_worksheet(gc.open(doc_name))
    except (gspread.SpreadsheetNotFound, gspread.exceptions.APIError, gspread.exceptions.GSpreadException) as e:
        utils.error_print(f"Could not open spreadsheet at index {index}: {e}")
        return None


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
        try:
            return google_api.get_worksheet(gc.open(matched_name))
        except (gspread.SpreadsheetNotFound, gspread.exceptions.APIError, gspread.exceptions.GSpreadException) as e:
            utils.error_print(f"Could not open spreadsheet '{name}': {e}")
            return None
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

def handle_error_message(input_file_fails: int, e: Exception) -> None:
    """Handle error messages with progressive detail based on failure count."""
    if input_file_fails == 1:
        utils.error_print(f"Could not access spreadsheet: {e}", 
            [f"Please try again. Type '{config.COMMAND_LIST_ALL}' to see available documents."])
    elif input_file_fails == 2:
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

def select_spreadsheet(gc: Any, doc_list: List[str]) -> Any:
    """Interactive spreadsheet selection. Returns the selected worksheet."""
    input_file_fails = 0
    
    while True:
        input_name = input(f"\nPlease enter the index, name, URL or key of the spreadsheet ('{config.COMMAND_LIST_ALL}' for list, '{config.COMMAND_LIST_NEW}' for list of newest, '{config.COMMAND_OPEN_LAST}' to immediately open latest, '{config.COMMAND_SETTINGS}' to change settings):\n>> ")
        try:
            # Handle URL
            if input_name.startswith(config.COMMAND_HTTP_PREFIX):
                worksheet = open_spreadsheet_by_url(gc, input_name)
                if worksheet:
                    return worksheet
                continue
            
            # Handle 'all' command
            if input_name.startswith(config.COMMAND_LIST_ALL):
                handle_list_all_command(doc_list)
                continue
            
            # Handle 'new' command
            if input_name.startswith(config.COMMAND_LIST_NEW):
                handle_list_new_command(doc_list)
                continue
            
            # Handle 'last' command
            if input_name.startswith(config.COMMAND_OPEN_LAST):
                latest = google_api.get_all_spreadsheets(gc).split(',')[0]
                worksheet = open_spreadsheet_by_name(gc, doc_list, latest)
                if worksheet:
                    return worksheet
                continue
            
            # Handle numeric index
            if input_name[0].isdigit():
                worksheet = open_spreadsheet_by_index(gc, doc_list, int(input_name))
                if worksheet:
                    return worksheet
                continue
            
            # Handle 'settings' command
            if input_name.startswith(config.COMMAND_SETTINGS):
                utils.set_program_settings()
                continue
            
            # Handle name search
            worksheet = open_spreadsheet_by_name(gc, doc_list, input_name)
            if worksheet:
                return worksheet
                
        except (gspread.SpreadsheetNotFound, gspread.exceptions.APIError, gspread.exceptions.GSpreadException) as e:
            input_file_fails += 1
            handle_error_message(input_file_fails, e)

def select_mode_and_generate(worksheet: Any) -> List[Any]:
    """Interactive mode selection. Returns the clips list for processing."""
    mode_map = {
        'b': 'batch', 'batch': 'batch',
        'l': 'line', 'line': 'line',
        'r': 'range', 'range': 'range',
        'c': 'category', 'cat': 'category', 'category': 'category',
        'br': 'browse', 'browse': 'browse',
        'test': 'test'
    }
    
    while True:
        input_mode = input('\nSelect mode: (b)atch, (r)ange, (c)ategory, (l)ine, or (br)owse\n>> ').strip().lower()
        
        if not input_mode:
            utils.info_print("  Please enter a mode (b, r, c, l, or br).")
            continue
        
        try:
            # Check for two-character modes first (like 'br' for browse)
            mode = mode_map.get(input_mode[:2]) or mode_map.get(input_mode[0]) or mode_map.get(input_mode)
            if mode == 'browse':
                # Browse mode doesn't generate clips, just displays data
                spreadsheet.browse_spreadsheet(worksheet)
                return []
            elif mode:
                return spreadsheet.generate_list(worksheet, mode)
            else:
                utils.info_print(f"  Unknown mode '{input_mode}'. Available modes:")
                utils.info_print("    b or batch   - Generate all clips in the spreadsheet")
                utils.info_print("    r or range   - Generate clips from a range of rows")
                utils.info_print("    c or category - Generate clips by category")
                utils.info_print("    l or line    - Generate clips from specific line(s)")
                utils.info_print("    br or browse - Browse spreadsheet rows interactively")
        except gspread.exceptions.GSpreadException as e:
            utils.error_print(f"Google Sheets API error: {e}")
            utils.debug_print(f"ERROR Message '{e}', Attempting reconnect")

def process_clips(clips_list: List[Any]) -> int:
    """Process and generate video clips from the clips list. Returns count of videos generated."""
    ic(len(clips_list))
    # Check if clips_list is empty
    if not clips_list:
        utils.warning_print("No clips to process. No timestamps were found or selected.")
        return 0
    
    utils.verbose_print('\n* ffmpeg is set to never prompt for input and will always overwrite.\n  Only warns if close to crashing.\n')
    videos_generated = 0
    videos_skipped = 0
    missing_videos = set()  # Track unique missing video files

    for clip in clips_list:
        ic(clip)
        clip = files.clean_issue(clip)
        
        # Skip if no valid timestamps were parsed
        if not clip['times']:
            videos_skipped += 1
            continue
        
        base_video = f"{clip['study']}_{clip['participant']}{config.FILEFORMAT}"
        
        # Check if base video exists (only warn once per unique file)
        if not os.path.isfile(base_video) and base_video not in missing_videos:
            missing_videos.add(base_video)
            utils.error_print(f"Source video file not found: '{base_video}'",
                [f"Expected location: {os.path.join(os.getcwd(), base_video)}",
                 f"Clips for participant '{clip['participant']}' in study '{clip['study']}' will be skipped."])
        
        for vid_in, vid_out in clip['times']:
            try:
                # Ensure all components are strings and handle unicode properly
                vid_name = files.get_unique_filename(
                    f"[{clip['category']}] {clip['study']} {clip['participant']} {clip['desc']}{config.FILEFORMAT}"
                )
                ic(vid_name)
            except (TypeError, UnicodeEncodeError, UnicodeDecodeError) as e:
                ic(e, clip)
                utils.error_print(f"Character encoding issue occurred: {e}",
                    [f"Category: '{clip['category']}', Study: '{clip['study']}', Participant: '{clip['participant']}'",
                     "Try simplifying the description or category names to use only ASCII characters."])
                videos_skipped += 1
                break

            completed = video.run_ffmpeg(
                input_file=base_video,
                output_file=vid_name,
                start_pos=vid_in,
                end_pos=vid_out,
                reencode=config.REENCODING
            )
            if completed:
                videos_generated += 1
            else:
                videos_skipped += 1
    
    # Report summary if any videos were skipped
    if videos_skipped > 0:
        utils.verbose_print(f"\n* Summary: {videos_generated} video(s) generated, {videos_skipped} skipped due to errors.")
    if missing_videos:
        utils.verbose_print(f"* Missing source video files: {len(missing_videos)}")

    return videos_generated

def setup_encoding() -> None:
    """Ensure UTF-8 encoding for stdout/stderr to handle unicode properly."""
    if sys.stdout.encoding.lower() != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

def parse_cli_mode_args(args: Any) -> Tuple[Optional[List[int]], Optional[int], Optional[int]]:
    """Parse CLI arguments for line and range modes.
    
    Args:
        args: Parsed command-line arguments
        
    Returns:
        tuple: (cli_line_numbers, cli_range_start, cli_range_end)
    """
    cli_line_numbers = None
    cli_range_start = None
    cli_range_end = None
    
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
    
    return (cli_line_numbers, cli_range_start, cli_range_end)

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
        if args.spreadsheet.startswith(config.COMMAND_HTTP_PREFIX):
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
    
    if worksheet:
        ic(worksheet.title)
    utils.verbose_print('\nConnected to Google Drive!')
    return worksheet

def run_cli_mode(worksheet: Any, args: Any, cli_line_numbers: Optional[List[int]], cli_range_start: Optional[int], cli_range_end: Optional[int]) -> None:
    """Execute CLI mode - run once and exit.
    
    Args:
        worksheet: Selected worksheet
        args: Parsed command-line arguments
        cli_line_numbers: Parsed line numbers (if line mode)
        cli_range_start: Range start (if range mode)
        cli_range_end: Range end (if range mode)
    """
    skip_prompts = args.yes
    
    if args.batch:
        clips_list = spreadsheet.generate_list(worksheet, 'batch', skip_prompts=skip_prompts)
    elif args.lines:
        clips_list = spreadsheet.generate_list(worksheet, 'line', line_numbers=cli_line_numbers, skip_prompts=skip_prompts)
    elif args.range:
        clips_list = spreadsheet.generate_list(worksheet, 'range', range_start=cli_range_start, range_end=cli_range_end, skip_prompts=skip_prompts)
    
    videos_generated = process_clips(clips_list)
    
    if not config.REENCODING:
        utils.verbose_print('* No re-encoding done, expect:\n- inaccurate start and end timings\n- lossy frames until first keyframe\n- bad timecodes at the end\n')
    utils.info_print(f'All done, created {videos_generated} videos!\nFiles are in {os.getcwd()}\n')

def run_interactive_mode(worksheet: Any) -> None:
    """Execute interactive mode - main processing loop.
    
    Args:
        worksheet: Selected worksheet
    """
    while True:
        clips_list = select_mode_and_generate(worksheet)
        videos_generated = process_clips(clips_list)

        if not config.REENCODING:
            utils.info_print('* No re-encoding done, expect:\n- inaccurate start and end timings\n- lossy frames until first keyframe\n- bad timecodes at the end\n')
        utils.info_print(f'All done, created {videos_generated} videos!\nFiles are in {os.getcwd()}\n')
        
        yn = input('Continue working (y) or quit the program (n)? y/n\n>> ')
        if yn == 'n':
            break

def main() -> None:
    """Main entry point for clipgen."""
    setup_encoding()
    
    # Parse command-line arguments
    args = utils.parse_arguments()
    ic(args)
    
    # Determine if running in CLI mode (any mode argument provided)
    cli_mode = args.batch or args.lines or args.range
    
    # Set verbose mode: silent by default in CLI mode, verbose in interactive mode
    config.VERBOSE = not cli_mode or args.verbose
    
    # Parse CLI arguments for line and range modes
    cli_line_numbers, cli_range_start, cli_range_end = parse_cli_mode_args(args)
    
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
        run_cli_mode(worksheet, args, cli_line_numbers, cli_range_start, cli_range_end)
    else:
        run_interactive_mode(worksheet)

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        utils.info_print('\nInterrupted by user')
        sys.exit(0)
