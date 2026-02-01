# -*- coding: utf-8 -*-
"""Utility functions for clipgen."""

import argparse
from datetime import datetime, timedelta
from typing import List, Optional, Tuple, Union

from icecream import ic

import config


def parse_arguments() -> argparse.Namespace:
    """Parse command-line arguments for non-interactive mode."""
    parser = argparse.ArgumentParser(
        description='clipgen - Video clip generator from Google Sheets timestamps.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python clipgen.py                    Interactive mode (default)
  python clipgen.py -b                 Batch mode - generate all clips
  python clipgen.py -l 5               Single line mode - line 5
  python clipgen.py -l 1+4+5           Multi-line mode - lines 1, 4, and 5
  python clipgen.py -l 1,4,5           Multi-line mode (comma separator)
  python clipgen.py -r 1-10            Range mode - lines 1 through 10
  python clipgen.py -c "P01.11"        Cell mode - single cell (participant P01, row 11)
  python clipgen.py -c "P01.11 + P03.11" Cell mode - multiple cells
  python clipgen.py -p P01             Participant mode - all clips for participant P01
  python clipgen.py -p "P01,P03"       Participant mode - all clips for P01 and P03
  python clipgen.py -b -s "Study Name" Batch mode with specific spreadsheet
  python clipgen.py -l 5 -y            Line mode, skip confirmation prompts
  python clipgen.py -b -v              Batch mode with verbose output
  python clipgen.py -R "11, 13-16, P01, \\"Observations\\""  Reel mode - one combined video

Note: Non-interactive mode (using -b, -l, -r, -c, -p, or -R) is silent by default,
      only showing errors and the final summary. Use -v for full output.
'''
    )
    
    # Mode arguments (mutually exclusive)
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument('-b', '--batch', action='store_true',
        help='Batch mode: generate all possible clips')
    mode_group.add_argument('-l', '--lines', type=str, metavar='LINES',
        help='Line mode: specify line numbers separated by + or , (e.g., 1+4+5 or 1,4,5)')
    mode_group.add_argument('-r', '--range', type=str, metavar='RANGE',
        help='Range mode: specify start-end line range (e.g., 1-10)')
    mode_group.add_argument('-c', '--cell', type=str, metavar='CELLS',
        help='Cell mode: specify cells as participant.row (e.g., P01.11 or P01.11 + P03.11)')
    mode_group.add_argument('-p', '--participant', type=str, metavar='ID',
        help='Participant mode: generate all clips for one or more participants (e.g., P01 or P01,P03)')
    mode_group.add_argument('-R', '--reel', type=str, metavar='SELECTORS',
        help='Reel mode: combine selectors (e.g. "11, 13-16, P01, \\"Observations\\"") into one video')
    
    # Optional arguments
    parser.add_argument('-s', '--spreadsheet', type=str, metavar='NAME',
        help='Spreadsheet name, URL, or index number')
    parser.add_argument('-y', '--yes', action='store_true',
        help='Skip confirmation prompts (auto-confirm)')
    parser.add_argument('-v', '--verbose', action='store_true',
        help='Enable verbose output in non-interactive mode (shows all messages)')
    
    return parser.parse_args()

def debug_print(message: str) -> None:
    """Print debug messages when DEBUGGING is enabled."""
    if config.DEBUGGING:
        print(f'! DEBUG {message}')

def verbose_print(message: str) -> None:
    """Print informational messages when VERBOSE is enabled.
    
    In interactive mode, VERBOSE is always True.
    In CLI mode, VERBOSE is False unless -v flag is used.
    """
    if config.VERBOSE:
        print(message)

def error_print(message: str, details: Optional[List[str]] = None) -> None:
    """Print error messages. Always displayed regardless of verbosity.
    
    Args:
        message: Primary error message
        details: Optional list of detail lines to print (indented)
    """
    print(f"! ERROR {message}")
    if details:
        for detail in details:
            print(f"  {detail}")

def warning_print(message: str, details: Optional[List[str]] = None) -> None:
    """Print warning messages. Always displayed regardless of verbosity.
    
    Args:
        message: Primary warning message
        details: Optional list of detail lines to print (indented)
    """
    print(f"! WARNING {message}")
    if details:
        for detail in details:
            print(f"  {detail}")

def info_print(message: str) -> None:
    """Print informational messages. Always displayed regardless of verbosity.
    
    Args:
        message: Informational message
    """
    print(message)

def normalize_study_name(raw_name: str) -> str:
    """Convert study name to a filesystem-safe format.
    Preserves unicode characters for international study names."""
    # Ensure we're working with a string
    name = str(raw_name)
    name = name.lower()
    name = name.replace('study ', 'study')
    name = name.replace(' ', '_')
    return name

def sanitize_filename(text: str) -> str:
    """Remove or replace characters that are unsafe for filenames.
    Preserves unicode characters to support international filenames."""
    # Ensure text is a string and handle unicode properly
    text = str(text)
    
    # Characters that need special replacement
    text = text.replace('\\', '-')
    text = text.replace('/', '-')
    text = text.replace('?', '_')
    # Characters to remove entirely (filesystem-unsafe characters)
    for char in ['\'', '\"', '.', '>', '<', '|', ':']:
        text = text.replace(char, '')
    return text

def add_duration(start_time: str) -> Union[str, int]:
    """Add default duration to a start timestamp.
    
    Adds DEFAULT_DURATION_SECONDS to the given start timestamp to create
    an end timestamp. Used when only a start time is provided.
    
    Args:
        start_time: Start timestamp in format MM:SS or HH:MM:SS
        
    Returns:
        The new timestamp string with duration added, or -1 if the timestamp
        format is invalid.
    """
    try:
        if len(start_time) <= 5:
            start_datetime = datetime.strptime(str(start_time), '%M:%S')
            new_time = start_datetime + timedelta(seconds=config.DEFAULT_DURATION_SECONDS)
            return new_time.strftime('%M:%S')
        else:
            start_datetime = datetime.strptime(start_time, '%H:%M:%S')
            new_time = start_datetime + timedelta(seconds=config.DEFAULT_DURATION_SECONDS)
            return new_time.strftime('%H:%M:%S')
    except ValueError:
        warning_print(f"Could not parse single timestamp '{start_time}' to add default duration.",
            [f"Expected format: MM:SS or HH:MM:SS (e.g., 12:34 or 1:23:45)",
             "This timestamp will be skipped."])
        return -1

def _parse_single_timestamp_token(token: str) -> Optional[Tuple[str, str]]:
    """Parse one token into a (start_time, end_time) pair, or None if invalid/skip.
    
    Handles: dash range (start-end), colon single time (add default duration), blank (None), else None.
    """
    if token == '':
        return None
    if '-' in token:
        dash_pos = token.find('-')
        if dash_pos > 0 and token[dash_pos - 1].isdigit():
            return (token[:dash_pos], token[dash_pos + 1:])
        return None
    if ':' in token:
        colon_pos = token.find(':')
        if colon_pos > 0 and token[colon_pos - 1].isdigit():
            end_time = add_duration(token)
            if isinstance(end_time, str):
                return (token, end_time)
        return None
    return None


def parse_timestamps(cell_value: str, cell_ref: Optional[str] = None) -> List[Tuple[str, str]]:
    """Parse timestamp pairs from a cell value string.
    
    Args:
        cell_value: The raw cell value containing timestamps
        cell_ref: Optional cell reference (e.g., 'B5') for error messages
    
    Returns:
        A list of (start_time, end_time) tuples.
    """
    if config.DEBUGGING:
        ic(cell_value, cell_ref)
    parsed_timestamps = []
    skipped_timestamps = []
    raw_times = cell_value.lower().replace('+', ' ').replace(';', ' ').replace(',', ' ').split()
    if config.DEBUGGING:
        ic(raw_times)
    debug_print(f'raw_times content after split is {raw_times}')
    debug_print(f'Timestamp list raw_times is {len(raw_times)} entries long')

    for i in range(len(raw_times)):
        debug_print(f'Cleaning timestamp {raw_times[i]}')
        raw_times[i] = raw_times[i].strip().rstrip(',').rstrip('-').replace('.', ':')
        pair = _parse_single_timestamp_token(raw_times[i])
        if pair is not None:
            if config.DEBUGGING and len(pair) == 2:
                ic(pair)
            parsed_timestamps.append(pair)
        elif raw_times[i]:
            skipped_timestamps.append(raw_times[i])
    
    # Report skipped timestamps if any
    if skipped_timestamps:
        if config.DEBUGGING:
            ic(skipped_timestamps)
        cell_info = f" in cell {cell_ref}" if cell_ref else ""
        details = []
        for ts in skipped_timestamps[:config.MAX_SKIPPED_TIMESTAMPS_TO_SHOW]:
            details.append(f"    '{ts}'")
        if len(skipped_timestamps) > config.MAX_SKIPPED_TIMESTAMPS_TO_SHOW:
            details.append(f"    ... and {len(skipped_timestamps) - config.MAX_SKIPPED_TIMESTAMPS_TO_SHOW} more")
        details.append("  Expected formats: MM:SS-MM:SS, HH:MM:SS-HH:MM:SS, or single timestamps like MM:SS")
        warning_print(f"Skipped {len(skipped_timestamps)} unparseable timestamp(s){cell_info}:", details)

    if config.DEBUGGING:
        ic(parsed_timestamps)
    return parsed_timestamps

def set_program_settings() -> bool:
    """Interactive function to change program settings.
    
    Allows user to modify settings like REENCODING, FILEFORMAT, and DEBUGGING.
    
    Returns:
        True if a setting was changed, False otherwise.
    """
    SETTINGS_OPTIONS = ['REENCODING', 'FILEFORMAT', 'DEBUGGING', 'MAX_FILESIZE_MB']

    info_print('\nWhich setting? Available:\n')
    info_print(', '.join(SETTINGS_OPTIONS))
    setting_to_change = input('\n>> ')

    info_print(f"* Current value for '{setting_to_change}' is '{getattr(config, setting_to_change)}'")

    new_value = input('\nWhich new value?\n>> ')

    info_print(f"* '{setting_to_change}' SET TO '{new_value}'")

    if setting_to_change != '':
        setattr(config, setting_to_change, new_value)
        return True
    return False

def get_current_time() -> str:
    """Get current time as formatted string.
    
    Returns:
        Current time in format 'YYYY-MM-DD HH:MM:SS'
    """
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')
