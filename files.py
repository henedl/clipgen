# -*- coding: utf-8 -*-
"""File and filename operations for clipgen."""

import os
from typing import Any, Dict

import gspread
from icecream import ic

import config
import utils


def double_digits(number: str) -> str:
    """Convert a number string to double-digit format.
    
    Args:
        number: String representation of a number
        
    Returns:
        String with leading zero if number < 10, otherwise returns original string
    """
    try:
        if int(number) < 10:
            return '0' + number
        return number
    except TypeError:
        return number

def format_filesize(size_bytes: float, precision: int = 2) -> str:
    """Format byte size as human-readable string.
    
    Args:
        size_bytes: Size in bytes
        precision: Number of decimal places (default: 2)
        
    Returns:
        Formatted string with appropriate unit (B, KB, MB, GB, TB)
    """
    suffixes = ['B', 'KB', 'MB', 'GB', 'TB']
    suffix_index = 0
    while size_bytes > 1024 and suffix_index < 4:
        suffix_index += 1
        size_bytes = size_bytes / 1024
    return f'{size_bytes:.{precision}f}{suffixes[suffix_index]}'

def get_unique_filename(filename: str) -> str:
    """Generate a unique filename by appending an incremented number.
    
    If a file with the given name already exists, appends '-1', '-2', etc.
    until a unique filename is found. Also truncates if filename exceeds max length.
    
    Args:
        filename: Original filename
        
    Returns:
        Unique filename that doesn't exist in the filesystem
    """
    step = 1
    while True:
        if os.path.isfile(filename):
            if step < 2:
                suffix_pos = filename.find(config.FILEFORMAT)
                filename = filename[0:suffix_pos] + '-' + str(step) + config.FILEFORMAT
            else:
                dash_pos = filename.rfind('-')
                filename = filename[0:dash_pos] + '-' + str(step) + config.FILEFORMAT
            step += 1
        else:
            filename = truncate_filename(filename, step)
            break
    return filename

def truncate_filename(filename: str, step: int = 1) -> str:
    """Truncate filenames that exceed maximum length.
    
    Truncates to MAX_FILENAME_LENGTH (255 chars on Windows), preserving
    file extension and step number if present.
    
    Args:
        filename: Filename to truncate
        step: Step number for unique filename generation (default: 1)
        
    Returns:
        Truncated filename that fits within max length
    """
    if len(filename) > config.MAX_FILENAME_LENGTH:
        if step > 1:
            utils.debug_print(f'Filename was longer than {config.MAX_FILENAME_LENGTH} chars ({filename}, length {len(filename)})')
            filename = filename[0:config.MAX_FILENAME_LENGTH-(1+len(str(step))+len(config.FILEFORMAT))] + '-' + str(step) + config.FILEFORMAT
        else:
            filename = filename[0:config.MAX_FILENAME_LENGTH-(len(config.FILEFORMAT))] + config.FILEFORMAT
    return filename

def clean_issue(issue: Dict[str, Any]) -> Dict[str, Any]:
    """Parse timestamps and sanitize description/category for filename use.
    
    Processes a clip issue dictionary by:
    - Parsing timestamps from the cell value
    - Cleaning and sanitizing the description
    - Sanitizing the category (defaults to 'uncategorized' if empty)
    
    Args:
        issue: Dictionary containing 'cell', 'desc', 'category', 'study', 'participant'
        
    Returns:
        Modified issue dictionary with 'times' added and sanitized fields
    """
    ic(issue)
    utils.debug_print(f"clean_issue() received issue with cell contents {issue['cell'].value}")
    utils.debug_print('Will attempt to split the cell contents')
    
    # Get cell reference for error messages
    cell_ref = gspread.utils.rowcol_to_a1(issue['cell'].row, issue['cell'].col)
    
    # Parse timestamps from cell value
    issue['times'] = utils.parse_timestamps(issue['cell'].value, cell_ref=cell_ref)
    ic(issue['times'])
    
    # Warn if no valid timestamps were parsed
    if not issue['times']:
        utils.warning_print(f"No valid timestamps found in cell {cell_ref}",
            [f"Cell contents: '{issue['cell'].value}'",
             f"Participant: {issue['participant']}, Description: {issue['desc'][:50]}..."])

    # Clean description: remove bracketed prefix and sanitize
    # Handle case where description doesn't contain ']'
    bracket_pos = issue['desc'].rfind(']')
    if bracket_pos >= 0:
        desc = issue['desc'][bracket_pos+1:].strip()
    else:
        desc = issue['desc'].strip()
    issue['desc'] = utils.sanitize_filename(desc)
    ic(issue['desc'])
    
    # Sanitize category (handle None/empty)
    if issue['category']:
        issue['category'] = utils.sanitize_filename(issue['category'])
    else:
        issue['category'] = 'uncategorized'
    ic(issue['category'])

    ic(issue)
    return issue
