# -*- coding: utf-8 -*-
"""File and filename operations for clipgen."""

import os
import re
from typing import Any, Dict, List

import gspread
from icecream import ic

import config
import utils


def pad_number_two_digits(number: str) -> str:
    """Pad a numeric string to two digits (e.g. '5' -> '05').
    
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
    # Keep dividing by 1024 until size is under 1024 or we reach TB (index 4)
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
                # First collision: insert step number before file extension
                # "file.mp4" -> "file-1.mp4"
                suffix_pos = filename.find(config.FILEFORMAT)
                filename = filename[0:suffix_pos] + '-' + str(step) + config.FILEFORMAT
            else:
                # Subsequent collisions: replace existing step number
                # "file-1.mp4" -> "file-2.mp4"
                dash_pos = filename.rfind('-')
                filename = filename[0:dash_pos] + '-' + str(step) + config.FILEFORMAT
            step += 1
        else:
            # Found a unique name; truncate if needed and return
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
            # Reserve space for: dash + step number + extension (e.g., "-2.mp4")
            max_base_len = config.MAX_FILENAME_LENGTH - (1 + len(str(step)) + len(config.FILEFORMAT))
            filename = filename[0:max_base_len] + '-' + str(step) + config.FILEFORMAT
        else:
            # Reserve space for extension only (e.g., ".mp4")
            max_base_len = config.MAX_FILENAME_LENGTH - len(config.FILEFORMAT)
            filename = filename[0:max_base_len] + config.FILEFORMAT
    return filename

def prepare_clip(clip: Dict[str, Any]) -> Dict[str, Any]:
    """Parse timestamps and sanitize description/category for filename use.
    
    Mutates the input clip dict: adds 'times' (list of (start, end) timestamp pairs)
    and overwrites 'desc' and 'category' with sanitized values.
    
    Expected input keys: 'cell', 'desc', 'category', 'study', 'participant'.
    
    Args:
        clip: Clip record dict containing 'cell', 'desc', 'category', 'study', 'participant'
        
    Returns:
        The same dict with 'times' added and sanitized 'desc' and 'category'
    """
    if config.DEBUGGING:
        ic(clip)
    utils.debug_print(f"prepare_clip() received clip with cell contents {clip['cell'].value}")
    utils.debug_print('Will attempt to split the cell contents')
    
    # Get cell reference for error messages
    cell_ref = gspread.utils.rowcol_to_a1(clip['cell'].row, clip['cell'].col)
    
    # Parse timestamps from cell value
    clip['times'] = utils.parse_timestamps(clip['cell'].value, cell_ref=cell_ref)
    if config.DEBUGGING:
        ic(clip['times'])
    
    # Warn if no valid timestamps were parsed
    if not clip['times']:
        utils.warning_print(f"No valid timestamps found in cell {cell_ref}",
            [f"Cell contents: '{clip['cell'].value}'",
             f"Participant: {clip['participant']}, Description: {clip['desc'][:50]}..."])

    # Clean description: remove bracketed prefix like "[TAG] actual description"
    # and sanitize for use in filename
    bracket_pos = clip['desc'].rfind(']')
    if bracket_pos >= 0:
        # Strip everything up to and including the last ']'
        desc = clip['desc'][bracket_pos+1:].strip()
    else:
        # No bracket found; use description as-is
        desc = clip['desc'].strip()
    clip['desc'] = utils.sanitize_filename(desc)
    if config.DEBUGGING:
        ic(clip['desc'])
    
    # Sanitize category (handle None/empty)
    if clip['category']:
        clip['category'] = utils.sanitize_filename(clip['category'])
    else:
        clip['category'] = 'uncategorized'
    if config.DEBUGGING:
        ic(clip['category'])
        ic(clip)
    return clip


def is_source_video(filename: str) -> bool:
    """Check if filename matches source video pattern (study_P01.mp4, study_G02.mp4).
    
    Source videos follow the naming convention {study}_{participant}.mp4 where
    participant starts with P or G followed by digits.
    
    Args:
        filename: Filename to check
        
    Returns:
        True if filename matches source video pattern, False otherwise
    """
    return bool(re.search(r'_[PG]\d+\.mp4$', filename, re.IGNORECASE))


def discover_clips() -> List[str]:
    """Find generated clips in the current working directory.
    
    Scans for .mp4 files and excludes source videos (those matching the
    pattern study_P01.mp4, study_G02.mp4, etc.).
    
    Returns:
        Sorted list of clip filenames (alphabetically)
    """
    clips = []
    for f in os.listdir('.'):
        if f.endswith(config.FILEFORMAT) and not is_source_video(f):
            clips.append(f)
    return sorted(clips)
