# -*- coding: utf-8 -*-
"""File and filename operations for clipgen."""

import os
import re
from typing import Any, Dict, List, Optional

import gspread
from icecream import ic

import config
import utils

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

def get_unique_filename(filename: str, file_format: Optional[str] = None) -> str:
    """Generate a unique filename by appending an incremented number.
    
    If a file with the given name already exists, appends '-1', '-2', etc.
    until a unique filename is found. Also truncates if filename exceeds max length.
    
    Args:
        filename: Original filename
        file_format: File extension to preserve (defaults to config.FILEFORMAT)
        
    Returns:
        Unique filename that doesn't exist in the filesystem
    """
    file_extension = file_format or config.FILEFORMAT
    suffix_counter = 1
    while True:
        if os.path.isfile(filename):
            if suffix_counter < 2:
                # First collision: insert step number before file extension
                # "file.mp4" -> "file-1.mp4"
                suffix_pos = filename.rfind(file_extension)
                if suffix_pos >= 0:
                    filename = filename[0:suffix_pos] + '-' + str(suffix_counter) + file_extension
                else:
                    filename = filename + '-' + str(suffix_counter)
            else:
                # Subsequent collisions: replace existing step number
                # "file-1.mp4" -> "file-2.mp4"
                dash_pos = filename.rfind('-')
                if dash_pos >= 0:
                    base = filename[0:dash_pos]
                else:
                    suffix_pos = filename.rfind(file_extension)
                    base = filename[0:suffix_pos] if suffix_pos >= 0 else filename
                filename = base + '-' + str(suffix_counter) + file_extension
            suffix_counter += 1
        else:
            filename = truncate_filename(filename, suffix_counter, file_extension)
            break
    return filename

def truncate_filename(filename: str, step: int = 1, file_format: Optional[str] = None) -> str:
    """Truncate filenames that exceed maximum length.
    
    Truncates to MAX_FILENAME_LENGTH (255 chars on Windows), preserving
    file extension and step number if present.
    
    Args:
        filename: Filename to truncate
        step: Step number for unique filename generation (default: 1)
        file_format: File extension to preserve (defaults to config.FILEFORMAT)
        
    Returns:
        Truncated filename that fits within max length
    """
    file_extension = file_format or config.FILEFORMAT
    if len(filename) > config.MAX_FILENAME_LENGTH:
        if step > 1:
            utils.debug_print(f'Filename was longer than {config.MAX_FILENAME_LENGTH} chars ({filename}, length {len(filename)})')
            # Reserve space for: dash + step number + extension (e.g., "-2.mp4")
            max_base_len = config.MAX_FILENAME_LENGTH - (1 + len(str(step)) + len(file_extension))
            filename = filename[0:max_base_len] + '-' + str(step) + file_extension
        else:
            # Reserve space for extension only (e.g., ".mp4")
            max_base_len = config.MAX_FILENAME_LENGTH - len(file_extension)
            filename = filename[0:max_base_len] + file_extension
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
    
    # Parse inline annotations (e.g. !key), then parse timestamps from cleaned value.
    cleaned_cell_value, segment_annotations, cell_annotations = utils.parse_cell_annotations(clip['cell'].value)
    clip['cell_annotations'] = sorted(cell_annotations)
    clip['segment_annotations'] = {key: sorted(indexes) for key, indexes in segment_annotations.items()}
    clip['times'] = utils.parse_timestamps(cleaned_cell_value, cell_ref=cell_ref)
    selected_segment_indexes = clip.get('selected_segment_indexes')
    if selected_segment_indexes is not None:
        selected_set = set(selected_segment_indexes)
        clip['times'] = [pair for index, pair in enumerate(clip['times']) if index in selected_set]
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
    for filename in os.listdir('.'):
        if filename.endswith(config.FILEFORMAT) and not is_source_video(filename):
            clips.append(filename)
    return sorted(clips)
