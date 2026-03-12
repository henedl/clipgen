# -*- coding: utf-8 -*-
"""File and filename operations for clipgen."""

import re
from pathlib import Path
from typing import List, Optional

import gspread
from icecream import ic

import config
import utils
from utils import ClipRecord

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
        Unique filename path as a string that doesn't exist in the filesystem.
    """
    base_path = utils.resolve_output_path(filename)
    file_extension = file_format or config.FILEFORMAT
    suffix_counter = 1
    current_name = str(base_path)
    while True:
        if Path(current_name).is_file():
            if suffix_counter < 2:
                suffix_pos = current_name.rfind(file_extension)
                if suffix_pos >= 0:
                    current_name = current_name[0:suffix_pos] + '-' + str(suffix_counter) + file_extension
                else:
                    current_name = current_name + '-' + str(suffix_counter)
            else:
                dash_pos = current_name.rfind('-')
                if dash_pos >= 0:
                    base = current_name[0:dash_pos]
                else:
                    suffix_pos = current_name.rfind(file_extension)
                    base = current_name[0:suffix_pos] if suffix_pos >= 0 else current_name
                current_name = base + '-' + str(suffix_counter) + file_extension
            suffix_counter += 1
        else:
            current_name = truncate_filename(current_name, suffix_counter, file_extension)
            break
    return current_name

def truncate_filename(filename: str, step: int = 1, file_format: Optional[str] = None) -> str:
    """Truncate filenames that exceed maximum length.
    
    Truncates to MAX_FILENAME_LENGTH (255 chars on Windows), preserving
    file extension and step number if present.
    
    Args:
        filename: Filename to truncate
        step: Step number for unique filename generation (default: 1)
        file_format: File extension to preserve (defaults to config.FILEFORMAT)
        
    Returns:
        Truncated filename (path string) that fits within max length.
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


def get_source_video_filename(study: str, participant: str, override: Optional[str] = None) -> str:
    """Resolve the expected source video filename for a clip.

    Args:
        study: Normalized study name
        participant: Normalized participant ID
        override: Optional filename override from the spreadsheet

    Returns:
        Filename to use when looking for the source video.
    """
    if override is not None:
        override = override.strip()
    if override:
        # If override includes an extension, respect it as-is.
        suffix = Path(override).suffix
        if suffix:
            return override
        # No extension present: append configured default file format.
        return override + config.FILEFORMAT
    return f"{study}_{participant}{config.FILEFORMAT}"

def prepare_clip(clip: ClipRecord) -> ClipRecord:
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
    timestamp_baseline = clip.get('timestamp_baseline')
    if timestamp_baseline:
        clip['times'] = utils.convert_clock_pairs_to_relative(clip['times'], timestamp_baseline, cell_ref=cell_ref)
    selected_segment_indexes = clip.get('selected_segment_indexes')
    if selected_segment_indexes is not None:
        selected_set = set(selected_segment_indexes)
        clip['times'] = [pair for index, pair in enumerate(clip['times']) if index in selected_set]
    if config.DEBUGGING:
        ic(clip['times'])
    
    # Warn if no valid timestamps were parsed, except cells with only ignored tokens (e.g. "x").
    if not clip['times'] and utils.has_non_ignored_timestamp_content(cleaned_cell_value):
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
    return bool(re.search(config.SOURCE_VIDEO_PATTERN, filename, re.IGNORECASE))


def discover_clips() -> List[str]:
    """Find generated clips in the effective output directory.
    
    Scans for .mp4 files and excludes source videos (those matching the
    pattern study_P01.mp4, study_G02.mp4, etc.).
    
    Returns:
        Sorted list of clip filenames (relative to the output directory)
    """
    base_dir = utils.get_effective_output_dir()
    return sorted(
        p.name for p in base_dir.iterdir()
        if p.name.endswith(config.FILEFORMAT) and not is_source_video(p.name)
    )
