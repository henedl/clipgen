# -*- coding: utf-8 -*-
"""File and filename operations for clipgen."""

import re
import unicodedata
from pathlib import Path

import gspread
from icecream import ic

import config
import utils
from utils import ClipRecord


def _safe_truncate(text: str, max_chars: int) -> str:
    """Truncate to ``max_chars`` code points, then drop trailing combining marks.

    Plain ``text[:n]`` can split a grapheme cluster (e.g. emoji + skin-tone
    modifier, or letter + combining accent), leaving an orphan combining
    character. Drop those at the truncation boundary so the visible glyph
    survives intact.
    """
    if max_chars <= 0:
        return ""
    truncated = text[:max_chars]
    # Trim trailing combining marks / joiners that would render orphaned.
    while truncated and (
        unicodedata.category(truncated[-1]) in ("Mn", "Mc", "Me")
        or truncated[-1] in ("‍", "️")
    ):
        truncated = truncated[:-1]
    return truncated


def get_unique_filename(filename: str, file_format: str | None = None) -> str:
    """Generate a unique filename by appending an incremented number.

    If a file with the given name already exists, appends '-1', '-2', etc.
    until a unique filename is found. Also truncates if filename exceeds max length.

    Args:
        filename: Original filename
        file_format: File extension to preserve (defaults to config.FILEFORMAT)

    Returns:
        Unique filename path as a string that doesn't exist in the filesystem.
    """
    file_extension = file_format or config.FILEFORMAT
    resolved = Path(utils.resolve_output_path(filename))
    directory = resolved.parent
    name = resolved.name
    # Strip extension to get base name
    if name.endswith(file_extension):
        base = name[: -len(file_extension)]
    else:
        base = name
    # Truncate base if needed (reserve space for extension)
    max_base = config.MAX_FILENAME_LENGTH - len(file_extension)
    base = _safe_truncate(base, max_base)
    candidate = directory / (base + file_extension)
    counter = 1
    while candidate.is_file():
        suffix = f"-{counter}"
        truncated_base = _safe_truncate(base, max_base - len(suffix))
        candidate = directory / (truncated_base + suffix + file_extension)
        counter += 1
    return str(candidate)


def get_source_video_filename(
    study: str, participant: str, override: str | None = None
) -> str:
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

    # Pre-parsed fast path: callers (e.g. --ss-clips, --transcript-clips) build
    # synthetic clips with timestamps already resolved. Skip the cell-based parse
    # but still sanitize desc/category for use in filenames.
    if clip.get("times"):
        clip["cell_annotations"] = list(clip.get("cell_annotations") or [])
        clip["segment_annotations"] = dict(clip.get("segment_annotations") or {})
        raw_desc = clip.get("desc", "")
        bracket_pos = raw_desc.rfind("]")
        cleaned_desc = (
            raw_desc[bracket_pos + 1 :].strip()
            if bracket_pos >= 0
            else raw_desc.strip()
        )
        clip["desc"] = utils.sanitize_filename(cleaned_desc)
        clip["category"] = (
            utils.sanitize_filename(clip["category"])
            if clip.get("category")
            else "uncategorized"
        )
        return clip

    utils.debug_print(
        f"prepare_clip() received clip with cell contents {clip['cell'].value}"
    )
    utils.debug_print("Will attempt to split the cell contents")

    # Get cell reference for error messages
    cell_ref = gspread.utils.rowcol_to_a1(clip["cell"].row, clip["cell"].col)

    # Parse inline annotations (e.g. !key), then parse timestamps from cleaned value.
    cleaned_cell_value, segment_annotations, cell_annotations = (
        utils.parse_cell_annotations(clip["cell"].value)
    )
    clip["cell_annotations"] = sorted(cell_annotations)
    clip["segment_annotations"] = {
        key: sorted(indexes) for key, indexes in segment_annotations.items()
    }
    clip["times"] = utils.parse_timestamps(cleaned_cell_value, cell_ref=cell_ref)
    timestamp_baseline = clip.get("timestamp_baseline")
    if timestamp_baseline:
        clip["times"] = utils.convert_clock_pairs_to_relative(
            clip["times"], timestamp_baseline, cell_ref=cell_ref
        )
    selected_segment_indexes = clip.get("selected_segment_indexes")
    if selected_segment_indexes is not None:
        selected_set = set(selected_segment_indexes)
        clip["times"] = [
            pair for index, pair in enumerate(clip["times"]) if index in selected_set
        ]
    if config.DEBUGGING:
        ic(clip["times"])

    # Warn if no valid timestamps were parsed, except cells with only ignored tokens (e.g. "x").
    if not clip["times"] and utils.has_non_ignored_timestamp_content(
        cleaned_cell_value
    ):
        # Only show this detailed per-cell warning at verbose verbosity.
        if getattr(config, "VERBOSITY", config.STANDARD) >= config.VERBOSE:
            utils.warning_print(
                f"No valid timestamps found in cell {cell_ref}",
                [
                    f"Cell contents: '{clip['cell'].value}'",
                    f"Participant: {clip['participant']}, Description: {clip['desc'][:50]}...",
                ],
            )

    # Clean description: remove bracketed prefix and sanitize for use in filename
    raw_desc = clip["desc"]
    bracket_pos = raw_desc.rfind("]")
    cleaned_desc = (
        raw_desc[bracket_pos + 1 :].strip() if bracket_pos >= 0 else raw_desc.strip()
    )
    clip["desc"] = utils.sanitize_filename(cleaned_desc)
    if config.DEBUGGING:
        ic(clip["desc"])

    # Sanitize category (handle None/empty)
    if clip["category"]:
        clip["category"] = utils.sanitize_filename(clip["category"])
    else:
        clip["category"] = "uncategorized"
    if config.DEBUGGING:
        ic(clip["category"])
        ic(clip)
    return clip


def discover_clips() -> list[str]:
    """Find generated clips in the effective output directory.

    Scans for .mp4 files and excludes source videos (those matching the
    pattern study_P01.mp4, study_G02.mp4, etc.).

    Returns:
        Sorted list of clip filenames (relative to the output directory)
    """
    base_dir = utils.get_effective_output_dir()
    return sorted(
        p.name
        for p in base_dir.iterdir()
        if p.name.endswith(config.FILEFORMAT)
        and not re.search(config.SOURCE_VIDEO_PATTERN, p.name, re.IGNORECASE)
    )
