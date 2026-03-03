# -*- coding: utf-8 -*-
"""Utility functions for clipgen."""

import argparse
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Set, Tuple, TypedDict, Union

from icecream import ic

import config


# ---- Shared type definitions ----

class ClipRecord(TypedDict, total=False):
    """Clip record built by spreadsheet layer, enriched by prepare_clip.

    Always present after _make_clip_record: cell, desc, study, participant, category.
    Added by prepare_clip: times, cell_annotations, segment_annotations.
    Optionally set before prepare_clip: selected_segment_indexes, timestamp_baseline.
    """
    cell: Any  # gspread.Cell or ExcelSheetAdapter equivalent
    desc: str
    study: str
    participant: str
    category: str
    timestamp_baseline: str
    times: List[Tuple[str, str]]
    cell_annotations: List[str]
    segment_annotations: Dict[str, List[int]]
    selected_segment_indexes: List[int]


class ReelInput(TypedDict):
    """Parsed reel selector input from parse_reel_input."""
    batch: bool
    filter: bool
    timeline: bool
    lines: List[int]
    ranges: List[Tuple[int, int]]
    categories: List[str]
    cells: List[Tuple[str, int]]
    participants: List[str]


class BrowseRow(TypedDict, total=False):
    """Row data for browse mode display."""
    row_num: int
    category: str
    description: str
    timestamps: Dict[str, str]


# Rich library integration with graceful fallback
try:
    from rich.console import Console
    from rich.panel import Panel
    from rich.progress import Progress, BarColumn, TextColumn, TimeElapsedColumn, MofNCompleteColumn, SpinnerColumn
    from rich.table import Table
    from rich.text import Text
    from rich.theme import Theme
    RICH_AVAILABLE = True
except ImportError:
    RICH_AVAILABLE = False
    Progress = None  # type: ignore

# Custom theme for clipgen - only create if Rich is available
_CLIPGEN_THEME = Theme({
    "error": "bold red",
    "error.prefix": "bold red",
    "error.detail": "dim",
    "warning": "bold yellow",
    "warning.prefix": "bold yellow",
    "warning.detail": "dim",
    "success": "bold green",
    "success.prefix": "bold green",
    "info": "cyan",
    "verbose": "dim",
    "debug": "magenta",
}) if RICH_AVAILABLE else None

# Global console instance
console = Console(theme=_CLIPGEN_THEME, highlight=False) if RICH_AVAILABLE else None


def _use_rich() -> bool:
    """Check if Rich output should be used."""
    return RICH_AVAILABLE and console is not None and getattr(config, 'RICH_COLORS', True)


def _use_panels() -> bool:
    """Check if Rich panels should be used for errors/warnings/success."""
    return getattr(config, 'RICH_PANELS', True)


def _use_progress() -> bool:
    """Check if Rich progress bars should be used."""
    return RICH_AVAILABLE and console is not None and getattr(config, 'RICH_PROGRESS', True)


def create_progress_bar(description: str = "Processing"):
    """Create a Rich Progress instance configured for clipgen, or None if unavailable.

    Args:
        description: Default task description

    Returns:
        Configured Progress instance if Rich is available and enabled, else None
    """
    if not _use_progress():
        return None

    return Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(bar_width=40),
        MofNCompleteColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TimeElapsedColumn(),
        console=console,
        transient=False,  # Keep progress visible after completion
    )


def parse_arguments() -> argparse.Namespace:
    """Parse command-line arguments for non-interactive mode.
    
    Exactly one of the mode flags (-b, -l, -r, -c, -p, -f, -M, -R, -T) may be given;
    if none is given, the program runs in interactive mode. Optional flags (-s, -y,
    -v, --screen, --gif) may be combined with any mode.
    
    Returns:
        argparse.Namespace with attributes: batch, lines, range, cell,
        participant, filter, mixed, reel, timeline (mode flags/values),
        spreadsheet, yes, verbose, screen, gif.
    """
    parser = argparse.ArgumentParser(
        description='clipgen - Video clip generator from Google Sheets timestamps.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python clipgen.py                        Interactive mode (default)
  python clipgen.py -b                     Batch mode - generate all clips
  python clipgen.py -l 5                   Single line mode - line 5
  python clipgen.py -l 1+4+5               Multi-line mode - lines 1, 4, and 5
  python clipgen.py -l 1,4,5               Multi-line mode (comma separator)
  python clipgen.py -r 1-10                Range mode - lines 1 through 10
  python clipgen.py -c "P01.11"            Cell mode - single cell (participant P01, row 11)
  python clipgen.py -c "P01.11 + P03.11"   Cell mode - multiple cells
  python clipgen.py -p P01                 Participant mode - all clips for participant P01
  python clipgen.py -p "P01,P03"           Participant mode - all clips for P01 and P03
  python clipgen.py -f                     Filter mode - only key-marked clips/timestamps
  python clipgen.py -M "5, P01.11, 13-16"  Mixed mode - combine selectors for individual outputs
  python clipgen.py -b -s "Study Name"     Batch mode with specific spreadsheet
  python clipgen.py -l 5 -y                Line mode, skip confirmation prompts
  python clipgen.py -b -v                  Batch mode with verbose output
  python clipgen.py -R "11, 13-16, P01, \\"Observations\\""  Reel mode - one combined video
  python clipgen.py -T P01                 Timeline mode - chronological reel for participant P01
  python clipgen.py -b --screen            Batch mode screenshots (.png)
  python clipgen.py -l 5 --gif             Line mode GIF output (.gif)
 
Note: Non-interactive mode (using -b, -l, -r, -c, -p, -f, -M, -R, or -T) is silent by default,
      only showing errors and the final summary. Use -v for full output.
'''
    )

    # Mode arguments: only one of -b/-l/-r/-c/-p/-f/-M/-R/-T may be set at a time
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
    mode_group.add_argument('-f', '--filter', action='store_true',
        help='Filter mode: generate only key-marked clips/timestamps')
    mode_group.add_argument('-M', '--mixed', type=str, metavar='SELECTORS',
        help='Mixed mode: combine selectors (e.g. "5, P01.11, 13-16") for individual clips/screenshots/GIFs')
    mode_group.add_argument('-R', '--reel', type=str, metavar='SELECTORS',
        help='Reel mode: combine selectors (e.g. "11, 13-16, P01, \\"Observations\\"") into one video')
    mode_group.add_argument('-T', '--timeline', type=str, metavar='PARTICIPANT',
        help='Timeline mode: chronological reel for one participant (e.g., P01)')

    format_group = parser.add_mutually_exclusive_group()
    format_group.add_argument('--screen', action='store_true',
        help='Output screenshots (.png) instead of video clips')
    format_group.add_argument('--gif', action='store_true',
        help='Output animated GIFs instead of video clips')

    # Optional arguments (can be used with any mode)
    parser.add_argument('-s', '--spreadsheet', type=str, metavar='NAME',
        help='Spreadsheet name, URL, or index number')
    parser.add_argument('-y', '--yes', action='store_true',
        help='Skip confirmation prompts (auto-confirm)')
    parser.add_argument('-v', '--verbose', action='store_true',
        help='Enable verbose output in non-interactive mode (shows all messages)')
    parser.add_argument('--viewer', action='store_true',
        help='Generate a timeline HTML viewer file (clips_viewer.html) for this run')

    return parser.parse_args()

def debug_print(message: str) -> None:
    """Print debug messages when DEBUGGING is enabled."""
    if config.DEBUGGING:
        if _use_rich():
            console.print(f"[debug]! DEBUG[/debug] {message}")
        else:
            print(f'! DEBUG {message}')


def verbose_print(message: str) -> None:
    """Print informational messages when VERBOSE is enabled.

    In interactive mode, VERBOSE is always True.
    In CLI mode, VERBOSE is False unless -v flag is used.
    """
    if config.VERBOSE:
        if _use_rich() and console is not None:
            console.print(message, style="verbose")
        else:
            print(message)


def _styled_print(
    message: str,
    *,
    prefix: str = "",
    prefix_style: Optional[str] = None,
    message_style: Optional[str] = None,
    details: Optional[List[str]] = None,
    details_style: Optional[str] = None,
    panel_border_style: Optional[str] = None,
) -> None:
    if _use_rich() and console is not None:
        content = Text()
        if prefix:
            content.append(prefix, style=prefix_style)
        content.append(message, style=message_style)
        if details:
            for detail in details:
                content.append(f"\n  {detail}", style=details_style)

        if panel_border_style and _use_panels():
            console.print(Panel(content, border_style=panel_border_style, padding=(0, 1)))
        else:
            console.print(content)
        return

    print(f"{prefix}{message}")
    if details:
        for detail in details:
            print(f"  {detail}")


def error_print(message: str, details: Optional[List[str]] = None) -> None:
    """Print error messages. Always displayed regardless of verbosity.

    Args:
        message: Primary error message
        details: Optional list of detail lines to print (indented)
    """
    _styled_print(
        message,
        prefix="! ERROR ",
        prefix_style="error.prefix",
        details=details,
        details_style="error.detail",
        panel_border_style="red",
    )


def warning_print(message: str, details: Optional[List[str]] = None) -> None:
    """Print warning messages. Always displayed regardless of verbosity.

    Args:
        message: Primary warning message
        details: Optional list of detail lines to print (indented)
    """
    _styled_print(
        message,
        prefix="! WARNING ",
        prefix_style="warning.prefix",
        details=details,
        details_style="warning.detail",
        panel_border_style="yellow",
    )


def success_print(message: str) -> None:
    """Print success/completion messages. Always displayed regardless of verbosity.

    Args:
        message: Success message
    """
    _styled_print(
        message,
        prefix="✓ ",
        prefix_style="success.prefix",
        panel_border_style="green",
    )


def info_print(message: str) -> None:
    """Print informational messages. Always displayed regardless of verbosity.

    Args:
        message: Informational message
    """
    _styled_print(message, message_style="info")


def create_browse_table(
    rows_data: List[BrowseRow],
    participant_headers: List[str]
) -> Optional['Table']:
    """Create a Rich Table for browse mode display.

    Args:
        rows_data: List of dicts with keys: row_num, category, description, timestamps (dict)
        participant_headers: List of participant IDs for column headers

    Returns:
        Rich Table object if Rich is available, None otherwise
    """
    if not _use_rich():
        return None

    table = Table(
        show_header=True,
        header_style="bold cyan",
        border_style="dim",
        row_styles=["", "dim"],
        expand=False,
    )

    # Add fixed columns
    table.add_column("Row", justify="right", style="bold", width=4)
    table.add_column("Category", style="yellow", max_width=15, overflow="ellipsis")
    table.add_column("Description", max_width=config.BROWSE_DESCRIPTION_MAX_WIDTH, overflow="ellipsis")

    # Add participant columns
    for p_id in participant_headers:
        table.add_column(p_id, max_width=config.BROWSE_TIMESTAMP_MAX_WIDTH, overflow="fold")

    # Add data rows
    for row in rows_data:
        row_values = [
            str(row['row_num']),
            row['category'],
            row['description'],
        ]
        # Add timestamp values for each participant
        for participant_id in participant_headers:
            timestamp = row['timestamps'].get(participant_id, '-')
            row_values.append(timestamp if timestamp else '-')

        table.add_row(*row_values)

    return table


def format_browse_rows_plain(
    rows_data: List[BrowseRow],
    participant_headers: List[str]
) -> str:
    """Format browse rows as plain text (fallback when Rich unavailable).

    Args:
        rows_data: List of dicts with keys: row_num, category, description, timestamps (dict)
        participant_headers: List of participant IDs

    Returns:
        Formatted plain text string
    """
    lines = ['-' * 60]

    for row in rows_data:
        lines.append(f"Row {row['row_num']}")
        lines.append(f"  Category: {row['category']}")
        lines.append(f"  Description: {row['description']}")

        # Get participant timestamps
        participant_data = []
        for participant_id in participant_headers:
            timestamp = row['timestamps'].get(participant_id)
            if timestamp:
                participant_data.append(f"    {participant_id}: {timestamp}")

        if participant_data:
            lines.append('  Participants:')
            lines.extend(participant_data)
        else:
            lines.append('  Participants: (no timestamps)')

        lines.append('  ---')

    return '\n'.join(lines)


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


def normalize_participant_id(participant_value: str) -> str:
    """Strip known annotation tokens from a participant header value."""
    if not participant_value:
        return ''

    known_tokens = set(get_known_annotation_map().keys())
    cleaned_parts = []
    for part in participant_value.split():
        token = part.strip()
        if token and token.lower() not in known_tokens:
            cleaned_parts.append(token)
    return ' '.join(cleaned_parts).strip()


def index_to_letter(idx: int) -> str:
    """Convert 0-based index to letter label (0='A', 25='Z', 26='AA', etc.).

    Args:
        idx: 0-based index

    Returns:
        Letter label (A-Z, then AA, AB, etc.)
    """
    column_label = ''
    idx += 1  # Convert to 1-based for calculation
    while idx > 0:
        idx -= 1
        column_label = chr(ord('A') + (idx % 26)) + column_label
        idx //= 26
    return column_label


def letter_to_index(letter: str) -> int:
    """Convert letter label to 0-based index ('A'=0, 'Z'=25, 'AA'=26, etc.).

    Args:
        letter: Letter label (case-insensitive)

    Returns:
        0-based index, or -1 if invalid
    """
    letter = letter.upper().strip()
    if not letter or not letter.isalpha():
        return -1
    column_index = 0
    for char in letter:
        column_index = column_index * 26 + (ord(char) - ord('A') + 1)
    return column_index - 1


def _split_timestamp_tokens(cell_value: str) -> List[str]:
    """Split a cell value into normalized timestamp/annotation tokens."""
    return cell_value.lower().replace('+', ' ').replace(';', ' ').replace(',', ' ').split()


def _clean_timestamp_token(token: str) -> str:
    """Normalize one token before timestamp parsing."""
    return token.strip().rstrip(',').rstrip('-').replace('.', ':')


def get_ignored_timestamp_tokens() -> Set[str]:
    """Return configured ignored non-timestamp tokens in normalized form."""
    configured_tokens = getattr(config, 'IGNORED_TIMESTAMP_TOKENS', set())
    normalized_tokens: Set[str] = set()
    for token in configured_tokens:
        cleaned = _clean_timestamp_token(str(token).strip().lower())
        if cleaned:
            normalized_tokens.add(cleaned)
    return normalized_tokens


def has_non_ignored_timestamp_content(cell_value: str) -> bool:
    """Return True when cell content is more than ignored timestamp tokens.

    Cells containing only ignored tokens (e.g. "x") should not produce the
    generic "No valid timestamps found" warning.
    """
    ignored_tokens = get_ignored_timestamp_tokens()
    for raw_token in _split_timestamp_tokens(cell_value):
        token = _clean_timestamp_token(raw_token)
        if not token:
            continue
        if _parse_single_timestamp_token(token) is not None:
            return True
        if token not in ignored_tokens:
            return True
    return False


def get_known_annotation_map() -> Dict[str, str]:
    """Return configured annotation tokens mapped to normalized annotation IDs."""
    configured_map = getattr(config, 'ANNOTATION_KEYPHRASES', {'!key': 'key'})
    normalized_map: Dict[str, str] = {}
    for token, annotation_id in configured_map.items():
        normalized_map[str(token).strip().lower()] = str(annotation_id).strip().lower()
    return normalized_map


def parse_cell_annotations(
    cell_value: str,
    annotation_map: Optional[Dict[str, str]] = None
) -> Tuple[str, Dict[str, Set[int]], Set[str]]:
    """Extract inline annotation tokens and map them to parsed timestamp indexes.

    Semantics: an annotation token marks the preceding parseable timestamp token.
    The returned cleaned cell value has annotation tokens removed.
    """
    known_annotations = annotation_map or get_known_annotation_map()
    cleaned_tokens: List[str] = []
    segment_annotations: Dict[str, Set[int]] = {}
    cell_annotations: Set[str] = set()
    parsed_timestamp_count = 0

    for raw_token in _split_timestamp_tokens(cell_value):
        token = _clean_timestamp_token(raw_token)
        if not token:
            continue

        annotation_id = known_annotations.get(token)
        if annotation_id:
            cell_annotations.add(annotation_id)
            if parsed_timestamp_count > 0:
                segment_annotations.setdefault(annotation_id, set()).add(parsed_timestamp_count - 1)
            continue

        cleaned_tokens.append(token)
        if _parse_single_timestamp_token(token) is not None:
            parsed_timestamp_count += 1

    return (' '.join(cleaned_tokens), segment_annotations, cell_annotations)


def add_duration(start_time: str) -> Optional[str]:
    """Add default duration to a start timestamp.
    
    Adds DEFAULT_DURATION_SECONDS to the given start timestamp to create
    an end timestamp. Used when only a start time is provided.
    
    Args:
        start_time: Start timestamp in format MM:SS or HH:MM:SS
        
    Returns:
        The new timestamp string with duration added, or None if the timestamp
        format is invalid.
    """
    try:
        if len(start_time) <= config.MAX_MMSS_LENGTH:
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
        return None


def timestamp_to_seconds(ts_str: str) -> Optional[float]:
    """Convert MM:SS or HH:MM:SS timestamp string to seconds.

    Args:
        ts_str: Timestamp string in MM:SS or HH:MM:SS format

    Returns:
        Total seconds as float, or None if the timestamp cannot be parsed.
    """
    ts = (ts_str or '').strip()
    if not ts:
        return None

    formats = ['%M:%S', '%H:%M:%S']
    for fmt in formats:
        try:
            parsed = datetime.strptime(ts, fmt)
            return float(parsed.hour * config.SECONDS_PER_HOUR + parsed.minute * config.SECONDS_PER_MINUTE + parsed.second)
        except ValueError:
            continue
    return None

def _parse_single_timestamp_token(token: str) -> Optional[Tuple[str, str]]:
    """Parse one token into a (start_time, end_time) pair, or None if invalid/skip.

    Handles: dash range (start-end), single timestamp with colon (add default
    duration), blank token (skip), or unrecognized format (skip).

    Args:
        token: A single timestamp token, e.g. "1:23-1:45", "2:30", or "".

    Returns:
        (start_time, end_time) tuple if parseable, else None (caller skips).

    Examples:
        "1:23-1:45" -> ("1:23", "1:45"); "2:30" -> ("2:30", "2:45") with default duration.
    """
    if token == '':
        return None
    # Dash range: "start-end". Require a digit before the dash so we don't
    # treat a leading dash (e.g. "-5") or non-time dash as a range.
    if '-' in token:
        dash_pos = token.find('-')
        if dash_pos > 0 and token[dash_pos - 1].isdigit():
            return (token[:dash_pos], token[dash_pos + 1:])
        return None
    # Single timestamp with colon: use as start and add default duration for end.
    # Require a digit before the first colon so we only match time-like strings.
    if ':' in token:
        colon_pos = token.find(':')
        if colon_pos > 0 and token[colon_pos - 1].isdigit():
            end_time = add_duration(token)
            if end_time is not None:
                return (token, end_time)
        return None
    return None


def parse_timestamps(cell_value: str, cell_ref: Optional[str] = None) -> List[Tuple[str, str]]:
    """Parse timestamp pairs from a cell value string.

    Pipeline: (1) Normalize delimiters to spaces and split into tokens,
    (2) Clean each token and parse into (start, end) pairs,
    (3) Report any unparseable tokens as warnings.

    Supported formats: "MM:SS-MM:SS", "HH:MM:SS-HH:MM:SS", or a single time
    "MM:SS"/"HH:MM:SS" (end time is start + default duration). Delimiters
    between multiple pairs: space, comma, semicolon, or plus.

    Args:
        cell_value: The raw cell value containing timestamps
        cell_ref: Optional cell reference (e.g., 'B5') for error messages

    Returns:
        A list of (start_time, end_time) tuples. Invalid tokens are skipped
        and reported via warning_print.
    """
    if config.DEBUGGING:
        ic(cell_value, cell_ref)
    parsed_timestamps = []
    skipped_timestamps = []
    ignored_tokens = get_ignored_timestamp_tokens()
    # Unify delimiters (+, ;, ,) to spaces so split() yields one token per time or range
    raw_times = _split_timestamp_tokens(cell_value)
    if config.DEBUGGING:
        ic(raw_times)
    debug_print(f'raw_times content after split is {raw_times}')
    debug_print(f'Timestamp list raw_times is {len(raw_times)} entries long')

    # Clean each token (strip, normalize trailing punctuation, use colon for decimals) and parse
    raw_times = [_clean_timestamp_token(t) for t in raw_times]
    for token in raw_times:
        debug_print(f'Cleaning timestamp {token}')
        pair = _parse_single_timestamp_token(token)
        if pair is not None:
            if config.DEBUGGING and len(pair) == 2:
                ic(pair)
            parsed_timestamps.append(pair)
        elif token and token not in ignored_tokens:
            skipped_timestamps.append(token)

    # Report skipped timestamps: list up to MAX_SKIPPED_TIMESTAMPS_TO_SHOW, then "... and N more"
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


def _clock_to_seconds(ts: str) -> Optional[int]:
    """Parse a clock-style timestamp into total seconds.

    Accepts HH:MM:SS, HH:MM, or MM:SS. Returns None if parsing fails.
    """
    value = (ts or '').strip()
    if not value:
        return None

    # Choose format based on number of components to avoid treating "22:00"
    # as 22 minutes instead of 22 hours.
    parts = value.split(':')
    if len(parts) == 3:
        try:
            parsed = datetime.strptime(value, '%H:%M:%S')
            return parsed.hour * config.SECONDS_PER_HOUR + parsed.minute * config.SECONDS_PER_MINUTE + parsed.second
        except ValueError:
            return None
    if len(parts) == 2:
        # Prefer HH:MM for clock-style values like "22:00"; fall back to MM:SS.
        for fmt in ('%H:%M', '%M:%S'):
            try:
                parsed = datetime.strptime(value, fmt)
                return parsed.hour * config.SECONDS_PER_HOUR + parsed.minute * config.SECONDS_PER_MINUTE + parsed.second
            except ValueError:
                continue
        return None
    return None


def _seconds_to_timestamp(total_seconds: int) -> str:
    """Format a non-negative number of seconds as H:MM:SS or M:SS."""
    if total_seconds < 0:
        total_seconds = 0
    hours, rem = divmod(total_seconds, config.SECONDS_PER_HOUR)
    minutes, seconds = divmod(rem, config.SECONDS_PER_MINUTE)
    if hours > 0:
        return f"{hours:d}:{minutes:02d}:{seconds:02d}"
    return f"{minutes:d}:{seconds:02d}"


def convert_clock_pairs_to_relative(
    pairs: List[Tuple[str, str]],
    baseline: str,
    cell_ref: Optional[str] = None,
) -> List[Tuple[str, str]]:
    """Convert absolute clock (start, end) pairs to relative offsets using a baseline.

    Returns a new list of (start, end) pairs in relative time. Invalid or
    non-positive-length segments are skipped with warnings.
    """
    baseline_seconds = _clock_to_seconds(baseline)
    if baseline_seconds is None:
        cell_info = f" in cell {cell_ref}" if cell_ref else ""
        warning_print(
            f"Ignoring invalid baseline timestamp '{baseline}'{cell_info}.",
            ["Baseline must use clock format like HH:MM:SS or MM:SS.",
             "Timestamps in this column will be treated as relative."],
        )
        return pairs

    result: List[Tuple[str, str]] = []
    skipped: List[str] = []
    for start_str, end_str in pairs:
        start_s = _clock_to_seconds(start_str)
        end_s = _clock_to_seconds(end_str)
        if start_s is None or end_s is None:
            skipped.append(f"{start_str}-{end_str}")
            continue
        start_rel = start_s - baseline_seconds
        end_rel = end_s - baseline_seconds
        if start_rel < 0 or end_rel <= 0 or end_rel <= start_rel:
            skipped.append(f"{start_str}-{end_str}")
            continue
        result.append(
            (_seconds_to_timestamp(start_rel), _seconds_to_timestamp(end_rel))
        )

    if skipped:
        cell_info = f" in cell {cell_ref}" if cell_ref else ""
        details = [f"    '{s}'" for s in skipped]
        details.append("  These segments were before the baseline, invalid, or zero/negative length.")
        warning_print(
            f"Skipped {len(skipped)} clock-based timestamp segment(s){cell_info} when converting to relative:",
            details,
        )

    return result

class QuitProgram(Exception):
    """Signal that the user requested to quit the program from an interactive prompt."""


class TopToSpreadsheet(Exception):
    """Signal that the user requested to return to spreadsheet selection."""


class BackToModeSelection(Exception):
    """Signal that the user requested to return to the main mode selection prompt."""


def read_user_input(prompt: str) -> str:
    """Read user input and handle global control keywords.

    Recognizes the following keywords when they appear as the first token:
    - 'quit' / 'exit' -> quit program
    - 'top'           -> return to spreadsheet selection
    - 'back'          -> return to main mode selection
    """
    raw = input(prompt)
    value = raw.strip()
    if not value:
        return value

    first_token = value.split()[0].lower()
    if first_token in ('quit', 'exit'):
        info_print('Exiting clipgen.')
        raise QuitProgram()
    if first_token == 'top':
        info_print('Returning to spreadsheet selection.')
        raise TopToSpreadsheet()
    if first_token == 'back':
        info_print('Returning to mode selection.')
        raise BackToModeSelection()

    return value


def set_program_settings() -> bool:
    """Interactive function to change program settings.
    
    Allows user to modify settings like REENCODING, FILEFORMAT, and DEBUGGING.
    
    Returns:
        True if a setting was changed, False otherwise.
    """
    SETTINGS_OPTIONS = ['REENCODING', 'FILEFORMAT', 'DEBUGGING', 'MAX_FILESIZE_MB']

    info_print('\nWhich setting? Available:\n')
    info_print(', '.join(SETTINGS_OPTIONS))
    setting_to_change = read_user_input('\n>> ')

    info_print(f"* Current value for '{setting_to_change}' is '{getattr(config, setting_to_change)}'")

    new_value = read_user_input('\nWhich new value?\n>> ')

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
