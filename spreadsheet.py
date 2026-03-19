# -*- coding: utf-8 -*-
"""Spreadsheet data processing for clipgen.

Expected spreadsheet structure:
- Row 0 (first row): Study name in cell A1; may be empty (falls back to spreadsheet title).
- Header cells: Contains required columns ID, Observation, Category (exact names from config).
- Participant columns: Immediately follow the ID column; headers start with P or G (e.g. P01, P02, G01).
  Each participant column holds timestamp values; non-empty cells become clip candidates.
- Category column: Labels each observation row; used for category and reel selection.
- Observation column: Human-readable description for each row; used in clip metadata.

Coordinate system:
- gspread uses 1-based row/col (e.g. Cell(row=1, col=1) is A1).
- sheet.get_all_values() returns a list of lists with 0-based indices: sheet_data[row][col].
- Conversions: sheet row = row_idx + 1, sheet col = col_idx + 1; header cell .row/.col are 1-based.

Clip record (returned by generation functions):
- Dict with keys: cell (gspread Cell), desc (observation text), study (normalized name),
  participant (header value for that column), category (row category). The 'times' key
  is added later by prepare_clip when resolving timestamps to actual time ranges.
"""

import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Set, Tuple

import gspread
from icecream import ic

import config
import utils
from utils import ClipRecord, ReelInput


# ---- Sheet context dataclass ----


@dataclass
class SheetContext:
    """Bundles all sheet metadata needed by generation functions.

    Created once per generate_list() call via build_sheet_context().
    Avoids threading the same 7-8 parameters through every function.
    """

    sheet_data: List[List[str]]
    id_cell: Any
    observation_cell: Any
    category_cell: Any
    num_participants: int
    study_name: str
    baseline_row_idx: Optional[int] = None
    filename_row_idx: Optional[int] = None

    @property
    def header_row(self) -> List[str]:
        return self.sheet_data[self.id_cell.row - 1] if self.id_cell.row > 0 else []

    @property
    def first_data_row_idx(self) -> int:
        """0-based index into sheet_data of the first data row (row after Observation header)."""
        return self.observation_cell.row


# ---- Validation and context building ----


def validate_spreadsheet_headers(sheet: Any) -> Optional[Tuple[Any, Any, Any]]:
    """Validate that required headers exist in the spreadsheet.

    Args:
        sheet: The gspread worksheet object

    Returns:
        tuple: (id_cell, observation_cell, category_cell) if all headers found,
               None if any header is missing
    """
    id_cell = sheet.find(config.ID_HEADER)
    observation_cell = sheet.find(config.OBSERVATION_HEADER)
    category_cell = sheet.find(config.CATEGORY_HEADER)

    missing_headers = []
    if id_cell is None:
        missing_headers.append(f"'{config.ID_HEADER}'")
    if observation_cell is None:
        missing_headers.append(f"'{config.OBSERVATION_HEADER}'")
    if category_cell is None:
        missing_headers.append(f"'{config.CATEGORY_HEADER}'")

    if missing_headers:
        utils.error_print(
            f"Required header(s) not found in spreadsheet: {', '.join(missing_headers)}",
            [
                f"The spreadsheet must contain columns with these exact headers: {config.ID_HEADER}, {config.OBSERVATION_HEADER}, {config.CATEGORY_HEADER}",
                "Please check your spreadsheet structure.",
            ],
        )
        return None

    return (id_cell, observation_cell, category_cell)


_BASELINE_ROW_CACHE: Dict[int, Optional[int]] = {}


def _detect_baseline_row(sheet_data: List[List[str]]) -> Optional[int]:
    """Detect the sheet-wide baseline row marked by a 'Baseline time' label.

    Scans the sheet once for a row where any cell value starts with 'Baseline time'
    (case-insensitive). That row is treated as the baseline row: each non-empty
    cell in that row provides the baseline timestamp for its column.
    """
    cache_key = id(sheet_data)
    if cache_key in _BASELINE_ROW_CACHE:
        return _BASELINE_ROW_CACHE[cache_key]

    baseline_row_idx: Optional[int] = None
    for row_idx, row in enumerate(sheet_data):
        for value in row:
            if not value:
                continue
            if str(value).strip().lower().startswith("baseline time"):
                baseline_row_idx = row_idx
                break
        if baseline_row_idx is not None:
            break

    _BASELINE_ROW_CACHE[cache_key] = baseline_row_idx
    return baseline_row_idx


def get_num_participants(header_row: List[str], id_cell: Any, col_count: int) -> int:
    """Count the number of participant columns in the worksheet.

    Looks for columns starting with participant prefixes (P or G) and stops
    when it encounters the NOTES_COLUMN.

    Args:
        header_row: List of header cell values
        id_cell: The ID header cell object
        col_count: Total number of columns in the worksheet

    Returns:
        Number of participant columns found
    """
    num_participants = sum(
        1
        for j in range(col_count)
        if header_row[j] and header_row[j][0] in config.PARTICIPANT_PREFIXES
    )
    utils.standard_print(
        f"Found {num_participants} participants in total, spanning columns {id_cell.col + 1} to {num_participants + id_cell.col + 1}."
    )
    return num_participants


def build_sheet_context(sheet: Any) -> Optional[SheetContext]:
    """Validate headers, load sheet data, and build a SheetContext.

    Returns None if validation fails (missing headers, empty sheet, no participants).
    """
    header_result = validate_spreadsheet_headers(sheet)
    if header_result is None:
        return None

    id_cell, observation_cell, category_cell = header_result
    if config.DEBUGGING:
        ic(id_cell, observation_cell, category_cell)

    sheet_data = sheet.get_all_values()
    utils.debug_print(f"Sheet dumped into memory at {utils.get_current_time()}")

    if len(sheet_data) <= 1:
        utils.error_print(
            "Spreadsheet appears to be empty (no data rows found).",
            [f"The spreadsheet only has {len(sheet_data)} row(s)."],
        )
        return None

    baseline_row_idx = _detect_baseline_row(sheet_data)

    study_name = sheet_data[0][0]
    if study_name == "":
        study_name = sheet.spreadsheet.title
    utils.standard_print(f"\nBeginning work on {study_name}.")
    study_name = utils.normalize_study_name(study_name)

    num_participants = get_num_participants(
        sheet.row_values(id_cell.row), id_cell, sheet.col_count
    )

    filename_cell = None
    filename_row_idx: Optional[int] = None
    try:
        filename_cell = sheet.find(config.FILENAME_HEADER)
    except Exception:
        filename_cell = None
    if filename_cell is not None:
        filename_row_idx = filename_cell.row - 1

    if num_participants == 0:
        utils.warning_print(
            "No participant columns found in the spreadsheet.",
            [
                f"Looking for columns starting with: {', '.join(config.PARTICIPANT_PREFIXES)}",
                "Check that participant column headers start with 'P' or 'G' (e.g., P01, P02, G01).",
            ],
        )
        return None

    return SheetContext(
        sheet_data=sheet_data,
        id_cell=id_cell,
        observation_cell=observation_cell,
        category_cell=category_cell,
        num_participants=num_participants,
        study_name=study_name,
        baseline_row_idx=baseline_row_idx,
        filename_row_idx=filename_row_idx,
    )


# ---- Parsing utilities ----


def get_participant_list(
    header_row: List[str], id_cell: Any, num_participants: int
) -> List[str]:
    """Return list of participant IDs from the header row.

    Args:
        header_row: List of header cell values
        id_cell: The ID header cell object
        num_participants: Number of participant columns

    Returns:
        List of participant IDs (e.g. ['P01', 'P02', 'G01'])
    """
    participants = []
    for j in range(id_cell.col, id_cell.col + num_participants):
        if j < len(header_row):
            participant_id = utils.normalize_participant_id(header_row[j]).strip()
            if participant_id:
                participants.append(participant_id)
    return participants


def parse_participant_selection(input_str: str) -> List[str]:
    """Parse a participant selection string into a list of IDs.

    Splits on + or , and returns non-empty stripped tokens.

    Args:
        input_str: User input (e.g. "P01 + P03" or "P01, P03" or "1, 3")

    Returns:
        List of participant ID strings (e.g. ['P01', 'P03'])
    """
    if not input_str or not input_str.strip():
        return []
    # Support both + and , as separators
    combined = input_str.replace(",", "+")
    return [s.strip() for s in combined.split("+") if s.strip()]


def parse_cell_specifications(cell_input: str) -> List[Tuple[str, int]]:
    """Parse cell specification string into list of (participant_id, row_number) tuples.

    Expected format: "P01.11" (participant_id.row_number). Multiple cells separated by
    + or , e.g. "P01.11 + P03.11" or "P01.11, P03.09". Participant ID must start with
    P or G; row_number must be a positive integer (1-based sheet row).

    Args:
        cell_input: String like "P01.11" or "P01.11 + P03.11 + P03.09"

    Returns:
        List of (participant_id, row_number) tuples

    Raises:
        ValueError: If format is invalid
    """
    # Support both + and , as separators; normalize to + then split
    cell_str = cell_input.replace(",", "+")
    specs = []

    for spec in cell_str.split("+"):
        spec = spec.strip()
        if not spec:
            continue

        # Exactly one dot: participant_id.row_number (split('.', 1) avoids splitting IDs like P01.02)
        if "." not in spec:
            raise ValueError(
                f'Invalid cell specification "{spec}". Expected format: P01.11'
            )

        parts = spec.split(".", 1)
        if len(parts) != 2:
            raise ValueError(
                f'Invalid cell specification "{spec}". Expected format: P01.11'
            )

        participant_id = parts[0].strip()
        row_str = parts[1].strip()

        # Validate participant ID format (should start with P or G)
        if not participant_id or participant_id[0] not in config.PARTICIPANT_PREFIXES:
            raise ValueError(
                f'Invalid participant ID "{participant_id}". Must start with {", ".join(config.PARTICIPANT_PREFIXES)}'
            )

        # Validate row number
        try:
            row_number = int(row_str)
            if row_number < 1:
                raise ValueError(f"Row number must be positive. Got: {row_number}")
        except ValueError as e:
            if "invalid literal" in str(e):
                raise ValueError(
                    f'Invalid row number "{row_str}". Must be a positive integer.'
                )
            raise

        specs.append((participant_id, row_number))

    return specs


def parse_reel_input(input_string: str) -> ReelInput:
    """Parse mixed reel selector input into structured selectors.

    Supports: batch, keyword, timeline, lines (e.g. 11, 12), ranges (e.g. 13-16),
    categories (quoted), cells (e.g. P01.11), participants (e.g. P01, P02).

    Args:
        input_string: Raw user input (e.g. '11, 13-16, P01, P02.15, "Observations"')

    Returns:
        Dict with keys: batch (bool), keyword (bool), timeline (bool), lines (list of int),
        ranges (list of (int,int)), categories (list of str), cells (list of (str,int)),
        participants (list of str)
    """
    result: ReelInput = {
        "batch": False,
        "keyword": False,
        "timeline": False,
        "lines": [],
        "ranges": [],
        "categories": [],
        "cells": [],
        "participants": [],
    }
    if not input_string or not input_string.strip():
        return result

    # Extract quoted strings (categories) first so commas inside quotes don't split the string
    rest = input_string.strip()
    while True:
        match = re.search(
            r'"([^"]*)"', rest
        )  # "category name" -> group(1) is category name
        if not match:
            break
        result["categories"].append(match.group(1).strip())
        rest = rest[: match.start()] + " " + rest[match.end() :]

    # Split remaining by comma; each part is one token. Token type is inferred in fixed order below.
    # Order of checks per token (do not reorder): batch, keyword, timeline, range, line, cell, participant
    parts = [p.strip() for p in rest.split(",") if p.strip()]
    seen_lines = set()
    seen_ranges = set()
    seen_cells = set()
    seen_participants = set()

    for part in parts:
        token = part.strip()
        if not token:
            continue
        # Batch keyword: literal "batch"
        if token.lower() == "batch":
            result["batch"] = True
            continue
        # Keyword mode: literal "keyword"
        if token.lower() == "keyword":
            result["keyword"] = True
            continue
        # Timeline keyword: literal "timeline"
        if token.lower() == "timeline":
            result["timeline"] = True
            continue
        # Range: "13-16" -> start, end (digits hyphen digits)
        range_match = re.match(r"^(\d+)\s*-\s*(\d+)$", token)
        if range_match:
            start, end = int(range_match.group(1)), int(range_match.group(2))
            if start <= end and (start, end) not in seen_ranges:
                seen_ranges.add((start, end))
                result["ranges"].append((start, end))
            continue
        # Single line number: plain digits (e.g. "11", "12")
        if token.isdigit():
            line_num = int(token)
            if line_num not in seen_lines:
                seen_lines.add(line_num)
                result["lines"].append(line_num)
            continue
        # Cell: "P01.11" or "G02.5" -> (participant_id, row_number); regex: P or G, word chars, dot, digits
        if "." in token:
            cell_match = re.match(r"^([PG]\w*)\.(\d+)$", token, re.IGNORECASE)
            if cell_match:
                pid, row_str = cell_match.group(1), cell_match.group(2)
                if pid[0].upper() in config.PARTICIPANT_PREFIXES:
                    row_num = int(row_str)
                    if row_num >= 1:
                        key = (pid.upper() if len(pid) <= 3 else pid, row_num)
                        if key not in seen_cells:
                            seen_cells.add(key)
                            result["cells"].append((pid, row_num))
            continue
        # Participant only: "P01" or "G02" (no dot) -> include all rows for that participant
        if token[0].upper() in config.PARTICIPANT_PREFIXES and "." not in token:
            key = token.upper() if len(token) <= 3 else token
            if key not in seen_participants:
                seen_participants.add(key)
                result["participants"].append(token)

    return result


def detect_mode_from_input(input_string: str) -> Tuple[Optional[str], dict]:
    """Detect mode from input syntax (for implicit mode selection).

    Only line, range, cell, and participant modes are auto-detected. Batch, browse,
    reel, and category require explicit mode selection. If input matches mixed types
    (e.g. "5, P01.11"), returns (None, {}) so the user can use reel mode explicitly.

    Args:
        input_string: Raw user input after worksheet selection.

    Returns:
        (mode, kwargs_dict) where kwargs_dict contains arguments for generate_list(),
        or (None, {}) if input does not match exactly one auto-detectable pattern.
    """
    if not input_string or not input_string.strip():
        return (None, {})

    parsed = parse_reel_input(input_string)

    # Ignore batch and categories for auto-detection
    has_lines = len(parsed["lines"]) > 0
    has_ranges = len(parsed["ranges"]) > 0
    has_cells = len(parsed["cells"]) > 0
    has_participants = len(parsed["participants"]) > 0

    types_present = sum([has_lines, has_ranges, has_cells, has_participants])
    if types_present == 0:
        return (None, {})
    if types_present > 1:
        return (None, {})

    if has_lines:
        return ("line", {"line_numbers": parsed["lines"]})
    if has_ranges:
        # Only auto-detect range when there is exactly one range
        if len(parsed["ranges"]) != 1:
            return (None, {})
        start, end = parsed["ranges"][0]
        return ("range", {"range_start": start, "range_end": end})
    if has_cells:
        return ("cell", {"cell_specs": parsed["cells"]})
    if has_participants:
        # participant_id is a string that parse_participant_selection() will parse
        participant_id = ",".join(parsed["participants"])
        return ("participant", {"participant_id": participant_id})

    return (None, {})


# ---- Core record building ----


def find_participant_column(
    header_row: List[str], id_cell: Any, participant_id: str
) -> Optional[int]:
    """Find the column index for a given participant ID.

    Args:
        header_row: List of header cell values
        id_cell: The ID header cell object
        participant_id: Participant ID to find (e.g., "P01")

    Returns:
        Column index (0-based) if found, None otherwise
    """
    normalized_target = utils.normalize_participant_id(participant_id).lower()
    for col_idx in range(id_cell.col - 1, len(header_row)):
        header_value = utils.normalize_participant_id(header_row[col_idx]).strip()
        if header_value.lower() == normalized_target:
            return col_idx
    return None


def _make_clip_record(
    ctx: SheetContext, row_idx: int, col_idx: int, cell_value: str
) -> ClipRecord:
    """Build one clip record dict for a cell at (row_idx, col_idx).

    Clip record structure (used by downstream prepare_clip/video code):
    - cell: gspread Cell with 1-based row/col and the timestamp value.
    - desc: Observation/description text from the same row (observation column).
    - study: Normalized study name for output paths.
    - participant: Participant ID from the header row for this column.
    - category: Category label from the same row (category column).
    'times' is added later when timestamps are resolved to [start, end] ranges.
    """
    # gspread Cell uses 1-based coordinates; convert from 0-based list indices
    cell = gspread.cell.Cell(row_idx + 1, col_idx + 1, cell_value)
    # observation_cell.col is 1-based; convert to 0-based for sheet_data
    desc_col = ctx.observation_cell.col - 1
    category_col = ctx.observation_cell.col - 2
    # Header row in sheet_data is 0-based; id_cell.row is 1-based
    participant_row = ctx.id_cell.row - 1
    desc = ""
    if 0 <= desc_col < len(ctx.sheet_data[row_idx]):
        desc = ctx.sheet_data[row_idx][desc_col]
    participant = ""
    if 0 <= participant_row < len(ctx.sheet_data) and col_idx < len(
        ctx.sheet_data[participant_row]
    ):
        participant = utils.normalize_participant_id(
            ctx.sheet_data[participant_row][col_idx]
        )
    timestamp_baseline = ""
    if ctx.baseline_row_idx is not None:
        if 0 <= ctx.baseline_row_idx < len(ctx.sheet_data) and col_idx < len(
            ctx.sheet_data[ctx.baseline_row_idx]
        ):
            timestamp_baseline = ctx.sheet_data[ctx.baseline_row_idx][col_idx].strip()
    category = ""
    if 0 <= category_col < len(ctx.sheet_data[row_idx]):
        category = ctx.sheet_data[row_idx][category_col]
    result: ClipRecord = {
        "cell": cell,
        "desc": desc,
        "study": ctx.study_name,
        "participant": participant,
        "category": category,
    }
    if timestamp_baseline:
        result["timestamp_baseline"] = timestamp_baseline
    if ctx.filename_row_idx is not None:
        if 0 <= ctx.filename_row_idx < len(ctx.sheet_data) and col_idx < len(
            ctx.sheet_data[ctx.filename_row_idx]
        ):
            filename_override = ctx.sheet_data[ctx.filename_row_idx][col_idx].strip()
            if filename_override:
                result["source_filename"] = filename_override
    return result


def get_line_timestamps(ctx: SheetContext, line_index: int) -> List[ClipRecord]:
    """Extract timestamp data from a single row in the spreadsheet as clip records.

    Processes all participant columns in the specified row and creates
    clip issue dictionaries for each timestamp found.

    Args:
        ctx: Sheet context with all metadata
        line_index: Zero-based row index into sheet_data (sheet row = line_index + 1)

    Returns:
        List of clip records, one per timestamp found
    """
    if config.DEBUGGING:
        ic(line_index, ctx.num_participants, ctx.study_name)
    utils.debug_print(
        f"Running method get_line_timestamps, starting line index {line_index} (real sheet line {line_index + 1})"
    )

    if line_index < 0 or line_index >= len(ctx.sheet_data):
        if config.DEBUGGING:
            ic(line_index, len(ctx.sheet_data))
        utils.error_print(
            f"Line index {line_index} (row {line_index + 1}) is out of bounds.",
            [f"Spreadsheet has {len(ctx.sheet_data)} rows."],
        )
        return []

    clips = []
    try:
        for col_index, value in enumerate(ctx.sheet_data[line_index]):
            utils.debug_print(f"Item {col_index} with value '{value}' being processed.")
            if col_index < ctx.id_cell.col:
                utils.debug_print(f"Skipping item {col_index} with value '{value}'")
            elif col_index == ctx.id_cell.col + ctx.num_participants:
                utils.debug_print(
                    f"Exit for-loop, reached final column {col_index} (real sheet column {col_index + 1})."
                )
                break
            elif value is None or value == "":
                pass
            else:
                issue = _make_clip_record(ctx, line_index, col_index, value)
                if config.DEBUGGING:
                    ic(
                        issue.get("participant"),
                        issue.get("desc"),
                        issue.get("category"),
                    )
                    ic(issue)
                utils.debug_print(
                    f"Participant ID at R{ctx.id_cell.row},C{col_index} -> '{issue.get('participant')}'"
                )
                utils.debug_print(
                    f"Description at R{line_index},C{ctx.observation_cell.col - 1} -> '{issue.get('desc')}'"
                )
                utils.debug_print(
                    f"Timestamp at R{issue['cell'].row - 1},C{issue['cell'].col - 1} -> '{issue['cell'].value}'"
                )
                utils.debug_print(
                    f"Actual cell at address {gspread.utils.rowcol_to_a1(issue['cell'].row, issue['cell'].col)}"
                )
                clips.append(issue)
                display_value = value.replace("\n", " ")
                cell_addr = gspread.utils.rowcol_to_a1(
                    issue["cell"].row, issue["cell"].col
                )
                utils.verbose_print(
                    f"+ Found timestamp: {display_value} at address {cell_addr}"
                )
    except IndexError as e:
        if config.DEBUGGING:
            ic(e, line_index)
        utils.error_print(
            f"Index error while reading row {line_index + 1}: {e}",
            ["The spreadsheet structure may be malformed."],
        )

    utils.debug_print(
        f"Line completed, returning list of {len(clips)} potential clips."
    )
    if config.DEBUGGING:
        ic(clips)
    return clips


# ---- Mode generators ----


def _validate_row_range(
    start_line: int, end_line: int, max_row: int
) -> Optional[Tuple[int, int]]:
    """Validate row range; print error and return None if invalid.

    Args:
        start_line: Start row (1-based)
        end_line: End row (1-based)
        max_row: Number of rows in sheet (len(sheet_data))

    Returns:
        (start_line, end_line) if valid, None otherwise
    """
    if start_line < 1 or end_line < 1:
        utils.error_print(
            f"Line numbers must be positive. Got start={start_line}, end={end_line}"
        )
        return None
    if start_line > max_row or end_line > max_row:
        utils.error_print(
            f"Line number(s) out of range. Spreadsheet has {max_row} rows.",
            [f"Requested: lines {start_line} to {end_line}"],
        )
        return None
    if start_line > end_line:
        utils.error_print(
            f"Start line ({start_line}) must be less than or equal to end line ({end_line})."
        )
        return None
    return (start_line, end_line)


def generate_list(
    sheet: Any,
    mode: str,
    line_numbers: Optional[List[int]] = None,
    range_start: Optional[int] = None,
    range_end: Optional[int] = None,
    skip_prompts: bool = False,
    cell_specs: Optional[List[Tuple[str, int]]] = None,
    participant_id: Optional[str] = None,
    reel_input: Optional[str] = None,
    categories: Optional[List[str]] = None,
) -> List[ClipRecord]:
    """Generate clip records from a sheet based on mode and resolved parameters.

    This function is pure: it takes resolved parameters (no interactive prompts).
    Interactive prompts are handled by clipgen.py using functions from interactive.py.

    Args:
        sheet: The gspread worksheet object
        mode: One of 'batch', 'line', 'range', 'category', 'cell', 'participant', 'keyword', 'reel'
        line_numbers: List of line numbers for 'line' mode
        range_start: Start line for 'range' mode
        range_end: End line for 'range' mode
        skip_prompts: If True, skip confirmation prompts (CLI -y flag, for batch/keyword)
        cell_specs: List of (participant_id, row_number) tuples for 'cell' mode
        participant_id: Participant ID(s) for 'participant' mode (comma/plus-separated string)
        reel_input: Reel selector string for 'reel' mode
        categories: List of category names for 'category' mode

    Returns:
        List of clip records
    """
    if config.DEBUGGING:
        ic(mode, line_numbers, range_start, range_end)

    ctx = build_sheet_context(sheet)
    if ctx is None:
        return []

    # Generate clips according to the selected mode
    if mode == "batch":
        utils.standard_print("Batch mode: generating all possible clips...")
        return generate_batch_timestamps(ctx)

    if mode == "category":
        if not categories:
            utils.error_print(
                "Category mode requires categories list.",
                ["Pass categories via CLI (-C) or use interactive mode."],
            )
            return []
        return generate_category_timestamps(ctx, categories)

    if mode == "line":
        if line_numbers is None:
            utils.error_print(
                "Line mode requires line_numbers list.",
                ["Pass line numbers via CLI (-l) or use interactive mode."],
            )
            return []
        return generate_line_timestamps(ctx, line_numbers)

    if mode == "range":
        if range_start is None or range_end is None:
            utils.error_print(
                "Range mode requires range_start and range_end.",
                ["Pass range via CLI (-r) or use interactive mode."],
            )
            return []
        max_row = len(ctx.sheet_data)
        valid = _validate_row_range(range_start, range_end, max_row)
        if valid is None:
            return []
        range_start, range_end = valid
        utils.standard_print(f"Range mode: lines {range_start} to {range_end}")
        utils.standard_print(
            f"Lines selected: {ctx.sheet_data[range_start - 1][ctx.observation_cell.col - 1]} to {ctx.sheet_data[range_end - 1][ctx.observation_cell.col - 1]}"
        )
        return generate_range_timestamps(ctx, range_start, range_end)

    if mode == "cell":
        if cell_specs is None:
            utils.error_print(
                "Cell mode requires cell_specs list.",
                ["Pass cell specs via CLI (-c) or use interactive mode."],
            )
            return []
        utils.standard_print(f"Cell mode: processing {len(cell_specs)} cell(s)")
        return generate_cell_timestamps(ctx, cell_specs)

    if mode == "participant":
        if participant_id is None:
            utils.error_print(
                "Participant mode requires participant_id.",
                ["Pass participant ID via CLI (-p) or use interactive mode."],
            )
            return []
        available_list = get_participant_list(
            ctx.header_row, ctx.id_cell, ctx.num_participants
        )
        participant_ids = parse_participant_selection(participant_id)
        if not participant_ids:
            utils.error_print(
                "No participant ID(s) provided.",
                ["Use format: P01 or P01+P03 or P01, P03"],
            )
            return []
        invalid = []
        for pid in participant_ids:
            if find_participant_column(ctx.header_row, ctx.id_cell, pid) is None:
                invalid.append(pid)
        if invalid:
            utils.error_print(
                f"Participant(s) not found in spreadsheet headers: {', '.join(invalid)}",
                [f"Available participants: {', '.join(available_list)}"],
            )
            return []
        utils.standard_print(
            f"Participant mode: generating all clips for {', '.join(participant_ids)}"
        )
        clips = []
        for pid in participant_ids:
            clips.extend(generate_participant_timestamps(ctx, pid))
        return clips

    if mode == "keyword":
        utils.standard_print("Keyword mode: generating key-marked clips...")
        return generate_keyword_timestamps(ctx)

    if mode == "reel":
        if reel_input is None or not reel_input.strip():
            utils.info_print("Reel mode: no input provided.")
            return []
        utils.standard_print("Reel mode: parsing selectors and collecting clips...")
        return generate_reel_timestamps(ctx, reel_input.strip())

    return []


def generate_batch_timestamps(ctx: SheetContext) -> List[ClipRecord]:
    """Generate clip records for all rows in batch mode."""
    utils.debug_print("Running method generate_batch_timestamps()")
    clips = []
    for i in range(ctx.first_data_row_idx, len(ctx.sheet_data)):
        if ctx.filename_row_idx is not None and i == ctx.filename_row_idx:
            continue
        utils.debug_print(f"Batching on line {i} (real sheet line {i + 1})")
        clips.extend(get_line_timestamps(ctx, i))
    return clips


def generate_keyword_timestamps(ctx: SheetContext) -> List[ClipRecord]:
    """Generate key-marked clips from the entire sheet based on cell content.

    Semantics:
    - Segment-level: annotation tokens like `!key` in a timestamp cell mark the
      preceding parseable timestamp segment(s) within that cell.
    - Header/participant-level annotations in the header row are ignored here;
      keyword mode is driven purely by per-cell annotations.
    """
    clips = generate_batch_timestamps(ctx)
    if not clips:
        return []

    filtered_clips: List[ClipRecord] = []
    for clip in clips:
        _, segment_annotations, _ = utils.parse_cell_annotations(clip["cell"].value)
        key_indexes = sorted(segment_annotations.get("key", set()))
        if key_indexes:
            clip["selected_segment_indexes"] = key_indexes
            filtered_clips.append(clip)

    return filtered_clips


def collect_categories(ctx: SheetContext) -> List[str]:
    """Scan sheet and return unique categories in order of first appearance."""
    categories = []
    category_col = ctx.category_cell.col - 1
    for i in range(ctx.first_data_row_idx, len(ctx.sheet_data)):
        category = ctx.sheet_data[i][category_col].strip()
        if category and category not in categories:
            categories.append(category)
    return categories


def generate_category_timestamps(
    ctx: SheetContext, selected_categories: List[str]
) -> List[ClipRecord]:
    """Generate clip records for all rows matching any of the selected categories."""
    utils.debug_print("Starting method generate_category_timestamps()")
    clips = []
    category_col = ctx.category_cell.col - 1
    for i in range(ctx.first_data_row_idx, len(ctx.sheet_data)):
        if ctx.filename_row_idx is not None and i == ctx.filename_row_idx:
            continue
        row_category = ctx.sheet_data[i][category_col].strip()
        if row_category in selected_categories:
            utils.debug_print(f"Row {i + 1} matches category '{row_category}'")
            clips.extend(get_line_timestamps(ctx, i))
    return clips


def generate_line_timestamps(
    ctx: SheetContext, line_numbers: List[int]
) -> List[ClipRecord]:
    """Generate clip records for one or more line/row numbers.

    Args:
        ctx: Sheet context with all metadata
        line_numbers: List of 1-based row numbers (already validated)

    Returns:
        List of clip records
    """
    valid_lines = []
    utils.standard_print(
        f"Line mode: processing lines {', '.join(str(n) for n in line_numbers)}"
    )
    utils.standard_print("Selected issues:")
    for line_num in line_numbers:
        if line_num < 1 or line_num > len(ctx.sheet_data):
            utils.standard_print(f"  Line {line_num}: [INVALID - out of range]")
        elif ctx.filename_row_idx is not None and line_num - 1 == ctx.filename_row_idx:
            utils.standard_print(
                f"  Line {line_num}: [RESERVED - filename overrides row]"
            )
        else:
            desc = ctx.sheet_data[line_num - 1][ctx.observation_cell.col - 1]
            utils.standard_print(f"  Line {line_num}: {desc}")
            valid_lines.append(line_num)

    if not valid_lines:
        utils.info_print("No valid lines found. Exiting.")
        return []

    clips = []
    for line_num in valid_lines:
        utils.debug_print(
            f"Calling get_line_timestamps() from generate_line_timestamps() for line {line_num}"
        )
        clips.extend(get_line_timestamps(ctx, line_num - 1))

    utils.debug_print(
        f"Printing return of get_line_timestamps() in generate_line_timestamps(): {len(clips)} total clips"
    )
    utils.debug_print(str(clips))
    return clips


def generate_range_timestamps(
    ctx: SheetContext, start_line: int, end_line: int
) -> List[ClipRecord]:
    """Generate clip records for a range of rows.

    Args:
        ctx: Sheet context with all metadata
        start_line: Starting row number (1-based)
        end_line: Ending row number (1-based, inclusive)
    """
    clips = []
    for i in range(start_line - 1, end_line):
        utils.debug_print(f"Batching on line {i}")
        clips.extend(get_line_timestamps(ctx, i))
    return clips


def generate_participant_timestamps(
    ctx: SheetContext, participant_id: str
) -> List[ClipRecord]:
    """Generate clip records for all rows in a single participant's column."""
    utils.debug_print("Starting method generate_participant_timestamps()")
    col_idx = find_participant_column(ctx.header_row, ctx.id_cell, participant_id)
    if col_idx is None:
        return []
    clips = []
    for row_idx in range(ctx.first_data_row_idx, len(ctx.sheet_data)):
        if ctx.filename_row_idx is not None and row_idx == ctx.filename_row_idx:
            continue
        if col_idx >= len(ctx.sheet_data[row_idx]):
            continue
        cell_value = ctx.sheet_data[row_idx][col_idx]
        if not cell_value or not cell_value.strip():
            continue
        issue = _make_clip_record(ctx, row_idx, col_idx, cell_value)
        clips.append(issue)
        display_value = cell_value.replace("\n", " ")
        cell_addr = gspread.utils.rowcol_to_a1(issue["cell"].row, issue["cell"].col)
        utils.verbose_print(
            f"+ Found timestamp: {display_value} at row {row_idx + 1} ({cell_addr})"
        )
    return clips


def generate_cell_timestamps(
    ctx: SheetContext, cell_specs: List[Tuple[str, int]]
) -> List[ClipRecord]:
    """Generate clip records for specific cells.

    Args:
        ctx: Sheet context with all metadata
        cell_specs: List of (participant_id, row_number) tuples

    Returns:
        List of clip records
    """
    utils.debug_print("Starting method generate_cell_timestamps()")
    clips = []

    for participant_id, row_number in cell_specs:
        col_idx = find_participant_column(ctx.header_row, ctx.id_cell, participant_id)
        if col_idx is None:
            utils.warning_print(
                f"Participant '{participant_id}' not found in spreadsheet headers.",
                [
                    f"Available participants start with: {', '.join(config.PARTICIPANT_PREFIXES)}"
                ],
            )
            continue

        row_idx = row_number - 1
        if row_idx < 0 or row_idx >= len(ctx.sheet_data):
            utils.warning_print(
                f"Row {row_number} is out of range.",
                [
                    f"Spreadsheet has {len(ctx.sheet_data)} rows (valid range: 1-{len(ctx.sheet_data)})."
                ],
            )
            continue
        if ctx.filename_row_idx is not None and row_idx == ctx.filename_row_idx:
            utils.warning_print(
                f"Row {row_number} is reserved for filename overrides and will be skipped."
            )
            continue

        if col_idx >= len(ctx.sheet_data[row_idx]):
            utils.warning_print(
                f"Column index {col_idx + 1} is out of range for row {row_number}."
            )
            continue

        cell_value = ctx.sheet_data[row_idx][col_idx]
        if not cell_value or cell_value.strip() == "":
            utils.verbose_print(
                f"Cell {participant_id}.{row_number} is empty, skipping."
            )
            continue

        issue = _make_clip_record(ctx, row_idx, col_idx, cell_value)
        # Use actual header value for participant when available
        participant_row = ctx.id_cell.row - 1
        if (
            0 <= participant_row < len(ctx.sheet_data)
            and col_idx < len(ctx.sheet_data[participant_row])
            and ctx.sheet_data[participant_row][col_idx]
        ):
            issue["participant"] = utils.normalize_participant_id(
                ctx.sheet_data[participant_row][col_idx]
            )
        clips.append(issue)
        display_value = cell_value.replace("\n", " ")
        cell_addr = gspread.utils.rowcol_to_a1(issue["cell"].row, issue["cell"].col)
        utils.verbose_print(
            f"+ Found timestamp: {display_value} at cell {participant_id}.{row_number} ({cell_addr})"
        )

    return clips


def sort_clips_chronologically(clips: List[ClipRecord]) -> None:
    """Sort clip records in-place by earliest start timestamp in each clip cell.

    Cells are normalized through parse_cell_annotations first, then parsed via
    parse_timestamps. Unparseable/empty timestamps are sorted last.
    """

    def _clip_start_seconds(clip: Any) -> float:
        cell_value = str(clip.get("cell").value if clip.get("cell") is not None else "")
        cleaned_value, _, _ = utils.parse_cell_annotations(cell_value)
        parsed_times = utils.parse_timestamps(cleaned_value)
        if not parsed_times:
            return float("inf")
        first_start = parsed_times[0][0]
        seconds = utils.timestamp_to_seconds(first_start)
        if seconds is None:
            return float("inf")
        return seconds

    clips.sort(key=_clip_start_seconds)


def generate_reel_timestamps(
    ctx: SheetContext, reel_input_string: str
) -> List[ClipRecord]:
    """Generate clip records for reel mode by combining multiple selector types and deduplicating.

    Parses reel input (batch, keyword, timeline, lines, ranges, categories, cells, participants),
    collects timestamps from each selector, deduplicates by cell (row, col), and returns a single
    ordered list.
    """
    selectors = parse_reel_input(reel_input_string)
    if selectors["timeline"] and len(selectors["participants"]) != 1:
        utils.error_print(
            "Timeline selector requires exactly one participant.",
            [
                "Use timeline with one participant, e.g. 'timeline, P01'.",
                "Timeline reels are generated per participant.",
            ],
        )
        return []

    has_any = (
        selectors["batch"]
        or selectors["keyword"]
        or selectors["timeline"]
        or selectors["lines"]
        or selectors["ranges"]
        or selectors["categories"]
        or selectors["cells"]
        or selectors["participants"]
    )
    if not has_any:
        return []

    all_issues: List[ClipRecord] = []

    if selectors["keyword"]:
        all_issues.extend(generate_keyword_timestamps(ctx))
    if selectors["batch"]:
        all_issues.extend(generate_batch_timestamps(ctx))
    if selectors["lines"]:
        all_issues.extend(generate_line_timestamps(ctx, selectors["lines"]))
    for start_line, end_line in selectors["ranges"]:
        all_issues.extend(generate_range_timestamps(ctx, start_line, end_line))
    if selectors["categories"]:
        all_issues.extend(generate_category_timestamps(ctx, selectors["categories"]))
    if selectors["cells"]:
        all_issues.extend(generate_cell_timestamps(ctx, selectors["cells"]))
    for participant_id in selectors["participants"]:
        all_issues.extend(generate_participant_timestamps(ctx, participant_id))

    # Deduplicate by cell (row, col)
    seen: Set[Tuple[int, int]] = set()
    deduped: List[ClipRecord] = []
    for issue in all_issues:
        key = (issue["cell"].row, issue["cell"].col)
        if key not in seen:
            seen.add(key)
            deduped.append(issue)

    if selectors["timeline"]:
        timeline_participant = utils.normalize_participant_id(
            selectors["participants"][0]
        ).lower()
        deduped = [
            issue
            for issue in deduped
            if utils.normalize_participant_id(issue.get("participant", "")).lower()
            == timeline_participant
        ]
        sort_clips_chronologically(deduped)
    else:
        deduped.sort(key=lambda issue: (issue["cell"].row, issue["cell"].col))

    return deduped
