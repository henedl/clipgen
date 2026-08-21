"""Spreadsheet data processing for clipgen.

Expected spreadsheet layout (see README.md for a reference example):
- Row 0 (A1): Study name; optional — falls back to spreadsheet title when empty.
- Header row: Must contain required columns ID, Observation, Category (exact names from config).
  Optional columns: Severity (numeric -4..+2 or labels like Critical/High; enables severity mode)
  and Filename (overrides the source video filename per participant column).
- Participant columns: Immediately follow the ID column; headers start with P or G (e.g. P01, G02).
  Each column holds timestamp strings; non-empty cells become clip candidates.
- Observation column: Human-readable description per row; included in clip metadata and filenames.
- Category column: Label per row; used for category and reel selection.

Optional baseline time row:
- A row whose cells include the label "Baseline time" marks that row as a clock baseline.
- Per-participant baseline timestamps in that row (e.g. "09:12:00" under P01) mean that
  participant column uses absolute/clock timestamps.
- During files.prepare_clip(), absolute (start, end) pairs are converted to relative offsets
  by subtracting the per-column baseline via utils.convert_clock_pairs_to_relative().
- Participant columns without a baseline cell use relative timestamps as-is.
- If the marker row is absent entirely, all columns are treated as relative.
- Baseline row placement is tied to header/``id_cell`` row math (offsets from
  ``id_cell.row``); changing that offset without aligning tests and sheet
  layout has broken baseline timestamp handling before.

Optional Filename row (source video override):
- A "Filename" header marks a row whose per-participant cells override the source video
  filename for that column; the value is stored verbatim as clip['source_filename'] and is
  never parsed as timestamps (so a '+' here is safe, unlike '+' inside a timestamp cell).
- A cell may list several files plus-separated, order matters — "morning.mp4 + afternoon.mp4" —
  to treat a participant's session as one continuous timeline spanning multiple source videos.
  Without an override, the pipeline auto-detects numbered parts on disk (study_P01-1.mp4, ...).

Coordinate system:
- gspread uses 1-based row/col (e.g. Cell(row=1, col=1) is A1).
- sheet.get_all_values() returns a list of lists with 0-based indices: sheet_data[row][col].
- Conversions: sheet row = row_idx + 1, sheet col = col_idx + 1; header cell .row/.col are 1-based.

Clip record (returned by generation functions):
- Dict with keys: cell (gspread Cell), desc (observation text), study (normalized name),
  participant (header value for that column), category (row category), severity (label or "").
  The 'times' key is added later by files.prepare_clip() when resolving timestamps to time ranges.
"""

import re
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Any, NamedTuple

import config
import google_api
import profiling
import utils
from utils import ClipRecord, ReelInput


# ---- Sheet context dataclass ----


@dataclass
class SheetContext:
    """Bundles all sheet metadata needed by generation functions.

    Created once per generate_list() call via build_sheet_context().
    Avoids threading the same 7-8 parameters through every function.
    """

    sheet_data: list[list[str]]
    id_cell: Any
    observation_cell: Any
    category_cell: Any
    num_participants: int
    study_name: str
    baseline_row_idx: int | None = None
    filename_row_idx: int | None = None
    severity_cell: Any = None

    @property
    def header_row(self) -> list[str]:
        return self.sheet_data[self.id_cell.row - 1] if self.id_cell.row > 0 else []

    @property
    def first_data_row_idx(self) -> int:
        """0-based index into sheet_data of the first data row (row after Observation header)."""
        return self.observation_cell.row


# ---- Validation and context building ----


class _CellLike(NamedTuple):
    """Minimal cell reference with 1-based row and col."""

    row: int
    col: int


def _find_in_data(sheet_data: list[list[str]], text: str) -> _CellLike | None:
    """Find first cell with exact text match in pre-loaded sheet data.

    Returns _CellLike(row, col) with 1-based coordinates, or None.
    """
    for row_idx, row in enumerate(sheet_data):
        for col_idx, cell_value in enumerate(row):
            if cell_value == text:
                return _CellLike(row=row_idx + 1, col=col_idx + 1)
    return None


def validate_spreadsheet_headers(
    sheet_data: list[list[str]],
) -> tuple[_CellLike, _CellLike, _CellLike] | None:
    """Validate that required headers exist in pre-loaded sheet data.

    Args:
        sheet_data: Pre-loaded sheet data (list of lists from get_all_values()).

    Returns:
        tuple: (id_cell, observation_cell, category_cell) if all headers found,
               None if any header is missing
    """
    id_cell = _find_in_data(sheet_data, config.ID_HEADER)
    observation_cell = _find_in_data(sheet_data, config.OBSERVATION_HEADER)
    category_cell = _find_in_data(sheet_data, config.CATEGORY_HEADER)

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

    if id_cell is None or observation_cell is None or category_cell is None:
        return None
    return (id_cell, observation_cell, category_cell)


def _detect_baseline_row(sheet_data: list[list[str]]) -> int | None:
    """Detect the sheet-wide baseline row marked by a 'Baseline time' label.

    Scans the sheet once for a row where any cell value starts with 'Baseline time'
    (case-insensitive). That row is treated as the baseline row: each non-empty
    cell in that row provides the baseline timestamp for its column.
    """
    for row_idx, row in enumerate(sheet_data):
        for value in row:
            if not value:
                continue
            if str(value).strip().lower().startswith("baseline time"):
                return row_idx
    return None


def get_num_participants(header_row: list[str], id_cell: Any, col_count: int) -> int:
    """Count the number of participant columns in the worksheet.

    Scans every column after ID and counts those whose header starts with one of
    ``config.PARTICIPANT_PREFIXES`` (P / G). Layout-agnostic: assumes nothing about
    where Observation, Category or other non-participant columns sit relative to
    ID. *id_cell* is the ID header cell (1-based ``col``).
    """
    start_col = (
        id_cell.col
    )  # id_cell.col is 1-based, so this is the 0-based index of the next column
    num_participants = sum(
        1
        for j in range(start_col, col_count)
        if j < len(header_row)
        and header_row[j]
        and header_row[j][0] in config.PARTICIPANT_PREFIXES
    )
    utils.standard_print(
        f"Found {num_participants} participants in total, spanning columns {id_cell.col + 1} to {num_participants + id_cell.col}."
    )
    return num_participants


def build_sheet_context(sheet: Any) -> SheetContext | None:
    """Validate headers, load sheet data, and build a SheetContext.

    Makes exactly one API call (get_all_values); all header lookups use local data.
    Returns None if validation fails (missing headers, empty sheet, no participants).
    """
    sheet_data = google_api.get_sheet_values(sheet)
    utils.debug_print(f"Sheet dumped into memory at {utils.get_current_time()}")

    if len(sheet_data) <= 1:
        utils.error_print(
            "Spreadsheet appears to be empty (no data rows found).",
            [f"The spreadsheet only has {len(sheet_data)} row(s)."],
        )
        return None

    header_result = validate_spreadsheet_headers(sheet_data)
    if header_result is None:
        return None

    id_cell, observation_cell, category_cell = header_result
    if config.DEBUGGING:
        config.debug_ic(id_cell, observation_cell, category_cell)

    baseline_row_idx = _detect_baseline_row(sheet_data)

    study_name = sheet_data[0][0]
    if study_name == "":
        # A second API round-trip, taken only when A1 is blank — so the
        # "exactly one API call" claim above holds for the common case but not
        # this one. Counted rather than hidden.
        with profiling.span("sheets.spreadsheet_title"):
            study_name = sheet.spreadsheet.title
    utils.standard_print(f"\nBeginning work on {study_name}.")
    study_name = utils.normalize_study_name(study_name)

    header_row = sheet_data[id_cell.row - 1]
    col_count = max(len(row) for row in sheet_data)
    num_participants = get_num_participants(header_row, id_cell, col_count)

    filename_cell = _find_in_data(sheet_data, config.FILENAME_HEADER)
    filename_row_idx: int | None = None
    if filename_cell is not None:
        filename_row_idx = filename_cell.row - 1

    severity_cell = _find_in_data(sheet_data, config.SEVERITY_HEADER)

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
        severity_cell=severity_cell,
    )


# ---- Parsing utilities ----


def get_participant_list(
    header_row: list[str], id_cell: Any, num_participants: int
) -> list[str]:
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


def parse_participant_selection(input_str: str) -> list[str]:
    """Parse a participant selection string into a list of IDs.

    Splits on + or , and returns non-empty stripped tokens.

    Args:
        input_str: User input (e.g. "P01 + P03" or "P01, P03" or "1, 3")

    Returns:
        List of participant ID strings (e.g. ['P01', 'P03'])
    """
    # Support both + and , as separators
    return utils.split_selector_tokens(input_str)


def parse_cell_specifications(cell_input: str) -> list[tuple[str, int]]:
    """Parse cell specification string into list of (participant_id, row_number) tuples.

    Format is "P01.11" (participant_id.row_number), multiple cells separated by +
    or , — "P01.11 + P03.11", "P01.11, P03.09". Participant ID must start with P or
    G, row_number must be a positive 1-based sheet row. Raises ValueError on a
    malformed input.
    """
    # Support both + and , as separators; normalize to + then split
    specs = []

    for spec in utils.split_selector_tokens(cell_input):
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
        except ValueError:
            raise ValueError(
                f'Invalid row number "{row_str}". Must be a positive integer.'
            )
        if row_number < 1:
            raise ValueError(f"Row number must be positive. Got: {row_number}")

        specs.append((participant_id, row_number))

    return specs


def parse_reel_input(input_string: str) -> ReelInput:
    """Parse mixed reel selector input into structured selectors.

    Supports: batch, keyword, chronologic, lines (e.g. 11, 12), ranges (e.g. 13-16),
    categories (quoted), cells (e.g. P01.11), participants (e.g. P01, P02).

    Args:
        input_string: Raw user input (e.g. '11, 13-16, P01, P02.15, "Observations"')

    Returns:
        Dict with keys: batch (bool), keyword (bool), chronologic (bool), lines (list of int),
        ranges (list of (int,int)), categories (list of str), cells (list of (str,int)),
        participants (list of str)
    """
    result: ReelInput = {
        "batch": False,
        "keyword": False,
        "chronologic": False,
        "severity": False,
        "highlights": False,
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
    # Order of checks per token (do not reorder): batch, keyword, chronologic, range, line, cell, participant
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
        # Severity ordering: literal "severity"
        if token.lower() == "severity":
            result["severity"] = True
            continue
        # Highlights keyword: literal "highlights"
        if token.lower() == "highlights":
            result["highlights"] = True
            continue
        # Chronologic keyword: literal "chronologic"
        if token.lower() == "chronologic":
            result["chronologic"] = True
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


def detect_mode_from_input(input_string: str) -> tuple[str | None, dict[str, Any]]:
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
    header_row: list[str], id_cell: Any, participant_id: str
) -> int | None:
    """Find the column index for a given participant ID.

    Args:
        header_row: List of header cell values
        id_cell: The ID header cell object
        participant_id: Participant ID to find (e.g., "P01")

    Returns:
        Column index (0-based) if found, None otherwise
    """
    normalized_target = utils.normalize_participant_id(participant_id).lower()
    for col_idx in range(id_cell.col, len(header_row)):
        header_value = utils.normalize_participant_id(header_row[col_idx]).strip()
        if header_value.lower() == normalized_target:
            return col_idx
    return None


def participant_filename_overrides(
    ctx: SheetContext, user_overrides: dict[str, str] | None = None
) -> dict[str, str | None]:
    """Map each participant id to its source-video filename override (or None).

    Two sources, in increasing precedence:

    1. The sheet's Filename row (``ctx.filename_row_idx``) — a per-column
       override; a cell may list several files plus-separated for a multi-video
       participant (``morning.mp4 + afternoon.mp4``).
    2. The user's own overrides, set from the Start overlay's preview rows and
       persisted per user in ``start.json``. They win, because clipgen cannot
       write the sheet back and this is the only fix a user has that does not
       mean leaving the app. Defaults to ``config.FILENAME_OVERRIDES`` (the
       *open* source's map); pass *user_overrides* explicitly to resolve against
       some other source, as the Start overlay's preview route does.

    Returns ``{}`` when there is neither. Used by the studio/transcripts/
    screenspace servers to resolve a participant's source video(s) via
    ``files.resolve_source_video_paths``.
    """
    if user_overrides is None:
        user_overrides = config.FILENAME_OVERRIDES
    overrides: dict[str, str | None] = {}
    participants = get_participant_list(
        ctx.header_row, ctx.id_cell, ctx.num_participants
    )
    if ctx.filename_row_idx is not None:
        row_data = (
            ctx.sheet_data[ctx.filename_row_idx]
            if ctx.filename_row_idx < len(ctx.sheet_data)
            else []
        )
        for p_idx, pid in enumerate(participants):
            col_idx = ctx.id_cell.col + p_idx
            value = row_data[col_idx].strip() if col_idx < len(row_data) else ""
            overrides[pid] = value or None
    for pid in participants:
        user_value = user_overrides.get(pid)
        if user_value:
            overrides[pid] = user_value
    return overrides


def _make_clip_record(
    ctx: SheetContext, row_idx: int, col_idx: int, cell_value: Any
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
    # Lazy import: keeps gspread (and its heavy google.auth/cryptography chain)
    # off the CLI startup path; only spreadsheet parsing needs the Cell type.
    import gspread

    # Excel cells may yield numbers, datetimes, or None; gspread always returns
    # strings. Coerce here so downstream timestamp parsing (which calls .lower())
    # never sees a non-string.
    cell_value_str = "" if cell_value is None else str(cell_value)
    # gspread Cell uses 1-based coordinates; convert from 0-based list indices
    cell = gspread.cell.Cell(row_idx + 1, col_idx + 1, cell_value_str)
    # observation_cell.col is 1-based; convert to 0-based for sheet_data
    desc_col = ctx.observation_cell.col - 1
    category_col = ctx.category_cell.col - 1
    # Header row in sheet_data is 0-based; id_cell.row is 1-based
    participant_row = ctx.id_cell.row - 1
    desc = ""
    if 0 <= desc_col < len(ctx.sheet_data[row_idx]):
        desc = str(ctx.sheet_data[row_idx][desc_col] or "")
    participant = ""
    if 0 <= participant_row < len(ctx.sheet_data) and col_idx < len(
        ctx.sheet_data[participant_row]
    ):
        participant = utils.normalize_participant_id(
            str(ctx.sheet_data[participant_row][col_idx] or "")
        )
    timestamp_baseline = ""
    # Kept nested (not collapsed) to match the severity_cell guard below: row
    # presence and row/col bounds are separate checks in this layer.
    if ctx.baseline_row_idx is not None:  # noqa: SIM102
        if 0 <= ctx.baseline_row_idx < len(ctx.sheet_data) and col_idx < len(
            ctx.sheet_data[ctx.baseline_row_idx]
        ):
            timestamp_baseline = str(
                ctx.sheet_data[ctx.baseline_row_idx][col_idx] or ""
            ).strip()
    category = ""
    if 0 <= category_col < len(ctx.sheet_data[row_idx]):
        category = str(ctx.sheet_data[row_idx][category_col] or "")
    severity = ""
    if ctx.severity_cell is not None:
        severity_col = ctx.severity_cell.col - 1
        if 0 <= severity_col < len(ctx.sheet_data[row_idx]):
            severity = utils.normalize_severity(ctx.sheet_data[row_idx][severity_col])
    result: ClipRecord = {
        "cell": cell,
        "desc": desc,
        "study": ctx.study_name,
        "participant": participant,
        "category": category,
        "severity": severity,
    }
    if timestamp_baseline:
        result["timestamp_baseline"] = timestamp_baseline
    # Same precedence as participant_filename_overrides: the user's Start-overlay
    # override wins over the sheet's Filename row. Listing a participant against
    # one file while cutting their clips from another is the "wrong output, no
    # error" class, so both paths have to agree.
    filename_override = config.FILENAME_OVERRIDES.get(participant, "")
    if (
        not filename_override
        and ctx.filename_row_idx is not None
        and 0 <= ctx.filename_row_idx < len(ctx.sheet_data)
        and col_idx < len(ctx.sheet_data[ctx.filename_row_idx])
    ):
        filename_override = ctx.sheet_data[ctx.filename_row_idx][col_idx].strip()
    if filename_override:
        result["source_filename"] = filename_override
    return result


def get_line_timestamps(ctx: SheetContext, line_index: int) -> list[ClipRecord]:
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
        config.debug_ic(line_index, ctx.num_participants, ctx.study_name)
    utils.debug_print(
        f"Running method get_line_timestamps, starting line index {line_index} (real sheet line {line_index + 1})"
    )

    if line_index < 0 or line_index >= len(ctx.sheet_data):
        if config.DEBUGGING:
            config.debug_ic(line_index, len(ctx.sheet_data))
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
                    config.debug_ic(
                        issue.get("participant"),
                        issue.get("desc"),
                        issue.get("category"),
                    )
                    config.debug_ic(issue)
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
                    f"Actual cell at address {utils.safe_cell_a1(issue['cell'].row, issue['cell'].col)}"
                )
                clips.append(issue)
                display_value = value.replace("\n", " ")
                cell_addr = utils.safe_cell_a1(issue["cell"].row, issue["cell"].col)
                utils.verbose_print(
                    f"+ Found timestamp: {display_value} at address {cell_addr}"
                )
    except IndexError as e:
        if config.DEBUGGING:
            config.debug_ic(e, line_index)
        utils.error_print(
            f"Index error while reading row {line_index + 1}: {e}",
            ["The spreadsheet structure may be malformed."],
        )

    utils.debug_print(
        f"Line completed, returning list of {len(clips)} potential clips."
    )
    if config.DEBUGGING:
        config.debug_ic(clips)
    return clips


# ---- Sheet-wide collectors ----
#
# Discovery helpers that scan the whole sheet to enumerate the unique values of
# a single dimension (categories, severities, annotation IDs). Used by
# interactive prompts to populate selection menus before any clip records are
# generated; not called from the generate_* functions below.


def collect_categories(ctx: SheetContext) -> list[str]:
    """Scan sheet and return unique categories in order of first appearance."""
    categories = []
    category_col = ctx.category_cell.col - 1
    for i in range(ctx.first_data_row_idx, len(ctx.sheet_data)):
        if category_col < len(ctx.sheet_data[i]):
            category = ctx.sheet_data[i][category_col].strip()
            if category and category not in categories:
                categories.append(category)
    return categories


def collect_severities(ctx: SheetContext) -> tuple[list[str], dict[str, int]]:
    """Scan sheet and return unique severity values sorted most severe first, plus row counts."""
    if ctx.severity_cell is None:
        return [], {}
    severities: list[str] = []
    counts: dict[str, int] = {}
    severity_col = ctx.severity_cell.col - 1
    for i in range(ctx.first_data_row_idx, len(ctx.sheet_data)):
        if severity_col < len(ctx.sheet_data[i]):
            raw = ctx.sheet_data[i][severity_col].strip()
            if raw:
                normalized = utils.normalize_severity(raw)
                if normalized:
                    counts[normalized] = counts.get(normalized, 0) + 1
                    if normalized not in severities:
                        severities.append(normalized)
    severities.sort(key=utils.severity_sort_key)
    return severities, counts


def collect_annotations(ctx: SheetContext) -> tuple[list[str], dict[str, int]]:
    """Scan sheet and return unique annotation IDs found in timestamp cells, plus cell counts."""
    annotation_ids: list[str] = []
    counts: dict[str, int] = {}
    for i in range(ctx.first_data_row_idx, len(ctx.sheet_data)):
        if ctx.filename_row_idx is not None and i == ctx.filename_row_idx:
            continue
        for col_idx in range(ctx.id_cell.col, ctx.id_cell.col + ctx.num_participants):
            if col_idx >= len(ctx.sheet_data[i]):
                continue
            cell_value = ctx.sheet_data[i][col_idx].strip()
            if not cell_value:
                continue
            _, segment_annotations, _ = utils.parse_cell_annotations(cell_value)
            for aid in segment_annotations:
                counts[aid] = counts.get(aid, 0) + 1
                if aid not in annotation_ids:
                    annotation_ids.append(aid)
    annotation_ids.sort()
    return annotation_ids, counts


# ---- Mode generators ----


def _validate_row_range(
    start_line: int, end_line: int, max_row: int
) -> tuple[int, int] | None:
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
    ctx: SheetContext | None = None,
    line_numbers: list[int] | None = None,
    range_start: int | None = None,
    range_end: int | None = None,
    skip_prompts: bool = False,
    cell_specs: list[tuple[str, int]] | None = None,
    participant_id: str | None = None,
    reel_input: str | None = None,
    categories: list[str] | None = None,
    severities: list[str] | None = None,
    annotation_ids: list[str] | None = None,
) -> list[ClipRecord]:
    """Generate clip records from a sheet based on mode and resolved parameters.

    Pure: takes resolved parameters and never prompts — interactive.py owns the
    prompts. *mode* is one of 'batch', 'line', 'range', 'category', 'cell',
    'participant', 'keyword', 'reel', and each mode reads its own parameter
    (*line_numbers*, *range_start*/*range_end*, *cell_specs* as
    ``(participant_id, row_number)`` tuples, *participant_id* as a comma/plus
    separated string, *reel_input*, *categories*).

    Passing a pre-built *ctx* skips the sheet API call. *skip_prompts* (the CLI
    ``--no-input`` flag) drops the batch/keyword confirmations.
    """
    if config.DEBUGGING:
        config.debug_ic(mode, line_numbers, range_start, range_end)

    if ctx is None:
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
        utils.standard_print("Keyword mode: generating annotated clips...")
        return generate_keyword_timestamps(ctx, annotation_ids=annotation_ids)

    if mode == "severity":
        if not severities:
            utils.error_print(
                "Severity mode requires severities list.",
                ["Pass severities via CLI (-S) or use interactive mode."],
            )
            return []
        utils.standard_print(f"Severity mode: filtering by {', '.join(severities)}...")
        return generate_severity_timestamps(ctx, severities)

    if mode == "reel":
        if reel_input is None or not reel_input.strip():
            utils.info_print("Reel mode: no input provided.")
            return []
        utils.standard_print("Reel mode: parsing selectors and collecting clips...")
        return generate_reel_timestamps(ctx, reel_input.strip())

    return []


def generate_batch_timestamps(ctx: SheetContext) -> list[ClipRecord]:
    """Generate clip records for all rows in batch mode."""
    utils.debug_print("Running method generate_batch_timestamps()")
    clips = []
    for i in range(ctx.first_data_row_idx, len(ctx.sheet_data)):
        if ctx.filename_row_idx is not None and i == ctx.filename_row_idx:
            continue
        utils.debug_print(f"Batching on line {i} (real sheet line {i + 1})")
        clips.extend(get_line_timestamps(ctx, i))
    return clips


def generate_keyword_timestamps(
    ctx: SheetContext, annotation_ids: list[str] | None = None
) -> list[ClipRecord]:
    """Generate annotation-marked clips from the entire sheet based on cell content.

    Args:
        ctx: Sheet context with all metadata
        annotation_ids: Annotation IDs to filter by. None means all known annotations.
    """
    clips = generate_batch_timestamps(ctx)
    if not clips:
        return []

    target_ids = (
        annotation_ids
        if annotation_ids
        else list(utils.get_known_annotation_map().values())
    )

    filtered_clips: list[ClipRecord] = []
    for clip in clips:
        _, segment_annotations, _ = utils.parse_cell_annotations(clip["cell"].value)
        matched_indexes: set = set()
        for aid in target_ids:
            matched_indexes.update(segment_annotations.get(aid, set()))
        if matched_indexes:
            clip["selected_segment_indexes"] = sorted(matched_indexes)
            filtered_clips.append(clip)

    return filtered_clips


def generate_category_timestamps(
    ctx: SheetContext, selected_categories: list[str]
) -> list[ClipRecord]:
    """Generate clip records for all rows matching any of the selected categories."""
    utils.debug_print("Starting method generate_category_timestamps()")
    clips = []
    category_col = ctx.category_cell.col - 1
    for i in range(ctx.first_data_row_idx, len(ctx.sheet_data)):
        if ctx.filename_row_idx is not None and i == ctx.filename_row_idx:
            continue
        if category_col >= len(ctx.sheet_data[i]):
            continue
        row_category = ctx.sheet_data[i][category_col].strip()
        if row_category in selected_categories:
            utils.debug_print(f"Row {i + 1} matches category '{row_category}'")
            clips.extend(get_line_timestamps(ctx, i))
    return clips


def generate_severity_timestamps(
    ctx: SheetContext, selected_severities: list[str]
) -> list[ClipRecord]:
    """Generate clip records for all rows matching any of the selected severities."""
    if ctx.severity_cell is None:
        utils.warning_print("No Severity column found in the spreadsheet.")
        return []
    clips = []
    severity_col = ctx.severity_cell.col - 1
    selected_lower = {s.lower() for s in selected_severities}
    for i in range(ctx.first_data_row_idx, len(ctx.sheet_data)):
        if ctx.filename_row_idx is not None and i == ctx.filename_row_idx:
            continue
        if severity_col < len(ctx.sheet_data[i]):
            raw = ctx.sheet_data[i][severity_col].strip()
            normalized = utils.normalize_severity(raw)
            if normalized.lower() in selected_lower:
                clips.extend(get_line_timestamps(ctx, i))
    return clips


def generate_line_timestamps(
    ctx: SheetContext, line_numbers: list[int]
) -> list[ClipRecord]:
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
) -> list[ClipRecord]:
    """Generate clip records for a range of rows.

    Args:
        ctx: Sheet context with all metadata
        start_line: Starting row number (1-based)
        end_line: Ending row number (1-based, inclusive)
    """
    clips = []
    for i in range(start_line - 1, end_line):
        # Skip the Filename override row like every other generator; its cells
        # hold per-column source-video names, not timestamps.
        if ctx.filename_row_idx is not None and i == ctx.filename_row_idx:
            continue
        utils.debug_print(f"Batching on line {i}")
        clips.extend(get_line_timestamps(ctx, i))
    return clips


def generate_participant_timestamps(
    ctx: SheetContext, participant_id: str
) -> list[ClipRecord]:
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
        cell_addr = utils.safe_cell_a1(issue["cell"].row, issue["cell"].col)
        utils.verbose_print(
            f"+ Found timestamp: {display_value} at row {row_idx + 1} ({cell_addr})"
        )
    return clips


def generate_cell_timestamps(
    ctx: SheetContext, cell_specs: list[tuple[str, int]]
) -> list[ClipRecord]:
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
        cell_addr = utils.safe_cell_a1(issue["cell"].row, issue["cell"].col)
        utils.verbose_print(
            f"+ Found timestamp: {display_value} at cell {participant_id}.{row_number} ({cell_addr})"
        )

    return clips


# ---- Clip sorting ----


def sort_clips_chronologically(clips: list[ClipRecord]) -> None:
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


def sort_clips_by_severity(clips: list[ClipRecord]) -> None:
    """Sort clip records in-place by severity (most severe first).

    Clips without severity or with unrecognized severity sort last.
    """
    clips.sort(key=lambda clip: utils.severity_sort_key(clip.get("severity", "")))


# ---- Highlights reel scoring ----

# Largest friction magnitude in the severity scale (critical = -4 → 4), used to
# normalize highlight severity scores into 0-1. Derived from config so the scale
# can't drift from the canonical severity labels.
_MAX_SEVERITY_MAGNITUDE = -min(config.SEVERITY_LABEL_TO_NUMERIC.values())


def _clip_duration_seconds(clip: Any) -> float:
    """Estimate a clip's total duration in seconds from its raw cell value."""
    cell_value = str(clip.get("cell").value if clip.get("cell") is not None else "")
    cleaned_value, _, _ = utils.parse_cell_annotations(cell_value)
    parsed_times = utils.parse_timestamps(cleaned_value)
    if not parsed_times:
        return float(config.DEFAULT_DURATION_SECONDS)
    total = 0.0
    for start_str, end_str in parsed_times:
        s = utils.timestamp_to_seconds(start_str)
        e = utils.timestamp_to_seconds(end_str)
        if s is not None and e is not None:
            total += max(0.0, e - s)
    return total if total > 0 else float(config.DEFAULT_DURATION_SECONDS)


def _clip_highlight_score(clip: Any, lowered_filenames: Sequence[str]) -> float:
    """Score a clip for highlights reel selection.

    Higher score = more important. Combines severity, uniqueness (no existing
    artifact), and keyword annotation, weighted by config constants.
    ``lowered_filenames`` must already be lowercased by the caller.
    """
    # Severity: friction magnitude normalized to 0-1 (positive/neutral labels score 0).
    sev_label = clip.get("severity", "").strip().lower()
    numeric = config.SEVERITY_LABEL_TO_NUMERIC.get(sev_label, 0)
    sev = max(0, -numeric) / _MAX_SEVERITY_MAGNITUDE

    # Uniqueness: 1.0 if no matching artifact exists
    study = (clip.get("study") or "").lower()
    participant = (clip.get("participant") or "").lower()
    desc = (clip.get("desc") or "").lower()[:30]
    has_existing = (
        any(
            study in f and participant in f and desc and desc in f
            for f in lowered_filenames
        )
        if study and participant and desc
        else False
    )
    uniq = 0.0 if has_existing else 1.0

    # Keyword: 1.0 if cell has any annotation (e.g. !key)
    cell_value = str(clip.get("cell").value if clip.get("cell") is not None else "")
    _, _, cell_annotations = utils.parse_cell_annotations(cell_value)
    kw = 1.0 if cell_annotations else 0.0

    return (
        sev * config.HIGHLIGHTS_WEIGHT_SEVERITY
        + uniq * config.HIGHLIGHTS_WEIGHT_UNIQUENESS
        + kw * config.HIGHLIGHTS_WEIGHT_KEYWORD
    )


def score_and_truncate_clips(
    clips: list[ClipRecord],
    existing_filenames: set[str],
    duration_budget: int,
) -> list[ClipRecord]:
    """Score clips by importance and select the best ones that fit within a duration budget.

    Clips are scored by severity, uniqueness, and keyword annotations, then
    sorted highest-first. Clips are accumulated until the total duration exceeds
    the budget.
    """
    lowered = [f.lower() for f in existing_filenames]
    scored = sorted(
        clips,
        key=lambda c: _clip_highlight_score(c, lowered),
        reverse=True,
    )
    result: list[ClipRecord] = []
    total_seconds = 0.0
    for clip in scored:
        dur = _clip_duration_seconds(clip)
        if total_seconds + dur > duration_budget and result:
            break
        result.append(clip)
        total_seconds += dur
    return result


def generate_reel_timestamps(
    ctx: SheetContext, reel_input_string: str
) -> list[ClipRecord]:
    """Generate clip records for reel mode by combining multiple selector types and deduplicating.

    Parses reel input (batch, keyword, chronologic, lines, ranges, categories, cells, participants),
    collects timestamps from each selector, deduplicates by cell (row, col), and returns a single
    ordered list.
    """
    selectors = parse_reel_input(reel_input_string)
    if selectors.get("highlights") and (
        selectors["chronologic"] or selectors.get("severity")
    ):
        utils.error_print(
            "Highlights selector cannot be combined with chronologic or severity ordering.",
            [
                "Use highlights on its own or with other clip selectors (batch, lines, etc.)."
            ],
        )
        return []
    if selectors["chronologic"] and len(selectors["participants"]) != 1:
        utils.error_print(
            "Chronologic selector requires exactly one participant.",
            [
                "Use chronologic with one participant, e.g. 'chronologic, P01'.",
                "Chronologic reels are generated per participant.",
            ],
        )
        return []

    has_any = (
        selectors["batch"]
        or selectors["keyword"]
        or selectors["chronologic"]
        or selectors.get("severity")
        or selectors.get("highlights")
        or selectors["lines"]
        or selectors["ranges"]
        or selectors["categories"]
        or selectors["cells"]
        or selectors["participants"]
    )
    if not has_any:
        return []

    # Highlights alone defaults to batch (score all available clips)
    if selectors.get("highlights") and not any(
        [
            selectors["batch"],
            selectors["keyword"],
            selectors["lines"],
            selectors["ranges"],
            selectors["categories"],
            selectors["cells"],
            selectors["participants"],
        ]
    ):
        selectors["batch"] = True

    all_issues: list[ClipRecord] = []

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
    seen: set[tuple[int, int]] = set()
    deduped: list[ClipRecord] = []
    for issue in all_issues:
        key = (issue["cell"].row, issue["cell"].col)
        if key not in seen:
            seen.add(key)
            deduped.append(issue)

    if selectors["chronologic"]:
        chronologic_participant = utils.normalize_participant_id(
            selectors["participants"][0]
        ).lower()
        deduped = [
            issue
            for issue in deduped
            if utils.normalize_participant_id(issue.get("participant", "")).lower()
            == chronologic_participant
        ]
        sort_clips_chronologically(deduped)
    elif selectors.get("highlights"):
        import files

        existing_filenames = set(files.discover_clips())
        deduped = score_and_truncate_clips(
            deduped, existing_filenames, config.HIGHLIGHTS_REEL_DURATION_SECONDS
        )
    elif selectors.get("severity"):
        sort_clips_by_severity(deduped)
    # Default: preserve insertion order from the selector generators above.
    # For the studio reel button, cells arrive in panel/drag order so the
    # composed reel matches the on-screen card order. Selector-based CLI
    # inputs (batch/categories/lines/keyword) naturally walk the sheet
    # row-major, so they keep producing row-major reels without an explicit
    # sort here.

    return deduped
