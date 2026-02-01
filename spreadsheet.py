# -*- coding: utf-8 -*-
"""Spreadsheet data processing for clipgen."""

import re
import webbrowser
from typing import Any, Dict, List, Optional, Set, Tuple

import gspread
from icecream import ic

import config
import google_api
import utils


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
        utils.error_print(f"Required header(s) not found in spreadsheet: {', '.join(missing_headers)}",
            [f"The spreadsheet must contain columns with these exact headers: {config.ID_HEADER}, {config.OBSERVATION_HEADER}, {config.CATEGORY_HEADER}",
             "Please check your spreadsheet structure."])
        return None
    
    return (id_cell, observation_cell, category_cell)


def _validate_row_range(start_line: int, end_line: int, max_row: int) -> Optional[Tuple[int, int]]:
    """Validate row range; print error and return None if invalid.
    
    Args:
        start_line: Start row (1-based)
        end_line: End row (1-based)
        max_row: Number of rows in sheet (len(sheet_data))
        
    Returns:
        (start_line, end_line) if valid, None otherwise
    """
    if start_line < 1 or end_line < 1:
        utils.error_print(f"Line numbers must be positive. Got start={start_line}, end={end_line}")
        return None
    if start_line > max_row or end_line > max_row:
        utils.error_print(
            f"Line number(s) out of range. Spreadsheet has {max_row} rows.",
            [f"Requested: lines {start_line} to {end_line}"],
        )
        return None
    if start_line > end_line:
        utils.error_print(f"Start line ({start_line}) must be less than or equal to end line ({end_line}).")
        return None
    return (start_line, end_line)


def _interactive_category_selection(categories: List[str]) -> List[str]:
    """Interactive category selection: show numbered list, parse input, validate, confirm. Returns selected category names."""
    utils.info_print('\nAvailable categories:')
    for i, cat in enumerate(categories, 1):
        utils.info_print(f'  {i}. {cat}')
    while True:
        selection = input('\nEnter category numbers (comma-separated, e.g., "1,3,5") or "all":\n>> ')
        if selection.lower() == 'all':
            return categories
        try:
            indices = [int(x.strip()) for x in selection.split(',')]
            selected_categories = []
            invalid_indices = []
            for idx in indices:
                if 1 <= idx <= len(categories):
                    if categories[idx - 1] not in selected_categories:
                        selected_categories.append(categories[idx - 1])
                else:
                    invalid_indices.append(idx)
            if invalid_indices:
                utils.info_print(f'  Invalid index(es): {", ".join(str(i) for i in invalid_indices)}')
            if selected_categories:
                utils.info_print('\nSelected categories:')
                for cat in selected_categories:
                    utils.info_print(f'  - {cat}')
                yn = input('\nIs this correct? y/n\n>> ')
                if yn == 'y':
                    return selected_categories
            else:
                utils.info_print('No valid categories selected. Please try again.')
        except ValueError:
            utils.info_print('Please enter valid numbers separated by commas.')


def _make_clip_record(
    sheet_data: List[List[str]],
    row_idx: int,
    col_idx: int,
    id_cell: Any,
    observation_cell: Any,
    study_name: str,
    cell_value: str,
) -> Dict[str, Any]:
    """Build one clip record dict for a cell at (row_idx, col_idx).
    
    Args:
        sheet_data: Sheet data matrix
        row_idx: 0-based row index
        col_idx: 0-based column index
        id_cell: ID header cell
        observation_cell: Observation header cell
        study_name: Normalized study name
        cell_value: Value of the timestamp cell (sheet_data[row_idx][col_idx])
        
    Returns:
        Dict with keys: cell, desc, study, participant, category
    """
    cell = gspread.cell.Cell(row_idx + 1, col_idx + 1, cell_value)
    desc_col = observation_cell.col - 1
    category_col = observation_cell.col - 2
    participant_row = id_cell.row - 1
    desc = ""
    if 0 <= desc_col < len(sheet_data[row_idx]):
        desc = sheet_data[row_idx][desc_col]
    participant = ""
    if 0 <= participant_row < len(sheet_data) and col_idx < len(sheet_data[participant_row]):
        participant = sheet_data[participant_row][col_idx]
    category = ""
    if 0 <= category_col < len(sheet_data[row_idx]):
        category = sheet_data[row_idx][category_col]
    return {
        "cell": cell,
        "desc": desc,
        "study": study_name,
        "participant": participant,
        "category": category,
    }


def generate_list(sheet: Any, mode: str, line_numbers: Optional[List[int]] = None, range_start: Optional[int] = None, range_end: Optional[int] = None, skip_prompts: bool = False, cell_specs: Optional[List[Tuple[str, int]]] = None, participant_id: Optional[str] = None, reel_input: Optional[str] = None) -> List[Any]:
    """Goes through a sheet, bundles values from timestamp columns and descriptions columns into tuples.
    
    Args:
        sheet: The gspread worksheet object
        mode: One of 'batch', 'line', 'range', 'category', 'cell', 'participant', 'reel', 'select'
        line_numbers: Optional list of line numbers for 'line' mode (CLI)
        range_start: Optional start line for 'range' mode (CLI)
        range_end: Optional end line for 'range' mode (CLI)
        skip_prompts: If True, skip confirmation prompts (CLI -y flag)
        cell_specs: Optional list of (participant_id, row_number) tuples for 'cell' mode (CLI)
        participant_id: Optional participant ID for 'participant' mode (CLI)
        reel_input: Optional reel selector string for 'reel' mode (e.g. '11, 13-16, P01, "Observations"')
        
    Returns:
        List of clip records (dicts with keys: cell, desc, study, participant, category; 'times' added later by prepare_clip).
    """
    if config.DEBUGGING:
        ic(mode, line_numbers, range_start, range_end)
    # Validate required headers exist
    header_result = validate_spreadsheet_headers(sheet)
    if header_result is None:
        return []
    
    id_cell, observation_cell, category_cell = header_result
    if config.DEBUGGING:
        ic(id_cell, observation_cell, category_cell)
    clips = []

    # Sheet data is a list of lists, which forms a matrix
    # - sheet_data[row][col] where indices start at 0 (real spreadsheet starts at 1)
    sheet_data = sheet.get_all_values()
    utils.debug_print(f'Sheet dumped into memory at {utils.get_current_time()}')
    
    # Check if sheet is empty or has only headers
    if len(sheet_data) <= 1:
        utils.error_print("Spreadsheet appears to be empty (no data rows found).",
            [f"The spreadsheet only has {len(sheet_data)} row(s)."])
        return []

    # Determine the study name.
    study_name = sheet_data[0][0]
    if study_name == '':
        study_name = sheet.spreadsheet.title
    utils.verbose_print(f'\nBeginning work on {study_name}.')

    # Normalize study name for filesystem use
    study_name = utils.normalize_study_name(study_name)

    # Get number of participants needed to loop through the worksheet
    num_participants = get_num_participants(sheet.row_values(id_cell.row), id_cell, sheet.col_count)
    
    # Warn if no participants found
    if num_participants == 0:
        utils.warning_print("No participant columns found in the spreadsheet.",
            [f"Looking for columns starting with: {', '.join(config.PARTICIPANT_PREFIXES)}",
             "Check that participant column headers start with 'P' or 'G' (e.g., P01, P02, G01)."])
        return []

    # Generate clips according to the selected mode.
    # --- Batch mode ---
    if mode == 'batch':
        if skip_prompts:
            utils.verbose_print('Batch mode: generating all possible clips...')
            clips = generate_batch_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name)
        else:
            yn = input('\nWarning: This will generate all possible clips. Do you want to proceed? y/n\n>> ')
            if yn == 'y':
                clips = generate_batch_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name)
    # --- Category mode ---
    elif mode == 'category':
        categories = collect_categories(sheet_data, id_cell, category_cell)
        if not categories:
            utils.info_print('\nNo categories found in the spreadsheet.')
            return []
        selected_categories = _interactive_category_selection(categories)
        clips = generate_category_timestamps(sheet_data, id_cell, observation_cell, category_cell, num_participants, study_name, selected_categories)
    # --- Line mode ---
    elif mode == 'line':
        clips = generate_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name, line_numbers, skip_prompts)
    # --- Range mode ---
    elif mode == 'range':
        max_row = len(sheet_data)
        if range_start is not None and range_end is not None:
            # CLI mode - use provided range with bounds validation
            valid = _validate_row_range(range_start, range_end, max_row)
            if valid is None:
                return []
            range_start, range_end = valid
            utils.verbose_print(f'Range mode: lines {range_start} to {range_end}')
            utils.verbose_print(f'Lines selected: {sheet_data[range_start-1][observation_cell.col-1]} to {sheet_data[range_end-1][observation_cell.col-1]}')
            clips = generate_range_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name, range_start, range_end)
        else:
            # Interactive mode
            while True:
                try:
                    start_line = int(input('\nWhich starting line (row number only)?\n>> '))
                    end_line = int(input('\nWhich ending line (row number only)?\n>> '))
                except ValueError:
                    utils.info_print('\nInvalid input. Please enter row numbers as integers.')
                    continue
                if _validate_row_range(start_line, end_line, max_row) is None:
                    continue
                utils.info_print(f'Lines selected: {sheet_data[start_line-1][observation_cell.col-1]} to {sheet_data[end_line-1][observation_cell.col-1]}')
                yn = input('Is this correct? y/n\n>> ')
                if yn == 'y':
                    break
            clips = generate_range_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name, start_line, end_line)
    # --- Cell mode ---
    elif mode == 'cell':
        if cell_specs is not None:
            # CLI mode - use provided cell specifications
            utils.verbose_print(f'Cell mode: processing {len(cell_specs)} cell(s)')
            clips = generate_cell_timestamps(sheet_data, id_cell, observation_cell, study_name, cell_specs)
        else:
            # Interactive mode
            while True:
                try:
                    cell_input = input('\nEnter cell specification(s) (e.g., P01.11 or P01.11 + P03.11):\n>> ')
                    if not cell_input.strip():
                        utils.info_print('Please enter at least one cell specification.')
                        continue
                    
                    # Parse cell specifications
                    try:
                        parsed_specs = parse_cell_specifications(cell_input)
                    except ValueError as e:
                        utils.info_print(f'Invalid format: {e}')
                        utils.info_print('Expected format: P01.11 or P01.11 + P03.11')
                        continue
                    
                    # Preview selected cells
                    utils.info_print('\nSelected cells:')
                    header_row = sheet_data[id_cell.row - 1] if id_cell.row > 0 else []
                    valid_specs = []
                    for participant_id, row_number in parsed_specs:
                        col_idx = find_participant_column(header_row, id_cell, participant_id)
                        if col_idx is None:
                            utils.info_print(f'  {participant_id}.{row_number}: [INVALID - participant not found]')
                        elif row_number < 1 or row_number > len(sheet_data):
                            utils.info_print(f'  {participant_id}.{row_number}: [INVALID - row out of range]')
                        else:
                            row_idx = row_number - 1
                            cell_value = ''
                            if col_idx < len(sheet_data[row_idx]):
                                cell_value = sheet_data[row_idx][col_idx]
                            desc = ''
                            desc_col = observation_cell.col - 1
                            if desc_col >= 0 and desc_col < len(sheet_data[row_idx]):
                                desc = sheet_data[row_idx][desc_col]
                            if cell_value and cell_value.strip():
                                utils.info_print(f'  {participant_id}.{row_number}: {cell_value.replace(chr(10), " ")} (row: {desc[:config.DESCRIPTION_PREVIEW_LENGTH] if desc else "N/A"})')
                            else:
                                utils.info_print(f'  {participant_id}.{row_number}: [EMPTY]')
                            valid_specs.append((participant_id, row_number))
                    
                    if not valid_specs:
                        utils.info_print('\nNo valid cells found. Please try again.')
                        continue
                    
                    utils.info_print('')
                    yn = input('Are these the correct cells? y/n\n>> ')
                    if yn == 'y':
                        clips = generate_cell_timestamps(sheet_data, id_cell, observation_cell, study_name, valid_specs)
                        break
                except KeyboardInterrupt:
                    utils.info_print('\nCancelled by user.')
                    return []
    # --- Participant mode ---
    elif mode == 'participant':
        header_row = sheet_data[id_cell.row - 1] if id_cell.row > 0 else []
        available_list = get_participant_list(header_row, id_cell, num_participants)
        if participant_id is not None:
            # CLI mode - parse and validate participant ID(s)
            participant_ids = parse_participant_selection(participant_id)
            if not participant_ids:
                utils.error_print("No participant ID(s) provided.",
                    ['Use format: P01 or P01+P03 or P01, P03'])
                return []
            invalid = []
            for pid in participant_ids:
                if find_participant_column(header_row, id_cell, pid) is None:
                    invalid.append(pid)
            if invalid:
                utils.error_print(f"Participant(s) not found in spreadsheet headers: {', '.join(invalid)}",
                    [f"Available participants: {', '.join(available_list)}"])
                return []
            utils.verbose_print(f'Participant mode: generating all clips for {", ".join(participant_ids)}')
            clips = []
            for pid in participant_ids:
                clips.extend(generate_participant_timestamps(sheet_data, id_cell, observation_cell, study_name, pid))
        else:
            # Interactive mode
            if not available_list:
                utils.info_print('\nNo participants found in the spreadsheet.')
                return []
            utils.info_print('\nAvailable participants:')
            for i, pid in enumerate(available_list, 1):
                utils.info_print(f'  {i}. {pid}')
            while True:
                selection = input('\nEnter participant number(s) or ID(s), separated by + or , (e.g., 1, 3 or P01, P03):\n>> ').strip()
                if not selection:
                    utils.info_print('Please enter one or more participant numbers or IDs.')
                    continue
                tokens = parse_participant_selection(selection)
                if not tokens:
                    utils.info_print('No valid participant(s) entered. Use + or , as separator.')
                    continue
                chosen_ids = []
                invalid_tokens = []
                for token in tokens:
                    if token.isdigit():
                        idx = int(token)
                        if 1 <= idx <= len(available_list):
                            chosen_ids.append(available_list[idx - 1])
                        else:
                            invalid_tokens.append(token)
                    else:
                        col_idx = find_participant_column(header_row, id_cell, token)
                        if col_idx is not None:
                            chosen_ids.append(header_row[col_idx] if col_idx < len(header_row) else token)
                        else:
                            invalid_tokens.append(token)
                if invalid_tokens:
                    utils.info_print(f"Not found: {', '.join(invalid_tokens)}. Available: {', '.join(available_list)}")
                    continue
                # Dedupe preserving order
                seen = set()
                unique_ids = []
                for pid in chosen_ids:
                    if pid not in seen:
                        seen.add(pid)
                        unique_ids.append(pid)
                utils.info_print(f'\nSelected participant(s): {", ".join(unique_ids)}')
                yn = input('Generate all clips for these participants? y/n\n>> ')
                if yn == 'y':
                    clips = []
                    for pid in unique_ids:
                        clips.extend(generate_participant_timestamps(sheet_data, id_cell, observation_cell, study_name, pid))
                    break
    # --- Reel mode ---
    elif mode == 'reel':
        if reel_input is None or not reel_input.strip():
            utils.info_print('\nReel mode: no input provided.')
            return []
        utils.verbose_print('Reel mode: parsing selectors and collecting clips...')
        clips = generate_reel_timestamps(
            sheet_data,
            id_cell,
            observation_cell,
            category_cell,
            num_participants,
            study_name,
            reel_input.strip(),
        )
    elif mode == 'select':
        pass

    return clips

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
    num_participants = 0
    for j in range(0, col_count):
        if len(header_row[j]) > 0:
            if header_row[j][0] in config.PARTICIPANT_PREFIXES:
                num_participants += 1
    utils.verbose_print(f'Found {num_participants} participants in total, spanning columns {id_cell.col+1} to {num_participants+id_cell.col+1}.')
    return num_participants

def get_participant_list(header_row: List[str], id_cell: Any, num_participants: int) -> List[str]:
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
        if j < len(header_row) and header_row[j].strip():
            participants.append(header_row[j].strip())
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
    combined = input_str.replace(',', '+')
    return [s.strip() for s in combined.split('+') if s.strip()]

def parse_cell_specifications(cell_input: str) -> List[Tuple[str, int]]:
    """Parse cell specification string into list of (participant_id, row_number) tuples.
    
    Args:
        cell_input: String like "P01.11" or "P01.11 + P03.11 + P03.09"
        
    Returns:
        List of (participant_id, row_number) tuples
        
    Raises:
        ValueError: If format is invalid
    """
    # Support both + and , as separators
    cell_str = cell_input.replace(',', '+')
    specs = []
    
    for spec in cell_str.split('+'):
        spec = spec.strip()
        if not spec:
            continue
            
        # Split by dot to get participant_id and row_number
        if '.' not in spec:
            raise ValueError(f'Invalid cell specification "{spec}". Expected format: P01.11')
        
        parts = spec.split('.', 1)
        if len(parts) != 2:
            raise ValueError(f'Invalid cell specification "{spec}". Expected format: P01.11')
        
        participant_id = parts[0].strip()
        row_str = parts[1].strip()
        
        # Validate participant ID format (should start with P or G)
        if not participant_id or participant_id[0] not in config.PARTICIPANT_PREFIXES:
            raise ValueError(f'Invalid participant ID "{participant_id}". Must start with {", ".join(config.PARTICIPANT_PREFIXES)}')
        
        # Validate row number
        try:
            row_number = int(row_str)
            if row_number < 1:
                raise ValueError(f'Row number must be positive. Got: {row_number}')
        except ValueError as e:
            if 'invalid literal' in str(e):
                raise ValueError(f'Invalid row number "{row_str}". Must be a positive integer.')
            raise
        
        specs.append((participant_id, row_number))
    
    return specs


def parse_reel_input(input_string: str) -> dict:
    """Parse mixed reel selector input into structured selectors.

    Supports: batch, lines (e.g. 11, 12), ranges (e.g. 13-16), categories (quoted),
    cells (e.g. P01.11), participants (e.g. P01, P02).

    Args:
        input_string: Raw user input (e.g. '11, 13-16, P01, P02.15, "Observations"')

    Returns:
        Dict with keys: batch (bool), lines (list of int), ranges (list of (int,int)),
        categories (list of str), cells (list of (str,int)), participants (list of str)
    """
    import re
    result = {
        'batch': False,
        'lines': [],
        'ranges': [],
        'categories': [],
        'cells': [],
        'participants': [],
    }
    if not input_string or not input_string.strip():
        return result

    # Extract quoted strings (categories) first so commas inside don't split
    rest = input_string.strip()
    while True:
        match = re.search(r'"([^"]*)"', rest)
        if not match:
            break
        result['categories'].append(match.group(1).strip())
        rest = rest[:match.start()] + ' ' + rest[match.end():]

    # Split remaining by comma, then by space for tokens that might be space-separated
    # Order of checks per token (do not reorder): batch, range, line, cell, participant
    parts = [p.strip() for p in rest.split(',') if p.strip()]
    seen_lines = set()
    seen_ranges = set()
    seen_cells = set()
    seen_participants = set()

    for part in parts:
        token = part.strip()
        if not token:
            continue
        # Batch keyword
        if token.lower() == 'batch':
            result['batch'] = True
            continue
        # Range: digits-digits
        range_match = re.match(r'^(\d+)\s*-\s*(\d+)$', token)
        if range_match:
            start, end = int(range_match.group(1)), int(range_match.group(2))
            if start <= end and (start, end) not in seen_ranges:
                seen_ranges.add((start, end))
                result['ranges'].append((start, end))
            continue
        # Single line number
        if token.isdigit():
            line_num = int(token)
            if line_num not in seen_lines:
                seen_lines.add(line_num)
                result['lines'].append(line_num)
            continue
        # Cell: P##.## or G##.## (participant.row)
        if '.' in token:
            cell_match = re.match(r'^([PG]\w*)\.(\d+)$', token, re.IGNORECASE)
            if cell_match:
                pid, row_str = cell_match.group(1), cell_match.group(2)
                if pid[0].upper() in config.PARTICIPANT_PREFIXES:
                    row_num = int(row_str)
                    if row_num >= 1:
                        key = (pid.upper() if len(pid) <= 3 else pid, row_num)
                        if key not in seen_cells:
                            seen_cells.add(key)
                            result['cells'].append((pid, row_num))
            continue
        # Participant only: P## or G## (no dot, valid participant-like token)
        if token[0].upper() in config.PARTICIPANT_PREFIXES and '.' not in token:
            key = token.upper() if len(token) <= 3 else token
            if key not in seen_participants:
                seen_participants.add(key)
                result['participants'].append(token)

    return result


def find_participant_column(header_row: List[str], id_cell: Any, participant_id: str) -> Optional[int]:
    """Find the column index for a given participant ID.
    
    Args:
        header_row: List of header cell values
        id_cell: The ID header cell object
        participant_id: Participant ID to find (e.g., "P01")
        
    Returns:
        Column index (0-based) if found, None otherwise
    """
    for col_idx in range(id_cell.col - 1, len(header_row)):
        header_value = header_row[col_idx].strip()
        if header_value.lower() == participant_id.lower():
            return col_idx
    return None

def generate_participant_timestamps(sheet_data: List[List[str]], id_cell: Any, observation_cell: Any, study_name: str, participant_id: str) -> List[Any]:
    """Generate timestamps for all rows in a single participant's column.

    Args:
        sheet_data: The sheet data matrix
        id_cell: The ID header cell
        observation_cell: The observation header cell
        study_name: Normalized study name
        participant_id: Participant ID (e.g. P01)

    Returns:
        List of clip issue dictionaries
    """
    utils.debug_print('Starting method generate_participant_timestamps()')
    header_row = sheet_data[id_cell.row - 1] if id_cell.row > 0 else []
    col_idx = find_participant_column(header_row, id_cell, participant_id)
    if col_idx is None:
        return []
    clips = []
    for row_idx in range(id_cell.row, len(sheet_data)):
        if col_idx >= len(sheet_data[row_idx]):
            continue
        cell_value = sheet_data[row_idx][col_idx]
        if not cell_value or not cell_value.strip():
            continue
        issue = _make_clip_record(sheet_data, row_idx, col_idx, id_cell, observation_cell, study_name, cell_value)
        clips.append(issue)
        utils.verbose_print(f"+ Found timestamp: {cell_value.replace(chr(10), ' ')} at row {row_idx + 1} ({gspread.utils.rowcol_to_a1(issue['cell'].row, issue['cell'].col)})")
    return clips

def generate_cell_timestamps(sheet_data: List[List[str]], id_cell: Any, observation_cell: Any, study_name: str, cell_specs: List[Tuple[str, int]]) -> List[Any]:
    """Generate timestamps for specific cells.
    
    Args:
        sheet_data: The sheet data matrix
        id_cell: The ID header cell
        observation_cell: The observation header cell
        study_name: Normalized study name
        cell_specs: List of (participant_id, row_number) tuples
        
    Returns:
        List of clip issue dictionaries
    """
    utils.debug_print('Starting method generate_cell_timestamps()')
    clips = []
    header_row = sheet_data[id_cell.row - 1] if id_cell.row > 0 else []
    
    for participant_id, row_number in cell_specs:
        # Find the column for this participant
        col_idx = find_participant_column(header_row, id_cell, participant_id)
        if col_idx is None:
            utils.warning_print(f"Participant '{participant_id}' not found in spreadsheet headers.",
                [f"Available participants start with: {', '.join(config.PARTICIPANT_PREFIXES)}"])
            continue
        
        # Validate row number (convert to 0-based index)
        row_idx = row_number - 1
        if row_idx < 0 or row_idx >= len(sheet_data):
            utils.warning_print(f"Row {row_number} is out of range.",
                [f"Spreadsheet has {len(sheet_data)} rows (valid range: 1-{len(sheet_data)})."])
            continue
        
        # Get the cell value
        if col_idx >= len(sheet_data[row_idx]):
            utils.warning_print(f"Column index {col_idx + 1} is out of range for row {row_number}.")
            continue
        
        cell_value = sheet_data[row_idx][col_idx]
        if not cell_value or cell_value.strip() == '':
            utils.verbose_print(f"Cell {participant_id}.{row_number} is empty, skipping.")
            continue
        
        issue = _make_clip_record(sheet_data, row_idx, col_idx, id_cell, observation_cell, study_name, cell_value)
        # Use actual header value for participant when available
        participant_row = id_cell.row - 1
        if 0 <= participant_row < len(sheet_data) and col_idx < len(sheet_data[participant_row]) and sheet_data[participant_row][col_idx]:
            issue['participant'] = sheet_data[participant_row][col_idx]
        clips.append(issue)
        utils.verbose_print(f"+ Found timestamp: {cell_value.replace(chr(10), ' ')} at cell {participant_id}.{row_number} ({gspread.utils.rowcol_to_a1(issue['cell'].row, issue['cell'].col)})")
    
    return clips

def generate_batch_timestamps(sheet_data: List[List[str]], id_cell: Any, observation_cell: Any, num_participants: int, study_name: str) -> List[Any]:
    """Generate timestamps for all rows in batch mode.
    
    Args:
        sheet_data: The sheet data matrix
        id_cell: The ID header cell
        observation_cell: The observation header cell
        num_participants: Number of participant columns
        study_name: Normalized study name
        
    Returns:
        List of clip issue dictionaries
    """
    utils.debug_print('Running method generate_batch_timestamps()')
    clips = []
    for i in range(id_cell.row+1, len(sheet_data)):
        utils.debug_print(f'Batching on line {i} (real sheet line {i+1})')
        clips.extend(get_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, i, study_name))
    return clips

def collect_categories(sheet_data: List[List[str]], id_cell: Any, category_cell: Any) -> List[str]:
    """Scan sheet and return unique categories in order of first appearance.
    
    Args:
        sheet_data: The sheet data matrix
        id_cell: The ID header cell (unused but kept for consistency)
        category_cell: The category header cell
        
    Returns:
        List of unique category names in order of first appearance
    """
    categories = []
    category_col = category_cell.col - 1  # Convert from 1-indexed to 0-indexed
    
    # Start from the row after the category header
    for i in range(category_cell.row, len(sheet_data)):        
        category = sheet_data[i][category_col].strip()
        if category and category not in categories:
            categories.append(category)
    
    return categories

def generate_category_timestamps(sheet_data: List[List[str]], id_cell: Any, observation_cell: Any, category_cell: Any, num_participants: int, study_name: str, selected_categories: List[str]) -> List[Any]:
    """Generate timestamps for all rows matching any of the selected categories.
    
    Args:
        sheet_data: The sheet data matrix
        id_cell: The ID header cell
        observation_cell: The observation header cell
        category_cell: The category header cell
        num_participants: Number of participant columns
        study_name: Normalized study name
        selected_categories: List of category names to include
        
    Returns:
        List of clip issue dictionaries
    """
    utils.debug_print('Starting method generate_category_timestamps()')
    clips = []
    category_col = category_cell.col - 1  # Convert from 1-indexed to 0-indexed
    
    # Start from the row after the category header
    for i in range(category_cell.row, len(sheet_data)):
        row_category = sheet_data[i][category_col].strip()
        if row_category in selected_categories:
            utils.debug_print(f"Row {i+1} matches category '{row_category}'")
            clips.extend(get_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, i, study_name))
    
    return clips

def generate_line_timestamps(sheet_data: List[List[str]], id_cell: Any, observation_cell: Any, num_participants: int, study_name: str, cli_line_numbers: Optional[List[int]] = None, skip_prompts: bool = False) -> List[Any]:
    """Generate videos for one or more line/row numbers.
    
    Args:
        sheet_data: The sheet data matrix
        id_cell: The ID header cell
        observation_cell: The observation header cell
        num_participants: Number of participant columns
        study_name: Normalized study name
        cli_line_numbers: Optional list of line numbers from CLI (skips interactive input)
        skip_prompts: If True, skip confirmation prompts
        
    Returns:
        List of clip issue dictionaries
    """
    valid_lines = []
    
    if cli_line_numbers is not None:
        # CLI mode - use provided line numbers
        utils.verbose_print(f'\nLine mode: processing lines {", ".join(str(n) for n in cli_line_numbers)}')
        utils.verbose_print('\nSelected issues:')
        for line_num in cli_line_numbers:
            if line_num < 1 or line_num > len(sheet_data):
                utils.verbose_print(f'  Line {line_num}: [INVALID - out of range]')
            else:
                desc = sheet_data[line_num-1][observation_cell.col-1]
                utils.verbose_print(f'  Line {line_num}: {desc}')
                valid_lines.append(line_num)
        
        if not valid_lines:
            utils.info_print('\nNo valid lines found. Exiting.')
            return []
    else:
        # Interactive mode
        while True:
            try:
                line_input = input('\nWhich issue(s)? Enter row number(s), comma-separated for multiple.\n>> ')
                # Parse comma-separated line numbers
                line_numbers = [int(num.strip()) for num in line_input.split(',')]
            except ValueError:
                utils.info_print('\nTry again. Enter row numbers as integers, separated by commas.')
                continue
            
            # Preview all selected lines
            utils.info_print('\nSelected issues:')
            valid_lines = []
            for line_num in line_numbers:
                if line_num < 1 or line_num > len(sheet_data):
                    utils.info_print(f'  Line {line_num}: [INVALID - out of range]')
                else:
                    desc = sheet_data[line_num-1][observation_cell.col-1]
                    utils.info_print(f'  Line {line_num}: {desc}')
                    valid_lines.append(line_num)
            
            if not valid_lines:
                utils.info_print('\nNo valid lines selected. Please try again.')
                continue
            
            utils.info_print('')
            yn = input('Are these the correct issues? y/n\n>> ')
            if yn == 'y':
                break

    # Collect clips from all valid lines
    clips = []
    for line_num in valid_lines:
        utils.debug_print(f'Calling get_line_timestamps() from generate_line_timestamps() for line {line_num}')
        line_clips = get_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, line_num-1, study_name)
        clips.extend(line_clips)
    
    utils.debug_print(f'Printing return of get_line_timestamps() in generate_line_timestamps(): {len(clips)} total clips')
    utils.debug_print(str(clips))

    return clips

def get_line_timestamps(sheet_data: List[List[str]], id_cell: Any, observation_cell: Any, num_participants: int, line_index: int, study_name: str) -> List[Any]:
    """Extract timestamp data from a single row in the spreadsheet.
    
    Processes all participant columns in the specified row and creates
    clip issue dictionaries for each timestamp found.
    
    Args:
        sheet_data: The sheet data matrix
        id_cell: The ID header cell
        observation_cell: The observation header cell
        num_participants: Number of participant columns
        line_index: Zero-based row index to process
        study_name: Normalized study name
        
    Returns:
        List of clip issue dictionaries, one per timestamp found
    """
    if config.DEBUGGING:
        ic(line_index, num_participants, study_name)
    utils.debug_print(f'Running method get_line_timestamps, starting line index {line_index} (real sheet line {line_index+1})')
    
    # Bounds checking
    if line_index < 0 or line_index >= len(sheet_data):
        if config.DEBUGGING:
            ic(line_index, len(sheet_data))
        utils.error_print(f"Line index {line_index} (row {line_index+1}) is out of bounds.",
            [f"Spreadsheet has {len(sheet_data)} rows."])
        return []

    clips = []
    try:
        for col_index, value in enumerate(sheet_data[line_index]):
            utils.debug_print(f"Item {col_index} with value '{value}' being processed.")
            if col_index < id_cell.col:
                utils.debug_print(f"Skipping item {col_index} with value '{value}'")
            elif col_index == id_cell.col + num_participants:
                utils.debug_print(f'Exit for-loop, reached final column {col_index} (real sheet column {col_index+1}).')
                break
            elif value is None or value == '':
                pass
            else:
                # Build one clip record per non-empty participant cell in this row
                issue = _make_clip_record(sheet_data, line_index, col_index, id_cell, observation_cell, study_name, value)
                if config.DEBUGGING:
                    ic(issue.get("participant"), issue.get("desc"), issue.get("category"))
                    ic(issue)
                utils.debug_print(f"Participant ID at R{id_cell.row},C{col_index} -> '{issue.get('participant')}'")
                utils.debug_print(f"Description at R{line_index},C{observation_cell.col-1} -> '{issue.get('desc')}'")
                utils.debug_print(f"Timestamp at R{issue['cell'].row-1},C{issue['cell'].col-1} -> '{issue['cell'].value}'")
                utils.debug_print(f'Actual cell at address {gspread.utils.rowcol_to_a1(issue["cell"].row, issue["cell"].col)}')
                clips.append(issue)
                utils.verbose_print(f"+ Found timestamp: {value.replace(chr(10), ' ')} at address {gspread.utils.rowcol_to_a1(issue['cell'].row, issue['cell'].col)}")
    except IndexError as e:
        if config.DEBUGGING:
            ic(e, line_index)
        utils.error_print(f"Index error while reading row {line_index+1}: {e}",
            ["The spreadsheet structure may be malformed."])

    utils.debug_print(f'Line completed, returning list of {len(clips)} potential clips.')
    if config.DEBUGGING:
        ic(clips)
    return clips

def generate_range_timestamps(sheet_data: List[List[str]], id_cell: Any, observation_cell: Any, num_participants: int, study_name: str, start_line: int, end_line: int) -> List[Any]:
    """Generate timestamps for a range of rows.
    
    Args:
        sheet_data: The sheet data matrix
        id_cell: The ID header cell
        observation_cell: The observation header cell
        num_participants: Number of participant columns
        study_name: Normalized study name
        start_line: Starting row number (1-based)
        end_line: Ending row number (1-based, inclusive)
        
    Returns:
        List of clip issue dictionaries
    """
    clips = []
    for i in range(start_line-1, end_line):
        utils.debug_print(f'Batching on line {i}')
        clips.extend(get_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, i, study_name))
    return clips


def generate_reel_timestamps(
    sheet_data: List[List[str]],
    id_cell: Any,
    observation_cell: Any,
    category_cell: Any,
    num_participants: int,
    study_name: str,
    reel_input_string: str,
) -> List[Any]:
    """Generate timestamps for reel mode by combining multiple selector types and deduplicating.

    Parses reel input (batch, lines, ranges, categories, cells, participants), collects
    timestamps from each selector, deduplicates by cell (row, col), and returns a single
    ordered list of issue dicts.

    Args:
        sheet_data: The sheet data matrix
        id_cell: The ID header cell
        observation_cell: The observation header cell
        category_cell: The category header cell
        num_participants: Number of participant columns
        study_name: Normalized study name
        reel_input_string: Raw user input (e.g. '11, 13-16, P01, "Observations"')

    Returns:
        List of clip issue dictionaries, deduplicated by cell and sorted by row then column
    """
    selectors = parse_reel_input(reel_input_string)
    has_any = (
        selectors['batch']
        or selectors['lines']
        or selectors['ranges']
        or selectors['categories']
        or selectors['cells']
        or selectors['participants']
    )
    if not has_any:
        return []

    all_issues: List[Any] = []

    if selectors['batch']:
        all_issues.extend(
            generate_batch_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name)
        )

    if selectors['lines']:
        all_issues.extend(
            generate_line_timestamps(
                sheet_data,
                id_cell,
                observation_cell,
                num_participants,
                study_name,
                cli_line_numbers=selectors['lines'],
                skip_prompts=True,
            )
        )

    for start_line, end_line in selectors['ranges']:
        all_issues.extend(
            generate_range_timestamps(
                sheet_data,
                id_cell,
                observation_cell,
                num_participants,
                study_name,
                start_line,
                end_line,
            )
        )

    if selectors['categories']:
        all_issues.extend(
            generate_category_timestamps(
                sheet_data,
                id_cell,
                observation_cell,
                category_cell,
                num_participants,
                study_name,
                selectors['categories'],
            )
        )

    if selectors['cells']:
        all_issues.extend(
            generate_cell_timestamps(
                sheet_data,
                id_cell,
                observation_cell,
                study_name,
                selectors['cells'],
            )
        )

    for participant_id in selectors['participants']:
        all_issues.extend(
            generate_participant_timestamps(
                sheet_data,
                id_cell,
                observation_cell,
                study_name,
                participant_id,
            )
        )

    # Deduplicate by cell (row, col) so each timestamp is included once
    seen: Set[Tuple[int, int]] = set()
    deduped: List[Any] = []
    for issue in all_issues:
        key = (issue['cell'].row, issue['cell'].col)
        if key not in seen:
            seen.add(key)
            deduped.append(issue)

    # Sort by row then column for logical playback order
    deduped.sort(key=lambda issue: (issue['cell'].row, issue['cell'].col))

    return deduped


def browse_spreadsheet(sheet: Any) -> None:
    """Interactive browse mode for viewing spreadsheet rows line by line.
    
    Allows users to navigate through the spreadsheet to inspect issues
    before generating clips. Shows row number, category, description,
    and participant/group timestamps for each row.
    """
    # Validate required headers exist
    header_result = validate_spreadsheet_headers(sheet)
    if header_result is None:
        return
    
    id_cell, observation_cell, category_cell = header_result
    
    # Load sheet data
    sheet_data = sheet.get_all_values()
    utils.debug_print(f'Sheet dumped into memory at {utils.get_current_time()}')
    
    # Check if sheet is empty or has only headers
    if len(sheet_data) <= 1:
        utils.error_print("Spreadsheet appears to be empty (no data rows found).")
        return
    
    # Get participant info
    header_row = sheet.row_values(id_cell.row)
    num_participants = get_num_participants(header_row, id_cell, sheet.col_count)
    
    if num_participants == 0:
        utils.warning_print("No participant columns found in the spreadsheet.",
            [f"Looking for columns starting with: {', '.join(config.PARTICIPANT_PREFIXES)}"])
        return
    
    # Calculate bounds for data rows (after header)
    first_data_row = id_cell.row  # 0-indexed, this is the first row after the header
    last_data_row = len(sheet_data) - 1  # 0-indexed
    total_data_rows = last_data_row - first_data_row + 1
    
    # Current position (0-indexed into sheet_data)
    current_row = first_data_row
    
    # Get participant column headers for display
    participant_headers = []
    for col_idx in range(id_cell.col, id_cell.col + num_participants):
        if col_idx < len(header_row):
            participant_headers.append(header_row[col_idx])
    
    utils.info_print(f'\n=== Browse Mode ===')
    utils.info_print(f'Total data rows: {total_data_rows} (rows {first_data_row + 1} to {last_data_row + 1})')
    utils.info_print(f'Participants: {", ".join(participant_headers)}')
    utils.info_print(f'\nCommands: up/u, down/d, pageup/pu, pagedown/pd, jump/j <row>, open/o, quit/q')
    utils.info_print(f'Press Enter to move down one row.\n')
    
    def display_rows(start_row, num_rows):
        """Display num_rows starting from start_row (0-indexed)."""
        utils.info_print('-' * 60)
        for i in range(num_rows):
            row_idx = start_row + i
            if row_idx > last_data_row:
                break
            
            row_data = sheet_data[row_idx]
            row_num = row_idx + 1  # 1-indexed for display
            
            # Get category
            category_col = category_cell.col - 1  # 0-indexed
            category = row_data[category_col] if category_col < len(row_data) else ''
            
            # Get description (observation)
            desc_col = observation_cell.col - 1  # 0-indexed
            description = row_data[desc_col] if desc_col < len(row_data) else ''
            
            utils.info_print(f'Row {row_num}')
            utils.info_print(f'  Category: {category if category else "(empty)"}')
            utils.info_print(f'  Description: {description if description else "(empty)"}')
            
            # Get participant timestamps
            has_timestamps = False
            participant_data = []
            for j, participant_id in enumerate(participant_headers):
                col_idx = id_cell.col + j  # 0-indexed
                if col_idx < len(row_data):
                    timestamp_value = row_data[col_idx]
                    if timestamp_value and timestamp_value.strip():
                        # Replace newlines with commas for display
                        timestamp_display = timestamp_value.replace('\n', ', ').replace('\r', '')
                        participant_data.append(f'    {participant_id}: {timestamp_display}')
                        has_timestamps = True
            
            if has_timestamps:
                utils.info_print('  Participants:')
                for p_data in participant_data:
                    utils.info_print(p_data)
            else:
                utils.info_print('  Participants: (no timestamps)')
            
            utils.info_print('  ---')
        
        # Show position info
        displayed_end = min(start_row + num_rows, last_data_row + 1)
        utils.info_print(f'\nShowing rows {start_row + 1}-{displayed_end} of {last_data_row + 1}')
    
    # Initial display
    display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
    
    # Navigation loop
    while True:
        user_input = input('\n>> ').strip().lower()
        
        if user_input in ('quit', 'q'):
            utils.info_print('Exiting browse mode.')
            break
        elif user_input in ('up', 'u'):
            if current_row > first_data_row:
                current_row -= 1
                display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
            else:
                utils.info_print('Already at the first row.')
        elif user_input in ('down', 'd', ''):
            if current_row < last_data_row:
                current_row += 1
                display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
            else:
                utils.info_print('Already at the last row.')
        elif user_input in ('pageup', 'pu'):
            new_row = max(first_data_row, current_row - config.BROWSE_LINES_TO_DISPLAY)
            if new_row != current_row:
                current_row = new_row
                display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
            else:
                utils.info_print('Already at the first row.')
        elif user_input in ('pagedown', 'pd'):
            new_row = min(last_data_row, current_row + config.BROWSE_LINES_TO_DISPLAY)
            if new_row != current_row:
                current_row = new_row
                display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
            else:
                utils.info_print('Already at the last row.')
        elif user_input.startswith('jump ') or user_input.startswith('j '):
            try:
                parts = user_input.split()
                if len(parts) >= 2:
                    target_row = int(parts[1]) - 1  # Convert to 0-indexed
                    if target_row < first_data_row:
                        utils.info_print(f'Row number must be at least {first_data_row + 1}.')
                    elif target_row > last_data_row:
                        utils.info_print(f'Row number must be at most {last_data_row + 1}.')
                    else:
                        current_row = target_row
                        display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
                else:
                    utils.info_print('Usage: jump <row_number> or j <row_number>')
            except ValueError:
                utils.info_print('Invalid row number. Usage: jump <row_number> or j <row_number>')
        elif user_input in ('open', 'o'):
            try:
                spreadsheet_url = sheet.spreadsheet.url
                utils.info_print(f'Opening spreadsheet in browser: {sheet.spreadsheet.title}')
                webbrowser.open(spreadsheet_url)
                utils.info_print('Spreadsheet opened in your default browser.')
            except AttributeError as e:
                utils.error_print('Could not retrieve spreadsheet URL.', [f'Error: {e}'])
            except Exception as e:
                utils.error_print('Could not open browser.', [f'Error: {e}'])
        else:
            utils.info_print('Unknown command. Available: up/u, down/d, pageup/pu, pagedown/pd, jump/j <row>, open/o, quit/q')
            utils.info_print('Press Enter to move down one row.')
