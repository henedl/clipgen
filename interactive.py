# -*- coding: utf-8 -*-
"""Interactive prompt helpers for clipgen.

All user-facing interactive prompts (line selection, range selection, cell
selection, participant selection, category selection, batch/filter confirmation,
and browse mode) live here. Generation functions in spreadsheet.py are kept pure:
they take resolved parameters and return clip records, never prompting the user.

Each prompt function takes a SheetContext for data preview and returns the
resolved parameters (or None if the user cancels / enters invalid input and
the caller should re-prompt or abort).
"""

import webbrowser
from typing import Any, List, Optional, Tuple

import config
import spreadsheet
import utils
from spreadsheet import SheetContext


# ---- Prompt helpers ----


def prompt_batch_confirm(ctx: SheetContext) -> bool:
    """Show clip count and confirm batch mode. Returns True to proceed."""
    clips = spreadsheet.generate_batch_timestamps(ctx)
    num_data_rows = (
        len(ctx.sheet_data)
        - (ctx.id_cell.row + 1)
        - (1 if ctx.filename_row_idx is not None else 0)
    )
    msg = f"\nThis will generate {len(clips)} clips (from {num_data_rows} data rows and {ctx.num_participants} participant column(s)). Proceed? y/n\n>> "
    yn = utils.read_user_input(msg)
    return yn == "y"


def prompt_category_selection(ctx: SheetContext) -> Optional[List[str]]:
    """Show categories, prompt for selection, return selected names or None."""
    all_categories = spreadsheet.collect_categories(ctx)
    if not all_categories:
        utils.info_print("No categories found in the spreadsheet.")
        return None
    return _interactive_category_selection(all_categories)


def prompt_line_selection(ctx: SheetContext) -> Optional[List[int]]:
    """Prompt for line numbers, validate bounds, show descriptions, confirm.

    Returns validated line numbers list, or None if user cancels.
    """
    while True:
        try:
            line_input = utils.read_user_input(
                "\nWhich issue(s)? Enter row number(s), comma-separated for multiple.\n>> "
            )
            parsed_numbers = [int(num.strip()) for num in line_input.split(",")]
        except ValueError:
            utils.info_print(
                "Try again. Enter row numbers as integers, separated by commas."
            )
            continue
        utils.info_print("Selected issues:")
        valid_lines = []
        for line_num in parsed_numbers:
            if line_num < 1 or line_num > len(ctx.sheet_data):
                utils.info_print(f"  Line {line_num}: [INVALID - out of range]")
            elif (
                ctx.filename_row_idx is not None
                and line_num - 1 == ctx.filename_row_idx
            ):
                utils.info_print(
                    f"  Line {line_num}: [RESERVED - filename overrides row]"
                )
            else:
                desc = ctx.sheet_data[line_num - 1][ctx.observation_cell.col - 1]
                utils.info_print(f"  Line {line_num}: {desc}")
                valid_lines.append(line_num)
        if not valid_lines:
            utils.info_print("No valid lines selected. Please try again.")
            continue
        utils.info_print("")
        yn = utils.read_user_input("Are these the correct issues? y/n\n>> ")
        if yn == "y":
            return valid_lines


def prompt_range_selection(ctx: SheetContext) -> Optional[Tuple[int, int]]:
    """Prompt for start/end row, validate, confirm. Returns (start, end) or None."""
    max_row = len(ctx.sheet_data)
    while True:
        try:
            start_line_str = utils.read_user_input(
                "\nWhich starting line (row number only)?\n>> "
            )
            end_line_str = utils.read_user_input(
                "\nWhich ending line (row number only)?\n>> "
            )
            start_line = int(start_line_str)
            end_line = int(end_line_str)
        except ValueError:
            utils.info_print("Invalid input. Please enter row numbers as integers.")
            continue
        if spreadsheet._validate_row_range(start_line, end_line, max_row) is None:
            continue
        utils.info_print(
            f"Lines selected: {ctx.sheet_data[start_line - 1][ctx.observation_cell.col - 1]} to {ctx.sheet_data[end_line - 1][ctx.observation_cell.col - 1]}"
        )
        yn = utils.read_user_input("Is this correct? y/n\n>> ")
        if yn == "y":
            return (start_line, end_line)


def prompt_cell_selection(ctx: SheetContext) -> Optional[List[Tuple[str, int]]]:
    """Prompt for cell specs (e.g. P01.11), validate, preview, confirm.

    Returns list of (participant_id, row_number) tuples, or None if user cancels.
    """
    while True:
        try:
            cell_input = utils.read_user_input(
                "\nEnter cell specification(s) (e.g., P01.11 or P01.11 + P03.11):\n>> "
            )
            if not cell_input.strip():
                utils.info_print("Please enter at least one cell specification.")
                continue
            try:
                parsed_specs = spreadsheet.parse_cell_specifications(cell_input)
            except ValueError as e:
                utils.info_print(f"Invalid format: {e}")
                utils.info_print("Expected format: P01.11 or P01.11 + P03.11")
                continue
            utils.info_print("Selected cells:")
            valid_specs = []
            for participant_id, row_number in parsed_specs:
                col_idx = spreadsheet.find_participant_column(
                    ctx.header_row, ctx.id_cell, participant_id
                )
                if col_idx is None:
                    utils.info_print(
                        f"  {participant_id}.{row_number}: [INVALID - participant not found]"
                    )
                elif row_number < 1 or row_number > len(ctx.sheet_data):
                    utils.info_print(
                        f"  {participant_id}.{row_number}: [INVALID - row out of range]"
                    )
                else:
                    row_idx = row_number - 1
                    cell_value = ""
                    if col_idx < len(ctx.sheet_data[row_idx]):
                        cell_value = ctx.sheet_data[row_idx][col_idx]
                    desc = ""
                    desc_col = ctx.observation_cell.col - 1
                    if desc_col >= 0 and desc_col < len(ctx.sheet_data[row_idx]):
                        desc = ctx.sheet_data[row_idx][desc_col]
                    if cell_value and cell_value.strip():
                        display_value = cell_value.replace("\n", " ")
                        desc_preview = (
                            desc[: config.DESCRIPTION_PREVIEW_LENGTH] if desc else "N/A"
                        )
                        utils.info_print(
                            f"  {participant_id}.{row_number}: {display_value} (row: {desc_preview})"
                        )
                    else:
                        utils.info_print(f"  {participant_id}.{row_number}: [EMPTY]")
                    valid_specs.append((participant_id, row_number))
            if not valid_specs:
                utils.info_print("No valid cells found. Please try again.")
                continue
            utils.info_print("")
            yn = utils.read_user_input("Are these the correct cells? y/n\n>> ")
            if yn == "y":
                return valid_specs
        except KeyboardInterrupt:
            utils.info_print("Cancelled by user.")
            return None


def prompt_participant_selection(ctx: SheetContext) -> Optional[List[str]]:
    """Show available participants, prompt for selection, validate, confirm.

    Returns list of participant IDs, or None if no participants found.
    """
    available_list = spreadsheet.get_participant_list(
        ctx.header_row, ctx.id_cell, ctx.num_participants
    )
    if not available_list:
        utils.info_print("No participants found in the spreadsheet.")
        return None
    utils.info_print("Available participants:")
    for i, pid in enumerate(available_list, 1):
        utils.info_print(f"  {i}. {pid}")
    while True:
        selection = utils.read_user_input(
            "\nEnter participant number(s) or ID(s), separated by + or , (e.g., 1, 3 or P01, P03):\n>> "
        ).strip()
        if not selection:
            utils.info_print("Please enter one or more participant numbers or IDs.")
            continue
        tokens = spreadsheet.parse_participant_selection(selection)
        if not tokens:
            utils.info_print(
                "No valid participant(s) entered. Use + or , as separator."
            )
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
                col_idx = spreadsheet.find_participant_column(
                    ctx.header_row, ctx.id_cell, token
                )
                if col_idx is not None:
                    if col_idx < len(ctx.header_row):
                        chosen_ids.append(
                            utils.normalize_participant_id(ctx.header_row[col_idx])
                        )
                    else:
                        chosen_ids.append(token)
                else:
                    invalid_tokens.append(token)
        if invalid_tokens:
            utils.info_print(
                f"Not found: {', '.join(invalid_tokens)}. Available: {', '.join(available_list)}"
            )
            continue
        seen = set()
        unique_ids = []
        for pid in chosen_ids:
            if pid not in seen:
                seen.add(pid)
                unique_ids.append(pid)
        utils.info_print(f"Selected participant(s): {', '.join(unique_ids)}")
        yn = utils.read_user_input(
            "Generate all clips for these participants? y/n\n>> "
        )
        if yn == "y":
            return unique_ids


def prompt_filter_confirm() -> bool:
    """Confirm filter mode. Returns True to proceed."""
    yn = utils.read_user_input(
        "\nFilter mode will include only key-marked timestamps (per-cell annotations). Do you want to proceed? y/n\n>> "
    )
    return yn == "y"


# ---- Browse mode ----


def browse_spreadsheet(sheet: Any) -> None:
    """Interactive browse mode for viewing spreadsheet rows line by line.

    Allows users to navigate through the spreadsheet to inspect issues
    before generating clips. Shows row number, category, description,
    and participant/group timestamps for each row.
    """

    def _load_browse_data() -> tuple:
        header_result = spreadsheet.validate_spreadsheet_headers(sheet)
        if header_result is None:
            return (None, None)
        sheet_data = sheet.get_all_values()
        return (header_result, sheet_data)

    header_result, sheet_data = (
        utils.run_with_spinner("Loading spreadsheet...", _load_browse_data)
        if utils.use_progress()
        else _load_browse_data()
    )
    if header_result is None:
        return

    id_cell, observation_cell, category_cell = header_result
    utils.debug_print(f"Sheet dumped into memory at {utils.get_current_time()}")

    # Check if sheet is empty or has only headers
    if len(sheet_data) <= 1:
        utils.error_print("Spreadsheet appears to be empty (no data rows found).")
        return

    # Get participant info
    header_row = sheet.row_values(id_cell.row)
    num_participants = spreadsheet.get_num_participants(
        header_row, id_cell, sheet.col_count
    )

    if num_participants == 0:
        utils.warning_print(
            "No participant columns found in the spreadsheet.",
            [
                f"Looking for columns starting with: {', '.join(config.PARTICIPANT_PREFIXES)}"
            ],
        )
        return

    # Browse state: all row indices are 0-based (into sheet_data). id_cell.row is 1-based.
    first_data_row = id_cell.row
    last_data_row = len(sheet_data) - 1
    total_data_rows = last_data_row - first_data_row + 1

    current_row = first_data_row

    # Participant column headers for display (id_cell.col is 1-based)
    participant_headers = []
    for col_idx in range(id_cell.col, id_cell.col + num_participants):
        if col_idx < len(header_row):
            participant_headers.append(header_row[col_idx])

    utils.print_mode_heading("Browse mode", "mode.browse")
    utils.info_print(
        f"Total data rows: {total_data_rows} (rows {first_data_row + 1} to {last_data_row + 1})"
    )
    utils.info_print(f"Participants: {', '.join(participant_headers)}")
    utils.info_print(
        "Commands: up/u, down/d, pageup/pu, pagedown/pd, jump/j <row>, open/o, quit/q"
    )
    utils.info_print("Press Enter to move down one row.")

    def display_rows(start_row, num_rows):
        """Display num_rows starting from start_row (0-indexed into sheet_data)."""
        category_col = category_cell.col - 1
        desc_col = observation_cell.col - 1

        rows_data = []
        for i in range(num_rows):
            row_idx = start_row + i
            if row_idx > last_data_row:
                break

            row_data = sheet_data[row_idx]
            row_num = row_idx + 1

            category = row_data[category_col] if category_col < len(row_data) else ""
            description = row_data[desc_col] if desc_col < len(row_data) else ""

            timestamps = {}
            for j, participant_id in enumerate(participant_headers):
                col_idx = id_cell.col + j
                if col_idx < len(row_data):
                    timestamp_value = row_data[col_idx]
                    if timestamp_value and timestamp_value.strip():
                        timestamps[participant_id] = timestamp_value.replace(
                            "\n", ", "
                        ).replace("\r", "")

            rows_data.append(
                {
                    "row_num": row_num,
                    "category": category or "(empty)",
                    "description": description or "(empty)",
                    "timestamps": timestamps,
                }
            )

        if utils._use_rich():
            table = utils.create_browse_table(rows_data, participant_headers)
            if table and utils.console is not None:
                utils.console.print(table)
        else:
            output = utils.format_browse_rows_plain(rows_data, participant_headers)
            print(output)

        displayed_end = min(start_row + num_rows, last_data_row + 1)
        utils.info_print(
            f"Showing rows {start_row + 1}-{displayed_end} of {last_data_row + 1}"
        )

    # Initial display
    display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)

    # Navigation loop
    while True:
        user_input = utils.read_user_input("\n>> ").strip().lower()

        if user_input in ("quit", "q"):
            utils.info_print("Exiting browse mode.")
            break
        elif user_input in ("up", "u"):
            if current_row > first_data_row:
                current_row -= 1
                display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
            else:
                utils.info_print("Already at the first row.")
        elif user_input in ("down", "d", ""):
            if current_row < last_data_row:
                current_row += 1
                display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
            else:
                utils.info_print("Already at the last row.")
        elif user_input in ("pageup", "pu"):
            new_row = max(first_data_row, current_row - config.BROWSE_LINES_TO_DISPLAY)
            if new_row != current_row:
                current_row = new_row
                display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
            else:
                utils.info_print("Already at the first row.")
        elif user_input in ("pagedown", "pd"):
            new_row = min(last_data_row, current_row + config.BROWSE_LINES_TO_DISPLAY)
            if new_row != current_row:
                current_row = new_row
                display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
            else:
                utils.info_print("Already at the last row.")
        elif user_input.startswith("jump ") or user_input.startswith("j "):
            try:
                parts = user_input.split()
                if len(parts) >= 2:
                    target_row = int(parts[1]) - 1
                    if target_row < first_data_row:
                        utils.info_print(
                            f"Row number must be at least {first_data_row + 1}."
                        )
                    elif target_row > last_data_row:
                        utils.info_print(
                            f"Row number must be at most {last_data_row + 1}."
                        )
                    else:
                        current_row = target_row
                        display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
                else:
                    utils.info_print("Usage: jump <row_number> or j <row_number>")
            except ValueError:
                utils.info_print(
                    "Invalid row number. Usage: jump <row_number> or j <row_number>"
                )
        elif user_input in ("open", "o"):
            spreadsheet_url = getattr(getattr(sheet, "spreadsheet", None), "url", None)
            if not spreadsheet_url:
                utils.info_print(
                    "Opening in browser is not available for local Excel files."
                )
            else:
                try:
                    utils.info_print(
                        f"Opening spreadsheet in browser: {sheet.spreadsheet.title}"
                    )
                    webbrowser.open(spreadsheet_url)
                    utils.info_print("Spreadsheet opened in your default browser.")
                except OSError as e:
                    utils.error_print("Could not open browser.", [f"Error: {e}"])
        else:
            utils.info_print(
                "Unknown command. Available: up/u, down/d, pageup/pu, pagedown/pd, jump/j <row>, open/o, quit/q"
            )
            utils.info_print("Press Enter to move down one row.")


# ---- Internal helpers ----


def _interactive_category_selection(categories: List[str]) -> List[str]:
    """Interactively select one or more category names from a numbered list."""
    utils.info_print("Available categories:")
    for i, cat in enumerate(categories, 1):
        utils.info_print(f"  {i}. {cat}")
    while True:
        selection = utils.read_user_input(
            '\nEnter category numbers (comma-separated, e.g., "1,3,5") or "all":\n>> '
        )
        if selection.lower() == "all":
            return categories
        try:
            indices = [int(x.strip()) for x in selection.split(",")]
            selected_categories = []
            invalid_indices = []
            for idx in indices:
                if 1 <= idx <= len(categories):
                    if categories[idx - 1] not in selected_categories:
                        selected_categories.append(categories[idx - 1])
                else:
                    invalid_indices.append(idx)
            if invalid_indices:
                utils.info_print(
                    f"  Invalid index(es): {', '.join(str(i) for i in invalid_indices)}"
                )
            if selected_categories:
                utils.info_print("Selected categories:")
                for cat in selected_categories:
                    utils.info_print(f"  - {cat}")
                yn = utils.read_user_input("\nIs this correct? y/n\n>> ")
                if yn == "y":
                    return selected_categories
            else:
                utils.info_print("No valid categories selected. Please try again.")
        except ValueError:
            utils.info_print("Please enter valid numbers separated by commas.")
