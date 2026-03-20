# -*- coding: utf-8 -*-
"""Interactive prompt helpers for clipgen.

All user-facing interactive prompts (line selection, range selection, cell
selection, participant selection, category selection, batch/keyword confirmation,
and browse mode) live here. Generation functions in spreadsheet.py are kept pure:
they take resolved parameters and return clip records, never prompting the user.

Each prompt function takes a SheetContext for data preview and returns the
resolved parameters (or None if the user cancels / enters invalid input and
the caller should re-prompt or abort).
"""

import shutil
import sys
import termios
import tty
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
        - ctx.first_data_row_idx
        - (1 if ctx.filename_row_idx is not None else 0)
    )
    msg = f"\nThis will generate {len(clips)} clips (from {num_data_rows} data rows and {ctx.num_participants} participant column(s)). Proceed? [y/n]\n>> "
    yn = utils.read_user_input(msg)
    return yn == "y"


def prompt_category_selection(ctx: SheetContext) -> Optional[List[str]]:
    """Show categories, prompt for selection, return selected names or None."""
    all_categories = spreadsheet.collect_categories(ctx)
    if not all_categories:
        utils.info_print("No categories found in the spreadsheet.")
        return None
    utils.info_print("Available categories:")
    for i, cat in enumerate(all_categories, 1):
        utils.info_print(f"  {i}. {cat}")
    while True:
        selection = utils.read_user_input(
            '\nEnter category numbers (comma-separated, e.g., "1,3,5") or "all":\n>> '
        )
        if selection.lower() == "all":
            return all_categories
        try:
            indices = [int(x.strip()) for x in selection.split(",")]
            selected_categories = []
            invalid_indices = []
            for idx in indices:
                if 1 <= idx <= len(all_categories):
                    if all_categories[idx - 1] not in selected_categories:
                        selected_categories.append(all_categories[idx - 1])
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
                yn = utils.read_user_input("\nIs this correct? [y/n]\n>> ")
                if yn == "y":
                    return selected_categories
            else:
                utils.info_print("No valid categories selected. Please try again.")
        except ValueError:
            tokens = [t.strip() for t in selection.split(",") if t.strip()]
            matched_categories = []
            unmatched = []
            for token in tokens:
                exact = next(
                    (c for c in all_categories if c.lower() == token.lower()), None
                )
                if exact:
                    if exact not in matched_categories:
                        matched_categories.append(exact)
                    continue
                suggestion = utils.suggest_close_match(token, all_categories)
                if suggestion is not None:
                    if suggestion not in matched_categories:
                        matched_categories.append(suggestion)
                else:
                    unmatched.append(token)
            if unmatched:
                utils.info_print(f"Could not match: {', '.join(unmatched)}")
            if matched_categories:
                utils.info_print("Selected categories:")
                for cat in matched_categories:
                    utils.info_print(f"  - {cat}")
                yn = utils.read_user_input("\nIs this correct? [y/n]\n>> ")
                if yn == "y":
                    return matched_categories
            else:
                utils.info_print(
                    "No valid categories matched. Enter numbers (e.g. 1,3,5) or category names."
                )


def prompt_severity_selection(ctx: SheetContext) -> Optional[List[str]]:
    """Show severities sorted most severe first, prompt for selection, return selected or None."""
    all_severities, severity_counts = spreadsheet.collect_severities(ctx)
    if not all_severities:
        utils.info_print("No severity values found in the spreadsheet.")
        return None
    utils.info_print("Available severity levels (most severe first):")
    for i, sev in enumerate(all_severities, 1):
        count = severity_counts.get(sev, 0)
        count_label = "1 row" if count == 1 else f"{count} rows"
        display = f"{utils.format_severity_display(sev)} \u2014 {count_label}"
        style = utils.get_severity_style(sev)
        if utils._use_rich() and utils.console is not None and style:
            utils.console.print(f"  {i}. [{style}]{display}[/{style}]")
        else:
            utils.info_print(f"  {i}. {display}")
    while True:
        selection = utils.read_user_input(
            '\nEnter severity numbers (comma-separated, e.g., "1,2") or "all":\n>> '
        )
        if selection.lower() == "all":
            return all_severities
        try:
            indices = [int(x.strip()) for x in selection.split(",")]
            selected = []
            invalid_indices = []
            for idx in indices:
                if 1 <= idx <= len(all_severities):
                    if all_severities[idx - 1] not in selected:
                        selected.append(all_severities[idx - 1])
                else:
                    invalid_indices.append(idx)
            if invalid_indices:
                utils.info_print(
                    f"  Invalid index(es): {', '.join(str(i) for i in invalid_indices)}"
                )
            if selected:
                utils.info_print("Selected severities:")
                for sev in selected:
                    utils.info_print(f"  - {utils.format_severity_display(sev)}")
                yn = utils.read_user_input("\nIs this correct? [y/n]\n>> ")
                if yn == "y":
                    return selected
            else:
                utils.info_print("No valid severities selected. Please try again.")
        except ValueError:
            tokens = [t.strip() for t in selection.split(",") if t.strip()]
            matched = []
            unmatched = []
            for token in tokens:
                normalized = utils.normalize_severity(token)
                exact = next(
                    (s for s in all_severities if s.lower() == normalized.lower()),
                    None,
                )
                if exact:
                    if exact not in matched:
                        matched.append(exact)
                else:
                    unmatched.append(token)
            if unmatched:
                utils.info_print(f"Could not match: {', '.join(unmatched)}")
            if matched:
                utils.info_print("Selected severities:")
                for sev in matched:
                    utils.info_print(f"  - {utils.format_severity_display(sev)}")
                yn = utils.read_user_input("\nIs this correct? [y/n]\n>> ")
                if yn == "y":
                    return matched
            else:
                utils.info_print(
                    "No valid severities matched. Enter numbers or severity names."
                )


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
        yn = utils.read_user_input("Are these the correct issues? [y/n]\n>> ")
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
        yn = utils.read_user_input("Is this correct? [y/n]\n>> ")
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
            yn = utils.read_user_input("Are these the correct cells? [y/n]\n>> ")
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
            still_invalid = []
            for token in invalid_tokens:
                suggestion = utils.suggest_close_match(token, available_list)
                if suggestion is not None:
                    chosen_ids.append(suggestion)
                else:
                    still_invalid.append(token)
            if still_invalid:
                utils.info_print(
                    f"Not found: {', '.join(still_invalid)}. Available: {', '.join(available_list)}"
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
            "Generate all clips for these participants? [y/n]\n>> "
        )
        if yn == "y":
            return unique_ids


def prompt_keyword_confirm() -> bool:
    """Confirm keyword mode. Returns True to proceed."""
    yn = utils.read_user_input(
        "\nKeyword mode will include only key-marked timestamps (per-cell annotations). Do you want to proceed? [y/n]\n>> "
    )
    return yn == "y"


# ---- Browse mode ----


def browse_spreadsheet(sheet: Any, *, process_fn=None) -> None:
    """Interactive browse mode for viewing spreadsheet rows line by line.

    Allows users to navigate through the spreadsheet to inspect issues
    before generating clips. Shows row number, category, description,
    and participant/group timestamps for each row.

    When process_fn is provided, users can generate clips/screenshots/GIFs
    directly from browse by typing selectors (line numbers, ranges, cells,
    participants, quoted categories).

    Args:
        sheet: Worksheet object (gspread or Excel adapter).
        process_fn: Optional callback (clips_list, output_format) -> (outputs_generated, artifacts).
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

    # Browse state: all row indices are 0-based (into sheet_data).
    # Data starts immediately below the Observation header row.
    first_data_row = observation_cell.row
    last_data_row = len(sheet_data) - 1
    total_data_rows = last_data_row - first_data_row + 1

    current_row = first_data_row
    output_format = "clip"

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

    def _print_search_bar_top():
        w = max(40, min(shutil.get_terminal_size().columns - 2, 80))
        control_label = "\u2191/\u2193 or Enter to navigate, pageup/pu, pagedown/pd, jump/j <row>, open/o, quit/q \u2014 or type to search"
        if process_fn is not None:
            format_indicator = output_format.upper()
            generate_label = (
                f'Generate: lines, ranges, P01.11, P01, "Category"'
                f" | Format: screen, gif, clip  [{format_indicator}]"
            )
            control_label = f"{control_label}\n{generate_label}"
        label = " Search or command "
        bar = "\u2500" * (w - 2 - len(label))
        top = f"{control_label}\n\n\u2500{label}{bar}\u2500"
        if utils._use_rich() and utils.console is not None:
            utils.console.print(f"[dim]{top}[/dim]")
        else:
            print(top)

    def _search_rows(query, *, fuzzy=False):
        query_lower = query.lower()
        matches = []
        cat_col = category_cell.col - 1
        desc_col_idx = observation_cell.col - 1

        if fuzzy:
            from difflib import SequenceMatcher

            query_words = query_lower.split()

        for row_idx in range(first_data_row, last_data_row + 1):
            row_data = sheet_data[row_idx]
            desc = row_data[desc_col_idx] if desc_col_idx < len(row_data) else ""
            cat = row_data[cat_col] if cat_col < len(row_data) else ""

            if not fuzzy:
                if query_lower in desc.lower():
                    matches.append(row_idx)
                    continue
                if query_lower in cat.lower():
                    matches.append(row_idx)
                    continue
                for j in range(num_participants):
                    col_idx = id_cell.col + j
                    if col_idx < len(row_data):
                        cell_val = row_data[col_idx]
                        if cell_val and query_lower in cell_val.lower():
                            matches.append(row_idx)
                            break
            else:
                text_words = (desc + " " + cat).lower().split()
                if text_words and all(
                    any(
                        SequenceMatcher(None, qw, tw).ratio() >= 0.75
                        for tw in text_words
                    )
                    for qw in query_words
                ):
                    matches.append(row_idx)

        return matches

    # Initial display
    display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)

    _NOOP = "\x00_noop"

    def _read_browse_key() -> str:
        """Read arrow keys instantly or fall back to line input for text commands.

        Renders a search bar around the input: top border is printed before
        reading, right edge + bottom border after the user finishes typing.
        """
        print()  # blank line before search bar
        _print_search_bar_top()
        prompt = "⌕ "

        if not sys.stdin.isatty():
            return utils.read_user_input(prompt)

        sys.stdout.write(prompt)
        sys.stdout.flush()

        fd = sys.stdin.fileno()
        old_settings = termios.tcgetattr(fd)
        try:
            tty.setraw(fd)
            ch = sys.stdin.read(1)

            if ch == "\x1b":
                seq = sys.stdin.read(2)
                termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
                sys.stdout.write("\r\n")
                if len(seq) == 2 and seq[0] == "[":
                    if seq[1] == "A":
                        return "up"
                    if seq[1] == "B":
                        return "down"
                return _NOOP

            if ch in ("\r", "\n"):
                termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
                sys.stdout.write("\r\n")
                return ""

            if ch == "\x03":
                termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
                sys.stdout.write("\r\n")
                raise KeyboardInterrupt

            if ch == "\x04":
                termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
                sys.stdout.write("\r\n")
                raise EOFError

        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)

        # Printable character: echo it and read the rest as a normal line
        sys.stdout.write(ch)
        sys.stdout.flush()
        rest = input()
        full_line = ch + rest

        value = full_line.strip()
        if not value:
            return value
        first_token = value.split()[0].lower()
        if first_token in ("quit", "exit"):
            utils.info_print("Exiting clipgen.")
            raise utils.QuitProgram()
        if first_token == "top":
            utils.info_print("Returning to spreadsheet selection.")
            raise utils.TopToSpreadsheet()
        if first_token == "back":
            utils.info_print("Returning to mode selection.")
            raise utils.BackToModeSelection()
        return value

    # Navigation loop
    while True:
        raw_input = _read_browse_key().strip()
        user_input = raw_input.lower()

        if user_input == _NOOP:
            continue
        elif user_input in ("quit", "q"):
            utils.info_print("Exiting browse mode.")
            break
        elif user_input in ("up", "u"):
            if current_row > first_data_row:
                new_row = max(
                    first_data_row, current_row - config.BROWSE_LINES_TO_SCROLL
                )
                current_row = new_row
                display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
            else:
                utils.info_print("Already at the first row.")
        elif user_input in ("down", "d", ""):
            if current_row < last_data_row:
                new_row = min(
                    last_data_row, current_row + config.BROWSE_LINES_TO_SCROLL
                )
                current_row = new_row
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
        elif user_input in ("screen", "sc"):
            output_format = "screen"
            utils.info_print("Switched to screenshot mode.")
        elif user_input == "gif":
            output_format = "gif"
            utils.info_print("Switched to GIF mode.")
        elif user_input == "clip":
            output_format = "clip"
            utils.info_print("Switched to clip mode.")
        else:
            # Try selector parsing for artifact generation
            if process_fn is not None:
                parsed = spreadsheet.parse_reel_input(raw_input)
                has_selectors = (
                    parsed.get("batch")
                    or parsed.get("keyword")
                    or parsed["lines"]
                    or parsed["ranges"]
                    or parsed["cells"]
                    or parsed["participants"]
                    or parsed["categories"]
                )
                if has_selectors:
                    if parsed.get("timeline"):
                        utils.info_print(
                            "Timeline selector is not supported in browse mode."
                        )
                    else:
                        clips = spreadsheet.generate_list(
                            sheet, "reel", reel_input=raw_input
                        )
                        if clips:
                            format_label = {
                                "clip": "clip(s)",
                                "screen": "screenshot(s)",
                                "gif": "GIF(s)",
                            }.get(output_format, "file(s)")
                            utils.info_print(
                                f"Generating {len(clips)} {format_label}..."
                            )
                            outputs_generated, _ = process_fn(clips, output_format)
                            utils.info_print(
                                f"Done! Generated {outputs_generated} {format_label}."
                            )
                        else:
                            utils.info_print("No clips found for the given selectors.")
                    continue

            # Treat unrecognized input as a search query
            matches = _search_rows(user_input)
            if not matches:
                matches = _search_rows(user_input, fuzzy=True)
                if matches:
                    utils.info_print(
                        f"No exact matches for '{user_input}'. Showing approximate matches:"
                    )
            if matches:
                match_row_nums = [str(m + 1) for m in matches]
                if len(match_row_nums) > 20:
                    shown = ", ".join(match_row_nums[:20])
                    extra = len(match_row_nums) - 20
                    utils.info_print(
                        f"Found {len(matches)} matching row(s): {shown} \u2026 and {extra} more"
                    )
                else:
                    utils.info_print(
                        f"Found {len(matches)} matching row(s): {', '.join(match_row_nums)}"
                    )
                current_row = matches[0]
                display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
            else:
                utils.info_print(f"No rows matching '{user_input}'.")
