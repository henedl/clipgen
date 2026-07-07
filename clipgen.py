# -*- coding: utf-8 -*-
"""clipgen - Video clip generator from Google Sheets timestamps.

This program will help quickly create video snippets from longer video files, based on timestamps in a spreadsheet!
Check out README.md for more detailed information about setting up and using clipgen.

Data flow: Spreadsheet -> Worksheet -> clip records (cell, desc, study, participant, category; 'times' added by prepare_clip) -> ffmpeg clips or reel.

This script supports full unicode/UTF-8 for international characters in:
- Study names
- Participant IDs
- Category names
- Descriptions
- File paths
"""

import sys
from pathlib import Path
from typing import Any, Callable

import gspread

import config
import files
import google_api
import interactive
import spreadsheet
import utils
import video
import viewer
from utils import ClipRecord

# Re-exported from pipeline.py for backward compatibility
from pipeline import (
    is_excel_worksheet as _is_excel_worksheet,
    process_clips,
    process_reel,
    regenerate_from_manifest,
)

# ---- Mode configuration ----

MODE_ALIASES = {
    "b": "batch",
    "batch": "batch",
    "l": "line",
    "line": "line",
    "r": "range",
    "range": "range",
    "c": "category",
    "cat": "category",
    "category": "category",
    "ce": "cell",
    "cell": "cell",
    "p": "participant",
    "participant": "participant",
    "k": "keyword",
    "keyword": "keyword",
    "sv": "severity",
    "severity": "severity",
    "s": "screen",
    "screen": "screen",
    "g": "gif",
    "gif": "gif",
    "re": "reel",
    "reel": "reel",
    "rl": "reellate",
    "reellate": "reellate",
    "br": "browse",
    "browse": "browse",
    "v": "viewer",
    "viewer": "viewer",
    "tv": "timeline-viewer",
    "timeline-viewer": "timeline-viewer",
    "rg": "regenerate",
    "regenerate": "regenerate",
    "gv": "gallery",
    "gallery": "gallery",
    "st": "studio",
    "studio": "studio",
    "ss": "screenspace",
    "screenspace": "screenspace",
    "tr": "transcripts",
    "transcripts": "transcripts",
    "ex": "export",
    "export": "export",
    "se": "settings",
    "settings": "settings",
}

FORMAT_MODE_ALIASES = {
    alias: mode
    for alias, mode in MODE_ALIASES.items()
    if mode
    in {
        "batch",
        "line",
        "range",
        "category",
        "cell",
        "participant",
        "keyword",
        "severity",
    }
}

# Dispatch table for standard interactive modes: mode -> (prompt_fn, generate_fn).
# prompt_fn(ctx) returns a result or None/False to cancel.
# generate_fn(ctx, result) returns a list of clip records.
_STANDARD_MODES = {
    "batch": (
        lambda ctx: interactive.prompt_batch_confirm(ctx),
        lambda ctx, _: spreadsheet.generate_batch_timestamps(ctx),
    ),
    "line": (
        lambda ctx: interactive.prompt_line_selection(ctx),
        lambda ctx, v: spreadsheet.generate_line_timestamps(ctx, v),
    ),
    "range": (
        lambda ctx: interactive.prompt_range_selection(ctx),
        lambda ctx, v: spreadsheet.generate_range_timestamps(ctx, *v),
    ),
    "category": (
        lambda ctx: interactive.prompt_category_selection(ctx),
        lambda ctx, v: spreadsheet.generate_category_timestamps(ctx, v),
    ),
    "cell": (
        lambda ctx: interactive.prompt_cell_selection(ctx),
        lambda ctx, v: spreadsheet.generate_cell_timestamps(ctx, v),
    ),
    "participant": (
        lambda ctx: interactive.prompt_participant_selection(ctx),
        lambda ctx, pids: [
            c
            for pid in pids
            for c in spreadsheet.generate_participant_timestamps(ctx, pid)
        ],
    ),
    "keyword": (
        lambda ctx: interactive.prompt_keyword_selection(ctx),
        lambda ctx, v: spreadsheet.generate_keyword_timestamps(ctx, annotation_ids=v),
    ),
    "severity": (
        lambda ctx: interactive.prompt_severity_selection(ctx),
        lambda ctx, v: spreadsheet.generate_severity_timestamps(ctx, v),
    ),
}


# ---- Spreadsheet opening and selection ----


def _open_worksheet(open_callable: Callable[[], Any], error_context: str) -> Any | None:
    """Try to open a worksheet via a callable; catch gspread errors and print a consistent message."""
    try:
        return google_api.get_worksheet(open_callable())
    except (
        gspread.SpreadsheetNotFound,
        gspread.exceptions.APIError,
        gspread.exceptions.GSpreadException,
    ) as e:
        utils.error_print(f"Could not open spreadsheet {error_context}: {e}")
        return None


def open_spreadsheet_by_url(
    gspread_client: Any, url: str, *, use_spinner: bool = False
) -> Any | None:
    """Open a spreadsheet by URL."""

    def open_fn() -> Any | None:
        return _open_worksheet(lambda: gspread_client.open_by_url(url), "by URL")

    if use_spinner:
        return utils.run_with_spinner("Opening document by URL...", open_fn)
    return open_fn()


def open_spreadsheet_by_index(
    gspread_client: Any, doc_list: list[str], index: int, *, use_spinner: bool = False
) -> Any | None:
    """Open a spreadsheet by 1-based index number from the document list."""
    if index < 1 or index > len(doc_list):
        utils.error_print(
            f"Invalid index {index}. Must be between 1 and {len(doc_list)}"
        )
        return None
    doc_name = doc_list[index - 1].strip()
    if not use_spinner:
        utils.standard_print(f"Opening document: {doc_name}")

    def open_fn() -> Any | None:
        return _open_worksheet(
            lambda: gspread_client.open(doc_name), f"at index {index}"
        )

    if use_spinner:
        return utils.run_with_spinner(f"Opening document: {doc_name}...", open_fn)
    return open_fn()


def open_spreadsheet_by_name(
    gspread_client: Any,
    doc_list: list[str],
    name: str,
    *,
    use_spinner: bool = False,
    prompt_prefix: str = "No exact match found. Did you mean",
) -> Any | None:
    """Open a spreadsheet by name search against the document list."""
    chosen_index = google_api.find_spreadsheet_by_name(name, doc_list)
    if chosen_index < 0:
        suggestion = utils.suggest_close_match(
            name,
            doc_list,
            prompt_prefix=prompt_prefix,
        )
        if suggestion is None:
            return None
        chosen_index = doc_list.index(suggestion)
    matched_name = doc_list[chosen_index].strip()
    if not use_spinner:
        utils.standard_print(f"Opening document: {matched_name}")

    def open_fn() -> Any | None:
        return _open_worksheet(lambda: gspread_client.open(matched_name), f"'{name}'")

    if use_spinner:
        return utils.run_with_spinner(f"Opening document: {matched_name}...", open_fn)
    return open_fn()


def _handle_spreadsheet_command(
    gspread_client: Any, doc_list: list[str], input_name: str
) -> Any | None:
    """Handle one spreadsheet selection command. Returns worksheet when one was opened, None to show prompt again."""
    if not input_name:
        return None
    # Handle 'excel' for local .xlsx
    if input_name.strip().lower() == config.COMMAND_EXCEL:
        import excel_io

        return excel_io.select_excel_file()
    # Handle URL
    if input_name.startswith(config.COMMAND_HTTP_PREFIX):
        return open_spreadsheet_by_url(gspread_client, input_name, use_spinner=True)
    # Handle 'all' command
    if input_name.startswith(config.COMMAND_LIST_ALL):
        utils.info_print("Available documents:")
        for i, doc in enumerate(doc_list):
            utils.info_print(f"{i + 1}. {doc.strip()}")
        return None
    # Handle 'new' command
    if input_name.startswith(config.COMMAND_LIST_NEW):
        utils.info_print("Newest documents: (modified or opened most recently)")
        for i in range(min(config.NUM_NEWEST_DOCS_TO_SHOW, len(doc_list))):
            utils.info_print(f"{i + 1}. {doc_list[i].strip()}")
        return None
    # Handle 'last' command
    if input_name.startswith(config.COMMAND_OPEN_LAST):
        # doc_list is already the get_all_spreadsheets() result (newest first,
        # same source the 'new' command reads); reuse it rather than paying
        # another rate-limited Google Sheets round-trip for data in hand.
        latest_spreadsheet_name = doc_list[0]
        return open_spreadsheet_by_name(
            gspread_client, doc_list, latest_spreadsheet_name, use_spinner=True
        )
    # Handle numeric index
    if input_name[0].isdigit():
        return open_spreadsheet_by_index(
            gspread_client, doc_list, int(input_name), use_spinner=True
        )
    # Handle 'settings' command
    if input_name.startswith(config.COMMAND_SETTINGS):
        utils.set_program_settings()
        return None
    # Handle name search
    return open_spreadsheet_by_name(
        gspread_client, doc_list, input_name, use_spinner=True
    )


def select_spreadsheet(gspread_client: Any, doc_list: list[str]) -> Any:
    """Interactive spreadsheet selection. Returns the selected worksheet."""
    if utils.NO_INPUT_MODE:
        utils.error_print(
            "Spreadsheet selection requires -s in non-interactive mode.",
            [
                "Pass -s <name|url|index|./file.xlsx>",
                "Or omit --no-input to select interactively.",
            ],
        )
        sys.exit(2)
    consecutive_open_failures = 0
    utils.print_mode_heading("Spreadsheet selection", "mode.spreadsheet")

    while True:
        try:
            input_name = utils.read_user_input(
                f"\nPlease enter the index, name, URL, or '{config.COMMAND_EXCEL}' for local file "
                f"('{config.COMMAND_LIST_ALL}' for list, '{config.COMMAND_LIST_NEW}' for list of newest, "
                f"'{config.COMMAND_OPEN_LAST}' to immediately open latest, '{config.COMMAND_SETTINGS}' to change settings):\n>> "
            )
        except utils.TopToSpreadsheet:
            # Already at spreadsheet selection; bubble up so main can restart selection loop.
            raise
        except utils.BackToModeSelection:
            # From this context, going "back" is equivalent to going to spreadsheet selection.
            raise utils.TopToSpreadsheet()
        try:
            worksheet = _handle_spreadsheet_command(
                gspread_client, doc_list, input_name
            )
            if worksheet is not None:
                return worksheet
        except (
            gspread.SpreadsheetNotFound,
            gspread.exceptions.APIError,
            gspread.exceptions.GSpreadException,
        ) as e:
            consecutive_open_failures += 1
            if consecutive_open_failures == 1:
                utils.error_print(
                    f"Could not access spreadsheet: {e}",
                    [
                        f"Please try again. Type '{config.COMMAND_LIST_ALL}' to see available documents."
                    ],
                )
            elif consecutive_open_failures == 2:
                utils.error_print(
                    "Spreadsheet not found or not accessible.",
                    [
                        "Common causes:",
                        "  - The spreadsheet name is misspelled",
                        "  - The spreadsheet hasn't been shared with your service account",
                        "    (Share it with the email in credentials.json 'client_email' field)",
                        "  - The spreadsheet doesn't contain any worksheets",
                        "",
                        f"  Type '{config.COMMAND_LIST_ALL}' to see accessible documents, or '{config.COMMAND_LIST_NEW}' for recent ones.",
                    ],
                )
            else:
                utils.error_print(
                    str(e),
                    [
                        f"Tip: Use the document index number (1, 2, 3...) from the '{config.COMMAND_LIST_ALL}' list."
                    ],
                )
        except (utils.QuitProgram, utils.TopToSpreadsheet, utils.BackToModeSelection):
            raise
        except Exception as e:
            utils.error_print(f"Could not open document: {e}")


# ---- Shared helpers ----

_SELECTION_MODE_HELP = [
    "    b or batch   - Generate all clips in the spreadsheet",
    "    r or range   - Generate clips from a range of rows",
    "    c or category - Generate clips by category",
    "    l or line    - Generate clips from specific line(s)",
    "    ce or cell   - Generate clips from specific cell(s) (e.g., P01.11)",
    "    p or participant - Generate all clips for one participant",
    "    k or keyword - Generate only annotated clips/timestamps (e.g., !key)",
    "    sv or severity - Generate clips by severity level",
]

_ALL_MODE_HELP = _SELECTION_MODE_HELP + [
    "    s or screen  - Generate screenshots (.png)",
    "    g or gif     - Generate GIFs (.gif)",
    "    re or reel   - Combine selectors into one highlight reel video",
    "    rl or reellate - Combine existing clips into a highlight reel",
    "    gv or gallery - Generate gallery from interval screenshots/GIFs of a video",
    "    br or browse - Browse spreadsheet rows interactively",
]


def _resolve_unrecognized_input(
    worksheet: Any, user_input: str, *, help_lines: list[str]
) -> list[ClipRecord] | None:
    """Try auto-detection and mixed-selector parsing for input that didn't match a mode alias.

    Attempts, in order: single-type auto-detection (line/range/cell/participant)
    then mixed-selector parsing. Prints appropriate messages.
    Returns clip list on success, None to signal the caller should re-prompt.
    """
    detected_mode, detected_kwargs = spreadsheet.detect_mode_from_input(user_input)
    if detected_mode:
        utils.standard_print(f"  {detected_mode.capitalize()} mode detected.")
        return spreadsheet.generate_list(worksheet, detected_mode, **detected_kwargs)

    parsed = spreadsheet.parse_reel_input(user_input)
    selector_types = [
        ("batch", bool(parsed.get("batch"))),
        ("keyword", bool(parsed.get("keyword"))),
        ("lines", len(parsed["lines"]) > 0),
        ("ranges", len(parsed["ranges"]) > 0),
        ("cells", len(parsed["cells"]) > 0),
        ("participants", len(parsed["participants"]) > 0),
        ("categories", len(parsed["categories"]) > 0),
    ]
    non_empty_types = [name for name, present in selector_types if present]
    has_chronologic = bool(parsed.get("chronologic"))

    if not non_empty_types:
        utils.info_print(f"  Unknown mode or input '{user_input}'. Available modes:")
        for line in help_lines:
            utils.info_print(line)
        return None
    if has_chronologic:
        utils.info_print(
            "  Chronologic selector is only supported for reel/chronologic modes."
        )
        utils.info_print(
            "  Use 're' or 'reel' for a combined reel video, or -T on the command line."
        )
        return None

    selector_summary = ", ".join(non_empty_types)
    utils.standard_print(
        f"  Mixed selectors detected ({selector_summary}). Generating from combined selectors."
    )
    return spreadsheet.generate_list(worksheet, "reel", reel_input=user_input)


def _run_standard_mode(mode: str, worksheet: Any) -> list[ClipRecord] | None:
    """Run a standard interactive mode (batch/line/range/category/cell/participant/keyword).

    Prompts the user for mode-specific input, then generates clips.
    Returns clip list on success, None if the user cancels or context fails.
    """
    utils.print_mode_heading(f"{mode.capitalize()} mode", f"mode.{mode}")
    ctx = spreadsheet.build_sheet_context(worksheet)
    if ctx is None:
        return None
    prompt_fn, gen_fn = _STANDARD_MODES[mode]
    result = prompt_fn(ctx)
    if result is None or result is False:
        return None
    return gen_fn(ctx, result)


def _print_run_summary(message: str) -> None:
    """Print a run summary block with newlines and a Summary header."""
    utils.info_print("")
    utils.print_mode_heading("Summary", "mode.selection")
    utils.info_print(message)
    utils.info_print("")


# ---- Clip processing pipeline (moved to pipeline.py) ----

# ---- Interactive mode flows ----


def _print_reencoding_warning(printer: Callable[[str], None]) -> None:
    """Print a reminder about trade-offs when ffmpeg runs without re-encoding."""
    printer(
        "* No re-encoding done, expect:\n- inaccurate start and end timings\n- lossy frames until first keyframe\n- bad timecodes at the end\n"
    )


def _print_completion_message(
    outputs_generated: int, output_format: str, is_reel: bool
) -> None:
    """Print a summary of generated outputs tailored to format and reel mode."""
    output_dir = utils.get_effective_output_dir()
    if is_reel:
        _print_run_summary(f"All done, created 1 reel!\nFiles are in {output_dir}")
        return
    noun = {"screen": "screenshots", "gif": "GIFs"}.get(output_format, "videos")
    _print_run_summary(
        f"All done, created {outputs_generated} {noun}!\nFiles are in {output_dir}"
    )


def _prompt_chronologic_participant_selection(
    worksheet: Any,
) -> tuple[str | None, spreadsheet.SheetContext | None]:
    """Prompt user to pick exactly one participant for chronologic reels.

    Returns (selected_participant_id, ctx). The context built here is handed
    back so the caller can pass it to generate_list and avoid a second fetch.
    """
    ctx = spreadsheet.build_sheet_context(worksheet)
    if ctx is None:
        return None, None

    available_list = spreadsheet.get_participant_list(
        ctx.header_row, ctx.id_cell, ctx.num_participants
    )
    if not available_list:
        utils.info_print("No participants found in the spreadsheet.")
        return None, None

    utils.print_mode_heading("Chronologic participant", "mode.chronologic")
    utils.info_print("Chronologic mode requires exactly one participant.")
    utils.info_print("Available participants:")
    for i, pid in enumerate(available_list, 1):
        utils.info_print(f"  {i}. {pid}")

    while True:
        selection = utils.read_user_input(
            "\nEnter one participant number or ID (e.g., 1 or P01):\n>> "
        )
        if not selection:
            utils.info_print("Please enter one participant.")
            continue
        tokens = spreadsheet.parse_participant_selection(selection)
        if len(tokens) != 1:
            utils.info_print("Please provide exactly one participant.")
            continue

        token = tokens[0]
        if token.isdigit():
            idx = int(token)
            if 1 <= idx <= len(available_list):
                return available_list[idx - 1], ctx
            utils.info_print(
                f"Not found: {token}. Available: {', '.join(available_list)}"
            )
            continue

        col_idx = spreadsheet.find_participant_column(
            ctx.header_row, ctx.id_cell, token
        )
        if col_idx is None:
            utils.info_print(
                f"Not found: {token}. Available: {', '.join(available_list)}"
            )
            continue
        if col_idx < len(ctx.header_row):
            return utils.normalize_participant_id(ctx.header_row[col_idx]), ctx
        return token, ctx


def _run_reel_mode_interactive(
    worksheet: Any,
) -> tuple[list[ClipRecord], bool, str | None]:
    """Run reel mode UI: instructions, input, generate_list, preview, confirm, output filename.
    Returns (clips_list, True, reel_output_file or None) when user confirms; caller may loop on continue.
    """
    utils.print_mode_heading("Reel mode", "mode.reel")

    utils.info_print("Combine selectors into one video. Syntax:")
    utils.info_print("  batch                    - all clips")
    utils.info_print("  keyword                  - annotated clips only")
    utils.info_print(
        "  chronologic              - chronological reel (requires exactly one participant)"
    )
    utils.info_print(
        "  severity                 - order reel by severity (most severe first)"
    )
    utils.info_print(
        "  highlights               - auto-select best clips within time budget"
    )
    utils.info_print("  11, 12, 13-16, 18        - lines and ranges")
    utils.info_print('  "Observations", "Onboarding" - categories (quoted)')
    utils.info_print("  P01.11, P02.15           - cells (participant.row)")
    utils.info_print("  P01, P02                 - participants (all their clips)")
    utils.info_print('  Example: chronologic, P01, 11, 13-16, "Observations"')
    reel_input = utils.read_user_input(
        "\nEnter reel selectors (combine any of the above, comma-separated):\n>> "
    )
    if not reel_input:
        utils.info_print("No input. Skipping reel.")
        return ([], False, None)

    # Reused by generate_list below when the chronologic prompt builds it; stays
    # None for every other reel path (generate_list then fetches once itself).
    reel_ctx: spreadsheet.SheetContext | None = None
    parsed_reel = spreadsheet.parse_reel_input(reel_input)
    if parsed_reel.get("highlights") and (
        parsed_reel["severity"] or parsed_reel["chronologic"]
    ):
        utils.error_print(
            "Cannot combine highlights with severity or chronologic ordering.",
            ["Use highlights on its own for auto-ranked highlight reel."],
        )
        return ([], False, None)
    if parsed_reel["severity"] and parsed_reel["chronologic"]:
        utils.error_print(
            "Cannot combine severity and chronologic ordering.",
            [
                "Use one ordering at a time: either chronologic (chronological) or severity."
            ],
        )
        return ([], False, None)
    if parsed_reel["chronologic"]:
        if len(parsed_reel["participants"]) > 1:
            utils.error_print(
                "Chronologic selector supports only one participant.",
                ["Please provide exactly one participant (e.g., chronologic, P01)."],
            )
            return ([], False, None)
        if len(parsed_reel["participants"]) == 0:
            selected_pid, reel_ctx = _prompt_chronologic_participant_selection(
                worksheet
            )
            if not selected_pid:
                return ([], False, None)
            reel_input = f"{reel_input}, {selected_pid}"
            parsed_reel["participants"] = [selected_pid]

    clips_list = spreadsheet.generate_list(
        worksheet, "reel", ctx=reel_ctx, reel_input=reel_input
    )
    if not clips_list:
        utils.info_print("No clips matched. Try different selectors.")
        return ([], False, None)
    utils.info_print(
        f"Preview: {len(clips_list)} clip(s) will be included (deduplicated by cell)."
    )
    for i, clip in enumerate(clips_list[: config.REEL_PREVIEW_CLIP_COUNT]):
        desc = (clip.get("desc") or "")[: config.DESCRIPTION_PREVIEW_LENGTH]
        utils.info_print(
            f"  {i + 1}. [{clip.get('category', '')}] {clip.get('participant', '')} row {clip['cell'].row}: {desc}..."
        )
    if len(clips_list) > config.REEL_PREVIEW_CLIP_COUNT:
        utils.info_print(
            f"  ... and {len(clips_list) - config.REEL_PREVIEW_CLIP_COUNT} more"
        )
    yn = utils.read_user_input("\nGenerate reel? [y/n]\n>> ")
    if yn.strip().lower() != "y":
        return ([], False, None)

    study_name = clips_list[0].get("study", "").strip() if clips_list else ""
    default_filename = (
        f"{study_name}_reel{config.FILEFORMAT}"
        if study_name
        else f"reel{config.FILEFORMAT}"
    )
    if parsed_reel["chronologic"] and parsed_reel["participants"]:
        chronologic_pid = utils.normalize_participant_id(
            parsed_reel["participants"][0]
        ).strip()
        if study_name and chronologic_pid:
            default_filename = (
                f"{study_name}_{chronologic_pid}_chronologic{config.FILEFORMAT}"
            )
        elif chronologic_pid:
            default_filename = f"{chronologic_pid}_chronologic{config.FILEFORMAT}"
        else:
            default_filename = f"chronologic{config.FILEFORMAT}"
    elif parsed_reel.get("highlights"):
        default_filename = (
            f"{study_name}_highlights{config.FILEFORMAT}"
            if study_name
            else f"highlights{config.FILEFORMAT}"
        )
    elif parsed_reel["severity"]:
        default_filename = (
            f"{study_name}_severity_reel{config.FILEFORMAT}"
            if study_name
            else f"severity_reel{config.FILEFORMAT}"
        )

    output_file = utils.read_user_input(
        f"\nOutput filename (Enter for default {default_filename}):\n>> "
    )
    reel_output_file = None
    if output_file:
        reel_output_file = (
            output_file
            if output_file.endswith(config.FILEFORMAT)
            else output_file + config.FILEFORMAT
        )
    else:
        reel_output_file = files.get_unique_filename(default_filename)
    return (clips_list, True, reel_output_file)


def _parse_clip_selection(selection_input: str, num_clips: int) -> list[int]:
    """Parse user selection input into list of clip indices.

    Supports formats: "A + B + C", "A, B, C", "A B C", or mixed.

    Args:
        selection_input: User's selection string (e.g., "A + B" or "A, C")
        num_clips: Total number of available clips (for validation)

    Returns:
        List of valid 0-based indices, deduplicated and in selection order
    """
    # Normalize separators to spaces
    normalized = selection_input.replace("+", " ").replace(",", " ")
    tokens = normalized.split()

    indices = []
    seen = set()
    for token in tokens:
        idx = utils.letter_to_index(token)
        if idx >= 0 and idx < num_clips and idx not in seen:
            indices.append(idx)
            seen.add(idx)
    return indices


def _run_reellate_mode_interactive() -> tuple[bool, str | None]:
    """Run reel-late mode UI: discover clips, display list, select, concatenate.

    Returns (True, output_file) when reel was generated; (False, None) otherwise.
    """
    clips = files.discover_clips()
    output_dir = utils.get_effective_output_dir()
    utils.print_mode_heading("Reel-late mode", "mode.reellate")

    if not clips:
        utils.info_print("No clips found in the output directory.")
        utils.info_print("  Source videos (like study_P01.mp4) are excluded.")
        utils.info_print(
            "  Generate some clips first, then use this mode to combine them."
        )
        return (False, None)

    utils.info_print("Combine existing clips into a highlight reel.")
    utils.info_print(f"Found {len(clips)} clip(s) in {output_dir}:")

    # Display indexed list
    for i, clip in enumerate(clips):
        letter = utils.index_to_letter(i)
        utils.info_print(f'  {letter}. "{clip}"')

    utils.info_print("Select clips to include (order preserved). Syntax:")
    utils.info_print("  A + B + C    - combine clips A, B, and C")
    utils.info_print("  A, B, C      - same as above")
    utils.info_print("  A B C        - same as above")

    selection_input = utils.read_user_input("\nEnter clip selection:\n>> ")
    if not selection_input:
        utils.info_print("No selection. Skipping reel.")
        return (False, None)

    indices = _parse_clip_selection(selection_input, len(clips))
    if not indices:
        utils.warning_print(
            "No valid clips selected.",
            ["Use letters from the list above (e.g., A + B + C)"],
        )
        return (False, None)

    selected_clips = [clips[i] for i in indices]

    # Preview selection
    utils.info_print(f"Selected {len(selected_clips)} clip(s):")
    for i, clip in enumerate(selected_clips):
        utils.info_print(f'  {i + 1}. "{clip}"')

    yn = utils.read_user_input("\nGenerate reel from these clips? [y/n]\n>> ").lower()
    if yn != "y":
        utils.info_print("Cancelled.")
        return (False, None)

    output_file = utils.read_user_input(
        '\nOutput filename (Enter for default "reel.mp4"):\n>> '
    )
    if not output_file:
        output_file = files.get_unique_filename(f"reel{config.FILEFORMAT}")
    elif not output_file.endswith(config.FILEFORMAT):
        output_file = output_file + config.FILEFORMAT

    resolved_clips = [str(utils.resolve_output_path(name)) for name in selected_clips]

    def _concat_reellate() -> bool:
        return video.concatenate_clips(
            resolved_clips, output_file, reencode_on_fail=True
        )

    ok = (
        utils.run_with_spinner(
            f"Concatenating {len(selected_clips)} clips into {output_file}...",
            _concat_reellate,
        )
        if utils.use_progress()
        else _concat_reellate()
    )

    if ok:
        return (True, output_file)
    return (False, None)


def _run_format_mode_interactive(worksheet: Any, output_format: str) -> None:
    """Run interactive flow for screen/gif output formats.

    Prompts the user for a timestamp-selection mode, generates a clip list, then renders
    screenshots/GIFs using the regular clip processing pipeline.
    """
    format_display_name = "Screenshot" if output_format == "screen" else "GIF"
    output_label = "screenshots" if output_format == "screen" else "GIFs"
    utils.print_mode_heading(f"{format_display_name} mode", "mode.format")

    utils.info_print("Choose how to select timestamps.")
    while True:
        selection = utils.read_user_input(
            "\nSelect source rows for this output:\n"
            "  Modes: (b)atch, (r)ange, (c)ategory, (l)ine, (ce)ll, (p)articipant, (k)eyword, (sv) severity\n"
            '  Or enter mixed selectors directly: e.g. 5, P01.11, 13-16, "Observations"\n>> '
        )
        if not selection:
            utils.info_print(
                "  Please enter a mode or direct input (e.g. P01.11, 5, 7, 13-16, P01)."
            )
            continue

        mode = FORMAT_MODE_ALIASES.get(selection.lower())
        if mode is None and len(selection.split()) == 1 and selection.strip():
            full_names = [k for k in FORMAT_MODE_ALIASES if len(k) > 2]
            suggestion = utils.suggest_close_match(selection.strip(), full_names)
            if suggestion is not None:
                mode = FORMAT_MODE_ALIASES[suggestion]
        if mode:
            clips_list = spreadsheet.generate_list(worksheet, mode)
            break

        clips_list = _resolve_unrecognized_input(
            worksheet, selection, help_lines=_SELECTION_MODE_HELP
        )
        if clips_list is not None:
            break

    outputs_generated, artifacts = process_clips(
        clips_list, output_format=output_format, include_severity=(mode == "severity")
    )
    if artifacts:
        viewer.INTERACTIVE_ARTIFACTS.extend(artifacts)
        if config.MANIFEST_ENABLED:
            viewer.save_manifest(
                viewer.INTERACTIVE_ARTIFACTS,
                new_reels=viewer.INTERACTIVE_REELS,
                study=artifacts[0].get("study", ""),
                worksheet_title=getattr(worksheet, "title", ""),
                is_excel=_is_excel_worksheet(worksheet),
                mode="interactive",
            )
    _print_run_summary(
        f"All done, created {outputs_generated} {output_label}!\nFiles are in {utils.get_effective_output_dir()}"
    )


def _run_viewer_mode(worksheet: Any) -> None:
    """Generate the timeline viewer from session artifacts, falling back to manifest."""
    artifacts = list(viewer.INTERACTIVE_ARTIFACTS)
    mode_label = "interactive"

    if not artifacts:
        manifest_artifacts = viewer.load_manifest_artifacts()
        if not manifest_artifacts:
            utils.info_print(
                "No artifacts in this session and no manifest file found.\n"
                "Generate clips first, or run a prior session with --manifest to save one."
            )
            return
        count = len(manifest_artifacts)
        yn = utils.read_user_input(
            f"No artifacts in this session, but found {count} artifact(s) in manifest.\n"
            "Generate viewer from manifest? [y/n]\n>> "
        )
        if yn.strip().lower() != "y":
            return
        artifacts = manifest_artifacts
        mode_label = "manifest"

    study = artifacts[0].get("study", "")
    participant = artifacts[0].get("participant", "")
    ss_events = viewer.load_screenspace_events_for_viewer()
    data = viewer.finalize_timeline_data(
        artifacts,
        study=study,
        participant=participant,
        worksheet_title=""
        if mode_label == "manifest"
        else getattr(worksheet, "title", ""),
        is_excel=False if mode_label == "manifest" else _is_excel_worksheet(worksheet),
        mode=mode_label,
        screenspace_events=ss_events or None,
    )
    viewer_path = viewer.generate_timeline_viewer(data)
    if viewer_path:
        utils.info_print(f"Timeline viewer created: {viewer_path}")


def _run_regenerate_mode() -> None:
    """Regenerate all media artifacts and reels from saved manifest."""
    existing_artifacts, existing_reels = viewer.load_manifest_both()
    if not existing_artifacts and not existing_reels:
        utils.info_print(
            "No manifest file found.\nGenerate clips first with --manifest to save one."
        )
        return
    media_count = sum(1 for a in existing_artifacts if a.get("type") != "transcript")
    reel_count = len(existing_reels)
    total = media_count + reel_count
    if total == 0:
        utils.info_print(
            "Manifest contains only transcript artifacts; nothing to regenerate."
        )
        return
    yn = utils.read_user_input(
        f"Found {media_count} media artifact(s) and {reel_count} reel(s) in manifest.\n"
        "Regenerate all? [y/n]\n>> "
    )
    if yn.strip().lower() != "y":
        return
    regenerated = regenerate_from_manifest(existing_artifacts, reels=existing_reels)
    utils.info_print(f"Regenerated {regenerated} of {total} item(s).")


def _run_gallery_mode_interactive() -> None:
    """Interactive gallery mode: select a video, generate interval captures, build gallery viewer."""
    utils.print_mode_heading("Gallery mode", "mode.gallery")

    input_dir = utils.get_effective_input_dir()
    videos = sorted(p for p in input_dir.glob(f"*{config.FILEFORMAT}") if p.is_file())
    if not videos:
        utils.error_print(
            f"No {config.FILEFORMAT} files found in {input_dir}.",
            ["Place a video file in the input directory and try again."],
        )
        return

    utils.info_print("Available videos:")
    for i, v in enumerate(videos, 1):
        size_mb = v.stat().st_size / 1_000_000
        utils.info_print(f"  {i}. {v.name}  ({size_mb:.0f} MB)")

    selection = utils.read_user_input("Select a video (number or filename):\n>> ")
    video_path: Path | None = None
    try:
        idx = int(selection) - 1
        if 0 <= idx < len(videos):
            video_path = videos[idx]
    except ValueError:
        for v in videos:
            if v.name == selection.strip() or v.stem == selection.strip():
                video_path = v
                break

    if video_path is None:
        utils.error_print(f"Could not find video matching '{selection}'.")
        return

    fmt_input = (
        utils.read_user_input("Generate (s)creenshots or (g)ifs? [s]\n>> ")
        .strip()
        .lower()
    )
    output_format = "gif" if fmt_input in ("g", "gif") else "screen"

    interval_input = utils.read_user_input(
        f"Capture interval in seconds [{config.GALLERY_INTERVAL_SECONDS}]:\n>> "
    ).strip()
    try:
        interval = (
            int(interval_input) if interval_input else config.GALLERY_INTERVAL_SECONDS
        )
    except ValueError:
        interval = config.GALLERY_INTERVAL_SECONDS
    if interval <= 0:
        interval = config.GALLERY_INTERVAL_SECONDS

    gif_duration = config.GALLERY_GIF_DURATION_SECONDS
    if output_format == "gif":
        dur_input = utils.read_user_input(
            f"GIF duration in seconds [{config.GALLERY_GIF_DURATION_SECONDS}]:\n>> "
        ).strip()
        try:
            gif_duration = (
                int(dur_input) if dur_input else config.GALLERY_GIF_DURATION_SECONDS
            )
        except ValueError:
            gif_duration = config.GALLERY_GIF_DURATION_SECONDS
        if gif_duration <= 0:
            gif_duration = config.GALLERY_GIF_DURATION_SECONDS

    artifacts = video.generate_interval_captures(
        str(video_path),
        interval_seconds=interval,
        output_format=output_format,
        gif_duration_seconds=gif_duration,
    )
    if not artifacts:
        return

    duration = video.get_file_duration(str(video_path)) or 0
    data = viewer.finalize_gallery_data(
        artifacts,
        source_video=video_path.name,
        video_duration=duration,
        output_format=output_format,
        interval=interval,
    )
    gallery_path = viewer.generate_gallery_viewer(data)
    if gallery_path:
        utils.info_print(f"Gallery viewer created: {gallery_path}")


def _dispatch_interactive_mode(
    mode: str | None,
    worksheet: Any,
    raw_input: str,
    gspread_client: Any = None,
) -> tuple[list[ClipRecord], bool, str | None] | None:
    """Dispatch a resolved mode or raw input. Returns result tuple, or None to re-prompt."""
    # Special modes with their own interactive flows
    if mode == "browse":

        def _browse_process_fn(clips_list, output_format):
            outputs_generated, artifacts = process_clips(
                clips_list, output_format=output_format
            )
            if artifacts:
                viewer.INTERACTIVE_ARTIFACTS.extend(artifacts)
                if config.MANIFEST_ENABLED:
                    viewer.save_manifest(
                        viewer.INTERACTIVE_ARTIFACTS,
                        new_reels=viewer.INTERACTIVE_REELS,
                        study=artifacts[0].get("study", ""),
                        worksheet_title=getattr(worksheet, "title", ""),
                        is_excel=_is_excel_worksheet(worksheet),
                        mode="interactive",
                    )
            if not config.REENCODING:
                _print_reencoding_warning(utils.info_print)
            return (outputs_generated, artifacts)

        interactive.browse_spreadsheet(worksheet, process_fn=_browse_process_fn)
        return ([], False, None)
    if mode == "viewer":
        _run_viewer_mode(worksheet)
        return None
    if mode == "regenerate":
        _run_regenerate_mode()
        return None
    if mode == "gallery":
        _run_gallery_mode_interactive()
        return None
    if mode == "studio":
        import server

        server.start_combined_server(
            worksheet=worksheet,
            default_page="studio",
            gspread_client=gspread_client,
        )
        return None
    if mode == "screenspace":
        import server

        server.start_combined_server(
            worksheet=worksheet,
            default_page="screenspace",
            gspread_client=gspread_client,
        )
        return None
    if mode == "transcripts":
        import server

        server.start_combined_server(
            worksheet=worksheet,
            default_page="transcripts",
            gspread_client=gspread_client,
        )
        return None
    if mode == "export":
        import data_export

        data_export.run_cli_export()
        return ([], False, None)
    if mode == "timeline-viewer":
        clips_list = spreadsheet.generate_list(worksheet, "batch", skip_prompts=True)
        outputs_generated, artifacts = process_clips(clips_list, output_format="clip")
        if not config.REENCODING:
            _print_reencoding_warning(utils.info_print)
        _print_completion_message(outputs_generated, "clip", is_reel=False)
        if artifacts:
            viewer.INTERACTIVE_ARTIFACTS.extend(artifacts)
            if config.MANIFEST_ENABLED:
                viewer.save_manifest(
                    viewer.INTERACTIVE_ARTIFACTS,
                    new_reels=viewer.INTERACTIVE_REELS,
                    study=artifacts[0].get("study", ""),
                    worksheet_title=getattr(worksheet, "title", ""),
                    is_excel=_is_excel_worksheet(worksheet),
                    mode="timeline-viewer",
                )
            study = artifacts[0].get("study", "")
            ss_events = viewer.load_screenspace_events_for_viewer()
            data = viewer.finalize_timeline_data(
                artifacts,
                study=study,
                worksheet_title=getattr(worksheet, "title", ""),
                is_excel=_is_excel_worksheet(worksheet),
                mode="timeline-viewer",
                screenspace_events=ss_events or None,
            )
            viewer_path = viewer.generate_timeline_viewer(
                data,
                template_name="timeline-viewer.html",
                output_basename="timeline_viewer.html",
            )
            if viewer_path:
                utils.info_print(f"Participant timeline viewer created: {viewer_path}")
        else:
            utils.warning_print(
                "No artifacts were generated; skipping timeline viewer."
            )
        return None
    if mode == "settings":
        utils.set_program_settings()
        return None
    if mode == "reel":
        clips, confirmed, reel_file = _run_reel_mode_interactive(worksheet)
        return (clips, True, reel_file) if confirmed else None
    if mode == "reellate":
        success, output_file = _run_reellate_mode_interactive()
        if success:
            _print_run_summary(
                f"Reel created: {output_file}\nFiles are in {utils.get_effective_output_dir()}"
            )
            return ([], False, None)
        return None
    if mode in ("screen", "gif"):
        _run_format_mode_interactive(worksheet, mode)
        return ([], False, None)

    # Standard modes with interactive prompts
    if mode in _STANDARD_MODES:
        clips = _run_standard_mode(mode, worksheet)
        return (clips or [], False, None)
    if mode:
        return (spreadsheet.generate_list(worksheet, mode), False, None)

    # No alias match -- try auto-detection and mixed selectors
    clips = _resolve_unrecognized_input(worksheet, raw_input, help_lines=_ALL_MODE_HELP)
    if clips is None:
        return None
    return (clips, False, None)


def run_interactive_mode(worksheet: Any, gspread_client: Any = None) -> None:
    """Execute interactive mode - main processing loop."""
    if utils.NO_INPUT_MODE:
        utils.error_print(
            "No CLI mode flag provided; cannot enter interactive mode under --no-input.",
            [
                "Pass one of: -b, -l, -r, -c, -p, -k, -S, -M, -R, -T, -H, --highlights",
                "Or a UI flag: --studio, --screenspace, --transcripts",
            ],
        )
        sys.exit(2)
    viewer.INTERACTIVE_ARTIFACTS.clear()
    viewer.INTERACTIVE_REELS.clear()

    while True:
        try:
            # Mode selection (inlined from select_mode_and_generate)
            utils.print_mode_heading("Mode selection", "mode.selection")
            input_mode = utils.read_user_input(
                "\nEnter mode or input directly:\n"
                "  Tools: (s)creen, (g)if, (re)el, (rl) reel-late, (rg) regenerate, (se)ttings \n"
                "  Front: (st) studio, (ss) screenspace, (tr)anscripts, (br)owse \n"
                "  Packs: (v)iewer, (tv) timeline-viewer, (gv) gallery, (ex)port \n"
                "  Modes: (b)atch, (r)ange, (c)ategory, (l)ine, (ce)ll, (p)articipant, (k)eyword, (sv) severity \n"
                '  Or enter mixed selectors directly: e.g. 5, P01.11, 13-16, "Observations"\n>> '
            )
            if not input_mode:
                utils.info_print(
                    "  Please enter a mode or direct input (e.g. P01.11, 5, 7, 13-16, P01)."
                )
                continue

            resolved_mode = MODE_ALIASES.get(input_mode.strip().lower())
            if (
                resolved_mode is None
                and len(input_mode.split()) == 1
                and input_mode.strip()
            ):
                full_mode_names = [k for k in MODE_ALIASES if len(k) > 2]
                suggestion = utils.suggest_close_match(
                    input_mode.strip(), full_mode_names
                )
                if suggestion is not None:
                    resolved_mode = MODE_ALIASES[suggestion]
            result = _dispatch_interactive_mode(
                resolved_mode, worksheet, input_mode, gspread_client=gspread_client
            )
            if result is None:
                continue
            clips_list, is_reel, reel_output_file = result

            if not clips_list and not is_reel:
                yn = utils.read_user_input(
                    "Continue working (y) or quit the program (n)? [y/n]\n>> "
                )
                if yn.strip().lower() == "n":
                    break
                continue
            if is_reel:
                outputs_generated, reel_records = process_reel(
                    clips_list,
                    output_file=reel_output_file,
                )
                if reel_records:
                    viewer.INTERACTIVE_REELS.extend(reel_records)
                    if config.MANIFEST_ENABLED:
                        viewer.save_manifest(
                            viewer.INTERACTIVE_ARTIFACTS,
                            new_reels=viewer.INTERACTIVE_REELS,
                            study=reel_records[0].get("study", ""),
                            worksheet_title=getattr(worksheet, "title", ""),
                            is_excel=_is_excel_worksheet(worksheet),
                            mode="interactive",
                        )
            else:
                outputs_generated, artifacts = process_clips(
                    clips_list, include_severity=(resolved_mode == "severity")
                )
                if artifacts:
                    viewer.INTERACTIVE_ARTIFACTS.extend(artifacts)
                    if config.MANIFEST_ENABLED:
                        viewer.save_manifest(
                            viewer.INTERACTIVE_ARTIFACTS,
                            new_reels=viewer.INTERACTIVE_REELS,
                            study=artifacts[0].get("study", ""),
                            worksheet_title=getattr(worksheet, "title", ""),
                            is_excel=_is_excel_worksheet(worksheet),
                            mode="interactive",
                        )

            if not config.REENCODING:
                _print_reencoding_warning(utils.info_print)
            _print_completion_message(
                outputs_generated, output_format="clip", is_reel=is_reel
            )

            yn = utils.read_user_input(
                "Continue working (y) or quit the program (n)? [y/n]\n>> "
            )
            if yn.strip().lower() == "n":
                break
        except gspread.exceptions.GSpreadException as e:
            utils.error_print(f"Google Sheets API error: {e}")
            utils.debug_print(f"ERROR Message '{e}', Attempting reconnect")
        except utils.TopToSpreadsheet:
            # Escalate to main to trigger spreadsheet reselection.
            raise
        except utils.BackToModeSelection:
            # Return to main mode selection prompt.
            continue
        except utils.QuitProgram:
            break


if __name__ == "__main__":
    from cli import main

    try:
        main()
    except KeyboardInterrupt:
        utils.info_print("Interrupted by user")
        sys.exit(0)
