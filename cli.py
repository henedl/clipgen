# -*- coding: utf-8 -*-
"""CLI entry point and argument parsing for clipgen.

Handles command-line argument parsing, CLI mode detection, setup,
and CLI-specific dispatch. The main() function is the program entry point,
called from clipgen.py's __main__ guard.
"""

import argparse
import io
import os
import sys
from pathlib import Path
from typing import Any, List, NamedTuple, Optional, Tuple

import gspread
from icecream import ic

import clipgen
import config
import excel_io
import files
import google_api
import spreadsheet
import utils
import video
import viewer
from utils import ClipRecord


# ---- CLI data structures ----


class CliModeArgs(NamedTuple):
    line_numbers: Optional[List[int]]
    range_start: Optional[int]
    range_end: Optional[int]
    cell_specs: Optional[List[Tuple[str, int]]]


# ---- Argument parsing ----


def parse_arguments() -> argparse.Namespace:
    """Parse command-line arguments for non-interactive mode.

    Exactly one of the mode flags (-b, -l, -r, -C, -c, -p, -f, -M, -R, -T) may be given;
    if none is given, the program runs in interactive mode. Optional flags (-s, -y,
    -v, --screen, --gif) may be combined with any mode.

    Returns:
        argparse.Namespace with attributes: batch, lines, range, category, cell,
        participant, filter, mixed, reel, timeline (mode flags/values),
        spreadsheet, yes, verbose, screen, gif, input, output.
    """
    parser = argparse.ArgumentParser(
        description="clipgen - Video clip generator from Google Sheets timestamps.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python clipgen.py                        Interactive mode (default)
  python clipgen.py -b                     Batch mode - generate all clips
  python clipgen.py -C "Observations"      Category mode - rows with category "Observations"
  python clipgen.py -C "Observations,Onboarding"
                                           Category mode - multiple categories
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
  python clipgen.py --timeline-viewer      Generate per-participant timeline viewer

Note: Non-interactive mode (using -b, -l, -r, -C, -c, -p, -f, -M, -R, or -T) is silent by default,
      only showing errors and the final summary. Use -v for full output.
""",
    )

    # Mode arguments: only one of -b/-l/-r/-C/-c/-p/-f/-M/-R/-T may be set at a time
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument(
        "-b",
        "--batch",
        action="store_true",
        help="Batch mode: generate all possible clips",
    )
    mode_group.add_argument(
        "-l",
        "--lines",
        type=str,
        metavar="LINES",
        help="Line mode: specify line numbers separated by + or , (e.g., 1+4+5 or 1,4,5)",
    )
    mode_group.add_argument(
        "-r",
        "--range",
        type=str,
        metavar="RANGE",
        help="Range mode: specify start-end line range (e.g., 1-10)",
    )
    mode_group.add_argument(
        "-C",
        "--category",
        type=str,
        metavar="CATEGORIES",
        help='Category mode: specify one or more category names (comma- or plus-separated, e.g., "Observations,Onboarding")',
    )
    mode_group.add_argument(
        "-c",
        "--cell",
        type=str,
        metavar="CELLS",
        help="Cell mode: specify cells as participant.row (e.g., P01.11 or P01.11 + P03.11)",
    )
    mode_group.add_argument(
        "-p",
        "--participant",
        type=str,
        metavar="ID",
        help="Participant mode: generate all clips for one or more participants (e.g., P01 or P01,P03)",
    )
    mode_group.add_argument(
        "-f",
        "--filter",
        action="store_true",
        help="Filter mode: generate only key-marked clips/timestamps",
    )
    mode_group.add_argument(
        "-M",
        "--mixed",
        type=str,
        metavar="SELECTORS",
        help='Mixed mode: combine selectors (e.g. "5, P01.11, 13-16") for individual clips/screenshots/GIFs',
    )
    mode_group.add_argument(
        "-R",
        "--reel",
        type=str,
        metavar="SELECTORS",
        help='Reel mode: combine selectors (e.g. "11, 13-16, P01, \\"Observations\\"") into one video',
    )
    mode_group.add_argument(
        "-T",
        "--timeline",
        type=str,
        metavar="PARTICIPANT",
        help="Timeline mode: chronological reel for one participant (e.g., P01)",
    )

    format_group = parser.add_mutually_exclusive_group()
    format_group.add_argument(
        "--screen",
        action="store_true",
        help="Output screenshots (.png) instead of video clips",
    )
    format_group.add_argument(
        "--gif", action="store_true", help="Output animated GIFs instead of video clips"
    )

    # Transcription arguments
    parser.add_argument(
        "--transcribe",
        action="store_true",
        help="Generate transcript files alongside artifacts",
    )
    parser.add_argument(
        "--transcript-format",
        type=str,
        choices=["md", "srt", "vtt"],
        metavar="FMT",
        help="Transcript format: md (default), srt, or vtt",
    )

    # Optional arguments (can be used with any mode)
    parser.add_argument(
        "-s",
        "--spreadsheet",
        type=str,
        metavar="NAME",
        help="Spreadsheet name, URL, or index number",
    )
    parser.add_argument(
        "-y",
        "--yes",
        action="store_true",
        help="Skip confirmation prompts (auto-confirm)",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Increase verbosity (-v = verbose output; default is quiet in CLI, standard in interactive mode)",
    )
    parser.add_argument(
        "--viewer",
        action="store_true",
        help="Generate a timeline HTML viewer file (clips_viewer.html) for this run",
    )
    parser.add_argument(
        "--manifest",
        action="store_true",
        help="Write artifact metadata to a cumulative manifest JSON file; combine with --viewer to regenerate viewer from manifest",
    )
    parser.add_argument(
        "--timeline-viewer",
        action="store_true",
        help="Batch-export all clips and generate a per-participant timeline HTML viewer",
    )
    parser.add_argument(
        "-i",
        "--input",
        type=str,
        metavar="DIR",
        help="Input directory where source videos are located (defaults to current working directory when unset)",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        metavar="DIR",
        help="Output directory where generated artifacts will be written (defaults to current working directory when unset)",
    )

    titlecard_group = parser.add_mutually_exclusive_group()
    titlecard_group.add_argument(
        "--titlecards",
        dest="titlecards",
        action="store_true",
        help="Enable titlecards for generated video clips for this run",
    )
    titlecard_group.add_argument(
        "--no-titlecards",
        dest="titlecards",
        action="store_false",
        help="Disable titlecards for generated video clips for this run",
    )

    parser.set_defaults(titlecards=None)

    return parser.parse_args()


def parse_cli_mode_args(args: Any) -> CliModeArgs:
    """Parse CLI arguments for line, range, and cell modes.

    Args:
        args: Parsed command-line arguments

    Returns:
        Parsed mode argument values as CliModeArgs
    """
    cli_line_numbers = None
    cli_range_start = None
    cli_range_end = None
    cli_cell_specs = None

    if args.lines:
        try:
            # Support both + and , as separators
            line_str = args.lines.replace(",", "+")
            cli_line_numbers = [int(num.strip()) for num in line_str.split("+")]
        except ValueError:
            utils.error_print(
                f'Invalid line numbers "{args.lines}". Use format: 1+4+5 or 1,4,5'
            )
            sys.exit(1)

    if args.range:
        try:
            parts = args.range.split("-")
            if len(parts) != 2:
                raise ValueError("Range must have exactly two parts")
            cli_range_start = int(parts[0].strip())
            cli_range_end = int(parts[1].strip())
            if cli_range_start > cli_range_end:
                utils.error_print(
                    f"Range start ({cli_range_start}) must be less than or equal to end ({cli_range_end})"
                )
                sys.exit(1)
        except ValueError:
            utils.error_print(f'Invalid range "{args.range}". Use format: 1-10')
            sys.exit(1)

    if args.cell:
        try:
            cli_cell_specs = spreadsheet.parse_cell_specifications(args.cell)
        except ValueError as e:
            utils.error_print(f"Invalid cell specification: {e}")
            sys.exit(1)

    return CliModeArgs(cli_line_numbers, cli_range_start, cli_range_end, cli_cell_specs)


# ---- Setup utilities ----


def setup_encoding() -> None:
    """Ensure UTF-8 encoding for stdout/stderr to handle unicode properly."""
    encoding = sys.stdout.encoding
    if not encoding or encoding.lower() != "utf-8":
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", errors="replace"
        )
        sys.stderr = io.TextIOWrapper(
            sys.stderr.buffer, encoding="utf-8", errors="replace"
        )


def get_runtime_working_dir() -> str:
    """Return the runtime working directory.

    Source runs use the script directory; frozen one-file builds use the
    executable directory so local assets resolve from where the binary lives.
    """
    if getattr(sys, "frozen", False):
        return str(Path(sys.executable).resolve().parent)
    return str(Path(__file__).resolve().parent)


# ---- Google authentication ----


def authenticate_google() -> Any:
    """Authenticate with Google Sheets API.

    Returns:
        Google client connection object
    """
    try:
        utils.debug_print("Attempting login...")
        gspread_client = gspread.oauth(credentials_filename="credentials.json")
        utils.debug_print("Login successful!")
        return gspread_client
    except gspread.exceptions.GSpreadException as e:
        utils.error_print(
            "Could not authenticate with Google.",
            [
                f"Error details: {e}",
                f"Credentials file location: {Path.cwd() / 'credentials.json'}",
                "",
                "Troubleshooting steps:",
                "  1. Ensure 'credentials.json' exists in the working directory",
                "  2. Verify the credentials file is valid JSON",
                "  3. Check that the service account has access to Google Sheets API",
                "  4. For OAuth flow, delete any existing token files and re-authenticate",
            ],
        )
        sys.exit(1)


# ---- Worksheet selection ----


def select_worksheet(
    gspread_client: Any, doc_list: List[str], args: Any, cli_mode: bool
) -> Any:
    """Select worksheet based on command-line arguments or interactive selection.

    Args:
        gspread_client: Google client connection
        doc_list: List of available spreadsheet names
        args: Parsed command-line arguments
        cli_mode: Whether running in CLI mode

    Returns:
        Worksheet object
    """
    worksheet = None
    if args.spreadsheet:
        # CLI-specified spreadsheet
        raw = args.spreadsheet.strip()
        raw_lower = raw.lower()
        if raw_lower == config.COMMAND_EXCEL:
            # -s excel: use single .xlsx in cwd, else error
            paths = excel_io.list_excel_in_cwd()
            if not paths:
                utils.error_print(
                    "No .xlsx files found in the current directory.",
                    [
                        "Place an Excel file (.xlsx) in the working directory or use -s path/to/file.xlsx"
                    ],
                )
                sys.exit(1)
            if len(paths) > 1:
                utils.error_print(
                    f"Multiple .xlsx files found ({len(paths)}). Specify one with -s path/to/file.xlsx",
                    [Path(p).name for p in paths],
                )
                sys.exit(1)
            worksheet = excel_io.open_excel_workbook(paths[0])
            if not worksheet:
                sys.exit(1)
        elif raw_lower.endswith(".xlsx"):
            # -s path/to/file.xlsx
            path = str(Path.cwd() / raw) if not Path(raw).is_absolute() else raw
            worksheet = excel_io.open_excel_workbook(path)
            if not worksheet:
                utils.error_print(f'Could not open Excel file "{args.spreadsheet}"')
                sys.exit(1)
        elif args.spreadsheet.startswith(config.COMMAND_HTTP_PREFIX):
            worksheet = clipgen.open_spreadsheet_by_url(
                gspread_client, args.spreadsheet
            )
        elif args.spreadsheet.isdigit():
            worksheet = clipgen.open_spreadsheet_by_index(
                gspread_client, doc_list, int(args.spreadsheet)
            )
        else:
            worksheet = clipgen.open_spreadsheet_by_name(
                gspread_client, doc_list, args.spreadsheet
            )

        if not worksheet:
            utils.error_print(
                f'Could not find or open spreadsheet "{args.spreadsheet}"'
            )
            sys.exit(1)
    else:
        # Auto-connect if working directory name matches a spreadsheet
        cwd_name = Path.cwd().name
        worksheet = clipgen.open_spreadsheet_by_name(
            gspread_client,
            doc_list,
            cwd_name,
            prompt_prefix=(
                "Tried matching current working directory to existing spreadsheets, "
                "but no exact match. \n\nDid you mean"
            ),
        )
        if worksheet:
            utils.standard_print(
                f"Auto-connecting to spreadsheet: {worksheet.spreadsheet.title}"
            )
        elif cli_mode:
            # CLI mode requires a spreadsheet - can't prompt interactively
            utils.error_print(
                "No spreadsheet found matching working directory name.",
                ["Use -s to specify a spreadsheet name, URL, or index."],
            )
            sys.exit(1)
        else:
            worksheet = clipgen.select_spreadsheet(gspread_client, doc_list)

    if worksheet and config.DEBUGGING:
        ic(worksheet.title)
    if clipgen._is_excel_worksheet(worksheet):
        utils.standard_print("Using local Excel file.")
    else:
        utils.standard_print("Connected to Google Drive!")
    return worksheet


# ---- CLI clip generation ----


def _generate_cli_clips(
    worksheet: Any,
    args: Any,
    cli_mode_args: CliModeArgs,
) -> List[ClipRecord]:
    """Resolve CLI arguments into a list of clip records."""
    skip_prompts = args.yes
    mixed_selectors = getattr(args, "mixed", None)
    output_format = "screen" if args.screen else "gif" if args.gif else "clip"

    selection_mode_set = bool(
        args.batch
        or args.lines
        or args.range
        or args.category
        or args.cell
        or args.participant
        or args.filter
        or mixed_selectors
        or args.reel
        or args.timeline
    )

    def _parse_cli_categories(raw: Optional[str]) -> List[str]:
        """Parse CLI category string into a list of category names."""
        if not raw:
            return []
        combined = raw.replace(",", "+")
        seen = set()
        result: List[str] = []
        for token in combined.split("+"):
            name = token.strip()
            if not name:
                continue
            if name not in seen:
                seen.add(name)
                result.append(name)
        return result

    cli_categories = _parse_cli_categories(getattr(args, "category", None))

    mode_dispatch: List[tuple] = [
        (
            args.batch or (output_format != "clip" and not selection_mode_set),
            "batch",
            {},
        ),
        (args.lines, "line", {"line_numbers": cli_mode_args.line_numbers}),
        (
            args.range,
            "range",
            {
                "range_start": cli_mode_args.range_start,
                "range_end": cli_mode_args.range_end,
            },
        ),
        (args.category, "category", {"categories": cli_categories}),
        (args.cell, "cell", {"cell_specs": cli_mode_args.cell_specs}),
        (args.participant, "participant", {"participant_id": args.participant}),
        (args.filter, "filter", {}),
        (mixed_selectors, "reel", {"reel_input": mixed_selectors}),
        (args.reel, "reel", {"reel_input": args.reel}),
        (args.timeline, "reel", {"reel_input": f"timeline, {args.timeline}"}),
    ]

    for condition, mode, kwargs in mode_dispatch:
        if condition:
            return spreadsheet.generate_list(
                worksheet, mode, skip_prompts=skip_prompts, **kwargs
            )
    return []


def _resolve_timeline_output_file(
    args: Any, clips_list: List[ClipRecord]
) -> Optional[str]:
    """Build the output filename for timeline reel mode."""
    if not args.timeline:
        return None
    participant_id = utils.normalize_participant_id(args.timeline).strip()
    study_name = clips_list[0].get("study", "").strip() if clips_list else ""
    if study_name and participant_id:
        return files.get_unique_filename(
            f"{study_name}_{participant_id}_timeline{config.FILEFORMAT}"
        )
    if participant_id:
        return files.get_unique_filename(
            f"{participant_id}_timeline{config.FILEFORMAT}"
        )
    return files.get_unique_filename(f"timeline{config.FILEFORMAT}")


# ---- CLI mode runner ----


def _run_timeline_viewer_mode(worksheet: Any, args: Any) -> None:
    """Export all clips via batch mode and generate a per-participant timeline viewer."""
    clips_list = spreadsheet.generate_list(worksheet, "batch", skip_prompts=True)
    outputs_generated, artifacts = clipgen.process_clips(
        clips_list, output_format="clip"
    )

    if not config.REENCODING:
        clipgen._print_reencoding_warning(utils.verbose_print)
    clipgen._print_completion_message(outputs_generated, "clip", is_reel=False)

    if not artifacts:
        utils.warning_print("No artifacts were generated; skipping timeline viewer.")
        return

    study = artifacts[0].get("study", "")
    data = viewer.finalize_timeline_data(
        artifacts,
        study=study,
        worksheet_title=getattr(worksheet, "title", ""),
        is_excel=clipgen._is_excel_worksheet(worksheet),
        mode="timeline-viewer",
        output_format="clip",
    )
    viewer_path = viewer.generate_timeline_viewer(
        data,
        template_name="timeline-viewer.html",
        output_basename="timeline_viewer.html",
    )
    if viewer_path:
        utils.info_print(f"Participant timeline viewer created: {viewer_path}")

    if getattr(args, "manifest", False):
        manifest_path = viewer.save_manifest(
            artifacts,
            study=study,
            worksheet_title=getattr(worksheet, "title", ""),
            is_excel=clipgen._is_excel_worksheet(worksheet),
            mode="timeline-viewer",
            output_format="clip",
        )
        if manifest_path:
            utils.info_print(f"Manifest updated: {manifest_path}")


def run_cli_mode(worksheet: Any, args: Any, cli_mode_args: CliModeArgs) -> None:
    """Execute CLI mode - run once and exit.

    Args:
        worksheet: Selected worksheet
        args: Parsed command-line arguments
        cli_mode_args: Parsed line/range/cell arguments
    """
    output_format = "screen" if args.screen else "gif" if args.gif else "clip"
    mixed_selectors = getattr(args, "mixed", None)

    if (args.reel or args.timeline) and output_format != "clip":
        utils.error_print(
            "Reel/timeline mode cannot be combined with --screen or --gif.",
            [
                "Use reel/timeline mode for a single .mp4 output, or use screen/gif with batch/line/range/category/cell/participant/filter selection."
            ],
        )
        sys.exit(1)

    if mixed_selectors:
        parsed_mixed = spreadsheet.parse_reel_input(mixed_selectors)
        if parsed_mixed.get("timeline"):
            utils.error_print(
                "Timeline selector is not supported in mixed mode.",
                [
                    "Use -T PARTICIPANT for a chronological reel,",
                    "or use -R with timeline selectors to create a single reel video.",
                ],
            )
            sys.exit(1)

    clips_list = _generate_cli_clips(worksheet, args, cli_mode_args)

    if args.reel or args.timeline:
        reel_output_file = _resolve_timeline_output_file(args, clips_list)
        outputs_generated, artifacts = clipgen.process_reel(
            clips_list,
            output_file=reel_output_file,
        )
    else:
        outputs_generated, artifacts = clipgen.process_clips(
            clips_list,
            output_format=output_format,
        )

    if not config.REENCODING:
        clipgen._print_reencoding_warning(utils.verbose_print)
    clipgen._print_completion_message(
        outputs_generated, output_format, is_reel=bool(args.reel or args.timeline)
    )

    if (
        getattr(args, "viewer", False) or getattr(args, "manifest", False)
    ) and artifacts:
        study = artifacts[0].get("study", "")
        participant = artifacts[0].get("participant", "")
        ws_title = getattr(worksheet, "title", "")
        is_excel = clipgen._is_excel_worksheet(worksheet)
        effective_mode = output_format if output_format != "clip" else "batch"

        if getattr(args, "viewer", False):
            data = viewer.finalize_timeline_data(
                artifacts,
                study=study,
                participant=participant,
                worksheet_title=ws_title,
                is_excel=is_excel,
                mode=effective_mode,
                output_format=output_format,
            )
            viewer_path = viewer.generate_timeline_viewer(data)
            if viewer_path:
                utils.info_print(f"Timeline viewer created: {viewer_path}")

        if getattr(args, "manifest", False):
            manifest_path = viewer.save_manifest(
                artifacts,
                study=study,
                participant=participant,
                worksheet_title=ws_title,
                is_excel=is_excel,
                mode=effective_mode,
                output_format=output_format,
            )
            if manifest_path:
                utils.info_print(f"Manifest updated: {manifest_path}")


# ---- Main entry point ----


def main() -> None:
    """Main entry point for clipgen."""
    setup_encoding()

    # Parse command-line arguments
    args = parse_arguments()
    if config.DEBUGGING:
        ic(args)

    # Determine if running in CLI mode (any mode argument provided)
    mixed_selectors = getattr(args, "mixed", None)
    timeline_viewer = getattr(args, "timeline_viewer", False)

    if timeline_viewer:
        conflicting = [
            args.batch,
            args.lines,
            args.range,
            args.category,
            args.cell,
            args.participant,
            args.filter,
            mixed_selectors,
            args.reel,
            args.timeline,
            args.screen,
            args.gif,
            args.viewer,
        ]
        if any(conflicting):
            utils.error_print(
                "--timeline-viewer cannot be combined with mode, format, or --viewer flags.",
                [
                    "Only -s (spreadsheet) and -v (verbose) may be used alongside --timeline-viewer."
                ],
            )
            sys.exit(1)

    cli_mode = (
        args.batch
        or args.lines
        or args.range
        or args.category
        or args.cell
        or args.participant
        or args.filter
        or mixed_selectors
        or args.reel
        or args.timeline
        or args.screen
        or args.gif
        or timeline_viewer
    )

    # Set verbosity: quiet by default in CLI mode, standard in interactive mode
    if cli_mode:
        config.VERBOSITY = config.VERBOSE if args.verbose else config.QUIET
    else:
        config.VERBOSITY = config.VERBOSE if args.verbose else config.STANDARD

    # Optional per-run override for titlecards setting
    if getattr(args, "titlecards", None) is not None:
        config.TITLECARDS_ENABLED = bool(args.titlecards)

    # Optional per-run overrides for input/output directories
    if getattr(args, "input", None) is not None:
        config.INPUT_DIR = args.input
    if getattr(args, "output", None) is not None:
        config.OUTPUT_DIR = args.output

    if getattr(args, "manifest", False):
        config.MANIFEST_ENABLED = True

    if getattr(args, "transcribe", False):
        config.TRANSCRIBE_ENABLED = True
    if getattr(args, "transcript_format", None):
        config.TRANSCRIBE_FORMAT = args.transcript_format

    # Parse CLI arguments for line, range, and cell modes
    cli_mode_args = parse_cli_mode_args(args)

    # Change working directory to runtime location (script/executable)
    os.chdir(get_runtime_working_dir())
    utils.standard_print(
        "-------------------------------------------------------------------------------"
    )
    utils.standard_print(
        f"Welcome to clipgen v{config.VERSIONNUM}\nWorking directory: {os.getcwd()}\nPlace video files and the credentials.json file in this directory."
    )
    utils.debug_print(
        "Debug mode is ON. Several limitations apply and more things will be printed."
    )

    # Sanity-check input/output directories before proceeding
    utils.validate_runtime_directories()

    if not video.check_ffmpeg_tools_available():
        sys.exit(1)

    # Standalone manifest → viewer: regenerate viewer from saved manifest, no spreadsheet needed
    if config.MANIFEST_ENABLED and getattr(args, "viewer", False) and not cli_mode:
        existing_artifacts = viewer.load_manifest_artifacts()
        if not existing_artifacts:
            utils.error_print(
                "No manifest found or manifest is empty.",
                [
                    f"Run a clip generation mode with --manifest first to create {config.MANIFEST_FILENAME}."
                ],
            )
            sys.exit(1)
        study = existing_artifacts[0].get("study", "")
        participant = existing_artifacts[0].get("participant", "")
        data = viewer.finalize_timeline_data(
            existing_artifacts, study=study, participant=participant, mode="manifest"
        )
        viewer_path = viewer.generate_timeline_viewer(data)
        if viewer_path:
            utils.info_print(f"Timeline viewer created from manifest: {viewer_path}")
        sys.exit(0)

    # Authenticate with Google (once per run)
    gspread_client = authenticate_google()
    doc_list = google_api.get_all_spreadsheets(gspread_client).split(",")

    # Outer loop so 'top' can return to spreadsheet selection
    while True:
        try:
            worksheet = select_worksheet(gspread_client, doc_list, args, cli_mode)

            # Execute based on mode
            if timeline_viewer:
                _run_timeline_viewer_mode(worksheet, args)
            elif cli_mode:
                run_cli_mode(worksheet, args, cli_mode_args)
            else:
                clipgen.run_interactive_mode(worksheet)
            break
        except utils.TopToSpreadsheet:
            # User requested to go back to spreadsheet selection; restart loop.
            continue
        except utils.QuitProgram:
            # Keyword-aware input requested exit; helper already printed context message.
            sys.exit(0)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        utils.info_print("Interrupted by user")
        sys.exit(0)
