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
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, NamedTuple

from icecream import ic

import clipgen
import config
import files
import spreadsheet
import transcripts
import utils
import video
import viewer
from utils import ClipRecord


# ---- CLI data structures ----


class CliModeArgs(NamedTuple):
    line_numbers: list[int] | None
    range_start: int | None
    range_end: int | None
    cell_specs: list[tuple[str, int]] | None


# ---- Argument parsing ----


def parse_arguments() -> argparse.Namespace:
    """Parse command-line arguments for non-interactive mode.

    At most one selection-mode flag (-b, -l, -r, -C, -c, -p, -k, -S, -M, -R, -T) may be
    given; if none is given, the program runs in interactive mode (see --help groups for
    all options: output format, transcription, paths, viewer/manifest, run flags).

    Returns:
        argparse.Namespace with mode flags/values, spreadsheet, yes, verbose, screen, gif,
        transcribe, transcript_format, pre_transcribe, viewer, manifest, timeline_viewer,
        input, output, titlecards, and related attributes.
    """
    parser = argparse.ArgumentParser(
        description=(
            "clipgen - Video clip generator from Google Sheets or local Excel timestamps. "
            "Interactive-only flows (e.g. browse, reellate) have no separate CLI flags; "
            "run without a selection-mode flag to use them."
        ),
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
  python clipgen.py -k                     Keyword mode - only key-marked clips/timestamps
  python clipgen.py -S "Critical,High"     Severity mode - Critical and High severity clips
  python clipgen.py -S "-4,-3"             Severity mode - using numeric values
  python clipgen.py -M "5, P01.11, 13-16"  Mixed mode - combine selectors for individual outputs
  python clipgen.py -b -s "Study Name"     Batch mode with specific spreadsheet
  python clipgen.py -s excel               Use the only .xlsx in cwd (fails if 0 or many)
  python clipgen.py -s ./notes.xlsx        Batch with local Excel workbook
  python clipgen.py -l 5 -y                Line mode, skip confirmation prompts
  python clipgen.py -b -v                  Batch mode with verbose output
  python clipgen.py -R "11, 13-16, P01, \\"Observations\\""  Reel mode - one combined video
  python clipgen.py -T P01                 Chronologic mode - chronological reel for participant P01
  python clipgen.py -b --screen            Batch mode screenshots (.png)
  python clipgen.py -l 5 --gif             Line mode GIF output (.gif)
  python clipgen.py --timeline-viewer      Generate per-participant timeline viewer
  python clipgen.py -b --transcribe        Batch with Markdown transcripts per clip
  python clipgen.py -b --transcribe --transcript-format vtt
  python clipgen.py -b --viewer --manifest Timeline viewer + manifest after batch run
  python clipgen.py --viewer               Regenerate clips_viewer.html from saved manifest
  python clipgen.py --regenerate            Regenerate all media artifacts from saved manifest
  python clipgen.py --studio                  Launch Studio web interface
  python clipgen.py --studio -s "My Study"   Studio with specific spreadsheet
  python clipgen.py -b -i ./videos -o ./out   Custom input/output directories
  python clipgen.py -b --titlecards        Enable titlecards for this run
  python clipgen.py -b --no-titlecards     Disable titlecards for this run

Note: Non-interactive mode (using -b, -l, -r, -C, -c, -p, -k, -S, -M, -R, or -T) is silent by default,
      only showing errors and the final summary. Use -v for full output.
""",
    )

    selection = parser.add_argument_group("selection mode (choose at most one)")
    mode_group = selection.add_mutually_exclusive_group()
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
        "-k",
        "--keyword",
        nargs="?",
        const=True,
        default=False,
        metavar="ANNOTATIONS",
        help='Keyword mode: generate only annotated clips. Optionally specify annotation types (e.g., "key,bug")',
    )
    mode_group.add_argument(
        "-S",
        "--severity",
        type=str,
        metavar="SEVERITIES",
        help='Severity mode: filter by severity levels (e.g., "Critical,High" or "-4,-3")',
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
        "--chronologic",
        type=str,
        metavar="PARTICIPANT",
        help="Chronologic mode: chronological reel for one participant (e.g., P01)",
    )
    mode_group.add_argument(
        "-H",
        "--highlights",
        nargs="?",
        const="highlights",
        type=str,
        metavar="DURATION",
        help="Highlights reel: auto-select best clips within time budget (default 180s). Optionally specify duration in seconds.",
    )

    output_fmt = parser.add_argument_group("output format (choose at most one)")
    format_group = output_fmt.add_mutually_exclusive_group()
    format_group.add_argument(
        "--screen",
        action="store_true",
        help="Output screenshots (.png) instead of video clips",
    )
    format_group.add_argument(
        "--gif", action="store_true", help="Output animated GIFs instead of video clips"
    )

    transcription = parser.add_argument_group("transcription")
    transcription.add_argument(
        "--transcribe",
        action="store_true",
        help="Generate transcript files alongside artifacts",
    )
    transcription.add_argument(
        "--transcript-format",
        type=str,
        choices=["md", "srt", "vtt"],
        metavar="FMT",
        help="Transcript format: md (default), srt, or vtt",
    )
    transcription.add_argument(
        "--pre-transcribe",
        nargs="*",
        metavar="ID",
        default=None,
        help="Pre-transcribe source videos. No IDs = all participants. Specify IDs to transcribe specific participants (e.g., P01 P03).",
    )
    transcription.add_argument(
        "--whisper-model",
        type=str,
        choices=["tiny", "base", "small", "medium", "large-v3"],
        metavar="MODEL",
        help="Whisper model for transcription: tiny, base, small, medium, large-v3 (default: base)",
    )
    transcription.add_argument(
        "--summarize",
        nargs="*",
        metavar="ID",
        default=None,
        help="Run the summary thinking agent over already-transcribed participants. "
        "No IDs = all transcribed. Existing summaries are kept unless -y is passed.",
    )
    transcription.add_argument(
        "--citations",
        nargs="*",
        metavar="ID",
        default=None,
        help="Run the citation thinking agent over participants that already have a summary. "
        "No IDs = all eligible. Existing citations are kept unless -y is passed.",
    )

    ai_opts = parser.add_argument_group("AI models")
    ai_opts.add_argument(
        "--ollama-model",
        type=str,
        metavar="MODEL",
        help="Ollama model for transcript summaries and citations (e.g. gemma3:4b)",
    )

    paths = parser.add_argument_group("spreadsheet & directories")
    paths.add_argument(
        "-s",
        "--spreadsheet",
        type=str,
        metavar="SOURCE",
        help=(
            "Google Sheet title, numeric index from your account list, full spreadsheet URL, "
            f"path to a local .xlsx file, or keyword {config.COMMAND_EXCEL!r} to use the "
            "only .xlsx in the current directory (errors if there are zero or multiple files)"
        ),
    )
    paths.add_argument(
        "-i",
        "--input",
        type=str,
        metavar="DIR",
        help="Input directory where source videos are located (defaults to current working directory when unset)",
    )
    paths.add_argument(
        "-o",
        "--output",
        type=str,
        metavar="DIR",
        help="Output directory where generated artifacts will be written (defaults to current working directory when unset)",
    )

    viewer_manifest = parser.add_argument_group("viewer & manifest")
    viewer_manifest.add_argument(
        "--viewer",
        action="store_true",
        help="Generate a timeline HTML viewer file (clips_viewer.html). With a mode flag, creates viewer from that run's artifacts. Alone, regenerates from saved manifest.",
    )
    viewer_manifest.add_argument(
        "--manifest",
        action="store_true",
        help="Write artifact metadata to a cumulative manifest JSON file alongside generated clips",
    )
    viewer_manifest.add_argument(
        "--export",
        action="store_true",
        help="Export analysis-ready JSON+CSV from manifests in the output directory (Screenspace events, Insights, Transcripts). Skips manifests that aren't present.",
    )
    viewer_manifest.add_argument(
        "--timeline-viewer",
        action="store_true",
        help="Batch-export all clips and generate a per-participant timeline HTML viewer",
    )
    viewer_manifest.add_argument(
        "--regenerate",
        action="store_true",
        help="Regenerate all media artifacts from saved manifest (no spreadsheet needed)",
    )
    viewer_manifest.add_argument(
        "--studio",
        action="store_true",
        help="Launch the Studio web interface for interactive artifact generation and reel building",
    )
    viewer_manifest.add_argument(
        "--insights",
        action="store_true",
        help="Launch Insights for authoring research findings from generated artifacts",
    )
    viewer_manifest.add_argument(
        "--screenspace",
        action="store_true",
        help="Launch the Screenspace analysis interface for video frame analysis",
    )
    viewer_manifest.add_argument(
        "--transcripts",
        action="store_true",
        help="Launch the Transcript workspace for viewing, editing, and managing transcriptions",
    )
    viewer_manifest.add_argument(
        "--gallery",
        type=str,
        nargs="?",
        const="",
        metavar="VIDEO",
        help="Generate a gallery viewer with interval screenshots/GIFs from a video file",
    )
    viewer_manifest.add_argument(
        "--interval",
        type=int,
        metavar="SECONDS",
        help=f"Capture interval in seconds for gallery mode (default: {config.GALLERY_INTERVAL_SECONDS})",
    )
    viewer_manifest.add_argument(
        "--bundle",
        action="store_true",
        help="Embed gallery images as base64 data URIs in the HTML (makes it fully self-contained)",
    )

    screenspace_cli = parser.add_argument_group("screenspace cli")
    ss_modes = screenspace_cli.add_mutually_exclusive_group()
    ss_modes.add_argument(
        "--ss-task",
        nargs=3,
        metavar=("TYPE", "PARTICIPANT", "REGION"),
        default=None,
        help=(
            "Run a Screenspace analysis task headlessly. "
            "TYPE is one of color, change, similarity, text, numbers, timelapse, "
            "template, flow, inactivity. REGION must already exist in the active "
            "manifest or in a stash (use --ss-list-regions / --ss-list-stashes)."
        ),
    )
    ss_modes.add_argument(
        "--ss-list-regions",
        action="store_true",
        help="List active Screenspace regions from the manifest and exit.",
    )
    ss_modes.add_argument(
        "--ss-list-stashes",
        action="store_true",
        help="List Screenspace region stashes and exit.",
    )
    ss_modes.add_argument(
        "--ss-list-tasks",
        nargs="?",
        const="",
        default=None,
        metavar="STATUS",
        help=(
            "List Screenspace tasks from the manifest. Optional STATUS filter: "
            "queued, running, completed, failed, cancelled, paused."
        ),
    )

    screenspace_cli.add_argument(
        "--ss-target-color",
        type=str,
        metavar="HEX",
        help="Target colour as #RRGGBB hex (color tool).",
    )
    screenspace_cli.add_argument(
        "--ss-tolerance",
        type=str,
        metavar="H,S,V",
        help="HSV tolerance triple as comma-separated ints (color tool).",
    )
    screenspace_cli.add_argument(
        "--ss-threshold",
        type=float,
        metavar="FLOAT",
        help="Match threshold (color, change, similarity, template, flow, inactivity).",
    )
    screenspace_cli.add_argument(
        "--ss-reference-timestamp",
        type=float,
        metavar="SECONDS",
        help="Reference frame timestamp (similarity, template).",
    )
    screenspace_cli.add_argument(
        "--ss-text",
        type=str,
        metavar="STR",
        help="Search string (text tool).",
    )
    screenspace_cli.add_argument(
        "--ss-fuzzy-threshold",
        type=float,
        metavar="FLOAT",
        help=f"OCR fuzzy-match threshold for the text tool (default: {config.SCREENSPACE_OCR_FUZZY_THRESHOLD}).",
    )
    screenspace_cli.add_argument(
        "--ss-operator",
        type=str,
        choices=["eq", "gt", "lt", "gte", "lte", "range"],
        metavar="OP",
        help="Numeric comparison operator (numbers tool).",
    )
    screenspace_cli.add_argument(
        "--ss-target-value",
        type=float,
        metavar="FLOAT",
        help="Target numeric value (numbers tool, non-range operators).",
    )
    screenspace_cli.add_argument(
        "--ss-range-min",
        type=float,
        metavar="FLOAT",
        help="Range minimum (numbers tool, range operator).",
    )
    screenspace_cli.add_argument(
        "--ss-range-max",
        type=float,
        metavar="FLOAT",
        help="Range maximum (numbers tool, range operator).",
    )
    screenspace_cli.add_argument(
        "--ss-speedup",
        type=float,
        metavar="FACTOR",
        help="Speed-up factor (timelapse tool).",
    )
    screenspace_cli.add_argument(
        "--ss-output-format",
        type=str,
        choices=["mp4", "gif"],
        metavar="FMT",
        help="Output format (timelapse tool).",
    )
    screenspace_cli.add_argument(
        "--ss-start",
        type=float,
        metavar="SECONDS",
        help="Start time in seconds (timelapse and other range-aware tools).",
    )
    screenspace_cli.add_argument(
        "--ss-end",
        type=float,
        metavar="SECONDS",
        help="End time in seconds (timelapse and other range-aware tools).",
    )
    screenspace_cli.add_argument(
        "--ss-interval",
        type=float,
        metavar="SECONDS",
        help=f"Frame sampling interval (default: {config.SCREENSPACE_DEFAULT_INTERVAL}).",
    )
    screenspace_cli.add_argument(
        "--ss-event-label",
        type=str,
        metavar="STR",
        help="Override the auto-generated event label written to the manifest.",
    )

    run_opts = parser.add_argument_group("run options")
    run_opts.add_argument(
        "-y",
        "--yes",
        action="store_true",
        help="Skip confirmation prompts (auto-confirm)",
    )
    run_opts.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Increase verbosity (-v = verbose output; default is quiet in CLI, standard in interactive mode)",
    )

    titlecards_grp = parser.add_argument_group("title cards (choose at most one)")
    titlecard_group = titlecards_grp.add_mutually_exclusive_group()
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

    filmstrip_grp = parser.add_argument_group("filmstrip (choose at most one)")
    filmstrip_group = filmstrip_grp.add_mutually_exclusive_group()
    filmstrip_group.add_argument(
        "--filmstrip",
        dest="filmstrip",
        action="store_true",
        help="Enable filmstrip thumbnail mode on timeline markers in the HTML viewer",
    )
    filmstrip_group.add_argument(
        "--no-filmstrip",
        dest="filmstrip",
        action="store_false",
        help="Disable filmstrip thumbnail mode on timeline markers in the HTML viewer",
    )
    parser.set_defaults(filmstrip=None)

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


def authenticate_google() -> Any | None:
    """Authenticate with Google Sheets API.

    Returns:
        Google client connection object on success, or None if authentication
        failed. Callers are responsible for deciding whether to fall back to a
        local Excel file or exit.
    """
    import gspread

    try:
        utils.debug_print("Attempting login...")
        gspread_client = gspread.oauth(credentials_filename="credentials.json")
        utils.debug_print("Login successful!")
        return gspread_client
    except (gspread.exceptions.GSpreadException, FileNotFoundError, OSError) as e:
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
        return None


# ---- Worksheet selection ----


def _is_excel_spreadsheet_arg(spreadsheet_arg: str | None) -> bool:
    """Return True if the -s argument points to a local Excel file."""
    if not spreadsheet_arg:
        return False
    raw = spreadsheet_arg.strip().lower()
    return raw == config.COMMAND_EXCEL or raw.endswith(".xlsx")


def select_worksheet(
    gspread_client: Any, doc_list: list[str], args: Any, cli_mode: bool
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
    import excel_io

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
) -> list[ClipRecord]:
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
        or args.keyword
        or args.severity
        or mixed_selectors
        or args.reel
        or args.chronologic
        or args.highlights
    )

    def _parse_cli_categories(raw: str | None) -> list[str]:
        """Parse CLI category string into a list of category names."""
        if not raw:
            return []
        combined = raw.replace(",", "+")
        seen = set()
        result: list[str] = []
        for token in combined.split("+"):
            name = token.strip()
            if not name:
                continue
            if name not in seen:
                seen.add(name)
                result.append(name)
        return result

    cli_categories = _parse_cli_categories(getattr(args, "category", None))

    def _parse_cli_severities(raw: str | None) -> list[str]:
        if not raw:
            return []
        combined = raw.replace(",", "+")
        seen = set()
        result: list[str] = []
        for token in combined.split("+"):
            name = utils.normalize_severity(token.strip())
            if not name:
                continue
            if name not in seen:
                seen.add(name)
                result.append(name)
        return result

    cli_severities = _parse_cli_severities(getattr(args, "severity", None))

    cli_annotation_ids = None
    if isinstance(args.keyword, str):
        combined = args.keyword.replace(",", "+")
        cli_annotation_ids = [
            t.strip().lower().lstrip("!") for t in combined.split("+") if t.strip()
        ] or None

    # Apply custom highlights duration if specified (e.g. -H 120)
    if args.highlights and args.highlights != "highlights":
        try:
            config.HIGHLIGHTS_REEL_DURATION_SECONDS = int(args.highlights)
        except ValueError:
            utils.warning_print(
                f"Invalid highlights duration '{args.highlights}', using default ({config.HIGHLIGHTS_REEL_DURATION_SECONDS}s)."
            )

    mode_dispatch: list[tuple] = [
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
        (args.keyword, "keyword", {"annotation_ids": cli_annotation_ids}),
        (args.severity, "severity", {"severities": cli_severities}),
        (mixed_selectors, "reel", {"reel_input": mixed_selectors}),
        (args.reel, "reel", {"reel_input": args.reel}),
        (args.chronologic, "reel", {"reel_input": f"chronologic, {args.chronologic}"}),
        (args.highlights, "reel", {"reel_input": "highlights, batch"}),
    ]

    for condition, mode, kwargs in mode_dispatch:
        if condition:
            return spreadsheet.generate_list(
                worksheet, mode, skip_prompts=skip_prompts, **kwargs
            )
    return []


def _resolve_chronologic_output_file(
    args: Any, clips_list: list[ClipRecord]
) -> str | None:
    """Build the output filename for chronologic reel mode."""
    if not args.chronologic:
        return None
    participant_id = utils.normalize_participant_id(args.chronologic).strip()
    study_name = clips_list[0].get("study", "").strip() if clips_list else ""
    if study_name and participant_id:
        return files.get_unique_filename(
            f"{study_name}_{participant_id}_chronologic{config.FILEFORMAT}"
        )
    if participant_id:
        return files.get_unique_filename(
            f"{participant_id}_chronologic{config.FILEFORMAT}"
        )
    return files.get_unique_filename(f"chronologic{config.FILEFORMAT}")


def _resolve_highlights_output_file(
    clips_list: list[ClipRecord],
) -> str | None:
    """Build the output filename for highlights reel mode."""
    study_name = clips_list[0].get("study", "").strip() if clips_list else ""
    if study_name:
        return files.get_unique_filename(f"{study_name}_highlights{config.FILEFORMAT}")
    return files.get_unique_filename(f"highlights{config.FILEFORMAT}")


# ---- CLI mode runner ----


def _run_gallery_cli(args: argparse.Namespace) -> None:
    """Generate interval captures from a video and build a gallery viewer (no spreadsheet needed)."""
    gallery_arg = getattr(args, "gallery", "")
    input_dir = utils.get_effective_input_dir()

    if gallery_arg:
        video_path = Path(gallery_arg)
        if not video_path.is_absolute():
            video_path = input_dir / video_path
    else:
        videos = sorted(
            p for p in input_dir.glob(f"*{config.FILEFORMAT}") if p.is_file()
        )
        if not videos:
            utils.error_print(
                f"No {config.FILEFORMAT} files found in {input_dir}.",
                [
                    "Specify a video file: --gallery path/to/video.mp4",
                    "Or place a video file in the input directory.",
                ],
            )
            return
        video_path = videos[0]
        utils.info_print(f"Using video: {video_path.name}")

    if not video_path.is_file():
        utils.error_print(f"Video file not found: '{video_path}'")
        return

    output_format = "gif" if getattr(args, "gif", False) else "screen"
    interval = getattr(args, "interval", None) or config.GALLERY_INTERVAL_SECONDS
    bundle = getattr(args, "bundle", False) or config.GALLERY_BUNDLE_ENABLED

    artifacts = video.generate_interval_captures(
        str(video_path),
        interval_seconds=interval,
        output_format=output_format,
        gif_duration_seconds=config.GALLERY_GIF_DURATION_SECONDS,
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
        bundle=bundle,
    )
    gallery_path = viewer.generate_gallery_viewer(data)
    if gallery_path:
        utils.info_print(f"Gallery viewer created: {gallery_path}")


def _run_pre_transcribe(worksheet: Any, args: Any) -> None:
    """Pre-transcribe source videos for specified participants."""
    ctx = spreadsheet.build_sheet_context(worksheet)
    if ctx is None:
        utils.error_print("Cannot build sheet context.")
        sys.exit(1)

    available = spreadsheet.get_participant_list(
        ctx.header_row, ctx.id_cell, ctx.num_participants
    )

    requested_ids = getattr(args, "pre_transcribe", [])
    if not requested_ids:
        target_ids = list(available)
    else:
        target_ids = []
        for raw_id in requested_ids:
            pid = utils.normalize_participant_id(raw_id).strip()
            if pid in available:
                target_ids.append(pid)
            else:
                utils.warning_print(
                    f"Participant '{raw_id}' not found in spreadsheet. "
                    f"Available: {', '.join(available)}"
                )

    if not target_ids:
        utils.error_print("No valid participants to transcribe.")
        return

    manifest = transcripts.load_transcripts_manifest()
    source_transcripts = manifest["source_transcripts"]
    corrections = manifest["corrections"]
    context_keywords = transcripts.get_corrections_keywords(corrections) or None

    skipped = 0
    transcribed = 0

    for pid in target_ids:
        if pid in source_transcripts:
            utils.info_print(f"Skipping {pid}: already transcribed.")
            skipped += 1
            continue

        col_idx = spreadsheet.find_participant_column(ctx.header_row, ctx.id_cell, pid)
        override = None
        if (
            col_idx is not None
            and ctx.filename_row_idx is not None
            and ctx.filename_row_idx < len(ctx.sheet_data)
            and col_idx < len(ctx.sheet_data[ctx.filename_row_idx])
        ):
            override = ctx.sheet_data[ctx.filename_row_idx][col_idx].strip() or None

        video_filename = files.get_source_video_filename(ctx.study_name, pid, override)
        video_path = utils.resolve_input_path(video_filename)
        if not video_path.is_file():
            utils.error_print(f"Source video not found for {pid}: {video_path}")
            continue

        utils.info_print(f"Transcribing {pid}: {video_path.name}...")
        result = transcripts.transcribe_video(
            str(video_path),
            context_keywords=context_keywords,
        )
        if result is None:
            utils.error_print(f"Transcription failed for {pid}.")
            continue

        source_transcripts[pid] = {
            "segments": [
                {"start": s["start"], "end": s["end"], "text": s["text"]}
                for s in result["segments"]
            ],
            "language": result["language"],
            "model": result["model"],
            "source_file": result["source_file"],
            "transcribed_at": datetime.now(timezone.utc).isoformat(),
        }
        transcribed += 1

        transcripts.save_transcripts_manifest(source_transcripts, corrections)
        utils.info_print(f"  {pid}: {len(result['segments'])} segments stored.")

    utils.info_print(
        f"Pre-transcription complete: {transcribed} transcribed, {skipped} skipped."
    )


# ---- Screenspace CLI ----


_SS_VALID_TASK_TYPES = (
    "color",
    "change",
    "similarity",
    "text",
    "numbers",
    "timelapse",
    "template",
    "flow",
    "inactivity",
)


def _ss_load_known_regions(manifest: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Combine active manifest regions with all stash regions (last-write-wins)."""
    known: dict[str, dict[str, Any]] = dict(manifest.get("regions", {}))
    for stash in manifest.get("stashes", []):
        known.update(stash.get("regions", {}))
    return known


def _ss_resolve_video_for_participant(participant_id: str) -> str | None:
    """Resolve the source video path for a participant via filename discovery.

    Mirrors how screenspace_server falls back when no spreadsheet is loaded.
    """
    discovered = utils.discover_participant_videos("")
    for entry in discovered:
        if entry["id"] == participant_id and entry.get("has_video"):
            return entry["video_path"]
    return None


def _ss_hex_to_hsv(hex_str: str) -> dict[str, int]:
    """Convert a #RRGGBB hex string to OpenCV HSV (H 0–180, S 0–255, V 0–255)."""
    import cv2
    import numpy as np

    s = hex_str.strip().lstrip("#")
    if len(s) != 6:
        raise ValueError(f"Invalid hex colour {hex_str!r} (expected #RRGGBB)")
    try:
        r, g, b = int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)
    except ValueError as exc:
        raise ValueError(f"Invalid hex colour {hex_str!r}") from exc
    bgr = np.array([[[b, g, r]]], dtype=np.uint8)
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)[0][0]
    return {"h": int(hsv[0]), "s": int(hsv[1]), "v": int(hsv[2])}


def _ss_parse_tolerance(tol_str: str) -> dict[str, int]:
    """Parse a comma-separated H,S,V tolerance triple."""
    parts = [p.strip() for p in tol_str.split(",")]
    if len(parts) != 3:
        raise ValueError(
            f"Tolerance must be three comma-separated ints (got {tol_str!r})"
        )
    try:
        h, s, v = (int(parts[0]), int(parts[1]), int(parts[2]))
    except ValueError as exc:
        raise ValueError(
            f"Tolerance must be three comma-separated ints (got {tol_str!r})"
        ) from exc
    return {"h": h, "s": s, "v": v}


def _ss_build_params(
    args: argparse.Namespace,
    task_type: str,
    region_coords: dict[str, int],
    video_path: str,
) -> dict[str, Any]:
    """Build a `parameters` dict for create_task() from per-tool CLI flags.

    Validates that the required flags for ``task_type`` are present. For
    ``similarity`` and ``template`` extracts the reference frame from the
    video at ``--ss-reference-timestamp`` (mirrors the server-side path in
    screenspace_server._extract_task_media).
    """
    import screenspace

    params: dict[str, Any] = {}

    if args.ss_interval is not None:
        params["interval"] = args.ss_interval
    if args.ss_start is not None:
        params["start_seconds"] = args.ss_start
    if args.ss_end is not None:
        params["end_seconds"] = args.ss_end
    if args.ss_event_label:
        params["event_label"] = args.ss_event_label

    if task_type == "color":
        if not args.ss_target_color:
            raise ValueError("color task requires --ss-target-color HEX")
        if not args.ss_tolerance:
            raise ValueError("color task requires --ss-tolerance H,S,V")
        params["target_color"] = _ss_hex_to_hsv(args.ss_target_color)
        params["tolerance"] = _ss_parse_tolerance(args.ss_tolerance)

    elif task_type == "change":
        if args.ss_threshold is None:
            raise ValueError("change task requires --ss-threshold FLOAT")
        params["threshold"] = args.ss_threshold

    elif task_type == "similarity":
        if args.ss_reference_timestamp is None:
            raise ValueError(
                "similarity task requires --ss-reference-timestamp SECONDS"
            )
        if args.ss_threshold is None:
            raise ValueError("similarity task requires --ss-threshold FLOAT")
        params["reference_timestamp"] = args.ss_reference_timestamp
        params["threshold"] = args.ss_threshold
        frame = video.extract_frame_at_timestamp(
            video_path, float(args.ss_reference_timestamp)
        )
        if frame is None:
            raise ValueError(
                f"Could not extract reference frame at {args.ss_reference_timestamp}s"
            )
        params["reference_frame"] = screenspace.extract_region(frame, region_coords)

    elif task_type == "text":
        if not args.ss_text:
            raise ValueError("text task requires --ss-text STR")
        params["search_string"] = args.ss_text
        if args.ss_fuzzy_threshold is not None:
            params["fuzzy_threshold"] = args.ss_fuzzy_threshold

    elif task_type == "numbers":
        if not args.ss_operator:
            raise ValueError("numbers task requires --ss-operator OP")
        params["operator"] = args.ss_operator
        if args.ss_operator == "range":
            if args.ss_range_min is None or args.ss_range_max is None:
                raise ValueError(
                    "numbers task with operator=range requires "
                    "--ss-range-min and --ss-range-max"
                )
            params["range_min"] = args.ss_range_min
            params["range_max"] = args.ss_range_max
        else:
            if args.ss_target_value is None:
                raise ValueError(
                    f"numbers task with operator={args.ss_operator} requires --ss-target-value"
                )
            params["target_value"] = args.ss_target_value

    elif task_type == "timelapse":
        if args.ss_speedup is None:
            raise ValueError("timelapse task requires --ss-speedup FACTOR")
        params["speedup_factor"] = args.ss_speedup
        if args.ss_output_format:
            params["output_format"] = args.ss_output_format

    elif task_type == "template":
        if args.ss_reference_timestamp is None:
            raise ValueError("template task requires --ss-reference-timestamp SECONDS")
        if args.ss_threshold is None:
            raise ValueError("template task requires --ss-threshold FLOAT")
        params["reference_timestamp"] = args.ss_reference_timestamp
        params["threshold"] = args.ss_threshold
        frame = video.extract_frame_at_timestamp(
            video_path, float(args.ss_reference_timestamp)
        )
        if frame is None:
            raise ValueError(
                f"Could not extract template frame at {args.ss_reference_timestamp}s"
            )
        params["template_image"] = screenspace.extract_region(frame, region_coords)

    elif task_type == "flow":
        if args.ss_threshold is None:
            raise ValueError("flow task requires --ss-threshold FLOAT (magnitude)")
        params["magnitude_threshold"] = args.ss_threshold

    elif task_type == "inactivity":
        if args.ss_threshold is None:
            raise ValueError("inactivity task requires --ss-threshold FLOAT")
        params["threshold"] = args.ss_threshold

    params.setdefault("cv_resolution_scale", config.SCREENSPACE_CV_RESOLUTION_SCALE)
    return params


def _run_ss_list_regions(args: argparse.Namespace) -> None:
    """List active Screenspace regions from the manifest."""
    import screenspace

    manifest = screenspace.load_screenspace_manifest()
    regions = manifest.get("regions", {})
    if not regions:
        utils.info_print("No active Screenspace regions in manifest.")
        return
    utils.info_print(f"Active regions ({len(regions)}):")
    for name in sorted(regions.keys()):
        rd = regions[name]
        sw = rd.get("source_width", "?")
        sh = rd.get("source_height", "?")
        utils.info_print(
            f"  {name}: x={rd.get('x', 0):.3f} y={rd.get('y', 0):.3f} "
            f"w={rd.get('w', 0):.3f} h={rd.get('h', 0):.3f}  source={sw}x{sh}"
        )


def _run_ss_list_stashes(args: argparse.Namespace) -> None:
    """List Screenspace region stashes."""
    import screenspace

    manifest = screenspace.load_screenspace_manifest()
    stashes = manifest.get("stashes", [])
    if not stashes:
        utils.info_print("No Screenspace region stashes in manifest.")
        return
    utils.info_print(f"Stashes ({len(stashes)}):")
    for stash in stashes:
        regions = stash.get("regions", {})
        names = ", ".join(sorted(regions.keys())) or "(empty)"
        utils.info_print(
            f"  {stash.get('id', '?')}  {stash.get('name', '(unnamed)')}: "
            f"{len(regions)} region(s) — {names}"
        )


def _run_ss_list_tasks(args: argparse.Namespace) -> None:
    """List Screenspace tasks from the manifest, optionally filtered by status."""
    import screenspace

    manifest = screenspace.load_screenspace_manifest()
    tasks = manifest.get("tasks", [])
    status_filter = (args.ss_list_tasks or "").strip().lower() or None
    if status_filter:
        tasks = [t for t in tasks if (t.get("status") or "").lower() == status_filter]
    if not tasks:
        if status_filter:
            utils.info_print(f"No Screenspace tasks with status={status_filter}.")
        else:
            utils.info_print("No Screenspace tasks in manifest.")
        return
    label = f" (status={status_filter})" if status_filter else ""
    utils.info_print(f"Tasks{label}: {len(tasks)}")
    for t in tasks:
        result = t.get("result")
        result_count = len(result) if isinstance(result, list) else 0
        utils.info_print(
            f"  {t.get('id', '?')}  {t.get('type', '?'):10s}  "
            f"{t.get('participant', '?'):8s}  region={t.get('region', '?'):16s}  "
            f"status={t.get('status', '?'):10s}  results={result_count}"
        )


def _run_ss_task(args: argparse.Namespace) -> None:
    """Run a Screenspace analysis task synchronously and persist the result."""
    import time

    import screenspace

    task_type, participant, region_name = args.ss_task
    if task_type not in _SS_VALID_TASK_TYPES:
        utils.error_print(
            f"Unknown screenspace task type {task_type!r}.",
            [f"Valid types: {', '.join(_SS_VALID_TASK_TYPES)}"],
        )
        sys.exit(1)

    manifest = screenspace.load_screenspace_manifest()
    known_regions = _ss_load_known_regions(manifest)
    if region_name not in known_regions:
        available = sorted(known_regions.keys())
        hint = (
            f"Available regions: {', '.join(available)}"
            if available
            else "No regions defined. Use --screenspace to define regions in the web UI first."
        )
        utils.error_print(f"Region {region_name!r} not found.", [hint])
        sys.exit(1)

    video_path = _ss_resolve_video_for_participant(participant)
    if video_path is None:
        utils.error_print(
            f"No video found for participant {participant!r}.",
            ["Place the source video in the input directory before running --ss-task."],
        )
        sys.exit(1)

    props = video.probe_video_properties(video_path)
    rd = known_regions[region_name]
    if props and props.get("width") and props.get("height"):
        region_coords = screenspace.denormalize_region(
            rd, props["width"], props["height"]
        )
    else:
        region_coords = {k: int(rd[k]) for k in ("x", "y", "w", "h") if k in rd}

    try:
        parameters = _ss_build_params(args, task_type, region_coords, video_path)
    except ValueError as exc:
        utils.error_print(str(exc))
        sys.exit(1)

    source_video = Path(video_path).name
    task = screenspace.create_task(
        task_type=task_type,
        participant=participant,
        source_video=source_video,
        video_path=video_path,
        region_name=region_name,
        region_coords=region_coords,
        parameters=parameters,
    )

    worker = screenspace.ScreenspaceWorker()
    worker.restore_tasks(manifest.get("tasks", []))
    worker.start()
    task_id = worker.enqueue(task)
    utils.info_print(f"Running {task_type} on {participant} (region: {region_name})...")

    last_progress = -1.0
    final_task: dict[str, Any] | None = None
    try:
        while True:
            current = worker.get_task(task_id)
            if current is None:
                break
            status = current.get("status", "")
            progress = float(current.get("progress", 0.0))
            if progress - last_progress > 0.05:
                utils.info_print(f"  {status}: {int(progress * 100)}%")
                last_progress = progress
            if status in ("completed", "failed", "cancelled"):
                final_task = current
                break
            time.sleep(0.25)
    finally:
        new_events = worker.drain_new_events()
        all_tasks = worker.get_all_tasks()
        events = list(manifest.get("events", [])) + new_events
        screenspace.save_screenspace_manifest(
            manifest.get("regions", {}),
            all_tasks,
            events,
            stashes=manifest.get("stashes", []),
        )
        worker.stop()

    if final_task is None:
        utils.error_print("Task did not complete (no final state).")
        sys.exit(1)

    status = final_task.get("status", "")
    if status == "completed":
        result = final_task.get("result")
        result_count = len(result) if isinstance(result, list) else 0
        if task_type == "timelapse":
            output = (
                result[0].get("output_path")
                if isinstance(result, list) and result
                else None
            )
            if output:
                utils.info_print(f"Timelapse written to {output}")
            else:
                utils.info_print("Timelapse complete.")
        else:
            utils.info_print(f"Completed: {result_count} result(s).")
    elif status == "failed":
        utils.error_print(f"Task failed: {final_task.get('error', 'unknown error')}")
        sys.exit(1)
    else:
        utils.info_print(f"Task ended with status={status}.")


# ---- Thinking-agent CLI ----


def _select_transcript_targets(
    requested: list[str] | None,
    source_transcripts: dict[str, Any],
) -> list[str]:
    """Resolve a requested participant list against transcripted participants.

    ``requested`` is the value of ``args.summarize`` or ``args.citations`` —
    None should never reach here, but ``[]`` means "all transcripted".
    Unknown IDs print a warning and are dropped.
    """
    if not requested:
        return list(source_transcripts.keys())
    targets: list[str] = []
    for raw_id in requested:
        pid = utils.normalize_participant_id(raw_id).strip()
        if pid in source_transcripts:
            targets.append(pid)
        else:
            utils.warning_print(
                f"Participant {raw_id!r} has no transcript; skipping. "
                f"Available: {', '.join(sorted(source_transcripts.keys()))}"
            )
    return targets


def _run_summarize(args: argparse.Namespace) -> None:
    """Run the summary thinking agent over already-transcribed participants."""
    import thinking_agents

    manifest = transcripts.load_transcripts_manifest()
    source_transcripts = manifest["source_transcripts"]
    corrections = manifest["corrections"]
    marks = manifest.get("marks")

    targets = _select_transcript_targets(args.summarize, source_transcripts)
    if not targets:
        utils.error_print("No transcribed participants to summarize.")
        return

    summarized = 0
    skipped = 0
    for pid in targets:
        entry = source_transcripts.get(pid)
        if not entry:
            utils.warning_print(f"{pid}: no transcript entry; skipping.")
            skipped += 1
            continue
        if entry.get("summary") and not args.yes:
            utils.info_print(f"{pid}: summary already present; skip (-y to overwrite).")
            skipped += 1
            continue
        utils.info_print(f"Summarizing {pid}...")
        summary = thinking_agents.summarize_transcript(entry.get("segments") or [])
        if not summary:
            utils.warning_print(
                f"{pid}: summary not produced (transcript too short or Ollama unavailable)."
            )
            skipped += 1
            continue
        entry["summary"] = summary
        transcripts.save_transcripts_manifest(source_transcripts, corrections, marks)
        utils.info_print(f"  {pid}: summary stored ({len(summary)} chars).")
        summarized += 1

    utils.info_print(f"Summary complete: {summarized} summarized, {skipped} skipped.")


def _run_citations(args: argparse.Namespace) -> None:
    """Run the citation thinking agent over participants with summaries."""
    import thinking_agents

    manifest = transcripts.load_transcripts_manifest()
    source_transcripts = manifest["source_transcripts"]
    corrections = manifest["corrections"]
    marks = manifest.get("marks")

    targets = _select_transcript_targets(args.citations, source_transcripts)
    if not targets:
        utils.error_print("No transcribed participants for citation generation.")
        return

    cited = 0
    skipped = 0
    for pid in targets:
        entry = source_transcripts.get(pid)
        if not entry:
            utils.warning_print(f"{pid}: no transcript entry; skipping.")
            skipped += 1
            continue
        if not entry.get("summary"):
            utils.warning_print(f"{pid}: no summary yet; run --summarize first.")
            skipped += 1
            continue
        if entry.get("citations") and not args.yes:
            utils.info_print(
                f"{pid}: citations already present; skip (-y to overwrite)."
            )
            skipped += 1
            continue
        utils.info_print(f"Finding citations for {pid}...")
        citations = thinking_agents.find_citations(
            entry["summary"], entry.get("segments") or []
        )
        if not citations:
            utils.warning_print(
                f"{pid}: no citations produced (Ollama unavailable or empty summary)."
            )
            skipped += 1
            continue
        entry["citations"] = citations
        transcripts.save_transcripts_manifest(source_transcripts, corrections, marks)
        total_refs = sum(len(c.get("refs") or []) for c in citations)
        utils.info_print(
            f"  {pid}: {len(citations)} claim(s), {total_refs} ref(s) stored."
        )
        cited += 1

    utils.info_print(f"Citations complete: {cited} processed, {skipped} skipped.")


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
    ss_events = viewer.load_screenspace_events_for_viewer()
    data = viewer.finalize_timeline_data(
        artifacts,
        study=study,
        worksheet_title=getattr(worksheet, "title", ""),
        is_excel=clipgen._is_excel_worksheet(worksheet),
        mode="timeline-viewer",
        output_format="clip",
        screenspace_events=ss_events or None,
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

    if (args.reel or args.chronologic or args.highlights) and output_format != "clip":
        utils.error_print(
            "Reel/chronologic/highlights mode cannot be combined with --screen or --gif.",
            [
                "Use reel/chronologic/highlights mode for a single .mp4 output, or use screen/gif with batch/line/range/category/cell/participant/keyword selection."
            ],
        )
        sys.exit(1)

    if mixed_selectors:
        parsed_mixed = spreadsheet.parse_reel_input(mixed_selectors)
        if parsed_mixed.get("chronologic"):
            utils.error_print(
                "Chronologic selector is not supported in mixed mode.",
                [
                    "Use -T PARTICIPANT for a chronological reel,",
                    "or use -R with chronologic selectors to create a single reel video.",
                ],
            )
            sys.exit(1)

    clips_list = _generate_cli_clips(worksheet, args, cli_mode_args)

    is_reel = bool(args.reel or args.chronologic or args.highlights)
    artifacts: list = []
    reel_records: list = []

    if is_reel:
        reel_output_file = _resolve_chronologic_output_file(args, clips_list)
        if args.highlights and reel_output_file is None:
            reel_output_file = _resolve_highlights_output_file(clips_list)
        outputs_generated, reel_records = clipgen.process_reel(
            clips_list,
            output_file=reel_output_file,
        )
    else:
        outputs_generated, artifacts = clipgen.process_clips(
            clips_list,
            output_format=output_format,
            include_severity=bool(args.severity),
        )

    if not config.REENCODING:
        clipgen._print_reencoding_warning(utils.verbose_print)
    clipgen._print_completion_message(
        outputs_generated,
        output_format,
        is_reel=is_reel,
    )

    ws_title = getattr(worksheet, "title", "")
    is_excel = clipgen._is_excel_worksheet(worksheet)
    effective_mode = output_format if output_format != "clip" else "batch"

    if (
        getattr(args, "viewer", False) or getattr(args, "manifest", False)
    ) and artifacts:
        study = artifacts[0].get("study", "")
        participant = artifacts[0].get("participant", "")

        if getattr(args, "viewer", False):
            ss_events = viewer.load_screenspace_events_for_viewer()
            data = viewer.finalize_timeline_data(
                artifacts,
                study=study,
                participant=participant,
                worksheet_title=ws_title,
                is_excel=is_excel,
                mode=effective_mode,
                output_format=output_format,
                screenspace_events=ss_events or None,
            )
            viewer_path = viewer.generate_timeline_viewer(data)
            if viewer_path:
                utils.info_print(f"Timeline viewer created: {viewer_path}")

        if getattr(args, "manifest", False):
            manifest_path = viewer.save_manifest(
                artifacts,
                new_reels=reel_records or None,
                study=study,
                participant=participant,
                worksheet_title=ws_title,
                is_excel=is_excel,
                mode=effective_mode,
                output_format=output_format,
            )
            if manifest_path:
                utils.info_print(f"Manifest updated: {manifest_path}")

    elif getattr(args, "manifest", False) and reel_records:
        study = reel_records[0].get("study", "")
        manifest_path = viewer.save_manifest(
            [],
            new_reels=reel_records,
            study=study,
            worksheet_title=ws_title,
            is_excel=is_excel,
            mode="reel",
        )
        if manifest_path:
            utils.info_print(f"Manifest updated: {manifest_path}")


# ---- Main entry point ----


_BASE_SELECTOR_ATTRS = (
    "batch",
    "lines",
    "range",
    "category",
    "cell",
    "participant",
    "keyword",
    "severity",
    "mixed",
    "reel",
    "chronologic",
    "screen",
    "gif",
    "viewer",
    "regenerate",
)


class _ModeSpec(NamedTuple):
    """Declarative description of one exclusive mode for conflict validation."""

    key: str  # attribute on args (also returned in result dict)
    truthy: Callable[[Any], bool]  # how to detect this mode is active
    error: str
    hint: str
    selector_attrs: tuple[str, ...] = _BASE_SELECTOR_ATTRS
    blocks_modes: tuple[str, ...] = ()  # earlier mode keys that conflict


_EXCLUSIVE_MODES: tuple[_ModeSpec, ...] = (
    _ModeSpec(
        key="timeline_viewer",
        truthy=lambda a: bool(getattr(a, "timeline_viewer", False)),
        error="--timeline-viewer cannot be combined with mode, format, or --viewer/--regenerate flags.",
        hint="Only -s (spreadsheet) and -v (verbose) may be used alongside --timeline-viewer.",
    ),
    _ModeSpec(
        key="studio",
        truthy=lambda a: bool(getattr(a, "studio", False)),
        error="--studio cannot be combined with mode, format, or --viewer/--regenerate flags.",
        hint="Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --studio.",
        blocks_modes=("timeline_viewer",),
    ),
    _ModeSpec(
        key="insights",
        truthy=lambda a: bool(getattr(a, "insights", False)),
        error="--insights cannot be combined with mode, format, or --viewer/--regenerate/--studio/--screenspace flags.",
        hint="Only -i/-o (directories) and -v (verbose) may be used alongside --insights.",
        blocks_modes=(
            "timeline_viewer",
            "studio",
            "screenspace",
            "transcripts",
            "export",
        ),
    ),
    _ModeSpec(
        key="screenspace",
        truthy=lambda a: bool(getattr(a, "screenspace", False)),
        error="--screenspace cannot be combined with mode, format, or --viewer/--regenerate/--studio/--insights flags.",
        hint="Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --screenspace.",
        blocks_modes=("timeline_viewer", "studio", "insights", "transcripts", "export"),
    ),
    _ModeSpec(
        key="transcripts",
        truthy=lambda a: bool(getattr(a, "transcripts", False)),
        error="--transcripts cannot be combined with mode, format, or --viewer/--regenerate/--studio/--insights/--screenspace flags.",
        hint="Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --transcripts.",
        blocks_modes=("timeline_viewer", "studio", "insights", "screenspace", "export"),
    ),
    _ModeSpec(
        key="gallery",
        # `gallery` carries an optional VIDEO arg, so use `is not None` to detect it.
        truthy=lambda a: getattr(a, "gallery", None) is not None,
        error="--gallery cannot be combined with selection modes, --viewer, --regenerate, --studio, or --timeline-viewer.",
        hint="Only --gif, --interval, --bundle, -i/-o (directories), and -v (verbose) may be used alongside --gallery.",
        # Gallery permits --screen and --gif as output-format toggles.
        selector_attrs=tuple(
            a for a in _BASE_SELECTOR_ATTRS if a not in ("screen", "gif")
        ),
        blocks_modes=(
            "timeline_viewer",
            "studio",
            "screenspace",
            "transcripts",
            "export",
        ),
    ),
    _ModeSpec(
        key="pre_transcribe",
        truthy=lambda a: getattr(a, "pre_transcribe", None) is not None,
        error="--pre-transcribe cannot be combined with mode, format, or --studio/--insights/--screenspace/--transcripts flags.",
        hint="Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --pre-transcribe.",
        # pre-transcribe additionally conflicts with --highlights.
        selector_attrs=_BASE_SELECTOR_ATTRS + ("highlights",),
        blocks_modes=(
            "timeline_viewer",
            "studio",
            "insights",
            "screenspace",
            "transcripts",
            "gallery",
            "export",
        ),
    ),
    _ModeSpec(
        key="export",
        truthy=lambda a: bool(getattr(a, "export", False)),
        error="--export cannot be combined with mode, format, or other standalone flags.",
        hint="Use --export with -i/-o (directories) and -v (verbose) only.",
        selector_attrs=_BASE_SELECTOR_ATTRS + ("highlights",),
        blocks_modes=(
            "timeline_viewer",
            "studio",
            "insights",
            "screenspace",
            "transcripts",
            "gallery",
            "pre_transcribe",
        ),
    ),
    _ModeSpec(
        key="ss_task",
        truthy=lambda a: getattr(a, "ss_task", None) is not None,
        error="--ss-task cannot be combined with mode, format, or other standalone flags.",
        hint="Use --ss-task with -i/-o (directories) and -v (verbose) only.",
        selector_attrs=_BASE_SELECTOR_ATTRS + ("highlights",),
        blocks_modes=(
            "timeline_viewer",
            "studio",
            "insights",
            "screenspace",
            "transcripts",
            "gallery",
            "pre_transcribe",
            "export",
            "ss_list_regions",
            "ss_list_stashes",
            "ss_list_tasks",
            "summarize",
            "citations",
        ),
    ),
    _ModeSpec(
        key="ss_list_regions",
        truthy=lambda a: bool(getattr(a, "ss_list_regions", False)),
        error="--ss-list-regions cannot be combined with other modes.",
        hint="Use --ss-list-regions on its own (with -i/-o for directories).",
        selector_attrs=_BASE_SELECTOR_ATTRS + ("highlights",),
        blocks_modes=(
            "timeline_viewer",
            "studio",
            "insights",
            "screenspace",
            "transcripts",
            "gallery",
            "pre_transcribe",
            "export",
            "ss_task",
            "ss_list_stashes",
            "ss_list_tasks",
            "summarize",
            "citations",
        ),
    ),
    _ModeSpec(
        key="ss_list_stashes",
        truthy=lambda a: bool(getattr(a, "ss_list_stashes", False)),
        error="--ss-list-stashes cannot be combined with other modes.",
        hint="Use --ss-list-stashes on its own (with -i/-o for directories).",
        selector_attrs=_BASE_SELECTOR_ATTRS + ("highlights",),
        blocks_modes=(
            "timeline_viewer",
            "studio",
            "insights",
            "screenspace",
            "transcripts",
            "gallery",
            "pre_transcribe",
            "export",
            "ss_task",
            "ss_list_regions",
            "ss_list_tasks",
            "summarize",
            "citations",
        ),
    ),
    _ModeSpec(
        key="ss_list_tasks",
        truthy=lambda a: getattr(a, "ss_list_tasks", None) is not None,
        error="--ss-list-tasks cannot be combined with other modes.",
        hint="Use --ss-list-tasks on its own (with -i/-o for directories).",
        selector_attrs=_BASE_SELECTOR_ATTRS + ("highlights",),
        blocks_modes=(
            "timeline_viewer",
            "studio",
            "insights",
            "screenspace",
            "transcripts",
            "gallery",
            "pre_transcribe",
            "export",
            "ss_task",
            "ss_list_regions",
            "ss_list_stashes",
            "summarize",
            "citations",
        ),
    ),
    _ModeSpec(
        key="summarize",
        truthy=lambda a: getattr(a, "summarize", None) is not None,
        error="--summarize cannot be combined with mode, format, or other standalone flags.",
        hint="Use --summarize with -i/-o (directories), -v (verbose), and --ollama-model.",
        selector_attrs=_BASE_SELECTOR_ATTRS + ("highlights",),
        blocks_modes=(
            "timeline_viewer",
            "studio",
            "insights",
            "screenspace",
            "transcripts",
            "gallery",
            "pre_transcribe",
            "export",
            "ss_task",
            "ss_list_regions",
            "ss_list_stashes",
            "ss_list_tasks",
            "citations",
        ),
    ),
    _ModeSpec(
        key="citations",
        truthy=lambda a: getattr(a, "citations", None) is not None,
        error="--citations cannot be combined with mode, format, or other standalone flags.",
        hint="Use --citations with -i/-o (directories), -v (verbose), and --ollama-model.",
        selector_attrs=_BASE_SELECTOR_ATTRS + ("highlights",),
        blocks_modes=(
            "timeline_viewer",
            "studio",
            "insights",
            "screenspace",
            "transcripts",
            "gallery",
            "pre_transcribe",
            "export",
            "ss_task",
            "ss_list_regions",
            "ss_list_stashes",
            "ss_list_tasks",
            "summarize",
        ),
    ),
)


def _validate_mode_conflicts(args: Any) -> dict[str, Any]:
    """Validate mutually exclusive mode flags and exit on conflict.

    Returns a dict keyed by mode name. Boolean for each mode plus
    ``"gallery_arg"`` (the optional VIDEO argument from --gallery).
    """
    active = {spec.key: spec.truthy(args) for spec in _EXCLUSIVE_MODES}
    for spec in _EXCLUSIVE_MODES:
        if not active[spec.key]:
            continue
        conflicts = [getattr(args, attr, None) for attr in spec.selector_attrs]
        conflicts.extend(active[m] for m in spec.blocks_modes)
        if any(conflicts):
            utils.error_print(spec.error, [spec.hint])
            sys.exit(1)

    result: dict[str, Any] = dict(active)
    result["gallery_arg"] = getattr(args, "gallery", None)
    return result


def _apply_config_overrides(args: Any, cli_mode: bool) -> CliModeArgs:
    """Apply per-run config overrides from CLI args.

    Returns parsed CLI mode arguments.
    """
    if cli_mode:
        config.VERBOSITY = config.VERBOSE if args.verbose else config.QUIET
    else:
        config.VERBOSITY = config.VERBOSE if args.verbose else config.STANDARD

    if getattr(args, "titlecards", None) is not None:
        config.TITLECARDS_ENABLED = bool(args.titlecards)
    if getattr(args, "filmstrip", None) is not None:
        config.FILMSTRIP_ENABLED = bool(args.filmstrip)
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
    if getattr(args, "whisper_model", None):
        config.TRANSCRIBE_MODEL = args.whisper_model
    if getattr(args, "ollama_model", None):
        config.OLLAMA_SUMMARY_MODEL = args.ollama_model

    return parse_cli_mode_args(args)


def _dispatch_standalone_mode(
    args: Any,
    cli_mode: bool,
    gallery_arg: Any,
) -> bool:
    """Handle standalone modes that don't need a spreadsheet.

    Returns True if a standalone mode was dispatched (caller should exit).
    """
    # Standalone viewer: regenerate viewer from saved manifest
    if getattr(args, "viewer", False) and not cli_mode:
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
        ss_events = viewer.load_screenspace_events_for_viewer()
        data = viewer.finalize_timeline_data(
            existing_artifacts,
            study=study,
            participant=participant,
            mode="manifest",
            screenspace_events=ss_events or None,
        )
        viewer_path = viewer.generate_timeline_viewer(data)
        if viewer_path:
            utils.info_print(f"Timeline viewer created from manifest: {viewer_path}")
        return True

    # Standalone insights builder
    if getattr(args, "insights", False):
        import server

        server.start_combined_server(worksheet=None, default_page="insights")
        return True

    # Standalone analysis-data export
    if getattr(args, "export", False):
        import data_export

        sys.exit(data_export.run_cli_export())

    # Standalone Screenspace CLI tasks (no UI)
    if getattr(args, "ss_list_regions", False):
        _run_ss_list_regions(args)
        return True
    if getattr(args, "ss_list_stashes", False):
        _run_ss_list_stashes(args)
        return True
    if getattr(args, "ss_list_tasks", None) is not None:
        _run_ss_list_tasks(args)
        return True
    if getattr(args, "ss_task", None) is not None:
        _run_ss_task(args)
        return True

    # Standalone thinking-agent CLI passes
    if getattr(args, "summarize", None) is not None:
        _run_summarize(args)
        return True
    if getattr(args, "citations", None) is not None:
        _run_citations(args)
        return True

    # Standalone screenspace (no spreadsheet)
    if getattr(args, "screenspace", False) and not args.spreadsheet:
        import server

        server.start_combined_server(worksheet=None, default_page="screenspace")
        return True

    # Standalone transcripts (no spreadsheet)
    if getattr(args, "transcripts", False) and not args.spreadsheet:
        import server

        server.start_combined_server(worksheet=None, default_page="transcripts")
        return True

    # Standalone gallery
    if gallery_arg is not None:
        _run_gallery_cli(args)
        return True

    # Standalone regenerate from manifest
    if getattr(args, "regenerate", False) and not cli_mode:
        existing_artifacts, existing_reels = viewer._load_manifest_both()
        if not existing_artifacts and not existing_reels:
            utils.error_print(
                "No manifest found or manifest is empty.",
                [
                    f"Run a clip generation mode with --manifest first to create {config.MANIFEST_FILENAME}."
                ],
            )
            sys.exit(1)
        media_count = sum(
            1 for a in existing_artifacts if a.get("type") != "transcript"
        )
        reel_count = len(existing_reels)
        total = media_count + reel_count
        utils.info_print(
            f"Found {media_count} media artifact(s) and {reel_count} reel(s) in manifest. Regenerating..."
        )
        regenerated = clipgen.regenerate_from_manifest(
            existing_artifacts, reels=existing_reels
        )
        utils.info_print(f"Regenerated {regenerated} of {total} item(s).")
        return True

    return False


def main() -> None:
    """Main entry point for clipgen."""
    setup_encoding()

    args = parse_arguments()
    if config.DEBUGGING:
        ic(args)

    # Validate mutually exclusive mode flags
    modes = _validate_mode_conflicts(args)
    timeline_viewer = modes["timeline_viewer"]
    gallery_arg = modes["gallery_arg"]
    pre_transcribe_mode = modes["pre_transcribe"]

    # Determine if running in CLI mode (any mode argument provided)
    mixed_selectors = getattr(args, "mixed", None)
    cli_mode = (
        args.batch
        or args.lines
        or args.range
        or args.category
        or args.cell
        or args.participant
        or args.keyword
        or args.severity
        or mixed_selectors
        or args.reel
        or args.chronologic
        or args.highlights
        or args.screen
        or args.gif
        or timeline_viewer
        or modes["ss_task"]
        or modes["ss_list_regions"]
        or modes["ss_list_stashes"]
        or modes["ss_list_tasks"]
        or modes["summarize"]
        or modes["citations"]
        or modes["export"]
        or pre_transcribe_mode
    )

    cli_mode_args = _apply_config_overrides(args, cli_mode)

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

    webp_formats = [
        name
        for name, value in (
            ("SCREENSHOT_FORMAT", config.SCREENSHOT_FORMAT),
            ("GIF_FORMAT", config.GIF_FORMAT),
        )
        if value.lower() == ".webp"
    ]
    if webp_formats and not video.check_webp_support():
        utils.error_print(
            "WebP output is configured but ffmpeg lacks libwebp support.",
            [
                f"Affected config: {', '.join(webp_formats)}",
                "Install an ffmpeg build with libwebp, or change the format(s) back to .png/.jpg/.gif in config.py.",
            ],
        )
        sys.exit(1)

    if _dispatch_standalone_mode(args, cli_mode, gallery_arg):
        sys.exit(0)

    # Authenticate with Google (once per run) – skip for local Excel files
    gspread_client = None
    doc_list: list[str] = []
    if not _is_excel_spreadsheet_arg(getattr(args, "spreadsheet", None)):
        import google_api

        gspread_client = authenticate_google()
        if gspread_client is None:
            # Auth failed. CLI mode or an explicit -s argument can't recover
            # interactively — point the user at the Excel option and exit.
            if cli_mode or getattr(args, "spreadsheet", None):
                utils.error_print(
                    "Google authentication failed.",
                    [
                        "Use -s path/to/file.xlsx to work with a local Excel file instead.",
                    ],
                )
                sys.exit(1)
            # Interactive mode: fall through with gspread_client=None; the
            # while loop below will prompt for an Excel file instead.
        else:
            doc_list = google_api.get_all_spreadsheets(gspread_client)

    # Outer loop so 'top' can return to spreadsheet selection
    while True:
        try:
            if gspread_client is None and not getattr(args, "spreadsheet", None):
                import excel_io

                worksheet = excel_io.prompt_for_excel_fallback()
                if worksheet is None:
                    sys.exit(0)
                utils.standard_print("Using local Excel file.")
            else:
                worksheet = select_worksheet(gspread_client, doc_list, args, cli_mode)

            if getattr(args, "studio", False):
                import server

                server.start_combined_server(worksheet=worksheet, default_page="studio")
                sys.exit(0)

            if getattr(args, "screenspace", False):
                import server

                server.start_combined_server(
                    worksheet=worksheet, default_page="screenspace"
                )
                sys.exit(0)

            if getattr(args, "transcripts", False):
                import server

                server.start_combined_server(
                    worksheet=worksheet, default_page="transcripts"
                )
                sys.exit(0)

            if pre_transcribe_mode:
                _run_pre_transcribe(worksheet, args)
                sys.exit(0)

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
