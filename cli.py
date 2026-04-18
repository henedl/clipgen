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
from typing import Any, NamedTuple

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
        help="Launch the Insights Builder for authoring research findings from generated artifacts",
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


def authenticate_google() -> Any:
    """Authenticate with Google Sheets API.

    Returns:
        Google client connection object
    """
    import gspread

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


def _validate_mode_conflicts(
    args: Any,
) -> tuple[bool, bool, bool, bool, bool, Any, bool]:
    """Validate mutually exclusive mode flags and exit on conflict.

    Returns:
        (timeline_viewer, studio_mode, insights_mode, screenspace_mode,
         transcripts_mode, gallery_arg, pre_transcribe_mode)
    """
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
            args.keyword,
            args.severity,
            mixed_selectors,
            args.reel,
            args.chronologic,
            args.screen,
            args.gif,
            args.viewer,
            getattr(args, "regenerate", False),
        ]
        if any(conflicting):
            utils.error_print(
                "--timeline-viewer cannot be combined with mode, format, or --viewer/--regenerate flags.",
                [
                    "Only -s (spreadsheet) and -v (verbose) may be used alongside --timeline-viewer."
                ],
            )
            sys.exit(1)

    studio_mode = getattr(args, "studio", False)
    if studio_mode:
        conflicting = [
            args.batch,
            args.lines,
            args.range,
            args.category,
            args.cell,
            args.participant,
            args.keyword,
            args.severity,
            mixed_selectors,
            args.reel,
            args.chronologic,
            args.screen,
            args.gif,
            args.viewer,
            getattr(args, "regenerate", False),
            timeline_viewer,
        ]
        if any(conflicting):
            utils.error_print(
                "--studio cannot be combined with mode, format, or --viewer/--regenerate flags.",
                [
                    "Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --studio."
                ],
            )
            sys.exit(1)

    insights_mode = getattr(args, "insights", False)
    if insights_mode:
        conflicting = [
            args.batch,
            args.lines,
            args.range,
            args.category,
            args.cell,
            args.participant,
            args.keyword,
            args.severity,
            mixed_selectors,
            args.reel,
            args.chronologic,
            args.screen,
            args.gif,
            args.viewer,
            getattr(args, "regenerate", False),
            timeline_viewer,
            studio_mode,
            getattr(args, "screenspace", False),
            getattr(args, "transcripts", False),
        ]
        if any(conflicting):
            utils.error_print(
                "--insights cannot be combined with mode, format, or --viewer/--regenerate/--studio/--screenspace flags.",
                [
                    "Only -i/-o (directories) and -v (verbose) may be used alongside --insights."
                ],
            )
            sys.exit(1)

    screenspace_mode = getattr(args, "screenspace", False)
    if screenspace_mode:
        conflicting = [
            args.batch,
            args.lines,
            args.range,
            args.category,
            args.cell,
            args.participant,
            args.keyword,
            args.severity,
            mixed_selectors,
            args.reel,
            args.chronologic,
            args.screen,
            args.gif,
            args.viewer,
            getattr(args, "regenerate", False),
            timeline_viewer,
            studio_mode,
            insights_mode,
            getattr(args, "transcripts", False),
        ]
        if any(conflicting):
            utils.error_print(
                "--screenspace cannot be combined with mode, format, or --viewer/--regenerate/--studio/--insights flags.",
                [
                    "Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --screenspace."
                ],
            )
            sys.exit(1)

    transcripts_mode = getattr(args, "transcripts", False)
    if transcripts_mode:
        conflicting = [
            args.batch,
            args.lines,
            args.range,
            args.category,
            args.cell,
            args.participant,
            args.keyword,
            args.severity,
            mixed_selectors,
            args.reel,
            args.chronologic,
            args.screen,
            args.gif,
            args.viewer,
            getattr(args, "regenerate", False),
            timeline_viewer,
            studio_mode,
            insights_mode,
            screenspace_mode,
        ]
        if any(conflicting):
            utils.error_print(
                "--transcripts cannot be combined with mode, format, or --viewer/--regenerate/--studio/--insights/--screenspace flags.",
                [
                    "Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --transcripts."
                ],
            )
            sys.exit(1)

    gallery_arg = getattr(args, "gallery", None)
    if gallery_arg is not None:
        conflicting = [
            args.batch,
            args.lines,
            args.range,
            args.category,
            args.cell,
            args.participant,
            args.keyword,
            args.severity,
            mixed_selectors,
            args.reel,
            args.chronologic,
            args.viewer,
            getattr(args, "regenerate", False),
            timeline_viewer,
            studio_mode,
            screenspace_mode,
            transcripts_mode,
        ]
        if any(conflicting):
            utils.error_print(
                "--gallery cannot be combined with selection modes, --viewer, --regenerate, --studio, or --timeline-viewer.",
                [
                    "Only --gif, --interval, --bundle, -i/-o (directories), and -v (verbose) may be used alongside --gallery."
                ],
            )
            sys.exit(1)

    pre_transcribe_mode = getattr(args, "pre_transcribe", None) is not None
    if pre_transcribe_mode:
        conflicting = [
            args.batch,
            args.lines,
            args.range,
            args.category,
            args.cell,
            args.participant,
            args.keyword,
            args.severity,
            mixed_selectors,
            args.reel,
            args.chronologic,
            args.highlights,
            args.screen,
            args.gif,
            args.viewer,
            getattr(args, "regenerate", False),
            timeline_viewer,
            studio_mode,
            insights_mode,
            screenspace_mode,
            transcripts_mode,
            gallery_arg is not None,
        ]
        if any(conflicting):
            utils.error_print(
                "--pre-transcribe cannot be combined with mode, format, or --studio/--insights/--screenspace/--transcripts flags.",
                [
                    "Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --pre-transcribe."
                ],
            )
            sys.exit(1)

    return (
        timeline_viewer,
        studio_mode,
        insights_mode,
        screenspace_mode,
        transcripts_mode,
        gallery_arg,
        pre_transcribe_mode,
    )


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
    (
        timeline_viewer,
        studio_mode,
        insights_mode,
        screenspace_mode,
        transcripts_mode,
        gallery_arg,
        pre_transcribe_mode,
    ) = _validate_mode_conflicts(args)

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

    if _dispatch_standalone_mode(args, cli_mode, gallery_arg):
        sys.exit(0)

    # Authenticate with Google (once per run) – skip for local Excel files
    gspread_client = None
    doc_list: list[str] = []
    if not _is_excel_spreadsheet_arg(getattr(args, "spreadsheet", None)):
        import google_api

        gspread_client = authenticate_google()
        doc_list = google_api.get_all_spreadsheets(gspread_client)

    # Outer loop so 'top' can return to spreadsheet selection
    while True:
        try:
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
