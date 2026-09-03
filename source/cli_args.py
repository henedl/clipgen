"""clipgen CLI argument parser.

Every ``--flag`` clipgen accepts is declared here, one ``add_argument_group``
per subject. Pure argparse: this module imports only ``config`` and ``utils``
so ``clipgen.py --help`` never loads cv2, torch, or a Flask app. Mode
detection, conflict validation, and dispatch stay in ``cli.py``, which
re-exports :func:`parse_arguments` for its callers and tests.
"""

import argparse
from typing import Any

import config
import utils


class _LicensesAction(argparse.Action):
    """Print the bundled third-party license notice and exit.

    Not `action="version"`: that action takes a *static* string, so the notice
    would be read on every single run just to build the parser. This one reads
    the ~78 KB file only when the flag is actually passed.
    """

    def __init__(
        self,
        option_strings: list[str],
        dest: str = argparse.SUPPRESS,
        default: str = argparse.SUPPRESS,
        help: str | None = None,
    ) -> None:
        super().__init__(
            option_strings=option_strings,
            dest=dest,
            default=default,
            nargs=0,
            help=help,
        )

    def __call__(
        self,
        parser: argparse.ArgumentParser,
        namespace: argparse.Namespace,
        values: Any,
        option_string: str | None = None,
    ) -> None:
        text = utils.get_licenses_text()
        if text is None:
            parser.exit(1, "THIRD-PARTY-LICENSES is missing from this installation.\n")
        # Bare print: standard_print() is verbosity-gated and Rich re-wraps verbatim
        # license text.
        print(text)
        parser.exit()


def _add_selection_args(parser: argparse.ArgumentParser) -> None:
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


def _add_output_format_args(parser: argparse.ArgumentParser) -> None:
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


def _add_transcription_args(parser: argparse.ArgumentParser) -> None:
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
        "--no-whisper-vad",
        action="store_true",
        help="Disable Silero VAD pre-filter (transcribe full audio including long silence)",
    )
    transcription.add_argument(
        "--whisper-hallucination-silence",
        type=float,
        metavar="SEC",
        help="Enable hallucination silence skip when SEC > 0 (seconds; slower, uses word timestamps)",
    )
    transcription.add_argument(
        "--summarize",
        nargs="*",
        metavar="ID",
        default=None,
        help="Run the summary thinking agent over already-transcribed participants. "
        "No IDs = all transcribed. Existing summaries are kept unless --no-input is passed.",
    )
    transcription.add_argument(
        "--citations",
        nargs="*",
        metavar="ID",
        default=None,
        help="Run the citation thinking agent over participants that already have a summary. "
        "No IDs = all eligible. Existing citations are kept unless --no-input is passed.",
    )
    transcription.add_argument(
        "--friction",
        nargs="*",
        metavar="ID",
        default=None,
        help="Run the friction thinking agent over participants that already have a summary. "
        "No IDs = all eligible. Existing friction results are kept unless --no-input is passed.",
    )


def _add_ai_args(parser: argparse.ArgumentParser) -> None:
    ai_opts = parser.add_argument_group("AI models")
    ai_opts.add_argument(
        "--llm-model",
        type=str,
        metavar="MODEL",
        help="AI model for transcript summaries, citations, and friction (HF ref or downloaded name)",
    )


def _add_path_args(parser: argparse.ArgumentParser) -> None:
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


def _add_viewer_args(parser: argparse.ArgumentParser) -> None:
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
        help="Export analysis-ready JSON+CSV from manifests in the output directory (Screenspace events, Transcripts). Skips manifests that aren't present.",
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
        "--workflows",
        action="store_true",
        help="Launch the Workflows node-canvas for chaining clip, Screenspace, and transcript actions",
    )
    viewer_manifest.add_argument(
        "--composer",
        action="store_true",
        help="Launch the Composer timeline for cutting source videos and reviewing markers",
    )
    viewer_manifest.add_argument(
        "--overview",
        action="store_true",
        help="Launch the Overview frontend (Metadata, Convergence, and the 3D similarity Map)",
    )
    viewer_manifest.add_argument(
        "--desktop",
        action="store_true",
        help="Open the web frontend in a native window instead of a browser (implies --studio when no frontend is given)",
    )
    viewer_manifest.add_argument(
        "--browser",
        action="store_true",
        help="Force the web frontend into the default browser, even for a double-clicked bundle",
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


def _add_screenspace_args(parser: argparse.ArgumentParser) -> None:
    screenspace_cli = parser.add_argument_group("screenspace cli")
    ss_modes = screenspace_cli.add_mutually_exclusive_group()
    ss_modes.add_argument(
        "--ss-task",
        nargs="+",
        metavar="TYPE PARTICIPANT [REGION]",
        default=None,
        help=(
            "Run a Screenspace analysis task headlessly. "
            "TYPE is one of color, change, similarity, text, numbers, timelapse, "
            "template, shape, flow, inactivity, scene, attention. REGION is optional "
            "and must "
            "already "
            "exist in the active manifest or in a stash (use --ss-list-regions / "
            "--ss-list-stashes); omit it (or pass 'full_frame') to scan the whole frame."
        ),
    )
    ss_modes.add_argument(
        "--ss-run-task",
        type=str,
        default=None,
        metavar="TASK_ID",
        help=(
            "Re-run a saved Screenspace task from the manifest by id, re-extracting "
            "reference frames from the source video. The only headless path for "
            "multitool tasks (build the chain in --screenspace, then run it here). "
            "Find ids with --ss-list-tasks."
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
        "--ss-color-mode",
        type=str,
        choices=("average", "presence"),
        default="average",
        help=(
            "Color tool match mode: 'average' (region average) or 'presence' "
            "(target color appears anywhere in the region). Default: average."
        ),
    )
    screenspace_cli.add_argument(
        "--ss-min-area",
        type=float,
        metavar="PERCENT",
        help=(
            "Minimum matching-pixel area as a percent (0-100) for "
            "--ss-color-mode presence. Default fires on any matching pixel."
        ),
    )
    screenspace_cli.add_argument(
        "--ss-threshold",
        type=float,
        metavar="FLOAT",
        help="Match threshold (color, change, similarity, template, shape, flow, inactivity).",
    )
    screenspace_cli.add_argument(
        "--ss-reference-timestamp",
        type=float,
        metavar="SECONDS",
        help="Reference frame timestamp (similarity, template, shape).",
    )
    screenspace_cli.add_argument(
        "--ss-scale-min",
        type=float,
        metavar="FLOAT",
        help="Shape scale ladder minimum (default from config).",
    )
    screenspace_cli.add_argument(
        "--ss-scale-max",
        type=float,
        metavar="FLOAT",
        help="Shape scale ladder maximum (default from config).",
    )
    screenspace_cli.add_argument(
        "--ss-scale-steps",
        type=int,
        metavar="INT",
        help="Shape scale ladder rungs (default from config).",
    )
    screenspace_cli.add_argument(
        "--ss-scale-y-min",
        type=float,
        metavar="FLOAT",
        help="Shape vertical scale ladder minimum (unlinks the axes).",
    )
    screenspace_cli.add_argument(
        "--ss-scale-y-max",
        type=float,
        metavar="FLOAT",
        help="Shape vertical scale ladder maximum (unlinks the axes).",
    )
    screenspace_cli.add_argument(
        "--ss-scale-y-steps",
        type=int,
        metavar="INT",
        help="Shape vertical scale ladder rungs.",
    )
    screenspace_cli.add_argument(
        "--ss-scene-ref",
        action="append",
        metavar="NAME:TIMESTAMP[:THRESHOLD]",
        help=(
            "Reference scene for the scene tool, repeatable. TIMESTAMP is in seconds; "
            "NAME must not contain ':'. Optional per-scene THRESHOLD (0-1) overrides "
            "--ss-threshold. Example: --ss-scene-ref menu:12.5 --ss-scene-ref game:30:0.8"
        ),
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


def _add_event_clip_args(parser: argparse.ArgumentParser) -> None:
    event_clips = parser.add_argument_group("event-driven clips")
    event_clips.add_argument(
        "--ss-clips",
        action="store_true",
        help=(
            "Cut clips from existing Screenspace events (reads the manifest). "
            "Filter with --ss-clips-detector / --ss-clips-region / --ss-clips-participant / "
            "--ss-clips-min-confidence / --ss-clips-event-type. Cluster nearby events with "
            "--cluster-gap and pad with --clip-pre / --clip-post."
        ),
    )
    event_clips.add_argument(
        "--transcript-clips",
        action="store_true",
        help=(
            "Cut clips from transcript segments or marks (reads the manifest). "
            "Filter with --transcript-clips-participant / --transcript-clips-mark / "
            "--transcript-clips-text. Cluster nearby segments with --cluster-gap."
        ),
    )
    event_clips.add_argument(
        "--ss-clips-detector",
        type=str,
        metavar="TYPE",
        help="Comma-separated detector types to include (e.g. 'change,color').",
    )
    event_clips.add_argument(
        "--ss-clips-region",
        type=str,
        metavar="NAME",
        help="Comma-separated region names to include.",
    )
    event_clips.add_argument(
        "--ss-clips-participant",
        type=str,
        metavar="ID",
        help="Comma-separated participant IDs to include (--ss-clips).",
    )
    event_clips.add_argument(
        "--ss-clips-min-confidence",
        type=float,
        metavar="FLOAT",
        help="Minimum event confidence (0.0-1.0).",
    )
    event_clips.add_argument(
        "--ss-clips-event-type",
        type=str,
        metavar="STR",
        help="Substring match against event_type (case-insensitive).",
    )
    event_clips.add_argument(
        "--transcript-clips-participant",
        type=str,
        metavar="ID",
        help="Comma-separated participant IDs to include (--transcript-clips).",
    )
    event_clips.add_argument(
        "--transcript-clips-mark",
        type=str,
        metavar="CATEGORY",
        help=(
            "Comma-separated mark categories. When set, only segments with at least "
            "one matching mark are included."
        ),
    )
    event_clips.add_argument(
        "--transcript-clips-text",
        type=str,
        metavar="STR",
        help="Substring match against segment text (case-insensitive).",
    )
    event_clips.add_argument(
        "--transcript-mark",
        type=str,
        metavar="TERM",
        help=(
            "Batch-mark transcript segments whose text contains TERM "
            "(case-insensitive substring, like the Transcripts UI search box). "
            "Quote multi-word terms. Requires --transcript-mark-category. "
            "Filter with --transcript-mark-participant. Optionally tag created "
            "marks with --transcript-mark-label."
        ),
    )
    event_clips.add_argument(
        "--transcript-mark-category",
        type=str,
        metavar="CATEGORY",
        help=(
            "Mark category to apply (e.g. 'pain_point', 'insight'). "
            "Must be a key in config.MARK_CATEGORIES."
        ),
    )
    event_clips.add_argument(
        "--transcript-mark-participant",
        type=str,
        metavar="ID",
        help="Comma-separated participant IDs to include (--transcript-mark). Omit to mark all.",
    )
    event_clips.add_argument(
        "--transcript-mark-label",
        type=str,
        metavar="TEXT",
        help="Optional label written onto every created or updated mark.",
    )
    event_clips.add_argument(
        "--cluster-gap",
        type=float,
        default=5.0,
        metavar="SECONDS",
        help="Cluster events whose gap is <= SECONDS into one clip (default: 5.0; 0 disables).",
    )
    event_clips.add_argument(
        "--clip-pre",
        type=float,
        default=5.0,
        metavar="SECONDS",
        help="Pad each cluster's start by SECONDS (default: 5.0).",
    )
    event_clips.add_argument(
        "--clip-post",
        type=float,
        default=5.0,
        metavar="SECONDS",
        help="Pad each cluster's end by SECONDS (default: 5.0).",
    )
    event_clips.add_argument(
        "--max-clip-duration",
        type=float,
        default=0.0,
        metavar="SECONDS",
        help="If > 0, cap each clip's duration; longer clusters are split (default: 0 = no cap).",
    )


def _add_run_args(parser: argparse.ArgumentParser) -> None:
    run_opts = parser.add_argument_group("run options")
    run_opts.add_argument(
        "--no-input",
        dest="no_input",
        action="store_true",
        help="Non-interactive mode: skip confirmation prompts and fail fast on prompts that would block on stdin (for programmatic use)",
    )
    run_opts.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Increase verbosity (-v = verbose output; default is quiet in CLI, standard in interactive mode)",
    )
    run_opts.add_argument(
        "--profile",
        action="store_true",
        help="Collect and print performance timings (grep 'profile |'; "
        "see agents/skills/profile/SKILL.md)",
    )
    run_opts.add_argument(
        "--profile-deep",
        metavar="LABEL",
        default=None,
        help="cProfile the spans whose profile label contains LABEL (implies "
        "--profile); prints a per-label function breakdown at exit",
    )
    run_opts.add_argument(
        "--settings",
        action="store_true",
        help="Open the interactive settings editor before running "
        "(changes apply to this run only; incompatible with --no-input)",
    )
    run_opts.add_argument(
        "--version",
        action="version",
        version=utils.get_version(),
        help="Print the clipgen version and exit",
    )
    run_opts.add_argument(
        "--licenses",
        action=_LicensesAction,
        help="Print third-party license notices for the bundled software and exit",
    )


def _add_titlecard_args(parser: argparse.ArgumentParser) -> None:
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


def _add_filmstrip_args(parser: argparse.ArgumentParser) -> None:
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
            f"clipgen v{utils.get_version()} - Video clip generator from Google Sheets "
            "or local Excel timestamps. "
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
  python clipgen.py -l 5 --no-input        Line mode, non-interactive (no prompts)
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
    _add_selection_args(parser)
    _add_output_format_args(parser)
    _add_transcription_args(parser)
    _add_ai_args(parser)
    _add_path_args(parser)
    _add_viewer_args(parser)
    _add_screenspace_args(parser)
    _add_event_clip_args(parser)
    _add_run_args(parser)
    _add_titlecard_args(parser)
    _add_filmstrip_args(parser)
    return parser.parse_args()
