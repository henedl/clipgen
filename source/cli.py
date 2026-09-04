"""CLI entry point and argument parsing for clipgen.

Handles command-line argument parsing, CLI mode detection, setup,
and CLI-specific dispatch. The main() function is the program entry point,
called from clipgen.py's __main__ guard.
"""

import argparse
import copy
import io
import os
import sys
import uuid
from collections.abc import Callable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, NamedTuple

import app
import config
import files
import profiling
import spreadsheet
import transcripts
import utils
import video
import viewer
from cli_args import parse_arguments
from utils import ClipRecord


# ---- CLI data structures ----


class CliModeArgs(NamedTuple):
    line_numbers: list[int] | None
    range_start: int | None
    range_end: int | None
    cell_specs: list[tuple[str, int]] | None


# ---- Argument parsing ----


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
            cli_line_numbers = [
                int(num) for num in utils.split_selector_tokens(args.lines)
            ]
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

    Source runs use the repo root — the directory holding ``clipgen.py``, one
    level above this module's ``source/`` — because that is the directory the
    user is working in. ``main`` chdirs here, so it decides where
    ``credentials.json`` is looked up, where ``.xlsx`` files are discovered, and
    what the default output directory is.

    Frozen builds use the directory the user sees the application in, so "put
    credentials.json next to the app" means what it says:

    * macOS ``.app`` — the folder *containing* the bundle. ``sys.executable``
      points inside ``Contents/MacOS``, which is invisible in Finder and part of
      the code signature, so writing there would both hide files from the user
      and invalidate the signature.
    * a one-dir build (Windows/Linux) — the folder containing the payload
      directory. One-dir puts ``clipgen.exe`` *inside* ``clipgen/`` alongside
      ``lib/``; the folder the user actually dragged out of the zip is its
      parent, so that is where they will drop ``credentials.json``.
    * a one-dir *installer* copy — Inno Setup lands that same layout at
      ``%LOCALAPPDATA%\\Programs\\clipgen``. Walking up one more level would
      put cwd in ``Programs`` itself, so stay in ``{app}``. Detected by an
      ``unins000.exe`` sibling (Inno's uninstaller) or a parent named
      ``Programs``.
    * anything else — the executable's own directory.
    """
    if not getattr(sys, "frozen", False):
        return str(Path(__file__).resolve().parent.parent)

    exe_dir = Path(sys.executable).resolve().parent
    # .../clipgen.app/Contents/MacOS/clipgen → .../  (the .app's parent)
    if exe_dir.name == "MacOS" and exe_dir.parent.name == "Contents":
        bundle = exe_dir.parent.parent
        if bundle.suffix == ".app":
            return str(bundle.parent)
    # One-dir build: _MEIPASS is lib/ under the exe dir. One-file's temp dir never
    # matches.
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass and Path(meipass).resolve().parent == exe_dir:
        # Portable zip: parent of the exe dir. Inno install: stay in {app}, not
        # Programs.
        if (exe_dir / "unins000.exe").is_file() or exe_dir.parent.name.lower() == (
            "programs"
        ):
            return str(exe_dir)
        return str(exe_dir.parent)
    return str(exe_dir)


# ---- Google authentication ----

CREDENTIALS_FILENAME = "credentials.json"


def credentials_search_paths() -> list[Path]:
    """Return the locations searched for ``credentials.json``, in priority order.

    1. The working directory — for a frozen bundle that means *beside the app*,
       which is where users are told to drop the file.
    2. gspread's own config dir (``~/.config/gspread``), where many existing
       installs already keep it.
    3. clipgen's per-user config dir, alongside ``start.json``.
    """
    paths = [Path.cwd() / CREDENTIALS_FILENAME]
    try:
        from gspread.auth import DEFAULT_CONFIG_DIR

        paths.append(Path(DEFAULT_CONFIG_DIR) / CREDENTIALS_FILENAME)
    except ImportError:
        pass
    import start_settings

    paths.append(start_settings.config_dir() / CREDENTIALS_FILENAME)
    return paths


def resolve_credentials_path() -> Path | None:
    """Return the first existing ``credentials.json``, or None if there is none."""
    for candidate in credentials_search_paths():
        if candidate.is_file():
            return candidate
    return None


def _cached_token_path() -> Path | None:
    """Return the cached gspread token path if it exists.

    gspread always writes the token to an absolute path under its own config
    dir (clipgen never overrides ``authorized_user_filename``), so this is the
    single place it can be.
    """
    try:
        from gspread.auth import DEFAULT_AUTHORIZED_USER_FILENAME
    except ImportError:
        return None
    token = Path(DEFAULT_AUTHORIZED_USER_FILENAME)
    return token if token.is_file() else None


def _try_silent_google_auth() -> Any | None:
    """Reuse a cached gspread token without ever launching the OAuth flow.

    Returns a gspread client when the cached ``authorized_user.json`` exists and
    loads; otherwise returns ``None`` without printing or prompting. Used by the
    frozen-binary launch path so a previously-authenticated user is not forced
    through "Connect Google" on every double-click — but a fresh install lands
    silently on the Start overlay's Connect CTA rather than blocking on an
    interactive flow.

    The gate is the *token*, not ``credentials.json``: gspread only reads the
    credentials file when no cached token exists, so requiring both would push
    users with a perfectly good token back through the connect flow purely
    because their credentials file lives somewhere else.
    """
    try:
        import gspread
    except ImportError:
        return None
    if _cached_token_path() is None:
        return None
    credentials = resolve_credentials_path()
    try:
        if credentials is None:
            return gspread.oauth()
        return gspread.oauth(credentials_filename=str(credentials))
    except Exception:
        return None


def authenticate_google() -> Any | None:
    """Authenticate with Google Sheets API.

    Returns:
        Google client connection object on success, or None if authentication
        failed. Callers are responsible for deciding whether to fall back to a
        local Excel file or exit.
    """
    import gspread

    credentials = resolve_credentials_path()
    try:
        utils.debug_print("Attempting login...")
        if credentials is None:
            # No file anywhere: gspread's own error names its default path.
            gspread_client = gspread.oauth()
        else:
            utils.debug_print(f"Using credentials at {credentials}")
            gspread_client = gspread.oauth(credentials_filename=str(credentials))
        utils.debug_print("Login successful!")
        return gspread_client
    # ValueError covers json.JSONDecodeError from a malformed credentials.json; the
    # hint below addresses it.
    except (
        gspread.exceptions.GSpreadException,
        FileNotFoundError,
        OSError,
        ValueError,
    ) as e:
        searched = [f"  - {p}" for p in credentials_search_paths()]
        utils.error_print(
            "Could not authenticate with Google.",
            [
                f"Error details: {e}",
                "",
                "Searched for credentials.json in:",
                *searched,
                "",
                "Troubleshooting steps:",
                "  1. Put 'credentials.json' in one of the locations above",
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


def _single_xlsx_fallback_path(reason: str) -> str | None:
    """Path of the sole .xlsx in the working directory, or None.

    CLI mode can't prompt, so when Google Sheets is unavailable and exactly
    one local Excel file is present, fall back to it (with a notice) instead
    of dead-ending. Ambiguous directories (zero or several .xlsx) return None
    and leave the caller's error path in charge.
    """
    import excel_io

    paths = excel_io.list_excel_in_cwd()
    if len(paths) != 1:
        return None
    utils.info_print(
        f"{reason}; falling back to local Excel file {Path(paths[0]).name}. "
        "Pass -s to choose a source explicitly."
    )
    return paths[0]


def select_worksheet(gspread_client: Any, args: Any, cli_mode: bool) -> Any:
    """Select worksheet based on command-line arguments or interactive selection.

    Args:
        gspread_client: Google client connection
        args: Parsed command-line arguments
        cli_mode: Whether running in CLI mode

    Returns:
        Worksheet object
    """
    import excel_io
    import google_api

    _doc_list_cache: list[list[str]] = []

    def get_doc_list() -> list[str]:
        # Rate-limited Drive listing: fetch once, and only when a path needs it.
        if not _doc_list_cache:
            _doc_list_cache.append(google_api.get_all_spreadsheets(gspread_client))
        return _doc_list_cache[0]

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
            worksheet = app.open_spreadsheet_by_url(gspread_client, args.spreadsheet)
        elif args.spreadsheet.isdigit():
            worksheet = app.open_spreadsheet_by_index(
                gspread_client, get_doc_list(), int(args.spreadsheet)
            )
        else:
            worksheet = app.open_spreadsheet_by_name(
                gspread_client, get_doc_list(), args.spreadsheet
            )

        if not worksheet:
            utils.error_print(
                f'Could not find or open spreadsheet "{args.spreadsheet}"'
            )
            sys.exit(1)
    else:
        # Auto-connect if working directory name matches a spreadsheet
        cwd_name = Path.cwd().name
        worksheet = app.open_spreadsheet_by_name(
            gspread_client,
            get_doc_list(),
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
            fallback = _single_xlsx_fallback_path(
                "No spreadsheet found matching working directory name"
            )
            if fallback:
                worksheet = excel_io.open_excel_workbook(fallback)
            if not worksheet:
                utils.error_print(
                    "No spreadsheet found matching working directory name.",
                    [
                        "Use -s to specify a spreadsheet name, URL, or index.",
                        "Or use -s excel / -s path/to/file.xlsx for a local Excel file.",
                    ],
                )
                sys.exit(1)
        else:
            worksheet = app.select_spreadsheet(gspread_client, get_doc_list())

    if worksheet and config.DEBUGGING:
        config.debug_ic(worksheet.title)
    if app._is_excel_worksheet(worksheet):
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
    skip_prompts = args.no_input
    mixed_selectors = getattr(args, "mixed", None)
    output_format = "screen" if args.screen else "gif" if args.gif else "clip"

    selection_mode_set = bool(
        any(getattr(args, a, None) for a in _SELECTION_ATTRS) or args.highlights
    )

    def _parse_cli_categories(raw: str | None) -> list[str]:
        """Parse CLI category string into a list of category names."""
        if not raw:
            return []
        seen = set()
        result: list[str] = []
        for name in utils.split_selector_tokens(raw):
            if name not in seen:
                seen.add(name)
                result.append(name)
        return result

    cli_categories = _parse_cli_categories(getattr(args, "category", None))

    def _parse_cli_severities(raw: str | None) -> list[str]:
        if not raw:
            return []
        seen = set()
        result: list[str] = []
        for token in utils.split_selector_tokens(raw):
            name = utils.normalize_severity(token)
            if not name:
                continue
            if name not in seen:
                seen.add(name)
                result.append(name)
        return result

    cli_severities = _parse_cli_severities(getattr(args, "severity", None))

    cli_annotation_ids = None
    if isinstance(args.keyword, str):
        cli_annotation_ids = [
            t.lower().lstrip("!") for t in utils.split_selector_tokens(args.keyword)
        ] or None

    # Apply custom highlights duration if specified (e.g. -H 120)
    if args.highlights and args.highlights != "highlights":
        try:
            duration = int(args.highlights)
        except ValueError:
            utils.warning_print(
                f"Invalid highlights duration '{args.highlights}', using default ({config.HIGHLIGHTS_REEL_DURATION_SECONDS}s)."
            )
        else:
            if duration <= 0:
                utils.warning_print(
                    f"Invalid highlights duration '{args.highlights}', using default ({config.HIGHLIGHTS_REEL_DURATION_SECONDS}s)."
                )
            else:
                config.HIGHLIGHTS_REEL_DURATION_SECONDS = duration

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
    participant_id = utils.normalize_participant_id(args.chronologic)
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
    # Mirrors the interactive gallery guard; a bare ``or`` would pass a negative
    # through.
    interval = getattr(args, "interval", None)
    if interval is None or interval <= 0:
        interval = config.GALLERY_INTERVAL_SECONDS
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
            pid = utils.normalize_participant_id(raw_id)
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
    known_terms = transcripts.get_known_terms(manifest) or None

    skipped = 0
    transcribed = 0
    overrides = spreadsheet.participant_filename_overrides(ctx)

    for pid in target_ids:
        if source_transcripts.get(pid, {}).get("segments"):
            utils.info_print(f"Skipping {pid}: already transcribed.")
            skipped += 1
            continue

        source_paths = files.resolve_source_video_paths(
            ctx.study_name, pid, overrides.get(pid), utils.get_effective_input_dir()
        )
        missing = [p for p in source_paths if not p.is_file()]
        if missing:
            utils.error_print(
                f"Source video not found for {pid}: "
                + ", ".join(str(p) for p in missing)
            )
            continue

        names = ", ".join(p.name for p in source_paths)
        utils.info_print(f"Transcribing {pid}: {names}...")
        if len(source_paths) == 1:
            # Single-video fast path: transcribe directly, no duration probe.
            result = transcripts.transcribe_video(
                str(source_paths[0]),
                context_keywords=context_keywords,
                known_terms=known_terms,
            )
        else:
            # Multi-video: shift each part's times by its cumulative start onto the
            # global timeline.
            timeline = video.build_source_timeline([str(p) for p in source_paths])
            result = (
                None
                if timeline is None
                else transcripts.transcribe_timeline(
                    timeline,
                    context_keywords=context_keywords,
                    known_terms=known_terms,
                )
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
            "transcribed_at": datetime.now(UTC).isoformat(),
        }
        if config.TRANSCRIBE_SPEAKERS:
            block = transcripts.label_speakers(
                [str(p) for p in source_paths],
                source_transcripts[pid]["segments"],
                None,
            )
            if block:
                source_transcripts[pid]["speakers"] = block
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
    "shape",
    "flow",
    "inactivity",
    "scene",
    "attention",
)


def _ss_resolve_videos_for_participant(participant_id: str) -> list[str]:
    """Resolve a participant's ordered source video path(s) via filename discovery.

    Honours ``config.FILENAME_OVERRIDES`` the same way the web tools do.
    A multi-video participant (numbered parts) returns all parts in timeline order;
    a normal participant returns a single-element list. Returns [] when unknown.
    """
    for entry in files.resolve_participant_videos():
        if entry["id"] == participant_id and entry.get("has_video"):
            return list(entry["video_paths"])
    return []


def _ss_frame_extractor(video_paths: list[str]) -> Callable[[float], "Any | None"]:
    """Return a ``frame_at(global_ts)`` closure mapping into the right sub-video.

    For reference-frame extraction (similarity/template/scene) over a participant
    whose recording spans several files: a global reference timestamp resolves to
    the owning sub-video. Single-video participants extract unchanged (no probe).
    """
    timeline = video.timeline_or_none(video_paths)

    def _extract(global_ts: float) -> "Any | None":
        if timeline is None:
            return video.extract_frame_at_timestamp(video_paths[0], global_ts)
        mapped = utils.map_global_to_segment(timeline, global_ts)
        if mapped is None:
            return None
        return video.extract_frame_at_timestamp(timeline[mapped[0]][0], mapped[1])

    return _extract


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


def _ss_parse_scene_ref(raw: str) -> dict[str, Any]:
    """Parse a NAME:TIMESTAMP[:THRESHOLD] scene reference (TIMESTAMP in seconds).

    Returns ``{"name", "timestamp"}`` plus optional ``"threshold"``. NAME must not
    contain ``:``; TIMESTAMP and THRESHOLD must be numeric.
    """
    parts = raw.split(":")
    if len(parts) not in (2, 3):
        raise ValueError(
            f"Scene reference must be NAME:TIMESTAMP[:THRESHOLD] (got {raw!r})"
        )
    name = parts[0].strip()
    if not name:
        raise ValueError(f"Scene reference NAME is required (got {raw!r})")
    try:
        ref: dict[str, Any] = {"name": name, "timestamp": float(parts[1])}
        if len(parts) == 3:
            ref["threshold"] = float(parts[2])
    except ValueError as exc:
        raise ValueError(
            f"Scene reference TIMESTAMP/THRESHOLD must be numeric (got {raw!r})"
        ) from exc
    if "threshold" in ref and not (0.0 <= ref["threshold"] <= 1.0):
        raise ValueError(
            f"Scene reference THRESHOLD must be between 0 and 1 (got {raw!r})"
        )
    return ref


def _ss_build_params(
    args: argparse.Namespace,
    task_type: str,
    region_coords: dict[str, int],
    frame_at: Callable[[float], "Any | None"],
) -> dict[str, Any]:
    """Build a `parameters` dict for create_task() from per-tool CLI flags.

    Validates that the required flags for ``task_type`` are present. For
    ``similarity`` and ``template`` extracts the reference frame at
    ``--ss-reference-timestamp`` via *frame_at* (which maps a global timestamp
    into the owning sub-video for multi-video participants; mirrors the
    server-side path in screenspace_server._extract_tool_media).
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
        if args.ss_color_mode == "presence":
            params["color_mode"] = "presence"
            if args.ss_min_area is not None:
                params["min_coverage"] = args.ss_min_area / 100.0

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
        frame = frame_at(float(args.ss_reference_timestamp))
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
        frame = frame_at(float(args.ss_reference_timestamp))
        if frame is None:
            raise ValueError(
                f"Could not extract template frame at {args.ss_reference_timestamp}s"
            )
        params["template_image"] = screenspace.extract_region(frame, region_coords)

    elif task_type == "shape":
        if args.ss_reference_timestamp is None:
            raise ValueError("shape task requires --ss-reference-timestamp SECONDS")
        params["reference_timestamp"] = args.ss_reference_timestamp
        if args.ss_threshold is not None:
            params["threshold"] = args.ss_threshold
        if args.ss_scale_min is not None:
            params["scale_min"] = args.ss_scale_min
        if args.ss_scale_max is not None:
            params["scale_max"] = args.ss_scale_max
        if args.ss_scale_steps is not None:
            params["scale_steps"] = args.ss_scale_steps
        if args.ss_scale_y_min is not None:
            params["scale_y_min"] = args.ss_scale_y_min
        if args.ss_scale_y_max is not None:
            params["scale_y_max"] = args.ss_scale_y_max
        if args.ss_scale_y_steps is not None:
            params["scale_y_steps"] = args.ss_scale_y_steps
        frame = frame_at(float(args.ss_reference_timestamp))
        if frame is None:
            raise ValueError(
                f"Could not extract shape frame at {args.ss_reference_timestamp}s"
            )
        params["shape_image"] = screenspace.extract_region(frame, region_coords)

    elif task_type == "scene":
        raw_refs = getattr(args, "ss_scene_ref", None) or []
        if not raw_refs:
            raise ValueError(
                "scene task requires at least one --ss-scene-ref NAME:TIMESTAMP[:THRESHOLD]"
            )
        parsed = [_ss_parse_scene_ref(r) for r in raw_refs]
        reference_scenes = []
        for ref in parsed:
            frame = frame_at(float(ref["timestamp"]))
            if frame is None:
                raise ValueError(
                    f"Could not extract scene frame for {ref['name']!r} "
                    f"at {ref['timestamp']}s"
                )
            entry: dict[str, Any] = {
                "name": ref["name"],
                "frame": screenspace.extract_region(frame, region_coords),
            }
            if "threshold" in ref:
                entry["threshold"] = ref["threshold"]
            reference_scenes.append(entry)
        params["reference_scenes"] = reference_scenes
        # Frames are stripped on manifest save; the input form keeps --ss-run-task
        # working.
        params["scene_references"] = parsed
        if args.ss_threshold is not None:
            params["threshold"] = args.ss_threshold

    elif task_type == "flow":
        if args.ss_threshold is None:
            raise ValueError("flow task requires --ss-threshold FLOAT (magnitude)")
        params["magnitude_threshold"] = args.ss_threshold

    elif task_type == "inactivity":
        if args.ss_threshold is None:
            raise ValueError("inactivity task requires --ss-threshold FLOAT")
        params["threshold"] = args.ss_threshold

    elif task_type == "attention":
        # All knobs have config defaults; --ss-threshold optionally overrides
        # the normalized peak-jump distance for shift events.
        if args.ss_threshold is not None:
            params["shift_threshold"] = args.ss_threshold

    params.setdefault("cv_resolution_scale", config.SCREENSPACE_CV_RESOLUTION_SCALE)
    return params


def _print_ss_table(
    title: str,
    columns: list[tuple[str, dict[str, Any]]],
    rows: list[list[str]],
    fallback_lines: list[str],
) -> None:
    """Render a Rich table when available, else fall back to plain info_print lines.

    columns is a list of (name, kwargs) tuples passed to Table.add_column.
    """
    if utils._use_rich() and utils.console is not None:
        from rich.table import Table

        table = Table(
            title=title,
            show_header=True,
            header_style="bold cyan",
            border_style="dim",
            row_styles=["", "dim"],
            expand=False,
        )
        for name, kwargs in columns:
            table.add_column(name, **kwargs)
        for row in rows:
            table.add_row(*row)
        utils.console.print(table)
    else:
        for line in fallback_lines:
            utils.info_print(line)


def _run_ss_list_regions(args: argparse.Namespace) -> None:
    """List active Screenspace regions from the manifest."""
    import screenspace

    manifest = screenspace.load_screenspace_manifest()
    regions = manifest.get("regions", {})
    if not regions:
        utils.info_print("No active Screenspace regions in manifest.")
        return

    rows: list[list[str]] = []
    fallback: list[str] = [f"Active regions ({len(regions)}):"]
    for name in sorted(regions.keys()):
        rd = regions[name]
        sw = rd.get("source_width", "?")
        sh = rd.get("source_height", "?")
        rows.append(
            [
                name,
                f"{rd.get('x', 0):.3f}",
                f"{rd.get('y', 0):.3f}",
                f"{rd.get('w', 0):.3f}",
                f"{rd.get('h', 0):.3f}",
                f"{sw}x{sh}",
            ]
        )
        fallback.append(
            f"  {name}: x={rd.get('x', 0):.3f} y={rd.get('y', 0):.3f} "
            f"w={rd.get('w', 0):.3f} h={rd.get('h', 0):.3f}  source={sw}x{sh}"
        )

    _print_ss_table(
        f"Active regions ({len(regions)})",
        [
            ("Name", {"style": "bold"}),
            ("x", {"justify": "right"}),
            ("y", {"justify": "right"}),
            ("w", {"justify": "right"}),
            ("h", {"justify": "right"}),
            ("Source", {"justify": "right"}),
        ],
        rows,
        fallback,
    )


def _run_ss_list_stashes(args: argparse.Namespace) -> None:
    """List Screenspace region stashes."""
    import screenspace

    manifest = screenspace.load_screenspace_manifest()
    stashes = manifest.get("stashes", [])
    if not stashes:
        utils.info_print("No Screenspace region stashes in manifest.")
        return

    rows: list[list[str]] = []
    fallback: list[str] = [f"Stashes ({len(stashes)}):"]
    for stash in stashes:
        regions = stash.get("regions", {})
        names = ", ".join(sorted(regions.keys())) or "(empty)"
        rows.append(
            [
                str(stash.get("id", "?")),
                str(stash.get("name", "(unnamed)")),
                str(len(regions)),
                names,
            ]
        )
        fallback.append(
            f"  {stash.get('id', '?')}  {stash.get('name', '(unnamed)')}: "
            f"{len(regions)} region(s) — {names}"
        )

    _print_ss_table(
        f"Stashes ({len(stashes)})",
        [
            ("ID", {"style": "bold"}),
            ("Name", {}),
            ("Regions", {"justify": "right"}),
            ("Names", {"overflow": "fold"}),
        ],
        rows,
        fallback,
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
    rows: list[list[str]] = []
    fallback: list[str] = [f"Tasks{label}: {len(tasks)}"]
    for t in tasks:
        result = t.get("result")
        result_count = len(result) if isinstance(result, list) else 0
        rows.append(
            [
                str(t.get("id", "?")),
                str(t.get("type", "?")),
                str(t.get("participant", "?")),
                str(t.get("region", "?")),
                str(t.get("status", "?")),
                str(result_count),
            ]
        )
        fallback.append(
            f"  {t.get('id', '?')}  {t.get('type', '?'):10s}  "
            f"{t.get('participant', '?'):8s}  region={t.get('region', '?'):16s}  "
            f"status={t.get('status', '?'):10s}  results={result_count}"
        )

    _print_ss_table(
        f"Tasks{label} ({len(tasks)})",
        [
            ("ID", {"style": "bold"}),
            ("Type", {}),
            ("Participant", {}),
            ("Region", {}),
            ("Status", {}),
            ("Results", {"justify": "right"}),
        ],
        rows,
        fallback,
    )


def _run_ss_task(args: argparse.Namespace) -> None:
    """Run a Screenspace analysis task synchronously and persist the result."""
    import screenspace

    # REGION is optional; two args scan the full frame.
    ss_task = list(args.ss_task)
    if len(ss_task) == 2:
        ss_task.append(screenspace.FULL_FRAME_REGION_NAME)
    if len(ss_task) != 3:
        utils.error_print(
            "--ss-task takes TYPE PARTICIPANT [REGION].",
            ["Omit REGION (or pass 'full_frame') to scan the whole frame."],
        )
        sys.exit(1)
    task_type, participant, region_name = ss_task
    if task_type not in _SS_VALID_TASK_TYPES:
        utils.error_print(
            f"Unknown screenspace task type {task_type!r}.",
            [f"Valid types: {', '.join(_SS_VALID_TASK_TYPES)}"],
        )
        sys.exit(1)
    if task_type == "attention" and region_name != screenspace.FULL_FRAME_REGION_NAME:
        # Attention is full-frame only (the server forces this too); a region would
        # mislabel events.
        utils.warning_print(
            f"attention is full-frame only; ignoring region {region_name!r}."
        )
        region_name = screenspace.FULL_FRAME_REGION_NAME

    manifest = screenspace.load_screenspace_manifest()

    # Active regions win over same-named stashed copies; same resolver as
    # --ss-run-task.
    try:
        _, resolved_region = screenspace.resolve_region_request(
            region_name, None, manifest
        )
    except ValueError:
        # The hint lists every known name, active and stashed.
        available = sorted(
            {
                *manifest.get("regions", {}),
                *(
                    name
                    for stash in manifest.get("stashes", [])
                    for name in stash.get("regions", {})
                ),
            }
        )
        hint = (
            f"Available regions: {', '.join(available)}"
            if available
            else "No regions defined. Use --screenspace to define regions in the web UI first."
        )
        utils.error_print(f"Region {region_name!r} not found.", [hint])
        sys.exit(1)

    video_paths = _ss_resolve_videos_for_participant(participant)
    if not video_paths:
        utils.error_print(
            f"No video found for participant {participant!r}.",
            ["Place the source video in the input directory before running --ss-task."],
        )
        sys.exit(1)

    # Parts share resolution; reference frames map global→sub-video via frame_at.
    props = video.probe_video_properties(video_paths[0])
    if props and props.get("width") and props.get("height"):
        region_coords = screenspace.denormalize_region(
            resolved_region, int(props["width"]), int(props["height"])
        )
    else:
        region_coords = {
            k: int(resolved_region[k])
            for k in ("x", "y", "w", "h")
            if k in resolved_region
        }

    try:
        parameters = _ss_build_params(
            args, task_type, region_coords, _ss_frame_extractor(video_paths)
        )
    except ValueError as exc:
        utils.error_print(str(exc))
        sys.exit(1)

    source_video = Path(video_paths[0]).name
    task = screenspace.create_task(
        task_type=task_type,
        participant=participant,
        source_video=source_video,
        video_paths=video_paths,
        region_name=region_name,
        region_coords=region_coords,
        parameters=parameters,
    )

    _ss_run_and_persist_task(task, manifest)


def _ss_run_and_persist_task(task: dict[str, Any], manifest: dict[str, Any]) -> None:
    """Enqueue a task, poll to completion, persist the manifest, and report.

    Shared by --ss-task (flag-built tasks) and --ss-run-task (manifest re-runs).
    Restores the manifest's historical tasks so they survive the save, then enqueues
    only ``task`` for execution.
    """
    import time

    import screenspace

    task_type = task.get("type", "")
    participant = task.get("participant", "")
    region_name = task.get("region", "")

    worker = screenspace.ScreenspaceWorker()
    worker.restore_tasks(manifest.get("tasks", []))
    worker.start()
    task_id = worker.enqueue(task)
    utils.info_print(f"Running {task_type} on {participant} (region: {region_name})...")

    final_task: dict[str, Any] | None = None
    try:
        with utils.progress_scope(f"{task_type} (queued)", 100) as ps:
            last_progress = -1.0
            while True:
                current = worker.get_task(task_id)
                if current is None:
                    break
                status = current.get("status", "")
                progress = float(current.get("progress", 0.0))
                if ps.live:
                    ps.update(
                        completed=int(progress * 100),
                        description=f"{task_type} ({status})",
                    )
                elif progress - last_progress > 0.05:
                    utils.info_print(f"  {status}: {int(progress * 100)}%")
                    last_progress = progress
                if status in ("completed", "failed", "cancelled"):
                    final_task = current
                    ps.update(completed=100)
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
            per_participant=manifest.get("per_participant", {}),
            pins=manifest.get("pins") or {},
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


def _ss_extract_scene_frames(
    scene_refs: list[dict[str, Any]],
    frame_at: Callable[[float], "Any | None"],
    region_coords: dict[str, int],
    *,
    context: str = "",
) -> list[dict[str, Any]]:
    """Build reference_scenes (with cropped frames) from saved scene_references.

    Mirrors the scene path of screenspace_server._extract_tool_media. *frame_at*
    maps a global timestamp into the owning sub-video for multi-video
    participants. Raises ValueError when a frame cannot be read.
    """
    import screenspace

    reference_scenes: list[dict[str, Any]] = []
    for ref in scene_refs:
        frame = frame_at(float(ref["timestamp"]))
        if frame is None:
            raise ValueError(
                f"{context}could not read frame for scene {ref.get('name')!r} "
                f"at {ref.get('timestamp')}s"
            )
        entry: dict[str, Any] = {
            "name": ref["name"],
            "frame": screenspace.extract_region(frame, region_coords),
        }
        if "threshold" in ref:
            entry["threshold"] = ref["threshold"]
        reference_scenes.append(entry)
    return reference_scenes


def _ss_reference_coords(
    parameters: dict[str, Any],
    region_coords: dict[str, Any],
    manifest: dict[str, Any],
    dims: tuple[int, int] | None,
) -> dict[str, Any]:
    """Pixel rect a template/shape re-run cuts its sample from.

    The persisted capture region (``reference_region``) wins over the run
    region, mirroring screenspace_server._prepare_task_media.
    """
    import screenspace

    ref_name = str(parameters.get("reference_region") or "").strip()
    if not ref_name:
        return region_coords
    _rn, ref_norm = screenspace.resolve_region_request(ref_name, None, manifest)
    if dims is None:
        return region_coords
    return screenspace.denormalize_region(ref_norm, dims[0], dims[1])


def _ss_rehydrate_task_media(
    task_type: str,
    parameters: dict[str, Any],
    frame_at: Callable[[float], "Any | None"],
    region_coords: dict[str, Any],
    manifest: dict[str, Any],
    dims: tuple[int, int] | None,
) -> None:
    """Re-extract reference frames/templates/scenes into a saved task's parameters.

    Mirrors screenspace_server._extract_tool_media and _prepare_multitool_steps so a
    manifest task (whose binary frame data was stripped on save) can be re-run.
    *frame_at* maps a global reference timestamp into the owning sub-video for
    multi-video participants. Mutates ``parameters`` in place; raises ValueError
    when a reference cannot be recovered (e.g. a multitool step built from an
    uploaded template image, which has no timestamp to re-extract from).
    """
    import screenspace

    def _extract_frame(ref_ts: float, coords: dict[str, Any], label: str) -> Any:
        frame = frame_at(float(ref_ts))
        if frame is None:
            raise ValueError(f"{label}: could not read reference frame")
        return screenspace.extract_region(frame, coords)

    if task_type == "similarity":
        if parameters.get("reference_timestamp") is None:
            raise ValueError("similarity task has no reference_timestamp to re-extract")
        parameters["reference_frame"] = _extract_frame(
            parameters["reference_timestamp"], region_coords, "similarity"
        )

    elif task_type == "template":
        if parameters.get("reference_timestamp") is None:
            raise ValueError(
                "template task built from an uploaded image cannot be re-run from the "
                "manifest (no reference timestamp was saved)"
            )
        coords = _ss_reference_coords(parameters, region_coords, manifest, dims)
        parameters["template_image"] = _extract_frame(
            parameters["reference_timestamp"],
            coords,
            "template",
        )
        screenspace.attach_capture_mask(
            parameters, "template_image", "template_mask", coords
        )

    elif task_type == "shape":
        if parameters.get("reference_timestamp") is None:
            raise ValueError(
                "shape task built from an uploaded image cannot be re-run from the "
                "manifest (no reference timestamp was saved)"
            )
        coords = _ss_reference_coords(parameters, region_coords, manifest, dims)
        parameters["shape_image"] = _extract_frame(
            parameters["reference_timestamp"],
            coords,
            "shape",
        )
        screenspace.attach_capture_mask(parameters, "shape_image", "shape_mask", coords)

    elif task_type == "scene":
        scene_refs = parameters.get("scene_references")
        if not scene_refs:
            raise ValueError("scene task has no scene_references to re-extract")
        parameters["reference_scenes"] = _ss_extract_scene_frames(
            scene_refs, frame_at, region_coords
        )

    elif task_type == "multitool":
        steps: list[dict[str, Any]] = parameters.get("steps", [])
        for i, step in enumerate(steps):
            stype = step.get("type", "")
            step_region_name = (step.get("region") or "").strip()
            step_region_ref = step.get("region_ref")
            if step_region_name or step_region_ref is not None:
                resolved_name, resolved_region = screenspace.resolve_region_request(
                    step_region_name, step_region_ref, manifest
                )
                step["region"] = resolved_name
                if dims is not None:
                    step_coords = screenspace.denormalize_region(
                        resolved_region, dims[0], dims[1]
                    )
                else:
                    step_coords = region_coords
            else:
                step_coords = region_coords
            step["region_coords"] = step_coords

            if stype == "similarity":
                if step.get("reference_timestamp") is None:
                    raise ValueError(f"Step {i}: no reference_timestamp to re-extract")
                step["reference_frame"] = _extract_frame(
                    step["reference_timestamp"], step_coords, f"Step {i}"
                )
            elif stype == "template":
                if step.get("reference_timestamp") is None:
                    raise ValueError(
                        f"Step {i}: template step built from an uploaded image cannot "
                        "be re-run from the manifest (no reference timestamp saved)"
                    )
                step["template_image"] = _extract_frame(
                    step["reference_timestamp"], step_coords, f"Step {i}"
                )
                screenspace.attach_capture_mask(
                    step, "template_image", "template_mask", step_coords
                )
            elif stype == "scene":
                step_refs = step.get("scene_references")
                if not step_refs:
                    raise ValueError(f"Step {i}: no scene_references to re-extract")
                step["reference_scenes"] = _ss_extract_scene_frames(
                    step_refs, frame_at, step_coords, context=f"Step {i}: "
                )


def _run_ss_rerun_task(args: argparse.Namespace) -> None:
    """Re-run a saved Screenspace task from the manifest by id.

    Re-extracts reference media from the source video (stripped on save), creates a
    fresh task run (new id, preserving the original), and persists the result. This is
    the only headless path for multitool tasks.
    """
    import screenspace

    task_id = args.ss_run_task
    manifest = screenspace.load_screenspace_manifest()
    saved = next((t for t in manifest.get("tasks", []) if t.get("id") == task_id), None)
    if saved is None:
        utils.error_print(
            f"Task {task_id!r} not found in the manifest.",
            ["List available tasks with --ss-list-tasks."],
        )
        sys.exit(1)

    task_type = saved.get("type", "")
    participant = saved.get("participant", "")
    region_name = saved.get("region", "")
    parameters = copy.deepcopy(saved.get("parameters", {}))

    video_paths = _ss_resolve_videos_for_participant(participant)
    if not video_paths:
        utils.error_print(
            f"No video found for participant {participant!r}.",
            ["Place the source video in the input directory before re-running."],
        )
        sys.exit(1)

    props = video.probe_video_properties(video_paths[0])
    dims: tuple[int, int] | None = None
    if props and props.get("width") and props.get("height"):
        dims = (int(props["width"]), int(props["height"]))

    # region_ref-aware resolution first; saved region_coords cover older tasks,
    # multitool steps, and missing dims.
    region_coords: dict[str, int] | None = None
    try:
        _, resolved_region = screenspace.resolve_region_request(
            region_name, saved.get("region_ref"), manifest
        )
        if dims is not None:
            region_coords = screenspace.denormalize_region(
                resolved_region, dims[0], dims[1]
            )
    except ValueError:
        pass  # fall through to saved coords below
    if region_coords is None and isinstance(saved.get("region_coords"), dict):
        region_coords = {
            k: int(saved["region_coords"][k])
            for k in ("x", "y", "w", "h")
            if k in saved["region_coords"]
        }
    if region_coords is None:
        utils.error_print(
            f"Region {region_name!r} for task {task_id!r} could not be resolved.",
            [
                "The named region is no longer in the manifest and no saved coords exist."
            ],
        )
        sys.exit(1)

    try:
        _ss_rehydrate_task_media(
            task_type,
            parameters,
            _ss_frame_extractor(video_paths),
            region_coords,
            manifest,
            dims,
        )
    except ValueError as exc:
        utils.error_print(f"Cannot re-run task {task_id!r}: {exc}")
        sys.exit(1)

    task = screenspace.create_task(
        task_type=task_type,
        participant=participant,
        source_video=Path(video_paths[0]).name,
        video_paths=video_paths,
        region_name=region_name,
        region_coords=region_coords,
        parameters=parameters,
    )

    _ss_run_and_persist_task(task, manifest)


# ---- Event-driven clip cutting (--ss-clips, --transcript-clips) ----


_SS_CLIPS_CELL_COL = 1  # synthetic cell column for --ss-clips artifacts
_TRANSCRIPT_CLIPS_CELL_COL = 2  # synthetic cell column for --transcript-clips


def _split_csv_set(value: str | None) -> set[str] | None:
    """Parse a comma-separated CLI value into a set of trimmed non-empty tokens.

    Returns None when the option was not supplied (so callers can distinguish
    "no filter" from "filter rejecting everything").
    """
    if value is None:
        return None
    items = {tok.strip() for tok in value.split(",")}
    items.discard("")
    return items if items else None


def _filter_screenspace_events(
    events: list[dict[str, Any]],
    *,
    detectors: set[str] | None,
    regions: set[str] | None,
    participants: set[str] | None,
    min_confidence: float | None,
    event_type_substr: str | None,
) -> list[dict[str, Any]]:
    """Apply CLI filters to Screenspace events; always drops excluded=True."""
    needle = event_type_substr.lower() if event_type_substr else None
    out: list[dict[str, Any]] = []
    for ev in events:
        if ev.get("excluded"):
            continue
        if detectors and ev.get("detector") not in detectors:
            continue
        if regions and ev.get("region") not in regions:
            continue
        if participants and ev.get("participant") not in participants:
            continue
        if min_confidence is not None:
            try:
                conf = float(ev.get("confidence", 0.0))
            except (TypeError, ValueError):
                conf = 0.0
            if conf < min_confidence:
                continue
        if needle is not None:
            label = str(ev.get("event_type", "")).lower()
            if needle not in label:
                continue
        out.append(ev)
    return out


def _filter_transcript_segments(
    manifest: dict[str, Any],
    *,
    participants: set[str] | None,
    mark_categories: set[str] | None,
    text_substr: str | None,
) -> list[tuple[str, dict[str, Any], list[dict[str, Any]]]]:
    """Return (participant_id, segment, attached_marks) tuples after filtering.

    When ``mark_categories`` is non-empty, only segments referenced by at least
    one matching mark are kept; the matching marks come along for clip metadata.
    Otherwise ``attached_marks`` is empty.
    """
    needle = text_substr.lower() if text_substr else None
    marks = manifest.get("marks") or []
    marks_by_segment: dict[str, list[dict[str, Any]]] = {}
    for mark in marks:
        if mark_categories and mark.get("category") not in mark_categories:
            continue
        seg_id = mark.get("segment_id")
        if not seg_id:
            continue
        marks_by_segment.setdefault(seg_id, []).append(mark)

    rows: list[tuple[str, dict[str, Any], list[dict[str, Any]]]] = []
    sources = manifest.get("source_transcripts") or {}
    for pid, entry in sources.items():
        if participants and pid not in participants:
            continue
        for idx, seg in enumerate(entry.get("segments") or []):
            seg_id = seg.get("id") or f"{pid}:{idx}"
            attached = marks_by_segment.get(seg_id, [])
            if mark_categories and not attached:
                continue
            if needle is not None and needle not in str(seg.get("text", "")).lower():
                continue
            rows.append((pid, seg, attached))
    return rows


def _cluster_groups(
    keyed_spans: list[tuple[Any, tuple[float, float], int]],
    *,
    gap: float,
    pad_pre: float,
    pad_post: float,
    max_duration: float,
) -> list[tuple[Any, tuple[float, float], list[int]]]:
    """Group items by ``key`` first, then run utils.cluster_spans within each group.

    Returns a list of (group_key, (cluster_start, cluster_end), member_indices),
    where ``member_indices`` index into the original ``keyed_spans`` list (not the
    per-group sublist) — so callers can look up the contributing items directly.
    """
    by_key: dict[Any, list[tuple[tuple[float, float], int]]] = {}
    for key, span, orig_idx in keyed_spans:
        by_key.setdefault(key, []).append((span, orig_idx))

    out: list[tuple[Any, tuple[float, float], list[int]]] = []
    for key, items in by_key.items():
        spans = [span for span, _ in items]
        local_to_orig = [orig for _, orig in items]
        clusters = utils.cluster_spans(
            spans,
            gap_seconds=gap,
            pad_pre=pad_pre,
            pad_post=pad_post,
            max_duration=max_duration,
        )
        for cs, ce, local_members in clusters:
            members = [local_to_orig[i] for i in local_members]
            out.append((key, (cs, ce), members))
    return out


def _build_clusters_from_ss_events(
    events: list[dict[str, Any]],
    *,
    gap: float,
    pad_pre: float,
    pad_post: float,
    max_duration: float,
) -> list[dict[str, Any]]:
    """Group SS events by (participant, source_video, detector) then cluster.

    Returns a list of cluster dicts: ``{participant, source_video, detector,
    region, start, end, regions, event_types, member_event_ids}``.
    """
    keyed: list[tuple[Any, tuple[float, float], int]] = []
    for idx, ev in enumerate(events):
        try:
            t_in = float(ev.get("time_in", 0.0))
            t_out = float(ev.get("time_out", t_in))
        except (TypeError, ValueError):
            continue
        t_out = max(t_out, t_in)
        key = (
            ev.get("participant", ""),
            ev.get("source_video", ""),
            ev.get("detector", ""),
        )
        keyed.append((key, (t_in, t_out), idx))

    clusters_raw = _cluster_groups(
        keyed,
        gap=gap,
        pad_pre=pad_pre,
        pad_post=pad_post,
        max_duration=max_duration,
    )

    out: list[dict[str, Any]] = []
    for key, (cs, ce), members in clusters_raw:
        participant, source_video, detector = key
        regions = sorted(
            {events[m].get("region", "") for m in members if events[m].get("region")}
        )
        event_types = sorted(
            {
                str(events[m].get("event_type", ""))
                for m in members
                if events[m].get("event_type")
            }
        )
        out.append(
            {
                "participant": participant,
                "source_video": source_video,
                "detector": detector,
                "region": regions[0] if len(regions) == 1 else "",
                "regions": regions,
                "event_types": event_types,
                "start": cs,
                "end": ce,
                "member_event_ids": [events[m].get("id", "") for m in members],
            }
        )
    out.sort(
        key=lambda c: (c["participant"], c["source_video"], c["detector"], c["start"])
    )
    return out


def _build_clusters_from_transcript_segments(
    rows: list[tuple[str, dict[str, Any], list[dict[str, Any]]]],
    transcripts_manifest: dict[str, Any],
    *,
    gap: float,
    pad_pre: float,
    pad_post: float,
    max_duration: float,
) -> list[dict[str, Any]]:
    """Group filtered (pid, segment, marks) by participant; cluster on segment spans.

    Returns a list of cluster dicts: ``{participant, source_video, start, end,
    text, mark_categories, member_segment_ids}``.
    """
    keyed: list[tuple[Any, tuple[float, float], int]] = []
    for idx, (pid, seg, _marks) in enumerate(rows):
        try:
            s = float(seg.get("start", 0.0))
            e = float(seg.get("end", s))
        except (TypeError, ValueError):
            continue
        e = max(e, s)
        keyed.append((pid, (s, e), idx))

    clusters_raw = _cluster_groups(
        keyed,
        gap=gap,
        pad_pre=pad_pre,
        pad_post=pad_post,
        max_duration=max_duration,
    )

    sources = transcripts_manifest.get("source_transcripts") or {}
    out: list[dict[str, Any]] = []
    for key, (cs, ce), members in clusters_raw:
        participant = key
        text = " ".join(
            str(rows[m][1].get("text", "")).strip() for m in members
        ).strip()
        mark_cats = sorted(
            {
                str(mark.get("category", ""))
                for m in members
                for mark in rows[m][2]
                if mark.get("category")
            }
        )
        seg_ids = [rows[m][1].get("id") for m in members if rows[m][1].get("id")]
        source_file = ""
        entry = sources.get(participant) or {}
        if isinstance(entry, dict):
            source_file = str(entry.get("source_file", "") or "")
        source_video = Path(source_file).name if source_file else ""
        out.append(
            {
                "participant": participant,
                "source_video": source_video,
                "start": cs,
                "end": ce,
                "text": text,
                "mark_categories": mark_cats,
                "member_segment_ids": seg_ids,
            }
        )
    out.sort(key=lambda c: (c["participant"], c["start"]))
    return out


def _truncate_for_filename(text: str, *, limit: int = 60) -> str:
    """Trim a text snippet for use in a clip description (filename-safe upstream)."""
    text = " ".join(text.split())
    if len(text) <= limit:
        return text
    return files.safe_truncate(text, limit).rstrip() + "…"


def _run_ss_clips(args: argparse.Namespace) -> None:
    """Cut clips from existing Screenspace events and append to the manifest."""
    import pipeline
    import screenspace

    manifest = screenspace.load_screenspace_manifest()
    raw_events = list(manifest.get("events") or [])
    if not raw_events:
        utils.warning_print(
            "No Screenspace events found.",
            [
                (
                    "Run --screenspace (UI) or --ss-task to generate events first, "
                    "or check your input/output directory."
                )
            ],
        )
        return

    filtered = _filter_screenspace_events(
        raw_events,
        detectors=_split_csv_set(args.ss_clips_detector),
        regions=_split_csv_set(args.ss_clips_region),
        participants=_split_csv_set(args.ss_clips_participant),
        min_confidence=args.ss_clips_min_confidence,
        event_type_substr=args.ss_clips_event_type,
    )
    if not filtered:
        utils.warning_print("No Screenspace events match the given filters.")
        return

    clusters = _build_clusters_from_ss_events(
        filtered,
        gap=args.cluster_gap,
        pad_pre=args.clip_pre,
        pad_post=args.clip_post,
        max_duration=args.max_clip_duration,
    )
    if not clusters:
        utils.warning_print("No clusters produced from filtered events.")
        return

    clips_list: list[ClipRecord] = []
    last_study = ""
    for idx, cluster in enumerate(clusters):
        source_video = cluster.get("source_video") or ""
        parsed = utils.parse_source_video_name(source_video) if source_video else None
        study = parsed[0] if parsed else ""
        participant = cluster.get("participant") or ""
        detector = cluster.get("detector") or ""
        region = cluster.get("region") or ""
        desc_parts = [detector]
        if region:
            desc_parts.append(region)
        desc = " ".join(p for p in desc_parts if p).strip() or "event"
        category = f"screenspace-{detector}" if detector else "screenspace"
        clips_list.extend(
            files.build_clip_records(
                participant=participant,
                source_filename=source_video,
                time_ranges=[(cluster["start"], cluster["end"])],
                description=desc,
                category=category,
                study=study,
                cell_col=_SS_CLIPS_CELL_COL,
                cell_row_base=idx,
            )
        )
        last_study = study or last_study

    count, artifacts = pipeline.process_clips(
        clips_list, output_format="clip", include_severity=False
    )
    if artifacts:
        viewer.save_manifest(
            artifacts,
            study=last_study,
            participant="",
            worksheet_title="",
            is_excel=False,
            mode="ss-clips",
        )
    utils.info_print(
        f"Generated {count} clip(s) from {len(filtered)} event(s) "
        f"in {len(clusters)} cluster(s)."
    )


def _run_transcript_clips(args: argparse.Namespace) -> None:
    """Cut clips from transcript segments/marks and append to the manifest."""
    import pipeline

    manifest = transcripts.load_transcripts_manifest()
    if not manifest.get("source_transcripts"):
        utils.warning_print(
            "No transcripts found.",
            [
                (
                    "Run --transcribe (with a clip mode) or --pre-transcribe / --transcripts "
                    "to generate transcripts first."
                )
            ],
        )
        return

    rows = _filter_transcript_segments(
        manifest,
        participants=_split_csv_set(args.transcript_clips_participant),
        mark_categories=_split_csv_set(args.transcript_clips_mark),
        text_substr=args.transcript_clips_text,
    )
    if not rows:
        utils.warning_print("No transcript segments match the given filters.")
        return

    clusters = _build_clusters_from_transcript_segments(
        rows,
        manifest,
        gap=args.cluster_gap,
        pad_pre=args.clip_pre,
        pad_post=args.clip_post,
        max_duration=args.max_clip_duration,
    )
    if not clusters:
        utils.warning_print("No clusters produced from filtered segments.")
        return

    mark_filter = _split_csv_set(args.transcript_clips_mark)
    clips_list: list[ClipRecord] = []
    last_study = ""
    for idx, cluster in enumerate(clusters):
        source_video = cluster.get("source_video") or ""
        parsed = utils.parse_source_video_name(source_video) if source_video else None
        study = parsed[0] if parsed else ""
        participant = cluster.get("participant") or ""
        text = cluster.get("text") or ""
        desc = _truncate_for_filename(text) if text else "transcript"
        if mark_filter and cluster.get("mark_categories"):
            primary = cluster["mark_categories"][0]
            category = f"mark-{primary}"
        else:
            category = "transcript"
        clips_list.extend(
            files.build_clip_records(
                participant=participant,
                source_filename=source_video,
                time_ranges=[(cluster["start"], cluster["end"])],
                description=desc,
                category=category,
                study=study,
                cell_col=_TRANSCRIPT_CLIPS_CELL_COL,
                cell_row_base=idx,
            )
        )
        last_study = study or last_study

    count, artifacts = pipeline.process_clips(
        clips_list, output_format="clip", include_severity=False
    )
    if artifacts:
        viewer.save_manifest(
            artifacts,
            study=last_study,
            participant="",
            worksheet_title="",
            is_excel=False,
            mode="transcript-clips",
        )
    utils.info_print(
        f"Generated {count} clip(s) from {len(rows)} segment(s) "
        f"in {len(clusters)} cluster(s)."
    )


def _post_marks_to_running_server(
    segment_ids: list[str],
    category: str,
    label: str | None,
) -> dict[str, Any] | None:
    """POST marks to a Transcripts server running on localhost.

    Returns the parsed response dict on success, or None when no server is
    reachable. Routing through the API keeps the server's in-memory manifest in
    sync — without this, a CLI-only disk write would be silently overwritten the
    next time the running server persists its (now-stale) state.
    """
    import json
    import urllib.error
    import urllib.request

    url = f"http://127.0.0.1:{config.SERVER_PORT}/transcripts/api/marks"
    payload = json.dumps(
        {"segment_ids": segment_ids, "category": category, "label": label}
    ).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=2.0) as resp:
            body = resp.read().decode("utf-8")
            data = json.loads(body)
            if isinstance(data, dict) and data.get("ok"):
                return data
            return None
    except (urllib.error.URLError, OSError, ValueError):
        return None


def _run_transcript_mark(args: argparse.Namespace) -> None:
    """Batch-mark transcript segments whose text contains a search term.

    Mirrors the Transcripts UI's "Mark all results" action: case-insensitive
    substring match against corrected segment text, upsert into the manifest's
    ``marks`` array (no duplicates per segment). When the Transcripts server is
    running locally, POST through its API so its in-memory state stays in sync;
    otherwise write the manifest directly.
    """
    term = (args.transcript_mark or "").strip()
    if not term:
        utils.error_print("--transcript-mark requires a non-empty TERM.")
        return

    category = (args.transcript_mark_category or "").strip()
    if category not in config.MARK_CATEGORIES:
        valid = ", ".join(sorted(config.MARK_CATEGORIES.keys()))
        utils.error_print(
            "--transcript-mark-category is required and must be a known category.",
            [f"Valid categories: {valid}"],
        )
        return

    label = args.transcript_mark_label
    participants_filter = _split_csv_set(args.transcript_mark_participant)

    manifest = transcripts.load_transcripts_manifest()
    source_transcripts = manifest["source_transcripts"]
    corrections = manifest.get("corrections", [])
    marks: list[dict[str, Any]] = list(manifest.get("marks") or [])

    if not source_transcripts:
        utils.warning_print(
            "No transcripts found.",
            [
                (
                    "Run --transcribe (with a clip mode) or --pre-transcribe / --transcripts "
                    "to generate transcripts first."
                )
            ],
        )
        return

    needle = term.lower()
    matching_seg_ids: list[tuple[str, str]] = []  # (participant, segment_id)
    for pid, entry in source_transcripts.items():
        if participants_filter and pid not in participants_filter:
            continue
        raw_segments = entry.get("segments") or []
        if not raw_segments:
            continue
        corrected = transcripts.apply_corrections(raw_segments, corrections)
        for raw, seg in zip(raw_segments, corrected):
            if needle in str(seg.get("text", "")).lower():
                seg_id = raw.get("id") or ""
                if seg_id:
                    matching_seg_ids.append((pid, seg_id))

    if not matching_seg_ids:
        utils.warning_print(f"No transcript segments contain {term!r}.")
        return

    seg_id_list = [sid for _, sid in matching_seg_ids]
    touched_participants = {pid for pid, _ in matching_seg_ids}
    total = len(seg_id_list)

    server_response = _post_marks_to_running_server(seg_id_list, category, label)
    if server_response is not None:
        utils.info_print(
            f"Marked {total} segment(s) across {len(touched_participants)} participant(s) "
            f"via running Transcripts server."
        )
        return

    existing_by_seg = {m["segment_id"]: m for m in marks if m.get("segment_id")}
    now = datetime.now(UTC).isoformat()
    created = 0
    updated = 0
    for sid in seg_id_list:
        existing = existing_by_seg.get(sid)
        if existing is not None:
            existing["category"] = category
            if label is not None:
                existing["label"] = label
            updated += 1
        else:
            new_mark = {
                "id": f"m_{uuid.uuid4().hex[:8]}",
                "segment_id": sid,
                "category": category,
                "label": label,
                "created": now,
            }
            marks.append(new_mark)
            existing_by_seg[sid] = new_mark
            created += 1

    transcripts.save_transcripts_manifest(source_transcripts, corrections, marks=marks)
    utils.info_print(
        f"Marked {total} segment(s) across {len(touched_participants)} participant(s) "
        f"(created {created}, updated {updated})."
    )


# ---- Thinking-agent CLI ----


def _select_transcript_targets(
    requested: list[str] | None,
    source_transcripts: dict[str, Any],
) -> list[str]:
    """Resolve a requested participant list against transcripted participants.

    ``requested`` is the value of ``args.summarize``, ``args.citations``, or
    ``args.friction`` — None should never reach here, but ``[]`` means "all
    transcripted".
    Unknown IDs print a warning and are dropped.
    """
    if not requested:
        return list(source_transcripts.keys())
    targets: list[str] = []
    for raw_id in requested:
        pid = utils.normalize_participant_id(raw_id)
        if pid in source_transcripts:
            targets.append(pid)
        else:
            utils.warning_print(
                f"Participant {raw_id!r} has no transcript; skipping. "
                f"Available: {', '.join(sorted(source_transcripts.keys()))}"
            )
    return targets


class _AgentPass(NamedTuple):
    """CLI wording for one thinking-agent pass; the registry supplies the rest."""

    key: str  # thinking_agents.AGENTS key
    args_attr: str  # argparse attribute holding the participant list
    no_targets: str
    start: str  # "{pid}" template
    not_produced: str  # "{pid}" template
    tally: str  # "{done}" / "{skipped}" template
    report: Callable[[str, Any], None]  # per-participant success line


# Flag that produces each agent's manifest field, for the prerequisite hint.
_AGENT_FLAGS = {
    "summary": "--summarize",
    "citations": "--citations",
    "friction": "--friction",
}


def _report_summary(pid: str, result: Any) -> None:
    utils.info_print(f"  {pid}: summary stored ({len(result)} chars).")


def _report_citations(pid: str, result: Any) -> None:
    total_refs = sum(len(c.get("refs") or []) for c in result)
    utils.info_print(f"  {pid}: {len(result)} claim(s), {total_refs} ref(s) stored.")


def _report_friction(pid: str, result: Any) -> None:
    moment_count = len(result.get("moments") or [])
    if result.get("llm_ok"):
        utils.info_print(f"  {pid}: {moment_count} friction moment(s) stored.")
    else:
        utils.warning_print(
            f"  {pid}: programmatic scores stored, but LLM moment detection "
            "failed (AI server unavailable?)."
        )


_SUMMARY_PASS = _AgentPass(
    key="summary",
    args_attr="summarize",
    no_targets="No transcribed participants to summarize.",
    start="Summarizing {pid}...",
    not_produced="{pid}: summary not produced (transcript too short or AI server unavailable).",
    tally="Summary complete: {done} summarized, {skipped} skipped.",
    report=_report_summary,
)
_CITATIONS_PASS = _AgentPass(
    key="citations",
    args_attr="citations",
    no_targets="No transcribed participants for citation generation.",
    start="Finding citations for {pid}...",
    not_produced="{pid}: no citations produced (AI server unavailable or empty summary).",
    tally="Citations complete: {done} processed, {skipped} skipped.",
    report=_report_citations,
)
_FRICTION_PASS = _AgentPass(
    key="friction",
    args_attr="friction",
    no_targets="No transcribed participants for friction analysis.",
    start="Scoring friction for {pid}...",
    not_produced="{pid}: friction not produced (missing transcript or summary).",
    tally="Friction complete: {done} processed, {skipped} skipped.",
    report=_report_friction,
)


def _run_thinking_agent(args: argparse.Namespace, spec: _AgentPass) -> None:
    """Run one registry agent over the selected transcripts, storing its field."""
    import thinking_agents

    agent = thinking_agents.get_agent(spec.key)
    assert agent is not None  # registered in thinking_agents.AGENTS
    field = agent["manifest_field"]

    manifest = transcripts.load_transcripts_manifest()
    source_transcripts = manifest["source_transcripts"]
    corrections = manifest["corrections"]
    marks = manifest.get("marks")

    targets = _select_transcript_targets(
        getattr(args, spec.args_attr), source_transcripts
    )
    if not targets:
        utils.error_print(spec.no_targets)
        return

    done = 0
    skipped = 0
    for pid in targets:
        entry = source_transcripts.get(pid)
        if not entry:
            utils.warning_print(f"{pid}: no transcript entry; skipping.")
            skipped += 1
            continue
        missing = next((dep for dep in agent["depends_on"] if not entry.get(dep)), None)
        if missing:
            utils.warning_print(
                f"{pid}: no {missing} yet; run {_AGENT_FLAGS[missing]} first."
            )
            skipped += 1
            continue
        if entry.get(field) and not args.no_input:
            utils.info_print(
                f"{pid}: {field} already present; skip (--no-input to overwrite)."
            )
            skipped += 1
            continue
        utils.info_print(spec.start.format(pid=pid))
        result = agent["run"](entry, None)
        if not result:
            utils.warning_print(spec.not_produced.format(pid=pid))
            skipped += 1
            continue
        entry[field] = result
        transcripts.save_transcripts_manifest(source_transcripts, corrections, marks)
        spec.report(pid, result)
        done += 1

    utils.info_print(spec.tally.format(done=done, skipped=skipped))


def _run_summarize(args: argparse.Namespace) -> None:
    _run_thinking_agent(args, _SUMMARY_PASS)


def _run_citations(args: argparse.Namespace) -> None:
    _run_thinking_agent(args, _CITATIONS_PASS)


def _run_friction_agent(args: argparse.Namespace) -> None:
    _run_thinking_agent(args, _FRICTION_PASS)


def _run_timeline_viewer_mode(worksheet: Any, args: Any) -> None:
    """Export all clips via batch mode and generate a per-participant timeline viewer."""
    clips_list = spreadsheet.generate_list(worksheet, "batch", skip_prompts=True)
    outputs_generated, artifacts = app.process_clips(clips_list, output_format="clip")

    if not config.REENCODING:
        app._print_reencoding_warning(utils.verbose_print)
    app._print_completion_message(outputs_generated, "clip", is_reel=False)

    if not artifacts:
        utils.warning_print("No artifacts were generated; skipping timeline viewer.")
        return

    study = artifacts[0].get("study", "")
    ss_events = viewer.load_screenspace_events_for_viewer()
    data = viewer.finalize_timeline_data(
        artifacts,
        study=study,
        worksheet_title=getattr(worksheet, "title", ""),
        is_excel=app._is_excel_worksheet(worksheet),
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

    if getattr(args, "manifest", False):
        manifest_path = viewer.save_manifest(
            artifacts,
            study=study,
            worksheet_title=getattr(worksheet, "title", ""),
            is_excel=app._is_excel_worksheet(worksheet),
            mode="timeline-viewer",
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
        outputs_generated, reel_records = app.process_reel(
            clips_list,
            output_file=reel_output_file,
        )
    else:
        outputs_generated, artifacts = app.process_clips(
            clips_list,
            output_format=output_format,
            include_severity=bool(args.severity),
        )

    if not config.REENCODING:
        app._print_reencoding_warning(utils.verbose_print)
    app._print_completion_message(
        outputs_generated,
        output_format,
        is_reel=is_reel,
    )

    ws_title = getattr(worksheet, "title", "")
    is_excel = app._is_excel_worksheet(worksheet)
    effective_mode = output_format if output_format != "clip" else "batch"

    wants_viewer_or_manifest = getattr(args, "viewer", False) or getattr(
        args, "manifest", False
    )
    if wants_viewer_or_manifest and (artifacts or reel_records):
        primary = artifacts[0] if artifacts else reel_records[0]
        study = primary.get("study", "")
        participant = primary.get("participant", "")

        if getattr(args, "viewer", False):
            ss_events = viewer.load_screenspace_events_for_viewer()
            data = viewer.finalize_timeline_data(
                artifacts,
                reels=reel_records or None,
                study=study,
                participant=participant,
                worksheet_title=ws_title,
                is_excel=is_excel,
                mode="reel" if is_reel and not artifacts else effective_mode,
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
                mode="reel" if is_reel and not artifacts else effective_mode,
            )
            if manifest_path:
                utils.info_print(f"Manifest updated: {manifest_path}")


# ---- Main entry point ----


# Clip-selection flags, shared by _generate_cli_clips, _BASE_SELECTOR_ATTRS, and
# main()'s cli_mode detection.
_SELECTION_ATTRS = (
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
)

_BASE_SELECTOR_ATTRS = _SELECTION_ATTRS + ("screen", "gif", "viewer", "regenerate")
# Headless CLI modes additionally conflict with --highlights.
_CLI_SELECTOR_ATTRS = _BASE_SELECTOR_ATTRS + ("highlights",)


class _ModeSpec(NamedTuple):
    """Declarative description of one exclusive mode for conflict validation.

    Every standalone mode is mutually exclusive with every other standalone
    mode; that all-pairs rule is derived in _validate_mode_conflicts, not
    hand-listed per spec.
    """

    key: str  # attribute on args (also returned in result dict)
    error: str
    hint: str
    is_value: bool = (
        False  # flag carries a value: active when not None, else when truthy
    )
    selector_attrs: tuple[str, ...] = _CLI_SELECTOR_ATTRS
    implies_cli_mode: bool = True  # headless CLI mode (vs. web/standalone dispatch)
    dispatch: Callable[[Any], None] | None = None  # spreadsheet-free runner, if any

    def active(self, args: Any) -> bool:
        value = getattr(args, self.key, None)
        return value is not None if self.is_value else bool(value)


_EXCLUSIVE_MODES: tuple[_ModeSpec, ...] = (
    _ModeSpec(
        key="timeline_viewer",
        error="--timeline-viewer cannot be combined with mode, format, or --viewer/--regenerate flags.",
        hint="Only -s (spreadsheet) and -v (verbose) may be used alongside --timeline-viewer.",
        selector_attrs=_BASE_SELECTOR_ATTRS,
    ),
    _ModeSpec(
        key="studio",
        error="--studio cannot be combined with mode, format, or --viewer/--regenerate flags.",
        hint="Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --studio.",
        selector_attrs=_BASE_SELECTOR_ATTRS,
        implies_cli_mode=False,
    ),
    _ModeSpec(
        key="screenspace",
        error="--screenspace cannot be combined with mode, format, or --viewer/--regenerate/--studio flags.",
        hint="Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --screenspace.",
        selector_attrs=_BASE_SELECTOR_ATTRS,
        implies_cli_mode=False,
    ),
    _ModeSpec(
        key="transcripts",
        error="--transcripts cannot be combined with mode, format, or --viewer/--regenerate/--studio/--screenspace flags.",
        hint="Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --transcripts.",
        selector_attrs=_BASE_SELECTOR_ATTRS,
        implies_cli_mode=False,
    ),
    _ModeSpec(
        key="workflows",
        error="--workflows cannot be combined with mode, format, or other web/CLI mode flags.",
        hint="Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --workflows.",
        selector_attrs=_BASE_SELECTOR_ATTRS,
        implies_cli_mode=False,
    ),
    _ModeSpec(
        key="composer",
        error="--composer cannot be combined with mode, format, or other web/CLI mode flags.",
        hint="Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --composer.",
        selector_attrs=_BASE_SELECTOR_ATTRS,
        implies_cli_mode=False,
    ),
    _ModeSpec(
        key="overview",
        error="--overview cannot be combined with mode, format, or other web/CLI mode flags.",
        hint="Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --overview.",
        selector_attrs=_BASE_SELECTOR_ATTRS,
        implies_cli_mode=False,
    ),
    _ModeSpec(
        key="gallery",
        # `gallery` carries an optional VIDEO arg, so presence means active.
        error="--gallery cannot be combined with selection modes, --viewer, --regenerate, --studio, or --timeline-viewer.",
        hint="Only --gif, --interval, --bundle, -i/-o (directories), and -v (verbose) may be used alongside --gallery.",
        # Gallery permits --screen and --gif as output-format toggles.
        selector_attrs=tuple(
            a for a in _BASE_SELECTOR_ATTRS if a not in ("screen", "gif")
        ),
        is_value=True,
        implies_cli_mode=False,
    ),
    _ModeSpec(
        key="pre_transcribe",
        error="--pre-transcribe cannot be combined with mode, format, or --studio/--screenspace/--transcripts flags.",
        hint="Only -s (spreadsheet), -i/-o (directories), and -v (verbose) may be used alongside --pre-transcribe.",
        # pre-transcribe additionally conflicts with --highlights.
        is_value=True,
    ),
    _ModeSpec(
        key="export",
        error="--export cannot be combined with mode, format, or other standalone flags.",
        hint="Use --export with -i/-o (directories) and -v (verbose) only.",
    ),
    _ModeSpec(
        key="ss_task",
        error="--ss-task cannot be combined with mode, format, or other standalone flags.",
        hint="Use --ss-task with -i/-o (directories) and -v (verbose) only.",
        is_value=True,
        dispatch=_run_ss_task,
    ),
    _ModeSpec(
        key="ss_run_task",
        error="--ss-run-task cannot be combined with mode, format, or other standalone flags.",
        hint="Use --ss-run-task with -i/-o (directories) and -v (verbose) only.",
        is_value=True,
        dispatch=_run_ss_rerun_task,
    ),
    _ModeSpec(
        key="ss_list_regions",
        error="--ss-list-regions cannot be combined with other modes.",
        hint="Use --ss-list-regions on its own (with -i/-o for directories).",
        dispatch=_run_ss_list_regions,
    ),
    _ModeSpec(
        key="ss_list_stashes",
        error="--ss-list-stashes cannot be combined with other modes.",
        hint="Use --ss-list-stashes on its own (with -i/-o for directories).",
        dispatch=_run_ss_list_stashes,
    ),
    _ModeSpec(
        key="ss_list_tasks",
        error="--ss-list-tasks cannot be combined with other modes.",
        hint="Use --ss-list-tasks on its own (with -i/-o for directories).",
        is_value=True,
        dispatch=_run_ss_list_tasks,
    ),
    _ModeSpec(
        key="summarize",
        error="--summarize cannot be combined with mode, format, or other standalone flags.",
        hint="Use --summarize with -i/-o (directories), -v (verbose), and --llm-model.",
        is_value=True,
        dispatch=_run_summarize,
    ),
    _ModeSpec(
        key="citations",
        error="--citations cannot be combined with mode, format, or other standalone flags.",
        hint="Use --citations with -i/-o (directories), -v (verbose), and --llm-model.",
        is_value=True,
        dispatch=_run_citations,
    ),
    _ModeSpec(
        key="friction",
        error="--friction cannot be combined with mode, format, or other standalone flags.",
        hint="Use --friction with -i/-o (directories), -v (verbose), and --llm-model.",
        is_value=True,
        dispatch=_run_friction_agent,
    ),
    _ModeSpec(
        key="ss_clips",
        error="--ss-clips cannot be combined with mode, format, or other standalone flags.",
        hint=(
            "Use --ss-clips with -i/-o (directories), -v (verbose), "
            "--cluster-gap / --clip-pre / --clip-post / --max-clip-duration, "
            "and --ss-clips-* filters."
        ),
        dispatch=_run_ss_clips,
    ),
    _ModeSpec(
        key="transcript_clips",
        error="--transcript-clips cannot be combined with mode, format, or other standalone flags.",
        hint=(
            "Use --transcript-clips with -i/-o (directories), -v (verbose), "
            "--cluster-gap / --clip-pre / --clip-post / --max-clip-duration, "
            "and --transcript-clips-* filters."
        ),
        dispatch=_run_transcript_clips,
    ),
    _ModeSpec(
        key="transcript_mark",
        error="--transcript-mark cannot be combined with mode, format, or other standalone flags.",
        hint=(
            "Use --transcript-mark with -i/-o (directories), -v (verbose), "
            "--transcript-mark-category, and optional "
            "--transcript-mark-participant / --transcript-mark-label."
        ),
        is_value=True,
        dispatch=_run_transcript_mark,
    ),
    _ModeSpec(
        key="regenerate",
        error="--regenerate cannot be combined with mode, format, or other standalone flags.",
        hint="Use --regenerate with -i/-o (directories) and -v (verbose) only.",
        # Exclude self from selector_attrs to avoid a false self-conflict.
        selector_attrs=tuple(a for a in _BASE_SELECTOR_ATTRS if a != "regenerate")
        + ("highlights",),
        implies_cli_mode=False,
    ),
)


def _validate_mode_conflicts(args: Any) -> dict[str, Any]:
    """Validate mutually exclusive mode flags and exit on conflict.

    Returns a dict keyed by mode name. Boolean for each mode plus
    ``"gallery_arg"`` (the optional VIDEO argument from --gallery).
    """
    active = {spec.key: spec.active(args) for spec in _EXCLUSIVE_MODES}
    for spec in _EXCLUSIVE_MODES:
        if not active[spec.key]:
            continue
        conflicts = [getattr(args, attr, None) for attr in spec.selector_attrs]
        # Every standalone mode is mutually exclusive with every other one.
        conflicts.extend(v for k, v in active.items() if k != spec.key)
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

    if getattr(args, "profile", False):
        profiling.enable()
    if getattr(args, "profile_deep", None):
        config.PROFILE_DEEP = args.profile_deep
        profiling.enable()

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
    if getattr(args, "no_whisper_vad", False):
        config.TRANSCRIBE_VAD_FILTER = False
    if getattr(args, "speakers", False):
        config.TRANSCRIBE_SPEAKERS = True
    if getattr(args, "whisper_hallucination_silence", None) is not None:
        config.TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD = (
            args.whisper_hallucination_silence
        )
    if getattr(args, "llm_model", None):
        config.LLM_SUMMARY_MODEL = args.llm_model

    return parse_cli_mode_args(args)


def _maybe_apply_persisted_dirs(args: Any) -> None:
    """Apply last-used input/output dirs from start_settings when CLI didn't set them.

    Only used by the web-frontend dispatch path. Skips silently if persistence
    is disabled or the saved paths no longer exist.
    """
    import start_settings

    settings = start_settings.load_start_settings()
    if not settings.get("persist_enabled", True):
        return
    if getattr(args, "input", None) is None:
        last_input = settings.get("last_input") or ""
        if last_input and Path(last_input).is_dir():
            config.INPUT_DIR = last_input
    if getattr(args, "output", None) is None:
        last_output = settings.get("last_output") or ""
        if last_output and Path(last_output).is_dir():
            config.OUTPUT_DIR = last_output


# Ordered so the first match wins, matching the mutually-exclusive frontend
# flags validated by _validate_mode_conflicts.
_WEB_MODES = (
    "studio",
    "screenspace",
    "transcripts",
    "workflows",
    "composer",
    "overview",
)


def _resolve_web_mode(args: Any) -> str | None:
    """Return the web frontend requested by *args*, or None if none was."""
    for mode in _WEB_MODES:
        if getattr(args, mode, False):
            return mode
    return None


def _use_desktop_window(args: Any) -> bool:
    """Decide whether a web frontend opens in a native window or the browser.

    ``--browser`` always wins. Otherwise a window is used when asked for with
    ``--desktop``, or when a frozen bundle was launched with no arguments at all
    — a Finder/Explorer double-click. Every other invocation, including every
    explicit CLI run of a frozen binary, keeps the browser behaviour.
    """
    if getattr(args, "browser", False):
        return False
    if getattr(args, "desktop", False):
        return True
    return bool(getattr(sys, "frozen", False)) and not sys.argv[1:]


def _make_worksheet_factory(args: Any) -> Any:
    """Deferred `-s` worksheet opener for window-first desktop launches.

    Returns ``factory(client) -> (worksheet, notice)``, run on the server's
    boot-build thread. That thread operates under ``NO_INPUT_MODE`` with no
    console, so everything the console path handles interactively degrades to
    a ``notice`` dict (``message`` + ``source_type``) — the boot continues
    sheetless and the Start overlay shows the message as the recovery surface,
    landed on the failed source's tab.
    """
    sheet_arg = getattr(args, "spreadsheet", None)
    source_type = "excel" if _is_excel_spreadsheet_arg(sheet_arg) else "google"

    def factory(client: Any) -> tuple[Any, dict[str, str] | None]:
        if source_type == "google" and client is None:
            # No cached token; a background thread must never start browser OAuth.
            message = (
                f"Google sign-in is needed to open '{sheet_arg}' — connect "
                "below, then pick it again."
            )
            return None, {"message": message, "source_type": source_type}
        try:
            return select_worksheet(client, args, cli_mode=False), None
        except BaseException as exc:
            # BaseException on purpose: select_worksheet's sys.exit must not kill the
            # boot build.
            utils.warning_print(f"Could not open spreadsheet '{sheet_arg}': {exc}")
            message = f"Could not open spreadsheet '{sheet_arg}' — pick a source below."
            return None, {"message": message, "source_type": source_type}

    return factory


def _launch_web_frontend(
    args: Any,
    default_page: str,
    worksheet: Any = None,
    gspread_client: Any = None,
    gspread_client_factory: Any = None,
    worksheet_factory: Any = None,
) -> None:
    """Serve *default_page*, in a desktop window or the default browser.

    Single funnel for all seven launch sites so the window-vs-browser decision
    lives in exactly one place. *gspread_client* is a client the caller already
    authenticated (console `-s` path); *gspread_client_factory* defers that work
    to the server's boot-build thread instead, and *worksheet_factory* defers
    the `-s` worksheet open the same way.
    """
    if _use_desktop_window(args):
        import desktop

        desktop.launch(
            worksheet=worksheet,
            default_page=default_page,
            gspread_client=gspread_client,
            gspread_client_factory=gspread_client_factory,
            worksheet_factory=worksheet_factory,
        )
        return

    import server

    server.start_combined_server(
        worksheet=worksheet,
        default_page=default_page,
        gspread_client=gspread_client,
        gspread_client_factory=gspread_client_factory,
        worksheet_factory=worksheet_factory,
    )


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
        existing_artifacts, existing_reels = viewer.load_manifest_both()
        if not existing_artifacts and not existing_reels:
            utils.error_print(
                "No manifest found or manifest is empty.",
                [
                    f"Run a clip generation mode with --manifest first to create {config.MANIFEST_FILENAME}."
                ],
            )
            sys.exit(1)
        primary = existing_artifacts[0] if existing_artifacts else existing_reels[0]
        study = primary.get("study", "")
        participant = primary.get("participant", "")
        ss_events = viewer.load_screenspace_events_for_viewer()
        data = viewer.finalize_timeline_data(
            existing_artifacts,
            reels=existing_reels or None,
            study=study,
            participant=participant,
            mode="manifest",
            screenspace_events=ss_events or None,
        )
        viewer_path = viewer.generate_timeline_viewer(data)
        if viewer_path:
            utils.info_print(f"Timeline viewer created from manifest: {viewer_path}")
        return True

    # Standalone analysis-data export
    if getattr(args, "export", False):
        import data_export

        sys.exit(data_export.run_cli_export())

    # Spreadsheet-free CLI runners declare themselves on _EXCLUSIVE_MODES.
    for spec in _EXCLUSIVE_MODES:
        if spec.dispatch is not None and spec.active(args):
            spec.dispatch(args)
            return True

    # Standalone web frontend, no spreadsheet; the Start overlay picks one in-app.
    web_mode = _resolve_web_mode(args)
    if web_mode is not None and not args.spreadsheet:
        profiling.mark("startup.web_dispatch")
        _maybe_apply_persisted_dirs(args)
        # Reuse the cached Google token; the factory defers gspread's slow import to
        # the boot thread.
        _launch_web_frontend(
            args,
            web_mode,
            worksheet=None,
            gspread_client_factory=_try_silent_google_auth,
        )
        return True

    # Standalone gallery
    if gallery_arg is not None:
        _run_gallery_cli(args)
        return True

    # Standalone regenerate from manifest
    if getattr(args, "regenerate", False) and not cli_mode:
        existing_artifacts, existing_reels = viewer.load_manifest_both()
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
        regenerated = app.regenerate_from_manifest(
            existing_artifacts, reels=existing_reels
        )
        utils.info_print(f"Regenerated {regenerated} of {total} item(s).")
        return True

    return False


def main() -> None:
    """Main entry point for clipgen."""
    setup_encoding()

    args = parse_arguments()
    profiling.mark("startup.args_parsed")

    # Frozen bundle with no args (Finder/Explorer double-click) lands in Studio.
    if getattr(sys, "frozen", False) and not sys.argv[1:]:
        args.studio = True

    # --desktop / --browser alone name a surface, not a frontend; default to Studio.
    surface_only = getattr(args, "desktop", False) or getattr(args, "browser", False)
    if surface_only and _resolve_web_mode(args) is None:
        args.studio = True

    utils.NO_INPUT_MODE = bool(getattr(args, "no_input", False))

    # A windowed launch has no console; set this before anything can abort startup.
    utils.GUI_LAUNCH = _use_desktop_window(args) and _resolve_web_mode(args) is not None
    # Bundled ffmpeg/ffprobe must beat system copies; Homebrew (appended next) is
    # the fallback. Run before shutil.which.
    utils.prepend_bundled_bin_to_path()
    # Finder gives GUI processes a bare PATH without Homebrew, hiding ffmpeg.
    utils.augment_path_for_gui_launch()

    if config.DEBUGGING:
        config.debug_ic(args)

    # Validate mutually exclusive mode flags
    modes = _validate_mode_conflicts(args)
    timeline_viewer = modes["timeline_viewer"]
    gallery_arg = modes["gallery_arg"]
    pre_transcribe_mode = modes["pre_transcribe"]

    # Headless modes set cli_mode via _ModeSpec.implies_cli_mode; web frontends,
    # gallery, and regenerate do not.
    cli_mode = (
        any(getattr(args, a, None) for a in _SELECTION_ATTRS)
        or args.highlights
        or args.screen
        or args.gif
        or any(modes[s.key] for s in _EXCLUSIVE_MODES if s.implies_cli_mode)
    )

    cli_mode_args = _apply_config_overrides(args, cli_mode)

    # Change working directory to runtime location (script/executable)
    os.chdir(get_runtime_working_dir())
    utils.standard_print(
        "-------------------------------------------------------------------------------"
    )
    utils.standard_print(
        f"Welcome to clipgen v{utils.get_version()}\nWorking directory: {os.getcwd()}\nPlace video files and the credentials.json file in this directory."
    )
    utils.debug_print(
        "Debug mode is ON. Several limitations apply and more things will be printed."
    )

    # Sanity-check input/output directories before proceeding
    utils.validate_runtime_directories()

    if getattr(args, "settings", False):
        if args.no_input:
            utils.error_print(
                "--settings requires interactive input and cannot be combined with --no-input."
            )
            sys.exit(1)
        # Reopen the grid after each change; empty input exits.
        while utils.set_program_settings():
            pass

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
        utils.fatal_startup_error(
            "WebP output is configured but ffmpeg lacks libwebp support.",
            [
                f"Affected config: {', '.join(webp_formats)}",
                "Install an ffmpeg build with libwebp, or change the format(s) back to .png/.jpg/.gif in config.py.",
            ],
        )
        sys.exit(1)

    if config.GIF_FORMAT.lower() == ".webm" and not video.check_vp9_support():
        utils.fatal_startup_error(
            "WebM output is configured but ffmpeg lacks libvpx-vp9 support.",
            [
                "Affected config: GIF_FORMAT",
                "Install an ffmpeg build with libvpx, or change GIF_FORMAT back to .gif/.webp in config.py.",
            ],
        )
        sys.exit(1)

    if config.TITLECARDS_ENABLED and not video.check_drawtext_support():
        utils.error_print(
            "Titlecards are enabled but ffmpeg lacks the drawtext filter.",
            [
                "drawtext requires libfreetype (often missing from Homebrew's default ffmpeg 8.x build).",
                "Install an ffmpeg build with libfreetype to restore titlecards.",
                "Disabling titlecards for this run.",
            ],
        )
        config.TITLECARDS_ENABLED = False

    profiling.mark("startup.ffmpeg_checks")
    if _dispatch_standalone_mode(args, cli_mode, gallery_arg):
        sys.exit(0)

    # Desktop -s launch has no console: defer auth and worksheet selection to the
    # boot thread.
    web_mode = _resolve_web_mode(args)
    if web_mode is not None and _use_desktop_window(args):
        profiling.mark("startup.web_dispatch")
        _launch_web_frontend(
            args,
            web_mode,
            worksheet=None,
            gspread_client_factory=_try_silent_google_auth,
            worksheet_factory=_make_worksheet_factory(args),
        )
        sys.exit(0)

    # Google auth once per run; skipped for Excel. select_worksheet fetches the
    # Drive listing lazily.
    gspread_client = None
    if not _is_excel_spreadsheet_arg(getattr(args, "spreadsheet", None)):
        gspread_client = authenticate_google()
        if gspread_client is None:  # noqa: SIM102 - the comment below belongs to the inner branch
            # Auth failed and no interactive recovery: try a sole local .xlsx, else
            # exit.
            if cli_mode or getattr(args, "spreadsheet", None):
                fallback = None
                if cli_mode and not getattr(args, "spreadsheet", None):
                    fallback = _single_xlsx_fallback_path(
                        "Google authentication failed"
                    )
                if fallback:
                    # select_worksheet routes .xlsx args through excel_io
                    # without touching the (absent) gspread client.
                    args.spreadsheet = fallback
                else:
                    utils.fatal_startup_error(
                        "Google authentication failed.",
                        [
                            "Use -s path/to/file.xlsx to work with a local Excel file instead.",
                        ],
                    )
                    sys.exit(1)
            # Interactive mode: the loop below prompts for an Excel file instead.

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
                worksheet = select_worksheet(gspread_client, args, cli_mode)

            # Filename overrides pick the source file per participant; web launches
            # seed them in server._seed_filename_overrides.
            meta = files.derive_sheet_meta(worksheet)
            if meta is not None:
                import start_settings

                config.FILENAME_OVERRIDES = start_settings.filename_overrides(
                    meta["type"], meta["id_or_path"], meta.get("worksheet", "")
                )

            web_mode = _resolve_web_mode(args)
            if web_mode is not None:
                _launch_web_frontend(
                    args,
                    web_mode,
                    worksheet=worksheet,
                    gspread_client=gspread_client,
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
                app.run_interactive_mode(worksheet, gspread_client=gspread_client)
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
