# -*- coding: utf-8 -*-
"""Utility functions for clipgen."""

import difflib
import functools
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, TypedDict, TypeVar

from icecream import ic

import config


# ---- Non-interactive mode flag ----
# Set to True from cli.py when --no-input is passed. Targeted guards in
# clipgen.py / excel_io.py / pipeline.py / video.py / utils.suggest_close_match
# read this flag to fail fast (or skip) instead of blocking on stdin.
NO_INPUT_MODE: bool = False


# ---- Shared type definitions ----


class ClipRecord(TypedDict, total=False):
    """Clip record built by spreadsheet layer, enriched by prepare_clip.

    Always present after _make_clip_record: cell, desc, study, participant, category.
    Added by prepare_clip: times, cell_annotations, segment_annotations.
    Optionally set before prepare_clip: selected_segment_indexes, timestamp_baseline,
    source_filename.
    """

    cell: Any  # gspread.Cell or ExcelSheetAdapter equivalent
    desc: str
    study: str
    participant: str
    category: str
    severity: str
    source_filename: str
    timestamp_baseline: str
    times: list[tuple[str, str]]
    cell_annotations: list[str]
    segment_annotations: dict[str, list[int]]
    selected_segment_indexes: list[int]


class ReelInput(TypedDict):
    """Parsed reel selector input from parse_reel_input."""

    batch: bool
    keyword: bool
    chronologic: bool
    severity: bool
    highlights: bool
    lines: list[int]
    ranges: list[tuple[int, int]]
    categories: list[str]
    cells: list[tuple[str, int]]
    participants: list[str]


class BrowseRow(TypedDict, total=False):
    """Row data for browse mode display."""

    row_num: int
    category: str
    description: str
    timestamps: dict[str, str]


# ---- Rich library integration with graceful fallback ----
try:
    from rich.console import Console
    from rich.panel import Panel
    from rich.progress import (
        Progress,
        BarColumn,
        TextColumn,
        TimeElapsedColumn,
        MofNCompleteColumn,
        SpinnerColumn,
    )
    from rich.table import Table
    from rich.text import Text
    from rich.theme import Theme

    RICH_AVAILABLE = True
except ImportError:
    RICH_AVAILABLE = False
    Progress = None  # type: ignore

# ---- Custom theme for clipgen - only create if Rich is available ----
_CLIPGEN_THEME = (
    Theme(
        {
            "error": "bold red",
            "error.prefix": "bold red",
            "error.detail": "dim",
            "warning": "bold yellow",
            "warning.prefix": "bold yellow",
            "warning.detail": "dim",
            "success": "bold green",
            "success.prefix": "bold green",
            "info": "cyan",
            "verbose": "dim",
            "debug": "magenta",
            "mode.spreadsheet": "bold blue",
            "mode.selection": "bold cyan",
            "mode.reel": "bold green",
            "mode.reellate": "bold magenta",
            "mode.format": "bold yellow",
            "mode.chronologic": "bold green",
            "mode.browse": "bold white",
            "mode.batch": "bold purple",
            "mode.range": "bold plum3",
            "mode.category": "bold cyan",
            "mode.line": "bold orange3",
            "mode.cell": "bold light_pink3",
            "mode.participant": "bold cyan",
            "mode.keyword": "bold cyan",
            "mode.severity": "bold #FF8C00",
            "mode.regenerate": "bold green",
            "severity.critical": "bold #8B0000",
            "severity.high": "bold red",
            "severity.medium": "bold #FF8C00",
            "severity.low": "bold yellow",
            "severity.na": "bold cyan",
            "severity.positive": "bold #90EE90",
            "severity.very_positive": "bold green",
        }
    )
    if RICH_AVAILABLE
    else None
)


# ---- Interactive decorations ----
console = Console(theme=_CLIPGEN_THEME, highlight=False) if RICH_AVAILABLE else None


# ---- Rich output helpers (defined before print functions that use them) ----


def _use_rich() -> bool:
    """Check if Rich output should be used."""
    return (
        RICH_AVAILABLE and console is not None and getattr(config, "RICH_COLORS", True)
    )


def _use_panels() -> bool:
    """Check if Rich panels should be used for errors/warnings/success."""
    return getattr(config, "RICH_PANELS", True)


def use_progress() -> bool:
    """Check if Rich progress bars should be used."""
    return (
        RICH_AVAILABLE
        and console is not None
        and getattr(config, "RICH_PROGRESS", True)
    )


# ---- Print functions ----


def debug_print(message: str) -> None:
    """Print debug messages when DEBUGGING is enabled."""
    if config.DEBUGGING:
        if _use_rich():
            c = console
            if c is not None:
                c.print(f"[debug]! DEBUG[/debug] {message}")
        else:
            print(f"! DEBUG {message}")


def verbose_print(message: str) -> None:
    """Print informational messages when VERBOSITY is set to the highest level."""
    if getattr(config, "VERBOSITY", config.STANDARD) >= config.VERBOSE:
        if _use_rich() and console is not None:
            console.print(message, style="verbose")
        else:
            print(message)


def standard_print(message: str) -> None:
    """Print informational messages for standard verbosity and above."""
    if getattr(config, "VERBOSITY", config.STANDARD) >= config.STANDARD:
        if _use_rich() and console is not None:
            console.print(message, style="verbose")
        else:
            print(message)


def _styled_print(
    message: str,
    *,
    prefix: str = "",
    prefix_style: str | None = None,
    message_style: str | None = None,
    details: list[str] | None = None,
    details_style: str | None = None,
    panel_border_style: str | None = None,
) -> None:
    if _use_rich() and console is not None:
        content = Text()
        if prefix:
            content.append(prefix, style=prefix_style)
        content.append(message, style=message_style)
        if details:
            for detail in details:
                content.append(f"\n  {detail}", style=details_style)

        if panel_border_style and _use_panels():
            console.print(
                Panel(content, border_style=panel_border_style, padding=(0, 1))
            )
        else:
            console.print(content)
        return

    print(f"{prefix}{message}")
    if details:
        for detail in details:
            print(f"  {detail}")


def error_print(message: str, details: list[str] | None = None) -> None:
    """Print error messages. Always displayed regardless of verbosity.

    Args:
        message: Primary error message
        details: Optional list of detail lines to print (indented)
    """
    _styled_print(
        message,
        prefix="! ERROR ",
        prefix_style="error.prefix",
        details=details,
        details_style="error.detail",
        panel_border_style="red",
    )


def warning_print(message: str, details: list[str] | None = None) -> None:
    """Print warning messages. Always displayed regardless of verbosity.

    Args:
        message: Primary warning message
        details: Optional list of detail lines to print (indented)
    """
    _styled_print(
        message,
        prefix="! WARNING ",
        prefix_style="warning.prefix",
        details=details,
        details_style="warning.detail",
        panel_border_style="yellow",
    )


def info_print(message: str) -> None:
    """Print informational messages. Always displayed regardless of verbosity.

    Args:
        message: Informational message
    """
    _styled_print(message, message_style="info")


# ---- Rich browse table and progress helpers ----


def create_browse_table(
    rows_data: list[BrowseRow], participant_headers: list[str]
) -> "Table | None":
    """Create a Rich Table for browse mode display.

    Args:
        rows_data: List of dicts with keys: row_num, category, description, timestamps (dict)
        participant_headers: List of participant IDs for column headers

    Returns:
        Rich Table object if Rich is available, None otherwise
    """
    if not _use_rich():
        return None

    table = Table(
        show_header=True,
        header_style="bold cyan",
        border_style="dim",
        row_styles=["", "dim"],
        expand=False,
    )

    # Add fixed columns
    table.add_column("Row", justify="right", style="bold", width=4)
    table.add_column("Category", style="yellow", max_width=15, overflow="ellipsis")
    table.add_column(
        "Description",
        max_width=config.BROWSE_DESCRIPTION_MAX_WIDTH,
        overflow="ellipsis",
    )

    # Add participant columns
    for p_id in participant_headers:
        table.add_column(
            p_id, max_width=config.BROWSE_TIMESTAMP_MAX_WIDTH, overflow="fold"
        )

    # Add data rows
    for row in rows_data:
        row_values = [
            str(row["row_num"]),
            row["category"],
            row["description"],
        ]
        # Add timestamp values for each participant
        for participant_id in participant_headers:
            timestamp = row["timestamps"].get(participant_id, "-")
            row_values.append(timestamp if timestamp else "-")

        table.add_row(*row_values)

    return table


def format_browse_rows_plain(
    rows_data: list[BrowseRow], participant_headers: list[str]
) -> str:
    """Format browse rows as plain text (fallback when Rich unavailable).

    Args:
        rows_data: List of dicts with keys: row_num, category, description, timestamps (dict)
        participant_headers: List of participant IDs

    Returns:
        Formatted plain text string
    """
    lines = ["-" * 60]

    for row in rows_data:
        lines.append(f"Row {row['row_num']}")
        lines.append(f"  Category: {row['category']}")
        lines.append(f"  Description: {row['description']}")

        # Get participant timestamps
        participant_data = []
        for participant_id in participant_headers:
            timestamp = row["timestamps"].get(participant_id)
            if timestamp:
                participant_data.append(f"    {participant_id}: {timestamp}")

        if participant_data:
            lines.append("  Participants:")
            lines.extend(participant_data)
        else:
            lines.append("  Participants: (no timestamps)")

        lines.append("  ---")

    return "\n".join(lines)


def create_progress_bar(description: str = "Processing"):
    """Create a Rich Progress instance configured for clipgen, or None if unavailable.

    Args:
        description: Default task description

    Returns:
        Configured Progress instance if Rich is available and enabled, else None
    """
    if not use_progress():
        return None

    return Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(bar_width=40),
        MofNCompleteColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TimeElapsedColumn(),
        console=console,
        transient=False,  # Keep progress visible after completion
    )


T = TypeVar("T")


def run_with_spinner(message: str, callback: Callable[[], T]) -> T:
    """Run callback with an indeterminate Rich spinner; if progress disabled, run callback only."""
    if not use_progress() or not RICH_AVAILABLE or console is None:
        return callback()
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console,
        transient=False,
    ) as progress:
        progress.add_task(message, total=None)
        return callback()


def print_mode_heading(label: str, style: str | None = None) -> None:
    """Print a one-line mode heading (bold, optional color). Plain fallback when Rich unavailable.
    style should be a theme key (e.g. 'mode.spreadsheet') so colors render correctly.
    No-op when VERBOSITY is below STANDARD (e.g. CLI mode without -v).
    """
    if getattr(config, "VERBOSITY", config.STANDARD) < config.STANDARD:
        return
    if _use_rich() and console is not None and style:
        console.print()
        console.print(f"[{style}]{label}[/{style}]")
    elif _use_rich() and console is not None:
        console.print()
        console.print(Text(label, style="bold"))
    else:
        print()
        print(f"=== {label} ===")


# ---- Directory and path utilities ----


def get_effective_input_dir() -> Path:
    """Return the effective input directory for source videos."""
    configured = getattr(config, "INPUT_DIR", "") or ""
    if configured:
        return Path(configured).expanduser()
    return Path.cwd()


def get_effective_output_dir() -> Path:
    """Return the effective output directory for generated artifacts."""
    configured = getattr(config, "OUTPUT_DIR", "") or ""
    if configured:
        return Path(configured).expanduser()
    return Path.cwd()


def validate_runtime_directories() -> None:
    """Validate that the effective input directory exists and ensure output directory is ready.

    Behavior:
    - If the effective input directory does not exist, print a warning with guidance and exit.
    - If the effective output directory does not exist, attempt to create it and print a message.
    """
    input_dir = get_effective_input_dir()
    if not input_dir.exists():
        warning_print(
            "Input directory does not exist.",
            [
                f"Configured input directory: {input_dir}",
                "Please create this directory, update INPUT_DIR in config.py,",
                "or pass an existing directory via the -i/--input CLI option.",
            ],
        )
        raise SystemExit(1)

    output_dir = get_effective_output_dir()
    if not output_dir.exists():
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            error_print(
                "Could not create output directory.",
                [
                    f"Target directory: {output_dir}",
                    f"Error: {error}",
                    "Please choose a different output directory or fix permissions,",
                    "then rerun clipgen.",
                ],
            )
            raise SystemExit(1)
        info_print(f"Output directory did not exist and was created: {output_dir}")


def resolve_input_path(name: str) -> Path:
    """Resolve a source filename against the effective input directory."""
    base = get_effective_input_dir()
    path = Path(name)
    if path.is_absolute():
        return path
    return base / path


def resolve_output_path(name: str) -> Path:
    """Resolve an output filename against the effective output directory."""
    base = get_effective_output_dir()
    path = Path(name)
    if path.is_absolute():
        return path
    return base / path


def require_optional(module_name: str, feature_label: str) -> None:
    """Raise ImportError with install instructions if *module_name* is missing."""
    try:
        __import__(module_name)
    except ImportError:
        raise ImportError(
            f"{module_name} is required for {feature_label}. Install with: uv add {module_name}"
        ) from None


def load_json_manifest(filename: str, *, default: Any = None) -> Any:
    """Load a JSON manifest from the output directory.

    Returns parsed data, or *default* on missing/corrupt file.
    """
    path = Path(get_effective_output_dir()) / filename
    if not path.is_file():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default


def save_json_manifest(
    filename: str, data: Any, *, warn_label: str = ""
) -> Path | None:
    """Write *data* as JSON to *filename* in the output directory.

    Writes via a sibling .tmp file and ``os.replace()`` so a crash or ENOSPC
    mid-write leaves the previous manifest intact rather than corrupted.
    Creates parent dirs. Returns the path on success, ``None`` on failure.
    Logs a warning on write/serialization failure using *warn_label*.
    """
    import os as _os

    path = Path(get_effective_output_dir()) / filename
    tmp = path.with_suffix(path.suffix + ".tmp")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(data, ensure_ascii=False, indent=2)
        tmp.write_text(payload, encoding="utf-8")
        _os.replace(tmp, path)
        return path
    except (OSError, TypeError, ValueError) as exc:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        if warn_label:
            warning_print(f"Could not write {warn_label}: {exc}")
        return None


def get_bundled_assets_root() -> Path:
    """Return the base directory for bundled project assets.

    In a PyInstaller one-file build, bundled data is extracted to sys._MEIPASS.
    In source runs, assets are relative to the source directory.
    """
    if getattr(sys, "frozen", False):
        return Path(
            getattr(sys, "_MEIPASS", str(Path(sys.executable).resolve().parent))
        )
    return Path(__file__).resolve().parent


@functools.cache
def get_version() -> str:
    """Return the project version, read from the `VERSION` file in `build/`.

    In a PyInstaller bundle the file is copied to the bundle root by `clipgen.spec`,
    so the source-tree `build/VERSION` and the bundled `VERSION` both resolve via
    `get_bundled_assets_root()`. Falls back to "0.0.0+unknown" if neither exists —
    a missing version must never crash the CLI banner.
    """
    root = get_bundled_assets_root()
    for candidate in (root / "VERSION", root / "build" / "VERSION"):
        try:
            return candidate.read_text(encoding="utf-8").strip()
        except OSError:
            continue
    return "0.0.0+unknown"


def terminate_subprocess(proc: subprocess.Popen, timeout: int = 5) -> None:
    """Terminate a subprocess, escalating to kill if it ignores SIGTERM."""
    proc.terminate()
    try:
        proc.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.wait()


# ---- Filename and study name helpers ----


def normalize_study_name(raw_name: str) -> str:
    """Convert study name to a filesystem-safe format.
    Preserves unicode characters for international study names."""
    # Ensure we're working with a string
    name = str(raw_name)
    name = name.lower()
    name = name.replace("study ", "study")
    name = name.replace(" ", "_")
    return name


def sanitize_filename(text: str) -> str:
    """Remove or replace characters that are unsafe for filenames.
    Preserves unicode characters to support international filenames."""
    # Ensure text is a string and handle unicode properly
    text = str(text)

    # Characters that need special replacement
    text = text.replace("\\", "-")
    text = text.replace("/", "-")
    text = text.replace("?", "_")
    # Characters to remove entirely (filesystem-unsafe characters)
    for char in ["'", '"', ".", ">", "<", "|", ":"]:
        text = text.replace(char, "")
    return text


# ---- Annotation and participant helpers ----


@functools.cache
def get_known_annotation_map() -> dict[str, str]:
    """Return configured annotation tokens mapped to normalized annotation IDs."""
    configured_map = getattr(config, "ANNOTATION_KEYPHRASES", {"!key": "key"})
    normalized_map: dict[str, str] = {}
    for token, annotation_id in configured_map.items():
        normalized_map[str(token).strip().lower()] = str(annotation_id).strip().lower()
    return normalized_map


def normalize_participant_id(participant_value: str) -> str:
    """Strip known annotation tokens from a participant header value."""
    if not participant_value:
        return ""

    known_tokens = set(get_known_annotation_map().keys())
    cleaned_parts = []
    for part in participant_value.split():
        token = part.strip()
        if token and token.lower() not in known_tokens:
            cleaned_parts.append(token)
    return " ".join(cleaned_parts).strip()


# ---- Severity helpers ----


def normalize_severity(raw_value: str) -> str:
    """Map a raw severity cell value to its canonical label.

    Accepts numeric strings ("-4" .. "2") or case-insensitive labels.
    Returns the canonical label or the stripped original if unrecognized.
    """
    stripped = raw_value.strip()
    if not stripped:
        return ""
    label = config.SEVERITY_NUMERIC_TO_LABEL.get(stripped)
    if label:
        return label
    lower = stripped.lower()
    if lower in config.SEVERITY_LABEL_TO_NUMERIC:
        return config.SEVERITY_NUMERIC_TO_LABEL[
            str(config.SEVERITY_LABEL_TO_NUMERIC[lower])
        ]
    return stripped


def severity_sort_key(severity_label: str) -> int:
    """Return a numeric sort key for a severity label (lower = more severe).

    Unrecognized labels sort last.
    """
    return config.SEVERITY_LABEL_TO_NUMERIC.get(severity_label.strip().lower(), 999)


def format_severity_display(severity_label: str) -> str:
    """Format a severity label for display, showing both numeric value and name.

    E.g. "Critical" -> "-4 (Critical)". Returns as-is if not in canonical map.
    """
    num = config.SEVERITY_LABEL_TO_NUMERIC.get(severity_label.strip().lower())
    if num is not None:
        return f"{num} ({severity_label})"
    return severity_label


def get_severity_style(severity_label: str) -> str:
    """Return the Rich theme style key for a severity label."""
    _STYLE_MAP = {
        "critical": "severity.critical",
        "high": "severity.high",
        "medium": "severity.medium",
        "low": "severity.low",
        "n/a": "severity.na",
        "positive": "severity.positive",
        "very positive": "severity.very_positive",
    }
    return _STYLE_MAP.get(severity_label.strip().lower(), "")


def severity_css_class(severity_label: str) -> str:
    """Map a canonical severity label to its tokens.css CSS class.

    "Critical" → "sev-critical", "Very Positive" → "sev-very-positive",
    "N/A" → "sev-na", unknown → "sev-unknown".
    """
    label = severity_label.strip().lower()
    if not label:
        return ""
    if label not in config.SEVERITY_LABEL_TO_NUMERIC:
        return "sev-unknown"
    return "sev-" + label.replace("/", "").replace(" ", "-")


# ---- Frontend config payload ----


def get_frontend_config() -> dict[str, Any]:
    """Return canonical config the JS layer needs (severity, timestamp tokens).

    Embedded in API responses (server.py /api/sheet-data) and exported viewer
    payloads (viewer.py finalize_*) so JS does not duplicate config.py.
    The hardcoded fallback in assets/web/utils.js mirrors this; the contract
    is asserted by tests/test_shared_constants.py.
    """
    severity = []
    for numeric_str, label in sorted(
        config.SEVERITY_NUMERIC_TO_LABEL.items(), key=lambda kv: int(kv[0])
    ):
        severity.append(
            {
                "label": label,
                "rank": int(numeric_str),
                "cssClass": severity_css_class(label),
            }
        )
    annotations = [
        {"id": ann_id, "token": token}
        for token, ann_id in sorted(get_known_annotation_map().items())
    ]
    return {
        "defaultDuration": config.DEFAULT_DURATION_SECONDS,
        "severity": severity,
        "annotationKeyphrases": sorted(get_known_annotation_map().keys()),
        "annotations": annotations,
        "ignoredTimestampTokens": sorted(get_ignored_timestamp_tokens()),
    }


# ---- Column index / letter conversion ----


def index_to_letter(idx: int) -> str:
    """Convert 0-based index to letter label (0='A', 25='Z', 26='AA', etc.).

    Args:
        idx: 0-based index

    Returns:
        Letter label (A-Z, then AA, AB, etc.)
    """
    column_label = ""
    idx += 1  # Convert to 1-based for calculation
    while idx > 0:
        idx -= 1
        column_label = chr(ord("A") + (idx % 26)) + column_label
        idx //= 26
    return column_label


def letter_to_index(letter: str) -> int:
    """Convert letter label to 0-based index ('A'=0, 'Z'=25, 'AA'=26, etc.).

    Args:
        letter: Letter label (case-insensitive)

    Returns:
        0-based index, or -1 if invalid
    """
    letter = letter.upper().strip()
    if not letter or not letter.isalpha():
        return -1
    column_index = 0
    for char in letter:
        column_index = column_index * 26 + (ord(char) - ord("A") + 1)
    return column_index - 1


def safe_cell_a1(row: int | None, col: int | None) -> str:
    """Convert 1-based row/col to A1 notation, returning '' on failure.

    Returns an empty string when row/col are missing, non-int, or non-positive
    (e.g. synthetic ClipRecords from --ss-clips / --transcript-clips use negative
    rows to namespace artifact ids — those produce no spreadsheet A1 reference).
    """
    import gspread.utils

    if not isinstance(row, int) or not isinstance(col, int):
        return ""
    if row < 1 or col < 1:
        return ""
    try:
        return gspread.utils.rowcol_to_a1(row, col)
    except Exception:
        return ""


def _clip_metadata_fields(
    clip: "ClipRecord", base_video: str, start_str: str, end_str: str
) -> dict[str, Any]:
    """Extract the shared per-segment metadata that every persisted record needs.

    Used by both ``build_artifact_record`` (manifest artifacts) and
    ``build_reel_component`` (reel-component records). The two shapes only differ
    by file-specific fields (id/file/type/thumbnail), so the body of every
    persisted record flows from one place.
    """
    cell = clip.get("cell")
    cell_row = getattr(cell, "row", None)
    cell_col = getattr(cell, "col", None)
    return {
        "start": timestamp_to_seconds(start_str),
        "end": timestamp_to_seconds(end_str),
        "study": clip.get("study", ""),
        "participant": clip.get("participant", ""),
        "category": clip.get("category", ""),
        "severity": clip.get("severity", ""),
        "description": clip.get("desc", ""),
        "cellRow": cell_row,
        "cellCol": cell_col,
        "cellA1": safe_cell_a1(cell_row, cell_col),
        "annotations": list(clip.get("cell_annotations", [])),
        "sourceVideo": base_video,
    }


def build_artifact_record(
    clip: "ClipRecord",
    base_video: str,
    out_path: str,
    start_str: str,
    end_str: str,
    *,
    artifact_type: str,
    seg_idx: int,
) -> dict[str, Any]:
    """Build one artifact dict from a clip record + one segment.

    Single source of truth for the artifact record shape used by clipgen_manifest
    and the timeline viewer. Callers may add or override fields after the call
    (e.g. transcripts append ``transcriptFormat``).
    """
    cell = clip.get("cell")
    cell_row = getattr(cell, "row", None)
    cell_col = getattr(cell, "col", None)
    return {
        "id": f"a{cell_row or 0}c{cell_col or 0}s{seg_idx}",
        "type": artifact_type,
        "file": Path(out_path).name,
        "thumbnail": "",
        **_clip_metadata_fields(clip, base_video, start_str, end_str),
    }


def build_reel_component(
    clip: "ClipRecord",
    base_video: str,
    start_str: str,
    end_str: str,
) -> dict[str, Any]:
    """Build one reel-component dict from a clip record + one segment.

    Reel components describe an input segment used to assemble a reel — they
    share the artifact record's per-segment metadata shape but omit the
    file/id/type fields (the rendered output is the reel itself, not the
    component). Stored in the ``components`` list of a reel manifest entry.
    """
    return _clip_metadata_fields(clip, base_video, start_str, end_str)


# ---- Timestamp parsing pipeline ----
#
# Reading order: token splitting/cleaning → add_duration → _parse_single_timestamp_token
# → higher-level parsers (has_non_ignored_timestamp_content, parse_cell_annotations,
# parse_timestamps).


def _split_timestamp_tokens(cell_value: str) -> list[str]:
    """Split a cell value into normalized timestamp/annotation tokens."""
    return (
        cell_value.lower().replace("+", " ").replace(";", " ").replace(",", " ").split()
    )


def _clean_timestamp_token(token: str) -> str:
    """Normalize one token before timestamp parsing."""
    return token.strip().rstrip(",").rstrip("-").replace(".", ":")


def get_ignored_timestamp_tokens() -> set[str]:
    """Return configured ignored non-timestamp tokens in normalized form."""
    configured_tokens = getattr(config, "IGNORED_TIMESTAMP_TOKENS", set())
    normalized_tokens: set[str] = set()
    for token in configured_tokens:
        cleaned = _clean_timestamp_token(str(token).strip().lower())
        if cleaned:
            normalized_tokens.add(cleaned)
    return normalized_tokens


def add_duration(start_time: str) -> str | None:
    """Add default duration to a start timestamp.

    Adds DEFAULT_DURATION_SECONDS to the given start timestamp to create
    an end timestamp. Used when only a start time is provided.

    Args:
        start_time: Start timestamp in format MM:SS or HH:MM:SS

    Returns:
        The new timestamp string with duration added, or None if the timestamp
        format is invalid.
    """
    start_seconds = timestamp_to_seconds(start_time)
    if start_seconds is None:
        warning_print(
            f"Could not parse single timestamp '{start_time}' to add default duration.",
            [
                "Expected format: MM:SS or HH:MM:SS (e.g., 12:34 or 1:23:45)",
                "This timestamp will be skipped.",
            ],
        )
        return None
    has_hours = start_time.count(":") >= 2
    return seconds_to_timestamp(
        int(start_seconds) + config.DEFAULT_DURATION_SECONDS,
        force_hours=has_hours,
    )


def _parse_single_timestamp_token(token: str) -> tuple[str, str | None] | None:
    """Parse one token into a (start_time, end_time) pair, or None if invalid/skip.

    Handles: dash range (start-end), single timestamp with colon (add default
    duration), blank token (skip), or unrecognized format (skip).

    Args:
        token: A single timestamp token, e.g. "1:23-1:45", "2:30", or "".

    Returns:
        (start_time, end_time) tuple if parseable, else None (caller skips).

    Examples:
        "1:23-1:45" -> ("1:23", "1:45"); "2:30" -> ("2:30", "2:45") with default duration.
    """
    if token == "":
        return None
    # Dash range: "start-end". Require a digit before the dash so we don't
    # treat a leading dash (e.g. "-5") or non-time dash as a range.
    if "-" in token:
        dash_pos = token.find("-")
        if dash_pos > 0 and token[dash_pos - 1].isdigit():
            return (token[:dash_pos], token[dash_pos + 1 :])
        return None
    # Single timestamp with colon: use as start and add default duration for end.
    # Require a digit before the first colon so we only match time-like strings.
    if ":" in token:
        colon_pos = token.find(":")
        if colon_pos > 0 and token[colon_pos - 1].isdigit():
            end_time = add_duration(token)
            if end_time is not None:
                return (token, end_time)
        return None
    return None


def has_non_ignored_timestamp_content(cell_value: str) -> bool:
    """Return True when cell content is more than ignored timestamp tokens.

    Cells containing only ignored tokens (e.g. "x") should not produce the
    generic "No valid timestamps found" warning.
    """
    ignored_tokens = get_ignored_timestamp_tokens()
    for raw_token in _split_timestamp_tokens(cell_value):
        token = _clean_timestamp_token(raw_token)
        if not token:
            continue
        if _parse_single_timestamp_token(token) is not None:
            return True
        if token not in ignored_tokens:
            return True
    return False


def parse_cell_annotations(
    cell_value: str, annotation_map: dict[str, str | None] | None = None
) -> tuple[str, dict[str, set[int]], set[str]]:
    """Extract inline annotation tokens and map them to parsed timestamp indexes.

    Semantics: an annotation token marks the preceding parseable timestamp token.
    The returned cleaned cell value has annotation tokens removed.
    """
    known_annotations = annotation_map or get_known_annotation_map()
    cleaned_tokens: list[str] = []
    segment_annotations: dict[str, set[int]] = {}
    cell_annotations: set[str] = set()
    parsed_timestamp_count = 0

    for raw_token in _split_timestamp_tokens(cell_value):
        token = _clean_timestamp_token(raw_token)
        if not token:
            continue

        annotation_id = known_annotations.get(token)
        if annotation_id:
            cell_annotations.add(annotation_id)
            if parsed_timestamp_count > 0:
                segment_annotations.setdefault(annotation_id, set()).add(
                    parsed_timestamp_count - 1
                )
            continue

        cleaned_tokens.append(token)
        if _parse_single_timestamp_token(token) is not None:
            parsed_timestamp_count += 1

    return (" ".join(cleaned_tokens), segment_annotations, cell_annotations)


def timestamp_to_seconds(ts_str: str) -> float | None:
    """Convert MM:SS or HH:MM:SS timestamp string to seconds.

    Args:
        ts_str: Timestamp string in MM:SS or HH:MM:SS format

    Returns:
        Total seconds as float, or None if the timestamp cannot be parsed.
    """
    ts = (ts_str or "").strip()
    if not ts:
        return None

    formats = ["%M:%S", "%H:%M:%S"]
    for fmt in formats:
        try:
            parsed = datetime.strptime(ts, fmt)
            return float(
                parsed.hour * config.SECONDS_PER_HOUR
                + parsed.minute * config.SECONDS_PER_MINUTE
                + parsed.second
            )
        except ValueError:
            continue
    return None


def parse_timestamps(
    cell_value: str, cell_ref: str | None = None
) -> list[tuple[str, str]]:
    """Parse timestamp pairs from a cell value string.

    Pipeline: (1) Normalize delimiters to spaces and split into tokens,
    (2) Clean each token and parse into (start, end) pairs,
    (3) Report any unparseable tokens as warnings.

    Supported formats: "MM:SS-MM:SS", "HH:MM:SS-HH:MM:SS", or a single time
    "MM:SS"/"HH:MM:SS" (end time is start + default duration). Delimiters
    between multiple pairs: space, comma, semicolon, or plus.

    Args:
        cell_value: The raw cell value containing timestamps
        cell_ref: Optional cell reference (e.g., 'B5') for error messages

    Returns:
        A list of (start_time, end_time) tuples. Invalid tokens are skipped
        and reported via warning_print.
    """
    if config.DEBUGGING:
        ic(cell_value, cell_ref)
    parsed_timestamps = []
    skipped_timestamps = []
    ignored_tokens = get_ignored_timestamp_tokens()
    # Unify delimiters (+, ;, ,) to spaces so split() yields one token per time or range
    raw_times = _split_timestamp_tokens(cell_value)
    if config.DEBUGGING:
        ic(raw_times)
    debug_print(f"raw_times content after split is {raw_times}")
    debug_print(f"Timestamp list raw_times is {len(raw_times)} entries long")

    # Clean each token (strip, normalize trailing punctuation, use colon for decimals) and parse
    raw_times = [_clean_timestamp_token(t) for t in raw_times]
    for token in raw_times:
        debug_print(f"Cleaning timestamp {token}")
        pair = _parse_single_timestamp_token(token)
        if pair is not None:
            if config.DEBUGGING and len(pair) == 2:
                ic(pair)
            parsed_timestamps.append(pair)
        elif token and token not in ignored_tokens:
            skipped_timestamps.append(token)

    # Report skipped timestamps: list up to MAX_SKIPPED_TIMESTAMPS_TO_SHOW, then "... and N more"
    if skipped_timestamps:
        if config.DEBUGGING:
            ic(skipped_timestamps)
        # Only show detailed skipped-timestamp warnings at verbose verbosity.
        if getattr(config, "VERBOSITY", config.STANDARD) >= config.VERBOSE:
            cell_info = f" in cell {cell_ref}" if cell_ref else ""
            details = []
            for ts in skipped_timestamps[: config.MAX_SKIPPED_TIMESTAMPS_TO_SHOW]:
                details.append(f"    '{ts}'")
            if len(skipped_timestamps) > config.MAX_SKIPPED_TIMESTAMPS_TO_SHOW:
                details.append(
                    f"    ... and {len(skipped_timestamps) - config.MAX_SKIPPED_TIMESTAMPS_TO_SHOW} more"
                )
            details.append(
                "  Expected formats: MM:SS-MM:SS, HH:MM:SS-HH:MM:SS, or single timestamps like MM:SS"
            )
            warning_print(
                f"Skipped {len(skipped_timestamps)} unparseable timestamp(s){cell_info}:",
                details,
            )

    if config.DEBUGGING:
        ic(parsed_timestamps)
    return parsed_timestamps


# ---- Clock/baseline timestamp conversion ----


def _clock_to_seconds(ts: str) -> int | None:
    """Parse a clock-style timestamp into total seconds.

    Accepts HH:MM:SS, HH:MM, or MM:SS. Returns None if parsing fails.
    """
    value = (ts or "").strip()
    if not value:
        return None

    # Choose format based on number of components to avoid treating "22:00"
    # as 22 minutes instead of 22 hours.
    parts = value.split(":")
    if len(parts) == 3:
        try:
            parsed = datetime.strptime(value, "%H:%M:%S")
            return (
                parsed.hour * config.SECONDS_PER_HOUR
                + parsed.minute * config.SECONDS_PER_MINUTE
                + parsed.second
            )
        except ValueError:
            return None
    if len(parts) == 2:
        # Prefer HH:MM for clock-style values like "22:00"; fall back to MM:SS.
        for fmt in ("%H:%M", "%M:%S"):
            try:
                parsed = datetime.strptime(value, fmt)
                return (
                    parsed.hour * config.SECONDS_PER_HOUR
                    + parsed.minute * config.SECONDS_PER_MINUTE
                    + parsed.second
                )
            except ValueError:
                continue
        return None
    return None


def seconds_to_timestamp(total_seconds: int, *, force_hours: bool = False) -> str:
    """Format a non-negative number of seconds as H:MM:SS or M:SS."""
    if total_seconds < 0:
        total_seconds = 0
    hours, rem = divmod(total_seconds, config.SECONDS_PER_HOUR)
    minutes, seconds = divmod(rem, config.SECONDS_PER_MINUTE)
    if hours > 0 or force_hours:
        return f"{hours:d}:{minutes:02d}:{seconds:02d}"
    return f"{minutes:d}:{seconds:02d}"


def convert_clock_pairs_to_relative(
    pairs: list[tuple[str, str]],
    baseline: str,
    cell_ref: str | None = None,
) -> list[tuple[str, str]]:
    """Convert absolute clock (start, end) pairs to relative offsets using a baseline.

    Returns a new list of (start, end) pairs in relative time. Invalid or
    non-positive-length segments are skipped with warnings.
    """
    baseline_seconds = _clock_to_seconds(baseline)
    if baseline_seconds is None:
        cell_info = f" in cell {cell_ref}" if cell_ref else ""
        warning_print(
            f"Ignoring invalid baseline timestamp '{baseline}'{cell_info}.",
            [
                "Baseline must use clock format like HH:MM:SS or MM:SS.",
                f"Dropping {len(pairs)} clock-style segment(s) in this column.",
            ],
        )
        return []

    # Overnight recordings: segments after midnight will be < baseline_seconds.
    # If the resulting wraparound offset is within a reasonable recording
    # window (<= 12 hours), treat it as a day rollover; otherwise reject as
    # a typo / pre-baseline entry.
    seconds_per_day = 24 * config.SECONDS_PER_HOUR
    wrap_window = 12 * config.SECONDS_PER_HOUR

    result: list[tuple[str, str]] = []
    skipped: list[str] = []
    for start_str, end_str in pairs:
        start_s = _clock_to_seconds(start_str)
        end_s = _clock_to_seconds(end_str)
        if start_s is None or end_s is None:
            skipped.append(f"{start_str}-{end_str}")
            continue
        start_rel = start_s - baseline_seconds
        end_rel = end_s - baseline_seconds
        if start_rel < 0:
            wrapped = start_rel + seconds_per_day
            if 0 <= wrapped <= wrap_window:
                start_rel = wrapped
                end_rel = end_rel + seconds_per_day
        elif end_rel < start_rel:
            # Span crosses midnight while baseline was pre-midnight.
            end_rel += seconds_per_day
        if start_rel < 0 or end_rel <= 0 or end_rel <= start_rel:
            skipped.append(f"{start_str}-{end_str}")
            continue
        result.append(
            (
                seconds_to_timestamp(start_rel, force_hours=True),
                seconds_to_timestamp(end_rel, force_hours=True),
            )
        )

    if skipped:
        cell_info = f" in cell {cell_ref}" if cell_ref else ""
        # Standard verbosity: short summary so users notice silent drops.
        # Verbose: full list and explanation.
        verbosity = getattr(config, "VERBOSITY", config.STANDARD)
        if verbosity >= config.VERBOSE:
            details = [f"    '{s}'" for s in skipped]
            details.append(
                "  These segments were before the baseline, invalid, or zero/negative length."
            )
            warning_print(
                f"Skipped {len(skipped)} clock-based timestamp segment(s){cell_info} when converting to relative:",
                details,
            )
        else:
            warning_print(
                f"Skipped {len(skipped)} clock-based timestamp segment(s){cell_info} "
                "(before baseline, invalid, or zero/negative length). "
                "Re-run with verbose verbosity for the full list.",
            )

    return result


def cluster_spans(
    spans: list[tuple[float, float]],
    *,
    gap_seconds: float,
    pad_pre: float = 0.0,
    pad_post: float = 0.0,
    max_duration: float = 0.0,
) -> list[tuple[float, float, list[int]]]:
    """Group (start, end) spans whose gap <= gap_seconds, then pad and optionally split.

    Returns a list of (cluster_start, cluster_end, member_indices) tuples, where
    member_indices are the original indices contributing to that cluster.

    - gap_seconds <= 0 yields one cluster per input span (no merging).
    - Padding is applied after clustering: cluster_start = max(0, start - pad_pre);
      cluster_end is not clamped to a video duration (callers may clamp if needed —
      ffmpeg tolerates an end past EOF).
    - When max_duration > 0 and a padded cluster exceeds it, the cluster is split
      into ceil(duration / max_duration) sub-clips that each claim all members.
    - Input is not mutated; spans are sorted internally by (start, end).
    """
    if not spans:
        return []
    indexed = sorted(enumerate(spans), key=lambda kv: (kv[1][0], kv[1][1]))
    groups: list[tuple[float, float, list[int]]] = []
    cur_start, cur_end = indexed[0][1]
    cur_members: list[int] = [indexed[0][0]]
    for orig_idx, (s, e) in indexed[1:]:
        if gap_seconds <= 0 or s - cur_end > gap_seconds:
            groups.append((cur_start, cur_end, cur_members))
            cur_start, cur_end, cur_members = s, e, [orig_idx]
        else:
            if e > cur_end:
                cur_end = e
            cur_members.append(orig_idx)
    groups.append((cur_start, cur_end, cur_members))

    out: list[tuple[float, float, list[int]]] = []
    for s, e, members in groups:
        ps = max(0.0, s - pad_pre)
        pe = e + pad_post
        if max_duration > 0 and (pe - ps) > max_duration:
            t = ps
            while t < pe:
                sub_end = min(t + max_duration, pe)
                out.append((t, sub_end, list(members)))
                t = sub_end
        else:
            out.append((ps, pe, members))
    return out


# ---- Interactive control flow ----


class QuitProgram(Exception):
    """Signal that the user requested to quit the program from an interactive prompt."""


class TopToSpreadsheet(Exception):
    """Signal that the user requested to return to spreadsheet selection."""


class BackToModeSelection(Exception):
    """Signal that the user requested to return to the main mode selection prompt."""


def check_navigation_keywords(value: str) -> None:
    """Raise a navigation exception if the first token is a control keyword.

    Recognizes: quit/exit, top, back. Does nothing for empty or non-keyword input.
    """
    if not value:
        return
    first_token = value.split()[0].lower()
    if first_token in ("quit", "exit"):
        info_print("Exiting clipgen.")
        raise QuitProgram()
    if first_token == "top":
        info_print("Returning to spreadsheet selection.")
        raise TopToSpreadsheet()
    if first_token == "back":
        info_print("Returning to mode selection.")
        raise BackToModeSelection()


def read_user_input(prompt: str) -> str:
    """Read user input and handle global control keywords.

    Recognizes the following keywords when they appear as the first token:
    - 'quit' / 'exit' -> quit program
    - 'top'           -> return to spreadsheet selection
    - 'back'          -> return to main mode selection
    """
    raw = input(prompt)
    value = raw.strip()
    check_navigation_keywords(value)
    return value


def suggest_close_match(
    user_input: str,
    valid_options: list[str],
    *,
    prompt_prefix: str = "Did you mean",
    cutoff: float = 0.6,
) -> str | None:
    """Find a fuzzy match and prompt user for confirmation. Returns matched option or None."""
    lower_to_original: dict[str, str] = {}
    for option in valid_options:
        key = option.strip().lower()
        if key not in lower_to_original:
            lower_to_original[key] = option
    matches = difflib.get_close_matches(
        user_input.strip().lower(), list(lower_to_original.keys()), n=1, cutoff=cutoff
    )
    if not matches:
        return None
    original = lower_to_original[matches[0]]
    display = original.strip()
    if NO_INPUT_MODE:
        return None
    yn = read_user_input(f"{prompt_prefix} '{display}'? [y/n]\n>> ")
    if yn.strip().lower() == "y":
        return original
    return None


def set_program_settings() -> bool:
    """Interactive settings screen with grid display and type-safe value changes.

    Returns:
        True if a setting was changed, False otherwise.
    """
    settings_list = list(config.SETTINGS_DESCRIPTIONS.keys())

    if _use_rich() and console is not None:
        table = Table(
            show_header=True,
            header_style="bold cyan",
            border_style="dim",
            expand=False,
        )
        table.add_column("#", justify="right", style="bold", width=3)
        table.add_column("Setting", style="yellow", min_width=12)
        table.add_column("Value", style="green", min_width=6)
        table.add_column("Description", max_width=60, overflow="fold")
        for i, name in enumerate(settings_list, 1):
            table.add_row(
                str(i),
                name,
                str(getattr(config, name, "?")),
                config.SETTINGS_DESCRIPTIONS[name],
            )
        console.print(table)
    else:
        for i, name in enumerate(settings_list, 1):
            val = getattr(config, name, "?")
            info_print(
                f"  {i:>2}. {name:<30} = {val!s:<10}  {config.SETTINGS_DESCRIPTIONS[name]}"
            )

    choice = read_user_input(
        "\nSetting to change (number or name, or empty to go back):\n>> "
    )
    if not choice:
        return False

    setting_name = None
    if choice.isdigit():
        idx = int(choice) - 1
        if 0 <= idx < len(settings_list):
            setting_name = settings_list[idx]
    else:
        upper = choice.strip().upper()
        if upper in config.SETTINGS_DESCRIPTIONS:
            setting_name = upper

    if setting_name is None:
        error_print(f"Unknown setting: '{choice}'")
        return False

    current_value = getattr(config, setting_name)
    info_print(f"  Current value: {current_value!r}")
    info_print(f"  {config.SETTINGS_DESCRIPTIONS[setting_name]}")

    new_raw = read_user_input("\nNew value (empty to cancel):\n>> ")
    if not new_raw:
        return False

    current_type = type(current_value)
    try:
        if current_type is bool:
            converted = new_raw.strip().lower() in ("true", "1", "yes", "on")
        elif current_type is int:
            converted = int(new_raw)
        elif current_type is float:
            converted = float(new_raw)
        else:
            converted = new_raw
    except (ValueError, TypeError):
        error_print(f"Invalid value '{new_raw}' for type {current_type.__name__}")
        return False

    setattr(config, setting_name, converted)
    info_print(f"  '{setting_name}' set to {converted!r}")
    return True


# ---- Miscellaneous utilities ----


def format_filesize(size_bytes: float, precision: int = 2) -> str:
    """Format byte size as human-readable string.

    Args:
        size_bytes: Size in bytes
        precision: Number of decimal places (default: 2)

    Returns:
        Formatted string with appropriate unit (B, KB, MB, GB, TB)
    """
    suffixes = ["B", "KB", "MB", "GB", "TB"]
    suffix_index = 0
    # Keep dividing by 1024 until size is under 1024 or we reach TB (index 4)
    while size_bytes > 1024 and suffix_index < 4:
        suffix_index += 1
        size_bytes = size_bytes / 1024
    return f"{size_bytes:.{precision}f}{suffixes[suffix_index]}"


def get_current_time() -> str:
    """Get current time as formatted string.

    Returns:
        Current time in format 'YYYY-MM-DD HH:MM:SS'
    """
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ---- Participant video discovery ----


def discover_participant_videos(study_name: str = "") -> list[dict[str, Any]]:
    """Scan input directory for source video files and return participant info dicts.

    Extracts participant IDs from filenames matching ``{study}_{participant}{FILEFORMAT}``.
    Only files whose participant segment starts with a recognised prefix
    (``config.PARTICIPANT_PREFIXES``) are included.

    Returns:
        List of ``{"id": str, "video_path": str, "has_video": True}`` dicts,
        sorted by filename.
    """
    input_dir = Path(get_effective_input_dir())
    if not input_dir.is_dir():
        return []
    participants: list[dict[str, Any]] = []
    for path in sorted(input_dir.glob(f"*{config.FILEFORMAT}")):
        name = path.stem
        parts = name.rsplit("_", 1)
        if len(parts) == 2:
            pid = parts[1]
            if pid and pid[0] in config.PARTICIPANT_PREFIXES:
                participants.append(
                    {
                        "id": pid,
                        "video_path": str(path),
                        "has_video": True,
                    }
                )
    return participants


# ---- Flask blueprint helpers ----


def register_static_routes(
    bp: Any,
    index_html: str,
    *,
    media_dir_getter: Any = None,
    media_error: str = "Media directory not configured",
    icons: bool = False,
    logos: bool = True,
) -> None:
    """Register standard static-file serving routes on a Flask Blueprint.

    Always registers ``/`` (index) and ``/<path:filename>`` (static assets).
    Optionally registers ``/icons/<path:filename>``, ``/logos/<path:filename>``,
    and ``/media/<path:filename>``.

    Args:
        bp: Flask Blueprint to register routes on.
        index_html: Filename of the HTML page served at ``/``.
        media_dir_getter: Callable returning the current media directory path.
            When provided, a ``/media/<path:filename>`` route is registered.
        media_error: Error message returned (500) when the media dir is falsy.
        icons: When True, registers ``/icons/<path:filename>`` from ``assets/icons/``.
        logos: When True (default), registers ``/logos/<path:filename>`` from
            ``assets/logos/`` so favicons and the brand mark are available to
            every served page.
    """
    from flask import Response, jsonify, send_from_directory

    assets_dir = get_bundled_assets_root() / "assets" / "web"

    @bp.route("/")
    def serve_index() -> Response:
        return send_from_directory(assets_dir, index_html)

    @bp.route("/<path:filename>")
    def serve_static(filename: str) -> Response:
        return send_from_directory(assets_dir, filename)

    if icons:
        icons_dir = get_bundled_assets_root() / "assets" / "icons"

        @bp.route("/icons/<path:filename>")
        def serve_icons(filename: str) -> Response:
            return send_from_directory(icons_dir, filename)

    if logos:
        logos_dir = get_bundled_assets_root() / "assets" / "logos"

        @bp.route("/logos/<path:filename>")
        def serve_logos(filename: str) -> Response:
            return send_from_directory(logos_dir, filename)

    if media_dir_getter is not None:

        @bp.route("/media/<path:filename>")
        def serve_media(filename: str) -> Response | tuple[Response, int]:
            d = media_dir_getter()
            if not d:
                return jsonify({"ok": False, "error": media_error}), 500
            return send_from_directory(d, filename)


# ---- Native folder picker ------------------------------------------------
#
# clipgen is a local tool, so when the Start overlay's Browse button is
# clicked we can open the host OS's native folder dialog and pipe the path
# back to the browser. The server-side approach below shells out to platform
# tooling so we don't need extra GUI dependencies.


def open_native_folder_picker(initial_dir: str = "") -> str | None:
    """Open a native folder picker and return the chosen path.

    Returns the chosen folder's absolute path on confirm, ``None`` when the
    user cancels or no native dialog is available on this platform.
    """
    import subprocess

    if sys.platform == "darwin":
        prompt = "Select a folder for clipgen"
        safe_initial = ""
        if initial_dir:
            try:
                if Path(initial_dir).is_dir():
                    safe_initial = initial_dir
            except OSError:
                pass
        if safe_initial:
            # Escape backslashes first, then double quotes, for safe embedding
            # in an AppleScript double-quoted string literal.
            escaped = safe_initial.replace("\\", "\\\\").replace('"', '\\"')
            script = (
                f'set chosenFolder to choose folder with prompt "{prompt}" '
                f'default location POSIX file "{escaped}"\n'
                "return POSIX path of chosenFolder"
            )
        else:
            script = (
                f'set chosenFolder to choose folder with prompt "{prompt}"\n'
                "return POSIX path of chosenFolder"
            )
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                timeout=300,
                check=False,
            )
        except (OSError, subprocess.TimeoutExpired):
            return None
        if result.returncode != 0:
            # User cancelled or AppleScript failed — both are non-errors here.
            return None
        path = result.stdout.strip().rstrip("/")
        return path or None

    # Tkinter fallback for Linux/Windows when available. uv-managed Pythons
    # often skip Tk, so this is best-effort — callers should accept None.
    try:
        import tkinter
        from tkinter import filedialog
    except ImportError:
        return None
    try:
        root = tkinter.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askdirectory(
            initialdir=initial_dir or str(Path.home()),
            title="Select a folder for clipgen",
        )
        root.destroy()
    except Exception:
        return None
    return path or None
