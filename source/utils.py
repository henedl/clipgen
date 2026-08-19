"""Utility functions for clipgen."""

import contextlib
import difflib
import functools
import json
import math
import os
import shutil
import subprocess
import sys
import threading
from collections.abc import Callable
from datetime import datetime
from numbers import Integral, Real
from pathlib import Path
from typing import Any, TypedDict, TypeVar

import config


# ---- Non-interactive mode flag ----
# Set to True from cli.py when --no-input is passed. Targeted guards in
# clipgen.py / excel_io.py / pipeline.py / video.py / utils.suggest_close_match
# read this flag to fail fast (or skip) instead of blocking on stdin.
NO_INPUT_MODE: bool = False

# ---- Windowed-launch flag ----
# Set to True from cli.py when the run will open a desktop window instead of a
# console. Read by fatal_startup_error: a Finder/Explorer launch has no attached
# terminal, so anything printed before the window exists is invisible.
GUI_LAUNCH: bool = False

# ---- Desktop window-chrome flag ----
# Set by desktop.launch_desktop before the server starts, cleared when the window
# closes. Holds the chrome style the native window uses ("macos"), or None for a
# browser launch. render_index_html() turns it into an html[data-desktop-chrome]
# attribute so the topnav can inset itself for the traffic lights; it has to be a
# server-side flag rather than a `window.pywebview` check because pywebview injects
# its bridge asynchronously, long after the bar has laid out.
DESKTOP_CHROME: str | None = None


# ---- Native (C/ObjC-level) stderr suppression ----


@contextlib.contextmanager
def suppress_native_stderr():
    """Temporarily silence OS file descriptor 2 (stderr) for the duration.

    Redirects fd 2 to /dev/null and restores it in ``finally``. Unlike
    reassigning ``sys.stderr``, this also swallows writes from C/Objective-C
    libraries (e.g. the dynamic linker / ObjC runtime), which write to the
    fd directly. Use sparingly and around the narrowest possible block.
    """
    saved_fd = os.dup(2)
    null_fd = os.open(os.devnull, os.O_WRONLY)
    try:
        os.dup2(null_fd, 2)
        yield
    finally:
        os.dup2(saved_fd, 2)
        os.close(null_fd)
        os.close(saved_fd)


_vision_libs_preloaded = False


def preload_vision_libs_quietly(on_phase: Callable[[str], None] | None = None) -> None:
    """Import ``cv2`` once, early, with native stderr silenced.

    The wheel's dylibs can emit native noise on load, and the import can take
    ~10s of disk I/O on a cold machine — pre-loading here under
    :func:`suppress_native_stderr` means later lazy ``import cv2`` calls find
    it already resident, and the server's boot page can narrate the wait
    (*on_phase* is called with ``"cv2"`` before the import). Idempotent; a
    missing install is a no-op. This used to preload PyAV too, whose bundled
    ``libavdevice`` duplicated cv2's and tripped a macOS duplicate-ObjC-class
    warning — PyAV is no longer a dependency (see transcripts._ensure_av_stub).
    """
    global _vision_libs_preloaded
    if _vision_libs_preloaded:
        return
    _vision_libs_preloaded = True
    if on_phase is not None:
        on_phase("cv2")
    with suppress_native_stderr():
        try:
            import cv2  # noqa: F401
        except ImportError:
            pass


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
    # Set by the pipeline only when a participant resolves to 2+ source videos
    # (one continuous timeline). Each entry is (path, duration, cumulative_start).
    # Absent for the single-video fast path. See video.build_source_timeline.
    source_timeline: list[tuple[str, int, int]]


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


# Package managers install here, and macOS does not put any of them on a GUI
# process's PATH. A Finder-launched .app gets only /usr/bin:/bin:/usr/sbin:/sbin,
# so shutil.which("ffmpeg") misses a perfectly good Homebrew install.
_GUI_PATH_DIRS = (
    "/opt/homebrew/bin",  # Homebrew on Apple Silicon
    "/usr/local/bin",  # Homebrew on Intel, and most manual installs
    "/opt/local/bin",  # MacPorts
)


def augment_path_for_gui_launch() -> list[str]:
    """Make package-manager binaries findable in a frozen macOS GUI launch.

    Returns the directories actually added.

    Double-clicking a .app used to work only by accident: the bundle shipped a
    shim that ran ``open -a Terminal``, and Terminal starts a login shell that
    sources the user's profile. Launching the binary directly (as the app now
    does) drops that inheritance, and startup aborted on a missing ffmpeg with no
    window and nothing on screen.

    Entries are *appended*, never prepended — this is about discoverability, not
    about overriding a resolution order the user already has. Source runs are
    left alone: they already carry the developer's real PATH.
    """
    if not getattr(sys, "frozen", False) or sys.platform != "darwin":
        return []
    current = [p for p in os.environ.get("PATH", "").split(os.pathsep) if p]
    added = [d for d in _GUI_PATH_DIRS if d not in current and Path(d).is_dir()]
    if added:
        os.environ["PATH"] = os.pathsep.join([*current, *added])
    return added


def prepend_bundled_bin_to_path() -> str | None:
    """Put the bundle's own tools first on PATH in a frozen launch.

    Returns the directory that was (or already is) at the head, or None when
    there is nothing to do (source runs, or a bundle without a bin/ dir).

    The desktop builds ship pinned ffmpeg/ffprobe under <bundle>/bin (see
    build/fetch_binaries.py). Unlike ``augment_path_for_gui_launch`` this
    *prepends*: the app must run the ffmpeg it was built and feature-verified
    with, not whatever an older Homebrew install resolves to — the escape
    hatch for power users is replacing the file inside the bundle. Source
    runs are left alone, matching the append helper.
    """
    if not getattr(sys, "frozen", False) or sys.platform not in ("darwin", "win32"):
        return None
    bin_dir = get_bundled_assets_root() / "bin"
    if not bin_dir.is_dir():
        return None
    current = [p for p in os.environ.get("PATH", "").split(os.pathsep) if p]
    bundled = str(bin_dir)
    if current[:1] != [bundled]:
        current = [bundled, *[p for p in current if p != bundled]]
        os.environ["PATH"] = os.pathsep.join(current)
    return bundled


def install_guidance_lines(
    *,
    brew_command: str,
    linux: list[str],
    windows: list[str],
    download_url: str,
    verify_commands: list[str],
) -> list[str]:
    """Build platform-specific "how do I install this" lines for a missing tool.

    Shared by the ffmpeg and Ollama guidance so the two stay in step — both are
    surfaced to users who have no console (a native alert, or the browser), and
    both used to hand macOS users ``brew install …`` unconditionally. That is a
    dead end on a machine without Homebrew, which is the *default* state of a
    fresh Mac: nothing else in clipgen requires it. So the macOS branch probes
    for brew and leads with the direct download when it is absent.

    ``download_url`` is always reachable in the output, on every platform — it
    is the one instruction that works regardless of what package manager the
    user does or does not have.

    **Line 0 is always a complete, actionable instruction on its own.** The
    browser surfaces (Overview's gate, the Settings note, the Transcripts
    summary hint) are one-liners and show only the first line; the full list
    goes to the terminal and to the runtime dialog. A brewless Mac briefly led
    with a bare "Homebrew is not installed." header here, which told those
    surfaces' users nothing they could act on.
    """
    if sys.platform == "darwin":
        if shutil.which("brew") is not None:
            platform_specific = [f"macOS: install with Homebrew: {brew_command}"]
        else:
            platform_specific = [
                f"macOS: download from {download_url}",
                "  Homebrew is not installed. To use it instead, install it from",
                f"  https://brew.sh, then: {brew_command}",
            ]
    elif sys.platform.startswith("linux"):
        platform_specific = list(linux)
    elif sys.platform.startswith("win"):
        platform_specific = list(windows)
    else:
        platform_specific = []

    if any(download_url in line for line in platform_specific):
        download_lines = []
    elif platform_specific:
        download_lines = [f"Or download from: {download_url}"]
    else:
        download_lines = [f"Download from: {download_url}"]

    return [
        *platform_specific,
        *download_lines,
        "Then verify in a new terminal:",
        *[f"  {command}" for command in verify_commands],
    ]


def _show_native_alert(title: str, message: str) -> None:
    """Best-effort native error dialog. Never raises."""
    try:
        if sys.platform == "darwin":
            # json.dumps yields an AppleScript-compatible quoted literal.
            script = (
                f"display alert {json.dumps(title)} "
                f"message {json.dumps(message)} as critical"
            )
            subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                timeout=120,
                check=False,
            )
        elif os.name == "nt":
            import ctypes

            # 0x10 = MB_ICONERROR
            ctypes.windll.user32.MessageBoxW(None, message, title, 0x10)
    except (OSError, subprocess.SubprocessError, AttributeError):
        # osascript missing/failing, or no user32 to call. A failed dialog must
        # never mask the error it was trying to report, so this is swallowed —
        # but only for the failures the two branches above can actually raise.
        pass


def fatal_startup_error(message: str, details: list[str] | None = None) -> None:
    """Report a startup failure that will end the process.

    Always prints, so console runs are unchanged. Additionally raises a native
    dialog when the run was going to open a window: a Finder/Explorer launch has
    no attached terminal, so ``error_print`` goes to a stdout nobody sees and the
    user gets a bouncing dock icon and silence. Every hard exit reachable before
    the window exists must route through here.
    """
    error_print(message, details)
    if GUI_LAUNCH:
        body = "\n".join(details or []) or message
        _show_native_alert("clipgen cannot start", body)


def validate_runtime_directories() -> None:
    """Validate that the effective input directory exists and ensure output directory is ready.

    Behavior:
    - If the effective input directory does not exist, print a warning with guidance and exit.
    - If the effective output directory does not exist, attempt to create it and print a message.
    """
    input_dir = get_effective_input_dir()
    if not input_dir.exists():
        fatal_startup_error(
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


def load_json_manifest(
    filename: str, *, default: Any = None, warn_label: str = ""
) -> Any:
    """Load a JSON manifest from the output directory.

    Returns parsed data, or *default* on missing/corrupt file. An unreadable
    (as opposed to missing) file logs a warning when *warn_label* is set,
    mirroring :func:`save_json_manifest`.
    """
    path = Path(get_effective_output_dir()) / filename
    if not path.is_file():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        if warn_label:
            warning_print(f"Could not read {warn_label}; using defaults.")
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


def remove_json_manifest(filename: str) -> None:
    """Delete a manifest and any stale ``.tmp`` sibling from the output dir.

    Used by the save wrappers when a manifest is semantically empty: rather than
    writing an empty artifact into the user's CWD, remove the file so a
    zero-interaction launch leaves no junk. No-op when nothing is on disk.
    """
    path = Path(get_effective_output_dir()) / filename
    for candidate in (path, path.with_suffix(path.suffix + ".tmp")):
        try:
            candidate.unlink(missing_ok=True)
        except OSError:
            pass


def sweep_stale_temp_artifacts() -> None:
    """Remove orphaned atomic-write tmps and reel temp-clips from the output dir.

    Targets only our own artifacts — ``*.json.tmp`` siblings (manifest atomic
    writes) and ``{TEMP_ARTIFACT_PREFIX}*`` reel temp-clips — so user files are
    never touched. Meant to run once at server startup, before any worker thread,
    to reclaim leftovers from a prior hard kill.
    """
    base = Path(get_effective_output_dir())
    if not base.is_dir():
        return
    for pattern in ("*.json.tmp", config.TEMP_ARTIFACT_PREFIX + "*"):
        for stale in base.glob(pattern):
            try:
                stale.unlink(missing_ok=True)
            except OSError:
                pass


def get_bundled_assets_root() -> Path:
    """Return the base directory for bundled project assets.

    In a PyInstaller build, bundled data lives under sys._MEIPASS, which is
    already the bundle root. In source runs this is the *repo* root: `assets/`,
    `build/VERSION` and `CHANGELOG.md` stay there while the Python modules live
    one level down in `source/`, hence the second `.parent`. Every asset lookup
    in the project funnels through here, so this is the only place that has to
    know about that split.
    """
    if getattr(sys, "frozen", False):
        return Path(
            getattr(sys, "_MEIPASS", str(Path(sys.executable).resolve().parent))
        )
    return Path(__file__).resolve().parent.parent


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


def get_licenses_path() -> Path | None:
    """Return the path of the bundled `THIRD-PARTY-LICENSES` notice, or None.

    Same `build/`-vs-bundle-root split as `get_version()`: `clipgen.spec` copies
    the file to the bundle root, while a source checkout keeps it in `build/`.
    """
    root = get_bundled_assets_root()
    for candidate in (
        root / "THIRD-PARTY-LICENSES",
        root / "build" / "THIRD-PARTY-LICENSES",
    ):
        if candidate.is_file():
            return candidate
    return None


def get_licenses_text() -> str | None:
    """Return the bundled `THIRD-PARTY-LICENSES` notice, or None if absent.

    Returns None rather than raising — `--licenses` reports that itself, and no
    other caller should hard-fail on a stripped-down installation.

    Deliberately not cached: this is a ~78 KB read on a path that exits
    immediately afterwards, so caching would only pin the string in memory.
    """
    path = get_licenses_path()
    if path is None:
        return None
    try:
        return path.read_text(encoding="utf-8")
    except OSError:
        return None


def terminate_subprocess(proc: subprocess.Popen, timeout: int = 5) -> None:
    """Terminate a subprocess, escalating to kill if it ignores SIGTERM."""
    proc.terminate()
    try:
        proc.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.wait()


def sanitize_floats(obj: Any) -> Any:
    """Recursively replace non-finite floats (NaN, ±Inf) with ``None`` and
    normalise numpy/Integral scalars to plain Python numbers, for JSON safety.

    Use before any ``jsonify`` response or manifest write that may carry numpy
    or cv2-derived values: NaN survives min/max clamps and serialises as invalid
    ``NaN`` in JSON, and raw numpy scalars are not JSON-serialisable.
    """
    if isinstance(obj, bool):
        return obj
    if isinstance(obj, Integral):
        return int(obj)
    if isinstance(obj, Real):
        value = float(obj)
        return value if math.isfinite(value) else None
    if hasattr(obj, "item") and callable(obj.item):
        try:
            return sanitize_floats(obj.item())
        except (TypeError, ValueError):
            pass
    if isinstance(obj, dict):
        return {k: sanitize_floats(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [sanitize_floats(v) for v in obj]
    return obj


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
    friction_categories = [
        {"key": key, "label": label}
        for key, label in config.FRICTION_CATEGORIES.items()
    ]
    return {
        "defaultDuration": config.DEFAULT_DURATION_SECONDS,
        "severity": severity,
        "annotationKeyphrases": sorted(get_known_annotation_map().keys()),
        "annotations": annotations,
        "ignoredTimestampTokens": sorted(get_ignored_timestamp_tokens()),
        "screenspaceOcrMinConfidence": config.SCREENSPACE_OCR_MIN_CONFIDENCE,
        "screenspaceOcrFuzzyThreshold": config.SCREENSPACE_OCR_FUZZY_THRESHOLD,
        "screenspaceMultitoolMaxOffset": config.SCREENSPACE_MULTITOOL_MAX_OFFSET_SECONDS,
        "screenspaceMaskFallbackTools": list(config.SCREENSPACE_MASK_FALLBACK_TOOLS),
        "frictionCategories": friction_categories,
        "frictionColorToken": "--color-friction",
        "frictionMomentLimit": config.FRICTION_MOMENT_LIMIT,
        "convergenceSources": list(config.CONVERGENCE_SOURCES),
        "cardScrubberSpriteCols": config.STUDIO_SCRUBBER_SPRITE_COLS,
        "cardScrubberSpriteRows": config.STUDIO_SCRUBBER_SPRITE_ROWS,
        "clipFormat": config.FILEFORMAT,
        "screenshotFormat": config.SCREENSHOT_FORMAT,
        "gifFormat": config.GIF_FORMAT,
        "composerAnnotationColor": config.COMPOSER_ANNOTATION_COLOR,
        "composerAnnotationColorSecondary": config.COMPOSER_ANNOTATION_COLOR_SECONDARY,
        "composerAnnotationStrokeWidth": config.COMPOSER_ANNOTATION_STROKE_WIDTH,
        "composerAnnotationStrokeStyle": config.COMPOSER_ANNOTATION_STROKE_STYLE,
        "composerAnnotationFontSize": config.COMPOSER_ANNOTATION_FONT_SIZE,
        "composerAnnotationSpanSeconds": config.COMPOSER_ANNOTATION_SPAN_SECONDS,
        "composerScrubMaxAudioSeconds": config.COMPOSER_SCRUB_MAX_AUDIO_SECONDS,
        "composerDoubleClickCuts": config.COMPOSER_DOUBLE_CLICK_CUTS,
        "mediaContainerWarning": config.MEDIA_CONTAINER_WARNING,
        "subtitleContainers": _subtitle_container_config(),
        "hotkeyOverrides": dict(config.HOTKEY_OVERRIDES),
        "profiling": config.PROFILING,
    }


def _subtitle_container_config() -> dict[str, list[str]]:
    """Container extensions the subtitle muxer can write, split by disposition.

    ``supported`` is every container ``video.mux_subtitles`` has a codec for —
    the Embed Subtitles dialog filters against it so a run that ffmpeg would
    reject at its ``codec is None`` guard is never promised in the summary.
    ``alwaysDefault`` is the mp4 family, whose muxer ships the subtitle track
    enabled no matter what ``-disposition:s:0`` says, so the dialog's
    "set as default" toggle has to declare itself a no-op there.

    Imported lazily: video.py pulls in the ffmpeg helpers, and utils is on the
    import path of everything.
    """
    import video

    return {
        "supported": sorted(video.SUBTITLE_CODEC_BY_CONTAINER),
        "alwaysDefault": sorted(video.SUBTITLE_ALWAYS_DEFAULT_CONTAINERS),
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


def pick_worksheet_title(
    titles: list[str], preferred_name: str | None = None
) -> str | None:
    """Choose a worksheet title from *titles*.

    Precedence: an explicit *preferred_name* if it is present in *titles*, else
    the first ``config.WORKSHEET_PRIORITY`` entry that matches, else the first
    title, else ``None`` when *titles* is empty. Shared by the Google
    (``google_api.get_worksheet``) and Excel (``excel_io``) pickers so the
    auto-selection rule lives in exactly one place.
    """
    if preferred_name and preferred_name in titles:
        return preferred_name
    for priority_name in config.WORKSHEET_PRIORITY:
        if priority_name in titles:
            return priority_name
    if titles:
        return titles[0]
    return None


def _resolve_segment_source_fields(
    clip: "ClipRecord",
    base_video: str,
    start_str: str,
    end_str: str,
    *,
    allow_split: bool,
) -> dict[str, Any]:
    """Resolve the source-video fields for one persisted segment record.

    ``sourceVideo`` is always a **basename** (matching ``pipeline.cut_global_range``);
    regeneration resolves it against the input dir via ``resolve_input_path``.
    Single-video (no ``source_timeline``): ``sourceVideo`` is *base_video*'s
    basename and the local times equal the global times. Multi-video: the global
    ``[start, end]`` is mapped onto ``clip['source_timeline']`` into the owning
    sub-video plus local offsets. When *allow_split* is True (video clips) and the
    range straddles a recording boundary, a ``parts`` list describes each piece so
    it can be re-cut and stitched; ``sourceVideo``/``localStart``/``localEnd``
    carry the first piece. When *allow_split* is False (screenshots/GIFs/
    transcripts) a single frame's position maps by start only — never split.

    ``start``/``end`` (global seconds) stay on the record for the timeline
    viewer; these fields drive regeneration, which re-cuts from ``sourceVideo``.
    """
    global_start = timestamp_to_seconds(start_str) or 0.0
    global_end = timestamp_to_seconds(end_str) or 0.0
    timeline = clip.get("source_timeline")
    if not timeline or len(timeline) < 2:
        return {
            "sourceVideo": Path(base_video).name,
            "localStart": global_start,
            "localEnd": global_end,
        }

    if allow_split:
        pieces = map_global_range_to_segments(timeline, global_start, global_end)
        if pieces:
            parts = [
                {
                    "sourceVideo": Path(timeline[index][0]).name,
                    "localStart": local_start,
                    "localEnd": local_end,
                }
                for index, local_start, local_end in pieces
            ]
            first = parts[0]
            fields: dict[str, Any] = {
                "sourceVideo": first["sourceVideo"],
                "localStart": first["localStart"],
                "localEnd": first["localEnd"],
            }
            if len(parts) > 1:
                fields["parts"] = parts
            return fields
        return {
            "sourceVideo": Path(base_video).name,
            "localStart": global_start,
            "localEnd": global_end,
        }

    mapped = map_global_to_segment(timeline, global_start)
    if mapped is None:
        return {
            "sourceVideo": Path(base_video).name,
            "localStart": global_start,
            "localEnd": global_end,
        }
    index, local_start = mapped
    seg_duration = timeline[index][1]
    local_end = min(float(seg_duration), local_start + (global_end - global_start))
    return {
        "sourceVideo": Path(timeline[index][0]).name,
        "localStart": local_start,
        "localEnd": local_end,
    }


def _clip_metadata_fields(
    clip: "ClipRecord",
    base_video: str,
    start_str: str,
    end_str: str,
    *,
    allow_split: bool = False,
) -> dict[str, Any]:
    """Extract the shared per-segment metadata that every persisted record needs.

    Used by both ``build_artifact_record`` (manifest artifacts) and
    ``build_reel_component`` (reel-component records). The two shapes only differ
    by file-specific fields (id/file/type/thumbnail), so the body of every
    persisted record flows from one place.

    ``start``/``end`` are GLOBAL seconds (the timeline viewer positions artifacts
    by them). ``sourceVideo``/``localStart``/``localEnd`` (and ``parts`` for a
    boundary-spanning clip) describe where the segment was actually cut from and
    drive regeneration — see :func:`_resolve_segment_source_fields`.
    """
    cell = clip.get("cell")
    cell_row = getattr(cell, "row", None)
    cell_col = getattr(cell, "col", None)
    fields: dict[str, Any] = {
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
    }
    fields.update(
        _resolve_segment_source_fields(
            clip, base_video, start_str, end_str, allow_split=allow_split
        )
    )
    return fields


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

    The artifact id is built from ``cell.row`` / ``cell.col`` and is the manifest
    dedup key. Callers must therefore provide either a real spreadsheet cell
    (positive row/col) or a synthetic cell with a unique ``(row, col)`` pair —
    see ``_make_synthetic_clip_record`` in ``cli.py``, which mints negative
    rows namespaced per-mode by ``cell_col``. Passing ``cell=None`` or a stub
    without ``.row``/``.col`` raises ``ValueError`` to prevent silent id
    collisions (two such records with the same ``seg_idx`` would dedup against
    each other in ``viewer.save_manifest``).
    """
    cell = clip.get("cell")
    cell_row = getattr(cell, "row", None)
    cell_col = getattr(cell, "col", None)
    if cell_row is None or cell_col is None:
        raise ValueError(
            "build_artifact_record requires a cell with row and col; "
            "synthetic records must use a unique (row, col) pair — see "
            "_make_synthetic_clip_record in cli.py for the negative-row "
            "convention."
        )
    return {
        "id": f"a{cell_row}c{cell_col}s{seg_idx}",
        "type": artifact_type,
        "file": Path(out_path).name,
        "thumbnail": "",
        # Only video clips may be stitched across a recording boundary; a single
        # screenshot/GIF frame maps by its start position and is never split.
        **_clip_metadata_fields(
            clip, base_video, start_str, end_str, allow_split=(artifact_type == "clip")
        ),
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
    return _clip_metadata_fields(clip, base_video, start_str, end_str, allow_split=True)


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


def split_selector_tokens(text: str) -> list[str]:
    """Split a selector string on ',' or '+' into stripped, non-empty tokens.

    Used by the CLI category/severity/keyword/line selectors and the spreadsheet
    participant/cell selectors, which all accept comma- or plus-separated lists.
    """
    if not text:
        return []
    return [tok.strip() for tok in text.replace(",", "+").split("+") if tok.strip()]


def _clean_timestamp_token(token: str) -> str:
    """Normalize one token before timestamp parsing."""
    return token.strip().rstrip(",").rstrip("-").replace(".", ":")


@functools.cache
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


def _parse_single_timestamp_token(token: str) -> tuple[str, str] | None:
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
    # treat a leading dash (e.g. "-5") or non-time dash as a range, and require
    # both halves to be real timestamps so half-garbage ranges ("1:23-abc",
    # "5-10", "1:23-1:45-2:00") are reported as skipped here instead of
    # failing deep in ffmpeg with the clip silently dropped.
    if "-" in token:
        dash_pos = token.find("-")
        if dash_pos > 0 and token[dash_pos - 1].isdigit():
            start_time = token[:dash_pos]
            end_time = token[dash_pos + 1 :]
            if (
                timestamp_to_seconds(start_time) is not None
                and timestamp_to_seconds(end_time) is not None
            ):
                return (start_time, end_time)
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

    parts = ts.split(":")
    if len(parts) not in (2, 3) or not all(p.isdigit() for p in parts):
        return None
    nums = [int(p) for p in parts]

    if len(parts) == 3:
        hours, minutes, seconds = nums
    else:
        # MM:SS — minutes may exceed 59 (e.g. "75:00" for a long session
        # written without an hours component), matching how the range form
        # "75:00-80:00" already flows through to ffmpeg.
        hours = 0
        minutes, seconds = nums

    minutes_per_hour = config.SECONDS_PER_HOUR // config.SECONDS_PER_MINUTE
    if seconds >= config.SECONDS_PER_MINUTE:
        return None
    if len(parts) == 3 and minutes >= minutes_per_hour:
        return None
    return float(
        hours * config.SECONDS_PER_HOUR + minutes * config.SECONDS_PER_MINUTE + seconds
    )


def parse_timestamps(
    cell_value: str, cell_ref: str | None = None
) -> list[tuple[str, str]]:
    """Parse timestamp pairs from a cell value string.

    Supported formats: "MM:SS-MM:SS", "HH:MM:SS-HH:MM:SS", or a single
    "MM:SS"/"HH:MM:SS" whose end becomes start + the default duration. Multiple
    pairs may be separated by space, comma, semicolon or plus.

    *cell_ref* (e.g. 'B5') only labels error messages. Invalid tokens are skipped
    and reported via warning_print rather than raising.
    """
    if config.DEBUGGING:
        config.debug_ic(cell_value, cell_ref)
    parsed_timestamps = []
    skipped_timestamps = []
    ignored_tokens = get_ignored_timestamp_tokens()
    # Unify delimiters (+, ;, ,) to spaces so split() yields one token per time or range
    raw_times = _split_timestamp_tokens(cell_value)
    if config.DEBUGGING:
        config.debug_ic(raw_times)
        debug_print(f"raw_times content after split is {raw_times}")
        debug_print(f"Timestamp list raw_times is {len(raw_times)} entries long")

    # Clean each token (strip, normalize trailing punctuation, use colon for decimals) and parse
    raw_times = [_clean_timestamp_token(t) for t in raw_times]
    for token in raw_times:
        if config.DEBUGGING:
            debug_print(f"Cleaning timestamp {token}")
        pair = _parse_single_timestamp_token(token)
        if pair is not None:
            if config.DEBUGGING and len(pair) == 2:
                config.debug_ic(pair)
            parsed_timestamps.append(pair)
        elif token and token not in ignored_tokens:
            skipped_timestamps.append(token)

    # Report skipped timestamps: list up to MAX_SKIPPED_TIMESTAMPS_TO_SHOW, then "... and N more"
    if skipped_timestamps:
        if config.DEBUGGING:
            config.debug_ic(skipped_timestamps)
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
        config.debug_ic(parsed_timestamps)
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


def seconds_to_timestamp(total_seconds: float, *, force_hours: bool = False) -> str:
    """Format a non-negative number of seconds as H:MM:SS or M:SS.

    Accepts an int or float; a float (e.g. a clamped end time) is truncated to
    whole seconds so the ``:d``/``:02d`` format specs can't crash.
    """
    total_seconds = int(total_seconds)
    total_seconds = max(total_seconds, 0)
    hours, rem = divmod(total_seconds, config.SECONDS_PER_HOUR)
    minutes, seconds = divmod(rem, config.SECONDS_PER_MINUTE)
    if hours > 0 or force_hours:
        return f"{hours:d}:{minutes:02d}:{seconds:02d}"
    return f"{minutes:d}:{seconds:02d}"


# ---- Multi-video timeline mapping ----
#
# When a participant's session spans several source videos (see
# video.build_source_timeline), spreadsheet timestamps are GLOBAL — relative to
# the concatenated recording. These pure helpers map a global second into the
# owning sub-video and the local offset within it. The ``timeline`` argument is
# the list of ``(path, duration, cumulative_start)`` tuples returned by
# build_source_timeline.


def _timeline_total_seconds(timeline: list[tuple[str, int, int]]) -> int:
    """Total duration of the concatenated timeline (last cumulative_start + duration)."""
    if not timeline:
        return 0
    _path, duration, cumulative = timeline[-1]
    return cumulative + duration


def map_global_to_segment(
    timeline: list[tuple[str, int, int]], global_seconds: float
) -> tuple[int, float] | None:
    """Map a global second to ``(segment_index, local_seconds)``.

    Returns the segment whose range ``[cumulative_start, cumulative_start +
    duration)`` contains *global_seconds*, with the offset within that segment.
    Returns ``None`` if *global_seconds* is negative or at/beyond the total
    timeline duration. Example: timeline video1=80s, video2=120s →
    ``map_global_to_segment(t, 124)`` is ``(1, 44.0)``.
    """
    if global_seconds < 0:
        return None
    if global_seconds >= _timeline_total_seconds(timeline):
        return None
    for index, (_path, duration, cumulative) in enumerate(timeline):
        if cumulative <= global_seconds < cumulative + duration:
            return (index, global_seconds - cumulative)
    return None


def resolve_timeline_segment(
    timeline: list[tuple[str, int, int]], global_seconds: float
) -> tuple[str, float] | None:
    """Map *global_seconds* to ``(sub_video_path, local_seconds)`` within a built
    source *timeline*, or ``None`` if it falls at/beyond the recording.

    Companion to :func:`map_global_to_segment` that also resolves the owning
    sub-video's path — the shared step behind multi-video thumbnail and frame
    lookups. Callers that already hold a (cached) timeline pass it straight in.
    """
    mapped = map_global_to_segment(timeline, global_seconds)
    if mapped is None:
        return None
    index, local_seconds = mapped
    return (timeline[index][0], local_seconds)


def map_global_range_to_segments(
    timeline: list[tuple[str, int, int]],
    start_seconds: float,
    end_seconds: float,
) -> list[tuple[int, float, float]] | None:
    """Map a global ``[start, end)`` range to per-segment ``(index, local_start, local_end)`` pieces.

    A range that lies within one sub-video yields a single piece; a range that
    straddles a recording boundary yields one piece per spanned sub-video, each
    clamped to that sub-video's ``[0, duration]``. An ``end`` past the timeline
    is clamped to the total duration. Returns ``None`` if ``end <= start`` or the
    start is out of range. Boundary spans are detected by ``len(result) > 1``.
    """
    if end_seconds <= start_seconds:
        return None
    total = _timeline_total_seconds(timeline)
    if start_seconds < 0 or start_seconds >= total:
        return None
    end_seconds = min(end_seconds, float(total))
    pieces: list[tuple[int, float, float]] = []
    for index, (_path, duration, cumulative) in enumerate(timeline):
        seg_end = cumulative + duration
        overlap_start = max(start_seconds, float(cumulative))
        overlap_end = min(end_seconds, float(seg_end))
        if overlap_end > overlap_start:
            pieces.append((index, overlap_start - cumulative, overlap_end - cumulative))
    return pieces or None


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
            cur_end = max(cur_end, e)
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


def apply_span_padding(
    start: float,
    end: float,
    *,
    pad_pre: float = 0.0,
    pad_post: float = 0.0,
    max_duration: float = 0.0,
    limit: float | None = None,
) -> tuple[float, float]:
    """Pad a single (start, end) second-span and clamp it to a single shorter span.

    Unlike cluster_spans (which splits an over-long span into back-to-back
    sub-clips), this shortens the end so the result stays one span — the
    behavior the Workflows artifact nodes want.

    - Pads are signed: positive extends outward, negative trims inward
      (start += -pad_pre, end += pad_post).
    - When ``limit`` is given (e.g. the source video's EOF in seconds), the end
      is capped there and the start kept at most ``limit - 1`` so a valid span
      survives — avoids run_ffmpeg skipping a clip padded past end-of-video.
    - When ``max_duration > 0`` and the padded span still exceeds it, the end is
      pulled in to ``start + max_duration``.
    - start is floored at 0; end is floored at start + 1s (min 1-second span).
      This is applied last, so it wins over a sub-1s ``max_duration`` cap — the
      result is never shorter than a second.
    """
    s = max(0.0, start - pad_pre)
    e = end + pad_post
    if limit is not None:
        e = min(e, limit)
        s = min(s, max(0.0, limit - 1))
    if max_duration > 0 and (e - s) > max_duration:
        e = s + max_duration
    e = max(e, s + 1)
    return s, e


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
    # Keep dividing by 1024 until size is under 1024 or we reach TB (index 4).
    # Use >= so an exact power of 1024 promotes to the next unit (1024 bytes ->
    # 1.00KB, not 1024.00B).
    while size_bytes >= 1024 and suffix_index < 4:
        suffix_index += 1
        size_bytes = size_bytes / 1024
    return f"{size_bytes:.{precision}f}{suffixes[suffix_index]}"


def get_current_time() -> str:
    """Get current time as formatted string.

    Returns:
        Current time in format 'YYYY-MM-DD HH:MM:SS'
    """
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ISO 639-1 (two-letter) -> ISO 639-2/T (three-letter), covering every language
# faster-whisper can report (its _LANGUAGE_CODES set). Media containers store
# only the three-letter form, so this is the hop between what Whisper detects
# and what a subtitle/audio track can be tagged with.
#
# Whisper's own set already includes two three-letter codes ("haw", "yue"),
# which need no entry — normalize_track_language passes any 3-letter code
# through. Deliberately the /T (terminological) variant: for the ~20 languages
# where ISO 639-2 has two codes, /T is what Matroska specifies and what current
# players expect ("deu" not "ger", "fra" not "fre"). ffmpeg stores whichever it
# is given verbatim, so the choice is ours to get right.
ISO639_1_TO_2 = {
    "af": "afr", "am": "amh", "ar": "ara", "as": "asm", "az": "aze",
    "ba": "bak", "be": "bel", "bg": "bul", "bn": "ben", "bo": "bod",
    "br": "bre", "bs": "bos", "ca": "cat", "cs": "ces", "cy": "cym",
    "da": "dan", "de": "deu", "el": "ell", "en": "eng", "es": "spa",
    "et": "est", "eu": "eus", "fa": "fas", "fi": "fin", "fo": "fao",
    "fr": "fra", "gl": "glg", "gu": "guj", "ha": "hau", "he": "heb",
    "hi": "hin", "hr": "hrv", "ht": "hat", "hu": "hun", "hy": "hye",
    "id": "ind", "is": "isl", "it": "ita", "ja": "jpn", "jv": "jav",
    # Whisper spells Javanese "jw"; ISO 639-1 spells it "jv". Both map over.
    "jw": "jav",
    "ka": "kat", "kk": "kaz", "km": "khm", "kn": "kan", "ko": "kor",
    "la": "lat", "lb": "ltz", "ln": "lin", "lo": "lao", "lt": "lit",
    "lv": "lav", "mg": "mlg", "mi": "mri", "mk": "mkd", "ml": "mal",
    "mn": "mon", "mr": "mar", "ms": "msa", "mt": "mlt", "my": "mya",
    "ne": "nep", "nl": "nld", "nn": "nno", "no": "nor", "oc": "oci",
    "pa": "pan", "pl": "pol", "ps": "pus", "pt": "por", "ro": "ron",
    "ru": "rus", "sa": "san", "sd": "snd", "si": "sin", "sk": "slk",
    "sl": "slv", "sn": "sna", "so": "som", "sq": "sqi", "sr": "srp",
    "su": "sun", "sv": "swe", "sw": "swa", "ta": "tam", "te": "tel",
    "tg": "tgk", "th": "tha", "tk": "tuk", "tl": "tgl", "tr": "tur",
    "tt": "tat", "uk": "ukr", "ur": "urd", "uz": "uzb", "vi": "vie",
    "yi": "yid", "yo": "yor", "zh": "zho",
}  # fmt: skip


def normalize_track_language(code: str | None) -> str:
    """Coerce *code* into a three-letter media-container language tag.

    Containers store only ISO 639-2, and they fail *silently* on anything else:
    measured on ffmpeg 8.1.2, an ``en`` tag is dropped outright by the mp4
    muxer, while ``unknown`` (transcripts.py's detection fallback) is truncated
    to the nonsense tag ``unk``. Neither raises, so an unnormalized code costs
    the track its language with no error anywhere.

    Returns ``"und"`` (the standard "undetermined" tag) for empty, unknown, or
    malformed input rather than passing junk through to the muxer.
    """
    text = (code or "").strip().lower()
    if not text:
        return "und"
    # Accept BCP 47 forms too ("en-US", "zh_Hans"): only the primary subtag is
    # meaningful to a container. TRANSCRIBE_LANGUAGE is user-editable, so this
    # is a plausible thing to be handed.
    primary = text.replace("_", "-").split("-")[0]
    if len(primary) == 3 and primary.isalpha():
        return primary
    return ISO639_1_TO_2.get(primary, "und")


# ---- Participant video discovery ----


def numbered_parts_are_contiguous(indices: list[int]) -> bool:
    """True if *indices* are exactly ``1..N`` with no gaps (after sorting).

    Guards the multi-video timeline: numbered source parts must be a gapless
    sequence starting at 1, otherwise concatenating them back-to-back would map
    global timestamps into the wrong sub-video.
    """
    return sorted(indices) == list(range(1, len(indices) + 1))


def split_source_stem(name: str) -> tuple[str, str]:
    """Split a source-video filename into ``(study, remainder)``.

    Strips a numbered-part ``-N`` suffix first so ``study_P01-2.mp4`` yields
    ``("study", "P01")``. When the stem has no ``_``, study is ``""`` and
    remainder is the stripped stem.
    """
    stem = Path(name).stem
    head, sep, tail = stem.rpartition("-")
    if sep and head and tail.isdigit():
        stem = head
    parts = stem.rsplit("_", 1)
    if len(parts) == 2:
        return (parts[0], parts[1])
    return ("", stem)


def participant_id_from_source_name(name: str) -> str | None:
    """Extract the participant id from a source-video filename, or None.

    Handles both the plain ``{study}_{participant}{FILEFORMAT}`` form and a
    numbered part ``{study}_{participant}-N{FILEFORMAT}`` — the ``-N`` suffix is
    stripped first so a part groups under its base participant id. Returns None
    when the trailing segment is not a recognised id (config.PARTICIPANT_PREFIXES).
    """
    study, pid = split_source_stem(name)
    if not pid:
        return None
    # A stem with no ``_`` maps to ``("", stripped_stem)``. That is not a
    # participant id — unlike ``_P01.mp4``, which splits to ``("", "P01")``.
    if not study and "_" not in Path(name).stem:
        return None
    if pid[0] not in config.PARTICIPANT_PREFIXES:
        return None
    # Reject ids with whitespace: real participant ids are clean tokens (P01,
    # G02). A space is the signature of a Finder/Explorer duplicate
    # ("study_P03 copy.mp4"), which is never a new participant — without this it
    # becomes a phantom participant in every tool's dropdown and (with the P6
    # watch-dir trigger) auto-launches a run for a bogus id.
    if any(ch.isspace() for ch in pid):
        return None
    return pid


# One participant-video scan per input-dir state, keyed dir -> (mtime_ns, result).
# The directory mtime advances on add/remove/rename, invalidating on real change
# (incl. the P6 watch-dir drop). Result is directory-only (study_name is unused),
# so it is shared across all callers regardless of the study_name they pass, and
# keying on the dir string (not a single slot) means a runtime input-dir switch
# selects a different entry rather than needing explicit invalidation.
_discover_videos_cache: dict[str, tuple[int | None, list[dict[str, Any]]]] = {}
_discover_videos_lock = threading.Lock()


def discover_participant_videos(study_name: str = "") -> list[dict[str, Any]]:
    """Scan the input directory and return one entry per participant.

    A participant's session may span several files (a recording that broke off,
    or a diary study); this groups the plain ``{study}_{pid}{FILEFORMAT}`` and/or
    the numbered parts ``{study}_{pid}-N{FILEFORMAT}`` into one entry with ordered
    ``video_paths``. The plain file wins when both it and numbered parts exist; a
    non-contiguous numbered set is skipped with a warning (see
    :func:`numbered_parts_are_contiguous`). Only ids starting with a recognised
    prefix (``config.PARTICIPANT_PREFIXES``) are included.

    Cached on the input directory's ``mtime_ns``, so the many hot callers
    (``/api/status``, the Workflows video-source node, the watch-dir daemon) share
    one glob/parse pass until the directory changes.

    Returns:
        ``{"id", "video_paths", "has_video"}`` dicts, sorted by participant id.
    """
    input_dir = Path(get_effective_input_dir())
    dir_str = str(input_dir)
    try:
        mtime_ns: int | None = (
            input_dir.stat().st_mtime_ns if input_dir.is_dir() else None
        )
    except OSError:
        mtime_ns = None

    with _discover_videos_lock:
        cached = _discover_videos_cache.get(dir_str)
        if cached is not None and cached[0] == mtime_ns:
            return cached[1]

        plain: dict[str, Path] = {}
        numbered: dict[str, list[tuple[int, Path]]] = {}
        if input_dir.is_dir():
            for path in sorted(input_dir.glob(f"*{config.FILEFORMAT}")):
                pid = participant_id_from_source_name(path.name)
                if pid is None:
                    continue
                head, sep, tail = path.stem.rpartition("-")
                if sep and head and tail.isdigit():
                    numbered.setdefault(pid, []).append((int(tail), path))
                else:
                    plain[pid] = path

        participants: list[dict[str, Any]] = []
        for pid in sorted(set(plain) | set(numbered)):
            if pid in plain:
                paths = [plain[pid]]
            else:
                parts = sorted(numbered[pid], key=lambda item: item[0])
                indices = [n for n, _ in parts]
                if not numbered_parts_are_contiguous(indices):
                    warning_print(
                        f"Numbered source videos for participant '{pid}' are "
                        f"non-contiguous (found parts {indices}); expected 1..N.",
                        [
                            (
                                "Skipping this participant; rename the parts to a "
                                "gapless 1..N sequence to enable concatenation."
                            ),
                        ],
                    )
                    continue
                paths = [p for _, p in parts]
            participants.append(
                {
                    "id": pid,
                    "video_paths": [str(p) for p in paths],
                    "has_video": paths[0].is_file(),
                }
            )

        _discover_videos_cache[dir_str] = (mtime_ns, participants)
        return participants


# ---- Flask blueprint helpers ----

# Live index pages embed this marker where the shared favicon + Google-fonts
# block belongs; render_index_html() expands it from assets/web/_head.html so
# the block lives in one place. Exported viewers don't use it (self-contained).
_HEAD_MARKER = "<!-- CLIPGEN_HEAD_HERE -->"

# Rendered-index cache: str(index path) -> (index_mtime_ns, head_mtime_ns|None,
# desktop_chrome, rendered). The `/` route re-renders on every GET; the assets never
# change while the server runs, so memoize by mtime (a live dev edit still bumps mtime
# and refreshes). head_mtime is None for pages without the marker (they never read
# _head.html). DESKTOP_CHROME is part of the key because it varies per launch, not per
# file, so an mtime-only key would serve a browser render into a desktop window.
_index_html_cache: dict[str, tuple[int | None, int | None, str | None, str]] = {}
_index_html_lock = threading.Lock()


def _desktop_chrome_head(chrome: str) -> str:
    """Inline ``<script>`` telling the page it is hosted in a native window.

    Runs in ``<head>``, so it lands before the deferred ``topnav.js`` reads the
    attribute — the bar lays out inset for the traffic lights on first paint rather
    than jumping. The two measurements come from config so AppKit (which positions
    the real buttons) and CSS (which reserves the space) cannot drift apart.
    """
    return (
        "\n  <script>(function () {\n"
        "    var d = document.documentElement;\n"
        f'    d.dataset.desktopChrome = "{chrome}";\n'
        f'    d.style.setProperty("--desktop-chrome-height", "{config.DESKTOP_CHROME_BAR_HEIGHT}px");\n'
        f'    d.style.setProperty("--desktop-traffic-inset", "{config.DESKTOP_TRAFFIC_LIGHT_INSET}px");\n'
        "  })();</script>"
    )


def render_index_html(assets_dir: Path, index_html: str) -> str:
    """Read an index page, expanding the shared ``<head>`` marker if present.

    Pages without the marker are returned unchanged, so this stays safe for any
    current or future index page. Results are memoized per index path and
    invalidated when the page (or, for marker pages, ``_head.html``) mtime changes.
    """
    index_path = assets_dir / index_html
    head_path = assets_dir / "_head.html"
    chrome = DESKTOP_CHROME
    try:
        index_mtime: int | None = index_path.stat().st_mtime_ns
    except OSError:
        index_mtime = None

    with _index_html_lock:
        cached = _index_html_cache.get(str(index_path))
        if cached is not None and cached[0] == index_mtime and cached[2] == chrome:
            head_mtime_cached = cached[1]
            if head_mtime_cached is None:
                return cached[3]
            try:
                head_mtime: int | None = head_path.stat().st_mtime_ns
            except OSError:
                head_mtime = None
            if head_mtime == head_mtime_cached:
                return cached[3]

        html = index_path.read_text(encoding="utf-8")
        head_mtime_used: int | None = None
        if _HEAD_MARKER in html:
            head = head_path.read_text(encoding="utf-8").rstrip("\n")
            if chrome:
                head += _desktop_chrome_head(chrome)
            html = html.replace(_HEAD_MARKER, head)
            try:
                head_mtime_used = head_path.stat().st_mtime_ns
            except OSError:
                head_mtime_used = None
        _index_html_cache[str(index_path)] = (
            index_mtime,
            head_mtime_used,
            chrome,
            html,
        )
        return html


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
        return Response(render_index_html(assets_dir, index_html), mimetype="text/html")

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
