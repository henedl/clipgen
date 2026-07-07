# -*- coding: utf-8 -*-
"""File and filename operations for clipgen."""

import os
import re
import threading
import unicodedata
from pathlib import Path

import gspread

import config
import utils
from utils import ClipRecord

# Per-template high-water counter so a batch of same-named reservations doesn't
# re-probe 0,1,2,… from scratch each call (which is O(n²) syscalls). Keyed on
# (directory, base, extension); only ever advances, so released slots leave
# harmless gaps but uniqueness is preserved. Guarded because get_unique_filename
# runs inside the clip ThreadPoolExecutor.
_unique_high_water: dict[tuple[str, str, str], int] = {}
_unique_high_water_lock = threading.Lock()


def safe_truncate(text: str, max_chars: int) -> str:
    """Truncate to ``max_chars`` code points, then drop trailing combining marks.

    Plain ``text[:n]`` can split a grapheme cluster (e.g. emoji + skin-tone
    modifier, or letter + combining accent), leaving an orphan combining
    character. Drop those at the truncation boundary so the visible glyph
    survives intact.
    """
    if max_chars <= 0:
        return ""
    truncated = text[:max_chars]
    # Trim trailing combining marks / joiners that would render orphaned.
    while truncated and (
        unicodedata.category(truncated[-1]) in ("Mn", "Mc", "Me")
        or truncated[-1] in ("‍", "️")
    ):
        truncated = truncated[:-1]
    return truncated


def get_unique_filename(filename: str, file_format: str | None = None) -> str:
    """Atomically reserve a unique output path.

    Creates an empty placeholder file at the returned path so that parallel
    clip / reel / gallery workers each receive a distinct path even when their
    filename templates collide. The placeholder is overwritten by the ffmpeg
    process (or file writer) that fills the artifact. Callers that abort before
    writing real content should remove the placeholder with
    ``release_reservation()``.

    If a file with the given name already exists, appends '-1', '-2', etc.
    until a free name is found. Also truncates if filename exceeds max length.

    Args:
        filename: Original filename
        file_format: File extension to preserve (defaults to config.FILEFORMAT)

    Returns:
        Unique filename path as a string, reserved on disk as an empty file.
    """
    file_extension = file_format or config.FILEFORMAT
    resolved = Path(utils.resolve_output_path(filename))
    directory = resolved.parent
    name = resolved.name
    # Strip extension to get base name
    if name.endswith(file_extension):
        base = name[: -len(file_extension)]
    else:
        base = name
    # Truncate base if needed (reserve space for extension)
    max_base = config.MAX_FILENAME_LENGTH - len(file_extension)
    base = safe_truncate(base, max_base)

    def _candidate(counter: int) -> Path:
        if counter == 0:
            return directory / (base + file_extension)
        suffix = f"-{counter}"
        truncated_base = safe_truncate(base, max_base - len(suffix))
        return directory / (truncated_base + suffix + file_extension)

    key = (str(directory), base, file_extension)
    with _unique_high_water_lock:
        counter = _unique_high_water.get(key, 0)

    def _record_high_water(used: int) -> None:
        with _unique_high_water_lock:
            _unique_high_water[key] = max(_unique_high_water.get(key, 0), used + 1)

    while True:
        candidate = _candidate(counter)
        try:
            fd = os.open(candidate, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o644)
        except FileExistsError:
            counter += 1
            continue
        except OSError:
            # Reservation not possible (missing/unwritable directory, name too
            # long, etc.). Fall back to non-atomic uniqueness so callers still
            # receive a usable path.
            while candidate.is_file():
                counter += 1
                candidate = _candidate(counter)
            _record_high_water(counter)
            return str(candidate)
        else:
            os.close(fd)
            _record_high_water(counter)
            return str(candidate)


def release_reservation(path: str | os.PathLike[str] | None) -> None:
    """Remove an unused placeholder reserved by ``get_unique_filename()``.

    Safe to call when *path* is missing or ``None``; used by callers that abort
    before writing real content so empty placeholders are not left on disk.
    """
    if not path:
        return
    try:
        Path(path).unlink(missing_ok=True)
    except OSError:
        pass


def _apply_default_extension(name: str) -> str:
    """Return *name* unchanged if it has an extension, else append FILEFORMAT."""
    return name if Path(name).suffix else name + config.FILEFORMAT


def get_source_video_filename(
    study: str, participant: str, override: str | None = None
) -> str:
    """Resolve the expected source video filename for a clip.

    Args:
        study: Normalized study name
        participant: Normalized participant ID
        override: Optional filename override from the spreadsheet

    Returns:
        Filename to use when looking for the source video.
    """
    if override is not None:
        override = override.strip()
    if override:
        # If override includes an extension, respect it as-is; else append default.
        return _apply_default_extension(override)
    return f"{study}_{participant}{config.FILEFORMAT}"


def get_source_video_filenames(
    study: str, participant: str, override: str | None = None
) -> list[str]:
    """Resolve the ordered list of expected source video filenames for a clip.

    Supports multiple source videos per participant that form one continuous
    concatenated timeline (e.g. a recording that broke off, or a diary study).

    - With an *override* (the spreadsheet ``Filename`` row): split on ``+`` so a
      cell like ``"morning.mp4 + afternoon.mp4"`` yields two files in order. Each
      part follows the same extension rule as :func:`get_source_video_filename`.
      Empty parts are dropped. Order is authoritative (concatenation order).
    - Without an override: returns the single plain name
      ``{study}_{participant}.mp4``. On-disk numbered-suffix auto-detection
      (``study_P01-1.mp4`` ...) is resolved later by the pipeline against the
      input directory via :func:`discover_numbered_source_videos`.

    Always returns at least one entry.
    """
    if override is not None:
        override = override.strip()
    if override:
        parts = [p.strip() for p in override.split("+")]
        names = [_apply_default_extension(p) for p in parts if p]
        if names:
            return names
    return [f"{study}_{participant}{config.FILEFORMAT}"]


def resolve_source_video_paths(
    study: str, participant: str, override: str | None, input_dir: Path
) -> list[Path]:
    """Resolve a participant's ordered source-video paths without fuzzy matching.

    Resolution order (same policy as the pipeline's interactive resolver, minus
    the fuzzy-match prompt — for non-interactive callers like the studio server
    and the ``--transcribe`` CLI):

    - *override* present (the spreadsheet ``Filename`` row) → its plus-separated
      parts, resolved in order.
    - else the plain ``{study}_{participant}.mp4`` when it exists on disk.
    - else on-disk numbered parts (``study_P01-1.mp4`` ...) when any exist.
    - else the (missing) plain path, so callers can report it via ``is_file()``.

    Always returns at least one path.
    """
    names = get_source_video_filenames(study, participant, override)

    def _against_input_dir(name: str) -> Path:
        path = Path(name)
        return path if path.is_absolute() else input_dir / path

    if override:
        return [_against_input_dir(n) for n in names]
    plain = _against_input_dir(names[0])
    if plain.is_file():
        return [plain]
    numbered = discover_numbered_source_videos(input_dir, study, participant)
    if numbered:
        return numbered
    return [plain]


def discover_numbered_source_videos(
    input_dir: Path, study: str, participant: str
) -> list[Path]:
    """Return numbered source-video parts for a participant, ordered by part number.

    Globs *input_dir* for ``{study}_{participant}-N{FILEFORMAT}`` files and sorts
    them by the integer N (so ``-2`` precedes ``-10``). Used only when no
    spreadsheet override is set and the plain ``{study}_{participant}{FILEFORMAT}``
    file is absent. Returns [] when no numbered parts exist.

    The parts must form a gapless ``1..N`` sequence: concatenating non-contiguous
    parts back-to-back would map global timestamps into the wrong sub-video. When
    a gap is found, a warning is emitted and [] is returned (treated as no valid
    multi-part sequence) rather than building a silently-wrong timeline.
    """
    prefix = f"{study}_{participant}"
    matches: list[tuple[int, Path]] = []
    for p in input_dir.glob(f"{prefix}-*{config.FILEFORMAT}"):
        m = re.search(
            config.NUMBERED_SOURCE_VIDEO_SUFFIX_PATTERN, p.name, re.IGNORECASE
        )
        if m:
            matches.append((int(m.group(1)), p))
    matches.sort(key=lambda item: item[0])
    indices = [n for n, _ in matches]
    if indices and not utils.numbered_parts_are_contiguous(indices):
        utils.warning_print(
            f"Numbered source videos for '{prefix}' are non-contiguous "
            f"(found parts {indices}); expected 1..N with no gaps.",
            [
                "Ignoring the numbered sequence; rename the parts to a gapless "
                "1..N sequence to enable concatenation.",
            ],
        )
        return []
    return [p for _, p in matches]


def prepare_clip(clip: ClipRecord) -> ClipRecord:
    """Parse timestamps and sanitize description/category for filename use.

    Mutates the input clip dict: adds 'times' (list of (start, end) timestamp pairs)
    and overwrites 'desc' and 'category' with sanitized values.

    Expected input keys: 'cell', 'desc', 'category', 'study', 'participant'.

    Args:
        clip: Clip record dict containing 'cell', 'desc', 'category', 'study', 'participant'

    Returns:
        The same dict with 'times' added and sanitized 'desc' and 'category'
    """
    if config.DEBUGGING:
        config.debug_ic(clip)

    # Pre-parsed fast path: callers (e.g. --ss-clips, --transcript-clips) build
    # synthetic clips with timestamps already resolved. Skip the cell-based parse
    # but still sanitize desc/category for use in filenames.
    if clip.get("times"):
        clip["cell_annotations"] = list(clip.get("cell_annotations") or [])
        clip["segment_annotations"] = dict(clip.get("segment_annotations") or {})
        raw_desc = clip.get("desc", "")
        bracket_pos = raw_desc.rfind("]")
        cleaned_desc = (
            raw_desc[bracket_pos + 1 :].strip()
            if bracket_pos >= 0
            else raw_desc.strip()
        )
        clip["desc"] = utils.sanitize_filename(cleaned_desc)
        clip["category"] = (
            utils.sanitize_filename(clip["category"])
            if clip.get("category")
            else "uncategorized"
        )
        return clip

    utils.debug_print(
        f"prepare_clip() received clip with cell contents {clip['cell'].value}"
    )
    utils.debug_print("Will attempt to split the cell contents")

    # Get cell reference for error messages
    cell_ref = gspread.utils.rowcol_to_a1(clip["cell"].row, clip["cell"].col)

    # Parse inline annotations (e.g. !key), then parse timestamps from cleaned value.
    cleaned_cell_value, segment_annotations, cell_annotations = (
        utils.parse_cell_annotations(clip["cell"].value)
    )
    clip["cell_annotations"] = sorted(cell_annotations)
    clip["segment_annotations"] = {
        key: sorted(indexes) for key, indexes in segment_annotations.items()
    }
    clip["times"] = utils.parse_timestamps(cleaned_cell_value, cell_ref=cell_ref)
    timestamp_baseline = clip.get("timestamp_baseline")
    if timestamp_baseline:
        clip["times"] = utils.convert_clock_pairs_to_relative(
            clip["times"], timestamp_baseline, cell_ref=cell_ref
        )
    selected_segment_indexes = clip.get("selected_segment_indexes")
    if selected_segment_indexes is not None:
        selected_set = set(selected_segment_indexes)
        clip["times"] = [
            pair for index, pair in enumerate(clip["times"]) if index in selected_set
        ]
    if config.DEBUGGING:
        config.debug_ic(clip["times"])

    # Warn if no valid timestamps were parsed, except cells with only ignored tokens (e.g. "x").
    if not clip["times"] and utils.has_non_ignored_timestamp_content(
        cleaned_cell_value
    ):
        # Only show this detailed per-cell warning at verbose verbosity.
        if getattr(config, "VERBOSITY", config.STANDARD) >= config.VERBOSE:
            utils.warning_print(
                f"No valid timestamps found in cell {cell_ref}",
                [
                    f"Cell contents: '{clip['cell'].value}'",
                    f"Participant: {clip['participant']}, Description: {clip['desc'][:50]}...",
                ],
            )

    # Clean description: remove bracketed prefix and sanitize for use in filename
    raw_desc = clip["desc"]
    bracket_pos = raw_desc.rfind("]")
    cleaned_desc = (
        raw_desc[bracket_pos + 1 :].strip() if bracket_pos >= 0 else raw_desc.strip()
    )
    clip["desc"] = utils.sanitize_filename(cleaned_desc)
    if config.DEBUGGING:
        config.debug_ic(clip["desc"])

    # Sanitize category (handle None/empty)
    if clip["category"]:
        clip["category"] = utils.sanitize_filename(clip["category"])
    else:
        clip["category"] = "uncategorized"
    if config.DEBUGGING:
        config.debug_ic(clip["category"])
        config.debug_ic(clip)
    return clip


_WORKFLOW_CELL_COL = 3  # synthetic cell column for Workflows clip artifacts
# (col 1 = --ss-clips, col 2 = --transcript-clips; the distinct column keeps
# synthetic artifact ids from colliding across the three sheet-free clip paths.)


def _make_synthetic_clip_record(
    *,
    cluster_idx: int,
    cell_col: int,
    study: str,
    participant: str,
    desc: str,
    category: str,
    severity: str,
    start_seconds: float,
    end_seconds: float,
    source_filename: str,
) -> ClipRecord:
    """Build a ClipRecord with synthetic cell + pre-filled times.

    Uses negative cell rows (unreachable for real spreadsheets) and a per-mode
    column to namespace artifact ids. The pre-filled ``times`` triggers the
    fast path in :func:`prepare_clip` so the cell value is never read.
    """
    from types import SimpleNamespace

    start_ts = utils.seconds_to_timestamp(int(start_seconds), force_hours=True)
    end_ts = utils.seconds_to_timestamp(
        max(int(end_seconds), int(start_seconds) + 1), force_hours=True
    )
    cell = SimpleNamespace(value="", row=-(cluster_idx + 1), col=cell_col)
    record: ClipRecord = {
        "cell": cell,
        "desc": desc,
        "study": study,
        "participant": participant,
        "category": category,
        "severity": severity,
        "times": [(start_ts, end_ts)],
        "source_filename": source_filename,
        "cell_annotations": [],
        "segment_annotations": {},
    }
    return record


def build_clip_records(
    *,
    participant: str,
    source_filename: str,
    time_ranges: list[tuple[float, float]],
    description: str,
    category: str = "workflow",
    study: str = "",
    severity: str = "",
    cell_col: int = _WORKFLOW_CELL_COL,
    cell_row_base: int = 0,
    cluster_gap: float | None = None,
    pad_pre: float = 0.0,
    pad_post: float = 0.0,
    max_duration: float | None = None,
) -> list[ClipRecord]:
    """Build sheet-free ClipRecords from explicit ``(start, end)`` second ranges.

    Pre-fills ``times`` (H:MM:SS) on synthetic cells so ``pipeline.process_clips``
    runs without a live spreadsheet — those records hit the fast path in
    :func:`prepare_clip`. This is the public, sheet-free clip-record entry point
    shared by the Workflows ``make_clips`` node, its typed-port adapters, and the
    CLI ``--ss-clips`` / ``--transcript-clips`` paths.

    Args:
        participant: Participant id stamped on every record (e.g. ``"P01"``).
        source_filename: Source video basename, or a ``" + "``-joined list of
            parts for a multi-video participant (resolved by
            ``pipeline._check_source_video``).
        time_ranges: ``(start_seconds, end_seconds)`` pairs to cut.
        description: Clip description applied to every produced record.
        category: Clip category applied to every produced record.
        study: Normalized study name for output paths.
        severity: Optional severity label.
        cell_col: Synthetic cell column for artifact-id namespacing.
        cell_row_base: Added to each record's synthetic (negative) cell row so
            callers can keep ids unique/stable across batches.
        cluster_gap: When set, merge ranges whose gap is within this many seconds
            via :func:`utils.cluster_spans` before building records.
        pad_pre / pad_post / max_duration: Forwarded to ``cluster_spans`` when
            ``cluster_gap`` is set.

    Returns:
        A list of ClipRecords ready for ``pipeline.process_clips`` /
        ``process_reel`` (not yet prepared — those call :func:`prepare_clip`).
    """
    if cluster_gap is not None:
        clustered = utils.cluster_spans(
            list(time_ranges),
            gap_seconds=cluster_gap,
            pad_pre=pad_pre,
            pad_post=pad_post,
            max_duration=max_duration or 0.0,
        )
        ranges = [(cs, ce) for cs, ce, _members in clustered]
    else:
        ranges = list(time_ranges)

    records: list[ClipRecord] = []
    for idx, (start_s, end_s) in enumerate(ranges):
        records.append(
            _make_synthetic_clip_record(
                cluster_idx=cell_row_base + idx,
                cell_col=cell_col,
                study=study,
                participant=participant,
                desc=description,
                category=category,
                severity=severity,
                start_seconds=start_s,
                end_seconds=end_s,
                source_filename=source_filename,
            )
        )
    return records


def discover_clips() -> list[str]:
    """Find generated clips in the effective output directory.

    Scans for .mp4 files and excludes source videos (those matching the
    pattern study_P01.mp4, study_G02.mp4, etc.).

    Returns:
        Sorted list of clip filenames (relative to the output directory)
    """
    base_dir = utils.get_effective_output_dir()
    return sorted(
        p.name
        for p in base_dir.iterdir()
        if p.name.endswith(config.FILEFORMAT)
        and not re.search(config.SOURCE_VIDEO_PATTERN, p.name, re.IGNORECASE)
    )
