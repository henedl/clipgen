"""File and filename operations for clipgen."""

import os
import re
import threading
import unicodedata
from pathlib import Path
from typing import Any

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
    """Atomically reserve a unique output path, as an empty placeholder file.

    The placeholder is what makes this atomic: parallel clip / reel / gallery
    workers each get a distinct path even when their filename templates collide.
    Whoever fills the artifact overwrites it; a caller that aborts before writing
    real content must remove it with ``release_reservation()``.

    Appends '-1', '-2', … when the name is taken, and truncates over-long names.
    *file_format* defaults to ``config.FILEFORMAT``.
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


def get_source_video_filenames(
    study: str, participant: str, override: str | None = None
) -> list[str]:
    """Resolve the ordered list of expected source video filenames for a clip.

    Supports multiple source videos per participant that form one continuous
    concatenated timeline (e.g. a recording that broke off, or a diary study).

    - With an *override* (the spreadsheet ``Filename`` row): split on ``+`` so a
      cell like ``"morning.mp4 + afternoon.mp4"`` yields two files in order. Each
      part follows the same default-extension rule (:func:`_apply_default_extension`).
      Empty parts are dropped. Order is authoritative (concatenation order).
    - Without an override: returns the single plain name built from
      ``config.SOURCE_FILENAME_PATTERN`` (default ``{study}_{participant}.mp4``).
      On-disk numbered-suffix auto-detection (``study_P01-1.mp4`` ...) is
      resolved later by the pipeline against the input directory via
      :func:`discover_numbered_source_videos`.

    Always returns at least one entry.
    """
    if override is not None:
        override = override.strip()
    if override:
        parts = [p.strip() for p in override.split("+")]
        names = [_apply_default_extension(p) for p in parts if p]
        if names:
            return names
    return [utils.format_source_video_stem(study, participant) + config.FILEFORMAT]


def resolve_source_video_paths(
    study: str, participant: str, override: str | None, input_dir: Path
) -> list[Path]:
    """Resolve a participant's ordered source-video paths without fuzzy matching.

    Resolution order (same policy as the pipeline's interactive resolver, minus
    the fuzzy-match prompt — for non-interactive callers like the studio server
    and the ``--transcribe`` CLI):

    - *override* present (the spreadsheet ``Filename`` row) → its plus-separated
      parts, resolved in order.
    - else the plain patterned name (``config.SOURCE_FILENAME_PATTERN``, default
      ``{study}_{participant}.mp4``) when it exists on disk.
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

    Globs *input_dir* for the patterned stem plus a ``-N`` suffix
    (``config.SOURCE_FILENAME_PATTERN``, e.g. ``study_P01-1.mp4``) and sorts
    the hits by the integer N (so ``-2`` precedes ``-10``). Used only when no
    spreadsheet override is set and the plain patterned file is absent.
    Returns [] when no numbered parts exist.

    The parts must form a gapless ``1..N`` sequence: concatenating non-contiguous
    parts back-to-back would map global timestamps into the wrong sub-video. When
    a gap is found, a warning is emitted and [] is returned (treated as no valid
    multi-part sequence) rather than building a silently-wrong timeline.
    """
    prefix = utils.format_source_video_stem(study, participant)
    suffix_re = re.compile(rf"-(\d+){re.escape(config.FILEFORMAT)}$", re.IGNORECASE)
    matches: list[tuple[int, Path]] = []
    for p in input_dir.glob(f"{prefix}-*{config.FILEFORMAT}"):
        m = suffix_re.search(p.name)
        if m:
            matches.append((int(m.group(1)), p))
    matches.sort(key=lambda item: item[0])
    indices = [n for n, _ in matches]
    if indices and not utils.numbered_parts_are_contiguous(indices):
        utils.warning_print(
            f"Numbered source videos for '{prefix}' are non-contiguous "
            f"(found parts {indices}); expected 1..N with no gaps.",
            [
                (
                    "Ignoring the numbered sequence; rename the parts to a gapless "
                    "1..N sequence to enable concatenation."
                ),
            ],
        )
        return []
    return [p for _, p in matches]


def _browser_seekable(video_paths: list[str]) -> bool | None:
    """Can a browser seek every one of this participant's source files?

    ``False`` as soon as any part is a fragmented MP4 — a participant whose
    second part is unseekable is just as broken as one whose first part is.
    ``None`` means "unknown" (non-MP4 container, missing file) and the UI must
    say nothing; see :func:`video.probe_container_seekability`.
    """
    import video  # lazy: video imports files, so this cannot be a top-level import

    saw_unknown = False
    for path in video_paths:
        probed = video.probe_container_seekability(path)
        if probed is None:
            saw_unknown = True
        elif not probed["browser_seekable"]:
            return False
    return None if saw_unknown else True


def resolve_participant_videos(sheet_context: Any = None) -> list[dict[str, Any]]:
    """Return every participant the tools should list — sheet columns, then disk.

    Neither source alone is the truth, so this is their union:

    - The sheet's participant columns, resolved override-aware through
      :func:`resolve_source_video_paths`. A column whose file is missing stays in
      the list with ``has_video: False`` so the gap is visible rather than silent.
    - Every participant :func:`utils.discover_participant_videos` finds on disk
      that the sheet does not mention, carrying its own discovered paths verbatim.
      Discovery accepts any ``*_P<x>`` / ``*_G<x>`` name regardless of study
      prefix, so ``study_P13.mp4`` must NOT be re-resolved against the sheet's
      study name — that would invent a ``clipgen-test_P13.mp4`` that is not there.
    - Every participant with a user filename override
      (``config.FILENAME_OVERRIDES``) that neither of the above produced. An
      override names a file discovery cannot pattern-match, so without this a
      mind-map participant pointed at ``recording 3.mp4`` would stay invisible.

    Sheet order first, then disk-only ids sorted, then override-only ids sorted;
    deduped by id, sheet winning. ``sheet_context=None`` makes this a plain scan
    with every entry ``in_sheet: False``.

    The input directory is derived here rather than passed in — handing the sheet
    half a different directory than ``discover_participant_videos`` (which always
    derives its own) would be a silent-mismatch bug.

    Returns:
        ``[{"id", "video_paths", "has_video", "in_sheet", "browser_seekable"}]``,
        freshly built. Never the memoized dicts ``discover_participant_videos``
        returns — those are shared with the Workflows blueprint and must not be
        stamped on.
    """
    import spreadsheet  # lazy: files.py is CLI-hot; spreadsheet pulls in google_api

    study = ""
    sheet_ids: list[str] = []
    overrides: dict[str, str | None] = {}
    if sheet_context is not None:
        study = getattr(sheet_context, "study_name", "")
        sheet_ids = spreadsheet.get_participant_list(
            sheet_context.header_row,
            sheet_context.id_cell,
            sheet_context.num_participants,
        )
        overrides = spreadsheet.participant_filename_overrides(sheet_context)

    input_dir = Path(utils.get_effective_input_dir())
    entries: list[dict[str, Any]] = []
    seen: set[str] = set()
    for pid in sheet_ids:
        if pid in seen:  # a sheet with a duplicated column header
            continue
        seen.add(pid)
        paths = resolve_source_video_paths(study, pid, overrides.get(pid), input_dir)
        path_strs = [str(p) for p in paths]
        entries.append(
            {
                "id": pid,
                "video_paths": path_strs,
                "has_video": paths[0].is_file(),
                "in_sheet": True,
                "browser_seekable": _browser_seekable(path_strs),
            }
        )
    for found in utils.discover_participant_videos():  # already sorted by id
        if found["id"] in seen:
            continue
        seen.add(found["id"])
        user_override = config.FILENAME_OVERRIDES.get(found["id"])
        if user_override:
            # The user pointed this participant at a specific file; what
            # discovery happened to match by name is not it.
            paths = resolve_source_video_paths(
                "", found["id"], user_override, input_dir
            )
            found_paths = [str(p) for p in paths]
            has_video = paths[0].is_file()
        else:
            found_paths = list(found["video_paths"])  # copy — source is memoized
            has_video = found["has_video"]
        entries.append(
            {
                "id": found["id"],
                "video_paths": found_paths,
                "has_video": has_video,
                "in_sheet": False,
                "browser_seekable": _browser_seekable(found_paths),
            }
        )
    for pid in sorted(config.FILENAME_OVERRIDES):
        if pid in seen:
            continue
        override = config.FILENAME_OVERRIDES[pid]
        if not override:
            continue
        # study is irrelevant here: resolve_source_video_paths ignores it when
        # an override is present.
        paths = resolve_source_video_paths("", pid, override, input_dir)
        path_strs = [str(p) for p in paths]
        entries.append(
            {
                "id": pid,
                "video_paths": path_strs,
                "has_video": paths[0].is_file(),
                "in_sheet": False,
                "browser_seekable": _browser_seekable(path_strs),
            }
        )
    return entries


def derive_sheet_meta(worksheet: Any) -> dict[str, str] | None:
    """Return ``{type, id_or_path, label, worksheet}`` identifying *worksheet*.

    This triple keys the per-source filename overrides in ``start.json`` and
    the Start overlay's recent-projects entries. Shared by the server (picker
    display, override seeding) and the CLI launch path, so both derive the
    same identity for the same worksheet.
    """
    if worksheet is None:
        return None
    try:
        import excel_io

        if isinstance(worksheet, excel_io.ExcelSheetAdapter):
            path = getattr(worksheet, "_workbook_path", "") or ""
            if not path:
                return None
            return {
                "type": "excel",
                "id_or_path": path,
                "label": Path(path).name,
                "worksheet": getattr(worksheet, "title", ""),
            }
    except ImportError:
        pass  # no excel_io in this build; fall through to the gspread branch
    # gspread Worksheet (or anything quacking like one): use the parent
    # spreadsheet title as both the identifier and the label.
    parent = getattr(worksheet, "spreadsheet", None)
    title = getattr(parent, "title", "") if parent is not None else ""
    if not title:
        return None
    return {
        "type": "google",
        "id_or_path": title,
        "label": title,
        "worksheet": getattr(worksheet, "title", ""),
    }


def find_participant_record(
    sheet_context: Any, participant_id: str
) -> dict[str, Any] | None:
    """First :func:`resolve_participant_videos` record matching *participant_id*.

    A fresh resolve on every call — for routes that must see the live input
    directory (e.g. after ``POST /api/dirs``) rather than a blueprint's
    ``_participants`` cache.
    """
    for participant in resolve_participant_videos(sheet_context):
        if participant["id"] == participant_id:
            return participant
    return None


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
    cell_ref = utils.safe_cell_a1(clip["cell"].row, clip["cell"].col)

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
    if not clip["times"] and utils.has_non_ignored_timestamp_content(  # noqa: SIM102 - the comment below belongs to the inner branch
        cleaned_cell_value
    ):
        # Only show this detailed per-cell warning at verbose verbosity.
        if getattr(config, "VERBOSITY", config.STANDARD) >= config.VERBOSE:
            utils.warning_print(
                f"No valid timestamps found in cell {cell_ref}",
                [
                    f"Cell contents: '{clip['cell'].value}'",
                    f"Participant: {clip['participant']}, Description: {(clip.get('desc') or '')[:50]}...",
                ],
            )

    # Clean description: remove bracketed prefix and sanitize for use in filename.
    # `.get` mirrors the fast path above: a record built without desc/category is
    # legal (synthetic clips, sheet rows with empty columns) and must not KeyError.
    raw_desc = clip.get("desc") or ""
    bracket_pos = raw_desc.rfind("]")
    cleaned_desc = (
        raw_desc[bracket_pos + 1 :].strip() if bracket_pos >= 0 else raw_desc.strip()
    )
    clip["desc"] = utils.sanitize_filename(cleaned_desc)
    if config.DEBUGGING:
        config.debug_ic(clip["desc"])

    # Sanitize category (handle missing/None/empty)
    if clip.get("category"):
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
        source_filename: Source video basename, or a ``" + "``-joined list of
            parts for a multi-video participant (resolved by
            ``pipeline._check_source_video``).
        cell_col: Synthetic cell column for artifact-id namespacing.
        cell_row_base: Added to each record's synthetic (negative) cell row so
            callers can keep ids unique/stable across batches.
        cluster_gap: When set, merge ranges closer than this many seconds via
            :func:`utils.cluster_spans` first; *pad_pre* / *pad_post* /
            *max_duration* are forwarded to it and ignored otherwise.

    Returns:
        ClipRecords ready for ``pipeline.process_clips`` / ``process_reel``, but
        not yet prepared — those call :func:`prepare_clip`.
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

    Scans for FILEFORMAT files and excludes source videos — anything matching
    the configured ``SOURCE_FILENAME_PATTERN`` (study_P01.mp4, study_G02-2.mp4,
    etc.), via :func:`utils.compile_source_video_regex`.

    Returns:
        Sorted list of clip filenames (relative to the output directory)
    """
    base_dir = utils.get_effective_output_dir()
    source_re = utils.compile_source_video_regex()
    return sorted(
        p.name
        for p in base_dir.iterdir()
        if p.name.endswith(config.FILEFORMAT) and not source_re.fullmatch(p.name)
    )
