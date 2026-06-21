# -*- coding: utf-8 -*-
"""Analysis-ready data export for Screenspace and Transcripts.

Reads the on-disk manifests and produces flat, one-row-per-atomic-unit
records suitable for direct loading into pandas / spreadsheets / BI tools.

Exposed builders:
    build_screenspace_events(manifest, *, include_excluded, participants, detectors)
    build_transcript_segments(manifest)

Serialization:
    to_csv(records, *, preferred_column_order)
    to_json(records)

Bundle writer (used by the --export CLI flag):
    write_export_bundle(output_dir) -> list[Path]
    run_cli_export() -> int
"""

from __future__ import annotations

import csv
import io
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

import config
import utils


# ---- Column ordering ----------------------------------------------------

SCREENSPACE_EVENT_COLUMNS: tuple[str, ...] = (
    "id",
    "participant",
    "source_video",
    "detector",
    "event_type",
    "region",
    "time_in",
    "time_out",
    "duration",
    "confidence",
    "excluded",
    "task_id",
)

SCREENSPACE_PIN_COLUMNS: tuple[str, ...] = (
    "participant",
    "id",
    "timestamp",
    "polarity",
    "label",
    "created_at",
)

_TRANSCRIPT_SEGMENT_BASE_COLS = (
    "participant",
    "segment_id",
    "start",
    "end",
    "duration",
    "text",
    "language",
    "model",
    "source_file",
    "transcribed_at",
    "mark_categories",
    "mark_labels",
)

_FRICTION_MOMENT_COLS = (
    "participant",
    "segment_ids",
    "category",
    "rationale",
    "score",
    "model",
    "computed_at",
)

_FRICTION_SEGMENT_COLS = (
    "participant",
    "segment_id",
    "score",
    "categories",
    "markers",
)


# ---- Helpers ------------------------------------------------------------


def _flatten_value_for_csv(value: Any) -> Any:
    """Coerce list/dict values into CSV-safe strings.

    Lists are joined with ';'. Dicts and other non-primitive types are
    serialized as compact JSON. Strings, numbers, booleans, and None pass
    through.
    """
    if value is None or isinstance(value, (str, int, float, bool)):
        return utils.sanitize_floats(value)
    if isinstance(value, list):
        if all(
            isinstance(item, (str, int, float, bool)) or item is None for item in value
        ):
            return ";".join(
                "" if (safe := utils.sanitize_floats(item)) is None else str(safe)
                for item in value
            )
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def _ordered_columns(
    records: list[dict[str, Any]],
    preferred: tuple[str, ...] | list[str],
) -> list[str]:
    """Build a column list: preferred order first, then alphabetical tail."""
    seen_keys: set[str] = set()
    for rec in records:
        seen_keys.update(rec.keys())
    head = [c for c in preferred if c in seen_keys]
    tail = sorted(k for k in seen_keys if k not in head)
    return head + tail


# ---- Builders -----------------------------------------------------------


def build_screenspace_events(
    manifest: dict[str, Any],
    *,
    include_excluded: bool = False,
    participants: list[str] | None = None,
    detectors: list[str] | None = None,
) -> list[dict[str, Any]]:
    """Flatten Screenspace events into analysis-ready rows.

    Hoists every key in ``event["metadata"]`` to the top level so each
    detector's tool-specific values (magnitude, score, text_found, ...) become
    their own columns.
    """
    raw_events = manifest.get("events", []) or []
    participant_set = set(participants) if participants else None
    detector_set = set(detectors) if detectors else None

    records: list[dict[str, Any]] = []
    for ev in raw_events:
        if not isinstance(ev, dict):
            continue
        if not include_excluded and ev.get("excluded"):
            continue
        if participant_set and ev.get("participant") not in participant_set:
            continue
        if detector_set and ev.get("detector") not in detector_set:
            continue

        time_in = ev.get("time_in", 0.0)
        time_out = ev.get("time_out", time_in)
        try:
            duration = float(time_out) - float(time_in)
        except (TypeError, ValueError):
            duration = 0.0

        record: dict[str, Any] = {
            "id": ev.get("id", ""),
            "participant": ev.get("participant", ""),
            "source_video": ev.get("source_video", ""),
            "detector": ev.get("detector", ""),
            "event_type": ev.get("event_type", ""),
            "region": ev.get("region", ""),
            "time_in": utils.sanitize_floats(time_in),
            "time_out": utils.sanitize_floats(time_out),
            "duration": utils.sanitize_floats(round(duration, 4)),
            "confidence": utils.sanitize_floats(ev.get("confidence", 0.0)),
            "excluded": bool(ev.get("excluded", False)),
            "task_id": ev.get("task_id", ""),
        }
        metadata = ev.get("metadata", {}) or {}
        if isinstance(metadata, dict):
            for k, v in metadata.items():
                if k in record:
                    continue
                record[k] = utils.sanitize_floats(v)
        records.append(record)
    return records


def build_screenspace_pins(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    """One row per calibration pin — provenance for tuned thresholds.

    Flattens ``manifest["pins"]`` (``{participant: [pin, ...]}``) into a flat
    table so a consumer can see which frames a researcher calibrated a detector
    against (and with what polarity) when reviewing the chosen parameters.
    """
    pins_by_participant = manifest.get("pins", {}) or {}
    if not isinstance(pins_by_participant, dict):
        return []
    records: list[dict[str, Any]] = []
    for participant_id, pins in pins_by_participant.items():
        if not isinstance(pins, list):
            continue
        for pin in pins:
            if not isinstance(pin, dict):
                continue
            records.append(
                {
                    "participant": participant_id,
                    "id": pin.get("id", ""),
                    "timestamp": utils.sanitize_floats(pin.get("timestamp", 0.0)),
                    "polarity": pin.get("polarity", ""),
                    "label": pin.get("label", ""),
                    "created_at": pin.get("created_at", ""),
                }
            )
    return records


def build_transcript_segments(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    """One row per transcript segment, with participant and mark info joined."""
    source = manifest.get("source_transcripts", {}) or {}
    marks = manifest.get("marks", []) or []

    marks_by_segment: dict[str, list[dict[str, Any]]] = {}
    for mk in marks:
        if not isinstance(mk, dict):
            continue
        seg_id = mk.get("segment_id")
        if not seg_id:
            continue
        marks_by_segment.setdefault(seg_id, []).append(mk)

    records: list[dict[str, Any]] = []
    for participant_id, entry in source.items():
        if not isinstance(entry, dict):
            continue
        language = entry.get("language", "")
        model = entry.get("model", "")
        source_file = entry.get("source_file", "")
        transcribed_at = entry.get("transcribed_at", "")
        for idx, seg in enumerate(entry.get("segments", []) or []):
            if not isinstance(seg, dict):
                continue
            seg_id = seg.get("id") or f"{participant_id}:{idx}"
            try:
                start = float(seg.get("start", 0.0))
                end = float(seg.get("end", 0.0))
            except (TypeError, ValueError):
                start = 0.0
                end = 0.0
            seg_marks = marks_by_segment.get(seg_id, [])
            records.append(
                {
                    "participant": participant_id,
                    "segment_id": seg_id,
                    "start": utils.sanitize_floats(start),
                    "end": utils.sanitize_floats(end),
                    "duration": utils.sanitize_floats(round(end - start, 4)),
                    "text": seg.get("text", ""),
                    "language": language,
                    "model": model,
                    "source_file": source_file,
                    "transcribed_at": transcribed_at,
                    "mark_categories": [m.get("category", "") for m in seg_marks],
                    "mark_labels": [m.get("label", "") for m in seg_marks],
                }
            )
    return records


def build_friction_moments(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    """One row per LLM-detected friction moment, across all participants."""
    source = manifest.get("source_transcripts", {}) or {}
    records: list[dict[str, Any]] = []
    for participant_id, entry in source.items():
        if not isinstance(entry, dict):
            continue
        friction_data = entry.get("friction")
        if not isinstance(friction_data, dict):
            continue
        computed_at = friction_data.get("computed_at", "")
        model = friction_data.get("model", "")
        for moment in friction_data.get("moments", []) or []:
            if not isinstance(moment, dict):
                continue
            records.append(
                {
                    "participant": participant_id,
                    "segment_ids": [str(s) for s in (moment.get("segment_ids") or [])],
                    "category": moment.get("category", ""),
                    "rationale": moment.get("rationale", ""),
                    "score": utils.sanitize_floats(moment.get("score", 0.0)),
                    "model": model,
                    "computed_at": computed_at,
                }
            )
    return records


def build_friction_segments(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    """One row per scored segment with friction (score > 0), for reproducibility.

    Zero-score segments are omitted to keep the export lean; a consumer can
    reconstruct positional order and treat missing segments as 0 via the
    ``participant:index`` segment IDs.
    """
    source = manifest.get("source_transcripts", {}) or {}
    records: list[dict[str, Any]] = []
    for participant_id, entry in source.items():
        if not isinstance(entry, dict):
            continue
        friction_data = entry.get("friction")
        if not isinstance(friction_data, dict):
            continue
        for seg in friction_data.get("segments", []) or []:
            if not isinstance(seg, dict):
                continue
            score = seg.get("score", 0.0)
            if not score:
                continue
            records.append(
                {
                    "participant": participant_id,
                    "segment_id": seg.get("id", ""),
                    "score": utils.sanitize_floats(score),
                    "categories": list(seg.get("categories") or []),
                    "markers": list(seg.get("markers") or []),
                }
            )
    return records


# ---- Serialization ------------------------------------------------------


def to_csv(
    records: list[dict[str, Any]],
    *,
    preferred_column_order: tuple[str, ...] | list[str] = (),
) -> str:
    """Serialize records to a CSV string.

    Columns: ``preferred_column_order`` first (those that actually appear in
    the data), then any remaining keys sorted alphabetically. List/dict
    values are flattened via :func:`_flatten_value_for_csv`.
    """
    columns = _ordered_columns(records, preferred_column_order)
    if not columns:
        return ""
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=columns, extrasaction="ignore")
    writer.writeheader()
    for rec in records:
        flat = {col: _flatten_value_for_csv(rec.get(col)) for col in columns}
        writer.writerow(flat)
    return buf.getvalue()


def to_json(records: list[dict[str, Any]]) -> str:
    """Serialize records to a JSON string with an export envelope."""
    payload = {
        "exported_at": datetime.now(timezone.utc).isoformat(),
        "version": utils.get_version(),
        "records": records,
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


# ---- Bundle writer ------------------------------------------------------


_SurfaceBuilder = Callable[[dict[str, Any]], list[dict[str, Any]]]

_SURFACES: tuple[tuple[str, _SurfaceBuilder, str, tuple[str, ...]], ...] = (
    (
        "screenspace_events",
        build_screenspace_events,
        config.SCREENSPACE_MANIFEST_FILENAME,
        SCREENSPACE_EVENT_COLUMNS,
    ),
    (
        "screenspace_pins",
        build_screenspace_pins,
        config.SCREENSPACE_MANIFEST_FILENAME,
        SCREENSPACE_PIN_COLUMNS,
    ),
    (
        "transcripts",
        build_transcript_segments,
        config.TRANSCRIPTS_MANIFEST_FILENAME,
        _TRANSCRIPT_SEGMENT_BASE_COLS,
    ),
    (
        "friction_moments",
        build_friction_moments,
        config.TRANSCRIPTS_MANIFEST_FILENAME,
        _FRICTION_MOMENT_COLS,
    ),
    (
        "friction_segments",
        build_friction_segments,
        config.TRANSCRIPTS_MANIFEST_FILENAME,
        _FRICTION_SEGMENT_COLS,
    ),
)


def write_export_bundle(output_dir: Path | None = None) -> list[Path]:
    """Write JSON+CSV exports for every manifest present in *output_dir*.

    Manifests that don't exist on disk are silently skipped. Returns the
    list of files actually written.
    """
    base = Path(output_dir) if output_dir else Path(utils.get_effective_output_dir())
    written: list[Path] = []
    summaries: list[str] = []

    progress = utils.create_progress_bar()

    def _process_surface(
        output_basename: str,
        builder: _SurfaceBuilder,
        manifest_filename: str,
        preferred_cols: tuple[str, ...],
    ) -> None:
        manifest_path = base / manifest_filename
        if not manifest_path.is_file():
            return
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as err:
            utils.warning_print(
                f"Skipping {manifest_filename}: could not read manifest.",
                [f"Error: {err}"],
            )
            return
        records = builder(manifest)
        if not records:
            summaries.append(
                f"Skipping {output_basename}: manifest is present but contains no records."
            )
            return
        json_path = base / f"clipgen_export_{output_basename}.json"
        csv_path = base / f"clipgen_export_{output_basename}.csv"
        json_path.write_text(to_json(records), encoding="utf-8")
        csv_path.write_text(
            to_csv(records, preferred_column_order=preferred_cols),
            encoding="utf-8",
        )
        written.extend([json_path, csv_path])
        summaries.append(
            f"Exported {len(records)} record(s) to {json_path.name} and {csv_path.name}"
        )

    if progress:
        with progress:
            ptask = progress.add_task("Exporting", total=len(_SURFACES))
            for (
                output_basename,
                builder,
                manifest_filename,
                preferred_cols,
            ) in _SURFACES:
                progress.update(ptask, description=f"Exporting {output_basename}")
                _process_surface(
                    output_basename, builder, manifest_filename, preferred_cols
                )
                progress.update(ptask, advance=1)
    else:
        for (
            output_basename,
            builder,
            manifest_filename,
            preferred_cols,
        ) in _SURFACES:
            _process_surface(
                output_basename, builder, manifest_filename, preferred_cols
            )

    for line in summaries:
        utils.info_print(line)
    return written


def run_cli_export() -> int:
    """Entry point for the ``--export`` CLI flag.

    Reads manifests from the effective output directory, writes export
    files, prints a summary, and returns an exit code (0 on success).
    """
    output_dir = Path(utils.get_effective_output_dir())
    utils.info_print(f"Exporting analysis data from manifests in {output_dir}")
    written = write_export_bundle(output_dir)
    if not written:
        utils.warning_print(
            "No exports written.",
            [
                "No manifest files were found in the output directory.",
                f"Expected one or more of: {config.SCREENSPACE_MANIFEST_FILENAME}, "
                f"{config.TRANSCRIPTS_MANIFEST_FILENAME}",
            ],
        )
        return 1
    utils.info_print(f"Wrote {len(written)} export file(s).")
    return 0


__all__ = [
    "build_screenspace_events",
    "build_screenspace_pins",
    "build_transcript_segments",
    "build_friction_moments",
    "build_friction_segments",
    "to_csv",
    "to_json",
    "write_export_bundle",
    "run_cli_export",
    "SCREENSPACE_EVENT_COLUMNS",
    "SCREENSPACE_PIN_COLUMNS",
]
