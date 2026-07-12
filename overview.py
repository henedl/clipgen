# -*- coding: utf-8 -*-
"""Overview: per-participant feature matrix + Flask blueprint for /overview/.

Serves the Overview frontend (assets/web/overview.html + the overview-*.js
satellites), whose Map tab is a 3D similarity space: participants are
positioned so spatial distance reflects behavioral similarity, exposing
clusters and outliers across a study.

Split of responsibilities (thin server, thick client):

- This module builds the RAW (un-normalized) participant x feature matrix
  from the three on-disk manifests and serves it via ``GET /overview/api/data``.
- The client (overview-map.js) does z-scoring, feature-group weighting, the 3D PCA
  projection, and outlier scoring, so weight sliders re-layout instantly
  without server round-trips. Raw values are shipped because the explain
  panel needs them ("12.0 events/min vs cohort mean 5.1").

Feature groups (one slider each in the UI):

- ``observations``   — spreadsheet clip artifacts: category/severity shares.
- ``screenspace``    — detector event rates and confidences.
- ``transcript``     — friction markers/min per category, speech rate.
- ``session_shape``  — WHEN things happen: normalized-session-time histograms.

All features are rates or shares (never raw counts, except one total per
group) so long sessions don't masquerade as outliers. A participant missing
from a source entirely gets ``None`` for that group's cells plus an
``availability`` flag, letting the client impute the cohort mean for layout
while excluding those features from the outlier score.

Duration denominators are deliberately per-source and never mixed: screenspace
rates divide by the participant's last event time; transcript rates divide by
the last segment end. The two can disagree for the same participant (different
recordings/trims), and mixing them would silently skew rates.
"""

from __future__ import annotations

import math
import re
from pathlib import Path
from typing import Any, cast

from flask import Blueprint, request

from server_utils import err, ok

import config
import data_export
import friction
import screenspace_manifest
import transcripts
import utils
import viewer

# Feature-group keys, in payload/UI order. Mirrored client-side from the
# payload itself (columns carry their group), so no JS constant to sync.
GROUP_KEYS: tuple[str, ...] = (
    "observations",
    "screenspace",
    "transcript",
    "session_shape",
)

GROUP_LABELS: dict[str, str] = {
    "observations": "Observations",
    "screenspace": "Screenspace",
    "transcript": "Transcript",
    "session_shape": "Session shape",
}

# Bin count for the session-shape histograms. 8 keeps the group's dimension
# count (2 x bins) comparable to the other groups after the client's
# 1/sqrt(group_size) equalization.
SESSION_SHAPE_BINS: int = 8

_SLUG_RE = re.compile(r"[^a-z0-9]+")


def _slug(text: str) -> str:
    """Filesystem/JS-safe feature-key fragment: lowercase, runs -> ``_``."""
    return _SLUG_RE.sub("_", str(text).strip().lower()).strip("_") or "unknown"


def _column(key: str, group: str, label: str) -> dict[str, str]:
    return {"key": key, "group": group, "label": label}


def _minutes(seconds: float) -> float:
    return seconds / 60.0 if seconds > 0 else 0.0


# ---- Group builders -------------------------------------------------------
#
# Each builder returns ``(columns, values)`` where columns is the ordered
# feature list ``[{"key", "group", "label"}, ...]`` (identical for every
# participant) and values maps ``participant -> {key: float}``. A participant
# absent from ``values`` has no data for the whole group (-> None cells).


def build_observation_features(
    artifacts: list[dict[str, Any]],
) -> tuple[list[dict[str, str]], dict[str, dict[str, float]]]:
    """Category/severity shares + total count from spreadsheet clip artifacts."""
    by_participant: dict[str, list[dict[str, Any]]] = {}
    categories: set[str] = set()
    severities: set[str] = set()
    for art in artifacts:
        if not isinstance(art, dict):
            continue
        pid = art.get("participant", "")
        if not pid:
            continue
        by_participant.setdefault(pid, []).append(art)
        categories.add(art.get("category") or "uncategorized")
        if art.get("severity"):
            severities.add(art["severity"])

    cat_list = sorted(categories)
    sev_list = sorted(severities)
    columns = [
        _column(f"obs_cat_{_slug(c)}", "observations", f"{c} share of observations")
        for c in cat_list
    ]
    columns += [
        _column(f"obs_sev_{_slug(s)}", "observations", f"{s} share of observations")
        for s in sev_list
    ]
    columns.append(_column("obs_total", "observations", "observation count"))

    values: dict[str, dict[str, float]] = {}
    for pid, arts in by_participant.items():
        total = len(arts)
        row: dict[str, float] = {}
        for c in cat_list:
            n = sum(1 for a in arts if (a.get("category") or "uncategorized") == c)
            row[f"obs_cat_{_slug(c)}"] = round(n / total, 4)
        for s in sev_list:
            n = sum(1 for a in arts if a.get("severity") == s)
            row[f"obs_sev_{_slug(s)}"] = round(n / total, 4)
        row["obs_total"] = float(total)
        values[pid] = row
    return columns, values


def _screenspace_rows_by_participant(
    event_rows: list[dict[str, Any]],
) -> dict[str, list[dict[str, Any]]]:
    by_participant: dict[str, list[dict[str, Any]]] = {}
    for row in event_rows:
        pid = row.get("participant", "")
        if pid:
            by_participant.setdefault(pid, []).append(row)
    return by_participant


def _screenspace_duration_seconds(rows: list[dict[str, Any]]) -> float:
    """Session-length proxy: the participant's last event end time."""
    latest = 0.0
    for row in rows:
        try:
            latest = max(latest, float(row.get("time_out") or 0.0))
        except (TypeError, ValueError):
            continue
    return latest


def build_screenspace_features(
    event_rows: list[dict[str, Any]],
) -> tuple[list[dict[str, str]], dict[str, dict[str, float]]]:
    """Per-detector rates + confidences from flattened screenspace events.

    ``event_rows`` come from :func:`data_export.build_screenspace_events`
    (excluded events already dropped).
    """
    by_participant = _screenspace_rows_by_participant(event_rows)
    detectors = sorted({r.get("detector", "") for r in event_rows if r.get("detector")})

    columns = [
        _column(f"ss_rate_{_slug(d)}", "screenspace", f"{d} events/min")
        for d in detectors
    ]
    columns += [
        _column(f"ss_conf_{_slug(d)}", "screenspace", f"{d} mean confidence")
        for d in detectors
    ]
    columns.append(_column("ss_total_rate", "screenspace", "events/min (all)"))
    columns.append(_column("ss_nav_share", "screenspace", "navigational share"))

    values: dict[str, dict[str, float]] = {}
    for pid, rows in by_participant.items():
        minutes = _minutes(_screenspace_duration_seconds(rows))
        row_out: dict[str, float] = {}
        for d in detectors:
            d_rows = [r for r in rows if r.get("detector") == d]
            rate = len(d_rows) / minutes if minutes > 0 else 0.0
            confs = [
                float(r["confidence"])
                for r in d_rows
                if isinstance(r.get("confidence"), (int, float))
            ]
            row_out[f"ss_rate_{_slug(d)}"] = round(rate, 4)
            row_out[f"ss_conf_{_slug(d)}"] = (
                round(sum(confs) / len(confs), 4) if confs else 0.0
            )
        row_out["ss_total_rate"] = round(len(rows) / minutes, 4) if minutes > 0 else 0.0
        row_out["ss_nav_share"] = round(
            sum(1 for r in rows if r.get("navigational")) / len(rows), 4
        )
        values[pid] = row_out
    return columns, values


def _transcript_duration_seconds(segments: list[dict[str, Any]]) -> float:
    """Session-length proxy: the last segment's end time."""
    latest = 0.0
    for seg in segments:
        try:
            latest = max(latest, float(seg.get("end") or 0.0))
        except (TypeError, ValueError):
            continue
    return latest


def _friction_scored_and_stats(
    entry: dict[str, Any], segments: list[dict[str, Any]], duration: float
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Scored friction rows + session stats, from the manifest when present.

    Falls back to running the deterministic scorer directly (friction.py —
    no Ollama involved) when the friction agent never ran for this entry, so
    the transcript feature group works on transcribe-only studies.
    """
    friction_data = entry.get("friction")
    if isinstance(friction_data, dict) and isinstance(friction_data.get("stats"), dict):
        return list(friction_data.get("segments") or []), friction_data["stats"]
    scored = friction.score_segments(segments)
    return scored, friction.compute_stats(scored, duration)


def build_transcript_features(
    source_transcripts: dict[str, Any],
) -> tuple[list[dict[str, str]], dict[str, dict[str, float]]]:
    """Friction category rates + speech rate from transcript entries."""
    columns = [
        _column(f"tr_fric_{_slug(c)}", "transcript", f"{c} markers/min")
        for c in friction.CATEGORY_ORDER
    ]
    columns.append(_column("tr_markers_per_min", "transcript", "friction markers/min"))
    columns.append(_column("tr_segment_count", "transcript", "transcript segments"))
    columns.append(_column("tr_words_per_min", "transcript", "words/min"))
    columns.append(_column("tr_duration_min", "transcript", "session minutes"))

    values: dict[str, dict[str, float]] = {}
    for pid, entry in source_transcripts.items():
        if not isinstance(entry, dict):
            continue
        segments = [s for s in (entry.get("segments") or []) if isinstance(s, dict)]
        if not segments:
            continue
        duration = _transcript_duration_seconds(segments)
        minutes = _minutes(duration)
        _, stats = _friction_scored_and_stats(entry, segments, duration)
        by_category = stats.get("by_category", {}) or {}

        row: dict[str, float] = {}
        for c in friction.CATEGORY_ORDER:
            count = by_category.get(c, 0) or 0
            row[f"tr_fric_{_slug(c)}"] = round(count / minutes, 4) if minutes else 0.0
        row["tr_markers_per_min"] = float(stats.get("markers_per_minute", 0.0) or 0.0)
        row["tr_segment_count"] = float(len(segments))
        words = sum(len((s.get("text") or "").split()) for s in segments)
        row["tr_words_per_min"] = round(words / minutes, 4) if minutes else 0.0
        row["tr_duration_min"] = round(minutes, 4)
        values[pid] = row
    return columns, values


def _histogram_shares(positions: list[tuple[float, float]], bins: int) -> list[float]:
    """Weighted histogram over normalized positions in [0, 1], as shares.

    ``positions`` is ``[(normalized_time, weight), ...]``. Returns ``bins``
    shares summing to 1 when any weight lands, else all zeros.
    """
    counts = [0.0] * bins
    for pos, weight in positions:
        if weight <= 0:
            continue
        idx = min(int(max(pos, 0.0) * bins), bins - 1)
        counts[idx] += weight
    total = sum(counts)
    if total <= 0:
        return counts
    return [round(c / total, 4) for c in counts]


def build_session_shape_features(
    event_rows: list[dict[str, Any]],
    source_transcripts: dict[str, Any],
    *,
    bins: int = SESSION_SHAPE_BINS,
) -> tuple[list[dict[str, str]], dict[str, dict[str, float]]]:
    """WHEN-in-the-session histograms (each normalized by that source's length).

    Two sub-histograms per participant: screenspace event density and friction
    score density. A participant with only one of the two sources gets zeros
    for the other's bins (the group is available as long as either source is).
    """
    columns = [
        _column(
            f"shape_ss_bin{i}",
            "session_shape",
            f"event density {i + 1}/{bins} of session",
        )
        for i in range(bins)
    ]
    columns += [
        _column(
            f"shape_fric_bin{i}",
            "session_shape",
            f"friction density {i + 1}/{bins} of session",
        )
        for i in range(bins)
    ]

    ss_by_participant = _screenspace_rows_by_participant(event_rows)
    values: dict[str, dict[str, float]] = {}

    participants = set(ss_by_participant)
    for pid, entry in source_transcripts.items():
        if isinstance(entry, dict) and entry.get("segments"):
            participants.add(pid)

    for pid in participants:
        row: dict[str, float] = {}

        ss_rows = ss_by_participant.get(pid, [])
        ss_duration = _screenspace_duration_seconds(ss_rows)
        ss_positions: list[tuple[float, float]] = []
        if ss_duration > 0:
            for r in ss_rows:
                try:
                    mid = (
                        float(r.get("time_in") or 0.0) + float(r.get("time_out") or 0.0)
                    ) / 2.0
                except (TypeError, ValueError):
                    continue
                ss_positions.append((mid / ss_duration, 1.0))
        for i, share in enumerate(_histogram_shares(ss_positions, bins)):
            row[f"shape_ss_bin{i}"] = share

        entry = source_transcripts.get(pid)
        fric_positions: list[tuple[float, float]] = []
        if isinstance(entry, dict):
            segments = [s for s in (entry.get("segments") or []) if isinstance(s, dict)]
            duration = _transcript_duration_seconds(segments)
            if segments and duration > 0:
                scored, _ = _friction_scored_and_stats(entry, segments, duration)
                score_by_id = {
                    str(s.get("id")): float(s.get("score") or 0.0)
                    for s in scored
                    if isinstance(s, dict)
                }
                for idx, seg in enumerate(segments):
                    seg_id = str(seg.get("id") or idx)
                    score = score_by_id.get(seg_id, 0.0)
                    if score <= 0:
                        continue
                    try:
                        mid = (
                            float(seg.get("start") or 0.0)
                            + float(seg.get("end") or 0.0)
                        ) / 2.0
                    except (TypeError, ValueError):
                        continue
                    fric_positions.append((mid / duration, score))
        for i, share in enumerate(_histogram_shares(fric_positions, bins)):
            row[f"shape_fric_bin{i}"] = share

        values[pid] = row
    return columns, values


# ---- Assembly -------------------------------------------------------------


def build_feature_matrix() -> dict[str, Any]:
    """Assemble the /overview/api/data payload from the three on-disk manifests.

    The only I/O-touching function in this module. Missing manifests yield
    empty shapes (the loaders never raise for absent files), so the payload
    degrades to ``participants: []`` on a fresh output directory.
    """
    ss_manifest = screenspace_manifest.load_screenspace_manifest()
    event_rows = data_export.build_screenspace_events(ss_manifest)
    tr_manifest = transcripts.load_transcripts_manifest()
    source_transcripts = tr_manifest.get("source_transcripts", {}) or {}
    artifacts = viewer.load_manifest_artifacts()

    groups: list[tuple[str, list[dict[str, str]], dict[str, dict[str, float]]]] = [
        ("observations", *build_observation_features(artifacts)),
        ("screenspace", *build_screenspace_features(event_rows)),
        ("transcript", *build_transcript_features(source_transcripts)),
        (
            "session_shape",
            *build_session_shape_features(event_rows, source_transcripts),
        ),
    ]

    columns = [col for _, cols, _ in groups for col in cols]
    participants = sorted({pid for _, _, vals in groups for pid in vals})

    matrix: list[list[float | None]] = []
    availability: dict[str, dict[str, bool]] = {}
    for pid in participants:
        row: list[float | None] = []
        avail: dict[str, bool] = {}
        for group_key, cols, vals in groups:
            has_group = pid in vals
            avail[group_key] = has_group
            if has_group:
                row.extend(vals[pid].get(col["key"], 0.0) for col in cols)
            else:
                row.extend(None for _ in cols)
        matrix.append(row)
        availability[pid] = avail

    return utils.sanitize_floats(
        {
            "participants": participants,
            "columns": columns,
            "matrix": matrix,
            "availability": availability,
            "groups": [{"key": k, "label": GROUP_LABELS[k]} for k in GROUP_KEYS],
            "config": utils.get_frontend_config(),
        }
    )


# ---- Flask blueprint -------------------------------------------------------

overview_bp = Blueprint("overview", __name__)


# Registered before register_static_routes so the catch-all /<path:filename>
# static route can never shadow the API (mirrors the other blueprints).
@overview_bp.route("/api/data")
def api_map_data():
    return ok(**build_feature_matrix())


def _clean_convergence_offsets(raw: object) -> dict[str, dict[str, float]]:
    """Normalize nested per-lane convergence offsets to {pid: {source: float}}.

    Drops: non-string/empty participant ids, non-dict participant values,
    unknown source keys (outside config.CONVERGENCE_SOURCES), non-numeric /
    non-finite / zero lane values, and participants left with no lanes.
    """
    cleaned: dict[str, dict[str, float]] = {}
    if not isinstance(raw, dict):
        return cleaned
    raw_map = cast(dict[str, Any], raw)
    for pid, lanes in raw_map.items():
        if not isinstance(pid, str) or not pid or not isinstance(lanes, dict):
            continue
        lane_map = cast(dict[str, Any], lanes)
        clean_lanes: dict[str, float] = {}
        for source, value in lane_map.items():
            if source not in config.CONVERGENCE_SOURCES:
                continue
            try:
                num = float(value)
            except (TypeError, ValueError):
                continue
            if math.isfinite(num) and num != 0:
                clean_lanes[source] = num
        if clean_lanes:
            cleaned[pid] = clean_lanes
    return cleaned


@overview_bp.route("/api/convergence/offsets")
def api_convergence_offsets_get():
    """Return persisted per-lane convergence display offsets (seconds, signed).

    Independent from /studio/api/sheet/baseline: baselines convert sheet
    wall-clock to video-time (sheet-only). Offsets shift a participant's
    events per data source so misaligned recording start times — or a single
    drifting source such as spreadsheet timestamps — can be nudged until lanes
    line up visually in the Convergence tab.

    Response: {"ok": true, "offsets": {"P01": {"sheet": 12.5, "screenspace": 12.5}}}
    """
    data = utils.load_json_manifest(
        config.CONVERGENCE_OFFSETS_FILENAME, default={"offsets": {}}
    )
    raw = data.get("offsets") if isinstance(data, dict) else None
    return ok(offsets=_clean_convergence_offsets(raw))


@overview_bp.route("/api/convergence/offsets", methods=["PUT"])
def api_convergence_offsets_put():
    """Persist per-lane convergence display offsets.

    Body: {"offsets": {"P01": {"sheet": 12.5, ...}, ...}}. Unknown sources,
    zeros, and non-finite values are dropped per lane; participants left with
    no lanes are dropped. When the cleaned dict is empty, the manifest file is
    deleted so a clean output dir has no leftover empty manifest.
    """
    data = request.get_json(silent=True) or {}
    raw = data.get("offsets")
    if not isinstance(raw, dict):
        return err("Invalid offsets payload")

    cleaned = _clean_convergence_offsets(raw)

    settings_path = (
        Path(utils.get_effective_output_dir()) / config.CONVERGENCE_OFFSETS_FILENAME
    )
    if not cleaned:
        if settings_path.is_file():
            try:
                settings_path.unlink()
            except OSError:
                pass
    else:
        utils.save_json_manifest(
            config.CONVERGENCE_OFFSETS_FILENAME, {"offsets": cleaned}
        )

    return ok(offsets=cleaned)


utils.register_static_routes(overview_bp, "overview.html", icons=True)
