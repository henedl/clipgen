"""Overview: Flask blueprint for /overview/.

Serves the Overview frontend (assets/web/overview.html + the overview-*.js
satellites) — cohort-level tabs for Metadata, Convergence, and Reports. The
tabs are thin over other blueprints' data (sheet/baseline via ../studio/,
transcripts + report agent via ../transcripts/); the only state owned here is
the Convergence tab's per-lane display offsets, persisted in the output dir.
"""

from __future__ import annotations

import math
from typing import Any, cast

from flask import Blueprint, request

from server_utils import err, ok

import config
import utils

overview_bp = Blueprint("overview", __name__)


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


# Registered before register_static_routes so the catch-all static route cannot shadow it.
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
    data = utils.load_manifest_section("convergence", default={})
    raw = data.get("offsets") if isinstance(data, dict) else None
    return ok(offsets=_clean_convergence_offsets(raw))


@overview_bp.route("/api/convergence/offsets", methods=["PUT"])
def api_convergence_offsets_put():
    """Persist per-lane convergence display offsets.

    Body: {"offsets": {"P01": {"sheet": 12.5, ...}, ...}}. Unknown sources,
    zeros, and non-finite values are dropped per lane; participants left with
    no lanes are dropped. When the cleaned dict is empty, the section is
    removed so a clean output dir has no leftover empty manifest.
    """
    data = request.get_json(silent=True) or {}
    raw = data.get("offsets")
    if not isinstance(raw, dict):
        return err("Invalid offsets payload")

    cleaned = _clean_convergence_offsets(raw)

    utils.save_manifest_section(
        "convergence", {"offsets": cleaned} if cleaned else None
    )

    return ok(offsets=cleaned)


utils.register_static_routes(overview_bp, "overview.html", icons=True)
