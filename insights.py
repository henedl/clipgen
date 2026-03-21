# -*- coding: utf-8 -*-
"""Insights data model and manifest persistence for clipgen.

Handles CRUD operations for insight records and read/write of
the insights manifest JSON file.
"""

from __future__ import annotations

import json
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

import config
import utils


def _empty_manifest() -> Dict[str, Any]:
    return {"meta": {}, "insights": []}


def load_insights_manifest() -> Dict[str, Any]:
    """Load the insights manifest from the output directory.

    Returns a dict with 'meta' and 'insights' keys.
    Handles missing 'summary' field on legacy insights (defaults to "").
    """
    manifest_path = (
        Path(utils.get_effective_output_dir()) / config.INSIGHTS_MANIFEST_FILENAME
    )
    if not manifest_path.is_file():
        return _empty_manifest()
    try:
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return _empty_manifest()

    if not isinstance(data, dict):
        return _empty_manifest()

    for insight in data.get("insights", []):
        insight.setdefault("summary", "")

    return {
        "meta": data.get("meta", {}),
        "insights": data.get("insights", []),
    }


def save_insights_manifest(
    meta: Dict[str, Any], insights: List[Dict[str, Any]]
) -> Optional[Path]:
    """Write the full insights manifest to disk.

    Returns the manifest path on success, or None on failure.
    """
    meta = dict(meta)
    meta["generatedAt"] = datetime.now(timezone.utc).isoformat()
    meta.setdefault("version", config.VERSIONNUM)

    manifest_path = (
        Path(utils.get_effective_output_dir()) / config.INSIGHTS_MANIFEST_FILENAME
    )
    try:
        manifest_path.write_text(
            json.dumps(
                {"meta": meta, "insights": insights}, ensure_ascii=False, indent=2
            ),
            encoding="utf-8",
        )
        return manifest_path
    except OSError as e:
        utils.warning_print(f"Could not write insights manifest: {e}")
        return None


def create_insight(
    title: str = "", severity: str = "", status: str = "draft"
) -> Dict[str, Any]:
    """Create a new insight dict with all fields initialized."""
    now = datetime.now(timezone.utc).isoformat()
    return {
        "id": f"ins_{uuid.uuid4().hex[:8]}",
        "title": title or "Untitled insight",
        "summary": "",
        "severity": severity,
        "status": status,
        "createdAt": now,
        "updatedAt": now,
        "causes": {"narrative": "", "artifacts": []},
        "behaviors": {"narrative": "", "artifacts": []},
        "impacts": {"narrative": "", "artifacts": []},
        "timelineContext": "",
    }


def update_insight(
    insights: List[Dict[str, Any]],
    insight_id: str,
    updates: Dict[str, Any],
) -> Optional[Dict[str, Any]]:
    """Find an insight by ID, merge updates, set updatedAt.

    Returns the updated insight, or None if not found.
    """
    for insight in insights:
        if insight["id"] == insight_id:
            for key, value in updates.items():
                if key not in ("id", "createdAt"):
                    insight[key] = value
            insight["updatedAt"] = datetime.now(timezone.utc).isoformat()
            return insight
    return None


def delete_insight(insights: List[Dict[str, Any]], insight_id: str) -> bool:
    """Remove an insight by ID. Returns True if found and removed."""
    for i, insight in enumerate(insights):
        if insight["id"] == insight_id:
            insights.pop(i)
            return True
    return False
