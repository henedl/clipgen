# -*- coding: utf-8 -*-
"""Insights data model and manifest persistence for clipgen.

Insight record shape:
  {
    "id": "ins_<8hex>",       # uuid4 hex prefix, assigned by create_insight()
    "title": str,
    "summary": str,
    "severity": str,          # Critical/High/Medium/Low/N/A/Positive/Very Positive
    "status": str,            # "draft" or "final"
    "createdAt": str,         # ISO 8601 UTC
    "updatedAt": str,         # ISO 8601 UTC
    "causes":    {"narrative": str, "artifacts": [ids]},
    "behaviors": {"narrative": str, "artifacts": [ids]},
    "impacts":   {"narrative": str, "artifacts": [ids]},
    "timelineContext": str,
  }

Key functions:
  load_insights_manifest() / save_insights_manifest() – read/write insights_manifest.json
  create_insight()  – returns a new insight dict with all fields initialized
  update_insight()  – merge a partial dict into an existing insight by id
  delete_insight()  – remove an insight by id
"""

import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import config
import utils


def _empty_manifest() -> dict[str, Any]:
    return {"meta": {}, "insights": []}


def load_insights_manifest() -> dict[str, Any]:
    """Load the insights manifest from the output directory.

    Returns a dict with 'meta' and 'insights' keys.
    """
    return utils.load_json_manifest(
        config.INSIGHTS_MANIFEST_FILENAME, default=_empty_manifest()
    )


def save_insights_manifest(
    meta: dict[str, Any], insights: list[dict[str, Any]]
) -> Path | None:
    """Write the full insights manifest to disk.

    Returns the manifest path on success, or None on failure.
    """
    meta = dict(meta)
    meta["generatedAt"] = datetime.now(timezone.utc).isoformat()
    meta.setdefault("version", utils.get_version())

    return utils.save_json_manifest(
        config.INSIGHTS_MANIFEST_FILENAME,
        {"meta": meta, "insights": insights},
        warn_label="insights manifest",
    )


def create_insight(
    title: str = "", severity: str = "", status: str = "draft"
) -> dict[str, Any]:
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
    insights: list[dict[str, Any]],
    insight_id: str,
    updates: dict[str, Any],
) -> dict[str, Any] | None:
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


def delete_insight(insights: list[dict[str, Any]], insight_id: str) -> bool:
    """Remove an insight by ID. Returns True if found and removed."""
    for i, insight in enumerate(insights):
        if insight["id"] == insight_id:
            insights.pop(i)
            return True
    return False
