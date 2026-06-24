"""Workflows: node-based scripting engine (data model + persistence + run engine).

Workflows is clipgen's fourth top-level frontend (next to Studio, Screenspace,
and Transcripts). It is a free-form 2D node canvas where users drag "blueprint
cards" — each wrapping one backend action — and wire typed outputs into typed
inputs to chain capabilities across all three domains (artifact generation,
Screenspace analysis, transcription + thinking agents). A ``WorkflowRunner``
executes the resulting DAG.

This module is the backend home for, in later milestones:

* ``NODE_TYPES`` — a declarative single-source-of-truth registry of node types
  (ports + param schema + executor), modelled on ``thinking_agents.AGENTS``.
* ``ADAPTERS`` — the typed-port coercion table (e.g. ``events -> clipRecords``).
* ``WorkflowRunner`` — DAG topo-sort + sequential ready-set execution, calling
  the existing pure functions directly with the uniform ``on_progress`` /
  ``cancel_flag`` contract.

M0 (scaffold) ships only the manifest persistence helpers below; the catalog
and runner land in subsequent milestones. See ``plans/WORKFLOWS-PLAN.md``.

Manifest shape (``workflows_manifest.json`` in the output directory)::

    {
        "blueprints": [ {id, name, nodes, edges, viewport, trigger} ],
        "stashes":    [ {id, name, nodes, edges, createdAt} ],
        "runs":       [ {id, blueprintId, status, nodeStates, startedAt, completedAt} ]
    }

``trigger`` is reserved (always ``null`` in v1) for the deferred auto-launch
phase — kept in the schema so the seam exists without building anything.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import config
import utils


def empty_workflows_manifest() -> dict[str, list[Any]]:
    """Return a fresh, empty workflows manifest with all top-level keys present."""
    return {"blueprints": [], "stashes": [], "runs": []}


def load_workflows_manifest() -> dict[str, Any]:
    """Load ``workflows_manifest.json`` from the output dir (empty default).

    Missing or corrupt files fall back to :func:`empty_workflows_manifest` so
    callers always get the full key set, never a partial dict.
    """
    data = utils.load_json_manifest(config.WORKFLOWS_MANIFEST_FILENAME, default=None)
    if not isinstance(data, dict):
        return empty_workflows_manifest()
    # Backfill any missing top-level keys so callers can index unconditionally.
    base = empty_workflows_manifest()
    base.update({k: v for k, v in data.items() if k in base})
    return base


def save_workflows_manifest(
    blueprints: list[dict[str, Any]],
    stashes: list[dict[str, Any]] | None = None,
    runs: list[dict[str, Any]] | None = None,
) -> Path | None:
    """Persist the workflows manifest atomically; returns the path or ``None``."""
    payload = {
        "blueprints": blueprints,
        "stashes": stashes or [],
        "runs": runs or [],
    }
    return utils.save_json_manifest(
        config.WORKFLOWS_MANIFEST_FILENAME,
        payload,
        warn_label="workflows manifest",
    )
