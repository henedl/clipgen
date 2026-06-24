"""Workflows: node-based scripting engine (data model + persistence + run engine).

Workflows is clipgen's fourth top-level frontend (next to Studio, Screenspace,
and Transcripts). It is a free-form 2D node canvas where users drag "blueprint
cards" — each wrapping one backend action — and wire typed outputs into typed
inputs to chain capabilities across all three domains (artifact generation,
Screenspace analysis, transcription + thinking agents). A ``WorkflowRunner``
executes the resulting DAG.

This module is the backend home for:

* ``NODE_TYPES`` — a declarative single-source-of-truth registry of node types
  (typed ports + param schema), modelled on ``thinking_agents.AGENTS``. Present
  as of M1; ``serialize_catalog`` feeds the frontend via ``/api/catalog``.

and, in later milestones:

* ``ADAPTERS`` — the typed-port coercion table (e.g. ``events -> clipRecords``).
* ``WorkflowRunner`` — DAG topo-sort + sequential ready-set execution, calling
  the existing pure functions directly with the uniform ``on_progress`` /
  ``cancel_flag`` contract.

M1 ships the manifest persistence helpers **and** the declarative ``NODE_TYPES``
catalog below. The catalog is declarative-only here — each node's ``execute``
callable (plus ``ADAPTERS`` and the ``WorkflowRunner``) lands in M3. See
``plans/WORKFLOWS-PLAN.md``.

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
from typing import Any, Callable, NotRequired, TypedDict

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


# ---------------------------------------------------------------------------
# Node catalog (declarative single source of truth)
# ---------------------------------------------------------------------------
#
# Modelled on ``thinking_agents.AGENTS``: a data-driven, enumerable registry
# the frontend renders generically. Each node carries the typed ports the DAG
# needs (the wire vocabulary lives in plans/WORKFLOWS-PLAN.md). ``execute`` is
# deferred to M3 — these entries are declarative-only, so adding a node in a
# later milestone is "append a NodeType + (in M3) an executor", zero frontend
# edits. ``serialize_catalog`` strips any ``execute`` for the JSON endpoint.


class Port(TypedDict):
    """A typed input/output socket on a node. ``optional`` inputs may be unwired."""

    name: str
    type: str
    optional: NotRequired[bool]


class ParamSpec(TypedDict):
    """A node parameter the frontend renders an editor for (editors land in M2)."""

    name: str
    type: str  # number | string | enum | bool | participant | region | ...
    default: Any
    label: NotRequired[str]
    choices: NotRequired[list[Any]]
    min: NotRequired[float]
    max: NotRequired[float]


class NodeType(TypedDict):
    """One node in the catalog. ``execute`` is filled in M3 (see module docstring)."""

    id: str
    label: str
    domain: str  # artifact | screenspace | transcript | thinking | control
    category: str  # human-facing palette group label
    inputs: list[Port]
    outputs: list[Port]
    params: list[ParamSpec]
    requires: list[str]  # subset of {"sheet", "videoDir"}
    execute: NotRequired[Callable[..., dict[str, Any]]]


# Curated v1 node set (plans/WORKFLOWS-PLAN.md). Keyed by id so the frontend can
# both iterate (palette) and look up a placed node's type. Ports/params may be
# refined when the M3 executors are wired against the underlying functions.
NODE_TYPES: dict[str, NodeType] = {
    # ---- Sources ----
    "video_source": {
        "id": "video_source",
        "label": "Video Source",
        "domain": "artifact",
        "category": "Source",
        "inputs": [],
        "outputs": [
            {"name": "video", "type": "video"},
            {"name": "participant", "type": "participant"},
        ],
        "params": [
            {
                "name": "participant",
                "type": "participant",
                "default": "",
                "label": "Participant",
            },
        ],
        "requires": ["videoDir"],
    },
    "sheet_selection": {
        "id": "sheet_selection",
        "label": "Sheet Selection",
        "domain": "artifact",
        "category": "Source",
        "inputs": [],
        "outputs": [{"name": "clips", "type": "clipRecords"}],
        "params": [
            {
                "name": "selector",
                "type": "string",
                "default": "",
                "label": "Selector (cells/lines/category)",
            },
        ],
        "requires": ["sheet"],
    },
    "region": {
        "id": "region",
        "label": "Region",
        "domain": "screenspace",
        "category": "Source",
        "inputs": [],
        "outputs": [{"name": "region", "type": "region"}],
        "params": [
            {"name": "name", "type": "string", "default": "", "label": "Region name"},
        ],
        "requires": ["videoDir"],
    },
    # ---- Transcript ----
    "transcribe": {
        "id": "transcribe",
        "label": "Transcribe",
        "domain": "transcript",
        "category": "Transcript",
        "inputs": [{"name": "video", "type": "video"}],
        "outputs": [
            {"name": "transcript", "type": "transcript"},
            {"name": "segments", "type": "segments"},
        ],
        "params": [
            {
                "name": "language",
                "type": "string",
                "default": "auto",
                "label": "Language",
            },
        ],
        "requires": ["videoDir"],
    },
    "find_word": {
        "id": "find_word",
        "label": "Find Word",
        "domain": "transcript",
        "category": "Transcript",
        "inputs": [{"name": "segments", "type": "segments"}],
        "outputs": [
            {"name": "timeRange", "type": "timeRange"},
            {"name": "timestamps", "type": "timestamps"},
        ],
        "params": [
            {
                "name": "word",
                "type": "string",
                "default": "",
                "label": "Word or phrase",
            },
            {
                "name": "pad",
                "type": "number",
                "default": 2,
                "min": 0,
                "max": 30,
                "label": "Pad (seconds)",
            },
        ],
        "requires": [],
    },
    # ---- Thinking (Ollama) ----
    "summarize": {
        "id": "summarize",
        "label": "Summarize",
        "domain": "thinking",
        "category": "Thinking",
        "inputs": [{"name": "transcript", "type": "transcript"}],
        "outputs": [{"name": "summary", "type": "summary"}],
        "params": [],
        "requires": [],
    },
    "citations": {
        "id": "citations",
        "label": "Citations",
        "domain": "thinking",
        "category": "Thinking",
        "inputs": [
            {"name": "summary", "type": "summary"},
            {"name": "segments", "type": "segments"},
        ],
        "outputs": [{"name": "citations", "type": "citations"}],
        "params": [],
        "requires": [],
    },
    "friction": {
        "id": "friction",
        "label": "Friction",
        "domain": "thinking",
        "category": "Thinking",
        "inputs": [
            {"name": "segments", "type": "segments"},
            {"name": "summary", "type": "summary", "optional": True},
        ],
        "outputs": [{"name": "friction", "type": "friction"}],
        "params": [],
        "requires": [],
    },
    # ---- Screenspace ----
    "ss_scan": {
        "id": "ss_scan",
        "label": "SS Scan",
        "domain": "screenspace",
        "category": "Screenspace",
        "inputs": [
            {"name": "video", "type": "video"},
            {"name": "region", "type": "region", "optional": True},
            {"name": "timeRange", "type": "timeRange", "optional": True},
        ],
        "outputs": [{"name": "events", "type": "events"}],
        "params": [
            {
                "name": "tool",
                "type": "enum",
                "default": "text",
                "choices": ["text", "color", "change", "similarity"],
                "label": "Tool",
            },
        ],
        "requires": ["videoDir"],
    },
    # ---- Artifact ----
    "make_clips": {
        "id": "make_clips",
        "label": "Make Clips",
        "domain": "artifact",
        "category": "Artifact",
        "inputs": [
            {"name": "clips", "type": "clipRecords", "optional": True},
            {"name": "video", "type": "video", "optional": True},
            {"name": "timeRange", "type": "timeRange", "optional": True},
        ],
        "outputs": [{"name": "artifacts", "type": "artifacts"}],
        "params": [
            {
                "name": "description",
                "type": "string",
                "default": "",
                "label": "Description",
            },
        ],
        "requires": ["videoDir"],
    },
    "build_reel": {
        "id": "build_reel",
        "label": "Build Reel",
        "domain": "artifact",
        "category": "Artifact",
        "inputs": [{"name": "clips", "type": "clipRecords"}],
        "outputs": [
            {"name": "artifacts", "type": "artifacts"},
            {"name": "manifest", "type": "manifest"},
        ],
        "params": [
            {"name": "name", "type": "string", "default": "reel", "label": "Reel name"},
        ],
        "requires": ["videoDir"],
    },
    "timeline_viewer": {
        "id": "timeline_viewer",
        "label": "Timeline Viewer",
        "domain": "artifact",
        "category": "Artifact",
        "inputs": [
            {"name": "artifacts", "type": "artifacts", "optional": True},
            {"name": "events", "type": "events", "optional": True},
            {"name": "segments", "type": "segments", "optional": True},
        ],
        "outputs": [{"name": "viewer", "type": "viewerHtml"}],
        "params": [],
        "requires": [],
    },
    # ---- Control ----
    "gate": {
        "id": "gate",
        "label": "Gate",
        "domain": "control",
        "category": "Control",
        "inputs": [{"name": "value", "type": "scalar"}],
        "outputs": [{"name": "pass", "type": "scalar"}],
        "params": [
            {
                "name": "op",
                "type": "enum",
                "default": ">=",
                "choices": [">=", ">", "<=", "<", "==", "!="],
                "label": "Comparison",
            },
            {"name": "threshold", "type": "number", "default": 0, "label": "Threshold"},
        ],
        "requires": [],
    },
}


def serialize_catalog() -> list[dict[str, Any]]:
    """Return ``NODE_TYPES`` as ordered JSON-safe dicts (without ``execute``).

    Drives ``GET /workflows/api/catalog``. Deliberately *not* routed through
    ``utils.get_frontend_config`` — the catalog is large and Workflows-specific,
    and bolting it onto the shared config would pollute the
    ``tests/test_shared_constants`` contract.
    """
    return [
        {k: v for k, v in node.items() if k != "execute"}
        for node in NODE_TYPES.values()
    ]
