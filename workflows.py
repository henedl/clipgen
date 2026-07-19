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
* ``NodeContext`` + the per-node ``execute`` callables (M3) — each a thin adapter
  over an existing pure function, keyed by output-port name. Every domain value a
  node emits embeds a "source descriptor" so the adapters below stay pure.
* ``ADAPTERS`` (M3) — the typed-port coercion table (e.g. ``events -> clipRecords``),
  pure ``value -> value`` callables the runner applies when an output type differs
  from the consuming input type.
* ``WorkflowRunner`` (M4) — DAG topo-sort + sequential ready-set execution, calling
  the executors directly with the uniform ``on_progress`` / ``cancel_flag`` /
  ``cancel_event`` contract ``NodeContext`` carries. ``topo_order`` rejects cycles;
  control edges (a gate's ``control`` output) gate downstream without feeding data.

See ``plans/archive/WORKFLOWS-PLAN.md``.

Manifest shape (``workflows_manifest.json`` in the output directory)::

    {
        "blueprints": [ {id, name, nodes, edges, viewport, trigger} ],
        "stashes":    [ {id, name, nodes, edges, createdAt} ],
        "runs":       [ {id, blueprintId, status, nodeStates, startedAt, completedAt} ]
    }

``trigger`` holds the watch-dir auto-launch binding (P6): ``null`` (or
``{"type": "watch_dir", "enabled": false}``) when disarmed, or
``{"type": "watch_dir", "enabled": true}`` on the single armed blueprint. The
directory watcher in ``workflows_server`` auto-runs the armed blueprint (one run
per just-arrived participant). The ``type`` field leaves room for future trigger
kinds (transcript_complete / scan_event chaining) without a schema migration.
"""

from __future__ import annotations

import copy
import json
import os
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, NotRequired, TypedDict, cast

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


def _is_empty_workflows_manifest(payload: dict[str, Any]) -> bool:
    """True when nothing worth persisting exists.

    A blueprint counts only if it carries graph content (nodes or edges); a bare
    auto-created "Untitled" with empty nodes/edges is treated as empty even when
    renamed, so a zero-interaction Workflows launch writes no file. Any stash or
    run is user-meaningful and keeps the manifest.
    """
    if payload.get("stashes") or payload.get("runs"):
        return False
    for blueprint in payload.get("blueprints", []):
        if blueprint.get("nodes") or blueprint.get("edges"):
            return False
    return True


def save_workflows_manifest(
    blueprints: list[dict[str, Any]],
    stashes: list[dict[str, Any]] | None = None,
    runs: list[dict[str, Any]] | None = None,
) -> Path | None:
    """Persist the workflows manifest atomically; returns the path or ``None``.

    Skips the write (and removes any stale file) when the manifest is empty, so
    an unused canvas leaves no junk in the output dir.
    """
    payload = {
        "blueprints": blueprints,
        "stashes": stashes or [],
        "runs": runs or [],
    }
    if _is_empty_workflows_manifest(payload):
        utils.remove_json_manifest(config.WORKFLOWS_MANIFEST_FILENAME)
        return None
    return utils.save_json_manifest(
        config.WORKFLOWS_MANIFEST_FILENAME,
        payload,
        warn_label="workflows manifest",
    )


# ---------------------------------------------------------------------------
# Execution context
# ---------------------------------------------------------------------------


@dataclass
class NodeContext:
    """Run-wide context handed to every node executor.

    Carries launch context (output dir + optional sheet) plus the uniform
    progress and cancellation seams. Executors forward ``cancel_flag`` to
    scan/clip callees and ``cancel_event`` to thinking-agent callees — the dual
    contract the backend functions already expose (see
    ``plans/archive/WORKFLOWS-PLAN.md`` and AGENTS.md). ``WorkflowRunner`` builds
    one ``NodeContext`` per run; tests construct it directly.
    """

    input_dir: Path
    output_dir: Path
    sheet_context: Any = None
    worksheet: Any = None
    on_progress: Callable[[float], None] = lambda _fraction: None
    cancel_event: threading.Event = field(default_factory=threading.Event)

    def cancel_flag(self) -> bool:
        """Predicate form of the cancel signal (scan/clip callees want this)."""
        return self.cancel_event.is_set()

    def resolve_videos(self, participant: str) -> list[str]:
        """Return a participant's ordered source video path(s), or ``[]`` if none.

        Multi-part participants (a session split across numbered files) return all
        parts in timeline order — mirrors ``utils.discover_participant_videos``.
        """
        for entry in utils.discover_participant_videos():
            if entry.get("id") == participant and entry.get("has_video"):
                return list(entry["video_paths"])
        return []


# ---------------------------------------------------------------------------
# Node catalog (declarative single source of truth)
# ---------------------------------------------------------------------------
#
# Modelled on ``thinking_agents.AGENTS``: a data-driven, enumerable registry
# the frontend renders generically. Each node carries the typed ports the DAG
# needs (the wire vocabulary lives in plans/archive/WORKFLOWS-PLAN.md). Adding
# a node is "append a NodeType + an executor" (executors section below), zero
# frontend edits. ``serialize_catalog`` strips ``execute`` for the JSON endpoint.


class Port(TypedDict):
    """A typed input/output socket on a node. ``optional`` inputs may be unwired."""

    name: str
    type: str
    optional: NotRequired[bool]


class ParamSpec(TypedDict):
    """A node parameter the frontend renders an editor for."""

    name: str
    type: str  # number | string | enum | bool | participant | region | ...
    default: Any
    label: NotRequired[str]
    choices: NotRequired[list[Any]]
    min: NotRequired[float]
    max: NotRequired[float]
    # P5: an empty value here is a guaranteed no-op/failure — the pre-run
    # validation panel surfaces it as an error and disables Run.
    required: NotRequired[bool]


class NodeType(TypedDict):
    """One node in the catalog. ``execute`` is bound in the executors section below."""

    id: str
    label: str
    description: str  # one-line palette tooltip / on-card help text
    domain: str  # artifact | screenspace | transcript | thinking | control
    category: str  # human-facing palette group label
    inputs: list[Port]
    outputs: list[Port]
    params: list[ParamSpec]
    requires: list[str]  # subset of {"sheet", "videoDir"}
    execute: NotRequired[Callable[..., dict[str, Any]]]
    # Hidden from the palette but kept in the catalog (e.g. the per-detector
    # ss_<tool> nodes, which the unified Detect node + Multitool read for specs).
    hidden: NotRequired[bool]
    # Whether a ss_<tool> detector can be a Multitool chain step (the frontend
    # derives its step-type list from this flag — see _MULTITOOL_STEP_TOOLS).
    multitoolStep: NotRequired[bool]


# Shared by the three Ollama thinking nodes: a free-text override of the model
# name (blank → the configured default). A ``string`` rather than ``enum`` because
# the installed Ollama models are environment-specific and not known server-side.
_OLLAMA_MODEL_PARAM: ParamSpec = {
    "name": "model",
    "type": "string",
    "default": "",
    "label": "Ollama model (blank = default)",
}


# Curated v1 node set (plans/archive/WORKFLOWS-PLAN.md). Keyed by id so the
# frontend can both iterate (palette) and look up a placed node's type.
NODE_TYPES: dict[str, NodeType] = {
    # ---- Sources ----
    "video_source": {
        "id": "video_source",
        "label": "Video Source",
        "description": "A participant's source video, picked from the video directory.",
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
                "required": True,
            },
        ],
        "requires": ["videoDir"],
    },
    "sheet_selection": {
        "id": "sheet_selection",
        "label": "Sheet Selection",
        "description": "Clip records pulled from the spreadsheet by a cell, line, or category selector.",
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
    "time_range": {
        "id": "time_range",
        "label": "Time Range",
        "description": "Manually entered start–end time ranges to drive downstream clips.",
        "domain": "artifact",
        "category": "Source",
        "inputs": [],
        "outputs": [{"name": "timeRange", "type": "timeRange"}],
        "params": [
            {
                "name": "ranges",
                "type": "string",
                "default": "",
                "label": "Ranges (e.g. 1:23-1:45, 2:00-2:30)",
                "required": True,
            },
        ],
        "requires": [],
    },
    "region": {
        "id": "region",
        "label": "Region",
        "description": "A named screen region that scopes Screenspace detection to part of the frame.",
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
        "description": "Transcribe a video's audio into a timestamped transcript and segments.",
        "domain": "transcript",
        "category": "Transcript",
        "inputs": [{"name": "video", "type": "video"}],
        "outputs": [
            {"name": "transcript", "type": "transcript"},
            {"name": "segments", "type": "segments"},
        ],
        "params": [
            {
                "name": "model",
                "type": "enum",
                "default": config.TRANSCRIBE_MODEL,
                "choices": ["tiny", "base", "small", "medium", "large-v3"],
                "label": "Whisper model",
            },
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
        "description": "Find a word or phrase in transcript segments and emit its time ranges.",
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
                "required": True,
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
    "transcript_export": {
        "id": "transcript_export",
        "label": "Transcript Export",
        "description": "Write a transcript file (Markdown, SRT, or VTT) from a transcript or segments.",
        "domain": "transcript",
        "category": "Transcript",
        "inputs": [
            {"name": "transcript", "type": "transcript", "optional": True},
            {"name": "segments", "type": "segments", "optional": True},
        ],
        "outputs": [{"name": "artifacts", "type": "artifacts"}],
        "params": [
            {
                "name": "format",
                "type": "enum",
                "default": config.TRANSCRIBE_FORMAT,
                "choices": ["md", "srt", "vtt"],
                "label": "Format",
            },
        ],
        "requires": [],
    },
    # ---- Thinking (Ollama) ----
    "summarize": {
        "id": "summarize",
        "label": "Summarize",
        "description": "Summarize a transcript into a paragraph and key bullet points (Ollama).",
        "domain": "thinking",
        "category": "Thinking",
        "inputs": [{"name": "transcript", "type": "transcript"}],
        "outputs": [{"name": "summary", "type": "summary"}],
        "params": [_OLLAMA_MODEL_PARAM],
        "requires": [],
    },
    "citations": {
        "id": "citations",
        "label": "Citations",
        "description": "Link summary claims back to the transcript segments that support them (Ollama).",
        "domain": "thinking",
        "category": "Thinking",
        "inputs": [
            {"name": "summary", "type": "summary"},
            {"name": "segments", "type": "segments"},
        ],
        "outputs": [{"name": "citations", "type": "citations"}],
        "params": [_OLLAMA_MODEL_PARAM],
        "requires": [],
    },
    "friction": {
        "id": "friction",
        "label": "Friction",
        "description": "Score transcript segments for usability friction and surface the rough moments.",
        "domain": "thinking",
        "category": "Thinking",
        "inputs": [
            {"name": "segments", "type": "segments"},
            {"name": "summary", "type": "summary", "optional": True},
        ],
        "outputs": [{"name": "friction", "type": "friction"}],
        "params": [_OLLAMA_MODEL_PARAM],
        "requires": [],
    },
    # ---- Screenspace ----
    # The ten per-detector nodes (ss_text … ss_boundary) are appended below the
    # literal from ``_SS_DETECTOR_SPECS`` so each tool's real params reach the
    # scan (the old single ``ss_scan`` passed ``parameters={}``).
    "multitool": {
        "id": "multitool",
        "label": "Multitool",
        "description": "Chain two or more per-frame detectors, matching frames where all of them fire.",
        "domain": "screenspace",
        "category": "Screenspace",
        "inputs": [
            {"name": "video", "type": "video"},
            {"name": "region", "type": "region", "optional": True},
        ],
        "outputs": [{"name": "events", "type": "events"}],
        "params": [
            {
                "name": "steps",
                "type": "step-list",
                "default": [],
                "label": "Steps (≥2, chained per frame)",
            },
        ],
        "requires": ["videoDir"],
    },
    "timelapse": {
        "id": "timelapse",
        "label": "Timelapse",
        "description": "Condense a video into a sped-up timelapse clip or GIF.",
        "domain": "screenspace",
        "category": "Screenspace",
        "inputs": [
            {"name": "video", "type": "video"},
            {"name": "region", "type": "region", "optional": True},
        ],
        "outputs": [{"name": "artifacts", "type": "artifacts"}],
        "params": [
            {
                "name": "speedup_factor",
                "type": "number",
                "default": 10.0,
                "min": 1,
                "label": "Speed-up ×",
            },
            {
                "name": "output_format",
                "type": "enum",
                "default": "mp4",
                "choices": ["mp4", "gif"],
                "label": "Format",
            },
            {
                "name": "sample_interval",
                "type": "number",
                "default": 0.0,
                "min": 0,
                "label": "Sample interval (s)",
            },
        ],
        "requires": ["videoDir"],
    },
    "heatmap": {
        "id": "heatmap",
        "label": "Heatmap",
        "description": "Render a heatmap image from detector events (match the upstream detector style).",
        "domain": "screenspace",
        "category": "Screenspace",
        "inputs": [{"name": "events", "type": "events"}],
        "outputs": [{"name": "artifacts", "type": "artifacts"}],
        "params": [
            {
                "name": "style",
                "type": "enum",
                "default": "change",
                "choices": ["template", "flow", "change"],
                "label": "Style (needs matching upstream detector)",
            },
            {
                "name": "output",
                "type": "enum",
                "default": "image",
                "choices": ["image", "gif", "rolling_gif"],
                "label": "Output",
            },
            {
                "name": "frames",
                "type": "number",
                "default": 24,
                "min": 2,
                "label": "GIF frames",
            },
            {
                "name": "window",
                "type": "number",
                "default": 6,
                "min": 1,
                "label": "Rolling window (frames)",
            },
        ],
        "requires": ["videoDir"],
    },
    # ---- Artifact ----
    "highlights": {
        "id": "highlights",
        "label": "Highlights",
        "description": "Score clips by severity, uniqueness, and annotations, keeping the best within a time budget.",
        "domain": "artifact",
        "category": "Artifact",
        "inputs": [{"name": "clips", "type": "clipRecords"}],
        "outputs": [{"name": "clips", "type": "clipRecords"}],
        "params": [
            {
                "name": "budget",
                "type": "number",
                "default": config.HIGHLIGHTS_REEL_DURATION_SECONDS,
                "min": 1,
                "label": "Budget (seconds)",
            },
        ],
        "requires": [],
    },
    "make_clips": {
        "id": "make_clips",
        "label": "Make Clips",
        "description": "Cut clips, screenshots, or GIFs from clip records, a video, or time ranges.",
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
            {
                "name": "output_format",
                "type": "enum",
                "default": "clip",
                "choices": ["clip", "screen", "gif"],
                "label": "Output format",
            },
            {
                "name": "titlecards",
                "type": "bool",
                "default": False,
                "label": "Titlecards",
            },
            {
                "name": "titlecard_duration",
                "type": "number",
                "default": config.TITLECARD_DURATION_SECONDS,
                "min": 1,
                "max": 30,
                "label": "Titlecard duration (s)",
            },
            # Pad fields omit "min" so the number input accepts negatives
            # (negative = trim inward); max_duration keeps min 0 (0 = no cap).
            {
                "name": "pad_start",
                "type": "number",
                "default": 0,
                "label": "Pad start (s)",
            },
            {"name": "pad_end", "type": "number", "default": 0, "label": "Pad end (s)"},
            {
                "name": "max_duration",
                "type": "number",
                "default": 0,
                "min": 0,
                "label": "Max duration (s, 0 = none)",
            },
        ],
        "requires": ["videoDir"],
    },
    "interval_captures": {
        "id": "interval_captures",
        "label": "Interval Captures",
        "description": "Sample a screenshot or GIF at a fixed interval across the video (or each input time range).",
        "domain": "artifact",
        "category": "Artifact",
        "inputs": [
            {"name": "video", "type": "video"},
            {"name": "timeRange", "type": "timeRange", "optional": True},
        ],
        "outputs": [{"name": "artifacts", "type": "artifacts"}],
        "params": [
            {
                "name": "interval",
                "type": "number",
                "default": config.GALLERY_INTERVAL_SECONDS,
                "min": 1,
                "label": "Interval (s)",
                "required": True,
            },
            {
                "name": "output_format",
                "type": "enum",
                "default": "screen",
                "choices": ["screen", "gif"],
                "label": "Output format",
            },
            {
                "name": "gif_duration",
                "type": "number",
                "default": config.GALLERY_GIF_DURATION_SECONDS,
                "min": 1,
                "label": "GIF duration (s)",
            },
        ],
        "requires": ["videoDir"],
    },
    "build_reel": {
        "id": "build_reel",
        "label": "Build Reel",
        "description": "Concatenate clips into a single reel video plus a manifest.",
        "domain": "artifact",
        "category": "Artifact",
        "inputs": [{"name": "clips", "type": "clipRecords"}],
        "outputs": [
            {"name": "artifacts", "type": "artifacts"},
            {"name": "manifest", "type": "manifest"},
        ],
        "params": [
            {"name": "name", "type": "string", "default": "reel", "label": "Reel name"},
            {
                "name": "chronological",
                "type": "bool",
                "default": False,
                "label": "Chronological order",
            },
            # Pad fields omit "min" so the number input accepts negatives
            # (negative = trim inward); max_duration keeps min 0 (0 = no cap).
            {
                "name": "pad_start",
                "type": "number",
                "default": 0,
                "label": "Pad start (s)",
            },
            {"name": "pad_end", "type": "number", "default": 0, "label": "Pad end (s)"},
            {
                "name": "max_duration",
                "type": "number",
                "default": 0,
                "min": 0,
                "label": "Max duration (s, 0 = none)",
            },
        ],
        "requires": ["videoDir"],
    },
    "data_export": {
        "id": "data_export",
        "label": "Data Export",
        "description": "Export events and segments as analysis-ready JSON/CSV tables.",
        "domain": "artifact",
        "category": "Artifact",
        "inputs": [
            {"name": "events", "type": "events", "optional": True},
            {"name": "segments", "type": "segments", "optional": True},
        ],
        "outputs": [{"name": "artifacts", "type": "artifacts"}],
        "params": [
            {
                "name": "format",
                "type": "enum",
                "default": "both",
                "choices": ["both", "json", "csv"],
                "label": "Format",
            },
        ],
        "requires": [],
    },
    "timeline_viewer": {
        "id": "timeline_viewer",
        "label": "Timeline Viewer",
        "description": "Bundle artifacts, events, and segments into a standalone timeline HTML viewer.",
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
    "measure": {
        "id": "measure",
        "label": "Measure",
        "description": "Reduce events, clips, or segments to a single number (count, confidence, or duration).",
        "domain": "control",
        "category": "Control",
        "inputs": [
            {"name": "events", "type": "events", "optional": True},
            {"name": "clips", "type": "clipRecords", "optional": True},
            {"name": "segments", "type": "segments", "optional": True},
        ],
        "outputs": [{"name": "value", "type": "scalar"}],
        "params": [
            {
                "name": "metric",
                "type": "enum",
                "default": "count",
                "choices": ["count", "max_confidence", "total_duration"],
                "label": "Metric",
            },
        ],
        "requires": [],
    },
    "gate": {
        "id": "gate",
        "label": "Gate",
        "description": "Compare a measured value to a threshold to allow or skip downstream nodes.",
        "domain": "control",
        "category": "Control",
        "inputs": [{"name": "value", "type": "scalar"}],
        # ``pass`` is a CONTROL output: it carries no data, it gates. The runner
        # skips a node when an upstream gate completed with ``pass`` False, and
        # excludes control edges from a node's data inputs. The universal
        # ``__gate__`` input port the frontend renders is also ``control``-typed,
        # so a gate can wire into any node (exact-match) as a control dependency.
        "outputs": [{"name": "pass", "type": "control"}],
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
    "gate_collection": {
        "id": "gate_collection",
        "label": "Threshold Gate",
        "description": "Measure events, clips, or segments and gate downstream nodes on the result (Measure + Gate in one node).",
        "domain": "control",
        "category": "Control",
        "inputs": [
            {"name": "events", "type": "events", "optional": True},
            {"name": "clips", "type": "clipRecords", "optional": True},
            {"name": "segments", "type": "segments", "optional": True},
        ],
        # ``pass`` gates exactly like the plain Gate's output (see that node).
        "outputs": [{"name": "pass", "type": "control"}],
        "params": [
            {
                "name": "metric",
                "type": "enum",
                "default": "count",
                "choices": ["count", "max_confidence", "total_duration"],
                "label": "Metric",
            },
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


# Per-detector Screenspace nodes. Each entry's params are lifted from the matching
# ``screenspace_tools`` class (the knobs its ``scan`` reads). The three
# reference-based detectors (template/similarity/scene) self-extract their
# reference from the node's region at ``reference_seconds`` so the canvas needs no
# upload UI. ``_build_ss_scan_params`` (below) assembles these flat params into the
# nested ``scan_params`` each scan expects.
_SS_DETECTOR_LABELS: dict[str, str] = {
    "text": "Detect Text",
    "color": "Detect Color",
    "change": "Detect Change",
    "similarity": "Detect Similarity",
    "numbers": "Detect Numbers",
    "template": "Detect Template",
    "flow": "Detect Motion",
    "scene": "Detect Scene",
    "inactivity": "Detect Inactivity",
    "boundary": "Detect Boundary",
}

_SS_DETECTOR_DESCRIPTIONS: dict[str, str] = {
    "text": "Detect when target text appears in the region via OCR.",
    "color": "Detect when a target color appears or covers enough of the region.",
    "change": "Detect visual change between frames in the region.",
    "similarity": "Detect frames matching a reference image sampled from the region.",
    "numbers": "Read numbers in the region via OCR and compare them to a target.",
    "template": "Detect a reference template (sampled from the region) appearing in the frame.",
    "flow": "Detect motion in the region via optical flow.",
    "scene": "Detect scene changes against a reference fingerprint sampled from the region.",
    "inactivity": "Detect stretches of inactivity (no change) in the region.",
    "boundary": "Detect UI boundaries or edges appearing in the region.",
}

_INTERVAL_PARAM: ParamSpec = {
    "name": "interval",
    "type": "number",
    "default": 0,
    "min": 0,
    "label": "Interval (s, 0=auto)",
}

_SS_DETECTOR_SPECS: dict[str, list[ParamSpec]] = {
    "text": [
        {
            "name": "search_string",
            "type": "string",
            "default": "",
            "label": "Search text",
            "required": True,
        },
        {
            "name": "fuzzy_threshold",
            "type": "number",
            "default": 80,
            "min": 0,
            "max": 100,
            "label": "Fuzzy match %",
        },
        {
            "name": "interval",
            "type": "number",
            "default": 2.0,
            "min": 0,
            "label": "Interval (s, 0=auto)",
        },
    ],
    "color": [
        {
            "name": "color_h",
            "type": "number",
            "default": 0,
            "min": 0,
            "max": 179,
            "label": "Hue",
        },
        {
            "name": "color_s",
            "type": "number",
            "default": 0,
            "min": 0,
            "max": 255,
            "label": "Saturation",
        },
        {
            "name": "color_v",
            "type": "number",
            "default": 0,
            "min": 0,
            "max": 255,
            "label": "Value",
        },
        {
            "name": "tol_h",
            "type": "number",
            "default": 10,
            "min": 0,
            "max": 179,
            "label": "Hue tol",
        },
        {
            "name": "tol_s",
            "type": "number",
            "default": 50,
            "min": 0,
            "max": 255,
            "label": "Sat tol",
        },
        {
            "name": "tol_v",
            "type": "number",
            "default": 50,
            "min": 0,
            "max": 255,
            "label": "Val tol",
        },
        {
            "name": "color_mode",
            "type": "enum",
            "default": "average",
            "choices": ["average", "presence"],
            "label": "Mode",
        },
        {
            "name": "min_coverage",
            "type": "number",
            "default": 0.0,
            "min": 0,
            "max": 1,
            "label": "Min coverage",
        },
        _INTERVAL_PARAM,
    ],
    "change": [
        {
            "name": "threshold",
            "type": "number",
            "default": config.SCREENSPACE_CHANGE_RATIO_THRESHOLD,
            "min": 0,
            "max": 1,
            "label": "Change ratio",
        },
        {
            "name": "noise_threshold",
            "type": "number",
            "default": config.SCREENSPACE_NOISE_THRESHOLD,
            "min": 0,
            "label": "Noise threshold",
        },
        {
            "name": "require_consecutive",
            "type": "number",
            "default": 1,
            "min": 1,
            "label": "Consecutive frames",
        },
        _INTERVAL_PARAM,
    ],
    "similarity": [
        {
            "name": "reference_seconds",
            "type": "number",
            "default": 0.0,
            "min": 0,
            "label": "Reference time (s)",
        },
        {
            "name": "threshold",
            "type": "number",
            "default": config.SCREENSPACE_SSIM_THRESHOLD,
            "min": 0,
            "max": 1,
            "label": "SSIM threshold",
        },
        _INTERVAL_PARAM,
    ],
    "numbers": [
        {
            "name": "operator",
            "type": "enum",
            "default": "gt",
            "choices": ["gt", "lt", "gte", "lte", "eq", "range"],
            "label": "Operator",
        },
        {
            "name": "target_value",
            "type": "number",
            "default": 0,
            "label": "Target value",
        },
        {"name": "range_min", "type": "number", "default": 0, "label": "Range min"},
        {"name": "range_max", "type": "number", "default": 0, "label": "Range max"},
        {
            "name": "integers_only",
            "type": "bool",
            "default": False,
            "label": "Integers only",
        },
        {
            "name": "interval",
            "type": "number",
            "default": 2.0,
            "min": 0,
            "label": "Interval (s, 0=auto)",
        },
    ],
    "template": [
        {
            "name": "reference_seconds",
            "type": "number",
            "default": 0.0,
            "min": 0,
            "label": "Reference time (s)",
        },
        {
            "name": "threshold",
            "type": "number",
            "default": config.SCREENSPACE_TEMPLATE_MATCH_THRESHOLD,
            "min": 0,
            "max": 1,
            "label": "Match threshold",
        },
        {
            "name": "template_scale",
            "type": "number",
            "default": 1.0,
            "min": 0,
            "label": "Template scale",
        },
        _INTERVAL_PARAM,
    ],
    "flow": [
        {
            "name": "magnitude_threshold",
            "type": "number",
            "default": config.SCREENSPACE_FLOW_MAGNITUDE_THRESHOLD,
            "min": 0,
            "label": "Magnitude threshold",
        },
        {
            "name": "require_consecutive",
            "type": "number",
            "default": 1,
            "min": 1,
            "label": "Consecutive frames",
        },
        _INTERVAL_PARAM,
    ],
    "scene": [
        {
            "name": "reference_seconds",
            "type": "number",
            "default": 0.0,
            "min": 0,
            "label": "Reference time (s)",
        },
        {
            "name": "threshold",
            "type": "number",
            "default": config.SCREENSPACE_SCENE_SIMILARITY_THRESHOLD,
            "min": 0,
            "max": 1,
            "label": "Scene threshold",
        },
        _INTERVAL_PARAM,
    ],
    "inactivity": [
        {
            "name": "threshold",
            "type": "number",
            "default": config.SCREENSPACE_INACTIVITY_PHASH_THRESHOLD,
            "min": 0,
            "label": "Phash threshold",
        },
        {
            "name": "min_duration",
            "type": "number",
            "default": 0.0,
            "min": 0,
            "label": "Min duration (s)",
        },
        _INTERVAL_PARAM,
    ],
    "boundary": [
        {
            "name": "threshold",
            "type": "number",
            "default": 0,
            "min": 0,
            "label": "Threshold (0=auto)",
        },
        {
            "name": "min_gap",
            "type": "number",
            "default": 0.0,
            "min": 0,
            "label": "Min gap (s)",
        },
        {
            "name": "metric",
            "type": "enum",
            "default": config.SCREENSPACE_BOUNDARY_METRIC,
            "choices": ["hybrid", "phash", "scene"],
            "label": "Metric",
        },
        _INTERVAL_PARAM,
    ],
}

# Detectors whose scan needs a reference frame self-extracted from the node region.
_SS_REFERENCE_DETECTORS = frozenset({"similarity", "template", "scene"})

# Detectors usable as a Multitool chain step: the per-frame (``check_frame``)
# detectors that need no uploaded reference. Single source of truth — served to
# the frontend via each node's ``multitoolStep`` flag so the step editor derives
# the list instead of hardcoding it (the "no duplicated JS constants" rule).
# ``tests/test_workflows_executors`` cross-checks this against the actual tool
# classes (override ``check_frame`` AND not reference-based) so it can't drift.
_MULTITOOL_STEP_TOOLS = frozenset(
    {"color", "change", "flow", "text", "numbers", "inactivity"}
)

for _ss_tool in _SS_DETECTOR_SPECS:
    NODE_TYPES[f"ss_{_ss_tool}"] = {
        "id": f"ss_{_ss_tool}",
        "label": _SS_DETECTOR_LABELS[_ss_tool],
        "description": _SS_DETECTOR_DESCRIPTIONS[_ss_tool],
        "domain": "screenspace",
        "category": "Screenspace",
        "inputs": [
            {"name": "video", "type": "video"},
            {"name": "region", "type": "region", "optional": True},
            {"name": "timeRange", "type": "timeRange", "optional": True},
        ],
        "outputs": [{"name": "events", "type": "events"}],
        "params": list(_SS_DETECTOR_SPECS[_ss_tool]),
        "requires": ["videoDir"],
        # Hidden from the palette: the unified "detect" node below is the
        # palette-facing entry. These stay in the catalog as the per-detector
        # spec source (Detect editor + Multitool steps) and keep old blueprints
        # and built-in recipes that reference ss_<tool> directly runnable.
        "hidden": True,
        # Whether this detector can be a Multitool step (the frontend derives the
        # step-type list from this flag — see _MULTITOOL_STEP_TOOLS).
        "multitoolStep": _ss_tool in _MULTITOOL_STEP_TOOLS,
    }

# Unified palette-facing detector: one node whose ``detector`` dropdown swaps the
# per-detector param set (the frontend reads it from the hidden ss_<tool> nodes
# above). Dispatches to the same _run_ss_detector body the ss_<tool> nodes use.
NODE_TYPES["detect"] = {
    "id": "detect",
    "label": "Detect",
    "description": "Detect a region condition (text, colour, motion, numbers). Pick the detector.",
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
            "name": "detector",
            "type": "enum",
            "default": "text",
            "choices": list(_SS_DETECTOR_SPECS.keys()),
            "label": "Detector",
        }
    ],
    "requires": ["videoDir"],
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


def serialize_adapters() -> list[list[str]]:
    """Return the ``ADAPTERS`` keys as JSON-safe ``[src, dst, description]`` rows.

    Drives the ``adapters`` field of ``GET /workflows/api/catalog`` so the
    frontend's ``canConnect`` can accept the same coercions the runner applies
    (``_gather_inputs``), and so a coerced wire's tooltip can explain the
    transformation. Serving the table — rather than duplicating it in JS — keeps
    UI wire-validity in lockstep with the runner (``ADAPTERS`` defined below is the
    single source of truth; ``tests/test_workflows_api`` guards parity).
    """
    return [
        [src, dst, _ADAPTER_DESCRIPTIONS.get((src, dst), "")] for src, dst in ADAPTERS
    ]


# ---------------------------------------------------------------------------
# Built-in recipes (P4) — read-only stashes served alongside user stashes
# ---------------------------------------------------------------------------
#
# These are the headline graphs shipped as ready-to-stamp sub-graphs. They are
# *code, not data*: ``GET /api/stashes`` prepends them to the user's persisted
# stashes, so they are never seeded into the manifest (no migration/dedup). The
# ``builtin`` flag makes them read-only — the stash CRUD routes reject renaming
# or deleting them. Node/edge shapes mirror the on-canvas blueprint shapes
# (``{id, type, params, position}`` / ``{id, from, fromPort, to, toPort}``) with
# stash-local ids; the frontend remaps to fresh ``n_``/``e_`` ids on instantiate.

BUILTIN_STASHES: list[dict[str, Any]] = [
    {
        "id": "builtin_clip_word",
        "name": "Transcribe → Find Word → Make Clips → Viewer",
        "builtin": True,
        "createdAt": "",
        "nodes": [
            {
                "id": "s1",
                "type": "video_source",
                "params": {"participant": ""},
                "position": {"x": 40, "y": 160},
            },
            {
                "id": "s2",
                "type": "transcribe",
                "params": {"language": "auto"},
                "position": {"x": 340, "y": 80},
            },
            {
                "id": "s3",
                "type": "find_word",
                "params": {"word": "", "pad": 2},
                "position": {"x": 640, "y": 80},
            },
            {
                "id": "s4",
                "type": "make_clips",
                "params": {
                    "description": "",
                    "output_format": "clip",
                    "titlecards": False,
                    "titlecard_duration": config.TITLECARD_DURATION_SECONDS,
                },
                "position": {"x": 940, "y": 160},
            },
            {
                "id": "s5",
                "type": "timeline_viewer",
                "params": {},
                "position": {"x": 1240, "y": 160},
            },
        ],
        "edges": [
            {
                "id": "se1",
                "from": "s1",
                "fromPort": "video",
                "to": "s2",
                "toPort": "video",
            },
            {
                "id": "se2",
                "from": "s2",
                "fromPort": "segments",
                "to": "s3",
                "toPort": "segments",
            },
            {
                "id": "se3",
                "from": "s3",
                "fromPort": "timeRange",
                "to": "s4",
                "toPort": "timeRange",
            },
            {
                "id": "se4",
                "from": "s1",
                "fromPort": "video",
                "to": "s4",
                "toPort": "video",
            },
            {
                "id": "se5",
                "from": "s4",
                "fromPort": "artifacts",
                "to": "s5",
                "toPort": "artifacts",
            },
        ],
    },
    {
        "id": "builtin_highlights_reel",
        "name": "Sheet Selection → Highlights → Build Reel → Viewer",
        "builtin": True,
        "createdAt": "",
        "nodes": [
            {
                "id": "s1",
                "type": "sheet_selection",
                "params": {"selector": ""},
                "position": {"x": 40, "y": 160},
            },
            {
                "id": "s2",
                "type": "highlights",
                "params": {"budget": config.HIGHLIGHTS_REEL_DURATION_SECONDS},
                "position": {"x": 340, "y": 160},
            },
            {
                "id": "s3",
                "type": "build_reel",
                "params": {"name": "reel"},
                "position": {"x": 640, "y": 160},
            },
            {
                "id": "s4",
                "type": "timeline_viewer",
                "params": {},
                "position": {"x": 940, "y": 160},
            },
        ],
        "edges": [
            {
                "id": "se1",
                "from": "s1",
                "fromPort": "clips",
                "to": "s2",
                "toPort": "clips",
            },
            {
                "id": "se2",
                "from": "s2",
                "fromPort": "clips",
                "to": "s3",
                "toPort": "clips",
            },
            {
                "id": "se3",
                "from": "s3",
                "fromPort": "artifacts",
                "to": "s4",
                "toPort": "artifacts",
            },
        ],
    },
    {
        # Demonstrates the collection-algebra family (limit_events): keep only the
        # top-N most confident detections before cutting clips.
        "id": "builtin_top_detections",
        "name": "Detect → Top Events → Make Clips → Viewer",
        "builtin": True,
        "createdAt": "",
        "nodes": [
            {
                "id": "s1",
                "type": "video_source",
                "params": {"participant": ""},
                "position": {"x": 40, "y": 160},
            },
            {
                "id": "s2",
                "type": "detect",
                "params": {"detector": "text"},
                "position": {"x": 340, "y": 80},
            },
            {
                "id": "s3",
                "type": "limit_events",
                "params": {"sort_by": "confidence", "order": "desc", "take": 10},
                "position": {"x": 640, "y": 80},
            },
            {
                "id": "s4",
                "type": "make_clips",
                "params": {
                    "description": "",
                    "output_format": "clip",
                    "titlecards": False,
                    "titlecard_duration": config.TITLECARD_DURATION_SECONDS,
                },
                "position": {"x": 940, "y": 160},
            },
            {
                "id": "s5",
                "type": "timeline_viewer",
                "params": {},
                "position": {"x": 1240, "y": 160},
            },
        ],
        "edges": [
            {
                "id": "se1",
                "from": "s1",
                "fromPort": "video",
                "to": "s2",
                "toPort": "video",
            },
            {
                "id": "se2",
                "from": "s2",
                "fromPort": "events",
                "to": "s3",
                "toPort": "in",
            },
            {
                "id": "se3",
                "from": "s3",
                "fromPort": "out",
                "to": "s4",
                "toPort": "clips",
            },
            {
                "id": "se4",
                "from": "s1",
                "fromPort": "video",
                "to": "s4",
                "toPort": "video",
            },
            {
                "id": "se5",
                "from": "s4",
                "fromPort": "artifacts",
                "to": "s5",
                "toPort": "artifacts",
            },
        ],
    },
    {
        "id": "builtin_transcript_exports",
        "name": "Transcribe → Transcript + Data Export",
        "builtin": True,
        "createdAt": "",
        "nodes": [
            {
                "id": "s1",
                "type": "video_source",
                "params": {"participant": ""},
                "position": {"x": 40, "y": 160},
            },
            {
                "id": "s2",
                "type": "transcribe",
                "params": {"language": "auto"},
                "position": {"x": 340, "y": 160},
            },
            {
                "id": "s3",
                "type": "transcript_export",
                "params": {"format": config.TRANSCRIBE_FORMAT},
                "position": {"x": 640, "y": 80},
            },
            {
                "id": "s4",
                "type": "data_export",
                "params": {"format": "both"},
                "position": {"x": 640, "y": 240},
            },
        ],
        "edges": [
            {
                "id": "se1",
                "from": "s1",
                "fromPort": "video",
                "to": "s2",
                "toPort": "video",
            },
            {
                "id": "se2",
                "from": "s2",
                "fromPort": "transcript",
                "to": "s3",
                "toPort": "transcript",
            },
            {
                "id": "se3",
                "from": "s2",
                "fromPort": "segments",
                "to": "s4",
                "toPort": "segments",
            },
        ],
    },
]


# ---------------------------------------------------------------------------
# Executors (M3) — thin adapters over existing pure functions
# ---------------------------------------------------------------------------
#
# Each executor has the uniform shape ``execute(ctx, inputs, params) -> {port:
# value}`` (keyed by OUTPUT-port name). Backend modules are imported lazily
# inside each executor (mirrors ``cli._run_ss_clips``) to avoid import cost and
# cycles — Workflows sits at the top of the dependency DAG. The concrete value
# carried on each wire is documented in ``plans/archive/WORKFLOWS-PLAN.md``; the unifying
# primitive is a "source descriptor" embedded in every domain value so the pure
# ``ADAPTERS`` (value -> value, no ctx/params) can still reach a clip's source.

_DEFAULT_EVENT_CLUSTER_GAP = 5.0  # seconds; matches the CLI --cluster-gap default


def _study_from_filename(filename: str) -> str:
    """Derive the study name from a ``{study}_{pid}`` basename ('' when absent)."""
    head, _sep, _tail = Path(filename).stem.rpartition("_")
    return head


def _source_descriptor(participant: str, video_paths: list[str]) -> dict[str, Any]:
    """Build the source descriptor embedded in every domain value."""
    first = Path(video_paths[0]).name if video_paths else ""
    return {
        "participant": participant,
        "study": _study_from_filename(first),
        "source_filename": first,
        "video_paths": list(video_paths),
    }


def _source_from_events(events: list[dict[str, Any]]) -> dict[str, Any]:
    """Best-effort source descriptor derived from a screenspace event list."""
    if not events:
        return {}
    first = events[0]
    source_video = str(first.get("source_video", "") or "")
    return {
        "participant": str(first.get("participant", "") or ""),
        "study": _study_from_filename(source_video),
        "source_filename": source_video,
        "video_paths": [source_video] if source_video else [],
    }


def _clip_source_filename(source: dict[str, Any]) -> str:
    """Source-video override for clip cutting (``+``-joined for multi-part).

    ``pipeline._check_source_video`` resolves a ``" + "``-joined override into all
    parts (building the clip's ``source_timeline``); a single basename resolves by
    exact match. So a multi-video participant flows through clip cutting with the
    right global-time mapping without any timeline code here.
    """
    paths = source.get("video_paths") or []
    if len(paths) >= 2:
        return " + ".join(Path(p).name for p in paths)
    return str(source.get("source_filename", "") or "")


# ---- Sources ----


def _exec_video_source(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    participant = str(params.get("participant", "") or "")
    video_paths = ctx.resolve_videos(participant) if participant else []
    return {
        "video": _source_descriptor(participant, video_paths),
        "participant": participant,
    }


def _exec_sheet_selection(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import spreadsheet

    selector = str(params.get("selector", "") or "").strip()
    if ctx.sheet_context is None or not selector:
        return {"clips": {"records": [], "study": ""}}
    records = spreadsheet.generate_list(
        ctx.worksheet,
        "reel",
        ctx=ctx.sheet_context,
        reel_input=selector,
        skip_prompts=True,
    )
    study = str(getattr(ctx.sheet_context, "study_name", "") or "")
    return {"clips": {"records": records, "study": study}}


def _exec_region(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import screenspace

    name = str(params.get("name", "") or "").strip()
    coords: dict[str, Any] | None = None
    if name:
        regions = screenspace.load_screenspace_manifest().get("regions") or {}
        entry = regions.get(name)
        if isinstance(entry, dict):
            coords = {k: entry[k] for k in ("x", "y", "w", "h") if k in entry}
    return {"region": {"name": name, "coords": coords}}


def _exec_time_range(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    """Manual in/out times → timeRange (scan windows for SS nodes, cuts for clips).

    Parses the same ``MM:SS``/``HH:MM:SS`` (range or single) syntax as a sheet
    cell. Source is left empty — SS detectors take their video from the ``video``
    input, and ``make_clips`` falls back to its wired ``video`` for the source.
    """
    raw = str(params.get("ranges", "") or "").strip()
    ranges: list[tuple[float, float]] = []
    for start_str, end_str in utils.parse_timestamps(raw) if raw else []:
        start = utils.timestamp_to_seconds(start_str)
        end = utils.timestamp_to_seconds(end_str)
        if start is not None and end is not None:
            ranges.append((start, max(start, end)))
    return {"timeRange": {"ranges": ranges, "source": {}}}


# ---- Transcript ----


def _exec_transcribe(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import transcripts
    import video

    src = inputs.get("video") or {}
    paths = list(src.get("video_paths") or [])
    language = str(params.get("language", "") or "").strip()
    lang = None if language in ("", "auto") else language
    model_name = str(params.get("model", "") or "").strip() or None

    result: Any = None
    if len(paths) >= 2:
        timeline = video.build_source_timeline(paths)
        if timeline is not None:
            result = transcripts.transcribe_timeline(
                timeline,
                model_name=model_name,
                language=lang,
                cancel_flag=ctx.cancel_flag,
            )
    elif paths:
        result = transcripts.transcribe_video(
            paths[0],
            model_name=model_name,
            language=lang,
            cancel_flag=ctx.cancel_flag,
        )

    if result is None:
        result = {
            "segments": [],
            "language": lang or "",
            "source_file": paths[0] if paths else "",
            "model": "",
        }
    transcript_val = dict(result)
    transcript_val["source"] = src
    return {
        "transcript": transcript_val,
        "segments": {"segments": result.get("segments", []), "source": src},
    }


def _exec_find_word(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    seg_val = inputs.get("segments") or {}
    segments = seg_val.get("segments") or []
    source = seg_val.get("source") or {}
    word = str(params.get("word", "") or "").strip().lower()
    pad = float(params.get("pad", 0) or 0)

    ranges: list[tuple[float, float]] = []
    times: list[float] = []
    if word:
        for seg in segments:
            if word in str(seg.get("text", "")).lower():
                start = float(seg.get("start", 0.0) or 0.0)
                end = float(seg.get("end", start) or start)
                ranges.append((max(0.0, start - pad), end + pad))
                times.append(start)
    return {
        "timeRange": {"ranges": ranges, "source": source},
        "timestamps": {"times": times, "source": source},
    }


def _exec_transcript_export(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import transcripts

    transcript_in = inputs.get("transcript") or {}
    seg_in = inputs.get("segments") or {}
    src = transcript_in.get("source") or seg_in.get("source") or {}
    study = str(src.get("study", "") or "")
    fmt = str(params.get("format", "") or "") or config.TRANSCRIBE_FORMAT
    if fmt not in ("md", "srt", "vtt"):
        fmt = "md"

    # Prefer the full transcript (carries language/model); a bare segments wire
    # still exports, just with empty metadata.
    if transcript_in.get("segments"):
        base: dict[str, Any] = transcript_in
    elif seg_in.get("segments"):
        base = {
            "segments": seg_in.get("segments"),
            "language": "",
            "model": "",
            "source_file": str(src.get("source_filename", "") or ""),
        }
    else:
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "No transcript or segments wired",
        }
    result = cast(
        transcripts.TranscriptResult,
        {
            "segments": list(base.get("segments") or []),
            "language": str(base.get("language", "") or ""),
            "model": str(base.get("model", "") or ""),
            "source_file": str(base.get("source_file", "") or ""),
        },
    )

    participant = str(src.get("participant", "") or "")
    stem = f"transcript_{participant}" if participant else "transcript"
    ext = transcripts.get_transcript_extension(fmt)
    output_path = files.get_unique_filename(f"{stem}{ext}", file_format=ext)
    if not transcripts.write_transcript(result, output_path, fmt=fmt):
        files.release_reservation(output_path)
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "Transcript couldn't be written",
        }
    # "export" (not "transcript") — the viewer routes it to the Attachments
    # pane's document card; "transcript" is a timeline card type there.
    rec = _attachment_artifact("export", output_path, src, f"Transcript ({fmt})")
    return {"artifacts": {"artifacts": [rec], "study": study, "count": 1}}


# ---- Thinking (Ollama) ----


def _exec_summarize(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import ollama_client
    import thinking_agents

    transcript = inputs.get("transcript") or {}
    segments = transcript.get("segments") or []
    if not ollama_client.is_available():
        return {"summary": "", "__note__": "Ollama not available. Summary skipped"}
    summary = thinking_agents.summarize_transcript(
        segments, model=params.get("model") or None, cancel_event=ctx.cancel_event
    )
    return {"summary": summary or ""}


def _exec_citations(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import ollama_client
    import thinking_agents

    summary = str(inputs.get("summary") or "")
    seg_val = inputs.get("segments") or {}
    segments = seg_val.get("segments") or []
    if not ollama_client.is_available():
        return {"citations": [], "__note__": "Ollama not available. Citations skipped"}
    cites = thinking_agents.find_citations(
        summary,
        segments,
        model=params.get("model") or None,
        cancel_event=ctx.cancel_event,
    )
    return {"citations": cites or []}


def _exec_friction(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import friction
    import ollama_client
    import thinking_agents

    seg_val = inputs.get("segments") or {}
    segments = seg_val.get("segments") or []
    summary = str(inputs.get("summary") or "")
    if not ollama_client.is_available():
        return {"friction": [], "__note__": "Ollama not available. Friction skipped"}
    scored = friction.score_segments(segments)
    candidates = friction.select_candidates(scored)
    moments = thinking_agents.find_friction_moments(
        summary,
        segments,
        candidates,
        model=params.get("model") or None,
        cancel_event=ctx.cancel_event,
    )
    return {"friction": moments or []}


# ---- Screenspace ----


def _build_ss_scan_params(tool_name: str, params: dict[str, Any]) -> dict[str, Any]:
    """Assemble a detector's flat node params into the scan's ``parameters`` dict.

    Flat number/enum/bool params from the canvas are reshaped into the nested
    structure each ``screenspace_tools`` scan reads (e.g. color's ``target_color``
    /``tolerance`` HSV dicts). Reference-frame extraction is handled separately by
    the caller (it needs the video path + region).
    """

    def _num(key: str, default: float = 0.0) -> float:
        val = params.get(key)
        return float(val) if val not in (None, "") else float(default)

    if tool_name == "color":
        return {
            "target_color": {
                "h": _num("color_h"),
                "s": _num("color_s"),
                "v": _num("color_v"),
            },
            "tolerance": {
                "h": _num("tol_h", 10),
                "s": _num("tol_s", 50),
                "v": _num("tol_v", 50),
            },
            "color_mode": str(params.get("color_mode", "average") or "average"),
            "min_coverage": _num("min_coverage"),
            "interval": _num("interval"),
        }
    if tool_name == "change":
        return {
            "threshold": _num("threshold"),
            "noise_threshold": _num("noise_threshold"),
            "require_consecutive": int(_num("require_consecutive", 1)),
            "interval": _num("interval"),
        }
    if tool_name == "flow":
        return {
            "magnitude_threshold": _num("magnitude_threshold"),
            "require_consecutive": int(_num("require_consecutive", 1)),
            "interval": _num("interval"),
        }
    if tool_name == "text":
        return {
            "search_string": str(params.get("search_string", "") or ""),
            "fuzzy_threshold": _num("fuzzy_threshold", 80),
            "interval": _num("interval", 2.0),
        }
    if tool_name == "numbers":
        out: dict[str, Any] = {
            "operator": str(params.get("operator", "gt") or "gt"),
            "target_value": _num("target_value"),
            "integers_only": bool(params.get("integers_only", False)),
            "interval": _num("interval", 2.0),
        }
        if params.get("range_min") not in (None, ""):
            out["range_min"] = _num("range_min")
        if params.get("range_max") not in (None, ""):
            out["range_max"] = _num("range_max")
        return out
    if tool_name == "similarity":
        return {"threshold": _num("threshold"), "interval": _num("interval")}
    if tool_name == "scene":
        return {"threshold": _num("threshold"), "interval": _num("interval")}
    if tool_name == "template":
        return {
            "threshold": _num("threshold"),
            "template_scale": _num("template_scale", 1.0),
            "interval": _num("interval"),
        }
    if tool_name == "inactivity":
        return {
            "threshold": _num("threshold"),
            "min_duration": _num("min_duration"),
            "interval": _num("interval"),
        }
    if tool_name == "boundary":
        return {
            "threshold": _num("threshold"),
            "min_gap": _num("min_gap"),
            "metric": str(
                params.get("metric", "") or config.SCREENSPACE_BOUNDARY_METRIC
            ),
            "interval": _num("interval"),
        }
    return {"interval": _num("interval")}


def _attach_ss_reference(
    tool_name: str,
    base_params: dict[str, Any],
    params: dict[str, Any],
    video_path: str,
    region_coords: dict[str, int],
) -> bool:
    """Self-extract a reference frame from ``video_path`` at ``reference_seconds``.

    Similarity/scene/template need reference image data that the canvas can't
    upload; instead we crop the node's region from a frame at ``reference_seconds``.
    Returns False when the frame can't be read so the executor can short-circuit.
    """
    import screenspace
    import video as video_mod

    ref_ts = float(params.get("reference_seconds", 0.0) or 0.0)
    frame = video_mod.extract_frame_at_timestamp(video_path, ref_ts)
    if frame is None:
        return False
    crop = screenspace.extract_region(frame, region_coords)
    if tool_name == "similarity":
        base_params["reference_frame"] = crop
    elif tool_name == "template":
        base_params["template_image"] = crop
    elif tool_name == "scene":
        base_params["reference_scenes"] = [{"name": "ref", "frame": crop}]
    return True


def _run_ss_detector(
    ctx: NodeContext,
    inputs: dict[str, Any],
    params: dict[str, Any],
    tool_name: str,
) -> dict[str, Any]:
    import screenspace
    import screenspace_manifest
    import screenspace_worker

    src = inputs.get("video") or {}
    paths = list(src.get("video_paths") or [])
    tool = screenspace.TOOLS.get(tool_name)
    if not paths or tool is None:
        note = "No video wired" if not paths else f"Unknown detector: {tool_name}"
        return {
            "events": {"events": [], "source": src, "raw_results": []},
            "__note__": note,
        }

    # Unwired region scans the whole frame (zero-size coords would make the scan a
    # silent no-op — see _resolve_region_coords).
    region_name, region_coords = _resolve_region_coords(
        inputs.get("region") or {}, paths[0]
    )

    base_params = _build_ss_scan_params(tool_name, params)
    if tool_name in _SS_REFERENCE_DETECTORS and not _attach_ss_reference(
        tool_name, base_params, params, paths[0], region_coords
    ):
        return {
            "events": {"events": [], "source": src, "raw_results": []},
            "__note__": "Couldn't read the reference frame at the given time",
        }

    task = screenspace_manifest.create_task(
        tool_name,
        str(src.get("participant", "") or ""),
        str(src.get("source_filename", "") or ""),
        paths,
        region_name,
        region_coords,
        parameters=base_params,
    )

    # Scan windows: each timeRange span scanned separately, else the whole video.
    windows = list((inputs.get("timeRange") or {}).get("ranges") or [])
    raw_results: list[dict[str, Any]] = []
    scan_targets = windows or [None]
    # Each underlying scan reports progress on its own 0->1 scale (multi-video
    # scans even force on_progress(1.0) at the end), so map each window into its
    # span-weighted slice of the job to keep job-level progress monotonic.
    spans = [
        max(0.0, float(w[1]) - float(w[0])) if w is not None else 1.0
        for w in scan_targets
    ]
    total_span = sum(spans)
    if total_span <= 0:
        spans = [1.0] * len(scan_targets)
        total_span = float(len(scan_targets))
    done_span = 0.0
    for window, win_span in zip(scan_targets, spans):
        if ctx.cancel_flag():
            break
        scan_params = dict(task["parameters"])
        if window is not None:
            scan_params["start_seconds"] = float(window[0])
            scan_params["end_seconds"] = float(window[1])
        frac_start = done_span / total_span
        frac_end = (done_span + win_span) / total_span

        def window_progress(
            p: float, _a: float = frac_start, _b: float = frac_end
        ) -> None:
            ctx.on_progress(_a + p * (_b - _a))

        returned = screenspace_worker.dispatch_tool_scan(
            tool,
            paths,
            region_coords,
            scan_params,
            task_id=task["id"],
            scan_mode="normal",
            on_progress=window_progress,
            cancel_flag=ctx.cancel_flag,
            on_result=None,
            fast_opts=None,
        )
        raw_results.extend(returned or [])
        done_span += win_span

    events = screenspace_manifest.generate_events_from_results(task, raw_results)
    # raw_results rides along for the heatmap node (template/flow/change); every
    # other consumer reads only ``events`` and ignores it.
    return {"events": {"events": events, "source": src, "raw_results": raw_results}}


def _make_ss_executor(
    tool_name: str,
) -> Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]:
    """Bind ``_run_ss_detector`` to one tool so all ten nodes share one body."""

    def _exec(
        ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
    ) -> dict[str, Any]:
        return _run_ss_detector(ctx, inputs, params, tool_name)

    return _exec


def _exec_detect(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    """Unified detector node: dispatch to the chosen tool's scan body."""
    tool_name = str(params.get("detector", "text") or "text")
    return _run_ss_detector(ctx, inputs, params, tool_name)


def _resolve_region_coords(
    region_in: dict[str, Any], video_path: str
) -> tuple[str, dict[str, int]]:
    """Resolve a region input to pixel coords, defaulting to the **full frame**.

    A region port is optional on every Screenspace node; when it is unwired (or a
    Region node names nothing / a coord-less entry), this returns the whole
    frame's pixel dimensions rather than zero-size coords. That matters because
    ``scan_video_frames`` rejects a zero-size region and skips the scan entirely
    (a silent no-op), and the reference detectors would otherwise crop an empty
    reference frame. Probes ``video_path`` once for the frame size.
    """
    import screenspace
    import video

    region_name = str(region_in.get("name", "") or "")
    props = video.probe_video_properties(video_path) or {}
    width = int(props.get("width", 0) or 0)
    height = int(props.get("height", 0) or 0)
    coords: dict[str, int] = {"x": 0, "y": 0, "w": 0, "h": 0}
    norm = region_in.get("coords")
    if isinstance(norm, dict) and norm and width > 0 and height > 0:
        coords = screenspace.denormalize_region(norm, width, height)
    if coords["w"] <= 0 or coords["h"] <= 0:
        # Full-frame fallback: unwired region, or a region that didn't resolve.
        coords = {"x": 0, "y": 0, "w": width, "h": height}
    return region_name, coords


def _exec_multitool(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import screenspace
    import screenspace_manifest
    import screenspace_worker

    src = inputs.get("video") or {}
    paths = list(src.get("video_paths") or [])
    raw_steps = list(params.get("steps") or [])
    if not paths or len(raw_steps) < 2:
        return {"events": {"events": [], "source": src, "raw_results": []}}

    region_name, region_coords = _resolve_region_coords(
        inputs.get("region") or {}, paths[0]
    )

    # Reshape each flat step into the {type, region_coords, logic, …} shape
    # scan_multitool expects, reusing the per-detector param builder.
    steps: list[dict[str, Any]] = []
    for idx, raw in enumerate(raw_steps):
        step_type = str(raw.get("type", "") or "")
        if step_type not in screenspace.TOOLS:
            continue
        step = _build_ss_scan_params(step_type, raw)
        step["type"] = step_type
        step["region_coords"] = region_coords
        if idx > 0:
            step["logic"] = str(raw.get("logic", "AND") or "AND").upper()
        steps.append(step)
    if len(steps) < 2:
        return {"events": {"events": [], "source": src, "raw_results": []}}

    task = screenspace_manifest.create_task(
        "multitool",
        str(src.get("participant", "") or ""),
        str(src.get("source_filename", "") or ""),
        paths,
        region_name,
        region_coords,
        parameters={"steps": steps},
    )
    raw_results = (
        screenspace_worker.dispatch_tool_scan(
            screenspace.TOOLS["multitool"],
            paths,
            region_coords,
            task["parameters"],
            task_id=task["id"],
            scan_mode="normal",
            on_progress=ctx.on_progress,
            cancel_flag=ctx.cancel_flag,
            on_result=None,
            fast_opts=None,
        )
        or []
    )
    events = screenspace_manifest.generate_events_from_results(task, raw_results)
    return {"events": {"events": events, "source": src, "raw_results": raw_results}}


# ---- Artifact ----


def _exec_highlights(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import spreadsheet

    clips_in = inputs.get("clips") or {}
    records = list(clips_in.get("records") or [])
    study = str(clips_in.get("study", "") or "")
    if not records:
        return {"clips": {"records": [], "study": study}}
    budget = int(params.get("budget", config.HIGHLIGHTS_REEL_DURATION_SECONDS) or 0)
    if budget <= 0:
        budget = config.HIGHLIGHTS_REEL_DURATION_SECONDS
    # Uniqueness is scored against artifacts already in the output dir (mirrors the
    # ``-H`` CLI path: spreadsheet.generate_reel_timestamps -> score_and_truncate_clips).
    existing_filenames = set(files.discover_clips())
    selected = spreadsheet.score_and_truncate_clips(records, existing_filenames, budget)
    return {"clips": {"records": selected, "study": study}}


def _artifact_padding_params(params: dict[str, Any]) -> tuple[float, float, float]:
    """Read the pad-start/pad-end/max-duration node params for the pipeline.

    Returns ``(pad_pre, pad_post, max_duration)``; all default to a no-op (0.0).
    Pads are signed (negative = trim inward). Shared by the Make Clips and Build
    Reel executors.
    """

    def _num(key: str) -> float:
        try:
            return float(params.get(key) or 0)
        except (TypeError, ValueError):
            return 0.0

    return _num("pad_start"), _num("pad_end"), _num("max_duration")


def _exec_make_clips(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import pipeline

    description = str(params.get("description", "") or "").strip() or "workflow"
    output_format = str(params.get("output_format", "clip") or "clip")
    if output_format not in ("clip", "screen", "gif"):
        output_format = "clip"
    titlecards = bool(params.get("titlecards", False))
    raw_card_dur = params.get("titlecard_duration")
    titlecard_duration = int(raw_card_dur) if raw_card_dur not in (None, "") else None

    records: list[Any] = []
    study = ""
    clips_in = inputs.get("clips")
    if isinstance(clips_in, dict) and clips_in.get("records"):
        records = list(clips_in["records"])
        study = str(clips_in.get("study", "") or "")
    else:
        tr = inputs.get("timeRange")
        if isinstance(tr, dict) and tr.get("ranges"):
            source = tr.get("source") or inputs.get("video") or {}
            study = str(source.get("study", "") or "")
            records = files.build_clip_records(
                participant=str(source.get("participant", "") or ""),
                source_filename=_clip_source_filename(source),
                time_ranges=[(float(s), float(e)) for s, e in tr["ranges"]],
                description=description,
                study=study,
            )

    if not records:
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "No clips to render. Wire clips, a time range, or a video",
        }
    pad_pre, pad_post, max_duration = _artifact_padding_params(params)
    count, artifacts = pipeline.process_clips(
        records,
        output_format=output_format,
        include_severity=False,
        cancel_flag=ctx.cancel_flag,
        titlecards_enabled=titlecards,
        titlecard_duration_seconds=titlecard_duration,
        pad_pre=pad_pre,
        pad_post=pad_post,
        max_duration=max_duration,
    )
    return {"artifacts": {"artifacts": artifacts, "study": study, "count": count}}


def _exec_interval_captures(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    """Sample a video into screenshots/GIFs at a fixed interval.

    Iterates each wired time range (or the whole video when none is wired) at
    ``interval`` seconds, expands the samples into point/GIF clip records, and
    reuses ``process_clips`` so the artifacts match Make Clips exactly. Fixes the
    one-artifact-per-range limit of Make Clips' screen/gif output.
    """
    import files
    import pipeline
    import video as video_mod

    src = inputs.get("video") or {}
    paths = list(src.get("video_paths") or [])
    study = str(src.get("study", "") or "")
    empty = {"artifacts": {"artifacts": [], "study": study, "count": 0}}
    if not paths:
        return {**empty, "__note__": "No video wired"}

    interval = int(
        float(params.get("interval", config.GALLERY_INTERVAL_SECONDS) or 0)
        or config.GALLERY_INTERVAL_SECONDS
    )
    if interval < 1:
        interval = 1
    fmt = (
        "gif" if str(params.get("output_format", "screen") or "") == "gif" else "screen"
    )
    gif_dur = float(
        params.get("gif_duration", config.GALLERY_GIF_DURATION_SECONDS)
        or config.GALLERY_GIF_DURATION_SECONDS
    )

    ranges = [
        (float(s), float(e))
        for s, e in ((inputs.get("timeRange") or {}).get("ranges") or [])
    ]
    if not ranges:
        duration = video_mod.get_file_duration(paths[0]) or 0
        if duration <= 0:
            return {**empty, "__note__": "Couldn't read the video duration"}
        ranges = [(0.0, float(duration))]

    # Expand each window into per-interval sample points (a point for a
    # screenshot, a [t, t+gif_dur] window for a GIF).
    sample_ranges: list[tuple[float, float]] = []
    for start, end in ranges:
        t = start
        while t < end:
            sample_ranges.append((t, t + gif_dur if fmt == "gif" else t))
            t += interval
    if not sample_ranges:
        return {**empty, "__note__": "No sample points in the given interval/range"}

    records = files.build_clip_records(
        participant=str(src.get("participant", "") or ""),
        source_filename=_clip_source_filename(src),
        time_ranges=sample_ranges,
        description="sample",
        study=study,
    )
    count, artifacts = pipeline.process_clips(
        records,
        output_format=fmt,
        include_severity=False,
        cancel_flag=ctx.cancel_flag,
    )
    return {"artifacts": {"artifacts": artifacts, "study": study, "count": count}}


def _attachment_artifact(
    art_type: str, output_path: str, source: dict[str, Any], description: str
) -> dict[str, Any]:
    """Build an artifact record for a single non-timeline output (timelapse/heatmap).

    ``start``/``end`` are 0 so the viewer keeps it out of the timeline track and
    surfaces it in the Attachments panel instead (branch on ``type``).
    """
    name = Path(output_path).name
    return {
        "id": f"{art_type}-{name}",
        "type": art_type,
        "file": name,
        "thumbnail": "",
        "start": 0,
        "end": 0,
        "study": str(source.get("study", "") or ""),
        "participant": str(source.get("participant", "") or ""),
        "category": "",
        "severity": "",
        "description": description,
        "sourceVideo": str(source.get("source_filename", "") or ""),
        "annotations": [],
    }


def _exec_timelapse(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import screenspace_scans

    src = inputs.get("video") or {}
    paths = list(src.get("video_paths") or [])
    study = str(src.get("study", "") or "")
    if not paths:
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "No video wired",
        }

    # _resolve_region_coords already falls back to the full frame when no region
    # is wired; a still-zero size means the probe failed (unreadable video).
    _name, region_coords = _resolve_region_coords(inputs.get("region") or {}, paths[0])
    if region_coords["w"] <= 0 or region_coords["h"] <= 0:
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "Couldn't read the video",
        }

    out_format = str(params.get("output_format", "mp4") or "mp4")
    if out_format not in ("mp4", "gif"):
        out_format = "mp4"
    output_path = files.get_unique_filename(
        f"timelapse.{out_format}", file_format=f".{out_format}"
    )
    result = screenspace_scans.generate_timelapse(
        paths[0],
        region_coords,
        float(params.get("speedup_factor", 10.0) or 10.0),
        output_path,
        output_format=out_format,
        sample_interval=float(params.get("sample_interval", 0.0) or 0.0),
        on_progress=ctx.on_progress,
        cancel_flag=ctx.cancel_flag,
    )
    if not result:
        files.release_reservation(output_path)
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "Timelapse couldn't be generated",
        }
    rec = _attachment_artifact("timelapse", result, src, "Timelapse")
    return {"artifacts": {"artifacts": [rec], "study": study, "count": 1}}


def _exec_heatmap(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import screenspace_heatmap
    import video

    events_in = inputs.get("events") or {}
    src = events_in.get("source") or {}
    study = str(src.get("study", "") or "")
    results = list(events_in.get("raw_results") or [])
    style = str(params.get("style", "change") or "change")
    paths = list(src.get("video_paths") or [])
    if not results or not paths or style not in ("template", "flow", "change"):
        note = (
            "No detector results. Wire a matching template/flow/change detector"
            if not results
            else "No video for the heatmap"
        )
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": note,
        }

    props = video.probe_video_properties(paths[0]) or {}
    width = int(props.get("width", 0) or 0) or 1920
    height = int(props.get("height", 0) or 0) or 1080
    output = str(params.get("output", "image") or "image")
    if output not in ("image", "gif", "rolling_gif"):
        output = "image"
    if output == "image":
        output_path = files.get_unique_filename("heatmap.png", file_format=".png")
        if style == "template":
            result = screenspace_heatmap.generate_template_heatmap(
                results, width, height, output_path
            )
        elif style == "flow":
            result = screenspace_heatmap.generate_flow_heatmap(
                results, width, height, output_path
            )
        else:
            result = screenspace_heatmap.generate_change_heatmap(
                results, width, height, output_path
            )
        failure_note = "Heatmap couldn't be generated"
    else:
        output_path = files.get_unique_filename("heatmap.gif", file_format=".gif")
        num_frames = int(float(params.get("frames", 24) or 24))
        if output == "gif":
            result = screenspace_heatmap.generate_heatmap_gif(
                results,
                width,
                height,
                output_path,
                heatmap_type=style,
                num_frames=num_frames,
            )
        else:
            result = screenspace_heatmap.generate_rolling_heatmap_gif(
                results,
                width,
                height,
                output_path,
                heatmap_type=style,
                num_frames=num_frames,
                window_frames=int(float(params.get("window", 6) or 6)),
            )
        failure_note = "Not enough detector results for an animated heatmap"
    if not result:
        files.release_reservation(output_path)
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": failure_note,
        }
    rec = _attachment_artifact("heatmap", result, src, f"{style.title()} heatmap")
    return {"artifacts": {"artifacts": [rec], "study": study, "count": 1}}


def _reel_start_seconds(rec: Any) -> float:
    """Earliest start (seconds) of a clip record, for chronological reels.

    Adapter-built records carry pre-resolved ``times`` (H:MM:SS); sheet records
    resolve lazily from their ``cell`` (matching
    ``spreadsheet.sort_clips_chronologically``). Unparseable records sort last.
    """
    times = rec.get("times")
    if not times:
        cell = rec.get("cell")
        cell_value = str(getattr(cell, "value", "") or "") if cell is not None else ""
        cleaned, _, _ = utils.parse_cell_annotations(cell_value)
        times = utils.parse_timestamps(cleaned)
    if not times:
        return float("inf")
    seconds = utils.timestamp_to_seconds(times[0][0])
    return seconds if seconds is not None else float("inf")


def _exec_build_reel(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import pipeline

    clips_in = inputs.get("clips") or {}
    records = list(clips_in.get("records") or [])
    study = str(clips_in.get("study", "") or "")
    if params.get("chronological"):
        records = sorted(records, key=_reel_start_seconds)
    if not records:
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "manifest": {"path": None, "records": []},
            "__note__": "No clips to build a reel from",
        }
    # Honor the node's reel name: reserve a unique output path (process_reel
    # treats a supplied output_file as a reservation and releases it on failure).
    name = utils.sanitize_filename(str(params.get("name", "") or "").strip()) or "reel"
    output_file = files.get_unique_filename(f"{name}{config.FILEFORMAT}")
    pad_pre, pad_post, max_duration = _artifact_padding_params(params)
    count, reels = pipeline.process_reel(
        records,
        output_file=output_file,
        cancel_flag=ctx.cancel_flag,
        pad_pre=pad_pre,
        pad_post=pad_post,
        max_duration=max_duration,
    )
    return {
        "artifacts": {"artifacts": reels, "study": study, "count": count},
        "manifest": {"path": None, "records": reels},
    }


def _exec_data_export(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import data_export
    import files

    events_in = inputs.get("events") or {}
    seg_in = inputs.get("segments") or {}
    src = events_in.get("source") or seg_in.get("source") or {}
    study = str(src.get("study", "") or "")
    fmt = str(params.get("format", "both") or "both")
    if fmt not in ("both", "json", "csv"):
        fmt = "both"

    participant = str(src.get("participant", "") or "")
    suffix = f"_{participant}" if participant else ""
    # (stem, rows, preferred CSV column order, description) per wired surface.
    surfaces: list[tuple[str, list[dict[str, Any]], tuple[str, ...], str]] = []
    events = list(events_in.get("events") or [])
    if events:
        surfaces.append(
            (
                f"export_events{suffix}",
                data_export.build_screenspace_events(
                    {"events": events}, include_excluded=True
                ),
                data_export.SCREENSPACE_EVENT_COLUMNS,
                "Events export",
            )
        )
    segments = list(seg_in.get("segments") or [])
    if segments:
        # build_transcript_segments reads a transcripts-manifest envelope;
        # synthesize one around the wired segments.
        manifest = {
            "source_transcripts": {
                (participant or "unknown"): {
                    "segments": segments,
                    "source_file": str(src.get("source_filename", "") or ""),
                }
            }
        }
        surfaces.append(
            (
                f"export_segments{suffix}",
                data_export.build_transcript_segments(manifest),
                data_export._TRANSCRIPT_SEGMENT_BASE_COLS,
                "Segments export",
            )
        )
    if not surfaces:
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "No events or segments wired",
        }

    records: list[dict[str, Any]] = []
    for stem, rows, columns, description in surfaces:
        writes: list[tuple[str, str, str]] = []  # (extension, payload, label)
        if fmt in ("both", "json"):
            writes.append((".json", data_export.to_json(rows), "JSON"))
        if fmt in ("both", "csv"):
            writes.append(
                (
                    ".csv",
                    data_export.to_csv(rows, preferred_column_order=columns),
                    "CSV",
                )
            )
        for ext, payload, label in writes:
            output_path = files.get_unique_filename(f"{stem}{ext}", file_format=ext)
            try:
                Path(output_path).write_text(payload, encoding="utf-8")
            except OSError:
                files.release_reservation(output_path)
                return {
                    "artifacts": {"artifacts": [], "study": study, "count": 0},
                    "__note__": "Export couldn't be written",
                }
            records.append(
                _attachment_artifact(
                    "export", output_path, src, f"{description} ({label})"
                )
            )
    return {"artifacts": {"artifacts": records, "study": study, "count": len(records)}}


def _exec_timeline_viewer(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import viewer

    artifacts_in = inputs.get("artifacts") or {}
    incoming = list(artifacts_in.get("artifacts") or [])
    study = str(artifacts_in.get("study", "") or "")
    # build_reel emits reel records (carrying ``components``, no start/end) onto
    # the same ``artifacts`` wire as clip/screen/gif artifacts. The viewer renders
    # them from separate slots — timeline artifacts on the timeline, reels in the
    # Attachments pane — so split them out here (otherwise reels land in the
    # timeline slot, get filtered for lack of start/end, and the viewer is empty).
    reels = [
        a for a in incoming if isinstance(a, dict) and a.get("components") is not None
    ]
    artifacts = [
        a
        for a in incoming
        if not (isinstance(a, dict) and a.get("components") is not None)
    ]
    ss_events = (inputs.get("events") or {}).get("events") or None
    data = viewer.finalize_timeline_data(
        artifacts, reels=reels or None, study=study, screenspace_events=ss_events
    )
    path = viewer.generate_timeline_viewer(data, output_basename="workflow_viewer.html")
    return {"viewer": {"path": str(path) if path else None}}


# ---- Control ----


_GATE_OPS: dict[str, Callable[[float, float], bool]] = {
    ">=": lambda a, b: a >= b,
    ">": lambda a, b: a > b,
    "<=": lambda a, b: a <= b,
    "<": lambda a, b: a < b,
    "==": lambda a, b: a == b,
    "!=": lambda a, b: a != b,
}


def _reduce_collection(metric: str, inputs: dict[str, Any]) -> float:
    """Reduce a wired collection to one scalar (the measure / gate-collection core).

    Reads whichever of events / clipRecords / segments is wired (events first).
    ``max_confidence`` only applies to events; it falls back to 0 otherwise.
    """
    events = (inputs.get("events") or {}).get("events")
    records = (inputs.get("clips") or {}).get("records")
    segments = (inputs.get("segments") or {}).get("segments")

    if events is not None:
        items = list(events)
        if metric == "count":
            return float(len(items))
        if metric == "max_confidence":
            confs = [float(e.get("confidence", 0.0) or 0.0) for e in items]
            return max(confs) if confs else 0.0
        return float(
            sum(
                max(
                    0.0,
                    float(e.get("time_out", 0.0) or 0.0)
                    - float(e.get("time_in", 0.0) or 0.0),
                )
                for e in items
            )
        )

    if records is not None:
        items = list(records)
        if metric == "count":
            return float(len(items))
        if metric == "max_confidence":
            return 0.0
        total = 0.0
        for rec in items:
            for start_str, end_str in rec.get("times") or []:
                start = utils.timestamp_to_seconds(start_str)
                end = utils.timestamp_to_seconds(end_str)
                if start is not None and end is not None:
                    total += max(0.0, end - start)
        return total

    if segments is not None:
        items = list(segments)
        if metric == "count":
            return float(len(items))
        if metric == "max_confidence":
            return 0.0
        return float(
            sum(
                max(
                    0.0,
                    float(s.get("end", 0.0) or 0.0) - float(s.get("start", 0.0) or 0.0),
                )
                for s in items
            )
        )

    return 0.0


def _exec_measure(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    """Reduce a wired collection to one scalar for a downstream gate."""
    metric = str(params.get("metric", "count") or "count")
    return {"value": _reduce_collection(metric, inputs)}


def _apply_gate(value: float, params: dict[str, Any]) -> bool:
    """Compare *value* to a threshold per the node's ``op`` (shared gate logic)."""
    fn = _GATE_OPS.get(str(params.get("op", ">=") or ">="))
    if fn is None:
        return False
    try:
        return bool(fn(value, float(params.get("threshold", 0))))
    except (TypeError, ValueError):
        return False


def _exec_gate(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    raw = inputs.get("value")
    if raw is None:
        return {"pass": False}
    try:
        value = float(raw)
    except (TypeError, ValueError):
        return {"pass": False}
    return {"pass": _apply_gate(value, params)}


def _exec_gate_collection(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    """Reduce a wired collection to a scalar then gate it — the measure+gate pair
    fused into one node (see :func:`_exec_measure` and :func:`_exec_gate`)."""
    metric = str(params.get("metric", "count") or "count")
    value = _reduce_collection(metric, inputs)
    return {"pass": _apply_gate(value, params)}


# ---- Collection-algebra control nodes (filter / merge / partition / limit / dedup) ----
#
# These thin / combine / branch / cap / dedup the collections that already flow
# through the graph (events, clipRecords, segments) — the collections *are* the
# iteration, so no per-item ``foreach`` (which would force runtime DAG expansion,
# breaking the static ``topo_order`` model). All are pure single-pass nodes: no
# runner changes. Per-type families (mirroring the ss_* split) keep every port
# exact-typed, so no adapters or frontend ``canConnect`` changes are needed.

# One ``{field, op, value}`` clause reuses the gate's comparison table; ``contains``
# is the one string-only addition (kept out of ``_GATE_OPS``, which must stay
# numeric for the gate). ``none`` is the limit node's "keep input order" sentinel.
_COLLECTION_OPS: list[str] = list(_GATE_OPS.keys()) + ["contains"]
_SORT_NONE = "none"

# kind -> envelope metadata. ``port`` is the wire type; ``key`` is the inner list
# key in the envelope; ``preserve`` are the envelope keys carried through unchanged
# (source lineage / study / raw_results); ``fields`` drive the predicate enum and
# ``sort_fields`` the limit sort enum (numeric-only, so the sort key stays
# comparable). ``recount`` (artifacts) rewrites the envelope's ``count`` to the
# kept length. dedup is span-based, registered for events + clips + timeRanges.
#
# Both the pre-clip side (``clipRecords``, ``timeRange``) and the post-clip side
# (``artifacts`` from make_clips/timelapse/heatmap/build_reel) get families, so a
# stream can be thinned/capped/combined before *or* after it becomes artifacts.
_COLLECTION_KINDS: dict[str, dict[str, Any]] = {
    "events": {
        "port": "events",
        "key": "events",
        "label": "Events",
        "preserve": ("source", "raw_results"),
        "fields": ["confidence", "duration", "start"],
        "sort_fields": ["confidence", "duration", "start"],
    },
    "clips": {
        "port": "clipRecords",
        # Labelled "Clip Selections" (not "Clips") so the family reads as operating
        # on pre-render clip specs from sheet_selection/highlights — not the
        # rendered ``artifacts`` Make Clips emits (which has its own artifacts
        # family). The node ids stay ``*_clips`` so saved blueprints are unaffected.
        "key": "records",
        "label": "Clip Selections",
        "preserve": ("study",),
        "fields": ["duration", "category", "severity", "desc"],
        "sort_fields": ["duration"],
    },
    "segments": {
        "port": "segments",
        "key": "segments",
        "label": "Segments",
        "preserve": ("source",),
        "fields": ["text", "duration", "start"],
        "sort_fields": ["duration", "start"],
    },
    "timerange": {
        "port": "timeRange",
        "key": "ranges",
        "label": "Time Ranges",
        "preserve": ("source",),
        "fields": ["duration", "start"],
        "sort_fields": ["duration", "start"],
    },
    "artifacts": {
        "port": "artifacts",
        "key": "artifacts",
        "label": "Artifacts",
        "preserve": ("study",),
        "recount": True,
        "fields": ["duration", "start", "type", "category", "severity", "participant"],
        "sort_fields": ["duration", "start"],
    },
}


def _collection_field(kind: str, item: Any, field: str) -> float | str | None:
    """Extract one predicate/sort field from a collection item.

    Numeric fields return ``float``; text fields (``category``/``severity``/
    ``desc``/``text``/``type``/``participant``) return ``str``. Clip ``duration``
    sums the record's ``times`` spans, mirroring ``_exec_measure``'s total-duration
    path. ``timerange`` items are ``(start, end)`` tuples, not dicts.
    """
    if kind == "timerange":
        if not isinstance(item, (list, tuple)) or len(item) < 2:
            return None
        start = float(item[0] or 0.0)
        end = float(item[1] or 0.0)
        if field == "duration":
            return max(0.0, end - start)
        if field == "start":
            return start
        return None
    if not isinstance(item, dict):
        return None
    if kind == "artifacts":
        if field == "duration":
            return max(
                0.0,
                float(item.get("end", 0.0) or 0.0)
                - float(item.get("start", 0.0) or 0.0),
            )
        if field == "start":
            return float(item.get("start", 0.0) or 0.0)
        if field in ("type", "category", "severity", "participant"):
            return str(item.get(field, "") or "")
        return None
    if kind == "events":
        if field == "confidence":
            return float(item.get("confidence", 0.0) or 0.0)
        if field == "duration":
            return max(
                0.0,
                float(item.get("time_out", 0.0) or 0.0)
                - float(item.get("time_in", 0.0) or 0.0),
            )
        if field == "start":
            return float(item.get("time_in", 0.0) or 0.0)
    elif kind == "clips":
        if field == "duration":
            total = 0.0
            for start_str, end_str in item.get("times") or []:
                start = utils.timestamp_to_seconds(start_str)
                end = utils.timestamp_to_seconds(end_str)
                if start is not None and end is not None:
                    total += max(0.0, end - start)
            return total
        if field in ("category", "severity", "desc"):
            return str(item.get(field, "") or "")
    elif kind == "segments":
        if field == "text":
            return str(item.get("text", "") or "")
        if field == "duration":
            return max(
                0.0,
                float(item.get("end", 0.0) or 0.0)
                - float(item.get("start", 0.0) or 0.0),
            )
        if field == "start":
            return float(item.get("start", 0.0) or 0.0)
    return None


def _eval_predicate(field_val: float | str | None, op: str, raw_value: Any) -> bool:
    """Evaluate one ``{field, op, value}`` clause against an extracted field value.

    ``contains`` is a case-insensitive substring test; text fields support only
    ``==``/``!=`` (case-insensitive, trimmed). Numeric fields route through
    ``_GATE_OPS`` after a ``float()`` coerce (mirrors ``_exec_gate``'s guard).
    """
    if field_val is None:
        return False
    if op == "contains":
        return str(raw_value).strip().lower() in str(field_val).lower()
    if isinstance(field_val, str):
        a = field_val.strip().lower()
        b = str(raw_value).strip().lower()
        if op == "==":
            return a == b
        if op == "!=":
            return a != b
        return False  # ordering ops don't apply to text fields
    fn = _GATE_OPS.get(op)
    if fn is None:
        return False
    try:
        return fn(float(field_val), float(raw_value))
    except (TypeError, ValueError):
        return False


def _eval_clauses(kind: str, item: Any, params: dict[str, Any]) -> bool:
    """Evaluate the primary ``{field, op, value}`` clause plus the optional
    second clause, combined with AND/OR when ``combine`` isn't "off"."""
    meta = _COLLECTION_KINDS[kind]
    field = str(params.get("field") or meta["fields"][0])
    op = str(params.get("op") or ">=")
    first = _eval_predicate(
        _collection_field(kind, item, field), op, params.get("value")
    )
    combine = str(params.get("combine") or "off")
    if combine not in ("AND", "OR"):
        return first
    field2 = str(params.get("field2") or meta["fields"][0])
    op2 = str(params.get("op2") or ">=")
    second = _eval_predicate(
        _collection_field(kind, item, field2), op2, params.get("value2")
    )
    return (first and second) if combine == "AND" else (first or second)


def _wrap_collection(
    kind: str, src_envelope: dict[str, Any], items: list[Any]
) -> dict[str, Any]:
    """Re-wrap an item list in the kind's envelope, preserving source lineage."""
    meta = _COLLECTION_KINDS[kind]
    out: dict[str, Any] = {meta["key"]: items}
    for preserve_key in meta["preserve"]:
        if preserve_key in src_envelope:
            out[preserve_key] = src_envelope[preserve_key]
    if meta.get("recount"):
        out["count"] = len(items)  # artifacts carry a count; keep it honest
    return out


def _make_filter_executor(
    kind: str,
) -> Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]:
    """Keep items matching the clause(s) (see ``_eval_clauses``); same type in/out."""
    meta = _COLLECTION_KINDS[kind]

    def _exec(
        ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
    ) -> dict[str, Any]:
        env = inputs.get("in") or {}
        items = list(env.get(meta["key"]) or [])
        kept = [it for it in items if _eval_clauses(kind, it, params)]
        return {"out": _wrap_collection(kind, env, kept)}

    return _exec


def _make_partition_executor(
    kind: str,
) -> Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]:
    """Split one collection into ``matched`` / ``unmatched`` — the gate's missing
    data-level branch. Two same-typed outputs (runner stores the whole result
    dict; consumers read per-port)."""
    meta = _COLLECTION_KINDS[kind]

    def _exec(
        ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
    ) -> dict[str, Any]:
        env = inputs.get("in") or {}
        items = list(env.get(meta["key"]) or [])
        matched: list[Any] = []
        unmatched: list[Any] = []
        for it in items:
            target = matched if _eval_clauses(kind, it, params) else unmatched
            target.append(it)
        return {
            "matched": _wrap_collection(kind, env, matched),
            "unmatched": _wrap_collection(kind, env, unmatched),
        }

    return _exec


def _make_merge_executor(
    kind: str,
) -> Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]:
    """Union 2-3 same-type collections into one (fixes the one-wire-per-input
    wall). Preserves the first wired input's lineage; concatenates ``raw_results``
    for events so a downstream heatmap still sees per-frame coverage."""
    meta = _COLLECTION_KINDS[kind]

    def _exec(
        ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
    ) -> dict[str, Any]:
        merged: list[Any] = []
        raw: list[Any] = []
        base_env: dict[str, Any] = {}
        for port in ("in1", "in2", "in3"):
            env = inputs.get(port)
            if not isinstance(env, dict):
                continue
            if not base_env:
                base_env = env
            seq = env.get(meta["key"])
            if isinstance(seq, list):
                merged.extend(seq)
            extra = env.get("raw_results")
            if isinstance(extra, list):
                raw.extend(extra)
        out = _wrap_collection(kind, base_env, merged)
        if kind == "events":
            out["raw_results"] = raw
        return {"out": out}

    return _exec


def _make_limit_executor(
    kind: str,
) -> Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]:
    """Optionally sort by a numeric field, then keep the first N items."""
    meta = _COLLECTION_KINDS[kind]

    def _exec(
        ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
    ) -> dict[str, Any]:
        env = inputs.get("in") or {}
        items = list(env.get(meta["key"]) or [])
        sort_by = str(params.get("sort_by") or _SORT_NONE)
        order = str(params.get("order") or "desc")
        if sort_by != _SORT_NONE:

            def _key(it: dict[str, Any]) -> float:
                val = _collection_field(kind, it, sort_by)
                return float(val) if isinstance(val, (int, float)) else 0.0

            items.sort(key=_key, reverse=(order == "desc"))
        try:
            take = int(params.get("take", 0) or 0)
        except (TypeError, ValueError):
            take = 0
        if take > 0:
            items = items[:take]
        return {"out": _wrap_collection(kind, env, items)}

    return _exec


def _dedup_events(items: list[dict[str, Any]], gap: float) -> list[dict[str, Any]]:
    """Merge events whose spans overlap or sit within ``gap`` seconds; the merged
    event spans the union and keeps the max confidence."""
    ordered = sorted(items, key=lambda e: float(e.get("time_in", 0.0) or 0.0))
    out: list[dict[str, Any]] = []
    for ev in ordered:
        t_in = float(ev.get("time_in", 0.0) or 0.0)
        t_out = float(ev.get("time_out", t_in) or t_in)
        if out:
            prev = out[-1]
            p_in = float(prev.get("time_in", 0.0) or 0.0)
            p_out = float(prev.get("time_out", p_in) or p_in)
            if t_in <= p_out + gap:
                # Keep the higher-confidence member's fields, union the span.
                ev_conf = float(ev.get("confidence", 0.0) or 0.0)
                prev_conf = float(prev.get("confidence", 0.0) or 0.0)
                merged = dict(ev if ev_conf > prev_conf else prev)
                merged["time_in"] = min(p_in, t_in)
                merged["time_out"] = max(p_out, t_out)
                merged["confidence"] = max(prev_conf, ev_conf)
                out[-1] = merged
                continue
        out.append(dict(ev))
    return out


def _dedup_clips(records: list[dict[str, Any]], gap: float) -> list[dict[str, Any]]:
    """Drop clip records whose overall time-span overlaps (within ``gap``) a
    record already kept; keeps the first, extending the covered span."""

    def _span(rec: dict[str, Any]) -> tuple[float, float] | None:
        starts: list[float] = []
        ends: list[float] = []
        for start_str, end_str in rec.get("times") or []:
            start = utils.timestamp_to_seconds(start_str)
            end = utils.timestamp_to_seconds(end_str)
            if start is not None:
                starts.append(start)
            if end is not None:
                ends.append(end)
        if not starts or not ends:
            return None
        return (min(starts), max(ends))

    indexed = sorted(
        ((_span(rec), rec) for rec in records),
        key=lambda pair: pair[0][0] if pair[0] else 0.0,
    )
    out: list[dict[str, Any]] = []
    last_span: tuple[float, float] | None = None
    for span, rec in indexed:
        if span and last_span and span[0] <= last_span[1] + gap:
            last_span = (last_span[0], max(last_span[1], span[1]))
            continue
        out.append(rec)
        # Keep an untimed record but never let its None span clobber the tracker —
        # otherwise the next overlap check short-circuits and later duplicates leak.
        if span is not None:
            last_span = span
    return out


def _dedup_timeranges(items: list[Any], gap: float) -> list[tuple[float, float]]:
    """Merge ``(start, end)`` windows that overlap or sit within ``gap`` seconds."""
    spans: list[tuple[float, float]] = []
    for it in items:
        if isinstance(it, (list, tuple)) and len(it) >= 2:
            spans.append((float(it[0] or 0.0), float(it[1] or 0.0)))
    spans.sort()
    out: list[tuple[float, float]] = []
    for start, end in spans:
        if out and start <= out[-1][1] + gap:
            out[-1] = (out[-1][0], max(out[-1][1], end))
        else:
            out.append((start, end))
    return out


def _make_dedup_executor(
    kind: str,
) -> Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]:
    """Span-merge near-duplicate items (events, clips, and time ranges)."""
    meta = _COLLECTION_KINDS[kind]

    def _exec(
        ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
    ) -> dict[str, Any]:
        env = inputs.get("in") or {}
        items = list(env.get(meta["key"]) or [])
        try:
            gap = max(0.0, float(params.get("gap", 0) or 0))
        except (TypeError, ValueError):
            gap = 0.0
        if kind == "events":
            items = _dedup_events(items, gap)
        elif kind == "clips":
            items = _dedup_clips(items, gap)
        elif kind == "timerange":
            items = _dedup_timeranges(items, gap)
        return {"out": _wrap_collection(kind, env, items)}

    return _exec


def _predicate_params(kind: str) -> list[ParamSpec]:
    """The shared ``{field, op, value}`` clause for filter / partition nodes."""
    meta = _COLLECTION_KINDS[kind]
    return [
        {
            "name": "field",
            "type": "enum",
            "default": meta["fields"][0],
            "choices": list(meta["fields"]),
            "label": "Field",
        },
        {
            "name": "op",
            "type": "enum",
            "default": ">=",
            "choices": list(_COLLECTION_OPS),
            "label": "Comparison",
        },
        {
            "name": "value",
            "type": "string",
            "default": "",
            "label": "Value",
            "required": True,
        },
        # Optional second clause. "off" keeps the node single-clause; value2 is
        # deliberately not required so validation stays quiet in that case.
        {
            "name": "combine",
            "type": "enum",
            "default": "off",
            "choices": ["off", "AND", "OR"],
            "label": "Second clause",
        },
        {
            "name": "field2",
            "type": "enum",
            "default": meta["fields"][0],
            "choices": list(meta["fields"]),
            "label": "Field 2",
        },
        {
            "name": "op2",
            "type": "enum",
            "default": ">=",
            "choices": list(_COLLECTION_OPS),
            "label": "Comparison 2",
        },
        {
            "name": "value2",
            "type": "string",
            "default": "",
            "label": "Value 2",
        },
    ]


def _limit_params(kind: str) -> list[ParamSpec]:
    """Sort-key + order + count for the limit (top-N) node."""
    meta = _COLLECTION_KINDS[kind]
    return [
        {
            "name": "sort_by",
            "type": "enum",
            "default": _SORT_NONE,
            "choices": [_SORT_NONE] + list(meta["sort_fields"]),
            "label": "Sort by",
        },
        {
            "name": "order",
            "type": "enum",
            "default": "desc",
            "choices": ["desc", "asc"],
            "label": "Order",
        },
        {
            "name": "take",
            "type": "number",
            "default": 10,
            "min": 0,
            "label": "Keep first N",
            "required": True,
        },
    ]


# ---------------------------------------------------------------------------
# Typed-port adapters (M3) — pure value -> value, applied by the runner (M4)
# ---------------------------------------------------------------------------
#
# An adapter coerces an output value whose port type differs from the consuming
# input's type. They are pure (no ctx/params), so each value self-carries its
# source descriptor (see executors above).


def _adapt_transcript_to_segments(value: dict[str, Any]) -> dict[str, Any]:
    return {
        "segments": value.get("segments", []),
        "source": value.get("source", {}),
    }


def _adapt_segments_to_timerange(value: dict[str, Any]) -> dict[str, Any]:
    """Project every segment to a ``(start, end)`` window (no text filter).

    Distinct from the ``find_word`` node, which filters segments by a search term;
    this is the default coercion when ``segments`` is wired straight into a
    ``timeRange`` input.
    """
    ranges: list[tuple[float, float]] = []
    for seg in value.get("segments") or []:
        start = float(seg.get("start", 0.0) or 0.0)
        end = float(seg.get("end", start) or start)
        ranges.append((start, max(start, end)))
    return {"ranges": ranges, "source": value.get("source", {})}


def _adapt_timerange_to_cliprecords(value: dict[str, Any]) -> dict[str, Any]:
    import files

    source = value.get("source") or {}
    ranges = [(float(s), float(e)) for s, e in (value.get("ranges") or [])]
    records = files.build_clip_records(
        participant=str(source.get("participant", "") or ""),
        source_filename=_clip_source_filename(source),
        time_ranges=ranges,
        description="workflow",
        study=str(source.get("study", "") or ""),
    )
    return {"records": records, "study": str(source.get("study", "") or "")}


def _adapt_events_to_timerange(value: dict[str, Any]) -> dict[str, Any]:
    events = value.get("events") or []
    ranges: list[tuple[float, float]] = []
    for ev in events:
        t_in = float(ev.get("time_in", 0.0) or 0.0)
        t_out = float(ev.get("time_out", t_in) or t_in)
        ranges.append((t_in, max(t_in, t_out)))
    source = value.get("source") or _source_from_events(events)
    return {"ranges": ranges, "source": source}


def _adapt_cliprecords_to_timerange(value: dict[str, Any]) -> dict[str, Any]:
    """Project clip records to ``(start, end)`` windows — e.g. *sheet cells →
    SS scan windows*. Each record's resolved ``times`` (via ``files.prepare_clip``
    when not pre-filled) are converted to seconds; the inverse of
    :func:`_adapt_timerange_to_cliprecords`.
    """
    import files

    ranges: list[tuple[float, float]] = []
    source: dict[str, Any] = {}
    for rec in value.get("records") or []:
        prepared = (
            rec
            if rec.get("times")
            else files.prepare_clip(cast(utils.ClipRecord, dict(rec)))
        )
        for start_str, end_str in prepared.get("times") or []:
            start = utils.timestamp_to_seconds(start_str)
            end = utils.timestamp_to_seconds(end_str)
            if start is not None and end is not None:
                ranges.append((start, max(start, end)))
        if not source and rec.get("participant"):
            source = {
                "participant": str(rec.get("participant", "") or ""),
                "study": str(rec.get("study", value.get("study", "")) or ""),
                "source_filename": "",
                "video_paths": [],
            }
    return {"ranges": ranges, "source": source}


def _adapt_events_to_cliprecords(value: dict[str, Any]) -> dict[str, Any]:
    import files

    events = value.get("events") or []
    source = value.get("source") or _source_from_events(events)
    spans: list[tuple[float, float]] = []
    for ev in events:
        t_in = float(ev.get("time_in", 0.0) or 0.0)
        t_out = float(ev.get("time_out", t_in) or t_in)
        spans.append((t_in, max(t_in, t_out)))
    records = files.build_clip_records(
        participant=str(source.get("participant", "") or ""),
        source_filename=_clip_source_filename(source),
        time_ranges=spans,
        description="event",
        category="workflow",
        study=str(source.get("study", "") or ""),
        cluster_gap=_DEFAULT_EVENT_CLUSTER_GAP,
    )
    return {"records": records, "study": str(source.get("study", "") or "")}


def _adapt_video_to_scalar(value: dict[str, Any]) -> float:
    import video

    paths = value.get("video_paths") or []
    if not paths:
        return 0.0
    props = video.probe_video_properties(paths[0]) or {}
    return float(props.get("duration", 0.0) or 0.0)


ADAPTERS: dict[tuple[str, str], Callable[[Any], Any]] = {
    ("transcript", "segments"): _adapt_transcript_to_segments,
    ("segments", "timeRange"): _adapt_segments_to_timerange,
    ("timeRange", "clipRecords"): _adapt_timerange_to_cliprecords,
    ("clipRecords", "timeRange"): _adapt_cliprecords_to_timerange,
    ("events", "timeRange"): _adapt_events_to_timerange,
    ("events", "clipRecords"): _adapt_events_to_cliprecords,
    ("video", "scalar"): _adapt_video_to_scalar,
}

# Plain-language description of what each adapter does, served alongside the
# table (see ``serialize_adapters``) so a coerced (dashed) wire's tooltip can
# explain the transformation — not just that one happened. One per ADAPTERS key
# (guarded by ``tests/test_workflows_api``); a missing one degrades to no suffix.
_ADAPTER_DESCRIPTIONS: dict[tuple[str, str], str] = {
    ("transcript", "segments"): "use the transcript's segments",
    ("segments", "timeRange"): "use each segment's time span",
    ("timeRange", "clipRecords"): "make a clip from each time range",
    ("clipRecords", "timeRange"): "use each clip's time span",
    ("events", "timeRange"): "use each event's time span",
    ("events", "clipRecords"): "make a clip from each event (clustered)",
    ("video", "scalar"): "use the video's duration",
}


# Wire each declarative NodeType to its executor. ``serialize_catalog`` strips
# ``execute`` again for the JSON catalog endpoint, so this stays server-internal.
_EXECUTORS: dict[
    str, Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]
] = {
    "video_source": _exec_video_source,
    "sheet_selection": _exec_sheet_selection,
    "region": _exec_region,
    "time_range": _exec_time_range,
    "transcribe": _exec_transcribe,
    "find_word": _exec_find_word,
    "transcript_export": _exec_transcript_export,
    "data_export": _exec_data_export,
    "summarize": _exec_summarize,
    "citations": _exec_citations,
    "friction": _exec_friction,
    "multitool": _exec_multitool,
    "highlights": _exec_highlights,
    "make_clips": _exec_make_clips,
    "interval_captures": _exec_interval_captures,
    "build_reel": _exec_build_reel,
    "timelapse": _exec_timelapse,
    "heatmap": _exec_heatmap,
    "measure": _exec_measure,
    "timeline_viewer": _exec_timeline_viewer,
    "gate": _exec_gate,
    "gate_collection": _exec_gate_collection,
}

# The ten per-detector Screenspace nodes share one body via the factory above;
# the unified ``detect`` node dispatches into the same body by ``detector`` param.
for _ss_tool in _SS_DETECTOR_SPECS:
    _EXECUTORS[f"ss_{_ss_tool}"] = _make_ss_executor(_ss_tool)
_EXECUTORS["detect"] = _exec_detect

# Collection-algebra control nodes — per-type families, all factory-generated.
# Registered here (NODE_TYPES + _EXECUTORS together) so the attach loop below
# wires their ``execute`` like any other node. Category "Collection" groups them
# apart from measure/gate in the palette.
for _kind, _meta in _COLLECTION_KINDS.items():
    _T = _meta["port"]
    _name = _meta["label"]
    _lname = _name.lower()
    NODE_TYPES[f"filter_{_kind}"] = {
        "id": f"filter_{_kind}",
        "label": f"Filter {_name}",
        "domain": "control",
        "category": "Collection",
        "description": f"Keep only the {_lname} matching a field/comparison/value test.",
        "inputs": [{"name": "in", "type": _T}],
        "outputs": [{"name": "out", "type": _T}],
        "params": _predicate_params(_kind),
        "requires": [],
    }
    _EXECUTORS[f"filter_{_kind}"] = _make_filter_executor(_kind)
    NODE_TYPES[f"partition_{_kind}"] = {
        "id": f"partition_{_kind}",
        "label": f"Partition {_name}",
        "domain": "control",
        "category": "Collection",
        "description": f"Split {_lname} into matched and unmatched branches by a test.",
        "inputs": [{"name": "in", "type": _T}],
        "outputs": [
            {"name": "matched", "type": _T},
            {"name": "unmatched", "type": _T},
        ],
        "params": _predicate_params(_kind),
        "requires": [],
    }
    _EXECUTORS[f"partition_{_kind}"] = _make_partition_executor(_kind)
    NODE_TYPES[f"merge_{_kind}"] = {
        "id": f"merge_{_kind}",
        "label": f"Merge {_name}",
        "domain": "control",
        "category": "Collection",
        "description": f"Combine two or three {_lname} streams into one.",
        "inputs": [
            {"name": "in1", "type": _T},
            {"name": "in2", "type": _T, "optional": True},
            {"name": "in3", "type": _T, "optional": True},
        ],
        "outputs": [{"name": "out", "type": _T}],
        "params": [],
        "requires": [],
    }
    _EXECUTORS[f"merge_{_kind}"] = _make_merge_executor(_kind)
    NODE_TYPES[f"limit_{_kind}"] = {
        "id": f"limit_{_kind}",
        "label": f"Limit {_name}",
        "domain": "control",
        "category": "Collection",
        "description": f"Optionally sort {_lname} by a field, then keep the first N.",
        "inputs": [{"name": "in", "type": _T}],
        "outputs": [{"name": "out", "type": _T}],
        "params": _limit_params(_kind),
        "requires": [],
    }
    _EXECUTORS[f"limit_{_kind}"] = _make_limit_executor(_kind)

# dedup is span-based -> events + clips + time ranges.
for _kind in ("events", "clips", "timerange"):
    _T = _COLLECTION_KINDS[_kind]["port"]
    _name = _COLLECTION_KINDS[_kind]["label"]
    _lname = _name.lower()
    NODE_TYPES[f"dedup_{_kind}"] = {
        "id": f"dedup_{_kind}",
        "label": f"Dedup {_name}",
        "domain": "control",
        "category": "Collection",
        "description": f"Merge overlapping or near-duplicate {_lname} into single spans.",
        "inputs": [{"name": "in", "type": _T}],
        "outputs": [{"name": "out", "type": _T}],
        "params": [
            {
                "name": "gap",
                "type": "number",
                "default": 0,
                "min": 0,
                "label": "Merge gap (s)",
            },
        ],
        "requires": [],
    }
    _EXECUTORS[f"dedup_{_kind}"] = _make_dedup_executor(_kind)

for _node_id, _executor in _EXECUTORS.items():
    NODE_TYPES[_node_id]["execute"] = _executor


# ---------------------------------------------------------------------------
# Run engine (M4) — DAG topo-sort + sequential ready-set execution
# ---------------------------------------------------------------------------
#
# ``WorkflowRunner`` runs one blueprint on a daemon thread (the server spawns it).
# It calls the executors directly with the uniform ``NodeContext`` contract, so a
# cross-domain DAG gets clean end-to-end progress + cancellation without routing
# through the per-domain worker queues. Execution is strictly sequential (the v1
# decision: Whisper/Ollama are single-resource); intra-node parallelism (e.g.
# ``process_clips``' own pool) still applies.

# Run + per-node status constants (mirrors screenspace_manifest's TASK_STATUS_*).
RUN_STATUS_QUEUED = "queued"
RUN_STATUS_RUNNING = "running"
RUN_STATUS_COMPLETED = "completed"
RUN_STATUS_FAILED = "failed"
RUN_STATUS_CANCELLED = "cancelled"

NODE_STATUS_QUEUED = "queued"
NODE_STATUS_RUNNING = "running"
NODE_STATUS_COMPLETED = "completed"
NODE_STATUS_FAILED = "failed"
NODE_STATUS_SKIPPED = "skipped"

# Canvas-only sticky-note pseudo-node (frontend-created, not in NODE_TYPES).
# Notes live in blueprint["nodes"] so they ride save/undo/copy/import for free;
# the runner filters them out so they never execute or appear in run snapshots.
NOTE_NODE_TYPE = "note"

_PROGRESS_NOTIFY_INTERVAL = (
    0.5  # seconds; throttle SSE notifies (copy screenspace_worker)
)


class WorkflowCycleError(ValueError):
    """Raised by :func:`topo_order` when the graph is not a DAG (rejected at submit)."""


def _now_iso() -> str:
    """UTC ISO-8601 timestamp for run/node start+complete stamps."""
    return datetime.now(timezone.utc).isoformat()


def topo_order(nodes: list[dict[str, Any]], edges: list[dict[str, Any]]) -> list[str]:
    """Return node ids in execution order (Kahn); raise on a cycle.

    Stable: ties break by the given node order, so a run is reproducible. Edges
    referencing unknown node ids are ignored (a stale wire never blocks a run).
    """
    ids = [n["id"] for n in nodes]
    id_set = set(ids)
    adj: dict[str, list[str]] = {nid: [] for nid in ids}
    indeg: dict[str, int] = {nid: 0 for nid in ids}
    for edge in edges:
        src, dst = edge.get("from"), edge.get("to")
        if src in id_set and dst in id_set:
            adj[src].append(dst)
            indeg[dst] += 1
    ready = [nid for nid in ids if indeg[nid] == 0]
    order: list[str] = []
    while ready:
        nid = ready.pop(0)
        order.append(nid)
        for nxt in adj[nid]:
            indeg[nxt] -= 1
            if indeg[nxt] == 0:
                ready.append(nxt)
    if len(order) != len(ids):
        raise WorkflowCycleError("Workflow graph contains a cycle")
    return order


def blueprint_participant_nodes(blueprint: dict[str, Any]) -> list[dict[str, Any]]:
    """The blueprint's participant-bound source nodes (``video_source``).

    Whole-study batch (P3) fans out over these — a blueprint with none can't be
    rebound per participant, so the batch endpoint rejects it.
    """
    return [n for n in blueprint.get("nodes", []) if n.get("type") == "video_source"]


def bind_participant(blueprint: dict[str, Any], participant: str) -> dict[str, Any]:
    """Deep-copy ``blueprint`` and rebind every ``video_source`` node to ``participant``.

    Pure: the original is never mutated. Non-source nodes are untouched — only the
    participant-bound sources are rewritten, so one blueprint can run once per
    participant in a batch (P3).
    """
    clone = copy.deepcopy(blueprint)
    for node in clone.get("nodes", []):
        if node.get("type") == "video_source":
            params = node.setdefault("params", {})
            params["participant"] = participant
    return clone


def _port_type(type_id: str, port_name: str | None, direction: str) -> str | None:
    """Look up a port's wire type from ``NODE_TYPES`` ('in' or 'out' direction)."""
    node_type = NODE_TYPES.get(type_id)
    if not node_type or not port_name:
        return None
    ports = node_type["outputs"] if direction == "out" else node_type["inputs"]
    for port in ports:
        if port["name"] == port_name:
            return port["type"]
    return None


def _port_optional(type_id: str | None, port_name: str | None) -> bool:
    """True if a node type's *input* port is declared ``optional`` (else False)."""
    node_type = NODE_TYPES.get(type_id) if type_id else None
    if not node_type or not port_name:
        return False
    for port in node_type["inputs"]:
        if port["name"] == port_name:
            return bool(port.get("optional"))
    return False


def _summarize_value(value: Any) -> Any:
    """Shrink a node output to a JSON-safe summary (counts + terminal pointers).

    The full result (raw frames, whole segment lists, ClipRecords) stays in the
    runner's in-memory store; the snapshot ships only what a panel needs.
    """
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, list):
        return {"count": len(value)}
    if isinstance(value, dict):
        out: dict[str, Any] = {}
        for list_key in (
            "events",
            "segments",
            "artifacts",
            "records",
            "ranges",
            "times",
        ):
            seq = value.get(list_key)
            if isinstance(seq, list):
                out["count"] = len(seq)
                break
        for scalar_key in ("path", "study", "name", "participant", "count"):
            sval = value.get(scalar_key)
            if isinstance(sval, (str, int, float, bool)):
                out[scalar_key] = sval
        return out
    return type(value).__name__


def _node_result_summary(result: Any) -> dict[str, Any]:
    """Per-port summary of a completed node's output (see :func:`_summarize_value`)."""
    if not isinstance(result, dict):
        return {}
    return {port: _summarize_value(val) for port, val in result.items()}


# ---- Per-node result sidecars (P5) ----------------------------------------
#
# The snapshot ships only counts/pointers; the *full* inspectable result is
# written to ``<output_dir>/workflow_runs/<run_id>/<node_id>.json`` so the
# run-history UI can lazily fetch and render it after the runner is evicted.

# Output port types the run-history UI renders on row-expand. Plumbing types
# (video/region/timeRange handles, control) are persisted for resume (below)
# but hidden from the inspector.
_INSPECTABLE_PORT_TYPES = frozenset(
    {
        "artifacts",
        "events",
        "segments",
        "summary",
        "citations",
        "friction",
        "manifest",
        "viewerHtml",
        "scalar",
    }
)

# Output port types persisted in the sidecar — everything JSON-safe, so a
# later resume (``compute_resume_plan``) can reload a completed node's outputs
# verbatim. Only ``clipRecords`` is excluded: its records carry gspread
# ``Cell`` objects that don't survive JSON, so clipRecords producers always
# re-run on resume (cheap — one Sheets read).
_SIDECAR_PORT_TYPES = _INSPECTABLE_PORT_TYPES | frozenset(
    {
        "transcript",
        "timeRange",
        "timestamps",
        "video",
        "participant",
        "region",
        "control",  # a gate's pass verdict is a JSON bool; resume needs it
    }
)


def run_results_dir(output_dir: Path | str, run_id: str) -> Path:
    """Directory holding one run's per-node result sidecars (P5)."""
    return Path(output_dir) / "workflow_runs" / run_id


def _project_inspectable(port_type: str, value: Any) -> Any:
    """Strip a port value down to its inspectable (JSON-safe) projection.

    The only heavy/non-serializable rider is an ``events`` value's
    ``raw_results`` — per-frame scoring kept for the heatmap node, which can
    carry numpy arrays and is large. Drop it; the rest of every inspectable type
    is already JSON-safe (paths, counts, text, timestamps).
    """
    if port_type == "events" and isinstance(value, dict):
        return {k: v for k, v in value.items() if k != "raw_results"}
    return value


def _filter_result_ports(
    node_type_id: str, result: Any, allowed: frozenset[str]
) -> dict[str, Any]:
    """Keep only the result ports whose declared type is in ``allowed`` (projected)."""
    if not isinstance(result, dict):
        return {}
    node_type = NODE_TYPES.get(node_type_id)
    out_types = {p["name"]: p["type"] for p in (node_type or {}).get("outputs", [])}
    kept: dict[str, Any] = {}
    for port, val in result.items():
        ptype = out_types.get(port)
        if ptype in allowed:
            kept[port] = _project_inspectable(ptype, val)
    return kept


def _inspectable_result(node_type_id: str, result: Any) -> dict[str, Any]:
    """The UI-inspectable projection of a node result."""
    return _filter_result_ports(node_type_id, result, _INSPECTABLE_PORT_TYPES)


def inspectable_sidecar_view(payload: dict[str, Any]) -> dict[str, Any]:
    """Project a stored sidecar payload down to its UI-inspectable ports.

    Sidecars persist every JSON-safe port (for resume) plus a self-describing
    ``__type__`` key; the run-history inspector should render only the
    inspectable subset, exactly as it did before the widening.
    """
    node_type_id = str(payload.get("__type__", "") or "")
    out_types = {
        p["name"]: p["type"]
        for p in (NODE_TYPES.get(node_type_id) or {}).get("outputs", [])
    }
    return {
        port: val
        for port, val in payload.items()
        if port != "__type__" and out_types.get(port) in _INSPECTABLE_PORT_TYPES
    }


def write_node_sidecar(
    output_dir: Path | str, run_id: str, node_id: str, node_type_id: str, result: Any
) -> bool:
    """Atomically write a node's JSON-safe result ports to its run sidecar.

    Persists every ``_SIDECAR_PORT_TYPES`` port plus a self-describing
    ``__type__`` key (consumed by resume + the read-time inspectable filter).
    Returns True iff a sidecar file now exists for this node. JSON-sanitizes via
    :func:`utils.sanitize_floats` (non-finite floats / numpy scalars). A bad
    ``node_id`` (path separators) or an empty payload writes nothing and
    returns False.
    """
    if not node_id or node_id != os.path.basename(node_id) or node_id in (".", ".."):
        return False
    payload = _filter_result_ports(node_type_id, result, _SIDECAR_PORT_TYPES)
    if not payload:
        return False
    payload["__type__"] = node_type_id
    path = run_results_dir(output_dir, run_id) / f"{node_id}.json"
    tmp = path.with_suffix(".json.tmp")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp.write_text(
            json.dumps(utils.sanitize_floats(payload), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        os.replace(tmp, path)
        return True
    except (OSError, TypeError, ValueError) as exc:
        utils.warning_print(f"workflow sidecar write failed ({node_id}): {exc}")
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        return False


# Collection nodes that pass an events value's ``raw_results`` through
# unchanged (see ``_COLLECTION_KINDS["events"]["preserve"]`` / merge's concat).
# The resume planner walks heatmap ancestry through these.
_RAW_RESULTS_PRESERVING = frozenset(
    {
        "filter_events",
        "partition_events",
        "merge_events",
        "limit_events",
        "dedup_events",
    }
)


def compute_resume_plan(
    blueprint: dict[str, Any],
    prior_node_states: dict[str, Any],
    load_sidecar: Callable[[str], dict[str, Any] | None],
) -> tuple[dict[str, dict[str, Any]], list[str]]:
    """Plan a resume: which prior-run nodes can be reused as seeds vs. re-run.

    Pure. ``load_sidecar(node_id)`` returns the stored sidecar payload (with
    its ``__type__`` key) or None. Returns ``(seed_results, notes)`` where
    ``seed_results`` feeds :class:`WorkflowRunner`'s ``seed_results`` and
    ``notes`` carries human-readable degradation reasons.

    A node re-runs when: it didn't complete in the prior run; its id/type
    changed since (graph edited between runs); its sidecar is missing or
    doesn't cover every declared output port (e.g. clipRecords producers,
    which are never sidecar-persisted); OR any ancestor re-runs (fresh inputs
    invalidate the cached output). Additionally, a re-running ``heatmap``
    forces its completed events-ancestry to re-run too — sidecars project out
    the heavy ``raw_results`` the heatmap consumes, so a seeded events
    producer would make the resumed heatmap silently emit nothing.
    """
    nodes = [n for n in blueprint.get("nodes", []) if n.get("type") != NOTE_NODE_TYPE]
    by_id = {n["id"]: n for n in nodes}
    children: dict[str, list[str]] = {n["id"]: [] for n in nodes}
    parents: dict[str, list[str]] = {n["id"]: [] for n in nodes}
    for edge in blueprint.get("edges", []):
        src, dst = edge.get("from"), edge.get("to")
        if src in by_id and dst in by_id:
            children[src].append(dst)
            parents[dst].append(src)

    notes: list[str] = []
    rerun: set[str] = set()
    seeds: dict[str, dict[str, Any]] = {}
    for n in nodes:
        nid = n["id"]
        prior = prior_node_states.get(nid) or {}
        if prior.get("status") != NODE_STATUS_COMPLETED:
            rerun.add(nid)
            continue
        payload = load_sidecar(nid)
        if not isinstance(payload, dict):
            rerun.add(nid)
            continue
        if str(payload.get("__type__", "")) != str(n.get("type", "")):
            rerun.add(nid)  # the node changed type since the prior run
            continue
        declared = (NODE_TYPES.get(str(n.get("type", ""))) or {}).get("outputs", [])
        stored = {k: v for k, v in payload.items() if k != "__type__"}
        if not declared or any(p["name"] not in stored for p in declared):
            # Some output port wasn't persisted (non-JSON-safe type, or the
            # executor omitted it) — downstream would see None; re-run instead.
            rerun.add(nid)
            continue
        seeds[nid] = stored

    def _close_under_descendants() -> None:
        stack = list(rerun)
        while stack:
            for child in children.get(stack.pop(), []):
                if child not in rerun:
                    rerun.add(child)
                    stack.append(child)

    # Fixpoint: descendant closure and the heatmap raw_results rule feed each
    # other (a forced ancestor invalidates its own seeded descendants).
    while True:
        before = len(rerun)
        _close_under_descendants()
        for n in nodes:
            if n.get("type") != "heatmap" or n["id"] not in rerun:
                continue
            stack = list(parents.get(n["id"], []))
            while stack:
                pid = stack.pop()
                if pid in rerun:
                    continue
                rerun.add(pid)
                if by_id[pid].get("type") in _RAW_RESULTS_PRESERVING:
                    stack.extend(parents.get(pid, []))
        if len(rerun) == before:
            break

    seed_results = {nid: val for nid, val in seeds.items() if nid not in rerun}
    if not seed_results:
        notes.append("No prior results could be reused — running everything")
    return seed_results, notes


class WorkflowRunner:
    """Executes one blueprint DAG; ``run()`` is the daemon-thread target.

    Holds per-node state + an in-memory per-(node,port) result store. ``on_update``
    is invoked on every status transition (immediately) and on throttled progress
    ticks; the server wires it to the SSE notify. ``snapshot()`` returns the
    JSON-safe run record (counts + status, large blobs stripped) for the API.
    """

    def __init__(
        self,
        run_id: str,
        blueprint: dict[str, Any],
        ctx: NodeContext,
        on_update: Callable[[], None] | None = None,
        participant: str = "",
        batch_id: str = "",
        triggered: bool = False,
        target_node_id: str = "",
        seed_results: dict[str, dict[str, Any]] | None = None,
        seed_note: str = "",
    ) -> None:
        self.run_id = run_id
        self.blueprint_id = str(blueprint.get("id", "") or "")
        # Partial run (P11): when set, only this node and its transitive ancestors
        # execute; the rest are marked skipped. Empty → run the whole graph.
        self.target_node_id = target_node_id
        # Pre-seeded results: {node_id: result} for nodes whose output is
        # already known — participant-independent sources a batch coordinator
        # computed once (P3), or completed nodes reloaded from a prior run's
        # sidecars on resume. A seeded node is stored as if it ran, skipping
        # its executor. ``seed_note`` (resume) is surfaced on each seeded node.
        self._seed_results = seed_results or {}
        self._seed_note = seed_note
        # Batch identity (P3): empty for a normal single run; a child run carries
        # its participant + parent batch id so the snapshot can be grouped.
        self.participant = participant
        self.batch_id = batch_id
        # Watch-dir trigger (P6): True when this run was auto-launched by the
        # directory watcher (surfaced as a badge in the run history).
        self.triggered = triggered
        # Sticky notes are canvas annotations, not executable nodes — drop them
        # before node_states is built so they never run, fail as "No executor",
        # or pad the snapshot's node counts.
        self.nodes = [
            n for n in blueprint.get("nodes", []) if n.get("type") != NOTE_NODE_TYPE
        ]
        self.edges = list(blueprint.get("edges", []))
        self.ctx = ctx
        self.on_update = on_update or (lambda: None)
        self._nodes_by_id = {n["id"]: n for n in self.nodes}
        self.node_states: dict[str, dict[str, Any]] = {
            n["id"]: {
                "status": NODE_STATUS_QUEUED,
                "progress": 0.0,
                "error": None,
                # Non-fatal note for a degraded-but-completed node (Ollama down,
                # nothing wired, an adapter that couldn't coerce) — distinct from
                # ``error`` (which means FAILED). Surfaced in the run history.
                "note": None,
                "started_at": None,
                "completed_at": None,
            }
            for n in self.nodes
        }
        self._results: dict[str, dict[str, Any]] = {}
        # Node ids with an inspectable result sidecar on disk (P5); surfaced as
        # ``hasResult`` in the snapshot so the UI knows it can fetch on demand.
        self._sidecars: set[str] = set()
        self.status = RUN_STATUS_QUEUED
        self.started_at: str | None = None
        self.completed_at: str | None = None
        self._lock = threading.Lock()
        self._last_notify = 0.0

    # ---- control ----

    def cancel(self) -> None:
        """Signal the run-wide cancel event (checked between nodes + by executors)."""
        self.ctx.cancel_event.set()

    # ---- graph helpers ----

    def _incoming(self, node_id: str) -> list[dict[str, Any]]:
        return [e for e in self.edges if e.get("to") == node_id]

    def _deps(self, node_id: str) -> set[str]:
        deps: set[str] = set()
        for edge in self._incoming(node_id):
            src = edge.get("from")
            if isinstance(src, str) and src in self._nodes_by_id:
                deps.add(src)
        return deps

    def _ancestors_inclusive(self, node_id: str) -> set[str]:
        """``node_id`` plus everything it transitively depends on (for partial runs)."""
        seen: set[str] = set()
        stack = [node_id]
        while stack:
            nid = stack.pop()
            if nid in seen:
                continue
            seen.add(nid)
            stack.extend(self._deps(nid))
        return seen

    def _gate_blocks(self, node_id: str) -> bool:
        """True if ``node_id`` is a gate that completed with ``pass`` False."""
        node = self._nodes_by_id.get(node_id)
        if not node or node.get("type") not in ("gate", "gate_collection"):
            return False
        return (self._results.get(node_id) or {}).get("pass") is False

    def _should_skip(self, node_id: str) -> bool:
        """Decide whether ``node_id`` is skipped before it runs.

        Skipped when a *required* input's producer failed/skipped (the node can't
        run without it) or when a gate gating this node blocks it. A dead producer
        feeding an *optional* input is tolerated — the node runs with that input
        absent (executors already read ``inputs.get(...)``), so a muted or broken
        branch of a ``merge`` (or any optional input) doesn't sink the whole
        downstream.
        """
        node = self._nodes_by_id.get(node_id)
        node_type = node.get("type") if node else None
        for edge in self._incoming(node_id):
            dep = edge.get("from")
            if not isinstance(dep, str) or dep not in self._nodes_by_id:
                continue
            status = self.node_states[dep]["status"]
            out_type = _port_type(
                self._nodes_by_id[dep].get("type"), edge.get("fromPort"), "out"
            )
            if out_type == "control":
                # A gate edge: skip if the gate can't pass us through — it blocked
                # (``pass`` False) or it never completed (failed/skipped).
                if status in (NODE_STATUS_FAILED, NODE_STATUS_SKIPPED):
                    return True
                if status == NODE_STATUS_COMPLETED and self._gate_blocks(dep):
                    return True
                continue
            # A data edge: only a *required* input's dead producer forces a skip.
            if status in (
                NODE_STATUS_FAILED,
                NODE_STATUS_SKIPPED,
            ) and not _port_optional(node_type, edge.get("toPort")):
                return True
        return False

    def _gather_inputs(self, node: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
        """Map upstream results onto this node's input ports, applying adapters.

        Returns ``(inputs, notes)``. ``notes`` carries any adapter-failure
        messages so the runner can surface them on the node — a coercion that
        raises otherwise degrades the input to ``None`` invisibly (only a server
        log). Control edges (a gate's ``control`` output) establish a dependency
        but carry no data, so they are excluded here — they never clobber a real
        input.
        """
        inputs: dict[str, Any] = {}
        notes: list[str] = []
        for edge in self._incoming(node["id"]):
            src_id = edge.get("from")
            to_port = edge.get("toPort")
            from_port = edge.get("fromPort")
            if not (
                isinstance(src_id, str)
                and isinstance(to_port, str)
                and isinstance(from_port, str)
            ):
                continue
            src_node = self._nodes_by_id.get(src_id)
            if src_node is None:
                continue
            out_type = _port_type(src_node["type"], from_port, "out")
            if out_type == "control":
                continue
            in_type = _port_type(node["type"], to_port, "in")
            value = (self._results.get(src_id) or {}).get(from_port)
            if out_type is not None and in_type is not None and out_type != in_type:
                adapter = ADAPTERS.get((out_type, in_type))
                if adapter is not None:
                    try:
                        value = adapter(value)
                    except Exception as exc:  # adapter failure → empty input + note
                        utils.warning_print(
                            f"workflow adapter {out_type}->{in_type} failed: {exc}"
                        )
                        notes.append(f"Couldn't convert {out_type} to {in_type}")
                        value = None
            inputs[to_port] = value
        return inputs, notes

    # ---- state + notify ----

    def _set_node(self, node_id: str, **changes: Any) -> None:
        with self._lock:
            self.node_states[node_id].update(changes)

    def _notify(self, force: bool = False) -> None:
        now = time.monotonic()
        if force or now - self._last_notify >= _PROGRESS_NOTIFY_INTERVAL:
            self._last_notify = now
            try:
                self.on_update()
            except Exception:
                pass

    def _make_progress(self, node_id: str) -> Callable[[float], None]:
        def _on_progress(fraction: float) -> None:
            with self._lock:
                self.node_states[node_id]["progress"] = max(
                    0.0, min(1.0, float(fraction))
                )
            self._notify()

        return _on_progress

    # ---- run ----

    def run(self) -> None:
        """Execute the DAG in topological order. Safe to call once, on a thread."""
        self.status = RUN_STATUS_RUNNING
        self.started_at = _now_iso()
        self._notify(force=True)
        try:
            order = topo_order(self.nodes, self.edges)
        except WorkflowCycleError:
            self.status = RUN_STATUS_FAILED
            self.completed_at = _now_iso()
            self._notify(force=True)
            return

        # Partial run: keep only the target node and its ancestors; the rest are
        # skipped up front (they never execute and don't block completion).
        if self.target_node_id and self.target_node_id in self._nodes_by_id:
            keep = self._ancestors_inclusive(self.target_node_id)
            for nid in order:
                if nid not in keep:
                    self._set_node(
                        nid, status=NODE_STATUS_SKIPPED, completed_at=_now_iso()
                    )
            order = [nid for nid in order if nid in keep]

        for node_id in order:
            node = self._nodes_by_id[node_id]
            if self.ctx.cancel_event.is_set():
                self._set_node(
                    node_id, status=NODE_STATUS_SKIPPED, completed_at=_now_iso()
                )
                continue
            # Seeded result (batch precompute or resume): the node's output is
            # authoritatively known — store it as if it just ran, skipping its
            # executor. Checked BEFORE the mute/skip gates: a resume seed for a
            # completed node must survive even when a (re-run) parent upstream
            # is currently marked skipped/muted — the prior run already proved
            # this node's output. (Batch seeds are parentless sources, so the
            # ordering change is behavior-neutral for them.)
            if node_id in self._seed_results:
                seeded = self._seed_results[node_id]
                with self._lock:
                    self._results[node_id] = seeded
                if write_node_sidecar(
                    self.ctx.output_dir, self.run_id, node_id, node["type"], seeded
                ) and _inspectable_result(node["type"], seeded):
                    with self._lock:
                        self._sidecars.add(node_id)
                self._set_node(
                    node_id,
                    status=NODE_STATUS_COMPLETED,
                    progress=1.0,
                    completed_at=_now_iso(),
                    note=self._seed_note or None,
                )
                self._notify(force=True)
                continue

            # A muted node is skipped intrinsically; _should_skip then propagates
            # SKIPPED to its whole downstream subtree (same as a blocking gate).
            if node.get("disabled"):
                self._set_node(
                    node_id, status=NODE_STATUS_SKIPPED, completed_at=_now_iso()
                )
                self._notify(force=True)
                continue
            if self._should_skip(node_id):
                self._set_node(
                    node_id, status=NODE_STATUS_SKIPPED, completed_at=_now_iso()
                )
                self._notify(force=True)
                continue

            inputs, input_notes = self._gather_inputs(node)
            params = node.get("params", {}) or {}
            executor = NODE_TYPES.get(node["type"], {}).get("execute")
            self._set_node(
                node_id, status=NODE_STATUS_RUNNING, started_at=_now_iso(), progress=0.0
            )
            self._notify(force=True)
            if executor is None:
                self._set_node(
                    node_id,
                    status=NODE_STATUS_FAILED,
                    error=f"No executor for node type: {node.get('type')}",
                    completed_at=_now_iso(),
                )
                self._notify(force=True)
                continue

            self.ctx.on_progress = self._make_progress(node_id)
            try:
                result = executor(self.ctx, inputs, params)
                result = result if isinstance(result, dict) else {}
                # A reserved ``__note__`` key lets an executor flag a non-fatal
                # degraded outcome (e.g. Ollama unavailable, nothing wired) that
                # still completes — surfaced on the node, never stored as a result
                # port. Merge it with any adapter-coercion notes from gathering.
                notes = list(input_notes)
                exec_note = result.pop("__note__", None)
                if exec_note:
                    notes.append(str(exec_note))
                with self._lock:
                    self._results[node_id] = result
                # Persist the JSON-safe result ports (resume reloads them; the
                # run-history UI fetches the inspectable subset on demand even
                # after this runner is evicted from memory). ``hasResult`` only
                # advertises sidecars with something the inspector can render.
                if write_node_sidecar(
                    self.ctx.output_dir, self.run_id, node_id, node["type"], result
                ) and _inspectable_result(node["type"], result):
                    with self._lock:
                        self._sidecars.add(node_id)
                self._set_node(
                    node_id,
                    status=NODE_STATUS_COMPLETED,
                    progress=1.0,
                    completed_at=_now_iso(),
                    note="; ".join(notes) if notes else None,
                )
            except Exception as exc:
                self._set_node(
                    node_id,
                    status=NODE_STATUS_FAILED,
                    error=str(exc),
                    completed_at=_now_iso(),
                )
            self._notify(force=True)

        if self.ctx.cancel_event.is_set():
            self.status = RUN_STATUS_CANCELLED
        elif any(s["status"] == NODE_STATUS_FAILED for s in self.node_states.values()):
            self.status = RUN_STATUS_FAILED
        else:
            self.status = RUN_STATUS_COMPLETED
        self.completed_at = _now_iso()
        self._notify(force=True)

    # ---- serialization ----

    def snapshot(self) -> dict[str, Any]:
        """JSON-safe run record for the API/SSE/manifest (counts + status only)."""
        with self._lock:
            node_states = {
                nid: {
                    **{k: v for k, v in st.items() if not k.startswith("_")},
                    "hasResult": nid in self._sidecars,
                }
                for nid, st in self.node_states.items()
            }
            results = {
                nid: _node_result_summary(res) for nid, res in self._results.items()
            }
        return {
            "id": self.run_id,
            "blueprintId": self.blueprint_id,
            "batchId": self.batch_id,
            "participant": self.participant,
            "triggered": self.triggered,
            "status": self.status,
            "nodeStates": node_states,
            "results": results,
            "startedAt": self.started_at,
            "completedAt": self.completed_at,
        }
