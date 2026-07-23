"""Workflows node catalog: declarative registry + typed-port adapters.

The data half of the workflows engine, split out of ``workflows.py`` (which
keeps the executors, the import-time wiring, and the facade; the run engine
lives in ``workflows_runner``). Owns:

* ``NodeContext`` — the run-wide context handed to every node executor.
* ``NODE_TYPES`` — the declarative single-source-of-truth node registry
  (typed ports + param schema); ``serialize_catalog`` feeds ``/api/catalog``.
* ``BUILTIN_STASHES`` — the read-only built-in recipe graphs.
* The source-descriptor helpers embedded in every domain value.
* ``ADAPTERS`` + ``_ADAPTER_DESCRIPTIONS`` — the typed-port coercion table the
  runner applies (pure ``value -> value``; late-imports ``files``/``video``).

**Wiring caveat:** the collection-algebra node types and every node's
``execute`` key are attached by ``workflows.py`` at import time (it mutates
the ``NODE_TYPES`` dict imported from here). Import ``workflows``, not this
module, before running graphs — this module alone is an unwired catalog.
"""

from __future__ import annotations

import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, NotRequired, TypedDict, cast

import config
import utils

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
# Source descriptors — embedded in every domain value so the pure adapters
# (below) and the executors (in workflows.py) can reach a value's source
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Typed-port adapters (M3) — pure value -> value, applied by the runner (M4)
# ---------------------------------------------------------------------------
#
# An adapter coerces an output value whose port type differs from the consuming
# input's type. They are pure (no ctx/params), so each value self-carries its
# source descriptor (see the executors in ``workflows.py``).


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
