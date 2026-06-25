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

See ``plans/WORKFLOWS-PLAN.md``.

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

import threading
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
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
# Execution context
# ---------------------------------------------------------------------------


@dataclass
class NodeContext:
    """Run-wide context handed to every node executor.

    Carries launch context (output dir + optional sheet) plus the uniform
    progress and cancellation seams. Executors forward ``cancel_flag`` to
    scan/clip callees and ``cancel_event`` to thinking-agent callees — the dual
    contract the backend functions already expose (see ``plans/WORKFLOWS-PLAN.md``
    and AGENTS.md). M4's ``WorkflowRunner`` builds one ``NodeContext`` per run; in
    M3 it is constructed directly by tests.
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


# ---------------------------------------------------------------------------
# Executors (M3) — thin adapters over existing pure functions
# ---------------------------------------------------------------------------
#
# Each executor has the uniform shape ``execute(ctx, inputs, params) -> {port:
# value}`` (keyed by OUTPUT-port name). Backend modules are imported lazily
# inside each executor (mirrors ``cli._run_ss_clips``) to avoid import cost and
# cycles — Workflows sits at the top of the dependency DAG. The concrete value
# carried on each wire is documented in ``plans/WORKFLOWS-PLAN.md``; the unifying
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

    result: Any = None
    if len(paths) >= 2:
        timeline = video.build_source_timeline(paths)
        if timeline is not None:
            result = transcripts.transcribe_timeline(
                timeline, language=lang, cancel_flag=ctx.cancel_flag
            )
    elif paths:
        result = transcripts.transcribe_video(
            paths[0], language=lang, cancel_flag=ctx.cancel_flag
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


# ---- Thinking (Ollama) ----


def _exec_summarize(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import ollama_client
    import thinking_agents

    transcript = inputs.get("transcript") or {}
    segments = transcript.get("segments") or []
    if not ollama_client.is_available():
        return {"summary": ""}
    summary = thinking_agents.summarize_transcript(
        segments, cancel_event=ctx.cancel_event
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
        return {"citations": []}
    cites = thinking_agents.find_citations(
        summary, segments, cancel_event=ctx.cancel_event
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
        return {"friction": []}
    scored = friction.score_segments(segments)
    candidates = friction.select_candidates(scored)
    moments = thinking_agents.find_friction_moments(
        summary, segments, candidates, cancel_event=ctx.cancel_event
    )
    return {"friction": moments or []}


# ---- Screenspace ----


def _exec_ss_scan(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import screenspace
    import screenspace_manifest
    import screenspace_worker
    import video

    src = inputs.get("video") or {}
    paths = list(src.get("video_paths") or [])
    tool_name = str(params.get("tool", "text") or "text")
    tool = screenspace.TOOLS.get(tool_name)
    if not paths or tool is None:
        return {"events": {"events": [], "source": src}}

    # Region: denormalize against the first part's dimensions, else whole frame.
    region_in = inputs.get("region") or {}
    region_name = str(region_in.get("name", "") or "")
    region_coords: dict[str, int] = {"x": 0, "y": 0, "w": 0, "h": 0}
    norm = region_in.get("coords")
    if isinstance(norm, dict) and norm:
        props = video.probe_video_properties(paths[0]) or {}
        w = int(props.get("width", 0) or 0)
        h = int(props.get("height", 0) or 0)
        if w > 0 and h > 0:
            region_coords = screenspace.denormalize_region(norm, w, h)

    task = screenspace_manifest.create_task(
        tool_name,
        str(src.get("participant", "") or ""),
        str(src.get("source_filename", "") or ""),
        paths,
        region_name,
        region_coords,
        parameters={},
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
    return {"events": {"events": events, "source": src}}


# ---- Artifact ----


def _exec_make_clips(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import pipeline

    description = str(params.get("description", "") or "").strip() or "workflow"

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
        return {"artifacts": {"artifacts": [], "study": study, "count": 0}}
    count, artifacts = pipeline.process_clips(
        records,
        output_format="clip",
        include_severity=False,
        cancel_flag=ctx.cancel_flag,
    )
    return {"artifacts": {"artifacts": artifacts, "study": study, "count": count}}


def _exec_build_reel(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import pipeline

    clips_in = inputs.get("clips") or {}
    records = list(clips_in.get("records") or [])
    study = str(clips_in.get("study", "") or "")
    if not records:
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "manifest": {"path": None, "records": []},
        }
    # Honor the node's reel name: reserve a unique output path (process_reel
    # treats a supplied output_file as a reservation and releases it on failure).
    name = utils.sanitize_filename(str(params.get("name", "") or "").strip()) or "reel"
    output_file = files.get_unique_filename(f"{name}{config.FILEFORMAT}")
    count, reels = pipeline.process_reel(
        records, output_file=output_file, cancel_flag=ctx.cancel_flag
    )
    return {
        "artifacts": {"artifacts": reels, "study": study, "count": count},
        "manifest": {"path": None, "records": reels},
    }


def _exec_timeline_viewer(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import viewer

    artifacts_in = inputs.get("artifacts") or {}
    artifacts = list(artifacts_in.get("artifacts") or [])
    study = str(artifacts_in.get("study", "") or "")
    ss_events = (inputs.get("events") or {}).get("events") or None
    data = viewer.finalize_timeline_data(
        artifacts, study=study, screenspace_events=ss_events
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


def _exec_gate(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    op = str(params.get("op", ">=") or ">=")
    fn = _GATE_OPS.get(op)
    raw = inputs.get("value")
    if fn is None or raw is None:
        return {"pass": False}
    try:
        result = fn(float(raw), float(params.get("threshold", 0)))
    except (TypeError, ValueError):
        result = False
    return {"pass": bool(result)}


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
    ("events", "timeRange"): _adapt_events_to_timerange,
    ("events", "clipRecords"): _adapt_events_to_cliprecords,
    ("video", "scalar"): _adapt_video_to_scalar,
}


# Wire each declarative NodeType to its executor. ``serialize_catalog`` strips
# ``execute`` again for the JSON catalog endpoint, so this stays server-internal.
_EXECUTORS: dict[
    str, Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]
] = {
    "video_source": _exec_video_source,
    "sheet_selection": _exec_sheet_selection,
    "region": _exec_region,
    "transcribe": _exec_transcribe,
    "find_word": _exec_find_word,
    "summarize": _exec_summarize,
    "citations": _exec_citations,
    "friction": _exec_friction,
    "ss_scan": _exec_ss_scan,
    "make_clips": _exec_make_clips,
    "build_reel": _exec_build_reel,
    "timeline_viewer": _exec_timeline_viewer,
    "gate": _exec_gate,
}

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
    ) -> None:
        self.run_id = run_id
        self.blueprint_id = str(blueprint.get("id", "") or "")
        self.nodes = list(blueprint.get("nodes", []))
        self.edges = list(blueprint.get("edges", []))
        self.ctx = ctx
        self.on_update = on_update or (lambda: None)
        self._nodes_by_id = {n["id"]: n for n in self.nodes}
        self.node_states: dict[str, dict[str, Any]] = {
            n["id"]: {
                "status": NODE_STATUS_QUEUED,
                "progress": 0.0,
                "error": None,
                "started_at": None,
                "completed_at": None,
            }
            for n in self.nodes
        }
        self._results: dict[str, dict[str, Any]] = {}
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

    def _gate_blocks(self, node_id: str) -> bool:
        """True if ``node_id`` is a gate that completed with ``pass`` False."""
        node = self._nodes_by_id.get(node_id)
        if not node or node.get("type") != "gate":
            return False
        return (self._results.get(node_id) or {}).get("pass") is False

    def _should_skip(self, node_id: str) -> bool:
        """A node is skipped if any upstream failed/skipped, or is a blocking gate."""
        for dep in self._deps(node_id):
            status = self.node_states[dep]["status"]
            if status in (NODE_STATUS_FAILED, NODE_STATUS_SKIPPED):
                return True
            if status == NODE_STATUS_COMPLETED and self._gate_blocks(dep):
                return True
        return False

    def _gather_inputs(self, node: dict[str, Any]) -> dict[str, Any]:
        """Map upstream results onto this node's input ports, applying adapters.

        Control edges (a gate's ``control`` output) establish a dependency but
        carry no data, so they are excluded here — they never clobber a real input.
        """
        inputs: dict[str, Any] = {}
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
                    except Exception as exc:  # adapter failure → empty input
                        utils.warning_print(
                            f"workflow adapter {out_type}->{in_type} failed: {exc}"
                        )
                        value = None
            inputs[to_port] = value
        return inputs

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

        for node_id in order:
            node = self._nodes_by_id[node_id]
            if self.ctx.cancel_event.is_set():
                self._set_node(
                    node_id, status=NODE_STATUS_SKIPPED, completed_at=_now_iso()
                )
                continue
            if self._should_skip(node_id):
                self._set_node(
                    node_id, status=NODE_STATUS_SKIPPED, completed_at=_now_iso()
                )
                self._notify(force=True)
                continue

            inputs = self._gather_inputs(node)
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
                with self._lock:
                    self._results[node_id] = result or {}
                self._set_node(
                    node_id,
                    status=NODE_STATUS_COMPLETED,
                    progress=1.0,
                    completed_at=_now_iso(),
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
                nid: {k: v for k, v in st.items() if not k.startswith("_")}
                for nid, st in self.node_states.items()
            }
            results = {
                nid: _node_result_summary(res) for nid, res in self._results.items()
            }
        return {
            "id": self.run_id,
            "blueprintId": self.blueprint_id,
            "status": self.status,
            "nodeStates": node_states,
            "results": results,
            "startedAt": self.started_at,
            "completedAt": self.completed_at,
        }
