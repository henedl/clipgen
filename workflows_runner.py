"""Workflows run engine: DAG topo-sort + sequential ready-set execution.

The engine half of the workflows engine, split out of ``workflows.py`` (which
keeps the executors, the import-time wiring, and the facade; the declarative
catalog lives in ``workflows_catalog``). Owns the run/node status constants,
the auto-run trigger types, ``topo_order``/``bind_participant``, per-node
result sidecars, resume planning, and ``WorkflowRunner``.

Reads ``NODE_TYPES[...]["execute"]`` and ``ADAPTERS`` only at call time —
after ``workflows.py``'s import-time wiring has attached the executors — so
importing ``workflows`` (the facade) remains the supported entry point for
running graphs.
"""

from __future__ import annotations

import copy
import json
import os
import threading
import time
from collections.abc import Callable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import utils
from workflows_catalog import ADAPTERS, NODE_TYPES, NodeContext

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

# Run + per-node status constants. Deliberately duplicated from
# screenspace_manifest's TASK_STATUS_* (and transcripts' status strings): the
# only viable import direction would drag screenspace_tools' top-level cv2
# into the workflows import chain, and a shared module for five strings
# fails the repo's minimalism bar. Keep in sync by eye.
RUN_STATUS_QUEUED = "queued"
RUN_STATUS_RUNNING = "running"
RUN_STATUS_COMPLETED = "completed"
# Every node ran, but at least one produced a result we know is incomplete (an
# input that failed to coerce, or a result sidecar that could not be written).
# Distinct from COMPLETED so the run history can't show green over lost data,
# and distinct from FAILED because the outputs that did land are usable.
RUN_STATUS_DEGRADED = "degraded"
RUN_STATUS_FAILED = "failed"
RUN_STATUS_CANCELLED = "cancelled"

NODE_STATUS_QUEUED = "queued"
NODE_STATUS_RUNNING = "running"
NODE_STATUS_COMPLETED = "completed"
NODE_STATUS_FAILED = "failed"
NODE_STATUS_DEGRADED = "degraded"
NODE_STATUS_SKIPPED = "skipped"

# Canvas-only sticky-note pseudo-node (frontend-created, not in NODE_TYPES).
# Notes live in blueprint["nodes"] so they ride save/undo/copy/import for free;
# the runner filters them out so they never execute or appear in run snapshots.
NOTE_NODE_TYPE = "note"

# Auto-run trigger types: new_video (the original watch-dir P6 trigger) plus
# the chaining triggers (a transcript or Screenspace scan completing fires an
# armed blueprint for that participant). Served through /api/catalog context so
# the frontend picker never duplicates the list; workflows_server's watcher
# polls each type's source (input dir / transcripts manifest / screenspace
# manifest) only while a blueprint of that type is armed.
TRIGGER_TYPES: list[dict[str, str]] = [
    {"id": "new_video", "label": "New video lands"},
    {"id": "transcript_complete", "label": "Transcript completes"},
    {"id": "scan_event", "label": "Screenspace scan completes"},
]
TRIGGER_TYPE_IDS = frozenset(t["id"] for t in TRIGGER_TYPES)

_PROGRESS_NOTIFY_INTERVAL = (
    0.5  # seconds; throttle SSE notifies (copy screenspace_worker)
)


class WorkflowCycleError(ValueError):
    """Raised by :func:`topo_order` when the graph is not a DAG (rejected at submit)."""


def _now_iso() -> str:
    """UTC ISO-8601 timestamp for run/node start+complete stamps."""
    return datetime.now(UTC).isoformat()


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
        # A wire missing either endpoint (or carrying a non-string one) is
        # malformed, not just stale — skip it for the same reason. `in id_set`
        # below already excluded these (None is never a node id), but only
        # incidentally: it reads as a membership test, and ty cannot use it to
        # narrow `.get()`'s Optional. This states the contract instead.
        if not isinstance(src, str) or not isinstance(dst, str):
            continue
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
) -> str:
    """Atomically write a node's JSON-safe result ports to its run sidecar.

    Persists every ``_SIDECAR_PORT_TYPES`` port plus a self-describing
    ``__type__`` key (consumed by resume + the read-time inspectable filter).
    Returns ``"written"`` when a sidecar now exists, ``"empty"`` when there was
    nothing to persist (bad ``node_id`` or no sidecar-able ports — not a problem),
    and ``"failed"`` when the write itself errored. The caller must tell those
    last two apart: collapsing them hides a full disk behind a missing
    ``hasResult`` badge, and a later resume then silently re-executes every node
    because no sidecar exists to seed from. JSON-sanitizes via
    :func:`utils.sanitize_floats` (non-finite floats / numpy scalars).
    """
    if not node_id or node_id != os.path.basename(node_id) or node_id in (".", ".."):
        return "empty"
    payload = _filter_result_ports(node_type_id, result, _SIDECAR_PORT_TYPES)
    if not payload:
        return "empty"
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
        return "written"
    except (OSError, TypeError, ValueError) as exc:
        utils.warning_print(f"workflow sidecar write failed ({node_id}): {exc}")
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        return "failed"


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
        trigger_type: str = "",
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
        # Auto-run triggers (P6 + chaining): True when this run was launched by
        # the watcher (surfaced as a badge in the run history); ``trigger_type``
        # records which trigger fired it (new_video / transcript_complete / …).
        self.triggered = triggered
        self.trigger_type = trigger_type
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
                if status in (
                    NODE_STATUS_COMPLETED,
                    NODE_STATUS_DEGRADED,
                ) and self._gate_blocks(dep):
                    return True
                continue
            # A data edge: only a *required* input's dead producer forces a skip.
            if status in (
                NODE_STATUS_FAILED,
                NODE_STATUS_SKIPPED,
            ) and not _port_optional(node_type, edge.get("toPort")):
                return True
        return False

    def _gather_inputs(
        self, node: dict[str, Any]
    ) -> tuple[dict[str, Any], list[str], bool]:
        """Map upstream results onto this node's input ports, applying adapters.

        Returns ``(inputs, notes, degraded)``. ``notes`` carries any adapter-failure
        messages so the runner can surface them on the node — a coercion that
        raises otherwise degrades the input to ``None`` invisibly (only a server
        log). ``degraded`` is True when that happened: the node still runs, but on
        less data than the graph promised it, so the run must not report a clean
        completion. Control edges (a gate's ``control`` output) establish a
        dependency but carry no data, so they are excluded here — they never
        clobber a real input.
        """
        inputs: dict[str, Any] = {}
        notes: list[str] = []
        degraded = False
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
                        degraded = True
            inputs[to_port] = value
        return inputs, notes, degraded

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
            except Exception as exc:
                # A broken progress listener must never abort the run itself.
                utils.verbose_print(f"workflow progress notify failed: {exc}")

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
                # Compare against "written" explicitly: every return value is a
                # truthy string, so a truthiness check here would advertise a
                # hasResult badge for a node whose sidecar was never written
                # (404 when the inspector fetches it) and hide a failed write
                # behind a green COMPLETED — the same failure the execute path
                # below surfaces as DEGRADED.
                sidecar = write_node_sidecar(
                    self.ctx.output_dir, self.run_id, node_id, node["type"], seeded
                )
                if sidecar == "written" and _inspectable_result(node["type"], seeded):
                    with self._lock:
                        self._sidecars.add(node_id)
                seed_notes = [self._seed_note] if self._seed_note else []
                if sidecar == "failed":
                    seed_notes.append("Result sidecar could not be written")
                self._set_node(
                    node_id,
                    status=(
                        NODE_STATUS_DEGRADED
                        if sidecar == "failed"
                        else NODE_STATUS_COMPLETED
                    ),
                    progress=1.0,
                    completed_at=_now_iso(),
                    note="; ".join(seed_notes) if seed_notes else None,
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

            inputs, input_notes, inputs_degraded = self._gather_inputs(node)
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
                sidecar = write_node_sidecar(
                    self.ctx.output_dir, self.run_id, node_id, node["type"], result
                )
                if sidecar == "written" and _inspectable_result(node["type"], result):
                    with self._lock:
                        self._sidecars.add(node_id)
                if sidecar == "failed":
                    notes.append("Result sidecar could not be written")
                degraded = inputs_degraded or sidecar == "failed"
                self._set_node(
                    node_id,
                    status=(
                        NODE_STATUS_DEGRADED if degraded else NODE_STATUS_COMPLETED
                    ),
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
        elif any(
            s["status"] == NODE_STATUS_DEGRADED for s in self.node_states.values()
        ):
            self.status = RUN_STATUS_DEGRADED
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
            "triggerType": self.trigger_type,
            "status": self.status,
            "nodeStates": node_states,
            "results": results,
            "startedAt": self.started_at,
            "completedAt": self.completed_at,
        }
