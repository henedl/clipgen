"""Workflows Flask blueprint — serves the node-canvas page and its REST API.

Registered at ``/workflows`` by ``server.build_combined_app`` (mutually exclusive
launch with the other web modes, but all blueprints are always mounted). M1
ships the static page routes, module-state init, the ``/api/catalog`` node
registry, and full blueprint CRUD (the canvas autosave target). M4 adds the run
lifecycle: ``POST /api/runs`` spawns a :class:`workflows.WorkflowRunner` on a
daemon thread, with per-run SSE (``/api/runs/<id>/stream``) + a polling fallback
(``GET /api/runs/<id>``), mirroring ``screenspace_server``'s task stream.

Module-level state (``_input_dir``, ``_sheet_context``, ``_worksheet``,
``_manifest``, ``_runs``) is initialized by :func:`_init_workflows_state`,
mirroring the Screenspace and Transcripts blueprints. Mutations hold
``_manifest_lock`` and persist via :func:`_persist_locked` (mirrors
``screenspace_server._do_persist``). Live runner progress stays in ``_runs``;
the manifest is written only at run creation + terminal, never per progress tick.
"""

from __future__ import annotations

import copy
import json
import queue
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from flask import Blueprint, Response, jsonify, request

import utils
import workflows

# ---- Module state (initialized by _init_workflows_state) ----

_input_dir: str = ""
_sheet_context: Any = None
_worksheet: Any = None
_manifest: dict[str, Any] = {}
_manifest_lock = threading.Lock()

# ---- Run state (M4) ----

# Live runners by id (authoritative for in-flight progress); the manifest holds
# the persisted history. SSE clients are (run_id, queue) pairs scoped to one run.
_runs: dict[str, workflows.WorkflowRunner] = {}
_runs_lock = threading.Lock()
_sse_clients: list[tuple[str, queue.Queue[str]]] = []
_sse_lock = threading.Lock()
_MAX_RUN_HISTORY = 50  # cap persisted runs (small ephemeral tool; keep most recent)


# ---- Blueprint ----

workflows_bp = Blueprint("workflows", __name__)

utils.register_static_routes(
    workflows_bp,
    "workflows.html",
    media_dir_getter=lambda: _input_dir,
    media_error="Input directory not configured",
    icons=True,
)


def _persist_locked() -> None:
    """Write the workflows manifest to disk. Caller must hold ``_manifest_lock``.

    Passes the live ``stashes``/``runs`` through so a blueprint write never
    clobbers them (the save helper rewrites the whole file).
    """
    workflows.save_workflows_manifest(
        _manifest.get("blueprints", []),
        _manifest.get("stashes", []),
        _manifest.get("runs", []),
    )


# ---- Catalog ----


@workflows_bp.route("/api/catalog")
def api_catalog() -> Any:
    """Serialized node-type catalog + launch-context flags (for palette grey-out).

    ``context.sheet`` / ``context.videoDir`` tell the frontend which nodes'
    ``requires`` are satisfiable in this launch, so it can disable the rest;
    ``context.participants`` populates the Video-Source participant dropdown
    (reusing the one discovery call ``videoDir`` already needs — no extra route).
    """
    videos = utils.discover_participant_videos()
    return jsonify(
        {
            "ok": True,
            "catalog": workflows.serialize_catalog(),
            # Adapter pairs the runner coerces across (events→clipRecords, …) so
            # the frontend's canConnect accepts the same wires the runner runs.
            "adapters": workflows.serialize_adapters(),
            "context": {
                "sheet": _sheet_context is not None,
                "videoDir": bool(videos),
                "participants": [v["id"] for v in videos],
            },
        }
    )


# ---- Blueprint CRUD (the canvas autosave target) ----


@workflows_bp.route("/api/blueprints")
def api_blueprints() -> Any:
    """Return the persisted workflow blueprints."""
    with _manifest_lock:
        blueprints = copy.deepcopy(_manifest.get("blueprints", []))
    return jsonify({"ok": True, "blueprints": blueprints})


@workflows_bp.route("/api/blueprints", methods=["POST"])
def api_blueprints_create() -> Any:
    """Create a new (usually empty) blueprint and return it with its assigned id."""
    data = request.get_json(silent=True) or {}
    blueprint = {
        "id": "bp_" + uuid.uuid4().hex[:8],
        "name": (data.get("name") or "Untitled").strip() or "Untitled",
        "nodes": data.get("nodes", []),
        "edges": data.get("edges", []),
        "viewport": data.get("viewport", {"x": 0, "y": 0, "zoom": 1}),
        "trigger": None,  # reserved for the deferred auto-launch phase
        "createdAt": datetime.now(timezone.utc).isoformat(),
    }
    with _manifest_lock:
        _manifest.setdefault("blueprints", []).append(blueprint)
        _persist_locked()
    return jsonify({"ok": True, "blueprint": blueprint})


@workflows_bp.route("/api/blueprints/<bp_id>", methods=["PUT"])
def api_blueprints_update(bp_id: str) -> Any:
    """Update a blueprint's name/nodes/edges/viewport (the debounced autosave)."""
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "JSON body required"}), 400
    with _manifest_lock:
        blueprints = _manifest.get("blueprints", [])
        blueprint = next((b for b in blueprints if b.get("id") == bp_id), None)
        if blueprint is None:
            return jsonify({"ok": False, "error": "Blueprint not found"}), 404
        for key in ("name", "nodes", "edges", "viewport"):
            if key in data:
                blueprint[key] = data[key]
        _persist_locked()
    return jsonify({"ok": True, "blueprint": blueprint})


@workflows_bp.route("/api/blueprints/<bp_id>", methods=["DELETE"])
def api_blueprints_delete(bp_id: str) -> Any:
    """Delete a blueprint by id."""
    with _manifest_lock:
        blueprints = _manifest.get("blueprints", [])
        idx = next((i for i, b in enumerate(blueprints) if b.get("id") == bp_id), None)
        if idx is None:
            return jsonify({"ok": False, "error": "Blueprint not found"}), 404
        blueprints.pop(idx)
        _persist_locked()
    return jsonify({"ok": True})


# ---- Run lifecycle (M4) ----


def _build_node_context(cancel_event: threading.Event) -> workflows.NodeContext:
    """Build the per-run ``NodeContext`` from the active launch context."""
    input_dir = Path(_input_dir) if _input_dir else utils.get_effective_input_dir()
    return workflows.NodeContext(
        input_dir=input_dir,
        output_dir=utils.get_effective_output_dir(),
        sheet_context=_sheet_context,
        worksheet=_worksheet,
        cancel_event=cancel_event,
    )


def _run_snapshot(run_id: str) -> dict[str, Any] | None:
    """Live runner snapshot if in-flight, else the persisted manifest record."""
    with _runs_lock:
        runner = _runs.get(run_id)
    if runner is not None:
        return runner.snapshot()
    with _manifest_lock:
        for record in _manifest.get("runs", []):
            if record.get("id") == run_id:
                return copy.deepcopy(record)
    return None


def _persist_run(snapshot: dict[str, Any] | None) -> None:
    """Upsert a run snapshot into the manifest history (capped, most-recent kept)."""
    if not snapshot:
        return
    with _manifest_lock:
        runs = _manifest.setdefault("runs", [])
        idx = next(
            (i for i, r in enumerate(runs) if r.get("id") == snapshot.get("id")), None
        )
        if idx is None:
            runs.append(snapshot)
        else:
            runs[idx] = snapshot
        if len(runs) > _MAX_RUN_HISTORY:
            del runs[0 : len(runs) - _MAX_RUN_HISTORY]
        _persist_locked()


def _notify_run_clients(run_id: str) -> None:
    """Wake the SSE clients watching ``run_id`` (coalesce on a full queue).

    Mirrors ``screenspace_server._notify_sse_clients``: on overflow discard one
    stale entry and leave a single ``update`` marker so the client still re-emits
    fresh run state once it catches up.
    """
    with _sse_lock:
        for rid, client_q in _sse_clients:
            if rid != run_id:
                continue
            try:
                client_q.put_nowait("update")
            except queue.Full:
                try:
                    client_q.get_nowait()
                except queue.Empty:
                    pass
                try:
                    client_q.put_nowait("update")
                except queue.Full:
                    pass


def _sse_run_payload(run_id: str) -> str:
    """SSE ``data:`` line carrying the current run snapshot."""
    snap = _run_snapshot(run_id)
    return "data: " + json.dumps({"ok": snap is not None, "run": snap}) + "\n\n"


@workflows_bp.route("/api/runs", methods=["POST"])
def api_run_create() -> Any:
    """Validate a blueprint's DAG, spawn a runner thread, return the run snapshot."""
    data = request.get_json(silent=True) or {}
    bp_id = data.get("blueprintId")
    with _manifest_lock:
        blueprint = next(
            (b for b in _manifest.get("blueprints", []) if b.get("id") == bp_id), None
        )
        blueprint = copy.deepcopy(blueprint) if blueprint else None
    if blueprint is None:
        return jsonify({"ok": False, "error": "Blueprint not found"}), 404
    try:
        workflows.topo_order(blueprint.get("nodes", []), blueprint.get("edges", []))
    except workflows.WorkflowCycleError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    run_id = "run_" + uuid.uuid4().hex[:8]
    cancel_event = threading.Event()
    ctx = _build_node_context(cancel_event)
    runner = workflows.WorkflowRunner(
        run_id, blueprint, ctx, on_update=lambda: _notify_run_clients(run_id)
    )
    with _runs_lock:
        _runs[run_id] = runner
    _persist_run(runner.snapshot())

    def _run_and_finalize() -> None:
        try:
            runner.run()
        finally:
            _persist_run(runner.snapshot())
            _notify_run_clients(run_id)

    threading.Thread(
        target=_run_and_finalize, daemon=True, name=f"workflow-{run_id}"
    ).start()
    return jsonify({"ok": True, "run": runner.snapshot()})


@workflows_bp.route("/api/runs")
def api_runs_list() -> Any:
    """Recent runs (live runners override manifest history), newest first."""
    bp_filter = request.args.get("blueprintId")
    merged: dict[str, dict[str, Any]] = {}
    with _manifest_lock:
        for record in _manifest.get("runs", []):
            merged[record.get("id")] = copy.deepcopy(record)
    with _runs_lock:
        live = list(_runs.items())
    for run_id, runner in live:
        merged[run_id] = runner.snapshot()
    runs = list(merged.values())
    if bp_filter:
        runs = [r for r in runs if r.get("blueprintId") == bp_filter]
    runs.sort(key=lambda r: r.get("startedAt") or "", reverse=True)
    return jsonify({"ok": True, "runs": runs})


@workflows_bp.route("/api/runs/<run_id>")
def api_run_get(run_id: str) -> Any:
    """Polling fallback for one run's live/persisted snapshot."""
    snap = _run_snapshot(run_id)
    if snap is None:
        return jsonify({"ok": False, "error": "Run not found"}), 404
    return jsonify({"ok": True, "run": snap})


@workflows_bp.route("/api/runs/<run_id>/cancel", methods=["POST"])
def api_run_cancel(run_id: str) -> Any:
    """Signal an in-flight run's cancel event (no-op once finished)."""
    with _runs_lock:
        runner = _runs.get(run_id)
    if runner is None:
        return jsonify({"ok": False, "error": "Run not found or finished"}), 404
    runner.cancel()
    return jsonify({"ok": True})


@workflows_bp.route("/api/runs/<run_id>/stream")
def api_run_stream(run_id: str) -> Response:
    """SSE stream of one run's snapshot (mirrors screenspace_server.api_tasks_stream)."""
    client_q: queue.Queue[str] = queue.Queue(maxsize=64)
    with _sse_lock:
        _sse_clients.append((run_id, client_q))

    def generate():  # type: ignore[no-untyped-def]
        try:
            yield _sse_run_payload(run_id)
            while True:
                try:
                    client_q.get(timeout=15)
                    while not client_q.empty():
                        try:
                            client_q.get_nowait()
                        except queue.Empty:
                            break
                    yield _sse_run_payload(run_id)
                except queue.Empty:
                    yield ": keepalive\n\n"
        except GeneratorExit:
            pass
        finally:
            with _sse_lock:
                try:
                    _sse_clients.remove((run_id, client_q))
                except ValueError:
                    pass

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


def _init_workflows_state(
    sheet_context: Any = None,
    participant_list: list[str] | None = None,
    worksheet: Any = None,
) -> None:
    """Initialize module-level state for Workflows routes.

    Loads the workflows manifest and records the active input dir + sheet
    context + worksheet (the latter feeds the ``sheet_selection`` executor).
    ``participant_list`` is accepted for parity with the other blueprints' init
    signatures; per-participant video paths are resolved on demand.
    """
    global _input_dir, _sheet_context, _worksheet, _manifest  # noqa: PLW0603

    _input_dir = str(utils.get_effective_input_dir())
    _sheet_context = sheet_context
    _worksheet = worksheet
    _manifest = workflows.load_workflows_manifest()
