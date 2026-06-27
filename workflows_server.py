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
import os
import queue
import shutil
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

# ---- Batch state (P3: whole-study fan-out) ----

# Live batch coordinators by id, symmetric to ``_runs`` (history is *derived* by
# grouping persisted runs on their ``batchId`` tag — no separate manifest key).
# Each value is ``{blueprintId, participants, runIds, cancel_event, status,
# createdAt}``. A batch SSE client is a (batch_id, queue) pair.
_batches: dict[str, dict[str, Any]] = {}
_batches_lock = threading.Lock()
_batch_sse_clients: list[tuple[str, queue.Queue[str]]] = []
_batch_sse_lock = threading.Lock()
_RUN_TERMINAL = {
    workflows.RUN_STATUS_COMPLETED,
    workflows.RUN_STATUS_FAILED,
    workflows.RUN_STATUS_CANCELLED,
}


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
    Only participants with an actual video file are listed, matching the set the
    batch endpoint will fan out over (so the dropdown never offers a participant a
    run can't resolve).
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
                "participants": [v["id"] for v in videos if v.get("has_video")],
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


# ---- Stash CRUD (M5: save/instantiate sub-graphs) ----
#
# A stash is a reusable sub-graph fragment ({id, name, nodes, edges, createdAt,
# builtin}). The server does CRUD only; the frontend instantiates a stash onto
# the canvas (id remap + position offset) client-side. ``GET`` prepends the
# read-only built-in recipes (P4) ahead of the user's persisted stashes. The
# same single-combined-manifest locking the blueprint routes use applies here.


@workflows_bp.route("/api/stashes")
def api_stashes() -> Any:
    """Return the built-in recipes (read-only) followed by the user's stashes."""
    with _manifest_lock:
        user_stashes = copy.deepcopy(_manifest.get("stashes", []))
    return jsonify({"ok": True, "stashes": workflows.BUILTIN_STASHES + user_stashes})


@workflows_bp.route("/api/stashes", methods=["POST"])
def api_stashes_create() -> Any:
    """Save a selected sub-graph as a named stash and return it with its id."""
    data = request.get_json(silent=True) or {}
    nodes = data.get("nodes", [])
    if not nodes:
        return jsonify({"ok": False, "error": "No nodes to stash"}), 400
    stash = {
        "id": "stash_" + uuid.uuid4().hex[:8],
        "name": (data.get("name") or "Stash").strip() or "Stash",
        "nodes": nodes,
        "edges": data.get("edges", []),
        "builtin": False,  # P4 built-ins are served from code; CRUD guards on this
        "createdAt": datetime.now(timezone.utc).isoformat(),
    }
    with _manifest_lock:
        _manifest.setdefault("stashes", []).append(stash)
        _persist_locked()
    return jsonify({"ok": True, "stash": stash})


@workflows_bp.route("/api/stashes/<stash_id>", methods=["PUT"])
def api_stashes_update(stash_id: str) -> Any:
    """Rename a user stash. Built-in recipes are read-only (403)."""
    if any(s["id"] == stash_id for s in workflows.BUILTIN_STASHES):
        return jsonify({"ok": False, "error": "Built-in recipes are read-only"}), 403
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "JSON body required"}), 400
    with _manifest_lock:
        stashes = _manifest.get("stashes", [])
        stash = next((s for s in stashes if s.get("id") == stash_id), None)
        if stash is None:
            return jsonify({"ok": False, "error": "Stash not found"}), 404
        if "name" in data:
            stash["name"] = (data["name"] or stash["name"]).strip() or stash["name"]
        _persist_locked()
    return jsonify({"ok": True, "stash": stash})


@workflows_bp.route("/api/stashes/<stash_id>", methods=["DELETE"])
def api_stashes_delete(stash_id: str) -> Any:
    """Delete a user stash by id. Built-in recipes are read-only (403)."""
    if any(s["id"] == stash_id for s in workflows.BUILTIN_STASHES):
        return jsonify({"ok": False, "error": "Built-in recipes are read-only"}), 403
    with _manifest_lock:
        stashes = _manifest.get("stashes", [])
        idx = next((i for i, s in enumerate(stashes) if s.get("id") == stash_id), None)
        if idx is None:
            return jsonify({"ok": False, "error": "Stash not found"}), 404
        stashes.pop(idx)
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


def _trim_run_history(runs: list[dict[str, Any]]) -> list[str]:
    """Cap persisted run history in place, evicting whole units oldest-first.

    A *unit* is one loose run or one batch's whole set of child runs (grouped by
    ``batchId``). Evicting by unit keeps a batch's children together — a partial
    batch would make its derived summary lie — and the newest unit plus any live
    batch are always kept, so a single large batch is never split or dropped (the
    cap yields to integrity). Without this, a 50+-participant batch would shed its
    earliest children and 404 on drill-in.

    Returns the run ids that were dropped, so the caller can prune their per-node
    result sidecars (P5) in lockstep.
    """
    if len(runs) <= _MAX_RUN_HISTORY:
        return []
    with _batches_lock:
        live = set(_batches)
    # Group into units preserving first-seen (≈ oldest-first) order.
    order: list[tuple[str, str]] = []  # ("batch"|"run", key)
    members: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for rec in runs:
        bid = rec.get("batchId")
        key = ("batch", bid) if bid else ("run", str(rec.get("id")))
        if key not in members:
            members[key] = []
            order.append(key)
        members[key].append(rec)
    kept = sum(len(members[k]) for k in order)
    drop: set[tuple[str, str]] = set()
    for i, key in enumerate(order):
        if kept <= _MAX_RUN_HISTORY or i == len(order) - 1:
            break  # under cap, or never drop the newest unit
        if key[0] == "batch" and key[1] in live:
            continue  # never evict a live batch's children
        drop.add(key)
        kept -= len(members[key])
    if not drop:
        return []
    dropped_ids = [str(r.get("id")) for k in drop for r in members[k] if r.get("id")]
    runs[:] = [r for k in order if k not in drop for r in members[k]]
    return dropped_ids


def _prune_run_sidecars(run_ids: list[str]) -> None:
    """Delete the per-node result sidecar dirs for evicted runs (best-effort)."""
    base = utils.get_effective_output_dir()
    for rid in run_ids:
        shutil.rmtree(workflows.run_results_dir(base, rid), ignore_errors=True)


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
        dropped = _trim_run_history(runs)
        _persist_locked()
    # Outside the manifest lock — filesystem cleanup mustn't hold it.
    if dropped:
        _prune_run_sidecars(dropped)


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
            # Evict the terminal runner: its summary now lives in the manifest
            # (``_run_snapshot`` falls back to it) and its inspectable per-node
            # results are already on disk as sidecars (written during ``run()``),
            # so holding the runner would only leak its full in-memory results.
            with _runs_lock:
                _runs.pop(run_id, None)

    threading.Thread(
        target=_run_and_finalize, daemon=True, name=f"workflow-{run_id}"
    ).start()
    return jsonify({"ok": True, "run": runner.snapshot()})


def _merged_runs() -> dict[str, dict[str, Any]]:
    """All run snapshots keyed by id (live runners override manifest history)."""
    merged: dict[str, dict[str, Any]] = {}
    with _manifest_lock:
        for record in _manifest.get("runs", []):
            merged[record.get("id")] = copy.deepcopy(record)
    with _runs_lock:
        live = list(_runs.items())
    for run_id, runner in live:
        merged[run_id] = runner.snapshot()
    return merged


@workflows_bp.route("/api/runs")
def api_runs_list() -> Any:
    """Recent runs (live runners override manifest history), newest first."""
    bp_filter = request.args.get("blueprintId")
    runs = list(_merged_runs().values())
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


@workflows_bp.route("/api/runs/<run_id>/nodes/<node_id>/result")
def api_run_node_result(run_id: str, node_id: str) -> Any:
    """Serve a node's inspectable result sidecar written by the runner (P5).

    Lazily fetched by the run-history UI on row-expand. Returns the raw stored
    payload (already JSON-sanitized at write time). 404 when no sidecar exists.
    """
    # Both ids are path segments — reject anything that could escape the run dir.
    safe_node = node_id == os.path.basename(node_id) and node_id not in (".", "..")
    safe_run = run_id == os.path.basename(run_id) and run_id not in (".", "..")
    if not (safe_node and safe_run):
        return jsonify({"ok": False, "error": "Invalid id"}), 404
    path = (
        workflows.run_results_dir(utils.get_effective_output_dir(), run_id)
        / f"{node_id}.json"
    )
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return jsonify({"ok": False, "error": "No result for node"}), 404
    return jsonify({"ok": True, "result": payload})


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


# ---- Batch lifecycle (P3: whole-study fan-out) ----


def _notify_batch_clients(batch_id: str) -> None:
    """Wake the SSE clients watching ``batch_id`` (coalesce on a full queue)."""
    with _batch_sse_lock:
        for bid, client_q in _batch_sse_clients:
            if bid != batch_id:
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


def _aggregate_batch_status(child_statuses: list[str], cancelled: bool) -> str:
    """Roll child run statuses up to one batch status.

    Any non-terminal child ⇒ the batch is still ``running`` (or ``queued`` before
    the first child starts). Once all children are terminal: ``cancelled`` if the
    batch was cancelled or any child was, else ``failed`` if any failed, else
    ``completed``.
    """
    if any(s not in _RUN_TERMINAL for s in child_statuses):
        return (
            workflows.RUN_STATUS_RUNNING
            if any(s == workflows.RUN_STATUS_RUNNING for s in child_statuses)
            else workflows.RUN_STATUS_QUEUED
        )
    if cancelled or any(s == workflows.RUN_STATUS_CANCELLED for s in child_statuses):
        return workflows.RUN_STATUS_CANCELLED
    if any(s == workflows.RUN_STATUS_FAILED for s in child_statuses):
        return workflows.RUN_STATUS_FAILED
    return workflows.RUN_STATUS_COMPLETED


def _batch_summary(batch_id: str) -> dict[str, Any] | None:
    """JSON-safe batch summary (counts + per-child status), live or historical.

    A live batch reads its identity from ``_batches``; a finished one is rebuilt by
    grouping the persisted runs that carry this ``batchId``. Returns ``None`` if no
    such batch exists in either place.
    """
    all_runs = _merged_runs()
    with _batches_lock:
        live = _batches.get(batch_id)
        record = dict(live) if live else None
        cancelled = bool(live and live["cancel_event"].is_set())
    if record is not None:
        run_ids = list(record.get("runIds", []))
        blueprint_id = record.get("blueprintId", "")
        participants = list(record.get("participants", []))
        created_at = record.get("createdAt")
    else:
        grouped = [
            r for r in all_runs.values() if r.get("batchId") == batch_id and r.get("id")
        ]
        if not grouped:
            return None
        grouped.sort(key=lambda r: r.get("startedAt") or "")
        run_ids = [r["id"] for r in grouped]
        blueprint_id = grouped[0].get("blueprintId", "")
        participants = [r.get("participant", "") for r in grouped]
        created_at = grouped[0].get("startedAt")

    children: list[dict[str, Any]] = []
    counts: dict[str, int] = {}
    child_statuses: list[str] = []
    for run_id, participant in zip(run_ids, participants):
        snap = all_runs.get(run_id)
        status = (snap or {}).get("status", workflows.RUN_STATUS_QUEUED)
        child_statuses.append(status)
        counts[status] = counts.get(status, 0) + 1
        children.append({"runId": run_id, "participant": participant, "status": status})
    status = _aggregate_batch_status(child_statuses, cancelled)
    return {
        "id": batch_id,
        "blueprintId": blueprint_id,
        "participants": participants,
        "createdAt": created_at,
        "status": status,
        "counts": counts,
        "children": children,
    }


def _sse_batch_payload(batch_id: str) -> str:
    """SSE ``data:`` line carrying the current batch summary."""
    summary = _batch_summary(batch_id)
    return "data: " + json.dumps({"ok": summary is not None, "batch": summary}) + "\n\n"


def _run_batch(batch_id: str, blueprint: dict[str, Any]) -> None:
    """Coordinator thread: run the blueprint once per participant, sequentially.

    Continue-on-error is mandatory — one bad participant must not sink the batch.
    A cancel signals the in-flight child (via its runner) and short-circuits every
    remaining child to ``cancelled``. Each child runner is evicted from ``_runs``
    once persisted (its summary lives in the manifest thereafter).
    """
    with _batches_lock:
        record = _batches.get(batch_id)
    if record is None:
        return
    cancel_event: threading.Event = record["cancel_event"]
    plan = list(zip(record["runIds"], record["participants"]))

    for run_id, participant in plan:
        child_cancel = threading.Event()
        if cancel_event.is_set():
            child_cancel.set()  # short-circuits run() to all-skipped + cancelled
        ctx = _build_node_context(child_cancel)
        child_bp = workflows.bind_participant(blueprint, participant)

        def _on_child_update(rid: str = run_id) -> None:
            _notify_run_clients(rid)
            _notify_batch_clients(batch_id)

        runner = workflows.WorkflowRunner(
            run_id,
            child_bp,
            ctx,
            on_update=_on_child_update,
            participant=participant,
            batch_id=batch_id,
        )
        with _runs_lock:
            _runs[run_id] = runner
        # A cancel that lands between the check above and run() still reaches this
        # child: the cancel endpoint cancels every live runner tagged to the batch.
        try:
            runner.run()
        except Exception as exc:  # belt-and-suspenders; run() catches per node
            utils.error_print(f"workflow batch child {participant} crashed: {exc}")
        _persist_run(runner.snapshot())
        _notify_run_clients(run_id)
        _notify_batch_clients(batch_id)
        with _runs_lock:
            _runs.pop(run_id, None)

    with _batches_lock:
        _batches.pop(batch_id, None)
    _notify_batch_clients(batch_id)


@workflows_bp.route("/api/batches", methods=["POST"])
def api_batch_create() -> Any:
    """Fan a blueprint out across participants, one sequential run each (P3)."""
    data = request.get_json(silent=True) or {}
    bp_id = data.get("blueprintId")
    with _manifest_lock:
        blueprint = next(
            (b for b in _manifest.get("blueprints", []) if b.get("id") == bp_id), None
        )
        blueprint = copy.deepcopy(blueprint) if blueprint else None
    if blueprint is None:
        return jsonify({"ok": False, "error": "Blueprint not found"}), 404
    bp_id = str(blueprint.get("id", "") or "")
    if not workflows.blueprint_participant_nodes(blueprint):
        return (
            jsonify(
                {"ok": False, "error": "Blueprint has no Video Source to fan out over"}
            ),
            400,
        )
    try:
        workflows.topo_order(blueprint.get("nodes", []), blueprint.get("edges", []))
    except workflows.WorkflowCycleError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    available = [
        v["id"] for v in utils.discover_participant_videos() if v.get("has_video")
    ]
    requested = data.get("participants")
    if requested:
        participants = [p for p in requested if p in available]
    else:
        participants = available
    if not participants:
        return jsonify({"ok": False, "error": "No participants with video found"}), 400

    batch_id = "batch_" + uuid.uuid4().hex[:8]
    # Child runs are persisted as they execute, not up front: pre-persisting one
    # queued record per participant would flood the run-history cap and could evict
    # this batch's own not-yet-started children. The live ``_batches`` record makes
    # the batch (and its queued children) visible immediately via ``_batch_summary``.
    run_ids = ["run_" + uuid.uuid4().hex[:8] for _ in participants]
    with _batches_lock:
        _batches[batch_id] = {
            "id": batch_id,
            "blueprintId": bp_id,
            "participants": participants,
            "runIds": run_ids,
            "cancel_event": threading.Event(),
            "createdAt": datetime.now(timezone.utc).isoformat(),
        }

    threading.Thread(
        target=_run_batch,
        args=(batch_id, blueprint),
        daemon=True,
        name=f"workflow-batch-{batch_id}",
    ).start()
    return jsonify({"ok": True, "batch": _batch_summary(batch_id)})


@workflows_bp.route("/api/batches")
def api_batches_list() -> Any:
    """Recent batches (derived by grouping runs on ``batchId``), newest first."""
    bp_filter = request.args.get("blueprintId")
    batch_ids: list[str] = []
    seen: set[str] = set()
    with _batches_lock:
        for bid in _batches:
            if bid not in seen:
                batch_ids.append(bid)
                seen.add(bid)
    for run in _merged_runs().values():
        bid = run.get("batchId")
        if bid and bid not in seen:
            batch_ids.append(bid)
            seen.add(bid)
    summaries = [s for s in (_batch_summary(bid) for bid in batch_ids) if s]
    if bp_filter:
        summaries = [s for s in summaries if s.get("blueprintId") == bp_filter]
    summaries.sort(key=lambda s: s.get("createdAt") or "", reverse=True)
    return jsonify({"ok": True, "batches": summaries})


@workflows_bp.route("/api/batches/<batch_id>")
def api_batch_get(batch_id: str) -> Any:
    """One batch's summary + per-child run snapshots (drill-in)."""
    summary = _batch_summary(batch_id)
    if summary is None:
        return jsonify({"ok": False, "error": "Batch not found"}), 404
    all_runs = _merged_runs()
    runs = [all_runs[c["runId"]] for c in summary["children"] if c["runId"] in all_runs]
    return jsonify({"ok": True, "batch": summary, "runs": runs})


@workflows_bp.route("/api/batches/<batch_id>/cancel", methods=["POST"])
def api_batch_cancel(batch_id: str) -> Any:
    """Cancel a batch: stop launching children and cancel the in-flight one."""
    with _batches_lock:
        record = _batches.get(batch_id)
    if record is None:
        return jsonify({"ok": False, "error": "Batch not found or finished"}), 404
    record["cancel_event"].set()
    with _runs_lock:
        for runner in _runs.values():
            if getattr(runner, "batch_id", "") == batch_id:
                runner.cancel()
    return jsonify({"ok": True})


@workflows_bp.route("/api/batches/<batch_id>/stream")
def api_batch_stream(batch_id: str) -> Response:
    """SSE stream of one batch's summary (mirrors :func:`api_run_stream`)."""
    client_q: queue.Queue[str] = queue.Queue(maxsize=64)
    with _batch_sse_lock:
        _batch_sse_clients.append((batch_id, client_q))

    def generate():  # type: ignore[no-untyped-def]
        try:
            yield _sse_batch_payload(batch_id)
            while True:
                try:
                    client_q.get(timeout=15)
                    while not client_q.empty():
                        try:
                            client_q.get_nowait()
                        except queue.Empty:
                            break
                    yield _sse_batch_payload(batch_id)
                except queue.Empty:
                    yield ": keepalive\n\n"
        except GeneratorExit:
            pass
        finally:
            with _batch_sse_lock:
                try:
                    _batch_sse_clients.remove((batch_id, client_q))
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
