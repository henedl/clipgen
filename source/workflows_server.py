"""Workflows Flask blueprint — serves the node-canvas page and its REST API.

Registered at ``/workflows`` by ``server.build_combined_app``. Static page
routes, module-state init, the ``/api/catalog`` node registry, blueprint CRUD
(the canvas autosave target), and the run lifecycle: ``POST /api/runs`` spawns a
:class:`workflows.WorkflowRunner` on a daemon thread, with per-run SSE
(``/api/runs/<id>/stream``) plus a polling fallback, mirroring
``screenspace_server``'s task stream.

Module state (``_sheet_context``, ``_worksheet``, ``_manifest``, ``_runs``) is
initialized by :func:`_init_workflows_state`, its sheet half re-pointed on a
worksheet swap by :func:`repin_sheet_state`. Mutations hold ``_manifest_lock``
and persist via :func:`_persist_locked`. Live runner progress stays in ``_runs``;
the manifest is written only at run creation and terminal, never per tick.
"""

from __future__ import annotations

import copy
import json
import os
import shutil
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from flask import Blueprint, Response, request

import config
import utils
import workflows
from server_utils import err, find_by_id, make_sse_channel, ok, remove_by_id

# ---- Module state (initialized by _init_workflows_state) ----

_sheet_context: Any = None
_worksheet: Any = None
_manifest: dict[str, Any] = {}
_manifest_lock = threading.Lock()

# ---- Run state ----

# Live runners by id (authoritative for in-flight progress); the manifest holds
# the persisted history. SSE clients are (run_id, queue) pairs scoped to one run.
_runs: dict[str, workflows.WorkflowRunner] = {}
_runs_lock = threading.Lock()
# SSE clients scoped to one run: notify with the run_id key. ``_sse_clients`` is
# the channel's live registry (``(run_id, queue)`` tuples); see make_sse_channel.
_notify_run_clients, _run_stream, _sse_clients = make_sse_channel()
_MAX_RUN_HISTORY = 50  # cap persisted runs (small ephemeral tool; keep most recent)

# ---- Batch state (whole-study fan-out) ----

# Live batch coordinators by id, symmetric to ``_runs`` (history is *derived* by
# grouping persisted runs on their ``batchId`` tag — no separate manifest key).
# Each value is ``{blueprintId, participants, runIds, cancel_event, status,
# createdAt}``. A batch SSE client is a (batch_id, queue) pair.
_batches: dict[str, dict[str, Any]] = {}
_batches_lock = threading.Lock()
# A batch SSE client is a (batch_id, queue) pair; notify with the batch_id key.
_notify_batch_clients, _batch_stream, _batch_sse_clients = make_sse_channel()
_RUN_TERMINAL = {
    workflows.RUN_STATUS_COMPLETED,
    workflows.RUN_STATUS_DEGRADED,
    workflows.RUN_STATUS_FAILED,
    workflows.RUN_STATUS_CANCELLED,
}

# ---- Auto-run trigger state ----
#
# One polling daemon thread checks each trigger type's source while a blueprint
# of that type is armed — the input dir for new videos, the transcripts manifest
# for fresh ``transcribed_at`` stamps, the screenspace manifest for completed
# tasks — firing one run per arrival. Baselines seed at startup and re-seed on
# arm so the pre-existing backlog never fires, and a new video must stat
# identically across two consecutive polls (the partial-copy guard). With nothing
# armed a tick does no I/O at all.
_watch_seen: set[str] = set()  # pids already accounted for (never fire again)
_watch_pending: dict[str, tuple[int, float]] = {}  # pid -> last-poll (size, mtime)
# Chaining-trigger baselines: completions already accounted for (never re-fire).
_watch_transcript_baseline: dict[str, str] = {}  # pid -> transcribed_at stamp
_watch_scan_seen: set[str] = set()  # completed screenspace task ids
# mtime-gated parse caches so an unchanged manifest is never re-read per poll.
_watch_transcript_cache: tuple[float, dict[str, str]] = (-1.0, {})
_watch_scan_cache: tuple[float, dict[str, str]] = (-1.0, {})
_watch_lock = threading.Lock()
_watch_thread: threading.Thread | None = None
_watch_stop = threading.Event()  # tests only; production never sets it


# ---- Blueprint ----

workflows_bp = Blueprint("workflows", __name__)

utils.register_static_routes(
    workflows_bp,
    "workflows.html",
    # Per request, not a snapshot — POST /api/dirs moves config.INPUT_DIR
    # mid-session and never re-inits this blueprint. See transcripts_bp.
    media_dir_getter=lambda: str(utils.get_effective_input_dir()),
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
    return ok(
        # Bootstrap channel for shared frontend config (hotkey overrides etc.);
        # this page has no sheet-data fetch, so the config rides along here.
        config=utils.get_frontend_config(),
        catalog=workflows.serialize_catalog(),
        # Adapter pairs the runner coerces across (events→clipRecords, …) so
        # the frontend's canConnect accepts the same wires the runner runs.
        adapters=workflows.serialize_adapters(),
        context={
            "sheet": _sheet_context is not None,
            "videoDir": bool(videos),
            "participants": [v["id"] for v in videos if v.get("has_video")],
            # Where a run's artifacts land — surfaced in the run panel so the
            # user knows where to find their clips/reels/viewers.
            "outputDir": str(utils.get_effective_output_dir()),
            # Auto-run trigger types for the toolbar picker (no duplicated
            # Python↔JS constants; workflows.TRIGGER_TYPES is the source).
            "triggerTypes": workflows.TRIGGER_TYPES,
        },
    )


# ---- Blueprint CRUD (the canvas autosave target) ----


@workflows_bp.route("/api/blueprints")
def api_blueprints() -> Any:
    """Return the persisted workflow blueprints."""
    with _manifest_lock:
        blueprints = copy.deepcopy(_manifest.get("blueprints", []))
    return ok(blueprints=blueprints)


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
        # Auto-run binding: null when never armed, else {"type": <TRIGGER_TYPES
        # id>, "enabled": bool} — see api_blueprint_trigger.
        "trigger": None,
        "createdAt": datetime.now(UTC).isoformat(),
    }
    with _manifest_lock:
        _manifest.setdefault("blueprints", []).append(blueprint)
        _persist_locked()
    return ok(blueprint=blueprint)


@workflows_bp.route("/api/blueprints/<bp_id>", methods=["PUT"])
def api_blueprints_update(bp_id: str) -> Any:
    """Update a blueprint's name/nodes/edges/viewport (the debounced autosave)."""
    data = request.get_json(silent=True)
    if not data:
        return err("JSON body required")
    with _manifest_lock:
        blueprints = _manifest.get("blueprints", [])
        blueprint = find_by_id(blueprints, bp_id)
        if blueprint is None:
            return err("Blueprint not found", 404)
        for key in ("name", "nodes", "edges", "viewport"):
            if key in data:
                blueprint[key] = data[key]
        _persist_locked()
    return ok(blueprint=blueprint)


@workflows_bp.route("/api/blueprints/<bp_id>", methods=["DELETE"])
def api_blueprints_delete(bp_id: str) -> Any:
    """Delete a blueprint by id."""
    with _manifest_lock:
        if remove_by_id(_manifest.get("blueprints", []), bp_id) is None:
            return err("Blueprint not found", 404)
        _persist_locked()
    return ok()


@workflows_bp.route("/api/blueprints/<bp_id>/trigger", methods=["PUT"])
def api_blueprint_trigger(bp_id: str) -> Any:
    """Arm/disarm an auto-run trigger on a blueprint.

    ``type`` picks the trigger source: ``new_video`` (the watch-dir trigger),
    ``transcript_complete``, or ``scan_event``. A *single* blueprint may be
    armed **per trigger type**: arming one disables the same-type trigger on
    every other, so a "new video → transcribe" pipeline and a "transcript done →
    export" pipeline can chain. Kept on its own endpoint so the debounced canvas
    autosave (the generic ``PUT`` above, which never sends ``trigger``) can't
    clobber the armed state. Arming requires a runnable graph — a passing DAG
    check and at least one ``video_source`` node to bind the firing participant
    onto (uniform across trigger types).
    """
    data = request.get_json(silent=True) or {}
    enabled = bool(data.get("enabled"))
    trigger_type = str(data.get("type") or "new_video")
    if trigger_type not in workflows.TRIGGER_TYPE_IDS:
        return err("Unknown trigger type")
    with _manifest_lock:
        blueprints = _manifest.get("blueprints", [])
        target = find_by_id(blueprints, bp_id)
        if target is None:
            return err("Blueprint not found", 404)
        if enabled:
            try:
                workflows.topo_order(target.get("nodes", []), target.get("edges", []))
            except workflows.WorkflowCycleError as exc:
                return err(str(exc))
            if not workflows.blueprint_participant_nodes(target):
                return err("Add a Video Source node before arming auto-run")
            for b in blueprints:
                if b is target:
                    b["trigger"] = {"type": trigger_type, "enabled": True}
                else:
                    b["trigger"] = _disarmed_trigger(b.get("trigger"), trigger_type)
        else:
            # Disarm whatever is currently bound on this blueprint (the client
            # may not know its type); fall back to the requested type.
            current = target.get("trigger")
            off_type = (
                str(current.get("type"))
                if isinstance(current, dict) and current.get("type")
                else trigger_type
            )
            target["trigger"] = {"type": off_type, "enabled": False}
        _persist_locked()
        result = copy.deepcopy(target)
    # Re-baseline on arm so the current backlog (present videos, already-finished
    # transcripts/scans) never retro-fires. The poll maintains no baselines while
    # nothing is armed, so this re-seed is what upholds that promise.
    if enabled:
        _seed_watch_seen()
    return ok(blueprint=result)


def _disarmed_trigger(trigger: Any, trigger_type: str) -> Any:
    """Force a same-type trigger off; leave any other trigger type untouched."""
    if isinstance(trigger, dict) and trigger.get("type") == trigger_type:
        return {"type": trigger_type, "enabled": False}
    return trigger


# ---- Stash CRUD (save/instantiate sub-graphs) ----
#
# A stash is a reusable sub-graph fragment ({id, name, nodes, edges, createdAt,
# builtin}). The server does CRUD only; the frontend instantiates one onto the
# canvas (id remap + position offset) client-side. ``GET`` prepends the read-only
# built-in recipes ahead of the user's persisted stashes, under the same
# combined-manifest locking the blueprint routes use.


@workflows_bp.route("/api/stashes")
def api_stashes() -> Any:
    """Return the built-in recipes (read-only) followed by the user's stashes."""
    with _manifest_lock:
        user_stashes = copy.deepcopy(_manifest.get("stashes", []))
    return ok(stashes=workflows.BUILTIN_STASHES + user_stashes)


@workflows_bp.route("/api/stashes", methods=["POST"])
def api_stashes_create() -> Any:
    """Save a selected sub-graph as a named stash and return it with its id."""
    data = request.get_json(silent=True) or {}
    nodes = data.get("nodes", [])
    if not nodes:
        return err("No nodes to stash")
    stash = {
        "id": "stash_" + uuid.uuid4().hex[:8],
        "name": (data.get("name") or "Stash").strip() or "Stash",
        "nodes": nodes,
        "edges": data.get("edges", []),
        "builtin": False,  # P4 built-ins are served from code; CRUD guards on this
        "createdAt": datetime.now(UTC).isoformat(),
    }
    with _manifest_lock:
        _manifest.setdefault("stashes", []).append(stash)
        _persist_locked()
    return ok(stash=stash)


@workflows_bp.route("/api/stashes/<stash_id>", methods=["PUT"])
def api_stashes_update(stash_id: str) -> Any:
    """Rename a user stash. Built-in recipes are read-only (403)."""
    if any(s["id"] == stash_id for s in workflows.BUILTIN_STASHES):
        return err("Built-in recipes are read-only", 403)
    data = request.get_json(silent=True)
    if not data:
        return err("JSON body required")
    with _manifest_lock:
        stashes = _manifest.get("stashes", [])
        stash = find_by_id(stashes, stash_id)
        if stash is None:
            return err("Stash not found", 404)
        if "name" in data:
            stash["name"] = (data["name"] or stash["name"]).strip() or stash["name"]
        _persist_locked()
    return ok(stash=stash)


@workflows_bp.route("/api/stashes/<stash_id>", methods=["DELETE"])
def api_stashes_delete(stash_id: str) -> Any:
    """Delete a user stash by id. Built-in recipes are read-only (403)."""
    if any(s["id"] == stash_id for s in workflows.BUILTIN_STASHES):
        return err("Built-in recipes are read-only", 403)
    with _manifest_lock:
        if remove_by_id(_manifest.get("stashes", []), stash_id) is None:
            return err("Stash not found", 404)
        _persist_locked()
    return ok()


# ---- Run lifecycle ----


def _build_node_context(cancel_event: threading.Event) -> workflows.NodeContext:
    """Build the per-run ``NodeContext`` from the active launch context.

    Both directories resolve live, for the same reason the media route does:
    the Start overlay's folder picker moves ``config.INPUT_DIR`` long after
    this blueprint was initialized.
    """
    return workflows.NodeContext(
        input_dir=utils.get_effective_input_dir(),
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


def _run_meta_index() -> dict[str, dict[str, Any]]:
    """Per-run scalar metadata keyed by id, without deep-copying run records.

    The batch-summary and batch-discover paths only read a handful of scalars per
    child run (status + grouping fields), yet fire on every child node transition
    and every discover poll. Reading scalars here avoids ``_merged_runs``'s
    whole-history deep copy on those hot paths.
    """
    meta: dict[str, dict[str, Any]] = {}
    with _manifest_lock:
        for record in _manifest.get("runs", []):
            rid = record.get("id")
            if rid:
                meta[rid] = {
                    "id": rid,
                    "status": record.get("status", workflows.RUN_STATUS_QUEUED),
                    "batchId": record.get("batchId", ""),
                    "participant": record.get("participant", ""),
                    "blueprintId": record.get("blueprintId", ""),
                    "startedAt": record.get("startedAt"),
                }
    with _runs_lock:
        live = list(_runs.items())
    for run_id, runner in live:
        meta[run_id] = {
            "id": run_id,
            "status": runner.status,
            "batchId": runner.batch_id,
            "participant": runner.participant,
            "blueprintId": runner.blueprint_id,
            "startedAt": runner.started_at,
        }
    return meta


def _trim_run_history(runs: list[dict[str, Any]]) -> list[str]:
    """Cap persisted run history in place, evicting whole units oldest-first.

    A *unit* is one loose run or one batch's whole set of child runs (grouped by
    ``batchId``). Evicting by unit keeps a batch's children together — a partial
    batch would make its derived summary lie — and the newest unit plus any live
    batch are always kept, so a single large batch is never split or dropped (the
    cap yields to integrity). Without this, a 50+-participant batch would shed its
    earliest children and 404 on drill-in.

    Returns the run ids that were dropped, so the caller can prune their per-node
    result sidecars in lockstep.
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


def _sse_run_payload(run_id: str) -> str:
    """SSE ``data:`` line carrying the current run snapshot."""
    snap = _run_snapshot(run_id)
    return "data: " + json.dumps({"ok": snap is not None, "run": snap}) + "\n\n"


def _launch_run(
    blueprint: dict[str, Any],
    participant: str = "",
    triggered: bool = False,
    trigger_type: str = "",
    target_node_id: str = "",
    seed_results: dict[str, dict[str, Any]] | None = None,
    seed_note: str = "",
) -> dict[str, Any]:
    """Create + spawn one run on a daemon thread; return the initial snapshot.

    Shared by ``POST /api/runs`` and the watch-dir trigger. The blueprint is
    assumed already validated (``topo_order``) and, for a triggered run, already
    participant-bound by the caller. ``target_node_id`` restricts the run to
    that node and its ancestors. ``seed_results`` (resume) pre-completes nodes
    whose output was reloaded from a prior run's sidecars.
    """
    run_id = "run_" + uuid.uuid4().hex[:8]
    cancel_event = threading.Event()
    ctx = _build_node_context(cancel_event)
    runner = workflows.WorkflowRunner(
        run_id,
        blueprint,
        ctx,
        on_update=lambda: _notify_run_clients(run_id),
        participant=participant,
        triggered=triggered,
        trigger_type=trigger_type,
        target_node_id=target_node_id,
        seed_results=seed_results,
        seed_note=seed_note,
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
            # Evict the terminal runner: its summary is in the manifest and its
            # per-node results are on disk as sidecars, so keeping it would only
            # leak the full in-memory results.
            with _runs_lock:
                _runs.pop(run_id, None)

    threading.Thread(
        target=_run_and_finalize, daemon=True, name=f"workflow-{run_id}"
    ).start()
    return runner.snapshot()


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
        return err("Blueprint not found", 404)
    try:
        workflows.topo_order(blueprint.get("nodes", []), blueprint.get("edges", []))
    except workflows.WorkflowCycleError as exc:
        return err(str(exc))
    # Optional partial run: restrict to this node + its ancestors. Reject an
    # unknown id rather than silently running the whole graph, so a stale
    # selection surfaces as a clear error.
    target = str(data.get("targetNodeId") or "")
    if target and not any(n.get("id") == target for n in blueprint.get("nodes", [])):
        return err("Unknown target node")

    # Optional resume: reload the prior run's completed-node sidecars as seeds and
    # execute only what failed or changed, plus everything downstream. Resumes
    # against the CURRENT blueprint (same semantics as Re-run) — an edited graph
    # just seeds fewer nodes. Sidecars load into memory here, so a concurrent
    # history-trim pruning that run's dir mid-flight is harmless.
    participant = ""
    seed_results: dict[str, dict[str, Any]] | None = None
    seed_note = ""
    resume_from = str(data.get("resumeFromRunId") or "")
    if resume_from:
        prior = _run_snapshot(resume_from)
        if prior is None:
            return err("Run to resume not found", 404)
        if prior.get("status") in (
            workflows.RUN_STATUS_QUEUED,
            workflows.RUN_STATUS_RUNNING,
        ):
            return err("Run is still in progress")
        # A batch child / triggered run keeps its participant binding.
        participant = str(prior.get("participant") or "")
        if participant:
            blueprint = workflows.bind_participant(blueprint, participant)
        results_dir = workflows.run_results_dir(
            utils.get_effective_output_dir(), resume_from
        )

        def _load_sidecar(node_id: str) -> dict[str, Any] | None:
            if node_id != os.path.basename(node_id) or node_id in (".", ".."):
                return None
            try:
                loaded = json.loads(
                    (results_dir / f"{node_id}.json").read_text(encoding="utf-8")
                )
            except (OSError, ValueError):
                return None
            return loaded if isinstance(loaded, dict) else None

        seed_results, _plan_notes = workflows.compute_resume_plan(
            blueprint, prior.get("nodeStates") or {}, _load_sidecar
        )
        seed_note = f"Reused from run {resume_from}"

    return ok(
        run=_launch_run(
            blueprint,
            participant=participant,
            target_node_id=target,
            seed_results=seed_results,
            seed_note=seed_note,
        )
    )


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
    return ok(runs=runs)


@workflows_bp.route("/api/runs/<run_id>")
def api_run_get(run_id: str) -> Any:
    """Polling fallback for one run's live/persisted snapshot."""
    snap = _run_snapshot(run_id)
    if snap is None:
        return err("Run not found", 404)
    return ok(run=snap)


@workflows_bp.route("/api/runs/<run_id>/nodes/<node_id>/result")
def api_run_node_result(run_id: str, node_id: str) -> Any:
    """Serve a node's inspectable result sidecar written by the runner.

    Lazily fetched by the run-history UI on row-expand. Returns the raw stored
    payload (already JSON-sanitized at write time). 404 when no sidecar exists.
    """
    # Both ids are path segments — reject anything that could escape the run dir.
    safe_node = node_id == os.path.basename(node_id) and node_id not in (".", "..")
    safe_run = run_id == os.path.basename(run_id) and run_id not in (".", "..")
    if not (safe_node and safe_run):
        return err("Invalid id", 404)
    path = (
        workflows.run_results_dir(utils.get_effective_output_dir(), run_id)
        / f"{node_id}.json"
    )
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return err("No result for node", 404)
    # Sidecars persist every JSON-safe port for resume; the inspector renders
    # only the inspectable subset (and never the __type__ marker).
    view = workflows.inspectable_sidecar_view(payload)
    if not view:
        return err("No result for node", 404)
    return ok(result=view)


@workflows_bp.route("/api/runs/<run_id>/cancel", methods=["POST"])
def api_run_cancel(run_id: str) -> Any:
    """Signal an in-flight run's cancel event (no-op once finished)."""
    with _runs_lock:
        runner = _runs.get(run_id)
    if runner is None:
        return err("Run not found or finished", 404)
    runner.cancel()
    return ok()


@workflows_bp.route("/api/runs/<run_id>/stream")
def api_run_stream(run_id: str) -> Response:
    """SSE stream of one run's snapshot (mirrors screenspace_server.api_tasks_stream)."""
    return _run_stream(lambda: _sse_run_payload(run_id), key=run_id)


# ---- Batch lifecycle (whole-study fan-out) ----


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
    if any(s == workflows.RUN_STATUS_DEGRADED for s in child_statuses):
        return workflows.RUN_STATUS_DEGRADED
    return workflows.RUN_STATUS_COMPLETED


def _batch_summary(batch_id: str) -> dict[str, Any] | None:
    """JSON-safe batch summary (counts + per-child status), live or historical.

    A live batch reads its identity from ``_batches``; a finished one is rebuilt by
    grouping the persisted runs that carry this ``batchId``. Returns ``None`` if no
    such batch exists in either place.
    """
    all_meta = _run_meta_index()
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
            r for r in all_meta.values() if r.get("batchId") == batch_id and r.get("id")
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
        meta = all_meta.get(run_id)
        status = (meta or {}).get("status", workflows.RUN_STATUS_QUEUED)
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


# Source node types whose result is participant-independent across a batch and
# expensive enough to compute once and seed into every child. ``sheet_selection``
# calls the heavily rate-limited Google Sheets API; ``bind_participant`` never
# rebinds it, so re-running it per participant is N identical API round-trips.
# (``region``/``time_range`` are cheap, local, and not worth the bookkeeping.)
_BATCH_CACHEABLE_TYPES = {"sheet_selection"}


def _precompute_shared_nodes(
    blueprint: dict[str, Any], ctx: workflows.NodeContext
) -> dict[str, dict[str, Any]]:
    """Run a batch's participant-independent source nodes once.

    Returns ``{node_id: result}`` to seed into every child runner, so a batch hits
    the rate-limited Sheets API once instead of once per participant. A node that
    raises is simply omitted — the child re-runs it normally (no behavior change).
    """
    seeded: dict[str, dict[str, Any]] = {}
    for node in blueprint.get("nodes", []):
        if node.get("type") not in _BATCH_CACHEABLE_TYPES or node.get("disabled"):
            continue
        executor = workflows.NODE_TYPES.get(node["type"], {}).get("execute")
        if executor is None:
            continue
        try:
            result = executor(ctx, {}, node.get("params", {}) or {})
        except Exception as exc:  # the child run re-runs this; don't sink the batch
            utils.warning_print(f"workflow batch precompute failed: {exc}")
            continue
        if isinstance(result, dict):
            seeded[str(node["id"])] = result
    return seeded


def _run_batch_child(
    run_id: str,
    participant: str,
    batch_id: str,
    blueprint: dict[str, Any],
    seed_results: dict[str, dict[str, Any]],
    batch_cancel: threading.Event,
) -> None:
    """One batch child, start to finish (register → run → persist → evict).

    Self-contained so the coordinator can run children sequentially or in a
    thread pool: every shared-state touch is already behind ``_runs_lock`` /
    ``_manifest_lock``, per-run sidecar dirs are keyed by run id, and
    ``files.get_unique_filename`` reserves output paths atomically.
    """
    child_cancel = threading.Event()
    if batch_cancel.is_set():
        child_cancel.set()  # short-circuits run() to all-skipped + cancelled
    ctx = _build_node_context(child_cancel)
    child_bp = workflows.bind_participant(blueprint, participant)

    def _on_child_update() -> None:
        _notify_run_clients(run_id)
        _notify_batch_clients(batch_id)

    runner = workflows.WorkflowRunner(
        run_id,
        child_bp,
        ctx,
        on_update=_on_child_update,
        participant=participant,
        batch_id=batch_id,
        # Every child gets its own deep copy: downstream executors mutate
        # seeded values in place (files.prepare_clip adds `times` to sheet
        # records), so one shared dict would cross-contaminate siblings —
        # quasi-benign sequentially, an outright race with workers > 1.
        seed_results=copy.deepcopy(seed_results),
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


def _run_batch(batch_id: str, blueprint: dict[str, Any]) -> None:
    """Coordinator thread: run the blueprint once per participant.

    Sequential by default; ``config.WORKFLOWS_BATCH_WORKERS`` > 1 opts into a
    thread pool over the children (clamped to 4 — heavy nodes serialize on
    Whisper/ffmpeg/OCR anyway, so wide pools mostly add memory pressure).
    Continue-on-error is mandatory — one bad participant must not sink the
    batch. A cancel signals the in-flight children (via their runners) and
    short-circuits every not-yet-started child to ``cancelled``. Each child
    runner is evicted from ``_runs`` once persisted (its summary lives in the
    manifest thereafter).
    """
    with _batches_lock:
        record = _batches.get(batch_id)
    if record is None:
        return
    cancel_event: threading.Event = record["cancel_event"]
    plan = list(zip(record["runIds"], record["participants"]))

    # Compute participant-independent sources (sheet_selection) once and seed them
    # into every child, so an N-participant batch hits the Sheets API once, not N.
    seed_results = _precompute_shared_nodes(
        blueprint, _build_node_context(threading.Event())
    )

    workers = max(1, min(4, int(config.WORKFLOWS_BATCH_WORKERS or 1)))
    if workers == 1 or len(plan) <= 1:
        for run_id, participant in plan:
            _run_batch_child(
                run_id, participant, batch_id, blueprint, seed_results, cancel_event
            )
    else:
        with ThreadPoolExecutor(
            max_workers=min(workers, len(plan)),
            thread_name_prefix=f"workflow-{batch_id}",
        ) as pool:
            futures = [
                pool.submit(
                    _run_batch_child,
                    run_id,
                    participant,
                    batch_id,
                    blueprint,
                    seed_results,
                    cancel_event,
                )
                for run_id, participant in plan
            ]
            for future in futures:
                future.result()  # child bodies swallow their own errors

    with _batches_lock:
        _batches.pop(batch_id, None)
    _notify_batch_clients(batch_id)


@workflows_bp.route("/api/batches", methods=["POST"])
def api_batch_create() -> Any:
    """Fan a blueprint out across participants, one sequential run each."""
    data = request.get_json(silent=True) or {}
    bp_id = data.get("blueprintId")
    with _manifest_lock:
        blueprint = next(
            (b for b in _manifest.get("blueprints", []) if b.get("id") == bp_id), None
        )
        blueprint = copy.deepcopy(blueprint) if blueprint else None
    if blueprint is None:
        return err("Blueprint not found", 404)
    bp_id = str(blueprint.get("id", "") or "")
    if not workflows.blueprint_participant_nodes(blueprint):
        return err("Blueprint has no Video Source to fan out over")
    try:
        workflows.topo_order(blueprint.get("nodes", []), blueprint.get("edges", []))
    except workflows.WorkflowCycleError as exc:
        return err(str(exc))

    available = [
        v["id"] for v in utils.discover_participant_videos() if v.get("has_video")
    ]
    requested = data.get("participants")
    if requested:
        participants = [p for p in requested if p in available]
    else:
        participants = available
    if not participants:
        return err("No participants with video found")

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
            "createdAt": datetime.now(UTC).isoformat(),
        }

    threading.Thread(
        target=_run_batch,
        args=(batch_id, blueprint),
        daemon=True,
        name=f"workflow-batch-{batch_id}",
    ).start()
    return ok(batch=_batch_summary(batch_id))


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
    for run in _run_meta_index().values():
        bid = run.get("batchId")
        if bid and bid not in seen:
            batch_ids.append(bid)
            seen.add(bid)
    summaries = [s for s in (_batch_summary(bid) for bid in batch_ids) if s]
    if bp_filter:
        summaries = [s for s in summaries if s.get("blueprintId") == bp_filter]
    summaries.sort(key=lambda s: s.get("createdAt") or "", reverse=True)
    return ok(batches=summaries)


@workflows_bp.route("/api/batches/<batch_id>")
def api_batch_get(batch_id: str) -> Any:
    """One batch's summary + per-child run snapshots (drill-in)."""
    summary = _batch_summary(batch_id)
    if summary is None:
        return err("Batch not found", 404)
    all_runs = _merged_runs()
    runs = [all_runs[c["runId"]] for c in summary["children"] if c["runId"] in all_runs]
    return ok(batch=summary, runs=runs)


@workflows_bp.route("/api/batches/<batch_id>/cancel", methods=["POST"])
def api_batch_cancel(batch_id: str) -> Any:
    """Cancel a batch: stop launching children and cancel the in-flight one."""
    with _batches_lock:
        record = _batches.get(batch_id)
    if record is None:
        return err("Batch not found or finished", 404)
    record["cancel_event"].set()
    with _runs_lock:
        for runner in _runs.values():
            if getattr(runner, "batch_id", "") == batch_id:
                runner.cancel()
    return ok()


@workflows_bp.route("/api/batches/<batch_id>/stream")
def api_batch_stream(batch_id: str) -> Response:
    """SSE stream of one batch's summary (mirrors :func:`api_run_stream`)."""
    return _batch_stream(lambda: _sse_batch_payload(batch_id), key=batch_id)


# ---- Auto-run trigger watcher ----


def _trigger_enabled(trigger: Any, trigger_type: str) -> bool:
    """True if ``trigger`` is an armed binding of the given type."""
    return (
        isinstance(trigger, dict)
        and trigger.get("type") == trigger_type
        and bool(trigger.get("enabled"))
    )


def _manifest_mtime(filename: str) -> float:
    """The manifest file's mtime in the output dir, or 0.0 when absent."""
    try:
        return os.stat(Path(utils.get_effective_output_dir()) / filename).st_mtime
    except OSError:
        return 0.0


def _transcript_markers() -> dict[str, str]:
    """``{pid: transcribed_at}`` for every transcribed participant.

    mtime-gated: the (potentially large, all-segments) transcripts manifest is
    re-parsed only when its file actually changed — a handful of times per
    session, not once per poll tick.
    """
    global _watch_transcript_cache
    mtime = _manifest_mtime(config.TRANSCRIPTS_MANIFEST_FILENAME)
    if mtime == _watch_transcript_cache[0]:
        return _watch_transcript_cache[1]
    markers: dict[str, str] = {}
    if mtime:
        manifest = (
            utils.load_json_manifest(config.TRANSCRIPTS_MANIFEST_FILENAME, default={})
            or {}
        )
        source = manifest.get("source_transcripts", {}) or {}
        if isinstance(source, dict):
            for pid, entry in source.items():
                if isinstance(entry, dict) and entry.get("transcribed_at"):
                    markers[str(pid)] = str(entry["transcribed_at"])
    _watch_transcript_cache = (mtime, markers)
    return markers


def _scan_markers() -> dict[str, str]:
    """``{task_id: participant}`` for every completed Screenspace task (mtime-gated)."""
    global _watch_scan_cache
    mtime = _manifest_mtime(config.SCREENSPACE_MANIFEST_FILENAME)
    if mtime == _watch_scan_cache[0]:
        return _watch_scan_cache[1]
    markers: dict[str, str] = {}
    if mtime:
        manifest = (
            utils.load_json_manifest(config.SCREENSPACE_MANIFEST_FILENAME, default={})
            or {}
        )
        for task in manifest.get("tasks", []) or []:
            if (
                isinstance(task, dict)
                and task.get("status") == "completed"
                and task.get("id")
            ):
                markers[str(task["id"])] = str(task.get("participant", "") or "")
    _watch_scan_cache = (mtime, markers)
    return markers


def _seed_watch_seen() -> None:
    """Baseline every trigger source so the existing backlog never auto-fires.

    Present videos, already-transcribed participants, and already-completed
    scans are all recorded; the watcher fires only for arrivals/completions
    that happen *after* this call.
    """
    global _watch_transcript_cache, _watch_scan_cache
    with _watch_lock:
        _watch_seen.clear()
        _watch_pending.clear()
        for entry in utils.discover_participant_videos():
            if entry.get("has_video"):
                _watch_seen.add(str(entry["id"]))
        # Force a fresh parse (the cached mtime may predate this call).
        _watch_transcript_cache = (-1.0, {})
        _watch_scan_cache = (-1.0, {})
        _watch_transcript_baseline.clear()
        _watch_transcript_baseline.update(_transcript_markers())
        _watch_scan_seen.clear()
        _watch_scan_seen.update(_scan_markers())


def _stat_first_video(video_paths: list[str]) -> tuple[int, float] | None:
    """``(size, mtime)`` of a participant's first video, or ``None`` if unreadable."""
    if not video_paths:
        return None
    try:
        st = os.stat(video_paths[0])
    except OSError:
        return None  # mid-rename / vanished — treat as not-yet-stable
    return (st.st_size, st.st_mtime)


def _armed_blueprint_locked(trigger_type: str) -> dict[str, Any] | None:
    """A deepcopy of the single armed blueprint of this type, or ``None``.
    Caller holds lock.

    Single-active-per-type is enforced on write (``api_blueprint_trigger``);
    this still defends by returning ``None`` when zero or (defensively) more
    than one are armed, so an inconsistent manifest never fans out unexpectedly.
    """
    armed = [
        b
        for b in _manifest.get("blueprints", [])
        if _trigger_enabled(b.get("trigger"), trigger_type)
    ]
    return copy.deepcopy(armed[0]) if len(armed) == 1 else None


def _maybe_fire_trigger(participant: str, trigger_type: str) -> None:
    """Auto-run this type's armed blueprint for one arrival/completion."""
    with _manifest_lock:
        blueprint = _armed_blueprint_locked(trigger_type)
    if blueprint is None:
        return  # disarmed/ambiguous — the marker is already baselined, won't refire
    bound = workflows.bind_participant(blueprint, participant)
    try:
        workflows.topo_order(bound.get("nodes", []), bound.get("edges", []))
    except workflows.WorkflowCycleError:
        utils.warning_print(
            f"auto-run trigger: armed blueprint has a cycle; skipping {participant}"
        )
        return
    _launch_run(
        bound, participant=participant, triggered=True, trigger_type=trigger_type
    )


def _poll_new_videos() -> None:
    """The original watch-dir tick: fire on newly-arrived, stable participants.

    A pid fires only after it stats identically across two consecutive polls
    (the partial-copy guard).
    """
    entries = {
        str(e["id"]): e
        for e in utils.discover_participant_videos()
        if e.get("has_video")
    }
    fire: list[str] = []
    with _watch_lock:
        current = set(entries)
        # Drop pendings whose file vanished (e.g. a copy that was aborted).
        for pid in [p for p in _watch_pending if p not in current]:
            _watch_pending.pop(pid, None)
        for pid, entry in entries.items():
            if pid in _watch_seen:
                continue
            stat = _stat_first_video(entry.get("video_paths", []))
            if stat is None:
                continue
            if _watch_pending.get(pid) == stat:
                # Stable across two polls -> mark seen + queue a fire.
                _watch_seen.add(pid)
                _watch_pending.pop(pid, None)
                fire.append(pid)
            else:
                _watch_pending[pid] = stat
    for pid in fire:
        _maybe_fire_trigger(pid, "new_video")


def _poll_transcript_completions() -> None:
    """Fire once per (participant, transcribed_at) the baseline hasn't seen.

    A re-transcription bumps ``transcribed_at`` and fires again — deliberate:
    the chained graph should re-process the fresh transcript. Workflow-launched
    Transcribe nodes never write the transcripts manifest, so a triggered graph
    can't re-fire itself.
    """
    markers = _transcript_markers()
    fire: list[str] = []
    with _watch_lock:
        for pid, stamp in markers.items():
            if _watch_transcript_baseline.get(pid) != stamp:
                _watch_transcript_baseline[pid] = stamp
                fire.append(pid)
    for pid in fire:
        _maybe_fire_trigger(pid, "transcript_complete")


def _poll_scan_completions() -> None:
    """Fire once per newly-completed Screenspace task (keyed by task id)."""
    markers = _scan_markers()
    fire: list[str] = []
    with _watch_lock:
        for task_id, pid in markers.items():
            if task_id not in _watch_scan_seen:
                _watch_scan_seen.add(task_id)
                if pid:
                    fire.append(pid)
    for pid in fire:
        _maybe_fire_trigger(pid, "scan_event")


def _watch_poll_once() -> None:
    """One watcher tick: check each trigger type's source, but only while a
    blueprint of that type is armed — with nothing armed the tick does no I/O
    (no glob, no stat, no manifest parse), the common case even in Studio /
    Screenspace / Transcripts launches where this daemon also runs. The
    no-retro-fire guarantee is upheld by re-baselining on arm
    (``api_blueprint_trigger``), not by maintaining baselines here.
    """
    with _manifest_lock:
        armed = {
            t["id"]: _armed_blueprint_locked(t["id"]) is not None
            for t in workflows.TRIGGER_TYPES
        }
    if armed.get("new_video"):
        _poll_new_videos()
    if armed.get("transcript_complete"):
        _poll_transcript_completions()
    if armed.get("scan_event"):
        _poll_scan_completions()


def _watch_loop() -> None:
    """Daemon body: poll until stopped, never dying on a transient error."""
    while not _watch_stop.is_set():
        try:
            _watch_poll_once()
        except Exception as exc:  # a poll error must not kill the daemon
            utils.warning_print(f"watch-dir trigger poll failed: {exc}")
        _watch_stop.wait(config.WORKFLOWS_WATCH_POLL_SECONDS)


def _start_watch_thread() -> None:
    """Start the watch-dir daemon (idempotent; daemon dies with the process)."""
    global _watch_thread
    with _watch_lock:
        if _watch_thread is not None and _watch_thread.is_alive():
            return
        _watch_stop.clear()
        _watch_thread = threading.Thread(
            target=_watch_loop, daemon=True, name="workflow-watch-dir"
        )
        _watch_thread.start()


def _init_workflows_state(
    sheet_context: Any = None,
    worksheet: Any = None,
) -> None:
    """Initialize module-level state for Workflows routes.

    Loads the workflows manifest and records the active sheet context +
    worksheet (the latter feeds the ``sheet_selection`` executor), then seeds
    the watch-dir baseline and starts the trigger daemon. Per-participant
    video paths and the input dir are resolved on demand.

    Called once, from ``build_combined_app``. A worksheet swap goes through
    :func:`repin_sheet_state` instead — re-running this would reload the
    manifest, reseed the watch baseline and restart the trigger daemon on
    every spreadsheet the user opens.
    """
    global _sheet_context, _worksheet, _manifest

    _sheet_context = sheet_context
    _worksheet = worksheet
    _manifest = workflows.load_workflows_manifest()
    # Reclaim a stale empty manifest (e.g. an abandoned auto-created "Untitled"
    # blueprint) left by a prior session: the guarded save removes the file when
    # empty and is an idempotent rewrite otherwise.
    with _manifest_lock:
        _persist_locked()
    _seed_watch_seen()
    _start_watch_thread()


def repin_sheet_state(sheet_context: Any = None, worksheet: Any = None) -> None:
    """Point the blueprint at a newly opened (or closed) worksheet.

    The sheet-only half of :func:`_init_workflows_state`, called by
    ``server._swap_worksheet``. Without it this blueprint kept whatever sheet
    the *process* started with — normally none, since a desktop launch has no
    ``-s`` — so a spreadsheet opened from the Start overlay never reached the
    canvas or any run's ``NodeContext``.
    """
    global _sheet_context, _worksheet

    _sheet_context = sheet_context
    _worksheet = worksheet
