"""Workflows Flask blueprint — serves the node-canvas page and its REST API.

Registered at ``/workflows`` by ``server.build_combined_app`` (mutually exclusive
launch with the other web modes, but all blueprints are always mounted). M1
ships the static page routes, module-state init, the ``/api/catalog`` node
registry, and full blueprint CRUD (the canvas autosave target). The run
lifecycle (SSE + polling) lands in M4 — see ``plans/WORKFLOWS-PLAN.md``.

Module-level state (``_input_dir``, ``_sheet_context``, ``_manifest``) is
initialized by :func:`_init_workflows_state`, mirroring the Screenspace and
Transcripts blueprints. Mutations hold ``_manifest_lock`` and persist via
:func:`_persist_locked` (mirrors ``screenspace_server._do_persist``).
"""

from __future__ import annotations

import copy
import threading
import uuid
from datetime import datetime, timezone
from typing import Any

from flask import Blueprint, jsonify, request

import utils
import workflows

# ---- Module state (initialized by _init_workflows_state) ----

_input_dir: str = ""
_sheet_context: Any = None
_manifest: dict[str, Any] = {}
_manifest_lock = threading.Lock()


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
    ``requires`` are satisfiable in this launch, so it can disable the rest.
    """
    return jsonify(
        {
            "ok": True,
            "catalog": workflows.serialize_catalog(),
            "context": {
                "sheet": _sheet_context is not None,
                "videoDir": bool(utils.discover_participant_videos()),
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


def _init_workflows_state(
    sheet_context: Any = None,
    participant_list: list[str] | None = None,
) -> None:
    """Initialize module-level state for Workflows routes.

    Loads the workflows manifest and records the active input dir + sheet
    context. ``participant_list`` is accepted for parity with the other
    blueprints' init signatures (used by later milestones); M0 does not yet
    resolve per-participant video paths.
    """
    global _input_dir, _sheet_context, _manifest  # noqa: PLW0603

    _input_dir = str(utils.get_effective_input_dir())
    _sheet_context = sheet_context
    _manifest = workflows.load_workflows_manifest()
