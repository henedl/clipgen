"""Workflows Flask blueprint — serves the node-canvas page and its REST API.

Registered at ``/workflows`` by ``server.build_combined_app`` (mutually exclusive
launch with the other web modes, but all blueprints are always mounted). M0
ships the static page routes, module-state init, and a read-only blueprint-list
endpoint; canvas CRUD, the ``/api/catalog`` node registry, and the run lifecycle
(SSE + polling) land in later milestones — see ``plans/WORKFLOWS-PLAN.md``.

Module-level state (``_input_dir``, ``_sheet_context``, ``_manifest``) is
initialized by :func:`_init_workflows_state`, mirroring the Screenspace and
Transcripts blueprints.
"""

from __future__ import annotations

import threading
from typing import Any

from flask import Blueprint, jsonify

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


@workflows_bp.route("/api/blueprints")
def api_blueprints() -> Any:
    """Return the persisted workflow blueprints (read-only in M0)."""
    with _manifest_lock:
        blueprints = list(_manifest.get("blueprints", []))
    return jsonify({"ok": True, "blueprints": blueprints})


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
