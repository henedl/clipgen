"""Run the combined Flask app on a real loopback port, then take it fully down.

This is a thin wrapper over ``server.serve_combined_app`` /
``server.stop_combined_app`` — the same pair the desktop shell uses — so the
smoke harness exercises the shipped code path rather than a parallel copy of it.
Those functions own the three hazards this module used to handle itself:
``utils.NO_INPUT_MODE``, werkzeug's ``block_on_close`` (one open SSE stream at
teardown would otherwise hang the run with no output), and the three background
workers that ``build_combined_app`` starts as *module* globals.

Only the harness-specific bits stay here: an ephemeral port, and silencing
werkzeug's per-request logging.
"""

import logging
from typing import Any

import server
from server import LiveServer

__all__ = ["LiveServer", "start", "stop"]


def start(worksheet: Any) -> LiveServer:
    """Build the combined app and serve it on an ephemeral loopback port."""
    # A page load pulls a few hundred assets, and werkzeug logs a line per
    # request. Under pytest that is merely captured; under shot.py it buries the
    # screenshot path and the --eval result the caller actually came for.
    logging.getLogger("werkzeug").setLevel(logging.ERROR)
    return server.serve_combined_app(worksheet=worksheet, port=0, default_page="studio")


def stop(live: LiveServer) -> None:
    """Shut the socket down, then stop every thread ``build_combined_app`` started."""
    server.stop_combined_app(live)
