"""Run the combined Flask app on a real loopback port, then take it fully down.

``server.start_combined_server`` is unusable here: it calls ``webbrowser.open``
unconditionally and then blocks forever. ``build_combined_app`` returns the live
``Flask`` instance instead — but it deliberately does less than the launcher, so
three things are this module's job.

1. ``utils.NO_INPUT_MODE`` — only the launcher sets it. Without it, any pipeline
   code path that prompts would block a Flask worker thread forever.
2. ``srv.block_on_close`` — werkzeug's ``ThreadedWSGIServer`` overrides
   ``daemon_threads`` but leaves ``socketserver``'s ``block_on_close = True``, so
   ``server_close()`` joins every tracked connection thread. One still-open SSE
   stream at teardown hangs the run with no output. This is the single nastiest
   failure mode in the harness.
3. The three background workers. ``build_combined_app`` starts a Screenspace
   worker, a Transcripts worker and the Workflows watch-dir thread, and they are
   *module* globals rather than app-scoped — they outlive the app object and keep
   polling whatever ``config.INPUT_DIR`` points at next.
"""

import logging
import threading
from dataclasses import dataclass
from typing import Any

from werkzeug.serving import ThreadedWSGIServer, make_server

import screenspace_server
import server
import transcripts_server
import workflows_server


@dataclass
class LiveServer:
    url: str
    srv: Any
    thread: threading.Thread


def start(worksheet: Any) -> LiveServer:
    """Build the combined app and serve it on an ephemeral loopback port."""
    # A page load pulls a few hundred assets, and werkzeug logs a line per
    # request. Under pytest that is merely captured; under shot.py it buries the
    # screenshot path and the --eval result the caller actually came for.
    logging.getLogger("werkzeug").setLevel(logging.ERROR)
    app = server.build_combined_app(worksheet=worksheet, default_page="studio")
    srv = make_server("127.0.0.1", 0, app, threaded=True)
    # `threaded=True` guarantees this, but make_server's return type is the base
    # class, and block_on_close only exists on the ThreadingMixIn subclass.
    assert isinstance(srv, ThreadedWSGIServer)
    srv.block_on_close = False  # see module docstring
    thread = threading.Thread(
        target=srv.serve_forever, daemon=True, name="ui-check-server"
    )
    thread.start()
    return LiveServer(f"http://127.0.0.1:{srv.server_port}", srv, thread)


def stop(live: LiveServer) -> None:
    """Shut the socket down, then stop every thread ``build_combined_app`` started."""
    live.srv.shutdown()
    live.srv.server_close()
    live.thread.join(timeout=10)

    if screenspace_server._worker is not None:
        screenspace_server._worker.stop()
        screenspace_server._worker = None
    if transcripts_server._worker is not None:
        transcripts_server._worker.stop()
        transcripts_server._worker = None
    # `_watch_stop` is annotated in-source as "tests only; production never sets
    # it" — this is exactly its intended use.
    workflows_server._watch_stop.set()
    if workflows_server._watch_thread is not None:
        workflows_server._watch_thread.join(timeout=5)
        workflows_server._watch_thread = None
    workflows_server._watch_stop.clear()
