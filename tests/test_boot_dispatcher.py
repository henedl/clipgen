"""Boot dispatcher: the WSGI shim that serves a boot page until the app is built.

The dispatcher is what lets the server socket bind in well under a second while
the heavy build (av/cv2 preload + blueprint imports) runs in the background.
These tests exercise it as a plain WSGI callable — no sockets, no Flask.
"""

import json
import threading
from typing import Any

import pytest

import server
import utils


def _boot_state(**overrides: Any) -> dict[str, Any]:
    state: dict[str, Any] = {
        "ready": False,
        "phase": "starting",
        "message": "Starting clipgen…",
        "error": None,
        "app": None,
    }
    state.update(overrides)
    return state


def _call(app: Any, path: str) -> tuple[str, dict[str, str], bytes]:
    captured: dict[str, Any] = {}

    def start_response(status: str, headers: list) -> None:
        captured["status"] = status
        captured["headers"] = dict(headers)

    environ = {"PATH_INFO": path, "REQUEST_METHOD": "GET"}
    body = b"".join(app(environ, start_response))
    return captured["status"], captured["headers"], body


def _fake_app(environ: Any, start_response: Any) -> list:
    start_response("200 OK", [("Content-Type", "text/plain")])
    return [b"real-app"]


def test_boot_page_served_while_building():
    dispatcher = server._make_boot_dispatcher(_boot_state())
    status, headers, body = _call(dispatcher, "/studio/")
    assert status == "200 OK"
    assert "text/html" in headers["Content-Type"]
    assert headers["Cache-Control"] == "no-store"
    assert b"api/boot-status" in body  # the poll loop shipped with the page


def test_api_gets_503_envelope_while_building():
    dispatcher = server._make_boot_dispatcher(_boot_state())
    for path in ("/api/status", "/transcripts/api/participants"):
        status, _headers, body = _call(dispatcher, path)
        assert status.startswith("503")
        payload = json.loads(body)
        assert payload["ok"] is False
        assert payload["error"]


def test_boot_status_reflects_phase_and_survives_swap():
    state = _boot_state(
        phase="vision_libs", message="Loading computer-vision libraries…"
    )
    dispatcher = server._make_boot_dispatcher(state)

    _, _, body = _call(dispatcher, "/api/boot-status")
    payload = json.loads(body)
    assert payload == {
        "ready": False,
        "phase": "vision_libs",
        "message": "Loading computer-vision libraries…",
        "error": None,
    }

    # After the swap the dispatcher still owns the route — a boot-page poll
    # landing just after the swap must see ready, not the real app's 404.
    state["app"] = _fake_app
    state["ready"] = True
    state["phase"] = "ready"
    _, _, body = _call(dispatcher, "/api/boot-status")
    assert json.loads(body)["ready"] is True


def test_requests_delegate_to_installed_app():
    state = _boot_state(app=_fake_app, ready=True, phase="ready")
    dispatcher = server._make_boot_dispatcher(state)
    for path in ("/studio/", "/api/status"):
        status, _, body = _call(dispatcher, path)
        assert status == "200 OK"
        assert body == b"real-app"


def test_build_error_lands_on_boot_page():
    state = _boot_state(error="SystemExit: 1")
    dispatcher = server._make_boot_dispatcher(state)
    _, _, body = _call(dispatcher, "/api/boot-status")
    assert json.loads(body)["error"] == "SystemExit: 1"


def test_boot_page_splices_desktop_chrome(monkeypatch):
    monkeypatch.setattr(utils, "DESKTOP_CHROME", "dark")
    dispatcher = server._make_boot_dispatcher(_boot_state())
    _, _, body = _call(dispatcher, "/studio/")
    assert b"desktopChrome" in body
    assert b"CLIPGEN_BOOT_CHROME" not in body  # marker consumed by the splice


def test_serve_combined_app_raises_on_build_failure(monkeypatch):
    """block_until_ready surfaces a build-thread death (incl. SystemExit — a
    bad worksheet exits, and swallowing that in a daemon thread would leave the
    boot page spinning forever) as a RuntimeError with the cause."""

    def _exit_build(**kwargs: Any) -> Any:
        raise SystemExit(1)

    monkeypatch.setattr(utils, "preload_vision_libs_quietly", lambda **kwargs: None)
    monkeypatch.setattr(utils, "sweep_stale_temp_artifacts", lambda: None)
    monkeypatch.setattr(server, "build_combined_app", _exit_build)

    with pytest.raises(RuntimeError, match="SystemExit"):
        server.serve_combined_app(port=0, block_until_ready=True)


def test_serve_combined_app_returns_before_build_completes(monkeypatch):
    """Without block_until_ready the socket is live while the build is still
    running — the whole point of the two-phase boot."""
    release = threading.Event()

    def _slow_build(**kwargs: Any) -> Any:
        release.wait(timeout=10)
        return _fake_app

    monkeypatch.setattr(utils, "preload_vision_libs_quietly", lambda **kwargs: None)
    monkeypatch.setattr(utils, "sweep_stale_temp_artifacts", lambda: None)
    monkeypatch.setattr(server, "build_combined_app", _slow_build)

    live = server.serve_combined_app(port=0)
    try:
        assert live.boot["ready"] is False
        assert live.thread.is_alive()
        release.set()
        assert live.ready.wait(timeout=10)
        assert live.boot["ready"] is True
        assert live.boot["app"] is _fake_app
    finally:
        release.set()
        live.srv.shutdown()
        live.srv.server_close()
