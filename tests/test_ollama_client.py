"""Tests for ollama_client transport layer.

Agent-specific behavior (summarization, citation linking) lives in
tests/test_thinking_agents.py.
"""

import http.server
import io
import json
import socketserver
import threading
import time
import urllib.error
from contextlib import contextmanager
from unittest.mock import MagicMock, patch

import ollama_client


def _ndjson_lines(*chunks: dict) -> list[bytes]:
    """Encode each chunk dict as a separate NDJSON line (bytes)."""
    return [(json.dumps(c) + "\n").encode("utf-8") for c in chunks]


def _make_streaming_resp(lines: list[bytes]) -> MagicMock:
    """Build a mock response object whose readline() yields *lines* then b""."""
    queue = list(lines) + [b""]
    closed = {"flag": False}

    def _readline() -> bytes:
        if closed["flag"]:
            return b""
        if not queue:
            return b""
        return queue.pop(0)

    def _close() -> None:
        closed["flag"] = True

    resp = MagicMock()
    resp.readline.side_effect = _readline
    resp.close.side_effect = _close
    return resp


class TestIsAvailable:
    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_true_when_server_responds(self, mock_urlopen):
        mock_urlopen.return_value.__enter__ = lambda s: s
        mock_urlopen.return_value.__exit__ = lambda s, *a: None
        assert ollama_client.is_available() is True

    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_false_on_connection_error(self, mock_urlopen):
        mock_urlopen.side_effect = urllib.error.URLError("Connection refused")
        assert ollama_client.is_available() is False

    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_false_on_timeout(self, mock_urlopen):
        mock_urlopen.side_effect = OSError("timed out")
        assert ollama_client.is_available() is False


class TestListModels:
    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_models_on_success(self, mock_urlopen):
        response_data = json.dumps(
            {
                "models": [
                    {
                        "name": "qwen3.5:0.8b",
                        "size": 531490688,
                        "details": {
                            "parameter_size": "0.8B",
                            "quantization_level": "Q4_K_M",
                            "family": "qwen3.5",
                        },
                    },
                    {
                        "name": "gemma3:4b",
                        "size": 2100000000,
                        "details": {
                            "parameter_size": "4B",
                            "quantization_level": "Q4_0",
                            "family": "gemma3",
                        },
                    },
                ]
            }
        ).encode("utf-8")
        mock_resp = io.BytesIO(response_data)
        mock_urlopen.return_value.__enter__ = lambda s: mock_resp
        mock_urlopen.return_value.__exit__ = lambda s, *a: None
        result = ollama_client.list_models()
        assert result is not None
        assert len(result) == 2
        assert result[0]["name"] == "qwen3.5:0.8b"
        assert result[0]["size_bytes"] == 531490688
        assert result[0]["parameter_size"] == "0.8B"
        assert result[0]["family"] == "qwen3.5"
        assert result[1]["name"] == "gemma3:4b"

    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_none_on_connection_error(self, mock_urlopen):
        mock_urlopen.side_effect = urllib.error.URLError("Connection refused")
        assert ollama_client.list_models() is None

    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_empty_list_for_no_models(self, mock_urlopen):
        response_data = json.dumps({"models": []}).encode("utf-8")
        mock_resp = io.BytesIO(response_data)
        mock_urlopen.return_value.__enter__ = lambda s: mock_resp
        mock_urlopen.return_value.__exit__ = lambda s, *a: None
        result = ollama_client.list_models()
        assert result is not None
        assert len(result) == 0


class TestUnloadModel:
    @patch("ollama_client.urllib.request.urlopen")
    def test_sends_keep_alive_zero(self, mock_urlopen):
        mock_resp = io.BytesIO(b"{}")
        mock_urlopen.return_value.__enter__ = lambda s: mock_resp
        mock_urlopen.return_value.__exit__ = lambda s, *a: None

        assert ollama_client.unload_model("qwen3.5:9b") is True

        call_args = mock_urlopen.call_args
        req = call_args[0][0]
        body = json.loads(req.data.decode("utf-8"))
        assert body["model"] == "qwen3.5:9b"
        assert body["keep_alive"] == 0
        assert body["stream"] is False

    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_false_on_connection_error(self, mock_urlopen):
        mock_urlopen.side_effect = urllib.error.URLError("Connection refused")
        assert ollama_client.unload_model("qwen3.5:9b") is False


class TestGenerate:
    @patch("ollama_client.urllib.request.urlopen")
    def test_streams_and_concatenates_chunks(self, mock_urlopen):
        lines = _ndjson_lines(
            {"response": "Hello", "done": False},
            {"response": " world", "done": False},
            {"response": "", "done": True},
        )
        mock_urlopen.return_value = _make_streaming_resp(lines)
        result = ollama_client.generate("test prompt")
        assert result == "Hello world"

    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_think_tags_verbatim(self, mock_urlopen):
        # Transport is pure: <think> stripping is a response-parsing concern that
        # lives in thinking_agents, so generate() must pass reasoning blocks through
        # untouched (only trims surrounding whitespace).
        lines = _ndjson_lines(
            {"response": "<think>Let me think...</think>", "done": False},
            {"response": "\n\nHere is the summary.", "done": True},
        )
        mock_urlopen.return_value = _make_streaming_resp(lines)
        result = ollama_client.generate("test prompt")
        assert result == "<think>Let me think...</think>\n\nHere is the summary."

    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_none_on_empty_stream(self, mock_urlopen):
        lines = _ndjson_lines({"response": "", "done": True})
        mock_urlopen.return_value = _make_streaming_resp(lines)
        result = ollama_client.generate("test prompt")
        assert result is None

    @patch("ollama_client.urllib.request.urlopen")
    def test_skips_malformed_json_lines(self, mock_urlopen):
        valid_lines = _ndjson_lines(
            {"response": "Hello", "done": False},
            {"response": " world", "done": True},
        )
        # Inject a malformed line between the two valid ones
        lines = [valid_lines[0], b"not json\n", valid_lines[1]]
        mock_urlopen.return_value = _make_streaming_resp(lines)
        result = ollama_client.generate("test prompt")
        assert result == "Hello world"

    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_none_on_connection_error(self, mock_urlopen):
        mock_urlopen.side_effect = urllib.error.URLError("Connection refused")
        result = ollama_client.generate("test prompt")
        assert result is None

    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_none_on_http_error(self, mock_urlopen):
        from email.message import Message

        hdrs = Message()
        mock_urlopen.side_effect = urllib.error.HTTPError(
            url="",
            code=404,
            msg="Not Found",
            hdrs=hdrs,
            fp=None,
        )
        result = ollama_client.generate("test prompt")
        assert result is None

    @patch("ollama_client.urllib.request.urlopen")
    def test_uses_config_model_by_default(self, mock_urlopen):
        lines = _ndjson_lines({"response": "ok", "done": True})
        mock_urlopen.return_value = _make_streaming_resp(lines)

        ollama_client.generate("test prompt")

        call_args = mock_urlopen.call_args
        req = call_args[0][0]
        body = json.loads(req.data.decode("utf-8"))
        assert body["model"] == "qwen3.5:9b"

    @patch("ollama_client.urllib.request.urlopen")
    def test_uses_custom_model(self, mock_urlopen):
        lines = _ndjson_lines({"response": "ok", "done": True})
        mock_urlopen.return_value = _make_streaming_resp(lines)

        ollama_client.generate("test prompt", model="llama3.1:8b")

        call_args = mock_urlopen.call_args
        req = call_args[0][0]
        body = json.loads(req.data.decode("utf-8"))
        assert body["model"] == "llama3.1:8b"

    @patch("ollama_client.urllib.request.urlopen")
    def test_includes_system_prompt(self, mock_urlopen):
        lines = _ndjson_lines({"response": "ok", "done": True})
        mock_urlopen.return_value = _make_streaming_resp(lines)

        ollama_client.generate("test prompt", system="You are helpful.")

        call_args = mock_urlopen.call_args
        req = call_args[0][0]
        body = json.loads(req.data.decode("utf-8"))
        assert body["system"] == "You are helpful."

    @patch("ollama_client.urllib.request.urlopen")
    def test_request_body_uses_stream_true(self, mock_urlopen):
        lines = _ndjson_lines({"response": "ok", "done": True})
        mock_urlopen.return_value = _make_streaming_resp(lines)

        ollama_client.generate("test prompt")

        call_args = mock_urlopen.call_args
        req = call_args[0][0]
        body = json.loads(req.data.decode("utf-8"))
        assert body["stream"] is True

    @patch("ollama_client.urllib.request.urlopen")
    def test_cancel_event_set_before_call_returns_none(self, mock_urlopen):
        lines = _ndjson_lines({"response": "should not appear", "done": True})
        resp = _make_streaming_resp(lines)
        mock_urlopen.return_value = resp
        evt = threading.Event()
        evt.set()
        result = ollama_client.generate("test prompt", cancel_event=evt)
        assert result is None
        # The watcher thread closed the response on event.set(), and the
        # finally block also calls close(); either way it must be closed.
        assert resp.close.call_count >= 1

    @patch("ollama_client.urllib.request.urlopen")
    def test_cancel_event_during_stream_aborts(self, mock_urlopen):
        # readline blocks on a synchronisation event so the test thread can
        # trigger cancellation while the worker is waiting for data.
        block_event = threading.Event()

        def _readline_blocking() -> bytes:
            # Block until the test releases us — simulating a slow model.
            block_event.wait(timeout=2.0)
            return b""

        resp = MagicMock()
        resp.readline.side_effect = _readline_blocking
        # When close() is called, unblock the readline so the worker exits.
        resp.close.side_effect = lambda: block_event.set()
        mock_urlopen.return_value = resp

        evt = threading.Event()
        # Trigger cancel from a side thread so we exercise the watcher path.
        timer = threading.Timer(0.05, evt.set)
        timer.start()
        try:
            result = ollama_client.generate("test prompt", cancel_event=evt)
        finally:
            timer.cancel()
            block_event.set()  # safety net
        assert result is None
        assert resp.close.called

    @patch("ollama_client.urllib.request.urlopen")
    def test_no_cancel_event_works_normally(self, mock_urlopen):
        # When no cancel_event is passed, no watcher thread spawns and the
        # request completes normally.
        lines = _ndjson_lines({"response": "fine", "done": True})
        mock_urlopen.return_value = _make_streaming_resp(lines)
        result = ollama_client.generate("test prompt")
        assert result == "fine"

    @patch("ollama_client.urllib.request.urlopen")
    def test_on_token_fires_per_chunk_and_returns_joined(self, mock_urlopen):
        # on_token sees each streamed piece live, in order; the return value is
        # still the joined text (the done chunk's empty response is not emitted).
        lines = _ndjson_lines(
            {"response": "Hello", "done": False},
            {"response": " world", "done": False},
            {"response": "", "done": True},
        )
        mock_urlopen.return_value = _make_streaming_resp(lines)
        seen: list[str] = []
        result = ollama_client.generate("test prompt", on_token=seen.append)
        assert result == "Hello world"
        assert seen == ["Hello", " world"]

    @patch("ollama_client.urllib.request.urlopen")
    def test_on_token_error_does_not_break_stream(self, mock_urlopen):
        # A raising callback is swallowed — the accumulated result is unaffected.
        lines = _ndjson_lines(
            {"response": "Hello", "done": False},
            {"response": " world", "done": True},
        )
        mock_urlopen.return_value = _make_streaming_resp(lines)

        def _boom(_piece: str) -> None:
            raise RuntimeError("callback blew up")

        result = ollama_client.generate("test prompt", on_token=_boom)
        assert result == "Hello world"


class _StubOllamaHandler(http.server.BaseHTTPRequestHandler):
    """Test HTTP handler that mimics Ollama /api/generate streaming.

    Behavior is selected by the request body's ``prompt`` field:
      - "complete" → send 3 NDJSON chunks, finish quickly.
      - "block"    → send 1 chunk, then sleep so the connection blocks
                     readline() until shut down.
    """

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        body_bytes = self.rfile.read(length)
        body = json.loads(body_bytes.decode("utf-8"))
        prompt = body.get("prompt", "")
        self.send_response(200)
        self.send_header("Content-Type", "application/x-ndjson")
        self.send_header("Transfer-Encoding", "chunked")
        self.end_headers()

        def _write_chunk(piece: str) -> None:
            data = (piece + "\n").encode("utf-8")
            self.wfile.write(b"%x\r\n" % len(data))
            self.wfile.write(data)
            self.wfile.write(b"\r\n")
            self.wfile.flush()

        try:
            if prompt == "complete":
                _write_chunk('{"response":"Hello","done":false}')
                _write_chunk('{"response":" world","done":false}')
                _write_chunk('{"response":"","done":true}')
                self.wfile.write(b"0\r\n\r\n")
            elif prompt == "block":
                _write_chunk('{"response":"Hello","done":false}')
                # Block; the test will close the socket via cancel.
                time.sleep(30)
        except (BrokenPipeError, ConnectionResetError, OSError):
            return

    def log_message(self, *_args, **_kwargs) -> None:
        return


@contextmanager
def _stub_ollama_server():
    srv = socketserver.ThreadingTCPServer(("127.0.0.1", 0), _StubOllamaHandler)
    srv.daemon_threads = True
    port = srv.server_address[1]
    thread = threading.Thread(target=srv.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{port}"
    finally:
        srv.shutdown()
        srv.server_close()


class TestGenerateAgainstRealServer:
    """End-to-end tests against a real HTTP server.

    Mocking urlopen is not sufficient to exercise the cancel path because the
    abort relies on shutting down the underlying socket so a blocked
    readline() returns. These tests run against a stub HTTP server that
    blocks until the cancel actually closes the connection.
    """

    def test_streaming_completes_and_returns_text(self, monkeypatch):
        with _stub_ollama_server() as url:
            monkeypatch.setattr(ollama_client.config, "OLLAMA_BASE_URL", url)
            result = ollama_client.generate("complete")
            assert result == "Hello world"

    def test_cancel_aborts_blocked_request_quickly(self, monkeypatch):
        with _stub_ollama_server() as url:
            monkeypatch.setattr(ollama_client.config, "OLLAMA_BASE_URL", url)
            evt = threading.Event()
            # Trigger cancel shortly after the request starts so the watcher
            # has to shut down the underlying socket to unblock readline().
            timer = threading.Timer(0.2, evt.set)
            timer.start()
            start = time.monotonic()
            try:
                result = ollama_client.generate("block", cancel_event=evt)
            finally:
                timer.cancel()
            elapsed = time.monotonic() - start
            assert result is None
            # Real abort should return well under the 30s server sleep and
            # under the urlopen connect timeout. Use a generous bound to
            # avoid flake on slow CI.
            assert elapsed < 5.0, f"Cancel took {elapsed:.2f}s — abort is broken"


class TestAutoStartServer:
    """Tests for the auto-start behavior when Ollama is not running."""

    @patch("ollama_client._start_server", return_value=True)
    @patch("ollama_client.urllib.request.urlopen")
    def test_retries_after_connection_refused(self, mock_urlopen, mock_start):
        """On ConnectionRefusedError, start server and retry successfully."""
        lines = _ndjson_lines({"response": "retried ok", "done": True})

        # First call: connection refused; second call: streaming success
        mock_urlopen.side_effect = [
            urllib.error.URLError(ConnectionRefusedError("Connection refused")),
            _make_streaming_resp(lines),
        ]
        result = ollama_client.generate("test prompt")
        assert result == "retried ok"
        mock_start.assert_called_once()

    @patch("ollama_client._start_server", return_value=False)
    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_none_when_server_start_fails(self, mock_urlopen, mock_start):
        """If server fails to start, return None without retrying."""
        mock_urlopen.side_effect = urllib.error.URLError(
            ConnectionRefusedError("Connection refused")
        )
        result = ollama_client.generate("test prompt")
        assert result is None
        mock_start.assert_called_once()

    @patch("ollama_client._start_server")
    @patch("ollama_client.urllib.request.urlopen")
    def test_does_not_start_server_on_other_errors(self, mock_urlopen, mock_start):
        """Non-ConnectionRefused errors should not trigger auto-start."""
        mock_urlopen.side_effect = urllib.error.URLError("DNS resolution failed")
        result = ollama_client.generate("test prompt")
        assert result is None
        mock_start.assert_not_called()

    @patch("ollama_client.shutil.which", return_value="/usr/local/bin/ollama")
    @patch("ollama_client.subprocess.Popen")
    @patch("ollama_client.is_available")
    def test_start_server_polls_until_available(
        self, mock_avail, mock_popen, mock_which
    ):
        """_start_server polls is_available until it returns True."""
        mock_avail.side_effect = [False, False, True]
        assert ollama_client._start_server() is True
        assert mock_avail.call_count == 3
        mock_popen.assert_called_once()

    @patch("ollama_client.shutil.which", return_value=None)
    def test_start_server_returns_false_when_binary_missing(self, mock_which):
        """_start_server returns False when ollama is not installed."""
        assert ollama_client._start_server() is False

    @patch("ollama_client.shutil.which", return_value=None)
    @patch("ollama_client.utils.warning_print")
    def test_start_server_shows_install_guidance(self, mock_warn, mock_which):
        """_start_server shows install guidance when ollama is not in PATH."""
        ollama_client._start_server()
        mock_warn.assert_called_once()
        assert "not installed" in mock_warn.call_args[0][0].lower()
        details = mock_warn.call_args[1]["details"]
        assert any("ollama --version" in line for line in details)


class TestIsModelInstalled:
    @patch("ollama_client.list_models")
    def test_exact_match(self, mock_list):
        mock_list.return_value = [{"name": "qwen3.5:9b"}, {"name": "llama3:latest"}]
        assert ollama_client.is_model_installed("qwen3.5:9b") is True

    @patch("ollama_client.list_models")
    def test_implicit_latest_tag(self, mock_list):
        mock_list.return_value = [{"name": "llama3:latest"}]
        assert ollama_client.is_model_installed("llama3") is True

    @patch("ollama_client.list_models")
    def test_matches_other_tag_when_untagged(self, mock_list):
        mock_list.return_value = [{"name": "qwen3.5:9b"}]
        assert ollama_client.is_model_installed("qwen3.5") is True

    @patch("ollama_client.list_models")
    def test_missing_model(self, mock_list):
        mock_list.return_value = [{"name": "llama3:latest"}]
        assert ollama_client.is_model_installed("qwen3.5:9b") is False

    @patch("ollama_client.list_models")
    def test_tagged_request_does_not_prefix_match(self, mock_list):
        # An explicit tag must match exactly — no base-prefix fallback.
        mock_list.return_value = [{"name": "qwen3.5:9b"}]
        assert ollama_client.is_model_installed("qwen3.5:32b") is False

    @patch("ollama_client.list_models")
    def test_server_unreachable(self, mock_list):
        mock_list.return_value = None
        assert ollama_client.is_model_installed("qwen3.5:9b") is False

    def test_empty_model_name(self):
        assert ollama_client.is_model_installed("") is False

    @patch("ollama_client.list_models")
    def test_reuses_supplied_installed_list(self, mock_list):
        # When the caller passes the installed list, no /api/tags call is made.
        installed = [{"name": "qwen3.5:9b"}]
        assert ollama_client.is_model_installed("qwen3.5:9b", installed) is True
        mock_list.assert_not_called()


class TestPullModel:
    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_true_on_success_status(self, mock_urlopen):
        lines = _ndjson_lines(
            {"status": "pulling manifest"},
            {"status": "downloading", "total": 100, "completed": 50},
            {"status": "success"},
        )
        mock_urlopen.return_value = _make_streaming_resp(lines)
        assert ollama_client.pull_model("qwen3.5:9b") is True

    @patch("ollama_client.urllib.request.urlopen")
    def test_reports_progress(self, mock_urlopen):
        lines = _ndjson_lines(
            {"status": "downloading", "total": 100, "completed": 25},
            {"status": "success"},
        )
        mock_urlopen.return_value = _make_streaming_resp(lines)
        seen: list[dict] = []
        ollama_client.pull_model("m", on_progress=seen.append)
        assert any(c.get("completed") == 25 for c in seen)

    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_false_on_error_line(self, mock_urlopen):
        lines = _ndjson_lines({"error": "model 'nope' not found"})
        mock_urlopen.return_value = _make_streaming_resp(lines)
        assert ollama_client.pull_model("nope") is False

    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_false_without_success(self, mock_urlopen):
        lines = _ndjson_lines({"status": "pulling manifest"})
        mock_urlopen.return_value = _make_streaming_resp(lines)
        assert ollama_client.pull_model("m") is False

    def test_empty_model_name(self):
        assert ollama_client.pull_model("") is False
