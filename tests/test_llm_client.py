"""Tests for the llm_client transport layer (llama-server router mode).

Agent-specific behavior (summarization, citation linking) lives in
tests/test_thinking_agents.py.
"""

import email.message
import http.client
import http.server
import json
import socketserver
import threading
import time
import urllib.error
from contextlib import contextmanager
from unittest.mock import MagicMock, patch

import pytest

import llm_client


@pytest.fixture(autouse=True)
def _isolated_models_dir(tmp_path, monkeypatch):
    """Point the models dir into tmp so tests never see real downloads."""
    monkeypatch.setattr(
        llm_client.start_settings, "config_dir", lambda: tmp_path / "cfg"
    )


def _chunk(piece: str, finish: str | None = None) -> dict:
    return {"choices": [{"delta": {"content": piece}, "finish_reason": finish}]}


def _sse_lines(*chunks: dict, done: bool = True) -> list[bytes]:
    """Encode each chunk dict as an SSE data line, plus the [DONE] marker."""
    lines = [b"data: " + json.dumps(c).encode("utf-8") + b"\n" for c in chunks]
    if done:
        lines.append(b"data: [DONE]\n")
    return lines


def _make_streaming_resp(lines: list[bytes]) -> MagicMock:
    """Build a mock response object whose readline() yields *lines* then b""."""
    queue = list(lines) + [b""]
    closed = {"flag": False}

    def _readline() -> bytes:
        if closed["flag"] or not queue:
            return b""
        return queue.pop(0)

    def _close() -> None:
        closed["flag"] = True

    resp = MagicMock()
    resp.readline.side_effect = _readline
    resp.close.side_effect = _close
    return resp


class TestModelName:
    def test_hf_ref_maps_to_deterministic_stem(self):
        assert (
            llm_client.model_name("unsloth/Qwen3.5-9B-GGUF:Q4_K_M")
            == "unsloth--Qwen3.5-9B-GGUF--Q4_K_M"
        )

    def test_bare_stem_passes_through(self):
        assert llm_client.model_name("my-model-q4") == "my-model-q4"

    def test_gguf_suffix_is_tolerated(self):
        assert llm_client.model_name("my-model-q4.gguf") == "my-model-q4"

    def test_ref_without_quant(self):
        assert llm_client.model_name("acme/tiny") == "acme--tiny"

    def test_model_file_lands_in_models_dir(self):
        path = llm_client.model_file("acme/tiny:Q8_0")
        assert path == llm_client.models_dir() / "acme--tiny--Q8_0.gguf"


class TestIsAvailable:
    @patch("llm_client.urllib.request.urlopen")
    def test_returns_true_when_server_responds(self, mock_urlopen):
        mock_urlopen.return_value.__enter__ = MagicMock()
        mock_urlopen.return_value.__exit__ = MagicMock()
        assert llm_client.is_available() is True
        assert "/health" in mock_urlopen.call_args[0][0].full_url

    @patch("llm_client.urllib.request.urlopen")
    def test_returns_false_on_connection_error(self, mock_urlopen):
        mock_urlopen.side_effect = urllib.error.URLError("refused")
        assert llm_client.is_available() is False

    @patch("llm_client.urllib.request.urlopen")
    def test_returns_false_on_timeout(self, mock_urlopen):
        mock_urlopen.side_effect = TimeoutError()
        assert llm_client.is_available() is False


class TestListModels:
    def test_scans_gguf_files(self, tmp_path):
        directory = llm_client.models_dir()
        directory.mkdir(parents=True)
        (directory / "beta.gguf").write_bytes(b"GGUF" + b"\0" * 10)
        (directory / "alpha.gguf").write_bytes(b"GGUF")
        (directory / "notes.txt").write_text("ignored")
        models = llm_client.list_models()
        assert [m["name"] for m in models] == ["alpha", "beta"]
        assert models[1]["size_bytes"] == 14

    def test_missing_dir_returns_empty(self):
        assert llm_client.list_models() == []


class TestIsModelInstalled:
    def test_file_present(self):
        directory = llm_client.models_dir()
        directory.mkdir(parents=True)
        (directory / "acme--tiny--Q8_0.gguf").write_bytes(b"GGUF")
        assert llm_client.is_model_installed("acme/tiny:Q8_0") is True

    def test_file_absent(self):
        assert llm_client.is_model_installed("acme/tiny:Q8_0") is False

    def test_empty_model_name(self):
        assert llm_client.is_model_installed("") is False

    def test_reuses_supplied_installed_list(self):
        installed = [{"name": "acme--tiny--Q8_0", "size_bytes": 4}]
        assert llm_client.is_model_installed("acme/tiny:Q8_0", installed) is True
        assert llm_client.is_model_installed("other/x:Q4_K_M", installed) is False


class TestUnloadModel:
    @patch("llm_client.urllib.request.urlopen")
    def test_posts_router_unload(self, mock_urlopen):
        mock_urlopen.return_value.__enter__ = MagicMock()
        mock_urlopen.return_value.__exit__ = MagicMock()
        assert llm_client.unload_model("acme/tiny:Q8_0") is True
        req = mock_urlopen.call_args[0][0]
        assert req.full_url.endswith("/models/unload")
        assert json.loads(req.data.decode()) == {"model": "acme--tiny--Q8_0"}

    @patch("llm_client.urllib.request.urlopen")
    def test_returns_false_on_connection_error(self, mock_urlopen):
        mock_urlopen.side_effect = urllib.error.URLError("refused")
        assert llm_client.unload_model("m") is False


class TestGenerate:
    @patch("llm_client.urllib.request.urlopen")
    def test_streams_and_concatenates_chunks(self, mock_urlopen):
        mock_urlopen.return_value = _make_streaming_resp(
            _sse_lines(_chunk("Hello"), _chunk(" world"), _chunk("", "stop"))
        )
        assert llm_client.generate("hi") == "Hello world"

    @patch("llm_client.urllib.request.urlopen")
    def test_request_shape(self, mock_urlopen):
        mock_urlopen.return_value = _make_streaming_resp(
            _sse_lines(_chunk("x", "stop"))
        )
        llm_client.generate("prompt text", model="acme/tiny:Q8_0", system="be brief")
        req = mock_urlopen.call_args[0][0]
        assert req.full_url.endswith("/v1/chat/completions")
        body = json.loads(req.data.decode())
        assert body["model"] == "acme--tiny--Q8_0"
        assert body["stream"] is True
        assert body["chat_template_kwargs"] == {"enable_thinking": False}
        assert body["messages"] == [
            {"role": "system", "content": "be brief"},
            {"role": "user", "content": "prompt text"},
        ]

    @patch("llm_client.urllib.request.urlopen")
    def test_no_system_message_when_absent(self, mock_urlopen):
        mock_urlopen.return_value = _make_streaming_resp(
            _sse_lines(_chunk("x", "stop"))
        )
        llm_client.generate("prompt text")
        body = json.loads(mock_urlopen.call_args[0][0].data.decode())
        assert [m["role"] for m in body["messages"]] == ["user"]

    @patch("llm_client.urllib.request.urlopen")
    def test_uses_config_model_by_default(self, mock_urlopen, monkeypatch):
        monkeypatch.setattr(llm_client.config, "LLM_SUMMARY_MODEL", "acme/tiny:Q8_0")
        mock_urlopen.return_value = _make_streaming_resp(
            _sse_lines(_chunk("x", "stop"))
        )
        llm_client.generate("hi")
        body = json.loads(mock_urlopen.call_args[0][0].data.decode())
        assert body["model"] == "acme--tiny--Q8_0"

    @patch("llm_client.urllib.request.urlopen")
    def test_returns_none_on_empty_stream(self, mock_urlopen):
        mock_urlopen.return_value = _make_streaming_resp(_sse_lines(_chunk("", "stop")))
        assert llm_client.generate("hi") is None

    @patch("llm_client.urllib.request.urlopen")
    def test_skips_malformed_sse_lines(self, mock_urlopen):
        lines = [
            b": keep-alive comment\n",
            b"data: {not json}\n",
            b"garbage without prefix\n",
        ] + _sse_lines(_chunk("ok", "stop"))
        mock_urlopen.return_value = _make_streaming_resp(lines)
        assert llm_client.generate("hi") == "ok"

    @patch("llm_client.urllib.request.urlopen")
    def test_error_chunk_returns_none(self, mock_urlopen):
        lines = [b'data: {"error": {"message": "boom"}}\n']
        mock_urlopen.return_value = _make_streaming_resp(lines)
        assert llm_client.generate("hi") is None

    @patch("llm_client.start_server")
    @patch("llm_client.urllib.request.urlopen")
    def test_returns_none_on_http_error(self, mock_urlopen, mock_start):
        mock_urlopen.side_effect = urllib.error.HTTPError(
            "url", 400, "Bad Request", email.message.Message(), None
        )
        assert llm_client.generate("hi") is None
        mock_start.assert_not_called()

    @patch("llm_client.urllib.request.urlopen")
    def test_cancel_event_set_before_call_returns_none(self, mock_urlopen):
        mock_urlopen.return_value = _make_streaming_resp(
            _sse_lines(_chunk("never", "stop"))
        )
        evt = threading.Event()
        evt.set()
        assert llm_client.generate("hi", cancel_event=evt) is None

    @patch("llm_client.urllib.request.urlopen")
    def test_on_token_fires_per_chunk_and_returns_joined(self, mock_urlopen):
        mock_urlopen.return_value = _make_streaming_resp(
            _sse_lines(_chunk("a"), _chunk("b"), _chunk("", "stop"))
        )
        seen: list[str] = []
        result = llm_client.generate("hi", on_token=seen.append)
        assert result == "ab"
        assert seen == ["a", "b"]

    @patch("llm_client.urllib.request.urlopen")
    def test_on_token_error_does_not_break_stream(self, mock_urlopen):
        mock_urlopen.return_value = _make_streaming_resp(
            _sse_lines(_chunk("a"), _chunk("b"), _chunk("", "stop"))
        )

        def _boom(_piece: str) -> None:
            raise RuntimeError("callback broke")

        assert llm_client.generate("hi", on_token=_boom) == "ab"


class TestGenerateStreamTruncation:
    @patch("llm_client.urllib.request.urlopen")
    def test_eof_without_finish_reason_returns_none(self, mock_urlopen):
        # No finish_reason and no [DONE]: a mid-flight unload or server death.
        mock_urlopen.return_value = _make_streaming_resp(
            _sse_lines(_chunk("partial "), _chunk("text"), done=False)
        )
        assert llm_client.generate("hi") is None

    @patch("llm_client.urllib.request.urlopen")
    def test_finish_reason_without_done_marker_still_returns_text(self, mock_urlopen):
        mock_urlopen.return_value = _make_streaming_resp(
            _sse_lines(_chunk("full"), _chunk("", "stop"), done=False)
        )
        assert llm_client.generate("hi") == "full"

    @patch("llm_client.urllib.request.urlopen")
    def test_non_dict_chunks_are_skipped(self, mock_urlopen):
        lines = [b"data: [1, 2, 3]\n"] + _sse_lines(_chunk("ok", "stop"))
        mock_urlopen.return_value = _make_streaming_resp(lines)
        assert llm_client.generate("hi") == "ok"

    @patch("llm_client.urllib.request.urlopen")
    def test_incomplete_read_is_handled_not_raised(self, mock_urlopen):
        resp = MagicMock()
        resp.readline.side_effect = http.client.IncompleteRead(b"partial")
        mock_urlopen.return_value = resp
        assert llm_client.generate("hi") is None


class _SSEChatHandler(http.server.BaseHTTPRequestHandler):
    """Test HTTP handler that mimics llama-server /v1/chat/completions SSE.

    Behavior is selected by the request body's user message:
      - "complete" → send 3 SSE chunks + [DONE], finish quickly.
      - "block"    → send 1 chunk, then sleep so the connection blocks
                     readline() until shut down.
    """

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(length).decode("utf-8"))
        prompt = body.get("messages", [{}])[-1].get("content", "")
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream")
        self.send_header("Transfer-Encoding", "chunked")
        self.end_headers()

        def _write_sse(obj) -> None:
            data = (
                b"data: "
                + (obj if isinstance(obj, bytes) else json.dumps(obj).encode("utf-8"))
                + b"\n\n"
            )
            self.wfile.write(b"%x\r\n" % len(data))
            self.wfile.write(data)
            self.wfile.write(b"\r\n")
            self.wfile.flush()

        try:
            if prompt == "complete":
                _write_sse(
                    {
                        "choices": [
                            {"delta": {"content": "Hello"}, "finish_reason": None}
                        ]
                    }
                )
                _write_sse(
                    {
                        "choices": [
                            {"delta": {"content": " world"}, "finish_reason": None}
                        ]
                    }
                )
                _write_sse({"choices": [{"delta": {}, "finish_reason": "stop"}]})
                _write_sse(b"[DONE]")
                self.wfile.write(b"0\r\n\r\n")
            elif prompt == "block":
                _write_sse(
                    {
                        "choices": [
                            {"delta": {"content": "Hello"}, "finish_reason": None}
                        ]
                    }
                )
                # Block; the test will close the socket via cancel.
                time.sleep(30)
        except (BrokenPipeError, ConnectionResetError, OSError):
            return

    def log_message(self, *_args, **_kwargs) -> None:
        return


@contextmanager
def _stub_chat_server():
    srv = socketserver.ThreadingTCPServer(("127.0.0.1", 0), _SSEChatHandler)
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
        with _stub_chat_server() as url:
            monkeypatch.setattr(llm_client.config, "LLM_BASE_URL", url)
            assert llm_client.generate("complete") == "Hello world"

    def test_cancel_aborts_blocked_request_quickly(self, monkeypatch):
        with _stub_chat_server() as url:
            monkeypatch.setattr(llm_client.config, "LLM_BASE_URL", url)
            evt = threading.Event()
            # Trigger cancel shortly after the request starts so the watcher
            # has to shut down the underlying socket to unblock readline().
            timer = threading.Timer(0.2, evt.set)
            timer.start()
            start = time.monotonic()
            try:
                result = llm_client.generate("block", cancel_event=evt)
            finally:
                timer.cancel()
            elapsed = time.monotonic() - start
            assert result is None
            # Real abort should return well under the 30s server sleep and
            # under the urlopen connect timeout. Use a generous bound to
            # avoid flake on slow CI.
            assert elapsed < 5.0, f"Cancel took {elapsed:.2f}s — abort is broken"


class TestAutoStartServer:
    """Tests for the auto-start behavior when the server is not running."""

    @patch("llm_client.start_server")
    @patch("llm_client.urllib.request.urlopen")
    def test_retries_after_connection_refused(self, mock_urlopen, mock_start):
        refused = urllib.error.URLError(ConnectionRefusedError("refused"))
        mock_urlopen.side_effect = [
            refused,
            _make_streaming_resp(_sse_lines(_chunk("ok", "stop"))),
        ]
        mock_start.return_value = True
        assert llm_client.generate("hi") == "ok"
        mock_start.assert_called_once()

    @patch("llm_client.start_server")
    @patch("llm_client.urllib.request.urlopen")
    def test_returns_none_when_server_start_fails(self, mock_urlopen, mock_start):
        mock_urlopen.side_effect = urllib.error.URLError(
            ConnectionRefusedError("refused")
        )
        mock_start.return_value = False
        assert llm_client.generate("hi") is None

    @patch("llm_client.start_server")
    @patch("llm_client.urllib.request.urlopen")
    def test_does_not_start_server_on_other_errors(self, mock_urlopen, mock_start):
        mock_urlopen.side_effect = urllib.error.URLError("some DNS failure")
        assert llm_client.generate("hi") is None
        mock_start.assert_not_called()

    @patch("llm_client.subprocess.Popen")
    @patch("llm_client.is_available")
    @patch("llm_client.shutil.which")
    def test_start_server_polls_until_available(
        self, mock_which, mock_available, mock_popen, monkeypatch
    ):
        mock_which.return_value = "/usr/local/bin/llama-server"
        mock_available.side_effect = [False, False, True]
        mock_popen.return_value.poll.return_value = None
        monkeypatch.setattr(llm_client, "_START_POLL_INTERVAL", 0.01)
        assert llm_client.start_server() is True
        args = mock_popen.call_args[0][0]
        assert args[0] == "/usr/local/bin/llama-server"
        assert "--models-dir" in args
        assert "--no-webui" in args
        assert "--models-max" in args

    @patch("llm_client.shutil.which")
    def test_start_server_returns_false_when_binary_missing(self, mock_which):
        mock_which.return_value = None
        assert llm_client.start_server() is False

    @patch("llm_client.utils.warning_print")
    @patch("llm_client.shutil.which")
    def test_start_server_shows_install_guidance(self, mock_which, mock_warn):
        mock_which.return_value = None
        llm_client.start_server()
        details = mock_warn.call_args.kwargs.get("details") or []
        assert any("llama" in line.lower() for line in details)

    @patch("llm_client.subprocess.Popen")
    @patch("llm_client.is_available")
    @patch("llm_client.shutil.which")
    def test_start_server_reports_immediate_death(
        self, mock_which, mock_available, mock_popen, monkeypatch
    ):
        mock_which.return_value = "/usr/local/bin/llama-server"
        mock_available.return_value = False
        mock_popen.return_value.poll.return_value = 1
        monkeypatch.setattr(llm_client, "_START_POLL_INTERVAL", 0.01)
        monkeypatch.setattr(llm_client, "_START_TIMEOUT", 0.2)
        assert llm_client.start_server() is False


def _tree_entry(path: str, size: int, sha: str | None = "a" * 64) -> dict:
    entry: dict = {"path": path, "size": size}
    if sha is not None:
        entry["lfs"] = {"oid": sha, "size": size}
    return entry


class TestDownloadModel:
    def _mock_hf(self, tree, payload: bytes = b"GGUF-payload"):
        """urlopen side effect serving the tree listing then the file body."""

        def _side_effect(req, timeout=None):
            url = req.full_url
            resp = MagicMock()
            if "/api/models/" in url:
                resp.read.return_value = json.dumps(tree).encode()
                resp.__enter__ = MagicMock(return_value=resp)
                resp.__exit__ = MagicMock(return_value=False)
                return resp
            chunks = [payload, b""]
            resp.read.side_effect = lambda *_a: chunks.pop(0)
            resp.headers = {"Content-Length": str(len(payload))}
            resp.headers = MagicMock()
            resp.headers.get.return_value = str(len(payload))
            resp.__enter__ = MagicMock(return_value=resp)
            resp.__exit__ = MagicMock(return_value=False)
            return resp

        return _side_effect

    def test_downloads_and_renames_into_place(self):
        import hashlib

        payload = b"GGUF-payload"
        sha = hashlib.sha256(payload).hexdigest()
        tree = [_tree_entry("tiny-Q8_0.gguf", len(payload), sha)]
        with patch("llm_client.urllib.request.urlopen") as mock_urlopen:
            mock_urlopen.side_effect = self._mock_hf(tree, payload)
            seen: list[dict] = []
            assert llm_client.download_model("acme/tiny:Q8_0", seen.append) is True
        target = llm_client.model_file("acme/tiny:Q8_0")
        assert target.read_bytes() == payload
        assert seen and seen[-1]["completed"] == len(payload)

    def test_rejects_sha_mismatch(self):
        tree = [_tree_entry("tiny-Q8_0.gguf", 12, "f" * 64)]
        with patch("llm_client.urllib.request.urlopen") as mock_urlopen:
            mock_urlopen.side_effect = self._mock_hf(tree)
            assert llm_client.download_model("acme/tiny:Q8_0") is False
        assert not llm_client.model_file("acme/tiny:Q8_0").exists()
        assert not list(llm_client.models_dir().glob("tmp*"))

    def test_size_check_when_sha_absent(self):
        tree = [_tree_entry("tiny-Q8_0.gguf", 999, sha=None)]
        with patch("llm_client.urllib.request.urlopen") as mock_urlopen:
            mock_urlopen.side_effect = self._mock_hf(tree)
            assert llm_client.download_model("acme/tiny:Q8_0") is False

    def test_quant_not_found(self):
        tree = [_tree_entry("tiny-Q2_K.gguf", 12)]
        with patch("llm_client.urllib.request.urlopen") as mock_urlopen:
            mock_urlopen.side_effect = self._mock_hf(tree)
            assert llm_client.download_model("acme/tiny:Q8_0") is False

    def test_sharded_model_rejected(self):
        tree = [
            _tree_entry("big-Q8_0-00001-of-00002.gguf", 12),
            _tree_entry("big-Q8_0-00002-of-00002.gguf", 12),
        ]
        with patch("llm_client.urllib.request.urlopen") as mock_urlopen:
            mock_urlopen.side_effect = self._mock_hf(tree)
            assert llm_client.download_model("acme/big:Q8_0") is False

    def test_gated_repo_reports_failure(self):
        with patch("llm_client.urllib.request.urlopen") as mock_urlopen:
            mock_urlopen.side_effect = urllib.error.HTTPError(
                "url", 401, "Unauthorized", email.message.Message(), None
            )
            assert llm_client.download_model("acme/tiny:Q8_0") is False

    def test_already_downloaded_returns_true_without_network(self):
        target = llm_client.model_file("acme/tiny:Q8_0")
        target.parent.mkdir(parents=True)
        target.write_bytes(b"GGUF")
        with patch("llm_client.urllib.request.urlopen") as mock_urlopen:
            assert llm_client.download_model("acme/tiny:Q8_0") is True
            mock_urlopen.assert_not_called()

    def test_bare_stem_is_not_downloadable(self):
        assert llm_client.download_model("local-model") is False

    def test_empty_ref(self):
        assert llm_client.download_model("") is False


class TestInstallGuidance:
    def test_guidance_mentions_llama_server(self):
        lines = llm_client.install_guidance_lines()
        assert lines
        assert any("llama" in line.lower() for line in lines)
