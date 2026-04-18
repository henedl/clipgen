# -*- coding: utf-8 -*-
"""Tests for ollama_client transport layer.

Agent-specific behavior (summarization, citation linking) lives in
tests/test_thinking_agents.py.
"""

import io
import json
import urllib.error
from unittest.mock import MagicMock, patch

import ollama_client


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


class TestGenerate:
    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_response_text_on_success(self, mock_urlopen):
        response_data = json.dumps({"response": "Hello world"}).encode("utf-8")
        mock_resp = io.BytesIO(response_data)
        mock_urlopen.return_value.__enter__ = lambda s: mock_resp
        mock_urlopen.return_value.__exit__ = lambda s, *a: None
        result = ollama_client.generate("test prompt")
        assert result == "Hello world"

    @patch("ollama_client.urllib.request.urlopen")
    def test_strips_think_tags_from_response(self, mock_urlopen):
        raw = "<think>Let me analyze this...</think>\n\nHere is the summary."
        response_data = json.dumps({"response": raw}).encode("utf-8")
        mock_resp = io.BytesIO(response_data)
        mock_urlopen.return_value.__enter__ = lambda s: mock_resp
        mock_urlopen.return_value.__exit__ = lambda s, *a: None
        result = ollama_client.generate("test prompt")
        assert result == "Here is the summary."

    @patch("ollama_client.urllib.request.urlopen")
    def test_returns_none_when_only_think_tags(self, mock_urlopen):
        raw = "<think>Thinking hard but producing nothing useful...</think>"
        response_data = json.dumps({"response": raw}).encode("utf-8")
        mock_resp = io.BytesIO(response_data)
        mock_urlopen.return_value.__enter__ = lambda s: mock_resp
        mock_urlopen.return_value.__exit__ = lambda s, *a: None
        result = ollama_client.generate("test prompt")
        assert result is None

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
    def test_returns_none_on_invalid_json(self, mock_urlopen):
        mock_resp = io.BytesIO(b"not json")
        mock_urlopen.return_value.__enter__ = lambda s: mock_resp
        mock_urlopen.return_value.__exit__ = lambda s, *a: None
        result = ollama_client.generate("test prompt")
        assert result is None

    @patch("ollama_client.urllib.request.urlopen")
    def test_uses_config_model_by_default(self, mock_urlopen):
        response_data = json.dumps({"response": "ok"}).encode("utf-8")
        mock_resp = io.BytesIO(response_data)
        mock_urlopen.return_value.__enter__ = lambda s: mock_resp
        mock_urlopen.return_value.__exit__ = lambda s, *a: None

        ollama_client.generate("test prompt")

        call_args = mock_urlopen.call_args
        req = call_args[0][0]
        body = json.loads(req.data.decode("utf-8"))
        assert body["model"] == "qwen3.5:9b"

    @patch("ollama_client.urllib.request.urlopen")
    def test_uses_custom_model(self, mock_urlopen):
        response_data = json.dumps({"response": "ok"}).encode("utf-8")
        mock_resp = io.BytesIO(response_data)
        mock_urlopen.return_value.__enter__ = lambda s: mock_resp
        mock_urlopen.return_value.__exit__ = lambda s, *a: None

        ollama_client.generate("test prompt", model="llama3.1:8b")

        call_args = mock_urlopen.call_args
        req = call_args[0][0]
        body = json.loads(req.data.decode("utf-8"))
        assert body["model"] == "llama3.1:8b"

    @patch("ollama_client.urllib.request.urlopen")
    def test_includes_system_prompt(self, mock_urlopen):
        response_data = json.dumps({"response": "ok"}).encode("utf-8")
        mock_resp = io.BytesIO(response_data)
        mock_urlopen.return_value.__enter__ = lambda s: mock_resp
        mock_urlopen.return_value.__exit__ = lambda s, *a: None

        ollama_client.generate("test prompt", system="You are helpful.")

        call_args = mock_urlopen.call_args
        req = call_args[0][0]
        body = json.loads(req.data.decode("utf-8"))
        assert body["system"] == "You are helpful."


class TestAutoStartServer:
    """Tests for the auto-start behavior when Ollama is not running."""

    @patch("ollama_client._start_server", return_value=True)
    @patch("ollama_client.urllib.request.urlopen")
    def test_retries_after_connection_refused(self, mock_urlopen, mock_start):
        """On ConnectionRefusedError, start server and retry successfully."""
        response_data = json.dumps({"response": "retried ok"}).encode("utf-8")
        mock_resp = io.BytesIO(response_data)

        # First call: connection refused; second call: success
        mock_urlopen.side_effect = [
            urllib.error.URLError(ConnectionRefusedError("Connection refused")),
            MagicMock(__enter__=lambda s: mock_resp, __exit__=lambda s, *a: None),
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
