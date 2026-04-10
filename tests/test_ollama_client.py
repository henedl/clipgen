# -*- coding: utf-8 -*-
"""Tests for ollama_client module."""

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
        assert body["model"] == "qwen3.5:0.8b"

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

    @patch("ollama_client.subprocess.Popen")
    @patch("ollama_client.is_available")
    def test_start_server_polls_until_available(self, mock_avail, mock_popen):
        """_start_server polls is_available until it returns True."""
        mock_avail.side_effect = [False, False, True]
        assert ollama_client._start_server() is True
        assert mock_avail.call_count == 3
        mock_popen.assert_called_once()

    @patch("ollama_client.subprocess.Popen")
    def test_start_server_returns_false_when_binary_missing(self, mock_popen):
        """_start_server returns False when ollama is not installed."""
        mock_popen.side_effect = FileNotFoundError("ollama not found")
        assert ollama_client._start_server() is False


class TestSummarizeTranscript:
    def test_returns_none_for_empty_segments(self):
        result = ollama_client.summarize_transcript([])
        assert result is None

    def test_returns_none_for_very_short_text(self):
        result = ollama_client.summarize_transcript([{"text": "Hi"}])
        assert result is None

    @patch("ollama_client.generate")
    def test_concatenates_segment_text(self, mock_generate):
        mock_generate.return_value = "A summary."
        segments = [
            {"text": "First segment text here."},
            {"text": "Second segment text here."},
            {"text": "Third segment more text."},
        ]
        ollama_client.summarize_transcript(segments)

        prompt = mock_generate.call_args[0][0]
        assert "First segment text here." in prompt
        assert "Second segment text here." in prompt
        assert "Third segment more text." in prompt

    @patch("ollama_client.generate")
    def test_returns_generate_result(self, mock_generate):
        mock_generate.return_value = "This is a summary.\n- Point one\n- Point two"
        segments = [
            {
                "text": "A sufficiently long segment of text for the minimum length check."
            },
        ]
        result = ollama_client.summarize_transcript(segments)
        assert result == "This is a summary.\n- Point one\n- Point two"

    @patch("ollama_client.generate")
    def test_returns_none_when_generate_fails(self, mock_generate):
        mock_generate.return_value = None
        segments = [
            {
                "text": "A sufficiently long segment of text for the minimum length check."
            },
        ]
        result = ollama_client.summarize_transcript(segments)
        assert result is None

    @patch("ollama_client.generate")
    def test_passes_model_override(self, mock_generate):
        mock_generate.return_value = "ok"
        segments = [
            {
                "text": "A sufficiently long segment of text for the minimum length check."
            },
        ]
        ollama_client.summarize_transcript(segments, model="llama3.1:8b")
        assert mock_generate.call_args[1]["model"] == "llama3.1:8b"

    @patch("ollama_client.generate")
    def test_uses_small_model_for_short_text(self, mock_generate):
        mock_generate.return_value = "ok"
        # Short text (under _LARGE_MODEL_THRESHOLD)
        segments = [
            {
                "text": "A sufficiently long segment of text for the minimum length check."
            },
        ]
        ollama_client.summarize_transcript(segments)
        assert mock_generate.call_args[1]["model"] == "qwen3.5:0.8b"

    @patch("ollama_client.generate")
    def test_uses_large_model_for_long_text(self, mock_generate):
        mock_generate.return_value = "ok"
        # Long text (over _LARGE_MODEL_THRESHOLD of 8000 chars)
        segments = [{"text": "word " * 1700}]
        ollama_client.summarize_transcript(segments)
        assert mock_generate.call_args[1]["model"] == "qwen3.5:9b"
