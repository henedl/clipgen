# -*- coding: utf-8 -*-
"""Ollama local LLM client for clipgen.

Provides a reusable interface to the Ollama REST API for text generation.
All functions fail gracefully (return None / False) and never raise on
network errors. On connection refused, generate() auto-starts ``ollama serve``
and retries once.

Key functions:
  is_available()            - check Ollama server connectivity
  generate()                - send a prompt and get a text response
  summarize_transcript()    - summarize transcript segments into paragraph + bullets
"""

import json
import re
import subprocess
import time
import urllib.error
import urllib.request
from typing import Any

import config
import utils

_HEALTH_TIMEOUT = 5  # seconds for connectivity check
_GENERATE_TIMEOUT = (
    300  # seconds — generous to allow cold model loading + long transcripts
)
_START_POLL_INTERVAL = 0.5  # seconds between health-check polls after starting server
_START_TIMEOUT = 10  # seconds to wait for server to become available after starting
_MAX_TRANSCRIPT_CHARS = 6000  # truncate long transcripts to fit model context window
_THINK_RE = re.compile(r"<think>[\s\S]*?</think>\s*", re.DOTALL)


def is_available() -> bool:
    """Check whether the Ollama server is reachable.

    Returns True if the server responds to GET /api/tags, False otherwise.
    Does not log on failure — callers decide how to handle unavailability.
    """
    try:
        req = urllib.request.Request(f"{config.OLLAMA_BASE_URL}/api/tags")
        with urllib.request.urlopen(req, timeout=_HEALTH_TIMEOUT):
            return True
    except (urllib.error.URLError, OSError, ValueError):
        return False


def _is_connection_refused(exc: Exception) -> bool:
    """Check whether an exception indicates connection refused (server not running)."""
    if isinstance(exc, urllib.error.URLError) and isinstance(
        exc.reason, ConnectionRefusedError
    ):
        return True
    if isinstance(exc, ConnectionRefusedError):
        return True
    return False


def _start_server() -> bool:
    """Attempt to start ``ollama serve`` and wait for it to become available.

    Returns True if the server is responding after startup, False otherwise.
    """
    utils.info_print("Starting Ollama server...")
    try:
        subprocess.Popen(
            ["ollama", "serve"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except FileNotFoundError:
        utils.warning_print("Ollama binary not found — is Ollama installed?")
        return False
    except OSError as exc:
        utils.warning_print(f"Failed to start Ollama: {exc}")
        return False

    deadline = time.monotonic() + _START_TIMEOUT
    while time.monotonic() < deadline:
        if is_available():
            utils.info_print("Ollama server started.")
            return True
        time.sleep(_START_POLL_INTERVAL)

    utils.warning_print("Ollama server did not start within timeout.")
    return False


def _do_generate(
    body: dict[str, Any],
) -> str | None:
    """Execute a single Ollama /api/generate request. Returns text or None."""
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        f"{config.OLLAMA_BASE_URL}/api/generate",
        data=data,
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=_GENERATE_TIMEOUT) as resp:
        result = json.loads(resp.read().decode("utf-8"))
        text = result.get("response", "")
        # Strip <think>...</think> blocks produced by reasoning models (e.g. qwen3.5)
        text = _THINK_RE.sub("", text).strip()
        if not text:
            utils.warning_print(
                f"Ollama returned empty response (model: {body.get('model')})"
            )
            return None
        return text


def generate(
    prompt: str,
    *,
    model: str | None = None,
    system: str | None = None,
) -> str | None:
    """Send a prompt to Ollama and return the generated text.

    On connection refused, automatically attempts to start ``ollama serve``
    and retries the request once.

    Args:
        prompt: The user prompt to send.
        model: Ollama model name. Defaults to config.OLLAMA_SUMMARY_MODEL.
        system: Optional system prompt.

    Returns the generated text string, or None on any failure.
    """
    resolved_model = model or config.OLLAMA_SUMMARY_MODEL
    body: dict[str, Any] = {
        "model": resolved_model,
        "prompt": prompt,
        "stream": False,
    }
    if system:
        body["system"] = system

    try:
        return _do_generate(body)
    except urllib.error.HTTPError as exc:
        utils.warning_print(f"Ollama generate failed (HTTP {exc.code}): {exc.reason}")
        return None
    except (urllib.error.URLError, OSError) as exc:
        if not _is_connection_refused(exc):
            utils.warning_print(f"Ollama generate failed (connection): {exc}")
            return None
        # Connection refused — try to start the server and retry once
        if not _start_server():
            return None
        try:
            return _do_generate(body)
        except (urllib.error.URLError, urllib.error.HTTPError, OSError) as retry_exc:
            utils.warning_print(f"Ollama generate failed after retry: {retry_exc}")
            return None
        except (json.JSONDecodeError, KeyError, ValueError) as retry_exc:
            utils.warning_print(f"Ollama generate failed (response): {retry_exc}")
            return None
    except (json.JSONDecodeError, KeyError, ValueError) as exc:
        utils.warning_print(f"Ollama generate failed (response): {exc}")
        return None


_SUMMARIZE_PROMPT = """\
Summarize this user research session transcript. Write a concise paragraph \
(2-4 sentences) describing what happened in the session. Then list the key \
topics or themes as bullet points (prefix each with "- ").

Transcript:
{text}"""

_MIN_TEXT_LENGTH = 50  # skip summarization for very short transcripts
_LARGE_MODEL_THRESHOLD = 8000  # chars — use the larger model above this


def summarize_transcript(
    segments: list[dict[str, Any]],
    *,
    model: str | None = None,
) -> str | None:
    """Summarize transcript segments into a paragraph + bullet points.

    Automatically selects between OLLAMA_SUMMARY_MODEL (short transcripts) and
    OLLAMA_SUMMARY_MODEL_LARGE (longer transcripts) unless an explicit model
    override is provided.

    Args:
        segments: List of transcript segment dicts with "text" keys.
        model: Ollama model override. Skips auto-selection when set.

    Returns the summary string, or None if segments are empty/too short
    or if generation fails.
    """
    text = " ".join(seg.get("text", "").strip() for seg in segments).strip()
    if len(text) < _MIN_TEXT_LENGTH:
        return None

    # Auto-select model based on transcript length
    if model is None:
        if len(text) > _LARGE_MODEL_THRESHOLD:
            model = config.OLLAMA_SUMMARY_MODEL_LARGE
        else:
            model = config.OLLAMA_SUMMARY_MODEL

    # Truncate to fit model context window — keep beginning and end for
    # a representative overview of the session
    if len(text) > _MAX_TRANSCRIPT_CHARS:
        half = _MAX_TRANSCRIPT_CHARS // 2
        text = text[:half] + "\n[...]\n" + text[-half:]

    utils.verbose_print(
        f"Summarizing transcript ({len(segments)} segments, "
        f"{len(text)} chars) with model {model}"
    )
    prompt = _SUMMARIZE_PROMPT.format(text=text)
    result = generate(prompt, model=model)
    if result:
        utils.verbose_print(f"Summary generated ({len(result)} chars)")
    return result
