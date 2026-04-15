# -*- coding: utf-8 -*-
"""Ollama local LLM transport for clipgen.

A thin, reusable HTTP wrapper around the Ollama REST API. All functions fail
gracefully (return None / False) and never raise on network errors. On
connection refused, ``generate()`` auto-starts ``ollama serve`` and retries
once.

This module is intentionally small — higher-level reasoning lives in
[thinking_agents.py](thinking_agents.py), which routes every call through
``generate()`` here.

Key functions:
  is_available()  - check Ollama server connectivity
  list_models()   - enumerate installed models with metadata
  generate()      - send a prompt and get a text response
"""

import json
import re
import shutil
import subprocess
import sys
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


def list_models() -> list[dict[str, Any]] | None:
    """List installed Ollama models with metadata.

    Calls GET /api/tags on the Ollama server and parses the response into a
    list of dicts with keys: name, size_bytes, parameter_size, quantization,
    family.  Returns None on any failure (server unreachable, bad JSON, etc.).
    """
    try:
        req = urllib.request.Request(f"{config.OLLAMA_BASE_URL}/api/tags")
        with urllib.request.urlopen(req, timeout=_HEALTH_TIMEOUT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            models = []
            for m in data.get("models", []):
                details = m.get("details", {})
                models.append(
                    {
                        "name": m.get("name", ""),
                        "size_bytes": m.get("size", 0),
                        "parameter_size": details.get("parameter_size", ""),
                        "quantization": details.get("quantization_level", ""),
                        "family": details.get("family", ""),
                    }
                )
            return models
    except (urllib.error.URLError, OSError, ValueError, json.JSONDecodeError):
        return None


def _is_connection_refused(exc: Exception) -> bool:
    """Check whether an exception indicates connection refused (server not running)."""
    if isinstance(exc, urllib.error.URLError) and isinstance(
        exc.reason, ConnectionRefusedError
    ):
        return True
    if isinstance(exc, ConnectionRefusedError):
        return True
    return False


def _ollama_install_guidance_lines() -> list[str]:
    """Return actionable install guidance based on the current platform."""
    platform_specific = []
    if sys.platform == "darwin":
        platform_specific = [
            "macOS: install with Homebrew: brew install ollama",
        ]
    elif sys.platform.startswith("linux"):
        platform_specific = [
            "Linux: curl -fsSL https://ollama.com/install.sh | sh",
        ]
    elif sys.platform.startswith("win"):
        platform_specific = [
            "Windows: winget install Ollama.Ollama",
        ]
    else:
        platform_specific = [
            "Download from: https://ollama.com/download",
        ]

    return platform_specific + [
        "Then verify in a new terminal:",
        "  ollama --version",
    ]


def _start_server() -> bool:
    """Attempt to start ``ollama serve`` and wait for it to become available.

    Returns True if the server is responding after startup, False otherwise.
    """
    if shutil.which("ollama") is None:
        utils.warning_print(
            "Ollama is not installed.",
            details=_ollama_install_guidance_lines(),
        )
        return False

    utils.info_print("Starting Ollama server...")
    try:
        subprocess.Popen(
            ["ollama", "serve"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
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
    think: bool | None = None,
) -> str | None:
    """Send a prompt to Ollama and return the generated text.

    On connection refused, automatically attempts to start ``ollama serve``
    and retries the request once.

    Args:
        prompt: The user prompt to send.
        model: Ollama model name. Defaults to config.OLLAMA_SUMMARY_MODEL.
        system: Optional system prompt.
        think: Explicitly enable/disable thinking mode for reasoning models.
            When False, the model skips chain-of-thought and responds directly.

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
    if think is not None:
        body["think"] = think

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
