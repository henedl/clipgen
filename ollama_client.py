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
  unload_model()  - ask Ollama to evict a model from memory immediately
"""

import json
import re
import shutil
import socket
import subprocess
import sys
import threading
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
# Wall-clock deadline for the entire streaming response. urlopen's timeout only
# applies to connect + first byte, so a slow trickle can stall a worker
# indefinitely. This bounds total elapsed time end-to-end.
_GENERATE_DEADLINE = 600  # seconds
_START_POLL_INTERVAL = 0.5  # seconds between health-check polls after starting server
_START_TIMEOUT = 10  # seconds to wait for server to become available after starting
_CANCEL_WATCHER_POLL = 1.0  # seconds; bounds abort latency during long quiet stretches
_THINK_RE = re.compile(r"<think>[\s\S]*?</think>\s*", re.DOTALL)

# Serializes _start_server() calls so two threads hitting connection-refused at
# the same time don't both spawn `ollama serve`.
_start_server_lock = threading.Lock()


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


def unload_model(model: str) -> bool:
    """Ask Ollama to evict *model* from memory immediately.

    Sends a minimal /api/generate request with ``keep_alive: 0``, which tells
    Ollama to unload the model as soon as the call returns. Returns True if
    the request was accepted, False on any failure.
    """
    body = {
        "model": model,
        "prompt": "",
        "stream": False,
        "keep_alive": 0,
    }
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        f"{config.OLLAMA_BASE_URL}/api/generate",
        data=data,
        headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=_HEALTH_TIMEOUT) as resp:
            resp.read()  # drain
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

    Serialized via ``_start_server_lock`` so concurrent connection-refused
    retries don't both spawn a server process — the first one to acquire the
    lock spawns it, the rest see the already-running server when they re-poll.

    Returns True if the server is responding after startup, False otherwise.
    """
    if shutil.which("ollama") is None:
        utils.warning_print(
            "Ollama is not installed.",
            details=_ollama_install_guidance_lines(),
        )
        return False

    with _start_server_lock:
        # Another thread may have already started the server while we waited.
        if is_available():
            return True

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


def _shutdown_response_socket(resp: Any) -> None:
    """Force-close the underlying socket of a urllib HTTPResponse.

    Plain ``resp.close()`` does NOT unblock a ``readline()`` blocked on a
    socket from another thread. We have to reach down to the raw socket and
    shut it down so the kernel returns from the syscall on the worker thread.

    We deliberately do not call ``resp.close()`` here — that mutates
    ``resp.fp`` and races with the worker's blocked read, which can surface
    as ``AttributeError`` when the read wakes up. The worker's ``finally``
    block will close the response after the loop returns.
    """
    try:
        sock = resp.fp.raw._sock  # type: ignore[attr-defined]
    except AttributeError:
        sock = None
    if sock is not None:
        try:
            sock.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass


def _do_generate(
    body: dict[str, Any],
    cancel_event: threading.Event | None = None,
) -> str | None:
    """Execute a single Ollama /api/generate request (streaming).

    Reads the NDJSON stream chunk-by-chunk, accumulating ``response`` fields
    until a chunk arrives with ``done: true``. When *cancel_event* is set, a
    watcher thread shuts down the underlying socket so the blocked readline()
    unblocks and the function returns ``None`` promptly — freeing Ollama for
    another run.
    """
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        f"{config.OLLAMA_BASE_URL}/api/generate",
        data=data,
        headers={"Content-Type": "application/json"},
    )
    parts: list[str] = []
    done_event = threading.Event()
    watcher: threading.Thread | None = None
    deadline = time.monotonic() + _GENERATE_DEADLINE
    deadline_exceeded = False
    resp: Any = None
    try:
        resp = urllib.request.urlopen(req, timeout=_GENERATE_TIMEOUT)
        if cancel_event is not None:

            def _watch() -> None:
                while not done_event.is_set():
                    if cancel_event.wait(timeout=_CANCEL_WATCHER_POLL):
                        _shutdown_response_socket(resp)
                        return

            watcher = threading.Thread(
                target=_watch, daemon=True, name="ollama-cancel-watcher"
            )
            watcher.start()

        if cancel_event is not None and cancel_event.is_set():
            return None
        while True:
            if time.monotonic() >= deadline:
                deadline_exceeded = True
                _shutdown_response_socket(resp)
                return None
            try:
                line = resp.readline()
            except (OSError, ValueError, AttributeError):
                # Raised when the watcher shuts down the socket mid-read, or
                # when we shut it down above on deadline. http.client may
                # surface this as AttributeError on Python 3.13+ when the
                # chunked-encoding state machine tries to advance past a
                # closed fp.
                return None
            if not line:
                break
            if cancel_event is not None and cancel_event.is_set():
                return None
            if time.monotonic() >= deadline:
                deadline_exceeded = True
                _shutdown_response_socket(resp)
                return None
            try:
                chunk = json.loads(line.decode("utf-8"))
            except (json.JSONDecodeError, UnicodeDecodeError):
                continue  # skip malformed lines, keep reading
            piece = chunk.get("response", "")
            if piece:
                parts.append(piece)
            if chunk.get("done"):
                break
    finally:
        done_event.set()
        if resp is not None:
            try:
                resp.close()
            except Exception:
                pass
        if watcher is not None:
            watcher.join(timeout=0.5)
        if deadline_exceeded:
            utils.warning_print(
                f"Ollama generate exceeded {_GENERATE_DEADLINE}s deadline "
                f"(model: {body.get('model')}); aborting."
            )

    if cancel_event is not None and cancel_event.is_set():
        return None
    text = "".join(parts)
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
    cancel_event: threading.Event | None = None,
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
        cancel_event: Optional event used to abort the in-flight request. When
            set, the underlying HTTP response is closed and ``None`` is
            returned promptly so Ollama is freed for another run.

    Returns the generated text string, or None on any failure or cancellation.
    """
    resolved_model = model or config.OLLAMA_SUMMARY_MODEL
    body: dict[str, Any] = {
        "model": resolved_model,
        "prompt": prompt,
        "stream": True,
    }
    if system:
        body["system"] = system
    if think is not None:
        body["think"] = think

    try:
        return _do_generate(body, cancel_event=cancel_event)
    except urllib.error.HTTPError as exc:
        utils.warning_print(f"Ollama generate failed (HTTP {exc.code}): {exc.reason}")
        return None
    except (urllib.error.URLError, OSError) as exc:
        if cancel_event is not None and cancel_event.is_set():
            return None
        if not _is_connection_refused(exc):
            utils.warning_print(f"Ollama generate failed (connection): {exc}")
            return None
        # Connection refused — try to start the server and retry once
        if not _start_server():
            return None
        try:
            return _do_generate(body, cancel_event=cancel_event)
        except (
            urllib.error.URLError,
            urllib.error.HTTPError,
            OSError,
        ) as retry_exc:
            if cancel_event is not None and cancel_event.is_set():
                return None
            utils.warning_print(f"Ollama generate failed after retry: {retry_exc}")
            return None
        except (json.JSONDecodeError, KeyError, ValueError) as retry_exc:
            if cancel_event is not None and cancel_event.is_set():
                return None
            utils.warning_print(f"Ollama generate failed (response): {retry_exc}")
            return None
    except (json.JSONDecodeError, KeyError, ValueError) as exc:
        if cancel_event is not None and cancel_event.is_set():
            return None
        utils.warning_print(f"Ollama generate failed (response): {exc}")
        return None
