"""Local LLM transport for clipgen, backed by llama.cpp's ``llama-server``.

A thin HTTP wrapper around llama-server running in router mode
(``--models-dir``): server lifecycle, model-dir scanning, GGUF downloads from
Hugging Face, generation over ``/v1/chat/completions``, and model unload.
Every function fails gracefully (returns None / False) and never raises on a
network error; on connection refused, ``generate()`` auto-starts the server
and retries once.

The router registers each ``*.gguf`` in the models dir under its filename
stem and spawns one child process per loaded model. ``model_name()`` maps a
user-facing model value — a Hugging Face ref like ``unsloth/Qwen3.5-9B-GGUF:Q4_K_M``
or a bare stem for a hand-dropped file — onto that id; the same mapping names
the file ``download_model()`` writes, so ref, file, and router id always agree.

Deliberately small — higher-level reasoning lives in
[thinking_agents.py](thinking_agents.py), which routes every call through
``generate()`` here.
"""

import atexit
import hashlib
import http.client
import json
import shutil
import socket
import subprocess
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from collections.abc import Callable
from pathlib import Path
from typing import Any

import config
import profiling
import start_settings
import utils

_HEALTH_TIMEOUT = 5  # seconds for connectivity check
_GENERATE_TIMEOUT = (
    300  # seconds — generous to allow cold model loading + long transcripts
)
# Wall-clock deadline for the entire streaming response. urlopen's timeout only
# applies to connect + first byte, so a slow trickle can stall a worker
# indefinitely. This bounds total elapsed time end-to-end.
_GENERATE_DEADLINE = 600  # seconds
# The router answers 503 while a model instance is still coming up. The gate
# run only ever saw requests block until ready, so this is cheap insurance.
_LOAD_RETRY_WINDOW = 120  # seconds
_LOAD_RETRY_INTERVAL = 2.0  # seconds
_START_POLL_INTERVAL = 0.5  # seconds between health-check polls after starting server
_START_TIMEOUT = 15  # seconds to wait for server to become available after starting
_CANCEL_WATCHER_POLL = 1.0  # seconds; bounds abort latency during long quiet stretches
_DOWNLOAD_CHUNK = 1024 * 1024
_DOWNLOAD_TIMEOUT = 120  # seconds; per-socket-read stall, HF is normally fast
# Wall-clock backstop for the whole download, same reasoning as
# _GENERATE_DEADLINE: _DOWNLOAD_TIMEOUT bounds a stall *between* reads, so a
# server trickling a byte per window never trips it. Generous — a 9B Q4 GGUF
# is ~6 GB, which is ~55 min on a 15 Mbit line.
_DOWNLOAD_DEADLINE = 5400  # seconds
_HF_API_TIMEOUT = 30  # seconds; the tree listing is a small JSON response
# Loaded models the router keeps resident before LRU-evicting. Two covers the
# "friction on its own model while summary stays warm" case without inviting
# three 9B models into RAM.
_MODELS_MAX = "2"

# Serializes start_server() calls so two threads hitting connection-refused at
# the same time don't both spawn a router process.
_start_server_lock = threading.Lock()

# The router we spawned, if any — terminated at exit. SIGTERM, never SIGKILL:
# the router reaps its per-model children only on a clean shutdown (gate
# scenario 6: a killed router orphans them).
_server_proc: subprocess.Popen[bytes] | None = None


def models_dir() -> Path:
    """Where clipgen keeps its GGUF models, under the per-user config dir."""
    return start_settings.config_dir() / "models"


def model_name(value: str) -> str:
    """Router model id for a settings value (HF ref or bare stem).

    A Hugging Face ref (``user/repo[:QUANT]``) maps to the deterministic stem
    ``download_model()`` writes; anything else is treated as the stem of a
    file already in the models dir, with a tolerated ``.gguf`` suffix.
    """
    value = (value or "").strip()
    if "/" in value:
        return value.replace("/", "--").replace(":", "--")
    return value.removesuffix(".gguf")


def model_file(value: str) -> Path:
    """Local GGUF path for a settings value."""
    return models_dir() / f"{model_name(value)}.gguf"


def is_available() -> bool:
    """Check whether the llama-server router is reachable.

    Returns True if the server responds to GET /health, False otherwise.
    Does not log on failure — callers decide how to handle unavailability.
    """
    try:
        req = urllib.request.Request(f"{config.LLM_BASE_URL}/health")
        with urllib.request.urlopen(req, timeout=_HEALTH_TIMEOUT):
            return True
    except (urllib.error.URLError, OSError, ValueError, http.client.HTTPException):
        return False


def list_models() -> list[dict[str, Any]]:
    """List downloaded GGUF models from the models dir.

    A filesystem scan, not a server call, so it answers even while the server
    is stopped. Returns dicts with keys: name (router id / stem), size_bytes.
    """
    directory = models_dir()
    if not directory.is_dir():
        return []
    models = []
    for path in sorted(directory.glob("*.gguf")):
        try:
            size = path.stat().st_size
        except OSError:
            continue
        models.append({"name": path.stem, "size_bytes": size})
    return models


def is_model_installed(
    model: str, installed: list[dict[str, Any]] | None = None
) -> bool:
    """Return True if *model*'s GGUF file exists in the models dir.

    Pass *installed* (the result of ``list_models()``) to skip the stat when
    the caller already holds the scan.
    """
    if not model:
        return False
    name = model_name(model)
    if installed is not None:
        return any(m["name"] == name for m in installed)
    return model_file(model).is_file()


def unload_model(model: str) -> bool:
    """Ask the router to evict *model*'s instance from memory now.

    Returns True if the request was accepted, False on any failure (including
    the model not being loaded — harmless for the delayed-unload caller).
    """
    body = {"model": model_name(model)}
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        f"{config.LLM_BASE_URL}/models/unload",
        data=data,
        headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=_HEALTH_TIMEOUT) as resp:
            resp.read()  # drain
            return True
    except (urllib.error.URLError, OSError, ValueError, http.client.HTTPException):
        return False


def _is_connection_refused(exc: Exception) -> bool:
    """Check whether an exception indicates connection refused (server not running)."""
    if isinstance(exc, urllib.error.URLError) and isinstance(
        exc.reason, ConnectionRefusedError
    ):
        return True
    return isinstance(exc, ConnectionRefusedError)


def install_guidance_lines() -> list[str]:
    """Return actionable install guidance based on the current platform.

    Public because the guidance has to reach the browser as well as the
    terminal — ``/api/models`` ships it so the Transcripts and Overview pages
    can tell a user *how* to get llama.cpp instead of only that it is missing.
    Frozen builds bundle ``llama-server``, so this only ever shows on
    source-tree runs.
    """
    return utils.install_guidance_lines(
        brew_command="brew install llama.cpp",
        linux=["Linux: download a llama.cpp release and put llama-server on PATH"],
        windows=["Windows: winget install ggml.llamacpp"],
        download_url="https://github.com/ggml-org/llama.cpp/releases",
        verify_commands=["llama-server --version"],
    )


def resolve_server_bin() -> str | None:
    """The ``llama-server`` binary clipgen should run, or None when absent.

    PATH resolution only: frozen builds prepend their bundled ``bin/`` before
    anything calls ``shutil.which``, so the bundled copy is found the same way
    a brew install is.
    """
    return shutil.which("llama-server")


def is_installed() -> bool:
    """Return True when a ``llama-server`` binary is reachable on PATH.

    Deliberately distinct from ``is_available()``, which reports whether the
    *server* answers. The two states need opposite advice — "start it, then
    refresh" is useless to someone who never installed it — so every surface
    that gates on the runtime reads both rather than collapsing them into one
    flag.
    """
    return resolve_server_bin() is not None


def _base_host_port() -> tuple[str, str]:
    """Host and port parsed from ``config.LLM_BASE_URL``."""
    parsed = urllib.parse.urlparse(config.LLM_BASE_URL)
    return parsed.hostname or "127.0.0.1", str(parsed.port or 8790)


def _terminate_server() -> None:
    """SIGTERM our router at exit so it reaps its model children."""
    proc = _server_proc
    if proc is None or proc.poll() is not None:
        return
    proc.terminate()
    try:
        proc.wait(timeout=5)
    except subprocess.TimeoutExpired:
        proc.kill()


atexit.register(_terminate_server)


def start_server() -> bool:
    """Start ``llama-server`` in router mode and wait for it to answer.

    Serialized via ``_start_server_lock`` so concurrent connection-refused
    retries don't both spawn a server process — the first one to acquire the
    lock spawns it, the rest see the already-running server when they re-poll.

    Returns True if the server is responding after startup, False otherwise.
    """
    binary = resolve_server_bin()
    if binary is None:
        utils.warning_print(
            "llama-server is not installed.",
            details=install_guidance_lines(),
        )
        return False

    with _start_server_lock:
        # Another thread may have already started the server while we waited.
        if is_available():
            return True

        directory = models_dir()
        try:
            directory.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            utils.warning_print(f"Could not create the models dir: {exc}")
            return False

        host, port = _base_host_port()
        utils.info_print("Starting AI server...")
        try:
            proc = subprocess.Popen(
                [
                    binary,
                    "--models-dir",
                    str(directory),
                    "--host",
                    host,
                    "--port",
                    port,
                    "--no-webui",
                    "--models-max",
                    _MODELS_MAX,
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except OSError as exc:
            utils.warning_print(f"Failed to start AI server: {exc}")
            return False

        global _server_proc
        _server_proc = proc

        deadline = time.monotonic() + _START_TIMEOUT
        while time.monotonic() < deadline:
            if is_available():
                utils.info_print("AI server started.")
                return True
            # The router dying instantly (port already held, broken install)
            # would otherwise burn the whole timeout with the lock held and
            # report only a generic "did not start".
            code = proc.poll()
            if code is not None:
                _server_proc = None
                utils.warning_print(
                    f"AI server exited immediately (code {code}) — "
                    "is the port already in use?"
                )
                return False
            time.sleep(_START_POLL_INTERVAL)

    utils.warning_print("AI server did not start within timeout.")
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


def _delta_piece(chunk: dict[str, Any]) -> tuple[str, bool]:
    """Extract (content piece, finish seen) from one SSE chat chunk."""
    choices = chunk.get("choices")
    if not isinstance(choices, list) or not choices:
        return "", False
    choice = choices[0]
    if not isinstance(choice, dict):
        return "", False
    delta = choice.get("delta")
    piece = delta.get("content") or "" if isinstance(delta, dict) else ""
    return piece, choice.get("finish_reason") is not None


@profiling.timed("llm.generate")
def _do_generate(
    body: dict[str, Any],
    cancel_event: threading.Event | None = None,
    on_token: Callable[[str], None] | None = None,
) -> str | None:
    """Execute a single streaming /v1/chat/completions request.

    Reads the SSE stream line-by-line, accumulating ``delta.content`` pieces
    until a chunk carries ``finish_reason`` (the completion marker — an EOF
    without one means the stream was truncated, e.g. by a mid-flight unload,
    and the partial text is discarded). When *cancel_event* is set, a watcher
    thread shuts down the underlying socket so the blocked readline()
    unblocks and the function returns ``None`` promptly — freeing the server
    for another run.

    When *on_token* is provided it is called with each content piece as it
    streams in, letting callers surface partial text live. A raising callback
    is swallowed so it can never break the read loop; the accumulated return
    value is unchanged.
    """
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        f"{config.LLM_BASE_URL}/v1/chat/completions",
        data=data,
        headers={"Content-Type": "application/json"},
    )
    parts: list[str] = []
    done_event = threading.Event()
    watcher: threading.Thread | None = None
    deadline = time.monotonic() + _GENERATE_DEADLINE
    deadline_exceeded = False
    saw_finish = False
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
                target=_watch, daemon=True, name="llm-cancel-watcher"
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
            except (OSError, ValueError, AttributeError, http.client.HTTPException):
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
            payload = line.strip()
            if not payload.startswith(b"data:"):
                continue  # skip SSE comments / blank keep-alives
            payload = payload[5:].strip()
            if payload == b"[DONE]":
                break
            try:
                chunk = json.loads(payload.decode("utf-8"))
            except (json.JSONDecodeError, UnicodeDecodeError):
                continue  # skip malformed lines, keep reading
            if not isinstance(chunk, dict):
                continue
            error = chunk.get("error")
            if error:
                utils.warning_print(f"AI generate failed: {error}")
                return None
            piece, finished = _delta_piece(chunk)
            if piece:
                parts.append(piece)
                if on_token is not None:
                    try:
                        on_token(piece)
                    except Exception as exc:
                        utils.verbose_print(f"LLM on_token callback failed: {exc}")
            if finished:
                saw_finish = True
    finally:
        done_event.set()
        if resp is not None:
            try:
                resp.close()
            except OSError:
                pass  # already-dead socket; nothing to recover
        if watcher is not None:
            watcher.join(timeout=0.5)
        if deadline_exceeded:
            utils.warning_print(
                f"AI generate exceeded {_GENERATE_DEADLINE}s deadline "
                f"(model: {body.get('model')}); aborting."
            )

    if cancel_event is not None and cancel_event.is_set():
        return None
    if not saw_finish:
        # The stream ended (EOF) without a finish_reason — the server
        # restarted, the model was unloaded mid-run, or the connection
        # dropped. The accumulated text is a truncated prefix; returning it
        # would commit a half-written result as a finished one.
        utils.warning_print(
            f"AI stream ended before completion (model: {body.get('model')})"
        )
        return None
    text = "".join(parts).strip()
    if not text:
        utils.warning_print(f"AI returned empty response (model: {body.get('model')})")
        return None
    return text


def _generate_with_load_retry(
    body: dict[str, Any],
    cancel_event: threading.Event | None,
    on_token: Callable[[str], None] | None,
) -> str | None:
    """Run ``_do_generate``, absorbing 503s while the model instance loads."""
    retry_deadline = time.monotonic() + _LOAD_RETRY_WINDOW
    while True:
        try:
            return _do_generate(body, cancel_event=cancel_event, on_token=on_token)
        except urllib.error.HTTPError as exc:
            if exc.code != 503 or time.monotonic() >= retry_deadline:
                raise
            if cancel_event is not None and cancel_event.is_set():
                return None
            time.sleep(_LOAD_RETRY_INTERVAL)


def generate(
    prompt: str,
    *,
    model: str | None = None,
    system: str | None = None,
    cancel_event: threading.Event | None = None,
    on_token: Callable[[str], None] | None = None,
) -> str | None:
    """Send a prompt to the local LLM and return the generated text.

    Uses ``/v1/chat/completions`` — never the raw ``/completion`` endpoint —
    so the model's own chat template from the GGUF metadata is applied; raw
    completion would feed an instruct model untemplated text and silently
    degrade output quality. On connection refused, automatically attempts to
    start the server and retries the request once.

    Args:
        prompt: The user prompt to send.
        model: Model value (HF ref or stem). Defaults to config.LLM_SUMMARY_MODEL.
        system: Optional system prompt.
        cancel_event: Optional event used to abort the in-flight request. When
            set, the underlying HTTP response is closed and ``None`` is
            returned promptly so the model is freed for another run.
        on_token: Optional callback invoked with each streamed content piece,
            for surfacing partial text live. Errors from it are swallowed.

    Returns the generated text string, or None on any failure or cancellation.
    """
    resolved_model = model or config.LLM_SUMMARY_MODEL
    messages: list[dict[str, str]] = []
    if system:
        messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": prompt})
    body: dict[str, Any] = {
        "model": model_name(resolved_model),
        "messages": messages,
        "stream": True,
        # Honored by templates with a think toggle (Qwen); ignored elsewhere.
        # The default reasoning_format already keeps chain-of-thought out of
        # delta.content, and _strip_think() upstream is the final backstop.
        "chat_template_kwargs": {"enable_thinking": False},
    }

    try:
        return _generate_with_load_retry(body, cancel_event, on_token)
    except urllib.error.HTTPError as exc:
        utils.warning_print(f"AI generate failed (HTTP {exc.code}): {exc.reason}")
        return None
    except (urllib.error.URLError, OSError, http.client.HTTPException) as exc:
        if cancel_event is not None and cancel_event.is_set():
            return None
        if not _is_connection_refused(exc):
            utils.warning_print(f"AI generate failed (connection): {exc}")
            return None
        # Connection refused — try to start the server and retry once
        if not start_server():
            return None
        try:
            return _generate_with_load_retry(body, cancel_event, on_token)
        except (
            urllib.error.URLError,
            urllib.error.HTTPError,
            OSError,
            http.client.HTTPException,
        ) as retry_exc:
            if cancel_event is not None and cancel_event.is_set():
                return None
            utils.warning_print(f"AI generate failed after retry: {retry_exc}")
            return None
        except (json.JSONDecodeError, KeyError, ValueError) as retry_exc:
            if cancel_event is not None and cancel_event.is_set():
                return None
            utils.warning_print(f"AI generate failed (response): {retry_exc}")
            return None
    except (json.JSONDecodeError, KeyError, ValueError) as exc:
        if cancel_event is not None and cancel_event.is_set():
            return None
        utils.warning_print(f"AI generate failed (response): {exc}")
        return None


def _split_hf_ref(ref: str) -> tuple[str, str]:
    """Split ``user/repo[:QUANT]`` into (repo, quant). Quant defaults to Q4_K_M."""
    if ":" in ref:
        repo, quant = ref.rsplit(":", 1)
        return repo, quant or "Q4_K_M"
    return ref, "Q4_K_M"


def _resolve_hf_file(ref: str) -> dict[str, Any] | None:
    """Find the single GGUF file for *ref* in its Hugging Face repo.

    Lists the repo tree and matches the requested quant against GGUF
    filenames (case-insensitive). Returns ``{path, size, sha256}`` (sha256
    may be empty when the API omits the LFS oid), or None with a warning on
    any failure — repo missing, gated (HTTP 401/403), quant absent, or the
    match being a sharded multi-file model, which clipgen does not support.
    """
    repo, quant = _split_hf_ref(ref)
    url = f"https://huggingface.co/api/models/{repo}/tree/main?recursive=true"
    req = urllib.request.Request(url, headers={"User-Agent": "clipgen"})
    try:
        with urllib.request.urlopen(req, timeout=_HF_API_TIMEOUT) as resp:
            tree = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        if exc.code in (401, 403):
            utils.warning_print(f"Model repo {repo} is gated (HTTP {exc.code}).")
        elif exc.code == 404:
            utils.warning_print(f"Model repo {repo} not found.")
        else:
            utils.warning_print(f"Model lookup failed (HTTP {exc.code}): {repo}")
        return None
    except (
        urllib.error.URLError,
        OSError,
        ValueError,
        http.client.HTTPException,
    ) as exc:
        utils.warning_print(f"Model lookup failed: {exc}")
        return None

    if not isinstance(tree, list):
        utils.warning_print(f"Unexpected model listing for {repo}.")
        return None
    needle = quant.lower()
    matches = [
        entry
        for entry in tree
        if isinstance(entry, dict)
        and entry.get("path", "").lower().endswith(".gguf")
        and needle in Path(entry.get("path", "")).stem.lower()
    ]
    if not matches:
        utils.warning_print(f"No {quant} GGUF found in {repo}.")
        return None
    if len(matches) > 1 or "-of-" in matches[0].get("path", "").lower():
        utils.warning_print(f"{repo}:{quant} is sharded; not supported.")
        return None
    entry = matches[0]
    lfs = entry.get("lfs") or {}
    return {
        "path": entry.get("path", ""),
        "size": int(entry.get("size") or lfs.get("size") or 0),
        "sha256": str(lfs.get("oid") or ""),
    }


def download_model(
    ref: str,
    on_progress: Callable[[dict[str, Any]], None] | None = None,
) -> bool:
    """Download the GGUF for a Hugging Face ref into the models dir.

    Consent lives with the caller — this is only ever reached after the user
    confirms the download dialog. Streams to a temp file beside the target,
    verifying the SHA256 from the repo's LFS metadata incrementally (size-only
    when the API omits it), then renames into place under ``model_file(ref)``.
    Progress dicts are shaped like the transcribe-model download chunks
    (``status``/``completed``/``total``). Never raises; returns False on any
    failure. Idempotent: an already-downloaded model returns True immediately.
    """
    ref = (ref or "").strip()
    if not ref:
        return False
    if "/" not in ref:
        utils.warning_print(f"Not a downloadable model ref: {ref}")
        return False
    target = model_file(ref)
    if target.is_file():
        return True

    resolved = _resolve_hf_file(ref)
    if resolved is None:
        return False
    repo, _ = _split_hf_ref(ref)
    url = (
        f"https://huggingface.co/{repo}/resolve/main/"
        f"{urllib.parse.quote(resolved['path'])}"
    )

    directory = models_dir()
    try:
        directory.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        utils.warning_print(f"Model download failed (create dir): {exc}")
        return False
    _sweep_stale_downloads(directory)

    download_path: Path | None = None
    digest = hashlib.sha256()
    received = 0
    request = urllib.request.Request(url, headers={"User-Agent": "clipgen"})
    deadline = time.monotonic() + _DOWNLOAD_DEADLINE
    try:
        with tempfile.NamedTemporaryFile(
            dir=directory, suffix=".part", delete=False
        ) as tmp:
            download_path = Path(tmp.name)
        with (
            urllib.request.urlopen(request, timeout=_DOWNLOAD_TIMEOUT) as response,
            download_path.open("wb") as out,
        ):
            total = int(response.headers.get("Content-Length") or resolved["size"])
            while chunk := response.read(_DOWNLOAD_CHUNK):
                if time.monotonic() > deadline:
                    raise TimeoutError(
                        f"download exceeded the {_DOWNLOAD_DEADLINE}s deadline "
                        f"after {received} of {total} bytes"
                    )
                digest.update(chunk)
                out.write(chunk)
                received += len(chunk)
                if on_progress is not None:
                    on_progress(
                        {
                            "status": "downloading model",
                            "completed": received,
                            "total": total,
                        }
                    )
    except (urllib.error.URLError, OSError, ValueError) as exc:
        utils.warning_print(f"Model download failed: {exc}")
        if download_path is not None:
            download_path.unlink(missing_ok=True)
        return False

    expected = resolved["sha256"]
    if expected and digest.hexdigest() != expected:
        utils.warning_print(f"Model download corrupt (SHA256 mismatch): {ref}")
        download_path.unlink(missing_ok=True)
        return False
    if not expected and resolved["size"] and received != resolved["size"]:
        utils.warning_print(f"Model download incomplete: {ref}")
        download_path.unlink(missing_ok=True)
        return False

    try:
        download_path.replace(target)
    except OSError as exc:
        utils.warning_print(f"Model download failed (rename): {exc}")
        download_path.unlink(missing_ok=True)
        return False
    utils.info_print(f"Downloaded model {ref}.")
    return True


def _sweep_stale_downloads(directory: Path) -> None:
    """Remove partial ``.part`` temp files left by an interrupted download.

    The download runs on a daemon thread, which is killed at interpreter
    shutdown *without* unwinding, so quitting clipgen mid-download leaves the
    partial file behind — at ~6 GB per model that adds up fast. Best-effort
    by design; a sweep failure must not block the download that follows.
    """
    for stale in directory.glob("tmp*.part"):
        try:
            if stale.is_file():
                stale.unlink()
        except OSError:
            pass
