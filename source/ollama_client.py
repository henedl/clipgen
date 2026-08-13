"""Ollama local LLM transport for clipgen.

A thin, reusable HTTP wrapper around the Ollama REST API. All functions fail
gracefully (return None / False) and never raise on network errors. On
connection refused, ``generate()`` auto-starts ``ollama serve`` and retries
once.

This module is intentionally small — higher-level reasoning lives in
[thinking_agents.py](thinking_agents.py), which routes every call through
``generate()`` here.

Key functions:
  is_available()      - check Ollama server connectivity
  is_installed()      - check whether the `ollama` binary is reachable at all
  resolve_ollama_bin()- the binary actually used: PATH first, managed second
  start_server()      - spawn `ollama serve` and wait for it to answer
  install_guidance_lines() - platform-specific "how to install Ollama" text
  can_install_managed() / install_managed() - consent-gated in-app download of
    the official standalone CLI into clipgen's config dir (macOS only)
  list_models()       - enumerate installed models with metadata
  is_model_installed()- check whether a specific model is installed locally
  generate()          - send a prompt and get a text response
  pull_model()        - download (install) a model, streaming progress
  unload_model()      - ask Ollama to evict a model from memory immediately
"""

import hashlib
import json
import os
import shutil
import socket
import subprocess
import sys
import tarfile
import tempfile
import threading
import time
import urllib.error
import urllib.request
from collections.abc import Callable
from pathlib import Path
from typing import Any

import config
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
_START_POLL_INTERVAL = 0.5  # seconds between health-check polls after starting server
_START_TIMEOUT = 10  # seconds to wait for server to become available after starting
# Per-read timeout for a model pull. Layers stream steadily once a download
# starts; this bounds a stalled connection without aborting a healthy pull.
_PULL_TIMEOUT = 300  # seconds
_CANCEL_WATCHER_POLL = 1.0  # seconds; bounds abort latency during long quiet stretches

# Serializes start_server() calls so two threads hitting connection-refused at
# the same time don't both spawn `ollama serve`.
_start_server_lock = threading.Lock()

# Pinned release assets for the consent-gated in-app install. Both SHA256s
# come from the release's published sha256sum.txt and are re-verified after
# download. macOS gets the standalone CLI tarball (flat — the `ollama` binary
# with llama-server and the GGML runner libraries beside it, designed to run
# from any directory). Windows gets the official OllamaSetup.exe run silently:
# the standalone zip there is ~1 GB of GPU runner DLLs with no PATH handling,
# while the installer is a per-user Inno Setup (no UAC) that also manages
# updates — at the cost of a much larger download, which the consent dialog
# states honestly via managed_install_size_mb().
OLLAMA_DOWNLOAD_VERSION = "0.32.5"
_OLLAMA_DARWIN_URL = (
    "https://github.com/ollama/ollama/releases/download/"
    f"v{OLLAMA_DOWNLOAD_VERSION}/ollama-darwin.tgz"
)
_OLLAMA_DARWIN_SHA256 = (
    "5789dd037a86adb328c72c11fc45e6c558452d07e5b50814a8bdb7b0fbdbcd81"
)
_OLLAMA_DARWIN_SIZE_BYTES = 145_747_028  # fallback when Content-Length is absent
_OLLAMA_WINDOWS_URL = (
    "https://github.com/ollama/ollama/releases/download/"
    f"v{OLLAMA_DOWNLOAD_VERSION}/OllamaSetup.exe"
)
_OLLAMA_WINDOWS_SHA256 = (
    "b7eeef038ddcbd09ac665b11872baff1bc9b42794be41b5ef187b2f4b16a4498"
)
_OLLAMA_WINDOWS_SIZE_BYTES = 1_563_078_600
_INSTALL_CHUNK = 1024 * 1024
_INSTALL_TIMEOUT = 120  # seconds; connect + first byte, GitHub is normally fast
_INSTALLER_RUN_TIMEOUT = 900  # seconds; silent Inno Setup unpacks ~2 GB of runners
_VERSION_PROBE_TIMEOUT = 15  # seconds; `ollama --version` sanity check
# 0 off-Windows, so passing it unconditionally leaves darwin/linux unchanged.
_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


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


def is_model_installed(
    model: str, installed: list[dict[str, Any]] | None = None
) -> bool:
    """Return True if *model* is among the locally installed Ollama models.

    Matches the requested name against installed tags: exact match,
    ``model:latest`` (Ollama's implicit tag), or any tag sharing the same base
    when *model* carries no explicit ``:tag``. Returns False when the server is
    unreachable or the model is absent. Pass *installed* (the result of
    ``list_models()``) to avoid a redundant ``/api/tags`` round-trip when the
    caller already holds the list.
    """
    if not model:
        return False
    if installed is None:
        installed = list_models()
    if not installed:
        return False
    names = {m["name"] for m in installed}
    if model in names:
        return True
    if ":" not in model:
        if f"{model}:latest" in names:
            return True
        prefix = f"{model}:"
        return any(name.startswith(prefix) for name in names)
    return False


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
    return isinstance(exc, ConnectionRefusedError)


def install_guidance_lines() -> list[str]:
    """Return actionable install guidance based on the current platform.

    Public because the guidance now has to reach the browser as well as the
    terminal — ``/api/models`` ships it so the Transcripts and Overview pages
    can tell a user *how* to get Ollama instead of only that it is missing.
    """
    return utils.install_guidance_lines(
        brew_command="brew install ollama",
        linux=["Linux: curl -fsSL https://ollama.com/install.sh | sh"],
        windows=["Windows: winget install Ollama.Ollama"],
        download_url="https://ollama.com/download",
        verify_commands=["ollama --version"],
    )


def _managed_ollama_dir() -> Path:
    """Where the in-app install lives, under clipgen's own config dir."""
    return start_settings.config_dir() / "tools" / "ollama"


def _windows_install_path() -> Path:
    """Where OllamaSetup.exe installs per-user: %LOCALAPPDATA%\\Programs\\Ollama."""
    local_appdata = os.environ.get(
        "LOCALAPPDATA", str(Path.home() / "AppData" / "Local")
    )
    return Path(local_appdata) / "Programs" / "Ollama" / "ollama.exe"


def managed_ollama_path() -> Path | None:
    """Path to the managed ``ollama`` binary, or None when absent.

    On macOS, probes both the flat layout the current tarball ships (``ollama``
    at the root, runners beside it) and a ``bin/`` layout in case a future
    release restructures the archive. On Windows, probes the official
    installer's per-user location directly: OllamaSetup.exe writes its PATH
    entry to the registry, which the already-running process never sees, so
    ``shutil.which`` cannot find a fresh install (nor one the user ran
    themselves — probing here makes both discoverable without a restart).
    """
    if sys.platform == "win32":
        exe = _windows_install_path()
        # No os.X_OK probe: it is meaningless for .exe files on Windows.
        return exe if exe.is_file() else None
    base = _managed_ollama_dir()
    for candidate in (base / "ollama", base / "bin" / "ollama"):
        if candidate.is_file() and os.access(candidate, os.X_OK):
            return candidate
    return None


def resolve_ollama_bin() -> str | None:
    """The ``ollama`` binary clipgen should run, or None when there is none.

    PATH wins over the managed copy: a user-managed install (Ollama.app, brew)
    self-updates and carries the menu-bar app, so clipgen's downloaded copy is
    strictly the fallback for machines with nothing — matching the codebase's
    stance that PATH augmentation is about discoverability, not overriding a
    resolution order the user already has.
    """
    on_path = shutil.which("ollama")
    if on_path:
        return on_path
    managed = managed_ollama_path()
    return str(managed) if managed is not None else None


def is_installed() -> bool:
    """Return True when an ``ollama`` binary is reachable (PATH or managed).

    Deliberately distinct from ``is_available()``, which reports whether the
    *server* answers. The two states need opposite advice — "start it, then
    refresh" is useless to someone who never installed it — so every surface
    that gates on Ollama reads both rather than collapsing them into one flag.
    """
    return resolve_ollama_bin() is not None


def start_server() -> bool:
    """Attempt to start ``ollama serve`` and wait for it to become available.

    Serialized via ``_start_server_lock`` so concurrent connection-refused
    retries don't both spawn a server process — the first one to acquire the
    lock spawns it, the rest see the already-running server when they re-poll.

    Returns True if the server is responding after startup, False otherwise.
    """
    if not is_installed():
        utils.warning_print(
            "Ollama is not installed.",
            details=install_guidance_lines(),
        )
        return False

    with _start_server_lock:
        # Another thread may have already started the server while we waited.
        if is_available():
            return True

        utils.info_print("Starting Ollama server...")
        try:
            subprocess.Popen(
                [resolve_ollama_bin() or "ollama", "serve"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=_NO_WINDOW,
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


def can_install_managed() -> bool:
    """Whether the in-app Ollama download is supported on this platform."""
    return sys.platform in ("darwin", "win32")


def managed_install_size_mb() -> int:
    """Approximate download size, for consent labels ("~140 MB" / "~1.5 GB")."""
    if sys.platform == "win32":
        return round(_OLLAMA_WINDOWS_SIZE_BYTES / 1e6)
    return round(_OLLAMA_DARWIN_SIZE_BYTES / 1e6)


def _managed_binary_works(binary: Path) -> bool:
    """Sanity-run ``ollama --version`` on the managed copy."""
    try:
        probe = subprocess.run(
            [str(binary), "--version"],
            capture_output=True,
            timeout=_VERSION_PROBE_TIMEOUT,
            check=False,
            creationflags=_NO_WINDOW,
        )
    except (OSError, subprocess.TimeoutExpired):
        return False
    return probe.returncode == 0


def _download_pinned(
    url: str,
    expected_sha256: str,
    size_fallback: int,
    suffix: str,
    target_dir: Path,
    on_progress: Callable[[dict[str, Any]], None] | None,
) -> Path | None:
    """Stream a pinned release asset to a temp file in ``target_dir``.

    Verifies the SHA256 incrementally; on any failure (network, disk, hash
    mismatch) the temp file is removed and None is returned. On success the
    caller owns the returned path and must unlink it when done.
    """
    with tempfile.NamedTemporaryFile(
        dir=target_dir, suffix=suffix, delete=False
    ) as tmp:
        download_path = Path(tmp.name)
    digest = hashlib.sha256()
    received = 0
    request = urllib.request.Request(
        url, headers={"User-Agent": "clipgen-ollama-install"}
    )
    try:
        with (
            urllib.request.urlopen(request, timeout=_INSTALL_TIMEOUT) as response,
            download_path.open("wb") as out,
        ):
            total = int(response.headers.get("Content-Length") or size_fallback)
            while chunk := response.read(_INSTALL_CHUNK):
                digest.update(chunk)
                out.write(chunk)
                received += len(chunk)
                if on_progress is not None:
                    on_progress(
                        {
                            "status": "downloading Ollama",
                            "completed": received,
                            "total": total,
                        }
                    )
    except (urllib.error.URLError, OSError, ValueError) as exc:
        utils.warning_print(f"Ollama install failed (download): {exc}")
        download_path.unlink(missing_ok=True)
        return None

    if digest.hexdigest() != expected_sha256:
        utils.warning_print(
            "Ollama install failed: downloaded file does not match the "
            f"pinned SHA256 for v{OLLAMA_DOWNLOAD_VERSION}."
        )
        download_path.unlink(missing_ok=True)
        return None
    return download_path


def _extract_darwin_archive(archive_path: Path, target_dir: Path) -> bool:
    """Unpack the verified darwin tarball into the managed dir."""
    try:
        with tarfile.open(archive_path) as archive:
            # filter="tar" blocks absolute-path/traversal escapes while
            # keeping file modes — this hash-verified archive ships
            # executables, which the stricter "data" filter would strip.
            archive.extractall(target_dir, filter="tar")
    except (tarfile.TarError, OSError) as exc:
        utils.warning_print(f"Ollama install failed (extract): {exc}")
        return False
    return True


def _run_windows_installer(
    setup_path: Path,
    on_progress: Callable[[dict[str, Any]], None] | None,
) -> bool:
    """Run the verified OllamaSetup.exe silently and wait for it to finish."""
    if on_progress is not None:
        # No completed/total — the frontend renders total-less statuses as
        # plain text, same as "unpacking Ollama" on macOS.
        on_progress({"status": "running installer"})
    try:
        proc = subprocess.run(
            [str(setup_path), "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART"],
            capture_output=True,
            timeout=_INSTALLER_RUN_TIMEOUT,
            check=False,
            creationflags=_NO_WINDOW,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        utils.warning_print(f"Ollama install failed (installer): {exc}")
        return False
    if proc.returncode != 0:
        utils.warning_print(
            f"Ollama install failed: installer exited with code {proc.returncode}."
        )
        return False
    return True


def install_managed(
    on_progress: Callable[[dict[str, Any]], None] | None = None,
) -> bool:
    """Download and install the pinned Ollama release for this platform.

    Consent lives with the caller — this is only ever reached from an explicit
    user action (the Transcripts install dialog). On macOS, streams the pinned
    CLI tarball to a temp file inside the managed dir, verifies its SHA256,
    extracts, and sanity-runs ``--version``. On Windows, streams the official
    OllamaSetup.exe the same way, then runs it silently — a per-user Inno Setup
    install to %LOCALAPPDATA%\\Programs\\Ollama with no UAC prompt. Its
    ``skipifsilent`` post-install entries mean the tray app does not auto-start;
    that's fine, the frontend POSTs ``api/models/ollama/start`` right after.
    Progress dicts are shaped like ``pull_model`` chunks
    (``status``/``completed``/``total``) so the frontend reuses its pull
    rendering. Never raises; returns False on any failure, matching the
    module's error style. Idempotent: an already-working install returns True
    immediately.
    """
    if not can_install_managed():
        return False
    existing = managed_ollama_path()
    if existing is not None and _managed_binary_works(existing):
        return True

    target_dir = _managed_ollama_dir()
    try:
        target_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        utils.warning_print(f"Ollama install failed (create dir): {exc}")
        return False

    on_windows = sys.platform == "win32"
    download_path = _download_pinned(
        url=_OLLAMA_WINDOWS_URL if on_windows else _OLLAMA_DARWIN_URL,
        expected_sha256=_OLLAMA_WINDOWS_SHA256 if on_windows else _OLLAMA_DARWIN_SHA256,
        size_fallback=_OLLAMA_WINDOWS_SIZE_BYTES
        if on_windows
        else _OLLAMA_DARWIN_SIZE_BYTES,
        suffix=".exe" if on_windows else ".tgz",
        target_dir=target_dir,
        on_progress=on_progress,
    )
    if download_path is None:
        return False
    try:
        if on_windows:
            if not _run_windows_installer(download_path, on_progress):
                return False
        else:
            if on_progress is not None:
                on_progress({"status": "unpacking Ollama"})
            if not _extract_darwin_archive(download_path, target_dir):
                return False
    finally:
        download_path.unlink(missing_ok=True)

    installed = managed_ollama_path()
    if installed is None and not on_windows:
        # Belt and braces for a future archive that drops the exec bit.
        fallback = target_dir / "ollama"
        if fallback.is_file():
            try:
                fallback.chmod(0o755)
            except OSError:
                pass
            installed = managed_ollama_path()
    if installed is None or not _managed_binary_works(installed):
        utils.warning_print(
            "Ollama install failed: installed binary is missing or does not run."
        )
        return False
    if on_progress is not None:
        on_progress({"status": "success"})
    destination = installed.parent if on_windows else target_dir
    utils.info_print(f"Installed Ollama v{OLLAMA_DOWNLOAD_VERSION} to {destination}.")
    return True


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
    on_token: Callable[[str], None] | None = None,
) -> str | None:
    """Execute a single Ollama /api/generate request (streaming).

    Reads the NDJSON stream chunk-by-chunk, accumulating ``response`` fields
    until a chunk arrives with ``done: true``. When *cancel_event* is set, a
    watcher thread shuts down the underlying socket so the blocked readline()
    unblocks and the function returns ``None`` promptly — freeing Ollama for
    another run.

    When *on_token* is provided it is called with each ``response`` piece as it
    streams in, letting callers surface partial text live. A raising callback is
    swallowed so it can never break the read loop; the accumulated return value
    is unchanged.
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
                if on_token is not None:
                    try:
                        on_token(piece)
                    except Exception as exc:
                        utils.verbose_print(f"Ollama on_token callback failed: {exc}")
            if chunk.get("done"):
                break
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
                f"Ollama generate exceeded {_GENERATE_DEADLINE}s deadline "
                f"(model: {body.get('model')}); aborting."
            )

    if cancel_event is not None and cancel_event.is_set():
        return None
    text = "".join(parts).strip()
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
    on_token: Callable[[str], None] | None = None,
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
        on_token: Optional callback invoked with each streamed response piece,
            for surfacing partial text live. Errors from it are swallowed.

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
        return _do_generate(body, cancel_event=cancel_event, on_token=on_token)
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
        if not start_server():
            return None
        try:
            return _do_generate(body, cancel_event=cancel_event, on_token=on_token)
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


def _do_pull(model: str, on_progress: Callable[[dict[str, Any]], None] | None) -> bool:
    """Stream POST /api/pull for *model*, returning True on a ``success`` line.

    Reads the NDJSON progress stream line-by-line, forwarding each status dict
    to *on_progress*. An ``error`` field anywhere in the stream aborts with
    False.
    """
    data = json.dumps({"model": model, "stream": True}).encode("utf-8")
    req = urllib.request.Request(
        f"{config.OLLAMA_BASE_URL}/api/pull",
        data=data,
        headers={"Content-Type": "application/json"},
    )
    succeeded = False
    resp: Any = None
    try:
        resp = urllib.request.urlopen(req, timeout=_PULL_TIMEOUT)
        while True:
            line = resp.readline()
            if not line:
                break
            try:
                chunk = json.loads(line.decode("utf-8"))
            except (json.JSONDecodeError, UnicodeDecodeError):
                continue  # skip malformed lines, keep reading
            if not isinstance(chunk, dict):
                continue
            err = chunk.get("error")
            if err:
                utils.warning_print(f"Ollama pull failed: {err}")
                return False
            if on_progress is not None:
                on_progress(chunk)
            if chunk.get("status") == "success":
                succeeded = True
    finally:
        if resp is not None:
            try:
                resp.close()
            except OSError:
                pass  # already-dead socket; nothing to recover
    return succeeded


def pull_model(
    model: str,
    on_progress: Callable[[dict[str, Any]], None] | None = None,
) -> bool:
    """Download (install) an Ollama model via streaming POST /api/pull.

    Reports progress through *on_progress*, called once per NDJSON status line
    (keys include ``status``, ``total``, ``completed``). On connection refused,
    auto-starts ``ollama serve`` and retries once. Returns True only when the
    stream ends with a ``success`` status; False on any failure.
    """
    if not model:
        return False
    try:
        return _do_pull(model, on_progress)
    except urllib.error.HTTPError as exc:
        utils.warning_print(f"Ollama pull failed (HTTP {exc.code}): {exc.reason}")
        return False
    except (urllib.error.URLError, OSError) as exc:
        if not _is_connection_refused(exc):
            utils.warning_print(f"Ollama pull failed (connection): {exc}")
            return False
        # Connection refused — try to start the server and retry once.
        if not start_server():
            return False
        try:
            return _do_pull(model, on_progress)
        except (urllib.error.URLError, urllib.error.HTTPError, OSError) as retry_exc:
            utils.warning_print(f"Ollama pull failed after retry: {retry_exc}")
            return False
