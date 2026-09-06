"""In-app self-update for the frozen desktop builds.

Checks GitHub Releases for a newer tag, downloads the asset that matches this
install, and applies it in place with a relaunch. Only a frozen desktop launch
is supported; every other run (source tree, browser mode) reports
``supported: False`` and the frontend shows nothing.

Three install shapes, each with its own apply path:

* ``mac-app`` — ``…/clipgen.app``. Mount the DMG, ``ditto`` the bundle to a
  staging sibling on the same volume, then a detached shell helper waits for
  this process to exit, swaps the bundles and ``open``s the new one. urllib
  downloads carry no quarantine flag, so Gatekeeper's "damaged" dialog never
  fires. Measured 2026-09-06 on Sequoia with a local ad-hoc build: the swap
  succeeded in ``/Applications``, ``~/Applications`` and a plain folder, with
  no App Management prompt. Should a future macOS refuse the rename, the
  helper leaves the staged copy and logs the failure, which the next launch
  reports as ``last_error``.
* ``win-inno`` — the per-user installer (``unins000.exe`` beside the exe).
  Run the new ``setup.exe`` silently with ``/RELAUNCH=1``; the ``[Run]`` entry
  in ``build/clipgen.iss`` relaunches the app once files are in place.
* ``win-zip`` — the portable folder. Extract beside it, then a detached
  PowerShell helper renames old/new and relaunches.

State machine (``_status["phase"]``): ``idle`` → ``checking`` → ``available``
→ ``downloading`` → ``ready`` → ``applying``, with ``error`` reachable from any
step. The automatic launch check honours a cooldown persisted in the config
dir's ``update.json``; a manual check bypasses it.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
import zipfile
from collections.abc import Callable
from pathlib import Path, PureWindowsPath
from typing import Any, Literal

import config
import start_settings
import utils

Shape = Literal["mac-app", "win-inno", "win-zip", "unsupported"]

RELEASES_API = (
    "https://api.github.com/repos/"
    + config.REPO_URL.removeprefix("https://github.com/").strip("/")
    + "/releases/latest"
)
CHECK_COOLDOWN_SECONDS = 6 * 3600
STATE_FILENAME = "update.json"
UPDATES_DIRNAME = "updates"
APPLY_LOG = "apply.log"
# Seconds the helper waits for this process to exit before giving up.
HELPER_WAIT_SECONDS = 60

_REQUEST_TIMEOUT = 10.0
_DOWNLOAD_TIMEOUT = 30.0
_DOWNLOAD_CHUNK = 1 << 20
_DOWNLOAD_DEADLINE = 3600.0
_VERSION_RE = re.compile(r"^v?(\d+)\.(\d+)\.(\d+)$")

_lock = threading.Lock()
_status: dict[str, Any] = {
    "phase": "idle",
    "supported": False,
    "checked": False,
    "current": None,
    "version": None,
    "release_url": None,
    "asset": None,
    "path": None,
    "completed": 0,
    "total": 0,
    "error": None,
    "last_error": None,
    "skipped": None,
}
# The release the current phase refers to; None until a check succeeds.
_latest: dict[str, Any] | None = None


# ---- Install shape -----------------------------------------------------------


def install_shape() -> Shape:
    """Classify this process's install; mirrors cli.get_runtime_working_dir."""
    if not getattr(sys, "frozen", False):
        return "unsupported"
    exe_dir = Path(sys.executable).resolve().parent
    if exe_dir.name == "MacOS" and exe_dir.parent.name == "Contents":
        if exe_dir.parent.parent.suffix == ".app":
            return "mac-app"
        return "unsupported"
    meipass = getattr(sys, "_MEIPASS", None)
    if (
        sys.platform == "win32"
        and meipass
        and Path(meipass).resolve().parent == exe_dir
    ):
        if (exe_dir / "unins000.exe").is_file():
            return "win-inno"
        return "win-zip"
    return "unsupported"


def install_root() -> Path | None:
    """The .app bundle or the folder holding clipgen.exe; None if unsupported."""
    shape = install_shape()
    exe = Path(sys.executable).resolve()
    if shape == "mac-app":
        return exe.parent.parent.parent
    if shape in ("win-inno", "win-zip"):
        return exe.parent
    return None


def is_supported() -> bool:
    """Only a frozen desktop window can be replaced and relaunched."""
    return utils.GUI_LAUNCH and install_shape() != "unsupported"


# ---- Versions and assets -----------------------------------------------------


def parse_version(text: str) -> tuple[int, int, int] | None:
    """``v1.2.3`` or ``1.2.3`` → (1, 2, 3); anything else (dev builds) → None."""
    match = _VERSION_RE.match((text or "").strip())
    if match is None:
        return None
    return int(match.group(1)), int(match.group(2)), int(match.group(3))


def is_newer(latest: str, current: str) -> bool:
    """True when both parse and *latest* is strictly greater."""
    new = parse_version(latest)
    cur = parse_version(current)
    return new is not None and cur is not None and new > cur


def asset_name(tag: str, shape: Shape) -> str | None:
    """The release asset build-binaries.yml publishes for *shape*."""
    suffix = {
        "mac-app": "macos.dmg",
        "win-inno": "setup.exe",
        "win-zip": "windows.zip",
    }.get(shape)
    if suffix is None:
        return None
    return f"clipgen-{tag}-{suffix}"


def pick_asset(release: dict[str, Any], shape: Shape) -> dict[str, Any] | None:
    """Exact-name match against the release's assets."""
    wanted = asset_name(str(release.get("tag") or ""), shape)
    if wanted is None:
        return None
    for asset in release.get("assets") or []:
        if asset.get("name") == wanted:
            return asset
    return None


def _normalize_release(payload: dict[str, Any]) -> dict[str, Any]:
    assets = []
    for raw in payload.get("assets") or []:
        digest = str(raw.get("digest") or "")
        assets.append(
            {
                "name": str(raw.get("name") or ""),
                "size": int(raw.get("size") or 0),
                "sha256": digest.removeprefix("sha256:") if digest else "",
                "url": str(raw.get("browser_download_url") or ""),
            }
        )
    return {
        "tag": str(payload.get("tag_name") or ""),
        "url": str(payload.get("html_url") or config.REPO_URL + "/releases"),
        "assets": assets,
    }


# ---- Check -------------------------------------------------------------------


def check_latest(*, force: bool = False) -> dict[str, Any] | None:
    """The latest release, or None when skipped, offline, or unparsable.

    Not forced: skipped when the launch check is disabled, and answered from
    ``update.json`` inside the cooldown. GitHub allows 60 unauthenticated
    requests an hour per address; an ETag revalidation costs none of them.
    """
    state = start_settings.load_config_json(STATE_FILENAME, default={}) or {}
    cached = state.get("latest") if isinstance(state.get("latest"), dict) else None
    now = time.time()
    if not force:
        if not config.UPDATE_CHECK_ON_LAUNCH:
            return None
        last = float(state.get("last_check") or 0)
        if now - last < CHECK_COOLDOWN_SECONDS:
            return cached
    headers = {
        "User-Agent": f"clipgen/{utils.get_version()}",
        "Accept": "application/vnd.github+json",
    }
    if cached and state.get("etag"):
        headers["If-None-Match"] = str(state["etag"])
    request = urllib.request.Request(RELEASES_API, headers=headers)
    etag = state.get("etag")
    try:
        with urllib.request.urlopen(request, timeout=_REQUEST_TIMEOUT) as response:
            etag = response.headers.get("ETag") or etag
            release = _normalize_release(json.loads(response.read().decode("utf-8")))
    except urllib.error.HTTPError as exc:
        if exc.code != 304 or cached is None:
            utils.warning_print(f"Update check failed: HTTP {exc.code}")
            return None
        release = cached
    except (urllib.error.URLError, OSError, ValueError) as exc:
        utils.warning_print(f"Update check failed: {exc}")
        return None
    if not release.get("tag"):
        return None
    _save_state(last_check=now, etag=etag, latest=release)
    return release


def _save_state(**changes: Any) -> None:
    """Merge *changes* into update.json; other keys (skipped tag) survive."""
    state = start_settings.load_config_json(STATE_FILENAME, default={}) or {}
    state.update(changes)
    start_settings.save_config_json(STATE_FILENAME, state)


def _skipped_tag() -> str | None:
    state = start_settings.load_config_json(STATE_FILENAME, default={}) or {}
    tag = state.get("skipped")
    return str(tag) if tag else None


def start_check(*, force: bool = False) -> bool:
    """Enter ``checking`` synchronously so the route's snapshot already shows it."""
    if not force and not config.UPDATE_CHECK_ON_LAUNCH:
        return False
    with _lock:
        if _status["phase"] in ("checking", "downloading", "applying"):
            return False
        # A launch check (every page load) must not disturb an offered update.
        if not force and _status["phase"] in ("available", "ready"):
            return False
        _status.update(phase="checking", current=utils.get_version(), error=None)
    return True


def run_check(*, force: bool = False) -> None:
    """start_check plus finish_check in one call (tests, CLI)."""
    if start_check(force=force):
        finish_check(force=force)


def finish_check(*, force: bool = False) -> None:
    """Thread body behind /api/update/check; moves the phase on from checking."""
    global _latest
    current = utils.get_version()
    release = check_latest(force=force)
    shape = install_shape()
    skipped = _skipped_tag()
    if force and skipped:
        # A manual check means "show me anyway"; the skip is forgotten.
        _save_state(skipped=None)
        skipped = None
    with _lock:
        # A skipped or offline launch check must not read as "up to date".
        _status["checked"] = release is not None or force
        if release is None:
            _latest = None
            _status.update(version=None, release_url=None, asset=None)
            if force:
                _status.update(phase="error", error="Could not reach GitHub")
            else:
                _status["phase"] = "idle"
            return
        if not is_newer(release["tag"], current):
            _latest = None
            _status.update(phase="idle", version=None, release_url=release["url"])
            return
        if skipped == release["tag"] and not force:
            _latest = None
            _status.update(
                phase="idle", version=None, release_url=release["url"], skipped=skipped
            )
            return
        _status["skipped"] = None
        asset = pick_asset(release, shape)
        _latest = release
        # Read before the reset below: a verified download for this same asset stays ready.
        already = _status.get("path")
        same_file = (
            asset is not None
            and _status.get("asset") == asset["name"]
            and bool(already)
            and Path(str(already)).is_file()
        )
        _status.update(
            version=release["tag"], release_url=release["url"], asset=None, path=None
        )
        if asset is None:
            _status.update(phase="error", error="No download for this platform")
            return
        _status["asset"] = asset["name"]
        if same_file:
            _status.update(
                phase="ready",
                path=already,
                total=asset["size"],
                completed=asset["size"],
            )
            return
        existing = _existing_download(asset)
        if existing is not None:
            _status.update(phase="ready", path=str(existing), total=asset["size"])
            _status["completed"] = asset["size"]
        else:
            _status.update(phase="available", completed=0, total=asset["size"])


# ---- Download ----------------------------------------------------------------


def _updates_dir() -> Path:
    return start_settings.config_dir() / UPDATES_DIRNAME


def _existing_download(asset: dict[str, Any]) -> Path | None:
    """A previously downloaded asset that still verifies, else None."""
    path = _updates_dir() / asset["name"]
    if not path.is_file():
        return None
    if _verify_file(path, asset) is None:
        return path
    path.unlink(missing_ok=True)
    return None


def _verify_file(path: Path, asset: dict[str, Any]) -> str | None:
    """Size and sha256 against the release metadata; an error string or None."""
    try:
        size = path.stat().st_size
    except OSError as exc:
        return str(exc)
    if asset.get("size") and size != asset["size"]:
        return f"size mismatch ({size} of {asset['size']} bytes)"
    expected = asset.get("sha256") or ""
    if not expected:
        return None
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(_DOWNLOAD_CHUNK):
            digest.update(chunk)
    if digest.hexdigest() != expected:
        return "checksum mismatch"
    return None


def download_update(
    asset: dict[str, Any],
    on_progress: Callable[[int, int], None] | None = None,
) -> Path | None:
    """Stream *asset* into the updates dir and verify it. Never raises.

    Same shape as llm_client.download_model: a ``.part`` temp beside the
    target, streamed sha256, a wall-clock deadline, and None on any failure.
    """
    directory = _updates_dir()
    target = directory / asset["name"]
    try:
        directory.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        utils.warning_print(f"Update download failed (create dir): {exc}")
        return None
    for stale in directory.glob("*.part"):
        stale.unlink(missing_ok=True)

    part: Path | None = None
    received = 0
    request = urllib.request.Request(
        asset["url"], headers={"User-Agent": f"clipgen/{utils.get_version()}"}
    )
    deadline = time.monotonic() + _DOWNLOAD_DEADLINE
    try:
        with tempfile.NamedTemporaryFile(
            dir=directory, suffix=".part", delete=False
        ) as tmp:
            part = Path(tmp.name)
        with (
            urllib.request.urlopen(request, timeout=_DOWNLOAD_TIMEOUT) as response,
            part.open("wb") as out,
        ):
            total = int(response.headers.get("Content-Length") or asset["size"])
            while chunk := response.read(_DOWNLOAD_CHUNK):
                if time.monotonic() > deadline:
                    raise TimeoutError("download exceeded its deadline")
                out.write(chunk)
                received += len(chunk)
                if on_progress is not None:
                    on_progress(received, total)
    except (urllib.error.URLError, OSError, ValueError) as exc:
        utils.warning_print(f"Update download failed: {exc}")
        if part is not None:
            part.unlink(missing_ok=True)
        return None

    problem = _verify_file(part, asset)
    if problem is not None:
        utils.warning_print(f"Update download corrupt: {problem}")
        part.unlink(missing_ok=True)
        return None
    try:
        part.replace(target)
    except OSError as exc:
        utils.warning_print(f"Update download failed (rename): {exc}")
        part.unlink(missing_ok=True)
        return None
    return target


def start_download() -> bool:
    """Enter ``downloading`` synchronously; False unless an update is available."""
    with _lock:
        if _status["phase"] != "available" or _latest is None:
            return False
        _status.update(phase="downloading", completed=0, error=None)
    return True


def run_download() -> None:
    """start_download plus finish_download in one call."""
    if start_download():
        finish_download()


def finish_download() -> None:
    """Thread body behind /api/update/download."""
    with _lock:
        release = _latest
        name = _status.get("asset")
    if release is None:
        return
    asset = next((a for a in release["assets"] if a["name"] == name), None)
    if asset is None:
        with _lock:
            _status.update(phase="error", error="No download for this platform")
        return

    def _progress(completed: int, total: int) -> None:
        with _lock:
            _status["completed"] = completed
            _status["total"] = total

    path = download_update(asset, on_progress=_progress)
    with _lock:
        if path is None:
            _status.update(phase="error", error="Download failed")
        else:
            _status.update(phase="ready", path=str(path))


# ---- Apply -------------------------------------------------------------------


def _writable(directory: Path) -> bool:
    """A real create-and-delete probe; os.access lies on some volumes."""
    try:
        with tempfile.NamedTemporaryFile(dir=directory, prefix=".clipgen-probe"):
            pass
    except OSError:
        return False
    return True


def _spawn_detached(command: list[str], cwd: Path) -> None:
    """Start *command* so it outlives this process."""
    kwargs: dict[str, Any] = {
        "stdin": subprocess.DEVNULL,
        "stdout": subprocess.DEVNULL,
        "stderr": subprocess.DEVNULL,
        "close_fds": True,
        "cwd": str(cwd),
    }
    if os.name == "nt":
        kwargs["creationflags"] = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(
            subprocess, "CREATE_NEW_PROCESS_GROUP", 0
        )
    else:
        kwargs["start_new_session"] = True
    subprocess.Popen(command, **kwargs)


def _run(command: list[str], timeout: float) -> None:
    subprocess.run(
        command,
        check=True,
        timeout=timeout,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def render_mac_helper(pid: int, live: Path, staged: Path, log: Path) -> str:
    """The shell script that swaps the bundles once *pid* has exited."""
    q = shlex.quote
    return "\n".join(
        [
            "#!/bin/sh",
            "# clipgen update helper; rewritten on every apply.",
            f"PID={pid}",
            f"LIVE={q(str(live))}",
            f"STAGED={q(str(staged))}",
            f"STAGING={q(str(staged.parent))}",
            f"LOG={q(str(log))}",
            "i=0",
            'while kill -0 "$PID" 2>/dev/null; do',
            "  sleep 0.5",
            "  i=$((i+1))",
            f'  if [ "$i" -ge {HELPER_WAIT_SECONDS * 2} ]; then',
            '    echo "clipgen did not exit" >"$LOG"; exit 1',
            "  fi",
            "done",
            'xattr -dr com.apple.quarantine "$STAGED" 2>/dev/null',
            'if ! mv "$LIVE" "$LIVE.old" 2>>"$LOG"; then',
            '  echo "could not move the old app aside" >>"$LOG"; exit 2',
            "fi",
            'if ! mv "$STAGED" "$LIVE" 2>>"$LOG"; then',
            '  mv "$LIVE.old" "$LIVE"',
            '  echo "could not move the new app into place" >>"$LOG"; exit 3',
            "fi",
            'rm -rf "$LIVE.old" "$STAGING"',
            'rm -f "$LOG"',
            'open -n "$LIVE"',
            "",
        ]
    )


def _ps_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def render_win_helper(pid: int, root: Path, staged: Path, log: Path) -> str:
    """The PowerShell script that swaps the folders once *pid* has exited."""
    # PureWindowsPath so the name splits on backslashes when rendered on macOS/CI.
    win_root = PureWindowsPath(root)
    old = win_root.with_name(win_root.name + ".old")
    q = _ps_quote
    return "\n".join(
        [
            "# clipgen update helper; rewritten on every apply.",
            "$ErrorActionPreference = 'Stop'",
            f"$target = {pid}",
            f"$root = {q(str(win_root))}",
            f"$staged = {q(str(staged))}",
            f"$old = {q(str(old))}",
            f"$log = {q(str(log))}",
            "try {",
            f"  Wait-Process -Id $target -Timeout {HELPER_WAIT_SECONDS} -ErrorAction SilentlyContinue",
            "  $moved = $false",
            "  for ($i = 0; $i -lt 5; $i++) {",
            f"    try {{ Rename-Item -LiteralPath $root -NewName {q(old.name)}; $moved = $true; break }}",
            "    catch { Start-Sleep -Seconds 1 }",
            "  }",
            "  if (-not $moved) { throw 'could not move the old folder aside' }",
            f"  try {{ Rename-Item -LiteralPath $staged -NewName {q(win_root.name)} }}",
            "  catch {",
            f"    Rename-Item -LiteralPath $old -NewName {q(win_root.name)}",
            "    throw 'could not move the new folder into place'",
            "  }",
            "  Remove-Item -LiteralPath $old -Recurse -Force -ErrorAction SilentlyContinue",
            "  Remove-Item -LiteralPath $log -Force -ErrorAction SilentlyContinue",
            "  Start-Process -FilePath (Join-Path $root 'clipgen.exe')",
            "} catch {",
            "  Set-Content -LiteralPath $log -Value $_.Exception.Message",
            "}",
            "",
        ]
    )


def extract_bundle(archive: Path, dest: Path) -> str | None:
    """Unpack the portable zip's single top-level folder into *dest*."""
    try:
        with zipfile.ZipFile(archive) as zf:
            names = zf.namelist()
            if not names:
                return "archive is empty"
            top = names[0].split("/", 1)[0]
            if not top:
                return "archive has no top-level folder"
            for info in zf.infolist():
                parts = Path(info.filename).parts
                if not parts or parts[0] != top or any(p in ("..", "") for p in parts):
                    return f"unexpected archive member: {info.filename}"
                if len(parts) == 1:
                    continue
                out = dest.joinpath(*parts[1:])
                if info.is_dir():
                    out.mkdir(parents=True, exist_ok=True)
                    continue
                out.parent.mkdir(parents=True, exist_ok=True)
                with zf.open(info) as src, out.open("wb") as dst:
                    shutil.copyfileobj(src, dst, _DOWNLOAD_CHUNK)
    except (zipfile.BadZipFile, OSError) as exc:
        return str(exc)
    if not (dest / "clipgen.exe").is_file():
        return "archive has no clipgen.exe"
    return None


def _apply_mac(dmg: Path, bundle: Path) -> str | None:
    parent = bundle.parent
    if str(parent).startswith("/Volumes/") or not _writable(parent):
        return "not-writable"
    mount = Path(tempfile.mkdtemp(prefix="clipgen-dmg-"))
    staging = parent / ".clipgen-update"
    staged = staging / bundle.name
    try:
        _run(
            [
                "hdiutil",
                "attach",
                "-nobrowse",
                "-readonly",
                "-noverify",
                "-noautoopen",
                "-mountpoint",
                str(mount),
                str(dmg),
            ],
            timeout=120,
        )
        source = next(iter(sorted(mount.glob("*.app"))), None)
        if source is None:
            return "disk image holds no app"
        shutil.rmtree(staging, ignore_errors=True)
        staging.mkdir()
        _run(["ditto", str(source), str(staged)], timeout=600)
        _run(["codesign", "--verify", "--deep", str(staged)], timeout=120)
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired, OSError) as exc:
        shutil.rmtree(staging, ignore_errors=True)
        return f"could not stage the new app: {exc}"
    finally:
        subprocess.run(
            ["hdiutil", "detach", str(mount), "-force"],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        shutil.rmtree(mount, ignore_errors=True)
    log = _updates_dir() / APPLY_LOG
    script = _updates_dir() / "apply.sh"
    script.write_text(
        render_mac_helper(os.getpid(), bundle, staged, log), encoding="utf-8"
    )
    _spawn_detached(["/bin/sh", str(script)], cwd=parent)
    return None


def _apply_inno(setup: Path) -> str | None:
    _spawn_detached(
        [
            str(setup),
            "/SILENT",
            "/NORESTART",
            "/SUPPRESSMSGBOXES",
            "/CLOSEAPPLICATIONS",
            "/RELAUNCH=1",
        ],
        cwd=setup.parent,
    )
    return None


def _apply_zip(archive: Path, root: Path) -> str | None:
    parent = root.parent
    if not _writable(parent):
        return "not-writable"
    staged = parent / (root.name + ".new")
    shutil.rmtree(staged, ignore_errors=True)
    problem = extract_bundle(archive, staged)
    if problem is not None:
        shutil.rmtree(staged, ignore_errors=True)
        return f"could not unpack the update: {problem}"
    log = _updates_dir() / APPLY_LOG
    script = _updates_dir() / "apply.ps1"
    script.write_text(
        render_win_helper(os.getpid(), root, staged, log), encoding="utf-8"
    )
    _spawn_detached(
        [
            "powershell",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script),
        ],
        cwd=parent,
    )
    return None


def apply_update(path: Path) -> str | None:
    """Hand the downloaded asset to the platform installer; error text or None.

    On success a detached helper (or the installer) owns the rest, and the
    caller must quit so the files can be replaced.
    """
    shape = install_shape()
    root = install_root()
    if root is None:
        return "unsupported install"
    try:
        if shape == "mac-app":
            return _apply_mac(path, root)
        if shape == "win-inno":
            return _apply_inno(path)
        return _apply_zip(path, root)
    except OSError as exc:
        return str(exc)


def start_apply() -> bool:
    """Enter ``applying`` synchronously; False unless a download is ready."""
    with _lock:
        if _status["phase"] != "ready" or not _status.get("path"):
            return False
        _status.update(phase="applying", error=None)
    return True


def run_apply() -> None:
    """start_apply plus finish_apply in one call."""
    if start_apply():
        finish_apply()


def finish_apply() -> None:
    """Thread body behind /api/update/apply; quits the app on success."""
    with _lock:
        path = Path(_status["path"])
    problem = apply_update(path)
    if problem is not None:
        with _lock:
            _status.update(phase="ready" if problem == "not-writable" else "error")
            _status["error"] = (
                "This install cannot be replaced in place"
                if problem == "not-writable"
                else problem
            )
        return
    request_quit()


def skip_version() -> bool:
    """Hide the offered release until a newer one or a manual check."""
    global _latest
    with _lock:
        tag = _status.get("version")
        if _status["phase"] not in ("available", "ready") or not tag:
            return False
        _latest = None
        _status.update(phase="idle", skipped=tag, error=None)
    _save_state(skipped=tag)
    return True


def request_quit() -> None:
    """Close the desktop window so the process unwinds and exits."""
    import desktop

    desktop.request_quit()


def reveal_download() -> bool:
    """Show the downloaded asset in the file manager."""
    with _lock:
        path = _status.get("path")
    if not path or not Path(path).is_file():
        return False
    return utils.reveal_in_file_manager(Path(path))


# ---- Startup -----------------------------------------------------------------


def sweep_updates_dir() -> None:
    """Report a failed apply, drop partial and outdated downloads, clear staging."""
    directory = _updates_dir()
    with _lock:
        _status["supported"] = is_supported()
        _status["current"] = utils.get_version()
        _status["skipped"] = _skipped_tag()
    log = directory / APPLY_LOG
    try:
        if log.is_file():
            text = log.read_text(encoding="utf-8", errors="replace").strip()
            with _lock:
                _status["last_error"] = text or "unknown failure"
            log.unlink()
    except OSError:
        pass
    current = utils.get_version()
    for entry in list(directory.glob("*")) if directory.is_dir() else []:
        try:
            if entry.suffix == ".part":
                entry.unlink()
                continue
            match = re.match(r"^clipgen-(v[\d.]+)-", entry.name)
            if match and not is_newer(match.group(1), current):
                entry.unlink()
        except OSError:
            pass
    root = install_root()
    if root is None:
        return
    leftovers = [root.parent / ".clipgen-update"]
    if install_shape() == "win-zip":
        leftovers += [
            root.with_name(root.name + ".new"),
            root.with_name(root.name + ".old"),
        ]
    for stale in leftovers:
        shutil.rmtree(stale, ignore_errors=True)


def status() -> dict[str, Any]:
    """A snapshot of the update state for the API."""
    with _lock:
        snapshot = dict(_status)
    snapshot["supported"] = is_supported()
    snapshot["auto_check"] = bool(config.UPDATE_CHECK_ON_LAUNCH)
    return snapshot


def reset_for_tests() -> None:
    """Return the module state to its import-time shape."""
    global _latest
    with _lock:
        _latest = None
        _status.update(
            phase="idle",
            checked=False,
            version=None,
            release_url=None,
            asset=None,
            path=None,
            completed=0,
            total=0,
            error=None,
            last_error=None,
            skipped=None,
        )
