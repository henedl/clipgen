"""Locate Playwright and a Chromium binary without letting either resolve the other.

Two independent problems, both solved here:

**The import.** ``playwright`` lives in the optional ``ui`` extra, so it is
absent in CI. A static ``from playwright.sync_api import sync_playwright`` would
be an ``unresolved-import`` there, and a ``# ty: ignore`` for it flips to
``unused-ignore-comment`` on a machine where the extra *is* installed — no
suppression is correct in both states. ``importlib.import_module`` returns a
``ModuleType``, whose stub declares ``__getattr__ -> Any``, so ty stays quiet
either way.

**The binary.** The pip package hard-codes the browser build number it expects,
and that drifts from whatever ``playwright install`` last downloaded — playwright
1.61 demanded ``chromium_headless_shell-1228`` on a machine holding 1217 and
refused to launch, while the same 1217 binary passed as ``executable_path``
worked immediately. So we never let playwright pick: glob the cache, take the
highest build, pass it explicitly.
"""

import importlib
import os
import sys
from pathlib import Path
from typing import Any

from _ui_fixtures import UiUnavailable

# Ordered by preference. The purpose-built headless shell first; the full
# Chromium build is the fallback for installs that only carry it.
_CANDIDATES: tuple[tuple[str, tuple[str, ...]], ...] = (
    (
        "chromium_headless_shell-*",
        ("*/chrome-headless-shell", "*/chrome-headless-shell.exe"),
    ),
    ("chromium-*", ("*/*.app/Contents/MacOS/*", "*/chrome", "*/chrome.exe")),
)

LAUNCH_ARGS = [
    "--no-sandbox",
    "--disable-dev-shm-usage",
]


def sync_playwright() -> Any:
    """Return ``playwright.sync_api.sync_playwright``, or raise :class:`UiUnavailable`."""
    try:
        module = importlib.import_module("playwright.sync_api")
    except ImportError as exc:
        raise UiUnavailable(
            "playwright is not installed. It lives in the optional `ui` extra, "
            "which is never installed by CI or a plain `uv sync`:\n"
            "  uv sync --extra dev --extra ui"
        ) from exc
    return module.sync_playwright


def playwright_error() -> type[BaseException]:
    """The ``playwright.sync_api.Error`` class, for catching in-page failures.

    Fetched through the same dynamic import as :func:`sync_playwright` so a
    caller can narrow its ``except`` without a static playwright import.
    """
    return importlib.import_module("playwright.sync_api").Error


def browsers_root() -> Path:
    """The ms-playwright download cache for this platform."""
    override = os.environ.get("PLAYWRIGHT_BROWSERS_PATH")
    if override:
        return Path(override).expanduser()
    # The developer's real home: tests/conftest.py sandboxes Path.home().
    home = Path(os.path.expanduser("~"))
    if sys.platform == "darwin":
        return home / "Library" / "Caches" / "ms-playwright"
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA") or str(home / "AppData" / "Local")
        return Path(base) / "ms-playwright"
    return home / ".cache" / "ms-playwright"


def _build_number(name: str) -> int:
    """``1217`` from ``chromium-1217``; ``-1`` when the suffix isn't numeric."""
    tail = name.rsplit("-", 1)[-1]
    return int(tail) if tail.isdigit() else -1


def resolve_chromium(prefer_full: bool = False) -> Path:
    """Return the highest-numbered installed Chromium executable.

    ``prefer_full=True`` flips the candidate order so the full Chromium build
    wins over the headless shell: the shell's paint/compositor timings are
    only indicative (no GPU/compositor parity), so perf captures that care
    about paint fidelity opt into the full build (``shot.py --full-chromium``).

    Raises :class:`UiUnavailable` with the install command when none is found.
    """
    root = browsers_root()
    candidates = tuple(reversed(_CANDIDATES)) if prefer_full else _CANDIDATES
    for dir_glob, exe_globs in candidates:
        directories = sorted(
            (path for path in root.glob(dir_glob) if path.is_dir()),
            key=lambda path: _build_number(path.name),
            reverse=True,
        )
        for directory in directories:
            for exe_glob in exe_globs:
                for exe in sorted(directory.glob(exe_glob)):
                    if exe.is_file() and os.access(exe, os.X_OK):
                        return exe
    raise UiUnavailable(
        f"No Playwright Chromium found under {root}.\n"
        "Install one (~150 MB, downloads outside the repo):\n"
        "  uv run --extra ui playwright install chromium\n"
        "Or point PLAYWRIGHT_BROWSERS_PATH at an existing cache."
    )
