"""Persistent per-user clipgen state: Start overlay settings and window geometry.

Stores last-used input/output directories and last-used spreadsheet so the
frontend Start overlay can prefill its inputs across launches, plus the desktop
window's last size and position. The user can toggle persistence off via
``persist_enabled``; when disabled, the recording helpers short-circuit but
``persist_enabled`` itself is still written so the toggle survives sessions.

Settings file location:

- Windows: ``%LOCALAPPDATA%\\clipgen\\start.json`` (fallback:
  ``~/AppData/Local/clipgen/start.json`` when ``LOCALAPPDATA`` is unset)
- macOS / Linux: ``~/.config/clipgen/start.json``

Every mutating helper holds ``_write_lock`` across its whole load-mutate-save
sequence. ``save_start_settings`` rewrites the entire dict, so two interleaved
read-modify-write cycles silently drop one side's field — and the writers are
genuinely concurrent: the Start overlay records from Flask request threads while
``desktop.py`` records window geometry from a debounce timer thread.
"""

import json
import os
import sys
import threading
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import utils


RECENTS_CAP = 12

# Guards load -> mutate -> save in every helper below. Deliberately not inside
# load_start_settings/save_start_settings: locking the halves separately would
# still let two cycles interleave, which is the actual race.
_write_lock = threading.Lock()


def config_dir() -> Path:
    """Return the per-user clipgen config directory for this platform.

    Public because it is the canonical answer to "where does per-user clipgen
    state live": desktop.py keeps the webview profile alongside start.json, and
    cli.py searches here for credentials.json.
    """
    if sys.platform == "win32":
        local_appdata = os.environ.get("LOCALAPPDATA")
        if local_appdata:
            return Path(local_appdata) / "clipgen"
        return Path.home() / "AppData" / "Local" / "clipgen"
    return Path.home() / ".config" / "clipgen"


def _settings_path() -> Path:
    return config_dir() / "start.json"


def _defaults() -> dict[str, Any]:
    return {
        "persist_enabled": True,
        "last_input": "",
        "last_output": "",
        "recent_inputs": [],
        "recent_outputs": [],
        "last_spreadsheet": None,
        "recent_spreadsheets": [],
        # Full session records: (name, input, output, spreadsheet|None,
        # last_opened). Powers the Start overlay's "Recently opened" rail and
        # lets a click restore the name plus all three picker values at once.
        "recent_projects": [],
        # Desktop window rect: {"x", "y", "width", "height"} or None for
        # "use the defaults". See desktop.py.
        "window": None,
    }


def load_start_settings() -> dict[str, Any]:
    """Return the persisted settings, falling back to defaults on missing/corrupt file."""
    path = _settings_path()
    if not path.is_file():
        return _defaults()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return _defaults()
    if not isinstance(data, dict):
        return _defaults()
    merged = _defaults()
    merged.update(data)
    return merged


def save_start_settings(settings: dict[str, Any]) -> None:
    """Persist *settings* to the platform config path via atomic replace."""
    path = _settings_path()
    tmp = path.with_suffix(path.suffix + ".tmp")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp.write_text(
            json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        os.replace(tmp, path)
    except (OSError, TypeError, ValueError) as exc:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        utils.warning_print(f"Could not write start settings: {exc}")


def _prepend_dedup(items: list[Any], new_item: Any, key: Any = None) -> list[Any]:
    if key is None:
        deduped = [x for x in items if x != new_item]
    else:
        new_key = key(new_item)
        deduped = [x for x in items if key(x) != new_key]
    return [new_item] + deduped[: RECENTS_CAP - 1]


def record_recent_input(path: str) -> None:
    """Record *path* as the last-used input directory."""
    with _write_lock:
        settings = load_start_settings()
        if not settings.get("persist_enabled", True):
            return
        if not path:
            return
        settings["last_input"] = path
        settings["recent_inputs"] = _prepend_dedup(
            settings.get("recent_inputs", []), path
        )
        save_start_settings(settings)


def record_recent_output(path: str) -> None:
    """Record *path* as the last-used output directory."""
    with _write_lock:
        settings = load_start_settings()
        if not settings.get("persist_enabled", True):
            return
        if not path:
            return
        settings["last_output"] = path
        settings["recent_outputs"] = _prepend_dedup(
            settings.get("recent_outputs", []), path
        )
        save_start_settings(settings)


def record_recent_spreadsheet(
    type_: str, id_or_path: str, label: str, worksheet: str = ""
) -> None:
    """Record a spreadsheet selection as last/recent."""
    with _write_lock:
        settings = load_start_settings()
        if not settings.get("persist_enabled", True):
            return
        if type_ not in ("google", "excel") or not id_or_path:
            return
        entry = {
            "type": type_,
            "id_or_path": id_or_path,
            "label": label or id_or_path,
            "worksheet": worksheet,
        }
        settings["last_spreadsheet"] = entry
        settings["recent_spreadsheets"] = _prepend_dedup(
            settings.get("recent_spreadsheets", []),
            entry,
            key=lambda x: (x.get("type"), x.get("id_or_path")),
        )
        save_start_settings(settings)


def _spreadsheet_key(spreadsheet: dict[str, Any] | None) -> tuple[str, str] | None:
    if not spreadsheet:
        return None
    type_ = spreadsheet.get("type") or ""
    id_or_path = spreadsheet.get("id_or_path") or ""
    if not type_ or not id_or_path:
        return None
    return (type_, id_or_path)


def _project_key(project: dict[str, Any]) -> tuple[Any, Any, tuple[str, str] | None]:
    return (
        project.get("input"),
        project.get("output"),
        _spreadsheet_key(project.get("spreadsheet")),
    )


def record_project_session(
    input_dir: str,
    output_dir: str,
    spreadsheet: dict[str, Any] | None = None,
    name: str | None = None,
) -> None:
    """Record an "Open workspace" event as a full project session.

    A project is identified by the triple ``(input, output, spreadsheet_key)``
    where ``spreadsheet_key`` is ``(type, id_or_path)`` or ``None``. Re-opening
    the same triple moves the entry to the front rather than duplicating it.

    *name* is the user's optional label for the project. It is metadata, never
    part of the identity triple, and is tri-state: ``None`` keeps whatever the
    matching entry already had, ``""`` clears it, and a non-empty string sets
    it. The carry-over matters because only the Start overlay knows the name —
    the CLI-launch and Studio sheet-switch call sites pass nothing, and without
    it every relaunch would silently wipe the label.
    """
    with _write_lock:
        settings = load_start_settings()
        if not settings.get("persist_enabled", True):
            return
        if not input_dir or not output_dir:
            return
        projects: list[Any] = settings.get("recent_projects", [])
        entry: dict[str, Any] = {
            "name": "",
            "input": input_dir,
            "output": output_dir,
            "spreadsheet": spreadsheet if spreadsheet else None,
            "last_opened": datetime.now(UTC).isoformat(),
        }
        if name is None:
            key = _project_key(entry)
            for existing in projects:
                if isinstance(existing, dict) and _project_key(existing) == key:
                    entry["name"] = existing.get("name") or ""
                    break
        else:
            entry["name"] = name.strip()
        settings["recent_projects"] = _prepend_dedup(projects, entry, key=_project_key)
        save_start_settings(settings)


def load_window_geometry() -> dict[str, int] | None:
    """Return the persisted desktop window rect, or None if absent/malformed.

    Shape only: the caller owns every judgement about whether the rect is
    *usable* (min size, on-screen), because that needs screen geometry this
    module has no business knowing about.
    """
    window = load_start_settings().get("window")
    if not isinstance(window, dict):
        return None
    rect: dict[str, int] = {}
    for key in ("x", "y", "width", "height"):
        value = window.get(key)
        # bool is an int subclass and would sail through int(); reject it.
        if not isinstance(value, (int, float)) or isinstance(value, bool):
            return None
        rect[key] = int(value)
    return rect


def record_window_geometry(x: int, y: int, width: int, height: int) -> None:
    """Record the desktop window's rect for the next launch."""
    with _write_lock:
        settings = load_start_settings()
        if not settings.get("persist_enabled", True):
            return
        if width <= 0 or height <= 0:
            return
        settings["window"] = {
            "x": int(x),
            "y": int(y),
            "width": int(width),
            "height": int(height),
        }
        save_start_settings(settings)


def clear_window_geometry() -> None:
    """Forget the stored window rect so the next launch uses the defaults.

    Ungated by ``persist_enabled``, like ``set_persist_enabled``: an explicit
    reset has to take effect whatever the toggle says.
    """
    with _write_lock:
        settings = load_start_settings()
        settings["window"] = None
        save_start_settings(settings)


def set_persist_enabled(enabled: bool) -> None:
    """Toggle the persist_enabled flag and save."""
    with _write_lock:
        settings = load_start_settings()
        settings["persist_enabled"] = bool(enabled)
        save_start_settings(settings)
