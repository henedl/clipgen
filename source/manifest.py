"""The sectioned ``clipgen.json`` manifest: one state file per output dir.

Every tool keeps its project state as one section (``clips``, ``screenspace``,
``transcripts``, ``composer``, ``workflows``, ``stashes``, ``convergence``).
Sections are cached as raw JSON text keyed by the file's ``(mtime_ns, size)``,
written through a ``.tmp`` sibling and ``os.replace``, and serialized by a
process lock plus the cross-process :func:`utils.file_lock`. A section is
dropped when empty; the file disappears with its last section.
"""

from __future__ import annotations

import json
import os
import threading
import time
from pathlib import Path
from collections.abc import Callable
from typing import Any

import config
import profiling
from utils import (
    file_lock,
    get_effective_output_dir,
    warning_print,
)

_MANIFEST_LOCK = threading.Lock()

_manifest_cache: dict[str, Any] = {
    "path": None,
    "stamp": None,
    "sections": {},
    # An unreadable file disables saves, so one bad read can't wipe every section.
    "broken": False,
}


def _manifest_path() -> Path:
    return Path(get_effective_output_dir()) / config.MANIFEST_FILENAME


def _manifest_stamp(path: Path) -> tuple[int, int] | None:
    """(mtime_ns, size) of *path*, or None when absent."""
    try:
        st = path.stat()
    except OSError:
        return None
    return (st.st_mtime_ns, st.st_size)


def _dump_section(data: Any) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2)


def _read_sections(path: Path) -> dict[str, str] | None:
    """Parse the whole file into per-section JSON text; None when unreadable."""
    try:
        doc = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        warning_print(
            f"Could not read {path.name}; saving is paused until it is fixed."
        )
        return None
    if not isinstance(doc, dict):
        warning_print(f"{path.name} is not a JSON object; saving is paused.")
        return None
    return {str(k): _dump_section(v) for k, v in doc.items()}


def _sections_locked() -> dict[str, str]:
    """Cached section texts, refreshed when the file's path or stamp changed."""
    path = _manifest_path()
    stamp = _manifest_stamp(path)
    if _manifest_cache["path"] != str(path) or _manifest_cache["stamp"] != stamp:
        sections = _read_sections(path) if stamp else {}
        _manifest_cache["path"] = str(path)
        _manifest_cache["stamp"] = stamp
        _manifest_cache["sections"] = sections or {}
        _manifest_cache["broken"] = sections is None
    return _manifest_cache["sections"]


# Re-indented section text keyed by the text; unchanged multi-MB sections skip
# re-indenting. Pruned on write.
_manifest_indent_cache: dict[str, str] = {}


def _indented_section(text: str) -> str:
    """*text* shifted one nesting level, memoized (texts are reused verbatim)."""
    cached = _manifest_indent_cache.get(text)
    if cached is None:
        cached = _manifest_indent_cache[text] = text.replace("\n", "\n  ")
    return cached


def _write_sections_locked(sections: dict[str, str]) -> Path | None:
    """Atomically rewrite the file from section texts; delete it when empty."""
    path = _manifest_path()
    tmp = path.with_suffix(path.suffix + ".tmp")
    if not sections:
        for candidate in (path, tmp):
            try:
                candidate.unlink(missing_ok=True)
            except OSError:
                pass
        _manifest_cache["path"] = str(path)
        _manifest_cache["stamp"] = None
        _manifest_cache["sections"] = {}
        _manifest_cache["broken"] = False
        _manifest_indent_cache.clear()
        return None
    # Sections are already indent=2; json.dumps escapes newlines, so nesting is a
    # plain re-indent.
    body = ",\n".join(
        f"  {json.dumps(key)}: {_indented_section(text)}"
        for key, text in sorted(sections.items())
    )
    if len(_manifest_indent_cache) > len(sections):
        live = set(sections.values())
        for stale in [t for t in _manifest_indent_cache if t not in live]:
            del _manifest_indent_cache[stale]
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp.write_text("{\n" + body + "\n}\n", encoding="utf-8")
        os.replace(tmp, path)
    except OSError as exc:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        warning_print(f"Could not write {path.name}: {exc}")
        return None
    _manifest_cache["path"] = str(path)
    _manifest_cache["stamp"] = _manifest_stamp(path)
    _manifest_cache["sections"] = sections
    _manifest_cache["broken"] = False
    return path


def load_manifest_section(section: str, *, default: Any = None) -> Any:
    """Parse one section of the output-dir manifest; *default* when absent.

    Every call returns fresh objects, so callers may mutate the result freely.
    Profiled as ``manifest.load <section>`` with the section's text size: the
    parse is per call, so a poll that reloads a multi-MB section shows here.
    """
    t0 = time.perf_counter() if config.PROFILING else 0.0
    with _MANIFEST_LOCK:
        text = _sections_locked().get(section)
    if text is None:
        return default
    data = json.loads(text)
    if t0:
        profiling.add(
            f"manifest.load {section}", time.perf_counter() - t0, nbytes=len(text)
        )
    return data


def save_manifest_section(section: str, data: Any) -> Path | None:
    """Persist one section; ``None`` removes it. Deletes the file when empty.

    Writes via a sibling .tmp file and ``os.replace()`` so a crash mid-write
    leaves the previous manifest intact. Returns the path, or ``None`` on
    failure or removal.

    A save whose text matches the stored section skips the write entirely —
    idempotent persists (startup rewrites, debounced saves with no delta) cost
    nothing and leave the mtime alone, so pollers see no phantom change.
    """
    text: str | None = None
    t0 = time.perf_counter() if config.PROFILING else 0.0
    if data is not None:
        try:
            text = _dump_section(data)
        except (TypeError, ValueError) as exc:
            warning_print(f"Could not serialize {section}: {exc}")
            return None
    with _MANIFEST_LOCK, file_lock(_manifest_path()):
        # The stamp check inside re-reads what another process wrote meanwhile.
        sections = dict(_sections_locked())
        if _manifest_cache["broken"]:
            return None
        if text is None:
            if section not in sections and sections:
                return _manifest_path()
            sections.pop(section, None)
        else:
            if sections.get(section) == text:
                if t0:
                    profiling.count(f"manifest.save.unchanged {section}")
                return _manifest_path()
            sections[section] = text
        written = _write_sections_locked(sections)
    if t0:
        # Dump + re-indent + write; bytes is this section alone, not the file.
        profiling.add(
            f"manifest.save {section}",
            time.perf_counter() - t0,
            nbytes=len(text) if text else 0,
        )
    return written


def memo_section(
    memo: dict[str, tuple[tuple[int, int] | None, Any]],
    key: str,
    section: str,
    build: Callable[[dict[str, Any]], Any],
) -> Any:
    """Rebuild a derived view of *section* only when the file's stamp changed.

    Pollers keep a ``memo`` dict of ``key -> (stamp, value)`` and call this per
    poll; the parse runs once per on-disk change instead of once per request.
    Delete ``memo[key]`` to force a fresh parse. Decoded sections are not cached
    directly on purpose: the deep copy a mutable return would need costs as much
    as the parse.
    """
    stamp = _manifest_stamp(_manifest_path())
    cached = memo.get(key)
    if cached is not None and stamp == cached[0]:
        return cached[1]
    value = (
        build(load_manifest_section(section, default={}) or {}) if stamp else build({})
    )
    memo[key] = (stamp, value)
    return value


def manifest_sections() -> set[str]:
    """Names of the sections present on disk, without parsing them."""
    with _MANIFEST_LOCK:
        return set(_sections_locked())


def manifest_mtime() -> int:
    """The manifest file's mtime_ns, or 0 when absent."""
    stamp = _manifest_stamp(_manifest_path())
    return stamp[0] if stamp else 0


def _reset_manifest_cache() -> None:
    """Drop the in-memory section cache. Intended for test fixtures."""
    with _MANIFEST_LOCK:
        _manifest_cache["path"] = None
        _manifest_cache["stamp"] = None
        _manifest_cache["sections"] = {}
        _manifest_indent_cache.clear()
        _manifest_cache["broken"] = False


def sweep_stale_temp_artifacts() -> None:
    """Remove orphaned atomic-write tmps and reel temp-clips from the output dir.

    Targets only our own artifacts — ``*.json.tmp`` siblings (manifest atomic
    writes) and ``{TEMP_ARTIFACT_PREFIX}*`` reel temp-clips — so user files are
    never touched. Meant to run once at server startup, before any worker thread,
    to reclaim leftovers from a prior hard kill.
    """
    base = Path(get_effective_output_dir())
    if not base.is_dir():
        return
    for pattern in ("*.json.tmp", config.TEMP_ARTIFACT_PREFIX + "*"):
        for stale in base.glob(pattern):
            try:
                stale.unlink(missing_ok=True)
            except OSError:
                pass
