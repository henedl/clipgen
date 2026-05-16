# -*- coding: utf-8 -*-
"""Parse ``CHANGELOG.md`` into structured entries for the Start overlay.

Each entry corresponds to a level-2 heading in the form::

    ## <version> — <YYYY-MM-DD> — <tool>

followed by a bolded title line and an optional body paragraph. The first
heading is treated as the most recent entry; ordering in the file wins (we
do not re-sort by version).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import utils


_HEADING_RE = re.compile(
    r"^##\s+(?P<version>\S+)\s*[—-]\s*(?P<date>\S+)\s*[—-]\s*(?P<tool>.+?)\s*$"
)
_TITLE_RE = re.compile(r"^\*\*(?P<title>.+?)\*\*\s*$")


def _changelog_path() -> Path:
    return utils.get_bundled_assets_root() / "CHANGELOG.md"


def load_entries(limit: int = 20) -> list[dict[str, Any]]:
    """Return up to *limit* changelog entries in file order (newest first)."""
    path = _changelog_path()
    if not path.is_file():
        return []
    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return []
    entries: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    body_lines: list[str] = []

    def _flush() -> None:
        if current is None:
            return
        current["body"] = " ".join(body_lines).strip()
        entries.append(current)

    for raw in text.splitlines():
        line = raw.rstrip()
        heading = _HEADING_RE.match(line)
        if heading is not None:
            _flush()
            body_lines = []
            current = {
                "version": heading.group("version").strip(),
                "date": heading.group("date").strip(),
                "tool": heading.group("tool").strip(),
                "title": "",
                "body": "",
            }
            continue
        if current is None:
            continue
        if not current["title"]:
            title = _TITLE_RE.match(line)
            if title is not None:
                current["title"] = title.group("title").strip()
                continue
            if line:
                current["title"] = line.strip()
                continue
        elif line:
            body_lines.append(line.strip())
    _flush()
    return entries[:limit]
