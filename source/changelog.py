"""Parse ``CHANGELOG.md`` into structured entries for the Start overlay.

Each release is one level-2 heading::

    ## <version> — <YYYY-MM-DD>

followed by one line per change::

    **<Tool>:** <Feat|Fix>: <one or two plain sentences>

A release groups every change that shipped under that version, so one entry
carries a list of changes rather than a single title/body pair.

The heading separator must carry whitespace on both sides. Without that the
date's own hyphens read as separators too, and ``2026-08-20`` silently parses
as date ``2026`` / tool ``20`` — which is exactly how the previous
three-part heading format failed when the tool field was dropped.

The first heading is the most recent entry; ordering in the file wins (we do
not re-sort by version).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import utils


_HEADING_RE = re.compile(r"^##\s+(?P<version>\S+)\s+[—-]\s+(?P<date>\S+)\s*$")
_CHANGE_RE = re.compile(
    r"^\*\*(?P<tool>[^:*]+):\*\*\s*(?:(?P<kind>Feat|Fix):\s*)?(?P<text>.+?)\s*$"
)


def _changelog_path() -> Path:
    return utils.get_bundled_assets_root() / "CHANGELOG.md"


def parse_entries(text: str) -> list[dict[str, Any]]:
    """Parse changelog markdown into ``{version, date, changes}`` releases.

    Pure, and takes the file's *text* rather than a path, so ``build/release_notes.py``
    can feed it a changelog read from anywhere while the heading regex — which has
    silently mis-parsed the format once already (see the module docstring) — stays in
    one place. Releases with no parsable change lines are kept here; dropping them is
    the caller's policy, and only the Start overlay wants it.
    """
    entries: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None

    for raw in text.splitlines():
        line = raw.strip()
        heading = _HEADING_RE.match(line)
        if heading is not None:
            current = {
                "version": heading.group("version"),
                "date": heading.group("date"),
                "changes": [],
            }
            entries.append(current)
            continue
        if current is None:
            continue
        change = _CHANGE_RE.match(line)
        if change is not None:
            current["changes"].append(
                {
                    "tool": change.group("tool").strip(),
                    "kind": (change.group("kind") or "").strip(),
                    "text": change.group("text").strip(),
                }
            )

    return entries


def load_entries(limit: int = 20) -> list[dict[str, Any]]:
    """Return up to *limit* releases in file order (newest first)."""
    path = _changelog_path()
    if not path.is_file():
        utils.warning_print(
            f"CHANGELOG.md not found at {path}; Start overlay 'Recent updates' will be empty."
        )
        return []
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        utils.warning_print(f"Could not read CHANGELOG.md at {path}: {exc}")
        return []

    # A heading with no change lines would render as an empty card.
    return [e for e in parse_entries(text) if e["changes"]][:limit]
