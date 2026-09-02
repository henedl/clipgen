"""Parse the SUMMARY table of ``THIRD-PARTY-LICENSES`` for the Start overlay.

The notice file opens with a fixed-width table listing every bundled component::

    ===============================================================================
    SUMMARY
    ===============================================================================

    Component                   Version    License
    -------------------------   --------   ---------------
    FFmpeg (ffmpeg + ffprobe)   8.1.x      GPL-3.0-or-later
    gspread                     6.2.1      MIT

Two shapes complicate a naive line split, and both are load-bearing:

* **Indented sub-rows** (``  FFmpeg DLL (in cv2)``, ``  PP-OCR models (bundled)``)
  belong to the component above them. They are flagged ``nested`` and never start
  a new license group.
* **Wrapped license cells.** A long license note continues on the next line with
  nothing in the component/version columns. Dropping that line would silently
  truncate the notice, so a lone field is appended to the previous row.

Columns are split on runs of two or more spaces rather than by character offset:
the sub-rows are indented, so their columns do not line up with the top-level
ones and any offset-based read would mangle them.

Ordering in the file wins — we never re-sort. The table already clusters
components by license, so grouping consecutive rows on ``group`` reproduces the
notice's own structure (same principle as :mod:`changelog`).
"""

from __future__ import annotations

import re
from typing import Any

import utils


# A section rule in the notice file: a run of '=' on its own line.
_RULE_RE = re.compile(r"^={10,}$")
# The table header, whose following '-----' divider is skipped with it.
_HEADER_RE = re.compile(r"^Component\s+Version\s+License\s*$")
_COLUMN_SPLIT_RE = re.compile(r"\s{2,}")
# Where a compound license stops naming one family: "MIT (macOS only)", "MPL-2.0 AND MIT".
_GROUP_SEPARATORS = (" (", " + ", " AND ")


def _licenses_text() -> str | None:
    """Return the raw notice file, or None when it is not present.

    Indirection over :func:`utils.get_licenses_text` so tests have one seam to
    monkeypatch, the way ``changelog._changelog_path`` is patched.
    """
    return utils.get_licenses_text()


def _license_group(license_text: str) -> str:
    """Reduce a license cell to the license family it belongs to."""
    for separator in _GROUP_SEPARATORS:
        index = license_text.find(separator)
        if index != -1:
            license_text = license_text[:index]
    return license_text.strip()


def load_components() -> list[dict[str, Any]]:
    """Return the SUMMARY table rows in file order.

    Each row is ``{"component", "version", "license", "group", "nested"}``.
    Returns ``[]`` when the notice file is missing or has no parseable table —
    a source checkout stripped of `build/` must not break the About tab.
    """
    text = _licenses_text()
    if not text:
        return []

    lines = text.splitlines()
    try:
        start = _summary_start(lines)
    except ValueError:
        utils.warning_print(
            "THIRD-PARTY-LICENSES has no SUMMARY table; "
            "the Start overlay's third-party list will be empty."
        )
        return []

    rows: list[dict[str, Any]] = []
    for raw in lines[start:]:
        if _RULE_RE.match(raw.strip()):
            break
        stripped = raw.strip()
        if not stripped or stripped.startswith("---"):
            continue
        parts = _COLUMN_SPLIT_RE.split(stripped)
        if len(parts) >= 3:
            license_text = " ".join(parts[2:])
            rows.append(
                {
                    "component": parts[0],
                    "version": parts[1],
                    "license": license_text,
                    "group": _license_group(license_text),
                    "nested": raw.startswith("  "),
                }
            )
        elif rows and len(parts) == 1:
            # A wrapped license cell: continues the row above.
            rows[-1]["license"] = f"{rows[-1]['license']} {parts[0]}"
            rows[-1]["group"] = _license_group(rows[-1]["license"])
    return rows


def _summary_start(lines: list[str]) -> int:
    """Return the index of the first table row after the SUMMARY header.

    Raises ValueError if the file carries no SUMMARY section, which means the
    notice was restructured and this parser needs revisiting.
    """
    for index, line in enumerate(lines):
        if index == 0 or line.strip() != "SUMMARY":
            continue
        if not _RULE_RE.match(lines[index - 1].strip()):
            continue
        for offset in range(index + 1, len(lines)):
            if _HEADER_RE.match(lines[offset].strip()):
                # Skip the header and the '-----' divider under it.
                return offset + 2
        break
    raise ValueError("no SUMMARY table found")
