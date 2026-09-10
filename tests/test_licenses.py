"""Tests for the THIRD-PARTY-LICENSES SUMMARY parser used by the Start overlay."""

import re
import tomllib
from pathlib import Path

import pytest

import licenses

_ROOT = Path(__file__).resolve().parent.parent


HEADER = """THIRD-PARTY SOFTWARE NOTICES AND LICENSES

===============================================================================
SUMMARY
===============================================================================

Component                   Version    License
-------------------------   --------   ---------------
"""

FOOTER = """

===============================================================================
MIT LICENSE
===============================================================================

gspread
Copyright (C) 2011-2023 Anton Burnashev
"""


@pytest.fixture
def notice(monkeypatch):
    """Install a fake notice file; the callable takes the table body."""

    def _install(table: str) -> None:
        monkeypatch.setattr(licenses, "_licenses_text", lambda: HEADER + table + FOOTER)

    return _install


def test_returns_empty_when_notice_missing(monkeypatch):
    monkeypatch.setattr(licenses, "_licenses_text", lambda: None)
    assert licenses.load_components() == []


def test_returns_empty_when_there_is_no_summary_table(monkeypatch):
    monkeypatch.setattr(licenses, "_licenses_text", lambda: "no table here\n")
    assert licenses.load_components() == []


def test_parses_plain_rows(notice):
    notice("gspread                     6.2.1      MIT\n")
    assert licenses.load_components() == [
        {
            "component": "gspread",
            "version": "6.2.1",
            "license": "MIT",
            "group": "MIT",
            "nested": False,
        }
    ]


def test_stops_at_the_next_section_rule(notice):
    notice("gspread                     6.2.1      MIT\n")
    # The MIT LICENSE heading and the copyright block after it are not rows.
    assert [row["component"] for row in licenses.load_components()] == ["gspread"]


def test_indented_rows_are_flagged_nested(notice):
    notice(
        "opencv-python-headless      4.14.0     MIT (bindings)\n"
        "  PP-OCR models (bundled)     v4/v5      Apache-2.0\n"
    )
    rows = licenses.load_components()
    assert [row["nested"] for row in rows] == [False, True]
    assert rows[1]["component"] == "PP-OCR models (bundled)"


def test_wrapped_license_cell_is_joined_onto_the_row_above(notice):
    """A dropped continuation would silently truncate the notice."""
    notice(
        "  FFmpeg DLL (in cv2)         8.x        LGPL-2.1 (Windows zip only; the macOS\n"
        "                                         cv2 is self-built without FFmpeg)\n"
    )
    rows = licenses.load_components()
    assert len(rows) == 1
    assert rows[0]["license"] == (
        "LGPL-2.1 (Windows zip only; the macOS cv2 is self-built without FFmpeg)"
    )
    assert rows[0]["group"] == "LGPL-2.1"


@pytest.mark.parametrize(
    "cell,expected",
    [
        ("MIT", "MIT"),
        ("MIT (macOS only)", "MIT"),
        ("MIT (bindings) + Apache 2.0 (OpenCV)", "MIT"),
        ("MPL-2.0 AND MIT", "MPL-2.0"),
        ("Apache-2.0 OR BSD-3-Clause", "Apache-2.0"),
        ("HPND (MIT-CMU)", "HPND"),
        ("GPL-3.0-or-later", "GPL-3.0-or-later"),
        ("SIL OFL 1.1", "SIL OFL 1.1"),
    ],
)
def test_license_group_reduces_to_the_family(cell, expected):
    assert licenses._license_group(cell) == expected


# ---- Against the file that actually ships ------------------------------


def test_real_notice_parses():
    """Catches a hand-edit that breaks the table's column alignment."""
    rows = licenses.load_components()
    assert len(rows) > 20
    for row in rows:
        assert row["component"], row
        assert row["license"], row
        assert row["group"], row


def test_real_notice_lists_every_bundled_asset_class():
    """Anything absent from the SUMMARY table is invisible in the About tab."""
    components = {row["component"] for row in licenses.load_components()}
    for expected in (
        "FFmpeg (ffmpeg + ffprobe)",
        "Heroicons",
        "Octicons",
        "Silero VAD",
        "WeSpeaker CAM++ (bundled)",
        "GEOS (in shapely)",
        "OpenSSL (in cryptography)",
        "Inter (web font)",
        "JetBrains Mono (web font)",
    ):
        assert expected in components, f"{expected} missing from the SUMMARY table"


def _real_table_lines() -> list[str]:
    """The raw SUMMARY table lines, header and divider excluded."""
    lines = (_ROOT / "build" / "THIRD-PARTY-LICENSES").read_text().splitlines()
    start = licenses._summary_start(lines)
    body: list[str] = []
    for raw in lines[start:]:
        if licenses._RULE_RE.match(raw.strip()):
            break
        if raw.strip() and not raw.strip().startswith("---"):
            body.append(raw)
    return body


def test_real_notice_drops_no_rows():
    """A row whose columns touch (one space) parses as two fields and vanishes."""
    for raw in _real_table_lines():
        parts = licenses._COLUMN_SPLIT_RE.split(raw.strip())
        assert len(parts) >= 3 or len(parts) == 1, f"unparseable table line: {raw!r}"


# Lock entries that never install: platform markers pin them to no real OS.
_NEVER_INSTALLED = {"av", "opencv-python", "qtpy"}


def _normalize(name: str) -> str:
    return re.sub(r"[-_.]+", "-", name).lower()


def test_every_locked_dependency_is_attributed():
    """Any package in the runtime dependency closure needs a SUMMARY row.

    Walks uv.lock from the project's own dependencies, ignoring markers so every
    platform's packages count. Extra rows (fonts, models, executables) are fine;
    only lock -> table is asserted.
    """
    lock = tomllib.loads((_ROOT / "uv.lock").read_text())
    packages = {pkg["name"]: pkg for pkg in lock["package"]}
    closure: set[str] = set()
    stack = [dep["name"] for dep in packages["clipgen"].get("dependencies", [])]
    while stack:
        name = stack.pop()
        if name in closure or name in _NEVER_INSTALLED:
            continue
        closure.add(name)
        stack.extend(dep["name"] for dep in packages[name].get("dependencies", []))

    attributed = {_normalize(row["component"]) for row in licenses.load_components()}
    missing = sorted(name for name in closure if _normalize(name) not in attributed)
    assert not missing, f"unattributed in THIRD-PARTY-LICENSES SUMMARY: {missing}"


def test_real_notice_groups_follow_file_order():
    rows = licenses.load_components()
    groups: list[str] = []
    for row in rows:
        if not row["nested"] and (not groups or groups[-1] != row["group"]):
            groups.append(row["group"])
    # Each license family appears exactly once: the table is clustered, so the
    # frontend can open a heading on change instead of re-sorting.
    assert len(groups) == len(set(groups)), groups


def test_real_notice_keeps_the_wrapped_cv2_cell_whole():
    rows = licenses.load_components()
    dll = next(r for r in rows if r["component"] == "FFmpeg DLL (in cv2)")
    assert dll["nested"] is True
    assert dll["license"].endswith("without FFmpeg)")
