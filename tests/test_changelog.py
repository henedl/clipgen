"""Tests for the CHANGELOG.md parser used by the Start overlay."""

import pytest

import changelog


@pytest.fixture
def changelog_file(tmp_path, monkeypatch):
    target = tmp_path / "CHANGELOG.md"
    monkeypatch.setattr(changelog, "_changelog_path", lambda: target)
    return target


def test_load_returns_empty_when_file_missing(changelog_file):
    assert changelog.load_entries() == []


def test_load_parses_releases_with_several_changes(changelog_file):
    changelog_file.write_text(
        "# Changelog\n"
        "\n"
        "## v0.16.0 — 2026-05-16\n"
        "**Core:** Feat: Refreshed Start screen.\n"
        "**Studio:** Fix: Cells drag onto the Artifact intake again.\n"
        "\n"
        "## v0.15.0 — 2026-05-03\n"
        "**Transcripts:** Feat: Cancel a summary partway through.\n",
        encoding="utf-8",
    )
    entries = changelog.load_entries()
    assert len(entries) == 2

    first = entries[0]
    assert first["version"] == "v0.16.0"
    assert first["date"] == "2026-05-16"
    assert first["changes"] == [
        {"tool": "Core", "kind": "Feat", "text": "Refreshed Start screen."},
        {
            "tool": "Studio",
            "kind": "Fix",
            "text": "Cells drag onto the Artifact intake again.",
        },
    ]
    assert entries[1]["changes"][0]["tool"] == "Transcripts"


def test_heading_separator_requires_surrounding_whitespace(changelog_file):
    """The date's own hyphens must not read as the version/date separator.

    This is the bug that silently emptied the panel when the heading format
    changed: ``## v0.16.0 — 2026-08-20`` parsed as date ``2026``, tool ``20``.
    """
    changelog_file.write_text(
        "## v0.16.0 — 2026-08-20\n**Core:** Feat: Something.\n",
        encoding="utf-8",
    )
    entry = changelog.load_entries()[0]
    assert entry["version"] == "v0.16.0"
    assert entry["date"] == "2026-08-20"


def test_change_line_without_a_kind_still_parses(changelog_file):
    changelog_file.write_text(
        "## v0.1.0 — 2026-01-01\n**Core:** Something happened.\n",
        encoding="utf-8",
    )
    change = changelog.load_entries()[0]["changes"][0]
    assert change == {"tool": "Core", "kind": "", "text": "Something happened."}


def test_release_with_no_change_lines_is_dropped(changelog_file):
    """An empty card reads as a rendering bug, so it never reaches the page."""
    changelog_file.write_text(
        "## v0.2.0 — 2026-01-02\n"
        "\n"
        "## v0.1.0 — 2026-01-01\n"
        "**Core:** Feat: Something.\n",
        encoding="utf-8",
    )
    entries = changelog.load_entries()
    assert [e["version"] for e in entries] == ["v0.1.0"]


def test_load_respects_limit(changelog_file):
    body = "# Changelog\n\n" + "\n".join(
        f"## v0.0.{i} — 2026-01-{i:02d}\n**Core:** Feat: Entry {i}.\n"
        for i in range(1, 6)
    )
    changelog_file.write_text(body, encoding="utf-8")
    entries = changelog.load_entries(limit=3)
    assert len(entries) == 3
    assert entries[0]["changes"][0]["text"] == "Entry 1."


def test_real_changelog_is_wellformed():
    """Guard the shipped file: every change line must carry a known tool."""
    known = {
        "Core",
        "Studio",
        "Screenspace",
        "Transcripts",
        "Workflows",
        "Composer",
        "Overview",
    }
    entries = changelog.load_entries(limit=999)
    assert entries, "CHANGELOG.md present but no entries parsed"
    changes = [c for e in entries for c in e["changes"]]
    assert changes
    assert {c["tool"] for c in changes} <= known
    assert {c["kind"] for c in changes} == {"Feat", "Fix"}
    # Markdown inside a line would render literally in the Start overlay.
    assert not [c for c in changes if "**" in c["text"] or "`" in c["text"]]
