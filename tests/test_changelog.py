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


def test_load_parses_curated_entries(changelog_file):
    changelog_file.write_text(
        "# Changelog\n"
        "\n"
        "## v0.10.145 — 2026-05-16 — Core\n"
        "**Refreshed Start screen**\n"
        "Two-column launcher with animated intro.\n"
        "\n"
        "## v0.10.140 — 2026-05-03 — Studio\n"
        "**Drag timestamp cells**\n"
        "Cells drag onto the Artifact intake.\n",
        encoding="utf-8",
    )
    entries = changelog.load_entries()
    assert len(entries) == 2
    first = entries[0]
    assert first["version"] == "v0.10.145"
    assert first["date"] == "2026-05-16"
    assert first["tool"] == "Core"
    assert first["title"] == "Refreshed Start screen"
    assert "Two-column" in first["body"]
    assert entries[1]["tool"] == "Studio"


def test_load_respects_limit(changelog_file):
    body = "# Changelog\n\n" + "\n".join(
        f"## v0.0.{i} — 2026-01-{i:02d} — Core\n**Entry {i}**\nBody {i}.\n"
        for i in range(1, 6)
    )
    changelog_file.write_text(body, encoding="utf-8")
    entries = changelog.load_entries(limit=3)
    assert len(entries) == 3
    assert entries[0]["title"] == "Entry 1"
