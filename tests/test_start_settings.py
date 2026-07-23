"""Tests for Start overlay persistence helpers (start_settings)."""

import json

import pytest

import start_settings


def test_settings_path_windows_localappdata(monkeypatch, tmp_path):
    monkeypatch.setattr(start_settings.sys, "platform", "win32")
    local = tmp_path / "Local"
    monkeypatch.setenv("LOCALAPPDATA", str(local))
    assert start_settings._settings_path() == local / "clipgen" / "start.json"


def test_settings_path_windows_fallback(monkeypatch, tmp_path):
    monkeypatch.setattr(start_settings.sys, "platform", "win32")
    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    monkeypatch.setattr(start_settings.Path, "home", lambda: tmp_path)
    assert (
        start_settings._settings_path()
        == tmp_path / "AppData" / "Local" / "clipgen" / "start.json"
    )


def test_settings_path_unix(monkeypatch, tmp_path):
    monkeypatch.setattr(start_settings.sys, "platform", "linux")
    monkeypatch.setattr(start_settings.Path, "home", lambda: tmp_path)
    assert (
        start_settings._settings_path()
        == tmp_path / ".config" / "clipgen" / "start.json"
    )


@pytest.fixture
def settings_file(tmp_path, monkeypatch):
    """Redirect _settings_path() to a per-test temp file."""
    target = tmp_path / "start.json"
    monkeypatch.setattr(start_settings, "_settings_path", lambda: target)
    return target


def test_load_returns_defaults_when_missing(settings_file):
    settings = start_settings.load_start_settings()
    assert settings["persist_enabled"] is True
    assert settings["recent_inputs"] == []
    assert settings["last_spreadsheet"] is None


def test_load_recovers_from_corrupt_file(settings_file):
    settings_file.write_text("not json", encoding="utf-8")
    settings = start_settings.load_start_settings()
    assert settings["persist_enabled"] is True
    assert settings["recent_inputs"] == []


def test_load_merges_partial_files(settings_file):
    settings_file.write_text(
        json.dumps({"persist_enabled": False, "last_input": "/x"}), encoding="utf-8"
    )
    settings = start_settings.load_start_settings()
    assert settings["persist_enabled"] is False
    assert settings["last_input"] == "/x"
    assert settings["recent_inputs"] == []


def test_save_and_reload_roundtrip(settings_file):
    payload = {
        "persist_enabled": True,
        "last_input": "/a",
        "last_output": "/b",
        "recent_inputs": ["/a"],
        "recent_outputs": ["/b"],
        "last_spreadsheet": {"type": "excel", "id_or_path": "/x.xlsx", "label": "x"},
        "recent_spreadsheets": [
            {"type": "excel", "id_or_path": "/x.xlsx", "label": "x"}
        ],
        "recent_projects": [],
    }
    start_settings.save_start_settings(payload)
    assert settings_file.is_file()
    reloaded = start_settings.load_start_settings()
    assert reloaded == payload


def test_record_recent_input_prepends_dedupes_caps(settings_file):
    for i in range(10):
        start_settings.record_recent_input(f"/dir{i}")
    settings = start_settings.load_start_settings()
    assert len(settings["recent_inputs"]) == start_settings.RECENTS_CAP
    assert settings["recent_inputs"][0] == "/dir9"
    assert settings["last_input"] == "/dir9"

    # dedupe: re-record an existing entry
    start_settings.record_recent_input("/dir3")
    settings = start_settings.load_start_settings()
    assert settings["recent_inputs"][0] == "/dir3"
    assert settings["recent_inputs"].count("/dir3") == 1


def test_record_recent_output_persists(settings_file):
    start_settings.record_recent_output("/out1")
    start_settings.record_recent_output("/out2")
    settings = start_settings.load_start_settings()
    assert settings["last_output"] == "/out2"
    assert settings["recent_outputs"] == ["/out2", "/out1"]


def test_record_recent_spreadsheet_keyed_on_type_and_id(settings_file):
    start_settings.record_recent_spreadsheet("excel", "/a.xlsx", "A")
    start_settings.record_recent_spreadsheet("google", "MyStudy", "MyStudy")
    start_settings.record_recent_spreadsheet("excel", "/a.xlsx", "A renamed", "Data")
    settings = start_settings.load_start_settings()
    assert len(settings["recent_spreadsheets"]) == 2
    assert settings["recent_spreadsheets"][0] == {
        "type": "excel",
        "id_or_path": "/a.xlsx",
        "label": "A renamed",
        "worksheet": "Data",
    }


def test_persist_disabled_short_circuits_recording(settings_file):
    start_settings.set_persist_enabled(False)
    start_settings.record_recent_input("/should-not-record")
    start_settings.record_recent_output("/should-not-record")
    start_settings.record_recent_spreadsheet("excel", "/x.xlsx", "x")
    start_settings.record_project_session(
        "/in", "/out", {"type": "excel", "id_or_path": "/x.xlsx", "label": "x"}
    )
    settings = start_settings.load_start_settings()
    assert settings["persist_enabled"] is False
    assert settings["recent_inputs"] == []
    assert settings["recent_outputs"] == []
    assert settings["recent_spreadsheets"] == []
    assert settings["recent_projects"] == []


def test_record_project_session_dedupes_on_triple(settings_file):
    sheet = {"type": "excel", "id_or_path": "/x.xlsx", "label": "x"}
    start_settings.record_project_session("/in", "/out", sheet)
    start_settings.record_project_session("/in", "/out", None)
    # Re-recording the same triple moves the entry to the front, not duplicates.
    start_settings.record_project_session("/in", "/out", sheet)
    settings = start_settings.load_start_settings()
    assert len(settings["recent_projects"]) == 2
    head = settings["recent_projects"][0]
    assert head["input"] == "/in"
    assert head["output"] == "/out"
    assert head["spreadsheet"] == sheet
    assert head["last_opened"]  # ISO-8601 string


def test_record_project_session_ignores_blank_dirs(settings_file):
    start_settings.record_project_session("", "/out", None)
    start_settings.record_project_session("/in", "", None)
    settings = start_settings.load_start_settings()
    assert settings["recent_projects"] == []


def test_persist_disabled_still_writes_the_flag(settings_file):
    start_settings.set_persist_enabled(False)
    assert settings_file.is_file()
    body = json.loads(settings_file.read_text(encoding="utf-8"))
    assert body["persist_enabled"] is False


def test_record_ignores_blank_paths(settings_file):
    start_settings.record_recent_input("")
    start_settings.record_recent_output("")
    start_settings.record_recent_spreadsheet("excel", "", "x")
    settings = start_settings.load_start_settings()
    assert settings["recent_inputs"] == []
    assert settings["recent_outputs"] == []
    assert settings["recent_spreadsheets"] == []


def test_record_ignores_unknown_spreadsheet_type(settings_file):
    start_settings.record_recent_spreadsheet("postgres", "abc", "abc")
    settings = start_settings.load_start_settings()
    assert settings["recent_spreadsheets"] == []
