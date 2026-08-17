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
        "window": None,
        "remember_window": True,
        "filename_overrides": {},
    }
    start_settings.save_start_settings(payload)
    assert settings_file.is_file()
    reloaded = start_settings.load_start_settings()
    assert reloaded == payload


def test_record_recent_input_prepends_dedupes_caps(settings_file):
    overflow = start_settings.RECENTS_CAP + 2
    for i in range(overflow):
        start_settings.record_recent_input(f"/dir{i}")
    settings = start_settings.load_start_settings()
    assert len(settings["recent_inputs"]) == start_settings.RECENTS_CAP
    assert settings["recent_inputs"][0] == f"/dir{overflow - 1}"
    assert settings["last_input"] == f"/dir{overflow - 1}"

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


def test_record_project_session_stores_and_trims_the_name(settings_file):
    start_settings.record_project_session("/in", "/out", None, name="  My study  ")
    settings = start_settings.load_start_settings()
    assert settings["recent_projects"][0]["name"] == "My study"


def test_record_project_session_name_is_tri_state(settings_file):
    start_settings.record_project_session("/in", "/out", None, name="My study")
    # None means "keep whatever is stored" — the CLI-launch and Studio
    # sheet-switch call sites pass nothing and must not wipe the label.
    start_settings.record_project_session("/in", "/out", None)
    assert start_settings.load_start_settings()["recent_projects"][0]["name"] == (
        "My study"
    )
    # An explicit empty string clears it.
    start_settings.record_project_session("/in", "/out", None, name="")
    assert start_settings.load_start_settings()["recent_projects"][0]["name"] == ""


def test_record_project_session_name_is_not_part_of_the_identity(settings_file):
    start_settings.record_project_session("/in", "/out", None, name="First")
    start_settings.record_project_session("/in", "/out", None, name="Renamed")
    settings = start_settings.load_start_settings()
    assert len(settings["recent_projects"]) == 1
    assert settings["recent_projects"][0]["name"] == "Renamed"


def test_record_project_session_name_does_not_leak_across_projects(settings_file):
    start_settings.record_project_session("/in", "/out", None, name="Named")
    start_settings.record_project_session("/other", "/out", None)
    settings = start_settings.load_start_settings()
    assert settings["recent_projects"][0]["input"] == "/other"
    assert settings["recent_projects"][0]["name"] == ""


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


def test_window_geometry_defaults_to_none(settings_file):
    assert start_settings.load_window_geometry() is None


def test_window_geometry_roundtrip(settings_file):
    start_settings.record_window_geometry(120, 64, 1600, 1000)
    assert start_settings.load_window_geometry() == {
        "x": 120,
        "y": 64,
        "width": 1600,
        "height": 1000,
    }


def test_window_geometry_rejects_malformed_entries(settings_file):
    for window in ({"x": 0, "y": 0, "width": 800}, {"x": "0"}, [], "1440x900"):
        settings_file.write_text(json.dumps({"window": window}), encoding="utf-8")
        assert start_settings.load_window_geometry() is None


def test_window_geometry_ignores_nonpositive_size(settings_file):
    start_settings.record_window_geometry(0, 0, 0, 900)
    assert start_settings.load_window_geometry() is None


def test_record_window_geometry_respects_its_own_toggle(settings_file):
    start_settings.set_remember_window(False)
    start_settings.record_window_geometry(10, 20, 1200, 800)
    assert start_settings.load_window_geometry() is None


def test_window_geometry_is_independent_of_persist_enabled(settings_file):
    """The two toggles gate different things; persist_enabled is project history."""
    start_settings.set_persist_enabled(False)
    start_settings.record_window_geometry(10, 20, 1200, 800)
    assert start_settings.load_window_geometry() == {
        "x": 10,
        "y": 20,
        "width": 1200,
        "height": 800,
    }


def test_disabling_remember_window_drops_the_stored_rect(settings_file):
    start_settings.record_window_geometry(10, 20, 1200, 800)
    start_settings.set_remember_window(False)
    assert start_settings.load_window_geometry() is None
    # The flag itself is still written, so the toggle survives the session.
    assert start_settings.load_start_settings()["remember_window"] is False


def test_clear_window_geometry_works_with_persist_disabled(settings_file):
    start_settings.record_window_geometry(10, 20, 1200, 800)
    start_settings.set_persist_enabled(False)
    start_settings.clear_window_geometry()
    assert start_settings.load_window_geometry() is None


# ---- Filename overrides ----


def test_filename_override_roundtrips(settings_file):
    start_settings.set_filename_override(
        "excel", "/study.xlsx", "Data", "P01", "morning.mp4 + afternoon.mp4"
    )
    assert start_settings.filename_overrides("excel", "/study.xlsx", "Data") == {
        "P01": "morning.mp4 + afternoon.mp4"
    }


def test_filename_overrides_are_per_worksheet(settings_file):
    """One workbook can hold several studies; P01 means a different person in each."""
    start_settings.set_filename_override(
        "excel", "/study.xlsx", "Wave1", "P01", "a.mp4"
    )
    start_settings.set_filename_override(
        "excel", "/study.xlsx", "Wave2", "P01", "b.mp4"
    )

    assert start_settings.filename_overrides("excel", "/study.xlsx", "Wave1") == {
        "P01": "a.mp4"
    }
    assert start_settings.filename_overrides("excel", "/study.xlsx", "Wave2") == {
        "P01": "b.mp4"
    }


def test_clearing_a_filename_override_drops_the_source_entry(settings_file):
    start_settings.set_filename_override(
        "mindnode", "/map.mindnode", "", "P01", "a.mp4"
    )
    remaining = start_settings.set_filename_override(
        "mindnode", "/map.mindnode", "", "P01", ""
    )
    assert remaining == {}
    assert start_settings.filename_overrides("mindnode", "/map.mindnode", "") == {}
    # Emptied sources are removed rather than left as {} husks.
    assert start_settings.load_start_settings()["filename_overrides"] == {}


def test_filename_overrides_are_independent_of_persist_enabled(settings_file):
    """An override is configuration, not project history — see set_filename_override."""
    start_settings.set_persist_enabled(False)
    start_settings.set_filename_override("excel", "/study.xlsx", "Data", "P01", "a.mp4")
    assert start_settings.filename_overrides("excel", "/study.xlsx", "Data") == {
        "P01": "a.mp4"
    }
