"""Integration tests for Start-overlay endpoints on the combined Flask app.

Exercises the routes registered directly on ``combined`` (not on a blueprint):

* ``GET  /api/changelog``
* ``GET  /api/dirs``           ``POST /api/dirs``
* ``POST /api/folder-picker``
* ``POST /api/sessions/record``
* ``GET  /api/spreadsheets/google``
* ``POST /api/spreadsheets/google/auth``

We build the live app via :func:`server.build_combined_app` and drive it
through ``test_client``. State globals (``_google_auth``, ``config.INPUT_DIR``,
etc.) are reset around each test so order doesn't matter.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytest.importorskip("flask")

import config
import server
import start_settings


@pytest.fixture
def app(monkeypatch, tmp_path):
    """Combined Flask app with the worksheet unset and dirs pinned to tmp."""
    in_dir = tmp_path / "in"
    out_dir = tmp_path / "out"
    in_dir.mkdir()
    out_dir.mkdir()

    # Pin config dirs so /api/dirs sees a known starting point.
    monkeypatch.setattr(config, "INPUT_DIR", str(in_dir), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(out_dir), raising=False)

    # Persist Start settings to tmp so tests don't touch the user's real file.
    monkeypatch.setattr(
        start_settings, "_settings_path", lambda: tmp_path / "start.json"
    )

    # Reset Google-auth state between tests.
    monkeypatch.setattr(server, "_google_auth", server._GoogleAuthState())

    # /api/spreadsheets/open reassigns these module globals directly (no
    # fixture owns them), so an opened workbook would otherwise outlive the
    # test and feed its participants to anything later in the same worker that
    # reads the live sheet. monkeypatch restores them however the route left them.
    monkeypatch.setattr(server, "_worksheet", None)
    monkeypatch.setattr(server, "_sheet_context", None)
    monkeypatch.setattr(server, "_sheet_payload_cache", None)
    monkeypatch.setattr(server, "_active_sheet_meta", None)

    return server.build_combined_app(worksheet=None, default_page="studio")


@pytest.fixture
def client(app):
    return app.test_client()


# ---------- /api/changelog --------------------------------------------------


def test_changelog_returns_entries(client):
    """Happy-path: real CHANGELOG.md parses into a non-empty entries list."""
    resp = client.get("/api/changelog")
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["ok"] is True
    assert isinstance(body["entries"], list)
    # The repo ships a populated CHANGELOG.md; if entries is empty, the
    # parser regex drifted from the file format.
    assert body["entries"], "CHANGELOG.md present but no entries parsed"
    first = body["entries"][0]
    assert {"version", "date", "tool", "title", "body"} <= set(first)


def test_changelog_handles_missing_file(client, monkeypatch):
    """A missing CHANGELOG.md returns ok with an empty list (warning logged)."""
    import changelog

    monkeypatch.setattr(changelog, "_changelog_path", lambda: Path("/nope/missing.md"))
    resp = client.get("/api/changelog")
    assert resp.status_code == 200
    assert resp.get_json() == {"ok": True, "entries": []}


# ---------- /api/dirs (GET + POST) ------------------------------------------


def test_dirs_get_returns_current(client, tmp_path):
    resp = client.get("/api/dirs")
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["ok"] is True
    assert body["input"] == str(tmp_path / "in")
    assert body["output"] == str(tmp_path / "out")


def test_dirs_post_updates_both(client, tmp_path):
    new_in = tmp_path / "alt_in"
    new_in.mkdir()
    new_out = tmp_path / "alt_out"  # doesn't exist yet — POST should mkdir
    resp = client.post(
        "/api/dirs",
        data=json.dumps({"input": str(new_in), "output": str(new_out)}),
        content_type="application/json",
    )
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["ok"] is True
    assert body["input"] == str(new_in)
    assert body["output"] == str(new_out)
    assert new_out.is_dir()  # Output dir auto-created


def test_dirs_post_rejects_nonexistent_input(client, tmp_path):
    bogus = tmp_path / "definitely_not_here"
    resp = client.post(
        "/api/dirs",
        data=json.dumps({"input": str(bogus)}),
        content_type="application/json",
    )
    assert resp.status_code == 400
    body = resp.get_json()
    assert body["ok"] is False
    assert "input" in body["errors"]


def test_dirs_post_rejects_output_under_file(client, tmp_path):
    """If the path points at an existing file (not a dir), mkdir fails."""
    blocker = tmp_path / "blocker"
    blocker.write_text("not a directory")
    resp = client.post(
        "/api/dirs",
        data=json.dumps({"output": str(blocker / "child")}),
        content_type="application/json",
    )
    assert resp.status_code == 400
    body = resp.get_json()
    assert body["ok"] is False
    assert "output" in body["errors"]


# ---------- /api/folder-picker ---------------------------------------------


def test_folder_picker_returns_chosen_path(client, monkeypatch):
    """The route returns whatever utils.open_native_folder_picker returns."""
    import utils

    monkeypatch.setattr(
        utils, "open_native_folder_picker", lambda initial="": "/picked"
    )
    resp = client.post(
        "/api/folder-picker",
        data=json.dumps({"initial": "/start"}),
        content_type="application/json",
    )
    assert resp.status_code == 200
    assert resp.get_json() == {"ok": True, "path": "/picked"}


def test_folder_picker_returns_null_on_cancel(client, monkeypatch):
    import utils

    monkeypatch.setattr(utils, "open_native_folder_picker", lambda initial="": None)
    resp = client.post(
        "/api/folder-picker",
        data=json.dumps({}),
        content_type="application/json",
    )
    assert resp.status_code == 200
    assert resp.get_json() == {"ok": True, "path": None}


# ---------- /api/sessions/record -------------------------------------------


def test_sessions_record_happy_path(client, tmp_path):
    payload = {
        "input": str(tmp_path / "in"),
        "output": str(tmp_path / "out"),
        "spreadsheet": {
            "type": "excel",
            "id_or_path": str(tmp_path / "book.xlsx"),
            "label": "book.xlsx",
        },
    }
    resp = client.post(
        "/api/sessions/record",
        data=json.dumps(payload),
        content_type="application/json",
    )
    assert resp.status_code == 200
    assert resp.get_json() == {"ok": True}
    saved = start_settings.load_start_settings()
    assert saved["recent_projects"], "Session record didn't persist"
    entry = saved["recent_projects"][0]
    assert entry["input"] == payload["input"]
    assert entry["output"] == payload["output"]
    assert entry["spreadsheet"]["type"] == "excel"


def test_sessions_record_no_sheet_path(client, tmp_path):
    """Spreadsheet omitted → still records a session with spreadsheet=None."""
    resp = client.post(
        "/api/sessions/record",
        data=json.dumps(
            {"input": str(tmp_path / "in"), "output": str(tmp_path / "out")}
        ),
        content_type="application/json",
    )
    assert resp.status_code == 200
    saved = start_settings.load_start_settings()
    assert saved["recent_projects"][0]["spreadsheet"] is None


def test_sessions_record_stores_the_project_name(client, tmp_path):
    resp = client.post(
        "/api/sessions/record",
        data=json.dumps(
            {
                "input": str(tmp_path / "in"),
                "output": str(tmp_path / "out"),
                "name": "  Coffee machine study  ",
            }
        ),
        content_type="application/json",
    )
    assert resp.status_code == 200
    saved = start_settings.load_start_settings()
    assert saved["recent_projects"][0]["name"] == "Coffee machine study"


def test_sessions_record_rejects_non_string_name(client, tmp_path):
    resp = client.post(
        "/api/sessions/record",
        data=json.dumps(
            {"input": str(tmp_path / "in"), "output": str(tmp_path / "out"), "name": 7}
        ),
        content_type="application/json",
    )
    assert resp.status_code == 400
    assert "name" in resp.get_json()["error"]


def test_sessions_record_omitted_name_keeps_the_stored_one(client, tmp_path):
    payload = {"input": str(tmp_path / "in"), "output": str(tmp_path / "out")}
    client.post(
        "/api/sessions/record",
        data=json.dumps({**payload, "name": "Kept"}),
        content_type="application/json",
    )
    client.post(
        "/api/sessions/record",
        data=json.dumps(payload),
        content_type="application/json",
    )
    saved = start_settings.load_start_settings()
    assert saved["recent_projects"][0]["name"] == "Kept"


def test_sessions_record_rejects_non_string_input(client):
    resp = client.post(
        "/api/sessions/record",
        data=json.dumps({"input": 42, "output": "/some/where"}),
        content_type="application/json",
    )
    assert resp.status_code == 400
    body = resp.get_json()
    assert body["ok"] is False
    assert "input" in body["error"]


def test_sessions_record_rejects_bad_spreadsheet_type(client, tmp_path):
    resp = client.post(
        "/api/sessions/record",
        data=json.dumps(
            {
                "input": str(tmp_path / "in"),
                "output": str(tmp_path / "out"),
                "spreadsheet": {"type": "csv", "id_or_path": "/x"},
            }
        ),
        content_type="application/json",
    )
    assert resp.status_code == 400
    body = resp.get_json()
    assert body["ok"] is False
    assert "google" in body["error"] and "excel" in body["error"]


def test_sessions_record_rejects_non_dict_spreadsheet(client, tmp_path):
    resp = client.post(
        "/api/sessions/record",
        data=json.dumps(
            {
                "input": str(tmp_path / "in"),
                "output": str(tmp_path / "out"),
                "spreadsheet": "not-a-dict",
            }
        ),
        content_type="application/json",
    )
    assert resp.status_code == 400
    body = resp.get_json()
    assert body["ok"] is False


def test_sessions_record_rejects_missing_spreadsheet_id(client, tmp_path):
    resp = client.post(
        "/api/sessions/record",
        data=json.dumps(
            {
                "input": str(tmp_path / "in"),
                "output": str(tmp_path / "out"),
                "spreadsheet": {"type": "google", "id_or_path": "  "},
            }
        ),
        content_type="application/json",
    )
    assert resp.status_code == 400
    body = resp.get_json()
    assert body["ok"] is False
    assert "id_or_path" in body["error"]


# ---------- /api/spreadsheets/excel + /worksheets --------------------------


def _write_workbook(path, titles):
    import openpyxl

    wb = openpyxl.Workbook()
    wb.active.title = titles[0]
    for t in titles[1:]:
        wb.create_sheet(t)
    wb.save(path)
    wb.close()


def test_excel_list_includes_modified(client, tmp_path):
    """Each .xlsx entry carries a numeric file-modified time for the picker."""
    _write_workbook(tmp_path / "in" / "study.xlsx", ["Sheet1"])
    resp = client.get("/api/spreadsheets/excel")
    body = resp.get_json()
    assert body["ok"] is True
    names = [f["name"] for f in body["files"]]
    assert "study.xlsx" in names
    entry = next(f for f in body["files"] if f["name"] == "study.xlsx")
    assert isinstance(entry["modified"], (int, float))
    assert entry["modified"] > 0


def test_worksheets_route_excel(client, tmp_path):
    """The worksheets route lists an .xlsx's tabs + the priority recommendation."""
    wb_path = tmp_path / "in" / "multi.xlsx"
    _write_workbook(wb_path, ["Intro", "Data", "Extra"])
    resp = client.get(
        "/api/spreadsheets/worksheets",
        query_string={"type": "excel", "id_or_path": str(wb_path)},
    )
    body = resp.get_json()
    assert body["ok"] is True
    assert body["worksheets"] == ["Intro", "Data", "Extra"]
    assert body["recommended"] == "Data"


def test_worksheets_route_rejects_bad_type(client):
    resp = client.get(
        "/api/spreadsheets/worksheets",
        query_string={"type": "bogus", "id_or_path": "x"},
    )
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False


def test_worksheets_route_google_requires_auth(client):
    resp = client.get(
        "/api/spreadsheets/worksheets",
        query_string={"type": "google", "id_or_path": "My Study"},
    )
    assert resp.get_json()["ok"] is False


def _write_preview_workbook(path, participants, filename_row=None):
    """Write a one-tab .xlsx with *participants* and an optional Filename row.

    ``filename_row`` maps a participant id to its Filename-row override cell.
    """
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Study"
    ws["A2"] = "ID"
    ws["D2"] = "Observation"
    ws["E2"] = "Category"
    for i, pid in enumerate(participants):
        ws.cell(row=2, column=2 + i, value=pid)
    if filename_row:
        ws["A3"] = "Filename"
        for i, pid in enumerate(participants):
            override = filename_row.get(pid)
            if override:
                ws.cell(row=3, column=2 + i, value=override)
    wb.save(path)
    wb.close()


def test_preview_reports_expected_filenames_and_disk_status(client, tmp_path):
    """Each participant gets its expected name plus whether it is on disk."""
    wb_path = tmp_path / "in" / "preview.xlsx"
    _write_preview_workbook(wb_path, ["P01", "P02"])
    (tmp_path / "in" / "study_P01.mp4").write_bytes(b"")

    resp = client.get(
        "/api/spreadsheets/preview",
        query_string={
            "type": "excel",
            "id_or_path": str(wb_path),
            "input_dir": str(tmp_path / "in"),
        },
    )
    body = resp.get_json()
    assert body["ok"] is True
    assert body["study"] == "study"
    assert body["participants"] == [
        {
            "id": "P01",
            "filenames": ["study_P01.mp4"],
            "found": True,
            "override": False,
        },
        {
            "id": "P02",
            "filenames": ["study_P02.mp4"],
            "found": False,
            "override": False,
        },
    ]


def test_preview_honours_the_filename_row_override(client, tmp_path):
    """A plus-separated Filename cell previews both parts, in order."""
    wb_path = tmp_path / "in" / "override.xlsx"
    _write_preview_workbook(
        wb_path, ["P01"], filename_row={"P01": "morning.mp4 + afternoon.mp4"}
    )
    (tmp_path / "in" / "morning.mp4").write_bytes(b"")

    body = client.get(
        "/api/spreadsheets/preview",
        query_string={
            "type": "excel",
            "id_or_path": str(wb_path),
            "input_dir": str(tmp_path / "in"),
        },
    ).get_json()
    entry = body["participants"][0]
    assert entry["filenames"] == ["morning.mp4", "afternoon.mp4"]
    assert entry["override"] is True
    assert entry["found"] is False  # afternoon.mp4 is missing


def test_preview_does_not_swap_the_active_sheet(client, tmp_path):
    """Previewing is read-only: no sheet is loaded as a side effect."""
    wb_path = tmp_path / "in" / "preview.xlsx"
    _write_preview_workbook(wb_path, ["P01"])
    assert client.get("/api/status").get_json()["sheet_loaded"] is False

    client.get(
        "/api/spreadsheets/preview",
        query_string={"type": "excel", "id_or_path": str(wb_path)},
    )
    assert client.get("/api/status").get_json()["sheet_loaded"] is False


def test_preview_reports_an_unreadable_worksheet(client, tmp_path):
    """A sheet with no clipgen headers fails with guidance, not a stack trace."""
    import openpyxl

    wb_path = tmp_path / "in" / "empty.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = "nothing useful"
    wb.active["A2"] = "still nothing"
    wb.save(wb_path)
    wb.close()

    resp = client.get(
        "/api/spreadsheets/preview",
        query_string={"type": "excel", "id_or_path": str(wb_path)},
    )
    body = resp.get_json()
    assert body["ok"] is False
    assert "Observation" in body["error"]


def test_preview_rejects_bad_type(client):
    resp = client.get(
        "/api/spreadsheets/preview",
        query_string={"type": "bogus", "id_or_path": "x"},
    )
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False


def test_preview_google_requires_auth(client):
    resp = client.get(
        "/api/spreadsheets/preview",
        query_string={"type": "google", "id_or_path": "My Study"},
    )
    assert resp.get_json()["ok"] is False


def _write_valid_workbook(path, tabs):
    """Write an .xlsx whose every tab has a minimal valid clipgen layout."""
    import openpyxl

    wb = openpyxl.Workbook()
    for i, (title, pid) in enumerate(tabs):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = title
        ws["A1"] = "Study"
        ws["A2"] = "ID"
        ws["B2"] = pid
        ws["D2"] = "Observation"
        ws["E2"] = "Category"
        ws["A3"] = "1"
        ws["B3"] = "00:10-00:20"
        ws["D3"] = "Obs"
        ws["E3"] = "CatA"
    wb.save(path)
    wb.close()


def test_open_excel_with_worksheet_reflected_in_status(client, tmp_path):
    """Opening a chosen tab loads it and surfaces it via /api/status."""
    wb_path = tmp_path / "in" / "multi.xlsx"
    _write_valid_workbook(
        wb_path, [("Intro", "P01"), ("Data", "P02"), ("Extra", "P03")]
    )
    resp = client.post(
        "/api/spreadsheets/open",
        json={"type": "excel", "id_or_path": str(wb_path), "worksheet": "Extra"},
    )
    assert resp.get_json()["ok"] is True
    st = client.get("/api/status").get_json()
    assert st["spreadsheet_worksheet"] == "Extra"
    assert st["spreadsheet_label"].endswith("(Extra)")
    # Confirm the Extra tab's data actually loaded (its participant is P03).
    sheet = client.get("/studio/api/sheet").get_json()
    assert sheet["participants"] == ["P03"]


def test_open_excel_records_the_project_name(client, tmp_path):
    """The Start overlay rides project_name along with the sheet open."""
    wb_path = tmp_path / "in" / "named.xlsx"
    _write_valid_workbook(wb_path, [("Data", "P01")])
    resp = client.post(
        "/api/spreadsheets/open",
        json={
            "type": "excel",
            "id_or_path": str(wb_path),
            "project_name": "Coffee machine study",
        },
    )
    assert resp.get_json()["ok"] is True
    saved = start_settings.load_start_settings()
    assert saved["recent_projects"][0]["name"] == "Coffee machine study"


def test_open_excel_rejects_non_string_project_name(client, tmp_path):
    wb_path = tmp_path / "in" / "named.xlsx"
    _write_valid_workbook(wb_path, [("Data", "P01")])
    resp = client.post(
        "/api/spreadsheets/open",
        json={"type": "excel", "id_or_path": str(wb_path), "project_name": 7},
    )
    assert resp.status_code == 400
    assert "project_name" in resp.get_json()["error"]


def test_open_excel_without_worksheet_uses_priority_tab(client, tmp_path):
    """No worksheet field → priority auto-pick ("Data"), not the first tab."""
    wb_path = tmp_path / "in" / "multi.xlsx"
    _write_valid_workbook(
        wb_path, [("Intro", "P01"), ("Data", "P02"), ("Extra", "P03")]
    )
    resp = client.post(
        "/api/spreadsheets/open",
        json={"type": "excel", "id_or_path": str(wb_path)},
    )
    assert resp.get_json()["ok"] is True
    st = client.get("/api/status").get_json()
    assert st["spreadsheet_worksheet"] == "Data"


# ---------- /api/spreadsheets/google + /auth -------------------------------


def test_google_list_unauthenticated(client):
    """No cached client → authenticated=False, no in-flight flag."""
    resp = client.get("/api/spreadsheets/google")
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["ok"] is True
    assert body["authenticated"] is False
    assert body["auth_in_flight"] is False
    assert body["sheets"] == []


def test_google_list_authenticated(client, monkeypatch):
    """Sheet metadata (incl. modifiedTime) comes from get_all_spreadsheet_meta."""
    server._google_auth.client = object()  # stand-in
    import google_api

    monkeypatch.setattr(
        google_api,
        "get_all_spreadsheet_meta",
        lambda _c: [
            {"name": "Alpha", "id": "a", "modifiedTime": "2026-07-20T10:00:00Z"},
            {"name": "Beta", "id": "b", "modifiedTime": ""},
        ],
    )
    resp = client.get("/api/spreadsheets/google")
    body = resp.get_json()
    assert body["authenticated"] is True
    assert [s["name"] for s in body["sheets"]] == ["Alpha", "Beta"]
    # id stays the name (open-by-name contract), modifiedTime flows through.
    assert body["sheets"][0]["id"] == "Alpha"
    assert body["sheets"][0]["modifiedTime"] == "2026-07-20T10:00:00Z"


def test_google_list_surfaces_api_error(client, monkeypatch):
    server._google_auth.client = object()
    import google_api

    def _boom(_c):
        raise RuntimeError("rate limited")

    monkeypatch.setattr(google_api, "get_all_spreadsheet_meta", _boom)
    resp = client.get("/api/spreadsheets/google")
    body = resp.get_json()
    assert body["authenticated"] is True
    assert "rate limited" in body["auth_error"]
    assert body["sheets"] == []


def test_google_auth_short_circuits_when_already_authenticated(client):
    """POST /auth returns authenticated=True without spawning a thread."""
    server._google_auth.client = object()
    resp = client.post("/api/spreadsheets/google/auth", json={})
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["ok"] is True
    assert body["started"] is False
    assert body["authenticated"] is True


def test_google_auth_short_circuits_when_in_flight(client):
    server._google_auth.in_flight = True
    resp = client.post("/api/spreadsheets/google/auth", json={})
    body = resp.get_json()
    assert body["started"] is False
    assert body["in_flight"] is True


def test_google_auth_runs_handler_when_idle(client, monkeypatch):
    """When no client and no flight, the thread runs cli.authenticate_google."""
    import threading

    # Force the spawned thread to run inline so we can assert end state.
    def fake_thread(target, daemon=True):
        target()

        class _Joined:
            def start(self_inner):
                pass

        return _Joined()

    monkeypatch.setattr(threading, "Thread", fake_thread)

    sentinel = object()
    import cli

    monkeypatch.setattr(cli, "authenticate_google", lambda: sentinel)

    resp = client.post("/api/spreadsheets/google/auth", json={})
    assert resp.status_code == 202
    body = resp.get_json()
    assert body["started"] is True
    # After the (synchronous) thread ran, state should reflect success.
    assert server._google_auth.client is sentinel
    assert server._google_auth.in_flight is False
    assert server._google_auth.error == ""


def test_google_auth_records_thread_error(client, monkeypatch):
    import threading

    def fake_thread(target, daemon=True):
        target()

        class _Joined:
            def start(self_inner):
                pass

        return _Joined()

    monkeypatch.setattr(threading, "Thread", fake_thread)

    import cli

    def _explode():
        raise RuntimeError("creds missing")

    monkeypatch.setattr(cli, "authenticate_google", _explode)

    client.post("/api/spreadsheets/google/auth", json={})
    assert server._google_auth.client is None
    assert "creds missing" in server._google_auth.error
    assert server._google_auth.in_flight is False


# ---------- sheet-switch guard during active generation ---------------------


def test_spreadsheets_open_rejected_during_generation(client, monkeypatch):
    """Switching spreadsheets is rejected with 409 while a clip generation is
    in progress, so the generated lists are not rebound under an active stream."""
    monkeypatch.setattr(server, "_generate_in_progress", True)
    resp = client.post(
        "/api/spreadsheets/open",
        json={"type": "excel", "id_or_path": "/tmp/whatever.xlsx"},
    )
    assert resp.status_code == 409
    data = resp.get_json()
    assert data["ok"] is False
    assert "in progress" in data["error"]


def test_spreadsheets_open_rejected_during_intake(client, monkeypatch):
    """An in-flight intake stream also blocks a spreadsheet switch."""
    monkeypatch.setattr(server, "_intake_active", 1)
    resp = client.post(
        "/api/spreadsheets/open",
        json={"type": "excel", "id_or_path": "/tmp/whatever.xlsx"},
    )
    assert resp.status_code == 409
    assert resp.get_json()["ok"] is False


def test_spreadsheets_open_rejected_during_timeline_viewer(client, monkeypatch):
    """A timeline-viewer build blocks a spreadsheet switch: it appends into the
    shared generated list/manifest, so a swap mid-build would rebind those under
    it and mix old-sheet artifacts into the new sheet."""
    monkeypatch.setattr(server, "_timeline_viewer_in_progress", True)
    resp = client.post(
        "/api/spreadsheets/open",
        json={"type": "excel", "id_or_path": "/tmp/whatever.xlsx"},
    )
    assert resp.status_code == 409
    assert resp.get_json()["ok"] is False


def test_spreadsheets_open_rejected_during_gallery(client, monkeypatch):
    """A gallery build also blocks a spreadsheet switch."""
    monkeypatch.setattr(server, "_gallery_in_progress", True)
    resp = client.post(
        "/api/spreadsheets/open",
        json={"type": "excel", "id_or_path": "/tmp/whatever.xlsx"},
    )
    assert resp.status_code == 409
    assert resp.get_json()["ok"] is False


def test_spreadsheets_close_rejected_during_reel(client, monkeypatch):
    """Closing the spreadsheet is rejected with 409 while a reel build runs."""
    monkeypatch.setattr(server, "_reel_in_progress", True)
    resp = client.post("/api/spreadsheets/close")
    assert resp.status_code == 409
    assert resp.get_json()["ok"] is False


def test_combined_server_forces_non_interactive(monkeypatch):
    """The web server has no console: start_combined_server must flip
    utils.NO_INPUT_MODE on so a missing source video is skipped-and-reported
    instead of blocking a Flask/daemon thread on input() (regression: watch-dir
    triggered runs + Studio generate hung on the fuzzy-match prompt)."""
    import utils

    class _FakeApp:
        def run(self, **kwargs):  # never actually serve
            pass

    monkeypatch.setattr(utils, "NO_INPUT_MODE", False)
    monkeypatch.setattr(server, "build_combined_app", lambda **kwargs: _FakeApp())
    monkeypatch.setattr(server.webbrowser, "open", lambda *a, **k: None)

    server.start_combined_server(port=0)
    assert utils.NO_INPUT_MODE is True
