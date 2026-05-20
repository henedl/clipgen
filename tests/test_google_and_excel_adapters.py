from typing import cast

import gspread
import google_api
import excel_io


def test_get_worksheet_prefers_priority_match(monkeypatch):
    class FakeWorksheet:
        def __init__(self, title):
            self.title = title

    class FakeSpreadsheet:
        def __init__(self):
            self._worksheets = [FakeWorksheet("Other"), FakeWorksheet("Sheet1")]

        def worksheets(self):
            return self._worksheets

        def worksheet(self, title):
            for ws in self._worksheets:
                if ws.title == title:
                    return ws
            raise KeyError(title)

    fake_spreadsheet = FakeSpreadsheet()
    ws = google_api.get_worksheet(cast(gspread.Spreadsheet, fake_spreadsheet))
    assert ws.title == "Sheet1"


def test_find_spreadsheet_by_name_and_guess(monkeypatch):
    monkeypatch.setattr(google_api.config, "DEBUGGING", False, raising=False)
    docs = ["Playtest Data Set", "Other Sheet"]

    idx_exact = google_api.find_spreadsheet_by_name("Other Sheet", docs)
    assert idx_exact == 1

    idx_guess = google_api.find_spreadsheet_by_name("Playtest", docs)
    assert idx_guess == 0


def test_find_spreadsheet_by_name_prefers_exact_over_guess(monkeypatch):
    monkeypatch.setattr(google_api.config, "DEBUGGING", False, raising=False)
    docs = ["Playtest Data Set", "Playtest"]

    idx = google_api.find_spreadsheet_by_name("Playtest", docs)
    assert idx == 1

    docs_reversed = ["Playtest", "Playtest Data Set"]
    idx_reversed = google_api.find_spreadsheet_by_name("Playtest", docs_reversed)
    assert idx_reversed == 0


def test_get_sheet_values_retries_on_rate_limit(monkeypatch):
    class FakeAPIError(gspread.exceptions.APIError):
        def __init__(self, status_code: int):
            self.response = type("R", (), {"status_code": status_code})()

        def __str__(self) -> str:
            return f"HTTP {self.response.status_code}"

    calls = {"n": 0}

    class FakeSheet:
        def get_all_values(self):
            calls["n"] += 1
            if calls["n"] == 1:
                raise FakeAPIError(429)
            return [["Study"], ["ID", "P01"]]

    monkeypatch.setattr(google_api.time, "sleep", lambda _s: None)
    assert google_api.get_sheet_values(FakeSheet()) == [
        ["Study"],
        ["ID", "P01"],
    ]
    assert calls["n"] == 2


def test_get_sheet_values_does_not_retry_client_errors():
    class FakeAPIError(gspread.exceptions.APIError):
        def __init__(self, status_code: int):
            self.response = type("R", (), {"status_code": status_code})()

        def __str__(self) -> str:
            return f"HTTP {self.response.status_code}"

    calls = {"n": 0}

    class FakeSheet:
        def get_all_values(self):
            calls["n"] += 1
            raise FakeAPIError(403)

    import pytest

    with pytest.raises(FakeAPIError):
        google_api.get_sheet_values(FakeSheet())
    assert calls["n"] == 1


def test_excel_sheet_adapter_basic_access(tmp_path, monkeypatch):
    # Build a tiny workbook with one sheet and some data.
    import openpyxl

    wb_path = tmp_path / "test.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Study"
    ws["A2"] = "ID"
    ws["B2"] = "P01"
    ws["D2"] = "Observation"
    ws["E2"] = "Category"
    ws["A3"] = "1"
    ws["B3"] = "00:10-00:20"
    ws["D3"] = "Obs one"
    ws["E3"] = "CatA"
    wb.save(wb_path)
    wb.close()

    adapter = excel_io.open_excel_workbook(str(wb_path))
    assert adapter is not None

    # Adapter should expose spreadsheet-like metadata.
    assert adapter.spreadsheet.title == "test"
    assert adapter.spreadsheet.url is None

    # Data access helpers.
    all_values = adapter.get_all_values()
    assert all_values[0][0] == "Study"

    id_cell = adapter.find("ID")
    assert id_cell is not None
    assert id_cell.row == 2

    row2 = adapter.row_values(2)
    assert "P01" in row2
    assert adapter.col_count >= 5


def _make_workbook(path):
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "ID"
    wb.save(path)
    wb.close()


def test_prompt_for_excel_fallback_picks_by_index(tmp_path, monkeypatch):
    wb_path = tmp_path / "alpha.xlsx"
    _make_workbook(wb_path)
    monkeypatch.chdir(tmp_path)

    inputs = iter(["1"])
    monkeypatch.setattr(excel_io.utils, "read_user_input", lambda _prompt: next(inputs))

    adapter = excel_io.prompt_for_excel_fallback()
    assert adapter is not None
    assert adapter.spreadsheet.title == "alpha"


def test_prompt_for_excel_fallback_accepts_path(tmp_path, monkeypatch):
    # No .xlsx in cwd — user pastes an absolute path.
    cwd = tmp_path / "empty"
    cwd.mkdir()
    elsewhere = tmp_path / "data" / "other.xlsx"
    elsewhere.parent.mkdir()
    _make_workbook(elsewhere)
    monkeypatch.chdir(cwd)

    inputs = iter([str(elsewhere)])
    monkeypatch.setattr(excel_io.utils, "read_user_input", lambda _prompt: next(inputs))

    adapter = excel_io.prompt_for_excel_fallback()
    assert adapter is not None
    assert adapter.spreadsheet.title == "other"


def test_prompt_for_excel_fallback_cancel_returns_none(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    inputs = iter([""])
    monkeypatch.setattr(excel_io.utils, "read_user_input", lambda _prompt: next(inputs))

    assert excel_io.prompt_for_excel_fallback() is None


def test_prompt_for_excel_fallback_help_then_cancel(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    inputs = iter(["help", ""])
    monkeypatch.setattr(excel_io.utils, "read_user_input", lambda _prompt: next(inputs))

    # 'help' should not exit the loop; subsequent empty input cancels.
    assert excel_io.prompt_for_excel_fallback() is None
