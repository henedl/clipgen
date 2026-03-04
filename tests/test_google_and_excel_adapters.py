from types import SimpleNamespace

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
    ws = google_api.get_worksheet(fake_spreadsheet)
    assert ws.title == "Sheet1"


def test_find_spreadsheet_by_name_and_guess(monkeypatch):
    monkeypatch.setattr(google_api.config, "DEBUGGING", False, raising=False)
    docs = ["Playtest Data Set", "Other Sheet"]

    idx_exact = google_api.find_spreadsheet_by_name("Other Sheet", docs)
    assert idx_exact == 1

    idx_guess = google_api.find_spreadsheet_by_name("Playtest", docs)
    assert idx_guess == 0


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

