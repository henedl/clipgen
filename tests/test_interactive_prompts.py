"""Tests for interactive.py prompt helpers.

interactive.py was deliberately split out of spreadsheet.py to keep the
prompt logic pure and testable: every function takes a SheetContext and
routes user input through utils.read_user_input(). These tests drive each
prompt with scripted input sequences and assert on the resolved selection.
"""

from types import SimpleNamespace

import google_api
import interactive
import utils
from spreadsheet import SheetContext


def _scripted(monkeypatch, responses):
    """Patch utils.read_user_input to return queued responses in order.

    Raises if a prompt is reached after the script is exhausted so an
    accidental extra prompt surfaces as a test failure, not a hang.
    """
    queue = list(responses)

    def _fake(_prompt=""):
        if not queue:
            raise AssertionError(
                f"read_user_input exhausted; unexpected extra prompt: {_prompt!r}"
            )
        return queue.pop(0)

    monkeypatch.setattr(utils, "read_user_input", _fake)
    return queue


def _sheet_data():
    # Row 0: study name. Row 1: header (row 2). Rows 2-4: data rows.
    return [
        ["Study", "", "", "", "", ""],
        ["ID", "P01", "P02", "Observation", "Category", "Severity"],
        ["1", "00:10-00:20", "", "Obs one", "CatA", "High"],
        ["2", "", "00:30-00:40 !key", "Obs two", "CatB", "Low"],
        ["3", "00:50-01:00", "01:10-01:20", "Obs three", "CatA", "Critical"],
    ]


_UNSET = object()


def _ctx(sheet_data=None, num_participants=2, severity_cell=_UNSET):
    """Build a SheetContext mirroring build_sheet_context output for tests."""
    if sheet_data is None:
        sheet_data = _sheet_data()
    sev = SimpleNamespace(row=2, col=6) if severity_cell is _UNSET else severity_cell
    return SheetContext(
        sheet_data=sheet_data,
        id_cell=SimpleNamespace(row=2, col=1),
        observation_cell=SimpleNamespace(row=2, col=4),
        category_cell=SimpleNamespace(row=2, col=5),
        num_participants=num_participants,
        study_name="study",
        severity_cell=sev,
    )


# ---- prompt_batch_confirm ----


def test_prompt_batch_confirm_yes(monkeypatch):
    _scripted(monkeypatch, ["y"])
    assert interactive.prompt_batch_confirm(_ctx()) is True


def test_prompt_batch_confirm_no(monkeypatch):
    _scripted(monkeypatch, ["n"])
    assert interactive.prompt_batch_confirm(_ctx()) is False


# ---- prompt_multi_selection (generic engine) ----


def _run_multi(items):
    """Invoke prompt_multi_selection with fixed display strings."""
    return interactive.prompt_multi_selection(
        items,
        header="Items:",
        prompt_text="pick",
        confirm_label="Selected:",
        no_match_msg="no match",
    )


def test_prompt_multi_selection_by_index(monkeypatch):
    _scripted(monkeypatch, ["1,3", "y"])
    assert _run_multi(["alpha", "beta", "gamma"]) == ["alpha", "gamma"]


def test_prompt_multi_selection_all_keyword(monkeypatch):
    _scripted(monkeypatch, ["all"])
    assert _run_multi(["a", "b"]) == ["a", "b"]


def test_prompt_multi_selection_by_text_match(monkeypatch):
    _scripted(monkeypatch, ["beta", "y"])
    assert _run_multi(["alpha", "beta"]) == ["beta"]


def test_prompt_multi_selection_dedupes_repeated_index(monkeypatch):
    _scripted(monkeypatch, ["1,1,2", "y"])
    assert _run_multi(["a", "b"]) == ["a", "b"]


def test_prompt_multi_selection_reprompts_after_all_invalid(monkeypatch):
    # First selection is entirely out of range -> reprompt, then a valid pick.
    _scripted(monkeypatch, ["9", "2", "y"])
    assert _run_multi(["a", "b", "c"]) == ["b"]


def test_prompt_multi_selection_declined_confirmation_reselects(monkeypatch):
    # Pick index 1, decline confirmation, then pick index 2 and confirm.
    _scripted(monkeypatch, ["1", "n", "2", "y"])
    assert _run_multi(["a", "b"]) == ["b"]


# ---- prompt_category_selection ----


def test_prompt_category_selection_happy(monkeypatch):
    _scripted(monkeypatch, ["1", "y"])
    assert interactive.prompt_category_selection(_ctx()) == ["CatA"]


def test_prompt_category_selection_all(monkeypatch):
    _scripted(monkeypatch, ["all"])
    assert interactive.prompt_category_selection(_ctx()) == ["CatA", "CatB"]


def test_prompt_category_selection_none_when_no_categories(monkeypatch):
    sheet = [
        ["Study"],
        ["ID", "P01", "P02", "Observation", "Category"],
        ["1", "00:10", "", "Obs one", ""],
    ]
    _scripted(monkeypatch, [])  # no prompt should be reached
    assert interactive.prompt_category_selection(_ctx(sheet_data=sheet)) is None


# ---- prompt_severity_selection ----


def test_prompt_severity_selection_happy(monkeypatch):
    monkeypatch.setattr(utils, "_use_rich", lambda: False)
    _scripted(monkeypatch, ["all"])
    result = interactive.prompt_severity_selection(_ctx())
    assert result is not None
    assert set(result) == {"Critical", "High", "Low"}


def test_prompt_severity_selection_by_name(monkeypatch):
    monkeypatch.setattr(utils, "_use_rich", lambda: False)
    _scripted(monkeypatch, ["high", "y"])
    assert interactive.prompt_severity_selection(_ctx()) == ["High"]


def test_prompt_severity_selection_none_without_severity_column(monkeypatch):
    _scripted(monkeypatch, [])
    assert interactive.prompt_severity_selection(_ctx(severity_cell=None)) is None


# ---- prompt_line_selection ----


def test_prompt_line_selection_happy(monkeypatch):
    _scripted(monkeypatch, ["3,4", "y"])
    assert interactive.prompt_line_selection(_ctx()) == [3, 4]


def test_prompt_line_selection_reprompts_on_non_integer(monkeypatch):
    _scripted(monkeypatch, ["abc", "3", "y"])
    assert interactive.prompt_line_selection(_ctx()) == [3]


def test_prompt_line_selection_filters_out_of_range(monkeypatch):
    # Row 99 is out of range; row 3 is valid.
    _scripted(monkeypatch, ["99,3", "y"])
    assert interactive.prompt_line_selection(_ctx()) == [3]


# ---- prompt_range_selection ----


def test_prompt_range_selection_happy(monkeypatch):
    _scripted(monkeypatch, ["3", "4", "y"])
    assert interactive.prompt_range_selection(_ctx()) == (3, 4)


def test_prompt_range_selection_reprompts_on_non_integer(monkeypatch):
    _scripted(monkeypatch, ["x", "3", "3", "4", "y"])
    assert interactive.prompt_range_selection(_ctx()) == (3, 4)


def test_prompt_range_selection_reprompts_on_inverted_range(monkeypatch):
    # start > end is rejected by _validate_row_range, then a valid range.
    _scripted(monkeypatch, ["4", "3", "3", "4", "y"])
    assert interactive.prompt_range_selection(_ctx()) == (3, 4)


# ---- prompt_cell_selection ----


def test_prompt_cell_selection_happy(monkeypatch):
    _scripted(monkeypatch, ["P01.3", "y"])
    assert interactive.prompt_cell_selection(_ctx()) == [("P01", 3)]


def test_prompt_cell_selection_multiple_specs(monkeypatch):
    _scripted(monkeypatch, ["P01.3 + P02.4", "y"])
    assert interactive.prompt_cell_selection(_ctx()) == [("P01", 3), ("P02", 4)]


def test_prompt_cell_selection_filters_unknown_participant(monkeypatch):
    # P99 is not a real participant; only the valid spec survives.
    _scripted(monkeypatch, ["P99.3 + P01.3", "y"])
    assert interactive.prompt_cell_selection(_ctx()) == [("P01", 3)]


def test_prompt_cell_selection_reprompts_on_empty_input(monkeypatch):
    _scripted(monkeypatch, ["", "P01.3", "y"])
    assert interactive.prompt_cell_selection(_ctx()) == [("P01", 3)]


# ---- prompt_participant_selection ----


def test_prompt_participant_selection_by_number(monkeypatch):
    _scripted(monkeypatch, ["1", "y"])
    assert interactive.prompt_participant_selection(_ctx()) == ["P01"]


def test_prompt_participant_selection_by_id(monkeypatch):
    _scripted(monkeypatch, ["P02", "y"])
    assert interactive.prompt_participant_selection(_ctx()) == ["P02"]


def test_prompt_participant_selection_dedupes_number_and_id(monkeypatch):
    # "1" and "P01" refer to the same participant; result must be deduped.
    _scripted(monkeypatch, ["1, P01", "y"])
    assert interactive.prompt_participant_selection(_ctx()) == ["P01"]


def test_prompt_participant_selection_none_when_no_participants(monkeypatch):
    _scripted(monkeypatch, [])
    assert interactive.prompt_participant_selection(_ctx(num_participants=0)) is None


# ---- prompt_keyword_selection ----


def test_prompt_keyword_selection_single_annotation_confirmed(monkeypatch):
    _scripted(monkeypatch, ["y"])
    assert interactive.prompt_keyword_selection(_ctx()) == ["key"]


def test_prompt_keyword_selection_single_annotation_declined(monkeypatch):
    _scripted(monkeypatch, ["n"])
    assert interactive.prompt_keyword_selection(_ctx()) is None


def test_prompt_keyword_selection_none_when_no_annotations(monkeypatch):
    sheet = _sheet_data()
    sheet[3][2] = "00:30-00:40"  # remove the only !key annotation
    _scripted(monkeypatch, [])
    assert interactive.prompt_keyword_selection(_ctx(sheet_data=sheet)) is None


# ---- browse_spreadsheet (smoke) ----


def test_browse_spreadsheet_loads_and_quits(monkeypatch):
    """Browse mode loads the sheet, renders the first page, and exits on quit.

    Non-tty stdin routes _read_browse_key() through utils.read_user_input,
    so a scripted "quit" drives one full loop iteration.
    """

    class _FakeSheet:
        def get_all_values(self):
            return _sheet_data()

        spreadsheet = SimpleNamespace(title="study", url=None)

    monkeypatch.setattr(google_api, "get_sheet_values", lambda s: s.get_all_values())
    monkeypatch.setattr(utils, "use_progress", lambda: False)
    _scripted(monkeypatch, ["quit"])
    # Returns cleanly without raising.
    interactive.browse_spreadsheet(_FakeSheet())


def test_browse_spreadsheet_returns_early_on_missing_headers(monkeypatch):
    class _FakeSheet:
        def get_all_values(self):
            return [["Study"], ["ID", "P01"]]  # missing Observation/Category

        spreadsheet = SimpleNamespace(title="study", url=None)

    monkeypatch.setattr(google_api, "get_sheet_values", lambda s: s.get_all_values())
    monkeypatch.setattr(utils, "use_progress", lambda: False)
    _scripted(monkeypatch, [])  # loop must not be reached
    interactive.browse_spreadsheet(_FakeSheet())
