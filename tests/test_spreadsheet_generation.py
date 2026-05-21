from types import SimpleNamespace

import spreadsheet
from spreadsheet import SheetContext
import files
import utils


def _make_cells():
    # Header row is spreadsheet row 2 in this program.
    return SimpleNamespace(
        id_cell=SimpleNamespace(row=2, col=1),
        observation_cell=SimpleNamespace(row=2, col=4),
        category_cell=SimpleNamespace(row=2, col=5),
    )


def _basic_sheet_data():
    # Row 0: study name
    # Row 1: header row (ID, P01, P02, Observation, Category)
    # Rows 2+: data rows
    return [
        ["Study"],
        ["ID", "P01", "P02", "Observation", "Category"],
        ["1", "00:10-00:20", "", "Obs one", "CatA"],
        ["2", "", "00:30-00:40", "Obs two", "CatB"],
        ["3", "00:50-01:00", "01:10-01:20", "Obs three", "CatA"],
    ]


def _make_context(
    sheet_data=None,
    cells=None,
    num_participants=2,
    baseline_row_idx=None,
    filename_row_idx=None,
):
    """Helper to build a SheetContext for tests."""
    if sheet_data is None:
        sheet_data = _basic_sheet_data()
    if cells is None:
        cells = _make_cells()
    return SheetContext(
        sheet_data=sheet_data,
        id_cell=cells.id_cell,
        observation_cell=cells.observation_cell,
        category_cell=cells.category_cell,
        num_participants=num_participants,
        study_name="study",
        baseline_row_idx=baseline_row_idx,
        filename_row_idx=filename_row_idx,
    )


def test_get_num_participants_and_participant_list():
    cells = _make_cells()
    sheet_data = _basic_sheet_data()
    header_row = sheet_data[cells.id_cell.row - 1]

    num = spreadsheet.get_num_participants(header_row, cells.id_cell, len(header_row))
    assert num == 2

    participants = spreadsheet.get_participant_list(header_row, cells.id_cell, num)
    assert participants == ["P01", "P02"]


def test_get_num_participants_ragged_header_row():
    """col_count may exceed header_row length (build_sheet_context derives it
    from the longest row in the sheet). The scan must not IndexError."""
    cells = _make_cells()
    header_row = ["ID", "P01"]

    num = spreadsheet.get_num_participants(header_row, cells.id_cell, 5)
    assert num == 1


def test_get_num_participants_with_observation_before_id():
    """Reference clipgen-test.xlsx layout: 'Observation' header appears in a
    row above (and to the left of) the ID/participants header row. Because
    _find_in_data returns the first match anywhere, observation_cell.col can
    be smaller than id_cell.col. The participant scan must remain layout-
    agnostic and still find every P*/G* column after ID."""
    id_cell = SimpleNamespace(row=2, col=6)
    header_row = ["", "", "", "", "", "ID", "P01", "P02", "P03", "P04"]

    num = spreadsheet.get_num_participants(header_row, id_cell, len(header_row))
    assert num == 4


def test_get_num_participants_handles_variable_column_layouts():
    """Different studies arrange columns differently. The count must reflect
    only the headers actually starting with P/G — never assume a fixed total
    or that participants are sandwiched between specific columns."""
    cases = [
        (["ID", "P01"], 1, 1),
        (["ID", "P01", "P02", "P03"], 1, 3),
        (["ID", "Notes", "P01", "Tags", "G01", "G02"], 1, 3),
        (["Notes", "ID"], 2, 0),
    ]
    for header_row, id_col, expected in cases:
        id_cell = SimpleNamespace(row=1, col=id_col)
        num = spreadsheet.get_num_participants(header_row, id_cell, len(header_row))
        assert num == expected, f"Expected {expected} for {header_row!r}, got {num}"


def test_build_sheet_context_with_observation_in_earlier_row():
    """End-to-end guard for the clipgen-test.xlsx layout that regressed
    in commit a9f66fb: 'Observation' appears in a header section several
    rows above (and to the left of) the row that holds 'ID' and the
    participant columns. build_sheet_context must still detect every P*
    column rather than aborting with 'No participant columns found'."""
    sheet_data = [
        ["Study", "", "", "", "", "", "", "", ""],
        ["intro", "", "", "", "", "ID", "P01", "P02", "P03"],
        ["names", "", "", "", "", "Metadata", "Alice", "Bob", "Carol"],
        ["Count", "Reported", "Severity", "Category", "Observation", "Summary"],
        ["1", "", "", "CatA", "Obs one", "00:10-00:20"],
    ]

    class _Sheet:
        def get_all_values(self):
            return sheet_data

        spreadsheet = SimpleNamespace(title="study")

    ctx = spreadsheet.build_sheet_context(_Sheet())
    assert ctx is not None
    assert ctx.num_participants == 3
    assert ctx.id_cell.col == 6
    assert ctx.observation_cell.col == 5


def test_collect_categories_skips_ragged_rows():
    sheet_data = [
        ["Study"],
        ["ID", "P01", "P02", "Observation", "Category"],
        ["1", "00:10", "", "Obs one", "CatA"],
        ["2", "00:20"],  # row missing Category column
        ["3", "00:30", "", "Obs three", "CatB"],
    ]
    ctx = _make_context(sheet_data=sheet_data)

    categories = spreadsheet.collect_categories(ctx)
    assert categories == ["CatA", "CatB"]


def test_generate_participant_timestamps_happy_path():
    ctx = _make_context()

    clips_p01 = spreadsheet.generate_participant_timestamps(ctx, "P01")
    assert len(clips_p01) == 2
    assert {clip["desc"] for clip in clips_p01} == {"Obs one", "Obs three"}
    assert all(clip["participant"] == "P01" for clip in clips_p01)

    clips_p02 = spreadsheet.generate_participant_timestamps(ctx, "P02")
    assert len(clips_p02) == 2
    assert {clip["desc"] for clip in clips_p02} == {"Obs two", "Obs three"}
    assert all(clip["participant"] == "P02" for clip in clips_p02)


def test_generate_participant_timestamps_missing_participant_returns_empty():
    ctx = _make_context()

    clips = spreadsheet.generate_participant_timestamps(ctx, "P99")
    assert clips == []


def test_generate_cell_timestamps_valid_and_invalid_specs():
    ctx = _make_context()

    specs = [("P01", 3), ("P02", 4), ("P99", 3), ("P01", 99)]

    clips = spreadsheet.generate_cell_timestamps(ctx, specs)

    # Only the two valid specs should produce clips.
    assert len(clips) == 2
    coords = {(clip["cell"].row, clip["cell"].col) for clip in clips}
    assert coords == {(3, 2), (4, 3)}
    assert {clip["participant"] for clip in clips} == {"P01", "P02"}


def test_collect_categories_and_generate_category_timestamps():
    ctx = _make_context()

    categories = spreadsheet.collect_categories(ctx)
    assert categories == ["CatA", "CatB"]

    selected = ["CatA"]
    clips = spreadsheet.generate_category_timestamps(ctx, selected)
    # CatA appears on rows 3 and 5 (1-based); both rows have timestamps in at least one participant.
    assert len(clips) == 3
    assert {clip["desc"] for clip in clips} == {"Obs one", "Obs three"}


def test_generate_line_and_range_timestamps_cli_paths():
    ctx = _make_context()

    # Line mode with CLI numbers selects specific rows.
    line_clips = spreadsheet.generate_line_timestamps(ctx, [3, 4])
    assert {clip["desc"] for clip in line_clips} == {"Obs one", "Obs two"}

    # Range mode uses inclusive 1-based start/end.
    range_clips = spreadsheet.generate_range_timestamps(ctx, 3, 4)
    assert {clip["desc"] for clip in range_clips} == {"Obs one", "Obs two"}


def test_baseline_and_relative_timestamps_integration(monkeypatch):
    # Sheet layout with a clock baseline row for P01 only.
    # Row 0: Study
    # Row 1: Baseline row (clock times)
    # Row 2: Header row
    # Rows 3+: Data rows
    sheet_data = [
        ["Study"],
        ["Baseline time", "09:12:00", "", "", ""],
        ["ID", "P01", "P02", "Observation", "Category"],
        ["1", "09:15:00-09:16:30", "", "Obs one", "CatA"],
        ["2", "", "00:20-00:40", "Obs two", "CatB"],
    ]

    cells = SimpleNamespace(
        id_cell=SimpleNamespace(row=3, col=1),
        observation_cell=SimpleNamespace(row=3, col=4),
        category_cell=SimpleNamespace(row=3, col=5),
    )
    baseline_row_idx = spreadsheet._detect_baseline_row(sheet_data)

    ctx = _make_context(
        sheet_data=sheet_data,
        cells=cells,
        num_participants=2,
        baseline_row_idx=baseline_row_idx,
    )

    # Simplify annotation parsing so we focus on timestamp + baseline behavior.
    def fake_parse_cell_annotations(value):
        return value, {}, set()

    monkeypatch.setattr(utils, "parse_cell_annotations", fake_parse_cell_annotations)
    monkeypatch.setattr(utils, "has_non_ignored_timestamp_content", lambda _v: True)

    # Two participants (P01 has a baseline; P02 does not).
    clips_p01 = spreadsheet.generate_participant_timestamps(ctx, "P01")
    clips_p02 = spreadsheet.generate_participant_timestamps(ctx, "P02")

    prepared_p01 = [files.prepare_clip(clip) for clip in clips_p01]
    prepared_p02 = [files.prepare_clip(clip) for clip in clips_p02]

    assert len(prepared_p01) == 1
    assert len(prepared_p02) == 1

    # P01 clip should use the baseline row (09:12:00) and convert to relative times.
    assert prepared_p01[0]["times"] == [("0:03:00", "0:04:30")]

    # P02 clip has no baseline and should remain relative as entered.
    assert prepared_p02[0]["times"] == [("00:20", "00:40")]


# ---- Local header lookup tests ----


def test_find_in_data_finds_exact_match():
    data = [["Study"], ["ID", "P01", "P02", "Observation", "Category"]]
    result = spreadsheet._find_in_data(data, "Observation")
    assert result is not None
    assert result.row == 2
    assert result.col == 4


def test_find_in_data_returns_none_for_missing():
    data = [["ID", "P01"]]
    assert spreadsheet._find_in_data(data, "Missing") is None


def test_validate_headers_with_local_data():
    data = [["Study"], ["ID", "P01", "P02", "Observation", "Category"]]
    result = spreadsheet.validate_spreadsheet_headers(data)
    assert result is not None
    id_cell, obs_cell, cat_cell = result
    assert id_cell.row == 2 and id_cell.col == 1
    assert obs_cell.row == 2 and obs_cell.col == 4
    assert cat_cell.row == 2 and cat_cell.col == 5


def test_validate_headers_missing_required(capsys):
    data = [["Study"], ["ID", "P01", "P02"]]
    result = spreadsheet.validate_spreadsheet_headers(data)
    assert result is None
