from types import SimpleNamespace

import spreadsheet
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


def test_get_num_participants_and_participant_list():
    cells = _make_cells()
    sheet_data = _basic_sheet_data()
    header_row = sheet_data[cells.id_cell.row - 1]

    num = spreadsheet.get_num_participants(header_row, cells.id_cell, col_count=len(header_row))
    assert num == 2

    participants = spreadsheet.get_participant_list(header_row, cells.id_cell, num)
    assert participants == ["P01", "P02"]


def test_generate_participant_timestamps_happy_path():
    cells = _make_cells()
    sheet_data = _basic_sheet_data()
    id_cell = cells.id_cell
    observation_cell = cells.observation_cell

    clips_p01 = spreadsheet.generate_participant_timestamps(
        sheet_data, id_cell, observation_cell, study_name="study", participant_id="P01"
    )
    assert len(clips_p01) == 2
    assert {clip["desc"] for clip in clips_p01} == {"Obs one", "Obs three"}
    assert all(clip["participant"] == "P01" for clip in clips_p01)

    clips_p02 = spreadsheet.generate_participant_timestamps(
        sheet_data, id_cell, observation_cell, study_name="study", participant_id="P02"
    )
    assert len(clips_p02) == 2
    assert {clip["desc"] for clip in clips_p02} == {"Obs two", "Obs three"}
    assert all(clip["participant"] == "P02" for clip in clips_p02)


def test_generate_participant_timestamps_missing_participant_returns_empty():
    cells = _make_cells()
    sheet_data = _basic_sheet_data()

    clips = spreadsheet.generate_participant_timestamps(
        sheet_data, cells.id_cell, cells.observation_cell, study_name="study", participant_id="P99"
    )
    assert clips == []


def test_generate_cell_timestamps_valid_and_invalid_specs():
    cells = _make_cells()
    sheet_data = _basic_sheet_data()
    id_cell = cells.id_cell
    observation_cell = cells.observation_cell

    specs = [("P01", 3), ("P02", 4), ("P99", 3), ("P01", 99)]

    clips = spreadsheet.generate_cell_timestamps(
        sheet_data, id_cell, observation_cell, study_name="study", cell_specs=specs
    )

    # Only the two valid specs should produce clips.
    assert len(clips) == 2
    coords = {(clip["cell"].row, clip["cell"].col) for clip in clips}
    assert coords == {(3, 2), (4, 3)}
    assert {clip["participant"] for clip in clips} == {"P01", "P02"}


def test_collect_categories_and_generate_category_timestamps():
    cells = _make_cells()
    sheet_data = _basic_sheet_data()
    id_cell = cells.id_cell
    observation_cell = cells.observation_cell
    category_cell = cells.category_cell

    categories = spreadsheet.collect_categories(sheet_data, id_cell, category_cell)
    assert categories == ["CatA", "CatB"]

    selected = ["CatA"]
    clips = spreadsheet.generate_category_timestamps(
        sheet_data,
        id_cell,
        observation_cell,
        category_cell,
        num_participants=2,
        study_name="study",
        selected_categories=selected,
    )
    # CatA appears on rows 3 and 5 (1-based); both rows have timestamps in at least one participant.
    assert len(clips) == 3
    assert {clip["desc"] for clip in clips} == {"Obs one", "Obs three"}


def test_generate_line_and_range_timestamps_cli_paths():
    cells = _make_cells()
    sheet_data = _basic_sheet_data()
    id_cell = cells.id_cell
    observation_cell = cells.observation_cell

    # Line mode with CLI numbers selects specific rows.
    line_clips = spreadsheet.generate_line_timestamps(
        sheet_data,
        id_cell,
        observation_cell,
        num_participants=2,
        study_name="study",
        cli_line_numbers=[3, 4],
        skip_prompts=True,
    )
    assert {clip["desc"] for clip in line_clips} == {"Obs one", "Obs two"}

    # Range mode uses inclusive 1-based start/end.
    range_clips = spreadsheet.generate_range_timestamps(
        sheet_data,
        id_cell,
        observation_cell,
        num_participants=2,
        study_name="study",
        start_line=3,
        end_line=4,
    )
    assert {clip["desc"] for clip in range_clips} == {"Obs one", "Obs two"}


def test_baseline_and_relative_timestamps_integration(monkeypatch):
    # Sheet layout with a clock baseline row for P01 only.
    # Row 0: Study
    # Row 1: Baseline row (clock times)
    # Row 2: Header row
    # Rows 3+: Data rows
    sheet_data = [
        ["Study"],
        ["", "09:12:00", "", "", ""],
        ["ID", "P01", "P02", "Observation", "Category"],
        ["1", "09:15:00-09:16:30", "", "Obs one", "CatA"],
        ["2", "", "00:20-00:40", "Obs two", "CatB"],
    ]

    id_cell = SimpleNamespace(row=3, col=1)
    observation_cell = SimpleNamespace(row=3, col=4)

    # Simplify annotation parsing so we focus on timestamp + baseline behavior.
    def fake_parse_cell_annotations(value):
        return value, {}, set()

    monkeypatch.setattr(utils, "parse_cell_annotations", fake_parse_cell_annotations)
    monkeypatch.setattr(utils, "has_non_ignored_timestamp_content", lambda _v: True)

    # Two participants (P01 has a baseline; P02 does not).
    clips_p01 = spreadsheet.generate_participant_timestamps(
        sheet_data, id_cell, observation_cell, study_name="study", participant_id="P01"
    )
    clips_p02 = spreadsheet.generate_participant_timestamps(
        sheet_data, id_cell, observation_cell, study_name="study", participant_id="P02"
    )

    prepared_p01 = [files.prepare_clip(clip) for clip in clips_p01]
    prepared_p02 = [files.prepare_clip(clip) for clip in clips_p02]

    assert len(prepared_p01) == 1
    assert len(prepared_p02) == 1

    # P01 clip should use the baseline row (09:12:00) and convert to relative times.
    assert prepared_p01[0]["times"] == [("3:00", "4:30")]

    # P02 clip has no baseline and should remain relative as entered.
    assert prepared_p02[0]["times"] == [("00:20", "00:40")]

