import pytest

import spreadsheet
from types import SimpleNamespace

from conftest import make_sheet_context as _make_context
from utils import ClipRecord


def test_parse_reel_input_parses_mixed_selectors():
    parsed = spreadsheet.parse_reel_input(
        'chronologic, 11, 13-16, P01.5, P02, "Usability"'
    )
    assert parsed["chronologic"] is True
    assert parsed["lines"] == [11]
    assert parsed["ranges"] == [(13, 16)]
    assert parsed["cells"] == [("P01", 5)]
    assert parsed["participants"] == ["P02"]
    assert parsed["categories"] == ["Usability"]


def test_parse_reel_input_handles_quoted_categories_and_dedupes():
    parsed = spreadsheet.parse_reel_input(
        '"Usability, Round 2", 11, 11, 13-16, 13-16, P01.5, P01.5, P02, P02'
    )
    assert parsed["categories"] == ["Usability, Round 2"]
    assert parsed["lines"] == [11]
    assert parsed["ranges"] == [(13, 16)]
    assert parsed["cells"] == [("P01", 5)]
    assert parsed["participants"] == ["P02"]


def test_parse_reel_input_ignores_malformed_tokens():
    parsed = spreadsheet.parse_reel_input('PXX.abc, not-a-time, "Valid Category"')
    assert parsed["lines"] == []
    assert parsed["ranges"] == []
    assert parsed["cells"] == []
    assert parsed["participants"] == []
    assert parsed["categories"] == ["Valid Category"]


def test_parse_reel_input_parses_highlights():
    parsed = spreadsheet.parse_reel_input("highlights")
    assert parsed["highlights"] is True
    assert parsed["batch"] is False
    assert parsed["keyword"] is False


def test_parse_reel_input_highlights_with_selectors():
    parsed = spreadsheet.parse_reel_input("highlights, P01, 11")
    assert parsed["highlights"] is True
    assert parsed["participants"] == ["P01"]
    assert parsed["lines"] == [11]


def test_parse_cell_specifications_parses_multiple_cells():
    assert spreadsheet.parse_cell_specifications("P01.11 + P03.9") == [
        ("P01", 11),
        ("P03", 9),
    ]


def test_parse_cell_specifications_rejects_non_numeric_row():
    with pytest.raises(ValueError, match=r'Invalid row number "abc"'):
        spreadsheet.parse_cell_specifications("P01.abc")


def test_parse_cell_specifications_rejects_non_positive_row():
    with pytest.raises(ValueError, match="Row number must be positive"):
        spreadsheet.parse_cell_specifications("P01.0")
    with pytest.raises(ValueError, match="Row number must be positive"):
        spreadsheet.parse_cell_specifications("P01.-3")


def test_parse_cell_specifications_rejects_missing_dot():
    with pytest.raises(ValueError, match="Invalid cell specification"):
        spreadsheet.parse_cell_specifications("P0111")


def test_parse_cell_specifications_rejects_bad_participant_prefix():
    with pytest.raises(ValueError, match="Invalid participant ID"):
        spreadsheet.parse_cell_specifications("X01.5")


def test_detect_mode_from_input_rejects_mixed_types():
    mode, kwargs = spreadsheet.detect_mode_from_input("5, P01.11")
    assert mode is None
    assert kwargs == {}


def test_detect_mode_from_input_single_range():
    mode, kwargs = spreadsheet.detect_mode_from_input("10-20")
    assert mode == "range"
    assert kwargs == {"range_start": 10, "range_end": 20}


def test_detect_mode_from_input_cell_list():
    mode, kwargs = spreadsheet.detect_mode_from_input("P01.11, P02.12")
    assert mode == "cell"
    assert kwargs == {"cell_specs": [("P01", 11), ("P02", 12)]}


def test_detect_mode_from_input_participant_list():
    mode, kwargs = spreadsheet.detect_mode_from_input("P01, P02")
    assert mode == "participant"
    assert kwargs == {"participant_id": "P01,P02"}


def test_detect_mode_from_input_multiple_ranges_rejected():
    mode, kwargs = spreadsheet.detect_mode_from_input("10-20, 30-40")
    assert mode is None
    assert kwargs == {}


def test_detect_mode_from_input_cells_and_participants_rejected():
    mode, kwargs = spreadsheet.detect_mode_from_input("P01.11, P02")
    assert mode is None
    assert kwargs == {}


def test_generate_reel_timestamps_chronologic_requires_one_participant(fake_sheet_meta):
    ctx = _make_context(
        sheet_data=[["Study"]],
        id_cell=fake_sheet_meta.id_cell,
        observation_cell=fake_sheet_meta.observation_cell,
        category_cell=fake_sheet_meta.category_cell,
        num_participants=1,
        study_name="study",
    )
    clips = spreadsheet.generate_reel_timestamps(ctx, "chronologic")
    assert clips == []


def test_make_clip_record_attaches_timestamp_baseline(fake_sheet_meta):
    # Sheet layout:
    # Row 0: Study
    # Row 1: Baseline row (clock times)
    # Row 2: ID header row
    # Row 3+: Data rows
    sheet_data = [
        ["Study"],
        ["Baseline time", "09:12:00", "", ""],
        ["ID", "P01", "Observation", "Category"],
        ["1", "09:15:00-09:16:30", "Obs one", "CatA"],
    ]
    # For this sheet layout, the header row is spreadsheet row 3.
    id_cell = SimpleNamespace(row=3, col=1)
    observation_cell = SimpleNamespace(row=3, col=3)
    category_cell = SimpleNamespace(row=3, col=4)
    baseline_row_idx = spreadsheet._detect_baseline_row(sheet_data)

    ctx = _make_context(
        sheet_data=sheet_data,
        id_cell=id_cell,
        observation_cell=observation_cell,
        category_cell=category_cell,
        num_participants=1,
        study_name="study",
        baseline_row_idx=baseline_row_idx,
    )

    clip = spreadsheet._make_clip_record(
        ctx, row_idx=3, col_idx=1, cell_value=sheet_data[3][1]
    )
    assert clip["participant"] == "P01"
    assert clip.get("timestamp_baseline") == "09:12:00"


def test_generate_reel_timestamps_preserves_input_cell_order(
    monkeypatch, fake_sheet_meta, make_clip
):
    """Explicit cell selectors compose in input order, not row-major.

    Studio's reel button sends cells in panel/drag order; the composed reel
    must match that order so the on-screen card sequence equals the output
    sequence. Previously the function force-sorted by (row, col), which broke
    the Studio UX whenever the user rearranged cards or queued cells from
    non-adjacent rows.
    """
    # Three cells: caller sends row 7 first, then row 3, then row 5. The
    # composed reel should match — not (3, 5, 7).
    clip_r7 = make_clip(row=7, col=2)
    clip_r3 = make_clip(row=3, col=2)
    clip_r5 = make_clip(row=5, col=2)

    monkeypatch.setattr(
        spreadsheet,
        "parse_reel_input",
        lambda _input: {
            "batch": False,
            "keyword": False,
            "chronologic": False,
            "highlights": False,
            "severity": False,
            "lines": [],
            "ranges": [],
            "categories": [],
            "cells": [("P01", 7), ("P01", 3), ("P01", 5)],
            "participants": [],
        },
    )
    monkeypatch.setattr(
        spreadsheet,
        "generate_cell_timestamps",
        lambda ctx, cells: [clip_r7, clip_r3, clip_r5],
    )

    ctx = _make_context(
        sheet_data=[["Study"]],
        id_cell=fake_sheet_meta.id_cell,
        observation_cell=fake_sheet_meta.observation_cell,
        category_cell=fake_sheet_meta.category_cell,
        num_participants=1,
        study_name="study",
    )

    clips = spreadsheet.generate_reel_timestamps(ctx, "P01.7, P01.3, P01.5")
    assert [c["cell"].row for c in clips] == [7, 3, 5]


def test_generate_reel_timestamps_dedupes_cells(
    monkeypatch, fake_sheet_meta, make_clip
):
    duplicate_clip = make_clip(row=4, col=2)

    monkeypatch.setattr(
        spreadsheet,
        "parse_reel_input",
        lambda _input: {
            "batch": False,
            "keyword": True,
            "chronologic": False,
            "highlights": False,
            "lines": [4],
            "ranges": [],
            "categories": [],
            "cells": [],
            "participants": [],
        },
    )
    monkeypatch.setattr(
        spreadsheet, "generate_keyword_timestamps", lambda ctx: [duplicate_clip]
    )
    monkeypatch.setattr(
        spreadsheet, "generate_line_timestamps", lambda ctx, lines: [duplicate_clip]
    )

    ctx = _make_context(
        sheet_data=[["Study"]],
        id_cell=fake_sheet_meta.id_cell,
        observation_cell=fake_sheet_meta.observation_cell,
        category_cell=fake_sheet_meta.category_cell,
        num_participants=1,
        study_name="study",
    )
    clips = spreadsheet.generate_reel_timestamps(ctx, "keyword, 4")

    assert len(clips) == 1
    assert (clips[0]["cell"].row, clips[0]["cell"].col) == (4, 2)


def test_generate_keyword_timestamps_honors_header_and_segment_annotations(
    monkeypatch, fake_sheet_meta
):
    cells = fake_sheet_meta
    sheet_data = [
        ["Study"],
        ["ID", "P01 !key", "P02", "Observation", "Category"],
        ["1", "00:10-00:20 !key", "", "Obs one", "CatA"],
        ["2", "", "00:30-00:40 !key", "Obs two", "CatB"],
    ]

    # Use real batch generation so filter logic can inspect header and cell annotations.
    monkeypatch.setattr(
        spreadsheet, "get_num_participants", lambda header_row, _id, _col_count: 2
    )

    ctx = _make_context(
        sheet_data=sheet_data,
        id_cell=cells.id_cell,
        observation_cell=cells.observation_cell,
        category_cell=cells.category_cell,
        num_participants=2,
        study_name="study",
    )
    clips = spreadsheet.generate_keyword_timestamps(ctx)

    coords = {(clip["cell"].row, clip["cell"].col) for clip in clips}
    # Both cells with segment-level !key annotations should be included.
    assert coords == {(3, 2), (4, 3)}
    segment_indexes = [
        clip.get("selected_segment_indexes") for clip in clips if clip["cell"].col == 3
    ][0]
    assert segment_indexes == [0]


def test_sort_clips_chronologically_orders_by_start_time():
    import gspread

    clips: list[ClipRecord] = [
        {"cell": gspread.cell.Cell(3, 2, "00:30-00:40")},
        {"cell": gspread.cell.Cell(2, 2, "01:00-01:10")},
        {"cell": gspread.cell.Cell(4, 2, "invalid")},
        {"cell": gspread.cell.Cell(1, 2, "00:05-00:10")},
    ]

    spreadsheet.sort_clips_chronologically(clips)

    ordered_values = [clip["cell"].value for clip in clips]
    assert ordered_values[:3] == ["00:05-00:10", "00:30-00:40", "01:00-01:10"]
    assert ordered_values[-1] == "invalid"
