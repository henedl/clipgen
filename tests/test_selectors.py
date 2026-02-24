import spreadsheet


def test_parse_reel_input_parses_mixed_selectors():
    parsed = spreadsheet.parse_reel_input('timeline, 11, 13-16, P01.5, P02, "Usability"')
    assert parsed["timeline"] is True
    assert parsed["lines"] == [11]
    assert parsed["ranges"] == [(13, 16)]
    assert parsed["cells"] == [("P01", 5)]
    assert parsed["participants"] == ["P02"]
    assert parsed["categories"] == ["Usability"]


def test_detect_mode_from_input_rejects_mixed_types():
    mode, kwargs = spreadsheet.detect_mode_from_input("5, P01.11")
    assert mode is None
    assert kwargs == {}


def test_generate_reel_timestamps_timeline_requires_one_participant(fake_sheet_meta):
    clips = spreadsheet.generate_reel_timestamps(
        sheet_data=[["Study"]],
        id_cell=fake_sheet_meta.id_cell,
        observation_cell=fake_sheet_meta.observation_cell,
        category_cell=fake_sheet_meta.category_cell,
        num_participants=1,
        study_name="study",
        reel_input_string="timeline",
    )
    assert clips == []


def test_generate_reel_timestamps_dedupes_cells(monkeypatch, fake_sheet_meta, make_clip):
    duplicate_clip = make_clip(row=4, col=2)

    monkeypatch.setattr(
        spreadsheet,
        "parse_reel_input",
        lambda _input: {
            "batch": False,
            "filter": True,
            "timeline": False,
            "lines": [4],
            "ranges": [],
            "categories": [],
            "cells": [],
            "participants": [],
        },
    )
    monkeypatch.setattr(spreadsheet, "generate_filter_timestamps", lambda *args, **kwargs: [duplicate_clip])
    monkeypatch.setattr(spreadsheet, "generate_line_timestamps", lambda *args, **kwargs: [duplicate_clip])

    clips = spreadsheet.generate_reel_timestamps(
        sheet_data=[["Study"]],
        id_cell=fake_sheet_meta.id_cell,
        observation_cell=fake_sheet_meta.observation_cell,
        category_cell=fake_sheet_meta.category_cell,
        num_participants=1,
        study_name="study",
        reel_input_string="filter, 4",
    )

    assert len(clips) == 1
    assert (clips[0]["cell"].row, clips[0]["cell"].col) == (4, 2)
