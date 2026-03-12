from pathlib import Path
from typing import cast

import files
import clipgen
import utils
import spreadsheet
from utils import ClipRecord


def test_truncate_filename_respects_max_length(monkeypatch, tmp_path):
    monkeypatch.setattr(files.config, "MAX_FILENAME_LENGTH", 20, raising=False)
    long_base = "a" * 50
    filename = f"{long_base}.mp4"

    truncated_step1 = files.truncate_filename(filename, step=1, file_format=".mp4")
    assert len(truncated_step1) <= 20
    assert truncated_step1.endswith(".mp4")

    truncated_step2 = files.truncate_filename(filename, step=12, file_format=".mp4")
    assert len(truncated_step2) <= 20
    assert truncated_step2.endswith("-12.mp4")


def test_is_source_video_and_discover_clips(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "study_P01.mp4").write_text("video")
    (tmp_path / "study_P01-clip-1.mp4").write_text("clip1")
    (tmp_path / "study_P01-clip-2.mp4").write_text("clip2")

    assert files.is_source_video("study_P01.mp4") is True
    assert files.is_source_video("study_P01-clip-1.mp4") is False

    clips = files.discover_clips()
    assert sorted(clips) == ["study_P01-clip-1.mp4", "study_P01-clip-2.mp4"]


def test_prepare_clip_sanitizes_and_sets_defaults(monkeypatch, make_clip):
    raw_clip = make_clip()
    raw_clip["cell"].value = "00:10-00:20"
    raw_clip["desc"] = "[TAG] Description / with ? chars"
    raw_clip["category"] = ""

    def fake_parse_cell_annotations(value):
        return value, {}, set()

    monkeypatch.setattr(utils, "parse_cell_annotations", fake_parse_cell_annotations)
    monkeypatch.setattr(utils, "has_non_ignored_timestamp_content", lambda _v: True)

    prepared = files.prepare_clip(raw_clip)

    assert prepared["times"] == [("00:10", "00:20")]
    assert "TAG" not in prepared["desc"]
    assert prepared["category"] == "uncategorized"


def test_build_artifact_records_for_clip_and_finalize_timeline_data(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    clip: ClipRecord = {
        "cell": type("Cell", (), {"row": 4, "col": 2})(),
        "study": "study",
        "participant": "P01",
        "category": "CatA",
        "desc": "Obs one",
        "cell_annotations": ["key"],
    }
    base_video = "study_P01.mp4"
    segment_details = [("out_1.mp4", "00:10", "00:20")]

    artifacts = clipgen.build_artifact_records_for_clip(
        clip,
        base_video,
        segment_details,
        output_format="clip",
    )

    assert len(artifacts) == 1
    a = artifacts[0]
    assert a["file"] == "out_1.mp4"
    assert a["start"] == utils.timestamp_to_seconds("00:10")
    assert a["end"] == utils.timestamp_to_seconds("00:20")
    assert a["cellRow"] == 4
    assert a["cellCol"] == 2
    assert "key" in a["annotations"]

    data = clipgen.finalize_timeline_data(
        artifacts,
        study="study",
        participant="P01",
        worksheet_title="Sheet",
        is_excel=False,
        mode="batch",
        output_format="clip",
    )

    assert data["meta"]["study"] == "study"
    assert data["meta"]["participant"] == "P01"
    assert data["meta"]["sourceSpreadsheet"] == "Sheet"
    assert data["meta"]["sourceFileType"] == "google"
    assert data["timeline"]["duration"] > artifacts[0]["end"]


def test_baseline_row_detection_and_relative_conversion():
    # Sheet layout:
    # Row 0: study name
    # Row 1: baseline marker row with 'Baseline time' label and per-participant baseline
    # Row 2: headers: ID, P01, P02, Observation, Category
    # Row 3: data row with clock-style timestamps for both participants
    sheet_data = [
        ["study", "", "", "", ""],
        ["Baseline time", "09:12:00", "", "", ""],
        ["ID", "P01", "P02", "Observation", "Category"],
        ["1", "09:13:00-09:14:00", "09:20:00-09:21:00", "Observation one", "CatA"],
    ]

    # Dummy header cells matching the header row (row index 2 → row=3 in gspread terms)
    id_cell = type("Cell", (), {"row": 3, "col": 1})()
    observation_cell = type("Cell", (), {"row": 3, "col": 4})()
    baseline_row_idx = spreadsheet._detect_baseline_row(sheet_data)
    assert baseline_row_idx == 1

    # Two participants: P01 and P02
    clips = spreadsheet.get_line_timestamps(
        sheet_data,
        id_cell,
        observation_cell,
        num_participants=2,
        line_index=3,
        study_name="study",
        baseline_row_idx=baseline_row_idx,
    )

    # Expect clip records for both participants
    assert len(clips) == 2

    p01_clip = cast(ClipRecord, clips[0])
    p02_clip = cast(ClipRecord, clips[1])

    # P01 column has a baseline in the marker row, P02 does not
    assert p01_clip.get("timestamp_baseline") == "09:12:00"
    assert "timestamp_baseline" not in p02_clip

    prepared_p01 = files.prepare_clip(p01_clip)
    prepared_p02 = files.prepare_clip(p02_clip)

    # P01 times should be converted to relative offsets from 09:12:00
    assert prepared_p01["times"] == [("1:00", "2:00")]
    # P02 times should remain absolute clock values (no baseline applied)
    assert prepared_p02["times"] == [("09:20:00", "09:21:00")]


def test_no_baseline_row_means_relative_timestamps_only():
    # Same layout as above, but without any 'Baseline time' marker row.
    sheet_data = [
        ["study", "", "", "", ""],
        ["", "", "", "", ""],
        ["ID", "P01", "P02", "Observation", "Category"],
        ["1", "09:13:00-09:14:00", "09:20:00-09:21:00", "Observation one", "CatA"],
    ]

    id_cell = type("Cell", (), {"row": 3, "col": 1})()
    observation_cell = type("Cell", (), {"row": 3, "col": 4})()
    baseline_row_idx = spreadsheet._detect_baseline_row(sheet_data)
    assert baseline_row_idx is None

    clips = spreadsheet.get_line_timestamps(
        sheet_data,
        id_cell,
        observation_cell,
        num_participants=2,
        line_index=3,
        study_name="study",
        baseline_row_idx=baseline_row_idx,
    )

    assert len(clips) == 2
    p01_clip = cast(ClipRecord, clips[0])
    p02_clip = cast(ClipRecord, clips[1])

    assert "timestamp_baseline" not in p01_clip
    assert "timestamp_baseline" not in p02_clip

    prepared_p01 = files.prepare_clip(p01_clip)
    prepared_p02 = files.prepare_clip(p02_clip)

    # Without a baseline row, times remain absolute clock values
    assert prepared_p01["times"] == [("09:13:00", "09:14:00")]
    assert prepared_p02["times"] == [("09:20:00", "09:21:00")]

