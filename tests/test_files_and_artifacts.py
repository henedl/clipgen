import concurrent.futures
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

import config
import files
import pipeline
import utils
import video
import viewer
import spreadsheet
from spreadsheet import SheetContext
from utils import ClipRecord


def _make_context(
    sheet_data,
    id_cell,
    observation_cell,
    category_cell=None,
    num_participants=2,
    study_name="study",
    baseline_row_idx=None,
    filename_row_idx=None,
):
    """Helper to build a SheetContext for tests."""
    if category_cell is None:
        category_cell = SimpleNamespace(row=id_cell.row, col=5)
    return SheetContext(
        sheet_data=sheet_data,
        id_cell=id_cell,
        observation_cell=observation_cell,
        category_cell=category_cell,
        num_participants=num_participants,
        study_name=study_name,
        baseline_row_idx=baseline_row_idx,
        filename_row_idx=filename_row_idx,
    )


def test_get_unique_filename_truncates_long_names(monkeypatch, tmp_path):
    monkeypatch.setattr(files.config, "MAX_FILENAME_LENGTH", 20, raising=False)
    monkeypatch.setattr(files.config, "OUTPUT_DIR", str(tmp_path), raising=False)
    long_base = "a" * 50

    # First call: no collision, but name is truncated
    result1 = files.get_unique_filename(f"{long_base}.mp4")
    assert len(Path(result1).name) <= 20
    assert result1.endswith(".mp4")

    # Create file to trigger collision
    Path(result1).write_text("exists")

    # Second call: collision, gets "-1" suffix, still within limit
    result2 = files.get_unique_filename(f"{long_base}.mp4")
    assert len(Path(result2).name) <= 20
    assert result2.endswith(".mp4")
    assert "-1" in Path(result2).name
    assert result2 != result1


def test_get_unique_filename_sequential_same_name_batch(monkeypatch, tmp_path):
    """A batch of same-named reservations yields name, name-1, name-2, … with no
    duplicates — and the per-template high-water counter keeps it O(n)."""
    monkeypatch.setattr(files.config, "OUTPUT_DIR", str(tmp_path), raising=False)

    results = [files.get_unique_filename("clip.mp4") for _ in range(5)]
    names = [Path(p).name for p in results]

    assert names == ["clip.mp4", "clip-1.mp4", "clip-2.mp4", "clip-3.mp4", "clip-4.mp4"]
    assert len(set(results)) == len(results)
    for path in results:
        assert Path(path).is_file()  # each reserved as a placeholder


def test_get_unique_filename_reserves_distinct_paths_under_threads(
    monkeypatch, tmp_path
):
    """Concurrent callers requesting the same template each receive a distinct
    path, atomically reserved on disk — no two workers can pick the same name."""
    monkeypatch.setattr(files.config, "OUTPUT_DIR", str(tmp_path), raising=False)

    def reserve():
        return files.get_unique_filename("clip.mp4")

    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
        results = [f.result() for f in [pool.submit(reserve) for _ in range(40)]]

    assert len(set(results)) == len(results)  # every path is unique
    for path in results:
        assert Path(path).is_file()  # each path reserved as a placeholder


def test_release_reservation_removes_unused_placeholder(monkeypatch, tmp_path):
    """An unused reservation is removed; release is a no-op on missing/None."""
    monkeypatch.setattr(files.config, "OUTPUT_DIR", str(tmp_path), raising=False)

    path = files.get_unique_filename("clip.mp4")
    assert Path(path).is_file()

    files.release_reservation(path)
    assert not Path(path).exists()

    # Safe to call again, and on None.
    files.release_reservation(path)
    files.release_reservation(None)


def test_discover_clips_excludes_source_videos(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "study_P01.mp4").write_text("video")
    (tmp_path / "study_P01-clip-1.mp4").write_text("clip1")
    (tmp_path / "study_P01-clip-2.mp4").write_text("clip2")

    clips = files.discover_clips()
    assert sorted(clips) == ["study_P01-clip-1.mp4", "study_P01-clip-2.mp4"]


def test_discover_clips_excludes_numbered_source_videos(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "study_P01.mp4").write_text("video")  # plain source
    (tmp_path / "study_P01-1.mp4").write_text("part1")  # numbered source part
    (tmp_path / "study_P01-2.mp4").write_text("part2")  # numbered source part
    (tmp_path / "[cat] study P01 desc.mp4").write_text("clip")  # generated clip

    clips = files.discover_clips()
    assert clips == ["[cat] study P01 desc.mp4"]


def test_prepare_clip_pre_parsed_fast_path_keeps_times_and_sanitizes_desc():
    # Synthetic clips (e.g. --ss-clips) arrive with times already parsed and a
    # SimpleNamespace cell. prepare_clip should skip the cell-based parse,
    # leave times untouched, and still sanitize desc/category for filenames.
    cell = SimpleNamespace(value="", row=-1, col=1)
    clip: ClipRecord = {
        "cell": cell,
        "study": "study",
        "participant": "P01",
        "category": "",
        "desc": "[ignored-bracket] description / with ?",
        "times": [("0:00:10", "0:00:20")],
    }
    prepared = files.prepare_clip(clip)
    # Times preserved exactly.
    assert prepared["times"] == [("0:00:10", "0:00:20")]
    # Bracketed prefix stripped, special chars removed.
    assert "ignored-bracket" not in prepared["desc"]
    assert "?" not in prepared["desc"]
    # Empty category becomes 'uncategorized'.
    assert prepared["category"] == "uncategorized"
    # Annotation containers initialized.
    assert prepared["cell_annotations"] == []
    assert prepared["segment_annotations"] == {}


def test_build_clip_records_feeds_process_clips(monkeypatch, tmp_path):
    """build_clip_records produces synthetic records that process_clips accepts,
    namespaced under the reserved Workflows cell column (col 3)."""
    monkeypatch.setattr(Path, "is_file", lambda self: True)
    monkeypatch.setattr(utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    monkeypatch.setattr(
        files,
        "get_unique_filename",
        lambda _t, file_format=None: str(tmp_path / f"out{file_format or '.mp4'}"),
    )
    monkeypatch.setattr(video, "run_ffmpeg", Mock(return_value=True))

    records = files.build_clip_records(
        participant="P01",
        source_filename="study_P01.mp4",
        time_ranges=[(10.0, 20.0)],
        description="hit",
        study="study",
    )
    assert len(records) == 1
    assert records[0]["cell"].col == files._WORKFLOW_CELL_COL
    assert records[0]["cell"].row == -1
    # Pre-filled times trigger prepare_clip's fast path inside process_clips.
    assert records[0]["times"] == [("0:00:10", "0:00:20")]

    count, artifacts = pipeline.process_clips(records, output_format="clip")
    assert count == 1
    assert artifacts[0]["id"] == "a-1c3s0"
    assert artifacts[0]["cellCol"] == files._WORKFLOW_CELL_COL


def test_build_clip_records_clusters_adjacent_ranges():
    merged = files.build_clip_records(
        participant="P01",
        source_filename="study_P01.mp4",
        time_ranges=[(10.0, 12.0), (14.0, 16.0)],
        description="x",
        cluster_gap=5.0,
    )
    assert len(merged) == 1  # gap within 5s → single merged record

    split = files.build_clip_records(
        participant="P01",
        source_filename="study_P01.mp4",
        time_ranges=[(10.0, 12.0), (40.0, 42.0)],
        description="x",
        cluster_gap=5.0,
    )
    assert len(split) == 2  # gap exceeded → two records


def test_build_clip_records_no_clustering_keeps_each_range():
    recs = files.build_clip_records(
        participant="P01",
        source_filename="study_P01.mp4",
        time_ranges=[(10.0, 12.0), (14.0, 16.0)],
        description="x",
    )
    assert len(recs) == 2  # no cluster_gap → one record per range
    assert [r["cell"].row for r in recs] == [-1, -2]


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


def test_build_artifact_records_for_clip_and_finalize_timeline_data(
    tmp_path, monkeypatch
):
    monkeypatch.chdir(tmp_path)

    clip: ClipRecord = {
        "cell": type("Cell", (), {"row": 4, "col": 2})(),
        "study": "study",
        "participant": "P01",
        "category": "CatA",
        "desc": "Obs one",
        "cell_annotations": ["key"],
        "times": [("00:10", "00:20")],
    }
    base_video = "study_P01.mp4"
    segment_details = [("out_1.mp4", 0)]

    artifacts = viewer.build_artifact_records_for_clip(
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

    data = viewer.finalize_timeline_data(
        artifacts,
        study="study",
        participant="P01",
        worksheet_title="Sheet",
        is_excel=False,
        mode="batch",
    )

    assert data["meta"]["study"] == "study"
    assert data["meta"]["participant"] == "P01"
    assert data["meta"]["sourceSpreadsheet"] == "Sheet"
    assert data["meta"]["sourceFileType"] == "google"
    assert data["timeline"]["duration"] > artifacts[0]["end"]


def test_build_artifact_records_for_clip_stores_card_image_fields(
    tmp_path, monkeypatch
):
    monkeypatch.chdir(tmp_path)
    clip: ClipRecord = {
        "cell": type("Cell", (), {"row": 4, "col": 2})(),
        "study": "study",
        "participant": "P01",
        "category": "CatA",
        "desc": "Obs one",
        "times": [("00:10", "00:20")],
    }

    artifacts = viewer.build_artifact_records_for_clip(
        clip,
        "study_P01.mp4",
        [("out_1.mp4", 0)],
        output_format="clip",
        titlecards=True,
        titlecard_duration=3,
        titlecard_image="hero.png",
        endcard_image="__color__",
    )
    a = artifacts[0]
    assert a["titlecards"] is True
    assert a["titlecardImage"] == "hero.png"
    assert a["endcardImage"] == "__color__"

    # Non-clip outputs never carry titlecard metadata.
    screens = viewer.build_artifact_records_for_clip(
        clip,
        "study_P01.mp4",
        [("shot.png", 0)],
        output_format="screen",
        titlecards=True,
        titlecard_image="hero.png",
    )
    assert "titlecardImage" not in screens[0]


def test_build_artifact_record_raises_when_cell_missing():
    """Refuse cells without row/col so future callers cannot silently mint
    colliding ids of the form ``a0c0s{seg_idx}``. Two such records would
    dedup against each other in ``viewer.save_manifest``."""
    clip_no_cell: ClipRecord = {
        "cell": None,
        "study": "study",
        "participant": "P01",
        "category": "CatA",
        "desc": "Obs",
        "cell_annotations": [],
        "times": [("00:10", "00:20")],
    }
    with pytest.raises(ValueError, match="cell with row and col"):
        utils.build_artifact_record(
            clip_no_cell,
            "study_P01.mp4",
            "out_1.mp4",
            "00:10",
            "00:20",
            artifact_type="clip",
            seg_idx=0,
        )

    clip_partial_cell: ClipRecord = {
        **clip_no_cell,
        "cell": SimpleNamespace(row=3),
    }
    with pytest.raises(ValueError, match="cell with row and col"):
        utils.build_artifact_record(
            clip_partial_cell,
            "study_P01.mp4",
            "out_1.mp4",
            "00:10",
            "00:20",
            artifact_type="clip",
            seg_idx=0,
        )


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
    id_cell = SimpleNamespace(row=3, col=1)
    observation_cell = SimpleNamespace(row=3, col=4)
    category_cell = SimpleNamespace(row=3, col=5)
    baseline_row_idx = spreadsheet._detect_baseline_row(sheet_data)
    assert baseline_row_idx == 1

    ctx = _make_context(
        sheet_data=sheet_data,
        id_cell=id_cell,
        observation_cell=observation_cell,
        category_cell=category_cell,
        num_participants=2,
        study_name="study",
        baseline_row_idx=baseline_row_idx,
    )

    # Two participants: P01 and P02
    clips = spreadsheet.get_line_timestamps(ctx, 3)

    # Expect clip records for both participants
    assert len(clips) == 2

    p01_clip = clips[0]
    p02_clip = clips[1]

    # P01 column has a baseline in the marker row, P02 does not
    assert p01_clip.get("timestamp_baseline") == "09:12:00"
    assert "timestamp_baseline" not in p02_clip

    prepared_p01 = files.prepare_clip(p01_clip)
    prepared_p02 = files.prepare_clip(p02_clip)

    # P01 times should be converted to relative offsets from 09:12:00
    assert prepared_p01["times"] == [("0:01:00", "0:02:00")]
    # P02 times should remain absolute clock values (no baseline applied)
    assert prepared_p02["times"] == [("09:20:00", "09:21:00")]


def test_baseline_subtracts_from_both_ends_of_multiple_ranges():
    # Regression: when a cell holds multiple baselined HH:MM:SS-HH:MM:SS pairs,
    # baseline must be subtracted from both ends of *each* pair independently.
    sheet_data = [
        ["study", "", "", "", ""],
        ["Baseline time", "09:12:00", "", "", ""],
        ["ID", "P01", "P02", "Observation", "Category"],
        ["1", "09:13:00-09:14:00, 09:15:30-09:16:00", "", "Obs", "CatA"],
    ]
    id_cell = SimpleNamespace(row=3, col=1)
    observation_cell = SimpleNamespace(row=3, col=4)
    category_cell = SimpleNamespace(row=3, col=5)
    ctx = _make_context(
        sheet_data=sheet_data,
        id_cell=id_cell,
        observation_cell=observation_cell,
        category_cell=category_cell,
        num_participants=2,
        study_name="study",
        baseline_row_idx=spreadsheet._detect_baseline_row(sheet_data),
    )
    clips = spreadsheet.get_line_timestamps(ctx, 3)
    prepared = files.prepare_clip(clips[0])
    assert prepared["times"] == [("0:01:00", "0:02:00"), ("0:03:30", "0:04:00")]


def test_no_baseline_row_means_relative_timestamps_only():
    # Same layout as above, but without any 'Baseline time' marker row.
    sheet_data = [
        ["study", "", "", "", ""],
        ["", "", "", "", ""],
        ["ID", "P01", "P02", "Observation", "Category"],
        ["1", "09:13:00-09:14:00", "09:20:00-09:21:00", "Observation one", "CatA"],
    ]

    id_cell = SimpleNamespace(row=3, col=1)
    observation_cell = SimpleNamespace(row=3, col=4)
    category_cell = SimpleNamespace(row=3, col=5)
    baseline_row_idx = spreadsheet._detect_baseline_row(sheet_data)
    assert baseline_row_idx is None

    ctx = _make_context(
        sheet_data=sheet_data,
        id_cell=id_cell,
        observation_cell=observation_cell,
        category_cell=category_cell,
        num_participants=2,
        study_name="study",
        baseline_row_idx=baseline_row_idx,
    )

    clips = spreadsheet.get_line_timestamps(ctx, 3)

    assert len(clips) == 2
    p01_clip = clips[0]
    p02_clip = clips[1]

    assert "timestamp_baseline" not in p01_clip
    assert "timestamp_baseline" not in p02_clip

    prepared_p01 = files.prepare_clip(p01_clip)
    prepared_p02 = files.prepare_clip(p02_clip)

    # Without a baseline row, times remain absolute clock values
    assert prepared_p01["times"] == [("09:13:00", "09:14:00")]
    assert prepared_p02["times"] == [("09:20:00", "09:21:00")]
