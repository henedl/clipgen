"""Tests for multiple source videos per participant (one continuous timeline)."""

import threading
import time
from typing import cast
from unittest.mock import Mock

import pytest

import config
import files
import pipeline
import transcripts
import utils
import video
from utils import ClipRecord


# ---- Pure timeline mapping (utils) ----

TIMELINE = [("video1.mp4", 80, 0), ("video2.mp4", 120, 80)]  # total 200s


def test_map_global_to_segment_basic():
    # 124s global -> 44s into video2 (124 - 80).
    assert utils.map_global_to_segment(TIMELINE, 124) == (1, 44.0)


def test_map_global_to_segment_boundary_goes_to_next_segment():
    # Exactly at the boundary belongs to the start of the following segment.
    assert utils.map_global_to_segment(TIMELINE, 80) == (1, 0.0)
    assert utils.map_global_to_segment(TIMELINE, 79) == (0, 79.0)


def test_map_global_to_segment_out_of_range():
    assert utils.map_global_to_segment(TIMELINE, -1) is None
    assert utils.map_global_to_segment(TIMELINE, 200) is None  # at/after total
    assert utils.map_global_to_segment(TIMELINE, 250) is None


def test_resolve_timeline_segment_returns_path_and_offset():
    # 124s global -> video2 at 44s in; boundary belongs to the next segment.
    assert utils.resolve_timeline_segment(TIMELINE, 124) == ("video2.mp4", 44.0)
    assert utils.resolve_timeline_segment(TIMELINE, 80) == ("video2.mp4", 0.0)
    assert utils.resolve_timeline_segment(TIMELINE, 79) == ("video1.mp4", 79.0)


def test_resolve_timeline_segment_out_of_range():
    assert utils.resolve_timeline_segment(TIMELINE, -1) is None
    assert utils.resolve_timeline_segment(TIMELINE, 200) is None  # at/after total


def test_map_global_range_within_one_segment():
    pieces = utils.map_global_range_to_segments(TIMELINE, 124, 130)
    assert pieces == [(1, 44.0, 50.0)]


def test_map_global_range_spans_boundary():
    # 60s -> 90s straddles the 80s boundary: tail of video1 + head of video2.
    pieces = utils.map_global_range_to_segments(TIMELINE, 60, 90)
    assert pieces == [(0, 60.0, 80.0), (1, 0.0, 10.0)]


def test_map_global_range_clamps_end_to_total():
    pieces = utils.map_global_range_to_segments(TIMELINE, 150, 9999)
    assert pieces == [(1, 70.0, 120.0)]


def test_map_global_range_invalid():
    assert utils.map_global_range_to_segments(TIMELINE, 50, 50) is None
    assert utils.map_global_range_to_segments(TIMELINE, 50, 40) is None
    assert utils.map_global_range_to_segments(TIMELINE, 200, 210) is None


# ---- Duration timeline (video) ----


def test_build_source_timeline_cumulative(monkeypatch):
    durations = {"a.mp4": 80, "b.mp4": 120, "c.mp4": 30}
    monkeypatch.setattr(video, "get_file_duration", lambda p: durations[p])
    timeline = video.build_source_timeline(["a.mp4", "b.mp4", "c.mp4"])
    assert timeline == [("a.mp4", 80, 0), ("b.mp4", 120, 80), ("c.mp4", 30, 200)]


def test_build_source_timeline_none_when_unprobeable(monkeypatch):
    monkeypatch.setattr(
        video, "get_file_duration", lambda p: None if p == "b.mp4" else 10
    )
    assert video.build_source_timeline(["a.mp4", "b.mp4"]) is None


def test_build_source_timeline_order_survives_parallel_probes(monkeypatch):
    """Probes complete in reverse order; the timeline still follows input order."""
    durations = {"a.mp4": 80, "b.mp4": 120, "c.mp4": 30}
    delays = {"a.mp4": 0.06, "b.mp4": 0.03, "c.mp4": 0.0}
    completed: list[str] = []
    lock = threading.Lock()

    def slow_duration(path):
        time.sleep(delays[path])
        with lock:
            completed.append(path)
        return durations[path]

    monkeypatch.setattr(video, "get_file_duration", slow_duration)
    timeline = video.build_source_timeline(["a.mp4", "b.mp4", "c.mp4"])

    assert timeline == [("a.mp4", 80, 0), ("b.mp4", 120, 80), ("c.mp4", 30, 200)]
    # The probes really did finish out of order — ordering above is enforced, not luck.
    assert completed == ["c.mp4", "b.mp4", "a.mp4"]


# ---- Filename resolution (files) ----


def test_get_source_video_filenames_split_on_plus():
    assert files.get_source_video_filenames("s", "P01", "a.mp4 + b.mov") == [
        "a.mp4",
        "b.mov",
    ]
    # Missing extensions get the default appended per part.
    assert files.get_source_video_filenames("s", "P01", "a + b") == ["a.mp4", "b.mp4"]
    # Empty parts dropped; whitespace stripped.
    assert files.get_source_video_filenames("s", "P01", " a.mp4 +  + b.mp4 ") == [
        "a.mp4",
        "b.mp4",
    ]


def test_get_source_video_filenames_no_override():
    assert files.get_source_video_filenames("study", "P01") == ["study_P01.mp4"]


def test_discover_numbered_source_videos_integer_order(tmp_path):
    # A contiguous 1..10 set proves integer (not lexical) sorting: -2 before -10.
    for n in range(1, 11):
        (tmp_path / f"study_P01-{n}.mp4").write_text("v")
    (tmp_path / "study_P02-1.mp4").write_text("other")  # different participant
    found = files.discover_numbered_source_videos(tmp_path, "study", "P01")
    assert [p.name for p in found] == [f"study_P01-{n}.mp4" for n in range(1, 11)]


def test_discover_numbered_source_videos_empty(tmp_path):
    assert files.discover_numbered_source_videos(tmp_path, "study", "P01") == []


def test_discover_numbered_source_videos_non_contiguous_returns_empty(tmp_path):
    # Gap (missing -2) → no valid sequence; guard returns [] rather than a
    # silently-wrong back-to-back concatenation.
    (tmp_path / "study_P01-1.mp4").write_text("v1")
    (tmp_path / "study_P01-3.mp4").write_text("v3")
    assert files.discover_numbered_source_videos(tmp_path, "study", "P01") == []
    # Not starting at 1 is also rejected.
    (tmp_path / "study_P02-2.mp4").write_text("v2")
    (tmp_path / "study_P02-3.mp4").write_text("v3")
    assert files.discover_numbered_source_videos(tmp_path, "study", "P02") == []


def test_resolve_source_video_paths_plain_wins(tmp_path):
    (tmp_path / "study_P01.mp4").write_text("plain")
    (tmp_path / "study_P01-1.mp4").write_text("part1")
    (tmp_path / "study_P01-2.mp4").write_text("part2")
    paths = files.resolve_source_video_paths("study", "P01", None, tmp_path)
    assert [p.name for p in paths] == ["study_P01.mp4"]


def test_resolve_source_video_paths_numbered_when_plain_absent(tmp_path):
    (tmp_path / "study_P01-1.mp4").write_text("part1")
    (tmp_path / "study_P01-2.mp4").write_text("part2")
    paths = files.resolve_source_video_paths("study", "P01", None, tmp_path)
    assert [p.name for p in paths] == ["study_P01-1.mp4", "study_P01-2.mp4"]


def test_resolve_source_video_paths_override(tmp_path):
    paths = files.resolve_source_video_paths("study", "P01", "x.mp4 + y.mp4", tmp_path)
    assert [p.name for p in paths] == ["x.mp4", "y.mp4"]


# ---- Source resolution + timeline build (pipeline) ----


def _ctx_input_dir(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)


def test_check_source_video_single_video_no_probe(monkeypatch, tmp_path, make_clip):
    _ctx_input_dir(monkeypatch, tmp_path)
    (tmp_path / "study_P01.mp4").write_text("v")
    clip = cast(ClipRecord, make_clip())
    build = Mock()
    monkeypatch.setattr(pipeline.video, "build_source_timeline", build)

    result = pipeline._check_source_video(clip, set(), "skip", {})

    assert result == str(tmp_path / "study_P01.mp4")
    assert "source_timeline" not in clip
    build.assert_not_called()


def test_check_source_video_numbered_builds_timeline(monkeypatch, tmp_path, make_clip):
    _ctx_input_dir(monkeypatch, tmp_path)
    (tmp_path / "study_P01-1.mp4").write_text("v1")
    (tmp_path / "study_P01-2.mp4").write_text("v2")
    clip = cast(ClipRecord, make_clip())
    sentinel = [
        (str(tmp_path / "study_P01-1.mp4"), 80, 0),
        (str(tmp_path / "study_P01-2.mp4"), 120, 80),
    ]
    monkeypatch.setattr(pipeline.video, "build_source_timeline", lambda paths: sentinel)

    result = pipeline._check_source_video(clip, set(), "skip", {})

    assert result == str(tmp_path / "study_P01-1.mp4")
    assert clip["source_timeline"] == sentinel


def test_check_source_video_override_plus_list(monkeypatch, tmp_path, make_clip):
    _ctx_input_dir(monkeypatch, tmp_path)
    (tmp_path / "a.mp4").write_text("v1")
    (tmp_path / "b.mp4").write_text("v2")
    clip = cast(ClipRecord, make_clip())
    clip["source_filename"] = "a.mp4 + b.mp4"
    captured = {}

    def fake_build(paths):
        captured["paths"] = paths
        return [(paths[0], 10, 0), (paths[1], 20, 10)]

    monkeypatch.setattr(pipeline.video, "build_source_timeline", fake_build)

    result = pipeline._check_source_video(clip, set(), "skip", {})

    assert result == str(tmp_path / "a.mp4")
    assert captured["paths"] == [str(tmp_path / "a.mp4"), str(tmp_path / "b.mp4")]
    assert "source_timeline" in clip


# ---- Cut mapping + stitch (pipeline) ----


def _multi_clip(make_clip, times) -> ClipRecord:
    clip = cast(ClipRecord, dict(make_clip()))
    clip["times"] = list(times)
    clip["severity"] = ""
    clip["source_timeline"] = [("video1.mp4", 80, 0), ("video2.mp4", 120, 80)]
    return clip


def test_single_video_cut_unchanged_and_no_mapping(monkeypatch, make_clip):
    clip = cast(ClipRecord, dict(make_clip()))
    clip["times"] = [("2:04", "2:10")]
    clip["severity"] = ""
    run_ffmpeg = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", run_ffmpeg)
    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *a, **k: "out.mp4"
    )

    generated, _, _ = pipeline._process_single_clip_segments(
        clip, "study_P01.mp4", set(), output_format="clip", collect_paths=True
    )

    assert generated == 1
    _, kwargs = run_ffmpeg.call_args
    assert kwargs["input_file"] == "study_P01.mp4"
    assert kwargs["start_pos"] == "2:04"
    assert kwargs["end_pos"] == "2:10"


def test_multi_video_clip_maps_into_second_video(monkeypatch, make_clip):
    clip = _multi_clip(make_clip, [("2:04", "2:10")])  # global 124-130
    run_ffmpeg = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", run_ffmpeg)
    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *a, **k: "out.mp4"
    )

    generated, _, _ = pipeline._process_single_clip_segments(
        clip, "video1.mp4", set(), output_format="clip", collect_paths=True
    )

    assert generated == 1
    _, kwargs = run_ffmpeg.call_args
    assert kwargs["input_file"] == "video2.mp4"
    assert kwargs["start_pos"] == "0:00:44"
    assert kwargs["end_pos"] == "0:00:50"


def test_multi_video_titlecard_wraps_at_clip_resolution(monkeypatch, make_clip):
    """A clip cut from a later part must not be wrapped at the first part's resolution.

    The pipeline passes no resolution to wrap_clip_with_cards, which probes the
    generated clip itself — so a video2 cut whose resolution differs from video1
    still gets titlecards instead of a silently-skipped concat mismatch.
    """
    clip = _multi_clip(make_clip, [("2:04", "2:10")])  # global 124-130 -> video2 @44s

    monkeypatch.setattr(pipeline.config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(pipeline.config, "REENCODING", False)
    monkeypatch.setattr(pipeline.config, "DEBUGGING", False)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", lambda **_k: True)
    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *a, **k: "out.mp4"
    )

    # The pipeline must not probe a source video to force a resolution onto wrap.
    def fail_probe(_path):
        raise AssertionError("pipeline should not probe a source video for wrap")

    monkeypatch.setattr(pipeline.video, "probe_video_properties", fail_probe)

    wrap_calls = []

    def fake_wrap(_clip, out_name, resolution=None, **_kwargs):
        wrap_calls.append((out_name, resolution))
        return (True, True)

    monkeypatch.setattr(pipeline.titlecards, "wrap_clip_with_cards", fake_wrap)

    generated, _, _ = pipeline._process_single_clip_segments(
        clip, "video1.mp4", set(), output_format="clip", collect_paths=True
    )

    assert generated == 1
    # Wrap receives no forced resolution -> it self-probes the actual clip.
    assert wrap_calls == [("out.mp4", None)]


def test_multi_video_clip_stitches_across_boundary(monkeypatch, make_clip):
    clip = _multi_clip(make_clip, [("1:00", "1:30")])  # global 60-90 spans boundary
    run_ffmpeg = Mock(return_value=True)
    concat = Mock(return_value=True)
    # First get_unique_filename call is the final out_name; then one per piece.
    names = iter(["out.mp4", "part1.mp4", "part2.mp4"])
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", run_ffmpeg)
    monkeypatch.setattr(pipeline.video, "concatenate_clips", concat)
    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *a, **k: next(names)
    )
    monkeypatch.setattr(pipeline.Path, "unlink", lambda self, **k: None)

    generated, _, _ = pipeline._process_single_clip_segments(
        clip, "video1.mp4", set(), output_format="clip", collect_paths=True
    )

    assert generated == 1
    assert run_ffmpeg.call_count == 2  # one per piece
    concat_args, _ = concat.call_args
    assert concat_args[0] == ["part1.mp4", "part2.mp4"]
    assert concat_args[1] == "out.mp4"


def test_multi_video_screenshot_uses_start_segment(monkeypatch, make_clip):
    clip = _multi_clip(make_clip, [("2:04", "3:04")])  # start 124s -> video2 @44s
    shot = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "extract_screenshot", shot)
    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *a, **k: "out.png"
    )

    pipeline._process_single_clip_segments(
        clip, "video1.mp4", set(), output_format="screen", collect_paths=True
    )

    _, kwargs = shot.call_args
    assert kwargs["input_file"] == "video2.mp4"
    assert kwargs["timestamp"] == "0:00:44"


def test_multi_video_gif_duration_clamped_to_segment_end(monkeypatch, make_clip):
    # Start near the end of video2 so the remaining time is below the default.
    clip = _multi_clip(make_clip, [("3:17", "4:17")])  # 197s -> video2 @117s, 3s left
    gif = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "extract_gif", gif)
    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *a, **k: "out.gif"
    )

    pipeline._process_single_clip_segments(
        clip, "video1.mp4", set(), output_format="gif", collect_paths=True
    )

    _, kwargs = gif.call_args
    assert kwargs["input_file"] == "video2.mp4"
    assert kwargs["duration_seconds"] == min(config.DEFAULT_GIF_DURATION_SECONDS, 3)


# ---- Artifact records carry local + global fields ----


def test_artifact_record_single_video_local_equals_global(make_clip):
    clip = cast(ClipRecord, dict(make_clip()))
    clip["cell_annotations"] = []
    record = utils.build_artifact_record(
        clip,
        "study_P01.mp4",
        "out.mp4",
        "0:10",
        "0:20",
        artifact_type="clip",
        seg_idx=0,
    )
    assert record["sourceVideo"] == "study_P01.mp4"
    assert record["start"] == 10.0 and record["end"] == 20.0
    assert record["localStart"] == 10.0 and record["localEnd"] == 20.0
    assert "parts" not in record


def test_artifact_record_multi_video_maps_local(make_clip):
    clip = _multi_clip(make_clip, [("2:04", "2:10")])
    clip["cell_annotations"] = []
    record = utils.build_artifact_record(
        clip,
        "video1.mp4",
        "out.mp4",
        "2:04",
        "2:10",
        artifact_type="clip",
        seg_idx=0,
    )
    assert record["sourceVideo"] == "video2.mp4"
    assert record["start"] == 124.0 and record["end"] == 130.0  # global preserved
    assert record["localStart"] == 44.0 and record["localEnd"] == 50.0
    assert "parts" not in record


def test_artifact_record_boundary_clip_has_parts(make_clip):
    clip = _multi_clip(make_clip, [("1:00", "1:30")])
    clip["cell_annotations"] = []
    record = utils.build_artifact_record(
        clip,
        "video1.mp4",
        "out.mp4",
        "1:00",
        "1:30",
        artifact_type="clip",
        seg_idx=0,
    )
    assert record["parts"] == [
        {"sourceVideo": "video1.mp4", "localStart": 60.0, "localEnd": 80.0},
        {"sourceVideo": "video2.mp4", "localStart": 0.0, "localEnd": 10.0},
    ]
    # Top-level fields carry the first piece.
    assert record["sourceVideo"] == "video1.mp4"


def test_artifact_record_screenshot_never_splits(make_clip):
    # A screenshot's [start,end] window would straddle the boundary, but a single
    # frame must map by start only — never split.
    clip = _multi_clip(make_clip, [("1:00", "1:30")])
    clip["cell_annotations"] = []
    record = utils.build_artifact_record(
        clip,
        "video1.mp4",
        "out.png",
        "1:00",
        "1:30",
        artifact_type="screen",
        seg_idx=0,
    )
    assert "parts" not in record
    assert record["sourceVideo"] == "video1.mp4"
    assert record["localStart"] == 60.0


def test_artifact_record_full_path_base_video_normalized_to_basename(make_clip):
    # Producers may hold a full path in base_video; the persisted sourceVideo must
    # be a basename (matching pipeline.cut_global_range) so sheet and intake
    # manifests share one shape. Regression for the path-shape alignment fix.
    clip = cast(ClipRecord, dict(make_clip()))
    clip["cell_annotations"] = []
    record = utils.build_artifact_record(
        clip,
        "/srv/input/study_P01.mp4",
        "out.mp4",
        "0:10",
        "0:20",
        artifact_type="clip",
        seg_idx=0,
    )
    assert record["sourceVideo"] == "study_P01.mp4"


def test_artifact_record_full_path_timeline_normalized_to_basenames(make_clip):
    # Multi-video: full-path source_timeline entries must persist as basenames at
    # both the top level and inside the boundary-spanning parts list.
    clip = cast(ClipRecord, dict(make_clip()))
    clip["cell_annotations"] = []
    clip["times"] = [("1:00", "1:30")]
    clip["severity"] = ""
    clip["source_timeline"] = [
        ("/srv/input/video1.mp4", 80, 0),
        ("/srv/input/video2.mp4", 120, 80),
    ]
    record = utils.build_artifact_record(
        clip,
        "/srv/input/video1.mp4",
        "out.mp4",
        "1:00",
        "1:30",
        artifact_type="clip",
        seg_idx=0,
    )
    assert record["sourceVideo"] == "video1.mp4"
    assert [p["sourceVideo"] for p in record["parts"]] == ["video1.mp4", "video2.mp4"]


# ---- Regenerate from manifest ----


def test_regenerate_boundary_artifact_stitches_parts(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path), raising=False)
    (tmp_path / "video1.mp4").write_text("v1")
    (tmp_path / "video2.mp4").write_text("v2")
    run_ffmpeg = Mock(return_value=True)
    concat = Mock(return_value=True)
    names = iter(["t1.mp4", "t2.mp4"])
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", run_ffmpeg)
    monkeypatch.setattr(pipeline.video, "concatenate_clips", concat)
    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *a, **k: next(names)
    )
    monkeypatch.setattr(pipeline.Path, "unlink", lambda self, **k: None)

    artifact = {
        "type": "clip",
        "file": "clip.mp4",
        "sourceVideo": "video1.mp4",
        "localStart": 60.0,
        "localEnd": 80.0,
        "parts": [
            {"sourceVideo": "video1.mp4", "localStart": 60.0, "localEnd": 80.0},
            {"sourceVideo": "video2.mp4", "localStart": 0.0, "localEnd": 10.0},
        ],
    }
    assert pipeline._regenerate_single_artifact(artifact, set()) is True
    assert run_ffmpeg.call_count == 2
    concat_args, _ = concat.call_args
    assert concat_args[0] == ["t1.mp4", "t2.mp4"]


def test_regenerate_single_artifact_uses_local_times(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path), raising=False)
    (tmp_path / "video2.mp4").write_text("v2")
    run_ffmpeg = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", run_ffmpeg)

    artifact = {
        "type": "clip",
        "file": "clip.mp4",
        "sourceVideo": "video2.mp4",
        "start": 124.0,
        "end": 130.0,
        "localStart": 44.0,
        "localEnd": 50.0,
    }
    assert pipeline._regenerate_single_artifact(artifact, set()) is True
    _, kwargs = run_ffmpeg.call_args
    assert kwargs["input_file"] == str(tmp_path / "video2.mp4")
    assert kwargs["start_pos"] == "0:44"  # local, not global
    assert kwargs["end_pos"] == "0:50"


# ---- Transcribe all parts (merged global timeline) ----


def test_transcribe_timeline_shifts_segment_times(monkeypatch):
    results = {
        "a.mp4": {
            "segments": [{"start": 0.0, "end": 5.0, "text": "one"}],
            "language": "en",
            "source_file": "a.mp4",
            "model": "base",
        },
        "b.mp4": {
            "segments": [{"start": 1.0, "end": 4.0, "text": "two"}],
            "language": "en",
            "source_file": "b.mp4",
            "model": "base",
        },
    }
    monkeypatch.setattr(transcripts, "transcribe_video", lambda p, **kwargs: results[p])
    timeline = [("a.mp4", 80, 0), ("b.mp4", 120, 80)]
    merged = transcripts.transcribe_timeline(timeline)
    assert merged is not None
    assert merged["segments"] == [
        {"start": 0.0, "end": 5.0, "text": "one"},
        {"start": 81.0, "end": 84.0, "text": "two"},  # shifted by 80
    ]
    assert merged["source_file"] == "a.mp4 + b.mp4"


def test_transcribe_timeline_shifts_word_times(monkeypatch):
    from typing import Any

    results: dict[str, Any] = {
        "a.mp4": {
            "segments": [
                {
                    "start": 0.0,
                    "end": 5.0,
                    "text": "one",
                    "words": [{"start": 0.5, "end": 4.5, "text": "one"}],
                }
            ],
            "language": "en",
            "source_file": "a.mp4",
            "model": "base",
        },
        "b.mp4": {
            "segments": [
                {
                    "start": 1.0,
                    "end": 4.0,
                    "text": "two",
                    "words": [{"start": 1.25, "end": 3.75, "text": "two"}],
                }
            ],
            "language": "en",
            "source_file": "b.mp4",
            "model": "base",
        },
    }
    streamed: list = []

    def _fake_transcribe(path, on_segment=None, **kwargs):
        result = results[path]
        if on_segment is not None:
            for seg in result["segments"]:
                on_segment(seg["end"], seg)
        return result

    monkeypatch.setattr(transcripts, "transcribe_video", _fake_transcribe)
    timeline = [("a.mp4", 80, 0), ("b.mp4", 120, 80)]
    merged = transcripts.transcribe_timeline(
        timeline, on_segment=lambda end, seg: streamed.append(seg)
    )
    assert merged is not None
    assert merged["segments"][0]["words"] == [{"start": 0.5, "end": 4.5, "text": "one"}]
    assert merged["segments"][1]["words"] == [
        {"start": 81.25, "end": 83.75, "text": "two"}
    ]
    # The streaming shifter rebases word times identically.
    assert [s["words"] for s in streamed] == [s["words"] for s in merged["segments"]]


def test_transcribe_timeline_none_on_failure(monkeypatch):
    monkeypatch.setattr(transcripts, "transcribe_video", lambda p, **kwargs: None)
    assert transcripts.transcribe_timeline([("a.mp4", 10, 0)]) is None


def test_transcribe_timeline_window_inside_one_part_skips_the_others(monkeypatch):
    """A global in/out window wholly inside part 2 never touches part 1, and the
    part-local results land back on the global timeline."""
    calls: list = []

    def fake_video(path, **kwargs):
        calls.append((path, kwargs.get("start_seconds"), kwargs.get("end_seconds")))
        # transcribe_video returns times on the *file's* timeline.
        return {
            "segments": [{"start": 11.0, "end": 12.0, "text": "two"}],
            "language": "en",
            "source_file": path,
            "model": "base",
        }

    monkeypatch.setattr(transcripts, "transcribe_video", fake_video)
    timeline = [("a.mp4", 80, 0), ("b.mp4", 120, 80)]
    merged = transcripts.transcribe_timeline(
        timeline, start_seconds=90, end_seconds=100
    )
    assert merged is not None
    assert calls == [("b.mp4", 10.0, 20.0)]  # part-local window; a.mp4 skipped
    assert merged["segments"] == [{"start": 91.0, "end": 92.0, "text": "two"}]


def test_transcribe_timeline_window_spanning_boundary(monkeypatch):
    """A window straddling the part boundary gives each part its own sub-window,
    unbounded on the side that runs to the part's edge."""
    calls: list = []

    def fake_video(path, **kwargs):
        calls.append((path, kwargs.get("start_seconds"), kwargs.get("end_seconds")))
        return {"segments": [], "language": "en", "source_file": path, "model": "base"}

    monkeypatch.setattr(transcripts, "transcribe_video", fake_video)
    timeline = [("a.mp4", 80, 0), ("b.mp4", 120, 80)]
    merged = transcripts.transcribe_timeline(
        timeline, start_seconds=60, end_seconds=100
    )
    assert merged is not None
    assert calls == [("a.mp4", 60.0, None), ("b.mp4", None, 20.0)]


def test_transcribe_timeline_attempted_part_failure_still_aborts(monkeypatch):
    """Skipped-by-window is not failed — but a part that was attempted and
    returned None aborts the job exactly as before."""

    def fake_video(path, **kwargs):
        return (
            None
            if path == "b.mp4"
            else {
                "segments": [],
                "language": "en",
                "source_file": path,
                "model": "base",
            }
        )

    monkeypatch.setattr(transcripts, "transcribe_video", fake_video)
    timeline = [("a.mp4", 80, 0), ("b.mp4", 120, 80)]
    assert transcripts.transcribe_timeline(timeline, start_seconds=60) is None


def test_transcribe_segments_multi_video_uses_global_timeline(
    make_clip, monkeypatch, tmp_path
):
    """The CLI pipeline path transcribes a multi-video clip via the global
    timeline (not per-file) and writes a transcript clipped to the segment window
    on the global timeline, with the artifact stamped with global clip times."""
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path), raising=False)
    clip = cast(ClipRecord, dict(make_clip()))
    clip["cell_annotations"] = []
    clip["participant"] = "P01"
    clip["times"] = [("1:00", "1:30")]  # global 60-90s, straddles the 80s boundary
    clip["severity"] = ""
    clip["source_timeline"] = [("video1.mp4", 80, 0), ("video2.mp4", 120, 80)]

    # Full transcript already on the participant's GLOBAL timeline.
    full = {
        "segments": [
            {"start": 70.0, "end": 75.0, "text": "before boundary"},
            {"start": 85.0, "end": 88.0, "text": "after boundary"},
        ],
        "language": "en",
        "source_file": "video1.mp4 + video2.mp4",
        "model": "base",
    }
    monkeypatch.setattr(transcripts, "transcribe_timeline", lambda timeline, **kw: full)
    monkeypatch.setattr(
        transcripts, "transcribe_video", lambda *a, **k: pytest.fail("used single path")
    )
    monkeypatch.setattr(
        transcripts,
        "load_transcripts_manifest",
        lambda: {"source_transcripts": {}, "corrections": []},
    )

    written = {}

    def fake_write(clipped, path):
        written["segments"] = clipped["segments"]
        return True

    monkeypatch.setattr(transcripts, "write_transcript", fake_write)
    monkeypatch.setattr(
        files, "get_unique_filename", lambda name, **k: str(tmp_path / name)
    )

    artifacts: list[dict] = []
    pipeline._transcribe_segments(
        clip, "video1.mp4", [("clip_out.mp4", 0)], artifacts, {}, None
    )

    assert len(artifacts) == 1
    art = artifacts[0]
    assert art["type"] == "transcript"
    assert art["start"] == 60.0 and art["end"] == 90.0  # global clip window
    # Both global segments fall in [60, 90]; offset_to_zero rebases to clip start.
    assert written["segments"] == [
        {"start": 10.0, "end": 15.0, "text": "before boundary"},
        {"start": 25.0, "end": 28.0, "text": "after boundary"},
    ]


# ---- Studio hover thumbnail maps into the right sub-video ----


def test_studio_thumbnail_maps_to_sub_video(monkeypatch, tmp_path):
    Flask = pytest.importorskip("flask").Flask
    import spreadsheet
    import server

    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    (tmp_path / "a.mp4").write_text("v1")
    (tmp_path / "b.mp4").write_text("v2")

    sheet_data = [
        ["ID", "P01"],  # header row (row 1)
        ["Filename", "a.mp4 + b.mp4"],  # filename row (row 2)
    ]
    ctx = spreadsheet.SheetContext(
        sheet_data=sheet_data,
        id_cell=type("C", (), {"row": 1, "col": 1})(),
        observation_cell=type("C", (), {"row": 1, "col": 3})(),
        category_cell=type("C", (), {"row": 1, "col": 4})(),
        num_participants=1,
        study_name="study",
        filename_row_idx=1,
    )
    monkeypatch.setattr(server, "_sheet_context", ctx)
    monkeypatch.setattr(
        server.video,
        "build_source_timeline",
        lambda paths: [(paths[0], 80, 0), (paths[1], 120, 80)],
    )
    captured = {}

    def fake_thumb(path, sec, width):
        captured["path"] = path
        captured["sec"] = sec
        return b"jpeg"

    monkeypatch.setattr(server.video, "extract_thumbnail_bytes", fake_thumb)

    app = Flask(__name__)
    app.register_blueprint(server.studio_bp, url_prefix="/studio")
    with app.test_client() as c:
        resp = c.get("/studio/api/thumbnail/P01/124")  # global 124s -> video2 @44s

    assert resp.status_code == 200
    assert captured["path"] == str(tmp_path / "b.mp4")
    assert captured["sec"] == 44


# ---- discover_participant_videos grouping (shared bug-fixes) ----


def test_discover_participant_videos_groups_numbered_parts(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    (tmp_path / "study_P01-1.mp4").write_text("v1")
    (tmp_path / "study_P01-2.mp4").write_text("v2")
    found = utils.discover_participant_videos("study")
    # Regression: numbered parts must group under base id P01, never "P01-1".
    ids = [p["id"] for p in found]
    assert ids == ["P01"]
    assert [_basename(p) for p in found[0]["video_paths"]] == [
        "study_P01-1.mp4",
        "study_P01-2.mp4",
    ]
    assert found[0]["has_video"] is True


def test_discover_participant_videos_plain_wins(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    (tmp_path / "study_P01.mp4").write_text("v")
    (tmp_path / "study_P01-1.mp4").write_text("v1")
    found = utils.discover_participant_videos("study")
    assert len(found) == 1
    assert [_basename(p) for p in found[0]["video_paths"]] == ["study_P01.mp4"]


def test_discover_participant_videos_skips_non_contiguous(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    (tmp_path / "study_P01-1.mp4").write_text("v1")
    (tmp_path / "study_P01-3.mp4").write_text("v3")  # gap → skipped
    (tmp_path / "study_P02.mp4").write_text("v")  # normal participant kept
    ids = [p["id"] for p in utils.discover_participant_videos("study")]
    assert ids == ["P02"]


def _basename(path_str):
    from pathlib import Path

    return Path(path_str).name


def test_participant_id_from_source_name():
    assert utils.participant_id_from_source_name("study_P01.mp4") == "P01"
    assert utils.participant_id_from_source_name("study_P01-2.mp4") == "P01"
    assert utils.participant_id_from_source_name("my-study_G02-10.mp4") == "G02"
    assert utils.participant_id_from_source_name("random.mp4") is None
    # A Finder/Explorer duplicate ("… copy.mp4") yields a whitespace id, which is
    # never a real participant — reject it so it can't become a phantom
    # participant or auto-launch a watch-dir-triggered run for a bogus id.
    assert utils.participant_id_from_source_name("study_P03 copy.mp4") is None
    assert utils.participant_id_from_source_name("study_P03 copy 2.mp4") is None


def test_parse_source_video_name():
    assert utils.parse_source_video_name("study_P01.mp4") == ("study", "P01", None)
    assert utils.parse_source_video_name("study_P01-2.mp4") == ("study", "P01", 2)
    # Study may contain the separator (greedy rsplit-like semantics) and dashes.
    assert utils.parse_source_video_name("my_study_P01.mp4") == (
        "my_study",
        "P01",
        None,
    )
    assert utils.parse_source_video_name("my-study_G02-10.mp4") == (
        "my-study",
        "G02",
        10,
    )
    # Empty study accepted; a stem with no separator is not a source video.
    assert utils.parse_source_video_name("_P01.mp4") == ("", "P01", None)
    assert utils.parse_source_video_name("random.mp4") is None
    assert utils.parse_source_video_name("study_P03 copy.mp4") is None
    # A lowercase prefix groups under the configured casing.
    assert utils.parse_source_video_name("study_p01.mp4") == ("study", "P01", None)


def test_parse_source_video_name_custom_patterns(monkeypatch):
    monkeypatch.setattr(config, "SOURCE_FILENAME_PATTERN", "{participant}_{study}")
    assert utils.parse_source_video_name("P01_study.mp4") == ("study", "P01", None)
    assert utils.parse_source_video_name("P01_study-2.mp4") == ("study", "P01", 2)
    assert utils.parse_source_video_name("study_P01.mp4") is None

    monkeypatch.setattr(config, "SOURCE_FILENAME_PATTERN", "{participant}")
    assert utils.parse_source_video_name("P01.mp4") == ("", "P01", None)
    assert utils.parse_source_video_name("G02-3.mp4") == ("", "G02", 3)
    assert utils.parse_source_video_name("P01 copy.mp4") is None
    assert utils.parse_source_video_name("random.mp4") is None

    monkeypatch.setattr(config, "SOURCE_FILENAME_PATTERN", "{study} {participant}")
    assert utils.parse_source_video_name("my study P01.mp4") == (
        "my study",
        "P01",
        None,
    )
    assert utils.parse_source_video_name("my_study_P01.mp4") is None


def test_format_source_video_stem(monkeypatch):
    assert utils.format_source_video_stem("study", "P01") == "study_P01"
    monkeypatch.setattr(config, "SOURCE_FILENAME_PATTERN", "{participant}")
    # str.format ignores the unused study kwarg.
    assert utils.format_source_video_stem("study", "P01") == "P01"


def test_get_source_video_filenames_custom_pattern(monkeypatch):
    monkeypatch.setattr(config, "SOURCE_FILENAME_PATTERN", "{participant}_{study}")
    assert files.get_source_video_filenames("study", "P01") == ["P01_study.mp4"]
    # Per-participant overrides still beat the pattern.
    assert files.get_source_video_filenames("study", "P01", "x.mp4") == ["x.mp4"]


def test_discover_numbered_source_videos_custom_pattern(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "SOURCE_FILENAME_PATTERN", "{participant}_{study}")
    (tmp_path / "P01_study-1.mp4").write_text("v1")
    (tmp_path / "P01_study-2.mp4").write_text("v2")
    found = files.discover_numbered_source_videos(tmp_path, "study", "P01")
    assert [p.name for p in found] == ["P01_study-1.mp4", "P01_study-2.mp4"]


def test_discover_numbered_source_videos_glob_metachars_in_pattern(
    monkeypatch, tmp_path
):
    # "[" / "]" are legal filename characters and pass pattern validation, but
    # they are Path.glob metacharacters — an unescaped glob reads them as a
    # character class and silently finds no parts, so regex-based participant
    # discovery would list files that clip resolution then reports missing.
    monkeypatch.setattr(config, "SOURCE_FILENAME_PATTERN", "[{study}] {participant}")
    (tmp_path / "[study] P01-1.mp4").write_text("v1")
    (tmp_path / "[study] P01-2.mp4").write_text("v2")
    found = files.discover_numbered_source_videos(tmp_path, "study", "P01")
    assert [p.name for p in found] == ["[study] P01-1.mp4", "[study] P01-2.mp4"]
    # Numbered-part resolution and regex discovery agree on the same files.
    paths = files.resolve_source_video_paths("study", "P01", None, tmp_path)
    assert [p.name for p in paths] == ["[study] P01-1.mp4", "[study] P01-2.mp4"]
    assert utils.parse_source_video_name("[study] P01-1.mp4") == ("study", "P01", 1)


def test_discover_participant_videos_custom_pattern(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    monkeypatch.setattr(config, "SOURCE_FILENAME_PATTERN", "{participant}")
    (tmp_path / "P01.mp4").write_text("v")
    (tmp_path / "G02-1.mp4").write_text("v1")
    (tmp_path / "G02-2.mp4").write_text("v2")
    (tmp_path / "random.mp4").write_text("x")
    found = utils.discover_participant_videos()
    assert [p["id"] for p in found] == ["G02", "P01"]
    assert [_basename(p) for p in found[0]["video_paths"]] == [
        "G02-1.mp4",
        "G02-2.mp4",
    ]


def test_discover_participant_videos_rescans_on_pattern_change(monkeypatch, tmp_path):
    # A settings PUT changes the pattern without touching the directory mtime;
    # the memo must not serve the old pattern's participant list.
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    (tmp_path / "P01_study.mp4").write_text("v")
    assert utils.discover_participant_videos() == []
    monkeypatch.setattr(config, "SOURCE_FILENAME_PATTERN", "{participant}_{study}")
    assert [p["id"] for p in utils.discover_participant_videos()] == ["P01"]


def test_validate_source_filename_pattern():
    assert utils.validate_source_filename_pattern("{study}_{participant}") is None
    assert utils.validate_source_filename_pattern("{participant}") is None
    assert utils.validate_source_filename_pattern("{study} {participant}") is None
    assert utils.validate_source_filename_pattern("session-{participant}") is None
    # Each rejection rule.
    assert utils.validate_source_filename_pattern("") is not None
    assert utils.validate_source_filename_pattern("   ") is not None
    assert utils.validate_source_filename_pattern("{study}_{p") is not None
    assert utils.validate_source_filename_pattern("{foo}_{participant}") is not None
    assert utils.validate_source_filename_pattern("{}_{participant}") is not None
    assert utils.validate_source_filename_pattern("{study}") is not None
    assert (
        utils.validate_source_filename_pattern("{participant}_{participant}")
        is not None
    )
    assert utils.validate_source_filename_pattern("a/{participant}") is not None
    assert utils.validate_source_filename_pattern("a:{participant}") is not None


def test_numbered_parts_are_contiguous():
    assert utils.numbered_parts_are_contiguous([1, 2, 3]) is True
    assert utils.numbered_parts_are_contiguous([2, 1, 3]) is True  # order-insensitive
    assert utils.numbered_parts_are_contiguous([1, 3]) is False
    assert utils.numbered_parts_are_contiguous([2, 3]) is False
    assert utils.numbered_parts_are_contiguous([]) is True


# ---- Screenspace: global event offsets + per-event source ----


def test_offset_result_times_shifts_point_and_span():
    import screenspace

    point = {"timestamp": 10.0}
    screenspace._offset_result_times(point, 80)
    assert point["timestamp"] == 90.0

    span = {"timestamp": 5.0, "start": 5.0, "end": 12.0}
    screenspace._offset_result_times(span, 80)
    assert span == {"timestamp": 85.0, "start": 85.0, "end": 92.0}

    untouched = {"timestamp": 3.0}
    screenspace._offset_result_times(untouched, 0)  # no-op
    assert untouched["timestamp"] == 3.0


def test_generate_events_uses_per_result_source_video():
    import screenspace

    task = {
        "id": "ss_1",
        "type": "color",
        "participant": "P01",
        "source_video": "study_P01-1.mp4",
        "region": "r",
        "parameters": {},
    }
    raw = [
        {"timestamp": 5.0, "_confidence": 0.9, "_source_video": "study_P01-1.mp4"},
        {"timestamp": 130.0, "_confidence": 0.9, "_source_video": "study_P01-2.mp4"},
    ]
    events = screenspace.generate_events_from_results(task, raw)
    assert events[0]["source_video"] == "study_P01-1.mp4"
    assert events[0]["time_in"] == 5.0
    assert events[1]["source_video"] == "study_P01-2.mp4"  # per-result override
    assert events[1]["time_in"] == 130.0  # global (offset already applied upstream)


# ---- Screenspace: global->sub-video frame mapping ----


def test_screenspace_map_participant_time(monkeypatch):
    import screenspace_server

    monkeypatch.setattr(
        screenspace_server,
        "_participants",
        [
            {"id": "P01", "video_paths": ["/in/study_P01.mp4"], "has_video": True},
            {
                "id": "P02",
                "video_paths": ["/in/study_P02-1.mp4", "/in/study_P02-2.mp4"],
                "has_video": True,
            },
        ],
    )
    # Single video → identity map (no probe, no stat).
    assert screenspace_server._map_participant_time("P01", 12.0) == (
        "/in/study_P01.mp4",
        12.0,
    )
    # Multi video → owning sub-video + local offset.
    monkeypatch.setattr(
        screenspace_server,
        "_participant_timeline",
        lambda pid: [("/in/study_P02-1.mp4", 80, 0), ("/in/study_P02-2.mp4", 120, 80)],
    )
    assert screenspace_server._map_participant_time("P02", 124.0) == (
        "/in/study_P02-2.mp4",
        44.0,
    )
    # Out of range → None.
    assert screenspace_server._map_participant_time("P02", 9999.0) is None
    # Unknown participant → None.
    assert screenspace_server._map_participant_time("PX", 1.0) is None


# ---- Transcripts worker: single vs multi-video execution ----


@pytest.fixture(autouse=True)
def _stub_whisper_load(monkeypatch):
    """Keep _execute_task's model preload out of these tests.

    The preload (added with the loading_model phase) runs before the
    transcribe_video/transcribe_timeline stubs below get a look in, and it is
    not stubbed by them — so without this these tests constructed a *real*
    WhisperModel, downloading ~140 MB on any machine with a cold HF cache and
    leaving the shared transcripts._cached_model populated for whatever ran
    next. That shared cache is what made them fail intermittently depending on
    file order.
    """
    monkeypatch.setattr(transcripts, "_load_model", lambda name=None: object())


def test_transcript_worker_multi_video_builds_timeline(monkeypatch):
    worker = transcripts.TranscriptWorker()
    task = transcripts.create_transcript_task("P01", ["a.mp4", "b.mp4"])
    worker._tasks[task["id"]] = task

    # _execute_task does `import video as video_mod`, so patch the video module.
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda p: {"audio_codec": "aac", "duration": 80.0},
    )
    monkeypatch.setattr(
        video,
        "timeline_or_none",
        lambda paths: [("a.mp4", 80, 0), ("b.mp4", 120, 80)],
    )
    captured = {}

    def fake_timeline(timeline, **kwargs):
        captured["timeline"] = timeline
        return {
            "segments": [{"start": 81.0, "end": 84.0, "text": "two"}],
            "language": "en",
            "model": "base",
            "source_file": "a.mp4 + b.mp4",
        }

    monkeypatch.setattr(transcripts, "transcribe_timeline", fake_timeline)
    monkeypatch.setattr(
        transcripts, "transcribe_video", lambda *a, **k: pytest.fail("used single path")
    )

    worker._execute_task(task)
    assert task["status"] == transcripts.TASK_STATUS_COMPLETED
    assert task["result"]["source_file"] == "a.mp4 + b.mp4"
    assert captured["timeline"][1][0] == "b.mp4"


def test_transcript_worker_records_the_audio_track_it_used(monkeypatch):
    """The task's audio_index reaches transcribe_timeline and lands on the result.

    The manifest entry is what explains a transcript that changed because
    auto-detect moved off track 1, so it has to survive the worker.
    """
    worker = transcripts.TranscriptWorker()
    task = transcripts.create_transcript_task("P01", ["a.mp4", "b.mp4"], audio_index=1)
    worker._tasks[task["id"]] = task

    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda p: {
            "audio_codec": "aac",
            "duration": 80.0,
            "audio_tracks": [
                {"index": 0, "label": "System"},
                {"index": 1, "label": "Interview"},
            ],
            "audio_track_count": 2,
        },
    )
    monkeypatch.setattr(
        video, "timeline_or_none", lambda paths: [("a.mp4", 80, 0), ("b.mp4", 120, 80)]
    )
    captured = {}

    def fake_timeline(timeline, **kwargs):
        captured.update(kwargs)
        return {
            "segments": [],
            "language": "en",
            "model": "base",
            "source_file": "a.mp4 + b.mp4",
        }

    monkeypatch.setattr(transcripts, "transcribe_timeline", fake_timeline)

    worker._execute_task(task)
    assert task["status"] == transcripts.TASK_STATUS_COMPLETED
    assert captured["audio_index"] == 1
    assert task["result"]["audio_index"] == 1
    assert task["result"]["audio_track_label"] == "Interview"


def test_transcript_worker_rejects_out_of_range_audio_track(monkeypatch):
    worker = transcripts.TranscriptWorker()
    task = transcripts.create_transcript_task("P01", ["a.mp4"], audio_index=4)
    worker._tasks[task["id"]] = task

    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda p: {
            "audio_codec": "aac",
            "duration": 10.0,
            "audio_tracks": [{"index": 0, "label": "Mic"}],
            "audio_track_count": 1,
        },
    )
    monkeypatch.setattr(video, "timeline_or_none", lambda paths: None)
    monkeypatch.setattr(
        transcripts, "transcribe_video", lambda *a, **k: pytest.fail("must not run")
    )

    worker._execute_task(task)
    assert task["status"] == transcripts.TASK_STATUS_FAILED
    assert "Audio track 5 does not exist" in task["error"]


def test_transcribe_timeline_resolves_the_audio_track_once(monkeypatch):
    """Auto-detect runs on part 1 only; every part gets the same concrete index.

    Per-part detection could splice two different microphones into one transcript
    if the recorder's stream order shifted between files.
    """
    probes = {
        "a.mp4": [{"index": 0, "label": "System"}, {"index": 1, "label": "Mic"}],
        "b.mp4": [{"index": 0, "label": "Mic"}, {"index": 1, "label": "System"}],
    }
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda p: {"audio_tracks": probes[p], "audio_track_count": 2},
    )
    seen: list[int | None] = []

    def fake_video(path, **kwargs):
        seen.append(kwargs.get("audio_index"))
        return {"segments": [], "language": "en", "model": "base", "source_file": path}

    monkeypatch.setattr(transcripts, "transcribe_video", fake_video)

    transcripts.transcribe_timeline([("a.mp4", 80, 0), ("b.mp4", 120, 80)])
    assert seen == [1, 1]


def test_transcript_worker_single_video_no_timeline(monkeypatch):
    worker = transcripts.TranscriptWorker()
    task = transcripts.create_transcript_task("P01", ["solo.mp4"])
    worker._tasks[task["id"]] = task

    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda p: {"audio_codec": "aac", "duration": 50.0},
    )
    build = Mock()
    monkeypatch.setattr(video, "build_source_timeline", build)
    monkeypatch.setattr(
        transcripts,
        "transcribe_video",
        lambda *a, **k: {
            "segments": [],
            "language": "en",
            "model": "base",
            "source_file": "solo.mp4",
        },
    )
    monkeypatch.setattr(
        transcripts,
        "transcribe_timeline",
        lambda *a, **k: pytest.fail("used timeline for single video"),
    )

    worker._execute_task(task)
    assert task["status"] == transcripts.TASK_STATUS_COMPLETED
    build.assert_not_called()  # single-video fast path: no duration probe via timeline


def test_transcript_worker_window_reaches_transcribe_and_result(monkeypatch):
    """The task's in/out window is clamped, forwarded to the transcribe call,
    recorded on the result, and used as the progress denominator."""
    worker = transcripts.TranscriptWorker()
    task = transcripts.create_transcript_task(
        "P01", ["solo.mp4"], start_seconds=10.0, end_seconds=40.0
    )
    worker._tasks[task["id"]] = task

    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda p: {"audio_codec": "aac", "duration": 50.0},
    )
    monkeypatch.setattr(video, "timeline_or_none", lambda paths: None)
    captured = {}

    def fake_video(path, on_segment=None, **kwargs):
        captured.update(kwargs)
        # Emitted times are global (already shifted by the window start).
        seg = {"start": 20.0, "end": 25.0, "text": "mid"}
        if on_segment is not None:
            on_segment(seg["end"], seg)
        captured["progress_at_seg"] = task["progress"]
        return {
            "segments": [seg],
            "language": "en",
            "model": "base",
            "source_file": "solo.mp4",
        }

    monkeypatch.setattr(transcripts, "transcribe_video", fake_video)

    worker._execute_task(task)
    assert task["status"] == transcripts.TASK_STATUS_COMPLETED
    assert captured["start_seconds"] == 10.0
    assert captured["end_seconds"] == 40.0
    # Progress uses the window as denominator: (25 - 10) / (40 - 10) = 0.5.
    assert captured["progress_at_seg"] == pytest.approx(0.5)
    assert task["result"]["start_seconds"] == 10.0
    assert task["result"]["end_seconds"] == 40.0


def test_transcript_worker_unbounded_result_carries_null_window(monkeypatch):
    """No markers → both provenance keys present as None, so a full re-run's
    manifest merge overwrites any previous window rather than leaving it stale."""
    worker = transcripts.TranscriptWorker()
    task = transcripts.create_transcript_task("P01", ["solo.mp4"])
    worker._tasks[task["id"]] = task

    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda p: {"audio_codec": "aac", "duration": 50.0},
    )
    monkeypatch.setattr(video, "timeline_or_none", lambda paths: None)
    captured = {}

    def fake_video(path, **kwargs):
        captured.update(kwargs)
        return {"segments": [], "language": "en", "model": "base", "source_file": path}

    monkeypatch.setattr(transcripts, "transcribe_video", fake_video)

    worker._execute_task(task)
    assert task["status"] == transcripts.TASK_STATUS_COMPLETED
    assert captured["start_seconds"] is None
    assert captured["end_seconds"] is None
    assert task["result"]["start_seconds"] is None
    assert task["result"]["end_seconds"] is None


def test_transcript_worker_fails_window_outside_video(monkeypatch):
    worker = transcripts.TranscriptWorker()
    task = transcripts.create_transcript_task("P01", ["solo.mp4"], start_seconds=60.0)
    worker._tasks[task["id"]] = task

    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda p: {"audio_codec": "aac", "duration": 50.0},
    )
    monkeypatch.setattr(video, "timeline_or_none", lambda paths: None)
    monkeypatch.setattr(
        transcripts, "transcribe_video", lambda *a, **k: pytest.fail("must not run")
    )

    worker._execute_task(task)
    assert task["status"] == transcripts.TASK_STATUS_FAILED
    assert "Marker range" in task["error"]


# ---- Screenspace: _dispatch scans each part with global-offset results ----


def test_screenspace_dispatch_multi_video_offsets_events(monkeypatch):
    import screenspace

    worker = screenspace.ScreenspaceWorker()
    task = screenspace.create_task(
        "color",
        "P01",
        "s",
        ["a.mp4", "b.mp4"],
        "r",
        {"x": 0, "y": 0, "w": 1, "h": 1},
        parameters={"start_seconds": 60.0, "end_seconds": 90.0},
    )
    monkeypatch.setattr(
        video, "timeline_or_none", lambda paths: [("a.mp4", 80, 0), ("b.mp4", 120, 80)]
    )

    scanned = []

    class FakeTool:
        supports_fast_scan = False
        fast_scan_region_dim = 0
        fast_scan_extra_opts: dict = {}

        def scan(self, vp, rc, params, **kw):
            scanned.append((vp, params["start_seconds"], params["end_seconds"]))
            result = {"timestamp": 5.0, "_confidence": 0.9}
            if kw.get("on_result"):
                kw["on_result"](dict(result))
            return [result]

    monkeypatch.setitem(screenspace.TOOLS, "color", FakeTool())

    results = worker._dispatch(task, lambda p: None, lambda: False, None)

    # 60..90 straddles the 80s boundary → tail of part a + head of part b.
    assert scanned == [("a.mp4", 60.0, 80.0), ("b.mp4", 0.0, 10.0)]
    # Returned result times are shifted to the global timeline + tagged per part.
    assert sorted(r["timestamp"] for r in results) == [5.0, 85.0]
    assert {r["_source_video"] for r in results} == {"a.mp4", "b.mp4"}


def test_screenspace_dispatch_single_video_unchanged(monkeypatch):
    import screenspace

    worker = screenspace.ScreenspaceWorker()
    task = screenspace.create_task(
        "color",
        "P01",
        "s",
        ["solo.mp4"],
        "r",
        {"x": 0, "y": 0, "w": 1, "h": 1},
        parameters={"start_seconds": 0.0, "end_seconds": 30.0},
    )
    build = Mock()
    monkeypatch.setattr(video, "build_source_timeline", build)
    scanned = []

    class FakeTool:
        supports_fast_scan = False
        fast_scan_region_dim = 0
        fast_scan_extra_opts: dict = {}

        def scan(self, vp, rc, params, **kw):
            scanned.append((vp, params["start_seconds"], params["end_seconds"]))
            return [{"timestamp": 3.0, "_confidence": 0.5}]

    monkeypatch.setitem(screenspace.TOOLS, "color", FakeTool())

    results = worker._dispatch(task, lambda p: None, lambda: False, None)
    assert scanned == [("solo.mp4", 0.0, 30.0)]  # original file + global range
    assert results[0]["timestamp"] == 3.0  # no offset
    build.assert_not_called()  # single-video fast path: no duration probe


# ---- Review-fix regressions ----


def test_ss_cli_resolves_all_parts(monkeypatch, tmp_path):
    import cli

    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    (tmp_path / "study_P01-1.mp4").write_text("v1")
    (tmp_path / "study_P01-2.mp4").write_text("v2")
    paths = cli._ss_resolve_videos_for_participant("P01")
    assert [_basename(p) for p in paths] == ["study_P01-1.mp4", "study_P01-2.mp4"]
    assert cli._ss_resolve_videos_for_participant("PX") == []


def test_ss_cli_honours_filename_overrides(monkeypatch, tmp_path):
    import cli

    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    (tmp_path / "study_P01.mp4").write_text("pattern")
    (tmp_path / "other.mp4").write_text("override")
    monkeypatch.setattr(
        config, "FILENAME_OVERRIDES", {"P01": "other.mp4"}, raising=False
    )
    paths = cli._ss_resolve_videos_for_participant("P01")
    assert [_basename(p) for p in paths] == ["other.mp4"]
