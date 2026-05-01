from unittest.mock import Mock

import clipgen
import config
import pipeline


def _prepared_clip(raw_clip, times):
    prepared = dict(raw_clip)
    prepared["times"] = list(times)
    return prepared


def test_process_clips_counts_generated_segments_for_all_formats(
    monkeypatch, make_clip
):
    raw_clip = make_clip()
    times = [("00:10", "00:20"), ("00:30", "00:40")]
    monkeypatch.setattr(
        clipgen.files, "prepare_clip", lambda clip: _prepared_clip(clip, times)
    )
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: True)
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)

    call_counter = {"n": 0}

    def unique_name(_template, file_format=None):
        call_counter["n"] += 1
        return f"out_{call_counter['n']}{file_format or '.mp4'}"

    monkeypatch.setattr(clipgen.files, "get_unique_filename", unique_name)

    run_ffmpeg = Mock(return_value=True)
    extract_screenshot = Mock(return_value=True)
    extract_gif = Mock(return_value=True)
    monkeypatch.setattr(clipgen.video, "run_ffmpeg", run_ffmpeg)
    monkeypatch.setattr(clipgen.video, "extract_screenshot", extract_screenshot)
    monkeypatch.setattr(clipgen.video, "extract_gif", extract_gif)

    assert clipgen.process_clips([raw_clip], output_format="clip")[0] == 2
    assert clipgen.process_clips([raw_clip], output_format="screen")[0] == 2
    assert clipgen.process_clips([raw_clip], output_format="gif")[0] == 2

    assert run_ffmpeg.call_count == 2
    assert extract_screenshot.call_count == 2
    assert extract_gif.call_count == 2


def test_prepare_clip_converts_clock_timestamps_to_relative(monkeypatch, make_clip):
    # Arrange a clip with a baseline and clock-style timestamps.
    raw_clip = make_clip(row=3, col=2)
    raw_clip["timestamp_baseline"] = "09:12:00"

    def fake_parse_cell_annotations(value):
        # Return cleaned value and empty annotations.
        return value, {}, set()

    monkeypatch.setattr(
        clipgen.utils, "parse_cell_annotations", fake_parse_cell_annotations
    )
    monkeypatch.setattr(
        clipgen.utils, "has_non_ignored_timestamp_content", lambda _v: True
    )

    # 09:15:00-09:16:30 should become 3:00-4:30 relative to 09:12:00 baseline.
    raw_cell_value = "09:15:00-09:16:30"
    raw_clip["cell"].value = raw_cell_value

    prepared = clipgen.files.prepare_clip(raw_clip)
    assert prepared["times"] == [("0:03:00", "0:04:30")]


def test_process_clips_skips_when_source_video_missing(monkeypatch, make_clip):
    raw_clip = make_clip()
    monkeypatch.setattr(
        clipgen.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: False)
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    run_ffmpeg = Mock(return_value=True)
    monkeypatch.setattr(clipgen.video, "run_ffmpeg", run_ffmpeg)

    assert clipgen.process_clips([raw_clip], output_format="clip")[0] == 0
    run_ffmpeg.assert_not_called()


def test_process_reel_concatenates_and_cleans_temp_parts(monkeypatch, make_clip):
    raw_clips = [make_clip(row=3, col=2), make_clip(row=4, col=2)]
    monkeypatch.setattr(
        clipgen.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)

    generated_parts = []

    def unique_name(_template, file_format=None):
        next_name = f"_reel_part_{len(generated_parts) + 1}{file_format or '.mp4'}"
        generated_parts.append(next_name)
        return next_name

    monkeypatch.setattr(clipgen.files, "get_unique_filename", unique_name)
    monkeypatch.setattr(clipgen.video, "run_ffmpeg", lambda **_kwargs: True)

    def path_is_file(self):
        return str(self).endswith(".mp4")

    monkeypatch.setattr(clipgen.Path, "is_file", path_is_file)

    concat = Mock(return_value=True)
    unlink = Mock()
    monkeypatch.setattr(clipgen.video, "concatenate_clips", concat)
    monkeypatch.setattr(clipgen.Path, "unlink", unlink)

    result, reel_records = clipgen.process_reel(raw_clips, output_file="reel.mp4")
    assert result == 1
    concat.assert_called_once()
    concat_args = concat.call_args.args[0]
    assert concat_args == generated_parts
    assert unlink.call_count == len(generated_parts)

    assert len(reel_records) == 1
    reel = reel_records[0]
    assert reel["id"].startswith("reel_")
    assert reel["file"] == "reel.mp4"
    assert len(reel["components"]) == 2
    assert reel["components"][0]["cellRow"] == 3
    assert reel["components"][1]["cellRow"] == 4


def test_process_reel_returns_zero_when_no_segments_generated(monkeypatch, make_clip):
    raw_clips = [make_clip(row=3, col=2)]
    monkeypatch.setattr(
        clipgen.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(
        clipgen.files,
        "get_unique_filename",
        lambda *_args, **_kwargs: "_reel_part_1.mp4",
    )
    monkeypatch.setattr(clipgen.video, "run_ffmpeg", lambda **_kwargs: False)
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: True)
    concat = Mock(return_value=True)
    monkeypatch.setattr(clipgen.video, "concatenate_clips", concat)

    result, _artifacts = clipgen.process_reel(raw_clips, output_file="reel.mp4")
    assert result == 0
    concat.assert_not_called()


def test_process_clips_parallel_generates_all(monkeypatch, make_clip):
    """Multiple clips processed in parallel should all generate successfully."""
    clips = [make_clip(row=i, col=2) for i in range(3, 7)]
    monkeypatch.setattr(
        clipgen.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: True)
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 4)

    call_counter = {"n": 0}

    def unique_name(_template, file_format=None):
        call_counter["n"] += 1
        return f"out_{call_counter['n']}{file_format or '.mp4'}"

    monkeypatch.setattr(clipgen.files, "get_unique_filename", unique_name)
    run_ffmpeg = Mock(return_value=True)
    monkeypatch.setattr(clipgen.video, "run_ffmpeg", run_ffmpeg)

    count, _artifacts = clipgen.process_clips(clips, output_format="clip")
    assert count == 4
    assert run_ffmpeg.call_count == 4


def test_process_clips_accepts_pre_parsed_synthetic_clip(monkeypatch):
    """Synthetic clips (e.g. --ss-clips) skip prepare_clip's cell parse and
    produce artifacts with deterministic ids derived from the negative cell row."""
    from types import SimpleNamespace

    from utils import ClipRecord

    cell = SimpleNamespace(value="", row=-1, col=1)
    clip: ClipRecord = {
        "cell": cell,
        "study": "mystudy",
        "participant": "P01",
        "category": "screenspace-change",
        "desc": "change region1",
        "severity": "",
        "times": [("0:00:10", "0:00:20")],
        "source_filename": "mystudy_P01.mp4",
    }
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: True)
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    monkeypatch.setattr(
        clipgen.files,
        "get_unique_filename",
        lambda _template, file_format=None: f"out{file_format or '.mp4'}",
    )
    run_ffmpeg = Mock(return_value=True)
    monkeypatch.setattr(clipgen.video, "run_ffmpeg", run_ffmpeg)

    count, artifacts = clipgen.process_clips([clip], output_format="clip")
    assert count == 1
    assert len(artifacts) == 1
    a = artifacts[0]
    # Negative-row + col 1 produces stable id namespace; collision-proof vs.
    # spreadsheet artifacts (which always have positive row/col).
    assert a["id"] == "a-1c1s0"
    assert a["cellRow"] == -1
    assert a["cellCol"] == 1
    # safe_cell_a1 returns "" for negative rows.
    assert a["cellA1"] == ""
    assert run_ffmpeg.call_count == 1


def test_resolve_clip_workers(monkeypatch):
    """Auto-detect returns min(4, cpu_count); explicit values pass through."""
    import pipeline

    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 0)
    monkeypatch.setattr(pipeline.os, "cpu_count", lambda: 8)
    assert pipeline._resolve_clip_workers() == 4

    monkeypatch.setattr(pipeline.os, "cpu_count", lambda: 2)
    assert pipeline._resolve_clip_workers() == 2

    monkeypatch.setattr(pipeline.os, "cpu_count", lambda: None)
    assert pipeline._resolve_clip_workers() == 1

    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 6)
    assert pipeline._resolve_clip_workers() == 6

    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    assert pipeline._resolve_clip_workers() == 1


def test_run_clip_pipeline_cancel_flag(monkeypatch):
    """cancel_flag should stop processing remaining clips."""
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)

    processed = []

    def per_clip_fn(clip, missing_videos):
        processed.append(clip["id"])
        return clip["id"]

    clips = [{"id": i, "desc": "", "participant": ""} for i in range(5)]

    # Cancel after the first clip is processed
    call_count = {"n": 0}

    def cancel_after_first():
        call_count["n"] += 1
        return call_count["n"] > 1

    results, _ = pipeline._run_clip_pipeline(
        clips,
        empty_warning="",
        intro_message="",
        task_label="test",
        per_clip_fn=per_clip_fn,
        cancel_flag=cancel_after_first,
    )
    # Should have processed only the first clip before cancel was detected
    assert len(results) == 1
    assert processed == [0]


def test_process_clips_forwards_cancel_flag_to_segments(monkeypatch, make_clip):
    """process_clips should forward cancel_flag into _process_single_clip_segments."""
    raw_clip = make_clip()
    monkeypatch.setattr(
        clipgen.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: True)
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)

    captured = {}

    def fake_segments(*args, **kwargs):
        captured["cancel_flag"] = kwargs.get("cancel_flag")
        return (1, [])

    monkeypatch.setattr(pipeline, "_process_single_clip_segments", fake_segments)

    sentinel = lambda: False  # noqa: E731
    clipgen.process_clips([raw_clip], output_format="clip", cancel_flag=sentinel)
    assert captured["cancel_flag"] is sentinel


def test_process_clips_sequential_short_circuits_on_cancel(monkeypatch, make_clip):
    """Sequential branch should stop calling _process_single_clip_segments after cancel."""
    clips = [make_clip(row=i, col=2) for i in range(3, 6)]
    monkeypatch.setattr(
        clipgen.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: True)
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)

    cancelled = {"flag": False}

    def seg_side_effect(*_args, **_kwargs):
        # Trip cancel after the first segment finishes; no further calls expected.
        cancelled["flag"] = True
        return (1, [])

    seg_mock = Mock(side_effect=seg_side_effect)
    monkeypatch.setattr(pipeline, "_process_single_clip_segments", seg_mock)

    clipgen.process_clips(
        clips, output_format="clip", cancel_flag=lambda: cancelled["flag"]
    )
    assert seg_mock.call_count == 1


def test_process_single_clip_segments_forwards_cancel_to_video(monkeypatch, make_clip):
    """_process_single_clip_segments should pass cancel_flag to each video helper."""
    raw_clip = _prepared_clip(make_clip(), [("00:10", "00:20")])

    monkeypatch.setattr(
        clipgen.files, "get_unique_filename", lambda *_a, **_k: "out.mp4"
    )

    captured = {}

    def fake_run_ffmpeg(**kwargs):
        captured["clip"] = kwargs.get("cancel_flag")
        return True

    def fake_screenshot(**kwargs):
        captured["screen"] = kwargs.get("cancel_flag")
        return True

    def fake_gif(**kwargs):
        captured["gif"] = kwargs.get("cancel_flag")
        return True

    monkeypatch.setattr(pipeline.video, "run_ffmpeg", fake_run_ffmpeg)
    monkeypatch.setattr(pipeline.video, "extract_screenshot", fake_screenshot)
    monkeypatch.setattr(pipeline.video, "extract_gif", fake_gif)
    monkeypatch.setattr(config, "TITLECARDS_ENABLED", False)

    sentinel = lambda: False  # noqa: E731
    pipeline._process_single_clip_segments(
        raw_clip, "src.mp4", set(), output_format="clip", cancel_flag=sentinel
    )
    pipeline._process_single_clip_segments(
        raw_clip, "src.mp4", set(), output_format="screen", cancel_flag=sentinel
    )
    pipeline._process_single_clip_segments(
        raw_clip, "src.mp4", set(), output_format="gif", cancel_flag=sentinel
    )

    assert captured["clip"] is sentinel
    assert captured["screen"] is sentinel
    assert captured["gif"] is sentinel


def test_process_single_clip_segments_unlinks_partial_on_cancel(
    monkeypatch, make_clip, tmp_path
):
    """When ffmpeg returns False due to cancel, the partial output should be removed."""
    raw_clip = _prepared_clip(make_clip(), [("00:10", "00:20"), ("00:30", "00:40")])

    out_path = tmp_path / "out.mp4"
    monkeypatch.setattr(
        clipgen.files,
        "get_unique_filename",
        lambda *_a, **_k: str(out_path),
    )

    # Simulate ffmpeg writing a partial file then returning False because cancel fired.
    cancel_state = {"set": False}

    def fake_run_ffmpeg(**_kwargs):
        out_path.write_bytes(b"partial")
        cancel_state["set"] = True
        return False

    monkeypatch.setattr(pipeline.video, "run_ffmpeg", fake_run_ffmpeg)
    monkeypatch.setattr(config, "TITLECARDS_ENABLED", False)

    generated, paths = pipeline._process_single_clip_segments(
        raw_clip,
        "src.mp4",
        set(),
        output_format="clip",
        cancel_flag=lambda: cancel_state["set"],
    )

    assert generated == 0
    assert paths == []
    assert not out_path.exists()
