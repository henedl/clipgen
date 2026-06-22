from unittest.mock import Mock

import clipgen
import config
import pipeline
import viewer


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


def test_large_input_videos_cached(tmp_path, monkeypatch):
    """_large_input_videos should glob/stat the input dir once per mtime."""
    pipeline._fuzzy_input_videos_cache.clear()
    monkeypatch.setattr(config, "MIN_SOURCE_VIDEO_SIZE_MB", 1)
    (tmp_path / "study_P01.mp4").write_bytes(b"x" * 2_000_000)
    first = pipeline._large_input_videos(tmp_path)
    second = pipeline._large_input_videos(tmp_path)
    assert first == second
    assert len(first) == 1
    assert first[0][1].name == "study_P01.mp4"


def test_process_clips_forwards_titlecard_options_to_wrap(monkeypatch, make_clip):
    """Per-request titlecard settings reach wrap_clip_with_cards without config overrides."""
    raw_clip = make_clip()
    monkeypatch.setattr(
        clipgen.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: True)
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    monkeypatch.setattr(config, "TITLECARDS_ENABLED", False)
    monkeypatch.setattr(
        clipgen.files, "get_unique_filename", lambda *_a, **_k: "out.mp4"
    )
    monkeypatch.setattr(clipgen.video, "run_ffmpeg", lambda **_k: True)
    monkeypatch.setattr(
        clipgen.video,
        "probe_video_properties",
        lambda *_a, **_k: {"width": 1280, "height": 720},
    )

    captured = {}

    def fake_wrap(_clip, _out_name, **kwargs):
        captured.update(kwargs)
        return True

    monkeypatch.setattr(pipeline.titlecards, "wrap_clip_with_cards", fake_wrap)

    pipeline.process_clips(
        [raw_clip],
        output_format="clip",
        titlecards_enabled=True,
        titlecard_duration_seconds=6,
    )

    assert captured["titlecards_enabled"] is True
    assert captured["titlecard_duration_seconds"] == 6


def test_build_reel_transcript_uses_request_titlecard_duration(monkeypatch):
    """Reel transcript offsets honor per-request titlecard duration."""
    components = [
        {
            "participant": "P01",
            "start": 0.0,
            "end": 10.0,
            "cellRow": 1,
            "cellCol": 2,
        }
    ]
    monkeypatch.setattr(
        pipeline.transcripts,
        "load_transcripts_manifest",
        lambda: {
            "source_transcripts": {
                "P01": {
                    "segments": [{"start": 1.0, "end": 2.0, "text": "hi"}],
                    "language": "en",
                    "source_file": "t.json",
                    "model": "tiny",
                }
            },
            "corrections": [],
        },
    )
    monkeypatch.setattr(
        pipeline.transcripts,
        "apply_corrections",
        lambda segments, _corrections: segments,
    )
    monkeypatch.setattr(
        pipeline.transcripts,
        "filter_segments",
        lambda full, start, end, offset_to_zero=True: {
            "segments": [{"start": 1.0, "end": 2.0, "text": "hi"}]
        },
    )

    merged = pipeline._build_reel_transcript(
        components,
        titlecards_enabled=True,
        titlecard_duration_seconds=7,
    )

    assert merged[0]["start"] == 8.0
    assert merged[0]["end"] == 9.0


def test_process_single_clip_segments_releases_reservation_on_ffmpeg_failure(
    monkeypatch, make_clip, tmp_path
):
    """Failed segment encodes must not leave get_unique_filename placeholders."""
    input_dir = tmp_path / "in"
    output_dir = tmp_path / "out"
    input_dir.mkdir()
    output_dir.mkdir()
    (input_dir / "study_P01.mp4").write_bytes(b"\x00")
    monkeypatch.setattr(config, "INPUT_DIR", str(input_dir), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(output_dir), raising=False)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", lambda **_k: False)

    raw_clip = pipeline.files.prepare_clip(make_clip())
    generated, paths = pipeline._process_single_clip_segments(
        raw_clip, str(input_dir / "study_P01.mp4"), set(), output_format="clip"
    )

    assert generated == 0
    assert paths == []
    assert list(output_dir.iterdir()) == []


def test_process_reel_releases_reservation_on_concat_failure(
    monkeypatch, make_clip, tmp_path
):
    """Concat failure must release the reserved final reel path placeholder."""
    output_dir = tmp_path / "out"
    output_dir.mkdir()
    part_path = output_dir / "part.mp4"
    part_path.write_bytes(b"clip")
    monkeypatch.setattr(config, "OUTPUT_DIR", str(output_dir), raising=False)
    monkeypatch.setattr(
        pipeline,
        "_run_clip_pipeline",
        lambda clips_list, **kwargs: ([([(str(part_path), 0)], [])], set()),
    )
    monkeypatch.setattr(pipeline.video, "concatenate_clips", lambda *a, **k: False)
    monkeypatch.setattr(pipeline.utils, "use_progress", lambda: False)

    clips = [make_clip()]
    result, records = clipgen.process_reel(clips)
    assert result == 0
    assert records == []
    assert list(output_dir.iterdir()) == []


def test_process_reel_releases_caller_reserved_output_when_no_clips(
    monkeypatch, make_clip, tmp_path
):
    """A caller-supplied output reservation (e.g. a chronologic reel) must be
    reclaimed when no clips are generated, not left as a 0-byte placeholder."""
    output_dir = tmp_path / "out"
    output_dir.mkdir()
    monkeypatch.setattr(config, "OUTPUT_DIR", str(output_dir), raising=False)
    monkeypatch.setattr(
        pipeline, "_run_clip_pipeline", lambda clips_list, **kwargs: ([], set())
    )
    monkeypatch.setattr(pipeline.utils, "use_progress", lambda: False)

    reserved = pipeline.files.get_unique_filename("study_P01_chronologic.mp4")
    assert pipeline.Path(reserved).is_file()  # 0-byte placeholder created

    result, records = clipgen.process_reel([make_clip()], output_file=reserved)
    assert result == 0
    assert records == []
    assert not pipeline.Path(reserved).is_file()  # placeholder reclaimed
    assert list(output_dir.iterdir()) == []


def test_regenerate_reel_releases_reservation_on_ffmpeg_failure(monkeypatch, tmp_path):
    """A reel part whose ffmpeg encode fails must not leave a 0-byte placeholder
    behind from get_unique_filename's atomic reservation."""
    input_dir = tmp_path / "in"
    output_dir = tmp_path / "out"
    input_dir.mkdir()
    output_dir.mkdir()
    (input_dir / "study_P01.mp4").write_bytes(b"\x00")
    monkeypatch.setattr(config, "INPUT_DIR", str(input_dir), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(output_dir), raising=False)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", lambda **_k: False)

    reel = {
        "file": "reel.mp4",
        "components": [{"sourceVideo": "study_P01.mp4", "start": 0, "end": 10}],
    }
    ok = pipeline._regenerate_reel(reel, set())

    assert ok is False
    # The reserved placeholder for the failed part was cleaned up.
    assert list(output_dir.iterdir()) == []


def test_regenerate_from_manifest_parallel(monkeypatch):
    """Independent artifacts regenerate concurrently when workers >= 2."""
    artifacts = [
        {
            "id": f"a{i}",
            "type": "clip",
            "file": f"clip{i}.mp4",
            "sourceVideo": "study_P01.mp4",
            "start": 0,
            "end": 10,
            "description": f"clip {i}",
            "participant": "P01",
        }
        for i in range(4)
    ]
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 4)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(pipeline.utils, "print_mode_heading", lambda *a, **k: None)

    calls: list[str] = []

    def fake_regenerate(artifact, missing_videos):
        calls.append(artifact["id"])
        return True

    monkeypatch.setattr(pipeline, "_regenerate_single_artifact", fake_regenerate)

    count = pipeline.regenerate_from_manifest(artifacts)
    assert count == 4
    assert sorted(calls) == ["a0", "a1", "a2", "a3"]


def test_process_clips_skips_titlecard_cache_clear_when_disabled(
    monkeypatch, make_clip
):
    """clear_titlecard_cache=False keeps a per-cell worker from purging the
    shared endcard cache mid-flight; the default True clears it once."""
    raw_clip = make_clip()
    monkeypatch.setattr(
        clipgen.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: True)
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    monkeypatch.setattr(config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(
        clipgen.files, "get_unique_filename", lambda *_a, **_k: "out.mp4"
    )
    monkeypatch.setattr(clipgen.video, "run_ffmpeg", lambda **_k: True)
    monkeypatch.setattr(
        clipgen.video,
        "probe_video_properties",
        lambda *_a, **_k: {"width": 1280, "height": 720},
    )
    monkeypatch.setattr(
        pipeline.titlecards, "wrap_clip_with_cards", lambda *_a, **_k: True
    )

    clear_mock = Mock()
    monkeypatch.setattr(pipeline.titlecards, "clear_endcard_cache", clear_mock)

    clipgen.process_clips([raw_clip], output_format="clip", clear_titlecard_cache=False)
    clear_mock.assert_not_called()

    clipgen.process_clips([raw_clip], output_format="clip")
    clear_mock.assert_called_once()


def test_process_reel_forwards_cancel_and_titlecard_options_to_segments(
    monkeypatch, make_clip
):
    """process_reel must pass its cancel_flag and per-request titlecard options
    into _process_single_clip_segments for each reel part."""
    raw_clip = make_clip()
    monkeypatch.setattr(
        clipgen.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: True)
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)

    captured = {}

    def fake_segments(*_args, **kwargs):
        captured.update(kwargs)
        return (1, [("_reel_part_1.mp4", 0)])

    monkeypatch.setattr(pipeline, "_process_single_clip_segments", fake_segments)
    monkeypatch.setattr(clipgen.video, "concatenate_clips", lambda *_a, **_k: True)
    monkeypatch.setattr(pipeline, "_build_reel_transcript", lambda *_a, **_k: [])

    sentinel = lambda: False  # noqa: E731
    clipgen.process_reel(
        [raw_clip],
        output_file="reel.mp4",
        cancel_flag=sentinel,
        titlecards_enabled=True,
        titlecard_duration_seconds=5,
    )

    assert captured["cancel_flag"] is sentinel
    assert captured["titlecards_enabled"] is True
    assert captured["titlecard_duration_seconds"] == 5


def test_process_reel_dedups_missing_video_across_parallel_clips(
    monkeypatch, make_clip
):
    """Parallel reel clips that all reference the same missing source video
    produce exactly one missing-video error, not one per worker thread."""
    clips = [make_clip(row=i, col=2) for i in range(3, 7)]
    monkeypatch.setattr(
        clipgen.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 4)
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: False)
    monkeypatch.setattr(pipeline, "_large_input_videos", lambda _d: [])

    errors: list[tuple] = []
    monkeypatch.setattr(
        pipeline.utils, "error_print", lambda *a, **_k: errors.append(a)
    )

    result, _ = clipgen.process_reel(clips, output_file="reel.mp4")

    assert result == 0
    assert len(errors) == 1


def test_build_artifact_records_for_clip_stamps_titlecard_state(make_clip):
    """Clip artifacts record their titlecard state; screen/gif artifacts do not."""
    clip = _prepared_clip(make_clip(), [("00:10", "00:20")])
    segment_details = [("out.mp4", 0)]

    clip_recs = viewer.build_artifact_records_for_clip(
        clip,
        "study_P01.mp4",
        segment_details,
        "clip",
        titlecards=True,
        titlecard_duration=5,
    )
    assert clip_recs[0]["titlecards"] is True
    assert clip_recs[0]["titlecardDuration"] == 5

    screen_recs = viewer.build_artifact_records_for_clip(
        clip,
        "study_P01.mp4",
        segment_details,
        "screen",
        titlecards=True,
        titlecard_duration=5,
    )
    assert "titlecards" not in screen_recs[0]
    assert "titlecardDuration" not in screen_recs[0]


def test_regenerate_single_artifact_applies_titlecards_from_manifest(monkeypatch):
    """A clip artifact recorded with titlecards is re-wrapped on regeneration;
    one without titlecards is not."""
    monkeypatch.setattr(pipeline.utils, "resolve_input_path", lambda p: p)
    monkeypatch.setattr(pipeline.utils, "resolve_output_path", lambda p: p)
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", lambda **_k: True)

    wrap_calls: list[tuple] = []

    def fake_wrap(clip, out_path, **kwargs):
        wrap_calls.append((clip, out_path, kwargs))
        return True

    monkeypatch.setattr(pipeline.titlecards, "wrap_clip_with_cards", fake_wrap)

    with_cards = {
        "type": "clip",
        "file": "clip.mp4",
        "sourceVideo": "study_P01.mp4",
        "start": 0,
        "end": 10,
        "description": "an observation",
        "titlecards": True,
        "titlecardDuration": 4,
    }
    assert pipeline._regenerate_single_artifact(with_cards, set()) is True
    assert len(wrap_calls) == 1
    clip_arg, _out_arg, kwargs = wrap_calls[0]
    assert clip_arg["desc"] == "an observation"
    assert kwargs["titlecards_enabled"] is True
    assert kwargs["titlecard_duration_seconds"] == 4

    # Without the titlecards flag, no wrap happens.
    wrap_calls.clear()
    without_cards = dict(with_cards)
    without_cards["titlecards"] = False
    assert pipeline._regenerate_single_artifact(without_cards, set()) is True
    assert wrap_calls == []
