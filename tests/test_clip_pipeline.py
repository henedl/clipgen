from unittest.mock import Mock

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
        pipeline.files, "prepare_clip", lambda clip: _prepared_clip(clip, times)
    )
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)

    call_counter = {"n": 0}

    def unique_name(_template, file_format=None):
        call_counter["n"] += 1
        return f"out_{call_counter['n']}{file_format or '.mp4'}"

    monkeypatch.setattr(pipeline.files, "get_unique_filename", unique_name)

    run_ffmpeg = Mock(return_value=True)
    extract_screenshot = Mock(return_value=True)
    extract_gif = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", run_ffmpeg)
    monkeypatch.setattr(pipeline.video, "extract_screenshot", extract_screenshot)
    monkeypatch.setattr(pipeline.video, "extract_gif", extract_gif)

    assert pipeline.process_clips([raw_clip], output_format="clip")[0] == 2
    assert pipeline.process_clips([raw_clip], output_format="screen")[0] == 2
    assert pipeline.process_clips([raw_clip], output_format="gif")[0] == 2

    assert run_ffmpeg.call_count == 2
    assert extract_screenshot.call_count == 2
    assert extract_gif.call_count == 2


def _padding_test_setup(monkeypatch, make_clip, times):
    raw_clip = make_clip()
    monkeypatch.setattr(
        pipeline.files, "prepare_clip", lambda clip: _prepared_clip(clip, times)
    )
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    monkeypatch.setattr(
        pipeline.files,
        "get_unique_filename",
        lambda _t, file_format=None: f"out{file_format or '.mp4'}",
    )
    # Large EOF so pad_post never clamps in these assertions.
    monkeypatch.setattr(pipeline.video, "get_file_duration", lambda _p: 10_000)
    run_ffmpeg = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", run_ffmpeg)
    return raw_clip, run_ffmpeg


def test_process_clips_pads_and_clamps_cut_timestamps(monkeypatch, make_clip):
    raw_clip, run_ffmpeg = _padding_test_setup(
        monkeypatch,
        make_clip,
        [("00:10", "01:10")],  # 60s span
    )
    # pad start 3s earlier, end 2s later → (7, 72); max_duration 10 caps end to 17.
    pipeline.process_clips(
        [raw_clip], output_format="clip", pad_pre=3.0, pad_post=2.0, max_duration=10.0
    )
    _, kwargs = run_ffmpeg.call_args
    assert kwargs["start_pos"] == "0:00:07"
    assert kwargs["end_pos"] == "0:00:17"


def test_process_clips_no_padding_leaves_timestamps_untouched(monkeypatch, make_clip):
    raw_clip, run_ffmpeg = _padding_test_setup(
        monkeypatch, make_clip, [("00:10", "00:20")]
    )
    pipeline.process_clips([raw_clip], output_format="clip")
    _, kwargs = run_ffmpeg.call_args
    # Default (no-op) path passes the original strings straight through.
    assert kwargs["start_pos"] == "00:10"
    assert kwargs["end_pos"] == "00:20"


def test_process_clips_gif_fractional_max_duration_floors_to_one(
    monkeypatch, make_clip
):
    raw_clip = make_clip()
    monkeypatch.setattr(
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    monkeypatch.setattr(
        pipeline.files,
        "get_unique_filename",
        lambda _t, file_format=None: f"out{file_format or '.gif'}",
    )
    extract_gif = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "extract_gif", extract_gif)
    # A sub-1s cap must floor the gif to 1s, never truncate to a 0s (invalid) gif.
    pipeline.process_clips([raw_clip], output_format="gif", max_duration=0.5)
    _, kwargs = extract_gif.call_args
    assert kwargs["duration_seconds"] == 1


def test_prepare_clip_converts_clock_timestamps_to_relative(monkeypatch, make_clip):
    # Arrange a clip with a baseline and clock-style timestamps.
    raw_clip = make_clip(row=3, col=2)
    raw_clip["timestamp_baseline"] = "09:12:00"

    def fake_parse_cell_annotations(value):
        # Return cleaned value and empty annotations.
        return value, {}, set()

    monkeypatch.setattr(
        pipeline.utils, "parse_cell_annotations", fake_parse_cell_annotations
    )
    monkeypatch.setattr(
        pipeline.utils, "has_non_ignored_timestamp_content", lambda _v: True
    )

    # 09:15:00-09:16:30 should become 3:00-4:30 relative to 09:12:00 baseline.
    raw_cell_value = "09:15:00-09:16:30"
    raw_clip["cell"].value = raw_cell_value

    prepared = pipeline.files.prepare_clip(raw_clip)
    assert prepared["times"] == [("0:03:00", "0:04:30")]


def test_process_clips_skips_when_source_video_missing(monkeypatch, make_clip):
    raw_clip = make_clip()
    monkeypatch.setattr(
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: False)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    run_ffmpeg = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", run_ffmpeg)

    assert pipeline.process_clips([raw_clip], output_format="clip")[0] == 0
    run_ffmpeg.assert_not_called()


def test_process_reel_concatenates_and_cleans_temp_parts(monkeypatch, make_clip):
    raw_clips = [make_clip(row=3, col=2), make_clip(row=4, col=2)]
    monkeypatch.setattr(
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)

    generated_parts = []

    def unique_name(_template, file_format=None):
        next_name = f"_reel_part_{len(generated_parts) + 1}{file_format or '.mp4'}"
        generated_parts.append(next_name)
        return next_name

    monkeypatch.setattr(pipeline.files, "get_unique_filename", unique_name)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", lambda **_kwargs: True)

    def path_is_file(self):
        return str(self).endswith(".mp4")

    monkeypatch.setattr(pipeline.Path, "is_file", path_is_file)

    concat = Mock(return_value=True)
    unlink = Mock()
    monkeypatch.setattr(pipeline.video, "concatenate_clips", concat)
    monkeypatch.setattr(pipeline.Path, "unlink", unlink)

    result, reel_records = pipeline.process_reel(raw_clips, output_file="reel.mp4")
    assert result == 1
    concat.assert_called_once()
    concat_args = concat.call_args.args[0]
    # process_reel generates segments via a ThreadPoolExecutor, so the order in
    # which unique_name() appends to generated_parts is nondeterministic. Compare
    # as a set: the parts-in-segment-order contract is covered by the cellRow
    # assertions on reel["components"] below.
    assert sorted(concat_args) == sorted(generated_parts)
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
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(
        pipeline.files,
        "get_unique_filename",
        lambda *_args, **_kwargs: "_reel_part_1.mp4",
    )
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", lambda **_kwargs: False)
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    concat = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "concatenate_clips", concat)

    result, _artifacts = pipeline.process_reel(raw_clips, output_file="reel.mp4")
    assert result == 0
    concat.assert_not_called()


def test_process_clips_parallel_generates_all(monkeypatch, make_clip):
    """Multiple clips processed in parallel should all generate successfully."""
    clips = [make_clip(row=i, col=2) for i in range(3, 7)]
    monkeypatch.setattr(
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 4)

    call_counter = {"n": 0}

    def unique_name(_template, file_format=None):
        call_counter["n"] += 1
        return f"out_{call_counter['n']}{file_format or '.mp4'}"

    monkeypatch.setattr(pipeline.files, "get_unique_filename", unique_name)
    run_ffmpeg = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", run_ffmpeg)

    count, _artifacts = pipeline.process_clips(clips, output_format="clip")
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
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    monkeypatch.setattr(
        pipeline.files,
        "get_unique_filename",
        lambda _template, file_format=None: f"out{file_format or '.mp4'}",
    )
    run_ffmpeg = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", run_ffmpeg)

    count, artifacts = pipeline.process_clips([clip], output_format="clip")
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
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)

    captured = {}

    def fake_segments(*args, **kwargs):
        captured["cancel_flag"] = kwargs.get("cancel_flag")
        return (1, [], True)

    monkeypatch.setattr(pipeline, "_process_single_clip_segments", fake_segments)

    sentinel = lambda: False
    pipeline.process_clips([raw_clip], output_format="clip", cancel_flag=sentinel)
    assert captured["cancel_flag"] is sentinel


def test_process_clips_sequential_short_circuits_on_cancel(monkeypatch, make_clip):
    """Sequential branch should stop calling _process_single_clip_segments after cancel."""
    clips = [make_clip(row=i, col=2) for i in range(3, 6)]
    monkeypatch.setattr(
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)

    cancelled = {"flag": False}

    def seg_side_effect(*_args, **_kwargs):
        # Trip cancel after the first segment finishes; no further calls expected.
        cancelled["flag"] = True
        return (1, [], True)

    seg_mock = Mock(side_effect=seg_side_effect)
    monkeypatch.setattr(pipeline, "_process_single_clip_segments", seg_mock)

    pipeline.process_clips(
        clips, output_format="clip", cancel_flag=lambda: cancelled["flag"]
    )
    assert seg_mock.call_count == 1


def test_process_single_clip_segments_forwards_cancel_to_video(monkeypatch, make_clip):
    """_process_single_clip_segments should pass cancel_flag to each video helper."""
    raw_clip = _prepared_clip(make_clip(), [("00:10", "00:20")])

    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *_a, **_k: "out.mp4"
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

    sentinel = lambda: False
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
        pipeline.files,
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

    generated, paths, _ = pipeline._process_single_clip_segments(
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
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    monkeypatch.setattr(config, "TITLECARDS_ENABLED", False)
    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *_a, **_k: "out.mp4"
    )
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", lambda **_k: True)
    monkeypatch.setattr(
        pipeline.video,
        "probe_video_properties",
        lambda *_a, **_k: {"width": 1280, "height": 720},
    )

    captured = {}

    def fake_wrap(_clip, _out_name, **kwargs):
        captured.update(kwargs)
        return (True, True)

    monkeypatch.setattr(pipeline.titlecards, "wrap_clip_with_cards", fake_wrap)

    pipeline.process_clips(
        [raw_clip],
        output_format="clip",
        titlecards_enabled=True,
        titlecard_duration_seconds=6,
    )

    assert captured["titlecards_enabled"] is True
    assert captured["titlecard_duration_seconds"] == 6


def test_process_clips_compresses_after_titlecard_wrap(monkeypatch, make_clip):
    """The size cap is enforced on the final clip AFTER the wrap, not before it.

    Regression for the compress-before-wrap bug: compression used to run inside
    run_ffmpeg (before the CRF wrap discarded it), wasting passes and letting the
    output exceed the cap.
    """
    raw_clip = make_clip()
    monkeypatch.setattr(
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    monkeypatch.setattr(config, "MAX_FILESIZE_MB", 50)
    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *_a, **_k: "out.mp4"
    )
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", lambda **_k: True)

    order: list[str] = []

    def fake_wrap(_clip, _out_name, **_kwargs):
        order.append("wrap")
        return (True, True)

    def fake_enforce(_path, **_kwargs):
        order.append("enforce")

    monkeypatch.setattr(pipeline.titlecards, "wrap_clip_with_cards", fake_wrap)
    monkeypatch.setattr(pipeline.video, "enforce_filesize_limit", fake_enforce)

    pipeline.process_clips([raw_clip], output_format="clip", titlecards_enabled=True)

    assert order == ["wrap", "enforce"]


def test_reel_part_cut_is_not_size_capped(monkeypatch, make_clip):
    """Reel parts (enforce_size=False) skip compression; the cap applies to the reel-less final clip only."""
    monkeypatch.setattr(config, "MAX_FILESIZE_MB", 50)
    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *_a, **_k: "out.mp4"
    )
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", lambda **_k: True)
    monkeypatch.setattr(config, "TITLECARDS_ENABLED", False)

    enforce = Mock()
    monkeypatch.setattr(pipeline.video, "enforce_filesize_limit", enforce)

    prepared = _prepared_clip(make_clip(), [("00:10", "00:20")])

    # enforce_size=False (reel-part path) → no compression.
    pipeline._process_single_clip_segments(
        prepared, "src.mp4", set(), collect_paths=True, enforce_size=False
    )
    assert enforce.call_count == 0

    # enforce_size defaults True (final clip) → compression runs once.
    pipeline._process_single_clip_segments(
        prepared, "src.mp4", set(), collect_paths=True
    )
    assert enforce.call_count == 1


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


def _mock_two_component_transcript(monkeypatch):
    """Two participants, one segment each at local 1.0–2.0; passthrough filters."""
    monkeypatch.setattr(
        pipeline.transcripts,
        "load_transcripts_manifest",
        lambda: {
            "source_transcripts": {
                "P01": {"segments": [{"start": 1.0, "end": 2.0, "text": "a"}]},
                "P02": {"segments": [{"start": 1.0, "end": 2.0, "text": "b"}]},
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
            "segments": [{"start": 1.0, "end": 2.0, "text": "seg"}]
        },
    )


def test_build_reel_transcript_advances_offset_by_endcard(monkeypatch):
    """The second component starts after clip + titlecard + endcard of the first."""
    _mock_two_component_transcript(monkeypatch)
    # Endcard renders (skip=False) → cumulative offset includes its duration.
    monkeypatch.setattr(
        pipeline.titlecards,
        "resolve_card_background",
        lambda kind: (None, True, False, "black"),
    )
    components = [
        {"participant": "P01", "start": 0.0, "end": 10.0},
        {"participant": "P02", "start": 0.0, "end": 8.0},
    ]

    merged = pipeline._build_reel_transcript(
        components, titlecards_enabled=True, titlecard_duration_seconds=7
    )

    # comp1 segment: 1.0 + offset(0) + titlecard(7)
    assert merged[0]["start"] == 8.0
    # offset after comp1 = clip(10) + titlecard(7) + endcard(7) = 24
    # comp2 segment: 1.0 + offset(24) + titlecard(7) = 32
    assert merged[1]["start"] == 32.0
    assert merged[1]["end"] == 33.0


def test_build_reel_transcript_skips_endcard_when_none(monkeypatch):
    """A 'none' endcard adds no outro duration to the cumulative offset."""
    _mock_two_component_transcript(monkeypatch)
    # Endcard skipped (skip=True) → cumulative offset omits its duration.
    monkeypatch.setattr(
        pipeline.titlecards,
        "resolve_card_background",
        lambda kind: (None, False, True, "black"),
    )
    components = [
        {"participant": "P01", "start": 0.0, "end": 10.0},
        {"participant": "P02", "start": 0.0, "end": 8.0},
    ]

    merged = pipeline._build_reel_transcript(
        components, titlecards_enabled=True, titlecard_duration_seconds=7
    )

    # offset after comp1 = clip(10) + titlecard(7) + endcard(0) = 17
    # comp2 segment: 1.0 + offset(17) + titlecard(7) = 25
    assert merged[1]["start"] == 25.0
    assert merged[1]["end"] == 26.0


def test_process_reel_clears_endcard_cache():
    """process_reel purges the shared per-process endcard temp cache on exit.

    Covers the CLI, interactive, and Studio /api/reel paths, which all route
    through process_reel; previously only process_clips and /api/reel-direct
    cleared it, so sheet/CLI reels leaked cached endcards.
    """
    pipeline.titlecards._endcard_cache["k"] = "/tmp/nonexistent_endcard.mp4"
    try:
        # An empty clip list returns early but still runs the wrapper's finally.
        assert pipeline.process_reel([]) == (0, [])
        assert pipeline.titlecards._endcard_cache == {}
    finally:
        pipeline.titlecards._endcard_cache.clear()


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
    generated, paths, _ = pipeline._process_single_clip_segments(
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
        lambda clips_list, **kwargs: ([([(str(part_path), 0)], [], [], True)], set()),
    )
    monkeypatch.setattr(pipeline.video, "concatenate_clips", lambda *a, **k: False)
    monkeypatch.setattr(pipeline.utils, "use_progress", lambda: False)

    clips = [make_clip()]
    result, records = pipeline.process_reel(clips)
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

    result, records = pipeline.process_reel([make_clip()], output_file=reserved)
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


def test_regenerate_reel_aborts_when_component_source_missing(monkeypatch, tmp_path):
    """A missing source for any component must fail the whole reel rather than
    silently concatenating the surviving components into a shorter reel."""
    input_dir = tmp_path / "in"
    output_dir = tmp_path / "out"
    input_dir.mkdir()
    output_dir.mkdir()
    (input_dir / "study_P01.mp4").write_bytes(b"\x00")
    monkeypatch.setattr(config, "INPUT_DIR", str(input_dir), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(output_dir), raising=False)

    # The present component would cut fine; the second source is absent.
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", lambda **_k: True)
    concat_calls: list = []
    monkeypatch.setattr(
        pipeline.video,
        "concatenate_clips",
        lambda *a, **k: concat_calls.append(a) or True,
    )

    missing: set[str] = set()
    reel = {
        "file": "reel.mp4",
        "components": [
            {"sourceVideo": "study_P01.mp4", "start": 0, "end": 10},
            {"sourceVideo": "gone_P02.mp4", "start": 0, "end": 10},
        ],
    }
    ok = pipeline._regenerate_reel(reel, missing)

    assert ok is False
    assert concat_calls == []  # never concatenated a partial reel
    assert any("gone_P02.mp4" in p for p in missing)  # missing source recorded
    # No truncated output or leftover temp segments remain.
    assert list(output_dir.iterdir()) == []


def test_regenerate_reel_aborts_on_component_ffmpeg_failure(monkeypatch, tmp_path):
    """An ffmpeg failure on any component aborts the reel without concatenating
    or leaving temp segments behind."""
    input_dir = tmp_path / "in"
    output_dir = tmp_path / "out"
    input_dir.mkdir()
    output_dir.mkdir()
    (input_dir / "study_P01.mp4").write_bytes(b"\x00")
    (input_dir / "study_P02.mp4").write_bytes(b"\x00")
    monkeypatch.setattr(config, "INPUT_DIR", str(input_dir), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(output_dir), raising=False)

    # First component's cut succeeds, second fails.
    calls = {"n": 0}

    def fake_ffmpeg(**_k):
        calls["n"] += 1
        return calls["n"] == 1

    monkeypatch.setattr(pipeline.video, "run_ffmpeg", fake_ffmpeg)
    concat_calls: list = []
    monkeypatch.setattr(
        pipeline.video,
        "concatenate_clips",
        lambda *a, **k: concat_calls.append(a) or True,
    )

    reel = {
        "file": "reel.mp4",
        "components": [
            {"sourceVideo": "study_P01.mp4", "start": 0, "end": 10},
            {"sourceVideo": "study_P02.mp4", "start": 0, "end": 10},
        ],
    }
    ok = pipeline._regenerate_reel(reel, set())

    assert ok is False
    assert concat_calls == []  # never concatenated a partial reel
    # The first component's temp segment and the failed reservation are cleaned up.
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
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    monkeypatch.setattr(config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *_a, **_k: "out.mp4"
    )
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", lambda **_k: True)
    monkeypatch.setattr(
        pipeline.video,
        "probe_video_properties",
        lambda *_a, **_k: {"width": 1280, "height": 720},
    )
    monkeypatch.setattr(
        pipeline.titlecards, "wrap_clip_with_cards", lambda *_a, **_k: (True, True)
    )

    clear_mock = Mock()
    monkeypatch.setattr(pipeline.titlecards, "clear_endcard_cache", clear_mock)

    pipeline.process_clips(
        [raw_clip], output_format="clip", clear_titlecard_cache=False
    )
    clear_mock.assert_not_called()

    pipeline.process_clips([raw_clip], output_format="clip")
    clear_mock.assert_called_once()


def test_process_reel_forwards_cancel_and_titlecard_options_to_segments(
    monkeypatch, make_clip
):
    """process_reel must pass its cancel_flag and per-request titlecard options
    into _process_single_clip_segments for each reel part."""
    raw_clip = make_clip()
    monkeypatch.setattr(
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)

    captured = {}

    def fake_segments(*_args, **kwargs):
        captured.update(kwargs)
        return (1, [("_reel_part_1.mp4", 0)], True)

    monkeypatch.setattr(pipeline, "_process_single_clip_segments", fake_segments)
    monkeypatch.setattr(pipeline.video, "concatenate_clips", lambda *_a, **_k: True)
    monkeypatch.setattr(pipeline, "_build_reel_transcript", lambda *_a, **_k: [])

    sentinel = lambda: False
    pipeline.process_reel(
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
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 4)
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: False)
    monkeypatch.setattr(pipeline, "_large_input_videos", lambda _d: [])

    errors: list[tuple] = []
    monkeypatch.setattr(
        pipeline.utils, "error_print", lambda *a, **_k: errors.append(a)
    )

    result, _ = pipeline.process_reel(clips, output_file="reel.mp4")

    assert result == 0
    # The missing-video error is reported once, not once per worker thread.
    missing_errors = [e for e in errors if "Source video file not found" in e[0]]
    assert len(missing_errors) == 1
    # ...and the reel aborts rather than shipping a reel built from no clips.
    assert any("Reel aborted" in e[0] for e in errors)


def test_process_reel_aborts_when_a_clip_fails_to_cut(monkeypatch, make_clip):
    """A reel is all-or-nothing: one failed segment must not ship a short reel.

    Concatenating the survivors produces a video that looks complete but silently
    omits moments, and `compute_reel_id` hashes the truncated component list, so the
    generate-cache would then serve that truncated reel even after the source is
    fixed. Guards the abort path.
    """
    clips = [make_clip(row=3, col=2), make_clip(row=4, col=2)]
    monkeypatch.setattr(
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline, "_large_input_videos", lambda _d: [])

    # Second clip's cut fails; the first succeeds.
    calls: list[str] = []

    def fake_ffmpeg(**kwargs):
        calls.append(kwargs["output_file"])
        return len(calls) == 1

    monkeypatch.setattr(pipeline.video, "run_ffmpeg", fake_ffmpeg)

    concat_calls: list = []
    monkeypatch.setattr(
        pipeline.video,
        "concatenate_clips",
        lambda *a, **k: concat_calls.append(a) or True,
    )

    errors: list[tuple] = []
    monkeypatch.setattr(
        pipeline.utils, "error_print", lambda *a, **_k: errors.append(a)
    )

    generated, records = pipeline.process_reel(clips, output_file="reel.mp4")

    assert generated == 0
    assert records == []
    # Crucially: no concatenation was attempted, so no partial reel exists.
    assert concat_calls == []
    assert any("Reel aborted" in e[0] for e in errors)


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
        return (True, True)

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


# ---- pure helpers ----


def test_is_excel_worksheet_true_for_local(make_clip):
    from types import SimpleNamespace

    excel = SimpleNamespace(spreadsheet=SimpleNamespace(url=None))
    assert pipeline.is_excel_worksheet(excel) is True


def test_is_excel_worksheet_false_for_gsheet_and_missing(make_clip):
    from types import SimpleNamespace

    gsheet = SimpleNamespace(spreadsheet=SimpleNamespace(url="https://x"))
    assert pipeline.is_excel_worksheet(gsheet) is False
    assert pipeline.is_excel_worksheet(SimpleNamespace()) is False


def test_resolve_clip_workers_explicit_and_auto(monkeypatch):
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 7)
    assert pipeline._resolve_clip_workers() == 7

    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 0)
    monkeypatch.setattr(pipeline.os, "cpu_count", lambda: 16)
    assert pipeline._resolve_clip_workers() == 4  # capped at 4

    monkeypatch.setattr(pipeline.os, "cpu_count", lambda: 2)
    assert pipeline._resolve_clip_workers() == 2  # below cap -> cpu count


def test_resolve_titlecard_options_defaults_and_overrides(monkeypatch):
    monkeypatch.setattr(config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(config, "TITLECARD_DURATION_SECONDS", 3)
    # None -> fall back to config.
    assert pipeline._resolve_titlecard_options(None, None) == (True, 3)
    # Explicit values win over the config defaults.
    assert pipeline._resolve_titlecard_options(False, 5) == (False, 5)


def test_compute_reel_id_is_deterministic_and_order_independent():
    a = {"cellRow": 3, "cellCol": 2, "start": "0", "end": "10"}
    b = {"cellRow": 4, "cellCol": 2, "start": "5", "end": "15"}
    id_ab = pipeline.compute_reel_id([a, b])
    id_ba = pipeline.compute_reel_id([b, a])
    assert id_ab == id_ba  # sorted internally
    assert id_ab.startswith("reel_")
    # A different component set yields a different id.
    c = {"cellRow": 9, "cellCol": 9, "start": "0", "end": "1"}
    assert pipeline.compute_reel_id([a, c]) != id_ab


def test_run_clip_pipeline_cancel_captures_started_clip_results(monkeypatch):
    """On cancel, results from futures that already started running must not be
    dropped — otherwise the files they produced (e.g. reel _reel_part_* segments)
    are orphaned because the caller's cancel cleanup never sees their paths."""
    import threading
    import time

    monkeypatch.setattr(pipeline, "_resolve_clip_workers", lambda: 2)

    ran: list[int] = []
    ran_lock = threading.Lock()

    def per_clip(clip, _missing):
        with ran_lock:
            ran.append(clip["id"])
        time.sleep(0.1)  # keep the future running across the cancel
        return ([(f"/tmp/_reel_part_{clip['id']}.mp4", 0)], [])

    def cancel_flag():
        # Cancel as soon as at least one clip has begun executing.
        with ran_lock:
            return len(ran) >= 1

    clips = [{"id": i, "desc": f"c{i}", "participant": "P01"} for i in range(6)]
    results, _missing = pipeline._run_clip_pipeline(
        clips,
        empty_warning="",
        intro_message="",
        task_label="t",
        per_clip_fn=per_clip,
        parallel=True,
        cancel_flag=cancel_flag,
    )

    # Every clip that started running is represented in the results (none
    # orphaned); clips cancelled before starting are absent.
    assert len(results) == len(ran)
    assert all(r is not None for r in results)
    # Each captured result is the proper (segment_paths, components) tuple shape.
    for segment_paths, components in results:
        assert isinstance(segment_paths, list)


def test_parallel_map_ordered_collects_in_original_order():
    items = [3, 1, 2, 4]
    results: list = [None] * len(items)
    pipeline._parallel_map_ordered(
        items,
        lambda n: n * 10,
        workers=4,
        results=results,
        on_error=lambda idx, exc: "err",
    )
    assert results == [30, 10, 20, 40]


def test_parallel_map_ordered_on_error_fills_failing_slot():
    items = [1, 2, 3]
    results: list = [None] * len(items)

    def worker(n):
        if n == 2:
            raise ValueError("boom")
        return n

    pipeline._parallel_map_ordered(
        items,
        worker,
        workers=3,
        results=results,
        on_error=lambda idx, exc: f"failed:{idx}",
    )
    assert results == [1, "failed:1", 3]


def test_parallel_map_ordered_cancel_leaves_prefilled_sentinel():
    items = list(range(6))
    sentinel = (0, [])
    results: list = [sentinel] * len(items)

    # Cancel immediately: the as_completed loop breaks before collecting, so
    # every slot keeps the caller's pre-filled sentinel.
    pipeline._parallel_map_ordered(
        items,
        lambda n: (1, [str(n)]),
        workers=2,
        results=results,
        on_error=lambda idx, exc: sentinel,
        cancel_flag=lambda: True,
    )
    assert all(r == sentinel for r in results)


def test_process_reel_records_cards_that_actually_landed(monkeypatch, make_clip):
    """A part whose card wrap soft-failed must not be recorded as carded.

    `wrap_clip_with_cards` leaves a usable but *unwrapped* clip on a soft
    failure. If the reel record still claims `titlecards: true`, the generate
    cache matches on that flag and skips rebuilding, so the card can never be
    applied — the same bug fixed for process_clips and /api/reel-direct.
    """
    raw_clip = make_clip()
    monkeypatch.setattr(
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(pipeline.video, "concatenate_clips", lambda *_a, **_k: True)
    monkeypatch.setattr(pipeline, "_build_reel_transcript", lambda *_a, **_k: [])

    # The part was cut, but its wrap soft-failed → cards_applied False.
    monkeypatch.setattr(
        pipeline,
        "_process_single_clip_segments",
        lambda *_a, **_k: (1, [("_reel_part_1.mp4", 0)], False),
    )
    generated, records = pipeline.process_reel(
        [raw_clip], output_file="reel.mp4", titlecards_enabled=True
    )
    assert generated == 1
    assert records[0]["titlecards"] is False
    assert records[0]["titlecardDuration"] == 0

    # Control: when the wrap succeeds the reel is recorded as carded.
    monkeypatch.setattr(
        pipeline,
        "_process_single_clip_segments",
        lambda *_a, **_k: (1, [("_reel_part_2.mp4", 0)], True),
    )
    _generated, records = pipeline.process_reel(
        [raw_clip], output_file="reel2.mp4", titlecards_enabled=True
    )
    assert records[0]["titlecards"] is True


def test_completion_message_reports_aborted_reel(monkeypatch, capsys):
    import app

    monkeypatch.setattr(app.utils, "get_effective_output_dir", lambda: "/tmp/out")
    app._print_completion_message(0, "clip", is_reel=True)
    aborted = capsys.readouterr().out
    assert "created 1 reel" not in aborted
    assert "No reel was created" in aborted

    app._print_completion_message(1, "clip", is_reel=True)
    assert "created 1 reel" in capsys.readouterr().out


def test_process_clips_parallel_survives_one_failing_clip(monkeypatch, make_clip):
    """A clip that raises is reported and skipped; the batch still completes."""
    clips = [make_clip(row=i, col=2) for i in range(3, 7)]
    monkeypatch.setattr(
        pipeline.files,
        "prepare_clip",
        lambda clip: _prepared_clip(clip, [("00:10", "00:20")]),
    )
    monkeypatch.setattr(pipeline.Path, "is_file", lambda self: True)
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 4)
    monkeypatch.setattr(
        pipeline.files,
        "get_unique_filename",
        lambda template, file_format=None: f"{template}{file_format or ''}",
    )

    import threading

    calls = {"n": 0}
    lock = threading.Lock()

    def run_ffmpeg(*_args, **_kwargs):
        with lock:
            calls["n"] += 1
            first = calls["n"] == 1
        if first:
            raise RuntimeError("boom")
        return True

    monkeypatch.setattr(pipeline.video, "run_ffmpeg", run_ffmpeg)

    count, _artifacts = pipeline.process_clips(clips, output_format="clip")
    assert count == 3


def test_regenerate_gif_artifact_keeps_default_gif_length(monkeypatch, tmp_path):
    """The stored span is the clip's; the GIF stays capped like generation."""
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path), raising=False)
    (tmp_path / "study_P01.mp4").write_text("v")
    extract_gif = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "extract_gif", extract_gif)

    artifact = {
        "type": "gif",
        "file": "clip.gif",
        "sourceVideo": "study_P01.mp4",
        "localStart": 0.0,
        "localEnd": 60.0,
    }
    assert pipeline._regenerate_single_artifact(artifact, set()) is True
    _, kwargs = extract_gif.call_args
    assert kwargs["duration_seconds"] == config.DEFAULT_GIF_DURATION_SECONDS


def test_regenerate_artifact_rounds_fractional_local_times(monkeypatch, tmp_path):
    """Rounding matches the original cut; truncation drifted by up to a second."""
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path), raising=False)
    (tmp_path / "study_P01.mp4").write_text("v")
    run_ffmpeg = Mock(return_value=True)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", run_ffmpeg)

    artifact = {
        "type": "clip",
        "file": "clip.mp4",
        "sourceVideo": "study_P01.mp4",
        "localStart": 10.6,
        "localEnd": 20.4,
    }
    assert pipeline._regenerate_single_artifact(artifact, set()) is True
    _, kwargs = run_ffmpeg.call_args
    assert kwargs["start_pos"] == "0:11"
    assert kwargs["end_pos"] == "0:20"
