from unittest.mock import Mock

import clipgen


def _prepared_clip(raw_clip, times):
    prepared = dict(raw_clip)
    prepared["times"] = list(times)
    return prepared


def test_process_clips_counts_generated_segments_for_all_formats(monkeypatch, make_clip):
    raw_clip = make_clip()
    times = [("00:10", "00:20"), ("00:30", "00:40")]
    monkeypatch.setattr(clipgen.files, "prepare_clip", lambda clip: _prepared_clip(clip, times))
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: True)
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)

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

    assert clipgen.process_clips([raw_clip], output_format="clip") == 2
    assert clipgen.process_clips([raw_clip], output_format="screen") == 2
    assert clipgen.process_clips([raw_clip], output_format="gif") == 2

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

    monkeypatch.setattr(clipgen.utils, "parse_cell_annotations", fake_parse_cell_annotations)
    monkeypatch.setattr(clipgen.utils, "has_non_ignored_timestamp_content", lambda _v: True)

    # 09:15:00-09:16:30 should become 3:00-4:30 relative to 09:12:00 baseline.
    raw_cell_value = "09:15:00-09:16:30"
    raw_clip["cell"].value = raw_cell_value

    prepared = clipgen.files.prepare_clip(raw_clip)
    assert prepared["times"] == [("3:00", "4:30")]


def test_process_clips_skips_when_source_video_missing(monkeypatch, make_clip):
    raw_clip = make_clip()
    monkeypatch.setattr(
        clipgen.files, "prepare_clip", lambda clip: _prepared_clip(clip, [("00:10", "00:20")])
    )
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: False)
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)
    run_ffmpeg = Mock(return_value=True)
    monkeypatch.setattr(clipgen.video, "run_ffmpeg", run_ffmpeg)

    assert clipgen.process_clips([raw_clip], output_format="clip") == 0
    run_ffmpeg.assert_not_called()


def test_process_reel_concatenates_and_cleans_temp_parts(monkeypatch, make_clip):
    raw_clips = [make_clip(row=3, col=2), make_clip(row=4, col=2)]
    monkeypatch.setattr(
        clipgen.files, "prepare_clip", lambda clip: _prepared_clip(clip, [("00:10", "00:20")])
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

    result = clipgen.process_reel(raw_clips, output_file="reel.mp4")
    assert result == 1
    concat.assert_called_once()
    concat_args = concat.call_args.args[0]
    assert concat_args == generated_parts
    assert unlink.call_count == len(generated_parts)


def test_process_reel_returns_zero_when_no_segments_generated(monkeypatch, make_clip):
    raw_clips = [make_clip(row=3, col=2)]
    monkeypatch.setattr(
        clipgen.files, "prepare_clip", lambda clip: _prepared_clip(clip, [("00:10", "00:20")])
    )
    monkeypatch.setattr(clipgen.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(clipgen.files, "get_unique_filename", lambda *_args, **_kwargs: "_reel_part_1.mp4")
    monkeypatch.setattr(clipgen.video, "run_ffmpeg", lambda **_kwargs: False)
    monkeypatch.setattr(clipgen.Path, "is_file", lambda self: True)
    concat = Mock(return_value=True)
    monkeypatch.setattr(clipgen.video, "concatenate_clips", concat)

    result = clipgen.process_reel(raw_clips, output_file="reel.mp4")
    assert result == 0
    concat.assert_not_called()
