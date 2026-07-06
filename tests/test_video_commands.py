import json
import subprocess

import utils
import video

_MATCHING_PROPS = {
    "width": 1920,
    "height": 1080,
    "video_codec": "h264",
    "audio_codec": "aac",
}


def test_build_ffmpeg_cut_command_includes_expected_flags():
    cmd_copy = video.build_ffmpeg_cut_command(
        input_file="in.mp4",
        output_file="out.mp4",
        start_pos="00:10",
        duration_seconds=5,
        reencode=False,
        audio_normalize=True,
    )
    assert "-c:v" in cmd_copy
    assert "copy" in cmd_copy
    assert "-af" in cmd_copy
    assert "loudnorm=I=-16:TP=-1.5:LRA=11" in cmd_copy

    cmd_reencode = video.build_ffmpeg_cut_command(
        input_file="in.mp4",
        output_file="out.mp4",
        start_pos="00:10",
        duration_seconds=5,
        reencode=True,
        audio_normalize=False,
    )
    assert "-c:v" not in cmd_reencode
    assert "-af" not in cmd_reencode
    assert "out.mp4" in cmd_reencode


def test_concatenate_clips_reencode_fallback(monkeypatch):
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "verify_output_file", lambda *_args, **_kwargs: True)
    monkeypatch.setattr(
        video, "probe_video_properties", lambda _path: dict(_MATCHING_PROPS)
    )

    captured_commands = []
    results = [
        subprocess.CompletedProcess(
            args=["ffmpeg"], returncode=1, stderr="copy failed"
        ),
        subprocess.CompletedProcess(args=["ffmpeg"], returncode=0, stderr=""),
    ]

    def fake_run_ffmpeg_process(command, **_kwargs):
        captured_commands.append(command)
        return results.pop(0)

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run_ffmpeg_process)

    ok = video.concatenate_clips(["a.mp4", "b.mp4"], "reel.mp4", reencode_on_fail=True)
    assert ok is True
    assert len(captured_commands) == 2
    assert "-c" in captured_commands[0] and "copy" in captured_commands[0]
    assert "-c:v" in captured_commands[1] and "libx264" in captured_commands[1]
    assert "-c:a" in captured_commands[1] and "aac" in captured_commands[1]


# -- probe_video_properties tests --


def test_probe_video_properties_parses_output(monkeypatch, tmp_path):
    video._video_properties_cache.clear()
    clip = tmp_path / "clip.mp4"
    clip.write_bytes(b"x")
    fake_json = json.dumps(
        {
            "streams": [
                {
                    "codec_type": "video",
                    "codec_name": "h264",
                    "width": 1920,
                    "height": 1080,
                },
                {"codec_type": "audio", "codec_name": "aac"},
            ]
        }
    )
    monkeypatch.setattr(video.subprocess, "check_output", lambda _cmd, **_kw: fake_json)

    result = video.probe_video_properties(str(clip))
    assert result == {
        "width": 1920,
        "height": 1080,
        "video_codec": "h264",
        "audio_codec": "aac",
        "fps": 0.0,
        "duration": 0.0,
        "nb_frames": 0,
    }


def test_probe_video_properties_no_audio(monkeypatch, tmp_path):
    video._video_properties_cache.clear()
    clip = tmp_path / "clip.mp4"
    clip.write_bytes(b"x")
    fake_json = json.dumps(
        {
            "streams": [
                {
                    "codec_type": "video",
                    "codec_name": "hevc",
                    "width": 1280,
                    "height": 720,
                },
            ]
        }
    )
    monkeypatch.setattr(video.subprocess, "check_output", lambda _cmd, **_kw: fake_json)

    result = video.probe_video_properties(str(clip))
    assert result is not None
    assert result["audio_codec"] is None
    assert result["video_codec"] == "hevc"
    assert result["width"] == 1280


def test_probe_video_properties_failure(monkeypatch, tmp_path):
    video._video_properties_cache.clear()
    clip = tmp_path / "clip.mp4"
    clip.write_bytes(b"x")

    def raise_cpe(_cmd, **_kw):
        raise subprocess.CalledProcessError(returncode=1, cmd="ffprobe")

    monkeypatch.setattr(video.subprocess, "check_output", raise_cpe)
    assert video.probe_video_properties(str(clip)) is None


def test_probe_video_properties_file_not_found():
    video._video_properties_cache.clear()
    assert video.probe_video_properties("/nonexistent/missing.mp4") is None


# -- probe_max_keyframe_gap tests --


def test_probe_max_keyframe_gap_uses_max_not_median(monkeypatch, tmp_path):
    video._keyframe_gap_cache.clear()
    clip = tmp_path / "clip.mp4"
    clip.write_bytes(b"x")
    # Keyframes at 0, 1, 2, 8 (gaps 1, 1, 6). The max (6.0) is returned — a
    # single long-GOP stretch must not be masked by the surrounding short gaps
    # (median would be 1.0). Interior non-keyframe packets are ignored.
    csv = "\n".join(
        [
            "0.000000,K__",
            "0.033000,__",
            "1.000000,K__",
            "2.000000,K__",
            "8.000000,K__",
        ]
    )
    monkeypatch.setattr(video.subprocess, "check_output", lambda _cmd, **_kw: csv)
    assert video.probe_max_keyframe_gap(str(clip)) == 6.0


def test_probe_max_keyframe_gap_failure_returns_none(monkeypatch, tmp_path):
    video._keyframe_gap_cache.clear()
    clip = tmp_path / "clip.mp4"
    clip.write_bytes(b"x")

    def raise_cpe(_cmd, **_kw):
        raise subprocess.CalledProcessError(returncode=1, cmd="ffprobe")

    monkeypatch.setattr(video.subprocess, "check_output", raise_cpe)
    assert video.probe_max_keyframe_gap(str(clip)) is None


def test_probe_max_keyframe_gap_single_keyframe_returns_none(monkeypatch, tmp_path):
    video._keyframe_gap_cache.clear()
    clip = tmp_path / "clip.mp4"
    clip.write_bytes(b"x")
    # Only one keyframe in the probe window → GOP longer than the window; can't
    # confirm short cadence, so treat as unknown (None → do not skip).
    csv = "\n".join(["0.000000,K__", "0.033000,__", "1.000000,__"])
    monkeypatch.setattr(video.subprocess, "check_output", lambda _cmd, **_kw: csv)
    assert video.probe_max_keyframe_gap(str(clip)) is None


def test_probe_max_keyframe_gap_file_not_found():
    video._keyframe_gap_cache.clear()
    assert video.probe_max_keyframe_gap("/nonexistent/missing.mp4") is None


def test_probe_video_properties_reprobes_after_mtime_change(monkeypatch, tmp_path):
    """A re-encoded file (new mtime_ns) invalidates the cached props."""
    import os

    video._video_properties_cache.clear()
    clip = tmp_path / "clip.mp4"
    clip.write_bytes(b"original")

    fake_streams = {
        "streams": [
            {
                "codec_type": "video",
                "codec_name": "h264",
                "width": 1920,
                "height": 1080,
                "r_frame_rate": "30/1",
            },
        ],
        "format": {},
    }
    call_count = {"n": 0}

    def fake_check_output(_cmd, **_kw):
        call_count["n"] += 1
        # Second probe sees a different width to prove we re-probed.
        if call_count["n"] >= 2:
            fake_streams["streams"][0]["width"] = 1280
        return json.dumps(fake_streams)

    monkeypatch.setattr(video.subprocess, "check_output", fake_check_output)

    first = video.probe_video_properties(str(clip))
    assert first is not None and first["width"] == 1920
    cached = video.probe_video_properties(str(clip))
    assert cached is first  # same dict object — cache hit
    assert call_count["n"] == 1

    # Bump mtime forward to simulate a re-encode in place.
    bumped = clip.stat().st_mtime_ns + 10**9
    os.utime(clip, ns=(bumped, bumped))

    fresh = video.probe_video_properties(str(clip))
    assert fresh is not None and fresh["width"] == 1280
    assert call_count["n"] == 2


# -- concatenate mismatch detection tests --


def test_concatenate_clips_matching_uses_demuxer(monkeypatch):
    """Identical properties → concat demuxer path (fast, stream copy)."""
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)
    monkeypatch.setattr(
        video, "probe_video_properties", lambda _p: dict(_MATCHING_PROPS)
    )

    captured = []

    def fake_run(command, **_kw):
        captured.append(command)
        return subprocess.CompletedProcess(args=command, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    ok = video.concatenate_clips(["a.mp4", "b.mp4"], "reel.mp4")
    assert ok is True
    assert len(captured) == 1
    assert "-f" in captured[0] and "concat" in captured[0]
    assert "-c" in captured[0] and "copy" in captured[0]
    assert "-filter_complex" not in captured[0]


def test_concatenate_clips_resolution_mismatch_uses_filter_complex(monkeypatch):
    """Different resolutions → filter_complex path with scale+pad."""
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)

    props_by_path = {
        "a.mp4": {
            "width": 1920,
            "height": 1080,
            "video_codec": "h264",
            "audio_codec": "aac",
        },
        "b.mp4": {
            "width": 1280,
            "height": 720,
            "video_codec": "h264",
            "audio_codec": "aac",
        },
    }
    monkeypatch.setattr(video, "probe_video_properties", lambda p: props_by_path.get(p))

    captured = []

    def fake_run(command, **_kw):
        captured.append(command)
        return subprocess.CompletedProcess(args=command, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    ok = video.concatenate_clips(["a.mp4", "b.mp4"], "reel.mp4")
    assert ok is True
    assert len(captured) == 1
    cmd = captured[0]
    assert "-filter_complex" in cmd
    fc_idx = cmd.index("-filter_complex")
    fc_str = cmd[fc_idx + 1]
    assert "scale=1920:1080" in fc_str
    assert "pad=1920:1080" in fc_str
    assert "concat=n=2:v=1:a=1" in fc_str
    assert "-f" not in cmd  # no concat demuxer


def test_concatenate_clips_warns_on_mismatch(monkeypatch):
    """Mismatch detection prints warnings."""
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)

    props_by_path = {
        "a.mp4": {
            "width": 1920,
            "height": 1080,
            "video_codec": "h264",
            "audio_codec": "aac",
        },
        "b.mp4": {
            "width": 1280,
            "height": 720,
            "video_codec": "hevc",
            "audio_codec": "aac",
        },
    }
    monkeypatch.setattr(video, "probe_video_properties", lambda p: props_by_path.get(p))
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda cmd, **_kw: subprocess.CompletedProcess(
            args=cmd, returncode=0, stderr=""
        ),
    )

    warnings = []
    original_warning = video.utils.warning_print

    def capture_warning(msg, details=None):
        warnings.append(msg)
        original_warning(msg, details)

    monkeypatch.setattr(video.utils, "warning_print", capture_warning)

    video.concatenate_clips(["a.mp4", "b.mp4"], "reel.mp4")
    warning_text = " ".join(warnings)
    assert "Resolution mismatch" in warning_text
    assert "1920x1080" in warning_text
    assert "1280x720" in warning_text
    assert "Video codec mismatch" in warning_text


def test_concatenate_clips_mixed_audio_presence(monkeypatch):
    """One clip with audio, one without → anullsrc for the silent clip."""
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)
    monkeypatch.setattr(video, "get_file_duration", lambda _p: 10)

    props_by_path = {
        "a.mp4": {
            "width": 1920,
            "height": 1080,
            "video_codec": "h264",
            "audio_codec": "aac",
        },
        "b.mp4": {
            "width": 1920,
            "height": 1080,
            "video_codec": "h264",
            "audio_codec": None,
        },
    }
    monkeypatch.setattr(video, "probe_video_properties", lambda p: props_by_path.get(p))

    captured = []

    def fake_run(command, **_kw):
        captured.append(command)
        return subprocess.CompletedProcess(args=command, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    ok = video.concatenate_clips(["a.mp4", "b.mp4"], "reel.mp4")
    assert ok is True
    cmd = captured[0]
    assert "-filter_complex" in cmd
    fc_str = cmd[cmd.index("-filter_complex") + 1]
    assert "anullsrc" in fc_str
    assert "concat=n=2:v=1:a=1" in fc_str


def test_concatenate_clips_all_no_audio(monkeypatch):
    """All clips lack audio → concat with a=0, no audio mapping."""
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)

    no_audio_props = {
        "width": 1280,
        "height": 720,
        "video_codec": "h264",
        "audio_codec": None,
    }
    monkeypatch.setattr(
        video, "probe_video_properties", lambda _p: dict(no_audio_props)
    )

    # Same resolution, but audio mismatch is about presence—need at least one
    # mismatch to trigger filter_complex. Use resolution mismatch instead.
    call_count = [0]

    def props_alternating(_p):
        call_count[0] += 1
        if call_count[0] % 2 == 1:
            return {
                "width": 1920,
                "height": 1080,
                "video_codec": "h264",
                "audio_codec": None,
            }
        return {
            "width": 1280,
            "height": 720,
            "video_codec": "h264",
            "audio_codec": None,
        }

    monkeypatch.setattr(video, "probe_video_properties", props_alternating)

    captured = []

    def fake_run(command, **_kw):
        captured.append(command)
        return subprocess.CompletedProcess(args=command, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    ok = video.concatenate_clips(["a.mp4", "b.mp4"], "reel.mp4")
    assert ok is True
    cmd = captured[0]
    fc_str = cmd[cmd.index("-filter_complex") + 1]
    assert "concat=n=2:v=1:a=0" in fc_str
    assert "[outa]" not in fc_str
    assert '"-map", "[outa]"' not in str(cmd)


def test_get_duration_valid_and_invalid_values(monkeypatch):
    # Valid MM:SS
    assert video.get_duration("00:10", "00:20") == 10
    # Valid HH:MM:SS
    assert video.get_duration("01:00:00", "01:10:00") == 600
    # INVALID_END_TIMESTAMP sentinel returns None
    assert video.get_duration("00:10", video.INVALID_END_TIMESTAMP) is None
    # Completely invalid format returns None
    assert video.get_duration("not-a-time", "also-bad") is None


def test_get_duration_mixed_format_and_overflow_minutes():
    """Pairs whose ends need different formats, and MM:SS with minutes >= 60,
    must yield a real duration. The timestamp parser emits both shapes; a
    shared-format strptime previously returned None and silently dropped the
    clip in run_ffmpeg."""
    # Mixed M:SS / H:MM:SS pair (a clip straddling the hour boundary).
    assert video.get_duration("59:50", "1:00:10") == 20
    # MM:SS with minutes >= 60 (a long session written without an hours field).
    assert video.get_duration("75:00", "80:00") == 300
    # Both ends carrying hours still work.
    assert video.get_duration("1:00:10", "1:00:30") == 20


def test_get_duration_accepts_every_pair_the_parser_emits():
    """Tie the parser contract to the cutter: any (start, end) pair that
    parse_timestamps produces must yield a duration, not None — including
    single timestamps whose default-duration end crosses the hour."""
    for cell in ("59:50", "75:00", "58:30", "0:10"):
        start, end = utils.parse_timestamps(cell)[0]
        assert video.get_duration(start, end) is not None, cell
    # Explicit range cell with overflow minutes.
    start, end = utils.parse_timestamps("75:00-80:00")[0]
    assert video.get_duration(start, end) == 300


def test_calculate_target_bitrate_typical_and_min_floor():
    kbps = video.calculate_target_bitrate(target_size_mb=50, duration_seconds=600)
    assert kbps > 100
    small = video.calculate_target_bitrate(target_size_mb=1, duration_seconds=5)
    assert small >= 100
    zero_duration = video.calculate_target_bitrate(
        target_size_mb=10, duration_seconds=0
    )
    assert zero_duration == 100


def test_get_file_duration_returns_rounded_probe_duration(monkeypatch, tmp_path):
    """After probe_video_properties, duration must not depend on a prior cache hit."""
    video_f = tmp_path / "video.mp4"
    video_f.write_bytes(b"x")
    video._file_duration_cache.clear()
    video._video_properties_cache.clear()

    def fake_probe(_path: str) -> dict:
        return {
            "width": 1920,
            "height": 1080,
            "video_codec": "h264",
            "audio_codec": None,
            "fps": 30.0,
            "duration": 99.4,
            "nb_frames": 0,
        }

    monkeypatch.setattr(video, "probe_video_properties", fake_probe)
    assert video.get_file_duration(str(video_f)) == 99
    key = (str(video_f.resolve()), video_f.stat().st_mtime_ns)
    assert video._file_duration_cache[key] == 99


def test_get_file_duration_error_paths(monkeypatch):
    # Missing file
    monkeypatch.setattr(video.os.path, "isfile", lambda _path: False)
    assert video.get_file_duration("missing.mp4") is None

    # ffprobe not found
    monkeypatch.setattr(video.os.path, "isfile", lambda _path: True)

    def raise_fnf(_cmd, **_kwargs):
        raise FileNotFoundError()

    monkeypatch.setattr(video.subprocess, "check_output", raise_fnf)
    assert video.get_file_duration("video.mp4") is None

    # ffprobe CalledProcessError
    def raise_cpe(_cmd, **_kwargs):
        raise video.subprocess.CalledProcessError(returncode=1, cmd="ffprobe")

    monkeypatch.setattr(video.subprocess, "check_output", raise_cpe)
    assert video.get_file_duration("video.mp4") is None


def test_verify_output_file_rejects_empty_placeholder(tmp_path):
    """A zero-byte reservation placeholder (ffmpeg exited 0 but wrote nothing)
    must fail verification so the caller releases it instead of registering a
    bogus zero-byte artifact."""
    empty = tmp_path / "out.mp4"
    empty.write_bytes(b"")  # what get_unique_filename leaves on disk
    assert video.verify_output_file(str(empty), "ffmpeg") is False

    missing = tmp_path / "nope.mp4"
    assert video.verify_output_file(str(missing), "ffmpeg") is False

    real = tmp_path / "real.mp4"
    real.write_bytes(b"\x00\x01\x02")
    assert video.verify_output_file(str(real), "ffmpeg") is True


def test_check_ffmpeg_tools_available_missing(monkeypatch):
    monkeypatch.setattr(video.shutil, "which", lambda _tool: None)
    ok = video.check_ffmpeg_tools_available()
    assert ok is False


_FFMPEG_ENCODERS_WITH_WEBP = (
    "Encoders:\n"
    " V..... = Video\n"
    " ------\n"
    " V..... libwebp_anim    libwebp WebP image (codec webp)\n"
    " V..... libwebp         libwebp WebP image (codec webp)\n"
)

_FFMPEG_ENCODERS_NO_WEBP = (
    "Encoders:\n V..... = Video\n ------\n V..... libx264         H.264 (codec h264)\n"
)

# The codecs listing mentions 'webp' even on builds without libwebp; the old
# check matched this and produced a false positive. Used here as a regression
# guard.
_FFMPEG_CODECS_WITHOUT_LIBWEBP_ENCODER = (
    "Codecs:\n D..... = Decoder\n .EV..L webp     WebP (encoders: )\n"
)


def test_check_webp_support_detects_libwebp(monkeypatch):
    monkeypatch.setattr(video, "_webp_support_cache", None)
    monkeypatch.setattr(
        video.subprocess,
        "run",
        lambda *_a, **_kw: subprocess.CompletedProcess(
            [], 0, stdout=_FFMPEG_ENCODERS_WITH_WEBP, stderr=""
        ),
    )
    assert video.check_webp_support() is True

    monkeypatch.setattr(video, "_webp_support_cache", None)
    monkeypatch.setattr(
        video.subprocess,
        "run",
        lambda *_a, **_kw: subprocess.CompletedProcess(
            [], 0, stdout=_FFMPEG_ENCODERS_NO_WEBP, stderr=""
        ),
    )
    assert video.check_webp_support() is False


def test_check_webp_support_ignores_codecs_listing(monkeypatch):
    """Regression: the codecs listing is not authoritative; only -encoders is."""
    monkeypatch.setattr(video, "_webp_support_cache", None)
    monkeypatch.setattr(
        video.subprocess,
        "run",
        lambda *_a, **_kw: subprocess.CompletedProcess(
            [], 0, stdout=_FFMPEG_CODECS_WITHOUT_LIBWEBP_ENCODER, stderr=""
        ),
    )
    assert video.check_webp_support() is False


_FFMPEG_FILTERS_WITH_DRAWTEXT = (
    "Filters:\n"
    "  T.. = Timeline support\n"
    " ... drawbox           V->V       Draw a colored box.\n"
    " T.. drawtext          V->V       Draw text on top of video.\n"
)

_FFMPEG_FILTERS_NO_DRAWTEXT = (
    "Filters:\n"
    "  T.. = Timeline support\n"
    " ... drawbox           V->V       Draw a colored box.\n"
)

_FFMPEG_ENCODERS_WITH_VP9 = (
    "Encoders:\n"
    " V..... = Video\n"
    " ------\n"
    " V..... libvpx-vp9      libvpx VP9 (codec vp9)\n"
)

_FFMPEG_ENCODERS_NO_VP9 = (
    "Encoders:\n V..... = Video\n ------\n V..... libx264         H.264 (codec h264)\n"
)


def test_check_drawtext_support_detects_filter(monkeypatch):
    monkeypatch.setattr(video, "_drawtext_support_cache", None)
    monkeypatch.setattr(
        video.subprocess,
        "run",
        lambda *_a, **_kw: subprocess.CompletedProcess(
            [], 0, stdout=_FFMPEG_FILTERS_WITH_DRAWTEXT, stderr=""
        ),
    )
    assert video.check_drawtext_support() is True

    monkeypatch.setattr(video, "_drawtext_support_cache", None)
    monkeypatch.setattr(
        video.subprocess,
        "run",
        lambda *_a, **_kw: subprocess.CompletedProcess(
            [], 0, stdout=_FFMPEG_FILTERS_NO_DRAWTEXT, stderr=""
        ),
    )
    assert video.check_drawtext_support() is False


def test_check_vp9_support_detects_encoder(monkeypatch):
    monkeypatch.setattr(video, "_vp9_support_cache", None)
    monkeypatch.setattr(
        video.subprocess,
        "run",
        lambda *_a, **_kw: subprocess.CompletedProcess(
            [], 0, stdout=_FFMPEG_ENCODERS_WITH_VP9, stderr=""
        ),
    )
    assert video.check_vp9_support() is True

    monkeypatch.setattr(video, "_vp9_support_cache", None)
    monkeypatch.setattr(
        video.subprocess,
        "run",
        lambda *_a, **_kw: subprocess.CompletedProcess(
            [], 0, stdout=_FFMPEG_ENCODERS_NO_VP9, stderr=""
        ),
    )
    assert video.check_vp9_support() is False


def test_extract_gif_rejects_webm_when_vp9_unsupported(monkeypatch):
    monkeypatch.setattr(video, "check_vp9_support", lambda: False)
    monkeypatch.setattr(video, "_vp9_missing_warned", False)

    called = {"run": False}

    def fake_run(*_a, **_kw):
        called["run"] = True
        return subprocess.CompletedProcess([], 0, stdout="", stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    ok = video.extract_gif("/in.mp4", "/out.webm", "0:10", 3)
    assert ok is False
    assert called["run"] is False


def test_extract_gif_includes_webp_quality_for_webp_output(monkeypatch):
    import config as cfg

    monkeypatch.setattr(cfg, "DEBUGGING", False)
    monkeypatch.setattr(video, "get_file_duration", lambda _f: 600)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video.Path, "stat", lambda self: type("S", (), {"st_size": 1})()
    )
    monkeypatch.setattr(video, "check_webp_support", lambda: True)

    captured: dict = {}

    def fake_run(cmd, **_kw):
        captured["cmd"] = cmd
        return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    ok = video.extract_gif("/in.mp4", "/out.webp", "0:10", 3)
    assert ok is True
    assert "-quality" in captured["cmd"]
    assert str(cfg.WEBP_QUALITY) in captured["cmd"]

    captured.clear()
    ok = video.extract_gif("/in.mp4", "/out.gif", "0:10", 3)
    assert ok is True
    assert "-quality" not in captured["cmd"]


def test_extract_gif_uses_vp9_for_webm_output(monkeypatch):
    import config as cfg

    monkeypatch.setattr(cfg, "DEBUGGING", False)
    monkeypatch.setattr(video, "get_file_duration", lambda _f: 600)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video.Path, "stat", lambda self: type("S", (), {"st_size": 1})()
    )
    monkeypatch.setattr(video, "check_vp9_support", lambda: True)

    captured: dict = {}

    def fake_run(cmd, **_kw):
        captured["cmd"] = cmd
        return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    ok = video.extract_gif("/in.mp4", "/out.webm", "0:10", 3)
    assert ok is True
    cmd = captured["cmd"]
    assert "libvpx-vp9" in cmd
    assert "-an" in cmd
    # WebM container does not honor -loop; that flag must not be added.
    assert "-loop" not in cmd
    # WebP-only quality flag must not leak into the WebM command either.
    assert "-quality" not in cmd


def test_extract_gif_rejects_webp_when_unsupported(monkeypatch):
    monkeypatch.setattr(video, "check_webp_support", lambda: False)
    monkeypatch.setattr(video, "_webp_missing_warned", False)

    called = {"run": False}

    def fake_run(*_a, **_kw):
        called["run"] = True
        return subprocess.CompletedProcess([], 0, stdout="", stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    ok = video.extract_gif("/in.mp4", "/out.webp", "0:10", 3)
    assert ok is False
    assert called["run"] is False


def _patch_batch_screenshots(monkeypatch, captured: dict, ext: str) -> None:
    import config as cfg

    monkeypatch.setattr(cfg, "SCREENSHOT_FORMAT", ext)
    monkeypatch.setattr(video, "check_webp_support", lambda: True)

    def fake_run(cmd, **_kw):
        captured["cmd"] = cmd
        return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)
    # The artifact-collection loop needs the per-frame files to "exist" so the
    # function returns successfully; pretend they all do.
    monkeypatch.setattr(video.os.path, "isfile", lambda _p: True)
    monkeypatch.setattr(video.shutil, "move", lambda _src, _dst: None)
    monkeypatch.setattr(
        video.files, "get_unique_filename", lambda name, file_format: name
    )


def test_batch_extract_screenshots_forces_image2_muxer_for_webp(monkeypatch):
    """Regression: a single ffmpeg pass with `%04d.webp` could auto-select the
    webp (animation) muxer and bundle every frame into one animated file. The
    fix forces `-f image2` and `-c:v libwebp` so each interval produces a
    separate still WebP."""
    import config as cfg

    captured: dict = {}
    _patch_batch_screenshots(monkeypatch, captured, ".webp")

    artifacts = video._batch_extract_screenshots("/in.mp4", [0, 10, 20], 10)
    assert artifacts is not None
    cmd = captured["cmd"]
    assert "-f" in cmd and "image2" in cmd
    f_idx = cmd.index("-f")
    assert cmd[f_idx + 1] == "image2"
    assert "-c:v" in cmd
    cv_idx = cmd.index("-c:v")
    assert cmd[cv_idx + 1] == "libwebp"
    assert "-quality" in cmd
    assert str(cfg.WEBP_QUALITY) in cmd
    assert cmd[-1].endswith("frame_%04d.webp")


def test_batch_extract_screenshots_png_uses_image2_without_libwebp(monkeypatch):
    """`-f image2` is the right muxer for any `%d`-pattern still output, so it
    is added unconditionally. `-c:v libwebp` / `-quality` must stay scoped to
    .webp output."""
    captured: dict = {}
    _patch_batch_screenshots(monkeypatch, captured, ".png")

    artifacts = video._batch_extract_screenshots("/in.mp4", [0, 10, 20], 10)
    assert artifacts is not None
    cmd = captured["cmd"]
    assert "-f" in cmd and "image2" in cmd
    assert "libwebp" not in cmd
    assert "-quality" not in cmd
    assert cmd[-1].endswith("frame_%04d.png")


# ---- extract_frame_at_timestamp ----


def test_extract_frame_at_timestamp_returns_frame(monkeypatch):
    """Successful extraction returns a numpy array with correct shape."""
    import numpy as np

    w, h = 320, 240
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _: {"width": w, "height": h, "fps": 30.0, "duration": 10.0},
    )
    fake_frame = np.zeros(h * w * 3, dtype=np.uint8)
    monkeypatch.setattr(
        video.subprocess,
        "run",
        lambda *a, **kw: subprocess.CompletedProcess(a, 0, stdout=fake_frame.tobytes()),
    )

    frame = video.extract_frame_at_timestamp("/fake.mp4", 1.5)
    assert frame is not None
    assert frame.shape == (h, w, 3)


def test_extract_frame_at_timestamp_returns_none_on_probe_failure(monkeypatch):
    """Returns None when probe_video_properties fails."""
    monkeypatch.setattr(video, "probe_video_properties", lambda _: None)
    assert video.extract_frame_at_timestamp("/fake.mp4", 0.0) is None


def test_extract_frame_at_timestamp_returns_none_on_short_output(monkeypatch):
    """Returns None when ffmpeg outputs fewer bytes than expected."""
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _: {"width": 320, "height": 240, "fps": 30.0, "duration": 10.0},
    )
    monkeypatch.setattr(
        video.subprocess,
        "run",
        lambda *a, **kw: subprocess.CompletedProcess(a, 0, stdout=b"short"),
    )
    assert video.extract_frame_at_timestamp("/fake.mp4", 0.0) is None


def test_extract_frame_at_timestamp_debug_mode(monkeypatch):
    """In debug mode, returns a stub frame without calling ffmpeg."""
    monkeypatch.setattr(video.config, "DEBUGGING", True)
    frame = video.extract_frame_at_timestamp("/fake.mp4", 0.0)
    assert frame is not None
    assert frame.shape == (1080, 1920, 3)


# ---- accurate_seek_args ----


def test_accurate_seek_args_zero_returns_empty_lists():
    pre, post = video.accurate_seek_args(0.0)
    assert pre == []
    assert post == []


def test_accurate_seek_args_within_preseek_window_skips_pre():
    """For ts within the preseek window, the entire seek goes after -i."""
    ts = video.FFMPEG_PRESEEK_SECONDS / 2
    pre, post = video.accurate_seek_args(ts)
    assert pre == []
    assert post == ["-ss", str(ts)]


def test_accurate_seek_args_splits_for_far_target():
    """For ts beyond the preseek window, we get a fast pre + small post."""
    ts = video.FFMPEG_PRESEEK_SECONDS + 12.345
    pre, post = video.accurate_seek_args(ts)
    assert pre == ["-ss", str(ts - video.FFMPEG_PRESEEK_SECONDS)]
    assert post == ["-ss", str(video.FFMPEG_PRESEEK_SECONDS)]


# ---- two-stage seek wiring in extract_frame_at_timestamp ----


def _captured_run(captured: dict):
    def _run(*args, **kwargs):
        captured["cmd"] = list(args[0])
        return subprocess.CompletedProcess(args, 0, stdout=b"")

    return _run


def test_extract_frame_at_timestamp_uses_two_stage_seek(monkeypatch):
    """Far targets get -ss before -i AND -ss after -i (frame-accurate)."""
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _: {"width": 320, "height": 240, "fps": 30.0, "duration": 60.0},
    )
    captured: dict = {}
    monkeypatch.setattr(video.subprocess, "run", _captured_run(captured))

    video.extract_frame_at_timestamp("/fake.mp4", 12.5)
    cmd = captured["cmd"]
    i_idx = cmd.index("-i")
    pre = cmd[:i_idx]
    post = cmd[i_idx + 2 :]
    assert "-ss" in pre, "expected fast pre-input seek"
    assert "-ss" in post, "expected accurate post-input seek"


def test_extract_frame_at_timestamp_post_only_seek_for_near_target(monkeypatch):
    """For ts inside the preseek window, only post-input -ss is emitted."""
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _: {"width": 320, "height": 240, "fps": 30.0, "duration": 60.0},
    )
    captured: dict = {}
    monkeypatch.setattr(video.subprocess, "run", _captured_run(captured))

    video.extract_frame_at_timestamp("/fake.mp4", 0.5)
    cmd = captured["cmd"]
    i_idx = cmd.index("-i")
    assert "-ss" not in cmd[:i_idx]
    assert "-ss" in cmd[i_idx + 2 :]


# ---- two-stage seek + float ts in extract_thumbnail_bytes ----


def test_extract_thumbnail_bytes_preserves_float_timestamp(monkeypatch, tmp_path):
    """The float timestamp must reach ffmpeg without int() truncation."""
    fake = tmp_path / "video.mp4"
    fake.write_bytes(b"x")
    captured: dict = {}
    monkeypatch.setattr(video.subprocess, "run", _captured_run(captured))

    video.extract_thumbnail_bytes(str(fake), 12.75, width=200)
    cmd = captured["cmd"]
    seek_values = [cmd[i + 1] for i, a in enumerate(cmd) if a == "-ss"]
    assert seek_values, "expected at least one -ss flag"
    assert any("12.75" in str(v) or "10.75" in str(v) for v in seek_values), (
        f"float seconds should appear in either pre- or post-seek (got {seek_values})"
    )
    # And no integer-truncated whole-second-only command.
    i_idx = cmd.index("-i")
    assert "-ss" in cmd[i_idx + 2 :]


# ---- card-scrubber media helpers (sprite sheet + audio segment) ----


def test_extract_sprite_sheet_bytes_debug_mode(monkeypatch):
    monkeypatch.setattr(video.config, "DEBUGGING", True)
    assert video.extract_sprite_sheet_bytes("x.mp4", 0.0, 5.0, 5, 5) is None


def test_extract_sprite_sheet_bytes_missing_file(monkeypatch):
    monkeypatch.setattr(video.config, "DEBUGGING", False)
    assert video.extract_sprite_sheet_bytes("/nope.mp4", 0.0, 5.0, 5, 5) is None


def test_extract_sprite_sheet_bytes_builds_tile_command(monkeypatch, tmp_path):
    monkeypatch.setattr(video.config, "DEBUGGING", False)
    fake = tmp_path / "v.mp4"
    fake.write_bytes(b"x")
    captured: dict = {}
    monkeypatch.setattr(video.subprocess, "run", _captured_run(captured))

    video.extract_sprite_sheet_bytes(str(fake), 2.0, 5.0, 4, 3)
    cmd = captured["cmd"]
    vf = cmd[cmd.index("-vf") + 1]
    assert "tile=4x3" in vf  # cols x rows grid
    assert "fps=12/5.0" in vf  # cols*rows frames over the duration


def test_extract_audio_segment_bytes_debug_mode(monkeypatch):
    monkeypatch.setattr(video.config, "DEBUGGING", True)
    assert video.extract_audio_segment_bytes("x.mp4", 0.0, 5.0) is None


def test_extract_audio_segment_bytes_missing_file(monkeypatch):
    monkeypatch.setattr(video.config, "DEBUGGING", False)
    assert video.extract_audio_segment_bytes("/nope.mp4", 0.0, 5.0) is None


def test_extract_audio_segment_bytes_builds_wav_command(monkeypatch, tmp_path):
    monkeypatch.setattr(video.config, "DEBUGGING", False)
    fake = tmp_path / "v.mp4"
    fake.write_bytes(b"x")
    captured: dict = {}
    monkeypatch.setattr(video.subprocess, "run", _captured_run(captured))

    video.extract_audio_segment_bytes(str(fake), 1.0, 4.0, sample_rate=16000)
    cmd = captured["cmd"]
    assert "-vn" in cmd
    assert cmd[cmd.index("-ac") + 1] == "1"
    assert cmd[cmd.index("-ar") + 1] == "16000"
    assert "pcm_s16le" in cmd
    assert cmd[cmd.index("-f") + 1] == "wav"


# -- cancel_flag forwarding --


def _captured_run_ffmpeg(captured):
    def _fake(_cmd, **kwargs):
        captured.append(kwargs.get("cancel_flag"))
        return subprocess.CompletedProcess(args=["ffmpeg"], returncode=0, stderr="")

    return _fake


def test_run_ffmpeg_forwards_cancel_flag(monkeypatch):
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video.Path, "stat", lambda self: type("_S", (), {"st_size": 1})()
    )
    monkeypatch.setattr(video, "get_duration", lambda *_a: 5)
    monkeypatch.setattr(video, "get_file_duration", lambda *_a: 60)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)
    # Even with a size cap set, run_ffmpeg is a pure cut and must NOT compress —
    # enforcement moved to the callers, applied after any titlecard wrap/concat.
    monkeypatch.setattr(video.config, "MAX_FILESIZE_MB", 50)

    def _no_compress(*_a, **_kw):
        raise AssertionError("run_ffmpeg must not call compress_to_size")

    monkeypatch.setattr(video, "compress_to_size", _no_compress)

    captured: list = []
    monkeypatch.setattr(video, "run_ffmpeg_process", _captured_run_ffmpeg(captured))

    sentinel = lambda: False  # noqa: E731
    ok = video.run_ffmpeg(
        "in.mp4", "out.mp4", "00:10", "00:15", reencode=False, cancel_flag=sentinel
    )
    assert ok is True
    assert captured == [sentinel]


def test_enforce_filesize_limit_noop_when_disabled(monkeypatch):
    monkeypatch.setattr(video.config, "MAX_FILESIZE_MB", 0)

    def _no_compress(*_a, **_kw):
        raise AssertionError("compress_to_size must not run when the cap is disabled")

    monkeypatch.setattr(video, "compress_to_size", _no_compress)
    # Should return without touching compress_to_size.
    video.enforce_filesize_limit("out.mp4")


def test_enforce_filesize_limit_compresses_and_forwards_cancel_flag(monkeypatch):
    monkeypatch.setattr(video.config, "MAX_FILESIZE_MB", 50)

    captured: list = []

    def _fake_compress(path, target_mb, *, cancel_flag=None):
        captured.append((path, target_mb, cancel_flag))
        return True

    monkeypatch.setattr(video, "compress_to_size", _fake_compress)

    sentinel = lambda: False  # noqa: E731
    video.enforce_filesize_limit("out.mp4", cancel_flag=sentinel)
    assert captured == [("out.mp4", 50, sentinel)]


def test_extract_screenshot_forwards_cancel_flag(monkeypatch):
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "get_file_duration", lambda *_a: 60)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)
    monkeypatch.setattr(
        video.Path, "stat", lambda self: type("_S", (), {"st_size": 1})()
    )

    captured: list = []
    monkeypatch.setattr(video, "run_ffmpeg_process", _captured_run_ffmpeg(captured))

    sentinel = lambda: False  # noqa: E731
    ok = video.extract_screenshot("in.mp4", "out.png", "00:10", cancel_flag=sentinel)
    assert ok is True
    assert captured == [sentinel]


def test_extract_gif_forwards_cancel_flag(monkeypatch):
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "get_file_duration", lambda *_a: 60)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)
    monkeypatch.setattr(
        video.Path, "stat", lambda self: type("_S", (), {"st_size": 1})()
    )

    captured: list = []
    monkeypatch.setattr(video, "run_ffmpeg_process", _captured_run_ffmpeg(captured))

    sentinel = lambda: False  # noqa: E731
    ok = video.extract_gif("in.mp4", "out.gif", "00:10", 5, cancel_flag=sentinel)
    assert ok is True
    assert captured == [sentinel]


def test_compress_to_size_forwards_cancel_flag_to_both_passes(monkeypatch, tmp_path):
    big = tmp_path / "big.mp4"
    big.write_bytes(b"x" * 1024)
    monkeypatch.setattr(video, "get_file_duration", lambda *_a: 10)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)

    captured: list = []
    monkeypatch.setattr(video, "run_ffmpeg_process", _captured_run_ffmpeg(captured))

    # Stub os.replace and unlink so the test doesn't move/delete real files.
    monkeypatch.setattr(video.os, "replace", lambda *_a, **_kw: None)

    sentinel = lambda: False  # noqa: E731
    # Target tiny size so compression runs both passes.
    video.compress_to_size(str(big), 0.0001, cancel_flag=sentinel)
    assert captured == [sentinel, sentinel]


# -- mux_subtitles tests --


def _write_dummy_pair(tmp_path, video_name="in.mp4", srt_name="in.srt"):
    """Create a stand-in video + SRT file so existence checks pass."""
    src = tmp_path / video_name
    src.write_bytes(b"\x00")
    srt = tmp_path / srt_name
    srt.write_text("1\n00:00:00,000 --> 00:00:01,000\nhi\n", encoding="utf-8")
    return src, srt


def test_mux_subtitles_mp4_uses_mov_text(monkeypatch, tmp_path):
    src, srt = _write_dummy_pair(tmp_path)
    out = tmp_path / "out.mp4"

    captured = {}

    def fake_run(command, **_kwargs):
        captured["command"] = command
        return subprocess.CompletedProcess(args=command, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)

    ok = video.mux_subtitles(str(src), str(srt), str(out))
    assert ok is True
    cmd = captured["command"]
    assert cmd[:2] == ["ffmpeg", "-y"]
    assert "-c:s" in cmd and cmd[cmd.index("-c:s") + 1] == "mov_text"
    assert "-c" in cmd and cmd[cmd.index("-c") + 1] == "copy"
    assert "-map" in cmd
    assert "0" in cmd and "1:0" in cmd
    assert cmd[-1] == str(out)


def test_mux_subtitles_mkv_uses_srt(monkeypatch, tmp_path):
    src, srt = _write_dummy_pair(tmp_path, video_name="in.mkv")
    out = tmp_path / "out.mkv"

    captured = {}

    def fake_run(command, **_kwargs):
        captured["command"] = command
        return subprocess.CompletedProcess(args=command, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)

    ok = video.mux_subtitles(str(src), str(srt), str(out))
    assert ok is True
    cmd = captured["command"]
    assert cmd[cmd.index("-c:s") + 1] == "srt"


def test_mux_subtitles_unsupported_container_returns_false(monkeypatch, tmp_path):
    src, srt = _write_dummy_pair(tmp_path)
    out = tmp_path / "out.avi"

    invoked = []
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda *a, **kw: invoked.append(True) or None,
    )

    ok = video.mux_subtitles(str(src), str(srt), str(out))
    assert ok is False
    assert invoked == []


def test_mux_subtitles_returns_false_on_ffmpeg_error(monkeypatch, tmp_path):
    src, srt = _write_dummy_pair(tmp_path)
    out = tmp_path / "out.mp4"

    def fake_run(command, **_kwargs):
        return subprocess.CompletedProcess(args=command, returncode=1, stderr="boom")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    ok = video.mux_subtitles(str(src), str(srt), str(out))
    assert ok is False


def test_mux_subtitles_missing_input_returns_false(tmp_path):
    srt = tmp_path / "in.srt"
    srt.write_text("1\n00:00:00,000 --> 00:00:01,000\nhi\n", encoding="utf-8")
    ok = video.mux_subtitles(
        str(tmp_path / "missing.mp4"), str(srt), str(tmp_path / "out.mp4")
    )
    assert ok is False


def test_batch_extract_screenshots_seeks_only_for_offset_grid(monkeypatch):
    """An offset (interval-aligned) grid seeks the input; a zero grid does not."""
    captured = {}

    def fake_run(command, **_kwargs):
        captured["cmd"] = command
        return None  # short-circuit fallback; we only inspect the built command

    monkeypatch.setattr(video.config, "SCREENSHOT_FORMAT", ".png")
    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    # Multi-video part aligned to the global grid -> first capture at 5s.
    video._batch_extract_screenshots("in.mp4", [5, 15], 10)
    assert "-ss" in captured["cmd"]
    assert captured["cmd"][captured["cmd"].index("-ss") + 1] == "5"

    # Zero-aligned grid -> no input seek (unchanged behavior).
    video._batch_extract_screenshots("in.mp4", [0, 10], 10)
    assert "-ss" not in captured["cmd"]
