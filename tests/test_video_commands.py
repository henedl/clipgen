import json
import os
import subprocess
import threading
import time
from pathlib import Path

import pytest

import config
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


# -- hardware encoder selection --


def test_check_videotoolbox_support_needs_darwin_and_encoder(monkeypatch):
    """macOS gate first, then the -encoders listing; both cached."""
    probes = []

    def fake_probe(listing_arg, tokens):
        probes.append((listing_arg, tokens))
        return True

    monkeypatch.setattr(video, "_probe_ffmpeg_listing", fake_probe)

    monkeypatch.setattr(video.sys, "platform", "linux")
    video._videotoolbox_support_cache = None
    assert video.check_videotoolbox_support() is False
    assert probes == []  # never shells out off macOS

    monkeypatch.setattr(video.sys, "platform", "darwin")
    video._videotoolbox_support_cache = None
    assert video.check_videotoolbox_support() is True
    assert probes == [("-encoders", {"h264_videotoolbox"})]

    # Cached: a second call does not re-probe.
    assert video.check_videotoolbox_support() is True
    assert len(probes) == 1


def test_resolve_video_encoder_honors_knob(monkeypatch):
    monkeypatch.setattr(video, "check_videotoolbox_support", lambda: True)

    monkeypatch.setattr(config, "FFMPEG_VIDEO_ENCODER", "auto")
    assert video.resolve_video_encoder() == "h264_videotoolbox"

    monkeypatch.setattr(config, "FFMPEG_VIDEO_ENCODER", "h264_videotoolbox")
    assert video.resolve_video_encoder() == "h264_videotoolbox"

    monkeypatch.setattr(config, "FFMPEG_VIDEO_ENCODER", "libx264")
    assert video.resolve_video_encoder() == "libx264"

    # A select setting is coerced with str() and never validated against its
    # options list, so junk must degrade quietly rather than raise.
    monkeypatch.setattr(config, "FFMPEG_VIDEO_ENCODER", "nonsense")
    assert video.resolve_video_encoder() == "libx264"


def test_resolve_video_encoder_degrades_when_unsupported(monkeypatch):
    """Explicit hardware request on a machine without it warns once, then libx264."""
    monkeypatch.setattr(video, "check_videotoolbox_support", lambda: False)
    monkeypatch.setattr(config, "FFMPEG_VIDEO_ENCODER", "h264_videotoolbox")
    warnings = []
    monkeypatch.setattr(
        video.utils, "warning_print", lambda msg, *a, **kw: warnings.append(msg)
    )

    assert video.resolve_video_encoder() == "libx264"
    assert video.resolve_video_encoder() == "libx264"
    assert len(warnings) == 1


@pytest.mark.parametrize("choice", ["auto", "h264_videotoolbox"])
def test_resolve_video_encoder_respects_sticky_failure(monkeypatch, choice):
    """A runtime failure is session-sticky for both hardware modes.

    Including the *explicit* choice: on hardware that lists the encoder but can't
    run it, honoring the request again would spend a doomed attempt per encode.
    """
    monkeypatch.setattr(video, "check_videotoolbox_support", lambda: True)
    monkeypatch.setattr(config, "FFMPEG_VIDEO_ENCODER", choice)
    assert video.resolve_video_encoder() == "h264_videotoolbox"

    video._hw_encode_failed = True
    assert video.resolve_video_encoder() == "libx264"


def test_video_encoder_args_software_omits_unset_flags():
    """Sites that passed no crf/preset keep relying on libx264's own defaults."""
    assert video.video_encoder_args("libx264") == ["-c:v", "libx264"]
    assert video.video_encoder_args("libx264", crf=23, preset="fast") == [
        "-c:v",
        "libx264",
        "-preset",
        "fast",
        "-crf",
        "23",
    ]
    assert video.video_encoder_args("libx264", crf=20, preset="veryfast") == [
        "-c:v",
        "libx264",
        "-preset",
        "veryfast",
        "-crf",
        "20",
    ]


def test_video_encoder_args_videotoolbox_maps_crf_to_quality():
    """VideoToolbox has no CRF: -q:v carries the quality, clamped to 30-80."""
    assert video.video_encoder_args("h264_videotoolbox") == [
        "-c:v",
        "h264_videotoolbox",
        "-q:v",
        "54",  # crf 23 (libx264's default) -> 100 - 46
        "-allow_sw",
        "1",
    ]
    assert "-q:v" in video.video_encoder_args("h264_videotoolbox", crf=20)
    assert "-crf" not in video.video_encoder_args("h264_videotoolbox", crf=20)
    assert "-pass" not in video.video_encoder_args("h264_videotoolbox", crf=20)
    # Clamped at both ends rather than emitting a nonsense -q:v.
    assert video.video_encoder_args("h264_videotoolbox", crf=51)[3] == "30"
    assert video.video_encoder_args("h264_videotoolbox", crf=0)[3] == "80"


def test_run_ffmpeg_encode_falls_back_to_libx264_once(monkeypatch):
    """A failing hardware encode retries in software and disables hw for the run."""
    encoders = []
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda cmd, **_kw: subprocess.CompletedProcess(
            args=cmd, returncode=1 if cmd[0] == "h264_videotoolbox" else 0, stderr=""
        ),
    )

    def build(encoder):
        encoders.append(encoder)
        return [encoder]

    result = video.run_ffmpeg_encode(build, encoder="h264_videotoolbox")
    assert result is not None and result.returncode == 0
    assert encoders == ["h264_videotoolbox", "libx264"]
    assert video._hw_encode_failed is True

    # The flag sticks, so the next resolve keeps everything in software.
    monkeypatch.setattr(video, "check_videotoolbox_support", lambda: True)
    monkeypatch.setattr(config, "FFMPEG_VIDEO_ENCODER", "auto")
    assert video.resolve_video_encoder() == "libx264"


def test_run_ffmpeg_encode_no_retry_on_success_or_cancel(monkeypatch):
    """Success runs once; a cancel (None) must not burn the retry or set the flag."""
    calls = []

    def fake_run(cmd, **_kwargs):
        calls.append(cmd)
        return subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)
    video.run_ffmpeg_encode(lambda enc: [enc], encoder="h264_videotoolbox")
    assert calls == [["h264_videotoolbox"]]
    assert video._hw_encode_failed is False

    calls.clear()
    monkeypatch.setattr(video, "run_ffmpeg_process", lambda cmd, **_kw: None)
    assert (
        video.run_ffmpeg_encode(lambda enc: [enc], encoder="h264_videotoolbox") is None
    )
    assert video._hw_encode_failed is False


def test_concatenate_clips_reencode_uses_hardware_encoder(monkeypatch):
    """With hardware encoding on, the concat fallback targets VideoToolbox."""
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "verify_output_file", lambda *_args, **_kwargs: True)
    monkeypatch.setattr(
        video, "probe_video_properties", lambda _path: dict(_MATCHING_PROPS)
    )
    monkeypatch.setattr(video, "check_videotoolbox_support", lambda: True)
    monkeypatch.setattr(config, "FFMPEG_VIDEO_ENCODER", "auto")

    captured = []
    results = [
        subprocess.CompletedProcess(
            args=["ffmpeg"], returncode=1, stderr="copy failed"
        ),
        subprocess.CompletedProcess(args=["ffmpeg"], returncode=0, stderr=""),
    ]
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda command, **_kw: (captured.append(command), results.pop(0))[1],
    )

    assert video.concatenate_clips(
        ["a.mp4", "b.mp4"], "reel.mp4", reencode_on_fail=True
    )
    assert "h264_videotoolbox" in captured[1]
    assert "libx264" not in captured[1]


def test_detect_clip_mismatches_parallels_clip_paths(monkeypatch):
    """props_list stays index-aligned with clip_paths, None where a probe fails."""

    def fake_probe(path):
        return None if path == "b.mp4" else dict(_MATCHING_PROPS)

    monkeypatch.setattr(video, "probe_video_properties", fake_probe)

    props_list, res_mismatch, audio_mismatch = video._detect_clip_mismatches(
        ["a.mp4", "b.mp4", "c.mp4"]
    )
    assert [p is None for p in props_list] == [False, True, False]
    assert res_mismatch is False
    assert audio_mismatch is False


def test_detect_clip_mismatches_single_clip_stays_inline(monkeypatch):
    """One clip probes on the calling thread — no pool for the common case."""
    probe_threads = []

    def recording_probe(_path):
        probe_threads.append(threading.current_thread())
        return dict(_MATCHING_PROPS)

    monkeypatch.setattr(video, "probe_video_properties", recording_probe)

    video._detect_clip_mismatches(["only.mp4"])
    assert probe_threads == [threading.current_thread()]


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
                    "pix_fmt": "yuv420p",
                },
                {
                    "codec_type": "audio",
                    "codec_name": "aac",
                    "sample_rate": "48000",
                    "channels": 2,
                    "channel_layout": "stereo",
                },
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
        "pix_fmt": "yuv420p",
        "audio_sample_rate": 48000,
        "audio_channels": 2,
        "audio_channel_layout": "stereo",
        "audio_tracks": [
            {
                "index": 0,
                "codec": "aac",
                "channels": 2,
                "title": "",
                "language": "",
                "handler": "",
                "label": "Track 1",
            }
        ],
        "audio_track_count": 1,
        "fps": 0.0,
        "duration": 0.0,
        "nb_frames": 0,
        "start_time": 0.0,
    }


def test_probe_video_properties_parses_container_start_time(monkeypatch, tmp_path):
    video._video_properties_cache.clear()
    clip = tmp_path / "clip.ts"
    clip.write_bytes(b"x")
    fake_json = json.dumps(
        {
            "streams": [
                {
                    "codec_type": "video",
                    "codec_name": "mpeg2video",
                    "width": 160,
                    "height": 120,
                }
            ],
            "format": {"duration": "8.0", "start_time": "11.4"},
        }
    )
    monkeypatch.setattr(video.subprocess, "check_output", lambda _cmd, **_kw: fake_json)
    result = video.probe_video_properties(str(clip))
    assert result is not None
    assert result["start_time"] == 11.4
    assert result["duration"] == 8.0


def test_probe_video_properties_multiple_audio_tracks(monkeypatch, tmp_path):
    """Every audio stream is enumerated; the first also fills the flat fields."""
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
                {
                    "codec_type": "audio",
                    "codec_name": "aac",
                    "channels": 2,
                    "tags": {"title": "Microphone"},
                },
                {
                    "codec_type": "audio",
                    "codec_name": "opus",
                    "channels": 1,
                    # Generic muxer handler + language → label falls back to lang.
                    "tags": {"handler_name": "SoundHandler", "language": "eng"},
                },
                {
                    "codec_type": "audio",
                    "codec_name": "aac",
                    "channels": 2,
                    # "und" is MP4's undefined-language default → ordinal label.
                    "tags": {"language": "und"},
                },
            ]
        }
    )
    monkeypatch.setattr(video.subprocess, "check_output", lambda _cmd, **_kw: fake_json)

    result = video.probe_video_properties(str(clip))
    assert result is not None
    assert result["audio_track_count"] == 3
    labels = [t["label"] for t in result["audio_tracks"]]
    assert labels == ["Microphone", "ENG", "Track 3"]
    assert [t["index"] for t in result["audio_tracks"]] == [0, 1, 2]
    # Flat top-level fields describe the first audio stream only.
    assert result["audio_codec"] == "aac"
    assert result["audio_channels"] == 2


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
    # No audio stream → audio params are absent/zero.
    assert result["audio_sample_rate"] == 0
    assert result["audio_channels"] == 0
    assert result["audio_channel_layout"] is None
    assert result["audio_tracks"] == []
    assert result["audio_track_count"] == 0


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


# -- pick_speech_audio_track tests --


def _named_tracks(*labels):
    """Track dicts shaped like probe_video_properties' audio_tracks entries."""
    return [
        {"index": i, "title": label, "handler": "", "label": label}
        for i, label in enumerate(labels)
    ]


@pytest.mark.parametrize(
    ("labels", "expected"),
    [
        # A named mic beats a named system-audio track.
        (("Microphone", "System Audio"), 0),
        (("System Audio", "Participant"), 1),
        # Nothing named → track 0, i.e. exactly what whisper did before this
        # heuristic existed. This is the no-silent-behaviour-change guard.
        (("Track 1", "Track 2"), 0),
        # Only a *negative* signal still moves off track 0.
        (("Speakers", "Track 2"), 1),
        # Two speech-looking tracks tie → lowest index wins.
        (("Mic", "Interview"), 0),
        # Word boundaries: "mic" must not match "dynamic".
        (("Dynamic Range",), 0),
        (("Screen Recording", "Participant Mic"), 1),
        ((), 0),
    ],
)
def test_pick_speech_audio_track(labels, expected):
    assert video.pick_speech_audio_track(_named_tracks(*labels)) == expected


def test_pick_speech_audio_track_reads_handler_and_returns_track_index():
    """Scores the handler name too, and returns the track's own index field."""
    tracks = [
        {"index": 0, "title": "", "handler": "Screen Capture", "label": "Track 1"},
        {"index": 1, "title": "", "handler": "Interview Mic", "label": "Track 2"},
    ]
    assert video.pick_speech_audio_track(tracks) == 1


# -- extract_audio_track tests --


def _stub_ffmpeg_writes_output(monkeypatch, tmp_path, fail_copy=False):
    """Route the audio-track cache to tmp_path and fake ffmpeg by writing the
    output file. Returns the list that records each subprocess.run cmd."""
    monkeypatch.setattr(config, "DEBUGGING", False)
    monkeypatch.setattr(video.tempfile, "gettempdir", lambda: str(tmp_path))
    calls: list[list[str]] = []

    def _run(cmd, **_kw):
        calls.append(cmd)
        if fail_copy and "copy" in cmd:
            raise subprocess.CalledProcessError(1, "ffmpeg")
        Path(cmd[-1]).write_bytes(b"audio")
        return subprocess.CompletedProcess(cmd, 0)

    monkeypatch.setattr(video.subprocess, "run", _run)
    return calls


def test_extract_audio_track_stream_copy_and_cache(monkeypatch, tmp_path):
    src = tmp_path / "v.mp4"
    src.write_bytes(b"x")
    calls = _stub_ffmpeg_writes_output(monkeypatch, tmp_path)

    out = video.extract_audio_track(str(src), 0)
    assert out is not None and out.is_file()
    assert len(calls) == 1  # AAC copy succeeded → no re-encode retry
    assert "copy" in calls[0]

    # Second call hits the on-disk cache — ffmpeg is not invoked again.
    out2 = video.extract_audio_track(str(src), 0)
    assert out2 == out
    assert len(calls) == 1


def test_extract_audio_track_falls_back_to_aac(monkeypatch, tmp_path):
    src = tmp_path / "v.mp4"
    src.write_bytes(b"x")
    calls = _stub_ffmpeg_writes_output(monkeypatch, tmp_path, fail_copy=True)

    out = video.extract_audio_track(str(src), 1)
    assert out is not None and out.is_file()
    assert len(calls) == 2  # copy failed, AAC re-encode succeeded
    assert "aac" in calls[1]


def test_extract_audio_track_failure_returns_none(monkeypatch, tmp_path):
    src = tmp_path / "v.mp4"
    src.write_bytes(b"x")
    monkeypatch.setattr(config, "DEBUGGING", False)
    monkeypatch.setattr(video.tempfile, "gettempdir", lambda: str(tmp_path))

    def _raise(_cmd, **_kw):
        raise subprocess.CalledProcessError(1, "ffmpeg")

    monkeypatch.setattr(video.subprocess, "run", _raise)
    assert video.extract_audio_track(str(src), 0) is None


def test_extract_audio_track_file_not_found():
    assert video.extract_audio_track("/nonexistent/missing.mp4", 0) is None


def test_extract_audio_track_concurrent_calls_run_ffmpeg_once(monkeypatch, tmp_path):
    """Two threads racing on the same track must not both invoke ffmpeg.

    Without the per-key lock both would write the same deterministic .partial
    path with -y and one would rename a half-written file into the cache, which
    is then served for the source's whole mtime lifetime.
    """
    src = tmp_path / "v.mp4"
    src.write_bytes(b"x")
    monkeypatch.setattr(config, "DEBUGGING", False)
    monkeypatch.setattr(video.tempfile, "gettempdir", lambda: str(tmp_path))

    calls: list[list[str]] = []
    tmp_names: list[str] = []
    started = threading.Event()
    calls_lock = threading.Lock()

    def _run(cmd, **_kw):
        with calls_lock:
            calls.append(cmd)
            tmp_names.append(Path(cmd[-1]).name)
        started.set()
        time.sleep(0.05)  # hold the lock long enough for the racer to pile up
        Path(cmd[-1]).write_bytes(b"audio")
        return subprocess.CompletedProcess(cmd, 0)

    monkeypatch.setattr(video.subprocess, "run", _run)

    results: list[Path | None] = []
    results_lock = threading.Lock()

    def _extract():
        out = video.extract_audio_track(str(src), 0)
        with results_lock:
            results.append(out)

    threads = [threading.Thread(target=_extract) for _ in range(2)]
    threads[0].start()
    started.wait(timeout=5)
    threads[1].start()
    for t in threads:
        t.join(timeout=10)

    assert len(calls) == 1, "second caller should have used the cached result"
    assert len(results) == 2
    assert results[0] == results[1]
    assert results[0] is not None and results[0].read_bytes() == b"audio"
    # Scratch paths are per-caller, so even a different track of the same file
    # can never collide on one .partial name.
    assert ".partial." in tmp_names[0]


def test_extract_audio_track_prunes_superseded_extractions(monkeypatch, tmp_path):
    """A re-encoded source keys a new cache file; the old mtime's must go."""
    src = tmp_path / "v.mp4"
    src.write_bytes(b"x")
    _stub_ffmpeg_writes_output(monkeypatch, tmp_path)

    first = video.extract_audio_track(str(src), 0)
    assert first is not None and first.is_file()

    # Rewrite the source so mtime_ns (and therefore the cache key) changes.
    time.sleep(0.01)
    src.write_bytes(b"xy")
    os.utime(src, ns=(time.time_ns(), time.time_ns()))

    second = video.extract_audio_track(str(src), 0)
    assert second is not None and second.is_file()
    assert second != first
    assert not first.exists(), "superseded extraction should be pruned"


# -- probe_max_keyframe_gap tests --


def test_probe_max_keyframe_gap_uses_max_not_median(monkeypatch, tmp_path):
    video._keyframe_gap_cache.clear()
    clip = tmp_path / "clip.mp4"
    clip.write_bytes(b"x")
    # Keyframes at 0, 1, 2, 8 (gaps 1, 1, 6). The max (6.0) is returned — a
    # single long-GOP stretch must not be masked by the surrounding short gaps
    # (median would be 1.0). Interior non-keyframe packets are ignored.
    csv = "0.000000,K__\n0.033000,__\n1.000000,K__\n2.000000,K__\n8.000000,K__"
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
    csv = "0.000000,K__\n0.033000,__\n1.000000,__"
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


def test_accurate_seek_args_zero_returns_empty_list():
    assert video.accurate_seek_args(0.0) == []


def test_accurate_seek_args_is_a_single_pre_input_seek():
    """One pre-input -ss, exact float preserved — no two-stage split.

    Pre-input -ss is frame-accurate on every decoded output (ffmpeg
    decodes-and-discards from the prior keyframe); the old split decoded
    ~2 s of extra frames per extraction for a bit-identical result.
    """
    assert video.accurate_seek_args(12.345) == ["-ss", "12.345"]
    assert video.accurate_seek_args(0.5) == ["-ss", "0.5"]


def test_accurate_seek_pre_post_zero_start_is_single_pre_input():
    assert video.accurate_seek_pre_post(12.345) == (["-ss", "12.345"], [])
    assert video.accurate_seek_pre_post(0.0) == ([], [])
    assert video.accurate_seek_pre_post(0.5, container_start=0.0) == (
        ["-ss", "0.5"],
        [],
    )


def test_accurate_seek_pre_post_nonzero_start_uses_two_stage():
    """MPEG-TS start_time ≠ 0: post-input -ss counts decoded media."""
    assert video.accurate_seek_pre_post(0.5, container_start=11.4) == (
        [],
        ["-ss", "0.5"],
    )
    assert video.accurate_seek_pre_post(3.7, container_start=11.4) == (
        ["-ss", "1.7"],
        ["-ss", "2.0"],
    )


# ---- two-stage seek wiring in extract_frame_at_timestamp ----


def _captured_run(captured: dict):
    def _run(*args, **kwargs):
        captured["cmd"] = list(args[0])
        return subprocess.CompletedProcess(args, 0, stdout=b"")

    return _run


def test_extract_frame_at_timestamp_seeks_before_input_only(monkeypatch):
    """The seek is one pre-input -ss; no post-input -ss survives.

    A post-input -ss on top of a pre-input one is the obsolete two-stage
    idiom — it decodes ~2 s of frames per extraction that ffmpeg then
    discards, for a bit-identical result.
    """
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
    assert cmd[:i_idx][-2:] == ["-ss", "12.5"], "expected pre-input seek"
    assert "-ss" not in cmd[i_idx + 2 :], "post-input seek is the obsolete split"


def test_extract_frame_at_timestamp_two_stage_when_start_time_nonzero(monkeypatch):
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _: {
            "width": 320,
            "height": 240,
            "fps": 30.0,
            "duration": 60.0,
            "start_time": 11.4,
        },
    )
    captured: dict = {}
    monkeypatch.setattr(video.subprocess, "run", _captured_run(captured))

    video.extract_frame_at_timestamp("/fake.ts", 0.5)
    cmd = captured["cmd"]
    i_idx = cmd.index("-i")
    assert "-ss" not in cmd[:i_idx]
    assert cmd[i_idx + 1 : i_idx + 3] == ["/fake.ts", "-ss"]
    assert cmd[i_idx + 3] == "0.5"


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
    # One pre-input seek carrying the exact float — no int() truncation.
    assert seek_values == ["12.75"]
    assert cmd.index("-ss") < cmd.index("-i")


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

    sentinel = lambda: False
    ok = video.run_ffmpeg(
        "in.mp4", "out.mp4", "00:10", "00:15", reencode=False, cancel_flag=sentinel
    )
    assert ok is True
    assert captured == [sentinel]


# -- end-of-recording spans are shortened, not dropped --


def _run_ffmpeg_harness(monkeypatch, captured, *, file_duration):
    """Stub out the filesystem + ffmpeg around run_ffmpeg / extract_gif."""
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video.Path, "stat", lambda self: type("_S", (), {"st_size": 1})()
    )
    monkeypatch.setattr(video, "get_file_duration", lambda *_a: file_duration)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)
    monkeypatch.setattr(video.config, "MAX_FILESIZE_MB", 0)

    def _fake(cmd, **_kwargs):
        captured.append(list(cmd))
        return subprocess.CompletedProcess(args=["ffmpeg"], returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", _fake)


def test_run_ffmpeg_shortens_a_clip_running_past_the_end(monkeypatch):
    """A bare end-of-session timestamp overshoots by DEFAULT_DURATION_SECONDS.

    ffmpeg stops at EOF anyway, and the multi-video path already clamps, so the
    clip is cut short rather than dropped.
    """
    captured: list = []
    _run_ffmpeg_harness(monkeypatch, captured, file_duration=100)

    ok = video.run_ffmpeg("in.mp4", "out.mp4", "01:20", "02:20", reencode=False)

    assert ok is True
    cmd = captured[0]
    assert cmd[cmd.index("-t") + 1] == "20"  # 100 - 80, not the requested 60


def test_run_ffmpeg_leaves_an_in_bounds_clip_alone(monkeypatch):
    captured: list = []
    _run_ffmpeg_harness(monkeypatch, captured, file_duration=100)

    assert video.run_ffmpeg("in.mp4", "out.mp4", "00:10", "00:40", reencode=False)
    cmd = captured[0]
    assert cmd[cmd.index("-t") + 1] == "30"


def test_run_ffmpeg_still_skips_a_start_past_the_end(monkeypatch):
    captured: list = []
    _run_ffmpeg_harness(monkeypatch, captured, file_duration=100)

    assert (
        video.run_ffmpeg("in.mp4", "out.mp4", "02:00", "03:00", reencode=False) is False
    )
    assert captured == []


def test_run_ffmpeg_skips_when_clamping_leaves_nothing(monkeypatch):
    """Defensive floor: a sub-second tail truncates to 0 and must not be cut."""
    captured: list = []
    _run_ffmpeg_harness(monkeypatch, captured, file_duration=80.5)

    assert (
        video.run_ffmpeg("in.mp4", "out.mp4", "01:20", "02:20", reencode=False) is False
    )
    assert captured == []


def test_extract_gif_shortens_a_range_running_past_the_end(monkeypatch):
    captured: list = []
    _run_ffmpeg_harness(monkeypatch, captured, file_duration=100)

    assert video.extract_gif("in.mp4", "out.gif", "01:30", 60) is True
    cmd = captured[0]
    assert cmd[cmd.index("-t") + 1] == "10"  # 100 - 90


def test_extract_gif_still_skips_a_start_past_the_end(monkeypatch):
    captured: list = []
    _run_ffmpeg_harness(monkeypatch, captured, file_duration=100)

    assert video.extract_gif("in.mp4", "out.gif", "02:00", 5) is False
    assert captured == []


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

    sentinel = lambda: False
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

    sentinel = lambda: False
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

    sentinel = lambda: False
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

    sentinel = lambda: False
    # Target tiny size so compression runs both passes.
    video.compress_to_size(str(big), 0.0001, cancel_flag=sentinel)
    assert captured == [sentinel, sentinel]


def test_compress_to_size_stays_on_libx264_with_hardware_available(
    monkeypatch, tmp_path
):
    """Size capping is bitrate targeting — deliberately never hardware-encoded.

    VideoToolbox has no -pass and overshoots a -b:v target badly (measured 2.3x),
    so a hardware attempt would be discarded and the two-pass paid for anyway.
    """
    big = tmp_path / "big.mp4"
    big.write_bytes(b"x" * 4096)
    monkeypatch.setattr(video, "get_file_duration", lambda *_a: 10)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)
    monkeypatch.setattr(video, "check_videotoolbox_support", lambda: True)
    monkeypatch.setattr(config, "FFMPEG_VIDEO_ENCODER", "auto")
    monkeypatch.setattr(video.os, "replace", lambda *_a, **_kw: None)

    commands: list[list[str]] = []

    def fake_run(command, **_kwargs):
        commands.append(command)
        return subprocess.CompletedProcess(args=command, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    video.compress_to_size(str(big), 0.001)
    assert len(commands) == 2  # still exactly the two libx264 passes
    assert all("h264_videotoolbox" not in cmd for cmd in commands)
    assert commands[0][commands[0].index("-pass") + 1] == "1"
    assert commands[1][commands[1].index("-pass") + 1] == "2"


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
    # Video + audio only, never a bare `-map 0`: mapping the source's own
    # subtitle streams would push the new track off output index 0, so the
    # -metadata/-disposition arguments below would land on the wrong one.
    maps = [cmd[i + 1] for i, tok in enumerate(cmd) if tok == "-map"]
    assert maps == ["0:v?", "0:a?", "1:0"]
    assert cmd[cmd.index("-disposition:s:0") + 1] == "default"
    assert cmd[-1] == str(out)


def test_mux_subtitles_normalizes_the_language_to_three_letters(monkeypatch, tmp_path):
    """Whisper reports ISO 639-1 ("en"), containers store only ISO 639-2.

    Measured on ffmpeg 8.1.2: the mp4 muxer drops an "en" tag outright and
    truncates transcripts.py's "unknown" fallback to the nonsense tag "unk" —
    both silently. So the code is normalized before it reaches the muxer, not
    trusted.
    """
    src, srt = _write_dummy_pair(tmp_path)

    def fake_run(command, **_kwargs):
        captured["command"] = command
        return subprocess.CompletedProcess(args=command, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)

    for given, expected in [
        ("en", "eng"),  # the case that shipped untagged mp4s
        ("zh", "zho"),
        ("pt-BR", "por"),  # BCP 47 keeps only the primary subtag
        ("eng", "eng"),  # already 639-2 -> untouched
        ("unknown", "und"),  # transcripts.py's detection fallback
        ("", "und"),
    ]:
        captured: dict = {}
        out = tmp_path / f"out-{expected}.mp4"
        assert video.mux_subtitles(
            str(src), str(srt), str(out), track_language=given
        ), given
        cmd = captured["command"]
        lang_arg = cmd[cmd.index("-metadata:s:s:0") + 1]
        assert lang_arg == f"language={expected}", f"{given!r} -> {lang_arg}"


def test_mux_subtitles_set_default_false_clears_the_disposition(monkeypatch, tmp_path):
    """set_default=False writes an explicit 0 rather than dropping the flag.

    A container whose only subtitle stream is the one we add can mark it default
    on its own, so omitting the flag would not reliably leave the track off.
    """
    src, srt = _write_dummy_pair(tmp_path)
    out = tmp_path / "out.mp4"

    captured = {}

    def fake_run(command, **_kwargs):
        captured["command"] = command
        return subprocess.CompletedProcess(args=command, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)

    ok = video.mux_subtitles(str(src), str(srt), str(out), set_default=False)
    assert ok is True
    cmd = captured["command"]
    assert cmd[cmd.index("-disposition:s:0") + 1] == "0"


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

    monkeypatch.setattr(video.config, "SCREENSHOT_FORMAT", ".png")
    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    # Multi-video part aligned to the global grid -> first capture at 5s.
    video._batch_extract_screenshots("in.mp4", [5, 15], 10)
    assert "-ss" in captured["cmd"]
    assert captured["cmd"][captured["cmd"].index("-ss") + 1] == "5"

    # Zero-aligned grid -> no input seek (unchanged behavior).
    video._batch_extract_screenshots("in.mp4", [0, 10], 10)
    assert "-ss" not in captured["cmd"]


def test_extract_sprite_sheet_seek_composites_grid(monkeypatch, tmp_path):
    """seek_frames=True grabs one frame per slot center and tiles them with PIL."""
    from io import BytesIO

    from PIL import Image

    monkeypatch.setattr(video.config, "DEBUGGING", False)
    src = tmp_path / "clip.mp4"
    src.write_bytes(b"x")  # existence check only; ffmpeg is stubbed

    frame = BytesIO()
    Image.new("RGB", (8, 6), (200, 30, 30)).save(frame, format="JPEG")
    frame_bytes = frame.getvalue()

    calls = []

    def fake_run(cmd, **_kwargs):
        calls.append(cmd)
        return subprocess.CompletedProcess(args=cmd, returncode=0, stdout=frame_bytes)

    monkeypatch.setattr(video.subprocess, "run", fake_run)

    data = video.extract_sprite_sheet_bytes(
        str(src), 10.0, 100.0, 2, 2, seek_frames=True
    )
    assert data is not None
    sheet = Image.open(BytesIO(data))
    assert sheet.size == (16, 12)  # 2x2 grid of 8x6 frames
    # One fast input-seek per frame, sampled at slot centers (pool order varies).
    seeks = sorted(float(c[c.index("-ss") + 1]) for c in calls)
    assert seeks == [10.0 + (i + 0.5) * 25.0 for i in range(4)]


def test_extract_sprite_sheet_seek_all_grabs_fail_returns_none(monkeypatch, tmp_path):
    monkeypatch.setattr(video.config, "DEBUGGING", False)
    src = tmp_path / "clip.mp4"
    src.write_bytes(b"x")

    def fake_run(cmd, **_kwargs):
        return subprocess.CompletedProcess(args=cmd, returncode=1, stdout=b"")

    monkeypatch.setattr(video.subprocess, "run", fake_run)
    assert (
        video.extract_sprite_sheet_bytes(str(src), 0.0, 10.0, 2, 2, seek_frames=True)
        is None
    )
