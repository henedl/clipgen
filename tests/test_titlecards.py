import subprocess


import titlecards
import video


def test_build_titlecard_frame_uses_png_background_and_scaling(monkeypatch, make_clip):
    clip = make_clip(desc="My description")

    commands = []

    def fake_run_ffmpeg_process(cmd, **_kwargs):
        commands.append(cmd)
        return subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run_ffmpeg_process)
    monkeypatch.setattr(video, "verify_output_file", lambda *_args, **_kwargs: True)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)

    path = titlecards.build_titlecard_frame(clip, "1280x720")
    assert path is not None
    assert commands, "ffmpeg was not invoked for titlecard generation"

    cmd = commands[0]
    joined = " ".join(cmd)
    assert "assets/titlecard.png" in joined
    assert "scale=1280x720:force_original_aspect_ratio=decrease" in joined
    assert "pad=1280:720:(ow-iw)/2:(oh-ih)/2" in joined
    assert "drawtext=text='My description'" in joined


def test_build_titlecard_frame_uses_black_background_when_no_asset(
    monkeypatch, make_clip
):
    clip = make_clip(desc="No asset")

    commands = []

    def fake_run_ffmpeg_process(cmd, **_kwargs):
        commands.append(cmd)
        return subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run_ffmpeg_process)
    monkeypatch.setattr(video, "verify_output_file", lambda *_args, **_kwargs: True)

    def fake_is_file(self):
        return False

    monkeypatch.setattr(titlecards.Path, "is_file", fake_is_file)

    path = titlecards.build_titlecard_frame(clip, "1920x1080")
    assert path is not None
    cmd = commands[0]
    joined = " ".join(cmd)
    assert "assets/titlecard.png" not in joined
    assert "color=c=black:s=1920x1080" in joined
    assert "drawtext=text='No asset'" in joined


def test_wrap_clip_with_cards_respects_disabled_flag(monkeypatch, make_clip):
    clip = make_clip()

    calls = {"build_title": 0, "build_end": 0, "ffmpeg": 0}

    monkeypatch.setattr(titlecards.config, "TITLECARDS_ENABLED", False)
    monkeypatch.setattr(
        titlecards,
        "build_titlecard_frame",
        lambda *_a, **_k: calls.__setitem__("build_title", calls["build_title"] + 1),
    )
    monkeypatch.setattr(
        titlecards,
        "get_or_build_endcard",
        lambda *_a, **_k: calls.__setitem__("build_end", calls["build_end"] + 1),
    )
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda *_a, **_k: calls.__setitem__("ffmpeg", calls["ffmpeg"] + 1),
    )

    ok = titlecards.wrap_clip_with_cards(clip, "clip.mp4")
    assert ok is True
    assert calls == {"build_title": 0, "build_end": 0, "ffmpeg": 0}


def test_wrap_clip_with_cards_single_encode_happy_path(monkeypatch, make_clip):
    clip = make_clip()

    monkeypatch.setattr(titlecards.config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _p: {
            "width": 1280,
            "height": 720,
            "video_codec": "h264",
            "audio_codec": "aac",
            "fps": 30.0,
            "duration": 12.0,
            "nb_frames": 360,
        },
    )
    monkeypatch.setattr(
        titlecards, "build_titlecard_frame", lambda *_a, **_k: "titlecard.mp4"
    )
    monkeypatch.setattr(
        titlecards, "get_or_build_endcard", lambda *_a, **_k: "endcard.mp4"
    )
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_k: True)

    commands = []

    def fake_run_ffmpeg_process(cmd, **_kwargs):
        commands.append(cmd)
        return subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run_ffmpeg_process)
    replaced = {}
    monkeypatch.setattr(
        titlecards.os,
        "replace",
        lambda src, dst: replaced.update({"src": src, "dst": dst}),
    )

    ok = titlecards.wrap_clip_with_cards(clip, "clip.mp4")
    assert ok is True
    assert len(commands) == 1, "expected exactly one ffmpeg invocation"

    cmd = commands[0]
    joined = " ".join(cmd)
    # All three video inputs (title, clip, end) should be present.
    assert cmd.count("-i") == 5  # 3 video inputs + 2 silent anullsrc audio inputs
    assert "titlecard.mp4" in cmd
    assert "endcard.mp4" in cmd
    assert "clip.mp4" in cmd
    # Single concat covers all three segments with audio.
    assert "concat=n=3:v=1:a=1" in joined
    # Clip audio is normalized before concat.
    assert "aresample=48000" in joined
    # Silent audio pads the two cards.
    assert joined.count("anullsrc=channel_layout=stereo:sample_rate=48000") == 2
    # Output replaces the original clip.
    assert replaced.get("dst") == "clip.mp4"


def test_wrap_clip_with_cards_title_only_fallback(monkeypatch, make_clip):
    clip = make_clip()

    monkeypatch.setattr(titlecards.config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _p: {
            "width": 1920,
            "height": 1080,
            "video_codec": "h264",
            "audio_codec": "aac",
            "fps": 30.0,
            "duration": 8.0,
            "nb_frames": 240,
        },
    )
    monkeypatch.setattr(
        titlecards, "build_titlecard_frame", lambda *_a, **_k: "titlecard.mp4"
    )
    # Endcard build fails.
    monkeypatch.setattr(titlecards, "get_or_build_endcard", lambda *_a, **_k: None)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_k: True)

    commands = []
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda cmd, **_k: (
            commands.append(cmd)
            or subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")
        ),
    )
    monkeypatch.setattr(titlecards.os, "replace", lambda src, dst: None)

    ok = titlecards.wrap_clip_with_cards(clip, "clip.mp4")
    assert ok is True
    assert len(commands) == 1
    joined = " ".join(commands[0])
    assert "concat=n=2:v=1:a=1" in joined
    assert "titlecard.mp4" in commands[0]
    assert "endcard.mp4" not in commands[0]


def test_wrap_clip_with_cards_no_audio_drops_audio_stream(monkeypatch, make_clip):
    clip = make_clip()

    monkeypatch.setattr(titlecards.config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _p: {
            "width": 1280,
            "height": 720,
            "video_codec": "h264",
            "audio_codec": None,
            "fps": 30.0,
            "duration": 4.0,
            "nb_frames": 120,
        },
    )
    monkeypatch.setattr(
        titlecards, "build_titlecard_frame", lambda *_a, **_k: "titlecard.mp4"
    )
    monkeypatch.setattr(
        titlecards, "get_or_build_endcard", lambda *_a, **_k: "endcard.mp4"
    )
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_k: True)

    commands = []
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda cmd, **_k: (
            commands.append(cmd)
            or subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")
        ),
    )
    monkeypatch.setattr(titlecards.os, "replace", lambda src, dst: None)

    ok = titlecards.wrap_clip_with_cards(clip, "clip.mp4")
    assert ok is True
    cmd = commands[0]
    joined = " ".join(cmd)
    # No audio stream in output: concat a=0 and no anullsrc inputs.
    assert "concat=n=3:v=1:a=0" in joined
    assert "anullsrc" not in joined
    assert "-c:a" not in cmd


def test_wrap_clip_with_cards_no_cards_is_noop(monkeypatch, make_clip):
    clip = make_clip()

    monkeypatch.setattr(titlecards.config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _p: {
            "width": 1280,
            "height": 720,
            "video_codec": "h264",
            "audio_codec": "aac",
            "fps": 30.0,
            "duration": 4.0,
            "nb_frames": 120,
        },
    )
    monkeypatch.setattr(titlecards, "build_titlecard_frame", lambda *_a, **_k: None)
    monkeypatch.setattr(titlecards, "get_or_build_endcard", lambda *_a, **_k: None)

    ffmpeg_calls = []
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda cmd, **_k: ffmpeg_calls.append(cmd) or None,
    )

    ok = titlecards.wrap_clip_with_cards(clip, "clip.mp4")
    assert ok is True
    assert ffmpeg_calls == []


def test_pipeline_skips_per_output_probe(monkeypatch, make_clip):
    """Regression: pipeline should probe the source video, not each generated output."""
    import pipeline

    clip = make_clip(value="00:10-00:20")
    clip["times"] = [("00:10", "00:20"), ("00:30", "00:40")]
    clip["severity"] = ""

    monkeypatch.setattr(pipeline.config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(pipeline.config, "REENCODING", False)
    monkeypatch.setattr(pipeline.config, "DEBUGGING", False)

    probe_calls = []

    def fake_probe(path):
        probe_calls.append(path)
        return {
            "width": 1280,
            "height": 720,
            "video_codec": "h264",
            "audio_codec": "aac",
            "fps": 30.0,
            "duration": 60.0,
            "nb_frames": 1800,
        }

    monkeypatch.setattr(pipeline.video, "probe_video_properties", fake_probe)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", lambda **_k: True)

    unique_names = iter(["out1.mp4", "out2.mp4"])
    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *_a, **_k: next(unique_names)
    )

    wrap_calls = []

    def fake_wrap(_clip, out_name, resolution=None, **_kwargs):
        wrap_calls.append((out_name, resolution))
        return True

    monkeypatch.setattr(pipeline.titlecards, "wrap_clip_with_cards", fake_wrap)

    generated, _paths = pipeline._process_single_clip_segments(
        clip, "source.mp4", set(), output_format="clip"
    )

    assert generated == 2
    # Source is probed once for two segments; outputs are not probed separately.
    assert probe_calls == ["source.mp4"]
    assert wrap_calls == [
        ("out1.mp4", "1280x720"),
        ("out2.mp4", "1280x720"),
    ]
