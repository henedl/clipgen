"""Real-ffmpeg I/O tests: the pipeline must write usable media, not just argv.

Everything else in the default suite mocks ``run_ffmpeg`` and asserts arguments;
these prove the end of the pipe — a cut clip a player can open, a readable
screenshot, a reel that is the sum of its parts (or, on a failed part, no reel
at all — the silent-short-reel class, locked here on the *real* concat path;
the mock-only version lives in ``test_clip_pipeline.py``), and the
frame-accuracy of the single pre-input seek (an assumption about ffmpeg itself,
unmockable by construction).

Deliberately capped at four (see ``agents/skills/test/SKILL.md``): titlecards
(machine-dependent ``drawtext``), OCR, Whisper, remux (already covered by
``test_container_seekability.py``), and padding/format/worker combinatorics all
stay mocked elsewhere.
"""

import json
import shutil
import subprocess

import pytest
from PIL import Image

import config
import pipeline
import video


def _have_ffmpeg() -> bool:
    return shutil.which("ffmpeg") is not None and shutil.which("ffprobe") is not None


requires_ffmpeg = pytest.mark.skipif(
    not _have_ffmpeg(), reason="needs real ffmpeg/ffprobe to cut media"
)


# Subprocess bound; each call here finishes well under a second.
_TIMEOUT = 30
_FPS = 15
_SIZE = (160, 120)


def _encode(path) -> None:
    """Encode 1 s red then 2 s blue, a keyframe every second, no B-frames."""
    command = ["ffmpeg", "-y", "-v", "error"]
    for color, seconds in (("red", 1), ("blue", 2)):
        command += [
            "-f",
            "lavfi",
            "-i",
            f"color=c={color}:s=160x120:r={_FPS}:d={seconds}",
        ]
    command += ["-filter_complex", "[0:v][1:v]concat=n=2:v=1:a=0[v]", "-map", "[v]"]
    command += ["-c:v", "libx264", "-pix_fmt", "yuv420p", "-g", str(_FPS), "-bf", "0"]
    command += ["-x264-params", f"keyint={_FPS}:min-keyint={_FPS}:scenecut=0"]
    command += ["-f", "mp4", str(path)]
    subprocess.run(command, check=True, capture_output=True, timeout=_TIMEOUT)


def _stream_duration(path) -> float:
    """The video stream's precise duration from ffprobe JSON."""
    command = ["ffprobe", "-v", "error", "-select_streams", "v:0"]
    command += ["-show_entries", "stream=duration", "-of", "json", str(path)]
    result = subprocess.run(
        command, check=True, capture_output=True, text=True, timeout=_TIMEOUT
    )
    return float(json.loads(result.stdout)["streams"][0]["duration"])


def _decode_centers(path) -> list[tuple[int, ...]]:
    """Decode every frame, failing on any decode error; return center RGB pixels."""
    command = ["ffmpeg", "-v", "error", "-xerror", "-i", str(path)]
    command += ["-f", "rawvideo", "-pix_fmt", "rgb24", "pipe:1"]
    result = subprocess.run(command, check=True, capture_output=True, timeout=_TIMEOUT)
    assert result.stderr == b"", result.stderr.decode(errors="replace")
    width, height = _SIZE
    frame_size = width * height * 3
    assert result.stdout and len(result.stdout) % frame_size == 0
    center = ((height // 2) * width + width // 2) * 3
    return [
        tuple(result.stdout[i + center : i + center + 3])
        for i in range(0, len(result.stdout), frame_size)
    ]


def _color(rgb) -> str:
    """Name the dominant channel: it must exceed 150, the others stay below 100."""
    names = ("red", "green", "blue")
    for idx, name in enumerate(names):
        others = [v for j, v in enumerate(rgb) if j != idx]
        if rgb[idx] > 150 and all(v < 100 for v in others):
            return name
    return f"unclear {tuple(rgb)}"


@pytest.fixture(scope="module")
def source_dir(tmp_path_factory):
    """One shared red/blue source at the ``{study}_{participant}.mp4`` contract path."""
    if not _have_ffmpeg():
        pytest.skip("needs real ffmpeg/ffprobe")
    directory = tmp_path_factory.mktemp("io_input")
    _encode(directory / "study_P01.mp4")
    return directory


@pytest.fixture
def output_dir(monkeypatch, source_dir, tmp_path):
    """Point the pipeline at the shared source and a fresh output dir."""
    out = tmp_path / "out"
    out.mkdir()
    monkeypatch.setattr(config, "INPUT_DIR", str(source_dir), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(out), raising=False)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    # Stream copy, no loudnorm, real ffmpeg: the output is the cut itself.
    monkeypatch.setattr(config, "REENCODING", False)
    monkeypatch.setattr(config, "AUDIO_NORMALIZE", False)
    monkeypatch.setattr(config, "DEBUGGING", False)
    # Keep Rich progress/spinner rendering out of the test output.
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(pipeline.utils, "use_progress", lambda: False)
    return out


@requires_ffmpeg
def test_cut_writes_playable_clip(make_clip, output_dir):
    clip = make_clip(value="0:01-0:02")

    generated, records = pipeline.process_clips(
        [clip], output_format="clip", titlecards_enabled=False
    )

    assert generated == 1
    assert len(records) == 1
    outputs = list(output_dir.glob("*.mp4"))
    assert len(outputs) == 1
    # One second of the blue span, not the 3 s source or the red opening.
    assert abs(_stream_duration(outputs[0]) - 1.0) <= 1 / _FPS
    centers = _decode_centers(outputs[0])
    assert _color(centers[len(centers) // 2]) == "blue"


@requires_ffmpeg
def test_screenshot_is_a_readable_image(make_clip, output_dir):
    # Screenshots key off the start time; 0:02 sits a second from either edge.
    clip = make_clip(value="0:02-0:03")

    generated, _records = pipeline.process_clips(
        [clip], output_format="screen", titlecards_enabled=False
    )

    assert generated == 1
    outputs = list(output_dir.glob("*.png"))
    assert len(outputs) == 1
    with Image.open(outputs[0]) as image:
        image.load()
        assert image.size == _SIZE
        rgb = image.convert("RGB").getpixel((_SIZE[0] // 2, _SIZE[1] // 2))
    assert _color(rgb) == "blue"


@requires_ffmpeg
def test_reel_concatenates_both_windows_or_writes_nothing(
    monkeypatch, make_clip, output_dir
):
    def _clips():
        return [
            make_clip(row=3, value="0:01-0:02"),
            make_clip(row=4, value="0:00-0:01"),
        ]

    # One worker cuts in clip order; failing call two must leave no reel.
    with monkeypatch.context() as patch:
        calls: list[str] = []
        real_run_ffmpeg = video.run_ffmpeg

        def fail_second_cut(**kwargs):
            calls.append(kwargs["output_file"])
            if len(calls) == 2:
                return False
            return real_run_ffmpeg(**kwargs)

        patch.setattr(pipeline.video, "run_ffmpeg", fail_second_cut)
        errors: list[tuple] = []
        patch.setattr(pipeline.utils, "error_print", lambda *a, **_k: errors.append(a))

        generated, records = pipeline.process_reel(_clips(), titlecards_enabled=False)

        assert generated == 0
        assert records == []
        assert any("Reel aborted" in e[0] for e in errors)
        # The successful first part was unlinked and its reservation released.
        assert list(output_dir.iterdir()) == []

    # Clean run: blue then red, in request order, each window once.
    generated, records = pipeline.process_reel(_clips(), titlecards_enabled=False)

    assert generated == 1
    assert len(records) == 1
    reel_path = output_dir / "study_reel.mp4"
    assert reel_path.is_file()
    assert abs(_stream_duration(reel_path) - 2.0) <= 2 / _FPS
    centers = _decode_centers(reel_path)
    first, second = centers[len(centers) // 4], centers[3 * len(centers) // 4]
    assert (_color(first), _color(second)) == ("blue", "red")


@requires_ffmpeg
def test_single_seek_matches_two_stage_seek_exactly(tmp_path):
    """The single pre-input -ss must decode the exact frame the old split did.

    accurate_seek_pre_post emits one pre-input -ss and relies on ffmpeg
    decoding-and-discarding from the prior keyframe when the output is
    re-encoded. That is a claim about ffmpeg's behavior, not clipgen's, so it
    can only be locked against the real binary: every extracted frame must be
    bit-identical to the obsolete two-stage form (pre-input -ss to target-2s
    plus post-input -ss 2) it replaced. Needs its own 5 s fixture (not the
    shared red/blue source) so the timestamps straddle the old 2 s split threshold
    — both branches of the old idiom — as well as keyframe boundaries.
    """
    path = str(tmp_path / "seek.mp4")
    command = ["ffmpeg", "-y", "-v", "error"]
    command += ["-f", "lavfi", "-i", "testsrc=duration=5:size=160x120:rate=15"]
    command += ["-c:v", "libx264", "-pix_fmt", "yuv420p", "-g", "15", "-f", "mp4", path]
    subprocess.run(command, check=True, capture_output=True, timeout=_TIMEOUT)

    def raw_frame(pre: list[str], post: list[str]) -> bytes:
        cmd = ["ffmpeg", *pre, "-i", path, *post, "-frames:v", "1"]
        cmd += ["-pix_fmt", "bgr24", "-f", "rawvideo", "-loglevel", "error", "pipe:1"]
        return subprocess.run(
            cmd, capture_output=True, check=True, timeout=_TIMEOUT
        ).stdout

    for ts in (0.4, 1.75, 3.0, 4.6):
        old = raw_frame(
            ["-ss", str(ts - 2.0)] if ts > 2.0 else [],
            ["-ss", "2.0"] if ts > 2.0 else ["-ss", str(ts)],
        )
        new = raw_frame(video.accurate_seek_pre_post(ts)[0], [])
        assert len(new) > 0
        assert new == old, f"frame at t={ts} drifted off the two-stage result"
