"""Tests for the transcripts module (formatting, filtering, write/read roundtrips)."""

from pathlib import Path

import pytest

import config
import transcripts
import utils
from transcripts import TranscriptResult, TranscriptSegment


# ---------------------------------------------------------------------------
# Fixtures / helpers
# ---------------------------------------------------------------------------


def _sample_result(source="study_P01.mp4") -> TranscriptResult:
    return TranscriptResult(
        segments=[
            TranscriptSegment(start=0.0, end=12.5, text="First segment."),
            TranscriptSegment(start=12.5, end=25.0, text="Second segment."),
            TranscriptSegment(start=30.0, end=45.0, text="Third segment."),
            TranscriptSegment(start=3661.5, end=3675.0, text="Over one hour in."),
        ],
        language="en",
        source_file=source,
        model="base",
    )


def _empty_result() -> TranscriptResult:
    return TranscriptResult(
        segments=[], language="en", source_file="empty.mp4", model="base"
    )


# ---------------------------------------------------------------------------
# Timestamp formatting helpers
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "fmt,seconds,expected",
    [
        ("display", 0.0, "0:00"),
        ("display", 5.0, "0:05"),
        ("display", 125.0, "2:05"),
        ("display", 3661.5, "1:01:01"),
        ("srt", 0.0, "00:00:00,000"),
        ("srt", 12.5, "00:00:12,500"),
        ("srt", 3661.5, "01:01:01,500"),
        ("vtt", 0.0, "00:00.000"),
        ("vtt", 12.5, "00:12.500"),
        ("vtt", 3661.5, "01:01:01.500"),
    ],
)
def test_format_timestamp(fmt, seconds, expected):
    assert transcripts._format_timestamp(seconds, fmt) == expected


# ---------------------------------------------------------------------------
# filter_segments
# ---------------------------------------------------------------------------


class TestFilterSegments:
    def test_filters_to_range(self):
        result = _sample_result()
        filtered = transcripts.filter_segments(result, 10.0, 30.0)
        assert len(filtered["segments"]) == 2
        assert filtered["segments"][0]["text"] == "First segment."
        assert filtered["segments"][1]["text"] == "Second segment."

    def test_no_overlap_returns_empty(self):
        result = _sample_result()
        filtered = transcripts.filter_segments(result, 100.0, 200.0)
        assert filtered["segments"] == []

    def test_preserves_metadata(self):
        result = _sample_result()
        filtered = transcripts.filter_segments(result, 0.0, 13.0)
        assert filtered["language"] == "en"
        assert filtered["model"] == "base"
        assert filtered["source_file"] == "study_P01.mp4"

    def test_offset_to_zero(self):
        result = _sample_result()
        filtered = transcripts.filter_segments(result, 10.0, 30.0, offset_to_zero=True)
        assert len(filtered["segments"]) == 2
        # First segment started at 0.0, clipped start_sec=10.0 → max(0, 0-10)=0
        assert filtered["segments"][0]["start"] == 0.0
        assert filtered["segments"][0]["end"] == 2.5  # 12.5 - 10
        assert filtered["segments"][1]["start"] == 2.5  # 12.5 - 10
        assert filtered["segments"][1]["end"] == 15.0  # 25.0 - 10

    def test_offset_to_zero_clamps_negative_start(self):
        result = TranscriptResult(
            segments=[TranscriptSegment(start=5.0, end=15.0, text="Spanning.")],
            language="en",
            source_file="x.mp4",
            model="base",
        )
        filtered = transcripts.filter_segments(result, 10.0, 20.0, offset_to_zero=True)
        assert filtered["segments"][0]["start"] == 0.0
        assert filtered["segments"][0]["end"] == 5.0

    def test_empty_segments(self):
        filtered = transcripts.filter_segments(_empty_result(), 0.0, 100.0)
        assert filtered["segments"] == []


# ---------------------------------------------------------------------------
# Markdown format / roundtrip
# ---------------------------------------------------------------------------


class TestMarkdownFormat:
    def test_format_contains_header_and_segments(self):
        text = transcripts._format_markdown(_sample_result())
        assert "# Transcript: study_P01.mp4" in text
        assert "**Source:** study_P01.mp4" in text
        assert "**Model:** base" in text
        assert "**Language:** en" in text
        assert "**[0:00 - 0:12]**" in text
        assert "First segment." in text
        assert "**[1:01:01 - 1:01:15]**" in text

    def test_empty_segments_produces_header_only(self):
        text = transcripts._format_markdown(_empty_result())
        assert "# Transcript:" in text
        assert "**[" not in text

    def test_write_and_read_roundtrip(self, tmp_path):
        result = _sample_result()
        path = str(tmp_path / "transcript.md")
        assert transcripts.write_transcript(result, path, fmt="md")

        loaded = transcripts.read_transcript(path)
        assert loaded is not None
        assert loaded["language"] == "en"
        assert loaded["model"] == "base"
        assert len(loaded["segments"]) == len(result["segments"])
        for orig, parsed in zip(result["segments"], loaded["segments"]):
            assert parsed["text"] == orig["text"]
            assert parsed["start"] == int(orig["start"])  # md loses sub-second
            assert parsed["end"] == int(orig["end"])


# ---------------------------------------------------------------------------
# SRT format / roundtrip
# ---------------------------------------------------------------------------


class TestSrtFormat:
    def test_format_structure(self):
        text = transcripts._format_srt(_sample_result())
        assert text.startswith("1\n00:00:00,000 --> 00:00:12,500\nFirst segment.")
        assert "\n\n2\n" in text

    def test_empty_segments(self):
        assert transcripts._format_srt(_empty_result()) == ""

    def test_write_and_read_roundtrip(self, tmp_path):
        result = _sample_result()
        path = str(tmp_path / "transcript.srt")
        assert transcripts.write_transcript(result, path, fmt="srt")

        loaded = transcripts.read_transcript(path)
        assert loaded is not None
        assert len(loaded["segments"]) == len(result["segments"])
        for orig, parsed in zip(result["segments"], loaded["segments"]):
            assert parsed["text"] == orig["text"]
            assert abs(parsed["start"] - orig["start"]) < 0.01
            assert abs(parsed["end"] - orig["end"]) < 0.01


# ---------------------------------------------------------------------------
# VTT format / roundtrip
# ---------------------------------------------------------------------------


class TestVttFormat:
    def test_format_starts_with_webvtt(self):
        text = transcripts._format_vtt(_sample_result())
        assert text.startswith("WEBVTT\n")
        assert "00:12.500 --> 00:25.000" in text

    def test_hour_timestamps(self):
        text = transcripts._format_vtt(_sample_result())
        assert "01:01:01.500 --> 01:01:15.000" in text

    def test_empty_segments(self):
        text = transcripts._format_vtt(_empty_result())
        assert text.startswith("WEBVTT")
        assert "-->" not in text

    def test_write_and_read_roundtrip(self, tmp_path):
        result = _sample_result()
        path = str(tmp_path / "transcript.vtt")
        assert transcripts.write_transcript(result, path, fmt="vtt")

        loaded = transcripts.read_transcript(path)
        assert loaded is not None
        assert len(loaded["segments"]) == len(result["segments"])
        for orig, parsed in zip(result["segments"], loaded["segments"]):
            assert parsed["text"] == orig["text"]
            assert abs(parsed["start"] - orig["start"]) < 0.01
            assert abs(parsed["end"] - orig["end"]) < 0.01


# ---------------------------------------------------------------------------
# get_transcript_extension
# ---------------------------------------------------------------------------


class TestGetTranscriptExtension:
    def test_defaults_to_config(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_FORMAT", "vtt")
        assert transcripts.get_transcript_extension() == ".vtt"

    def test_explicit_format(self):
        assert transcripts.get_transcript_extension("srt") == ".srt"
        assert transcripts.get_transcript_extension("md") == ".md"
        assert transcripts.get_transcript_extension("vtt") == ".vtt"

    def test_unknown_format_falls_back_to_md(self):
        assert transcripts.get_transcript_extension("txt") == ".md"


# ---------------------------------------------------------------------------
# write_transcript / read_transcript edge cases
# ---------------------------------------------------------------------------


class TestWriteRead:
    def test_write_uses_config_format(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_FORMAT", "srt")
        path = str(tmp_path / "out.srt")
        assert transcripts.write_transcript(_sample_result(), path)
        content = Path(path).read_text()
        assert content.startswith("1\n")

    def test_read_nonexistent_returns_none(self):
        assert transcripts.read_transcript("/no/such/file.md") is None

    def test_write_bad_path_returns_false(self):
        assert not transcripts.write_transcript(_sample_result(), "/no/such/dir/out.md")


# ---------------------------------------------------------------------------
# _build_transcribe_kwargs / transcribe_video Whisper options
# ---------------------------------------------------------------------------


class TestBuildTranscribeKwargs:
    def test_defaults_include_vad(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_VAD_FILTER", True)
        kwargs = transcripts._build_transcribe_kwargs(
            language="en", initial_prompt="test"
        )
        assert kwargs["vad_filter"] is True
        assert kwargs["language"] == "en"
        assert kwargs["initial_prompt"] == "test"
        assert kwargs["no_speech_threshold"] == config.TRANSCRIBE_NO_SPEECH_THRESHOLD
        assert (
            kwargs["compression_ratio_threshold"]
            == config.TRANSCRIBE_COMPRESSION_RATIO_THRESHOLD
        )

    def test_vad_can_be_disabled(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_VAD_FILTER", False)
        kwargs = transcripts._build_transcribe_kwargs(language=None, initial_prompt="")
        assert kwargs["vad_filter"] is False

    def test_vad_parameters_included_when_vad_on(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_VAD_FILTER", True)
        monkeypatch.setattr(config, "TRANSCRIBE_VAD_THRESHOLD", 0.3)
        monkeypatch.setattr(config, "TRANSCRIBE_VAD_SPEECH_PAD_MS", 400)
        monkeypatch.setattr(config, "TRANSCRIBE_VAD_MIN_SILENCE_MS", 2000)
        kwargs = transcripts._build_transcribe_kwargs(language=None, initial_prompt="")
        assert kwargs["vad_parameters"] == {
            "threshold": 0.3,
            "speech_pad_ms": 400,
            "min_silence_duration_ms": 2000,
        }

    def test_vad_parameters_omitted_when_vad_off(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_VAD_FILTER", False)
        kwargs = transcripts._build_transcribe_kwargs(language=None, initial_prompt="")
        assert "vad_parameters" not in kwargs

    def test_beam_size_from_config(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_BEAM_SIZE", 2)
        kwargs = transcripts._build_transcribe_kwargs(language=None, initial_prompt="")
        assert kwargs["beam_size"] == 2

    def test_hallucination_threshold_enables_word_timestamps(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD", 2.0)
        kwargs = transcripts._build_transcribe_kwargs(language=None, initial_prompt="")
        assert kwargs["hallucination_silence_threshold"] == 2.0
        assert kwargs["word_timestamps"] is True

    def test_hallucination_threshold_zero_omits_word_timestamps(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD", 0.0)
        kwargs = transcripts._build_transcribe_kwargs(language=None, initial_prompt="")
        assert "hallucination_silence_threshold" not in kwargs
        assert "word_timestamps" not in kwargs


class TestTranscribeVideoWhisperKwargs:
    def test_transcribe_passes_kwargs_to_model(self, monkeypatch):
        captured: dict = {}

        class FakeSeg:
            text = " hello"
            start = 0.0
            end = 1.0

        class FakeInfo:
            language = "en"

        class FakeModel:
            def transcribe(self, path: str, **kwargs):  # noqa: ARG002
                captured.update(kwargs)
                return iter([FakeSeg()]), FakeInfo()

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(
            transcripts, "_load_model", lambda model_name=None: FakeModel()
        )
        monkeypatch.setattr(config, "TRANSCRIBE_VAD_FILTER", True)
        monkeypatch.setattr(config, "TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD", 0.0)

        result = transcripts.transcribe_video("/fake/video.mp4")
        assert result is not None
        assert captured["vad_filter"] is True
        assert "word_timestamps" not in captured

    def test_no_audio_stream_returns_none_without_loading_model(
        self, monkeypatch, capsys
    ):
        monkeypatch.setattr(config, "DEBUGGING", False)

        def _fake_probe(_path):
            return {
                "width": 1920,
                "height": 1080,
                "video_codec": "h264",
                "audio_codec": None,
                "fps": 60.0,
                "duration": 10.0,
                "nb_frames": 600,
            }

        import video as video_mod

        monkeypatch.setattr(video_mod, "probe_video_properties", _fake_probe)

        def _fail_load(model_name=None):  # noqa: ARG001
            raise AssertionError("_load_model must not be called when audio is absent")

        monkeypatch.setattr(transcripts, "_load_model", _fail_load)

        result = transcripts.transcribe_video("/fake/no_audio.mp4")
        assert result is None
        assert "No audio stream" in capsys.readouterr().out


# ---------------------------------------------------------------------------
# transcribe_video in debug mode
# ---------------------------------------------------------------------------


class TestTranscribeVideoDebug:
    def test_debug_returns_stub(self, monkeypatch):
        monkeypatch.setattr(config, "DEBUGGING", True)
        result = transcripts.transcribe_video("/fake/video.mp4")
        assert result is not None
        assert result["segments"] == []
        assert result["language"] == "en"
        assert result["source_file"] == "/fake/video.mp4"
        assert result["model"] == config.TRANSCRIBE_MODEL

    def test_debug_respects_language_arg(self, monkeypatch):
        monkeypatch.setattr(config, "DEBUGGING", True)
        result = transcripts.transcribe_video("/fake/video.mp4", language="de")
        assert result is not None
        assert result["language"] == "de"


# ---------------------------------------------------------------------------
# apply_corrections
# ---------------------------------------------------------------------------


class TestApplyCorrections:
    def test_no_corrections_returns_copy(self):
        segs = [transcripts.TranscriptSegment(start=0, end=1, text="hello")]
        result = transcripts.apply_corrections(segs, [])
        assert result == segs
        assert result is not segs

    def test_single_correction(self):
        segs = [transcripts.TranscriptSegment(start=0, end=1, text="teh quick fox")]
        corrections = [{"from": "teh", "to": "the"}]
        result = transcripts.apply_corrections(segs, corrections)
        assert result[0]["text"] == "the quick fox"

    def test_multiple_corrections(self):
        segs = [transcripts.TranscriptSegment(start=0, end=1, text="teh quck fox")]
        corrections = [
            {"from": "teh", "to": "the"},
            {"from": "quck", "to": "quick"},
        ]
        result = transcripts.apply_corrections(segs, corrections)
        assert result[0]["text"] == "the quick fox"

    def test_no_mutation_of_input(self):
        segs = [transcripts.TranscriptSegment(start=0, end=1, text="teh fox")]
        original_text = segs[0]["text"]
        transcripts.apply_corrections(segs, [{"from": "teh", "to": "the"}])
        assert segs[0]["text"] == original_text

    def test_case_insensitive(self):
        segs = [transcripts.TranscriptSegment(start=0, end=1, text="Teh quick TEH fox")]
        corrections = [{"from": "teh", "to": "the"}]
        result = transcripts.apply_corrections(segs, corrections)
        assert result[0]["text"] == "the quick the fox"

    def test_correction_not_found(self):
        segs = [transcripts.TranscriptSegment(start=0, end=1, text="hello world")]
        corrections = [{"from": "xyz", "to": "abc"}]
        result = transcripts.apply_corrections(segs, corrections)
        assert result[0]["text"] == "hello world"

    def test_skips_empty_from_or_to(self):
        segs = [transcripts.TranscriptSegment(start=0, end=1, text="hello")]
        corrections = [{"from": "", "to": "x"}, {"from": "y", "to": ""}]
        result = transcripts.apply_corrections(segs, corrections)
        assert result[0]["text"] == "hello"


# ---------------------------------------------------------------------------
# get_corrections_keywords
# ---------------------------------------------------------------------------


class TestGetCorrectionsKeywords:
    def test_empty(self):
        assert transcripts.get_corrections_keywords([]) == []

    def test_extracts_unique_to_values(self):
        corrections = [
            {"from": "a", "to": "alpha"},
            {"from": "b", "to": "beta"},
            {"from": "c", "to": "alpha"},  # duplicate
        ]
        result = transcripts.get_corrections_keywords(corrections)
        assert result == ["alpha", "beta"]

    def test_strips_whitespace(self):
        corrections = [{"from": "a", "to": "  alpha  "}]
        result = transcripts.get_corrections_keywords(corrections)
        assert result == ["alpha"]

    def test_skips_empty_to(self):
        corrections = [{"from": "a", "to": ""}, {"from": "b", "to": "  "}]
        assert transcripts.get_corrections_keywords(corrections) == []


# ---------------------------------------------------------------------------
# Transcripts manifest I/O
# ---------------------------------------------------------------------------


class TestTranscriptsManifest:
    def test_empty_manifest_default(self):
        m = transcripts._empty_transcripts_manifest()
        assert m == {"source_transcripts": {}, "corrections": [], "marks": []}

    def test_load_missing_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        m = transcripts.load_transcripts_manifest()
        assert m == {"source_transcripts": {}, "corrections": [], "marks": []}

    def test_save_and_load_roundtrip(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        source = {
            "P01": {
                "segments": [{"start": 0.0, "end": 1.0, "text": "hello"}],
                "language": "en",
                "model": "base",
                "source_file": "video.mp4",
                "transcribed_at": "2025-01-01T00:00:00+00:00",
            }
        }
        corrections = [
            {
                "id": "c1",
                "from": "helo",
                "to": "hello",
                "created": "2025-01-01T00:00:00+00:00",
            }
        ]
        path = transcripts.save_transcripts_manifest(source, corrections)
        assert path is not None

        loaded = transcripts.load_transcripts_manifest()
        assert loaded["corrections"] == corrections
        assert "P01" in loaded["source_transcripts"]
        assert len(loaded["source_transcripts"]["P01"]["segments"]) == 1

    def test_segment_ids_assigned_on_save(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        source = {
            "P01": {
                "segments": [
                    {"start": 0.0, "end": 1.0, "text": "first"},
                    {"start": 1.0, "end": 2.0, "text": "second"},
                ],
                "language": "en",
                "model": "base",
                "source_file": "v.mp4",
                "transcribed_at": "2025-01-01T00:00:00+00:00",
            }
        }
        transcripts.save_transcripts_manifest(source, [])
        loaded = transcripts.load_transcripts_manifest()
        segs = loaded["source_transcripts"]["P01"]["segments"]
        assert segs[0]["id"] == "P01:0"
        assert segs[1]["id"] == "P01:1"

    def test_corrections_preserved(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        corrections = [
            {
                "id": "c1",
                "from": "teh",
                "to": "the",
                "created": "2025-01-01T00:00:00+00:00",
            },
            {
                "id": "c2",
                "from": "recieve",
                "to": "receive",
                "created": "2025-01-01T00:00:00+00:00",
            },
        ]
        transcripts.save_transcripts_manifest({}, corrections)
        loaded = transcripts.load_transcripts_manifest()
        assert loaded["corrections"] == corrections

    def test_empty_save_writes_no_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        path = transcripts.save_transcripts_manifest({}, [])
        assert path is None
        assert not (tmp_path / config.TRANSCRIPTS_MANIFEST_FILENAME).exists()

    def test_emptying_existing_manifest_removes_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        manifest = tmp_path / config.TRANSCRIPTS_MANIFEST_FILENAME
        transcripts.save_transcripts_manifest(
            {}, [{"id": "c1", "from": "a", "to": "b", "created": "2025-01-01T00:00:00"}]
        )
        assert manifest.is_file()
        # Pass marks=[] explicitly to defeat the marks-from-disk preservation.
        transcripts.save_transcripts_manifest({}, [], marks=[])
        assert not manifest.exists()

    def test_load_corrupt_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        (tmp_path / config.TRANSCRIPTS_MANIFEST_FILENAME).write_text("not json")
        m = transcripts.load_transcripts_manifest()
        assert m == {"source_transcripts": {}, "corrections": [], "marks": []}

    def test_load_returns_independent_deep_copies(self, tmp_path, monkeypatch):
        """Mutating a returned entry in place must not corrupt the cache."""
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        transcripts._reset_transcripts_manifest_cache()
        source = {
            "P01": {
                "segments": [{"start": 0.0, "end": 1.0, "text": "hi"}],
                "language": "en",
                "model": "base",
                "source_file": "v.mp4",
                "transcribed_at": "2025-01-01T00:00:00+00:00",
            }
        }
        transcripts.save_transcripts_manifest(source, [])

        first = transcripts.load_transcripts_manifest()
        # Deep-mutate a nested entry exactly like --summarize / --citations do.
        first["source_transcripts"]["P01"]["summary"] = "leaked"

        second = transcripts.load_transcripts_manifest()
        assert "summary" not in second["source_transcripts"]["P01"]

    def test_repeated_load_reuses_cache(self, tmp_path, monkeypatch):
        """A second load with an unchanged file must not re-read/parse from disk."""
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        transcripts._reset_transcripts_manifest_cache()
        transcripts.save_transcripts_manifest(
            {}, [{"id": "c1", "from": "a", "to": "b", "created": "2025-01-01T00:00:00"}]
        )

        calls = {"n": 0}
        real_load = utils.load_json_manifest

        def _counting_load(*args, **kwargs):
            calls["n"] += 1
            return real_load(*args, **kwargs)

        monkeypatch.setattr(utils, "load_json_manifest", _counting_load)

        transcripts.load_transcripts_manifest()  # miss -> one disk read
        transcripts.load_transcripts_manifest()  # hit  -> no disk read
        assert calls["n"] == 1

    def test_save_busts_cache(self, tmp_path, monkeypatch):
        """After a save the next load must reflect the new data, not the cache."""
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        transcripts._reset_transcripts_manifest_cache()
        transcripts.save_transcripts_manifest(
            {}, [{"id": "c1", "from": "a", "to": "b", "created": "2025-01-01T00:00:00"}]
        )
        assert len(transcripts.load_transcripts_manifest()["corrections"]) == 1

        transcripts.save_transcripts_manifest(
            {},
            [
                {"id": "c1", "from": "a", "to": "b", "created": "2025-01-01T00:00:00"},
                {"id": "c2", "from": "c", "to": "d", "created": "2025-01-01T00:00:00"},
            ],
        )
        assert len(transcripts.load_transcripts_manifest()["corrections"]) == 2


# ---------------------------------------------------------------------------
# ManifestSegment type
# ---------------------------------------------------------------------------


class TestManifestSegment:
    def test_fields(self):
        seg = transcripts.ManifestSegment(id="P01:0", start=0.0, end=1.0, text="hi")
        assert seg["id"] == "P01:0"
        assert seg["start"] == 0.0
        assert seg["end"] == 1.0
        assert seg["text"] == "hi"


# ---------------------------------------------------------------------------
# Shared participant video discovery
# ---------------------------------------------------------------------------


class TestDiscoverParticipantVideos:
    def test_discovers_participant_videos(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path))
        (tmp_path / "study_P01.mp4").touch()
        (tmp_path / "study_P02.mp4").touch()
        (tmp_path / "study_G01.mp4").touch()
        (tmp_path / "other.mp4").touch()  # no underscore-separated participant ID
        result = utils.discover_participant_videos()
        ids = [p["id"] for p in result]
        assert "P01" in ids
        assert "P02" in ids
        assert "G01" in ids
        assert len(result) == 3
        for p in result:
            assert p["has_video"] is True

    def test_empty_dir(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path))
        result = utils.discover_participant_videos()
        assert result == []

    def test_nonexistent_dir(self, monkeypatch):
        monkeypatch.setattr(config, "INPUT_DIR", "/nonexistent/path")
        result = utils.discover_participant_videos()
        assert result == []

    def test_ignores_non_participant_prefix(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path))
        (tmp_path / "study_X01.mp4").touch()
        (tmp_path / "study_notes.mp4").touch()
        result = utils.discover_participant_videos()
        assert result == []


# ---------------------------------------------------------------------------
# TranscriptWorker and task creation
# ---------------------------------------------------------------------------


class TestCreateTranscriptTask:
    def test_task_shape(self):
        task = transcripts.create_transcript_task("P01", ["/path/to/video.mp4"])
        assert task["id"].startswith("tr_")
        assert len(task["id"]) == 11  # "tr_" + 8 hex chars
        assert task["participant"] == "P01"
        assert task["video_paths"] == ["/path/to/video.mp4"]
        assert task["status"] == "queued"
        assert task["progress"] == 0.0
        assert task["result"] is None
        assert task["error"] is None
        assert task["created_at"] is not None
        assert task["completed_at"] is None

    def test_unique_ids(self):
        t1 = transcripts.create_transcript_task("P01", ["/path/v.mp4"])
        t2 = transcripts.create_transcript_task("P01", ["/path/v.mp4"])
        assert t1["id"] != t2["id"]


class TestTranscriptWorker:
    def test_restore_tasks(self):
        worker = transcripts.TranscriptWorker()
        tasks = [
            {"id": "tr_abc12345", "status": "completed", "participant": "P01"},
            {"id": "tr_def67890", "status": "failed", "participant": "P02"},
        ]
        worker.restore_tasks(tasks)
        all_tasks = worker.get_all_tasks()
        assert len(all_tasks) == 2
        ids = {t["id"] for t in all_tasks}
        assert "tr_abc12345" in ids
        assert "tr_def67890" in ids

    def test_get_task(self):
        worker = transcripts.TranscriptWorker()
        worker.restore_tasks([{"id": "tr_test1234", "status": "completed"}])
        task = worker.get_task("tr_test1234")
        assert task is not None
        assert task["id"] == "tr_test1234"
        assert worker.get_task("nonexistent") is None

    def test_get_all_tasks_slim_omits_partial_segments(self):
        """include_partials=False drops the growing segment tail, reports count."""
        worker = transcripts.TranscriptWorker()
        task = transcripts.create_transcript_task("P01", ["/v.mp4"])
        worker.enqueue(task)
        with worker._lock:
            worker._tasks[task["id"]]["partial_segments"] = [
                {"start": 0.0, "end": 1.0, "text": "a"},
                {"start": 1.0, "end": 2.0, "text": "b"},
            ]
        slim = worker.get_all_tasks(include_partials=False)[0]
        assert "partial_segments" not in slim
        assert slim["partial_count"] == 2
        # Default path still carries the full array.
        full = worker.get_all_tasks()[0]
        assert len(full["partial_segments"]) == 2

    def test_get_partial_segments_tail(self):
        """Since-cursor returns exactly the appended tail plus the running total."""
        worker = transcripts.TranscriptWorker()
        task = transcripts.create_transcript_task("P01", ["/v.mp4"])
        worker.enqueue(task)
        segs = [{"start": float(i), "end": i + 1.0, "text": str(i)} for i in range(3)]
        with worker._lock:
            worker._tasks[task["id"]]["partial_segments"] = segs
        tail, total = worker.get_partial_segments(task["id"], since=1)
        assert total == 3
        assert [s["text"] for s in tail] == ["1", "2"]
        # A cursor at/after the end yields no new segments.
        assert worker.get_partial_segments(task["id"], since=3) == ([], 3)
        # Unknown task → empty tail (the endpoint just reports no new segments).
        assert worker.get_partial_segments("nope", 0) == ([], 0)

    def test_cancel_queued_task(self):
        worker = transcripts.TranscriptWorker()
        task = transcripts.create_transcript_task("P01", ["/v.mp4"])
        worker.enqueue(task)
        assert worker.cancel(task["id"]) is True
        t = worker.get_task(task["id"])
        assert t is not None
        assert t["status"] == "cancelled"

    def test_cancel_running_task(self):
        worker = transcripts.TranscriptWorker()
        task = transcripts.create_transcript_task("P01", ["/v.mp4"])
        worker.enqueue(task)
        # Simulate the worker picking up the task
        with worker._lock:
            worker._tasks[task["id"]]["status"] = "running"
        assert worker.cancel(task["id"]) is True
        t = worker.get_task(task["id"])
        assert t is not None
        assert t["_cancelled"] is True

    def test_cancel_nonexistent(self):
        worker = transcripts.TranscriptWorker()
        assert worker.cancel("nonexistent") is False

    def test_execute_task_empty_video_paths_fails_cleanly(self, monkeypatch):
        """A task with no video files fails cleanly instead of raising IndexError."""
        from unittest.mock import Mock

        import video as video_mod

        probe = Mock()
        monkeypatch.setattr(video_mod, "probe_video_properties", probe)

        worker = transcripts.TranscriptWorker()
        task = transcripts.create_transcript_task("P01", [])
        task["status"] = "running"

        worker._execute_task(task)

        assert task["status"] == transcripts.TASK_STATUS_FAILED
        assert task["partial_segments"] == []
        assert task["error"]
        probe.assert_not_called()

    def test_execute_task_cancelled_before_model_load(self, monkeypatch):
        """A task cancelled before model load aborts without loading a model."""
        from unittest.mock import Mock

        import video as video_mod

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(video_mod, "probe_video_properties", lambda *_a, **_k: None)
        monkeypatch.setattr(
            transcripts, "load_transcripts_manifest", lambda: {"corrections": []}
        )
        load_model = Mock()
        monkeypatch.setattr(transcripts, "_load_model", load_model)

        worker = transcripts.TranscriptWorker()
        task = transcripts.create_transcript_task("P01", ["/v.mp4"])
        task["status"] = "running"
        task["_cancelled"] = True

        worker._execute_task(task)

        assert task["status"] == "cancelled"
        assert task["partial_segments"] == []
        load_model.assert_not_called()

    def test_execute_task_cancelled_between_segments(self, monkeypatch):
        """Cancelling mid-transcription stops between segments and clears the
        partial-segment buffer."""
        from types import SimpleNamespace

        import video as video_mod

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(
            video_mod,
            "probe_video_properties",
            lambda *_a, **_k: {"duration": 100.0, "audio_codec": "aac"},
        )
        monkeypatch.setattr(
            transcripts, "load_transcripts_manifest", lambda: {"corrections": []}
        )

        worker = transcripts.TranscriptWorker()
        task = transcripts.create_transcript_task("P01", ["/v.mp4"])
        task["status"] = "running"

        class FakeSeg:
            def __init__(self, t):
                self.start = float(t)
                self.end = float(t + 1)
                self.text = f"seg {t}"

        def fake_segments():
            for i in range(5):
                if i == 1:
                    # Trip cancellation once the first segment is consumed.
                    task["_cancelled"] = True
                yield FakeSeg(i)

        class FakeModel:
            def transcribe(self, path, **kwargs):  # noqa: ARG002
                return fake_segments(), SimpleNamespace(language="en")

        monkeypatch.setattr(transcripts, "_load_model", lambda *_a, **_k: FakeModel())

        worker._execute_task(task)

        assert task["status"] == "cancelled"
        assert task["partial_segments"] == []

    def test_execute_task_fails_fast_when_video_has_no_audio(self, monkeypatch):
        """A video with no audio stream fails immediately with a friendly error
        and never loads the Whisper model."""
        from unittest.mock import Mock

        import video as video_mod

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(
            video_mod,
            "probe_video_properties",
            lambda *_a, **_k: {"duration": 100.0, "audio_codec": None},
        )
        monkeypatch.setattr(
            transcripts, "load_transcripts_manifest", lambda: {"corrections": []}
        )
        load_model = Mock()
        monkeypatch.setattr(transcripts, "_load_model", load_model)

        worker = transcripts.TranscriptWorker()
        task = transcripts.create_transcript_task("P01", ["/v.mp4"])
        task["status"] = "running"

        worker._execute_task(task)

        assert task["status"] == "failed"
        assert "No audio stream" in task["error"]
        assert task["partial_segments"] == []
        load_model.assert_not_called()

    def test_remove_task(self):
        worker = transcripts.TranscriptWorker()
        task = transcripts.create_transcript_task("P01", ["/v.mp4"])
        worker.enqueue(task)
        assert worker.remove_task(task["id"]) is True
        assert worker.get_task(task["id"]) is None

    def test_remove_running_task_sets_cancelled(self):
        worker = transcripts.TranscriptWorker()
        task = transcripts.create_transcript_task("P01", ["/v.mp4"])
        worker.enqueue(task)
        with worker._lock:
            worker._tasks[task["id"]]["status"] = "running"
        assert worker.remove_task(task["id"]) is True
        # Task is removed, but the _cancelled flag was set before removal
        assert worker.get_task(task["id"]) is None

    def test_remove_nonexistent(self):
        worker = transcripts.TranscriptWorker()
        assert worker.remove_task("nonexistent") is False

    def test_is_alive_before_start(self):
        worker = transcripts.TranscriptWorker()
        assert worker.is_alive is False

    def test_start_and_stop(self):
        worker = transcripts.TranscriptWorker()
        worker.start()
        assert worker.is_alive is True
        worker.stop()
        assert worker.is_alive is False


# ---------------------------------------------------------------------------
# Whisper model cache detection (gates silent downloads)
# ---------------------------------------------------------------------------


class TestLoadModelCpuThreads:
    @staticmethod
    def _capture_model(monkeypatch):
        """Patch faster_whisper.WhisperModel to record constructor kwargs and
        reset the module-level model cache so the load actually runs."""
        seen: dict = {}

        class FakeModel:
            def __init__(self, name, **kwargs):
                seen["name"] = name
                seen["kwargs"] = kwargs

        import faster_whisper

        monkeypatch.setattr(faster_whisper, "WhisperModel", FakeModel)
        monkeypatch.setattr(transcripts, "_cached_model", None)
        monkeypatch.setattr(transcripts, "_cached_model_name", None)
        return seen

    def test_cpu_threads_passed_when_set(self, monkeypatch):
        seen = self._capture_model(monkeypatch)
        monkeypatch.setattr(config, "TRANSCRIBE_CPU_THREADS", 7)
        transcripts._load_model("base")
        assert seen["kwargs"].get("cpu_threads") == 7

    def test_cpu_threads_auto_resolves_to_cpu_count(self, monkeypatch):
        seen = self._capture_model(monkeypatch)
        monkeypatch.setattr(config, "TRANSCRIBE_CPU_THREADS", 0)
        monkeypatch.setattr(transcripts.os, "cpu_count", lambda: 6)
        transcripts._load_model("base")
        assert seen["kwargs"].get("cpu_threads") == 6

    def test_cpu_threads_omitted_when_unresolvable(self, monkeypatch):
        seen = self._capture_model(monkeypatch)
        monkeypatch.setattr(config, "TRANSCRIBE_CPU_THREADS", 0)
        monkeypatch.setattr(transcripts.os, "cpu_count", lambda: None)
        transcripts._load_model("base")
        assert "cpu_threads" not in seen["kwargs"]


class TestIsWhisperModelCached:
    def test_debugging_short_circuits_true(self, monkeypatch):
        monkeypatch.setattr(config, "DEBUGGING", True)
        assert transcripts.is_whisper_model_cached("large-v3") is True

    def test_true_when_download_model_resolves(self, monkeypatch):
        monkeypatch.setattr(config, "DEBUGGING", False)
        import faster_whisper.utils as fwu

        monkeypatch.setattr(fwu, "download_model", lambda *a, **k: "/cache/path")
        assert transcripts.is_whisper_model_cached("base") is True

    def test_false_when_download_model_raises(self, monkeypatch):
        monkeypatch.setattr(config, "DEBUGGING", False)
        import faster_whisper.utils as fwu

        def _boom(*a, **k):
            raise FileNotFoundError("not in cache")

        monkeypatch.setattr(fwu, "download_model", _boom)
        assert transcripts.is_whisper_model_cached("large-v3") is False

    def test_passes_local_files_only_and_name(self, monkeypatch):
        monkeypatch.setattr(config, "DEBUGGING", False)
        import faster_whisper.utils as fwu

        seen: dict = {}

        def _capture(name, **kwargs):
            seen["name"] = name
            seen["local_files_only"] = kwargs.get("local_files_only")
            return "/cache/path"

        monkeypatch.setattr(fwu, "download_model", _capture)
        transcripts.is_whisper_model_cached("medium")
        assert seen["name"] == "medium"
        assert seen["local_files_only"] is True
