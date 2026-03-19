"""Tests for the transcripts module (formatting, filtering, write/read roundtrips)."""

from pathlib import Path

import config
import transcripts
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


class TestFmtDisplay:
    def test_seconds_only(self):
        assert transcripts._fmt_display(5.0) == "0:05"

    def test_minutes_and_seconds(self):
        assert transcripts._fmt_display(125.0) == "2:05"

    def test_hours(self):
        assert transcripts._fmt_display(3661.5) == "1:01:01"

    def test_zero(self):
        assert transcripts._fmt_display(0.0) == "0:00"


class TestFmtSrt:
    def test_basic(self):
        assert transcripts._fmt_srt(0.0) == "00:00:00,000"

    def test_with_milliseconds(self):
        assert transcripts._fmt_srt(12.5) == "00:00:12,500"

    def test_over_an_hour(self):
        assert transcripts._fmt_srt(3661.5) == "01:01:01,500"


class TestFmtVtt:
    def test_under_an_hour(self):
        assert transcripts._fmt_vtt(12.5) == "00:12.500"

    def test_over_an_hour(self):
        assert transcripts._fmt_vtt(3661.5) == "01:01:01.500"

    def test_zero(self):
        assert transcripts._fmt_vtt(0.0) == "00:00.000"


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
