"""Tests for the transcripts module (formatting, filtering, write/read roundtrips)."""

from pathlib import Path

import numpy as np
import pytest

import config
import transcripts
import utils
from transcripts import TranscriptResult, TranscriptSegment


# ---------------------------------------------------------------------------
# Fixtures / helpers
# ---------------------------------------------------------------------------

# One second of silence in the exact shape video.decode_audio_pcm hands the
# model: 16 kHz mono float32.
_FAKE_AUDIO = np.zeros(16000, dtype=np.float32)


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

    def test_offset_to_zero_shifts_words(self):
        result = TranscriptResult(
            segments=[
                TranscriptSegment(
                    start=12.0,
                    end=14.0,
                    text="hello world",
                    words=[
                        {"start": 9.5, "end": 12.5, "text": "hello"},
                        {"start": 12.6, "end": 14.0, "text": "world"},
                    ],
                )
            ],
            language="en",
            source_file="x.mp4",
            model="base",
        )
        filtered = transcripts.filter_segments(result, 10.0, 20.0, offset_to_zero=True)
        seg = filtered["segments"][0]
        assert seg["start"] == 2.0
        # First word started before the clip window: clamped to 0, end intact.
        assert seg["words"] == [
            {"start": 0.0, "end": 2.5, "text": "hello"},
            {"start": 2.6, "end": 4.0, "text": "world"},
        ]

    def test_non_offset_path_keeps_words(self):
        result = TranscriptResult(
            segments=[
                TranscriptSegment(
                    start=1.0,
                    end=2.0,
                    text="hi",
                    words=[{"start": 1.0, "end": 2.0, "text": "hi"}],
                )
            ],
            language="en",
            source_file="x.mp4",
            model="base",
        )
        filtered = transcripts.filter_segments(result, 0.0, 10.0)
        assert filtered["segments"][0]["words"] == [
            {"start": 1.0, "end": 2.0, "text": "hi"}
        ]


# ---------------------------------------------------------------------------
# Energy edge-snap
# ---------------------------------------------------------------------------


def _tone_burst_audio(
    lead: float = 0.5, tone: float = 1.0, tail: float = 0.7, sr: int = 16000
):
    """Quiet noise floor + a loud 440 Hz burst + quiet tail (deterministic)."""

    def _sine(duration, freq, amp):
        t = np.arange(int(duration * sr)) / sr
        return (amp * np.sin(2 * np.pi * freq * t)).astype(np.float32)

    return np.concatenate(
        [_sine(lead, 3000, 0.001), _sine(tone, 440, 0.3), _sine(tail, 3000, 0.001)]
    )


class TestEnergySnap:
    """Speech spans exactly [0.5, 1.5] in _tone_burst_audio. Frames are 30 ms on
    a 10 ms hop, so onset lands in the frame at 0.48 (new start 0.45 after the
    30 ms lead-in) and the last speech frame starts at 1.49 (new end 1.58 after
    frame width + release)."""

    def test_snaps_start_to_onset_and_trims_end(self):
        snap = transcripts._build_energy_snapper(_tone_burst_audio())
        assert snap is not None
        seg = TranscriptSegment(start=0.2, end=2.0, text="x")
        snap(seg, 0.0)
        assert seg["start"] == pytest.approx(0.45)
        assert seg["end"] == pytest.approx(1.58)

    def test_end_never_extends(self):
        snap = transcripts._build_energy_snapper(_tone_burst_audio())
        assert snap is not None
        seg = TranscriptSegment(start=0.7, end=1.2, text="x")  # ends mid-speech
        snap(seg, 0.0)
        assert seg["end"] == pytest.approx(1.2)

    def test_end_unchanged_when_no_speech_in_window(self):
        # End overshoot larger than the search window: quiet trailing audio must
        # not be amputated on a guess, so the end stays put.
        snap = transcripts._build_energy_snapper(_tone_burst_audio())
        assert snap is not None
        seg = TranscriptSegment(start=0.2, end=2.15, text="x")
        snap(seg, 0.0)
        assert seg["end"] == pytest.approx(2.15)

    def test_start_never_crosses_prev_end(self):
        snap = transcripts._build_energy_snapper(_tone_burst_audio())
        assert snap is not None
        seg = TranscriptSegment(start=0.5, end=2.0, text="x")
        snap(seg, 0.47)
        assert seg["start"] == pytest.approx(0.47)

    def test_start_never_crosses_first_word_and_words_clamped(self):
        snap = transcripts._build_energy_snapper(_tone_burst_audio())
        assert snap is not None
        seg = TranscriptSegment(
            start=0.2,
            end=2.0,
            text="a b",
            words=[
                {"start": 0.4, "end": 0.46, "text": "a"},
                {"start": 1.0, "end": 1.9, "text": "b"},
            ],
        )
        snap(seg, 0.0)
        # Onset says 0.45, but the first word ends at 0.46: clamp to 0.41.
        assert seg["start"] == pytest.approx(0.41)
        words = seg["words"]
        assert words[0]["start"] == pytest.approx(0.41)  # pulled into the span
        assert words[-1]["end"] == seg["end"]  # trimmed with the segment
        # Monotonic non-decreasing across the list.
        flat = [t for w in words for t in (w["start"], w["end"])]
        assert flat == sorted(flat)

    def test_short_first_word_cannot_undo_prev_end_floor(self):
        # First word ends only 20 ms after prev_end: the word-margin ceiling
        # (word end - 50 ms) lands below prev_end and must lose to it, or the
        # segment would overlap its predecessor and steal the playhead highlight.
        snap = transcripts._build_energy_snapper(_tone_burst_audio())
        assert snap is not None
        seg = TranscriptSegment(
            start=0.5,
            end=2.0,
            text="a b",
            words=[
                {"start": 0.45, "end": 0.48, "text": "a"},
                {"start": 1.0, "end": 1.9, "text": "b"},
            ],
        )
        snap(seg, 0.46)
        assert seg["start"] == pytest.approx(0.46)

    def test_sub_min_duration_segment_cannot_overlap_predecessor(self):
        # Segment shorter than the 100 ms minimum: the min-duration ceiling
        # (end - 100 ms) lands below prev_end and must lose to it.
        snap = transcripts._build_energy_snapper(_tone_burst_audio())
        assert snap is not None
        seg = TranscriptSegment(start=0.5, end=0.58, text="x")
        snap(seg, 0.5)
        assert seg["start"] == pytest.approx(0.5)
        assert seg["end"] == pytest.approx(0.58)

    def test_low_dynamic_range_disables_snap(self):
        constant = _tone_burst_audio(lead=0.0, tone=2.0, tail=0.0)
        assert transcripts._build_energy_snapper(constant) is None
        assert transcripts._build_energy_snapper(np.zeros(32000, np.float32)) is None

    def test_short_audio_disables_snap(self):
        assert (
            transcripts._build_energy_snapper(_tone_burst_audio(0.1, 0.3, 0.1)) is None
        )
        assert transcripts._build_energy_snapper(None) is None


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

    def test_hallucination_threshold_zero_omits_the_kwarg(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD", 0.0)
        monkeypatch.setattr(config, "TRANSCRIBE_WORD_TIMESTAMPS", False)
        kwargs = transcripts._build_transcribe_kwargs(language=None, initial_prompt="")
        assert "hallucination_silence_threshold" not in kwargs
        assert "word_timestamps" not in kwargs

    def test_word_timestamps_default_on(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_WORD_TIMESTAMPS", True)
        kwargs = transcripts._build_transcribe_kwargs(language=None, initial_prompt="")
        assert kwargs["word_timestamps"] is True

    def test_word_timestamps_can_be_disabled(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_WORD_TIMESTAMPS", False)
        monkeypatch.setattr(config, "TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD", 0.0)
        kwargs = transcripts._build_transcribe_kwargs(language=None, initial_prompt="")
        assert "word_timestamps" not in kwargs

    def test_hallucination_forces_word_timestamps_even_when_knob_off(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_WORD_TIMESTAMPS", False)
        monkeypatch.setattr(config, "TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD", 2.0)
        kwargs = transcripts._build_transcribe_kwargs(language=None, initial_prompt="")
        assert kwargs["word_timestamps"] is True

    def test_hotwords_omitted_by_default(self):
        kwargs = transcripts._build_transcribe_kwargs(language=None, initial_prompt="")
        assert "hotwords" not in kwargs

    def test_hotwords_passed_through(self):
        kwargs = transcripts._build_transcribe_kwargs(
            language=None, initial_prompt="", hotwords="Frobnicator, Widget Bay"
        )
        assert kwargs["hotwords"] == "Frobnicator, Widget Bay"

    def test_empty_hotwords_omitted(self):
        kwargs = transcripts._build_transcribe_kwargs(
            language=None, initial_prompt="", hotwords=""
        )
        assert "hotwords" not in kwargs


class TestTranscribeVideoWhisperKwargs:
    def test_transcribe_passes_kwargs_to_model(self, monkeypatch):
        import video as video_mod

        captured: dict = {}

        class FakeSeg:
            text = " hello"
            start = 0.0
            end = 1.0

        class FakeInfo:
            language = "en"

        class FakeModel:
            def transcribe(self, audio, **kwargs):
                captured.update(kwargs)
                return iter([FakeSeg()]), FakeInfo()

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(
            transcripts, "_load_model", lambda model_name=None: FakeModel()
        )
        monkeypatch.setattr(
            video_mod, "decode_audio_pcm", lambda _path, _idx=0, **_kw: _FAKE_AUDIO
        )
        monkeypatch.setattr(config, "TRANSCRIBE_VAD_FILTER", True)
        monkeypatch.setattr(config, "TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD", 0.0)
        monkeypatch.setattr(config, "TRANSCRIBE_WORD_TIMESTAMPS", False)

        result = transcripts.transcribe_video("/fake/video.mp4")
        assert result is not None
        assert captured["vad_filter"] is True
        assert "word_timestamps" not in captured

    def test_known_terms_become_hotwords(self, monkeypatch):
        import video as video_mod

        captured: dict = {}

        class FakeSeg:
            text = " hello"
            start = 0.0
            end = 1.0

        class FakeInfo:
            language = "en"

        class FakeModel:
            def transcribe(self, audio, **kwargs):
                captured.update(kwargs)
                return iter([FakeSeg()]), FakeInfo()

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(
            transcripts, "_load_model", lambda model_name=None: FakeModel()
        )
        monkeypatch.setattr(
            video_mod, "decode_audio_pcm", lambda _path, _idx=0, **_kw: _FAKE_AUDIO
        )

        result = transcripts.transcribe_video(
            "/fake/video.mp4", known_terms=["Frobnicator", "Widget Bay"]
        )
        assert result is not None
        assert captured["hotwords"] == "Frobnicator, Widget Bay"

    def test_segments_carry_rounded_words_and_tightened_bounds(self, monkeypatch):
        """Word timing rides on the segment; segment bounds tighten to the words."""
        from types import SimpleNamespace

        import video as video_mod

        class FakeSeg:
            text = " hello world"
            start = 0.0  # looser than the words: includes the VAD pad
            end = 2.0
            words = [
                SimpleNamespace(start=0.401, end=0.9, word=" hello"),
                SimpleNamespace(start=0.95, end=1.402, word=" world"),
            ]

        class FakeInfo:
            language = "en"

        class FakeModel:
            def transcribe(self, audio, **_kwargs):
                return iter([FakeSeg()]), FakeInfo()

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(config, "TRANSCRIBE_EDGE_SNAP", False)
        monkeypatch.setattr(
            transcripts, "_load_model", lambda model_name=None: FakeModel()
        )
        monkeypatch.setattr(
            video_mod, "decode_audio_pcm", lambda _path, _idx=0, **_kw: _FAKE_AUDIO
        )

        streamed: list = []
        result = transcripts.transcribe_video(
            "/fake/video.mp4", on_segment=lambda end, seg: streamed.append((end, seg))
        )
        assert result is not None
        seg = result["segments"][0]
        assert seg["start"] == 0.4  # words[0].start, rounded to 2 dp
        assert seg["end"] == 1.4  # words[-1].end, rounded to 2 dp
        assert seg["words"] == [
            {"start": 0.4, "end": 0.9, "text": "hello"},
            {"start": 0.95, "end": 1.4, "text": "world"},
        ]
        # The streamed partial is the same (finalized) dict as the result's.
        assert streamed == [(1.4, seg)]

    def test_segments_without_words_keep_whisper_bounds(self, monkeypatch):
        import video as video_mod

        class FakeSeg:
            text = " hello"
            start = 0.25
            end = 1.75

        class FakeInfo:
            language = "en"

        class FakeModel:
            def transcribe(self, audio, **_kwargs):
                return iter([FakeSeg()]), FakeInfo()

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(config, "TRANSCRIBE_EDGE_SNAP", False)
        monkeypatch.setattr(
            transcripts, "_load_model", lambda model_name=None: FakeModel()
        )
        monkeypatch.setattr(
            video_mod, "decode_audio_pcm", lambda _path, _idx=0, **_kw: _FAKE_AUDIO
        )

        result = transcripts.transcribe_video("/fake/video.mp4")
        assert result is not None
        seg = result["segments"][0]
        assert seg["start"] == 0.25
        assert seg["end"] == 1.75
        assert "words" not in seg

    def test_edge_snap_applied_before_streaming(self, monkeypatch):
        """The snapper sees each segment with the previous (snapped) end as floor,
        and mutations land before on_segment fires."""
        import video as video_mod

        class FakeSeg:
            def __init__(self, start, end, text):
                self.start = start
                self.end = end
                self.text = text

        class FakeInfo:
            language = "en"

        class FakeModel:
            def transcribe(self, audio, **_kwargs):
                return iter([FakeSeg(0.0, 2.0, "one"), FakeSeg(2.5, 4.0, "two")]), (
                    FakeInfo()
                )

        snapped_with: list = []

        def _fake_snapper(segment, prev_end):
            snapped_with.append((segment["text"], prev_end))
            segment["end"] = segment["end"] - 0.5  # pretend we trimmed silence

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(config, "TRANSCRIBE_EDGE_SNAP", True)
        monkeypatch.setattr(
            transcripts, "_build_energy_snapper", lambda _audio: _fake_snapper
        )
        monkeypatch.setattr(
            transcripts, "_load_model", lambda model_name=None: FakeModel()
        )
        monkeypatch.setattr(
            video_mod, "decode_audio_pcm", lambda _path, _idx=0, **_kw: _FAKE_AUDIO
        )

        streamed: list = []
        result = transcripts.transcribe_video(
            "/fake/video.mp4", on_segment=lambda end, seg: streamed.append(end)
        )
        assert result is not None
        # Second segment's floor is the first's *snapped* end (1.5, not 2.0).
        assert snapped_with == [("one", 0.0), ("two", 1.5)]
        # on_segment received the snapped end times.
        assert streamed == [1.5, 3.5]

    def test_window_shifts_segments_onto_file_timeline(self, monkeypatch):
        """With start_seconds set, the decode is windowed and every stored /
        streamed segment (words included) is shifted back by the window start."""
        from types import SimpleNamespace

        import video as video_mod

        class FakeSeg:
            text = " hello"
            start = 1.0  # window-relative: the decoded array starts at -ss
            end = 2.0
            words = [SimpleNamespace(start=1.1, end=1.9, word=" hello")]

        class FakeInfo:
            language = "en"

        class FakeModel:
            def transcribe(self, audio, **_kwargs):
                return iter([FakeSeg()]), FakeInfo()

        decode_kwargs = {}

        def fake_decode(_path, _idx=0, **kw):
            decode_kwargs.update(kw)
            return _FAKE_AUDIO

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(config, "TRANSCRIBE_EDGE_SNAP", False)
        monkeypatch.setattr(
            transcripts, "_load_model", lambda model_name=None: FakeModel()
        )
        monkeypatch.setattr(video_mod, "decode_audio_pcm", fake_decode)

        streamed: list = []
        result = transcripts.transcribe_video(
            "/fake/video.mp4",
            start_seconds=10.0,
            end_seconds=30.0,
            on_segment=lambda end, seg: streamed.append((end, seg)),
        )
        assert result is not None
        assert decode_kwargs == {"start_seconds": 10.0, "duration_seconds": 20.0}
        seg = result["segments"][0]
        # Bounds tighten to the words (window-relative 1.1–1.9), then shift.
        assert seg["start"] == 11.1
        assert seg["end"] == 11.9
        assert seg["words"] == [{"start": 11.1, "end": 11.9, "text": "hello"}]
        # The streamed partial is the shifted dict, with the shifted end time.
        assert streamed == [(11.9, seg)]

    def test_window_snap_runs_before_shift(self, monkeypatch):
        """The snapper (built on the windowed array) sees window-relative times
        and the window-relative _prev_end floor; the shift happens after."""
        import video as video_mod

        class FakeSeg:
            def __init__(self, start, end, text):
                self.start = start
                self.end = end
                self.text = text

        class FakeInfo:
            language = "en"

        class FakeModel:
            def transcribe(self, audio, **_kwargs):
                return iter([FakeSeg(0.0, 2.0, "one"), FakeSeg(2.5, 4.0, "two")]), (
                    FakeInfo()
                )

        snapped_with: list = []

        def _fake_snapper(segment, prev_end):
            snapped_with.append((segment["start"], prev_end))
            segment["end"] = segment["end"] - 0.5

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(config, "TRANSCRIBE_EDGE_SNAP", True)
        monkeypatch.setattr(
            transcripts, "_build_energy_snapper", lambda _audio: _fake_snapper
        )
        monkeypatch.setattr(
            transcripts, "_load_model", lambda model_name=None: FakeModel()
        )
        monkeypatch.setattr(
            video_mod, "decode_audio_pcm", lambda _path, _idx=0, **_kw: _FAKE_AUDIO
        )

        streamed: list = []
        result = transcripts.transcribe_video(
            "/fake/video.mp4",
            start_seconds=100.0,
            on_segment=lambda end, seg: streamed.append(end),
        )
        assert result is not None
        # Window-relative starts and floors — not 100.x / 101.5.
        assert snapped_with == [(0.0, 0.0), (2.5, 1.5)]
        # Streamed (and stored) ends are on the file's timeline.
        assert streamed == [101.5, 103.5]
        assert [s["end"] for s in result["segments"]] == [101.5, 103.5]

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

        def _fail_load(model_name=None):
            raise AssertionError("_load_model must not be called when audio is absent")

        monkeypatch.setattr(transcripts, "_load_model", _fail_load)

        result = transcripts.transcribe_video("/fake/no_audio.mp4")
        assert result is None
        assert "No audio stream" in capsys.readouterr().out


# ---------------------------------------------------------------------------
# transcribe_video audio-track selection
# ---------------------------------------------------------------------------


def _multitrack_probe(*labels):
    """probe_video_properties stub for a file with the given named audio tracks."""

    def _probe(_path):
        return {
            "width": 1920,
            "height": 1080,
            "video_codec": "h264",
            "audio_codec": "aac",
            "fps": 60.0,
            "duration": 10.0,
            "nb_frames": 600,
            "audio_tracks": [
                {"index": i, "title": label, "handler": "", "label": label}
                for i, label in enumerate(labels)
            ],
            "audio_track_count": len(labels),
        }

    return _probe


class TestTranscribeVideoAudioTrack:
    """The selected stream is decoded to PCM by video.decode_audio_pcm
    (``-map 0:a:N``). These pin which stream actually reaches the model."""

    def _install_model(self, monkeypatch):
        captured: dict = {}

        class FakeSeg:
            text = " hello"
            start = 0.0
            end = 1.0

        class FakeInfo:
            language = "en"

        class FakeModel:
            def transcribe(self, audio, **_kwargs):
                captured["audio"] = audio
                return iter([FakeSeg()]), FakeInfo()

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(
            transcripts, "_load_model", lambda model_name=None: FakeModel()
        )
        return captured

    @staticmethod
    def _install_decode(monkeypatch):
        import video as video_mod

        calls: list = []
        monkeypatch.setattr(
            video_mod,
            "decode_audio_pcm",
            lambda path, idx=0, **_kw: calls.append((path, idx)) or _FAKE_AUDIO,
        )
        return calls

    def test_track_zero_decodes_stream_zero(self, monkeypatch):
        import video as video_mod

        captured = self._install_model(monkeypatch)
        calls = self._install_decode(monkeypatch)
        monkeypatch.setattr(
            video_mod, "probe_video_properties", _multitrack_probe("Mic", "System")
        )

        result = transcripts.transcribe_video("/fake/video.mp4", audio_index=0)
        assert result is not None
        assert calls == [("/fake/video.mp4", 0)]
        assert captured["audio"] is _FAKE_AUDIO
        # The PCM array is transient; source_file stays the video.
        assert result["source_file"] == "/fake/video.mp4"

    def test_nonzero_track_decodes_that_stream(self, monkeypatch):
        import video as video_mod

        captured = self._install_model(monkeypatch)
        calls = self._install_decode(monkeypatch)
        monkeypatch.setattr(
            video_mod, "probe_video_properties", _multitrack_probe("System", "Mic")
        )

        result = transcripts.transcribe_video("/fake/video.mp4", audio_index=1)
        assert result is not None
        assert calls == [("/fake/video.mp4", 1)]
        assert captured["audio"] is _FAKE_AUDIO
        assert result["source_file"] == "/fake/video.mp4"

    def test_auto_detects_the_speech_track(self, monkeypatch):
        import video as video_mod

        captured = self._install_model(monkeypatch)
        calls = self._install_decode(monkeypatch)
        monkeypatch.setattr(
            video_mod,
            "probe_video_properties",
            _multitrack_probe("System Audio", "Participant Mic"),
        )

        result = transcripts.transcribe_video("/fake/video.mp4")
        assert result is not None
        assert calls == [("/fake/video.mp4", 1)]
        assert captured["audio"] is _FAKE_AUDIO

    def test_failed_decode_fails_loudly(self, monkeypatch, capsys):
        """Never fall back to another track — that transcribes the wrong audio."""
        import video as video_mod

        self._install_model(monkeypatch)
        monkeypatch.setattr(
            video_mod, "probe_video_properties", _multitrack_probe("System", "Mic")
        )
        monkeypatch.setattr(
            video_mod, "decode_audio_pcm", lambda _path, _idx=0, **_kw: None
        )

        assert transcripts.transcribe_video("/fake/video.mp4", audio_index=1) is None
        assert "Could not decode audio track 2" in capsys.readouterr().out

    def test_out_of_range_track_returns_none_without_loading_model(
        self, monkeypatch, capsys
    ):
        import video as video_mod

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(
            video_mod, "probe_video_properties", _multitrack_probe("Mic", "System")
        )

        def _fail_load(model_name=None):
            raise AssertionError("_load_model must not be called for a bad index")

        monkeypatch.setattr(transcripts, "_load_model", _fail_load)

        assert transcripts.transcribe_video("/fake/video.mp4", audio_index=5) is None
        assert "has 2 audio track(s)" in capsys.readouterr().out


# ---------------------------------------------------------------------------
# transcribe_video in debug mode
# ---------------------------------------------------------------------------


class TestDecodeAudioPcmWindow:
    """decode_audio_pcm's in/out window rides on the ffmpeg argv: input-side
    -ss (so PTS zero = window start) plus -t after the -map."""

    def _decode(self, monkeypatch, **kwargs):
        import video as video_mod

        captured = {}

        class FakeResult:
            returncode = 0
            stdout = np.zeros(16, dtype=np.float32).tobytes()
            stderr = b""

        def fake_run(cmd, **_kw):
            captured["cmd"] = cmd
            return FakeResult()

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(video_mod.subprocess, "run", fake_run)
        result = video_mod.decode_audio_pcm("/fake/video.mp4", 0, **kwargs)
        assert result is not None
        return captured["cmd"]

    def test_windowed_argv_has_ss_before_input_and_t(self, monkeypatch):
        cmd = self._decode(monkeypatch, start_seconds=10.0, duration_seconds=20.0)
        assert cmd.index("-ss") < cmd.index("-i")
        assert cmd[cmd.index("-ss") + 1] == "10.000"
        assert cmd.index("-t") > cmd.index("-map")
        assert cmd[cmd.index("-t") + 1] == "20.000"

    def test_unwindowed_argv_is_unchanged(self, monkeypatch):
        cmd = self._decode(monkeypatch)
        assert "-ss" not in cmd
        assert "-t" not in cmd


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

    def test_preserves_words(self):
        words: list[transcripts.TranscriptWord] = [
            {"start": 0.0, "end": 0.5, "text": "teh"},
            {"start": 0.5, "end": 1.0, "text": "fox"},
        ]
        segs = [
            transcripts.TranscriptSegment(start=0, end=1, text="teh fox", words=words)
        ]
        result = transcripts.apply_corrections(segs, [{"from": "teh", "to": "the"}])
        assert result[0]["text"] == "the fox"
        assert result[0]["words"] == words

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
# get_known_terms
# ---------------------------------------------------------------------------


class TestGetKnownTerms:
    def test_missing_key(self):
        assert transcripts.get_known_terms({}) == []

    def test_strips_and_drops_empties(self):
        m = {"known_terms": ["  Frobnicator  ", "", "   "]}
        assert transcripts.get_known_terms(m) == ["Frobnicator"]

    def test_dedupes_case_insensitively_keeping_first_spelling(self):
        m = {"known_terms": ["Frobnicator", "frobnicator", "Widget Bay"]}
        assert transcripts.get_known_terms(m) == ["Frobnicator", "Widget Bay"]

    def test_preserves_order(self):
        m = {"known_terms": ["b", "a", "c"]}
        assert transcripts.get_known_terms(m) == ["b", "a", "c"]


# ---------------------------------------------------------------------------
# Dictionary CSV interchange
# ---------------------------------------------------------------------------


class TestDictionaryCsv:
    def test_round_trip(self):
        corrections = [{"from": "teh", "to": "the"}]
        terms = ["Frobnicator", 'a 27" display']
        text = transcripts.dictionary_to_csv(corrections, terms)
        assert transcripts.parse_dictionary_csv(text) == (corrections, terms)

    def test_empty_dictionary_still_has_a_header(self):
        text = transcripts.dictionary_to_csv([], [])
        assert text.startswith("type,from,to")
        assert transcripts.parse_dictionary_csv(text) == ([], [])

    def test_export_skips_half_filled_corrections(self):
        text = transcripts.dictionary_to_csv([{"from": "teh", "to": ""}], [])
        assert "teh" not in text

    def test_parse_skips_unknown_types_and_blank_rows(self):
        text = "type,from,to\nnonsense,a,b\ncorrection,,x\nterm,,\ncorrection,teh,the\n"
        assert transcripts.parse_dictionary_csv(text) == (
            [{"from": "teh", "to": "the"}],
            [],
        )

    def test_term_may_sit_in_either_column(self):
        text = "type,from,to\nterm,Frobnicator,\n"
        assert transcripts.parse_dictionary_csv(text) == ([], ["Frobnicator"])

    def test_parse_trims_whitespace(self):
        text = "type,from,to\ncorrection,  teh , the  \nterm,,  Widget \n"
        corrections, terms = transcripts.parse_dictionary_csv(text)
        assert corrections == [{"from": "teh", "to": "the"}]
        assert terms == ["Widget"]

    def test_garbage_input_yields_nothing(self):
        assert transcripts.parse_dictionary_csv("not a csv at all") == ([], [])


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
        assert m == {
            "source_transcripts": {},
            "corrections": [],
            "marks": [],
            "known_terms": [],
        }

    def test_load_missing_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        m = transcripts.load_transcripts_manifest()
        assert m == {
            "source_transcripts": {},
            "corrections": [],
            "marks": [],
            "known_terms": [],
        }

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
        assert not (tmp_path / config.MANIFEST_FILENAME).exists()

    def test_known_terms_round_trip(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        transcripts.save_transcripts_manifest({}, [], known_terms=["Frobnicator"])
        loaded = transcripts.load_transcripts_manifest()
        assert loaded["known_terms"] == ["Frobnicator"]

    def test_terms_only_manifest_persists(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        path = transcripts.save_transcripts_manifest({}, [], known_terms=["Widget"])
        assert path is not None

    def test_known_terms_preserved_when_none(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        transcripts.save_transcripts_manifest({}, [], known_terms=["Widget"])
        transcripts.save_transcripts_manifest(
            {}, [{"id": "c1", "from": "a", "to": "b"}]
        )
        loaded = transcripts.load_transcripts_manifest()
        assert loaded["known_terms"] == ["Widget"]

    def test_emptying_existing_manifest_removes_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        manifest = tmp_path / config.MANIFEST_FILENAME
        transcripts.save_transcripts_manifest(
            {}, [{"id": "c1", "from": "a", "to": "b", "created": "2025-01-01T00:00:00"}]
        )
        assert manifest.is_file()
        # Pass marks=[] explicitly to defeat the marks-from-disk preservation.
        transcripts.save_transcripts_manifest({}, [], marks=[])
        assert not manifest.exists()

    def test_load_corrupt_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        (tmp_path / config.MANIFEST_FILENAME).write_text("not json")
        m = transcripts.load_transcripts_manifest()
        assert m == {
            "source_transcripts": {},
            "corrections": [],
            "marks": [],
            "known_terms": [],
        }

    def test_load_returns_independent_copies(self, tmp_path, monkeypatch):
        """Mutating a returned entry in place must not corrupt the cache."""
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        utils._reset_manifest_cache()
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
        utils._reset_manifest_cache()
        transcripts.save_transcripts_manifest(
            {}, [{"id": "c1", "from": "a", "to": "b", "created": "2025-01-01T00:00:00"}]
        )
        utils._reset_manifest_cache()

        calls = {"n": 0}
        real_read = utils._read_sections

        def _counting_read(*args, **kwargs):
            calls["n"] += 1
            return real_read(*args, **kwargs)

        monkeypatch.setattr(utils, "_read_sections", _counting_read)

        transcripts.load_transcripts_manifest()  # miss -> one disk read
        transcripts.load_transcripts_manifest()  # hit  -> no disk read
        assert calls["n"] == 1

    def test_save_busts_cache(self, tmp_path, monkeypatch):
        """After a save the next load must reflect the new data, not the cache."""
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        utils._reset_manifest_cache()
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

    def test_caches_scan_for_unchanged_dir(self, tmp_path, monkeypatch):
        """A second call on an unchanged dir is served from the mtime cache."""
        monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path))
        (tmp_path / "study_P01.mp4").touch()

        calls = {"n": 0}
        real_glob = Path.glob

        def counting_glob(self, pattern):
            calls["n"] += 1
            return real_glob(self, pattern)

        monkeypatch.setattr(Path, "glob", counting_glob)

        first = utils.discover_participant_videos()
        second = utils.discover_participant_videos()
        assert [p["id"] for p in first] == ["P01"]
        assert first == second
        assert calls["n"] == 1  # second call did not re-glob


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
        assert task["phase"] == "queued"
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
    def test_get_task(self):
        worker = transcripts.TranscriptWorker()
        task_id = worker.enqueue(transcripts.create_transcript_task("P01", ["/v.mp4"]))
        task = worker.get_task(task_id)
        assert task is not None
        assert task["id"] == task_id
        assert worker.get_task("nonexistent") is None

    def test_cancel_all_cancels_queued_and_flags_running(self):
        """cancel_all marks queued tasks cancelled and flags running ones."""
        worker = transcripts.TranscriptWorker()
        queued_id = worker.enqueue(
            transcripts.create_transcript_task("P01", ["/v.mp4"])
        )
        running_id = worker.enqueue(
            transcripts.create_transcript_task("P02", ["/v.mp4"])
        )
        with worker._lock:
            worker._tasks[running_id]["status"] = transcripts.TASK_STATUS_RUNNING
        worker.cancel_all()
        queued = worker.get_task(queued_id)
        running = worker.get_task(running_id)
        assert queued is not None and running is not None
        assert queued["status"] == transcripts.TASK_STATUS_CANCELLED
        assert running["_cancelled"] is True

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
            video_mod, "decode_audio_pcm", lambda _path, _idx=0, **_kw: _FAKE_AUDIO
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
            def transcribe(self, path, **kwargs):
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

    def test_execute_task_phase_progression_in_debug(self, monkeypatch):
        """Debug mode skips the loading_model phase by design (stub results,
        no Whisper) and lands on transcribing with a start stamp."""
        import video as video_mod

        monkeypatch.setattr(config, "DEBUGGING", True)
        monkeypatch.setattr(video_mod, "timeline_or_none", lambda *_a: None)
        monkeypatch.setattr(
            video_mod,
            "probe_video_properties",
            lambda *_a, **_k: {"duration": 100.0, "audio_codec": "aac"},
        )
        monkeypatch.setattr(
            transcripts, "load_transcripts_manifest", lambda: {"corrections": []}
        )
        monkeypatch.setattr(transcripts, "_resolve_audio_index", lambda *_a: 0)

        worker = transcripts.TranscriptWorker()
        task = transcripts.create_transcript_task("P01", ["/v.mp4"])
        task["status"] = "running"
        assert task["phase"] == "queued"

        worker._execute_task(task)

        assert task["status"] == "completed"
        assert task["phase"] == "transcribing"
        assert task["transcribe_started_at"]

    def test_execute_task_model_load_failure_fails_task(self, monkeypatch):
        """A model that fails to construct fails the task under the
        loading_model phase instead of dying inside transcribe_video."""
        from unittest.mock import Mock

        import video as video_mod

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(video_mod, "timeline_or_none", lambda *_a: None)
        monkeypatch.setattr(
            video_mod,
            "probe_video_properties",
            lambda *_a, **_k: {"duration": 100.0, "audio_codec": "aac"},
        )
        monkeypatch.setattr(
            transcripts, "load_transcripts_manifest", lambda: {"corrections": []}
        )
        monkeypatch.setattr(transcripts, "_resolve_audio_index", lambda *_a: 0)
        load_model = Mock(return_value=None)
        monkeypatch.setattr(transcripts, "_load_model", load_model)

        worker = transcripts.TranscriptWorker()
        task = transcripts.create_transcript_task("P01", ["/v.mp4"])
        task["status"] = "running"

        worker._execute_task(task)

        assert task["status"] == "failed"
        assert "model failed to load" in task["error"]
        assert task["phase"] == "loading_model"
        load_model.assert_called_once()

    def test_execute_task_model_load_raising_fails_task(self, monkeypatch):
        """_load_model raises as well as returning None (run_with_spinner calls
        its callback bare, and the WhisperModel construction has no except), so
        the preload must catch it rather than let it escape _execute_task."""
        import video as video_mod

        monkeypatch.setattr(config, "DEBUGGING", False)
        monkeypatch.setattr(video_mod, "timeline_or_none", lambda *_a: None)
        monkeypatch.setattr(
            video_mod,
            "probe_video_properties",
            lambda *_a, **_k: {"duration": 100.0, "audio_codec": "aac"},
        )
        monkeypatch.setattr(
            transcripts, "load_transcripts_manifest", lambda: {"corrections": []}
        )
        monkeypatch.setattr(transcripts, "_resolve_audio_index", lambda *_a: 0)

        def _raising_load(*_a, **_k):
            raise RuntimeError("Library cublas64_12.dll is not found")

        monkeypatch.setattr(transcripts, "_load_model", _raising_load)

        worker = transcripts.TranscriptWorker()
        task = transcripts.create_transcript_task("P01", ["/v.mp4"])
        task["status"] = "running"

        worker._execute_task(task)

        assert task["status"] == "failed"
        assert "cublas64_12" in task["error"]
        assert task["completed_at"]

    def test_worker_loop_survives_a_raising_task(self, monkeypatch):
        """A task that raises anywhere outside _execute_task's own try must not
        kill the worker thread: there is no restart path, so the next task
        would queue into a loop nothing drains."""
        import time

        monkeypatch.setattr(config, "DEBUGGING", False)

        seen = []

        # Keyed on the participant, not execution order: both tasks enqueue at
        # the same priority and the tie breaks on a random task id, so which
        # one the worker picks up first is not deterministic.
        def _boom(task):
            seen.append(task["id"])
            if task["participant"] == "P01":
                raise RuntimeError("model exploded")
            task["status"] = "completed"

        worker = transcripts.TranscriptWorker()
        monkeypatch.setattr(worker, "_execute_task", _boom)
        worker.start()
        try:
            first = worker.enqueue(
                transcripts.create_transcript_task("P01", ["/v.mp4"])
            )
            second = worker.enqueue(
                transcripts.create_transcript_task("P02", ["/v.mp4"])
            )
            deadline = time.monotonic() + 5
            while len(seen) < 2 and time.monotonic() < deadline:
                time.sleep(0.02)
        finally:
            worker.stop()

        assert len(seen) == 2, "worker died on the first task"
        failed = worker.get_task(first)
        completed = worker.get_task(second)
        assert failed is not None and completed is not None
        assert failed["status"] == "failed"
        assert "model exploded" in failed["error"]
        assert completed["status"] == "completed"

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

        import sys
        import types
        from typing import Any

        fake_fw: Any = types.ModuleType("faster_whisper")
        fake_fw.WhisperModel = FakeModel
        monkeypatch.setitem(sys.modules, "faster_whisper", fake_fw)
        monkeypatch.setattr(
            transcripts, "is_whisper_model_cached", lambda *_a, **_k: True
        )
        monkeypatch.setattr(transcripts, "_cached_model", None)
        monkeypatch.setattr(transcripts, "_cached_model_name", None)
        monkeypatch.setattr(transcripts, "_cached_model_key", None)
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

    def test_changing_the_device_reloads_the_cached_model(self, monkeypatch):
        """TRANSCRIBE_DEVICE is a user-editable Studio setting, but device is
        read at WhisperModel() time — keying the cache on the model name alone
        meant a cpu→cuda change saved, persisted and displayed while every
        later transcription silently kept running the model loaded for cpu."""
        seen = self._capture_model(monkeypatch)
        loads: list[str] = []
        monkeypatch.setattr(config, "TRANSCRIBE_DEVICE", "cpu")
        transcripts._load_model("base")
        loads.append(seen["kwargs"]["device"])

        # Same name, different device → must construct again.
        monkeypatch.setattr(config, "TRANSCRIBE_DEVICE", "cuda")
        assert transcripts.is_transcription_model_loaded() is False
        transcripts._load_model("base")
        loads.append(seen["kwargs"]["device"])
        assert loads == ["cpu", "cuda"]

        # ...and an unchanged signature must still hit the cache.
        seen.clear()
        transcripts._load_model("base")
        assert seen == {}, "an unchanged load signature must not reconstruct"

    def test_changing_the_thread_count_reloads_the_cached_model(self, monkeypatch):
        """Same contract for the other construction-time settings."""
        seen = self._capture_model(monkeypatch)
        monkeypatch.setattr(config, "TRANSCRIBE_CPU_THREADS", 4)
        transcripts._load_model("base")
        assert seen["kwargs"]["cpu_threads"] == 4
        monkeypatch.setattr(config, "TRANSCRIBE_CPU_THREADS", 8)
        transcripts._load_model("base")
        assert seen["kwargs"]["cpu_threads"] == 8

    def test_download_gate_stays_keyed_on_the_name(self, monkeypatch):
        """is_whisper_model_cached answers "is this model downloaded", which
        device and thread count do not affect — it must not start reporting
        False (and re-prompting for a download) after a device change."""
        self._capture_model(monkeypatch)
        monkeypatch.setattr(config, "TRANSCRIBE_DEVICE", "cpu")
        transcripts._load_model("base")
        assert transcripts.is_whisper_model_cached("base") is True
        monkeypatch.setattr(config, "TRANSCRIBE_DEVICE", "cuda")
        assert transcripts.is_whisper_model_cached("base") is True

    def test_device_is_always_passed_explicitly(self, monkeypatch):
        """Never leave it to faster-whisper's ``device="auto"`` default."""
        seen = self._capture_model(monkeypatch)
        monkeypatch.setattr(config, "TRANSCRIBE_DEVICE", "auto")
        transcripts._load_model("base")
        assert seen["kwargs"].get("device") in ("auto", "cpu")


class TestResolveTranscribeDevice:
    """The frozen bundle ships no CUDA runtime, so an auto-selected GPU can
    only fail at first inference with ``Library cublas64_12.dll is not found
    or cannot be loaded``."""

    def test_auto_is_cpu_when_frozen(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_DEVICE", "auto")
        monkeypatch.setattr(transcripts.sys, "frozen", True, raising=False)
        assert transcripts._resolve_transcribe_device() == "cpu"

    def test_auto_stays_auto_from_source(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_DEVICE", "auto")
        monkeypatch.delattr(transcripts.sys, "frozen", raising=False)
        assert transcripts._resolve_transcribe_device() == "auto"

    def test_explicit_cuda_survives_freezing(self, monkeypatch):
        """A user who installed cuBLAS/cuDNN themselves gets what they asked
        for — including CTranslate2's own error if they were wrong."""
        monkeypatch.setattr(config, "TRANSCRIBE_DEVICE", "cuda")
        monkeypatch.setattr(transcripts.sys, "frozen", True, raising=False)
        assert transcripts._resolve_transcribe_device() == "cuda"

    def test_case_and_whitespace_tolerated(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_DEVICE", "  CPU ")
        assert transcripts._resolve_transcribe_device() == "cpu"

    def test_unknown_value_falls_back_to_auto(self, monkeypatch):
        monkeypatch.setattr(config, "TRANSCRIBE_DEVICE", "mps")
        monkeypatch.delattr(transcripts.sys, "frozen", raising=False)
        assert transcripts._resolve_transcribe_device() == "auto"


class TestIsWhisperModelCached:
    def test_debugging_short_circuits_true(self, monkeypatch):
        monkeypatch.setattr(config, "DEBUGGING", True)
        assert transcripts.is_whisper_model_cached("large-v3") is True

    def test_true_when_snapshot_resolves(self, monkeypatch):
        monkeypatch.setattr(config, "DEBUGGING", False)
        import huggingface_hub

        monkeypatch.setattr(
            huggingface_hub, "snapshot_download", lambda *a, **k: "/cache/path"
        )
        assert transcripts.is_whisper_model_cached("base") is True

    def test_false_when_snapshot_raises(self, monkeypatch):
        monkeypatch.setattr(config, "DEBUGGING", False)
        import huggingface_hub

        def _boom(*a, **k):
            raise FileNotFoundError("not in cache")

        monkeypatch.setattr(huggingface_hub, "snapshot_download", _boom)
        assert transcripts.is_whisper_model_cached("large-v3") is False

    def test_passes_local_files_only_and_repo_id(self, monkeypatch):
        monkeypatch.setattr(config, "DEBUGGING", False)
        import huggingface_hub

        seen: dict = {}

        def _capture(repo_id, **kwargs):
            seen["repo_id"] = repo_id
            seen["local_files_only"] = kwargs.get("local_files_only")
            return "/cache/path"

        monkeypatch.setattr(huggingface_hub, "snapshot_download", _capture)
        transcripts.is_whisper_model_cached("medium")
        assert seen["repo_id"] == "Systran/faster-whisper-medium"
        assert seen["local_files_only"] is True

    def test_full_repo_id_is_passed_through(self, monkeypatch):
        """download_model's rule: anything with a '/' is already a repo id."""
        monkeypatch.setattr(config, "DEBUGGING", False)
        import huggingface_hub

        seen: dict = {}

        def _capture(repo_id, **kwargs):
            seen["repo_id"] = repo_id
            return "/cache/path"

        monkeypatch.setattr(huggingface_hub, "snapshot_download", _capture)
        transcripts.is_whisper_model_cached("mobiuslabs/faster-whisper-large-v3-turbo")
        assert seen["repo_id"] == "mobiuslabs/faster-whisper-large-v3-turbo"

    def test_mirror_stays_pinned_to_faster_whisper(self):
        """The hf-direct check must keep matching faster_whisper's own logic.

        is_whisper_model_cached deliberately avoids importing faster_whisper
        (~600 ms package import on the /api/models request path) and mirrors
        download_model instead: repo id from the Systran prefix, cache probe
        via snapshot_download with the same allow_patterns. This test pays the
        import once to pin both halves of that mirror, so a faster-whisper
        upgrade that remaps a curated model or fetches different files fails
        here instead of silently mis-gating downloads.
        """
        import inspect
        import re

        # av is deliberately not installed; the package import needs the stub.
        transcripts._ensure_av_stub()
        import faster_whisper.utils as fwu

        for m in transcripts.WHISPER_MODELS:
            assert fwu._MODELS.get(m["name"]) == (
                transcripts._WHISPER_REPO_PREFIX + m["name"]
            ), f"faster_whisper remapped '{m['name']}'"
        source = inspect.getsource(fwu.download_model)
        match = re.search(r"allow_patterns = \[(.*?)\]", source, re.DOTALL)
        assert match, "download_model no longer builds a literal allow_patterns"
        theirs = set(re.findall(r'"([^"]+)"', match.group(1)))
        assert theirs == set(transcripts._WHISPER_ALLOW_PATTERNS)


def test_faster_whisper_imports_with_av_stub_only():
    """The PyAV-free import contract, pinned from both ends.

    PyAV is overridden out of the dependency tree (pyproject.toml): its wheel
    bundles a second ~40 MB FFmpeg, and faster-whisper only calls ``av`` to
    decode *path* inputs while clipgen always passes ffmpeg-decoded ndarrays.
    Two ways this can silently rot, both caught here: the real ``av``
    distribution reappearing (reinstating it must be a deliberate change —
    it drags licensing sections of build/THIRD-PARTY-LICENSES and the
    clipgen.spec exclude back with it), and a faster-whisper upgrade that
    starts *executing* ``av`` at import time, which the empty stub module
    cannot satisfy.
    """
    import importlib.metadata

    with pytest.raises(importlib.metadata.PackageNotFoundError):
        importlib.metadata.distribution("av")

    transcripts._ensure_av_stub()
    import faster_whisper

    assert hasattr(faster_whisper, "WhisperModel")


class TestConfirmModelDownload:
    """The CLI consent gate in front of a first Whisper download."""

    @staticmethod
    def _uncached(monkeypatch, *, interactive):
        monkeypatch.setattr(
            transcripts, "is_whisper_model_cached", lambda n=None: False
        )
        monkeypatch.setattr(transcripts, "_stdin_is_interactive", lambda: interactive)
        monkeypatch.setattr(utils, "NO_INPUT_MODE", False)

    def test_cached_model_never_prompts(self, monkeypatch):
        monkeypatch.setattr(transcripts, "is_whisper_model_cached", lambda n=None: True)
        monkeypatch.setattr(
            utils, "read_user_input", lambda _p: pytest.fail("should not prompt")
        )
        assert transcripts._confirm_model_download("large-v3") is True

    def test_non_interactive_stdin_proceeds_without_prompting(self, monkeypatch):
        """A piped or closed stdin has nothing to answer with, and input()
        raises there — a scripted --transcribe must download, not crash. This is
        also how CI hits it: pytest's capture stub reports isatty() False."""
        self._uncached(monkeypatch, interactive=False)
        monkeypatch.setattr(
            utils, "read_user_input", lambda _p: pytest.fail("should not prompt")
        )
        assert transcripts._confirm_model_download("base") is True

    def test_no_input_mode_proceeds_without_prompting(self, monkeypatch):
        self._uncached(monkeypatch, interactive=True)
        monkeypatch.setattr(utils, "NO_INPUT_MODE", True)
        monkeypatch.setattr(
            utils, "read_user_input", lambda _p: pytest.fail("should not prompt")
        )
        assert transcripts._confirm_model_download("base") is True

    @pytest.mark.parametrize(
        "answer,expected", [("y", True), ("yes", True), ("n", False), ("", False)]
    )
    def test_interactive_answer_decides(self, monkeypatch, answer, expected):
        self._uncached(monkeypatch, interactive=True)
        monkeypatch.setattr(utils, "read_user_input", lambda _p: answer)
        assert transcripts._confirm_model_download("large-v3") is expected

    def test_declining_aborts_the_load(self, monkeypatch):
        """_load_model returns None rather than downloading behind the spinner."""
        self._uncached(monkeypatch, interactive=True)
        monkeypatch.setattr(utils, "read_user_input", lambda _p: "n")
        monkeypatch.setattr(transcripts, "_cached_model", None)
        monkeypatch.setattr(transcripts, "_cached_model_name", None)
        transcripts._ensure_av_stub()  # av is deliberately not installed
        import faster_whisper

        monkeypatch.setattr(
            faster_whisper,
            "WhisperModel",
            lambda *a, **k: pytest.fail("must not load after a decline"),
        )
        assert transcripts._load_model("large-v3") is None


class TestApplyCorrectionsBoundaries:
    def test_word_boundary_does_not_rewrite_substrings(self):
        """ "the" -> "they" must not turn "there" into "theyre"."""
        segs = [transcripts.TranscriptSegment(start=0, end=1, text="the cat sat there")]
        result = transcripts.apply_corrections(segs, [{"from": "the", "to": "they"}])
        assert result[0]["text"] == "they cat sat there"

    def test_punctuation_edges_still_match(self):
        """\\b against a punctuation edge never matches, so it is only applied
        to word-character edges."""
        segs = [transcripts.TranscriptSegment(start=0, end=1, text="use e.g. this")]
        result = transcripts.apply_corrections(
            segs, [{"from": "e.g.", "to": "for example"}]
        )
        assert result[0]["text"] == "use for example this"

    def test_replacement_is_literal_not_template(self):
        """A backslash in the replacement is text, not a re.sub escape —
        previously this raised re.error and 500'd every corrected route."""
        segs = [transcripts.TranscriptSegment(start=0, end=1, text="the path")]
        result = transcripts.apply_corrections(
            segs, [{"from": "path", "to": "C:\\Users\\path \\g<0>"}]
        )
        assert result[0]["text"] == "the C:\\Users\\path \\g<0>"
