"""Unit tests for transcribe_bench's row reduction and validity checks."""

import json

import pytest

import config
import profiling
import transcribe_bench


@pytest.fixture
def make_doc(monkeypatch):
    """A real export document seeded with labels; no git/ffmpeg probes."""
    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(profiling, "environment", lambda: {"python": "3.12"})

    def build(seed):
        monkeypatch.setattr(config, "PROFILE_DEEP", "")
        profiling.reset()
        for label, (secs, n) in seed.items():
            profiling.add(label, secs, n)
        return profiling.export(reset=True)

    yield build
    profiling.reset()


SEED = {
    "transcribe.prepare": (2.0, 1),
    "transcribe.decode": (8.0, 42),
    "transcribe.model_load": (0.4, 1),
    "transcribe.decode_audio": (0.2, 1),
    "speakers.diarize": (1.5, 1),
    "speakers.cluster": (0.1, 1),
}


def test_summarize_reads_whisper_and_speaker_labels(make_doc):
    row = transcribe_bench.summarize(make_doc(SEED))
    assert row["whisper_s"] == pytest.approx(10.0)
    assert row["segments"] == 42
    assert row["speakers_s"] == pytest.approx(1.5)
    assert row["speakers_cluster_s"] == pytest.approx(0.1)


def _store(out_dir, segments):
    entry = {"segments": segments}
    doc = {"transcripts": {"source_transcripts": {"P01": entry}}}
    (out_dir / "clipgen.json").write_text(json.dumps(doc))


def test_check_work_needs_segments_and_a_stored_transcript(tmp_path, make_doc):
    assert transcribe_bench.check_work("whisper", make_doc({}), tmp_path) == [
        "no segments decoded (VAD found no speech?)",
        "no transcript stored for P01",
    ]
    _store(tmp_path, [{"start": 0, "end": 1, "text": "hi"}])
    assert transcribe_bench.check_work("whisper", make_doc(SEED), tmp_path) == []


def test_speakers_scenario_requires_the_speaker_pass(tmp_path, make_doc):
    _store(tmp_path, [{"start": 0, "end": 1, "text": "hi"}])
    no_pass = {k: v for k, v in SEED.items() if not k.startswith("speakers")}
    assert transcribe_bench.check_work("speakers", make_doc(no_pass), tmp_path) == [
        "speaker pass did not run (model missing?)"
    ]


def test_fixture_spec_records_speech_audio():
    spec = transcribe_bench.fixture_spec(300)
    assert spec["audio"] == "speech" and spec["has_audio"] is True
