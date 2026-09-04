"""speakers.py: fbank shape, clustering, label assignment, multi-part mapping."""

import numpy as np
import pytest

import config
import speakers
import video


# ---- Features --------------------------------------------------------------


def test_fbank_frame_count_follows_kaldi_snip_edges_false():
    rng = np.random.default_rng(0)
    one_second = rng.standard_normal(16000).astype(np.float32)
    feats = speakers._fbank(one_second)
    assert feats.shape == (100, speakers.N_MELS)
    assert feats.dtype == np.float32
    assert np.isfinite(feats).all()
    assert speakers._fbank(one_second[:8000]).shape == (50, speakers.N_MELS)
    assert speakers._fbank(one_second[:100]).shape == (1, speakers.N_MELS)
    assert speakers._fbank(np.zeros(0, dtype=np.float32)).shape == (0, speakers.N_MELS)


def test_fbank_tone_peaks_in_a_stable_low_bin():
    t = np.arange(16000) / 16000
    tone = (0.5 * np.sin(2 * np.pi * 440 * t)).astype(np.float32)
    feats = speakers._fbank(tone)
    peak = int(np.argmax(feats.mean(axis=0)))
    assert 5 <= peak <= 25
    assert int(np.argmax(speakers._fbank(tone).mean(axis=0))) == peak


def test_mel_filterbank_shape_and_nyquist_column_is_silent():
    bank = speakers._mel_filterbank()
    assert bank.shape == (speakers.N_FFT // 2 + 1, speakers.N_MELS)
    assert bank[-1].sum() == 0.0
    assert (bank.sum(axis=0) > 0).all()


# ---- Clustering ------------------------------------------------------------


def _centroids(n_dims: int = 16):
    return np.eye(n_dims, dtype=np.float32)[:3]


def _jittered(centroid, count, seed):
    rng = np.random.default_rng(seed)
    pts = centroid + 0.05 * rng.standard_normal((count, centroid.shape[0]))
    return pts / np.linalg.norm(pts, axis=1, keepdims=True)


def test_cluster_separates_three_speakers_numbered_by_first_appearance():
    a, b, c = _centroids()
    pts = np.concatenate([_jittered(b, 4, 1), _jittered(a, 4, 2), _jittered(c, 4, 3)])
    labels = speakers.cluster(pts, max_speakers=8)
    assert labels[:4] == [0] * 4
    assert labels[4:8] == [1] * 4
    assert labels[8:] == [2] * 4


def test_cluster_respects_max_speakers():
    a, b, c = _centroids()
    pts = np.concatenate([_jittered(a, 3, 1), _jittered(b, 3, 2), _jittered(c, 3, 3)])
    labels = speakers.cluster(pts, max_speakers=2)
    assert len(set(labels)) == 2


def test_cluster_identical_vectors_collapse_to_one():
    pts = np.tile(_centroids()[0], (5, 1))
    assert speakers.cluster(pts, max_speakers=8) == [0] * 5


def test_cluster_zero_threshold_keeps_every_point_apart():
    pts = _jittered(_centroids()[0], 4, 7)
    assert speakers.cluster(pts, max_speakers=8, threshold=0.0) == [0, 1, 2, 3]


def test_cluster_empty():
    assert speakers.cluster(np.zeros((0, 4), dtype=np.float32), 4) == []


# ---- assign_speakers ---------------------------------------------------------


def _segments(*spans):
    return [{"start": s, "end": e, "text": "x"} for s, e in spans]


def _parity_embed(monkeypatch):
    """Even-indexed calls embed as speaker A, odd as B; records every call."""
    calls: list[int] = []

    def fake_embed(samples):
        calls.append(len(samples))
        vec = np.zeros(8, dtype=np.float32)
        vec[len(calls) % 2] = 1.0
        return vec

    monkeypatch.setattr(speakers, "embed", fake_embed)
    return calls


def test_assign_speakers_labels_and_short_segments_inherit(monkeypatch):
    calls = _parity_embed(monkeypatch)
    segs = _segments((0, 2), (2, 4), (4, 4.2), (4.2, 6), (6, 8))
    audio = np.ones(16000 * 10, dtype=np.float32)

    def audio_for(seg):
        return audio[int(seg["start"] * 16000) : int(seg["end"] * 16000)]

    progress: list[float] = []
    speakers.assign_speakers(segs, audio_for, 4, on_progress=progress.append)
    assert len(calls) == 4
    labels = [s["speaker"] for s in segs]
    assert labels[0] != labels[1]
    assert labels[1] != labels[3]
    # 4.0–4.2 sits nearer the 4.2–6 segment's midpoint than the 2–4 one.
    assert labels[2] == labels[3]
    assert set(labels) == {"1", "2"}
    assert progress[-1] == 1.0


def test_assign_speakers_skips_missing_audio_and_clears_stale_labels(monkeypatch):
    _parity_embed(monkeypatch)
    segs = _segments((0, 2), (2, 4))
    segs[0]["speaker"] = "9"
    speakers.assign_speakers(segs, lambda seg: None, 4)
    assert all("speaker" not in s for s in segs)


def test_assign_speakers_cancel_raises(monkeypatch):
    _parity_embed(monkeypatch)
    segs = _segments((0, 2), (2, 4))
    with pytest.raises(speakers.DiarizationCancelled):
        speakers.assign_speakers(
            segs,
            lambda seg: np.ones(32000, dtype=np.float32),
            4,
            cancel_flag=lambda: True,
        )


def test_assign_speakers_pools_adjacent_runs_before_clustering(monkeypatch):
    monkeypatch.setattr(
        speakers, "embed", lambda samples: np.array([1.0, 0, 0, 0], dtype=np.float32)
    )
    seen: list[int] = []

    def spy_cluster(embeddings, max_speakers, threshold=speakers.CLUSTER_THRESHOLD):
        seen.append(len(embeddings))
        return [0] * len(embeddings)

    monkeypatch.setattr(speakers, "cluster", spy_cluster)
    segs = _segments((0, 2), (2, 4), (4, 6))
    speakers.assign_speakers(segs, lambda seg: np.ones(32000, dtype=np.float32), 4)
    assert seen == [1]
    assert [s["speaker"] for s in segs] == ["1", "1", "1"]


def test_absorb_small_clusters_folds_singletons_into_nearest():
    a, b, _c = _centroids(4)
    vectors = np.stack([a, a, b, b, (a * 0.9 + b * 0.1)])
    vectors = vectors / np.linalg.norm(vectors, axis=1, keepdims=True)
    labels = [0, 0, 1, 1, 2]
    weights = [10.0, 10.0, 10.0, 10.0, 1.0]
    assert speakers._absorb_small_clusters(labels, vectors, weights) == [0, 0, 1, 1, 0]
    # Nothing to absorb when every cluster carries enough speech.
    assert speakers._absorb_small_clusters(labels, vectors, [10.0] * 5) == labels
    # Everything tiny: leave as is rather than collapse to nothing.
    assert speakers._absorb_small_clusters(labels, vectors, [1.0] * 5) == labels


# ---- diarize_entry -----------------------------------------------------------


def test_diarize_entry_debugging_stub_alternates_two_labels(monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True)
    decoded: list = []
    monkeypatch.setattr(
        video, "decode_audio_pcm", lambda *a, **k: decoded.append(a) or None
    )
    segs = _segments((0, 1), (1, 2), (2, 3))
    block = speakers.diarize_entry(["/v.mp4"], segs, 0, None, max_speakers=4)
    assert [s["speaker"] for s in segs] == ["1", "2", "1"]
    assert block is not None and block["count"] == 2 and block["enabled"] is True
    assert block["labels"] == {}
    assert decoded == []


def test_diarize_entry_decodes_each_part_once_and_clusters_once(monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", False)
    monkeypatch.setattr(speakers, "is_speaker_model_available", lambda: True)
    decodes: list[tuple] = []

    def fake_decode(path, audio_index, *, start_seconds=None, duration_seconds=None):
        decodes.append((path, start_seconds, duration_seconds))
        length = int((duration_seconds or 10) * 16000)
        return np.arange(length, dtype=np.float32)

    monkeypatch.setattr(video, "decode_audio_pcm", fake_decode)
    windows: list[tuple[float, float]] = []

    def fake_embed(samples):
        windows.append((float(samples[0]), len(samples)))
        vec = np.zeros(4, dtype=np.float32)
        vec[len(windows) - 1] = 1.0  # distinct per call so nothing pools
        return vec

    monkeypatch.setattr(speakers, "embed", fake_embed)
    cluster_calls: list[int] = []

    def fake_cluster(embeddings, max_speakers, threshold=speakers.CLUSTER_THRESHOLD):
        cluster_calls.append(len(embeddings))
        return list(range(len(embeddings)))

    monkeypatch.setattr(speakers, "cluster", fake_cluster)
    timeline = [("/a.mp4", 100, 0), ("/b.mp4", 100, 100)]
    segs = _segments((10, 12), (130, 133), (150, 152))
    block = speakers.diarize_entry(
        ["/a.mp4", "/b.mp4"], segs, 1, timeline, max_speakers=4
    )
    assert [d[0] for d in decodes] == ["/a.mp4", "/b.mp4"]
    assert decodes[0][1] == 10 and decodes[0][2] == pytest.approx(2)
    assert decodes[1][1] == 30 and decodes[1][2] == pytest.approx(22)
    # Global 130 s is local 30 s: offset 0 in the window starting at 30.
    assert windows[1] == (0.0, 3 * 16000)
    assert windows[2] == (20 * 16000, 2 * 16000)
    assert cluster_calls == [3]
    assert block is not None and block["count"] == 3


def test_diarize_entry_returns_none_and_strips_when_decode_fails(monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", False)
    monkeypatch.setattr(speakers, "is_speaker_model_available", lambda: True)
    monkeypatch.setattr(video, "decode_audio_pcm", lambda *a, **k: None)
    segs = _segments((0, 2), (2, 4))
    segs[0]["speaker"] = "1"
    assert speakers.diarize_entry(["/v.mp4"], segs, 0, None, max_speakers=4) is None
    assert all("speaker" not in s for s in segs)


def test_diarize_entry_without_model_returns_none(monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", False)
    monkeypatch.setattr(speakers, "is_speaker_model_available", lambda: False)
    segs = _segments((0, 2))
    assert speakers.diarize_entry(["/v.mp4"], segs, 0, None, max_speakers=4) is None


def test_speaker_display_name_prefers_rename():
    assert speakers.speaker_display_name("1", {"1": "Moderator"}) == "Moderator"
    assert speakers.speaker_display_name("2", {"1": "Moderator"}) == "Speaker 2"
    assert speakers.speaker_display_name("2", None) == "Speaker 2"
    assert speakers.speaker_display_name("2", {"2": ""}) == "Speaker 2"


# ---- Real model smoke (only where the vendored file exists) ------------------


@pytest.mark.skipif(
    speakers.vendored_model_path() is None, reason="speaker model not vendored"
)
def test_vendored_model_embeds_unit_vectors():
    rng = np.random.default_rng(3)
    noise = (0.1 * rng.standard_normal(32000)).astype(np.float32)
    vec = speakers.embed(noise)
    assert vec.ndim == 1 and vec.shape[0] > 64
    assert np.isfinite(vec).all()
    assert np.linalg.norm(vec) == pytest.approx(1.0, abs=1e-4)
    _sess, meta = speakers._load_session()
    assert meta.get("sample_rate", "16000") == "16000"
    same = speakers.embed(noise)
    assert float(vec @ same) == pytest.approx(1.0, abs=1e-4)


def test_remap_speaker_ids_keeps_old_ids_by_overlap():
    previous = {"a": "1", "b": "1", "c": "2", "d": "2", "e": "2"}
    # The fresh run swapped the ids and split one line into a third voice.
    fresh = {"a": "2", "b": "2", "c": "1", "d": "1", "e": "3"}
    assert speakers.remap_speaker_ids(previous, fresh) == {"2": "1", "1": "2", "3": "3"}


def test_remap_speaker_ids_without_history_is_identity_and_fills_gaps():
    assert speakers.remap_speaker_ids({}, {"a": "1", "b": "2"}) == {"1": "1", "2": "2"}
    # Old id 1 vanished; the unmatched fresh id takes the lowest free slot.
    assert speakers.remap_speaker_ids({"a": "2"}, {"a": "1", "b": "2"}) == {
        "1": "2",
        "2": "1",
    }
