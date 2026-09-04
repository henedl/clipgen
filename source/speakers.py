"""Speaker attribution for transcripts: who said each Whisper segment.

A standalone pass over an existing transcript, not part of ``transcribe_video``:
``diarize_entry`` re-decodes the audio, embeds every segment with the bundled
WeSpeaker CAM++ ONNX model (``build/fetch_binaries.py`` ``SPEAKER_MODEL_PINS``),
clusters the embeddings, and writes ``segment["speaker"] = "1".."N"``. That
lets the worker run it right after transcription *and* the Transcripts page
run it on demand for a participant that was transcribed long ago.

The whole pipeline is numpy + onnxruntime, both already shipped: no torch,
no scipy. Features are Kaldi ``compute-fbank-feats`` reimplemented in numpy
(``_fbank``), matching what sherpa-onnx feeds these models: 16 kHz, 25 ms
povey-windowed frames every 10 ms, ``snip_edges=false`` symmetric padding,
DC removal, 0.97 pre-emphasis, 512-point power spectrum, 80 unnormalised
triangular mel bins between 20 Hz and 7600 Hz, natural log with a float32
epsilon floor, no dither. The model's own metadata decides two things at
load time: ``normalize_samples`` ("0" means int16-range input, so the float
PCM is scaled by 32768) and ``feature_normalize_type`` ("global-mean"
subtracts the per-bin mean over time before inference).

Clustering is average-linkage agglomerative on cosine distance
(``cluster``), stopping at ``CLUSTER_THRESHOLD`` and never above
``max_speakers``. Speakers talk in runs, so ``assign_speakers`` first pools
consecutive near-identical segments (``ADJACENT_MERGE_DISTANCE``), which
takes an hour of speech from ~2000 segments to a few hundred units before
the N² matrix is built. Segments too short to embed reliably
(``MIN_EMBED_SECONDS``) inherit the label of the nearest labelled segment.

Module-level imports stay light on purpose: ``thinking_agents`` and
``data_export`` import this for ``speaker_display_name`` and must not pay for
numpy or onnxruntime.
"""

from __future__ import annotations

import functools
import threading
from collections.abc import Callable
from datetime import UTC, datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any

import config
import profiling
import utils

if TYPE_CHECKING:
    import numpy as np

SAMPLE_RATE = 16000
FRAME_LEN = 400
FRAME_SHIFT = 160
N_FFT = 512
N_MELS = 80
PREEMPHASIS = 0.97
LOW_FREQ_HZ = 20.0
HIGH_FREQ_HZ = SAMPLE_RATE / 2 - 400.0
LOG_FLOOR = 1.1920929e-07
# Below this a segment gets no embedding; it inherits a neighbour's label.
MIN_EMBED_SECONDS = 0.4
# Consecutive segments closer than this pool into one clustering unit.
ADJACENT_MERGE_DISTANCE = 0.25
# Clusters further apart than this stay apart; 0.5 left singletons on real dialogue.
CLUSTER_THRESHOLD = 0.6
# Clusters with less speech than this fold into their nearest neighbour.
MIN_CLUSTER_SECONDS = 5.0
MODEL_FILENAME = "speaker_embed.onnx"

_session_lock = threading.Lock()
_session: tuple[Any, dict[str, str]] | None = None


class DiarizationCancelled(Exception):
    """Raised when the cancel flag trips mid-pass."""


def vendored_model_path() -> Path | None:
    """The bundled speaker model: frozen ``speaker_models/`` or ``build/vendor/speakers/``."""
    root = utils.get_bundled_assets_root()
    for base in (root / "speaker_models", root / "build" / "vendor" / "speakers"):
        onnx = base / MODEL_FILENAME
        if onnx.is_file():
            return onnx
    return None


def is_speaker_model_available() -> bool:
    return vendored_model_path() is not None


def speaker_display_name(speaker: str, labels: dict[str, str] | None) -> str:
    """User rename when present, else ``Speaker N``."""
    return (labels or {}).get(speaker) or f"Speaker {speaker}"


def speakers_block(count: int) -> dict[str, Any]:
    """Fresh manifest ``speakers`` block after a detection run."""
    return {
        "enabled": True,
        "labels": {},
        "count": count,
        "detected_at": datetime.now(UTC).isoformat(),
    }


# ---------------------------------------------------------------------------
# Features
# ---------------------------------------------------------------------------


@functools.cache
def _povey_window() -> np.ndarray:
    import numpy as np

    n = np.arange(FRAME_LEN, dtype=np.float64)
    return ((0.5 - 0.5 * np.cos(2 * np.pi * n / (FRAME_LEN - 1))) ** 0.85).astype(
        np.float32
    )


@functools.cache
def _mel_filterbank() -> np.ndarray:
    """Kaldi triangular mel bins as a ``(N_FFT // 2 + 1, N_MELS)`` matrix."""
    import numpy as np

    def mel(hz: float) -> float:
        return 1127.0 * np.log(1.0 + hz / 700.0)

    n_bins = N_FFT // 2 + 1
    fft_hz = np.arange(n_bins) * SAMPLE_RATE / N_FFT
    fft_mel = 1127.0 * np.log(1.0 + fft_hz / 700.0)
    mel_low, mel_high = mel(LOW_FREQ_HZ), mel(HIGH_FREQ_HZ)
    delta = (mel_high - mel_low) / (N_MELS + 1)
    bank = np.zeros((n_bins, N_MELS), dtype=np.float32)
    for m in range(N_MELS):
        left = mel_low + m * delta
        centre = left + delta
        right = centre + delta
        up = (fft_mel - left) / delta
        down = (right - fft_mel) / delta
        weight = np.where(fft_mel <= centre, up, down)
        bank[:, m] = np.clip(weight, 0.0, None)
    return bank


def _fbank(samples: np.ndarray) -> np.ndarray:
    """Kaldi fbank, ``snip_edges=false``: ``(num_frames, N_MELS)`` float32."""
    import numpy as np

    x = np.asarray(samples, dtype=np.float32)
    n = x.shape[0]
    num_frames = (n + FRAME_SHIFT // 2) // FRAME_SHIFT
    if num_frames <= 0:
        return np.zeros((0, N_MELS), dtype=np.float32)
    pad_left = (FRAME_LEN - FRAME_SHIFT) // 2
    last_end = (num_frames - 1) * FRAME_SHIFT - pad_left + FRAME_LEN
    pad_right = max(0, last_end - n)
    padded = np.pad(x, (pad_left, pad_right), mode="symmetric")
    frames = np.lib.stride_tricks.sliding_window_view(padded, FRAME_LEN)[::FRAME_SHIFT][
        :num_frames
    ]
    frames = frames - frames.mean(axis=1, keepdims=True)
    pre = frames.copy()
    pre[:, 1:] -= PREEMPHASIS * frames[:, :-1]
    pre[:, 0] -= PREEMPHASIS * frames[:, 0]
    pre *= _povey_window()
    power = np.abs(np.fft.rfft(pre, n=N_FFT)) ** 2
    mel = power @ _mel_filterbank()
    return np.log(np.maximum(mel, LOG_FLOOR)).astype(np.float32)


# ---------------------------------------------------------------------------
# Embeddings
# ---------------------------------------------------------------------------


def _load_session() -> tuple[Any, dict[str, str]]:
    """Cached onnxruntime session plus the model's metadata map."""
    global _session
    if _session is not None:
        return _session
    with _session_lock:
        if _session is not None:
            return _session
        path = vendored_model_path()
        if path is None:
            raise FileNotFoundError("Speaker model is not installed.")
        import onnxruntime as ort

        with profiling.span("speakers.session_load"):
            opts = ort.SessionOptions()
            opts.log_severity_level = 3
            sess = ort.InferenceSession(
                str(path), opts, providers=["CPUExecutionProvider"]
            )
        meta = dict(sess.get_modelmeta().custom_metadata_map)
        meta["_input"] = sess.get_inputs()[0].name
        if meta.get("sample_rate", str(SAMPLE_RATE)) != str(SAMPLE_RATE):
            raise ValueError(f"Speaker model expects {meta['sample_rate']} Hz audio.")
        _session = (sess, meta)
        return _session


def embed(samples: np.ndarray) -> np.ndarray:
    """Unit-length speaker embedding for float32 PCM in [-1, 1]."""
    import numpy as np

    sess, meta = _load_session()
    if meta.get("normalize_samples", "1") == "0":
        samples = np.asarray(samples, dtype=np.float32) * 32768.0
    feats = _fbank(samples)
    if meta.get("feature_normalize_type", "") == "global-mean":
        feats = feats - feats.mean(axis=0, keepdims=True)
    with profiling.span("speakers.embed"):
        out = sess.run(None, {meta["_input"]: feats[None]})[0][0]
    out = np.asarray(out, dtype=np.float32)
    return out / max(float(np.linalg.norm(out)), 1e-8)


# ---------------------------------------------------------------------------
# Clustering
# ---------------------------------------------------------------------------


def cluster(
    embeddings: np.ndarray, max_speakers: int, threshold: float = CLUSTER_THRESHOLD
) -> list[int]:
    """Average-linkage cosine clustering; labels numbered by first appearance."""
    import numpy as np

    n = len(embeddings)
    if n == 0:
        return []
    max_speakers = max(1, max_speakers)
    x = np.asarray(embeddings, dtype=np.float32)
    x = x / np.clip(np.linalg.norm(x, axis=1, keepdims=True), 1e-8, None)
    dist = (1.0 - x @ x.T).astype(np.float32)
    np.fill_diagonal(dist, np.inf)
    sizes = np.ones(n, dtype=np.float32)
    members: list[list[int]] = [[i] for i in range(n)]
    active = set(range(n))
    while len(active) > 1:
        i, j = divmod(int(np.argmin(dist)), n)
        if dist[i, j] > threshold and len(active) <= max_speakers:
            break
        merged = (sizes[i] * dist[i] + sizes[j] * dist[j]) / (sizes[i] + sizes[j])
        dist[i, :] = merged
        dist[:, i] = merged
        dist[i, i] = np.inf
        dist[j, :] = np.inf
        dist[:, j] = np.inf
        sizes[i] += sizes[j]
        members[i] += members[j]
        active.discard(j)
    labels = [0] * n
    for rank, root in enumerate(sorted(active, key=lambda r: min(members[r]))):
        for m in members[root]:
            labels[m] = rank
    return labels


def _merge_adjacent(
    indices: list[int], vectors: np.ndarray, durations: list[float]
) -> list[tuple[list[int], np.ndarray, float]]:
    """Pool consecutive near-identical embeddings into duration-weighted units."""
    import numpy as np

    units: list[tuple[list[int], np.ndarray, float]] = []
    for idx, vec, dur in zip(indices, vectors, durations, strict=True):
        if units:
            _members, mean, weight = units[-1]
            if 1.0 - float(mean @ vec) < ADJACENT_MERGE_DISTANCE:
                pooled = mean * weight + vec * dur
                pooled = pooled / max(float(np.linalg.norm(pooled)), 1e-8)
                units[-1] = (_members + [idx], pooled, weight + dur)
                continue
        units.append(([idx], vec, dur))
    return units


def _absorb_small_clusters(
    labels: list[int], vectors: np.ndarray, weights: list[float]
) -> list[int]:
    """Fold clusters under MIN_CLUSTER_SECONDS into the nearest larger centroid."""
    import numpy as np

    totals: dict[int, float] = {}
    for label, w in zip(labels, weights, strict=True):
        totals[label] = totals.get(label, 0.0) + w
    keep = [c for c, total in totals.items() if total >= MIN_CLUSTER_SECONDS]
    if not keep or len(keep) == len(totals):
        return labels
    centroids = {}
    for c in keep:
        members = [vectors[i] for i, lab in enumerate(labels) if lab == c]
        mean = np.mean(members, axis=0)
        centroids[c] = mean / max(float(np.linalg.norm(mean)), 1e-8)
    out = list(labels)
    for i, lab in enumerate(labels):
        if lab in centroids:
            continue
        out[i] = max(keep, key=lambda c: float(centroids[c] @ vectors[i]))
    order: dict[int, int] = {}
    for lab in out:
        order.setdefault(lab, len(order))
    return [order[lab] for lab in out]


def assign_speakers(
    segments: list[Any],
    audio_for_segment: Callable[[Any], np.ndarray | None],
    max_speakers: int,
    *,
    cancel_flag: Callable[[], bool] | None = None,
    on_progress: Callable[[float], None] | None = None,
) -> None:
    """Label *segments* in place with ``speaker`` ids; short ones inherit a neighbour."""
    import numpy as np

    max_speakers = max(2, min(8, int(max_speakers)))
    embedded: list[int] = []
    vectors: list[np.ndarray] = []
    durations: list[float] = []
    total = len(segments)
    for idx, seg in enumerate(segments):
        if cancel_flag and cancel_flag():
            raise DiarizationCancelled
        seg.pop("speaker", None)
        length = float(seg["end"]) - float(seg["start"])
        if length < MIN_EMBED_SECONDS:
            continue
        samples = audio_for_segment(seg)
        if samples is None or len(samples) < MIN_EMBED_SECONDS * SAMPLE_RATE:
            continue
        embedded.append(idx)
        vectors.append(embed(samples))
        durations.append(length)
        if on_progress:
            on_progress((idx + 1) / total)
    if not embedded:
        return
    with profiling.span("speakers.cluster"):
        units = _merge_adjacent(embedded, np.stack(vectors), durations)
        unit_vecs = np.stack([vec for _m, vec, _w in units])
        labels = cluster(unit_vecs, max_speakers)
        labels = _absorb_small_clusters(labels, unit_vecs, [w for _m, _v, w in units])
    for (member_idx, _vec, _w), label in zip(units, labels, strict=True):
        for m in member_idx:
            segments[m]["speaker"] = str(label + 1)
    labelled_mid = np.array(
        [(segments[i]["start"] + segments[i]["end"]) / 2 for i in embedded]
    )
    for seg in segments:
        if "speaker" not in seg:
            mid = (seg["start"] + seg["end"]) / 2
            nearest = embedded[int(np.argmin(np.abs(labelled_mid - mid)))]
            seg["speaker"] = segments[nearest]["speaker"]
    if on_progress:
        on_progress(1.0)


# ---------------------------------------------------------------------------
# Orchestration over a participant's video(s)
# ---------------------------------------------------------------------------


def diarize_entry(
    video_paths: list[str],
    segments: list[Any],
    audio_index: int,
    timeline: list[tuple[str, int, int]] | None,
    *,
    max_speakers: int,
    cancel_flag: Callable[[], bool] | None = None,
    on_progress: Callable[[float], None] | None = None,
) -> dict[str, Any] | None:
    """Label a transcript's segments across all parts; returns the ``speakers`` block.

    *timeline* is ``video.timeline_or_none(video_paths)``: ``(path, duration,
    cumulative_start)`` per part, or None for a single file. Each part is
    decoded once over the span its segments cover, segment windows are sliced
    out of that array, and the whole session is clustered together so a
    speaker keeps one id across parts. Returns None when audio cannot be
    decoded or the model is missing; the caller decides how to report that.
    """
    if config.DEBUGGING:
        width = min(2, max(1, max_speakers))
        for i, seg in enumerate(segments):
            seg["speaker"] = str(i % width + 1)
        return speakers_block(min(width, len(segments)))
    if not segments:
        return speakers_block(0)
    if not is_speaker_model_available():
        utils.warning_print("Speaker model is not installed; skipping speakers.")
        return None
    import video as video_mod

    parts = timeline or [(video_paths[0], 0, 0)]
    spans = _part_spans(parts, segments)
    decoded: dict[int, tuple[float, Any]] = {}

    def _decode(part_index: int) -> tuple[float, Any] | None:
        if part_index in decoded:
            return decoded[part_index]
        path = parts[part_index][0]
        win_start, win_end = spans[part_index]
        with profiling.span("speakers.decode"):
            samples = video_mod.decode_audio_pcm(
                path,
                audio_index,
                start_seconds=win_start or None,
                duration_seconds=max(win_end - win_start, MIN_EMBED_SECONDS),
            )
        if samples is None:
            return None
        decoded.clear()
        decoded[part_index] = (win_start, samples)
        return decoded[part_index]

    failed = False

    def audio_for_segment(seg: Any) -> Any:
        nonlocal failed
        part_index = _part_for(parts, float(seg["start"]))
        hit = _decode(part_index)
        if hit is None:
            failed = True
            return None
        win_start, samples = hit
        cumulative = parts[part_index][2]
        a = int((float(seg["start"]) - cumulative - win_start) * SAMPLE_RATE)
        b = int((float(seg["end"]) - cumulative - win_start) * SAMPLE_RATE)
        return samples[max(a, 0) : max(b, 0)]

    with profiling.span("speakers.diarize"):
        assign_speakers(
            segments,
            audio_for_segment,
            max_speakers,
            cancel_flag=cancel_flag,
            on_progress=on_progress,
        )
    if failed:
        utils.warning_print("Could not decode audio for speaker detection.")
        for seg in segments:
            seg.pop("speaker", None)
        return None
    count = len({seg["speaker"] for seg in segments if "speaker" in seg})
    return speakers_block(count)


def _part_for(parts: list[tuple[str, int, int]], start: float) -> int:
    """Index of the part whose cumulative window holds *start*; the last takes the tail."""
    for i in range(len(parts) - 1, -1, -1):
        if start >= parts[i][2]:
            return i
    return 0


def _part_spans(
    parts: list[tuple[str, int, int]], segments: list[Any]
) -> list[tuple[float, float]]:
    """Per part, the local ``(min start, max end)`` its segments cover."""
    spans: list[tuple[float, float]] = [(0.0, 0.0)] * len(parts)
    seen: dict[int, tuple[float, float]] = {}
    for seg in segments:
        i = _part_for(parts, float(seg["start"]))
        cumulative = parts[i][2]
        lo = max(0.0, float(seg["start"]) - cumulative)
        hi = max(lo, float(seg["end"]) - cumulative)
        prev = seen.get(i)
        seen[i] = (lo, hi) if prev is None else (min(prev[0], lo), max(prev[1], hi))
    for i, span in seen.items():
        spans[i] = span
    return spans
