"""PII detection for transcripts with Desert Ant Labs' Redact model.

A standalone pass over an existing transcript, like ``speakers.py``: the
worker runs ``redact_entry`` on a participant's corrected segments and writes
``segment["pii"]`` (``[{label, start, end, score, text}]``, code-point offsets
into that segment's text) plus ``segment["pii_crc"]`` (crc32 of the text the
offsets index). The text itself is never rewritten; every reader that wants
placeholders calls ``entry_spans`` + ``render_text`` at read time, which also
re-anchors spans after a correction moved the text and numbers them so the
same surface gets the same ``[GIVEN_NAME_1]`` across a participant.

The model (``ASSETS``) is an opt-in download into ``models_dir()``: a 23M
parameter 6-layer BERT token classifier over 89 BIOES tags (int8 TFLite),
evaluated in numpy by ``tflite_numpy``. Text is tokenized by the model's own unigram
tokenizer (``redact_tokenizer.bin``: ``RDTK`` magic, one version byte,
``<4i`` unk/bos/eos/count, ``count`` float32 scores, ``count`` uint16 piece
lengths, UTF-8 pieces; Viterbi over NFKC text with runs of spaces squeezed
to one ``▁``) and fed in overlapping windows of ``MAX_CONTENT`` tokens
(``input_ids``/``attention_mask`` int32 ``[1, SEQ]`` → logits ``[1, SEQ,
labels]``). The span algebra (BIOES decode, best-score dedupe across
windows, name hysteresis, overlap resolution, word snapping) follows the
vendor SDK; the deterministic checksum layer (IBAN, Luhn, IMEI) and the
US-address regexes are not ported — spoken transcripts rarely carry them.

Module-level imports stay light on purpose: numpy and ``tflite_numpy``
load inside functions, so ``data_export`` and the server can import the
read-time helpers for free.
"""

from __future__ import annotations

import json
import struct
import threading
import unicodedata
import zlib
from collections.abc import Callable, Iterable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import config
import profiling
import utils

MODEL_TAG = "v0.4.0"
MODEL_FILENAME = "redact.tflite"
TOKENIZER_FILENAME = "redact_tokenizer.bin"
LABELS_FILENAME = "labels.json"
_HF_BASE = f"https://huggingface.co/desert-ant-labs/redact/resolve/{MODEL_TAG}"
# Pinned to the tag; the sha256 values are the repo's LFS oids.
ASSETS: list[dict[str, Any]] = [
    {
        "filename": MODEL_FILENAME,
        "url": f"{_HF_BASE}/{MODEL_FILENAME}",
        "size": 24529472,
        "sha256": "ee36727f07e3237569e71427bfe661463a82e526f7e97e92b0fa583cad16ed27",
    },
    {
        "filename": TOKENIZER_FILENAME,
        "url": f"{_HF_BASE}/{TOKENIZER_FILENAME}",
        "size": 391416,
        "sha256": "81bc354dc99285e05b8f70edca9df3afbd83dfa4b0cea18d0448ab17ce58aae2",
    },
    {
        "filename": LABELS_FILENAME,
        "url": f"{_HF_BASE}/{LABELS_FILENAME}",
        "size": 3842,
        "sha256": "228835734235f85ec4bf31391a296aa6e8a36684e1171e841ac94f7b3f9b430a",
    },
]
MODEL_SIZE_MB = round(sum(a["size"] for a in ASSETS) / (1024 * 1024))
LICENSE_NAME = "Desert Ant Labs Source-Available License 1.0"
LICENSE_URL = "https://license.desertant.com/1.0"
MODEL_URL = "https://huggingface.co/desert-ant-labs/redact"
VENDOR_URL = "https://desertant.com"

SEQ = 256
MAX_CONTENT = SEQ - 2
STRIDE = 64
STEP = MAX_CONTENT - STRIDE
# Tags below this confidence are read as "O" before decoding.
LOW_SCORE = 0.3
PAD_ID = 1
# Regex-only in the vendor SDK; the classifier never emits it.
DETERMINISTIC_LABELS = frozenset({"IMEI"})
OPTIONAL_LABELS = frozenset({"ORG"})
NAME_LABELS = frozenset({"GIVEN_NAME", "SURNAME"})
_META = "▁"
_CONNECT = frozenset({"-", "'", "’"})
_TRIM = " \t\n\r.,;:!?()[]{}\"'«»‘’“”"
SEGMENT_JOIN = "\n"

_runtime_lock = threading.Lock()
_runtime: tuple[dict[str, Any], Tokenizer, dict[int, str]] | None = None


class RedactCancelled(Exception):
    """Raised when the cancel flag trips mid-pass."""


# ---------------------------------------------------------------------------
# Model files
# ---------------------------------------------------------------------------


def models_dir() -> Path:
    """Where the downloaded Redact files live, under the per-user config dir."""
    import start_settings

    return start_settings.config_dir() / "redact_models"


def is_redact_model_available() -> bool:
    """True when every asset is present; a hand-copied set counts."""
    base = models_dir()
    return all((base / a["filename"]).is_file() for a in ASSETS)


def download(on_progress: Callable[[dict[str, Any]], None] | None = None) -> bool:
    """Fetch the missing assets; progress sums across files. Never raises."""
    import llm_client

    base = models_dir()
    total = sum(int(a["size"]) for a in ASSETS)
    done_before = 0
    for asset in ASSETS:
        target = base / asset["filename"]
        if target.is_file() and target.stat().st_size == asset["size"]:
            done_before += int(asset["size"])
            continue

        def _relay(chunk: dict[str, Any], _base: int = done_before) -> None:
            if on_progress is not None:
                on_progress(
                    {
                        "status": "downloading model",
                        "completed": _base + int(chunk.get("completed") or 0),
                        "total": total,
                    }
                )

        if not llm_client.stream_download(
            asset["url"],
            target,
            sha256=asset["sha256"],
            size=int(asset["size"]),
            on_progress=_relay,
        ):
            return False
        done_before += int(asset["size"])
    utils.info_print("Downloaded the Redact model.")
    return True


def remove() -> None:
    """Delete the downloaded assets and drop the cached runtime."""
    base = models_dir()
    for asset in ASSETS:
        (base / asset["filename"]).unlink(missing_ok=True)
    _drop_runtime()


def redaction_block(count: int, *, min_score: float, org: bool) -> dict[str, Any]:
    """Fresh manifest ``redaction`` block after a detection run."""
    return {
        "enabled": True,
        "count": count,
        "min_score": min_score,
        "org": org,
        "detected_at": datetime.now(UTC).isoformat(),
    }


# ---------------------------------------------------------------------------
# Tokenizer
# ---------------------------------------------------------------------------


class Tokenizer:
    """The model's unigram tokenizer, parsed from the ``RDTK`` blob."""

    def __init__(self, data: bytes) -> None:
        if len(data) < 21 or data[:4] != b"RDTK":
            raise ValueError("invalid tokenizer header")
        offset = 5
        unk_id, bos_id, eos_id, count = struct.unpack_from("<4i", data, offset)
        offset += 16
        if count <= 0 or count > (len(data) - offset) // 6:
            raise ValueError("invalid tokenizer vocabulary count")
        scores = struct.unpack_from(f"<{count}f", data, offset)
        offset += count * 4
        lengths = struct.unpack_from(f"<{count}H", data, offset)
        offset += count * 2
        pieces: dict[str, int] = {}
        longest = 1
        for piece_id, length in enumerate(lengths):
            end = offset + length
            if end > len(data):
                raise ValueError("truncated tokenizer piece")
            piece = data[offset:end].decode("utf-8")
            offset = end
            pieces[piece] = piece_id
            longest = max(longest, len(piece))
        if offset != len(data):
            raise ValueError("unexpected trailing tokenizer data")
        if not all(0 <= i < count for i in (unk_id, bos_id, eos_id)):
            raise ValueError("tokenizer special id out of range")
        self.bos_id = bos_id
        self.eos_id = eos_id
        self.unk_id = unk_id
        self.vocab_size = count
        self._scores = scores
        self._pieces = pieces
        self._max_len = min(longest, 32)
        self._unknown_penalty = min(scores) - 10.0

    def tokenize(self, text: str) -> list[tuple[int, str]]:
        """Viterbi-optimal ``(id, surface)`` pairs; surfaces keep their ``▁``."""
        squeezed: list[str] = []
        last_space = True
        for ch in unicodedata.normalize("NFKC", text):
            if ch == " ":
                if last_space:
                    continue
                last_space = True
            else:
                last_space = False
            squeezed.append(ch)
        if squeezed and squeezed[-1] == " ":
            squeezed.pop()
        s = _META + "".join(_META if c == " " else c for c in squeezed)
        n = len(s)
        best = [-1e18] * (n + 1)
        best[0] = 0.0
        back_pos = [-1] * (n + 1)
        back_id = [-1] * (n + 1)
        pieces, scores, max_len = self._pieces, self._scores, self._max_len
        for end in range(1, n + 1):
            for start in range(max(0, end - max_len), end):
                tid = pieces.get(s[start:end])
                if tid is None:
                    continue
                score = best[start] + scores[tid]
                if score > best[end]:
                    best[end] = score
                    back_pos[end] = start
                    back_id[end] = tid
            unknown = best[end - 1] + self._unknown_penalty
            if unknown > best[end]:
                best[end] = unknown
                back_pos[end] = end - 1
                back_id[end] = self.unk_id
        out: list[tuple[int, str]] = []
        pos = n
        while pos > 0:
            start = back_pos[pos]
            out.append((back_id[pos], s[start:pos]))
            pos = start
        out.reverse()
        return out


def _reconstruct_offsets(
    text: str, tokens: list[tuple[int, str]]
) -> list[tuple[int, int]]:
    """Map token surfaces back onto *text* by scanning forward."""
    cursor = 0
    offsets: list[tuple[int, int]] = []
    for _tid, surface in tokens:
        surface = surface.replace(_META, "")
        if not surface:
            offsets.append((cursor, cursor))
            continue
        found = text.find(surface, cursor)
        if found < 0:
            offsets.append((cursor, cursor))
        else:
            offsets.append((found, found + len(surface)))
            cursor = found + len(surface)
    return offsets


# ---------------------------------------------------------------------------
# Runtime
# ---------------------------------------------------------------------------


def _load_runtime() -> tuple[dict[str, Any], Tokenizer, dict[int, str]]:
    """Model graph, tokenizer and id→label map, cached until ``_drop_runtime``."""
    global _runtime
    if _runtime is not None:
        return _runtime
    with _runtime_lock:
        if _runtime is not None:
            return _runtime
        import tflite_numpy

        base = models_dir()
        with profiling.span("redact.runtime_load"):
            graph = tflite_numpy.load_graph(base / MODEL_FILENAME)
            tokenizer = Tokenizer((base / TOKENIZER_FILENAME).read_bytes())
            labels = json.loads((base / LABELS_FILENAME).read_text(encoding="utf-8"))
            id2label = {int(k): str(v) for k, v in labels["id2label"].items()}
        _runtime = (graph, tokenizer, id2label)
        return _runtime


def _drop_runtime() -> None:
    """Free the ~90 MB of float weights; a reload takes ~60 ms."""
    global _runtime
    with _runtime_lock:
        _runtime = None


def label_families(id2label: dict[int, str]) -> set[str]:
    """Label families the classifier can emit (``B-EMAIL`` → ``EMAIL``)."""
    return {
        tag.split("-", 1)[1]
        for tag in id2label.values()
        if "-" in tag and tag.split("-", 1)[1] not in DETERMINISTIC_LABELS
    }


def enabled_labels(families: Iterable[str], *, org: bool) -> frozenset[str]:
    """Families to keep: everything except the optional ones unless asked."""
    return frozenset(f for f in families if org or f not in OPTIONAL_LABELS)


def _run_window(ids: list[int]) -> tuple[list[str], list[float]]:
    """Tags and confidences for one content window (bos/eos stripped)."""
    import numpy as np

    import tflite_numpy

    graph, tokenizer, id2label = _load_runtime()
    seq = [tokenizer.bos_id, *ids, tokenizer.eos_id]
    padded = np.full((1, SEQ), PAD_ID, dtype=np.int32)
    mask = np.zeros((1, SEQ), dtype=np.int32)
    padded[0, : len(seq)] = seq
    mask[0, : len(seq)] = 1
    inputs = graph["inputs"]
    ids_index = next((i for n, i in inputs.items() if "input_ids" in n), None)
    mask_index = next((i for n, i in inputs.items() if "attention_mask" in n), None)
    if len(inputs) != 2 or ids_index is None or mask_index is None:
        raise ValueError(f"unexpected Redact model inputs: {sorted(inputs)}")
    feeds = {ids_index: padded, mask_index: mask}
    logits = tflite_numpy.run_graph(graph, feeds)[0]
    logits = logits[0, : len(seq)].astype(np.float32)
    shifted = logits - logits.max(axis=-1, keepdims=True)
    probs = np.exp(shifted)
    probs /= probs.sum(axis=-1, keepdims=True)
    picks = probs.argmax(axis=-1)
    tags = [id2label.get(int(i), "O") for i in picks][1:-1]
    scores = [float(p) for p in probs[np.arange(len(seq)), picks]][1:-1]
    return tags, scores


# ---------------------------------------------------------------------------
# Span algebra
# ---------------------------------------------------------------------------


def bioes_to_spans(
    tags: list[str], offsets: list[tuple[int, int]]
) -> list[tuple[str, int, int]]:
    """Decode BIOES tags over token offsets into ``(label, start, end)``."""
    out: list[tuple[str, int, int]] = []
    label: str | None = None
    start = end = 0

    def close() -> None:
        nonlocal label
        if label is not None and end > start:
            out.append((label, start, end))
        label = None

    for tag, (tok_start, tok_end) in zip(tags, offsets, strict=True):
        if tok_end <= tok_start:
            continue
        if tag == "O" or "-" not in tag:
            close()
            continue
        prefix, family = tag.split("-", 1)
        if prefix == "S":
            close()
            out.append((family, tok_start, tok_end))
        elif prefix == "B":
            close()
            label, start, end = family, tok_start, tok_end
        elif prefix == "I":
            if label == family:
                end = tok_end
            else:
                close()
                label, start, end = family, tok_start, tok_end
        elif prefix == "E":
            if label == family:
                end = tok_end
                close()
            else:
                close()
                out.append((family, tok_start, tok_end))
        else:
            close()
    close()
    return out


def _is_word_char(ch: str) -> bool:
    return ch.isalnum() or ch == "_"


def _snap(text: str, start: int, end: int) -> tuple[int, int]:
    """Grow a span over adjacent word characters and in-word connectors."""
    start = max(0, min(start, len(text)))
    end = max(0, min(end, len(text)))
    while start > 0 and (
        _is_word_char(text[start - 1])
        or (
            text[start - 1] in _CONNECT
            and start >= 2
            and _is_word_char(text[start - 2])
        )
    ):
        start -= 1
    while end < len(text) and (
        _is_word_char(text[end])
        or (
            text[end] in _CONNECT
            and end + 1 < len(text)
            and _is_word_char(text[end + 1])
        )
    ):
        end += 1
    while end > start and text[end - 1] in _TRIM:
        end -= 1
    while start < end and text[start] in _TRIM:
        start += 1
    return start, end


def _merge_spans(spans: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Merge overlaps: same label unions, different labels keep the longer."""
    out: list[dict[str, Any]] = []
    for span in sorted(
        spans, key=lambda s: (s["start"], -(s["end"] - s["start"]), s["label"])
    ):
        if not out or span["start"] >= out[-1]["end"]:
            out.append(dict(span))
        elif span["label"] == out[-1]["label"]:
            out[-1]["end"] = max(out[-1]["end"], span["end"])
            out[-1]["score"] = max(out[-1]["score"], span["score"])
        elif span["end"] - span["start"] > out[-1]["end"] - out[-1]["start"]:
            out[-1] = dict(span)
    return out


def _name_like_gap(gap: str) -> bool:
    if not gap.strip() or len(gap) > 20 or any(c in gap for c in ',;:/&|()[]{}"<>\n\t'):
        return False
    return all(
        len(w.strip(".-'’")) <= 1 or w.strip(".-'’")[0].isupper() for w in gap.split()
    )


def _adjacent_names(text: str, a: dict[str, Any], b: dict[str, Any]) -> bool:
    left, right = sorted((a, b), key=lambda s: s["start"])
    if right["start"] < left["end"]:
        return True
    gap = text[left["end"] : right["start"]]
    return (gap.isspace() and len(gap) <= 3) or _name_like_gap(gap)


def _hysteresis(
    text: str, spans: list[dict[str, Any]], min_score: float
) -> list[dict[str, Any]]:
    """Keep confident spans, plus weak name parts touching a kept name."""
    kept = [s for s in spans if s["score"] >= min_score]
    weak = [s for s in spans if s["score"] < min_score and s["label"] in NAME_LABELS]
    changed = True
    while changed and weak:
        changed = False
        for span in list(weak):
            if any(
                k["label"] in NAME_LABELS and _adjacent_names(text, span, k)
                for k in kept
            ):
                kept.append(span)
                weak.remove(span)
                changed = True
                break
    return kept


# ---------------------------------------------------------------------------
# Detection
# ---------------------------------------------------------------------------


def _stub_spans(texts: list[str]) -> list[list[dict[str, Any]]]:
    """DEBUGGING stand-in: every token containing ``@`` is an EMAIL."""
    out: list[list[dict[str, Any]]] = []
    for text in texts:
        spans: list[dict[str, Any]] = []
        pos = 0
        for token in text.split(" "):
            if "@" in token:
                start, end = _snap(text, pos, pos + len(token))
                if end > start:
                    spans.append(
                        {
                            "label": "EMAIL",
                            "start": start,
                            "end": end,
                            "score": 1.0,
                            "text": text[start:end],
                        }
                    )
            pos += len(token) + 1
        out.append(spans)
    return out


def _document_spans(
    text: str,
    *,
    min_score: float,
    labels: frozenset[str],
    cancel_flag: Callable[[], bool] | None,
    on_progress: Callable[[float], None] | None,
) -> list[dict[str, Any]]:
    """Windowed inference over one document; spans carry code-point offsets."""
    _graph, tokenizer, _id2label = _load_runtime()
    tokens = tokenizer.tokenize(text)
    offsets = _reconstruct_offsets(text, tokens)
    ids = [tid for tid, _ in tokens]
    low = min(LOW_SCORE, min_score)
    best: dict[tuple[int, int, str], dict[str, Any]] = {}
    index = 0
    while tokens:
        if cancel_flag is not None and cancel_flag():
            raise RedactCancelled
        end = min(index + MAX_CONTENT, len(tokens))
        tags, probs = _run_window(ids[index:end])
        window_offsets = offsets[index:end]
        usable = [t if p >= low else "O" for t, p in zip(tags, probs, strict=True)]
        for label, start, stop in bioes_to_spans(usable, window_offsets):
            score = max(
                (
                    p
                    for (a, b), p in zip(window_offsets, probs, strict=True)
                    if b > a and max(a, start) < min(b, stop)
                ),
                default=0.0,
            )
            key = (start, stop, label)
            if key not in best or score > best[key]["score"]:
                best[key] = {
                    "label": label,
                    "start": start,
                    "end": stop,
                    "score": score,
                }
        if on_progress is not None:
            on_progress(end / len(tokens))
        if end == len(tokens):
            break
        index += STEP
    kept = _hysteresis(text, list(best.values()), min_score)
    snapped = []
    for span in kept:
        start, stop = _snap(text, span["start"], span["end"])
        if stop > start and span["label"] in labels:
            snapped.append({**span, "start": start, "end": stop})
    return _merge_spans(snapped)


def detect_spans(
    texts: list[str],
    *,
    min_score: float,
    org: bool,
    cancel_flag: Callable[[], bool] | None = None,
    on_progress: Callable[[float], None] | None = None,
) -> list[list[dict[str, Any]]]:
    """PII spans per text, detected over the texts joined as one document.

    Windows then see context across segment boundaries. A span that straddles
    a boundary is cut at it. Each span carries its surface as ``text`` so a
    reader can re-anchor it after the segment text changes.
    """
    if config.DEBUGGING:
        return _stub_spans(texts)
    if not texts:
        return []
    try:
        _graph, _tokenizer, id2label = _load_runtime()
        labels = enabled_labels(label_families(id2label), org=org)
        with profiling.span("redact.detect"):
            spans = _document_spans(
                SEGMENT_JOIN.join(texts),
                min_score=min_score,
                labels=labels,
                cancel_flag=cancel_flag,
                on_progress=on_progress,
            )
    finally:
        _drop_runtime()
    out: list[list[dict[str, Any]]] = [[] for _ in texts]
    starts: list[int] = []
    pos = 0
    for text in texts:
        starts.append(pos)
        pos += len(text) + len(SEGMENT_JOIN)
    for span in spans:
        for i, text in enumerate(texts):
            seg_start = starts[i]
            seg_end = seg_start + len(text)
            start = max(span["start"], seg_start) - seg_start
            end = min(span["end"], seg_end) - seg_start
            if end <= start:
                continue
            out[i].append(
                {
                    "label": span["label"],
                    "start": start,
                    "end": end,
                    "score": round(float(span["score"]), 3),
                    "text": text[start:end],
                }
            )
    return out


def text_crc(text: str) -> int:
    return zlib.crc32(text.encode("utf-8"))


def redact_entry(
    segments: list[Any],
    *,
    min_score: float,
    org: bool,
    cancel_flag: Callable[[], bool] | None = None,
    on_progress: Callable[[float], None] | None = None,
) -> dict[str, Any] | None:
    """Tag *segments* in place; returns the ``redaction`` block, None if no model."""
    if not segments:
        return redaction_block(0, min_score=min_score, org=org)
    if not config.DEBUGGING and not is_redact_model_available():
        utils.warning_print("Redact model is not installed; skipping redaction.")
        return None
    texts = [str(seg.get("text") or "") for seg in segments]
    per_segment = detect_spans(
        texts,
        min_score=min_score,
        org=org,
        cancel_flag=cancel_flag,
        on_progress=on_progress,
    )
    count = 0
    for seg, text, spans in zip(segments, texts, per_segment, strict=True):
        seg["pii"] = spans
        seg["pii_crc"] = text_crc(text)
        count += len(spans)
    return redaction_block(count, min_score=min_score, org=org)


# ---------------------------------------------------------------------------
# Read-time helpers (pure)
# ---------------------------------------------------------------------------


def _anchor(text: str, spans: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Re-find each span's surface after the text changed; lost ones drop."""
    out: list[dict[str, Any]] = []
    cursor = 0
    for span in sorted(spans, key=lambda s: s["start"]):
        surface = str(span.get("text") or "")
        if not surface:
            continue
        found = text.find(surface, cursor)
        if found < 0:
            continue
        out.append({**span, "start": found, "end": found + len(surface)})
        cursor = found + len(surface)
    return out


def entry_spans(
    segments: list[Any],
    texts: list[str],
    excluded: list[dict[str, Any]] | None = None,
) -> list[list[dict[str, Any]]]:
    """Numbered spans per segment against *texts*, the text readers show.

    *excluded* is the entry's restore list (``[{label, key}]``); a matching
    span stays in the list flagged ``excluded`` so readers can offer to redact
    it again, and every renderer skips it.
    """
    per_segment: list[list[dict[str, Any]]] = []
    for seg, text in zip(segments, texts, strict=True):
        spans = list(seg.get("pii") or [])
        if not spans:
            per_segment.append([])
        elif seg.get("pii_crc") == text_crc(text):
            per_segment.append([dict(s) for s in spans])
        else:
            per_segment.append(_anchor(text, spans))
    numbered = number_spans(per_segment)
    keys = {(e.get("label"), e.get("key")) for e in excluded or []}
    if keys:
        for spans in numbered:
            for span in spans:
                if (span["label"], surface_key(str(span.get("text") or ""))) in keys:
                    span["excluded"] = True
    return numbered


def surface_key(surface: str) -> str:
    """Case- and punctuation-insensitive identity of a detected surface."""
    folded = unicodedata.normalize("NFKC", surface).casefold()
    return " ".join(folded.strip(_TRIM).split())


def number_spans(per_segment: list[list[dict[str, Any]]]) -> list[list[dict[str, Any]]]:
    """Assign ``[LABEL_N]`` placeholders; the same surface keeps its number."""
    counters: dict[str, int] = {}
    seen: dict[tuple[str, str], str] = {}
    for spans in per_segment:
        for span in sorted(spans, key=lambda s: s["start"]):
            key = (span["label"], surface_key(str(span.get("text") or "")))
            placeholder = seen.get(key)
            if placeholder is None:
                counters[span["label"]] = counters.get(span["label"], 0) + 1
                placeholder = f"[{span['label']}_{counters[span['label']]}]"
                seen[key] = placeholder
            span["placeholder"] = placeholder
    return per_segment


def render_text(text: str, spans: list[dict[str, Any]]) -> str:
    """Replace each span with its placeholder, right to left; restored ones stay."""
    out = text
    for span in sorted(spans, key=lambda s: s["start"], reverse=True):
        if span.get("excluded"):
            continue
        start, end = int(span["start"]), int(span["end"])
        if 0 <= start < end <= len(out):
            out = out[:start] + str(span.get("placeholder") or "[PII]") + out[end:]
    return out


def utf16_spans(text: str, spans: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Spans with UTF-16 unit offsets, the units JavaScript slices by."""

    def units(index: int) -> int:
        return len(text[:index].encode("utf-16-le")) // 2

    return [
        {**span, "start": units(int(span["start"])), "end": units(int(span["end"]))}
        for span in spans
    ]


def strip_entry(segments: list[Any]) -> None:
    """Drop every stored span; used when a participant turns redaction off."""
    for seg in segments:
        seg.pop("pii", None)
        seg.pop("pii_crc", None)
