"""redact.py: tokenizer, BIOES decode, windowing, numbering, download."""

import json
import os
import struct
from pathlib import Path
from typing import Any

import pytest

import config
import redact

_PIECES = ["<unk>", "<s>", "</s>", "▁", "▁Anna", "a", "@", "n", "▁x", "y"]


def _blob(pieces=_PIECES, scores=None, unk=0, bos=1, eos=2) -> bytes:
    scores = scores or [-1.0] * len(pieces)
    out = bytearray(b"RDTK\x01")
    out += struct.pack("<4i", unk, bos, eos, len(pieces))
    out += struct.pack(f"<{len(pieces)}f", *scores)
    encoded = [p.encode("utf-8") for p in pieces]
    out += struct.pack(f"<{len(pieces)}H", *[len(e) for e in encoded])
    for e in encoded:
        out += e
    return bytes(out)


# ---- Tokenizer -------------------------------------------------------------


def test_parse_tokenizer_reads_header_and_pieces():
    tok = redact.Tokenizer(_blob())
    assert (tok.unk_id, tok.bos_id, tok.eos_id, tok.vocab_size) == (0, 1, 2, 10)


def test_parse_tokenizer_rejects_bad_magic_and_trailing_bytes():
    with pytest.raises(ValueError):
        redact.Tokenizer(b"NOPE" + _blob()[4:])
    with pytest.raises(ValueError):
        redact.Tokenizer(_blob() + b"\x00")


def test_tokenize_prefers_whole_piece_and_falls_back_to_unk():
    tok = redact.Tokenizer(_blob())
    assert tok.tokenize("Anna") == [(4, "▁Anna")]
    ids = [i for i, _ in tok.tokenize("x  y")]
    assert ids == [8, 3, 9]


def test_reconstruct_offsets_scans_forward():
    tok = redact.Tokenizer(_blob())
    text = "Anna x"
    tokens = tok.tokenize(text)
    offsets = redact._reconstruct_offsets(text, tokens)
    assert offsets[0] == (0, 4)
    assert offsets[-1] == (5, 6)


# ---- Span algebra ----------------------------------------------------------


def _offs(n):
    return [(i * 2, i * 2 + 1) for i in range(n)]


def test_bioes_decodes_runs_singletons_and_orphans():
    tags = ["B-EMAIL", "I-EMAIL", "E-EMAIL", "O", "S-CITY", "I-PHONE", "E-CITY"]
    spans = redact.bioes_to_spans(tags, _offs(len(tags)))
    assert spans == [
        ("EMAIL", 0, 5),
        ("CITY", 8, 9),
        ("PHONE", 10, 11),
        ("CITY", 12, 13),
    ]


def test_bioes_skips_zero_width_tokens():
    spans = redact.bioes_to_spans(["S-CITY", "S-CITY"], [(0, 0), (2, 3)])
    assert spans == [("CITY", 2, 3)]


def test_snap_grows_over_word_chars_and_trims_punctuation():
    text = "Hi Anna-Lena, bye"
    assert redact._snap(text, 5, 8) == (3, 12)
    assert redact._snap("(anna@ex.se).", 1, 12) == (1, 11)


def test_merge_spans_unions_same_label_and_keeps_longer():
    spans = [
        {"label": "EMAIL", "start": 0, "end": 4, "score": 0.5},
        {"label": "EMAIL", "start": 3, "end": 8, "score": 0.9},
        {"label": "URL", "start": 6, "end": 20, "score": 0.7},
    ]
    out = redact._merge_spans(spans)
    assert [(s["label"], s["start"], s["end"]) for s in out] == [("URL", 6, 20)] or [
        (s["label"], s["start"], s["end"]) for s in out
    ] == [("EMAIL", 0, 8)]


def test_hysteresis_keeps_weak_name_next_to_strong_one():
    text = "Anna Lindqvist"
    strong = {"label": "GIVEN_NAME", "start": 0, "end": 4, "score": 0.9}
    weak = {"label": "SURNAME", "start": 5, "end": 14, "score": 0.4}
    lone = {"label": "CITY", "start": 5, "end": 14, "score": 0.4}
    assert redact._hysteresis(text, [strong, weak], 0.6) == [strong, weak]
    assert redact._hysteresis(text, [strong, lone], 0.6) == [strong]


# ---- Windowing -------------------------------------------------------------


class _WordTokenizer:
    bos_id = 1
    eos_id = 2

    def tokenize(self, text):
        return [(i, "▁" + w) for i, w in enumerate(text.split())]


def test_document_spans_windows_with_stride_and_dedupes(monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", False)
    words = [f"w{i}" for i in range(400)]
    words[200] = "anna@x.se"
    text = " ".join(words)
    id2label = {0: "O", 1: "S-EMAIL"}
    monkeypatch.setattr(
        redact, "_load_runtime", lambda: (None, _WordTokenizer(), id2label)
    )
    calls: list[list[int]] = []

    def fake_window(ids):
        calls.append(list(ids))
        # Token 200 is an email in every window that sees it, with a rising score.
        tags = ["S-EMAIL" if i == 200 else "O" for i in ids]
        return tags, [0.7 + 0.1 * len(calls) if i == 200 else 0.99 for i in ids]

    monkeypatch.setattr(redact, "_run_window", fake_window)
    spans = redact._document_spans(
        text,
        min_score=0.6,
        labels=frozenset({"EMAIL"}),
        cancel_flag=None,
        on_progress=None,
    )
    assert len(calls[0]) == redact.MAX_CONTENT
    assert calls[1][0] == redact.STEP
    assert calls[-1][-1] == 399
    assert len(spans) == 1
    assert text[spans[0]["start"] : spans[0]["end"]] == "anna@x.se"
    assert spans[0]["score"] == pytest.approx(0.9)


def test_document_spans_honours_cancel(monkeypatch):
    monkeypatch.setattr(
        redact, "_load_runtime", lambda: (None, _WordTokenizer(), {0: "O"})
    )
    monkeypatch.setattr(
        redact, "_run_window", lambda ids: (["O"] * len(ids), [1.0] * len(ids))
    )
    with pytest.raises(redact.RedactCancelled):
        redact._document_spans(
            "a b c",
            min_score=0.6,
            labels=frozenset(),
            cancel_flag=lambda: True,
            on_progress=None,
        )


def test_detect_spans_splits_at_segment_boundaries(monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", False)
    id2label = {0: "O", 1: "B-PHONE", 2: "E-PHONE"}
    monkeypatch.setattr(
        redact, "_load_runtime", lambda: (None, _WordTokenizer(), id2label)
    )
    # "070" ends segment one and "123" starts segment two; the span straddles them.
    monkeypatch.setattr(
        redact,
        "_run_window",
        lambda ids: (["O", "B-PHONE", "E-PHONE", "O"], [1.0, 0.9, 0.9, 1.0]),
    )
    out = redact.detect_spans(["call 070", "123 now"], min_score=0.6, org=False)
    assert [(s["label"], s["text"]) for s in out[0]] == [("PHONE", "070")]
    assert [(s["label"], s["text"]) for s in out[1]] == [("PHONE", "123")]


def test_detect_spans_drops_the_runtime_after_a_pass(monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", False)
    monkeypatch.setattr(redact, "_runtime", ("graph", _WordTokenizer(), {0: "O"}))
    monkeypatch.setattr(
        redact, "_run_window", lambda ids: (["O"] * len(ids), [1.0] * len(ids))
    )
    redact.detect_spans(["a b"], min_score=0.6, org=False)
    assert redact._runtime is None


@pytest.mark.parametrize(
    "inputs",
    [
        {"serving_default_input_ids": 0},
        {"serving_default_ids": 0, "serving_default_attention_mask": 1},
        {"input_ids": 0, "attention_mask": 1, "token_type_ids": 2},
    ],
)
def test_run_window_rejects_unexpected_inputs(monkeypatch, inputs):
    graph = {"inputs": inputs}
    monkeypatch.setattr(
        redact, "_load_runtime", lambda: (graph, _WordTokenizer(), {0: "O"})
    )
    with pytest.raises(ValueError, match="unexpected Redact model inputs"):
        redact._run_window([5, 6])


def test_enabled_labels_drop_org_unless_asked():
    fams = redact.label_families({0: "O", 1: "B-ORG", 2: "S-EMAIL", 3: "S-IMEI"})
    assert fams == {"ORG", "EMAIL"}
    assert redact.enabled_labels(fams, org=False) == {"EMAIL"}
    assert redact.enabled_labels(fams, org=True) == {"EMAIL", "ORG"}


def test_debugging_stub_tags_at_tokens(monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True)
    segs = [{"text": "mail anna@ex.se, please"}, {"text": "no contact"}]
    block = redact.redact_entry(segs, min_score=0.6, org=False)
    assert block is not None and block["count"] == 1
    assert segs[0]["pii"] == [
        {"label": "EMAIL", "start": 5, "end": 15, "score": 1.0, "text": "anna@ex.se"}
    ]
    assert segs[0]["pii_crc"] == redact.text_crc("mail anna@ex.se, please")
    assert segs[1]["pii"] == []


def test_redact_entry_without_model_returns_none(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "DEBUGGING", False)
    monkeypatch.setattr(redact, "models_dir", lambda: tmp_path / "none")
    assert redact.redact_entry([{"text": "x"}], min_score=0.6, org=False) is None


# ---- Read-time helpers -----------------------------------------------------


def _seg(text, *spans):
    return {
        "text": text,
        "pii": [
            {"label": lab, "start": a, "end": b, "score": 0.9, "text": text[a:b]}
            for lab, a, b in spans
        ],
        "pii_crc": redact.text_crc(text),
    }


def test_number_spans_reuses_number_for_same_surface():
    segs = [
        _seg("Anna met Erik.", ("GIVEN_NAME", 0, 4), ("GIVEN_NAME", 9, 13)),
        _seg("anna again", ("GIVEN_NAME", 0, 4)),
    ]
    spans = redact.entry_spans(segs, [s["text"] for s in segs])
    assert [s["placeholder"] for s in spans[0]] == ["[GIVEN_NAME_1]", "[GIVEN_NAME_2]"]
    assert spans[1][0]["placeholder"] == "[GIVEN_NAME_1]"


def test_entry_spans_reanchors_after_text_moved_and_drops_lost():
    seg = _seg("Hi Anna, mail anna@ex.se", ("GIVEN_NAME", 3, 7), ("EMAIL", 14, 24))
    moved = "Hello there Anna, mail anna@ex.se"
    spans = redact.entry_spans([seg], [moved])[0]
    assert [(s["label"], moved[s["start"] : s["end"]]) for s in spans] == [
        ("GIVEN_NAME", "Anna"),
        ("EMAIL", "anna@ex.se"),
    ]
    gone = "Hello there, mail someone@else.se"
    assert redact.entry_spans([seg], [gone])[0] == []


def test_render_text_substitutes_right_to_left():
    seg = _seg("Anna at anna@ex.se", ("GIVEN_NAME", 0, 4), ("EMAIL", 8, 18))
    spans = redact.entry_spans([seg], [seg["text"]])[0]
    assert redact.render_text(seg["text"], spans) == "[GIVEN_NAME_1] at [EMAIL_1]"


def test_entry_spans_flags_excluded_and_render_skips_them():
    seg = _seg("Anna met Erik.", ("GIVEN_NAME", 0, 4), ("GIVEN_NAME", 9, 13))
    excluded = [{"label": "GIVEN_NAME", "key": redact.surface_key("anna,")}]
    spans = redact.entry_spans([seg], [seg["text"]], excluded)[0]
    assert [s.get("excluded", False) for s in spans] == [True, False]
    # Numbering is assigned before the flag, so the kept span keeps its number.
    assert spans[1]["placeholder"] == "[GIVEN_NAME_2]"
    assert redact.render_text(seg["text"], spans) == "Anna met [GIVEN_NAME_2]."


def test_utf16_spans_count_astral_characters_twice():
    text = "\U0001f41c Anna"
    spans = [{"label": "GIVEN_NAME", "start": 2, "end": 6, "score": 1.0}]
    out = redact.utf16_spans(text, spans)
    assert (out[0]["start"], out[0]["end"]) == (3, 7)


def test_strip_entry_drops_spans():
    seg = _seg("Anna", ("GIVEN_NAME", 0, 4))
    redact.strip_entry([seg])
    assert "pii" not in seg and "pii_crc" not in seg


# ---- Assets and download ---------------------------------------------------


def test_assets_are_pinned_to_the_tag():
    assert [a["filename"] for a in redact.ASSETS] == [
        redact.MODEL_FILENAME,
        redact.TOKENIZER_FILENAME,
        redact.LABELS_FILENAME,
    ]
    for asset in redact.ASSETS:
        assert f"/resolve/{redact.MODEL_TAG}/" in asset["url"]
        assert len(asset["sha256"]) == 64 and int(asset["sha256"], 16)
        assert asset["size"] > 0


def test_download_streams_missing_files_and_is_idempotent(monkeypatch, tmp_path):
    import llm_client

    monkeypatch.setattr(redact, "models_dir", lambda: tmp_path / "rm")
    pulled: list[str] = []

    def fake_stream(url, target, *, sha256, size, on_progress=None):
        pulled.append(target.name)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(b"x" * size)
        if on_progress:
            on_progress(
                {"status": "downloading model", "completed": size, "total": size}
            )
        return True

    monkeypatch.setattr(llm_client, "stream_download", fake_stream)
    seen: list[dict] = []
    assert redact.download(seen.append) is True
    assert pulled == [a["filename"] for a in redact.ASSETS]
    assert redact.is_redact_model_available()
    assert (
        seen[-1]["completed"]
        == seen[-1]["total"]
        == sum(a["size"] for a in redact.ASSETS)
    )
    pulled.clear()
    assert redact.download() is True
    assert pulled == []
    redact.remove()
    assert not redact.is_redact_model_available()


def test_download_failure_leaves_model_unavailable(monkeypatch, tmp_path):
    import llm_client

    monkeypatch.setattr(redact, "models_dir", lambda: tmp_path / "rm")
    monkeypatch.setattr(llm_client, "stream_download", lambda *a, **k: False)
    assert redact.download() is False
    assert not redact.is_redact_model_available()


# ---- Real model smoke (only where the files exist) ---------------------------


def _real_model_dir() -> Path | None:
    candidates = [os.environ.get("CLIPGEN_REDACT_MODEL_DIR", "")]
    candidates.append(
        str(Path(__file__).resolve().parent.parent / ".context" / "redact-spike")
    )
    for c in candidates:
        if c and all((Path(c) / a["filename"]).is_file() for a in redact.ASSETS):
            return Path(c)
    return None


@pytest.mark.skipif(_real_model_dir() is None, reason="Redact model not downloaded")
def test_real_model_finds_name_email_and_city(monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", False)
    monkeypatch.setattr(redact, "models_dir", _real_model_dir)
    monkeypatch.setattr(redact, "_runtime", None)
    segs: list[dict[str, Any]] = [
        {"text": "Jag heter Anna Lindqvist och bor i Umeå."},
        {"text": "Mejla anna.lindqvist@example.se"},
    ]
    block = redact.redact_entry(segs, min_score=0.6, org=False)
    assert block is not None and block["count"] >= 4
    labels = {s["label"]: s["text"] for seg in segs for s in seg["pii"]}
    assert labels["GIVEN_NAME"] == "Anna"
    assert labels["SURNAME"] == "Lindqvist"
    assert labels["CITY"] == "Umeå"
    assert labels["EMAIL"] == "anna.lindqvist@example.se"


# ---- LiteRT-verified reference (build/redact_parity.py) ---------------------

_REFERENCE = Path(__file__).resolve().parent / "fixtures" / "redact_reference.json"


def test_reference_fixture_matches_pinned_assets():
    """A model bump fails here until it is re-verified against LiteRT."""
    ref = json.loads(_REFERENCE.read_text(encoding="utf-8"))
    pinned = {a["filename"]: a["sha256"] for a in redact.ASSETS}
    assert (ref["model_tag"], ref["assets"]) == (redact.MODEL_TAG, pinned), (
        "redact.ASSETS changed; run `uv run --with ai-edge-litert==2.2.0 "
        "build/redact_parity.py --write` and commit the fixture it writes"
    )


@pytest.mark.skipif(_real_model_dir() is None, reason="Redact model not downloaded")
def test_real_model_reproduces_reference(monkeypatch):
    """The numpy runtime replays LiteRT's tags and spans on the pinned model."""
    ref = json.loads(_REFERENCE.read_text(encoding="utf-8"))
    monkeypatch.setattr(config, "DEBUGGING", False)
    monkeypatch.setattr(redact, "models_dir", _real_model_dir)
    monkeypatch.setattr(redact, "_runtime", None)
    windows: list[list[str]] = []
    real_window = redact._run_window

    def recording(ids):
        tags, probs = real_window(ids)
        windows.append(tags)
        return tags, probs

    monkeypatch.setattr(redact, "_run_window", recording)
    spans = redact.detect_spans(ref["texts"], min_score=ref["min_score"], org=True)
    assert len(windows) == len(ref["windows"])
    pairs = [
        (a, b)
        for got, want in zip(windows, ref["windows"], strict=True)
        for a, b in zip(got, want, strict=True)
    ]
    # Another BLAS may flip a near-tie token; the spans themselves must not move.
    assert sum(a == b for a, b in pairs) >= 0.995 * len(pairs)
    got_spans = [[[s["label"], s["start"], s["end"]] for s in seg] for seg in spans]
    assert got_spans == ref["spans"]
