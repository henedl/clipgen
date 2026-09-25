"""Verify the numpy Redact runtime against LiteRT and record the reference fixture.

clipgen runs Redact's TFLite graph with ``source/tflite_numpy.py``, not LiteRT. Any
change to the pinned model (``redact.ASSETS``) or to the evaluator must be checked
against the vendor runtime before it merges, and this script is that check:

    uv run --with ai-edge-litert==2.2.0 build/redact_parity.py          # report only
    uv run --with ai-edge-litert==2.2.0 build/redact_parity.py --write  # + fixture

It runs ``redact.detect_spans`` over ``REFERENCE_TEXTS`` twice (LiteRT windows, then
numpy windows) and compares the op graph, every window's argmax tags, and the
final spans. ``--write`` refuses unless all three agree, then stores LiteRT's
output in ``tests/fixtures/redact_reference.json``. ``test_redact.py`` fails while
that fixture's hashes differ from ``ASSETS``, and replays it against the numpy
runtime wherever the model is present (CI fetches it with ``--fetch DIR``).

LiteRT is never a dependency: ``--with`` installs it into a throwaway environment.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib
import importlib.metadata
import json
import sys
from pathlib import Path
from typing import Any

_ROOT = Path(__file__).resolve().parents[1]
FIXTURE = _ROOT / "tests" / "fixtures" / "redact_reference.json"

# `source/` holds the modules; same sys.path insert as tests/conftest.py.
sys.path.insert(0, str(_ROOT / "source"))

import config
import redact

MIN_SCORE = 0.6

# Fictional people and numbers across the model's scripts and label families.
REFERENCE_TEXTS = [
    "Hi, I'm Anna Kovács from Stockholm, you can email me at anna.k@example.se or call 070 123 45 67.",
    "My colleague Erik Johansson works at Visma on Drottninggatan 12 in Uppsala.",
    "so yeah I just clicked on the settings button and then it asked me for my password again, my wife Maria usually handles this",
    "Nous avons parlé avec Pierre Dubois à Lyon, il habite 14 rue de la République.",
    "Ich heiße Jürgen Schmidt und wohne in der Hauptstraße 5, 10115 Berlin.",
    "Mi chiamo Giulia Rossi, il mio numero è 333 1234567.",
    "Me llamo Carlos García y vivo en Madrid, calle Mayor 3.",
    "Jeg heter Ingrid Nilsen og bor i Bergen.",
    "Nazywam się Piotr Kowalski, mój email to piotr.kowalski@wp.pl.",
    "um okay so the the thing is I don't really know where to click here, is it this one? oh it's loading",
    "My social security number is 123-45-6789 and my card is 4111 1111 1111 1111.",
    "Visit https://example.com/login and use IP 192.168.1.14 to connect.",
    "Olen Mikko Virtanen Helsingistä.",
    "Ik ben Sanne de Vries uit Amsterdam, bel me op 06 12345678.",
    "So Tom from Spotify said Linda should ask Bob at the Microsoft office in Dublin.",
    "Participant: I usually use Excel for this. Moderator: Okay, and what about Jira?",
    "Το όνομά μου είναι Γιώργος Παπαδόπουλος και μένω στην Αθήνα.",
    "Казвам се Иван Петров и живея в София.",
    "Én Kovács Péter vagyok, Budapesten lakom.",
    "The account number is GB29NWBK60161331926819 and my passport is 123456789.",
    " ".join(
        ["Anna told me that Johan Berg and Sara Lind met in Göteborg last week."] * 12
    ),
]


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def run_reference(window: Any = None) -> tuple[list[list[str]], list[list[list[Any]]]]:
    """Per-window tags and per-text spans, optionally through a replacement window."""
    real_window = redact._run_window
    windows: list[list[str]] = []

    def recording(ids: list[int]) -> tuple[list[str], list[float]]:
        tags, probs = (window or real_window)(ids)
        windows.append(list(tags))
        return tags, probs

    # setattr: ty rejects assigning a closure over a module function.
    setattr(redact, "_run_window", recording)  # noqa: B010
    try:
        spans = redact.detect_spans(REFERENCE_TEXTS, min_score=MIN_SCORE, org=True)
    finally:
        setattr(redact, "_run_window", real_window)  # noqa: B010
    return windows, [[[s["label"], s["start"], s["end"]] for s in seg] for seg in spans]


def _litert_window(interpreter: Any) -> Any:
    import numpy as np

    details = {d["name"]: d["index"] for d in interpreter.get_input_details()}
    output = interpreter.get_output_details()[0]["index"]

    def window(ids: list[int]) -> tuple[list[str], list[float]]:
        _, tokenizer, id2label = redact._load_runtime()
        seq = [tokenizer.bos_id, *ids, tokenizer.eos_id]
        padded = np.full((1, redact.SEQ), redact.PAD_ID, dtype=np.int32)
        mask = np.zeros((1, redact.SEQ), dtype=np.int32)
        padded[0, : len(seq)] = seq
        mask[0, : len(seq)] = 1
        for name, index in details.items():
            interpreter.set_tensor(index, padded if "input_ids" in name else mask)
        interpreter.invoke()
        logits = interpreter.get_tensor(output)[0, : len(seq)].astype(np.float32)
        probs = np.exp(logits - logits.max(axis=-1, keepdims=True))
        probs /= probs.sum(axis=-1, keepdims=True)
        picks = probs.argmax(axis=-1)
        tags = [id2label.get(int(i), "O") for i in picks][1:-1]
        return tags, [float(probs[n, i]) for n, i in enumerate(picks)][1:-1]

    return window


def verify(model_dir: Path, write: bool) -> int:
    import tflite_numpy

    for asset in redact.ASSETS:
        path = model_dir / asset["filename"]
        if not path.is_file() or _sha256(path) != asset["sha256"]:
            print(f"{path}: missing or not the pinned sha256; run --fetch first")
            return 1
    # A string import: LiteRT is absent from the typecheck environment by design.
    try:
        litert = importlib.import_module("ai_edge_litert.interpreter")
    except ImportError:
        print(
            "LiteRT is missing: uv run --with ai-edge-litert==2.2.0 build/redact_parity.py"
        )
        return 2

    config.DEBUGGING = False
    setattr(redact, "models_dir", lambda: model_dir)  # noqa: B010
    model_path = str(model_dir / redact.MODEL_FILENAME)

    ours = [(op[0], op[1], op[2]) for op in tflite_numpy.load_graph(model_path)["ops"]]
    plain = litert.Interpreter(
        model_path=model_path,
        experimental_op_resolver_type=litert.OpResolverType.BUILTIN_WITHOUT_DEFAULT_DELEGATES,
    )
    theirs = [
        (
            o["op_name"],
            tuple(int(i) for i in o["inputs"]),
            tuple(int(i) for i in o["outputs"]),
        )
        for o in plain._get_ops_details()
    ]
    graph_ok = ours == theirs
    print(
        f"graph: {len(ours)} ops, {'identical' if graph_ok else 'DIFFERENT'} to LiteRT's"
    )

    interpreter = litert.Interpreter(model_path=model_path, num_threads=4)
    interpreter.allocate_tensors()
    ref_windows, ref_spans = run_reference(_litert_window(interpreter))
    np_windows, np_spans = run_reference()
    pairs = [
        (a, b)
        for wa, wb in zip(ref_windows, np_windows, strict=True)
        for a, b in zip(wa, wb, strict=True)
    ]
    agree = sum(a == b for a, b in pairs)
    tags_ok = agree == len(pairs)
    spans_ok = ref_spans == np_spans
    print(f"tags: {agree}/{len(pairs)} agree over {len(ref_windows)} windows")
    print(
        f"spans: {'identical' if spans_ok else 'DIFFERENT'} ({sum(map(len, ref_spans))} total)"
    )
    for text, a, b in zip(REFERENCE_TEXTS, ref_spans, np_spans, strict=True):
        if a != b:
            print(f"  {text[:50]!r}\n    litert {a}\n    numpy  {b}")

    if not (graph_ok and tags_ok and spans_ok):
        print("numpy runtime disagrees with LiteRT; fixture left unchanged")
        return 1
    if write:
        FIXTURE.parent.mkdir(parents=True, exist_ok=True)
        fixture = {
            "model_tag": redact.MODEL_TAG,
            "assets": {a["filename"]: a["sha256"] for a in redact.ASSETS},
            "verified_with": f"ai-edge-litert {importlib.metadata.version('ai-edge-litert')}",
            "min_score": MIN_SCORE,
            "texts": REFERENCE_TEXTS,
            "windows": ref_windows,
            "spans": ref_spans,
        }
        FIXTURE.write_text(
            json.dumps(fixture, ensure_ascii=False, indent=1) + "\n", encoding="utf-8"
        )
        print(f"wrote {FIXTURE.relative_to(_ROOT)}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--model-dir", type=Path, help="default: the config dir's redact_models/"
    )
    parser.add_argument(
        "--write", action="store_true", help="record the fixture when all checks agree"
    )
    parser.add_argument(
        "--fetch",
        type=Path,
        metavar="DIR",
        help="download the pinned assets into DIR and exit",
    )
    args = parser.parse_args()
    if args.fetch:
        setattr(redact, "models_dir", lambda: args.fetch)  # noqa: B010
        args.fetch.mkdir(parents=True, exist_ok=True)
        return 0 if redact.download() else 1
    return verify(args.model_dir or redact.models_dir(), args.write)


if __name__ == "__main__":
    sys.exit(main())
