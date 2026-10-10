"""In-process growth check for pure functions whose input grows with session length.

The media benches time whole runs; this times the pure steps inside them
over synthetic inputs of growing size, so a quadratic step shows up as a
fitted exponent in seconds, not as a three-hour session that never ends. No
media, no subprocesses, no models.

Usage (from the repo root):

    uv run python tests/perf/scale_bench.py                    # every function
    uv run python tests/perf/scale_bench.py --only cluster --sizes 500,1000,2000
    uv run python tests/perf/scale_bench.py --project 6000     # extrapolate

Functions:

- ``cluster``  ``speakers.cluster`` over N random unit embeddings: the
  average-linkage merge loop the speaker pass runs once per transcript.
  Random vectors are the worst case — nothing is close, so every merge runs.
  N is the unit count after adjacent-segment merging; fast turn-taking keeps
  it near the segment count (~2000 per hour).
- ``redact-map``  the span-to-segment mapping in ``redact.detect_spans``
  over N segments with one span per five, the model stubbed out.
- ``friction``  ``friction.score_segments`` over N segments (a linear
  control: its exponent should read ~1).

Each size keeps the fastest of ``--repeat`` timings. The fit's exponent is
taken between the two largest sizes; ``--project N`` extrapolates the time
at N along that power law, labelled as a projection.
"""

from __future__ import annotations

import argparse
import sys
import time
from collections.abc import Callable
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "source"))

import bench_common as bc

DEFAULT_SIZES = (250, 500, 1000, 2000)
EMBED_DIM = 192
_PHRASES = (
    "okay so I am looking at the screen and I cannot find the button",
    "um wait where is it I guess I would tap here but it looks disabled",
)
WORDS = " ".join(_PHRASES).split()


def _segments(n: int) -> list[dict[str, Any]]:
    out = []
    for i in range(n):
        words = [WORDS[(i * 7 + k) % len(WORDS)] for k in range(12)]
        out.append({"id": str(i), "start": i * 3.0, "text": " ".join(words)})
    return out


def prepare_cluster(n: int) -> Callable[[], Any]:
    import numpy as np

    import speakers

    rng = np.random.default_rng(n)
    emb = rng.standard_normal((n, EMBED_DIM)).astype(np.float32)
    return lambda: speakers.cluster(emb, max_speakers=4)


def prepare_redact_map(n: int) -> Callable[[], Any]:
    import redact

    texts = [s["text"] for s in _segments(n)]
    starts, pos = [], 0
    for text in texts:
        starts.append(pos)
        pos += len(text) + len(redact.SEGMENT_JOIN)
    spans = [
        {
            "label": "GIVEN_NAME",
            "start": starts[i] + 3,
            "end": starts[i] + 8,
            "score": 0.9,
        }
        for i in range(0, n, 5)
    ]

    stubs: dict[str, Any] = {
        "_load_runtime": lambda: (None, None, {}),
        "_drop_runtime": lambda: None,
        "_document_spans": lambda *_a, **_k: spans,
        "label_families": lambda _m: set(),
    }

    def run() -> Any:
        saved = {name: getattr(redact, name) for name in stubs}
        for name, stub in stubs.items():
            setattr(redact, name, stub)
        try:
            return redact.detect_spans(texts, min_score=0.6, org=False)
        finally:
            for name, original in saved.items():
                setattr(redact, name, original)

    return run


def prepare_friction(n: int) -> Callable[[], Any]:
    import friction

    segments = _segments(n)
    return lambda: friction.score_segments(segments)


FUNCTIONS: dict[str, Callable[[int], Callable[[], Any]]] = {
    "cluster": prepare_cluster,
    "redact-map": prepare_redact_map,
    "friction": prepare_friction,
}


def time_once(fn: Callable[[], Any], repeat: int) -> float:
    """Fastest of *repeat* calls: the least noisy estimate of the work itself."""
    best = float("inf")
    for _ in range(max(1, repeat)):
        t0 = time.perf_counter()
        fn()
        best = min(best, time.perf_counter() - t0)
    return best


def project(points: list[tuple[float, float]], target: float) -> float | None:
    """Time at *target* along the power law through the two largest points."""
    fit = bc.fit_scaling(points)
    if fit["exponent"] is None:
        return None
    x2, y2 = max(points)
    return y2 * (target / x2) ** fit["exponent"]


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument(
        "--only",
        default=",".join(FUNCTIONS),
        help=f"comma-separated functions (default: {', '.join(FUNCTIONS)})",
    )
    ap.add_argument(
        "--sizes",
        default=",".join(str(n) for n in DEFAULT_SIZES),
        help="comma-separated input sizes N",
    )
    ap.add_argument("--repeat", type=int, default=3, help="timings per size")
    ap.add_argument(
        "--project",
        type=int,
        default=6000,
        help="extrapolate each function to this N (default: 6000, ~3 h of segments)",
    )
    args = ap.parse_args()
    names = [n.strip() for n in args.only.split(",") if n.strip()]
    unknown = [n for n in names if n not in FUNCTIONS]
    if unknown:
        print(f"unknown functions: {', '.join(unknown)} (know: {', '.join(FUNCTIONS)})")
        return bc.EXIT_USAGE
    sizes = sorted({int(v) for v in args.sizes.split(",") if v.strip()})
    if len(sizes) < 2:
        print("--sizes needs at least two values")
        return bc.EXIT_USAGE

    points: dict[str, list[tuple[float, float]]] = {}
    for name in names:
        for n in sizes:
            fn = FUNCTIONS[name](n)
            seconds = time_once(fn, args.repeat if n <= 1000 else 1)
            points.setdefault(name, []).append((float(n), seconds))
            print(f"  {name:<11} N={n:<6} {seconds:.4f}s")

    bc.print_scaling(points, unit="1k items", per=1000)
    print()
    for name, pts in points.items():
        guess = project(pts, args.project)
        if guess is not None:
            print(f"{name:<16}projected at N={args.project}: {guess:.2f}s")
    return bc.EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
