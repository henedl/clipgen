"""Guards for the tests/perf bench scripts' compare and fail-on logic."""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path

import pytest

_spec = importlib.util.spec_from_file_location(
    "scan_bench", Path(__file__).parent / "perf" / "scan_bench.py"
)
assert _spec is not None and _spec.loader is not None  # a checked-in file
scan_bench = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(scan_bench)


def test_regressions_flags_a_failed_run():
    rows = {"color": {"callback_s": 0.0}, "ocr": {"callback_s": 1.2}}
    base = {"color": {"callback_s": 2.0}, "ocr": {"callback_s": 1.0}}
    hit = dict(scan_bench.regressions(rows, base, "callback_s", 10))
    assert hit["color"] == float("inf")
    assert hit["ocr"] == pytest.approx(20.0)


def test_regressions_ignores_a_missing_baseline():
    assert (
        scan_bench.regressions({"x": {"callback_s": 0.0}}, {}, "callback_s", 10) == []
    )


def test_baseline_rows_refuses_another_fixture_length(tmp_path):
    snap = tmp_path / "base.json"
    snap.write_text(json.dumps({"meta": {"duration": 15}, "tools": {"a": {}}}))
    ap = argparse.ArgumentParser()
    assert scan_bench.baseline_rows(ap, snap, "tools", 15) == {"a": {}}
    with pytest.raises(SystemExit):
        scan_bench.baseline_rows(ap, snap, "tools", 120)


def test_probe_duration_survives_a_missing_ffprobe(monkeypatch, tmp_path):
    def missing(*a, **k):
        raise FileNotFoundError("ffprobe")

    monkeypatch.setattr(scan_bench.subprocess, "run", missing)
    assert scan_bench.probe_duration(tmp_path / "v.mp4") is None
