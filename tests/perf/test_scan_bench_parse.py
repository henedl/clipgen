"""Unit tests for scan_bench's row reduction, validity and fixture checks.

The bench itself is a subprocess driver (never collected — no test_ prefix);
the reduction is the part that silently rots if the profiling labels it
reads change shape, so it is pinned against a real ``profiling.export()``.
"""

import json
from pathlib import Path
from types import SimpleNamespace

import pytest

import bench_common as bc
import config
import profiling
import scan_bench


@pytest.fixture
def make_doc(monkeypatch):
    """A real export document seeded with labels; no git/ffmpeg probes."""
    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(profiling, "environment", lambda: {"python": "3.12"})

    def build(seed, *, deep=""):
        monkeypatch.setattr(config, "PROFILE_DEEP", deep)
        profiling.reset()
        for label, (secs, n, peak) in seed.items():
            profiling.add(label, secs, n, peak=peak)
        return profiling.export(reset=True)

    yield build
    profiling.reset()


SEED = {
    "scan.callback.color": (3.688, 962, 0.007),
    "scan.decode_wait": (0.798, 962, None),
    "scan.fast_filter": (0.012, 962, None),
    "heatmap.gifs": (1.008, 1, None),
    "heatmap.grid_layers": (0.001, 1, None),
    "ffprobe.run": (0.039, 1, None),
}


def _fake_result(doc, *, valid=True, reasons=(), elapsed=5.0):
    return {
        "valid": valid,
        "reasons": list(reasons),
        "elapsed_s": elapsed,
        "returncode": 0 if valid else 1,
        "doc": doc,
        "output": "",
    }


def test_summarize_reduces_export_to_row(make_doc):
    doc = make_doc(SEED)
    row = scan_bench.summarize("color", doc)
    assert row["callback_s"] == pytest.approx(3.688)
    assert row["frames"] == 962
    assert row["callback_avg_ms"] == pytest.approx(3.834, abs=0.01)
    assert row["decode_s"] == pytest.approx(0.798)
    assert row["filter_s"] == pytest.approx(0.012)
    assert row["heatmap_s"] == pytest.approx(1.009)
    assert row["peak_rss_mb"] == doc["peak_rss_mb"]


def test_summarize_handles_missing_labels(make_doc):
    row = scan_bench.summarize("flow", make_doc({}))
    assert row["callback_s"] == 0.0
    assert row["frames"] == 0
    assert row["callback_avg_ms"] == 0.0


def test_task_completed_reads_manifest(tmp_path):
    assert scan_bench.task_completed(tmp_path) is False
    (tmp_path / "clipgen.json").write_text(
        json.dumps({"screenspace": {"tasks": [{"status": "failed"}]}})
    )
    assert scan_bench.task_completed(tmp_path) is False
    (tmp_path / "clipgen.json").write_text(
        json.dumps({"screenspace": {"tasks": [{"status": "completed"}]}})
    )
    assert scan_bench.task_completed(tmp_path) is True


def test_run_tool_seeds_region_and_requires_completion(tmp_path, monkeypatch, make_doc):
    doc = make_doc({"scan.callback.shape": (1.0, 10, 0.2)})
    captured = {}

    def fake_run(args, out_dir, *, required_labels, timeout=0):
        captured["args"] = args
        captured["required"] = required_labels
        manifest = json.loads((out_dir / "clipgen.json").read_text())
        assert manifest["screenspace"]["regions"]["bench"] == scan_bench.BENCH_REGION
        return _fake_result(doc)

    monkeypatch.setattr(bc, "run_clipgen", fake_run)
    result = scan_bench.run_tool("shape", tmp_path / "in", tmp_path / "out", 0.1)
    assert captured["args"][:4] == ["--ss-task", "shape", "P01", "bench"]
    assert captured["required"] == ["scan.callback.shape"]
    # The seeded manifest has no completed task, so the run is not valid.
    assert result["valid"] is False
    assert "task did not complete" in result["reasons"]


def test_run_tool_accepts_a_completed_task(tmp_path, monkeypatch, make_doc):
    doc = make_doc({"scan.callback.color": (1.0, 10, 0.2)})

    def fake_run(args, out_dir, *, required_labels, timeout=0):
        (out_dir / "clipgen.json").write_text(
            json.dumps({"screenspace": {"tasks": [{"status": "completed"}]}})
        )
        return _fake_result(doc)

    monkeypatch.setattr(bc, "run_clipgen", fake_run)
    result = scan_bench.run_tool("color", tmp_path / "in", tmp_path / "out", 0.1)
    assert result["valid"] is True


def test_run_tool_deep_is_diagnostic_not_invalid(tmp_path, monkeypatch, make_doc):
    doc = make_doc({"scan.callback.color": (1.0, 10, 0.2)}, deep="scan.callback")

    def fake_run(args, out_dir, *, required_labels, timeout=0):
        assert "--profile-deep" in args
        (out_dir / "clipgen.json").write_text(
            json.dumps({"screenspace": {"tasks": [{"status": "completed"}]}})
        )
        return _fake_result(doc, valid=False, reasons=["deep run"])

    monkeypatch.setattr(bc, "run_clipgen", fake_run)
    result = scan_bench.run_tool(
        "color", tmp_path / "in", tmp_path / "out", 0.1, deep=True
    )
    assert result["valid"] is True


def test_ensure_fixture_rebuilds_on_spec_mismatch(tmp_path, monkeypatch):
    video = tmp_path / "bench_P01.mp4"
    video.write_bytes(b"old")
    probes = iter([dict(scan_bench.fixture_spec(120))])
    monkeypatch.setattr(bc, "probe_fixture", lambda _p: next(probes))
    calls = []

    def fake_run(cmd, **kwargs):
        calls.append(cmd[0])
        video.write_bytes(b"new")
        return SimpleNamespace(stdout="", stderr="", returncode=0)

    monkeypatch.setattr(scan_bench.subprocess, "run", fake_run)
    assert scan_bench.ensure_fixture(tmp_path, 15) == video
    assert video.read_bytes() == b"new"
    assert calls == ["ffmpeg"]


def test_ensure_fixture_keeps_a_matching_file(tmp_path, monkeypatch):
    video = tmp_path / "bench_P01.mp4"
    video.write_bytes(b"keep")
    monkeypatch.setattr(bc, "probe_fixture", lambda _p: scan_bench.fixture_spec(15))

    def fail(*_a, **_k):
        raise AssertionError("ffmpeg must not run")

    monkeypatch.setattr(scan_bench.subprocess, "run", fail)
    scan_bench.ensure_fixture(tmp_path, 15)
    assert video.read_bytes() == b"keep"


def test_ensure_fixture_rebuilds_a_wrong_size(tmp_path, monkeypatch):
    """Duration alone is not identity: a 640x360 leftover must go."""
    video = tmp_path / "bench_P01.mp4"
    video.write_bytes(b"old")
    probed = dict(scan_bench.fixture_spec(15), width=640, height=360)
    monkeypatch.setattr(bc, "probe_fixture", lambda _p: probed)
    ran = []
    monkeypatch.setattr(
        scan_bench.subprocess,
        "run",
        lambda cmd, **k: ran.append(cmd[0]) or SimpleNamespace(returncode=0),
    )
    scan_bench.ensure_fixture(tmp_path, 15)
    assert ran == ["ffmpeg"]


def test_shape_is_in_default_sweep():
    assert "shape" in scan_bench.DEFAULT_TOOLS
    assert scan_bench.TOOL_FLAGS["shape"] == [
        "--ss-reference-timestamp",
        "1",
        "--ss-threshold",
        "0.55",
    ]


def test_build_row_keeps_every_sample(make_doc):
    docs = [make_doc({"scan.callback.color": (s, 100, 0.1)}) for s in (1.0, 1.2, 5.0)]
    results = [_fake_result(d, elapsed=e) for d, e in zip(docs, (3.0, 3.1, 9.0))]
    row = bc.build_row(
        results, lambda d: scan_bench.summarize("color", d), scan_bench.METRICS
    )
    assert row["valid"] is True
    assert len(row["samples"]) == 3
    assert row["stats"]["elapsed_s"]["median"] == pytest.approx(3.1)
    assert row["stats"]["callback_s"]["median"] == pytest.approx(1.2)
    assert Path  # keep the import used for future path assertions
