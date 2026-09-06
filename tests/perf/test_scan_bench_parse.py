"""Unit tests for scan_bench's profile-report parser and row reduction.

The bench itself is a subprocess driver (never collected — no test_ prefix);
the parser is the part that silently rots if the `profile |` line shape
changes, so it is pinned here against a verbatim report snippet.
"""

import json
from types import SimpleNamespace

import pytest

import scan_bench

REPORT = """\
profile | scan color bench_P01.mp4: decode_wait=0.798s/n=962  fast_filter=0.000s/n=962  callback=3.688s/n=962
profile | scan.callback.color                 3.688s  n=962  avg=3.8ms  max=7.0ms
profile | scan.decode_wait                    0.798s  n=962  avg=0.8ms
profile | scan.fast_filter                    0.012s  n=962  avg=0.0ms
profile | heatmap.gifs                        1.008s  n=1  avg=1007.7ms  max=1007.7ms
profile | heatmap.grid_layers                 0.001s  n=1  avg=0.5ms  max=0.5ms
profile | ffprobe.run                         0.039s  n=1  avg=39.0ms  max=39.0ms
profile | peak_rss                            104.7MB
"""

SPACED = """\
profile | route /api/models                   0.120s  n=3  avg=40.0ms  max=80.0ms  bytes=1.2MB  first=80.0ms
profile | manifest.load clips                 0.010s  n=2  avg=5.0ms  bytes=256B
profile | stream /studio/api/generate         8.400s  n=1  avg=8400.0ms  bytes=4.0KB
"""


def test_parse_profile_extracts_labels_and_rss():
    parsed = scan_bench.parse_profile(REPORT)
    color = parsed["scan.callback.color"]
    assert color["seconds"] == 3.688
    assert color["n"] == 962
    assert color["avg"] == pytest.approx(0.0038)
    assert color["max"] == pytest.approx(0.007)
    assert parsed["scan.decode_wait"]["seconds"] == 0.798
    assert parsed["peak_rss"]["mb"] == 104.7
    # Per-scan summary lines are not totals and must not parse.
    assert "scan" not in parsed


def test_parse_profile_keeps_spaced_labels_and_extra_fields():
    parsed = scan_bench.parse_profile(SPACED)
    route = parsed["route /api/models"]
    assert route["seconds"] == 0.120
    assert route["n"] == 3
    assert route["max"] == pytest.approx(0.080)
    assert route["first"] == pytest.approx(0.080)
    assert route["bytes"] == pytest.approx(1.2 * 1024 * 1024)
    assert parsed["manifest.load clips"]["n"] == 2
    assert parsed["manifest.load clips"]["bytes"] == 256
    assert parsed["stream /studio/api/generate"]["bytes"] == pytest.approx(4.0 * 1024)


def test_parse_profile_round_trips_report(monkeypatch, capsys):
    import config
    import profiling

    monkeypatch.setattr(config, "PROFILING", True)
    profiling.reset()
    profiling.add("route /api/models", 0.08, nbytes=1024)
    profiling.add("route /api/models", 0.02, nbytes=512)
    profiling.add("scan.callback.color", 1.5, 100, peak=0.02)
    profiling.report()
    parsed = scan_bench.parse_profile(capsys.readouterr().out)
    assert parsed["route /api/models"]["n"] == 2
    assert parsed["route /api/models"]["bytes"] > 0
    assert parsed["scan.callback.color"]["seconds"] == pytest.approx(1.5)
    assert parsed["scan.callback.color"]["max"] == pytest.approx(0.02)
    profiling.reset()


def test_summarize_reduces_to_comparison_row():
    row = scan_bench.summarize("color", scan_bench.parse_profile(REPORT))
    assert row["callback_s"] == 3.688
    assert row["frames"] == 962
    assert abs(row["callback_avg_ms"] - 3.834) < 0.01
    assert row["decode_s"] == 0.798
    assert row["filter_s"] == 0.012
    assert abs(row["heatmap_s"] - 1.009) < 1e-9
    assert row["peak_rss_mb"] == 104.7


def test_keep_best_prefers_success_over_failed_first_run():
    failed = {"callback_s": 0.0, "frames": 0}
    ok = {"callback_s": 0.5, "frames": 962}
    faster = {"callback_s": 0.4, "frames": 962}
    # A failed first run must be replaced by any later success...
    assert scan_bench.keep_best(scan_bench.keep_best(None, failed), ok) is ok
    # ...a later failure never displaces a success...
    assert scan_bench.keep_best(scan_bench.keep_best(None, ok), failed) is ok
    # ...and among successes the minimum callback wins.
    assert scan_bench.keep_best(scan_bench.keep_best(None, ok), faster) is faster
    assert scan_bench.keep_best(scan_bench.keep_best(None, faster), ok) is faster


def test_summarize_handles_missing_labels():
    row = scan_bench.summarize("flow", scan_bench.parse_profile(""))
    assert row["callback_s"] == 0.0
    assert row["frames"] == 0
    assert row["callback_avg_ms"] == 0.0
    assert row["filter_s"] == 0.0


def test_regressions_flags_callback_increase():
    rows = {"color": {"callback_s": 1.2}}
    base = {"color": {"callback_s": 1.0}}
    hit = scan_bench.regressions(rows, base, "callback_s", 10)
    assert hit == [("color", pytest.approx(20.0))]
    assert scan_bench.regressions(rows, base, "callback_s", 25) == []


def test_ensure_fixture_rebuilds_on_duration_mismatch(tmp_path, monkeypatch):
    video = tmp_path / "bench_P01.mp4"
    video.write_bytes(b"old")
    calls = []

    def fake_run(cmd, **kwargs):
        calls.append(cmd[0])
        if cmd[0] == "ffprobe":
            return SimpleNamespace(stdout="120.0\n", stderr="", returncode=0)
        video.write_bytes(b"new")
        return SimpleNamespace(stdout="", stderr="", returncode=0)

    monkeypatch.setattr(scan_bench.subprocess, "run", fake_run)
    out = scan_bench.ensure_fixture(tmp_path, 15)
    assert out == video
    assert video.read_bytes() == b"new"
    assert calls[0] == "ffprobe"
    assert "ffmpeg" in calls


def test_ensure_fixture_keeps_matching_duration(tmp_path, monkeypatch):
    video = tmp_path / "bench_P01.mp4"
    video.write_bytes(b"keep")

    def fake_run(cmd, **kwargs):
        assert cmd[0] == "ffprobe"
        return SimpleNamespace(stdout="15.0\n", stderr="", returncode=0)

    monkeypatch.setattr(scan_bench.subprocess, "run", fake_run)
    scan_bench.ensure_fixture(tmp_path, 15)
    assert video.read_bytes() == b"keep"


def test_shape_is_in_default_sweep():
    assert "shape" in scan_bench.DEFAULT_TOOLS
    assert scan_bench.TOOL_FLAGS["shape"] == [
        "--ss-reference-timestamp",
        "1",
        "--ss-threshold",
        "0.55",
    ]


def test_shape_run_seeds_region(tmp_path, monkeypatch):
    captured = {}
    report = REPORT.replace("color", "shape")

    def fake_run(cmd, **kwargs):
        captured["cmd"] = cmd
        return SimpleNamespace(stdout=report, stderr="", returncode=0)

    monkeypatch.setattr(scan_bench.subprocess, "run", fake_run)
    parsed = scan_bench.run_tool("shape", tmp_path / "in", tmp_path / "out", 0.1)
    manifest_path = tmp_path / "out" / "bench-shape" / "clipgen.json"
    manifest = json.loads(manifest_path.read_text())

    assert parsed["scan.callback.shape"]["seconds"] == 3.688
    assert captured["cmd"][3:7] == ["--ss-task", "shape", "P01", "bench"]
    assert manifest["screenspace"]["regions"]["bench"] == scan_bench.BENCH_REGION
