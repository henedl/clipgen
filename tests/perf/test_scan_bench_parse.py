"""Unit tests for scan_bench's profile-report parser and row reduction.

The bench itself is a subprocess driver (never collected — no test_ prefix);
the parser is the part that silently rots if the `profile |` line shape
changes, so it is pinned here against a verbatim report snippet.
"""

import json
from types import SimpleNamespace

import scan_bench

REPORT = """\
profile | scan color bench_P01.mp4: decode_wait=0.798s/n=962  fast_filter=0.000s/n=962  callback=3.688s/n=962
profile | scan.callback.color                 3.688s  n=962  avg=3.8ms  max=7.0ms
profile | scan.decode_wait                    0.798s  n=962  avg=0.8ms
profile | heatmap.gifs                        1.008s  n=1  avg=1007.7ms  max=1007.7ms
profile | heatmap.grid_layers                 0.001s  n=1  avg=0.5ms  max=0.5ms
profile | ffprobe.run                         0.039s  n=1  avg=39.0ms  max=39.0ms
profile | peak_rss                            104.7MB
"""


def test_parse_profile_extracts_labels_and_rss():
    parsed = scan_bench.parse_profile(REPORT)
    assert parsed["scan.callback.color"] == {"seconds": 3.688, "n": 962}
    assert parsed["scan.decode_wait"]["seconds"] == 0.798
    assert parsed["peak_rss"]["mb"] == 104.7
    # The per-scan summary line is not a totals label and must not parse.
    assert "scan" not in parsed


def test_summarize_reduces_to_comparison_row():
    row = scan_bench.summarize("color", scan_bench.parse_profile(REPORT))
    assert row["callback_s"] == 3.688
    assert row["frames"] == 962
    assert abs(row["callback_avg_ms"] - 3.834) < 0.01
    assert row["decode_s"] == 0.798
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
        return SimpleNamespace(stdout=report, stderr="")

    monkeypatch.setattr(scan_bench.subprocess, "run", fake_run)
    parsed = scan_bench.run_tool("shape", tmp_path / "in", tmp_path / "out", 0.1)
    manifest_path = tmp_path / "out" / "bench-shape" / "clipgen.json"
    manifest = json.loads(manifest_path.read_text())

    assert parsed["scan.callback.shape"]["seconds"] == 3.688
    assert captured["cmd"][3:7] == ["--ss-task", "shape", "P01", "bench"]
    assert manifest["screenspace"]["regions"]["bench"] == scan_bench.BENCH_REGION
