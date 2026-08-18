"""Unit tests for scan_bench's profile-report parser and row reduction.

The bench itself is a subprocess driver (never collected — no test_ prefix);
the parser is the part that silently rots if the `profile |` line shape
changes, so it is pinned here against a verbatim report snippet.
"""

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


def test_summarize_handles_missing_labels():
    row = scan_bench.summarize("flow", scan_bench.parse_profile(""))
    assert row["callback_s"] == 0.0
    assert row["frames"] == 0
    assert row["callback_avg_ms"] == 0.0
