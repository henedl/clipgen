"""Unit tests for clip_bench's row reduction and best-run keeping.

The bench itself is a subprocess driver (never collected — no test_ prefix);
the reduction is the part that silently rots if the `profile |` labels it
reads change shape, so it is pinned here against a verbatim report snippet.
"""

import clip_bench
from scan_bench import parse_profile

REPORT = """\
profile | pipeline.clip                       6.702s  n=13  avg=515.5ms  max=823.6ms
profile | titlecard.wrap                      5.987s  n=13  avg=460.5ms  max=739.8ms
profile | ffmpeg.run                          5.970s  n=27  avg=221.1ms  max=410.5ms
profile | pipeline.pool_wall                  1.879s  n=1  avg=1879.0ms  max=1879.0ms
profile | ffprobe.run                         0.543s  n=14  avg=38.8ms  max=41.1ms
profile | titlecard.copy                      0.000s  n=13
profile | peak_rss                             65.1MB
"""


def test_summarize_reduces_report_to_row():
    row = clip_bench.summarize(parse_profile(REPORT))
    assert row["clips"] == 13
    assert row["clip_s"] == 6.702
    assert row["pool_wall_s"] == 1.879
    assert 3.5 < row["parallelism"] < 3.6
    assert row["ffmpeg_n"] == 27
    assert row["cards_s"] == 5.987
    assert row["cards_copy"] == 13
    assert row["cards_reencode"] == 0
    assert row["peak_rss_mb"] == 65.1


def test_summarize_handles_missing_labels():
    row = clip_bench.summarize(parse_profile("profile | ffprobe.run 0.1s n=1"))
    assert row["clips"] == 0
    assert row["parallelism"] == 0.0


def test_keep_best_prefers_fastest_successful_run():
    ok = clip_bench.summarize(parse_profile(REPORT))
    failed = clip_bench.summarize(parse_profile(""))
    assert clip_bench.keep_best(None, failed) is failed
    assert clip_bench.keep_best(failed, ok) is ok
    assert clip_bench.keep_best(ok, failed) is ok
    slower = dict(ok, clip_s=ok["clip_s"] + 1)
    assert clip_bench.keep_best(ok, slower) is ok
