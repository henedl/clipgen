"""Unit tests for clip_bench's row reduction and expected-work checks.

The bench itself is a subprocess driver (never collected — no test_ prefix);
the reduction is the part that silently rots if the profiling labels it
reads change shape, so it is pinned against a real ``profiling.export()``.
"""

import pytest

import bench_common as bc
import bench_fixtures as bf
import clip_bench
import config
import profiling

SEED = {
    "pipeline.clip": (6.702, 13, 0.8236),
    "titlecard.wrap": (5.987, 13, 0.7398),
    "ffmpeg.run.card": (3.9, 14, 0.4105),
    "ffmpeg.run.concat": (1.22, 13, 0.12),
    "ffmpeg.run.cut": (0.85, 13, 0.09),
    "pipeline.pool_wall": (1.879, 1, None),
    "ffprobe.run": (0.543, 14, 0.0411),
    "titlecard.copy": (0.0, 13, None),
}


@pytest.fixture
def make_doc(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(config, "PROFILE_DEEP", "")
    monkeypatch.setattr(profiling, "environment", lambda: {"python": "3.12"})

    def build(seed):
        profiling.reset()
        for label, (secs, n, peak) in seed.items():
            profiling.add(label, secs, n, peak=peak)
        return profiling.export(reset=True)

    yield build
    profiling.reset()


def test_summarize_reduces_export_to_row(make_doc):
    row = clip_bench.summarize(make_doc(SEED))
    assert row["clips"] == 13
    assert row["clip_s"] == pytest.approx(6.702)
    assert row["pool_wall_s"] == pytest.approx(1.879)
    assert 3.5 < row["parallelism"] < 3.6
    assert row["ffmpeg_n"] == 40
    assert row["ffmpeg_s"] == pytest.approx(5.970)
    assert row["ffmpeg_kinds"]["card"] == {"s": pytest.approx(3.9), "n": 14}
    assert set(row["ffmpeg_kinds"]) == {"card", "concat", "cut"}
    assert row["cards_s"] == pytest.approx(5.987)
    assert row["cards_copy"] == 13
    assert row["cards_reencode"] == 0


def test_summarize_handles_missing_labels(make_doc):
    row = clip_bench.summarize(make_doc({"ffprobe.run": (0.1, 1, 0.1)}))
    assert row["clips"] == 0
    assert row["parallelism"] == 0.0


def test_clip_cell_range_matches_legacy_windows_on_long_fixture():
    assert clip_bench.clip_cell_range(0, 120) == "0:02-0:05"
    assert clip_bench.clip_cell_range(4, 120) == "0:14-0:21"


def test_clip_cell_range_clamps_to_short_fixture():
    assert clip_bench.clip_cell_range(4, 15) == "0:07-0:14"


def test_check_work_requires_clip_count_and_files(tmp_path, make_doc):
    short = make_doc({"pipeline.clip": (1.0, 13, 0.1)})
    reasons = clip_bench.check_work("clips", short, tmp_path)
    assert any("13 clips cut" in r for r in reasons)
    assert any("0 output files" in r for r in reasons)

    full = make_doc({"pipeline.clip": (1.0, clip_bench.EXPECTED_CLIPS, 0.1)})
    for i in range(clip_bench.EXPECTED_CLIPS):
        (tmp_path / f"clip{i}.mp4").write_bytes(b"x")
    assert clip_bench.check_work("clips", full, tmp_path) == []
    # A carded scenario without card work is disabled titlecards, not speed.
    assert clip_bench.check_work("clips-cards", full, tmp_path) == [
        "0 clips carded, expected 20"
    ]
    carded = make_doc(
        {
            "pipeline.clip": (1.0, clip_bench.EXPECTED_CLIPS, 0.1),
            "titlecard.wrap": (2.0, clip_bench.EXPECTED_CLIPS, 0.2),
        }
    )
    assert clip_bench.check_work("clips-cards", carded, tmp_path) == []
    full = carded
    # A reel needs one file; its `_reel_part_*` segments are not outputs.
    reel_dir = tmp_path / "reel"
    reel_dir.mkdir()
    (reel_dir / "_reel_part_0.mp4").write_bytes(b"x")
    assert clip_bench.check_work("reel", full, reel_dir) == [
        "0 output files, expected 1"
    ]
    (reel_dir / "reel.mp4").write_bytes(b"x")
    assert clip_bench.check_work("reel", full, reel_dir) == []


def test_run_scenario_invalidates_short_work(tmp_path, monkeypatch, make_doc):
    doc = make_doc({"pipeline.clip": (1.0, 5, 0.1)})

    def fake_run(args, out_dir, *, required_labels, timeout=0):
        assert required_labels == ["pipeline.clip"]
        assert "--no-input" in args
        return {
            "valid": True,
            "reasons": [],
            "elapsed_s": 2.0,
            "returncode": 0,
            "doc": doc,
            "output": "",
        }

    monkeypatch.setattr(bc, "run_clipgen", fake_run)
    result = clip_bench.run_scenario("clips", tmp_path / "in", tmp_path / "out")
    assert result["valid"] is False
    assert result["reasons"][0].startswith("5 clips cut")


def test_sheet_cells_spread_long_clips_across_the_source():
    cells = clip_bench.sheet_cells(3600, 180)
    assert cells[0] == "0:01-3:01"
    assert cells[clip_bench.EXPECTED_CLIPS - 1] == "54:08-57:08"
    assert len(cells) == clip_bench.DATA_ROWS
    assert clip_bench.clip_media_seconds(cells) == 180 * clip_bench.EXPECTED_CLIPS
    # Without --clip-seconds the legacy first-minute windows stay.
    assert clip_bench.sheet_cells(120, None)[0] == "0:02-0:05"


def test_clip_cells_reach_past_an_hour():
    cells = bf.clip_cells(4, duration=10_800, clip_seconds=60)
    assert cells[-1] == "2:14:14-2:15:14"
    assert bf.format_hms(59) == "0:59" and bf.format_hms(3600) == "1:00:00"


def test_compress_scenario_requires_compress_passes(tmp_path, make_doc):
    for i in range(clip_bench.EXPECTED_CLIPS):
        (tmp_path / f"clip{i}.mp4").write_bytes(b"x")
    n = clip_bench.EXPECTED_CLIPS
    skipped = make_doc({"pipeline.clip": (1.0, n, 0.1)})
    assert (
        "compress passes"
        in clip_bench.check_work("clips-compress", skipped, tmp_path)[0]
    )
    ran = make_doc(
        {"pipeline.clip": (1.0, n, 0.1), "ffmpeg.run.compress": (9.0, 2 * n, 0.5)}
    )
    assert clip_bench.check_work("clips-compress", ran, tmp_path) == []


def test_run_scenario_passes_overrides_only_where_needed(tmp_path, monkeypatch):
    seen = []

    def fake_run(args, out_dir, *, required_labels, timeout=0, overrides=None):
        seen.append(overrides)
        return {"valid": False, "reasons": ["x"], "doc": {}, "output": ""}

    monkeypatch.setattr(bc, "run_clipgen", fake_run)
    clip_bench.run_scenario("clips", tmp_path / "in", tmp_path / "out")
    clip_bench.run_scenario("clips-reencode", tmp_path / "in", tmp_path / "out")
    assert seen == [None, {"REENCODING": True}]
