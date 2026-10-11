"""Guards for tests/perf/bench_common.py: validity, aggregation, comparison."""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
from types import SimpleNamespace

import pytest

_spec = importlib.util.spec_from_file_location(
    "bench_common", Path(__file__).parent / "perf" / "bench_common.py"
)
assert _spec is not None and _spec.loader is not None  # a checked-in file
bc = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(bc)


def _row(valid=True, *, elapsed=1.0, count=10, reasons=()):
    return {
        "valid": valid,
        "reasons": list(reasons),
        "samples": [{"elapsed_s": elapsed, "frames": count}] if valid else [],
        "stats": {
            "elapsed_s": bc.aggregate([{"elapsed_s": elapsed}], "elapsed_s"),
            "frames": bc.aggregate([{"frames": count}], "frames"),
        },
    }


# ---------- aggregation --------------------------------------------------------


def test_aggregate_median_is_robust_to_an_outlier():
    samples = [{"t": 1.0}, {"t": 1.1}, {"t": 9.0}]
    agg = bc.aggregate(samples, "t")
    assert agg["median"] == pytest.approx(1.1)
    assert agg["min"] == 1.0
    assert agg["max"] == 9.0
    assert agg["mad"] == pytest.approx(0.1)
    assert agg["n"] == 3


def test_aggregate_empty():
    assert bc.aggregate([], "t") == {
        "median": 0.0,
        "min": 0.0,
        "max": 0.0,
        "mad": 0.0,
        "n": 0,
    }


def test_build_row_one_bad_repetition_voids_the_row():
    good = {"valid": True, "reasons": [], "elapsed_s": 1.0, "doc": {}}
    bad = {"valid": False, "reasons": ["exit 1"], "elapsed_s": 0.1, "doc": {}}
    row = bc.build_row([good, bad, good], lambda d: {"x": 1}, ("elapsed_s",))
    assert row["valid"] is False
    assert row["reasons"] == ["exit 1"]
    assert len(row["samples"]) == 2  # the good samples are kept for inspection


# ---------- comparison ---------------------------------------------------------


def test_compare_flags_invalid_rows_never_as_speedups():
    findings = bc.compare(
        {"a": _row(False, reasons=["exit 1"])},
        {"a": _row(elapsed=2.0)},
        metric="elapsed_s",
        pct=10,
        abs_s=None,
        count_key="frames",
    )
    assert findings == [("invalid", "a", "exit 1")]


def test_compare_missing_baseline_is_incomparable():
    findings = bc.compare(
        {"a": _row()}, {}, metric="elapsed_s", pct=10, abs_s=None, count_key="frames"
    )
    assert findings[0][:2] == ("incomparable", "a")


def test_compare_work_count_mismatch_is_incomparable():
    findings = bc.compare(
        {"a": _row(count=10)},
        {"a": _row(count=12)},
        metric="elapsed_s",
        pct=10,
        abs_s=None,
        count_key="frames",
    )
    assert findings[0][:2] == ("incomparable", "a")
    assert "frames 10 vs baseline 12" in findings[0][2]


def test_compare_regression_on_percent():
    findings = bc.compare(
        {"a": _row(elapsed=1.2)},
        {"a": _row(elapsed=1.0)},
        metric="elapsed_s",
        pct=10,
        abs_s=None,
        count_key="frames",
    )
    assert findings[0][:2] == ("regression", "a")
    assert "+20.0%" in findings[0][2]
    assert (
        bc.compare(
            {"a": _row(elapsed=1.2)},
            {"a": _row(elapsed=1.0)},
            metric="elapsed_s",
            pct=25,
            abs_s=None,
            count_key="frames",
        )
        == []
    )


def test_compare_needs_both_thresholds_when_both_given():
    # +20% but only +0.2 s: the absolute bar is not cleared.
    assert (
        bc.compare(
            {"a": _row(elapsed=1.2)},
            {"a": _row(elapsed=1.0)},
            metric="elapsed_s",
            pct=10,
            abs_s=0.5,
            count_key="frames",
        )
        == []
    )
    # +20% and +2 s: both cleared.
    findings = bc.compare(
        {"a": _row(elapsed=12.0)},
        {"a": _row(elapsed=10.0)},
        metric="elapsed_s",
        pct=10,
        abs_s=0.5,
        count_key="frames",
    )
    assert findings[0][0] == "regression"


def test_compare_absolute_only():
    findings = bc.compare(
        {"a": _row(elapsed=1.6)},
        {"a": _row(elapsed=1.0)},
        metric="elapsed_s",
        pct=None,
        abs_s=0.5,
        count_key="frames",
    )
    assert findings[0][0] == "regression"


def test_status_and_exit_code_priority():
    findings = [
        ("regression", "a", ""),
        ("invalid", "b", ""),
        ("incomparable", "c", ""),
    ]
    assert bc.status_of("a", findings) == "regression"
    assert bc.status_of("b", findings) == "invalid"
    assert bc.status_of("d", findings) == "ok"
    assert bc.exit_code(findings) == bc.EXIT_INVALID
    assert bc.exit_code([("incomparable", "c", ""), ("regression", "a", "")]) == 4
    assert bc.exit_code([("regression", "a", "")]) == 1
    assert bc.exit_code([]) == 0


# ---------- environment and workload identity ----------------------------------


def test_env_diff_ignores_commit_but_not_settings():
    a = {"commit": "aaa", "platform": "mac", "python": "3.12.1", "settings": {"X": 1}}
    b = {"commit": "bbb", "platform": "mac", "python": "3.12.9", "settings": {"X": 2}}
    assert bc.env_diff(a, b) == ["X: 1 -> 2"]
    b["platform"] = "linux"
    assert bc.env_diff(a, b)[0].startswith("platform:")
    b["python"] = "3.13.0"
    assert any(line.startswith("python:") for line in bc.env_diff(a, b))


def test_meta_diff_covers_fixture_spec_and_workload():
    base = {"fixture": {"spec": {"duration": 120}}, "workload": {"interval": 0.1}}
    same = {
        "fixture": {"spec": {"duration": 120}, "probed": {}},
        "workload": {"interval": 0.1},
    }
    assert bc.meta_diff(base, same) == []
    other = {"fixture": {"spec": {"duration": 15}}, "workload": {"interval": 0.5}}
    assert bc.meta_diff(base, other) == [
        "fixture.duration: 120 -> 15",
        "workload.interval: 0.1 -> 0.5",
    ]


def test_load_baseline_refuses_mismatch_and_env(tmp_path, capsys):
    snap = tmp_path / "base.json"
    meta = {
        "fixture": {"spec": {"duration": 15}},
        "workload": {},
        "env": {"chip": "M1"},
    }
    snap.write_text(json.dumps({"meta": meta, "tools": {"a": {}}}))
    rows, problems = bc.load_baseline(snap, "tools", meta, allow_env=False)
    assert rows == {"a": {}} and problems == []
    other = dict(meta, env={"chip": "M3"})
    rows, problems = bc.load_baseline(snap, "tools", other, allow_env=False)
    assert rows is None and problems == ["chip: M1 -> M3"]
    rows, problems = bc.load_baseline(snap, "tools", other, allow_env=True)
    assert rows == {"a": {}}
    assert "env: chip: M1 -> M3" in capsys.readouterr().out
    wrong = dict(meta, fixture={"spec": {"duration": 120}})
    rows, problems = bc.load_baseline(snap, "tools", wrong, allow_env=True)
    assert rows is None and problems == ["fixture.duration: 15 -> 120"]


def test_fixture_matches():
    spec = {"duration": 15, "width": 1280, "fps": 30.0}
    assert bc.fixture_matches({"duration": 15.3, "width": 1280, "fps": 30.0}, spec)
    assert not bc.fixture_matches({"duration": 16, "width": 1280, "fps": 30.0}, spec)
    assert not bc.fixture_matches({"duration": 15, "width": 640, "fps": 30.0}, spec)
    assert not bc.fixture_matches(None, spec)


def test_probe_fixture_parses_ffprobe_json(monkeypatch, tmp_path):
    payload = json.dumps(
        {
            "streams": [
                {
                    "codec_type": "video",
                    "width": 1280,
                    "height": 720,
                    "r_frame_rate": "30/1",
                },
                {"codec_type": "audio"},
            ],
            "format": {"duration": "15.02"},
        }
    )
    monkeypatch.setattr(
        bc.subprocess,
        "run",
        lambda *a, **k: SimpleNamespace(stdout=payload, returncode=0),
    )
    assert bc.probe_fixture(tmp_path / "v.mp4") == {
        "duration": 15.02,
        "width": 1280,
        "height": 720,
        "fps": 30.0,
        "has_audio": True,
    }


def test_probe_fixture_survives_a_missing_ffprobe(monkeypatch, tmp_path):
    def missing(*a, **k):
        raise FileNotFoundError("ffprobe")

    monkeypatch.setattr(bc.subprocess, "run", missing)
    assert bc.probe_fixture(tmp_path / "v.mp4") is None


# ---------- the fresh-process run ----------------------------------------------


def _fake_subprocess(monkeypatch, *, returncode=0, doc=None):
    def run(cmd, **kwargs):
        path = Path(cmd[cmd.index("--profile-output") + 1])
        if doc is not None:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(doc))
        return SimpleNamespace(stdout="out\n", stderr="", returncode=returncode)

    monkeypatch.setattr(bc.subprocess, "run", run)


def test_run_clipgen_valid(monkeypatch, tmp_path):
    _fake_subprocess(
        monkeypatch, doc={"mode": "profile", "labels": {"pipeline.clip": {}}}
    )
    result = bc.run_clipgen([], tmp_path, required_labels=["pipeline.clip"])
    assert result["valid"] is True
    assert result["elapsed_s"] >= 0
    assert result["doc"]["mode"] == "profile"


def test_run_clipgen_failed_exit_with_partial_labels_is_invalid(monkeypatch, tmp_path):
    _fake_subprocess(
        monkeypatch,
        returncode=1,
        doc={"mode": "profile", "labels": {"ffprobe.run": {}}},
    )
    result = bc.run_clipgen([], tmp_path, required_labels=["pipeline.clip"])
    assert result["valid"] is False
    assert result["reasons"] == ["exit 1", "missing pipeline.clip"]


def test_run_clipgen_missing_json_is_invalid(monkeypatch, tmp_path):
    _fake_subprocess(monkeypatch)
    result = bc.run_clipgen([], tmp_path, required_labels=[])
    assert result["reasons"] == ["no profile.json"]


def test_run_clipgen_rejects_deep_runs(monkeypatch, tmp_path):
    _fake_subprocess(monkeypatch, doc={"mode": "deep", "labels": {"x": {}}})
    result = bc.run_clipgen([], tmp_path, required_labels=["x"])
    assert result["reasons"] == ["deep run"]


def test_validate_common_args_rejects_deep_with_save():
    ap = argparse.ArgumentParser()
    bc.add_common_args(ap, "callback")
    ap.add_argument("--deep", action="store_true")
    with pytest.raises(SystemExit):
        bc.validate_common_args(ap, ap.parse_args(["--deep", "--save", "x.json"]))
    with pytest.raises(SystemExit):
        bc.validate_common_args(ap, ap.parse_args(["--fail-abs", "1"]))
    bc.validate_common_args(ap, ap.parse_args(["--deep"]))
    assert ap.parse_args([]).runs is None  # the tier fills it


# ---------- long media: scaling, rates, tiers, identity -----------------------------


def test_fit_scaling_reads_linear_flat_and_superlinear():
    linear = bc.fit_scaling([(600, 10.0), (1800, 30.0), (3600, 60.0)])
    assert linear["verdict"] == "linear"
    assert linear["slope_per_hour"] == pytest.approx(60.0)
    assert linear["exponent"] == pytest.approx(1.0)
    assert bc.fit_scaling([(600, 2.0), (3600, 2.05)])["verdict"] == "flat"
    cubic = bc.fit_scaling([(1000, 0.37), (2000, 2.99)])
    assert cubic["verdict"] == "superlinear"
    assert cubic["exponent"] == pytest.approx(3.0, abs=0.05)
    assert bc.fit_scaling([(600, 1.0)])["verdict"] == "n/a"


def test_fit_scaling_exponent_ignores_startup_overhead():
    """A fixed 2 s launch must not read a linear cost as sublinear."""
    pts = [(x, 2.0 + x / 60) for x in (60, 1800, 3600)]
    assert bc.fit_scaling(pts)["verdict"] == "linear"


def test_add_rates_normalizes_by_media_length():
    sample = bc.add_rates({"elapsed_s": 40.0}, 3600.0)
    assert sample["xrt"] == pytest.approx(90.0)
    assert sample["s_per_media_hour"] == pytest.approx(40.0)
    assert bc.add_rates({"elapsed_s": 0.0}, 10.0)["xrt"] == 0.0


def test_build_row_adds_rates_only_with_media():
    results = [{"valid": True, "reasons": [], "elapsed_s": 2.0, "doc": {}}]
    plain = bc.build_row(results, lambda _d: {}, ("elapsed_s",))
    assert "xrt" not in plain["samples"][0]
    rated = bc.build_row(results, lambda _d: {}, ("elapsed_s",), media_s=120.0)
    assert rated["samples"][0]["xrt"] == pytest.approx(60.0)


def test_timeout_scales_with_media():
    assert bc.timeout_for(60, None) == 1800.0
    assert bc.timeout_for(10_800, None) == 32_400.0
    assert bc.timeout_for(10_800, 99) == 99.0


def test_parse_scale_needs_two_positive_durations():
    assert bc.parse_scale("3600, 600,600") == [600, 3600]
    for bad in ("600", "0,600", ""):
        with pytest.raises(ValueError):
            bc.parse_scale(bad)


def _tier_parser():
    ap = argparse.ArgumentParser()
    bc.add_common_args(ap, "callback")
    tiers = {"smoke": {"duration": 120}, "session": {"duration": 3600, "runs": 1}}
    bc.add_long_args(ap, tiers)
    ap.add_argument("--duration", type=int, default=None)
    return ap, tiers


def test_apply_tier_fills_only_unset_flags():
    ap, tiers = _tier_parser()
    args = ap.parse_args(["--tier", "session"])
    bc.apply_tier(ap, args, tiers)
    assert (args.duration, args.runs) == (3600, 1)
    args = ap.parse_args(["--tier", "session", "--duration", "60", "--runs", "2"])
    bc.apply_tier(ap, args, tiers)
    assert (args.duration, args.runs) == (60, 2)
    args = ap.parse_args([])
    bc.apply_tier(ap, args, tiers)
    assert (args.duration, args.runs) == (120, 3)


def test_apply_tier_rejects_scale_with_save_or_source(tmp_path):
    ap, tiers = _tier_parser()
    with pytest.raises(SystemExit):
        bc.apply_tier(ap, ap.parse_args(["--scale", "60,120", "--save", "x"]), tiers)
    with pytest.raises(SystemExit):
        bc.apply_tier(ap, ap.parse_args(["--source", str(tmp_path / "no")]), tiers)


def test_keyframe_gap_is_matched_within_tolerance():
    spec = {"duration": 3600, "keyframe_gap": 8.333}
    assert bc.fixture_matches({"duration": 3600.08, "keyframe_gap": 8.36}, spec)
    # A 1 s-GOP leftover must never pass for an OBS-like fixture.
    assert not bc.fixture_matches({"duration": 3600, "keyframe_gap": 1.0}, spec)


def test_probe_keyframe_gap_takes_the_widest(monkeypatch):
    out = "0.000,K__\n0.033,___\n8.333,K__\n16.666,K__\n19.0,K__\n"
    monkeypatch.setattr(
        bc.subprocess, "run", lambda *_a, **_k: SimpleNamespace(stdout=out)
    )
    assert bc.probe_keyframe_gap(Path("x.mp4")) == pytest.approx(8.333)


def test_ensure_video_rebuilds_on_an_unprobed_mismatch(tmp_path, monkeypatch):
    """Same geometry, different content: the sidecar tells them apart."""
    video = tmp_path / "v_P01.mp4"
    spec = {"duration": 60, "content": "screen"}
    built = []

    def build():
        built.append(1)
        video.write_bytes(b"x")

    monkeypatch.setattr(bc, "probe_fixture", lambda *_a, **_k: {"duration": 60.0})
    bc.ensure_video(video, spec, build)
    bc.ensure_video(video, spec, build)
    assert len(built) == 1
    bc.ensure_video(video, dict(spec, content="testsrc"), build)
    assert len(built) == 2


def test_run_clipgen_reports_a_timeout(tmp_path, monkeypatch):
    def hang(cmd, **kwargs):
        raise bc.subprocess.TimeoutExpired(cmd, kwargs["timeout"], output=b"partial")

    monkeypatch.setattr(bc.subprocess, "run", hang)
    result = bc.run_clipgen(["--x"], tmp_path, required_labels=[], timeout=5)
    assert result["valid"] is False
    assert result["reasons"][0] == "timeout after 5s"
    assert "partial" in result["output"]


def test_run_clipgen_overrides_launch_through_python(tmp_path, monkeypatch):
    seen = {}

    def fake(cmd, **_kwargs):
        seen["cmd"] = cmd
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    monkeypatch.setattr(bc.subprocess, "run", fake)
    bc.run_clipgen(["-r", "6"], tmp_path, required_labels=[], overrides={"A": 1})
    cmd = seen["cmd"]
    assert cmd[:4] == ["uv", "run", "python", "-c"]
    assert json.loads(cmd[5]) == {"A": 1}
    assert cmd[6:8] == ["-r", "6"]
    bc.run_clipgen(["-r", "6"], tmp_path, required_labels=[])
    assert seen["cmd"][:3] == ["uv", "run", "clipgen.py"]


def test_link_source_never_copies(tmp_path):
    source = tmp_path / "Session 7.mkv"
    source.write_bytes(b"media")
    link = bc.link_source(source, tmp_path / "in", "bench_P01.mp4")
    assert link.is_symlink() and link.resolve() == source.resolve()
    bc.link_source(source, tmp_path / "in", "bench_P01.mp4")  # relink is fine
    ident = bc.source_identity(source)
    assert ident["name"] == "Session 7.mkv" and ident["bytes"] == 5
    assert str(tmp_path) not in json.dumps(ident)
