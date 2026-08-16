"""Tests for the opt-in profiling layer (source/profiling.py).

Covers the accumulator contract (no-op when ``config.PROFILING`` is off,
accumulation when on), the report line shape agents grep for, and the
``/api/profile`` endpoint's opt-in gating on the combined Flask app.
"""

from __future__ import annotations

import pytest

import config
import profiling


@pytest.fixture(autouse=True)
def _clean_profiling(monkeypatch):
    """Every test starts with profiling off and an empty accumulator."""
    monkeypatch.setattr(config, "PROFILING", False)
    profiling.reset()
    yield
    profiling.reset()


# ---------- accumulator ------------------------------------------------------


def test_recording_is_noop_when_off():
    profiling.add("x", 1.0)
    profiling.count("y")
    with profiling.span("z"):
        pass
    assert profiling.snapshot() == {}


def test_add_and_count_accumulate_when_on(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("work", 0.5, 2)
    profiling.add("work", 0.25)
    profiling.count("hits", 3)
    snap = profiling.snapshot()
    assert snap["work"] == {"seconds": 0.75, "count": 3}
    assert snap["hits"] == {"seconds": 0.0, "count": 3}


def test_span_times_the_block(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    with profiling.span("block"):
        pass
    snap = profiling.snapshot()
    assert snap["block"]["count"] == 1
    assert snap["block"]["seconds"] >= 0.0


def test_span_records_on_exception(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    with pytest.raises(ValueError, match="boom"), profiling.span("failing"):
        raise ValueError("boom")
    assert profiling.snapshot()["failing"]["count"] == 1


def test_timed_decorator(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)

    @profiling.timed("fn")
    def fn(x):
        return x + 1

    assert fn(1) == 2
    assert fn(2) == 3
    assert profiling.snapshot()["fn"]["count"] == 2


def test_snapshot_sorted_by_seconds_desc(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("small", 0.1)
    profiling.add("big", 5.0)
    assert list(profiling.snapshot()) == ["big", "small"]


def test_reset_clears(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("x", 1.0)
    profiling.reset()
    assert profiling.snapshot() == {}


def test_label_cap_drops_new_labels(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(profiling, "_MAX_LABELS", 2)
    profiling.add("a", 1.0)
    profiling.add("b", 1.0)
    profiling.add("c", 1.0)  # dropped: at cap
    profiling.add("a", 1.0)  # existing label still accumulates
    snap = profiling.snapshot()
    assert set(snap) == {"a", "b"}
    assert snap["a"]["seconds"] == 2.0


def test_enable_flips_config_and_registers_report_once(monkeypatch):
    registered = []
    monkeypatch.setattr(profiling.atexit, "register", registered.append)
    monkeypatch.setattr(profiling, "_REPORT_REGISTERED", False)
    profiling.enable()
    profiling.enable()
    assert config.PROFILING is True
    assert registered == [profiling.report]


# ---------- report shape -----------------------------------------------------


def test_report_line_shape(monkeypatch, capsys):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("scan.callback", 12.8, 312)
    profiling.count("media_cache.hit", 42)
    profiling.report()
    lines = capsys.readouterr().out.splitlines()
    assert all(line.startswith("profile | ") for line in lines)
    assert lines[0].startswith("profile | scan.callback")  # seconds-desc order
    assert "n=312" in lines[0]
    assert "avg=" in lines[0]
    assert "n=42" in lines[1]
    assert "avg=" not in lines[1]  # pure counters carry no average


def test_report_silent_when_empty(capsys):
    profiling.report()
    assert capsys.readouterr().out == ""


def test_scan_summary_single_line(monkeypatch, capsys):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.scan_summary("bench_P01.mp4", [("decode_wait", 8.1, 300)])
    out = capsys.readouterr().out
    assert out.startswith("profile | scan bench_P01.mp4: ")
    assert "decode_wait=8.100s/n=300" in out


def test_scan_summary_noop_when_off(capsys):
    profiling.scan_summary("x", [("a", 1.0, 1)])
    assert capsys.readouterr().out == ""


# ---------- /api/profile endpoint ---------------------------------------------


@pytest.fixture(scope="module")
def combined_app():
    pytest.importorskip("flask")
    import server

    return server.build_combined_app(worksheet=None, default_page="studio")


@pytest.fixture
def client(combined_app):
    return combined_app.test_client()


def test_api_profile_404_when_off(client):
    resp = client.get("/api/profile")
    assert resp.status_code == 404
    assert resp.get_json()["ok"] is False


def test_api_profile_snapshot_and_reset(client, monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("scan.callback", 1.5, 10)

    resp = client.get("/api/profile")
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["ok"] is True
    assert body["profile"]["scan.callback"] == {"seconds": 1.5, "count": 10}
    # Request itself was timed by the route hook.
    assert any(label.startswith("route ") for label in profiling.snapshot())

    resp = client.get("/api/profile?reset=1")
    assert resp.status_code == 200
    resp = client.get("/api/profile")
    body = resp.get_json()
    assert "scan.callback" not in body["profile"]
