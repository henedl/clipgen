"""Tests for the resource probes and stack sampler (source/profiling_sampler.py).

Per-thread CPU and memory footprint exist on macOS and Linux only; those
assertions skip elsewhere. Each sampler test runs for about a tenth of a
second: long enough for a dozen ticks, short enough to stay off the slow list.
"""

from __future__ import annotations

import threading
import time

import pytest

import profiling_sampler as sampler

# Asked of the probes themselves: a container may hide what the platform has.
HAS_OS_PROBES = (
    sampler.thread_cpu(threading.current_thread()) is not None
    and "footprint_mb" in sampler.os_usage()
)
needs_os_probes = pytest.mark.skipif(not HAS_OS_PROBES, reason="no OS probes here")


@pytest.fixture(autouse=True)
def _clean_sampler(monkeypatch):
    """Every test starts and ends with the sampler stopped and empty."""
    sampler.stop()
    sampler.snapshot(reset=True)
    monkeypatch.setattr(sampler, "_HZ", 0)
    yield
    sampler.stop()
    sampler.snapshot(reset=True)


def _burn_cpu(seconds: float) -> None:
    end = time.thread_time() + seconds
    while time.thread_time() < end:
        pass


def _burn_while_sampled(seconds: float = 0.1, ticks: int = 10) -> None:
    # A busy thread holds the GIL; on a loaded runner the sampler can starve.
    deadline = time.monotonic() + 5.0
    _burn_cpu(seconds)
    while sampler._TICKS < ticks and time.monotonic() < deadline:
        _burn_cpu(0.005)


def _sample(target, *, hz: int = 500, **start_kwargs):
    thread = threading.Thread(target=target, name="burner-7")
    assert sampler.start(hz, **start_kwargs)
    thread.start()
    thread.join()
    sampler.stop()
    snap = sampler.snapshot()
    assert snap is not None
    return snap


def test_process_cpu_grows_with_work():
    before, children = sampler.process_cpu()
    _burn_cpu(0.02)
    after, _ = sampler.process_cpu()
    assert after - before >= 0.02
    assert children >= 0.0


@needs_os_probes
def test_os_usage_reports_footprint_and_peak():
    usage = sampler.os_usage()
    assert usage["footprint_mb"] > 1.0
    assert usage["peak_footprint_mb"] >= usage["footprint_mb"] * 0.5
    assert usage["disk_read_bytes"] >= 0


@needs_os_probes
def test_thread_cpu_tracks_one_thread():
    me = threading.current_thread()
    before = sampler.thread_cpu(me)
    _burn_cpu(0.02)
    after = sampler.thread_cpu(me)
    assert before is not None and after is not None
    assert after - before >= 0.015


def test_thread_cpu_of_unstarted_thread_is_none():
    assert sampler.thread_cpu(threading.Thread(target=lambda: None)) is None


def test_thread_label_folds_pool_counters():
    assert sampler._thread_label("ThreadPoolExecutor-3_12") == "ThreadPoolExecutor-N_N"
    assert sampler._thread_label("MainThread") == "MainThread"


def test_snapshot_is_none_before_any_start():
    assert sampler.snapshot() is None
    assert sampler.folded() == ""


def test_start_refuses_a_second_sampler():
    assert sampler.start(200)
    assert not sampler.start(200)
    assert sampler.running()
    sampler.stop()
    assert not sampler.running()


def test_sampler_sees_every_thread_in_wall_samples():
    snap = _sample(_burn_while_sampled)
    assert snap["ticks"] >= 5
    assert snap["faults"] == 0
    assert snap["threads"]["burner-N"]["samples"] >= 5
    assert "MainThread" in snap["threads"]
    wall = sampler.folded("wall")
    assert any(
        line.startswith("burner-N;") and "_burn_cpu" in line
        for line in wall.splitlines()
    )


@needs_os_probes
def test_sampler_weights_stacks_by_thread_cpu():
    snap = _sample(_burn_while_sampled)
    assert snap["threads"]["burner-N"]["cpu_s"] >= 0.05
    assert snap["threads"]["MainThread"]["cpu_s"] < 0.05
    top = snap["top_self"][0]
    assert top["frame"].endswith(":_burn_cpu") and top["cpu_s"] >= 0.05
    line = next(row for row in sampler.folded("cpu").splitlines() if "_burn_cpu" in row)
    stack, weight = line.rsplit(" ", 1)
    assert stack.split(";")[0] == "burner-N"
    assert int(weight) >= 50_000


@needs_os_probes
def test_idle_threads_carry_no_cpu_weight():
    snap = _sample(lambda: time.sleep(0.1))
    assert snap["threads"]["burner-N"]["samples"] >= 5
    assert snap["threads"]["burner-N"]["cpu_s"] < 0.02
    # Thread start-up may burn a few microseconds; anything more is a leak.
    weights = [
        int(line.rsplit(" ", 1)[1])
        for line in sampler.folded("cpu").splitlines()
        if line.startswith("burner-N;")
    ]
    assert sum(weights) < 20_000


@needs_os_probes
def test_memory_peak_names_what_was_running(monkeypatch):
    monkeypatch.setattr(sampler, "_MEM_GAP", 0.0)
    readings: list[float] = []
    snap = _sample(
        lambda: time.sleep(0.05),
        on_memory=readings.append,
        context=lambda: ["scan.callback.text"],
    )
    memory = snap["memory"]
    assert readings and memory["mb"] == max(readings)
    assert memory["active"] == ["scan.callback.text"]
    assert memory["timeline"][0][1] > 0


def test_reset_clears_the_window():
    _sample(lambda: _burn_cpu(0.03))
    first = sampler.snapshot(reset=True)
    after = sampler.snapshot()
    assert first is not None and after is not None
    assert first["ticks"] > 0
    assert after["ticks"] == 0 and after["threads"] == {}
    assert sampler.folded("wall") == ""


def test_frame_name_marks_own_and_library_code():
    own = sampler.frame_name(sampler.snapshot.__code__)
    assert own == "profiling_sampler:snapshot"
    lib = sampler.frame_name(threading.Thread.run.__code__)
    assert lib.endswith("/threading:Thread.run")
    assert sampler.frame_name("(stacks past the cap)") == "(stacks past the cap)"
