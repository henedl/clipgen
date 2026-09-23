"""hardware.profile() is stdlib-only and must never raise."""

from __future__ import annotations

import os

import hardware


def test_profile_reports_ints_and_a_gpu_class():
    hardware.profile.cache_clear()
    hw = hardware.profile()
    assert isinstance(hw["memory_mb"], int) and hw["memory_mb"] >= 0
    assert isinstance(hw["cpu_count"], int)
    assert hw["gpu"] in {"apple", "nvidia", "none"}
    assert hw["unified_memory"] == (hw["gpu"] == "apple")
    assert isinstance(hw["chip"], str)


def test_memory_probe_failure_reads_as_unknown(monkeypatch):
    def boom(_name):
        raise OSError("no sysconf")

    monkeypatch.setattr(hardware.sys, "platform", "linux")
    monkeypatch.setattr(os, "sysconf", boom)
    assert hardware.memory_bytes() == 0


def test_gpu_class_table(monkeypatch):
    monkeypatch.setattr(hardware.sys, "platform", "darwin")
    monkeypatch.setattr(hardware.platform, "machine", lambda: "arm64")
    assert hardware.gpu_class() == "apple"

    monkeypatch.setattr(hardware.platform, "machine", lambda: "x86_64")
    monkeypatch.setattr(hardware.shutil, "which", lambda _b: "/usr/bin/nvidia-smi")
    assert hardware.gpu_class() == "nvidia"

    monkeypatch.setattr(hardware.shutil, "which", lambda _b: None)
    assert hardware.gpu_class() == "none"
