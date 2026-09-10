"""Machine facts for model recommendations: memory, cores, GPU class.

Stdlib only. psutil is a C-extension wheel PyInstaller would have to bundle,
so memory comes from ``os.sysconf`` (macOS, Linux) or ``GlobalMemoryStatusEx``
(Windows). Every probe degrades to 0 / "" / "none"; ``profile()`` never raises.
"""

import ctypes
import functools
import os
import platform
import shutil
import subprocess
import sys
from typing import Any


def _memory_bytes_windows() -> int:
    class _MemoryStatus(ctypes.Structure):
        _fields_ = [
            ("dwLength", ctypes.c_ulong),
            ("dwMemoryLoad", ctypes.c_ulong),
            ("ullTotalPhys", ctypes.c_ulonglong),
            ("ullAvailPhys", ctypes.c_ulonglong),
            ("ullTotalPageFile", ctypes.c_ulonglong),
            ("ullAvailPageFile", ctypes.c_ulonglong),
            ("ullTotalVirtual", ctypes.c_ulonglong),
            ("ullAvailVirtual", ctypes.c_ulonglong),
            ("sullAvailExtendedVirtual", ctypes.c_ulonglong),
        ]

    if sys.platform != "win32":
        return 0
    status = _MemoryStatus()
    status.dwLength = ctypes.sizeof(_MemoryStatus)
    if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
        return 0
    return int(status.ullTotalPhys)


def memory_bytes() -> int:
    """Physical RAM in bytes; 0 when the platform gives no answer."""
    try:
        if sys.platform == "win32":
            return _memory_bytes_windows()
        return int(os.sysconf("SC_PHYS_PAGES") * os.sysconf("SC_PAGE_SIZE"))
    except (AttributeError, OSError, ValueError):
        return 0


def gpu_class() -> str:
    """``apple`` (unified memory, Metal), ``nvidia`` (CUDA), or ``none``."""
    if sys.platform == "darwin" and platform.machine() == "arm64":
        return "apple"
    if shutil.which("nvidia-smi"):
        return "nvidia"
    return "none"


def chip_name() -> str:
    """Marketing name like ``Apple M4 Pro``; empty when unknown."""
    if sys.platform == "darwin":
        try:
            out = subprocess.run(
                ["sysctl", "-n", "machdep.cpu.brand_string"],
                capture_output=True,
                text=True,
                timeout=1,
                check=False,
            ).stdout
            return out.strip()
        except (OSError, subprocess.SubprocessError):
            return ""
    return platform.processor() or ""


@functools.lru_cache(maxsize=1)
def profile() -> dict[str, Any]:
    """Cached machine summary; probed once per process."""
    gpu = gpu_class()
    return {
        "memory_mb": memory_bytes() // (1024 * 1024),
        "cpu_count": os.cpu_count() or 0,
        "arch": platform.machine() or "",
        "gpu": gpu,
        "chip": chip_name(),
        "unified_memory": gpu == "apple",
    }
