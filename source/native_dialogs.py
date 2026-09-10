"""Native OS dialogs and reveals: shell out to platform tooling, no GUI dependency."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

from utils import warning_print


def open_native_folder_picker(initial_dir: str = "") -> str | None:
    """Open a native folder picker and return the chosen path.

    Returns the chosen folder's absolute path on confirm, ``None`` when the
    user cancels or no native dialog is available on this platform.
    """
    import subprocess

    if sys.platform == "darwin":
        prompt = "Select a folder for clipgen"
        safe_initial = ""
        if initial_dir:
            try:
                if Path(initial_dir).is_dir():
                    safe_initial = initial_dir
            except OSError:
                pass
        if safe_initial:
            # Escape backslashes first, then double quotes, for safe embedding
            # in an AppleScript double-quoted string literal.
            escaped = safe_initial.replace("\\", "\\\\").replace('"', '\\"')
            script = (
                f'set chosenFolder to choose folder with prompt "{prompt}" '
                f'default location POSIX file "{escaped}"\n'
                "return POSIX path of chosenFolder"
            )
        else:
            script = (
                f'set chosenFolder to choose folder with prompt "{prompt}"\n'
                "return POSIX path of chosenFolder"
            )
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                timeout=300,
                check=False,
            )
        except (OSError, subprocess.TimeoutExpired):
            return None
        if result.returncode != 0:
            # User cancelled or AppleScript failed — both are non-errors here.
            return None
        path = result.stdout.strip().rstrip("/")
        return path or None

    # Tkinter fallback; uv-managed Pythons often lack Tk, so None is expected.
    try:
        import tkinter
        from tkinter import filedialog
    except ImportError:
        return None
    try:
        root = tkinter.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askdirectory(
            initialdir=initial_dir or str(Path.home()),
            title="Select a folder for clipgen",
        )
        root.destroy()
    except Exception:
        return None
    return path or None


def reveal_in_file_manager(path: Path) -> bool:
    """Show *path* in the OS file browser. Returns False if nothing could run.

    A file is revealed with its containing folder open and the file selected; a
    directory is simply opened. Linux has no portable "select this file", so it
    falls back to opening the parent.
    """
    is_dir = path.is_dir()
    if sys.platform == "darwin":
        command = ["open", str(path)] if is_dir else ["open", "-R", str(path)]
    elif os.name == "nt":
        command = ["explorer", str(path)] if is_dir else ["explorer", f"/select,{path}"]
    else:
        command = ["xdg-open", str(path if is_dir else path.parent)]
    try:
        subprocess.run(command, check=False)
    except OSError as exc:
        warning_print(f"Could not open {path}: {exc}")
        return False
    return True
