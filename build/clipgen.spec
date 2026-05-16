# clipgen.spec
# Build with: pyinstaller --clean --noconfirm build/clipgen.spec
#
# App icon: build/clipgen.icns (macOS) and build/clipgen.ico (Windows) are
# regenerated from the brand mark with `uv run build/render_icons.py`.
# PyInstaller picks the right format per host platform from the list below.

import sys

from PyInstaller.utils.hooks import collect_data_files, collect_submodules


hiddenimports = []
hiddenimports += collect_submodules("gspread")
hiddenimports += collect_submodules("openpyxl")
hiddenimports += collect_submodules("rich")

datas = []
datas += collect_data_files("gspread")
datas += collect_data_files("openpyxl")
datas += collect_data_files("rich")
datas += [("../assets", "assets")]
datas += [("VERSION", ".")]
datas += [("../CHANGELOG.md", ".")]


excludes = [
    # Test frameworks — not needed in production binary
    "pytest", "_pytest", "py", "pluggy", "iniconfig",
    # Torch submodules clipgen never uses
    "tensorboard", "torch.distributed",
    # GUI toolkit pulled in transitively — clipgen is CLI-only
    "tkinter", "_tkinter",
    # Other unused transitive dependencies
    "matplotlib", "IPython", "notebook", "jupyter",
]

a = Analysis(
    ["../clipgen.py"],
    pathex=[".."],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="clipgen",
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=("clipgen.icns" if sys.platform == "darwin" else "clipgen.ico"),
)

if sys.platform == "darwin":
    import stat
    from pathlib import Path

    _version = (Path(SPECPATH) / "VERSION").read_text().strip()
    app = BUNDLE(
        exe,
        name="clipgen.app",
        icon="clipgen.icns",
        bundle_identifier="se.signalresearch.clipgen",
        version=_version,
        info_plist={
            "CFBundleShortVersionString": _version,
            "CFBundleVersion": _version,
            "NSHighResolutionCapable": True,
            "LSMinimumSystemVersion": "11.0",
        },
    )

    # PyInstaller writes the CLI binary to Contents/MacOS/clipgen. clipgen is a
    # console app, so double-clicking the .app would silently launch the binary
    # with no visible terminal. Move the real binary aside and put a tiny
    # shell launcher in its place that opens Terminal pointing at it.
    _macos_dir = Path(DISTPATH) / "clipgen.app" / "Contents" / "MacOS"
    _real_bin = _macos_dir / "clipgen-bin"
    _launcher = _macos_dir / "clipgen"
    if _launcher.exists() and not _real_bin.exists():
        _launcher.rename(_real_bin)
        _launcher.write_text(
            '#!/bin/bash\n'
            'DIR="$(cd "$(dirname "$0")" && pwd)"\n'
            'open -a Terminal "$DIR/clipgen-bin"\n'
        )
        _launcher.chmod(_launcher.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)
