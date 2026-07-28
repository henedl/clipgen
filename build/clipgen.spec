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
# desktop.py imports webview lazily, so PyInstaller's static analysis never sees
# the platform backend modules (WKWebView via pyobjc, WebView2 via clr).
hiddenimports += collect_submodules("webview")

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
    # Other unused transitive dependencies
    "matplotlib", "IPython", "notebook", "jupyter",
]
# tkinter is intentionally NOT excluded: utils.open_native_folder_picker
# uses tkinter.filedialog as the non-macOS fallback for the Start overlay's
# Browse button. Excluding it silently breaks the Browse flow on Windows/Linux
# frozen builds.

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
    # Windows keeps a console so `clipgen.exe --gif ...` still prints — a
    # GUI-subsystem build has no stdout at all, and PyInstaller has no dual-mode
    # option. desktop.py hides that console at runtime when (and only when) this
    # process owns it, so a double-click shows just the window.
    # macOS has no such subsystem split: console=False only stops a Terminal
    # being spawned, and `clipgen.app/Contents/MacOS/clipgen --help` still prints.
    console=(sys.platform == "win32"),
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=("clipgen.icns" if sys.platform == "darwin" else "clipgen.ico"),
)

if sys.platform == "darwin":
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
            # macOS raises a "find devices on your local network?" prompt on
            # first launch because the app binds a TCP socket (the UI is served
            # over loopback to the embedded webview). Without this key the
            # prompt has no explanation, which reads as spyware. Denying it does
            # not break anything: nothing outside 127.0.0.1 is ever contacted.
            "NSLocalNetworkUsageDescription": (
                "clipgen serves its own interface to the app window on this "
                "Mac. Nothing is sent off your machine."
            ),
        },
    )
    # No launcher shim here any more. Contents/MacOS/clipgen is the real binary:
    # double-clicking opens the pywebview window (cli.main sends frozen + no-argv
    # straight to desktop mode), and running that same path from a terminal still
    # behaves as the CLI. The old shim renamed the binary to clipgen-bin and
    # dropped in a bash script that ran `open -a Terminal`, which is exactly the
    # "it's really a CLI tool" seam this change removes.
