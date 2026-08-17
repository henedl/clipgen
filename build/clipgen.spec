# clipgen.spec
# Build with: pyinstaller --clean --noconfirm build/clipgen.spec
#
# App icon: build/clipgen.icns (macOS) and build/clipgen.ico (Windows) are
# regenerated from the brand mark with `uv run build/render_icons.py`.
# PyInstaller picks the right format per host platform from the list below.

import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules


# Relative paths in a spec are resolved against two *different* bases, which is
# a trap worth spelling out: `datas`/`binaries` resolve against SPECPATH (this
# file's directory), but `pathex` resolves against the process CWD. Since the
# documented build command runs from the repo root, a relative `pathex` entry
# silently points outside the repo — and PyInstaller reports no error, it just
# ships an app with no application code in it. Derive it from SPECPATH instead.
_REPO_ROOT = Path(SPECPATH).parent  # noqa: F821 — SPECPATH is injected by PyInstaller

hiddenimports = []
hiddenimports += collect_submodules("gspread")
hiddenimports += collect_submodules("openpyxl")
hiddenimports += collect_submodules("rich")
# desktop.py imports webview lazily, so PyInstaller's static analysis never sees
# the platform backend modules (WKWebView via pyobjc, WebView2 via clr).
hiddenimports += collect_submodules("webview")
# rapidocr's public API is a lazy __getattr__ table (importlib.import_module
# with computed names), invisible to static analysis. collect_submodules skips
# the optional engines whose deps we don't ship (torch/paddle/openvino) with a
# warning; the onnxruntime engine imports cleanly and is what clipgen uses.
hiddenimports += collect_submodules("rapidocr")
if sys.platform == "darwin":
    # desktop_chrome.py reaches these through importlib.import_module (a literal
    # `import AppKit` is an unresolved-import error on the Linux typecheck CI),
    # and a string import is invisible to PyInstaller's static analysis. The
    # cocoa backend happens to pull both in today; do not rely on that.
    hiddenimports += ["AppKit", "PyObjCTools.AppHelper"]

datas = []
datas += collect_data_files("gspread")
datas += collect_data_files("openpyxl")
datas += collect_data_files("rich")
# faster-whisper loads its bundled Silero VAD model (assets/silero_vad_v6.onnx
# as of 1.2.1) via __file__ arithmetic, and PyInstaller ships no hook for it —
# without this the frozen app computes the right path to a file that was never
# copied in, and every default (VAD-on) transcription dies with NO_SUCHFILE.
_fw_datas = collect_data_files("faster_whisper")
if not any(src.endswith(".onnx") for src, _dest in _fw_datas):
    raise SystemExit(
        "clipgen.spec: collect_data_files('faster_whisper') found no .onnx VAD "
        "model. faster-whisper moved or renamed its assets; frozen transcription "
        "would break at runtime. Inspect the installed package and update this spec."
    )
datas += _fw_datas
# rapidocr resolves its yaml configs and bundled default (ch/en) det/cls/rec
# models via package-relative paths; no PyInstaller contrib hook exists for it.
# Guarded like faster-whisper above: a wheel that stops shipping models would
# otherwise freeze an app whose OCR downloads at runtime or dies.
_rapidocr_datas = collect_data_files("rapidocr")
if not any(src.endswith(".onnx") for src, _dest in _rapidocr_datas):
    raise SystemExit(
        "clipgen.spec: collect_data_files('rapidocr') found no bundled .onnx "
        "models. rapidocr moved or stopped shipping its default models; frozen "
        "OCR would break at runtime. Inspect the installed package."
    )
datas += _rapidocr_datas
# Vendored non-default recognition models (fetch_binaries.py OCR_MODEL_PINS),
# guarded like ffmpeg below: a build that skipped the fetch must fail here.
# "ocr_models" matches screenspace_ocr._vendored_rec_model's frozen lookup.
_ocr_vendor = Path(SPECPATH) / "vendor" / "ocr"  # noqa: F821
_ocr_models = ["latin_rec.onnx", "japan_rec.onnx", "korean_rec.onnx"]
_missing_models = [n for n in _ocr_models if not (_ocr_vendor / n).is_file()]
if _missing_models:
    raise SystemExit(
        f"clipgen.spec: vendored OCR models missing from {_ocr_vendor}: "
        f"{sorted(_missing_models)}. Run `uv run build/fetch_binaries.py` first."
    )
datas += [(str(_ocr_vendor / n), "ocr_models") for n in _ocr_models]
datas += [("../assets", "assets")]
datas += [("VERSION", ".")]
datas += [("../CHANGELOG.md", ".")]
# Attribution for the bundled GPL/LGPL/MPL components is a distribution
# obligation, so the notice ships *inside* the bundle where `--licenses` can
# read it — not only beside it in the DMG/zip, where a user can delete it.
datas += [("THIRD-PARTY-LICENSES", ".")]


excludes = [
    # Test frameworks — not needed in production binary
    "pytest", "_pytest", "py", "pluggy", "iniconfig",
    # Other unused transitive dependencies
    "matplotlib", "IPython", "notebook", "jupyter",
    # PyAV. Overridden out of the dependency tree (pyproject.toml), so it is
    # not in the build venv either — this entry documents that on purpose:
    # faster_whisper imports av at module scope but never calls it on the
    # ndarray input path clipgen uses, and transcripts._ensure_av_stub()
    # satisfies the import at runtime. Its wheel would re-add a second ~40 MB
    # FFmpeg beside the vendored ffmpeg/ffprobe binaries below.
    "av",
    # Replaced by in-tree code (screenspace_primitives.PHash and
    # structural_similarity) and removed from the dependency tree; these
    # entries document that on purpose. scipy/pywt rode in only as their
    # transitive dependencies (~110 MB unpacked, scipy alone 71 MB).
    "imagehash", "skimage", "scipy", "pywt",
]

# Bundled ffmpeg/ffprobe: pinned static GPL builds, fetched by
# `uv run build/fetch_binaries.py` into build/vendor/<platform>/bin/ (see the
# PINS block there for provenance; THIRD-PARTY-LICENSES carries the license).
# They land in <bundle>/bin/, which cli.main prepends to PATH on frozen
# launches, so every "ffmpeg" argv in the codebase resolves to these copies.
# Guarded like source/ below: a build that skipped the fetch step must fail
# here, not ship an app that dies on its startup ffmpeg check.
_vendor_platform = "macos-arm64" if sys.platform == "darwin" else "windows-x64"
_vendor_bin = Path(SPECPATH) / "vendor" / _vendor_platform / "bin"  # noqa: F821
_tool_names = ["ffmpeg", "ffprobe"] if sys.platform == "darwin" else ["ffmpeg.exe", "ffprobe.exe"]
_missing_tools = [name for name in _tool_names if not (_vendor_bin / name).is_file()]
if _missing_tools:
    raise SystemExit(
        f"clipgen.spec: bundled video tools missing from {_vendor_bin}: "
        f"{sorted(_missing_tools)}. Run `uv run build/fetch_binaries.py` first."
    )
binaries = [(str(_vendor_bin / name), "bin") for name in _tool_names]
# tkinter is intentionally NOT excluded: utils.open_native_folder_picker
# uses tkinter.filedialog as the non-macOS fallback for the Start overlay's
# Browse button. Excluding it silently breaks the Browse flow on Windows/Linux
# frozen builds.

a = Analysis(
    # The entry script is the repo-root launcher. Everything it imports lives in
    # `source/`, so that directory must be on the analysis path or modulegraph
    # resolves nothing past `from cli import main` — yielding a bundle whose only
    # Python file is the launcher itself. Guarded immediately below.
    ["../clipgen.py"],
    pathex=[str(_REPO_ROOT / "source")],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
)

# Fail loudly on an empty bundle. A wrong `pathex` does not error — PyInstaller
# reports "Build complete!" and ships an app whose only Python file is the
# launcher, which then dies on its first import with nothing on screen. Compare
# what Analysis actually resolved against what is on disk, so the build breaks
# here rather than in a user's Applications folder.
_wanted = {p.stem for p in (_REPO_ROOT / "source").glob("*.py")}
_found = {name for name, _path, _kind in a.pure}
_missing = _wanted - _found
if _missing:
    raise SystemExit(
        f"clipgen.spec: Analysis resolved none/some of source/ "
        f"({len(_found & _wanted)}/{len(_wanted)} modules). Missing: "
        f"{sorted(_missing)}. Check `pathex` — relative entries resolve against "
        f"the CWD, not this spec's directory."
    )

pyz = PYZ(a.pure)

# strip must never run on Windows — PyInstaller's own docs call --strip "not
# recommended for Windows". It shells out to whatever `strip` is on PATH, and
# the GitHub windows runner has a GNU binutils one (it reports the aarch64
# WebView2Loader.dll as "file format not recognized" and rewrites everything
# else). GNU strip mangles MSVC-built PE DLLs: the file still exists and looks
# healthy, but LoadLibrary fails at launch with error 998 "Invalid access to
# memory location", so the exe flashes a console and dies before any Python
# runs. Confirmed from the CI build log: "Executing: strip ...\python312.dll".
# UPX is off for the same reason preemptively: it never ran in CI (no upx on
# the runner today), but if a future image adds one it would corrupt the same
# DLLs the same silent way. macOS keeps strip (Apple's strip understands
# mach-O, and the bundle is re-signed afterwards).
_strip = sys.platform == "darwin"

# One-dir, not one-file. One-file re-extracts the whole archive to a *new* temp
# directory on every launch, so every large dylib (cv2, onnxruntime) loads cold —
# no OS page cache, and macOS re-validates each code signature from scratch.
# Measured double-click to first HTTP response: 17.6 s one-file vs 1.1 s one-dir.
# PyInstaller also deprecated one-file + windowed on macOS ("clashes with macOS's
# security") and makes it an error in v7.0.
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,  # binaries/datas are gathered by COLLECT below
    name="clipgen",
    debug=False,
    bootloader_ignore_signals=False,
    strip=_strip,
    upx=False,
    # Belt-and-braces should UPX ever be re-enabled: it corrupts signed
    # mach-O binaries, and on Windows a packed ffmpeg.exe is a known-broken
    # combination. Mirrored in COLLECT below.
    upx_exclude=["ffmpeg*", "ffprobe*"],
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
    # "lib" instead of PyInstaller's default "_internal": friendlier in
    # Explorer, and safe — runtime code never hardcodes the name (both
    # cli.get_runtime_working_dir and utils.get_bundled_assets_root derive
    # from sys._MEIPASS). An EXE kwarg, not COLLECT: the bootloader bakes the
    # payload dir name into the executable. macOS is unaffected — BUNDLE
    # reorganizes everything into Contents/Frameworks + Contents/Resources.
    contents_directory="lib",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=_strip,
    upx=False,
    upx_exclude=["ffmpeg*", "ffprobe*"],  # see EXE above
    name="clipgen",
)

if sys.platform == "darwin":
    _version = (Path(SPECPATH) / "VERSION").read_text().strip()
    app = BUNDLE(
        coll,
        name="clipgen.app",
        icon="clipgen.icns",
        bundle_identifier="se.signalresearch.clipgen",
        version=_version,
        info_plist={
            "CFBundleShortVersionString": _version,
            "CFBundleVersion": _version,
            "NSHighResolutionCapable": True,
            # onnxruntime and numpy ship macosx_14_0 wheels; the self-built
            # OpenCV wheel is MACOSX_DEPLOYMENT_TARGET=13.0. Advertising 11
            # let Gatekeeper install on 11–13, then import failed at launch.
            "LSMinimumSystemVersion": "14.0",
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
