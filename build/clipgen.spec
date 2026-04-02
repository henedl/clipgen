# clipgen.spec
# Build with: pyinstaller --clean --noconfirm build/clipgen.spec

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
    icon=None,
)
