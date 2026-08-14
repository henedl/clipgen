"""Guard the source layout: `source/` holds every module, and `pyproject.toml`
lists them all.

The project is a flat set of top-level modules in ``source/``, imported by bare
name (``import config``). ``source/`` is not a package — no ``__init__.py`` — it
is just put on ``sys.path`` by the ``clipgen.py`` launcher for real runs and by
``tests/conftest.py`` under pytest.

Two drifts, both silent from the source tree:

* A module missing from ``[tool.setuptools] py-modules`` imports fine locally, but
  ``uv pip install .`` ships only the listed ones, so installed and frozen
  environments die with ``ModuleNotFoundError``.
* A new module in the *repo root* also imports fine locally (the root is still
  ``sys.path[1]``) and PyInstaller still bundles it, so nothing fails until an
  install — by which time the flat root has quietly grown back.
"""

import os
import subprocess
import sys
import tomllib
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
_SOURCE = _ROOT / "source"


def _source_modules() -> set[str]:
    """Top-level importable modules: ``source/*.py`` minus private/dunder files."""
    return {p.stem for p in _SOURCE.glob("*.py") if not p.name.startswith("_")}


def _listed_modules() -> set[str]:
    data = tomllib.loads((_ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    return set(data["tool"]["setuptools"]["py-modules"])


def test_py_modules_covers_every_source_module() -> None:
    missing = _source_modules() - _listed_modules()
    assert not missing, (
        "source/ modules missing from pyproject [tool.setuptools] py-modules "
        f"(install would break `import` of these): {sorted(missing)}"
    )


def test_py_modules_has_no_phantom_entries() -> None:
    phantom = _listed_modules() - _source_modules()
    assert not phantom, (
        "pyproject py-modules lists modules with no matching source/*.py file: "
        f"{sorted(phantom)}"
    )


def test_package_dir_points_at_source() -> None:
    """``py-modules`` names resolve to ``source/<name>.py``, not the repo root."""
    data = tomllib.loads((_ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    assert data["tool"]["setuptools"]["package-dir"] == {"": "source"}


def test_repo_root_holds_only_the_launcher() -> None:
    """No product module may sit beside ``clipgen.py`` again."""
    root_modules = {p.name for p in _ROOT.glob("*.py")}
    assert root_modules == {"clipgen.py"}, (
        "the repo root must hold only the clipgen.py launcher; product modules "
        f"belong in source/. Found: {sorted(root_modules)}"
    )


def test_source_dir_is_not_a_package() -> None:
    """``source/__init__.py`` would break bare-name imports and ty's discovery.

    With an ``__init__.py`` present, ``source/`` becomes a package: ``import
    config`` no longer resolves from it, and ty stops treating it as a module
    root. Both failures are wholesale rather than local, so guard the absence.
    """
    assert not (_SOURCE / "__init__.py").exists()


def _fetch_binaries_module():
    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "fetch_binaries", _ROOT / "build" / "fetch_binaries.py"
    )
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_ffmpeg_pins_cover_both_platforms_and_are_wellformed() -> None:
    """The desktop bundle ships pinned ffmpeg/ffprobe; a malformed or partial
    PINS table only surfaces on the (slow, per-platform) release build."""
    pins = _fetch_binaries_module().PINS
    expected_targets = {
        "macos-arm64": {"ffmpeg", "ffprobe"},
        "windows-x64": {"ffmpeg.exe", "ffprobe.exe"},
    }
    assert set(pins) == set(expected_targets)
    for plat, archives in pins.items():
        targets = set()
        for archive in archives:
            assert archive["url"].startswith("https://"), archive["url"]
            assert _looks_like_sha256(archive["sha256"]), archive["url"]
            for member in archive["members"].values():
                assert _looks_like_sha256(member["sha256"]), member["target"]
                targets.add(member["target"])
        assert targets == expected_targets[plat], plat


def _looks_like_sha256(value: str) -> bool:
    return len(value) == 64 and all(c in "0123456789abcdef" for c in value)


def test_spec_guards_and_upx_excludes_the_vendored_tools() -> None:
    """The spec must refuse to build without the fetched binaries (a silent
    skip ships an app that dies on its startup ffmpeg check), and UPX must
    never touch them (it corrupts signed mach-O / packed ffmpeg.exe)."""
    spec_text = (_ROOT / "build" / "clipgen.spec").read_text(encoding="utf-8")
    assert "fetch_binaries.py" in spec_text, "vendor guard missing from spec"
    assert spec_text.count('upx_exclude=["ffmpeg*", "ffprobe*"]') == 2, (
        "both EXE and COLLECT must exclude the bundled video tools from UPX"
    )
    assert "binaries=binaries" in spec_text, (
        "Analysis must collect the vendored binaries list"
    )


def test_spec_never_strips_or_packs_on_windows() -> None:
    """``strip``/``upx`` must stay macOS-only, or Windows ships a dead exe.

    PyInstaller shells out to whatever ``strip`` is on PATH. The GitHub windows
    runner carries a GNU binutils ``strip``, which happily rewrites MSVC-built
    PE DLLs into images the Windows loader rejects — ``python312.dll`` included.
    The build stays green (``strip`` returns 0), every file is present and
    plausibly sized, and the shipped ``clipgen.exe`` dies before any Python runs
    with ``Failed to load Python DLL ... LoadLibrary: Invalid access to memory
    location``. That shipped in every Windows build until the ``_strip`` gate.
    """
    spec_text = (_ROOT / "build" / "clipgen.spec").read_text(encoding="utf-8")
    assert '_strip = sys.platform == "darwin"' in spec_text, (
        "the spec must gate `strip` to macOS; GNU strip corrupts Windows PE DLLs"
    )
    assert spec_text.count("strip=_strip") == 2, (
        "both EXE and COLLECT must take `strip` from the platform gate"
    )
    assert "strip=True" not in spec_text, (
        "`strip=True` is unconditional and breaks the Windows bundle"
    )
    assert spec_text.count("upx=False") == 2, (
        "UPX corrupts the same PE DLLs; keep it off in both EXE and COLLECT"
    )


def test_build_workflow_smoke_launches_both_bundles() -> None:
    """A green PyInstaller build proves nothing about whether the exe launches.

    The bootloader loads ``python312.dll`` before argv is parsed, so ``--help``
    is enough to catch a bundle whose DLLs were corrupted at build time — the
    one failure mode the feature-verification steps all miss.
    """
    workflow = (_ROOT / ".github" / "workflows" / "build-binaries.yml").read_text(
        encoding="utf-8"
    )
    assert "dist/clipgen/clipgen.exe --help" in workflow, (
        "the Windows leg must smoke-launch the frozen exe"
    )
    assert "./dist/clipgen.app/Contents/MacOS/clipgen --help" in workflow, (
        "the macOS leg must smoke-launch the frozen binary"
    )


def test_gitignore_allowlists_the_fetch_script() -> None:
    """``build/*`` is gitignored with a ``!`` allowlist; without its entry the
    fetch script exists locally but never lands in a commit."""
    gitignore = (_ROOT / ".gitignore").read_text(encoding="utf-8")
    assert "!build/fetch_binaries.py" in gitignore.splitlines()


def test_launcher_bootstraps_from_a_foreign_cwd() -> None:
    """``clipgen.py`` must find ``source/`` without help from the environment.

    Runs with ``PYTHONPATH`` cleared and from a directory that is not the repo,
    which is the shape of every real invocation. ``--help`` exits inside
    ``parse_arguments()``, before ``main`` chdirs or touches any media.
    """
    env = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    result = subprocess.run(
        [sys.executable, str(_ROOT / "clipgen.py"), "--help"],
        cwd=Path(sys.prefix),
        env=env,
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stderr or result.stdout
    assert "clipgen" in result.stdout
