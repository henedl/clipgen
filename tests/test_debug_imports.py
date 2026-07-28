import os
import subprocess
import sys
import textwrap
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "source"
DEBUG_MODULES = (
    "config",
    "cli",
    "files",
    "google_api",
    "pipeline",
    "spreadsheet",
    "utils",
    "video",
)


def _run_python(source: str) -> subprocess.CompletedProcess[str]:
    """Run *source* in a fresh interpreter that imports from the source tree.

    ``PYTHONPATH`` must point at ``source/``, not the repo root. CI runs
    ``uv pip install ".[dev]"``, so a stale copy of every module also sits in
    site-packages; pointing this at the root would silently fall through to that
    snapshot and pass in CI while failing locally.
    """
    env = os.environ.copy()
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    env["PYTHONPATH"] = str(SOURCE)
    return subprocess.run(
        [sys.executable, "-c", textwrap.dedent(source)],
        cwd=ROOT,
        env=env,
        check=False,
        capture_output=True,
        text=True,
    )


def test_default_debug_modules_do_not_import_icecream() -> None:
    modules = ", ".join(repr(name) for name in DEBUG_MODULES)
    result = _run_python(
        f"""
        import importlib
        import sys
        from pathlib import Path

        # Prove we loaded the source tree, not the site-packages snapshot CI
        # installs. Compare resolved directories rather than matching on the
        # name "source": a substring/segment test is both imprecise (any
        # checkout under a path containing "source" would pass) and, if split
        # on "/", wrong on Windows, where __file__ has no forward slashes.
        expected = Path({str(SOURCE)!r}).resolve()
        for name in ({modules},):
            mod = importlib.import_module(name)
            if Path(mod.__file__).resolve().parent != expected:
                raise SystemExit(f"{{name}} resolved to {{mod.__file__}}")

        if "icecream" in sys.modules:
            raise SystemExit("icecream imported during default startup")
        """
    )

    assert result.returncode == 0, result.stderr or result.stdout


def test_debug_ic_lazy_loads_icecream_when_debugging_enabled() -> None:
    result = _run_python(
        """
        import sys

        import config

        assert "icecream" not in sys.modules
        config.DEBUGGING = True
        config.debug_ic("probe")
        assert "icecream" in sys.modules
        """
    )

    assert result.returncode == 0, result.stderr or result.stdout


def test_debug_modules_do_not_import_icecream_directly() -> None:
    for module in DEBUG_MODULES:
        source = (SOURCE / f"{module}.py").read_text(encoding="utf-8")
        assert "from icecream import ic" not in source
