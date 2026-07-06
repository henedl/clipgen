import os
import subprocess
import sys
import textwrap
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
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
    env = os.environ.copy()
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    env["PYTHONPATH"] = str(ROOT)
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

        for name in ({modules},):
            importlib.import_module(name)

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
        source = (ROOT / f"{module}.py").read_text(encoding="utf-8")
        assert "from icecream import ic" not in source
