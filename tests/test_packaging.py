"""Guard the source layout: `source/` holds every module, and `pyproject.toml`
lists them all.

The project is a flat set of top-level modules that live in ``source/`` and are
imported by bare name (``import config``). ``source/`` is not a package — it has
no ``__init__.py`` — it is simply put on ``sys.path``, by the repo-root
``clipgen.py`` launcher for real runs and by ``tests/conftest.py`` under pytest.

Two ways that drifts, both silent from the source tree:

* A module missing from ``[tool.setuptools] py-modules`` still imports fine
  locally, but ``uv pip install .`` ships only the listed modules, so installed
  and frozen environments die with ``ModuleNotFoundError``.
* A new module dropped in the *repo root* also imports fine locally (the root is
  still ``sys.path[1]``) and PyInstaller still bundles it (``pathex`` includes
  ``..``), so nothing fails until an install — by which time the flat root has
  quietly grown back.
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
