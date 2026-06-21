"""Guard that ``pyproject.toml``'s ``py-modules`` list stays in sync with the
top-level modules on disk.

This is a flat (single-package-less) project: every importable module is a
root-level ``*.py`` listed under ``[tool.setuptools] py-modules``. A module
that exists on disk but is missing from that list still imports fine from the
source tree (so ``uv run pytest`` passes), but ``uv pip install .`` ships only
the listed modules — so installed/binary environments die with
``ModuleNotFoundError``. This regression test makes that drift fail loudly the
next time a god-file is split into new modules.
"""

import tomllib
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent


def _root_modules() -> set[str]:
    """Top-level importable modules: root ``*.py`` minus private/dunder files."""
    return {p.stem for p in _ROOT.glob("*.py") if not p.name.startswith("_")}


def _listed_modules() -> set[str]:
    data = tomllib.loads((_ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    return set(data["tool"]["setuptools"]["py-modules"])


def test_py_modules_covers_every_root_module() -> None:
    missing = _root_modules() - _listed_modules()
    assert not missing, (
        "Root modules missing from pyproject [tool.setuptools] py-modules "
        f"(install would break `import` of these): {sorted(missing)}"
    )


def test_py_modules_has_no_phantom_entries() -> None:
    phantom = _listed_modules() - _root_modules()
    assert not phantom, (
        "pyproject py-modules lists modules with no matching root .py file: "
        f"{sorted(phantom)}"
    )
