"""Enforce the documented module-layer DAG (agents/ARCHITECTURE.md).

The screenspace siblings wire deepest-first; a cycle is broken with a
function-local import, never a top-level one. This is also why ruff's
I001 import sorting stays ignored in pyproject.toml: order is a contract.
"""

import ast
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "source"

# Lower layers must never import higher ones.
_LAYERS = {
    "screenspace_primitives": 0,
    "screenspace_heatmap": 0,
    "screenspace_ocr": 1,
    "screenspace_frames": 2,
    "screenspace_scans": 3,
    "screenspace_tools": 4,
    "screenspace_multitool": 5,
    "screenspace_manifest": 5,
    "screenspace_worker": 6,
}

_PROJECT_MODULES = {path.stem for path in SOURCE.glob("*.py")}


def _module_level_imports(stem: str) -> set[str]:
    """Imported module names, skipping function bodies (sanctioned cycle breaks)."""
    tree = ast.parse((SOURCE / f"{stem}.py").read_text(encoding="utf-8"))
    found: set[str] = set()

    def visit(node: ast.AST) -> None:
        for child in ast.iter_child_nodes(node):
            if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                continue
            if isinstance(child, ast.Import):
                found.update(alias.name.split(".")[0] for alias in child.names)
            elif isinstance(child, ast.ImportFrom) and child.module:
                found.add(child.module.split(".")[0])
            visit(child)

    visit(tree)
    return found


def test_screenspace_siblings_respect_the_layer_order() -> None:
    for stem, layer in _LAYERS.items():
        siblings = _module_level_imports(stem) & _LAYERS.keys()
        too_high = {s for s in siblings if _LAYERS[s] >= layer}
        assert not too_high, f"{stem} (layer {layer}) imports {too_high}"


def test_primitives_and_friction_stay_pure() -> None:
    """primitives: no I/O or ffmpeg. friction: no LLM transport or I/O."""
    banned = {"subprocess", "requests", "urllib", "llm_client"}
    assert not _module_level_imports("screenspace_primitives") & banned
    friction = _module_level_imports("friction")
    assert not friction & banned
    assert not friction & _PROJECT_MODULES


def test_workflows_catalog_imports_only_config_and_utils() -> None:
    """Heavier deps (files, video) are late-imported inside adapters."""
    project = _module_level_imports("workflows_catalog") & _PROJECT_MODULES
    assert project <= {"config", "utils"}, project
