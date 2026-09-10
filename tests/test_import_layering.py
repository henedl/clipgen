"""Enforce the documented module-layer DAG (agents/ARCHITECTURE.md).

The screenspace siblings wire deepest-first; a cycle is broken with a
function-local import, never a top-level one. This is also why ruff's
I001 import sorting stays ignored in pyproject.toml: order is a contract.
"""

import ast
import re
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


def test_speakers_stays_import_light() -> None:
    """thinking_agents/data_export import it for names; numpy/onnxruntime load lazily."""
    project = _module_level_imports("speakers") & _PROJECT_MODULES
    assert project <= {"config", "utils", "profiling"}, project
    text = (SOURCE / "speakers.py").read_text(encoding="utf-8")
    assert not re.search(r"^import (numpy|onnxruntime)", text, re.MULTILINE)


# Whole-tree layers; module-level imports point strictly downhill.
_TREE_LAYERS = {
    "config": 0,
    "profiling": 1,
    "friction": 1,
    "utils": 2,
    "files": 3,
    "google_api": 3,
    "cli_args": 3,
    "excel_io": 3,
    "start_settings": 3,
    "changelog": 3,
    "licenses": 3,
    "mindnode": 3,
    "speakers": 3,
    "workflows_catalog": 3,
    "server_utils": 3,
    "manifest": 3,
    "native_dialogs": 3,
    "desktop_chrome": 4,
    "desktop_menu": 4,
    "video": 4,
    "spreadsheet": 4,
    "data_export": 4,
    "llm_client": 4,
    "updater": 4,
    "workflows_runner": 4,
    "titlecards": 5,
    "transcripts": 5,
    "viewer": 5,
    "interactive": 5,
    "thinking_agents": 5,
    "remux_server": 5,
    "workflows": 5,
    "pipeline": 6,
    "app": 7,
    "cli_event_clips": 7,
    "cli_screenspace": 7,
    "screenspace_server": 7,
    "transcripts_server": 7,
    "workflows_server": 7,
    "composer_server": 7,
    "overview": 7,
    "server": 8,
    "cli": 8,
    "desktop": 9,
}
_SCREENSPACE_FAMILY_LAYER = 5


def _tree_layer(stem: str) -> int | None:
    if stem in _TREE_LAYERS:
        return _TREE_LAYERS[stem]
    if stem.startswith("screenspace"):
        return _SCREENSPACE_FAMILY_LAYER
    return None


def test_every_module_has_a_layer() -> None:
    unplaced = sorted(s for s in _PROJECT_MODULES if _tree_layer(s) is None)
    assert not unplaced, f"assign a layer in _TREE_LAYERS: {unplaced}"


def test_modules_import_only_lower_layers() -> None:
    """The screenspace siblings order among themselves under the test above."""
    uphill: dict[str, set[str]] = {}
    for stem in sorted(_PROJECT_MODULES):
        layer = _tree_layer(stem)
        assert layer is not None
        for dep in _module_level_imports(stem) & _PROJECT_MODULES:
            if stem.startswith("screenspace") and dep.startswith("screenspace"):
                continue
            dep_layer = _tree_layer(dep)
            assert dep_layer is not None
            if dep_layer >= layer:
                uphill.setdefault(stem, set()).add(dep)
    assert not uphill, f"module-level imports going uphill: {uphill}"


def test_server_and_blueprints_stay_decoupled_at_import() -> None:
    """server.py mounts the blueprints; both sides reach across only inside functions."""
    blueprints = {
        "screenspace_server",
        "transcripts_server",
        "workflows_server",
        "composer_server",
        "overview",
    }
    assert not _module_level_imports("server") & blueprints
    for stem in blueprints:
        assert "server" not in _module_level_imports(stem), stem
