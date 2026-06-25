# clipgen-split-module — Split a Python god-file into facade + siblings

Splitting a large module (e.g. `screenspace.py` → 9 `screenspace_*` modules behind a re-export facade) has twice shipped install-breaking or test-breaking gaps: a module missing from `py-modules` (`b036aec`, risked again in `9fbc762`) and an incomplete facade re-export (`5683a96`: `screenspace._probe_video_meta` → `AttributeError`). This is a flat, package-less layout, so the traps are specific. Follow this checklist.

## Checklist

1. **List every new module in `pyproject.toml`.** Add each new root `*.py` to `[tool.setuptools] py-modules`. The source tree imports fine without this (so `pytest` passes), but `uv pip install .` ships only listed modules and installed/frozen environments die with `ModuleNotFoundError`. Guarded by `tests/test_packaging.py`.

2. **Re-export *every* public name from the facade — and every test-touched private name.** The thin facade module (`screenspace.py`) must re-export each name former callers and tests reach via `facade.NAME`. Audit by grepping the test suite for `facade.` accesses (`screenspace.`), including underscore-prefixed helpers (`_probe_video_meta`).

3. **Patch the owning module, not the facade.** Re-exporting only **rebinds** a name on the facade — it does not propagate. `mock.patch("screenspace.foo")` patches the facade's binding, which the sibling that actually calls `foo` never sees. Point patch targets at the owning sibling (`screenspace_frames.foo`). Fix existing tests' patch targets when moving a function.

4. **Keep the import DAG acyclic, deepest-first.** Wire siblings in the order documented in `agents/ARCHITECTURE.md` (e.g. `primitives` → `ocr` → `frames` → `scans` → `tools` → `multitool`/`manifest` → `worker`). Break any cycle with a function-local import (as `MultitoolTool.scan` does for `scan_multitool`), not a top-level one.

## Verify

1. `uv run --extra dev pytest -c tests/pytest.ini tests/test_packaging.py` (py-modules + no-phantom guards).
2. `uv run --extra dev pytest -c tests/pytest.ini` (full suite — catches missing re-exports and stale patch targets).
3. `/check`.
4. If feasible, sanity-check the install path: `uv pip install .` into a throwaway venv and `import` the facade — the only thing the source-tree test cannot exercise.
