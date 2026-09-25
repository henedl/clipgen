"""No new unreferenced function may appear in ``source/*.py``.

The Python twin of ``test_js_dead_functions.py``. Ruff and ty see unused
imports and variables, never an unused function: a helper can lose its last
caller in a refactor and live on indefinitely.

The signal is the same narrow one: a function or method whose name occurs
exactly **once** as a ``\\w+`` token across ``source/``, ``build/`` and
``clipgen.py``, that occurrence being its own definition. Dunder methods are
skipped (the interpreter calls them), and so is anything decorated: a Flask
route, a ``@property`` or a registry decorator is reached through the
decorator, not by name.

A name found once in production but also in ``tests/`` is **test-only**: a
fixture seam such as a cache reset. Those are legitimate but must be named in
``KNOWN_TEST_ONLY`` with a reason, so a helper kept alive only by its own unit
test fails too.

This is a **ratchet**: both lists may only shrink. Deleting a function means
deleting its entry in the same commit; a new orphan fails ``/check``.
"""

import re
from collections import Counter
from functools import cache
from pathlib import Path

from _py_source import PRODUCTION, ROOT, SOURCE, functions

_WORD = re.compile(r"\w+")

# Genuine orphans. Wave 1 of plans/CODE-QUALITY-PLAN.md emptied it; keep it so.
KNOWN_DEAD: frozenset[str] = frozenset()

# Production defines these for test fixtures only. The value says why it stays.
KNOWN_TEST_ONLY: dict[str, str] = {
    "_reset_manifest_cache": "fixtures drop the clipgen.json section cache",
    "_reset_screenspace_events_cache": "fixtures drop the viewer's events cache",
    "deep_reset": "profiling tests share one process",
    "ops_reset": "profiling tests share one process",
    "reset_for_tests": "updater tests restore import-time state",
}


def _read(paths: list[Path]) -> str:
    return "\n".join(path.read_text(encoding="utf-8") for path in paths)


def unreferenced(defining: list[Path], corpus: str) -> dict[str, str]:
    """Map each once-only, undecorated, non-dunder name to its ``file:line``.

    One tokenizing pass over the corpus, then a dict lookup per name: the
    per-name ``\\bname\\b`` scan is quadratic, as ``test_js_dead_functions`` found.
    """
    tokens = Counter(_WORD.findall(corpus))
    found: dict[str, str] = {}
    for path in defining:
        for fn in functions(path):
            if fn.decorated or (fn.name.startswith("__") and fn.name.endswith("__")):
                continue
            if tokens[fn.name] == 1:
                found.setdefault(fn.name, f"{path.name}:{fn.lineno}")
    return found


@cache
def _split() -> tuple[dict[str, str], dict[str, str]]:
    """(dead, test-only) for the real tree.

    This file is left out of the tests corpus: its own allowlist would
    otherwise make every listed name look test-referenced.
    """
    once = unreferenced(SOURCE, _read(PRODUCTION))
    this = Path(__file__).resolve()
    tests = [p for p in sorted((ROOT / "tests").rglob("*.py")) if p.resolve() != this]
    test_tokens = set(_WORD.findall(_read(tests)))
    dead = {n: s for n, s in once.items() if n not in test_tokens}
    test_only = {n: s for n, s in once.items() if n in test_tokens}
    return dead, test_only


def _listing(found: dict[str, str]) -> str:
    return "\n".join(f"  {name}  ({site})" for name, site in sorted(found.items()))


def test_no_new_unreferenced_functions():
    dead, _ = _split()
    unexpected = {n: s for n, s in dead.items() if n not in KNOWN_DEAD}
    assert not unexpected, (
        "These functions are defined in source/ and referenced nowhere. Call "
        "them or delete them:\n" + _listing(unexpected)
    )


def test_no_new_test_only_functions():
    _, test_only = _split()
    unexpected = {n: s for n, s in test_only.items() if n not in KNOWN_TEST_ONLY}
    assert not unexpected, (
        "These functions are called only from tests/. Delete the function and "
        "its tests, or, for a fixture seam such as a cache reset, add it to "
        "KNOWN_TEST_ONLY with a reason:\n" + _listing(unexpected)
    )


def test_baselines_have_no_stale_entries():
    """A name that is no longer an orphan must leave its list.

    Otherwise the lists become a graveyard of names deleted or revived long
    ago, and the next reader cannot tell which entries still mean anything.
    """
    dead, test_only = _split()
    stale_dead = sorted(KNOWN_DEAD - set(dead))
    stale_test = sorted(set(KNOWN_TEST_ONLY) - set(test_only))
    assert not stale_dead and not stale_test, (
        "Drop these entries; the functions were deleted, gained a production "
        f"caller, or changed category. KNOWN_DEAD: {stale_dead}; "
        f"KNOWN_TEST_ONLY: {stale_test}"
    )


def test_the_scan_detects_an_orphan(tmp_path):
    """Guard the detector: a broken scan must not read as a clean tree."""
    module = tmp_path / "sample.py"
    module.write_text(
        "def used():\n    return orphan_free()\n\n"
        "def orphan_free():\n    return 1\n\n"
        "def lonely():\n    return 2\n\n"
        "@decorator\ndef routed():\n    return 3\n\n"
        "class Box:\n    def __repr__(self):\n        return ''\n"
        "    def unused_method(self):\n        return 4\n",
        encoding="utf-8",
    )
    found = unreferenced([module], module.read_text(encoding="utf-8") + " used")
    assert set(found) == {"lonely", "unused_method"}
    assert sum(len(functions(p)) for p in SOURCE) > 1000, "scan saw no source"
