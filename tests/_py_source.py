"""Shared AST walk for the Python source-level ratchets.

``test_py_dead_functions`` and ``test_function_length`` both need every function
definition in the production tree. Parsing it costs ~0.2 s, so each file is
parsed once per process and the records are shared read-only.
"""

import ast
from functools import cache
from pathlib import Path
from typing import NamedTuple

ROOT = Path(__file__).resolve().parent.parent
SOURCE = sorted((ROOT / "source").glob("*.py"))
# Top level only: build/lib/ holds setuptools copies of source/ after an install.
PRODUCTION = SOURCE + sorted((ROOT / "build").glob("*.py")) + [ROOT / "clipgen.py"]

_NESTED = ("body", "orelse", "finalbody", "handlers", "cases")


class PyFunction(NamedTuple):
    qualname: str
    name: str
    lineno: int
    length: int
    decorated: bool


def _walk(nodes: list, prefix: str, out: list[PyFunction]) -> None:
    """Collect defs from statement lists only; expressions cannot hold a ``def``.

    Skipping expression nodes is what makes this 10x cheaper than ``ast.walk``.
    """
    for node in nodes:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            qualname = prefix + node.name
            if not isinstance(node, ast.ClassDef):
                length = (node.end_lineno or node.lineno) - node.lineno + 1
                out.append(
                    PyFunction(
                        qualname,
                        node.name,
                        node.lineno,
                        length,
                        bool(node.decorator_list),
                    )
                )
            _walk(node.body, qualname + ".", out)
            continue
        for field in _NESTED:
            child = getattr(node, field, None)
            if isinstance(child, list):
                _walk(child, prefix, out)


@cache
def functions(path: Path) -> tuple[PyFunction, ...]:
    """Every function and method defined in ``path``, nested ones included."""
    out: list[PyFunction] = []
    _walk(ast.parse(path.read_text(encoding="utf-8")).body, "", out)
    return tuple(out)
