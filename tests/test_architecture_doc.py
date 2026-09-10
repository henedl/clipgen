"""The ARCHITECTURE.md ``assets/web/`` inventory names every carved satellite.

Agents pick a file from that row when deciding where a function lives, and it
has drifted twice (a satellite landed, the row was edited later, the name never
appeared). Names may be spelled out or brace-expanded (``studio-{a,b}.js``).
"""

import re
from pathlib import Path

import pytest

from _frontend_source import WEB
from test_frontend_satellite_wiring import _GROUPS

_DOC = Path(__file__).resolve().parent.parent / "agents" / "ARCHITECTURE.md"
_BRACE = re.compile(r"([a-z]+)-\{([^}]*)\}\.js")


def _mentioned_files(text: str) -> set[str]:
    names = set(re.findall(r"[a-z][a-z0-9-]*\.js", text))
    for prefix, body in _BRACE.findall(text):
        names.update(f"{prefix}-{part.strip()}.js" for part in body.split(","))
    return names


@pytest.mark.parametrize("group, pattern", sorted(_GROUPS.items()))
def test_every_satellite_is_documented(group: str, pattern: str) -> None:
    mentioned = _mentioned_files(_DOC.read_text(encoding="utf-8"))
    satellites = {p.name for p in WEB.glob(pattern) if "-" in p.name}
    missing = sorted(satellites - mentioned)
    assert not missing, f"{group}: add to agents/ARCHITECTURE.md: {missing}"
