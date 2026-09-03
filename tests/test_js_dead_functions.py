"""No new unreferenced top-level function may appear in ``assets/web/*.js``.

Nothing else in the suite sees dead JS: ``node --check`` proves a file parses and
``test_frontend_satellite_wiring.py`` proves a call *resolves*, but neither asks
whether anything calls a function at all — so one can be written, wired to
nothing, and live in the bundle indefinitely.

The signal is deliberately narrow: a top-level ``function name(...)`` whose name
occurs exactly **once** across every ``.js``, ``.html`` and ``source/*.py`` file,
that occurrence being its own definition. At that threshold there is no judgement
call left — nothing references it under any spelling, including the
``window.Clipgen*`` publishes and the guarded hub delegators.

The looser "never appears in call position" variant measured ~174 hits, dominated
by legitimate handlers passed by reference and namespace exports. A decent
human-reviewed worklist and a terrible gate, so it is not implemented.

This is a **ratchet**, not a cleanup: the known-dead functions are listed so the
current state passes, and the list may only shrink. Deleting one means deleting
its entry in the same commit; adding another fails ``/check``.
"""

import re
from collections import Counter
from functools import cache
from pathlib import Path

from _frontend_source import WEB

_TOP_LEVEL_FUNCTION = re.compile(
    r"^\s*function\s+([A-Za-z_$][\w$]*)\s*\(", re.MULTILINE
)
_WORD = re.compile(r"\w+")

# Unreferenced as of the frontend-cleanup pass. Each is a genuine deletion
# candidate, tracked in the plan's drawdown step rather than removed here so this
# commit stays test-only. Shrink this set; never grow it.
KNOWN_DEAD: frozenset[str] = frozenset()


def _consumer_text() -> str:
    """Everything that could plausibly name a frontend function."""
    parts = [path.read_text(encoding="utf-8") for path in sorted(WEB.glob("*.js"))]
    parts += [path.read_text(encoding="utf-8") for path in sorted(WEB.glob("*.html"))]
    source = Path(WEB).parent.parent / "source"
    parts += [path.read_text(encoding="utf-8") for path in sorted(source.glob("*.py"))]
    return "\n".join(parts)


@cache
def _unreferenced() -> dict[str, str]:
    """Map each unreferenced function name to the ``file:line`` that defines it.

    Counted from a single tokenizing pass rather than one ``\\bname\\b`` scan per
    name: the corpus is ~4 MB and there are ~2000 names, so the per-name form cost
    69 s a call — and all three tests below call this. For a name made only of
    ``\\w`` characters the two are equivalent by definition, since ``\\b`` on both
    ends matches exactly when the name is a whole ``\\w+`` token. A ``$`` in a name
    breaks that equivalence (``$`` is not a ``\\w`` character), so those keep the
    original scan; there are none today, which is why the fast path is the one
    worth having.

    Cached because it is pure over files that cannot change mid-run, and the
    result is a shared read-only mapping — no test mutates it.
    """
    corpus = _consumer_text()
    definitions: dict[str, str] = {}
    for path in sorted(WEB.glob("*.js")):
        text = path.read_text(encoding="utf-8")
        for match in _TOP_LEVEL_FUNCTION.finditer(text):
            line = text.count("\n", 0, match.start()) + 1
            definitions.setdefault(match.group(1), f"{path.name}:{line}")

    tokens = Counter(_WORD.findall(corpus))

    def _occurrences(name: str) -> int:
        if "$" in name:
            return len(re.findall(rf"\b{re.escape(name)}\b", corpus))
        return tokens[name]

    return {name: site for name, site in definitions.items() if _occurrences(name) == 1}


def test_no_new_unreferenced_functions():
    unexpected = {
        name: site for name, site in _unreferenced().items() if name not in KNOWN_DEAD
    }
    assert not unexpected, (
        "These top-level functions are defined and referenced nowhere. Wire them "
        "up or delete them:\n"
        + "\n".join(f"  {name}  ({site})" for name, site in sorted(unexpected.items()))
    )


def test_known_dead_list_has_no_stale_entries():
    """A name that is no longer dead must leave the list.

    Without this the set silently becomes a graveyard of names that were deleted
    or revived years ago, and the next reader cannot tell which entries still
    mean anything.
    """
    stale = sorted(KNOWN_DEAD - set(_unreferenced()))
    assert not stale, (
        "KNOWN_DEAD lists functions that are no longer unreferenced (deleted, or "
        f"now called). Drop them from the set: {', '.join(stale)}"
    )


def test_the_scan_finds_the_known_dead_functions():
    """Guard the detector.

    If the regex or the corpus glob silently stops matching, both tests above go
    green while checking nothing. Pinning that the five known names are still
    *found* keeps a broken scan from reading as a clean codebase.
    """
    assert KNOWN_DEAD <= set(_unreferenced())
