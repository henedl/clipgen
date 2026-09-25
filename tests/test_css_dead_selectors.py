"""No new unreferenced class selector may appear in ``assets/web/*.css``.

A rule whose class nothing ever sets is invisible to every other check: the
page renders the same, the token ratchets still count it, and it survives
each refactor because deleting CSS feels risky.

A class counts as referenced when any ``.js``, ``.html`` or ``source/*.py``
file names it as a ``[\\w-]+`` token. Classes built by concatenation
(``"task-card-" + task.status``, ``f"sev-{label}"``) never appear whole, so a
string literal ending in ``-`` that feeds a ``+``, ``{``, or ``%s`` registers
its tail as a prefix, and any class starting with one counts as referenced.
That is deliberately generous: a missed orphan costs a few bytes, a false
alarm teaches people to baseline without looking.

This is a **ratchet**: ``KNOWN_ORPHANS`` names the orphans per file and may
only shrink. Deleting a rule means deleting its entry in the same commit.
"""

import re
from functools import cache

from _frontend_source import WEB

ROOT = WEB.parent.parent

_COMMENT = re.compile(r"/\*.*?\*/", re.DOTALL)
_STRING = re.compile(r"\"[^\"]*\"|'[^']*'")
_PRELUDE = re.compile(r"([^{};]*)\{")
_CLASS = re.compile(r"\.(-?[A-Za-z_][\w-]*)")
_TOKEN = re.compile(r"[\w-]+")
# The glue after a dash-ending literal: close-quote then +, f-string {, or %.
_BUILD = re.compile(r"-(?:[\"']\s*\+|\{|%s|%\()")
_TAIL = re.compile(r"[A-Za-z_][\w-]*$")

# Orphan classes per stylesheet. Wave 1 of plans/CODE-QUALITY-PLAN.md emptied it.
KNOWN_ORPHANS: dict[str, frozenset[str]] = {}


def css_classes(css: str) -> set[str]:
    """Class names used in selectors, ignoring comments, at-rules and strings."""
    names: set[str] = set()
    for match in _PRELUDE.finditer(_COMMENT.sub("", css)):
        prelude = match.group(1).strip()
        if prelude and not prelude.startswith("@"):
            names.update(_CLASS.findall(_STRING.sub('""', prelude)))
    return names


def references(text: str) -> tuple[set[str], tuple[str, ...]]:
    """(tokens, concatenation prefixes) named anywhere in ``text``.

    Prefixes come from anchoring on the glue and reading back, which is ~50x
    faster than a forward regex that tries every identifier start.
    """
    prefixes = set()
    for match in _BUILD.finditer(text):
        tail = _TAIL.search(text, max(0, match.start() - 80), match.start() + 1)
        if tail:
            prefixes.add(tail.group())
    return set(_TOKEN.findall(text)), tuple(sorted(prefixes))


def orphans(classes: set[str], tokens: set[str], prefixes: tuple[str, ...]) -> set[str]:
    return {
        name for name in classes if name not in tokens and not name.startswith(prefixes)
    }


@cache
def _scan() -> tuple[dict[str, frozenset[str]], int, tuple[str, ...]]:
    """(orphans per stylesheet, total classes seen, prefixes); clean files omitted."""
    consumers = (
        sorted(WEB.glob("*.js"))
        + sorted(WEB.glob("*.html"))
        + sorted((ROOT / "source").glob("*.py"))
    )
    tokens, prefixes = references(
        "\n".join(p.read_text(encoding="utf-8") for p in consumers)
    )
    found = {}
    seen = 0
    for path in sorted(WEB.glob("*.css")):
        classes = css_classes(path.read_text(encoding="utf-8"))
        seen += len(classes)
        dead = orphans(classes, tokens, prefixes)
        if dead:
            found[path.name] = frozenset(dead)
    return found, seen, prefixes


def test_no_new_orphan_selectors():
    new = {
        name: sorted(dead - KNOWN_ORPHANS.get(name, frozenset()))
        for name, dead in _scan()[0].items()
    }
    new = {name: dead for name, dead in new.items() if dead}
    assert not new, (
        "These classes are styled but no JS, HTML or Python file sets them. "
        "Delete the rules, or set the class where it belongs:\n"
        + "\n".join(
            f"  {name}: {', '.join(dead)}" for name, dead in sorted(new.items())
        )
    )


def test_known_orphans_have_no_stale_entries():
    """A class that is no longer an orphan must leave the baseline."""
    found = _scan()[0]
    stale = {
        name: sorted(dead - found.get(name, frozenset()))
        for name, dead in KNOWN_ORPHANS.items()
    }
    stale = {name: dead for name, dead in stale.items() if dead}
    assert not stale, (
        f"KNOWN_ORPHANS lists classes that were deleted or are now set. Drop them: {stale}"
    )


def test_the_scan_detects_an_orphan():
    """Guard the detector: a broken scan must not read as a clean tree."""
    css = (
        "/* .commented-out {} */\n"
        ".used-card, .card-state-open:hover { color: red; }\n"
        "@media (width > 1px) { .lonely-rule > .used-card { top: 0; } }\n"
        '[data-x=".not-a-class"] { margin: 0.5rem; }\n'
    )
    tokens, prefixes = references(
        'el("div", "used-card"); el("span", "card card-state-" + s);'
    )
    assert "card-state-" in prefixes
    assert orphans(css_classes(css), tokens, prefixes) == {"lonely-rule"}


def test_the_scan_sees_the_real_tree():
    """The live corpus must yield many classes and a known dynamic prefix."""
    _, seen, prefixes = _scan()
    assert seen > 1000
    assert "task-card-" in prefixes
