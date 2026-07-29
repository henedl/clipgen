"""Every ``var(--x)`` in the CSS must resolve to a property something defines.

An undefined custom property fails in one of two ways, and neither of them is
loud. Without a fallback the whole declaration is invalid and gets dropped, so
the element silently inherits — that is how ``.attachment-label`` in the timeline
viewer and four Screenspace hover rules lost their color and background outright.
With a fallback the literal simply wins forever, which is how Screenspace spent a
year rendering warnings in ``#f59e0b`` while every medium-severity chip elsewhere
in the app used ``#fbbf24``: the token the rules referenced had never existed.

Nothing else catches this. The browser reports no error, the page still boots so
the UI smoke passes, and a screenshot only shows it if you already know which
shade to expect. A one-pass scan is enough, so it runs in ``/check`` rather than
the opt-in browser harness.

Scope note: this checks that a name is *defined somewhere*, not that it is in
scope at the point of use. A property defined only on ``.co-icon-play`` and read
from an unrelated element would pass here and still resolve to nothing. That is a
narrower bug than the one this exists for, and catching it needs the runtime
census rather than a text scan.
"""

import re

from _frontend_source import WEB

# A declaration starts at the top of a block or after a semicolon. Anchoring on
# line-start instead would miss the single-line form the icon rules use
# (`.co-icon-play { --co-icon: url("icons/play.svg"); }`) and report every one of
# them as undefined.
_DEFINITION = re.compile(r"(?:^|[{;])\s*(--[\w-]+)\s*:", re.MULTILINE)
_REFERENCE = re.compile(r"var\(\s*(--[\w-]+)\s*(,?)")
# Custom properties written from code rather than declared in CSS.
_SET_PROPERTY = re.compile(r"""setProperty\(\s*["'`](--[\w-]+)""")
_INLINE_STYLE = re.compile(r"""["'`][^"'`]*?(--[\w-]+)\s*:""")


def _defined_in_css() -> set[str]:
    names: set[str] = set()
    for path in sorted(WEB.glob("*.css")):
        names |= set(_DEFINITION.findall(path.read_text(encoding="utf-8")))
    return names


def _set_at_runtime() -> set[str]:
    """Properties assigned from JS, HTML style attributes, or injected Python.

    ``--desktop-chrome-height`` and ``--desktop-traffic-inset`` are the live
    example: ``utils.render_index_html`` writes them onto ``<html>`` from
    ``config`` so the frameless macOS window can size its own title bar.
    """
    names: set[str] = set()
    for path in sorted(WEB.glob("*.js")) + sorted(WEB.glob("*.html")):
        text = path.read_text(encoding="utf-8")
        names |= set(_SET_PROPERTY.findall(text))
        names |= set(_INLINE_STYLE.findall(text))
    source = WEB.parent.parent / "source"
    for path in sorted(source.glob("*.py")):
        names |= set(_SET_PROPERTY.findall(path.read_text(encoding="utf-8")))
    return names


def _undefined_references() -> dict[str, list[str]]:
    """Map each unresolvable property to the ``file:line`` sites that read it."""
    known = _defined_in_css() | _set_at_runtime()
    found: dict[str, list[str]] = {}
    for path in sorted(WEB.glob("*.css")):
        for number, line in enumerate(
            path.read_text(encoding="utf-8").splitlines(), start=1
        ):
            for name, _ in _REFERENCE.findall(line):
                if name not in known:
                    found.setdefault(name, []).append(f"{path.name}:{number}")
    return found


def test_every_referenced_custom_property_is_defined():
    undefined = _undefined_references()
    assert not undefined, "CSS reads custom properties nothing defines:\n" + "\n".join(
        f"  {name} <- {', '.join(sites)}" for name, sites in sorted(undefined.items())
    )


def test_scan_sees_the_single_line_declaration_form():
    """Guard the regex itself.

    An earlier draft anchored definitions to line start and reported all ~30
    single-line ``--co-icon`` rules as undefined. The scan silently over-reporting
    is survivable; the fix for it — widening the allowlist — would have been the
    damaging part, so pin the shape that broke it.
    """
    names = _DEFINITION.findall('.co-icon-play { --co-icon: url("icons/play.svg"); }')
    assert names == ["--co-icon"]


def test_warning_token_resolves_to_the_shared_severity_amber():
    """``--color-warning`` must stay an alias, not regrow its own hex.

    It was introduced to end a split between Screenspace's ``#f59e0b`` warnings
    and the ``--severity-medium`` amber used everywhere else. Giving it a literal
    again would silently reopen that split.
    """
    tokens = (WEB / "tokens.css").read_text(encoding="utf-8")
    match = re.search(r"--color-warning\s*:\s*([^;]+);", tokens)
    assert match, "--color-warning is no longer defined in tokens.css"
    assert "var(--severity-medium)" in match.group(1)
