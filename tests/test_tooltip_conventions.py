"""Guard the single tooltip convention across every web page.

Tooltips have been re-fixed roughly eight times (`7e340085` double tooltip,
`ad0258d6` shared `.cg-tooltip`, `6f9c1d57` delegated singleton, `940160fa`
viewport clamping, `0e21c928` native `title` invisible on draggable rows,
`ed9323c9` hand-written key hints duplicating the registry's). The rules that
are mechanically checkable are asserted here.

Native ``title`` is *not* banned outright — ~150 uses are intentional. What is
banned is the combinations that actually shipped bugs.

Both static HTML and JS-constructed DOM are covered. The JS half matters most:
every instance of the draggable-row bug so far has been in a JS builder, which
is exactly where nothing was checking.
"""

import re

from _frontend_source import WEB, strip_comments

# One opening tag, non-greedy, no nested ">" — good enough for our hand-written
# HTML (no ">" appears inside an attribute value in assets/web).
_TAG_RE = re.compile(r"<[a-zA-Z][^>]*>")

# A trailing parenthesised key hint: "Play/Pause (Space)", "Set In (I)".
# Bounded length so prose parentheses ("(s)" aside, e.g. "no data for last 3s")
# don't trip it -- the hints are always short key names.
_KEY_HINT_RE = re.compile(r"\((?:[A-Za-z0-9]{1,6}|[^)]{0,12}[⌘⇧⌥⌃+][^)]{0,12})\)\s*$")

_LABEL_ATTR_RE = re.compile(r"(?:title|data-tooltip)=\"([^\"]*)\"")


def _html_files():
    return sorted(WEB.glob("*.html"))


def _tags():
    """Yield (path, tag_text) for every opening tag in every page."""
    for path in _html_files():
        for tag in _TAG_RE.finditer(path.read_text(encoding="utf-8")):
            yield path, tag.group(0)


def test_no_element_carries_both_title_and_data_tooltip():
    """Both attributes on one element renders two stacked tooltips (`7e340085`).

    Pick one: the `[data-tooltip]` singleton for styled/rich tooltips, or the
    native `title` for plain ones.
    """
    offenders = [
        f"{path.name}: {tag}"
        for path, tag in _tags()
        if "data-tooltip=" in tag and re.search(r"\stitle=", tag)
    ]
    assert not offenders, (
        "element carries both title= and data-tooltip=:\n" + "\n".join(offenders)
    )


def test_hotkey_elements_do_not_hand_write_the_key_hint():
    """The hotkey registry appends the resolved combo to a `[data-hotkey]`
    control's label on Alt-hold (see `resolvedCombos` in `hotkeys.js`).

    Hand-writing "(Space)" into the title as well renders the hint twice, and
    goes stale the moment the user rebinds the key. `ed9323c9` stripped these
    from Composer and Workflows; this keeps them from coming back anywhere.
    """
    offenders = []
    for path, tag in _tags():
        if "data-hotkey=" not in tag:
            continue
        for label in _LABEL_ATTR_RE.findall(tag):
            if _KEY_HINT_RE.search(label):
                offenders.append(f'{path.name}: "{label}"')
    assert not offenders, (
        "[data-hotkey] control hand-writes a key hint the registry already "
        "appends; drop the parenthesised hint:\n" + "\n".join(offenders)
    )


# `el.draggable = true` or `el.setAttribute("draggable", "true")`
_DRAGGABLE_RE = re.compile(
    r"\b([A-Za-z_$][\w$]*)\.(?:draggable\s*=\s*true"
    r"|setAttribute\(\s*[\"']draggable[\"']\s*,\s*[\"']true[\"'])"
)
# `el.title = ...` on that same identifier (not `document.title`).
_TITLE_ASSIGN_RE = re.compile(
    r"\b([A-Za-z_$][\w$]*)\.(?:title\s*=|setAttribute\(\s*[\"']title[\"'])"
)


def _js_functions(src: str):
    """Yield rough function bodies, so we only pair up names in the same scope.

    Splitting on `function` is crude, but a builder that makes an element
    draggable and sets its title does both within one function — which is all
    this needs to see. Being scope-local is what keeps it from pairing an
    unrelated `card.title` in a different builder.
    """
    parts = re.split(r"\bfunction\b", src)
    yield from parts[1:]


def test_tooltip_icon_sidecar_rides_along_with_a_tooltip():
    """`data-tooltip-icon` is opt-in decoration on the [data-tooltip] singleton.

    `showFor` in utils.js returns early when there is no `data-tooltip`, so an
    element carrying only the icon attribute shows nothing at all — a silent
    no-op with no error, exactly the failure mode this file exists to catch.
    """
    offenders = [
        f"{path.name}: {tag}"
        for path, tag in _tags()
        if "data-tooltip-icon=" in tag and not re.search(r"\sdata-tooltip=", tag)
    ]
    assert not offenders, (
        "data-tooltip-icon without data-tooltip never renders:\n" + "\n".join(offenders)
    )


def test_tooltip_icon_sidecar_names_an_icon_that_exists():
    """The sidecar is a CSS mask built from the attribute value; a mistyped name
    renders an invisible zero-content span beside the text, with no 404 visible
    to the user."""
    icons = WEB.parent / "icons"
    missing = []
    for path, tag in _tags():
        m = re.search(r'data-tooltip-icon="([^"]+)"', tag)
        if m and not (icons / f"{m.group(1)}.svg").is_file():
            missing.append(f"{path.name}: {m.group(1)}.svg")
    assert not missing, "data-tooltip-icon names a nonexistent icon:\n" + "\n".join(
        missing
    )


def test_native_title_is_not_set_on_a_draggable_element():
    """Native `title` does not render on an element with `draggable="true"`.

    The hint silently disappears, which is how `0e21c928` (Workflows palette
    rows) and the stash list shipped without their affordance text. Use the
    `[data-tooltip]` singleton on draggable elements instead.

    Only flags the *same identifier* getting both in one function — a `title`
    on a child of a draggable ancestor still renders and is not the bug.
    """
    offenders = []
    for path in sorted(WEB.glob("*.js")):
        src = strip_comments(path.read_text(encoding="utf-8"))
        for body in _js_functions(src):
            draggable = {m.group(1) for m in _DRAGGABLE_RE.finditer(body)}
            if not draggable:
                continue
            for m in _TITLE_ASSIGN_RE.finditer(body):
                if m.group(1) in draggable:
                    offenders.append(
                        f"{path.name}: {m.group(1)} is draggable and sets .title"
                    )
    assert not offenders, (
        "native title on a draggable element never renders; use data-tooltip:\n"
        + "\n".join(sorted(set(offenders)))
    )
