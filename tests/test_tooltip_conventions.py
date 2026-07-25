"""Guard the single tooltip convention across every web page.

Tooltips have been re-fixed roughly eight times (`7e340085` double tooltip,
`ad0258d6` shared `.cg-tooltip`, `6f9c1d57` delegated singleton, `940160fa`
viewport clamping, `0e21c928` native `title` invisible on draggable rows,
`ed9323c9` hand-written key hints duplicating the registry's). The rules that
are mechanically checkable are asserted here.

Native ``title`` is *not* banned outright — ~150 uses are intentional. What is
banned is the two combinations that actually shipped bugs.
"""

import re

from _frontend_source import WEB

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
