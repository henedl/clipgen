"""Shared helpers for the frontend source-level test modules.

Every ``test_*_frontend_source`` / ``*_wiring`` / ``*_source`` module needs the
same few things: the ``assets/web`` path, the hub+satellite glob-concat read,
the JS comment stripper, and the ES5-discipline assertion. One definition here;
the assertions themselves stay in the test files.
"""

import re
from pathlib import Path

WEB = Path(__file__).resolve().parent.parent / "assets" / "web"


def read(name: str) -> str:
    """Read one ``assets/web`` file as UTF-8 text."""
    return (WEB / name).read_text(encoding="utf-8")


def concat_js(prefix: str) -> str:
    """Concatenate a page's hub + satellite sources (``<prefix>*.js``, sorted).

    Satellites sort before the hub ("-" < "."), so ``src.index(a)..src.index(b)``
    slices still resolve within the single file that owns both anchors.
    """
    return "".join(
        p.read_text(encoding="utf-8") for p in sorted(WEB.glob(prefix + "*.js"))
    )


def strip_comments(src: str) -> str:
    """Drop ``/* */`` blocks and full-line ``//`` comments from a JS source."""
    src = re.sub(r"/\*.*?\*/", "", src, flags=re.DOTALL)
    return re.sub(r"^\s*//.*$", "", src, flags=re.MULTILINE)


def assert_es5(src: str, name: str) -> None:
    """House style: no arrow functions, no async/await.

    (``img.decoding = "async"`` is a DOM property, not the keyword — hence the
    word-boundary patterns rather than a bare substring check.)
    """
    assert "=>" not in src, f"{name} uses an arrow function"
    assert not re.search(r"\basync function\b|\bawait\s", src), (
        f"{name} uses async/await"
    )
