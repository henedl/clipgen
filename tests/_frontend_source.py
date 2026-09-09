"""Shared helpers for the frontend source-level test modules.

Every ``test_*_frontend_source`` / ``*_wiring`` / ``*_source`` module needs the
same few things: the ``assets/web`` path, the hub+satellite glob-concat read,
the JS comment stripper, and the ES5-discipline assertion. One definition here;
the assertions themselves stay in the test files.
"""

import re
import shutil
import subprocess
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


NODE = shutil.which("node")


def slice_between(src: str, start: str, end: str) -> str:
    """The source between two literal anchors (start inclusive, end exclusive)."""
    i = src.index(start)
    return src[i : src.index(end, i)]


def node_eval(snippet: str, probe: str, prelude: str = "") -> list[str]:
    """Run *snippet* + *probe* under node and return stdout lines.

    Pure-function regression tests slice a page script between two anchors
    (see ``slice_between``) and probe it with ``console.log``; *prelude* stubs
    whatever the slice needs (a fake ``document``, a helper it calls).
    """
    assert NODE is not None, "node missing; guard the test with skipif"
    result = subprocess.run(
        [NODE, "-e", prelude + "\n" + snippet + "\n" + probe],
        check=True,
        capture_output=True,
        text=True,
    )
    return result.stdout.split("\n")[:-1]
