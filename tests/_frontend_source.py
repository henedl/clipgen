"""Shared helpers for the frontend source-level test modules.

Every ``test_*_frontend_source`` / ``*_wiring`` / ``*_source`` module needs the
same few things: the ``assets/web`` path, the hub+satellite glob-concat read,
the JS comment stripper, and the ES5-discipline assertion. One definition here;
the assertions themselves stay in the test files.
"""

import json
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


NODE_TIMEOUT = 20


def node_eval(snippet: str, probe: str, prelude: str = "") -> list[str]:
    """Run *snippet* + *probe* under node and return stdout lines.

    Pure-function regression tests slice a page script between two anchors
    (see ``slice_between``) and probe it with ``console.log``; *prelude* stubs
    whatever the slice needs (a fake ``document``, a helper it calls).
    """
    assert NODE is not None, "node missing; guard the test with skipif"
    result = subprocess.run(
        [NODE, "-e", prelude + "\n" + snippet + "\n" + probe],
        capture_output=True,
        text=True,
        timeout=NODE_TIMEOUT,
        check=False,
    )
    assert result.returncode == 0, f"node exited {result.returncode}:\n{result.stderr}"
    return result.stdout.split("\n")[:-1]


# Settle-or-fail: a rejection or a never-settled promise exits node nonzero.
_ASYNC_PROBE = """
var __settled = false;
process.on("exit", function (code) {
  if (!__settled && code === 0) {
    console.error("probe promise never settled");
    process.exitCode = 1;
  }
});
Promise.resolve().then(function () { return (%s); }).then(
  function (v) { __settled = true; console.log(JSON.stringify(v)); },
  function (e) { __settled = true; console.error((e && e.stack) || e); process.exitCode = 1; }
);
"""


def node_eval_async(snippet: str, expr: str, prelude: str = ""):
    """Await the promise *expr* under node and return its JSON-decoded value."""
    lines = node_eval(snippet, _ASYNC_PROBE % expr, prelude)
    assert lines, "async probe printed nothing"
    return json.loads(lines[-1])
