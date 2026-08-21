"""Parse every ``assets/web/*.js`` file with ``node --check``.

The cheapest possible frontend gate, and until now the only one nobody ran: a
stray unbalanced brace in a page script was invisible to the whole suite, since
every other frontend test greps source text rather than parsing it.
``node --check`` is a real parser, needs no config file, and costs ~1.6 s across
all of ``assets/web``.

This is deliberately *not* part of the opt-in browser harness in ``tests/ui``
(see ``agents/skills/ui-check/SKILL.md``) — it belongs in ``/check`` and in CI,
where GitHub's runners ship node. It skips cleanly where node is absent.

Syntax is all it proves. The cross-file ``ReferenceError`` class that
``test_frontend_satellite_wiring.py`` approximates statically, and that
``tests/ui`` catches at runtime, is out of scope here.
"""

import shutil
import subprocess

import pytest

from _frontend_source import WEB, assert_es5, strip_comments

NODE = shutil.which("node")

# Sorted so the parametrized ids are stable across machines.
JS_FILES = sorted(p.name for p in WEB.glob("*.js"))


def test_js_files_were_discovered() -> None:
    """Guard the glob itself: a silently empty parametrize would pass vacuously."""
    assert len(JS_FILES) > 40, f"expected the full assets/web bundle, got {JS_FILES}"


@pytest.mark.skipif(NODE is None, reason="node not installed; JS syntax gate skipped")
@pytest.mark.parametrize("name", JS_FILES)
def test_js_parses(name: str) -> None:
    assert NODE is not None  # narrowed by the skipif above
    result = subprocess.run(
        [NODE, "--check", str(WEB / name)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, f"{name} is not valid JavaScript:\n{result.stderr}"


@pytest.mark.parametrize("name", JS_FILES)
def test_js_is_es5(name: str) -> None:
    """House style: every page script is ES5 (no arrows, no async/await)."""
    assert_es5(strip_comments((WEB / name).read_text(encoding="utf-8")), name)
