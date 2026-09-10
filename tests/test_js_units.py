"""Executed regression tests for pure utils.js helpers (node, no browser).

Slices are anchored on literal source text; a moved helper fails loudly here
rather than silently losing coverage.
"""

import json

import pytest

from _frontend_source import NODE, node_eval, read, slice_between

pytestmark = pytest.mark.skipif(NODE is None, reason="node not installed")

_UTILS = read("utils.js")
_TIME = slice_between(_UTILS, "var pad2 = ", "// ---- Elapsed-time")
_MARKDOWN = slice_between(
    _UTILS, "var clipgenRenderInlineMarkdown = ", "var hexToRgba = "
)
_PENDING = slice_between(_UTILS, "var isPending = ", "var _apiJson = ")
_HIGHLIGHT = slice_between(
    _UTILS, "var clipgenHighlightMatches = ", "// Escapes, then converts"
)
_ESCAPE_STUB = 'var escapeHtml = function (s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); };'
_CLAMP = slice_between(_UTILS, "var clamp = ", "\n\n")


def _probe(expr: str, prelude: str = "", snippet: str = _TIME) -> list:
    return json.loads(
        node_eval(snippet, f"console.log(JSON.stringify({expr}));", prelude)[0]
    )


def test_format_duration_pads_and_promotes_hours():
    out = _probe(
        "[formatDuration(0), formatDuration(65), formatDuration(3600), formatDuration(3661.6), formatDuration(null), formatDuration(NaN)]"
    )
    assert out == ["0:00", "1:05", "1:00:00", "1:01:02", "--:--", "--:--"]


def test_format_time_decimals_round_not_floor():
    out = _probe(
        "[formatTime(69.96, {decimals: 1}), formatTime(59.97, {decimals: 1}), formatTime(3599.9), formatTime(65.4)]"
    )
    assert out == ["1:10.0", "1:00.0", "59:59", "1:05"]


def test_highlight_matches_raw_text_not_entities():
    out = _probe(
        '[clipgenHighlightMatches("A&B", "&", "hl"), clipgenHighlightMatches("A&B", "amp", "hl"), clipgenHighlightMatches("<x>", "", "hl")]',
        prelude=_ESCAPE_STUB,
        snippet=_HIGHLIGHT,
    )
    assert out == ['A<span class="hl">&amp;</span>B', "A&amp;B", "&lt;x&gt;"]


def test_inline_markdown_escapes_before_marking_up():
    out = _probe(
        '[clipgenRenderInlineMarkdown("<b>x</b> `a<b` **bold** *em* snake_case"), clipgenRenderInlineMarkdown(null)]',
        prelude=_ESCAPE_STUB,
        snippet=_MARKDOWN,
    )
    assert (
        out[0]
        == "&lt;b&gt;x&lt;/b&gt; <code>a&lt;b</code> <strong>bold</strong> <em>em</em> snake_case"
    )
    assert out[1] == ""


def test_is_pending_reads_the_generating_flag_only():
    out = _probe(
        "[isPending({ok: false, generating: true}), isPending({ok: true, generating: true}), isPending({ok: false}), isPending(null)]",
        snippet=_PENDING,
    )
    assert out == [True, True, False, False]


def test_clamp_bounds_both_ends():
    out = _probe("[clamp(5, 0, 3), clamp(-1, 0, 3), clamp(2, 0, 3)]", snippet=_CLAMP)
    assert out == [3, 0, 2]
