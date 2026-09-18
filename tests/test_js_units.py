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


# ---- Composer undo/redo history (composer.js) ----

_COMPOSER = read("composer.js")
_HISTORY = slice_between(
    _COMPOSER, "  var _undoStack = [];", "  // ---- User-facing cut actions"
)

# A fake server: create allocates a fresh id, edits on unknown ids reject.
_HISTORY_STUBS = """
var qs = function () { return {}; };
var server = { cuts: {}, anns: {}, next: 0 };
var fresh = function (prefix) { server.next += 1; return prefix + server.next; };
var known = function (store, id) {
  return id in store ? Promise.resolve() : Promise.reject(new Error("No " + id));
};
var applyCreate = function (cut) {
  var c = { id: fresh("cut_"), start: cut.start, end: cut.end };
  server.cuts[c.id] = c; return Promise.resolve(c);
};
var applyDelete = function (id) {
  return known(server.cuts, id).then(function () { delete server.cuts[id]; });
};
var applyTimes = function (id, t) {
  return known(server.cuts, id).then(function () { server.cuts[id].end = t.end; });
};
var applyAnnCreate = function (ann) {
  var a = { id: fresh("ann_"), text: ann.text };
  server.anns[a.id] = a; return Promise.resolve(a);
};
var applyAnnDelete = function (id) {
  return known(server.anns, id).then(function () { delete server.anns[id]; });
};
var applyAnnPatch = function (id, p) {
  return known(server.anns, id).then(function () { server.anns[id].text = p.text; });
};
var applyTrim = function () { return Promise.resolve(); };
var failures = [];
var opFailed = function (e) { failures.push(String(e && e.message)); };
var settle = function () { return new Promise(function (r) { setTimeout(r, 0); }); };
var steps = function (fns) {
  return fns.reduce(function (p, fn) { return p.then(fn).then(settle); }, Promise.resolve());
};
"""


def _run_history(script: str) -> dict:
    """Drive the sliced history code with *script* and dump the outcome."""
    probe = (
        "steps(["
        + script
        + "]).then(function () { console.log(JSON.stringify({"
        + "failures: failures, undo: _undoStack.length, redo: _redoStack.length,"
        + " cuts: Object.keys(server.cuts), anns: Object.keys(server.anns),"
        + " server: server })); });"
    )
    return json.loads(node_eval(_HISTORY, probe, _HISTORY_STUBS)[0])


def test_composer_history_survives_cut_delete_and_restore():
    """create → edit → delete, undo ×3 then redo ×3 must never hit a stale id."""
    out = _run_history(
        """
        function () { return applyCreate({start: 1, end: 2}).then(function (c) {
          recordOp({type: "create", cut: c});
          recordOp({type: "edit", id: c.id, before: {start: 1, end: 2}, after: {start: 1, end: 5}});
          return applyTimes(c.id, {end: 5});
        }); },
        function () { return applyDelete("cut_1").then(function () {
          recordOp({type: "delete", cut: {id: "cut_1", start: 1, end: 5}});
        }); },
        function () { undo(); }, function () { undo(); }, function () { undo(); },
        function () { redo(); }, function () { redo(); }, function () { redo(); }
        """
    )
    assert out["failures"] == []
    assert (out["undo"], out["redo"]) == (3, 0)
    assert out["cuts"] == []  # redo of the delete removed the restored cut


def test_composer_history_remaps_grouped_annotation_edits():
    """A group holding an ann-edit on a restored annotation follows the new id."""
    out = _run_history(
        """
        function () { return applyAnnCreate({text: "a"}).then(function (a) {
          recordOp({type: "ann-create", annotation: a});
          recordOp({type: "ann-group", ops: [
            {type: "ann-edit", id: a.id, field: "text", before: "a", after: "b"}
          ]});
          return applyAnnPatch(a.id, {text: "b"});
        }); },
        function () { return applyAnnDelete("ann_1").then(function () {
          recordOp({type: "ann-delete", annotation: {id: "ann_1", text: "b"}});
        }); },
        function () { undo(); }, function () { undo(); },
        function () { redo(); }
        """
    )
    assert out["failures"] == []
    assert (out["undo"], out["redo"]) == (2, 1)
    assert next(iter(out["server"]["anns"].values()))["text"] == "b"


# ---- Transcript inline-edit word diff (transcripts.js) ----

_TRANSCRIPTS = read("transcripts.js")
_EXTRACT = slice_between(
    _TRANSCRIPTS, "  function extractCorrections(", "  function saveCorrections("
)


def _diff(old: str, new: str) -> list:
    return _probe(
        "extractCorrections(" + json.dumps(old) + ", " + json.dumps(new) + ")",
        snippet=_EXTRACT,
    )


def test_extract_corrections_keeps_insertions_and_deletions():
    """Pure insert / delete groups surface with one empty side, never vanish."""
    assert _diff("I like it", "I really like it") == [{"from": "", "to": "really"}]
    assert _diff("I really like it", "I like it") == [{"from": "really", "to": ""}]
    assert _diff("I like it", "I love it") == [{"from": "like", "to": "love"}]
    assert _diff("I like it today", "I really love it") == [
        {"from": "like", "to": "really love"},
        {"from": "today", "to": ""},
    ]
    assert _diff("same text", "same  text") == []
