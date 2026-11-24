"""Executed regression tests for pure utils.js helpers (node, no browser).

Slices are anchored on literal source text; a moved helper fails loudly here
rather than silently losing coverage.
"""

import json

import pytest

from _frontend_source import NODE, node_eval, node_eval_async, read, slice_between

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
_NDJSON = slice_between(
    _UTILS, "// ---- NDJSON streaming reader ----", "// ---- File downloads ----"
)


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


_TASK_PICK = slice_between(
    _TRANSCRIPTS,
    "  function _taskForSelectedParticipant(",
    "  // Diarization rides the same task list",
)
_TASK_PICK_STUBS = r"""
var state = { selectedParticipant: "P01", tasks: [] };
function _isSpeakerTask(t) { return t.kind === "speakers"; }
function _isRedactTask(t) { return t.kind === "redact"; }
function pick(tasks) { state.tasks = tasks; var t = _taskForSelectedParticipant(); return t && t.id; }
"""


def test_status_task_pick_prefers_active_then_newest():
    """A retry that completed must replace the older failure in the indicator."""
    out = _probe(
        "["
        'pick([{id: "a", participant: "P01", status: "failed", created_at: 1},'
        ' {id: "b", participant: "P01", status: "completed", created_at: 2}]),'
        'pick([{id: "a", participant: "P01", status: "completed", created_at: 1},'
        ' {id: "b", participant: "P01", status: "failed", created_at: 2}]),'
        'pick([{id: "a", participant: "P01", status: "queued", created_at: 3},'
        ' {id: "b", participant: "P01", status: "running", created_at: 1},'
        ' {id: "c", participant: "P01", status: "failed", created_at: 4}]),'
        'pick([{id: "a", participant: "P01", status: "completed", created_at: 1},'
        ' {id: "b", participant: "P01", status: "running", created_at: 2, kind: "redact"}])'
        "]",
        prelude=_TASK_PICK_STUBS,
        snippet=_TASK_PICK,
    )
    assert out == ["b", "b", "b", "a"]


_EDIT_TEXT = slice_between(
    _TRANSCRIPTS, "  function _segmentEditText(", "  function startSegmentEditing("
)

# A .segment-text node whose first child is the mark label badge.
_FAKE_TEXT_EL = r"""
function fakeTextEl(badgeText, wordText) {
  var parts = badgeText === null ? [wordText] : [badgeText, wordText];
  return { cloneNode: function () {
    var kids = parts.slice();
    return {
      querySelector: function () {
        return badgeText === null ? null : { remove: function () { kids.shift(); } };
      },
      get textContent() { return kids.join(""); },
    };
  } };
}
"""


def test_segment_edit_text_drops_the_mark_badge():
    """The label badge must not leak into a saved correction."""
    out = _probe(
        '[_segmentEditText(fakeTextEl("Friction · confusion", "I can not find it")),'
        ' _segmentEditText(fakeTextEl(null, "plain line"))]',
        prelude=_FAKE_TEXT_EL,
        snippet=_EDIT_TEXT,
    )
    assert out == ["I can not find it", "plain line"]


# ---- NDJSON streaming (utils.js) ----

# Each chunk is a string, a byte array, or an Error the reader rejects with.
_STREAM_STUBS = r"""
var enc = new TextEncoder();
function fakeResponse(chunks) {
  var i = 0;
  return { ok: true, status: 200, body: { getReader: function () {
    return { read: function () {
      if (i >= chunks.length) return Promise.resolve({ done: true });
      var c = chunks[i++];
      if (c instanceof Error) return Promise.reject(c);
      return Promise.resolve({ done: false, value: typeof c === "string" ? enc.encode(c) : c });
    } };
  } } };
}
function outcome(p, lines) {
  return p.then(
    function () { return { lines: lines }; },
    function (e) { return { lines: lines, error: e.message, name: e.name, status: e.status, bodyText: e.bodyText }; }
  );
}
function collect(chunks) {
  var lines = [];
  return outcome(readNDJSONStream(fakeResponse(chunks), function (l) { lines.push(l); }), lines);
}
function all(cases) {
  var keys = Object.keys(cases), out = {};
  return keys.reduce(function (p, k) {
    return p.then(cases[k]).then(function (v) { out[k] = v; });
  }, Promise.resolve()).then(function () { return out; });
}
// Bytes 6..8 hold the euro sign; slicing at 7 splits it.
var euro = enc.encode('{"t":"€"}\n');
"""


def test_read_ndjson_stream_splits_buffers_and_trims():
    out = node_eval_async(
        _NDJSON,
        """all({
          several: function () { return collect(['{"a":1}\\n{"a":2}\\n{"a":3}\\n']); },
          split: function () { return collect(['{"a":', '1}', '\\n']); },
          multibyte: function () { return collect([euro.slice(0, 7), euro.slice(7)]); },
          blanks: function () { return collect(['\\n  {"a":1}  \\r\\n\\r\\n\\n', '\\t{"b":2}\\n']); },
          tail: function () { return collect(['{"a":1}\\n{"a"', ':2}  ']); },
          empty: function () { return collect([]); },
          raw: function () { return collect(['not json\\n']); },
          types: function () {
            var seen = [];
            return readNDJSONStream(fakeResponse(['{"a":1}\\n']), function (l) { seen.push(typeof l); })
              .then(function () { return seen; });
          }
        })""",
        _STREAM_STUBS,
    )
    assert out["several"] == {"lines": ['{"a":1}', '{"a":2}', '{"a":3}']}
    assert out["split"] == {"lines": ['{"a":1}']}
    assert out["multibyte"] == {"lines": ['{"t":"€"}']}
    assert out["blanks"] == {"lines": ['{"a":1}', '{"b":2}']}
    assert out["tail"] == {"lines": ['{"a":1}', '{"a":2}']}
    assert out["empty"] == {"lines": []}
    assert out["raw"] == {"lines": ["not json"]}
    assert out["types"] == ["string"]


def test_read_ndjson_stream_rejects_on_failures():
    out = node_eval_async(
        _NDJSON,
        """all({
          noBody: function () { return outcome(readNDJSONStream({}, function () {}), []); },
          noReader: function () { return outcome(readNDJSONStream({ body: {} }, function () {}), []); },
          readFails: function () { return collect(['{"a":1}\\n', new Error("boom")]); },
          onLineThrows: function () {
            return outcome(readNDJSONStream(fakeResponse(['x\\n']), function () {
              throw new Error("bad line");
            }), []);
          }
        })""",
        _STREAM_STUBS,
    )
    unsupported = {
        "lines": [],
        "error": "Streaming response not supported",
        "name": "Error",
    }
    assert out["noBody"] == unsupported
    assert out["noReader"] == unsupported
    assert out["readFails"] == {"lines": ['{"a":1}'], "error": "boom", "name": "Error"}
    assert out["onLineThrows"] == {"lines": [], "error": "bad line", "name": "Error"}


_FETCH_STUBS = """
var calls = [];
var reply;
var fetch = function (url, init) { calls.push({ url: url, init: init }); return reply(); };
var abortErr = new DOMException("The operation was aborted.", "AbortError");
function errorReply(status, text) {
  return function () {
    return Promise.resolve({ ok: false, status: status, text: text });
  };
}
function post(opts) {
  var lines = [];
  opts = opts || {};
  opts.onLine = opts.onLine === false ? undefined : function (l) { lines.push(l); };
  return outcome(apiPostNDJSON("api/generate", { cells: [1, 2] }, opts), lines);
}
"""


def test_api_post_ndjson_sends_request_and_streams():
    out = node_eval_async(
        _NDJSON,
        """(function () {
          var ac = new AbortController();
          reply = function () { return Promise.resolve(fakeResponse(['{"a":1}\\n', '{"b":2}'])); };
          return post({ signal: ac.signal }).then(function (res) {
            var c = calls[0];
            res.request = { url: c.url, method: c.init.method, headers: c.init.headers,
              body: c.init.body, sameSignal: c.init.signal === ac.signal, count: calls.length };
            return post({ onLine: false }).then(function (bare) { res.bare = bare; return res; });
          });
        })()""",
        _STREAM_STUBS + _FETCH_STUBS,
    )
    assert out["request"] == {
        "url": "api/generate",
        "method": "POST",
        "headers": {"Content-Type": "application/json"},
        "body": '{"cells":[1,2]}',
        "sameSignal": True,
        "count": 1,
    }
    assert out["lines"] == ['{"a":1}', '{"b":2}']
    assert "error" not in out
    assert out["bare"] == {"lines": []}


def test_api_post_ndjson_errors_and_aborts():
    out = node_eval_async(
        _NDJSON,
        """all({
          http: function () {
            reply = errorReply(409, function () { return Promise.resolve('{"error":"busy"}'); });
            return post();
          },
          bodyFails: function () {
            reply = errorReply(502, function () { return Promise.reject(new Error("gone")); });
            return post();
          },
          fetchAbort: function () {
            reply = function () { return Promise.reject(abortErr); };
            return apiPostNDJSON("x", {}).then(null, function (e) { return e === abortErr; });
          },
          readAbort: function () {
            reply = function () { return Promise.resolve(fakeResponse(['{"a":1}\\n', abortErr])); };
            return apiPostNDJSON("x", {}).then(null, function (e) { return e === abortErr; });
          }
        })""",
        _STREAM_STUBS + _FETCH_STUBS,
    )
    assert out["http"] == {
        "lines": [],
        "error": "Server error 409",
        "name": "Error",
        "status": 409,
        "bodyText": '{"error":"busy"}',
    }
    assert out["bodyFails"] == {
        "lines": [],
        "error": "Server error 502",
        "name": "Error",
        "status": 502,
        "bodyText": "",
    }
    assert out["fetchAbort"] is True
    assert out["readAbort"] is True


# ---- Studio generate cancellation (studio-generate.js) ----

_GENERATE = read("studio-generate.js")
_CANCEL = slice_between(
    _GENERATE,
    "  function isGenerateFetchAborted(err) {",
    "  // Sheet and intake halves",
) + slice_between(
    _GENERATE, "  function onCancelGenerate() {", "  STUDIO.onGenerate = onGenerate;"
)

# A rejected cancel post with no handler crashes node, failing the probe.
_CANCEL_STUBS = """
var state = {};
var hidden = [];
var qs = function (sel) {
  return { classList: { add: function (c) { hidden.push(sel + "." + c); } } };
};
var posts = [];
var apiPost = function (path) { posts.push(path); return Promise.reject(new Error("offline")); };
var tick = function () { return new Promise(function (r) { setTimeout(r, 0); }); };
"""


def test_on_cancel_generate_aborts_every_controller():
    out = node_eval_async(
        _CANCEL,
        """(function () {
          var first = new AbortController(), last = new AbortController();
          var hit = [];
          var broken = { abort: function () { hit.push("broken"); throw new Error("x"); } };
          state.activeGenerateAborts = [first, broken, last];
          onCancelGenerate();
          var res = {
            aborted: [first.signal.aborted, hit, last.signal.aborted],
            cancelled: state.generateCancelledByUser,
            hidden: hidden.slice(),
            left: state.activeGenerateAborts,
            posts: posts.slice()
          };
          state.activeGenerateAborts = [];
          onCancelGenerate();
          delete state.activeGenerateAborts;
          onCancelGenerate();
          res.afterAbsent = state.activeGenerateAborts;
          res.postCount = posts.length;
          return tick().then(function () { return res; });
        })()""",
        _CANCEL_STUBS,
    )
    assert out["aborted"] == [True, ["broken"], True]
    assert out["cancelled"] is True
    assert out["hidden"] == ["#cancelGenerateBtn.hidden"]
    assert out["left"] == []
    assert out["posts"] == ["api/generate/cancel", "api/generate-intake/cancel"]
    assert out["afterAbsent"] == []
    assert out["postCount"] == 6


def test_is_generate_fetch_aborted_classifies_errors():
    out = node_eval_async(
        _CANCEL,
        """(function () {
          var abortErr = new DOMException("aborted", "AbortError");
          var plain = new Error("HTTP 500");
          var res = [isGenerateFetchAborted(abortErr), isGenerateFetchAborted(plain),
            isGenerateFetchAborted(null)];
          state.generateCancelledByUser = true;
          res.push(isGenerateFetchAborted(plain));
          return res;
        })()""",
        _CANCEL_STUBS,
    )
    assert out == [True, False, False, True]


_RESULT_ROW = slice_between(
    read("screenspace-results.js"),
    "function buildResultRow(",
    "// Which detectors offer a certainty cutoff",
)
_RESULT_ROW_STUBS = """
function el(tag, cls, text) {
  var kids = [];
  return { dataset: {}, cls: cls, text: text || "", kids: kids,
    appendChild: function (c) { kids.push(c); } };
}
function formatTime(t) { return String(t); }
"""


def test_result_row_accepts_streamed_color_instants():
    """A running color scan streams {timestamp, _confidence}; treating it as a
    span threw on r.duration and stalled the Results panel."""
    out = _probe(
        "[buildResultRow({timestamp: 4, _confidence: 0.9}, 0, null, false, {type: 'color'}).kids.map(function (k) { return k.text; }),"
        " buildResultRow({start: 1, end: 3, duration: 2}, 0, null, false, {type: 'color'}).kids.map(function (k) { return k.text; })]",
        prelude=_RESULT_ROW_STUBS,
        snippet=_RESULT_ROW,
    )
    assert out == [["4"], ["1 – 3", "2.0s"]]


_SET_RANGE = slice_between(
    read("screenspace-utils.js"), "function setRangeValue(", "// Click the numeral"
)
_SET_RANGE_STUBS = """
function numberOrDefault(v, d) { var n = parseFloat(v); return isNaN(n) ? d : n; }
function slider(min, step) {
  return { min: String(min), step: String(step), value: "",
    getAttribute: function () { return String(step); } };
}
function typed(min, step, value) {
  var s = slider(min, step);
  setRangeValue(s, value);
  return s.step;
}
"""


def test_typed_slider_value_keeps_off_step_precision():
    """A range input snaps its value to `step`, so a typed 137 on a step-5
    slider read back as 135 unless the step is refined first."""
    out = _probe(
        "[typed(25, 5, 137), typed(25, 5, 140), typed(0.5, 0.01, 0.755),"
        " typed(0.5, 0.01, 0.7), typed(0.5, 0.5, 2.25), typed(0, 1, 30)]",
        prelude=_SET_RANGE_STUBS,
        snippet=_SET_RANGE,
    )
    assert out == ["1", "5", "0.001", "0.01", "0.05", "1"]


_ADJUST = slice_between(
    read("screenspace.js"), "function ssAdjustControl(", "// Only the tool region"
)
_ADJUST_STUBS = """
function numberOrDefault(v, d) { var n = parseFloat(v); return isNaN(n) ? d : n; }
function Event() {}
function nudge(min, declared, live, value, dir) {
  var c = { type: "range", min: String(min), max: "400", step: String(live),
    value: String(value), dispatchEvent: function () {},
    getAttribute: function () { return declared === null ? null : String(declared); } };
  ssAdjustControl(c, dir);
  return [c.value, c.step];
}
"""


def test_keyboard_nudge_leaves_a_typed_fine_step():
    """Typing 137 leaves step 1 on a step-5 slider; arrow-key navigation fires
    synthetic input events, so it kept nudging by the fine step."""
    out = _probe(
        "[nudge(25, 5, 1, 137, 1), nudge(25, 5, 1, 137, -1), nudge(25, 5, 5, 135, 1),"
        " nudge(0.5, 0.01, 0.001, 0.755, 1), nudge(0.5, 0.01, 0.001, 0.755, -1),"
        " nudge(0, null, 2, 10, 1)]",
        prelude=_ADJUST_STUBS,
        snippet=_ADJUST,
    )
    assert out == [
        [140, "5"],
        [135, "5"],
        [140, "5"],
        [0.76, "0.01"],
        [0.75, "0.01"],
        [12, "2"],
    ]
