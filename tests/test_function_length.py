"""No new function may reach 150 lines, and the existing giants may not grow.

Covers Python (``source/``, ``build/``, ``clipgen.py``; measured by the AST as
``end_lineno - lineno + 1``) and every ``function name(`` in
``assets/web/*.js`` (brace-matched, skipping strings, comments and regex
literals). Nested functions count on their own as well as inside their parent.

``KNOWN_LONG`` pins every function at or over ``LIMIT`` to its length when the
ratchet landed. A listed function may grow by at most ``GROWTH`` over its pin,
so edits to a giant cannot keep inflating it; one that drops below ``LIMIT``
must leave the list. The fix for a failure is a split in the style of Wave 5
in ``plans/CODE-QUALITY-PLAN.md``: carve named phases, not arbitrary halves.
"""

import re
from functools import cache
from pathlib import Path

from _frontend_source import WEB
from _py_source import PRODUCTION, ROOT, functions

LIMIT = 150
GROWTH = 1.10

_JS_FUNCTION = re.compile(r"^[ \t]*function\s+([A-Za-z_$][\w$]*)\s*\(", re.MULTILINE)
_JS_STOP = re.compile(r"[\"'`/{}]")
_REGEX_LEAD = set("(,=:[!&|?{};+-*%<>~^")
_REGEX_WORD = re.compile(r"\b(?:return|typeof|case|in|of|void|delete)$")

# Lengths as of 2026-09-25. Shrink or delete entries; never add one.
KNOWN_LONG: dict[str, int] = {
    "assets/web/composer-annotate.js::initGestures": 251,
    "assets/web/composer-annotate.js::initPalette": 174,
    "assets/web/composer-timeline.js::initTimeline": 301,
    "assets/web/composer-timeline.js::renderTimelineImpl": 157,
    "assets/web/primitives.js::createSwimLane": 325,
    "assets/web/screenspace-model-view.js::_doRefreshModelView": 224,
    "assets/web/screenspace-multitool-params.js::renderMultitoolParams": 326,
    "assets/web/screenspace-overlay-interaction.js::initRegionDrawing": 526,
    "assets/web/screenspace-overlay.js::renderOverlay": 308,
    "assets/web/screenspace-params.js::renderColorParams": 158,
    "assets/web/screenspace-results.js::initResultsPanel": 187,
    "assets/web/screenspace-results.js::renderResultsImpl": 320,
    "assets/web/screenspace-run.js::gatherWorkflowParams": 171,
    "assets/web/screenspace-sample-editor.js::openSampleModal": 207,
    "assets/web/screenspace-tasks.js::initTaskQueue": 220,
    "assets/web/screenspace-tasks.js::restoreTaskToWorkflow": 234,
    "assets/web/screenspace-timeline.js::initTimeline": 162,
    "assets/web/screenspace-timeline.js::renderTimelineImpl": 212,
    "assets/web/screenspace.js::initKeyboard": 176,
    "assets/web/settings-modal.js::_buildLlmModelsBlock": 208,
    "assets/web/start-overlay.js::bind": 161,
    "assets/web/studio-generate.js::onGenerate": 204,
    "assets/web/studio-trim.js::openTrimPopover": 181,
    "assets/web/studio.js::bindButtons": 175,
    "assets/web/transcripts-batch.js::createBatchJobModal": 166,
    "assets/web/transcripts.js::_confirmModelInstallNow": 153,
    "assets/web/video-controls.js::attachAudioPanel": 336,
    "source/cli.py::main": 211,
    "source/cli_args.py::_add_screenspace_args": 215,
    "source/composer_server.py::_run_overlay_export": 155,
    "source/interactive.py::browse_spreadsheet": 331,
    "source/pipeline.py::_process_reel": 271,
    "source/pipeline.py::_process_single_clip_segments": 199,
    "source/pipeline.py::process_clips": 227,
    "source/screenspace_multitool.py::scan_multitool": 210,
    "source/screenspace_scans.py::scan_attention": 173,
    "source/server.py::_process_intake_item": 154,
    "source/server.py::api_generate": 225,
    "source/server.py::api_generate.stream": 153,
    "source/server.py::api_reel_direct": 222,
    "source/server.py::api_reel_direct.stream": 186,
    "source/server.py::api_reel_direct.stream.work": 169,
    "source/server.py::serve_combined_app": 180,
    "source/titlecards.py::_build_card_frame": 154,
    "source/titlecards.py::wrap_clip_with_cards": 217,
    "source/transcripts.py::TranscriptWorker._execute_task": 164,
    "source/transcripts.py::transcribe_video": 222,
    "source/video.py::compress_to_size": 158,
    "source/workflows.py::_exec_post_process": 181,
}


def _js_body_end(src: str, start: int) -> int:
    """Index of the ``}`` closing the first ``{`` at or after ``start``; -1 if none.

    A ``/`` opens a regex after an operator, an opening bracket, or a keyword,
    and division otherwise. Regex literals never span lines, which bounds a
    wrong guess to one line.
    """
    depth = 0
    i = start
    n = len(src)
    while True:
        match = _JS_STOP.search(src, i)
        if not match:
            return -1
        i = match.start()
        ch = src[i]
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return i
        elif ch in "\"'`":
            j = i + 1
            while j < n and src[j] != ch:
                j += 2 if src[j] == "\\" else 1
            i = j
        else:
            nxt = src[i + 1 : i + 2]
            if nxt == "/":
                j = src.find("\n", i)
                i = n if j < 0 else j
                continue
            if nxt == "*":
                j = src.find("*/", i + 2)
                i = n if j < 0 else j + 2
                continue
            before = src[max(0, i - 12) : i].rstrip()
            if not before or before[-1] in _REGEX_LEAD or _REGEX_WORD.search(before):
                j = i + 1
                in_class = False
                while j < n and src[j] != "\n":
                    c = src[j]
                    if c == "\\":
                        j += 2
                        continue
                    if c == "[":
                        in_class = True
                    elif c == "]":
                        in_class = False
                    elif c == "/" and not in_class:
                        break
                    j += 1
                i = j
        i += 1


def js_lengths(src: str) -> dict[str, int]:
    """Line count per ``function name(``; a repeated name keeps its longest."""
    lengths: dict[str, int] = {}
    for match in _JS_FUNCTION.finditer(src):
        end = _js_body_end(src, match.end())
        assert end >= 0, f"unclosed function {match.group(1)}"
        length = src.count("\n", match.start(), end) + 1
        name = match.group(1)
        lengths[name] = max(length, lengths.get(name, 0))
    return lengths


@cache
def _measure() -> dict[str, int]:
    """``path::name`` to length for every function in the tree."""
    found: dict[str, int] = {}
    for path in PRODUCTION:
        rel = path.relative_to(ROOT).as_posix()
        for fn in functions(path):
            key = f"{rel}::{fn.qualname}"
            found[key] = max(fn.length, found.get(key, 0))
    for path in sorted(WEB.glob("*.js")):
        rel = Path(path).relative_to(ROOT).as_posix()
        for name, length in js_lengths(path.read_text(encoding="utf-8")).items():
            found[f"{rel}::{name}"] = length
    return found


_HOW = (
    "Split it into named phases (see Wave 5 in plans/CODE-QUALITY-PLAN.md) "
    "rather than raising the pin."
)


def test_no_new_long_functions():
    new = {
        key: length
        for key, length in _measure().items()
        if length >= LIMIT and key not in KNOWN_LONG
    }
    assert not new, f"These functions are {LIMIT}+ lines. {_HOW}\n" + "\n".join(
        f"  {key}: {length}" for key, length in sorted(new.items())
    )


def test_long_functions_do_not_grow():
    found = _measure()
    grown = {
        key: (pin, found[key])
        for key, pin in KNOWN_LONG.items()
        if key in found and found[key] > pin * GROWTH
    }
    assert not grown, (
        f"These functions grew over {GROWTH:.0%} of their KNOWN_LONG pin. {_HOW}\n"
        + "\n".join(
            f"  {key}: {pin} -> {now}" for key, (pin, now) in sorted(grown.items())
        )
    )


def test_known_long_has_no_stale_entries():
    """A split, renamed or deleted function must leave the list."""
    found = _measure()
    stale = sorted(k for k in KNOWN_LONG if found.get(k, 0) < LIMIT)
    assert not stale, f"Drop these KNOWN_LONG entries; they are gone or short: {stale}"


def test_js_measure_skips_literals():
    """Braces inside strings, comments and regex literals must not count."""
    src = (
        "function a(x) {\n"
        "  var s = \"}\" + '{' + x / 2;\n"
        "  // }\n"
        "  /* } */\n"
        "  var r = /[}]\\//g;\n"
        "  return r.test(s) ? { k: margin / 2 } : 0;\n"
        "}\n"
        "function b() { return 1; }\n"
    )
    assert js_lengths(src) == {"a": 7, "b": 1}
