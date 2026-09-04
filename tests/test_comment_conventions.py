"""AGENTS.md caps comment blocks at 15 words; this scan makes the cap mechanical.

A block is a run of same-indent comment lines, a ``/* */`` block, or one
trailing comment. Docstrings, the JS file-header block, and tool directives
(``noqa``, ``type:``, ``eslint``) are exempt. Splitting one block into adjacent
blocks does not help: a blank-line-free run counts as one.

``source/``, ``build/``, ``clipgen.py``, and ``assets/web/`` must be clean.
``tests/`` carries frozen per-file counts that ratchet both ways, the same
pattern as ``test_source_conventions.py``: growth fails, and a fixed file must
drop from the baseline so the win is locked in.

One pass: every file is read exactly once (test-perf budget).
"""

import io
import re
import tokenize
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LIMIT = 15

_PY_DIRECTIVE = re.compile(r"^(noqa|type:|ty:|pragma|ruff:|fmt:|isort:)")
_JS_DIRECTIVE = re.compile(r"^(eslint|global |jshint)")
_JS_REGEX_LEAD = set("(,=:[!&|?{};+-*%<>~^")
_JS_REGEX_WORDS = ("return", "typeof", "case", "in", "of", "void", "delete")


def _words(text: str) -> int:
    return len(text.split())


def python_blocks(text: str) -> list[tuple[int, str]]:
    """(line, body) per comment block; directives skipped, docstrings ignored."""
    blocks: list[tuple[int, str]] = []
    run: list[str] = []
    run_line = 0
    run_col = -1
    last_comment_line = -1
    try:
        tokens = list(tokenize.generate_tokens(io.StringIO(text).readline))
    except (tokenize.TokenError, SyntaxError):
        return blocks

    def flush() -> None:
        nonlocal run
        if run:
            blocks.append((run_line, " ".join(run)))
        run = []

    for tok in tokens:
        if tok.type == tokenize.COMMENT:
            body = tok.string.lstrip("#").strip()
            if _PY_DIRECTIVE.match(body):
                continue
            line, col = tok.start
            if run and line == last_comment_line + 1 and col == run_col:
                run.append(body)
            else:
                flush()
                run, run_line, run_col = [body], line, col
            last_comment_line = line
        elif tok.type in (
            tokenize.NL,
            tokenize.NEWLINE,
            tokenize.INDENT,
            tokenize.DEDENT,
        ):
            continue
        elif run and tok.start[0] > last_comment_line:
            flush()
    flush()
    return blocks


def _js_regex_start(line: str, i: int) -> bool:
    before = line[:i].rstrip()
    if not before:
        return True
    if before[-1] in _JS_REGEX_LEAD:
        return True
    return any(before.endswith(w) for w in _JS_REGEX_WORDS)


def _js_skip_literal(line: str, i: int) -> int:
    """Index just past the string or regex literal starting at ``i``."""
    quote = line[i]
    j = i + 1
    in_class = False
    while j < len(line):
        ch = line[j]
        if ch == "\\":
            j += 2
            continue
        if quote == "/":
            if ch == "[":
                in_class = True
            elif ch == "]":
                in_class = False
            elif ch == "/" and not in_class:
                return j + 1
        elif ch == quote:
            return j + 1
        j += 1
    return len(line)


def js_blocks(text: str) -> list[tuple[int, str]]:
    """(line, body) per comment block; the line-1 ``/* */`` header is skipped."""
    blocks: list[tuple[int, str]] = []
    lines = text.split("\n")
    run: list[str] = []
    run_line = 0
    run_indent = -1
    in_block = False
    block_line = 0
    block_buf: list[str] = []

    def flush_run() -> None:
        nonlocal run
        if run:
            blocks.append((run_line, " ".join(run)))
        run = []

    for idx, raw in enumerate(lines, start=1):
        line = raw
        i = 0
        code_seen = False
        if in_block:
            end = line.find("*/")
            if end < 0:
                block_buf.append(line)
                continue
            block_buf.append(line[:end])
            in_block = False
            if block_line != 1:
                body = " ".join(re.sub(r"^\s*\*+\s?", "", s).strip() for s in block_buf)
                blocks.append((block_line, body.strip()))
            i = end + 2
        while i < len(line):
            ch = line[i]
            if ch in "\"'`":
                code_seen = True
                i = _js_skip_literal(line, i)
                continue
            if ch == "/" and i + 1 < len(line):
                nxt = line[i + 1]
                if nxt == "/":
                    body = line[i + 2 :].strip()
                    indent = len(line) - len(line.lstrip())
                    if _JS_DIRECTIVE.match(body):
                        break
                    if code_seen:
                        flush_run()
                        blocks.append((idx, body))
                    elif run and idx == run_line + len(run) and indent == run_indent:
                        run.append(body)
                    else:
                        flush_run()
                        run, run_line, run_indent = [body], idx, indent
                    break
                if nxt == "*":
                    flush_run()
                    end = line.find("*/", i + 2)
                    if end < 0:
                        in_block = True
                        block_line = idx
                        block_buf = [line[i + 2 :]]
                        break
                    if idx != 1:
                        blocks.append((idx, line[i + 2 : end].strip(" *")))
                    i = end + 2
                    continue
                if _js_regex_start(line, i):
                    code_seen = True
                    i = _js_skip_literal(line, i)
                    continue
            if not ch.isspace():
                code_seen = True
            i += 1
        else:
            if code_seen or not line.strip():
                flush_run()
    flush_run()
    return blocks


def _scan(path: Path) -> list[tuple[int, int, str]]:
    text = path.read_text(encoding="utf-8")
    blocks = python_blocks(text) if path.suffix == ".py" else js_blocks(text)
    return [(ln, _words(b), b) for ln, b in blocks if _words(b) > LIMIT]


def _report(path: Path, hits: list[tuple[int, int, str]]) -> str:
    rel = path.relative_to(ROOT)
    return "\n".join(f"  {rel}:{ln}: {n} words: {b[:90]}" for ln, n, b in hits)


_STRICT_FILES = sorted(
    [ROOT / "clipgen.py"]
    + list((ROOT / "source").glob("*.py"))
    + list((ROOT / "build").glob("*.py"))
    + list((ROOT / "assets" / "web").glob("*.js"))
)
_TEST_FILES = sorted(
    p for p in (ROOT / "tests").rglob("*.py") if p.name != Path(__file__).name
)

_STRICT_HITS = {p: _scan(p) for p in _STRICT_FILES}
_TEST_HITS = {p.relative_to(ROOT).as_posix(): _scan(p) for p in _TEST_FILES}

# Over-limit blocks per tests/ file as of the 2026-09 pass. Ratchets down only.
_TEST_BASELINE: dict[str, int] = {
    "tests/conftest.py": 1,
    "tests/screenspace/test_attention.py": 11,
    "tests/screenspace/test_change_similarity.py": 4,
    "tests/screenspace/test_inactivity.py": 1,
    "tests/screenspace/test_manifest_events.py": 1,
    "tests/screenspace/test_multitool_scoring.py": 24,
    "tests/screenspace/test_region_mask.py": 4,
    "tests/screenspace/test_scan_pipeline.py": 6,
    "tests/screenspace/test_shape.py": 4,
    "tests/screenspace/test_static_gating.py": 3,
    "tests/screenspace/test_text_numbers.py": 14,
    "tests/screenspace/test_worker.py": 7,
    "tests/test_boot_dispatcher.py": 1,
    "tests/test_cli_args.py": 3,
    "tests/test_cli_event_clip_args.py": 2,
    "tests/test_cli_screenspace_args.py": 2,
    "tests/test_clip_pipeline.py": 8,
    "tests/test_composer_server.py": 3,
    "tests/test_container_seekability.py": 5,
    "tests/test_css_custom_properties.py": 1,
    "tests/test_css_token_discipline.py": 1,
    "tests/test_data_export.py": 2,
    "tests/test_desktop_chrome.py": 4,
    "tests/test_desktop_window.py": 1,
    "tests/test_export_copy.py": 1,
    "tests/test_files_and_artifacts.py": 3,
    "tests/test_friction_agent.py": 4,
    "tests/test_friction_scorer.py": 2,
    "tests/test_friction_smoke.py": 1,
    "tests/test_frontend_satellite_wiring.py": 3,
    "tests/test_google_and_excel_adapters.py": 1,
    "tests/test_head_partial.py": 1,
    "tests/test_hotkeys_frontend_source.py": 5,
    "tests/test_icon_conventions.py": 1,
    "tests/test_js_dead_functions.py": 1,
    "tests/test_licenses.py": 1,
    "tests/test_llm_client.py": 2,
    "tests/test_media_play_frontend_source.py": 2,
    "tests/test_mindnode.py": 5,
    "tests/test_motion_wiring.py": 27,
    "tests/test_multi_video.py": 7,
    "tests/test_overview.py": 1,
    "tests/test_packaging.py": 4,
    "tests/test_participant_merge.py": 5,
    "tests/test_pipeline_io.py": 2,
    "tests/test_profiling.py": 7,
    "tests/test_resource_lifecycle.py": 1,
    "tests/test_screenspace_api.py": 15,
    "tests/test_screenspace_preview.py": 5,
    "tests/test_scrollbar_theming.py": 1,
    "tests/test_select_styling.py": 1,
    "tests/test_selectors.py": 2,
    "tests/test_shared_constants.py": 4,
    "tests/test_shot.py": 1,
    "tests/test_spreadsheet_generation.py": 4,
    "tests/test_start_endpoints.py": 6,
    "tests/test_start_overlay_source.py": 5,
    "tests/test_start_settings.py": 1,
    "tests/test_studio_api.py": 18,
    "tests/test_studio_frontend_source.py": 5,
    "tests/test_thinking_agents.py": 5,
    "tests/test_titlecards.py": 2,
    "tests/test_tooltip_conventions.py": 2,
    "tests/test_transcripts.py": 5,
    "tests/test_transcripts_api.py": 31,
    "tests/test_transcripts_dom_wiring.py": 6,
    "tests/test_utils_timestamps.py": 6,
    "tests/test_video_commands.py": 10,
    "tests/test_viewer_data.py": 2,
    "tests/test_viewer_inline.py": 5,
    "tests/test_workflows_api.py": 40,
    "tests/test_workflows_collection_ops.py": 1,
    "tests/test_workflows_executors.py": 15,
    "tests/test_workflows_frontend_source.py": 19,
    "tests/test_workflows_runner.py": 25,
    "tests/ui/_ui_browser.py": 1,
    "tests/ui/_ui_fixtures.py": 9,
    "tests/ui/_ui_pages.py": 3,
    "tests/ui/_ui_server.py": 2,
    "tests/ui/_ui_session.py": 1,
    "tests/ui/_ui_states.py": 2,
    "tests/ui/conftest.py": 2,
    "tests/ui/shot.py": 6,
    "tests/ui/test_ui_journeys.py": 3,
    "tests/ui/test_ui_smoke.py": 2,
}


def test_comment_blocks_within_limit() -> None:
    """Every source, build, and web comment block is <= LIMIT words."""
    over = {p: h for p, h in _STRICT_HITS.items() if h}
    assert not over, (
        f"comment blocks over {LIMIT} words (AGENTS.md 'Be concise'); "
        "rewrite from scratch, never trim word-by-word:\n"
        + "\n".join(_report(p, h) for p, h in sorted(over.items()))
    )


def test_test_comment_blocks_ratchet() -> None:
    """tests/ over-limit counts never grow; a fixed file leaves the baseline."""
    found = {name: len(h) for name, h in _TEST_HITS.items() if h}
    grew = {n: c for n, c in found.items() if c > _TEST_BASELINE.get(n, 0)}
    assert not grew, (
        f"new comment blocks over {LIMIT} words in tests: {grew}\n"
        + "\n".join(_report(ROOT / n, _TEST_HITS[n]) for n in sorted(grew))
    )
    shrank = {n: c for n, c in _TEST_BASELINE.items() if found.get(n, 0) < c}
    assert not shrank, f"nice — ratchet _TEST_BASELINE down for {shrank}"


def test_test_baseline_has_no_dead_entries() -> None:
    dead = sorted(set(_TEST_BASELINE) - set(_TEST_HITS))
    assert not dead, f"_TEST_BASELINE names files that no longer exist: {dead}"
