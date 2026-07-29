"""Open one page, photograph it, and optionally run JavaScript inside it.

The six-page smoke costs ~30 s cold, which is the wrong loop for iterating on a
single surface. This does one page in a few seconds, and adds the thing the old
"paste this DevTools snippet to the human" workflow was standing in for:
``--eval`` runs arbitrary JS in the live page and prints what it returns. Read
computed styles, count rendered nodes, dump ``state``, call
``el.getAnimations()`` — whatever you would have asked someone else to type.

    uv run --extra ui python tests/ui/shot.py studio
    uv run --extra ui python tests/ui/shot.py studio --selector "#sheetGrid"
    uv run --extra ui python tests/ui/shot.py studio --theme light
    uv run --extra ui python tests/ui/shot.py studio --state settings
    uv run --extra ui python tests/ui/shot.py screenspace --all-states
    uv run --extra ui python tests/ui/shot.py transcripts \
        --eval "return document.querySelectorAll('.pill-wrap').length"
    uv run --extra ui python tests/ui/shot.py overview --eval-file /tmp/probe.js --wait 2000

``--theme light`` is worth reaching for: light is a shipped feature reachable from
every page's theme toggle, and until it was added here nothing in the harness had
ever rendered it. ``--state`` and ``--all-states`` (see ``_ui_states``) reach the
modals and tabs the six-page smoke never opens — ``--all-states`` drives them all
from a single boot, which is the one thing looping this script cannot do.

Not a test: ``norecursedirs = ui`` plus the non-``test_`` filename keep pytest
away from it. The JS it runs only ever reaches a loopback server serving
generated fixture data.
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))  # the _ui_* helpers
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "source"))

import _ui_browser
import _ui_fixtures
import _ui_pages
import _ui_session
import _ui_states


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="shot.py",
        description="Screenshot and probe one clipgen page in a headless browser.",
    )
    parser.add_argument("page", choices=sorted(_ui_pages.PAGES))
    parser.add_argument(
        "--selector",
        help="Clip the screenshot to this element instead of the full page.",
    )
    parser.add_argument(
        "--eval",
        dest="evaluate",
        help="JavaScript body to run in the page; its return value is printed as JSON.",
    )
    parser.add_argument(
        "--eval-file",
        type=Path,
        help="Read the --eval body from a file (for anything longer than a shell line).",
    )
    parser.add_argument(
        "--wait", type=int, default=0, help="Extra settle milliseconds before capture."
    )
    parser.add_argument("--out", type=Path, help="Override the screenshot path.")
    parser.add_argument(
        "--viewport", default="1600x1000", help="Viewport size, e.g. 1280x800."
    )
    parser.add_argument(
        "--theme",
        choices=_ui_session.THEMES,
        default="dark",
        help="Boot the page in this theme. Light is a shipped feature the "
        "harness otherwise never renders.",
    )
    parser.add_argument(
        "--state",
        help="Drive into one UI state before capturing, e.g. settings, "
        "cheatsheet, palette, start, tab:map, tool:color. Pass an unknown name "
        "to be told which states this page has.",
    )
    parser.add_argument(
        "--all-states",
        action="store_true",
        help="Capture every reachable state on this page from one boot, to "
        "<page>-<state>.png. Unreachable states are reported, not skipped.",
    )
    return parser.parse_args(argv)


def _viewport(raw: str) -> dict[str, int]:
    try:
        width, height = (int(part) for part in raw.lower().split("x", 1))
    except ValueError:
        raise SystemExit(f"--viewport expects WIDTHxHEIGHT, got {raw!r}") from None
    return {"width": width, "height": height}


def _capture(page: Any, args: argparse.Namespace, out: Path) -> None:
    if args.selector:
        page.locator(args.selector).first.screenshot(path=str(out))
    else:
        page.screenshot(path=str(out), full_page=True)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    viewport = _viewport(args.viewport)
    body = args.evaluate
    if args.eval_file is not None:
        body = args.eval_file.read_text(encoding="utf-8")

    out = args.out or (_ui_fixtures.SHOT_DIR / f"{args.page}.png")
    out.parent.mkdir(parents=True, exist_ok=True)

    log = _ui_pages.PageLog()
    result: Any = None
    eval_error: str | None = None
    shots: list[Path] = []
    states: list[_ui_states.StateResult] = []
    state_problem: str = ""
    try:
        with _ui_session.ui_session(viewport=viewport, theme=args.theme) as session:
            page = session.context.new_page()
            _ui_pages.wire_listeners(page, log)
            try:
                _ui_pages.open_and_settle(
                    page, session.origin, args.page, log, args.wait
                )
                if args.all_states:
                    # Capture the boot state too: --all-states should mean all.
                    _capture(page, args, out)
                    shots.append(out)
                    for state in _ui_states.each_state(page, args.page):
                        states.append(state)
                        if not state.reached:
                            continue
                        shot = out.with_name(
                            f"{out.stem}-{state.name.replace(':', '-')}{out.suffix}"
                        )
                        _capture(page, args, shot)
                        shots.append(shot)
                elif args.state:
                    state = _ui_states.enter_named(page, args.page, args.state)
                    states.append(state)
                    if not state.reached:
                        state_problem = state.detail
                if body:
                    try:
                        result = page.evaluate(f"() => {{ {body} }}")
                    except _ui_browser.playwright_error() as exc:
                        # A throwing snippet is a normal outcome when probing, not
                        # a harness crash. Report the JS error, keep the screenshot.
                        eval_error = str(exc).splitlines()[0]
            finally:
                if not args.all_states:
                    _capture(page, args, out)
                    shots.append(out)
    except _ui_fixtures.UiUnavailable as exc:
        print(exc, file=sys.stderr)
        return 1

    for shot in shots:
        print(f"screenshot: {shot.resolve()}")
    # Report every state we tried, reached or not. A silently skipped state reads
    # as covered, which is the failure mode this whole harness exists to avoid.
    for state in states:
        mark = "ok  " if state.reached else "MISS"
        suffix = f" — {state.detail}" if state.detail else ""
        print(f"state: {mark} {state.name}{suffix}")
    if body and eval_error is None:
        print("eval: " + json.dumps(result, ensure_ascii=False, default=str))
    if eval_error is not None:
        print(f"eval failed: {eval_error}", file=sys.stderr)
    for text in log.page_errors:
        print(f"[pageerror] {text}", file=sys.stderr)
    for text in log.console_errors:
        print(f"[console]   {text}", file=sys.stderr)
    if log.timeout:
        print(f"[timeout]   {log.timeout}", file=sys.stderr)
    if state_problem:
        print(f"state failed: {state_problem}", file=sys.stderr)
    # An explicitly requested --state that could not be reached is a failure; an
    # unreachable state during --all-states is reported above and is not, since
    # some states legitimately do not exist in this fixture.
    return 1 if (log.fatal or eval_error is not None or state_problem) else 0


if __name__ == "__main__":
    raise SystemExit(main())
