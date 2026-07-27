"""Open one page, photograph it, and optionally run JavaScript inside it.

The six-page smoke costs ~30 s cold, which is the wrong loop for iterating on a
single surface. This does one page in a few seconds, and adds the thing the old
"paste this DevTools snippet to the human" workflow was standing in for:
``--eval`` runs arbitrary JS in the live page and prints what it returns. Read
computed styles, count rendered nodes, dump ``state``, call
``el.getAnimations()`` — whatever you would have asked someone else to type.

    uv run --extra ui python tests/ui/shot.py studio
    uv run --extra ui python tests/ui/shot.py studio --selector "#sheetGrid"
    uv run --extra ui python tests/ui/shot.py transcripts \
        --eval "return document.querySelectorAll('.pill-wrap').length"
    uv run --extra ui python tests/ui/shot.py overview --eval-file /tmp/probe.js --wait 2000

Not a test: ``norecursedirs = ui`` plus the non-``test_`` filename keep pytest
away from it. The JS it runs only ever reaches a loopback server serving
generated fixture data.
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

import _ui_browser
import _ui_fixtures
import _ui_pages
import _ui_server
import config
import start_settings
import utils


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

    try:
        _ui_fixtures.ensure_inputs()
        chromium_path = _ui_browser.resolve_chromium()
        playwright_factory = _ui_browser.sync_playwright()
    except _ui_fixtures.UiUnavailable as exc:
        print(exc, file=sys.stderr)
        return 1
    _ui_fixtures.ensure_run_dirs()

    # Same config redirection the pytest fixtures do, minus the monkeypatch: this
    # is a one-shot process, so plain assignment is enough.
    config.INPUT_DIR = str(_ui_fixtures.INPUT_DIR)
    config.OUTPUT_DIR = str(_ui_fixtures.OUTPUT_DIR)
    # Keep stdout to the screenshot path and the --eval result; sheet-loading
    # chatter would bury both.
    config.VERBOSITY = config.QUIET
    # Redirect the Start-overlay settings file, or build_combined_app's
    # record_project_session would prepend these fixture dirs to the real
    # "Recently opened" rail. setattr rather than plain assignment because the
    # attribute's declared type is the original function; this is the same
    # rebinding monkeypatch does in tests/ui/conftest.py.
    setattr(  # noqa: B010
        start_settings,
        "_settings_path",
        lambda: _ui_fixtures.SETTINGS_DIR / "start.json",
    )
    utils.NO_INPUT_MODE = True

    workbook, reason = _ui_fixtures.open_workbook()
    if reason:
        print(reason, file=sys.stderr)
        return 1

    log = _ui_pages.PageLog()
    result: Any = None
    eval_error: str | None = None
    live = _ui_server.start(workbook)
    playwright = playwright_factory().start()
    try:
        browser = playwright.chromium.launch(
            executable_path=str(chromium_path),
            headless=True,
            args=_ui_browser.LAUNCH_ARGS,
        )
        context = browser.new_context(viewport=viewport, device_scale_factor=1)
        context.add_init_script(
            "try { sessionStorage.setItem('clipgen.startOverlayDismissed', '1'); }"
            " catch (e) {}"
        )
        page = context.new_page()
        _ui_pages.wire_listeners(page, log)
        try:
            _ui_pages.open_and_settle(page, live.url, args.page, log, args.wait)
            if body:
                try:
                    result = page.evaluate(f"() => {{ {body} }}")
                except _ui_browser.playwright_error() as exc:
                    # A throwing snippet is a normal outcome when probing, not a
                    # harness crash. Report the JS error, keep the screenshot.
                    eval_error = str(exc).splitlines()[0]
        finally:
            _capture(page, args, out)
        context.close()
        browser.close()
    finally:
        playwright.stop()
        _ui_server.stop(live)

    print(f"screenshot: {out.resolve()}")
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
    return 1 if (log.fatal or eval_error is not None) else 0


if __name__ == "__main__":
    raise SystemExit(main())
