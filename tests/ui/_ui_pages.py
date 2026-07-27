"""Page table, error capture, and the readiness sequence — shared by the smoke and shot.py.

Readiness deliberately avoids ``networkidle``: several pages hold an open
``EventSource``, so the network never goes quiet, and "quiet" would not mean
"painted" anyway. Three explicit steps instead —

1. ``domcontentloaded``,
2. ``nav.topnav``, which ``topnav.js`` mounts on every one of the six pages, so
   it separates "the bundle threw before boot" from "a fetch never resolved",
3. a per-page selector that only appears once that page's own data has rendered.

Those per-page selectors are reverse-engineered from each page's render function
and are the part of this harness most likely to rot. A rename turns a pass into a
confusing timeout, so the failure message says which stage timed out.

Each selector is deliberately chosen to be *absent in the zero state*. That
matters more than it sounds: the first draft waited on
``#coParticipantSelect option[value="P01"]``, which exists as soon as the dropdown
is populated — so Composer passed while photographing an empty "Select a
participant to load their source video" screen. A green run that photographs
nothing is worse than a red one. Pages that need a participant loaded get one via
the app's own ``/composer/#P01`` deep-link (``clipgenHashParticipant`` in
utils.js), not by synthesizing clicks.
"""

from dataclasses import dataclass, field
from typing import Any

# page segment -> (url fragment, a selector that only exists once data rendered)
PAGES: dict[str, tuple[str, str]] = {
    "studio": ("", "#sheetGrid tbody tr"),
    "screenspace": ("#P01", "#videoPlayer[src]"),
    "transcripts": ("#P01", "#segmentList .segment-row"),
    "workflows": ("", "#wfBlueprintSelect option"),
    "composer": ("#P01", "#coCutList .co-cut-item"),
    "overview": ("", "#ovStudyName:not(:empty)"),
}

# console.error text we refuse to treat as a defect.
_CONSOLE_ALLOWLIST = (
    # HTTP status is owned by the response handler below; keeping this would
    # double-report every non-2xx.
    "Failed to load resource",
    # Headless has no GPU. Overview's three.js falls back to SwiftShader and says so.
    "Automatic fallback to software WebGL",
    "GroupMarkerNotSet",
    "Failed to create WebGL context",
    "SharedImageManager",
    "Fallback to SwiftShader",
)

# A non-2xx on these resource types is always a bug. XHR/image/media are not:
# /api/thumbnail, /api/sprite and /api/preview legitimately 404 when no frames
# have been extracted, which is the normal state of a fresh fixture project.
_FATAL_RESOURCE_TYPES = frozenset({"document", "script", "stylesheet"})

_SETTLE_MS = 700


@dataclass
class PageLog:
    page_errors: list[str] = field(default_factory=list)
    console_errors: list[str] = field(default_factory=list)
    non_2xx: list[tuple[str, int, str]] = field(default_factory=list)
    request_failures: list[tuple[str, str]] = field(default_factory=list)
    # Set when a readiness wait timed out. Recorded rather than raised, because a
    # boot-time ReferenceError usually shows up *as* a timeout — the page simply
    # never renders — and the pageerror naming the symbol is the far more useful
    # message of the two. See format_failure.
    timeout: str = ""

    @property
    def fatal(self) -> bool:
        return bool(
            self.page_errors
            or self.console_errors
            or self.timeout
            or [item for item in self.non_2xx if item[2] in _FATAL_RESOURCE_TYPES]
        )


def wire_listeners(page: Any, log: PageLog) -> None:
    """Attach the four capture handlers before any navigation happens."""

    def on_console(message: Any) -> None:
        if message.type != "error":
            return
        text = message.text
        if any(allowed in text for allowed in _CONSOLE_ALLOWLIST):
            return
        log.console_errors.append(text)

    def on_response(response: Any) -> None:
        if response.status < 400:
            return
        log.non_2xx.append(
            (response.url, response.status, response.request.resource_type)
        )

    def on_request_failed(request: Any) -> None:
        failure = request.failure or ""
        # SSE streams and in-flight fetches are aborted by design on close.
        if "ERR_ABORTED" in failure:
            return
        log.request_failures.append((request.url, failure))

    page.on("pageerror", lambda exc: log.page_errors.append(str(exc)))
    page.on("console", on_console)
    page.on("response", on_response)
    page.on("requestfailed", on_request_failed)


def open_and_settle(
    page: Any, base_url: str, name: str, log: PageLog, extra_wait_ms: int = 0
) -> None:
    """Navigate to a page and wait until its own data has rendered.

    A readiness timeout lands in ``log.timeout`` instead of raising, so the
    caller can still screenshot the page and — crucially — lead with any
    ``pageerror`` it captured. The timeout is the symptom; the exception thrown
    during boot is the cause.
    """
    import _ui_browser

    fragment, ready = PAGES[name]
    page.goto(
        f"{base_url}/{name}/{fragment}", wait_until="domcontentloaded", timeout=20_000
    )
    try:
        page.wait_for_selector("nav.topnav", state="attached", timeout=15_000)
        page.wait_for_selector(ready, state="attached", timeout=25_000)
    except _ui_browser.playwright_error() as exc:
        log.timeout = str(exc).splitlines()[0]
    page.wait_for_timeout(_SETTLE_MS + extra_wait_ms)


def format_failure(name: str, log: PageLog, screenshot: str, report: str) -> str:
    lines = [
        (
            f"{name}: {len(log.page_errors)} page error(s), "
            f"{len(log.console_errors)} console error(s)"
        )
    ]
    # Page errors first: when a boot-time throw stops the page rendering, the
    # timeout below is only the downstream symptom.
    lines += [f"  [pageerror] {text}" for text in log.page_errors]
    lines += [f"  [console]   {text}" for text in log.console_errors]
    lines += [
        f"  [{kind} {status}] {url}"
        for url, status, kind in log.non_2xx
        if kind in _FATAL_RESOURCE_TYPES
    ]
    if log.timeout:
        lines.append(f"  [timeout]   {log.timeout}")
    lines.append(f"  screenshot: {screenshot}")
    lines.append(f"  report:     {report}")
    return "\n".join(lines)
