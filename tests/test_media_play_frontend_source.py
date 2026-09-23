"""Source-level guard: every ``play()`` on a media element handles its promise.

``HTMLMediaElement.play()`` returns a promise that *rejects* whenever playback
is interrupted before it starts — a ``pause()`` in the same tick, a ``load()``,
a source-swapping seek, or autoplay policy. All of those are ordinary here
(clicking a transcript segment and immediately pausing is one gesture), but an
unhandled rejection surfaces as an uncaught page error, which is console noise
for users and a hard failure in ``/ui-check``.

Six of the twelve call sites had drifted into calling it bare, so the contract
is pinned here rather than left to review: go through
``ClipgenVideoControls.safePlay`` (video-controls.js), or handle the promise
inline on the pages that cannot load that module.
"""

import re

from _frontend_source import WEB, read

# Files that may call play() bare on their own line, and why. Both attach a
# handler within two lines, which the assertion below verifies rather than
# trusts — an allowlist entry buys a *local* handler, not an exemption.
#
# gallery.js and viewer.js are inlined into exported standalone viewers, which
# ship without video-controls.js, so they cannot reach safePlay().
INLINE_HANDLER_ALLOWLIST = {
    "gallery.js",
    "viewer.js",
    "video-controls.js",  # defines safePlay(); its own mix path needs both arms
}

# `.play()` on anything, ignoring the definition inside safePlay's own comment.
_PLAY_RE = re.compile(r"\.play\(\)")


def _js_files():
    return sorted(p for p in WEB.glob("*.js"))


def _handled_within(src: str, idx: int) -> bool:
    """Is the play() at *idx* followed by promise handling within two lines?"""
    tail = src[idx:].split("\n")[:3]
    return any(".then(" in line or ".catch(" in line for line in tail)


def test_every_play_call_handles_its_promise():
    offenders = []
    for path in _js_files():
        src = path.read_text(encoding="utf-8")
        for match in _PLAY_RE.finditer(src):
            line_start = src.rfind("\n", 0, match.start()) + 1
            line = src[line_start : src.find("\n", match.start())]
            if line.lstrip().startswith(("*", "//")):
                continue  # prose in a comment, not a call
            if path.name in INLINE_HANDLER_ALLOWLIST:
                if not _handled_within(src, line_start):
                    offenders.append(
                        f"{path.name}: allowlisted but unhandled -> {line.strip()}"
                    )
                continue
            offenders.append(f"{path.name}: {line.strip()}")
    assert not offenders, (
        "play() promises must be handled — call "
        "window.ClipgenVideoControls.safePlay(el[, onRejected]) instead:\n  "
        + "\n  ".join(offenders)
    )


def test_safeplay_is_exported_and_defensive():
    src = read("video-controls.js")
    assert "safePlay: safePlay," in src, "safePlay must be on the public API"
    body = src[src.index("function safePlay(") : src.index("function nextSpeed(")]
    # A missing element must be a no-op, and a browser returning no promise
    # (or a rejection with no caller handler) must not throw either.
    assert "if (!videoEl) return;" in body
    assert "p && p.catch" in body


def test_pages_that_use_safeplay_load_video_controls():
    """safePlay is only reachable on pages that actually load the module."""
    users = {
        path.name.split("-")[0].split(".")[0]
        for path in _js_files()
        if "ClipgenVideoControls.safePlay" in path.read_text(encoding="utf-8")
    }
    assert users, "expected the shared helper to have callers"
    for page in users:
        markup = read(f"{page}.html")
        assert "video-controls.js" in markup, (
            f"{page}.js calls safePlay but {page}.html does not load "
            "video-controls.js — it would throw at runtime"
        )
