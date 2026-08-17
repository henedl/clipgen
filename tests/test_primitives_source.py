"""Source locks for shared UI primitives.

The Overview swim lane used to bind mouseenter/leave/click on every marker.
On a 200×12 sheet that was 7,592 listeners (three per event). Delegation on
the wrap is the fix; this test keeps the per-marker path from coming back.
"""

from _frontend_source import read


def test_swim_lane_does_not_bind_per_marker_listeners() -> None:
    src = read("primitives.js")
    start = src.index("function renderEvents()")
    end = src.index("function bindDelegates()", start)
    body = src[start:end]
    assert "addEventListener" not in body
    assert "document.createDocumentFragment" in body
    assert "function bindDelegates()" in src
