"""Regression checks for desktop window geometry persistence.

The window itself cannot exist under pytest, so these cover the decision layer:
``_window_kwargs`` (a saved rect in, create_window kwargs out) and the Shift
reset gate. The rules being pinned are the ones whose failure modes are
user-hostile — an app that opens off-screen with no way to drag it back, and a
window that cannot be closed.
"""

import desktop
import pytest


class FakeScreen:
    """Stands in for a pywebview Screen, which only needs four ints here."""

    def __init__(self, x, y, width, height):
        self.x = x
        self.y = y
        self.width = width
        self.height = height


LAPTOP = FakeScreen(0, 0, 1920, 1080)
SIDECAR = FakeScreen(1920, 0, 2560, 1440)


@pytest.fixture
def stored(monkeypatch):
    """A writable stand-in for the persisted rect, with a cleared() flag."""
    box = {"rect": None, "cleared": False}

    def clear():
        box["rect"] = None
        box["cleared"] = True

    monkeypatch.setattr(
        desktop.start_settings, "load_window_geometry", lambda: box["rect"]
    )
    monkeypatch.setattr(desktop.start_settings, "clear_window_geometry", clear)
    return box


def test_defaults_when_nothing_is_stored():
    assert desktop._window_kwargs(None, [LAPTOP]) == {"width": 1440, "height": 900}


def test_saved_rect_on_screen_is_honoured():
    saved = {"x": 200, "y": 100, "width": 1600, "height": 1000}
    assert desktop._window_kwargs(saved, [LAPTOP]) == saved


def test_size_is_clamped_to_the_window_minimum():
    saved = {"x": 0, "y": 0, "width": 320, "height": 200}
    kwargs = desktop._window_kwargs(saved, [LAPTOP])
    assert (kwargs["width"], kwargs["height"]) == desktop._WINDOW_MIN_SIZE


def test_size_is_clamped_to_the_largest_screen():
    saved = {"x": 0, "y": 0, "width": 9000, "height": 9000}
    kwargs = desktop._window_kwargs(saved, [LAPTOP, SIDECAR])
    assert (kwargs["width"], kwargs["height"]) == (2560, 1440)


def test_rect_on_a_secondary_screen_keeps_its_position():
    saved = {"x": 2200, "y": 300, "width": 1400, "height": 900}
    assert "x" in desktop._window_kwargs(saved, [LAPTOP, SIDECAR])


def test_offscreen_rect_drops_its_position_but_keeps_its_size():
    """The monitor it was saved on is gone. Size survives; placement is the OS's."""
    saved = {"x": 4000, "y": 2000, "width": 1600, "height": 1000}
    kwargs = desktop._window_kwargs(saved, [LAPTOP])
    assert kwargs == {"width": 1600, "height": 1000}


def test_barely_visible_rect_is_treated_as_offscreen():
    # 40px of the window pokes onto the screen — not enough of the drag strip
    # to grab, so it must not be restored there.
    saved = {"x": -1560, "y": 0, "width": 1600, "height": 1000}
    assert "x" not in desktop._window_kwargs(saved, [LAPTOP])


def test_position_is_dropped_when_no_screens_are_known():
    saved = {"x": 200, "y": 100, "width": 1600, "height": 1000}
    assert desktop._window_kwargs(saved, []) == {"width": 1600, "height": 1000}


def test_restore_uses_the_stored_rect(stored, monkeypatch):
    stored["rect"] = {"x": 200, "y": 100, "width": 1600, "height": 1000}
    monkeypatch.setattr(desktop, "_reset_requested", False)
    monkeypatch.setattr("webview.screens", [LAPTOP])
    assert desktop._restore_geometry() == stored["rect"]


def test_shift_at_launch_resets_and_clears_the_stored_rect(stored, monkeypatch):
    stored["rect"] = {"x": 200, "y": 100, "width": 1600, "height": 1000}
    monkeypatch.setattr(desktop, "_reset_requested", True)
    assert desktop._restore_geometry() == {"width": 1440, "height": 900}
    # Cleared immediately rather than at quit, so a crash still leaves defaults.
    assert stored["cleared"] is True


def test_sample_reset_modifier_latches(monkeypatch):
    monkeypatch.setattr(desktop, "_reset_requested", False)
    monkeypatch.setattr(desktop, "_shift_held", lambda: True)
    desktop._sample_reset_modifier()
    assert desktop._reset_requested is True
    # A later sample that misses the gesture must not un-latch it.
    monkeypatch.setattr(desktop, "_shift_held", lambda: False)
    desktop._sample_reset_modifier()
    assert desktop._reset_requested is True


def test_shift_held_is_false_on_platforms_without_a_probe(monkeypatch):
    monkeypatch.setattr(desktop.sys, "platform", "linux")
    assert desktop._shift_held() is False


def test_closing_handler_returns_none(monkeypatch):
    """closing is a locking event: a False return cancels the close."""
    monkeypatch.setattr(desktop, "_geometry", None)
    assert desktop._on_closing() is None


def test_move_and_resize_update_only_their_own_axis(monkeypatch):
    monkeypatch.setattr(desktop, "_geometry", (10, 20, 1440, 900))
    monkeypatch.setattr(desktop, "_schedule_geometry_persist", lambda: None)
    desktop._on_moved(300, 150)
    assert desktop._geometry == (300, 150, 1440, 900)
    desktop._on_resized(1600, 1000)
    assert desktop._geometry == (300, 150, 1600, 1000)


def test_events_before_the_window_is_shown_are_ignored(monkeypatch):
    """Without a seed from _on_shown there is no full rect worth persisting."""
    monkeypatch.setattr(desktop, "_geometry", None)
    desktop._on_moved(300, 150)
    desktop._on_resized(1600, 1000)
    assert desktop._geometry is None
