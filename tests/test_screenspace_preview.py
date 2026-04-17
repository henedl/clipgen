"""Smoke tests for the Screenspace Model View preview builder.

Verifies that :func:`screenspace_preview.build_preview` returns a
PNG-encodable BGR array for every supported tool type, given a synthetic
frame and region.
"""

import cv2
import numpy as np
import pytest

import screenspace_preview


@pytest.fixture
def synthetic_frame() -> np.ndarray:
    """A 320x240 BGR frame with a varying gradient plus a contrasting patch."""
    frame = np.zeros((240, 320, 3), dtype=np.uint8)
    for y in range(240):
        frame[y, :, 0] = y % 256  # blue gradient
    for x in range(320):
        frame[:, x, 1] = x % 256  # green gradient
    frame[50:110, 80:180] = (20, 200, 220)  # distinct patch
    return frame


@pytest.fixture
def prev_frame() -> np.ndarray:
    """A slightly different frame so change/flow produce non-trivial output."""
    frame = np.zeros((240, 320, 3), dtype=np.uint8)
    for y in range(240):
        frame[y, :, 0] = (y + 5) % 256
    for x in range(320):
        frame[:, x, 1] = (x + 5) % 256
    frame[60:120, 90:190] = (30, 190, 210)
    return frame


@pytest.fixture
def region() -> dict[str, int]:
    return {"x": 60, "y": 40, "w": 120, "h": 80}


ALL_TOOLS = [
    "color",
    "change",
    "similarity",
    "text",
    "numbers",
    "timelapse",
    "template",
    "flow",
    "scene",
    "inactivity",
]


@pytest.mark.parametrize("tool", ALL_TOOLS)
def test_build_preview_returns_encodable_image(
    synthetic_frame: np.ndarray,
    prev_frame: np.ndarray,
    region: dict[str, int],
    tool: str,
) -> None:
    """Every tool produces a non-empty BGR image that PNG-encodes successfully."""
    params: dict = {}
    if tool == "template":
        params["template_image"] = synthetic_frame[50:110, 80:180].copy()

    img = screenspace_preview.build_preview(
        synthetic_frame, prev_frame, region, tool, params
    )
    assert isinstance(img, np.ndarray)
    assert img.ndim == 3 and img.shape[2] == 3
    assert img.dtype == np.uint8
    assert img.size > 0

    png_bytes = screenspace_preview.encode_png(img)
    assert png_bytes and len(png_bytes) > 50
    # Re-decode round-trip to confirm well-formed PNG.
    decoded = cv2.imdecode(np.frombuffer(png_bytes, dtype=np.uint8), cv2.IMREAD_COLOR)
    assert decoded is not None
    assert decoded.shape[2] == 3


def test_missing_region_returns_placeholder(synthetic_frame: np.ndarray) -> None:
    """Tools that need a region degrade gracefully when none is supplied."""
    img = screenspace_preview.build_preview(synthetic_frame, None, None, "change", {})
    assert isinstance(img, np.ndarray)
    assert img.size > 0


def test_multitool_previews_first_step(
    synthetic_frame: np.ndarray, region: dict[str, int]
) -> None:
    """Multitool forwards to its first step's preview."""
    params = {"steps": [{"type": "scene", "parameters": {}}]}
    img = screenspace_preview.build_preview(
        synthetic_frame, None, region, "multitool", params
    )
    assert img.size > 0


def test_unknown_tool_returns_placeholder(synthetic_frame: np.ndarray) -> None:
    img = screenspace_preview.build_preview(synthetic_frame, None, None, "bogus", {})
    assert img.size > 0
