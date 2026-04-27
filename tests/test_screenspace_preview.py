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


# ---- Overlay layers ----


def _all_overlay_pairs() -> list[tuple[str, str, str]]:
    out: list[tuple[str, str, str]] = []
    for tool, layers in screenspace_preview.OVERLAY_LAYERS.items():
        for lid, _label, scope in layers:
            out.append((tool, lid, scope))
    return out


@pytest.mark.parametrize("tool,layer,scope", _all_overlay_pairs())
def test_build_overlay_layer_shape_matches_scope(
    synthetic_frame: np.ndarray,
    prev_frame: np.ndarray,
    region: dict[str, int],
    tool: str,
    layer: str,
    scope: str,
) -> None:
    """Region-scoped layers match the region rect; frame-scoped layers match the frame."""
    params: dict = {}
    if tool == "template":
        # Use a slice that includes the gradient (non-zero std) so the
        # heatmap layer can be computed.
        params["template_image"] = synthetic_frame[10:50, 10:50].copy()

    img = screenspace_preview.build_overlay_layer(
        synthetic_frame, prev_frame, region, tool, layer, params
    )
    assert img is not None, f"overlay layer {tool}/{layer} returned None"
    assert img.dtype == np.uint8
    assert img.ndim == 3 and img.shape[2] == 3

    if scope == "region":
        assert img.shape[:2] == (region["h"], region["w"])
    elif scope == "frame":
        assert img.shape[:2] == synthetic_frame.shape[:2]


def test_overlay_layer_excluded_tools_have_no_entry() -> None:
    """timelapse and inactivity are intentionally not overlay-eligible."""
    assert "timelapse" not in screenspace_preview.OVERLAY_LAYERS
    assert "inactivity" not in screenspace_preview.OVERLAY_LAYERS


def test_overlay_layer_unknown_returns_none(
    synthetic_frame: np.ndarray, region: dict[str, int]
) -> None:
    img = screenspace_preview.build_overlay_layer(
        synthetic_frame, None, region, "change", "not_a_layer", {}
    )
    assert img is None


def test_overlay_layer_no_region_returns_none(synthetic_frame: np.ndarray) -> None:
    img = screenspace_preview.build_overlay_layer(
        synthetic_frame, None, None, "color", "region", {}
    )
    assert img is None


def test_overlay_change_without_prev_only_yields_gray_blur(
    synthetic_frame: np.ndarray, region: dict[str, int]
) -> None:
    """change/abs_diff and change/mask need a previous frame; gray_blur does not."""
    gray_blur = screenspace_preview.build_overlay_layer(
        synthetic_frame, None, region, "change", "gray_blur", {}
    )
    assert gray_blur is not None and gray_blur.shape[:2] == (region["h"], region["w"])
    assert (
        screenspace_preview.build_overlay_layer(
            synthetic_frame, None, region, "change", "abs_diff", {}
        )
        is None
    )
    assert (
        screenspace_preview.build_overlay_layer(
            synthetic_frame, None, region, "change", "mask", {}
        )
        is None
    )


def test_overlay_layer_encodes_at_native_resolution(
    synthetic_frame: np.ndarray, region: dict[str, int]
) -> None:
    """encode_png(cap_width=False) keeps the overlay's native size."""
    img = screenspace_preview.build_overlay_layer(
        synthetic_frame, None, region, "color", "region", {}
    )
    assert img is not None
    png = screenspace_preview.encode_png(img, cap_width=False)
    decoded = cv2.imdecode(np.frombuffer(png, dtype=np.uint8), cv2.IMREAD_COLOR)
    assert decoded is not None
    assert decoded.shape[:2] == (region["h"], region["w"])


def test_overlay_multitool_inherits_first_step(
    synthetic_frame: np.ndarray, region: dict[str, int]
) -> None:
    params = {"steps": [{"type": "scene", "parameters": {}}]}
    img = screenspace_preview.build_overlay_layer(
        synthetic_frame, None, region, "multitool", "edges", params
    )
    assert img is not None
    assert img.shape[:2] == (region["h"], region["w"])


def test_overlay_layer_scope_helper() -> None:
    assert screenspace_preview.overlay_layer_scope("change", "mask") == "region"
    assert (
        screenspace_preview.overlay_layer_scope("template", "match_heatmap") == "frame"
    )
    assert screenspace_preview.overlay_layer_scope("timelapse", "anything") is None
    assert screenspace_preview.overlay_layer_scope("change", "not_a_layer") is None
