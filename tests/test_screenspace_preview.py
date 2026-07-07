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


def test_template_heatmap_with_mask_matches_scan_prep(
    synthetic_frame: np.ndarray,
) -> None:
    """A masked template heatmap uses the scan's binarized-mask prep helper."""
    template = synthetic_frame[10:50, 10:50].copy()
    # Soft-alpha mask: an opaque center on a semi-transparent border. The scan
    # binarizes this (>=128 -> 255); the preview must do the same so its heatmap
    # reflects the real scan rather than a blurred approximation.
    mask = np.full((40, 40), 100, dtype=np.uint8)
    mask[10:30, 10:30] = 255
    params = {"template_image": template, "template_mask": mask}

    img = screenspace_preview.build_overlay_layer(
        synthetic_frame, None, None, "template", "match_heatmap", params
    )
    assert img is not None
    assert img.dtype == np.uint8
    assert img.ndim == 3 and img.shape[2] == 3
    # Frame-scoped: the heatmap covers the whole frame after border replication.
    assert img.shape[:2] == synthetic_frame.shape[:2]


def test_composite_template_preview_binarizes_mask(
    synthetic_frame: np.ndarray,
) -> None:
    """The composite template preview binarizes the mask like the real scan.

    ``_prepare_template`` thresholds the alpha mask at 127, so a soft mask and
    its pre-binarized equivalent must yield an identical composite. A blurred
    mask (the previous behavior) would let the semi-transparent border leak into
    ``TM_CCOEFF_NORMED`` and diverge between the two.
    """
    template = synthetic_frame[10:50, 10:50].copy()
    soft = np.full((40, 40), 100, dtype=np.uint8)  # semi-transparent border
    soft[10:30, 10:30] = 255  # opaque center
    hard = np.where(soft >= 128, 255, 0).astype(np.uint8)

    img_soft = screenspace_preview.build_preview(
        synthetic_frame,
        None,
        None,
        "template",
        {"template_image": template, "template_mask": soft},
    )
    img_hard = screenspace_preview.build_preview(
        synthetic_frame,
        None,
        None,
        "template",
        {"template_image": template, "template_mask": hard},
    )
    assert np.array_equal(img_soft, img_hard)


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


# ---- Drawing-primitive scaling ----
#
# Region-scope overlay layers are returned at native region resolution and
# the browser scales them down to fit the display rect. So drawing primitives
# (Flow arrows, Canny edges) must scale with the source region's pixel size,
# otherwise they shrink to invisible widths on full-frame views of large
# videos. These tests assert that primitives grow in absolute pixel count
# when the region grows — a proxy for "arrows / edges actually got bigger".


def _green_arrow_pixel_count(img: np.ndarray) -> int:
    """Count BGR (40, 220, 40) arrow pixels."""
    b, g, r = cv2.split(img)
    return int(((g > 150) & (r < 100) & (b < 100)).sum())


def test_overlay_flow_arrow_scales_with_region() -> None:
    """Larger flow regions produce visibly larger arrows.

    With the previous hardcoded ``scale=4.0`` and ``thickness=1``, both region
    sizes drew identically-sized arrows; with proportional scaling the larger
    region's arrows are both longer (per grid cell) and thicker.
    """
    rng = np.random.default_rng(42)
    frame_h, frame_w = 900, 700
    base = rng.integers(0, 256, (frame_h, frame_w, 3), dtype=np.uint8)
    prev_full = base
    # Horizontal shift produces a strong, uniform optical-flow signal.
    curr_full = np.roll(base, shift=3, axis=1)

    small_region = {"x": 50, "y": 50, "w": 150, "h": 200}
    large_region = {"x": 50, "y": 50, "w": 600, "h": 800}

    small = screenspace_preview.build_overlay_layer(
        curr_full, prev_full, small_region, "flow", "flow_vectors", {}
    )
    large = screenspace_preview.build_overlay_layer(
        curr_full, prev_full, large_region, "flow", "flow_vectors", {}
    )
    assert small is not None and large is not None
    assert small.shape[:2] == (small_region["h"], small_region["w"])
    assert large.shape[:2] == (large_region["h"], large_region["w"])

    small_green = _green_arrow_pixel_count(small)
    large_green = _green_arrow_pixel_count(large)
    assert small_green > 0, "small region produced no arrows"
    assert large_green > small_green * 4, (
        f"larger region arrows should occupy substantially more pixels "
        f"(small={small_green}, large={large_green})"
    )


def test_overlay_scene_edges_thicken_with_region() -> None:
    """Larger scene/edges regions produce visibly thicker Canny edges.

    Canny output is 1-px regardless of source size; without proportional
    dilation, edges become hairlines when the browser scales a large region
    down to display size.
    """
    frame_h, frame_w = 900, 700
    frame = np.zeros((frame_h, frame_w, 3), dtype=np.uint8)
    # Horizontal stripe pattern — uniform edge density per unit area.
    for y in range(0, frame_h, 30):
        frame[y : y + 15, :] = 255

    small_region = {"x": 50, "y": 50, "w": 150, "h": 200}
    large_region = {"x": 50, "y": 50, "w": 600, "h": 800}

    small = screenspace_preview.build_overlay_layer(
        frame, None, small_region, "scene", "edges", {}
    )
    large = screenspace_preview.build_overlay_layer(
        frame, None, large_region, "scene", "edges", {}
    )
    assert small is not None and large is not None

    def edge_density(img: np.ndarray) -> float:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        return float((gray > 0).sum()) / (gray.shape[0] * gray.shape[1])

    small_density = edge_density(small)
    large_density = edge_density(large)
    assert small_density > 0
    assert large_density > small_density * 1.5, (
        f"larger region edges should be substantially denser due to dilation "
        f"(small={small_density:.3f}, large={large_density:.3f})"
    )
