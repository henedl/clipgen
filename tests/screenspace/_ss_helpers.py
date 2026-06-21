"""Shared frame/region builders for the Screenspace test suite.

Imported by sibling test modules via ``from _ss_helpers import ...`` (the same
bare-module import pattern already used elsewhere in ``tests/``). Not collected
as tests itself (filename does not match ``test_*``).
"""

import numpy as np

_ICON_BG = 50  # flat gray background used by the icon-frame fixtures


def _make_icon(size: int, seed: int = 7) -> np.ndarray:
    """Build a square icon with a textured center and a flat border.

    The border matches the fixture background colour so that Gaussian blur
    at the icon boundary does not distort cv2.matchTemplate correlation.
    """
    import cv2

    rng = np.random.RandomState(seed)
    base = 40
    icon = np.full((base, base, 3), _ICON_BG, dtype=np.uint8)
    # Leave a 5px flat border; fill the center with high-contrast texture
    icon[5:-5, 5:-5] = rng.randint(150, 255, (base - 10, base - 10, 3), dtype=np.uint8)
    if size == base:
        return icon
    return cv2.resize(icon, (size, size), interpolation=cv2.INTER_AREA)


def _make_icon_frame(
    frame_w: int,
    frame_h: int,
    icon_positions: list[tuple[int, int, int]],
    seed: int = 7,
) -> np.ndarray:
    """Build a frame with identical icons (possibly at different sizes) placed
    at *icon_positions* (list of ``(x, y, size)``)."""
    frame = np.full((frame_h, frame_w, 3), _ICON_BG, dtype=np.uint8)
    for x, y, s in icon_positions:
        frame[y : y + s, x : x + s] = _make_icon(s, seed=seed)
    return frame


def _gray_with_red_patch(patch=10):
    """100x100 gray region with a ``patch``x``patch`` dark-red corner block.

    The region average is gray (so ``color_matches`` misses the red), but the
    red is present per-pixel for ``color_present`` to find.
    """
    region = np.full((100, 100, 3), 128, dtype=np.uint8)  # neutral gray
    region[0:patch, 0:patch] = [0, 0, 139]  # BGR dark red
    return region
