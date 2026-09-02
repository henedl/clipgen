"""Render the clipgen brand mark to .icns (macOS) and .ico (Windows).

Run from the project root:

    uv run build/render_icons.py

Outputs:
    build/clipgen.icns
    build/clipgen.ico

The script draws the three nested L-shapes from the brand mark spec
(viewBox 0-100, stroke-width 11, square caps, miter joins) onto a solid
dark canvas. The mark is inset to ~22% on every side so it reads well
inside the macOS dock/squircle frame.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw
import itertools

BG = (10, 10, 10, 255)  # #0a0a0a
FG = (250, 250, 247, 255)  # #fafaf7

# Paths in the 100u mark viewBox; each tuple is a sequence of (x, y) corners.
PATHS: list[list[tuple[float, float]]] = [
    [(18, 18), (82, 18), (82, 82)],
    [(18, 40), (60, 40), (60, 82)],
    [(18, 62), (38, 62), (38, 82)],
]
STROKE = 11.0  # matches stroke-width="11" in mark-dark.svg
INSET_FRAC = 0.22  # mark area is ~56% of the canvas (matches apple-touch-icon.svg)

ICNS_SIZES = [16, 32, 64, 128, 256, 512, 1024]
ICO_SIZES = [16, 24, 32, 48, 64, 128, 256]


def render(size: int) -> Image.Image:
    """Render the brand mark at the given square pixel size."""
    img = Image.new("RGBA", (size, size), BG)
    draw = ImageDraw.Draw(img)

    mark_size = size * (1.0 - 2 * INSET_FRAC)
    offset = (size - mark_size) / 2.0
    scale = mark_size / 100.0
    half = (STROKE * scale) / 2.0

    def to_canvas(x: float, y: float) -> tuple[float, float]:
        return (offset + x * scale, offset + y * scale)

    for corners in PATHS:
        for a, b in itertools.pairwise(corners):
            x1, y1 = to_canvas(*a)
            x2, y2 = to_canvas(*b)
            left = min(x1, x2) - half
            right = max(x1, x2) + half
            top = min(y1, y2) - half
            bottom = max(y1, y2) + half
            draw.rectangle((left, top, right, bottom), fill=FG)

    return img


def write_icns(out: Path) -> None:
    """Build a .icns via macOS iconutil, falling back to PIL multi-frame PNG."""
    iconutil = shutil.which("iconutil")
    if iconutil is None:
        # No iconutil: skip rather than ship an invalid .icns.
        print("iconutil not found; skipping .icns generation (macOS only)")
        return

    with tempfile.TemporaryDirectory() as tmp:
        iconset = Path(tmp) / "clipgen.iconset"
        iconset.mkdir()
        # Apple-required iconset filenames map sizes to @1x/@2x pairs.
        mapping = [
            (16, "icon_16x16.png"),
            (32, "icon_16x16@2x.png"),
            (32, "icon_32x32.png"),
            (64, "icon_32x32@2x.png"),
            (128, "icon_128x128.png"),
            (256, "icon_128x128@2x.png"),
            (256, "icon_256x256.png"),
            (512, "icon_256x256@2x.png"),
            (512, "icon_512x512.png"),
            (1024, "icon_512x512@2x.png"),
        ]
        for size, name in mapping:
            render(size).save(iconset / name, "PNG")
        subprocess.run(
            [iconutil, "-c", "icns", str(iconset), "-o", str(out)],
            check=True,
        )
    print(f"wrote {out}")


def write_ico(out: Path) -> None:
    """Build a multi-resolution .ico via Pillow."""
    base = render(max(ICO_SIZES))
    base.save(out, format="ICO", sizes=[(s, s) for s in ICO_SIZES])
    print(f"wrote {out}")


def main() -> int:
    out_dir = Path(__file__).resolve().parent
    write_icns(out_dir / "clipgen.icns")
    write_ico(out_dir / "clipgen.ico")
    return 0


if __name__ == "__main__":
    sys.exit(main())
