"""Ratchet hardcoded hex colors in JS toward getComputedStyle reads.

AGENTS.md: JS reads colors via getComputedStyle(...).getPropertyValue(...),
never hex literals. Legacy counts are frozen per file; a count may only go
down. Some frozen sites are arguably permanent (color-picker palette rows,
screenspace-color math constants) — they still live here, not in code, so
every new literal is a deliberate decision.
"""

import re

from _frontend_source import WEB, strip_comments

# Frozen per-file hex-literal counts. New violations fail; cleanups
# update the number downward.
_BASELINE = {
    "color-picker.js": 10,
    "composer-annotate.js": 10,
    "composer-timeline.js": 3,
    "composer.js": 1,
    "dev-token-tweak.js": 9,
    "overview-convergence.js": 2,
    "overview-metadata.js": 2,
    "screenspace-color.js": 12,
    "screenspace-multitool-params.js": 1,
    "screenspace-overlay.js": 6,
    "screenspace-timeline.js": 3,
    "screenspace-utils.js": 1,
    "screenspace.js": 3,
    "settings-modal.js": 4,
    "transcripts-video.js": 4,
    "utils.js": 34,
    "viewer.js": 3,
    "workflows-canvas.js": 2,
}

_HEX = re.compile(r"#[0-9a-fA-F]{3,8}\b")


def _counts() -> dict[str, int]:
    found = {}
    for path in sorted(WEB.glob("*.js")):
        hits = len(_HEX.findall(strip_comments(path.read_text(encoding="utf-8"))))
        if hits:
            found[path.name] = hits
    return found


def test_no_new_hex_colors_in_js() -> None:
    found = _counts()
    grew = {n: c for n, c in found.items() if c > _BASELINE.get(n, 0)}
    assert not grew, (
        f"new hex literals in {grew} — read the color via "
        "getComputedStyle(...).getPropertyValue(...) from a tokens.css var"
    )


def test_hex_baseline_ratchets_down() -> None:
    found = _counts()
    shrank = {n: c for n, c in _BASELINE.items() if found.get(n, 0) < c}
    assert not shrank, f"nice — ratchet the baseline down for {shrank}"


def test_the_scan_sees_hex_spellings() -> None:
    assert _HEX.findall('ctx.fillStyle = "#888";') == ["#888"]
    assert _HEX.findall('var c = "#0891b2cc";') == ["#0891b2cc"]
    assert not _HEX.findall('qs("#summaryStream")')
