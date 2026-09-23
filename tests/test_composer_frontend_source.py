"""Static regression checks for the Composer frontend sources.

The page is a hub (composer.js) plus satellites sharing state through
window.ClipgenComposer (CO). These pin the contracts the satellite-wiring
test cannot see: script order, namespace hygiene, and the ES5 house style.
"""

import re

from _frontend_source import WEB as _WEB
from _frontend_source import assert_es5, read

_SATELLITES = (
    "composer-markers.js",
    "composer-timeline.js",
    "composer-annotate.js",
    "composer-scrub.js",
)


def test_script_order_hub_before_satellites():
    scripts = re.findall(r'<script src="([^"]+)"', read("composer.html"))
    hub = scripts.index("composer.js")
    for satellite in _SATELLITES:
        assert hub < scripts.index(satellite), f"{satellite} loads before the hub"


def test_every_satellite_is_loaded():
    scripts = set(re.findall(r'<script src="([^"]+)"', read("composer.html")))
    on_disk = {p.name for p in _WEB.glob("composer-*.js")}
    assert on_disk <= scripts, sorted(on_disk - scripts)


def test_satellites_reach_the_hub_through_the_namespace():
    for name in _SATELLITES:
        src = read(name)
        assert "window.ClipgenComposer" in src, f"{name} never opens the CO namespace"
        assert "window._studio" not in src


def test_composer_sources_are_es5():
    for path in sorted(_WEB.glob("composer*.js")):
        assert_es5(path.read_text(encoding="utf-8"), path.name)
