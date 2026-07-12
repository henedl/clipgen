"""Static regression checks for Overview frontend sources.

The Overview page is a hub (overview.js) + tab satellites
(overview-{map,convergence,metadata}.js) sharing state through the
window.ClipgenOverview (OV) namespace. These checks pin the contracts the
satellite-wiring test can't see (bare cross-file *reads*, script order,
namespace hygiene).
"""

import re
from pathlib import Path

_WEB = Path(__file__).resolve().parent.parent / "assets" / "web"
OVERVIEW_HTML = _WEB / "overview.html"


def _overview_js_files():
    return sorted(_WEB.glob("overview*.js"))


def test_no_studio_globals_in_overview_sources():
    """The moved tabs must read window.ClipgenOverview, never window._studio*."""
    for path in _overview_js_files():
        assert "_studio" not in path.read_text(encoding="utf-8"), (
            f"{path.name} still references a window._studio* global"
        )


def test_script_order_three_before_map_hub_before_satellites():
    """Load order is a contract: vendored THREE before overview-map.js, and
    the hub (overview.js) before every satellite."""
    html = OVERVIEW_HTML.read_text(encoding="utf-8")
    scripts = re.findall(r'<script src="([^"]+)"', html)
    assert scripts.index("vendor/three.min.js") < scripts.index("overview-map.js")
    hub = scripts.index("overview.js")
    for satellite in (
        "overview-map.js",
        "overview-convergence.js",
        "overview-metadata.js",
    ):
        assert hub < scripts.index(satellite), f"{satellite} loads before the hub"
    assert scripts.index("intake-cluster.js") < hub


def test_hub_reads_metadata_cluster_setting_from_sheet_payload():
    """The STUDIO_METADATA_CLUSTER_SCREENSPACE setting reaches Overview via the
    ../studio/api/sheet payload (server.py includes it in both branches)."""
    src = (_WEB / "overview.js").read_text(encoding="utf-8")
    assert "metadataClusterScreenspace" in src
    assert "data.metadataClusterScreenspace" in src
    assert '"../studio/api/sheet"' in src


def test_metadata_screenspace_clustering():
    """overview-metadata.js counts Screenspace as clusters (default on)."""
    src = (_WEB / "overview-metadata.js").read_text(encoding="utf-8")
    assert "SS_CLUSTER_THRESHOLD_SEC" in src
    assert "function clustersToEventLike(clusters)" in src
    assert "state.metadataClusterScreenspace !== false" in src
    assert '"Screenspace clusters"' in src
    # Collisions keep raw events (they cluster internally already).
    assert "computeCollisions(activeP, rows, events, marks" in src


def test_map_layers_and_drilldown_present():
    """The Map tab's link layers + drill-down surface (edges computed in the
    full weighted feature space, not the projection; burst capped)."""
    src = (_WEB / "overview-map.js").read_text(encoding="utf-8")
    assert "function computeSimilarityEdges()" in src
    assert "function placeAnchors()" in src
    assert "function buildParticipantItems(pid)" in src
    assert "function showBurst(idx, items)" in src
    # Edges come from the weighted z matrix, never projected coordinates.
    edges_body = src[src.index("function computeSimilarityEdges()") :]
    edges_body = edges_body[: edges_body.index("\n  }")]
    assert "state.weighted" in edges_body
    assert "state.coords" not in edges_body
    assert re.search(r"var BURST_CAP = \d+;", src)


def test_es5_discipline_in_overview_sources():
    """House style: no arrows / async-await in the Overview page scripts."""
    for path in _overview_js_files():
        src = path.read_text(encoding="utf-8")
        assert "=>" not in src, f"{path.name} uses an arrow function"
        # `img.decoding = "async"` is a DOM property, not the keyword.
        assert not re.search(r"\basync function\b|\bawait\s", src), (
            f"{path.name} uses async/await"
        )
