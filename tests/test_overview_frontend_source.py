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


def test_tab_order_and_wip_badge():
    """Tabs read Metadata | Convergence | Map; the Map tab carries the WIP
    pill (the similarity map is the newest surface)."""
    html = OVERVIEW_HTML.read_text(encoding="utf-8")
    tabs = re.findall(r'data-tab="(\w+)"', html)
    assert tabs == ["metadata", "convergence", "map"]
    map_tab = html[html.index('data-tab="map"') :]
    map_tab = map_tab[: map_tab.index("</button>")]
    assert "ov-wip-badge" in map_tab


def test_staleness_is_version_based_and_running_check_is_strict():
    """The 'data has changed' banners compare the hub dataVersion (no length
    heuristics), and only queued/running task statuses count as running —
    failed/cancelled/paused tasks must not claim an analysis is in flight."""
    hub = (_WEB / "overview.js").read_text(encoding="utf-8")
    assert "state.dataVersion++" in hub
    md = (_WEB / "overview-metadata.js").read_text(encoding="utf-8")
    assert "mdState._snapshot = { version: state.dataVersion }" in md
    assert 's === "queued" || s === "running"' in md
    cv = (_WEB / "overview-convergence.js").read_text(encoding="utf-8")
    assert "cvState._snapshot = { version: state.dataVersion }" in cv
    for src in (md, cv):
        assert "_snapshot.ss" not in src  # old length-compare heuristic
    # The Map follows the same contract: re-activation with a newer hub
    # version re-fetches api/data (a stale matrix froze the layout after
    # Refresh while sibling tabs updated).
    mp = (_WEB / "overview-map.js").read_text(encoding="utf-8")
    assert "hubDataVersion() !== _matrixVersion" in mp
    assert "function reloadData()" in mp


def test_hub_wires_shared_chrome_buttons():
    """Theme toggle + Settings must be wired on this page (TopNav only renders
    the buttons; each surface wires them — a missed call means dead chrome)."""
    hub = (_WEB / "overview.js").read_text(encoding="utf-8")
    assert "initThemeToggle()" in hub
    assert "openSettingsModal" in hub


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


def test_map_color_by_choropleth():
    """The "color by" choropleth: one authority computes the idle dot look
    (dotBaseColor/dotBaseScale) and the replay glow composes on it — a
    hardcoded ramp inside applyReplayGlow would silently override the
    choropleth mid-replay."""
    src = (_WEB / "overview-map.js").read_text(encoding="utf-8")
    assert "function setColorBy(" in src
    assert "colorBy: null" in src
    assert "function dotBaseColor(" in src
    assert "function dotBaseScale(" in src
    glow_body = src[src.index("function applyReplayGlow()") :]
    glow_body = glow_body[: glow_body.index("\n  }")]
    assert "dotBaseColor(" in glow_body
    assert "dotBaseScale(" in glow_body
    html = OVERVIEW_HTML.read_text(encoding="utf-8")
    assert 'id="mapColorBySection"' in html
    assert 'id="mapColorByChip"' in html
    css = (_WEB / "overview.css").read_text(encoding="utf-8")
    assert ".map-feature-name" in css


def test_map_compare_arc():
    """The pairwise compare: distances come from the weighted z matrix (never
    projected coordinates), and the arc layer follows the standard
    rebuild/sync pattern so it travels through the re-layout lerp."""
    src = (_WEB / "overview-map.js").read_text(encoding="utf-8")
    assert "function computePairDiff(" in src
    assert "function rebuildCompare()" in src
    assert "function syncComparePositions()" in src
    diff_body = src[src.index("function computePairDiff(") :]
    diff_body = diff_body[: diff_body.index("\n  }")]
    assert "state.zRaw" in diff_body
    assert "state.coords" not in diff_body
    html = OVERVIEW_HTML.read_text(encoding="utf-8")
    assert 'id="mapCompareBtn"' in html
    css = (_WEB / "overview.css").read_text(encoding="utf-8")
    assert ".map-compare-swatch" in css
    assert "#mapCompareBtn.is-armed" in css


def test_map_direct_axes_mode():
    """Manual-axes layout: coords come from state.zRaw (unweighted — the
    group weights are a similarity lens, not an axis warp), PCA components
    are still computed for the outlier/edge/anchor layers, and the pickers
    UI exists with its CSS."""
    src = (_WEB / "overview-map.js").read_text(encoding="utf-8")
    assert 'layoutMode: "pca"' in src
    assert "function computeLayoutCoords()" in src
    assert "function renderAxisPickers()" in src
    body = src[src.index("function computeLayoutCoords()") :]
    body = body[: body.index("\n  }")]
    assert "state.zRaw" in body
    assert "pcaProject(state.weighted" in body  # components live on regardless
    html = OVERVIEW_HTML.read_text(encoding="utf-8")
    assert 'id="mapLayoutMode"' in html
    assert 'id="mapAxisPickers"' in html
    css = (_WEB / "overview.css").read_text(encoding="utf-8")
    assert ".map-axis-pickers" in css
    assert ".map-axis-picker-row" in css


def test_hidden_utility_class_defined():
    """overview.css must define the generic `.hidden` utility (CSS toggle
    completeness, agents/CODE-REVIEW.md): the moved Convergence/Metadata
    satellites toggle `.hidden` on their status banners and controls, and on
    Studio that rule came from studio.css — which this page doesn't load.
    Without it the "analysis running"/"data changed" banners render always."""
    css = (_WEB / "overview.css").read_text(encoding="utf-8")
    assert re.search(r"^\.hidden \{\n  display: none !important;", css, re.M)


def test_es5_discipline_in_overview_sources():
    """House style: no arrows / async-await in the Overview page scripts."""
    for path in _overview_js_files():
        src = path.read_text(encoding="utf-8")
        assert "=>" not in src, f"{path.name} uses an arrow function"
        # `img.decoding = "async"` is a DOM property, not the keyword.
        assert not re.search(r"\basync function\b|\bawait\s", src), (
            f"{path.name} uses async/await"
        )
