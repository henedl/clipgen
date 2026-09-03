"""Static regression checks for Overview frontend sources.

The Overview page is a hub (overview.js) + tab satellites
(overview-{convergence,metadata,reports}.js) sharing state through the
window.ClipgenOverview (OV) namespace. These checks pin the contracts the
satellite-wiring test can't see (bare cross-file *reads*, script order,
namespace hygiene).
"""

import re

from _frontend_source import WEB as _WEB
from _frontend_source import assert_es5

OVERVIEW_HTML = _WEB / "overview.html"


def _overview_js_files():
    return sorted(_WEB.glob("overview*.js"))


def test_no_studio_globals_in_overview_sources():
    """The moved tabs must read window.ClipgenOverview, never window._studio*."""
    for path in _overview_js_files():
        assert "_studio" not in path.read_text(encoding="utf-8"), (
            f"{path.name} still references a window._studio* global"
        )


def test_script_order_hub_before_satellites():
    """Load order is a contract: the hub (overview.js) before every satellite."""
    html = OVERVIEW_HTML.read_text(encoding="utf-8")
    scripts = re.findall(r'<script src="([^"]+)"', html)
    hub = scripts.index("overview.js")
    for satellite in (
        "overview-convergence.js",
        "overview-metadata.js",
        "overview-reports.js",
    ):
        assert hub < scripts.index(satellite), f"{satellite} loads before the hub"
    assert scripts.index("intake-cluster.js") < hub


def test_tab_order_and_wip_badge():
    """Tabs read Metadata | Convergence | Reports; the Reports tab carries the
    WIP pill (the newest surface)."""
    html = OVERVIEW_HTML.read_text(encoding="utf-8")
    tabs = re.findall(r'data-tab="(\w+)"', html)
    assert tabs == ["metadata", "convergence", "reports"]
    button = html[html.index('data-tab="reports"') :]
    button = button[: button.index("</button>")]
    assert "cg-wip-badge" in button, "reports tab lost its WIP badge"


def test_staleness_is_version_based_and_running_check_is_strict():
    """The 'data has changed' banners compare the hub dataVersion (no length
    heuristics), and only queued/running task statuses count as running —
    failed/cancelled/paused tasks must not claim an analysis is in flight."""
    hub = (_WEB / "overview.js").read_text(encoding="utf-8")
    assert "state.dataVersion++" in hub
    assert "tabState._snapshot = { version: state.dataVersion }" in hub
    md = (_WEB / "overview-metadata.js").read_text(encoding="utf-8")
    assert "createStalenessTracker(mdState)" in md
    assert 's === "queued" || s === "running"' in md
    cv = (_WEB / "overview-convergence.js").read_text(encoding="utf-8")
    assert "cvState._snapshot = { version: state.dataVersion }" in cv
    assert "createStalenessTracker(cvState)" in cv
    sm = (_WEB / "overview-reports.js").read_text(encoding="utf-8")
    assert "createStalenessTracker(rpState)" in sm
    for src in (md, cv, sm):
        assert "_snapshot.ss" not in src  # old length-compare heuristic


def test_hub_wires_shared_chrome_buttons():
    """Theme toggle + Settings must be wired on this page (TopNav only renders
    the buttons; each surface wires them — a missed call means dead chrome)."""
    hub = (_WEB / "overview.js").read_text(encoding="utf-8")
    assert "initThemeToggle()" in hub
    # Settings wiring goes through settings-modal.js's shared helper, which
    # owns the #settingsBtn lookup.
    assert "wireSettingsButton" in hub


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


def test_hidden_utility_class_defined():
    """tokens.css must define the generic `.hidden` utility (CSS toggle
    completeness, agents/CODE-REVIEW.md): the Convergence/Metadata satellites
    toggle `.hidden` on their status banners and controls, and every app page
    (this one included) gets the rule from tokens.css. Without it the
    "analysis running"/"data changed" banners render always."""
    css = (_WEB / "tokens.css").read_text(encoding="utf-8")
    assert re.search(r"^\.hidden \{\n  display: none !important;", css, re.MULTILINE)


def test_es5_discipline_in_overview_sources():
    """House style: no arrows / async-await in the Overview page scripts."""
    for path in _overview_js_files():
        assert_es5(path.read_text(encoding="utf-8"), path.name)
