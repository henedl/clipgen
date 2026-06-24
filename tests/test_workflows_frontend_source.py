"""Static regression checks for Workflows frontend sources (M0 scaffold).

Globs ``workflows*.js`` so assertions stay valid as the hub is carved into
satellites (workflows-canvas / -wires / -nodes / -runs / -stashes) in later
milestones — mirroring tests/test_studio_frontend_source.py.
"""

from pathlib import Path

_WEB = Path(__file__).resolve().parent.parent / "assets" / "web"
WORKFLOWS_HTML = _WEB / "workflows.html"
WORKFLOWS_CSS = _WEB / "workflows.css"


def _workflows_js() -> str:
    return "".join(
        p.read_text(encoding="utf-8") for p in sorted(_WEB.glob("workflows*.js"))
    )


def test_hub_publishes_namespace_and_state():
    src = _workflows_js()
    # The hub must establish the shared namespace satellites attach to, and
    # publish `state` onto it (routed through WF.state, not a bare cross-file var).
    assert "window.ClipgenWorkflows" in src
    assert "WF.state = state" in src
    assert "var state = {" in src


def test_page_loads_hub_and_shared_chrome():
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    assert 'data-frontend="workflows"' in html
    assert "<topnav-mount" in html
    # Dependency order: utils + topnav before the page hub.
    assert html.index("utils.js") < html.index("workflows.js")
    assert html.index("topnav.js") < html.index("workflows.js")


def test_css_uses_shared_tokens_not_hardcoded_surfaces():
    css = WORKFLOWS_CSS.read_text(encoding="utf-8")
    # Layout should lean on design tokens for surfaces/spacing.
    assert "var(--bg)" in css
    assert "var(--space-4)" in css


def test_hub_wires_topnav_chrome():
    """TopNav renders theme toggle + Settings on every page; the hub must wire
    them (they appear but do nothing otherwise)."""
    src = _workflows_js()
    assert "initThemeToggle(" in src
    assert "#settingsBtn" in src
    assert "openSettingsModal(" in src


def test_start_overlay_treats_workflows_as_video_tool():
    """Video-only --workflows launches must not auto-open the Start overlay
    when videos are already present (parity with Screenspace/Transcripts)."""
    overlay = (_WEB / "start-overlay.js").read_text(encoding="utf-8")
    assert 'path.indexOf("/workflows/") === 0' in overlay


def test_satellites_read_shared_state_through_wf():
    """The carve gotcha: satellites must read `WF.state`, not redeclare a
    divergent `var state = {`. Only the hub owns the literal state object."""
    for name in ("workflows-nodes.js", "workflows-canvas.js"):
        src = (_WEB / name).read_text(encoding="utf-8")
        assert "var state = WF.state" in src, name
        assert "var state = {" not in src, name


def test_page_loads_satellites_after_hub_with_toolbar():
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    # Satellites load after the hub (load-order contract).
    assert html.index("workflows.js") < html.index("workflows-nodes.js")
    assert html.index("workflows.js") < html.index("workflows-canvas.js")
    # Canvas + blueprint-switcher structure the M1 hub/satellites target.
    assert 'id="wfWorld"' in html
    assert 'id="wfToolbar"' in html
    assert 'id="wfBlueprintSelect"' in html
    # Per code-review rule: text inputs need autocomplete off.
    assert 'id="wfBlueprintName"' in html
    assert 'autocomplete="off"' in html


def test_hub_and_satellites_publish_canvas_hooks():
    src = _workflows_js()
    # Hub-owned orchestration.
    for fn in ("WF.scheduleSave", "WF.renderPalette", "WF.openBlueprint"):
        assert fn in src, fn
    # Satellite-owned rendering / interaction attached back onto WF.
    for fn in (
        "WF.renderAllNodes",
        "WF.renderNode",
        "WF.initCanvas",
        "WF.applyViewport",
    ):
        assert fn in src, fn
    # The catalog is fetched (not hardcoded), and the world layer is transformed.
    assert "api/catalog" in src
    assert "scale(" in src and "translate(" in src
