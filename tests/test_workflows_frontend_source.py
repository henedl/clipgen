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
    for name in ("workflows-nodes.js", "workflows-canvas.js", "workflows-wires.js"):
        src = (_WEB / name).read_text(encoding="utf-8")
        assert "var state = WF.state" in src, name
        assert "var state = {" not in src, name


def test_page_loads_satellites_after_hub_with_toolbar():
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    # Satellites load after the hub (load-order contract); -wires destructures
    # canvas-published helpers via WF.* late-binding, but still loads last.
    assert html.index("workflows.js") < html.index("workflows-nodes.js")
    assert html.index("workflows.js") < html.index("workflows-canvas.js")
    assert html.index("workflows-canvas.js") < html.index("workflows-wires.js")
    # Canvas + blueprint-switcher structure the M1 hub/satellites target.
    assert 'id="wfWorld"' in html
    assert 'id="wfToolbar"' in html
    assert 'id="wfBlueprintSelect"' in html
    # Per code-review rule: text inputs need autocomplete off.
    assert 'id="wfBlueprintName"' in html
    assert 'autocomplete="off"' in html


def test_canvas_is_gated_until_a_blueprint_is_active():
    """Edits must not be possible (or silently lost) before a blueprint loads,
    and a load failure must surface a persistent, retryable error — not a canvas
    that looks editable but never saves."""
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    assert 'id="wfCanvasOverlay"' in html
    assert 'id="wfOverlayRetry"' in html
    # The blueprint toolbar starts disabled (no active blueprint yet).
    assert html.count("disabled") >= 4

    src = _workflows_js()
    # Readiness is the authoritative gate: interaction handlers no-op until
    # ready, the hub flips it only after a blueprint opens, and a failed load
    # shows the error state.
    assert "if (!state.ready) return" in src
    assert 'setCanvasState("ready")' in src
    assert 'setCanvasState("error")' in src


def test_hub_and_satellites_publish_canvas_hooks():
    src = _workflows_js()
    # Hub-owned orchestration (nodeContextMet is published for the nodes
    # satellite so palette grey-out logic isn't duplicated).
    for fn in (
        "WF.scheduleSave",
        "WF.renderPalette",
        "WF.openBlueprint",
        "WF.nodeContextMet",
    ):
        assert fn in src, fn
    # Satellite-owned rendering / interaction attached back onto WF.
    for fn in (
        "WF.renderAllNodes",
        "WF.renderNode",
        "WF.initCanvas",
        "WF.applyViewport",
        "WF.autoArrange",  # "Clean up" auto-layout
        # Wires satellite (M2): connector layer + drag-to-connect + validation.
        "WF.initWires",
        "WF.renderWires",
        "WF.startWireDrag",
        "WF.isConnecting",
        "WF.cancelConnect",
        "WF.selectEdge",
        "WF.removeEdge",
        "WF.canConnect",
    ):
        assert fn in src, fn
    # The catalog is fetched (not hardcoded), and the world layer is transformed.
    assert "api/catalog" in src
    assert "scale(" in src and "translate(" in src


def test_wires_satellite_present_and_typed():
    """M2 wire layer: the SVG element, the connector satellite, and exact-type
    validation (the M3 adapter seam) all ship."""
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    assert 'id="wfWires"' in html
    assert 'id="wfWireDelete"' in html
    assert "workflows-wires.js" in html
    wires = (_WEB / "workflows-wires.js").read_text(encoding="utf-8")
    # Edges are typed; M2 connects only on exact type match.
    assert "outType === inType" in wires
    # Endpoints come from live node position + cached port offsets.
    assert "portWorldPos" in wires


def test_css_defines_wire_and_param_styles():
    css = WORKFLOWS_CSS.read_text(encoding="utf-8")
    # Wire layer + connector styling.
    assert ".wf-wires" in css
    assert ".wf-wire" in css
    assert "non-scaling-stroke" in css  # constant on-screen wire width under zoom
    # Param editors + on-canvas validation cues.
    assert ".wf-param" in css
    assert ".wf-node.invalid" in css
    assert ".wf-node.disabled" in css
    # Connect-mode compatibility glow / dim (cards + palette).
    assert ".wf-compatible" in css
    assert ".wf-dim" in css


def test_connect_ux_polish_present():
    """Click-to-arm wiring, drop-snapping, compatibility highlight, and the
    'Clean up' auto-layout all ship."""
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    assert 'id="wfCleanUp"' in html
    wires = (_WEB / "workflows-wires.js").read_text(encoding="utf-8")
    # Click-to-arm: a second click resolves where the cursor lands.
    assert "armConnect" in wires
    assert "elementFromPoint" in wires
    # Drop-snapping picks the nearest compatible port on the target card.
    assert "connectToCard" in wires
    # Compatibility highlight spans cards and palette items.
    assert "applyConnectHighlight" in wires
    assert "wf-compatible" in wires
    canvas = (_WEB / "workflows-canvas.js").read_text(encoding="utf-8")
    assert "function autoArrange" in canvas
    # Review fix: selecting nodes clears any wire selection (Delete hits nodes).
    assert "state.selectedEdge = null" in canvas
    # Review fix: the wire-delete button doesn't fall through to pan.
    assert "#wfWireDelete" in canvas


def test_in_flight_connect_is_torn_down_on_context_change():
    """An armed wire must not outlive its source node: switching blueprints and
    re-laying-out the canvas both cancel the in-flight connection (else its
    document listeners persist and the next click can persist a dangling edge)."""
    hub = (_WEB / "workflows.js").read_text(encoding="utf-8")
    canvas = (_WEB / "workflows-canvas.js").read_text(encoding="utf-8")
    wires = (_WEB / "workflows-wires.js").read_text(encoding="utf-8")
    assert "WF.cancelConnect = endConnect" in wires
    assert "WF.cancelConnect()" in hub  # openBlueprint
    assert "WF.cancelConnect()" in canvas  # autoArrange
