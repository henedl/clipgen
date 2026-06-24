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
