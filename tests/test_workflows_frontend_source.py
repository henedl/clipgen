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
    for name in (
        "workflows-nodes.js",
        "workflows-canvas.js",
        "workflows-wires.js",
        "workflows-runs.js",
        "workflows-stashes.js",
        "workflows-validate.js",
    ):
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
    """Wire layer: the SVG element, the connector satellite, and adapter-aware
    type validation (exact match OR a served ADAPTERS coercion) all ship."""
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    assert 'id="wfWires"' in html
    assert 'id="wfWireDelete"' in html
    assert "workflows-wires.js" in html
    wires = (_WEB / "workflows-wires.js").read_text(encoding="utf-8")
    # canConnect keeps the exact-match clause AND consults the served adapter
    # table (P1), so the UI accepts the same coercions the runner applies.
    assert "outType === inType" in wires
    assert "state.adapters" in wires
    # Adapter-bridged wires get a distinct ("coerced") render class.
    assert "wf-wire-coerced" in wires
    # Endpoints come from live node position + cached port offsets.
    assert "portWorldPos" in wires


def test_css_defines_wire_and_param_styles():
    css = WORKFLOWS_CSS.read_text(encoding="utf-8")
    # Wire layer + connector styling.
    assert ".wf-wires" in css
    assert ".wf-wire" in css
    assert ".wf-wire-coerced" in css  # dashed cue for adapter-bridged wires (P1)
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


def test_run_panel_satellite_present_and_wired():
    """M4 run engine: the Run/Stop controls, the runs satellite (loaded last),
    and the hub<->satellite run interface all ship."""
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    assert 'id="wfRunBtn"' in html
    assert 'id="wfStopBtn"' in html
    assert 'id="wfRuns"' in html
    # Runs satellite loads after the wires satellite (load-order contract).
    assert html.index('src="workflows-wires.js"') < html.index(
        'src="workflows-runs.js"'
    )

    src = _workflows_js()
    # Hub-owned: flushSave is awaited before a run so the server has the latest
    # blueprint; the satellite owns the run lifecycle + rendering.
    for fn in (
        "WF.flushSave",
        "WF.initRuns",
        "WF.startRun",
        "WF.stopRun",
        "WF.refreshRuns",
        "WF.renderRuns",
    ):
        assert fn in src, fn
    # SSE stream with a polling fallback (mirrors screenspace-tasks).
    runs = (_WEB / "workflows-runs.js").read_text(encoding="utf-8")
    assert "EventSource" in runs
    assert "createPoller" in runs
    assert "api/runs" in runs


def test_batch_via_all_participants_option():
    """P3 whole-study fan-out: triggered by a Video Source's "All participants"
    dropdown option (no separate header button) — Run delegates to a batch, with
    its own SSE/poll transport and batch card style."""
    src = _workflows_js()
    # Frontend-only sentinel on the hub; the nodes satellite offers it as an option
    # and the runs satellite branches Run on it (no duplicated string literal).
    assert "WF.ALL_PARTICIPANTS" in src
    nodes = (_WEB / "workflows-nodes.js").read_text(encoding="utf-8")
    assert "All participants" in nodes
    # The single Run button fans out when a source is set to "All".
    runs = (_WEB / "workflows-runs.js").read_text(encoding="utf-8")
    assert "blueprintWantsBatch" in runs
    assert "batches:" in src  # state.batches on the hub
    assert "api/batches" in runs
    assert "subscribeBatch" in runs

    # The old standalone control is gone.
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    assert "wfRunAllBtn" not in html

    css = WORKFLOWS_CSS.read_text(encoding="utf-8")
    assert ".wf-batch-card" in css


def test_stash_satellite_present_and_wired():
    """M5 stashes + P4 built-in recipes: the sidebar list, the "Stash selection"
    toolbar control, the satellite (loaded last), and the hub<->satellite stash
    interface all ship; instantiation rides the canvas drop path via a new MIME."""
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    assert 'id="wfStashList"' in html
    assert 'id="wfSaveStash"' in html
    # Stashes satellite loads after the runs satellite (load-order contract).
    assert html.index('src="workflows-runs.js"') < html.index(
        'src="workflows-stashes.js"'
    )

    src = _workflows_js()
    # Hub-owned: state.stashes lives on the hub; the satellite is init'd + loaded
    # through guarded delegators.
    assert "stashes:" in src  # state.stashes on the hub
    assert "WF.initStashes" in src
    assert "WF.loadStashes" in src
    # Satellite-owned stash lifecycle attached back onto WF.
    for fn in (
        "WF.renderStashPalette",
        "WF.saveSelectionAsStash",
        "WF.instantiateStash",
        "WF.renameStash",
        "WF.deleteStash",
        "WF.syncStashButton",
    ):
        assert fn in src, fn
    # Server does CRUD only; the client GETs the list and instantiates locally.
    stashes = (_WEB / "workflows-stashes.js").read_text(encoding="utf-8")
    assert "api/stashes" in stashes
    # The canvas reaches the stash instantiator via the new drag MIME.
    canvas = (_WEB / "workflows-canvas.js").read_text(encoding="utf-8")
    assert "application/x-wf-stash" in canvas
    assert "WF.instantiateStash(" in canvas
    # randomId is published for the satellite's id remapping on instantiate.
    assert "WF.randomId" in canvas


def test_css_defines_stash_styles():
    css = WORKFLOWS_CSS.read_text(encoding="utf-8")
    assert ".wf-stash-list" in css
    assert ".wf-stash-item" in css
    assert ".wf-stash-builtin" in css  # read-only built-in variant


def test_universal_control_port_for_gate_wiring():
    """A Gate's `control` output wires into any node via a universal optional
    `__gate__` input (exact-match) — the M4 control-edge gating decision."""
    src = _workflows_js()
    assert "__gate__" in src
    assert '"control"' in src or "'control'" in src


def test_node_descriptions_use_data_tooltip_singleton():
    """Palette rows and the on-card `?` glyph surface the catalog description
    through the [data-tooltip] singleton (utils.clipgenInitDataTooltips), not
    native `title` — native tooltips don't render on draggable=true palette
    rows and are unstyled. Guards against regressing to `.title =`."""
    src = _workflows_js()
    # Palette row + help glyph both wire description via data-tooltip.
    assert 'setAttribute("data-tooltip", node.description)' in src
    assert 'setAttribute("data-tooltip", type.description)' in src
    # The help glyph itself ships (mask-icon span carrying the tooltip).
    assert "wf-node-help" in src
    css = WORKFLOWS_CSS.read_text(encoding="utf-8")
    assert ".wf-node-help" in css


def test_css_defines_run_panel_and_node_status_styles():
    css = WORKFLOWS_CSS.read_text(encoding="utf-8")
    assert ".wf-run-card" in css
    assert ".wf-run-status-running" in css
    assert ".wf-node.run-completed" in css
    assert ".wf-node.run-failed" in css
    assert ".wf-node-progress" in css


def test_validation_satellite_present_and_wired():
    """P5 pre-run validation: the Issues panel, the satellite (loaded last), and
    the hub<->satellite validation interface ship; the hub recomputes on every
    edit (scheduleSave) and gates Run on errors."""
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    assert 'id="wfValidation"' in html
    assert "workflows-validate.js" in html
    # Validation satellite loads after the stashes satellite (load-order contract).
    assert html.index('src="workflows-stashes.js"') < html.index(
        'src="workflows-validate.js"'
    )

    src = _workflows_js()
    # Hub publishes nothing new, but the satellite attaches its interface onto WF.
    for fn in ("WF.nodeIssues", "WF.graphHasCycle", "WF.refreshValidation"):
        assert fn in src, fn
    # state.validation lives on the hub; recomputed on every edit (not debounced)
    # and once on blueprint load.
    assert "validation:" in src
    hub = (_WEB / "workflows.js").read_text(encoding="utf-8")
    assert "WF.refreshValidation()" in hub  # scheduleSave + openBlueprint
    # The runs satellite re-gates Run from validation errors; the nodes satellite
    # shares the per-node cue via WF.nodeIssues.
    runs = (_WEB / "workflows-runs.js").read_text(encoding="utf-8")
    assert "WF.syncRunButton" in runs
    assert "state.validation" in runs
    nodes = (_WEB / "workflows-nodes.js").read_text(encoding="utf-8")
    assert "WF.nodeIssues" in nodes
    # Clicking an Issues row reveals its node (published by the canvas satellite).
    canvas = (_WEB / "workflows-canvas.js").read_text(encoding="utf-8")
    assert "WF.focusNode" in canvas


def test_lazy_node_results_and_rerun_present():
    """P5 run-history UX: node rows lazily fetch their stored result sidecar on
    expand, and terminal run cards offer a Re-run."""
    runs = (_WEB / "workflows-runs.js").read_text(encoding="utf-8")
    # hasResult drives the expandable row; the result is fetched from the sidecar
    # endpoint and cached on the run object.
    assert "hasResult" in runs
    assert "/nodes/" in runs and "/result" in runs
    assert "_nodeResults" in runs
    assert "wf-run-rerun" in runs
    css = WORKFLOWS_CSS.read_text(encoding="utf-8")
    assert ".wf-result-panel" in css
    assert ".wf-validation" in css
    assert ".wf-issue-error" in css
    assert ".wf-issue-warning" in css


def test_watch_dir_trigger_toggle_and_badge():
    """P6 watch-dir triggers: the toolbar toggle, its hub-owned PUT + single-active
    mirror, the validate satellite re-gate, and the run-history triggered badge."""
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    assert 'id="wfTriggerBtn"' in html

    hub = (_WEB / "workflows.js").read_text(encoding="utf-8")
    # Hub owns the toggle (touches state.blueprints + the dedicated trigger PUT).
    assert "/trigger" in hub
    assert "function toggleTrigger" in hub
    assert "WF.syncTriggerButton" in hub  # published for the validate satellite
    assert '"#wfTriggerBtn"' in hub  # gated alongside the other toolbar controls

    # The validate satellite re-gates the toggle on every edit (can't arm errors).
    validate = (_WEB / "workflows-validate.js").read_text(encoding="utf-8")
    assert "WF.syncTriggerButton" in validate

    # The run history tags auto-launched runs (rides the snapshot's `triggered`).
    runs = (_WEB / "workflows-runs.js").read_text(encoding="utf-8")
    assert "run.triggered" in runs
    assert "wf-run-triggered" in runs
    # A triggered run isn't started by this client, so the panel must discover it:
    # a low-frequency idle poll refreshes the run list so it surfaces live.
    assert "discoverTick" in runs
    assert "runsChanged" in runs

    css = WORKFLOWS_CSS.read_text(encoding="utf-8")
    assert ".wf-trigger-icon" in css
    assert ".wf-trigger-armed" in css
    assert ".wf-run-triggered" in css


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


def test_palette_search_filters_nodes():
    """The palette has a search box that filters nodes by label/description/category."""
    hub = (_WEB / "workflows.js").read_text(encoding="utf-8")
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    assert 'id="wfPaletteSearch"' in html
    assert 'autocomplete="off"' in html  # text input must not browser-autofill
    assert "paletteNodeMatches" in hub
    # The search input re-renders the palette on every keystroke.
    assert 'qs("#wfPaletteSearch")' in hub


def test_blueprint_import_export():
    """Blueprint JSON export/import live in the TopNav Quick Actions menu."""
    hub = (_WEB / "workflows.js").read_text(encoding="utf-8")
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    assert 'id="wfImportFile"' in html  # hidden input the Import action triggers
    assert "function exportBlueprint" in hub
    assert "function importBlueprint" in hub
    # Registered as TopNav quick actions, not toolbar buttons.
    assert "ClipgenTopNav.setQuickActions" in hub
    assert "Export blueprint JSON" in hub
    assert "Import blueprint JSON" in hub
    # Import reuses the existing create endpoint (no new server route).
    assert 'apiPost("api/blueprints"' in hub
    # Export reuses the canonical Blob-download idiom.
    assert "URL.createObjectURL" in hub


def test_copy_paste_duplicate_nodes():
    """Canvas supports Ctrl/Cmd+C/V/D, reusing the stash id-remap for paste."""
    canvas = (_WEB / "workflows-canvas.js").read_text(encoding="utf-8")
    stashes = (_WEB / "workflows-stashes.js").read_text(encoding="utf-8")
    assert "function copySelection" in canvas
    assert "function pasteClipboard" in canvas
    assert "function duplicateSelection" in canvas
    # Clipboard shortcuts gate on the cmd/ctrl modifier and skip text fields.
    assert "e.metaKey || e.ctrlKey" in canvas
    # Paste delegates to the stashes satellite's published id-remap (late-bound).
    assert "WF.instantiateSubgraph" in canvas
    assert "WF.instantiateSubgraph = instantiateSubgraph" in stashes


def test_node_mute_toggle():
    """Nodes have a mute toggle; the runner-facing flag is `disabled`, validation
    skips muted nodes, and the card dims via .wf-node-muted."""
    nodes = (_WEB / "workflows-nodes.js").read_text(encoding="utf-8")
    validate = (_WEB / "workflows-validate.js").read_text(encoding="utf-8")
    css = WORKFLOWS_CSS.read_text(encoding="utf-8")
    assert "wf-node-mute" in nodes
    assert "node.disabled = !node.disabled" in nodes
    assert "wf-node-muted" in nodes
    assert ".wf-node-muted" in css
    # A muted node must not block Run.
    assert "if (node.disabled) return { errors: [], warnings: [] }" in validate


def test_validation_warns_on_node_with_no_inputs_wired():
    """A node whose data inputs are all optional + unwired (e.g. make_clips,
    measure) warns instead of passing validation and running empty. Suppressed
    when the clearer orphan "not connected" message fires instead."""
    validate = (_WEB / "workflows-validate.js").read_text(encoding="utf-8")
    assert "Wire at least one input" in validate
    assert "willShowOrphan" in validate


def test_collection_palette_grouped_by_operation():
    """The Collection category renders as collapsible operation sub-groups."""
    hub = (_WEB / "workflows.js").read_text(encoding="utf-8")
    css = WORKFLOWS_CSS.read_text(encoding="utf-8")
    assert "appendCollectionGroups" in hub
    assert 'cat === "Collection"' in hub
    # Sub-groups use native <details>/<summary> for collapse (no JS toggle).
    assert 'el("details", "wf-palette-subgroup")' in hub
    assert ".wf-palette-subgroup-label" in css


def test_undo_redo_history():
    """The hub keeps a coalesced snapshot history; canvas binds Ctrl/Cmd+Z / +Y."""
    hub = (_WEB / "workflows.js").read_text(encoding="utf-8")
    canvas = (_WEB / "workflows-canvas.js").read_text(encoding="utf-8")
    assert "function undo" in hub and "function redo" in hub
    assert "WF.undo = undo" in hub and "WF.redo = redo" in hub
    # Capture rides the scheduleSave chokepoint and resets on blueprint switch.
    assert "captureHistory" in hub
    assert "resetHistory()" in hub  # openBlueprint
    # Canvas calls the hub's history late-bound.
    assert "WF.undo" in canvas and "WF.redo" in canvas


def test_unified_detect_node():
    """A single Detect node with a detector dropdown; hidden ss_* stay as specs."""
    hub = (_WEB / "workflows.js").read_text(encoding="utf-8")
    nodes = (_WEB / "workflows-nodes.js").read_text(encoding="utf-8")
    validate = (_WEB / "workflows-validate.js").read_text(encoding="utf-8")
    # Palette skips hidden catalog nodes.
    assert "if (node.hidden) return;" in hub
    # Detect editor swaps the param set per detector, derived from the catalog.
    assert "function buildDetectEditor" in nodes
    assert "function detectTypes" in nodes
    assert 'node.type === "detect"' in nodes
    # Detect's per-detector required params still gate Run.
    assert 'node.type === "detect"' in validate


def test_run_to_selected_node():
    """'Run to here' lives in the Run split-button menu and starts a partial run."""
    hub = (_WEB / "workflows.js").read_text(encoding="utf-8")
    runs = (_WEB / "workflows-runs.js").read_text(encoding="utf-8")
    html = WORKFLOWS_HTML.read_text(encoding="utf-8")
    css = WORKFLOWS_CSS.read_text(encoding="utf-8")
    # Split-button: primary Run + caret menu with a Run-to-here item.
    assert 'id="wfRunMenuBtn"' in html
    assert 'id="wfRunToItem"' in html
    assert ".wf-run-menu" in css
    assert "function initRunMenu" in hub
    assert "WF.startRun(sel[0])" in hub
    # The run-start passes the target through; the item needs exactly one node.
    assert "body.targetNodeId = targetNodeId" in runs
    assert "state.selection.length === 1" in runs


def test_theme_toggle_icons_styled():
    """The TopNav renders #themeToggle, but each page must supply the sun/moon
    icon visuals + the position:relative anchor (missing → blank toggle)."""
    css = WORKFLOWS_CSS.read_text(encoding="utf-8")
    assert ".theme-icon-sun" in css
    assert ".theme-icon-moon" in css
    assert '#themeToggle[data-theme="dark"] .theme-icon-moon' in css
    # The absolutely-positioned icons need a positioned button to anchor to.
    assert "#themeToggle {" in css
