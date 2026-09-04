"""Static regression checks for Screenspace frontend sources.

Two surfaces whose failure modes are silent. The calibration slider annotation
(the hairline at the pin-derived cutoff plus the pass/fail tint on the value
readout): a renamed class or custom property, or a `background` shorthand
reintroduced on the shared slider rule, makes the mark stop rendering with no
error anywhere. And the workflow param store: params are DOM-only, so the rules
about which ids may be remembered live in the code that reads them.
"""

import re

from _frontend_source import assert_es5, read

CALIBRATION_JS = read("screenspace-calibration.js")
SCREENSPACE_CSS = read("screenspace.css")
SCREENSPACE_JS = read("screenspace.js")
TASKS_JS = read("screenspace-tasks.js")
MODEL_VIEW_JS = read("screenspace-model-view.js")
OVERLAY_JS = read("screenspace-overlay.js")
INTERACTION_JS = read("screenspace-overlay-interaction.js")
SAMPLE_EDITOR_JS = read("screenspace-sample-editor.js")


def _shared_slider_rule() -> str:
    """The declaration block shared by all three Screenspace range variants."""
    start = SCREENSPACE_CSS.index(
        '.param-control input[type="range"],\n.scene-ref-thresh {'
    )
    return SCREENSPACE_CSS[start : SCREENSPACE_CSS.index("}", start)]


def test_shared_slider_rule_uses_background_color_not_shorthand():
    """`background:` would implicitly reset background-image to none, so the
    .cal-mark hairline would only survive on source-order luck."""
    rule = _shared_slider_rule()
    assert re.search(r"^\s*background-color:", rule, re.MULTILINE), (
        "shared slider rule must set background-color"
    )
    assert not re.search(r"^\s*background:", rule, re.MULTILINE), (
        "shared slider rule must not use the `background` shorthand — it resets "
        "background-image and silently kills the .cal-mark hairline"
    )
    assert "background-repeat: no-repeat;" in rule


def test_slider_thumb_size_variable_is_defined_where_the_thumbs_read_it():
    """The thumb pseudo-elements and .cal-mark both read --slider-thumb-size;
    it has to be declared on the originating element to inherit into them."""
    assert "--slider-thumb-size: 12px;" in _shared_slider_rule()
    for pseudo in ("::-webkit-slider-thumb", "::-moz-range-thumb"):
        block_start = SCREENSPACE_CSS.index(
            f'.param-control input[type="range"]{pseudo},'
        )
        block = SCREENSPACE_CSS[block_start : SCREENSPACE_CSS.index("}", block_start)]
        assert block.count("var(--slider-thumb-size") == 2, (
            f"{pseudo} should size both width and height from --slider-thumb-size"
        )


def test_calibration_slider_classes_have_css_rules():
    """CSS toggle completeness: every class the updater adds must be styled."""
    updater_start = CALIBRATION_JS.index("function updateCalibrationSliderMarks(")
    updater = CALIBRATION_JS[
        updater_start : CALIBRATION_JS.index("\n  }", updater_start)
    ]
    toggled = set(re.findall(r"classList\.(?:add|remove)\(([^)]*)\)", updater))
    classes = {c for group in toggled for c in re.findall(r'"([\w-]+)"', group)}
    assert classes == {"cal-mark", "cal-pass", "cal-fail"}, classes
    for name in classes:
        assert f".{name}" in SCREENSPACE_CSS, f".{name} is toggled but never styled"


def test_calibration_custom_properties_round_trip():
    """Custom-property names have no toggle-completeness guard of their own, and
    a rename typo degrades silently — pin the JS↔CSS pairing both ways."""
    written = set(
        re.findall(r'(?:setProperty|removeProperty)\("(--cal-[\w-]+)"', CALIBRATION_JS)
    )
    assert written == {"--cal-mark-frac"}, written
    for name in written:
        assert f"var({name})" in SCREENSPACE_CSS, f"{name} is set in JS but never read"
    read_in_css = set(re.findall(r"var\((--cal-mark-[\w-]+)", SCREENSPACE_CSS))
    # --cal-mark-x is derived inside the rule itself, not written from JS.
    assert read_in_css - {"--cal-mark-x"} == written, read_in_css


def test_calibration_satellite_is_es5():
    assert_es5(CALIBRATION_JS, "screenspace-calibration.js")


def test_multitool_step_ids_are_excluded_from_the_param_store():
    """Step ids are positional (paramSimThresh_mt1), so a remembered value would
    land on a different step after a delete-and-reindex. _paramControls is the
    one gate — snapshot, restore and the reset buttons all go through it."""
    start = SCREENSPACE_JS.index("function _paramControls(")
    body = SCREENSPACE_JS[start : SCREENSPACE_JS.index("\n  }", start)]
    assert "_MT_STEP_ID.test(" in body, (
        "_paramControls must filter multitool step ids out of the param store"
    )
    assert re.search(r"var _MT_STEP_ID = /_mt\\d\+\$/;", SCREENSPACE_JS)


def test_task_restore_does_not_inherit_remembered_param_values():
    """Editing a task must show that task's parameters, so a param it doesn't
    carry reads as absent rather than as whatever was last typed."""
    assert "renderWorkflowParams({ defaults: true });" in TASKS_JS


def test_param_restore_fires_change_as_well_as_input():
    """Listeners that only watch `change` (the numbers operator, which toggles
    the range/target rows) would otherwise leave the UI denying the restored
    value while the scan runs it."""
    start = SCREENSPACE_JS.index("function _restoreParamValues(")
    body = SCREENSPACE_JS[start : SCREENSPACE_JS.index("\n  }", start)]
    for event in ("input", "change"):
        assert f'new Event("{event}", {{ bubbles: true }})' in body, (
            f"_restoreParamValues must dispatch a bubbling {event} event"
        )


def test_param_reset_button_classes_have_css_rules():
    for name in ("param-reset", "param-reset-icon"):
        assert f".{name}" in SCREENSPACE_CSS, f".{name} is built in JS but never styled"


def _tasks_fn(name: str) -> str:
    start = TASKS_JS.index(f"function {name}(")
    return TASKS_JS[start : TASKS_JS.index("\n  }", start)]


def test_task_stream_retires_the_poller_when_it_reconnects():
    """Both transports running at once is invisible — the UI just costs twice
    the requests. onError starts the poller; only onOpen can retire it."""
    body = _tasks_fn("startSSE")
    assert "onUnsupported: startPolling" in body, (
        "startSSE must fall back to polling where EventSource is absent"
    )
    open_handler = body[body.index("onOpen:") : body.index("onUnsupported:")]
    assert "stopPolling()" in open_handler, (
        "a reconnected stream must stop the fallback poller, or a drop-and-recover "
        "leaves SSE and the 3s poller running together"
    )


def test_frame_preload_runs_behind_the_selected_participant():
    """An unbounded preload races the one frame the user is waiting for: every
    request is a server-side extraction against a 3-slot capture pool."""
    boot = SCREENSPACE_JS[SCREENSPACE_JS.index('apiGet("api/participants")') :]
    boot = boot[: boot.index('apiGet("api/regions")')]
    assert boot.index("selectParticipant(pickId") < boot.index(
        "queueFrameZeroPreload("
    ), "the selected participant's own frame request must be issued first"
    assert re.search(r"var PRELOAD_CONCURRENCY = \d+;", SCREENSPACE_JS)


def test_frame_preload_drops_results_for_a_replaced_source_file():
    """A queued preload can now resolve seconds after selectParticipant saw a
    newer mtime and dropped the stale blob; storing it anyway would repaint the
    old frame with no error anywhere."""
    start = SCREENSPACE_JS.index("function _preloadFrameZero(")
    body = SCREENSPACE_JS[start : SCREENSPACE_JS.index("\n  }", start)]
    assert "_videoVersions[item.pid]" in body, (
        "_preloadFrameZero must compare the enqueue-time version before storing"
    )
    assert "_preloadStopped" in body, (
        "a preload that resolves after pagehide must revoke its own blob URL"
    )


def test_shape_model_view_has_meta_and_sends_capture_mask():
    assert "shape:" in MODEL_VIEW_JS[MODEL_VIEW_JS.index("var MODEL_VIEW_META") :]
    assert "ref_mask=" in MODEL_VIEW_JS
    assert "function _encodeMaskContours(" in MODEL_VIEW_JS


def test_shape_axis_labels_relabel_when_unlinked():
    """Unlinked axes relabel the base ladder Width and reveal the Height rows;
    every label variant needs a tooltip key or the hover lookup goes silent."""
    start = SCREENSPACE_JS.index("function syncAxisRows()")
    body = SCREENSPACE_JS[start : SCREENSPACE_JS.index("\n    }", start)]
    assert '"Width scale min"' in body and '"Scale min"' in body
    tips = SCREENSPACE_JS[SCREENSPACE_JS.index("    shape: {") :]
    tips = tips[: tips.index("\n    },")]
    for label in ("Width scale", "Height scale"):
        for suffix in ("min", "max", "steps"):
            assert f'"{label} {suffix}"' in tips, f"missing tooltip {label} {suffix}"


def test_shape_draw_mode_wiring():
    """Escape must exit draw mode before it starts clearing regions, and the
    draw button is opt-in — the shared capture row must not give it to Template."""
    esc = SCREENSPACE_JS.index("} else if (state.shapeDraw) {")
    assert esc < SCREENSPACE_JS.index(
        "} else if (state.pendingRegion || state.activeRegion) {"
    )
    assert 'renderRefCaptureRow(container, "Shape", { draw: true })' in SCREENSPACE_JS
    assert 'renderRefCaptureRow(container, "Template");' in SCREENSPACE_JS
    assert "SS.cancelShapeDraw = cancelShapeDraw;" in INTERACTION_JS
    assert "if (state.shapeDraw) cancelShapeDraw();" in INTERACTION_JS


def test_sample_editor_uses_blocking_modal_and_is_es5():
    assert "openBlockingModal(" in SAMPLE_EDITOR_JS
    assert "closeBlockingModal(" in SAMPLE_EDITOR_JS
    assert_es5(SAMPLE_EDITOR_JS, "screenspace-sample-editor.js")


def test_results_tab_count_element_exists():
    """Five writes target #resultCount; without the span they went nowhere."""
    assert 'qs("#resultCount")' in read("screenspace-results.js")
    assert 'id="resultCount"' in read("screenspace.html")


def test_model_view_previews_focused_multitool_step():
    """The server previews plain tools; multitool must send the focused step."""
    body = MODEL_VIEW_JS[MODEL_VIEW_JS.index("function _doRefreshModelView") :]
    assert "_focusStep()" in body
    assert 'sfx = "_mt" + stepIdx' in body
    assert "_collectPreviewParams(tool, sfx)" in body
    assert "(state.multitoolSteps || [])[0]" not in MODEL_VIEW_JS


def test_model_view_overlay_uses_preview_region():
    """The overlay must use the region that produced its preview image."""
    assert "state.overlayImageRegion = overlayRegion;" in MODEL_VIEW_JS
    start = OVERLAY_JS.index("// Model-view overlay")
    body = OVERLAY_JS[start : OVERLAY_JS.index("ctx.globalAlpha = 1.0;", start)]
    assert "state.overlayImageTool === SS._previewToolKey()" in body
    assert "var oRegion = state.overlayImageRegion;" in body


def test_multitool_branch_refreshes_model_view():
    """The multitool branch returns early; without its own refresh the preview went stale."""
    start = SCREENSPACE_JS.index(
        'if (type === "multitool") {\n      SS.renderMultitoolParams'
    )
    body = SCREENSPACE_JS[start : SCREENSPACE_JS.index("return;", start)]
    assert "refreshModelView();" in body
    assert "_updateOverlayUi();" in body


def test_multitool_step_edits_refresh_model_view():
    """Step rows skip addParamRow, so the step list needs its own input listener."""
    js = read("screenspace-multitool-params.js")
    assert 'stepsDiv.addEventListener("input", onStepEdit)' in js
    assert 'stepsDiv.addEventListener("change", onStepEdit)' in js
    assert "SS.setMultitoolFocus(" in js


def test_multitool_focus_classes_have_css_rules():
    for name in (
        ".multitool-step.is-selected",
        ".multitool-step-chevron",
        ".preview-section-focus",
        ".cal-track.is-focus",
    ):
        assert name in SCREENSPACE_CSS, f"{name} is set in JS but never styled"
    assert 'id="modelViewFocus"' in read("screenspace.html")
