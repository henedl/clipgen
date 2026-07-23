/* clipgen Screenspace multitool-params satellite — screenspace-multitool-params.js
 *
 * Renders the multitool step list (per-step parameter panels, drag reorder,
 * drop-to-import from the task queue). Carved out of screenspace.js to shrink
 * the page script; loaded after it. Reads the hub's shared state + helpers via
 * window.ClipgenScreenspace (set up in screenspace.js) and publishes back the
 * two entry points the hub calls. Function bodies are unchanged from when they
 * lived inline in screenspace.js — the locals below stand in for the closure.
 */
(function () {
  "use strict";

  var SS = window.ClipgenScreenspace;
  var state = SS.state;
  var activeRegionRef = SS.activeRegionRef,
    allAvailableRegionRefs = SS.allAvailableRegionRefs,
    availableRegionRefByKey = SS.availableRegionRefByKey,
    normalizeRegionRef = SS.normalizeRegionRef,
    regionRefKey = SS.regionRefKey,
    regionRefLabel = SS.regionRefLabel,
    regionRefPayload = SS.regionRefPayload,
    taskTypeColor = SS.taskTypeColor,
    buildTypeIcon = SS.buildTypeIcon,
    iconSpan = SS.iconSpan,
    buildNormalizeControl = SS.buildNormalizeControl,
    buildColorModeControl = SS.buildColorModeControl,
    _colorMode = SS._colorMode,
    _updateMinAreaReadout = SS._updateMinAreaReadout,
    activatePipette = SS.activatePipette,
    deactivatePipette = SS.deactivatePipette,
    renderWorkflowParams = SS.renderWorkflowParams,
    updateRunButton = SS.updateRunButton,
    findTask = SS.findTask,
    refreshCalibration = SS.refreshCalibration;

  // ---- Multitool step list (drag reorder + drop-to-import from task queue) ----

  var MULTITOOL_ALLOWED_TYPES = [
    { value: "color", label: "Color" },
    { value: "change", label: "Change" },
    { value: "similarity", label: "Similarity" },
    { value: "text", label: "Text" },
    { value: "numbers", label: "Numbers" },
    { value: "template", label: "Template" },
    { value: "flow", label: "Flow" },
    { value: "scene", label: "Scene" },
    { value: "inactivity", label: "Inactivity" },
  ];

  function renderMultitoolStepBody(body, stepType, idx) {
    var sfx = "_mt" + idx;

    // Per-step region selector
    var regionRow = el("div", "param-row");
    regionRow.appendChild(el("span", "param-label", "Region"));
    var regionCtrl = el("div", "param-control");
    var regionSel = document.createElement("select");
    regionSel.className = "multitool-step-region-select";
    regionSel.id = "paramStepRegion" + sfx;
    var regionRefs = allAvailableRegionRefs();
    var selectedRef = normalizeRegionRef(state.multitoolSteps[idx].region_ref)
      || (state.multitoolSteps[idx].region ? activeRegionRef(state.multitoolSteps[idx].region) : null);
    var selectedKey = selectedRef ? regionRefKey(selectedRef) : "";
    regionRefs.forEach(function (ref) {
      var opt = document.createElement("option");
      opt.value = regionRefKey(ref);
      opt.textContent = regionRefLabel(ref);
      if (opt.value === selectedKey) opt.selected = true;
      regionSel.appendChild(opt);
    });
    if (!state.multitoolSteps[idx].region && regionRefs.length > 0) {
      var defaultRegionRef = state.runRegions.length > 0 ? normalizeRegionRef(state.runRegions[0]) : null;
      defaultRegionRef = defaultRegionRef || (state.activeRegion ? activeRegionRef(state.activeRegion) : null) || regionRefs[0];
      state.multitoolSteps[idx].region = defaultRegionRef.name;
      state.multitoolSteps[idx].region_ref = regionRefPayload(defaultRegionRef);
      regionSel.value = regionRefKey(defaultRegionRef);
    }
    (function (capturedIdx) {
      regionSel.addEventListener("change", function () {
        var ref = availableRegionRefByKey(regionSel.value);
        state.multitoolSteps[capturedIdx].region = ref ? ref.name : "";
        state.multitoolSteps[capturedIdx].region_ref = ref ? regionRefPayload(ref) : null;
        // Region drives the color presence min-area pixel estimate.
        _updateMinAreaReadout("_mt" + capturedIdx);
      });
    })(idx);
    regionCtrl.appendChild(regionSel);
    regionRow.appendChild(regionCtrl);
    body.appendChild(regionRow);

    var renderer = MULTITOOL_PARAM_RENDERERS[stepType];
    if (renderer) renderer(body, idx, sfx);
    // _initial is consumed at first render; drop it so adding/removing other steps later
    // doesn't overwrite the user's in-progress edits with the original saved values.
    delete state.multitoolSteps[idx]._initial;
  }

  function _mtRenderColor(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    var initColor = init.target_color || {};
    var initH = numberOrDefault(initColor.h, 0);
    var initS = numberOrDefault(initColor.s, 0);
    var initV = numberOrDefault(initColor.v, 0);
    var initTol = init.tolerance ? Math.round(init.tolerance.h * 100 / 90) : 30;
    var row1 = el("div", "param-row");
    row1.appendChild(el("span", "param-label", "Hex color"));
    var ctrl1 = el("div", "param-control");
    var hexIn = document.createElement("input");
    hexIn.type = "text"; hexIn.id = "paramColorHex" + sfx; hexIn.autocomplete = "off";
    hexIn.className = "color-hex-input"; hexIn.placeholder = "#000000"; hexIn.maxLength = 7;
    hexIn.style.width = "5.5rem";
    var initRgb = hsvToRgb(initH, initS, initV);
    hexIn.value = rgbToHex(initRgb.r, initRgb.g, initRgb.b);
    ctrl1.appendChild(hexIn);
    var pipBtn = el("button", "btn btn-small btn-pipette");
    pipBtn.title = "Pick color from frame";
    pipBtn.appendChild(buildTypeIcon("color"));
    pipBtn.addEventListener("click", function () {
      if (state.pipetteActive) { deactivatePipette(); return; }
      state._mtPipetteStep = idx;
      activatePipette();
    });
    ctrl1.appendChild(pipBtn);
    var hH = document.createElement("input"); hH.type = "hidden"; hH.id = "paramColorH" + sfx; hH.value = String(initH);
    var hS = document.createElement("input"); hS.type = "hidden"; hS.id = "paramColorS" + sfx; hS.value = String(initS);
    var hV = document.createElement("input"); hV.type = "hidden"; hV.id = "paramColorV" + sfx; hV.value = String(initV);
    ctrl1.appendChild(hH); ctrl1.appendChild(hS); ctrl1.appendChild(hV);
    row1.appendChild(ctrl1);
    body.appendChild(row1);
    hexIn.addEventListener("input", function () {
      var hex = hexIn.value.replace("#", "");
      if (hex.length === 6) {
        var r = parseInt(hex.substring(0, 2), 16) || 0;
        var g = parseInt(hex.substring(2, 4), 16) || 0;
        var b = parseInt(hex.substring(4, 6), 16) || 0;
        // OpenCV HSV: H 0-180, S 0-255, V 0-255
        var rr = r / 255, gg = g / 255, bb = b / 255;
        var mx = Math.max(rr, gg, bb), mn = Math.min(rr, gg, bb);
        var d = mx - mn, h = 0, s = mx === 0 ? 0 : d / mx, v = mx;
        if (d !== 0) {
          if (mx === rr) h = ((gg - bb) / d + (gg < bb ? 6 : 0)) / 6;
          else if (mx === gg) h = ((bb - rr) / d + 2) / 6;
          else h = ((rr - gg) / d + 4) / 6;
        }
        hH.value = Math.round(h * 180);
        hS.value = Math.round(s * 255);
        hV.value = Math.round(v * 255);
      }
    });
    _mtAddNumberRow(body, "Tolerance", "paramColorTol" + sfx, 0, 100, initTol, 1);

    // Match mode (average vs presence) + presence-only min-area row, mirroring
    // the single-tool color panel.
    var initMode = _colorMode(init.color_mode);
    // A restored presence step with no min_coverage means "any presence" (0%);
    // a fresh/average step defaults to 1% for when the user switches to presence.
    var initMinArea = initMode === "presence"
      ? (init.min_coverage != null ? init.min_coverage * 100 : 0)
      : 1;
    var modeRow = el("div", "param-row");
    modeRow.appendChild(el("span", "param-label", "Mode"));
    var modeCtrl = el("div", "param-control");
    var minAreaRow = el("div", "param-row" + (initMode === "presence" ? "" : " hidden"));
    modeCtrl.appendChild(
      buildColorModeControl("paramColorMode" + sfx, initMode, true, function (mode) {
        minAreaRow.classList.toggle("hidden", mode !== "presence");
      })
    );
    modeRow.appendChild(modeCtrl);
    body.appendChild(modeRow);
    minAreaRow.appendChild(el("span", "param-label", "Min area %"));
    var minAreaCtrl = el("div", "param-control");
    var minAreaInput = numberInput("paramColorMinArea" + sfx, 0, 100, initMinArea, 1);
    minAreaCtrl.appendChild(minAreaInput);
    var minAreaVal = el("span", "param-value param-value--minarea");
    minAreaVal.id = "paramColorMinAreaVal" + sfx;
    minAreaCtrl.appendChild(minAreaVal);
    minAreaRow.appendChild(minAreaCtrl);
    body.appendChild(minAreaRow);
    minAreaInput.addEventListener("input", function () {
      _updateMinAreaReadout(sfx);
    });
    _updateMinAreaReadout(sfx);
  }

  function _mtRenderChange(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    _mtAddNumberRow(body, "Threshold", "paramChangeThresh" + sfx, 0.01, 0.50, numberOrDefault(init.threshold, 0.03), 0.01);
    _mtAddNumberRow(body, "Noise", "paramChangeNoise" + sfx, 0, 100, intOrDefault(init.noise_threshold, 30), 1);
  }

  function _mtRenderSimilarity(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    _mtAddCaptureRefRow(body, idx, "Reference", "paramSimRef" + sfx);
    _mtAddNumberRow(body, "Threshold", "paramSimThresh" + sfx, 0.50, 1.00, numberOrDefault(init.threshold, 0.90), 0.01);
  }

  function _mtRenderText(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    var r1 = el("div", "param-row");
    r1.appendChild(el("span", "param-label", "Search"));
    var c1 = el("div", "param-control");
    var searchIn = document.createElement("input");
    searchIn.type = "text"; searchIn.id = "paramTextSearch" + sfx; searchIn.autocomplete = "off";
    searchIn.placeholder = "Search text...";
    searchIn.value = init.search_string || "";
    searchIn.style.flex = "1"; searchIn.style.fontSize = "var(--text-xs)";
    searchIn.style.padding = "var(--space-1)";
    searchIn.style.border = "1px solid var(--color-border)";
    searchIn.style.borderRadius = "var(--radius-sm)";
    searchIn.style.background = "var(--color-bg)";
    searchIn.style.color = "var(--color-text)";
    c1.appendChild(searchIn);
    r1.appendChild(c1);
    body.appendChild(r1);
    _mtAddNumberRow(body, "Fuzzy", "paramTextFuzzy" + sfx, 0.50, 1.00, numberOrDefault(init.fuzzy_threshold, CLIPGEN_CONFIG.screenspaceOcrFuzzyThreshold), 0.01);
    _mtAddNumberRow(body, "Min OCR", "paramTextOcrConf" + sfx, 0.00, 1.00, numberOrDefault(init.ocr_confidence_threshold, CLIPGEN_CONFIG.screenspaceOcrMinConfidence), 0.01);
    _mtAddCheckboxRow(body, "Enhance ROI", "paramTextOcrPreprocess" + sfx, init.ocr_preprocess);
    var nr = el("div", "param-row");
    nr.appendChild(el("span", "param-label", "Normalize"));
    var nc = el("div", "param-control");
    nc.appendChild(buildNormalizeControl("paramTextOcrNormalize" + sfx, init.ocr_normalize, true));
    nr.appendChild(nc);
    body.appendChild(nr);
  }

  function _mtRenderNumbers(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    var r1 = el("div", "param-row");
    r1.appendChild(el("span", "param-label", "Operator"));
    var c1 = el("div", "param-control");
    var sel = document.createElement("select");
    sel.id = "paramNumOperator" + sfx;
    sel.style.fontSize = "var(--text-xs)";
    [["gt",">"],["lt","<"],["eq","="],["gte","≥"],["lte","≤"],["range","range"]].forEach(function (pair) {
      var opt = document.createElement("option"); opt.value = pair[0]; opt.textContent = pair[1];
      sel.appendChild(opt);
    });
    sel.value = init.operator || "gt";
    c1.appendChild(sel);
    r1.appendChild(c1);
    body.appendChild(r1);
    var r2 = el("div", "param-row");
    r2.appendChild(el("span", "param-label", "Target"));
    var c2 = el("div", "param-control");
    var targetIn = document.createElement("input");
    targetIn.type = "number"; targetIn.id = "paramNumTarget" + sfx; targetIn.style.width = "4rem";
    targetIn.style.fontSize = "var(--text-xs)";
    targetIn.value = numberOrDefault(init.target_value, 0);
    c2.appendChild(targetIn);
    r2.appendChild(c2);
    body.appendChild(r2);
    _mtAddNumberRow(body, "Min OCR", "paramNumOcrConf" + sfx, 0.00, 1.00, numberOrDefault(init.ocr_confidence_threshold, CLIPGEN_CONFIG.screenspaceOcrMinConfidence), 0.01);
    _mtAddCheckboxRow(body, "Enhance ROI", "paramNumOcrPreprocess" + sfx, init.ocr_preprocess);
    _mtAddCheckboxRow(body, "Integers", "paramNumIntegersOnly" + sfx, init.integers_only);
  }

  function _mtRenderTemplate(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    var step = state.multitoolSteps[idx];

    // Capture-or-upload row, mirroring the single-tool template workflow but
    // scoped per step (state on the step object, not the global uploadedTemplate).
    var row = el("div", "param-row");
    row.appendChild(el("span", "param-label", "Template"));
    var ctrl = el("div", "param-control");

    var info = el("span", "param-value template-upload-info");
    function renderInfo() {
      info.innerHTML = "";
      if (step._upload) {
        var thumb = document.createElement("img");
        thumb.decoding = "async";
        thumb.src = "data:image/png;base64," + step._upload.data;
        thumb.alt = "Uploaded template";
        thumb.title = step._upload.name;
        info.appendChild(thumb);
        var clearBtn = el("button", "btn btn-small", "×");
        clearBtn.addEventListener("click", function () {
          step._upload = null;
          renderInfo();
          refreshCalibration({ debounce: true });
        });
        info.appendChild(clearBtn);
      } else if (step._refTs !== undefined) {
        info.appendChild(el("span", null, formatTime(step._refTs, { decimals: 1 })));
      } else {
        info.appendChild(el("span", null, "—"));
      }
    }

    var capBtn = el("button", "btn btn-small ss-template-icon-btn ss-template-icon-btn--capture");
    capBtn.setAttribute("type", "button");
    capBtn.title = "Capture Frame";
    capBtn.setAttribute("aria-label", "Capture Frame");
    capBtn.appendChild(el("span", "ss-template-icon-btn__glyph"));
    capBtn.addEventListener("click", function () {
      step._refTs = state.currentTimestamp;
      step._upload = null;
      renderInfo();
      refreshCalibration({ debounce: true });
    });
    ctrl.appendChild(capBtn);

    var fileInput = document.createElement("input");
    fileInput.type = "file";
    fileInput.accept = "image/png";
    fileInput.style.display = "none";
    fileInput.addEventListener("change", function () {
      var file = fileInput.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function (e) {
        step._upload = { name: file.name, data: e.target.result.split(",")[1] };
        step._refTs = undefined;
        renderInfo();
        refreshCalibration({ debounce: true });
        showToast("Template loaded");
      };
      reader.readAsDataURL(file);
    });
    var uploadBtn = el("button", "btn btn-small ss-template-icon-btn ss-template-icon-btn--upload");
    uploadBtn.setAttribute("type", "button");
    uploadBtn.title = "Upload PNG";
    uploadBtn.setAttribute("aria-label", "Upload PNG");
    uploadBtn.appendChild(el("span", "ss-template-icon-btn__glyph"));
    uploadBtn.addEventListener("click", function () { fileInput.click(); });
    ctrl.appendChild(uploadBtn);
    ctrl.appendChild(fileInput);

    renderInfo();
    ctrl.appendChild(info);
    row.appendChild(ctrl);
    body.appendChild(row);

    _mtAddNumberRow(body, "Threshold", "paramTemplateThresh" + sfx, 0.50, 1.00, numberOrDefault(init.threshold, 0.70), 0.01);
    _mtAddNumberRow(body, "Scale %", "paramTemplateScale" + sfx, 25, 200, init.template_scale != null ? Math.round(init.template_scale * 100) : 100, 5);
  }

  function _mtRenderFlow(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    _mtAddNumberRow(body, "Magnitude", "paramFlowMag" + sfx, 0.5, 20, numberOrDefault(init.magnitude_threshold, 2.0), 0.5);
  }

  function _mtRenderScene(body, idx, sfx) {
    var sceneList = el("div", "scene-reference-list");
    sceneList.id = "mtSceneList" + sfx;
    if (!state.multitoolSteps[idx]._scenes) state.multitoolSteps[idx]._scenes = [];
    function renderMtScenes() {
      sceneList.innerHTML = "";
      state.multitoolSteps[idx]._scenes.forEach(function (ref, refIdx) {
        if (ref.threshold === undefined) ref.threshold = 0.75;
        var item = el("div", "scene-ref-item");
        item.appendChild(el("span", "scene-ref-name", ref.name));
        item.appendChild(el("span", "param-value", formatTime(ref.timestamp, { decimals: 1 })));
        var threshSlider = document.createElement("input");
        threshSlider.type = "range";
        threshSlider.min = "0.50"; threshSlider.max = "1.00"; threshSlider.step = "0.01";
        threshSlider.value = String(ref.threshold);
        threshSlider.className = "scene-ref-thresh";
        var threshVal = el("span", "param-value", String(ref.threshold));
        threshSlider.addEventListener("input", (function (ri) {
          return function () {
            state.multitoolSteps[idx]._scenes[ri].threshold = parseFloat(threshSlider.value);
            threshVal.textContent = threshSlider.value;
          };
        })(refIdx));
        item.appendChild(threshSlider);
        item.appendChild(threshVal);
        var rmBtn = el("button", "btn btn-small", "\u00d7");
        rmBtn.addEventListener("click", (function (ri) {
          return function () {
            state.multitoolSteps[idx]._scenes.splice(ri, 1);
            renderMtScenes();
            refreshCalibration({ debounce: true });
          };
        })(refIdx));
        item.appendChild(rmBtn);
        sceneList.appendChild(item);
      });
    }
    renderMtScenes();
    body.appendChild(sceneList);
    var addScRow = el("div", "param-row");
    addScRow.appendChild(el("span", "param-label", "Add Scene"));
    var addScCtrl = el("div", "param-control");
    var scNameInp = textInput("paramSceneName" + sfx, "e.g. menu, gameplay");
    addScCtrl.appendChild(scNameInp);
    var scCapBtn = el("button", "btn btn-small", "Capture");
    scCapBtn.addEventListener("click", function () {
      var name = scNameInp.value.trim();
      if (!name) { showToast("Enter a scene name"); return; }
      state.multitoolSteps[idx]._scenes.push({
        name: name, timestamp: state.currentTimestamp, threshold: 0.75,
      });
      scNameInp.value = "";
      renderMtScenes();
      refreshCalibration({ debounce: true });
      showToast("Scene '" + name + "' at " + formatTime(state.currentTimestamp, { decimals: 1 }));
    });
    addScCtrl.appendChild(scCapBtn);
    addScRow.appendChild(addScCtrl);
    body.appendChild(addScRow);
  }

  function _mtRenderInactivity(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    var r1 = el("div", "param-row");
    r1.appendChild(el("span", "param-label", "Sensitivity"));
    var c1 = el("div", "param-control");
    var inactSlider = rangeInput("paramInactThresh" + sfx, 0, 30, intOrDefault(init.threshold, 10), 1);
    c1.appendChild(inactSlider);
    var inactVal = el("span", "param-value");
    inactVal.textContent = inactSlider.value;
    inactSlider.addEventListener("input", function () { inactVal.textContent = inactSlider.value; });
    c1.appendChild(inactVal);
    r1.appendChild(c1);
    body.appendChild(r1);
  }

  // Helpers used by multiple per-type renderers above.
  function _mtAddNumberRow(body, label, id, min, max, def, step) {
    var r = el("div", "param-row");
    r.appendChild(el("span", "param-label", label));
    var c = el("div", "param-control");
    c.appendChild(numberInput(id, min, max, def, step));
    r.appendChild(c);
    body.appendChild(r);
  }
  function _mtAddCheckboxRow(body, label, id, checked) {
    var r = el("div", "param-row");
    r.appendChild(el("span", "param-label", label));
    var c = el("div", "param-control");
    var cb = document.createElement("input");
    cb.type = "checkbox";
    cb.id = id;
    cb.checked = !!checked;
    c.appendChild(cb);
    r.appendChild(c);
    body.appendChild(r);
  }
  function _mtAddCaptureRefRow(body, idx, label, tsLabelId) {
    var r = el("div", "param-row");
    r.appendChild(el("span", "param-label", label));
    var c = el("div", "param-control");
    var capBtn = el("button", "btn btn-small", "Capture Frame");
    var tsLabel = el("span", "param-value", "\u2014");
    tsLabel.id = tsLabelId;
    if (state.multitoolSteps[idx]._refTs !== undefined) {
      tsLabel.textContent = formatTime(state.multitoolSteps[idx]._refTs, { decimals: 1 });
    }
    capBtn.addEventListener("click", function () {
      state.multitoolSteps[idx]._refTs = state.currentTimestamp;
      tsLabel.textContent = formatTime(state.currentTimestamp, { decimals: 1 });
      refreshCalibration({ debounce: true });
    });
    c.appendChild(capBtn);
    c.appendChild(tsLabel);
    r.appendChild(c);
    body.appendChild(r);
  }

  var MULTITOOL_PARAM_RENDERERS = {
    color:      _mtRenderColor,
    change:     _mtRenderChange,
    similarity: _mtRenderSimilarity,
    text:       _mtRenderText,
    numbers:    _mtRenderNumbers,
    template:   _mtRenderTemplate,
    flow:       _mtRenderFlow,
    scene:      _mtRenderScene,
    inactivity: _mtRenderInactivity,
  };

  var _multitoolDragMidpoints = null;
  var _multitoolDragOverRaf = null;
  var _multitoolPendingDragOver = null;

  function _cacheMultitoolDragMidpoints(container) {
    var cards = container.querySelectorAll(".multitool-step:not(.dragging)");
    var mids = new Array(cards.length);
    for (var i = 0; i < cards.length; i++) {
      var r = cards[i].getBoundingClientRect();
      mids[i] = r.top + r.height / 2;
    }
    _multitoolDragMidpoints = mids;
  }

  function getMultitoolDropIndex(container, clientY) {
    var mids = _multitoolDragMidpoints;
    if (!mids) {
      _cacheMultitoolDragMidpoints(container);
      mids = _multitoolDragMidpoints;
    }
    for (var i = 0; i < mids.length; i++) {
      if (clientY < mids[i]) return i;
    }
    return mids.length;
  }

  function clearMultitoolDragIndicators(container) {
    var cards = container.querySelectorAll(".multitool-step.drag-over");
    for (var i = 0; i < cards.length; i++) cards[i].classList.remove("drag-over");
    container.classList.remove("drag-over-append");
  }

  function taskToMultitoolStep(task) {
    var type = task.type;
    var allowed = MULTITOOL_ALLOWED_TYPES.some(function (t) { return t.value === type; });
    if (!allowed) return null;
    var params = task.parameters || {};
    var step = { type: type, collapsed: false, logic: "AND" };
    if (task.region) step.region = task.region;
    if (task.region_ref) step.region_ref = task.region_ref;
    if (params.reference_timestamp !== undefined) step._refTs = params.reference_timestamp;
    if (params.scene_references) step._scenes = params.scene_references.map(function (ref) {
      return { name: ref.name, timestamp: ref.timestamp, threshold: numberOrDefault(ref.threshold, 0.75) };
    });
    step._initial = params;
    return step;
  }

  // Walk the current step list, read each step's per-input DOM values, and
  // store them on step._initial in the same shape the _mtRender* helpers
  // expect. Call before any list mutation (add / remove / reorder / import)
  // so the upcoming renderWorkflowParams() restores values instead of
  // collapsing them back to per-input defaults. Region, _refTs and _scenes
  // already live on the step object, so they don't need snapshotting.
  function snapshotMultitoolStepValues() {
    state.multitoolSteps.forEach(function (step, idx) {
      var sfx = "_mt" + idx;
      var init = step._initial || {};
      if (step.type === "color") {
        var prevColor = init.target_color || {};
        init.target_color = {
          h: numberOrDefault((qs("#paramColorH" + sfx) || {}).value, prevColor.h || 0),
          s: numberOrDefault((qs("#paramColorS" + sfx) || {}).value, prevColor.s || 0),
          v: numberOrDefault((qs("#paramColorV" + sfx) || {}).value, prevColor.v || 0),
        };
        var tol = numberOrDefault((qs("#paramColorTol" + sfx) || {}).value, 30);
        init.tolerance = {
          h: Math.round(tol * 90 / 100),
          s: Math.round(tol * 128 / 100),
          v: Math.round(tol * 128 / 100),
        };
        var snapMode = (qs("#paramColorMode" + sfx) || {}).value;
        if (snapMode === "presence") {
          init.color_mode = "presence";
          init.min_coverage = numberOrDefault((qs("#paramColorMinArea" + sfx) || {}).value, 1) / 100;
        } else {
          init.color_mode = "average";
          delete init.min_coverage;
        }
      } else if (step.type === "change") {
        init.threshold = numberOrDefault((qs("#paramChangeThresh" + sfx) || {}).value, init.threshold);
        init.noise_threshold = intOrDefault((qs("#paramChangeNoise" + sfx) || {}).value, init.noise_threshold);
      } else if (step.type === "similarity") {
        init.threshold = numberOrDefault((qs("#paramSimThresh" + sfx) || {}).value, init.threshold);
      } else if (step.type === "text") {
        var searchEl = qs("#paramTextSearch" + sfx);
        if (searchEl) init.search_string = searchEl.value;
        init.fuzzy_threshold = numberOrDefault((qs("#paramTextFuzzy" + sfx) || {}).value, init.fuzzy_threshold);
        init.ocr_confidence_threshold = numberOrDefault((qs("#paramTextOcrConf" + sfx) || {}).value, init.ocr_confidence_threshold);
        init.ocr_preprocess = !!((qs("#paramTextOcrPreprocess" + sfx) || {}).checked);
        init.ocr_normalize = (qs("#paramTextOcrNormalize" + sfx) || {}).value || "off";
      } else if (step.type === "numbers") {
        var opEl = qs("#paramNumOperator" + sfx);
        if (opEl) init.operator = opEl.value;
        init.target_value = numberOrDefault((qs("#paramNumTarget" + sfx) || {}).value, init.target_value);
        init.ocr_confidence_threshold = numberOrDefault((qs("#paramNumOcrConf" + sfx) || {}).value, init.ocr_confidence_threshold);
        init.ocr_preprocess = !!((qs("#paramNumOcrPreprocess" + sfx) || {}).checked);
        init.integers_only = !!((qs("#paramNumIntegersOnly" + sfx) || {}).checked);
      } else if (step.type === "template") {
        init.threshold = numberOrDefault((qs("#paramTemplateThresh" + sfx) || {}).value, init.threshold);
        var tScalePct = numberOrDefault((qs("#paramTemplateScale" + sfx) || {}).value, 100);
        init.template_scale = tScalePct / 100;
      } else if (step.type === "flow") {
        init.magnitude_threshold = numberOrDefault((qs("#paramFlowMag" + sfx) || {}).value, init.magnitude_threshold);
      } else if (step.type === "inactivity") {
        init.threshold = intOrDefault((qs("#paramInactThresh" + sfx) || {}).value, init.threshold);
      }
      step._initial = init;
    });
  }

  function renderMultitoolParams(container) {
    var stepsDiv = el("div", "multitool-steps");
    state.multitoolSteps.forEach(function (step, idx) {
      if (idx > 0) {
        var opRow = el("div", "multitool-operator-row");
        var rail = el("div", "multitool-operator-rail");
        rail.appendChild(el("div", "multitool-operator-line"));
        var opBtn = el("button", "multitool-operator-btn");
        opBtn.type = "button";
        var current = (step.logic || "AND").toUpperCase();
        opBtn.classList.add(current === "NOT" ? "is-not" : "is-and");
        opBtn.title = current === "NOT"
          ? "NOT: frame rejected if this matches (click to switch to AND)"
          : "AND: frame must also match (click to switch to NOT)";
        opBtn.appendChild(el("span", "multitool-operator-icon"));
        (function (capturedIdx) {
          opBtn.addEventListener("click", function (e) {
            e.stopPropagation();
            var s = state.multitoolSteps[capturedIdx];
            var next = (s.logic || "AND").toUpperCase() === "NOT" ? "AND" : "NOT";
            s.logic = next;
            opBtn.classList.toggle("is-not", next === "NOT");
            opBtn.classList.toggle("is-and", next === "AND");
            opBtn.title = next === "NOT"
              ? "NOT: frame rejected if this matches (click to switch to AND)"
              : "AND: frame must also match (click to switch to NOT)";
            refreshCalibration({ debounce: true });
          });
        })(idx);
        rail.appendChild(opBtn);
        rail.appendChild(el("div", "multitool-operator-line"));
        opRow.appendChild(rail);

        // Offset window: a pill that reveals min/max second inputs. The window
        // is measured relative to the previous step's matched frame (see
        // scan_multitool's offset path). Presence of `step.offset` = enabled.
        var maxOffset = (CLIPGEN_CONFIG && CLIPGEN_CONFIG.screenspaceMultitoolMaxOffset) || 30;
        var offWrap = el("div", "multitool-offset");
        var offBtn = el("button", "multitool-offset-btn");
        offBtn.type = "button";
        var offActive = !!step.offset;
        offBtn.classList.toggle("is-active", offActive);
        offBtn.appendChild(el("span", "multitool-offset-icon"));
        offWrap.appendChild(offBtn);

        var offFields = el("div", "multitool-offset-fields" + (offActive ? "" : " hidden"));
        var minCtrl = el("div", "param-control");
        var offMin = numberInput("paramMtOffsetMin" + idx, -maxOffset, maxOffset,
          step.offset ? step.offset.min : 0, 0.1);
        minCtrl.appendChild(offMin);
        var offSep = el("span", "multitool-offset-sep");
        var maxCtrl = el("div", "param-control");
        var offMax = numberInput("paramMtOffsetMax" + idx, -maxOffset, maxOffset,
          step.offset ? step.offset.max : 5, 0.1);
        maxCtrl.appendChild(offMax);
        offFields.appendChild(minCtrl);
        offFields.appendChild(offSep);
        offFields.appendChild(maxCtrl);
        offFields.appendChild(el("span", "param-value", "s"));
        offWrap.appendChild(offFields);

        function setOffsetTitle(active) {
          offBtn.title = active
            ? "Offset window on: match within a time window of the previous step's frame (not evaluated in calibration; click to disable)"
            : "Offset: match within a time window relative to the previous step";
        }
        setOffsetTitle(offActive);

        (function (capturedIdx) {
          function syncOffsetFields() {
            var s = state.multitoolSteps[capturedIdx];
            if (!s.offset) return;
            var mn = numberOrDefault(offMin.value, 0);
            var mx = numberOrDefault(offMax.value, 5);
            s.offset.min = mn;
            s.offset.max = mx;
            var invalid = mn > mx;
            offMin.classList.toggle("is-invalid", invalid);
            offMax.classList.toggle("is-invalid", invalid);
            refreshCalibration({ debounce: true });
          }
          offBtn.addEventListener("click", function (e) {
            e.stopPropagation();
            var s = state.multitoolSteps[capturedIdx];
            if (s.offset) {
              delete s.offset;
              offBtn.classList.remove("is-active");
              offFields.classList.add("hidden");
              offMin.classList.remove("is-invalid");
              offMax.classList.remove("is-invalid");
              setOffsetTitle(false);
            } else {
              s.offset = {
                min: numberOrDefault(offMin.value, 0),
                max: numberOrDefault(offMax.value, 5),
              };
              offBtn.classList.add("is-active");
              offFields.classList.remove("hidden");
              setOffsetTitle(true);
            }
            refreshCalibration({ debounce: true });
          });
          offMin.addEventListener("input", syncOffsetFields);
          offMax.addEventListener("input", syncOffsetFields);
        })(idx);

        opRow.appendChild(offWrap);
        stepsDiv.appendChild(opRow);
      }
      var card = el("div", "multitool-step");
      card.dataset.stepIdx = String(idx);
      var header = el("div", "multitool-step-header");

      // Drag handle
      var dragHandle = el("span", "multitool-step-drag-handle");
      dragHandle.appendChild(iconSpan("bars-2"));
      dragHandle.addEventListener("mousedown", function () { card.setAttribute("draggable", "true"); });
      dragHandle.addEventListener("mouseup", function () { card.removeAttribute("draggable"); });
      header.appendChild(dragHandle);

      header.appendChild(el("span", "multitool-step-num", String(idx + 1)));
      var typeSpan = el("span", "multitool-step-type");
      typeSpan.style.color = taskTypeColor(step.type);
      var icon = buildTypeIcon(step.type);
      if (icon) typeSpan.appendChild(icon);
      typeSpan.appendChild(document.createTextNode(" " + step.type.charAt(0).toUpperCase() + step.type.slice(1)));
      header.appendChild(typeSpan);
      var removeBtn = el("button", "multitool-step-remove");
      removeBtn.title = "Remove step";
      removeBtn.appendChild(iconSpan("x-mark"));
      (function (capturedIdx) {
        removeBtn.addEventListener("click", function (e) {
          e.stopPropagation();
          snapshotMultitoolStepValues();
          state.multitoolSteps.splice(capturedIdx, 1);
          renderWorkflowParams();
          updateRunButton();
        });
      })(idx);
      header.appendChild(removeBtn);
      card.appendChild(header);

      var body = el("div", "multitool-step-body" + (step.collapsed ? " collapsed" : ""));
      renderMultitoolStepBody(body, step.type, idx);
      card.appendChild(body);

      header.addEventListener("click", function (e) {
        if (e.target.closest(".multitool-step-drag-handle") || e.target.closest(".multitool-step-remove")) return;
        step.collapsed = !step.collapsed;
        body.classList.toggle("collapsed", step.collapsed);
      });

      stepsDiv.appendChild(card);
    });

    if (state.multitoolSteps.length === 0) {
      // Visible drop target so a Task card has somewhere to land when the
      // step list is empty (an empty flex container is 0px tall and never
      // receives dragover events).
      var emptyDz = el("div", "multitool-empty-dropzone",
        "Drag a Task here, or use + Add Step below");
      stepsDiv.appendChild(emptyDz);
    }

    // Drag-and-drop reordering
    stepsDiv.addEventListener("dragstart", function (e) {
      var card = e.target.closest(".multitool-step");
      if (!card) { e.preventDefault(); return; }
      card.classList.add("dragging");
      e.dataTransfer.setData("text/plain", card.dataset.stepIdx);
      e.dataTransfer.effectAllowed = "move";
      _cacheMultitoolDragMidpoints(stepsDiv);
    });
    stepsDiv.addEventListener("dragend", function (e) {
      var card = e.target.closest(".multitool-step");
      if (card) {
        card.classList.remove("dragging");
        card.removeAttribute("draggable");
      }
      if (_multitoolDragOverRaf != null) {
        cancelAnimationFrame(_multitoolDragOverRaf);
        _multitoolDragOverRaf = null;
      }
      _multitoolPendingDragOver = null;
      clearMultitoolDragIndicators(stepsDiv);
      _multitoolDragMidpoints = null;
    });
    stepsDiv.addEventListener("dragover", function (e) {
      if (e.dataTransfer.types.indexOf("text/plain") < 0) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      _multitoolPendingDragOver = {
        clientY: e.clientY,
        isTaskDrop: e.dataTransfer.types.indexOf("application/x-task-id") >= 0,
      };
      if (_multitoolDragOverRaf != null) return;
      _multitoolDragOverRaf = requestAnimationFrame(function () {
        _multitoolDragOverRaf = null;
        var pending = _multitoolPendingDragOver;
        if (!pending) return;
        clearMultitoolDragIndicators(stepsDiv);
        if (pending.isTaskDrop) {
          stepsDiv.classList.add("drag-over-append");
        } else {
          var cards = stepsDiv.querySelectorAll(".multitool-step:not(.dragging)");
          var insertIdx = getMultitoolDropIndex(stepsDiv, pending.clientY);
          if (insertIdx < cards.length) {
            cards[insertIdx].classList.add("drag-over");
          } else {
            stepsDiv.classList.add("drag-over-append");
          }
        }
      });
    });
    stepsDiv.addEventListener("dragleave", function (e) {
      var card = e.target.closest(".multitool-step");
      if (card) card.classList.remove("drag-over");
      if (!stepsDiv.contains(e.relatedTarget)) {
        stepsDiv.classList.remove("drag-over-append");
      }
    });
    stepsDiv.addEventListener("drop", function (e) {
      e.preventDefault();
      clearMultitoolDragIndicators(stepsDiv);
      stepsDiv.classList.remove("drag-over-append");

      // Check for task card drop (import task as step)
      var taskId = e.dataTransfer.getData("application/x-task-id");
      if (taskId) {
        var task = findTask(taskId);
        if (!task) return;
        var step = taskToMultitoolStep(task);
        if (!step) {
          showToast(task.type + " cannot be added as a multitool step");
          return;
        }
        snapshotMultitoolStepValues();
        state.multitoolSteps.push(step);
        renderWorkflowParams();
        showToast("Imported " + task.type + " task as step");
        return;
      }

      // Step reorder. _cacheMultitoolDragMidpoints excludes the dragging
      // card, so getMultitoolDropIndex already returns an index aligned with
      // the array AFTER the dragging step is spliced out — no further
      // adjustment needed.
      var fromIdx = parseInt(e.dataTransfer.getData("text/plain"), 10);
      if (isNaN(fromIdx)) return;
      var toIdx = getMultitoolDropIndex(stepsDiv, e.clientY);
      if (fromIdx === toIdx) return;
      snapshotMultitoolStepValues();
      var moved = state.multitoolSteps.splice(fromIdx, 1)[0];
      state.multitoolSteps.splice(toIdx, 0, moved);
      renderWorkflowParams();
    });

    container.appendChild(stepsDiv);

    // Add Step row
    var addRow = el("div", "multitool-add-row");
    var sel = document.createElement("select");
    sel.id = "mtAddTypeSelect";
    MULTITOOL_ALLOWED_TYPES.forEach(function (t) {
      var opt = document.createElement("option");
      opt.value = t.value; opt.textContent = t.label;
      sel.appendChild(opt);
    });
    addRow.appendChild(sel);
    var addBtn = el("button", "btn btn-small", "+ Add Step");
    addBtn.addEventListener("click", function () {
      var chosen = sel.value;
      snapshotMultitoolStepValues();
      state.multitoolSteps.push({ type: chosen, collapsed: false, logic: "AND" });
      renderWorkflowParams();
      updateRunButton();
    });
    addRow.appendChild(addBtn);
    container.appendChild(addRow);

    if (state.multitoolSteps.length === 1) {
      container.appendChild(el("div", "multitool-hint", "Add at least 2 tool steps to create a multi-factor filter."));
    }
  }

  // ---- Published back to the hub (screenspace.js calls these) ----
  SS.renderMultitoolParams = renderMultitoolParams;
  SS.clearMultitoolDragIndicators = clearMultitoolDragIndicators;
  // The global dragend handler lives in the hub (it also clears task-list drag);
  // expose the multitool-drag half so the hub doesn't touch our internal state.
  SS.cancelMultitoolDrag = function () {
    if (_multitoolDragOverRaf != null) {
      cancelAnimationFrame(_multitoolDragOverRaf);
      _multitoolDragOverRaf = null;
    }
    _multitoolPendingDragOver = null;
  };
})();
