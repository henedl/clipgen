/* clipgen Screenspace params satellite — screenspace-params.js
 *
 * The single-tool parameter panels (color, change, similarity, text, numbers,
 * template, shape, scene, timelapse, …), renderWorkflowParams, and the
 * reset-to-default buttons, carved out of screenspace.js. Loads right after
 * screenspace-model-view.js: tasks / calibration / multitool-params /
 * overlay-interaction destructure renderWorkflowParams, refTimeChip and
 * updateParamResetButtons at load time. Functions owned by later satellites
 * (overlay, overlay-interaction, sample-editor, color) are reached late-bound
 * through SS. Function bodies are unchanged from the hub.
 */
(function () {
  "use strict";

  var SS = window.ClipgenScreenspace;
  var state = SS.state;
  var _updateMinAreaReadout = SS._updateMinAreaReadout,
    _updateOverlayUi = SS._updateOverlayUi,
    activatePipette = SS.activatePipette,
    applyColorMode = SS.applyColorMode,
    buildColorModeControl = SS.buildColorModeControl,
    buildNormalizeControl = SS.buildNormalizeControl,
    buildTypeIcon = SS.buildTypeIcon,
    deactivatePipette = SS.deactivatePipette,
    hideToolInfoTooltip = SS.hideToolInfoTooltip,
    loadFrame = SS.loadFrame,
    normalizeRegionRef = SS.normalizeRegionRef,
    refreshCalibration = SS.refreshCalibration,
    refreshModelView = SS.refreshModelView,
    renderCalibration = SS.renderCalibration,
    renderIntervalSlot = SS.renderIntervalSlot,
    toolSupportsFastScan = SS.toolSupportsFastScan,
    updateCalibrationVisibility = SS.updateCalibrationVisibility,
    updateRunButton = SS.updateRunButton;

  var _paletteDocListeners = null;


  function renderColorParams(container) {
    var pickerGroup = el("div", "color-picker-group");

    var palette = document.createElement("canvas");
    palette.id = "colorPalette";
    palette.className = "color-palette-canvas";
    pickerGroup.appendChild(palette);

    var bright = document.createElement("canvas");
    bright.id = "colorBrightness";
    bright.className = "color-brightness-strip";
    pickerGroup.appendChild(bright);

    var inputRow = el("div", "color-input-row");
    var preview = el("div", "color-preview");
    preview.id = "colorPreview";
    inputRow.appendChild(preview);

    var hexInput = document.createElement("input");
    hexInput.type = "text";
    hexInput.autocomplete = "off";
    hexInput.id = "paramColorHex";
    hexInput.className = "color-hex-input";
    hexInput.placeholder = "#000000";
    hexInput.maxLength = 7;
    inputRow.appendChild(hexInput);

    var pipetteBtn = el("button", "btn btn-small btn-pipette");
    pipetteBtn.id = "pipetteBtn";
    pipetteBtn.appendChild(buildTypeIcon("color"));
    pipetteBtn.title = "Pick color from video frame";
    pipetteBtn.addEventListener("click", function () {
      if (state.pipetteActive) deactivatePipette();
      else activatePipette();
    });
    inputRow.appendChild(pipetteBtn);

    var sampleBtn = el("button", "btn btn-small color-sample-btn");
    sampleBtn.id = "colorSampleBtn";
    sampleBtn.addEventListener("click", SS.sampleColorFromRegion);
    inputRow.appendChild(sampleBtn);
    SS.updateColorSampleBtnLabel();

    pickerGroup.appendChild(inputRow);

    var hiddenH = document.createElement("input");
    hiddenH.type = "hidden"; hiddenH.id = "paramColorH"; hiddenH.value = "90";
    var hiddenS = document.createElement("input");
    hiddenS.type = "hidden"; hiddenS.id = "paramColorS"; hiddenS.value = "200";
    var hiddenV = document.createElement("input");
    hiddenV.type = "hidden"; hiddenV.id = "paramColorV"; hiddenV.value = "200";
    pickerGroup.appendChild(hiddenH);
    pickerGroup.appendChild(hiddenS);
    pickerGroup.appendChild(hiddenV);
    SS.setColorHiddenInputs({ h: hiddenH, s: hiddenS, v: hiddenV, hex: hexInput });

    container.appendChild(pickerGroup);

    var paletteDragging = false;
    var brightDragging = false;
    function pickFromPalette(e) {
      var rect = palette.getBoundingClientRect();
      var x = clamp(e.clientX - rect.left, 0, rect.width);
      var y = clamp(e.clientY - rect.top, 0, rect.height);
      var h = Math.round((x / rect.width) * 180);
      var s = Math.round((1 - y / rect.height) * 255);
      var curV = numberOrDefault(hiddenV.value, 0);
      SS.setTargetColor(h, s, curV);
    }
    palette.addEventListener("mousedown", function (e) {
      e.preventDefault();
      paletteDragging = true;
      pickFromPalette(e);
    });

    function pickFromBrightness(e) {
      var rect = bright.getBoundingClientRect();
      var x = clamp(e.clientX - rect.left, 0, rect.width);
      var v = Math.round((x / rect.width) * 255);
      var curH = numberOrDefault(hiddenH.value, 0);
      var curS = numberOrDefault(hiddenS.value, 0);
      SS.setTargetColor(curH, curS, v);
    }
    bright.addEventListener("mousedown", function (e) {
      e.preventDefault();
      brightDragging = true;
      pickFromBrightness(e);
    });

    if (_paletteDocListeners) {
      document.removeEventListener("mousemove", _paletteDocListeners.move);
      document.removeEventListener("mouseup", _paletteDocListeners.up);
    }
    function onDocMove(e) {
      if (paletteDragging) pickFromPalette(e);
      if (brightDragging) pickFromBrightness(e);
    }
    function onDocUp() { paletteDragging = false; brightDragging = false; }
    document.addEventListener("mousemove", onDocMove);
    document.addEventListener("mouseup", onDocUp);
    _paletteDocListeners = { move: onDocMove, up: onDocUp };

    hexInput.addEventListener("input", function () {
      var rgb = hexToRgb(hexInput.value);
      if (rgb) {
        var hsv = rgbToHsv(rgb.r, rgb.g, rgb.b);
        hiddenH.value = hsv.h;
        hiddenS.value = hsv.s;
        hiddenV.value = hsv.v;
        SS.updateColorPreview();
        SS.renderColorPalette();
        SS.renderBrightnessStrip();
      }
    });

    var tolSlider = rangeInput("paramColorTol", 0, 100, 30);
    addParamRow(container, "Tolerance", tolSlider, "paramColorTolVal");
    tolSlider.addEventListener("input", function () {
      SS.renderColorPalette();
    });

    // Match mode (average vs presence) + the presence-only "Min area" row.
    var minAreaRow;
    var modeRow = el("div", "param-row");
    modeRow.appendChild(el("span", "param-label", "Mode"));
    var modeControl = el("div", "param-control");
    modeControl.appendChild(
      buildColorModeControl("paramColorMode", "average", false, function (mode) {
        if (minAreaRow) minAreaRow.classList.toggle("hidden", mode !== "presence");
      })
    );
    modeRow.appendChild(modeControl);
    container.appendChild(modeRow);
    minAreaRow = addParamRow(
      container, "Min area %", rangeInput("paramColorMinArea", 0, 100, 1, 1), "paramColorMinAreaVal"
    );
    minAreaRow.id = "paramColorMinAreaRow";
    minAreaRow.classList.add("hidden");
    // Region-aware readout; runs after addParamRow's generic listener and overwrites the
    // raw value.
    var minAreaVal = qs("#paramColorMinAreaVal");
    if (minAreaVal) minAreaVal.classList.add("param-value--minarea");
    var minAreaSlider = qs("#paramColorMinArea");
    if (minAreaSlider) {
      minAreaSlider.addEventListener("input", function () {
        _updateMinAreaReadout("");
      });
    }
    _updateMinAreaReadout("");

    renderIntervalSlot("paramColorInterval", 0.5, 60, 1.0, 0.5);

    SS.renderColorPalette();
    SS.renderBrightnessStrip();
    SS.updateColorPreview();
    var initRgb = hsvToRgb(90, 200, 200);
    hexInput.value = rgbToHex(initRgb.r, initRgb.g, initRgb.b);
  }

  function renderSimilarityParams(container) {
    var refRow = el("div", "param-row");
    var refLabel = el("span", "param-label", "Reference");
    var refControl = el("div", "param-control");
    var refBtn = el("button", "btn btn-small", "Capture Current Frame");
    refBtn.addEventListener("click", function () {
      state.referenceTimestamp = state.currentTimestamp;
      state.capturedRefPreview = captureRefSnapshot(state.currentTimestamp);
      renderWorkflowParams();
      showToast("Reference frame captured at " + formatTime(state.currentTimestamp, { decimals: 1 }));
    });
    refControl.appendChild(refBtn);
    if (state.referenceTimestamp !== null) {
      refControl.appendChild(refTimeChip(state.referenceTimestamp));
      var simSnapInfo = refSnapshotInfo("reference");
      if (simSnapInfo) refControl.appendChild(simSnapInfo);
    }
    refRow.appendChild(refLabel);
    refRow.appendChild(refControl);
    container.appendChild(refRow);
    addParamRow(container, "Threshold", rangeInput("paramSimThresh", 0.50, 1.00, 0.90, 0.01), "paramSimThreshVal");
    renderIntervalSlot("paramSimInterval", 0.5, 60, 1.0, 0.5);
  }

  function renderTextParams(container) {
    addParamRow(container, "Search text", textInput("paramTextSearch", "Enter text to find..."));
    addParamRow(container, "Fuzzy Thr.", rangeInput("paramTextFuzzy", 0.50, 1.00, numberOrDefault(CLIPGEN_CONFIG.screenspaceOcrFuzzyThreshold, 0.75), 0.01), "paramTextFuzzyVal");
    addParamRow(container, "Min OCR conf.", rangeInput("paramTextOcrConf", 0.00, 1.00, numberOrDefault(CLIPGEN_CONFIG.screenspaceOcrMinConfidence, 0.6), 0.01), "paramTextOcrConfVal");
    renderIntervalSlot("paramTextInterval", 0.5, 60, 2.0, 0.5);
    var langRow = el("div", "param-row");
    langRow.appendChild(el("span", "param-label", "Language"));
    var langControl = el("div", "param-control");
    var langSel = document.createElement("select");
    langSel.id = "paramTextLang";
    [["en", "English"], ["es", "Spanish"], ["fr", "French"], ["de", "German"],
     ["ja", "Japanese"], ["ko", "Korean"], ["zh", "Chinese"]].forEach(function (pair) {
      var opt = el("option", null, pair[1]);
      opt.value = pair[0];
      langSel.appendChild(opt);
    });
    langControl.appendChild(langSel);
    langRow.appendChild(langControl);
    container.appendChild(langRow);
    var ppCb = document.createElement("input");
    ppCb.type = "checkbox";
    ppCb.id = "paramTextOcrPreprocess";
    addParamRow(container, "Enhance ROI", ppCb);
    addParamRow(container, "Normalize", buildNormalizeControl("paramTextOcrNormalize", "off"));
    addParamRow(container, "Consecutive", numberInput("paramTextConsecutive", 1, 10, 1, 1));
  }

  function renderNumbersParams(container) {
    var opRow = el("div", "param-row");
    opRow.appendChild(el("span", "param-label", "Operator"));
    var opControl = el("div", "param-control");
    var opSel = document.createElement("select");
    opSel.id = "paramNumOperator";
    [["gt", "Greater than (>)"], ["lt", "Less than (<)"], ["eq", "Equal to (=)"],
     ["gte", "Greater or equal (\u2265)"], ["lte", "Less or equal (\u2264)"], ["range", "In range"]].forEach(function (pair) {
      var opt = el("option", null, pair[1]);
      opt.value = pair[0];
      opSel.appendChild(opt);
    });
    opSel.addEventListener("change", function () {
      var rangeRow = qs("#paramNumRangeRow");
      var targetRow = qs("#paramNumTargetRow");
      if (opSel.value === "range") {
        if (rangeRow) rangeRow.style.display = "";
        if (targetRow) targetRow.style.display = "none";
      } else {
        if (rangeRow) rangeRow.style.display = "none";
        if (targetRow) targetRow.style.display = "";
      }
    });
    opControl.appendChild(opSel);
    opRow.appendChild(opControl);
    container.appendChild(opRow);
    var targetRow = el("div", "param-row");
    targetRow.id = "paramNumTargetRow";
    targetRow.appendChild(el("span", "param-label", "Target value"));
    var targetCtrl = el("div", "param-control");
    targetCtrl.appendChild(numberInput("paramNumTarget", -999999, 999999, 100, 1));
    targetRow.appendChild(targetCtrl);
    container.appendChild(targetRow);
    var numRangeRow = el("div", "param-row");
    numRangeRow.id = "paramNumRangeRow";
    numRangeRow.style.display = "none";
    numRangeRow.appendChild(el("span", "param-label", "Range"));
    var rangeCtrl = el("div", "param-control");
    rangeCtrl.appendChild(numberInput("paramNumMin", -999999, 999999, 0, 1));
    rangeCtrl.appendChild(el("span", "param-value", "\u2013"));
    rangeCtrl.appendChild(numberInput("paramNumMax", -999999, 999999, 100, 1));
    numRangeRow.appendChild(rangeCtrl);
    container.appendChild(numRangeRow);
    addParamRow(container, "Min OCR conf.", rangeInput("paramNumOcrConf", 0.00, 1.00, numberOrDefault(CLIPGEN_CONFIG.screenspaceOcrMinConfidence, 0.6), 0.01), "paramNumOcrConfVal");
    var ppCb = document.createElement("input");
    ppCb.type = "checkbox";
    ppCb.id = "paramNumOcrPreprocess";
    addParamRow(container, "Enhance ROI", ppCb);
    var ioCb = document.createElement("input");
    ioCb.type = "checkbox";
    ioCb.id = "paramNumIntegersOnly";
    addParamRow(container, "Integers only", ioCb);
    renderIntervalSlot("paramNumInterval", 0.5, 60, 2.0, 0.5);
    addParamRow(container, "Consecutive", numberInput("paramNumConsecutive", 1, 10, 1, 1));
  }

  function renderTimelapseParams(container) {
    addParamRow(container, "Speed", numberInput("paramTlSpeed", 2, 100, 10, 1));
    addParamRow(container, "Sample every", numberInput("paramTlSampleInterval", 0, 60, 0, 0.5), "paramTlSampleIntervalVal");
    var siHint = el("span", "param-hint", "seconds (0 = every frame)");
    container.lastChild.querySelector(".param-control").appendChild(siHint);
    var fmtRow = el("div", "param-row");
    fmtRow.appendChild(el("span", "param-label", "Format"));
    var fmtControl = el("div", "param-control");
    var fmtSel = document.createElement("select");
    fmtSel.id = "paramTlFormat";
    [["mp4", "Video (.mp4)"], ["gif", "GIF (.gif)"]].forEach(function (pair) {
      var opt = el("option", null, pair[1]);
      opt.value = pair[0];
      fmtSel.appendChild(opt);
    });
    fmtControl.appendChild(fmtSel);
    fmtRow.appendChild(fmtControl);
    container.appendChild(fmtRow);
  }

  // Client-side same-frame crop of the region for the reference row; the run re-extracts
  // server-side.
  function captureRefSnapshot(ts) {
    var img = state.frameImage;
    var regs = state.previewRegions || state.regions;
    var name = state.activeRegion && regs[state.activeRegion] ? state.activeRegion : null;
    if (!name) {
      for (var i = state.runRegions.length - 1; i >= 0; i--) {
        var ref = normalizeRegionRef(state.runRegions[i]);
        if (ref && ref.name && regs[ref.name]) { name = ref.name; break; }
      }
    }
    var r = name && regs[name];
    if (!img || !img.naturalWidth || !r) return null;
    var sw = Math.max(1, Math.round(r.w * img.naturalWidth));
    var sh = Math.max(1, Math.round(r.h * img.naturalHeight));
    var c = document.createElement("canvas");
    c.width = sw;
    c.height = sh;
    c.getContext("2d").drawImage(
      img,
      r.x * img.naturalWidth, r.y * img.naturalHeight, sw, sh,
      0, 0, sw, sh
    );
    return { dataUrl: c.toDataURL("image/png"), region: name, ts: ts };
  }

  // Last-capture thumbnail + region label. A task-Edit restore has no pixels, so thumbnail
  // is optional.
  function refSnapshotInfo(labelText, editable) {
    var snap = state.capturedRefPreview;
    if (!snap || snap.ts !== state.referenceTimestamp) return null;
    if (!snap.dataUrl && !snap.region) return null;
    var capInfo = el("span", "param-value template-upload-info");
    if (snap.dataUrl) {
      var capThumb = document.createElement("img");
      capThumb.decoding = "async";
      capThumb.className = "ss-sample-thumb";
      capThumb.src = snap.dataUrl;
      capThumb.alt = "Captured " + labelText.toLowerCase();
      capThumb.title = snap.region;
      capThumb.addEventListener("click", function () {
        SS.openSampleModal({
          mode: editable ? "edit" : "view",
          title: "Captured " + labelText.toLowerCase(),
          dataUrl: snap.dataUrl,
          regionName: snap.region,
          onApply: function (b64) {
            // An edited capture becomes an upload; the server cannot re-derive its pixels.
            applyEditedSample((snap.region || "sample") + "-edited.png", b64);
          },
        });
      });
      capInfo.appendChild(capThumb);
    }
    if (snap.region) capInfo.appendChild(el("span", "param-hint", snap.region));
    return capInfo;
  }

  // Install an edited sample as the uploaded reference (Template/Shape).
  function applyEditedSample(name, b64) {
    state.uploadedTemplate = { name: name, data: b64 };
    state.referenceTimestamp = null;
    state.capturedRefPreview = null;
    state.templateOverlayPos = null;
    var previewImg = new Image();
    previewImg.onload = function () { SS.renderOverlay(); };
    previewImg.src = "data:image/png;base64," + b64;
    state.uploadedTemplateImg = previewImg;
    renderWorkflowParams();
    updateRunButton();
    refreshModelView({ debounce: true });
  }

  // Template + Shape reference row; the drag overlay stays template-only. opts.draw adds
  // paint-on-frame.
  function renderRefCaptureRow(container, labelText, opts) {
    var tmplRefRow = el("div", "param-row");
    tmplRefRow.appendChild(el("span", "param-label", labelText));
    var tmplRefCtrl = el("div", "param-control");
    var tmplCapBtn = el("button", "btn btn-small ss-template-icon-btn ss-template-icon-btn--capture");
    tmplCapBtn.setAttribute("type", "button");
    tmplCapBtn.title = "Capture Region";
    tmplCapBtn.setAttribute("aria-label", "Capture Region");
    var tmplCapGlyph = el("span", "ss-template-icon-btn__glyph");
    tmplCapBtn.appendChild(tmplCapGlyph);
    tmplCapBtn.addEventListener("click", function () {
      var snap = captureRefSnapshot(state.currentTimestamp);
      // Refuse a region-less capture: the sample would silently become the run region.
      if (!snap) {
        showToast("Select or draw a region to capture from");
        return;
      }
      state.referenceTimestamp = state.currentTimestamp;
      state.uploadedTemplate = null;
      state.capturedRefPreview = snap;
      renderWorkflowParams();
      showToast(labelText + " captured at " + formatTime(state.currentTimestamp, { decimals: 1 }));
    });
    tmplRefCtrl.appendChild(tmplCapBtn);

    var tmplFileInput = document.createElement("input");
    tmplFileInput.type = "file";
    tmplFileInput.accept = "image/png";
    tmplFileInput.style.display = "none";
    tmplFileInput.addEventListener("change", function () {
      var file = tmplFileInput.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function (e) {
        var dataUrl = e.target.result;
        var b64 = dataUrl.split(",")[1];
        state.uploadedTemplate = { name: file.name, data: b64 };
        state.referenceTimestamp = null;
        state.capturedRefPreview = null;
        state.templateOverlayPos = null;
        var previewImg = new Image();
        previewImg.onload = function () { SS.renderOverlay(); };
        previewImg.src = dataUrl;
        state.uploadedTemplateImg = previewImg;
        renderWorkflowParams();
        showToast(labelText + " loaded");
      };
      reader.readAsDataURL(file);
    });
    var tmplUploadBtn = el("button", "btn btn-small ss-template-icon-btn ss-template-icon-btn--upload");
    tmplUploadBtn.setAttribute("type", "button");
    tmplUploadBtn.title = "Upload PNG";
    tmplUploadBtn.setAttribute("aria-label", "Upload PNG");
    var tmplUploadGlyph = el("span", "ss-template-icon-btn__glyph");
    tmplUploadBtn.appendChild(tmplUploadGlyph);
    tmplUploadBtn.addEventListener("click", function () { tmplFileInput.click(); });
    tmplRefCtrl.appendChild(tmplUploadBtn);
    tmplRefCtrl.appendChild(tmplFileInput);

    if (opts && opts.draw) {
      var drawBtn = el("button", "btn btn-small ss-template-icon-btn ss-template-icon-btn--draw");
      drawBtn.setAttribute("type", "button");
      drawBtn.title = "Draw shape on frame";
      drawBtn.setAttribute("aria-label", "Draw shape on frame");
      drawBtn.appendChild(el("span", "ss-template-icon-btn__glyph"));
      // The row re-renders wholesale; active look derives from state alone.
      drawBtn.classList.toggle("active", !!state.shapeDraw);
      drawBtn.addEventListener("click", SS.toggleShapeDraw);
      tmplRefCtrl.appendChild(drawBtn);
    }

    if (state.uploadedTemplate) {
      if (!state.uploadedTemplateImg) {
        var liveImg = new Image();
        liveImg.onload = function () { SS.renderOverlay(); };
        liveImg.src = "data:image/png;base64," + state.uploadedTemplate.data;
        state.uploadedTemplateImg = liveImg;
      }
      var uploadInfo = el("span", "param-value template-upload-info");
      var uploadThumb = document.createElement("img");
      uploadThumb.decoding = "async";
      uploadThumb.className = "ss-sample-thumb";
      uploadThumb.src = "data:image/png;base64," + state.uploadedTemplate.data;
      uploadThumb.alt = "Uploaded " + labelText.toLowerCase();
      uploadThumb.title = state.uploadedTemplate.name;
      uploadThumb.addEventListener("click", function () {
        var up = state.uploadedTemplate;
        if (!up) return;
        SS.openSampleModal({
          mode: "edit",
          title: up.name || "Uploaded " + labelText.toLowerCase(),
          dataUrl: "data:image/png;base64," + up.data,
          onApply: function (b64) {
            applyEditedSample(up.name || "sample.png", b64);
          },
        });
      });
      uploadInfo.appendChild(uploadThumb);
      var clearBtn = el("button", "btn btn-small", "\u00d7");
      clearBtn.addEventListener("click", function () {
        state.uploadedTemplate = null;
        state.uploadedTemplateImg = null;
        state.templateOverlayPos = null;
        renderWorkflowParams();
        SS.renderOverlay();
      });
      uploadInfo.appendChild(clearBtn);
      tmplRefCtrl.appendChild(uploadInfo);
    } else if (state.referenceTimestamp !== null) {
      tmplRefCtrl.appendChild(refTimeChip(state.referenceTimestamp));
      // Task-Edit restores have no snapshot; the ts guard falls back to the time chip.
      var capInfo = refSnapshotInfo(labelText, true);
      if (capInfo) tmplRefCtrl.appendChild(capInfo);
    }
    tmplRefRow.appendChild(tmplRefCtrl);
    container.appendChild(tmplRefRow);
  }

  function renderTemplateParams(container) {
    renderRefCaptureRow(container, "Template");
    addParamRow(container, "Threshold", rangeInput("paramTemplateThresh", 0.50, 1.00, 0.70, 0.01), "paramTemplateThreshVal");
    addParamRow(container, "Template scale", rangeInput("paramTemplateScale", 25, 200, 100, 5), "paramTemplateScaleVal");
    var scaleHint = el("span", "param-hint", "% \u2014 resize the uploaded PNG before matching");
    container.lastChild.querySelector(".param-control").appendChild(scaleHint);
    var scaleSlider = qs("#paramTemplateScale");
    if (scaleSlider) {
      state.templateScalePreview = numberOrDefault(scaleSlider.value, 100) / 100;
      scaleSlider.addEventListener("input", function () {
        state.templateScalePreview = numberOrDefault(scaleSlider.value, 100) / 100;
        SS.renderOverlay();
      });
    }

    renderIntervalSlot("paramTemplateInterval", 0.5, 60, 1.0, 0.5);
  }

  function renderShapeParams(container) {
    renderRefCaptureRow(container, "Shape", { draw: true });
    addParamRow(container, "Threshold", rangeInput("paramShapeThresh", 0.30, 1.00, 0.55, 0.01), "paramShapeThreshVal");
    addParamRow(container, "Scale min", rangeInput("paramShapeScaleMin", 25, 400, 50, 5), "paramShapeScaleMinVal");
    var rowXMin = container.lastChild;
    addParamRow(container, "Scale max", rangeInput("paramShapeScaleMax", 25, 400, 200, 5), "paramShapeScaleMaxVal");
    var rowXMax = container.lastChild;
    addParamRow(container, "Scale steps", numberInput("paramShapeSteps", 1, 12, 7, 1));
    var rowXSteps = container.lastChild;
    // Unlinked axes: width-only sliders plus a height ladder; the sweep multiplies cost.
    var linkCb = document.createElement("input");
    linkCb.type = "checkbox";
    linkCb.id = "paramShapeLinkAxes";
    linkCb.checked = true;
    addParamRow(container, "Link axes", linkCb);
    addParamRow(container, "Height scale min", rangeInput("paramShapeScaleYMin", 25, 400, 90, 5), "paramShapeScaleYMinVal");
    var rowYMin = container.lastChild;
    addParamRow(container, "Height scale max", rangeInput("paramShapeScaleYMax", 25, 400, 110, 5), "paramShapeScaleYMaxVal");
    var rowYMax = container.lastChild;
    addParamRow(container, "Height scale steps", numberInput("paramShapeStepsY", 1, 12, 3, 1));
    var rowYSteps = container.lastChild;
    function syncAxisRows() {
      var show = !linkCb.checked;
      rowYMin.style.display = show ? "" : "none";
      rowYMax.style.display = show ? "" : "none";
      rowYSteps.style.display = show ? "" : "none";
      // Unlinked, the base ladder is width-only; say so in its labels.
      rowXMin.firstChild.textContent = show ? "Width scale min" : "Scale min";
      rowXMax.firstChild.textContent = show ? "Width scale max" : "Scale max";
      rowXSteps.firstChild.textContent = show ? "Width scale steps" : "Scale steps";
    }
    linkCb.addEventListener("change", syncAxisRows);
    syncAxisRows();
    renderIntervalSlot("paramShapeInterval", 0.5, 60, 1.0, 0.5);
  }

  function renderSceneParams(container) {
    var sceneList = el("div", "scene-reference-list");
    sceneList.id = "sceneRefList";
    state.sceneReferences.forEach(function (ref, i) {
      if (ref.threshold === undefined) ref.threshold = 0.75;
      var item = el("div", "scene-ref-item");
      item.appendChild(el("span", "scene-ref-name", ref.name));
      if (ref._thumb) {
        var scThumbWrap = el("span", "param-value template-upload-info");
        var scThumb = document.createElement("img");
        scThumb.decoding = "async";
        scThumb.className = "ss-sample-thumb";
        scThumb.src = ref._thumb;
        scThumb.alt = "Scene sample";
        if (ref._thumbRegion) scThumb.title = ref._thumbRegion;
        scThumb.addEventListener("click", function () {
          SS.openSampleModal({ mode: "view", title: ref.name || "Scene sample", dataUrl: ref._thumb });
        });
        scThumbWrap.appendChild(scThumb);
        item.appendChild(scThumbWrap);
      }
      item.appendChild(refTimeChip(ref.timestamp));
      var threshSlider = document.createElement("input");
      threshSlider.type = "range";
      threshSlider.min = "0.50";
      threshSlider.max = "1.00";
      threshSlider.step = "0.01";
      threshSlider.value = String(ref.threshold);
      threshSlider.className = "scene-ref-thresh";
      var threshVal = el("span", "param-value", String(ref.threshold));
      threshSlider.addEventListener("input", (function (idx) {
        return function () {
          state.sceneReferences[idx].threshold = parseFloat(threshSlider.value);
          threshVal.textContent = threshSlider.value;
        };
      })(i));
      item.appendChild(threshSlider);
      item.appendChild(threshVal);
      var rmBtn = el("button", "btn btn-small", "\u00d7");
      rmBtn.addEventListener("click", function () {
        state.sceneReferences.splice(i, 1);
        renderWorkflowParams();
      });
      item.appendChild(rmBtn);
      sceneList.appendChild(item);
    });
    container.appendChild(sceneList);
    var addScRow = el("div", "param-row");
    addScRow.appendChild(el("span", "param-label", "Add Scene"));
    var addScCtrl = el("div", "param-control");
    var scNameInp = textInput("paramSceneName", "e.g. menu, gameplay");
    addScCtrl.appendChild(scNameInp);
    var scCapBtn = el("button", "btn btn-small", "Capture");
    scCapBtn.addEventListener("click", function () {
      var nameEl = qs("#paramSceneName");
      var name = nameEl ? nameEl.value.trim() : "";
      if (!name) { showToast("Enter a scene name"); return; }
      // Underscore fields are display-only; gather strips them before the server sees
      // refs.
      var scSnap = captureRefSnapshot(state.currentTimestamp);
      state.sceneReferences.push({
        name: name,
        timestamp: state.currentTimestamp,
        threshold: 0.75,
        _thumb: scSnap && scSnap.dataUrl,
        _thumbRegion: scSnap && scSnap.region,
      });
      renderWorkflowParams();
      showToast("Scene '" + name + "' at " + formatTime(state.currentTimestamp, { decimals: 1 }));
    });
    addScCtrl.appendChild(scCapBtn);
    addScRow.appendChild(addScCtrl);
    container.appendChild(addScRow);
    renderIntervalSlot("paramSceneInterval", 0.5, 60, 1.0, 0.5);
  }

  // DOM-only param inputs snapshot by tool-prefixed id across rebuilds; multitool steps
  // (positional ids) self-restore.
  var _MT_STEP_ID = /_mt\d+$/;

  function _paramControlValue(el) {
    return el.type === "checkbox" ? el.checked : el.value;
  }

  // The id'd form controls under `root` that participate in save/restore/reset.
  function _paramControls(root) {
    var out = [];
    var nodes = root.querySelectorAll("[id]");
    for (var i = 0; i < nodes.length; i++) {
      var tag = nodes[i].tagName;
      if (tag !== "INPUT" && tag !== "SELECT" && tag !== "TEXTAREA") continue;
      if (_MT_STEP_ID.test(nodes[i].id)) continue;
      out.push(nodes[i]);
    }
    return out;
  }

  function _snapshotParamValues() {
    var map = {};
    ["#workflowParams", "#workflowIntervalSlot"].forEach(function (id) {
      var root = qs(id);
      if (!root) return;
      _paramControls(root).forEach(function (el) {
        map[el.id] = _paramControlValue(el);
      });
    });
    return map;
  }

  function _mergeParamMap(target, src) {
    Object.keys(src).forEach(function (id) { target[id] = src[id]; });
  }

  function _restoreParamValues(map) {
    Object.keys(map).forEach(function (id) {
      var el = document.getElementById(id);
      if (!el) return;
      var tag = el.tagName;
      if (tag !== "INPUT" && tag !== "SELECT" && tag !== "TEXTAREA") return;
      var saved = map[id];
      if (el.type === "checkbox") {
        if (el.checked === saved) return;
        el.checked = saved;
      } else if (el.type === "hidden" && el.parentNode
                 && el.parentNode.classList.contains("cg-segtrack")) {
        // Capsule visuals are CSS off the track, so el.value alone desyncs; apply* also
        // re-toggles rows.
        if (el.value === String(saved)) return;
        if (id === "paramColorMode") applyColorMode(id, saved);
        else segTrackSetValue(el.parentNode, saved);
      } else {
        if (el.value === String(saved)) return;
        el.value = saved;
      }
      // Fire input for readouts and model view, change for change-only listeners (numbers
      // operator row).
      el.dispatchEvent(new Event("input", { bubbles: true }));
      el.dispatchEvent(new Event("change", { bubbles: true }));
    });
  }

  // Only SS.setTargetColor repaints the swatch, hex field and palette; feed it the restored
  // values.
  function _restoreColorTarget() {
    if (state.activeWorkflow !== "color") return;
    var c = SS.getColorHiddenInputs();
    if (!c) return;
    SS.setTargetColor(
      numberOrDefault(c.h.value, 90), numberOrDefault(c.s.value, 200), numberOrDefault(c.v.value, 200)
    );
  }

  // `opts.defaults` skips the restore; task restore writes its own saved params next.
  function renderWorkflowParams(opts) {
    if (state.shapeDraw && state.activeWorkflow !== "shape") SS.cancelShapeDraw();
    _mergeParamMap(state.paramValues, _snapshotParamValues());
    _renderWorkflowParamsBuild();
    // Defaults are read off the just-built panel, not a second table that could drift.
    _mergeParamMap(state.paramDefaults, _snapshotParamValues());
    if (!(opts && opts.defaults)) {
      _restoreParamValues(state.paramValues);
      _restoreColorTarget();
    }
    updateParamResetButtons();
  }

  // ---- Reset-to-default ----
  // Per row, not control: the Range row's two inputs are one parameter.
  function _buildParamResetButton(row) {
    var btn = el("button", "param-reset hidden");
    btn.type = "button";
    var icon = el("span", "param-reset-icon");
    applyIconMask(icon, "arrow-path", "/screenspace/icons/");
    btn.appendChild(icon);
    btn.addEventListener("click", function () {
      var map = {};
      _paramControls(row).forEach(function (c) {
        if (state.paramDefaults[c.id] !== undefined) map[c.id] = state.paramDefaults[c.id];
      });
      // No _restoreColorTarget: the H/S/V inputs sit outside any .param-row.
      _restoreParamValues(map);
      updateParamResetButtons();
    });
    return btn;
  }

  function _syncParamResetButton(row) {
    var ctrl = row.querySelector(".param-control");
    if (!ctrl) return;
    var btn = row.querySelector(".param-reset");
    var eligible = _paramControls(row).filter(function (c) {
      return state.paramDefaults[c.id] !== undefined;
    });
    if (!eligible.length) {
      if (btn) btn.parentNode.removeChild(btn);
      return;
    }
    var changed = eligible.some(function (c) {
      return _paramControlValue(c) !== state.paramDefaults[c.id];
    });
    if (!btn) {
      btn = _buildParamResetButton(row);
      ctrl.appendChild(btn);
    }
    var tip = "Reset to default";
    if (eligible.length === 1 && eligible[0].type !== "checkbox") {
      tip += " (" + state.paramDefaults[eligible[0].id] + ")";
    }
    btn.setAttribute("data-tooltip", tip);
    btn.setAttribute("aria-label", tip);
    btn.classList.toggle("hidden", !changed);
  }

  function updateParamResetButtons() {
    var container = qs("#workflowParams");
    if (container) {
      _paramRows(container).forEach(_syncParamResetButton);
    }
    // The interval control has no .param-row wrapper, so it is its own row.
    var slot = qs("#workflowIntervalSlot");
    if (slot) _syncParamResetButton(slot);
  }

  function _paramRows(container) {
    return Array.prototype.slice.call(container.querySelectorAll(".param-row"));
  }

  // One delegated pair per stable container (like initCalibration) catches hand-assembled
  // rows too.
  function initParamResets() {
    ["workflowParams", "workflowIntervalSlot"].forEach(function (id) {
      var container = qs("#" + id);
      if (!container) return;
      container.addEventListener("input", updateParamResetButtons);
      container.addEventListener("change", updateParamResetButtons);
    });
  }

  function _renderWorkflowParamsBuild() {
    var container = qs("#workflowParams");
    container.innerHTML = "";
    SS.setColorHiddenInputs(null);
    hideToolInfoTooltip(true); // also unpins
    var intervalSlot = qs("#workflowIntervalSlot");
    if (intervalSlot) intervalSlot.innerHTML = "";
    var type = state.activeWorkflow;

    var regionPickerWrap = qs("#runRegionPicker");
    // Multitool uses per-step regions; boundary is full-frame only — both hide
    // the global region picker.
    if (regionPickerWrap) {
      regionPickerWrap.style.display =
        type === "multitool" || type === "boundary" || type === "attention"
          ? "none" : "";
    }

    if (type === "multitool") {
      SS.renderMultitoolParams(container);
      renderIntervalSlot("paramMultitoolInterval", 0.5, 60, 1.0, 0.5);
      var mtEventLabel = textInput("paramEventLabel", "e.g. low_health");
      mtEventLabel.className = "param-input-half";
      addParamRow(container, "Event label", mtEventLabel);
      var dfCb = document.createElement("input");
      dfCb.type = "checkbox";
      dfCb.id = "paramDetectFirst";
      addParamRow(container, "Detect first", dfCb);
      // Every multitool mutation lands here; task-import and reorder paths skip
      // updateRunButton.
      updateRunButton();
      _updateOverlayUi();
      refreshModelView();
      updateCalibrationVisibility();
      // Multitool returns early, so mirror the bottom-of-function calibration reset here.
      state.calibrationResult = null;
      renderCalibration();
      if (!state.suppressCalibrationRefresh) refreshCalibration();
      return;
    }

    if (type === "color") renderColorParams(container);
    else if (type === "change") {
      addParamRow(container, "Threshold", rangeInput("paramChangeThresh", 0.01, 0.50, 0.03, 0.01), "paramChangeThreshVal");
      addParamRow(container, "Noise Thr.", rangeInput("paramChangeNoise", 0, 100, 30, 1), "paramChangeNoiseVal");
      renderIntervalSlot("paramChangeInterval", 0.5, 60, 1.0, 0.5);
      addParamRow(container, "Consecutive", numberInput("paramChangeConsecutive", 1, 10, 1, 1));
    }
    else if (type === "similarity") renderSimilarityParams(container);
    else if (type === "text") renderTextParams(container);
    else if (type === "numbers") renderNumbersParams(container);
    else if (type === "timelapse") renderTimelapseParams(container);
    else if (type === "template") renderTemplateParams(container);
    else if (type === "shape") renderShapeParams(container);
    else if (type === "flow") {
      addParamRow(container, "Magnitude", rangeInput("paramFlowMag", 0.5, 20.0, 2.0, 0.5), "paramFlowMagVal");
      renderIntervalSlot("paramFlowInterval", 0.5, 60, 1.0, 0.5);
      addParamRow(container, "Consecutive", numberInput("paramFlowConsecutive", 1, 10, 1, 1));
    }
    else if (type === "scene") renderSceneParams(container);
    else if (type === "inactivity") {
      addParamRow(container, "Sensitivity", rangeInput("paramInactThresh", 0, 30, 10, 1), "paramInactThreshVal");
      addParamRow(container, "Min duration (s)", numberInput("paramInactMinDur", 0.5, 60, 2.0, 0.5));
      renderIntervalSlot("paramInactInterval", 0.5, 60, 1.0, 0.5);
    }
    else if (type === "boundary") {
      // Metric: Auto sends nothing (server default, currently Hybrid); pHash is the v1
      // spike detector.
      var metricSel = document.createElement("select");
      metricSel.id = "paramBoundaryMetric";
      [["", "Auto"], ["scene", "Scene"], ["phash", "pHash"], ["hybrid", "Hybrid"]].forEach(function (pair) {
        var opt = el("option", null, pair[1]);
        opt.value = pair[0];
        metricSel.appendChild(opt);
      });
      addParamRow(container, "Metric", metricSel);
      // Full phash Hamming span (0..64); drives the pHash metric and Hybrid's spike check.
      addParamRow(container, "Sensitivity", rangeInput("paramBoundaryThresh", 0, 64, 14, 1), "paramBoundaryThreshVal");
      addParamRow(container, "Min gap (s)", numberInput("paramBoundaryMinGap", 0.5, 60, 3.0, 0.5));
      renderIntervalSlot("paramBoundaryInterval", 0.5, 60, 1.0, 0.5);
    }
    else if (type === "attention") {
      // Peak-jump distance as a fraction of the screen diagonal, and the EMA smoothing
      // alpha.
      addParamRow(container, "Sensitivity", rangeInput("paramAttnShift", 0.05, 0.50, 0.15, 0.01), "paramAttnShiftVal");
      addParamRow(container, "Smoothing", rangeInput("paramAttnSmooth", 0.1, 1.0, 0.6, 0.05), "paramAttnSmoothVal");
      // Channel weights (defaults mirror SCREENSPACE_ATTENTION_WEIGHT_*); Faces at 0
      // disables the Haar channel.
      addParamRow(container, "Spectral wt.", rangeInput("paramAttnWSpectral", 0, 2.0, 1.0, 0.05), "paramAttnWSpectralVal");
      addParamRow(container, "Contrast wt.", rangeInput("paramAttnWContrast", 0, 2.0, 0.7, 0.05), "paramAttnWContrastVal");
      addParamRow(container, "Motion wt.", rangeInput("paramAttnWMotion", 0, 2.0, 1.2, 0.05), "paramAttnWMotionVal");
      addParamRow(container, "Faces wt.", rangeInput("paramAttnWFace", 0, 2.0, 0, 0.05), "paramAttnWFaceVal");
      addParamRow(container, "Center bias", rangeInput("paramAttnCenterBias", 0, 1.0, 0.25, 0.05), "paramAttnCenterBiasVal");
      renderIntervalSlot("paramAttnInterval", 0.5, 60, 0.5, 0.5);
    }

    if (type !== "timelapse") {
      var eventLabel = textInput("paramEventLabel", "e.g. low_health");
      eventLabel.className = "param-input-half";
      addParamRow(container, "Event label", eventLabel);
      // Boundary and attention emit transitions, not detections, so "Detect first" doesn't
      // apply.
      if (type !== "boundary" && type !== "attention") {
        var dfCb = document.createElement("input");
        dfCb.type = "checkbox";
        dfCb.id = "paramDetectFirst";
        addParamRow(container, "Detect first", dfCb);
      }
    }

    var scanPicker = qs("#runScanModePicker");
    // Timelapse has no scan modes; boundary runs its own coarse pass. Hide for both.
    if (scanPicker) {
      scanPicker.style.display = toolSupportsFastScan(type) ? "" : "none";
    }
    var scanBtn = scanPicker && scanPicker.querySelector(".scan-toggle-btn");
    if (scanBtn && scanBtn._updateScanState) scanBtn._updateScanState();

    updateRunButton();
    _updateOverlayUi();
    refreshModelView();
    updateCalibrationVisibility();
    // Drop stale scores before the new evaluation; renderCalibration() resets
    // calibrationGreen synchronously.
    state.calibrationResult = null;
    renderCalibration();
    if (!state.suppressCalibrationRefresh) refreshCalibration();
  }

  // Click-to-seek reference timestamp chip shared by every capturing tool; `textId` marks
  // the in-place-updated span.
  function refTimeChip(seconds, textId) {
    var wrap = el("span", "param-value param-value--ref");
    var text = el("span", null, formatTime(seconds, { decimals: 1 }));
    if (textId) text.id = textId;
    wrap.appendChild(text);

    var btn = el("button", "ref-seek-btn");
    btn.type = "button";
    btn.setAttribute("data-tooltip", "Jump to this frame");
    btn.setAttribute("aria-label", "Jump to this frame");
    var icon = el("span", "ref-seek-icon");
    applyIconMask(icon, "arrow-up-right", "/screenspace/icons/");
    btn.appendChild(icon);
    btn.addEventListener("click", function () {
      loadFrame(seconds);
    });
    wrap.appendChild(btn);
    return wrap;
  }

  function addParamRow(container, label, control, valueDisplayId) {
    var row = el("div", "param-row");
    row.appendChild(el("span", "param-label", label));
    var ctrl = el("div", "param-control");
    ctrl.appendChild(control);
    if (valueDisplayId) {
      var valSpan = el("span", "param-value");
      valSpan.id = valueDisplayId;
      valSpan.textContent = control.value;
      ctrl.appendChild(valSpan);
      control.addEventListener("input", function () {
        valSpan.textContent = control.value;
        if (state.activeWorkflow === "color" && control.id && control.id.startsWith("paramColor")) {
          SS.updateColorPreview();
        }
        refreshModelView({ debounce: true });
      });
    } else {
      control.addEventListener("input", function () {
        refreshModelView({ debounce: true });
      });
    }
    row.appendChild(ctrl);
    container.appendChild(row);
    return row;
  }

  SS.initParamResets = initParamResets;
  SS.refTimeChip = refTimeChip;
  SS.renderWorkflowParams = renderWorkflowParams;
  SS.updateParamResetButtons = updateParamResetButtons;
})();
