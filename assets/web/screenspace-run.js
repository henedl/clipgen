/* clipgen Screenspace run satellite — screenspace-run.js
 *
 * The Run button and parameter gathering carved out of screenspace.js:
 * initRunButton (queues one task per selected participant, or a multitool
 * chain) and gatherWorkflowParams / gatherMultitoolStepParams (the single
 * save path every tool's panel feeds). Loads last: it destructures the
 * tasks satellite's renderTaskList / startSSE and the hub's region-ref
 * helpers at load time. Function bodies are unchanged from the hub.
 */
(function () {
  "use strict";

  var SS = window.ClipgenScreenspace;
  var state = SS.state;
  var activeRegionRef = SS.activeRegionRef,
    fullFrameRegionRef = SS.fullFrameRegionRef,
    normalizeRegionRef = SS.normalizeRegionRef,
    regionRefLabel = SS.regionRefLabel,
    regionRefPayload = SS.regionRefPayload,
    renderTaskList = SS.renderTaskList,
    startSSE = SS.startSSE,
    toolSupportsFastScan = SS.toolSupportsFastScan;


  function initRunButton() {
    qs("#runBtn").addEventListener("click", function () {
      var type = state.activeWorkflow;
      // Boundary and Attention are full-frame only: always scan the whole
      // frame, ignoring any selected region.
      var isFullFrameTool = type === "boundary" || type === "attention";
      var regions = isFullFrameTool
        ? [fullFrameRegionRef()]
        : (state.runRegions.length > 0
            ? state.runRegions
            : (state.activeRegion ? [activeRegionRef(state.activeRegion)] : []));
      // Multitool uses per-step regions; skip global region requirement
      var isMultitool = type === "multitool";
      // Template/shape with an uploaded image can run without a region
      // (full-frame scan).
      if (!isMultitool && !isFullFrameTool && regions.length === 0
          && !((type === "template" || type === "shape") && state.uploadedTemplate)) return;
      if (regions.length === 0) regions = [""];
      var participants = state.runParticipants.length > 0
        ? state.runParticipants
        : (state.selectedParticipant ? [state.selectedParticipant] : []);
      if (participants.length === 0) return;
      var params = gatherWorkflowParams(type);
      if (params === null) return;
      if (state.scanMode === "fast" && toolSupportsFastScan(type)) params.scan_mode = "fast";

      if (state.inMarker !== null) params.start_seconds = state.inMarker;
      if (state.outMarker !== null) params.end_seconds = state.outMarker;

      var chain = Promise.resolve();
      if (isMultitool) {
        // Multitool: one task per participant, first step's region as top-level
        var mtRegion = (params.steps && params.steps.length > 0) ? (params.steps[0].region || "") : "";
        participants.forEach(function (pid) {
          chain = chain.then(function () {
            var body = {
              type: type,
              participant: pid,
              region: mtRegion,
              parameters: params,
            };
            return apiPost("api/tasks", body).then(function (data) {
              if (data.ok) {
                if (!state.tasks.some(function (t) { return t.id === data.task.id; })) {
                  state.tasks.push(data.task);
                }
                renderTaskList();
              } else {
                showToast(data.error || "Failed to create task for " + pid);
              }
            });
          });
        });
      } else {
        participants.forEach(function (pid) {
          regions.forEach(function (regionRef) {
            chain = chain.then(function () {
              var normalizedRegion = normalizeRegionRef(regionRef);
              var body = {
                type: type,
                participant: pid,
                region: normalizedRegion ? normalizedRegion.name : "",
                parameters: params,
              };
              if (normalizedRegion) body.region_ref = regionRefPayload(normalizedRegion);
              return apiPost("api/tasks", body).then(function (data) {
                if (data.ok) {
                  if (!state.tasks.some(function (t) { return t.id === data.task.id; })) {
                    state.tasks.push(data.task);
                  }
                  renderTaskList();
                } else {
                  showToast(data.error || "Failed to create task for " + pid + " / " + regionRefLabel(normalizedRegion));
                }
              });
            });
          });
        });
      }
      var totalTasks = isMultitool ? participants.length : participants.length * regions.length;
      chain.then(function () {
        showToast(totalTasks + " task" + (totalTasks !== 1 ? "s" : "") + " queued: " + type);
        startSSE();
      }).catch(function (err) { showToast("Error: " + err.message); });
    });
  }

  function gatherMultitoolStepParams(stepType, idx, opts) {
    // opts.silent drops missing-input toasts for the calibration strip's per-keystroke
    // probes.
    var silent = !!(opts && opts.silent);
    function toast(msg) { if (!silent) showToast(msg); }
    var sfx = "_mt" + idx;
    // Suffix-aware readers; rawNum deliberately yields NaN for inputs whose emptiness is
    // checked.
    function num(id, d) { return numberOrDefault((qs("#" + id + sfx) || {}).value, d); }
    function intv(id, d) { return intOrDefault((qs("#" + id + sfx) || {}).value, d); }
    function chk(id) { return !!((qs("#" + id + sfx) || {}).checked); }
    function str(id, d) { return (qs("#" + id + sfx) || {}).value || d; }
    function rawNum(id) { return parseFloat((qs("#" + id + sfx) || {}).value); }
    var p = {};
    if (stepType === "color") {
      p.target_color = {
        h: num("paramColorH", 0),
        s: num("paramColorS", 0),
        v: num("paramColorV", 0),
      };
      var tol = num("paramColorTol", 30);
      p.tolerance = {
        h: Math.round(tol * 90 / 100),
        s: Math.round(tol * 128 / 100),
        v: Math.round(tol * 128 / 100),
      };
      if (str("paramColorMode", "") === "presence") {
        p.color_mode = "presence";
        p.min_coverage = num("paramColorMinArea", 1) / 100;
      }
    } else if (stepType === "change") {
      p.threshold = num("paramChangeThresh", 0.03);
      p.noise_threshold = intv("paramChangeNoise", 30);
    } else if (stepType === "similarity") {
      var step = state.multitoolSteps[idx];
      if (!step || step._refTs === undefined) {
        toast("Step " + (idx + 1) + ": capture a reference frame first");
        return null;
      }
      p.reference_timestamp = step._refTs;
      p.threshold = num("paramSimThresh", 0.90);
    } else if (stepType === "text") {
      p.search_string = str("paramTextSearch", "");
      if (!p.search_string.trim()) {
        toast("Step " + (idx + 1) + ": enter a search string");
        return null;
      }
      p.fuzzy_threshold = num("paramTextFuzzy", CLIPGEN_CONFIG.screenspaceOcrFuzzyThreshold);
      p.ocr_confidence_threshold = num("paramTextOcrConf", CLIPGEN_CONFIG.screenspaceOcrMinConfidence);
      p.ocr_preprocess = chk("paramTextOcrPreprocess");
      p.ocr_normalize = str("paramTextOcrNormalize", "off");
    } else if (stepType === "numbers") {
      p.operator = str("paramNumOperator", "gt");
      p.target_value = rawNum("paramNumTarget");
      if (isNaN(p.target_value)) {
        toast("Step " + (idx + 1) + ": enter a valid target number");
        return null;
      }
      p.ocr_confidence_threshold = num("paramNumOcrConf", CLIPGEN_CONFIG.screenspaceOcrMinConfidence);
      p.ocr_preprocess = chk("paramNumOcrPreprocess");
      p.integers_only = chk("paramNumIntegersOnly");
    } else if (stepType === "template") {
      step = state.multitoolSteps[idx];
      if (step && step._upload) {
        p.template_image_data = step._upload.data;
      } else if (step && step._refTs !== undefined) {
        p.reference_timestamp = step._refTs;
      } else {
        toast("Step " + (idx + 1) + ": capture a template frame or upload a PNG");
        return null;
      }
      p.threshold = num("paramTemplateThresh", 0.70);
      var tScalePct = rawNum("paramTemplateScale");
      if (!isNaN(tScalePct) && tScalePct > 0 && tScalePct !== 100) {
        p.template_scale = tScalePct / 100;
      }
    } else if (stepType === "flow") {
      p.magnitude_threshold = num("paramFlowMag", 2.0);
    } else if (stepType === "scene") {
      step = state.multitoolSteps[idx];
      if (!step || !step._scenes || step._scenes.length === 0) {
        toast("Step " + (idx + 1) + ": add at least one scene reference");
        return null;
      }
      p.scene_references = step._scenes.map(function (ref) {
        return { name: ref.name, timestamp: ref.timestamp, threshold: numberOrDefault(ref.threshold, 0.75) };
      });
    } else if (stepType === "inactivity") {
      p.threshold = intv("paramInactThresh", 10);
    }
    var stepRegionRef = normalizeRegionRef(state.multitoolSteps[idx].region_ref)
      || (state.multitoolSteps[idx].region ? activeRegionRef(state.multitoolSteps[idx].region) : null);
    p.region = stepRegionRef ? stepRegionRef.name : "";
    if (stepRegionRef) p.region_ref = regionRefPayload(stepRegionRef);
    return p;
  }

  function gatherWorkflowParams(type, opts) {
    // opts.silent drops missing-input toasts for the calibration strip's per-keystroke
    // probes.
    var silent = !!(opts && opts.silent);
    function toast(msg) { if (!silent) showToast(msg); }
    var sfx = "";
    // Suffix-aware readers; rawNum deliberately yields NaN for inputs whose emptiness is
    // checked.
    function num(id, d) { return numberOrDefault((qs("#" + id + sfx) || {}).value, d); }
    function intv(id, d) { return intOrDefault((qs("#" + id + sfx) || {}).value, d); }
    function chk(id) { return !!((qs("#" + id + sfx) || {}).checked); }
    function str(id, d) { return (qs("#" + id + sfx) || {}).value || d; }
    function rawNum(id) { return parseFloat((qs("#" + id + sfx) || {}).value); }
    var params = {};
    if (type === "multitool") {
      if (state.multitoolSteps.length < 2) {
        toast("Add at least 2 steps");
        return null;
      }
      params.steps = [];
      for (var i = 0; i < state.multitoolSteps.length; i++) {
        var stepP = gatherMultitoolStepParams(state.multitoolSteps[i].type, i, opts);
        if (stepP === null) return null;
        stepP.type = state.multitoolSteps[i].type;
        if (i > 0) {
          stepP.logic = (state.multitoolSteps[i].logic || "AND").toUpperCase();
          var off = state.multitoolSteps[i].offset;
          if (off && isFinite(off.min) && isFinite(off.max)) {
            if (Number(off.min) > Number(off.max)) {
              toast("Step " + (i + 1) + ": offset min must be ≤ max");
              return null;
            }
            stepP.offset = { min: Number(off.min), max: Number(off.max) };
          }
        }
        params.steps.push(stepP);
      }
      params.interval = num("paramMultitoolInterval", 1.0);
      var mtLabelEl = qs("#paramEventLabel");
      if (mtLabelEl && mtLabelEl.value.trim()) params.event_label = mtLabelEl.value.trim();
      var mtDfEl = qs("#paramDetectFirst");
      if (mtDfEl && mtDfEl.checked) params.detect_first = true;
      return params;
    } else if (type === "color") {
      params.target_color = {
        h: num("paramColorH", 0),
        s: num("paramColorS", 0),
        v: num("paramColorV", 0),
      };
      var tol = num("paramColorTol", 30);
      params.tolerance = {
        h: Math.round(tol * 90 / 100),
        s: Math.round(tol * 128 / 100),
        v: Math.round(tol * 128 / 100),
      };
      if (str("paramColorMode", "") === "presence") {
        params.color_mode = "presence";
        params.min_coverage = num("paramColorMinArea", 1) / 100;
      }
      params.interval = num("paramColorInterval", 1.0);
    } else if (type === "change") {
      params.threshold = num("paramChangeThresh", 0.03);
      params.noise_threshold = intv("paramChangeNoise", 30);
      params.interval = num("paramChangeInterval", 1.0);
      var rcChange = intv("paramChangeConsecutive", 1);
      if (rcChange > 1) params.require_consecutive = rcChange;
    } else if (type === "similarity") {
      if (state.referenceTimestamp === null) {
        toast("Capture a reference frame first");
        return null;
      }
      params.reference_timestamp = state.referenceTimestamp;
      params.threshold = num("paramSimThresh", 0.90);
      params.interval = num("paramSimInterval", 1.0);
    } else if (type === "text") {
      params.search_string = str("paramTextSearch", "");
      if (!params.search_string.trim()) {
        toast("Enter a search string");
        return null;
      }
      params.fuzzy_threshold = num("paramTextFuzzy", CLIPGEN_CONFIG.screenspaceOcrFuzzyThreshold);
      params.ocr_confidence_threshold = num("paramTextOcrConf", CLIPGEN_CONFIG.screenspaceOcrMinConfidence);
      params.ocr_preprocess = chk("paramTextOcrPreprocess");
      params.ocr_normalize = str("paramTextOcrNormalize", "off");
      params.interval = num("paramTextInterval", 2.0);
      var lang = str("paramTextLang", "en");
      params.languages = [lang];
      var rcText = intv("paramTextConsecutive", 1);
      if (rcText > 1) params.require_consecutive = rcText;
    } else if (type === "numbers") {
      var op = str("paramNumOperator", "gt");
      params.operator = op;
      if (op === "range") {
        params.range_min = rawNum("paramNumMin");
        params.range_max = rawNum("paramNumMax");
        if (isNaN(params.range_min) || isNaN(params.range_max)) {
          toast("Enter valid min and max values");
          return null;
        }
        if (params.range_min > params.range_max) {
          toast("Min must be less than or equal to max");
          return null;
        }
      } else {
        params.target_value = rawNum("paramNumTarget");
        if (isNaN(params.target_value)) {
          toast("Enter a valid target number");
          return null;
        }
      }
      params.ocr_confidence_threshold = num("paramNumOcrConf", CLIPGEN_CONFIG.screenspaceOcrMinConfidence);
      params.ocr_preprocess = chk("paramNumOcrPreprocess");
      params.integers_only = chk("paramNumIntegersOnly");
      params.interval = num("paramNumInterval", 2.0);
      var rcNum = intv("paramNumConsecutive", 1);
      if (rcNum > 1) params.require_consecutive = rcNum;
    } else if (type === "timelapse") {
      params.speedup_factor = num("paramTlSpeed", 10);
      var si = rawNum("paramTlSampleInterval");
      if (si > 0) params.sample_interval = si;
      params.output_format = str("paramTlFormat", "mp4");
    } else if (type === "template") {
      if (state.uploadedTemplate) {
        params.template_image_data = state.uploadedTemplate.data;
        if (state.uploadedTemplate.name) params.template_name = state.uploadedTemplate.name;
      } else if (state.referenceTimestamp !== null) {
        params.reference_timestamp = state.referenceTimestamp;
        // Pin the sample to its capture region; the run target only scopes the search.
        var tplSnap = state.capturedRefPreview;
        if (tplSnap && tplSnap.ts === state.referenceTimestamp && tplSnap.region) {
          params.reference_region = tplSnap.region;
        }
      } else {
        toast("Capture a template region or upload a PNG");
        return null;
      }
      params.threshold = num("paramTemplateThresh", 0.70);
      params.interval = num("paramTemplateInterval", 1.0);
      var scalePct = rawNum("paramTemplateScale");
      if (!isNaN(scalePct) && scalePct > 0 && scalePct !== 100) {
        params.template_scale = scalePct / 100;
      }
    } else if (type === "shape") {
      if (state.uploadedTemplate) {
        params.shape_image_data = state.uploadedTemplate.data;
        if (state.uploadedTemplate.name) params.shape_name = state.uploadedTemplate.name;
      } else if (state.referenceTimestamp !== null) {
        params.reference_timestamp = state.referenceTimestamp;
        // Pin the sample to its capture region; the run target only scopes the search.
        var capSnap = state.capturedRefPreview;
        if (capSnap && capSnap.ts === state.referenceTimestamp && capSnap.region) {
          params.reference_region = capSnap.region;
        }
      } else {
        toast("Capture a shape region or upload a PNG");
        return null;
      }
      params.threshold = num("paramShapeThresh", 0.55);
      params.scale_min = num("paramShapeScaleMin", 50) / 100;
      params.scale_max = num("paramShapeScaleMax", 200) / 100;
      params.scale_steps = intv("paramShapeSteps", 7);
      var linkEl = qs("#paramShapeLinkAxes");
      if (linkEl && !linkEl.checked) {
        params.scale_y_min = num("paramShapeScaleYMin", 90) / 100;
        params.scale_y_max = num("paramShapeScaleYMax", 110) / 100;
        params.scale_y_steps = intv("paramShapeStepsY", 3);
      }
      params.interval = num("paramShapeInterval", 1.0);
    } else if (type === "flow") {
      params.magnitude_threshold = num("paramFlowMag", 2.0);
      params.interval = num("paramFlowInterval", 1.0);
      var rcFlow = intv("paramFlowConsecutive", 1);
      if (rcFlow > 1) params.require_consecutive = rcFlow;
    } else if (type === "scene") {
      if (state.sceneReferences.length === 0) {
        toast("Add at least one scene reference");
        return null;
      }
      params.scene_references = state.sceneReferences.map(function (ref) {
        return { name: ref.name, timestamp: ref.timestamp, threshold: numberOrDefault(ref.threshold, 0.75) };
      });
      params.interval = num("paramSceneInterval", 1.0);
    } else if (type === "inactivity") {
      params.threshold = intv("paramInactThresh", 10);
      params.min_duration = num("paramInactMinDur", 2.0);
      params.interval = num("paramInactInterval", 1.0);
    } else if (type === "boundary") {
      params.threshold = intv("paramBoundaryThresh", 14);
      params.min_gap = num("paramBoundaryMinGap", 3.0);
      params.interval = num("paramBoundaryInterval", 1.0);
      // Auto ("") omits metric so the server applies its configured default.
      var boundaryMetric = str("paramBoundaryMetric", "");
      if (boundaryMetric) params.metric = boundaryMetric;
    } else if (type === "attention") {
      params.shift_threshold = num("paramAttnShift", 0.15);
      params.ema_alpha = num("paramAttnSmooth", 0.6);
      params.weight_spectral = num("paramAttnWSpectral", 1.0);
      params.weight_contrast = num("paramAttnWContrast", 0.7);
      params.weight_motion = num("paramAttnWMotion", 1.2);
      params.weight_face = num("paramAttnWFace", 0);
      params.center_bias = num("paramAttnCenterBias", 0.25);
      params.interval = num("paramAttnInterval", 0.5);
    }
    var labelEl = qs("#paramEventLabel");
    if (labelEl && labelEl.value.trim()) {
      params.event_label = labelEl.value.trim();
    }
    var dfEl = qs("#paramDetectFirst");
    if (dfEl && dfEl.checked) {
      params.detect_first = true;
    }
    return params;
  }

  SS.initRunButton = initRunButton;
  SS.gatherWorkflowParams = gatherWorkflowParams;
  SS.gatherMultitoolStepParams = gatherMultitoolStepParams;
})();
