/* clipgen Screenspace — model-view (preprocessed preview) satellite.
 *
 * Carved out of screenspace.js (the hub) following the hub+satellite convention
 * (see screenspace-overlay/timeline/...). Owns the "Model view" section of the
 * right pane's Preview tab: the live preprocessed-frame preview (api/preview),
 * the overlay-layer catalog + toggle/dropdown UI, the preview-region resolvers,
 * and the color "Min area %" readout.
 *
 * It is a read of the hub's shared `state` plus a few hub helpers, all reached
 * through window.ClipgenScreenspace (SS). apiGet / qs / numberOrDefault /
 * _formatMinAreaReadout are ambient utils.js / screenspace-utils.js globals.
 * renderOverlay lives in screenspace-overlay.js (loaded AFTER this file), so it
 * is reached late-bound via SS.renderOverlay(...). The hub keeps same-named
 * delegators (initModelView / refreshModelView / _updateMinAreaReadout /
 * _updateOverlayUi / _overlayEligibleForActiveTool) for its own call sites.
 *
 * Load order: right after screenspace.js and BEFORE screenspace-overlay.js,
 * screenspace-tasks.js, screenspace-multitool-params.js and
 * screenspace-calibration.js — each destructures one of this file's published
 * helpers (SS._overlayEligibleForActiveTool / SS._updateMinAreaReadout /
 * SS._previewRegionRef) at load time.
 */

(function () {
  "use strict";

  var SS = window.ClipgenScreenspace;
  var state = SS.state;
  // Hub helpers, published before this file loads; other helpers are ambient globals.
  var normalizeRegionRef = SS.normalizeRegionRef,
    activeRegionRef = SS.activeRegionRef;

  // ---- Model view (preprocessed preview) ----

  var _modelViewGen = 0;
  var _modelViewTimer = 0;

  var MODEL_VIEW_META = {
    color: "Downscaled region (≤64 px) with mean HSV vs. target swatch.",
    change: "Gray-blur, colorized abs-diff, and changed pixels tinted on the frame (prev = 1 s earlier).",
    similarity: "Gray-blurred region (≤256 px); reference and SSIM difference map appear once captured.",
    text: "Grayscale region fed to OCR.",
    numbers: "Grayscale region fed to OCR.",
    timelapse: "Region crop. FFmpeg encodes this unmodified.",
    template: "Gray-blurred frame, template, and normalized match heatmap.",
    shape: "Edge ridges and scale-swept match heatmap.",
    flow: "Prev + current gray frames with dense optical-flow vectors.",
    scene: "Region (≤128 px), Canny edges, and 8-bin hue histogram.",
    inactivity: "Region and pHash bit grid (white = 1, black = 0).",
    boundary: "Full frame; Auto/Scene/Hybrid use a content fingerprint vs. the current period, pHash compares consecutive samples.",
    attention: "Full frame (\u2264256 px): spectral residual, Lab contrast, frame-diff motion, and the combined center-weighted saliency map.",
    multitool: "Preview of the first tool step.",
  };

  function initModelView() {
    // Restore persisted overlay preferences (sessionStorage, per-tab).
    try {
      var stored = sessionStorage.getItem("ss_overlayEnabled");
      if (stored === "1") state.overlayEnabled = true;
    } catch (_) { /* sessionStorage may be unavailable */ }
    try {
      var layer = sessionStorage.getItem("ss_overlayLayer");
      if (layer) state.overlayLayer = layer;
    } catch (_) { /* ignore */ }

    var toggle = qs("#modelViewOverlayToggle");
    if (toggle) {
      toggle.checked = !!state.overlayEnabled;
      toggle.addEventListener("change", function () {
        state.overlayEnabled = !!toggle.checked;
        try { sessionStorage.setItem("ss_overlayEnabled", state.overlayEnabled ? "1" : "0"); } catch (_) { /* ignore */ }
        var curTs = Number(state.currentTimestamp || 0).toFixed(3);
        if (state.overlayEnabled && (!state.overlayImage || state.overlayImageTimestamp !== curTs || state.overlayImageTool !== state.activeWorkflow)) {
          refreshModelView();
        }
        SS.renderOverlay();
      });
    }

    var sel = qs("#modelViewOverlayLayer");
    if (sel) {
      sel.addEventListener("change", function () {
        state.overlayLayer = sel.value || null;
        try { sessionStorage.setItem("ss_overlayLayer", state.overlayLayer || ""); } catch (_) { /* ignore */ }
        // Force a refetch of the overlay image at the new layer.
        if (state.overlayImageObjectUrl) {
          URL.revokeObjectURL(state.overlayImageObjectUrl);
          state.overlayImageObjectUrl = null;
        }
        state.overlayImage = null;
        state.overlayImageTimestamp = null;
        state.overlayImageTool = null;
        refreshModelView();
        SS.renderOverlay();
      });
    }

    // Fetch the overlay-layer catalog once at init.
    apiGet("api/preview/layers")
      .then(function (data) {
        if (data && data.ok && data.layers) {
          state.overlayLayerSpec = data.layers;
          _updateOverlayUi();
        }
      })
      .catch(function () { /* leave catalog empty; toggle stays disabled */ });
  }

  function _activeOverlayTool() {
    var tool = state.activeWorkflow;
    if (tool === "multitool") {
      var first = (state.multitoolSteps || [])[0];
      tool = first && first.type ? first.type : null;
    }
    return tool;
  }

  function _activeOverlayLayers() {
    var tool = _activeOverlayTool();
    if (!tool) return [];
    return state.overlayLayerSpec[tool] || [];
  }

  function _overlayEligibleForActiveTool() {
    return _activeOverlayLayers().length > 0;
  }

  // Drive the <select> so its change handler stays the one persist + refetch path.
  function cycleOverlayLayer() {
    var sel = qs("#modelViewOverlayLayer");
    if (!sel || sel.options.length < 2) return;
    sel.selectedIndex = (sel.selectedIndex + 1) % sel.options.length;
    sel.dispatchEvent(new Event("change", { bubbles: true }));
    // The picker lives in the Preview tab; name the layer for everyone else.
    showToast(sel.options[sel.selectedIndex].textContent);
  }

  function _resolveOverlayLayer() {
    var layers = _activeOverlayLayers();
    if (!layers.length) return null;
    if (state.overlayLayer) {
      for (var i = 0; i < layers.length; i++) {
        if (layers[i].id === state.overlayLayer) return layers[i];
      }
    }
    return layers[0];
  }

  function _updateOverlayUi() {
    var toggle = qs("#modelViewOverlayToggle");
    var sel = qs("#modelViewOverlayLayer");
    var label = toggle && toggle.parentElement;
    var layers = _activeOverlayLayers();
    var eligible = layers.length > 0;

    if (toggle) {
      toggle.disabled = !eligible;
      if (label) {
        if (eligible) {
          label.classList.remove("disabled");
          label.removeAttribute("title");
        } else {
          label.classList.add("disabled");
          label.setAttribute("title", "This tool's preview isn't pixel-aligned to the frame");
        }
      }
      if (!eligible && toggle.checked) {
        toggle.checked = false;
        // Don't persist away the user's preference; just disable for this tool.
      }
    }

    if (sel) {
      if (layers.length > 1) {
        sel.classList.remove("hidden");
        sel.innerHTML = "";
        var resolved = _resolveOverlayLayer();
        layers.forEach(function (layer) {
          var opt = document.createElement("option");
          opt.value = layer.id;
          opt.textContent = layer.label;
          if (resolved && layer.id === resolved.id) opt.selected = true;
          sel.appendChild(opt);
        });
      } else {
        sel.classList.add("hidden");
        sel.innerHTML = "";
      }
    }
  }

  // Inactive Preview tab hides the image; overlay and B-blink draw on the video instead.
  function refreshModelView(opts) {
    if (state.rightPaneTab !== "preview" && !state.overlayEnabled && !state.overlayBlinkActive) return;
    if (_modelViewTimer) {
      clearTimeout(_modelViewTimer);
      _modelViewTimer = 0;
    }
    if (opts && opts.debounce) {
      _modelViewTimer = setTimeout(_doRefreshModelView, 150);
    } else {
      _doRefreshModelView();
    }
  }

  var _FULL_FRAME_REGION_STRING = "0.000000,0.000000,1.000000,1.000000";

  // Region {x,y,w,h} (normalized when source_width is set, else canvas pixels) → fraction string.
  function _regionDataToString(r) {
    if (r.source_width) {
      return [r.x, r.y, r.w, r.h]
        .map(function (v) { return Number(v).toFixed(6); })
        .join(",");
    }
    var canvas = qs("#overlayCanvas");
    if (!canvas.width || !canvas.height) return _FULL_FRAME_REGION_STRING;
    return [r.x / canvas.width, r.y / canvas.height, r.w / canvas.width, r.h / canvas.height]
      .map(function (v) { return Number(v).toFixed(6); })
      .join(",");
  }

  // Any region ref (active / stash / full-frame) → coordinate string, or null if unresolved.
  function _regionStringForRef(ref) {
    var r = normalizeRegionRef(ref);
    if (!r) return null;
    if (r.source === "full_frame") return _FULL_FRAME_REGION_STRING;
    var data = null;
    if (r.source === "stash") {
      for (var i = 0; i < state.stashes.length; i++) {
        if (state.stashes[i].id === r.stash_id) {
          data = state.stashes[i].regions[r.name];
          break;
        }
      }
    } else {
      data = state.regions[r.name];
    }
    if (!data) return null;
    return _regionDataToString(data);
  }

  // Preview target: last run-region toggled on, else the active chip. Multitool/boundary hide the picker.
  function _previewRegionRef() {
    if (state.activeWorkflow !== "multitool" && state.activeWorkflow !== "boundary") {
      for (var i = state.runRegions.length - 1; i >= 0; i--) {
        var ref = normalizeRegionRef(state.runRegions[i]);
        if (!ref) continue;
        return ref;
      }
    }
    if (state.activeRegion && state.regions[state.activeRegion]) {
      return activeRegionRef(state.activeRegion);
    }
    return null;
  }

  function _normalizedRegionString() {
    if (state.pendingRegion) {
      var p = state.pendingRegion;
      var c = qs("#overlayCanvas");
      if (!c.width || !c.height) return _FULL_FRAME_REGION_STRING;
      return [p.x / c.width, p.y / c.height, p.w / c.width, p.h / c.height]
        .map(function (v) { return Number(v).toFixed(6); })
        .join(",");
    }
    var regionStr = _regionStringForRef(_previewRegionRef());
    return regionStr || _FULL_FRAME_REGION_STRING;
  }

  function _hasActiveOrPendingRegion() {
    if (state.pendingRegion) return true;
    var ref = _previewRegionRef();
    return !!(ref && ref.source !== "full_frame");
  }

  function _encodeMaskContours(contours) {
    if (!contours || !contours.length) return null;
    return contours
      .map(function (contour) {
        return contour
          .map(function (pt) { return pt[0].toFixed(4) + "," + pt[1].toFixed(4); })
          .join(";");
      })
      .join("|");
  }

  // Bbox-relative contours as "u,v;u,v|…" for mask=, or null for rects. Pending draws are canvas-absolute.
  function _regionMaskString() {
    var contours = null;
    if (state.pendingRegion) {
      var p = state.pendingRegion;
      if (p.points && p.points.length > 0 && p.w > 0 && p.h > 0) {
        contours = p.points.map(function (contour) {
          return contour.map(function (pt) {
            return [(pt[0] - p.x) / p.w, (pt[1] - p.y) / p.h];
          });
        });
      }
    } else {
      var data = _regionObjectForRef(_previewRegionRef());
      if (data && data.points && data.points.length > 0) contours = data.points;
    }
    return _encodeMaskContours(contours);
  }

  // Shaped region but the tool only analyzes its bounding rect (config-mirrored list).
  function _maskFallbackActive() {
    if (CLIPGEN_CONFIG.screenspaceMaskFallbackTools.indexOf(state.activeWorkflow) === -1) {
      return false;
    }
    if (state.pendingRegion) {
      return !!(state.pendingRegion.points && state.pendingRegion.points.length > 0);
    }
    var data = _regionObjectForRef(_previewRegionRef());
    return !!(data && data.points && data.points.length > 0);
  }

  // Stored {x,y,w,h} (frame fractions) for a region ref; null for full-frame / unresolved.
  function _regionObjectForRef(ref) {
    var r = normalizeRegionRef(ref);
    if (!r || r.source === "full_frame") return null;
    if (r.source === "stash") {
      for (var i = 0; i < state.stashes.length; i++) {
        if (state.stashes[i].id === r.stash_id) {
          return state.stashes[i].regions[r.name] || null;
        }
      }
      return null;
    }
    return state.regions[r.name] || null;
  }

  // Source-pixel area the color step analyzes; null when video size is unknown.
  function _colorRegionPixelArea(sfx) {
    var info = state.videoInfo;
    if (!info || !info.width || !info.height) return null;
    var frameArea = info.width * info.height;
    var ref;
    if (sfx && sfx.indexOf("_mt") === 0) {
      var idx = parseInt(sfx.slice(3), 10);
      var step = state.multitoolSteps[idx];
      ref = step ? (step.region_ref || (step.region ? activeRegionRef(step.region) : null)) : null;
    } else if (state.pendingRegion) {
      var c = qs("#overlayCanvas");
      if (!c || !c.width || !c.height) return frameArea;
      var pw = (state.pendingRegion.w / c.width) * info.width;
      var ph = (state.pendingRegion.h / c.height) * info.height;
      return Math.max(1, Math.round(pw * ph));
    } else {
      ref = _previewRegionRef();
    }
    var r = _regionObjectForRef(ref);
    if (!r) return frameArea; // full frame or unresolved
    return Math.max(1, Math.round((r.w * info.width) * (r.h * info.height)));
  }

  // "Min area %" readout: percentage plus approximate pixel count, or the 0% "any presence" note.
  function _updateMinAreaReadout(sfx) {
    sfx = sfx || "";
    var slider = qs("#paramColorMinArea" + sfx);
    var out = qs("#paramColorMinAreaVal" + sfx);
    if (!slider || !out) return;
    out.textContent = _formatMinAreaReadout(
      numberOrDefault(slider.value, 0), _colorRegionPixelArea(sfx)
    );
  }

  function _collectPreviewParams(tool) {
    var out = {};
    if (tool === "color") {
      var c = SS.getColorHiddenInputs();
      if (c) {
        out.h = c.h.value; out.s = c.s.value; out.v = c.v.value;
      }
    } else if (tool === "change") {
      var n = qs("#paramChangeNoise");
      if (n) out.noise = n.value;
    } else if (tool === "flow") {
      var m = qs("#paramFlowMag");
      if (m) out.magnitude = m.value;
    } else if (tool === "text") {
      var tp = qs("#paramTextOcrPreprocess");
      if (tp && tp.checked) out.ocr_preprocess = "1";
    } else if (tool === "numbers") {
      var np = qs("#paramNumOcrPreprocess");
      if (np && np.checked) out.ocr_preprocess = "1";
    } else if (tool === "shape") {
      var st = qs("#paramShapeThresh");
      if (st) out.threshold = st.value;
      var smin = qs("#paramShapeScaleMin");
      if (smin) out.scale_min = (parseFloat(smin.value) || 0) / 100;
      var smax = qs("#paramShapeScaleMax");
      if (smax) out.scale_max = (parseFloat(smax.value) || 0) / 100;
      var sst = qs("#paramShapeSteps");
      if (sst) out.scale_steps = sst.value;
      var slink = qs("#paramShapeLinkAxes");
      if (slink && !slink.checked) {
        var symin = qs("#paramShapeScaleYMin");
        if (symin) out.scale_y_min = (parseFloat(symin.value) || 0) / 100;
        var symax = qs("#paramShapeScaleYMax");
        if (symax) out.scale_y_max = (parseFloat(symax.value) || 0) / 100;
        var systeps = qs("#paramShapeStepsY");
        if (systeps) out.scale_y_steps = systeps.value;
      }
    } else if (tool === "attention") {
      var attnIds = {
        weight_spectral: "paramAttnWSpectral",
        weight_contrast: "paramAttnWContrast",
        weight_motion: "paramAttnWMotion",
        weight_face: "paramAttnWFace",
        center_bias: "paramAttnCenterBias",
      };
      Object.keys(attnIds).forEach(function (key) {
        var input = qs("#" + attnIds[key]);
        if (input) out[key] = input.value;
      });
    }
    return out;
  }

  function _doRefreshModelView() {
    var gen = ++_modelViewGen;
    var meta = qs("#modelViewMeta");
    var img = qs("#modelViewImage");
    if (!meta || !img) return;
    // Clear the shimmer first so every early return starts flat; the fetch branch re-adds it.
    meta.classList.remove("cg-shimmer");

    if (!state.selectedParticipant) {
      meta.textContent = "Select a participant to preview.";
      img.removeAttribute("src");
      return;
    }

    var tool = state.activeWorkflow;
    var regionStr = _normalizedRegionString();
    var hasRegion = _hasActiveOrPendingRegion();

    if (tool === "template" || tool === "shape") {
      var snapRegion = state.capturedRefPreview
        && state.capturedRefPreview.ts === state.referenceTimestamp
        && state.capturedRefPreview.region;
      if (state.uploadedTemplate && state.uploadedTemplate.data) {
        // POST with the upload — region optional
      } else if (state.referenceTimestamp != null) {
        // Shape's sample rides its capture region, so a Full-frame run target still previews.
        if (!hasRegion && !snapRegion) {
          meta.textContent = "Select or draw a region to preview the captured reference.";
          img.removeAttribute("src");
          return;
        }
      } else {
        meta.textContent = "Capture a reference region or upload a PNG to preview.";
        img.removeAttribute("src");
        return;
      }
    }

    var params = _collectPreviewParams(tool);
    var qsParts = ["tool=" + encodeURIComponent(tool)];
    if (regionStr) qsParts.push("region=" + regionStr);
    var maskStr = _regionMaskString();
    if (regionStr && maskStr) qsParts.push("mask=" + encodeURIComponent(maskStr));
    if (tool === "change" || tool === "flow" || tool === "attention") {
      // Attention's motion channel compares at its own sampling interval.
      var prevGap = 1;
      if (tool === "attention") {
        var attnInterval = qs("#paramAttnInterval");
        prevGap = attnInterval ? (parseFloat(attnInterval.value) || 0.5) : 0.5;
      }
      var prevTs = Math.max(0, (state.currentTimestamp || 0) - prevGap);
      qsParts.push("prev=" + prevTs.toFixed(3));
    }
    if (tool === "similarity" && state.referenceTimestamp != null) {
      qsParts.push("ref=" + Number(state.referenceTimestamp).toFixed(3));
    }
    if (
      (tool === "template" || tool === "shape") &&
      state.referenceTimestamp != null &&
      !(state.uploadedTemplate && state.uploadedTemplate.data)
    ) {
      qsParts.push("ref=" + Number(state.referenceTimestamp).toFixed(3));
      if (snapRegion) {
        var capRect = (state.previewRegions || state.regions)[snapRegion];
        if (capRect) {
          qsParts.push("ref_region=" + [capRect.x, capRect.y, capRect.w, capRect.h]
            .map(function (v) { return Number(v).toFixed(6); }).join(","));
          var capMask = _encodeMaskContours(capRect.points);
          if (capMask) qsParts.push("ref_mask=" + encodeURIComponent(capMask));
        }
      }
    }
    Object.keys(params).forEach(function (k) {
      qsParts.push(encodeURIComponent(k) + "=" + encodeURIComponent(params[k]));
    });
    qsParts.push("_=" + gen);

    var ts = Number(state.currentTimestamp || 0).toFixed(3);
    var url = "api/preview/" + encodeURIComponent(state.selectedParticipant)
      + "/" + ts + "?" + qsParts.join("&");

    meta.classList.add("cg-shimmer");
    meta.textContent = "Loading preview…";

    function applyPreviewError() {
      if (gen !== _modelViewGen) return;
      meta.classList.remove("cg-shimmer");
      meta.textContent = "Preview unavailable.";
      img.removeAttribute("src");
    }

    function applyPreviewOkFromBlob(blob) {
      if (gen !== _modelViewGen) return;
      if (img._modelViewObjectUrl) {
        URL.revokeObjectURL(img._modelViewObjectUrl);
        img._modelViewObjectUrl = null;
      }
      var u = URL.createObjectURL(blob);
      img._modelViewObjectUrl = u;
      img.src = u;
      meta.classList.remove("cg-shimmer");
      var metaText = MODEL_VIEW_META[tool] || "";
      if (!hasRegion) {
        metaText = (metaText ? metaText + " " : "") + "(Full frame — no region selected.)";
      } else if (_maskFallbackActive()) {
        metaText = (metaText ? metaText + " " : "")
          + "(Shaped region: this tool analyzes the bounding box.)";
      }
      meta.textContent = metaText;
    }

    function _refetchOverlayLayer() {
      if (gen !== _modelViewGen) return;
      var resolved = _resolveOverlayLayer();
      if (!resolved) {
        if (state.overlayImageObjectUrl) {
          URL.revokeObjectURL(state.overlayImageObjectUrl);
          state.overlayImageObjectUrl = null;
        }
        state.overlayImage = null;
        state.overlayImageTimestamp = null;
        state.overlayImageTool = null;
        return;
      }
      var layerQs = qsParts.concat(["layer=" + encodeURIComponent(resolved.id)]);
      var layerUrl = "api/preview/" + encodeURIComponent(state.selectedParticipant)
        + "/" + ts + "?" + layerQs.join("&");
      function fetchAsImage(blob) {
        if (gen !== _modelViewGen) return;
        if (state.overlayImageObjectUrl) {
          URL.revokeObjectURL(state.overlayImageObjectUrl);
          state.overlayImageObjectUrl = null;
        }
        var ou = URL.createObjectURL(blob);
        state.overlayImageObjectUrl = ou;
        var oi = new Image();
        oi.onload = function () {
          if (gen !== _modelViewGen) return;
          state.overlayImage = oi;
          state.overlayImageScope = resolved.scope;
          state.overlayImageTimestamp = ts;
          state.overlayImageTool = tool;
          SS.renderOverlay();
        };
        oi.src = ou;
      }
      if (useTemplatePost) {
        apiPostBlob(layerUrl, uploadPostBody())
          .then(fetchAsImage)
          .catch(function () { /* leave previous overlay image */ });
      } else {
        apiGetBlob(layerUrl)
          .then(fetchAsImage)
          .catch(function () { /* leave previous overlay image */ });
      }
    }

    var useTemplatePost = (tool === "template" || tool === "shape")
      && state.uploadedTemplate && state.uploadedTemplate.data;
    // Shape uploads ride their own field so the server routes them correctly.
    function uploadPostBody() {
      var body = {};
      body[tool === "shape" ? "shape_image_data" : "template_image_data"] = state.uploadedTemplate.data;
      return body;
    }
    if (useTemplatePost) {
      apiPostBlob(url, uploadPostBody())
        .then(function (blob) {
          if (gen !== _modelViewGen) return;
          applyPreviewOkFromBlob(blob);
          _refetchOverlayLayer();
        })
        .catch(function () {
          applyPreviewError();
        });
      return;
    }

    var tmp = new Image();
    tmp.onload = function () {
      if (gen !== _modelViewGen) return;
      if (img._modelViewObjectUrl) {
        URL.revokeObjectURL(img._modelViewObjectUrl);
        img._modelViewObjectUrl = null;
      }
      img.src = tmp.src;
      meta.classList.remove("cg-shimmer");
      meta.textContent = MODEL_VIEW_META[tool] || "";
      _refetchOverlayLayer();
    };
    tmp.onerror = function () {
      applyPreviewError();
    };
    tmp.src = url;
  }

  // ---- Published to the hub; later satellites destructure the helpers at load ----
  SS.initModelView = initModelView;
  SS.cycleOverlayLayer = cycleOverlayLayer;
  SS.refreshModelView = refreshModelView;
  SS._updateOverlayUi = _updateOverlayUi;
  SS._overlayEligibleForActiveTool = _overlayEligibleForActiveTool;
  SS._updateMinAreaReadout = _updateMinAreaReadout;
  SS._previewRegionRef = _previewRegionRef;
})();
