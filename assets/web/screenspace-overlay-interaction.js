/* clipgen Screenspace — region draw/drag/resize interaction satellite.
 *
 * Carved out of screenspace.js (the hub) following the hub+satellite convention
 * (see screenspace-overlay/timeline/model-view/...). Owns the region-editor's
 * pointer interaction: the overlay-canvas mousedown/move/up state machine (new-
 * region draw via rect / freehand lasso / magic-wand flood fill, region
 * move/resize, template drag, pipette sampling), the region tool toggle, the
 * region chips + toolbar buttons, the region-name modal, saveRegionUpdate, and
 * the overlay-rect cache + render RAF. Shaped regions carry bbox-relative
 * `points` (+ `shape`); geometry helpers (simplifyPolygon, floodFillMask,
 * traceMaskContour, polygonBounds/Area) are screenspace-utils.js globals.
 *
 * Reads the hub's shared `state` and a set of hub + model-view helpers through
 * window.ClipgenScreenspace (SS), all published before this file loads (the hub
 * loads first, then screenspace-model-view.js). renderOverlay
 * (screenspace-overlay.js) and setTargetColor (screenspace-color.js) load AFTER
 * this file, so they are reached late-bound through thin local wrappers. qs/el/
 * clamp/normalizeRect/showToast/apiPost/apiDelete/rgbToHsv/rgbToHex are ambient
 * utils.js / screenspace-utils.js globals.
 *
 * Load order: right after screenspace-model-view.js and BEFORE
 * screenspace-overlay.js (destructures SS.computeLabelRect),
 * screenspace-timeline.js (SS.invalidateOverlayRect) and screenspace-tasks.js
 * (SS.renderRegionChips / SS.updateRegionButtons). The hub keeps same-named
 * delegators for initRegionDrawing / renderRegionChips / updateRegionButtons /
 * hideRegionNameModal / invalidateOverlayRect.
 */

(function () {
  "use strict";

  var SS = window.ClipgenScreenspace;
  var state = SS.state;
  // Hub + model-view helpers, published before this file loads.
  var regionToPixels = SS.regionToPixels,
    regionColorForIndex = SS.regionColorForIndex,
    iconSpan = SS.iconSpan,
    templateOverlayBounds = SS.templateOverlayBounds,
    renderRunRegionPicker = SS.renderRunRegionPicker,
    updateRunButton = SS.updateRunButton,
    renderWorkflowParams = SS.renderWorkflowParams,
    getThemeColors = SS.getThemeColors,
    deactivatePipette = SS.deactivatePipette,
    refreshCalibration = SS.refreshCalibration,
    refreshModelView = SS.refreshModelView,
    _updateMinAreaReadout = SS._updateMinAreaReadout,
    pauseVideo = SS.pauseVideo,
    stashRegions = SS.stashRegions,
    pinCurrentFrame = SS.pinCurrentFrame,
    togglePinTrayVisibility = SS.togglePinTrayVisibility,
    clearAllPins = SS.clearAllPins,
    updatePinButtons = SS.updatePinButtons;

  // renderOverlay (screenspace-overlay.js) and setTargetColor
  // (screenspace-color.js) load AFTER this file — reach them late-bound.
  function renderOverlay() { return SS.renderOverlay && SS.renderOverlay.apply(null, arguments); }
  function setTargetColor(h, s, v) { return SS.setTargetColor && SS.setTargetColor(h, s, v); }
  function updateColorSampleBtnLabel() { return SS.updateColorSampleBtnLabel && SS.updateColorSampleBtnLabel(); }

  // Overlay-rect cache + render RAF; the timeline satellite drops it via SS.invalidateOverlayRect.
  var _overlayRaf = 0;
  var _cachedOverlayRect = null;

  // Magic-wand scrub state: frame ImageData read once at press, coalescing RAF.
  var _wandRaf = 0;
  var _wandFrame = null;
  var WAND_SCRUB_SENSITIVITY = 0.4; // tolerance units per horizontal px

  function invalidateOverlayRect() {
    _cachedOverlayRect = null;
  }

  // ---- Region drawing ----

  function canvasCoords(canvas, event, cachedRect) {
    var rect = cachedRect || canvas.getBoundingClientRect();
    var scaleX = canvas.width / rect.width;
    var scaleY = canvas.height / rect.height;
    return {
      x: Math.round((event.clientX - rect.left) * scaleX),
      y: Math.round((event.clientY - rect.top) * scaleY),
    };
  }

  function scheduleOverlayRender() {
    if (_overlayRaf) return;
    _overlayRaf = requestAnimationFrame(function () {
      _overlayRaf = 0;
      renderOverlay();
    });
  }

  function flushOverlayRender() {
    if (_overlayRaf) {
      cancelAnimationFrame(_overlayRaf);
      _overlayRaf = 0;
    }
    renderOverlay();
  }

  function computeLabelRect(r, name, ctx, s) {
    var fontSize = Math.round(12 * s);
    var labelH = Math.round(16 * s);
    var pad = Math.round(4 * s);
    var gripW = Math.round(8 * s);
    var gripPadL = Math.round(3 * s);
    ctx.font = "bold " + fontSize + "px -apple-system, BlinkMacSystemFont, sans-serif";
    var textW = ctx.measureText(name).width;
    var labelW = gripPadL + gripW + pad + textW + pad;
    return { x: r.x, y: r.y - labelH, w: labelW, h: labelH, gripPadL: gripPadL, gripW: gripW, pad: pad, fontSize: fontSize };
  }

  function findHitRegion(px, py, s, ctx) {
    if (!state.showRegionOverlays) return null;
    var names = Object.keys(state.regions);
    var handleSize = Math.round(14 * s);
    for (var i = names.length - 1; i >= 0; i--) {
      var name = names[i];
      var r = regionToPixels(state.regions[name]);
      if (px >= r.x + r.w - handleSize && px <= r.x + r.w && py >= r.y + r.h - handleSize && py <= r.y + r.h) {
        return { name: name, handle: "resize" };
      }
      if (state.showRegionLabels) {
        var lr = computeLabelRect(r, name, ctx, s);
        if (px >= lr.x && px <= lr.x + lr.w && py >= lr.y && py <= lr.y + lr.h) {
          return { name: name, handle: "move" };
        }
      }
    }
    return null;
  }

  // Saved region's shape in absolute canvas pixels: {contours} (stored ones are bbox-relative) or {rect}.
  function regionShapeAbs(region) {
    var px = regionToPixels(region);
    if (region.points && region.points.length > 0) {
      return {
        contours: region.points.map(function (contour) {
          return contour.map(function (p) {
            return [px.x + p[0] * px.w, px.y + p[1] * px.h];
          });
        }),
      };
    }
    return { rect: px };
  }

  function saveRegionUpdate(name) {
    var region = state.regions[name];
    if (!region) return;
    var canvas = qs("#overlayCanvas");
    var px = regionToPixels(region);
    var body = { name: name, x: px.x, y: px.y, w: px.w, h: px.h, canvas_width: canvas.width, canvas_height: canvas.height };
    if (region.points && region.points.length > 0) {
      // Stored contours are bbox-relative; the API takes canvas-pixel absolutes and recomputes the bbox.
      body.points = regionShapeAbs(region).contours;
      body.shape = region.shape || "lasso";
    }
    if (region.description) body.description = region.description;
    apiPost("api/regions", body)
      .then(function (data) {
        if (data.ok) {
          var saved = data.region;
          if (region.description) saved.description = region.description;
          state.regions[name] = saved;
          renderOverlay();
          refreshCalibration({ debounce: true });
        }
      })
      .catch(function () { showToast("Failed to update region"); });
  }

  // Saves a boolean-edit or merge result as "combo"; an axis-aligned result becomes an unmasked rect.
  function saveRegionShape(name, contours, successToast, onDone) {
    var canvas = qs("#overlayCanvas");
    var region = state.regions[name] || {};
    var rect = contoursToAxisRect(contours, 2);
    var body = rect
      ? { name: name, x: rect.x, y: rect.y, w: rect.w, h: rect.h }
      : (function () {
          var b = contoursBounds(contours);
          return { name: name, x: b.x, y: b.y, w: b.w, h: b.h, points: contours, shape: "combo" };
        })();
    body.canvas_width = canvas.width;
    body.canvas_height = canvas.height;
    if (region.description) body.description = region.description;
    // Always resolves (the catch swallows failures) so callers can chain saves of one region.
    return apiPost("api/regions", body)
      .then(function (data) {
        if (!data.ok) {
          showToast(data.error || "Failed to update region");
          return;
        }
        var saved = data.region;
        if (region.description) saved.description = region.description;
        state.regions[name] = saved;
        renderOverlay();
        refreshModelView({ debounce: true });
        refreshCalibration({ debounce: true });
        if (successToast) showToast(successToast);
        if (onDone) onDone();
      })
      .catch(function () { showToast("Failed to update region"); });
  }

  // Pending shaped region, or null when under finishDrawingRegion's 5x5 minimum or the area floor.
  function pendingShapedRegion(contours, shape) {
    contours = contours.filter(function (c) { return c.length >= 3; });
    if (!contours.length) return null;
    var bounds = contoursBounds(contours);
    var x = Math.floor(bounds.x);
    var y = Math.floor(bounds.y);
    var w = Math.ceil(bounds.x + bounds.w) - x; // containing integer box, so normalized points stay in [0, 1]
    var h = Math.ceil(bounds.y + bounds.h) - y;
    if (w <= 5 || h <= 5 || contoursArea(contours) < 64) return null;
    return { x: x, y: y, w: w, h: h, points: contours, shape: shape };
  }

  var COMBINE_VERBS = { add: "Added to", subtract: "Subtracted from", intersect: "Intersected" };

  // Rasterize base ∪/∖/∩ shape and trace back to contours; null when nothing survives.
  function combineToContours(baseShape, shape, op) {
    var overlay = qs("#overlayCanvas");
    if (!overlay.width || !overlay.height) return null;
    var w = overlay.width, h = overlay.height;
    var mask = rasterizeShapesMask([baseShape], w, h);
    combineShapeMasks(mask, rasterizeShapesMask([shape], w, h), op);
    var displayW = overlay.getBoundingClientRect().width || w;
    var s = w / displayW;
    // Drop <8px² specks, then simplify toward the server's 400-vertex cap.
    var contours = simplifyContours(maskToContours(mask, w, h, 8), 2 * s, 400);
    if (!contours.length || contoursArea(contours) < 64) return null;
    return contours;
  }

  // Per-region promise chain: a second combine within one round-trip must read the first's saved base.
  var _regionCombineChain = {};

  // Boolean-edit the active or unsaved pending region with a drawn shape: shift=add, alt=subtract, both=intersect.
  function applyRegionCombine(op, shape) {
    var name = state.activeRegion;
    if (name && state.regions[name]) {
      var run = function () { return combineIntoRegion(op, shape, name); };
      var prev = _regionCombineChain[name];
      _regionCombineChain[name] = prev ? prev.then(run) : run();
      return;
    }
    combineIntoPending(op, shape);
  }

  function combineIntoRegion(op, shape, name) {
    // Re-read now: a queued combine may have replaced it, or the region was deleted meanwhile.
    var region = state.regions[name];
    if (!region) return Promise.resolve();
    var contours = combineToContours(regionShapeAbs(region), shape, op);
    if (!contours) {
      showToast("Nothing would remain of '" + name + "', so this was not applied");
      return Promise.resolve();
    }
    return saveRegionShape(name, contours, COMBINE_VERBS[op] + " region '" + name + "'");
  }

  function combineIntoPending(op, shape) {
    var pending = state.pendingRegion;
    if (!pending) return;
    // Pending contours are already canvas-pixel absolute (unlike a saved
    // region's bbox-relative points, which regionShapeAbs() denormalizes).
    var baseShape = pending.points && pending.points.length > 0
      ? { contours: pending.points }
      : { rect: { x: pending.x, y: pending.y, w: pending.w, h: pending.h } };
    var contours = combineToContours(baseShape, shape, op);
    // In place, no round-trip; axis-aligned results collapse to a rect like saveRegionShape.
    var rect = contours ? contoursToAxisRect(contours, 2) : null;
    var updated = !contours
      ? null
      : rect
        ? { x: rect.x, y: rect.y, w: rect.w, h: rect.h }
        : pendingShapedRegion(contours, "combo");
    if (!updated) {
      showToast("Nothing would remain, so this was not applied");
      return;
    }
    state.pendingRegion = updated;
    renderOverlay();
    updateRegionButtons();
    showToast(COMBINE_VERBS[op] + " region");
  }

  // ---- Overlay interaction state machine ----
  // Document-level listeners keep drags alive off-canvas.
  function initRegionDrawing() {
    var overlay = qs("#overlayCanvas");

    function finishDrawingRegion(e) {
      if (!state.drawingRegion) return false;
      var rect = _cachedOverlayRect || overlay.getBoundingClientRect();
      var pos = canvasCoords(overlay, e, rect);
      pos.x = clamp(pos.x, 0, overlay.width);
      pos.y = clamp(pos.y, 0, overlay.height);
      state.drawingRegion.endX = pos.x;
      state.drawingRegion.endY = pos.y;
      var combine = state.drawingRegion.combine;
      var r = normalizeRect(
        state.drawingRegion.startX, state.drawingRegion.startY,
        state.drawingRegion.endX, state.drawingRegion.endY
      );
      state.drawingRegion = null;
      _cachedOverlayRect = null;
      if (r.w > 5 && r.h > 5) {
        if (combine) applyRegionCombine(combine, { rect: r });
        else state.pendingRegion = r;
      } else if (r.w > 0 || r.h > 0) {
        // Silent for a click without drag; a dragged too-small box gets a hint.
        showToast("Region too small. Drag a larger area");
      }
      flushOverlayRender();
      updateRegionButtons();
      return true;
    }

    // Append a point once the cursor moved ~3 display px, so the trail stays bounded.
    function appendLassoPoint(pos, s) {
      var pts = state.drawingLasso.points;
      var last = pts[pts.length - 1];
      var x = clamp(pos.x, 0, overlay.width);
      var y = clamp(pos.y, 0, overlay.height);
      var dx = x - last[0], dy = y - last[1];
      if (dx * dx + dy * dy >= 9 * s * s) {
        pts.push([x, y]);
        scheduleOverlayRender();
      }
    }

    // Simplify a raw point trail toward <=100 vertices with growing epsilon.
    function simplifyForRegion(pts, s) {
      var simplified = simplifyPolygon(pts, 2 * s);
      var epsilon = 2 * s;
      while (simplified.length > 100) {
        epsilon *= 1.5;
        simplified = simplifyPolygon(simplified, epsilon);
      }
      return simplified;
    }

    // Close and simplify the trail into a pending polygon, or boolean-apply a combine draw.
    function finishDrawingLasso() {
      if (!state.drawingLasso) return false;
      var pts = state.drawingLasso.points;
      var combine = state.drawingLasso.combine;
      state.drawingLasso = null;
      _cachedOverlayRect = null;
      var displayW = overlay.getBoundingClientRect().width || overlay.width;
      var s = overlay.width / displayW;
      var simplified = simplifyForRegion(pts, s);
      if (combine) {
        if (simplified.length >= 3 && polygonArea(simplified) >= 8) {
          applyRegionCombine(combine, { contours: [simplified] });
        } else if (pts.length > 2) {
          showToast("Shape too small. Draw a larger area");
        }
        flushOverlayRender();
        updateRegionButtons();
        return true;
      }
      var pending = pendingShapedRegion([simplified], "lasso");
      if (pending) {
        state.pendingRegion = pending;
      } else if (pts.length > 2) {
        showToast("Shape too small. Draw a larger area");
      }
      flushOverlayRender();
      updateRegionButtons();
      return true;
    }

    // Wand: press caches frame pixels + seed (canvases share dimensions), drag scrubs tolerance, release commits.

    // Flood the cached frame at the scrub tolerance; stash the outer contour for the painter.
    function computeWandPreview() {
      var f = _wandFrame;
      if (!f || !state.wandDragging) return;
      var mask = floodFillMask(f.data, f.w, f.h, f.seedX, f.seedY, state.wandDragging.tolerance);
      var pts = mask ? simplifyForRegion(traceMaskContour(mask, f.w, f.h), f.s) : [];
      state.wandDragging.previewPoints = pts.length >= 3 ? [pts] : null;
    }

    // floodFillMask is O(w*h); coalesce the many per-frame mousemoves onto one RAF.
    function scheduleWandRecompute() {
      if (_wandRaf) return;
      _wandRaf = requestAnimationFrame(function () {
        _wandRaf = 0;
        if (!state.wandDragging) { _wandFrame = null; return; }
        computeWandPreview();
        renderOverlay();
      });
    }

    function beginWandDrag(e, pos, s, combine) {
      var frameCanvas = qs("#frameCanvas");
      // Gate on frameImage, not canvas width: an unsized canvas reports 300x150 and would flood blank.
      if (!frameCanvas || !state.frameImage) {
        flushOverlayRender(); // caller cleared the region; repaint or the UI goes stale
        renderRegionChips();
        updateRegionButtons();
        return;
      }
      var w = frameCanvas.width, h = frameCanvas.height;
      _wandFrame = {
        data: frameCanvas.getContext("2d").getImageData(0, 0, w, h).data,
        w: w, h: h, seedX: pos.x, seedY: pos.y, s: s,
      };
      // seedX/seedY/headOffsetPx feed the painter's drag chrome; it can't see _wandFrame.
      state.wandDragging = {
        startClientX: e.clientX,
        startTolerance: state.wandTolerance,
        tolerance: state.wandTolerance,
        combine: combine || null,
        previewPoints: null,
        seedX: pos.x,
        seedY: pos.y,
        headOffsetPx: 0,
      };
      // Set on the overlay too; its crosshair rule and inline hover cursor beat body's.
      overlay.style.cursor = "ew-resize";
      document.body.style.cursor = "ew-resize";
      document.body.style.userSelect = "none";
      computeWandPreview();
      flushOverlayRender();
      updateRegionButtons();
    }

    function updateWandDragFromEvent(e) {
      if (!state.wandDragging) return;
      var inp = qs("#wandToleranceInput");
      var lo = parseInt(inp.min, 10) || 4;
      var hi = parseInt(inp.max, 10) || 120;
      var delta = e.clientX - state.wandDragging.startClientX;
      var tol = clamp(Math.round(state.wandDragging.startTolerance + delta * WAND_SCRUB_SENSITIVITY), lo, hi);
      state.wandDragging.tolerance = tol;
      // Clamped tolerance so the track stops at min/max; CSS px so the painter scales live.
      var wd = state.wandDragging;
      wd.headOffsetPx = (tol - wd.startTolerance) / WAND_SCRUB_SENSITIVITY;
      state.wandTolerance = tol;
      inp.value = String(tol);
      qs("#wandToleranceValue").textContent = String(tol);
      scheduleWandRecompute();
    }

    function endWandDrag() {
      if (_wandRaf) { cancelAnimationFrame(_wandRaf); _wandRaf = 0; }
      // The hover pass re-establishes the overlay cursor on the next move.
      overlay.style.cursor = "";
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      _wandFrame = null;
      _cachedOverlayRect = null;
    }

    function commitWandDrag() {
      var wd = state.wandDragging;
      if (!wd) return;
      // A quick release can beat the RAF; re-flood synchronously so the commit matches the readout.
      if (_wandRaf) { cancelAnimationFrame(_wandRaf); _wandRaf = 0; }
      computeWandPreview();
      var contour = wd.previewPoints && wd.previewPoints[0];
      var combine = wd.combine;
      endWandDrag();
      state.wandDragging = null;
      if (combine) {
        if (contour && contour.length >= 3 && polygonArea(contour) >= 8) {
          applyRegionCombine(combine, { contours: [contour] });
        } else {
          showToast("No contiguous area found. Adjust tolerance and try again");
        }
      } else {
        var pending = contour ? pendingShapedRegion([contour], "wand") : null;
        if (pending) state.pendingRegion = pending;
        else showToast("No contiguous area found. Adjust tolerance and try again");
      }
      flushOverlayRender();
      updateRegionButtons();
    }

    function cancelWandDrag() {
      if (!state.wandDragging) return;
      endWandDrag();
      state.wandDragging = null;
      flushOverlayRender();
      updateRegionButtons();
    }
    // Published here since the wand cluster is local; Escape and the frame loader abort scrubs.
    SS.cancelWandDrag = cancelWandDrag;

    // ---- Shape-draw mode ----
    // Strokes fill an offscreen mask; Apply uploads its bbox crop.

    // The draw button lives in the re-rendered params panel; sync by class.
    function syncDrawButton() {
      var btn = qs(".ss-template-icon-btn--draw");
      if (btn) btn.classList.toggle("active", !!state.shapeDraw);
    }

    function toggleShapeDraw() {
      if (state.shapeDraw) { cancelShapeDraw(); return; }
      if (!state.frameImage) { showToast("Load a frame first"); return; }
      var mask = document.createElement("canvas");
      mask.width = overlay.width;
      mask.height = overlay.height;
      state.shapeDraw = {
        canvas: mask, ctx: mask.getContext("2d"),
        stroking: false, erasing: false, lastX: 0, lastY: 0,
      };
      qs("#shapeDrawWrap").classList.remove("collapsed");
      syncDrawButton();
      flushOverlayRender();
      showToast("Paint over the shape — Shift-drag erases");
    }

    function cancelShapeDraw() {
      if (!state.shapeDraw) return;
      state.shapeDraw = null; // releases the mask canvas
      qs("#shapeDrawWrap").classList.add("collapsed");
      syncDrawButton();
      flushOverlayRender();
    }
    // Published like cancelWandDrag: Escape, the frame loader, and tool switches abort sessions.
    SS.toggleShapeDraw = toggleShapeDraw;
    SS.cancelShapeDraw = cancelShapeDraw;

    function beginShapeStroke(e, pos, s) {
      var sd = state.shapeDraw;
      sd.stroking = true;
      sd.erasing = e.shiftKey;
      var mctx = sd.ctx;
      mctx.globalCompositeOperation = sd.erasing ? "destination-out" : "source-over";
      // Opaque accent strokes; the painter applies the highlighter alpha.
      mctx.strokeStyle = getThemeColors().accent;
      mctx.fillStyle = mctx.strokeStyle;
      mctx.lineCap = "round";
      mctx.lineJoin = "round";
      mctx.lineWidth = Math.max(1, state.shapeBrushSize * s);
      mctx.beginPath();
      mctx.arc(pos.x, pos.y, mctx.lineWidth / 2, 0, Math.PI * 2);
      mctx.fill();
      sd.lastX = pos.x;
      sd.lastY = pos.y;
      scheduleOverlayRender();
    }

    function extendShapeStroke(pos) {
      var sd = state.shapeDraw;
      var mctx = sd.ctx;
      mctx.beginPath();
      mctx.moveTo(sd.lastX, sd.lastY);
      mctx.lineTo(pos.x, pos.y);
      mctx.stroke();
      sd.lastX = pos.x;
      sd.lastY = pos.y;
      scheduleOverlayRender();
    }

    function endShapeStroke() {
      if (state.shapeDraw) state.shapeDraw.stroking = false;
    }

    // Crop the painted bbox from the frame; hand it to the upload path.
    function commitShapeDraw() {
      var sd = state.shapeDraw;
      if (!sd) return;
      var w = sd.canvas.width, h = sd.canvas.height;
      var data = sd.ctx.getImageData(0, 0, w, h).data;
      var minX = w, minY = h, maxX = -1, maxY = -1;
      for (var y = 0; y < h; y++) {
        var rowBase = y * w * 4;
        for (var x = 0; x < w; x++) {
          if (data[rowBase + x * 4 + 3] > 0) {
            if (x < minX) minX = x;
            if (x > maxX) maxX = x;
            if (y < minY) minY = y;
            if (y > maxY) maxY = y;
          }
        }
      }
      // The backend erodes the alpha and needs enough surviving edge pixels.
      if (maxX - minX < 16 || maxY - minY < 16) {
        showToast("Painted shape too small. Draw a larger area");
        return;
      }
      var bw = maxX - minX + 1, bh = maxY - minY + 1;
      var crop = document.createElement("canvas");
      crop.width = bw;
      crop.height = bh;
      var cctx = crop.getContext("2d");
      cctx.drawImage(state.frameImage, minX, minY, bw, bh, 0, 0, bw, bh);
      cctx.globalCompositeOperation = "destination-in";
      cctx.drawImage(sd.canvas, minX, minY, bw, bh, 0, 0, bw, bh);
      var dataUrl = crop.toDataURL("image/png");
      state.uploadedTemplate = { name: "drawn-shape.png", data: dataUrl.split(",")[1] };
      state.referenceTimestamp = null;
      state.capturedRefPreview = null;
      state.templateOverlayPos = null;
      var previewImg = new Image();
      previewImg.onload = function () { renderOverlay(); };
      previewImg.src = dataUrl;
      state.uploadedTemplateImg = previewImg;
      cancelShapeDraw();
      renderWorkflowParams();
      updateRunButton();
      refreshModelView({ debounce: true });
      showToast("Shape captured from drawing");
    }

    qs("#shapeDrawApplyBtn").addEventListener("click", commitShapeDraw);
    qs("#shapeDrawCancelBtn").addEventListener("click", cancelShapeDraw);

    overlay.addEventListener("mousedown", function (e) {
      if (e.button !== 0) return;
      if (state.videoPlaying) pauseVideo();
      if (state.pipetteActive) {
        var frameCanvas = qs("#frameCanvas");
        if (!frameCanvas) { deactivatePipette(); return; }
        var pos = canvasCoords(frameCanvas, e);
        var frameCtx = frameCanvas.getContext("2d");
        var pixel = frameCtx.getImageData(pos.x, pos.y, 1, 1).data;
        var hsv = rgbToHsv(pixel[0], pixel[1], pixel[2]);
        if (state._mtPipetteStep !== undefined && state._mtPipetteStep >= 0) {
          var mtIdx = state._mtPipetteStep;
          var mtSfx = "_mt" + mtIdx;
          var mtHEl = qs("#paramColorH" + mtSfx);
          var mtSEl = qs("#paramColorS" + mtSfx);
          var mtVEl = qs("#paramColorV" + mtSfx);
          if (mtHEl) mtHEl.value = clamp(Math.round(hsv.h), 0, 180);
          if (mtSEl) mtSEl.value = clamp(Math.round(hsv.s), 0, 255);
          if (mtVEl) mtVEl.value = clamp(Math.round(hsv.v), 0, 255);
          var mtHex = qs("#paramColorHex" + mtSfx);
          if (mtHex) mtHex.value = rgbToHex(pixel[0], pixel[1], pixel[2]);
          state._mtPipetteStep = -1;
          // Programmatic sets fire no DOM event, so nudge calibration directly (as setTargetColor does).
          refreshCalibration({ debounce: true });
        } else {
          setTargetColor(hsv.h, hsv.s, hsv.v);
        }
        deactivatePipette();
        showToast("Sampled color from frame");
        return;
      }
      _cachedOverlayRect = overlay.getBoundingClientRect();
      pos = canvasCoords(overlay, e, _cachedOverlayRect);
      var displayW = _cachedOverlayRect.width || overlay.width;
      // Hidden overlay: no size, and `s` would go NaN into hit tests and stored coords.
      if (!displayW || !overlay.width) return;
      var s = overlay.width / displayW;
      var ctx = overlay.getContext("2d");
      // Shape-draw mode swallows every canvas press: paint, or Shift-erase.
      if (state.shapeDraw) {
        beginShapeStroke(e, pos, s);
        return;
      }
      // Shift/alt with a target (active saved region, else the pending one) boolean-edits it; see applyRegionCombine.
      var combineBase =
        state.activeRegion && state.regions[state.activeRegion]
          ? "active"
          : state.pendingRegion
            ? "pending"
            : null;
      if ((e.shiftKey || e.altKey) && combineBase) {
        var combineOp = e.shiftKey && e.altKey ? "intersect" : (e.altKey ? "subtract" : "add");
        // Keep the pending region when it's the target; drop it when editing a saved one.
        if (combineBase === "active") state.pendingRegion = null;
        if (state.regionTool === "wand") {
          beginWandDrag(e, pos, s, combineOp);
          return;
        }
        if (state.regionTool === "lasso") {
          state.drawingLasso = { points: [[pos.x, pos.y]], combine: combineOp };
        } else {
          state.drawingRegion = { startX: pos.x, startY: pos.y, endX: pos.x, endY: pos.y, combine: combineOp };
        }
        updateRegionButtons();
        return;
      }
      var tHit = templateOverlayBounds();
      if (tHit
          && pos.x >= tHit.x && pos.x <= tHit.x + tHit.w
          && pos.y >= tHit.y && pos.y <= tHit.y + tHit.h) {
        state.draggingTemplate = { offsetX: pos.x - tHit.x, offsetY: pos.y - tHit.y };
        document.body.style.cursor = "grabbing";
        document.body.style.userSelect = "none";
        scheduleOverlayRender();
        return;
      }
      var hit = findHitRegion(pos.x, pos.y, s, ctx);
      if (hit && hit.handle === "resize") {
        var origR = state.regions[hit.name];
        state.resizingRegion = { name: hit.name, origRegion: { x: origR.x, y: origR.y, w: origR.w, h: origR.h } };
        state.activeRegion = hit.name;
        state.pendingRegion = null;
        document.body.style.cursor = "nwse-resize";
        document.body.style.userSelect = "none";
        renderRegionChips();
        updateRegionButtons();
        return;
      }
      if (hit && hit.handle === "move") {
        var r = regionToPixels(state.regions[hit.name]);
        var origM = state.regions[hit.name];
        state.draggingRegion = { name: hit.name, offsetX: pos.x - r.x, offsetY: pos.y - r.y, origRegion: { x: origM.x, y: origM.y, w: origM.w, h: origM.h } };
        state.activeRegion = hit.name;
        state.pendingRegion = null;
        document.body.style.cursor = "grabbing";
        document.body.style.userSelect = "none";
        renderRegionChips();
        updateRegionButtons();
        return;
      }
      state.pendingRegion = null;
      state.activeRegion = null;
      if (state.regionTool === "wand") {
        beginWandDrag(e, pos, s, null);
        return;
      }
      if (state.regionTool === "lasso") {
        state.drawingLasso = { points: [[pos.x, pos.y]] };
      } else {
        state.drawingRegion = { startX: pos.x, startY: pos.y, endX: pos.x, endY: pos.y };
      }
      updateRegionButtons();
    });

    overlay.addEventListener("mousemove", function (e) {
      if (state.pipetteActive) return;
      if (state.wandDragging) { updateWandDragFromEvent(e); return; }
      var rect = _cachedOverlayRect || overlay.getBoundingClientRect();
      var pos = canvasCoords(overlay, e, rect);
      var displayW = rect.width || overlay.width;
      var s = overlay.width / displayW;
      if (state.shapeDraw) {
        if (state.shapeDraw.stroking) {
          extendShapeStroke(pos);
        } else {
          // Keep region affordances dead while draw mode is armed.
          state.hoveredRegion = null;
          overlay.style.cursor = "crosshair";
        }
        return;
      }
      if (state.draggingTemplate) {
        var tImg = state.uploadedTemplateImg;
        var tw = Math.max(1, Math.round(tImg.naturalWidth * (state.templateScalePreview || 1.0)));
        var thh = Math.max(1, Math.round(tImg.naturalHeight * (state.templateScalePreview || 1.0)));
        var nx = clamp(pos.x - state.draggingTemplate.offsetX, 0, overlay.width - tw);
        var ny = clamp(pos.y - state.draggingTemplate.offsetY, 0, overlay.height - thh);
        state.templateOverlayPos = { x: nx, y: ny };
        scheduleOverlayRender();
        return;
      }
      if (state.resizingRegion) {
        var rName = state.resizingRegion.name;
        var rPx = regionToPixels(state.regions[rName]);
        var minSize = Math.round(20 * s);
        var newW = clamp(pos.x - rPx.x, minSize, overlay.width - rPx.x);
        var newH = clamp(pos.y - rPx.y, minSize, overlay.height - rPx.y);
        state.regions[rName] = Object.assign({}, state.regions[rName], { w: newW / overlay.width, h: newH / overlay.height });
        scheduleOverlayRender();
        return;
      }
      if (state.draggingRegion) {
        var d = state.draggingRegion;
        var dPx = regionToPixels(state.regions[d.name]);
        var newX = clamp(pos.x - d.offsetX, 0, overlay.width - dPx.w);
        var newY = clamp(pos.y - d.offsetY, 0, overlay.height - dPx.h);
        state.regions[d.name] = Object.assign({}, state.regions[d.name], { x: newX / overlay.width, y: newY / overlay.height });
        scheduleOverlayRender();
        return;
      }
      if (state.drawingLasso) {
        appendLassoPoint(pos, s);
        return;
      }
      if (state.drawingRegion) {
        state.drawingRegion.endX = pos.x;
        state.drawingRegion.endY = pos.y;
        scheduleOverlayRender();
        return;
      }
      // Combine modifier held with a target: advertise draw mode ("copy" shows + for shift-add).
      if ((e.shiftKey || e.altKey) &&
          ((state.activeRegion && state.regions[state.activeRegion]) || state.pendingRegion)) {
        state.hoveredRegion = null;
        overlay.style.cursor = e.shiftKey && !e.altKey ? "copy" : "crosshair";
        scheduleOverlayRender();
        return;
      }
      var ctx = overlay.getContext("2d");
      var tBoundsHover = templateOverlayBounds();
      var overTemplate = tBoundsHover
        && pos.x >= tBoundsHover.x && pos.x <= tBoundsHover.x + tBoundsHover.w
        && pos.y >= tBoundsHover.y && pos.y <= tBoundsHover.y + tBoundsHover.h;
      var hit = overTemplate ? null : findHitRegion(pos.x, pos.y, s, ctx);
      state.hoveredRegion = hit;
      if (overTemplate) {
        overlay.style.cursor = "grab";
      } else {
        overlay.style.cursor = hit ? (hit.handle === "resize" ? "nwse-resize" : "grab") : "crosshair";
      }
      scheduleOverlayRender();
    });

    overlay.addEventListener("mouseup", function (e) {
      if (state.pipetteActive) return;
      // Mirror mousedown's button guard, else a right/middle click mid-drag commits the gesture.
      if (e.button !== 0) return;
      if (state.wandDragging) { commitWandDrag(); return; }
      if (state.shapeDraw) { endShapeStroke(); return; }
      if (state.draggingTemplate) {
        _cachedOverlayRect = null;
        state.draggingTemplate = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        flushOverlayRender();
        return;
      }
      if (state.resizingRegion) {
        _cachedOverlayRect = null;
        var rName = state.resizingRegion.name;
        state.resizingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(rName);
        flushOverlayRender();
        updateRegionButtons();
        return;
      }
      if (state.draggingRegion) {
        _cachedOverlayRect = null;
        var dName = state.draggingRegion.name;
        state.draggingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(dName);
        flushOverlayRender();
        updateRegionButtons();
        return;
      }
      if (finishDrawingLasso()) return;
      finishDrawingRegion(e);
    });

    // Blur or a context menu mid-scrub would otherwise leave wandDragging set forever, killing other drags.
    window.addEventListener("blur", cancelWandDrag);
    overlay.addEventListener("contextmenu", cancelWandDrag);
    // A stroke that loses focus mid-drag just ends; the session survives.
    window.addEventListener("blur", endShapeStroke);

    // Document-level listeners so drag/resize continues outside the canvas
    document.addEventListener("mousemove", function (e) {
      if (state.wandDragging) {
        // No button held: the mouseup happened somewhere we never saw it.
        if (e.buttons === 0) cancelWandDrag();
        else updateWandDragFromEvent(e);
        return;
      }
      if (state.shapeDraw && state.shapeDraw.stroking) {
        if (e.buttons === 0) { endShapeStroke(); return; }
        var sdRect = _cachedOverlayRect || overlay.getBoundingClientRect();
        extendShapeStroke(canvasCoords(overlay, e, sdRect));
        return;
      }
      if (state.drawingLasso) {
        var lassoRect = _cachedOverlayRect || overlay.getBoundingClientRect();
        var lassoPos = canvasCoords(overlay, e, lassoRect);
        appendLassoPoint(lassoPos, overlay.width / (lassoRect.width || overlay.width));
        return;
      }
      if (!state.resizingRegion && !state.draggingRegion && !state.draggingTemplate) return;
      var rect = _cachedOverlayRect || overlay.getBoundingClientRect();
      var pos = canvasCoords(overlay, e, rect);
      var displayW = rect.width || overlay.width;
      var s = overlay.width / displayW;
      if (state.draggingTemplate) {
        var tImg = state.uploadedTemplateImg;
        var tw = Math.max(1, Math.round(tImg.naturalWidth * (state.templateScalePreview || 1.0)));
        var thh = Math.max(1, Math.round(tImg.naturalHeight * (state.templateScalePreview || 1.0)));
        var nx = clamp(pos.x - state.draggingTemplate.offsetX, 0, overlay.width - tw);
        var ny = clamp(pos.y - state.draggingTemplate.offsetY, 0, overlay.height - thh);
        state.templateOverlayPos = { x: nx, y: ny };
        scheduleOverlayRender();
        return;
      }
      if (state.resizingRegion) {
        var rName = state.resizingRegion.name;
        var rPx = regionToPixels(state.regions[rName]);
        var minSize = Math.round(20 * s);
        var newW = clamp(pos.x - rPx.x, minSize, overlay.width - rPx.x);
        var newH = clamp(pos.y - rPx.y, minSize, overlay.height - rPx.y);
        state.regions[rName] = Object.assign({}, state.regions[rName], { w: newW / overlay.width, h: newH / overlay.height });
        scheduleOverlayRender();
      } else if (state.draggingRegion) {
        var d = state.draggingRegion;
        var dPx = regionToPixels(state.regions[d.name]);
        var newX = clamp(pos.x - d.offsetX, 0, overlay.width - dPx.w);
        var newY = clamp(pos.y - d.offsetY, 0, overlay.height - dPx.h);
        state.regions[d.name] = Object.assign({}, state.regions[d.name], { x: newX / overlay.width, y: newY / overlay.height });
        scheduleOverlayRender();
      }
    });

    document.addEventListener("mouseup", function (e) {
      if (e.button !== 0) return; // see the overlay mouseup guard above
      if (state.wandDragging) { commitWandDrag(); return; }
      if (state.shapeDraw && state.shapeDraw.stroking) { endShapeStroke(); return; }
      if (finishDrawingLasso()) return;
      if (finishDrawingRegion(e)) return;
      if (state.draggingTemplate) {
        _cachedOverlayRect = null;
        state.draggingTemplate = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        flushOverlayRender();
        return;
      }
      if (state.resizingRegion) {
        _cachedOverlayRect = null;
        var rName = state.resizingRegion.name;
        state.resizingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(rName);
        flushOverlayRender();
        updateRegionButtons();
      } else if (state.draggingRegion) {
        _cachedOverlayRect = null;
        var dName = state.draggingRegion.name;
        state.draggingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(dName);
        flushOverlayRender();
        updateRegionButtons();
      } else {
        _cachedOverlayRect = null;
      }
    });

    qs("#saveRegionBtn").addEventListener("click", function () {
      if (!state.pendingRegion) return;
      showRegionNameModal();
    });

    qs("#clearSelectionBtn").addEventListener("click", function () {
      state.pendingRegion = null;
      state.activeRegion = null;
      renderOverlay();
      updateRegionButtons();
    });

    qs("#mergeRegionsBtn").addEventListener("click", mergeSelectedRegions);

    qs("#deleteRegionBtn").addEventListener("click", function () {
      if (!state.activeRegion) return;
      var name = state.activeRegion;
      // The deleted region is always the active one, so its chip carries `.active`.
      var chip = qs("#regionChips .region-chip.active");
      apiDelete("api/regions/" + encodeURIComponent(name))
        .then(function (data) {
          if (!data.ok) return;
          var commit = function () {
            delete state.regions[name];
            state.activeRegion = null;
            renderRegionChips();
            renderOverlay();
            updateRegionButtons();
            updateRunButton();
            showToast("Region '" + name + "' deleted");
          };
          if (chip && window.ClipgenMotion) ClipgenMotion.animateOut(chip, "delete").then(commit);
          else commit();
        })
        .catch(function () { showToast("Failed to delete region"); });
    });

    qs("#deleteAllRegionsBtn").addEventListener("click", function () {
      if (Object.keys(state.regions).length === 0) return;
      if (!window.confirm("Delete all regions? Stashed regions are not affected.")) return;
      var chips = qsa("#regionChips .region-chip");
      apiDelete("api/regions")
        .then(function (data) {
          if (!data.ok) return;
          var commit = function () {
            state.regions = {};
            state.activeRegion = null;
            state.pendingRegion = null;
            renderRegionChips();
            renderOverlay();
            updateRegionButtons();
            updateRunButton();
            showToast("All regions deleted");
          };
          if (chips.length && window.ClipgenMotion) ClipgenMotion.animateOutAll(chips, "delete").then(commit);
          else commit();
        })
        .catch(function () { showToast("Failed to delete regions"); });
    });

    // Region tool track. Sits outside #workflowParams, so native title tooltips, not initParamTooltips.
    var regionToolTrack = createSegTrack({
      value: state.regionTool,
      options: [
        { value: "rect", icon: "squares-2x2", title: "Rectangle: drag to select", hotkey: "screenspace.regionRect" },
        { value: "lasso", icon: "pencil", title: "Lasso: draw a freehand shape", hotkey: "screenspace.regionLasso" },
        { value: "wand", icon: "sparkles", title: "Magic wand: click a similar-colored area", hotkey: "screenspace.regionWand" },
      ],
      basePath: "/screenspace/icons/",
      onChange: setRegionTool,
    });
    qs("#regionActions").insertBefore(regionToolTrack, qs("#wandToleranceWrap"));
    function setRegionTool(tool) {
      state.regionTool = tool;
      // Abandon the previous tool's in-progress draw, else mouseup could still commit it.
      if (state.wandDragging) cancelWandDrag();
      if (state.shapeDraw) cancelShapeDraw();
      state.drawingLasso = null;
      state.drawingRegion = null;
      // Idempotent after a click (the track already moved); needed when
      // setRegionTool is called programmatically.
      segTrackSetValue(regionToolTrack, tool);
      // Eased width collapse (.wand-tolerance in screenspace.css) so the track glides, not jumps.
      qs("#wandToleranceWrap").classList.toggle("collapsed", tool !== "wand");
    }
    // Registered here because setRegionTool is local. Ungated on purpose: mid-draw presses abort and switch.
    window.ClipgenHotkeys.register([
      { id: "screenspace.regionRect", repeat: false, handler: function () { setRegionTool("rect"); } },
      { id: "screenspace.regionLasso", repeat: false, handler: function () { setRegionTool("lasso"); } },
      { id: "screenspace.regionWand", repeat: false, handler: function () { setRegionTool("wand"); } },
    ]);

    var wandToleranceInput = qs("#wandToleranceInput");
    wandToleranceInput.value = String(state.wandTolerance);
    wandToleranceInput.addEventListener("input", function () {
      state.wandTolerance = parseInt(wandToleranceInput.value, 10) || 32;
      qs("#wandToleranceValue").textContent = String(state.wandTolerance);
    });

    var shapeBrushInput = qs("#shapeBrushInput");
    shapeBrushInput.value = String(state.shapeBrushSize);
    shapeBrushInput.addEventListener("input", function () {
      state.shapeBrushSize = parseInt(shapeBrushInput.value, 10) || 24;
      qs("#shapeBrushValue").textContent = String(state.shapeBrushSize);
    });

    // Toggle region labels
    var toggleLabelsBtn = qs("#toggleLabelsBtn");
    toggleLabelsBtn.appendChild(iconSpan("tag"));
    toggleLabelsBtn.addEventListener("click", function () {
      state.showRegionLabels = !state.showRegionLabels;
      updateRegionButtons();
      renderOverlay();
    });

    // Toggle region visibility
    var toggleRegionsBtn = qs("#toggleRegionsBtn");
    toggleRegionsBtn.appendChild(iconSpan("eye"));
    toggleRegionsBtn.addEventListener("click", function () {
      state.showRegionOverlays = !state.showRegionOverlays;
      updateRegionButtons();
      renderOverlay();
    });

    // Stash all regions
    var stashBtn = qs("#stashRegionsBtn");
    stashBtn.appendChild(iconSpan("archive-box-arrow-down"));
    stashBtn.addEventListener("click", stashRegions);

    // Pin current frame as a positive / negative calibration anchor
    var pinPosBtn = qs("#pinPositiveBtn");
    pinPosBtn.appendChild(iconSpan("check-circle"));
    pinPosBtn.addEventListener("click", function () { pinCurrentFrame("positive"); });
    var pinNegBtn = qs("#pinNegativeBtn");
    pinNegBtn.appendChild(iconSpan("x-circle"));
    pinNegBtn.addEventListener("click", function () { pinCurrentFrame("negative"); });
    qs("#togglePinTrayBtn").addEventListener("click", togglePinTrayVisibility);
    qs("#clearPinsBtn").addEventListener("click", clearAllPins);
    updatePinButtons();

    // Region name modal
    qs("#regionNameCancel").addEventListener("click", hideRegionNameModal);
    qs("#regionNameSave").addEventListener("click", function () {
      var name = qs("#regionNameInput").value.trim();
      if (!name) return;
      var desc = qs("#regionDescInput").value.trim();
      var r = state.pendingRegion;
      var canvas = qs("#overlayCanvas");
      var body = { name: name, x: r.x, y: r.y, w: r.w, h: r.h, canvas_width: canvas.width, canvas_height: canvas.height };
      if (r.points && r.points.length > 0) {
        body.points = r.points; // contours already canvas-pixel absolute for a pending draw
        body.shape = r.shape;
      }
      if (desc) body.description = desc;
      apiPost("api/regions", body)
        .then(function (data) {
          if (data.ok) {
            var saved = data.region;
            if (desc) saved.description = desc;
            state.regions[name] = saved;
            state.pendingRegion = null;
            state.activeRegion = name;
            renderRegionChips();
            renderOverlay();
            updateRegionButtons();
            updateRunButton();
            hideRegionNameModal();
            showToast("Region '" + name + "' saved");
          } else {
            showToast(data.error || "Failed to save region");
          }
        })
        .catch(function () { showToast("Failed to save region"); });
    });

    qs("#regionNameInput").addEventListener("keydown", function (e) {
      if (e.key === "Enter") qs("#regionNameSave").click();
      if (e.key === "Escape") {
        // Stop the page Escape cascade, else cancelling the prompt also discards the pending region.
        e.stopPropagation();
        hideRegionNameModal();
      }
    });

    var chipsEl = qs("#regionChips");
    clipgenWheelToHorizontal(chipsEl);
    chipsEl.addEventListener("scroll", updateRegionChipsOverflow);
  }

  var _regionNameModalPrevFocus = null;

  function showRegionNameModal() {
    var r = state.pendingRegion;
    _regionNameModalPrevFocus = document.activeElement;
    qs("#regionNameInput").value = "";
    qs("#regionDescInput").value = "";
    var coordsText = r ? (r.x + ", " + r.y + " \u2014 " + r.w + "\u00d7" + r.h + " px") : "";
    if (r && r.points && r.points.length > 0) {
      coordsText += " \u00b7 " + contoursTotalPoints(r.points) + " pts (" + (r.shape || "lasso") + ")";
    }
    qs("#regionCoords").textContent = coordsText;
    qs("#regionNameModal").classList.remove("hidden");
    qs("#regionNameInput").focus();
  }

  function hideRegionNameModal() {
    qs("#regionNameModal").classList.add("hidden");
    if (_regionNameModalPrevFocus && typeof _regionNameModalPrevFocus.focus === "function") {
      try { _regionNameModalPrevFocus.focus(); } catch (_) {}
    }
    _regionNameModalPrevFocus = null;
  }

  // Shift-clicked merge set; the active region is the target. renderRegionChips prunes stale names.
  var _mergeSelection = [];

  // Union selected regions into the active one and delete the rest; disjoint shapes stay multi-contour.
  function mergeSelectedRegions() {
    var target = state.activeRegion;
    var others = _mergeSelection.filter(function (n) { return state.regions[n] && n !== target; });
    var overlay = qs("#overlayCanvas");
    if (!target || !state.regions[target] || !others.length || !overlay.width) return;
    var w = overlay.width, h = overlay.height;
    var shapes = [target].concat(others).map(function (n) { return regionShapeAbs(state.regions[n]); });
    var mask = rasterizeShapesMask(shapes, w, h);
    var displayW = overlay.getBoundingClientRect().width || w;
    var contours = simplifyContours(maskToContours(mask, w, h, 8), 2 * (w / displayW), 400);
    if (!contours.length) return;
    saveRegionShape(target, contours, null, function () {
      Promise.all(others.map(function (n) {
        return apiDelete("api/regions/" + encodeURIComponent(n));
      }))
        .then(function () {
          others.forEach(function (n) { delete state.regions[n]; });
          _mergeSelection = [];
          renderRegionChips();
          renderOverlay();
          updateRegionButtons();
          updateRunButton();
          showToast("Merged " + (others.length + 1) + " regions into '" + target + "'");
        })
        .catch(function () { showToast("Merged shapes, but failed to delete a source region"); });
    });
  }

  function updateRegionButtons() {
    var hasPending = !!state.pendingRegion;
    var hasActive = !!state.activeRegion;
    var hasRegions = Object.keys(state.regions).length > 0;
    qs("#saveRegionBtn").classList.toggle("hidden", !hasPending);
    qs("#clearSelectionBtn").classList.toggle("hidden", !hasPending && !hasActive);
    qs("#deleteRegionBtn").classList.toggle("hidden", !hasActive);
    qs("#mergeRegionsBtn").classList.toggle("hidden", !hasActive || _mergeSelection.length === 0);
    qs("#deleteAllRegionsBtn").classList.toggle("hidden", !hasRegions);

    var toggleLabelsBtn = qs("#toggleLabelsBtn");
    var toggleRegionsBtn = qs("#toggleRegionsBtn");
    toggleLabelsBtn.classList.toggle("hidden", !hasRegions);
    toggleRegionsBtn.classList.toggle("hidden", !hasRegions);
    qs("#stashRegionsBtn").classList.toggle("hidden", !hasRegions);
    toggleLabelsBtn.classList.toggle("active", state.showRegionLabels);
    toggleRegionsBtn.classList.toggle("active", state.showRegionOverlays);

    toggleRegionsBtn.innerHTML = "";
    toggleRegionsBtn.appendChild(state.showRegionOverlays ? iconSpan("eye") : iconSpan("eye-slash"));
  }

  // Names at the last render so only new pills animate; null skips the first render.
  var _prevRegionNames = null;

  function renderRegionChips() {
    var container = qs("#regionChips");
    container.innerHTML = "";
    var names = Object.keys(state.regions);
    // Drop merge-set entries that vanished (deleted, stashed, merged away)
    // or became the active region itself.
    _mergeSelection = _mergeSelection.filter(function (n) {
      return state.regions[n] && n !== state.activeRegion;
    });
    if (names.length === 0) {
      _prevRegionNames = {};
      var hint = el("span", "region-hint", "Click and drag on the video to create a region");
      container.appendChild(hint);
      // Still refresh the picker so stale runRegions don't reach preview/calibrate and 400.
      renderRunRegionPicker();
      updateColorSampleBtnLabel();
      return;
    }
    var newPrev = {};
    names.forEach(function (name, i) {
      var color = regionColorForIndex(i);
      var chip = el(
        "div",
        "region-chip"
          + (name === state.activeRegion ? " active" : "")
          + (_mergeSelection.indexOf(name) >= 0 ? " merge-selected" : "")
      );
      chip.style.color = color;
      chip.setAttribute("draggable", "true");
      chip.dataset.regionName = name;
      chip.dataset.regionIdx = String(i);
      var dot = el("span", "region-chip-dot");
      dot.style.background = color;
      chip.appendChild(dot);
      chip.appendChild(document.createTextNode(name));
      chip.addEventListener("click", function (e) {
        if (state.regionSuppressNextClick) {
          state.regionSuppressNextClick = false;
          e.preventDefault();
          e.stopPropagation();
          return;
        }
        // Shift-click with another region active toggles the merge set instead of re-activating.
        if (e.shiftKey && state.activeRegion && state.activeRegion !== name) {
          var idx = _mergeSelection.indexOf(name);
          if (idx >= 0) _mergeSelection.splice(idx, 1);
          else _mergeSelection.push(name);
          renderRegionChips();
          updateRegionButtons();
          return;
        }
        _mergeSelection = [];
        if (state.activeRegion === name) {
          state.activeRegion = null;
        } else {
          state.activeRegion = name;
          state.pendingRegion = null;
        }
        renderRegionChips();
        renderOverlay();
        updateRegionButtons();
        updateRunButton();
      });
      container.appendChild(chip);
      // Animate in only pills new since the last render; reuses the stash-card landing animation.
      if (_prevRegionNames && !_prevRegionNames[name] && window.ClipgenMotion) {
        ClipgenMotion.animateIn(chip, "stashLand");
      }
      newPrev[name] = true;
    });
    _prevRegionNames = newPrev;
    renderRunRegionPicker();
    updateColorSampleBtnLabel();
    updateRegionChipsOverflow();
    _updateMinAreaReadout(""); // region change → refresh color presence pixel estimate
    refreshModelView({ debounce: true });
    refreshCalibration({ debounce: true });
  }

  function updateRegionChipsOverflow() {
    var chips = qs("#regionChips");
    var wrapper = qs("#regionChipsScroll");
    wrapper.classList.toggle("has-overflow", chips.scrollWidth > chips.clientWidth && chips.scrollLeft + chips.clientWidth < chips.scrollWidth - 1);
  }

  // ---- Publish to the hub + sibling satellites ----
  SS.initRegionDrawing = initRegionDrawing;
  SS.renderRegionChips = renderRegionChips;
  SS.updateRegionButtons = updateRegionButtons;
  SS.computeLabelRect = computeLabelRect;
  SS.invalidateOverlayRect = invalidateOverlayRect;
  SS.hideRegionNameModal = hideRegionNameModal;
})();
