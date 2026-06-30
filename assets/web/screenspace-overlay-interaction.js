/* clipgen Screenspace — region draw/drag/resize interaction satellite.
 *
 * Carved out of screenspace.js (the hub) following the hub+satellite convention
 * (see screenspace-overlay/timeline/model-view/...). Owns the region-editor's
 * pointer interaction: the overlay-canvas mousedown/move/up state machine (new-
 * region draw, region move/resize, template drag, pipette sampling), the region
 * chips + toolbar buttons, the region-name modal, saveRegionUpdate, and the
 * overlay-rect cache + render RAF.
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

  // Overlay-rect cache + render RAF, moved from the hub (this satellite owns
  // ~every read/write). The hub keeps an invalidateOverlayRect delegator for its
  // Escape handler; the timeline satellite drops the rect via SS.invalidateOverlayRect.
  var _overlayRaf = 0;
  var _cachedOverlayRect = null;

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

  function saveRegionUpdate(name) {
    var region = state.regions[name];
    if (!region) return;
    var canvas = qs("#overlayCanvas");
    var px = regionToPixels(region);
    var body = { name: name, x: px.x, y: px.y, w: px.w, h: px.h, canvas_width: canvas.width, canvas_height: canvas.height };
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

  // ---- Overlay interaction state machine ----
  //
  // Mousedown picks a mode based on what's under the cursor:
  //   - inside the template overlay        → state.draggingTemplate
  //   - on a region's resize handle        → state.resizingRegion
  //   - inside a region body               → state.draggingRegion
  //   - empty area                         → state.drawingRegion (new region)
  //
  // Mousemove updates whichever mode is active; mouseup commits and clears it.
  // Document-level move/up listeners mirror the canvas handlers so a drag
  // continues smoothly when the cursor leaves the overlay (the canvas-level
  // handlers stop firing, the document ones take over).
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
      var r = normalizeRect(
        state.drawingRegion.startX, state.drawingRegion.startY,
        state.drawingRegion.endX, state.drawingRegion.endY
      );
      state.drawingRegion = null;
      _cachedOverlayRect = null;
      if (r.w > 5 && r.h > 5) {
        state.pendingRegion = r;
      } else if (r.w > 0 || r.h > 0) {
        // Drop the draw silently for click-without-drag (zero size), but
        // surface a hint when the user actually dragged a too-small box —
        // otherwise the picker just snaps closed with no feedback.
        showToast("Region too small — drag a larger area");
      }
      flushOverlayRender();
      updateRegionButtons();
      return true;
    }

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
          // Hidden inputs set programmatically — no DOM event fires, so nudge
          // calibration directly (mirrors setTargetColor on the single-tool path).
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
      // Bail if the overlay has no usable size (e.g. hidden / display:none
      // container). Without this, `s` becomes NaN and propagates into hit
      // testing and any region coords stored downstream.
      if (!displayW || !overlay.width) return;
      var s = overlay.width / displayW;
      var ctx = overlay.getContext("2d");
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
      state.drawingRegion = { startX: pos.x, startY: pos.y, endX: pos.x, endY: pos.y };
      state.pendingRegion = null;
      state.activeRegion = null;
      updateRegionButtons();
    });

    overlay.addEventListener("mousemove", function (e) {
      if (state.pipetteActive) return;
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
      if (state.drawingRegion) {
        state.drawingRegion.endX = pos.x;
        state.drawingRegion.endY = pos.y;
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
      finishDrawingRegion(e);
    });

    // Document-level listeners so drag/resize continues outside the canvas
    document.addEventListener("mousemove", function (e) {
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

    qs("#deleteRegionBtn").addEventListener("click", function () {
      if (!state.activeRegion) return;
      var name = state.activeRegion;
      apiDelete("api/regions/" + encodeURIComponent(name))
        .then(function (data) {
          if (data.ok) {
            delete state.regions[name];
            state.activeRegion = null;
            renderRegionChips();
            renderOverlay();
            updateRegionButtons();
            updateRunButton();
            showToast("Region '" + name + "' deleted");
          }
        })
        .catch(function () { showToast("Failed to delete region"); });
    });

    qs("#deleteAllRegionsBtn").addEventListener("click", function () {
      if (Object.keys(state.regions).length === 0) return;
      if (!window.confirm("Delete all regions? Stashed regions are not affected.")) return;
      apiDelete("api/regions")
        .then(function (data) {
          if (data.ok) {
            state.regions = {};
            state.activeRegion = null;
            state.pendingRegion = null;
            renderRegionChips();
            renderOverlay();
            updateRegionButtons();
            updateRunButton();
            showToast("All regions deleted");
          }
        })
        .catch(function () { showToast("Failed to delete regions"); });
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
      if (e.key === "Escape") hideRegionNameModal();
    });

    // Convert vertical scroll to horizontal on region chips
    var chipsEl = qs("#regionChips");
    chipsEl.addEventListener("wheel", function (e) {
      if (chipsEl.scrollWidth > chipsEl.clientWidth) {
        e.preventDefault();
        chipsEl.scrollLeft += e.deltaY;
      }
    }, { passive: false });
    chipsEl.addEventListener("scroll", updateRegionChipsOverflow);
  }

  var _regionNameModalPrevFocus = null;

  function showRegionNameModal() {
    var r = state.pendingRegion;
    _regionNameModalPrevFocus = document.activeElement;
    qs("#regionNameInput").value = "";
    qs("#regionDescInput").value = "";
    qs("#regionCoords").textContent = r ? (r.x + ", " + r.y + " \u2014 " + r.w + "\u00d7" + r.h + " px") : "";
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

  function updateRegionButtons() {
    var hasPending = !!state.pendingRegion;
    var hasActive = !!state.activeRegion;
    var hasRegions = Object.keys(state.regions).length > 0;
    qs("#saveRegionBtn").classList.toggle("hidden", !hasPending);
    qs("#clearSelectionBtn").classList.toggle("hidden", !hasPending && !hasActive);
    qs("#deleteRegionBtn").classList.toggle("hidden", !hasActive);
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

  function renderRegionChips() {
    var container = qs("#regionChips");
    container.innerHTML = "";
    var names = Object.keys(state.regions);
    if (names.length === 0) {
      var hint = el("span", "region-hint", "Click and drag on the video to create a region");
      container.appendChild(hint);
      // Still refresh the run-region picker so it prunes any runRegions that
      // referenced now-deleted / stashed regions (delete-all, delete-last,
      // stash-all). Otherwise a stale region_ref survives and reaches the
      // preview / calibrate endpoints, which 400 ("Region '<name>' not found").
      renderRunRegionPicker();
      return;
    }
    names.forEach(function (name, i) {
      var color = regionColorForIndex(i);
      var chip = el("div", "region-chip" + (name === state.activeRegion ? " active" : ""));
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
    });
    renderRunRegionPicker();
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
  // Hub calls these through same-named delegators (region-editor init, the
  // Escape handler, the stashing / chip-reorder code). computeLabelRect is
  // destructured at load by screenspace-overlay.js, invalidateOverlayRect by
  // screenspace-timeline.js, renderRegionChips / updateRegionButtons by
  // screenspace-tasks.js.
  SS.initRegionDrawing = initRegionDrawing;
  SS.renderRegionChips = renderRegionChips;
  SS.updateRegionButtons = updateRegionButtons;
  SS.computeLabelRect = computeLabelRect;
  SS.invalidateOverlayRect = invalidateOverlayRect;
  SS.hideRegionNameModal = hideRegionNameModal;
})();
