/* clipgen Screenspace — sample-preview modal satellite.
 *
 * Owns the click-to-open modal behind every sample thumbnail: view mode is a
 * plain zoomed image; edit mode (Shape/Template) adds an eraser that alphas
 * pixels out of the sample — drag erases, Shift-drag restores. Apply hands the
 * edited RGBA back to the caller as base64, which the hub feeds through the
 * uploaded-image path (shape_image_data / template_image_data).
 *
 * Reads the hub's shared `state` (region polygons for the captured-sample
 * initial alpha) through window.ClipgenScreenspace (SS). qs/el/showToast/
 * openBlockingModal/closeBlockingModal are ambient utils.js globals. Loads
 * after the hub; the hub reaches openSampleModal through a guarded delegator.
 */

(function () {
  "use strict";

  var SS = window.ClipgenScreenspace;
  var state = SS.state;

  var _overlay = null;
  var _brushSize = 24; // display px, persists across opens

  // Shaped regions: rasterize contours as alpha, mirroring the server's attach_capture_mask.
  function applyRegionAlpha(ctx, regionName, w, h) {
    var regs = state.previewRegions || state.regions;
    var r = regionName && regs && regs[regionName];
    if (!r || !r.points || !r.points.length) return;
    ctx.save();
    ctx.globalCompositeOperation = "destination-in";
    // Default (opaque) fill is fine: only alpha survives destination-in.
    ctx.beginPath();
    r.points.forEach(function (contour) {
      if (contour.length < 3) return;
      ctx.moveTo(contour[0][0] * w, contour[0][1] * h);
      for (var i = 1; i < contour.length; i++) {
        ctx.lineTo(contour[i][0] * w, contour[i][1] * h);
      }
      ctx.closePath();
    });
    ctx.fill();
    ctx.restore();
  }

  function eraseSeg(ctx, x0, y0, x1, y1, radius) {
    ctx.save();
    ctx.globalCompositeOperation = "destination-out";
    ctx.lineWidth = radius * 2;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    ctx.beginPath();
    ctx.moveTo(x0, y0);
    ctx.lineTo(x1, y1);
    ctx.stroke();
    ctx.restore();
  }

  // Restore original pixels under circle stamps stepped along the segment.
  function restoreSeg(ctx, original, x0, y0, x1, y1, radius) {
    var dx = x1 - x0, dy = y1 - y0;
    var steps = Math.max(1, Math.ceil(Math.sqrt(dx * dx + dy * dy) / (radius / 2)));
    for (var i = 0; i <= steps; i++) {
      var x = x0 + (dx * i) / steps;
      var y = y0 + (dy * i) / steps;
      ctx.save();
      ctx.beginPath();
      ctx.arc(x, y, radius, 0, Math.PI * 2);
      ctx.clip();
      ctx.clearRect(x - radius, y - radius, radius * 2, radius * 2);
      ctx.drawImage(original, 0, 0);
      ctx.restore();
    }
  }

  // spec: { mode: "edit"|"view", title, dataUrl, regionName, onApply(b64) }
  function openSampleModal(spec) {
    if (!spec || !spec.dataUrl) return;
    if (!_overlay) {
      _overlay = el("div", "ss-sample-modal modal-overlay cg-modal-overlay hidden");
      _overlay.appendChild(el("div", "ss-sample-modal__card cg-modal-card"));
      document.body.appendChild(_overlay);
    }
    var overlay = _overlay;
    var card = overlay.querySelector(".ss-sample-modal__card");
    card.innerHTML = "";
    card.appendChild(el("h3", "", spec.title || "Sample"));
    var stage = el("div", "ss-sample-modal__stage");
    card.appendChild(stage);

    var cleanups = [];
    function close() {
      cleanups.forEach(function (fn) { fn(); });
      overlay.classList.add("hidden");
      closeBlockingModal(overlay);
    }

    if (spec.mode !== "edit") {
      var img = document.createElement("img");
      img.className = "ss-sample-modal__img";
      img.decoding = "async";
      img.src = spec.dataUrl;
      img.alt = spec.title || "Sample";
      stage.appendChild(img);
      var viewActions = el("div", "modal-actions");
      var closeBtn = el("button", "btn btn-small", "Close");
      closeBtn.type = "button";
      closeBtn.addEventListener("click", close);
      viewActions.appendChild(closeBtn);
      card.appendChild(viewActions);
    } else {
      var work = document.createElement("canvas");
      work.className = "ss-sample-canvas";
      var original = document.createElement("canvas");
      var wctx = work.getContext("2d");
      var srcImg = new Image();
      srcImg.onload = function () {
        work.width = srcImg.naturalWidth;
        work.height = srcImg.naturalHeight;
        original.width = srcImg.naturalWidth;
        original.height = srcImg.naturalHeight;
        // Original stays unmasked so Shift-restore can recover any pixel.
        original.getContext("2d").drawImage(srcImg, 0, 0);
        wctx.drawImage(srcImg, 0, 0);
        applyRegionAlpha(wctx, spec.regionName, work.width, work.height);
      };
      srcImg.src = spec.dataUrl;
      stage.appendChild(work);

      // Brush ghost ring on the whole stage: the checkerboard margin is paintable too.
      stage.classList.add("ss-sample-modal__stage--edit");
      var ghost = el("div", "ss-sample-modal__ghost hidden");
      stage.appendChild(ghost);
      function sizeGhost() {
        ghost.style.width = _brushSize + "px";
        ghost.style.height = _brushSize + "px";
      }
      sizeGhost();
      stage.addEventListener("mouseenter", function () {
        ghost.classList.remove("hidden");
      });
      stage.addEventListener("mouseleave", function () {
        ghost.classList.add("hidden");
      });
      stage.addEventListener("mousemove", function (e) {
        var rect = stage.getBoundingClientRect();
        ghost.style.left = e.clientX - rect.left + "px";
        ghost.style.top = e.clientY - rect.top + "px";
      });

      // Stroke-level undo/redo: full-canvas snapshots, capped.
      var history = [];
      var redoStack = [];
      var undoBtn, redoBtn;
      function snapshot() {
        return wctx.getImageData(0, 0, work.width, work.height);
      }
      function syncHistButtons() {
        undoBtn.disabled = !history.length;
        redoBtn.disabled = !redoStack.length;
      }
      function pushHistory() {
        history.push(snapshot());
        if (history.length > 20) history.shift();
        redoStack.length = 0;
        syncHistButtons();
      }
      function undo() {
        if (!history.length) return;
        redoStack.push(snapshot());
        wctx.putImageData(history.pop(), 0, 0);
        syncHistButtons();
      }
      function redo() {
        if (!redoStack.length) return;
        history.push(snapshot());
        wctx.putImageData(redoStack.pop(), 0, 0);
        syncHistButtons();
      }

      var hint = el("div", "ss-sample-modal__hint",
        "Drag to erase — Shift-drag restores");
      card.appendChild(hint);

      var controls = el("div", "ss-sample-modal__controls");
      controls.appendChild(el("span", "param-label", "Brush"));
      var brush = document.createElement("input");
      brush.type = "range";
      brush.min = "4";
      brush.max = "96";
      brush.step = "1";
      brush.value = String(_brushSize);
      var brushVal = el("span", "ss-sample-modal__brush-val", String(_brushSize));
      brush.addEventListener("input", function () {
        _brushSize = parseInt(brush.value, 10) || 24;
        brushVal.textContent = String(_brushSize);
        sizeGhost();
      });
      controls.appendChild(brush);
      controls.appendChild(brushVal);
      undoBtn = el("button", "btn btn-small ss-sample-modal__hist");
      undoBtn.type = "button";
      undoBtn.title = "Undo";
      undoBtn.setAttribute("aria-label", "Undo");
      undoBtn.appendChild(iconMaskSpan("arrow-uturn-left", {
        className: "ss-sample-modal__hist-icon",
        basePath: "/screenspace/icons/",
      }));
      undoBtn.addEventListener("click", undo);
      redoBtn = el("button", "btn btn-small ss-sample-modal__hist");
      redoBtn.type = "button";
      redoBtn.title = "Redo";
      redoBtn.setAttribute("aria-label", "Redo");
      redoBtn.appendChild(iconMaskSpan("arrow-uturn-right", {
        className: "ss-sample-modal__hist-icon",
        basePath: "/screenspace/icons/",
      }));
      redoBtn.addEventListener("click", redo);
      syncHistButtons();
      controls.appendChild(undoBtn);
      controls.appendChild(redoBtn);
      card.appendChild(controls);

      var stroking = null; // {restoring, lastX, lastY}
      function canvasPos(e) {
        var rect = work.getBoundingClientRect();
        return {
          x: (e.clientX - rect.left) * (work.width / rect.width),
          y: (e.clientY - rect.top) * (work.height / rect.height),
          scale: work.width / rect.width,
        };
      }
      function stamp(pos) {
        var radius = Math.max(1, (_brushSize / 2) * pos.scale);
        if (stroking.restoring) {
          restoreSeg(wctx, original, stroking.lastX, stroking.lastY, pos.x, pos.y, radius);
        } else {
          eraseSeg(wctx, stroking.lastX, stroking.lastY, pos.x, pos.y, radius);
        }
        stroking.lastX = pos.x;
        stroking.lastY = pos.y;
      }
      stage.addEventListener("mousedown", function (e) {
        if (e.button !== 0 || !work.width) return;
        e.preventDefault();
        pushHistory();
        var pos = canvasPos(e);
        stroking = { restoring: e.shiftKey, lastX: pos.x, lastY: pos.y };
        stamp(pos);
      });
      // Document-level so strokes survive leaving the canvas; out-of-bounds stamps
      // clip harmlessly.
      function strokeMove(e) {
        if (!stroking) return;
        if (e.buttons === 0) { stroking = null; return; }
        stamp(canvasPos(e));
      }
      function endStroke() { stroking = null; }
      document.addEventListener("mousemove", strokeMove);
      document.addEventListener("mouseup", endStroke);
      cleanups.push(function () {
        document.removeEventListener("mousemove", strokeMove);
        document.removeEventListener("mouseup", endStroke);
      });

      var actions = el("div", "modal-actions");
      var cancelBtn = el("button", "btn btn-small", "Cancel");
      cancelBtn.type = "button";
      cancelBtn.addEventListener("click", close);
      var applyBtn = el("button", "btn btn-small btn-primary", "Apply");
      applyBtn.type = "button";
      applyBtn.addEventListener("click", function () {
        var b64 = work.toDataURL("image/png").split(",")[1];
        close();
        if (spec.onApply) spec.onApply(b64);
      });
      actions.appendChild(cancelBtn);
      actions.appendChild(applyBtn);
      card.appendChild(actions);
    }

    overlay.classList.remove("hidden");
    openBlockingModal(overlay, {
      trapFocus: true,
      restoreFocus: true,
      onEscape: close,
      onBackdropClick: close,
    });
  }

  // ---- Publish to the hub ----
  SS.openSampleModal = openSampleModal;
})();
