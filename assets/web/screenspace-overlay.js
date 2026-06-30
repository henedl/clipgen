/* clipgen Screenspace — overlay-render satellite.
 *
 * Carved out of screenspace.js (the hub) following the hub+satellite convention
 * (see screenspace-tasks/results/...). renderOverlay paints the region editor's
 * canvas chrome: saved/preview regions + labels + handles, the in-progress draw
 * rect, the pending region, the draggable template preview, template/flow result
 * overlays, the heatmap composite, and the model-view comparator.
 *
 * It is a pure read of the hub's shared `state` plus a handful of hub helpers,
 * all reached through the window.ClipgenScreenspace (SS) namespace. hexToRgba
 * and qs are ambient utils.js globals (scope chain). The hub keeps a same-named
 * renderOverlay delegator for its ~33 call sites. Loaded by screenspace.html
 * right after screenspace.js and BEFORE screenspace-tasks.js / -results.js,
 * which destructure SS.renderOverlay at load time.
 */

(function () {
  "use strict";

  var SS = window.ClipgenScreenspace;
  var state = SS.state;
  // Hub helpers (published synchronously during the hub's load, before this
  // file runs). hexToRgba / qs are ambient utils.js globals.
  var regionToPixels = SS.regionToPixels,
    regionColorForIndex = SS.regionColorForIndex,
    computeLabelRect = SS.computeLabelRect,
    getThemeColors = SS.getThemeColors,
    templateOverlayBounds = SS.templateOverlayBounds,
    taskTypeColor = SS.taskTypeColor,
    _overlayEligibleForActiveTool = SS._overlayEligibleForActiveTool;

  function renderOverlay() {
    var canvas = qs("#overlayCanvas");
    if (!canvas.width || !canvas.height) return;
    var ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    // Scale factor: canvas pixels per display pixel, so chrome looks
    // the same physical size regardless of the underlying video resolution.
    var displayW = canvas.getBoundingClientRect().width || canvas.width;
    var s = canvas.width / displayW;

    // Draw saved regions (or preview regions when hovering a stash)
    var drawRegions = state.previewRegions || state.regions;
    if (state.showRegionOverlays) {
      var names = Object.keys(drawRegions);
      names.forEach(function (name, i) {
        var r = regionToPixels(drawRegions[name]);
        var color = regionColorForIndex(i);
        var isActive = (name === state.activeRegion);
        var isHovered = state.hoveredRegion && state.hoveredRegion.name === name;
        var showHandles = isActive || isHovered;
        ctx.strokeStyle = color;
        ctx.lineWidth = (isActive ? 2 : 1) * s;
        ctx.setLineDash(isActive ? [] : [6 * s, 3 * s]);
        ctx.strokeRect(r.x, r.y, r.w, r.h);
        if (isActive) {
          ctx.fillStyle = hexToRgba(color, 0.12);
          ctx.fillRect(r.x, r.y, r.w, r.h);
        }
        ctx.setLineDash([]);

        // Label with grip indicator
        if (state.showRegionLabels) {
          var lr = computeLabelRect(r, name, ctx, s);
          ctx.fillStyle = hexToRgba(color, 0.85);
          ctx.fillRect(lr.x, lr.y, lr.w, lr.h);
          if (showHandles) {
            ctx.fillStyle = "rgba(255,255,255,0.5)";
            var dotR = Math.round(1 * s);
            var gripColGap = Math.round(3 * s);
            var gripRowGap = Math.round(3 * s);
            var gx = lr.x + lr.gripPadL + dotR + Math.round(1 * s);
            var gy = lr.y + Math.round(lr.h / 2) - gripRowGap;
            for (var row = 0; row < 3; row++) {
              for (var col = 0; col < 2; col++) {
                ctx.beginPath();
                ctx.arc(gx + col * gripColGap, gy + row * gripRowGap, dotR, 0, Math.PI * 2);
                ctx.fill();
              }
            }
          }
          ctx.fillStyle = "#fff";
          ctx.fillText(name, lr.x + lr.gripPadL + lr.gripW + lr.pad, r.y - Math.round(4 * s));
        }

        // Resize handle (bottom-right corner, 3-dot triangle)
        if (showHandles) {
          var dotRr = Math.round(1.5 * s);
          var handlePad = Math.round(5 * s);
          var dotSpacing = Math.round(4 * s);
          var bx = r.x + r.w - handlePad;
          var by = r.y + r.h - handlePad;
          ctx.fillStyle = hexToRgba(color, 0.9);
          ctx.beginPath(); ctx.arc(bx, by, dotRr, 0, Math.PI * 2); ctx.fill();
          ctx.beginPath(); ctx.arc(bx - dotSpacing, by, dotRr, 0, Math.PI * 2); ctx.fill();
          ctx.beginPath(); ctx.arc(bx, by - dotSpacing, dotRr, 0, Math.PI * 2); ctx.fill();
        }
      });
    }

    // Drawing in progress
    if (state.drawingRegion) {
      var d = state.drawingRegion;
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = 1.5 * s;
      ctx.setLineDash([4 * s, 3 * s]);
      ctx.strokeRect(d.startX, d.startY, d.endX - d.startX, d.endY - d.startY);
      ctx.setLineDash([]);
      // Dimensions
      var w = Math.abs(d.endX - d.startX);
      var h = Math.abs(d.endY - d.startY);
      if (w > 20 && h > 20) {
        ctx.font = Math.round(11 * s) + "px " + getThemeColors().fontMono;
        ctx.fillStyle = "rgba(255,255,255,0.9)";
        ctx.fillText(w + "\u00d7" + h, Math.min(d.startX, d.endX) + Math.round(4 * s), Math.max(d.startY, d.endY) + Math.round(14 * s));
      }
    }

    // Pending (unsaved) region
    if (state.pendingRegion) {
      var p = state.pendingRegion;
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = 1.5 * s;
      ctx.setLineDash([]);
      ctx.strokeRect(p.x, p.y, p.w, p.h);
      ctx.fillStyle = "rgba(255, 255, 255, 0.08)";
      ctx.fillRect(p.x, p.y, p.w, p.h);
      ctx.font = Math.round(11 * s) + "px " + getThemeColors().fontMono;
      ctx.fillStyle = "rgba(255,255,255,0.9)";
      ctx.fillText(p.w + "\u00d7" + p.h + " px", p.x + Math.round(4 * s), p.y + p.h + Math.round(14 * s));
    }

    // Template preview: overlay the uploaded PNG at its effective in-video
    // size (native PNG pixels * template_scale) so the user can see how
    // large it will match. Canvas is sized in native video pixels, so the
    // display-to-video ratio is applied automatically when the browser
    // scales the canvas to fit the viewport. The overlay is draggable so
    // the user can position it against specific elements in the frame.
    var tBounds = templateOverlayBounds();
    if (tBounds) {
      var tImg = state.uploadedTemplateImg;
      ctx.globalAlpha = state.draggingTemplate ? 0.9 : 0.75;
      ctx.drawImage(tImg, tBounds.x, tBounds.y, tBounds.w, tBounds.h);
      ctx.globalAlpha = 1.0;
      ctx.strokeStyle = taskTypeColor("template");
      ctx.lineWidth = (state.draggingTemplate ? 2 : 1.5) * s;
      ctx.setLineDash([4 * s, 3 * s]);
      ctx.strokeRect(tBounds.x, tBounds.y, tBounds.w, tBounds.h);
      ctx.setLineDash([]);
      ctx.font = Math.round(11 * s) + "px " + getThemeColors().fontMono;
      ctx.fillStyle = taskTypeColor("template");
      ctx.fillText(tBounds.w + "\u00d7" + tBounds.h + " px",
        tBounds.x + Math.round(4 * s),
        tBounds.y + tBounds.h + Math.round(14 * s));
    }

    // Result overlay: template match bounding boxes / flow motion arrows
    if (state.resultOverlay) {
      ctx.setLineDash([]);
      if (state.resultOverlay.type === "template") {
        var matches = state.resultOverlay.data.matches || [];
        matches.forEach(function (m) {
          ctx.strokeStyle = taskTypeColor("template");
          ctx.lineWidth = 2 * s;
          ctx.strokeRect(m.x, m.y, m.w, m.h);
          ctx.fillStyle = hexToRgba(taskTypeColor("template"), 0.15);
          ctx.fillRect(m.x, m.y, m.w, m.h);
          ctx.font = Math.round(11 * s) + "px " + getThemeColors().fontMono;
          ctx.fillStyle = taskTypeColor("template");
          ctx.fillText((m.score * 100).toFixed(0) + "%", m.x + Math.round(3 * s), m.y - Math.round(4 * s));
        });
      } else if (state.resultOverlay.type === "flow") {
        var grid = state.resultOverlay.data.flow_grid || [];
        var fRegion = state.resultOverlay.region;
        if (fRegion && fRegion.w && grid.length) {
          var maxMag = 0;
          grid.forEach(function (c) { if (c.mag > maxMag) maxMag = c.mag; });
          if (maxMag > 0) {
            grid.forEach(function (c) {
              var px = fRegion.x + c.x * fRegion.w;
              var py = fRegion.y + c.y * fRegion.h;
              var norm = Math.min(c.mag / maxMag, 1);
              var arrowLen = norm * 20 * s;
              var rad = c.ang * Math.PI / 180;
              var ex = px + Math.cos(rad) * arrowLen;
              var ey = py + Math.sin(rad) * arrowLen;
              var alpha = norm * 0.8 + 0.2;
              ctx.strokeStyle = "rgba(99, 102, 241, " + alpha + ")";
              ctx.lineWidth = 1.5 * s;
              ctx.beginPath();
              ctx.moveTo(px, py);
              ctx.lineTo(ex, ey);
              ctx.stroke();
              var headLen = 4 * s;
              ctx.beginPath();
              ctx.moveTo(ex, ey);
              ctx.lineTo(ex - headLen * Math.cos(rad - 0.4), ey - headLen * Math.sin(rad - 0.4));
              ctx.moveTo(ex, ey);
              ctx.lineTo(ex - headLen * Math.cos(rad + 0.4), ey - headLen * Math.sin(rad + 0.4));
              ctx.stroke();
            });
          }
        }
      }
    }

    // Heatmap image overlay (semi-transparent composite)
    if (state.heatmapOverlay && state.heatmapOverlay._img) {
      var hm = state.heatmapOverlay;
      ctx.globalAlpha = 0.5;
      if (hm.type === "template") {
        ctx.drawImage(hm._img, 0, 0, canvas.width, canvas.height);
      } else if (hm.type === "flow" || hm.type === "change") {
        var rPx = hm.region_coords;
        if (rPx && rPx.w) {
          ctx.drawImage(hm._img, rPx.x, rPx.y, rPx.w, rPx.h);
        }
      }
      ctx.globalAlpha = 1.0;
    }

    // Model-view overlay (toggle or held-key blink comparator)
    var overlayActive = (state.overlayEnabled || state.overlayBlinkActive)
      && state.overlayImage
      && _overlayEligibleForActiveTool();
    if (overlayActive) {
      ctx.globalAlpha = state.overlayBlinkActive ? 1.0 : 0.7;
      var scope = state.overlayImageScope || "region";
      if (scope === "frame") {
        ctx.drawImage(state.overlayImage, 0, 0, canvas.width, canvas.height);
      } else {
        var oRegion = state.pendingRegion
          ? state.pendingRegion
          : (state.activeRegion ? state.regions[state.activeRegion] : null);
        if (oRegion) {
          var oPx = regionToPixels(oRegion);
          if (oPx && oPx.w && oPx.h) {
            ctx.drawImage(state.overlayImage, oPx.x, oPx.y, oPx.w, oPx.h);
          }
        } else {
          // No region active — overlay covers the whole frame.
          ctx.drawImage(state.overlayImage, 0, 0, canvas.width, canvas.height);
        }
      }
      ctx.globalAlpha = 1.0;
    }
  }

  SS.renderOverlay = renderOverlay;
})();
