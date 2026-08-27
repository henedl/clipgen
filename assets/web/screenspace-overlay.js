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

  // Build the canvas path for a shaped region's contours (one closed subpath
  // per contour): points are bbox-relative (0-1 of the region's own rect),
  // r is the pixel bbox. Contours are disjoint, so plain nonzero fill works.
  function traceRegionPolygonPath(ctx, contours, r) {
    ctx.beginPath();
    contours.forEach(function (points) {
      if (points.length < 3) return;
      ctx.moveTo(r.x + points[0][0] * r.w, r.y + points[0][1] * r.h);
      for (var i = 1; i < points.length; i++) {
        ctx.lineTo(r.x + points[i][0] * r.w, r.y + points[i][1] * r.h);
      }
      ctx.closePath();
    });
  }

  // In-progress draw stroke: white for a plain new-region draw, green/red/
  // amber while a shift-add / alt-subtract / shift+alt-intersect edit of the
  // active region is being drawn (Photoshop selection-modifier semantics).
  var COMBINE_STROKES = { add: "#34d399", subtract: "#f87171", intersect: "#fbbf24" };

  function drawStrokeColor(combine) {
    return COMBINE_STROKES[combine] || "#ffffff";
  }

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
        var region = drawRegions[name];
        var r = regionToPixels(region);
        var color = regionColorForIndex(i);
        var isActive = (name === state.activeRegion);
        var isHovered = state.hoveredRegion && state.hoveredRegion.name === name;
        var showHandles = isActive || isHovered;
        ctx.strokeStyle = color;
        ctx.lineWidth = (isActive ? 2 : 1) * s;
        ctx.setLineDash(isActive ? [] : [6 * s, 3 * s]);
        var shaped = region.points && region.points.length > 0;
        if (shaped) {
          // Shaped region: stroke/fill the polygon (points are bbox-relative);
          // label bar and resize handle keep rendering at the bbox below.
          traceRegionPolygonPath(ctx, region.points, r);
          ctx.stroke();
          if (isActive) {
            ctx.fillStyle = hexToRgba(color, 0.12);
            ctx.fill();
          }
        } else {
          ctx.strokeRect(r.x, r.y, r.w, r.h);
          if (isActive) {
            ctx.fillStyle = hexToRgba(color, 0.12);
            ctx.fillRect(r.x, r.y, r.w, r.h);
          }
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

    // Freehand lasso in progress: solid trail + dashed closing segment.
    if (state.drawingLasso && state.drawingLasso.points.length > 1) {
      var lp = state.drawingLasso.points;
      ctx.strokeStyle = drawStrokeColor(state.drawingLasso.combine);
      ctx.lineWidth = 1.5 * s;
      ctx.setLineDash([]);
      ctx.beginPath();
      ctx.moveTo(lp[0][0], lp[0][1]);
      for (var li = 1; li < lp.length; li++) ctx.lineTo(lp[li][0], lp[li][1]);
      ctx.stroke();
      ctx.setLineDash([4 * s, 3 * s]);
      ctx.beginPath();
      ctx.moveTo(lp[lp.length - 1][0], lp[lp.length - 1][1]);
      ctx.lineTo(lp[0][0], lp[0][1]);
      ctx.stroke();
      ctx.setLineDash([]);
    }

    // Drawing in progress
    if (state.drawingRegion) {
      var d = state.drawingRegion;
      ctx.strokeStyle = drawStrokeColor(state.drawingRegion.combine);
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

    // Shape-draw strokes: the opaque mask rendered as a translucent highlight.
    if (state.shapeDraw) {
      ctx.save();
      ctx.globalAlpha = 0.35;
      ctx.drawImage(state.shapeDraw.canvas, 0, 0);
      ctx.restore();
    }

    // Live magic-wand scrub preview: the current flood contour, stroked white
    // for a new region or in the modifier color for a shift/alt combine, with a
    // running tolerance readout. Drawn only while the press-drag is active;
    // release swaps it for a pending region or a boolean edit.
    if (state.wandDragging) {
      var wd = state.wandDragging;
      var wcol = drawStrokeColor(wd.combine);
      if (wd.previewPoints) {
        ctx.strokeStyle = wcol;
        ctx.lineWidth = 1.5 * s;
        ctx.setLineDash([]);
        ctx.beginPath();
        wd.previewPoints.forEach(function (contour) {
          if (contour.length < 3) return;
          ctx.moveTo(contour[0][0], contour[0][1]);
          for (var wi = 1; wi < contour.length; wi++) ctx.lineTo(contour[wi][0], contour[wi][1]);
          ctx.closePath();
        });
        ctx.stroke();
        ctx.fillStyle = hexToRgba(wcol, 0.12);
        ctx.fill();
      }

      // Drag chrome, painted on top of the contour and *not* gated on it: the
      // flood can find nothing contiguous, and without this the whole drag
      // would have zero visual response. The anchor marks where the press
      // landed, the horizontal track shows how far the tolerance scrub has
      // travelled, and the readout rides the head next to the pointer.
      // headOffsetPx is the scrub distance in CSS pixels; mapping it here with
      // the live `s` keeps the head under the cursor even when a panel toggle
      // or window resize changes the canvas-to-display ratio mid-drag.
      var headX = wd.seedX + wd.headOffsetPx * s;
      var dragged = Math.abs(headX - wd.seedX) >= 2 * s;
      if (dragged) {
        ctx.strokeStyle = hexToRgba(wcol, 0.55);
        ctx.lineWidth = 1 * s;
        ctx.setLineDash([3 * s, 3 * s]);
        ctx.beginPath();
        ctx.moveTo(wd.seedX, wd.seedY);
        ctx.lineTo(headX, wd.seedY);
        ctx.stroke();
        ctx.setLineDash([]);
      }
      // Dark ring on each dot so they stay legible over light frame content.
      ctx.lineWidth = 1 * s;
      ctx.strokeStyle = "rgba(0,0,0,0.55)";
      ctx.fillStyle = wcol;
      ctx.beginPath(); ctx.arc(wd.seedX, wd.seedY, 2.5 * s, 0, Math.PI * 2); ctx.fill(); ctx.stroke();
      if (dragged) {
        ctx.beginPath(); ctx.arc(headX, wd.seedY, 2 * s, 0, Math.PI * 2); ctx.fill(); ctx.stroke();
      }
      ctx.font = Math.round(11 * s) + "px " + getThemeColors().fontMono;
      ctx.fillStyle = "rgba(255,255,255,0.9)";
      // Ride the head, but flip to its left when scrubbing leftward (otherwise
      // the readout sits on top of the track and anchor dot) and clamp to the
      // canvas so a seed near an edge doesn't push it out of frame.
      var wlabel = "tol " + wd.tolerance;
      var wpad = Math.round(6 * s);
      var wlw = ctx.measureText(wlabel).width;
      var wlx = headX >= wd.seedX ? headX + wpad : headX - wpad - wlw;
      var wmargin = Math.round(2 * s);
      ctx.fillText(
        wlabel,
        clamp(wlx, wmargin, Math.max(wmargin, canvas.width - wlw - wmargin)),
        clamp(wd.seedY - Math.round(8 * s), Math.round(11 * s), canvas.height - wmargin)
      );
    }

    // Pending (unsaved) region
    if (state.pendingRegion) {
      var p = state.pendingRegion;
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = 1.5 * s;
      ctx.setLineDash([]);
      if (p.points && p.points.length > 0) {
        // Pending shaped region: contours are canvas-pixel absolute (not yet
        // normalized by the server).
        ctx.beginPath();
        p.points.forEach(function (contour) {
          if (contour.length < 3) return;
          ctx.moveTo(contour[0][0], contour[0][1]);
          for (var pi = 1; pi < contour.length; pi++) ctx.lineTo(contour[pi][0], contour[pi][1]);
          ctx.closePath();
        });
        ctx.stroke();
        ctx.fillStyle = "rgba(255, 255, 255, 0.08)";
        ctx.fill();
      } else {
        ctx.strokeRect(p.x, p.y, p.w, p.h);
        ctx.fillStyle = "rgba(255, 255, 255, 0.08)";
        ctx.fillRect(p.x, p.y, p.w, p.h);
      }
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
      if (state.resultOverlay.type === "template" || state.resultOverlay.type === "shape") {
        var boxColor = taskTypeColor(state.resultOverlay.type);
        var matches = state.resultOverlay.data.matches || [];
        matches.forEach(function (m) {
          ctx.strokeStyle = boxColor;
          ctx.lineWidth = 2 * s;
          ctx.strokeRect(m.x, m.y, m.w, m.h);
          ctx.fillStyle = hexToRgba(boxColor, 0.15);
          ctx.fillRect(m.x, m.y, m.w, m.h);
          ctx.font = Math.round(11 * s) + "px " + getThemeColors().fontMono;
          ctx.fillStyle = boxColor;
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
      if (hm.type === "template" || hm.type === "shape" || hm.type === "attention") {
        // Frame-scoped heatmaps cover the whole canvas (attention's
        // region_coords are {0,0,0,0} — the region branch would draw nothing).
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
