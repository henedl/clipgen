/* clipgen Screenspace — timeline satellite.
 *
 * Carved out of screenspace.js (the hub) following the hub+satellite convention
 * (see screenspace-overlay/tasks/results). Owns the timeline canvas: the ruler,
 * zoom/pan, scrubbing, in/out markers, boundary-flag rail, playhead overlay,
 * result tooltips, and the type legend.
 *
 * Reaches hub helpers through window.ClipgenScreenspace (SS): loadFrame /
 * seekPlayhead (frame viewer, stay in the hub), taskTypeColor, getThemeColors.
 * findTask / focusedTaskId live in screenspace-tasks.js, which loads AFTER this
 * file, so they are called late-bound as SS.findTask(...) / SS.focusedTaskId(...)
 * rather than destructured. formatTime/formatDuration/clamp/hexToRgba/el/qs and
 * the drawTimelineRuler/niceTimeInterval/drawAmplitudeBands canvas helpers are
 * ambient utils.js globals (scope chain).
 *
 * The hub keeps same-named delegators (initTimeline/renderTimeline/renderPlayhead)
 * for the entry points its own code calls; the tasks and results satellites
 * destructure SS.renderTimeline / SS.updateMarkerInfo at load, so this file is
 * loaded by screenspace.html after screenspace.js and BEFORE those two.
 */

(function () {
  "use strict";

  var SS = window.ClipgenScreenspace;
  var state = SS.state;
  // Hub helpers (published during the hub's load, before this file runs).
  var loadFrame = SS.loadFrame,
    seekPlayhead = SS.seekPlayhead,
    taskTypeColor = SS.taskTypeColor,
    getThemeColors = SS.getThemeColors,
    buildTypeIcon = SS.buildTypeIcon,
    iconSpan = SS.iconSpan,
    invalidateOverlayRect = SS.invalidateOverlayRect;

  var TIMELINE_CANVAS_HEIGHT = 64;
  var _timelineHitRects = [];
  var _cachedTimelineRect = null;

  function getTimelineRect(canvas) {
    if (!_cachedTimelineRect) _cachedTimelineRect = canvas.getBoundingClientRect();
    return _cachedTimelineRect;
  }

  function initTimeline() {
    qs("#zoomInBtn").appendChild(iconSpan("plus", "ss-icon--sm"));
    qs("#zoomOutBtn").appendChild(iconSpan("minus", "ss-icon--sm"));

    var storedTimeline = getStoredUIState("screenspace");
    if (storedTimeline.amplitudeGraphEnabled === true) {
      state.amplitudeGraphEnabled = true;
      qs("#amplitudeGraphBtn").classList.add("active");
    }
    var canvas = qs("#timelineCanvas");
    sizeTimelineCanvas();
    window.addEventListener("resize", function () {
      invalidateOverlayRect();
      _cachedTimelineRect = null;
      sizeTimelineCanvas();
    });
    window.addEventListener("scroll", function () {
      invalidateOverlayRect();
      _cachedTimelineRect = null;
    }, true);

    canvas.addEventListener("click", function (e) {
      if (state.timelineDragging) return;
      var ts = timelineXToTime(e);
      if (ts !== null) {
        state.resultOverlay = null;
        loadFrame(ts);
      }
    });

    canvas.addEventListener("wheel", function (e) {
      e.preventDefault();
      if (!state.videoInfo) return;
      var dur = state.videoInfo.duration;
      var zoomFactor = e.deltaY < 0 ? 1.3 : 1 / 1.3;
      var mouseTs = timelineXToTime(e);
      var oldZoom = state.timelineZoom;
      state.timelineZoom = clamp(oldZoom * zoomFactor, 1, 200);
      if (mouseTs !== null && state.timelineZoom > 1) {
        var rect = canvas.getBoundingClientRect();
        var frac = (e.clientX - rect.left) / rect.width;
        var visLen = dur / state.timelineZoom;
        state.timelineOffset = clamp(mouseTs - frac * visLen, 0, dur - visLen);
      }
      if (state.timelineZoom <= 1) state.timelineOffset = 0;
      renderTimeline();
    }, { passive: false });

    var dragStart = null;
    var scrubbing = false;
    var scrubRaf = 0;
    canvas.addEventListener("mousedown", function (e) {
      hideSsTooltip();
      _lastTimelineHit = null;
      if (state.timelineZoom > 1) {
        // Zoomed in: drag to pan
        dragStart = { x: e.clientX, offset: state.timelineOffset };
        state.timelineDragging = false;
      } else {
        // Not zoomed: drag to scrub
        scrubbing = true;
        var ts = timelineXToTime(e);
        if (ts !== null) loadFrame(ts);
      }
    });
    document.addEventListener("mousemove", function (e) {
      if (scrubbing) {
        var ts = timelineXToTime(e);
        if (ts !== null) {
          seekPlayhead(ts);
          if (!scrubRaf) {
            scrubRaf = requestAnimationFrame(function () {
              scrubRaf = 0;
              loadFrame(state.currentTimestamp);
            });
          }
        }
        return;
      }
      if (!dragStart) return;
      var dx = e.clientX - dragStart.x;
      if (Math.abs(dx) > 3) state.timelineDragging = true;
      if (!state.videoInfo) return;
      var dur = state.videoInfo.duration;
      var visLen = dur / state.timelineZoom;
      var rect = canvas.getBoundingClientRect();
      var dtSec = -(dx / rect.width) * visLen;
      state.timelineOffset = clamp(dragStart.offset + dtSec, 0, dur - visLen);
      renderTimeline();
    });
    document.addEventListener("mouseup", function () {
      if (scrubbing) {
        scrubbing = false;
        if (scrubRaf) { cancelAnimationFrame(scrubRaf); scrubRaf = 0; }
        loadFrame(state.currentTimestamp);
      }
      if (dragStart) {
        setTimeout(function () { state.timelineDragging = false; }, 50);
        dragStart = null;
      }
    });

    var _ssTooltipRaf = 0;
    var _lastTimelineHit = null;
    canvas.addEventListener("mousemove", function (e) {
      if (scrubbing || dragStart) {
        hideSsTooltip();
        _lastTimelineHit = null;
        return;
      }
      if (_ssTooltipRaf) return;
      var cx = e.clientX;
      var cy = e.clientY;
      _ssTooltipRaf = requestAnimationFrame(function () {
        _ssTooltipRaf = 0;
        var hit = hitTestTimeline(cx, cy);
        if (hit) {
          _lastTimelineHit = hit;
          showSsTooltip(hit, cx, cy);
        } else if (_lastTimelineHit) {
          _lastTimelineHit = null;
          hideSsTooltip();
        }
      });
    });
    canvas.addEventListener("mouseleave", function () {
      _lastTimelineHit = null;
      hideSsTooltip();
    });

    qs("#zoomInBtn").addEventListener("click", function () {
      if (!state.videoInfo) return;
      state.timelineZoom = clamp(state.timelineZoom * 1.5, 1, 200);
      clampTimelineOffset();
      renderTimeline();
    });
    qs("#zoomOutBtn").addEventListener("click", function () {
      state.timelineZoom = clamp(state.timelineZoom / 1.5, 1, 200);
      if (state.timelineZoom <= 1) state.timelineOffset = 0;
      clampTimelineOffset();
      renderTimeline();
    });
    qs("#zoomResetBtn").addEventListener("click", function () {
      state.timelineZoom = 1;
      state.timelineOffset = 0;
      renderTimeline();
    });
    qs("#setInBtn").addEventListener("click", function () {
      state.inMarker = state.currentTimestamp;
      if (state.outMarker !== null && state.inMarker > state.outMarker) state.outMarker = null;
      updateMarkerInfo();
      renderTimeline();
    });
    qs("#setOutBtn").addEventListener("click", function () {
      state.outMarker = state.currentTimestamp;
      if (state.inMarker !== null && state.outMarker < state.inMarker) state.inMarker = null;
      updateMarkerInfo();
      renderTimeline();
    });
    qs("#clearMarkersBtn").addEventListener("click", function () {
      state.inMarker = null;
      state.outMarker = null;
      updateMarkerInfo();
      renderTimeline();
    });
    qs("#amplitudeGraphBtn").addEventListener("click", function () {
      state.amplitudeGraphEnabled = !state.amplitudeGraphEnabled;
      this.classList.toggle("active", state.amplitudeGraphEnabled);
      setStoredUIStateField("screenspace", "amplitudeGraphEnabled", state.amplitudeGraphEnabled);
      renderTimeline();
    });
  }

  function clampTimelineOffset() {
    if (!state.videoInfo) return;
    var dur = state.videoInfo.duration;
    var visLen = dur / state.timelineZoom;
    state.timelineOffset = clamp(state.timelineOffset, 0, Math.max(0, dur - visLen));
  }

  function sizeTimelineCanvas() {
    var canvas = qs("#timelineCanvas");
    _cachedTimelineRect = null;
    var rect = canvas.getBoundingClientRect();
    canvas.width = Math.floor(rect.width);
    canvas.height = TIMELINE_CANVAS_HEIGHT;
    var ph = qs("#playheadCanvas");
    ph.width = canvas.width;
    ph.height = canvas.height;
    renderTimeline();
    renderPlayhead();
  }

  function timelineXToTime(event) {
    if (!state.videoInfo) return null;
    var canvas = qs("#timelineCanvas");
    var rect = getTimelineRect(canvas);
    var frac = (event.clientX - rect.left) / rect.width;
    frac = clamp(frac, 0, 1);
    var dur = state.videoInfo.duration;
    var visLen = dur / state.timelineZoom;
    return state.timelineOffset + frac * visLen;
  }

  function updateMarkerInfo() {
    var info = qs("#markerInfo");
    var clearBtn = qs("#clearMarkersBtn");
    if (state.inMarker === null && state.outMarker === null) {
      info.textContent = "";
      clearBtn.classList.add("hidden");
      return;
    }
    clearBtn.classList.remove("hidden");
    var parts = [];
    if (state.inMarker !== null) parts.push("In: " + formatTime(state.inMarker, { decimals: 1 }));
    if (state.outMarker !== null) parts.push("Out: " + formatTime(state.outMarker, { decimals: 1 }));
    info.textContent = parts.join("  ");
  }

  function renderTimeline() {
    var canvas = qs("#timelineCanvas");
    var ctx = canvas.getContext("2d");
    var w = canvas.width;
    var h = canvas.height;
    var dur = state.videoInfo ? state.videoInfo.duration : 0;

    var tc = getThemeColors();

    ctx.clearRect(0, 0, w, h);

    // Background
    ctx.fillStyle = tc.surfaceAlt;
    ctx.fillRect(0, 0, w, h);

    if (dur <= 0) {
      ctx.fillStyle = tc.textDim;
      ctx.font = "12px -apple-system, sans-serif";
      ctx.textAlign = "center";
      ctx.fillText("No video loaded", w / 2, h / 2 + 4);
      ctx.textAlign = "start";
      return;
    }

    var visLen = dur / state.timelineZoom;
    var visStart = state.timelineOffset;
    var visEnd = visStart + visLen;

    function timeToX(t) {
      return ((t - visStart) / visLen) * w;
    }

    // Time ruler ticks
    drawTimelineRuler(ctx, {
      visStart: visStart,
      visEnd: visEnd,
      interval: niceTimeInterval(visLen, { maxTicks: 20 }),
      timeToX: timeToX,
      colors: { border: tc.border, textDim: tc.textDim, fontMono: tc.fontMono },
      tickHeight: 8,
      labelY: 18,
      format: formatDuration,
    });

    // In/Out marker shading — scrim "outside" the active range against the
    // timeline's surfaceAlt background. fg works in both themes (fg is white in
    // dark → lightens, dark in light → darkens; both differentiate the range).
    if (state.inMarker !== null || state.outMarker !== null) {
      ctx.fillStyle = hexToRgba(tc.fg, 0.12);
      if (state.inMarker !== null) {
        var inX = timeToX(state.inMarker);
        ctx.fillRect(0, 0, Math.max(0, inX), h);
      }
      if (state.outMarker !== null) {
        var outX = timeToX(state.outMarker);
        ctx.fillRect(outX, 0, w - outX, h);
      }
      // Marker lines
      ctx.strokeStyle = "#16a34a";
      ctx.lineWidth = 2;
      if (state.inMarker !== null) {
        var ix = timeToX(state.inMarker);
        ctx.beginPath(); ctx.moveTo(ix, 0); ctx.lineTo(ix, h); ctx.stroke();
      }
      if (state.outMarker !== null) {
        var ox = timeToX(state.outMarker);
        ctx.beginPath(); ctx.moveTo(ox, 0); ctx.lineTo(ox, h); ctx.stroke();
      }
    }

    // Optional amplitude band: per-task-type event-density curves above the
    // result markers. Drawn before markers so markers stay on top visually.
    var ampOn = state.amplitudeGraphEnabled;
    var AMP_BAND_H = 22;
    var AMP_BAND_GAP = 2;
    var resultY = ampOn ? 24 + AMP_BAND_H + AMP_BAND_GAP : 24;
    var resultH = h - resultY - 6;
    var focused = SS.focusedTaskId();
    _timelineHitRects = [];

    if (ampOn) {
      var seriesByType = {};
      state.tasks.forEach(function (task) {
        if (!task.result || task.status === "cancelled") return;
        if (task.participant && task.participant !== state.selectedParticipant) return;
        if (task.type === "timelapse") return;
        // Boundaries are orientation scaffolding, not events — keep them out of
        // the event-density amplitude band (they render as flags above the timeline).
        if (task.type === "boundary") return;
        if (!seriesByType[task.type]) {
          seriesByType[task.type] = { key: task.type, color: taskTypeColor(task.type), timestamps: [] };
        }
        var dst = seriesByType[task.type].timestamps;
        var results = task.result || [];
        for (var ri = 0; ri < results.length; ri++) {
          var r = results[ri];
          var ts = r.timestamp !== undefined ? r.timestamp : r.start;
          if (ts !== undefined) dst.push(ts);
        }
      });
      var seriesList = Object.keys(seriesByType).map(function (k) { return seriesByType[k]; });
      var focusedTask = focused ? SS.findTask(focused) : null;
      var dimKey = focusedTask ? focusedTask.type : null;
      drawAmplitudeBands(ctx, {
        x: 0,
        y: 24,
        w: w,
        h: AMP_BAND_H,
        visStart: visStart,
        visEnd: visEnd,
        series: seriesList,
        binPx: 2,
        dimKey: dimKey,
      });
    }

    // Build excluded-timestamp lookup per task from cached events
    var excludedByTask = {};
    Object.keys(state.taskEvents).forEach(function (tid) {
      var exSet = {};
      (state.taskEvents[tid] || []).forEach(function (ev) {
        if (ev.excluded) exSet[ev.time_in.toFixed(2)] = true;
      });
      excludedByTask[tid] = exSet;
    });

    state.tasks.forEach(function (task) {
      if (!task.result || task.status === "cancelled") return;
      if (task.participant && task.participant !== state.selectedParticipant) return;
      var color = taskTypeColor(task.type);
      var dimmed = focused && task.id !== focused;
      var taskExcluded = excludedByTask[task.id] || {};
      if ((task.type === "color" || task.type === "inactivity") && task.status === "completed") {
        // Completed color: merged spans
        task.result.forEach(function (span) {
          var isExcluded = taskExcluded[span.start.toFixed(2)];
          ctx.fillStyle = hexToRgba(color, isExcluded ? 0.05 : (dimmed ? 0.10 : 0.35));
          var x1 = timeToX(span.start);
          var x2 = timeToX(span.end);
          var rw = Math.max(x2 - x1, 2);
          ctx.fillRect(x1, resultY, rw, resultH);
          _timelineHitRects.push({ x1: x1, x2: x1 + rw, y: resultY, h: resultH, task: task, result: span });
        });
      } else if (task.type === "timelapse") {
        // No timeline markers for timelapse
      } else if (task.type === "boundary") {
        // Boundaries are orientation scaffolding, not findings — rendered as
        // flags above the timeline (renderBoundaryFlags), not in-band ticks.
      } else {
        // Point markers (change, similarity, text, numbers, template, flow, scene, running color)
        ctx.lineWidth = 1.5;
        var results = task.result || [];
        results.forEach(function (r) {
          var ts = r.timestamp !== undefined ? r.timestamp : r.start;
          if (ts === undefined) return;
          var isExcluded = taskExcluded[ts.toFixed(2)];
          var sceneDimmed = task.type === "scene" && state.hoveredResultSceneName !== null
            && r.scene_name !== state.hoveredResultSceneName;
          if (isExcluded) {
            ctx.strokeStyle = hexToRgba(color, 0.15);
            ctx.setLineDash([3, 3]);
          } else if (dimmed || sceneDimmed) {
            ctx.strokeStyle = hexToRgba(color, 0.15);
            ctx.setLineDash([]);
          } else {
            ctx.strokeStyle = color;
            ctx.setLineDash([]);
          }
          var x = timeToX(ts);
          ctx.beginPath();
          ctx.moveTo(x, resultY);
          ctx.lineTo(x, resultY + resultH);
          ctx.stroke();
          ctx.setLineDash([]);
          _timelineHitRects.push({ x1: x - 3, x2: x + 3, y: resultY, h: resultH, task: task, result: r });
        });
      }
    });

    // Calibration pin ticks: small downward triangles just above the result
    // band, colored by polarity (green = positive, red = negative). The hovered
    // pin (cross-highlight from the tray) gets a fuller glyph.
    if (state.pins && state.pins.length) {
      var bodyStyle = getComputedStyle(document.body);
      var pinColors = {
        positive: bodyStyle.getPropertyValue("--color-pin-positive").trim() || "#4ade80",
        negative: bodyStyle.getPropertyValue("--color-pin-negative").trim() || "#ef4444",
      };
      var PIN_GLYPH_H = 7;
      var pinTop = Math.max(resultY - PIN_GLYPH_H - 1, 0);
      state.pins.forEach(function (pin) {
        var px = timeToX(pin.timestamp);
        if (px < 0 || px > w) return;
        var color = pinColors[pin.polarity] || "#888";
        var hovered = state.hoveredPinId === pin.id;
        ctx.fillStyle = hovered ? color : hexToRgba(color, 0.85);
        ctx.beginPath();
        ctx.moveTo(px - 4, pinTop);
        ctx.lineTo(px + 4, pinTop);
        ctx.lineTo(px, pinTop + PIN_GLYPH_H);
        ctx.closePath();
        ctx.fill();
        if (hovered) {
          ctx.strokeStyle = color;
          ctx.lineWidth = 1;
          ctx.beginPath();
          ctx.moveTo(px, pinTop + PIN_GLYPH_H);
          ctx.lineTo(px, resultY + resultH);
          ctx.stroke();
        }
      });
    }

    renderBoundaryFlags(visStart, visLen, w, excludedByTask, focused);
    renderTimelineLegend();
    renderPlayhead();
  }

  // Boundary events render as flag glyphs in #boundaryFlagRail, planted on top
  // of the timeline (outside the result band) — they are orientation
  // scaffolding, not clip candidates. Rebuilt on every pan/zoom/resize/focus.
  function renderBoundaryFlags(visStart, visLen, w, excludedByTask, focused) {
    var rail = qs("#boundaryFlagRail");
    if (!rail) return;
    rail.innerHTML = "";
    if (visLen <= 0) return;
    var color = taskTypeColor("boundary");
    var frag = document.createDocumentFragment();
    state.tasks.forEach(function (task) {
      if (task.type !== "boundary") return;
      if (!task.result || task.status === "cancelled") return;
      if (task.participant && task.participant !== state.selectedParticipant) return;
      var dimmed = focused && task.id !== focused;
      var taskExcluded = excludedByTask[task.id] || {};
      task.result.forEach(function (r) {
        var ts = r.timestamp !== undefined ? r.timestamp : r.start;
        if (ts === undefined) return;
        var x = ((ts - visStart) / visLen) * w;
        if (x < 0 || x > w) return;
        var flag = el("span", "boundary-flag");
        if (taskExcluded[ts.toFixed(2)]) flag.classList.add("excluded");
        else if (dimmed) flag.classList.add("dimmed");
        flag.style.left = x + "px";
        flag.style.color = color;
        flag.appendChild(el("span", "boundary-flag-icon"));
        // Per-flag hover/click: the rail itself is pointer-events:none, so each
        // flag (pointer-events:auto) owns its own listeners — leaving a flag into
        // empty rail space (or off-page) reliably fires its mouseleave. Hover only
        // redraws the playhead overlay, never the rail, so the element persists.
        wireBoundaryFlag(flag, task, r, ts);
        frag.appendChild(flag);
      });
    });
    rail.appendChild(frag);
  }

  function wireBoundaryFlag(flag, task, r, ts) {
    function enter(e) {
      showSsTooltip({ task: task, result: r }, e.clientX, e.clientY);
      if (state.hoveredBoundaryTs !== ts) {
        state.hoveredBoundaryTs = ts;
        renderPlayhead();
      }
    }
    flag.addEventListener("mouseenter", enter);
    flag.addEventListener("mousemove", enter);
    flag.addEventListener("mouseleave", function () {
      hideSsTooltip();
      if (state.hoveredBoundaryTs !== null) {
        state.hoveredBoundaryTs = null;
        renderPlayhead();
      }
    });
    flag.addEventListener("click", function () {
      state.resultOverlay = null;
      loadFrame(ts);
    });
  }

  function renderPlayhead() {
    var canvas = qs("#playheadCanvas");
    if (!canvas.width || !canvas.height) return;
    var ctx = canvas.getContext("2d");
    var w = canvas.width;
    var h = canvas.height;
    var dur = state.videoInfo ? state.videoInfo.duration : 0;
    ctx.clearRect(0, 0, w, h);
    if (dur <= 0) return;
    var visLen = dur / state.timelineZoom;
    var visStart = state.timelineOffset;

    // Transient locator: faint hairline under the hovered boundary flag. Drawn on
    // the playhead overlay (not the timeline canvas) so hovering never rebuilds the
    // flag rail — keeping each flag's mouseleave reliable.
    if (state.hoveredBoundaryTs !== null && state.hoveredBoundaryTs !== undefined) {
      var hbx = ((state.hoveredBoundaryTs - visStart) / visLen) * w;
      if (hbx >= 0 && hbx <= w) {
        ctx.strokeStyle = hexToRgba(taskTypeColor("boundary"), 0.4);
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(hbx, 0);
        ctx.lineTo(hbx, h);
        ctx.stroke();
      }
    }

    var px = ((state.currentTimestamp - visStart) / visLen) * w;
    var tc = getThemeColors();
    ctx.strokeStyle = tc.accent;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(px, 0);
    ctx.lineTo(px, h);
    ctx.stroke();
    ctx.fillStyle = tc.accent;
    ctx.beginPath();
    ctx.moveTo(px - 5, 0);
    ctx.lineTo(px + 5, 0);
    ctx.lineTo(px, 6);
    ctx.closePath();
    ctx.fill();
  }

  function hitTestTimeline(clientX, clientY) {
    var canvas = qs("#timelineCanvas");
    var rect = getTimelineRect(canvas);
    var mx = clientX - rect.left;
    var my = clientY - rect.top;
    for (var i = _timelineHitRects.length - 1; i >= 0; i--) {
      var hr = _timelineHitRects[i];
      if (mx >= hr.x1 && mx <= hr.x2 && my >= hr.y && my <= hr.y + hr.h) {
        return hr;
      }
    }
    return null;
  }

  function showSsTooltip(hit, clientX, clientY) {
    var tip = qs("#ssTooltip");
    if (!tip) return;
    var color = taskTypeColor(hit.task.type);
    tip.innerHTML = "";
    tip.style.borderLeft = "3px solid " + color;

    var header = el("div", "ss-tooltip-header");
    var icon = buildTypeIcon(hit.task.type);
    if (icon) {
      icon.style.color = color;
      icon.style.flexShrink = "0";
      header.appendChild(icon);
    }
    var label = hit.task.type.charAt(0).toUpperCase() + hit.task.type.slice(1);
    header.appendChild(el("strong", "", label));
    tip.appendChild(header);

    var r = hit.result;
    var timeStr;
    if (r.start !== undefined && r.end !== undefined) {
      timeStr = formatTime(r.start, { decimals: 1 }) + " \u2013 " + formatTime(r.end, { decimals: 1 });
    } else {
      var ts = r.timestamp !== undefined ? r.timestamp : r.start;
      timeStr = formatTime(ts, { decimals: 1 });
    }
    tip.appendChild(el("span", "ss-tooltip-time", timeStr));

    var details = el("div", "ss-tooltip-details");
    details.appendChild(el("span", "", hit.task.participant + " \u00b7 " + (hit.task.region || "")));
    if (hit.task.type === "inactivity" && r.duration !== undefined) {
      details.appendChild(el("span", "", "Duration: " + r.duration.toFixed(1) + "s"));
      if (r.avg_distance !== undefined) details.appendChild(el("span", "", "Avg distance: " + r.avg_distance));
    } else if (hit.task.type === "color" && r.duration !== undefined) {
      details.appendChild(el("span", "", "Duration: " + r.duration.toFixed(1) + "s"));
    } else if (hit.task.type === "change" && r.magnitude !== undefined) {
      details.appendChild(el("span", "", "Magnitude: " + (r.magnitude * 100).toFixed(1) + "%"));
    } else if (hit.task.type === "similarity" && r.score !== undefined) {
      details.appendChild(el("span", "", "Score: " + (r.score * 100).toFixed(1) + "%"));
    } else if (hit.task.type === "text" && r.text_found) {
      details.appendChild(el("span", "", "Found: " + r.text_found));
      if (r.confidence !== undefined) details.appendChild(el("span", "", "Confidence: " + (r.confidence * 100).toFixed(0) + "%"));
    } else if (hit.task.type === "numbers" && r.number_found !== undefined) {
      details.appendChild(el("span", "", "Found: " + r.number_found));
      if (r.confidence !== undefined) details.appendChild(el("span", "", "Confidence: " + (r.confidence * 100).toFixed(0) + "%"));
    } else if (hit.task.type === "template" && r.best_score !== undefined) {
      details.appendChild(el("span", "", "Score: " + (r.best_score * 100).toFixed(1) + "%"));
      if (r.match_count !== undefined) details.appendChild(el("span", "", "Matches: " + r.match_count));
    } else if (hit.task.type === "flow" && r.magnitude !== undefined) {
      details.appendChild(el("span", "", "Magnitude: " + r.magnitude.toFixed(2)));
      if (r.angle !== undefined) details.appendChild(el("span", "", "Direction: " + r.angle.toFixed(0) + "\u00b0"));
    } else if (hit.task.type === "scene" && r.scene_name) {
      details.appendChild(el("span", "", "Scene: " + r.scene_name));
      if (r.score !== undefined) details.appendChild(el("span", "", "Score: " + (r.score * 100).toFixed(1) + "%"));
    } else if (hit.task.type === "boundary") {
      if (r.scene_label) details.appendChild(el("span", "", "Enters: " + r.scene_label));
      if (r.distance !== undefined) details.appendChild(el("span", "", "Distance: " + r.distance));
    }
    tip.appendChild(details);

    tip.classList.remove("hidden");
    positionSsTooltip(tip, clientX, clientY);
  }

  function positionSsTooltip(tip, clientX, clientY) {
    var x = clientX + 12;
    var y = clientY + 12;
    var rect = tip.getBoundingClientRect();
    if (x + rect.width > window.innerWidth - 8) {
      x = clientX - rect.width - 12;
    }
    if (y + rect.height > window.innerHeight - 8) {
      y = clientY - rect.height - 12;
    }
    tip.style.left = x + "px";
    tip.style.top = y + "px";
  }

  function hideSsTooltip() {
    var tip = qs("#ssTooltip");
    if (tip) tip.classList.add("hidden");
  }


  function renderTimelineLegend() {
    var container = qs("#timelineLegend");
    container.innerHTML = "";
    var hasTypes = {};
    state.tasks.forEach(function (t) {
      if ((t.status === "completed" || t.status === "running") && t.result) hasTypes[t.type] = true;
    });
    var types = Object.keys(hasTypes);
    if (types.length === 0) return;
    var focused = SS.focusedTaskId();
    var focusedType = focused ? (SS.findTask(focused) || {}).type : null;
    types.forEach(function (type) {
      var item = el("span", "legend-item");
      var dot = el("span", "legend-dot");
      dot.style.background = taskTypeColor(type);
      item.appendChild(dot);
      item.appendChild(document.createTextNode(type));
      if (focusedType && type !== focusedType) {
        item.style.opacity = "0.3";
      }
      container.appendChild(item);
    });
  }


  SS.initTimeline = initTimeline;
  SS.renderTimeline = renderTimeline;
  SS.renderPlayhead = renderPlayhead;
  SS.updateMarkerInfo = updateMarkerInfo;
})();
