/* clipgen Composer — timeline satellite (forked from screenspace-timeline.js).
 *
 * Owns the timeline canvas pair: the ruler, zoom/pan, click-to-seek, the three
 * marker lanes (sheet / screenspace / transcript, colored by the shared
 * --stream-* tokens), the cuts band with draggable edge/body handles, hover
 * tooltips, and the playhead overlay (playhead + pending in-point). Reads and
 * mutates hub state through window.ClipgenComposer (CO); commits finished
 * drags through CO.commitCutTimes / CO.selectCut. drawTimelineRuler /
 * niceTimeInterval / getCanvasThemeColors / getCSSVar / formatTime /
 * formatDuration / clamp / hexToRgba / el / qs are ambient utils.js globals.
 */
(function () {
  "use strict";

  var CO = window.ClipgenComposer;
  var state = CO.state;

  var RULER_H = 18;
  var CUT_TRACK_H = 26; // the single cuts track, directly under the ruler
  var LANE_GAP = 3;
  var ROW_H = 15;       // one marker sub-row (14px bar + 1px gap)
  var MAX_LANE_ROWS = 8; // unfold ceiling; denser overlaps collapse onto the last row
  var EDGE_SLOP = 5;    // px hit zone around a cut edge
  var MIN_CUT_SECONDS = 0.2;
  var SOURCES = ["sheet", "screenspace", "transcript"];

  var _hitRects = [];          // marker + cut hover rects, rebuilt per render
  var _cachedRect = null;
  var _laneColors = null;

  function laneColors() {
    if (!_laneColors) {
      _laneColors = {
        sheet: getCSSVar("--stream-sheet", "#d97706"),
        screenspace: getCSSVar("--stream-screenspace", "#2563eb"),
        transcript: getCSSVar("--stream-transcript", "#16a34a"),
      };
    }
    return _laneColors;
  }

  function canvasEl() { return qs("#coTimelineCanvas"); }

  // Minimal sub-row count for a lane's markers: greedy interval packing over
  // markers sorted by start yields the optimal (lowest) row count.
  function neededRows(markers) {
    var sorted = markers.slice().sort(function (a, b) { return a.start - b.start; });
    var rowEnds = [];
    sorted.forEach(function (m) {
      for (var r = 0; r < rowEnds.length; r++) {
        if (rowEnds[r] <= m.start) {
          rowEnds[r] = m.end;
          return;
        }
      }
      rowEnds.push(m.end);
    });
    return Math.min(Math.max(rowEnds.length, 1), MAX_LANE_ROWS);
  }

  function laneRows(source) {
    if (!state.sourceToggles[source]) return 0; // hidden lane occupies no space
    if (state.laneFolds[source]) return 1;
    return neededRows(state.markers[source] || []);
  }

  // Vertical layout, driven by fold state: ruler → cuts track (the user's
  // working track) → one lane per visible source, each sized to its row count.
  // The strip's height follows via updateTimelineHeight(), not vice versa.
  function layout() {
    var cutY = RULER_H + 2;
    var y = cutY + CUT_TRACK_H + 4;
    var lanes = {};
    SOURCES.forEach(function (source) {
      var rows = laneRows(source);
      lanes[source] = { y: y, rows: rows, h: rows * ROW_H };
      if (rows) y += rows * ROW_H + LANE_GAP;
    });
    return { cutY: cutY, cutH: CUT_TRACK_H, lanes: lanes, canvasH: y + 2 };
  }

  // Grow/shrink the whole timeline strip to fit the current fold state by
  // writing the shared --co-timeline-height var (both the shell and the strip
  // are sized from it); the wrapper's ResizeObserver then re-renders.
  var _sectionExtra = null;

  function updateTimelineHeight() {
    var section = qs("#coTimelineSection");
    var wrapper = qs("#coTimelineWrapper");
    if (_sectionExtra === null) {
      _sectionExtra = Math.max(section.offsetHeight - wrapper.offsetHeight, 24);
    }
    var target = layout().canvasH + _sectionExtra;
    var current = parseFloat(
      getComputedStyle(document.documentElement)
        .getPropertyValue("--co-timeline-height")
    );
    if (Math.abs(target - current) < 1) return;
    document.documentElement.style.setProperty(
      "--co-timeline-height", target + "px"
    );
    // Fallback when no ResizeObserver is wired (old engines): resize directly.
    if (typeof ResizeObserver !== "function") sizeCanvases();
  }

  function getRect() {
    if (!_cachedRect) _cachedRect = canvasEl().getBoundingClientRect();
    return _cachedRect;
  }

  function visWindow() {
    var visLen = state.duration / state.zoom;
    return { start: state.offset, len: visLen, end: state.offset + visLen };
  }

  function timeToX(t, w) {
    var vis = visWindow();
    return ((t - vis.start) / vis.len) * w;
  }

  function xToTime(clientX) {
    if (!state.duration) return null;
    var rect = getRect();
    var frac = clamp((clientX - rect.left) / rect.width, 0, 1);
    var vis = visWindow();
    return vis.start + frac * vis.len;
  }

  function clampOffset() {
    var visLen = state.duration / state.zoom;
    state.offset = clamp(state.offset, 0, Math.max(0, state.duration - visLen));
  }

  function sizeCanvases() {
    var canvas = canvasEl();
    _cachedRect = null;
    var rect = canvas.getBoundingClientRect();
    canvas.width = Math.floor(rect.width);
    canvas.height = Math.floor(rect.height);
    var ph = qs("#coPlayheadCanvas");
    ph.width = canvas.width;
    ph.height = canvas.height;
    renderTimeline();
  }

  // ---- Rendering ----

  // Greedy sub-row packing inside one source lane: overlapping markers go to
  // the first free sub-row; overflow past maxRows collapses onto the last.
  function assignRows(markers, maxRows) {
    var sorted = markers.slice().sort(function (a, b) { return a.start - b.start; });
    var rowEnds = [];
    sorted.forEach(function (m) {
      var placed = false;
      for (var r = 0; r < rowEnds.length && r < maxRows; r++) {
        if (rowEnds[r] <= m.start) {
          m._row = r;
          rowEnds[r] = m.end;
          placed = true;
          break;
        }
      }
      if (!placed) {
        if (rowEnds.length < maxRows) {
          m._row = rowEnds.length;
          rowEnds.push(m.end);
        } else {
          m._row = maxRows - 1;
          rowEnds[maxRows - 1] = Math.max(rowEnds[maxRows - 1], m.end);
        }
      }
    });
    return sorted;
  }

  function renderTimeline() {
    var canvas = canvasEl();
    if (!canvas.width || !canvas.height) return;
    var ctx = canvas.getContext("2d");
    var w = canvas.width;
    var h = canvas.height;
    var tc = getCanvasThemeColors();

    ctx.clearRect(0, 0, w, h);
    ctx.fillStyle = tc.surfaceAlt;
    ctx.fillRect(0, 0, w, h);
    _hitRects = [];

    if (!state.duration) {
      ctx.fillStyle = tc.textDim;
      ctx.font = "12px " + tc.fontMono;
      ctx.textAlign = "center";
      ctx.fillText("Select a participant to load the timeline", w / 2, h / 2 + 4);
      ctx.textAlign = "start";
      renderPlayhead();
      return;
    }

    var vis = visWindow();
    function tx(t) { return timeToX(t, w); }

    drawTimelineRuler(ctx, {
      visStart: vis.start,
      visEnd: vis.end,
      interval: niceTimeInterval(vis.len, { maxTicks: 20 }),
      timeToX: tx,
      colors: { border: tc.border, textDim: tc.textDim, fontMono: tc.fontMono },
      tickHeight: 8,
      labelY: RULER_H - 2,
      format: formatDuration,
    });

    var L = layout();

    // Cuts track — the single working track for all in/out pairs, directly
    // under the ruler so committed cuts land visibly on the timeline.
    ctx.strokeStyle = tc.border;
    ctx.lineWidth = 1;
    ctx.strokeRect(0.5, L.cutY - 0.5, w - 1, L.cutH + 1);
    CO.participantCuts().forEach(function (cut) {
      if (cut.end < vis.start || cut.start > vis.end) return;
      var x1 = tx(cut.start);
      var x2 = tx(cut.end);
      var rw = Math.max(x2 - x1, 2);
      var selected = cut.id === state.selectedCutId;
      ctx.fillStyle = hexToRgba(tc.accent, selected ? 0.45 : 0.25);
      ctx.fillRect(x1, L.cutY, rw, L.cutH);
      // Edge handles
      ctx.fillStyle = selected ? tc.accent : hexToRgba(tc.accent, 0.8);
      ctx.fillRect(x1 - 1, L.cutY, 3, L.cutH);
      ctx.fillRect(x2 - 2, L.cutY, 3, L.cutH);
      if (selected) {
        ctx.strokeStyle = tc.accent;
        ctx.strokeRect(x1 + 0.5, L.cutY + 0.5, rw - 1, L.cutH - 1);
      }
      _hitRects.push({ x1: x1, x2: x2, y: L.cutY, h: L.cutH, cut: cut });
    });

    // Marker lanes (one per visible source, below the cuts track)
    var colors = laneColors();
    SOURCES.forEach(function (source) {
      var lane = L.lanes[source];
      if (!lane.rows) return;
      var markers = state.markers[source] || [];
      if (!markers.length) return;
      var color = colors[source];
      assignRows(markers, lane.rows).forEach(function (m) {
        if (m.end < vis.start || m.start > vis.end) return;
        var x1 = tx(m.start);
        var x2 = tx(Math.max(m.end, m.start));
        var rw = Math.max(x2 - x1, 2);
        var y = lane.y + m._row * ROW_H;
        ctx.fillStyle = hexToRgba(color, 0.55);
        ctx.fillRect(x1, y, rw, ROW_H - 1);
        _hitRects.push({ x1: x1, x2: x1 + rw, y: y, h: ROW_H - 1, marker: m });
      });
    });

    renderLaneRail(L);
    renderPlayhead();
  }

  // Per-lane fold buttons in the DOM rail (rebuilt on every render, like
  // screenspace's boundary flags). Hidden lanes get no button.
  function renderLaneRail(L) {
    var rail = qs("#coLaneRail");
    if (!rail) return;
    rail.innerHTML = "";
    if (!state.duration) return;
    var frag = document.createDocumentFragment();
    SOURCES.forEach(function (source) {
      var lane = L.lanes[source];
      if (!lane.rows) return;
      var folded = !!state.laneFolds[source];
      var btn = el("button", "co-lane-fold");
      btn.type = "button";
      btn.setAttribute("data-source", source);
      btn.title = (folded ? "Unfold " : "Fold ") + source + " lane";
      btn.setAttribute("aria-label", btn.title);
      btn.style.top = lane.y + "px";
      btn.appendChild(el("span",
        "co-btn-icon " + (folded ? "co-icon-fold-closed" : "co-icon-fold-open")));
      btn.addEventListener("click", function (e) {
        e.stopPropagation();
        state.laneFolds[source] = !state.laneFolds[source];
        if (CO.persistLaneUi) CO.persistLaneUi();
        updateTimelineHeight();
        renderTimeline();
      });
      frag.appendChild(btn);
    });
    rail.appendChild(frag);
  }

  function renderPlayhead() {
    var canvas = qs("#coPlayheadCanvas");
    if (!canvas.width || !canvas.height) return;
    var ctx = canvas.getContext("2d");
    var w = canvas.width;
    var h = canvas.height;
    ctx.clearRect(0, 0, w, h);
    if (!state.duration) return;
    var tc = getCanvasThemeColors();

    // Pending in-point: dashed hairline until its out-point commits the cut.
    if (state.pendingIn !== null) {
      var ix = timeToX(state.pendingIn, w);
      if (ix >= 0 && ix <= w) {
        ctx.strokeStyle = tc.positive;
        ctx.lineWidth = 2;
        ctx.setLineDash([4, 3]);
        ctx.beginPath();
        ctx.moveTo(ix, 0);
        ctx.lineTo(ix, h);
        ctx.stroke();
        ctx.setLineDash([]);
      }
    }

    var px = timeToX(state.playhead, w);
    if (px < -6 || px > w + 6) return;
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

  // ---- Hit testing ----

  function hitTest(clientX, clientY) {
    var rect = getRect();
    var mx = clientX - rect.left;
    var my = clientY - rect.top;
    for (var i = _hitRects.length - 1; i >= 0; i--) {
      var hr = _hitRects[i];
      if (mx >= hr.x1 && mx <= hr.x2 && my >= hr.y && my <= hr.y + hr.h) return hr;
    }
    return null;
  }

  // Cut-edge hit: returns {cut, edge: "start"|"end"} when clientX/Y is within
  // EDGE_SLOP of a cut boundary inside the cuts band.
  function hitTestCutEdge(clientX, clientY) {
    var rect = getRect();
    var mx = clientX - rect.left;
    var my = clientY - rect.top;
    var w = canvasEl().width;
    var L = layout();
    if (my < L.cutY || my > L.cutY + L.cutH) return null;
    var cuts = CO.participantCuts();
    var best = null;
    var bestDist = EDGE_SLOP + 1;
    cuts.forEach(function (cut) {
      [["start", timeToX(cut.start, w)], ["end", timeToX(cut.end, w)]].forEach(function (pair) {
        var dist = Math.abs(mx - pair[1]);
        if (dist <= EDGE_SLOP && dist < bestDist) {
          best = { cut: cut, edge: pair[0] };
          bestDist = dist;
        }
      });
    });
    return best;
  }

  function hitTestCutBody(clientX, clientY) {
    var hit = hitTest(clientX, clientY);
    return hit && hit.cut ? hit.cut : null;
  }

  // ---- Tooltip ----

  function showTooltip(hit, clientX, clientY) {
    var tip = qs("#coTooltip");
    tip.innerHTML = "";
    var colors = laneColors();
    if (hit.marker) {
      var m = hit.marker;
      tip.style.borderLeft = "3px solid " + colors[m.source];
      var head = el("span", "co-tooltip-source", m.source + (m.eventType ? " · " + m.eventType : ""));
      tip.appendChild(head);
      tip.appendChild(el("span", "co-tooltip-time",
        formatTime(m.start, { decimals: 1 }) +
        (m.end > m.start ? " – " + formatTime(m.end, { decimals: 1 }) : "")));
      if (m.label) tip.appendChild(el("span", "co-tooltip-label", m.label));
    } else if (hit.cut) {
      var tc = getCanvasThemeColors();
      tip.style.borderLeft = "3px solid " + tc.accent;
      tip.appendChild(el("span", "co-tooltip-source", "cut"));
      tip.appendChild(el("span", "co-tooltip-time",
        formatTime(hit.cut.start, { decimals: 1 }) + " – " +
        formatTime(hit.cut.end, { decimals: 1 }) +
        " (" + formatDuration(hit.cut.end - hit.cut.start) + ")"));
    }
    tip.classList.remove("hidden");
    var x = clientX + 12;
    var y = clientY - tip.offsetHeight - 12;
    var tr = tip.getBoundingClientRect();
    if (x + tr.width > window.innerWidth - 8) x = clientX - tr.width - 12;
    if (y < 8) y = clientY + 12;
    tip.style.left = x + "px";
    tip.style.top = y + "px";
  }

  function hideTooltip() {
    qs("#coTooltip").classList.add("hidden");
  }

  // ---- Interaction ----

  function initTimeline() {
    var canvas = canvasEl();

    qs("#coZoomInBtn").appendChild(el("span", "co-btn-icon co-icon-zoom-in"));
    qs("#coZoomOutBtn").appendChild(el("span", "co-btn-icon co-icon-zoom-out"));

    sizeCanvases();
    if (typeof ResizeObserver === "function") {
      var obs = new ResizeObserver(function () { sizeCanvases(); });
      obs.observe(qs("#coTimelineWrapper"));
      window.addEventListener("pagehide", function () { obs.disconnect(); });
    } else {
      window.addEventListener("resize", sizeCanvases);
    }
    window.addEventListener("scroll", function () { _cachedRect = null; }, true);

    canvas.addEventListener("wheel", function (e) {
      e.preventDefault();
      if (!state.duration) return;
      var zoomFactor = e.deltaY < 0 ? 1.3 : 1 / 1.3;
      var mouseTs = xToTime(e.clientX);
      state.zoom = clamp(state.zoom * zoomFactor, 1, 200);
      if (mouseTs !== null && state.zoom > 1) {
        var rect = getRect();
        var frac = (e.clientX - rect.left) / rect.width;
        var visLen = state.duration / state.zoom;
        state.offset = clamp(mouseTs - frac * visLen, 0, state.duration - visLen);
      }
      if (state.zoom <= 1) state.offset = 0;
      renderTimeline();
    }, { passive: false });

    // One pointer gesture at a time: edge drag > body drag > pan (zoomed) > scrub.
    var drag = null;   // {type, cut?, edge?, startX, startOffset?, origStart?, origEnd?, moved}
    var _dragRaf = 0;

    function scheduleRender() {
      if (_dragRaf) return;
      _dragRaf = requestAnimationFrame(function () {
        _dragRaf = 0;
        renderTimeline();
      });
    }

    canvas.addEventListener("pointerdown", function (e) {
      if (!state.duration) return;
      hideTooltip();
      var edgeHit = hitTestCutEdge(e.clientX, e.clientY);
      if (edgeHit) {
        drag = {
          type: "edge",
          cut: edgeHit.cut,
          edge: edgeHit.edge,
          origStart: edgeHit.cut.start,
          origEnd: edgeHit.cut.end,
          moved: false,
        };
        CO.selectCut(edgeHit.cut.id);
        canvas.classList.add("co-drag-edge");
      } else {
        var bodyCut = hitTestCutBody(e.clientX, e.clientY);
        var markerHit = bodyCut ? null : hitTest(e.clientX, e.clientY);
        if (bodyCut) {
          drag = {
            type: "body",
            cut: bodyCut,
            startX: e.clientX,
            origStart: bodyCut.start,
            origEnd: bodyCut.end,
            moved: false,
          };
          CO.selectCut(bodyCut.id);
        } else if (markerHit && markerHit.marker) {
          // Click a marker → seek to its in point (same as the cut track).
          drag = { type: "marker", marker: markerHit.marker, startX: e.clientX, moved: false };
        } else if (state.zoom > 1) {
          drag = { type: "pan", startX: e.clientX, startOffset: state.offset, moved: false };
        } else {
          drag = { type: "scrub", moved: false };
          var ts = xToTime(e.clientX);
          if (ts !== null) CO.seekVideo(ts);
        }
      }
      state.dragging = true;
      canvas.setPointerCapture(e.pointerId);
    });

    canvas.addEventListener("pointermove", function (e) {
      if (!drag) {
        // Hover: edge cursor affordance + tooltip.
        if (!state.duration) return;
        var edge = hitTestCutEdge(e.clientX, e.clientY);
        canvas.classList.toggle("co-drag-edge", !!edge);
        var hit = edge ? null : hitTest(e.clientX, e.clientY);
        if (hit) showTooltip(hit, e.clientX, e.clientY);
        else hideTooltip();
        return;
      }
      var ts;
      if (drag.type === "edge") {
        ts = xToTime(e.clientX);
        if (ts === null) return;
        drag.moved = true;
        var cut = drag.cut;
        if (drag.edge === "start") {
          cut.start = clamp(ts, 0, cut.end - MIN_CUT_SECONDS);
        } else {
          cut.end = clamp(ts, cut.start + MIN_CUT_SECONDS, state.duration);
        }
        scheduleRender();
      } else if (drag.type === "body") {
        var visLen = state.duration / state.zoom;
        var rect = getRect();
        var dt = ((e.clientX - drag.startX) / rect.width) * visLen;
        if (Math.abs(e.clientX - drag.startX) > 3) drag.moved = true;
        if (!drag.moved) return;
        var span = drag.origEnd - drag.origStart;
        var newStart = clamp(drag.origStart + dt, 0, state.duration - span);
        drag.cut.start = newStart;
        drag.cut.end = newStart + span;
        canvas.classList.add("co-drag-body");
        scheduleRender();
      } else if (drag.type === "pan") {
        var rect2 = getRect();
        var visLen2 = state.duration / state.zoom;
        var dx = e.clientX - drag.startX;
        if (Math.abs(dx) > 3) drag.moved = true;
        state.offset = clamp(drag.startOffset - (dx / rect2.width) * visLen2,
          0, Math.max(0, state.duration - visLen2));
        scheduleRender();
      } else if (drag.type === "marker") {
        if (Math.abs(e.clientX - drag.startX) > 3) drag.moved = true;
      } else if (drag.type === "scrub") {
        drag.moved = true;
        ts = xToTime(e.clientX);
        if (ts !== null) CO.seekVideo(ts);
      }
    });

    function endDrag(e) {
      if (!drag) return;
      var d = drag;
      drag = null;
      state.dragging = false;
      canvas.classList.remove("co-drag-edge", "co-drag-body");
      if (canvas.hasPointerCapture && canvas.hasPointerCapture(e.pointerId)) {
        canvas.releasePointerCapture(e.pointerId);
      }
      if ((d.type === "edge" || d.type === "body") && d.moved) {
        CO.commitCutTimes(d.cut, { start: d.origStart, end: d.origEnd });
      } else if (d.type === "body" && !d.moved) {
        CO.seekVideo(d.cut.start);
      } else if (d.type === "marker" && !d.moved) {
        CO.seekVideo(d.marker.start);
      }
    }
    canvas.addEventListener("pointerup", endDrag);
    canvas.addEventListener("pointercancel", endDrag);

    canvas.addEventListener("mouseleave", hideTooltip);

    qs("#coZoomInBtn").addEventListener("click", function () {
      if (!state.duration) return;
      state.zoom = clamp(state.zoom * 1.5, 1, 200);
      clampOffset();
      renderTimeline();
    });
    qs("#coZoomOutBtn").addEventListener("click", function () {
      state.zoom = clamp(state.zoom / 1.5, 1, 200);
      if (state.zoom <= 1) state.offset = 0;
      clampOffset();
      renderTimeline();
    });
    qs("#coZoomResetBtn").addEventListener("click", function () {
      state.zoom = 1;
      state.offset = 0;
      renderTimeline();
    });
  }

  CO.initTimeline = initTimeline;
  CO.renderTimeline = renderTimeline;
  CO.renderPlayhead = renderPlayhead;
  CO.updateTimelineHeight = updateTimelineHeight;
  // Theme flips resample the --stream-* lane colors on the next render.
  CO.invalidateLaneColors = function () { _laneColors = null; };
})();
