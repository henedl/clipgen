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
  var STEP_TRACK_H = 8; // minor/major tick strip just under the timestamps
  var CUT_TRACK_H = 26; // the single cuts track, directly under the ruler
  var LANE_GAP = 3;
  var ROW_H = 15;       // one marker sub-row (14px bar + 1px gap)
  var THUMB_ROW_H = 42; // marker sub-row with thumbnail strips (41px bar + 1px gap)
  var THUMB_CUT_H = 42; // cuts track height with thumbnail strips
  var MAX_LANE_ROWS = 8; // unfold ceiling; denser overlaps collapse onto the last row
  var EDGE_SLOP = 5;    // px hit zone around a cut edge
  var MIN_CUT_SECONDS = 0.2;
  var SOURCES = ["sheet", "screenspace", "transcript"];

  var _hitRects = [];          // marker + cut hover rects, rebuilt per render
  var _cachedRect = null;
  var _laneColors = null;
  var _chromeColors = null;

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

  // Subtle timeline chrome: alternating-row stripe fill + minor step-track tick
  // color. Cached like laneColors() and re-sampled on theme flips.
  function chromeColors() {
    if (!_chromeColors) {
      _chromeColors = {
        stripe: getCSSVar("--co-timeline-stripe", "rgba(255,255,255,0.02)"),
        grid: getCSSVar("--co-timeline-grid", "rgba(255,255,255,0.06)"),
      };
    }
    return _chromeColors;
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

  // Marker sub-rows and the cuts track always use their large (thumb-sized)
  // heights so each bar can fit a recognizable frame; the Thumbs toggle only
  // controls whether thumbnails are drawn, not the geometry. Only the
  // annotations lane stays flat (its spans are drawings, not video content).
  function laneRowH() { return THUMB_ROW_H; }
  function cutTrackH() { return THUMB_CUT_H; }

  // Annotation spans as packable {start, end, ann} entries for the lane.
  function annotationSpans() {
    var annotations = CO.participantAnnotations ? CO.participantAnnotations() : [];
    return annotations.map(function (ann) {
      return { start: ann.span.start, end: ann.span.end, ann: ann };
    });
  }

  // Vertical layout, driven by fold state: ruler → cuts track (the user's
  // working track) → the annotations lane right under it (both are Composer-
  // owned, editable content; only present when annotations exist) → one lane
  // per visible read-only source. Every lane folds to a single row and
  // unfolds to its minimal packed row count. The strip's height follows via
  // updateTimelineHeight(), not vice versa.
  function layout() {
    var cutY = RULER_H + STEP_TRACK_H + 2;
    var y = cutY + cutTrackH() + 4;
    var spans = annotationSpans();
    var annRows = !spans.length
      ? 0
      : state.laneFolds.annotations
        ? 1
        : neededRows(spans);
    var annotationsLane = { y: y, rows: annRows, h: annRows * ROW_H };
    if (annRows) y += annRows * ROW_H + LANE_GAP;
    var lanes = {};
    SOURCES.forEach(function (source) {
      var rows = laneRows(source);
      lanes[source] = { y: y, rows: rows, h: rows * laneRowH() };
      if (rows) y += rows * laneRowH() + LANE_GAP;
    });
    return {
      cutY: cutY,
      cutH: cutTrackH(),
      lanes: lanes,
      annotationsLane: annotationsLane,
      canvasH: y + 2,
    };
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

  // Pan the viewport minimally so *t* sits inside the visible window with a
  // small edge margin. No-op when fully zoomed out (the whole timeline is
  // visible) or while the user is mid pan/scrub drag (don't fight the gesture).
  function revealTime(t) {
    if (state.zoom <= 1 || !state.duration) return;
    if (state.dragging) return;
    var visLen = state.duration / state.zoom;
    var margin = visLen * 0.12; // breathing room at the edges
    var lo = state.offset + margin;
    var hi = state.offset + visLen - margin;
    if (t >= lo && t <= hi) return; // already comfortably in view
    var target = t < lo ? t - margin : t - visLen + margin;
    target = clamp(target, 0, Math.max(0, state.duration - visLen));
    if (target === state.offset) return; // viewport already pinned — no redraw
    state.offset = target;
    renderTimeline();
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

  // Minor + major tick marks in the step-track strip under the timestamps.
  // Minors are subdivided only while they stay legible (>= ~8px apart), so the
  // ruler declutters automatically as the view zooms out.
  function drawStepTrack(ctx, tx, vis, interval, w, chrome, tc) {
    var majorPx = (interval / vis.len) * w;
    var minorDiv = majorPx / 5 >= 8 ? 5 : majorPx / 2 >= 8 ? 2 : 1;
    if (minorDiv > 1) {
      var minorInt = interval / minorDiv;
      ctx.fillStyle = chrome.grid;
      for (var t = Math.ceil(vis.start / minorInt) * minorInt;
           t <= vis.end + 1e-6; t += minorInt) {
        ctx.fillRect(Math.round(tx(t)), RULER_H, 1, 3);
      }
    }
    ctx.fillStyle = tc.textDim;
    for (var T = Math.ceil(vis.start / interval) * interval;
         T <= vis.end + 1e-6; T += interval) {
      ctx.fillRect(Math.round(tx(T)), RULER_H, 1, 7);
    }
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
    var L = layout();
    var interval = niceTimeInterval(vis.len, { maxTicks: 20 });
    var chrome = chromeColors();

    // Subtle alternating row bands, ledger-style: fill every other lane band
    // (cuts → annotations → each visible source lane) under the content.
    var bands = [{ y: L.cutY, h: L.cutH }];
    if (L.annotationsLane.rows) {
      bands.push({ y: L.annotationsLane.y, h: L.annotationsLane.h });
    }
    SOURCES.forEach(function (source) {
      var lane = L.lanes[source];
      if (lane.rows) bands.push({ y: lane.y, h: lane.h });
    });
    ctx.fillStyle = chrome.stripe;
    bands.forEach(function (b, i) {
      if (i % 2 === 1) ctx.fillRect(0, b.y, w, b.h);
    });

    // Step track: minor + major tick marks just under the timestamps (a
    // video-editor-style ruler), instead of full-height gridlines that clutter
    // the lanes. Majors align with the ruler labels; minors subdivide them.
    drawStepTrack(ctx, tx, vis, interval, w, chrome, tc);

    // Labels only — the ruler's own ticks are suppressed (tickHeight 0) so the
    // step track owns the tick marks.
    drawTimelineRuler(ctx, {
      visStart: vis.start,
      visEnd: vis.end,
      interval: interval,
      timeToX: tx,
      colors: { border: tc.border, textDim: tc.textDim, fontMono: tc.fontMono },
      tickHeight: 0,
      labelY: RULER_H - 6,
      format: formatDuration,
    });

    // Cuts track — the single working track for all in/out pairs, directly
    // under the ruler so committed cuts land visibly on the timeline. Its
    // boundary reads from the alternating band, so no separator outline.
    // Chronological numbering shared with the cut list's index badges.
    var cutIndexById = {};
    if (CO.sortedCuts) {
      CO.sortedCuts().forEach(function (c, i) { cutIndexById[c.id] = i + 1; });
    }
    CO.participantCuts().forEach(function (cut) {
      if (cut.end < vis.start || cut.start > vis.end) return;
      var x1 = tx(cut.start);
      var x2 = tx(cut.end);
      var rw = Math.max(x2 - x1, 2);
      var selected = cut.id === state.selectedCutId;
      ctx.fillStyle = hexToRgba(tc.accent, selected ? 0.45 : 0.25);
      ctx.fillRect(x1, L.cutY, rw, L.cutH);
      if (state.markerThumbnails && CO.drawMarkerThumbs) {
        CO.drawMarkerThumbs(ctx, "cut:" + cut.id, cut.start, cut.end,
          x1, L.cutY, rw, L.cutH, tc.accent);
      }
      // Edge handles
      ctx.fillStyle = selected ? tc.accent : hexToRgba(tc.accent, 0.8);
      ctx.fillRect(x1 - 1, L.cutY, 3, L.cutH);
      ctx.fillRect(x2 - 2, L.cutY, 3, L.cutH);
      if (selected) {
        ctx.strokeStyle = tc.accent;
        ctx.strokeRect(x1 + 0.5, L.cutY + 0.5, rw - 1, L.cutH - 1);
      }
      // Index badge, top-left (skipped on slivers too narrow to carry it).
      var cutIdx = cutIndexById[cut.id];
      if (cutIdx && rw >= 18) {
        var label = String(cutIdx);
        ctx.font = "600 9px " + tc.fontMono;
        var bw = ctx.measureText(label).width + 6;
        ctx.fillStyle = hexToRgba(tc.accent, 0.9);
        ctx.fillRect(x1 + 2, L.cutY + 2, bw, 11);
        ctx.fillStyle = tc.bg;
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.fillText(label, x1 + 2 + bw / 2, L.cutY + 8);
        ctx.textAlign = "start";
        ctx.textBaseline = "alphabetic";
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
      var rowH = laneRowH();
      assignRows(markers, lane.rows).forEach(function (m) {
        if (m.end < vis.start || m.start > vis.end) return;
        var x1 = tx(m.start);
        var x2 = tx(Math.max(m.end, m.start));
        var rw = Math.max(x2 - x1, 2);
        var y = lane.y + m._row * rowH;
        ctx.fillStyle = hexToRgba(color, 0.55);
        ctx.fillRect(x1, y, rw, rowH - 1);
        if (state.markerThumbnails && CO.drawMarkerThumbs) {
          CO.drawMarkerThumbs(ctx, m.key, m.start, m.end, x1, y, rw, rowH - 1, color);
        }
        // Trimmed affordance: a solid underline in the lane color marks a
        // marker whose span deviates from its source (right-click resets).
        // Drawn after any thumbnails so it stays visible on top.
        if (m.trimmed) {
          ctx.fillStyle = color;
          ctx.fillRect(x1, y + rowH - 3, rw, 2);
        }
        _hitRects.push({ x1: x1, x2: x1 + rw, y: y, h: rowH - 1, marker: m });
      });
    });

    // Annotations lane (accent-colored spans; selection matches the overlay).
    // Dimmed — not skipped — while the layer is hidden, so the spans stay
    // findable and draggable.
    var annLane = L.annotationsLane;
    if (annLane.rows) {
      var annDim = state.annHidden ? 0.3 : 1;
      assignRows(annotationSpans(), annLane.rows).forEach(function (entry) {
        if (entry.end < vis.start || entry.start > vis.end) return;
        var ax1 = tx(entry.start);
        var ax2 = tx(entry.end);
        var arw = Math.max(ax2 - ax1, 2);
        var ay = annLane.y + entry._row * ROW_H;
        var annSelected = CO.isAnnotationSelected(entry.ann.id);
        var annColor = (entry.ann.style && entry.ann.style.color) || tc.accent;
        ctx.fillStyle = hexToRgba(annColor, (annSelected ? 0.8 : 0.5) * annDim);
        ctx.fillRect(ax1, ay, arw, ROW_H - 1);
        if (annSelected) {
          ctx.strokeStyle = hexToRgba(annColor, annDim);
          ctx.lineWidth = 1;
          ctx.strokeRect(ax1 + 0.5, ay + 0.5, arw - 1, ROW_H - 2);
        }
        _hitRects.push({
          x1: ax1, x2: ax1 + arw, y: ay, h: ROW_H - 1,
          annotation: entry.ann, start: entry.start, end: entry.end,
        });
      });
    }

    renderLaneRail(L);
    renderPlayhead();
  }

  // Per-lane fold buttons in the DOM rail (rebuilt on every render, like
  // screenspace's boundary flags). Hidden/empty lanes get no button.
  function renderLaneRail(L) {
    var rail = qs("#coLaneRail");
    if (!rail) return;
    rail.innerHTML = "";
    if (!state.duration) return;
    var frag = document.createDocumentFragment();

    function foldButton(source, laneY, laneH) {
      var folded = !!state.laneFolds[source];
      var btn = el("button", "co-lane-fold");
      btn.type = "button";
      btn.setAttribute("data-source", source);
      btn.title = (folded ? "Unfold " : "Fold ") + source + " lane";
      btn.setAttribute("aria-label", btn.title);
      // Vertically center in the lane (CSS translateY(-50%) does the rest).
      btn.style.top = (laneY + laneH / 2) + "px";
      btn.appendChild(el("span",
        "co-btn-icon " + (folded ? "co-icon-fold-closed" : "co-icon-fold-open")));
      btn.addEventListener("click", function (e) {
        e.stopPropagation();
        state.laneFolds[source] = !state.laneFolds[source];
        if (CO.persistLaneUi) CO.persistLaneUi();
        updateTimelineHeight();
        renderTimeline();
      });
      return btn;
    }

    SOURCES.forEach(function (source) {
      var lane = L.lanes[source];
      if (!lane.rows) return;
      frag.appendChild(foldButton(source, lane.y, lane.h));
    });
    if (L.annotationsLane.rows) {
      frag.appendChild(
        foldButton("annotations", L.annotationsLane.y, L.annotationsLane.h));
    }
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

  // Lane-edge hit: like cut edges, but against the marker/annotation hit
  // rects (their vertical band is per-row). Returns {marker?, annotation?,
  // edge} or null.
  function hitTestLaneEdge(clientX, clientY) {
    var rect = getRect();
    var mx = clientX - rect.left;
    var my = clientY - rect.top;
    var best = null;
    var bestDist = EDGE_SLOP + 1;
    for (var i = _hitRects.length - 1; i >= 0; i--) {
      var hr = _hitRects[i];
      if (!hr.marker && !hr.annotation) continue;
      if (my < hr.y || my > hr.y + hr.h) continue;
      // Point markers (start === end) have no meaningful edges to trim.
      if (hr.marker && hr.marker.end <= hr.marker.start) continue;
      var startDist = Math.abs(mx - hr.x1);
      var endDist = Math.abs(mx - hr.x2);
      if (startDist <= EDGE_SLOP && startDist < bestDist) {
        best = { marker: hr.marker, annotation: hr.annotation, edge: "start" };
        bestDist = startDist;
      }
      if (endDist <= EDGE_SLOP && endDist < bestDist) {
        best = { marker: hr.marker, annotation: hr.annotation, edge: "end" };
        bestDist = endDist;
      }
    }
    return best;
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
      if (m.trimmed) {
        tip.appendChild(el("span", "co-tooltip-time",
          "trimmed from " + formatTime(m.origStart, { decimals: 1 }) +
          " – " + formatTime(m.origEnd, { decimals: 1 }) +
          " · right-click to reset"));
      }
      if (m.label) tip.appendChild(el("span", "co-tooltip-label", m.label));
    } else if (hit.cut) {
      var tc = getCanvasThemeColors();
      tip.style.borderLeft = "3px solid " + tc.accent;
      tip.appendChild(el("span", "co-tooltip-source", "cut"));
      tip.appendChild(el("span", "co-tooltip-time",
        formatTime(hit.cut.start, { decimals: 1 }) + " – " +
        formatTime(hit.cut.end, { decimals: 1 }) +
        " (" + formatDuration(hit.cut.end - hit.cut.start) + ")"));
    } else if (hit.annotation) {
      var ann = hit.annotation;
      var annColor = (ann.style && ann.style.color) ||
        getCanvasThemeColors().accent;
      tip.style.borderLeft = "3px solid " + annColor;
      tip.appendChild(el("span", "co-tooltip-source", "annotation · " + ann.type));
      tip.appendChild(el("span", "co-tooltip-time",
        formatTime(ann.span.start, { decimals: 1 }) + " – " +
        formatTime(ann.span.end, { decimals: 1 })));
      if (ann.type === "text" && ann.geometry.text) {
        tip.appendChild(el("span", "co-tooltip-label", ann.geometry.text));
      }
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
      // Set early so the scrub branch's immediate seek sees an active drag and
      // revealTime() no-ops (the clicked time is on-screen anyway).
      state.dragging = true;
      hideTooltip();
      if (CO.scrubHoverEnd) CO.scrubHoverEnd();
      var edgeHit = hitTestCutEdge(e.clientX, e.clientY);
      var laneEdgeHit = edgeHit ? null : hitTestLaneEdge(e.clientX, e.clientY);
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
      } else if (laneEdgeHit && laneEdgeHit.marker) {
        // Marker trim drag: same gesture as a cut edge, committed as a
        // non-destructive trim override on pointer-up.
        drag = {
          type: "marker-edge",
          marker: laneEdgeHit.marker,
          edge: laneEdgeHit.edge,
          origStart: laneEdgeHit.marker.start,
          origEnd: laneEdgeHit.marker.end,
          // The trim in force before this drag (null = untrimmed) — the undo
          // payload, captured before the drag mutates anything.
          beforeTrim: state.trims[laneEdgeHit.marker.key]
            ? {
                start: state.trims[laneEdgeHit.marker.key].start,
                end: state.trims[laneEdgeHit.marker.key].end,
              }
            : null,
          moved: false,
        };
        canvas.classList.add("co-drag-edge");
      } else if (laneEdgeHit && laneEdgeHit.annotation) {
        // Annotation span drag: adjusts when the annotation is visible.
        drag = {
          type: "ann-edge",
          ann: laneEdgeHit.annotation,
          edge: laneEdgeHit.edge,
          beforeSpan: {
            start: laneEdgeHit.annotation.span.start,
            end: laneEdgeHit.annotation.span.end,
          },
          moved: false,
        };
        CO.selectAnnotation(laneEdgeHit.annotation.id);
        canvas.classList.add("co-drag-edge");
      } else {
        var bodyCut = hitTestCutBody(e.clientX, e.clientY);
        var laneHit = bodyCut ? null : hitTest(e.clientX, e.clientY);
        var markerHit = laneHit && laneHit.marker ? laneHit : null;
        var annHit = laneHit && laneHit.annotation ? laneHit : null;
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
        } else if (markerHit) {
          // Click a marker → seek to its in point (same as the cut track).
          drag = { type: "marker", marker: markerHit.marker, startX: e.clientX, moved: false };
        } else if (annHit) {
          // Whole-span translate, mirroring the cut-body drag; a no-move
          // release stays a click (select + seek).
          drag = {
            type: "annotation",
            ann: annHit.annotation,
            startX: e.clientX,
            origStart: annHit.annotation.span.start,
            origEnd: annHit.annotation.span.end,
            moved: false,
          };
          CO.selectAnnotation(annHit.annotation.id);
        } else if (state.zoom > 1) {
          drag = { type: "pan", startX: e.clientX, startOffset: state.offset, moved: false };
        } else {
          drag = { type: "scrub", moved: false };
          var ts = xToTime(e.clientX);
          if (ts !== null) CO.seekVideo(ts);
        }
      }
      canvas.setPointerCapture(e.pointerId);
    });

    canvas.addEventListener("pointermove", function (e) {
      if (!drag) {
        // Hover: edge cursor affordance + tooltip + opt-in audio scrub.
        if (!state.duration) return;
        var edge = hitTestCutEdge(e.clientX, e.clientY) ||
          hitTestLaneEdge(e.clientX, e.clientY);
        canvas.classList.toggle("co-drag-edge", !!edge);
        var hit = edge ? null : hitTest(e.clientX, e.clientY);
        if (hit) showTooltip(hit, e.clientX, e.clientY);
        else hideTooltip();
        var span = hit && (hit.marker || hit.cut);
        if (state.markerAudioScrub && span && span.end > span.start &&
            CO.scrubHoverMove) {
          var ts0 = xToTime(e.clientX);
          var frac = clamp((ts0 - span.start) / (span.end - span.start), 0, 1);
          CO.scrubHoverMove(hit.marker ? hit.marker.key : "cut:" + hit.cut.id,
            span.start, span.end, frac,
            { left: hit.x1, top: hit.y, width: hit.x2 - hit.x1, height: hit.h });
        } else if (CO.scrubHoverEnd) {
          CO.scrubHoverEnd();
        }
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
      } else if (drag.type === "marker-edge") {
        ts = xToTime(e.clientX);
        if (ts === null) return;
        drag.moved = true;
        var marker = drag.marker;
        if (drag.edge === "start") {
          marker.start = clamp(ts, 0, marker.end - MIN_CUT_SECONDS);
        } else {
          marker.end = clamp(ts, marker.start + MIN_CUT_SECONDS, state.duration);
        }
        scheduleRender();
      } else if (drag.type === "ann-edge") {
        ts = xToTime(e.clientX);
        if (ts === null) return;
        drag.moved = true;
        var span = drag.ann.span;
        if (drag.edge === "start") {
          span.start = clamp(ts, 0, span.end - MIN_CUT_SECONDS);
        } else {
          span.end = clamp(ts, span.start + MIN_CUT_SECONDS, state.duration);
        }
        scheduleRender();
      } else if (drag.type === "body") {
        var visLen = state.duration / state.zoom;
        var rect = getRect();
        var dt = ((e.clientX - drag.startX) / rect.width) * visLen;
        if (Math.abs(e.clientX - drag.startX) > 3) drag.moved = true;
        if (!drag.moved) return;
        var bodyLen = drag.origEnd - drag.origStart;
        var newStart = clamp(drag.origStart + dt, 0, state.duration - bodyLen);
        drag.cut.start = newStart;
        drag.cut.end = newStart + bodyLen;
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
      } else if (drag.type === "annotation") {
        // Translate the whole visibility span, preserving its length.
        var visLenA = state.duration / state.zoom;
        var rectA = getRect();
        var dtA = ((e.clientX - drag.startX) / rectA.width) * visLenA;
        if (Math.abs(e.clientX - drag.startX) > 3) drag.moved = true;
        if (!drag.moved) return;
        var spanLen = drag.origEnd - drag.origStart;
        var annStart = clamp(drag.origStart + dtA, 0, state.duration - spanLen);
        drag.ann.span.start = annStart;
        drag.ann.span.end = annStart + spanLen;
        canvas.classList.add("co-drag-body");
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
      } else if (d.type === "marker-edge" && d.moved) {
        CO.commitMarkerTrim(d.marker, d.beforeTrim,
          { start: d.origStart, end: d.origEnd });
      } else if (d.type === "marker-edge" && !d.moved) {
        // A no-move edge grab restores the pre-drag values (nothing changed).
        d.marker.start = d.origStart;
        d.marker.end = d.origEnd;
      } else if (d.type === "ann-edge" && d.moved) {
        CO.commitAnnotationField(d.ann, "span", d.beforeSpan);
      } else if (d.type === "ann-edge" && !d.moved) {
        d.ann.span.start = d.beforeSpan.start;
        d.ann.span.end = d.beforeSpan.end;
      } else if (d.type === "annotation" && d.moved) {
        CO.commitAnnotationField(d.ann, "span",
          { start: d.origStart, end: d.origEnd });
      } else if (d.type === "body" && !d.moved) {
        CO.seekVideo(d.cut.start);
      } else if (d.type === "marker" && !d.moved) {
        CO.seekVideo(d.marker.start);
      } else if (d.type === "annotation" && !d.moved) {
        CO.seekVideo(d.ann.span.start);
      }
    }
    canvas.addEventListener("pointerup", endDrag);
    canvas.addEventListener("pointercancel", endDrag);

    // Double-click on empty timeline space sets the pending in point, then a
    // second double-click commits the out point (config-gated; the preceding
    // single clicks already scrubbed the playhead to the clicked time).
    canvas.addEventListener("dblclick", function (e) {
      if (!CLIPGEN_CONFIG.composerDoubleClickCuts) return;
      if (!state.duration || !state.participant) return;
      if (hitTestCutEdge(e.clientX, e.clientY) ||
          hitTestLaneEdge(e.clientX, e.clientY) ||
          hitTestCutBody(e.clientX, e.clientY) ||
          hitTest(e.clientX, e.clientY)) {
        return; // clicks on cuts/markers/annotations keep their own semantics
      }
      var ts = xToTime(e.clientX);
      if (ts === null) return;
      if (CO.seekVideo) CO.seekVideo(ts);
      if (state.pendingIn === null) {
        if (CO.setInPoint) CO.setInPoint();
      } else if (CO.setOutPoint) {
        CO.setOutPoint();
      }
    });

    // Right-click a trimmed marker to reset its trim (tooltip advertises this).
    canvas.addEventListener("contextmenu", function (e) {
      var hit = hitTest(e.clientX, e.clientY);
      if (hit && hit.marker && hit.marker.trimmed) {
        e.preventDefault();
        hideTooltip();
        CO.resetTrim(hit.marker);
      }
    });

    canvas.addEventListener("mouseleave", function () {
      hideTooltip();
      if (CO.scrubHoverEnd) CO.scrubHoverEnd();
    });

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
  CO.revealTime = revealTime;
  CO.renderPlayhead = renderPlayhead;
  CO.updateTimelineHeight = updateTimelineHeight;
  // Theme flips resample the --stream-* lane + chrome colors on the next render.
  CO.invalidateLaneColors = function () { _laneColors = null; _chromeColors = null; };
})();
