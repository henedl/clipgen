/* clipgen design primitives — factory helpers shared across Studio,
 * Screenspace, and Transcripts. Each factory returns a DOM element
 * (not an HTML string) so callers can wire event handlers cleanly and
 * appending hundreds of cards in a frame stays cheap.
 *
 * Mutator helpers (e.g. setButtonProgress) take an existing element or
 * an element id and update its state in-place; they pair with a CSS
 * class defined in primitives.css.
 *
 * Hue resolution falls back to categoryHue(label) from utils.js when
 * the caller omits an explicit `hue`.
 *
 * Surface (window.ClipgenPrimitives):
 *   createFilterChip, createParticipantPill, createDensityTimeline,
 *   createSparkBars, createBtn, createSwimLane, createKpiCard,
 *   createCoverageMatrix, setButtonProgress
 */

(function (global) {
  function setDataset(el, dataset) {
    if (!dataset) return;
    Object.keys(dataset).forEach(function (k) {
      var v = dataset[k];
      if (v == null) return;
      el.dataset[k] = String(v);
    });
  }

  function resolveHue(label, hue) {
    if (typeof hue === "number") return hue;
    if (typeof global.categoryHue === "function") {
      return global.categoryHue(label || "");
    }
    return 220;
  }

  function fmtHue(hue, lightness, chroma, alpha) {
    var l = lightness != null ? lightness : 0.7;
    var c = chroma != null ? chroma : 0.16;
    if (alpha == null || alpha >= 1) {
      return "oklch(" + l + " " + c + " " + hue + ")";
    }
    return "oklch(" + l + " " + c + " " + hue + " / " + alpha + ")";
  }

  // ---- FilterChip ----

  function createFilterChip(opts) {
    opts = opts || {};
    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "filter-chip";
    var hue = resolveHue(opts.label, opts.hue);
    // `opts.color` (CSS color string) overrides the oklch(hue) path so
    // detector chips can pin to the canonical `--color-task-*` tokens.
    var dotColor = opts.color || fmtHue(hue);

    if (opts.dot !== false) {
      var dot = document.createElement("span");
      dot.className = "filter-chip-dot";
      dot.style.color = dotColor;
      btn.appendChild(dot);
    }

    var labelEl = document.createElement("span");
    labelEl.className = "filter-chip-label";
    labelEl.textContent = opts.label || "";
    btn.appendChild(labelEl);

    if (opts.count != null) {
      var countEl = document.createElement("span");
      countEl.className = "filter-chip-count";
      countEl.textContent = String(opts.count);
      btn.appendChild(countEl);
    }

    if (opts.active) {
      btn.classList.add("is-active");
      // Active state still uses oklch for the bg/border tints since they need
      // explicit alpha; the dot pulls from `dotColor` so the canonical token
      // shows on detector chips when `opts.color` is supplied.
      btn.style.setProperty("--cg-chip-fg", dotColor);
      btn.style.setProperty("--cg-chip-bg", fmtHue(hue, 0.7, 0.16, 0.12));
      btn.style.setProperty("--cg-chip-border", fmtHue(hue, 0.7, 0.16, 0.45));
    }

    setDataset(btn, opts.dataset);
    if (typeof opts.onClick === "function") {
      btn.addEventListener("click", opts.onClick);
    }
    return btn;
  }

  // ---- ParticipantPill ----

  function createParticipantPill(opts) {
    opts = opts || {};
    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "participant-pill cg-mono";
    btn.textContent = opts.id || "";
    if (opts.active) btn.classList.add("is-active");
    setDataset(btn, opts.dataset);
    if (typeof opts.onClick === "function") {
      btn.addEventListener("click", opts.onClick);
    }
    return btn;
  }

  // ---- DensityTimeline ----
  //
  // events: [{ t: 0..1, count, hue?, label? }, ...]
  // marker: 0..1 (optional), tickCount: number of tick labels (default 6),
  // durationSec: number — used to format tick labels as M:SS / H:MM:SS.

  function fmtTick(sec) {
    var s = Math.max(0, Math.floor(sec));
    var h = Math.floor(s / 3600);
    var m = Math.floor((s - h * 3600) / 60);
    var rs = s - h * 3600 - m * 60;
    if (h > 0) {
      return h + ":" + (m < 10 ? "0" + m : m) + ":" + (rs < 10 ? "0" + rs : rs);
    }
    return m + ":" + (rs < 10 ? "0" + rs : rs);
  }

  function createDensityTimeline(opts) {
    opts = opts || {};
    var wrap = document.createElement("div");
    wrap.className = "density-timeline";
    if (opts.height) wrap.style.setProperty("--cg-density-h", opts.height + "px");

    var ticks = document.createElement("div");
    ticks.className = "density-timeline-ticks cg-mono";
    wrap.appendChild(ticks);

    var track = document.createElement("div");
    track.className = "density-timeline-track";
    wrap.appendChild(track);

    function renderTicks(durationSec, tickCount) {
      ticks.textContent = "";
      var n = Math.max(2, tickCount || 6);
      for (var i = 0; i < n; i++) {
        var s = (durationSec || 0) * (i / (n - 1));
        var t = document.createElement("span");
        t.textContent = fmtTick(s);
        ticks.appendChild(t);
      }
    }

    function renderBars(events, marker) {
      track.textContent = "";
      var max = 1;
      (events || []).forEach(function (e) {
        if (e && typeof e.count === "number" && e.count > max) max = e.count;
      });
      (events || []).forEach(function (e, idx) {
        if (!e || typeof e.t !== "number") return;
        var bar = document.createElement("div");
        bar.className = "density-timeline-bar";
        var hue = resolveHue(e.label, e.hue);
        var alpha = Math.min(1, 0.35 + 0.6 * ((e.count || 1) / max));
        // Span from start to end so bar width reflects clip length; `min-width`
        // in CSS keeps short spans visible and clickable. Zero-span events (e.g.
        // unpadded transcript bookmarks) have no positive width, so center the
        // min-width marker on the timestamp like createSwimLane — but via a left
        // offset (half the 4px min-width) rather than translateX, so the
        // class-based hover scaleY transform is left intact.
        var widthPct = typeof e.tEnd === "number" && e.tEnd > e.t ? (e.tEnd - e.t) * 100 : 0;
        if (widthPct > 0) {
          bar.style.left = e.t * 100 + "%";
          bar.style.width = widthPct + "%";
        } else {
          bar.style.left = "calc(" + e.t * 100 + "% - 2px)";
        }
        // `e.color` (CSS color string) overrides the oklch(hue) path so
        // detector bars can pin to the canonical `--color-task-*` tokens.
        bar.style.background = e.color
          ? "color-mix(in oklch, " + e.color + " " + Math.round(alpha * 100) + "%, transparent)"
          : fmtHue(hue, 0.7, 0.16, alpha);
        bar.dataset.idx = idx;
        if (typeof opts.onBarMouseEnter === "function") {
          bar.addEventListener("mouseenter", function () { opts.onBarMouseEnter(idx); });
        }
        if (typeof opts.onBarMouseLeave === "function") {
          bar.addEventListener("mouseleave", function () { opts.onBarMouseLeave(idx); });
        }
        if (typeof opts.onBarClick === "function") {
          bar.addEventListener("click", function (ev) { opts.onBarClick(idx, ev); });
        }
        track.appendChild(bar);
      });
      if (marker != null) {
        var mk = document.createElement("div");
        mk.className = "density-timeline-marker";
        mk.style.left = "calc(" + (marker * 100) + "% - 1px)";
        track.appendChild(mk);
      }
    }

    wrap.update = function (events, marker, durationSec, tickCount) {
      if (durationSec != null || tickCount != null) {
        renderTicks(durationSec != null ? durationSec : opts.durationSec, tickCount != null ? tickCount : opts.tickCount);
      }
      renderBars(events, marker);
    };

    wrap.setHovered = function (idx) {
      var bars = track.querySelectorAll(".density-timeline-bar");
      for (var i = 0; i < bars.length; i++) {
        var b = bars[i];
        if (idx == null || idx === -1) {
          b.classList.remove("is-hover", "is-dim");
        } else if (parseInt(b.dataset.idx, 10) === idx) {
          b.classList.add("is-hover");
          b.classList.remove("is-dim");
        } else {
          b.classList.add("is-dim");
          b.classList.remove("is-hover");
        }
      }
    };

    renderTicks(opts.durationSec, opts.tickCount);
    renderBars(opts.events, opts.marker);
    return wrap;
  }

  // ---- SparkBars ----

  function createSparkBars(opts) {
    opts = opts || {};
    var wrap = document.createElement("div");
    wrap.className = "spark-bars";
    if (opts.height) wrap.style.height = opts.height + "px";
    var data = opts.data || [];
    var hue = resolveHue(null, opts.hue);
    var max = 1;
    data.forEach(function (v) { if (v > max) max = v; });
    data.forEach(function (v) {
      var bar = document.createElement("span");
      var ratio = max > 0 ? (v / max) : 0;
      bar.style.height = Math.max(2, ratio * 100) + "%";
      var alpha = 0.45 + ratio * 0.5;
      bar.style.background = fmtHue(hue, 0.6, 0.14, Math.min(0.95, alpha));
      wrap.appendChild(bar);
    });
    return wrap;
  }

  // ---- SwimLane ----
  //
  // participants: array of participant IDs (strings) shown as lane labels.
  // events: [{ p, t, tEnd?, source?, label, intensity? }, ...] — `p` matches a
  //   participant id; `t` and `tEnd` are normalized 0..1 times; if `tEnd` is
  //   omitted, the marker renders at minimum width. `source` selects a sub-lane
  //   when `sources` length > 1; `label` resolves to a hue via categoryHue().
  // sources: ordered list of source ids (default ['sheet']). Each participant
  //   gets one sub-row per source.
  // clusters: [{ t0, t1, hue, n }, ...] — cluster bands that span the whole timeline.
  // durationSec: numeric duration; tick labels are interpolated to mm:ss / h:mm:ss.

  function fmtSwimTick(sec) {
    var s = Math.max(0, Math.floor(sec));
    var h = Math.floor(s / 3600);
    var m = Math.floor((s - h * 3600) / 60);
    var rs = s - h * 3600 - m * 60;
    var pad = function (n) { return n < 10 ? "0" + n : "" + n; };
    if (h > 0) return h + ":" + pad(m) + ":" + pad(rs);
    return pad(m) + ":" + pad(rs);
  }

  function createSwimLane(opts) {
    opts = opts || {};
    var AXIS_H = 22;
    var TICK_COUNT = opts.tickCount || 11;
    var MIN_MARKER_PX = 6;

    var wrap = document.createElement("div");
    wrap.className = "cg-swim-lane";

    var axis = document.createElement("div");
    axis.className = "cg-swim-axis";
    wrap.appendChild(axis);

    var lanes = document.createElement("div");
    lanes.className = "cg-swim-lanes";
    wrap.appendChild(lanes);

    var labels = document.createElement("div");
    labels.className = "cg-swim-labels cg-mono";
    wrap.appendChild(labels);

    var state = {
      participants: opts.participants || [],
      sources: opts.sources && opts.sources.length ? opts.sources : ["sheet"],
      events: opts.events || [],
      clusters: opts.clusters || [],
      durationSec: opts.durationSec || 0,
      height: opts.height || 320,
      subRowH: opts.subRowH || null,
      eventEls: [],
      clusterEls: [],
    };

    function effectiveSubRowH() {
      if (state.subRowH) return state.subRowH;
      var totalRows = state.participants.length * state.sources.length;
      return totalRows > 0 ? (state.height - AXIS_H) / totalRows : 0;
    }

    function effectiveHeight() {
      if (state.subRowH) {
        return AXIS_H + state.participants.length * state.sources.length * state.subRowH;
      }
      return state.height;
    }

    function applyHeight() {
      var h = effectiveHeight();
      wrap.style.setProperty("--cg-swim-h", h + "px");
      wrap.style.setProperty("--cg-swim-row-h", effectiveSubRowH() + "px");
      lanes.style.height = (h - AXIS_H) + "px";
    }

    function renderAxisTicks() {
      var existing = axis.querySelectorAll(".cg-swim-axis-tick");
      for (var i = 0; i < existing.length; i++) existing[i].remove();
      for (var k = 0; k < TICK_COUNT; k++) {
        var ts = state.durationSec * (k / (TICK_COUNT - 1));
        var tick = document.createElement("span");
        tick.className = "cg-swim-axis-tick cg-mono";
        tick.style.left = (k / (TICK_COUNT - 1) * 100) + "%";
        if (k === 0) tick.style.transform = "translateX(0)";
        else if (k === TICK_COUNT - 1) tick.style.transform = "translateX(-100%)";
        else tick.style.transform = "translateX(-50%)";
        tick.textContent = fmtSwimTick(ts);
        axis.appendChild(tick);
      }
    }

    function renderClusterBands() {
      var existing = wrap.querySelectorAll(".cg-swim-cluster-band, .cg-swim-cluster-connector");
      for (var i = 0; i < existing.length; i++) existing[i].remove();
      state.clusterEls = [];
      var axisFrag = document.createDocumentFragment();
      var laneFrag = document.createDocumentFragment();
      state.clusters.forEach(function (cl, idx) {
        var hue = cl.hue != null ? cl.hue : 220;
        var leftPct = cl.t0 * 100;
        var widthPct = (cl.t1 - cl.t0) * 100;

        var axisBand = document.createElement("div");
        axisBand.className = "cg-swim-cluster-band cg-swim-cluster-axis";
        axisBand.style.left = leftPct + "%";
        axisBand.style.width = widthPct + "%";
        axisBand.style.setProperty("--cg-cluster-hue", hue);
        axisBand.dataset.clusterIdx = idx;
        axisFrag.appendChild(axisBand);

        var laneBand = document.createElement("div");
        laneBand.className = "cg-swim-cluster-band cg-swim-cluster-lanes";
        laneBand.style.left = leftPct + "%";
        laneBand.style.width = widthPct + "%";
        laneBand.style.setProperty("--cg-cluster-hue", hue);
        laneBand.dataset.clusterIdx = idx;
        laneFrag.appendChild(laneBand);

        var connector = document.createElement("div");
        connector.className = "cg-swim-cluster-connector";
        connector.style.left = (((cl.t0 + cl.t1) / 2) * 100) + "%";
        connector.style.setProperty("--cg-cluster-hue", hue);
        laneFrag.appendChild(connector);

        state.clusterEls.push({ axis: axisBand, lane: laneBand, connector: connector });
      });
      axis.appendChild(axisFrag);
      lanes.appendChild(laneFrag);
    }

    function renderLaneSeparators() {
      var existing = lanes.querySelectorAll(".cg-swim-row");
      for (var i = 0; i < existing.length; i++) existing[i].remove();
      var subH = effectiveSubRowH();
      var pCount = state.participants.length;
      var sCount = state.sources.length;
      for (var pi = 0; pi < pCount; pi++) {
        for (var si = 0; si < sCount; si++) {
          var row = document.createElement("div");
          row.className = "cg-swim-row";
          row.style.top = ((pi * sCount + si) * subH) + "px";
          row.style.height = subH + "px";
          row.dataset.participantIdx = pi;
          row.dataset.participant = state.participants[pi];
          row.dataset.sourceIdx = si;
          if (si === sCount - 1) row.classList.add("is-participant-end");
          if (pi === pCount - 1 && si === sCount - 1) row.classList.add("is-last");
          lanes.appendChild(row);
        }
      }
    }

    function renderEvents() {
      var existing = lanes.querySelectorAll(".cg-swim-event");
      for (var i = 0; i < existing.length; i++) existing[i].remove();
      // Sparse, indexed by events[] so getEventsForParticipant / setHovered
      // address by event idx rather than packed render order.
      state.eventEls = new Array(state.events.length);
      var subH = effectiveSubRowH();
      var sCount = state.sources.length;
      var defaultSource = state.sources[0];
      var pIndex = {};
      for (var pi = 0; pi < state.participants.length; pi++) {
        pIndex[state.participants[pi]] = pi;
      }
      var frag = document.createDocumentFragment();
      state.events.forEach(function (e, idx) {
        var pIdx = pIndex[e.p];
        if (pIdx == null) return;
        var sIdx = state.sources.indexOf(e.source || defaultSource);
        if (sIdx < 0) sIdx = 0;
        var hue = resolveHue(e.label, e.hue);
        var laneIdx = pIdx * sCount + sIdx;
        var rowTop = laneIdx * subH;
        var markerH = Math.max(6, Math.min(subH - 2, 12));
        var top = rowTop + (subH - markerH) / 2;
        var leftPct = Math.max(0, Math.min(1, e.t)) * 100;
        var widthPct = 0;
        // Navigational events (boundaries) always render as a thin point tick —
        // orientation scaffolding, not a span — even if the cluster carries one.
        if (!e.navigational && typeof e.tEnd === "number" && e.tEnd > e.t) {
          widthPct = Math.min(1, e.tEnd - e.t) * 100;
        }
        var marker = document.createElement("div");
        marker.className = "cg-swim-event";
        if (e.navigational) marker.classList.add("cg-swim-event--navigational");
        marker.style.left = leftPct + "%";
        marker.style.top = top + "px";
        marker.style.height = markerH + "px";
        if (widthPct > 0) {
          marker.style.width = widthPct + "%";
          marker.style.minWidth = MIN_MARKER_PX + "px";
          marker.style.transform = "none";
          marker.classList.add("has-duration");
        } else {
          marker.style.width = e.navigational ? "2px" : "10px";
          marker.style.transform = "translateX(-50%)";
        }
        marker.style.background = "oklch(0.78 0.14 " + hue + ")";
        marker.style.boxShadow = "0 0 0 1px oklch(0.18 0.04 " + hue + " / 0.55)";
        marker.dataset.idx = idx;
        marker.dataset.participant = e.p;
        marker.dataset.sourceIdx = sIdx;
        frag.appendChild(marker);
        state.eventEls[idx] = marker;
      });
      lanes.appendChild(frag);
    }

    // Hover/click on thousands of markers via three listeners on the wrap, not
    // three per marker. mouseenter doesn't bubble, so hover uses mouseover/out
    // with a relatedTarget guard. Bound once — renderEvents only replaces nodes.
    function bindDelegates() {
      if (typeof opts.onEventHover === "function") {
        lanes.addEventListener("mouseover", function (ev) {
          var marker = ev.target.closest(".cg-swim-event");
          if (!marker || !lanes.contains(marker)) return;
          if (ev.relatedTarget && marker.contains(ev.relatedTarget)) return;
          var idx = parseInt(marker.dataset.idx, 10);
          if (isNaN(idx) || !state.events[idx]) return;
          opts.onEventHover(idx, state.events[idx], true, marker);
        });
        lanes.addEventListener("mouseout", function (ev) {
          var marker = ev.target.closest(".cg-swim-event");
          if (!marker || !lanes.contains(marker)) return;
          if (ev.relatedTarget && marker.contains(ev.relatedTarget)) return;
          var idx = parseInt(marker.dataset.idx, 10);
          if (isNaN(idx) || !state.events[idx]) return;
          opts.onEventHover(idx, state.events[idx], false, marker);
        });
      }
      if (typeof opts.onEventClick === "function" || typeof opts.onClusterClick === "function") {
        wrap.addEventListener("click", function (ev) {
          var marker = ev.target.closest(".cg-swim-event");
          if (marker && wrap.contains(marker) && typeof opts.onEventClick === "function") {
            var midx = parseInt(marker.dataset.idx, 10);
            if (!isNaN(midx) && state.events[midx]) {
              opts.onEventClick(midx, state.events[midx], ev, marker);
            }
            return;
          }
          if (typeof opts.onClusterClick !== "function") return;
          var band = ev.target.closest(".cg-swim-cluster-band");
          if (!band || !wrap.contains(band)) return;
          var cidx = parseInt(band.dataset.clusterIdx, 10);
          if (isNaN(cidx) || !state.clusters[cidx]) return;
          opts.onClusterClick(cidx, state.clusters[cidx], ev);
        });
      }
    }

    function renderLabels() {
      labels.textContent = "";
      var subH = effectiveSubRowH();
      var participantH = subH * state.sources.length;
      state.participants.forEach(function (p, pi) {
        var label = document.createElement("div");
        label.className = "cg-swim-label";
        label.style.height = participantH + "px";
        label.dataset.participant = p;
        label.dataset.participantIdx = pi;
        label.textContent = p;
        labels.appendChild(label);
      });
    }

    function renderAll() {
      applyHeight();
      renderAxisTicks();
      renderClusterBands();
      renderLaneSeparators();
      renderEvents();
      renderLabels();
    }

    wrap.update = function (events, clusters, durationSec, participants, height) {
      if (events) state.events = events;
      if (clusters) state.clusters = clusters;
      if (durationSec != null) state.durationSec = durationSec;
      if (participants) state.participants = participants;
      if (height != null) state.height = height;
      renderAll();
    };

    wrap.setHovered = function (idx) {
      state.eventEls.forEach(function (el, i) {
        if (!el) return;
        if (idx == null || idx === -1) {
          el.classList.remove("is-hover", "is-dim");
        } else if (i === idx) {
          el.classList.add("is-hover");
          el.classList.remove("is-dim");
        } else {
          el.classList.add("is-dim");
          el.classList.remove("is-hover");
        }
      });
    };

    wrap.setSelectedCluster = function (idx) {
      state.clusterEls.forEach(function (cl, i) {
        var on = i === idx;
        cl.axis.classList.toggle("is-selected", on);
        cl.lane.classList.toggle("is-selected", on);
        cl.connector.classList.toggle("is-selected", on);
      });
    };

    wrap.setHoveredCluster = function (idx) {
      state.clusterEls.forEach(function (cl, i) {
        var on = i === idx;
        cl.axis.classList.toggle("is-hovered", on);
        cl.lane.classList.toggle("is-hovered", on);
        cl.connector.classList.toggle("is-hovered", on);
      });
    };

    wrap.getLabelForParticipant = function (pid) {
      return labels.querySelector('.cg-swim-label[data-participant="' + pid + '"]');
    };
    wrap.getRowsForParticipant = function (pid) {
      return lanes.querySelectorAll('.cg-swim-row[data-participant="' + pid + '"]');
    };
    wrap.getEventsForParticipant = function (pid) {
      var out = [];
      for (var i = 0; i < state.events.length; i++) {
        if (state.events[i].p === pid && state.eventEls[i]) out.push(state.eventEls[i]);
      }
      return out;
    };
    wrap.getEventsForParticipantSource = function (pid, sourceIdx) {
      var out = [];
      for (var i = 0; i < state.events.length; i++) {
        if (state.events[i].p === pid &&
            state.sources.indexOf(state.events[i].source) === sourceIdx &&
            state.eventEls[i]) {
          out.push(state.eventEls[i]);
        }
      }
      return out;
    };
    wrap.getLanesPxPerSec = function () {
      var w = lanes.getBoundingClientRect().width;
      var d = state.durationSec || 1;
      return w / d;
    };

    bindDelegates();
    renderAll();
    return wrap;
  }

  // ---- KpiCard ----

  function createKpiCard(opts) {
    opts = opts || {};
    var card = document.createElement("div");
    card.className = "cg-kpi";

    var rail = document.createElement("div");
    rail.className = "cg-kpi-rail";
    rail.style.background = opts.accent || "var(--accent)";
    card.appendChild(rail);

    var labelEl = document.createElement("div");
    labelEl.className = "cg-kpi-label";
    labelEl.textContent = opts.label || "";
    card.appendChild(labelEl);

    var valueRow = document.createElement("div");
    valueRow.className = "cg-kpi-value-row";
    var valueEl = document.createElement("span");
    valueEl.className = "cg-kpi-value cg-mono";
    valueEl.textContent = opts.value != null ? String(opts.value) : "";
    valueRow.appendChild(valueEl);
    if (opts.sub) {
      var subEl = document.createElement("span");
      subEl.className = "cg-kpi-sub cg-mono";
      subEl.textContent = opts.sub;
      valueRow.appendChild(subEl);
    }
    card.appendChild(valueRow);

    if (opts.spark) {
      var sparkWrap = document.createElement("div");
      sparkWrap.className = "cg-kpi-spark";
      sparkWrap.appendChild(opts.spark);
      card.appendChild(sparkWrap);
    }
    setDataset(card, opts.dataset);
    return card;
  }

  // ---- CoverageMatrix ----
  //
  // rows: [{ p, sheet, screenspace, transcript }]; counts ≥ 0.
  // Hues: sheet=280, screenspace=220, transcript=145 (from the prototype).

  function _covCellBg(n, hue, max) {
    if (!n) return "transparent";
    var intensity = Math.min(1, n / max);
    return "oklch(0.32 " + (0.02 + intensity * 0.10).toFixed(3) + " " + hue + " / " + (0.25 + intensity * 0.55).toFixed(3) + ")";
  }
  function _covCellFg(n, hue, max) {
    if (!n) return "var(--fg-dim)";
    var intensity = Math.min(1, n / max);
    return "oklch(" + (0.78 - intensity * 0.05).toFixed(3) + " 0.14 " + hue + ")";
  }

  function createCoverageMatrix(opts) {
    opts = opts || {};
    var table = document.createElement("table");
    table.className = "cg-cov-table";

    var thead = document.createElement("thead");
    var headRow = document.createElement("tr");
    [
      { label: "Participant", align: "left" },
      { label: "Sheet", align: "center" },
      { label: "Screenspace", align: "center" },
      { label: "Transcript", align: "center" },
      { label: "Distribution", align: "center" },
    ].forEach(function (cell) {
      var th = document.createElement("th");
      th.className = "cg-cov-th cg-cov-th-" + cell.align;
      th.textContent = cell.label;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);

    var tbody = document.createElement("tbody");
    table.appendChild(tbody);

    function renderRows(rows, maxCount) {
      tbody.textContent = "";
      var max = maxCount || 1;
      rows.forEach(function (r) {
        var v1 = r.sheet || 0, v2 = r.screenspace || 0, v3 = r.transcript || 0;
        if (!maxCount) {
          if (v1 > max) max = v1;
          if (v2 > max) max = v2;
          if (v3 > max) max = v3;
        }
      });
      rows.forEach(function (r) {
        var tr = document.createElement("tr");
        var pTd = document.createElement("td");
        pTd.className = "cg-cov-td cg-cov-td-left cg-mono";
        pTd.textContent = r.p;
        tr.appendChild(pTd);

        [
          { v: r.sheet || 0, hue: 280 },
          { v: r.screenspace || 0, hue: 220 },
          { v: r.transcript || 0, hue: 145 },
        ].forEach(function (cell) {
          var td = document.createElement("td");
          td.className = "cg-cov-td cg-cov-td-center cg-mono";
          td.style.background = _covCellBg(cell.v, cell.hue, max);
          td.style.color = _covCellFg(cell.v, cell.hue, max);
          td.textContent = String(cell.v);
          tr.appendChild(td);
        });

        var distTd = document.createElement("td");
        distTd.className = "cg-cov-td cg-cov-td-dist";
        var total = (r.sheet || 0) + (r.screenspace || 0) + (r.transcript || 0);
        if (total === 0) {
          var dash = document.createElement("span");
          dash.className = "cg-cov-empty";
          dash.textContent = "—";
          distTd.appendChild(dash);
        } else {
          var bar = document.createElement("div");
          bar.className = "cg-cov-bar";
          [
            { flex: r.sheet || 0, hue: 280 },
            { flex: r.screenspace || 0, hue: 220 },
            { flex: r.transcript || 0, hue: 145 },
          ].forEach(function (seg) {
            if (!seg.flex) return;
            var s = document.createElement("div");
            s.className = "cg-cov-bar-seg";
            s.style.flex = String(seg.flex);
            s.style.background = "oklch(0.65 0.16 " + seg.hue + ")";
            bar.appendChild(s);
          });
          distTd.appendChild(bar);
        }
        tr.appendChild(distTd);
        tbody.appendChild(tr);
      });
    }

    renderRows(opts.rows || [], opts.maxCount);
    table.update = function (rows, maxCount) { renderRows(rows || [], maxCount); };
    return table;
  }

  // ---- Btn ----

  function createBtn(opts) {
    opts = opts || {};
    var btn = document.createElement("button");
    btn.type = "button";
    var classes = ["cg-btn"];
    var variant = opts.variant || "ghost";
    if (variant !== "ghost") classes.push("cg-btn-" + variant);
    var size = opts.size || "md";
    if (size !== "md") classes.push("cg-btn-" + size);
    btn.className = classes.join(" ");

    if (opts.icon) {
      var ic = document.createElement("span");
      ic.className = "cg-btn-icon";
      applyIconMask(ic, opts.icon);
      btn.appendChild(ic);
    }
    if (opts.label) {
      var label = document.createElement("span");
      label.textContent = opts.label;
      btn.appendChild(label);
    }
    if (opts.disabled) btn.disabled = true;
    setDataset(btn, opts.dataset);
    if (typeof opts.onClick === "function") {
      btn.addEventListener("click", opts.onClick);
    }
    return btn;
  }

  // Render a .cg-btn as a left-to-right progress bar. Pairs with the
  // .cg-btn-progress rule + --progress var in primitives.css. Accepts an
  // element or an element id; pass a fraction in [0, 1] to set the fill,
  // or null/-1 to clear it.
  function setButtonProgress(btn, fraction) {
    if (typeof btn === "string") btn = document.getElementById(btn);
    if (!btn) return;
    if (fraction == null || fraction < 0) {
      btn.classList.remove("cg-btn-progress");
      btn.style.removeProperty("--progress");
      return;
    }
    btn.classList.add("cg-btn-progress");
    var clamped = Math.max(0, Math.min(1, fraction));
    btn.style.setProperty("--progress", (clamped * 100).toFixed(1) + "%");
  }

  global.ClipgenPrimitives = {
    createFilterChip: createFilterChip,
    createParticipantPill: createParticipantPill,
    createDensityTimeline: createDensityTimeline,
    createSparkBars: createSparkBars,
    createBtn: createBtn,
    createSwimLane: createSwimLane,
    createKpiCard: createKpiCard,
    createCoverageMatrix: createCoverageMatrix,
    setButtonProgress: setButtonProgress,
  };
})(window);
