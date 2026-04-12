/* Convergence Browser – Phase 1 + Phase 2 */

(function () {
  "use strict";

  var cvState = {
    active: false,
    initialized: false,
    baselines: null,
    events: [],
    filteredEvents: [],
    convergenceZones: [],
    selection: null,
    filters: {
      streams: [],
      eventTypes: [],
      minParticipants: 2,
      windowSec: 5,
      timeRange: null,
    },
    dataVersion: 0,
    duration: 0,
    participants: [],
    _snapshot: null,
  };

  var _summaryHitRects = [];

  // --- Utilities ---

  function debounce(fn, ms) {
    var timer;
    return function () {
      var ctx = this, args = arguments;
      clearTimeout(timer);
      timer = setTimeout(function () { fn.apply(ctx, args); }, ms);
    };
  }

  function getState() {
    return window._studioState;
  }

  function getEventTypeColor(source, eventType) {
    if (source === "screenspace") return DETECTOR_COLORS[eventType] || "#888";
    if (source === "transcript") return (MARK_CATEGORIES[eventType] || {}).color || "#0891b2";
    return XREF_BADGES.sheet.color;
  }

  function stddev(nums) {
    if (nums.length < 2) return 0;
    var sum = 0;
    for (var i = 0; i < nums.length; i++) sum += nums[i];
    var mean = sum / nums.length;
    var sqSum = 0;
    for (var j = 0; j < nums.length; j++) sqSum += (nums[j] - mean) * (nums[j] - mean);
    return Math.sqrt(sqSum / nums.length);
  }

  function parseAccentColor() {
    try {
      var raw = getComputedStyle(document.documentElement).getPropertyValue("--color-accent").trim();
      if (!raw) return { r: 59, g: 130, b: 246 };
      if (raw.charAt(0) === "#") {
        return {
          r: parseInt(raw.slice(1, 3), 16),
          g: parseInt(raw.slice(3, 5), 16),
          b: parseInt(raw.slice(5, 7), 16),
        };
      }
      var m = raw.match(/(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
      if (m) return { r: +m[1], g: +m[2], b: +m[3] };
    } catch (_) { /* fall through */ }
    return { r: 59, g: 130, b: 246 };
  }

  // --- Data Collection ---

  function collectAllEvents() {
    var state = getState();
    var events = [];
    var participantSet = {};

    // Sheet events
    if (state.sheetData && state.sheetData.rows) {
      var participants = state.sheetData.participants || [];
      for (var r = 0; r < state.sheetData.rows.length; r++) {
        var row = state.sheetData.rows[r];
        for (var p = 0; p < participants.length; p++) {
          var pid = participants[p];
          var cell = row.cells[pid];
          if (!cell || !cell.valid) continue;
          var segs = window._studioParseClipTimestamps(cell.value);
          var baselineOffset = (cvState.baselines && cvState.baselines[pid]) || 0;
          for (var s = 0; s < segs.length; s++) {
            var start = segs[s].startSeconds - baselineOffset;
            var end = start + segs[s].duration;
            events.push({
              participant: pid,
              start: Math.max(0, start),
              end: Math.max(0, end),
              source: "sheet",
              eventType: row.category || "uncategorized",
              label: row.observation || "",
              id: "sh_" + row.rowNum + "_" + pid + "_" + s,
              rawData: row,
            });
            participantSet[pid] = true;
          }
        }
      }
    }

    // Screenspace events (raw, not clusters)
    for (var i = 0; i < state.intakeEvents.length; i++) {
      var ev = state.intakeEvents[i];
      events.push({
        participant: ev.participant,
        start: ev.time_in,
        end: ev.time_out,
        source: "screenspace",
        eventType: ev.event_type || ev.detector || "unknown",
        label: (ev.event_type || ev.detector || "") + " detection",
        id: ev.id || ("ss_" + i),
        rawData: ev,
      });
      participantSet[ev.participant] = true;
    }

    // Transcript marks
    for (var j = 0; j < state.trIntakeMarks.length; j++) {
      var mark = state.trIntakeMarks[j];
      events.push({
        participant: mark.participant,
        start: mark.start,
        end: mark.end,
        source: "transcript",
        eventType: mark.category || "bookmark",
        label: mark.text || mark.label || "",
        id: mark.id || mark.segment_id || ("tr_" + j),
        rawData: mark,
      });
      participantSet[mark.participant] = true;
    }

    // Duration
    var maxEnd = 0;
    for (var k = 0; k < events.length; k++) {
      if (events[k].end > maxEnd) maxEnd = events[k].end;
    }
    cvState.duration = Math.max(maxEnd * 1.05, 60);

    // Participants: spreadsheet order first, then others alphabetically
    var ordered = [];
    if (state.sheetData && state.sheetData.participants) {
      for (var pi = 0; pi < state.sheetData.participants.length; pi++) {
        ordered.push(state.sheetData.participants[pi]);
      }
    }
    var others = [];
    var pids = Object.keys(participantSet);
    for (var qi = 0; qi < pids.length; qi++) {
      if (ordered.indexOf(pids[qi]) < 0) others.push(pids[qi]);
    }
    others.sort();
    cvState.participants = ordered.concat(others);

    cvState.events = events;

    // Snapshot for staleness detection
    cvState._snapshot = {
      ss: state.intakeEvents.length,
      tr: state.trIntakeMarks.length,
      sh: state.sheetData ? state.sheetData.rows.length : 0,
    };
  }

  // --- Convergence Algorithm ---

  function computeConvergenceZones(events, windowSec, minParticipants) {
    if (!events.length) return [];

    // Sort by start time
    var sorted = events.slice().sort(function (a, b) { return a.start - b.start; });

    // For each event, find distinct participants within ±windowSec
    var qualifying = []; // indices of events that meet threshold
    for (var i = 0; i < sorted.length; i++) {
      var center = sorted[i].start;
      var seen = {};
      for (var j = 0; j < sorted.length; j++) {
        if (sorted[j].start > center + windowSec) break;
        if (sorted[j].end < center - windowSec && sorted[j].start < center - windowSec) continue;
        // Event j overlaps the window [center - W, center + W]
        if (sorted[j].start <= center + windowSec && sorted[j].end >= center - windowSec) {
          seen[sorted[j].participant] = true;
        }
      }
      if (Object.keys(seen).length >= minParticipants) {
        qualifying.push(i);
      }
    }

    if (!qualifying.length) return [];

    // Build zones from consecutive qualifying events
    var zones = [];
    var zStart = sorted[qualifying[0]].start;
    var zEnd = sorted[qualifying[0]].end;
    var zEvents = [sorted[qualifying[0]]];

    for (var q = 1; q < qualifying.length; q++) {
      var evt = sorted[qualifying[q]];
      // Merge if overlapping or within window
      if (evt.start <= zEnd + windowSec) {
        zEnd = Math.max(zEnd, evt.end);
        zEvents.push(evt);
      } else {
        zones.push(buildZone(zStart, zEnd, zEvents, windowSec));
        zStart = evt.start;
        zEnd = evt.end;
        zEvents = [evt];
      }
    }
    zones.push(buildZone(zStart, zEnd, zEvents, windowSec));

    return zones;
  }

  function buildZone(start, end, events, windowSec) {
    var pSet = {};
    var starts = [];
    for (var i = 0; i < events.length; i++) {
      pSet[events[i].participant] = true;
      starts.push(events[i].start);
    }
    var participants = Object.keys(pSet);
    var tightness = stddev(starts);
    var totalP = cvState.participants.length || 1;
    var strength = (participants.length / totalP) * (1 / (1 + tightness / Math.max(windowSec, 1)));
    return {
      start: start,
      end: end,
      participantCount: participants.length,
      participants: participants,
      events: events,
      tightness: tightness,
      strength: strength,
    };
  }

  // --- Filter Pipeline ---

  function applyFilters() {
    var filtered = cvState.events;
    var streams = cvState.filters.streams;
    var types = cvState.filters.eventTypes;
    var range = cvState.filters.timeRange;

    if (streams.length > 0) {
      filtered = filtered.filter(function (e) { return streams.indexOf(e.source) >= 0; });
    }
    if (types.length > 0) {
      filtered = filtered.filter(function (e) { return types.indexOf(e.eventType) >= 0; });
    }
    if (range) {
      filtered = filtered.filter(function (e) { return e.end >= range.start && e.start <= range.end; });
    }

    cvState.filteredEvents = filtered;
    cvState.convergenceZones = computeConvergenceZones(
      filtered,
      cvState.filters.windowSec,
      cvState.filters.minParticipants
    );
  }

  function recalculate() {
    collectAllEvents();
    populateEventTypePills();
    applyFilters();
    render();
  }

  // --- Event Type Pills ---

  function populateEventTypePills() {
    var container = qs("#cvEventTypePills");
    if (!container) return;

    var streams = cvState.filters.streams;
    var typeMap = {}; // source -> Set<eventType>
    for (var i = 0; i < cvState.events.length; i++) {
      var e = cvState.events[i];
      if (streams.length > 0 && streams.indexOf(e.source) < 0) continue;
      if (!typeMap[e.source]) typeMap[e.source] = {};
      typeMap[e.source][e.eventType] = true;
    }

    var frag = document.createDocumentFragment();
    var sources = ["sheet", "screenspace", "transcript"];
    for (var si = 0; si < sources.length; si++) {
      var src = sources[si];
      var types = typeMap[src];
      if (!types) continue;
      var keys = Object.keys(types).sort();
      for (var ki = 0; ki < keys.length; ki++) {
        var type = keys[ki];
        var pill = document.createElement("button");
        pill.className = "cv-event-type-pill" +
          (cvState.filters.eventTypes.indexOf(type) >= 0 ? " active" : "");
        pill.textContent = type;
        pill.dataset.eventType = type;
        pill.style.setProperty("--det-color", getEventTypeColor(src, type));
        pill.addEventListener("click", (function (t) {
          return function () { toggleEventType(t); };
        })(type));
        frag.appendChild(pill);
      }
    }

    container.innerHTML = "";
    container.appendChild(frag);
  }

  function toggleEventType(type) {
    var idx = cvState.filters.eventTypes.indexOf(type);
    if (idx >= 0) {
      cvState.filters.eventTypes.splice(idx, 1);
    } else {
      cvState.filters.eventTypes.push(type);
    }
    // Update pill active states
    var pills = qsa(".cv-event-type-pill");
    for (var i = 0; i < pills.length; i++) {
      pills[i].classList.toggle("active", cvState.filters.eventTypes.indexOf(pills[i].dataset.eventType) >= 0);
    }
    onFilterChange();
  }

  // --- Filter Controls ---

  var STREAM_DEFS = [
    { key: "all", label: "All streams", color: "var(--color-accent)" },
    { key: "sheet", label: "Sheet", color: XREF_BADGES.sheet.color },
    { key: "screenspace", label: "Screenspace", color: XREF_BADGES.screenspace.color },
    { key: "transcript", label: "Transcript", color: XREF_BADGES.transcript.color },
  ];

  function buildFilterControls() {
    var controls = qs("#convergenceControls");
    var filters = qs("#convergenceFilters");
    if (!controls || !filters) return;

    // --- Controls bar: min participants + window ---
    var minLabel = el("label", "intake-cluster-label");
    minLabel.textContent = "Min participants ";
    var minInput = document.createElement("input");
    minInput.type = "number";
    minInput.min = "2";
    minInput.value = String(cvState.filters.minParticipants);
    minInput.className = "intake-cluster-input";
    minInput.autocomplete = "off";
    minInput.addEventListener("input", debouncedFilterChange);
    minLabel.appendChild(minInput);

    var winLabel = el("label", "intake-cluster-label");
    winLabel.textContent = "Window ";
    var winInput = document.createElement("input");
    winInput.type = "number";
    winInput.min = "1";
    winInput.max = "60";
    winInput.value = String(cvState.filters.windowSec);
    winInput.className = "intake-cluster-input";
    winInput.autocomplete = "off";
    winInput.addEventListener("input", debouncedFilterChange);
    var winSuffix = document.createTextNode("\u00B1s");
    winLabel.appendChild(winInput);
    winLabel.appendChild(winSuffix);

    controls.appendChild(minLabel);
    controls.appendChild(winLabel);

    // --- Filters bar: stream toggles + event type pills ---
    for (var i = 0; i < STREAM_DEFS.length; i++) {
      var def = STREAM_DEFS[i];
      var btn = document.createElement("button");
      btn.className = "cv-stream-toggle" + (def.key === "all" ? " active" : "");
      btn.textContent = def.label;
      btn.dataset.stream = def.key;
      btn.style.setProperty("--det-color", def.color);
      btn.addEventListener("click", (function (key) {
        return function () { onStreamToggle(key); };
      })(def.key));
      filters.appendChild(btn);
    }

    var typePills = document.createElement("div");
    typePills.id = "cvEventTypePills";
    filters.appendChild(typePills);
  }

  function onStreamToggle(stream) {
    if (stream === "all") {
      cvState.filters.streams = [];
    } else {
      var idx = cvState.filters.streams.indexOf(stream);
      if (idx >= 0) {
        cvState.filters.streams.splice(idx, 1);
      } else {
        cvState.filters.streams.push(stream);
      }
    }

    // Update pill active states
    var pills = qsa(".cv-stream-toggle");
    var isAll = cvState.filters.streams.length === 0;
    for (var i = 0; i < pills.length; i++) {
      var key = pills[i].dataset.stream;
      if (key === "all") {
        pills[i].classList.toggle("active", isAll);
      } else {
        pills[i].classList.toggle("active", cvState.filters.streams.indexOf(key) >= 0);
      }
    }

    // Clear event type filter when streams change
    cvState.filters.eventTypes = [];
    populateEventTypePills();
    onFilterChange();
  }

  function syncFilterInputs() {
    var controls = qs("#convergenceControls");
    if (!controls) return;
    var inputs = controls.querySelectorAll("input[type=number]");
    if (inputs[0]) cvState.filters.minParticipants = Math.max(2, parseInt(inputs[0].value, 10) || 2);
    if (inputs[1]) cvState.filters.windowSec = Math.max(1, parseInt(inputs[1].value, 10) || 5);
  }

  function onFilterChange() {
    syncFilterInputs();
    applyFilters();
    render();
  }

  var debouncedFilterChange = debounce(onFilterChange, 250);

  // --- Rendering ---

  function render() {
    renderTimeline();
    requestAnimationFrame(renderCanvases);
  }

  function renderTimeline() {
    var container = qs("#convergenceTimeline");
    if (!container) return;

    container.innerHTML = "";

    // Empty state
    if (cvState.events.length === 0) {
      var empty = el("div", "cv-empty-state");
      empty.textContent = "No events loaded. Load data from multiple participants to see convergence.";
      container.appendChild(empty);
      return;
    }

    if (cvState.filteredEvents.length === 0) {
      var emptyFiltered = el("div", "cv-empty-state");
      emptyFiltered.textContent = "No events match the current filters.";
      container.appendChild(emptyFiltered);
      return;
    }

    var frag = document.createDocumentFragment();

    // Time axis
    var axis = el("div", "cv-time-axis");
    var tickInterval = window._studioIntakeComputeTickInterval(cvState.duration);
    for (var t = tickInterval; t <= cvState.duration; t += tickInterval) {
      var tick = el("div", "cv-tick");
      tick.style.left = (t / cvState.duration * 100) + "%";
      var tickLabel = el("span", "cv-tick-label");
      tickLabel.textContent = window._studioFormatDuration(t);
      tick.appendChild(tickLabel);
      axis.appendChild(tick);
    }
    frag.appendChild(axis);

    // Summary lane
    var summaryWrap = el("div", "cv-summary-lane-wrap");
    var summaryCanvas = document.createElement("canvas");
    summaryCanvas.className = "cv-summary-canvas";
    summaryWrap.appendChild(summaryCanvas);
    frag.appendChild(summaryWrap);

    // Staleness banner placeholder
    var staleBanner = el("div", "cv-stale-banner hidden");
    staleBanner.id = "cvStaleBanner";
    var staleText = el("span", "");
    staleText.textContent = "New data available";
    var staleBtn = document.createElement("button");
    staleBtn.className = "cv-stale-refresh";
    staleBtn.textContent = "Refresh";
    staleBtn.addEventListener("click", function () { recalculate(); });
    staleBanner.appendChild(staleText);
    staleBanner.appendChild(staleBtn);
    frag.appendChild(staleBanner);

    // No convergence message
    if (cvState.convergenceZones.length === 0 && cvState.filteredEvents.length > 0) {
      var noZones = el("div", "cv-no-convergence");
      noZones.textContent = "No convergence detected. Try widening the window or lowering the threshold.";
      frag.appendChild(noZones);
    }

    // Participant rows
    var rowsContainer = el("div", "cv-participant-rows");

    // Index filtered events by participant and source for fast lookup
    var eventIndex = {}; // pid -> source -> events[]
    for (var ei = 0; ei < cvState.filteredEvents.length; ei++) {
      var fe = cvState.filteredEvents[ei];
      if (!eventIndex[fe.participant]) eventIndex[fe.participant] = {};
      if (!eventIndex[fe.participant][fe.source]) eventIndex[fe.participant][fe.source] = [];
      eventIndex[fe.participant][fe.source].push(fe);
    }

    for (var pi = 0; pi < cvState.participants.length; pi++) {
      var pid = cvState.participants[pi];
      var row = el("div", "cv-participant-row");

      var label = el("div", "cv-participant-label");
      label.textContent = pid;
      row.appendChild(label);

      var tracks = el("div", "cv-tracks-container");

      // Row shading canvas (behind markers)
      var rowCanvas = document.createElement("canvas");
      rowCanvas.className = "cv-row-shading";
      rowCanvas.dataset.participant = pid;
      tracks.appendChild(rowCanvas);

      // Sub-tracks
      var sources = ["sheet", "screenspace", "transcript"];
      for (var si = 0; si < sources.length; si++) {
        var src = sources[si];
        var subTrack = el("div", "cv-sub-track");
        subTrack.dataset.source = src;

        // Add markers for this participant + source
        var pEvents = (eventIndex[pid] && eventIndex[pid][src]) || [];
        for (var mi = 0; mi < pEvents.length; mi++) {
          var mev = pEvents[mi];
          var marker = el("div", "cv-event-marker");
          marker.style.left = (mev.start / cvState.duration * 100) + "%";
          marker.style.width = Math.max((mev.end - mev.start) / cvState.duration * 100, 0.3) + "%";
          marker.style.background = getEventTypeColor(mev.source, mev.eventType);
          marker.title = mev.eventType + " (" + formatTime(mev.start) + ")";
          subTrack.appendChild(marker);
        }

        tracks.appendChild(subTrack);
      }

      row.appendChild(tracks);
      rowsContainer.appendChild(row);
    }

    frag.appendChild(rowsContainer);
    container.appendChild(frag);
  }

  function renderCanvases() {
    renderSummaryLane();
    renderAllRowShading();
  }

  // --- Summary Lane Canvas ---

  function renderSummaryLane() {
    var canvas = qs(".cv-summary-canvas");
    if (!canvas) return;
    var w = canvas.clientWidth;
    if (w <= 0) return;

    var dpr = window.devicePixelRatio || 1;
    canvas.width = w * dpr;
    canvas.height = 32 * dpr;
    var ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

    var cs = getComputedStyle(document.documentElement);
    var surfaceAlt = cs.getPropertyValue("--color-surface-alt").trim() || "#f1ece4";
    ctx.fillStyle = surfaceAlt;
    ctx.fillRect(0, 0, w, 32);

    _summaryHitRects = [];
    var zones = cvState.convergenceZones;
    if (!zones.length) return;

    // Normalize strength for alpha mapping
    var maxStrength = 0;
    for (var i = 0; i < zones.length; i++) {
      if (zones[i].strength > maxStrength) maxStrength = zones[i].strength;
    }
    if (maxStrength === 0) maxStrength = 1;

    var accent = parseAccentColor();

    for (var zi = 0; zi < zones.length; zi++) {
      var zone = zones[zi];
      var x1 = (zone.start / cvState.duration) * w;
      var x2 = (zone.end / cvState.duration) * w;
      var zw = Math.max(x2 - x1, 2);
      var alpha = 0.15 + (zone.strength / maxStrength) * 0.65;

      ctx.fillStyle = "rgba(" + accent.r + "," + accent.g + "," + accent.b + "," + alpha + ")";
      ctx.fillRect(x1, 0, zw, 32);

      _summaryHitRects.push({ x1: x1, x2: x1 + zw, y: 0, h: 32, zoneIdx: zi });
    }
  }

  // --- Per-Participant Row Shading ---

  function renderAllRowShading() {
    var canvases = qsa(".cv-row-shading");
    for (var i = 0; i < canvases.length; i++) {
      renderRowShading(canvases[i], canvases[i].dataset.participant);
    }
  }

  function renderRowShading(canvas, participantId) {
    var w = canvas.clientWidth;
    var h = canvas.clientHeight;
    if (w <= 0 || h <= 0) return;

    var dpr = window.devicePixelRatio || 1;
    canvas.width = w * dpr;
    canvas.height = h * dpr;
    var ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

    var zones = cvState.convergenceZones;
    if (!zones.length) return;

    var accent = parseAccentColor();

    for (var zi = 0; zi < zones.length; zi++) {
      var zone = zones[zi];
      // Check if this participant contributed
      if (zone.participants.indexOf(participantId) < 0) continue;

      // Count this participant's events in the zone
      var pCount = 0;
      for (var ei = 0; ei < zone.events.length; ei++) {
        if (zone.events[ei].participant === participantId) pCount++;
      }
      var ratio = pCount / Math.max(zone.events.length, 1);
      var alpha = Math.min(ratio * 0.3, 0.15);

      var x1 = (zone.start / cvState.duration) * w;
      var x2 = (zone.end / cvState.duration) * w;
      var zw = Math.max(x2 - x1, 2);

      ctx.fillStyle = "rgba(" + accent.r + "," + accent.g + "," + accent.b + "," + alpha + ")";
      ctx.fillRect(x1, 0, zw, h);
    }
  }

  // --- Data Freshness ---

  function checkStaleness() {
    if (!cvState._snapshot || !cvState.active) return;
    var state = getState();
    var stale = (state.intakeEvents.length !== cvState._snapshot.ss) ||
      (state.trIntakeMarks.length !== cvState._snapshot.tr) ||
      ((state.sheetData ? state.sheetData.rows.length : 0) !== cvState._snapshot.sh);

    var banner = qs("#cvStaleBanner");
    if (banner) {
      if (stale) {
        banner.classList.remove("hidden");
      } else {
        banner.classList.add("hidden");
      }
    }
  }

  // --- Lifecycle ---

  function activate() {
    cvState.active = true;
    if (!cvState.initialized) {
      buildFilterControls();
      cvState.initialized = true;
    }

    if (cvState.baselines === null) {
      // First activation: fetch baselines then recalculate
      apiGet("api/sheet/baseline").then(function (data) {
        var parsed = {};
        if (data.ok && data.baselines) {
          var keys = Object.keys(data.baselines);
          for (var i = 0; i < keys.length; i++) {
            var val = data.baselines[keys[i]];
            if (val) {
              var sec = window._studioParseTimestampToSeconds(val);
              parsed[keys[i]] = isNaN(sec) ? 0 : sec;
            }
          }
        }
        cvState.baselines = parsed;
        recalculate();
      }).catch(function () {
        cvState.baselines = {};
        recalculate();
      });
    } else {
      checkStaleness();
    }
  }

  function deactivate() {
    cvState.active = false;
  }

  function init() {
    document.addEventListener("visibilitychange", function () {
      if (!document.hidden && cvState.active) {
        checkStaleness();
      }
    });
  }

  // --- Window exports ---
  window.convergenceActivate = activate;
  window.convergenceDeactivate = deactivate;
  window.convergenceInit = init;
  window.convergenceResize = debounce(function () {
    if (!cvState.active) return;
    renderCanvases();
  }, 200);

  init();
})();
