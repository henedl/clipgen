/* Convergence Browser.
 *
 * Pulls events from multiple participants/streams (Screenspace detector
 * events + Transcripts marks + sheet timestamps) onto a single timeline,
 * then clusters moments where many participants do the same thing within a
 * short window into "convergence zones". Renders a marker timeline and a
 * frame-preview panel.
 *
 * `cvState._snapshot` records the last (ss, tr, sh) input lengths the view
 * was built against, so we can detect when upstream data changed and the
 * view needs to be rebuilt.
 */

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
      clusterSec: 5,
      timeRange: null,
    },
    dataVersion: 0,
    duration: 0,
    participants: [],
    hoveredMarkerIdx: -1,
    sortByDensity: false,
    _snapshot: null,
  };

  var _summaryHitRects = [];
  var _cvFrameCache = {};
  var _cvFramePreviewEl = null;
  var _cvHoverDebounce = null;
  var _cvZoneTooltipEl = null;

  // --- Utilities ---

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

  // --- Frame Preview ---

  function cvFrameUrl(source, participant, startSec) {
    if (source === "screenspace") {
      return "../screenspace/api/video/frame/" + encodeURIComponent(participant)
        + "/" + Math.floor(startSec) + "?w=240";
    }
    return "api/thumbnail/" + encodeURIComponent(participant) + "/" + Math.floor(startSec);
  }

  function cvEnsureFramePreview() {
    if (_cvFramePreviewEl) return _cvFramePreviewEl;
    var wrap = el("div", "cv-frame-preview hidden");
    var img = document.createElement("img");
    img.className = "cv-frame-preview-img";
    img.alt = "";
    wrap.appendChild(img);
    var lbl = el("div", "cv-frame-preview-label");
    wrap.appendChild(lbl);
    document.body.appendChild(wrap);
    _cvFramePreviewEl = wrap;
    return wrap;
  }

  function cvShowFramePreview(markerEl, event) {
    var preview = cvEnsureFramePreview();
    var img = preview.querySelector("img");
    var lbl = preview.querySelector(".cv-frame-preview-label");
    var url = cvFrameUrl(event.source, event.participant, event.start);

    var cached = _cvFrameCache[url];
    if (cached && cached !== "error" && cached !== "loading") {
      img.src = cached;
    } else if (cached === "error") {
      cvHideFramePreview();
      return;
    } else if (!cached) {
      _cvFrameCache[url] = "loading";
      img.src = "";
      // TODO: utils.apiGet returns JSON; needs an apiGetBlob helper to migrate.
      fetch(url)
        .then(function (r) {
          if (!r.ok) throw new Error("status " + r.status);
          return r.blob();
        })
        .then(function (blob) {
          var objUrl = URL.createObjectURL(blob);
          _cvFrameCache[url] = objUrl;
          // Only update if still showing this preview
          if (!preview.classList.contains("hidden") && img.parentNode) {
            img.src = objUrl;
          }
        })
        .catch(function () {
          _cvFrameCache[url] = "error";
          cvHideFramePreview();
        });
    }

    lbl.textContent = event.participant + " \u00b7 " + formatTime(event.start);
    preview.classList.remove("hidden");
    positionTooltipAnchored(preview, markerEl.getBoundingClientRect());
  }

  function cvHideFramePreview() {
    clearTimeout(_cvHoverDebounce);
    _cvHoverDebounce = null;
    if (_cvFramePreviewEl) _cvFramePreviewEl.classList.add("hidden");
  }

  // --- Zone Tooltip ---

  function cvEnsureZoneTooltip() {
    if (_cvZoneTooltipEl) return _cvZoneTooltipEl;
    var tip = el("div", "cv-zone-tooltip hidden");
    document.body.appendChild(tip);
    _cvZoneTooltipEl = tip;
    return tip;
  }

  // --- Density Sort ---

  function cvComputeDensityOrder() {
    var counts = {};
    for (var i = 0; i < cvState.participants.length; i++) {
      counts[cvState.participants[i]] = 0;
    }
    for (var zi = 0; zi < cvState.convergenceZones.length; zi++) {
      var zone = cvState.convergenceZones[zi];
      for (var ei = 0; ei < zone.events.length; ei++) {
        var pid = zone.events[ei].participant;
        if (counts[pid] !== undefined) counts[pid]++;
      }
    }
    var ordered = cvState.participants.slice();
    ordered.sort(function (a, b) { return (counts[b] || 0) - (counts[a] || 0); });
    return ordered;
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
          var baselineOffset = (cvState.baselines && cvState.baselines[pid]) || 0;
          var segs = parseClipSegmentsForCell(cell.value, baselineOffset, CLIPGEN_CONFIG.defaultDuration);
          for (var s = 0; s < segs.length; s++) {
            var start = segs[s].startSeconds;
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

    // Screenspace events (clustered)
    var clusterSec = cvState.filters.clusterSec;
    var ssClusters = window._studioClusterIntakeEvents
      ? window._studioClusterIntakeEvents(state.intakeEvents, clusterSec)
      : [];
    for (var i = 0; i < ssClusters.length; i++) {
      var cl = ssClusters[i];
      var clCount = cl.events ? cl.events.length : 1;
      events.push({
        participant: cl.participant,
        start: cl.start,
        end: cl.end,
        source: "screenspace",
        eventType: cl.event_type || cl.detector || "unknown",
        label: (cl.event_type || cl.detector || "") + " detection"
          + (clCount > 1 ? " (" + clCount + " events)" : ""),
        id: "ss_cl_" + i,
        rawData: cl,
        clusterCount: clCount,
      });
      participantSet[cl.participant] = true;
    }

    // Transcript marks (clustered)
    var trClusters = window._studioClusterTranscriptMarks
      ? window._studioClusterTranscriptMarks(state.trIntakeMarks, clusterSec)
      : [];
    for (var j = 0; j < trClusters.length; j++) {
      var tc = trClusters[j];
      var tcCount = tc.marks ? tc.marks.length : 1;
      events.push({
        participant: tc.participant,
        start: tc.start,
        end: tc.end,
        source: "transcript",
        eventType: tc.category || "bookmark",
        label: (tc.label || tc.text || "")
          + (tcCount > 1 ? " (" + tcCount + " marks)" : ""),
        id: "tr_cl_" + j,
        rawData: tc,
        clusterCount: tcCount,
      });
      participantSet[tc.participant] = true;
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
  //
  // Two passes:
  //   1. For each event, count distinct participants whose events overlap
  //      the window [start - W, start + W]. Events meeting `minParticipants`
  //      are "qualifying".
  //   2. Walk qualifying events in time order and merge any two within W of
  //      each other into a single zone. This collapses contiguous bursts of
  //      activity into one entry instead of one-zone-per-event.
  //
  // The inner loop in pass 1 is O(n²) in the worst case but breaks early on
  // the sorted array once we pass the right edge of the window.

  function computeConvergenceZones(events, windowSec, minParticipants) {
    if (!events.length) return [];

    var sorted = events.slice().sort(function (a, b) { return a.start - b.start; });

    // Pass 1: per-event distinct-participant count within ±windowSec
    var qualifying = []; // indices into sorted[] that meet threshold
    for (var i = 0; i < sorted.length; i++) {
      var center = sorted[i].start;
      var seen = {};
      for (var j = 0; j < sorted.length; j++) {
        if (sorted[j].start > center + windowSec) break;
        // Skip if event j ends before the window opens.
        if (sorted[j].end < center - windowSec && sorted[j].start < center - windowSec) continue;
        // Event j overlaps the window [center - W, center + W].
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
    clearSelection();
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

    var clusterLabel = el("label", "intake-cluster-label");
    clusterLabel.textContent = "Cluster ";
    var clusterInput = document.createElement("input");
    clusterInput.type = "number";
    clusterInput.id = "cvClusterThreshold";
    clusterInput.min = "1";
    clusterInput.max = "60";
    clusterInput.value = String(cvState.filters.clusterSec);
    clusterInput.className = "intake-cluster-input";
    clusterInput.autocomplete = "off";
    clusterInput.addEventListener("input", debouncedRecalculate);
    var clusterSuffix = document.createTextNode("s");
    clusterLabel.appendChild(clusterInput);
    clusterLabel.appendChild(clusterSuffix);

    var sortBtn = document.createElement("button");
    sortBtn.className = "cv-sort-toggle";
    sortBtn.id = "cvSortToggle";
    sortBtn.textContent = "Sort by density";
    sortBtn.addEventListener("click", function () {
      cvState.sortByDensity = !cvState.sortByDensity;
      sortBtn.classList.toggle("active", cvState.sortByDensity);
      render();
    });

    controls.appendChild(minLabel);
    controls.appendChild(winLabel);
    controls.appendChild(clusterLabel);
    controls.appendChild(sortBtn);

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
    if (inputs[2]) cvState.filters.clusterSec = Math.max(1, parseInt(inputs[2].value, 10) || 5);
  }

  var debouncedRecalculate = debounce(function () {
    syncFilterInputs();
    recalculate();
  }, 250);

  function onFilterChange() {
    syncFilterInputs();
    applyFilters();
    // Preserve or clear selection after filter change
    if (cvState.selection) {
      if (cvState.selection.zone) {
        // Zone-based selection: clear if zone no longer exists
        clearSelection();
      } else {
        // Drag-based: re-collect events in the same time range
        var sel = cvState.selection;
        var events = [];
        for (var i = 0; i < cvState.filteredEvents.length; i++) {
          var ev = cvState.filteredEvents[i];
          if (ev.start < sel.end && ev.end > sel.start) events.push(ev);
        }
        if (events.length === 0) {
          clearSelection();
        } else {
          sel.events = events;
        }
      }
    }
    render();
  }

  var debouncedFilterChange = debounce(onFilterChange, 250);

  // --- Tick Interval (pixel-aware) ---

  function computeConvergenceTickInterval(duration, trackWidthPx) {
    // Each tick label is centered on its mark. To prevent overlap the
    // minimum distance between ticks must be at least one full label
    // width plus comfortable padding.  10px monospace "H:MM:SS" ≈ 50px,
    // "M:SS" ≈ 30px; add 30px padding so labels breathe.
    var slotWidth = duration >= 3600 ? 80 : 60;
    var maxTicks = Math.max(2, Math.floor(trackWidthPx / slotWidth));
    var candidates = [1, 2, 5, 10, 15, 30, 60, 120, 300, 600, 900, 1800, 3600,
      7200, 10800, 21600, 43200];
    for (var i = 0; i < candidates.length; i++) {
      if (duration / candidates[i] <= maxTicks) return candidates[i];
    }
    return 43200;
  }

  // --- Rendering ---

  function render() {
    renderTimeline();
    requestAnimationFrame(function () {
      renderCanvases();
      if (cvState.selection) {
        renderSelectionOverlay();
        renderDetailPanel();
      }
    });
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

    // Time axis (flex row: spacer + tick area, matching participant row layout)
    var axis = el("div", "cv-time-axis");
    var axisSpacer = el("div", "cv-axis-spacer");
    axis.appendChild(axisSpacer);
    var axisTrack = el("div", "cv-axis-track");
    var trackWidthPx = (container.clientWidth || 500) - 52;
    var tickInterval = computeConvergenceTickInterval(cvState.duration, trackWidthPx);
    for (var t = tickInterval; t <= cvState.duration; t += tickInterval) {
      var tick = el("div", "cv-tick");
      tick.style.left = (t / cvState.duration * 100) + "%";
      var tickLabel = el("span", "cv-tick-label");
      tickLabel.textContent = window._studioFormatDuration(t);
      tick.appendChild(tickLabel);
      axisTrack.appendChild(tick);
    }
    axis.appendChild(axisTrack);
    frag.appendChild(axis);

    // Summary lane (flex row: spacer + canvas, matching participant row layout)
    var summaryWrap = el("div", "cv-summary-lane-wrap");
    var summarySpacer = el("div", "cv-axis-spacer");
    summaryWrap.appendChild(summarySpacer);
    var summaryTrack = el("div", "cv-summary-track");
    var summaryCanvas = document.createElement("canvas");
    summaryCanvas.className = "cv-summary-canvas";
    summaryTrack.appendChild(summaryCanvas);
    summaryWrap.appendChild(summaryTrack);
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
    var filteredIdxById = {}; // event.id -> index in filteredEvents
    for (var ei = 0; ei < cvState.filteredEvents.length; ei++) {
      var fe = cvState.filteredEvents[ei];
      if (!eventIndex[fe.participant]) eventIndex[fe.participant] = {};
      if (!eventIndex[fe.participant][fe.source]) eventIndex[fe.participant][fe.source] = [];
      eventIndex[fe.participant][fe.source].push(fe);
      filteredIdxById[fe.id] = ei;
    }

    var displayParticipants = cvState.sortByDensity
      ? cvComputeDensityOrder() : cvState.participants;

    for (var pi = 0; pi < displayParticipants.length; pi++) {
      var pid = displayParticipants[pi];
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
          var tooltipText = mev.eventType + " (" + formatTime(mev.start) + ")";
          if (mev.clusterCount && mev.clusterCount > 1) {
            tooltipText += " [" + mev.clusterCount + " events]";
          }
          marker.title = tooltipText;
          if (filteredIdxById[mev.id] !== undefined) marker.dataset.cvIdx = filteredIdxById[mev.id];
          subTrack.appendChild(marker);
        }

        tracks.appendChild(subTrack);
      }

      row.appendChild(tracks);
      rowsContainer.appendChild(row);
    }

    frag.appendChild(rowsContainer);
    container.appendChild(frag);

    // Attach interaction handlers
    summaryCanvas = container.querySelector(".cv-summary-canvas");
    if (summaryCanvas) {
      summaryCanvas.addEventListener("click", handleSummaryClick);

      // Zone tooltip on summary lane hover
      summaryCanvas.addEventListener("mousemove", function (e) {
        var rect = summaryCanvas.getBoundingClientRect();
        var mx = e.clientX - rect.left;
        var hit = null;
        for (var i = 0; i < _summaryHitRects.length; i++) {
          if (mx >= _summaryHitRects[i].x1 && mx <= _summaryHitRects[i].x2) {
            hit = _summaryHitRects[i]; break;
          }
        }
        var tip = cvEnsureZoneTooltip();
        if (hit) {
          var zone = cvState.convergenceZones[hit.zoneIdx];
          if (zone) {
            tip.textContent = zone.participantCount + " participant"
              + (zone.participantCount !== 1 ? "s" : "")
              + " \u00b7 " + formatTime(zone.start) + "\u2013" + formatTime(zone.end)
              + " \u00b7 " + zone.events.length + " event"
              + (zone.events.length !== 1 ? "s" : "");
            tip.classList.remove("hidden");
            positionTooltipAnchored(tip, {
              left: e.clientX - 4, top: rect.top,
              width: 8, height: rect.height, bottom: rect.bottom,
            });
          }
        } else {
          tip.classList.add("hidden");
        }
      });
      summaryCanvas.addEventListener("mouseleave", function () {
        cvEnsureZoneTooltip().classList.add("hidden");
      });
    }

    // Marker hover: frame preview + dimming
    rowsContainer.addEventListener("mouseover", function (e) {
      var marker = e.target.closest(".cv-event-marker");
      if (!marker) return;
      var idx = parseInt(marker.dataset.cvIdx, 10);
      if (isNaN(idx)) return;
      clearTimeout(_cvHoverDebounce);
      _cvHoverDebounce = setTimeout(function () {
        var ev = cvState.filteredEvents[idx];
        if (!ev) return;
        cvShowFramePreview(marker, ev);
        cvState.hoveredMarkerIdx = idx;
        var rows = marker.closest(".cv-participant-rows");
        if (rows) rows.classList.add("cv-markers-dimmed");
        renderSummaryLane();
      }, 60);
    });
    rowsContainer.addEventListener("mouseout", function (e) {
      var marker = e.target.closest(".cv-event-marker");
      if (!marker) return;
      clearTimeout(_cvHoverDebounce);
      _cvHoverDebounce = null;
      cvHideFramePreview();
      cvState.hoveredMarkerIdx = -1;
      var rows = marker.closest(".cv-participant-rows");
      if (rows) rows.classList.remove("cv-markers-dimmed");
      renderSummaryLane();
    });

    rowsContainer.addEventListener("mousedown", onDragMousedown);
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

      // Highlight zone containing hovered marker
      if (cvState.hoveredMarkerIdx !== -1) {
        var hovEvt = cvState.filteredEvents[cvState.hoveredMarkerIdx];
        if (hovEvt && hovEvt.start >= zone.start && hovEvt.start <= zone.end) {
          ctx.strokeStyle = "rgba(" + accent.r + "," + accent.g + "," + accent.b + ",0.9)";
          ctx.lineWidth = 2;
          ctx.strokeRect(x1 + 1, 1, zw - 2, 30);
        }
      }

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

  // --- Selection ---

  function setSelection(start, end, zone) {
    // Clamp to valid range
    start = Math.max(0, start);
    end = Math.min(cvState.duration, end);
    if (end - start < 0.5) return;

    var events = [];
    for (var i = 0; i < cvState.filteredEvents.length; i++) {
      var ev = cvState.filteredEvents[i];
      if (ev.start < end && ev.end > start) events.push(ev);
    }
    if (events.length === 0) return;

    cvState.selection = { start: start, end: end, zone: zone || null, events: events };
    renderSelectionOverlay();
    renderDetailPanel();
  }

  function clearSelection() {
    cvState.selection = null;
    var overlays = qsa(".cv-selection-overlay");
    for (var i = 0; i < overlays.length; i++) overlays[i].remove();
    var preview = qs(".cv-drag-preview");
    if (preview) preview.remove();
    var detail = qs("#convergenceDetail");
    if (detail) {
      detail.classList.add("hidden");
      detail.innerHTML = "";
    }
  }

  function renderSelectionOverlay() {
    // Remove existing overlays
    var old = qsa(".cv-selection-overlay");
    for (var i = 0; i < old.length; i++) old[i].remove();

    if (!cvState.selection) return;
    var sel = cvState.selection;
    var leftPct = (sel.start / cvState.duration * 100) + "%";
    var widthPct = ((sel.end - sel.start) / cvState.duration * 100) + "%";

    // Summary lane overlay (inside the track, not the spacer)
    var summaryTrack = qs(".cv-summary-track");
    if (summaryTrack) {
      var so = el("div", "cv-selection-overlay");
      so.style.left = leftPct;
      so.style.width = widthPct;
      summaryTrack.appendChild(so);
    }

    // Per-participant row overlays
    var tracks = qsa(".cv-tracks-container");
    for (var t = 0; t < tracks.length; t++) {
      var ro = el("div", "cv-selection-overlay cv-selection-overlay-row");
      ro.style.left = leftPct;
      ro.style.width = widthPct;
      tracks[t].appendChild(ro);
    }
  }

  // --- Summary Lane Click ---

  function handleSummaryClick(e) {
    var canvas = qs(".cv-summary-canvas");
    if (!canvas) return;
    var rect = canvas.getBoundingClientRect();
    var mouseX = e.clientX - rect.left;

    for (var i = 0; i < _summaryHitRects.length; i++) {
      var hr = _summaryHitRects[i];
      if (mouseX >= hr.x1 && mouseX <= hr.x2) {
        var zone = cvState.convergenceZones[hr.zoneIdx];
        if (zone) {
          setSelection(zone.start, zone.end, zone);
          return;
        }
      }
    }
    // Clicked empty space on summary lane
    clearSelection();
  }

  // --- Drag-to-Select ---

  var _drag = { active: false, startX: 0, startTime: 0, preview: null, moved: false,
    tracksRect: null };

  function timeFromMouseX(e) {
    // Use cached rect during drag for consistency, otherwise query live
    var rect = _drag.tracksRect;
    if (!rect) {
      var rowsContainer = qs(".cv-participant-rows");
      if (!rowsContainer) return 0;
      var tracksEl = rowsContainer.querySelector(".cv-tracks-container");
      if (!tracksEl) return 0;
      rect = tracksEl.getBoundingClientRect();
    }
    var frac = Math.max(0, Math.min(1, (e.clientX - rect.left) / rect.width));
    return frac * cvState.duration;
  }

  function onDragMousedown(e) {
    if (e.button !== 0) return;
    if (e.target.closest(".cv-event-marker")) return;
    if (!e.target.closest(".cv-tracks-container") && !e.target.closest(".cv-sub-track")) return;

    // Cache the tracks rect for consistent coordinate conversion during drag
    var tracksEl = e.target.closest(".cv-tracks-container") || e.target.closest(".cv-sub-track").parentElement;
    _drag.tracksRect = tracksEl ? tracksEl.getBoundingClientRect() : null;
    _drag.startX = e.clientX;
    _drag.startTime = timeFromMouseX(e);
    _drag.active = true;
    _drag.moved = false;
    _drag.preview = null;

    document.addEventListener("mousemove", onDragMousemove);
    document.addEventListener("mouseup", onDragMouseup);
    e.preventDefault();
  }

  function onDragMousemove(e) {
    if (!_drag.active) return;
    if (Math.abs(e.clientX - _drag.startX) < 5) return;
    _drag.moved = true;

    var curTime = timeFromMouseX(e);
    var s = Math.min(_drag.startTime, curTime);
    var en = Math.max(_drag.startTime, curTime);

    if (!_drag.preview) {
      // Clear any existing selection while dragging
      clearSelection();
      _drag.preview = el("div", "cv-drag-preview");
      var rowsContainer = qs(".cv-participant-rows");
      if (rowsContainer) rowsContainer.appendChild(_drag.preview);
    }

    // Position preview using pixels relative to the tracks container,
    // offset by the label column width within the rows container
    if (_drag.tracksRect) {
      rowsContainer = qs(".cv-participant-rows");
      var rowsRect = rowsContainer ? rowsContainer.getBoundingClientRect() : _drag.tracksRect;
      var labelOffset = _drag.tracksRect.left - rowsRect.left;
      var leftPx = labelOffset + (s / cvState.duration) * _drag.tracksRect.width;
      var widthPx = ((en - s) / cvState.duration) * _drag.tracksRect.width;
      _drag.preview.style.left = leftPx + "px";
      _drag.preview.style.width = widthPx + "px";
    }
  }

  function onDragMouseup(e) {
    document.removeEventListener("mousemove", onDragMousemove);
    document.removeEventListener("mouseup", onDragMouseup);

    if (!_drag.active) return;
    _drag.active = false;

    if (_drag.preview) _drag.preview.remove();
    _drag.preview = null;

    if (!_drag.moved) return;

    var curTime = timeFromMouseX(e);
    var s = Math.min(_drag.startTime, curTime);
    var en = Math.max(_drag.startTime, curTime);

    _drag.tracksRect = null;
    if (en - s < 1) return; // minimum 1-second range
    setSelection(s, en, null);
  }

  // --- Detail Panel ---

  function renderDetailPanel() {
    var panel = qs("#convergenceDetail");
    if (!panel || !cvState.selection) return;

    panel.innerHTML = "";
    panel.classList.remove("hidden");

    var sel = cvState.selection;
    var frag = document.createDocumentFragment();

    // Header
    var header = el("div", "cv-detail-header");
    var headerText = el("span", "cv-detail-header-text");
    var participantSet = {};
    for (var i = 0; i < sel.events.length; i++) participantSet[sel.events[i].participant] = true;
    var participantCount = Object.keys(participantSet).length;
    headerText.textContent = formatTime(sel.start) + " \u2013 " + formatTime(sel.end)
      + " \u00b7 " + participantCount + " participant" + (participantCount !== 1 ? "s" : "")
      + " \u00b7 " + sel.events.length + " event" + (sel.events.length !== 1 ? "s" : "");
    header.appendChild(headerText);

    var closeBtn = document.createElement("button");
    closeBtn.className = "cv-detail-close";
    closeBtn.textContent = "\u00d7";
    closeBtn.title = "Close";
    closeBtn.addEventListener("click", clearSelection);
    header.appendChild(closeBtn);
    frag.appendChild(header);

    // Actions bar
    var actions = el("div", "cv-detail-actions");
    var addAllArt = document.createElement("button");
    addAllArt.className = "cv-detail-btn";
    addAllArt.textContent = "Add All to Artifacts";
    addAllArt.addEventListener("click", function () { dispatchAllToArtifacts(); });
    actions.appendChild(addAllArt);

    var addAllReel = document.createElement("button");
    addAllReel.className = "cv-detail-btn";
    addAllReel.textContent = "Add All to Reel";
    addAllReel.addEventListener("click", function () { dispatchAllToReel(); });
    actions.appendChild(addAllReel);
    frag.appendChild(actions);

    // Group events by participant (maintain participant order)
    var eventsByPid = {};
    for (var j = 0; j < sel.events.length; j++) {
      var ev = sel.events[j];
      if (!eventsByPid[ev.participant]) eventsByPid[ev.participant] = [];
      eventsByPid[ev.participant].push(ev);
    }

    var participantsDiv = el("div", "cv-detail-participants");
    var detailParticipants = cvState.sortByDensity
      ? cvComputeDensityOrder() : cvState.participants;
    for (var pi = 0; pi < detailParticipants.length; pi++) {
      var pid = detailParticipants[pi];
      var pEvents = eventsByPid[pid];
      if (!pEvents || !pEvents.length) continue;

      var pSection = el("div", "cv-detail-participant");
      var pidLabel = el("div", "cv-detail-pid");
      pidLabel.textContent = pid + " (" + pEvents.length + " event" + (pEvents.length !== 1 ? "s" : "") + ")";
      pSection.appendChild(pidLabel);

      var eventsDiv = el("div", "cv-detail-events");
      for (var ei = 0; ei < pEvents.length; ei++) {
        eventsDiv.appendChild(buildDetailEventRow(pEvents[ei]));
      }
      pSection.appendChild(eventsDiv);
      participantsDiv.appendChild(pSection);
    }
    frag.appendChild(participantsDiv);

    panel.appendChild(frag);
  }

  function buildDetailEventRow(event) {
    var row = el("div", "cv-detail-event");

    // Time range
    var time = el("span", "cv-detail-time");
    time.textContent = formatTime(event.start) + " \u2013 " + formatTime(event.end);
    row.appendChild(time);

    // Source badge
    var badge = el("span", "cv-detail-source-badge");
    var dot = el("span", "cv-detail-source-dot");
    dot.style.background = getEventTypeColor(event.source, event.eventType);
    badge.appendChild(dot);
    badge.appendChild(document.createTextNode(event.source));
    row.appendChild(badge);

    // Event type
    var typeSpan = el("span", "cv-detail-type");
    typeSpan.textContent = event.eventType;
    row.appendChild(typeSpan);

    // Label (truncated)
    if (event.label) {
      var labelSpan = el("span", "cv-detail-label");
      labelSpan.textContent = event.label.length > 60 ? event.label.substring(0, 60) + "\u2026" : event.label;
      labelSpan.title = event.label;
      row.appendChild(labelSpan);
    }

    // Cross-reference badges
    if (window._studioFindOverlappingData && window._studioBuildXrefBadges) {
      var xref = window._studioFindOverlappingData(event.participant, event.start, event.end);
      var badges = window._studioBuildXrefBadges(xref, event.source);
      if (badges) {
        badges.style.position = "relative";
        badges.style.bottom = "auto";
        badges.style.left = "auto";
        row.appendChild(badges);
      }
    }

    // Action buttons
    var btnWrap = el("span", "cv-detail-event-actions");

    var artBtn = document.createElement("button");
    artBtn.className = "cv-detail-add-art";
    artBtn.textContent = "Artifact";
    artBtn.title = "Add to Artifacts";
    artBtn.addEventListener("click", (function (ev) {
      return function (e) {
        e.stopPropagation();
        dispatchToArtifacts(ev);
      };
    })(event));
    btnWrap.appendChild(artBtn);

    var reelBtn = document.createElement("button");
    reelBtn.className = "cv-detail-add-reel";
    reelBtn.textContent = "Reel";
    reelBtn.title = "Add to Reel";
    reelBtn.addEventListener("click", (function (ev) {
      return function (e) {
        e.stopPropagation();
        dispatchToReel(ev);
      };
    })(event));
    btnWrap.appendChild(reelBtn);

    row.appendChild(btnWrap);
    return row;
  }

  // --- Queue Dispatch ---

  function buildQueueItem(event) {
    var item = {
      participant: event.participant,
      start: event.start,
      end: event.end,
    };
    if (event.source === "screenspace") {
      item.desc = event.eventType;
      item.source = "screenspace";
      item.event_type = event.eventType;
      if (event.rawData.events && event.rawData.events.length > 0) {
        item.event_ids = event.rawData.events.map(function (e) { return e.id; });
      } else {
        item.event_ids = [event.rawData.id || event.id];
      }
    } else if (event.source === "transcript") {
      item.desc = event.eventType || "transcript";
      item.source = "transcript";
      if (event.rawData.marks && event.rawData.marks.length > 0) {
        item.mark_ids = event.rawData.marks.map(function (m) { return m.id || m.segment_id; });
      } else {
        item.mark_ids = [event.rawData.id || event.rawData.segment_id || event.id];
      }
    } else {
      // Sheet events: use grid item format (no source) so thumbnails route
      // through api/thumbnail and generation through the sheet path
      var rawRow = event.rawData;
      var cellValue = (rawRow.cells && rawRow.cells[event.participant])
        ? rawRow.cells[event.participant].value : "";
      var parts = event.id.split("_");
      var segIdx = parseInt(parts[parts.length - 1]) || 0;
      var baselineOffset2 = (cvState.baselines && cvState.baselines[event.participant]) || 0;
      var segs = parseClipSegmentsForCell(cellValue, baselineOffset2, CLIPGEN_CONFIG.defaultDuration);
      item.row = rawRow.rowNum;
      item.desc = rawRow.observation || event.label || event.eventType;
      item.timestamp = cellValue;
      item.segIdx = segIdx;
      item.segTotal = segs.length;
    }
    return item;
  }

  function dispatchToArtifacts(event) {
    var state = getState();
    state.artifactQueue.push(buildQueueItem(event));
    if (window._studioRenderArtifactQueue) window._studioRenderArtifactQueue();
  }

  function dispatchToReel(event) {
    var state = getState();
    state.reelQueue.push(buildQueueItem(event));
    if (window._studioRenderReelQueue) window._studioRenderReelQueue();
  }

  function dispatchAllToArtifacts() {
    if (!cvState.selection) return;
    var state = getState();
    var events = cvState.selection.events;
    for (var i = 0; i < events.length; i++) {
      state.artifactQueue.push(buildQueueItem(events[i]));
    }
    if (window._studioRenderArtifactQueue) window._studioRenderArtifactQueue();
  }

  function dispatchAllToReel() {
    if (!cvState.selection) return;
    var state = getState();
    var events = cvState.selection.events;
    for (var i = 0; i < events.length; i++) {
      state.reelQueue.push(buildQueueItem(events[i]));
    }
    if (window._studioRenderReelQueue) window._studioRenderReelQueue();
  }

  // --- Lifecycle ---

  function activate() {
    cvState.active = true;
    if (!cvState.initialized) {
      buildFilterControls();
      cvState.initialized = true;
    }

    if (cvState.baselines === null) {
      // First activation: fetch baselines (server returns seconds)
      apiGet("api/sheet/baseline").then(function (data) {
        cvState.baselines = (data.ok && data.baselines) ? data.baselines : {};
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

    // Keyboard shortcuts
    document.addEventListener("keydown", function (e) {
      if (!cvState.active) return;

      if (e.key === "Escape" && cvState.selection) {
        clearSelection();
        return;
      }

      if ((e.key === "ArrowLeft" || e.key === "ArrowRight") && cvState.selection && cvState.selection.zone) {
        var zones = cvState.convergenceZones;
        if (!zones.length) return;
        var currentIdx = -1;
        var selZone = cvState.selection.zone;
        for (var i = 0; i < zones.length; i++) {
          if (zones[i].start === selZone.start && zones[i].end === selZone.end) {
            currentIdx = i; break;
          }
        }
        if (currentIdx === -1) return;
        var nextIdx = e.key === "ArrowLeft"
          ? Math.max(0, currentIdx - 1)
          : Math.min(zones.length - 1, currentIdx + 1);
        if (nextIdx !== currentIdx) {
          var nextZone = zones[nextIdx];
          setSelection(nextZone.start, nextZone.end, nextZone);
        }
        e.preventDefault();
      }
    });
  }

  // --- Window exports ---
  window.convergenceActivate = activate;
  window.convergenceDeactivate = deactivate;
  window.convergenceResize = debounce(function () {
    if (!cvState.active) return;
    render();
  }, 200);

  init();
})();
