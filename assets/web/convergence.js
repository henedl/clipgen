/* Convergence Browser.
 *
 * Pulls events from multiple participants/streams (Screenspace detector
 * events + Transcripts marks + sheet timestamps) onto a single timeline,
 * then clusters moments where many participants do the same thing within a
 * short window into "convergence zones". Renders a SwimLane visualization
 * with cluster callouts; clicking a callout opens an inline detail panel below.
 *
 * `cvState._snapshot` records the last (ss, tr, sh) input lengths the view
 * was built against, so we can detect when upstream data changed and the
 * view needs to be rebuilt.
 */

(function () {
  "use strict";

  var P = window.ClipgenPrimitives || {};

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
      windowSec: 10,
      clusterSec: 10,
      timeRange: null,
    },
    dataVersion: 0,
    duration: 0,
    participants: [],
    sortByDensity: false,
    swimLaneEl: null,
    _snapshot: null,
  };

  var _cvFrameCache = {};
  var _cvFramePreviewEl = null;
  var _cvHoverDebounce = null;

  // --- Utilities ---

  function getState() {
    return window._studioState;
  }

  function getEventTypeColor(source, eventType) {
    if (source === "screenspace") return DETECTOR_COLORS[eventType] || "#888";
    if (source === "transcript") return (MARK_CATEGORIES[eventType] || {}).color || "#0891b2";
    return XREF_BADGES.sheet.color;
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
      fetch(url)
        .then(function (r) {
          if (!r.ok) throw new Error("status " + r.status);
          return r.blob();
        })
        .then(function (blob) {
          var objUrl = URL.createObjectURL(blob);
          _cvFrameCache[url] = objUrl;
          if (!preview.classList.contains("hidden") && img.parentNode) {
            img.src = objUrl;
          }
        })
        .catch(function () {
          _cvFrameCache[url] = "error";
          cvHideFramePreview();
        });
    }

    lbl.textContent = event.participant + " · " + formatTime(event.start);
    preview.classList.remove("hidden");
    positionTooltipAnchored(preview, markerEl.getBoundingClientRect());
  }

  function cvHideFramePreview() {
    clearTimeout(_cvHoverDebounce);
    _cvHoverDebounce = null;
    if (_cvFramePreviewEl) _cvFramePreviewEl.classList.add("hidden");
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
  //      each other into a single zone.

  function computeConvergenceZones(events, windowSec, minParticipants) {
    if (!events.length) return [];

    var sorted = events.slice().sort(function (a, b) { return a.start - b.start; });

    var qualifying = [];
    for (var i = 0; i < sorted.length; i++) {
      var center = sorted[i].start;
      var seen = {};
      for (var j = 0; j < sorted.length; j++) {
        if (sorted[j].start > center + windowSec) break;
        if (sorted[j].end < center - windowSec && sorted[j].start < center - windowSec) continue;
        if (sorted[j].start <= center + windowSec && sorted[j].end >= center - windowSec) {
          seen[sorted[j].participant] = true;
        }
      }
      if (Object.keys(seen).length >= minParticipants) {
        qualifying.push(i);
      }
    }

    if (!qualifying.length) return [];

    var zones = [];
    var zStart = sorted[qualifying[0]].start;
    var zEnd = sorted[qualifying[0]].end;
    var zEvents = [sorted[qualifying[0]]];

    for (var q = 1; q < qualifying.length; q++) {
      var evt = sorted[qualifying[q]];
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
    populateEventTypeChips();
    applyFilters();
    render();
    // collectAllEvents just refreshed _snapshot to current lengths, so
    // checkStaleness will clear the `is-stale` class on the toolbar button.
    checkStaleness();
  }

  // --- Cluster hue helper ---

  function zoneDominantHue(zone) {
    // Pick the most common eventType in the zone and resolve its hue.
    var counts = {};
    for (var i = 0; i < zone.events.length; i++) {
      var t = zone.events[i].eventType || "uncategorized";
      counts[t] = (counts[t] || 0) + 1;
    }
    var bestKey = null, bestCount = 0;
    for (var k in counts) {
      if (counts[k] > bestCount) { bestCount = counts[k]; bestKey = k; }
    }
    return categoryHue(bestKey || "uncategorized");
  }

  // --- Filter UI: chips, stream toggle, controls ---

  function populateEventTypeChips() {
    var container = qs("#cvEventTypeChips");
    if (!container) return;
    container.textContent = "";

    var streams = cvState.filters.streams;
    var typeMap = {};
    for (var i = 0; i < cvState.events.length; i++) {
      var e = cvState.events[i];
      if (streams.length > 0 && streams.indexOf(e.source) < 0) continue;
      typeMap[e.eventType] = (typeMap[e.eventType] || 0) + 1;
    }
    var keys = Object.keys(typeMap).sort();
    keys.forEach(function (type) {
      var chip = P.createFilterChip({
        label: type,
        active: cvState.filters.eventTypes.indexOf(type) >= 0,
        count: typeMap[type],
        onClick: function () { toggleEventType(type); },
      });
      chip.dataset.eventType = type;
      container.appendChild(chip);
    });
  }

  function toggleEventType(type) {
    var idx = cvState.filters.eventTypes.indexOf(type);
    if (idx >= 0) {
      cvState.filters.eventTypes.splice(idx, 1);
    } else {
      cvState.filters.eventTypes.push(type);
    }
    populateEventTypeChips();
    onFilterChange();
  }

  var STREAM_DEFS = [
    { key: "all", label: "All streams" },
    { key: "sheet", label: "Sheet" },
    { key: "screenspace", label: "Screenspace" },
    { key: "transcript", label: "Transcript" },
  ];

  function buildFilterControls() {
    var controls = qs("#convergenceControls");
    var filters = qs("#convergenceFilters");
    if (!controls || !filters) return;

    controls.textContent = "";
    filters.textContent = "";

    // Controls: Min participants / Window±s / Cluster s + sort dropdown
    function addNumberControl(text, suffix, value, onChange) {
      var label = el("label", "cv-control-label");
      var labelText = document.createTextNode(text);
      label.appendChild(labelText);
      var input = document.createElement("input");
      input.type = "number";
      input.className = "cv-control-input cg-mono";
      input.value = String(value);
      input.autocomplete = "off";
      input.addEventListener("input", onChange);
      label.appendChild(input);
      if (suffix) {
        var s = el("span", "cv-control-suffix cg-mono");
        s.textContent = suffix;
        label.appendChild(s);
      }
      return { label: label, input: input };
    }

    var minCtl = addNumberControl("Min participants", null, cvState.filters.minParticipants, debouncedFilterChange);
    minCtl.input.min = "2";
    var winCtl = addNumberControl("Window", "±s", cvState.filters.windowSec, debouncedFilterChange);
    winCtl.input.min = "1";
    winCtl.input.max = "60";
    var clusCtl = addNumberControl("Cluster", "s", cvState.filters.clusterSec, debouncedRecalculate);
    clusCtl.input.id = "cvClusterThreshold";
    clusCtl.input.min = "1";
    clusCtl.input.max = "60";

    var sortSelect = document.createElement("select");
    sortSelect.className = "cv-control-select";
    sortSelect.id = "cvSortSelect";
    var optDensity = document.createElement("option");
    optDensity.value = "density";
    optDensity.textContent = "Sort by density";
    var optTime = document.createElement("option");
    optTime.value = "time";
    optTime.textContent = "Sort by time";
    sortSelect.appendChild(optTime);
    sortSelect.appendChild(optDensity);
    sortSelect.value = cvState.sortByDensity ? "density" : "time";
    sortSelect.addEventListener("change", function () {
      cvState.sortByDensity = (sortSelect.value === "density");
      render();
    });

    var refreshBtn = P.createBtn({
      label: "Refresh",
      icon: "arrow-path",
      size: "sm",
      onClick: function () {
        bootstrapIntakeData().then(recalculate);
      },
    });
    refreshBtn.id = "cvRefreshBtn";
    refreshBtn.classList.add("cv-refresh-btn");
    refreshBtn.title = "Re-fetch upstream data and recompute convergence";

    controls.appendChild(minCtl.label);
    controls.appendChild(winCtl.label);
    controls.appendChild(clusCtl.label);
    controls.appendChild(sortSelect);
    controls.appendChild(refreshBtn);

    // Filters: stream buttons + divider + chip row
    var streamRow = el("div", "cv-stream-row");
    STREAM_DEFS.forEach(function (def) {
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "cv-stream-btn" + (def.key === "all" ? " is-active" : "");
      btn.textContent = def.label;
      btn.dataset.stream = def.key;
      btn.addEventListener("click", function () { onStreamToggle(def.key); });
      streamRow.appendChild(btn);
    });

    var divider = el("div", "cv-stream-divider");
    streamRow.appendChild(divider);

    var typeChips = el("div", "cv-event-type-row");
    typeChips.id = "cvEventTypeChips";
    streamRow.appendChild(typeChips);

    filters.appendChild(streamRow);
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

    var btns = qsa(".cv-stream-btn");
    var isAll = cvState.filters.streams.length === 0;
    for (var i = 0; i < btns.length; i++) {
      var key = btns[i].dataset.stream;
      if (key === "all") {
        btns[i].classList.toggle("is-active", isAll);
      } else {
        btns[i].classList.toggle("is-active", cvState.filters.streams.indexOf(key) >= 0);
      }
    }

    cvState.filters.eventTypes = [];
    populateEventTypeChips();
    onFilterChange();
  }

  function syncFilterInputs() {
    var controls = qs("#convergenceControls");
    if (!controls) return;
    var inputs = controls.querySelectorAll("input[type=number]");
    if (inputs[0]) cvState.filters.minParticipants = Math.max(2, parseInt(inputs[0].value, 10) || 2);
    if (inputs[1]) cvState.filters.windowSec = Math.max(1, parseInt(inputs[1].value, 10) || 10);
    if (inputs[2]) cvState.filters.clusterSec = Math.max(1, parseInt(inputs[2].value, 10) || 10);
  }

  var debouncedRecalculate = debounce(function () {
    syncFilterInputs();
    recalculate();
  }, 250);

  function onFilterChange() {
    syncFilterInputs();
    applyFilters();
    if (cvState.selection) {
      if (cvState.selection.zone) {
        clearSelection();
      } else {
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

  // --- Rendering ---

  function render() {
    renderTimeline();
    if (cvState.selection && cvState.selection.zone) {
      // Re-render the inline detail panel after the swim-lane rebuild.
      var idx = cvState.convergenceZones.indexOf(cvState.selection.zone);
      if (idx >= 0) renderDetailInline(idx);
    }
  }

  function renderTimeline() {
    var container = qs("#convergenceTimeline");
    if (!container) return;
    container.textContent = "";
    cvState.swimLaneEl = null;

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

    if (cvState.convergenceZones.length === 0) {
      var noZones = el("div", "cv-no-convergence");
      noZones.textContent = "No convergence detected. Try widening the window or lowering the threshold.";
      container.appendChild(noZones);
    }

    var participants = cvState.sortByDensity ? cvComputeDensityOrder() : cvState.participants;
    var duration = cvState.duration || 1;
    var SOURCES = ["sheet", "screenspace", "transcript"];

    var swimEvents = cvState.filteredEvents.map(function (e) {
      return {
        p: e.participant,
        source: e.source,
        t: e.start / duration,
        tEnd: e.end / duration,
        label: e.eventType,
        _ref: e,
      };
    });
    var swimClusters = cvState.convergenceZones.map(function (z) {
      return {
        t0: z.start / duration,
        t1: z.end / duration,
        hue: zoneDominantHue(z),
        n: z.participantCount,
      };
    });

    var SUB_ROW_H = 14;
    var swimLane = P.createSwimLane({
      participants: participants,
      sources: SOURCES,
      events: swimEvents,
      clusters: swimClusters,
      durationSec: duration,
      subRowH: SUB_ROW_H,
      onEventHover: function (idx, ev, hover) {
        clearTimeout(_cvHoverDebounce);
        if (!hover) { cvHideFramePreview(); return; }
        var origEv = ev._ref;
        if (!origEv) return;
        _cvHoverDebounce = setTimeout(function () {
          var marker = swimLane.querySelectorAll(".cg-swim-event")[idx];
          if (marker) cvShowFramePreview(marker, origEv);
        }, 60);
      },
      onEventClick: function (idx, ev, mouseEv) {
        var origEv = ev._ref;
        if (!origEv) return;
        for (var zi = 0; zi < cvState.convergenceZones.length; zi++) {
          var z = cvState.convergenceZones[zi];
          if (origEv.start >= z.start && origEv.start <= z.end &&
              z.events.indexOf(origEv) >= 0) {
            setSelection(z.start, z.end, z);
            renderDetailInline(zi);
            mouseEv.stopPropagation();
            return;
          }
        }
      },
      onClusterClick: function (idx, cl, mouseEv) {
        var zone = cvState.convergenceZones[idx];
        if (!zone) return;
        setSelection(zone.start, zone.end, zone);
        renderDetailInline(idx);
        mouseEv.stopPropagation();
      },
    });
    swimLane.classList.add("cv-swimlane");
    container.appendChild(swimLane);
    cvState.swimLaneEl = swimLane;

    // Cluster callouts row
    if (cvState.convergenceZones.length > 0) {
      var callouts = el("div", "cv-cluster-callouts");
      callouts.id = "cvClusterCallouts";
      cvState.convergenceZones.forEach(function (z, idx) {
        callouts.appendChild(buildCalloutCard(z, idx));
      });
      container.appendChild(callouts);
    }
  }

  function buildCalloutCard(zone, idx) {
    var hue = zoneDominantHue(zone);
    var card = el("div", "cv-callout");
    card.dataset.clusterIdx = idx;
    card.style.background = "oklch(0.20 0.06 " + hue + " / 0.5)";
    card.style.borderColor = "oklch(0.5 0.12 " + hue + " / 0.4)";

    var n = el("span", "cv-callout-n cg-mono");
    n.textContent = zone.participantCount + " participant" + (zone.participantCount !== 1 ? "s" : "");
    card.appendChild(n);

    var dur = el("span", "cv-callout-text");
    dur.textContent = "converged within ~" + Math.round(zone.end - zone.start) + "s";
    card.appendChild(dur);

    var ts = el("span", "cv-callout-ts cg-mono");
    ts.textContent = "· " + formatTime(zone.start);
    card.appendChild(ts);

    card.addEventListener("mouseenter", function () {
      if (cvState.swimLaneEl) cvState.swimLaneEl.setHoveredCluster(idx);
    });
    card.addEventListener("mouseleave", function () {
      if (cvState.swimLaneEl) cvState.swimLaneEl.setHoveredCluster(-1);
    });
    card.addEventListener("click", function () {
      // Toggle: clicking the active card again clears the inline panel.
      if (cvState.selection && cvState.selection.zone === zone) {
        clearSelection();
        return;
      }
      setSelection(zone.start, zone.end, zone);
      renderDetailInline(idx);
    });
    return card;
  }

  function syncActiveCallout(activeIdx) {
    var callouts = qsa(".cv-callout");
    for (var i = 0; i < callouts.length; i++) {
      callouts[i].classList.toggle("is-active", parseInt(callouts[i].dataset.clusterIdx, 10) === activeIdx);
    }
  }

  // --- Data Freshness ---

  function checkStaleness() {
    if (!cvState._snapshot || !cvState.active) return;
    var state = getState();
    var stale = (state.intakeEvents.length !== cvState._snapshot.ss) ||
      (state.trIntakeMarks.length !== cvState._snapshot.tr) ||
      ((state.sheetData ? state.sheetData.rows.length : 0) !== cvState._snapshot.sh);

    var btn = qs("#cvRefreshBtn");
    if (btn) {
      btn.classList.toggle("is-stale", stale);
      btn.title = stale
        ? "New upstream data available — click to refresh"
        : "Re-fetch upstream data and recompute convergence";
    }
  }

  // --- Selection ---

  function setSelection(start, end, zone) {
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
    if (cvState.swimLaneEl && zone) {
      var zi = cvState.convergenceZones.indexOf(zone);
      if (zi >= 0) cvState.swimLaneEl.setSelectedCluster(zi);
    }
  }

  function clearSelection() {
    cvState.selection = null;
    if (cvState.swimLaneEl) cvState.swimLaneEl.setSelectedCluster(-1);
    closeDetailInline();
  }

  // --- Detail Panel (inline) ---

  function ensureDetailHost() {
    var timeline = qs("#convergenceTimeline");
    if (!timeline) return null;
    var host = timeline.querySelector("#convergenceDetail");
    if (host) {
      host.className = "cv-detail-inline hidden";
      return host;
    }
    host = document.createElement("div");
    host.id = "convergenceDetail";
    host.className = "cv-detail-inline hidden";
    timeline.appendChild(host);
    return host;
  }

  function renderDetailInline(clusterIdx) {
    if (clusterIdx == null || clusterIdx < 0) return;
    var zone = cvState.convergenceZones[clusterIdx];
    if (!zone) return;
    var host = ensureDetailHost();
    buildDetailContent(host);
    host.classList.remove("hidden");
    syncActiveCallout(clusterIdx);
  }

  function closeDetailInline() {
    var host = qs("#convergenceDetail");
    if (host) {
      host.classList.add("hidden");
      host.textContent = "";
    }
    syncActiveCallout(-1);
  }

  function buildDetailContent(host) {
    host.textContent = "";
    if (!cvState.selection) return;

    var sel = cvState.selection;

    var header = el("div", "cv-detail-header");
    var headerText = el("span", "cv-detail-header-text");
    var participantSet = {};
    for (var i = 0; i < sel.events.length; i++) participantSet[sel.events[i].participant] = true;
    var participantCount = Object.keys(participantSet).length;
    headerText.textContent = formatTime(sel.start) + " – " + formatTime(sel.end)
      + " · " + participantCount + " participant" + (participantCount !== 1 ? "s" : "")
      + " · " + sel.events.length + " event" + (sel.events.length !== 1 ? "s" : "");
    header.appendChild(headerText);

    var closeBtn = P.createBtn({ icon: "x-mark", size: "sm", variant: "bare", onClick: clearSelection });
    closeBtn.classList.add("cv-detail-close");
    header.appendChild(closeBtn);
    host.appendChild(header);

    var actions = el("div", "cv-detail-actions");
    actions.appendChild(P.createBtn({
      label: "Add All to Artifacts", icon: "plus", size: "sm",
      onClick: dispatchAllToArtifacts,
    }));
    actions.appendChild(P.createBtn({
      label: "Add All to Reel", icon: "plus", size: "sm",
      onClick: dispatchAllToReel,
    }));
    host.appendChild(actions);

    var eventsByPid = {};
    for (var j = 0; j < sel.events.length; j++) {
      var ev = sel.events[j];
      if (!eventsByPid[ev.participant]) eventsByPid[ev.participant] = [];
      eventsByPid[ev.participant].push(ev);
    }
    var participantsDiv = el("div", "cv-detail-participants");
    var detailParticipants = cvState.sortByDensity ? cvComputeDensityOrder() : cvState.participants;
    for (var pi = 0; pi < detailParticipants.length; pi++) {
      var pid = detailParticipants[pi];
      var pEvents = eventsByPid[pid];
      if (!pEvents || !pEvents.length) continue;
      var pSection = el("div", "cv-detail-participant");
      var pillRow = el("div", "cv-detail-pid-row");
      pillRow.appendChild(P.createParticipantPill({ id: pid, active: true }));
      var countSpan = el("span", "cv-detail-pid-count cg-mono");
      countSpan.textContent = pEvents.length + " event" + (pEvents.length !== 1 ? "s" : "");
      pillRow.appendChild(countSpan);
      pSection.appendChild(pillRow);
      var eventsDiv = el("div", "cv-detail-events");
      for (var ei = 0; ei < pEvents.length; ei++) {
        eventsDiv.appendChild(buildDetailEventRow(pEvents[ei]));
      }
      pSection.appendChild(eventsDiv);
      participantsDiv.appendChild(pSection);
    }
    host.appendChild(participantsDiv);
  }

  function buildDetailEventRow(event) {
    var row = el("div", "cv-detail-event");

    var time = el("span", "cv-detail-time cg-mono");
    time.textContent = formatTime(event.start) + " – " + formatTime(event.end);
    row.appendChild(time);

    var badge = el("span", "cv-detail-source-badge");
    var dot = el("span", "cv-detail-source-dot");
    dot.style.background = getEventTypeColor(event.source, event.eventType);
    badge.appendChild(dot);
    badge.appendChild(document.createTextNode(event.source));
    row.appendChild(badge);

    var typeSpan = el("span", "cv-detail-type");
    typeSpan.textContent = event.eventType;
    row.appendChild(typeSpan);

    if (event.label) {
      var labelSpan = el("span", "cv-detail-label");
      labelSpan.textContent = event.label.length > 60 ? event.label.substring(0, 60) + "…" : event.label;
      labelSpan.title = event.label;
      row.appendChild(labelSpan);
    }

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

    var btnWrap = el("span", "cv-detail-event-actions");
    var artBtn = P.createBtn({
      label: "Artifact", icon: "plus", size: "sm", variant: "bare",
      onClick: function (e) { e.stopPropagation(); dispatchToArtifacts(event); },
    });
    artBtn.classList.add("cv-detail-add-art");
    btnWrap.appendChild(artBtn);
    var reelBtn = P.createBtn({
      label: "Reel", icon: "plus", size: "sm", variant: "bare",
      onClick: function (e) { e.stopPropagation(); dispatchToReel(event); },
    });
    reelBtn.classList.add("cv-detail-add-reel");
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

  // --- Bootstrap ---
  //
  // Studio loads Screenspace events + Transcript marks lazily — those pollers
  // only kick in when their respective Intake tabs activate. If the user lands
  // on Convergence first (or refreshes while it's active), state.intakeEvents
  // and state.trIntakeMarks are still empty and Convergence would only see
  // sheet data. Fetch both endpoints once on first activation so all three
  // streams are populated before the swim lane renders.
  function bootstrapIntakeData() {
    var s = getState();
    var jobs = [];
    jobs.push(
      apiGet("../screenspace/api/events?excluded=false")
        .then(function (data) {
          if (data && data.ok && Array.isArray(data.events)) {
            s.intakeEvents = data.events;
          }
        })
        .catch(function () {})
    );
    jobs.push(
      apiGet("../transcripts/api/marks")
        .then(function (data) {
          if (data && data.ok && Array.isArray(data.marks)) {
            s.trIntakeMarks = data.marks.filter(function (m) { return m.valid; });
          }
        })
        .catch(function () {})
    );
    return Promise.all(jobs);
  }

  // --- Lifecycle ---

  function activate() {
    cvState.active = true;
    if (!cvState.initialized) {
      buildFilterControls();
      cvState.initialized = true;
    }

    if (cvState.baselines === null) {
      // First activation: fetch baselines + bootstrap intake/transcript data
      // in parallel, then compute. Subsequent activations use the staleness
      // banner so the user controls when to pull in new upstream data.
      Promise.all([
        apiGet("api/sheet/baseline").catch(function () { return { ok: false }; }),
        bootstrapIntakeData(),
      ]).then(function (results) {
        var data = results[0];
        cvState.baselines = (data && data.ok && data.baselines) ? data.baselines : {};
        recalculate();
      });
    } else {
      checkStaleness();
    }
  }

  function deactivate() {
    cvState.active = false;
    closeDetailInline();
  }

  function init() {
    document.addEventListener("visibilitychange", function () {
      if (!document.hidden && cvState.active) {
        checkStaleness();
      }
    });

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
          renderDetailInline(nextIdx);
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
