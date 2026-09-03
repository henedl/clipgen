/* Convergence Browser — a tab of the Overview page (overview.html).
 *
 * Pulls events from multiple participants/streams (Screenspace detector
 * events + Transcripts marks + sheet timestamps) onto a single timeline,
 * then clusters moments where many participants do the same thing within a
 * short window into "convergence zones". Renders a SwimLane visualization
 * with cluster callouts; clicking a callout opens an inline detail panel below.
 *
 * Satellite of the overview.js hub: shares state and helpers through the
 * window.ClipgenOverview (OV) namespace (lazy reads inside activate/getState —
 * never top-level destructures). Data arrives via the hub's ensureData();
 * only the per-participant alignment offsets are fetched/persisted here
 * (GET/PUT api/convergence/offsets on the overview blueprint).
 *
 * `cvState._snapshot` records the hub data version the view was built
 * against, so we can detect when a refetch brought new upstream data and the
 * view needs to be rebuilt.
 */

(function () {
  "use strict";

  var P = window.ClipgenPrimitives || {};

  var cvState = {
    active: false,
    initialized: false,
    baselines: null,
    // Per-lane offsets in seconds, {pid: {source: seconds}}; zero keys omitted. Stacks on `baselines`.
    offsets: {},
    // Transient {pid: true}: edit lanes independently. Coupled (default) moves all lanes. Never persisted.
    uncoupled: {},
    offsetsLoaded: false,
    editing: null,           // participant id currently in "unlocked" edit mode
    _dragTx: null,           // scratchpad for the active drag (snapshot of markers + start coords)
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
    // Cross-refs flipped while hidden; a hidden panel measures zero width, so re-render on activate.
    crossRefsStale: false,
    _snapshot: null,
  };

  var _cvFrameCache = createBlobCache();
  var _cvFramePreviewEl = null;
  var _cvHoverDebounce = null;
  // Bumped per preview show so a slow fetch can't paint over a newer hover.
  var _cvPreviewSeq = 0;

  // --- Utilities ---

  function getState() {
    return window.ClipgenOverview.state;
  }

  // Per-lane offset lookup. Returns 0 when the participant or source is unset.
  function offsetFor(pid, source) {
    var o = cvState.offsets && cvState.offsets[pid];
    return (o && o[source]) || 0;
  }

  function getEventTypeColor(source, eventType) {
    if (source === "screenspace") return DETECTOR_COLORS[eventType] || "#888";
    if (source === "transcript") return (MARK_CATEGORIES[eventType] || {}).color || "#0891b2";
    if (source === "composer") return XREF_BADGES.composer.color;
    return XREF_BADGES.sheet.color;
  }

  // --- Frame Preview ---

  function cvFrameUrl(source, participant, startSec) {
    if (source === "screenspace") {
      return "../screenspace/api/video/frame/" + encodeURIComponent(participant)
        + "/" + Math.floor(startSec) + "?w=240";
    }
    return "../studio/api/thumbnail/" + encodeURIComponent(participant) + "/" + Math.floor(startSec);
  }

  function cvEnsureFramePreview() {
    if (_cvFramePreviewEl) return _cvFramePreviewEl;
    var wrap = el("div", "cv-frame-preview hidden");
    var img = document.createElement("img");
    img.decoding = "async";
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
    // Preview at raw video time, never offset-adjusted display time.
    var url = cvFrameUrl(event.source, event.participant, event.rawStart);
    var seq = ++_cvPreviewSeq;

    var cached = _cvFrameCache.get(url);
    if (cached && cached !== "error" && cached !== "loading") {
      img.src = cached;
    } else if (cached === "error") {
      cvHideFramePreview();
      return;
    } else if (!cached) {
      _cvFrameCache.mark(url, "loading");
      img.src = "";
      fetch(url)
        .then(function (r) {
          if (!r.ok) throw new Error("status " + r.status);
          return r.blob();
        })
        .then(function (blob) {
          var objUrl = _cvFrameCache.setBlob(url, blob);
          if (seq === _cvPreviewSeq && !preview.classList.contains("hidden") && img.parentNode) {
            img.src = objUrl;
          }
        })
        .catch(function () {
          _cvFrameCache.mark(url, "error");
          if (seq === _cvPreviewSeq) cvHideFramePreview();
        });
    }

    // Boundary (navigational) events carry a recurrence-aware scene label.
    var sceneLabel = "";
    if (event.navigational && event.rawData && event.rawData.events && event.rawData.events[0]) {
      var md = event.rawData.events[0].metadata;
      if (md && md.scene_label) sceneLabel = md.scene_label;
    }
    lbl.textContent = event.participant + " · " + formatTime(event.rawStart)
      + (sceneLabel ? " · " + sceneLabel : "");
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

    // start/end are display time (raw + offsetFor); rawStart/rawEnd keep source video time.
    function applyOffset(pid, source, rawStart, rawEnd) {
      var off = offsetFor(pid, source);
      return {
        rawStart: Math.max(0, rawStart),
        rawEnd: Math.max(0, rawEnd),
        // Never clamp to 0: negative offsets legitimately go off-edge; clamping fakes zones at t=0.
        start: rawStart + off,
        end: rawEnd + off,
      };
    }

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
            var rawStartSh = segs[s].startSeconds;
            var rawEndSh = rawStartSh + segs[s].duration;
            var tSh = applyOffset(pid, "sheet", rawStartSh, rawEndSh);
            events.push({
              participant: pid,
              start: tSh.start,
              end: tSh.end,
              rawStart: tSh.rawStart,
              rawEnd: tSh.rawEnd,
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
    var ssClusters = window.ClipgenIntakeCluster.clusterIntakeEvents(state.intakeEvents, clusterSec);
    for (var i = 0; i < ssClusters.length; i++) {
      var cl = ssClusters[i];
      var clCount = cl.events ? cl.events.length : 1;
      // Navigational ticks sit at the real boundary, not the cluster's ±5s padded window.
      var sStart = cl.start;
      var sEnd = cl.end;
      if (cl.navigational && cl.events && cl.events.length) {
        sStart = cl.events[0].time_in;
        sEnd = cl.events[cl.events.length - 1].time_out;
      }
      var tSs = applyOffset(cl.participant, "screenspace", sStart, sEnd);
      events.push({
        participant: cl.participant,
        start: tSs.start,
        end: tSs.end,
        rawStart: tSs.rawStart,
        rawEnd: tSs.rawEnd,
        source: "screenspace",
        eventType: cl.event_type || cl.detector || "unknown",
        label: (cl.event_type || cl.detector || "") + " detection"
          + (clCount > 1 ? " (" + clCount + " events)" : ""),
        id: "ss_cl_" + i,
        rawData: cl,
        clusterCount: clCount,
        navigational: !!cl.navigational,
      });
      participantSet[cl.participant] = true;
    }

    // Transcript marks (clustered)
    var trClusters = window.ClipgenIntakeCluster.clusterTranscriptMarks(state.trIntakeMarks, clusterSec);
    for (var j = 0; j < trClusters.length; j++) {
      var tc = trClusters[j];
      var tcCount = tc.marks ? tc.marks.length : 1;
      var tTr = applyOffset(tc.participant, "transcript", tc.start, tc.end);
      events.push({
        participant: tc.participant,
        start: tTr.start,
        end: tTr.end,
        rawStart: tTr.rawStart,
        rawEnd: tTr.rawEnd,
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

    // Composer cuts render as spans (like sheet ranges) and seed convergence zones.
    var composerCuts = state.composerCuts || [];
    for (var ci = 0; ci < composerCuts.length; ci++) {
      var cut = composerCuts[ci];
      if (!cut || !cut.participant) continue;
      if (typeof cut.start !== "number" || typeof cut.end !== "number") continue;
      var tCo = applyOffset(cut.participant, "composer", cut.start, cut.end);
      var cutLabel = cut.label || "cut";
      events.push({
        participant: cut.participant,
        start: tCo.start,
        end: tCo.end,
        rawStart: tCo.rawStart,
        rawEnd: tCo.rawEnd,
        source: "composer",
        eventType: cutLabel,
        label: cutLabel,
        id: "co_" + (cut.id || ci),
        rawData: cut,
      });
      participantSet[cut.participant] = true;
    }

    // Max of raw and offset ends: a negative offset must not shrink the timeline.
    var maxEnd = 0;
    for (var k = 0; k < events.length; k++) {
      if (events[k].end > maxEnd) maxEnd = events[k].end;
      var rawE = (typeof events[k].rawEnd === "number") ? events[k].rawEnd : events[k].end;
      if (rawE > maxEnd) maxEnd = rawE;
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

    // Staleness snapshot: the hub's data version, which only advances when
    // OV.refreshData() actually refetched (see checkStaleness).
    cvState._snapshot = { version: state.dataVersion };
  }

  // --- Convergence Algorithm ---
  // Qualify events with minParticipants inside ±W; merge qualifiers within W.

  function computeConvergenceZones(events, windowSec, minParticipants) {
    if (!events.length) return [];

    var sorted = events.slice().sort(function (a, b) { return a.start - b.start; });

    var qualifying = [];
    var left = 0;
    for (var i = 0; i < sorted.length; i++) {
      var center = sorted[i].start;
      var windowStart = center - windowSec;
      var windowEnd = center + windowSec;
      // Sorted by start: the left edge only advances; still-overlapping intervals hold it.
      while (
        left < sorted.length &&
        sorted[left].start < windowStart &&
        sorted[left].end < windowStart
      ) {
        left++;
      }
      var seen = {};
      for (var j = left; j < sorted.length; j++) {
        if (sorted[j].start > windowEnd) break;
        if (sorted[j].end < windowStart && sorted[j].start < windowStart) continue;
        if (sorted[j].start <= windowEnd && sorted[j].end >= windowStart) {
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
    var rawStarts = [];
    var rawEnds = [];
    for (var i = 0; i < events.length; i++) {
      pSet[events[i].participant] = true;
      starts.push(events[i].start);
      var rs = (typeof events[i].rawStart === "number") ? events[i].rawStart : events[i].start;
      var re = (typeof events[i].rawEnd === "number") ? events[i].rawEnd : events[i].end;
      rawStarts.push(rs);
      rawEnds.push(re);
    }
    var participants = Object.keys(pSet);
    var tightness = stddev(starts);
    var totalP = cvState.participants.length || 1;
    var strength = (participants.length / totalP) * (1 / (1 + tightness / Math.max(windowSec, 1)));
    return {
      start: start,
      end: end,
      rawStart: median(rawStarts),
      rawEnd: median(rawEnds),
      participantCount: participants.length,
      participants: participants,
      events: events,
      tightness: tightness,
      strength: strength,
    };
  }

  function median(values) {
    if (!values || !values.length) return 0;
    var arr = values.slice().sort(function (a, b) { return a - b; });
    var mid = Math.floor(arr.length / 2);
    return arr.length % 2 ? arr[mid] : (arr[mid - 1] + arr[mid]) / 2;
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
    // Navigational events are scaffolding: drawn as ticks, never seeding zones.
    cvState.convergenceZones = computeConvergenceZones(
      filtered.filter(function (e) { return !e.navigational; }),
      cvState.filters.windowSec,
      cvState.filters.minParticipants
    );
  }

  function recalculate() {
    // Re-read baselines each time: Refresh reassigns state.convergenceBaselines to a new object.
    cvState.baselines = getState().convergenceBaselines || {};
    clearSelection();
    clipgenPerf.span("overview.computeConvergence", function () {
      collectAllEvents();
      populateEventTypeChips();
      applyFilters();
    });
    render();
    // _snapshot is current now, so checkStaleness clears the Refresh button paint.
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
    { key: "composer", label: "Composer" },
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

    var resetBtn = P.createBtn({
      label: "Reset offsets",
      icon: "arrow-uturn-left",
      size: "sm",
      onClick: resetAllOffsets,
    });
    resetBtn.id = "cvResetOffsetsBtn";
    resetBtn.classList.add("cv-reset-btn", "hidden");
    resetBtn.title = "Reset per-participant convergence alignment offsets";

    controls.appendChild(minCtl.label);
    controls.appendChild(winCtl.label);
    controls.appendChild(clusCtl.label);
    controls.appendChild(sortSelect);
    controls.appendChild(resetBtn);

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
    return clipgenPerf.span("overview.renderConvergence", renderImpl);
  }

  function renderImpl() {
    cvState.crossRefsStale = false;
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
    // A debounced recalculate can land mid-drag; commit the delta before rebuilding.
    _cvAbortDrag();
    container.textContent = "";
    cvState.swimLaneEl = null;

    if (cvState.events.length === 0) {
      var empty = el("div", "cv-empty-state");
      empty.textContent = "No events loaded yet. Add intake or load a multi-participant sheet to see convergence.";
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
    var SOURCES = CLIPGEN_CONFIG.convergenceSources;

    var swimEvents = cvState.filteredEvents.map(function (e) {
      return {
        p: e.participant,
        source: e.source,
        t: e.start / duration,
        tEnd: e.end / duration,
        label: e.eventType,
        navigational: !!e.navigational,
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
      onEventHover: function (idx, ev, hover, marker) {
        clearTimeout(_cvHoverDebounce);
        if (!hover) { cvHideFramePreview(); return; }
        var origEv = ev._ref;
        if (!origEv) return;
        _cvHoverDebounce = setTimeout(function () {
          if (marker && marker.isConnected) cvShowFramePreview(marker, origEv);
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

    initOffsetEditing(swimLane, participants);

    // Cluster callouts row
    if (cvState.convergenceZones.length > 0) {
      var callouts = el("div", "cv-cluster-callouts");
      callouts.id = "cvClusterCallouts";
      cvState.convergenceZones.forEach(function (z, idx) {
        callouts.appendChild(buildCalloutCard(z, idx));
      });
      container.appendChild(callouts);
    }

    // Reset-offsets toolbar button is hidden until at least one offset is set.
    var resetBtn = qs("#cvResetOffsetsBtn");
    if (resetBtn) {
      var hasOffsets = cvState.offsets && Object.keys(cvState.offsets).length > 0;
      resetBtn.classList.toggle("hidden", !hasOffsets);
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
    // Raw (median) time stays invariant under per-participant offsets.
    ts.textContent = "· " + formatTime(typeof zone.rawStart === "number" ? zone.rawStart : zone.start);
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

  var checkStaleness = window.ClipgenOverview.createStalenessTracker(cvState).check;

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
    // Show the underlying raw cluster span so the time matches the callout.
    var hdrStart = (sel.zone && typeof sel.zone.rawStart === "number") ? sel.zone.rawStart : sel.start;
    var hdrEnd = (sel.zone && typeof sel.zone.rawEnd === "number") ? sel.zone.rawEnd : sel.end;
    headerText.textContent = formatTime(hdrStart) + " – " + formatTime(hdrEnd)
      + " · " + participantCount + " participant" + (participantCount !== 1 ? "s" : "")
      + " · " + sel.events.length + " event" + (sel.events.length !== 1 ? "s" : "");
    header.appendChild(headerText);

    var closeBtn = P.createBtn({ icon: "x-mark", size: "sm", variant: "bare", onClick: clearSelection });
    closeBtn.classList.add("cv-detail-close");
    header.appendChild(closeBtn);
    host.appendChild(header);

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
    // Raw video time: what the user sees when jumping to the source.
    time.textContent = formatTime(event.rawStart) + " – " + formatTime(event.rawEnd);
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

    if (window.ClipgenOverview.findOverlappingData) {
      // Overlap uses raw video time like the other panels. buildXrefBadges is a utils.js global.
      var xref = window.ClipgenOverview.findOverlappingData(event.participant, event.rawStart, event.rawEnd);
      var badges = buildXrefBadges(xref, event.source);
      if (badges) {
        badges.style.position = "relative";
        badges.style.bottom = "auto";
        badges.style.left = "auto";
        row.appendChild(badges);
      }
    }
    return row;
  }

  // Queue dispatch was dropped: Studio's queues are in-page state Overview cannot reach.

  // --- Lifecycle ---

  function activate() {
    cvState.active = true;
    if (!cvState.initialized) {
      buildFilterControls();
      cvState.initialized = true;
    }

    if (cvState.baselines === null) {
      // First activation: the hub's ensureData() plus the convergence-only alignment offsets.
      Promise.all([
        window.ClipgenOverview.ensureData(),
        apiGet("api/convergence/offsets").catch(function () { return { ok: false }; }),
      ]).then(function (results) {
        cvState.baselines = getState().convergenceBaselines || {};
        var oData = results[1];
        cvState.offsets = (oData && oData.ok && oData.offsets) ? oData.offsets : {};
        cvState.uncoupled = {};
        cvState.offsetsLoaded = true;
        recalculate();
      });
    } else if (cvState._snapshot && cvState._snapshot.version !== getState().dataVersion) {
      // The hub refetched; recalculate() re-reads baselines and clears the paint.
      recalculate();
    } else {
      if (cvState.crossRefsStale) render();
      checkStaleness();
    }
  }

  // --- Offsets persistence ---

  var _cvSaveOffsetsTimer = null;
  function cvSaveOffsets() {
    if (!cvState.offsetsLoaded) return;
    clearTimeout(_cvSaveOffsetsTimer);
    _cvSaveOffsetsTimer = setTimeout(function () {
      _cvSaveOffsetsTimer = null;
      fetch("api/convergence/offsets", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ offsets: cvState.offsets || {} }),
      }).catch(function () {
        // Persistence is best-effort; client state remains correct in-session.
      });
    }, 500);
  }

  // --- Per-participant alignment offset editor ---

  function formatOffsetBadge(seconds) {
    if (!seconds) return "";
    var sign = seconds > 0 ? "+" : "−";
    var abs = Math.abs(seconds);
    var rounded = abs >= 10 ? Math.round(abs) : Math.round(abs * 10) / 10;
    return sign + rounded + "s";
  }

  // Parse and clamp to ±max(duration, 60) seconds; non-numeric gives 0.
  function clampOffset(seconds) {
    var num = typeof seconds === "number" ? seconds : parseFloat(seconds);
    if (!isFinite(num)) num = 0;
    var maxAbs = Math.max(cvState.duration || 0, 60);
    return Math.max(-maxAbs, Math.min(maxAbs, num));
  }

  // Drop empty entries so Object.keys(cvState.offsets).length means "any offsets set".
  function pruneParticipant(pid) {
    var o = cvState.offsets[pid];
    if (o && Object.keys(o).length === 0) delete cvState.offsets[pid];
  }

  // source null writes every lane (coupled). Sub-0.05s values clear the lane.
  function commitOffset(pid, source, seconds) {
    if (!pid) return;
    var num = clampOffset(seconds);
    var sources = source ? [source] : CLIPGEN_CONFIG.convergenceSources;
    if (!cvState.offsets[pid]) cvState.offsets[pid] = {};
    for (var i = 0; i < sources.length; i++) {
      var s = sources[i];
      if (Math.abs(num) < 0.05) delete cvState.offsets[pid][s];
      else cvState.offsets[pid][s] = num;
    }
    pruneParticipant(pid);
    recalculate();
    cvSaveOffsets();
  }

  function resetAllOffsets() {
    if (!cvState.offsets || Object.keys(cvState.offsets).length === 0) return;
    if (!window.confirm("Reset offsets for all participants?")) return;
    cvState.offsets = {};
    cvState.uncoupled = {};
    cvState.editing = null;
    recalculate();
    cvSaveOffsets();
  }

  function setEditingParticipant(pid) {
    // Single editor: switching participants locks the current one; its offset is already committed.
    if (cvState.editing === pid) {
      cvState.editing = null;
    } else {
      cvState.editing = pid || null;
    }
    applyEditingClasses();
  }

  // commitOffset prunes zeroed lanes, so an entry implies an adjusted lane; check anyway.
  function participantHasOffset(pid) {
    var o = cvState.offsets && cvState.offsets[pid];
    return !!(o && Object.keys(o).length);
  }

  // Locked-label badge: shared value when coupled, "split" when lanes diverge (survives reload).
  function offsetBadgeText(pid) {
    if (!participantHasOffset(pid)) return "";
    var sources = CLIPGEN_CONFIG.convergenceSources;
    var first = offsetFor(pid, sources[0]);
    for (var i = 1; i < sources.length; i++) {
      if (offsetFor(pid, sources[i]) !== first) return "split";
    }
    return formatOffsetBadge(first);
  }

  // Tooltip listing each lane, so diverged lanes read without unlocking.
  function offsetSummaryTitle(pid) {
    var sources = CLIPGEN_CONFIG.convergenceSources;
    var parts = [];
    for (var i = 0; i < sources.length; i++) {
      var v = offsetFor(pid, sources[i]);
      parts.push(sources[i] + " " + (formatOffsetBadge(v) || "0s"));
    }
    return parts.join(" · ");
  }

  // Number input; source null is the coupled input. Escape reverts, change commits.
  function makeOffsetInput(pid, source) {
    var input = document.createElement("input");
    input.type = "number";
    input.className = (source ? "cv-offset-lane-input" : "cv-offset-input") + " cg-mono";
    input.step = "0.5";
    input.autocomplete = "off";
    input.value = offsetFor(pid, source || CLIPGEN_CONFIG.convergenceSources[0]).toFixed(1);
    if (source) {
      input.dataset.source = source;
      input.title = source + " offset (s)";
    }
    input.addEventListener("mousedown", function (e) { e.stopPropagation(); });
    input.addEventListener("keydown", function (e) {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      else if (e.key === "Escape") {
        input.value = offsetFor(pid, source || CLIPGEN_CONFIG.convergenceSources[0]).toFixed(1);
        input.blur();
      }
    });
    input.addEventListener("change", function () {
      commitOffset(pid, source, parseFloat(input.value));
    });
    return input;
  }

  // Sync an unlocked label's inputs to stored offsets without stealing focus.
  function syncLabelInputs(lbl, pid) {
    if (!lbl) return;
    var inp = lbl.querySelector(".cv-offset-input");
    if (inp) inp.value = offsetFor(pid, CLIPGEN_CONFIG.convergenceSources[0]).toFixed(1);
    var laneInputs = lbl.querySelectorAll(".cv-offset-lane-input");
    for (var li = 0; li < laneInputs.length; li++) {
      laneInputs[li].value = offsetFor(pid, laneInputs[li].dataset.source).toFixed(1);
    }
  }

  function applyEditingClasses() {
    if (!cvState.swimLaneEl) return;
    var sw = cvState.swimLaneEl;
    sw.classList.toggle("is-editing", !!cvState.editing);

    // Events carry is-unlocked too: it disables marker pointer-events that would intercept drags.
    var labels = sw.querySelectorAll(".cg-swim-label");
    for (var i = 0; i < labels.length; i++) {
      labels[i].classList.remove("is-unlocked");
    }
    var rows = sw.querySelectorAll(".cg-swim-row");
    for (var j = 0; j < rows.length; j++) {
      rows[j].classList.remove("is-unlocked");
    }
    var allEvents = sw.querySelectorAll(".cg-swim-event");
    for (var ei = 0; ei < allEvents.length; ei++) {
      allEvents[ei].classList.remove("is-unlocked");
    }
    if (cvState.editing) {
      var lbl = sw.getLabelForParticipant(cvState.editing);
      if (lbl) lbl.classList.add("is-unlocked");
      var rs = sw.getRowsForParticipant(cvState.editing);
      for (var k = 0; k < rs.length; k++) rs[k].classList.add("is-unlocked");
      var evs = sw.getEventsForParticipant(cvState.editing);
      for (var m = 0; m < evs.length; m++) evs[m].classList.add("is-unlocked");
      // Sync values only; never steal focus during a recalculate-driven rerender.
      syncLabelInputs(lbl, cvState.editing);
    }
  }

  function setUncoupled(pid, on) {
    if (on) cvState.uncoupled[pid] = true;
    else delete cvState.uncoupled[pid];
    // Rebuild only this label; no recalculate or save, the offsets are unchanged.
    if (cvState.swimLaneEl) {
      buildOffsetLabel(cvState.swimLaneEl, pid);
      applyEditingClasses();
    }
  }

  function buildOffsetLabel(swimLane, pid) {
    var label = swimLane.getLabelForParticipant(pid);
    if (!label) return;
    var uncoupled = !!cvState.uncoupled[pid];
    label.classList.add("cv-offset-label");
    label.classList.toggle("is-adjusted", participantHasOffset(pid));
    label.classList.toggle("is-uncoupled", uncoupled);
    label.textContent = "";

    var lockBtn = document.createElement("button");
    lockBtn.type = "button";
    lockBtn.className = "cv-offset-lock-btn";
    lockBtn.title = "Unlock to drag or type an alignment offset";
    lockBtn.setAttribute("aria-label", "Toggle alignment for " + pid);
    lockBtn.addEventListener("mousedown", function (e) { e.stopPropagation(); });
    lockBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      setEditingParticipant(pid);
    });
    label.appendChild(lockBtn);

    // Uncouple toggle: offset each data source independently.
    var coupleBtn = document.createElement("button");
    coupleBtn.type = "button";
    coupleBtn.className = "cv-offset-couple-btn";
    coupleBtn.title = uncoupled
      ? "Re-couple lanes (offset all sources together)"
      : "Uncouple lanes (offset each source separately)";
    coupleBtn.setAttribute("aria-label", "Toggle per-lane offsets for " + pid);
    coupleBtn.addEventListener("mousedown", function (e) { e.stopPropagation(); });
    coupleBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      setUncoupled(pid, !cvState.uncoupled[pid]);
    });
    label.appendChild(coupleBtn);

    var pidSpan = el("span", "cv-offset-pid");
    pidSpan.textContent = pid;
    label.appendChild(pidSpan);

    var badge = el("span", "cv-offset-badge");
    badge.textContent = offsetBadgeText(pid);
    if (participantHasOffset(pid)) badge.title = offsetSummaryTitle(pid);
    label.appendChild(badge);

    // Coupled input and per-lane inputs are both built; CSS picks via .is-uncoupled.is-unlocked.
    label.appendChild(makeOffsetInput(pid, null));

    var laneWrap = el("div", "cv-offset-lane-inputs");
    var sources = CLIPGEN_CONFIG.convergenceSources;
    for (var i = 0; i < sources.length; i++) {
      laneWrap.appendChild(makeOffsetInput(pid, sources[i]));
    }
    label.appendChild(laneWrap);
  }

  function initOffsetEditing(swimLane, participants) {
    if (!swimLane) return;
    for (var i = 0; i < participants.length; i++) {
      buildOffsetLabel(swimLane, participants[i]);
    }
    applyEditingClasses();
    renderVoidOverlays(swimLane, participants);
    installDragHandlers(swimLane);
  }

  // Hatched per-lane overlay where a source has no data: positive offset left, negative right.
  function renderVoidOverlays(swimLane, participants) {
    var lanes = swimLane.querySelector(".cg-swim-lanes");
    if (!lanes) return;
    var existing = lanes.querySelectorAll(".cv-offset-void");
    for (var i = 0; i < existing.length; i++) existing[i].remove();
    if (!cvState.duration || cvState.duration <= 0) return;
    var sources = CLIPGEN_CONFIG.convergenceSources;
    for (var p = 0; p < participants.length; p++) {
      var pid = participants[p];
      for (var s = 0; s < sources.length; s++) {
        var off = offsetFor(pid, sources[s]);
        if (off) applyVoidForSource(swimLane, pid, s, off);
      }
    }
  }

  function _voidKeySelector(pid, sourceIdx) {
    return '.cv-offset-void[data-participant="' + pid + '"][data-source-idx="' + sourceIdx + '"]';
  }

  function applyVoidForSource(swimLane, pid, sourceIdx, off) {
    var lanes = swimLane.querySelector(".cg-swim-lanes");
    if (!lanes) return;
    var existing = lanes.querySelector(_voidKeySelector(pid, sourceIdx));
    var rows = swimLane.getRowsForParticipant(pid);
    var row = rows && rows[sourceIdx];
    if (Math.abs(off) < 0.05 || !cvState.duration || !row) {
      if (existing) existing.remove();
      return;
    }
    var top = parseFloat(row.style.top) || 0;
    var height = parseFloat(row.style.height) || 0;

    var pctOfDuration = Math.abs(off) / cvState.duration * 100;
    pctOfDuration = Math.max(0, Math.min(100, pctOfDuration));

    var voidEl = existing || document.createElement("div");
    if (!existing) {
      voidEl.className = "cv-offset-void";
      voidEl.dataset.participant = pid;
      voidEl.dataset.sourceIdx = sourceIdx;
      lanes.appendChild(voidEl);
    }
    var source = CLIPGEN_CONFIG.convergenceSources[sourceIdx];
    voidEl.style.top = top + "px";
    voidEl.style.height = height + "px";
    if (off > 0) {
      voidEl.style.left = "0";
      voidEl.style.right = "auto";
      voidEl.style.width = pctOfDuration + "%";
      voidEl.title = pid + " · " + source + " · no data for first " + Math.round(off) + "s";
    } else {
      voidEl.style.right = "0";
      voidEl.style.left = "auto";
      voidEl.style.width = pctOfDuration + "%";
      voidEl.title = pid + " · " + source + " · no data for last " + Math.round(-off) + "s";
    }
  }

  // Commit the in-flight drag before renderTimeline() rebuilds; defer a recalculate for the committed value.
  function _cvAbortDrag() {
    var tx = cvState._dragTx;
    if (!tx) return;
    cvState._dragTx = null;
    _cvDragLiveInput = null;
    document.body.style.userSelect = "";
    var deltaSec = (_cvDragLastX - tx.startX) / tx.pxPerSec;
    if (Math.abs(deltaSec) >= 0.05) {
      var num = clampOffset(tx.baseOffset + deltaSec);
      var sources = tx.source ? [tx.source] : CLIPGEN_CONFIG.convergenceSources;
      if (!cvState.offsets[tx.pid]) cvState.offsets[tx.pid] = {};
      for (var i = 0; i < sources.length; i++) {
        var s = sources[i];
        if (Math.abs(num) < 0.05) delete cvState.offsets[tx.pid][s];
        else cvState.offsets[tx.pid][s] = num;
      }
      pruneParticipant(tx.pid);
      cvSaveOffsets();
      // Runs inside recalculate() via renderTimeline(); defer the re-render.
      setTimeout(recalculate, 0);
    }
  }

  // mousemove/mouseup bind once at init; mousedown rebinds per swim-lane rebuild.
  var _cvDragRafPending = false;
  var _cvDragLastX = 0;
  var _cvDragLiveInput = null;

  function installDragHandlers(swimLane) {
    swimLane.addEventListener("mousedown", function (e) {
      if (!cvState.editing) return;
      // Lock button + input mousedowns stopPropagation themselves.
      var pid = cvState.editing;
      var hit = e.target.closest('[data-participant="' + pid + '"]');
      if (!hit) return;
      e.preventDefault();
      var pxPerSec = swimLane.getLanesPxPerSec();
      if (!pxPerSec || !isFinite(pxPerSec) || pxPerSec <= 0) return;

      // Uncoupled drags on a source row scope to that lane; otherwise all lanes move.
      var rowHit = e.target.closest(".cg-swim-row[data-source-idx]");
      var sIdx = (cvState.uncoupled[pid] && rowHit)
        ? parseInt(rowHit.dataset.sourceIdx, 10) : -1;
      var source = sIdx >= 0 ? CLIPGEN_CONFIG.convergenceSources[sIdx] : null;

      var lbl = swimLane.getLabelForParticipant(pid);
      var liveInput = source
        ? (lbl && lbl.querySelector('.cv-offset-lane-input[data-source="' + source + '"]'))
        : (lbl && lbl.querySelector(".cv-offset-input"));
      cvState._dragTx = {
        pid: pid,
        source: source,
        sourceIdx: sIdx,
        startX: e.clientX,
        pxPerSec: pxPerSec,
        baseOffset: offsetFor(pid, source || CLIPGEN_CONFIG.convergenceSources[0]),
        markers: source
          ? swimLane.getEventsForParticipantSource(pid, sIdx)
          : swimLane.getEventsForParticipant(pid),
        liveInput: liveInput,
        swimLane: swimLane,
      };
      _cvDragLastX = e.clientX;
      if (lbl) lbl.classList.add("is-dragging");
      // Highlight the grabbed lane (uncoupled) or all lanes (coupled).
      var rs = source ? [rowHit] : swimLane.getRowsForParticipant(pid);
      for (var r = 0; r < rs.length; r++) { if (rs[r]) rs[r].classList.add("is-dragging"); }
      _cvDragLiveInput = liveInput;
      document.body.style.userSelect = "none";
    });
  }

  function _cvOnDocMouseMove(e) {
    var tx = cvState._dragTx;
    if (!tx) return;
    _cvDragLastX = e.clientX;
    if (_cvDragRafPending) return;
    _cvDragRafPending = true;
    requestAnimationFrame(function () {
      _cvDragRafPending = false;
      var tx2 = cvState._dragTx;
      if (!tx2) return;
      var deltaPx = _cvDragLastX - tx2.startX;
      var deltaSec = deltaPx / tx2.pxPerSec;
      for (var i = 0; i < tx2.markers.length; i++) {
        var m = tx2.markers[i];
        // Stack the drag translation in front of any baseline transform (e.g. translateX(-50%)).
        var orig = m.dataset._origTransform;
        if (orig === undefined) {
          orig = m.style.transform || "";
          m.dataset._origTransform = orig;
        }
        m.style.transform = "translateX(" + deltaPx + "px) " + orig;
      }
      if (_cvDragLiveInput) {
        _cvDragLiveInput.value = (tx2.baseOffset + deltaSec).toFixed(1);
      }
      // Live-update the void overlays; coupled drags touch all lanes, uncoupled only the grabbed one.
      if (tx2.swimLane) {
        var liveOff = tx2.baseOffset + deltaSec;
        if (tx2.source) {
          applyVoidForSource(tx2.swimLane, tx2.pid, tx2.sourceIdx, liveOff);
        } else {
          var sources = CLIPGEN_CONFIG.convergenceSources;
          for (var v = 0; v < sources.length; v++) {
            applyVoidForSource(tx2.swimLane, tx2.pid, v, liveOff);
          }
        }
      }
    });
  }

  function _cvOnDocMouseUp(e) {
    var tx = cvState._dragTx;
    if (!tx) return;
    // Never `||` here: e.clientX is 0 at the viewport's left edge.
    var deltaPx = e.clientX - tx.startX;
    var deltaSec = deltaPx / tx.pxPerSec;
    for (var i = 0; i < tx.markers.length; i++) {
      var m = tx.markers[i];
      if (m.dataset._origTransform !== undefined) {
        m.style.transform = m.dataset._origTransform;
        delete m.dataset._origTransform;
      }
    }
    var sw = tx.swimLane;
    if (sw) {
      var lbl = sw.getLabelForParticipant(tx.pid);
      if (lbl) lbl.classList.remove("is-dragging");
      var rs = sw.getRowsForParticipant(tx.pid);
      for (var r = 0; r < rs.length; r++) rs[r].classList.remove("is-dragging");
    }
    document.body.style.userSelect = "";
    cvState._dragTx = null;
    _cvDragLiveInput = null;
    // Match _cvAbortDrag's threshold; sub-0.05s jiggles skip the recalculate.
    if (Math.abs(deltaSec) >= 0.05) {
      commitOffset(tx.pid, tx.source, tx.baseOffset + deltaSec);
    }
  }

  function deactivate() {
    cvState.active = false;
    closeDetailInline();
    // The preview hangs off document.body and survives panel hiding; cancel the debounce too.
    clearTimeout(_cvHoverDebounce);
    _cvHoverDebounce = null;
    cvHideFramePreview();
  }

  function init() {

    document.addEventListener("mousemove", _cvOnDocMouseMove);
    document.addEventListener("mouseup", _cvOnDocMouseUp);

    // Zone navigation hotkeys; the dispatcher's typing guard covers the offset inputs.
    function _zoneNavReady() {
      return cvState.active && !!(cvState.selection && cvState.selection.zone);
    }
    function _moveZoneSelection(dir) {
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
      var nextIdx = dir < 0
        ? Math.max(0, currentIdx - 1)
        : Math.min(zones.length - 1, currentIdx + 1);
      if (nextIdx !== currentIdx) {
        var nextZone = zones[nextIdx];
        setSelection(nextZone.start, nextZone.end, nextZone);
        renderDetailInline(nextIdx);
      }
    }
    window.ClipgenHotkeys.register([
      { id: "overview.zonePrev", when: _zoneNavReady, handler: function () { _moveZoneSelection(-1); } },
      { id: "overview.zoneNext", when: _zoneNavReady, handler: function () { _moveZoneSelection(1); } },
    ]);
    window.ClipgenHotkeys.registerEscape(function () {
      if (cvState.active && cvState.selection) {
        clearSelection();
        return true;
      }
      return false;
    });
  }

  // Badges are baked into the detail rows, so a settings flip needs render().
  function renderCrossRefs() {
    if (cvState.active) render();
    else cvState.crossRefsStale = true;
  }

  // --- Hub exports (OV namespace) ---
  window.ClipgenOverview.convergenceActivate = activate;
  window.ClipgenOverview.convergenceRenderCrossRefs = renderCrossRefs;
  window.ClipgenOverview.convergenceDeactivate = deactivate;
  window.ClipgenOverview.convergenceResize = debounce(function () {
    if (!cvState.active) return;
    render();
  }, 200);

  init();
})();
