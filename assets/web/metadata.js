/* Metadata Overview — Tab 5
 *
 * Read-only aggregate statistics across all loaded sessions and streams.
 * Computation + display only — all data comes from window._studioState.
 *
 * Lifecycle: window.metadataActivate / metadataDeactivate / metadataResize
 * Pattern follows convergence.js (IIFE, window exports, state via _studioState).
 */
(function () {
  "use strict";

  // --- Aliases (set in init) ---
  var state;
  var parseClipTimestamps;
  var hexToRgba;
  var clusterIntakeEvents;
  var clusterTranscriptMarks;
  var ROW_FUNCTIONS;
  var SEVERITY_ORDER;
  var severityRank;

  var MD_HISTOGRAM_HEIGHT = 140;

  var mdState = {
    active: false,
    initialized: false,
    cache: null,
    _snapshot: null,
    baselines: null,
    filterParticipants: [],
    collapsedSections: {},
    collisionWindow: 5,
  };

  // --- Helpers ---

  function debounce(fn, ms) {
    var timer;
    return function () {
      clearTimeout(timer);
      timer = setTimeout(fn, ms);
    };
  }

  function median(arr) {
    if (!arr.length) return 0;
    var sorted = arr.slice().sort(function (a, b) { return a - b; });
    var mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  }

  function getStudyName() {
    return (state.sheetData && state.sheetData.study) || "study";
  }

  // Baseline-aware timestamp parsing (matches convergence.js pattern)
  function clockToSeconds(ts) {
    var parts = ts.split(":");
    if (parts.length === 3)
      return parseInt(parts[0], 10) * 3600 + parseInt(parts[1], 10) * 60 + parseInt(parts[2], 10);
    if (parts.length === 2)
      return parseInt(parts[0], 10) * 3600 + parseInt(parts[1], 10) * 60;
    return NaN;
  }

  function parseClockSegments(raw) {
    var DEFAULT_DUR = (state.sheetData && state.sheetData.defaultDuration) || 60;
    var cleaned = raw.toLowerCase().replace(/!key/g, "").replace(/[+;,]/g, " ");
    var tokens = cleaned.split(/\s+/).filter(function (t) { return t && t !== "x"; });
    var segments = [];
    for (var i = 0; i < tokens.length; i++) {
      var tok = tokens[i].replace(/\.$/, "").replace(/\./g, ":");
      var dashIdx = -1;
      for (var d = 1; d < tok.length; d++) {
        if (tok[d] === "-" && tok[d - 1] >= "0" && tok[d - 1] <= "9") { dashIdx = d; break; }
      }
      if (dashIdx > 0) {
        var s = clockToSeconds(tok.substring(0, dashIdx));
        var e = clockToSeconds(tok.substring(dashIdx + 1));
        if (!isNaN(s) && !isNaN(e))
          segments.push({ startSeconds: Math.floor(s), duration: Math.max(0, e - s) });
      } else if (tok.indexOf(":") > 0) {
        var sec = clockToSeconds(tok);
        if (!isNaN(sec))
          segments.push({ startSeconds: Math.floor(sec), duration: DEFAULT_DUR });
      }
    }
    return segments;
  }

  function parseSheetTimestamps(cellValue, participant) {
    var baselineOffset = (mdState.baselines && mdState.baselines[participant]) || 0;
    var segs = baselineOffset
      ? parseClockSegments(cellValue)
      : parseClipTimestamps(cellValue);
    if (baselineOffset) {
      for (var i = 0; i < segs.length; i++) {
        segs[i].startSeconds = Math.max(0, segs[i].startSeconds - baselineOffset);
      }
    }
    return segs;
  }

  // --- Participant helpers ---

  function getAllParticipants() {
    var seen = {};
    var list = [];
    if (state.sheetData && state.sheetData.participants) {
      for (var i = 0; i < state.sheetData.participants.length; i++) {
        var p = state.sheetData.participants[i];
        if (!seen[p]) { seen[p] = true; list.push(p); }
      }
    }
    for (var j = 0; j < state.intakeEvents.length; j++) {
      var sp = state.intakeEvents[j].participant;
      if (sp && !seen[sp]) { seen[sp] = true; list.push(sp); }
    }
    for (var k = 0; k < state.trIntakeMarks.length; k++) {
      var tp = state.trIntakeMarks[k].participant;
      if (tp && !seen[tp]) { seen[tp] = true; list.push(tp); }
    }
    list.sort();
    return list;
  }

  function getFilteredEvents(participants) {
    var events = [];
    for (var i = 0; i < state.intakeEvents.length; i++) {
      var ev = state.intakeEvents[i];
      if (ev.excluded) continue;
      if (participants.length && participants.indexOf(ev.participant) === -1) continue;
      events.push(ev);
    }
    return events;
  }

  function getFilteredMarks(participants) {
    var marks = [];
    for (var i = 0; i < state.trIntakeMarks.length; i++) {
      var m = state.trIntakeMarks[i];
      if (participants.length && participants.indexOf(m.participant) === -1) continue;
      marks.push(m);
    }
    return marks;
  }

  // --- Computation engine ---

  function computeCoverage(participants, rows, events, marks) {
    var cov = {};
    for (var i = 0; i < participants.length; i++) {
      cov[participants[i]] = { sheet: 0, screenspace: 0, transcript: 0 };
    }
    // Sheet
    for (var r = 0; r < rows.length; r++) {
      for (var p = 0; p < participants.length; p++) {
        var pid = participants[p];
        var cell = rows[r].cells[pid];
        if (cell && cell.valid) {
          cov[pid].sheet += parseSheetTimestamps(cell.value, pid).length;
        }
      }
    }
    // Screenspace
    for (var e = 0; e < events.length; e++) {
      var ep = events[e].participant;
      if (cov[ep]) cov[ep].screenspace++;
    }
    // Transcript
    for (var m = 0; m < marks.length; m++) {
      var mp = marks[m].participant;
      if (cov[mp]) cov[mp].transcript++;
    }
    return cov;
  }

  function computeEventTypeStats(events, participants) {
    var groups = {};
    for (var i = 0; i < events.length; i++) {
      var ev = events[i];
      var key = ev.event_type || "(unnamed)";
      if (!groups[key]) {
        groups[key] = {
          event_type: key,
          detector: ev.detector || "",
          total_count: 0,
          first_sec: Infinity,
          last_sec: -Infinity,
          sum_time: 0,
          sum_confidence: 0,
          sum_duration: 0,
          participants_seen: {},
          per_participant: {},
        };
      }
      var g = groups[key];
      g.total_count++;
      if (ev.time_in < g.first_sec) g.first_sec = ev.time_in;
      if (ev.time_out > g.last_sec) g.last_sec = ev.time_out;
      g.sum_time += ev.time_in;
      g.sum_confidence += (ev.confidence || 0);
      g.sum_duration += ((ev.time_out || 0) - (ev.time_in || 0));
      g.participants_seen[ev.participant] = true;
      if (!g.per_participant[ev.participant]) {
        g.per_participant[ev.participant] = { count: 0, sum_time: 0 };
      }
      g.per_participant[ev.participant].count++;
      g.per_participant[ev.participant].sum_time += ev.time_in;
    }
    var result = [];
    var keys = Object.keys(groups);
    for (var j = 0; j < keys.length; j++) {
      var g2 = groups[keys[j]];
      var pp = {};
      var ppKeys = Object.keys(g2.per_participant);
      for (var k = 0; k < ppKeys.length; k++) {
        var ppd = g2.per_participant[ppKeys[k]];
        pp[ppKeys[k]] = { count: ppd.count, mean_time: ppd.sum_time / ppd.count };
      }
      result.push({
        event_type: g2.event_type,
        detector: g2.detector,
        total_count: g2.total_count,
        participant_coverage: Object.keys(g2.participants_seen).length,
        participant_total: participants.length,
        first_sec: g2.first_sec === Infinity ? 0 : g2.first_sec,
        last_sec: g2.last_sec === -Infinity ? 0 : g2.last_sec,
        mean_time: g2.total_count ? g2.sum_time / g2.total_count : 0,
        mean_confidence: g2.total_count ? g2.sum_confidence / g2.total_count : 0,
        mean_duration: g2.total_count ? g2.sum_duration / g2.total_count : 0,
        per_participant: pp,
      });
    }
    result.sort(function (a, b) { return b.total_count - a.total_count; });
    return result;
  }

  function computeTranscriptCategoryStats(marks, participants) {
    var groups = {};
    for (var i = 0; i < marks.length; i++) {
      var m = marks[i];
      var cat = m.category || "bookmark";
      if (!groups[cat]) {
        groups[cat] = {
          category: cat,
          total_count: 0,
          first_sec: Infinity,
          last_sec: -Infinity,
          participants_seen: {},
          per_participant: {},
        };
      }
      var g = groups[cat];
      g.total_count++;
      var tin = m.time_in !== undefined ? m.time_in : m.start;
      var tout = m.time_out !== undefined ? m.time_out : m.end;
      if (tin < g.first_sec) g.first_sec = tin;
      if (tout > g.last_sec) g.last_sec = tout;
      g.participants_seen[m.participant] = true;
      if (!g.per_participant[m.participant]) g.per_participant[m.participant] = 0;
      g.per_participant[m.participant]++;
    }
    var result = [];
    var keys = Object.keys(groups);
    for (var j = 0; j < keys.length; j++) {
      var g2 = groups[keys[j]];
      result.push({
        category: g2.category,
        total_count: g2.total_count,
        participant_coverage: Object.keys(g2.participants_seen).length,
        participant_total: participants.length,
        first_sec: g2.first_sec === Infinity ? 0 : g2.first_sec,
        last_sec: g2.last_sec === -Infinity ? 0 : g2.last_sec,
        per_participant: g2.per_participant,
      });
    }
    result.sort(function (a, b) { return b.total_count - a.total_count; });
    return result;
  }

  function computeObservationStats(rows, participants) {
    var result = [];
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var totalTs = ROW_FUNCTIONS.Count(row, participants);
      var uniqueP = ROW_FUNCTIONS.Unique(row, participants);
      var earliest = Infinity;
      var latest = -Infinity;
      for (var p = 0; p < participants.length; p++) {
        var cell = row.cells[participants[p]];
        if (!cell || !cell.valid) continue;
        var segs = parseSheetTimestamps(cell.value, participants[p]);
        for (var s = 0; s < segs.length; s++) {
          if (segs[s].startSeconds < earliest) earliest = segs[s].startSeconds;
          var endSec = segs[s].startSeconds + segs[s].duration;
          if (endSec > latest) latest = endSec;
        }
      }
      result.push({
        observation: row.observation || row.name || "",
        category: row.category || "",
        severity: row.severity || "",
        total_timestamps: totalTs,
        unique_participants: uniqueP,
        participant_total: participants.length,
        earliest_sec: earliest === Infinity ? null : earliest,
        latest_sec: latest === -Infinity ? null : latest,
      });
    }
    return result;
  }

  function computeSeverityDistribution(rows) {
    var dist = {};
    for (var i = 0; i < SEVERITY_ORDER.length; i++) {
      dist[SEVERITY_ORDER[i].label] = 0;
    }
    for (var j = 0; j < rows.length; j++) {
      var sev = (rows[j].severity || "").trim();
      if (!sev) continue;
      // Match against known labels (case-insensitive)
      var matched = false;
      for (var k = 0; k < SEVERITY_ORDER.length; k++) {
        if (SEVERITY_ORDER[k].label.toLowerCase() === sev.toLowerCase()) {
          dist[SEVERITY_ORDER[k].label]++;
          matched = true;
          break;
        }
      }
      if (!matched) {
        if (!dist[sev]) dist[sev] = 0;
        dist[sev]++;
      }
    }
    return dist;
  }

  function computeCategoryBreakdown(rows, participants) {
    var groups = {};
    for (var i = 0; i < rows.length; i++) {
      var cat = rows[i].category || "(uncategorized)";
      if (!groups[cat]) {
        groups[cat] = { category: cat, count: 0, participants_seen: {} };
      }
      groups[cat].count++;
      for (var p = 0; p < participants.length; p++) {
        var cell = rows[i].cells[participants[p]];
        if (cell && cell.valid) {
          groups[cat].participants_seen[participants[p]] = true;
        }
      }
    }
    var result = [];
    var keys = Object.keys(groups);
    for (var j = 0; j < keys.length; j++) {
      var g = groups[keys[j]];
      result.push({
        category: g.category,
        count: g.count,
        participant_coverage: Object.keys(g.participants_seen).length,
        participant_total: participants.length,
      });
    }
    result.sort(function (a, b) { return b.count - a.count; });
    return result;
  }

  function computeSessionSummary(participants, rows, events, marks) {
    var summary = [];
    for (var i = 0; i < participants.length; i++) {
      var pid = participants[i];
      var sheetValid = 0;
      var sheetTs = 0;
      for (var r = 0; r < rows.length; r++) {
        var cell = rows[r].cells[pid];
        if (cell && cell.valid) {
          sheetValid++;
          sheetTs += parseSheetTimestamps(cell.value, pid).length;
        }
      }
      var ssEvents = 0;
      var ssTypes = {};
      for (var e = 0; e < events.length; e++) {
        if (events[e].participant === pid) {
          ssEvents++;
          ssTypes[events[e].event_type || ""] = true;
        }
      }
      var trMarks = 0;
      var trByCat = {};
      for (var m = 0; m < marks.length; m++) {
        if (marks[m].participant === pid) {
          trMarks++;
          var cat = marks[m].category || "bookmark";
          if (!trByCat[cat]) trByCat[cat] = 0;
          trByCat[cat]++;
        }
      }
      summary.push({
        participant: pid,
        sheet_valid_cells: sheetValid,
        sheet_timestamps: sheetTs,
        ss_events: ssEvents,
        ss_event_types: Object.keys(ssTypes).length,
        tr_marks: trMarks,
        tr_by_category: trByCat,
        outlier_flags: [],
      });
    }
    // Outlier detection: flag if value > 3x median
    if (summary.length >= 2) {
      var fields = ["sheet_timestamps", "ss_events", "tr_marks"];
      for (var f = 0; f < fields.length; f++) {
        var vals = [];
        for (var s = 0; s < summary.length; s++) vals.push(summary[s][fields[f]]);
        var med = median(vals);
        if (med > 0) {
          for (var s2 = 0; s2 < summary.length; s2++) {
            if (summary[s2][fields[f]] > med * 3) {
              summary[s2].outlier_flags.push(fields[f]);
            }
          }
        }
      }
    }
    return summary;
  }

  function computeHistogramData(participants, rows, events, marks) {
    // Collect all timestamps with stream labels
    var allTimes = [];
    // Sheet
    for (var r = 0; r < rows.length; r++) {
      for (var p = 0; p < participants.length; p++) {
        var cell = rows[r].cells[participants[p]];
        if (cell && cell.valid) {
          var segs = parseSheetTimestamps(cell.value, participants[p]);
          for (var s = 0; s < segs.length; s++) {
            allTimes.push({ time: segs[s].startSeconds, stream: "sheet" });
          }
        }
      }
    }
    // Screenspace
    for (var e = 0; e < events.length; e++) {
      allTimes.push({ time: events[e].time_in, stream: "screenspace" });
    }
    // Transcript
    for (var m = 0; m < marks.length; m++) {
      var tin = marks[m].time_in !== undefined ? marks[m].time_in : marks[m].start;
      allTimes.push({ time: tin, stream: "transcript" });
    }
    if (!allTimes.length) return { bins: [], binWidth: 0, maxTime: 0, maxCount: 0 };

    var maxTime = 0;
    for (var t = 0; t < allTimes.length; t++) {
      if (allTimes[t].time > maxTime) maxTime = allTimes[t].time;
    }
    if (maxTime <= 0) return { bins: [], binWidth: 0, maxTime: 0, maxCount: 0 };

    // Target 40-60 bins
    var numBins = Math.max(1, Math.min(60, Math.max(40, Math.round(maxTime / 30))));
    var binWidth = maxTime / numBins;
    var bins = [];
    for (var b = 0; b < numBins; b++) {
      bins.push({ sheet: 0, screenspace: 0, transcript: 0 });
    }
    for (var tt = 0; tt < allTimes.length; tt++) {
      var idx = Math.min(numBins - 1, Math.floor(allTimes[tt].time / binWidth));
      bins[idx][allTimes[tt].stream]++;
    }
    var maxCount = 0;
    for (var bc = 0; bc < bins.length; bc++) {
      var total = bins[bc].sheet + bins[bc].screenspace + bins[bc].transcript;
      if (total > maxCount) maxCount = total;
    }
    return { bins: bins, binWidth: binWidth, maxTime: maxTime, maxCount: maxCount };
  }

  function computeCollisions(participants, rows, events, marks, windowSec) {
    // Build per-participant interval lists for each stream
    // Screenspace: use clusters (5s threshold) to avoid double-counting
    var ssClusters = clusterIntakeEvents(events, 5);
    // Transcript: use clusters (5s threshold)
    var trClusters = clusterTranscriptMarks(marks, 5);

    // Group by participant
    var ssByP = {}, trByP = {}, shByP = {};
    for (var i = 0; i < participants.length; i++) {
      ssByP[participants[i]] = [];
      trByP[participants[i]] = [];
      shByP[participants[i]] = [];
    }
    for (var a = 0; a < ssClusters.length; a++) {
      var sc = ssClusters[a];
      if (ssByP[sc.participant]) ssByP[sc.participant].push({ start: sc.start, end: sc.end });
    }
    for (var b = 0; b < trClusters.length; b++) {
      var tc = trClusters[b];
      if (trByP[tc.participant]) trByP[tc.participant].push({ start: tc.start, end: tc.end });
    }
    for (var r = 0; r < rows.length; r++) {
      for (var p = 0; p < participants.length; p++) {
        var pid = participants[p];
        var cell = rows[r].cells[pid];
        if (!cell || !cell.valid) continue;
        var segs = parseSheetTimestamps(cell.value, pid);
        for (var s = 0; s < segs.length; s++) {
          shByP[pid].push({ start: segs[s].startSeconds, end: segs[s].startSeconds + segs[s].duration });
        }
      }
    }

    // Sort each list by start time
    function sortIntervals(arr) {
      arr.sort(function (a, b) { return a.start - b.start; });
    }
    for (var sp = 0; sp < participants.length; sp++) {
      sortIntervals(ssByP[participants[sp]]);
      sortIntervals(trByP[participants[sp]]);
      sortIntervals(shByP[participants[sp]]);
    }

    function countPairCollisions(listA, listB, w) {
      // For each interval in A, check if any in B overlaps within ±w
      // Returns { aHits, bHits } — count of items in A/B that have at least one match
      var aHit = 0, bHitSet = {};
      for (var i = 0; i < listA.length; i++) {
        var a = listA[i];
        var matched = false;
        for (var j = 0; j < listB.length; j++) {
          var b = listB[j];
          if (b.start > a.end + w) break; // sorted, no more overlaps possible
          if (a.start - w < b.end && a.end + w > b.start) {
            matched = true;
            bHitSet[j] = true;
          }
        }
        if (matched) aHit++;
      }
      return { aHits: aHit, bHits: Object.keys(bHitSet).length };
    }

    function computePair(byA, byB, w) {
      var totalCollisions = 0, participantsWith = 0;
      var totalA = 0, totalB = 0, totalAHits = 0, totalBHits = 0;
      for (var i = 0; i < participants.length; i++) {
        var pid = participants[i];
        var la = byA[pid], lb = byB[pid];
        totalA += la.length;
        totalB += lb.length;
        if (!la.length || !lb.length) continue;
        var result = countPairCollisions(la, lb, w);
        totalAHits += result.aHits;
        totalBHits += result.bHits;
        totalCollisions += result.aHits;
        if (result.aHits > 0) participantsWith++;
      }
      return {
        collision_count: totalCollisions,
        participants_with: participantsWith,
        participants_total: participants.length,
        pct_a: totalA > 0 ? Math.round((totalAHits / totalA) * 100) : 0,
        pct_b: totalB > 0 ? Math.round((totalBHits / totalB) * 100) : 0,
        total_a: totalA,
        total_b: totalB,
      };
    }

    return {
      window_seconds: windowSec,
      screenspace_spreadsheet: computePair(ssByP, shByP, windowSec),
      screenspace_transcript: computePair(ssByP, trByP, windowSec),
      transcript_spreadsheet: computePair(trByP, shByP, windowSec),
    };
  }

  function isRowEmpty(row, participants) {
    for (var j = 0; j < participants.length; j++) {
      var c = row.cells[participants[j]];
      if (c && c.hasText) return false;
    }
    return true;
  }

  function computeAllStats(participants) {
    var allP = getAllParticipants();
    var activeP = participants.length ? participants : allP;
    var events = getFilteredEvents(participants);
    var marks = getFilteredMarks(participants);
    var allRows = state.sheetData ? state.sheetData.rows : [];
    // Filter out empty rows (no participant has text) for stats
    var rows = [];
    for (var i = 0; i < allRows.length; i++) {
      if (!isRowEmpty(allRows[i], activeP)) rows.push(allRows[i]);
    }

    return {
      participants: activeP,
      allParticipants: allP,
      hasSheet: !!(state.sheetData && state.sheetData.rows && state.sheetData.rows.length),
      hasScreenspace: state.intakeEvents.length > 0,
      hasTranscript: state.trIntakeMarks.length > 0,
      coverage: computeCoverage(activeP, rows, events, marks),
      eventTypeStats: computeEventTypeStats(events, activeP),
      transcriptCategoryStats: computeTranscriptCategoryStats(marks, activeP),
      observationStats: computeObservationStats(rows, activeP),
      severityDist: computeSeverityDistribution(rows),
      categoryBreakdown: computeCategoryBreakdown(rows, activeP),
      sessionSummary: computeSessionSummary(activeP, rows, events, marks),
      histogramData: computeHistogramData(activeP, rows, events, marks),
      collisionStats: computeCollisions(activeP, rows, events, marks, mdState.collisionWindow),
    };
  }

  // --- Rendering ---

  function renderAll(cache) {
    var panel = qs("#metadataPanel");
    if (!panel) return;
    panel.innerHTML = "";

    // No data at all?
    if (!cache.hasSheet && !cache.hasScreenspace && !cache.hasTranscript) {
      var empty = el("div", "drop-target-empty md-empty-global");
      empty.textContent = "No data loaded. Statistics will appear once data is available from at least one stream.";
      panel.appendChild(empty);
      return;
    }

    panel.appendChild(renderHeaderBar(cache));
    renderFreshnessBanner(panel);
    panel.appendChild(renderSummaryStrip(cache));
    panel.appendChild(renderSection("coverage", "Data Coverage Matrix",
      cache.participants.length + " participants", renderCoverageBody, cache, false, null));
    panel.appendChild(renderSection("event-types", "Per Event Type \u2014 Screenspace",
      cache.eventTypeStats.length + " types", renderEventTypeBody, cache, !cache.hasScreenspace,
      "No screenspace events available."));
    panel.appendChild(renderSection("tr-categories", "Per Category \u2014 Transcript",
      cache.transcriptCategoryStats.length + " categories", renderTranscriptCategoryBody, cache, !cache.hasTranscript,
      "No transcript marks available."));
    panel.appendChild(renderSection("observations", "Per Observation \u2014 Spreadsheet",
      cache.observationStats.length + " observations", renderObservationBody, cache, !cache.hasSheet,
      "No spreadsheet data available."));
    panel.appendChild(renderSection("severity", "Severity Distribution",
      null, renderSeverityBody, cache, !cache.hasSheet,
      "No spreadsheet data available."));
    panel.appendChild(renderSection("cat-breakdown", "Category Breakdown \u2014 Spreadsheet",
      null, renderCategoryBreakdownBody, cache, !cache.hasSheet,
      "No spreadsheet data available."));
    var streamCount = (cache.hasScreenspace ? 1 : 0) + (cache.hasSheet ? 1 : 0) + (cache.hasTranscript ? 1 : 0);
    panel.appendChild(renderSection("collisions", "Cross-Stream Collisions",
      null, renderCollisionBody, cache, streamCount < 2,
      "Cross-stream collisions require data from at least two streams."));
    panel.appendChild(renderSection("sessions", "Session-Level Summary",
      cache.sessionSummary.length + " participants", renderSessionSummaryBody, cache, false, null));

    // Render histogram canvas after DOM is in place
    setTimeout(renderHistogram, 0);
  }

  // --- Header bar ---

  function renderHeaderBar(cache) {
    var bar = el("div", "md-header-bar");

    var title = el("h2", "md-title", "Metadata Overview");
    bar.appendChild(title);

    var pills = el("div", "md-participant-pills");
    var allP = cache.allParticipants;
    for (var i = 0; i < allP.length; i++) {
      var pill = el("button", "md-pill", allP[i]);
      pill.dataset.participant = allP[i];
      if (mdState.filterParticipants.indexOf(allP[i]) >= 0) {
        pill.classList.add("active");
      }
      pill.addEventListener("click", function () {
        var p = this.dataset.participant;
        var idx = mdState.filterParticipants.indexOf(p);
        if (idx >= 0) {
          mdState.filterParticipants.splice(idx, 1);
        } else {
          mdState.filterParticipants.push(p);
        }
        refresh();
      });
      pills.appendChild(pill);
    }
    bar.appendChild(pills);

    var windowLabel = el("label", "md-collision-window-label");
    windowLabel.textContent = "Collision window ";
    var windowInput = document.createElement("input");
    windowInput.type = "number";
    windowInput.min = "1";
    windowInput.max = "60";
    windowInput.value = String(mdState.collisionWindow);
    windowInput.className = "md-collision-input";
    windowInput.autocomplete = "off";
    windowInput.addEventListener("change", function () {
      var val = parseInt(this.value, 10);
      if (isNaN(val) || val < 1) val = 1;
      if (val > 60) val = 60;
      this.value = String(val);
      mdState.collisionWindow = val;
      recomputeCollisions();
    });
    windowLabel.appendChild(windowInput);
    windowLabel.appendChild(document.createTextNode(" s"));
    bar.appendChild(windowLabel);

    var actions = el("div", "md-header-actions");

    var refreshBtn = el("button", "btn btn-small btn-icon md-refresh-btn");
    refreshBtn.innerHTML = '<span class="md-icon md-icon-refresh"></span> Refresh';
    refreshBtn.title = "Re-fetch data and recompute statistics";
    refreshBtn.addEventListener("click", function () {
      mdState._snapshot = null;
      refresh();
    });
    actions.appendChild(refreshBtn);

    var exportBtn = el("button", "btn btn-small btn-icon md-export-btn");
    exportBtn.innerHTML = '<span class="md-icon md-icon-export"></span> JSON';
    exportBtn.title = "Download metadata as JSON";
    exportBtn.addEventListener("click", exportJSON);
    actions.appendChild(exportBtn);

    var csvBtn = el("button", "btn btn-small btn-icon md-export-btn");
    csvBtn.innerHTML = '<span class="md-icon md-icon-export"></span> CSV';
    csvBtn.title = "Download metadata as CSV (4 files)";
    csvBtn.addEventListener("click", exportCSV);
    actions.appendChild(csvBtn);

    bar.appendChild(actions);
    return bar;
  }

  // --- Data freshness banner ---

  function renderFreshnessBanner(panel) {
    var banner = el("div", "md-freshness-banner hidden");
    banner.id = "mdFreshnessBanner";
    banner.innerHTML = '<span class="md-icon md-icon-warning"></span> Screenspace analysis is still running \u2014 statistics may be incomplete. <button class="btn btn-small md-banner-refresh">Refresh</button>';
    var refreshBtn = banner.querySelector(".md-banner-refresh");
    if (refreshBtn) {
      refreshBtn.addEventListener("click", function () {
        mdState._snapshot = null;
        refresh();
      });
    }
    panel.appendChild(banner);

    // Check task status
    apiGet("../screenspace/api/tasks").then(function (data) {
      var running = false;
      if (data.tasks) {
        for (var i = 0; i < data.tasks.length; i++) {
          var s = data.tasks[i].status;
          if (s !== "completed" && s !== "error") { running = true; break; }
        }
      }
      if (running) banner.classList.remove("hidden");
    }).catch(function () {});

    // Staleness banner
    var staleBanner = el("div", "md-stale-banner hidden");
    staleBanner.id = "mdStaleBanner";
    staleBanner.innerHTML = '<span class="md-icon md-icon-info"></span> Data has changed \u2014 <button class="btn btn-small md-banner-refresh">Refresh to update statistics</button>';
    var staleRefresh = staleBanner.querySelector(".md-banner-refresh");
    if (staleRefresh) {
      staleRefresh.addEventListener("click", function () {
        mdState._snapshot = null;
        refresh();
      });
    }
    panel.appendChild(staleBanner);
  }

  // --- Summary charts strip ---

  function renderSummaryStrip(cache) {
    var strip = el("div", "md-summary-strip");

    // Histogram canvas
    var histContainer = el("div", "md-histogram-container");
    var histCanvas = document.createElement("canvas");
    histCanvas.id = "mdHistogramCanvas";
    histContainer.appendChild(histCanvas);
    var histTooltip = el("div", "md-histogram-tooltip hidden");
    histTooltip.id = "mdHistogramTooltip";
    histContainer.appendChild(histTooltip);
    strip.appendChild(histContainer);

    return strip;
  }

  // --- Section wrapper ---

  function renderSection(key, title, countText, bodyFn, cache, isEmpty, emptyMsg) {
    var section = el("div", "md-section");
    section.dataset.section = key;
    if (mdState.collapsedSections[key]) section.classList.add("collapsed");

    var header = el("div", "md-section-header");
    var chevron = el("span", "md-section-chevron");
    header.appendChild(chevron);
    var h3 = el("h3", "", title);
    header.appendChild(h3);
    if (countText) {
      var count = el("span", "md-section-count", countText);
      header.appendChild(count);
    }
    header.addEventListener("click", function () {
      var collapsed = section.classList.toggle("collapsed");
      mdState.collapsedSections[key] = collapsed;
    });
    section.appendChild(header);

    var body = el("div", "md-section-body");
    if (isEmpty) {
      var emptyEl = el("div", "drop-target-empty");
      emptyEl.textContent = emptyMsg;
      body.appendChild(emptyEl);
    } else {
      bodyFn(body, cache);
    }
    section.appendChild(body);
    return section;
  }

  // --- Section 1: Coverage Matrix ---

  function renderCoverageBody(body, cache) {
    var table = el("table", "md-table md-coverage-table");
    var thead = el("thead");
    var hrow = el("tr");
    hrow.appendChild(el("th", "", "Participant"));
    hrow.appendChild(el("th", "", "Sheet"));
    hrow.appendChild(el("th", "", "Screenspace"));
    hrow.appendChild(el("th", "", "Transcript"));
    thead.appendChild(hrow);
    table.appendChild(thead);

    var tbody = el("tbody");
    var participants = cache.participants;
    var maxSheet = 0, maxSS = 0, maxTR = 0;
    for (var i = 0; i < participants.length; i++) {
      var c = cache.coverage[participants[i]];
      if (c.sheet > maxSheet) maxSheet = c.sheet;
      if (c.screenspace > maxSS) maxSS = c.screenspace;
      if (c.transcript > maxTR) maxTR = c.transcript;
    }

    var hm = getComputedStyle(document.documentElement).getPropertyValue("--color-heatmap").trim() || "168, 130, 214";

    for (var j = 0; j < participants.length; j++) {
      var pid = participants[j];
      var cov = cache.coverage[pid];
      var row = el("tr");
      row.classList.add("md-coverage-row");
      row.dataset.participant = pid;

      var nameCell = el("td", "md-coverage-participant", pid);
      row.appendChild(nameCell);

      row.appendChild(makeCoverageCell(cov.sheet, maxSheet, hm));
      row.appendChild(makeCoverageCell(cov.screenspace, maxSS, hm));
      row.appendChild(makeCoverageCell(cov.transcript, maxTR, hm));

      nameCell.addEventListener("click", (function (p) {
        return function () { drillDownParticipant(p); };
      })(pid));

      tbody.appendChild(row);
    }
    table.appendChild(tbody);
    body.appendChild(table);
  }

  function makeCoverageCell(value, max, hm) {
    var td = el("td", "md-coverage-cell");
    td.textContent = value;
    if (value === 0) {
      td.classList.add("md-coverage-zero");
    } else if (max > 0) {
      var t = value / max;
      td.style.backgroundColor = "rgba(" + hm + ", " + (t * 0.45).toFixed(3) + ")";
    }
    return td;
  }

  // --- Section 2: Per Event Type ---

  function renderEventTypeBody(body, cache) {
    var data = cache.eventTypeStats;
    if (!data.length) {
      body.appendChild(el("div", "drop-target-empty", "No event types found."));
      return;
    }
    var maxCount = 0;
    for (var i = 0; i < data.length; i++) {
      if (data[i].total_count > maxCount) maxCount = data[i].total_count;
    }

    var table = el("table", "md-table md-sortable-table");
    var thead = el("thead");
    var hrow = el("tr");
    var cols = [
      { key: "event_type", label: "Event Type" },
      { key: "detector", label: "Detector" },
      { key: "total_count", label: "Count" },
      { key: "participant_coverage", label: "Participants" },
      { key: "first_sec", label: "First" },
      { key: "last_sec", label: "Last" },
      { key: "mean_time", label: "Mean Time" },
      { key: "mean_confidence", label: "Confidence" },
      { key: "mean_duration", label: "Duration" },
    ];
    for (var c = 0; c < cols.length; c++) {
      var th = el("th", "", cols[c].label);
      th.dataset.sort = cols[c].key;
      hrow.appendChild(th);
    }
    thead.appendChild(hrow);
    table.appendChild(thead);

    var tbody = el("tbody");
    table.appendChild(tbody);
    body.appendChild(table);

    function renderRows(sortedData) {
      tbody.innerHTML = "";
      for (var i = 0; i < sortedData.length; i++) {
        var d = sortedData[i];
        var row = el("tr", "md-event-type-row");
        row.dataset.eventType = d.event_type;

        row.appendChild(el("td", "md-et-name", d.event_type));
        row.appendChild(el("td", "md-et-detector", d.detector));

        var countTd = el("td", "md-et-count");
        countTd.appendChild(el("span", "md-count-value", String(d.total_count)));
        var bar = el("span", "md-inline-bar");
        bar.style.width = (maxCount > 0 ? (d.total_count / maxCount) * 100 : 0) + "%";
        countTd.appendChild(bar);
        row.appendChild(countTd);

        row.appendChild(el("td", "", d.participant_coverage + "/" + d.participant_total));
        row.appendChild(el("td", "md-time-cell", formatTime(d.first_sec)));
        row.appendChild(el("td", "md-time-cell", formatTime(d.last_sec)));
        row.appendChild(el("td", "md-time-cell", formatTime(d.mean_time)));
        row.appendChild(el("td", "", d.mean_confidence.toFixed(2)));
        row.appendChild(el("td", "md-time-cell", d.mean_duration.toFixed(1) + "s"));

        // Drill-down
        row.addEventListener("click", (function (et) {
          return function () { drillDownEventType(et); };
        })(d.event_type));

        tbody.appendChild(row);
      }
    }

    renderRows(data);
    makeSortable(table, data, cols, renderRows);
  }

  // --- Sortable table mechanism ---

  function makeSortable(table, data, columns, renderRowsFn) {
    var headers = table.querySelectorAll("th[data-sort]");
    var currentSort = { col: null, asc: true };

    for (var i = 0; i < headers.length; i++) {
      headers[i].classList.add("md-sortable");
      headers[i].addEventListener("click", function () {
        var col = this.dataset.sort;
        if (currentSort.col === col) {
          currentSort.asc = !currentSort.asc;
        } else {
          currentSort.col = col;
          currentSort.asc = true;
        }

        // Update indicators
        for (var j = 0; j < headers.length; j++) {
          headers[j].classList.remove("md-sort-asc", "md-sort-desc");
        }
        this.classList.add(currentSort.asc ? "md-sort-asc" : "md-sort-desc");

        // Sort data
        var sorted = data.slice().sort(function (a, b) {
          var va = a[col], vb = b[col];
          if (typeof va === "string") {
            va = va.toLowerCase();
            vb = (vb || "").toLowerCase();
            if (va < vb) return currentSort.asc ? -1 : 1;
            if (va > vb) return currentSort.asc ? 1 : -1;
            return 0;
          }
          return currentSort.asc ? va - vb : vb - va;
        });
        renderRowsFn(sorted);
      });
    }
  }

  // --- Section 3: Per Category — Transcript ---

  function renderTranscriptCategoryBody(body, cache) {
    var data = cache.transcriptCategoryStats;
    if (!data.length) {
      body.appendChild(el("div", "drop-target-empty", "No transcript marks found."));
      return;
    }

    // Sentiment bar
    var sentBar = el("div", "md-stacked-bar md-sentiment-bar");
    // Order: pain_point first, delight last for sentiment reading
    var catOrder = ["pain_point", "task", "insight", "quote", "bookmark", "delight"];
    var totalMarks = 0;
    var catMap = {};
    for (var i = 0; i < data.length; i++) {
      catMap[data[i].category] = data[i];
      totalMarks += data[i].total_count;
    }
    for (var c = 0; c < catOrder.length; c++) {
      var catData = catMap[catOrder[c]];
      if (!catData || !catData.total_count) continue;
      var seg = el("div", "md-stacked-segment");
      var pct = (catData.total_count / totalMarks) * 100;
      seg.style.flexBasis = pct + "%";
      var catInfo = MARK_CATEGORIES[catOrder[c]] || { color: "#888", label: catOrder[c] };
      seg.style.backgroundColor = catInfo.color;
      seg.title = catInfo.label + ": " + catData.total_count + " (" + pct.toFixed(1) + "%)";
      if (pct > 8) seg.textContent = catData.total_count;
      sentBar.appendChild(seg);
    }
    body.appendChild(sentBar);

    // Table
    var table = el("table", "md-table");
    var thead = el("thead");
    var hrow = el("tr");
    hrow.appendChild(el("th", "", "Category"));
    hrow.appendChild(el("th", "", "Count"));
    hrow.appendChild(el("th", "", "Participants"));
    hrow.appendChild(el("th", "", "First"));
    hrow.appendChild(el("th", "", "Last"));
    thead.appendChild(hrow);
    table.appendChild(thead);

    var tbody = el("tbody");
    for (var j = 0; j < data.length; j++) {
      var d = data[j];
      var row = el("tr", "md-tr-category-row");
      row.dataset.category = d.category;
      var catInfo2 = MARK_CATEGORIES[d.category] || { color: "#888", label: d.category };
      var nameTd = el("td", "");
      var dot = el("span", "md-cat-dot");
      dot.style.backgroundColor = catInfo2.color;
      nameTd.appendChild(dot);
      nameTd.appendChild(document.createTextNode(" " + catInfo2.label));
      row.appendChild(nameTd);
      row.appendChild(el("td", "", String(d.total_count)));
      row.appendChild(el("td", "", d.participant_coverage + "/" + d.participant_total));
      row.appendChild(el("td", "md-time-cell", formatTime(d.first_sec)));
      row.appendChild(el("td", "md-time-cell", formatTime(d.last_sec)));

      row.addEventListener("click", (function (cat) {
        return function () { drillDownTranscriptCategory(cat); };
      })(d.category));

      tbody.appendChild(row);
    }
    table.appendChild(tbody);
    body.appendChild(table);
  }

  // --- Section 4: Per Observation ---

  function renderObservationBody(body, cache) {
    var data = cache.observationStats;
    if (!data.length) {
      body.appendChild(el("div", "drop-target-empty", "No observations found."));
      return;
    }

    var table = el("table", "md-table md-sortable-table");
    var thead = el("thead");
    var hrow = el("tr");
    var cols = [
      { key: "observation", label: "Observation" },
      { key: "category", label: "Category" },
      { key: "severity", label: "Severity" },
      { key: "total_timestamps", label: "Timestamps" },
      { key: "unique_participants", label: "Participants" },
      { key: "earliest_sec", label: "Earliest" },
      { key: "latest_sec", label: "Latest" },
    ];
    for (var c = 0; c < cols.length; c++) {
      var th = el("th", "", cols[c].label);
      th.dataset.sort = cols[c].key;
      hrow.appendChild(th);
    }
    thead.appendChild(hrow);
    table.appendChild(thead);

    var tbody = el("tbody");
    table.appendChild(tbody);
    body.appendChild(table);

    function renderRows(sortedData) {
      tbody.innerHTML = "";
      for (var i = 0; i < sortedData.length; i++) {
        var d = sortedData[i];
        var row = el("tr");

        var obsTd = el("td", "md-obs-name");
        obsTd.textContent = d.observation;
        obsTd.title = d.observation;
        row.appendChild(obsTd);

        row.appendChild(el("td", "md-obs-category", d.category));

        var sevTd = el("td", "md-obs-severity");
        if (d.severity) {
          sevTd.textContent = d.severity;
          var sevCls = severityClass(d.severity);
          if (sevCls) sevTd.classList.add(sevCls);
        }
        row.appendChild(sevTd);

        row.appendChild(el("td", "", String(d.total_timestamps)));

        var partTd = el("td", "md-obs-participants");
        partTd.appendChild(el("span", "md-count-value", d.unique_participants + "/" + d.participant_total));
        var bar = el("span", "md-inline-bar");
        bar.style.width = (d.participant_total > 0 ? (d.unique_participants / d.participant_total) * 100 : 0) + "%";
        partTd.appendChild(bar);
        row.appendChild(partTd);

        row.appendChild(el("td", "md-time-cell", d.earliest_sec !== null ? formatTime(d.earliest_sec) : "\u2014"));
        row.appendChild(el("td", "md-time-cell", d.latest_sec !== null ? formatTime(d.latest_sec) : "\u2014"));

        tbody.appendChild(row);
      }
    }

    renderRows(data);
    makeSortable(table, data, cols, renderRows);
  }

  // --- Section 5: Severity Distribution ---

  function renderSeverityBar(dist, compact) {
    var total = 0;
    var entries = [];
    for (var i = 0; i < SEVERITY_ORDER.length; i++) {
      var label = SEVERITY_ORDER[i].label;
      var count = dist[label] || 0;
      if (count > 0) {
        entries.push({ label: label, count: count });
        total += count;
      }
    }
    // Any non-standard severities
    var keys = Object.keys(dist);
    for (var k = 0; k < keys.length; k++) {
      var found = false;
      for (var s = 0; s < SEVERITY_ORDER.length; s++) {
        if (SEVERITY_ORDER[s].label === keys[k]) { found = true; break; }
      }
      if (!found && dist[keys[k]] > 0) {
        entries.push({ label: keys[k], count: dist[keys[k]] });
        total += dist[keys[k]];
      }
    }
    if (!total) return null;

    var bar = el("div", "md-stacked-bar md-severity-bar" + (compact ? " md-compact" : ""));
    for (var j = 0; j < entries.length; j++) {
      var seg = el("div", "md-stacked-segment");
      var pct = (entries[j].count / total) * 100;
      seg.style.flexBasis = pct + "%";
      var cls = severityClass(entries[j].label);
      if (cls) seg.classList.add(cls);
      seg.title = entries[j].label + ": " + entries[j].count;
      if (!compact && pct > 8) seg.textContent = entries[j].count;
      bar.appendChild(seg);
    }
    return bar;
  }

  function renderSeverityBody(body, cache) {
    var bar = renderSeverityBar(cache.severityDist, false);
    if (bar) {
      body.appendChild(bar);
    } else {
      body.appendChild(el("div", "drop-target-empty", "No severity data."));
      return;
    }

    // Legend / exact counts
    var legend = el("div", "md-severity-legend");
    for (var i = 0; i < SEVERITY_ORDER.length; i++) {
      var label = SEVERITY_ORDER[i].label;
      var count = cache.severityDist[label] || 0;
      if (count === 0) continue;
      var item = el("span", "md-severity-legend-item");
      var dot = el("span", "md-sev-dot");
      var cls = severityClass(label);
      if (cls) dot.classList.add(cls);
      item.appendChild(dot);
      item.appendChild(document.createTextNode(" " + label + ": " + count));
      legend.appendChild(item);
    }
    body.appendChild(legend);
  }

  // --- Section 6: Category Breakdown ---

  function renderCategoryBreakdownBody(body, cache) {
    var data = cache.categoryBreakdown;
    if (!data.length) {
      body.appendChild(el("div", "drop-target-empty", "No categories found."));
      return;
    }
    var maxCount = data[0].count; // Already sorted descending

    for (var i = 0; i < data.length; i++) {
      var d = data[i];
      var row = el("div", "md-cat-bar-row");

      var label = el("span", "md-cat-bar-label", d.category);
      row.appendChild(label);

      var barContainer = el("div", "md-cat-bar-container");
      var bar = el("div", "md-cat-bar");
      bar.style.width = (maxCount > 0 ? (d.count / maxCount) * 100 : 0) + "%";
      barContainer.appendChild(bar);
      row.appendChild(barContainer);

      var countLabel = el("span", "md-cat-bar-count", String(d.count));
      row.appendChild(countLabel);

      var covLabel = el("span", "md-cat-bar-coverage", d.participant_coverage + "/" + d.participant_total + " participants");
      row.appendChild(covLabel);

      body.appendChild(row);
    }
  }

  // --- Section 7: Cross-Stream Collisions ---

  function renderCollisionBody(body, cache) {
    var cs = cache.collisionStats;
    if (!cs) {
      body.appendChild(el("div", "drop-target-empty", "No collision data."));
      return;
    }

    var pairs = [
      { key: "screenspace_spreadsheet", labelA: "Screenspace", labelB: "Spreadsheet" },
      { key: "screenspace_transcript", labelA: "Screenspace", labelB: "Transcript" },
      { key: "transcript_spreadsheet", labelA: "Transcript", labelB: "Spreadsheet" },
    ];

    var note = el("div", "md-collision-note",
      "Collisions within \u00b1" + cs.window_seconds + "s window. Based on clusters, not raw events.");
    body.appendChild(note);

    for (var i = 0; i < pairs.length; i++) {
      var pair = pairs[i];
      var data = cs[pair.key];
      // Skip pairs where both streams are empty
      if (data.total_a === 0 && data.total_b === 0) continue;

      var row = el("div", "md-collision-pair");

      var pairLabel = el("span", "md-collision-pair-label",
        pair.labelA + " \u2194 " + pair.labelB);
      row.appendChild(pairLabel);

      var stats = el("span", "md-collision-pair-stats");
      stats.innerHTML =
        "<strong>" + data.collision_count + "</strong> collisions &middot; " +
        data.participants_with + "/" + data.participants_total + " participants &middot; " +
        data.pct_a + "% of " + pair.labelA.toLowerCase() + " &middot; " +
        data.pct_b + "% of " + pair.labelB.toLowerCase();
      row.appendChild(stats);

      // Visual: two small percentage bars
      var bars = el("div", "md-collision-bars");
      var barA = el("div", "md-collision-bar md-collision-bar-a");
      barA.style.width = data.pct_a + "%";
      barA.title = data.pct_a + "% of " + pair.labelA.toLowerCase() + " clusters overlap";
      bars.appendChild(barA);
      var barB = el("div", "md-collision-bar md-collision-bar-b");
      barB.style.width = data.pct_b + "%";
      barB.title = data.pct_b + "% of " + pair.labelB.toLowerCase() + " items overlap";
      bars.appendChild(barB);
      row.appendChild(bars);

      body.appendChild(row);
    }
  }

  // --- Section 8: Session Summary ---

  function renderSmallMultiples(cache, compact) {
    var data = cache.sessionSummary;
    if (!data.length) return null;

    var container = el("div", "md-small-multiples" + (compact ? " md-compact" : ""));
    // Find max across all participants for shared scale
    var maxVal = 0;
    for (var i = 0; i < data.length; i++) {
      if (data[i].sheet_timestamps > maxVal) maxVal = data[i].sheet_timestamps;
      if (data[i].ss_events > maxVal) maxVal = data[i].ss_events;
      if (data[i].tr_marks > maxVal) maxVal = data[i].tr_marks;
    }

    for (var j = 0; j < data.length; j++) {
      var d = data[j];
      var row = el("div", "md-sm-row");
      if (d.outlier_flags.length) row.classList.add("md-outlier");

      var nameEl = el("span", "md-sm-name", d.participant);
      row.appendChild(nameEl);

      var bars = el("div", "md-sm-bars");
      bars.appendChild(makeSmBar("sheet", d.sheet_timestamps, maxVal, d.outlier_flags));
      bars.appendChild(makeSmBar("screenspace", d.ss_events, maxVal, d.outlier_flags));
      bars.appendChild(makeSmBar("transcript", d.tr_marks, maxVal, d.outlier_flags));
      row.appendChild(bars);

      if (!compact) {
        var counts = el("span", "md-sm-counts",
          d.sheet_timestamps + " / " + d.ss_events + " / " + d.tr_marks);
        row.appendChild(counts);
      }

      if (d.outlier_flags.length && !compact) {
        var warn = el("span", "md-outlier-icon");
        warn.title = "Outlier: " + d.outlier_flags.join(", ");
        row.appendChild(warn);
      }

      container.appendChild(row);
    }

    if (data.length < 2 && !compact) {
      var note = el("div", "md-outlier-note", "Outlier detection requires multiple participants.");
      container.appendChild(note);
    }

    return container;
  }

  function makeSmBar(stream, value, maxVal, outlierFlags) {
    var bar = el("div", "md-sm-bar md-stream-" + stream);
    bar.style.width = (maxVal > 0 ? (value / maxVal) * 100 : 0) + "%";
    bar.title = stream + ": " + value;
    var fieldMap = { sheet: "sheet_timestamps", screenspace: "ss_events", transcript: "tr_marks" };
    if (outlierFlags.indexOf(fieldMap[stream]) >= 0) {
      bar.classList.add("md-sm-bar-outlier");
    }
    return bar;
  }

  function renderSessionSummaryBody(body, cache) {
    var multiples = renderSmallMultiples(cache, false);
    if (multiples) {
      body.appendChild(multiples);
    } else {
      body.appendChild(el("div", "drop-target-empty", "No participant data."));
    }
  }

  // --- Section 9: Temporal Density Histogram ---

  var _histogramHitRects = [];

  function renderHistogram() {
    var canvas = qs("#mdHistogramCanvas");
    if (!canvas || !mdState.cache) return;
    var data = mdState.cache.histogramData;

    var container = canvas.parentElement;
    var rect = container.getBoundingClientRect();
    if (rect.width < 10) return;

    var dpr = window.devicePixelRatio || 1;
    var w = Math.floor(rect.width);
    var h = MD_HISTOGRAM_HEIGHT;
    canvas.width = w * dpr;
    canvas.height = h * dpr;
    canvas.style.width = w + "px";
    canvas.style.height = h + "px";

    var ctx = canvas.getContext("2d");
    ctx.scale(dpr, dpr);

    // Theme colors
    var cs = getComputedStyle(document.documentElement);
    var bgColor = cs.getPropertyValue("--color-surface-alt").trim() || "#f1ece4";
    var borderColor = cs.getPropertyValue("--color-border").trim() || "#e0ddd7";
    var textDim = cs.getPropertyValue("--color-text-dim").trim() || "#6b7280";
    var fontMono = cs.getPropertyValue("--font-mono").trim() || "monospace";

    ctx.clearRect(0, 0, w, h);
    ctx.fillStyle = bgColor;
    ctx.fillRect(0, 0, w, h);

    if (!data || !data.bins.length) {
      ctx.fillStyle = textDim;
      ctx.font = "12px " + fontMono;
      ctx.textAlign = "center";
      ctx.fillText("No events to plot", w / 2, h / 2);
      _histogramHitRects = [];
      return;
    }

    var STREAM_COLORS = {
      sheet: "rgba(234, 179, 8, 0.7)",
      screenspace: "rgba(52, 152, 219, 0.7)",
      transcript: "rgba(16, 163, 74, 0.7)",
    };

    var padLeft = 4;
    var padRight = 4;
    var padTop = 8;
    var padBottom = 20;
    var chartW = w - padLeft - padRight;
    var chartH = h - padTop - padBottom;
    var barW = chartW / data.bins.length;
    var maxC = data.maxCount || 1;

    _histogramHitRects = [];

    // Draw stacked bars
    for (var i = 0; i < data.bins.length; i++) {
      var bin = data.bins[i];
      var x = padLeft + i * barW;
      var stackOrder = ["sheet", "screenspace", "transcript"];
      var yOffset = 0;
      for (var s = 0; s < stackOrder.length; s++) {
        var count = bin[stackOrder[s]];
        if (count <= 0) continue;
        var barH = (count / maxC) * chartH;
        ctx.fillStyle = STREAM_COLORS[stackOrder[s]];
        ctx.fillRect(x + 0.5, padTop + chartH - yOffset - barH, barW - 1, barH);
        yOffset += barH;
      }
      _histogramHitRects.push({
        x: x, w: barW,
        bin: bin,
        startSec: i * data.binWidth,
        endSec: (i + 1) * data.binWidth,
      });
    }

    // Bottom border
    ctx.strokeStyle = borderColor;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(padLeft, padTop + chartH + 0.5);
    ctx.lineTo(padLeft + chartW, padTop + chartH + 0.5);
    ctx.stroke();

    // Time labels
    ctx.fillStyle = textDim;
    ctx.font = "9px " + fontMono;
    ctx.textAlign = "center";
    var labelCount = Math.min(6, data.bins.length);
    var labelStep = Math.max(1, Math.floor(data.bins.length / labelCount));
    for (var lb = 0; lb < data.bins.length; lb += labelStep) {
      var lx = padLeft + lb * barW + barW / 2;
      ctx.fillText(formatTime(lb * data.binWidth), lx, h - 4);
    }
    // Last label
    ctx.fillText(formatTime(data.maxTime), padLeft + chartW - 10, h - 4);
  }

  function initHistogramHover() {
    var canvas = qs("#mdHistogramCanvas");
    var tooltip = qs("#mdHistogramTooltip");
    if (!canvas || !tooltip) return;

    canvas.addEventListener("mousemove", function (e) {
      var rect = canvas.getBoundingClientRect();
      var mx = e.clientX - rect.left;
      var hit = null;
      for (var i = 0; i < _histogramHitRects.length; i++) {
        var hr = _histogramHitRects[i];
        if (mx >= hr.x && mx < hr.x + hr.w) { hit = hr; break; }
      }
      if (hit) {
        var total = hit.bin.sheet + hit.bin.screenspace + hit.bin.transcript;
        tooltip.innerHTML =
          "<strong>" + formatTime(hit.startSec) + " \u2013 " + formatTime(hit.endSec) + "</strong><br>" +
          "Sheet: " + hit.bin.sheet + " &middot; SS: " + hit.bin.screenspace + " &middot; TR: " + hit.bin.transcript +
          " &middot; Total: " + total;
        tooltip.style.left = (e.clientX - rect.left) + "px";
        tooltip.classList.remove("hidden");
      } else {
        tooltip.classList.add("hidden");
      }
    });

    canvas.addEventListener("mouseleave", function () {
      tooltip.classList.add("hidden");
    });
  }

  // --- Drill-down helpers ---

  function drillDownEventType(eventType) {
    state.intakeFilterText = eventType;
    var searchInput = qs("#intakeFilterSearch");
    if (searchInput) searchInput.value = eventType;
    switchToTab("intake");
  }

  function drillDownTranscriptCategory(category) {
    state.trIntakeFilterCategory = category;
    switchToTab("transcript-intake");
  }

  function drillDownParticipant(participant) {
    state.intakeFilterParticipants = [participant];
    switchToTab("intake");
  }

  function switchToTab(tabName) {
    state.activePreviewTab = tabName;
    var allTabs = qsa(".preview-tab");
    for (var i = 0; i < allTabs.length; i++) {
      allTabs[i].classList.remove("active");
      if (allTabs[i].dataset.tab === tabName) allTabs[i].classList.add("active");
    }
    if (window._studioSyncPreviewTab) window._studioSyncPreviewTab();
  }

  // --- JSON Export ---

  function exportJSON() {
    var cache = mdState.cache;
    if (!cache) return;
    var data = {
      study: getStudyName(),
      exported_at: new Date().toISOString(),
      participants: cache.participants,
      coverage_matrix: cache.coverage,
      event_type_stats: cache.eventTypeStats,
      transcript_category_stats: cache.transcriptCategoryStats,
      observation_stats: cache.observationStats,
      severity_distribution: cache.severityDist,
      category_breakdown: cache.categoryBreakdown,
      session_summary: cache.sessionSummary,
      cross_stream_collisions: cache.collisionStats,
    };
    var blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = getStudyName().replace(/[^a-zA-Z0-9_-]/g, "_") + "_metadata.json";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // --- CSV Export ---

  function csvEscape(val) {
    var s = String(val == null ? "" : val);
    if (s.indexOf(",") >= 0 || s.indexOf('"') >= 0 || s.indexOf("\n") >= 0) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  function csvRow(fields) {
    return fields.map(csvEscape).join(",");
  }

  function downloadCSV(filename, content) {
    var blob = new Blob([content], { type: "text/csv" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function exportCSV() {
    var cache = mdState.cache;
    if (!cache) return;
    var prefix = getStudyName().replace(/[^a-zA-Z0-9_-]/g, "_");
    var delay = 0;

    // 1. Events CSV
    if (cache.eventTypeStats.length) {
      var evLines = [csvRow(["event_type", "detector", "total_count", "participant_coverage",
        "participant_total", "first_occurrence_sec", "last_occurrence_sec",
        "mean_time_sec", "mean_confidence", "mean_duration_sec"])];
      for (var i = 0; i < cache.eventTypeStats.length; i++) {
        var e = cache.eventTypeStats[i];
        evLines.push(csvRow([e.event_type, e.detector, e.total_count, e.participant_coverage,
          e.participant_total, e.first_sec.toFixed(1), e.last_sec.toFixed(1),
          e.mean_time.toFixed(1), e.mean_confidence.toFixed(2), e.mean_duration.toFixed(1)]));
      }
      setTimeout(function () { downloadCSV(prefix + "_metadata_events.csv", evLines.join("\n")); }, delay);
      delay += 200;
    }

    // 2. Sessions CSV
    if (cache.sessionSummary.length) {
      var cats = ["pain_point", "delight", "quote", "insight", "task", "bookmark"];
      var sesHeader = ["participant", "spreadsheet_valid_cells", "spreadsheet_timestamps",
        "screenspace_events", "screenspace_event_types", "transcript_marks"];
      for (var c = 0; c < cats.length; c++) sesHeader.push("transcript_" + cats[c] + "s");
      sesHeader.push("outlier");
      var sesLines = [csvRow(sesHeader)];
      for (var j = 0; j < cache.sessionSummary.length; j++) {
        var s = cache.sessionSummary[j];
        var row = [s.participant, s.sheet_valid_cells, s.sheet_timestamps,
          s.ss_events, s.ss_event_types, s.tr_marks];
        for (var k = 0; k < cats.length; k++) row.push(s.tr_by_category[cats[k]] || 0);
        row.push(s.outlier_flags.length > 0 ? "true" : "false");
        sesLines.push(csvRow(row));
      }
      setTimeout(function () { downloadCSV(prefix + "_metadata_sessions.csv", sesLines.join("\n")); }, delay);
      delay += 200;
    }

    // 3. Collisions CSV
    if (cache.collisionStats) {
      var colLines = [csvRow(["pair", "window_seconds", "collision_count", "participants_with_collisions"])];
      var pairs = ["screenspace_spreadsheet", "screenspace_transcript", "transcript_spreadsheet"];
      for (var p = 0; p < pairs.length; p++) {
        var cd = cache.collisionStats[pairs[p]];
        colLines.push(csvRow([pairs[p], cache.collisionStats.window_seconds,
          cd.collision_count, cd.participants_with]));
      }
      setTimeout(function () { downloadCSV(prefix + "_metadata_collisions.csv", colLines.join("\n")); }, delay);
      delay += 200;
    }

    // 4. Observations CSV
    if (cache.observationStats.length) {
      var obsLines = [csvRow(["observation", "category", "severity", "total_timestamps",
        "unique_participants", "first_occurrence_sec", "last_occurrence_sec"])];
      for (var o = 0; o < cache.observationStats.length; o++) {
        var ob = cache.observationStats[o];
        obsLines.push(csvRow([ob.observation, ob.category, ob.severity,
          ob.total_timestamps, ob.unique_participants,
          ob.earliest_sec !== null ? ob.earliest_sec.toFixed(1) : "",
          ob.latest_sec !== null ? ob.latest_sec.toFixed(1) : ""]));
      }
      setTimeout(function () { downloadCSV(prefix + "_metadata_observations.csv", obsLines.join("\n")); }, delay);
    }
  }

  // --- Staleness detection ---

  function takeSnapshot() {
    mdState._snapshot = {
      ss: state.intakeEvents.length,
      tr: state.trIntakeMarks.length,
      sh: state.sheetData ? state.sheetData.rows.length : 0,
    };
  }

  function checkStaleness() {
    if (!mdState._snapshot || !mdState.active) return;
    var stale =
      (state.intakeEvents.length !== mdState._snapshot.ss) ||
      (state.trIntakeMarks.length !== mdState._snapshot.tr) ||
      ((state.sheetData ? state.sheetData.rows.length : 0) !== mdState._snapshot.sh);
    var banner = qs("#mdStaleBanner");
    if (banner) {
      if (stale) {
        banner.classList.remove("hidden");
      } else {
        banner.classList.add("hidden");
      }
    }
  }

  // --- Core lifecycle ---

  function refresh() {
    mdState.cache = computeAllStats(mdState.filterParticipants);
    renderAll(mdState.cache);
    takeSnapshot();
    setTimeout(initHistogramHover, 0);
  }

  function recomputeCollisions() {
    if (!mdState.cache) return;
    var events = getFilteredEvents(mdState.filterParticipants);
    var marks = getFilteredMarks(mdState.filterParticipants);
    var allRows = state.sheetData ? state.sheetData.rows : [];
    var rows = [];
    for (var i = 0; i < allRows.length; i++) {
      if (!isRowEmpty(allRows[i], mdState.cache.participants)) rows.push(allRows[i]);
    }
    mdState.cache.collisionStats = computeCollisions(
      mdState.cache.participants, rows, events, marks, mdState.collisionWindow);
    // Re-render only the collision section
    var section = qs('.md-section[data-section="collisions"]');
    if (section) {
      var body = section.querySelector(".md-section-body");
      if (body) {
        body.innerHTML = "";
        var streamCount = (mdState.cache.hasScreenspace ? 1 : 0) +
          (mdState.cache.hasSheet ? 1 : 0) + (mdState.cache.hasTranscript ? 1 : 0);
        if (streamCount < 2) {
          body.appendChild(el("div", "drop-target-empty",
            "Cross-stream collisions require data from at least two streams."));
        } else {
          renderCollisionBody(body, mdState.cache);
        }
      }
    }
  }

  function activate() {
    mdState.active = true;
    if (!state) {
      state = window._studioState;
      parseClipTimestamps = window._studioParseClipTimestamps;
      hexToRgba = window._studioHexToRgba;
      clusterIntakeEvents = window._studioClusterIntakeEvents;
      clusterTranscriptMarks = window._studioClusterTranscriptMarks;
      ROW_FUNCTIONS = window._studioROW_FUNCTIONS;
      SEVERITY_ORDER = window._studioSEVERITY_ORDER;
      severityRank = window._studioSeverityRank;
    }
    if (mdState._snapshot) {
      checkStaleness();
    }
    if (mdState.baselines === null) {
      // First activation: fetch baselines for clock-time correction
      apiGet("api/sheet/baseline").then(function (data) {
        mdState.baselines = (data.ok && data.baselines) ? data.baselines : {};
        refresh();
      }).catch(function () {
        mdState.baselines = {};
        refresh();
      });
    } else {
      refresh();
    }
  }

  function deactivate() {
    mdState.active = false;
  }

  function resize() {
    if (!mdState.active) return;
    renderHistogram();
  }

  // --- Visibility change ---

  document.addEventListener("visibilitychange", function () {
    if (!document.hidden && mdState.active) {
      checkStaleness();
    }
  });

  // --- Window exports ---
  window.metadataActivate = activate;
  window.metadataDeactivate = deactivate;
  window.metadataResize = debounce(resize, 200);
})();
