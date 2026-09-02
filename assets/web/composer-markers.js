/* clipgen Composer — markers satellite.
 *
 * Fetches and merges the three read-only marker sources for the active
 * participant — Sheet timestamps, Screenspace events, Transcript segments +
 * marks — into CO.state.markers, applying the persisted Convergence per-lane
 * offsets so lanes line up with the video the same way the Convergence Browser
 * does. Each fetch fails soft (empty lane) so Composer works without a
 * spreadsheet or the other tools' manifests. Also owns the lane toggle pills
 * and their `PUT api/ui` persistence.
 *
 * Cross-blueprint fetches (../studio/, ../screenspace/, ../transcripts/) are
 * the established convergence.js pattern. parseClipSegmentsForCell /
 * CLIPGEN_CONFIG / qs are ambient utils.js globals.
 */
(function () {
  "use strict";

  var CO = window.ClipgenComposer;
  var state = CO.state;

  var SOURCES = ["sheet", "screenspace", "transcript"];

  // Stale-response guard: participant switches mid-fetch drop the old payload.
  var _loadVersion = 0;

  function offsetFor(offsets, pid, source) {
    var perPid = offsets && offsets[pid];
    var off = perPid && perPid[source];
    return typeof off === "number" && isFinite(off) ? off : 0;
  }

  function loadMarkers(pid) {
    var version = ++_loadVersion;
    // Overview owns the offsets route (agents/ARCHITECTURE.md); ../studio/ 404s
    // and every lane renders at 0.
    apiGet("../overview/api/convergence/offsets")
      .catch(function () { return {}; })
      .then(function (offData) {
        var offsets = (offData && offData.offsets) || {};
        loadSheetMarkers(pid, version, offsets);
        loadScreenspaceMarkers(pid, version, offsets);
        loadTranscriptMarkers(pid, version, offsets);
      });
  }

  function commitLane(pid, version, source, markers) {
    if (version !== _loadVersion || pid !== state.participant) return;
    // Apply persisted trims; origStart/origEnd keep the source span for tooltip and reset.
    markers.forEach(function (m) {
      var trim = state.trims[m.key];
      if (!trim) return;
      m.origStart = m.start;
      m.origEnd = m.end;
      m.start = trim.start;
      m.end = trim.end;
      m.trimmed = true;
    });
    state.markers[source] = markers;
    if (CO.updateTimelineHeight) CO.updateTimelineHeight();
    if (CO.renderTimeline) CO.renderTimeline();
    if (CO.renderSidebar) CO.renderSidebar();
  }

  // Sheet: one marker per timestamp pair, baseline-converted like Studio/Convergence.
  function loadSheetMarkers(pid, version, offsets) {
    var off = offsetFor(offsets, pid, "sheet");
    Promise.all([
      apiGet("../studio/api/sheet"),
      apiGet("../studio/api/sheet/baseline"),
    ]).then(function (results) {
      var sheet = results[0] || {};
      var baselines = (results[1] && results[1].baselines) || {};
      if (!sheet.rows || !sheet.rows.length) {
        commitLane(pid, version, "sheet", []);
        return;
      }
      var baselineOffset = baselines[pid] || 0;
      // Applies the Convergence sheet offset, unlike Studio's grid; saved trims
      // bake it in deliberately.
      var markers = [];
      sheet.rows.forEach(function (row) {
        var cell = row.cells && row.cells[pid];
        if (!cell || !cell.valid) return;
        var segs = parseClipSegmentsForCell(
          cell.value, baselineOffset, CLIPGEN_CONFIG.defaultDuration);
        segs.forEach(function (seg, idx) {
          markers.push({
            key: "sheet:" + row.rowNum + ":" + pid + ":" + idx,
            source: "sheet",
            start: seg.startSeconds + off,
            end: seg.startSeconds + seg.duration + off,
            label: row.observation || "",
            eventType: row.category || "uncategorized",
          });
        });
      });
      commitLane(pid, version, "sheet", markers);
    }).catch(function () {
      commitLane(pid, version, "sheet", []);
    });
  }

  // Merge gap mirroring Studio's #intakeClusterThreshold; clustering makes point
  // events grabbable by the edge.
  var SS_CLUSTER_SECONDS = 10;

  function loadScreenspaceMarkers(pid, version, offsets) {
    var off = offsetFor(offsets, pid, "screenspace");
    apiGet("../screenspace/api/events?excluded=false&participant=" + encodeURIComponent(pid))
      .then(function (data) {
        var events = ((data && data.events) || []).filter(function (ev) {
          // Boundaries are orientation scaffolding, not clip candidates —
          // same default as Studio's intake (intakeClusterSource).
          return !ev.navigational;
        });
        var clusters = window.ClipgenIntakeCluster.clusterIntakeEvents(
          events, SS_CLUSTER_SECONDS);
        var markers = clusters.map(function (cl) {
          var n = cl.events.length;
          var type = cl.event_type || cl.detector || "";
          return {
            // Keyed on the earliest event; a re-scan that adds an earlier one
            // orphans the trim.
            key: "screenspace:" + cl.events[0].id,
            source: "screenspace",
            start: cl.start + off,
            end: Math.max(cl.end, cl.start) + off,
            label: type + (n > 1 ? " · " + n + " events" : "")
              + (cl.region ? " · " + cl.region : ""),
            eventType: type,
            eventIds: cl.events.map(function (e) { return e.id; }),
          };
        });
        commitLane(pid, version, "screenspace", markers);
      })
      .catch(function () {
        commitLane(pid, version, "screenspace", []);
      });
  }

  // Transcript: only marked segments; a full transcript would carpet the lane.
  function loadTranscriptMarkers(pid, version, offsets) {
    var off = offsetFor(offsets, pid, "transcript");
    apiGet("../transcripts/api/transcript/" + encodeURIComponent(pid))
      .then(function (data) {
        var segments = (data && data.segments) || [];
        var markers = [];
        segments.forEach(function (seg) {
          var marks = seg.marks || [];
          if (!marks.length) return;
          marks.forEach(function (mark) {
            markers.push({
              // Defensive fallback only: server marks always carry an id, and
              // trimBadgeKey matches only ids.
              key: "transcript-mark:" + (mark.id || pid + ":" + seg.id),
              source: "transcript",
              start: seg.start + off,
              end: seg.end + off,
              label: mark.label || seg.text || "",
              eventType: mark.category || "bookmark",
            });
          });
        });
        commitLane(pid, version, "transcript", markers);
      })
      .catch(function () {
        commitLane(pid, version, "transcript", []);
      });
  }

  // ---- Lane toggles ----

  function syncSourcePills() {
    SOURCES.forEach(function (src) {
      var box = qs('.co-lane-check[data-source="' + src + '"] input');
      if (box) box.checked = !!state.sourceToggles[src];
    });
  }

  function persistUi() {
    apiPut("api/ui", {
      markerSources: state.sourceToggles,
      laneFolds: state.laneFolds,
      markerThumbnails: state.markerThumbnails,
      markerAudioScrub: state.markerAudioScrub,
      followPlayhead: state.followPlayhead,
    }).catch(function () {});
  }

  function rerenderLanes() {
    if (CO.updateTimelineHeight) CO.updateTimelineHeight();
    if (CO.renderTimeline) CO.renderTimeline();
  }

  function toggleSource(source) {
    if (!(source in state.sourceToggles)) return;
    state.sourceToggles[source] = !state.sourceToggles[source];
    syncSourcePills();
    persistUi();
    rerenderLanes();
  }

  function toggleAllSources() {
    // Any lane visible → hide all; all hidden → show all.
    var anyOn = SOURCES.some(function (src) { return state.sourceToggles[src]; });
    SOURCES.forEach(function (src) { state.sourceToggles[src] = !anyOn; });
    syncSourcePills();
    persistUi();
    rerenderLanes();
  }

  function initMarkerToggles() {
    // change, not click: label clicks already flip the box. blur(): a focused
    // checkbox swallows hotkeys.
    SOURCES.forEach(function (src) {
      var box = qs('.co-lane-check[data-source="' + src + '"] input');
      if (box) {
        box.addEventListener("change", function () {
          this.blur();
          toggleSource(src);
        });
      }
    });
    syncSourcePills();
  }

  CO.loadMarkers = loadMarkers;
  CO.initMarkerToggles = initMarkerToggles;
  CO.toggleSource = toggleSource;
  CO.toggleAllSources = toggleAllSources;
  CO.syncSourcePills = syncSourcePills;
  CO.persistLaneUi = persistUi;
})();
