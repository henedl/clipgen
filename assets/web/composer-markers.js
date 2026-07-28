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

  function fetchJson(path) {
    return fetch(path).then(function (r) { return r.json(); });
  }

  function loadMarkers(pid) {
    var version = ++_loadVersion;
    // Overview, not Studio: the convergence-offsets route is deliberately on the
    // overview blueprint (see agents/ARCHITECTURE.md) so its own satellites can
    // reach it page-relative. Pointing at ../studio/ 404s, and because the catch
    // below degrades to {} every lane silently rendered at offset 0.
    fetchJson("../overview/api/convergence/offsets")
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
    // Overlay persisted trims: the marker shows its trimmed span, keeping the
    // source span on origStart/origEnd for the tooltip and reset.
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

  // Sheet: one marker per timestamp pair in the participant's column, converted
  // through the baseline (wall-clock sheets) exactly like Studio/Convergence.
  function loadSheetMarkers(pid, version, offsets) {
    var off = offsetFor(offsets, pid, "sheet");
    Promise.all([
      fetchJson("../studio/api/sheet"),
      fetchJson("../studio/api/sheet/baseline"),
    ]).then(function (results) {
      var sheet = results[0] || {};
      var baselines = (results[1] && results[1].baselines) || {};
      if (!sheet.rows || !sheet.rows.length) {
        commitLane(pid, version, "sheet", []);
        return;
      }
      var baselineOffset = baselines[pid] || 0;
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

  // Merge gap for Screenspace events, mirroring Studio's intake default
  // (#intakeClusterThreshold). Clustering — via the shared
  // ClipgenIntakeCluster — turns bursts of point events into one grabbable
  // block (point clusters are padded ±5 s by the shared helper), so lane
  // spans are wide enough to trim; raw single-frame events were impossible
  // to grab by the edge.
  var SS_CLUSTER_SECONDS = 10;

  function loadScreenspaceMarkers(pid, version, offsets) {
    var off = offsetFor(offsets, pid, "screenspace");
    fetchJson("../screenspace/api/events?excluded=false&participant=" + encodeURIComponent(pid))
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
            // Keyed on the cluster's earliest event so trims survive reloads
            // while the event set is stable (clusterIntakeEvents sorts by time).
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

  // Transcript: marked segments become labeled markers; unmarked segments are
  // skipped (a full transcript would carpet the lane edge-to-edge).
  function loadTranscriptMarkers(pid, version, offsets) {
    var off = offsetFor(offsets, pid, "transcript");
    fetchJson("../transcripts/api/transcript/" + encodeURIComponent(pid))
      .then(function (data) {
        var segments = (data && data.segments) || [];
        var markers = [];
        segments.forEach(function (seg) {
          var marks = seg.marks || [];
          if (!marks.length) return;
          marks.forEach(function (mark) {
            markers.push({
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
    CO.apiSend("PUT", "api/ui", {
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
    // change (not click on the label) — a label click already flips the
    // checkbox natively; toggleSource then aligns state and re-syncs it.
    // blur() because a focused checkbox is a hotkeys.js typing target and
    // would swallow every shortcut until focus moved elsewhere.
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
