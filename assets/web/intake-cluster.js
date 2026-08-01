/* Shared intake clustering.
 *
 * Pure grouping helpers that collapse raw Screenspace events and Transcript
 * marks into time-adjacent clusters. No DOM, no module state — each function
 * takes its data plus a threshold (in seconds) and returns a fresh list of
 * clusters. Consumers: Studio (studio.js + its sub-tabs convergence.js /
 * metadata.js), Composer, and the Transcripts "Clip Marked Lines" action, which
 * clusters the same way so identical marks yield identical spans on every page.
 *
 * Loaded before each page's hub script so consumers can read it on init.
 *
 * Public API on window.ClipgenIntakeCluster:
 *   clusterIntakeEvents(events, thresholdSec)    — group Screenspace events by
 *                                                  participant + event_type, merging
 *                                                  runs closer than thresholdSec.
 *   clusterTranscriptMarks(marks, thresholdSec)  — group Transcript marks by
 *                                                  participant, merging runs closer
 *                                                  than thresholdSec.
 */
(function () {
  "use strict";

  function clusterIntakeEvents(events, thresholdSec) {
    if (!events.length) return [];
    var sorted = events.slice().sort(function (a, b) {
      if (a.participant !== b.participant) return a.participant < b.participant ? -1 : 1;
      if (a.event_type !== b.event_type) return a.event_type < b.event_type ? -1 : 1;
      return a.time_in - b.time_in;
    });
    var clusters = [];
    var cur = null;
    for (var i = 0; i < sorted.length; i++) {
      var ev = sorted[i];
      if (
        !cur ||
        ev.participant !== cur.participant ||
        ev.event_type !== cur.event_type ||
        // Navigational (boundary) events render as individual point ticks, so
        // never merge them — a merged cluster would draw only its first tick and
        // hide the rest. Each boundary gets its own cluster.
        ev.navigational ||
        ev.time_in - cur.end > thresholdSec
      ) {
        if (cur) clusters.push(cur);
        cur = {
          participant: ev.participant,
          source_video: ev.source_video,
          start: ev.time_in,
          end: ev.time_out,
          event_type: ev.event_type,
          detector: ev.detector,
          region: ev.region,
          // Clusters group by participant + event_type, so a boundary
          // cluster's events are uniformly navigational. Carry the flag so
          // timelines can render them distinctly and exclude them from zones.
          navigational: !!ev.navigational,
          events: [ev],
          confidence_avg: ev.confidence,
        };
      } else {
        cur.end = Math.max(cur.end, ev.time_out);
        cur.events.push(ev);
        var sum = 0;
        for (var j = 0; j < cur.events.length; j++) sum += cur.events[j].confidence;
        cur.confidence_avg = sum / cur.events.length;
      }
    }
    if (cur) clusters.push(cur);
    for (var k = 0; k < clusters.length; k++) {
      var c = clusters[k];
      // Navigational (boundary) events are precise instants — leave them at the
      // real time so the density timeline, card ranges, and clip windows don't
      // sit ±5s off (Viewer and Convergence undo this padding the same way; the
      // clip window for a navigational point is set in screenspaceClusterToItem).
      if (!c.navigational && c.start === c.end) {
        c.start = Math.max(0, c.start - 5);
        c.end = c.end + 5;
      }
    }
    return clusters;
  }

  function clusterTranscriptMarks(marks, thresholdSec) {
    if (!marks.length) return [];
    var sorted = marks.slice().sort(function (a, b) {
      if (a.participant !== b.participant) return a.participant < b.participant ? -1 : 1;
      return a.start - b.start;
    });
    var clusters = [];
    var cur = null;
    for (var i = 0; i < sorted.length; i++) {
      var m = sorted[i];
      if (!cur || m.participant !== cur.participant || m.start - cur.end > thresholdSec) {
        if (cur) clusters.push(cur);
        cur = {
          participant: m.participant,
          start: m.start,
          end: m.end,
          marks: [m],
          category: m.category || "bookmark",
          label: m.label || "",
          text: m.text || "",
          severity: m.severity || "",
        };
      } else {
        cur.end = Math.max(cur.end, m.end);
        cur.marks.push(m);
        if (m.text) cur.text += " " + m.text;
        if (m.label && !cur.label) cur.label = m.label;
        // Hoist the most-severe severity across the cluster (lower rank = worse).
        if (m.severity) {
          var curRank = severityRank(cur.severity);
          var mRank = severityRank(m.severity);
          if (mRank != null && (curRank == null || mRank < curRank)) cur.severity = m.severity;
        }
      }
    }
    if (cur) clusters.push(cur);
    return clusters;
  }

  window.ClipgenIntakeCluster = {
    clusterIntakeEvents: clusterIntakeEvents,
    clusterTranscriptMarks: clusterTranscriptMarks,
  };
})();
