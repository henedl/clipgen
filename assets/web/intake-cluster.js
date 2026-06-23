/* Shared intake clustering for Studio.
 *
 * Pure grouping helpers used by the Studio page (studio.js) and its sub-tabs
 * (convergence.js, metadata.js) to collapse raw Screenspace events and
 * Transcript marks into time-adjacent clusters. No DOM, no module state —
 * each function takes its data plus a threshold (in seconds) and returns a
 * fresh list of clusters.
 *
 * Loaded before studio.js so the consumers can read it on init.
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
      if (c.start === c.end) {
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
        };
      } else {
        cur.end = Math.max(cur.end, m.end);
        cur.marks.push(m);
        if (m.text) cur.text += " " + m.text;
        if (m.label && !cur.label) cur.label = m.label;
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
