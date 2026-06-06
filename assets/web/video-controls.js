/* Shared custom video-player speed controls for Screenspace and Transcripts.
 *
 * Both pages render their own <video> chrome (play / mute / speed buttons) and
 * keep their own VIDEO_SPEEDS list + button-label updater, but the speed-cycle
 * math and playback-rate application are identical. This module owns those two
 * shared bits.
 *
 * Public API on window.ClipgenVideoControls:
 *   nextSpeed(speeds, current)        — next entry in the cycle (wraps around).
 *   applyPlaybackRate(videoEl, rate)  — set the playback rate without pitch
 *                                       preservation. The time-stretch filter is
 *                                       CPU-heavy and causes visible judder at
 *                                       high speeds, so audio pitch rises instead.
 */
(function () {
  "use strict";

  function nextSpeed(speeds, current) {
    var idx = speeds.indexOf(current);
    return speeds[(idx + 1) % speeds.length];
  }

  function applyPlaybackRate(videoEl, rate) {
    if (!videoEl) return;
    videoEl.defaultPlaybackRate = rate;
    videoEl.playbackRate = rate;
    videoEl.preservesPitch = false;
    videoEl.mozPreservesPitch = false;
    videoEl.webkitPreservesPitch = false;
  }

  window.ClipgenVideoControls = {
    nextSpeed: nextSpeed,
    applyPlaybackRate: applyPlaybackRate,
  };
})();
