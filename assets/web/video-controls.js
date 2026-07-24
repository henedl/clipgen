/* Shared custom video-player controls for Screenspace, Transcripts, and Composer.
 *
 * Each page renders its own <video> chrome (play / mute / speed buttons) and
 * keeps its own VIDEO_SPEEDS list + button-label updater, but the speed-cycle
 * math, playback-rate application, and the audio-level popover are identical.
 * This module owns those shared bits.
 *
 * Public API on window.ClipgenVideoControls:
 *   nextSpeed(speeds, current)        — next entry in the cycle (wraps around).
 *   applyPlaybackRate(videoEl, rate)  — set the playback rate without pitch
 *                                       preservation. The time-stretch filter is
 *                                       CPU-heavy and causes visible judder at
 *                                       high speeds, so audio pitch rises instead.
 *   attachAudioPanel({video, button, getTracks})
 *                                     — glassy hover popover on a speaker button
 *                                       with a 0–200% master volume slider (100%
 *                                       = source). Returns an idempotent
 *                                       controller; see below.
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

  // ---- Audio-level popover (glass hover panel + Web Audio master gain) ----
  //
  // HTML5 <video>.volume caps at 1.0, so >100% ("boost") is impossible with the
  // element alone. We route the element through a Web Audio GainNode
  // (createMediaElementSource -> gain -> destination) whose gain runs 0..2.0.
  //
  // The graph is built lazily on the first real user gesture on the slider,
  // never on hover: autoplay policy only resumes an AudioContext inside a
  // user-activation handler, and connecting a suspended context to an
  // already-playing element would cut its sound. Until then the video plays
  // natively; the first interaction activates gain seamlessly.
  //
  // Mute stays each page's own concern (video.muted) — it still silences the
  // graph because the element's muted flag gates its MediaElementSource output.
  // The slider and mute are independent (mute overrides; slider at 0 also
  // silences). Volume is in-memory only and resets to 100% on reload.

  var AUDIO_CTX = null; // one AudioContext shared by every attached player

  function audioContext() {
    if (AUDIO_CTX) return AUDIO_CTX;
    var Ctor = window.AudioContext || window.webkitAudioContext;
    if (!Ctor) return null;
    try {
      AUDIO_CTX = new Ctor();
    } catch (e) {
      AUDIO_CTX = null;
    }
    return AUDIO_CTX;
  }

  function attachAudioPanel(opts) {
    opts = opts || {};
    var video = opts.video;
    var button = opts.button;
    if (!video || !button) return null;
    // Idempotent: a page may re-run its player init. createMediaElementSource
    // throws if called twice on one element, so reuse the existing controller
    // and just refresh its track getter.
    if (video.__cgAudioPanel) {
      if (opts.getTracks) video.__cgAudioPanel.getTracks = opts.getTracks;
      return video.__cgAudioPanel;
    }

    var ctrl = {
      getTracks: opts.getTracks || function () { return []; },
      gainPercent: 100,
    };
    video.__cgAudioPanel = ctrl;

    var srcNode = null;
    var gainNode = null;
    var graphFailed = false;

    // Build the element -> gain -> destination graph. Must run inside a user
    // gesture so the context resumes and playback isn't interrupted.
    function ensureGraph() {
      if (gainNode || graphFailed) return;
      var ctx = audioContext();
      if (!ctx) { graphFailed = true; return; }
      try {
        srcNode = ctx.createMediaElementSource(video);
      } catch (e) {
        // Element not routable (already sourced, cross-origin) -> native volume.
        graphFailed = true;
        return;
      }
      try {
        gainNode = ctx.createGain();
        gainNode.gain.value = ctrl.gainPercent / 100;
        srcNode.connect(gainNode);
        gainNode.connect(ctx.destination);
      } catch (e2) {
        // Gain path failed but the element is already tapped — restore a direct
        // passthrough so audio isn't lost, and drop to native-volume control
        // (element.volume still attenuates a MediaElementSource output).
        gainNode = null;
        try { srcNode.connect(ctx.destination); } catch (e3) {}
        graphFailed = true;
      }
      if (ctx.state === "suspended" && ctx.resume) ctx.resume();
    }

    function applyGain(percent) {
      ctrl.gainPercent = percent;
      if (gainNode) {
        gainNode.gain.value = percent / 100;
      } else {
        // No gain node (unbuilt or failed): approximate with native volume,
        // capped at 1.0 (boost above 100% is unavailable on this path).
        video.volume = Math.min(1, percent / 100);
      }
    }

    // ---- Popover DOM (built once, lazily on first open) ----
    var popover = null;
    var slider = null;
    var valueLabel = null;
    var caption = null;
    var closeTimer = null;

    function cancelClose() {
      if (closeTimer) { clearTimeout(closeTimer); closeTimer = null; }
    }
    function scheduleClose() {
      cancelClose();
      closeTimer = setTimeout(close, 120);
    }
    function close() {
      if (popover) popover.style.display = "none";
    }

    function onSliderChange() {
      // The change itself is the user gesture — arm the graph synchronously so
      // the AudioContext resumes within the activation (covers pointer + arrows).
      ensureGraph();
      var v = parseInt(slider.value, 10);
      if (isNaN(v)) v = 100;
      valueLabel.textContent = v + "%";
      applyGain(v);
    }

    function buildPopover() {
      popover = document.createElement("div");
      popover.className = "audio-popover";
      popover.setAttribute("role", "dialog");
      popover.setAttribute("aria-label", "Audio level");

      var row = document.createElement("div");
      row.className = "audio-popover-row";

      var label = document.createElement("span");
      label.className = "audio-popover-label";
      label.textContent = "Volume";

      slider = document.createElement("input");
      slider.type = "range";
      slider.className = "audio-slider";
      slider.min = "0";
      slider.max = "200";
      slider.step = "1";
      slider.value = String(ctrl.gainPercent);
      slider.setAttribute("aria-label", "Volume — 100% is source level");

      valueLabel = document.createElement("span");
      valueLabel.className = "audio-popover-value";
      valueLabel.textContent = ctrl.gainPercent + "%";

      row.appendChild(label);
      row.appendChild(slider);
      row.appendChild(valueLabel);
      popover.appendChild(row);

      // Read-only detected-track caption (shown only when >1 track). The
      // per-track sliders themselves land in a follow-up commit; this row list
      // is already shaped to grow one row per track.
      caption = document.createElement("div");
      caption.className = "audio-popover-caption";
      caption.style.display = "none";
      popover.appendChild(caption);

      slider.addEventListener("input", onSliderChange);
      // Double-click resets to the source level.
      slider.addEventListener("dblclick", function () {
        ensureGraph();
        slider.value = "100";
        valueLabel.textContent = "100%";
        applyGain(100);
      });
      popover.addEventListener("mouseenter", cancelClose);
      popover.addEventListener("mouseleave", scheduleClose);

      document.body.appendChild(popover);
    }

    function refreshCaption() {
      if (!caption) return;
      var tracks = [];
      try { tracks = ctrl.getTracks() || []; } catch (e) { tracks = []; }
      if (tracks.length > 1) {
        var names = tracks.map(function (t, i) {
          return (t && t.label) || ("Track " + (i + 1));
        });
        caption.textContent = tracks.length + " audio tracks: " + names.join(", ");
        caption.style.display = "";
      } else {
        caption.textContent = "";
        caption.style.display = "none";
      }
    }

    function open() {
      cancelClose();
      if (!popover) buildPopover();
      slider.value = String(ctrl.gainPercent);
      valueLabel.textContent = ctrl.gainPercent + "%";
      refreshCaption();
      // Measure hidden, then anchor to the button (bottom-left, viewport-clamped).
      popover.style.visibility = "hidden";
      popover.style.display = "flex";
      if (typeof positionPopoverAnchored === "function") {
        positionPopoverAnchored(popover, button.getBoundingClientRect());
      }
      popover.style.visibility = "";
    }

    button.addEventListener("mouseenter", open);
    button.addEventListener("mouseleave", scheduleClose);

    return ctrl;
  }

  window.ClipgenVideoControls = {
    nextSpeed: nextSpeed,
    applyPlaybackRate: applyPlaybackRate,
    attachAudioPanel: attachAudioPanel,
  };
})();
