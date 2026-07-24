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
  // The single-track graph is built lazily on the first real user gesture on the
  // slider, never on hover: autoplay policy only resumes an AudioContext inside a
  // user-activation handler, and connecting a suspended context to an
  // already-playing element would cut its sound. Until then the video plays
  // natively; the first interaction activates gain seamlessly.
  //
  // MULTITRACK: a browser plays only the container's default audio track, so
  // independent per-track volume needs each track as its own media source. When
  // a (single-file) participant has >1 track and the page supplies trackAudioUrl,
  // the module mutes the <video> (visual only) and plays N hidden <audio>
  // elements — one per extracted track (/api/.../audio-track/…) — each through
  // its own GainNode into a shared master, kept time-aligned with the video
  // (play/pause/seek/rate + drift correction). The popover then shows one slider
  // per track. Any audio-element error bails back to the video's own track so
  // there's always sound. Multi-part participants keep the single master slider.
  //
  // The page drives mode via ctrl.setMuted() (mute intent) and ctrl.refresh()
  // (call after the track list for the current participant is known). Volume is
  // in-memory only and resets to 100% on reload.

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

  var DRIFT_TOLERANCE = 0.15; // seconds of audio<->video slip before a nudge

  function attachAudioPanel(opts) {
    opts = opts || {};
    var video = opts.video;
    var button = opts.button;
    if (!video || !button) return null;
    // Idempotent: a page may re-run its player init. createMediaElementSource
    // throws if called twice on one element, so reuse the existing controller
    // and just refresh its getters.
    if (video.__cgAudioPanel) {
      if (opts.getTracks) video.__cgAudioPanel.getTracks = opts.getTracks;
      if (opts.trackAudioUrl) video.__cgAudioPanel.trackAudioUrl = opts.trackAudioUrl;
      return video.__cgAudioPanel;
    }

    var ctrl = {
      getTracks: opts.getTracks || function () { return []; },
      trackAudioUrl: opts.trackAudioUrl || null,
    };
    video.__cgAudioPanel = ctrl;

    var muted = false;         // unified mute intent (page drives via setMuted)
    var multiActive = false;   // multitrack mixing engaged for this participant
    var lastSig = null;        // track-set signature so refresh() is idempotent

    // ---- Single-track path: lazy video -> gain -> destination (commit 1) ----
    var videoSrcNode = null;
    var videoGain = null;
    var videoGraphFailed = false;
    var singlePercent = 100;

    function ensureVideoGraph() {
      if (videoGain || videoGraphFailed) return;
      var ctx = audioContext();
      if (!ctx) { videoGraphFailed = true; return; }
      try {
        videoSrcNode = ctx.createMediaElementSource(video);
      } catch (e) {
        videoGraphFailed = true; // not routable -> native volume fallback
        return;
      }
      try {
        videoGain = ctx.createGain();
        videoGain.gain.value = singlePercent / 100;
        videoSrcNode.connect(videoGain);
        videoGain.connect(ctx.destination);
      } catch (e2) {
        // Element already tapped — restore a passthrough so audio isn't lost.
        videoGain = null;
        try { videoSrcNode.connect(ctx.destination); } catch (e3) {}
        videoGraphFailed = true;
      }
      if (ctx.state === "suspended" && ctx.resume) ctx.resume();
    }

    function applySingleGain(percent) {
      singlePercent = percent;
      if (videoGain) videoGain.gain.value = percent / 100;
      // Boost above 100% is unavailable on the native fallback path.
      else if (videoGraphFailed) video.volume = Math.min(1, percent / 100);
    }

    // ---- Multitrack path: N synced <audio> elements + per-track gain ----
    var masterGain = null;
    var trackNodes = [];       // [{ el, gain, index }]
    var trackPercents = [];    // parallel gain %, defaults 100

    function resumeCtx() {
      var ctx = audioContext();
      if (ctx && ctx.state === "suspended" && ctx.resume) ctx.resume();
    }

    function teardownMulti() {
      for (var i = 0; i < trackNodes.length; i++) {
        var t = trackNodes[i];
        try { t.el.pause(); } catch (e) {}
        try { t.el.removeAttribute("src"); t.el.load(); } catch (e2) {}
        try { t.src.disconnect(); } catch (e3) {}
        try { t.gain.disconnect(); } catch (e4) {}
        if (t.el.parentNode) t.el.parentNode.removeChild(t.el);
      }
      trackNodes = [];
      if (masterGain) { try { masterGain.disconnect(); } catch (e5) {} masterGain = null; }
      multiActive = false;
    }

    function bailToSingle() {
      // Something went wrong loading a track — never leave the user with a muted
      // video and no sound. Drop the mix and play the video's own default track.
      teardownMulti();
      lastSig = "single";
      video.muted = muted;
      rowsDirty = true;
      if (isOpen()) rebuildRows();
    }

    function enterMulti(tracks) {
      teardownMulti();
      var ctx = audioContext();
      if (!ctx) { enterSingle(); return; }
      video.muted = true; // video is visual only; its baked track stays silent
      try {
        masterGain = ctx.createGain();
        masterGain.gain.value = muted ? 0 : 1;
        masterGain.connect(ctx.destination);
      } catch (e) { enterSingle(); return; }
      trackPercents = [];
      for (var i = 0; i < tracks.length; i++) {
        var url = safeTrackUrl(tracks[i].index);
        if (!url) { teardownMulti(); enterSingle(); return; }
        var el = document.createElement("audio");
        el.preload = "auto";
        el.src = url;
        el.className = "cg-audiotrack";
        el.addEventListener("error", bailToSingle);
        try {
          var src = ctx.createMediaElementSource(el);
          var g = ctx.createGain();
          g.gain.value = 1;
          src.connect(g);
          g.connect(masterGain);
          trackNodes.push({ el: el, gain: g, src: src, index: tracks[i].index });
          trackPercents.push(100);
          document.body.appendChild(el);
          el.load();
        } catch (e2) { teardownMulti(); enterSingle(); return; }
      }
      multiActive = true;
      if (!video.paused) syncPlay();
    }

    function enterSingle() {
      teardownMulti();
      video.muted = muted;
      // The video->gain graph stays lazy (built on first slider interaction).
    }

    function setTrackGain(i, percent) {
      trackPercents[i] = percent;
      if (trackNodes[i] && trackNodes[i].gain) {
        trackNodes[i].gain.gain.value = percent / 100;
      }
    }

    // ---- Video <-> audio-element sync (attached once; no-op unless multi) ----
    function forEachTrack(fn) {
      for (var i = 0; i < trackNodes.length; i++) fn(trackNodes[i].el, i);
    }
    function syncTime() {
      forEachTrack(function (el) {
        try { el.currentTime = video.currentTime; } catch (e) {}
      });
    }
    function syncPlay() {
      resumeCtx();
      syncTime();
      forEachTrack(function (el) {
        var p = el.play();
        if (p && p.catch) p.catch(function () {});
      });
    }
    video.addEventListener("play", function () { if (multiActive) syncPlay(); });
    video.addEventListener("pause", function () {
      if (multiActive) forEachTrack(function (el) { el.pause(); });
    });
    video.addEventListener("ended", function () {
      if (multiActive) forEachTrack(function (el) { el.pause(); });
    });
    video.addEventListener("seeking", function () { if (multiActive) syncTime(); });
    video.addEventListener("seeked", function () { if (multiActive) syncTime(); });
    video.addEventListener("ratechange", function () {
      if (multiActive) forEachTrack(function (el) { el.playbackRate = video.playbackRate; });
    });
    video.addEventListener("timeupdate", function () {
      if (!multiActive || video.seeking) return;
      forEachTrack(function (el) {
        if (el.readyState < 2) return;
        if (Math.abs(el.currentTime - video.currentTime) > DRIFT_TOLERANCE) {
          try { el.currentTime = video.currentTime; } catch (e) {}
        }
        if (!video.paused && el.paused) { var p = el.play(); if (p && p.catch) p.catch(function () {}); }
      });
    });
    // Defend the invariant: in multitrack mode the video must never regain audio
    // (page code re-applies video.muted on load), else its baked track doubles.
    video.addEventListener("volumechange", function () {
      if (multiActive && !video.muted) video.muted = true;
    });

    // ---- Shared mute + track helpers ----
    function safeTracks() {
      var t;
      try { t = ctrl.getTracks() || []; } catch (e) { t = []; }
      return t;
    }
    function safeTrackUrl(index) {
      if (!ctrl.trackAudioUrl) return null;
      try { return ctrl.trackAudioUrl(index) || null; } catch (e) { return null; }
    }
    function applyMute() {
      if (multiActive) {
        if (masterGain) masterGain.gain.value = muted ? 0 : 1;
        video.muted = true;
      } else {
        video.muted = muted;
      }
    }
    ctrl.setMuted = function (m) { muted = !!m; applyMute(); };

    // Reconfigure single vs multi for the current participant. Idempotent via a
    // track-set signature so repeat calls (poll-driven track fetches) are cheap.
    ctrl.refresh = function () {
      var tracks = safeTracks();
      var url0 = tracks.length > 1 ? safeTrackUrl(tracks[0].index) : null;
      var wantMulti = tracks.length > 1 && !!url0;
      var sig = wantMulti
        ? "multi|" + url0 + "|" + tracks.map(function (t) { return t.index; }).join(",")
        : "single";
      if (sig === lastSig) return;
      lastSig = sig;
      if (wantMulti) enterMulti(tracks); else enterSingle();
      rowsDirty = true;
      if (isOpen()) rebuildRows();
    };

    // ---- Popover DOM (glass panel; rows rebuilt per mode) ----
    var popover = null;
    var rowsContainer = null;
    var caption = null;
    var rowsDirty = true;
    var closeTimer = null;

    function isOpen() { return popover && popover.style.display !== "none"; }
    function cancelClose() {
      if (closeTimer) { clearTimeout(closeTimer); closeTimer = null; }
    }
    function scheduleClose() {
      cancelClose();
      closeTimer = setTimeout(close, 120);
    }
    function close() { if (popover) popover.style.display = "none"; }

    // Build one label + slider + %-readout row. onGesture runs first (inside the
    // input's user activation) so it can arm/resume the audio graph.
    function makeRow(labelText, percent, onGesture, onValue) {
      var row = document.createElement("div");
      row.className = "audio-popover-row";
      var label = document.createElement("span");
      label.className = "audio-popover-label";
      label.textContent = labelText;
      var slider = document.createElement("input");
      slider.type = "range";
      slider.className = "audio-slider";
      slider.min = "0";
      slider.max = "200";
      slider.step = "1";
      slider.value = String(percent);
      slider.setAttribute("aria-label", labelText + " — 100% is source level");
      var value = document.createElement("span");
      value.className = "audio-popover-value";
      value.textContent = percent + "%";
      function commit() {
        onGesture();
        var v = parseInt(slider.value, 10);
        if (isNaN(v)) v = 100;
        value.textContent = v + "%";
        onValue(v);
      }
      slider.addEventListener("input", commit);
      slider.addEventListener("dblclick", function () {
        slider.value = "100";
        commit();
      });
      row.appendChild(label);
      row.appendChild(slider);
      row.appendChild(value);
      return row;
    }

    function rebuildRows() {
      rowsDirty = false;
      rowsContainer.innerHTML = "";
      var tracks = safeTracks();
      if (multiActive) {
        for (var i = 0; i < trackNodes.length; i++) {
          (function (idx) {
            var tr = tracks[idx] || {};
            var name = tr.label || ("Track " + (idx + 1));
            rowsContainer.appendChild(
              makeRow(name, trackPercents[idx] || 100, resumeCtx, function (v) {
                setTrackGain(idx, v);
              })
            );
          })(i);
        }
        caption.style.display = "none";
      } else {
        rowsContainer.appendChild(
          makeRow("Volume", singlePercent, ensureVideoGraph, applySingleGain)
        );
        // Detected-but-unmixable tracks (e.g. a multi-part participant): note them.
        if (tracks.length > 1) {
          var names = tracks.map(function (t, i) {
            return (t && t.label) || ("Track " + (i + 1));
          });
          caption.textContent = tracks.length + " tracks: " + names.join(", ");
          caption.style.display = "";
        } else {
          caption.style.display = "none";
        }
      }
    }

    function buildPopover() {
      popover = document.createElement("div");
      popover.className = "audio-popover";
      popover.setAttribute("role", "dialog");
      popover.setAttribute("aria-label", "Audio levels");
      rowsContainer = document.createElement("div");
      rowsContainer.className = "audio-popover-rows";
      popover.appendChild(rowsContainer);
      caption = document.createElement("div");
      caption.className = "audio-popover-caption";
      caption.style.display = "none";
      popover.appendChild(caption);
      popover.addEventListener("mouseenter", cancelClose);
      popover.addEventListener("mouseleave", scheduleClose);
      document.body.appendChild(popover);
    }

    function open() {
      cancelClose();
      if (!popover) buildPopover();
      if (rowsDirty) rebuildRows();
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
