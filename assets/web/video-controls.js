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
  // HTML5 <video>.volume caps at 1.0, so ">100% boost" needs the element routed
  // through a Web Audio GainNode (createMediaElementSource -> gain -> destination)
  // running 0..2.0. That graph is built lazily on the first real gesture on the
  // slider, never on hover: autoplay policy only resumes an AudioContext inside a
  // user-activation handler, and connecting a suspended context to an
  // already-playing element would cut its sound.
  //
  // MULTITRACK: a browser plays only the container's default audio track, so
  // per-track volume needs each track as its own media source. Given a
  // single-file participant with >1 track and a page-supplied trackAudioUrl, the
  // <video> is muted (visual only) and N hidden <audio> elements play instead —
  // one per extracted track, each through its own GainNode into a shared master,
  // kept time-aligned with the video (play/pause/seek/rate + drift correction).
  // Any audio-element error bails back to the video's own track so there is always
  // sound; multi-part participants keep the single master slider.
  //
  // The page drives mode via ctrl.setMuted() and ctrl.refresh() (once the current
  // participant's track list is known). Volume is in-memory and resets on reload.

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

  // Audio<->video drift is corrected by trimming playbackRate (smooth, no seek),
  // never by reassigning currentTime mid-playback (that re-seeks and pops).
  var SYNC_DEADBAND = 0.05; // s of slip tolerated before nudging
  var SYNC_GAIN = 0.5;      // proportional correction gain
  var SYNC_MAX = 0.06;      // cap the rate trim at ±6% (inaudible tempo shift)
  var SYNC_HARD = 0.75;     // s — beyond this, a one-off hard resync (seek) is ok

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
    var mixRunning = false;    // per-track <audio> elements confirmed playing
    var lastSig = null;        // track-set signature so refresh() is idempotent
    var retryHandler = null;   // one-shot gesture listener re-arming a blocked mix
    // Bumped by teardownMulti (which every mode switch goes through). A
    // discarded track element can still deliver a queued "error" after the next
    // participant's mix is built, and that must not tear the new mix down —
    // each track's error handler captures the generation it was created in.
    var mixGeneration = 0;

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
      if (!ctx || ctx.state !== "suspended" || !ctx.resume) return;
      var p = ctx.resume();
      // resume() is async and can settle after the elements' play() promises.
      // onMixStarted bails while the context is still suspended, so re-run it
      // here once the graph is genuinely audible.
      if (p && p.then) {
        p.then(function () {
          if (multiActive && !mixRunning && !video.paused) onMixStarted();
        }, function () {});
      }
    }

    function teardownMulti() {
      disarmRetry();
      mixGeneration++;
      for (var i = 0; i < trackNodes.length; i++) {
        var t = trackNodes[i];
        try { t.el.pause(); } catch (e) {}
        if (t.onError) t.el.removeEventListener("error", t.onError);
        try { t.el.removeAttribute("src"); t.el.load(); } catch (e2) {}
        try { t.src.disconnect(); } catch (e3) {}
        try { t.gain.disconnect(); } catch (e4) {}
        if (t.el.parentNode) t.el.parentNode.removeChild(t.el);
      }
      trackNodes = [];
      if (masterGain) { try { masterGain.disconnect(); } catch (e5) {} masterGain = null; }
      multiActive = false;
      mixRunning = false;
    }

    function bailToSingle(gen) {
      // Something went wrong loading a track — never leave the user with a muted
      // video and no sound. Drop the mix and play the video's own default track.
      // Ignore a track that already belongs to a torn-down mix: its error can
      // land after the *next* participant's mix is up, and tearing that one down
      // would silently cost them per-track mixing until another switch.
      if (gen !== mixGeneration) return;
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
      // NB: don't mute the <video> yet — the tail decides based on play state so
      // an already-playing switch stays audible until the mix is confirmed.
      try {
        masterGain = ctx.createGain();
        masterGain.gain.value = muted ? 0 : 1;
        masterGain.connect(ctx.destination);
      } catch (e) { enterSingle(); return; }
      trackPercents = [];
      var gen = mixGeneration;
      for (var i = 0; i < tracks.length; i++) {
        var url = safeTrackUrl(tracks[i].index);
        if (!url) { teardownMulti(); enterSingle(); return; }
        var el = document.createElement("audio");
        el.preload = "auto";
        // Keep pitch stable while the sync logic trims playbackRate to converge.
        el.preservesPitch = true;
        el.mozPreservesPitch = true;
        el.webkitPreservesPitch = true;
        el.src = url;
        el.className = "cg-audiotrack";
        var onError = (function (g0) {
          return function () { bailToSingle(g0); };
        })(gen);
        el.addEventListener("error", onError);
        try {
          var src = ctx.createMediaElementSource(el);
          var g = ctx.createGain();
          g.gain.value = 1;
          src.connect(g);
          g.connect(masterGain);
          trackNodes.push({ el: el, gain: g, src: src, index: tracks[i].index, onError: onError });
          trackPercents.push(100);
          document.body.appendChild(el);
          el.load();
        } catch (e2) { teardownMulti(); enterSingle(); return; }
      }
      multiActive = true;
      mixRunning = false;
      if (video.paused) {
        // Paused: silent anyway. The next user-initiated play() runs playTracks
        // within that gesture, so the elements start reliably; mute now so the
        // baked track never leaks when play resumes.
        video.muted = true;
      } else {
        // Already playing (audio-info resolved after play): keep the video's own
        // track audible and only mute once the mix is confirmed running, so a
        // gesture-less play() rejection can't leave playback silent.
        playTracks();
      }
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
    // Separate media elements can't be sample-locked to the video, so align
    // currentTime only on discontinuities (play / seek) and otherwise converge
    // residual drift by trimming playbackRate. Reassigning currentTime during
    // playback re-seeks the element and pops the audio, so it's reserved for a
    // rare large slip.
    function forEachTrack(fn) {
      for (var i = 0; i < trackNodes.length; i++) fn(trackNodes[i].el, i);
    }
    function alignTime() {
      forEachTrack(function (el) {
        try { el.currentTime = video.currentTime; } catch (e) {}
      });
    }
    // Start (or restart) every track element, aligned to the video. play() can
    // reject when called outside a user gesture (mid-playback mode switch); in
    // that case we keep the video's own track audible and retry on the next
    // gesture, so playback is never left silent.
    function playTracks() {
      resumeCtx();
      var rate = video.playbackRate || 1;
      forEachTrack(function (el) {
        try { el.currentTime = video.currentTime; } catch (e) {}
        el.playbackRate = rate;
        var p = el.play();
        if (p && p.then) p.then(onMixStarted, onPlayBlocked);
        else onMixStarted(); // no promise (older browsers) → assume it started
      });
    }
    function onMixStarted() {
      if (!multiActive) return;
      // Element playback and AudioContext.resume() are gated independently, so
      // the tracks can "play" into a still-suspended context — routing the mix
      // nowhere. Muting the <video> at that point leaves *total* silence with
      // no retry armed, recoverable only by a manual pause→play. Treat it like
      // a blocked start instead; resumeCtx re-runs this once the graph is live.
      var ctx = audioContext();
      if (!ctx || ctx.state !== "running") { armRetry(); return; }
      disarmRetry();
      mixRunning = true;
      applyMute(); // now silence the <video> — the mix has taken over its audio
    }
    function onPlayBlocked() {
      // Autoplay policy blocked a gesture-less start. Leave the video audible and
      // finish the switch on the user's next interaction anywhere on the page.
      if (multiActive) armRetry();
    }
    function armRetry() {
      if (retryHandler) return;
      // A pointerdown anywhere re-attempts the mix within that user activation.
      // Keyboard users are covered too: any pause→play cycle re-runs playTracks
      // from the video's own "play" gesture (so no document keydown needed).
      retryHandler = function () {
        disarmRetry();
        if (multiActive && !video.paused) playTracks();
      };
      document.addEventListener("pointerdown", retryHandler, true);
    }
    function disarmRetry() {
      if (!retryHandler) return;
      document.removeEventListener("pointerdown", retryHandler, true);
      retryHandler = null;
    }
    function resyncTick() {
      // Only correct once the mix is confirmed running — before that the start /
      // retry path owns playback and the video's own track is intentionally live.
      if (!multiActive || !mixRunning || video.paused || video.seeking) return;
      var rate = video.playbackRate || 1;
      forEachTrack(function (el) {
        if (el.readyState < 3) return; // HAVE_FUTURE_DATA — don't fight the buffer
        if (el.paused) {
          var p = el.play();
          if (p && p.catch) p.catch(function () {});
          return;
        }
        var delta = el.currentTime - video.currentTime; // + => audio ahead
        var ad = delta < 0 ? -delta : delta;
        if (ad > SYNC_HARD) {
          // Way out of sync (long stall / tab throttled) — accept one seek pop.
          try { el.currentTime = video.currentTime; } catch (e) {}
          el.playbackRate = rate;
        } else if (ad > SYNC_DEADBAND) {
          // Proportional rate trim: converges smoothly and settles as drift → 0.
          var corr = delta * SYNC_GAIN;
          if (corr > SYNC_MAX) corr = SYNC_MAX;
          else if (corr < -SYNC_MAX) corr = -SYNC_MAX;
          el.playbackRate = rate * (1 - corr);
        } else {
          el.playbackRate = rate; // within deadband — run at the true rate
        }
      });
    }
    video.addEventListener("play", function () { if (multiActive) playTracks(); });
    video.addEventListener("pause", function () {
      if (multiActive) forEachTrack(function (el) { el.pause(); });
    });
    video.addEventListener("ended", function () {
      if (multiActive) forEachTrack(function (el) { el.pause(); });
    });
    video.addEventListener("seeking", function () { if (multiActive) alignTime(); });
    video.addEventListener("seeked", function () { if (multiActive) alignTime(); });
    video.addEventListener("ratechange", function () {
      if (multiActive) forEachTrack(function (el) { el.playbackRate = video.playbackRate; });
    });
    video.addEventListener("timeupdate", resyncTick);
    // Defend the invariant: once the mix is running the video must never regain
    // audio (page code re-applies video.muted on load), else its baked track
    // doubles. Before the mix starts the video stays audible on purpose, so this
    // is gated on mixRunning.
    video.addEventListener("volumechange", function () {
      if (multiActive && mixRunning && !video.muted) video.muted = true;
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
              // == null, not ||: 0% is a valid (fully attenuated) track.
              makeRow(name, trackPercents[idx] == null ? 100 : trackPercents[idx], resumeCtx, function (v) {
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
