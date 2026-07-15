/* clipgen Transcripts video satellite — transcripts-video.js
 *
 * The custom <video> player (play/pause/mute/speed/CC/PiP/collapse), the
 * multi-video client-side source switching, the timeline canvas (ruler, mark
 * markers, friction heatmap band, playhead, hover hit-testing), seeking, and
 * playhead↔segment-list sync. Loaded after transcripts.js; reads the hub's
 * shared state + getMarkForSegment through window.ClipgenTranscripts (TS) and
 * publishes its player/timeline entry points back. The hub reaches the bits it
 * used to poke directly through cancelPendingSeek / clearTimelineMarkers /
 * hasTimelineHover. Plain utils.js globals (qs/el/formatTime/MARK_CATEGORIES/
 * getCanvasThemeColors/drawTimelineRuler/niceTimeInterval/getCSSVar/hexToRgba/
 * getStoredUIState/setStoredUIStateField/clipgenInstallPausedFrameOverlay) and
 * window.ClipgenVideoControls are reached via the scope chain.
 */
(function () {
  "use strict";

  var TS = window.ClipgenTranscripts;
  var state = TS.state;
  var getMarkForSegment = TS.getMarkForSegment;

  // Speed cycle for the custom video controls. Transcript-friendly steps so
  // users can slow review or skim faster without large jumps.
  var VIDEO_SPEEDS = [0.75, 1, 1.25, 1.5, 2];

  var _markerHitRects = [];
  var _playheadRaf = 0;
  var _timelineTooltipRaf = 0;
  var _lastTimelineHit = null;
  var _timelineResizeObs = null;

  // Transcribe-progress band: while the selected participant has a running
  // transcription task, the timeline fills left→right in sync with progress
  // (a fraction of the video's duration processed). The eased display value is
  // driven toward its polled target by a small RAF loop; renderTimeline() only
  // *draws* it (never mutates the target/RAF) so it can't re-enter the ease.
  var _txFillTarget = 0;      // 0..1 target progress (latest poll)
  var _txFillDisplay = 0;     // 0..1 eased value currently drawn
  var _txFillRaf = 0;         // RAF handle for the ease loop
  var _txFillActive = false;  // a running transcription task drives the band
  var _txFillPid = null;      // participant the display belongs to (reset on switch)

  function setIconClass(span, klass) {
    if (!span) return;
    span.className = "player-btn-icon " + klass;
  }

  function updatePlayerButtons() {
    var playBtn = qs("#videoPlayBtn");
    var muteBtn = qs("#videoMuteBtn");
    var speedBtn = qs("#videoSpeedBtn");
    var ccBtn = qs("#videoCcBtn");
    var pipBtn = qs("#videoPipBtn");
    if (playBtn) {
      var pIcon = playBtn.querySelector(".player-btn-icon");
      setIconClass(pIcon, state.videoPlaying ? "player-icon-pause" : "player-icon-play");
      playBtn.title = state.videoPlaying ? "Pause (Space)" : "Play/Pause (Space)";
    }
    if (muteBtn) {
      var mIcon = muteBtn.querySelector(".player-btn-icon");
      setIconClass(mIcon, state.videoMuted ? "player-icon-mute-off" : "player-icon-mute");
      muteBtn.classList.toggle("active", !state.videoMuted);
    }
    if (speedBtn) {
      speedBtn.textContent = state.videoPlaybackRate + "x";
      speedBtn.classList.toggle("active", state.videoPlaybackRate !== 1);
    }
    if (ccBtn) {
      ccBtn.classList.toggle("active", state.ccEnabled);
      ccBtn.setAttribute("aria-pressed", state.ccEnabled ? "true" : "false");
    }
    if (pipBtn) {
      pipBtn.classList.toggle("active", state.pipEnabled);
      pipBtn.setAttribute("aria-pressed", state.pipEnabled ? "true" : "false");
    }
    var collapseBtn = qs("#videoCollapseBtn");
    if (collapseBtn) {
      collapseBtn.title = state.videoCollapsed ? "Show video" : "Hide video";
      collapseBtn.setAttribute("aria-expanded", state.videoCollapsed ? "false" : "true");
    }
    var followBtn = qs("#videoFollowBtn");
    if (followBtn) {
      followBtn.classList.toggle("active", !!state.autoFollow);
      followBtn.setAttribute("aria-pressed", state.autoFollow ? "true" : "false");
      followBtn.title = state.autoFollow
        ? "Auto-scroll transcript with playback (on)"
        : "Auto-scroll transcript with playback (off)";
    }
  }

  function applyPlaybackRate() {
    window.ClipgenVideoControls.applyPlaybackRate(qs("#videoPlayer"), state.videoPlaybackRate);
  }

  // ---- Multi-video timeline (client-side source switching) ----
  // For a participant whose recording spans several files, p.timeline carries
  // [{filename, duration, cumulativeStart}]. The <video> plays one part at a
  // time; these helpers present a single GLOBAL timeline to the controls so the
  // playhead, labels, and segment sync all use global time. Single-video
  // participants have state.videoTimeline === null and take the original path.
  function _timelineTotal(tl) {
    if (!tl || !tl.length) return 0;
    var last = tl[tl.length - 1];
    return last.cumulativeStart + last.duration;
  }
  function _partForGlobal(tl, g) {
    for (var i = 0; i < tl.length; i++) {
      if (g >= tl[i].cumulativeStart && g < tl[i].cumulativeStart + tl[i].duration) {
        return i;
      }
    }
    return tl.length - 1;
  }
  function _partMediaUrl(i) {
    var url = "media/" + state.videoTimeline[i].filename;
    if (state.videoVersion != null) url += "?v=" + encodeURIComponent(state.videoVersion);
    return url;
  }
  function videoDisplayDuration() {
    var v = qs("#videoPlayer");
    if (state.videoTimeline) return _timelineTotal(state.videoTimeline);
    return v && isFinite(v.duration) ? v.duration : 0;
  }
  function videoGlobalTime() {
    var v = qs("#videoPlayer");
    if (!v) return 0;
    return (v.currentTime || 0) + (state.videoTimeline ? state.videoOffset : 0);
  }
  function _switchToPart(i, localTime, autoplay) {
    var v = qs("#videoPlayer");
    state.videoActivePart = i;
    state.videoOffset = state.videoTimeline[i].cumulativeStart;
    v.src = _partMediaUrl(i);
    var onMeta = function () {
      v.removeEventListener("loadedmetadata", onMeta);
      v.currentTime = localTime;
      if (autoplay) v.play();
    };
    v.addEventListener("loadedmetadata", onMeta);
  }

  function updateTimeLabel() {
    var v = qs("#videoPlayer");
    var label = qs("#videoTime");
    if (!v || !label) return;
    var dur = videoDisplayDuration();
    label.textContent = formatTime(videoGlobalTime()) + " / " + formatTime(dur);
  }

  function applyCaptionMode() {
    var v = qs("#videoPlayer");
    if (!v || !v.textTracks || !v.textTracks.length) return;
    v.textTracks[0].mode = state.ccEnabled ? "showing" : "hidden";
  }

  function sizeTimelineCanvas() {
    var wrap = qs("#timelineCanvasWrapper");
    var c1 = qs("#timelineCanvas");
    var c2 = qs("#playheadCanvas");
    if (!wrap || !c1 || !c2) return;
    var rect = wrap.getBoundingClientRect();
    if (rect.width === 0) return;
    var dpr = window.devicePixelRatio || 1;
    [c1, c2].forEach(function (c) {
      var cssH = c === c2 ? c.offsetHeight || 48 : c.offsetHeight || 48;
      var cssW = rect.width;
      c.width = Math.round(cssW * dpr);
      c.height = Math.round(cssH * dpr);
      var ctx = c.getContext("2d");
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    });
  }

  // Draw the smoothed friction density band across the timeline ruler.
  // Per-pixel averaging of overlapping segment scores (mirrors the Screenspace
  // amplitude graph's binning) gives a continuous band without a separate
  // smoothing constant; alpha scales with score. Reads the shared friction state
  // the agents satellite writes (state.frictionHeatmapEnabled / frictionBySegId).
  function _drawFrictionBand(ctx, timeToX, bandY, bandH, cssW) {
    if (!state.frictionHeatmapEnabled) return;
    if (!state.segments.length) return;
    var fcolor = getCSSVar("--color-friction", "#ea580c");
    if (fcolor.charAt(0) !== "#") fcolor = "#ea580c";
    var numBins = Math.max(1, Math.floor(cssW));
    var sums = new Array(numBins);
    var counts = new Array(numBins);
    for (var b = 0; b < numBins; b++) { sums[b] = 0; counts[b] = 0; }
    var any = false;
    for (var i = 0; i < state.segments.length; i++) {
      var seg = state.segments[i];
      var frow = state.frictionBySegId[seg.id];
      var sc = frow ? (frow.score || 0) : 0;
      if (sc <= 0) continue;
      any = true;
      var x0 = Math.max(0, Math.floor(timeToX(seg.start)));
      var x1 = Math.min(numBins - 1, Math.floor(timeToX(seg.end || seg.start)));
      if (x1 < x0) x1 = x0;
      for (var x = x0; x <= x1; x++) { sums[x] += sc; counts[x] += 1; }
    }
    if (!any) return;
    for (var px = 0; px < numBins; px++) {
      if (!counts[px]) continue;
      var v = sums[px] / counts[px];
      if (v <= 0) continue;
      ctx.fillStyle = hexToRgba(fcolor, Math.min(0.85, 0.15 + v * 0.7));
      ctx.fillRect(px, bandY, 1, bandH);
    }
  }

  // The selected participant's running-transcription progress (0..1), or null
  // when it has no running task. Progress is media-time processed / duration,
  // so it maps directly onto the timeline's x-axis.
  function _selectedTranscribeProgress() {
    var pid = state.selectedParticipant;
    if (!pid || !state.tasks) return null;
    var prog = null;
    for (var i = 0; i < state.tasks.length; i++) {
      var t = state.tasks[i];
      if (t.participant === pid && t.status === "running") {
        prog = Math.max(0, Math.min(1, t.progress || 0));
      }
    }
    return prog;
  }

  function _prefersReducedMotion() {
    return (
      window.matchMedia &&
      window.matchMedia("(prefers-reduced-motion: reduce)").matches
    );
  }

  // Controller: recompute the band's target from the latest tasks and start (or
  // stop) the eased fill. Called every poll by the hub, and on participant
  // switch. Self-correcting — it turns the band on/off purely from state.
  function updateTranscribeFill() {
    var target = _selectedTranscribeProgress();
    if (target === null) {
      if (!_txFillActive && _txFillPid === null) return;
      _txFillActive = false;
      _txFillPid = null;
      if (_txFillRaf) {
        cancelAnimationFrame(_txFillRaf);
        _txFillRaf = 0;
      }
      renderTimeline(); // clear the band
      return;
    }
    if (_txFillPid !== state.selectedParticipant) {
      // New participant → reveal its band from empty.
      _txFillPid = state.selectedParticipant;
      _txFillDisplay = 0;
    }
    _txFillActive = true;
    _txFillTarget = target;
    if (_prefersReducedMotion()) {
      _txFillDisplay = target;
      renderTimeline();
      return;
    }
    _startTranscribeEase();
  }

  function _startTranscribeEase() {
    if (_txFillRaf) return;
    _txFillRaf = requestAnimationFrame(_stepTranscribeEase);
  }

  function _stepTranscribeEase() {
    _txFillRaf = 0;
    if (!_txFillActive) return;
    var diff = _txFillTarget - _txFillDisplay;
    if (Math.abs(diff) < 0.002) {
      _txFillDisplay = _txFillTarget;
      renderTimeline();
      return;
    }
    _txFillDisplay += diff * 0.18;
    renderTimeline();
    _startTranscribeEase();
  }

  // Grid pitch (px) of the dot texture, and its peak opacity at the very bottom
  // of the timeline. The pattern fades to nothing toward the top (see the
  // per-row globalAlpha below), so it reads as a faint speckle concentrated
  // along the bottom edge. Tune these two to taste.
  var _TX_DOT_STEP = 6;
  var _TX_DOT_PEAK_ALPHA = 0.4;

  // Paint the faint "unfilled" dot texture across the *untranscribed* remainder
  // of the timeline (full height, right of the fill front). The transcribed
  // region is left untouched so it stays the plain surfaceAlt background — the
  // dots simply look wiped away left→right as transcription advances. The dots
  // fade out toward the top via a per-row alpha ramp (strongest at the bottom).
  // Pure draw: reads _txFillActive/_txFillDisplay set by the controller.
  function _drawTranscribeBand(ctx, cssW, cssH, progress, theme) {
    var fillW = Math.round(Math.max(0, Math.min(1, progress)) * cssW);
    if (fillW >= cssW) return; // fully swept — nothing left to pattern

    ctx.save();
    ctx.beginPath();
    ctx.rect(fillW, 0, cssW - fillW, cssH); // untranscribed remainder only
    ctx.clip();
    ctx.fillStyle = theme.textDim;
    var step = _TX_DOT_STEP;
    for (var y = step / 2; y < cssH; y += step) {
      // Vertical gradient: ~invisible at the top, strongest at the bottom.
      var f = y / cssH;
      ctx.globalAlpha = _TX_DOT_PEAK_ALPHA * f * f;
      for (var x = step / 2; x < cssW; x += step) {
        ctx.beginPath();
        ctx.arc(x, y, 1, 0, Math.PI * 2);
        ctx.fill();
      }
    }
    ctx.restore();
  }

  function renderTimeline() {
    var canvas = qs("#timelineCanvas");
    if (!canvas) return;
    var ctx = canvas.getContext("2d");
    var v = qs("#videoPlayer");
    var cssW = canvas.offsetWidth;
    var cssH = canvas.offsetHeight;
    ctx.clearRect(0, 0, cssW, cssH);

    var theme = getCanvasThemeColors();

    ctx.fillStyle = theme.surfaceAlt;
    ctx.fillRect(0, 0, cssW, cssH);

    // Transcribe-progress dot texture (full height, behind the ruler + marks).
    // Drawn before the dur<=0 early return so it also shows during the
    // pre-metadata window.
    if (_txFillActive) {
      _drawTranscribeBand(ctx, cssW, cssH, _txFillDisplay, theme);
    }

    var dur = videoDisplayDuration();
    if (dur <= 0) {
      _markerHitRects = [];
      ctx.fillStyle = theme.textDim;
      ctx.font = "11px -apple-system, sans-serif";
      ctx.textAlign = "center";
      var placeholder = _txFillActive
        ? "Transcribing… " + Math.round(_txFillDisplay * 100) + "%"
        : state.selectedParticipant
          ? "Loading…"
          : "No video";
      ctx.fillText(placeholder, cssW / 2, cssH / 2 + 4);
      ctx.textAlign = "start";
      renderPlayhead();
      return;
    }

    function timeToX(t) { return (t / dur) * cssW; }

    drawTimelineRuler(ctx, {
      visStart: 0,
      visEnd: dur,
      interval: niceTimeInterval(dur, { targetTicks: 8 }),
      timeToX: timeToX,
      colors: { border: theme.border, textDim: theme.textDim, fontMono: theme.fontMono },
      tickHeight: 6,
      labelY: 16,
      format: formatTime,
    });

    var markerY = 22;
    var markerH = cssH - markerY - 4;
    _markerHitRects = [];

    // Friction heatmap band (behind marks).
    _drawFrictionBand(ctx, timeToX, markerY, markerH, cssW);

    for (var i = 0; i < state.segments.length; i++) {
      var seg = state.segments[i];
      var mark = getMarkForSegment(seg);
      if (!mark) continue;
      var cat = MARK_CATEGORIES[mark.category] || MARK_CATEGORIES.bookmark || { color: "#0891b2" };
      var color = mark.color || cat.color;
      var startX = timeToX(seg.start);
      var endX = seg.end ? timeToX(seg.end) : startX + 2;
      var barX = Math.max(0, startX - 1);
      var barW = Math.max(2, Math.min(endX - startX, 6));
      ctx.fillStyle = color;
      ctx.fillRect(barX, markerY, barW, markerH);
      _markerHitRects.push({
        x1: barX - 3,
        x2: barX + barW + 3,
        y: markerY,
        h: markerH,
        segIndex: i,
      });
    }

    renderPlayhead();
  }

  function renderPlayhead() {
    var canvas = qs("#playheadCanvas");
    var v = qs("#videoPlayer");
    if (!canvas || !v) return;
    var cssW = canvas.offsetWidth;
    var cssH = canvas.offsetHeight;
    var ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, cssW, cssH);
    var dur = videoDisplayDuration();
    if (dur <= 0) return;
    var px = (videoGlobalTime() / dur) * cssW;
    var accent = getCanvasThemeColors().accent;
    ctx.strokeStyle = accent;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(px, 0);
    ctx.lineTo(px, cssH);
    ctx.stroke();
    ctx.fillStyle = accent;
    ctx.beginPath();
    ctx.moveTo(px - 4, 0);
    ctx.lineTo(px + 4, 0);
    ctx.lineTo(px, 5);
    ctx.closePath();
    ctx.fill();
  }

  function timelineXToTime(event) {
    var canvas = qs("#timelineCanvas");
    var v = qs("#videoPlayer");
    if (!canvas || !v) return null;
    var dur = videoDisplayDuration();
    if (dur <= 0) return null;
    var rect = canvas.getBoundingClientRect();
    var frac = (event.clientX - rect.left) / rect.width;
    if (frac < 0) frac = 0; else if (frac > 1) frac = 1;
    return frac * dur;
  }

  // ---- Transcript timeline canvas (marks, playhead, hover hit-testing) ----

  function hitTestTimeline(clientX, clientY) {
    var canvas = qs("#timelineCanvas");
    if (!canvas) return null;
    var rect = canvas.getBoundingClientRect();
    var mx = clientX - rect.left;
    var my = clientY - rect.top;
    for (var i = _markerHitRects.length - 1; i >= 0; i--) {
      var hr = _markerHitRects[i];
      if (mx >= hr.x1 && mx <= hr.x2 && my >= hr.y && my <= hr.y + hr.h) return hr;
    }
    return null;
  }

  function showTimelineTooltip(hit, clientX, clientY) {
    var tip = qs("#trTooltip");
    if (!tip) return;
    var seg = state.segments[hit.segIndex];
    if (!seg) return;
    var mark = getMarkForSegment(seg);
    var cat = (mark && MARK_CATEGORIES[mark.category]) || MARK_CATEGORIES.bookmark || { label: "Mark", color: "#888" };
    var snippet = (seg.text || "").trim().slice(0, 80);
    if ((seg.text || "").length > 80) snippet += "…";
    var extraCount = (seg.marks && seg.marks.length > 1) ? (seg.marks.length - 1) : 0;
    var label = mark && mark.label ? " · " + mark.label : "";
    tip.textContent = "";
    var catSpan = el("span", "tr-tooltip-cat", cat.label);
    // Set color via property API rather than string-interpolating into a
    // style attribute — mark.color comes from a stash/manifest file and a
    // crafted value (e.g. `red" onmouseover=...`) would otherwise break out.
    catSpan.style.color = (mark && mark.color) || cat.color || "";
    tip.appendChild(catSpan);
    tip.appendChild(document.createTextNode(formatTime(seg.start) + label));
    if (extraCount > 0) {
      tip.appendChild(document.createTextNode(" "));
      var extraSpan = el("span", "", "+" + extraCount + " more");
      extraSpan.style.opacity = ".7";
      tip.appendChild(extraSpan);
    }
    tip.appendChild(document.createElement("br"));
    tip.appendChild(document.createTextNode(snippet));
    tip.classList.remove("hidden");
    var tipRect = tip.getBoundingClientRect();
    var x = clientX + 12;
    var y = clientY - tipRect.height - 12;
    if (x + tipRect.width > window.innerWidth - 8) x = window.innerWidth - tipRect.width - 8;
    if (y < 8) y = clientY + 16;
    if (y + tipRect.height > window.innerHeight - 8) y = Math.max(8, window.innerHeight - tipRect.height - 8);
    tip.style.left = x + "px";
    tip.style.top = y + "px";
  }

  function hideTimelineTooltip() {
    // Yield if a friction (hot-segment) tooltip currently owns the shared
    // element, so a canvas mouseleave doesn't clobber it mid-display.
    if (state.frictionTooltipShown) return;
    var tip = qs("#trTooltip");
    if (tip) tip.classList.add("hidden");
  }

  // Whether the timeline canvas currently has a marker hover. The friction
  // satellite's _hideFrictionTooltip reads this (via TS.hasTimelineHover) to
  // yield the shared #trTooltip to the canvas hover.
  function hasTimelineHover() {
    return !!_lastTimelineHit;
  }

  // Reset the marker hit-test rects. Called by the hub's renderEmptyState
  // before it repaints an empty timeline.
  function clearTimelineMarkers() {
    _markerHitRects = [];
  }

  function onMarkerClick(hit) {
    var seg = state.segments[hit.segIndex];
    if (!seg) return;
    seekVideo(seg.start);
    if (!state.cachedSegmentRows) {
      state.cachedSegmentRows = qs("#segmentList").querySelectorAll(".segment-row");
    }
    var row = state.cachedSegmentRows[hit.segIndex];
    if (row) scrollToSegment(row);
  }

  function initVideoPlayer() {
    var video = qs("#videoPlayer");
    if (!video) return;

    // Restore CC preference. We restore via getStoredUIState rather than a
    // module-level constant so that switching browsers / clearing storage
    // gives the user the documented default (off).
    var stored = getStoredUIState("transcripts");
    state.ccEnabled = !!(stored && stored.ccEnabled);
    state.videoCollapsed = !!(stored && stored.videoCollapsed);
    // Auto-follow defaults on; only an explicit stored `false` disables it.
    state.autoFollow = !(stored && stored.autoFollow === false);
    var section = qs("#videoSection");
    if (section) section.classList.toggle("video-collapsed", state.videoCollapsed);

    qs("#videoPlayBtn").addEventListener("click", function () {
      if (video.paused) video.play();
      else video.pause();
    });
    qs("#videoMuteBtn").addEventListener("click", function () {
      state.videoMuted = !state.videoMuted;
      video.muted = state.videoMuted;
      updatePlayerButtons();
    });
    qs("#videoSpeedBtn").addEventListener("click", function () {
      state.videoPlaybackRate = window.ClipgenVideoControls.nextSpeed(VIDEO_SPEEDS, state.videoPlaybackRate);
      applyPlaybackRate();
      updatePlayerButtons();
    });
    qs("#videoCcBtn").addEventListener("click", function () {
      state.ccEnabled = !state.ccEnabled;
      applyCaptionMode();
      setStoredUIStateField("transcripts", "ccEnabled", state.ccEnabled);
      updatePlayerButtons();
    });
    qs("#videoPipBtn").addEventListener("click", function () {
      state.pipEnabled = !state.pipEnabled;
      // When the user disables PiP while the player is detached, drop it
      // back into flow immediately. _setPipActive is provided by initPipScroll
      // and undefined until then; guard so the button still toggles state
      // even if scroll wiring failed for some reason.
      if (!state.pipEnabled && state.pipActive && typeof _setPipActive === "function") {
        _setPipActive(false);
      }
      updatePlayerButtons();
    });
    qs("#videoCollapseBtn").addEventListener("click", function () {
      state.videoCollapsed = !state.videoCollapsed;
      var sec = qs("#videoSection");
      if (sec) sec.classList.toggle("video-collapsed", state.videoCollapsed);
      setStoredUIStateField("transcripts", "videoCollapsed", state.videoCollapsed);
      updatePlayerButtons();
    });
    var followBtn = qs("#videoFollowBtn");
    if (followBtn) {
      followBtn.addEventListener("click", function () {
        state.autoFollow = !state.autoFollow;
        // A fresh enable shouldn't stay paused from an earlier manual scroll.
        if (state.autoFollow) _autoFollowPausedUntil = 0;
        setStoredUIStateField("transcripts", "autoFollow", state.autoFollow);
        updatePlayerButtons();
      });
    }

    initAutoFollowScrollPause();
    initShortcutsOverlay();

    video.addEventListener("play", function () {
      state.videoPlaying = true;
      updatePlayerButtons();
    });
    video.addEventListener("pause", function () {
      state.videoPlaying = false;
      updatePlayerButtons();
    });
    video.addEventListener("ended", function () {
      state.videoPlaying = false;
      updatePlayerButtons();
    });
    video.addEventListener("loadedmetadata", function () {
      sizeTimelineCanvas();
      applyCaptionMode();
      applyPlaybackRate();
      updateTimeLabel();
      renderTimeline();
    });

    // The <track> can finish loading after the video metadata; re-apply the
    // caption mode once cues are parsed or some browsers render nothing.
    var track = qs("#subtitleTrack");
    if (track) {
      track.addEventListener("load", applyCaptionMode);
    }
    video.addEventListener("durationchange", function () {
      updateTimeLabel();
      renderTimeline();
    });
    video.addEventListener("timeupdate", function () {
      // Multi-video: hand off to the next part as playback nears the boundary so
      // continuous playback spans the whole recording.
      if (state.videoTimeline) {
        var tl = state.videoTimeline;
        var i = state.videoActivePart;
        if (i < tl.length - 1 && video.currentTime >= tl[i].duration - 0.05) {
          _switchToPart(i + 1, 0.001, !video.paused);
        }
      }
      updateTimeLabel();
      if (_playheadRaf) return;
      _playheadRaf = requestAnimationFrame(function () {
        _playheadRaf = 0;
        renderPlayhead();
      });
    });

    // Keep the paused frame visible across tab switches. See utils.js.
    clipgenInstallPausedFrameOverlay(video);

    updatePlayerButtons();
  }

  function initTimelineCanvas() {
    var canvas = qs("#timelineCanvas");
    if (!canvas) return;
    sizeTimelineCanvas();

    if (typeof ResizeObserver === "function") {
      _timelineResizeObs = new ResizeObserver(function () {
        sizeTimelineCanvas();
        renderTimeline();
      });
      _timelineResizeObs.observe(qs("#timelineCanvasWrapper"));
      window.addEventListener("pagehide", function () {
        if (_timelineResizeObs) { _timelineResizeObs.disconnect(); _timelineResizeObs = null; }
      });
    } else {
      window.addEventListener("resize", function () {
        sizeTimelineCanvas();
        renderTimeline();
      });
    }

    canvas.addEventListener("click", function (e) {
      var hit = hitTestTimeline(e.clientX, e.clientY);
      if (hit) {
        onMarkerClick(hit);
        return;
      }
      var t = timelineXToTime(e);
      if (t !== null) seekVideo(t);
    });

    canvas.addEventListener("mousemove", function (e) {
      if (_timelineTooltipRaf) return;
      var cx = e.clientX;
      var cy = e.clientY;
      _timelineTooltipRaf = requestAnimationFrame(function () {
        _timelineTooltipRaf = 0;
        var hit = hitTestTimeline(cx, cy);
        if (hit) {
          _lastTimelineHit = hit;
          showTimelineTooltip(hit, cx, cy);
          canvas.style.cursor = "pointer";
        } else if (_lastTimelineHit) {
          _lastTimelineHit = null;
          hideTimelineTooltip();
          canvas.style.cursor = "pointer";
        }
      });
    });
    canvas.addEventListener("mouseleave", function () {
      _lastTimelineHit = null;
      hideTimelineTooltip();
    });
  }

  // ---- PiP scroll behaviour ----

  // Hoisted so initVideoPlayer's PiP-toggle handler can drop the player back
  // into flow when the user disables PiP. Assigned in initPipScroll.
  var _setPipActive = null;

  // Chrome strip height — TopNav (48) + subheader (44) + pill bar (56). Mirrored
  // in transcripts.css `#trMain { padding-top: 148px }`; PiP compensation has
  // to add the videoSection height on top of this so layout doesn't jump.
  var TR_CHROME_TOP = 148;

  function initPipScroll() {
    var section = qs("#videoSection");
    // #trMain is the scroll container (pass 6 floating-nav scroll-under) so
    // scrolled content slides under the chrome strip. PiP triggers when the
    // user scrolls past a commit; only returning to the very top dismisses
    // it. Asymmetric thresholds avoid the bounce caused by switching the
    // video section to position:fixed mid-scroll.
    var scroller = qs("#trMain");
    if (!section || !scroller) return;

    var ENTER_THRESHOLD = 140;
    var scrollRaf = 0;

    function setPipActive(active) {
      if (active === state.pipActive) return;
      if (active) {
        // Reserve the section's natural height on top of the chrome inset so
        // transcript content does not jump up when the player detaches from
        // flow. We restore the scroll position on the next frame because the
        // browser may rebase scrollTop relative to the new content size when
        // padding is added.
        var h = Math.round(section.getBoundingClientRect().height);
        if (h > 0) scroller.style.paddingTop = TR_CHROME_TOP + h + "px";
        var keepTop = scroller.scrollTop;
        state.pipActive = true;
        section.classList.add("pip");
        requestAnimationFrame(function () {
          scroller.scrollTop = keepTop;
          sizeTimelineCanvas();
          renderTimeline();
        });
      } else {
        var keepTop2 = scroller.scrollTop;
        state.pipActive = false;
        section.classList.remove("pip");
        // Empty inline override falls back to the CSS default (148px).
        scroller.style.paddingTop = "";
        requestAnimationFrame(function () {
          scroller.scrollTop = keepTop2;
          sizeTimelineCanvas();
          renderTimeline();
        });
      }
    }
    _setPipActive = setPipActive;

    function evaluatePip() {
      if (!state.pipEnabled) return;
      var top = scroller.scrollTop;
      if (state.pipActive) {
        // Only release PiP once the user is back at the very top, so the user
        // makes a deliberate "scroll up" or "click PiP" gesture rather than
        // having the player drop out as soon as they ease back.
        if (top <= 0) setPipActive(false);
      } else {
        if (top > ENTER_THRESHOLD) setPipActive(true);
      }
    }

    scroller.addEventListener("scroll", function () {
      if (scrollRaf) return;
      scrollRaf = requestAnimationFrame(function () {
        scrollRaf = 0;
        evaluatePip();
      });
    }, { passive: true });

    section.addEventListener("click", function (e) {
      if (!state.pipActive) return;
      if (e.target.closest(".player-btn") || e.target.closest("#timelineCanvasWrapper")) return;
      scroller.scrollTo({ top: 0, behavior: "smooth" });
    });
  }

  // ---- Keyboard ----

  function _markElForIndex(idx) {
    var list = qs("#segmentList");
    var row = list ? list.querySelector('.segment-row[data-index="' + idx + '"]') : null;
    return row ? row.querySelector(".segment-mark") : null;
  }

  // Move the active segment by *delta*, seeking + scrolling to it. Establishes an
  // active segment at an edge when none is selected yet.
  function _moveActiveSegment(delta) {
    var n = state.segments.length;
    if (!n) return;
    var cur = state.activeSegmentIndex;
    var next = cur < 0 ? (delta > 0 ? 0 : n - 1) : cur + delta;
    if (next < 0) next = 0;
    if (next > n - 1) next = n - 1;
    setActiveSegment(next, { seek: true, follow: true, force: true });
  }

  // Jump to the next/previous segment carrying a mark.
  function _jumpToMarkedSegment(dir) {
    var n = state.segments.length;
    if (!n) return;
    var i = state.activeSegmentIndex < 0 ? (dir > 0 ? -1 : n) : state.activeSegmentIndex;
    for (var step = 0; step < n; step++) {
      i += dir;
      if (i < 0 || i >= n) return;
      var marks = state.segments[i].marks;
      if (marks && marks.length > 0) {
        setActiveSegment(i, { seek: true, follow: true, force: true });
        return;
      }
    }
  }

  // Mark the active segment. categoryKey null = toggle (create, or open the
  // popover if already marked); a category key sets/creates that category.
  function _markActiveSegment(categoryKey) {
    var idx = state.activeSegmentIndex;
    var seg = idx >= 0 ? state.segments[idx] : null;
    if (!seg) return;
    var existing = seg.marks && seg.marks.length > 0 ? seg.marks[0] : null;
    if (categoryKey) {
      if (existing) TS.updateMarkCategory(existing.id, categoryKey);
      else { state.lastMarkCategory = categoryKey; TS.toggleMark(seg.id); }
    } else if (existing) {
      var anchor = _markElForIndex(idx);
      if (anchor) TS.showMarkPopover(anchor, seg.id, existing);
    } else {
      TS.toggleMark(seg.id);
    }
  }

  function onTranscriptKeydown(e) {
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    var t = e.target;
    if (t && t.matches && t.matches("input, textarea, select, [contenteditable=true]")) return;
    if (state.editingTextEl) return;

    // The cheatsheet works regardless of player/segment state.
    if (e.key === "?") { e.preventDefault(); toggleShortcuts(); return; }
    if (e.key === "Escape" && _shortcutsOpen()) { e.preventDefault(); hideShortcuts(); return; }

    var modal = qs("#correctionsModal");
    if (modal && !modal.classList.contains("hidden")) return;
    var pop = qs("#markPopover");
    if (pop && !pop.classList.contains("hidden")) return;

    var key = e.key;
    if (e.code === "Space" || key === " ") {
      var v = qs("#videoPlayer");
      if (!v || !v.src) return;
      e.preventDefault();
      if (v.paused) v.play();
      else v.pause();
      return;
    }

    // Remaining shortcuts act on the segment list.
    if (!state.segments || !state.segments.length) return;
    if (key === "j" || key === "ArrowDown") { e.preventDefault(); _moveActiveSegment(1); return; }
    if (key === "k" || key === "ArrowUp") { e.preventDefault(); _moveActiveSegment(-1); return; }
    if (key === "n") { e.preventDefault(); _jumpToMarkedSegment(1); return; }
    if (key === "p") { e.preventDefault(); _jumpToMarkedSegment(-1); return; }
    if (key === "m") { e.preventDefault(); _markActiveSegment(null); return; }
    if (key === ",") { e.preventDefault(); seekVideo(Math.max(0, videoGlobalTime() - 5)); return; }
    if (key === ".") { e.preventDefault(); seekVideo(videoGlobalTime() + 5); return; }
    if (key >= "1" && key <= "9") {
      var catKeys = Object.keys(MARK_CATEGORIES);
      var ci = parseInt(key, 10) - 1;
      if (ci < catKeys.length) { e.preventDefault(); _markActiveSegment(catKeys[ci]); }
      return;
    }
  }

  function initPlayerKeyboard() {
    document.addEventListener("keydown", onTranscriptKeydown);
  }

  // ---- Keyboard shortcuts cheatsheet ----

  function _shortcutsOpen() {
    var el = qs("#trShortcuts");
    return !!(el && !el.classList.contains("hidden"));
  }
  function showShortcuts() {
    var el = qs("#trShortcuts");
    if (el) el.classList.remove("hidden");
  }
  function hideShortcuts() {
    var el = qs("#trShortcuts");
    if (el) el.classList.add("hidden");
  }
  function toggleShortcuts() {
    if (_shortcutsOpen()) hideShortcuts();
    else showShortcuts();
  }

  function initShortcutsOverlay() {
    var btn = qs("#shortcutsBtn");
    if (btn) btn.addEventListener("click", function () { toggleShortcuts(); });
    var closeBtn = qs("#shortcutsClose");
    if (closeBtn) closeBtn.addEventListener("click", hideShortcuts);
    var overlay = qs("#trShortcuts");
    if (overlay) {
      overlay.addEventListener("click", function (e) {
        if (e.target === overlay) hideShortcuts();
      });
    }
  }

  var _pendingSeekTime = null;
  var _seekRaf = 0;
  var _pendingSeekListener = null;

  // Cancel an in-flight/deferred seek. Called by the hub's selectParticipant
  // when switching participants so a pending seek can't land on the new video.
  function cancelPendingSeek() {
    var video = qs("#videoPlayer");
    _pendingSeekTime = null;
    cancelAnimationFrame(_seekRaf);
    _seekRaf = 0;
    if (_pendingSeekListener) {
      if (video) video.removeEventListener("loadedmetadata", _pendingSeekListener);
      _pendingSeekListener = null;
    }
  }

  function seekVideo(time) {
    // *time* is GLOBAL. For a multi-video participant, switch the <video> source
    // to the part that owns it and seek the local offset; single-video falls
    // straight through to the original local seek.
    if (state.videoTimeline) {
      var tl = state.videoTimeline;
      var g = time < 0 ? 0 : Math.min(time, _timelineTotal(tl));
      var i = _partForGlobal(tl, g);
      var local = g - tl[i].cumulativeStart;
      if (i !== state.videoActivePart) {
        _switchToPart(i, local, true);
      } else {
        _seekLocal(local);
      }
      return;
    }
    _seekLocal(time);
  }

  function _seekLocal(time) {
    var video = qs("#videoPlayer");
    if (!video || !video.src) return;

    // Remove any previous deferred-seek listener
    if (_pendingSeekListener) {
      video.removeEventListener("loadedmetadata", _pendingSeekListener);
      _pendingSeekListener = null;
    }

    // If metadata hasn't loaded yet, defer the seek
    if (video.readyState < 1) {
      _pendingSeekTime = time;
      _pendingSeekListener = function () {
        video.removeEventListener("loadedmetadata", _pendingSeekListener);
        _pendingSeekListener = null;
        var t = _pendingSeekTime;
        _pendingSeekTime = null;
        if (t !== null) seekVideo(t);
      };
      video.addEventListener("loadedmetadata", _pendingSeekListener);
      return;
    }

    // Coalesce rapid seeks into one per animation frame
    _pendingSeekTime = time;
    cancelAnimationFrame(_seekRaf);
    _seekRaf = requestAnimationFrame(function () {
      var t = _pendingSeekTime;
      _pendingSeekTime = null;
      _seekRaf = 0;
      if (t === null) return;
      video.currentTime = t;
      if (video.paused) video.play();
    });
  }

  // ---- Video sync ----

  var _syncRaf = 0;

  function initVideoSync() {
    var video = qs("#videoPlayer");
    var save = function () { persistVideoTime(videoGlobalTime()); };
    video.addEventListener("timeupdate", function () {
      save();
      if (_syncRaf) return;
      _syncRaf = requestAnimationFrame(function () {
        _syncRaf = 0;
        highlightActiveSegment();
      });
    });
    // Native scrubbing on a paused video doesn't always fire timeupdate.
    video.addEventListener("seeked", save);
  }

  function persistVideoTime(t) {
    if (!state.selectedParticipant || !isFinite(t)) return;
    var stored = getStoredUIState("transcripts");
    var map = (stored.videoTimeByParticipant && typeof stored.videoTimeByParticipant === "object")
      ? stored.videoTimeByParticipant : {};
    map[state.selectedParticipant] = t;
    setStoredUIStateField("transcripts", "videoTimeByParticipant", map);
  }

  // Move the .active highlight to *newIndex* and optionally seek the video to its
  // start. Shared by playhead sync (highlightActiveSegment) and keyboard nav.
  // opts.follow scrolls the row into view; opts.seek jumps the video there.
  function setActiveSegment(newIndex, opts) {
    opts = opts || {};
    if (newIndex === state.activeSegmentIndex && !opts.force) return;

    if (!state.cachedSegmentRows) {
      var list = qs("#segmentList");
      state.cachedSegmentRows = list ? list.querySelectorAll(".segment-row") : [];
    }
    var rows = state.cachedSegmentRows;

    if (state.activeSegmentIndex >= 0 && state.activeSegmentIndex < rows.length) {
      rows[state.activeSegmentIndex].classList.remove("active");
    }
    state.activeSegmentIndex = newIndex;
    if (newIndex >= 0 && newIndex < rows.length) {
      rows[newIndex].classList.add("active");
      if (opts.follow) scrollToSegment(rows[newIndex]);
    }
    if (opts.seek && newIndex >= 0 && state.segments[newIndex]) {
      seekVideo(state.segments[newIndex].start);
    }
  }

  function highlightActiveSegment() {
    var video = qs("#videoPlayer");
    if (!video || !video.src) return;
    var t = videoGlobalTime();

    // Binary search for active segment (sorted, non-overlapping)
    var lo = 0, hi = state.segments.length - 1, newIndex = -1;
    while (lo <= hi) {
      var mid = (lo + hi) >> 1;
      if (state.segments[mid].start <= t) { newIndex = mid; lo = mid + 1; }
      else { hi = mid - 1; }
    }
    if (newIndex >= 0 && t >= state.segments[newIndex].end) newIndex = -1;

    if (newIndex === state.activeSegmentIndex) return;

    // Chase the playhead when auto-follow is on (or in PiP, where the embedded
    // player is hidden so there's nothing else to read against). Skip while the
    // user is reading elsewhere — a manual scroll pauses follow for a few seconds.
    var follow = (state.autoFollow || state.pipActive) && !_autoFollowPaused();
    setActiveSegment(newIndex, { follow: follow });
  }

  // ---- Auto-follow scroll-pause ----
  // A user scroll on #trMain pauses playhead-following for a few seconds so it
  // doesn't yank the transcript away while they read. scrollToSegment marks its
  // own programmatic scrolls so they don't count as a manual scroll.
  var AUTO_FOLLOW_PAUSE_MS = 3000;
  var _autoFollowPausedUntil = 0;
  var _ignoreScrollUntil = 0;

  function _autoFollowPaused() {
    return Date.now() < _autoFollowPausedUntil;
  }

  function initAutoFollowScrollPause() {
    var scroller = qs("#trMain");
    if (!scroller) return;
    scroller.addEventListener("scroll", function () {
      if (Date.now() < _ignoreScrollUntil) return; // our own scrollToSegment
      _autoFollowPausedUntil = Date.now() + AUTO_FOLLOW_PAUSE_MS;
    }, { passive: true });
  }

  function scrollToSegment(row) {
    // Pass 6 floating-nav scroll-under: #trMain is the scroll container; the
    // top 148px of its viewport sits *under* the fixed chrome strip, so the
    // visible top edge is at scroller.top + TR_CHROME_TOP, not at scroller.top.
    var scroller = qs("#trMain");
    if (!scroller) return;
    var rowRect = row.getBoundingClientRect();
    var scRect = scroller.getBoundingClientRect();
    var rowTopInScroll = rowRect.top - scRect.top + scroller.scrollTop;
    var rowBottomInScroll = rowTopInScroll + rowRect.height;
    var visibleTop = scroller.scrollTop + TR_CHROME_TOP;
    var visibleBottom = scroller.scrollTop + scroller.clientHeight;

    if (rowTopInScroll < visibleTop + 40) {
      _ignoreScrollUntil = Date.now() + 120;
      scroller.scrollTop = rowTopInScroll - TR_CHROME_TOP - 40;
    } else if (rowBottomInScroll > visibleBottom - 40) {
      _ignoreScrollUntil = Date.now() + 120;
      scroller.scrollTop = rowBottomInScroll - scroller.clientHeight + 40;
    }
  }

  // ---- Published back to the hub ----
  // Boot wires the init*; selectParticipant uses _partForGlobal/_partMediaUrl/
  // applyCaptionMode/cancelPendingSeek; renderEmptyState uses clearTimelineMarkers;
  // the segment list, loadTranscript, search, and the agents panel use seekVideo/
  // renderTimeline/scrollToSegment; the friction tooltip uses hasTimelineHover;
  // the task poller + selectParticipant use updateTranscribeFill.
  TS.initVideoPlayer = initVideoPlayer;
  TS.initTimelineCanvas = initTimelineCanvas;
  TS.initPipScroll = initPipScroll;
  TS.initVideoSync = initVideoSync;
  TS.initPlayerKeyboard = initPlayerKeyboard;
  TS.renderTimeline = renderTimeline;
  TS.updateTranscribeFill = updateTranscribeFill;
  TS.seekVideo = seekVideo;
  TS.scrollToSegment = scrollToSegment;
  TS.applyCaptionMode = applyCaptionMode;
  TS._partForGlobal = _partForGlobal;
  TS._partMediaUrl = _partMediaUrl;
  TS.cancelPendingSeek = cancelPendingSeek;
  TS.clearTimelineMarkers = clearTimelineMarkers;
  TS.hasTimelineHover = hasTimelineHover;
})();
