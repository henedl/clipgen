/* clipgen Timeline Viewer.
 *
 * Renders the artifact timeline + list + filmstrip overlay. Artifacts come
 * from `window.CLIPGEN_DATA` injected by the Python side — the same data
 * contract is used both in-app (Studio embeds this viewer) and in the
 * exported standalone viewer (artifacts inlined as base64), so we never
 * assume a live backend is reachable.
 *
 * Lazy thumbnails: clip cards request thumbnails as they scroll into view
 * via _thumbObserver / _thumbQueue, with results cached in _thumbCache for
 * the session. Filmstrip mode does the same with its own observer/queue.
 */
(function () {
  "use strict";

  var data = window.CLIPGEN_DATA || null;
  if (data) clipgenApplyConfig(data.config);

  var state = {
    artifacts: [],
    filtered: [],
    selectedId: null,
    duration: 0,
    listSort: null,
    expandedTracks: {},
  };

  var _thumbQueue = [];
  var _thumbActive = 0;
  var THUMB_CONCURRENCY = 3;
  var _thumbObserver = null;
  var _thumbCache = {};

  window.addEventListener("pagehide", function () {
    Object.keys(_thumbCache).forEach(function (id) {
      try { URL.revokeObjectURL(_thumbCache[id]); } catch (_) {}
      delete _thumbCache[id];
    });
  });

  var _filmstripEnabled = false;
  var _filmstripObserver = null;
  var _filmstripThumbQueue = [];
  var _filmstripThumbActive = 0;
  var FILMSTRIP_CONCURRENCY = 4;
  var FILMSTRIP_STORAGE_KEY = "clipgen-viewer-filmstrip";

  var _cardScrub = null; // { mediaEl, videoEl, raf, audioKey, audioUrl }

  // Opt-in: layer the shared card-scrubber's audio snippets + waveform overlay
  // onto the existing <video>-seek scrub. Audio is decoded in-browser from the
  // clip file (no server, no extra exported assets); degrades to silent if the
  // browser can't decode that codec.
  var _scrubAudioEnabled = false;
  var SCRUBAUDIO_STORAGE_KEY = "clipgen-viewer-scrubaudio";

  var _screenspaceVisible = true;
  var SCREENSPACE_STORAGE_KEY = "clipgen-viewer-screenspace";

  var _preview = null; // { id, videoEl, wrapEl } — currently previewed artifact
  var _hoverDebounce = null;
  var _seekRaf = 0;
  var _lastSeekProportion = null;

  var SORT_DEFAULT_DIR = {
    severity: "desc",
    chrono: "asc",
    duration: "desc",
    alpha: "asc",
  };

  // ---- Helpers ----

  function markerTypeClass(type) {
    var t = type || "clip";
    if (t === "transcript") return "transcript";
    if (t === "screen" || t === "gif") return t;
    return "clip";
  }

  function markerClasses(a) {
    var parts = ["artifact-marker", markerTypeClass(a.type)];
    var sev = (a.severity || "").trim();
    if (sev) {
      parts.push(severityClass(sev));
    }
    return parts.join(" ");
  }

  function sortedUniqueSeverities() {
    var seen = {};
    var labels = [];
    state.artifacts.forEach(function (a) {
      var s = (a.severity || "").trim();
      if (!s || seen[s]) return;
      seen[s] = true;
      labels.push(s);
    });
    labels.sort(function (a, b) {
      var na = severityRank(a);
      var nb = severityRank(b);
      if (na === null) na = 999;
      if (nb === null) nb = 999;
      if (na !== nb) return na - nb;
      return a.localeCompare(b);
    });
    return labels;
  }

  function applySeverityPill(pillEl, a) {
    if (!pillEl) return;
    var sev = (a.severity || "").trim();
    if (!sev) {
      pillEl.textContent = "";
      pillEl.classList.add("hidden");
      return;
    }
    pillEl.classList.remove("hidden");
    pillEl.textContent = sev;
    pillEl.className = "detail-badge detail-severity " + severityClass(sev);
  }

  // ---- Clip thumbnails ----

  // Off-DOM video element seeks to ~25% (clamped to 0.5–5s) and captures a
  // 320x180 JPEG poster. The 8s timeout guards against clips that never fire
  // 'seeked' (corrupt files, codec stalls); finish() is idempotent so timeout
  // and seek both safely call it. Blob URL is owned by _thumbCache.
  function generateClipThumbnail(mediaEl, artifact, callback) {
    var done = false;
    function finish() {
      if (done) return;
      done = true;
      clearTimeout(timer);
      video.onerror = null;
      video.onloadedmetadata = null;
      video.onseeked = null;
      video.src = "";
      video.load();
      callback();
    }

    var video = document.createElement("video");
    video.preload = "metadata";
    video.muted = true;
    video.playsInline = true;
    video.src = artifact.file;

    var timer = setTimeout(finish, 8000);

    video.onerror = function () {
      mediaEl.classList.remove("thumb-pending");
      finish();
    };

    video.onloadedmetadata = function () {
      var dur = video.duration;
      if (!dur || !isFinite(dur)) { finish(); return; }
      var seek = Math.min(Math.max(dur * 0.25, 0.5), 5, dur - 0.01);
      video.currentTime = Math.max(0, seek);
    };

    video.onseeked = function () {
      try {
        var canvas = document.createElement("canvas");
        canvas.width = 320;
        canvas.height = 180;
        var ctx = canvas.getContext("2d");
        ctx.drawImage(video, 0, 0, 320, 180);
        canvas.toBlob(function (blob) {
          if (!blob) { mediaEl.classList.remove("thumb-pending"); finish(); return; }
          var url = URL.createObjectURL(blob);
          if (_thumbCache[artifact.id]) {
            try { URL.revokeObjectURL(_thumbCache[artifact.id]); } catch (_) {}
          }
          _thumbCache[artifact.id] = url;
          var img = document.createElement("img");
          img.src = url;
          img.alt = artifact.description || "";
          mediaEl.classList.remove("thumb-pending");
          mediaEl.classList.add("thumb-loaded");
          mediaEl.appendChild(img);
          finish();
        }, "image/jpeg", 0.7);
      } catch (_) {
        mediaEl.classList.remove("thumb-pending");
        finish();
      }
    };
  }

  function processThumbQueue() {
    while (_thumbActive < THUMB_CONCURRENCY && _thumbQueue.length) {
      _thumbActive++;
      var item = _thumbQueue.shift();
      generateClipThumbnail(item.mediaEl, item.artifact, function () {
        _thumbActive--;
        processThumbQueue();
      });
    }
  }

  // ---- Sidebar video scrub ----

  function cardScrubEnter(mediaEl, ev) {
    if (!mediaEl) return;
    var card = mediaEl.closest(".artifact-card");
    if (!card) return;
    var id = card.dataset.id;
    var artifact = findArtifact(id);
    if (!artifact || artifact.type !== "clip") return;

    cardScrubLeave();

    var vid = document.createElement("video");
    vid.preload = "auto";
    vid.muted = true;
    vid.playsInline = true;
    vid.src = artifact.file;
    vid.className = "card-scrub-video";

    var img = mediaEl.querySelector("img");
    if (img) img.style.display = "none";
    mediaEl.appendChild(vid);
    mediaEl.classList.add("scrub-active");

    if (!mediaEl.querySelector(".scrub-progress")) {
      var bar = el("div", "scrub-progress");
      bar.appendChild(el("div", "scrub-progress-fill"));
      mediaEl.appendChild(bar);
    }

    _cardScrub = { mediaEl: mediaEl, videoEl: vid, raf: 0, audioKey: id, audioUrl: artifact.file };
    if (_scrubAudioEnabled && window.clipgenCardScrubber) {
      window.clipgenCardScrubber.loadAudioBuffer(artifact.file, id);
    }

    var rect = mediaEl.getBoundingClientRect();
    var frac = Math.max(0, Math.min(1, (ev.clientX - rect.left) / rect.width));
    vid.onloadedmetadata = function () {
      if (!_cardScrub || _cardScrub.videoEl !== vid) return;
      var dur = vid.duration;
      if (!dur || !isFinite(dur)) return;
      vid.currentTime = dur * frac;
      var fill = mediaEl.querySelector(".scrub-progress-fill");
      if (fill) fill.style.width = (frac * 100).toFixed(1) + "%";
    };
  }

  function cardScrubMove(mediaEl, ev) {
    if (!_cardScrub) return;
    if (!mediaEl || mediaEl !== _cardScrub.mediaEl) return;
    if (!mediaEl.isConnected) { cardScrubLeave(); return; }
    var vid = _cardScrub.videoEl;
    if (!vid.duration || !isFinite(vid.duration)) return;
    if (vid.readyState < 1) return;

    var clientX = ev.clientX;
    if (_cardScrub.raf) return;
    _cardScrub.raf = requestAnimationFrame(function () {
      if (!_cardScrub) return;
      _cardScrub.raf = 0;
      var rect = mediaEl.getBoundingClientRect();
      var frac = Math.max(0, Math.min(1, (clientX - rect.left) / rect.width));
      vid.currentTime = vid.duration * frac;

      var fill = mediaEl.querySelector(".scrub-progress-fill");
      if (fill) fill.style.width = (frac * 100).toFixed(1) + "%";

      if (_scrubAudioEnabled && window.clipgenCardScrubber && _cardScrub) {
        var cs = window.clipgenCardScrubber;
        cs.audioScrubAt(_cardScrub.audioKey, _cardScrub.audioUrl, vid.duration * frac);
        var wf = cs.extractWaveform(_cardScrub.audioKey);
        if (wf) {
          var canvas = cs.getOrCreateWaveformCanvas(mediaEl);
          if (canvas) cs.drawWaveform(canvas, wf, frac);
        }
      }
    });
  }

  function cardScrubLeave() {
    if (!_cardScrub) return;
    var mediaEl = _cardScrub.mediaEl;
    var vid = _cardScrub.videoEl;

    if (_cardScrub.raf) cancelAnimationFrame(_cardScrub.raf);

    try { vid.src = ""; vid.load(); } catch (_) {}
    if (vid.parentNode) vid.parentNode.removeChild(vid);

    mediaEl.classList.remove("scrub-active");
    var img = mediaEl.querySelector("img");
    if (img) img.style.display = "";

    var fill = mediaEl.querySelector(".scrub-progress-fill");
    if (fill) fill.style.width = "0%";

    if (window.clipgenCardScrubber) {
      window.clipgenCardScrubber.audioScrubStop();
      window.clipgenCardScrubber.clearWaveform(mediaEl);
    }

    _cardScrub = null;
  }

  // Lazy clip thumbnails. Called on initial render and again after any list
  // rebuild — we tear down the previous observer first so old card elements
  // (now detached) don't keep firing intersection callbacks. Each card is
  // either served from `_thumbCache` immediately or queued for generation;
  // unobserve fires once the card has been seen so we don't re-enqueue.
  function initClipThumbnails() {
    if (_thumbObserver) { _thumbObserver.disconnect(); _thumbObserver = null; }
    _thumbQueue = [];
    _thumbActive = 0;

    var cards = qsa("#artifactList .artifact-card");
    if (!cards.length) return;

    function enqueueCard(card) {
      var id = card.dataset.id;
      var a = findArtifact(id);
      if (!a || a.type !== "clip") return;
      var media = card.querySelector(".artifact-media");
      if (!media || media.querySelector("img")) return;
      if (_thumbCache[a.id]) {
        var img = document.createElement("img");
        img.src = _thumbCache[a.id];
        img.alt = a.description || "";
        media.classList.remove("thumb-pending");
        media.classList.add("thumb-loaded");
        media.appendChild(img);
        return;
      }
      _thumbQueue.push({ mediaEl: media, artifact: a });
      processThumbQueue();
    }

    if (typeof IntersectionObserver !== "undefined") {
      _thumbObserver = new IntersectionObserver(function (entries) {
        entries.forEach(function (entry) {
          if (!entry.isIntersecting) return;
          _thumbObserver.unobserve(entry.target);
          enqueueCard(entry.target);
        });
      }, { root: qs("#sidebar"), rootMargin: "200px 0px" });
      cards.forEach(function (card) { _thumbObserver.observe(card); });
    } else {
      cards.forEach(enqueueCard);
    }
  }

  // ---- Filmstrip mode ----

  function initFilmstripToggle() {
    var defaultOn = data && data.meta && data.meta.filmstripEnabled;
    var stored = null;
    try { stored = window.localStorage.getItem(FILMSTRIP_STORAGE_KEY); } catch (_) {}
    if (stored === "true") _filmstripEnabled = true;
    else if (stored === "false") _filmstripEnabled = false;
    else _filmstripEnabled = !!defaultOn;

    var btn = qs("#filmstripToggle");
    if (btn) {
      btn.addEventListener("click", toggleFilmstrip);
      updateFilmstripButton();
    }
  }

  function updateFilmstripButton() {
    var btn = qs("#filmstripToggle");
    if (btn) btn.setAttribute("aria-pressed", _filmstripEnabled ? "true" : "false");
  }

  function toggleFilmstrip() {
    _filmstripEnabled = !_filmstripEnabled;
    try { window.localStorage.setItem(FILMSTRIP_STORAGE_KEY, _filmstripEnabled ? "true" : "false"); } catch (_) {}
    updateFilmstripButton();
    if (_filmstripEnabled) applyFilmstripMode();
    else removeFilmstripMode();
  }

  function applyFilmstripMode() {
    if (_filmstripObserver) { _filmstripObserver.disconnect(); _filmstripObserver = null; }
    _filmstripThumbQueue = [];
    _filmstripThumbActive = 0;

    var markers = qsa(".artifact-marker");
    if (!markers.length) return;

    var needsObserver = [];

    markers.forEach(function (m) {
      var id = m.dataset.id;
      var a = findArtifact(id);
      if (!a) return;

      var hasSev = !!(a.severity || "").trim();

      if (a.type === "screen" || a.type === "gif") {
        m.style.backgroundImage = "url(" + encodeURI(a.file) + ")";
        m.classList.add("filmstrip-thumb");
        if (hasSev) m.classList.add("filmstrip-sev-border");
      } else if (a.type === "clip") {
        if (_thumbCache[a.id]) {
          m.style.backgroundImage = "url(" + _thumbCache[a.id] + ")";
          m.classList.add("filmstrip-thumb");
          if (hasSev) m.classList.add("filmstrip-sev-border");
        } else {
          m.classList.add("filmstrip-loading");
          if (hasSev) m.classList.add("filmstrip-sev-border");
          needsObserver.push({ el: m, artifact: a });
        }
      }
    });

    if (!needsObserver.length) return;

    function enqueueMarker(item) {
      _filmstripThumbQueue.push(item);
      processFilmstripThumbQueue();
    }

    if (typeof IntersectionObserver !== "undefined") {
      _filmstripObserver = new IntersectionObserver(function (entries) {
        entries.forEach(function (entry) {
          if (!entry.isIntersecting) return;
          _filmstripObserver.unobserve(entry.target);
          var id = entry.target.dataset.id;
          var a = findArtifact(id);
          if (a) enqueueMarker({ el: entry.target, artifact: a });
        });
      }, { rootMargin: "200px 0px" });
      needsObserver.forEach(function (item) { _filmstripObserver.observe(item.el); });
    } else {
      needsObserver.forEach(enqueueMarker);
    }
  }

  function generateFilmstripThumb(markerEl, artifact, callback) {
    var STRIP_HEIGHT = 88;
    var MIN_FRAMES = 2;
    var MAX_FRAMES = 20;

    var done = false;
    var perFrameTimer = null;
    function finish() {
      if (done) return;
      done = true;
      clearTimeout(timer);
      clearTimeout(perFrameTimer);
      video.onerror = null;
      video.onloadedmetadata = null;
      video.onseeked = null;
      callback();
      try { video.src = ""; video.load(); } catch (_) {}
    }

    var video = document.createElement("video");
    video.preload = "metadata";
    video.muted = true;
    video.playsInline = true;
    video.src = artifact.file;

    var timer = setTimeout(finish, 20000);

    video.onerror = function () {
      markerEl.classList.remove("filmstrip-loading");
      finish();
    };

    video.onloadedmetadata = function () {
      var dur = video.duration;
      if (!dur || !isFinite(dur)) { markerEl.classList.remove("filmstrip-loading"); finish(); return; }

      var clipStart = artifact.start || 0;
      var clipEnd = artifact.end || artifact.start || dur;
      var clipDur = clipEnd - clipStart;
      if (clipDur <= 0) clipDur = dur;

      var aspect = (video.videoWidth / video.videoHeight) || (16 / 9);
      var thumbW = Math.round(STRIP_HEIGHT * aspect);
      var markerH = markerEl.offsetHeight || STRIP_HEIGHT;
      var markerW = markerEl.offsetWidth || thumbW;
      var displayThumbW = thumbW * markerH / STRIP_HEIGHT;
      var numFrames = Math.min(MAX_FRAMES, Math.max(MIN_FRAMES, Math.ceil(markerW / displayThumbW)));

      var canvas = document.createElement("canvas");
      canvas.width = thumbW * numFrames;
      canvas.height = STRIP_HEIGHT;
      var ctx = canvas.getContext("2d");

      var seekTimes = [];
      for (var i = 0; i < numFrames; i++) {
        var t = clipStart + (clipDur * (i + 0.5)) / numFrames;
        seekTimes.push(Math.min(Math.max(0, t), dur - 0.01));
      }

      var frameIndex = 0;
      var seekGen = 0;

      function captureAndAdvance() {
        clearTimeout(perFrameTimer);
        try { ctx.drawImage(video, frameIndex * thumbW, 0, thumbW, STRIP_HEIGHT); } catch (_) {}
        frameIndex++;
        if (frameIndex < seekTimes.length) {
          seekToFrame(frameIndex);
        } else {
          video.onseeked = null;
          canvas.toBlob(function (blob) {
            if (!blob) { markerEl.classList.remove("filmstrip-loading"); finish(); return; }
            var url = URL.createObjectURL(blob);
            if (_thumbCache[artifact.id]) {
              try { URL.revokeObjectURL(_thumbCache[artifact.id]); } catch (_) {}
            }
            _thumbCache[artifact.id] = url;
            if (_filmstripEnabled && markerEl.classList.contains("filmstrip-loading")) {
              markerEl.style.backgroundImage = "url(" + url + ")";
              markerEl.classList.remove("filmstrip-loading");
              markerEl.classList.add("filmstrip-thumb");
            }
            finish();
          }, "image/jpeg", 0.7);
        }
      }

      function seekToFrame(idx) {
        var gen = ++seekGen;
        video.onseeked = function () {
          if (gen !== seekGen) return;
          captureAndAdvance();
        };
        perFrameTimer = setTimeout(function () {
          if (gen !== seekGen) return;
          captureAndAdvance();
        }, 800);
        video.currentTime = seekTimes[idx];
      }

      seekToFrame(0);
    };
  }

  function processFilmstripThumbQueue() {
    while (_filmstripThumbActive < FILMSTRIP_CONCURRENCY && _filmstripThumbQueue.length) {
      var item = _filmstripThumbQueue.shift();
      if (_thumbCache[item.artifact.id]) {
        if (_filmstripEnabled && item.el.classList.contains("filmstrip-loading")) {
          item.el.style.backgroundImage = "url(" + _thumbCache[item.artifact.id] + ")";
          item.el.classList.remove("filmstrip-loading");
          item.el.classList.add("filmstrip-thumb");
        }
        continue;
      }
      _filmstripThumbActive++;
      try {
        generateFilmstripThumb(item.el, item.artifact, function () {
          _filmstripThumbActive--;
          processFilmstripThumbQueue();
        });
      } catch (_) {
        item.el.classList.remove("filmstrip-loading");
        _filmstripThumbActive--;
      }
    }
  }

  function removeFilmstripMode() {
    if (_filmstripObserver) { _filmstripObserver.disconnect(); _filmstripObserver = null; }
    _filmstripThumbQueue = [];
    _filmstripThumbActive = 0;

    qsa(".artifact-marker.filmstrip-thumb, .artifact-marker.filmstrip-loading").forEach(function (m) {
      m.classList.remove("filmstrip-thumb", "filmstrip-sev-border", "filmstrip-loading");
      m.style.backgroundImage = "";
    });
  }

  // ---- Screenspace toggle ----

  function initScreenspaceToggle() {
    var hasEvents = data && data.screenspaceEvents && data.screenspaceEvents.length > 0;
    var btn = qs("#screenspaceToggle");
    if (!btn) return;
    if (!hasEvents) {
      btn.style.display = "none";
      return;
    }
    var stored = null;
    try { stored = window.localStorage.getItem(SCREENSPACE_STORAGE_KEY); } catch (_) {}
    if (stored === "false") _screenspaceVisible = false;
    else _screenspaceVisible = true;

    btn.addEventListener("click", toggleScreenspace);
    updateScreenspaceButton();
  }

  function updateScreenspaceButton() {
    var btn = qs("#screenspaceToggle");
    if (btn) btn.setAttribute("aria-pressed", _screenspaceVisible ? "true" : "false");
  }

  function toggleScreenspace() {
    _screenspaceVisible = !_screenspaceVisible;
    try { window.localStorage.setItem(SCREENSPACE_STORAGE_KEY, _screenspaceVisible ? "true" : "false"); } catch (_) {}
    updateScreenspaceButton();

    // Participant timelines mode: re-render to add/remove per-participant sub-tracks
    if (qs("#participantTimelines")) {
      var presentTypes = state.presentTypes || derivePresentTypes(state.artifacts);
      initParticipantTimelines(presentTypes);
      if (_filmstripEnabled) applyFilmstripMode();
      return;
    }

    // Unified track mode: show/hide the global screenspace track
    var wrap = qs("#screenspaceTrackWrap");
    var legend = qs("#screenspaceLegend");
    if (_screenspaceVisible) {
      var track = qs("#screenspaceTrack");
      if (track && !track.hasChildNodes() && data.screenspaceEvents && data.screenspaceEvents.length > 0) {
        renderScreenspaceTrack(data.screenspaceEvents, data.timeline.duration);
      } else {
        if (wrap) wrap.classList.remove("hidden");
        if (legend) legend.classList.remove("hidden");
      }
    } else {
      if (wrap) wrap.classList.add("hidden");
      if (legend) legend.classList.add("hidden");
    }
  }

  // ---- Scrub-audio toggle (opt-in; default off) ----

  function initScrubAudioToggle() {
    var btn = qs("#scrubAudioToggle");
    if (!btn) return;
    var stored = null;
    try { stored = window.localStorage.getItem(SCRUBAUDIO_STORAGE_KEY); } catch (_) {}
    _scrubAudioEnabled = stored === "true";
    btn.addEventListener("click", toggleScrubAudio);
    updateScrubAudioButton();
  }

  function updateScrubAudioButton() {
    var btn = qs("#scrubAudioToggle");
    if (btn) btn.setAttribute("aria-pressed", _scrubAudioEnabled ? "true" : "false");
  }

  function toggleScrubAudio() {
    _scrubAudioEnabled = !_scrubAudioEnabled;
    try { window.localStorage.setItem(SCRUBAUDIO_STORAGE_KEY, _scrubAudioEnabled ? "true" : "false"); } catch (_) {}
    updateScrubAudioButton();
    if (!_scrubAudioEnabled && window.clipgenCardScrubber) {
      window.clipgenCardScrubber.audioScrubStop();
      if (_cardScrub) window.clipgenCardScrubber.clearWaveform(_cardScrub.mediaEl);
    }
  }

  // ---- Initialization ----

  document.addEventListener("DOMContentLoaded", function () {
    initThemeToggle();

    if (!data || !data.artifacts) {
      showEmptyState();
      return;
    }

    state.artifacts = (data.artifacts || []).filter(function (a) {
      return a.id && a.file && (a.start != null || a.end != null);
    }).map(function (a, i) {
      return Object.assign({}, a, { _idx: i });
    });

    if (state.artifacts.length === 0) {
      showEmptyState();
      return;
    }

    var presentTypes = derivePresentTypes(state.artifacts);
    state.presentTypes = presentTypes;

    computeDuration();
    populateHeader();

    initScreenspaceToggle();

    if (qs("#participantTimelines")) {
      initParticipantTimelines(presentTypes);
    } else {
      initTypeLegend(presentTypes);
      initSeverityLegend();
      initTypeFilters(presentTypes);
      populateFilters();
      applyFilters();
      renderTimeline();
      renderList();
      initSortToolbar();
      bindFilterEvents();
    }

    if (!qs("#participantTimelines") && _screenspaceVisible && data.screenspaceEvents && data.screenspaceEvents.length > 0) {
      renderScreenspaceTrack(data.screenspaceEvents, data.timeline.duration);
    }

    initFilmstripToggle();
    initScrubAudioToggle();
    if (_filmstripEnabled) applyFilmstripMode();
  });

  function derivePresentTypes(artifacts) {
    var found = {};
    artifacts.forEach(function (a) {
      var t = a.type || "clip";
      found[t] = true;
    });
    var ordered = ["clip", "screen", "gif"];
    var types = [];
    ordered.forEach(function (t) {
      if (found[t]) types.push(t);
    });
    return types;
  }

  function showEmptyState() {
    var empty = qs("#emptyState");
    if (empty) empty.classList.remove("hidden");
    var layout = qs("#layout");
    if (layout) layout.style.display = "none";
    var tp = qs("#timelinePane");
    if (tp) tp.style.display = "none";
    var pp = qs("#playerPane");
    if (pp) pp.style.display = "none";
    var pt = qs("#participantTimelines");
    if (pt) pt.style.display = "none";
  }

  function computeDuration() {
    if (data.timeline && data.timeline.duration > 0) {
      state.duration = data.timeline.duration;
      return;
    }
    var maxTime = 0;
    state.artifacts.forEach(function (a) {
      var e = a.end || a.start || 0;
      if (e > maxTime) maxTime = e;
    });
    state.duration = maxTime > 0 ? maxTime * 1.05 : 1;
  }

  // ---- Header ----

  function populateHeader() {
    var meta = data.meta || {};
    setText("#studyName", meta.study || "");
    if (meta.generatedAt) {
      try {
        var d = new Date(meta.generatedAt);
        setText("#generatedAt", d.toLocaleString());
      } catch (_) {
        setText("#generatedAt", meta.generatedAt);
      }
    }
  }

  function setText(sel, val) {
    var node = qs(sel);
    if (node) node.textContent = val;
  }

  // ---- Filters ----

  function populateFilters() {
    var categories = uniqueFieldValues(state.artifacts, "category");
    var participants = uniqueFieldValues(state.artifacts, "participant");

    populateSelect(qs("#filterCategory"), categories, "All categories");
    populateSelect(qs("#filterParticipant"), participants, "All participants");

    var severities = sortedUniqueSeverities();
    var wrap = qs("#filterSeverityWrap");
    if (wrap) {
      if (severities.length === 0) {
        wrap.classList.add("hidden");
      } else {
        wrap.classList.remove("hidden");
        populateSelect(qs("#filterSeverity"), severities, "All severities");
      }
    }
  }

  function initTypeLegend(presentTypes) {
    var typeSet = {};
    presentTypes.forEach(function (t) {
      typeSet[t] = true;
    });

    var items = qsa("#timelineLegend .legend-item");
    items.forEach(function (item) {
      var swatch = item.querySelector(".legend-swatch");
      if (!swatch) return;
      var t = "";
      if (swatch.classList.contains("clip")) t = "clip";
      else if (swatch.classList.contains("screen")) t = "screen";
      else if (swatch.classList.contains("gif")) t = "gif";
      if (!t || !typeSet[t]) {
        item.style.display = "none";
      } else {
        item.style.display = "";
      }
    });
  }

  function initSeverityLegend() {
    var leg = qs("#severityLegend");
    if (!leg) return;
    var labels = sortedUniqueSeverities();
    if (labels.length === 0) {
      leg.innerHTML = "";
      leg.classList.add("hidden");
      return;
    }
    leg.classList.remove("hidden");
    leg.innerHTML = "";
    var prefix = document.createElement("span");
    prefix.textContent = "Severity: ";
    prefix.style.fontWeight = "600";
    leg.appendChild(prefix);
    labels.forEach(function (lab) {
      var wrap = el("span", "severity-legend-item");
      var sw = el("span", "legend-severity-swatch " + severityClass(lab));
      wrap.appendChild(sw);
      wrap.appendChild(document.createTextNode(lab));
      leg.appendChild(wrap);
    });
  }

  function initTypeFilters(presentTypes) {
    var typeSet = {};
    presentTypes.forEach(function (t) {
      typeSet[t] = true;
    });

    qsa("#filterType input[type=checkbox]").forEach(function (cb) {
      var t = cb.value;
      var label = cb.closest("label") || cb.parentElement;
      if (!typeSet[t]) {
        cb.checked = false;
        cb.disabled = true;
        if (label) label.style.display = "none";
      } else {
        cb.checked = true;
        cb.disabled = false;
        if (label) label.style.display = "";
      }
    });

    var fieldset = qs("#filterType");
    if (fieldset) {
      if (presentTypes.length <= 1) {
        fieldset.style.display = "none";
      } else {
        fieldset.style.display = "";
      }
    }
  }

  function bindFilterEvents() {
    var catSel = qs("#filterCategory");
    var partSel = qs("#filterParticipant");
    var sevSel = qs("#filterSeverity");
    var typeChecks = qsa("#filterType input[type=checkbox]");

    if (catSel) catSel.addEventListener("change", onFilterChange);
    if (partSel) partSel.addEventListener("change", onFilterChange);
    if (sevSel) sevSel.addEventListener("change", onFilterChange);
    typeChecks.forEach(function (cb) {
      cb.addEventListener("change", onFilterChange);
    });
  }

  function getActiveTypes() {
    var types = [];
    qsa("#filterType input[type=checkbox]").forEach(function (cb) {
      if (cb.checked) types.push(cb.value);
    });
    return types;
  }

  function applyFilters() {
    var cat = (qs("#filterCategory") || {}).value || "";
    var part = (qs("#filterParticipant") || {}).value || "";
    var types = getActiveTypes();
    var sevFilt = "";
    var sevWrap = qs("#filterSeverityWrap");
    if (sevWrap && !sevWrap.classList.contains("hidden")) {
      sevFilt = (qs("#filterSeverity") || {}).value || "";
    }

    var ids = {};
    state.filtered = state.artifacts.filter(function (a) {
      if (cat && a.category !== cat) return false;
      if (part && a.participant !== part) return false;
      if (sevFilt && (a.severity || "").trim() !== sevFilt) return false;
      if (types.indexOf(a.type) === -1) return false;
      ids[a.id] = true;
      return true;
    });

    state._filteredIds = ids;
  }

  function onFilterChange() {
    applyFilters();
    updateTimelineVisibility();
    updateListVisibility();
    updateCount();
    if (state.selectedId && !state._filteredIds[state.selectedId]) {
      clearSelection();
    }
  }

  // ---- Track layout algorithms ----

  function computeTrackAssignments(artifacts, duration) {
    var sorted = artifacts.slice().sort(function (a, b) {
      var sa = a.start || 0;
      var sb = b.start || 0;
      if (sa !== sb) return sa - sb;
      var da = (a.end || a.start || 0) - sa;
      var db = (b.end || b.start || 0) - sb;
      return da - db;
    });
    var minWidthSec = duration * 0.004;
    var trackEnds = [];
    var assignments = {};
    sorted.forEach(function (a) {
      var s = a.start || 0;
      var e = a.end || a.start || 0;
      var visualEnd = Math.max(e, s + minWidthSec);
      var placed = false;
      for (var i = 0; i < trackEnds.length; i++) {
        if (trackEnds[i] <= s) {
          trackEnds[i] = visualEnd;
          assignments[a.id] = i;
          placed = true;
          break;
        }
      }
      if (!placed) {
        assignments[a.id] = trackEnds.length;
        trackEnds.push(visualEnd);
      }
    });
    return { assignments: assignments, trackCount: trackEnds.length || 1 };
  }

  function computeCollapsedZIndices(artifacts) {
    var items = artifacts.map(function (a) {
      var s = a.start || 0;
      var e = a.end || a.start || 0;
      var rank = severityRank(a.severity);
      var sevVal = rank !== null ? rank : 999;
      return { id: a.id, duration: e - s, start: s, sevVal: sevVal };
    });
    items.sort(function (a, b) {
      if (a.sevVal !== b.sevVal) return b.sevVal - a.sevVal;
      if (a.duration !== b.duration) return b.duration - a.duration;
      return a.start - b.start;
    });
    var zMap = {};
    items.forEach(function (item, i) {
      zMap[item.id] = i + 1;
    });
    return zMap;
  }

  function applyTrackLayout(trackEl) {
    var trackId = trackEl._trackId;
    var markers = trackEl._trackMarkers;
    if (!markers || !markers.length) return;
    var isExpanded = !!state.expandedTracks[trackId];
    var isUnified = trackId === "unified";
    var markerHeight = isUnified ? 44 : 24;
    var topPad = isUnified ? 6 : 4;
    var gap = 4;

    if (isExpanded) {
      var artifactList = markers.map(function (m) { return m.artifact; });
      var packing = computeTrackAssignments(artifactList, state.duration);
      var numTracks = packing.trackCount;
      var expandedHeight = topPad + numTracks * (markerHeight + gap);
      trackEl.style.height = expandedHeight + "px";
      trackEl.classList.add("track-expanded");
      trackEl.classList.remove("track-collapsed");
      markers.forEach(function (m) {
        var row = packing.assignments[m.artifact.id] || 0;
        m.el.style.top = (topPad + row * (markerHeight + gap)) + "px";
        m.el.style.zIndex = "";
        m.el.dataset.collapsedZ = "";
      });
    } else {
      trackEl.style.height = "";
      trackEl.classList.remove("track-expanded");
      trackEl.classList.add("track-collapsed");
      var zMap = computeCollapsedZIndices(markers.map(function (m) { return m.artifact; }));
      markers.forEach(function (m) {
        m.el.style.top = topPad + "px";
        var z = zMap[m.artifact.id] || 1;
        m.el.dataset.collapsedZ = z;
        if (!m.el.classList.contains("selected")) {
          m.el.style.zIndex = z;
        }
      });
    }
  }

  function updateExpandButtonState(trackEl) {
    var btn = trackEl.parentNode && trackEl.parentNode.querySelector(".track-expand-btn");
    if (!btn) return;
    var markers = trackEl._trackMarkers;
    if (!markers || !markers.length) { btn.disabled = true; return; }
    var artifactList = markers.map(function (m) { return m.artifact; });
    var packing = computeTrackAssignments(artifactList, state.duration);
    var canExpand = packing.trackCount > 1;
    btn.disabled = !canExpand;
    if (!canExpand && state.expandedTracks[trackEl._trackId]) {
      state.expandedTracks[trackEl._trackId] = false;
      applyTrackLayout(trackEl);
      btn.classList.remove("expanded");
      btn.setAttribute("aria-expanded", "false");
      btn.title = "Expand tracks";
      btn.setAttribute("aria-label", btn.title);
    }
  }

  function toggleTrackExpand(trackEl) {
    var btn = trackEl.parentNode && trackEl.parentNode.querySelector(".track-expand-btn");
    if (btn && btn.disabled) return;
    var trackId = trackEl._trackId;
    state.expandedTracks[trackId] = !state.expandedTracks[trackId];
    applyTrackLayout(trackEl);
    if (btn) {
      var expanded = !!state.expandedTracks[trackId];
      btn.classList.toggle("expanded", expanded);
      btn.setAttribute("aria-expanded", expanded ? "true" : "false");
      btn.title = expanded ? "Collapse tracks" : "Expand tracks";
      btn.setAttribute("aria-label", btn.title);
    }
  }

  // ---- Timeline rendering ----

  function bindMarkerEvents(marker, a) {
    marker.addEventListener("mouseenter", function (ev) {
      onMarkerHover(ev);
      clearTimeout(_hoverDebounce);
      var target = ev.currentTarget;
      var cx = ev.clientX;
      _hoverDebounce = setTimeout(function () {
        var p = getMarkerProportion(target, cx);
        previewArtifact(a.id, a.type === "clip" ? p : 0);
      }, 60);
    });
    marker.addEventListener("mousemove", function (ev) {
      moveTooltip(ev);
      if (_preview && _preview.id === a.id && _preview.videoEl) {
        var p = getMarkerProportion(ev.currentTarget, ev.clientX);
        updatePreviewSeek(p);
      }
    });
    marker.addEventListener("mouseleave", function () {
      hideTooltip();
      clearTimeout(_hoverDebounce);
    });
    marker.addEventListener("click", function () {
      selectArtifact(a.id);
    });
  }

  function renderTimeline() {
    var track = qs("#timelineTrack");
    if (!track) return;
    track.innerHTML = "";

    var markers = [];
    state.artifacts.forEach(function (a) {
      var marker = el("div", markerClasses(a));
      marker.dataset.id = a.id;

      var startPct = ((a.start || 0) / state.duration) * 100;
      var endSec = a.end || a.start || 0;
      var widthPct = ((endSec - (a.start || 0)) / state.duration) * 100;
      var minWidth = 0.4;
      if (widthPct < minWidth) widthPct = minWidth;
      if (a.type === "screen") widthPct = Math.max(widthPct, 0.5);

      marker.style.left = startPct + "%";
      marker.style.width = widthPct + "%";

      bindMarkerEvents(marker, a);

      track.appendChild(marker);
      markers.push({ el: marker, artifact: a });
    });

    track._trackMarkers = markers;
    track._trackId = "unified";
    applyTrackLayout(track);
    updateExpandButtonState(track);

    var wrapBtn = track.parentNode && track.parentNode.querySelector(".track-expand-btn");
    if (wrapBtn && !wrapBtn._bound) {
      wrapBtn._bound = true;
      wrapBtn.addEventListener("click", function () {
        toggleTrackExpand(track);
      });
    }

    renderTicks();
    updateTimelineVisibility();
    updateCount();
  }

  function renderTicks() {
    var container = qs("#timelineTicks");
    if (!container) return;
    container.innerHTML = "";

    var numTicks = 8;
    var step = state.duration / numTicks;
    for (var i = 0; i <= numTicks; i++) {
      var tick = el("span", null, formatTime(i * step));
      container.appendChild(tick);
    }
  }

  function updateTimelineVisibility() {
    qsa(".artifact-marker").forEach(function (m) {
      var id = m.dataset.id;
      if (state._filteredIds[id]) {
        m.classList.remove("filtered-out");
      } else {
        m.classList.add("filtered-out");
      }
    });
  }

  // ---- List rendering ----

  function orderedArtifactsForList() {
    if (!state.listSort) return state.artifacts;
    var key = state.listSort.key;
    var dir = state.listSort.dir;
    return state.artifacts.slice().sort(function (a, b) {
      var r = 0;
      if (key === "severity") {
        var ae = !(a.severity || "").trim();
        var be = !(b.severity || "").trim();
        if (ae && be) r = 0;
        else if (ae) r = 1;
        else if (be) r = -1;
        else {
          var na = severityRank(a.severity);
          var nb = severityRank(b.severity);
          if (na === null) na = 999;
          if (nb === null) nb = 999;
          if (dir === "desc") r = na - nb;
          else r = nb - na;
        }
      } else if (key === "chrono") {
        var sa = Number(a.start);
        var sb = Number(b.start);
        if (isNaN(sa)) sa = 0;
        if (isNaN(sb)) sb = 0;
        r = sa - sb;
        if (dir === "desc") r = -r;
      } else if (key === "duration") {
        var da = artifactDurationSec(a);
        var db = artifactDurationSec(b);
        r = da - db;
        if (dir === "desc") r = -r;
      } else if (key === "alpha") {
        var ta = (a.description || "").trim();
        var tb = (b.description || "").trim();
        if (!ta && !tb) r = 0;
        else if (!ta) r = 1;
        else if (!tb) r = -1;
        else {
          r = ta.localeCompare(tb, undefined, { sensitivity: "base" });
          if (dir === "desc") r = -r;
        }
      }
      if (r !== 0) return r;
      return a._idx - b._idx;
    });
  }

  function sortToolbarLabel(key, dir, active) {
    if (key === "severity") {
      if (!active) return "Sort by severity";
      return dir === "desc"
        ? "Sort by severity: descending (most severe first)"
        : "Sort by severity: ascending (least severe first)";
    }
    if (key === "chrono") {
      if (!active) return "Sort by chronology (position in source)";
      return dir === "asc"
        ? "Sort by chronology: ascending (earliest in source first)"
        : "Sort by chronology: descending (latest in source first)";
    }
    if (key === "duration") {
      if (!active) return "Sort by duration";
      return dir === "desc"
        ? "Sort by duration: descending (longest first)"
        : "Sort by duration: ascending (shortest first)";
    }
    if (key === "alpha") {
      if (!active) return "Sort alphabetically (description)";
      return dir === "asc"
        ? "Sort alphabetically: ascending (A–Z)"
        : "Sort alphabetically: descending (Z–A)";
    }
    return "";
  }

  function updateSortToolbarUI() {
    var bar = qs("#artifactSortBar");
    if (!bar) return;
    qsa("#artifactSortBar .artifact-sort-btn").forEach(function (btn) {
      var key = btn.getAttribute("data-sort");
      btn.classList.remove("active", "sort-asc", "sort-desc");
      var dirEl = btn.querySelector(".sort-dir");
      if (dirEl) dirEl.textContent = "";
      var isActive = state.listSort && state.listSort.key === key;
      btn.setAttribute("aria-pressed", isActive ? "true" : "false");
      if (isActive) {
        var d = state.listSort.dir;
        btn.classList.add("active", d === "asc" ? "sort-asc" : "sort-desc");
        if (dirEl) dirEl.textContent = d === "asc" ? "\u2191" : "\u2193";
        btn.title = sortToolbarLabel(key, d, true);
        btn.setAttribute("aria-label", btn.title);
      } else {
        btn.title = sortToolbarLabel(key, null, false);
        btn.setAttribute("aria-label", btn.title);
      }
    });
  }

  function onSortButtonClick(ev) {
    var key = ev.currentTarget.getAttribute("data-sort");
    if (!key || !Object.prototype.hasOwnProperty.call(SORT_DEFAULT_DIR, key)) return;
    if (state.listSort && state.listSort.key === key) {
      state.listSort.dir = state.listSort.dir === "asc" ? "desc" : "asc";
    } else {
      state.listSort = { key: key, dir: SORT_DEFAULT_DIR[key] };
    }
    updateSortToolbarUI();
    renderList();
  }

  function initSortToolbar() {
    var bar = qs("#artifactSortBar");
    if (!bar || bar.dataset.bound === "1") return;
    bar.dataset.bound = "1";
    qsa("#artifactSortBar .artifact-sort-btn").forEach(function (btn) {
      btn.addEventListener("click", onSortButtonClick);
    });
    updateSortToolbarUI();
  }

  // Single pair of delegated listeners on #artifactList that dispatch to the
  // scrub/tooltip/click handlers based on the event target. Installed once on
  // first render and reused on subsequent renders — avoids attaching 4+
  // listeners per card (O(N) with the artifact count).
  var _artifactListDelegated = false;
  var _hoveredCardId = null;
  var _hoveredMediaEl = null;

  function _ensureArtifactListDelegation() {
    if (_artifactListDelegated) return;
    var list = qs("#artifactList");
    if (!list) return;
    _artifactListDelegated = true;

    list.addEventListener("click", function (ev) {
      var card = ev.target.closest && ev.target.closest(".artifact-card");
      if (!card || !list.contains(card)) return;
      var id = card.dataset.id;
      if (id) selectArtifact(id);
    });

    list.addEventListener("mouseover", function (ev) {
      var card = ev.target.closest && ev.target.closest(".artifact-card");
      if (!card || !list.contains(card)) return;
      var id = card.dataset.id;
      if (id !== _hoveredCardId) {
        _hoveredCardId = id;
        var a = findArtifact(id);
        if (a) showTooltipForArtifact(a, ev);
      }
      var media = ev.target.closest && ev.target.closest(".artifact-media");
      if (media && media !== _hoveredMediaEl) {
        _hoveredMediaEl = media;
        cardScrubEnter(media, ev);
      }
    });

    list.addEventListener("mouseout", function (ev) {
      var related = ev.relatedTarget;
      if (_hoveredCardId) {
        var card = document.querySelector(
          '#artifactList .artifact-card[data-id="' + _hoveredCardId + '"]'
        );
        if (card && (!related || !card.contains(related))) {
          _hoveredCardId = null;
          hideTooltip();
        }
      }
      if (_hoveredMediaEl) {
        if (!related || !_hoveredMediaEl.contains(related)) {
          _hoveredMediaEl = null;
          cardScrubLeave();
        }
      }
    });

    list.addEventListener("mousemove", function (ev) {
      if (_hoveredCardId) moveTooltip(ev);
      if (_hoveredMediaEl) cardScrubMove(_hoveredMediaEl, ev);
    });
  }

  function renderList() {
    var list = qs("#artifactList");
    if (!list) return;
    list.innerHTML = "";
    _hoveredCardId = null;
    _hoveredMediaEl = null;

    var frag = document.createDocumentFragment();
    orderedArtifactsForList().forEach(function (a) {
      if (a.type === "transcript") return;
      var card = el("div", "artifact-card");
      card.dataset.id = a.id;

      var media = el("div", "artifact-media");
      if (a.type === "screen" || a.type === "gif") {
        var img = document.createElement("img");
        img.src = a.file;
        img.alt = a.description || "";
        img.loading = "lazy";
        media.appendChild(img);
      } else if (a.type === "clip") {
        if (_thumbCache[a.id]) {
          var cimg = document.createElement("img");
          cimg.src = _thumbCache[a.id];
          cimg.alt = a.description || "";
          media.classList.add("thumb-loaded");
          media.appendChild(cimg);
        } else {
          media.classList.add("thumb-pending");
        }
      }
      card.appendChild(media);

      var meta = el("div", "artifact-meta");
      var badges = el("div", "artifact-badges");
      badges.appendChild(el("span", "badge badge-participant", a.participant));
      badges.appendChild(el("span", "badge badge-" + markerTypeClass(a.type), a.type || "clip"));
      if (a.category) badges.appendChild(el("span", "badge badge-category", a.category));
      var sev = (a.severity || "").trim();
      if (sev) badges.appendChild(el("span", "badge badge-severity " + severityClass(sev), sev));
      meta.appendChild(badges);
      meta.appendChild(el("div", "artifact-desc", a.description || "(no description)"));
      meta.appendChild(el("div", "artifact-time",
        formatTime(a.start) + (a.end != null ? " \u2013 " + formatTime(a.end) : "")
      ));
      card.appendChild(meta);

      frag.appendChild(card);
    });
    list.appendChild(frag);
    _ensureArtifactListDelegation();

    updateListVisibility();
    if (state.selectedId) {
      var sel = document.querySelector(
        '#artifactList .artifact-card[data-id="' + state.selectedId + '"]'
      );
      if (sel) sel.classList.add("selected");
    }
    initClipThumbnails();
  }

  function updateListVisibility() {
    qsa("#artifactList .artifact-card").forEach(function (card) {
      var id = card.dataset.id;
      var shouldShow = !!state._filteredIds[id];
      var isHidden = card.classList.contains("filtered-out");
      if (shouldShow && isHidden) {
        card.classList.remove("filtered-out");
      } else if (!shouldShow && !isHidden) {
        card.classList.add("filtered-out");
      }
    });
  }

  function updateCount() {
    var span = qs("#artifactCount");
    if (span) {
      span.textContent = "(" + state.filtered.length + " of " + state.artifacts.length + ")";
    }
  }

  // ---- Selection & detail ----

  function selectArtifactVisuals(id) {
    state.selectedId = id;

    qsa(".artifact-marker.selected").forEach(function (m) {
      m.classList.remove("selected");
      var storedZ = m.dataset.collapsedZ;
      m.style.zIndex = storedZ || "";
    });
    qsa("#artifactList .artifact-card.selected").forEach(function (c) {
      c.classList.remove("selected");
    });

    var marker = document.querySelector('.artifact-marker[data-id="' + id + '"]');
    if (marker) {
      marker.classList.add("selected");
      marker.style.zIndex = 1001;
    }
    var card = document.querySelector('#artifactList .artifact-card[data-id="' + id + '"]');
    if (card) {
      card.classList.add("selected");
      var sidebar = qs("#sidebar");
      if (sidebar) {
        var sRect = sidebar.getBoundingClientRect();
        var cRect = card.getBoundingClientRect();
        if (cRect.top < sRect.top) {
          sidebar.scrollTop += cRect.top - sRect.top;
        } else if (cRect.bottom > sRect.bottom) {
          sidebar.scrollTop += cRect.bottom - sRect.bottom;
        }
      }
    }
  }

  function selectArtifact(id) {
    if (state.selectedId === id) {
      clearSelection();
      return;
    }

    // If preview is already showing this artifact, promote it to active playback
    if (_preview && _preview.id === id && _preview.videoEl) {
      activatePreview();
      selectArtifactVisuals(id);
      return;
    }

    selectArtifactVisuals(id);

    var artifact = findArtifact(id);
    if (!artifact) return;

    if (qs("#playerPane")) {
      showPlayer(artifact);
    } else {
      showDetail(artifact);
    }
  }

  function clearSelection() {
    state.selectedId = null;
    _preview = null;
    qsa(".artifact-marker.selected").forEach(function (m) {
      m.classList.remove("selected");
      var storedZ = m.dataset.collapsedZ;
      m.style.zIndex = storedZ || "";
    });
    qsa("#artifactList .artifact-card.selected").forEach(function (c) {
      c.classList.remove("selected");
    });
    var empty = qs("#detailEmpty");
    var content = qs("#detailContent");
    if (empty) empty.classList.remove("hidden");
    if (content) content.classList.add("hidden");
    var playerEmpty = qs("#playerEmpty");
    var playerContent = qs("#playerContent");
    if (playerEmpty) playerEmpty.classList.remove("hidden");
    if (playerContent) playerContent.classList.add("hidden");
    var ps = qs("#playerSeverityPill");
    if (ps) {
      ps.textContent = "";
      ps.classList.add("hidden");
    }
  }

  function findArtifact(id) {
    for (var i = 0; i < state.artifacts.length; i++) {
      if (state.artifacts[i].id === id) return state.artifacts[i];
    }
    return null;
  }

  function showDetail(a) {
    var empty = qs("#detailEmpty");
    var content = qs("#detailContent");
    if (empty) empty.classList.add("hidden");
    if (content) content.classList.remove("hidden");

    populateDetailMeta(a);

    var preview = qs("#detailPreview");
    if (!preview) return;
    preview.innerHTML = "";
    _preview = null;

    if (!a.file) return;

    if (a.type === "screen") {
      var img = document.createElement("img");
      img.src = a.file;
      img.alt = a.description || "screenshot";
      preview.appendChild(img);
    } else if (a.type === "gif") {
      if (isVideoLoop(a.file)) {
        preview.appendChild(createLoopVideo(a.file, a.description || "gif"));
      } else {
        var gifImg = document.createElement("img");
        gifImg.src = a.file;
        gifImg.alt = a.description || "gif";
        preview.appendChild(gifImg);
      }
    } else {
      var vid = document.createElement("video");
      vid.controls = true;
      vid.preload = "metadata";
      vid.src = a.file;
      preview.appendChild(vid);
    }
  }

  // ---- Hover preview ----

  function getMarkerProportion(markerEl, clientX) {
    var rect = markerEl.getBoundingClientRect();
    if (rect.width <= 0) return 0;
    return Math.max(0, Math.min(1, (clientX - rect.left) / rect.width));
  }

  function previewArtifact(id, seekProportion) {
    var a = findArtifact(id);
    if (!a || !a.file) return;

    // Don't interrupt active playback
    if (state.selectedId && _preview && _preview.videoEl && !_preview.videoEl.paused) return;

    // Already previewing this artifact — just update seek
    if (_preview && _preview.id === id) {
      if (_preview.videoEl && seekProportion != null) {
        updatePreviewSeek(seekProportion);
      }
      return;
    }

    // Determine which preview container to use
    var isPlayerLayout = !!qs("#playerPane");
    var preview, empty, content;
    if (isPlayerLayout) {
      preview = qs("#playerPreview");
      empty = qs("#playerEmpty");
      content = qs("#playerContent");
    } else {
      preview = qs("#detailPreview");
      empty = qs("#detailEmpty");
      content = qs("#detailContent");
    }
    if (!preview) return;

    // Show the detail/player pane
    if (empty) empty.classList.add("hidden");
    if (content) content.classList.remove("hidden");

    // Populate metadata
    if (isPlayerLayout) {
      populatePlayerMeta(a);
    } else {
      populateDetailMeta(a);
    }

    // Clear existing preview
    preview.innerHTML = "";
    _preview = null;

    if (a.type === "screen") {
      var img = document.createElement("img");
      img.src = a.file;
      img.alt = a.description || "screenshot";
      preview.appendChild(img);
      _preview = { id: id, videoEl: null, wrapEl: null };
      return;
    }

    if (a.type === "gif") {
      if (isVideoLoop(a.file)) {
        preview.appendChild(createLoopVideo(a.file, a.description || "gif"));
      } else {
        var gifImg = document.createElement("img");
        gifImg.src = a.file;
        gifImg.alt = a.description || "gif";
        preview.appendChild(gifImg);
      }
      _preview = { id: id, videoEl: null, wrapEl: null };
      return;
    }

    // Clip: create video preview with play overlay
    var wrap = el("div", "video-preview-wrap");
    var vid = document.createElement("video");
    vid.preload = "auto";
    vid.muted = true;
    vid.playsInline = true;
    vid.src = a.file;

    var overlay = document.createElement("button");
    overlay.className = "video-play-overlay";
    overlay.type = "button";
    overlay.setAttribute("aria-label", "Play video");
    overlay.appendChild(el("span", "video-play-overlay-icon"));

    var timeBadge = el("span", "video-time-badge", "--:--");

    var proportion = seekProportion != null ? Math.max(0, Math.min(1, seekProportion)) : 0;

    vid.onloadedmetadata = function () {
      var dur = vid.duration;
      if (!dur || !isFinite(dur)) return;
      var seekTime = dur * proportion;
      vid.currentTime = seekTime;
      timeBadge.textContent = formatTime(seekTime) + " / " + formatTime(dur);
    };

    vid.onerror = function () {
      if (wrap.parentNode) wrap.parentNode.removeChild(wrap);
      _preview = { id: id, videoEl: null, wrapEl: null };
    };

    function doActivate() {
      activatePreview();
      var markerId = _preview ? _preview.id : id;
      selectArtifactVisuals(markerId);
    }

    overlay.addEventListener("click", function (ev) {
      ev.stopPropagation();
      doActivate();
    });
    function onVidClick() {
      if (vid.paused) doActivate();
    }
    vid.addEventListener("click", onVidClick);
    vid._previewClickHandler = onVidClick;

    wrap.appendChild(vid);
    wrap.appendChild(overlay);
    wrap.appendChild(timeBadge);
    preview.appendChild(wrap);

    _preview = { id: id, videoEl: vid, wrapEl: wrap, overlay: overlay, timeBadge: timeBadge };
    _lastSeekProportion = proportion;
  }

  function updatePreviewSeek(proportionX) {
    if (!_preview || !_preview.videoEl) return;
    var vid = _preview.videoEl;
    if (!vid.duration || !isFinite(vid.duration)) return;
    if (vid.readyState < 1) return; // metadata not yet loaded

    if (_seekRaf) return;
    _seekRaf = requestAnimationFrame(function () {
      _seekRaf = 0;
      var clamped = Math.max(0, Math.min(1, proportionX));
      var seekTime = vid.duration * clamped;
      vid.currentTime = seekTime;
      if (_preview && _preview.timeBadge) {
        _preview.timeBadge.textContent = formatTime(seekTime) + " / " + formatTime(vid.duration);
      }
      _lastSeekProportion = clamped;
    });
  }

  function activatePreview() {
    if (!_preview || !_preview.videoEl) return;
    var vid = _preview.videoEl;
    // Remove click-to-play so it doesn't fight native controls
    if (vid._previewClickHandler) {
      vid.removeEventListener("click", vid._previewClickHandler);
      vid._previewClickHandler = null;
    }
    vid.muted = false;
    vid.controls = true;
    vid.play();
    if (_preview.overlay) _preview.overlay.classList.add("hidden");
    if (_preview.timeBadge) _preview.timeBadge.classList.add("hidden");
  }

  function populateDetailMeta(a) {
    applySeverityPill(qs("#detailSeverityPill"), a);
    var badge = qs("#detailType");
    if (badge) {
      badge.textContent = (a.type || "clip").toUpperCase();
      badge.className = "detail-badge " + (a.type || "clip");
    }
    setText("#detailDescription", a.description || "(no description)");
    setText("#detailCategory", a.category || "\u2013");
    setText("#detailParticipant", a.participant || "\u2013");
    setText("#detailTime",
      formatTime(a.start) + (a.end != null ? " \u2013 " + formatTime(a.end) : ""));
    setText("#detailCell", a.cellA1 || (a.cellRow ? "R" + a.cellRow + "C" + a.cellCol : "\u2013"));
    setText("#detailSource", a.sourceVideo || "\u2013");
    setText("#detailFile", a.file || "\u2013");
    var intakeRow = qs("#detailIntakeRow");
    if (intakeRow) {
      if ((a.source === "screenspace" || a.source === "transcript") && a.intake_label) {
        setText("#detailIntakeDt", a.source === "transcript" ? "Transcript mark" : "Screenspace");
        setText("#detailIntakeLabel", a.intake_label);
        intakeRow.classList.remove("hidden");
      } else {
        intakeRow.classList.add("hidden");
      }
    }
    var trRow = qs("#detailTranscriptRow");
    if (trRow) {
      var txt = (a.transcriptText || "").trim();
      if (txt) {
        setText("#detailTranscript", txt);
        trRow.classList.remove("hidden");
      } else {
        trRow.classList.add("hidden");
      }
    }
  }

  function populatePlayerMeta(a) {
    applySeverityPill(qs("#playerSeverityPill"), a);
    var badge = qs("#playerType");
    if (badge) {
      badge.textContent = (a.type || "clip").toUpperCase();
      badge.className = "detail-badge " + (a.type || "clip");
    }
    setText("#playerDescription", a.description || "(no description)");
    var metaEl = qs("#playerMeta");
    if (metaEl) {
      var parts = [];
      if (a.participant) parts.push(escHtml(a.participant));
      parts.push(formatTime(a.start) + (a.end != null ? " \u2013 " + formatTime(a.end) : ""));
      if (a.category) parts.push(escHtml(a.category));
      metaEl.innerHTML = parts.join("&ensp;\u00B7&ensp;");
    }
  }

  // ---- Tooltip ----

  function onMarkerHover(ev) {
    var id = ev.currentTarget.dataset.id;
    var a = findArtifact(id);
    if (a) showTooltipForArtifact(a, ev);
  }

  function showTooltipForArtifact(a, ev) {
    var tip = qs("#tooltip");
    if (!tip) return;
    tip.style.borderLeft = "";

    var html = "<strong>" + escHtml(a.description || "(no description)") + "</strong><br>";
    html += '<span class="tooltip-time">' + formatTime(a.start);
    if (a.end != null) html += " – " + formatTime(a.end);
    html += "</span>";
    if (a.category) html += "<br>" + escHtml(a.category);
    if (a.participant) html += " · " + escHtml(a.participant);
    if ((a.severity || "").trim()) {
      html += "<br>" + escHtml(a.severity);
    }
    var transcript = (a.transcriptText || "").trim();
    if (transcript) {
      html +=
        '<div class="tooltip-transcript">' +
        '<span class="tooltip-transcript-label">Transcript</span>' +
        escHtml(transcript) +
        "</div>";
    }

    tip.innerHTML = html;
    tip.classList.remove("hidden");
    positionTooltip(tip, ev.clientX, ev.clientY);
  }

  function showTooltipForScreenspaceCluster(c, ev) {
    var tip = qs("#tooltip");
    if (!tip) return;
    var color = SS_DETECTOR_COLORS[c.type] || "#888";
    var avgConf = c.count > 0 ? (c.confSum / c.count) : 0;

    tip.innerHTML = "";
    tip.style.borderLeft = "3px solid " + color;

    var header = el("div", "ss-tooltip-header");
    var icon = buildDetectorIconSvg(c.type);
    if (icon) {
      icon.style.color = color;
      icon.style.flexShrink = "0";
      header.appendChild(icon);
    }
    var title = document.createElement("strong");
    title.textContent = c.eventType + " (" + c.type + ")";
    header.appendChild(title);
    tip.appendChild(header);

    var time = el("span", "tooltip-time");
    // A boundary is a single instant; padded/range display would misstate when
    // it occurred. Show one time for a point boundary, a span for a run.
    if (c.navigational && c.start === c.end) {
      time.textContent = formatTime(c.start);
    } else {
      time.textContent = formatTime(c.start) + " \u2013 " + formatTime(c.end);
    }
    tip.appendChild(time);

    var details = el("div", "ss-tooltip-details");
    details.appendChild(el("span", "", "Region: " + c.region));
    details.appendChild(el("span", "", "Participant: " + c.participant));
    details.appendChild(el("span", "", "Confidence: " + Math.round(avgConf * 100) + "%"));
    if (c.navigational && c.sceneLabel) {
      details.appendChild(el("span", "", "Enters: " + c.sceneLabel));
    }
    if (c.navigational && typeof c.distance === "number") {
      details.appendChild(el("span", "", "Distance: " + Math.round(c.distance)));
    }
    if (c.count > 1) details.appendChild(el("span", "", "Events: " + c.count));
    tip.appendChild(details);

    tip.classList.remove("hidden");
    positionTooltip(tip, ev.clientX, ev.clientY);
  }

  var _tooltipRaf = 0;
  function moveTooltip(ev) {
    var clientX = ev.clientX;
    var clientY = ev.clientY;
    if (_tooltipRaf) return;
    _tooltipRaf = requestAnimationFrame(function () {
      _tooltipRaf = 0;
      var tip = qs("#tooltip");
      if (tip && !tip.classList.contains("hidden")) {
        positionTooltip(tip, clientX, clientY);
      }
    });
  }

  function positionTooltip(tip, clientX, clientY) {
    var x = clientX + 12;
    var y = clientY + 12;
    var rect = tip.getBoundingClientRect();
    if (x + rect.width > window.innerWidth - 8) {
      x = clientX - rect.width - 12;
    }
    if (y + rect.height > window.innerHeight - 8) {
      y = clientY - rect.height - 12;
    }
    tip.style.left = x + "px";
    tip.style.top = y + "px";
  }

  function hideTooltip() {
    var tip = qs("#tooltip");
    if (tip) tip.classList.add("hidden");
  }

  function escHtml(str) {
    var div = document.createElement("div");
    div.textContent = str;
    return div.innerHTML;
  }

  // ---- Participant timeline viewer ----

  function initParticipantTimelines(presentTypes) {
    initTypeLegend(presentTypes);
    initSeverityLegend();

    var grouped = {};
    var participantOrder = [];
    state.artifacts.forEach(function (a) {
      var p = a.participant || "Unknown";
      if (!grouped[p]) {
        grouped[p] = [];
        participantOrder.push(p);
      }
      grouped[p].push(a);
    });
    participantOrder.sort();

    // Pre-cluster screenspace events by participant
    var ssEvents = (_screenspaceVisible && data.screenspaceEvents) ? data.screenspaceEvents : [];
    var ssClusters = ssEvents.length ? clusterScreenspaceEvents(ssEvents, state.duration) : [];
    var ssByParticipant = {};
    ssClusters.forEach(function (c) {
      if (!ssByParticipant[c.participant]) ssByParticipant[c.participant] = [];
      ssByParticipant[c.participant].push(c);
    });

    var container = qs("#participantRows");
    if (!container) return;
    container.innerHTML = "";

    var allDetectorsSeen = {};

    participantOrder.forEach(function (pid) {
      var row = el("div", "participant-row");

      var label = el("div", "participant-label", pid);
      row.appendChild(label);

      var tracksCol = el("div", "participant-tracks-col");

      var track = el("div", "participant-track");
      var markers = [];
      grouped[pid].forEach(function (a) {
        var marker = el("div", markerClasses(a));
        marker.dataset.id = a.id;

        var startPct = ((a.start || 0) / state.duration) * 100;
        var endSec = a.end || a.start || 0;
        var widthPct = ((endSec - (a.start || 0)) / state.duration) * 100;
        if (widthPct < 0.4) widthPct = 0.4;
        if (a.type === "screen") widthPct = Math.max(widthPct, 0.5);

        marker.style.left = startPct + "%";
        marker.style.width = widthPct + "%";

        bindMarkerEvents(marker, a);

        track.appendChild(marker);
        markers.push({ el: marker, artifact: a });
      });

      track._trackMarkers = markers;
      track._trackId = "participant-" + pid;
      applyTrackLayout(track);

      tracksCol.appendChild(track);

      // Add screenspace sub-track for this participant
      var pidClusters = ssByParticipant[pid];
      if (pidClusters && pidClusters.length) {
        var ssTrack = el("div", "participant-ss-track");
        var seen = renderScreenspaceMarkers(pidClusters, ssTrack, state.duration);
        Object.keys(seen).forEach(function (k) { allDetectorsSeen[k] = true; });
        tracksCol.appendChild(ssTrack);
      }

      row.appendChild(tracksCol);

      var expandBtn = el("button", "track-expand-btn");
      expandBtn.type = "button";
      expandBtn.setAttribute("aria-label", "Expand tracks");
      expandBtn.setAttribute("aria-expanded", "false");
      expandBtn.title = "Expand tracks";
      expandBtn.appendChild(el("span", "track-expand-icon"));
      expandBtn.addEventListener("click", (function (t) {
        return function () { toggleTrackExpand(t); };
      })(track));
      row.appendChild(expandBtn);
      updateExpandButtonState(track);

      container.appendChild(row);
    });

    // Build screenspace legend from all participants' events
    var legend = qs("#screenspaceLegend");
    if (Object.keys(allDetectorsSeen).length && _screenspaceVisible) {
      buildScreenspaceLegend(legend, allDetectorsSeen);
    } else if (legend) {
      legend.classList.add("hidden");
    }

    renderParticipantTicks();
  }

  function renderParticipantTicks() {
    var container = qs("#participantTicks");
    if (!container) return;
    container.innerHTML = "";

    var numTicks = 8;
    var step = state.duration / numTicks;
    for (var i = 0; i <= numTicks; i++) {
      var tick = el("span", null, formatTime(i * step));
      container.appendChild(tick);
    }
  }

  function showPlayer(a) {
    var empty = qs("#playerEmpty");
    var content = qs("#playerContent");
    if (empty) empty.classList.add("hidden");
    if (content) content.classList.remove("hidden");

    populatePlayerMeta(a);

    var preview = qs("#playerPreview");
    if (!preview) return;
    preview.innerHTML = "";
    _preview = null;

    if (!a.file) return;

    if (a.type === "screen") {
      var img = document.createElement("img");
      img.src = a.file;
      img.alt = a.description || "screenshot";
      preview.appendChild(img);
    } else if (a.type === "gif") {
      if (isVideoLoop(a.file)) {
        preview.appendChild(createLoopVideo(a.file, a.description || "gif"));
      } else {
        var gifImg = document.createElement("img");
        gifImg.src = a.file;
        gifImg.alt = a.description || "gif";
        preview.appendChild(gifImg);
      }
    } else {
      var vid = document.createElement("video");
      vid.controls = true;
      vid.autoplay = true;
      vid.src = a.file;
      preview.appendChild(vid);
    }
  }

  // ---- Screenspace track ----

  var SS_DETECTOR_COLORS = DETECTOR_COLORS;

  var SS_DETECTOR_ICON_PATHS = {
    multitool: { viewBox: "0 0 16 16", paths: [
      { d: "M8.91421 6.02513C9.2071 5.73223 9.68197 5.73223 9.97487 6.02513C11.3417 7.39196 11.3417 9.60804 9.97487 10.9749L7.97487 12.9749C6.60803 14.3417 4.39195 14.3417 3.02512 12.9749C1.82824 11.778 1.67995 9.93153 2.57781 8.57265C2.80615 8.22706 3.27142 8.13202 3.61701 8.36036C3.9626 8.5887 4.05765 9.05397 3.8293 9.39956C3.31651 10.1757 3.40282 11.2313 4.08578 11.9142C4.86683 12.6953 6.13316 12.6953 6.91421 11.9142L8.91421 9.91421C9.69525 9.13316 9.69525 7.86683 8.91421 7.08579C8.62131 6.79289 8.62131 6.31802 8.91421 6.02513Z", fillRule: "evenodd" },
      { d: "M7.08578 9.97487C6.79289 10.2678 6.31801 10.2678 6.02512 9.97487C4.65828 8.60804 4.65829 6.39196 6.02512 5.02513L8.02512 3.02513C9.39195 1.65829 11.608 1.65829 12.9749 3.02513C14.1717 4.22201 14.32 6.06847 13.4222 7.42735C13.1938 7.77294 12.7286 7.86798 12.383 7.63964C12.0374 7.4113 11.9423 6.94603 12.1707 6.60044C12.6835 5.82435 12.5972 4.76874 11.9142 4.08579C11.1332 3.30474 9.86683 3.30474 9.08578 4.08579L7.08578 6.08579C6.30473 6.86683 6.30473 8.13316 7.08578 8.91421C7.37867 9.20711 7.37867 9.68198 7.08578 9.97487Z", fillRule: "evenodd" }
    ]},
    color: { viewBox: "0 0 16 16", paths: [
      { d: "M15 4C15 5.39788 14.0439 6.57245 12.75 6.90549V8.5C12.75 8.69891 12.671 8.88968 12.5303 9.03033L12.0303 9.53033C11.7374 9.82322 11.2626 9.82322 10.9697 9.53033L10.25 8.81069L5.57322 13.4875C5.24503 13.8157 4.79992 14.0001 4.33579 14.0001H3.66421C3.59791 14.0001 3.53432 14.0264 3.48744 14.0733L2.78033 14.7804C2.63968 14.921 2.44891 15.0001 2.25 15.0001C2.05109 15.0001 1.86032 14.921 1.71967 14.7804L1.21967 14.2804C0.926777 13.9875 0.926777 13.5126 1.21967 13.2197L1.92678 12.5126C1.97366 12.4657 2 12.4021 2 12.3358V11.6643C2 11.2001 2.18437 10.755 2.51256 10.4268L7.18937 5.75003L6.46967 5.03033C6.17678 4.73744 6.17678 4.26256 6.46967 3.96967L6.96967 3.46967C7.11032 3.32902 7.30109 3.25 7.5 3.25H9.09451C9.42755 1.95608 10.6021 1 12 1C13.6569 1 15 2.34315 15 4ZM9.18937 7.75003L8.25003 6.81069L3.57322 11.4875C3.52634 11.5344 3.5 11.598 3.5 11.6643V12.3358C3.5 12.3938 3.49713 12.4514 3.49146 12.5086C3.54862 12.5029 3.60627 12.5001 3.66421 12.5001H4.33579C4.40209 12.5001 4.46568 12.4737 4.51256 12.4268L9.18937 7.75003Z", fillRule: "evenodd" }
    ]},
    change: { viewBox: "0 0 16 16", paths: [
      { d: "M9.58011 1.07655C9.88578 1.22638 10.0522 1.56328 9.98545 1.89709L9.16486 6H13.25C13.5437 6 13.8103 6.17136 13.9323 6.43847C14.0542 6.70558 14.0091 7.0193 13.8168 7.2412L7.31678 14.7412C7.09383 14.9984 6.72559 15.0733 6.41991 14.9234C6.11424 14.7736 5.94781 14.4367 6.01458 14.1029L6.83516 10H2.75001C2.45637 10 2.18974 9.82864 2.06777 9.56153C1.9458 9.29442 1.99093 8.9807 2.18324 8.7588L8.68324 1.2588C8.90619 1.00155 9.27444 0.92672 9.58011 1.07655Z", fillRule: "evenodd" }
    ]},
    similarity: { viewBox: "0 0 16 16", paths: [
      { d: "M2 4C2 2.89543 2.89543 2 4 2H12C13.1046 2 14 2.89543 14 4V12C14 13.1046 13.1046 14 12 14H4C2.89543 14 2 13.1046 2 12V4ZM12.5 9.70711C12.5 9.5745 12.4473 9.44732 12.3536 9.35355L11.3536 8.35355C11.1583 8.15829 10.8417 8.15829 10.6464 8.35355L9.35355 9.64645C9.15829 9.84171 8.84171 9.84171 8.64645 9.64645L6.35355 7.35355C6.15829 7.15829 5.84171 7.15829 5.64645 7.35355L3.64645 9.35355C3.55268 9.44732 3.5 9.5745 3.5 9.70711V12C3.5 12.2761 3.72386 12.5 4 12.5H12C12.2761 12.5 12.5 12.2761 12.5 12V9.70711ZM12 5C12 5.55228 11.5523 6 11 6C10.4477 6 10 5.55228 10 5C10 4.44772 10.4477 4 11 4C11.5523 4 12 4.44772 12 5Z", fillRule: "evenodd" }
    ]},
    text: { viewBox: "0 0 16 16", paths: [
      { d: "M11 5C11.299 5 11.5693 5.17751 11.6882 5.45179L14.9382 12.9518C15.1029 13.3319 14.9283 13.7735 14.5482 13.9382C14.1682 14.1029 13.7266 13.9283 13.5619 13.5482L12.8908 11.9997H9.10923L8.4382 13.5482C8.2735 13.9283 7.83189 14.1029 7.45182 13.9382C7.07176 13.7735 6.89717 13.3319 7.06186 12.9518L10.3119 5.45179C10.4307 5.17751 10.7011 5 11 5ZM9.75923 10.4997H12.2408L11 7.63628L9.75923 10.4997Z", fillRule: "evenodd" },
      { d: "M5.00003 1C5.41424 1 5.75003 1.33579 5.75003 1.75V3.01104C6.16299 3.02322 6.5735 3.04541 6.98131 3.0774C7.44038 3.11341 7.89601 3.16182 8.34786 3.22231C8.75842 3.27727 9.04668 3.65464 8.99172 4.06519C8.93676 4.47574 8.55938 4.76401 8.14883 4.70905C7.92894 4.67961 7.70808 4.65321 7.48628 4.6299C7.1301 5.85717 6.59808 7.00928 5.91941 8.05729C6.15555 8.36066 6.40658 8.65193 6.67142 8.92999C6.95709 9.22993 6.94553 9.70466 6.64559 9.99034C6.34565 10.276 5.87092 10.2644 5.58525 9.96451C5.38294 9.7521 5.18774 9.53284 5.00002 9.30711C4.18402 10.2884 3.22645 11.1474 2.15883 11.853C1.81326 12.0813 1.34799 11.9863 1.11962 11.6408C0.891239 11.2952 0.986242 10.8299 1.33181 10.6015C2.3813 9.90797 3.31021 9.04714 4.08066 8.05729C3.88359 7.75296 3.69887 7.43984 3.52724 7.11865C3.33202 6.75332 3.46992 6.29891 3.83524 6.10369C4.20057 5.90847 4.65498 6.04637 4.8502 6.4117C4.89895 6.50293 4.9489 6.59343 5.00002 6.68318C5.38798 6.00207 5.7083 5.27759 5.95187 4.51891C5.63619 4.50635 5.31887 4.5 5.00003 4.5C3.93193 4.5 2.88086 4.57121 1.85122 4.70905C1.44067 4.76401 1.0633 4.47574 1.00834 4.06519C0.95338 3.65464 1.24164 3.27727 1.65219 3.22231C2.50548 3.10808 3.37219 3.03692 4.25003 3.01104V1.75C4.25003 1.33579 4.58582 1 5.00003 1Z", fillRule: "evenodd" }
    ]},
    numbers: { viewBox: "0 0 16 16", paths: [
      { d: "M7.48677 2.89033C7.56427 2.48344 7.29725 2.09075 6.89035 2.01325C6.48345 1.93574 6.09077 2.20277 6.01326 2.60967L5.55827 4.99835H3.60963C3.19542 4.99835 2.85963 5.33414 2.85963 5.74835C2.85963 6.16257 3.19542 6.49835 3.60963 6.49835H5.27256L4.7016 9.49589H2.74963C2.33542 9.49589 1.99963 9.83168 1.99963 10.2459C1.99963 10.6601 2.33542 10.9959 2.74963 10.9959H4.41588L4.01326 13.1097C3.93576 13.5166 4.20278 13.9092 4.60968 13.9868C5.01658 14.0643 5.40926 13.7972 5.48677 13.3903L5.94285 10.9959H8.91589L8.51326 13.1097C8.43576 13.5166 8.70278 13.9092 9.10968 13.9868C9.51658 14.0643 9.90927 13.7972 9.98677 13.3903L10.4429 10.9959H12.3896C12.8038 10.9959 13.1396 10.6601 13.1396 10.2459C13.1396 9.83168 12.8038 9.49589 12.3896 9.49589H10.7286L11.2995 6.49835H13.2496C13.6638 6.49835 13.9996 6.16257 13.9996 5.74835C13.9996 5.33414 13.6638 4.99835 13.2496 4.99835H11.5852L11.9868 2.89033C12.0643 2.48344 11.7972 2.09075 11.3903 2.01325C10.9835 1.93574 10.5908 2.20277 10.5133 2.60967L10.0583 4.99835H7.08524L7.48677 2.89033ZM6.79953 6.49835L6.22857 9.49589H9.2016L9.77256 6.49835H6.79953Z", fillRule: "evenodd" }
    ]},
    template: { viewBox: "0 0 16 16", paths: [
      { d: "M2 3.5A1.5 1.5 0 0 1 3.5 2H5a.75.75 0 0 1 0 1.5H3.5v1.75a.75.75 0 0 1-1.5 0V3.5ZM11 2a.75.75 0 0 0 0 1.5h1.5v1.75a.75.75 0 0 0 1.5 0V3.5A1.5 1.5 0 0 0 12.5 2H11ZM2.75 10.75a.75.75 0 0 1 .75.75v1.5H5a.75.75 0 0 1 0 1.5H3.5A1.5 1.5 0 0 1 2 13v-1.5a.75.75 0 0 1 .75-.75ZM13.25 10.75a.75.75 0 0 1 .75.75V13a1.5 1.5 0 0 1-1.5 1.5H11a.75.75 0 0 1 0-1.5h1.5v-1.5a.75.75 0 0 1 .75-.75ZM10 8a2 2 0 1 1-4 0 2 2 0 0 1 4 0Z", fillRule: "evenodd" }
    ]},
    flow: { viewBox: "0 0 16 16", paths: [
      { d: "M5.28 10.22a.75.75 0 0 1 0 1.06l-1.47 1.47h8.44a.75.75 0 0 1 0 1.5H3.81l1.47 1.47a.75.75 0 0 1-1.06 1.06l-2.75-2.75a.75.75 0 0 1 0-1.06l2.75-2.75a.75.75 0 0 1 1.06 0ZM10.72.22a.75.75 0 0 1 1.06 0l2.75 2.75a.75.75 0 0 1 0 1.06l-2.75 2.75a.75.75 0 1 1-1.06-1.06l1.47-1.47H3.75a.75.75 0 0 1 0-1.5h8.44L10.72 1.28a.75.75 0 0 1 0-1.06Z", fillRule: "evenodd" }
    ]},
    scene: { viewBox: "0 0 16 16", paths: [
      { d: "M2 3.5A1.5 1.5 0 0 1 3.5 2h2A1.5 1.5 0 0 1 7 3.5v2A1.5 1.5 0 0 1 5.5 7h-2A1.5 1.5 0 0 1 2 5.5v-2ZM9 3.5A1.5 1.5 0 0 1 10.5 2h2A1.5 1.5 0 0 1 14 3.5v2A1.5 1.5 0 0 1 12.5 7h-2A1.5 1.5 0 0 1 9 5.5v-2ZM2 10.5A1.5 1.5 0 0 1 3.5 9h2A1.5 1.5 0 0 1 7 10.5v2A1.5 1.5 0 0 1 5.5 14h-2A1.5 1.5 0 0 1 2 12.5v-2ZM9 10.5A1.5 1.5 0 0 1 10.5 9h2A1.5 1.5 0 0 1 14 10.5v2A1.5 1.5 0 0 1 12.5 14h-2A1.5 1.5 0 0 1 9 12.5v-2Z" }
    ]},
    inactivity: { viewBox: "0 0 16 16", paths: [
      { d: "M15 8C15 11.866 11.866 15 8 15C4.13401 15 1 11.866 1 8C1 4.13401 4.13401 1 8 1C11.866 1 15 4.13401 15 8ZM5.5 5.5C5.5 5.22386 5.72386 5 6 5H6.5C6.77614 5 7 5.22386 7 5.5V10.5C7 10.7761 6.77614 11 6.5 11H6C5.72386 11 5.5 10.7761 5.5 10.5V5.5ZM9.5 5C9.22386 5 9 5.22386 9 5.5V10.5C9 10.7761 9.22386 11 9.5 11H10C10.2761 11 10.5 10.7761 10.5 10.5V5.5C10.5 5.22386 10.2761 5 10 5H9.5Z", fillRule: "evenodd" }
    ]}
  };

  function buildDetectorIconSvg(type) {
    var info = SS_DETECTOR_ICON_PATHS[type];
    if (!info) return null;
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "14");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", info.viewBox);
    svg.setAttribute("fill", "currentColor");
    info.paths.forEach(function (p) {
      var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
      path.setAttribute("d", p.d);
      if (p.fillRule) {
        path.setAttribute("fill-rule", p.fillRule);
        path.setAttribute("clip-rule", p.fillRule);
      }
      svg.appendChild(path);
    });
    return svg;
  }

  function clusterScreenspaceEvents(events, timelineDuration) {
    var sorted = events.slice().sort(function (a, b) {
      if (a.participant !== b.participant) return a.participant < b.participant ? -1 : 1;
      if (a.eventType !== b.eventType) return a.eventType < b.eventType ? -1 : 1;
      return a.timeIn - b.timeIn;
    });
    var clusters = [];
    var cur = null;
    for (var i = 0; i < sorted.length; i++) {
      var ev = sorted[i];
      if (
        !cur ||
        ev.participant !== cur.participant ||
        ev.eventType !== cur.eventType ||
        // Boundary ticks are individual points — never merge them, or a cluster
        // would render only its first tick and hide later boundaries.
        ev.navigational ||
        ev.timeIn - cur.end > 5
      ) {
        if (cur) clusters.push(cur);
        cur = {
          participant: ev.participant,
          start: ev.timeIn,
          end: ev.timeOut,
          eventType: ev.eventType,
          type: ev.type,
          region: ev.region,
          navigational: !!ev.navigational,
          distance: (ev.metadata && ev.metadata.distance) || 0,
          sceneLabel: (ev.metadata && ev.metadata.scene_label) || "",
          count: 1,
          confSum: ev.confidence,
        };
      } else {
        cur.end = Math.max(cur.end, ev.timeOut);
        cur.count++;
        cur.confSum += ev.confidence;
        var d = (ev.metadata && ev.metadata.distance) || 0;
        if (d > cur.distance) cur.distance = d;
      }
    }
    if (cur) clusters.push(cur);

    for (var j = 0; j < clusters.length; j++) {
      // Navigational (boundary) ticks must sit at the real boundary time — keep
      // their exact instant. Other point detections get a ±2s window so their
      // hover/clip context is usable.
      if (!clusters[j].navigational && clusters[j].start === clusters[j].end) {
        clusters[j].start = Math.max(0, clusters[j].start - 2);
        clusters[j].end = Math.min(timelineDuration, clusters[j].end + 2);
      }
    }
    return clusters;
  }

  function renderScreenspaceMarkers(clusters, trackEl, timelineDuration) {
    var detectorsSeen = {};
    clusters.forEach(function (c) {
      // Clamp to timeline bounds so markers don't spill past the track edge
      var clampedStart = Math.max(0, Math.min(c.start, timelineDuration));
      var clampedEnd = Math.max(0, Math.min(c.end, timelineDuration));
      if (clampedStart >= timelineDuration) return;
      var left = (clampedStart / timelineDuration) * 100;
      var width = Math.max(((clampedEnd - clampedStart) / timelineDuration) * 100, 0.4);
      var marker = document.createElement("div");
      marker.className = "screenspace-marker ss-type-" + c.type;
      // Navigational (boundary) events are orientation scaffolding — draw them
      // as thin, lighter ticks rather than findings spans.
      if (c.navigational) marker.className += " screenspace-marker--navigational";
      marker.style.left = left + "%";
      marker.style.width = c.navigational ? "1px" : (width + "%");
      marker.style.background = "var(--color-task-" + c.type + ", #888)";
      marker.addEventListener("mouseenter", function (cluster) {
        return function (ev) { showTooltipForScreenspaceCluster(cluster, ev); };
      }(c));
      marker.addEventListener("mousemove", moveTooltip);
      marker.addEventListener("mouseleave", function () {
        var tip = qs("#tooltip");
        if (tip) tip.style.borderLeft = "";
        hideTooltip();
      });
      trackEl.appendChild(marker);
      detectorsSeen[c.type] = true;
    });
    return detectorsSeen;
  }

  function buildScreenspaceLegend(legendEl, detectorsSeen) {
    if (!legendEl) return;
    legendEl.innerHTML = "";
    var types = Object.keys(detectorsSeen);
    if (!types.length) return;
    types.forEach(function (t) {
      var item = document.createElement("span");
      item.className = "legend-item";
      var swatch = document.createElement("span");
      swatch.className = "legend-swatch";
      swatch.style.background = "var(--color-task-" + t + ", #888)";
      item.appendChild(swatch);
      item.appendChild(document.createTextNode(" " + t.charAt(0).toUpperCase() + t.slice(1)));
      legendEl.appendChild(item);
    });
    legendEl.classList.remove("hidden");
  }

  function renderScreenspaceTrack(events, timelineDuration) {
    var wrap = qs("#screenspaceTrackWrap");
    var track = qs("#screenspaceTrack");
    if (!wrap || !track || !timelineDuration) return;

    var clusters = clusterScreenspaceEvents(events, timelineDuration);
    track.innerHTML = "";
    var detectorsSeen = renderScreenspaceMarkers(clusters, track, timelineDuration);
    wrap.classList.remove("hidden");
    buildScreenspaceLegend(qs("#screenspaceLegend"), detectorsSeen);
  }
})();
