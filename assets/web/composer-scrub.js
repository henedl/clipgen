/* clipgen Composer — marker/cut media satellite.
 *
 * Owns the two opt-in timeline media layers: thumbnail strips drawn into
 * marker/cut bars (server sprite sheets from /composer/api/sprite, cached as
 * preloaded Images with lazy, debounced, viewport-filtered fetching so
 * renderTimeline stays a pure cache hit per frame), and the hover audio scrub
 * + waveform overlay (card-scrubber.js primitives — the viewer's pattern, not
 * attach(), because the bars are canvas rects, not DOM). The timeline reaches
 * this satellite late-bound via CO.drawMarkerThumbs / CO.scrubHoverMove /
 * CO.scrubHoverEnd (this file loads last); the hub calls initMarkerScrub /
 * resetScrubMedia through guarded delegators. Both toggles persist through
 * CO.persistLaneUi → PUT api/ui.
 */
(function () {
  "use strict";

  var CO = window.ClipgenComposer;
  var state = CO.state;

  var FETCH_DEBOUNCE_MS = 150; // quiet period after the last cache miss
  var FETCH_CONCURRENCY = 4;   // max in-flight sprite Images
  var SPRITE_CACHE_MAX = 150;  // decoded sprite sheets kept (LRU by tick)
  var MIN_THUMB_WIDTH = 24;    // bars narrower than this stay flat

  var _sprites = {};      // mediaKey → {img, ready, failed, tick}
  var _spriteCount = 0;
  var _tick = 0;          // monotonic LRU counter (bumped on insert + hit)
  var _wanted = {};       // mediaKey → {start, end} misses from recent renders
  var _fetchTimer = 0;
  var _inflight = 0;
  var _mediaVersion = 0;  // bumped on participant switch; stale loads no-op
  var _renderQueued = false;

  // Span-aware keys: committing a trim or cut edit changes start/end, which
  // naturally invalidates that bar's sprite and audio without bookkeeping.
  function mediaKey(keyStem, start, end) {
    return keyStem + "|" + start.toFixed(2) + "|" + end.toFixed(2);
  }

  function mediaUrl(kind, start, end) {
    return "api/" + kind + "/" + encodeURIComponent(state.participant) +
      "?start=" + start + "&end=" + end;
  }

  // ---- Sprite cache + lazy fetch ----

  function armFetchTimer() {
    // Re-arming on every miss keeps fetches quiet during pan/zoom churn:
    // they only fire once renders have settled for FETCH_DEBOUNCE_MS.
    if (_fetchTimer) clearTimeout(_fetchTimer);
    _fetchTimer = setTimeout(flushFetches, FETCH_DEBOUNCE_MS);
  }

  function queueRender() {
    if (_renderQueued) return;
    _renderQueued = true;
    requestAnimationFrame(function () {
      _renderQueued = false;
      if (CO.renderTimeline) CO.renderTimeline();
    });
  }

  function evictIfNeeded() {
    while (_spriteCount > SPRITE_CACHE_MAX) {
      var oldest = null;
      var oldestTick = Infinity;
      for (var key in _sprites) {
        var entry = _sprites[key];
        if (!entry.ready && !entry.failed) continue; // never drop in-flight
        if (entry.tick < oldestTick) {
          oldestTick = entry.tick;
          oldest = key;
        }
      }
      if (oldest === null) return;
      delete _sprites[oldest];
      _spriteCount--;
    }
  }

  function startFetch(key, span) {
    var version = _mediaVersion;
    var img = new Image();
    var entry = { img: img, ready: false, failed: false, tick: ++_tick };
    _sprites[key] = entry;
    _spriteCount++;
    _inflight++;
    evictIfNeeded();
    img.onload = function () {
      _inflight--;
      if (version !== _mediaVersion) return; // participant switched mid-flight
      entry.ready = true;
      queueRender();
      armFetchTimer(); // a slot freed — drain any queued-past-concurrency keys
    };
    img.onerror = function () {
      _inflight--;
      if (version !== _mediaVersion) return;
      entry.failed = true; // no retry; the flat bar stays
      armFetchTimer();
    };
    img.src = mediaUrl("sprite", span.start, span.end);
  }

  function flushFetches() {
    _fetchTimer = 0;
    if (!state.participant) {
      _wanted = {};
      return;
    }
    if (state.dragging) {
      // Cut/marker edge drags mutate spans every frame — wait them out.
      armFetchTimer();
      return;
    }
    var visLen = state.duration / (state.zoom || 1);
    var visStart = state.offset;
    var visEnd = state.offset + visLen;
    var keys = Object.keys(_wanted);
    for (var i = 0; i < keys.length && _inflight < FETCH_CONCURRENCY; i++) {
      var key = keys[i];
      var span = _wanted[key];
      delete _wanted[key];
      // Scrolled/zoomed away since the miss: drop it. A later render of the
      // bar re-enqueues it.
      if (span.end < visStart || span.start > visEnd) continue;
      startFetch(key, span);
    }
    if (Object.keys(_wanted).length) armFetchTimer();
  }

  // Called from renderTimeline for every visible marker/cut bar. Cache hit →
  // draw and report true; miss → enqueue a lazy fetch and report false (the
  // caller's flat bar stays). Must stay allocation-light: it runs per bar on
  // every pan/zoom frame.
  function drawMarkerThumbs(ctx, keyStem, start, end, x, y, w, h, tintColor) {
    if (w < MIN_THUMB_WIDTH || end <= start || !state.participant) return false;
    var key = mediaKey(keyStem, start, end);
    var entry = _sprites[key];
    if (!entry) {
      // Don't enqueue mid-drag: every drag frame renders a different span and
      // would leave a trail of junk keys to fetch after the drag commits.
      if (!state.dragging) {
        _wanted[key] = { start: start, end: end };
        armFetchTimer();
      }
      return false;
    }
    if (!entry.ready) return false; // failed or still loading
    entry.tick = ++_tick;

    var cols = CLIPGEN_CONFIG.cardScrubberSpriteCols;
    var rows = CLIPGEN_CONFIG.cardScrubberSpriteRows;
    var frameCount = cols * rows;
    var img = entry.img;
    var fw = img.width / cols;
    var fh = img.height / rows;
    var aspect = fh > 0 ? fw / fh : 1;
    // As many frames as fit at the bar's height, stretched slightly so the
    // strip fills the full width, sampled evenly across the span.
    var n = Math.max(1, Math.min(Math.floor(w / (h * aspect)), frameCount));
    var slotW = w / n;
    ctx.save();
    ctx.beginPath();
    ctx.rect(x, y, w, h);
    ctx.clip();
    for (var i = 0; i < n; i++) {
      var f = Math.min(frameCount - 1, Math.floor(((i + 0.5) / n) * frameCount));
      ctx.drawImage(
        img,
        (f % cols) * fw, Math.floor(f / cols) * fh, fw, fh,
        x + i * slotW, y, slotW, h
      );
    }
    // Light lane-color tint so source identity survives the imagery.
    ctx.fillStyle = hexToRgba(tintColor, 0.18);
    ctx.fillRect(x, y, w, h);
    ctx.restore();
    return true;
  }

  // ---- Hover audio scrub + waveform ----

  var _waveHost = null;   // #coScrubWave overlay div, lazily created
  var _lastAudioKey = null;

  function scrubHoverMove(keyStem, start, end, frac, rect) {
    var cs = window.clipgenCardScrubber;
    if (!cs || end <= start || !state.participant) return;
    if (end - start > CLIPGEN_CONFIG.composerScrubMaxAudioSeconds) {
      // The server caps the WAV at this length; scrubbing a longer span would
      // misalign hover fraction ↔ audio time, so skip it entirely.
      scrubHoverEnd();
      return;
    }
    var key = mediaKey(keyStem, start, end);
    var url = mediaUrl("audio", start, end);
    if (key !== _lastAudioKey) {
      _lastAudioKey = key;
      cs.loadAudioBuffer(url, key); // warm the decode on bar entry
    }
    cs.audioScrubAt(key, url, frac * (end - start));

    var host = _waveHost;
    if (!host) {
      var wrapper = qs("#coTimelineWrapper");
      if (!wrapper) return;
      host = document.createElement("div");
      host.id = "coScrubWave";
      wrapper.appendChild(host); // after the canvases + rail → paints on top
      _waveHost = host;
    }
    // card-scrubber caches its canvas size at creation; a differently-sized
    // bar needs a fresh canvas.
    if (Math.abs(host.offsetWidth - rect.width) > 1 ||
        Math.abs(host.offsetHeight - rect.height) > 1) {
      host.innerHTML = "";
    }
    host.style.left = rect.left + "px";
    host.style.top = rect.top + "px";
    host.style.width = rect.width + "px";
    host.style.height = rect.height + "px";
    host.style.display = "";
    var waveform = cs.extractWaveform(key);
    if (waveform) {
      var canvas = cs.getOrCreateWaveformCanvas(host); // null on 0×0 host
      if (canvas) cs.drawWaveform(canvas, waveform, frac);
    } else {
      cs.clearWaveform(host); // hide a previous bar's bars until this decodes
    }
  }

  function scrubHoverEnd() {
    _lastAudioKey = null;
    var cs = window.clipgenCardScrubber;
    if (cs) cs.audioScrubStop();
    if (_waveHost) _waveHost.style.display = "none";
  }

  // Participant switch: everything cached belongs to the old video.
  function resetScrubMedia() {
    _mediaVersion++;
    _sprites = {};
    _spriteCount = 0;
    _wanted = {};
    if (_fetchTimer) {
      clearTimeout(_fetchTimer);
      _fetchTimer = 0;
    }
    scrubHoverEnd();
    if (_waveHost) _waveHost.innerHTML = ""; // drop the size-cached canvas
  }

  // ---- Toggles ----

  function syncScrubToggles() {
    var thumbs = qs("#coThumbsToggle");
    if (thumbs) {
      thumbs.setAttribute("aria-pressed", state.markerThumbnails ? "true" : "false");
    }
    var scrub = qs("#coScrubAudioToggle");
    if (scrub) {
      scrub.setAttribute("aria-pressed", state.markerAudioScrub ? "true" : "false");
    }
  }

  function initMarkerScrub() {
    var thumbs = qs("#coThumbsToggle");
    if (thumbs) {
      thumbs.addEventListener("click", function () {
        state.markerThumbnails = !state.markerThumbnails;
        syncScrubToggles();
        if (CO.persistLaneUi) CO.persistLaneUi();
        if (CO.updateTimelineHeight) CO.updateTimelineHeight();
        if (CO.renderTimeline) CO.renderTimeline();
      });
    }
    var scrub = qs("#coScrubAudioToggle");
    if (scrub) {
      scrub.addEventListener("click", function () {
        state.markerAudioScrub = !state.markerAudioScrub;
        syncScrubToggles();
        if (CO.persistLaneUi) CO.persistLaneUi();
        if (!state.markerAudioScrub) scrubHoverEnd();
      });
    }
    syncScrubToggles();
  }

  CO.initMarkerScrub = initMarkerScrub;
  CO.syncScrubToggles = syncScrubToggles;
  CO.drawMarkerThumbs = drawMarkerThumbs;
  CO.scrubHoverMove = scrubHoverMove;
  CO.scrubHoverEnd = scrubHoverEnd;
  CO.resetScrubMedia = resetScrubMedia;
})();
