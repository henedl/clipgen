/* clipgen Composer — marker/cut media satellite.
 *
 * Owns the two opt-in timeline media layers: thumbnail strips drawn into
 * marker/cut bars (server sprite sheets from /composer/api/sprite, cached as
 * preloaded Images with lazy, debounced, viewport-filtered fetching so
 * renderTimeline stays a pure cache hit per frame), and the hover audio scrub
 * + waveform overlay (card-scrubber.js primitives — the viewer's pattern, not
 * attach(), because the bars are canvas rects, not DOM).
 *
 * Strips are zoom-adaptive: a bar is covered by TILES, each one sprite sheet
 * whose per-frame duration comes from a power-of-two "slot seconds" ladder
 * matched to the current px/sec (one slot ≈ a 16:9 frame at the bar's
 * height). Zooming in crosses a ladder step and refetches finer tiles instead
 * of stretching the same frames; panning reuses tiles (they are anchored to
 * the bar's start, not the viewport). Only tiles intersecting the canvas
 * viewport are fetched or drawn, so long bars cost nothing offscreen.
 *
 * The timeline reaches this satellite late-bound via CO.drawMarkerThumbs /
 * CO.scrubHoverMove / CO.scrubHoverEnd (this file loads last); the hub calls
 * initMarkerScrub / resetScrubMedia through guarded delegators. Both toggles
 * persist through CO.persistLaneUi → PUT api/ui.
 */
(function () {
  "use strict";

  var CO = window.ClipgenComposer;
  var state = CO.state;

  var FETCH_DEBOUNCE_MS = 150; // quiet period after the last cache miss
  var FETCH_CONCURRENCY = 4;   // max in-flight sprite Images
  var SPRITE_CACHE_MAX = 150;  // decoded sprite tiles kept (LRU by tick)
  var MIN_THUMB_WIDTH = 24;    // bars narrower than this stay flat
  var SLOT_ASPECT = 16 / 9;    // layout aspect per slot; draws cover-crop the
                               // real frames, so a mismatch never distorts
  var MIN_SLOT_SECONDS = 0.25; // density ladder bounds (powers of two)
  var MAX_SLOT_SECONDS = 8192;

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
    var entry = { img: img, ready: false, failed: false, tick: ++_tick, span: span };
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

  // Smallest ladder step (power of two × MIN_SLOT_SECONDS) covering slotDur.
  // Snapping keeps tile keys stable across pans; only a zoom that crosses a
  // doubling boundary produces new keys (and thus new fetches).
  function quantizedSlotSeconds(slotDur) {
    var d = MIN_SLOT_SECONDS;
    while (d < slotDur && d < MAX_SLOT_SECONDS) d *= 2;
    return d;
  }

  // Draw one cached tile's slots into the bar. Each slot cover-crops its
  // frame (center crop to the slot's aspect) so frames are never stretched
  // or squashed regardless of ladder rounding.
  function drawTile(ctx, img, entrySpan, barGeom) {
    var cols = CLIPGEN_CONFIG.cardScrubberSpriteCols;
    var rows = CLIPGEN_CONFIG.cardScrubberSpriteRows;
    var frameCount = cols * rows;
    var fw = img.width / cols;
    var fh = img.height / rows;
    if (!fw || !fh) return false;
    var span = entrySpan.end - entrySpan.start;
    var D = barGeom.slotSeconds;
    var i0 = Math.max(0, Math.floor((barGeom.tFrom - entrySpan.start) / D));
    var t = entrySpan.start + i0 * D;
    while (t < entrySpan.end && t < barGeom.tTo) {
      var slotEnd = Math.min(t + D, entrySpan.end);
      var mid = (t + slotEnd) / 2;
      var f = Math.min(
        frameCount - 1,
        Math.max(0, Math.floor(((mid - entrySpan.start) / span) * frameCount))
      );
      var dx = barGeom.x + (t - barGeom.start) * barGeom.pxPerSec;
      var dw = (slotEnd - t) * barGeom.pxPerSec;
      var destAspect = dw / barGeom.h;
      var srcW = Math.min(fw, fh * destAspect);
      var srcH = Math.min(fh, fw / destAspect);
      ctx.drawImage(
        img,
        (f % cols) * fw + (fw - srcW) / 2,
        Math.floor(f / cols) * fh + (fh - srcH) / 2,
        srcW, srcH,
        dx, barGeom.y, dw, barGeom.h
      );
      t += D;
    }
    return true;
  }

  // Called from renderTimeline for every visible marker/cut bar. Cache hits
  // draw immediately; missing tiles enqueue a lazy fetch and leave the flat
  // bar showing. Must stay allocation-light: it runs per bar on every
  // pan/zoom frame.
  function drawMarkerThumbs(ctx, keyStem, start, end, x, y, w, h, tintColor) {
    if (w < MIN_THUMB_WIDTH || end <= start || !state.participant) return false;
    var frameCount =
      CLIPGEN_CONFIG.cardScrubberSpriteCols * CLIPGEN_CONFIG.cardScrubberSpriteRows;
    var pxPerSec = w / (end - start);
    var D = quantizedSlotSeconds((h * SLOT_ASPECT) / pxPerSec);
    var tileSpan = frameCount * D;

    // Clip to the canvas viewport: offscreen stretches of a long bar are
    // neither drawn nor fetched.
    var visX1 = Math.max(x, 0);
    var visX2 = Math.min(x + w, ctx.canvas.width);
    if (visX2 <= visX1) return false;
    var tFrom = start + (visX1 - x) / pxPerSec;
    var tTo = start + (visX2 - x) / pxPerSec;
    var n0 = Math.floor((tFrom - start) / tileSpan);
    var n1 = Math.floor((tTo - start) / tileSpan);

    var barGeom = {
      start: start, x: x, y: y, h: h,
      pxPerSec: pxPerSec, slotSeconds: D, tFrom: tFrom, tTo: tTo,
    };
    var barKey = mediaKey(keyStem, start, end) + "|" + D + "|";
    var drew = false;
    ctx.save();
    ctx.beginPath();
    ctx.rect(x, y, w, h);
    ctx.clip();
    for (var n = n0; n <= n1; n++) {
      var tileStart = start + n * tileSpan;
      var tileEnd = Math.min(tileStart + tileSpan, end);
      if (tileEnd <= tileStart) continue;
      var key = barKey + n;
      var entry = _sprites[key];
      if (!entry) {
        // Don't enqueue mid-drag: every drag frame renders a different span
        // and would leave a trail of junk keys to fetch after the commit.
        if (!state.dragging) {
          _wanted[key] = { start: tileStart, end: tileEnd };
          armFetchTimer();
        }
        continue;
      }
      if (!entry.ready) continue; // failed or still loading
      entry.tick = ++_tick;
      drew = drawTile(ctx, entry.img, entry.span, barGeom) || drew;
    }
    if (drew) {
      // Light lane-color tint so source identity survives the imagery.
      ctx.fillStyle = hexToRgba(tintColor, 0.18);
      ctx.fillRect(visX1, y, visX2 - visX1, h);
    }
    ctx.restore();
    return drew;
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
    var thumbs = qs("#coThumbsToggle input");
    if (thumbs) thumbs.checked = !!state.markerThumbnails;
    var scrub = qs("#coScrubAudioToggle input");
    if (scrub) scrub.checked = !!state.markerAudioScrub;
    var follow = qs("#coFollowToggle input");
    if (follow) follow.checked = !!state.followPlayhead;
  }

  function initMarkerScrub() {
    // change (not click on the label) — a label click already flips the
    // checkbox natively; the handler aligns state and re-syncs it. Hotkeys
    // still route through label.click(), which forwards to the checkbox.
    var thumbs = qs("#coThumbsToggle input");
    if (thumbs) {
      thumbs.addEventListener("change", function () {
        state.markerThumbnails = !state.markerThumbnails;
        syncScrubToggles();
        if (CO.persistLaneUi) CO.persistLaneUi();
        if (CO.renderTimeline) CO.renderTimeline();
      });
    }
    var scrub = qs("#coScrubAudioToggle input");
    if (scrub) {
      scrub.addEventListener("change", function () {
        state.markerAudioScrub = !state.markerAudioScrub;
        syncScrubToggles();
        if (CO.persistLaneUi) CO.persistLaneUi();
        if (!state.markerAudioScrub) scrubHoverEnd();
      });
    }
    var follow = qs("#coFollowToggle input");
    if (follow) {
      follow.addEventListener("change", function () {
        state.followPlayhead = !state.followPlayhead;
        syncScrubToggles();
        if (CO.persistLaneUi) CO.persistLaneUi();
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
