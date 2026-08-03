// Card scrubber: hover-to-scrub on a sprite-sheet thumbnail with synced audio
// snippets and a translucent waveform overlay.
//
// Audio is global: only one snippet plays at a time across all attached cards.
// CSS lives in card-scrubber.css (.waveform-canvas, .waveform-scrim).

(function () {
  // ---- Module-scope state (shared across all attached cards) ----

  var _spriteRaf = 0;
  var _audioCtx = null;
  var _audioBuffers = {};
  var _audioLoading = {};
  var _audioSource = null;
  var _audioGain = null;
  var _audioLastTime = -1;
  var _audioSnippetLen = 0.10;
  var _audioMinDelta = 0.04;
  var _audioFadeIn = 0.005;
  var _audioFadeOut = 0.01;

  var _waveformCache = {};
  var _WAVEFORM_SAMPLES = 200;

  // [{ el, onMove, onLeave }] — used by detachAll() to clean up every attachment.
  var _attached = [];

  // ---- Audio helpers ----

  function getAudioContext() {
    if (!_audioCtx) {
      _audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    }
    if (_audioCtx.state === "suspended") _audioCtx.resume();
    return _audioCtx;
  }

  function loadAudioBuffer(url, key) {
    if (_audioBuffers[key]) return Promise.resolve(_audioBuffers[key]);
    if (_audioLoading[key]) return _audioLoading[key];
    // arrayBuffer (audio) — apiGet only handles JSON, so use fetch directly.
    _audioLoading[key] = fetch(url)
      .then(function (r) { return r.arrayBuffer(); })
      .then(function (buf) { return getAudioContext().decodeAudioData(buf); })
      .then(function (decoded) {
        _audioBuffers[key] = decoded;
        delete _audioLoading[key];
        return decoded;
      })
      .catch(function () {
        delete _audioLoading[key];
        return null;
      });
    return _audioLoading[key];
  }

  function audioScrubAt(audioKey, audioUrl, timeSec) {
    if (Math.abs(timeSec - _audioLastTime) < _audioMinDelta) return;
    _audioLastTime = timeSec;

    var buf = _audioBuffers[audioKey];
    if (!buf) {
      loadAudioBuffer(audioUrl, audioKey);
      return;
    }
    if (timeSec < 0 || timeSec >= buf.duration) return;

    var ctx = getAudioContext();
    var now = ctx.currentTime;

    // Crossfade-out the previous snippet instead of hard-stopping (avoids clicks).
    if (_audioGain) {
      _audioGain.gain.cancelScheduledValues(now);
      _audioGain.gain.setValueAtTime(_audioGain.gain.value, now);
      _audioGain.gain.linearRampToValueAtTime(0, now + _audioFadeOut);
    }
    if (_audioSource) {
      var prev = _audioSource;
      var prevGain = _audioGain;
      setTimeout(function () {
        try { prev.stop(); } catch (_) {}
        prev.disconnect();
        if (prevGain) prevGain.disconnect();
      }, _audioFadeOut * 1000 + 5);
    }

    var gain = ctx.createGain();
    gain.connect(ctx.destination);
    gain.gain.setValueAtTime(0, now);
    gain.gain.linearRampToValueAtTime(1, now + _audioFadeIn);
    gain.gain.setValueAtTime(1, now + _audioSnippetLen - _audioFadeOut);
    gain.gain.linearRampToValueAtTime(0, now + _audioSnippetLen);

    var src = ctx.createBufferSource();
    src.buffer = buf;
    src.connect(gain);
    src.onended = function () {
      src.disconnect();
      gain.disconnect();
      if (_audioSource === src) { _audioSource = null; _audioGain = null; }
    };
    src.start(0, timeSec, _audioSnippetLen);

    _audioSource = src;
    _audioGain = gain;
  }

  function audioScrubStop() {
    _audioLastTime = -1;
    if (_audioSource) {
      var ctx = getAudioContext();
      var now = ctx.currentTime;
      var src = _audioSource;
      var gain = _audioGain;
      if (gain) {
        gain.gain.cancelScheduledValues(now);
        gain.gain.setValueAtTime(gain.gain.value, now);
        gain.gain.linearRampToValueAtTime(0, now + _audioFadeOut);
      }
      setTimeout(function () {
        try { src.stop(); } catch (_) {}
        src.disconnect();
        if (gain) gain.disconnect();
      }, _audioFadeOut * 1000 + 5);
      _audioSource = null;
      _audioGain = null;
    }
  }

  // ---- Waveform overlay ----

  function extractWaveform(audioKey) {
    if (_waveformCache[audioKey]) return _waveformCache[audioKey];
    var buf = _audioBuffers[audioKey];
    if (!buf) return null;
    var raw = buf.getChannelData(0);
    var len = raw.length;
    var bucketSize = Math.floor(len / _WAVEFORM_SAMPLES);
    if (bucketSize < 1) return null;
    var peaks = new Float32Array(_WAVEFORM_SAMPLES);
    var maxPeak = 0;
    for (var i = 0; i < _WAVEFORM_SAMPLES; i++) {
      var start = i * bucketSize;
      var end = start + bucketSize;
      var peak = 0;
      for (var j = start; j < end; j++) {
        var abs = raw[j] < 0 ? -raw[j] : raw[j];
        if (abs > peak) peak = abs;
      }
      peaks[i] = peak;
      if (peak > maxPeak) maxPeak = peak;
    }
    if (maxPeak > 0) {
      for (var k = 0; k < _WAVEFORM_SAMPLES; k++) peaks[k] /= maxPeak;
    }
    _waveformCache[audioKey] = peaks;
    return peaks;
  }

  function getOrCreateWaveformCanvas(mediaEl) {
    var existing = mediaEl.querySelector(".waveform-canvas");
    if (existing) {
      existing.style.display = "";
      var existingScrim = mediaEl.querySelector(".waveform-scrim");
      if (existingScrim) existingScrim.style.display = "";
      return existing;
    }
    var rect = mediaEl.getBoundingClientRect();
    // A not-yet-laid-out element would cache a 0x0 canvas; skip until it has size.
    if (rect.width < 1 || rect.height < 1) return null;
    var scrim = document.createElement("div");
    scrim.className = "waveform-scrim";
    mediaEl.appendChild(scrim);
    var canvas = document.createElement("canvas");
    canvas.className = "waveform-canvas";
    canvas.width = Math.round(rect.width);
    canvas.height = Math.round(rect.height * 0.28);
    mediaEl.appendChild(canvas);
    return canvas;
  }

  // Render the static bars into an offscreen canvas once, cached on the target
  // <canvas> via expando props. The bars never change with the playhead, so this
  // turns a per-frame 200-bar fill loop into a single drawImage() blit.
  function getBarsLayer(canvas, waveformData) {
    var w = canvas.width;
    var h = canvas.height;
    var cached = canvas._barsLayer;
    if (
      cached &&
      cached.width === w &&
      cached.height === h &&
      canvas._barsWave === waveformData
    ) {
      return cached;
    }
    var off = document.createElement("canvas");
    off.width = w;
    off.height = h;
    var octx = off.getContext("2d");
    var barCount = waveformData.length;
    var barW = w / barCount;
    octx.fillStyle = "rgba(255, 255, 255, 0.75)";
    for (var i = 0; i < barCount; i++) {
      var barH = waveformData[i] * h * 0.9;
      if (barH < 1) barH = 1;
      octx.fillRect(i * barW, h - barH, Math.max(barW - 0.5, 0.5), barH);
    }
    canvas._barsLayer = off;
    canvas._barsWave = waveformData;
    return off;
  }

  function drawWaveform(canvas, waveformData, frac) {
    var w = canvas.width;
    var h = canvas.height;
    var x = Math.round(frac * w);
    var bars = getBarsLayer(canvas, waveformData);
    // Nothing moved (playhead pixel + bars unchanged) — skip the redraw.
    if (canvas._lastWaveX === x && canvas._lastWaveBars === bars) return;
    canvas._lastWaveX = x;
    canvas._lastWaveBars = bars;
    var ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, w, h);
    ctx.drawImage(bars, 0, 0);
    ctx.fillStyle = "rgba(255, 255, 255, 0.9)";
    ctx.fillRect(x - 1, 0, 2, h);
  }

  // Hide (but keep cached) the waveform canvas + scrim on an element.
  function clearWaveform(mediaEl) {
    if (!mediaEl) return;
    var canvas = mediaEl.querySelector(".waveform-canvas");
    if (canvas) canvas.style.display = "none";
    var scrim = mediaEl.querySelector(".waveform-scrim");
    if (scrim) scrim.style.display = "none";
  }

  // ---- Public API ----

  // Attach scrubbing to a media element. The consumer is responsible for
  // setting backgroundImage (and any other static styling) on the element;
  // this module owns backgroundSize + backgroundPosition.
  //
  // opts:
  //   spriteData:   { cols, rows, frameCount, interval } — required
  //   audioFile:    string filename used as cache key — optional; enables audio + waveform
  //   audioBaseUrl: string prefix for audio fetch — optional, default "media/"
  //   restFrame:    frame shown before the first hover and after mouseleave —
  //                 optional, default 0. Screenspace's heatmap animations rest on
  //                 the last frame, which is the finished accumulation.
  //   onScrub:      called with (frac, frameIndex) on each scrub step — optional
  //
  // Returns a detach() function that removes listeners and any DOM additions.
  function attach(mediaEl, opts) {
    if (!mediaEl || !opts || !opts.spriteData) return function () {};
    var sd = opts.spriteData;
    var audioFile = opts.audioFile || null;
    var audioBaseUrl = opts.audioBaseUrl || "media/";
    // Consumers with a query-string audio endpoint pass an explicit audioUrl so we
    // don't encodeURIComponent the whole URL (which would corrupt "?start=&end=").
    var audioUrl = opts.audioUrl
      ? opts.audioUrl
      : audioFile
        ? audioBaseUrl + encodeURIComponent(audioFile)
        : null;

    // Map a frame index onto the sprite grid's background-position percentages.
    function framePosition(frameIndex) {
      frameIndex = Math.max(0, Math.min(frameIndex, sd.frameCount - 1));
      var col = frameIndex % sd.cols;
      var row = Math.floor(frameIndex / sd.cols);
      var xPct = sd.cols > 1 ? (col / (sd.cols - 1)) * 100 : 0;
      var yPct = sd.rows > 1 ? (row / (sd.rows - 1)) * 100 : 0;
      return xPct + "% " + yPct + "%";
    }

    var restFrame = opts.restFrame || 0;
    mediaEl.style.backgroundSize = (sd.cols * 100) + "% " + (sd.rows * 100) + "%";
    mediaEl.style.backgroundPosition = framePosition(restFrame);

    function onMove(e) {
      var clientX = e.clientX;
      if (_spriteRaf) return;
      _spriteRaf = requestAnimationFrame(function () {
        _spriteRaf = 0;
        var rect = mediaEl.getBoundingClientRect();
        var frac = (clientX - rect.left) / rect.width;
        var frameIndex = Math.floor(frac * sd.frameCount);
        frameIndex = Math.max(0, Math.min(frameIndex, sd.frameCount - 1));
        mediaEl.style.backgroundPosition = framePosition(frameIndex);
        if (opts.onScrub) opts.onScrub(Math.max(0, Math.min(frac, 1)), frameIndex);
        if (audioFile && audioUrl) {
          audioScrubAt(audioFile, audioUrl, frameIndex * sd.interval);
          var waveform = extractWaveform(audioFile);
          if (waveform) {
            var wfCanvas = getOrCreateWaveformCanvas(mediaEl);
            if (wfCanvas) drawWaveform(wfCanvas, waveform, frac);
          }
        }
      });
    }

    function onLeave() {
      if (sd) mediaEl.style.backgroundPosition = framePosition(restFrame);
      if (opts.onScrub) opts.onScrub(null, restFrame);
      audioScrubStop();
      clearWaveform(mediaEl);
    }

    mediaEl.addEventListener("mousemove", onMove);
    mediaEl.addEventListener("mouseleave", onLeave);

    var entry = { el: mediaEl, onMove: onMove, onLeave: onLeave };
    _attached.push(entry);

    return function detach() {
      mediaEl.removeEventListener("mousemove", onMove);
      mediaEl.removeEventListener("mouseleave", onLeave);
      var canvas = mediaEl.querySelector(".waveform-canvas");
      if (canvas) canvas.remove();
      var scrim = mediaEl.querySelector(".waveform-scrim");
      if (scrim) scrim.remove();
      var idx = _attached.indexOf(entry);
      if (idx >= 0) _attached.splice(idx, 1);
    };
  }

  // Detach every attachment and clear shared audio state. Useful when toggling
  // a global "fancy cards" switch off, or before a wholesale UI rebuild.
  function detachAll() {
    var copy = _attached.slice();
    for (var i = 0; i < copy.length; i++) {
      var entry = copy[i];
      entry.el.removeEventListener("mousemove", entry.onMove);
      entry.el.removeEventListener("mouseleave", entry.onLeave);
      var canvas = entry.el.querySelector(".waveform-canvas");
      if (canvas) canvas.remove();
      var scrim = entry.el.querySelector(".waveform-scrim");
      if (scrim) scrim.remove();
    }
    _attached.length = 0;
    audioScrubStop();
  }

  // Detach only attachments whose element has left the DOM. Consumers that
  // rebuild a card list (e.g. Studio re-rendering a queue's innerHTML) call this
  // before re-attaching so the _attached array doesn't leak stale entries.
  function detachStale() {
    for (var i = _attached.length - 1; i >= 0; i--) {
      var entry = _attached[i];
      if (!entry.el.isConnected) {
        entry.el.removeEventListener("mousemove", entry.onMove);
        entry.el.removeEventListener("mouseleave", entry.onLeave);
        _attached.splice(i, 1);
      }
    }
  }

  // Stop in-flight audio without detaching. Useful on transient UI events
  // (sidebar resize, modal open) that should silence playback but keep cards
  // wired up.
  function stopAll() {
    audioScrubStop();
  }

  window.clipgenCardScrubber = {
    attach: attach,
    detachAll: detachAll,
    detachStale: detachStale,
    stopAll: stopAll,
    // Primitives — let a consumer with its own hover handler (e.g. the viewer's
    // <video>-seek scrub) drive audio + waveform without a second attach().
    loadAudioBuffer: loadAudioBuffer,
    audioScrubAt: audioScrubAt,
    audioScrubStop: audioScrubStop,
    extractWaveform: extractWaveform,
    getOrCreateWaveformCanvas: getOrCreateWaveformCanvas,
    drawWaveform: drawWaveform,
    clearWaveform: clearWaveform,
  };
})();
