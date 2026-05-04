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
    var scrim = document.createElement("div");
    scrim.className = "waveform-scrim";
    mediaEl.appendChild(scrim);
    var canvas = document.createElement("canvas");
    canvas.className = "waveform-canvas";
    var rect = mediaEl.getBoundingClientRect();
    canvas.width = Math.round(rect.width);
    canvas.height = Math.round(rect.height * 0.28);
    mediaEl.appendChild(canvas);
    return canvas;
  }

  function drawWaveform(canvas, waveformData, frac) {
    var ctx = canvas.getContext("2d");
    var w = canvas.width;
    var h = canvas.height;
    ctx.clearRect(0, 0, w, h);
    var barCount = waveformData.length;
    var barW = w / barCount;
    ctx.fillStyle = "rgba(255, 255, 255, 0.75)";
    for (var i = 0; i < barCount; i++) {
      var barH = waveformData[i] * h * 0.9;
      if (barH < 1) barH = 1;
      ctx.fillRect(i * barW, h - barH, Math.max(barW - 0.5, 0.5), barH);
    }
    var x = Math.round(frac * w);
    ctx.fillStyle = "rgba(255, 255, 255, 0.9)";
    ctx.fillRect(x - 1, 0, 2, h);
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
  //
  // Returns a detach() function that removes listeners and any DOM additions.
  function attach(mediaEl, opts) {
    if (!mediaEl || !opts || !opts.spriteData) return function () {};
    var sd = opts.spriteData;
    var audioFile = opts.audioFile || null;
    var audioBaseUrl = opts.audioBaseUrl || "media/";
    var audioUrl = audioFile ? (audioBaseUrl + encodeURIComponent(audioFile)) : null;

    mediaEl.style.backgroundSize = (sd.cols * 100) + "% " + (sd.rows * 100) + "%";
    mediaEl.style.backgroundPosition = "0% 0%";

    function onMove(e) {
      var clientX = e.clientX;
      if (_spriteRaf) return;
      _spriteRaf = requestAnimationFrame(function () {
        _spriteRaf = 0;
        var rect = mediaEl.getBoundingClientRect();
        var frac = (clientX - rect.left) / rect.width;
        var frameIndex = Math.floor(frac * sd.frameCount);
        frameIndex = Math.max(0, Math.min(frameIndex, sd.frameCount - 1));
        var col = frameIndex % sd.cols;
        var row = Math.floor(frameIndex / sd.cols);
        var xPct = sd.cols > 1 ? (col / (sd.cols - 1)) * 100 : 0;
        var yPct = sd.rows > 1 ? (row / (sd.rows - 1)) * 100 : 0;
        mediaEl.style.backgroundPosition = xPct + "% " + yPct + "%";
        if (audioFile && audioUrl) {
          audioScrubAt(audioFile, audioUrl, frameIndex * sd.interval);
          var waveform = extractWaveform(audioFile);
          if (waveform) drawWaveform(getOrCreateWaveformCanvas(mediaEl), waveform, frac);
        }
      });
    }

    function onLeave() {
      mediaEl.style.backgroundPosition = "0% 0%";
      audioScrubStop();
      var canvas = mediaEl.querySelector(".waveform-canvas");
      if (canvas) canvas.style.display = "none";
      var scrim = mediaEl.querySelector(".waveform-scrim");
      if (scrim) scrim.style.display = "none";
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

  // Stop in-flight audio without detaching. Useful on transient UI events
  // (sidebar resize, modal open) that should silence playback but keep cards
  // wired up.
  function stopAll() {
    audioScrubStop();
  }

  window.clipgenCardScrubber = {
    attach: attach,
    detachAll: detachAll,
    stopAll: stopAll,
  };
})();
