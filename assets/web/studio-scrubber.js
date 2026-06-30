/* clipgen Studio — card-scrubber satellite (opt-in: STUDIO_CARD_SCRUBBER).
 *
 * Carved out of studio.js (the hub) following the hub+satellite convention used
 * by studio-intake.js. Hover-to-scrub for queue/intake cards: sprite-sheet
 * prefetch + warming and per-card wiring onto window.clipgenCardScrubber. The
 * cluster is self-contained — its only cross-file dependency is the hub's
 * mutable `state` (read for state.cardScrubberEnabled), reached through the
 * window.ClipgenStudio (STUDIO) namespace; CLIPGEN_CONFIG and
 * window.clipgenCardScrubber are ambient globals reached via the scope chain.
 *
 * The hub calls back in via same-named guarded delegators (attachQueueScrubbers,
 * resetScrubberPrefetch). Loaded by studio.html after studio.js and BEFORE
 * studio-intake.js, which destructures STUDIO.attachQueueScrubbers at load time.
 */

(function () {
  "use strict";

  var STUDIO = window.ClipgenStudio;
  var state = STUDIO.state;

  function scrubMediaUrl(kind, participant, start, end) {
    return "api/" + kind + "/" + encodeURIComponent(participant) +
      "?start=" + start + "&end=" + end;
  }

  // Background sprite-sheet warming so the first hover is instant (audio is
  // cheap and loads on demand; the sprite is the slow part — many ffmpeg frame
  // samples + tile). Throttled to a couple of passes at a time, and only fed
  // cards whose thumbnail loaded, so large intake lists don't flood the server.
  var _spritePrefetchQueue = [];
  var _spritePrefetchActive = 0;
  var SPRITE_PREFETCH_CONCURRENCY = 2;

  function enqueueSpritePrefetch(thumb) {
    if (thumb.dataset.scrubSpriteQueued) return;
    thumb.dataset.scrubSpriteQueued = "1";
    _spritePrefetchQueue.push(thumb);
    processSpritePrefetch();
  }

  function processSpritePrefetch() {
    while (_spritePrefetchActive < SPRITE_PREFETCH_CONCURRENCY && _spritePrefetchQueue.length) {
      var thumb = _spritePrefetchQueue.shift();
      if (!thumb.isConnected || thumb.dataset.scrubSpriteLoaded) continue;
      _spritePrefetchActive++;
      loadCardSprite(thumb, function () {
        _spritePrefetchActive--;
        processSpritePrefetch();
      });
    }
  }

  // Preload the sprite sheet and paint it as the thumb background; the
  // .card-scrub-ready CSS rule reveals it under the resting <img> on hover.
  // Idempotent — guards against duplicate in-flight or completed loads.
  function loadCardSprite(thumb, done) {
    if (thumb.dataset.scrubSpriteLoaded || thumb.dataset.scrubSpriteLoading) {
      if (done) done();
      return;
    }
    thumb.dataset.scrubSpriteLoading = "1";
    var spriteUrl = scrubMediaUrl(
      "sprite", thumb.dataset.participant, thumb.dataset.start, thumb.dataset.end,
    );
    var img = new Image();
    img.onload = function () {
      thumb.dataset.scrubSpriteLoading = "";
      thumb.dataset.scrubSpriteLoaded = "1";
      thumb.style.backgroundImage = 'url("' + spriteUrl + '")';
      thumb.classList.add("card-scrub-ready");
      if (done) done();
    };
    img.onerror = function () {
      thumb.dataset.scrubSpriteLoading = "";
      if (done) done();
    };
    img.src = spriteUrl;
  }

  // Wire one queue-card thumbnail for hover scrubbing, but only once its source
  // frame is confirmed to exist. An invalid/out-of-range sheet timestamp still
  // renders a card, yet its thumbnail 404s (→ queue-card-error, the <img> is
  // removed); gating on the thumbnail's successful load keeps us from wiring it
  // and firing sprite/audio requests that would only 404 too.
  function wireCardScrubber(thumb, cols, rows, frameCount) {
    var participant = thumb.dataset.participant;
    var start = thumb.dataset.start;
    var end = thumb.dataset.end;
    if (!participant || start === undefined || end === undefined) return;
    if (thumb.dataset.scrubWired) return;
    var dur = Number(end) - Number(start);
    if (!(dur > 0)) return;
    var img = thumb.querySelector("img");
    if (!img) return;

    function activate() {
      if (thumb.dataset.scrubWired) return;
      thumb.dataset.scrubWired = "1";
      var audioUrl = scrubMediaUrl("clip-audio", participant, start, end);
      window.clipgenCardScrubber.attach(thumb, {
        spriteData: {
          cols: cols,
          rows: rows,
          frameCount: frameCount,
          interval: dur / frameCount,
        },
        audioUrl: audioUrl,
        audioFile: audioUrl, // cache key
        audioBaseUrl: "",
      });
      // Hover loads immediately (jumps the prefetch queue); also warm in the
      // background so a hover that lands before the queue reaches it is instant.
      thumb.addEventListener("mouseenter", function () { loadCardSprite(thumb); });
      enqueueSpritePrefetch(thumb);
    }

    // Activate when the thumbnail has a real frame. A thumbnail that errors is
    // removed from the DOM, so its `load` never fires and the card stays unwired.
    if (img.complete && img.naturalWidth > 0) {
      activate();
    } else {
      img.addEventListener("load", function onLoad() {
        img.removeEventListener("load", onLoad);
        if (img.naturalWidth > 0) activate();
      });
    }
  }

  function attachQueueScrubbers(listEl) {
    if (!state.cardScrubberEnabled || !window.clipgenCardScrubber || !listEl) return;
    window.clipgenCardScrubber.detachStale();
    var cols = CLIPGEN_CONFIG.cardScrubberSpriteCols;
    var rows = CLIPGEN_CONFIG.cardScrubberSpriteRows;
    var frameCount = cols * rows;
    var thumbs = listEl.querySelectorAll(".queue-card-thumb");
    for (var i = 0; i < thumbs.length; i++) {
      wireCardScrubber(thumbs[i], cols, rows, frameCount);
    }
  }

  // Hub resets the prefetch backlog when the scrubber toggle flips (settings
  // apply) before re-rendering, so a stale queue can't re-warm shed cards.
  function resetScrubberPrefetch() {
    _spritePrefetchQueue = [];
  }

  STUDIO.attachQueueScrubbers = attachQueueScrubbers;
  STUDIO.resetScrubberPrefetch = resetScrubberPrefetch;
})();
