/* clipgen Gallery Viewer — gallery.js
 *
 * Same artifact payload contract as the timeline viewer: `window.CLIPGEN_DATA`
 * is injected by the export pipeline (see viewer.py). No live backend.
 *
 * Grid + lightbox: `createGalleryLoopVideo` prefers IntersectionObserver so
 * only in-viewport clips request decode/play; falls back to autoplay when IO
 * is unavailable.
 */
(function () {
  "use strict";

  var data = window.CLIPGEN_DATA || null;
  if (data) clipgenApplyConfig(data.config);
  var state = { artifacts: [], lightboxIndex: -1 };
  var _galleryVideoObserver = null;
  var _galleryVideoObserverBound = false;

  function _ensureGalleryVideoObserver() {
    if (_galleryVideoObserver || typeof IntersectionObserver === "undefined") return;
    _galleryVideoObserver = new IntersectionObserver(function (entries) {
      for (var i = 0; i < entries.length; i++) {
        var ent = entries[i];
        var video = ent.target;
        if (ent.isIntersecting) {
          video.play().catch(function () {});
        } else {
          video.pause();
        }
      }
    }, { rootMargin: "80px" });
  }

  function createGalleryLoopVideo(src, alt) {
    var video = createLoopVideo(src, alt);
    video.autoplay = false;
    video.removeAttribute("autoplay");
    video.preload = "metadata";
    _ensureGalleryVideoObserver();
    if (_galleryVideoObserver) {
      _galleryVideoObserver.observe(video);
    } else {
      video.autoplay = true;
    }
    return video;
  }

  // ---- Init ----

  document.addEventListener("DOMContentLoaded", function () {
    initThemeToggle();

    if (!data || !data.artifacts || data.artifacts.length === 0) {
      var grid = qs("#galleryGrid");
      if (grid) grid.classList.add("hidden");
      var empty = qs("#emptyState");
      if (empty) empty.classList.remove("hidden");
      return;
    }

    state.artifacts = data.artifacts;
    populateHeader();
    renderGrid();
    initLightbox();
  });

  // ---- Header ----

  function populateHeader() {
    var meta = data.meta || {};
    var el;

    el = qs("#sourceVideo");
    if (el && meta.sourceVideo) el.textContent = meta.sourceVideo;

    el = qs("#generatedAt");
    if (el && meta.generatedAt) {
      try {
        var d = new Date(meta.generatedAt);
        el.textContent = d.toLocaleDateString() + " " + d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
      } catch (_) {
        el.textContent = meta.generatedAt;
      }
    }

    el = qs("#artifactCount");
    if (el) {
      var n = state.artifacts.length;
      var label = meta.format === "gif" ? "GIF" : "screenshot";
      el.textContent = n + " " + label + (n !== 1 ? "s" : "");
    }

    el = qs("#intervalInfo");
    if (el && meta.interval) {
      el.textContent = "every " + meta.interval + "s";
      if (meta.videoDuration) el.textContent += " \u00b7 " + formatTime(meta.videoDuration) + " total";
    }
  }

  // ---- Grid ----

  function renderGrid() {
    var grid = qs("#galleryGrid");
    if (!grid) return;

    var frag = document.createDocumentFragment();
    for (var i = 0; i < state.artifacts.length; i++) {
      var a = state.artifacts[i];
      var card = document.createElement("div");
      card.className = "gallery-card";
      card.setAttribute("data-index", i);

      var src = a.data || a.file;
      var altText = a.timestamp_formatted || formatTime(a.timestamp);
      if (isVideoLoop(a.file)) {
        card.appendChild(createGalleryLoopVideo(src, altText));
      } else {
        var img = document.createElement("img");
        img.decoding = "async";
        img.src = src;
        img.alt = altText;
        img.loading = "lazy";
        card.appendChild(img);
      }

      var overlay = document.createElement("span");
      overlay.className = "timestamp-overlay";
      overlay.textContent = a.timestamp_formatted || formatTime(a.timestamp);
      card.appendChild(overlay);

      frag.appendChild(card);
    }
    grid.appendChild(frag);

    if (!_galleryVideoObserverBound) {
      _galleryVideoObserverBound = true;
      document.addEventListener("visibilitychange", function () {
        if (!_galleryVideoObserver) return;
        var videos = grid.querySelectorAll("video");
        for (var vi = 0; vi < videos.length; vi++) {
          if (document.hidden) videos[vi].pause();
        }
      });
    }

    grid.addEventListener("click", function (e) {
      var card = e.target.closest(".gallery-card");
      if (!card) return;
      var idx = parseInt(card.getAttribute("data-index"), 10);
      if (!isNaN(idx)) openLightbox(idx);
    });
  }

  // ---- Lightbox ----

  function initLightbox() {
    var lb = qs("#lightbox");
    if (!lb) return;

    qs("#lightboxClose").addEventListener("click", closeLightbox);
    qs("#lightboxPrev").addEventListener("click", function () { navigateLightbox(-1); });
    qs("#lightboxNext").addEventListener("click", function () { navigateLightbox(1); });

    lb.addEventListener("click", function (e) {
      if (e.target === lb) closeLightbox();
    });

    function lightboxOpen() { return state.lightboxIndex >= 0; }
    window.ClipgenHotkeys.register([
      { id: "gallery.prev", when: lightboxOpen, handler: function () { navigateLightbox(-1); } },
      { id: "gallery.next", when: lightboxOpen, handler: function () { navigateLightbox(1); } },
    ]);
    window.ClipgenHotkeys.registerEscape(function () {
      if (!lightboxOpen()) return false;
      closeLightbox();
      return true;
    });
  }

  function openLightbox(index) {
    if (index < 0 || index >= state.artifacts.length) return;
    state.lightboxIndex = index;

    var a = state.artifacts[index];
    var content = qs("#lightboxContent");
    content.innerHTML = "";
    var src = a.data || a.file;
    var altText = a.timestamp_formatted || formatTime(a.timestamp);
    if (isVideoLoop(a.file)) {
      content.appendChild(createLoopVideo(src, altText));
    } else {
      var img = document.createElement("img");
      img.decoding = "async";
      img.src = src;
      img.alt = altText;
      content.appendChild(img);
    }

    var caption = qs("#lightboxCaption");
    if (caption) {
      var parts = [a.timestamp_formatted || formatTime(a.timestamp)];
      if (a.file) parts.push(a.file);
      caption.textContent = parts.join("  \u00b7  ");
    }

    qs("#lightbox").classList.remove("hidden");
  }

  function closeLightbox() {
    state.lightboxIndex = -1;
    qs("#lightbox").classList.add("hidden");
  }

  function navigateLightbox(delta) {
    var len = state.artifacts.length;
    if (!len) return;
    var next = ((state.lightboxIndex + delta) % len + len) % len;
    openLightbox(next);
  }
})();
