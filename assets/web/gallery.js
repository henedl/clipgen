/* clipgen Gallery Viewer – gallery.js */
(function () {
  "use strict";

  var data = window.CLIPGEN_DATA || null;
  var state = { artifacts: [], lightboxIndex: -1 };
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

    for (var i = 0; i < state.artifacts.length; i++) {
      var a = state.artifacts[i];
      var card = document.createElement("div");
      card.className = "gallery-card";
      card.setAttribute("data-index", i);

      var img = document.createElement("img");
      img.src = a.data || a.file;
      img.alt = a.timestamp_formatted || formatTime(a.timestamp);
      img.loading = "lazy";
      card.appendChild(img);

      var overlay = document.createElement("span");
      overlay.className = "timestamp-overlay";
      overlay.textContent = a.timestamp_formatted || formatTime(a.timestamp);
      card.appendChild(overlay);

      grid.appendChild(card);
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

    document.addEventListener("keydown", function (e) {
      if (state.lightboxIndex < 0) return;
      if (e.key === "Escape") closeLightbox();
      else if (e.key === "ArrowLeft") navigateLightbox(-1);
      else if (e.key === "ArrowRight") navigateLightbox(1);
    });
  }

  function openLightbox(index) {
    if (index < 0 || index >= state.artifacts.length) return;
    state.lightboxIndex = index;

    var a = state.artifacts[index];
    var content = qs("#lightboxContent");
    content.innerHTML = "";
    var img = document.createElement("img");
    img.src = a.data || a.file;
    img.alt = a.timestamp_formatted || formatTime(a.timestamp);
    content.appendChild(img);

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
    var next = state.lightboxIndex + delta;
    if (next < 0) next = state.artifacts.length - 1;
    if (next >= state.artifacts.length) next = 0;
    openLightbox(next);
  }
})();
