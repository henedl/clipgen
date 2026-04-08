/* clipgen Insights Viewer */

(function () {
  "use strict";

  var THEME_STORAGE_KEY = "clipgen-insights-viewer-theme";

  var data = null;
  var artifactMap = {};

  // ---- Helpers ----

  function countArtifacts(insight) {
    var c = (insight.causes || {}).artifacts || [];
    var b = (insight.behaviors || {}).artifacts || [];
    var i = (insight.impacts || {}).artifacts || [];
    return c.length + b.length + i.length;
  }

  // ---- Theme ----

  function initThemeToggle() {
    var stored = null;
    try {
      stored = window.localStorage.getItem(THEME_STORAGE_KEY);
    } catch (_) {}
    var root = document.documentElement;
    if (stored === "light" || stored === "dark") {
      root.setAttribute("data-theme", stored);
    }
    updateThemeButton(stored);

    var btn = qs("#themeToggle");
    if (btn)
      btn.addEventListener("click", function () {
        var current = root.getAttribute("data-theme");
        var next;
        if (current === "dark") next = "light";
        else if (current === "light") next = "dark";
        else {
          try {
            next =
              window.matchMedia &&
              window.matchMedia("(prefers-color-scheme: dark)").matches
                ? "light"
                : "dark";
          } catch (_) {
            next = "dark";
          }
        }
        root.setAttribute("data-theme", next);
        try {
          window.localStorage.setItem(THEME_STORAGE_KEY, next);
        } catch (_) {}
        updateThemeButton(next);
      });
  }

  function updateThemeButton(explicitTheme) {
    var btn = qs("#themeToggle");
    if (!btn) return;
    var effective = explicitTheme;
    if (effective !== "light" && effective !== "dark") {
      var prefersDark = false;
      try {
        prefersDark =
          window.matchMedia &&
          window.matchMedia("(prefers-color-scheme: dark)").matches;
      } catch (_) {}
      effective = prefersDark ? "dark" : "light";
    }
    btn.setAttribute("data-theme", effective);
    btn.setAttribute("aria-pressed", effective === "dark" ? "true" : "false");
  }

  // ---- Render index ----

  function renderIndex() {
    var container = qs("#insightsIndex");
    container.innerHTML = "";

    if (!data.insights || data.insights.length === 0) {
      container.appendChild(
        el("div", "empty-state", "No insights to display.")
      );
      return;
    }

    var grid = el("div", "index-grid");
    for (var i = 0; i < data.insights.length; i++) {
      grid.appendChild(createIndexCard(data.insights[i]));
    }
    container.appendChild(grid);
  }

  function createIndexCard(insight) {
    var card = el("div", "index-card");
    card.dataset.insightId = insight.id;

    var header = el("div", "index-card-header");
    header.appendChild(el("span", "index-card-title", insight.title));
    if (insight.severity) {
      header.appendChild(
        el("span", "sev-pill " + severityClass(insight.severity), insight.severity)
      );
    }
    card.appendChild(header);

    if (insight.summary) {
      card.appendChild(
        el("div", "index-card-summary", truncate(insight.summary, 120))
      );
    }

    var count = countArtifacts(insight);
    card.appendChild(
      el("div", "index-card-count", count + " artifact" + (count !== 1 ? "s" : ""))
    );

    card.addEventListener("click", function () {
      showDetail(insight);
    });
    return card;
  }

  // ---- Render detail ----

  function showDetail(insight) {
    qs("#insightsIndex").classList.add("hidden");
    var detail = qs("#insightDetail");
    detail.classList.remove("hidden");
    detail.innerHTML = "";

    // Back link
    var back = el("a", "detail-back", "\u2190 Back to all insights");
    back.addEventListener("click", function (e) {
      e.preventDefault();
      detail.classList.add("hidden");
      qs("#insightsIndex").classList.remove("hidden");
    });
    detail.appendChild(back);

    // Title
    detail.appendChild(el("h2", "detail-title", insight.title));

    // Meta
    var meta = el("div", "detail-meta");
    if (insight.severity) {
      meta.appendChild(
        el("span", "sev-pill " + severityClass(insight.severity), insight.severity)
      );
    }
    detail.appendChild(meta);

    // Timeline context
    if (insight.timelineContext) {
      detail.appendChild(
        el("div", "detail-timeline-context", insight.timelineContext)
      );
    }

    // Summary
    if (insight.summary) {
      detail.appendChild(el("div", "detail-summary", insight.summary));
    }

    // Buckets
    var buckets = [
      { key: "causes", label: "Causes" },
      { key: "behaviors", label: "Behaviors" },
      { key: "impacts", label: "Impacts" },
    ];
    for (var i = 0; i < buckets.length; i++) {
      var bucket = insight[buckets[i].key];
      if (!bucket) continue;
      var hasContent =
        (bucket.narrative && bucket.narrative.trim()) ||
        (bucket.artifacts && bucket.artifacts.length > 0);
      if (!hasContent) continue;

      var section = el("div", "detail-bucket");
      section.appendChild(
        el(
          "div",
          "detail-bucket-label detail-bucket-label-" + buckets[i].key,
          buckets[i].label
        )
      );

      if (bucket.narrative && bucket.narrative.trim()) {
        section.appendChild(el("div", "detail-narrative", bucket.narrative));
      }

      if (bucket.artifacts && bucket.artifacts.length > 0) {
        var grid = el("div", "detail-artifacts-grid");
        for (var j = 0; j < bucket.artifacts.length; j++) {
          var art = artifactMap[bucket.artifacts[j]];
          if (art) grid.appendChild(createDetailArtifactCard(art));
        }
        section.appendChild(grid);
      }

      detail.appendChild(section);
    }
  }

  function createDetailArtifactCard(artifact) {
    var card = el("div", "detail-artifact-card");

    // Media
    var media = el("div", "detail-artifact-media");
    if (artifact.type === "clip") {
      var video = document.createElement("video");
      video.src = artifact.file;
      video.controls = true;
      video.preload = "metadata";
      media.appendChild(video);
    } else {
      var img = document.createElement("img");
      img.src = artifact.file;
      img.alt = artifact.description || "";
      img.loading = "lazy";
      media.appendChild(img);
    }
    card.appendChild(media);

    // Info
    var info = el("div", "detail-artifact-info");
    info.appendChild(el("span", "badge", artifact.participant));
    if (artifact.category) info.appendChild(el("span", "badge", artifact.category));
    info.appendChild(
      el("div", "detail-artifact-desc", artifact.description || "")
    );
    info.appendChild(
      el(
        "div",
        "detail-artifact-time",
        formatTime(artifact.start) + " \u2013 " + formatTime(artifact.end)
      )
    );
    card.appendChild(info);

    return card;
  }

  // ---- Init ----

  function init() {
    initThemeToggle();

    data = window.CLIPGEN_DATA;
    if (!data) {
      qs("#insightsIndex").appendChild(
        el("div", "empty-state", "No data available.")
      );
      return;
    }

    // Build artifact lookup
    var artifacts = data.artifacts || [];
    for (var i = 0; i < artifacts.length; i++) {
      artifactMap[artifacts[i].id] = artifacts[i];
    }

    // Header
    if (data.meta) {
      if (data.meta.study) {
        qs("#studyTitle").innerHTML =
          data.meta.study +
          ' <span class="header-light">Insights</span>';
      }
      var metaParts = [];
      if (data.meta.generatedAt) {
        try {
          metaParts.push(new Date(data.meta.generatedAt).toLocaleDateString());
        } catch (_) {}
      }
      if (data.insights) metaParts.push(data.insights.length + " insights");
      qs("#headerMeta").textContent = metaParts.join(" \u00B7 ");
    }

    // Footer link to timeline viewer
    if (data.meta && data.meta.timelineViewerFile) {
      qs("#timelineLink").href = data.meta.timelineViewerFile;
      qs("#viewerFooter").classList.remove("hidden");
    }

    renderIndex();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
