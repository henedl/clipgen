/* clipgen Timeline Viewer – viewer.js */
(function () {
  "use strict";

  var data = window.CLIPGEN_DATA || null;

  var state = {
    artifacts: [],
    filtered: [],
    selectedId: null,
    duration: 0,
  };

  // ---- Helpers ----

  function formatTime(sec) {
    if (sec == null || isNaN(sec)) return "--:--";
    sec = Math.round(sec);
    var h = Math.floor(sec / 3600);
    var m = Math.floor((sec % 3600) / 60);
    var s = sec % 60;
    if (h > 0) {
      return h + ":" + pad2(m) + ":" + pad2(s);
    }
    return m + ":" + pad2(s);
  }

  function pad2(n) {
    return n < 10 ? "0" + n : "" + n;
  }

  function qs(sel) {
    return document.querySelector(sel);
  }

  function qsa(sel) {
    return document.querySelectorAll(sel);
  }

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text) e.textContent = text;
    return e;
  }

  // ---- Initialization ----

  var THEME_STORAGE_KEY = "clipgen-viewer-theme";

  document.addEventListener("DOMContentLoaded", function () {
    initThemeToggle();

    if (!data || !data.artifacts) {
      showEmptyState();
      return;
    }

    state.artifacts = (data.artifacts || []).map(function (a, i) {
      return Object.assign({}, a, { _idx: i });
    });

    if (state.artifacts.length === 0) {
      showEmptyState();
      return;
    }

    var presentTypes = derivePresentTypes(state.artifacts);
    state.presentTypes = presentTypes;

    initTypeLegend(presentTypes);
    initTypeFilters(presentTypes);

    computeDuration();
    populateHeader();
    populateFilters();
    applyFilters();
    renderTimeline();
    renderList();
    bindFilterEvents();
  });

  function initThemeToggle() {
    applyStoredThemePreference();
    var btn = qs("#themeToggle");
    if (!btn) return;
    btn.addEventListener("click", function () {
      toggleThemePreference();
    });
  }

  function applyStoredThemePreference() {
    var stored = null;
    try {
      stored = window.localStorage.getItem(THEME_STORAGE_KEY);
    } catch (_) {}
    var root = document.documentElement;
    if (stored === "light" || stored === "dark") {
      root.setAttribute("data-theme", stored);
    } else {
      root.removeAttribute("data-theme");
    }
    updateThemeToggleButton(stored);
  }

  function toggleThemePreference() {
    var root = document.documentElement;
    var current = root.getAttribute("data-theme");
    var next;
    if (current === "dark") {
      next = "light";
    } else if (current === "light") {
      next = "dark";
    } else {
      var prefersDark = false;
      try {
        prefersDark =
          window.matchMedia &&
          window.matchMedia("(prefers-color-scheme: dark)").matches;
      } catch (_) {}
      next = prefersDark ? "light" : "dark";
    }
    root.setAttribute("data-theme", next);
    try {
      window.localStorage.setItem(THEME_STORAGE_KEY, next);
    } catch (_) {}
    updateThemeToggleButton(next);
  }

  function updateThemeToggleButton(explicitTheme) {
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
    var categories = uniqueValues("category");
    var participants = uniqueValues("participant");

    fillSelect("#filterCategory", categories, "All categories");
    fillSelect("#filterParticipant", participants, "All participants");
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

  function uniqueValues(field) {
    var seen = {};
    state.artifacts.forEach(function (a) {
      var v = a[field];
      if (v) seen[v] = true;
    });
    return Object.keys(seen).sort();
  }

  function fillSelect(sel, values, allLabel) {
    var select = qs(sel);
    if (!select) return;
    select.innerHTML = "";
    var optAll = document.createElement("option");
    optAll.value = "";
    optAll.textContent = allLabel;
    select.appendChild(optAll);
    values.forEach(function (v) {
      var opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v;
      select.appendChild(opt);
    });
  }

  function bindFilterEvents() {
    var catSel = qs("#filterCategory");
    var partSel = qs("#filterParticipant");
    var typeChecks = qsa("#filterType input[type=checkbox]");

    if (catSel) catSel.addEventListener("change", onFilterChange);
    if (partSel) partSel.addEventListener("change", onFilterChange);
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

    var ids = {};
    state.filtered = state.artifacts.filter(function (a) {
      if (cat && a.category !== cat) return false;
      if (part && a.participant !== part) return false;
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

  // ---- Timeline rendering ----

  function renderTimeline() {
    var track = qs("#timelineTrack");
    if (!track) return;
    track.innerHTML = "";

    state.artifacts.forEach(function (a) {
      var marker = el("div", "artifact-marker " + (a.type || "clip"));
      marker.dataset.id = a.id;

      var startPct = ((a.start || 0) / state.duration) * 100;
      var endSec = a.end || a.start || 0;
      var widthPct = ((endSec - (a.start || 0)) / state.duration) * 100;
      var minWidth = 0.4;
      if (widthPct < minWidth) widthPct = minWidth;
      if (a.type === "screen") widthPct = Math.max(widthPct, 0.5);

      marker.style.left = startPct + "%";
      marker.style.width = widthPct + "%";

      marker.addEventListener("mouseenter", onMarkerHover);
      marker.addEventListener("mousemove", moveTooltip);
      marker.addEventListener("mouseleave", hideTooltip);
      marker.addEventListener("click", function () {
        selectArtifact(a.id);
      });

      track.appendChild(marker);
    });

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

  function renderList() {
    var list = qs("#artifactList");
    if (!list) return;
    list.innerHTML = "";

    state.artifacts.forEach(function (a) {
      var li = document.createElement("li");
      li.dataset.id = a.id;

      var badge = el("span", "list-type-badge " + (a.type || "clip"), (a.type || "clip").toUpperCase());
      li.appendChild(badge);

      var info = el("div", "list-item-info");
      var desc = el("div", "list-item-desc", a.description || "(no description)");
      var meta = el("div", "list-item-meta",
        formatTime(a.start) + (a.end != null ? " – " + formatTime(a.end) : "") +
        (a.category ? "  •  " + a.category : "")
      );
      info.appendChild(desc);
      info.appendChild(meta);
      li.appendChild(info);

      li.addEventListener("click", function () {
        selectArtifact(a.id);
      });
      li.addEventListener("mouseenter", function (ev) {
        showTooltipForArtifact(a, ev);
      });
      li.addEventListener("mousemove", moveTooltip);
      li.addEventListener("mouseleave", hideTooltip);

      list.appendChild(li);
    });

    updateListVisibility();
  }

  function updateListVisibility() {
    qsa("#artifactList li").forEach(function (li) {
      var id = li.dataset.id;
      if (state._filteredIds[id]) {
        li.classList.remove("filtered-out");
      } else {
        li.classList.add("filtered-out");
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

  function selectArtifact(id) {
    if (state.selectedId === id) {
      clearSelection();
      return;
    }
    state.selectedId = id;

    qsa(".artifact-marker.selected").forEach(function (m) {
      m.classList.remove("selected");
    });
    qsa("#artifactList li.selected").forEach(function (li) {
      li.classList.remove("selected");
    });

    var marker = document.querySelector('.artifact-marker[data-id="' + id + '"]');
    if (marker) marker.classList.add("selected");
    var li = document.querySelector('#artifactList li[data-id="' + id + '"]');
    if (li) {
      li.classList.add("selected");
      li.scrollIntoView({ block: "nearest", behavior: "smooth" });
    }

    var artifact = findArtifact(id);
    if (artifact) showDetail(artifact);
  }

  function clearSelection() {
    state.selectedId = null;
    qsa(".artifact-marker.selected").forEach(function (m) {
      m.classList.remove("selected");
    });
    qsa("#artifactList li.selected").forEach(function (li) {
      li.classList.remove("selected");
    });
    var empty = qs("#detailEmpty");
    var content = qs("#detailContent");
    if (empty) empty.classList.remove("hidden");
    if (content) content.classList.add("hidden");
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

    var badge = qs("#detailType");
    if (badge) {
      badge.textContent = (a.type || "clip").toUpperCase();
      badge.className = "detail-badge " + (a.type || "clip");
    }

    setText("#detailDescription", a.description || "(no description)");
    setText("#detailCategory", a.category || "–");
    setText("#detailParticipant", a.participant || "–");
    setText("#detailTime",
      formatTime(a.start) + (a.end != null ? " – " + formatTime(a.end) : ""));
    setText("#detailCell", a.cellA1 || (a.cellRow ? "R" + a.cellRow + "C" + a.cellCol : "–"));
    setText("#detailSource", a.sourceVideo || "–");
    setText("#detailFile", a.file || "–");

    var preview = qs("#detailPreview");
    if (!preview) return;
    preview.innerHTML = "";

    if (!a.file) return;

    if (a.type === "screen") {
      var img = document.createElement("img");
      img.src = a.file;
      img.alt = a.description || "screenshot";
      preview.appendChild(img);
    } else if (a.type === "gif") {
      var gifImg = document.createElement("img");
      gifImg.src = a.file;
      gifImg.alt = a.description || "gif";
      preview.appendChild(gifImg);
    } else {
      var vid = document.createElement("video");
      vid.controls = true;
      vid.preload = "metadata";
      vid.src = a.file;
      preview.appendChild(vid);
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

    var html = "<strong>" + escHtml(a.description || "(no description)") + "</strong><br>";
    html += '<span class="tooltip-time">' + formatTime(a.start);
    if (a.end != null) html += " – " + formatTime(a.end);
    html += "</span>";
    if (a.category) html += "<br>" + escHtml(a.category);
    if (a.participant) html += " · " + escHtml(a.participant);

    tip.innerHTML = html;
    tip.classList.remove("hidden");
    positionTooltip(tip, ev);
  }

  function moveTooltip(ev) {
    var tip = qs("#tooltip");
    if (tip && !tip.classList.contains("hidden")) {
      positionTooltip(tip, ev);
    }
  }

  function positionTooltip(tip, ev) {
    var x = ev.clientX + 12;
    var y = ev.clientY + 12;
    var rect = tip.getBoundingClientRect();
    if (x + rect.width > window.innerWidth - 8) {
      x = ev.clientX - rect.width - 12;
    }
    if (y + rect.height > window.innerHeight - 8) {
      y = ev.clientY - rect.height - 12;
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
})();
