/* clipgen Timeline Viewer – viewer.js */
(function () {
  "use strict";

  var data = window.CLIPGEN_DATA || null;

  var state = {
    artifacts: [],
    filtered: [],
    selectedId: null,
    duration: 0,
    listSort: null,
    expandedTracks: {},
  };

  var SORT_DEFAULT_DIR = {
    severity: "desc",
    chrono: "asc",
    duration: "desc",
    alpha: "asc",
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

  var SEVERITY_SORT = {
    "sev-critical": -4,
    "sev-high": -3,
    "sev-medium": -2,
    "sev-low": -1,
    "sev-na": 0,
    "sev-positive": 1,
    "sev-very-positive": 2,
    "sev-unknown": 998,
  };

  function severityClassForLabel(raw) {
    if (!raw || !String(raw).trim()) return "";
    var k = String(raw).trim().toLowerCase();
    var map = {
      critical: "sev-critical",
      high: "sev-high",
      medium: "sev-medium",
      low: "sev-low",
      "n/a": "sev-na",
      positive: "sev-positive",
      "very positive": "sev-very-positive",
    };
    return map[k] || "sev-unknown";
  }

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
      parts.push(severityClassForLabel(sev));
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
      var ca = severityClassForLabel(a);
      var cb = severityClassForLabel(b);
      var na = SEVERITY_SORT.hasOwnProperty(ca) ? SEVERITY_SORT[ca] : 999;
      var nb = SEVERITY_SORT.hasOwnProperty(cb) ? SEVERITY_SORT[cb] : 999;
      if (na !== nb) return na - nb;
      return a.localeCompare(b);
    });
    return labels;
  }

  function listPillClasses(a) {
    var t = markerTypeClass(a.type);
    var parts = ["list-artifact-pill", t];
    var s = (a.severity || "").trim();
    if (s) {
      parts.push(severityClassForLabel(s));
    } else {
      parts.push("type-only");
    }
    return parts.join(" ");
  }

  function listPillText(a) {
    var s = (a.severity || "").trim();
    if (s) return s;
    return (a.type || "clip").toUpperCase();
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
    pillEl.className = "detail-badge detail-severity " + severityClassForLabel(sev);
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
    if (text !== undefined) e.textContent = text;
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

    computeDuration();
    populateHeader();

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
    var categories = uniqueValues("category");
    var participants = uniqueValues("participant");

    fillSelect("#filterCategory", categories, "All categories");
    fillSelect("#filterParticipant", participants, "All participants");

    var severities = sortedUniqueSeverities();
    var wrap = qs("#filterSeverityWrap");
    if (wrap) {
      if (severities.length === 0) {
        wrap.classList.add("hidden");
      } else {
        wrap.classList.remove("hidden");
        fillSelect("#filterSeverity", severities, "All severities");
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
      var sw = el("span", "legend-severity-swatch " + severityClassForLabel(lab));
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
      var cls = severityClassForLabel(a.severity);
      var sevVal = cls && SEVERITY_SORT.hasOwnProperty(cls) ? SEVERITY_SORT[cls] : 999;
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

  function toggleTrackExpand(trackEl) {
    var trackId = trackEl._trackId;
    state.expandedTracks[trackId] = !state.expandedTracks[trackId];
    applyTrackLayout(trackEl);
    var btn = trackEl.parentNode.querySelector(".track-expand-btn");
    if (btn) {
      var expanded = !!state.expandedTracks[trackId];
      btn.classList.toggle("expanded", expanded);
      btn.setAttribute("aria-expanded", expanded ? "true" : "false");
      btn.title = expanded ? "Collapse tracks" : "Expand tracks";
      btn.setAttribute("aria-label", btn.title);
    }
  }

  // ---- Timeline rendering ----

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

      marker.addEventListener("mouseenter", onMarkerHover);
      marker.addEventListener("mousemove", moveTooltip);
      marker.addEventListener("mouseleave", hideTooltip);
      marker.addEventListener("click", function () {
        selectArtifact(a.id);
      });

      track.appendChild(marker);
      markers.push({ el: marker, artifact: a });
    });

    track._trackMarkers = markers;
    track._trackId = "unified";
    applyTrackLayout(track);

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

  function artifactDurationSec(a) {
    var s = Number(a.start);
    var e = Number(a.end);
    if (isNaN(s)) s = 0;
    if (isNaN(e)) e = isNaN(s) ? 0 : s;
    var d = e - s;
    if (isNaN(d) || d < 0) return 0;
    return d;
  }

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
          var ca = severityClassForLabel(a.severity);
          var cb = severityClassForLabel(b.severity);
          var na = SEVERITY_SORT.hasOwnProperty(ca) ? SEVERITY_SORT[ca] : 999;
          var nb = SEVERITY_SORT.hasOwnProperty(cb) ? SEVERITY_SORT[cb] : 999;
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
    if (!key || !SORT_DEFAULT_DIR.hasOwnProperty(key)) return;
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

  function renderList() {
    var list = qs("#artifactList");
    if (!list) return;
    list.innerHTML = "";

    orderedArtifactsForList().forEach(function (a) {
      var li = document.createElement("li");
      li.dataset.id = a.id;

      var pill = el("span", listPillClasses(a), listPillText(a));
      pill.title = listPillText(a);
      li.appendChild(pill);

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
    if (state.selectedId) {
      var sel = document.querySelector(
        '#artifactList li[data-id="' + state.selectedId + '"]'
      );
      if (sel) sel.classList.add("selected");
    }
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
      var storedZ = m.dataset.collapsedZ;
      m.style.zIndex = storedZ || "";
    });
    qsa("#artifactList li.selected").forEach(function (li) {
      li.classList.remove("selected");
    });

    var marker = document.querySelector('.artifact-marker[data-id="' + id + '"]');
    if (marker) {
      marker.classList.add("selected");
      marker.style.zIndex = 1001;
    }
    var li = document.querySelector('#artifactList li[data-id="' + id + '"]');
    if (li) {
      li.classList.add("selected");
      var sidebar = qs("#sidebar");
      if (sidebar) {
        var sRect = sidebar.getBoundingClientRect();
        var lRect = li.getBoundingClientRect();
        if (lRect.top < sRect.top) {
          sidebar.scrollTop += lRect.top - sRect.top;
        } else if (lRect.bottom > sRect.bottom) {
          sidebar.scrollTop += lRect.bottom - sRect.bottom;
        }
      }
    }

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
    qsa(".artifact-marker.selected").forEach(function (m) {
      m.classList.remove("selected");
      var storedZ = m.dataset.collapsedZ;
      m.style.zIndex = storedZ || "";
    });
    qsa("#artifactList li.selected").forEach(function (li) {
      li.classList.remove("selected");
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

    applySeverityPill(qs("#detailSeverityPill"), a);

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
    if ((a.severity || "").trim()) {
      html += "<br>" + escHtml(a.severity);
    }

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

    var container = qs("#participantRows");
    if (!container) return;
    container.innerHTML = "";

    participantOrder.forEach(function (pid) {
      var row = el("div", "participant-row");

      var label = el("div", "participant-label", pid);
      row.appendChild(label);

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

        marker.addEventListener("mouseenter", onMarkerHover);
        marker.addEventListener("mousemove", moveTooltip);
        marker.addEventListener("mouseleave", hideTooltip);
        marker.addEventListener("click", function () {
          selectArtifact(a.id);
        });

        track.appendChild(marker);
        markers.push({ el: marker, artifact: a });
      });

      track._trackMarkers = markers;
      track._trackId = "participant-" + pid;
      applyTrackLayout(track);

      row.appendChild(track);

      var expandBtn = el("button", "track-expand-btn");
      expandBtn.type = "button";
      expandBtn.setAttribute("aria-label", "Expand tracks");
      expandBtn.setAttribute("aria-expanded", "false");
      expandBtn.title = "Expand tracks";
      expandBtn.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><polyline points="6 9 12 15 18 9"/></svg>';
      expandBtn.addEventListener("click", (function (t) {
        return function () { toggleTrackExpand(t); };
      })(track));
      row.appendChild(expandBtn);

      container.appendChild(row);
    });

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
      parts.push(formatTime(a.start) + (a.end != null ? " – " + formatTime(a.end) : ""));
      if (a.category) parts.push(escHtml(a.category));
      metaEl.innerHTML = parts.join("&ensp;·&ensp;");
    }

    var preview = qs("#playerPreview");
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
      vid.autoplay = true;
      vid.src = a.file;
      preview.appendChild(vid);
    }
  }
})();
