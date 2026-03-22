/* clipgen Insight Builder */

(function () {
  "use strict";

  var THEME_STORAGE_KEY = "clipgen-insights-theme";
  var SIDEBAR_WIDTH_KEY = "clipgen-insights-sidebar-width";
  var SEVERITY_OPTIONS = [
    "",
    "Critical",
    "High",
    "Medium",
    "Low",
    "N/A",
    "Positive",
    "Very Positive",
  ];

  var SORT_DEFAULT_DIR = {
    severity: "desc",
    chrono: "asc",
    duration: "desc",
    alpha: "asc",
  };

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

  var state = {
    artifacts: [],
    filteredArtifacts: [],
    insights: [],
    dirtyIds: {},
    sidebarCollapsed: false,
    expandedInsightIds: {},
    popoverArtifactId: null,
    listSort: null,
  };

  // ---- Helpers ----

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
  function truncate(str, max) {
    if (!str) return "";
    return str.length > max ? str.slice(0, max) + "\u2026" : str;
  }

  function severityClass(raw) {
    if (!raw) return "";
    var k = raw.trim().toLowerCase();
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

  function formatTime(seconds) {
    if (seconds == null) return "";
    var s = Math.round(seconds);
    var m = Math.floor(s / 60);
    var sec = s % 60;
    return m + ":" + (sec < 10 ? "0" : "") + sec;
  }

  function countBucketArtifacts(insight) {
    var c = (insight.causes || {}).artifacts || [];
    var b = (insight.behaviors || {}).artifacts || [];
    var i = (insight.impacts || {}).artifacts || [];
    return c.length + b.length + i.length;
  }

  function findArtifact(id) {
    for (var i = 0; i < state.artifacts.length; i++) {
      if (state.artifacts[i].id === id) return state.artifacts[i];
    }
    return null;
  }

  function showToast(msg) {
    var t = qs("#toast");
    t.textContent = msg;
    t.classList.remove("hidden");
    clearTimeout(t._timer);
    t._timer = setTimeout(function () {
      t.classList.add("hidden");
    }, 3000);
  }

  // ---- Theme ----

  function initThemeToggle() {
    applyStoredTheme();
    var btn = qs("#themeToggle");
    if (btn)
      btn.addEventListener("click", function () {
        toggleTheme();
      });
  }

  function applyStoredTheme() {
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
    updateThemeButton(stored);
  }

  function toggleTheme() {
    var root = document.documentElement;
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

  // ---- Data loading ----

  function loadData() {
    Promise.all([
      fetch("api/artifacts").then(function (r) {
        return r.json();
      }),
      fetch("api/insights").then(function (r) {
        return r.json();
      }),
    ]).then(function (results) {
      state.artifacts = results[0].artifacts || [];
      state.insights = results[1].insights || [];
      state.filteredArtifacts = state.artifacts.slice();
      updateCounts();
      populateFilters();
      applyFilters();
      renderArtifactGrid();
      renderInsightCards();
    });
  }

  function updateCounts() {
    qs("#artifactCount").textContent = state.artifacts.length + " artifacts";
    qs("#insightCount").textContent = state.insights.length + " insights";
  }

  // ---- Filter system (ported from viewer.js) ----

  function uniqueValues(field) {
    var seen = {};
    var vals = [];
    for (var i = 0; i < state.artifacts.length; i++) {
      var v = (state.artifacts[i][field] || "").trim();
      if (v && !seen[v]) {
        seen[v] = true;
        vals.push(v);
      }
    }
    return vals.sort();
  }

  function fillSelect(sel, values, allLabel) {
    while (sel.options.length > 0) sel.remove(0);
    var opt = document.createElement("option");
    opt.value = "";
    opt.textContent = allLabel;
    sel.appendChild(opt);
    for (var i = 0; i < values.length; i++) {
      var o = document.createElement("option");
      o.value = values[i];
      o.textContent = values[i];
      sel.appendChild(o);
    }
  }

  function populateFilters() {
    fillSelect(
      qs("#filterParticipant"),
      uniqueValues("participant"),
      "All participants"
    );
    fillSelect(
      qs("#filterCategory"),
      uniqueValues("category"),
      "All categories"
    );
    var sevs = uniqueValues("severity");
    var sevSel = qs("#filterSeverity");
    fillSelect(sevSel, sevs, "All severities");
    sevSel.parentElement.style.display = sevs.length ? "" : "none";

    var types = {};
    for (var i = 0; i < state.artifacts.length; i++)
      types[state.artifacts[i].type] = true;
    var typeFilters = qs("#typeFilters");
    var labels = typeFilters.querySelectorAll("label");
    for (var j = 0; j < labels.length; j++) {
      var cb = labels[j].querySelector("input");
      labels[j].style.display = types[cb.value] ? "" : "none";
    }
  }

  function getActiveTypes() {
    var cbs = qsa("#typeFilters input[type=checkbox]");
    var types = [];
    for (var i = 0; i < cbs.length; i++) {
      if (cbs[i].checked) types.push(cbs[i].value);
    }
    return types;
  }

  function applyFilters() {
    var participant = qs("#filterParticipant").value;
    var category = qs("#filterCategory").value;
    var severity = qs("#filterSeverity").value;
    var types = getActiveTypes();
    var search = (qs("#filterSearch").value || "").toLowerCase().trim();

    state.filteredArtifacts = state.artifacts.filter(function (a) {
      if (participant && a.participant !== participant) return false;
      if (category && a.category !== category) return false;
      if (severity && a.severity !== severity) return false;
      if (types.indexOf(a.type) === -1) return false;
      if (search) {
        var haystack = [a.description, a.participant, a.category, a.severity]
          .join(" ")
          .toLowerCase();
        if (haystack.indexOf(search) === -1) return false;
      }
      return true;
    });
    renderArtifactGrid();
  }

  function bindFilterEvents() {
    qs("#filterParticipant").addEventListener("change", applyFilters);
    qs("#filterCategory").addEventListener("change", applyFilters);
    qs("#filterSeverity").addEventListener("change", applyFilters);
    var cbs = qsa("#typeFilters input[type=checkbox]");
    for (var i = 0; i < cbs.length; i++)
      cbs[i].addEventListener("change", applyFilters);
    var searchInput = qs("#filterSearch");
    var searchTimer;
    searchInput.addEventListener("input", function () {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(applyFilters, 150);
    });
  }

  // ---- Sorting ----

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
    if (!state.listSort) return state.filteredArtifacts;
    var key = state.listSort.key;
    var dir = state.listSort.dir;
    return state.filteredArtifacts.slice().sort(function (a, b) {
      var r = 0;
      if (key === "severity") {
        var ae = !(a.severity || "").trim();
        var be = !(b.severity || "").trim();
        if (ae && be) r = 0;
        else if (ae) r = 1;
        else if (be) r = -1;
        else {
          var ca = severityClass(a.severity);
          var cb = severityClass(b.severity);
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
      return r;
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
        ? "Sort by chronology: ascending (earliest first)"
        : "Sort by chronology: descending (latest first)";
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
        ? "Sort alphabetically: ascending (A\u2013Z)"
        : "Sort alphabetically: descending (Z\u2013A)";
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
    renderArtifactGrid();
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

  // ---- Artifact grid rendering ----

  function renderArtifactGrid() {
    var grid = qs("#artifactGrid");
    grid.innerHTML = "";
    var ordered = orderedArtifactsForList();
    if (ordered.length === 0) {
      grid.appendChild(el("div", "empty-state", "No artifacts match filters."));
      return;
    }
    for (var i = 0; i < ordered.length; i++) {
      grid.appendChild(createArtifactCard(ordered[i]));
    }
  }

  function createArtifactCard(artifact) {
    var card = el("div", "artifact-card");
    card.setAttribute("draggable", "true");
    card.dataset.artifactId = artifact.id;

    // Media
    var media = el("div", "artifact-media");
    if (artifact.type === "clip" && artifact.thumbnail && artifact.spriteData) {
      var sd = artifact.spriteData;
      media.style.backgroundImage =
        'url("media/' + encodeURIComponent(artifact.thumbnail) + '")';
      media.style.backgroundSize = (sd.cols * 100) + "% " + (sd.rows * 100) + "%";
      media.dataset.spriteData = JSON.stringify(sd);
      if (artifact.file) media.dataset.audioFile = artifact.file;
      media.addEventListener("mousemove", spriteHover);
      media.addEventListener("mouseleave", spriteReset);
    } else {
      var img = document.createElement("img");
      var fileSrc = "media/" + encodeURIComponent(artifact.file);
      img.src = fileSrc;
      img.alt = artifact.description || "";
      img.loading = "lazy";
      media.appendChild(img);
      if (artifact.type === "gif") {
        setupGifHover(card, img, fileSrc);
      }
    }
    card.appendChild(media);

    // Meta
    var meta = el("div", "artifact-meta");
    var badges = el("div", "artifact-badges");
    badges.appendChild(el("span", "badge badge-participant", artifact.participant));
    badges.appendChild(el("span", "badge badge-" + artifact.type, artifact.type));
    if (artifact.category) badges.appendChild(el("span", "badge badge-category", artifact.category));
    var sev = (artifact.severity || "").trim();
    if (sev) badges.appendChild(el("span", "badge badge-severity " + severityClass(sev), sev));
    meta.appendChild(badges);
    meta.appendChild(el("div", "artifact-desc", truncate(artifact.description, 60)));
    meta.appendChild(el("div", "artifact-time", formatTime(artifact.start) + " \u2013 " + formatTime(artifact.end)));
    card.appendChild(meta);

    // Drag
    card.addEventListener("dragstart", function (e) {
      card.classList.add("dragging");
      e.dataTransfer.setData("text/plain", artifact.id);
      e.dataTransfer.effectAllowed = "copy";
    });
    card.addEventListener("dragend", function () {
      card.classList.remove("dragging");
    });

    // Click to add
    card.addEventListener("click", function (e) {
      if (e.target.closest(".artifact-media")) return;
      showAddPopover(e, artifact.id);
    });
    media.addEventListener("click", function (e) {
      e.stopPropagation();
      showAddPopover(e, artifact.id);
    });

    return card;
  }

  // ---- Sprite hover-to-scrub ----

  var _spriteRaf = 0;
  function spriteHover(e) {
    var target = e.currentTarget;
    var clientX = e.clientX;
    if (_spriteRaf) return;
    _spriteRaf = requestAnimationFrame(function () {
      _spriteRaf = 0;
      var sd;
      try {
        sd = JSON.parse(target.dataset.spriteData);
      } catch (_) {
        return;
      }
      var rect = target.getBoundingClientRect();
      var frac = (clientX - rect.left) / rect.width;
      var frameIndex = Math.floor(frac * sd.frameCount);
      frameIndex = Math.max(0, Math.min(frameIndex, sd.frameCount - 1));
      var col = frameIndex % sd.cols;
      var row = Math.floor(frameIndex / sd.cols);
      var xPct = sd.cols > 1 ? (col / (sd.cols - 1)) * 100 : 0;
      var yPct = sd.rows > 1 ? (row / (sd.rows - 1)) * 100 : 0;
      target.style.backgroundPosition = xPct + "% " + yPct + "%";
      audioScrubAt(target, frameIndex * sd.interval);
    });
  }

  function spriteReset(e) {
    e.currentTarget.style.backgroundPosition = "0% 0%";
    audioScrubStop();
  }

  // ---- Audio scrub (Web Audio API) ----

  var _audioCtx = null;
  var _audioBuffers = {};
  var _audioLoading = {};
  var _audioSource = null;
  var _audioLastTime = -1;
  var _audioSnippetLen = 0.08;
  var _audioMinDelta = 0.04;

  function getAudioContext() {
    if (!_audioCtx) {
      _audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    }
    if (_audioCtx.state === "suspended") _audioCtx.resume();
    return _audioCtx;
  }

  function loadAudioBuffer(filePath) {
    if (_audioBuffers[filePath]) return Promise.resolve(_audioBuffers[filePath]);
    if (_audioLoading[filePath]) return _audioLoading[filePath];
    _audioLoading[filePath] = fetch("media/" + encodeURIComponent(filePath))
      .then(function (r) { return r.arrayBuffer(); })
      .then(function (buf) { return getAudioContext().decodeAudioData(buf); })
      .then(function (decoded) {
        _audioBuffers[filePath] = decoded;
        delete _audioLoading[filePath];
        return decoded;
      })
      .catch(function () {
        delete _audioLoading[filePath];
        return null;
      });
    return _audioLoading[filePath];
  }

  function audioScrubAt(mediaEl, timeSec) {
    if (Math.abs(timeSec - _audioLastTime) < _audioMinDelta) return;
    _audioLastTime = timeSec;

    var filePath = mediaEl.dataset.audioFile;
    if (!filePath) return;

    var buf = _audioBuffers[filePath];
    if (!buf) {
      loadAudioBuffer(filePath);
      return;
    }

    if (timeSec < 0 || timeSec >= buf.duration) return;

    var ctx = getAudioContext();
    if (_audioSource) {
      try { _audioSource.stop(); } catch (_) {}
      _audioSource.disconnect();
    }
    _audioSource = ctx.createBufferSource();
    _audioSource.buffer = buf;
    _audioSource.connect(ctx.destination);
    _audioSource.start(0, timeSec, _audioSnippetLen);
  }

  function audioScrubStop() {
    _audioLastTime = -1;
    if (_audioSource) {
      try { _audioSource.stop(); } catch (_) {}
      _audioSource.disconnect();
      _audioSource = null;
    }
  }

  // ---- GIF hover-to-play ----

  function setupGifHover(card, img, gifSrc) {
    var staticSrc = "";
    var captured = false;

    function captureFirstFrame() {
      if (captured) return;
      captured = true;
      try {
        var canvas = document.createElement("canvas");
        canvas.width = img.naturalWidth;
        canvas.height = img.naturalHeight;
        canvas.getContext("2d").drawImage(img, 0, 0);
        staticSrc = canvas.toDataURL("image/png");
        img.src = staticSrc;
      } catch (_) {}
    }

    img.addEventListener("load", captureFirstFrame, { once: true });
    if (img.complete && img.naturalWidth > 0) captureFirstFrame();

    card.addEventListener("mouseenter", function () {
      if (staticSrc) img.src = gifSrc;
    });
    card.addEventListener("mouseleave", function () {
      if (staticSrc) img.src = staticSrc;
    });
  }

  // ---- Click-to-add popover ----

  function showAddPopover(e, artifactId) {
    e.stopPropagation();
    var popover = qs("#addPopover");
    state.popoverArtifactId = artifactId;

    // Determine target insights
    var expandedIds = Object.keys(state.expandedInsightIds);
    var insightSelect = qs("#popoverInsightSelect");
    insightSelect.innerHTML = "";

    if (expandedIds.length > 1) {
      insightSelect.classList.remove("hidden");
      for (var i = 0; i < state.insights.length; i++) {
        var ins = state.insights[i];
        if (!state.expandedInsightIds[ins.id]) continue;
        var btn = el("button", "popover-insight-item", truncate(ins.title, 30));
        btn.dataset.insightId = ins.id;
        btn.addEventListener("click", function () {
          state._popoverTargetInsight = this.dataset.insightId;
          insightSelect.classList.add("hidden");
        });
        insightSelect.appendChild(btn);
      }
    } else {
      insightSelect.classList.add("hidden");
      if (expandedIds.length === 1) {
        state._popoverTargetInsight = expandedIds[0];
      } else if (state.insights.length === 1) {
        state._popoverTargetInsight = state.insights[0].id;
      } else {
        state._popoverTargetInsight = null;
      }
    }

    // Position
    var rect = e.currentTarget.getBoundingClientRect();
    popover.style.top = Math.min(rect.bottom + 4, window.innerHeight - 160) + "px";
    popover.style.left = Math.min(rect.left, window.innerWidth - 170) + "px";
    popover.classList.remove("hidden");
  }

  function onPopoverBucketClick(e) {
    var bucket = e.target.dataset.bucket;
    if (!bucket) return;
    var insightId = state._popoverTargetInsight;
    if (!insightId) {
      showToast("Expand an insight card first.");
      hidePopover();
      return;
    }
    addArtifactToInsight(insightId, bucket, state.popoverArtifactId);
    hidePopover();
  }

  function hidePopover() {
    qs("#addPopover").classList.add("hidden");
    state.popoverArtifactId = null;
  }

  // ---- Insight CRUD ----

  function addArtifactToInsight(insightId, bucket, artifactId) {
    for (var i = 0; i < state.insights.length; i++) {
      if (state.insights[i].id === insightId) {
        var b = state.insights[i][bucket];
        if (!b) {
          state.insights[i][bucket] = { narrative: "", artifacts: [] };
          b = state.insights[i][bucket];
        }
        if (b.artifacts.indexOf(artifactId) === -1) {
          b.artifacts.push(artifactId);
          markDirty(insightId);
          renderInsightCards();
        }
        return;
      }
    }
  }

  function removeArtifactFromBucket(insightId, bucket, artifactId) {
    for (var i = 0; i < state.insights.length; i++) {
      if (state.insights[i].id !== insightId) continue;
      var arts = state.insights[i][bucket].artifacts;
      var idx = arts.indexOf(artifactId);
      if (idx !== -1) {
        arts.splice(idx, 1);
        markDirty(insightId);
        renderInsightCards();
      }
      return;
    }
  }

  function createNewInsight() {
    fetch("api/insights", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ title: "Untitled insight" }),
    })
      .then(function (r) {
        return r.json();
      })
      .then(function (data) {
        if (data.ok) {
          state.insights.push(data.insight);
          state.expandedInsightIds[data.insight.id] = true;
          updateCounts();
          renderInsightCards();
          showToast("Insight created.");
        }
      });
  }

  function deleteInsight(insightId) {
    if (!confirm("Delete this insight? This cannot be undone.")) return;
    fetch("api/insights/" + insightId, { method: "DELETE" })
      .then(function (r) {
        return r.json();
      })
      .then(function (data) {
        if (data.ok) {
          state.insights = state.insights.filter(function (ins) {
            return ins.id !== insightId;
          });
          delete state.dirtyIds[insightId];
          delete state.expandedInsightIds[insightId];
          updateCounts();
          updateDirtyUI();
          renderInsightCards();
          showToast("Insight deleted.");
        }
      });
  }

  // ---- Dirty tracking and saving ----

  function markDirty(insightId) {
    state.dirtyIds[insightId] = true;
    updateDirtyUI();
  }

  function updateDirtyUI() {
    var hasDirty = Object.keys(state.dirtyIds).length > 0;
    qs("#unsavedIndicator").classList.toggle("hidden", !hasDirty);
    qs("#saveAllBtn").disabled = !hasDirty;
    qs("#discardBtn").disabled = !hasDirty;
  }

  function saveAll() {
    var ids = Object.keys(state.dirtyIds);
    if (ids.length === 0) return;
    var promises = ids.map(function (id) {
      var insight = null;
      for (var i = 0; i < state.insights.length; i++) {
        if (state.insights[i].id === id) {
          insight = state.insights[i];
          break;
        }
      }
      if (!insight) return Promise.resolve();
      return fetch("api/insights/" + id, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(insight),
      }).then(function (r) {
        return r.json();
      });
    });
    Promise.all(promises).then(function () {
      state.dirtyIds = {};
      updateDirtyUI();
      showToast("All changes saved.");
    });
  }

  function saveInsight(insightId) {
    var insight = null;
    for (var i = 0; i < state.insights.length; i++) {
      if (state.insights[i].id === insightId) {
        insight = state.insights[i];
        break;
      }
    }
    if (!insight) return;
    fetch("api/insights/" + insightId, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(insight),
    })
      .then(function (r) {
        return r.json();
      })
      .then(function () {
        delete state.dirtyIds[insightId];
        updateDirtyUI();
        showToast("Insight saved.");
      });
  }

  function discardAll() {
    if (!confirm("Discard all unsaved changes?")) return;
    fetch("api/insights")
      .then(function (r) {
        return r.json();
      })
      .then(function (data) {
        state.insights = data.insights || [];
        state.dirtyIds = {};
        updateDirtyUI();
        renderInsightCards();
        showToast("Changes discarded.");
      });
  }

  // ---- Insight card rendering ----

  function renderInsightCards() {
    var container = qs("#insightCards");
    container.innerHTML = "";

    if (state.insights.length === 0) {
      var empty = el("div", "empty-state");
      empty.innerHTML =
        "<p>No insights yet. Click <strong>+ New Insight</strong> to start authoring.</p>";
      container.appendChild(empty);
      return;
    }

    for (var i = 0; i < state.insights.length; i++) {
      container.appendChild(createInsightCard(state.insights[i]));
    }
  }

  function createInsightCard(insight) {
    var isExpanded = !!state.expandedInsightIds[insight.id];
    var card = el("div", "insight-card" + (isExpanded ? " expanded" : ""));
    card.dataset.insightId = insight.id;

    // Collapsed header
    var header = el("div", "insight-collapsed");
    header.appendChild(el("span", "insight-chevron", "\u25B6"));
    header.appendChild(el("span", "insight-collapsed-title", insight.title || "Untitled insight"));

    if (insight.severity) {
      var sevPill = el("span", "sev-pill " + severityClass(insight.severity), insight.severity);
      header.appendChild(sevPill);
    }

    var statusCls = insight.status === "final" ? "status-final" : "status-draft";
    header.appendChild(el("span", "insight-status-badge " + statusCls, insight.status || "draft"));

    var count = countBucketArtifacts(insight);
    header.appendChild(el("span", "insight-collapsed-count", count + " artifact" + (count !== 1 ? "s" : "")));

    header.addEventListener("click", function () {
      if (state.expandedInsightIds[insight.id]) {
        delete state.expandedInsightIds[insight.id];
      } else {
        state.expandedInsightIds[insight.id] = true;
      }
      renderInsightCards();
    });
    card.appendChild(header);

    // Body (only rendered when expanded)
    if (isExpanded) {
      var body = el("div", "insight-body");
      body.style.display = "block";

      // Fields row
      var fields = el("div", "insight-fields");
      var titleInput = document.createElement("input");
      titleInput.type = "text";
      titleInput.value = insight.title || "";
      titleInput.placeholder = "Insight title";
      titleInput.addEventListener("input", function () {
        insight.title = this.value;
        markDirty(insight.id);
      });
      fields.appendChild(titleInput);

      var sevSelect = document.createElement("select");
      for (var s = 0; s < SEVERITY_OPTIONS.length; s++) {
        var opt = document.createElement("option");
        opt.value = SEVERITY_OPTIONS[s];
        opt.textContent = SEVERITY_OPTIONS[s] || "No severity";
        if (SEVERITY_OPTIONS[s] === insight.severity) opt.selected = true;
        sevSelect.appendChild(opt);
      }
      sevSelect.addEventListener("change", function () {
        insight.severity = this.value;
        markDirty(insight.id);
      });
      fields.appendChild(sevSelect);

      var statusSelect = document.createElement("select");
      ["draft", "final"].forEach(function (v) {
        var opt = document.createElement("option");
        opt.value = v;
        opt.textContent = v;
        if (v === insight.status) opt.selected = true;
        statusSelect.appendChild(opt);
      });
      statusSelect.addEventListener("change", function () {
        insight.status = this.value;
        markDirty(insight.id);
      });
      fields.appendChild(statusSelect);
      body.appendChild(fields);

      // Summary
      var summary = document.createElement("textarea");
      summary.className = "insight-summary";
      summary.value = insight.summary || "";
      summary.placeholder = "Describe the key finding and why it matters...";
      summary.addEventListener("input", function () {
        insight.summary = this.value;
        markDirty(insight.id);
      });
      body.appendChild(summary);

      // Buckets
      var buckets = ["causes", "behaviors", "impacts"];
      for (var bi = 0; bi < buckets.length; bi++) {
        body.appendChild(createBucketSection(insight, buckets[bi]));
      }

      // Timeline context
      var tcDiv = el("div", "insight-timeline-context");
      tcDiv.appendChild(el("label", "", "Timeline context"));
      var tcInput = document.createElement("input");
      tcInput.type = "text";
      tcInput.value = insight.timelineContext || "";
      tcInput.placeholder = "e.g. During onboarding (first 5 minutes)";
      tcInput.addEventListener("input", function () {
        insight.timelineContext = this.value;
        markDirty(insight.id);
      });
      tcDiv.appendChild(tcInput);
      body.appendChild(tcDiv);

      // Footer
      var footer = el("div", "insight-footer");
      var saveBtn = el("button", "btn btn-small insight-save-btn", "Save this insight");
      saveBtn.addEventListener("click", function () {
        saveInsight(insight.id);
      });
      footer.appendChild(saveBtn);

      var deleteBtn = el("button", "btn btn-small btn-danger", "Delete insight");
      deleteBtn.addEventListener("click", function () {
        deleteInsight(insight.id);
      });
      footer.appendChild(deleteBtn);
      body.appendChild(footer);

      card.appendChild(body);
    }

    return card;
  }

  // ---- Bucket section ----

  function createBucketSection(insight, bucketName) {
    var section = el("div", "bucket-section");
    section.appendChild(
      el("div", "bucket-label bucket-label-" + bucketName, bucketName)
    );

    var bucket = insight[bucketName] || { narrative: "", artifacts: [] };

    // Narrative
    var narrative = document.createElement("textarea");
    narrative.className = "bucket-narrative";
    var placeholders = {
      causes: "What led to the observed behavior?",
      behaviors: "What did users actually do?",
      impacts: "What were the consequences?",
    };
    narrative.placeholder = placeholders[bucketName] || "";
    narrative.value = bucket.narrative || "";
    narrative.addEventListener("input", function () {
      bucket.narrative = this.value;
      if (!insight[bucketName]) insight[bucketName] = bucket;
      markDirty(insight.id);
    });
    section.appendChild(narrative);

    // Drop zone
    var dropzone = el(
      "div",
      "bucket-dropzone bucket-dropzone-" + bucketName
    );
    dropzone.dataset.insightId = insight.id;
    dropzone.dataset.bucket = bucketName;

    if (bucket.artifacts.length === 0) {
      dropzone.appendChild(el("div", "bucket-empty-hint", "Drag artifacts here"));
    } else {
      for (var i = 0; i < bucket.artifacts.length; i++) {
        var art = findArtifact(bucket.artifacts[i]);
        if (art) {
          dropzone.appendChild(
            createPreviewCard(art, insight.id, bucketName, i)
          );
        }
      }
    }

    // Drop events
    dropzone.addEventListener("dragover", function (e) {
      e.preventDefault();
      e.dataTransfer.dropEffect = "copy";
      this.classList.add("drag-over");
    });
    dropzone.addEventListener("dragleave", function () {
      this.classList.remove("drag-over");
    });
    dropzone.addEventListener("drop", function (e) {
      e.preventDefault();
      this.classList.remove("drag-over");
      var artifactId = e.dataTransfer.getData("text/plain");
      if (!artifactId) return;
      var targetInsightId = this.dataset.insightId;
      var targetBucket = this.dataset.bucket;

      // Check if this is a reorder within the same bucket
      var dragSourceBucket = e.dataTransfer.getData("application/x-bucket");
      var dragSourceInsight = e.dataTransfer.getData("application/x-insight");
      if (
        dragSourceBucket === targetBucket &&
        dragSourceInsight === targetInsightId
      ) {
        // Reorder
        var ins = findInsightById(targetInsightId);
        if (!ins) return;
        var arts = ins[targetBucket].artifacts;
        var oldIdx = arts.indexOf(artifactId);
        if (oldIdx === -1) return;
        arts.splice(oldIdx, 1);

        // Calculate new position based on mouse
        var cards = this.querySelectorAll(".preview-card");
        var insertIdx = arts.length;
        for (var ci = 0; ci < cards.length; ci++) {
          var rect = cards[ci].getBoundingClientRect();
          if (e.clientY < rect.top + rect.height / 2) {
            insertIdx = ci;
            break;
          }
        }
        arts.splice(insertIdx, 0, artifactId);
        markDirty(targetInsightId);
        renderInsightCards();
      } else {
        addArtifactToInsight(targetInsightId, targetBucket, artifactId);
      }
    });

    section.appendChild(dropzone);
    return section;
  }

  function findInsightById(id) {
    for (var i = 0; i < state.insights.length; i++) {
      if (state.insights[i].id === id) return state.insights[i];
    }
    return null;
  }

  // ---- Preview cards (in buckets) ----

  function createPreviewCard(artifact, insightId, bucketName, index) {
    var card = el("div", "preview-card");
    card.setAttribute("draggable", "true");
    card.dataset.artifactId = artifact.id;
    card.dataset.index = index;

    // Media
    var media = el("div", "preview-media");
    if (artifact.type === "clip" && artifact.thumbnail && artifact.spriteData) {
      var sd = artifact.spriteData;
      media.style.backgroundImage =
        'url("media/' + encodeURIComponent(artifact.thumbnail) + '")';
      media.style.backgroundSize = (sd.cols * 100) + "% " + (sd.rows * 100) + "%";
      media.style.backgroundPosition = "0% 0%";
      media.dataset.spriteData = JSON.stringify(sd);
      if (artifact.file) media.dataset.audioFile = artifact.file;
      media.addEventListener("mousemove", spriteHover);
      media.addEventListener("mouseleave", spriteReset);
    } else {
      var img = document.createElement("img");
      var fileSrc = "media/" + encodeURIComponent(artifact.file);
      img.src = fileSrc;
      img.alt = artifact.description || "";
      img.loading = "lazy";
      media.appendChild(img);
      if (artifact.type === "gif") {
        setupGifHover(card, img, fileSrc);
      }
    }
    card.appendChild(media);

    // Remove button
    var removeBtn = el("button", "preview-remove", "\u00D7");
    removeBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      removeArtifactFromBucket(insightId, bucketName, artifact.id);
    });
    card.appendChild(removeBtn);

    // Info
    var info = el("div", "preview-info");
    var badges = el("div", "artifact-badges");
    badges.appendChild(el("span", "badge badge-participant", artifact.participant));
    badges.appendChild(el("span", "badge badge-" + artifact.type, artifact.type));
    if (artifact.category) badges.appendChild(el("span", "badge badge-category", artifact.category));
    info.appendChild(badges);
    info.appendChild(el("div", "preview-desc", artifact.description || ""));
    info.appendChild(el("div", "preview-time", formatTime(artifact.start) + " \u2013 " + formatTime(artifact.end)));
    card.appendChild(info);

    // Drag for reorder
    card.addEventListener("dragstart", function (e) {
      e.stopPropagation();
      card.classList.add("dragging");
      e.dataTransfer.setData("text/plain", artifact.id);
      e.dataTransfer.setData("application/x-bucket", bucketName);
      e.dataTransfer.setData("application/x-insight", insightId);
      e.dataTransfer.effectAllowed = "move";
    });
    card.addEventListener("dragend", function () {
      card.classList.remove("dragging");
    });

    return card;
  }

  // ---- Sidebar resize ----

  function initSidebarResize() {
    var handle = qs("#resizeHandle");
    var sidebar = qs("#mediaLibrary");
    var isResizing = false;

    // Restore saved width
    try {
      var saved = window.localStorage.getItem(SIDEBAR_WIDTH_KEY);
      if (saved) {
        var w = parseInt(saved, 10);
        if (w >= 280 && w <= sidebar.parentElement.offsetWidth * 0.6) {
          sidebar.style.width = w + "px";
        }
      }
    } catch (_) {}

    handle.addEventListener("mousedown", function (e) {
      e.preventDefault();
      isResizing = true;
      handle.classList.add("active");
      sidebar.classList.add("resizing");
      document.body.style.cursor = "col-resize";
      document.body.style.userSelect = "none";
    });

    var rafPending = false;
    document.addEventListener("mousemove", function (e) {
      if (!isResizing) return;
      if (rafPending) return;
      rafPending = true;
      var clientX = e.clientX;
      requestAnimationFrame(function () {
        var panelLeft = sidebar.parentElement.getBoundingClientRect().left;
        var maxW = sidebar.parentElement.offsetWidth * 0.6;
        var w = Math.max(280, Math.min(maxW, clientX - panelLeft));
        sidebar.style.width = w + "px";
        rafPending = false;
      });
    });

    document.addEventListener("mouseup", function () {
      if (!isResizing) return;
      isResizing = false;
      handle.classList.remove("active");
      sidebar.classList.remove("resizing");
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      try {
        window.localStorage.setItem(
          SIDEBAR_WIDTH_KEY,
          parseInt(sidebar.style.width, 10) || 420
        );
      } catch (_) {}
    });

    // Collapse toggle
    var _savedWidth = "";
    var collapsibles = sidebar.querySelectorAll(
      "#filterPanel, #artifactListHeader, #artifactGrid, #sidebarHeader h2"
    );
    collapsibles.forEach(function (el) {
      el.classList.add("sidebar-collapsible");
    });

    qs("#collapseBtn").addEventListener("click", function () {
      state.sidebarCollapsed = !state.sidebarCollapsed;
      if (state.sidebarCollapsed) {
        // Fade out content, then collapse
        _savedWidth = sidebar.style.width;
        collapsibles.forEach(function (el) {
          el.style.opacity = "0";
        });
        setTimeout(function () {
          sidebar.style.width = "";
          sidebar.classList.add("collapsed");
        }, 100);
      } else {
        // Expand width first, keep content hidden, then fade in
        // Use a wrapper class that keeps content hidden during width transition
        sidebar.classList.add("expanding");
        sidebar.classList.remove("collapsed");
        if (_savedWidth) sidebar.style.width = _savedWidth;
        // Wait for width transition to finish before showing content
        function onExpanded(e) {
          if (e.propertyName !== "width") return;
          sidebar.removeEventListener("transitionend", onExpanded);
          sidebar.classList.remove("expanding");
          collapsibles.forEach(function (el) {
            el.style.opacity = "0";
          });
          requestAnimationFrame(function () {
            requestAnimationFrame(function () {
              collapsibles.forEach(function (el) {
                el.style.opacity = "";
              });
            });
          });
        }
        sidebar.addEventListener("transitionend", onExpanded);
      }
    });
  }

  // ---- Global events ----

  function bindGlobalEvents() {
    // Save shortcut
    document.addEventListener("keydown", function (e) {
      if ((e.ctrlKey || e.metaKey) && e.key === "s") {
        e.preventDefault();
        saveAll();
      }
    });

    // Navigation guard
    window.addEventListener("beforeunload", function (e) {
      if (Object.keys(state.dirtyIds).length > 0) {
        e.preventDefault();
        e.returnValue = "";
      }
    });

    // Close popover on outside click
    document.addEventListener("click", function (e) {
      if (!e.target.closest("#addPopover") && !e.target.closest(".artifact-card")) {
        hidePopover();
      }
    });

    // Header buttons
    qs("#saveAllBtn").addEventListener("click", saveAll);
    qs("#discardBtn").addEventListener("click", discardAll);
    qs("#newInsightBtn").addEventListener("click", createNewInsight);

    // Popover bucket buttons
    var bucketBtns = qsa("#addPopover .popover-btn");
    for (var i = 0; i < bucketBtns.length; i++) {
      bucketBtns[i].addEventListener("click", onPopoverBucketClick);
    }

    // Generate viewer
    qs("#generateViewerBtn").addEventListener("click", function () {
      this.disabled = true;
      var btn = this;
      fetch("api/generate-viewer", { method: "POST" })
        .then(function (r) {
          return r.json();
        })
        .then(function (data) {
          btn.disabled = false;
          if (data.ok) {
            showToast("Viewer generated: " + data.file);
          } else {
            showToast("Error: " + (data.error || "Unknown error"));
          }
        })
        .catch(function () {
          btn.disabled = false;
          showToast("Failed to generate viewer.");
        });
    });
  }

  function checkNavLinks() {
    fetch("/api/status")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data.studio) {
          var link = qs("#studioLink");
          if (link) link.classList.remove("hidden");
        }
      })
      .catch(function () {});
  }

  // ---- Init ----

  function init() {
    initThemeToggle();
    initSidebarResize();
    bindFilterEvents();
    initSortToolbar();
    bindGlobalEvents();
    loadData();
    checkNavLinks();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
