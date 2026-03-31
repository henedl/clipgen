/* clipgen Screenspace */

(function () {
  "use strict";

  var THEME_STORAGE_KEY = "clipgen-screenspace-theme";
  var POLL_INTERVAL = 3000;
  var FRAME_STEP = 1.0;

  var TASK_COLORS = {
    color: "#8b5cf6",
    change: "#f97316",
    similarity: "#0ea5e9",
    text: "#10b981",
    numbers: "#eab308",
    timelapse: "#ec4899",
    template: "#f43f5e",
    flow: "#6366f1",
    scene: "#14b8a6",
  };

  var SS_TYPE_ICON_PATHS = {
    color: { viewBox: "0 0 16 16", paths: [
      { d: "M15 4C15 5.39788 14.0439 6.57245 12.75 6.90549V8.5C12.75 8.69891 12.671 8.88968 12.5303 9.03033L12.0303 9.53033C11.7374 9.82322 11.2626 9.82322 10.9697 9.53033L10.25 8.81069L5.57322 13.4875C5.24503 13.8157 4.79992 14.0001 4.33579 14.0001H3.66421C3.59791 14.0001 3.53432 14.0264 3.48744 14.0733L2.78033 14.7804C2.63968 14.921 2.44891 15.0001 2.25 15.0001C2.05109 15.0001 1.86032 14.921 1.71967 14.7804L1.21967 14.2804C0.926777 13.9875 0.926777 13.5126 1.21967 13.2197L1.92678 12.5126C1.97366 12.4657 2 12.4021 2 12.3358V11.6643C2 11.2001 2.18437 10.755 2.51256 10.4268L7.18937 5.75003L6.46967 5.03033C6.17678 4.73744 6.17678 4.26256 6.46967 3.96967L6.96967 3.46967C7.11032 3.32902 7.30109 3.25 7.5 3.25H9.09451C9.42755 1.95608 10.6021 1 12 1C13.6569 1 15 2.34315 15 4ZM9.18937 7.75003L8.25003 6.81069L3.57322 11.4875C3.52634 11.5344 3.5 11.598 3.5 11.6643V12.3358C3.5 12.3938 3.49713 12.4514 3.49146 12.5086C3.54862 12.5029 3.60627 12.5001 3.66421 12.5001H4.33579C4.40209 12.5001 4.46568 12.4737 4.51256 12.4268L9.18937 7.75003Z", fillRule: "evenodd" }
    ]},
    change: { viewBox: "0 0 16 16", paths: [
      { d: "M9.58011 1.07655C9.88578 1.22638 10.0522 1.56328 9.98545 1.89709L9.16486 6H13.25C13.5437 6 13.8103 6.17136 13.9323 6.43847C14.0542 6.70558 14.0091 7.0193 13.8168 7.2412L7.31678 14.7412C7.09383 14.9984 6.72559 15.0733 6.41991 14.9234C6.11424 14.7736 5.94781 14.4367 6.01458 14.1029L6.83516 10H2.75001C2.45637 10 2.18974 9.82864 2.06777 9.56153C1.9458 9.29442 1.99093 8.9807 2.18324 8.7588L8.68324 1.2588C8.90619 1.00155 9.27444 0.92672 9.58011 1.07655Z", fillRule: "evenodd" }
    ]},
    similarity: { viewBox: "0 0 16 16", paths: [
      { d: "M2 4C2 2.89543 2.89543 2 4 2H12C13.1046 2 14 2.89543 14 4V12C14 13.1046 13.1046 14 12 14H4C2.89543 14 2 13.1046 2 12V4ZM12.5 9.70711C12.5 9.5745 12.4473 9.44732 12.3536 9.35355L11.3536 8.35355C11.1583 8.15829 10.8417 8.15829 10.6464 8.35355L9.35355 9.64645C9.15829 9.84171 8.84171 9.84171 8.64645 9.64645L6.35355 7.35355C6.15829 7.15829 5.84171 7.15829 5.64645 7.35355L3.64645 9.35355C3.55268 9.44732 3.5 9.5745 3.5 9.70711V12C3.5 12.2761 3.72386 12.5 4 12.5H12C12.2761 12.5 12.5 12.2761 12.5 12V9.70711ZM12 5C12 5.55228 11.5523 6 11 6C10.4477 6 10 5.55228 10 5C10 4.44772 10.4477 4 11 4C11.5523 4 12 4.44772 12 5Z", fillRule: "evenodd" }
    ]},
    text: { viewBox: "0 0 16 16", paths: [
      { d: "M11 5C11.299 5 11.5693 5.17751 11.6882 5.45179L14.9382 12.9518C15.1029 13.3319 14.9283 13.7735 14.5482 13.9382C14.1682 14.1029 13.7266 13.9283 13.5619 13.5482L12.8908 11.9997H9.10923L8.4382 13.5482C8.2735 13.9283 7.83189 14.1029 7.45182 13.9382C7.07176 13.7735 6.89717 13.3319 7.06186 12.9518L10.3119 5.45179C10.4307 5.17751 10.7011 5 11 5ZM9.75923 10.4997H12.2408L11 7.63628L9.75923 10.4997Z", fillRule: "evenodd" },
      { d: "M5.00003 1C5.41424 1 5.75003 1.33579 5.75003 1.75V3.01104C6.16299 3.02322 6.5735 3.04541 6.98131 3.0774C7.44038 3.11341 7.89601 3.16182 8.34786 3.22231C8.75842 3.27727 9.04668 3.65464 8.99172 4.06519C8.93676 4.47574 8.55938 4.76401 8.14883 4.70905C7.92894 4.67961 7.70808 4.65321 7.48628 4.6299C7.1301 5.85717 6.59808 7.00928 5.91941 8.05729C6.15555 8.36066 6.40658 8.65193 6.67142 8.92999C6.95709 9.22993 6.94553 9.70466 6.64559 9.99034C6.34565 10.276 5.87092 10.2644 5.58525 9.96451C5.38294 9.7521 5.18774 9.53284 5.00002 9.30711C4.18402 10.2884 3.22645 11.1474 2.15883 11.853C1.81326 12.0813 1.34799 11.9863 1.11962 11.6408C0.891239 11.2952 0.986242 10.8299 1.33181 10.6015C2.3813 9.90797 3.31021 9.04714 4.08066 8.05729C3.88359 7.75296 3.69887 7.43984 3.52724 7.11865C3.33202 6.75332 3.46992 6.29891 3.83524 6.10369C4.20057 5.90847 4.65498 6.04637 4.8502 6.4117C4.89895 6.50293 4.9489 6.59343 5.00002 6.68318C5.38798 6.00207 5.7083 5.27759 5.95187 4.51891C5.63619 4.50635 5.31887 4.5 5.00003 4.5C3.93193 4.5 2.88086 4.57121 1.85122 4.70905C1.44067 4.76401 1.0633 4.47574 1.00834 4.06519C0.95338 3.65464 1.24164 3.27727 1.65219 3.22231C2.50548 3.10808 3.37219 3.03692 4.25003 3.01104V1.75C4.25003 1.33579 4.58582 1 5.00003 1Z", fillRule: "evenodd" }
    ]},
    numbers: { viewBox: "0 0 16 16", paths: [
      { d: "M7.48677 2.89033C7.56427 2.48344 7.29725 2.09075 6.89035 2.01325C6.48345 1.93574 6.09077 2.20277 6.01326 2.60967L5.55827 4.99835H3.60963C3.19542 4.99835 2.85963 5.33414 2.85963 5.74835C2.85963 6.16257 3.19542 6.49835 3.60963 6.49835H5.27256L4.7016 9.49589H2.74963C2.33542 9.49589 1.99963 9.83168 1.99963 10.2459C1.99963 10.6601 2.33542 10.9959 2.74963 10.9959H4.41588L4.01326 13.1097C3.93576 13.5166 4.20278 13.9092 4.60968 13.9868C5.01658 14.0643 5.40926 13.7972 5.48677 13.3903L5.94285 10.9959H8.91589L8.51326 13.1097C8.43576 13.5166 8.70278 13.9092 9.10968 13.9868C9.51658 14.0643 9.90927 13.7972 9.98677 13.3903L10.4429 10.9959H12.3896C12.8038 10.9959 13.1396 10.6601 13.1396 10.2459C13.1396 9.83168 12.8038 9.49589 12.3896 9.49589H10.7286L11.2995 6.49835H13.2496C13.6638 6.49835 13.9996 6.16257 13.9996 5.74835C13.9996 5.33414 13.6638 4.99835 13.2496 4.99835H11.5852L11.9868 2.89033C12.0643 2.48344 11.7972 2.09075 11.3903 2.01325C10.9835 1.93574 10.5908 2.20277 10.5133 2.60967L10.0583 4.99835H7.08524L7.48677 2.89033ZM6.79953 6.49835L6.22857 9.49589H9.2016L9.77256 6.49835H6.79953Z", fillRule: "evenodd" }
    ]},
    template: { viewBox: "0 0 16 16", paths: [
      { d: "M2 3.5A1.5 1.5 0 0 1 3.5 2H5a.75.75 0 0 1 0 1.5H3.5v1.75a.75.75 0 0 1-1.5 0V3.5ZM11 2a.75.75 0 0 0 0 1.5h1.5v1.75a.75.75 0 0 0 1.5 0V3.5A1.5 1.5 0 0 0 12.5 2H11ZM2.75 10.75a.75.75 0 0 1 .75.75v1.5H5a.75.75 0 0 1 0 1.5H3.5A1.5 1.5 0 0 1 2 13v-1.5a.75.75 0 0 1 .75-.75ZM13.25 10.75a.75.75 0 0 1 .75.75V13a1.5 1.5 0 0 1-1.5 1.5H11a.75.75 0 0 1 0-1.5h1.5v-1.5a.75.75 0 0 1 .75-.75ZM10 8a2 2 0 1 1-4 0 2 2 0 0 1 4 0Z", fillRule: "evenodd" }
    ]},
    flow: { viewBox: "0 0 16 16", paths: [
      { d: "M5.28 10.22a.75.75 0 0 1 0 1.06l-1.47 1.47h8.44a.75.75 0 0 1 0 1.5H3.81l1.47 1.47a.75.75 0 0 1-1.06 1.06l-2.75-2.75a.75.75 0 0 1 0-1.06l2.75-2.75a.75.75 0 0 1 1.06 0ZM10.72.22a.75.75 0 0 1 1.06 0l2.75 2.75a.75.75 0 0 1 0 1.06l-2.75 2.75a.75.75 0 1 1-1.06-1.06l1.47-1.47H3.75a.75.75 0 0 1 0-1.5h8.44L10.72 1.28a.75.75 0 0 1 0-1.06Z", fillRule: "evenodd" }
    ]},
    scene: { viewBox: "0 0 16 16", paths: [
      { d: "M2 3.5A1.5 1.5 0 0 1 3.5 2h2A1.5 1.5 0 0 1 7 3.5v2A1.5 1.5 0 0 1 5.5 7h-2A1.5 1.5 0 0 1 2 5.5v-2ZM9 3.5A1.5 1.5 0 0 1 10.5 2h2A1.5 1.5 0 0 1 14 3.5v2A1.5 1.5 0 0 1 12.5 7h-2A1.5 1.5 0 0 1 9 5.5v-2ZM2 10.5A1.5 1.5 0 0 1 3.5 9h2A1.5 1.5 0 0 1 7 10.5v2A1.5 1.5 0 0 1 5.5 14h-2A1.5 1.5 0 0 1 2 12.5v-2ZM9 10.5A1.5 1.5 0 0 1 10.5 9h2A1.5 1.5 0 0 1 14 10.5v2A1.5 1.5 0 0 1 12.5 14h-2A1.5 1.5 0 0 1 9 12.5v-2Z" }
    ]}
  };

  function buildTypeIconSvg(type) {
    var info = SS_TYPE_ICON_PATHS[type];
    if (!info) return null;
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "14");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", info.viewBox);
    svg.setAttribute("fill", "currentColor");
    info.paths.forEach(function (p) {
      var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
      path.setAttribute("d", p.d);
      if (p.fillRule) {
        path.setAttribute("fill-rule", p.fillRule);
        path.setAttribute("clip-rule", p.fillRule);
      }
      svg.appendChild(path);
    });
    return svg;
  }

  var TIMELINE_CANVAS_HEIGHT = 64;

  var _paletteDocListeners = null;

  var REGION_COLORS = [
    "#3b82f6", "#ef4444", "#22c55e", "#f59e0b",
    "#8b5cf6", "#06b6d4", "#ec4899", "#14b8a6",
  ];

  var state = {
    participants: [],
    selectedParticipant: null,
    videoInfo: null,
    currentTimestamp: 0,
    frameImage: null,
    frameLoading: false,
    regions: {},
    activeRegion: null,
    drawingRegion: null,
    pendingRegion: null,
    draggingRegion: null,
    resizingRegion: null,
    hoveredRegion: null,
    timelineZoom: 1,
    timelineOffset: 0,
    inMarker: null,
    outMarker: null,
    activeWorkflow: "color",
    referenceTimestamp: null,
    sceneReferences: [],
    tasks: [],
    selectedTaskId: null,
    hoveredTaskId: null,
    selectedTaskResults: null,
    pollTimer: null,
    queuePaused: false,
    timelineDragging: false,
    panelHeight: 340,
    previewMaxWidth: 100,
    taskFilter: null,
    pipetteActive: false,
    runParticipants: [],
    runRegions: [],
    taskEvents: {},
    showExcluded: true,
    showRegionLabels: true,
    showRegionOverlays: true,
    stashes: [],
    previewRegions: null,
    resultOverlay: null,
    heatmapOverlay: null,
    uploadedTemplate: null,
  };

  var _timelineHitRects = [];
  var _overlayRaf = 0;
  var _cachedOverlayRect = null;
  var _lastPollFingerprint = "";

  var _cachedThemeColors = null;

  function refreshThemeColors() {
    var cs = getComputedStyle(document.documentElement);
    _cachedThemeColors = {
      surfaceAlt: cs.getPropertyValue("--color-surface-alt").trim() || "#f1ece4",
      border: cs.getPropertyValue("--color-border").trim() || "#e0ddd7",
      textDim: cs.getPropertyValue("--color-text-dim").trim() || "#6b7280",
      accent: cs.getPropertyValue("--color-accent").trim() || "#1d4f72",
      fontMono: cs.getPropertyValue("--font-mono").trim() || "monospace",
    };
  }

  function getThemeColors() {
    if (!_cachedThemeColors) refreshThemeColors();
    return _cachedThemeColors;
  }

  // ---- Helpers ----

  function qs(sel) { return document.querySelector(sel); }
  function qsa(sel) { return document.querySelectorAll(sel); }
  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
  }

  function formatDuration(secs) {
    secs = Math.round(secs);
    if (secs >= 3600) {
      var h = Math.floor(secs / 3600);
      var m = Math.floor((secs % 3600) / 60);
      var s = secs % 60;
      return h + ":" + (m < 10 ? "0" : "") + m + ":" + (s < 10 ? "0" : "") + s;
    }
    var m2 = Math.floor(secs / 60);
    var s2 = secs % 60;
    return m2 + ":" + (s2 < 10 ? "0" : "") + s2;
  }

  function formatTimestamp(secs) {
    var m = Math.floor(secs / 60);
    var s = secs % 60;
    return m + ":" + (s < 10 ? "0" : "") + s.toFixed(1);
  }

  function parseTimestampInput(str) {
    str = (str || "").trim();
    var parts = str.split(":");
    if (parts.length === 3) return (+parts[0]) * 3600 + (+parts[1]) * 60 + (+parts[2]);
    if (parts.length === 2) return (+parts[0]) * 60 + (+parts[1]);
    var n = parseFloat(str);
    return isNaN(n) ? null : n;
  }

  function clamp(val, min, max) {
    return Math.max(min, Math.min(max, val));
  }

  function showToast(msg) {
    var t = qs("#toast");
    t.textContent = msg;
    t.classList.remove("hidden");
    clearTimeout(t._timer);
    t._timer = setTimeout(function () { t.classList.add("hidden"); }, 3000);
  }

  function regionColorForIndex(i) {
    return REGION_COLORS[i % REGION_COLORS.length];
  }

  function regionColorByName(name) {
    var names = Object.keys(state.regions);
    var idx = names.indexOf(name);
    return idx >= 0 ? regionColorForIndex(idx) : REGION_COLORS[0];
  }

  function regionToPixels(r) {
    if (!r.source_width) return r;
    var canvas = qs("#overlayCanvas");
    return {
      x: Math.round(r.x * canvas.width),
      y: Math.round(r.y * canvas.height),
      w: Math.round(r.w * canvas.width),
      h: Math.round(r.h * canvas.height),
    };
  }

  function taskTypeColor(type) {
    return TASK_COLORS[type] || "#888";
  }

  // ---- API helpers ----

  function apiGet(path) {
    return fetch(path).then(function (r) {
      if (!r.ok) throw new Error("Server error " + r.status);
      return r.json();
    });
  }

  function apiPost(path, body) {
    return fetch(path, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    }).then(function (r) {
      if (!r.ok) throw new Error("Server error " + r.status);
      return r.json();
    });
  }

  function apiDelete(path) {
    return fetch(path, { method: "DELETE" }).then(function (r) {
      if (!r.ok) throw new Error("Server error " + r.status);
      return r.json();
    });
  }

  function apiPut(path, body) {
    return fetch(path, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    }).then(function (r) {
      if (!r.ok) throw new Error("Server error " + r.status);
      return r.json();
    });
  }

  function frameUrl(pid, ts) {
    return "api/video/frame/" + encodeURIComponent(pid) + "/" + Number(ts).toFixed(6);
  }

  // ---- Theme toggle (matches Studio) ----

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
    try { stored = window.localStorage.getItem(THEME_STORAGE_KEY); } catch (_) {}
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
        prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
      } catch (_) {}
      next = prefersDark ? "light" : "dark";
    }
    root.setAttribute("data-theme", next);
    try { window.localStorage.setItem(THEME_STORAGE_KEY, next); } catch (_) {}
    updateThemeToggleButton(next);
    refreshThemeColors();
    renderTimeline();
  }

  function updateThemeToggleButton(theme) {
    var btn = qs("#themeToggle");
    if (!btn) return;
    btn.setAttribute("data-theme", theme || "");
    btn.setAttribute("aria-pressed", theme === "dark" ? "true" : "false");
  }

  // ---- Nav links ----

  function checkNavLinks() {
    fetch("../api/status").then(function (r) { return r.json(); }).then(function (data) {
      if (data.studio) qs("#studioLink").classList.remove("hidden");
      if (data.insights) qs("#insightsLink").classList.remove("hidden");
    }).catch(function () {});
  }

  // ---- Participants ----

  function renderParticipantSelect() {
    var sel = qs("#participantSelect");
    sel.innerHTML = "";
    if (state.participants.length === 0) {
      var opt = el("option", null, "No participants");
      opt.value = "";
      sel.appendChild(opt);
      return;
    }
    state.participants.forEach(function (p) {
      var opt = el("option", null, p.id);
      opt.value = p.id;
      sel.appendChild(opt);
    });
  }

  function renderRunParticipantPicker() {
    var wrap = qs("#runParticipantPicker");
    if (!wrap) return;
    wrap.innerHTML = "";
    if (state.participants.length <= 1) return;

    var btn = el("button", "run-picker-btn");
    btn.type = "button";
    updatePickerBtnText(btn);

    var panel = el("div", "run-picker-panel hidden");

    var toggleAll = el("span", "run-picker-toggle-all");
    toggleAll.textContent = state.runParticipants.length === state.participants.length ? "Deselect all" : "Select all";
    toggleAll.addEventListener("click", function () {
      var allSelected = state.runParticipants.length === state.participants.length;
      state.runParticipants = allSelected ? [] : state.participants.map(function (p) { return p.id; });
      var cbs = panel.querySelectorAll("input[type=checkbox]");
      for (var i = 0; i < cbs.length; i++) cbs[i].checked = !allSelected;
      toggleAll.textContent = allSelected ? "Select all" : "Deselect all";
      updatePickerBtnText(btn);
      updateRunButton();
    });
    panel.appendChild(toggleAll);

    state.participants.forEach(function (p) {
      var lbl = document.createElement("label");
      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.value = p.id;
      cb.checked = state.runParticipants.indexOf(p.id) >= 0;
      cb.addEventListener("change", function () {
        if (cb.checked) {
          if (state.runParticipants.indexOf(p.id) < 0) state.runParticipants.push(p.id);
        } else {
          state.runParticipants = state.runParticipants.filter(function (id) { return id !== p.id; });
        }
        toggleAll.textContent = state.runParticipants.length === state.participants.length ? "Deselect all" : "Select all";
        updatePickerBtnText(btn);
        updateRunButton();
      });
      lbl.appendChild(cb);
      lbl.appendChild(document.createTextNode(p.id));
      panel.appendChild(lbl);
    });

    btn.addEventListener("click", function (e) {
      e.stopPropagation();
      var open = !panel.classList.contains("hidden");
      panel.classList.toggle("hidden", open);
      btn.classList.toggle("open", !open);
    });

    wrap.appendChild(btn);
    wrap.appendChild(panel);
  }

  function updatePickerBtnText(btn) {
    var n = state.runParticipants.length;
    var text = n === 0 ? "No participants"
      : n === 1 ? state.runParticipants[0]
      : n + " participants";
    btn.innerHTML = "";
    btn.appendChild(el("span", "run-picker-btn-text", text));
    var chevron = el("span", "chevron");
    chevron.appendChild(svgChevronDownIcon());
    btn.appendChild(chevron);
  }

  function closeRunPicker() {
    var panels = qsa(".run-picker-panel");
    var btns = qsa(".run-picker-btn");
    for (var i = 0; i < panels.length; i++) panels[i].classList.add("hidden");
    for (var i = 0; i < btns.length; i++) btns[i].classList.remove("open");
  }

  function allAvailableRegionNames() {
    var names = Object.keys(state.regions);
    state.stashes.forEach(function (stash) {
      Object.keys(stash.regions).forEach(function (n) {
        if (names.indexOf(n) < 0) names.push(n);
      });
    });
    return names;
  }

  function renderRunRegionPicker() {
    var wrap = qs("#runRegionPicker");
    if (!wrap) return;
    wrap.innerHTML = "";
    var names = Object.keys(state.regions);
    var allNames = allAvailableRegionNames();
    // Remove any runRegions that no longer exist in active or stashes
    state.runRegions = state.runRegions.filter(function (r) { return allNames.indexOf(r) >= 0; });
    // Auto-select the active region when no explicit selection has been made
    if (state.runRegions.length === 0 && state.activeRegion && names.indexOf(state.activeRegion) >= 0) {
      state.runRegions = [state.activeRegion];
    }
    if (allNames.length === 0) return;

    var btn = el("button", "run-picker-btn");
    btn.type = "button";
    updateRegionPickerBtnText(btn);

    var panel = el("div", "run-picker-panel hidden");

    if (names.length > 0) {
      var toggleAll = el("span", "run-picker-toggle-all");
      toggleAll.textContent = state.runRegions.length === names.length ? "Deselect all" : "Select all";
      toggleAll.addEventListener("click", function () {
        var allSelected = state.runRegions.length === names.length;
        state.runRegions = allSelected ? [] : names.slice();
        var cbs = panel.querySelectorAll(".run-picker-active-region input[type=checkbox]");
        for (var i = 0; i < cbs.length; i++) cbs[i].checked = !allSelected;
        toggleAll.textContent = allSelected ? "Select all" : "Deselect all";
        updateRegionPickerBtnText(btn);
        updateRunButton();
      });
      panel.appendChild(toggleAll);

      names.forEach(function (name, idx) {
        var color = regionColorForIndex(idx);
        var lbl = document.createElement("label");
        lbl.className = "run-picker-active-region";
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.value = name;
        cb.checked = state.runRegions.indexOf(name) >= 0;
        cb.addEventListener("change", function () {
          if (cb.checked) {
            if (state.runRegions.indexOf(name) < 0) state.runRegions.push(name);
          } else {
            state.runRegions = state.runRegions.filter(function (r) { return r !== name; });
          }
          toggleAll.textContent = state.runRegions.length === names.length ? "Deselect all" : "Select all";
          updateRegionPickerBtnText(btn);
          updateRunButton();
        });
        lbl.appendChild(cb);
        var dot = el("span", "region-chip-dot");
        dot.style.background = color;
        lbl.appendChild(dot);
        var nameSpan = el("span", "run-picker-label-text", name);
        lbl.appendChild(nameSpan);
        panel.appendChild(lbl);
      });
    }

    // Stash folders
    state.stashes.forEach(function (stash) {
      var stashNames = Object.keys(stash.regions);
      if (stashNames.length === 0) return;

      var header = el("div", "stash-folder-header");
      var chevron = el("span", "chevron", "\u25B8");
      header.appendChild(chevron);
      header.appendChild(document.createTextNode(stash.name + " (" + stashNames.length + ")"));

      var content = el("div", "stash-folder-content");

      header.addEventListener("click", function () {
        var expanded = header.classList.toggle("expanded");
        content.classList.toggle("expanded", expanded);
      });

      panel.appendChild(header);

      stashNames.forEach(function (name, idx) {
        var color = regionColorForIndex(idx);
        var lbl = document.createElement("label");
        lbl.className = "stash-folder-item";
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.value = name;
        cb.checked = state.runRegions.indexOf(name) >= 0;
        cb.addEventListener("change", function () {
          if (cb.checked) {
            if (state.runRegions.indexOf(name) < 0) state.runRegions.push(name);
          } else {
            state.runRegions = state.runRegions.filter(function (r) { return r !== name; });
          }
          updateRegionPickerBtnText(btn);
          updateRunButton();
        });
        lbl.appendChild(cb);
        var dot = el("span", "region-chip-dot");
        dot.style.background = color;
        lbl.appendChild(dot);
        var nameSpan = el("span", "run-picker-label-text", name);
        lbl.appendChild(nameSpan);
        content.appendChild(lbl);
      });

      panel.appendChild(content);
    });

    btn.addEventListener("click", function (e) {
      e.stopPropagation();
      var open = !panel.classList.contains("hidden");
      panel.classList.toggle("hidden", open);
      btn.classList.toggle("open", !open);
    });

    wrap.appendChild(btn);
    wrap.appendChild(panel);
  }

  function updateRegionPickerBtnText(btn) {
    var n = state.runRegions.length;
    var text = n === 0 ? "No region"
      : n === 1 ? state.runRegions[0]
      : n + " regions";
    btn.innerHTML = "";
    btn.appendChild(el("span", "run-picker-btn-text", text));
    var chevron = el("span", "chevron");
    chevron.appendChild(svgChevronDownIcon());
    btn.appendChild(chevron);
  }

  function selectParticipant(pid, initialTimestamp) {
    state.selectedParticipant = pid;
    state.currentTimestamp = 0;
    state.videoInfo = null;
    state.referenceTimestamp = null;
    state.sceneReferences = [];
    state.resultOverlay = null;
    state.heatmapOverlay = null;
    qs("#participantSelect").value = pid;
    qs("#videoInfo").textContent = "";
    qs("#frameEmpty").classList.remove("hidden");

    apiGet("api/video/info/" + encodeURIComponent(pid))
      .then(function (data) {
        if (!data.ok) return;
        state.videoInfo = data.info;
        var parts = [];
        if (data.info.duration) parts.push(formatDuration(data.info.duration));
        if (data.info.width && data.info.height) parts.push(data.info.width + "x" + data.info.height);
        if (data.info.fps) parts.push(Math.round(data.info.fps) + "fps");
        qs("#videoInfo").textContent = parts.join(" \u00b7 ");
        renderTimeline();
        loadFrame(initialTimestamp !== undefined ? initialTimestamp : 0);
      })
      .catch(function () { showToast("Failed to load video info"); });
  }

  // ---- Frame viewer ----

  function seekPlayhead(timestamp) {
    state.currentTimestamp = timestamp;
    qs("#timestampInput").value = formatTimestamp(timestamp);
    renderPlayhead();
  }

  var _pendingFrameTs = null;
  var _loadedFrameTs = null;

  function loadFrame(timestamp) {
    if (!state.selectedParticipant) return;
    seekPlayhead(timestamp);
    if (state.frameLoading) {
      _pendingFrameTs = timestamp;
      return;
    }
    _fetchFrame(timestamp);
  }

  function _fetchFrame(timestamp) {
    state.frameLoading = true;
    _pendingFrameTs = null;
    _loadedFrameTs = timestamp;

    var img = new Image();
    img.onload = function () {
      state.frameImage = img;
      state.frameLoading = false;
      qs("#frameEmpty").classList.add("hidden");
      var canvas = qs("#frameCanvas");
      var overlay = qs("#overlayCanvas");
      if (canvas.width !== img.naturalWidth || canvas.height !== img.naturalHeight) {
        canvas.width = img.naturalWidth;
        canvas.height = img.naturalHeight;
        overlay.width = img.naturalWidth;
        overlay.height = img.naturalHeight;
      }
      var ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0);
      renderOverlay();
      renderTimeline();
      if (_pendingFrameTs !== null && _pendingFrameTs !== _loadedFrameTs) {
        _fetchFrame(_pendingFrameTs);
      }
    };
    img.onerror = function () {
      state.frameLoading = false;
      if (_pendingFrameTs !== null && _pendingFrameTs !== _loadedFrameTs) {
        _fetchFrame(_pendingFrameTs);
      }
    };
    img.src = frameUrl(state.selectedParticipant, timestamp);
  }

  function initFrameControls() {
    qs("#framePrev").appendChild(svgChevronLeftIcon());
    qs("#frameNext").appendChild(svgChevronRightIcon());
    var input = qs("#timestampInput");

    input.addEventListener("change", function () {
      var ts = parseTimestampInput(input.value);
      if (ts !== null && state.videoInfo) {
        ts = clamp(ts, 0, state.videoInfo.duration || 0);
        loadFrame(ts);
      }
    });

    input.addEventListener("keydown", function (e) {
      if (e.key === "Enter") input.blur();
    });

    qs("#framePrev").addEventListener("click", function () {
      if (!state.videoInfo) return;
      var ts = clamp(state.currentTimestamp - FRAME_STEP, 0, state.videoInfo.duration);
      loadFrame(ts);
    });

    qs("#frameNext").addEventListener("click", function () {
      if (!state.videoInfo) return;
      var ts = clamp(state.currentTimestamp + FRAME_STEP, 0, state.videoInfo.duration);
      loadFrame(ts);
    });
  }

  // ---- Region drawing ----

  function canvasCoords(canvas, event, cachedRect) {
    var rect = cachedRect || canvas.getBoundingClientRect();
    var scaleX = canvas.width / rect.width;
    var scaleY = canvas.height / rect.height;
    return {
      x: Math.round((event.clientX - rect.left) * scaleX),
      y: Math.round((event.clientY - rect.top) * scaleY),
    };
  }

  function scheduleOverlayRender() {
    if (_overlayRaf) return;
    _overlayRaf = requestAnimationFrame(function () {
      _overlayRaf = 0;
      renderOverlay();
    });
  }

  function flushOverlayRender() {
    if (_overlayRaf) {
      cancelAnimationFrame(_overlayRaf);
      _overlayRaf = 0;
    }
    renderOverlay();
  }

  function normalizeRect(x1, y1, x2, y2) {
    return {
      x: Math.min(x1, x2),
      y: Math.min(y1, y2),
      w: Math.abs(x2 - x1),
      h: Math.abs(y2 - y1),
    };
  }

  function computeLabelRect(r, name, ctx, s) {
    var fontSize = Math.round(12 * s);
    var labelH = Math.round(16 * s);
    var pad = Math.round(4 * s);
    var gripW = Math.round(8 * s);
    var gripPadL = Math.round(3 * s);
    ctx.font = "bold " + fontSize + "px -apple-system, BlinkMacSystemFont, sans-serif";
    var textW = ctx.measureText(name).width;
    var labelW = gripPadL + gripW + pad + textW + pad;
    return { x: r.x, y: r.y - labelH, w: labelW, h: labelH, gripPadL: gripPadL, gripW: gripW, pad: pad, fontSize: fontSize };
  }

  function findHitRegion(px, py, s, ctx) {
    if (!state.showRegionOverlays) return null;
    var names = Object.keys(state.regions);
    var handleSize = Math.round(14 * s);
    for (var i = names.length - 1; i >= 0; i--) {
      var name = names[i];
      var r = regionToPixels(state.regions[name]);
      if (px >= r.x + r.w - handleSize && px <= r.x + r.w && py >= r.y + r.h - handleSize && py <= r.y + r.h) {
        return { name: name, handle: "resize" };
      }
      if (state.showRegionLabels) {
        var lr = computeLabelRect(r, name, ctx, s);
        if (px >= lr.x && px <= lr.x + lr.w && py >= lr.y && py <= lr.y + lr.h) {
          return { name: name, handle: "move" };
        }
      }
    }
    return null;
  }

  function saveRegionUpdate(name) {
    var region = state.regions[name];
    if (!region) return;
    var canvas = qs("#overlayCanvas");
    var px = regionToPixels(region);
    var body = { name: name, x: px.x, y: px.y, w: px.w, h: px.h, canvas_width: canvas.width, canvas_height: canvas.height };
    if (region.description) body.description = region.description;
    apiPost("api/regions", body)
      .then(function (data) {
        if (data.ok) {
          var saved = data.region;
          if (region.description) saved.description = region.description;
          state.regions[name] = saved;
          renderOverlay();
        }
      })
      .catch(function () { showToast("Failed to update region"); });
  }

  function initRegionDrawing() {
    var overlay = qs("#overlayCanvas");

    overlay.addEventListener("mousedown", function (e) {
      if (e.button !== 0) return;
      if (state.pipetteActive) {
        var pos = canvasCoords(qs("#frameCanvas"), e);
        var frameCtx = qs("#frameCanvas").getContext("2d");
        var pixel = frameCtx.getImageData(pos.x, pos.y, 1, 1).data;
        var hsv = rgbToHsv(pixel[0], pixel[1], pixel[2]);
        setTargetColor(hsv.h, hsv.s, hsv.v);
        deactivatePipette();
        showToast("Sampled color from frame");
        return;
      }
      _cachedOverlayRect = overlay.getBoundingClientRect();
      var pos = canvasCoords(overlay, e, _cachedOverlayRect);
      var displayW = _cachedOverlayRect.width || overlay.width;
      var s = overlay.width / displayW;
      var ctx = overlay.getContext("2d");
      var hit = findHitRegion(pos.x, pos.y, s, ctx);
      if (hit && hit.handle === "resize") {
        var origR = state.regions[hit.name];
        state.resizingRegion = { name: hit.name, origRegion: { x: origR.x, y: origR.y, w: origR.w, h: origR.h } };
        state.activeRegion = hit.name;
        state.pendingRegion = null;
        document.body.style.cursor = "nwse-resize";
        document.body.style.userSelect = "none";
        renderRegionChips();
        updateRegionButtons();
        return;
      }
      if (hit && hit.handle === "move") {
        var r = regionToPixels(state.regions[hit.name]);
        var origM = state.regions[hit.name];
        state.draggingRegion = { name: hit.name, offsetX: pos.x - r.x, offsetY: pos.y - r.y, origRegion: { x: origM.x, y: origM.y, w: origM.w, h: origM.h } };
        state.activeRegion = hit.name;
        state.pendingRegion = null;
        document.body.style.cursor = "grabbing";
        document.body.style.userSelect = "none";
        renderRegionChips();
        updateRegionButtons();
        return;
      }
      state.drawingRegion = { startX: pos.x, startY: pos.y, endX: pos.x, endY: pos.y };
      state.pendingRegion = null;
      state.activeRegion = null;
      updateRegionButtons();
    });

    overlay.addEventListener("mousemove", function (e) {
      if (state.pipetteActive) return;
      var rect = _cachedOverlayRect || overlay.getBoundingClientRect();
      var pos = canvasCoords(overlay, e, rect);
      var displayW = rect.width || overlay.width;
      var s = overlay.width / displayW;
      if (state.resizingRegion) {
        var rName = state.resizingRegion.name;
        var rPx = regionToPixels(state.regions[rName]);
        var minSize = Math.round(20 * s);
        var newW = clamp(pos.x - rPx.x, minSize, overlay.width - rPx.x);
        var newH = clamp(pos.y - rPx.y, minSize, overlay.height - rPx.y);
        state.regions[rName] = Object.assign({}, state.regions[rName], { w: newW / overlay.width, h: newH / overlay.height });
        scheduleOverlayRender();
        return;
      }
      if (state.draggingRegion) {
        var d = state.draggingRegion;
        var dPx = regionToPixels(state.regions[d.name]);
        var newX = clamp(pos.x - d.offsetX, 0, overlay.width - dPx.w);
        var newY = clamp(pos.y - d.offsetY, 0, overlay.height - dPx.h);
        state.regions[d.name] = Object.assign({}, state.regions[d.name], { x: newX / overlay.width, y: newY / overlay.height });
        scheduleOverlayRender();
        return;
      }
      if (state.drawingRegion) {
        state.drawingRegion.endX = pos.x;
        state.drawingRegion.endY = pos.y;
        scheduleOverlayRender();
        return;
      }
      var ctx = overlay.getContext("2d");
      var hit = findHitRegion(pos.x, pos.y, s, ctx);
      state.hoveredRegion = hit;
      overlay.style.cursor = hit ? (hit.handle === "resize" ? "nwse-resize" : "grab") : "crosshair";
      scheduleOverlayRender();
    });

    overlay.addEventListener("mouseup", function (e) {
      if (state.pipetteActive) return;
      _cachedOverlayRect = null;
      if (state.resizingRegion) {
        var rName = state.resizingRegion.name;
        state.resizingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(rName);
        flushOverlayRender();
        updateRegionButtons();
        return;
      }
      if (state.draggingRegion) {
        var dName = state.draggingRegion.name;
        state.draggingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(dName);
        flushOverlayRender();
        updateRegionButtons();
        return;
      }
      if (!state.drawingRegion) return;
      var pos = canvasCoords(overlay, e);
      state.drawingRegion.endX = pos.x;
      state.drawingRegion.endY = pos.y;
      var r = normalizeRect(
        state.drawingRegion.startX, state.drawingRegion.startY,
        state.drawingRegion.endX, state.drawingRegion.endY
      );
      state.drawingRegion = null;
      if (r.w > 5 && r.h > 5) {
        state.pendingRegion = r;
      }
      flushOverlayRender();
      updateRegionButtons();
    });

    // Document-level listeners so drag/resize continues outside the canvas
    document.addEventListener("mousemove", function (e) {
      if (!state.resizingRegion && !state.draggingRegion) return;
      var rect = _cachedOverlayRect || overlay.getBoundingClientRect();
      var pos = canvasCoords(overlay, e, rect);
      var displayW = rect.width || overlay.width;
      var s = overlay.width / displayW;
      if (state.resizingRegion) {
        var rName = state.resizingRegion.name;
        var rPx = regionToPixels(state.regions[rName]);
        var minSize = Math.round(20 * s);
        var newW = clamp(pos.x - rPx.x, minSize, overlay.width - rPx.x);
        var newH = clamp(pos.y - rPx.y, minSize, overlay.height - rPx.y);
        state.regions[rName] = Object.assign({}, state.regions[rName], { w: newW / overlay.width, h: newH / overlay.height });
        scheduleOverlayRender();
      } else if (state.draggingRegion) {
        var d = state.draggingRegion;
        var dPx = regionToPixels(state.regions[d.name]);
        var newX = clamp(pos.x - d.offsetX, 0, overlay.width - dPx.w);
        var newY = clamp(pos.y - d.offsetY, 0, overlay.height - dPx.h);
        state.regions[d.name] = Object.assign({}, state.regions[d.name], { x: newX / overlay.width, y: newY / overlay.height });
        scheduleOverlayRender();
      }
    });

    document.addEventListener("mouseup", function () {
      _cachedOverlayRect = null;
      if (state.resizingRegion) {
        var rName = state.resizingRegion.name;
        state.resizingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(rName);
        flushOverlayRender();
        updateRegionButtons();
      } else if (state.draggingRegion) {
        var dName = state.draggingRegion.name;
        state.draggingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(dName);
        flushOverlayRender();
        updateRegionButtons();
      }
    });

    qs("#saveRegionBtn").addEventListener("click", function () {
      if (!state.pendingRegion) return;
      showRegionNameModal();
    });

    qs("#clearSelectionBtn").addEventListener("click", function () {
      state.pendingRegion = null;
      state.activeRegion = null;
      renderOverlay();
      updateRegionButtons();
    });

    qs("#deleteRegionBtn").addEventListener("click", function () {
      if (!state.activeRegion) return;
      var name = state.activeRegion;
      apiDelete("api/regions/" + encodeURIComponent(name))
        .then(function (data) {
          if (data.ok) {
            delete state.regions[name];
            state.activeRegion = null;
            renderRegionChips();
            renderOverlay();
            updateRegionButtons();
            updateRunButton();
            showToast("Region '" + name + "' deleted");
          }
        })
        .catch(function () { showToast("Failed to delete region"); });
    });

    // Toggle region labels
    var toggleLabelsBtn = qs("#toggleLabelsBtn");
    toggleLabelsBtn.appendChild(svgTagIcon());
    toggleLabelsBtn.addEventListener("click", function () {
      state.showRegionLabels = !state.showRegionLabels;
      updateRegionButtons();
      renderOverlay();
    });

    // Toggle region visibility
    var toggleRegionsBtn = qs("#toggleRegionsBtn");
    toggleRegionsBtn.appendChild(svgEyeIcon());
    toggleRegionsBtn.addEventListener("click", function () {
      state.showRegionOverlays = !state.showRegionOverlays;
      updateRegionButtons();
      renderOverlay();
    });

    // Stash all regions
    var stashBtn = qs("#stashRegionsBtn");
    stashBtn.appendChild(svgArchiveIcon());
    stashBtn.addEventListener("click", stashRegions);

    // Region name modal
    qs("#regionNameCancel").addEventListener("click", hideRegionNameModal);
    qs("#regionNameSave").addEventListener("click", function () {
      var name = qs("#regionNameInput").value.trim();
      if (!name) return;
      var desc = qs("#regionDescInput").value.trim();
      var r = state.pendingRegion;
      var canvas = qs("#overlayCanvas");
      var body = { name: name, x: r.x, y: r.y, w: r.w, h: r.h, canvas_width: canvas.width, canvas_height: canvas.height };
      if (desc) body.description = desc;
      apiPost("api/regions", body)
        .then(function (data) {
          if (data.ok) {
            var saved = data.region;
            if (desc) saved.description = desc;
            state.regions[name] = saved;
            state.pendingRegion = null;
            state.activeRegion = name;
            renderRegionChips();
            renderOverlay();
            updateRegionButtons();
            updateRunButton();
            hideRegionNameModal();
            showToast("Region '" + name + "' saved");
          } else {
            showToast(data.error || "Failed to save region");
          }
        })
        .catch(function () { showToast("Failed to save region"); });
    });

    qs("#regionNameInput").addEventListener("keydown", function (e) {
      if (e.key === "Enter") qs("#regionNameSave").click();
      if (e.key === "Escape") hideRegionNameModal();
    });

    // Convert vertical scroll to horizontal on region chips
    var chipsEl = qs("#regionChips");
    chipsEl.addEventListener("wheel", function (e) {
      if (chipsEl.scrollWidth > chipsEl.clientWidth) {
        e.preventDefault();
        chipsEl.scrollLeft += e.deltaY;
      }
    }, { passive: false });
    chipsEl.addEventListener("scroll", updateRegionChipsOverflow);
  }

  function showRegionNameModal() {
    var r = state.pendingRegion;
    qs("#regionNameInput").value = "";
    qs("#regionDescInput").value = "";
    qs("#regionCoords").textContent = r ? (r.x + ", " + r.y + " \u2014 " + r.w + "\u00d7" + r.h + " px") : "";
    qs("#regionNameModal").classList.remove("hidden");
    qs("#regionNameInput").focus();
  }

  function hideRegionNameModal() {
    qs("#regionNameModal").classList.add("hidden");
  }

  function updateRegionButtons() {
    var hasPending = !!state.pendingRegion;
    var hasActive = !!state.activeRegion;
    var hasRegions = Object.keys(state.regions).length > 0;
    qs("#saveRegionBtn").classList.toggle("hidden", !hasPending);
    qs("#clearSelectionBtn").classList.toggle("hidden", !hasPending && !hasActive);
    qs("#deleteRegionBtn").classList.toggle("hidden", !hasActive);

    var toggleLabelsBtn = qs("#toggleLabelsBtn");
    var toggleRegionsBtn = qs("#toggleRegionsBtn");
    toggleLabelsBtn.classList.toggle("hidden", !hasRegions);
    toggleRegionsBtn.classList.toggle("hidden", !hasRegions);
    qs("#stashRegionsBtn").classList.toggle("hidden", !hasRegions);
    toggleLabelsBtn.classList.toggle("active", state.showRegionLabels);
    toggleRegionsBtn.classList.toggle("active", state.showRegionOverlays);

    toggleRegionsBtn.innerHTML = "";
    toggleRegionsBtn.appendChild(state.showRegionOverlays ? svgEyeIcon() : svgEyeSlashIcon());
  }

  function renderRegionChips() {
    var container = qs("#regionChips");
    container.innerHTML = "";
    var names = Object.keys(state.regions);
    if (names.length === 0) {
      var hint = el("span", "region-hint", "Click and drag on the video to create a region");
      container.appendChild(hint);
      return;
    }
    names.forEach(function (name, i) {
      var color = regionColorForIndex(i);
      var chip = el("div", "region-chip" + (name === state.activeRegion ? " active" : ""));
      chip.style.color = color;
      var dot = el("span", "region-chip-dot");
      dot.style.background = color;
      chip.appendChild(dot);
      chip.appendChild(document.createTextNode(name));
      chip.addEventListener("click", function () {
        if (state.activeRegion === name) {
          state.activeRegion = null;
        } else {
          state.activeRegion = name;
          state.pendingRegion = null;
        }
        renderRegionChips();
        renderOverlay();
        updateRegionButtons();
        updateRunButton();
      });
      container.appendChild(chip);
    });
    renderRunRegionPicker();
    updateRegionChipsOverflow();
  }

  function updateRegionChipsOverflow() {
    var chips = qs("#regionChips");
    var wrapper = qs("#regionChipsScroll");
    wrapper.classList.toggle("has-overflow", chips.scrollWidth > chips.clientWidth && chips.scrollLeft + chips.clientWidth < chips.scrollWidth - 1);
  }

  // ---- Region stashing ----

  function stashRegions() {
    apiPost("api/stashes", {}).then(function (data) {
      if (!data.ok) return;
      state.stashes.push(data.stash);
      state.regions = {};
      state.activeRegion = null;
      state.pendingRegion = null;
      renderRegionChips();
      renderOverlay();
      updateRegionButtons();
      updateRunButton();
      renderStashCards();
      showToast("Regions stashed");
    });
  }

  function dismissStash(stashId) {
    apiDelete("api/stashes/" + stashId).then(function (data) {
      if (!data.ok) return;
      state.stashes = state.stashes.filter(function (s) { return s.id !== stashId; });
      // Remove any run-selected regions that belonged to this stash
      renderRunRegionPicker();
      renderStashCards();
      showToast("Stash dismissed");
    });
  }

  function restoreStash(stashId) {
    apiPost("api/stashes/" + stashId + "/restore", {}).then(function (data) {
      if (!data.ok) return;
      state.regions = data.regions || {};
      state.activeRegion = null;
      state.pendingRegion = null;
      state.stashes = state.stashes.filter(function (s) { return s.id !== stashId; });
      renderRegionChips();
      renderOverlay();
      updateRegionButtons();
      updateRunButton();
      renderStashCards();
      showToast("Regions restored");
    });
  }

  function renameStash(stashId, newName) {
    apiPut("api/stashes/" + stashId, { name: newName }).then(function (data) {
      if (!data.ok) return;
      for (var i = 0; i < state.stashes.length; i++) {
        if (state.stashes[i].id === stashId) {
          state.stashes[i].name = data.stash.name;
          break;
        }
      }
      renderRunRegionPicker();
    });
  }

  function renderStashCards() {
    var existing = qs("#stashArea");
    if (state.stashes.length === 0) {
      if (existing) existing.remove();
      return;
    }
    var area = existing || el("div");
    area.id = "stashArea";
    area.innerHTML = "";

    var MAX_DOTS = 5;

    state.stashes.forEach(function (stash) {
      var card = el("div", "stash-card");
      var regionNames = Object.keys(stash.regions);

      // Editable name
      var nameEl = el("span", "stash-card-name", stash.name);
      nameEl.setAttribute("contenteditable", "true");
      nameEl.setAttribute("spellcheck", "false");
      nameEl.addEventListener("blur", function () {
        var trimmed = nameEl.textContent.trim();
        if (trimmed && trimmed !== stash.name) {
          renameStash(stash.id, trimmed);
        } else {
          nameEl.textContent = stash.name;
        }
      });
      nameEl.addEventListener("keydown", function (e) {
        if (e.key === "Enter") { e.preventDefault(); nameEl.blur(); }
      });
      card.appendChild(nameEl);

      // Separator dot and count
      card.appendChild(el("span", "stash-card-sep", "\u00b7"));
      card.appendChild(el("span", "stash-card-count", regionNames.length + " region" + (regionNames.length !== 1 ? "s" : "")));

      // Colored dots (max 5, with fade-out on 6th)
      var dots = el("span", "stash-card-dots");
      var showCount = Math.min(regionNames.length, MAX_DOTS);
      for (var i = 0; i < showCount; i++) {
        var dot = el("span", "region-chip-dot");
        dot.style.background = regionColorForIndex(i);
        dot.title = regionNames[i];
        dots.appendChild(dot);
      }
      if (regionNames.length > MAX_DOTS) {
        var fadeDot = el("span", "region-chip-dot stash-dot-fade");
        fadeDot.style.background = regionColorForIndex(MAX_DOTS);
        dots.appendChild(fadeDot);
      }
      card.appendChild(dots);

      // Action buttons
      var actions = el("span", "stash-card-actions");

      var restoreBtn = el("button", "stash-card-action-btn");
      restoreBtn.title = "Restore regions";
      restoreBtn.appendChild(svgRestoreIcon());
      restoreBtn.addEventListener("click", function () { restoreStash(stash.id); });
      actions.appendChild(restoreBtn);

      var dismissBtn = el("button", "stash-card-action-btn");
      dismissBtn.title = "Dismiss stash";
      dismissBtn.appendChild(svgDismissIcon());
      dismissBtn.addEventListener("click", function () { dismissStash(stash.id); });
      actions.appendChild(dismissBtn);

      card.appendChild(actions);

      // Hover preview: show stashed regions on the overlay
      card.addEventListener("mouseenter", function () {
        state.previewRegions = stash.regions;
        renderOverlay();
      });
      card.addEventListener("mouseleave", function () {
        state.previewRegions = null;
        renderOverlay();
      });

      area.appendChild(card);
    });

    if (!existing) {
      var viewerSection = qs("#viewerSection");
      viewerSection.parentNode.insertBefore(area, viewerSection.nextSibling);
    }
  }

  function renderOverlay() {
    var canvas = qs("#overlayCanvas");
    if (!canvas.width || !canvas.height) return;
    var ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    // Scale factor: canvas pixels per display pixel, so chrome looks
    // the same physical size regardless of the underlying video resolution.
    var displayW = canvas.getBoundingClientRect().width || canvas.width;
    var s = canvas.width / displayW;

    // Draw saved regions (or preview regions when hovering a stash)
    var drawRegions = state.previewRegions || state.regions;
    if (state.showRegionOverlays) {
      var names = Object.keys(drawRegions);
      names.forEach(function (name, i) {
        var r = regionToPixels(drawRegions[name]);
        var color = regionColorForIndex(i);
        var isActive = (name === state.activeRegion);
        var isHovered = state.hoveredRegion && state.hoveredRegion.name === name;
        var showHandles = isActive || isHovered;
        ctx.strokeStyle = color;
        ctx.lineWidth = (isActive ? 2 : 1) * s;
        ctx.setLineDash(isActive ? [] : [6 * s, 3 * s]);
        ctx.strokeRect(r.x, r.y, r.w, r.h);
        if (isActive) {
          ctx.fillStyle = hexToRgba(color, 0.12);
          ctx.fillRect(r.x, r.y, r.w, r.h);
        }
        ctx.setLineDash([]);

        // Label with grip indicator
        if (state.showRegionLabels) {
          var lr = computeLabelRect(r, name, ctx, s);
          ctx.fillStyle = hexToRgba(color, 0.85);
          ctx.fillRect(lr.x, lr.y, lr.w, lr.h);
          if (showHandles) {
            ctx.fillStyle = "rgba(255,255,255,0.5)";
            var dotR = Math.round(1 * s);
            var gripColGap = Math.round(3 * s);
            var gripRowGap = Math.round(3 * s);
            var gx = lr.x + lr.gripPadL + dotR + Math.round(1 * s);
            var gy = lr.y + Math.round(lr.h / 2) - gripRowGap;
            for (var row = 0; row < 3; row++) {
              for (var col = 0; col < 2; col++) {
                ctx.beginPath();
                ctx.arc(gx + col * gripColGap, gy + row * gripRowGap, dotR, 0, Math.PI * 2);
                ctx.fill();
              }
            }
          }
          ctx.fillStyle = "#fff";
          ctx.fillText(name, lr.x + lr.gripPadL + lr.gripW + lr.pad, r.y - Math.round(4 * s));
        }

        // Resize handle (bottom-right corner, 3-dot triangle)
        if (showHandles) {
          var dotRr = Math.round(1.5 * s);
          var handlePad = Math.round(5 * s);
          var dotSpacing = Math.round(4 * s);
          var bx = r.x + r.w - handlePad;
          var by = r.y + r.h - handlePad;
          ctx.fillStyle = hexToRgba(color, 0.9);
          ctx.beginPath(); ctx.arc(bx, by, dotRr, 0, Math.PI * 2); ctx.fill();
          ctx.beginPath(); ctx.arc(bx - dotSpacing, by, dotRr, 0, Math.PI * 2); ctx.fill();
          ctx.beginPath(); ctx.arc(bx, by - dotSpacing, dotRr, 0, Math.PI * 2); ctx.fill();
        }
      });
    }

    // Drawing in progress
    if (state.drawingRegion) {
      var d = state.drawingRegion;
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = 1.5 * s;
      ctx.setLineDash([4 * s, 3 * s]);
      ctx.strokeRect(d.startX, d.startY, d.endX - d.startX, d.endY - d.startY);
      ctx.setLineDash([]);
      // Dimensions
      var w = Math.abs(d.endX - d.startX);
      var h = Math.abs(d.endY - d.startY);
      if (w > 20 && h > 20) {
        ctx.font = Math.round(11 * s) + "px " + getThemeColors().fontMono;
        ctx.fillStyle = "rgba(255,255,255,0.9)";
        ctx.fillText(w + "\u00d7" + h, Math.min(d.startX, d.endX) + Math.round(4 * s), Math.max(d.startY, d.endY) + Math.round(14 * s));
      }
    }

    // Pending (unsaved) region
    if (state.pendingRegion) {
      var p = state.pendingRegion;
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = 1.5 * s;
      ctx.setLineDash([]);
      ctx.strokeRect(p.x, p.y, p.w, p.h);
      ctx.fillStyle = "rgba(255, 255, 255, 0.08)";
      ctx.fillRect(p.x, p.y, p.w, p.h);
      ctx.font = Math.round(11 * s) + "px " + getThemeColors().fontMono;
      ctx.fillStyle = "rgba(255,255,255,0.9)";
      ctx.fillText(p.w + "\u00d7" + p.h + " px", p.x + Math.round(4 * s), p.y + p.h + Math.round(14 * s));
    }

    // Result overlay: template match bounding boxes / flow motion arrows
    if (state.resultOverlay) {
      ctx.setLineDash([]);
      if (state.resultOverlay.type === "template") {
        var matches = state.resultOverlay.data.matches || [];
        matches.forEach(function (m) {
          ctx.strokeStyle = taskTypeColor("template");
          ctx.lineWidth = 2 * s;
          ctx.strokeRect(m.x, m.y, m.w, m.h);
          ctx.fillStyle = hexToRgba(taskTypeColor("template"), 0.15);
          ctx.fillRect(m.x, m.y, m.w, m.h);
          ctx.font = Math.round(11 * s) + "px " + getThemeColors().fontMono;
          ctx.fillStyle = taskTypeColor("template");
          ctx.fillText((m.score * 100).toFixed(0) + "%", m.x + Math.round(3 * s), m.y - Math.round(4 * s));
        });
      } else if (state.resultOverlay.type === "flow") {
        var grid = state.resultOverlay.data.flow_grid || [];
        var fRegion = state.resultOverlay.region;
        if (fRegion && fRegion.w && grid.length) {
          var maxMag = 0;
          grid.forEach(function (c) { if (c.mag > maxMag) maxMag = c.mag; });
          if (maxMag > 0) {
            grid.forEach(function (c) {
              var px = fRegion.x + c.x * fRegion.w;
              var py = fRegion.y + c.y * fRegion.h;
              var norm = Math.min(c.mag / maxMag, 1);
              var arrowLen = norm * 20 * s;
              var rad = c.ang * Math.PI / 180;
              var ex = px + Math.cos(rad) * arrowLen;
              var ey = py + Math.sin(rad) * arrowLen;
              var alpha = norm * 0.8 + 0.2;
              ctx.strokeStyle = "rgba(99, 102, 241, " + alpha + ")";
              ctx.lineWidth = 1.5 * s;
              ctx.beginPath();
              ctx.moveTo(px, py);
              ctx.lineTo(ex, ey);
              ctx.stroke();
              var headLen = 4 * s;
              ctx.beginPath();
              ctx.moveTo(ex, ey);
              ctx.lineTo(ex - headLen * Math.cos(rad - 0.4), ey - headLen * Math.sin(rad - 0.4));
              ctx.moveTo(ex, ey);
              ctx.lineTo(ex - headLen * Math.cos(rad + 0.4), ey - headLen * Math.sin(rad + 0.4));
              ctx.stroke();
            });
          }
        }
      }
    }

    // Heatmap image overlay (semi-transparent composite)
    if (state.heatmapOverlay && state.heatmapOverlay._img) {
      var hm = state.heatmapOverlay;
      ctx.globalAlpha = 0.5;
      if (hm.type === "template") {
        ctx.drawImage(hm._img, 0, 0, canvas.width, canvas.height);
      } else if (hm.type === "flow") {
        var rPx = regionToPixels(state.regions[hm.region] || {});
        if (rPx && rPx.w) {
          ctx.drawImage(hm._img, rPx.x, rPx.y, rPx.w, rPx.h);
        }
      }
      ctx.globalAlpha = 1.0;
    }
  }

  function hexToRgba(hex, alpha) {
    var r = parseInt(hex.slice(1, 3), 16);
    var g = parseInt(hex.slice(3, 5), 16);
    var b = parseInt(hex.slice(5, 7), 16);
    return "rgba(" + r + "," + g + "," + b + "," + alpha + ")";
  }

  // ---- Timeline ----

  function initTimeline() {
    qs("#zoomInBtn").appendChild(svgPlusIcon());
    qs("#zoomOutBtn").appendChild(svgMinusIcon());
    var canvas = qs("#timelineCanvas");
    sizeTimelineCanvas();
    window.addEventListener("resize", function () {
      _cachedOverlayRect = null;
      sizeTimelineCanvas();
    });

    canvas.addEventListener("click", function (e) {
      if (state.timelineDragging) return;
      var ts = timelineXToTime(e);
      if (ts !== null) {
        state.resultOverlay = null;
        loadFrame(ts);
      }
    });

    canvas.addEventListener("wheel", function (e) {
      e.preventDefault();
      if (!state.videoInfo) return;
      var dur = state.videoInfo.duration;
      var zoomFactor = e.deltaY < 0 ? 1.3 : 1 / 1.3;
      var mouseTs = timelineXToTime(e);
      var oldZoom = state.timelineZoom;
      state.timelineZoom = clamp(oldZoom * zoomFactor, 1, 200);
      if (mouseTs !== null && state.timelineZoom > 1) {
        var rect = canvas.getBoundingClientRect();
        var frac = (e.clientX - rect.left) / rect.width;
        var visLen = dur / state.timelineZoom;
        state.timelineOffset = clamp(mouseTs - frac * visLen, 0, dur - visLen);
      }
      if (state.timelineZoom <= 1) state.timelineOffset = 0;
      renderTimeline();
    }, { passive: false });

    var dragStart = null;
    var scrubbing = false;
    var scrubRaf = 0;
    canvas.addEventListener("mousedown", function (e) {
      hideSsTooltip();
      _lastTimelineHit = null;
      if (state.timelineZoom > 1) {
        // Zoomed in: drag to pan
        dragStart = { x: e.clientX, offset: state.timelineOffset };
        state.timelineDragging = false;
      } else {
        // Not zoomed: drag to scrub
        scrubbing = true;
        var ts = timelineXToTime(e);
        if (ts !== null) loadFrame(ts);
      }
    });
    document.addEventListener("mousemove", function (e) {
      if (scrubbing) {
        var ts = timelineXToTime(e);
        if (ts !== null) {
          seekPlayhead(ts);
          if (!scrubRaf) {
            scrubRaf = requestAnimationFrame(function () {
              scrubRaf = 0;
              loadFrame(state.currentTimestamp);
            });
          }
        }
        return;
      }
      if (!dragStart) return;
      var dx = e.clientX - dragStart.x;
      if (Math.abs(dx) > 3) state.timelineDragging = true;
      if (!state.videoInfo) return;
      var dur = state.videoInfo.duration;
      var visLen = dur / state.timelineZoom;
      var rect = canvas.getBoundingClientRect();
      var dtSec = -(dx / rect.width) * visLen;
      state.timelineOffset = clamp(dragStart.offset + dtSec, 0, dur - visLen);
      renderTimeline();
    });
    document.addEventListener("mouseup", function () {
      if (scrubbing) {
        scrubbing = false;
        if (scrubRaf) { cancelAnimationFrame(scrubRaf); scrubRaf = 0; }
        loadFrame(state.currentTimestamp);
      }
      if (dragStart) {
        setTimeout(function () { state.timelineDragging = false; }, 50);
        dragStart = null;
      }
    });

    var _ssTooltipRaf = 0;
    var _lastTimelineHit = null;
    canvas.addEventListener("mousemove", function (e) {
      if (scrubbing || dragStart) {
        hideSsTooltip();
        _lastTimelineHit = null;
        return;
      }
      if (_ssTooltipRaf) return;
      var cx = e.clientX;
      var cy = e.clientY;
      _ssTooltipRaf = requestAnimationFrame(function () {
        _ssTooltipRaf = 0;
        var hit = hitTestTimeline(cx, cy);
        if (hit) {
          _lastTimelineHit = hit;
          showSsTooltip(hit, cx, cy);
        } else if (_lastTimelineHit) {
          _lastTimelineHit = null;
          hideSsTooltip();
        }
      });
    });
    canvas.addEventListener("mouseleave", function () {
      _lastTimelineHit = null;
      hideSsTooltip();
    });

    qs("#zoomInBtn").addEventListener("click", function () {
      if (!state.videoInfo) return;
      state.timelineZoom = clamp(state.timelineZoom * 1.5, 1, 200);
      clampTimelineOffset();
      renderTimeline();
    });
    qs("#zoomOutBtn").addEventListener("click", function () {
      state.timelineZoom = clamp(state.timelineZoom / 1.5, 1, 200);
      if (state.timelineZoom <= 1) state.timelineOffset = 0;
      clampTimelineOffset();
      renderTimeline();
    });
    qs("#zoomResetBtn").addEventListener("click", function () {
      state.timelineZoom = 1;
      state.timelineOffset = 0;
      renderTimeline();
    });
    qs("#setInBtn").addEventListener("click", function () {
      state.inMarker = state.currentTimestamp;
      if (state.outMarker !== null && state.inMarker > state.outMarker) state.outMarker = null;
      updateMarkerInfo();
      renderTimeline();
    });
    qs("#setOutBtn").addEventListener("click", function () {
      state.outMarker = state.currentTimestamp;
      if (state.inMarker !== null && state.outMarker < state.inMarker) state.inMarker = null;
      updateMarkerInfo();
      renderTimeline();
    });
    qs("#clearMarkersBtn").addEventListener("click", function () {
      state.inMarker = null;
      state.outMarker = null;
      updateMarkerInfo();
      renderTimeline();
    });
  }

  function clampTimelineOffset() {
    if (!state.videoInfo) return;
    var dur = state.videoInfo.duration;
    var visLen = dur / state.timelineZoom;
    state.timelineOffset = clamp(state.timelineOffset, 0, Math.max(0, dur - visLen));
  }

  function sizeTimelineCanvas() {
    var canvas = qs("#timelineCanvas");
    var rect = canvas.getBoundingClientRect();
    canvas.width = Math.floor(rect.width);
    canvas.height = TIMELINE_CANVAS_HEIGHT;
    var ph = qs("#playheadCanvas");
    ph.width = canvas.width;
    ph.height = canvas.height;
    renderTimeline();
    renderPlayhead();
  }

  function timelineXToTime(event) {
    if (!state.videoInfo) return null;
    var canvas = qs("#timelineCanvas");
    var rect = canvas.getBoundingClientRect();
    var frac = (event.clientX - rect.left) / rect.width;
    frac = clamp(frac, 0, 1);
    var dur = state.videoInfo.duration;
    var visLen = dur / state.timelineZoom;
    return state.timelineOffset + frac * visLen;
  }

  function updateMarkerInfo() {
    var info = qs("#markerInfo");
    var clearBtn = qs("#clearMarkersBtn");
    if (state.inMarker === null && state.outMarker === null) {
      info.textContent = "";
      clearBtn.classList.add("hidden");
      return;
    }
    clearBtn.classList.remove("hidden");
    var parts = [];
    if (state.inMarker !== null) parts.push("In: " + formatTimestamp(state.inMarker));
    if (state.outMarker !== null) parts.push("Out: " + formatTimestamp(state.outMarker));
    info.textContent = parts.join("  ");
  }

  function renderTimeline() {
    var canvas = qs("#timelineCanvas");
    var ctx = canvas.getContext("2d");
    var w = canvas.width;
    var h = canvas.height;
    var dur = state.videoInfo ? state.videoInfo.duration : 0;

    var tc = getThemeColors();

    ctx.clearRect(0, 0, w, h);

    // Background
    ctx.fillStyle = tc.surfaceAlt;
    ctx.fillRect(0, 0, w, h);

    if (dur <= 0) {
      ctx.fillStyle = tc.textDim;
      ctx.font = "12px -apple-system, sans-serif";
      ctx.textAlign = "center";
      ctx.fillText("No video loaded", w / 2, h / 2 + 4);
      ctx.textAlign = "start";
      return;
    }

    var visLen = dur / state.timelineZoom;
    var visStart = state.timelineOffset;
    var visEnd = visStart + visLen;

    function timeToX(t) {
      return ((t - visStart) / visLen) * w;
    }

    // Time ruler ticks
    var tickInterval = computeTickInterval(visLen);
    var firstTick = Math.ceil(visStart / tickInterval) * tickInterval;
    ctx.strokeStyle = tc.border;
    ctx.fillStyle = tc.textDim;
    ctx.font = "10px " + tc.fontMono;
    ctx.textAlign = "center";
    ctx.lineWidth = 1;
    for (var t = firstTick; t <= visEnd; t += tickInterval) {
      var x = timeToX(t);
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, 8);
      ctx.stroke();
      ctx.fillText(formatDuration(t), x, 18);
    }
    ctx.textAlign = "start";

    // In/Out marker shading
    if (state.inMarker !== null || state.outMarker !== null) {
      ctx.fillStyle = "rgba(0,0,0,0.12)";
      if (state.inMarker !== null) {
        var inX = timeToX(state.inMarker);
        ctx.fillRect(0, 0, Math.max(0, inX), h);
      }
      if (state.outMarker !== null) {
        var outX = timeToX(state.outMarker);
        ctx.fillRect(outX, 0, w - outX, h);
      }
      // Marker lines
      ctx.strokeStyle = "#16a34a";
      ctx.lineWidth = 2;
      if (state.inMarker !== null) {
        var ix = timeToX(state.inMarker);
        ctx.beginPath(); ctx.moveTo(ix, 0); ctx.lineTo(ix, h); ctx.stroke();
      }
      if (state.outMarker !== null) {
        var ox = timeToX(state.outMarker);
        ctx.beginPath(); ctx.moveTo(ox, 0); ctx.lineTo(ox, h); ctx.stroke();
      }
    }

    // Result markers from completed and running tasks
    var resultY = 24;
    var resultH = h - resultY - 6;
    var focused = focusedTaskId();
    _timelineHitRects = [];
    state.tasks.forEach(function (task) {
      if ((task.status !== "completed" && task.status !== "running") || !task.result) return;
      if (task.participant && task.participant !== state.selectedParticipant) return;
      var color = taskTypeColor(task.type);
      var dimmed = focused && task.id !== focused;
      if (task.type === "color" && task.status === "completed") {
        // Completed color: merged spans
        ctx.fillStyle = hexToRgba(color, dimmed ? 0.10 : 0.35);
        task.result.forEach(function (span) {
          var x1 = timeToX(span.start);
          var x2 = timeToX(span.end);
          var rw = Math.max(x2 - x1, 2);
          ctx.fillRect(x1, resultY, rw, resultH);
          _timelineHitRects.push({ x1: x1, x2: x1 + rw, y: resultY, h: resultH, task: task, result: span });
        });
      } else if (task.type === "timelapse") {
        // No timeline markers for timelapse
      } else {
        // Point markers (change, similarity, text, numbers, running color)
        ctx.strokeStyle = dimmed ? hexToRgba(color, 0.15) : color;
        ctx.lineWidth = 1.5;
        var results = task.result || [];
        results.forEach(function (r) {
          var ts = r.timestamp !== undefined ? r.timestamp : r.start;
          if (ts === undefined) return;
          var x = timeToX(ts);
          ctx.beginPath();
          ctx.moveTo(x, resultY);
          ctx.lineTo(x, resultY + resultH);
          ctx.stroke();
          _timelineHitRects.push({ x1: x - 3, x2: x + 3, y: resultY, h: resultH, task: task, result: r });
        });
      }
    });

    renderTimelineLegend();
    renderPlayhead();
  }

  function renderPlayhead() {
    var canvas = qs("#playheadCanvas");
    if (!canvas.width || !canvas.height) return;
    var ctx = canvas.getContext("2d");
    var w = canvas.width;
    var h = canvas.height;
    var dur = state.videoInfo ? state.videoInfo.duration : 0;
    ctx.clearRect(0, 0, w, h);
    if (dur <= 0) return;
    var visLen = dur / state.timelineZoom;
    var visStart = state.timelineOffset;
    var px = ((state.currentTimestamp - visStart) / visLen) * w;
    var tc = getThemeColors();
    ctx.strokeStyle = tc.accent;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(px, 0);
    ctx.lineTo(px, h);
    ctx.stroke();
    ctx.fillStyle = tc.accent;
    ctx.beginPath();
    ctx.moveTo(px - 5, 0);
    ctx.lineTo(px + 5, 0);
    ctx.lineTo(px, 6);
    ctx.closePath();
    ctx.fill();
  }

  function hitTestTimeline(clientX, clientY) {
    var canvas = qs("#timelineCanvas");
    var rect = canvas.getBoundingClientRect();
    var mx = clientX - rect.left;
    var my = clientY - rect.top;
    for (var i = _timelineHitRects.length - 1; i >= 0; i--) {
      var hr = _timelineHitRects[i];
      if (mx >= hr.x1 && mx <= hr.x2 && my >= hr.y && my <= hr.y + hr.h) {
        return hr;
      }
    }
    return null;
  }

  function showSsTooltip(hit, clientX, clientY) {
    var tip = qs("#ssTooltip");
    if (!tip) return;
    var color = taskTypeColor(hit.task.type);
    tip.innerHTML = "";
    tip.style.borderLeft = "3px solid " + color;

    var header = el("div", "ss-tooltip-header");
    var icon = buildTypeIconSvg(hit.task.type);
    if (icon) {
      icon.style.color = color;
      icon.style.flexShrink = "0";
      header.appendChild(icon);
    }
    var label = hit.task.type.charAt(0).toUpperCase() + hit.task.type.slice(1);
    header.appendChild(el("strong", "", label));
    tip.appendChild(header);

    var r = hit.result;
    var timeStr;
    if (r.start !== undefined && r.end !== undefined) {
      timeStr = formatTimestamp(r.start) + " \u2013 " + formatTimestamp(r.end);
    } else {
      var ts = r.timestamp !== undefined ? r.timestamp : r.start;
      timeStr = formatTimestamp(ts);
    }
    tip.appendChild(el("span", "ss-tooltip-time", timeStr));

    var details = el("div", "ss-tooltip-details");
    details.appendChild(el("span", "", hit.task.participant + " \u00b7 " + (hit.task.region || "")));
    if (hit.task.type === "color" && r.duration !== undefined) {
      details.appendChild(el("span", "", "Duration: " + r.duration.toFixed(1) + "s"));
    } else if (hit.task.type === "change" && r.magnitude !== undefined) {
      details.appendChild(el("span", "", "Magnitude: " + (r.magnitude * 100).toFixed(1) + "%"));
    } else if (hit.task.type === "similarity" && r.score !== undefined) {
      details.appendChild(el("span", "", "Score: " + (r.score * 100).toFixed(1) + "%"));
    } else if (hit.task.type === "text" && r.text_found) {
      details.appendChild(el("span", "", "Found: " + r.text_found));
      if (r.confidence !== undefined) details.appendChild(el("span", "", "Confidence: " + (r.confidence * 100).toFixed(0) + "%"));
    } else if (hit.task.type === "numbers" && r.number_found !== undefined) {
      details.appendChild(el("span", "", "Found: " + r.number_found));
    } else if (hit.task.type === "template" && r.best_score !== undefined) {
      details.appendChild(el("span", "", "Score: " + (r.best_score * 100).toFixed(1) + "%"));
      if (r.match_count !== undefined) details.appendChild(el("span", "", "Matches: " + r.match_count));
    } else if (hit.task.type === "flow" && r.magnitude !== undefined) {
      details.appendChild(el("span", "", "Magnitude: " + r.magnitude.toFixed(2)));
      if (r.angle !== undefined) details.appendChild(el("span", "", "Direction: " + r.angle.toFixed(0) + "\u00b0"));
    } else if (hit.task.type === "scene" && r.scene_name) {
      details.appendChild(el("span", "", "Scene: " + r.scene_name));
      if (r.score !== undefined) details.appendChild(el("span", "", "Score: " + (r.score * 100).toFixed(1) + "%"));
    }
    tip.appendChild(details);

    tip.classList.remove("hidden");
    positionSsTooltip(tip, clientX, clientY);
  }

  function positionSsTooltip(tip, clientX, clientY) {
    var x = clientX + 12;
    var y = clientY + 12;
    var rect = tip.getBoundingClientRect();
    if (x + rect.width > window.innerWidth - 8) {
      x = clientX - rect.width - 12;
    }
    if (y + rect.height > window.innerHeight - 8) {
      y = clientY - rect.height - 12;
    }
    tip.style.left = x + "px";
    tip.style.top = y + "px";
  }

  function hideSsTooltip() {
    var tip = qs("#ssTooltip");
    if (tip) tip.classList.add("hidden");
  }

  function computeTickInterval(visibleSeconds) {
    var candidates = [1, 2, 5, 10, 15, 30, 60, 120, 300, 600, 900, 1800, 3600];
    for (var i = 0; i < candidates.length; i++) {
      if (visibleSeconds / candidates[i] <= 20) return candidates[i];
    }
    return 3600;
  }

  function renderTimelineLegend() {
    var container = qs("#timelineLegend");
    container.innerHTML = "";
    var hasTypes = {};
    state.tasks.forEach(function (t) {
      if ((t.status === "completed" || t.status === "running") && t.result) hasTypes[t.type] = true;
    });
    var types = Object.keys(hasTypes);
    if (types.length === 0) return;
    var focused = focusedTaskId();
    var focusedType = focused ? (findTask(focused) || {}).type : null;
    types.forEach(function (type) {
      var item = el("span", "legend-item");
      var dot = el("span", "legend-dot");
      dot.style.background = taskTypeColor(type);
      item.appendChild(dot);
      item.appendChild(document.createTextNode(type));
      if (focusedType && type !== focusedType) {
        item.style.opacity = "0.3";
      }
      container.appendChild(item);
    });
  }

  // ---- Workflow tabs + params ----

  function initWorkflowTabs() {
    qsa(".wf-tab").forEach(function (tab) {
      tab.addEventListener("click", function () {
        state.activeWorkflow = tab.dataset.type;
        qsa(".wf-tab").forEach(function (t) { t.classList.remove("active"); });
        tab.classList.add("active");
        renderWorkflowParams();
      });
    });
    renderWorkflowParams();
  }

  function renderIntervalSlot(inputId, min, max, def, step) {
    var slot = qs("#workflowIntervalSlot");
    if (!slot) return;
    slot.innerHTML = "";
    slot.appendChild(el("span", null, "Interval\u202f(s)"));
    var ctrl = el("div", "param-control");
    ctrl.appendChild(numberInput(inputId, min, max, def, step));
    slot.appendChild(ctrl);
  }

  function renderWorkflowParams() {
    var container = qs("#workflowParams");
    container.innerHTML = "";
    var intervalSlot = qs("#workflowIntervalSlot");
    if (intervalSlot) intervalSlot.innerHTML = "";
    var type = state.activeWorkflow;

    if (type === "color") {
      // Color picker group: palette + brightness + hex/pipette
      var pickerGroup = el("div", "color-picker-group");

      // Hue-Saturation palette canvas
      var palette = document.createElement("canvas");
      palette.id = "colorPalette";
      palette.className = "color-palette-canvas";
      pickerGroup.appendChild(palette);

      // Brightness strip
      var bright = document.createElement("canvas");
      bright.id = "colorBrightness";
      bright.className = "color-brightness-strip";
      pickerGroup.appendChild(bright);

      // Color input row: preview swatch + hex input + pipette button
      var inputRow = el("div", "color-input-row");
      var preview = el("div", "color-preview");
      preview.id = "colorPreview";
      inputRow.appendChild(preview);

      var hexInput = document.createElement("input");
      hexInput.type = "text";
      hexInput.id = "paramColorHex";
      hexInput.className = "color-hex-input";
      hexInput.placeholder = "#000000";
      hexInput.maxLength = 7;
      inputRow.appendChild(hexInput);

      var pipetteBtn = el("button", "btn btn-small btn-pipette");
      pipetteBtn.id = "pipetteBtn";
      pipetteBtn.innerHTML = '<svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M15 4C15 5.39788 14.0439 6.57245 12.75 6.90549V8.5C12.75 8.69891 12.671 8.88968 12.5303 9.03033L12.0303 9.53033C11.7374 9.82322 11.2626 9.82322 10.9697 9.53033L10.25 8.81069L5.57322 13.4875C5.24503 13.8157 4.79992 14.0001 4.33579 14.0001H3.66421C3.59791 14.0001 3.53432 14.0264 3.48744 14.0733L2.78033 14.7804C2.63968 14.921 2.44891 15.0001 2.25 15.0001C2.05109 15.0001 1.86032 14.921 1.71967 14.7804L1.21967 14.2804C0.926777 13.9875 0.926777 13.5126 1.21967 13.2197L1.92678 12.5126C1.97366 12.4657 2 12.4021 2 12.3358V11.6643C2 11.2001 2.18437 10.755 2.51256 10.4268L7.18937 5.75003L6.46967 5.03033C6.17678 4.73744 6.17678 4.26256 6.46967 3.96967L6.96967 3.46967C7.11032 3.32902 7.30109 3.25 7.5 3.25H9.09451C9.42755 1.95608 10.6021 1 12 1C13.6569 1 15 2.34315 15 4ZM9.18937 7.75003L8.25003 6.81069L3.57322 11.4875C3.52634 11.5344 3.5 11.598 3.5 11.6643V12.3358C3.5 12.3938 3.49713 12.4514 3.49146 12.5086C3.54862 12.5029 3.60627 12.5001 3.66421 12.5001H4.33579C4.40209 12.5001 4.46568 12.4737 4.51256 12.4268L9.18937 7.75003Z"/></svg>';
      pipetteBtn.title = "Pick color from video frame";
      pipetteBtn.addEventListener("click", function () {
        if (state.pipetteActive) deactivatePipette();
        else activatePipette();
      });
      inputRow.appendChild(pipetteBtn);

      var sampleBtn = el("button", "btn btn-small", "From Region");
      sampleBtn.addEventListener("click", sampleColorFromRegion);
      inputRow.appendChild(sampleBtn);

      pickerGroup.appendChild(inputRow);

      // Hidden inputs for gatherWorkflowParams() contract
      var hiddenH = document.createElement("input");
      hiddenH.type = "hidden"; hiddenH.id = "paramColorH"; hiddenH.value = "90";
      var hiddenS = document.createElement("input");
      hiddenS.type = "hidden"; hiddenS.id = "paramColorS"; hiddenS.value = "200";
      var hiddenV = document.createElement("input");
      hiddenV.type = "hidden"; hiddenV.id = "paramColorV"; hiddenV.value = "200";
      pickerGroup.appendChild(hiddenH);
      pickerGroup.appendChild(hiddenS);
      pickerGroup.appendChild(hiddenV);

      container.appendChild(pickerGroup);

      // Palette canvas events
      var paletteDragging = false;
      var brightDragging = false;
      function pickFromPalette(e) {
        var rect = palette.getBoundingClientRect();
        var x = clamp(e.clientX - rect.left, 0, rect.width);
        var y = clamp(e.clientY - rect.top, 0, rect.height);
        var h = Math.round((x / rect.width) * 180);
        var s = Math.round((1 - y / rect.height) * 255);
        var curV = parseFloat(qs("#paramColorV").value) || 0;
        setTargetColor(h, s, curV);
      }
      palette.addEventListener("mousedown", function (e) {
        e.preventDefault();
        paletteDragging = true;
        pickFromPalette(e);
      });

      // Brightness strip events
      function pickFromBrightness(e) {
        var rect = bright.getBoundingClientRect();
        var x = clamp(e.clientX - rect.left, 0, rect.width);
        var v = Math.round((x / rect.width) * 255);
        var curH = parseFloat(qs("#paramColorH").value) || 0;
        var curS = parseFloat(qs("#paramColorS").value) || 0;
        setTargetColor(curH, curS, v);
      }
      bright.addEventListener("mousedown", function (e) {
        e.preventDefault();
        brightDragging = true;
        pickFromBrightness(e);
      });

      // Remove previous document-level listeners before adding new ones
      if (_paletteDocListeners) {
        document.removeEventListener("mousemove", _paletteDocListeners.move);
        document.removeEventListener("mouseup", _paletteDocListeners.up);
      }
      function onDocMove(e) {
        if (paletteDragging) pickFromPalette(e);
        if (brightDragging) pickFromBrightness(e);
      }
      function onDocUp() { paletteDragging = false; brightDragging = false; }
      document.addEventListener("mousemove", onDocMove);
      document.addEventListener("mouseup", onDocUp);
      _paletteDocListeners = { move: onDocMove, up: onDocUp };

      // Hex input event
      hexInput.addEventListener("input", function () {
        var rgb = hexToRgb(hexInput.value);
        if (rgb) {
          var hsv = rgbToHsv(rgb.r, rgb.g, rgb.b);
          // Update hidden inputs and visuals without re-setting the hex field
          var hEl = qs("#paramColorH"), sEl = qs("#paramColorS"), vEl = qs("#paramColorV");
          if (hEl) hEl.value = hsv.h;
          if (sEl) sEl.value = hsv.s;
          if (vEl) vEl.value = hsv.v;
          updateColorPreview();
          renderColorPalette();
          renderBrightnessStrip();
        }
      });

      // Unified tolerance slider
      var tolSlider = rangeInput("paramColorTol", 0, 100, 30);
      addParamRow(container, "Tolerance", tolSlider, "paramColorTolVal");
      tolSlider.addEventListener("input", function () {
        renderColorPalette();
      });
      renderIntervalSlot("paramColorInterval", 0.5, 60, 1.0, 0.5);

      // Initial render of palette and preview
      renderColorPalette();
      renderBrightnessStrip();
      updateColorPreview();
      var initRgb = hsvToRgb(90, 200, 200);
      hexInput.value = rgbToHex(initRgb.r, initRgb.g, initRgb.b);
    } else if (type === "change") {
      addParamRow(container, "Threshold", rangeInput("paramChangeThresh", 0.01, 0.50, 0.03, 0.01), "paramChangeThreshVal");
      addParamRow(container, "Noise Thr.", rangeInput("paramChangeNoise", 0, 100, 30, 1), "paramChangeNoiseVal");
      renderIntervalSlot("paramChangeInterval", 0.5, 60, 1.0, 0.5);
    } else if (type === "similarity") {
      var refRow = el("div", "param-row");
      var refLabel = el("span", "param-label", "Reference");
      var refControl = el("div", "param-control");
      var refBtn = el("button", "btn btn-small", "Capture Current Frame");
      refBtn.addEventListener("click", function () {
        state.referenceTimestamp = state.currentTimestamp;
        renderWorkflowParams();
        showToast("Reference frame captured at " + formatTimestamp(state.currentTimestamp));
      });
      refControl.appendChild(refBtn);
      if (state.referenceTimestamp !== null) {
        var refTs = el("span", "param-value", formatTimestamp(state.referenceTimestamp));
        refControl.appendChild(refTs);
      }
      refRow.appendChild(refLabel);
      refRow.appendChild(refControl);
      container.appendChild(refRow);
      addParamRow(container, "Threshold", rangeInput("paramSimThresh", 0.50, 1.00, 0.90, 0.01), "paramSimThreshVal");
      renderIntervalSlot("paramSimInterval", 0.5, 60, 1.0, 0.5);
    } else if (type === "text") {
      addParamRow(container, "Search text", textInput("paramTextSearch", "Enter text to find..."));
      addParamRow(container, "Fuzzy Thr.", rangeInput("paramTextFuzzy", 0.50, 1.00, 0.80, 0.01), "paramTextFuzzyVal");
      renderIntervalSlot("paramTextInterval", 0.5, 60, 2.0, 0.5);
      var langRow = el("div", "param-row");
      langRow.appendChild(el("span", "param-label", "Language"));
      var langControl = el("div", "param-control");
      var langSel = document.createElement("select");
      langSel.id = "paramTextLang";
      [["en", "English"], ["es", "Spanish"], ["fr", "French"], ["de", "German"],
       ["ja", "Japanese"], ["ko", "Korean"], ["zh", "Chinese"]].forEach(function (pair) {
        var opt = el("option", null, pair[1]);
        opt.value = pair[0];
        langSel.appendChild(opt);
      });
      langControl.appendChild(langSel);
      langRow.appendChild(langControl);
      container.appendChild(langRow);
    } else if (type === "numbers") {
      // Operator selector
      var opRow = el("div", "param-row");
      opRow.appendChild(el("span", "param-label", "Operator"));
      var opControl = el("div", "param-control");
      var opSel = document.createElement("select");
      opSel.id = "paramNumOperator";
      [["gt", "Greater than (>)"], ["lt", "Less than (<)"], ["eq", "Equal to (=)"],
       ["gte", "Greater or equal (\u2265)"], ["lte", "Less or equal (\u2264)"], ["range", "In range"]].forEach(function (pair) {
        var opt = el("option", null, pair[1]);
        opt.value = pair[0];
        opSel.appendChild(opt);
      });
      opSel.addEventListener("change", function () {
        var rangeRow = qs("#paramNumRangeRow");
        var targetRow = qs("#paramNumTargetRow");
        if (opSel.value === "range") {
          if (rangeRow) rangeRow.style.display = "";
          if (targetRow) targetRow.style.display = "none";
        } else {
          if (rangeRow) rangeRow.style.display = "none";
          if (targetRow) targetRow.style.display = "";
        }
      });
      opControl.appendChild(opSel);
      opRow.appendChild(opControl);
      container.appendChild(opRow);
      // Target value (for non-range operators)
      var targetRow = el("div", "param-row");
      targetRow.id = "paramNumTargetRow";
      targetRow.appendChild(el("span", "param-label", "Target value"));
      var targetCtrl = el("div", "param-control");
      targetCtrl.appendChild(numberInput("paramNumTarget", -999999, 999999, 100, 1));
      targetRow.appendChild(targetCtrl);
      container.appendChild(targetRow);
      // Range min/max (hidden by default)
      var numRangeRow = el("div", "param-row");
      numRangeRow.id = "paramNumRangeRow";
      numRangeRow.style.display = "none";
      numRangeRow.appendChild(el("span", "param-label", "Range"));
      var rangeCtrl = el("div", "param-control");
      rangeCtrl.appendChild(numberInput("paramNumMin", -999999, 999999, 0, 1));
      rangeCtrl.appendChild(el("span", "param-value", "\u2013"));
      rangeCtrl.appendChild(numberInput("paramNumMax", -999999, 999999, 100, 1));
      numRangeRow.appendChild(rangeCtrl);
      container.appendChild(numRangeRow);
      renderIntervalSlot("paramNumInterval", 0.5, 60, 2.0, 0.5);
    } else if (type === "timelapse") {
      addParamRow(container, "Speed", numberInput("paramTlSpeed", 2, 100, 10, 1));
      var fmtRow = el("div", "param-row");
      fmtRow.appendChild(el("span", "param-label", "Format"));
      var fmtControl = el("div", "param-control");
      var fmtSel = document.createElement("select");
      fmtSel.id = "paramTlFormat";
      [["mp4", "Video (.mp4)"], ["gif", "GIF (.gif)"]].forEach(function (pair) {
        var opt = el("option", null, pair[1]);
        opt.value = pair[0];
        fmtSel.appendChild(opt);
      });
      fmtControl.appendChild(fmtSel);
      fmtRow.appendChild(fmtControl);
      container.appendChild(fmtRow);
    } else if (type === "template") {
      var tmplRefRow = el("div", "param-row");
      tmplRefRow.appendChild(el("span", "param-label", "Template"));
      var tmplRefCtrl = el("div", "param-control");
      var tmplCapBtn = el("button", "btn btn-small", "Capture Region");
      tmplCapBtn.addEventListener("click", function () {
        state.referenceTimestamp = state.currentTimestamp;
        state.uploadedTemplate = null;
        renderWorkflowParams();
        showToast("Template captured at " + formatTimestamp(state.currentTimestamp));
      });
      tmplRefCtrl.appendChild(tmplCapBtn);

      // Upload PNG button
      var tmplFileInput = document.createElement("input");
      tmplFileInput.type = "file";
      tmplFileInput.accept = "image/png";
      tmplFileInput.style.display = "none";
      tmplFileInput.addEventListener("change", function () {
        var file = tmplFileInput.files[0];
        if (!file) return;
        var reader = new FileReader();
        reader.onload = function (e) {
          // Store base64 data (strip the data:image/png;base64, prefix)
          var dataUrl = e.target.result;
          var b64 = dataUrl.split(",")[1];
          state.uploadedTemplate = { name: file.name, data: b64 };
          state.referenceTimestamp = null;
          renderWorkflowParams();
          showToast("Template loaded: " + file.name);
        };
        reader.readAsDataURL(file);
      });
      var tmplUploadBtn = el("button", "btn btn-small", "Upload PNG");
      tmplUploadBtn.addEventListener("click", function () { tmplFileInput.click(); });
      tmplRefCtrl.appendChild(tmplUploadBtn);
      tmplRefCtrl.appendChild(tmplFileInput);

      // Status indicator
      if (state.uploadedTemplate) {
        var uploadInfo = el("span", "param-value template-upload-info");
        var uploadThumb = document.createElement("img");
        uploadThumb.src = "data:image/png;base64," + state.uploadedTemplate.data;
        uploadThumb.alt = state.uploadedTemplate.name;
        uploadInfo.appendChild(uploadThumb);
        uploadInfo.appendChild(document.createTextNode(state.uploadedTemplate.name));
        var clearBtn = el("button", "btn btn-small", "\u00d7");
        clearBtn.addEventListener("click", function () {
          state.uploadedTemplate = null;
          renderWorkflowParams();
        });
        uploadInfo.appendChild(clearBtn);
        tmplRefCtrl.appendChild(uploadInfo);
      } else if (state.referenceTimestamp !== null) {
        tmplRefCtrl.appendChild(el("span", "param-value", formatTimestamp(state.referenceTimestamp)));
      }
      tmplRefRow.appendChild(tmplRefCtrl);
      container.appendChild(tmplRefRow);
      addParamRow(container, "Threshold", rangeInput("paramTemplateThresh", 0.50, 1.00, 0.70, 0.01), "paramTemplateThreshVal");
      renderIntervalSlot("paramTemplateInterval", 0.5, 60, 1.0, 0.5);
    } else if (type === "flow") {
      addParamRow(container, "Magnitude", rangeInput("paramFlowMag", 0.5, 20.0, 2.0, 0.5), "paramFlowMagVal");
      renderIntervalSlot("paramFlowInterval", 0.5, 60, 1.0, 0.5);
    } else if (type === "scene") {
      var sceneList = el("div", "scene-reference-list");
      sceneList.id = "sceneRefList";
      state.sceneReferences.forEach(function (ref, i) {
        var item = el("div", "scene-ref-item");
        item.appendChild(el("span", "scene-ref-name", ref.name));
        item.appendChild(el("span", "param-value", formatTimestamp(ref.timestamp)));
        var rmBtn = el("button", "btn btn-small", "\u00d7");
        rmBtn.addEventListener("click", function () {
          state.sceneReferences.splice(i, 1);
          renderWorkflowParams();
        });
        item.appendChild(rmBtn);
        sceneList.appendChild(item);
      });
      container.appendChild(sceneList);
      var addScRow = el("div", "param-row");
      addScRow.appendChild(el("span", "param-label", "Add Scene"));
      var addScCtrl = el("div", "param-control");
      var scNameInp = textInput("paramSceneName", "e.g. menu, gameplay");
      addScCtrl.appendChild(scNameInp);
      var scCapBtn = el("button", "btn btn-small", "Capture");
      scCapBtn.addEventListener("click", function () {
        var nameEl = qs("#paramSceneName");
        var name = nameEl ? nameEl.value.trim() : "";
        if (!name) { showToast("Enter a scene name"); return; }
        state.sceneReferences.push({ name: name, timestamp: state.currentTimestamp });
        renderWorkflowParams();
        showToast("Scene '" + name + "' at " + formatTimestamp(state.currentTimestamp));
      });
      addScCtrl.appendChild(scCapBtn);
      addScRow.appendChild(addScCtrl);
      container.appendChild(addScRow);
      addParamRow(container, "Threshold", rangeInput("paramSceneThresh", 0.50, 1.00, 0.75, 0.01), "paramSceneThreshVal");
      renderIntervalSlot("paramSceneInterval", 0.5, 60, 1.0, 0.5);
    }

    if (type !== "timelapse") {
      addParamRow(container, "Event label", textInput("paramEventLabel", "e.g. low_health"));
      var dfCb = document.createElement("input");
      dfCb.type = "checkbox";
      dfCb.id = "paramDetectFirst";
      addParamRow(container, "Detect first", dfCb);
    }

    updateRunButton();
  }

  function addParamRow(container, label, control, valueDisplayId) {
    var row = el("div", "param-row");
    row.appendChild(el("span", "param-label", label));
    var ctrl = el("div", "param-control");
    ctrl.appendChild(control);
    if (valueDisplayId) {
      var valSpan = el("span", "param-value");
      valSpan.id = valueDisplayId;
      valSpan.textContent = control.value;
      ctrl.appendChild(valSpan);
      control.addEventListener("input", function () {
        valSpan.textContent = control.value;
        if (state.activeWorkflow === "color" && control.id && control.id.startsWith("paramColor")) {
          updateColorPreview();
        }
      });
    }
    row.appendChild(ctrl);
    container.appendChild(row);
  }

  function rangeInput(id, min, max, value, step) {
    var inp = document.createElement("input");
    inp.type = "range";
    inp.id = id;
    inp.min = min;
    inp.max = max;
    inp.value = value;
    if (step) inp.step = step;
    return inp;
  }

  function numberInput(id, min, max, value, step) {
    var inp = document.createElement("input");
    inp.type = "number";
    inp.id = id;
    inp.min = min;
    inp.max = max;
    inp.value = value;
    if (step) inp.step = step;
    return inp;
  }

  function textInput(id, placeholder) {
    var inp = document.createElement("input");
    inp.type = "text";
    inp.id = id;
    inp.placeholder = placeholder || "";
    return inp;
  }

  // ---- Color conversion utilities ----

  function rgbToHsv(r, g, b) {
    var rn = r / 255, gn = g / 255, bn = b / 255;
    var max = Math.max(rn, gn, bn), min = Math.min(rn, gn, bn);
    var d = max - min, hue = 0;
    if (d > 0) {
      if (max === rn) hue = ((gn - bn) / d) % 6;
      else if (max === gn) hue = (bn - rn) / d + 2;
      else hue = (rn - gn) / d + 4;
      hue = Math.round(hue * 30);
      if (hue < 0) hue += 180;
    }
    return { h: hue, s: max > 0 ? Math.round((d / max) * 255) : 0, v: Math.round(max * 255) };
  }

  function hsvToRgb(h, s, v) {
    var hDeg = h * 2, sn = s / 255, vn = v / 255;
    var c = vn * sn, x = c * (1 - Math.abs((hDeg / 60) % 2 - 1)), m = vn - c;
    var r1 = 0, g1 = 0, b1 = 0;
    if (hDeg < 60) { r1 = c; g1 = x; }
    else if (hDeg < 120) { r1 = x; g1 = c; }
    else if (hDeg < 180) { g1 = c; b1 = x; }
    else if (hDeg < 240) { g1 = x; b1 = c; }
    else if (hDeg < 300) { r1 = x; b1 = c; }
    else { r1 = c; b1 = x; }
    return { r: Math.round((r1 + m) * 255), g: Math.round((g1 + m) * 255), b: Math.round((b1 + m) * 255) };
  }

  function rgbToHex(r, g, b) {
    return "#" + ((1 << 24) | (r << 16) | (g << 8) | b).toString(16).slice(1);
  }

  function hexToRgb(hex) {
    hex = hex.replace(/^#/, "");
    if (hex.length === 3) hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
    if (!/^[0-9a-fA-F]{6}$/.test(hex)) return null;
    var n = parseInt(hex, 16);
    return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
  }

  function updateColorPreview() {
    var preview = qs("#colorPreview");
    if (!preview) return;
    var h = parseFloat((qs("#paramColorH") || {}).value) || 0;
    var s = parseFloat((qs("#paramColorS") || {}).value) || 0;
    var v = parseFloat((qs("#paramColorV") || {}).value) || 0;
    var rgb = hsvToRgb(h, s, v);
    preview.style.background = rgbToHex(rgb.r, rgb.g, rgb.b);
  }

  function setTargetColor(h, s, v) {
    h = clamp(Math.round(h), 0, 180);
    s = clamp(Math.round(s), 0, 255);
    v = clamp(Math.round(v), 0, 255);
    var hEl = qs("#paramColorH"), sEl = qs("#paramColorS"), vEl = qs("#paramColorV");
    if (hEl) hEl.value = h;
    if (sEl) sEl.value = s;
    if (vEl) vEl.value = v;
    var rgb = hsvToRgb(h, s, v);
    var hexEl = qs("#paramColorHex");
    if (hexEl) hexEl.value = rgbToHex(rgb.r, rgb.g, rgb.b);
    updateColorPreview();
    renderColorPalette();
    renderBrightnessStrip();
  }

  function sizeCanvasToDisplay(canvas) {
    var rect = canvas.getBoundingClientRect();
    var dpr = window.devicePixelRatio || 1;
    var w = Math.round(rect.width * dpr);
    var h = Math.round(rect.height * dpr);
    if (canvas.width !== w || canvas.height !== h) {
      canvas.width = w;
      canvas.height = h;
    }
    return { w: w, h: h, dpr: dpr };
  }

  function renderColorPalette() {
    var canvas = qs("#colorPalette");
    if (!canvas || !canvas.getBoundingClientRect().width) return;
    var size = sizeCanvasToDisplay(canvas);
    var w = size.w, h = size.h, dpr = size.dpr;
    var ctx = canvas.getContext("2d");

    // Hue spectrum (horizontal)
    var hueGrad = ctx.createLinearGradient(0, 0, w, 0);
    var stops = ["#ff0000", "#ffff00", "#00ff00", "#00ffff", "#0000ff", "#ff00ff", "#ff0000"];
    for (var i = 0; i < stops.length; i++) hueGrad.addColorStop(i / (stops.length - 1), stops[i]);
    ctx.fillStyle = hueGrad;
    ctx.fillRect(0, 0, w, h);

    // White-to-transparent overlay (bottom = white = low saturation)
    var satGrad = ctx.createLinearGradient(0, 0, 0, h);
    satGrad.addColorStop(0, "rgba(255,255,255,0)");
    satGrad.addColorStop(1, "rgba(255,255,255,1)");
    ctx.fillStyle = satGrad;
    ctx.fillRect(0, 0, w, h);

    // Black overlay for brightness
    var curV = parseFloat((qs("#paramColorV") || {}).value) || 0;
    var darkness = 1 - curV / 255;
    if (darkness > 0) {
      ctx.fillStyle = "rgba(0,0,0," + darkness + ")";
      ctx.fillRect(0, 0, w, h);
    }

    // Current position
    var curH = parseFloat((qs("#paramColorH") || {}).value) || 0;
    var curS = parseFloat((qs("#paramColorS") || {}).value) || 0;
    var cx = (curH / 180) * w;
    var cy = (1 - curS / 255) * h;

    // Tolerance range visualization
    var tol = parseFloat((qs("#paramColorTol") || {}).value) || 0;
    if (tol > 0) {
      var tolH = tol * 90 / 100;
      var tolS = tol * 128 / 100;
      var rx = (tolH / 180) * w;
      var ry = (tolS / 255) * h;
      ctx.fillStyle = "rgba(255,255,255,0.18)";
      ctx.strokeStyle = "rgba(255,255,255,0.4)";
      ctx.lineWidth = 1 * dpr;
      ctx.beginPath();
      ctx.rect(cx - rx, cy - ry, rx * 2, ry * 2);
      ctx.fill();
      ctx.stroke();
    }

    // Crosshair indicator
    var r = 5 * dpr;
    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, Math.PI * 2);
    ctx.strokeStyle = "#fff";
    ctx.lineWidth = 2 * dpr;
    ctx.stroke();
    ctx.beginPath();
    ctx.arc(cx, cy, r + 1 * dpr, 0, Math.PI * 2);
    ctx.strokeStyle = "#000";
    ctx.lineWidth = 1 * dpr;
    ctx.stroke();
  }

  function renderBrightnessStrip() {
    var canvas = qs("#colorBrightness");
    if (!canvas || !canvas.getBoundingClientRect().width) return;
    var size = sizeCanvasToDisplay(canvas);
    var w = size.w, h = size.h, dpr = size.dpr;
    var ctx = canvas.getContext("2d");
    var curH = parseFloat((qs("#paramColorH") || {}).value) || 0;
    var curS = parseFloat((qs("#paramColorS") || {}).value) || 0;

    // Gradient from black (left) to fully saturated color (right)
    var fullRgb = hsvToRgb(curH, curS, 255);
    var grad = ctx.createLinearGradient(0, 0, w, 0);
    grad.addColorStop(0, "#000000");
    grad.addColorStop(1, rgbToHex(fullRgb.r, fullRgb.g, fullRgb.b));
    ctx.fillStyle = grad;
    ctx.fillRect(0, 0, w, h);

    // Position indicator
    var curV = parseFloat((qs("#paramColorV") || {}).value) || 0;
    var ix = (curV / 255) * w;
    var r = 4 * dpr;
    ctx.beginPath();
    ctx.arc(clamp(ix, r, w - r), h / 2, r, 0, Math.PI * 2);
    ctx.fillStyle = "#fff";
    ctx.fill();
    ctx.strokeStyle = "#000";
    ctx.lineWidth = 1 * dpr;
    ctx.stroke();
  }

  function sampleColorFromRegion() {
    if (!state.frameImage || !state.activeRegion) {
      showToast("Select a saved region first");
      return;
    }
    var r = regionToPixels(state.regions[state.activeRegion]);
    var ctx = qs("#frameCanvas").getContext("2d");
    var imgData = ctx.getImageData(r.x, r.y, r.w, r.h);
    var data = imgData.data;
    var totalR = 0, totalG = 0, totalB = 0;
    var count = data.length / 4;
    for (var i = 0; i < data.length; i += 4) {
      totalR += data[i];
      totalG += data[i + 1];
      totalB += data[i + 2];
    }
    var hsv = rgbToHsv(Math.round(totalR / count), Math.round(totalG / count), Math.round(totalB / count));
    setTargetColor(hsv.h, hsv.s, hsv.v);
    showToast("Sampled color from " + state.activeRegion);
  }

  function activatePipette() {
    if (!state.frameImage) {
      showToast("Load a video frame first");
      return;
    }
    state.pipetteActive = true;
    var overlay = qs("#overlayCanvas");
    if (overlay) overlay.classList.add("pipette-active");
    var btn = qs("#pipetteBtn");
    if (btn) btn.classList.add("active");
  }

  function deactivatePipette() {
    state.pipetteActive = false;
    var overlay = qs("#overlayCanvas");
    if (overlay) overlay.classList.remove("pipette-active");
    var btn = qs("#pipetteBtn");
    if (btn) btn.classList.remove("active");
  }

  function updateRunButton() {
    var btn = qs("#runBtn");
    var hasRegion = state.runRegions.length > 0 || !!state.activeRegion;
    var hasParticipants = state.runParticipants.length > 0 || !!state.selectedParticipant;
    btn.disabled = !hasRegion || !hasParticipants;
    if (!hasRegion) {
      btn.setAttribute("data-tooltip", "Select a region first");
    } else if (!hasParticipants) {
      btn.setAttribute("data-tooltip", "Select participants to run");
    } else {
      btn.removeAttribute("data-tooltip");
    }
  }

  // ---- Run analysis ----

  function initRunButton() {
    qs("#runBtn").addEventListener("click", function () {
      var regions = state.runRegions.length > 0
        ? state.runRegions
        : (state.activeRegion ? [state.activeRegion] : []);
      if (regions.length === 0) return;
      var participants = state.runParticipants.length > 0
        ? state.runParticipants
        : (state.selectedParticipant ? [state.selectedParticipant] : []);
      if (participants.length === 0) return;

      var type = state.activeWorkflow;
      var params = gatherWorkflowParams(type);
      if (params === null) return;

      if (state.inMarker !== null) params.start_seconds = state.inMarker;
      if (state.outMarker !== null) params.end_seconds = state.outMarker;

      var chain = Promise.resolve();
      participants.forEach(function (pid) {
        regions.forEach(function (regionName) {
          chain = chain.then(function () {
            var body = {
              type: type,
              participant: pid,
              region: regionName,
              parameters: params,
            };
            return apiPost("api/tasks", body).then(function (data) {
              if (data.ok) {
                state.tasks.push(data.task);
                renderTaskList();
              } else {
                showToast(data.error || "Failed to create task for " + pid + " / " + regionName);
              }
            });
          });
        });
      });
      var totalTasks = participants.length * regions.length;
      chain.then(function () {
        showToast(totalTasks + " task" + (totalTasks !== 1 ? "s" : "") + " queued: " + type);
        startPolling();
      }).catch(function (err) { showToast("Error: " + err.message); });
    });
  }

  function gatherWorkflowParams(type) {
    var params = {};
    if (type === "color") {
      params.target_color = {
        h: parseFloat((qs("#paramColorH") || {}).value) || 0,
        s: parseFloat((qs("#paramColorS") || {}).value) || 0,
        v: parseFloat((qs("#paramColorV") || {}).value) || 0,
      };
      var tol = parseFloat((qs("#paramColorTol") || {}).value) || 30;
      params.tolerance = {
        h: Math.round(tol * 90 / 100),
        s: Math.round(tol * 128 / 100),
        v: Math.round(tol * 128 / 100),
      };
      params.interval = parseFloat((qs("#paramColorInterval") || {}).value) || 1.0;
    } else if (type === "change") {
      params.threshold = parseFloat((qs("#paramChangeThresh") || {}).value) || 0.03;
      params.noise_threshold = parseInt((qs("#paramChangeNoise") || {}).value) || 30;
      params.interval = parseFloat((qs("#paramChangeInterval") || {}).value) || 1.0;
    } else if (type === "similarity") {
      if (state.referenceTimestamp === null) {
        showToast("Capture a reference frame first");
        return null;
      }
      params.reference_timestamp = state.referenceTimestamp;
      params.threshold = parseFloat((qs("#paramSimThresh") || {}).value) || 0.90;
      params.interval = parseFloat((qs("#paramSimInterval") || {}).value) || 1.0;
    } else if (type === "text") {
      params.search_string = (qs("#paramTextSearch") || {}).value || "";
      if (!params.search_string.trim()) {
        showToast("Enter a search string");
        return null;
      }
      params.fuzzy_threshold = parseFloat((qs("#paramTextFuzzy") || {}).value) || 0.80;
      params.interval = parseFloat((qs("#paramTextInterval") || {}).value) || 2.0;
      var lang = (qs("#paramTextLang") || {}).value || "en";
      params.languages = [lang];
    } else if (type === "numbers") {
      var op = (qs("#paramNumOperator") || {}).value || "gt";
      params.operator = op;
      if (op === "range") {
        params.range_min = parseFloat((qs("#paramNumMin") || {}).value);
        params.range_max = parseFloat((qs("#paramNumMax") || {}).value);
        if (isNaN(params.range_min) || isNaN(params.range_max)) {
          showToast("Enter valid min and max values");
          return null;
        }
        if (params.range_min > params.range_max) {
          showToast("Min must be less than or equal to max");
          return null;
        }
      } else {
        params.target_value = parseFloat((qs("#paramNumTarget") || {}).value);
        if (isNaN(params.target_value)) {
          showToast("Enter a valid target number");
          return null;
        }
      }
      params.interval = parseFloat((qs("#paramNumInterval") || {}).value) || 2.0;
    } else if (type === "timelapse") {
      params.speedup_factor = parseFloat((qs("#paramTlSpeed") || {}).value) || 10;
      params.output_format = (qs("#paramTlFormat") || {}).value || "mp4";
    } else if (type === "template") {
      if (state.uploadedTemplate) {
        params.template_image_data = state.uploadedTemplate.data;
      } else if (state.referenceTimestamp !== null) {
        params.reference_timestamp = state.referenceTimestamp;
      } else {
        showToast("Capture a template region or upload a PNG");
        return null;
      }
      params.threshold = parseFloat((qs("#paramTemplateThresh") || {}).value) || 0.70;
      params.interval = parseFloat((qs("#paramTemplateInterval") || {}).value) || 1.0;
    } else if (type === "flow") {
      params.magnitude_threshold = parseFloat((qs("#paramFlowMag") || {}).value) || 2.0;
      params.interval = parseFloat((qs("#paramFlowInterval") || {}).value) || 1.0;
    } else if (type === "scene") {
      if (state.sceneReferences.length === 0) {
        showToast("Add at least one scene reference");
        return null;
      }
      params.scene_references = state.sceneReferences.map(function (ref) {
        return { name: ref.name, timestamp: ref.timestamp };
      });
      params.threshold = parseFloat((qs("#paramSceneThresh") || {}).value) || 0.75;
      params.interval = parseFloat((qs("#paramSceneInterval") || {}).value) || 1.0;
    }
    var labelEl = qs("#paramEventLabel");
    if (labelEl && labelEl.value.trim()) {
      params.event_label = labelEl.value.trim();
    }
    var dfEl = qs("#paramDetectFirst");
    if (dfEl && dfEl.checked) {
      params.detect_first = true;
    }
    return params;
  }

  // ---- Task queue ----
  // SVG icons use createElementNS() to build inline SVG from Heroicons paths (assets/icons/).
  // Pattern: create <svg> with createElementNS, set viewBox/width/height, then append <path>
  // elements with the d attribute copied from the relevant .svg file in assets/icons/.

  function svgEditIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "14");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    // pencil-square.svg from assets/icons
    var p1 = document.createElementNS("http://www.w3.org/2000/svg", "path");
    p1.setAttribute("d", "M13.4875 2.51256C12.804 1.82915 11.696 1.82915 11.0126 2.51256L6.75098 6.77417C6.49563 7.02951 6.29308 7.33265 6.15488 7.66628L5.30712 9.71282C5.19103 9.99307 5.25519 10.3157 5.46968 10.5302C5.68417 10.7447 6.00676 10.8088 6.28702 10.6928L8.33382 9.84501C8.66748 9.70681 8.97066 9.50423 9.22604 9.24886L13.4875 4.98744C14.1709 4.30402 14.1709 3.19598 13.4875 2.51256Z");
    svg.appendChild(p1);
    var p2 = document.createElementNS("http://www.w3.org/2000/svg", "path");
    p2.setAttribute("d", "M4.75 3.5C4.05964 3.5 3.5 4.05964 3.5 4.75V11.25C3.5 11.9404 4.05964 12.5 4.75 12.5H11.25C11.9404 12.5 12.5 11.9404 12.5 11.25V9C12.5 8.58579 12.8358 8.25 13.25 8.25C13.6642 8.25 14 8.58579 14 9V11.25C14 12.7688 12.7688 14 11.25 14H4.75C3.23122 14 2 12.7688 2 11.25V4.75C2 3.23122 3.23122 2 4.75 2H7C7.41421 2 7.75 2.33579 7.75 2.75C7.75 3.16421 7.41421 3.5 7 3.5H4.75Z");
    svg.appendChild(p2);
    return svg;
  }

  function svgEyeIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "14");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    // eye.svg from assets/icons
    var p1 = document.createElementNS("http://www.w3.org/2000/svg", "path");
    p1.setAttribute("d", "M8 9.5C8.82843 9.5 9.5 8.82843 9.5 8C9.5 7.17157 8.82843 6.5 8 6.5C7.17157 6.5 6.5 7.17157 6.5 8C6.5 8.82843 7.17157 9.5 8 9.5Z");
    svg.appendChild(p1);
    var p2 = document.createElementNS("http://www.w3.org/2000/svg", "path");
    p2.setAttribute("fill-rule", "evenodd");
    p2.setAttribute("clip-rule", "evenodd");
    p2.setAttribute("d", "M1.3794 8.28049C1.31616 8.09687 1.31625 7.89727 1.37965 7.71371C2.32719 4.97038 4.93238 3 7.99777 3C11.0653 3 13.672 4.97316 14.6179 7.71951C14.6811 7.90313 14.681 8.10274 14.6176 8.2863C13.6701 11.0296 11.0649 13 7.99952 13C4.93197 13 2.32527 11.0268 1.3794 8.28049ZM11 8C11 9.65685 9.65685 11 8 11C6.34315 11 5 9.65685 5 8C5 6.34315 6.34315 5 8 5C9.65685 5 11 6.34315 11 8Z");
    svg.appendChild(p2);
    return svg;
  }

  function svgEyeSlashIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "14");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    // eye-slash.svg from assets/icons
    var p1 = document.createElementNS("http://www.w3.org/2000/svg", "path");
    p1.setAttribute("fill-rule", "evenodd");
    p1.setAttribute("clip-rule", "evenodd");
    p1.setAttribute("d", "M3.28033 2.21967C2.98744 1.92678 2.51256 1.92678 2.21967 2.21967C1.92678 2.51256 1.92678 2.98744 2.21967 3.28033L12.7197 13.7803C13.0126 14.0732 13.4874 14.0732 13.7803 13.7803C14.0732 13.4874 14.0732 13.0126 13.7803 12.7197L12.4577 11.397C13.438 10.5863 14.1937 9.51366 14.6176 8.2863C14.681 8.10274 14.6811 7.90313 14.6179 7.71951C13.672 4.97316 11.0653 3 7.99777 3C6.85414 3 5.77457 3.27425 4.82123 3.76057L3.28033 2.21967ZM6.47602 5.41536L7.61147 6.55081C7.73539 6.51767 7.86563 6.5 8 6.5C8.82843 6.5 9.5 7.17157 9.5 8C9.5 8.13437 9.48233 8.26461 9.44919 8.38853L10.5846 9.52398C10.8486 9.07734 11 8.55636 11 8C11 6.34315 9.65685 5 8 5C7.44364 5 6.92266 5.15145 6.47602 5.41536Z");
    svg.appendChild(p1);
    var p2 = document.createElementNS("http://www.w3.org/2000/svg", "path");
    p2.setAttribute("d", "M7.81206 10.9942L9.62754 12.8097C9.10513 12.9341 8.56002 13 7.99952 13C4.93197 13 2.32527 11.0268 1.3794 8.28049C1.31616 8.09687 1.31625 7.89727 1.37965 7.71371C1.63675 6.96935 2.01588 6.28191 2.49314 5.67529L5.00579 8.18794C5.09895 9.69509 6.30491 10.901 7.81206 10.9942Z");
    svg.appendChild(p2);
    return svg;
  }

  function svgTagIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "14");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    // tag.svg from assets/icons
    var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("fill-rule", "evenodd");
    path.setAttribute("clip-rule", "evenodd");
    path.setAttribute("d", "M4.5 2C3.11929 2 2 3.11929 2 4.5V7.37868C2 8.04172 2.26339 8.67761 2.73223 9.14645L7.23223 13.6464C8.20854 14.6228 9.79145 14.6228 10.7678 13.6464L13.6464 10.7678C14.6228 9.79146 14.6228 8.20855 13.6464 7.23223L9.14645 2.73223C8.67761 2.26339 8.04172 2 7.37868 2H4.5ZM5 6C5.55228 6 6 5.55228 6 5C6 4.44772 5.55228 4 5 4C4.44772 4 4 4.44772 4 5C4 5.55228 4.44772 6 5 6Z");
    svg.appendChild(path);
    return svg;
  }

  function svgDragHandle() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "10");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    // bars-2.svg from assets/icons
    var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("fill-rule", "evenodd");
    path.setAttribute("clip-rule", "evenodd");
    path.setAttribute("d", "M2 4.75C2 4.33579 2.33579 4 2.75 4H13.25C13.6642 4 14 4.33579 14 4.75C14 5.16421 13.6642 5.5 13.25 5.5H2.75C2.33579 5.5 2 5.16421 2 4.75ZM2 11.25C2 10.8358 2.33579 10.5 2.75 10.5H13.25C13.6642 10.5 14 10.8358 14 11.25C14 11.6642 13.6642 12 13.25 12H2.75C2.33579 12 2 11.6642 2 11.25Z");
    svg.appendChild(path);
    return svg;
  }

  function svgDismissIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "14");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    // x-mark.svg from assets/icons
    var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", "M5.28033 4.21967C4.98744 3.92678 4.51256 3.92678 4.21967 4.21967C3.92678 4.51256 3.92678 4.98744 4.21967 5.28033L6.93934 8L4.21967 10.7197C3.92678 11.0126 3.92678 11.4874 4.21967 11.7803C4.51256 12.0732 4.98744 12.0732 5.28033 11.7803L8 9.06066L10.7197 11.7803C11.0126 12.0732 11.4874 12.0732 11.7803 11.7803C12.0732 11.4874 12.0732 11.0126 11.7803 10.7197L9.06066 8L11.7803 5.28033C12.0732 4.98744 12.0732 4.51256 11.7803 4.21967C11.4874 3.92678 11.0126 3.92678 10.7197 4.21967L8 6.93934L5.28033 4.21967Z");
    svg.appendChild(path);
    return svg;
  }

  function svgPauseIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "14");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    var p1 = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    p1.setAttribute("x", "3.5");
    p1.setAttribute("y", "2");
    p1.setAttribute("width", "3");
    p1.setAttribute("height", "12");
    p1.setAttribute("rx", "1");
    svg.appendChild(p1);
    var p2 = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    p2.setAttribute("x", "9.5");
    p2.setAttribute("y", "2");
    p2.setAttribute("width", "3");
    p2.setAttribute("height", "12");
    p2.setAttribute("rx", "1");
    svg.appendChild(p2);
    return svg;
  }

  var TASK_TYPE_ICON_FILES = {
    color: "eye-dropper",
    change: "bolt",
    similarity: "photo",
    text: "language",
    numbers: "hashtag",
    timelapse: "forward",
  };

  function svgPlayIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "14");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", "M3 3.732a1 1 0 0 1 1.514-.857l8.07 4.268a1 1 0 0 1 0 1.714l-8.07 4.268A1 1 0 0 1 3 12.268V3.732Z");
    svg.appendChild(path);
    return svg;
  }

  function svgCheckIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "14");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    // check.svg from assets/icons
    var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("fill-rule", "evenodd");
    path.setAttribute("clip-rule", "evenodd");
    path.setAttribute("d", "M12.416 3.37592C12.7607 3.60568 12.8538 4.07134 12.624 4.41598L7.62404 11.916C7.4994 12.1029 7.2975 12.2242 7.0739 12.2463C6.8503 12.2684 6.62855 12.1892 6.46967 12.0303L3.46967 9.03029C3.17678 8.73739 3.17678 8.26252 3.46967 7.96963C3.76256 7.67673 4.23744 7.67673 4.53033 7.96963L6.88343 10.3227L11.376 3.58393C11.6057 3.23929 12.0714 3.14616 12.416 3.37592Z");
    svg.appendChild(path);
    return svg;
  }

  function svgArchiveIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "14");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    // archive-box-arrow-down.svg from assets/icons
    var p1 = document.createElementNS("http://www.w3.org/2000/svg", "path");
    p1.setAttribute("d", "M2 3C2 2.44772 2.44772 2 3 2H13C13.5523 2 14 2.44772 14 3V4C14 4.55228 13.5523 5 13 5H3C2.44772 5 2 4.55228 2 4V3Z");
    svg.appendChild(p1);
    var p2 = document.createElementNS("http://www.w3.org/2000/svg", "path");
    p2.setAttribute("fill-rule", "evenodd");
    p2.setAttribute("clip-rule", "evenodd");
    p2.setAttribute("d", "M13 6H3V12C3 13.1046 3.89543 14 5 14H11C12.1046 14 13 13.1046 13 12V6ZM8.75 7.75C8.75 7.33579 8.41421 7 8 7C7.58579 7 7.25 7.33579 7.25 7.75V10.4393L6.03033 9.21967C5.73744 8.92678 5.26256 8.92678 4.96967 9.21967C4.67678 9.51256 4.67678 9.98744 4.96967 10.2803L7.46967 12.7803C7.76256 13.0732 8.23744 13.0732 8.53033 12.7803L11.0303 10.2803C11.3232 9.98744 11.3232 9.51256 11.0303 9.21967C10.7374 8.92678 10.2626 8.92678 9.96967 9.21967L8.75 10.4393V7.75Z");
    svg.appendChild(p2);
    return svg;
  }

  function svgRestoreIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "14");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    // arrow-up-tray.svg from assets/icons
    var p1 = document.createElementNS("http://www.w3.org/2000/svg", "path");
    p1.setAttribute("d", "M7.25 10.25C7.25 10.6642 7.58579 11 8 11C8.41421 11 8.75 10.6642 8.75 10.25L8.75 4.56066L10.9697 6.78033C11.2626 7.07322 11.7374 7.07322 12.0303 6.78033C12.3232 6.48744 12.3232 6.01256 12.0303 5.71967L8.53033 2.21967C8.23744 1.92678 7.76256 1.92678 7.46967 2.21967L3.96967 5.71967C3.67678 6.01256 3.67678 6.48744 3.96967 6.78033C4.26256 7.07322 4.73744 7.07322 5.03033 6.78033L7.25 4.56066L7.25 10.25Z");
    svg.appendChild(p1);
    var p2 = document.createElementNS("http://www.w3.org/2000/svg", "path");
    p2.setAttribute("d", "M3.5 9.75C3.5 9.33579 3.16421 9 2.75 9C2.33579 9 2 9.33579 2 9.75V11.25C2 12.7688 3.23122 14 4.75 14H11.25C12.7688 14 14 12.7688 14 11.25V9.75C14 9.33579 13.6642 9 13.25 9C12.8358 9 12.5 9.33579 12.5 9.75V11.25C12.5 11.9404 11.9404 12.5 11.25 12.5H4.75C4.05964 12.5 3.5 11.9404 3.5 11.25V9.75Z");
    svg.appendChild(p2);
    return svg;
  }

  function svgPlusIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "12");
    svg.setAttribute("height", "12");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", "M8.75 3.75C8.75 3.33579 8.41421 3 8 3C7.58579 3 7.25 3.33579 7.25 3.75V7.25H3.75C3.33579 7.25 3 7.58579 3 8C3 8.41421 3.33579 8.75 3.75 8.75L7.25 8.75V12.25C7.25 12.6642 7.58579 13 8 13C8.41421 13 8.75 12.6642 8.75 12.25V8.75L12.25 8.75C12.6642 8.75 13 8.41421 13 8C13 7.58579 12.6642 7.25 12.25 7.25H8.75V3.75Z");
    svg.appendChild(path);
    return svg;
  }

  function svgMinusIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "12");
    svg.setAttribute("height", "12");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", "M3.75 7.25C3.33579 7.25 3 7.58579 3 8C3 8.41421 3.33579 8.75 3.75 8.75L12.25 8.75C12.6642 8.75 13 8.41421 13 8C13 7.58579 12.6642 7.25 12.25 7.25H3.75Z");
    svg.appendChild(path);
    return svg;
  }

  function svgChevronLeftIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "12");
    svg.setAttribute("height", "12");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("fill-rule", "evenodd");
    path.setAttribute("clip-rule", "evenodd");
    path.setAttribute("d", "M9.78033 4.21967C10.0732 4.51256 10.0732 4.98744 9.78033 5.28033L7.06066 8L9.78033 10.7197C10.0732 11.0126 10.0732 11.4874 9.78033 11.7803C9.48744 12.0732 9.01256 12.0732 8.71967 11.7803L5.46967 8.53033C5.17678 8.23744 5.17678 7.76256 5.46967 7.46967L8.71967 4.21967C9.01256 3.92678 9.48744 3.92678 9.78033 4.21967Z");
    svg.appendChild(path);
    return svg;
  }

  function svgChevronRightIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "12");
    svg.setAttribute("height", "12");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("fill-rule", "evenodd");
    path.setAttribute("clip-rule", "evenodd");
    path.setAttribute("d", "M6.21967 4.21967C6.51256 3.92678 6.98744 3.92678 7.28033 4.21967L10.5303 7.46967C10.8232 7.76256 10.8232 8.23744 10.5303 8.53033L7.28033 11.7803C6.98744 12.0732 6.51256 12.0732 6.21967 11.7803C5.92678 11.4874 5.92678 11.0126 6.21967 10.7197L8.93934 8L6.21967 5.28033C5.92678 4.98744 5.92678 4.51256 6.21967 4.21967Z");
    svg.appendChild(path);
    return svg;
  }

  function svgChevronDownIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "10");
    svg.setAttribute("height", "10");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("fill-rule", "evenodd");
    path.setAttribute("clip-rule", "evenodd");
    path.setAttribute("d", "M4.21967 6.21967C4.51256 5.92678 4.98744 5.92678 5.28033 6.21967L8 8.93934L10.7197 6.21967C11.0126 5.92678 11.4874 5.92678 11.7803 6.21967C12.0732 6.51256 12.0732 6.98744 11.7803 7.28033L8.53033 10.5303C8.23744 10.8232 7.76256 10.8232 7.46967 10.5303L4.21967 7.28033C3.92678 6.98744 3.92678 6.51256 4.21967 6.21967Z");
    svg.appendChild(path);
    return svg;
  }

  function sortTasks() {
    // completed/failed at top (oldest first), then running, then queued (by priority), cancelled last
    var statusOrder = { completed: 0, failed: 1, running: 2, paused: 3, queued: 4, cancelled: 5 };
    state.tasks.sort(function (a, b) {
      var sa = statusOrder[a.status] !== undefined ? statusOrder[a.status] : 5;
      var sb = statusOrder[b.status] !== undefined ? statusOrder[b.status] : 5;
      if (sa !== sb) return sa - sb;
      if (a.status === "queued" && b.status === "queued") {
        return (a.priority || 100) - (b.priority || 100);
      }
      return (a.created_at || "").localeCompare(b.created_at || "");
    });
  }

  function initTaskQueue() {
    var taskListEl = qs("#taskList");

    // Click handler delegated on taskList
    taskListEl.addEventListener("click", function (e) {
      var card = e.target.closest(".task-card");
      if (!card) return;
      var taskId = card.dataset.taskId;

      // Dismiss button
      if (e.target.closest(".task-card-dismiss")) {
        apiDelete("api/tasks/" + taskId + "?dismiss=true")
          .then(function (data) {
            if (data.ok) {
              state.tasks = state.tasks.filter(function (t) { return t.id !== taskId; });
              if (state.hoveredTaskId === taskId) {
                state.hoveredTaskId = null;
              }
              if (state.selectedTaskId === taskId) {
                state.selectedTaskId = null;
                state.selectedTaskResults = null;
                renderResults();
              }
              renderTaskList();
              renderTimeline();
              showToast("Task dismissed");
            }
          })
          .catch(function () { showToast("Failed to dismiss task"); });
        return;
      }

      // Edit button
      if (e.target.closest(".task-card-edit")) {
        var task = findTask(taskId);
        if (task) restoreTaskToWorkflow(task);
        return;
      }

      // Select completed/paused/running task to view results; click again to deselect
      var task = findTask(taskId);
      if (task && (task.status === "completed" || task.status === "paused" || task.status === "running")) {
        if (state.selectedTaskId === taskId) {
          state.selectedTaskId = null;
          state.selectedTaskResults = null;
          renderResults();
          renderTaskList();
          renderTimeline();
        } else {
          if (task.participant && task.participant !== state.selectedParticipant) {
            selectParticipant(task.participant);
          }
          state.selectedTaskId = taskId;
          loadAndShowResults(taskId);
          renderTaskList();
        }
      }
    });

    // Hover handler for task focus (dim non-hovered timeline markers)
    taskListEl.addEventListener("mouseover", function (e) {
      var card = e.target.closest(".task-card");
      if (!card) return;
      var task = findTask(card.dataset.taskId);
      if (task && (task.status === "completed" || task.status === "running")) {
        if (state.hoveredTaskId !== task.id) {
          state.hoveredTaskId = task.id;
          renderTimeline();
        }
      } else if (state.hoveredTaskId) {
        state.hoveredTaskId = null;
        renderTimeline();
      }
    });

    taskListEl.addEventListener("mouseleave", function () {
      if (state.hoveredTaskId) {
        state.hoveredTaskId = null;
        renderTimeline();
      }
    });

    // Drag-and-drop: only initiate drag from the handle
    taskListEl.addEventListener("dragstart", function (e) {
      var card = e.target.closest(".task-card");
      if (!card) { e.preventDefault(); return; }
      var task = findTask(card.dataset.taskId);
      if (!task) { e.preventDefault(); return; }
      var allowed = task.status === "queued" || task.status === "completed" || task.status === "failed";
      if (!allowed) { e.preventDefault(); return; }
      card.classList.add("dragging");
      e.dataTransfer.setData("text/plain", card.dataset.taskId);
      e.dataTransfer.setData("application/x-task-status", task.status);
      e.dataTransfer.effectAllowed = "move";
    });

    taskListEl.addEventListener("dragend", function (e) {
      var card = e.target.closest(".task-card");
      if (card) {
        card.classList.remove("dragging");
        card.removeAttribute("draggable");
      }
      clearDragIndicators(taskListEl);
    });

    taskListEl.addEventListener("dragover", function (e) {
      if (e.dataTransfer.types.indexOf("text/plain") < 0) return;

      var cards = taskListEl.querySelectorAll(".task-card:not(.dragging)");
      var insertIdx = getDropIndex(taskListEl, e.clientY);

      // Determine boundary between finished (completed/failed) and queued zones
      var finishedCount = 0;
      var queuedStart = cards.length;
      for (var i = 0; i < cards.length; i++) {
        var t = findTask(cards[i].dataset.taskId);
        if (t && (t.status === "completed" || t.status === "failed")) finishedCount++;
        else { queuedStart = i; break; }
      }

      var dragStatus = e.dataTransfer.types.indexOf("application/x-task-status") >= 0
        ? "unknown" : "unknown";
      // Infer dragged status from position constraint:
      // We stored it in dataTransfer but can't read it during dragover (security).
      // Instead, find the dragging card to determine its status.
      var draggingCard = taskListEl.querySelector(".task-card.dragging");
      var draggingTask = draggingCard ? findTask(draggingCard.dataset.taskId) : null;
      var isQueuedDrag = draggingTask && draggingTask.status === "queued";
      var isFinishedDrag = draggingTask && (draggingTask.status === "completed" || draggingTask.status === "failed");

      // Queued tasks can't go above finished tasks
      if (isQueuedDrag && insertIdx < finishedCount) return;
      // Finished tasks can't go below into queued zone
      if (isFinishedDrag && insertIdx > finishedCount) return;

      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      clearDragIndicators(taskListEl);
      if (insertIdx < cards.length) {
        cards[insertIdx].classList.add("drag-over");
      }
    });

    taskListEl.addEventListener("dragleave", function (e) {
      var card = e.target.closest(".task-card");
      if (card) card.classList.remove("drag-over");
    });

    taskListEl.addEventListener("drop", function (e) {
      e.preventDefault();
      clearDragIndicators(taskListEl);
      var draggedId = e.dataTransfer.getData("text/plain");
      if (!draggedId) return;

      var draggedTask = findTask(draggedId);
      if (!draggedTask) return;
      var isQueued = draggedTask.status === "queued";

      if (isQueued) {
        // Reorder among queued tasks
        var queuedIds = [];
        state.tasks.forEach(function (t) {
          if (t.status === "queued") queuedIds.push(t.id);
        });
        var fromIdx = queuedIds.indexOf(draggedId);
        if (fromIdx < 0) return;
        queuedIds.splice(fromIdx, 1);
        var toIdx = getDropIndexAmongStatus(taskListEl, e.clientY, "queued");
        queuedIds.splice(toIdx, 0, draggedId);

        apiPut("api/tasks/reorder", { task_ids: queuedIds }).catch(function () {
          showToast("Failed to reorder tasks");
        });
        for (var i = 0; i < queuedIds.length; i++) {
          var t = findTask(queuedIds[i]);
          if (t) t.priority = i + 1;
        }
      } else {
        // Reorder finished tasks visually via created_at swapping
        var finishedTasks = [];
        state.tasks.forEach(function (t) {
          if (t.status === "completed" || t.status === "failed") finishedTasks.push(t);
        });
        var fromIdx2 = -1;
        for (var j = 0; j < finishedTasks.length; j++) {
          if (finishedTasks[j].id === draggedId) { fromIdx2 = j; break; }
        }
        if (fromIdx2 < 0) return;
        finishedTasks.splice(fromIdx2, 1);
        var toIdx2 = getDropIndexAmongStatus(taskListEl, e.clientY, "finished");
        finishedTasks.splice(toIdx2, 0, draggedTask);
        // Reassign created_at to maintain the visual order across polls
        var timestamps = finishedTasks.map(function (t) { return t.created_at; });
        timestamps.sort();
        for (var k = 0; k < finishedTasks.length; k++) {
          finishedTasks[k].created_at = timestamps[k];
        }
      }

      sortTasks();
      renderTaskList();
    });
  }

  function getDropIndex(container, clientY) {
    var cards = container.querySelectorAll(".task-card:not(.dragging)");
    for (var i = 0; i < cards.length; i++) {
      var rect = cards[i].getBoundingClientRect();
      if (clientY < rect.top + rect.height / 2) return i;
    }
    return cards.length;
  }

  function getDropIndexAmongStatus(container, clientY, group) {
    var cards = container.querySelectorAll(".task-card:not(.dragging)");
    var idx = 0;
    for (var i = 0; i < cards.length; i++) {
      var t = findTask(cards[i].dataset.taskId);
      if (!t) continue;
      var match = group === "queued"
        ? t.status === "queued"
        : (t.status === "completed" || t.status === "failed");
      if (!match) continue;
      var rect = cards[i].getBoundingClientRect();
      if (clientY < rect.top + rect.height / 2) return idx;
      idx++;
    }
    return idx;
  }

  function clearDragIndicators(container) {
    var cards = container.querySelectorAll(".task-card.drag-over");
    for (var i = 0; i < cards.length; i++) cards[i].classList.remove("drag-over");
  }

  function setInputValue(selector, value) {
    var inp = qs(selector);
    if (inp) inp.value = value;
  }

  function syncValueDisplays() {
    var inputs = qsa(".param-control input[type='range']");
    for (var i = 0; i < inputs.length; i++) {
      var valSpan = inputs[i].parentNode.querySelector(".param-value");
      if (valSpan) valSpan.textContent = inputs[i].value;
    }
  }

  function restoreTaskToWorkflow(task) {
    // Switch workflow tab
    state.activeWorkflow = task.type;
    qsa(".wf-tab").forEach(function (t) { t.classList.remove("active"); });
    var targetTab = qs('.wf-tab[data-type="' + task.type + '"]');
    if (targetTab) targetTab.classList.add("active");

    // Select participant
    if (task.participant) {
      state.selectedParticipant = task.participant;
      var sel = qs("#participantSelect");
      if (sel) sel.value = task.participant;
    }

    // Select region
    if (task.region && state.regions[task.region]) {
      state.activeRegion = task.region;
      state.pendingRegion = null;
      renderRegionChips();
      renderOverlay();
      updateRegionButtons();
    }

    // For similarity, restore reference timestamp before rendering params
    if (task.type === "similarity") {
      var params = task.parameters || {};
      if (params.reference_timestamp !== undefined) {
        state.referenceTimestamp = params.reference_timestamp;
      } else {
        showToast("Reference frame must be recaptured");
      }
    }

    // Rebuild param controls then set values
    renderWorkflowParams();

    var params = task.parameters || {};
    if (task.type === "color") {
      setInputValue("#paramColorH", params.target_color ? params.target_color.h : 90);
      setInputValue("#paramColorS", params.target_color ? params.target_color.s : 200);
      setInputValue("#paramColorV", params.target_color ? params.target_color.v : 200);
      var savedTol = params.tolerance ? Math.round(params.tolerance.h * 100 / 90) : 30;
      setInputValue("#paramColorTol", savedTol);
      setInputValue("#paramColorInterval", params.interval || 1.0);
      updateColorPreview();
    } else if (task.type === "change") {
      setInputValue("#paramChangeThresh", params.threshold || 0.03);
      setInputValue("#paramChangeNoise", params.noise_threshold || 30);
      setInputValue("#paramChangeInterval", params.interval || 1.0);
    } else if (task.type === "similarity") {
      setInputValue("#paramSimThresh", params.threshold || 0.90);
      setInputValue("#paramSimInterval", params.interval || 1.0);
    } else if (task.type === "text") {
      setInputValue("#paramTextSearch", params.search_string || "");
      setInputValue("#paramTextFuzzy", params.fuzzy_threshold || 0.80);
      setInputValue("#paramTextInterval", params.interval || 2.0);
      if (params.languages && params.languages[0]) {
        setInputValue("#paramTextLang", params.languages[0]);
      }
    } else if (task.type === "timelapse") {
      setInputValue("#paramTlSpeed", params.speedup_factor || 10);
      setInputValue("#paramTlFormat", params.output_format || "mp4");
    } else if (task.type === "template") {
      if (params.reference_timestamp !== undefined) {
        state.referenceTimestamp = params.reference_timestamp;
      }
      setInputValue("#paramTemplateThresh", params.threshold || 0.70);
      setInputValue("#paramTemplateInterval", params.interval || 1.0);
    } else if (task.type === "flow") {
      setInputValue("#paramFlowMag", params.magnitude_threshold || 2.0);
      setInputValue("#paramFlowInterval", params.interval || 1.0);
    } else if (task.type === "scene") {
      if (params.scene_references) {
        state.sceneReferences = params.scene_references.map(function (ref) {
          return { name: ref.name, timestamp: ref.timestamp };
        });
      }
      setInputValue("#paramSceneThresh", params.threshold || 0.75);
      setInputValue("#paramSceneInterval", params.interval || 1.0);
    }

    syncValueDisplays();
    updateRunButton();
    showToast("Restored " + task.type + " task parameters");
  }

  function findTask(id) {
    for (var i = 0; i < state.tasks.length; i++) {
      if (state.tasks[i].id === id) return state.tasks[i];
    }
    return null;
  }

  function focusedTaskId() {
    if (state.hoveredTaskId) {
      var ht = findTask(state.hoveredTaskId);
      if (ht && ht.status === "completed") return state.hoveredTaskId;
    }
    return state.selectedTaskId;
  }

  function updatePauseButton() {
    var btn = qs("#taskQueuePauseBtn");
    if (!btn) return;
    btn.innerHTML = "";
    if (state.queuePaused) {
      btn.appendChild(svgPlayIcon());
      btn.title = "Resume queue";
    } else {
      btn.appendChild(svgPauseIcon());
      btn.title = "Pause queue";
    }
  }

  function initPauseButton() {
    var btn = qs("#taskQueuePauseBtn");
    if (!btn) return;
    updatePauseButton();
    btn.addEventListener("click", function () {
      var endpoint = state.queuePaused ? "api/tasks/resume" : "api/tasks/pause";
      apiPost(endpoint)
        .then(function (data) {
          if (data.ok) {
            state.queuePaused = data.paused;
            updatePauseButton();
          }
        })
        .catch(function (err) { showToast("Error: " + err.message); });
    });
  }

  function initTaskFilters() {
    var doneBtn = qs("#taskFilterDoneBtn");
    var failedBtn = qs("#taskFilterFailedBtn");
    if (doneBtn) {
      doneBtn.appendChild(svgCheckIcon());
      doneBtn.addEventListener("click", function () { toggleTaskFilter("completed"); });
    }
    if (failedBtn) {
      failedBtn.appendChild(svgDismissIcon());
      failedBtn.addEventListener("click", function () { toggleTaskFilter("failed"); });
    }
  }

  function toggleTaskFilter(status) {
    state.taskFilter = state.taskFilter === status ? null : status;
    updateTaskFilterButtons();
    renderTaskList();
  }

  function updateTaskFilterButtons() {
    var doneBtn = qs("#taskFilterDoneBtn");
    var failedBtn = qs("#taskFilterFailedBtn");
    if (doneBtn) doneBtn.classList.toggle("active", state.taskFilter === "completed");
    if (failedBtn) failedBtn.classList.toggle("active", state.taskFilter === "failed");
  }

  function renderTaskList() {
    sortTasks();
    var container = qs("#taskList");
    var count = qs("#taskCount");
    var filtered = state.taskFilter
      ? state.tasks.filter(function (t) { return t.status === state.taskFilter; })
      : state.tasks;
    count.textContent = "(" + filtered.length + ")";
    updateTaskFilterButtons();
    if (state.tasks.length === 0) {
      container.innerHTML = '<div class="panel-empty">No tasks yet. Configure a workflow and click Run.</div>';
      return;
    }
    if (filtered.length === 0) {
      container.innerHTML = '<div class="panel-empty">No ' + state.taskFilter + ' tasks.</div>';
      return;
    }
    var frag = document.createDocumentFragment();
    filtered.forEach(function (task) {
      var card = el("div", "task-card task-card-" + task.status);
      card.dataset.taskId = task.id;
      if (task.id === state.selectedTaskId) card.classList.add("selected");

      // Drag handle for reorderable tasks (completed, failed, queued)
      var isDraggable = task.status === "queued" || task.status === "completed" || task.status === "failed";
      if (isDraggable) {
        var handle = el("span", "task-card-drag-handle");
        handle.appendChild(svgDragHandle());
        handle.addEventListener("mousedown", function () { card.setAttribute("draggable", "true"); });
        handle.addEventListener("mouseup", function () { card.removeAttribute("draggable"); });
        card.appendChild(handle);
      } else if (task.status === "running") {
        card.appendChild(el("span", "task-card-spinner"));
      } else if (task.status === "paused") {
        var pauseIcon = el("span", "task-card-pause-icon");
        pauseIcon.appendChild(svgPauseIcon());
        card.appendChild(pauseIcon);
      }

      // Type badge
      var badge = el("span", "task-card-type");
      badge.style.color = taskTypeColor(task.type);
      badge.title = task.type;
      var iconSpan = el("span", "task-card-type-icon");
      var iconFile = TASK_TYPE_ICON_FILES[task.type] || "squares-2x2";
      var iconUrl = 'url("/screenspace/icons/' + iconFile + '.svg")';
      iconSpan.style.maskImage = iconUrl;
      iconSpan.style.webkitMaskImage = iconUrl;
      badge.appendChild(iconSpan);
      card.appendChild(badge);

      // Info
      var info = el("div", "task-card-info");
      var meta = el("span", "task-card-meta");
      meta.textContent = task.participant + " \u00b7 " + (task.region || "");
      info.appendChild(meta);

      if (task.status === "running" || task.status === "paused") {
        var prog = el("div", "task-card-progress");
        var fill = el("div", "task-card-progress-fill");
        fill.style.width = Math.round((task.progress || 0) * 100) + "%";
        prog.appendChild(fill);
        info.appendChild(prog);
      }
      card.appendChild(info);

      // Status text
      var statusText = task.status;
      if (task.status === "running") {
        var rPct = Math.round((task.progress || 0) * 100);
        var rLen = Array.isArray(task.result) ? task.result.length : 0;
        statusText = rPct + "%" + (rLen ? " \u00b7 " + rLen + " result" + (rLen !== 1 ? "s" : "") : "");
      }
      if (task.status === "paused") {
        var pPct = Math.round((task.progress || 0) * 100);
        var pLen = Array.isArray(task.result) ? task.result.length : 0;
        statusText = "paused " + pPct + "%" + (pLen ? " \u00b7 " + pLen + " result" + (pLen !== 1 ? "s" : "") : "");
      }
      if (task.status === "failed" && task.error) {
        statusText = task.error;
        card.title = task.error;
      }
      if (task.status === "completed" && task.result) {
        var rLen = Array.isArray(task.result) ? task.result.length : (typeof task.result === "string" ? 1 : 0);
        statusText = rLen + " result" + (rLen !== 1 ? "s" : "");
      }
      card.appendChild(el("span", "task-card-status", statusText));

      // Edit button
      var editBtn = el("button", "task-card-edit");
      editBtn.title = "Edit";
      editBtn.appendChild(svgEditIcon());
      card.appendChild(editBtn);

      // Dismiss button
      var dismissBtn = el("button", "task-card-dismiss");
      dismissBtn.title = "Dismiss";
      dismissBtn.appendChild(svgDismissIcon());
      card.appendChild(dismissBtn);

      frag.appendChild(card);
    });
    container.innerHTML = "";
    container.appendChild(frag);
  }

  // ---- Polling ----

  function startPolling() {
    if (state.pollTimer) return;
    state.pollTimer = setInterval(pollTasks, POLL_INTERVAL);
  }

  function pollTasks() {
    var hasActive = state.tasks.some(function (t) {
      return t.status === "queued" || t.status === "running" || t.status === "paused";
    });
    if (!hasActive) {
      clearInterval(state.pollTimer);
      state.pollTimer = null;
      return;
    }

    apiGet("api/tasks")
      .then(function (data) {
        if (!data.ok) return;
        var oldSelected = state.selectedTaskId;
        var oldTask = oldSelected ? findTask(oldSelected) : null;
        var wasRunning = oldTask && (oldTask.status === "queued" || oldTask.status === "running");
        state.tasks = data.tasks;
        if (data.paused !== undefined) {
          state.queuePaused = data.paused;
          updatePauseButton();
        }
        var fp = JSON.stringify(data.tasks.map(function (t) {
          return t.id + ":" + t.status + ":" + t.progress;
        }));
        var changed = fp !== _lastPollFingerprint;
        _lastPollFingerprint = fp;
        if (changed) {
          renderTaskList();
          renderTimeline();
        }
        // Auto-update results for selected running task
        if (oldSelected) {
          var selTask = findTask(oldSelected);
          if (selTask && selTask.status === "running" && selTask.result) {
            state.selectedTaskResults = selTask.result;
            renderResults();
          }
        }
        // Auto-load results when selected task completes
        if (wasRunning && oldSelected) {
          var newTask = findTask(oldSelected);
          if (newTask && newTask.status === "completed") {
            loadAndShowResults(oldSelected);
          }
        }
      })
      .catch(function () {});
  }

  // ---- Results ----

  function initResultsPanel() {
    qs("#resultsList").addEventListener("click", function (e) {
      // Handle exclude toggle
      var btn = e.target.closest(".result-exclude-btn");
      if (btn && btn.dataset.eventId) {
        var evId = btn.dataset.eventId;
        var isExcluded = btn.dataset.excluded === "true";
        var endpoint = isExcluded ? "api/events/" + evId + "/include" : "api/events/" + evId + "/exclude";
        apiPut(endpoint).then(function () {
          var evts = state.taskEvents[state.selectedTaskId] || [];
          for (var i = 0; i < evts.length; i++) {
            if (evts[i].id === evId) { evts[i].excluded = !isExcluded; break; }
          }
          renderResults();
        });
        return;
      }
      var row = e.target.closest(".result-row");
      if (!row || !row.dataset.timestamp) return;
      var ts = parseFloat(row.dataset.timestamp);
      if (isNaN(ts)) return;
      var task = state.selectedTaskId ? findTask(state.selectedTaskId) : null;
      if (task && task.participant && task.participant !== state.selectedParticipant) {
        selectParticipant(task.participant, ts);
        return;
      }
      // Set result overlay for spatial visualization
      var ri = parseInt(row.dataset.resultIndex, 10);
      var rData = (!isNaN(ri) && state.selectedTaskResults) ? state.selectedTaskResults[ri] : null;
      if (task && rData && task.type === "template" && rData.matches) {
        state.resultOverlay = { type: "template", data: rData };
      } else if (task && rData && task.type === "flow" && rData.flow_grid) {
        state.resultOverlay = { type: "flow", data: rData, region: regionToPixels(state.regions[task.region] || {}) };
      } else {
        state.resultOverlay = null;
      }
      loadFrame(ts);
    });

    qs("#exportJsonBtn").addEventListener("click", function () { exportResults("json"); });
    qs("#exportCsvBtn").addEventListener("click", function () { exportResults("csv"); });

    qs("#showExcludedCb").addEventListener("change", function () {
      state.showExcluded = this.checked;
      renderResults();
    });
  }

  function loadAndShowResults(taskId) {
    state.resultOverlay = null;
    state.heatmapOverlay = null;
    apiGet("api/tasks/" + taskId + "/results")
      .then(function (data) {
        state.selectedTaskId = taskId;
        state.selectedTaskResults = data.results;
        return apiGet("api/events?task_id=" + taskId);
      })
      .then(function (evData) {
        state.taskEvents[taskId] = evData.events || [];
        renderResults();
        renderTaskList();
      })
      .catch(function () { showToast("Failed to load results"); });
  }

  function renderResults() {
    var container = qs("#resultsList");
    var countEl = qs("#resultCount");
    var actionsEl = qs("#resultsActions");
    var results = state.selectedTaskResults;
    var task = state.selectedTaskId ? findTask(state.selectedTaskId) : null;

    if (!results || !task) {
      container.innerHTML = '<div class="panel-empty">Click a task to view results.</div>';
      countEl.textContent = "";
      actionsEl.classList.add("hidden");
      return;
    }

    actionsEl.classList.remove("hidden");

    // Timelapse: single file result
    if (task.type === "timelapse" && typeof results === "string") {
      countEl.textContent = "";
      container.innerHTML = "";
      var wrapper = el("div", "timelapse-result");
      var ext = results.split(".").pop().toLowerCase();
      var filename = results.split("/").pop();
      if (ext === "gif") {
        var img = document.createElement("img");
        img.src = "media/" + filename;
        wrapper.appendChild(img);
      } else {
        var vid = document.createElement("video");
        vid.src = "media/" + filename;
        vid.controls = true;
        vid.muted = true;
        wrapper.appendChild(vid);
      }
      container.innerHTML = "";
      container.appendChild(wrapper);
      return;
    }

    if (!Array.isArray(results)) {
      container.innerHTML = '<div class="panel-empty">No results.</div>';
      countEl.textContent = "";
      return;
    }

    var events = state.taskEvents[state.selectedTaskId] || [];
    var eventsByTs = {};
    events.forEach(function (ev) {
      var key = ev.time_in.toFixed(2);
      if (!eventsByTs[key]) eventsByTs[key] = [];
      eventsByTs[key].push(ev);
    });

    // For color results (spans), build a consumed-index tracker per timestamp
    var eventTsIndex = {};

    var showToggle = qs("#showExcludedToggle");
    if (showToggle) showToggle.classList.toggle("hidden", events.length === 0);

    var visibleCount = 0;

    countEl.textContent = "(" + results.length + ")";
    container.innerHTML = "";

    // Heatmap artifact display (template, flow)
    if (task.heatmap && (task.type === "template" || task.type === "flow")) {
      var heatmapSection = el("div", "heatmap-result");
      var heatmapLabel = el("div", "heatmap-label");
      heatmapLabel.appendChild(document.createTextNode("Detection Heatmap"));
      var overlayBtn = el("button", "ss-btn ss-btn-sm", state.heatmapOverlay ? "Hide Overlay" : "Overlay on Frame");
      overlayBtn.addEventListener("click", function () {
        if (state.heatmapOverlay) {
          state.heatmapOverlay = null;
          overlayBtn.textContent = "Overlay on Frame";
          renderOverlay();
        } else {
          state.heatmapOverlay = { src: "media/" + task.heatmap, type: task.type, region: task.region };
          overlayBtn.textContent = "Hide Overlay";
          var hmImg = new Image();
          hmImg.onload = function () {
            if (state.heatmapOverlay) {
              state.heatmapOverlay._img = hmImg;
              renderOverlay();
            }
          };
          hmImg.src = "media/" + task.heatmap;
        }
      });
      heatmapLabel.appendChild(overlayBtn);
      heatmapSection.appendChild(heatmapLabel);
      var heatmapImg = document.createElement("img");
      heatmapImg.src = "media/" + task.heatmap;
      heatmapImg.alt = "Detection heatmap";
      heatmapSection.appendChild(heatmapImg);
      container.appendChild(heatmapSection);
    }

    results.forEach(function (r, rIdx) {
      // Find matching event for this result
      var ts = r.timestamp !== undefined ? r.timestamp : r.start;
      var tsKey = ts !== undefined ? ts.toFixed(2) : null;
      var matchedEvent = null;
      if (tsKey && eventsByTs[tsKey]) {
        var idx = eventTsIndex[tsKey] || 0;
        if (idx < eventsByTs[tsKey].length) {
          matchedEvent = eventsByTs[tsKey][idx];
          eventTsIndex[tsKey] = idx + 1;
        }
      }

      var isExcluded = matchedEvent && matchedEvent.excluded;
      if (isExcluded && !state.showExcluded) return;
      visibleCount++;

      var row = el("div", "result-row" + (isExcluded ? " excluded" : ""));
      row.dataset.resultIndex = rIdx;

      if (task.type === "color") {
        row.dataset.timestamp = r.start;
        row.appendChild(el("span", "result-timestamp", formatTimestamp(r.start) + " \u2013 " + formatTimestamp(r.end)));
        row.appendChild(el("span", "result-detail", r.duration.toFixed(1) + "s"));
      } else if (task.type === "change") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTimestamp(r.timestamp)));
        var bar = el("div", "result-bar");
        var fill = el("div", "result-bar-fill");
        fill.style.width = Math.round(Math.min(r.magnitude, 1) * 100) + "%";
        fill.style.background = taskTypeColor("change");
        bar.appendChild(fill);
        row.appendChild(bar);
        row.appendChild(el("span", "result-score", (r.magnitude * 100).toFixed(1) + "%"));
      } else if (task.type === "similarity") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTimestamp(r.timestamp)));
        row.appendChild(el("span", "result-score", (r.score * 100).toFixed(1) + "%"));
      } else if (task.type === "text") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTimestamp(r.timestamp)));
        row.appendChild(el("span", "result-detail", r.text_found || ""));
        row.appendChild(el("span", "result-score", (r.confidence * 100).toFixed(0) + "%"));
      } else if (task.type === "numbers") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTimestamp(r.timestamp)));
        row.appendChild(el("span", "result-detail", String(r.number_found)));
      } else if (task.type === "template") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTimestamp(r.timestamp)));
        row.appendChild(el("span", "result-score", (r.best_score * 100).toFixed(1) + "%"));
        row.appendChild(el("span", "result-detail", r.match_count + " match" + (r.match_count !== 1 ? "es" : "")));
      } else if (task.type === "flow") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTimestamp(r.timestamp)));
        var flowBar = el("div", "result-bar");
        var flowFill = el("div", "result-bar-fill");
        flowFill.style.width = Math.round(Math.min(r.magnitude / 20, 1) * 100) + "%";
        flowFill.style.background = taskTypeColor("flow");
        flowBar.appendChild(flowFill);
        row.appendChild(flowBar);
        row.appendChild(el("span", "result-score", r.magnitude.toFixed(2)));
      } else if (task.type === "scene") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTimestamp(r.timestamp)));
        row.appendChild(el("span", "result-detail", r.scene_name));
        row.appendChild(el("span", "result-score", (r.score * 100).toFixed(1) + "%"));
      }

      if (matchedEvent) {
        var btn = el("button", "result-exclude-btn");
        btn.dataset.eventId = matchedEvent.id;
        btn.dataset.excluded = isExcluded ? "true" : "false";
        btn.title = isExcluded ? "Include event" : "Exclude event";
        var icon = isExcluded ? svgDismissIcon() : svgCheckIcon();
        icon.setAttribute("width", "12");
        icon.setAttribute("height", "12");
        btn.appendChild(icon);
        row.appendChild(btn);
      }

      container.appendChild(row);
    });
  }

  function exportResults(format) {
    if (!state.selectedTaskResults || !state.selectedTaskId) return;
    var task = findTask(state.selectedTaskId);
    var filename = "screenspace_" + (task ? task.type : "results") + "_" + state.selectedTaskId;
    var results = state.selectedTaskResults;

    if (format === "json") {
      var blob = new Blob([JSON.stringify(results, null, 2)], { type: "application/json" });
      downloadBlob(blob, filename + ".json");
    } else {
      if (!Array.isArray(results) || results.length === 0) return;
      var keys = Object.keys(results[0]);
      var lines = [keys.join(",")];
      results.forEach(function (r) {
        lines.push(keys.map(function (k) {
          var v = r[k];
          if (typeof v === "string") return '"' + v.replace(/"/g, '""') + '"';
          return v;
        }).join(","));
      });
      var blob = new Blob([lines.join("\n")], { type: "text/csv" });
      downloadBlob(blob, filename + ".csv");
    }
  }

  function downloadBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ---- Keyboard shortcuts ----

  function initKeyboard() {
    document.addEventListener("keydown", function (e) {
      // Don't capture when typing in inputs
      if (e.target.tagName === "INPUT" || e.target.tagName === "TEXTAREA" || e.target.tagName === "SELECT") return;

      if (e.key === "ArrowLeft") {
        e.preventDefault();
        if (state.videoInfo) loadFrame(clamp(state.currentTimestamp - FRAME_STEP, 0, state.videoInfo.duration));
      } else if (e.key === "ArrowRight") {
        e.preventDefault();
        if (state.videoInfo) loadFrame(clamp(state.currentTimestamp + FRAME_STEP, 0, state.videoInfo.duration));
      } else if (e.key === "Escape") {
        if (state.pipetteActive) {
          deactivatePipette();
          return;
        }
        if (state.draggingRegion) {
          var orig = state.draggingRegion.origRegion;
          state.regions[state.draggingRegion.name] = Object.assign({}, state.regions[state.draggingRegion.name], orig);
          state.draggingRegion = null;
          document.body.style.cursor = "";
          document.body.style.userSelect = "";
          renderOverlay();
        } else if (state.resizingRegion) {
          var origR = state.resizingRegion.origRegion;
          state.regions[state.resizingRegion.name] = Object.assign({}, state.regions[state.resizingRegion.name], origR);
          state.resizingRegion = null;
          document.body.style.cursor = "";
          document.body.style.userSelect = "";
          renderOverlay();
        } else if (state.pendingRegion || state.activeRegion) {
          state.pendingRegion = null;
          state.activeRegion = null;
          renderOverlay();
          updateRegionButtons();
          updateRunButton();
        }
        hideRegionNameModal();
      }
    });
  }

  // ---- Panel divider ----

  function initPanelDivider() {
    var handle = qs("#panelDivider");
    var panel = qs("#bottomPanel");
    if (!handle || !panel) return;
    var dragging = false;
    var startY = 0;
    var startHeight = 0;

    var MIN_H = 120;
    var MAX_H = Math.round(window.innerHeight * 0.6);

    function onDown(e) {
      e.preventDefault();
      dragging = true;
      startY = e.clientY || (e.touches && e.touches[0].clientY) || 0;
      startHeight = state.panelHeight;
      handle.classList.add("active");
      document.body.style.cursor = "row-resize";
      document.body.style.userSelect = "none";
    }

    handle.addEventListener("mousedown", onDown);
    handle.addEventListener("touchstart", onDown, { passive: false });

    var rafPending = false;

    function onMove(e) {
      if (!dragging || rafPending) return;
      rafPending = true;
      var clientY = e.clientY || (e.touches && e.touches[0].clientY) || 0;
      requestAnimationFrame(function () {
        var delta = startY - clientY;
        state.panelHeight = Math.max(MIN_H, Math.min(MAX_H, startHeight + delta));
        panel.style.height = state.panelHeight + "px";
        rafPending = false;
      });
    }

    document.addEventListener("mousemove", onMove);
    document.addEventListener("touchmove", onMove, { passive: false });

    function onUp() {
      if (!dragging) return;
      dragging = false;
      handle.classList.remove("active");
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    }

    document.addEventListener("mouseup", onUp);
    document.addEventListener("touchend", onUp);
  }

  // ---- Preview resize ----

  function initPreviewResize() {
    var handle = qs("#previewResizeHandle");
    var container = qs("#frameContainer");
    if (!handle || !container) return;
    var dragging = false;
    var startX = 0;
    var startWidthPx = 0;
    var parentWidth = 0;

    var MIN_PCT = 30;
    var MAX_PCT = 100;

    function onDown(e) {
      e.preventDefault();
      e.stopPropagation();
      dragging = true;
      startX = e.clientX || (e.touches && e.touches[0].clientX) || 0;
      startWidthPx = container.getBoundingClientRect().width;
      parentWidth = container.parentElement.getBoundingClientRect().width;
      handle.classList.add("active");
      document.body.style.cursor = "nwse-resize";
      document.body.style.userSelect = "none";
    }

    handle.addEventListener("mousedown", onDown);
    handle.addEventListener("touchstart", onDown, { passive: false });

    var rafPending = false;

    function onMove(e) {
      if (!dragging || rafPending) return;
      rafPending = true;
      var clientX = e.clientX || (e.touches && e.touches[0].clientX) || 0;
      requestAnimationFrame(function () {
        var delta = clientX - startX;
        var newWidthPx = startWidthPx + delta;
        var pct = Math.max(MIN_PCT, Math.min(MAX_PCT, (newWidthPx / parentWidth) * 100));
        state.previewMaxWidth = Math.round(pct);
        container.style.maxWidth = state.previewMaxWidth + "%";
        rafPending = false;
      });
    }

    document.addEventListener("mousemove", onMove);
    document.addEventListener("touchmove", onMove, { passive: false });

    function onUp() {
      if (!dragging) return;
      dragging = false;
      handle.classList.remove("active");
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    }

    document.addEventListener("mouseup", onUp);
    document.addEventListener("touchend", onUp);

    handle.addEventListener("dblclick", function (e) {
      e.preventDefault();
      e.stopPropagation();
      state.previewMaxWidth = MAX_PCT;
      container.style.maxWidth = "";
    });
  }

  // ---- Init ----

  document.addEventListener("DOMContentLoaded", function () {
    initThemeToggle();
    initFrameControls();
    initRegionDrawing();
    initTimeline();
    initWorkflowTabs();
    initRunButton();
    initTaskQueue();
    initPauseButton();
    initTaskFilters();
    initResultsPanel();
    initPanelDivider();
    initPreviewResize();
    initKeyboard();
    checkNavLinks();

    // Participant select
    qs("#participantSelect").addEventListener("change", function () {
      var pid = this.value;
      if (pid) {
        selectParticipant(pid);
        state.runParticipants = [pid];
        renderRunParticipantPicker();
      }
    });

    // Close run picker on outside click
    document.addEventListener("click", function (e) {
      if (!e.target.closest(".run-picker-wrap")) closeRunPicker();
    });

    // Load initial data
    apiGet("api/participants")
      .then(function (data) {
        if (!data.ok) return;
        state.participants = (data.participants || []).filter(function (p) { return p.has_video; });
        renderParticipantSelect();
        if (state.participants.length > 0) {
          var first = state.participants[0].id;
          selectParticipant(first);
          state.runParticipants = [first];
        }
        renderRunParticipantPicker();
      })
      .catch(function () { showToast("Failed to load participants"); });

    apiGet("api/regions")
      .then(function (data) {
        if (data.ok) {
          state.regions = data.regions || {};
          renderRegionChips();
          updateRegionButtons();
          renderOverlay();
        }
      })
      .catch(function () {});

    apiGet("api/stashes")
      .then(function (data) {
        if (data.ok) {
          state.stashes = data.stashes || [];
          renderStashCards();
          renderRunRegionPicker();
        }
      })
      .catch(function () {});

    apiGet("api/tasks")
      .then(function (data) {
        if (data.ok) {
          state.tasks = data.tasks || [];
          renderTaskList();
          renderTimeline();
          if (state.tasks.some(function (t) { return t.status === "queued" || t.status === "running"; })) {
            startPolling();
          }
        }
      })
      .catch(function () {});
  });

})();
