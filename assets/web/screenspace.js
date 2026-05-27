/* clipgen Screenspace page.
 *
 * Frame canvas + overlay editor for defining regions and templates on a video
 * frame, then running detector tasks (color/change/similarity/text/numbers/
 * timelapse/template/flow/scene/inactivity, plus the multitool chain).
 *
 * Two patterns recur and are worth knowing up front:
 *
 *   - Request versioning. Frame fetches and participant switches are async,
 *     and a fast user can issue several before the first lands. Every async
 *     entry point bumps a `_*RequestVersion` counter and stale callbacks
 *     compare against the latest version on resolution. See `_fetchFrame`.
 *   - Overlay interactions. The overlay canvas is a small state machine over
 *     four mutually-exclusive modes (drag template / drag region / resize
 *     region / draw region), driven by `state.draggingTemplate`,
 *     `state.draggingRegion`, `state.resizingRegion`, `state.drawingRegion`.
 */

(function () {
  "use strict";

  var FRAME_STEP = 1.0;
  var VIDEO_SPEEDS = [0.5, 1, 2, 3, 5];

  var TASK_COLORS = DETECTOR_COLORS;

  var SS_TASK_ICON_TYPES = {
    multitool: 1, color: 1, change: 1, similarity: 1, text: 1,
    numbers: 1, template: 1, flow: 1, scene: 1, inactivity: 1,
  };

  // Build a span that renders the task icon via mask-image (see .ss-task-icon
  // family in screenspace.css). Color follows currentColor on the parent.
  function buildTypeIcon(type) {
    if (!SS_TASK_ICON_TYPES[type]) return null;
    var span = document.createElement("span");
    span.className = "ss-task-icon ss-task-icon--" + type;
    return span;
  }

  // Generic mask-image icon span. `name` is the basename of a file in
  // assets/icons/ (no extension); `sizeClass` is an optional .ss-icon
  // modifier (e.g. "ss-icon--sm"). See .ss-icon family in screenspace.css.
  function iconSpan(name, sizeClass) {
    var span = el("span", "ss-icon" + (sizeClass ? " " + sizeClass : ""));
    applyMaskIcon(span, 'url("/screenspace/icons/' + name + '.svg")');
    return span;
  }

  var TIMELINE_CANVAS_HEIGHT = 64;

  var _paletteDocListeners = null;
  // Cached HSV hidden inputs for the single-tool color picker. Populated by
  // renderColorParams() when the panel is built; reused by setTargetColor,
  // updateColorPreview, _collectPreviewParams, etc. so they don't re-query
  // the DOM on every drag tick. Null when no color tool panel is active.
  var _colorHiddenInputs = null;

  var REGION_COLOR_COUNT = 8;

  // Re-reads each call so dev-token-tweak widget overrides take effect live.
  function bottomPanelHeightFromToken() {
    var v = getComputedStyle(document.documentElement)
      .getPropertyValue("--bottom-panel-height")
      .trim();
    return parseInt(v, 10) || 400;
  }

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
    restoreMarkersOnEdit: true,
    activeWorkflow: "color",
    referenceTimestamp: null,
    sceneReferences: [],
    tasks: [],
    selectedTaskId: null,
    hoveredTaskId: null,
    selectedTaskResults: null,
    pollTimer: null,
    eventSource: null,
    queuePaused: false,
    timelineDragging: false,
    panelHeight: bottomPanelHeightFromToken(),
    panelHeightBeforeCollapse: bottomPanelHeightFromToken(),
    bottomCollapsed: false,
    previewMaxWidth: 100,
    taskFilter: null,
    pipetteActive: false,
    runParticipants: [],
    runRegions: [],
    scanMode: "normal",
    taskEvents: {},
    showExcluded: true,
    certaintyCutoff: 0,
    showRegionLabels: true,
    showRegionOverlays: true,
    stashes: [],
    previewRegions: null,
    resultOverlay: null,
    heatmapOverlay: null,
    uploadedTemplate: null,
    uploadedTemplateImg: null,
    templateScalePreview: 1.0,
    templateOverlayPos: null,
    draggingTemplate: null,
    multitoolSteps: [],
    hoveredResultSceneName: null,
    videoPlaying: false,
    videoMuted: false,
    videoPlaybackRate: 1,
    modelViewOpen: false,
    overlayEnabled: false,
    overlayLayer: null,
    overlayBlinkActive: false,
    overlayImage: null,
    overlayImageObjectUrl: null,
    overlayImageTimestamp: null,
    overlayImageTool: null,
    overlayLayerSpec: {},
    rightPaneTab: "queue",
    resultsSwitcherOpen: false,
    amplitudeGraphEnabled: false,
  };

  var _timelineHitRects = [];
  var _overlayRaf = 0;
  var _playheadRaf = 0;
  var _cachedOverlayRect = null;
  var _cachedTimelineRect = null;

  function getTimelineRect(canvas) {
    if (!_cachedTimelineRect) _cachedTimelineRect = canvas.getBoundingClientRect();
    return _cachedTimelineRect;
  }
  var _lastPollFingerprint = "";
  var _preloadedFrames = {};
  // Per-participant source-video mtime_ns, sourced from /api/participants and
  // /api/video/info. Used as a ?v= cache-bust suffix on frame and stream URLs
  // so a re-encoded or replaced source file invalidates HTTP, backend, and
  // blob caches together. Empty string means "no version known yet".
  var _videoVersions = {};

  window.addEventListener("pagehide", function () {
    Object.keys(_preloadedFrames).forEach(function (pid) {
      try { URL.revokeObjectURL(_preloadedFrames[pid]); } catch (_) {}
      delete _preloadedFrames[pid];
    });
  });
  var _participantRequestVersion = 0;
  var _frameRequestVersion = 0;
  var _resultsRequestVersion = 0;
  var _heatmapOverlayRequestVersion = 0;

  // Region palette is screenspace-specific (REGION_COLOR_COUNT entries from
  // --region-color-1..N); the common canvas colors come from the shared
  // getCanvasThemeColors() cache in utils.js, which auto-invalidates on
  // theme toggle.
  var _cachedRegionPalette = null;

  function refreshThemeColors() {
    invalidateCanvasThemeColors();
    _cachedRegionPalette = null;
  }

  function getThemeColors() {
    var base = getCanvasThemeColors();
    if (!_cachedRegionPalette) {
      var cs = getComputedStyle(document.documentElement);
      var palette = [];
      for (var i = 1; i <= REGION_COLOR_COUNT; i++) {
        palette.push(cs.getPropertyValue("--region-color-" + i).trim() || "#3b82f6");
      }
      _cachedRegionPalette = palette;
    }
    return {
      fg: base.fg,
      bg: base.bg,
      surfaceAlt: base.surfaceAlt,
      border: base.border,
      textDim: base.textDim,
      accent: base.accent,
      fontMono: base.fontMono,
      regionPalette: _cachedRegionPalette,
    };
  }

  // ---- Helpers ----

  function numberOrDefault(value, fallback) {
    var n = parseFloat(value);
    return isNaN(n) ? fallback : n;
  }

  function intOrDefault(value, fallback) {
    var n = parseInt(value, 10);
    return isNaN(n) ? fallback : n;
  }

  function regionColorForIndex(i) {
    var palette = getThemeColors().regionPalette;
    return palette[i % palette.length];
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

  function taskRegionPixels(task) {
    var r = task && task.region_coords;
    if (!r) return null;
    return {
      x: Math.round(Number(r.x) || 0),
      y: Math.round(Number(r.y) || 0),
      w: Math.round(Number(r.w) || 0),
      h: Math.round(Number(r.h) || 0),
    };
  }

  function taskTypeColor(type) {
    return TASK_COLORS[type] || "#888";
  }

  // Compact one-liner describing a multitool step's criteria, used in result rows.
  // Mirrors the keys gathered by gatherMultitoolStepParams so the displayed values
  // match what the user configured at task-create time.
  function formatMultitoolStepParams(step) {
    if (!step) return "";
    var t = step.type;
    if (t === "color") {
      var tc = step.target_color || {};
      return "H" + (tc.h || 0) + "° S" + (tc.s || 0) + " V" + (tc.v || 0);
    }
    if (t === "change") return ">" + ((step.threshold || 0) * 100).toFixed(0) + "%";
    if (t === "similarity") return "≥" + ((step.threshold || 0) * 100).toFixed(0) + "%";
    if (t === "text") return "“" + (step.search_string || "") + "”";
    if (t === "numbers") {
      var opSym = { gt: ">", lt: "<", eq: "=", gte: "≥", lte: "≤" }[step.operator] || step.operator || "";
      return (opSym + " " + step.target_value).trim();
    }
    if (t === "template") return "≥" + ((step.threshold || 0) * 100).toFixed(0) + "%";
    if (t === "flow") return ">" + (step.magnitude_threshold || 0).toFixed(1);
    if (t === "scene") {
      var refs = step.scene_references || [];
      if (refs.length === 1) return refs[0].name || "1 ref";
      return refs.length + " refs";
    }
    if (t === "inactivity") return "≥" + (step.threshold || 0) + "s";
    return "";
  }

  // Confidence bar mirrors the prototype's ConfBar: 4 px tall, hue-tinted fill,
  // opacity ramps from 0.4 (low) to 1.0 (full) so high-confidence rows feel
  // saturated while low ones recede.
  function buildConfBar(value, type) {
    var v = Math.max(0, Math.min(1, Number(value) || 0));
    var bar = el("div", "result-bar");
    var fill = el("div", "result-bar-fill");
    fill.style.width = Math.round(v * 100) + "%";
    fill.style.background = taskTypeColor(type);
    fill.style.opacity = (0.4 + v * 0.6).toFixed(2);
    bar.appendChild(fill);
    return bar;
  }

  function frameUrl(pid, ts) {
    var base = "api/video/frame/" + encodeURIComponent(pid) + "/" + Number(ts).toFixed(6);
    var v = _videoVersions[pid];
    return v ? base + "?v=" + encodeURIComponent(v) : base;
  }

  function videoStreamUrl(pid) {
    var base = "api/video/stream/" + encodeURIComponent(pid);
    var v = _videoVersions[pid];
    return v ? base + "?v=" + encodeURIComponent(v) : base;
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
    chevron.appendChild(iconSpan("chevron-down", "ss-icon--xs"));
    btn.appendChild(chevron);
  }

  function closeRunPicker() {
    var panels = qsa(".run-picker-panel");
    var btns = qsa(".run-picker-btn");
    for (var i = 0; i < panels.length; i++) panels[i].classList.add("hidden");
    for (i = 0; i < btns.length; i++) btns[i].classList.remove("open");
  }

  var FULL_FRAME_REGION_NAME = "full_frame";

  function activeRegionRef(name) {
    return { source: "active", name: name };
  }

  function stashRegionRef(stash, name) {
    return {
      source: "stash",
      stash_id: stash.id,
      stash_name: stash.name,
      name: name,
    };
  }

  function fullFrameRegionRef() {
    return { source: "full_frame", name: FULL_FRAME_REGION_NAME };
  }

  function isFullFrameRef(ref) {
    return !!ref && ref.source === "full_frame";
  }

  function normalizeRegionRef(ref) {
    if (!ref) return null;
    if (typeof ref === "string") {
      if (ref === FULL_FRAME_REGION_NAME) return fullFrameRegionRef();
      return activeRegionRef(ref);
    }
    if (ref.source === "full_frame") return fullFrameRegionRef();
    if (ref.source === "stash") {
      var stashName = ref.stash_name;
      if (!stashName) {
        for (var i = 0; i < state.stashes.length; i++) {
          if (state.stashes[i].id === ref.stash_id) {
            stashName = state.stashes[i].name;
            break;
          }
        }
      }
      return {
        source: "stash",
        stash_id: ref.stash_id,
        stash_name: stashName,
        name: ref.name,
      };
    }
    return activeRegionRef(ref.name);
  }

  function regionRefKey(ref) {
    var r = normalizeRegionRef(ref);
    if (!r) return "";
    if (r.source === "full_frame") return "full_frame";
    return r.source === "stash" ? "stash:" + r.stash_id + ":" + r.name : "active:" + r.name;
  }

  function regionRefLabel(ref) {
    var r = normalizeRegionRef(ref);
    if (!r) return "";
    if (r.source === "full_frame") return "Full frame";
    return r.source === "stash" ? r.name + " · " + (r.stash_name || "stash") : r.name;
  }

  function regionRefPayload(ref) {
    var r = normalizeRegionRef(ref);
    if (!r) return null;
    if (r.source === "full_frame") return { source: "full_frame" };
    if (r.source === "stash") {
      return { source: "stash", stash_id: r.stash_id, name: r.name };
    }
    return { source: "active", name: r.name };
  }

  function hasRunRegion(ref) {
    var key = regionRefKey(ref);
    return state.runRegions.some(function (r) { return regionRefKey(r) === key; });
  }

  function addRunRegion(ref) {
    if (!hasRunRegion(ref)) state.runRegions.push(normalizeRegionRef(ref));
  }

  function removeRunRegion(ref) {
    var key = regionRefKey(ref);
    state.runRegions = state.runRegions.filter(function (r) { return regionRefKey(r) !== key; });
  }

  function allAvailableRegionRefs() {
    var refs = [fullFrameRegionRef()];
    Object.keys(state.regions).forEach(function (name) {
      refs.push(activeRegionRef(name));
    });
    state.stashes.forEach(function (stash) {
      Object.keys(stash.regions).forEach(function (name) {
        refs.push(stashRegionRef(stash, name));
      });
    });
    return refs;
  }

  function buildFullFrameIcon() {
    var icon = el("span", "run-picker-fullframe-icon");
    applyMaskIcon(icon, 'url("/screenspace/icons/arrows-pointing-out.svg")');
    return icon;
  }

  function availableRegionRefByKey(key) {
    var refs = allAvailableRegionRefs();
    for (var i = 0; i < refs.length; i++) {
      if (regionRefKey(refs[i]) === key) return refs[i];
    }
    return null;
  }

  function renderRunRegionPicker() {
    var wrap = qs("#runRegionPicker");
    if (!wrap) return;
    wrap.innerHTML = "";
    var names = Object.keys(state.regions);
    var activeRefs = names.map(function (name) { return activeRegionRef(name); });
    var allRefs = allAvailableRegionRefs();
    var availableKeys = {};
    allRefs.forEach(function (ref) { availableKeys[regionRefKey(ref)] = true; });
    // Remove any runRegions that no longer exist in active or stashes.
    state.runRegions = state.runRegions
      .map(normalizeRegionRef)
      .filter(function (r) { return r && availableKeys[regionRefKey(r)]; });
    // Auto-select the active region when no explicit selection has been made
    if (state.runRegions.length === 0 && state.activeRegion && names.indexOf(state.activeRegion) >= 0) {
      state.runRegions = [activeRegionRef(state.activeRegion)];
    }

    var btn = el("button", "run-picker-btn");
    btn.type = "button";
    updateRegionPickerBtnText(btn);

    var panel = el("div", "run-picker-panel hidden");

    // Full-frame entry — always available, sits above the active/stash regions.
    var fullFrameRef = fullFrameRegionRef();
    var fullFrameLbl = document.createElement("label");
    fullFrameLbl.className = "run-picker-fullframe";
    var fullFrameCb = document.createElement("input");
    fullFrameCb.type = "checkbox";
    fullFrameCb.value = regionRefKey(fullFrameRef);
    fullFrameCb.checked = hasRunRegion(fullFrameRef);
    fullFrameCb.addEventListener("change", function () {
      if (fullFrameCb.checked) {
        addRunRegion(fullFrameRef);
      } else {
        removeRunRegion(fullFrameRef);
      }
      updateRegionPickerBtnText(btn);
      updateRunButton();
    });
    fullFrameLbl.appendChild(fullFrameCb);
    fullFrameLbl.appendChild(buildFullFrameIcon());
    fullFrameLbl.appendChild(el("span", "run-picker-label-text", "Full frame"));
    panel.appendChild(fullFrameLbl);

    if (names.length > 0) {
      var toggleAll = el("span", "run-picker-toggle-all");
      var allActiveSelected = activeRefs.every(function (ref) { return hasRunRegion(ref); });
      toggleAll.textContent = allActiveSelected ? "Deselect all" : "Select all";
      toggleAll.addEventListener("click", function () {
        var allSelected = activeRefs.every(function (ref) { return hasRunRegion(ref); });
        if (allSelected) {
          activeRefs.forEach(removeRunRegion);
        } else {
          activeRefs.forEach(addRunRegion);
        }
        var cbs = panel.querySelectorAll(".run-picker-active-region input[type=checkbox]");
        for (var i = 0; i < cbs.length; i++) cbs[i].checked = !allSelected;
        toggleAll.textContent = allSelected ? "Select all" : "Deselect all";
        updateRegionPickerBtnText(btn);
        updateRunButton();
      });
      panel.appendChild(toggleAll);

      names.forEach(function (name, idx) {
        var color = regionColorForIndex(idx);
        var ref = activeRegionRef(name);
        var lbl = document.createElement("label");
        lbl.className = "run-picker-active-region";
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.value = regionRefKey(ref);
        cb.checked = hasRunRegion(ref);
        cb.addEventListener("change", function () {
          if (cb.checked) {
            addRunRegion(ref);
          } else {
            removeRunRegion(ref);
          }
          toggleAll.textContent = activeRefs.every(function (activeRef) {
            return hasRunRegion(activeRef);
          }) ? "Deselect all" : "Select all";
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
        var ref = stashRegionRef(stash, name);
        var lbl = document.createElement("label");
        lbl.className = "stash-folder-item";
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.value = regionRefKey(ref);
        cb.checked = hasRunRegion(ref);
        cb.addEventListener("change", function () {
          if (cb.checked) {
            addRunRegion(ref);
          } else {
            removeRunRegion(ref);
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
      : n === 1 ? regionRefLabel(state.runRegions[0])
      : n + " regions";
    btn.innerHTML = "";
    if (n === 1 && isFullFrameRef(state.runRegions[0])) {
      btn.appendChild(buildFullFrameIcon());
    }
    btn.appendChild(el("span", "run-picker-btn-text", text));
    var chevron = el("span", "chevron");
    chevron.appendChild(iconSpan("chevron-down", "ss-icon--xs"));
    btn.appendChild(chevron);
  }

  var FAST_SCAN_DESCRIPTIONS = {
    color: "Lower resolution, skips unchanged frames",
    change: "Lower resolution, skips unchanged frames",
    similarity: "Lower resolution, skips unchanged frames",
    text: "Skips unchanged frames",
    numbers: "Skips unchanged frames",
    template: "Downscales template 2\u00D7, skips unchanged frames",
    flow: "Lower resolution, skips unchanged frames",
    scene: "Lower resolution, skips unchanged frames",
    inactivity: "Lower resolution, skips unchanged frames",
    multitool: "Skips unchanged frames, widens interval"
  };

  var PARAM_DESCRIPTIONS = {
    _shared: {
      "Event label":      "Tag added to each detected event for filtering",
      "Detect first":     "Stop after the first match is found",
      "Region":           "Which screen region this step analyzes",
    },
    color: {
      "Tolerance":        "How far from the target color still counts as a match",
      "Hex color":        "Target color in hex notation",
    },
    change: {
      "Threshold":        "Minimum pixel-change ratio to trigger a detection",
      "Noise Thr.":       "Ignore changes below this pixel intensity",
      "Noise":            "Ignore changes below this pixel intensity",
    },
    similarity: {
      "Reference":        "Capture a frame to compare against",
      "Threshold":        "Minimum similarity score to count as a match",
    },
    text: {
      "Search text":      "Exact or partial text to find on screen",
      "Search":           "Exact or partial text to find on screen",
      "Fuzzy Thr.":       "Minimum fuzzy-match score (1.0 = exact match)",
      "Fuzzy":            "Minimum fuzzy-match score (1.0 = exact match)",
      "Language":         "OCR language for text recognition",
    },
    numbers: {
      "Operator":         "Comparison operator for the detected number",
      "Target value":     "Number to compare the detected value against",
      "Target":           "Number to compare the detected value against",
      "Range":            "Min and max bounds for the in-range check",
    },
    timelapse: {
      "Speed":            "Playback speed multiplier for the output",
      "Sample every":     "Seconds between captured frames (0 = every frame)",
      "Format":           "Output file format: video or animated GIF",
    },
    template: {
      "Template":         "Capture or upload a reference image to match",
      "Threshold":        "Minimum match score to trigger a detection",
    },
    flow: {
      "Magnitude":        "Minimum optical-flow magnitude to count as motion",
    },
    scene: {
      "Add Scene":        "Name and capture a reference frame for a scene",
    },
    inactivity: {
      "Sensitivity":      "Pixel-change level below which the frame is idle",
      "Min duration (s)": "Seconds of stillness required to trigger",
    },
  };

  function renderScanModePicker() {
    var wrap = qs("#runScanModePicker");
    if (!wrap) return;
    wrap.innerHTML = "";

    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "scan-toggle-btn";

    var icon = el("span", "scan-toggle-icon");
    applyMaskIcon(icon, 'url("/screenspace/icons/chevron-double-right.svg")');
    btn.appendChild(icon);

    function updateState() {
      var isFast = state.scanMode === "fast";
      btn.classList.toggle("active", isFast);
    }
    updateState();
    btn._updateScanState = updateState;

    attachHoverTooltip(btn, function () {
      var isFast = state.scanMode === "fast";
      var label = isFast ? "Fast scan enabled" : "Enable fast scan";
      var desc = FAST_SCAN_DESCRIPTIONS[state.activeWorkflow];
      if (desc) label += "\n" + desc;
      return label;
    }, { align: "center", multiline: true });

    btn.addEventListener("click", function () {
      state.scanMode = state.scanMode === "fast" ? "normal" : "fast";
      updateState();
    });

    wrap.appendChild(btn);
  }

  function initParamTooltips() {
    var container = qs("#workflowParams");
    if (!container) return;

    var tooltip = createTooltip({ align: "center" });

    function getToolType(labelEl) {
      var stepCard = labelEl.closest(".multitool-step");
      if (stepCard) {
        var idx = parseInt(stepCard.dataset.stepIdx, 10);
        var step = state.multitoolSteps[idx];
        return step ? step.type : null;
      }
      return state.activeWorkflow;
    }

    function getDescription(labelText, toolType) {
      if (!toolType) return null;
      var toolMap = PARAM_DESCRIPTIONS[toolType];
      if (toolMap && toolMap[labelText]) return toolMap[labelText];
      var shared = PARAM_DESCRIPTIONS._shared;
      if (shared && shared[labelText]) return shared[labelText];
      return null;
    }

    container.addEventListener("mouseenter", function (e) {
      var label = e.target.closest(".param-label");
      if (!label) return;
      var text = label.textContent.trim();
      var desc = getDescription(text, getToolType(label));
      if (!desc) return;
      tooltip.show(label, desc);
    }, true);

    container.addEventListener("mouseleave", function (e) {
      if (e.target.closest && e.target.closest(".param-label")) {
        tooltip.hide();
      }
    }, true);
  }

  // Switching participants must clear the entire frame/overlay/playback
  // pipeline together — keeping any one of these around (frameImage, video
  // src, scene refs, pending fetch ids) would let the previous participant's
  // state leak into the new one. Bumping the request-version counters also
  // invalidates any in-flight frame/heatmap loads from the prior participant.
  function selectParticipant(pid, initialTimestamp) {
    var participantRequestVersion = ++_participantRequestVersion;
    _frameRequestVersion += 1;
    _heatmapOverlayRequestVersion += 1;
    state.selectedParticipant = pid;
    setStoredUIStateField("screenspace", "selectedParticipant", pid);
    state.currentTimestamp = 0;
    state.videoInfo = null;
    state.frameImage = null;
    state.frameLoading = false;
    state.referenceTimestamp = null;
    state.sceneReferences = [];
    state.resultOverlay = null;
    state.heatmapOverlay = null;
    // Reset video playback
    var videoEl = qs("#videoPlayer");
    if (state.videoPlaying) videoEl.pause();
    state.videoPlaying = false;
    videoEl.classList.remove("active");
    videoEl.removeAttribute("src");
    videoEl.load();
    applyPlaybackRate();
    qs("#frameCanvas").classList.remove("video-active");
    updateVideoButtons();
    _pendingFrameTs = null;
    _loadedFrameTs = null;
    qs("#participantSelect").value = pid;
    qs("#videoInfo").textContent = "";
    qs("#frameEmpty").classList.remove("hidden");
    setInfoParticipant(pid);

    apiGet("api/video/info/" + encodeURIComponent(pid))
      .then(function (data) {
        if (participantRequestVersion !== _participantRequestVersion || pid !== state.selectedParticipant) return;
        if (!data.ok) return;
        state.videoInfo = data.info;
        // If the server reports a different mtime than we last saw, the
        // source file was replaced \u2014 drop the stale frame-0 blob so the
        // next loadFrame(0) hits the API and the new ?v= URL.
        var newVersion = data.info.version != null ? String(data.info.version) : "";
        var prevVersion = _videoVersions[pid] || "";
        _videoVersions[pid] = newVersion;
        if (newVersion !== prevVersion && _preloadedFrames[pid]) {
          try { URL.revokeObjectURL(_preloadedFrames[pid]); } catch (_) {}
          delete _preloadedFrames[pid];
        }
        var parts = [];
        if (data.info.duration) parts.push(formatDuration(data.info.duration));
        if (data.info.width && data.info.height) parts.push(data.info.width + "x" + data.info.height);
        if (data.info.fps) parts.push(Math.round(data.info.fps) + "fps");
        qs("#videoInfo").textContent = parts.join(" \u00b7 ");
        renderTimeline();
        // Preload video source for instant playback
        qs("#videoPlayer").src = videoStreamUrl(pid);
        loadFrame(initialTimestamp !== undefined ? initialTimestamp : 0);
      })
      .catch(function () { showToast("Failed to load video info"); });

    apiGet("api/participants/" + encodeURIComponent(pid) + "/notes")
      .then(function (data) {
        if (participantRequestVersion !== _participantRequestVersion || pid !== state.selectedParticipant) return;
        if (!data.ok) return;
        renderInfoNotes(data.notes || "");
      })
      .catch(function () {});

    apiGet("api/participants/" + encodeURIComponent(pid) + "/issues")
      .then(function (data) {
        if (participantRequestVersion !== _participantRequestVersion || pid !== state.selectedParticipant) return;
        if (!data.ok) return;
        renderInfoIssues(data.issues || []);
      })
      .catch(function () {});
  }

  // ---- Info panel ----

  function setInfoParticipant(pid) {
    var el = qs("#ssInfoParticipant");
    if (el) el.textContent = pid || "\u2014";
    qs("#ssInfoNotes").value = "";
    qs("#ssInfoIssuesBlock").classList.add("hidden");
    qs("#ssInfoIssues").innerHTML = "";
  }

  function renderInfoNotes(notes) {
    var ta = qs("#ssInfoNotes");
    if (!ta) return;
    if (document.activeElement !== ta) ta.value = notes;
  }

  function renderInfoIssues(issues) {
    var block = qs("#ssInfoIssuesBlock");
    var list = qs("#ssInfoIssues");
    if (!block || !list) return;
    list.innerHTML = "";
    if (!issues || !issues.length) {
      block.classList.add("hidden");
      return;
    }
    block.classList.remove("hidden");
    var frag = document.createDocumentFragment();
    issues.forEach(function (issue) {
      var li = document.createElement("li");
      li.className = "ss-info-issue";
      var dot = document.createElement("span");
      dot.className = "ss-info-issue-dot " + (severityClass(issue.severity) || "");
      var text = document.createElement("span");
      text.className = "ss-info-issue-text";
      text.textContent = issue.observation || "(no observation)";
      li.appendChild(dot);
      li.appendChild(text);
      if (issue.timestamp != null) {
        var ts = document.createElement("span");
        ts.className = "ss-info-issue-ts";
        ts.textContent = formatTime(issue.timestamp);
        li.appendChild(ts);
        li.classList.add("ss-info-issue--clickable");
        li.addEventListener("click", (function (t) {
          return function () { loadFrame(t); };
        })(issue.timestamp));
      }
      frag.appendChild(li);
    });
    list.appendChild(frag);
  }

  function initInfoNotes() {
    var ta = qs("#ssInfoNotes");
    if (!ta) return;
    var saveTimer = null;
    ta.addEventListener("input", function () {
      var pid = state.selectedParticipant;
      if (!pid) return;
      var value = ta.value;
      if (saveTimer) clearTimeout(saveTimer);
      saveTimer = setTimeout(function () {
        if (pid !== state.selectedParticipant) return;
        apiPut("api/participants/" + encodeURIComponent(pid) + "/notes", { notes: value })
          .catch(function () { showToast("Failed to save notes"); });
      }, 500);
    });
  }

  function applyInfoPanelCollapsed(collapsed) {
    qs("#ssInfoPanel").classList.toggle("hidden", collapsed);
    qs("#ssInfoExpandBtn").classList.toggle("hidden", !collapsed);
  }

  function initInfoPanelCollapse() {
    var stored = getStoredUIState("screenspace");
    applyInfoPanelCollapsed(!!stored.infoPanelCollapsed);
    qs("#ssInfoCollapseBtn").addEventListener("click", function () {
      applyInfoPanelCollapsed(true);
      setStoredUIStateField("screenspace", "infoPanelCollapsed", true);
    });
    qs("#ssInfoExpandBtn").addEventListener("click", function () {
      applyInfoPanelCollapsed(false);
      setStoredUIStateField("screenspace", "infoPanelCollapsed", false);
    });
  }

  function applyInfoSectionCollapsed(section, collapsed) {
    if (!section) return;
    section.setAttribute("data-collapsed", collapsed ? "true" : "false");
    var header = section.querySelector(".ss-info-section-header");
    if (header) header.setAttribute("aria-expanded", collapsed ? "false" : "true");
  }

  function initInfoSections() {
    var stored = getStoredUIState("screenspace");
    var sections = (stored.infoSectionsCollapsed && typeof stored.infoSectionsCollapsed === "object")
      ? stored.infoSectionsCollapsed
      : {};
    var headers = document.querySelectorAll(".ss-info-section-header");
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i];
      var section = header.closest(".ss-info-section");
      var name = section ? section.getAttribute("data-section") : null;
      if (!name) continue;
      applyInfoSectionCollapsed(section, !!sections[name]);
      header.addEventListener("click", function () {
        var sec = this.closest(".ss-info-section");
        if (!sec) return;
        var n = sec.getAttribute("data-section");
        var st = getStoredUIState("screenspace");
        var s = (st.infoSectionsCollapsed && typeof st.infoSectionsCollapsed === "object")
          ? st.infoSectionsCollapsed
          : {};
        var newCollapsed = !s[n];
        s[n] = newCollapsed;
        setStoredUIStateField("screenspace", "infoSectionsCollapsed", s);
        applyInfoSectionCollapsed(sec, newCollapsed);
      });
    }
  }

  // ---- Frame viewer ----

  function seekPlayhead(timestamp) {
    state.currentTimestamp = timestamp;
    qs("#timestampInput").value = formatTime(timestamp, { decimals: 1 });
    renderPlayhead();
    persistVideoTime(timestamp);
  }

  var _pendingFrameTs = null;
  var _loadedFrameTs = null;

  function loadFrame(timestamp) {
    if (!state.selectedParticipant) return;
    if (state.videoPlaying) {
      var video = qs("#videoPlayer");
      video.pause();
      state.videoPlaying = false;
      video.classList.remove("active");
      qs("#frameCanvas").classList.remove("video-active");
      updateVideoButtons();
    }
    seekPlayhead(timestamp);
    refreshModelView({ debounce: true });
    if (state.frameLoading) {
      _pendingFrameTs = timestamp;
      return;
    }
    _fetchFrame(timestamp);
  }

  // Frame loads are async and the user can scrub faster than the network.
  // We coalesce: at most one in-flight image at a time. While one is loading,
  // newer requests park their timestamp in `_pendingFrameTs` (loadFrame above)
  // and the onload/onerror handler picks it up after the current load
  // completes. `frameRequestVersion` plus the participant check rejects stale
  // images whose participant has since changed, so we never paint a frame
  // belonging to a participant the user already switched away from.
  function _fetchFrame(timestamp) {
    var participantId = state.selectedParticipant;
    var frameRequestVersion = ++_frameRequestVersion;
    state.frameLoading = true;
    _pendingFrameTs = null;
    _loadedFrameTs = timestamp;

    // Use preloaded frame 0 if available
    var preloaded = (timestamp === 0 && _preloadedFrames[participantId])
      ? _preloadedFrames[participantId] : null;

    var img = new Image();
    img.onload = function () {
      if (frameRequestVersion !== _frameRequestVersion || participantId !== state.selectedParticipant) return;
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
      if (frameRequestVersion !== _frameRequestVersion || participantId !== state.selectedParticipant) return;
      state.frameLoading = false;
      if (_pendingFrameTs !== null && _pendingFrameTs !== _loadedFrameTs) {
        _fetchFrame(_pendingFrameTs);
      }
    };
    img.src = preloaded || frameUrl(participantId, timestamp);
  }

  function initFrameControls() {
    qs("#framePrev").appendChild(iconSpan("chevron-left", "ss-icon--sm"));
    qs("#frameNext").appendChild(iconSpan("chevron-right", "ss-icon--sm"));
    var input = qs("#timestampInput");

    input.addEventListener("change", function () {
      var ts = parseTimestamp(input.value);
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
      var ts = clamp(state.currentTimestamp - FRAME_STEP, 0, Math.max(0, state.videoInfo.duration - 0.001));
      loadFrame(ts);
    });

    qs("#frameNext").addEventListener("click", function () {
      if (!state.videoInfo) return;
      var ts = clamp(state.currentTimestamp + FRAME_STEP, 0, Math.max(0, state.videoInfo.duration - 0.001));
      loadFrame(ts);
    });
  }

  // ---- Video playback ----

  function initVideoPlayback() {
    var video = qs("#videoPlayer");
    var playBtn = qs("#videoPlayBtn");
    var muteBtn = qs("#videoMuteBtn");

    playBtn.appendChild(iconSpan("play"));
    muteBtn.appendChild(iconSpan("speaker-wave"));

    playBtn.addEventListener("click", function () {
      if (state.videoPlaying) {
        pauseVideo();
      } else {
        playVideo();
      }
    });

    muteBtn.addEventListener("click", function () {
      state.videoMuted = !state.videoMuted;
      video.muted = state.videoMuted;
      updateVideoButtons();
    });

    qs("#videoSpeedBtn").addEventListener("click", function () {
      var idx = VIDEO_SPEEDS.indexOf(state.videoPlaybackRate);
      state.videoPlaybackRate = VIDEO_SPEEDS[(idx + 1) % VIDEO_SPEEDS.length];
      applyPlaybackRate();
      updateVideoButtons();
    });

    updateVideoButtons();

    video.addEventListener("ended", function () {
      pauseVideo();
    });

    video.addEventListener("timeupdate", function () {
      if (!state.videoPlaying) return;
      var t = video.currentTime;
      state.currentTimestamp = t;
      qs("#timestampInput").value = formatTime(t, { decimals: 1 });
      persistVideoTime(t);
      if (!_playheadRaf) {
        _playheadRaf = requestAnimationFrame(function () {
          _playheadRaf = 0;
          renderPlayhead();
        });
      }
    });
  }

  function persistVideoTime(t) {
    if (!state.selectedParticipant || !isFinite(t)) return;
    var stored = getStoredUIState("screenspace");
    var map = (stored.videoTimeByParticipant && typeof stored.videoTimeByParticipant === "object")
      ? stored.videoTimeByParticipant : {};
    map[state.selectedParticipant] = t;
    setStoredUIStateField("screenspace", "videoTimeByParticipant", map);
  }

  function playVideo() {
    var video = qs("#videoPlayer");
    if (!state.selectedParticipant || !state.videoInfo) return;

    var expectedSrc = videoStreamUrl(state.selectedParticipant);
    if (!video.src || video.src.indexOf(expectedSrc) === -1) {
      video.src = expectedSrc;
    }

    video.currentTime = state.currentTimestamp;
    video.muted = state.videoMuted;

    video.classList.add("active");
    qs("#frameCanvas").classList.add("video-active");

    state.videoPlaying = true;
    updateVideoButtons();

    applyPlaybackRate();
    var playPromise = video.play();
    if (playPromise && playPromise.then) {
      playPromise.catch(function () {
        pauseVideo();
      });
    }
  }

  function pauseVideo() {
    var video = qs("#videoPlayer");
    video.pause();
    state.videoPlaying = false;

    var ts = video.currentTime || state.currentTimestamp;
    state.currentTimestamp = ts;

    video.classList.remove("active");
    qs("#frameCanvas").classList.remove("video-active");

    loadFrame(ts);
    updateVideoButtons();
  }

  function applyPlaybackRate() {
    var v = qs("#videoPlayer");
    v.defaultPlaybackRate = state.videoPlaybackRate;
    v.playbackRate = state.videoPlaybackRate;
    // Disable pitch preservation: the time-stretch filter is CPU-heavy and
    // causes visible judder at >=3x. Audio pitch will rise at high speeds.
    v.preservesPitch = false;
    v.mozPreservesPitch = false;
    v.webkitPreservesPitch = false;
  }

  function updateVideoButtons() {
    var playBtn = qs("#videoPlayBtn");
    var muteBtn = qs("#videoMuteBtn");

    playBtn.innerHTML = "";
    playBtn.appendChild(state.videoPlaying ? iconSpan("pause") : iconSpan("play"));
    playBtn.title = state.videoPlaying ? "Pause (Space)" : "Play/Pause (Space)";

    muteBtn.innerHTML = "";
    muteBtn.appendChild(state.videoMuted ? iconSpan("speaker-x-mark") : iconSpan("speaker-wave"));
    muteBtn.classList.toggle("active", !state.videoMuted);

    var speedBtn = qs("#videoSpeedBtn");
    if (speedBtn) {
      speedBtn.textContent = state.videoPlaybackRate + "x";
      speedBtn.classList.toggle("active", state.videoPlaybackRate !== 1);
    }
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

  // ---- Overlay interaction state machine ----
  //
  // Mousedown picks a mode based on what's under the cursor:
  //   - inside the template overlay        → state.draggingTemplate
  //   - on a region's resize handle        → state.resizingRegion
  //   - inside a region body               → state.draggingRegion
  //   - empty area                         → state.drawingRegion (new region)
  //
  // Mousemove updates whichever mode is active; mouseup commits and clears it.
  // Document-level move/up listeners mirror the canvas handlers so a drag
  // continues smoothly when the cursor leaves the overlay (the canvas-level
  // handlers stop firing, the document ones take over).
  function initRegionDrawing() {
    var overlay = qs("#overlayCanvas");

    function finishDrawingRegion(e) {
      if (!state.drawingRegion) return false;
      var rect = _cachedOverlayRect || overlay.getBoundingClientRect();
      var pos = canvasCoords(overlay, e, rect);
      pos.x = clamp(pos.x, 0, overlay.width);
      pos.y = clamp(pos.y, 0, overlay.height);
      state.drawingRegion.endX = pos.x;
      state.drawingRegion.endY = pos.y;
      var r = normalizeRect(
        state.drawingRegion.startX, state.drawingRegion.startY,
        state.drawingRegion.endX, state.drawingRegion.endY
      );
      state.drawingRegion = null;
      _cachedOverlayRect = null;
      if (r.w > 5 && r.h > 5) {
        state.pendingRegion = r;
      } else if (r.w > 0 || r.h > 0) {
        // Drop the draw silently for click-without-drag (zero size), but
        // surface a hint when the user actually dragged a too-small box —
        // otherwise the picker just snaps closed with no feedback.
        showToast("Region too small — drag a larger area");
      }
      flushOverlayRender();
      updateRegionButtons();
      return true;
    }

    overlay.addEventListener("mousedown", function (e) {
      if (e.button !== 0) return;
      if (state.videoPlaying) pauseVideo();
      if (state.pipetteActive) {
        var frameCanvas = qs("#frameCanvas");
        if (!frameCanvas) { deactivatePipette(); return; }
        var pos = canvasCoords(frameCanvas, e);
        var frameCtx = frameCanvas.getContext("2d");
        var pixel = frameCtx.getImageData(pos.x, pos.y, 1, 1).data;
        var hsv = rgbToHsv(pixel[0], pixel[1], pixel[2]);
        if (state._mtPipetteStep !== undefined && state._mtPipetteStep >= 0) {
          var mtIdx = state._mtPipetteStep;
          var mtSfx = "_mt" + mtIdx;
          var mtHEl = qs("#paramColorH" + mtSfx);
          var mtSEl = qs("#paramColorS" + mtSfx);
          var mtVEl = qs("#paramColorV" + mtSfx);
          if (mtHEl) mtHEl.value = clamp(Math.round(hsv.h), 0, 180);
          if (mtSEl) mtSEl.value = clamp(Math.round(hsv.s), 0, 255);
          if (mtVEl) mtVEl.value = clamp(Math.round(hsv.v), 0, 255);
          var mtHex = qs("#paramColorHex" + mtSfx);
          if (mtHex) mtHex.value = rgbToHex(pixel[0], pixel[1], pixel[2]);
          state._mtPipetteStep = -1;
        } else {
          setTargetColor(hsv.h, hsv.s, hsv.v);
        }
        deactivatePipette();
        showToast("Sampled color from frame");
        return;
      }
      _cachedOverlayRect = overlay.getBoundingClientRect();
      pos = canvasCoords(overlay, e, _cachedOverlayRect);
      var displayW = _cachedOverlayRect.width || overlay.width;
      // Bail if the overlay has no usable size (e.g. hidden / display:none
      // container). Without this, `s` becomes NaN and propagates into hit
      // testing and any region coords stored downstream.
      if (!displayW || !overlay.width) return;
      var s = overlay.width / displayW;
      var ctx = overlay.getContext("2d");
      var tHit = templateOverlayBounds();
      if (tHit
          && pos.x >= tHit.x && pos.x <= tHit.x + tHit.w
          && pos.y >= tHit.y && pos.y <= tHit.y + tHit.h) {
        state.draggingTemplate = { offsetX: pos.x - tHit.x, offsetY: pos.y - tHit.y };
        document.body.style.cursor = "grabbing";
        document.body.style.userSelect = "none";
        scheduleOverlayRender();
        return;
      }
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
      if (state.draggingTemplate) {
        var tImg = state.uploadedTemplateImg;
        var tw = Math.max(1, Math.round(tImg.naturalWidth * (state.templateScalePreview || 1.0)));
        var thh = Math.max(1, Math.round(tImg.naturalHeight * (state.templateScalePreview || 1.0)));
        var nx = clamp(pos.x - state.draggingTemplate.offsetX, 0, overlay.width - tw);
        var ny = clamp(pos.y - state.draggingTemplate.offsetY, 0, overlay.height - thh);
        state.templateOverlayPos = { x: nx, y: ny };
        scheduleOverlayRender();
        return;
      }
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
      var tBoundsHover = templateOverlayBounds();
      var overTemplate = tBoundsHover
        && pos.x >= tBoundsHover.x && pos.x <= tBoundsHover.x + tBoundsHover.w
        && pos.y >= tBoundsHover.y && pos.y <= tBoundsHover.y + tBoundsHover.h;
      var hit = overTemplate ? null : findHitRegion(pos.x, pos.y, s, ctx);
      state.hoveredRegion = hit;
      if (overTemplate) {
        overlay.style.cursor = "grab";
      } else {
        overlay.style.cursor = hit ? (hit.handle === "resize" ? "nwse-resize" : "grab") : "crosshair";
      }
      scheduleOverlayRender();
    });

    overlay.addEventListener("mouseup", function (e) {
      if (state.pipetteActive) return;
      if (state.draggingTemplate) {
        _cachedOverlayRect = null;
        state.draggingTemplate = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        flushOverlayRender();
        return;
      }
      if (state.resizingRegion) {
        _cachedOverlayRect = null;
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
        _cachedOverlayRect = null;
        var dName = state.draggingRegion.name;
        state.draggingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(dName);
        flushOverlayRender();
        updateRegionButtons();
        return;
      }
      finishDrawingRegion(e);
    });

    // Document-level listeners so drag/resize continues outside the canvas
    document.addEventListener("mousemove", function (e) {
      if (!state.resizingRegion && !state.draggingRegion && !state.draggingTemplate) return;
      var rect = _cachedOverlayRect || overlay.getBoundingClientRect();
      var pos = canvasCoords(overlay, e, rect);
      var displayW = rect.width || overlay.width;
      var s = overlay.width / displayW;
      if (state.draggingTemplate) {
        var tImg = state.uploadedTemplateImg;
        var tw = Math.max(1, Math.round(tImg.naturalWidth * (state.templateScalePreview || 1.0)));
        var thh = Math.max(1, Math.round(tImg.naturalHeight * (state.templateScalePreview || 1.0)));
        var nx = clamp(pos.x - state.draggingTemplate.offsetX, 0, overlay.width - tw);
        var ny = clamp(pos.y - state.draggingTemplate.offsetY, 0, overlay.height - thh);
        state.templateOverlayPos = { x: nx, y: ny };
        scheduleOverlayRender();
        return;
      }
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

    document.addEventListener("mouseup", function (e) {
      if (finishDrawingRegion(e)) return;
      if (state.draggingTemplate) {
        _cachedOverlayRect = null;
        state.draggingTemplate = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        flushOverlayRender();
        return;
      }
      if (state.resizingRegion) {
        _cachedOverlayRect = null;
        var rName = state.resizingRegion.name;
        state.resizingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(rName);
        flushOverlayRender();
        updateRegionButtons();
      } else if (state.draggingRegion) {
        _cachedOverlayRect = null;
        var dName = state.draggingRegion.name;
        state.draggingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(dName);
        flushOverlayRender();
        updateRegionButtons();
      } else {
        _cachedOverlayRect = null;
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
    toggleLabelsBtn.appendChild(iconSpan("tag"));
    toggleLabelsBtn.addEventListener("click", function () {
      state.showRegionLabels = !state.showRegionLabels;
      updateRegionButtons();
      renderOverlay();
    });

    // Toggle region visibility
    var toggleRegionsBtn = qs("#toggleRegionsBtn");
    toggleRegionsBtn.appendChild(iconSpan("eye"));
    toggleRegionsBtn.addEventListener("click", function () {
      state.showRegionOverlays = !state.showRegionOverlays;
      updateRegionButtons();
      renderOverlay();
    });

    // Stash all regions
    var stashBtn = qs("#stashRegionsBtn");
    stashBtn.appendChild(iconSpan("archive-box-arrow-down"));
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

  var _regionNameModalPrevFocus = null;

  function showRegionNameModal() {
    var r = state.pendingRegion;
    _regionNameModalPrevFocus = document.activeElement;
    qs("#regionNameInput").value = "";
    qs("#regionDescInput").value = "";
    qs("#regionCoords").textContent = r ? (r.x + ", " + r.y + " \u2014 " + r.w + "\u00d7" + r.h + " px") : "";
    qs("#regionNameModal").classList.remove("hidden");
    qs("#regionNameInput").focus();
  }

  function hideRegionNameModal() {
    qs("#regionNameModal").classList.add("hidden");
    if (_regionNameModalPrevFocus && typeof _regionNameModalPrevFocus.focus === "function") {
      try { _regionNameModalPrevFocus.focus(); } catch (_) {}
    }
    _regionNameModalPrevFocus = null;
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
    toggleRegionsBtn.appendChild(state.showRegionOverlays ? iconSpan("eye") : iconSpan("eye-slash"));
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
    refreshModelView({ debounce: true });
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
      restoreBtn.appendChild(iconSpan("arrow-up-tray"));
      restoreBtn.addEventListener("click", function () { restoreStash(stash.id); });
      actions.appendChild(restoreBtn);

      var dismissBtn = el("button", "stash-card-action-btn");
      dismissBtn.title = "Dismiss stash";
      dismissBtn.appendChild(iconSpan("x-mark"));
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

  function templateOverlayBounds() {
    if (state.activeWorkflow !== "template") return null;
    var tImg = state.uploadedTemplateImg;
    if (!tImg || !tImg.naturalWidth) return null;
    var canvas = qs("#overlayCanvas");
    if (!canvas || !canvas.width) return null;
    var scale = state.templateScalePreview || 1.0;
    var w = Math.max(1, Math.round(tImg.naturalWidth * scale));
    var h = Math.max(1, Math.round(tImg.naturalHeight * scale));
    var x, y;
    if (state.templateOverlayPos) {
      x = state.templateOverlayPos.x;
      y = state.templateOverlayPos.y;
    } else {
      var regs = state.previewRegions || state.regions;
      var activeR = state.activeRegion && regs[state.activeRegion];
      if (activeR) {
        var aPx = regionToPixels(activeR);
        x = aPx.x;
        y = aPx.y;
      } else {
        var displayW = canvas.getBoundingClientRect().width || canvas.width;
        var s = canvas.width / displayW;
        x = Math.round(10 * s);
        y = Math.round(10 * s);
      }
    }
    x = Math.max(0, Math.min(canvas.width - w, x));
    y = Math.max(0, Math.min(canvas.height - h, y));
    return { x: x, y: y, w: w, h: h };
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

    // Template preview: overlay the uploaded PNG at its effective in-video
    // size (native PNG pixels * template_scale) so the user can see how
    // large it will match. Canvas is sized in native video pixels, so the
    // display-to-video ratio is applied automatically when the browser
    // scales the canvas to fit the viewport. The overlay is draggable so
    // the user can position it against specific elements in the frame.
    var tBounds = templateOverlayBounds();
    if (tBounds) {
      var tImg = state.uploadedTemplateImg;
      ctx.globalAlpha = state.draggingTemplate ? 0.9 : 0.75;
      ctx.drawImage(tImg, tBounds.x, tBounds.y, tBounds.w, tBounds.h);
      ctx.globalAlpha = 1.0;
      ctx.strokeStyle = taskTypeColor("template");
      ctx.lineWidth = (state.draggingTemplate ? 2 : 1.5) * s;
      ctx.setLineDash([4 * s, 3 * s]);
      ctx.strokeRect(tBounds.x, tBounds.y, tBounds.w, tBounds.h);
      ctx.setLineDash([]);
      ctx.font = Math.round(11 * s) + "px " + getThemeColors().fontMono;
      ctx.fillStyle = taskTypeColor("template");
      ctx.fillText(tBounds.w + "\u00d7" + tBounds.h + " px",
        tBounds.x + Math.round(4 * s),
        tBounds.y + tBounds.h + Math.round(14 * s));
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
        var rPx = hm.region_coords;
        if (rPx && rPx.w) {
          ctx.drawImage(hm._img, rPx.x, rPx.y, rPx.w, rPx.h);
        }
      }
      ctx.globalAlpha = 1.0;
    }

    // Model-view overlay (toggle or held-key blink comparator)
    var overlayActive = (state.overlayEnabled || state.overlayBlinkActive)
      && state.overlayImage
      && _overlayEligibleForActiveTool();
    if (overlayActive) {
      ctx.globalAlpha = state.overlayBlinkActive ? 1.0 : 0.7;
      var scope = state.overlayImageScope || "region";
      if (scope === "frame") {
        ctx.drawImage(state.overlayImage, 0, 0, canvas.width, canvas.height);
      } else {
        var oRegion = state.pendingRegion
          ? state.pendingRegion
          : (state.activeRegion ? state.regions[state.activeRegion] : null);
        if (oRegion) {
          var oPx = regionToPixels(oRegion);
          if (oPx && oPx.w && oPx.h) {
            ctx.drawImage(state.overlayImage, oPx.x, oPx.y, oPx.w, oPx.h);
          }
        } else {
          // No region active — overlay covers the whole frame.
          ctx.drawImage(state.overlayImage, 0, 0, canvas.width, canvas.height);
        }
      }
      ctx.globalAlpha = 1.0;
    }
  }

  // ---- Timeline ----
  // initTimeline wires pointer events; renderTimeline paints ruler, markers,
  // optional amplitude band, and hit rects used by hover/click.

  function initTimeline() {
    qs("#zoomInBtn").appendChild(iconSpan("plus", "ss-icon--sm"));
    qs("#zoomOutBtn").appendChild(iconSpan("minus", "ss-icon--sm"));

    var storedTimeline = getStoredUIState("screenspace");
    if (storedTimeline.amplitudeGraphEnabled === true) {
      state.amplitudeGraphEnabled = true;
      qs("#amplitudeGraphBtn").classList.add("active");
    }
    var canvas = qs("#timelineCanvas");
    sizeTimelineCanvas();
    window.addEventListener("resize", function () {
      _cachedOverlayRect = null;
      _cachedTimelineRect = null;
      sizeTimelineCanvas();
    });
    window.addEventListener("scroll", function () {
      _cachedOverlayRect = null;
      _cachedTimelineRect = null;
    }, true);

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
    qs("#amplitudeGraphBtn").addEventListener("click", function () {
      state.amplitudeGraphEnabled = !state.amplitudeGraphEnabled;
      this.classList.toggle("active", state.amplitudeGraphEnabled);
      setStoredUIStateField("screenspace", "amplitudeGraphEnabled", state.amplitudeGraphEnabled);
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
    _cachedTimelineRect = null;
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
    var rect = getTimelineRect(canvas);
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
    if (state.inMarker !== null) parts.push("In: " + formatTime(state.inMarker, { decimals: 1 }));
    if (state.outMarker !== null) parts.push("Out: " + formatTime(state.outMarker, { decimals: 1 }));
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

    // In/Out marker shading — scrim "outside" the active range against the
    // timeline's surfaceAlt background. fg works in both themes (fg is white in
    // dark → lightens, dark in light → darkens; both differentiate the range).
    if (state.inMarker !== null || state.outMarker !== null) {
      ctx.fillStyle = hexToRgba(tc.fg, 0.12);
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

    // Optional amplitude band: per-task-type event-density curves above the
    // result markers. Drawn before markers so markers stay on top visually.
    var ampOn = state.amplitudeGraphEnabled;
    var AMP_BAND_H = 22;
    var AMP_BAND_GAP = 2;
    var resultY = ampOn ? 24 + AMP_BAND_H + AMP_BAND_GAP : 24;
    var resultH = h - resultY - 6;
    var focused = focusedTaskId();
    _timelineHitRects = [];

    if (ampOn) {
      var seriesByType = {};
      state.tasks.forEach(function (task) {
        if (!task.result || task.status === "cancelled") return;
        if (task.participant && task.participant !== state.selectedParticipant) return;
        if (task.type === "timelapse") return;
        if (!seriesByType[task.type]) {
          seriesByType[task.type] = { key: task.type, color: taskTypeColor(task.type), timestamps: [] };
        }
        var dst = seriesByType[task.type].timestamps;
        var results = task.result || [];
        for (var ri = 0; ri < results.length; ri++) {
          var r = results[ri];
          var ts = r.timestamp !== undefined ? r.timestamp : r.start;
          if (ts !== undefined) dst.push(ts);
        }
      });
      var seriesList = Object.keys(seriesByType).map(function (k) { return seriesByType[k]; });
      var focusedTask = focused ? findTask(focused) : null;
      var dimKey = focusedTask ? focusedTask.type : null;
      drawAmplitudeBands(ctx, {
        x: 0,
        y: 24,
        w: w,
        h: AMP_BAND_H,
        visStart: visStart,
        visEnd: visEnd,
        series: seriesList,
        binPx: 2,
        dimKey: dimKey,
      });
    }

    // Build excluded-timestamp lookup per task from cached events
    var excludedByTask = {};
    Object.keys(state.taskEvents).forEach(function (tid) {
      var exSet = {};
      (state.taskEvents[tid] || []).forEach(function (ev) {
        if (ev.excluded) exSet[ev.time_in.toFixed(2)] = true;
      });
      excludedByTask[tid] = exSet;
    });

    state.tasks.forEach(function (task) {
      if (!task.result || task.status === "cancelled") return;
      if (task.participant && task.participant !== state.selectedParticipant) return;
      var color = taskTypeColor(task.type);
      var dimmed = focused && task.id !== focused;
      var taskExcluded = excludedByTask[task.id] || {};
      if ((task.type === "color" || task.type === "inactivity") && task.status === "completed") {
        // Completed color: merged spans
        task.result.forEach(function (span) {
          var isExcluded = taskExcluded[span.start.toFixed(2)];
          ctx.fillStyle = hexToRgba(color, isExcluded ? 0.05 : (dimmed ? 0.10 : 0.35));
          var x1 = timeToX(span.start);
          var x2 = timeToX(span.end);
          var rw = Math.max(x2 - x1, 2);
          ctx.fillRect(x1, resultY, rw, resultH);
          _timelineHitRects.push({ x1: x1, x2: x1 + rw, y: resultY, h: resultH, task: task, result: span });
        });
      } else if (task.type === "timelapse") {
        // No timeline markers for timelapse
      } else {
        // Point markers (change, similarity, text, numbers, template, flow, scene, running color)
        ctx.lineWidth = 1.5;
        var results = task.result || [];
        results.forEach(function (r) {
          var ts = r.timestamp !== undefined ? r.timestamp : r.start;
          if (ts === undefined) return;
          var isExcluded = taskExcluded[ts.toFixed(2)];
          var sceneDimmed = task.type === "scene" && state.hoveredResultSceneName !== null
            && r.scene_name !== state.hoveredResultSceneName;
          if (isExcluded) {
            ctx.strokeStyle = hexToRgba(color, 0.15);
            ctx.setLineDash([3, 3]);
          } else if (dimmed || sceneDimmed) {
            ctx.strokeStyle = hexToRgba(color, 0.15);
            ctx.setLineDash([]);
          } else {
            ctx.strokeStyle = color;
            ctx.setLineDash([]);
          }
          var x = timeToX(ts);
          ctx.beginPath();
          ctx.moveTo(x, resultY);
          ctx.lineTo(x, resultY + resultH);
          ctx.stroke();
          ctx.setLineDash([]);
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
    var rect = getTimelineRect(canvas);
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
    var icon = buildTypeIcon(hit.task.type);
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
      timeStr = formatTime(r.start, { decimals: 1 }) + " \u2013 " + formatTime(r.end, { decimals: 1 });
    } else {
      var ts = r.timestamp !== undefined ? r.timestamp : r.start;
      timeStr = formatTime(ts, { decimals: 1 });
    }
    tip.appendChild(el("span", "ss-tooltip-time", timeStr));

    var details = el("div", "ss-tooltip-details");
    details.appendChild(el("span", "", hit.task.participant + " \u00b7 " + (hit.task.region || "")));
    if (hit.task.type === "inactivity" && r.duration !== undefined) {
      details.appendChild(el("span", "", "Duration: " + r.duration.toFixed(1) + "s"));
      if (r.avg_distance !== undefined) details.appendChild(el("span", "", "Avg distance: " + r.avg_distance));
    } else if (hit.task.type === "color" && r.duration !== undefined) {
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

  // ---- Tool info tooltip ----

  var TOOL_INFO = {
    multitool: "Multitool chains multiple analysis tools together. Add at least two tool steps — each subsequent tool only checks frames that passed the previous step, finding moments that match ALL criteria simultaneously.",
    color: "Color tool: finds frames where a region's average color matches your chosen target within the tolerance range. Useful for tracking UI state indicators, health bars, or any element identified by a specific color.",
    change: "Change tool: detects frames where pixel differences in the region exceed the threshold. Good for spotting sudden visual changes like screen transitions, pop-ups appearing, or loading states completing.",
    similarity: "Similarity tool: compares each frame against a captured reference using structural similarity (SSIM). Use it to find moments that look like a specific reference frame — e.g. a particular menu, dialog, or game state.",
    text: "Text tool: performs OCR on the region and fuzzy-matches against your search text. Useful for detecting when specific labels, error messages, or button text appear on screen.",
    numbers: "Numbers tool: reads numeric values from the region via OCR and compares them against your target using the selected operator. Great for monitoring scores, timers, counters, or any on-screen number.",
    timelapse: "Timelapse tool: generates a sped-up video or GIF of the region over the selected time range. Unlike other tools, this produces a single artifact rather than detecting individual frames.",
    template: "Template tool: searches the full video frame for a captured or uploaded reference image using template matching. Works across the entire frame, not just the selected region — ideal for finding icons, buttons, or UI elements wherever they appear.",
    flow: "Flow tool: detects motion in the region via dense optical flow. Higher magnitude thresholds filter out subtle movements. Useful for detecting player movement, animations starting, or activity in a specific area.",
    scene: "Scene tool: classifies each frame by comparing it to your captured reference scenes. Useful for building a timeline of when different screens, menus, or game levels are active.",
    inactivity: "Inactivity tool: detects spans of near-duplicate frames in a region using perceptual hashing. Surfaces loading screens, frozen states, or repeated animation loops. Set the minimum duration to filter out brief pauses."
  };

  var _toolInfoPinned = false;

  function showToolInfoTooltip(anchorEl) {
    var tip = qs("#toolInfoTooltip");
    if (!tip) return;
    var type = state.activeWorkflow;
    var text = TOOL_INFO[type] || "";
    tip.innerHTML = "";

    var header = el("div", "tool-info-header");
    header.appendChild(el("strong", null, type.charAt(0).toUpperCase() + type.slice(1)));
    var closeBtn = el("button", "tool-info-close hidden");
    closeBtn.appendChild(iconSpan("x-mark"));
    closeBtn.addEventListener("click", function () {
      hideToolInfoTooltip(true);
    });
    header.appendChild(closeBtn);
    tip.appendChild(header);

    tip.appendChild(el("p", "tool-info-body", text));

    tip.classList.remove("hidden");
    positionToolInfoTooltip(tip, anchorEl);
  }

  function positionToolInfoTooltip(tip, anchorEl) {
    var rect = anchorEl.getBoundingClientRect();
    var x = rect.left;
    var y = rect.bottom + 6;
    tip.style.left = x + "px";
    tip.style.top = y + "px";
    var tipRect = tip.getBoundingClientRect();
    if (tipRect.right > window.innerWidth - 8) {
      tip.style.left = (window.innerWidth - tipRect.width - 8) + "px";
    }
    if (tipRect.bottom > window.innerHeight - 8) {
      tip.style.top = (rect.top - tipRect.height - 6) + "px";
    }
  }

  function pinToolInfoTooltip() {
    _toolInfoPinned = true;
    var tip = qs("#toolInfoTooltip");
    if (!tip) return;
    var closeBtn = tip.querySelector(".tool-info-close");
    if (closeBtn) closeBtn.classList.remove("hidden");
  }

  function hideToolInfoTooltip(force) {
    if (_toolInfoPinned && !force) return;
    _toolInfoPinned = false;
    var tip = qs("#toolInfoTooltip");
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
        hideToolInfoTooltip(true);
        state.activeWorkflow = tab.dataset.type;
        setStoredUIStateField("screenspace", "activeWorkflow", state.activeWorkflow);
        qsa(".wf-tab").forEach(function (t) { t.classList.remove("active"); });
        tab.classList.add("active");
        renderWorkflowParams();
        updateRunButton();
      });
    });
    // Tool info icon in action row
    var infoBtn = qs("#toolInfoBtn");
    if (infoBtn) {
      infoBtn.addEventListener("mouseenter", function () {
        if (!_toolInfoPinned) showToolInfoTooltip(infoBtn);
      });
      infoBtn.addEventListener("mouseleave", function () {
        hideToolInfoTooltip(false);
      });
      infoBtn.addEventListener("click", function (e) {
        e.stopPropagation();
        if (_toolInfoPinned) {
          hideToolInfoTooltip(true);
        } else {
          showToolInfoTooltip(infoBtn);
          pinToolInfoTooltip();
        }
      });
    }
    renderWorkflowParams();
  }

  function renderIntervalSlot(inputId, min, max, def, step) {
    var slot = qs("#workflowIntervalSlot");
    if (!slot) return;
    slot.innerHTML = "";
    slot.setAttribute("data-tooltip", "Interval (seconds)");
    var iconWrap = el("div", "interval-icon");
    var iconMask = el("span", "interval-icon-mask");
    applyMaskIcon(iconMask, 'url("/screenspace/icons/clock.svg")');
    iconWrap.appendChild(iconMask);
    slot.appendChild(iconWrap);
    var ctrl = el("div", "param-control");
    ctrl.appendChild(numberInput(inputId, min, max, def, step));
    slot.appendChild(ctrl);
  }

  // ---- Multitool step list (drag reorder + drop-to-import from task queue) ----

  var MULTITOOL_ALLOWED_TYPES = [
    { value: "color", label: "Color" },
    { value: "change", label: "Change" },
    { value: "similarity", label: "Similarity" },
    { value: "text", label: "Text" },
    { value: "numbers", label: "Numbers" },
    { value: "template", label: "Template" },
    { value: "flow", label: "Flow" },
    { value: "scene", label: "Scene" },
    { value: "inactivity", label: "Inactivity" },
  ];

  function renderMultitoolStepBody(body, stepType, idx) {
    var sfx = "_mt" + idx;

    // Per-step region selector
    var regionRow = el("div", "param-row");
    regionRow.appendChild(el("span", "param-label", "Region"));
    var regionCtrl = el("div", "param-control");
    var regionSel = document.createElement("select");
    regionSel.className = "multitool-step-region-select";
    regionSel.id = "paramStepRegion" + sfx;
    var regionRefs = allAvailableRegionRefs();
    var selectedRef = normalizeRegionRef(state.multitoolSteps[idx].region_ref)
      || (state.multitoolSteps[idx].region ? activeRegionRef(state.multitoolSteps[idx].region) : null);
    var selectedKey = selectedRef ? regionRefKey(selectedRef) : "";
    regionRefs.forEach(function (ref) {
      var opt = document.createElement("option");
      opt.value = regionRefKey(ref);
      opt.textContent = regionRefLabel(ref);
      if (opt.value === selectedKey) opt.selected = true;
      regionSel.appendChild(opt);
    });
    if (!state.multitoolSteps[idx].region && regionRefs.length > 0) {
      var defaultRegionRef = state.runRegions.length > 0 ? normalizeRegionRef(state.runRegions[0]) : null;
      defaultRegionRef = defaultRegionRef || (state.activeRegion ? activeRegionRef(state.activeRegion) : null) || regionRefs[0];
      state.multitoolSteps[idx].region = defaultRegionRef.name;
      state.multitoolSteps[idx].region_ref = regionRefPayload(defaultRegionRef);
      regionSel.value = regionRefKey(defaultRegionRef);
    }
    (function (capturedIdx) {
      regionSel.addEventListener("change", function () {
        var ref = availableRegionRefByKey(regionSel.value);
        state.multitoolSteps[capturedIdx].region = ref ? ref.name : "";
        state.multitoolSteps[capturedIdx].region_ref = ref ? regionRefPayload(ref) : null;
      });
    })(idx);
    regionCtrl.appendChild(regionSel);
    regionRow.appendChild(regionCtrl);
    body.appendChild(regionRow);

    var renderer = MULTITOOL_PARAM_RENDERERS[stepType];
    if (renderer) renderer(body, idx, sfx);
    // _initial is consumed at first render; drop it so adding/removing other steps later
    // doesn't overwrite the user's in-progress edits with the original saved values.
    delete state.multitoolSteps[idx]._initial;
  }

  function _mtRenderColor(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    var initColor = init.target_color || {};
    var initH = numberOrDefault(initColor.h, 0);
    var initS = numberOrDefault(initColor.s, 0);
    var initV = numberOrDefault(initColor.v, 0);
    var initTol = init.tolerance ? Math.round(init.tolerance.h * 100 / 90) : 30;
    var row1 = el("div", "param-row");
    row1.appendChild(el("span", "param-label", "Hex color"));
    var ctrl1 = el("div", "param-control");
    var hexIn = document.createElement("input");
    hexIn.type = "text"; hexIn.id = "paramColorHex" + sfx; hexIn.autocomplete = "off";
    hexIn.className = "color-hex-input"; hexIn.placeholder = "#000000"; hexIn.maxLength = 7;
    hexIn.style.width = "5.5rem";
    var initRgb = hsvToRgb(initH, initS, initV);
    hexIn.value = rgbToHex(initRgb.r, initRgb.g, initRgb.b);
    ctrl1.appendChild(hexIn);
    var pipBtn = el("button", "btn btn-small btn-pipette");
    pipBtn.title = "Pick color from frame";
    pipBtn.appendChild(buildTypeIcon("color"));
    pipBtn.addEventListener("click", function () {
      if (state.pipetteActive) { deactivatePipette(); return; }
      state._mtPipetteStep = idx;
      activatePipette();
    });
    ctrl1.appendChild(pipBtn);
    var hH = document.createElement("input"); hH.type = "hidden"; hH.id = "paramColorH" + sfx; hH.value = String(initH);
    var hS = document.createElement("input"); hS.type = "hidden"; hS.id = "paramColorS" + sfx; hS.value = String(initS);
    var hV = document.createElement("input"); hV.type = "hidden"; hV.id = "paramColorV" + sfx; hV.value = String(initV);
    ctrl1.appendChild(hH); ctrl1.appendChild(hS); ctrl1.appendChild(hV);
    row1.appendChild(ctrl1);
    body.appendChild(row1);
    hexIn.addEventListener("input", function () {
      var hex = hexIn.value.replace("#", "");
      if (hex.length === 6) {
        var r = parseInt(hex.substring(0, 2), 16) || 0;
        var g = parseInt(hex.substring(2, 4), 16) || 0;
        var b = parseInt(hex.substring(4, 6), 16) || 0;
        // OpenCV HSV: H 0-180, S 0-255, V 0-255
        var rr = r / 255, gg = g / 255, bb = b / 255;
        var mx = Math.max(rr, gg, bb), mn = Math.min(rr, gg, bb);
        var d = mx - mn, h = 0, s = mx === 0 ? 0 : d / mx, v = mx;
        if (d !== 0) {
          if (mx === rr) h = ((gg - bb) / d + (gg < bb ? 6 : 0)) / 6;
          else if (mx === gg) h = ((bb - rr) / d + 2) / 6;
          else h = ((rr - gg) / d + 4) / 6;
        }
        hH.value = Math.round(h * 180);
        hS.value = Math.round(s * 255);
        hV.value = Math.round(v * 255);
      }
    });
    _mtAddNumberRow(body, "Tolerance", "paramColorTol" + sfx, 0, 100, initTol, 1);
  }

  function _mtRenderChange(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    _mtAddNumberRow(body, "Threshold", "paramChangeThresh" + sfx, 0.01, 0.50, numberOrDefault(init.threshold, 0.03), 0.01);
    _mtAddNumberRow(body, "Noise", "paramChangeNoise" + sfx, 0, 100, intOrDefault(init.noise_threshold, 30), 1);
  }

  function _mtRenderSimilarity(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    _mtAddCaptureRefRow(body, idx, "Reference", "paramSimRef" + sfx);
    _mtAddNumberRow(body, "Threshold", "paramSimThresh" + sfx, 0.50, 1.00, numberOrDefault(init.threshold, 0.90), 0.01);
  }

  function _mtRenderText(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    var r1 = el("div", "param-row");
    r1.appendChild(el("span", "param-label", "Search"));
    var c1 = el("div", "param-control");
    var searchIn = document.createElement("input");
    searchIn.type = "text"; searchIn.id = "paramTextSearch" + sfx; searchIn.autocomplete = "off";
    searchIn.placeholder = "Search text...";
    searchIn.value = init.search_string || "";
    searchIn.style.flex = "1"; searchIn.style.fontSize = "var(--text-xs)";
    searchIn.style.padding = "var(--space-1)";
    searchIn.style.border = "1px solid var(--color-border)";
    searchIn.style.borderRadius = "var(--radius-sm)";
    searchIn.style.background = "var(--color-bg)";
    searchIn.style.color = "var(--color-text)";
    c1.appendChild(searchIn);
    r1.appendChild(c1);
    body.appendChild(r1);
    _mtAddNumberRow(body, "Fuzzy", "paramTextFuzzy" + sfx, 0.50, 1.00, numberOrDefault(init.fuzzy_threshold, 0.80), 0.01);
  }

  function _mtRenderNumbers(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    var r1 = el("div", "param-row");
    r1.appendChild(el("span", "param-label", "Operator"));
    var c1 = el("div", "param-control");
    var sel = document.createElement("select");
    sel.id = "paramNumOperator" + sfx;
    sel.style.fontSize = "var(--text-xs)";
    [["gt",">"],["lt","<"],["eq","="],["gte","≥"],["lte","≤"],["range","range"]].forEach(function (pair) {
      var opt = document.createElement("option"); opt.value = pair[0]; opt.textContent = pair[1];
      sel.appendChild(opt);
    });
    sel.value = init.operator || "gt";
    c1.appendChild(sel);
    r1.appendChild(c1);
    body.appendChild(r1);
    var r2 = el("div", "param-row");
    r2.appendChild(el("span", "param-label", "Target"));
    var c2 = el("div", "param-control");
    var targetIn = document.createElement("input");
    targetIn.type = "number"; targetIn.id = "paramNumTarget" + sfx; targetIn.style.width = "4rem";
    targetIn.style.fontSize = "var(--text-xs)";
    targetIn.value = numberOrDefault(init.target_value, 0);
    c2.appendChild(targetIn);
    r2.appendChild(c2);
    body.appendChild(r2);
  }

  function _mtRenderTemplate(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    _mtAddCaptureRefRow(body, idx, "Template", "paramTemplateRef" + sfx);
    _mtAddNumberRow(body, "Threshold", "paramTemplateThresh" + sfx, 0.50, 1.00, numberOrDefault(init.threshold, 0.70), 0.01);
  }

  function _mtRenderFlow(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    _mtAddNumberRow(body, "Magnitude", "paramFlowMag" + sfx, 0.5, 20, numberOrDefault(init.magnitude_threshold, 2.0), 0.5);
  }

  function _mtRenderScene(body, idx, sfx) {
    var sceneList = el("div", "scene-reference-list");
    sceneList.id = "mtSceneList" + sfx;
    if (!state.multitoolSteps[idx]._scenes) state.multitoolSteps[idx]._scenes = [];
    function renderMtScenes() {
      sceneList.innerHTML = "";
      state.multitoolSteps[idx]._scenes.forEach(function (ref, refIdx) {
        if (ref.threshold === undefined) ref.threshold = 0.75;
        var item = el("div", "scene-ref-item");
        item.appendChild(el("span", "scene-ref-name", ref.name));
        item.appendChild(el("span", "param-value", formatTime(ref.timestamp, { decimals: 1 })));
        var threshSlider = document.createElement("input");
        threshSlider.type = "range";
        threshSlider.min = "0.50"; threshSlider.max = "1.00"; threshSlider.step = "0.01";
        threshSlider.value = String(ref.threshold);
        threshSlider.className = "scene-ref-thresh";
        var threshVal = el("span", "param-value", String(ref.threshold));
        threshSlider.addEventListener("input", (function (ri) {
          return function () {
            state.multitoolSteps[idx]._scenes[ri].threshold = parseFloat(threshSlider.value);
            threshVal.textContent = threshSlider.value;
          };
        })(refIdx));
        item.appendChild(threshSlider);
        item.appendChild(threshVal);
        var rmBtn = el("button", "btn btn-small", "\u00d7");
        rmBtn.addEventListener("click", (function (ri) {
          return function () {
            state.multitoolSteps[idx]._scenes.splice(ri, 1);
            renderMtScenes();
          };
        })(refIdx));
        item.appendChild(rmBtn);
        sceneList.appendChild(item);
      });
    }
    renderMtScenes();
    body.appendChild(sceneList);
    var addScRow = el("div", "param-row");
    addScRow.appendChild(el("span", "param-label", "Add Scene"));
    var addScCtrl = el("div", "param-control");
    var scNameInp = textInput("paramSceneName" + sfx, "e.g. menu, gameplay");
    addScCtrl.appendChild(scNameInp);
    var scCapBtn = el("button", "btn btn-small", "Capture");
    scCapBtn.addEventListener("click", function () {
      var name = scNameInp.value.trim();
      if (!name) { showToast("Enter a scene name"); return; }
      state.multitoolSteps[idx]._scenes.push({
        name: name, timestamp: state.currentTimestamp, threshold: 0.75,
      });
      scNameInp.value = "";
      renderMtScenes();
      showToast("Scene '" + name + "' at " + formatTime(state.currentTimestamp, { decimals: 1 }));
    });
    addScCtrl.appendChild(scCapBtn);
    addScRow.appendChild(addScCtrl);
    body.appendChild(addScRow);
  }

  function _mtRenderInactivity(body, idx, sfx) {
    var init = state.multitoolSteps[idx]._initial || {};
    var r1 = el("div", "param-row");
    r1.appendChild(el("span", "param-label", "Sensitivity"));
    var c1 = el("div", "param-control");
    var inactSlider = rangeInput("paramInactThresh" + sfx, 0, 30, intOrDefault(init.threshold, 10), 1);
    c1.appendChild(inactSlider);
    var inactVal = el("span", "param-value");
    inactVal.textContent = inactSlider.value;
    inactSlider.addEventListener("input", function () { inactVal.textContent = inactSlider.value; });
    c1.appendChild(inactVal);
    r1.appendChild(c1);
    body.appendChild(r1);
  }

  // Helpers used by multiple per-type renderers above.
  function _mtAddNumberRow(body, label, id, min, max, def, step) {
    var r = el("div", "param-row");
    r.appendChild(el("span", "param-label", label));
    var c = el("div", "param-control");
    c.appendChild(numberInput(id, min, max, def, step));
    r.appendChild(c);
    body.appendChild(r);
  }
  function _mtAddCaptureRefRow(body, idx, label, tsLabelId) {
    var r = el("div", "param-row");
    r.appendChild(el("span", "param-label", label));
    var c = el("div", "param-control");
    var capBtn = el("button", "btn btn-small", "Capture Frame");
    var tsLabel = el("span", "param-value", "\u2014");
    tsLabel.id = tsLabelId;
    if (state.multitoolSteps[idx]._refTs !== undefined) {
      tsLabel.textContent = formatTime(state.multitoolSteps[idx]._refTs, { decimals: 1 });
    }
    capBtn.addEventListener("click", function () {
      state.multitoolSteps[idx]._refTs = state.currentTimestamp;
      tsLabel.textContent = formatTime(state.currentTimestamp, { decimals: 1 });
    });
    c.appendChild(capBtn);
    c.appendChild(tsLabel);
    r.appendChild(c);
    body.appendChild(r);
  }

  var MULTITOOL_PARAM_RENDERERS = {
    color:      _mtRenderColor,
    change:     _mtRenderChange,
    similarity: _mtRenderSimilarity,
    text:       _mtRenderText,
    numbers:    _mtRenderNumbers,
    template:   _mtRenderTemplate,
    flow:       _mtRenderFlow,
    scene:      _mtRenderScene,
    inactivity: _mtRenderInactivity,
  };

  var _multitoolDragMidpoints = null;
  var _multitoolDragOverRaf = null;
  var _multitoolPendingDragOver = null;

  function _cacheMultitoolDragMidpoints(container) {
    var cards = container.querySelectorAll(".multitool-step:not(.dragging)");
    var mids = new Array(cards.length);
    for (var i = 0; i < cards.length; i++) {
      var r = cards[i].getBoundingClientRect();
      mids[i] = r.top + r.height / 2;
    }
    _multitoolDragMidpoints = mids;
  }

  function getMultitoolDropIndex(container, clientY) {
    var mids = _multitoolDragMidpoints;
    if (!mids) {
      _cacheMultitoolDragMidpoints(container);
      mids = _multitoolDragMidpoints;
    }
    for (var i = 0; i < mids.length; i++) {
      if (clientY < mids[i]) return i;
    }
    return mids.length;
  }

  function clearMultitoolDragIndicators(container) {
    var cards = container.querySelectorAll(".multitool-step.drag-over");
    for (var i = 0; i < cards.length; i++) cards[i].classList.remove("drag-over");
    container.classList.remove("drag-over-append");
  }

  function taskToMultitoolStep(task) {
    var type = task.type;
    var allowed = MULTITOOL_ALLOWED_TYPES.some(function (t) { return t.value === type; });
    if (!allowed) return null;
    var params = task.parameters || {};
    var step = { type: type, collapsed: false, logic: "AND" };
    if (task.region) step.region = task.region;
    if (task.region_ref) step.region_ref = task.region_ref;
    if (params.reference_timestamp !== undefined) step._refTs = params.reference_timestamp;
    if (params.scene_references) step._scenes = params.scene_references.map(function (ref) {
      return { name: ref.name, timestamp: ref.timestamp, threshold: numberOrDefault(ref.threshold, 0.75) };
    });
    step._initial = params;
    return step;
  }

  // Walk the current step list, read each step's per-input DOM values, and
  // store them on step._initial in the same shape the _mtRender* helpers
  // expect. Call before any list mutation (add / remove / reorder / import)
  // so the upcoming renderWorkflowParams() restores values instead of
  // collapsing them back to per-input defaults. Region, _refTs and _scenes
  // already live on the step object, so they don't need snapshotting.
  function snapshotMultitoolStepValues() {
    state.multitoolSteps.forEach(function (step, idx) {
      var sfx = "_mt" + idx;
      var init = step._initial || {};
      if (step.type === "color") {
        var prevColor = init.target_color || {};
        init.target_color = {
          h: numberOrDefault((qs("#paramColorH" + sfx) || {}).value, prevColor.h || 0),
          s: numberOrDefault((qs("#paramColorS" + sfx) || {}).value, prevColor.s || 0),
          v: numberOrDefault((qs("#paramColorV" + sfx) || {}).value, prevColor.v || 0),
        };
        var tol = numberOrDefault((qs("#paramColorTol" + sfx) || {}).value, 30);
        init.tolerance = {
          h: Math.round(tol * 90 / 100),
          s: Math.round(tol * 128 / 100),
          v: Math.round(tol * 128 / 100),
        };
      } else if (step.type === "change") {
        init.threshold = numberOrDefault((qs("#paramChangeThresh" + sfx) || {}).value, init.threshold);
        init.noise_threshold = intOrDefault((qs("#paramChangeNoise" + sfx) || {}).value, init.noise_threshold);
      } else if (step.type === "similarity") {
        init.threshold = numberOrDefault((qs("#paramSimThresh" + sfx) || {}).value, init.threshold);
      } else if (step.type === "text") {
        var searchEl = qs("#paramTextSearch" + sfx);
        if (searchEl) init.search_string = searchEl.value;
        init.fuzzy_threshold = numberOrDefault((qs("#paramTextFuzzy" + sfx) || {}).value, init.fuzzy_threshold);
      } else if (step.type === "numbers") {
        var opEl = qs("#paramNumOperator" + sfx);
        if (opEl) init.operator = opEl.value;
        init.target_value = numberOrDefault((qs("#paramNumTarget" + sfx) || {}).value, init.target_value);
      } else if (step.type === "template") {
        init.threshold = numberOrDefault((qs("#paramTemplateThresh" + sfx) || {}).value, init.threshold);
      } else if (step.type === "flow") {
        init.magnitude_threshold = numberOrDefault((qs("#paramFlowMag" + sfx) || {}).value, init.magnitude_threshold);
      } else if (step.type === "inactivity") {
        init.threshold = intOrDefault((qs("#paramInactThresh" + sfx) || {}).value, init.threshold);
      }
      step._initial = init;
    });
  }

  function renderMultitoolParams(container) {
    var stepsDiv = el("div", "multitool-steps");
    state.multitoolSteps.forEach(function (step, idx) {
      if (idx > 0) {
        var opRow = el("div", "multitool-operator-row");
        var rail = el("div", "multitool-operator-rail");
        rail.appendChild(el("div", "multitool-operator-line"));
        var opBtn = el("button", "multitool-operator-btn");
        opBtn.type = "button";
        var current = (step.logic || "AND").toUpperCase();
        opBtn.classList.add(current === "NOT" ? "is-not" : "is-and");
        opBtn.title = current === "NOT"
          ? "NOT — frame rejected if this matches (click to switch to AND)"
          : "AND — frame must also match (click to switch to NOT)";
        opBtn.appendChild(el("span", "multitool-operator-icon"));
        (function (capturedIdx) {
          opBtn.addEventListener("click", function (e) {
            e.stopPropagation();
            var s = state.multitoolSteps[capturedIdx];
            var next = (s.logic || "AND").toUpperCase() === "NOT" ? "AND" : "NOT";
            s.logic = next;
            opBtn.classList.toggle("is-not", next === "NOT");
            opBtn.classList.toggle("is-and", next === "AND");
            opBtn.title = next === "NOT"
              ? "NOT — frame rejected if this matches (click to switch to AND)"
              : "AND — frame must also match (click to switch to NOT)";
          });
        })(idx);
        rail.appendChild(opBtn);
        rail.appendChild(el("div", "multitool-operator-line"));
        opRow.appendChild(rail);
        stepsDiv.appendChild(opRow);
      }
      var card = el("div", "multitool-step");
      card.dataset.stepIdx = String(idx);
      var header = el("div", "multitool-step-header");

      // Drag handle
      var dragHandle = el("span", "multitool-step-drag-handle");
      dragHandle.appendChild(iconSpan("bars-2"));
      dragHandle.addEventListener("mousedown", function () { card.setAttribute("draggable", "true"); });
      dragHandle.addEventListener("mouseup", function () { card.removeAttribute("draggable"); });
      header.appendChild(dragHandle);

      header.appendChild(el("span", "multitool-step-num", String(idx + 1)));
      var typeSpan = el("span", "multitool-step-type");
      typeSpan.style.color = taskTypeColor(step.type);
      var icon = buildTypeIcon(step.type);
      if (icon) typeSpan.appendChild(icon);
      typeSpan.appendChild(document.createTextNode(" " + step.type.charAt(0).toUpperCase() + step.type.slice(1)));
      header.appendChild(typeSpan);
      var removeBtn = el("button", "multitool-step-remove");
      removeBtn.title = "Remove step";
      removeBtn.appendChild(iconSpan("x-mark"));
      (function (capturedIdx) {
        removeBtn.addEventListener("click", function (e) {
          e.stopPropagation();
          snapshotMultitoolStepValues();
          state.multitoolSteps.splice(capturedIdx, 1);
          renderWorkflowParams();
          updateRunButton();
        });
      })(idx);
      header.appendChild(removeBtn);
      card.appendChild(header);

      var body = el("div", "multitool-step-body" + (step.collapsed ? " collapsed" : ""));
      renderMultitoolStepBody(body, step.type, idx);
      card.appendChild(body);

      header.addEventListener("click", function (e) {
        if (e.target.closest(".multitool-step-drag-handle") || e.target.closest(".multitool-step-remove")) return;
        step.collapsed = !step.collapsed;
        body.classList.toggle("collapsed", step.collapsed);
      });

      stepsDiv.appendChild(card);
    });

    if (state.multitoolSteps.length === 0) {
      // Visible drop target so a Task card has somewhere to land when the
      // step list is empty (an empty flex container is 0px tall and never
      // receives dragover events).
      var emptyDz = el("div", "multitool-empty-dropzone",
        "Drag a Task here, or use + Add Step below");
      stepsDiv.appendChild(emptyDz);
    }

    // Drag-and-drop reordering
    stepsDiv.addEventListener("dragstart", function (e) {
      var card = e.target.closest(".multitool-step");
      if (!card) { e.preventDefault(); return; }
      card.classList.add("dragging");
      e.dataTransfer.setData("text/plain", card.dataset.stepIdx);
      e.dataTransfer.effectAllowed = "move";
      _cacheMultitoolDragMidpoints(stepsDiv);
    });
    stepsDiv.addEventListener("dragend", function (e) {
      var card = e.target.closest(".multitool-step");
      if (card) {
        card.classList.remove("dragging");
        card.removeAttribute("draggable");
      }
      if (_multitoolDragOverRaf != null) {
        cancelAnimationFrame(_multitoolDragOverRaf);
        _multitoolDragOverRaf = null;
      }
      _multitoolPendingDragOver = null;
      clearMultitoolDragIndicators(stepsDiv);
      _multitoolDragMidpoints = null;
    });
    stepsDiv.addEventListener("dragover", function (e) {
      if (e.dataTransfer.types.indexOf("text/plain") < 0) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      _multitoolPendingDragOver = {
        clientY: e.clientY,
        isTaskDrop: e.dataTransfer.types.indexOf("application/x-task-id") >= 0,
      };
      if (_multitoolDragOverRaf != null) return;
      _multitoolDragOverRaf = requestAnimationFrame(function () {
        _multitoolDragOverRaf = null;
        var pending = _multitoolPendingDragOver;
        if (!pending) return;
        clearMultitoolDragIndicators(stepsDiv);
        if (pending.isTaskDrop) {
          stepsDiv.classList.add("drag-over-append");
        } else {
          var cards = stepsDiv.querySelectorAll(".multitool-step:not(.dragging)");
          var insertIdx = getMultitoolDropIndex(stepsDiv, pending.clientY);
          if (insertIdx < cards.length) {
            cards[insertIdx].classList.add("drag-over");
          } else {
            stepsDiv.classList.add("drag-over-append");
          }
        }
      });
    });
    stepsDiv.addEventListener("dragleave", function (e) {
      var card = e.target.closest(".multitool-step");
      if (card) card.classList.remove("drag-over");
      if (!stepsDiv.contains(e.relatedTarget)) {
        stepsDiv.classList.remove("drag-over-append");
      }
    });
    stepsDiv.addEventListener("drop", function (e) {
      e.preventDefault();
      clearMultitoolDragIndicators(stepsDiv);
      stepsDiv.classList.remove("drag-over-append");

      // Check for task card drop (import task as step)
      var taskId = e.dataTransfer.getData("application/x-task-id");
      if (taskId) {
        var task = findTask(taskId);
        if (!task) return;
        var step = taskToMultitoolStep(task);
        if (!step) {
          showToast(task.type + " cannot be added as a multitool step");
          return;
        }
        snapshotMultitoolStepValues();
        state.multitoolSteps.push(step);
        renderWorkflowParams();
        showToast("Imported " + task.type + " task as step");
        return;
      }

      // Step reorder. _cacheMultitoolDragMidpoints excludes the dragging
      // card, so getMultitoolDropIndex already returns an index aligned with
      // the array AFTER the dragging step is spliced out — no further
      // adjustment needed.
      var fromIdx = parseInt(e.dataTransfer.getData("text/plain"), 10);
      if (isNaN(fromIdx)) return;
      var toIdx = getMultitoolDropIndex(stepsDiv, e.clientY);
      if (fromIdx === toIdx) return;
      snapshotMultitoolStepValues();
      var moved = state.multitoolSteps.splice(fromIdx, 1)[0];
      state.multitoolSteps.splice(toIdx, 0, moved);
      renderWorkflowParams();
    });

    container.appendChild(stepsDiv);

    // Add Step row
    var addRow = el("div", "multitool-add-row");
    var sel = document.createElement("select");
    sel.id = "mtAddTypeSelect";
    MULTITOOL_ALLOWED_TYPES.forEach(function (t) {
      var opt = document.createElement("option");
      opt.value = t.value; opt.textContent = t.label;
      sel.appendChild(opt);
    });
    addRow.appendChild(sel);
    var addBtn = el("button", "btn btn-small", "+ Add Step");
    addBtn.addEventListener("click", function () {
      var chosen = sel.value;
      snapshotMultitoolStepValues();
      state.multitoolSteps.push({ type: chosen, collapsed: false, logic: "AND" });
      renderWorkflowParams();
      updateRunButton();
    });
    addRow.appendChild(addBtn);
    container.appendChild(addRow);

    if (state.multitoolSteps.length === 1) {
      container.appendChild(el("div", "multitool-hint", "Add at least 2 tool steps to create a multi-factor filter."));
    }
  }

  // ---- Single-tool parameter panels (color, change, template, …) ----

  function renderColorParams(container) {
    var pickerGroup = el("div", "color-picker-group");

    var palette = document.createElement("canvas");
    palette.id = "colorPalette";
    palette.className = "color-palette-canvas";
    pickerGroup.appendChild(palette);

    var bright = document.createElement("canvas");
    bright.id = "colorBrightness";
    bright.className = "color-brightness-strip";
    pickerGroup.appendChild(bright);

    var inputRow = el("div", "color-input-row");
    var preview = el("div", "color-preview");
    preview.id = "colorPreview";
    inputRow.appendChild(preview);

    var hexInput = document.createElement("input");
    hexInput.type = "text";
    hexInput.autocomplete = "off";
    hexInput.id = "paramColorHex";
    hexInput.className = "color-hex-input";
    hexInput.placeholder = "#000000";
    hexInput.maxLength = 7;
    inputRow.appendChild(hexInput);

    var pipetteBtn = el("button", "btn btn-small btn-pipette");
    pipetteBtn.id = "pipetteBtn";
    pipetteBtn.appendChild(buildTypeIcon("color"));
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

    var hiddenH = document.createElement("input");
    hiddenH.type = "hidden"; hiddenH.id = "paramColorH"; hiddenH.value = "90";
    var hiddenS = document.createElement("input");
    hiddenS.type = "hidden"; hiddenS.id = "paramColorS"; hiddenS.value = "200";
    var hiddenV = document.createElement("input");
    hiddenV.type = "hidden"; hiddenV.id = "paramColorV"; hiddenV.value = "200";
    pickerGroup.appendChild(hiddenH);
    pickerGroup.appendChild(hiddenS);
    pickerGroup.appendChild(hiddenV);
    _colorHiddenInputs = { h: hiddenH, s: hiddenS, v: hiddenV, hex: hexInput };

    container.appendChild(pickerGroup);

    var paletteDragging = false;
    var brightDragging = false;
    function pickFromPalette(e) {
      var rect = palette.getBoundingClientRect();
      var x = clamp(e.clientX - rect.left, 0, rect.width);
      var y = clamp(e.clientY - rect.top, 0, rect.height);
      var h = Math.round((x / rect.width) * 180);
      var s = Math.round((1 - y / rect.height) * 255);
      var curV = numberOrDefault(hiddenV.value, 0);
      setTargetColor(h, s, curV);
    }
    palette.addEventListener("mousedown", function (e) {
      e.preventDefault();
      paletteDragging = true;
      pickFromPalette(e);
    });

    function pickFromBrightness(e) {
      var rect = bright.getBoundingClientRect();
      var x = clamp(e.clientX - rect.left, 0, rect.width);
      var v = Math.round((x / rect.width) * 255);
      var curH = numberOrDefault(hiddenH.value, 0);
      var curS = numberOrDefault(hiddenS.value, 0);
      setTargetColor(curH, curS, v);
    }
    bright.addEventListener("mousedown", function (e) {
      e.preventDefault();
      brightDragging = true;
      pickFromBrightness(e);
    });

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

    hexInput.addEventListener("input", function () {
      var rgb = hexToRgb(hexInput.value);
      if (rgb) {
        var hsv = rgbToHsv(rgb.r, rgb.g, rgb.b);
        hiddenH.value = hsv.h;
        hiddenS.value = hsv.s;
        hiddenV.value = hsv.v;
        updateColorPreview();
        renderColorPalette();
        renderBrightnessStrip();
      }
    });

    var tolSlider = rangeInput("paramColorTol", 0, 100, 30);
    addParamRow(container, "Tolerance", tolSlider, "paramColorTolVal");
    tolSlider.addEventListener("input", function () {
      renderColorPalette();
    });
    renderIntervalSlot("paramColorInterval", 0.5, 60, 1.0, 0.5);

    renderColorPalette();
    renderBrightnessStrip();
    updateColorPreview();
    var initRgb = hsvToRgb(90, 200, 200);
    hexInput.value = rgbToHex(initRgb.r, initRgb.g, initRgb.b);
  }

  function renderSimilarityParams(container) {
    var refRow = el("div", "param-row");
    var refLabel = el("span", "param-label", "Reference");
    var refControl = el("div", "param-control");
    var refBtn = el("button", "btn btn-small", "Capture Current Frame");
    refBtn.addEventListener("click", function () {
      state.referenceTimestamp = state.currentTimestamp;
      renderWorkflowParams();
      showToast("Reference frame captured at " + formatTime(state.currentTimestamp, { decimals: 1 }));
    });
    refControl.appendChild(refBtn);
    if (state.referenceTimestamp !== null) {
      var refTs = el("span", "param-value", formatTime(state.referenceTimestamp, { decimals: 1 }));
      refControl.appendChild(refTs);
    }
    refRow.appendChild(refLabel);
    refRow.appendChild(refControl);
    container.appendChild(refRow);
    addParamRow(container, "Threshold", rangeInput("paramSimThresh", 0.50, 1.00, 0.90, 0.01), "paramSimThreshVal");
    renderIntervalSlot("paramSimInterval", 0.5, 60, 1.0, 0.5);
  }

  function renderTextParams(container) {
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
  }

  function renderNumbersParams(container) {
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
    var targetRow = el("div", "param-row");
    targetRow.id = "paramNumTargetRow";
    targetRow.appendChild(el("span", "param-label", "Target value"));
    var targetCtrl = el("div", "param-control");
    targetCtrl.appendChild(numberInput("paramNumTarget", -999999, 999999, 100, 1));
    targetRow.appendChild(targetCtrl);
    container.appendChild(targetRow);
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
  }

  function renderTimelapseParams(container) {
    addParamRow(container, "Speed", numberInput("paramTlSpeed", 2, 100, 10, 1));
    addParamRow(container, "Sample every", numberInput("paramTlSampleInterval", 0, 60, 0, 0.5), "paramTlSampleIntervalVal");
    var siHint = el("span", "param-hint", "seconds (0 = every frame)");
    container.lastChild.querySelector(".param-control").appendChild(siHint);
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
  }

  function renderTemplateParams(container) {
    var tmplRefRow = el("div", "param-row");
    tmplRefRow.appendChild(el("span", "param-label", "Template"));
    var tmplRefCtrl = el("div", "param-control");
    var tmplCapBtn = el("button", "btn btn-small ss-template-icon-btn ss-template-icon-btn--capture");
    tmplCapBtn.setAttribute("type", "button");
    tmplCapBtn.title = "Capture Region";
    tmplCapBtn.setAttribute("aria-label", "Capture Region");
    var tmplCapGlyph = el("span", "ss-template-icon-btn__glyph");
    tmplCapBtn.appendChild(tmplCapGlyph);
    tmplCapBtn.addEventListener("click", function () {
      state.referenceTimestamp = state.currentTimestamp;
      state.uploadedTemplate = null;
      renderWorkflowParams();
      showToast("Template captured at " + formatTime(state.currentTimestamp, { decimals: 1 }));
    });
    tmplRefCtrl.appendChild(tmplCapBtn);

    var tmplFileInput = document.createElement("input");
    tmplFileInput.type = "file";
    tmplFileInput.accept = "image/png";
    tmplFileInput.style.display = "none";
    tmplFileInput.addEventListener("change", function () {
      var file = tmplFileInput.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function (e) {
        var dataUrl = e.target.result;
        var b64 = dataUrl.split(",")[1];
        state.uploadedTemplate = { name: file.name, data: b64 };
        state.referenceTimestamp = null;
        state.templateOverlayPos = null;
        var previewImg = new Image();
        previewImg.onload = function () { renderOverlay(); };
        previewImg.src = dataUrl;
        state.uploadedTemplateImg = previewImg;
        renderWorkflowParams();
        showToast("Template loaded");
      };
      reader.readAsDataURL(file);
    });
    var tmplUploadBtn = el("button", "btn btn-small ss-template-icon-btn ss-template-icon-btn--upload");
    tmplUploadBtn.setAttribute("type", "button");
    tmplUploadBtn.title = "Upload PNG";
    tmplUploadBtn.setAttribute("aria-label", "Upload PNG");
    var tmplUploadGlyph = el("span", "ss-template-icon-btn__glyph");
    tmplUploadBtn.appendChild(tmplUploadGlyph);
    tmplUploadBtn.addEventListener("click", function () { tmplFileInput.click(); });
    tmplRefCtrl.appendChild(tmplUploadBtn);
    tmplRefCtrl.appendChild(tmplFileInput);

    if (state.uploadedTemplate) {
      if (!state.uploadedTemplateImg) {
        var liveImg = new Image();
        liveImg.onload = function () { renderOverlay(); };
        liveImg.src = "data:image/png;base64," + state.uploadedTemplate.data;
        state.uploadedTemplateImg = liveImg;
      }
      var uploadInfo = el("span", "param-value template-upload-info");
      var uploadThumb = document.createElement("img");
      uploadThumb.src = "data:image/png;base64," + state.uploadedTemplate.data;
      uploadThumb.alt = "Uploaded template";
      uploadThumb.title = state.uploadedTemplate.name;
      uploadInfo.appendChild(uploadThumb);
      var clearBtn = el("button", "btn btn-small", "\u00d7");
      clearBtn.addEventListener("click", function () {
        state.uploadedTemplate = null;
        state.uploadedTemplateImg = null;
        state.templateOverlayPos = null;
        renderWorkflowParams();
        renderOverlay();
      });
      uploadInfo.appendChild(clearBtn);
      tmplRefCtrl.appendChild(uploadInfo);
    } else if (state.referenceTimestamp !== null) {
      tmplRefCtrl.appendChild(el("span", "param-value", formatTime(state.referenceTimestamp, { decimals: 1 })));
    }
    tmplRefRow.appendChild(tmplRefCtrl);
    container.appendChild(tmplRefRow);
    addParamRow(container, "Threshold", rangeInput("paramTemplateThresh", 0.50, 1.00, 0.70, 0.01), "paramTemplateThreshVal");
    addParamRow(container, "Template scale", rangeInput("paramTemplateScale", 25, 200, 100, 5), "paramTemplateScaleVal");
    var scaleHint = el("span", "param-hint", "% \u2014 resize the uploaded PNG before matching");
    container.lastChild.querySelector(".param-control").appendChild(scaleHint);
    var scaleSlider = qs("#paramTemplateScale");
    if (scaleSlider) {
      state.templateScalePreview = numberOrDefault(scaleSlider.value, 100) / 100;
      scaleSlider.addEventListener("input", function () {
        state.templateScalePreview = numberOrDefault(scaleSlider.value, 100) / 100;
        renderOverlay();
      });
    }

    renderIntervalSlot("paramTemplateInterval", 0.5, 60, 1.0, 0.5);
  }

  function renderSceneParams(container) {
    var sceneList = el("div", "scene-reference-list");
    sceneList.id = "sceneRefList";
    state.sceneReferences.forEach(function (ref, i) {
      if (ref.threshold === undefined) ref.threshold = 0.75;
      var item = el("div", "scene-ref-item");
      item.appendChild(el("span", "scene-ref-name", ref.name));
      item.appendChild(el("span", "param-value", formatTime(ref.timestamp, { decimals: 1 })));
      var threshSlider = document.createElement("input");
      threshSlider.type = "range";
      threshSlider.min = "0.50";
      threshSlider.max = "1.00";
      threshSlider.step = "0.01";
      threshSlider.value = String(ref.threshold);
      threshSlider.className = "scene-ref-thresh";
      var threshVal = el("span", "param-value", String(ref.threshold));
      threshSlider.addEventListener("input", (function (idx) {
        return function () {
          state.sceneReferences[idx].threshold = parseFloat(threshSlider.value);
          threshVal.textContent = threshSlider.value;
        };
      })(i));
      item.appendChild(threshSlider);
      item.appendChild(threshVal);
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
      state.sceneReferences.push({ name: name, timestamp: state.currentTimestamp, threshold: 0.75 });
      renderWorkflowParams();
      showToast("Scene '" + name + "' at " + formatTime(state.currentTimestamp, { decimals: 1 }));
    });
    addScCtrl.appendChild(scCapBtn);
    addScRow.appendChild(addScCtrl);
    container.appendChild(addScRow);
    renderIntervalSlot("paramSceneInterval", 0.5, 60, 1.0, 0.5);
  }

  function renderWorkflowParams() {
    var container = qs("#workflowParams");
    container.innerHTML = "";
    _colorHiddenInputs = null;
    hideToolInfoTooltip(true);
    _toolInfoPinned = false;
    var intervalSlot = qs("#workflowIntervalSlot");
    if (intervalSlot) intervalSlot.innerHTML = "";
    var type = state.activeWorkflow;

    var regionPickerWrap = qs("#runRegionPicker");
    if (regionPickerWrap) regionPickerWrap.style.display = type === "multitool" ? "none" : "";

    if (type === "multitool") {
      renderMultitoolParams(container);
      renderIntervalSlot("paramMultitoolInterval", 0.5, 60, 1.0, 0.5);
      addParamRow(container, "Event label", textInput("paramEventLabel", "e.g. low_health"));
      var dfCb = document.createElement("input");
      dfCb.type = "checkbox";
      dfCb.id = "paramDetectFirst";
      addParamRow(container, "Detect first", dfCb);
      // Every multitool mutation funnels through renderWorkflowParams; refresh
      // the Run button here so task-import / reorder paths (which don't call
      // updateRunButton explicitly) still enable the button once the list is
      // long enough.
      updateRunButton();
      return;
    }

    if (type === "color") renderColorParams(container);
    else if (type === "change") {
      addParamRow(container, "Threshold", rangeInput("paramChangeThresh", 0.01, 0.50, 0.03, 0.01), "paramChangeThreshVal");
      addParamRow(container, "Noise Thr.", rangeInput("paramChangeNoise", 0, 100, 30, 1), "paramChangeNoiseVal");
      renderIntervalSlot("paramChangeInterval", 0.5, 60, 1.0, 0.5);
    }
    else if (type === "similarity") renderSimilarityParams(container);
    else if (type === "text") renderTextParams(container);
    else if (type === "numbers") renderNumbersParams(container);
    else if (type === "timelapse") renderTimelapseParams(container);
    else if (type === "template") renderTemplateParams(container);
    else if (type === "flow") {
      addParamRow(container, "Magnitude", rangeInput("paramFlowMag", 0.5, 20.0, 2.0, 0.5), "paramFlowMagVal");
      renderIntervalSlot("paramFlowInterval", 0.5, 60, 1.0, 0.5);
    }
    else if (type === "scene") renderSceneParams(container);
    else if (type === "inactivity") {
      addParamRow(container, "Sensitivity", rangeInput("paramInactThresh", 0, 30, 10, 1), "paramInactThreshVal");
      addParamRow(container, "Min duration (s)", numberInput("paramInactMinDur", 0.5, 60, 2.0, 0.5));
      renderIntervalSlot("paramInactInterval", 0.5, 60, 1.0, 0.5);
    }

    if (type !== "timelapse") {
      addParamRow(container, "Event label", textInput("paramEventLabel", "e.g. low_health"));
      dfCb = document.createElement("input");
      dfCb.type = "checkbox";
      dfCb.id = "paramDetectFirst";
      addParamRow(container, "Detect first", dfCb);
    }

    var scanPicker = qs("#runScanModePicker");
    if (scanPicker) scanPicker.style.display = type === "timelapse" ? "none" : "";
    var scanBtn = scanPicker && scanPicker.querySelector(".scan-toggle-btn");
    if (scanBtn && scanBtn._updateScanState) scanBtn._updateScanState();

    updateRunButton();
    _updateOverlayUi();
    refreshModelView();
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
        refreshModelView({ debounce: true });
      });
    } else {
      control.addEventListener("input", function () {
        refreshModelView({ debounce: true });
      });
    }
    row.appendChild(ctrl);
    container.appendChild(row);
  }

  // ---- Model view (preprocessed preview) ----

  var _modelViewGen = 0;
  var _modelViewTimer = 0;

  var MODEL_VIEW_META = {
    color: "Downscaled region (≤64 px) with mean HSV vs. target swatch.",
    change: "Gray-blur + abs-diff + thresholded mask (prev = 1 s earlier).",
    similarity: "Gray-blurred region (≤256 px); reference appears once captured.",
    text: "Grayscale region fed to OCR.",
    numbers: "Grayscale region fed to OCR.",
    timelapse: "Region crop — FFmpeg encodes this unmodified.",
    template: "Gray-blurred frame, template, and normalized match heatmap.",
    flow: "Prev + current gray frames with dense optical-flow vectors.",
    scene: "Region (≤128 px), Canny edges, and 8-bin hue histogram.",
    inactivity: "Region and pHash bit grid (white = 1, black = 0).",
    multitool: "Preview of the first tool step.",
  };

  function initModelView() {
    var btn = qs("#modelViewToggle");
    if (btn) btn.addEventListener("click", toggleModelView);

    // Restore persisted overlay preferences (sessionStorage, per-tab).
    try {
      var stored = sessionStorage.getItem("ss_overlayEnabled");
      if (stored === "1") state.overlayEnabled = true;
    } catch (_) { /* sessionStorage may be unavailable */ }
    try {
      var layer = sessionStorage.getItem("ss_overlayLayer");
      if (layer) state.overlayLayer = layer;
    } catch (_) { /* ignore */ }

    var toggle = qs("#modelViewOverlayToggle");
    if (toggle) {
      toggle.checked = !!state.overlayEnabled;
      toggle.addEventListener("change", function () {
        state.overlayEnabled = !!toggle.checked;
        try { sessionStorage.setItem("ss_overlayEnabled", state.overlayEnabled ? "1" : "0"); } catch (_) { /* ignore */ }
        var curTs = Number(state.currentTimestamp || 0).toFixed(3);
        if (state.overlayEnabled && (!state.overlayImage || state.overlayImageTimestamp !== curTs || state.overlayImageTool !== state.activeWorkflow)) {
          refreshModelView();
        }
        renderOverlay();
      });
    }

    var sel = qs("#modelViewOverlayLayer");
    if (sel) {
      sel.addEventListener("change", function () {
        state.overlayLayer = sel.value || null;
        try { sessionStorage.setItem("ss_overlayLayer", state.overlayLayer || ""); } catch (_) { /* ignore */ }
        // Force a refetch of the overlay image at the new layer.
        if (state.overlayImageObjectUrl) {
          URL.revokeObjectURL(state.overlayImageObjectUrl);
          state.overlayImageObjectUrl = null;
        }
        state.overlayImage = null;
        state.overlayImageTimestamp = null;
        state.overlayImageTool = null;
        refreshModelView();
        renderOverlay();
      });
    }

    // Fetch the overlay-layer catalog once at init.
    apiGet("api/preview/layers")
      .then(function (data) {
        if (data && data.ok && data.layers) {
          state.overlayLayerSpec = data.layers;
          _updateOverlayUi();
        }
      })
      .catch(function () { /* leave catalog empty; toggle stays disabled */ });
  }

  function _activeOverlayTool() {
    var tool = state.activeWorkflow;
    if (tool === "multitool") {
      var first = (state.multitoolSteps || [])[0];
      tool = first && first.type ? first.type : null;
    }
    return tool;
  }

  function _activeOverlayLayers() {
    var tool = _activeOverlayTool();
    if (!tool) return [];
    return state.overlayLayerSpec[tool] || [];
  }

  function _overlayEligibleForActiveTool() {
    return _activeOverlayLayers().length > 0;
  }

  function _resolveOverlayLayer() {
    var layers = _activeOverlayLayers();
    if (!layers.length) return null;
    if (state.overlayLayer) {
      for (var i = 0; i < layers.length; i++) {
        if (layers[i].id === state.overlayLayer) return layers[i];
      }
    }
    return layers[0];
  }

  function _updateOverlayUi() {
    var toggle = qs("#modelViewOverlayToggle");
    var sel = qs("#modelViewOverlayLayer");
    var label = toggle && toggle.parentElement;
    var layers = _activeOverlayLayers();
    var eligible = layers.length > 0;

    if (toggle) {
      toggle.disabled = !eligible;
      if (label) {
        if (eligible) {
          label.classList.remove("disabled");
          label.removeAttribute("title");
        } else {
          label.classList.add("disabled");
          label.setAttribute("title", "This tool's preview isn't pixel-aligned to the frame");
        }
      }
      if (!eligible && toggle.checked) {
        toggle.checked = false;
        // Don't persist away the user's preference; just disable for this tool.
      }
    }

    if (sel) {
      if (layers.length > 1) {
        sel.classList.remove("hidden");
        sel.innerHTML = "";
        var resolved = _resolveOverlayLayer();
        layers.forEach(function (layer) {
          var opt = document.createElement("option");
          opt.value = layer.id;
          opt.textContent = layer.label;
          if (resolved && layer.id === resolved.id) opt.selected = true;
          sel.appendChild(opt);
        });
      } else {
        sel.classList.add("hidden");
        sel.innerHTML = "";
      }
    }
  }

  function toggleModelView() {
    state.modelViewOpen = !state.modelViewOpen;
    var panel = qs("#modelViewPanel");
    var body = qs("#modelViewBody");
    var btn = qs("#modelViewToggle");
    if (state.modelViewOpen) {
      panel.classList.remove("collapsed");
      body.classList.remove("hidden");
      btn.setAttribute("aria-expanded", "true");
      refreshModelView();
    } else {
      panel.classList.add("collapsed");
      body.classList.add("hidden");
      btn.setAttribute("aria-expanded", "false");
    }
  }

  function refreshModelView(opts) {
    if (!state.modelViewOpen && !state.overlayEnabled && !state.overlayBlinkActive) return;
    if (_modelViewTimer) {
      clearTimeout(_modelViewTimer);
      _modelViewTimer = 0;
    }
    if (opts && opts.debounce) {
      _modelViewTimer = setTimeout(_doRefreshModelView, 150);
    } else {
      _doRefreshModelView();
    }
  }

  var _FULL_FRAME_REGION_STRING = "0.000000,0.000000,1.000000,1.000000";

  function _normalizedRegionString() {
    if (state.pendingRegion) {
      var p = state.pendingRegion;
      var c = qs("#overlayCanvas");
      if (!c.width || !c.height) return _FULL_FRAME_REGION_STRING;
      return [p.x / c.width, p.y / c.height, p.w / c.width, p.h / c.height]
        .map(function (v) { return Number(v).toFixed(6); })
        .join(",");
    }
    if (state.activeRegion && state.regions[state.activeRegion]) {
      var r = state.regions[state.activeRegion];
      if (r.source_width) {
        return [r.x, r.y, r.w, r.h]
          .map(function (v) { return Number(v).toFixed(6); })
          .join(",");
      }
      var canvas = qs("#overlayCanvas");
      if (!canvas.width || !canvas.height) return _FULL_FRAME_REGION_STRING;
      return [r.x / canvas.width, r.y / canvas.height, r.w / canvas.width, r.h / canvas.height]
        .map(function (v) { return Number(v).toFixed(6); })
        .join(",");
    }
    return _FULL_FRAME_REGION_STRING;
  }

  function _hasActiveOrPendingRegion() {
    return !!(state.pendingRegion || (state.activeRegion && state.regions[state.activeRegion]));
  }

  function _collectPreviewParams(tool) {
    var out = {};
    if (tool === "color") {
      var c = _colorHiddenInputs;
      if (c) {
        out.h = c.h.value; out.s = c.s.value; out.v = c.v.value;
      }
    } else if (tool === "change") {
      var n = qs("#paramChangeNoise");
      if (n) out.noise = n.value;
    } else if (tool === "flow") {
      var m = qs("#paramFlowMag");
      if (m) out.magnitude = m.value;
    }
    return out;
  }

  function _doRefreshModelView() {
    var gen = ++_modelViewGen;
    var meta = qs("#modelViewMeta");
    var img = qs("#modelViewImage");
    if (!meta || !img) return;

    if (!state.selectedParticipant) {
      meta.textContent = "Select a participant to preview.";
      img.removeAttribute("src");
      return;
    }

    var tool = state.activeWorkflow;
    var regionStr = _normalizedRegionString();
    var hasRegion = _hasActiveOrPendingRegion();

    if (tool === "template") {
      if (state.uploadedTemplate && state.uploadedTemplate.data) {
        // POST with template_image_data — region optional
      } else if (state.referenceTimestamp != null) {
        if (!hasRegion) {
          meta.textContent = "Select or draw a region to preview the captured template.";
          img.removeAttribute("src");
          return;
        }
      } else {
        meta.textContent = "Capture a template region or upload a PNG to preview.";
        img.removeAttribute("src");
        return;
      }
    }

    var params = _collectPreviewParams(tool);
    var qsParts = ["tool=" + encodeURIComponent(tool)];
    if (regionStr) qsParts.push("region=" + regionStr);
    if (tool === "change" || tool === "flow") {
      var prevTs = Math.max(0, (state.currentTimestamp || 0) - 1);
      qsParts.push("prev=" + prevTs.toFixed(3));
    }
    if (tool === "similarity" && state.referenceTimestamp != null) {
      qsParts.push("ref=" + Number(state.referenceTimestamp).toFixed(3));
    }
    if (
      tool === "template" &&
      state.referenceTimestamp != null &&
      !(state.uploadedTemplate && state.uploadedTemplate.data)
    ) {
      qsParts.push("ref=" + Number(state.referenceTimestamp).toFixed(3));
    }
    Object.keys(params).forEach(function (k) {
      qsParts.push(encodeURIComponent(k) + "=" + encodeURIComponent(params[k]));
    });
    qsParts.push("_=" + gen);

    var ts = Number(state.currentTimestamp || 0).toFixed(3);
    var url = "api/preview/" + encodeURIComponent(state.selectedParticipant)
      + "/" + ts + "?" + qsParts.join("&");

    meta.textContent = "Loading preview…";

    function applyPreviewError() {
      if (gen !== _modelViewGen) return;
      meta.textContent = "Preview unavailable.";
      img.removeAttribute("src");
    }

    function applyPreviewOkFromBlob(blob) {
      if (gen !== _modelViewGen) return;
      if (img._modelViewObjectUrl) {
        URL.revokeObjectURL(img._modelViewObjectUrl);
        img._modelViewObjectUrl = null;
      }
      var u = URL.createObjectURL(blob);
      img._modelViewObjectUrl = u;
      img.src = u;
      var metaText = MODEL_VIEW_META[tool] || "";
      if (!hasRegion) {
        metaText = (metaText ? metaText + " " : "") + "(Full frame — no region selected.)";
      }
      meta.textContent = metaText;
    }

    function _refetchOverlayLayer() {
      if (gen !== _modelViewGen) return;
      var resolved = _resolveOverlayLayer();
      if (!resolved) {
        if (state.overlayImageObjectUrl) {
          URL.revokeObjectURL(state.overlayImageObjectUrl);
          state.overlayImageObjectUrl = null;
        }
        state.overlayImage = null;
        state.overlayImageTimestamp = null;
        state.overlayImageTool = null;
        return;
      }
      var layerQs = qsParts.concat(["layer=" + encodeURIComponent(resolved.id)]);
      var layerUrl = "api/preview/" + encodeURIComponent(state.selectedParticipant)
        + "/" + ts + "?" + layerQs.join("&");
      function fetchAsImage(blob) {
        if (gen !== _modelViewGen) return;
        if (state.overlayImageObjectUrl) {
          URL.revokeObjectURL(state.overlayImageObjectUrl);
          state.overlayImageObjectUrl = null;
        }
        var ou = URL.createObjectURL(blob);
        state.overlayImageObjectUrl = ou;
        var oi = new Image();
        oi.onload = function () {
          if (gen !== _modelViewGen) return;
          state.overlayImage = oi;
          state.overlayImageScope = resolved.scope;
          state.overlayImageTimestamp = ts;
          state.overlayImageTool = tool;
          renderOverlay();
        };
        oi.src = ou;
      }
      if (useTemplatePost) {
        fetch(layerUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ template_image_data: state.uploadedTemplate.data }),
        })
          .then(function (r) { if (!r.ok) throw new Error("layer http " + r.status); return r.blob(); })
          .then(fetchAsImage)
          .catch(function () { /* leave previous overlay image */ });
      } else {
        fetch(layerUrl)
          .then(function (r) { if (!r.ok) throw new Error("layer http " + r.status); return r.blob(); })
          .then(fetchAsImage)
          .catch(function () { /* leave previous overlay image */ });
      }
    }

    var useTemplatePost = tool === "template" && state.uploadedTemplate && state.uploadedTemplate.data;
    if (useTemplatePost) {
      // TODO: utils.apiPost returns JSON; needs an apiPostBlob helper to migrate.
      fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ template_image_data: state.uploadedTemplate.data }),
      })
        .then(function (r) {
          if (!r.ok) throw new Error("preview http " + r.status);
          return r.blob();
        })
        .then(function (blob) {
          if (gen !== _modelViewGen) return;
          applyPreviewOkFromBlob(blob);
          _refetchOverlayLayer();
        })
        .catch(function () {
          applyPreviewError();
        });
      return;
    }

    var tmp = new Image();
    tmp.onload = function () {
      if (gen !== _modelViewGen) return;
      if (img._modelViewObjectUrl) {
        URL.revokeObjectURL(img._modelViewObjectUrl);
        img._modelViewObjectUrl = null;
      }
      img.src = tmp.src;
      meta.textContent = MODEL_VIEW_META[tool] || "";
      _refetchOverlayLayer();
    };
    tmp.onerror = function () {
      applyPreviewError();
    };
    tmp.src = url;
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
    inp.autocomplete = "off";
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
    var c = _colorHiddenInputs;
    if (!preview || !c) return;
    var rgb = hsvToRgb(numberOrDefault(c.h.value, 0), numberOrDefault(c.s.value, 0), numberOrDefault(c.v.value, 0));
    preview.style.background = rgbToHex(rgb.r, rgb.g, rgb.b);
  }

  function setTargetColor(h, s, v) {
    h = clamp(Math.round(h), 0, 180);
    s = clamp(Math.round(s), 0, 255);
    v = clamp(Math.round(v), 0, 255);
    var c = _colorHiddenInputs;
    if (c) {
      c.h.value = h; c.s.value = s; c.v.value = v;
      var rgb = hsvToRgb(h, s, v);
      c.hex.value = rgbToHex(rgb.r, rgb.g, rgb.b);
    }
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
    var c = _colorHiddenInputs;
    var curH = c ? numberOrDefault(c.h.value, 0) : 0;
    var curS = c ? numberOrDefault(c.s.value, 0) : 0;
    var curV = c ? numberOrDefault(c.v.value, 0) : 0;
    var darkness = 1 - curV / 255;
    if (darkness > 0) {
      ctx.fillStyle = "rgba(0,0,0," + darkness + ")";
      ctx.fillRect(0, 0, w, h);
    }

    // Current position
    var cx = (curH / 180) * w;
    var cy = (1 - curS / 255) * h;

    // Tolerance range visualization
    var tol = numberOrDefault((qs("#paramColorTol") || {}).value, 0);
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
    var c = _colorHiddenInputs;
    var curH = c ? numberOrDefault(c.h.value, 0) : 0;
    var curS = c ? numberOrDefault(c.s.value, 0) : 0;
    var curV = c ? numberOrDefault(c.v.value, 0) : 0;

    // Gradient from black (left) to fully saturated color (right)
    var fullRgb = hsvToRgb(curH, curS, 255);
    var grad = ctx.createLinearGradient(0, 0, w, 0);
    grad.addColorStop(0, "#000000");
    grad.addColorStop(1, rgbToHex(fullRgb.r, fullRgb.g, fullRgb.b));
    ctx.fillStyle = grad;
    ctx.fillRect(0, 0, w, h);

    // Position indicator
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
    // Multitool uses per-step regions instead of a global region
    var isMultitool = state.activeWorkflow === "multitool";
    var multitoolReady = isMultitool && state.multitoolSteps.length >= 2;
    var multitoolHasRegions = multitoolReady && state.multitoolSteps.every(function (s) { return !!s.region; });
    if (isMultitool) {
      btn.disabled = !hasParticipants || !multitoolReady || !multitoolHasRegions;
      if (!hasParticipants) {
        btn.setAttribute("data-tooltip", "Select participants to run");
      } else if (!multitoolReady) {
        btn.setAttribute("data-tooltip", "Add at least 2 steps");
      } else if (!multitoolHasRegions) {
        btn.setAttribute("data-tooltip", "Each step needs a region");
      } else {
        btn.removeAttribute("data-tooltip");
      }
    } else {
      var isTemplate = state.activeWorkflow === "template";
      var hasUploadedTemplate = !!state.uploadedTemplate;
      // Template scans full frames regardless of region selection; the region
      // (or uploaded image) only supplies the template patch.
      var templateMissingPatch = isTemplate && !hasRegion && !hasUploadedTemplate;
      var nonTemplateMissingRegion = !isTemplate && !hasRegion;
      btn.disabled = nonTemplateMissingRegion || templateMissingPatch || !hasParticipants;
      if (templateMissingPatch) {
        btn.setAttribute("data-tooltip", "Upload a template image or pick a region first");
      } else if (nonTemplateMissingRegion) {
        btn.setAttribute("data-tooltip", "Select a region first");
      } else if (!hasParticipants) {
        btn.setAttribute("data-tooltip", "Select participants to run");
      } else {
        btn.removeAttribute("data-tooltip");
      }
    }
  }

  // ---- Run analysis ----

  function initRunButton() {
    qs("#runBtn").addEventListener("click", function () {
      var type = state.activeWorkflow;
      var regions = state.runRegions.length > 0
        ? state.runRegions
        : (state.activeRegion ? [activeRegionRef(state.activeRegion)] : []);
      // Multitool uses per-step regions; skip global region requirement
      var isMultitool = type === "multitool";
      // Template with uploaded image can run without a region (full-frame scan)
      if (!isMultitool && regions.length === 0 && !(type === "template" && state.uploadedTemplate)) return;
      if (regions.length === 0) regions = [""];
      var participants = state.runParticipants.length > 0
        ? state.runParticipants
        : (state.selectedParticipant ? [state.selectedParticipant] : []);
      if (participants.length === 0) return;
      var params = gatherWorkflowParams(type);
      if (params === null) return;
      if (state.scanMode === "fast" && type !== "timelapse") params.scan_mode = "fast";

      if (state.inMarker !== null) params.start_seconds = state.inMarker;
      if (state.outMarker !== null) params.end_seconds = state.outMarker;

      var chain = Promise.resolve();
      if (isMultitool) {
        // Multitool: one task per participant, first step's region as top-level
        var mtRegion = (params.steps && params.steps.length > 0) ? (params.steps[0].region || "") : "";
        participants.forEach(function (pid) {
          chain = chain.then(function () {
            var body = {
              type: type,
              participant: pid,
              region: mtRegion,
              parameters: params,
            };
            return apiPost("api/tasks", body).then(function (data) {
              if (data.ok) {
                if (!state.tasks.some(function (t) { return t.id === data.task.id; })) {
                  state.tasks.push(data.task);
                }
                renderTaskList();
              } else {
                showToast(data.error || "Failed to create task for " + pid);
              }
            });
          });
        });
      } else {
        participants.forEach(function (pid) {
          regions.forEach(function (regionRef) {
            chain = chain.then(function () {
              var normalizedRegion = normalizeRegionRef(regionRef);
              var body = {
                type: type,
                participant: pid,
                region: normalizedRegion ? normalizedRegion.name : "",
                parameters: params,
              };
              if (normalizedRegion) body.region_ref = regionRefPayload(normalizedRegion);
              return apiPost("api/tasks", body).then(function (data) {
                if (data.ok) {
                  if (!state.tasks.some(function (t) { return t.id === data.task.id; })) {
                    state.tasks.push(data.task);
                  }
                  renderTaskList();
                } else {
                  showToast(data.error || "Failed to create task for " + pid + " / " + regionRefLabel(normalizedRegion));
                }
              });
            });
          });
        });
      }
      var totalTasks = isMultitool ? participants.length : participants.length * regions.length;
      chain.then(function () {
        showToast(totalTasks + " task" + (totalTasks !== 1 ? "s" : "") + " queued: " + type);
        startSSE();
      }).catch(function (err) { showToast("Error: " + err.message); });
    });
  }

  function gatherMultitoolStepParams(stepType, idx) {
    var sfx = "_mt" + idx;
    var p = {};
    if (stepType === "color") {
      p.target_color = {
        h: numberOrDefault((qs("#paramColorH" + sfx) || {}).value, 0),
        s: numberOrDefault((qs("#paramColorS" + sfx) || {}).value, 0),
        v: numberOrDefault((qs("#paramColorV" + sfx) || {}).value, 0),
      };
      var tol = numberOrDefault((qs("#paramColorTol" + sfx) || {}).value, 30);
      p.tolerance = {
        h: Math.round(tol * 90 / 100),
        s: Math.round(tol * 128 / 100),
        v: Math.round(tol * 128 / 100),
      };
    } else if (stepType === "change") {
      p.threshold = numberOrDefault((qs("#paramChangeThresh" + sfx) || {}).value, 0.03);
      p.noise_threshold = intOrDefault((qs("#paramChangeNoise" + sfx) || {}).value, 30);
    } else if (stepType === "similarity") {
      var step = state.multitoolSteps[idx];
      if (!step || step._refTs === undefined) {
        showToast("Step " + (idx + 1) + ": capture a reference frame first");
        return null;
      }
      p.reference_timestamp = step._refTs;
      p.threshold = numberOrDefault((qs("#paramSimThresh" + sfx) || {}).value, 0.90);
    } else if (stepType === "text") {
      p.search_string = (qs("#paramTextSearch" + sfx) || {}).value || "";
      if (!p.search_string.trim()) {
        showToast("Step " + (idx + 1) + ": enter a search string");
        return null;
      }
      p.fuzzy_threshold = numberOrDefault((qs("#paramTextFuzzy" + sfx) || {}).value, 0.80);
    } else if (stepType === "numbers") {
      p.operator = (qs("#paramNumOperator" + sfx) || {}).value || "gt";
      p.target_value = parseFloat((qs("#paramNumTarget" + sfx) || {}).value);
      if (isNaN(p.target_value)) {
        showToast("Step " + (idx + 1) + ": enter a valid target number");
        return null;
      }
    } else if (stepType === "template") {
      step = state.multitoolSteps[idx];
      if (!step || step._refTs === undefined) {
        showToast("Step " + (idx + 1) + ": capture a template frame first");
        return null;
      }
      p.reference_timestamp = step._refTs;
      p.threshold = numberOrDefault((qs("#paramTemplateThresh" + sfx) || {}).value, 0.70);
    } else if (stepType === "flow") {
      p.magnitude_threshold = numberOrDefault((qs("#paramFlowMag" + sfx) || {}).value, 2.0);
    } else if (stepType === "scene") {
      step = state.multitoolSteps[idx];
      if (!step || !step._scenes || step._scenes.length === 0) {
        showToast("Step " + (idx + 1) + ": add at least one scene reference");
        return null;
      }
      p.scene_references = step._scenes.map(function (ref) {
        return { name: ref.name, timestamp: ref.timestamp, threshold: numberOrDefault(ref.threshold, 0.75) };
      });
    } else if (stepType === "inactivity") {
      p.threshold = intOrDefault((qs("#paramInactThresh" + sfx) || {}).value, 10);
    }
    var stepRegionRef = normalizeRegionRef(state.multitoolSteps[idx].region_ref)
      || (state.multitoolSteps[idx].region ? activeRegionRef(state.multitoolSteps[idx].region) : null);
    p.region = stepRegionRef ? stepRegionRef.name : "";
    if (stepRegionRef) p.region_ref = regionRefPayload(stepRegionRef);
    return p;
  }

  function gatherWorkflowParams(type) {
    var params = {};
    if (type === "multitool") {
      if (state.multitoolSteps.length < 2) {
        showToast("Add at least 2 steps");
        return null;
      }
      params.steps = [];
      for (var i = 0; i < state.multitoolSteps.length; i++) {
        var stepP = gatherMultitoolStepParams(state.multitoolSteps[i].type, i);
        if (stepP === null) return null;
        stepP.type = state.multitoolSteps[i].type;
        if (i > 0) {
          stepP.logic = (state.multitoolSteps[i].logic || "AND").toUpperCase();
        }
        params.steps.push(stepP);
      }
      params.interval = numberOrDefault((qs("#paramMultitoolInterval") || {}).value, 1.0);
      var mtLabelEl = qs("#paramEventLabel");
      if (mtLabelEl && mtLabelEl.value.trim()) params.event_label = mtLabelEl.value.trim();
      var mtDfEl = qs("#paramDetectFirst");
      if (mtDfEl && mtDfEl.checked) params.detect_first = true;
      return params;
    } else if (type === "color") {
      params.target_color = {
        h: numberOrDefault((qs("#paramColorH") || {}).value, 0),
        s: numberOrDefault((qs("#paramColorS") || {}).value, 0),
        v: numberOrDefault((qs("#paramColorV") || {}).value, 0),
      };
      var tol = numberOrDefault((qs("#paramColorTol") || {}).value, 30);
      params.tolerance = {
        h: Math.round(tol * 90 / 100),
        s: Math.round(tol * 128 / 100),
        v: Math.round(tol * 128 / 100),
      };
      params.interval = numberOrDefault((qs("#paramColorInterval") || {}).value, 1.0);
    } else if (type === "change") {
      params.threshold = numberOrDefault((qs("#paramChangeThresh") || {}).value, 0.03);
      params.noise_threshold = intOrDefault((qs("#paramChangeNoise") || {}).value, 30);
      params.interval = numberOrDefault((qs("#paramChangeInterval") || {}).value, 1.0);
    } else if (type === "similarity") {
      if (state.referenceTimestamp === null) {
        showToast("Capture a reference frame first");
        return null;
      }
      params.reference_timestamp = state.referenceTimestamp;
      params.threshold = numberOrDefault((qs("#paramSimThresh") || {}).value, 0.90);
      params.interval = numberOrDefault((qs("#paramSimInterval") || {}).value, 1.0);
    } else if (type === "text") {
      params.search_string = (qs("#paramTextSearch") || {}).value || "";
      if (!params.search_string.trim()) {
        showToast("Enter a search string");
        return null;
      }
      params.fuzzy_threshold = numberOrDefault((qs("#paramTextFuzzy") || {}).value, 0.80);
      params.interval = numberOrDefault((qs("#paramTextInterval") || {}).value, 2.0);
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
      params.interval = numberOrDefault((qs("#paramNumInterval") || {}).value, 2.0);
    } else if (type === "timelapse") {
      params.speedup_factor = numberOrDefault((qs("#paramTlSpeed") || {}).value, 10);
      var si = parseFloat((qs("#paramTlSampleInterval") || {}).value);
      if (si > 0) params.sample_interval = si;
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
      params.threshold = numberOrDefault((qs("#paramTemplateThresh") || {}).value, 0.70);
      params.interval = numberOrDefault((qs("#paramTemplateInterval") || {}).value, 1.0);
      var scalePct = parseFloat((qs("#paramTemplateScale") || {}).value);
      if (!isNaN(scalePct) && scalePct > 0 && scalePct !== 100) {
        params.template_scale = scalePct / 100;
      }
    } else if (type === "flow") {
      params.magnitude_threshold = numberOrDefault((qs("#paramFlowMag") || {}).value, 2.0);
      params.interval = numberOrDefault((qs("#paramFlowInterval") || {}).value, 1.0);
    } else if (type === "scene") {
      if (state.sceneReferences.length === 0) {
        showToast("Add at least one scene reference");
        return null;
      }
      params.scene_references = state.sceneReferences.map(function (ref) {
        return { name: ref.name, timestamp: ref.timestamp, threshold: numberOrDefault(ref.threshold, 0.75) };
      });
      params.interval = numberOrDefault((qs("#paramSceneInterval") || {}).value, 1.0);
    } else if (type === "inactivity") {
      params.threshold = intOrDefault((qs("#paramInactThresh") || {}).value, 10);
      params.min_duration = numberOrDefault((qs("#paramInactMinDur") || {}).value, 2.0);
      params.interval = numberOrDefault((qs("#paramInactInterval") || {}).value, 1.0);
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

  var TASK_TYPE_ICON_FILES = {
    multitool: "link",
    color: "eye-dropper",
    change: "bolt",
    similarity: "photo",
    text: "language",
    numbers: "hashtag",
    timelapse: "forward",
    template: "viewfinder-circle",
    flow: "arrows-right-left",
    scene: "squares-2x2",
    inactivity: "pause-circle",
  };

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

  var TOOL_LABELS = {
    multitool: "Multitool",
    color: "Color",
    change: "Change",
    similarity: "Similarity",
    text: "Text",
    numbers: "Numbers",
    timelapse: "Timelapse",
    template: "Template",
    flow: "Flow",
    scene: "Scene",
    inactivity: "Inactivity",
  };

  function selectableTasks() {
    return state.tasks.filter(function (t) {
      return t.status === "completed" || t.status === "paused" || t.status === "running";
    });
  }

  function setRightPaneTab(tab) {
    state.rightPaneTab = tab;
    setStoredUIStateField("screenspace", "rightPaneTab", tab);
    qsa("#rightPaneTabs .rp-tab").forEach(function (b) {
      b.classList.toggle("active", b.dataset.tab === tab);
    });
    var qp = qs("#taskQueuePanel");
    var rp = qs("#resultsPanel");
    if (qp) qp.classList.toggle("hidden", tab !== "queue");
    if (rp) rp.classList.toggle("hidden", tab !== "results");
    qsa("#rightPaneTabs .rp-tab-actions").forEach(function (a) {
      a.classList.toggle("hidden", a.dataset.for !== tab);
    });
    closeResultsSwitcher();
  }

  function updateResultsCrumb() {
    var crumbEl = qs("#resultsTabCrumb");
    if (!crumbEl) return;
    var task = state.selectedTaskId ? findTask(state.selectedTaskId) : null;
    if (!task) {
      crumbEl.textContent = "";
      crumbEl.style.color = "";
      return;
    }
    var sameType = state.tasks
      .filter(function (t) {
        return t.type === task.type &&
          (t.status === "completed" || t.status === "paused" || t.status === "running");
      });
    var idx = -1;
    for (var i = 0; i < sameType.length; i++) {
      if (sameType[i].id === task.id) { idx = i; break; }
    }
    var label = TOOL_LABELS[task.type] || task.type;
    var ordinal = idx >= 0 ? idx + 1 : 1;
    var participant = task.participant || "";
    crumbEl.textContent = ": " + label + " " + ordinal + (participant ? " \u00b7 " + participant : "");
    crumbEl.style.color = taskTypeColor(task.type);
  }

  function openResultsSwitcher() {
    var panel = qs("#resultsSwitcherPanel");
    if (!panel) return;
    panel.innerHTML = "";
    var tasks = selectableTasks();
    if (tasks.length === 0) {
      var empty = el("div", "rp-switcher-empty", "No completed tasks yet.");
      panel.appendChild(empty);
    } else {
      var frag = document.createDocumentFragment();
      tasks.forEach(function (t) {
        var item = el("button", "rp-switcher-item");
        item.type = "button";
        item.dataset.taskId = t.id;
        if (t.id === state.selectedTaskId) item.classList.add("active");
        var badge = el("span", "rp-switcher-item-badge");
        badge.style.background = taskTypeColor(t.type);
        item.appendChild(badge);
        var label = TOOL_LABELS[t.type] || t.type;
        var primary = el("span", null, label + " \u00b7 " + (t.participant || ""));
        item.appendChild(primary);
        if (t.region) {
          var meta = el("span", "rp-switcher-item-meta", t.region);
          item.appendChild(meta);
        }
        item.addEventListener("click", function (e) {
          e.stopPropagation();
          var taskId = t.id;
          closeResultsSwitcher();
          var task = findTask(taskId);
          if (!task) return;
          if (task.participant && task.participant !== state.selectedParticipant) {
            selectParticipant(task.participant);
          }
          state.selectedTaskId = taskId;
          setRightPaneTab("results");
          loadAndShowResults(taskId);
          renderTaskList();
        });
        frag.appendChild(item);
      });
      panel.appendChild(frag);
    }
    var tab = qs('.rp-tab[data-tab="results"]');
    var tabsEl = qs("#rightPaneTabs");
    if (tab && tabsEl) {
      var tabRect = tab.getBoundingClientRect();
      var tabsRect = tabsEl.getBoundingClientRect();
      var tabCenter = tabRect.left - tabsRect.left + tabRect.width / 2;
      panel.style.left = tabCenter + "px";
      panel.style.transform = "translateX(-50%)";
    }
    panel.classList.remove("hidden");
    state.resultsSwitcherOpen = true;
    if (tab) tab.classList.add("switcher-open");
  }

  function closeResultsSwitcher() {
    var panel = qs("#resultsSwitcherPanel");
    if (panel) panel.classList.add("hidden");
    state.resultsSwitcherOpen = false;
    var tab = qs('.rp-tab[data-tab="results"]');
    if (tab) tab.classList.remove("switcher-open");
  }

  function initRightPaneTabs() {
    qsa("#rightPaneTabs .rp-tab").forEach(function (btn) {
      btn.addEventListener("click", function (e) {
        e.stopPropagation();
        var tab = btn.dataset.tab;
        if (tab === "queue") {
          setRightPaneTab("queue");
          return;
        }
        if (tab === "results") {
          if (state.rightPaneTab === "results" && state.selectedTaskId) {
            if (state.resultsSwitcherOpen) closeResultsSwitcher();
            else openResultsSwitcher();
          } else {
            setRightPaneTab("results");
          }
        }
      });
    });
    document.addEventListener("click", function (e) {
      if (!state.resultsSwitcherOpen) return;
      if (e.target.closest("#resultsSwitcherPanel")) return;
      if (e.target.closest('.rp-tab[data-tab="results"]')) return;
      closeResultsSwitcher();
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
                setRightPaneTab("queue");
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
      task = findTask(taskId);
      if (task && (task.status === "completed" || task.status === "paused" || task.status === "running")) {
        if (state.selectedTaskId === taskId) {
          _resultsRequestVersion += 1;
          state.selectedTaskId = null;
          state.selectedTaskResults = null;
          renderResults();
          renderTaskList();
          renderTimeline();
          updateResultsCrumb();
          setRightPaneTab("queue");
        } else {
          if (task.participant && task.participant !== state.selectedParticipant) {
            selectParticipant(task.participant);
          }
          state.selectedTaskId = taskId;
          setRightPaneTab("results");
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
      e.dataTransfer.setData("application/x-task-id", card.dataset.taskId);
      e.dataTransfer.effectAllowed = "move";
      _cacheTaskDragMidpoints(taskListEl);
    });

    taskListEl.addEventListener("dragend", function (e) {
      var card = e.target.closest(".task-card");
      if (card) {
        card.classList.remove("dragging");
        card.removeAttribute("draggable");
      }
      if (_taskListDragOverRaf != null) {
        cancelAnimationFrame(_taskListDragOverRaf);
        _taskListDragOverRaf = null;
      }
      _taskListPendingDragOver = null;
      clearDragIndicators(taskListEl);
      _taskDragCache = null;
    });

    taskListEl.addEventListener("dragover", function (e) {
      if (e.dataTransfer.types.indexOf("text/plain") < 0) return;

      var cards = taskListEl.querySelectorAll(".task-card:not(.dragging)");
      var insertIdx = getDropIndex(taskListEl, e.clientY);

      // Determine boundary between finished (completed/failed) and queued zones
      var finishedCount = 0;
      for (var i = 0; i < cards.length; i++) {
        var t = findTask(cards[i].dataset.taskId);
        if (t && (t.status === "completed" || t.status === "failed")) finishedCount++;
        else break;
      }

      // Find the dragging card to determine its status (we can't read the
      // status payload off dataTransfer during dragover for security reasons).
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

      _taskListPendingDragOver = { insertIdx: insertIdx };
      if (_taskListDragOverRaf != null) return;
      _taskListDragOverRaf = requestAnimationFrame(function () {
        _taskListDragOverRaf = null;
        var pending = _taskListPendingDragOver;
        if (!pending) return;
        var idx = pending.insertIdx;
        var cardsNow = taskListEl.querySelectorAll(".task-card:not(.dragging)");
        clearDragIndicators(taskListEl);
        if (idx < cardsNow.length) {
          cardsNow[idx].classList.add("drag-over");
        } else {
          taskListEl.classList.add("drag-over-append");
        }
      });
    });

    taskListEl.addEventListener("dragleave", function (e) {
      var card = e.target.closest(".task-card");
      if (card) card.classList.remove("drag-over");
      if (!taskListEl.contains(e.relatedTarget)) {
        taskListEl.classList.remove("drag-over-append");
      }
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

  // Cached at dragstart: { all: number[], statusGrouped: { queued: number[], finished: number[] } }
  var _taskDragCache = null;
  var _taskListDragOverRaf = null;
  var _taskListPendingDragOver = null;

  function _cacheTaskDragMidpoints(container) {
    var cards = container.querySelectorAll(".task-card:not(.dragging)");
    var all = new Array(cards.length);
    var queued = [];
    var finished = [];
    for (var i = 0; i < cards.length; i++) {
      var r = cards[i].getBoundingClientRect();
      var mid = r.top + r.height / 2;
      all[i] = mid;
      var t = findTask(cards[i].dataset.taskId);
      if (!t) continue;
      if (t.status === "queued") queued.push(mid);
      else if (t.status === "completed" || t.status === "failed") finished.push(mid);
    }
    _taskDragCache = { all: all, queued: queued, finished: finished };
  }

  function getDropIndex(container, clientY) {
    if (!_taskDragCache) _cacheTaskDragMidpoints(container);
    var mids = _taskDragCache.all;
    for (var i = 0; i < mids.length; i++) {
      if (clientY < mids[i]) return i;
    }
    return mids.length;
  }

  function getDropIndexAmongStatus(container, clientY, group) {
    if (!_taskDragCache) _cacheTaskDragMidpoints(container);
    var mids = group === "queued" ? _taskDragCache.queued : _taskDragCache.finished;
    for (var i = 0; i < mids.length; i++) {
      if (clientY < mids[i]) return i;
    }
    return mids.length;
  }

  function clearDragIndicators(container) {
    var cards = container.querySelectorAll(".task-card.drag-over");
    for (var i = 0; i < cards.length; i++) cards[i].classList.remove("drag-over");
    container.classList.remove("drag-over-append");
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
    if (task.region_ref) {
      var restoredRef = normalizeRegionRef(task.region_ref);
      state.runRegions = restoredRef ? [restoredRef] : [];
      state.pendingRegion = null;
      if (restoredRef && restoredRef.source === "active" && state.regions[restoredRef.name]) {
        state.activeRegion = restoredRef.name;
      } else {
        state.activeRegion = null;
      }
      renderRegionChips();
      renderRunRegionPicker();
      renderOverlay();
      updateRegionButtons();
    } else if (task.region && state.regions[task.region]) {
      state.activeRegion = task.region;
      state.pendingRegion = null;
      state.runRegions = [activeRegionRef(task.region)];
      renderRegionChips();
      renderRunRegionPicker();
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

    // For scene, restore references into state before rendering so the list shows them.
    if (task.type === "scene") {
      var sceneParams = task.parameters || {};
      state.sceneReferences = (sceneParams.scene_references || []).map(function (ref) {
        return { name: ref.name, timestamp: ref.timestamp, threshold: numberOrDefault(ref.threshold, 0.75) };
      });
    }

    // For multitool, rebuild steps state before rendering. `_initial` carries the saved
    // per-step config so the _mtRender* functions can set input values at element creation
    // time (rather than via a post-render setInputValue pass).
    if (task.type === "multitool") {
      var mtParams = task.parameters || {};
      state.multitoolSteps = (mtParams.steps || []).map(function (s) {
        var step = { type: s.type, collapsed: true };
        step.logic = (s.logic || "AND").toUpperCase();
        if (s.region) step.region = s.region;
        if (s.region_ref) step.region_ref = s.region_ref;
        if (s.reference_timestamp !== undefined) step._refTs = s.reference_timestamp;
        if (s.scene_references) step._scenes = s.scene_references.map(function (ref) {
          return { name: ref.name, timestamp: ref.timestamp, threshold: numberOrDefault(ref.threshold, 0.75) };
        });
        step._initial = s;
        return step;
      });
    }

    // Rebuild param controls then set values
    renderWorkflowParams();

    params = task.parameters || {};
    if (task.type === "multitool") {
      setInputValue("#paramMultitoolInterval", numberOrDefault(params.interval, 1.0));
    } else if (task.type === "color") {
      var tc = params.target_color || {};
      var ch = numberOrDefault(tc.h, 90);
      var cs = numberOrDefault(tc.s, 200);
      var cv = numberOrDefault(tc.v, 200);
      var savedTol = params.tolerance ? Math.round(params.tolerance.h * 100 / 90) : 30;
      setInputValue("#paramColorTol", savedTol);
      setInputValue("#paramColorInterval", numberOrDefault(params.interval, 1.0));
      // setTargetColor writes hidden h/s/v + hex input + preview + palette + brightness strip.
      setTargetColor(ch, cs, cv);
    } else if (task.type === "change") {
      setInputValue("#paramChangeThresh", numberOrDefault(params.threshold, 0.03));
      setInputValue("#paramChangeNoise", intOrDefault(params.noise_threshold, 30));
      setInputValue("#paramChangeInterval", numberOrDefault(params.interval, 1.0));
    } else if (task.type === "similarity") {
      setInputValue("#paramSimThresh", numberOrDefault(params.threshold, 0.90));
      setInputValue("#paramSimInterval", numberOrDefault(params.interval, 1.0));
    } else if (task.type === "text") {
      setInputValue("#paramTextSearch", params.search_string || "");
      setInputValue("#paramTextFuzzy", numberOrDefault(params.fuzzy_threshold, 0.80));
      setInputValue("#paramTextInterval", numberOrDefault(params.interval, 2.0));
      if (params.languages && params.languages[0]) {
        setInputValue("#paramTextLang", params.languages[0]);
      }
    } else if (task.type === "numbers") {
      setInputValue("#paramNumOperator", params.operator || "gt");
      // Fire change so range/target row visibility tracks the restored operator
      // (listener attached in renderNumbersParams).
      var opSel = qs("#paramNumOperator");
      if (opSel) opSel.dispatchEvent(new Event("change"));
      if (params.operator === "range") {
        setInputValue("#paramNumMin", numberOrDefault(params.range_min, 0));
        setInputValue("#paramNumMax", numberOrDefault(params.range_max, 100));
      } else {
        setInputValue("#paramNumTarget", numberOrDefault(params.target_value, 100));
      }
      setInputValue("#paramNumInterval", numberOrDefault(params.interval, 2.0));
    } else if (task.type === "timelapse") {
      setInputValue("#paramTlSpeed", numberOrDefault(params.speedup_factor, 10));
      setInputValue("#paramTlFormat", params.output_format || "mp4");
      if (params.sample_interval !== undefined) {
        setInputValue("#paramTlSampleInterval", params.sample_interval);
      }
    } else if (task.type === "template") {
      if (params.reference_timestamp !== undefined) {
        state.referenceTimestamp = params.reference_timestamp;
      }
      setInputValue("#paramTemplateThresh", numberOrDefault(params.threshold, 0.70));
      setInputValue("#paramTemplateInterval", numberOrDefault(params.interval, 1.0));
      if (params.template_scale) {
        setInputValue("#paramTemplateScale", Math.round(params.template_scale * 100));
      }
    } else if (task.type === "flow") {
      setInputValue("#paramFlowMag", numberOrDefault(params.magnitude_threshold, 2.0));
      setInputValue("#paramFlowInterval", numberOrDefault(params.interval, 1.0));
    } else if (task.type === "scene") {
      setInputValue("#paramSceneInterval", numberOrDefault(params.interval, 1.0));
    } else if (task.type === "inactivity") {
      setInputValue("#paramInactThresh", intOrDefault(params.threshold, 10));
      setInputValue("#paramInactMinDur", numberOrDefault(params.min_duration, 2.0));
      setInputValue("#paramInactInterval", numberOrDefault(params.interval, 1.0));
    }

    // event_label and detect_first apply to every non-timelapse task type
    // (see gatherWorkflowParams for the symmetric save path).
    if (task.type !== "timelapse") {
      if (params.event_label) setInputValue("#paramEventLabel", params.event_label);
      if (params.detect_first) {
        var dfEl = qs("#paramDetectFirst");
        if (dfEl) dfEl.checked = true;
      }
    }

    if (state.restoreMarkersOnEdit) {
      var hasIn = params.start_seconds !== undefined && params.start_seconds !== null;
      var hasOut = params.end_seconds !== undefined && params.end_seconds !== null;
      if (hasIn || hasOut) {
        state.inMarker = hasIn ? params.start_seconds : null;
        state.outMarker = hasOut ? params.end_seconds : null;
        updateMarkerInfo();
        renderTimeline();
      }
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
      btn.appendChild(iconSpan("play"));
      btn.title = "Resume queue";
    } else {
      btn.appendChild(iconSpan("pause"));
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

  // ---- Task list: filter chips + card DOM (drag reorder uses _taskDragCache) ----

  function initTaskFilters() {
    var doneBtn = qs("#taskFilterDoneBtn");
    var failedBtn = qs("#taskFilterFailedBtn");
    if (doneBtn) {
      doneBtn.appendChild(iconSpan("check"));
      doneBtn.addEventListener("click", function () { toggleTaskFilter("completed"); });
    }
    if (failedBtn) {
      failedBtn.appendChild(iconSpan("x-mark"));
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
      container.innerHTML = "";
      container.appendChild(el("div", "panel-empty", "No tasks yet. Configure a workflow and click Run."));
      return;
    }
    if (filtered.length === 0) {
      container.innerHTML = "";
      container.appendChild(el("div", "panel-empty", "No " + state.taskFilter + " tasks."));
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
        handle.appendChild(iconSpan("bars-2"));
        handle.addEventListener("mousedown", function () { card.setAttribute("draggable", "true"); });
        handle.addEventListener("mouseup", function () { card.removeAttribute("draggable"); });
        card.appendChild(handle);
      } else if (task.status === "running") {
        card.appendChild(el("span", "task-card-spinner"));
      } else if (task.status === "paused") {
        var pauseIcon = el("span", "task-card-pause-icon");
        pauseIcon.appendChild(iconSpan("pause", "ss-icon--xs"));
        card.appendChild(pauseIcon);
      }

      // Type badge
      var badge = el("span", "task-card-type");
      badge.style.color = taskTypeColor(task.type);
      badge.title = task.type;
      var typeIconEl = el("span", "task-card-type-icon");
      var iconFile = TASK_TYPE_ICON_FILES[task.type] || "squares-2x2";
      applyMaskIcon(typeIconEl, 'url("/screenspace/icons/' + iconFile + '.svg")');
      badge.appendChild(typeIconEl);
      card.appendChild(badge);

      // Fast scan badge
      if ((task.parameters || {}).scan_mode === "fast") {
        var fb = el("span", "task-fast-badge");
        var bi = el("span", "task-fast-badge-icon");
        applyMaskIcon(bi, 'url("/screenspace/icons/chevron-double-right.svg")');
        fb.appendChild(bi);
        fb.appendChild(document.createTextNode("Fast"));
        card.appendChild(fb);
      }

      // Info
      var info = el("div", "task-card-info");
      var meta = el("span", "task-card-meta");
      var eventLabel = (task.parameters || {}).event_label;
      if (eventLabel) {
        meta.textContent = eventLabel;
      } else {
        meta.textContent = task.participant + " \u00b7 " + (task.region || "");
      }
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
        rLen = Array.isArray(task.result) ? task.result.length : (typeof task.result === "string" ? 1 : 0);
        statusText = rLen + " result" + (rLen !== 1 ? "s" : "");
      }
      card.appendChild(el("span", "task-card-status", statusText));

      // Edit button
      var editBtn = el("button", "task-card-edit");
      editBtn.title = "Edit";
      editBtn.appendChild(iconSpan("pencil-square"));
      card.appendChild(editBtn);

      // Dismiss button
      var dismissBtn = el("button", "task-card-dismiss");
      dismissBtn.title = "Dismiss";
      dismissBtn.appendChild(iconSpan("x-mark"));
      card.appendChild(dismissBtn);

      frag.appendChild(card);
    });
    var prevScrollTop = container.scrollTop;
    container.innerHTML = "";
    container.appendChild(frag);
    container.scrollTop = prevScrollTop;
    updateResultsCrumb();
  }

  // ---- SSE (Server-Sent Events) with polling fallback ----

  function handleTaskData(data) {
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
  }

  function startSSE() {
    if (state.eventSource) return;
    var es = new EventSource("api/tasks/stream");
    state.eventSource = es;

    es.onmessage = function (e) {
      var data;
      try { data = JSON.parse(e.data); } catch (_) { return; }
      handleTaskData(data);
    };

    es.onerror = function () {
      // Connection lost — fall back to polling
      es.close();
      state.eventSource = null;
      startPolling();
    };
  }

  // ---- Polling (fallback) ----

  function startPolling() {
    if (state.pollTimer) return;
    if (document.hidden) return;
    state.pollTimer = setInterval(pollTasks, POLL_INTERVAL);
  }

  function stopPolling() {
    if (state.pollTimer) {
      clearInterval(state.pollTimer);
      state.pollTimer = null;
    }
  }

  function pollTasks() {
    var hasActive = state.tasks.some(function (t) {
      return t.status === "queued" || t.status === "running" || t.status === "paused";
    });
    if (!hasActive) {
      stopPolling();
      return;
    }

    apiGet("api/tasks")
      .then(function (data) { handleTaskData(data); })
      .catch(function () {});
  }

  function stopSSE() {
    if (state.eventSource) {
      state.eventSource.close();
      state.eventSource = null;
    }
  }

  document.addEventListener("dragend", function () {
    if (_multitoolDragOverRaf != null) {
      cancelAnimationFrame(_multitoolDragOverRaf);
      _multitoolDragOverRaf = null;
    }
    _multitoolPendingDragOver = null;
    if (_taskListDragOverRaf != null) {
      cancelAnimationFrame(_taskListDragOverRaf);
      _taskListDragOverRaf = null;
    }
    _taskListPendingDragOver = null;
    var stepsDiv = document.querySelector(".multitool-steps");
    if (stepsDiv) clearMultitoolDragIndicators(stepsDiv);
    var taskListEl = qs("#taskList");
    if (taskListEl) clearDragIndicators(taskListEl);
  });

  document.addEventListener("visibilitychange", function () {
    if (document.hidden) {
      stopSSE();
      stopPolling();
      return;
    }
    var hasActive = state.tasks.some(function (t) {
      return t.status === "queued" || t.status === "running" || t.status === "paused";
    });
    if (hasActive) startSSE();
  });

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
        state.resultOverlay = { type: "flow", data: rData, region: taskRegionPixels(task) };
      } else {
        state.resultOverlay = null;
      }
      loadFrame(ts);
    });

    // Hover result rows to highlight matching scene markers on timeline
    qs("#resultsList").addEventListener("mouseover", function (e) {
      var row = e.target.closest(".result-row");
      if (!row) return;
      var task = state.selectedTaskId ? findTask(state.selectedTaskId) : null;
      if (!task || task.type !== "scene") return;
      var ri = parseInt(row.dataset.resultIndex, 10);
      var results = state.selectedTaskResults || [];
      if (isNaN(ri) || ri >= results.length) return;
      var sceneName = results[ri].scene_name;
      if (sceneName !== state.hoveredResultSceneName) {
        state.hoveredResultSceneName = sceneName;
        renderTimeline();
      }
    });

    qs("#resultsList").addEventListener("mouseleave", function () {
      if (state.hoveredResultSceneName !== null) {
        state.hoveredResultSceneName = null;
        renderTimeline();
      }
    });

    var showExcludedBtn = qs("#showExcludedBtn");
    function updateShowExcludedIcon() {
      var iconSpan = showExcludedBtn.querySelector(".rp-icon-btn-icon");
      iconSpan.classList.toggle("rp-icon-eye", state.showExcluded);
      iconSpan.classList.toggle("rp-icon-eye-slash", !state.showExcluded);
      showExcludedBtn.classList.toggle("active", state.showExcluded);
    }
    updateShowExcludedIcon();
    showExcludedBtn.addEventListener("click", function () {
      state.showExcluded = !state.showExcluded;
      updateShowExcludedIcon();
      renderResults();
    });
    attachHoverTooltip(showExcludedBtn, function () {
      return state.showExcluded ? "Hiding excluded results is off" : "Hiding excluded results is on";
    }, { align: "center" });

    function downloadEventsExport(format) {
      var url = "api/export/events?format=" + encodeURIComponent(format);
      if (!state.showExcluded) url += "&excluded=false";
      var a = document.createElement("a");
      a.href = url;
      var ext = format === "csv" ? "csv" : "json";
      a.download = "screenspace_events." + ext;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }

    var exportBtn = qs("#exportEventsBtn");
    var exportMenu = qs("#exportEventsMenu");
    if (exportBtn && exportMenu) {
      attachHoverTooltip(exportBtn, "Export events", { align: "center" });
      var closeExportMenu = function () {
        exportMenu.classList.add("hidden");
        exportBtn.setAttribute("aria-expanded", "false");
      };
      exportBtn.addEventListener("click", function (e) {
        e.stopPropagation();
        var open = exportMenu.classList.toggle("hidden");
        exportBtn.setAttribute("aria-expanded", open ? "false" : "true");
      });
      exportMenu.addEventListener("click", function (e) {
        var item = e.target.closest(".rp-export-item");
        if (!item) return;
        e.stopPropagation();
        downloadEventsExport(item.dataset.format);
        closeExportMenu();
      });
      document.addEventListener("click", function (e) {
        if (exportMenu.classList.contains("hidden")) return;
        if (e.target.closest("#exportEventsWrap")) return;
        closeExportMenu();
      });
    }

    var certaintySlider = qs("#certaintyCutoff");
    certaintySlider.addEventListener("input", function () {
      state.certaintyCutoff = parseInt(this.value, 10) / 100;
      renderResults();
    });
    attachHoverTooltip(certaintySlider, function () {
      return "Certainty threshold: " + certaintySlider.value + "%";
    }, { align: "center" });

    var exclBtn = qs("#excludeNonVisibleBtn");
    attachHoverTooltip(exclBtn, "Exclude results below the certainty threshold", { align: "center" });

    qs("#excludeNonVisibleBtn").addEventListener("click", function () {
      var events = state.taskEvents[state.selectedTaskId] || [];
      var task = state.selectedTaskId ? findTask(state.selectedTaskId) : null;
      if (!task || !events.length || state.certaintyCutoff <= 0) {
        showToast("Set a certainty threshold first");
        return;
      }
      var idsToExclude = [];
      events.forEach(function (ev) {
        if (!ev.excluded && ev.confidence < state.certaintyCutoff) {
          idsToExclude.push(ev.id);
        }
      });
      if (idsToExclude.length === 0) {
        showToast("No events below threshold");
        return;
      }
      apiPut("api/events/bulk-exclude", { ids: idsToExclude }).then(function () {
        idsToExclude.forEach(function (id) {
          for (var i = 0; i < events.length; i++) {
            if (events[i].id === id) { events[i].excluded = true; break; }
          }
        });
        renderResults();
        renderTimeline();
        showToast("Excluded " + clipgenPluralUnit(idsToExclude.length, "event", "events"));
      });
    });
  }

  function loadAndShowResults(taskId) {
    var resultsRequestVersion = ++_resultsRequestVersion;
    var selectedTaskId = taskId;
    _heatmapOverlayRequestVersion += 1;
    state.resultOverlay = null;
    state.heatmapOverlay = null;
    state.hoveredResultSceneName = null;
    state.certaintyCutoff = 0;
    var slider = qs("#certaintyCutoff");
    if (slider) slider.value = "0";
    apiGet("api/tasks/" + taskId + "/results")
      .then(function (data) {
        if (resultsRequestVersion !== _resultsRequestVersion || state.selectedTaskId !== selectedTaskId) return null;
        state.selectedTaskResults = data.results;
        return apiGet("api/events?task_id=" + selectedTaskId);
      })
      .then(function (evData) {
        if (!evData) return;
        if (resultsRequestVersion !== _resultsRequestVersion || state.selectedTaskId !== selectedTaskId) return;
        state.taskEvents[selectedTaskId] = evData.events || [];
        renderResults();
        renderTaskList();
        updateResultsCrumb();
      })
      .catch(function () { showToast("Failed to load results"); });
  }

  function renderResults() {
    var container = qs("#resultsList");
    var prevResultsScrollTop = container.scrollTop;
    var countEl = qs("#resultCount") || { textContent: "" };
    var actionsEl = qs("#resultsActions");
    var results = state.selectedTaskResults;
    var task = state.selectedTaskId ? findTask(state.selectedTaskId) : null;

    // Manage fast scan label — between panel-header and resultsList
    var fastLabel = qs("#fastScanLabel");
    if (!fastLabel) {
      fastLabel = el("div", "fast-scan-label hidden");
      fastLabel.id = "fastScanLabel";
      container.parentNode.insertBefore(fastLabel, container);
    }

    if (!results || !task) {
      container.innerHTML = '<div class="panel-empty">Click a task to view results.</div>';
      countEl.textContent = "";
      actionsEl.classList.add("hidden");
      fastLabel.classList.add("hidden");
      return;
    }

    actionsEl.classList.remove("hidden");

    if ((task.parameters || {}).scan_mode === "fast") {
      fastLabel.classList.remove("hidden");
      fastLabel.innerHTML = "";
      var fIcon = el("span", "fast-scan-label-icon");
      applyMaskIcon(fIcon, 'url("/screenspace/icons/chevron-double-right.svg")');
      fastLabel.appendChild(fIcon);
      fastLabel.appendChild(document.createTextNode("Fast scan results"));
      var rerunBtn = el("button", "ss-btn ss-btn-sm fast-scan-rerun-btn", "Re-Run Normal");
      (function (t) {
        rerunBtn.addEventListener("click", function () {
          var params = {};
          Object.keys(t.parameters || {}).forEach(function (k) { params[k] = t.parameters[k]; });
          delete params.scan_mode;
          var body = {
            type: t.type,
            participant: t.participant,
            region: t.region || "",
            parameters: params,
          };
          if (t.region_ref) body.region_ref = t.region_ref;
          apiPost("api/tasks", body).then(function (data) {
            if (data.ok) {
              state.tasks.push(data.task);
              renderTaskList();
              startSSE();
              showToast("Re-queued in Normal mode");
            } else {
              showToast(data.error || "Failed to re-queue task");
            }
          });
        });
      })(task);
      fastLabel.appendChild(rerunBtn);
    } else {
      fastLabel.classList.add("hidden");
    }

    // Show/hide certainty controls based on whether the tool has confidence scores
    var hasConf = task && { change: 1, similarity: 1, text: 1, template: 1, scene: 1, flow: 1, multitool: 1, inactivity: 1 }[task.type];
    var certWrap = qs("#certaintyCutoffWrap");
    var exclBtn = qs("#excludeNonVisibleBtn");
    if (certWrap) certWrap.classList.toggle("hidden", !hasConf);
    if (exclBtn) exclBtn.classList.toggle("hidden", !hasConf);

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

    var showToggle = qs("#showExcludedBtn");
    if (showToggle) showToggle.classList.toggle("hidden", events.length === 0);

    var frag = document.createDocumentFragment();

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
          _heatmapOverlayRequestVersion += 1;
          state.heatmapOverlay = null;
          overlayBtn.textContent = "Overlay on Frame";
          renderOverlay();
        } else {
          var overlaySrc = "media/" + task.heatmap;
          var overlayRequestVersion = ++_heatmapOverlayRequestVersion;
          state.heatmapOverlay = {
            src: overlaySrc,
            type: task.type,
            region_coords: taskRegionPixels(task),
          };
          overlayBtn.textContent = "Hide Overlay";
          var hmImg = new Image();
          hmImg.onload = function () {
            if (
              overlayRequestVersion === _heatmapOverlayRequestVersion
              && state.heatmapOverlay
              && state.heatmapOverlay.src === overlaySrc
            ) {
              state.heatmapOverlay._img = hmImg;
              renderOverlay();
            }
          };
          hmImg.src = overlaySrc;
        }
      });
      heatmapLabel.appendChild(overlayBtn);

      if (task.heatmap_gif) {
        var animBtn = el("button", "ss-btn ss-btn-sm", "Show Animation");
        var showingGif = false;
        animBtn.addEventListener("click", function () {
          showingGif = !showingGif;
          heatmapStaticImg.classList.toggle("hidden", showingGif);
          heatmapGifImg.classList.toggle("hidden", !showingGif);
          animBtn.textContent = showingGif ? "Show Static" : "Show Animation";
        });
        heatmapLabel.appendChild(animBtn);
      }

      heatmapSection.appendChild(heatmapLabel);
      var heatmapStaticImg = document.createElement("img");
      heatmapStaticImg.src = "media/" + task.heatmap;
      heatmapStaticImg.alt = "Detection heatmap";
      heatmapSection.appendChild(heatmapStaticImg);

      if (task.heatmap_gif) {
        var heatmapGifImg = document.createElement("img");
        heatmapGifImg.src = "media/" + task.heatmap_gif;
        heatmapGifImg.alt = "Heatmap accumulation animation";
        heatmapGifImg.className = "hidden";
        heatmapSection.appendChild(heatmapGifImg);
      }

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

      // Certainty filtering
      if (hasConf && state.certaintyCutoff > 0) {
        var confValue = null;
        if (task.type === "change") confValue = r.magnitude;
        else if (task.type === "similarity") confValue = r.score;
        else if (task.type === "text") confValue = r.confidence;
        else if (task.type === "template") confValue = r.best_score;
        else if (task.type === "flow") confValue = Math.min(r.magnitude / 10, 1);
        else if (task.type === "scene") confValue = r.score;
        else if (task.type === "multitool") confValue = r.min_confidence;
        else if (task.type === "inactivity") confValue = Math.min((r.duration || 0) / 30, 1);
        if (confValue !== null && confValue < state.certaintyCutoff) return;
      }

      var isExcluded = matchedEvent && matchedEvent.excluded;
      if (isExcluded && !state.showExcluded) return;

      var row = el("div", "result-row" + (isExcluded ? " excluded" : ""));
      row.dataset.resultIndex = rIdx;

      if (task.type === "color") {
        row.dataset.timestamp = r.start;
        row.appendChild(el("span", "result-timestamp", formatTime(r.start, { decimals: 1 }) + " \u2013 " + formatTime(r.end, { decimals: 1 })));
        row.appendChild(el("span", "result-detail", r.duration.toFixed(1) + "s"));
      } else if (task.type === "inactivity") {
        row.dataset.timestamp = r.start;
        row.appendChild(el("span", "result-timestamp", formatTime(r.start, { decimals: 1 }) + " \u2013 " + formatTime(r.end, { decimals: 1 })));
        row.appendChild(el("span", "result-detail", r.duration.toFixed(1) + "s"));
        row.appendChild(el("span", "result-score", "d:" + (r.avg_distance !== undefined ? r.avg_distance : "?")));
      } else if (task.type === "change") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
        row.appendChild(buildConfBar(Math.min(r.magnitude, 1), task.type));
        row.appendChild(el("span", "result-score", (r.magnitude * 100).toFixed(1) + "%"));
      } else if (task.type === "similarity") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
        row.appendChild(buildConfBar(r.score, task.type));
        row.appendChild(el("span", "result-score", (r.score * 100).toFixed(1) + "%"));
      } else if (task.type === "text") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
        row.appendChild(el("span", "result-detail", r.text_found || ""));
        row.appendChild(buildConfBar(r.confidence, task.type));
        row.appendChild(el("span", "result-score", (r.confidence * 100).toFixed(0) + "%"));
      } else if (task.type === "numbers") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
        row.appendChild(el("span", "result-detail", String(r.number_found)));
      } else if (task.type === "template") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
        row.appendChild(buildConfBar(r.best_score, task.type));
        row.appendChild(el("span", "result-score", (r.best_score * 100).toFixed(1) + "%"));
        row.appendChild(el("span", "result-detail", r.match_count + " match" + (r.match_count !== 1 ? "es" : "")));
      } else if (task.type === "flow") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
        row.appendChild(buildConfBar(Math.min(r.magnitude / 20, 1), task.type));
        row.appendChild(el("span", "result-score", r.magnitude.toFixed(2)));
      } else if (task.type === "scene") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
        row.appendChild(el("span", "result-detail", r.scene_name));
        row.appendChild(buildConfBar(r.score, task.type));
        row.appendChild(el("span", "result-score", (r.score * 100).toFixed(1) + "%"));
      } else if (task.type === "multitool") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
        var badges = el("span", "result-detail multitool-badges");
        var stepDefs = (task.parameters && task.parameters.steps) || [];
        var types = r.tool_types || stepDefs.map(function (s) { return s.type; });
        types.forEach(function (t, i) {
          var step = stepDefs[i] || { type: t };
          if (i > 0) {
            var logic = (step.logic || "AND").toUpperCase();
            var sep = el("span", "multitool-step-logic" + (logic === "NOT" ? " logic-not" : ""), logic);
            badges.appendChild(sep);
          }
          var badge = el("span", "multitool-type-badge");
          badge.style.color = taskTypeColor(t);
          var paramStr = formatMultitoolStepParams(step);
          badge.title = t + (paramStr ? ": " + paramStr : "");
          var icon = buildTypeIcon(t);
          if (icon) badge.appendChild(icon);
          if (paramStr) badge.appendChild(el("span", "multitool-step-params", paramStr));
          badges.appendChild(badge);
        });
        row.appendChild(badges);
        row.appendChild(el("span", "result-score", ((r.min_confidence || 0) * 100).toFixed(1) + "%"));
      }

      if (matchedEvent) {
        var btn = el("button", "result-exclude-btn");
        btn.dataset.eventId = matchedEvent.id;
        btn.dataset.excluded = isExcluded ? "true" : "false";
        btn.title = isExcluded ? "Include event" : "Exclude event";
        var icon = isExcluded ? iconSpan("x-mark", "ss-icon--sm") : iconSpan("check", "ss-icon--sm");
        btn.appendChild(icon);
        row.appendChild(btn);
      }

      frag.appendChild(row);
    });
    container.appendChild(frag);
    container.scrollTop = prevResultsScrollTop;
  }

  // ---- Keyboard shortcuts ----

  function initKeyboard() {
    function onKeyDown(e) {
      // Don't capture when typing in inputs
      var t = e.target;
      if (t && t.matches && t.matches("input, textarea, select, [contenteditable=true]")) return;

      if (e.key === "ArrowLeft") {
        e.preventDefault();
        if (state.videoInfo) loadFrame(clamp(state.currentTimestamp - FRAME_STEP, 0, Math.max(0, state.videoInfo.duration - 0.001)));
      } else if (e.key === "ArrowRight") {
        e.preventDefault();
        if (state.videoInfo) loadFrame(clamp(state.currentTimestamp + FRAME_STEP, 0, Math.max(0, state.videoInfo.duration - 0.001)));
      } else if (e.key === " ") {
        e.preventDefault();
        if (state.videoPlaying) {
          pauseVideo();
        } else {
          playVideo();
        }
      } else if (e.key === "b" || e.key === "B") {
        if (e.repeat) return;
        if (e.metaKey || e.ctrlKey || e.altKey) return;
        if (!_overlayEligibleForActiveTool()) return;
        e.preventDefault();
        state.overlayBlinkActive = true;
        var curTs = Number(state.currentTimestamp || 0).toFixed(3);
        if (!state.overlayImage || state.overlayImageTimestamp !== curTs || state.overlayImageTool !== state.activeWorkflow) {
          refreshModelView();
        }
        renderOverlay();
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
        } else if (state.drawingRegion) {
          state.drawingRegion = null;
          _cachedOverlayRect = null;
          renderOverlay();
          updateRegionButtons();
        } else if (state.pendingRegion || state.activeRegion) {
          state.pendingRegion = null;
          state.activeRegion = null;
          renderOverlay();
          updateRegionButtons();
          updateRunButton();
        }
        hideRegionNameModal();
      }
    }

    function onKeyUp(e) {
      if (e.key === "b" || e.key === "B") {
        if (state.overlayBlinkActive) {
          state.overlayBlinkActive = false;
          renderOverlay();
        }
      }
    }

    document.addEventListener("keydown", onKeyDown);
    document.addEventListener("keyup", onKeyUp);
    window.addEventListener("pagehide", function () {
      document.removeEventListener("keydown", onKeyDown);
      document.removeEventListener("keyup", onKeyUp);
    });

    // Defensive: clear blink state on window blur so a held key doesn't get
    // stuck on if the user alt-tabs.
    window.addEventListener("blur", function () {
      if (state.overlayBlinkActive) {
        state.overlayBlinkActive = false;
        renderOverlay();
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
      if (state.bottomCollapsed) return;
      e.preventDefault();
      dragging = true;
      startY = e.clientY || (e.touches && e.touches[0].clientY) || 0;
      startHeight = state.panelHeight;
      handle.classList.add("active");
      document.body.classList.add("panel-dragging");
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
      document.body.classList.remove("panel-dragging");
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    }

    document.addEventListener("mouseup", onUp);
    document.addEventListener("touchend", onUp);

    handle.addEventListener("dblclick", function (e) {
      e.preventDefault();
      toggleBottomPanel();
    });
  }

  function toggleBottomPanel() {
    var panel = qs("#bottomPanel");
    if (!panel || panel._transitioning) return;
    panel._transitioning = true;

    if (state.bottomCollapsed) {
      // --- Restore ---
      state.bottomCollapsed = false;
      var maxH = Math.round(window.innerHeight * 0.6);
      var targetH = Math.min(state.panelHeightBeforeCollapse || bottomPanelHeightFromToken(), maxH);

      document.body.classList.add("bottom-animating");
      document.body.classList.remove("bottom-collapsed");

      panel.style.height = "0px";
      panel.offsetHeight; // reflow — pin start frame
      panel.style.height = targetH + "px";

      onCollapseTransitionEnd(panel, function () {
        state.panelHeight = targetH;
        panel._transitioning = false;
        document.body.classList.remove("bottom-animating");
      });
    } else {
      // --- Collapse ---
      state.bottomCollapsed = true;
      state.panelHeightBeforeCollapse = state.panelHeight;

      var currentH = panel.offsetHeight;
      document.body.classList.add("bottom-animating");

      panel.style.height = currentH + "px";
      panel.offsetHeight; // reflow
      document.body.classList.add("bottom-collapsed");
      panel.style.height = "0px";

      onCollapseTransitionEnd(panel, function () {
        panel._transitioning = false;
        document.body.classList.remove("bottom-animating");
      });
    }
  }

  function onCollapseTransitionEnd(el, cb) {
    var fired = false;
    function done() {
      if (fired) return;
      fired = true;
      el.removeEventListener("transitionend", handler);
      cb();
    }
    function handler(e) {
      if (e.target === el && e.propertyName === "height") done();
    }
    el.addEventListener("transitionend", handler);
    setTimeout(done, 400);
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

  function initTopNavActions() {
    if (!window.ClipgenTopNav) return;
    function rebuild() {
      window.ClipgenTopNav.setQuickActions([
        window.ClipgenExportActions.exportQuickAction(),
      ]);
    }
    rebuild();
    window.ClipgenExportActions.refreshExportStatus(rebuild);
    window.ClipgenTopNav.onBeforeOpen(function () {
      window.ClipgenExportActions.refreshExportStatus(rebuild);
    });
  }

  // ---- Settings (server-side STUDIO_SETTINGS) ----
  //
  // Backed by /api/settings. We mirror the SCREENSPACE_RESTORE_MARKERS_ON_EDIT
  // flag onto state.restoreMarkersOnEdit so restoreTaskToWorkflow can read it
  // without a network call. The settings modal's onSave/onReset hooks call
  // applyScreenspaceSettingsSnapshot to keep state in sync after user edits.

  function applyScreenspaceSettingsSnapshot(applied, settings) {
    var v;
    if (applied && Object.prototype.hasOwnProperty.call(applied, "SCREENSPACE_RESTORE_MARKERS_ON_EDIT")) {
      v = applied.SCREENSPACE_RESTORE_MARKERS_ON_EDIT;
    } else if (settings) {
      for (var i = 0; i < settings.length; i++) {
        if (settings[i].name === "SCREENSPACE_RESTORE_MARKERS_ON_EDIT") {
          v = settings[i].value;
          break;
        }
      }
    }
    if (v !== undefined) state.restoreMarkersOnEdit = !!v;
  }

  function fetchScreenspaceSettings() {
    fetch("/api/settings")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (!data || !data.ok || !data.settings) return;
        applyScreenspaceSettingsSnapshot(null, data.settings);
      })
      .catch(function () { /* keep config-default state.restoreMarkersOnEdit */ });
  }

  document.addEventListener("DOMContentLoaded", function () {
    initThemeToggle(function () { refreshThemeColors(); renderTimeline(); });
    initFrameControls();
    initVideoPlayback();
    initRegionDrawing();
    initTimeline();
    initWorkflowTabs();
    initModelView();
    initParamTooltips();
    initRunButton();
    initTaskQueue();
    initRightPaneTabs();
    initPauseButton();
    initTaskFilters();
    initResultsPanel();
    initPanelDivider();
    initPreviewResize();
    initInfoNotes();
    initInfoPanelCollapse();
    initInfoSections();
    initKeyboard();
    initFrontendSwitcher();
    initTopNavActions();

    // Settings
    fetchScreenspaceSettings();
    var settingsBtn = qs("#settingsBtn");
    if (settingsBtn) {
      settingsBtn.addEventListener("click", function () {
        if (typeof window.openSettingsModal === "function") {
          window.openSettingsModal({
            initialTab: "Screenspace",
            onSave: function (applied, settings) {
              applyScreenspaceSettingsSnapshot(applied, settings);
            },
            onReset: function (scope, settings) {
              applyScreenspaceSettingsSnapshot(null, settings);
            },
          });
        }
      });
    }

    // Participant select
    qs("#participantSelect").addEventListener("change", function () {
      var pid = this.value;
      if (pid) {
        var stored = getStoredUIState("screenspace");
        var ts;
        if (stored.videoTimeByParticipant && typeof stored.videoTimeByParticipant[pid] === "number") {
          ts = stored.videoTimeByParticipant[pid];
        }
        selectParticipant(pid, ts);
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
        // Seed _videoVersions before any frameUrl/videoStreamUrl call so the
        // preload loop below already includes the ?v= cache-bust suffix.
        state.participants.forEach(function (p) {
          if (p.version != null) _videoVersions[p.id] = String(p.version);
        });
        renderParticipantSelect();
        // Preload frame 0 for all participants (instant first-frame display)
        state.participants.forEach(function (p) {
          var url = frameUrl(p.id, 0);
          // TODO: fire-and-forget blob preload; needs apiGetBlob helper to migrate.
          fetch(url)
            .then(function (r) { return r.blob(); })
            .then(function (blob) {
              if (_preloadedFrames[p.id]) {
                try { URL.revokeObjectURL(_preloadedFrames[p.id]); } catch (_) {}
              }
              _preloadedFrames[p.id] = URL.createObjectURL(blob);
            })
            .catch(function () {});
        });
        if (state.participants.length > 0) {
          var stored = getStoredUIState("screenspace");
          var pickId = state.participants[0].id;
          if (stored.selectedParticipant) {
            for (var spi = 0; spi < state.participants.length; spi++) {
              if (state.participants[spi].id === stored.selectedParticipant) {
                pickId = stored.selectedParticipant;
                break;
              }
            }
          }
          var initialTs;
          if (stored.videoTimeByParticipant && typeof stored.videoTimeByParticipant[pickId] === "number") {
            initialTs = stored.videoTimeByParticipant[pickId];
          }
          selectParticipant(pickId, initialTs);
          state.runParticipants = [pickId];
          if (stored.rightPaneTab === "queue" || stored.rightPaneTab === "results") {
            setRightPaneTab(stored.rightPaneTab);
          }
          if (stored.activeWorkflow) {
            var wfTab = qs('.wf-tab[data-type="' + CSS.escape(stored.activeWorkflow) + '"]');
            if (wfTab) wfTab.click();
          }
        }
        renderRunParticipantPicker();
        renderScanModePicker();
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
            startSSE();
          }
        }
      })
      .catch(function () {});
  });

})();
