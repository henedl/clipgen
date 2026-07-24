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

  var FRAME_STEP = 1.0; // fine step (,/. and Shift+arrow)
  var SEEK_STEP = 5.0; // coarse step (arrows + the timeline step buttons)
  var VIDEO_SPEEDS = [0.5, 1, 2, 3, 5];

  var TASK_COLORS = DETECTOR_COLORS;

  var SS_TASK_ICON_TYPES = {
    multitool: 1, color: 1, change: 1, similarity: 1, text: 1,
    numbers: 1, template: 1, flow: 1, scene: 1, inactivity: 1,
    boundary: 1, attention: 1, timelapse: 1,
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
    return iconMaskSpan(name, {
      className: "ss-icon" + (sizeClass ? " " + sizeClass : ""),
      basePath: "/screenspace/icons/",
    });
  }

  // Three-state OCR-normalize direction control: a segmented icon button-set
  // backed by a hidden input holding the mode string ("letters" | "off" |
  // "digits"). It folds easily-confused glyphs toward whichever canonical form
  // you pick before the fuzzy compare — see _normalize_ocr_text in
  // screenspace_ocr.py. Off sits in the middle: digit→letter | off | letter→digit.
  var NORMALIZE_MODES = [
    { mode: "letters", icon: "language", desc: "Fold digits to letters before matching (0→o, 1→l, 5→s). For word targets that OCR may read as digits" },
    { mode: "off", icon: "no-symbol", desc: "No character folding" },
    { mode: "digits", icon: "hashtag", desc: "Fold letters to digits before matching (O→0, l→1, S→5). For number targets that OCR may read as letters" },
  ];

  function _normalizeMode(mode) {
    return mode === "letters" || mode === "digits" ? mode : "off";
  }

  function buildNormalizeControl(id, mode, small) {
    var wrap = el("div", "ss-segctl" + (small ? " ss-segctl--sm" : ""));
    var hidden = document.createElement("input");
    hidden.type = "hidden";
    hidden.id = id;
    hidden.value = _normalizeMode(mode);
    wrap.appendChild(hidden);
    NORMALIZE_MODES.forEach(function (spec) {
      var btn = el("button", "ss-segctl-btn");
      btn.type = "button";
      // Description shown via the custom param tooltip (see initParamTooltips),
      // matching how param labels surface their help text.
      btn.setAttribute("data-desc", spec.desc);
      btn.setAttribute("data-mode", spec.mode);
      btn.appendChild(iconSpan(spec.icon, "ss-icon--xs"));
      if (spec.mode === hidden.value) btn.classList.add("active");
      btn.addEventListener("click", function (e) {
        e.preventDefault();
        hidden.value = spec.mode;
        var sibs = wrap.querySelectorAll(".ss-segctl-btn");
        for (var i = 0; i < sibs.length; i++) {
          sibs[i].classList.toggle("active", sibs[i] === btn);
        }
        // Bubbles to the .param-control wrapper so addParamRow's input handler
        // (live model preview) fires, mirroring a checkbox change.
        hidden.dispatchEvent(new Event("input", { bubbles: true }));
      });
      wrap.appendChild(btn);
    });
    return wrap;
  }

  // Reflect a mode string back onto an existing segmented control (used when
  // rehydrating a saved task into the editor).
  function applyNormalizeMode(id, mode) {
    var hidden = qs("#" + id);
    if (!hidden) return;
    var m = _normalizeMode(mode);
    hidden.value = m;
    var wrap = hidden.parentNode;
    if (!wrap) return;
    var btns = wrap.querySelectorAll(".ss-segctl-btn");
    for (var i = 0; i < btns.length; i++) {
      btns[i].classList.toggle("active", btns[i].getAttribute("data-mode") === m);
    }
  }

  // Two-state Color match-mode control: "average" (region's mean color) vs
  // "presence" (target color appears anywhere in the region, per-pixel). Backed
  // by a hidden input holding the mode string. See ColorTool in screenspace_tools.py.
  var COLOR_MODES = [
    { mode: "average", icon: "swatch", desc: "Match the region's average colour" },
    { mode: "presence", icon: "magnifying-glass-circle", desc: "Match when the target colour appears anywhere in the region (per-pixel)" },
  ];

  function _colorMode(mode) {
    return mode === "presence" ? "presence" : "average";
  }

  // Reflect a color mode back onto an existing segmented control + presence-only
  // min-area row (used when rehydrating a saved single-tool color task).
  function applyColorMode(id, mode) {
    var hidden = qs("#" + id);
    if (!hidden) return;
    var m = _colorMode(mode);
    hidden.value = m;
    var wrap = hidden.parentNode;
    if (wrap) {
      var btns = wrap.querySelectorAll(".ss-segctl-btn");
      for (var i = 0; i < btns.length; i++) {
        btns[i].classList.toggle("active", btns[i].getAttribute("data-mode") === m);
      }
    }
    var row = qs("#paramColorMinAreaRow");
    if (row) row.classList.toggle("hidden", m !== "presence");
  }

  // `onChange(mode)` fires after the active button flips, before the bubbling
  // input event — callers use it to show/hide the presence-only min-area row.
  function buildColorModeControl(id, mode, small, onChange) {
    var wrap = el("div", "ss-segctl" + (small ? " ss-segctl--sm" : ""));
    var hidden = document.createElement("input");
    hidden.type = "hidden";
    hidden.id = id;
    hidden.value = _colorMode(mode);
    wrap.appendChild(hidden);
    COLOR_MODES.forEach(function (spec) {
      var btn = el("button", "ss-segctl-btn");
      btn.type = "button";
      btn.setAttribute("data-desc", spec.desc);
      btn.setAttribute("data-mode", spec.mode);
      btn.appendChild(iconSpan(spec.icon, "ss-icon--xs"));
      if (spec.mode === hidden.value) btn.classList.add("active");
      btn.addEventListener("click", function (e) {
        e.preventDefault();
        hidden.value = spec.mode;
        var sibs = wrap.querySelectorAll(".ss-segctl-btn");
        for (var i = 0; i < sibs.length; i++) {
          sibs[i].classList.toggle("active", sibs[i] === btn);
        }
        if (onChange) onChange(spec.mode);
        // Bubbles to the .param-control wrapper so addParamRow's input handler
        // (live model preview) fires, mirroring a checkbox change.
        hidden.dispatchEvent(new Event("input", { bubbles: true }));
      });
      wrap.appendChild(btn);
    });
    return wrap;
  }

  var _paletteDocListeners = null;
  // Cached HSV hidden inputs for the single-tool color picker. Populated by
  // renderColorParams() when the panel is built; reused by setTargetColor,
  // updateColorPreview, _collectPreviewParams, etc. so they don't re-query
  // the DOM on every drag tick. Null when no color tool panel is active.

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
    // Shaped-region drawing: the active selector ("rect" | "lasso" | "wand"),
    // the in-progress freehand point trail, the wand's flood-fill RGB
    // tolerance, and the wand's press-drag-release scrub state (seed + live
    // preview contour). On state (not satellite vars) — the hub Escape handler
    // and the overlay painter both read them across file boundaries.
    regionTool: "rect",
    drawingLasso: null,
    wandTolerance: 32,
    wandDragging: null,
    pendingRegion: null,
    draggingRegion: null,
    resizingRegion: null,
    hoveredRegion: null,
    // Set by the hub's chip drag-reorder (initRegionDrag) to swallow the click
    // that fires right after a drop; read/cleared by renderRegionChips' chip
    // click handler in screenspace-overlay-interaction.js (cross-file — must
    // live on state, not a hub-local var).
    regionSuppressNextClick: false,
    timelineZoom: 1,
    timelineOffset: 0,
    inMarker: null,
    outMarker: null,
    restoreMarkersOnEdit: true,
    showConfidenceHistogram: false,
    // Grouped category tool nav (SCREENSPACE_GROUPED_TOOL_NAV). Init true to
    // match the Python default before /api/settings resolves.
    groupedToolNav: true,
    activeWorkflow: "color",
    // Panel-focus keyboard navigation (Shift+1..4 + arrows). focusRegion is the
    // surface the arrows drive: "video" (the default) means transport seek and
    // is what Escape returns to; "sidebar"/"tool"/"task"/"results" rove within a
    // panel. focusCursor indexes ssNavItems(focusRegion). navEditing is true
    // while a text control (notes / a param input) holds real focus for typing.
    // pickerCursor (>= 0) indexes into an open run-picker dropdown's options.
    focusRegion: "video",
    focusCursor: 0,
    navEditing: false,
    pickerCursor: -1,
    referenceTimestamp: null,
    sceneReferences: [],
    tasks: [],
    selectedTaskId: null,
    hoveredTaskId: null,
    selectedTaskResults: null,
    // Per-task result cache (taskId -> results array). Status ticks no longer
    // carry result lists; the timeline and results panel read from here, kept
    // current by appending result tails (see _syncTaskResults in tasks).
    taskResults: {},
    resultsLoading: false,
    resultsLazyObserver: null,
    // Coordination flags shared between the hub and screenspace-tasks.js:
    // resultsRequestVersion gates in-flight results fetches; suppressCalibration-
    // Refresh is set while restoreTaskToWorkflow rebuilds the param panel.
    resultsRequestVersion: 0,
    heatmapOverlayRequestVersion: 0,
    suppressCalibrationRefresh: false,
    poller: null,
    eventSource: null,
    sseFellBack: false,
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
    // True while runRegions holds only the implicit active-chip seed (no
    // explicit run-picker choice yet) — see renderRunRegionPicker.
    runRegionsSeeded: false,
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
    hoveredBoundaryTs: null,
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
    pins: [],
    maxPins: null,
    hoveredPinId: null,
    pinTrayHidden: false,
    calibrationOpen: false,
    calibrationResult: null,
    calibrationOcrWarmed: false,
    calibrationGreen: false,
  };

  var _playheadRaf = 0;
  // _lastPollFingerprint / _etaTrackers / _etaTicker moved into
  // screenspace-tasks.js along with the task-queue surface that owns them.
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
      positive: base.positive,
      fontMono: base.fontMono,
      regionPalette: _cachedRegionPalette,
    };
  }

  // ---- Helpers ----

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

  // ---- Multi-video timeline helpers ----
  // For a participant whose recording spans several files, api/video/info
  // returns ``parts`` ([{filename, duration, cumulativeStart}]) and a total
  // ``duration``. Frame display already works at global time (the backend maps
  // it); only the <video> play element switches source per part below.
  function _ssParts() {
    var info = state.videoInfo;
    return info && info.parts && info.parts.length > 1 ? info.parts : null;
  }
  function _ssPartForGlobal(parts, g) {
    for (var i = 0; i < parts.length; i++) {
      if (g >= parts[i].cumulativeStart && g < parts[i].cumulativeStart + parts[i].duration) {
        return i;
      }
    }
    return parts.length - 1;
  }
  function _ssStreamUrlForPart(pid, i) {
    var url = videoStreamUrl(pid);
    return url + (url.indexOf("?") >= 0 ? "&" : "?") + "part=" + i;
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
    // An explicit picker choice pins the selection — stop following the chip.
    state.runRegionsSeeded = false;
  }

  function removeRunRegion(ref) {
    var key = regionRefKey(ref);
    state.runRegions = state.runRegions.filter(function (r) { return regionRefKey(r) !== key; });
    state.runRegionsSeeded = false;
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
    applyIconMask(icon, "arrows-pointing-out", "/screenspace/icons/");
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
    // Auto-select the active region when no explicit selection has been made.
    // The seed is implicit (runRegionsSeeded): it keeps FOLLOWING the
    // highlighted chip on later renders until the user touches the picker
    // (addRunRegion/removeRunRegion clear the flag). Without the follow-up
    // re-seed, the first created region stayed pinned and the model preview
    // ignored chip selection.
    var seedRef = state.activeRegion && names.indexOf(state.activeRegion) >= 0
      ? activeRegionRef(state.activeRegion)
      : null;
    if (state.runRegions.length === 0) {
      if (seedRef) {
        state.runRegions = [seedRef];
        state.runRegionsSeeded = true;
      }
    } else if (
      state.runRegionsSeeded &&
      seedRef &&
      state.runRegions.length === 1 &&
      state.runRegions[0].source === "active" &&
      regionRefKey(state.runRegions[0]) !== regionRefKey(seedRef)
    ) {
      state.runRegions = [seedRef];
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
      refreshModelView({ debounce: true });
      refreshCalibration({ debounce: true });
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
        refreshModelView({ debounce: true });
        refreshCalibration({ debounce: true });
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
          refreshModelView({ debounce: true });
          refreshCalibration({ debounce: true });
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
          refreshModelView({ debounce: true });
          refreshCalibration({ debounce: true });
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

  // A tool supports fast scan iff it has a fast-scan description. This is the
  // single source of truth for the toggle's visibility and whether a submitted
  // task carries scan_mode:"fast" — timelapse (media output) and boundary (runs
  // its own coarse phash pass) are both absent above and so opt out.
  function toolSupportsFastScan(type) { return !!FAST_SCAN_DESCRIPTIONS[type]; }

  var PARAM_DESCRIPTIONS = {
    _shared: {
      "Event label":      "Tag added to each detected event for filtering",
      "Detect first":     "Stop after the first match is found",
      "Region":           "Which screen region this step analyzes",
    },
    color: {
      "Tolerance":        "How far from the target color still counts. Widen to catch more shades, tighten to be stricter",
      "Hex color":        "Target color in hex notation",
      "Mode":             "Average matches the region's mean colour; Presence fires when the target colour appears anywhere in the region (per-pixel)",
      "Min area %":       "Presence mode only: the minimum share of region pixels that must match before an event fires. 0% = any presence detected (no minimum size); raise it to ignore stray noise pixels. The readout shows the approximate pixel count for the current region.",
    },
    change: {
      "Threshold":        "How much of the region must change to trigger. Raise it to ignore minor flicker",
      "Noise Thr.":       "Ignore changes below this pixel intensity",
      "Noise":            "Ignore changes below this pixel intensity",
      "Consecutive":      "Require this many consecutive sampled frames to match before an event fires (suppresses single-frame flicker; reports the run's median time)",
    },
    similarity: {
      "Reference":        "Capture the frame you want later frames to match against",
      "Threshold":        "How close a frame must look to the reference. Lower it to allow looser matches",
    },
    text: {
      "Search text":      "Exact or partial text to find on screen",
      "Search":           "Exact or partial text to find on screen",
      "Fuzzy Thr.":       "How closely the text must match (1.0 = exact). Lower it to tolerate more misreads",
      "Fuzzy":            "How closely the text must match (1.0 = exact). Lower it to tolerate more misreads",
      "Min OCR conf.":    "Drop OCR readings below this confidence before fuzzy matching (raise to suppress noisy misreads)",
      "Min OCR":          "Drop OCR readings below this confidence before fuzzy matching (raise to suppress noisy misreads)",
      "Enhance ROI":      "Upscale small/low-contrast crops and apply CLAHE before OCR (slower; helps tiny HUD text)",
      "Normalize":        "Fold easily-confused glyphs before matching: digits to letters, off, or letters to digits. Pick the side that matches your search target (letters vs digits).",
      "Language":         "OCR language for text recognition",
      "Consecutive":      "Require this many consecutive sampled frames to match before an event fires (suppresses single-frame flicker; reports the run's median time)",
    },
    numbers: {
      "Operator":         "Comparison operator for the detected number",
      "Target value":     "Number to compare the detected value against",
      "Target":           "Number to compare the detected value against",
      "Range":            "Min and max bounds for the in-range check",
      "Min OCR conf.":    "Drop OCR readings below this confidence before parsing numbers (raise to suppress noisy misreads)",
      "Min OCR":          "Drop OCR readings below this confidence before parsing numbers (raise to suppress noisy misreads)",
      "Enhance ROI":      "Upscale small/low-contrast crops and apply CLAHE before OCR (slower; helps tiny HUD numbers)",
      "Integers only":    "Restrict OCR to digits only (drop . , -) so a separator glyph can't survive as a digit and inflate the value. For whole-number HUD targets. English only.",
      "Integers":         "Restrict OCR to digits only (drop . , -) so a separator glyph can't survive as a digit and inflate the value. For whole-number HUD targets. English only.",
      "Consecutive":      "Require this many consecutive sampled frames to match before an event fires (suppresses single-frame flicker; reports the run's median time)",
    },
    timelapse: {
      "Speed":            "Playback speed multiplier for the output",
      "Sample every":     "Seconds between captured frames (0 = every frame)",
      "Format":           "Output file format: video or animated GIF",
    },
    template: {
      "Template":         "Capture or upload the picture to search for anywhere on screen",
      "Threshold":        "How closely the picture must match. Lower it to allow looser matches",
    },
    flow: {
      "Magnitude":        "Minimum movement strength to count. Raise it to ignore small or slow motion",
      "Consecutive":      "Require this many consecutive sampled frames to match before an event fires (suppresses single-frame flicker; reports the run's median time)",
    },
    scene: {
      "Add Scene":        "Capture and name each screen you want to recognize",
    },
    inactivity: {
      "Sensitivity":      "How little movement still counts as idle. Raise it to treat more frames as still",
      "Min duration (s)": "Seconds of stillness required before a stall is reported",
    },
    boundary: {
      "Sensitivity":      "Minimum frame-to-frame change to call a scene boundary (higher = only the biggest jumps)",
      "Min gap (s)":      "Suppress further boundaries for this long after one fires (avoids storms during fast action)",
    },
    attention: {
      "Sensitivity":      "How far the predicted focus must jump (as a fraction of the screen) to mark an attention shift. Raise it to mark only big jumps",
      "Smoothing":        "How quickly the attention map follows each new frame (1.0 = instant, lower = steadier but slower to react)",
      "Spectral wt.":     "Weight of the 'unexpected detail' channel (spectral residual): odd shapes and busy areas that stand out from the scene",
      "Contrast wt.":     "Weight of the color/brightness contrast channel: elements that differ strongly from their surroundings",
      "Motion wt.":       "Weight of the motion channel: areas that changed since the previous sample. Usually the strongest pull on screens",
      "Faces wt.":        "Weight of the face channel (webcam picture-in-picture). 0 turns face detection off; UI avatars can trigger false hits",
      "Center bias":      "How much the map favors the screen center. 0 = no preference; photos-style footage tolerates more than UI recordings",
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
    applyIconMask(icon, "chevron-double-right", "/screenspace/icons/");
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
      // Segmented-control buttons (e.g. the Normalize direction set) carry their
      // own description on data-desc; reuse the same dark-pill tooltip.
      var seg = e.target.closest && e.target.closest(".ss-segctl-btn");
      if (seg) {
        var segDesc = seg.getAttribute("data-desc");
        if (segDesc) tooltip.show(seg, segDesc);
        return;
      }
      var label = e.target.closest(".param-label");
      if (!label) return;
      var text = label.textContent.trim();
      var desc = getDescription(text, getToolType(label));
      if (!desc) return;
      tooltip.show(label, desc);
    }, true);

    container.addEventListener("mouseleave", function (e) {
      if (
        e.target.closest &&
        (e.target.closest(".param-label") || e.target.closest(".ss-segctl-btn"))
      ) {
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
    state.heatmapOverlayRequestVersion += 1;
    state.selectedParticipant = pid;
    setStoredUIStateField("screenspace", "selectedParticipant", pid);
    // Fill the result cache for the newly selected participant's tasks so the
    // timeline draws their markers (the sync is participant-scoped, and no SSE
    // tick may follow when their tasks are already completed).
    if (SS.syncTaskResults) SS.syncTaskResults();
    state.currentTimestamp = 0;
    state.videoInfo = null;
    state.videoActivePart = 0;
    state.videoOffset = 0;
    state.frameImage = null;
    state.frameLoading = false;
    state.referenceTimestamp = null;
    state.sceneReferences = [];
    state.resultOverlay = null;
    state.heatmapOverlay = null;
    state.pins = [];
    state.hoveredPinId = null;
    // Tray visibility is per-participant: don't carry one participant's hidden
    // tray over to another that has pins (mirrors the reset when pinning).
    state.pinTrayHidden = false;
    // Drop the prior participant's calibration scores and invalidate any
    // in-flight /api/calibrate response so it can't repaint the strip; the pin
    // load below re-evaluates once the new participant's pins arrive.
    state.calibrationResult = null;
    if (SS.calBumpGen) SS.calBumpGen();
    updateCalibrationVisibility();
    renderCalibration();
    renderPinTray();
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
    // Skeleton frame + a textual "Loading…" in the subheader make the switch
    // legibly in-progress; both are replaced once video info resolves (or on the
    // error path below).
    qs("#videoInfo").textContent = "Loading…";
    qs("#frameEmpty").classList.remove("hidden");
    setInfoParticipant(pid);

    apiGet("api/video/info/" + encodeURIComponent(pid))
      .then(function (data) {
        if (participantRequestVersion !== _participantRequestVersion || pid !== state.selectedParticipant) return;
        if (!data.ok) { qs("#videoInfo").textContent = ""; return; }
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
        updatePinButtons();
        // Preload video source for instant playback
        qs("#videoPlayer").src = videoStreamUrl(pid);
        loadFrame(initialTimestamp !== undefined ? initialTimestamp : 0);
      })
      .catch(function () {
        // Clear the "Loading…" placeholder for the still-current participant so
        // it doesn't hang after a failed fetch.
        if (participantRequestVersion === _participantRequestVersion && pid === state.selectedParticipant) {
          qs("#videoInfo").textContent = "";
        }
        showToast("Failed to load video info");
      });

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

    apiGet("api/participants/" + encodeURIComponent(pid) + "/marks")
      .then(function (data) {
        if (participantRequestVersion !== _participantRequestVersion || pid !== state.selectedParticipant) return;
        if (!data.ok) return;
        if (data.categories) setMarkCategories(data.categories);
        renderInfoMarks(data.marks || []);
      })
      .catch(toastError("Failed to load marks"));

    apiGet("api/pins/" + encodeURIComponent(pid))
      .then(function (data) {
        if (participantRequestVersion !== _participantRequestVersion || pid !== state.selectedParticipant) return;
        if (!data.ok) return;
        state.pins = data.pins || [];
        state.maxPins = data.max_pins != null ? data.max_pins : null;
        renderPinTray();
        updatePinButtons();
        renderTimeline();
        updateCalibrationVisibility();
        refreshCalibration();
      })
      .catch(toastError("Failed to load pins"));
  }

  // ---- Calibration pins ----
  //
  // A pin marks "this frame matters" with a polarity (positive = the
  // condition is true here; negative = it must not fire here). Pins are
  // tool-agnostic and drive detector calibration (later phase). The tray below
  // the viewer shows on-demand thumbnails; the timeline shows polarity ticks.

  var PIN_THUMB_WIDTH = 140;

  function pinThumbUrl(pid, ts) {
    var u = frameUrl(pid, ts);
    return u + (u.indexOf("?") === -1 ? "?" : "&") + "w=" + PIN_THUMB_WIDTH;
  }

  function updatePinButtons() {
    var posBtn = qs("#pinPositiveBtn");
    var negBtn = qs("#pinNegativeBtn");
    if (!posBtn || !negBtn) return;
    var hasVideo = !!(state.selectedParticipant && state.videoInfo);
    var atCap = state.maxPins != null && state.pins.length >= state.maxPins;
    var disabled = !hasVideo || atCap;
    posBtn.disabled = disabled;
    negBtn.disabled = disabled;
    var capNote = atCap ? ". Limit reached (" + state.maxPins + ")" : "";
    posBtn.title = "Pin this frame as a positive (condition is true here)" + capNote;
    negBtn.title = "Pin this frame as a negative (condition must not fire here)" + capNote;
  }

  function pinCurrentFrame(polarity) {
    var pid = state.selectedParticipant;
    if (!pid || !state.videoInfo) return;
    if (state.maxPins != null && state.pins.length >= state.maxPins) {
      showToast("Pin limit reached (" + state.maxPins + ")");
      return;
    }
    var ts = state.currentTimestamp;
    apiPost("api/pins/" + encodeURIComponent(pid), { timestamp: ts, polarity: polarity })
      .then(function (data) {
        if (!data.ok) {
          showToast(data.error || "Failed to pin frame");
          return;
        }
        if (pid !== state.selectedParticipant) return;
        state.pins.push(data.pin);
        // Re-reveal the tray so a new pin is always visible, even if hidden.
        state.pinTrayHidden = false;
        renderPinTray();
        updatePinButtons();
        renderTimeline();
        updateCalibrationVisibility();
        refreshCalibration();
      })
      .catch(function () { showToast("Failed to pin frame"); });
  }

  function removePin(pinId) {
    apiDelete("api/pins/" + encodeURIComponent(pinId))
      .then(function (data) {
        if (!data.ok) {
          showToast(data.error || "Failed to remove pin");
          return;
        }
        state.pins = state.pins.filter(function (p) { return p.id !== pinId; });
        renderPinTray();
        updatePinButtons();
        renderTimeline();
        updateCalibrationVisibility();
        refreshCalibration();
      })
      .catch(function () { showToast("Failed to remove pin"); });
  }

  function togglePinTrayVisibility() {
    state.pinTrayHidden = !state.pinTrayHidden;
    renderPinTray();
  }

  function clearAllPins() {
    var pid = state.selectedParticipant;
    if (!pid || !state.pins.length) return;
    if (!window.confirm("Clear all " + state.pins.length + " pinned frame(s)? This cannot be undone.")) return;
    apiDelete("api/pins/" + encodeURIComponent(pid) + "/all")
      .then(function (data) {
        if (!data.ok) {
          showToast(data.error || "Failed to clear pins");
          return;
        }
        if (pid !== state.selectedParticipant) return;
        state.pins = [];
        state.hoveredPinId = null;
        renderPinTray();
        updatePinButtons();
        renderTimeline();
        updateCalibrationVisibility();
        refreshCalibration();
        showToast("All pins cleared");
      })
      .catch(function () { showToast("Failed to clear pins"); });
  }

  function togglePinPolarity(pinId) {
    var pin = state.pins.filter(function (p) { return p.id === pinId; })[0];
    if (!pin) return;
    var next = pin.polarity === "positive" ? "negative" : "positive";
    apiPut("api/pins/" + encodeURIComponent(pinId), { polarity: next })
      .then(function (data) {
        if (!data.ok || !data.pin) {
          showToast(data.error || "Failed to update pin");
          return;
        }
        pin.polarity = data.pin.polarity;
        renderPinTray();
        renderTimeline();
        refreshCalibration();
      })
      .catch(function () { showToast("Failed to update pin"); });
  }

  function renderPinTray() {
    var tray = qs("#pinTray");
    var list = qs("#pinTrayItems");
    if (!tray || !list) return;
    var pins = state.pins || [];
    var hasPins = pins.length > 0;
    // The toggle + clear controls only matter once there are pins.
    var tBtn = qs("#togglePinTrayBtn");
    if (tBtn) {
      tBtn.classList.toggle("hidden", !hasPins);
      tBtn.innerHTML = "";
      tBtn.appendChild(iconSpan(state.pinTrayHidden ? "eye-slash" : "eye"));
      tBtn.title = state.pinTrayHidden ? "Show pinned frames" : "Hide pinned frames";
      tBtn.classList.toggle("active", state.pinTrayHidden);
    }
    var clearBtn = qs("#clearPinsBtn");
    if (clearBtn) clearBtn.classList.toggle("hidden", !hasPins);
    tray.classList.toggle("hidden", !hasPins || state.pinTrayHidden);
    list.innerHTML = "";
    if (!hasPins) return;
    var pid = state.selectedParticipant;
    var sorted = pins.slice().sort(function (a, b) { return a.timestamp - b.timestamp; });
    var frag = document.createDocumentFragment();
    sorted.forEach(function (pin) {
      var item = el("div", "pin-tray-item pin-tray-item--" + pin.polarity);
      if (pin.stale) item.classList.add("pin-tray-item--stale");
      item.setAttribute("data-pin-id", pin.id);

      var img = document.createElement("img");
      img.decoding = "async";
      img.className = "pin-tray-thumb";
      img.alt = "";
      img.loading = "lazy";
      if (pid) img.src = pinThumbUrl(pid, pin.timestamp);
      item.appendChild(img);

      var meta = el("div", "pin-tray-meta");
      var dot = el("span", "pin-tray-polarity");
      dot.title = "Toggle polarity (" + pin.polarity + ")";
      dot.addEventListener("click", function (e) {
        e.stopPropagation();
        togglePinPolarity(pin.id);
      });
      meta.appendChild(dot);
      var time = el("span", "");
      time.textContent = formatTime(pin.timestamp, { decimals: 1 });
      meta.appendChild(time);
      if (pin.stale) {
        var staleTag = el("span", "pin-tray-stale-tag");
        staleTag.textContent = "stale";
        staleTag.title = "Timestamp is beyond the current video duration";
        meta.appendChild(staleTag);
      }
      item.appendChild(meta);

      var remove = el("button", "pin-tray-remove");
      remove.type = "button";
      remove.title = "Remove pin";
      remove.appendChild(iconSpan("x-mark", "ss-icon--xs"));
      remove.addEventListener("click", function (e) {
        e.stopPropagation();
        removePin(pin.id);
      });
      item.appendChild(remove);

      item.addEventListener("click", function () { loadFrame(pin.timestamp); });
      item.addEventListener("mouseenter", function () {
        state.hoveredPinId = pin.id;
        renderTimeline();
      });
      item.addEventListener("mouseleave", function () {
        if (state.hoveredPinId === pin.id) {
          state.hoveredPinId = null;
          renderTimeline();
        }
      });
      frag.appendChild(item);
    });
    list.appendChild(frag);
  }

  // ---- Info panel ----

  function setInfoParticipant(pid) {
    var el = qs("#ssInfoParticipant");
    if (el) el.textContent = pid || "\u2014";
    qs("#ssInfoNotes").value = "";
    qs("#ssInfoIssuesBlock").classList.add("hidden");
    qs("#ssInfoIssues").innerHTML = "";
    qs("#ssInfoMarksBlock").classList.add("hidden");
    qs("#ssInfoMarks").innerHTML = "";
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

  function renderInfoMarks(marks) {
    var block = qs("#ssInfoMarksBlock");
    var list = qs("#ssInfoMarks");
    if (!block || !list) return;
    list.innerHTML = "";
    if (!marks || !marks.length) {
      block.classList.add("hidden");
      return;
    }
    block.classList.remove("hidden");
    var frag = document.createDocumentFragment();
    marks.forEach(function (mark) {
      var li = document.createElement("li");
      li.className = "ss-info-issue";
      var cat = MARK_CATEGORIES[mark.category] || MARK_CATEGORIES.bookmark;
      var dot = document.createElement("span");
      dot.className = "ss-info-issue-dot";
      if (cat) dot.style.backgroundColor = cat.color;
      var text = document.createElement("span");
      text.className = "ss-info-issue-text";
      var label = (mark.label && mark.label.trim()) || mark.text || "(mark)";
      if (label.length > 120) label = label.slice(0, 117) + "…";
      text.textContent = label;
      li.appendChild(dot);
      li.appendChild(text);
      if (mark.start != null) {
        var ts = document.createElement("span");
        ts.className = "ss-info-issue-ts";
        ts.textContent = formatTime(mark.start);
        li.appendChild(ts);
        li.classList.add("ss-info-issue--clickable");
        li.addEventListener("click", (function (t) {
          return function () { loadFrame(t); };
        })(mark.start));
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
      var ts = clamp(state.currentTimestamp - SEEK_STEP, 0, Math.max(0, state.videoInfo.duration - 0.001));
      loadFrame(ts);
    });

    qs("#frameNext").addEventListener("click", function () {
      if (!state.videoInfo) return;
      var ts = clamp(state.currentTimestamp + SEEK_STEP, 0, Math.max(0, state.videoInfo.duration - 0.001));
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
      state.videoPlaybackRate = window.ClipgenVideoControls.nextSpeed(VIDEO_SPEEDS, state.videoPlaybackRate);
      applyPlaybackRate();
      updateVideoButtons();
    });

    updateVideoButtons();

    video.addEventListener("ended", function () {
      pauseVideo();
    });

    video.addEventListener("timeupdate", function () {
      if (!state.videoPlaying) return;
      var parts = _ssParts();
      var t;
      if (parts) {
        var i = state.videoActivePart || 0;
        // Hand off to the next part near the boundary for continuous playback.
        if (i < parts.length - 1 && video.currentTime >= parts[i].duration - 0.05) {
          state.videoActivePart = i + 1;
          state.videoOffset = parts[i + 1].cumulativeStart;
          video.src = _ssStreamUrlForPart(state.selectedParticipant, i + 1);
          var onMeta = function () {
            video.removeEventListener("loadedmetadata", onMeta);
            video.currentTime = 0.001;
            video.play();
          };
          video.addEventListener("loadedmetadata", onMeta);
          return;
        }
        t = video.currentTime + (state.videoOffset || 0);
      } else {
        t = video.currentTime;
      }
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

    var parts = _ssParts();
    if (parts) {
      // Multi-video: play the part that owns the global playhead, seeking local.
      var i = _ssPartForGlobal(parts, state.currentTimestamp);
      state.videoActivePart = i;
      state.videoOffset = parts[i].cumulativeStart;
      var wantSrc = _ssStreamUrlForPart(state.selectedParticipant, i);
      if (!video.src || video.src.indexOf("part=" + i) === -1) {
        video.src = wantSrc;
      }
      video.currentTime = state.currentTimestamp - state.videoOffset;
    } else {
      state.videoActivePart = 0;
      state.videoOffset = 0;
      var expectedSrc = videoStreamUrl(state.selectedParticipant);
      if (!video.src || video.src.indexOf(expectedSrc) === -1) {
        video.src = expectedSrc;
      }
      video.currentTime = state.currentTimestamp;
    }
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

    var ts = _ssParts()
      ? video.currentTime + (state.videoOffset || 0)
      : video.currentTime || state.currentTimestamp;
    state.currentTimestamp = ts;

    video.classList.remove("active");
    qs("#frameCanvas").classList.remove("video-active");

    loadFrame(ts);
    updateVideoButtons();
  }

  function applyPlaybackRate() {
    window.ClipgenVideoControls.applyPlaybackRate(qs("#videoPlayer"), state.videoPlaybackRate);
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

  // ---- Region drawing + overlay interaction (impl in screenspace-overlay-interaction.js) ----
  // The region draw/drag/resize state machine, region chips, the region-name
  // modal, and the overlay-rect cache live in the satellite. These thin
  // delegators forward the hub's own call sites (region-editor init, the Escape
  // handler, and the stashing / chip-reorder code) to it.
  function initRegionDrawing() { return SS.initRegionDrawing && SS.initRegionDrawing.apply(null, arguments); }
  function renderRegionChips() { return SS.renderRegionChips && SS.renderRegionChips.apply(null, arguments); }
  function updateRegionButtons() { return SS.updateRegionButtons && SS.updateRegionButtons.apply(null, arguments); }
  function hideRegionNameModal() { return SS.hideRegionNameModal && SS.hideRegionNameModal.apply(null, arguments); }
  function invalidateOverlayRect() { return SS.invalidateOverlayRect && SS.invalidateOverlayRect.apply(null, arguments); }

  // ---- Region stashing ----

  // Id of the stash just created, so renderStashCards() can play the landing
  // animation on exactly the new card (consumed + cleared on first render).
  var _justStashedStashId = null;

  function stashRegions() {
    var chips = qsa("#regionChips .region-chip");
    apiPost("api/stashes", {}).then(function (data) {
      if (!data.ok) return;
      var commit = function () {
        state.stashes.push(data.stash);
        _justStashedStashId = data.stash.id;
        state.regions = {};
        state.activeRegion = null;
        state.pendingRegion = null;
        renderRegionChips();
        renderOverlay();
        updateRegionButtons();
        updateRunButton();
        renderStashCards();
        showToast("Regions stashed");
      };
      // Pills stash out (jump + wiggle + dissolve), then the stash card lands.
      if (chips.length && window.ClipgenMotion) ClipgenMotion.animateOutAll(chips, "stash").then(commit);
      else commit();
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

  function copyRegionToStash(name, stashId) {
    if (!(name in state.regions)) return;
    apiPost("api/stashes/" + stashId + "/regions", { name: name })
      .then(function (data) {
        if (!data.ok) return;
        for (var i = 0; i < state.stashes.length; i++) {
          if (state.stashes[i].id === stashId) {
            state.stashes[i] = data.stash;
            break;
          }
        }
        renderStashCards(); // updated count + dots
        renderRunRegionPicker(); // stash folder now lists the new region
        showToast("Added “" + name + "” to " + data.stash.name);
      })
      .catch(function () {
        showToast("Failed to add region to stash");
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
      card.dataset.stashId = stash.id;
      if (stash.id === _justStashedStashId && window.ClipgenMotion) {
        ClipgenMotion.animateIn(card, "stashLand");
        _justStashedStashId = null;
      }
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
      bindStashDrop(area);
      var viewerSection = qs("#viewerSection");
      viewerSection.parentNode.insertBefore(area, viewerSection.nextSibling);
    }
  }

  // Delegated drop handlers so a dragged region chip can be copied into a stash.
  // Bound once on the persistent #stashArea node (its innerHTML is rebuilt each
  // render, but the element itself is reused).
  var _stashDragOverCard = null;

  function clearStashDragIndicators() {
    if (_stashDragOverCard) {
      _stashDragOverCard.classList.remove("drag-over");
      _stashDragOverCard = null;
    }
    qsa(".stash-card.drag-over").forEach(function (card) {
      card.classList.remove("drag-over");
    });
  }

  function bindStashDrop(area) {
    area.addEventListener("dragover", function (e) {
      if (!hasRegionDragPayload(e)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "copy";
      var card = e.target.closest(".stash-card");
      if (_stashDragOverCard && _stashDragOverCard !== card) {
        _stashDragOverCard.classList.remove("drag-over");
      }
      if (card) {
        card.classList.add("drag-over");
        _stashDragOverCard = card;
      }
    });
    area.addEventListener("dragleave", function (e) {
      var card = e.target.closest(".stash-card");
      if (card && !card.contains(e.relatedTarget)) {
        card.classList.remove("drag-over");
        if (_stashDragOverCard === card) _stashDragOverCard = null;
      }
    });
    area.addEventListener("drop", function (e) {
      if (!hasRegionDragPayload(e)) return;
      var card = e.target.closest(".stash-card");
      if (!card) return;
      e.preventDefault();
      clearStashDragIndicators();
      _regionDragDropped = true;
      var regionName = getDraggedRegionName(e);
      if (regionName) copyRegionToStash(regionName, card.dataset.stashId);
    });
  }

  // ---- Region chip drag (reorder within #regionChips + copy into stashes) ----
  // Horizontal mirror of the multitool/task vertical drag helpers: midpoints use
  // left+width/2 and compare clientX. Excluding the .dragging chip from the cache
  // keeps the drop index aligned with the post-splice array (no off-by-one).
  var REGION_DRAG_MIME = "application/x-region-name";
  var _regionDragMidpoints = null;
  var _regionDragOverRaf = null;
  var _regionPendingDragOverX = null;
  var _regionDragActive = false;
  var _regionDragMoved = false;
  var _regionDragDropped = false;

  function dataTransferHasType(dt, type) {
    if (!dt || !dt.types) return false;
    if (typeof dt.types.indexOf === "function") return dt.types.indexOf(type) >= 0;
    if (typeof dt.types.contains === "function") return dt.types.contains(type);
    for (var i = 0; i < dt.types.length; i++) {
      if (dt.types[i] === type) return true;
    }
    return false;
  }

  function hasRegionDragPayload(e) {
    // Firefox/Safari may hide custom MIME types during dragover. Since these
    // drags start inside this document, the local flag is the reliable signal.
    return _regionDragActive || dataTransferHasType(e.dataTransfer, REGION_DRAG_MIME);
  }

  function setRegionDragData(dt, idx, name) {
    if (!dt) return;
    dt.setData("text/plain", String(idx));
    try {
      dt.setData(REGION_DRAG_MIME, name);
    } catch (_) {
      // Some engines reject custom types; text/plain + local state covers us.
    }
  }

  function getDraggedRegionName(e) {
    var name = "";
    try {
      name = e.dataTransfer.getData(REGION_DRAG_MIME);
    } catch (_) {
      name = "";
    }
    if (name && Object.prototype.hasOwnProperty.call(state.regions, name)) return name;
    var fromIdx = parseInt(e.dataTransfer.getData("text/plain"), 10);
    if (isNaN(fromIdx)) return "";
    return Object.keys(state.regions)[fromIdx] || "";
  }

  function getDraggedRegionIndex(e) {
    var fromIdx = parseInt(e.dataTransfer.getData("text/plain"), 10);
    if (!isNaN(fromIdx)) return fromIdx;
    var name = getDraggedRegionName(e);
    return name ? Object.keys(state.regions).indexOf(name) : -1;
  }

  function _cacheRegionDragMidpoints(container) {
    var chips = container.querySelectorAll(".region-chip:not(.dragging)");
    var mids = new Array(chips.length);
    for (var i = 0; i < chips.length; i++) {
      var r = chips[i].getBoundingClientRect();
      mids[i] = r.left + r.width / 2;
    }
    _regionDragMidpoints = mids;
  }

  function getRegionDropIndex(container, clientX) {
    var mids = _regionDragMidpoints;
    if (!mids) {
      _cacheRegionDragMidpoints(container);
      mids = _regionDragMidpoints;
    }
    for (var i = 0; i < mids.length; i++) {
      if (clientX < mids[i]) return i;
    }
    return mids.length;
  }

  function clearRegionDragIndicators(container) {
    var chips = container.querySelectorAll(".region-chip.drag-over");
    for (var i = 0; i < chips.length; i++) chips[i].classList.remove("drag-over");
    container.classList.remove("drag-over-append");
  }

  function initRegionDrag() {
    var chips = qs("#regionChips");

    chips.addEventListener("dragstart", function (e) {
      var chip = e.target.closest(".region-chip");
      if (!chip) {
        e.preventDefault();
        return;
      }
      _regionDragActive = true;
      _regionDragMoved = false;
      _regionDragDropped = false;
      chip.classList.add("dragging");
      setRegionDragData(e.dataTransfer, chip.dataset.regionIdx, chip.dataset.regionName);
      e.dataTransfer.effectAllowed = "copyMove";
      _cacheRegionDragMidpoints(chips);
    });

    chips.addEventListener("dragend", function (e) {
      var chip = e.target.closest(".region-chip");
      if (chip) chip.classList.remove("dragging");
      if (_regionDragOverRaf != null) {
        cancelAnimationFrame(_regionDragOverRaf);
        _regionDragOverRaf = null;
      }
      _regionPendingDragOverX = null;
      clearRegionDragIndicators(chips);
      clearStashDragIndicators();
      _regionDragMidpoints = null;
      if (_regionDragMoved || _regionDragDropped) {
        state.regionSuppressNextClick = true;
        setTimeout(function () { state.regionSuppressNextClick = false; }, 250);
      }
      _regionDragActive = false;
      _regionDragMoved = false;
      _regionDragDropped = false;
    });

    chips.addEventListener("dragover", function (e) {
      if (!hasRegionDragPayload(e)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      _regionDragMoved = true;
      _regionPendingDragOverX = e.clientX;
      if (_regionDragOverRaf != null) return; // RAF-debounce ~60Hz dragover
      _regionDragOverRaf = requestAnimationFrame(function () {
        _regionDragOverRaf = null;
        if (_regionPendingDragOverX == null) return;
        clearRegionDragIndicators(chips);
        var visible = chips.querySelectorAll(".region-chip:not(.dragging)");
        var idx = getRegionDropIndex(chips, _regionPendingDragOverX);
        if (idx < visible.length) visible[idx].classList.add("drag-over");
        else chips.classList.add("drag-over-append");
      });
    });

    chips.addEventListener("dragleave", function (e) {
      var chip = e.target.closest(".region-chip");
      if (chip) chip.classList.remove("drag-over");
      if (!chips.contains(e.relatedTarget)) chips.classList.remove("drag-over-append");
    });

    chips.addEventListener("drop", function (e) {
      if (!hasRegionDragPayload(e)) return;
      e.preventDefault();
      _regionDragDropped = true;
      clearRegionDragIndicators(chips);
      var fromIdx = getDraggedRegionIndex(e);
      if (fromIdx < 0) return;
      var toIdx = getRegionDropIndex(chips, e.clientX);
      if (fromIdx === toIdx) return;
      var names = Object.keys(state.regions);
      var previousRegions = state.regions;
      var moved = names.splice(fromIdx, 1)[0];
      if (!moved) return;
      names.splice(toIdx, 0, moved);
      // Rebuild state.regions in the new order.
      var reordered = {};
      names.forEach(function (n) {
        reordered[n] = state.regions[n];
      });
      state.regions = reordered;
      // Region colors are position-based (regionColorForIndex), so reordering
      // recolors regions — intentional. Repaint chips and overlay together.
      renderRegionChips();
      renderOverlay();
      apiPut("api/regions/reorder", { names: names }).then(function (data) {
        if (!data || !data.ok) throw new Error((data && data.error) || "reorder failed");
      }).catch(function () {
        state.regions = previousRegions;
        renderRegionChips();
        renderOverlay();
        showToast("Failed to save region order");
      });
    });

    document.addEventListener("dragend", function () {
      clearStashDragIndicators();
    });
    document.addEventListener("drop", function () {
      clearStashDragIndicators();
    });
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

  // renderOverlay lives in screenspace-overlay.js; this hub delegator forwards
  // its ~33 hub call sites to the satellite implementation.
  function renderOverlay() {
    return SS.renderOverlay && SS.renderOverlay.apply(null, arguments);
  }

  // ---- Timeline — screenspace-timeline.js ----
  // The timeline canvas (ruler, zoom/pan, scrubbing, markers, boundary flags,
  // playhead, tooltips, legend) lives in screenspace-timeline.js. The hub keeps
  // same-named delegators for the entry points its own code calls; the tasks and
  // results satellites destructure SS.renderTimeline / SS.updateMarkerInfo at
  // load (timeline loads before them). seekPlayhead/loadFrame stay here (frame
  // viewer) and the satellite reaches them via SS.
  function initTimeline() { return SS.initTimeline && SS.initTimeline.apply(null, arguments); }
  function renderTimeline() { return SS.renderTimeline && SS.renderTimeline.apply(null, arguments); }
  function renderPlayhead() { return SS.renderPlayhead && SS.renderPlayhead.apply(null, arguments); }
  // ---- Tool info tooltip ----

  var TOOL_INFO = {
    multitool: "Combines several tools so a frame only matches when it passes every step. For example, a red health bar AND the word 'DEAD'. Add at least two steps; a step can also be set to exclude (match only when it does NOT apply). Get each tool working on its own first, then chain them here to pin down precise moments.",
    color: "Finds frames where the average color of your region matches a color you pick. Draw a small region over a solid-colored element and sample its color; widen Tolerance to catch more shades, tighten it to be stricter. Good for color-coded elements like a health bar or status light. To find a specific icon or picture instead, use Template.",
    change: "Flags frames where the picture inside your region differs from a moment earlier: sudden changes such as a screen transition, a pop-up appearing, or a loading screen finishing. Raise the Threshold if it fires on every small flicker. Unlike Flow (which measures movement) it reacts to any difference; unlike Boundary it watches only the region you draw, not the whole screen.",
    similarity: "Capture one reference frame, then this finds every later frame that looks almost identical: a strict, pixel-for-pixel match that's sensitive to lighting and layout shifts. Lower the Threshold to allow looser matches. Use it to catch when one exact state returns (a specific dialog or menu). For 'which screen are we on' across several screens that vary, use Scene instead.",
    text: "Reads on-screen text in your region (OCR) and flags frames matching your search words, allowing for small misreads. Draw a tight region around the text; raise the OCR confidence if you get false hits. Good for catching specific labels, error messages, or button text. To compare on-screen numbers (e.g. score over 1000), use Numbers.",
    numbers: "Reads a number from your region (OCR) and flags frames where it meets a rule you set: equals, greater than, less than, or within a range. Draw a tight region around just the number, then pick the operator and target value. Great for scores, timers, lives, or any changing count. For words rather than numbers, use Text.",
    timelapse: "Produces one sped-up video or GIF of your region over the time range you choose: a fast way to skim a long session. Unlike every other tool it doesn't mark individual moments on the timeline; it outputs a single clip. Set the speed, and optionally sample every N seconds for a shorter file.",
    template: "Capture or upload a small reference image, then this looks for that exact picture anywhere on screen, not just inside a region. Ideal for finding an icon, button, or logo wherever it appears. Lower the Threshold to allow looser matches. Unlike Color (which matches an average shade) it matches the picture itself; unlike Similarity it searches the whole frame, not one region.",
    flow: "Detects movement inside your region: a character running, an animation playing, or activity in one corner. Raise the strength threshold to ignore small or slow motion. Unlike Change (which fires on any pixel difference, including flicker) Flow responds only to real movement, so it stays steadier on noisy footage.",
    scene: "Capture and label several reference screens, then this tags each frame with whichever one it most resembles. This builds a timeline of which screen is showing (title, map, level, pause menu). It tolerates lighting and minor changes better than Similarity, and handles many screens at once where Similarity matches just one. Lower the Threshold if frames go untagged.",
    inactivity: "Finds stretches where your region barely changes for a while — loading screens, frozen states, or a player standing idle. It's the opposite of Change: it fires when nothing happens, not when something does. Set the minimum duration so brief pauses are ignored and only real stalls are reported.",
    boundary: "Scans the whole screen for period transitions: menu to gameplay, a level loading, a loading screen ending. Metric: Auto (recommended) uses a content fingerprint and only marks a change that holds for a moment and is backed by a hard cut, so camera motion and brief overlays don't fragment one continuous period; pHash is the simpler 'any big frame-to-frame jump' detector. Sensitivity tunes the hard-cut threshold; Min gap avoids clustered markers during fast action. After scanning, near-identical periods are merged and transient blips dissolved. These are orientation markers, not clip candidates; unlike Scene, it doesn't label the screens; it only marks where they change.",
    attention: "Predicts where a viewer is probably looking, with no eye-tracking hardware. It scores each sampled frame for what draws the eye (strong contrast, movement, unusual detail) and turns the whole scan into heatmaps: a static image, an accumulation animation, and a rolling replay similar to an eye-tracking gaze video. The timeline only gets a marker at an attention shift, when the predicted focus jumps to a different part of the screen. Raise Sensitivity so only big jumps count; raise Smoothing if shifts lag behind the action. The weight sliders control what counts as eye-catching, and the Model view re-renders live while you drag them. It works from the visuals alone, so treat the output as an informed guess about attention rather than a measurement."
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
  // ---- Workflow tabs + params ----

  function initWorkflowTabs() {
    qsa(".wf-tab").forEach(function (tab, i) {
      // Alt-hold hint: tabs 1–9 map to the selectTool digit combos (the flat
      // row's digit shortcuts). Tabs 10+ have no digit (only 9 combos exist).
      if (i < 9) {
        tab.setAttribute("data-hotkey", "screenspace.selectTool");
        tab.setAttribute("data-hotkey-combo", String(i));
      }
      tab.addEventListener("click", function () {
        hideToolInfoTooltip(true);
        state.activeWorkflow = tab.dataset.type;
        setStoredUIStateField("screenspace", "activeWorkflow", state.activeWorkflow);
        qsa(".wf-tab").forEach(function (t) { t.classList.remove("active"); });
        tab.classList.add("active");
        renderWorkflowParams();
        updateRunButton();
        // Every selection path (tab click, category-delegated click, cycleTool,
        // session restore) funnels through here, so this keeps the grouped
        // category nav's active chip in sync however the tool was chosen.
        syncToolCategoryNav();
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

  // Cycle the active tool (the Z/X hotkeys). Layout-aware: in grouped mode it
  // walks the category order (left-to-right, tool-by-tool within a category) so
  // the traversal matches what the user sees; in flat-tab mode it walks the tab
  // DOM order. Both delegate to selectWorkflowType/.click() so the persist +
  // param re-render + Run-button refresh run through the one existing path.
  function cycleTool(delta) {
    if (state.groupedToolNav) {
      var order = [];
      TOOL_CATEGORIES.forEach(function (c) { order.push.apply(order, c.tools); });
      var idx = order.indexOf(state.activeWorkflow);
      if (idx === -1) idx = 0;
      selectWorkflowType(order[(idx + delta + order.length) % order.length]);
      return;
    }
    var tabs = qsa(".wf-tab");
    if (!tabs.length) return;
    var cur = 0;
    for (var i = 0; i < tabs.length; i++) {
      if (tabs[i].classList.contains("active")) { cur = i; break; }
    }
    tabs[(cur + delta + tabs.length) % tabs.length].click();
  }

  // ---- Grouped category tool nav (SCREENSPACE_GROUPED_TOOL_NAV) ----
  // An optional presentational layer over the flat .wf-tab row. The tabs stay
  // in the DOM (hidden via CSS) in both modes; a category selection just
  // delegates to the matching tab's .click(), so state, persistence, param
  // rendering, cycleTool, and restore all run through the one existing path.
  // Ordered source of truth for grouping + order. Every category renders as a
  // dropdown (even single-tool ones, for a uniform look). Multitool is the lone
  // standalone chip — a direct-select button (no dropdown), alwaysIcon so its
  // icon shows even when unselected, since it is unlike the detector categories.
  var TOOL_CATEGORIES = [
    { label: "Multitool", tools: ["multitool"], alwaysIcon: true, standalone: true },
    { label: "Difference", tools: ["change", "similarity", "inactivity"] },
    { label: "Detection", tools: ["template", "color", "text", "numbers"] },
    { label: "Classification", tools: ["scene", "boundary"] },
    { label: "Attention", tools: ["flow", "attention"] },
    { label: "Utility", tools: ["timelapse"] },
  ];

  // Heroicon basenames per tool — mirrors the .ss-task-icon--<type> mask map in
  // screenspace.css. Used for the command-palette tool-switch entries (the chip
  // glyphs use the CSS masks via buildTypeIcon, not these names).
  var TOOL_ICON_NAMES = {
    multitool: "wrench-screwdriver", color: "eye-dropper", change: "bolt",
    similarity: "photo", text: "language", numbers: "hashtag",
    template: "viewfinder-circle", flow: "arrows-right-left", scene: "squares-2x2",
    inactivity: "pause-circle", boundary: "flag", timelapse: "forward",
    attention: "eye",
  };

  var _catNavBuilt = false;
  var _catOutsideBound = false;

  function _toolLabel(type) {
    return type ? type.charAt(0).toUpperCase() + type.slice(1) : "";
  }

  // Alt-hold digit hints for the category chips. While a dropdown is open the
  // digits route into that menu's items (see handleToolDigit), so the chip
  // hints are removed and the open menu's items carry the 1..N hints instead
  // (items hold data-hotkey from build time and only render while visible).
  function _setCatChipHints(enabled) {
    qsa("#workflowCategories .ss-cat-chip").forEach(function (chip, i) {
      var trig = chip.querySelector(".ss-cat-trigger");
      if (!trig) return;
      if (enabled && i < 9) {
        trig.setAttribute("data-hotkey", "screenspace.selectTool");
        trig.setAttribute("data-hotkey-combo", String(i));
      } else {
        trig.removeAttribute("data-hotkey");
      }
    });
  }

  function _refreshCatHints() {
    _setCatChipHints(!qs("#workflowCategories .ss-cat-chip.open"));
  }

  function closeCatMenus(except) {
    qsa("#workflowCategories .ss-cat-chip.open").forEach(function (chip) {
      if (chip !== except) {
        chip.classList.remove("open");
        var trig = chip.querySelector(".ss-cat-trigger");
        if (trig) trig.setAttribute("aria-expanded", "false");
      }
    });
    _refreshCatHints();
  }

  function openCatMenu(chip) {
    closeCatMenus(chip);
    chip.classList.add("open");
    var trig = chip.querySelector(".ss-cat-trigger");
    if (trig) trig.setAttribute("aria-expanded", "true");
    _refreshCatHints();
  }

  function toggleCatMenu(chip) {
    if (chip.classList.contains("open")) closeCatMenus(null);
    else openCatMenu(chip);
  }

  // Delegate to the hidden flat tab so the whole existing selection path runs
  // (state.activeWorkflow, persistence, renderWorkflowParams, updateRunButton,
  // and the syncToolCategoryNav() call at the tail of the tab handler).
  function selectWorkflowType(type) {
    var tab = qs('.wf-tab[data-type="' + type + '"]');
    if (tab) tab.click();
  }

  function buildToolCategoryNav() {
    var nav = qs("#workflowCategories");
    if (!nav) return;
    nav.innerHTML = "";
    var frag = document.createDocumentFragment();
    TOOL_CATEGORIES.forEach(function (cat) {
      // Every category is a dropdown (uniform look); only the standalone
      // Multitool is a direct-select chip.
      var isDropdown = !cat.standalone;
      // Wrapper (div) + trigger (button) + sibling menu — menu items must NOT
      // nest inside a <button> (invalid HTML). Mirrors #exportEventsWrap.
      var chip = el("div", "ss-cat-chip");
      chip.setAttribute("data-cat", cat.label);
      chip.setAttribute("data-tools", cat.tools.join(","));
      if (cat.alwaysIcon) chip.setAttribute("data-always-icon", "");
      var trigger = el("button", "ss-cat-trigger");
      trigger.type = "button";
      trigger.appendChild(el("span", "ss-cat-glyph"));
      trigger.appendChild(el("span", "ss-cat-text"));
      if (isDropdown) {
        trigger.setAttribute("aria-haspopup", "menu");
        trigger.setAttribute("aria-expanded", "false");
        trigger.appendChild(el("span", "ss-cat-chevron"));
        trigger.addEventListener("click", function (e) {
          e.stopPropagation();
          toggleCatMenu(chip);
        });
        chip.appendChild(trigger);
        var menu = el("div", "ss-cat-menu"); // visibility driven by chip.open
        menu.setAttribute("role", "menu");
        cat.tools.forEach(function (type, ti) {
          var item = el("button", "ss-cat-item");
          item.type = "button";
          item.setAttribute("data-type", type);
          item.setAttribute("role", "menuitem");
          // Alt-hold hint: while the dropdown is open, digit ti+1 selects this
          // item. Only rendered when the menu is visible (open chip).
          if (ti < 9) {
            item.setAttribute("data-hotkey", "screenspace.selectTool");
            item.setAttribute("data-hotkey-combo", String(ti));
          }
          var icon = buildTypeIcon(type);
          if (icon) item.appendChild(icon);
          item.appendChild(el("span", "ss-cat-item-label", _toolLabel(type)));
          item.addEventListener("click", function (e) {
            e.stopPropagation();
            closeCatMenus(null);
            selectWorkflowType(type);
          });
          menu.appendChild(item);
        });
        chip.appendChild(menu);
      } else {
        // Direct-select chip (single-tool category or standalone Multitool).
        trigger.addEventListener("click", function (e) {
          e.stopPropagation();
          closeCatMenus(null);
          selectWorkflowType(cat.tools[0]);
        });
        chip.appendChild(trigger);
      }
      frag.appendChild(chip);
    });
    nav.appendChild(frag);
    if (!_catOutsideBound) {
      document.addEventListener("click", function () { closeCatMenus(null); });
      _catOutsideBound = true;
    }
    syncToolCategoryNav();
    _refreshCatHints();
  }

  // Reflect state.activeWorkflow in the category chips: the owning segment gets
  // the solid tool-color fill (via data-active-type) and shows just the active
  // tool's icon + name (the category name is dropped to keep chips compact and
  // equally sized). Resting chips show the category name; the alwaysIcon chip
  // (Multitool) keeps its icon even when resting.
  function syncToolCategoryNav() {
    var nav = qs("#workflowCategories");
    if (!nav) return;
    var active = state.activeWorkflow;
    qsa("#workflowCategories .ss-cat-chip").forEach(function (chip) {
      var tools = (chip.getAttribute("data-tools") || "").split(",");
      var cat = chip.getAttribute("data-cat") || "";
      var alwaysIcon = chip.hasAttribute("data-always-icon");
      var isActive = tools.indexOf(active) !== -1;
      var glyph = chip.querySelector(".ss-cat-glyph");
      var text = chip.querySelector(".ss-cat-text");
      chip.classList.toggle("active", isActive);
      // Glyph: the active tool's icon when active; the (single) tool's icon on
      // an alwaysIcon resting chip; empty otherwise.
      var glyphType = isActive ? active : (alwaysIcon ? tools[0] : null);
      if (glyph) {
        glyph.innerHTML = "";
        if (glyphType) {
          var icon = buildTypeIcon(glyphType);
          if (icon) glyph.appendChild(icon);
        }
      }
      if (isActive) {
        chip.setAttribute("data-active-type", active);
        if (text) text.textContent = _toolLabel(active);
      } else {
        chip.removeAttribute("data-active-type");
        if (text) text.textContent = cat;
      }
      chip.querySelectorAll(".ss-cat-item").forEach(function (item) {
        item.classList.toggle("active", isActive && item.getAttribute("data-type") === active);
      });
    });
  }

  // Numeral hotkey (1–9) tool selection — ISO-friendly, works in both modes.
  // Old tab mode: digit N selects the Nth tool tab. Grouped mode: digit N
  // selects the Nth segment; a dropdown category opens and the next digit
  // selects the Nth tool within it (hotkey → numeral); the standalone Multitool
  // selects directly.
  function handleToolDigit(n) {
    if (!n || n < 1) return;
    if (!state.groupedToolNav) {
      var tabs = qsa(".wf-tab");
      if (tabs[n - 1]) tabs[n - 1].click();
      return;
    }
    var openChip = qs("#workflowCategories .ss-cat-chip.open");
    if (openChip) {
      var items = openChip.querySelectorAll(".ss-cat-item");
      if (items[n - 1]) items[n - 1].click(); // selects + closes
      return;
    }
    var chips = qsa("#workflowCategories .ss-cat-chip");
    var chip = chips[n - 1];
    if (!chip) return;
    if (chip.querySelector(".ss-cat-menu")) openCatMenu(chip);
    else selectWorkflowType((chip.getAttribute("data-tools") || "").split(",")[0]);
  }

  // Switch between the grouped category nav and the flat tab row based on
  // state.groupedToolNav (SCREENSPACE_GROUPED_TOOL_NAV). Builds the category
  // nav lazily on first enable.
  function applyToolNavMode() {
    var section = qs("#workflowSection");
    if (!section) return;
    var grouped = !!state.groupedToolNav;
    if (grouped && !_catNavBuilt) {
      buildToolCategoryNav();
      _catNavBuilt = true;
    }
    section.classList.toggle("ss-grouped-tools", grouped);
    var nav = qs("#workflowCategories");
    if (nav) nav.setAttribute("aria-hidden", grouped ? "false" : "true");
    if (grouped) syncToolCategoryNav();
    else closeCatMenus(null);
  }

  function renderIntervalSlot(inputId, min, max, def, step) {
    var slot = qs("#workflowIntervalSlot");
    if (!slot) return;
    slot.innerHTML = "";
    slot.setAttribute("data-tooltip", "Interval (seconds)");
    var iconWrap = el("div", "interval-icon");
    var iconMask = el("span", "interval-icon-mask");
    applyIconMask(iconMask, "clock", "/screenspace/icons/");
    iconWrap.appendChild(iconMask);
    slot.appendChild(iconWrap);
    var ctrl = el("div", "param-control");
    ctrl.appendChild(numberInput(inputId, min, max, def, step));
    slot.appendChild(ctrl);
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
    SS.setColorHiddenInputs({ h: hiddenH, s: hiddenS, v: hiddenV, hex: hexInput });

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

    // Match mode (average vs presence) + the presence-only "Min area" row.
    var minAreaRow;
    var modeRow = el("div", "param-row");
    modeRow.appendChild(el("span", "param-label", "Mode"));
    var modeControl = el("div", "param-control");
    modeControl.appendChild(
      buildColorModeControl("paramColorMode", "average", false, function (mode) {
        if (minAreaRow) minAreaRow.classList.toggle("hidden", mode !== "presence");
      })
    );
    modeRow.appendChild(modeControl);
    container.appendChild(modeRow);
    minAreaRow = addParamRow(
      container, "Min area %", rangeInput("paramColorMinArea", 0, 100, 1, 1), "paramColorMinAreaVal"
    );
    minAreaRow.id = "paramColorMinAreaRow";
    minAreaRow.classList.add("hidden");
    // Widen the readout so it can show "X% · ~N px" / "Any presence …" beside
    // the slider; addParamRow's generic listener runs first and writes the raw
    // value, ours runs after and replaces it with the region-aware readout.
    var minAreaVal = qs("#paramColorMinAreaVal");
    if (minAreaVal) minAreaVal.classList.add("param-value--minarea");
    var minAreaSlider = qs("#paramColorMinArea");
    if (minAreaSlider) {
      minAreaSlider.addEventListener("input", function () {
        _updateMinAreaReadout("");
      });
    }
    _updateMinAreaReadout("");

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
    addParamRow(container, "Fuzzy Thr.", rangeInput("paramTextFuzzy", 0.50, 1.00, numberOrDefault(CLIPGEN_CONFIG.screenspaceOcrFuzzyThreshold, 0.75), 0.01), "paramTextFuzzyVal");
    addParamRow(container, "Min OCR conf.", rangeInput("paramTextOcrConf", 0.00, 1.00, numberOrDefault(CLIPGEN_CONFIG.screenspaceOcrMinConfidence, 0.6), 0.01), "paramTextOcrConfVal");
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
    var ppCb = document.createElement("input");
    ppCb.type = "checkbox";
    ppCb.id = "paramTextOcrPreprocess";
    addParamRow(container, "Enhance ROI", ppCb);
    addParamRow(container, "Normalize", buildNormalizeControl("paramTextOcrNormalize", "off"));
    addParamRow(container, "Consecutive", numberInput("paramTextConsecutive", 1, 10, 1, 1));
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
    addParamRow(container, "Min OCR conf.", rangeInput("paramNumOcrConf", 0.00, 1.00, numberOrDefault(CLIPGEN_CONFIG.screenspaceOcrMinConfidence, 0.6), 0.01), "paramNumOcrConfVal");
    var ppCb = document.createElement("input");
    ppCb.type = "checkbox";
    ppCb.id = "paramNumOcrPreprocess";
    addParamRow(container, "Enhance ROI", ppCb);
    var ioCb = document.createElement("input");
    ioCb.type = "checkbox";
    ioCb.id = "paramNumIntegersOnly";
    addParamRow(container, "Integers only", ioCb);
    renderIntervalSlot("paramNumInterval", 0.5, 60, 2.0, 0.5);
    addParamRow(container, "Consecutive", numberInput("paramNumConsecutive", 1, 10, 1, 1));
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
      uploadThumb.decoding = "async";
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

  // Param inputs are DOM-only (rangeInput/numberInput carry hardcoded defaults,
  // not state), so any *same-tool* rebuild of the panel — Capture Current Frame,
  // a scene/template add, a picker toggle — would snap every value (Interval,
  // thresholds, …) back to its default. Snapshot values by id and restore them
  // across the rebuild. Tool *switches* still reset (ids are tool-prefixed, so
  // they don't match). Multitool step params are state-backed (step._initial) and
  // its per-step ids are positional, so for multitool we restore only the shared
  // #workflowIntervalSlot, never the step rows.
  var _lastRenderedTool = null;

  function _snapshotParamValues(intervalOnly) {
    var map = {};
    var sel = intervalOnly
      ? "#workflowIntervalSlot [id]"
      : "#workflowParams [id], #workflowIntervalSlot [id]";
    qsa(sel).forEach(function (el) {
      var tag = el.tagName;
      if (tag === "INPUT" || tag === "SELECT" || tag === "TEXTAREA") {
        map[el.id] = el.type === "checkbox" ? el.checked : el.value;
      }
    });
    return map;
  }

  function _restoreParamValues(map) {
    Object.keys(map).forEach(function (id) {
      var el = document.getElementById(id);
      if (!el) return;
      var tag = el.tagName;
      if (tag !== "INPUT" && tag !== "SELECT" && tag !== "TEXTAREA") return;
      var saved = map[id];
      if (el.type === "checkbox") {
        if (el.checked === saved) return;
        el.checked = saved;
      } else {
        if (el.value === String(saved)) return;
        el.value = saved;
      }
      // Fire input so value readouts + the model view reflect the restored value.
      el.dispatchEvent(new Event("input", { bubbles: true }));
    });
  }

  function renderWorkflowParams() {
    var sameTool = _lastRenderedTool === state.activeWorkflow;
    var saved = sameTool ? _snapshotParamValues(state.activeWorkflow === "multitool") : null;
    _renderWorkflowParamsBuild();
    _lastRenderedTool = state.activeWorkflow;
    if (saved) _restoreParamValues(saved);
  }

  function _renderWorkflowParamsBuild() {
    var container = qs("#workflowParams");
    container.innerHTML = "";
    SS.setColorHiddenInputs(null);
    hideToolInfoTooltip(true);
    _toolInfoPinned = false;
    var intervalSlot = qs("#workflowIntervalSlot");
    if (intervalSlot) intervalSlot.innerHTML = "";
    var type = state.activeWorkflow;

    var regionPickerWrap = qs("#runRegionPicker");
    // Multitool uses per-step regions; boundary is full-frame only — both hide
    // the global region picker.
    if (regionPickerWrap) {
      regionPickerWrap.style.display =
        type === "multitool" || type === "boundary" || type === "attention"
          ? "none" : "";
    }

    if (type === "multitool") {
      SS.renderMultitoolParams(container);
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
      updateCalibrationVisibility();
      // Drop the prior tool's scores so the strip doesn't render a stale axis;
      // multitool returns early, so mirror the bottom-of-function cleanup here.
      // renderCalibration() also resets calibrationGreen and the Run tooltip
      // synchronously, before the async re-eval arrives.
      state.calibrationResult = null;
      renderCalibration();
      if (!state.suppressCalibrationRefresh) refreshCalibration();
      return;
    }

    if (type === "color") renderColorParams(container);
    else if (type === "change") {
      addParamRow(container, "Threshold", rangeInput("paramChangeThresh", 0.01, 0.50, 0.03, 0.01), "paramChangeThreshVal");
      addParamRow(container, "Noise Thr.", rangeInput("paramChangeNoise", 0, 100, 30, 1), "paramChangeNoiseVal");
      renderIntervalSlot("paramChangeInterval", 0.5, 60, 1.0, 0.5);
      addParamRow(container, "Consecutive", numberInput("paramChangeConsecutive", 1, 10, 1, 1));
    }
    else if (type === "similarity") renderSimilarityParams(container);
    else if (type === "text") renderTextParams(container);
    else if (type === "numbers") renderNumbersParams(container);
    else if (type === "timelapse") renderTimelapseParams(container);
    else if (type === "template") renderTemplateParams(container);
    else if (type === "flow") {
      addParamRow(container, "Magnitude", rangeInput("paramFlowMag", 0.5, 20.0, 2.0, 0.5), "paramFlowMagVal");
      renderIntervalSlot("paramFlowInterval", 0.5, 60, 1.0, 0.5);
      addParamRow(container, "Consecutive", numberInput("paramFlowConsecutive", 1, 10, 1, 1));
    }
    else if (type === "scene") renderSceneParams(container);
    else if (type === "inactivity") {
      addParamRow(container, "Sensitivity", rangeInput("paramInactThresh", 0, 30, 10, 1), "paramInactThreshVal");
      addParamRow(container, "Min duration (s)", numberInput("paramInactMinDur", 0.5, 60, 2.0, 0.5));
      renderIntervalSlot("paramInactInterval", 0.5, 60, 1.0, 0.5);
    }
    else if (type === "boundary") {
      // Metric: Auto sends nothing (server applies its configured default,
      // currently Hybrid). Scene/Hybrid use a content fingerprint + period
      // model; pHash is the v1 consecutive-frame spike detector.
      var metricSel = document.createElement("select");
      metricSel.id = "paramBoundaryMetric";
      [["", "Auto"], ["scene", "Scene"], ["phash", "pHash"], ["hybrid", "Hybrid"]].forEach(function (pair) {
        var opt = el("option", null, pair[1]);
        opt.value = pair[0];
        metricSel.appendChild(opt);
      });
      addParamRow(container, "Metric", metricSel);
      // Range is the full phash Hamming span (8x8 hash → 0..64); higher values
      // fire only on near-total frame changes, which cuts noise on busy footage.
      // Drives the phash threshold (pHash metric, and Hybrid's spike check).
      addParamRow(container, "Sensitivity", rangeInput("paramBoundaryThresh", 0, 64, 14, 1), "paramBoundaryThreshVal");
      addParamRow(container, "Min gap (s)", numberInput("paramBoundaryMinGap", 0.5, 60, 3.0, 0.5));
      renderIntervalSlot("paramBoundaryInterval", 0.5, 60, 1.0, 0.5);
    }
    else if (type === "attention") {
      // Normalized peak-jump distance for a shift event (fraction of the
      // screen diagonal-ish; 0.15 default) and the EMA alpha for temporal
      // smoothing (1.0 = follow each frame instantly).
      addParamRow(container, "Sensitivity", rangeInput("paramAttnShift", 0.05, 0.50, 0.15, 0.01), "paramAttnShiftVal");
      addParamRow(container, "Smoothing", rangeInput("paramAttnSmooth", 0.1, 1.0, 0.6, 0.05), "paramAttnSmoothVal");
      // Channel weights (defaults mirror SCREENSPACE_ATTENTION_WEIGHT_*).
      // Faces at 0 disables the Haar face channel entirely; the Model view
      // re-renders live as these move, so tuning is visual.
      addParamRow(container, "Spectral wt.", rangeInput("paramAttnWSpectral", 0, 2.0, 1.0, 0.05), "paramAttnWSpectralVal");
      addParamRow(container, "Contrast wt.", rangeInput("paramAttnWContrast", 0, 2.0, 0.7, 0.05), "paramAttnWContrastVal");
      addParamRow(container, "Motion wt.", rangeInput("paramAttnWMotion", 0, 2.0, 1.2, 0.05), "paramAttnWMotionVal");
      addParamRow(container, "Faces wt.", rangeInput("paramAttnWFace", 0, 2.0, 0, 0.05), "paramAttnWFaceVal");
      addParamRow(container, "Center bias", rangeInput("paramAttnCenterBias", 0, 1.0, 0.25, 0.05), "paramAttnCenterBiasVal");
      renderIntervalSlot("paramAttnInterval", 0.5, 60, 0.5, 0.5);
    }

    if (type !== "timelapse") {
      addParamRow(container, "Event label", textInput("paramEventLabel", "e.g. low_health"));
      // Boundary marks period transitions and attention streams shift
      // moments, not discrete detections, so "Detect first" (stop after the
      // first hit) doesn't apply — omit it for both.
      if (type !== "boundary" && type !== "attention") {
        var dfCb = document.createElement("input");
        dfCb.type = "checkbox";
        dfCb.id = "paramDetectFirst";
        addParamRow(container, "Detect first", dfCb);
      }
    }

    var scanPicker = qs("#runScanModePicker");
    // Timelapse produces media (no scan modes); boundary runs its own coarse
    // phash pass and opts out of fast scan, so hide the toggle for both.
    if (scanPicker) {
      scanPicker.style.display = toolSupportsFastScan(type) ? "" : "none";
    }
    var scanBtn = scanPicker && scanPicker.querySelector(".scan-toggle-btn");
    if (scanBtn && scanBtn._updateScanState) scanBtn._updateScanState();

    updateRunButton();
    _updateOverlayUi();
    refreshModelView();
    updateCalibrationVisibility();
    // Drop the prior tool's scores so the strip doesn't briefly render the old
    // axis before the new evaluation returns. renderCalibration() also resets
    // calibrationGreen and the Run tooltip synchronously, before the async
    // re-eval arrives.
    state.calibrationResult = null;
    renderCalibration();
    if (!state.suppressCalibrationRefresh) refreshCalibration();
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
    return row;
  }

  // ---- Model view (impl in screenspace-model-view.js) ----
  // Thin delegators forward the hub's own call sites to the satellite once it
  // has registered SS.*. The satellite also publishes _overlayEligibleForActiveTool
  // (read by screenspace-overlay.js), _updateMinAreaReadout (tasks +
  // multitool-params) and _previewRegionRef (calibration) — see its tail.
  function initModelView() { return SS.initModelView && SS.initModelView.apply(null, arguments); }
  function refreshModelView(opts) { return SS.refreshModelView && SS.refreshModelView.apply(null, arguments); }
  function _updateOverlayUi() { return SS._updateOverlayUi && SS._updateOverlayUi.apply(null, arguments); }
  function _overlayEligibleForActiveTool() { return SS._overlayEligibleForActiveTool && SS._overlayEligibleForActiveTool.apply(null, arguments); }
  function _updateMinAreaReadout(sfx) { return SS._updateMinAreaReadout && SS._updateMinAreaReadout.apply(null, arguments); }

  // ---- Calibration strip (impl in screenspace-calibration.js) ----
  // Thin delegators keep the ~30 hub call sites unchanged and forward to the
  // satellite once it has registered SS.cal*. state.suppressCalibrationRefresh
  // is hub<->tasks coordination: set during restoreTaskToWorkflow (in
  // screenspace-tasks.js) and checked in the single-tool param panels here.
  function refreshCalibration(opts) { return SS.calRefresh && SS.calRefresh(opts); }
  function updateCalibrationThresholdLine() { return SS.calUpdateThresholdLine && SS.calUpdateThresholdLine(); }
  function renderCalibration() { return SS.calRender && SS.calRender(); }
  function updateCalibrationVisibility() { return SS.calVisibility && SS.calVisibility(); }
  function initCalibration() { return SS.calInit && SS.calInit(); }

  // ---- Color picker (impl in screenspace-color.js) ----
  // Thin delegators forward to the color satellite; keeps existing call sites
  // and the sampleColorFromRegion click-handler reference unchanged.
  function updateColorPreview() { return SS.updateColorPreview && SS.updateColorPreview(); }
  function setTargetColor(h, s, v) { return SS.setTargetColor && SS.setTargetColor(h, s, v); }
  function renderColorPalette() { return SS.renderColorPalette && SS.renderColorPalette(); }
  function renderBrightnessStrip() { return SS.renderBrightnessStrip && SS.renderBrightnessStrip(); }
  function sampleColorFromRegion() { return SS.sampleColorFromRegion && SS.sampleColorFromRegion(); }

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
    // A template step with an uploaded image scans the full frame, so it
    // satisfies the per-step region requirement without a region.
    var multitoolHasRegions = multitoolReady && state.multitoolSteps.every(function (s) {
      return !!s.region || (s.type === "template" && s._upload);
    });
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
      var isFullFrameTool = state.activeWorkflow === "boundary"
        || state.activeWorkflow === "attention";
      var hasUploadedTemplate = !!state.uploadedTemplate;
      // Template scans full frames regardless of region selection; the region
      // (or uploaded image) only supplies the template patch.
      var templateMissingPatch = isTemplate && !hasRegion && !hasUploadedTemplate;
      // Boundary and Attention are full-frame only — they need no region at all.
      var nonTemplateMissingRegion = !isTemplate && !isFullFrameTool && !hasRegion;
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
    // Surface calibration agreement as a subtle hover hint on an enabled Run
    // button (no blocking when red — researcher agency preserved). Shown
    // whenever pins are satisfied, independent of the strip's collapsed state;
    // calibrationGreen is only ever true after a green evaluation, which
    // implies pins exist.
    if (!btn.disabled && state.calibrationGreen) {
      btn.setAttribute("data-tooltip", "Calibrated: pins satisfied");
    }
  }

  // ---- Run analysis ----

  function initRunButton() {
    qs("#runBtn").addEventListener("click", function () {
      var type = state.activeWorkflow;
      // Boundary and Attention are full-frame only: always scan the whole
      // frame, ignoring any selected region.
      var isFullFrameTool = type === "boundary" || type === "attention";
      var regions = isFullFrameTool
        ? [fullFrameRegionRef()]
        : (state.runRegions.length > 0
            ? state.runRegions
            : (state.activeRegion ? [activeRegionRef(state.activeRegion)] : []));
      // Multitool uses per-step regions; skip global region requirement
      var isMultitool = type === "multitool";
      // Template with uploaded image can run without a region (full-frame scan)
      if (!isMultitool && !isFullFrameTool && regions.length === 0 && !(type === "template" && state.uploadedTemplate)) return;
      if (regions.length === 0) regions = [""];
      var participants = state.runParticipants.length > 0
        ? state.runParticipants
        : (state.selectedParticipant ? [state.selectedParticipant] : []);
      if (participants.length === 0) return;
      var params = gatherWorkflowParams(type);
      if (params === null) return;
      if (state.scanMode === "fast" && toolSupportsFastScan(type)) params.scan_mode = "fast";

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

  function gatherMultitoolStepParams(stepType, idx, opts) {
    // opts.silent suppresses missing-input toasts so the calibration strip can
    // probe params on every keystroke without spamming the user (the Run path
    // leaves it off and keeps the toasts).
    var silent = !!(opts && opts.silent);
    function toast(msg) { if (!silent) showToast(msg); }
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
      if (((qs("#paramColorMode" + sfx) || {}).value) === "presence") {
        p.color_mode = "presence";
        p.min_coverage = numberOrDefault((qs("#paramColorMinArea" + sfx) || {}).value, 1) / 100;
      }
    } else if (stepType === "change") {
      p.threshold = numberOrDefault((qs("#paramChangeThresh" + sfx) || {}).value, 0.03);
      p.noise_threshold = intOrDefault((qs("#paramChangeNoise" + sfx) || {}).value, 30);
    } else if (stepType === "similarity") {
      var step = state.multitoolSteps[idx];
      if (!step || step._refTs === undefined) {
        toast("Step " + (idx + 1) + ": capture a reference frame first");
        return null;
      }
      p.reference_timestamp = step._refTs;
      p.threshold = numberOrDefault((qs("#paramSimThresh" + sfx) || {}).value, 0.90);
    } else if (stepType === "text") {
      p.search_string = (qs("#paramTextSearch" + sfx) || {}).value || "";
      if (!p.search_string.trim()) {
        toast("Step " + (idx + 1) + ": enter a search string");
        return null;
      }
      p.fuzzy_threshold = numberOrDefault((qs("#paramTextFuzzy" + sfx) || {}).value, CLIPGEN_CONFIG.screenspaceOcrFuzzyThreshold);
      p.ocr_confidence_threshold = numberOrDefault((qs("#paramTextOcrConf" + sfx) || {}).value, CLIPGEN_CONFIG.screenspaceOcrMinConfidence);
      p.ocr_preprocess = !!((qs("#paramTextOcrPreprocess" + sfx) || {}).checked);
      p.ocr_normalize = (qs("#paramTextOcrNormalize" + sfx) || {}).value || "off";
    } else if (stepType === "numbers") {
      p.operator = (qs("#paramNumOperator" + sfx) || {}).value || "gt";
      p.target_value = parseFloat((qs("#paramNumTarget" + sfx) || {}).value);
      if (isNaN(p.target_value)) {
        toast("Step " + (idx + 1) + ": enter a valid target number");
        return null;
      }
      p.ocr_confidence_threshold = numberOrDefault((qs("#paramNumOcrConf" + sfx) || {}).value, CLIPGEN_CONFIG.screenspaceOcrMinConfidence);
      p.ocr_preprocess = !!((qs("#paramNumOcrPreprocess" + sfx) || {}).checked);
      p.integers_only = !!((qs("#paramNumIntegersOnly" + sfx) || {}).checked);
    } else if (stepType === "template") {
      step = state.multitoolSteps[idx];
      if (step && step._upload) {
        p.template_image_data = step._upload.data;
      } else if (step && step._refTs !== undefined) {
        p.reference_timestamp = step._refTs;
      } else {
        toast("Step " + (idx + 1) + ": capture a template frame or upload a PNG");
        return null;
      }
      p.threshold = numberOrDefault((qs("#paramTemplateThresh" + sfx) || {}).value, 0.70);
      var tScalePct = parseFloat((qs("#paramTemplateScale" + sfx) || {}).value);
      if (!isNaN(tScalePct) && tScalePct > 0 && tScalePct !== 100) {
        p.template_scale = tScalePct / 100;
      }
    } else if (stepType === "flow") {
      p.magnitude_threshold = numberOrDefault((qs("#paramFlowMag" + sfx) || {}).value, 2.0);
    } else if (stepType === "scene") {
      step = state.multitoolSteps[idx];
      if (!step || !step._scenes || step._scenes.length === 0) {
        toast("Step " + (idx + 1) + ": add at least one scene reference");
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

  function gatherWorkflowParams(type, opts) {
    // opts.silent suppresses missing-input toasts (used by the calibration
    // strip, which probes params continuously); the Run path omits it.
    var silent = !!(opts && opts.silent);
    function toast(msg) { if (!silent) showToast(msg); }
    var params = {};
    if (type === "multitool") {
      if (state.multitoolSteps.length < 2) {
        toast("Add at least 2 steps");
        return null;
      }
      params.steps = [];
      for (var i = 0; i < state.multitoolSteps.length; i++) {
        var stepP = gatherMultitoolStepParams(state.multitoolSteps[i].type, i, opts);
        if (stepP === null) return null;
        stepP.type = state.multitoolSteps[i].type;
        if (i > 0) {
          stepP.logic = (state.multitoolSteps[i].logic || "AND").toUpperCase();
          var off = state.multitoolSteps[i].offset;
          if (off && isFinite(off.min) && isFinite(off.max)) {
            if (Number(off.min) > Number(off.max)) {
              toast("Step " + (i + 1) + ": offset min must be ≤ max");
              return null;
            }
            stepP.offset = { min: Number(off.min), max: Number(off.max) };
          }
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
      if (((qs("#paramColorMode") || {}).value) === "presence") {
        params.color_mode = "presence";
        params.min_coverage = numberOrDefault((qs("#paramColorMinArea") || {}).value, 1) / 100;
      }
      params.interval = numberOrDefault((qs("#paramColorInterval") || {}).value, 1.0);
    } else if (type === "change") {
      params.threshold = numberOrDefault((qs("#paramChangeThresh") || {}).value, 0.03);
      params.noise_threshold = intOrDefault((qs("#paramChangeNoise") || {}).value, 30);
      params.interval = numberOrDefault((qs("#paramChangeInterval") || {}).value, 1.0);
      var rcChange = intOrDefault((qs("#paramChangeConsecutive") || {}).value, 1);
      if (rcChange > 1) params.require_consecutive = rcChange;
    } else if (type === "similarity") {
      if (state.referenceTimestamp === null) {
        toast("Capture a reference frame first");
        return null;
      }
      params.reference_timestamp = state.referenceTimestamp;
      params.threshold = numberOrDefault((qs("#paramSimThresh") || {}).value, 0.90);
      params.interval = numberOrDefault((qs("#paramSimInterval") || {}).value, 1.0);
    } else if (type === "text") {
      params.search_string = (qs("#paramTextSearch") || {}).value || "";
      if (!params.search_string.trim()) {
        toast("Enter a search string");
        return null;
      }
      params.fuzzy_threshold = numberOrDefault((qs("#paramTextFuzzy") || {}).value, CLIPGEN_CONFIG.screenspaceOcrFuzzyThreshold);
      params.ocr_confidence_threshold = numberOrDefault((qs("#paramTextOcrConf") || {}).value, CLIPGEN_CONFIG.screenspaceOcrMinConfidence);
      params.ocr_preprocess = !!((qs("#paramTextOcrPreprocess") || {}).checked);
      params.ocr_normalize = (qs("#paramTextOcrNormalize") || {}).value || "off";
      params.interval = numberOrDefault((qs("#paramTextInterval") || {}).value, 2.0);
      var lang = (qs("#paramTextLang") || {}).value || "en";
      params.languages = [lang];
      var rcText = intOrDefault((qs("#paramTextConsecutive") || {}).value, 1);
      if (rcText > 1) params.require_consecutive = rcText;
    } else if (type === "numbers") {
      var op = (qs("#paramNumOperator") || {}).value || "gt";
      params.operator = op;
      if (op === "range") {
        params.range_min = parseFloat((qs("#paramNumMin") || {}).value);
        params.range_max = parseFloat((qs("#paramNumMax") || {}).value);
        if (isNaN(params.range_min) || isNaN(params.range_max)) {
          toast("Enter valid min and max values");
          return null;
        }
        if (params.range_min > params.range_max) {
          toast("Min must be less than or equal to max");
          return null;
        }
      } else {
        params.target_value = parseFloat((qs("#paramNumTarget") || {}).value);
        if (isNaN(params.target_value)) {
          toast("Enter a valid target number");
          return null;
        }
      }
      params.ocr_confidence_threshold = numberOrDefault((qs("#paramNumOcrConf") || {}).value, CLIPGEN_CONFIG.screenspaceOcrMinConfidence);
      params.ocr_preprocess = !!((qs("#paramNumOcrPreprocess") || {}).checked);
      params.integers_only = !!((qs("#paramNumIntegersOnly") || {}).checked);
      params.interval = numberOrDefault((qs("#paramNumInterval") || {}).value, 2.0);
      var rcNum = intOrDefault((qs("#paramNumConsecutive") || {}).value, 1);
      if (rcNum > 1) params.require_consecutive = rcNum;
    } else if (type === "timelapse") {
      params.speedup_factor = numberOrDefault((qs("#paramTlSpeed") || {}).value, 10);
      var si = parseFloat((qs("#paramTlSampleInterval") || {}).value);
      if (si > 0) params.sample_interval = si;
      params.output_format = (qs("#paramTlFormat") || {}).value || "mp4";
    } else if (type === "template") {
      if (state.uploadedTemplate) {
        params.template_image_data = state.uploadedTemplate.data;
        if (state.uploadedTemplate.name) params.template_name = state.uploadedTemplate.name;
      } else if (state.referenceTimestamp !== null) {
        params.reference_timestamp = state.referenceTimestamp;
      } else {
        toast("Capture a template region or upload a PNG");
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
      var rcFlow = intOrDefault((qs("#paramFlowConsecutive") || {}).value, 1);
      if (rcFlow > 1) params.require_consecutive = rcFlow;
    } else if (type === "scene") {
      if (state.sceneReferences.length === 0) {
        toast("Add at least one scene reference");
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
    } else if (type === "boundary") {
      params.threshold = intOrDefault((qs("#paramBoundaryThresh") || {}).value, 14);
      params.min_gap = numberOrDefault((qs("#paramBoundaryMinGap") || {}).value, 3.0);
      params.interval = numberOrDefault((qs("#paramBoundaryInterval") || {}).value, 1.0);
      // Auto ("") omits metric so the server applies its configured default.
      var boundaryMetric = (qs("#paramBoundaryMetric") || {}).value || "";
      if (boundaryMetric) params.metric = boundaryMetric;
    } else if (type === "attention") {
      params.shift_threshold = numberOrDefault((qs("#paramAttnShift") || {}).value, 0.15);
      params.ema_alpha = numberOrDefault((qs("#paramAttnSmooth") || {}).value, 0.6);
      params.weight_spectral = numberOrDefault((qs("#paramAttnWSpectral") || {}).value, 1.0);
      params.weight_contrast = numberOrDefault((qs("#paramAttnWContrast") || {}).value, 0.7);
      params.weight_motion = numberOrDefault((qs("#paramAttnWMotion") || {}).value, 1.2);
      params.weight_face = numberOrDefault((qs("#paramAttnWFace") || {}).value, 0);
      params.center_bias = numberOrDefault((qs("#paramAttnCenterBias") || {}).value, 0.25);
      params.interval = numberOrDefault((qs("#paramAttnInterval") || {}).value, 0.5);
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

  // ---- Task queue (impl in screenspace-tasks.js) ----
  // The task-queue / SSE-polling / results-switcher surface lives in
  // screenspace-tasks.js. These thin delegators keep the hub's own call sites
  // (DOMContentLoaded init, timeline render, queue/run path) unchanged and
  // forward to the satellite once it has registered its SS.* entry points.
  function findTask(id) { return SS.findTask && SS.findTask(id); }
  function focusedTaskId() { return SS.focusedTaskId && SS.focusedTaskId(); }
  function renderTaskList() { return SS.renderTaskList && SS.renderTaskList(); }
  function startSSE() { return SS.startSSE && SS.startSSE(); }
  function setRightPaneTab(tab) { return SS.setRightPaneTab && SS.setRightPaneTab(tab); }
  function updateResultsCrumb() { return SS.updateResultsCrumb && SS.updateResultsCrumb(); }
  function initRightPaneTabs() { return SS.initRightPaneTabs && SS.initRightPaneTabs(); }
  function initPauseButton() { return SS.initPauseButton && SS.initPauseButton(); }
  function initTaskQueue() { return SS.initTaskQueue && SS.initTaskQueue(); }
  function initTaskFilters() { return SS.initTaskFilters && SS.initTaskFilters(); }

  // ---- Results (impl in screenspace-results.js) ----
  // The Results panel lives in screenspace-results.js. These thin delegators
  // keep the hub's own call sites (DOMContentLoaded init, histogram-toggle
  // re-render) unchanged and forward to the satellite once it has registered
  // its SS.* entry points.
  function initResultsPanel() { return SS.initResultsPanel && SS.initResultsPanel(); }
  function renderResults() { return SS.renderResults && SS.renderResults(); }

  // ---- Keyboard shortcuts (shared hotkeys.js registry) ----

  function _seekBy(delta) {
    if (!state.videoInfo) return;
    loadFrame(clamp(state.currentTimestamp + delta, 0, Math.max(0, state.videoInfo.duration - 0.001)));
  }

  // ---- Panel focus navigation (Shift+1..4 + arrows) ----
  //
  // Project convention: Shift+numeral targets a panel for keyboard focus; the
  // arrows then rove within it while bare numerals keep selecting tools. The
  // video player is the default surface (arrows = transport seek) that Escape
  // returns to. Selection is a painted cursor (.ss-nav-cursor), NOT real DOM
  // focus — real focus would flip the dispatcher's isTypingTarget check and
  // suppress the arrow handlers. The lone exception is the notes textarea, which
  // takes real focus for editing (Enter to edit, Escape to step back out).

  function ssVideoFocused() {
    return state.focusRegion === "video";
  }

  function ssElVisible(elm) {
    if (!elm) return false;
    var r = elm.getBoundingClientRect();
    return r.width > 0 || r.height > 0;
  }

  // The interactive control of a tool-region nav item: a param row wraps its
  // control (slider / select / checkbox / a button such as "Capture Current
  // Frame") in .param-control; the top-row run controls are already the control.
  function ssToolControl(item) {
    if (item && item.classList && item.classList.contains("param-row")) {
      return item.querySelector("input, select, textarea, button");
    }
    return item;
  }

  // Ordered, currently-visible items the arrows walk in each region.
  function ssNavItems(region) {
    if (region === "sidebar") {
      var items = [];
      var notes = qs("#ssInfoNotes");
      if (notes) items.push(notes);
      // Each collapsible section contributes its header (Enter toggles collapse,
      // so a collapsed Top-issues / Transcript-tags section can still be reached
      // and reopened) followed by its rows while expanded.
      qsa("#ssInfoPanel .ss-info-section").forEach(function (section) {
        if (section.classList.contains("hidden")) return;
        var header = section.querySelector(".ss-info-section-header");
        if (header) items.push(header);
        if (section.getAttribute("data-collapsed") !== "true") {
          qsa("#" + section.id + " li.ss-info-issue").forEach(function (li) { items.push(li); });
        }
      });
      return items;
    }
    if (region === "tool") {
      // The "top row" run controls (participant / interval / region / fast-mode),
      // then the active tool's parameter rows. The tool *selector* is not here:
      // bare numerals and Z/X switch tools, so Shift+2 focuses the panel itself.
      var toolItems = [];
      [
        qs("#runParticipantPicker .run-picker-btn"),
        qs("#workflowIntervalSlot input, #workflowIntervalSlot select"),
        qs("#runRegionPicker .run-picker-btn"),
        qs("#runScanModePicker button"),
      ].forEach(function (ctrl) {
        if (ssElVisible(ctrl)) toolItems.push(ctrl);
      });
      qsa("#workflowParams .param-row").forEach(function (row) { toolItems.push(row); });
      return toolItems;
    }
    if (region === "task") return qsa("#taskList .task-card");
    if (region === "results") return qsa("#resultsList .result-row");
    return [];
  }

  function ssClearNavPaint() {
    qsa(".ss-nav-cursor").forEach(function (n) { n.classList.remove("ss-nav-cursor"); });
  }

  // ---- Run-picker sub-navigation ----
  // Opening a participant/region picker with Enter drops the cursor into the
  // dropdown: arrows walk its options (the Select-all row + per-item checkbox
  // labels) and Enter toggles the focused one. pickerCursor >= 0 means we're
  // inside a dropdown, so the tool-region handlers delegate here.

  function ssOpenPicker() {
    return qs(".run-picker-panel:not(.hidden)");
  }

  function ssInPicker() {
    return state.pickerCursor >= 0 && !!ssOpenPicker();
  }

  function ssPickerItems() {
    var panel = ssOpenPicker();
    if (!panel) return [];
    return Array.prototype.slice.call(panel.querySelectorAll(".run-picker-toggle-all, label"));
  }

  function ssPaintPicker() {
    ssClearNavPaint();
    var items = ssPickerItems();
    if (!items.length) { state.pickerCursor = -1; return; }
    state.pickerCursor = clamp(state.pickerCursor, 0, items.length - 1);
    var cur = items[state.pickerCursor];
    if (cur) {
      cur.classList.add("ss-nav-cursor");
      if (cur.scrollIntoView) cur.scrollIntoView({ block: "nearest" });
    }
  }

  // Repaint the item cursor. Regions re-render their innerHTML (poller, results
  // load, param rebuild), so the index is always re-clamped and never trusted
  // stale; the next nav keypress self-heals a wiped cursor. Only the focused
  // item is highlighted (no whole-panel outline — matches Studio).
  function ssPaintNav() {
    ssClearNavPaint();
    if (state.focusRegion === "video") return;
    var items = ssNavItems(state.focusRegion);
    if (!items.length) { state.focusCursor = 0; return; }
    state.focusCursor = clamp(state.focusCursor, 0, items.length - 1);
    var cur = items[state.focusCursor];
    if (cur) {
      cur.classList.add("ss-nav-cursor");
      if (cur.scrollIntoView) cur.scrollIntoView({ block: "nearest" });
    }
  }

  function ssSetFocusRegion(region) {
    closeRunPicker(); // a transient dropdown doesn't survive a focus-region change
    state.focusRegion = region;
    state.navEditing = false;
    state.pickerCursor = -1;
    if (region === "video") { ssClearNavPaint(); return; }
    state.focusCursor = 0;
    ssPaintNav();
  }

  // Shift+N: reveal the target panel, then land the cursor in it. Declines
  // (stays put) when the panel has nothing to land on, mirroring Studio kbJumpTo.
  function ssFocusRegionByNumber(n) {
    var region;
    if (n === 1) {
      region = "sidebar";
      if (qs("#ssInfoPanel") && qs("#ssInfoPanel").classList.contains("hidden")) {
        applyInfoPanelCollapsed(false);
        setStoredUIStateField("screenspace", "infoPanelCollapsed", false);
      }
    } else if (n === 2) {
      region = "tool";
      if (state.bottomCollapsed) toggleBottomPanel();
    } else if (n === 3) {
      region = "task";
      setRightPaneTab("queue");
    } else if (n === 4) {
      region = "results";
      setRightPaneTab("results");
    } else {
      return;
    }
    if (!ssNavItems(region).length) return;
    // Taking over with the painted cursor: drop any lingering native DOM focus
    // (e.g. a tabbed-to top-nav button) so only one focus indicator shows.
    if (window.ClipgenHotkeys && window.ClipgenHotkeys.blurStrayFocus) {
      window.ClipgenHotkeys.blurStrayFocus();
    }
    ssSetFocusRegion(region);
  }

  function ssNavMove(delta) {
    if (ssInPicker()) {
      var picks = ssPickerItems();
      if (!picks.length) { state.pickerCursor = -1; return; }
      state.pickerCursor = clamp(state.pickerCursor + delta, 0, picks.length - 1);
      ssPaintPicker();
      return;
    }
    var items = ssNavItems(state.focusRegion);
    if (!items.length) return;
    state.focusCursor = clamp(state.focusCursor + delta, 0, items.length - 1);
    ssPaintNav();
  }

  // Left/Right nudges a control by one native step and fires the input event
  // addParamRow listens for (so the model view refreshes). We set .value rather
  // than real-focusing the range, which would double-apply the browser's own
  // arrow stepping. Sliders and number inputs step; selects cycle; buttons,
  // checkboxes and text inputs ignore horizontal (they act on Enter).
  function ssAdjustControl(ctrl, dir) {
    if (!ctrl) return;
    if (ctrl.type === "range" || ctrl.type === "number") {
      var step = parseFloat(ctrl.step) || 1;
      var value = parseFloat(ctrl.value);
      if (isNaN(value)) value = 0;
      value += step * dir;
      if (ctrl.min !== "") value = Math.max(value, parseFloat(ctrl.min));
      if (ctrl.max !== "") value = Math.min(value, parseFloat(ctrl.max));
      value = Math.round(value * 1e6) / 1e6; // trim fractional-step float drift
      ctrl.value = value;
      ctrl.dispatchEvent(new Event("input", { bubbles: true }));
      return;
    }
    if (ctrl.tagName === "SELECT" && ctrl.options.length) {
      ctrl.selectedIndex = Math.max(0, Math.min(ctrl.selectedIndex + dir, ctrl.options.length - 1));
      ctrl.dispatchEvent(new Event("input", { bubbles: true }));
      ctrl.dispatchEvent(new Event("change", { bubbles: true }));
    }
  }

  // Left/Right: adjust the focused item's control. Only the tool region has
  // horizontal controls; the other regions are vertical lists, so horizontal is
  // a consumed no-op — the handler still preventDefaults so the arrows never
  // leak to a video seek.
  function ssNavAdjust(dir) {
    if (state.focusRegion !== "tool" || ssInPicker()) return;
    var cur = ssNavItems("tool")[state.focusCursor];
    if (cur) ssAdjustControl(ssToolControl(cur), dir);
  }

  function ssNavActivate() {
    if (ssInPicker()) {
      var picks = ssPickerItems();
      var pick = picks[state.pickerCursor];
      if (!pick) return;
      var cb = pick.querySelector && pick.querySelector('input[type="checkbox"]');
      if (cb) {
        cb.checked = !cb.checked;
        cb.dispatchEvent(new Event("change", { bubbles: true }));
      } else if (pick.click) {
        pick.click(); // the Select-all / Deselect-all row
      }
      ssPaintPicker(); // reflect the toggle; keep the cursor in place
      return;
    }
    var items = ssNavItems(state.focusRegion);
    var cur = items[state.focusCursor];
    if (!cur) return;
    if (state.focusRegion === "sidebar") {
      if (cur.tagName === "TEXTAREA") {
        state.navEditing = true;
        cur.focus(); // Enter edits the notes; Escape steps back out to the cursor
      } else if (cur.classList.contains("ss-info-section-header")) {
        cur.click();  // toggle the section collapse
        ssPaintNav(); // rows appeared/disappeared; keep the cursor on the header
      } else {
        cur.click(); // clickable cross-ref row -> loadFrame(t)
      }
      return;
    }
    if (state.focusRegion === "task") {
      cur.click(); // selects a completed/paused/running task -> Results tab
      if (state.selectedTaskId) {
        // "Enter moves into Task Results": follow the selection into the panel.
        state.focusRegion = "results";
        state.focusCursor = 0;
        ssPaintNav();
      }
      return;
    }
    if (state.focusRegion === "results") {
      cur.click(); // -> loadFrame(row.dataset.timestamp)
      return;
    }
    // Tool region: operate the focused control. Buttons/checkboxes (run pickers,
    // fast-mode toggle, "Capture Current Frame", "Detect first") activate on
    // Enter; text/number/select controls take real focus so the user can type or
    // open them (Escape steps back out to the cursor). Sliders act on Left/Right.
    var ctrl = ssToolControl(cur);
    if (!ctrl) return;
    if (ctrl.tagName === "BUTTON") {
      var opensPicker = ctrl.classList.contains("run-picker-btn");
      ctrl.click();
      // A run picker opens its dropdown: drop the cursor into it so the arrows
      // navigate its options (Escape closes it and returns here).
      if (opensPicker && ssOpenPicker()) {
        state.pickerCursor = 0;
        ssPaintPicker();
      }
      return;
    }
    if (ctrl.type === "checkbox") { ctrl.click(); return; }
    if (ctrl.tagName === "SELECT" || ctrl.type === "text" || ctrl.type === "search" || ctrl.type === "number") {
      state.navEditing = true;
      ctrl.focus();
    }
  }

  function initKeyboard() {
    window.ClipgenHotkeys.register([
      {
        id: "transport.playPause",
        handler: function () {
          if (state.videoPlaying) pauseVideo();
          else playVideo();
        },
      },
      // Arrows are the coarse seek; ,/. step a single frame (the page's
      // pre-registry arrows were frame-steps — that role moved to ,/.).
      // Arrow-key transport is gated to video focus so that a focused panel
      // (Shift+1..4) fully owns the arrows; ,/. fine-step is never an arrow so it
      // works regardless of focus.
      { id: "transport.seekBack", when: ssVideoFocused, handler: function () { _seekBy(-SEEK_STEP); } },
      { id: "transport.seekFwd", when: ssVideoFocused, handler: function () { _seekBy(SEEK_STEP); } },
      { id: "transport.stepBack", handler: function () { _seekBy(-FRAME_STEP); } },
      { id: "transport.stepFwd", handler: function () { _seekBy(FRAME_STEP); } },
      // Shift+arrow mirrors the ,/. fine step so the 1 s / 5 s pair is discoverable
      // from the arrow keys alone (screenspace-scoped to avoid Composer's Shift+arrow).
      { id: "screenspace.stepBackFine", when: ssVideoFocused, handler: function () { _seekBy(-FRAME_STEP); } },
      { id: "screenspace.stepFwdFine", when: ssVideoFocused, handler: function () { _seekBy(FRAME_STEP); } },
      { id: "screenspace.setIn", handler: function () { if (SS.setInMark) SS.setInMark(); } },
      { id: "screenspace.setOut", handler: function () { if (SS.setOutMark) SS.setOutMark(); } },
      {
        id: "screenspace.blink",
        repeat: false,
        when: function () { return _overlayEligibleForActiveTool(); },
        handler: function () {
          state.overlayBlinkActive = true;
          var curTs = Number(state.currentTimestamp || 0).toFixed(3);
          if (!state.overlayImage || state.overlayImageTimestamp !== curTs || state.overlayImageTool !== state.activeWorkflow) {
            refreshModelView();
          }
          renderOverlay();
        },
        onRelease: function () {
          if (state.overlayBlinkActive) {
            state.overlayBlinkActive = false;
            renderOverlay();
          }
        },
      },
      {
        id: "global.primary",
        when: function () {
          var btn = qs("#runBtn");
          return !!(btn && !btn.disabled);
        },
        handler: function () { qs("#runBtn").click(); },
      },
      { id: "screenspace.togglePanel", handler: function () { toggleBottomPanel(); } },
      {
        id: "screenspace.toggleInfoPanel",
        when: function () { return !!qs("#ssInfoPanel"); },
        handler: function () {
          var collapsed = qs("#ssInfoPanel").classList.contains("hidden");
          var btn = qs(collapsed ? "#ssInfoExpandBtn" : "#ssInfoCollapseBtn");
          if (btn) btn.click();
        },
      },
      { id: "screenspace.cycleToolPrev", handler: function () { cycleTool(-1); } },
      { id: "screenspace.cycleToolNext", handler: function () { cycleTool(1); } },
      {
        id: "screenspace.selectTool",
        repeat: false,
        handler: function (e, combo) { handleToolDigit(parseInt(combo, 10)); },
      },
      // Shift+1..4 target a panel for focus; the arrows then rove within it.
      {
        id: "screenspace.focusRegion",
        repeat: false,
        handler: function (e, combo) {
          ssFocusRegionByNumber(parseInt(combo.replace("Shift+", ""), 10));
        },
      },
      {
        id: "screenspace.nav",
        when: function () { return state.focusRegion !== "video"; },
        handler: function (e, combo) {
          if (combo === "ArrowUp") ssNavMove(-1);
          else if (combo === "ArrowDown") ssNavMove(1);
          else if (combo === "ArrowLeft") ssNavAdjust(-1);
          else if (combo === "ArrowRight") ssNavAdjust(1);
        },
      },
      {
        id: "screenspace.navActivate",
        repeat: false,
        when: function () { return state.focusRegion !== "video"; },
        handler: function () { ssNavActivate(); },
      },
    ]);

    // Back-out cascade: leave the notes editor, then return panel focus to the
    // video, then an open run-picker dropdown, then the active pointer
    // interaction, then the pending/active region, then the region-name modal.
    window.ClipgenHotkeys.registerEscape(function () {
      if (state.navEditing) {
        var active = document.activeElement;
        if (active && active.blur) active.blur();
        state.navEditing = false;
        ssPaintNav();
        return true;
      }
      var openCat = qs("#workflowCategories .ss-cat-chip.open");
      if (openCat) {
        closeCatMenus(null);
        return true;
      }
      var openPicker = qs(".run-picker-panel:not(.hidden)");
      if (openPicker) {
        closeRunPicker();
        state.pickerCursor = -1;
        if (state.focusRegion === "tool") ssPaintNav(); // restore the cursor on the picker button
        return true;
      }
      // With transient popovers closed, Escape returns panel focus to the video.
      if (state.focusRegion !== "video") {
        ssSetFocusRegion("video");
        return true;
      }
      var consumed = true;
      if (state.pipetteActive) {
        deactivatePipette();
      } else if (state.draggingRegion) {
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
      } else if (state.wandDragging) {
        // Abort an in-progress wand scrub. It never sets pendingRegion until
        // release, so nulling the state is a complete cancel; the satellite's
        // cached frame ImageData is only read while wandDragging is truthy.
        state.wandDragging = null;
        invalidateOverlayRect();
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        renderOverlay();
        updateRegionButtons();
      } else if (state.drawingRegion || state.drawingLasso) {
        state.drawingRegion = null;
        state.drawingLasso = null;
        invalidateOverlayRect();
        renderOverlay();
        updateRegionButtons();
      } else if (state.pendingRegion || state.activeRegion) {
        state.pendingRegion = null;
        state.activeRegion = null;
        renderOverlay();
        updateRegionButtons();
        updateRunButton();
      } else {
        consumed = false;
      }
      // Mirrors the pre-registry behavior: the region-name modal never
      // survives an Escape press, whatever else was cancelled.
      var modal = qs("#regionNameModal");
      if (modal && !modal.classList.contains("hidden")) {
        hideRegionNameModal();
        consumed = true;
      }
      // A stray tabbed focus (top-nav button, source <select>, …) is dropped by
      // the shared Escape fallback in hotkeys.js when nothing here claims it.
      return consumed;
    });
  }

  // ---- Panel divider ----

  function initBottomPanelDivider() {
    var panel = qs("#bottomPanel");
    if (!panel) return;
    var panelMaxH = Math.round(window.innerHeight * 0.6);
    initPanelDivider({
      isCollapsed: function () {
        return state.bottomCollapsed;
      },
      getHeight: function () {
        return state.panelHeight;
      },
      setHeight: function (h) {
        state.panelHeight = h;
        panel.style.height = h + "px";
      },
      getBounds: function () {
        return { min: 120, max: panelMaxH };
      },
      onDragStart: function () {
        document.body.classList.add("panel-dragging");
      },
      onDragEnd: function () {
        document.body.classList.remove("panel-dragging");
      },
      onToggle: toggleBottomPanel,
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

  // One-click boundary detection: enqueue a full-frame boundary task for every
  // participant that has a source video (state.participants is already filtered
  // to has_video). New tasks are pushed to state.tasks + rendered immediately,
  // matching the Run button's enqueue path; results stream in over SSE. The
  // in-flight guard blocks duplicate posts if the action is triggered again
  // before the sequential enqueue chain finishes.
  var _boundaryEnqueueInFlight = false;

  function detectBoundariesForAll() {
    if (_boundaryEnqueueInFlight) return;
    var participants = state.participants || [];
    if (!participants.length) return;
    _boundaryEnqueueInFlight = true;
    var chain = Promise.resolve();
    participants.forEach(function (p) {
      var pid = p.id;
      chain = chain.then(function () {
        return apiPost("api/tasks", { type: "boundary", participant: pid })
          .then(function (data) {
            if (data.ok && data.task) {
              if (!state.tasks.some(function (t) { return t.id === data.task.id; })) {
                state.tasks.push(data.task);
              }
              renderTaskList();
            }
          })
          .catch(function () { return null; });
      });
    });
    chain.then(function () {
      _boundaryEnqueueInFlight = false;
      showToast("Queued boundary detection for "
        + clipgenPluralUnit(participants.length, "participant", "participants"));
      startSSE();
    }).catch(function (err) {
      _boundaryEnqueueInFlight = false;
      showToast("Error: " + err.message);
    });
  }

  function detectBoundariesQuickAction() {
    var count = (state.participants || []).length;
    var busy = _boundaryEnqueueInFlight;
    return {
      icon: "film",
      label: "Detect boundaries",
      action: detectBoundariesForAll,
      disabled: count === 0 || busy,
      title: busy
        ? "Boundary detection is already being queued…"
        : count === 0
          ? "Load a participant with a source video first to detect boundaries."
          : "Detect scene boundaries for every participant with a source video (" + count + ").",
    };
  }

  function initTopNavActions() {
    if (!window.ClipgenTopNav) return;
    function rebuild() {
      window.ClipgenTopNav.setQuickActions([
        detectBoundariesQuickAction(),
        window.ClipgenExportActions.exportQuickAction(),
      ]);
    }
    rebuild();
    window.ClipgenExportActions.refreshExportStatus(rebuild);
    window.ClipgenTopNav.onBeforeOpen(function () {
      // Rebuild on every open so the boundary action reflects the current
      // participant list (loaded async after init). refreshExportStatus only
      // re-runs rebuild when the export flag flips, so it can't do this alone —
      // without this the boundary item stays frozen in its init-time (empty,
      // disabled) state.
      rebuild();
      window.ClipgenExportActions.refreshExportStatus(rebuild);
    });
  }

  // Command palette (command-palette.js): additions beyond the auto-ingested
  // quick actions — Run, plus per-participant jumps (the provider runs on
  // every palette open, so the list tracks state.participants).
  function initCommandPalette() {
    if (!window.ClipgenCommandPalette) return;
    window.ClipgenCommandPalette.setParticipants(function () {
      return (state.participants || []).map(function (p) { return p.id; });
    });
    window.ClipgenCommandPalette.register("screenspace", function () {
      function clickIfPresent(sel) {
        var btn = qs(sel);
        if (btn) btn.click();
      }
      var cmds = [
        {
          id: "screenspace:run",
          title: "Run analysis tool",
          icon: "play",
          keywords: "scan task queue start",
          section: "Screenspace",
          enabled: function () {
            var btn = qs("#runBtn");
            return !!btn && !btn.disabled;
          },
          run: function () { qs("#runBtn").click(); },
        },
        {
          id: "screenspace:clear-task-filter",
          title: "Clear task filter",
          icon: "x-mark",
          keywords: "reset show all queue completed failed",
          section: "Screenspace",
          enabled: function () { return !!state.taskFilter; },
          // Click the active filter button; its handler toggles the filter off.
          run: function () {
            clickIfPresent(state.taskFilter === "failed"
              ? "#taskFilterFailedBtn" : "#taskFilterDoneBtn");
          },
        },
        {
          id: "screenspace:toggle-info",
          title: "Toggle info panel",
          icon: "bars-3-bottom-left",
          keywords: "collapse expand help drawer",
          section: "Screenspace",
          visible: function () { return !!qs("#ssInfoPanel"); },
          run: function () {
            var collapsed = qs("#ssInfoPanel").classList.contains("hidden");
            clickIfPresent(collapsed ? "#ssInfoExpandBtn" : "#ssInfoCollapseBtn");
          },
        },
        {
          id: "screenspace:toggle-model-view",
          title: "Toggle model view",
          icon: "eye",
          keywords: "preview preprocess collapse expand panel",
          section: "Screenspace",
          visible: function () { return !!qs("#modelViewToggle"); },
          run: function () { qs("#modelViewToggle").click(); },
        },
        {
          id: "screenspace:toggle-calibration",
          title: "Toggle calibration panel",
          icon: "adjustments-horizontal",
          keywords: "ocr pins collapse expand panel",
          section: "Screenspace",
          visible: function () { return !!qs("#calibrationToggle"); },
          run: function () { qs("#calibrationToggle").click(); },
        },
        {
          id: "screenspace:toggle-bottom",
          title: "Toggle bottom panel",
          icon: "chevron-up-down",
          keywords: "collapse expand results timeline drawer",
          section: "Screenspace",
          visible: function () { return !!qs("#bottomPanel"); },
          run: function () { toggleBottomPanel(); },
        },
      ];
      // Switch-to-tool commands, one per analysis tool (grouped by category).
      // run() delegates to selectWorkflowType so it works in both nav modes.
      TOOL_CATEGORIES.forEach(function (cat) {
        cat.tools.forEach(function (type) {
          cmds.push({
            id: "screenspace:tool-" + type,
            title: "Switch to " + _toolLabel(type) + " tool",
            icon: TOOL_ICON_NAMES[type] || "cube",
            keywords: "tool detector select analysis " + cat.label.toLowerCase() + " " + type,
            section: "Tools",
            run: function () { selectWorkflowType(type); },
          });
        });
      });
      // "Jump to … in Screenspace" = stays here and selects in place; the
      // palette's built-in provider adds the cross-page "Open … in <Page>".
      (state.participants || []).forEach(function (p) {
        cmds.push({
          id: "screenspace:p:" + p.id,
          title: "Jump to " + p.id + " in Screenspace",
          icon: "user",
          keywords: "participant select video",
          section: "Participants",
          run: function () {
            var sel = qs("#participantSelect");
            sel.value = p.id;
            sel.dispatchEvent(new Event("change"));
          },
        });
      });
      return cmds;
    });
  }

  // ---- Settings (server-side STUDIO_SETTINGS) ----
  //
  // Backed by /api/settings. We mirror the Screenspace-relevant flags onto
  // state (state.restoreMarkersOnEdit for restoreTaskToWorkflow,
  // state.showConfidenceHistogram for the Results-panel histogram) so the hot
  // paths can read them without a network call. The settings modal's
  // onSave/onReset hooks call applyScreenspaceSettingsSnapshot to keep state in
  // sync after user edits.

  function applyScreenspaceSettingsSnapshot(applied, settings) {
    function pick(name) {
      if (applied && Object.prototype.hasOwnProperty.call(applied, name)) {
        return applied[name];
      }
      if (settings) {
        for (var i = 0; i < settings.length; i++) {
          if (settings[i].name === name) return settings[i].value;
        }
      }
      return undefined;
    }
    var markers = pick("SCREENSPACE_RESTORE_MARKERS_ON_EDIT");
    if (markers !== undefined) state.restoreMarkersOnEdit = !!markers;
    var hist = pick("SCREENSPACE_SHOW_CONFIDENCE_HISTOGRAM");
    if (hist !== undefined) state.showConfidenceHistogram = !!hist;
    var grouped = pick("SCREENSPACE_GROUPED_TOOL_NAV");
    if (grouped !== undefined) state.groupedToolNav = !!grouped;
    applyToolNavMode(); // switch tool nav mode (and re-sync the active chip)
  }

  function fetchScreenspaceSettings() {
    apiGet("/api/settings")
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
    initRegionDrag();
    initTimeline();
    initWorkflowTabs();
    applyToolNavMode(); // build/show the grouped nav immediately (default on)
    initModelView();
    initCalibration();
    initParamTooltips();
    initRunButton();
    initTaskQueue();
    initRightPaneTabs();
    initPauseButton();
    initTaskFilters();
    initResultsPanel();
    initBottomPanelDivider();
    initPreviewResize();
    initInfoNotes();
    initInfoPanelCollapse();
    initInfoSections();
    initKeyboard();
    initFrontendSwitcher();
    initTopNavActions();
    initCommandPalette();

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
              renderResults(); // reflect a histogram-toggle change immediately
            },
            onReset: function (scope, settings) {
              applyScreenspaceSettingsSnapshot(null, settings);
              renderResults();
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
        if (data.config) clipgenApplyConfig(data.config);
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
          // Fire-and-forget preload; failures just skip the cache warm.
          apiGetBlob(url)
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
          // Deep link (#P07, from the Overview Map) beats the stored pick.
          var hashPid = clipgenHashParticipant();
          if (hashPid) {
            for (var hpi = 0; hpi < state.participants.length; hpi++) {
              if (state.participants[hpi].id === hashPid) {
                pickId = hashPid;
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
      .catch(toastError("Failed to load regions"));

    apiGet("api/stashes")
      .then(function (data) {
        if (data.ok) {
          state.stashes = data.stashes || [];
          renderStashCards();
          renderRunRegionPicker();
        }
      })
      .catch(toastError("Failed to load stashes"));

    apiGet("api/tasks")
      .then(function (data) {
        if (data.ok) {
          state.tasks = data.tasks || [];
          renderTaskList();
          renderTimeline();
          // Fill the per-task result cache so the timeline has markers on load.
          // handleTaskData does this on every tick, but with all tasks already
          // completed the SSE stream below never starts, so seed it here too.
          if (SS.reconcileResultCache) SS.reconcileResultCache(state.tasks);
          if (SS.syncTaskResults) SS.syncTaskResults();
          if (state.tasks.some(function (t) { return t.status === "queued" || t.status === "running"; })) {
            startSSE();
          }
        }
      })
      .catch(toastError("Failed to load tasks"));
  });

  // ---- Satellite interface (window.ClipgenScreenspace) ----
  // Published for the screenspace-*.js satellite files (multitool params,
  // color, calibration) that load after this script. They read the hub's
  // shared state + helpers through this object and attach their own published
  // functions back onto it — mirrors window.ClipgenStudio. Assigned
  // synchronously here (during the hub script's load) so the object is fully
  // populated before any satellite IIFE runs; the DOMContentLoaded init above
  // and all user-event handlers fire later still, by which point satellites
  // have registered their functions.
  var SS = (window.ClipgenScreenspace = window.ClipgenScreenspace || {});
  SS.state = state;
  // Calibration entry points are thin hub delegators (see the calibration strip
  // section above) forwarding to SS.calRefresh / SS.calUpdateThresholdLine.
  SS.refreshCalibration = refreshCalibration;
  SS.updateCalibrationThresholdLine = updateCalibrationThresholdLine;
  // Hub helpers the satellites call outward.
  SS.gatherWorkflowParams = gatherWorkflowParams;
  SS.loadFrame = loadFrame;
  SS.seekPlayhead = seekPlayhead;
  SS.taskTypeColor = taskTypeColor;
  SS.taskRegionPixels = taskRegionPixels;
  SS.regionRefPayload = regionRefPayload;
  SS.normalizeRegionRef = normalizeRegionRef;
  SS.activeRegionRef = activeRegionRef;
  SS.availableRegionRefByKey = availableRegionRefByKey;
  SS.allAvailableRegionRefs = allAvailableRegionRefs;
  // Additional hub helpers the multitool-params satellite reuses.
  SS.regionRefKey = regionRefKey;
  SS.regionRefLabel = regionRefLabel;
  SS.buildTypeIcon = buildTypeIcon;
  // Grouped tool nav sync — the tasks satellite (restoreTaskToWorkflow) calls
  // this late-bound after it moves the active .wf-tab manually.
  SS.syncToolCategoryNav = syncToolCategoryNav;
  SS.iconSpan = iconSpan;
  SS.buildNormalizeControl = buildNormalizeControl;
  SS.buildColorModeControl = buildColorModeControl;
  SS._colorMode = _colorMode;
  SS.activatePipette = activatePipette;
  SS.deactivatePipette = deactivatePipette;
  SS.renderWorkflowParams = renderWorkflowParams;
  SS.updateRunButton = updateRunButton;
  // Hub helpers the tasks + results satellites reuse. findTask /
  // restoreTaskToWorkflow / setInputValue / syncValueDisplays live in
  // screenspace-tasks.js, and initResultsPanel / loadAndShowResults /
  // renderResults in screenspace-results.js — each published from there, with
  // thin hub delegators for the entry points the hub's own code calls.
  SS.applyColorMode = applyColorMode;
  SS.applyNormalizeMode = applyNormalizeMode;
  // renderOverlay itself lives in screenspace-overlay.js and publishes SS.renderOverlay;
  // these are the hub helpers that satellite reads (regionToPixels/taskTypeColor below).
  SS.regionColorForIndex = regionColorForIndex;
  SS.getThemeColors = getThemeColors;
  SS.templateOverlayBounds = templateOverlayBounds;
  SS.renderRunRegionPicker = renderRunRegionPicker;
  SS.selectParticipant = selectParticipant;
  // Hub helper the color satellite reuses.
  SS.regionToPixels = regionToPixels;
  // Hub helpers the overlay-interaction satellite reads (the region draw/drag/
  // resize state machine + toolbar live there; these stay hub-side). It owns
  // computeLabelRect / renderRegionChips / updateRegionButtons / invalidateOverlayRect
  // and publishes those back.
  SS.pauseVideo = pauseVideo;
  SS.stashRegions = stashRegions;
  SS.pinCurrentFrame = pinCurrentFrame;
  SS.togglePinTrayVisibility = togglePinTrayVisibility;
  SS.clearAllPins = clearAllPins;
  SS.updatePinButtons = updatePinButtons;

})();
