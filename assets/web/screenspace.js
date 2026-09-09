/* clipgen Screenspace page.
 *
 * Frame canvas + overlay editor for defining regions and templates on a video
 * frame, then running detector tasks (color/change/similarity/text/numbers/
 * timelapse/template/flow/scene/inactivity/boundary/attention, plus the
 * multitool chain).
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
    numbers: 1, template: 1, shape: 1, flow: 1, scene: 1, inactivity: 1,
    boundary: 1, attention: 1, timelapse: 1,
  };

  // Task-type icon span via mask-image; see .ss-task-icon in screenspace.css.
  function buildTypeIcon(type) {
    if (!SS_TASK_ICON_TYPES[type]) return null;
    var span = document.createElement("span");
    span.className = "ss-task-icon ss-task-icon--" + type;
    return span;
  }

  // Mask-image icon span; `name` is an assets/icons basename, `sizeClass` an .ss-icon
  // modifier.
  function iconSpan(name, sizeClass) {
    return iconMaskSpan(name, {
      className: "ss-icon" + (sizeClass ? " " + sizeClass : ""),
      basePath: "/screenspace/icons/",
    });
  }

  // OCR normalize direction; folds confusable glyphs before fuzzy compare (see
  // _normalize_ocr_text).
  var NORMALIZE_MODES = [
    { value: "letters", icon: "language", desc: "Fold digits to letters before matching (0→o, 1→l, 5→s). For word targets that OCR may read as digits" },
    { value: "off", icon: "no-symbol", desc: "No character folding" },
    { value: "digits", icon: "hashtag", desc: "Fold letters to digits before matching (O→0, l→1, S→5). For number targets that OCR may read as letters" },
  ];

  function _normalizeMode(mode) {
    return mode === "letters" || mode === "digits" ? mode : "off";
  }

  function buildNormalizeControl(id, mode, small) {
    return createSegTrack({
      id: id,
      value: _normalizeMode(mode),
      options: NORMALIZE_MODES,
      size: small ? "sm" : null,
      basePath: "/screenspace/icons/",
    });
  }

  // Reflect a saved mode onto an existing segmented control.
  function applyNormalizeMode(id, mode) {
    var hidden = qs("#" + id);
    if (!hidden || !hidden.parentNode) return;
    segTrackSetValue(hidden.parentNode, _normalizeMode(mode));
  }

  // Color match mode: region mean color vs per-pixel presence; see ColorTool in
  // screenspace_tools.py.
  var COLOR_MODES = [
    { value: "average", icon: "swatch", desc: "Match the region's average colour" },
    { value: "presence", icon: "magnifying-glass-circle", desc: "Match when the target colour appears anywhere in the region (per-pixel)" },
  ];

  function _colorMode(mode) {
    return mode === "presence" ? "presence" : "average";
  }

  // Reflect a saved color mode onto its control and min-area row.
  function applyColorMode(id, mode) {
    var hidden = qs("#" + id);
    if (!hidden) return;
    var m = _colorMode(mode);
    if (hidden.parentNode) segTrackSetValue(hidden.parentNode, m);
    var row = qs("#paramColorMinAreaRow");
    if (row) row.classList.toggle("hidden", m !== "presence");
  }

  // `onChange(mode)` fires before the bubbling input event; callers toggle the min-area
  // row.
  function buildColorModeControl(id, mode, small, onChange) {
    return createSegTrack({
      id: id,
      value: _colorMode(mode),
      options: COLOR_MODES,
      size: small ? "sm" : null,
      basePath: "/screenspace/icons/",
      onChange: onChange,
    });
  }

  // HSV hidden inputs cached by renderColorParams(); spares DOM queries on drag ticks.

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
    tools: {}, // per-tool facts from /api/tools (fast scan, confidence)
    // Whether a spreadsheet is loaded at all — gates the off-sheet label suffix.
    hasSheet: false,
    selectedParticipant: null,
    videoInfo: null,
    audioPanel: null, // ClipgenVideoControls audio-popover controller
    currentTimestamp: 0,
    frameImage: null,
    frameLoading: false,
    regions: {},
    activeRegion: null,
    drawingRegion: null,
    // Shaped-region draw state; lives on state because the Escape handler and overlay
    // painter read it.
    regionTool: "rect",
    drawingLasso: null,
    wandTolerance: 32,
    wandDragging: null,
    // Shape-draw session with an offscreen native-size mask canvas; cross-file readers
    // keep it on state.
    shapeDraw: null,
    shapeBrushSize: 24,
    pendingRegion: null,
    draggingRegion: null,
    resizingRegion: null,
    hoveredRegion: null,
    // Set by initRegionDrag to swallow the post-drop click; cleared by renderRegionChips
    // (cross-file).
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
    // Panel-focus keyboard nav state; semantics in the Panel focus navigation section
    // below.
    focusRegion: "video",
    focusCursor: 0,
    focusAnchor: null,
    navEditing: false,
    pickerCursor: -1,
    referenceTimestamp: null,
    sceneReferences: [],
    tasks: [],
    selectedTaskId: null,
    hoveredTaskId: null,
    selectedTaskResults: null,
    // Per-task results (taskId -> array); status ticks carry none. Kept current by
    // _syncTaskResults.
    taskResults: {},
    resultsLoading: false,
    resultsLazyObserver: null,
    // Hub<->tasks flags: resultsRequestVersion gates results fetches;
    // suppressCalibrationRefresh spans restoreTaskToWorkflow's rebuild.
    resultsRequestVersion: 0,
    heatmapOverlayRequestVersion: 0,
    // Heatmap thumb play state keyed "<taskId>|<attachment>"; renderResults() rebuilds the
    // strip, so toggles live here.
    heatmapPlaying: {},
    suppressCalibrationRefresh: false,
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
    // True while runRegions holds only the implicit active-chip seed; see
    // renderRunRegionPicker.
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
    capturedRefPreview: null,
    uploadedTemplate: null,
    uploadedTemplateImg: null,
    templateScalePreview: 1.0,
    templateOverlayPos: null,
    draggingTemplate: null,
    multitoolSteps: [],
    multitoolFocus: 0,
    hoveredResultSceneName: null,
    hoveredBoundaryTs: null,
    videoPlaying: false,
    videoMuted: false,
    videoPlaybackRate: 1,
    overlayEnabled: false,
    overlayLayer: null,
    overlayBlinkActive: false,
    overlayImage: null,
    overlayImageObjectUrl: null,
    overlayImageScope: null,
    overlayImageRegion: null,
    overlayImageTimestamp: null,
    overlayImageTool: null,
    overlayLayerSpec: {},
    rightPaneTab: "preview",
    resultsSwitcherOpen: false,
    amplitudeGraphEnabled: false,
    pins: [],
    maxPins: null,
    hoveredPinId: null,
    pinTrayHidden: false,
    calibrationResult: null,
    calibrationOcrWarmed: false,
    calibrationGreen: false,
    // Param panel by control id: paramValues survives tool switches, paramDefaults backs
    // reset buttons. See _snapshotParamValues.
    paramValues: {},
    paramDefaults: {},
  };

  var _playheadRaf = 0;
  var _preloadedFrames = {};
  // Per-participant source mtime_ns; the ?v= cache-bust suffix on frame and stream URLs.
  // Empty = unknown.
  var _videoVersions = {};

  // Frame-0 preload concurrency. Bounded: the server's 3-slot capture pool must not starve
  // the visible frame.
  var PRELOAD_CONCURRENCY = 2;
  var _preloadQueue = [];
  var _preloadActive = 0;
  var _preloadStopped = false;

  function queueFrameZeroPreload(participantIds) {
    participantIds.forEach(function (pid) {
      // Enqueue-time version: a late preload must not restore a blob selectParticipant
      // already dropped.
      _preloadQueue.push({ pid: pid, version: _videoVersions[pid] || "" });
    });
    _pumpFrameZeroPreload();
  }

  function _pumpFrameZeroPreload() {
    while (!_preloadStopped && _preloadActive < PRELOAD_CONCURRENCY && _preloadQueue.length) {
      var item = _preloadQueue.shift();
      if (_preloadedFrames[item.pid]) continue;
      _preloadActive++;
      _preloadFrameZero(item);
    }
  }

  function _preloadFrameZero(item) {
    apiGetBlob(frameUrl(item.pid, 0))
      .then(function (blob) {
        var url = URL.createObjectURL(blob);
        if (_preloadStopped || (_videoVersions[item.pid] || "") !== item.version) {
          try { URL.revokeObjectURL(url); } catch (_) {}
          return;
        }
        if (_preloadedFrames[item.pid]) {
          try { URL.revokeObjectURL(_preloadedFrames[item.pid]); } catch (_) {}
        }
        _preloadedFrames[item.pid] = url;
      })
      // Failed preloads just skip the warm; the trailing .then always releases the slot.
      .catch(function () {})
      .then(function () {
        _preloadActive--;
        _pumpFrameZeroPreload();
      });
  }

  window.addEventListener("pagehide", function () {
    // Mark in-flight preloads unowned so late blobs revoke themselves.
    _preloadStopped = true;
    _preloadQueue.length = 0;
    Object.keys(_preloadedFrames).forEach(function (pid) {
      try { URL.revokeObjectURL(_preloadedFrames[pid]); } catch (_) {}
      delete _preloadedFrames[pid];
    });
    // Model-view blob URLs (state + <img> expando) are revoked on replacement; release
    // here too.
    if (state.overlayImageObjectUrl) {
      try { URL.revokeObjectURL(state.overlayImageObjectUrl); } catch (_) {}
      state.overlayImageObjectUrl = null;
    }
    var mvImg = qs("#modelViewImage");
    if (mvImg && mvImg._modelViewObjectUrl) {
      try { URL.revokeObjectURL(mvImg._modelViewObjectUrl); } catch (_) {}
      mvImg._modelViewObjectUrl = null;
    }
  });
  var _participantRequestVersion = 0;
  var _frameRequestVersion = 0;

  // Region palette is screenspace-only (--region-color-1..N); common canvas colors use
  // getCanvasThemeColors() in utils.js.
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
  // Frames work at global time; <video> alone switches per part.
  function _ssParts() {
    var info = state.videoInfo;
    return info && info.parts && info.parts.length > 1 ? info.parts : null;
  }
  function _ssStreamUrlForPart(pid, i) {
    var url = videoStreamUrl(pid);
    return url + (url.indexOf("?") >= 0 ? "&" : "?") + "part=" + i;
  }

  // ---- Participants ----

  // Plain-text <option>s: off-sheet participants get a label suffix, only when a sheet is
  // loaded.
  function participantLabel(p) {
    return state.hasSheet && p.in_sheet === false ? p.id + " (off-sheet)" : p.id;
  }

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
      var opt = el("option", null, participantLabel(p));
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
      lbl.appendChild(document.createTextNode(participantLabel(p)));
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
    // Implicit seed follows the active chip until the picker is touched; else full frame.
    var seedRef = state.activeRegion && names.indexOf(state.activeRegion) >= 0
      ? activeRegionRef(state.activeRegion)
      : null;
    if (state.runRegions.length === 0) {
      state.runRegions = [seedRef || fullFrameRegionRef()];
      state.runRegionsSeeded = true;
    } else if (
      state.runRegionsSeeded &&
      seedRef &&
      state.runRegions.length === 1 &&
      (state.runRegions[0].source === "active" || state.runRegions[0].source === "full_frame") &&
      regionRefKey(state.runRegions[0]) !== regionRefKey(seedRef)
    ) {
      state.runRegions = [seedRef];
    }

    var btn = el("button", "run-picker-btn");
    btn.type = "button";
    updateRegionPickerBtnText(btn);

    var panel = el("div", "run-picker-panel hidden");

    // Full-frame entry first; its separator hairline only shows when rows follow.
    var hasFollowingRows = names.length > 0 || state.stashes.some(function (stash) {
      return Object.keys(stash.regions).length > 0;
    });
    var fullFrameRef = fullFrameRegionRef();
    var fullFrameLbl = document.createElement("label");
    fullFrameLbl.className = "run-picker-fullframe" + (hasFollowingRows ? " has-following" : "");
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
    // The seed mutates runRegions; the regions/stashes load paths never re-gate Run
    // themselves.
    updateRunButton();
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

  // Fast-scan support comes from the server's tool catalog (AnalysisTool ClassVars).
  function toolSupportsFastScan(type) {
    var tool = state.tools[type];
    return !!(tool && tool.supports_fast_scan);
  }

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
      "Integers only":    "Accept only whole-number readings: any extracted value carrying a decimal point or sign is rejected. For whole-number HUD targets.",
      "Integers":         "Accept only whole-number readings: any extracted value carrying a decimal point or sign is rejected. For whole-number HUD targets.",
      "Consecutive":      "Require this many consecutive sampled frames to match before an event fires (suppresses single-frame flicker; reports the run's median time)",
    },
    timelapse: {
      "Speed":            "Playback speed multiplier for the output",
      "Sample every":     "Seconds between captured frames (0 = every frame)",
      "Format":           "Output file format: video or animated GIF",
    },
    template: {
      "Template":         "Capture or upload the picture to search for. The selected region scopes where",
      "Threshold":        "How closely the picture must match. Lower it to allow looser matches",
    },
    shape: {
      "Shape":            "Capture, upload, or draw the shape to search for. Only its outline is matched; the selected region scopes where",
      "Threshold":        "How closely the outline must match. Lower it to allow looser matches",
      "Scale min":        "Smallest size to search at, as a percent of the reference",
      "Scale max":        "Largest size to search at, as a percent of the reference",
      "Scale steps":      "How many sizes to try between min and max. More steps = finer size coverage, slower scan",
      "Link axes":        "Uncheck to search width and height independently — for buttons that stretch with their content. Every width is tried at every height, so the scan slows accordingly",
      "Width scale min":  "Smallest width to search at, as a percent of the reference",
      "Width scale max":  "Largest width to search at, as a percent of the reference",
      "Width scale steps": "How many widths to try between min and max",
      "Height scale min": "Smallest height to search at, as a percent of the reference",
      "Height scale max": "Largest height to search at, as a percent of the reference",
      "Height scale steps": "How many heights to try between min and max",
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
      var tool = state.tools[state.activeWorkflow];
      var desc = tool && tool.fast_scan_description;
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
      // Segmented-track buttons carry data-desc; reuse the dark-pill tooltip.
      var seg = e.target.closest && e.target.closest(".cg-segtrack-btn");
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
        (e.target.closest(".param-label") || e.target.closest(".cg-segtrack-btn"))
      ) {
        tooltip.hide();
      }
    }, true);
  }

  // Clear the whole frame/overlay/playback pipeline so nothing leaks between participants;
  // version bumps drop in-flight loads.
  function selectParticipant(pid, initialTimestamp) {
    var participantRequestVersion = ++_participantRequestVersion;
    _frameRequestVersion += 1;
    state.heatmapOverlayRequestVersion += 1;
    state.selectedParticipant = pid;
    setStoredUIStateField("screenspace", "selectedParticipant", pid);
    // Participant-scoped result sync; completed tasks get no SSE tick to do it later.
    if (SS.syncTaskResults) SS.syncTaskResults();
    state.currentTimestamp = 0;
    state.videoInfo = null;
    state.videoActivePart = 0;
    state.videoOffset = 0;
    state.frameImage = null;
    state.frameLoading = false;
    state.referenceTimestamp = null;
    state.sceneReferences = [];
    cancelShapeDraw();
    // In/out markers are per participant; swap in the incoming pair (nulls if none).
    if (SS.restoreMarkers) SS.restoreMarkers(pid);
    updateMarkerInfo();
    state.resultOverlay = null;
    state.heatmapOverlay = null;
    state.pins = [];
    state.hoveredPinId = null;
    // Tray visibility is per participant (mirrors the reset when pinning).
    state.pinTrayHidden = false;
    // Drop stale calibration scores and in-flight /api/calibrate responses; the pin load
    // re-evaluates.
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
    // Skeleton frame + "Loading…" subheader until video info resolves or fails.
    qs("#videoInfo").classList.add("cg-shimmer");
    qs("#videoInfo").textContent = "Loading…";
    qs("#frameEmpty").classList.remove("hidden");
    setInfoParticipant(pid);

    // Fragmented-MP4 warning + remux action, at the top of the info panel.
    if (window.clipgenMediaBanner) {
      var banded = null;
      for (var bi = 0; bi < state.participants.length; bi++) {
        if (state.participants[bi].id === pid) { banded = state.participants[bi]; break; }
      }
      window.clipgenMediaBanner.show(qs(".ss-info-content"), banded);
    }

    apiGet("api/video/info/" + encodeURIComponent(pid))
      .then(function (data) {
        if (participantRequestVersion !== _participantRequestVersion || pid !== state.selectedParticipant) return;
        if (!data.ok) { qs("#videoInfo").classList.remove("cg-shimmer"); qs("#videoInfo").textContent = ""; return; }
        state.videoInfo = data.info;
        // Duration is only known now, so a restored marker can't be range-checked
        // until this point.
        if (SS.clampMarkersToDuration) SS.clampMarkersToDuration(data.info.duration);
        // Reconfigure the audio popover for this participant's track layout.
        if (state.audioPanel) state.audioPanel.refresh();
        // A changed mtime means a replaced source: drop the stale frame-0 blob.
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
        qs("#videoInfo").classList.remove("cg-shimmer");
        qs("#videoInfo").textContent = parts.join(" \u00b7 ");
        renderTimeline();
        updatePinButtons();
        // Preload video source for instant playback
        qs("#videoPlayer").src = videoStreamUrl(pid);
        loadFrame(initialTimestamp !== undefined ? initialTimestamp : 0);
      })
      .catch(function () {
        // Clear "Loading…" for the still-current participant after a failed fetch.
        if (participantRequestVersion === _participantRequestVersion && pid === state.selectedParticipant) {
          qs("#videoInfo").classList.remove("cg-shimmer");
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
  // Pins mark frames with polarity; positive must fire, negative must not.

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
        var newCollapsed = !getStoredUIMapEntry("screenspace", "infoSectionsCollapsed", n, false);
        setStoredUIMapEntry("screenspace", "infoSectionsCollapsed", n, newCollapsed);
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

  // One frame in flight; later scrubs park in `_pendingFrameTs`; version checks drop stale
  // paints.
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
      // A live wand scrub traces the outgoing frame's pixels; cancel it on every reload
      // path.
      cancelWandDrag();
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
        // A resize invalidates the shape-draw mask's pixel space.
        cancelShapeDraw();
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

    // Hover mute for a 0–200% volume popover; per-track mixing works for single-file
    // participants only.
    state.audioPanel = window.ClipgenVideoControls.attachAudioPanel({
      video: video,
      button: muteBtn,
      getTracks: function () {
        return (state.videoInfo && state.videoInfo.audio_tracks) || [];
      },
      trackAudioUrl: function (idx) {
        var pid = state.selectedParticipant;
        if (!pid || !state.videoInfo) return null;
        // Per-track mixing is single-file only; multi-part keeps the master slider.
        if (state.videoInfo.parts && state.videoInfo.parts.length > 1) return null;
        var url = "api/video/audio-track/" + encodeURIComponent(pid) + "/" + idx;
        var v = _videoVersions[pid];
        return v ? url + "?v=" + encodeURIComponent(v) : url;
      },
    });

    muteBtn.addEventListener("click", function () {
      state.videoMuted = !state.videoMuted;
      if (state.audioPanel) state.audioPanel.setMuted(state.videoMuted);
      else video.muted = state.videoMuted;
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
            window.ClipgenVideoControls.safePlay(video);
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
    setStoredUIMapEntry("screenspace", "videoTimeByParticipant", state.selectedParticipant, t);
  }

  function playVideo() {
    var video = qs("#videoPlayer");
    if (!state.selectedParticipant || !state.videoInfo) return;

    var parts = _ssParts();
    if (parts) {
      // Multi-video: play the part that owns the global playhead, seeking local.
      var i = clipgenPartForGlobal(parts, state.currentTimestamp);
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
    // Route through the audio panel so multitrack mode keeps the <video> muted.
    if (state.audioPanel) state.audioPanel.setMuted(state.videoMuted);
    else video.muted = state.videoMuted;

    video.classList.add("active");
    qs("#frameCanvas").classList.add("video-active");

    state.videoPlaying = true;
    updateVideoButtons();

    applyPlaybackRate();
    // Rejection means playback never started; the button state must fall back.
    window.ClipgenVideoControls.safePlay(video, pauseVideo);
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

  // ---- Region drawing (impl in screenspace-overlay-interaction.js) ----
  // Thin delegators for the hub's own call sites.
  function initRegionDrawing() { return SS.initRegionDrawing && SS.initRegionDrawing.apply(null, arguments); }
  function renderRegionChips() { return SS.renderRegionChips && SS.renderRegionChips.apply(null, arguments); }
  function updateRegionButtons() { return SS.updateRegionButtons && SS.updateRegionButtons.apply(null, arguments); }
  function hideRegionNameModal() { return SS.hideRegionNameModal && SS.hideRegionNameModal.apply(null, arguments); }
  function invalidateOverlayRect() { return SS.invalidateOverlayRect && SS.invalidateOverlayRect.apply(null, arguments); }
  function cancelWandDrag() { return SS.cancelWandDrag && SS.cancelWandDrag.apply(null, arguments); }
  function toggleShapeDraw() { return SS.toggleShapeDraw && SS.toggleShapeDraw.apply(null, arguments); }
  function cancelShapeDraw() { return SS.cancelShapeDraw && SS.cancelShapeDraw.apply(null, arguments); }
  function openSampleModal() { return SS.openSampleModal && SS.openSampleModal.apply(null, arguments); }

  // ---- Region stashing + chip drag (impl in screenspace-regions.js) ----
  function initRegionDrag() { return SS.initRegionDrag && SS.initRegionDrag.apply(null, arguments); }
  function renderStashCards() { return SS.renderStashCards && SS.renderStashCards.apply(null, arguments); }
  function stashRegions() { return SS.stashRegions && SS.stashRegions.apply(null, arguments); }


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

  // Delegator; renderOverlay lives in screenspace-overlay.js.
  function renderOverlay() {
    return SS.renderOverlay && SS.renderOverlay.apply(null, arguments);
  }

  // ---- Timeline — screenspace-timeline.js ----
  // Delegators; the tasks and results satellites destructure SS.renderTimeline at load.
  function initTimeline() { return SS.initTimeline && SS.initTimeline.apply(null, arguments); }
  function renderTimeline() { return SS.renderTimeline && SS.renderTimeline.apply(null, arguments); }
  function renderPlayhead() { return SS.renderPlayhead && SS.renderPlayhead.apply(null, arguments); }
  function updateMarkerInfo() { return SS.updateMarkerInfo && SS.updateMarkerInfo.apply(null, arguments); }
  // ---- Tool info tooltip ----

  var TOOL_INFO = {
    multitool: "Combines several tools so a frame only matches when it passes every step. For example, a red health bar AND the word 'DEAD'. Add at least two steps; a step can also be set to exclude (match only when it does NOT apply). Get each tool working on its own first, then chain them here to pin down precise moments.",
    color: "Finds frames where the average color of your region matches a color you pick. Draw a small region over a solid-colored element and sample its color; widen Tolerance to catch more shades, tighten it to be stricter. Good for color-coded elements like a health bar or status light. To find a specific icon or picture instead, use Template.",
    change: "Flags frames where the picture inside your region differs from a moment earlier: sudden changes such as a screen transition, a pop-up appearing, or a loading screen finishing. Raise the Threshold if it fires on every small flicker. Unlike Flow (which measures movement) it reacts to any difference; unlike Boundary it watches only the region you draw, not the whole screen.",
    similarity: "Capture one reference frame, then this finds every later frame that looks almost identical: a strict, pixel-for-pixel match that's sensitive to lighting and layout shifts. Lower the Threshold to allow looser matches. Use it to catch when one exact state returns (a specific dialog or menu). For 'which screen are we on' across several screens that vary, use Scene instead.",
    text: "Reads on-screen text in your region (OCR) and flags frames matching your search words, allowing for small misreads. Draw a tight region around the text; raise the OCR confidence if you get false hits. Good for catching specific labels, error messages, or button text. To compare on-screen numbers (e.g. score over 1000), use Numbers.",
    numbers: "Reads a number from your region (OCR) and flags frames where it meets a rule you set: equals, greater than, less than, or within a range. Draw a tight region around just the number, then pick the operator and target value. Great for scores, timers, lives, or any changing count. For words rather than numbers, use Text.",
    timelapse: "Produces one sped-up video or GIF of your region over the time range you choose: a fast way to skim a long session. Unlike every other tool it doesn't mark individual moments on the timeline; it outputs a single clip. Set the speed, and optionally sample every N seconds for a shorter file.",
    template: "Capture or upload a small reference image, then this looks for that exact picture within the selected region — pick Full frame to search the whole screen. Ideal for finding an icon, button, or logo wherever it appears. Lower the Threshold to allow looser matches. Unlike Color (which matches an average shade) it matches the picture itself; unlike Similarity it can find the picture anywhere, not just where it was sampled.",
    shape: "Capture or upload a reference image, then this looks for its outline within the selected region, sweeping a range of sizes — pick Full frame to search the whole screen. Because only edges are matched, it finds the shape even when its colors change (dark mode, hover states, hollow vs filled) or it appears larger or smaller than the reference. Text inside the reference counts as part of the outline, so a button sampled with one label scores lower against the same button with another; sample the chrome without the label when labels vary. Use Template when the exact pixels matter. Thin or tiny outlines are hard to match; prefer references at least ~20 px across.",
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
      // Alt-hold hint: tabs 1–9 map to digit combos; tabs 10+ have none.
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
        // Every selection path funnels through here; keep the category nav in sync.
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

  // Z/X tool cycling in on-screen order; delegates to the tab click path.
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

  // ---- Grouped tool nav ----
  // Layer over hidden .wf-tab row; selecting delegates to tab clicks.
  var TOOL_CATEGORIES = [
    { label: "Multitool", tools: ["multitool"], alwaysIcon: true, standalone: true },
    { label: "Difference", tools: ["change", "similarity", "inactivity"], icon: "square-2-stack" },
    { label: "Detection", tools: ["template", "shape", "color", "text", "numbers"], icon: "magnifying-glass" },
    { label: "Classification", tools: ["scene", "boundary"], icon: "tag" },
    { label: "Attention", tools: ["flow", "attention"], icon: "cursor-arrow-rays" },
    { label: "Utility", tools: ["timelapse"], icon: "cog-6-tooth" },
  ];

  // Heroicon basenames per tool for the command palette; mirrors .ss-task-icon--<type> in
  // screenspace.css.
  var TOOL_ICON_NAMES = {
    multitool: "wrench-screwdriver", color: "eye-dropper", change: "bolt",
    similarity: "photo", text: "language", numbers: "hashtag",
    template: "viewfinder-circle", shape: "star", flow: "arrows-right-left",
    scene: "squares-2x2",
    inactivity: "pause-circle", boundary: "flag", timelapse: "forward",
    attention: "eye",
  };

  var _catNavBuilt = false;
  var _catOutsideBound = false;

  // Category glyph; mask set inline since no .ss-task-icon--<type> class exists.
  function buildCatIcon(name) {
    if (!name) return null;
    return iconMaskSpan(name, {
      className: "ss-task-icon",
      basePath: "/screenspace/icons/",
    });
  }

  // Alt-hold chip hints; while a dropdown is open its items carry the digits (see
  // handleToolDigit).
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

  // Delegate to the hidden flat tab so the whole selection path runs.
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
      // Menu items must not nest inside a <button> (invalid HTML); mirrors
      // #exportEventsWrap.
      var chip = el("div", "ss-cat-chip");
      chip.setAttribute("data-cat", cat.label);
      chip.setAttribute("data-tools", cat.tools.join(","));
      if (cat.icon) chip.setAttribute("data-cat-icon", cat.icon);
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
        // Names the category; the chip shows the active tool instead when selected.
        menu.setAttribute("aria-label", cat.label + " tools");
        var head = el("div", "ss-cat-menu-head");
        head.setAttribute("role", "presentation");
        var headIcon = buildCatIcon(cat.icon);
        if (headIcon) head.appendChild(headIcon);
        head.appendChild(el("span", "", cat.label + " tools"));
        menu.appendChild(head);
        cat.tools.forEach(function (type, ti) {
          var item = el("button", "ss-cat-item");
          item.type = "button";
          item.setAttribute("data-type", type);
          item.setAttribute("role", "menuitem");
          // Alt-hold hint: digit ti+1 selects this item while the menu is open.
          if (ti < 9) {
            item.setAttribute("data-hotkey", "screenspace.selectTool");
            item.setAttribute("data-hotkey-combo", String(ti));
          }
          var icon = buildTypeIcon(type);
          if (icon) item.appendChild(icon);
          item.appendChild(el("span", "ss-cat-item-label", toolLabel(type)));
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

  // Active chip: solid tool-color fill, tool icon + name only. Resting: category icon +
  // name.
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
      // Glyph: active tool's icon, else the tool's on alwaysIcon chips, else the
      // category's.
      var glyphType = isActive ? active : (alwaysIcon ? tools[0] : null);
      if (glyph) {
        glyph.innerHTML = "";
        var icon = glyphType
          ? buildTypeIcon(glyphType)
          : buildCatIcon(chip.getAttribute("data-cat-icon"));
        if (icon) glyph.appendChild(icon);
      }
      if (isActive) {
        chip.setAttribute("data-active-type", active);
        if (text) text.textContent = toolLabel(active);
      } else {
        chip.removeAttribute("data-active-type");
        if (text) text.textContent = cat;
      }
      chip.querySelectorAll(".ss-cat-item").forEach(function (item) {
        item.classList.toggle("active", isActive && item.getAttribute("data-type") === active);
      });
    });
  }

  // Digits 1–9: flat mode picks tab N; grouped picks segment N, then tool N inside.
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

  // Toggle grouped nav vs flat tabs per state.groupedToolNav; builds the nav lazily.
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

  // ---- Single-tool parameter panels (impl in screenspace-params.js) ----
  function initParamResets() { return SS.initParamResets && SS.initParamResets.apply(null, arguments); }
  function refTimeChip() { return SS.refTimeChip && SS.refTimeChip.apply(null, arguments); }
  function renderWorkflowParams() { return SS.renderWorkflowParams && SS.renderWorkflowParams.apply(null, arguments); }
  function updateParamResetButtons() { return SS.updateParamResetButtons && SS.updateParamResetButtons.apply(null, arguments); }

  // ---- Model view (impl in screenspace-model-view.js) ----
  // Delegators; the satellite also publishes _overlayEligibleForActiveTool,
  // _updateMinAreaReadout, _previewRegionRef.
  function initModelView() { return SS.initModelView && SS.initModelView.apply(null, arguments); }
  function refreshModelView(opts) { return SS.refreshModelView && SS.refreshModelView.apply(null, arguments); }
  function _updateOverlayUi() { return SS._updateOverlayUi && SS._updateOverlayUi.apply(null, arguments); }
  function _overlayEligibleForActiveTool() { return SS._overlayEligibleForActiveTool && SS._overlayEligibleForActiveTool.apply(null, arguments); }
  function _updateMinAreaReadout(sfx) { return SS._updateMinAreaReadout && SS._updateMinAreaReadout.apply(null, arguments); }

  // ---- Calibration strip (impl in screenspace-calibration.js) ----
  // Delegators; state.suppressCalibrationRefresh is set by restoreTaskToWorkflow, checked
  // here.
  function refreshCalibration(opts) { return SS.calRefresh && SS.calRefresh(opts); }
  function updateCalibrationThresholdLine() { return SS.calUpdateThresholdLine && SS.calUpdateThresholdLine(); }
  function renderCalibration() { return SS.calRender && SS.calRender(); }
  function updateCalibrationVisibility() { return SS.calVisibility && SS.calVisibility(); }
  function initCalibration() { return SS.calInit && SS.calInit(); }

  // ---- Color picker (impl in screenspace-color.js) ----
  // Delegators; sampleColorFromRegion's handler reference stays unchanged.
  function updateColorPreview() { return SS.updateColorPreview && SS.updateColorPreview(); }
  function setTargetColor(h, s, v) { return SS.setTargetColor && SS.setTargetColor(h, s, v); }
  function renderColorPalette() { return SS.renderColorPalette && SS.renderColorPalette(); }
  function renderBrightnessStrip() { return SS.renderBrightnessStrip && SS.renderBrightnessStrip(); }
  function sampleColorFromRegion() { return SS.sampleColorFromRegion && SS.sampleColorFromRegion(); }
  function updateColorSampleBtnLabel() { return SS.updateColorSampleBtnLabel && SS.updateColorSampleBtnLabel(); }

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
    // An uploaded template scans the full frame, so it needs no region.
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
      var isTemplate = state.activeWorkflow === "template"
        || state.activeWorkflow === "shape";
      var isFullFrameTool = state.activeWorkflow === "boundary"
        || state.activeWorkflow === "attention";
      var hasUploadedTemplate = !!state.uploadedTemplate;
      // Region or uploaded image must supply the reference patch; template ignores the run
      // region.
      var templateMissingPatch = isTemplate && !hasRegion && !hasUploadedTemplate;
      // Boundary and Attention are full-frame only — they need no region at all.
      var nonTemplateMissingRegion = !isTemplate && !isFullFrameTool && !hasRegion;
      btn.disabled = nonTemplateMissingRegion || templateMissingPatch || !hasParticipants;
      if (templateMissingPatch) {
        btn.setAttribute("data-tooltip", "Upload a reference image or pick a region first");
      } else if (nonTemplateMissingRegion) {
        btn.setAttribute("data-tooltip", "Select a region first");
      } else if (!hasParticipants) {
        btn.setAttribute("data-tooltip", "Select participants to run");
      } else {
        btn.removeAttribute("data-tooltip");
      }
    }
    // Calibration agreement is a hover hint, never a block; calibrationGreen implies pins
    // exist.
    if (!btn.disabled && state.calibrationGreen) {
      btn.setAttribute("data-tooltip", "Calibrated: pins satisfied");
    }
  }

  // ---- Run analysis (impl in screenspace-run.js) ----
  // Thin delegators for the hub's own call sites.
  function initRunButton() { return SS.initRunButton && SS.initRunButton.apply(null, arguments); }
  function gatherWorkflowParams() { return SS.gatherWorkflowParams && SS.gatherWorkflowParams.apply(null, arguments); }

  // ---- Task queue (impl in screenspace-tasks.js) ----
  // Thin delegators for the hub's own call sites.
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
  // Thin delegators for the hub's own call sites.
  function initResultsPanel() { return SS.initResultsPanel && SS.initResultsPanel(); }
  function renderResults() { return SS.renderResults && SS.renderResults(); }

  // ---- Keyboard shortcuts (shared hotkeys.js registry) ----

  function _seekBy(delta) {
    if (!state.videoInfo) return;
    loadFrame(clamp(state.currentTimestamp + delta, 0, Math.max(0, state.videoInfo.duration - 0.001)));
  }

  // ---- Panel focus navigation ----
  // Painted cursor (.ss-nav-cursor), not DOM focus, which would trip isTypingTarget.

  // True while a panel owns the arrows; a hidden, empty region hands them back.
  function ssNavFocused() {
    if (state.focusRegion === "video") return false;
    if (ssNavItems(state.focusRegion).length) return true;
    ssSetFocusRegion("video");
    return false;
  }

  function ssVideoFocused() {
    return !ssNavFocused();
  }

  function ssElVisible(elm) {
    if (!elm) return false;
    var r = elm.getBoundingClientRect();
    return r.width > 0 || r.height > 0;
  }

  // A tool item's control: the .param-control child, or the top-row control itself.
  function ssToolControl(item) {
    if (item && item.classList && item.classList.contains("param-row")) {
      return item.querySelector("input, select, textarea, button");
    }
    return item;
  }

  // Ordered visible items per region; the visibility filter keeps hidden rows from
  // claiming the arrows.
  function ssNavItems(region) {
    // slice: the task/results branches return a NodeList, which has no .filter.
    return Array.prototype.slice.call(_ssNavItemsRaw(region)).filter(ssElVisible);
  }

  function _ssNavItemsRaw(region) {
    if (region === "sidebar") {
      var items = [];
      var notes = qs("#ssInfoNotes");
      if (notes) items.push(notes);
      // Each section contributes its header (Enter toggles collapse), then its rows while
      // expanded.
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
      // Top-row run controls, then the tool's param rows; numerals and Z/X cover the
      // selector.
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
  // pickerCursor >= 0 means inside a dropdown; tool-region handlers delegate here.

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

  // Repaint the cursor; regions re-render innerHTML, so the index is re-clamped every
  // time.
  function ssPaintNav() {
    ssClearNavPaint();
    if (state.focusRegion === "video") return;
    var items = ssNavItems(state.focusRegion);
    if (!items.length) { state.focusCursor = 0; return; }
    state.focusCursor = clamp(state.focusCursor, 0, items.length - 1);
    var cur = items[state.focusCursor];
    if (cur) {
      cur.classList.add("ss-nav-cursor");
      state.focusAnchor = ssNavKey(cur);
      if (cur.scrollIntoView) cur.scrollIntoView({ block: "nearest" });
    }
  }

  // Stable nav item identity; the task queue re-sorts on every poll, so indices go stale.
  function ssNavKey(item) {
    if (!item) return null;
    if (item.dataset && item.dataset.taskId) return "task:" + item.dataset.taskId;
    if (item.id) return "id:" + item.id;
    return null;
  }

  // Re-anchor and repaint after a wholesale list rebuild wipes the painted cursor.
  function ssRefreshNav() {
    if (state.focusRegion === "video") return;
    var items = ssNavItems(state.focusRegion);
    // Nothing to land on; ssNavFocused hands the arrows back on the next keypress.
    if (!items.length) return;
    if (state.focusAnchor) {
      for (var i = 0; i < items.length; i++) {
        if (ssNavKey(items[i]) === state.focusAnchor) { state.focusCursor = i; break; }
      }
    }
    ssPaintNav();
  }

  function ssSetFocusRegion(region) {
    closeRunPicker(); // a transient dropdown doesn't survive a focus-region change
    state.focusRegion = region;
    state.navEditing = false;
    state.pickerCursor = -1;
    state.focusAnchor = null;
    if (region === "video") { ssClearNavPaint(); return; }
    state.focusCursor = 0;
    ssPaintNav();
  }

  // Clicks re-own the arrows: a focused-region item moves the cursor; other clicks refocus
  // the video.
  function ssSyncFocusToClick(e) {
    if (state.focusRegion === "video" || state.navEditing) return;
    if (ssInPicker()) return; // an open dropdown owns its own click handling
    var items = ssNavItems(state.focusRegion);
    for (var i = 0; i < items.length; i++) {
      if (items[i] === e.target || items[i].contains(e.target)) {
        state.focusCursor = i;
        ssPaintNav();
        return;
      }
    }
    ssSetFocusRegion("video");
  }

  // Shift+N: reveal the panel, then land the cursor; declines when empty (mirrors Studio
  // kbJumpTo).
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
    // Drop lingering native focus so only one focus indicator shows.
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

  // Step controls by setting .value and firing input; real focus would double-apply
  // browser arrow stepping.
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

  // Only the tool region has horizontal controls; elsewhere consume the key so arrows
  // never seek.
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
    // Enter activates buttons/checkboxes; text, number and select controls take real focus
    // (Escape returns).
    var ctrl = ssToolControl(cur);
    if (!ctrl) return;
    if (ctrl.tagName === "BUTTON") {
      var opensPicker = ctrl.classList.contains("run-picker-btn");
      ctrl.click();
      // A run picker opens its dropdown; drop the cursor in (Escape returns here).
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

  // The wand snapshots pixels at press, so frame-changing keys must wait out a pointer
  // drag.
  function noPointerDrag() {
    return !state.wandDragging && !state.drawingLasso && !state.drawingRegion;
  }

  // Non-zero while B is held; the gap on release tells a peek from a tap.
  var _blinkStart = 0;

  function initKeyboard() {
    window.ClipgenHotkeys.register([
      {
        id: "transport.playPause",
        when: noPointerDrag,
        handler: function () {
          if (state.videoPlaying) pauseVideo();
          else playVideo();
        },
      },
      // Arrows coarse-seek only with video focus so a focused panel owns them; ,/.
      // fine-step regardless.
      { id: "transport.seekBack", when: ssVideoFocused, handler: function () { _seekBy(-SEEK_STEP); } },
      { id: "transport.seekFwd", when: ssVideoFocused, handler: function () { _seekBy(SEEK_STEP); } },
      { id: "transport.stepBack", when: noPointerDrag, handler: function () { _seekBy(-FRAME_STEP); } },
      { id: "transport.stepFwd", when: noPointerDrag, handler: function () { _seekBy(FRAME_STEP); } },
      // Shift+arrow mirrors the ,/. fine step; screenspace-scoped to avoid Composer's
      // binding.
      { id: "screenspace.stepBackFine", when: ssVideoFocused, handler: function () { _seekBy(-FRAME_STEP); } },
      { id: "screenspace.stepFwdFine", when: ssVideoFocused, handler: function () { _seekBy(FRAME_STEP); } },
      { id: "screenspace.setIn", handler: function () { if (SS.setInMark) SS.setInMark(); } },
      { id: "screenspace.setOut", handler: function () { if (SS.setOutMark) SS.setOutMark(); } },
      // Hold to peek, tap to latch (like Composer's B); tapping the checkbox persists and
      // repaints.
      {
        id: "screenspace.blink",
        repeat: false,
        when: function () { return _overlayEligibleForActiveTool(); },
        handler: function () {
          if (_blinkStart) return; // blur can swallow a keyup; don't restack
          _blinkStart = Date.now();
          state.overlayBlinkActive = true;
          var curTs = Number(state.currentTimestamp || 0).toFixed(3);
          if (!state.overlayImage || state.overlayImageTimestamp !== curTs || state.overlayImageTool !== SS._previewToolKey()) {
            refreshModelView();
          }
          renderOverlay();
        },
        onRelease: function () {
          if (!_blinkStart) return;
          var tapped = Date.now() - _blinkStart < 250;
          _blinkStart = 0;
          state.overlayBlinkActive = false;
          renderOverlay();
          var toggle = qs("#modelViewOverlayToggle");
          if (tapped && toggle && !toggle.disabled) toggle.click();
        },
      },
      {
        id: "screenspace.cycleOverlayLayer",
        when: function () { return _overlayEligibleForActiveTool(); },
        handler: function () { SS.cycleOverlayLayer(); },
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
        when: ssNavFocused,
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
        when: ssNavFocused,
        handler: function () { ssNavActivate(); },
      },
    ]);

    // Capture phase: read the click against the current DOM before page handlers
    // re-render.
    document.addEventListener("mousedown", ssSyncFocusToClick, true);

    // Back-out cascade: notes editor, panel focus, picker dropdown, pointer interaction,
    // region, name modal.
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
        // Only the satellite's cancel path frees the cached full-frame ImageData (~33 MB
        // at 4K).
        cancelWandDrag();
      } else if (state.shapeDraw) {
        // Exit shape-draw mode before touching pending/active regions.
        cancelShapeDraw();
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
      // The region-name modal never survives Escape, whatever else was cancelled.
      var modal = qs("#regionNameModal");
      if (modal && !modal.classList.contains("hidden")) {
        hideRegionNameModal();
        consumed = true;
      }
      // Stray tabbed focus falls to hotkeys.js's shared Escape fallback when nothing
      // claims it.
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
    var container = qs("#frameContainer");
    if (!container) return;
    var MIN_PCT = 30;
    var MAX_PCT = 100;
    var startWidthPx = 0;
    var parentWidth = 0;
    initDragHandle(qs("#previewResizeHandle"), "x", {
      cursor: "nwse-resize",
      stopPropagation: true,
      onStart: function () {
        startWidthPx = container.getBoundingClientRect().width;
        parentWidth = container.parentElement.getBoundingClientRect().width;
        return true;
      },
      onDelta: function (delta) {
        var pct = ((startWidthPx + delta) / parentWidth) * 100;
        state.previewMaxWidth = Math.round(Math.max(MIN_PCT, Math.min(MAX_PCT, pct)));
        container.style.maxWidth = state.previewMaxWidth + "%";
      },
      onToggle: function () {
        state.previewMaxWidth = MAX_PCT;
        container.style.maxWidth = "";
      },
    });
  }

  // ---- Init ----

  // One-click boundary detection for every participant with video; the guard blocks
  // duplicate posts mid-chain.
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

  // Rebuilt on every open so the boundary item tracks participants.
  function initTopNavActions() {
    if (!window.ClipgenTopNav) return;
    window.ClipgenTopNav.installQuickActions(function () {
      return [
        detectBoundariesQuickAction(),
        window.ClipgenExportActions.exportQuickAction(),
      ];
    }, { rebuildOnOpen: true });
  }

  // Command palette additions: Run plus per-participant jumps; the provider re-runs on
  // every open.
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
      var palette = window.ClipgenCommandPalette;
      var cmds = [
        palette.buttonCommand("Screenspace", "screenspace:run", "Run analysis tool", "play",
          "scan task queue start", "runBtn"),
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
          id: "screenspace:show-preview",
          title: "Show preview panel",
          icon: "eye",
          keywords: "model view preprocess calibration pins ocr tab",
          section: "Screenspace",
          visible: function () { return !!qs('.rp-tab[data-tab="preview"]'); },
          run: function () { setRightPaneTab("preview"); },
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
      // One switch-to-tool command per tool; selectWorkflowType works in both nav modes.
      TOOL_CATEGORIES.forEach(function (cat) {
        cat.tools.forEach(function (type) {
          cmds.push({
            id: "screenspace:tool-" + type,
            title: "Switch to " + toolLabel(type) + " tool",
            icon: TOOL_ICON_NAMES[type] || "cube",
            keywords: "tool detector select analysis " + cat.label.toLowerCase() + " " + type,
            section: "Tools",
            run: function () { selectWorkflowType(type); },
          });
        });
      });
      return cmds.concat(palette.participantJumps("screenspace:p:", "Screenspace",
        "participant select video", (state.participants || []).map(function (p) { return p.id; }),
        function (pid) {
          var sel = qs("#participantSelect");
          sel.value = pid;
          sel.dispatchEvent(new Event("change"));
        }));
    });
  }

  // ---- Settings (server-side STUDIO_SETTINGS) ----
  // Flags mirror onto state; the modal's onSave/onReset call
  // applyScreenspaceSettingsSnapshot.

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
    initParamResets();
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
    if (window.wireSettingsButton) {
      window.wireSettingsButton({
        initialTab: "Screenspace",
        onApply: function (applied, settings) {
          applyScreenspaceSettingsSnapshot(applied, settings);
          renderResults(); // reflect a histogram-toggle change immediately
        },
      });
    }

    // Participant select
    qs("#participantSelect").addEventListener("change", function () {
      var pid = this.value;
      if (pid) {
        var ts = getStoredUIMapEntry("screenspace", "videoTimeByParticipant", pid);
        if (typeof ts !== "number") ts = undefined;
        selectParticipant(pid, ts);
        state.runParticipants = [pid];
        renderRunParticipantPicker();
      }
    });

    // Close run picker on outside click
    document.addEventListener("click", function (e) {
      if (!e.target.closest(".run-picker-wrap")) closeRunPicker();
    });

    // Load initial data; the tool catalog first so the pickers can consult it.
    apiGet("api/tools")
      .then(function (data) { if (data.ok) state.tools = data.tools || {}; })
      .catch(function () {})
      .then(function () { return apiGet("api/participants"); })
      .then(function (data) {
        if (!data.ok) return;
        if (data.config) clipgenApplyConfig(data.config);
        state.hasSheet = !!data.has_sheet;
        state.participants = (data.participants || []).filter(function (p) { return p.has_video; });
        // Seed _videoVersions first so the preload queue gets the ?v= suffix.
        state.participants.forEach(function (p) {
          if (p.version != null) _videoVersions[p.id] = String(p.version);
        });
        renderParticipantSelect();
        var pickId = null;
        if (state.participants.length > 0) {
          var stored = getStoredUIState("screenspace");
          // Deep link (#P07, from the Overview Map) beats the stored pick.
          pickId = clipgenPickParticipant(state.participants, {
            hashPid: clipgenHashParticipant(),
            storedId: stored.selectedParticipant,
          }) || state.participants[0].id;
          var initialTs = getStoredUIMapEntry("screenspace", "videoTimeByParticipant", pickId);
          if (typeof initialTs !== "number") initialTs = undefined;
          selectParticipant(pickId, initialTs);
          state.runParticipants = [pickId];
          // Match the live tab strip so a new tab never silently falls back.
          if (stored.rightPaneTab
              && qs('.rp-tab[data-tab="' + CSS.escape(stored.rightPaneTab) + '"]')) {
            setRightPaneTab(stored.rightPaneTab);
          }
          if (stored.activeWorkflow) {
            var wfTab = qs('.wf-tab[data-type="' + CSS.escape(stored.activeWorkflow) + '"]');
            if (wfTab) wfTab.click();
          }
        }
        // Warm frame 0 for other participants; the selected one last, its request already
        // went out.
        var preloadOrder = [];
        state.participants.forEach(function (p) {
          if (p.id !== pickId) preloadOrder.push(p.id);
        });
        if (pickId) preloadOrder.push(pickId);
        queueFrameZeroPreload(preloadOrder);
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
          // Seed the result cache; with every task completed the SSE stream never starts.
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
  // Assigned synchronously, so fully populated before any satellite IIFE runs.
  var SS = (window.ClipgenScreenspace = window.ClipgenScreenspace || {});
  SS.state = state;
  // Calibration entry points are hub delegators; see the calibration strip section.
  SS.refreshCalibration = refreshCalibration;
  SS.updateCalibrationThresholdLine = updateCalibrationThresholdLine;
  // Hub helpers the satellites call outward.
  SS.fullFrameRegionRef = fullFrameRegionRef;
  SS.hideToolInfoTooltip = hideToolInfoTooltip;
  SS.renderIntervalSlot = renderIntervalSlot;
  SS.renderCalibration = renderCalibration;
  SS.updateCalibrationVisibility = updateCalibrationVisibility;
  SS.toolSupportsFastScan = toolSupportsFastScan;
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
  // The tasks satellite calls this after moving the active .wf-tab manually.
  SS.syncToolCategoryNav = syncToolCategoryNav;
  SS.iconSpan = iconSpan;
  SS.buildNormalizeControl = buildNormalizeControl;
  SS.buildColorModeControl = buildColorModeControl;
  SS._colorMode = _colorMode;
  SS.activatePipette = activatePipette;
  SS.deactivatePipette = deactivatePipette;
  SS.updateRunButton = updateRunButton;
  // Helpers for the tasks + results satellites; their own entry points publish from there.
  SS.applyColorMode = applyColorMode;
  SS.applyNormalizeMode = applyNormalizeMode;
  // Helpers the overlay satellite reads; it publishes SS.renderOverlay itself.
  SS.regionColorForIndex = regionColorForIndex;
  SS.getThemeColors = getThemeColors;
  SS.templateOverlayBounds = templateOverlayBounds;
  SS.renderRunRegionPicker = renderRunRegionPicker;
  SS.selectParticipant = selectParticipant;
  // Hub helper the color satellite reuses.
  SS.regionToPixels = regionToPixels;
  // Helpers the overlay-interaction satellite reads; it publishes renderRegionChips,
  // computeLabelRect, updateRegionButtons, invalidateOverlayRect back.
  SS.pauseVideo = pauseVideo;
  SS.pinCurrentFrame = pinCurrentFrame;
  SS.togglePinTrayVisibility = togglePinTrayVisibility;
  SS.clearAllPins = clearAllPins;
  SS.updatePinButtons = updatePinButtons;
  // Satellites call this after rebuilding a nav region's list wholesale.
  SS.ssRefreshNav = ssRefreshNav;

})();
