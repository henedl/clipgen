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

  var _paletteDocListeners = null;
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

  var FAST_SCAN_DESCRIPTIONS = {
    color: "Lower resolution, skips unchanged frames",
    change: "Lower resolution, skips unchanged frames",
    similarity: "Lower resolution, skips unchanged frames",
    text: "Skips unchanged frames",
    numbers: "Skips unchanged frames",
    template: "Downscales template 2\u00D7, skips unchanged frames",
    shape: "Downscales reference 2\u00D7, skips unchanged frames; thin outlines may vanish",
    flow: "Lower resolution, skips unchanged frames",
    scene: "Lower resolution, skips unchanged frames",
    inactivity: "Lower resolution, skips unchanged frames",
    multitool: "Skips unchanged frames, widens interval"
  };

  // Single source of truth for fast-scan support; timelapse and boundary opt out by
  // omission.
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

  // ---- Region stashing ----

  // Stash id whose card gets the landing animation; consumed on first render.
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
    }).catch(toastError("Could not create stash"));
  }

  function dismissStash(stashId) {
    apiDelete("api/stashes/" + stashId).then(function (data) {
      if (!data.ok) return;
      state.stashes = state.stashes.filter(function (s) { return s.id !== stashId; });
      // Remove any run-selected regions that belonged to this stash
      renderRunRegionPicker();
      renderStashCards();
      showToast("Stash dismissed");
    }).catch(toastError("Could not delete stash"));
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
    }).catch(toastError("Could not restore stash"));
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
    }).catch(toastError("Could not rename stash"));
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

  // Drop handlers for copying a chip into a stash, bound once on the persistent
  // #stashArea.
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

  // ---- Region chip drag ----
  // Mirrors the vertical task drag; excluding .dragging keeps indices aligned.
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
    // Firefox/Safari may hide custom MIME types on dragover; the local flag is reliable.
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
      // Colors are position-based, so reordering recolors regions on purpose; repaint
      // both.
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

  function _toolLabel(type) {
    return type ? type.charAt(0).toUpperCase() + type.slice(1) : "";
  }

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

    var sampleBtn = el("button", "btn btn-small color-sample-btn");
    sampleBtn.id = "colorSampleBtn";
    sampleBtn.addEventListener("click", sampleColorFromRegion);
    inputRow.appendChild(sampleBtn);
    updateColorSampleBtnLabel();

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
    // Region-aware readout; runs after addParamRow's generic listener and overwrites the
    // raw value.
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
      state.capturedRefPreview = captureRefSnapshot(state.currentTimestamp);
      renderWorkflowParams();
      showToast("Reference frame captured at " + formatTime(state.currentTimestamp, { decimals: 1 }));
    });
    refControl.appendChild(refBtn);
    if (state.referenceTimestamp !== null) {
      refControl.appendChild(refTimeChip(state.referenceTimestamp));
      var simSnapInfo = refSnapshotInfo("reference");
      if (simSnapInfo) refControl.appendChild(simSnapInfo);
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

  // Client-side same-frame crop of the region for the reference row; the run re-extracts
  // server-side.
  function captureRefSnapshot(ts) {
    var img = state.frameImage;
    var regs = state.previewRegions || state.regions;
    var name = state.activeRegion && regs[state.activeRegion] ? state.activeRegion : null;
    if (!name) {
      for (var i = state.runRegions.length - 1; i >= 0; i--) {
        var ref = normalizeRegionRef(state.runRegions[i]);
        if (ref && ref.name && regs[ref.name]) { name = ref.name; break; }
      }
    }
    var r = name && regs[name];
    if (!img || !img.naturalWidth || !r) return null;
    var sw = Math.max(1, Math.round(r.w * img.naturalWidth));
    var sh = Math.max(1, Math.round(r.h * img.naturalHeight));
    var c = document.createElement("canvas");
    c.width = sw;
    c.height = sh;
    c.getContext("2d").drawImage(
      img,
      r.x * img.naturalWidth, r.y * img.naturalHeight, sw, sh,
      0, 0, sw, sh
    );
    return { dataUrl: c.toDataURL("image/png"), region: name, ts: ts };
  }

  // Last-capture thumbnail + region label. A task-Edit restore has no pixels, so thumbnail
  // is optional.
  function refSnapshotInfo(labelText, editable) {
    var snap = state.capturedRefPreview;
    if (!snap || snap.ts !== state.referenceTimestamp) return null;
    if (!snap.dataUrl && !snap.region) return null;
    var capInfo = el("span", "param-value template-upload-info");
    if (snap.dataUrl) {
      var capThumb = document.createElement("img");
      capThumb.decoding = "async";
      capThumb.className = "ss-sample-thumb";
      capThumb.src = snap.dataUrl;
      capThumb.alt = "Captured " + labelText.toLowerCase();
      capThumb.title = snap.region;
      capThumb.addEventListener("click", function () {
        openSampleModal({
          mode: editable ? "edit" : "view",
          title: "Captured " + labelText.toLowerCase(),
          dataUrl: snap.dataUrl,
          regionName: snap.region,
          onApply: function (b64) {
            // An edited capture becomes an upload; the server cannot re-derive its pixels.
            applyEditedSample((snap.region || "sample") + "-edited.png", b64);
          },
        });
      });
      capInfo.appendChild(capThumb);
    }
    if (snap.region) capInfo.appendChild(el("span", "param-hint", snap.region));
    return capInfo;
  }

  // Install an edited sample as the uploaded reference (Template/Shape).
  function applyEditedSample(name, b64) {
    state.uploadedTemplate = { name: name, data: b64 };
    state.referenceTimestamp = null;
    state.capturedRefPreview = null;
    state.templateOverlayPos = null;
    var previewImg = new Image();
    previewImg.onload = function () { renderOverlay(); };
    previewImg.src = "data:image/png;base64," + b64;
    state.uploadedTemplateImg = previewImg;
    renderWorkflowParams();
    updateRunButton();
    refreshModelView({ debounce: true });
  }

  // Template + Shape reference row; the drag overlay stays template-only. opts.draw adds
  // paint-on-frame.
  function renderRefCaptureRow(container, labelText, opts) {
    var tmplRefRow = el("div", "param-row");
    tmplRefRow.appendChild(el("span", "param-label", labelText));
    var tmplRefCtrl = el("div", "param-control");
    var tmplCapBtn = el("button", "btn btn-small ss-template-icon-btn ss-template-icon-btn--capture");
    tmplCapBtn.setAttribute("type", "button");
    tmplCapBtn.title = "Capture Region";
    tmplCapBtn.setAttribute("aria-label", "Capture Region");
    var tmplCapGlyph = el("span", "ss-template-icon-btn__glyph");
    tmplCapBtn.appendChild(tmplCapGlyph);
    tmplCapBtn.addEventListener("click", function () {
      var snap = captureRefSnapshot(state.currentTimestamp);
      // Refuse a region-less capture: the sample would silently become the run region.
      if (!snap) {
        showToast("Select or draw a region to capture from");
        return;
      }
      state.referenceTimestamp = state.currentTimestamp;
      state.uploadedTemplate = null;
      state.capturedRefPreview = snap;
      renderWorkflowParams();
      showToast(labelText + " captured at " + formatTime(state.currentTimestamp, { decimals: 1 }));
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
        state.capturedRefPreview = null;
        state.templateOverlayPos = null;
        var previewImg = new Image();
        previewImg.onload = function () { renderOverlay(); };
        previewImg.src = dataUrl;
        state.uploadedTemplateImg = previewImg;
        renderWorkflowParams();
        showToast(labelText + " loaded");
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

    if (opts && opts.draw) {
      var drawBtn = el("button", "btn btn-small ss-template-icon-btn ss-template-icon-btn--draw");
      drawBtn.setAttribute("type", "button");
      drawBtn.title = "Draw shape on frame";
      drawBtn.setAttribute("aria-label", "Draw shape on frame");
      drawBtn.appendChild(el("span", "ss-template-icon-btn__glyph"));
      // The row re-renders wholesale; active look derives from state alone.
      drawBtn.classList.toggle("active", !!state.shapeDraw);
      drawBtn.addEventListener("click", toggleShapeDraw);
      tmplRefCtrl.appendChild(drawBtn);
    }

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
      uploadThumb.className = "ss-sample-thumb";
      uploadThumb.src = "data:image/png;base64," + state.uploadedTemplate.data;
      uploadThumb.alt = "Uploaded " + labelText.toLowerCase();
      uploadThumb.title = state.uploadedTemplate.name;
      uploadThumb.addEventListener("click", function () {
        var up = state.uploadedTemplate;
        if (!up) return;
        openSampleModal({
          mode: "edit",
          title: up.name || "Uploaded " + labelText.toLowerCase(),
          dataUrl: "data:image/png;base64," + up.data,
          onApply: function (b64) {
            applyEditedSample(up.name || "sample.png", b64);
          },
        });
      });
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
      tmplRefCtrl.appendChild(refTimeChip(state.referenceTimestamp));
      // Task-Edit restores have no snapshot; the ts guard falls back to the time chip.
      var capInfo = refSnapshotInfo(labelText, true);
      if (capInfo) tmplRefCtrl.appendChild(capInfo);
    }
    tmplRefRow.appendChild(tmplRefCtrl);
    container.appendChild(tmplRefRow);
  }

  function renderTemplateParams(container) {
    renderRefCaptureRow(container, "Template");
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

  function renderShapeParams(container) {
    renderRefCaptureRow(container, "Shape", { draw: true });
    addParamRow(container, "Threshold", rangeInput("paramShapeThresh", 0.30, 1.00, 0.55, 0.01), "paramShapeThreshVal");
    addParamRow(container, "Scale min", rangeInput("paramShapeScaleMin", 25, 400, 50, 5), "paramShapeScaleMinVal");
    var rowXMin = container.lastChild;
    addParamRow(container, "Scale max", rangeInput("paramShapeScaleMax", 25, 400, 200, 5), "paramShapeScaleMaxVal");
    var rowXMax = container.lastChild;
    addParamRow(container, "Scale steps", numberInput("paramShapeSteps", 1, 12, 7, 1));
    var rowXSteps = container.lastChild;
    // Unlinked axes: width-only sliders plus a height ladder; the sweep multiplies cost.
    var linkCb = document.createElement("input");
    linkCb.type = "checkbox";
    linkCb.id = "paramShapeLinkAxes";
    linkCb.checked = true;
    addParamRow(container, "Link axes", linkCb);
    addParamRow(container, "Height scale min", rangeInput("paramShapeScaleYMin", 25, 400, 90, 5), "paramShapeScaleYMinVal");
    var rowYMin = container.lastChild;
    addParamRow(container, "Height scale max", rangeInput("paramShapeScaleYMax", 25, 400, 110, 5), "paramShapeScaleYMaxVal");
    var rowYMax = container.lastChild;
    addParamRow(container, "Height scale steps", numberInput("paramShapeStepsY", 1, 12, 3, 1));
    var rowYSteps = container.lastChild;
    function syncAxisRows() {
      var show = !linkCb.checked;
      rowYMin.style.display = show ? "" : "none";
      rowYMax.style.display = show ? "" : "none";
      rowYSteps.style.display = show ? "" : "none";
      // Unlinked, the base ladder is width-only; say so in its labels.
      rowXMin.firstChild.textContent = show ? "Width scale min" : "Scale min";
      rowXMax.firstChild.textContent = show ? "Width scale max" : "Scale max";
      rowXSteps.firstChild.textContent = show ? "Width scale steps" : "Scale steps";
    }
    linkCb.addEventListener("change", syncAxisRows);
    syncAxisRows();
    renderIntervalSlot("paramShapeInterval", 0.5, 60, 1.0, 0.5);
  }

  function renderSceneParams(container) {
    var sceneList = el("div", "scene-reference-list");
    sceneList.id = "sceneRefList";
    state.sceneReferences.forEach(function (ref, i) {
      if (ref.threshold === undefined) ref.threshold = 0.75;
      var item = el("div", "scene-ref-item");
      item.appendChild(el("span", "scene-ref-name", ref.name));
      if (ref._thumb) {
        var scThumbWrap = el("span", "param-value template-upload-info");
        var scThumb = document.createElement("img");
        scThumb.decoding = "async";
        scThumb.className = "ss-sample-thumb";
        scThumb.src = ref._thumb;
        scThumb.alt = "Scene sample";
        if (ref._thumbRegion) scThumb.title = ref._thumbRegion;
        scThumb.addEventListener("click", function () {
          openSampleModal({ mode: "view", title: ref.name || "Scene sample", dataUrl: ref._thumb });
        });
        scThumbWrap.appendChild(scThumb);
        item.appendChild(scThumbWrap);
      }
      item.appendChild(refTimeChip(ref.timestamp));
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
      // Underscore fields are display-only; gather strips them before the server sees
      // refs.
      var scSnap = captureRefSnapshot(state.currentTimestamp);
      state.sceneReferences.push({
        name: name,
        timestamp: state.currentTimestamp,
        threshold: 0.75,
        _thumb: scSnap && scSnap.dataUrl,
        _thumbRegion: scSnap && scSnap.region,
      });
      renderWorkflowParams();
      showToast("Scene '" + name + "' at " + formatTime(state.currentTimestamp, { decimals: 1 }));
    });
    addScCtrl.appendChild(scCapBtn);
    addScRow.appendChild(addScCtrl);
    container.appendChild(addScRow);
    renderIntervalSlot("paramSceneInterval", 0.5, 60, 1.0, 0.5);
  }

  // DOM-only param inputs snapshot by tool-prefixed id across rebuilds; multitool steps
  // (positional ids) self-restore.
  var _MT_STEP_ID = /_mt\d+$/;

  function _paramControlValue(el) {
    return el.type === "checkbox" ? el.checked : el.value;
  }

  // The id'd form controls under `root` that participate in save/restore/reset.
  function _paramControls(root) {
    var out = [];
    var nodes = root.querySelectorAll("[id]");
    for (var i = 0; i < nodes.length; i++) {
      var tag = nodes[i].tagName;
      if (tag !== "INPUT" && tag !== "SELECT" && tag !== "TEXTAREA") continue;
      if (_MT_STEP_ID.test(nodes[i].id)) continue;
      out.push(nodes[i]);
    }
    return out;
  }

  function _snapshotParamValues() {
    var map = {};
    ["#workflowParams", "#workflowIntervalSlot"].forEach(function (id) {
      var root = qs(id);
      if (!root) return;
      _paramControls(root).forEach(function (el) {
        map[el.id] = _paramControlValue(el);
      });
    });
    return map;
  }

  function _mergeParamMap(target, src) {
    Object.keys(src).forEach(function (id) { target[id] = src[id]; });
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
      } else if (el.type === "hidden" && el.parentNode
                 && el.parentNode.classList.contains("cg-segtrack")) {
        // Capsule visuals are CSS off the track, so el.value alone desyncs; apply* also
        // re-toggles rows.
        if (el.value === String(saved)) return;
        if (id === "paramColorMode") applyColorMode(id, saved);
        else segTrackSetValue(el.parentNode, saved);
      } else {
        if (el.value === String(saved)) return;
        el.value = saved;
      }
      // Fire input for readouts and model view, change for change-only listeners (numbers
      // operator row).
      el.dispatchEvent(new Event("input", { bubbles: true }));
      el.dispatchEvent(new Event("change", { bubbles: true }));
    });
  }

  // Only setTargetColor repaints the swatch, hex field and palette; feed it the restored
  // values.
  function _restoreColorTarget() {
    if (state.activeWorkflow !== "color") return;
    var c = SS.getColorHiddenInputs();
    if (!c) return;
    setTargetColor(
      numberOrDefault(c.h.value, 90), numberOrDefault(c.s.value, 200), numberOrDefault(c.v.value, 200)
    );
  }

  // `opts.defaults` skips the restore; task restore writes its own saved params next.
  function renderWorkflowParams(opts) {
    if (state.shapeDraw && state.activeWorkflow !== "shape") cancelShapeDraw();
    _mergeParamMap(state.paramValues, _snapshotParamValues());
    _renderWorkflowParamsBuild();
    // Defaults are read off the just-built panel, not a second table that could drift.
    _mergeParamMap(state.paramDefaults, _snapshotParamValues());
    if (!(opts && opts.defaults)) {
      _restoreParamValues(state.paramValues);
      _restoreColorTarget();
    }
    updateParamResetButtons();
  }

  // ---- Reset-to-default ----
  // Per row, not control: the Range row's two inputs are one parameter.
  function _buildParamResetButton(row) {
    var btn = el("button", "param-reset hidden");
    btn.type = "button";
    var icon = el("span", "param-reset-icon");
    applyIconMask(icon, "arrow-path", "/screenspace/icons/");
    btn.appendChild(icon);
    btn.addEventListener("click", function () {
      var map = {};
      _paramControls(row).forEach(function (c) {
        if (state.paramDefaults[c.id] !== undefined) map[c.id] = state.paramDefaults[c.id];
      });
      // No _restoreColorTarget: the H/S/V inputs sit outside any .param-row.
      _restoreParamValues(map);
      updateParamResetButtons();
    });
    return btn;
  }

  function _syncParamResetButton(row) {
    var ctrl = row.querySelector(".param-control");
    if (!ctrl) return;
    var btn = row.querySelector(".param-reset");
    var eligible = _paramControls(row).filter(function (c) {
      return state.paramDefaults[c.id] !== undefined;
    });
    if (!eligible.length) {
      if (btn) btn.parentNode.removeChild(btn);
      return;
    }
    var changed = eligible.some(function (c) {
      return _paramControlValue(c) !== state.paramDefaults[c.id];
    });
    if (!btn) {
      btn = _buildParamResetButton(row);
      ctrl.appendChild(btn);
    }
    var tip = "Reset to default";
    if (eligible.length === 1 && eligible[0].type !== "checkbox") {
      tip += " (" + state.paramDefaults[eligible[0].id] + ")";
    }
    btn.setAttribute("data-tooltip", tip);
    btn.setAttribute("aria-label", tip);
    btn.classList.toggle("hidden", !changed);
  }

  function updateParamResetButtons() {
    var container = qs("#workflowParams");
    if (container) {
      _paramRows(container).forEach(_syncParamResetButton);
    }
    // The interval control has no .param-row wrapper, so it is its own row.
    var slot = qs("#workflowIntervalSlot");
    if (slot) _syncParamResetButton(slot);
  }

  function _paramRows(container) {
    return Array.prototype.slice.call(container.querySelectorAll(".param-row"));
  }

  // One delegated pair per stable container (like initCalibration) catches hand-assembled
  // rows too.
  function initParamResets() {
    ["workflowParams", "workflowIntervalSlot"].forEach(function (id) {
      var container = qs("#" + id);
      if (!container) return;
      container.addEventListener("input", updateParamResetButtons);
      container.addEventListener("change", updateParamResetButtons);
    });
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
      var mtEventLabel = textInput("paramEventLabel", "e.g. low_health");
      mtEventLabel.className = "param-input-half";
      addParamRow(container, "Event label", mtEventLabel);
      var dfCb = document.createElement("input");
      dfCb.type = "checkbox";
      dfCb.id = "paramDetectFirst";
      addParamRow(container, "Detect first", dfCb);
      // Every multitool mutation lands here; task-import and reorder paths skip
      // updateRunButton.
      updateRunButton();
      updateCalibrationVisibility();
      // Multitool returns early, so mirror the bottom-of-function calibration reset here.
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
    else if (type === "shape") renderShapeParams(container);
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
      // Metric: Auto sends nothing (server default, currently Hybrid); pHash is the v1
      // spike detector.
      var metricSel = document.createElement("select");
      metricSel.id = "paramBoundaryMetric";
      [["", "Auto"], ["scene", "Scene"], ["phash", "pHash"], ["hybrid", "Hybrid"]].forEach(function (pair) {
        var opt = el("option", null, pair[1]);
        opt.value = pair[0];
        metricSel.appendChild(opt);
      });
      addParamRow(container, "Metric", metricSel);
      // Full phash Hamming span (0..64); drives the pHash metric and Hybrid's spike check.
      addParamRow(container, "Sensitivity", rangeInput("paramBoundaryThresh", 0, 64, 14, 1), "paramBoundaryThreshVal");
      addParamRow(container, "Min gap (s)", numberInput("paramBoundaryMinGap", 0.5, 60, 3.0, 0.5));
      renderIntervalSlot("paramBoundaryInterval", 0.5, 60, 1.0, 0.5);
    }
    else if (type === "attention") {
      // Peak-jump distance as a fraction of the screen diagonal, and the EMA smoothing
      // alpha.
      addParamRow(container, "Sensitivity", rangeInput("paramAttnShift", 0.05, 0.50, 0.15, 0.01), "paramAttnShiftVal");
      addParamRow(container, "Smoothing", rangeInput("paramAttnSmooth", 0.1, 1.0, 0.6, 0.05), "paramAttnSmoothVal");
      // Channel weights (defaults mirror SCREENSPACE_ATTENTION_WEIGHT_*); Faces at 0
      // disables the Haar channel.
      addParamRow(container, "Spectral wt.", rangeInput("paramAttnWSpectral", 0, 2.0, 1.0, 0.05), "paramAttnWSpectralVal");
      addParamRow(container, "Contrast wt.", rangeInput("paramAttnWContrast", 0, 2.0, 0.7, 0.05), "paramAttnWContrastVal");
      addParamRow(container, "Motion wt.", rangeInput("paramAttnWMotion", 0, 2.0, 1.2, 0.05), "paramAttnWMotionVal");
      addParamRow(container, "Faces wt.", rangeInput("paramAttnWFace", 0, 2.0, 0, 0.05), "paramAttnWFaceVal");
      addParamRow(container, "Center bias", rangeInput("paramAttnCenterBias", 0, 1.0, 0.25, 0.05), "paramAttnCenterBiasVal");
      renderIntervalSlot("paramAttnInterval", 0.5, 60, 0.5, 0.5);
    }

    if (type !== "timelapse") {
      var eventLabel = textInput("paramEventLabel", "e.g. low_health");
      eventLabel.className = "param-input-half";
      addParamRow(container, "Event label", eventLabel);
      // Boundary and attention emit transitions, not detections, so "Detect first" doesn't
      // apply.
      if (type !== "boundary" && type !== "attention") {
        var dfCb = document.createElement("input");
        dfCb.type = "checkbox";
        dfCb.id = "paramDetectFirst";
        addParamRow(container, "Detect first", dfCb);
      }
    }

    var scanPicker = qs("#runScanModePicker");
    // Timelapse has no scan modes; boundary runs its own coarse pass. Hide for both.
    if (scanPicker) {
      scanPicker.style.display = toolSupportsFastScan(type) ? "" : "none";
    }
    var scanBtn = scanPicker && scanPicker.querySelector(".scan-toggle-btn");
    if (scanBtn && scanBtn._updateScanState) scanBtn._updateScanState();

    updateRunButton();
    _updateOverlayUi();
    refreshModelView();
    updateCalibrationVisibility();
    // Drop stale scores before the new evaluation; renderCalibration() resets
    // calibrationGreen synchronously.
    state.calibrationResult = null;
    renderCalibration();
    if (!state.suppressCalibrationRefresh) refreshCalibration();
  }

  // Click-to-seek reference timestamp chip shared by every capturing tool; `textId` marks
  // the in-place-updated span.
  function refTimeChip(seconds, textId) {
    var wrap = el("span", "param-value param-value--ref");
    var text = el("span", null, formatTime(seconds, { decimals: 1 }));
    if (textId) text.id = textId;
    wrap.appendChild(text);

    var btn = el("button", "ref-seek-btn");
    btn.type = "button";
    btn.setAttribute("data-tooltip", "Jump to this frame");
    btn.setAttribute("aria-label", "Jump to this frame");
    var icon = el("span", "ref-seek-icon");
    applyIconMask(icon, "arrow-up-right", "/screenspace/icons/");
    btn.appendChild(icon);
    btn.addEventListener("click", function () {
      loadFrame(seconds);
    });
    wrap.appendChild(btn);
    return wrap;
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
      // Template/shape with an uploaded image can run without a region
      // (full-frame scan).
      if (!isMultitool && !isFullFrameTool && regions.length === 0
          && !((type === "template" || type === "shape") && state.uploadedTemplate)) return;
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
    // opts.silent drops missing-input toasts for the calibration strip's per-keystroke
    // probes.
    var silent = !!(opts && opts.silent);
    function toast(msg) { if (!silent) showToast(msg); }
    var sfx = "_mt" + idx;
    // Suffix-aware readers; rawNum deliberately yields NaN for inputs whose emptiness is
    // checked.
    function num(id, d) { return numberOrDefault((qs("#" + id + sfx) || {}).value, d); }
    function intv(id, d) { return intOrDefault((qs("#" + id + sfx) || {}).value, d); }
    function chk(id) { return !!((qs("#" + id + sfx) || {}).checked); }
    function str(id, d) { return (qs("#" + id + sfx) || {}).value || d; }
    function rawNum(id) { return parseFloat((qs("#" + id + sfx) || {}).value); }
    var p = {};
    if (stepType === "color") {
      p.target_color = {
        h: num("paramColorH", 0),
        s: num("paramColorS", 0),
        v: num("paramColorV", 0),
      };
      var tol = num("paramColorTol", 30);
      p.tolerance = {
        h: Math.round(tol * 90 / 100),
        s: Math.round(tol * 128 / 100),
        v: Math.round(tol * 128 / 100),
      };
      if (str("paramColorMode", "") === "presence") {
        p.color_mode = "presence";
        p.min_coverage = num("paramColorMinArea", 1) / 100;
      }
    } else if (stepType === "change") {
      p.threshold = num("paramChangeThresh", 0.03);
      p.noise_threshold = intv("paramChangeNoise", 30);
    } else if (stepType === "similarity") {
      var step = state.multitoolSteps[idx];
      if (!step || step._refTs === undefined) {
        toast("Step " + (idx + 1) + ": capture a reference frame first");
        return null;
      }
      p.reference_timestamp = step._refTs;
      p.threshold = num("paramSimThresh", 0.90);
    } else if (stepType === "text") {
      p.search_string = str("paramTextSearch", "");
      if (!p.search_string.trim()) {
        toast("Step " + (idx + 1) + ": enter a search string");
        return null;
      }
      p.fuzzy_threshold = num("paramTextFuzzy", CLIPGEN_CONFIG.screenspaceOcrFuzzyThreshold);
      p.ocr_confidence_threshold = num("paramTextOcrConf", CLIPGEN_CONFIG.screenspaceOcrMinConfidence);
      p.ocr_preprocess = chk("paramTextOcrPreprocess");
      p.ocr_normalize = str("paramTextOcrNormalize", "off");
    } else if (stepType === "numbers") {
      p.operator = str("paramNumOperator", "gt");
      p.target_value = rawNum("paramNumTarget");
      if (isNaN(p.target_value)) {
        toast("Step " + (idx + 1) + ": enter a valid target number");
        return null;
      }
      p.ocr_confidence_threshold = num("paramNumOcrConf", CLIPGEN_CONFIG.screenspaceOcrMinConfidence);
      p.ocr_preprocess = chk("paramNumOcrPreprocess");
      p.integers_only = chk("paramNumIntegersOnly");
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
      p.threshold = num("paramTemplateThresh", 0.70);
      var tScalePct = rawNum("paramTemplateScale");
      if (!isNaN(tScalePct) && tScalePct > 0 && tScalePct !== 100) {
        p.template_scale = tScalePct / 100;
      }
    } else if (stepType === "flow") {
      p.magnitude_threshold = num("paramFlowMag", 2.0);
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
      p.threshold = intv("paramInactThresh", 10);
    }
    var stepRegionRef = normalizeRegionRef(state.multitoolSteps[idx].region_ref)
      || (state.multitoolSteps[idx].region ? activeRegionRef(state.multitoolSteps[idx].region) : null);
    p.region = stepRegionRef ? stepRegionRef.name : "";
    if (stepRegionRef) p.region_ref = regionRefPayload(stepRegionRef);
    return p;
  }

  function gatherWorkflowParams(type, opts) {
    // opts.silent drops missing-input toasts for the calibration strip's per-keystroke
    // probes.
    var silent = !!(opts && opts.silent);
    function toast(msg) { if (!silent) showToast(msg); }
    var sfx = "";
    // Suffix-aware readers; rawNum deliberately yields NaN for inputs whose emptiness is
    // checked.
    function num(id, d) { return numberOrDefault((qs("#" + id + sfx) || {}).value, d); }
    function intv(id, d) { return intOrDefault((qs("#" + id + sfx) || {}).value, d); }
    function chk(id) { return !!((qs("#" + id + sfx) || {}).checked); }
    function str(id, d) { return (qs("#" + id + sfx) || {}).value || d; }
    function rawNum(id) { return parseFloat((qs("#" + id + sfx) || {}).value); }
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
      params.interval = num("paramMultitoolInterval", 1.0);
      var mtLabelEl = qs("#paramEventLabel");
      if (mtLabelEl && mtLabelEl.value.trim()) params.event_label = mtLabelEl.value.trim();
      var mtDfEl = qs("#paramDetectFirst");
      if (mtDfEl && mtDfEl.checked) params.detect_first = true;
      return params;
    } else if (type === "color") {
      params.target_color = {
        h: num("paramColorH", 0),
        s: num("paramColorS", 0),
        v: num("paramColorV", 0),
      };
      var tol = num("paramColorTol", 30);
      params.tolerance = {
        h: Math.round(tol * 90 / 100),
        s: Math.round(tol * 128 / 100),
        v: Math.round(tol * 128 / 100),
      };
      if (str("paramColorMode", "") === "presence") {
        params.color_mode = "presence";
        params.min_coverage = num("paramColorMinArea", 1) / 100;
      }
      params.interval = num("paramColorInterval", 1.0);
    } else if (type === "change") {
      params.threshold = num("paramChangeThresh", 0.03);
      params.noise_threshold = intv("paramChangeNoise", 30);
      params.interval = num("paramChangeInterval", 1.0);
      var rcChange = intv("paramChangeConsecutive", 1);
      if (rcChange > 1) params.require_consecutive = rcChange;
    } else if (type === "similarity") {
      if (state.referenceTimestamp === null) {
        toast("Capture a reference frame first");
        return null;
      }
      params.reference_timestamp = state.referenceTimestamp;
      params.threshold = num("paramSimThresh", 0.90);
      params.interval = num("paramSimInterval", 1.0);
    } else if (type === "text") {
      params.search_string = str("paramTextSearch", "");
      if (!params.search_string.trim()) {
        toast("Enter a search string");
        return null;
      }
      params.fuzzy_threshold = num("paramTextFuzzy", CLIPGEN_CONFIG.screenspaceOcrFuzzyThreshold);
      params.ocr_confidence_threshold = num("paramTextOcrConf", CLIPGEN_CONFIG.screenspaceOcrMinConfidence);
      params.ocr_preprocess = chk("paramTextOcrPreprocess");
      params.ocr_normalize = str("paramTextOcrNormalize", "off");
      params.interval = num("paramTextInterval", 2.0);
      var lang = str("paramTextLang", "en");
      params.languages = [lang];
      var rcText = intv("paramTextConsecutive", 1);
      if (rcText > 1) params.require_consecutive = rcText;
    } else if (type === "numbers") {
      var op = str("paramNumOperator", "gt");
      params.operator = op;
      if (op === "range") {
        params.range_min = rawNum("paramNumMin");
        params.range_max = rawNum("paramNumMax");
        if (isNaN(params.range_min) || isNaN(params.range_max)) {
          toast("Enter valid min and max values");
          return null;
        }
        if (params.range_min > params.range_max) {
          toast("Min must be less than or equal to max");
          return null;
        }
      } else {
        params.target_value = rawNum("paramNumTarget");
        if (isNaN(params.target_value)) {
          toast("Enter a valid target number");
          return null;
        }
      }
      params.ocr_confidence_threshold = num("paramNumOcrConf", CLIPGEN_CONFIG.screenspaceOcrMinConfidence);
      params.ocr_preprocess = chk("paramNumOcrPreprocess");
      params.integers_only = chk("paramNumIntegersOnly");
      params.interval = num("paramNumInterval", 2.0);
      var rcNum = intv("paramNumConsecutive", 1);
      if (rcNum > 1) params.require_consecutive = rcNum;
    } else if (type === "timelapse") {
      params.speedup_factor = num("paramTlSpeed", 10);
      var si = rawNum("paramTlSampleInterval");
      if (si > 0) params.sample_interval = si;
      params.output_format = str("paramTlFormat", "mp4");
    } else if (type === "template") {
      if (state.uploadedTemplate) {
        params.template_image_data = state.uploadedTemplate.data;
        if (state.uploadedTemplate.name) params.template_name = state.uploadedTemplate.name;
      } else if (state.referenceTimestamp !== null) {
        params.reference_timestamp = state.referenceTimestamp;
        // Pin the sample to its capture region; the run target only scopes the search.
        var tplSnap = state.capturedRefPreview;
        if (tplSnap && tplSnap.ts === state.referenceTimestamp && tplSnap.region) {
          params.reference_region = tplSnap.region;
        }
      } else {
        toast("Capture a template region or upload a PNG");
        return null;
      }
      params.threshold = num("paramTemplateThresh", 0.70);
      params.interval = num("paramTemplateInterval", 1.0);
      var scalePct = rawNum("paramTemplateScale");
      if (!isNaN(scalePct) && scalePct > 0 && scalePct !== 100) {
        params.template_scale = scalePct / 100;
      }
    } else if (type === "shape") {
      if (state.uploadedTemplate) {
        params.shape_image_data = state.uploadedTemplate.data;
        if (state.uploadedTemplate.name) params.shape_name = state.uploadedTemplate.name;
      } else if (state.referenceTimestamp !== null) {
        params.reference_timestamp = state.referenceTimestamp;
        // Pin the sample to its capture region; the run target only scopes the search.
        var capSnap = state.capturedRefPreview;
        if (capSnap && capSnap.ts === state.referenceTimestamp && capSnap.region) {
          params.reference_region = capSnap.region;
        }
      } else {
        toast("Capture a shape region or upload a PNG");
        return null;
      }
      params.threshold = num("paramShapeThresh", 0.55);
      params.scale_min = num("paramShapeScaleMin", 50) / 100;
      params.scale_max = num("paramShapeScaleMax", 200) / 100;
      params.scale_steps = intv("paramShapeSteps", 7);
      var linkEl = qs("#paramShapeLinkAxes");
      if (linkEl && !linkEl.checked) {
        params.scale_y_min = num("paramShapeScaleYMin", 90) / 100;
        params.scale_y_max = num("paramShapeScaleYMax", 110) / 100;
        params.scale_y_steps = intv("paramShapeStepsY", 3);
      }
      params.interval = num("paramShapeInterval", 1.0);
    } else if (type === "flow") {
      params.magnitude_threshold = num("paramFlowMag", 2.0);
      params.interval = num("paramFlowInterval", 1.0);
      var rcFlow = intv("paramFlowConsecutive", 1);
      if (rcFlow > 1) params.require_consecutive = rcFlow;
    } else if (type === "scene") {
      if (state.sceneReferences.length === 0) {
        toast("Add at least one scene reference");
        return null;
      }
      params.scene_references = state.sceneReferences.map(function (ref) {
        return { name: ref.name, timestamp: ref.timestamp, threshold: numberOrDefault(ref.threshold, 0.75) };
      });
      params.interval = num("paramSceneInterval", 1.0);
    } else if (type === "inactivity") {
      params.threshold = intv("paramInactThresh", 10);
      params.min_duration = num("paramInactMinDur", 2.0);
      params.interval = num("paramInactInterval", 1.0);
    } else if (type === "boundary") {
      params.threshold = intv("paramBoundaryThresh", 14);
      params.min_gap = num("paramBoundaryMinGap", 3.0);
      params.interval = num("paramBoundaryInterval", 1.0);
      // Auto ("") omits metric so the server applies its configured default.
      var boundaryMetric = str("paramBoundaryMetric", "");
      if (boundaryMetric) params.metric = boundaryMetric;
    } else if (type === "attention") {
      params.shift_threshold = num("paramAttnShift", 0.15);
      params.ema_alpha = num("paramAttnSmooth", 0.6);
      params.weight_spectral = num("paramAttnWSpectral", 1.0);
      params.weight_contrast = num("paramAttnWContrast", 0.7);
      params.weight_motion = num("paramAttnWMotion", 1.2);
      params.weight_face = num("paramAttnWFace", 0);
      params.center_bias = num("paramAttnCenterBias", 0.25);
      params.interval = num("paramAttnInterval", 0.5);
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
          if (!state.overlayImage || state.overlayImageTimestamp !== curTs || state.overlayImageTool !== state.activeWorkflow) {
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
      // Rebuild on every open so the boundary item tracks participants;
      // refreshExportStatus rebuilds on flag flips.
      rebuild();
      window.ClipgenExportActions.refreshExportStatus(rebuild);
    });
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
            title: "Switch to " + _toolLabel(type) + " tool",
            icon: TOOL_ICON_NAMES[type] || "cube",
            keywords: "tool detector select analysis " + cat.label.toLowerCase() + " " + type,
            section: "Tools",
            run: function () { selectWorkflowType(type); },
          });
        });
      });
      // "Jump to …" selects in place; the built-in provider adds cross-page "Open … in
      // <Page>".
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

    // Load initial data
    apiGet("api/participants")
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
  SS.gatherWorkflowParams = gatherWorkflowParams;
  SS.loadFrame = loadFrame;
  SS.seekPlayhead = seekPlayhead;
  SS.refTimeChip = refTimeChip;
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
  SS.renderWorkflowParams = renderWorkflowParams;
  SS.updateParamResetButtons = updateParamResetButtons;
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
  SS.stashRegions = stashRegions;
  SS.pinCurrentFrame = pinCurrentFrame;
  SS.togglePinTrayVisibility = togglePinTrayVisibility;
  SS.clearAllPins = clearAllPins;
  SS.updatePinButtons = updatePinButtons;
  // Satellites call this after rebuilding a nav region's list wholesale.
  SS.ssRefreshNav = ssRefreshNav;

})();
