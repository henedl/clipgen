/* clipgen Screenspace */

(function () {
  "use strict";

  var THEME_STORAGE_KEY = "clipgen-screenspace-theme";
  var POLL_INTERVAL = 2000;
  var FRAME_STEP = 1.0;

  var TASK_COLORS = {
    color: "#8b5cf6",
    change: "#f97316",
    similarity: "#0ea5e9",
    text: "#10b981",
    numbers: "#eab308",
    timelapse: "#ec4899",
  };

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
    tasks: [],
    selectedTaskId: null,
    hoveredTaskId: null,
    selectedTaskResults: null,
    pollTimer: null,
    queuePaused: false,
    timelineDragging: false,
    panelHeight: 260,
    previewMaxWidth: 100,
    taskFilter: null,
  };

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
    return "api/video/frame/" + encodeURIComponent(pid) + "/" + ts;
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

  function selectParticipant(pid) {
    state.selectedParticipant = pid;
    state.currentTimestamp = 0;
    state.videoInfo = null;
    state.referenceTimestamp = null;
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
        loadFrame(0);
      })
      .catch(function () { showToast("Failed to load video info"); });
  }

  // ---- Frame viewer ----

  function loadFrame(timestamp) {
    if (!state.selectedParticipant) return;
    if (state.frameLoading) return;
    state.frameLoading = true;
    state.currentTimestamp = timestamp;
    qs("#timestampInput").value = formatTimestamp(timestamp);

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
    };
    img.onerror = function () {
      state.frameLoading = false;
    };
    img.src = frameUrl(state.selectedParticipant, timestamp);
  }

  function initFrameControls() {
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

  function canvasCoords(canvas, event) {
    var rect = canvas.getBoundingClientRect();
    var scaleX = canvas.width / rect.width;
    var scaleY = canvas.height / rect.height;
    return {
      x: Math.round((event.clientX - rect.left) * scaleX),
      y: Math.round((event.clientY - rect.top) * scaleY),
    };
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
    var names = Object.keys(state.regions);
    var handleSize = Math.round(14 * s);
    for (var i = names.length - 1; i >= 0; i--) {
      var name = names[i];
      var r = regionToPixels(state.regions[name]);
      if (px >= r.x + r.w - handleSize && px <= r.x + r.w && py >= r.y + r.h - handleSize && py <= r.y + r.h) {
        return { name: name, handle: "resize" };
      }
      var lr = computeLabelRect(r, name, ctx, s);
      if (px >= lr.x && px <= lr.x + lr.w && py >= lr.y && py <= lr.y + lr.h) {
        return { name: name, handle: "move" };
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
      var pos = canvasCoords(overlay, e);
      var displayW = overlay.getBoundingClientRect().width || overlay.width;
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
      var pos = canvasCoords(overlay, e);
      var displayW = overlay.getBoundingClientRect().width || overlay.width;
      var s = overlay.width / displayW;
      if (state.resizingRegion) {
        var rName = state.resizingRegion.name;
        var rPx = regionToPixels(state.regions[rName]);
        var minSize = Math.round(20 * s);
        var newW = clamp(pos.x - rPx.x, minSize, overlay.width - rPx.x);
        var newH = clamp(pos.y - rPx.y, minSize, overlay.height - rPx.y);
        state.regions[rName] = Object.assign({}, state.regions[rName], { w: newW / overlay.width, h: newH / overlay.height });
        renderOverlay();
        return;
      }
      if (state.draggingRegion) {
        var d = state.draggingRegion;
        var dPx = regionToPixels(state.regions[d.name]);
        var newX = clamp(pos.x - d.offsetX, 0, overlay.width - dPx.w);
        var newY = clamp(pos.y - d.offsetY, 0, overlay.height - dPx.h);
        state.regions[d.name] = Object.assign({}, state.regions[d.name], { x: newX / overlay.width, y: newY / overlay.height });
        renderOverlay();
        return;
      }
      if (state.drawingRegion) {
        state.drawingRegion.endX = pos.x;
        state.drawingRegion.endY = pos.y;
        renderOverlay();
        return;
      }
      var ctx = overlay.getContext("2d");
      var hit = findHitRegion(pos.x, pos.y, s, ctx);
      state.hoveredRegion = hit;
      overlay.style.cursor = hit ? (hit.handle === "resize" ? "nwse-resize" : "grab") : "crosshair";
      renderOverlay();
    });

    overlay.addEventListener("mouseup", function (e) {
      if (state.resizingRegion) {
        var rName = state.resizingRegion.name;
        state.resizingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(rName);
        renderOverlay();
        updateRegionButtons();
        return;
      }
      if (state.draggingRegion) {
        var dName = state.draggingRegion.name;
        state.draggingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(dName);
        renderOverlay();
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
      renderOverlay();
      updateRegionButtons();
    });

    // Document-level listeners so drag/resize continues outside the canvas
    document.addEventListener("mousemove", function (e) {
      if (!state.resizingRegion && !state.draggingRegion) return;
      var pos = canvasCoords(overlay, e);
      var displayW = overlay.getBoundingClientRect().width || overlay.width;
      var s = overlay.width / displayW;
      if (state.resizingRegion) {
        var rName = state.resizingRegion.name;
        var rPx = regionToPixels(state.regions[rName]);
        var minSize = Math.round(20 * s);
        var newW = clamp(pos.x - rPx.x, minSize, overlay.width - rPx.x);
        var newH = clamp(pos.y - rPx.y, minSize, overlay.height - rPx.y);
        state.regions[rName] = Object.assign({}, state.regions[rName], { w: newW / overlay.width, h: newH / overlay.height });
        renderOverlay();
      } else if (state.draggingRegion) {
        var d = state.draggingRegion;
        var dPx = regionToPixels(state.regions[d.name]);
        var newX = clamp(pos.x - d.offsetX, 0, overlay.width - dPx.w);
        var newY = clamp(pos.y - d.offsetY, 0, overlay.height - dPx.h);
        state.regions[d.name] = Object.assign({}, state.regions[d.name], { x: newX / overlay.width, y: newY / overlay.height });
        renderOverlay();
      }
    });

    document.addEventListener("mouseup", function () {
      if (state.resizingRegion) {
        var rName = state.resizingRegion.name;
        state.resizingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(rName);
        renderOverlay();
        updateRegionButtons();
      } else if (state.draggingRegion) {
        var dName = state.draggingRegion.name;
        state.draggingRegion = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        saveRegionUpdate(dName);
        renderOverlay();
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
    qs("#saveRegionBtn").classList.toggle("hidden", !hasPending);
    qs("#clearSelectionBtn").classList.toggle("hidden", !hasPending && !hasActive);
    qs("#deleteRegionBtn").classList.toggle("hidden", !hasActive);
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

    // Draw saved regions
    var names = Object.keys(state.regions);
    names.forEach(function (name, i) {
      var r = regionToPixels(state.regions[name]);
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
        ctx.font = Math.round(11 * s) + "px " + getComputedStyle(document.documentElement).getPropertyValue("--font-mono").trim();
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
      ctx.font = Math.round(11 * s) + "px " + getComputedStyle(document.documentElement).getPropertyValue("--font-mono").trim();
      ctx.fillStyle = "rgba(255,255,255,0.9)";
      ctx.fillText(p.w + "\u00d7" + p.h + " px", p.x + Math.round(4 * s), p.y + p.h + Math.round(14 * s));
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
    var canvas = qs("#timelineCanvas");
    sizeTimelineCanvas();
    window.addEventListener("resize", sizeTimelineCanvas);

    canvas.addEventListener("click", function (e) {
      if (state.timelineDragging) return;
      var ts = timelineXToTime(e);
      if (ts !== null) loadFrame(ts);
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
    canvas.addEventListener("mousedown", function (e) {
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
        if (ts !== null) loadFrame(ts);
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
      scrubbing = false;
      if (dragStart) {
        setTimeout(function () { state.timelineDragging = false; }, 50);
        dragStart = null;
      }
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
    canvas.height = 64;
    renderTimeline();
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

    // Resolve CSS variables for theming
    var cs = getComputedStyle(document.documentElement);
    var surfaceAlt = cs.getPropertyValue("--color-surface-alt").trim() || "#f1ece4";
    var borderColor = cs.getPropertyValue("--color-border").trim() || "#e0ddd7";
    var textDim = cs.getPropertyValue("--color-text-dim").trim() || "#6b7280";
    var accent = cs.getPropertyValue("--color-accent").trim() || "#1d4f72";

    ctx.clearRect(0, 0, w, h);

    // Background
    ctx.fillStyle = surfaceAlt;
    ctx.fillRect(0, 0, w, h);

    if (dur <= 0) {
      ctx.fillStyle = textDim;
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
    ctx.strokeStyle = borderColor;
    ctx.fillStyle = textDim;
    ctx.font = "10px " + (cs.getPropertyValue("--font-mono").trim() || "monospace");
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

    // Result markers from completed tasks
    var resultY = 24;
    var resultH = h - resultY - 6;
    var focused = focusedTaskId();
    state.tasks.forEach(function (task) {
      if (task.status !== "completed" || !task.result) return;
      var color = taskTypeColor(task.type);
      var dimmed = focused && task.id !== focused;
      if (task.type === "color") {
        // Spans
        ctx.fillStyle = hexToRgba(color, dimmed ? 0.10 : 0.35);
        task.result.forEach(function (span) {
          var x1 = timeToX(span.start);
          var x2 = timeToX(span.end);
          ctx.fillRect(x1, resultY, Math.max(x2 - x1, 2), resultH);
        });
      } else if (task.type === "timelapse") {
        // No timeline markers for timelapse
      } else {
        // Point markers (change, similarity, text)
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
        });
      }
    });

    // Playhead
    var px = timeToX(state.currentTimestamp);
    ctx.strokeStyle = accent;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(px, 0);
    ctx.lineTo(px, h);
    ctx.stroke();
    // Playhead triangle
    ctx.fillStyle = accent;
    ctx.beginPath();
    ctx.moveTo(px - 5, 0);
    ctx.lineTo(px + 5, 0);
    ctx.lineTo(px, 6);
    ctx.closePath();
    ctx.fill();

    renderTimelineLegend();
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
      if (t.status === "completed" && t.result) hasTypes[t.type] = true;
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

  function renderWorkflowParams() {
    var container = qs("#workflowParams");
    container.innerHTML = "";
    var type = state.activeWorkflow;

    if (type === "color") {
      addParamRow(container, "Target Hue", rangeInput("paramColorH", 0, 180, 90), "paramColorHVal");
      addParamRow(container, "Target Sat", rangeInput("paramColorS", 0, 255, 200), "paramColorSVal");
      addParamRow(container, "Target Val", rangeInput("paramColorV", 0, 255, 200), "paramColorVVal");
      addParamRow(container, "Hue Tol.", rangeInput("paramColorTolH", 0, 90, 15), "paramColorTolHVal");
      addParamRow(container, "Sat Tol.", rangeInput("paramColorTolS", 0, 128, 40), "paramColorTolSVal");
      addParamRow(container, "Val Tol.", rangeInput("paramColorTolV", 0, 128, 40), "paramColorTolVVal");
      addParamRow(container, "Interval (s)", numberInput("paramColorInterval", 0.5, 60, 1.0, 0.5));
      // Color preview swatch
      var preview = el("div", "color-preview");
      preview.id = "colorPreview";
      container.appendChild(preview);
      updateColorPreview();
      // Sample from frame button
      var sampleBtn = el("button", "btn btn-small", "Sample from Frame");
      sampleBtn.addEventListener("click", sampleColorFromFrame);
      container.appendChild(sampleBtn);
    } else if (type === "change") {
      addParamRow(container, "Threshold", rangeInput("paramChangeThresh", 0.01, 0.50, 0.03, 0.01), "paramChangeThreshVal");
      addParamRow(container, "Noise Thr.", rangeInput("paramChangeNoise", 0, 100, 30, 1), "paramChangeNoiseVal");
      addParamRow(container, "Interval (s)", numberInput("paramChangeInterval", 0.5, 60, 1.0, 0.5));
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
      addParamRow(container, "Interval (s)", numberInput("paramSimInterval", 0.5, 60, 1.0, 0.5));
    } else if (type === "text") {
      addParamRow(container, "Search text", textInput("paramTextSearch", "Enter text to find..."));
      addParamRow(container, "Fuzzy Thr.", rangeInput("paramTextFuzzy", 0.50, 1.00, 0.80, 0.01), "paramTextFuzzyVal");
      addParamRow(container, "Interval (s)", numberInput("paramTextInterval", 0.5, 60, 2.0, 0.5));
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
      addParamRow(container, "Interval (s)", numberInput("paramNumInterval", 0.5, 60, 2.0, 0.5));
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

  function updateColorPreview() {
    var preview = qs("#colorPreview");
    if (!preview) return;
    var h = parseFloat((qs("#paramColorH") || {}).value) || 0;
    var s = parseFloat((qs("#paramColorS") || {}).value) || 0;
    var v = parseFloat((qs("#paramColorV") || {}).value) || 0;
    // Convert OpenCV HSV (H:0-180, S:0-255, V:0-255) to CSS hsl
    var hDeg = h * 2; // OpenCV H is 0-180, CSS is 0-360
    var sPercent = (s / 255) * 100;
    var lPercent = (v / 255) * (1 - s / 255 / 2) * 100;
    preview.style.background = "hsl(" + hDeg + "," + sPercent + "%," + Math.max(lPercent, 5) + "%)";
  }

  function sampleColorFromFrame() {
    if (!state.frameImage || !state.activeRegion) {
      showToast("Select a saved region first");
      return;
    }
    var r = regionToPixels(state.regions[state.activeRegion]);
    var canvas = qs("#frameCanvas");
    var ctx = canvas.getContext("2d");
    var imgData = ctx.getImageData(r.x, r.y, r.w, r.h);
    var data = imgData.data;
    // Average RGB then convert to HSV
    var totalR = 0, totalG = 0, totalB = 0;
    var count = data.length / 4;
    for (var i = 0; i < data.length; i += 4) {
      totalR += data[i];
      totalG += data[i + 1];
      totalB += data[i + 2];
    }
    var avgR = totalR / count / 255;
    var avgG = totalG / count / 255;
    var avgB = totalB / count / 255;
    var max = Math.max(avgR, avgG, avgB);
    var min = Math.min(avgR, avgG, avgB);
    var d = max - min;
    var hue = 0;
    if (d > 0) {
      if (max === avgR) hue = ((avgG - avgB) / d) % 6;
      else if (max === avgG) hue = (avgB - avgR) / d + 2;
      else hue = (avgR - avgG) / d + 4;
      hue = Math.round(hue * 30); // Convert to 0-180 range (OpenCV scale)
      if (hue < 0) hue += 180;
    }
    var sat = max > 0 ? Math.round((d / max) * 255) : 0;
    var val = Math.round(max * 255);

    var hEl = qs("#paramColorH");
    var sEl = qs("#paramColorS");
    var vEl = qs("#paramColorV");
    if (hEl) { hEl.value = hue; qs("#paramColorHVal").textContent = hue; }
    if (sEl) { sEl.value = sat; qs("#paramColorSVal").textContent = sat; }
    if (vEl) { vEl.value = val; qs("#paramColorVVal").textContent = val; }
    updateColorPreview();
    showToast("Sampled color from " + state.activeRegion);
  }

  function updateRunButton() {
    var btn = qs("#runBtn");
    var hasRegion = !!state.activeRegion;
    var hasParticipant = !!state.selectedParticipant;
    btn.disabled = !hasRegion || !hasParticipant;
    if (!hasRegion) {
      btn.setAttribute("data-tooltip", "Select a region first");
    } else if (!hasParticipant) {
      btn.setAttribute("data-tooltip", "Select a participant first");
    } else {
      btn.removeAttribute("data-tooltip");
    }
  }

  // ---- Run analysis ----

  function initRunButton() {
    qs("#runBtn").addEventListener("click", function () {
      if (!state.activeRegion || !state.selectedParticipant) return;
      var type = state.activeWorkflow;
      var params = gatherWorkflowParams(type);
      if (params === null) return;

      if (state.inMarker !== null) params.start_seconds = state.inMarker;
      if (state.outMarker !== null) params.end_seconds = state.outMarker;

      var body = {
        type: type,
        participant: state.selectedParticipant,
        region: state.activeRegion,
        parameters: params,
      };

      apiPost("api/tasks", body)
        .then(function (data) {
          if (data.ok) {
            state.tasks.push(data.task);
            renderTaskList();
            showToast("Task queued: " + type);
            startPolling();
          } else {
            showToast(data.error || "Failed to create task");
          }
        })
        .catch(function (err) { showToast("Error: " + err.message); });
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
      params.tolerance = {
        h: parseFloat((qs("#paramColorTolH") || {}).value) || 15,
        s: parseFloat((qs("#paramColorTolS") || {}).value) || 40,
        v: parseFloat((qs("#paramColorTolV") || {}).value) || 40,
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
    }
    return params;
  }

  // ---- Task queue ----

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

      // Select completed/paused task to view results; click again to deselect
      var task = findTask(taskId);
      if (task && (task.status === "completed" || task.status === "paused")) {
        if (state.selectedTaskId === taskId) {
          state.selectedTaskId = null;
          state.selectedTaskResults = null;
          renderResults();
          renderTaskList();
          renderTimeline();
        } else {
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
      if (task && task.status === "completed") {
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
      setInputValue("#paramColorTolH", params.tolerance ? params.tolerance.h : 15);
      setInputValue("#paramColorTolS", params.tolerance ? params.tolerance.s : 40);
      setInputValue("#paramColorTolV", params.tolerance ? params.tolerance.v : 40);
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
    container.innerHTML = "";
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
      badge.textContent = task.type;
      badge.style.background = taskTypeColor(task.type);
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
      if (task.status === "running") statusText = Math.round((task.progress || 0) * 100) + "%";
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

      container.appendChild(card);
    });
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
        renderTaskList();
        renderTimeline();
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
      var row = e.target.closest(".result-row");
      if (!row || !row.dataset.timestamp) return;
      var ts = parseFloat(row.dataset.timestamp);
      if (!isNaN(ts)) loadFrame(ts);
    });

    qs("#exportJsonBtn").addEventListener("click", function () { exportResults("json"); });
    qs("#exportCsvBtn").addEventListener("click", function () { exportResults("csv"); });
  }

  function loadAndShowResults(taskId) {
    apiGet("api/tasks/" + taskId + "/results")
      .then(function (data) {
        state.selectedTaskId = taskId;
        state.selectedTaskResults = data.results;
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
      container.innerHTML = '<div class="panel-empty">Click a completed task to view results.</div>';
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

    countEl.textContent = "(" + results.length + ")";
    container.innerHTML = "";

    results.forEach(function (r) {
      var row = el("div", "result-row");

      if (task.type === "color") {
        // Span result: start-end, duration
        row.dataset.timestamp = r.start;
        row.appendChild(el("span", "result-timestamp", formatTimestamp(r.start) + " \u2013 " + formatTimestamp(r.end)));
        row.appendChild(el("span", "result-detail", r.duration.toFixed(1) + "s"));
      } else if (task.type === "change") {
        row.dataset.timestamp = r.timestamp;
        row.appendChild(el("span", "result-timestamp", formatTimestamp(r.timestamp)));
        // Magnitude bar
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
      if (pid) selectParticipant(pid);
    });

    // Load initial data
    apiGet("api/participants")
      .then(function (data) {
        if (!data.ok) return;
        state.participants = (data.participants || []).filter(function (p) { return p.has_video; });
        renderParticipantSelect();
        if (state.participants.length > 0) {
          selectParticipant(state.participants[0].id);
        }
      })
      .catch(function () { showToast("Failed to load participants"); });

    apiGet("api/regions")
      .then(function (data) {
        if (data.ok) {
          state.regions = data.regions || {};
          renderRegionChips();
          renderOverlay();
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
