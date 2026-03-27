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
    timelineZoom: 1,
    timelineOffset: 0,
    inMarker: null,
    outMarker: null,
    activeWorkflow: "color",
    referenceTimestamp: null,
    tasks: [],
    selectedTaskId: null,
    selectedTaskResults: null,
    pollTimer: null,
    timelineDragging: false,
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
        qs("#timestampSlider").max = data.info.duration || 100;
        qs("#timestampSlider").value = 0;
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
    qs("#timestampSlider").value = timestamp;
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
    var slider = qs("#timestampSlider");
    var input = qs("#timestampInput");

    slider.addEventListener("input", function () {
      var ts = parseFloat(slider.value);
      if (!isNaN(ts)) loadFrame(ts);
    });

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

  function initRegionDrawing() {
    var overlay = qs("#overlayCanvas");

    overlay.addEventListener("mousedown", function (e) {
      if (e.button !== 0) return;
      var pos = canvasCoords(overlay, e);
      state.drawingRegion = { startX: pos.x, startY: pos.y, endX: pos.x, endY: pos.y };
      state.pendingRegion = null;
      state.activeRegion = null;
      updateRegionButtons();
    });

    overlay.addEventListener("mousemove", function (e) {
      if (!state.drawingRegion) return;
      var pos = canvasCoords(overlay, e);
      state.drawingRegion.endX = pos.x;
      state.drawingRegion.endY = pos.y;
      renderOverlay();
    });

    overlay.addEventListener("mouseup", function (e) {
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
      var body = { name: name, x: r.x, y: r.y, w: r.w, h: r.h };
      if (desc) body.description = desc;
      apiPost("api/regions", body)
        .then(function (data) {
          if (data.ok) {
            state.regions[name] = { x: r.x, y: r.y, w: r.w, h: r.h, description: desc };
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

    // Draw saved regions
    var names = Object.keys(state.regions);
    names.forEach(function (name, i) {
      var r = state.regions[name];
      var color = regionColorForIndex(i);
      var isActive = (name === state.activeRegion);
      ctx.strokeStyle = color;
      ctx.lineWidth = isActive ? 3 : 1.5;
      ctx.setLineDash(isActive ? [] : [8, 4]);
      ctx.strokeRect(r.x, r.y, r.w, r.h);
      if (isActive) {
        ctx.fillStyle = hexToRgba(color, 0.12);
        ctx.fillRect(r.x, r.y, r.w, r.h);
      }
      ctx.setLineDash([]);
      // Label
      ctx.font = "bold 13px -apple-system, BlinkMacSystemFont, sans-serif";
      var labelW = ctx.measureText(name).width + 8;
      ctx.fillStyle = hexToRgba(color, 0.85);
      ctx.fillRect(r.x, r.y - 18, labelW, 18);
      ctx.fillStyle = "#fff";
      ctx.fillText(name, r.x + 4, r.y - 5);
    });

    // Drawing in progress
    if (state.drawingRegion) {
      var d = state.drawingRegion;
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = 2;
      ctx.setLineDash([4, 4]);
      ctx.strokeRect(d.startX, d.startY, d.endX - d.startX, d.endY - d.startY);
      ctx.setLineDash([]);
      // Dimensions
      var w = Math.abs(d.endX - d.startX);
      var h = Math.abs(d.endY - d.startY);
      if (w > 20 && h > 20) {
        ctx.font = "12px " + getComputedStyle(document.documentElement).getPropertyValue("--font-mono").trim();
        ctx.fillStyle = "rgba(255,255,255,0.9)";
        ctx.fillText(w + "\u00d7" + h, Math.min(d.startX, d.endX) + 4, Math.max(d.startY, d.endY) + 16);
      }
    }

    // Pending (unsaved) region
    if (state.pendingRegion) {
      var p = state.pendingRegion;
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = 2;
      ctx.setLineDash([]);
      ctx.strokeRect(p.x, p.y, p.w, p.h);
      ctx.fillStyle = "rgba(255, 255, 255, 0.08)";
      ctx.fillRect(p.x, p.y, p.w, p.h);
      ctx.font = "12px " + getComputedStyle(document.documentElement).getPropertyValue("--font-mono").trim();
      ctx.fillStyle = "rgba(255,255,255,0.9)";
      ctx.fillText(p.w + "\u00d7" + p.h + " px", p.x + 4, p.y + p.h + 16);
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
      // Keep mouse position stable
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
    canvas.addEventListener("mousedown", function (e) {
      if (state.timelineZoom <= 1) return;
      dragStart = { x: e.clientX, offset: state.timelineOffset };
      state.timelineDragging = false;
    });
    document.addEventListener("mousemove", function (e) {
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
    var rect = canvas.parentElement.getBoundingClientRect();
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
    state.tasks.forEach(function (task) {
      if (task.status !== "completed" || !task.result) return;
      var color = taskTypeColor(task.type);
      if (task.type === "color") {
        // Spans
        ctx.fillStyle = hexToRgba(color, 0.35);
        task.result.forEach(function (span) {
          var x1 = timeToX(span.start);
          var x2 = timeToX(span.end);
          ctx.fillRect(x1, resultY, Math.max(x2 - x1, 2), resultH);
        });
      } else if (task.type === "timelapse") {
        // No timeline markers for timelapse
      } else {
        // Point markers (change, similarity, text)
        ctx.strokeStyle = color;
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
    types.forEach(function (type) {
      var item = el("span", "legend-item");
      var dot = el("span", "legend-dot");
      dot.style.background = taskTypeColor(type);
      item.appendChild(dot);
      item.appendChild(document.createTextNode(type));
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
    var r = state.regions[state.activeRegion];
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
    } else if (type === "timelapse") {
      params.speedup_factor = parseFloat((qs("#paramTlSpeed") || {}).value) || 10;
      params.output_format = (qs("#paramTlFormat") || {}).value || "mp4";
    }
    return params;
  }

  // ---- Task queue ----

  function initTaskQueue() {
    // Click handler delegated on taskList
    qs("#taskList").addEventListener("click", function (e) {
      var card = e.target.closest(".task-card");
      if (!card) return;
      var taskId = card.dataset.taskId;
      // Cancel button
      if (e.target.closest(".task-card-cancel")) {
        apiDelete("api/tasks/" + taskId)
          .then(function (data) {
            if (data.ok) {
              var task = findTask(taskId);
              if (task) task.status = "cancelled";
              renderTaskList();
              showToast("Task cancelled");
            }
          })
          .catch(function () { showToast("Failed to cancel task"); });
        return;
      }
      // Select completed task to view results
      var task = findTask(taskId);
      if (task && task.status === "completed") {
        state.selectedTaskId = taskId;
        loadAndShowResults(taskId);
        renderTaskList();
      }
    });
  }

  function findTask(id) {
    for (var i = 0; i < state.tasks.length; i++) {
      if (state.tasks[i].id === id) return state.tasks[i];
    }
    return null;
  }

  function renderTaskList() {
    var container = qs("#taskList");
    var count = qs("#taskCount");
    count.textContent = "(" + state.tasks.length + ")";
    if (state.tasks.length === 0) {
      container.innerHTML = '<div class="panel-empty">No tasks yet. Configure a workflow and click Run.</div>';
      return;
    }
    container.innerHTML = "";
    state.tasks.forEach(function (task) {
      var card = el("div", "task-card task-card-" + task.status);
      card.dataset.taskId = task.id;
      if (task.id === state.selectedTaskId) card.classList.add("selected");

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

      if (task.status === "running") {
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
      if (task.status === "failed" && task.error) {
        statusText = task.error;
        card.title = task.error;
      }
      if (task.status === "completed" && task.result) {
        var rLen = Array.isArray(task.result) ? task.result.length : (typeof task.result === "string" ? 1 : 0);
        statusText = rLen + " result" + (rLen !== 1 ? "s" : "");
      }
      card.appendChild(el("span", "task-card-status", statusText));

      // Cancel button
      if (task.status === "queued" || task.status === "running") {
        var cancelBtn = el("button", "task-card-cancel", "\u2715");
        cancelBtn.title = "Cancel";
        card.appendChild(cancelBtn);
      }

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
      return t.status === "queued" || t.status === "running";
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
        if (state.pendingRegion || state.activeRegion) {
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

  // ---- Init ----

  document.addEventListener("DOMContentLoaded", function () {
    initThemeToggle();
    initFrameControls();
    initRegionDrawing();
    initTimeline();
    initWorkflowTabs();
    initRunButton();
    initTaskQueue();
    initResultsPanel();
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
