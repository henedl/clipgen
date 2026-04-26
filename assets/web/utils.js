/* clipgen shared utilities – utils.js
 *
 * Common helpers extracted from individual page scripts.
 * Loaded before the page-specific JS in both Flask-served and
 * inlined/exported viewers.
 *
 * All declarations are global vars (no IIFE) so page scripts
 * running inside their own IIFEs can access them via scope chain.
 */

// ---- Feature flags ----

// Animated canvas dot-grid background (grid-bg.js). When false, falls back to
// the static CSS line-grid defined in tokens.css.
var CLIPGEN_ANIMATED_BG = true;

// ---- DOM helpers ----

var qs = function (sel) { return document.querySelector(sel); };
var qsa = function (sel) { return document.querySelectorAll(sel); };

var el = function (tag, cls, text) {
  var e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};

// True when an animated artifact's filename should render via <video> rather
// than <img>. Used by the gallery, viewer, and insights builder so they all
// agree which extensions are looping video. Keep as the single source of truth.
var isVideoLoop = function (filename) {
  return /\.webm$/i.test(filename || "");
};

// Create a looping, silent, autoplay <video> element for animated artifacts
// stored as .webm. Centralized so the attribute set is consistent everywhere.
var createLoopVideo = function (src, alt) {
  var v = document.createElement("video");
  v.src = src;
  v.autoplay = true;
  v.loop = true;
  v.muted = true;
  v.setAttribute("muted", "");
  v.setAttribute("playsinline", "");
  v.preload = "auto";
  if (alt) v.setAttribute("aria-label", alt);
  return v;
};

// ---- Hover tooltips (dark pill, pairs with .cg-tooltip in tokens.css) ----

var createTooltip = function (opts) {
  opts = opts || {};
  var cls = "cg-tooltip";
  if (opts.multiline) cls += " cg-tooltip--multiline";
  if (opts.align === "center") cls += " cg-tooltip--centered";
  var tip = document.createElement("div");
  tip.className = cls;
  document.body.appendChild(tip);
  return {
    el: tip,
    show: function (anchor, text) {
      tip.textContent = text || "";
      var r = anchor.getBoundingClientRect();
      // Set text before measuring so offsetHeight reflects final size.
      var x = opts.align === "center" ? r.left + r.width / 2 : r.left;
      var above = r.top - tip.offsetHeight - 6;
      tip.style.left = x + "px";
      tip.style.top = (above < 0 ? r.bottom + 6 : above) + "px";
      tip.classList.add("is-visible");
    },
    hide: function () {
      tip.classList.remove("is-visible");
    },
  };
};

var attachHoverTooltip = function (anchor, getText, opts) {
  var t = createTooltip(opts);
  var show = function () {
    var text = typeof getText === "function" ? getText() : getText;
    if (text) t.show(anchor, text);
  };
  anchor.addEventListener("mouseenter", show);
  anchor.addEventListener("focus", show);
  anchor.addEventListener("mouseleave", t.hide);
  anchor.addEventListener("blur", t.hide);
  return t;
};

// ---- Formatting ----

var pad2 = function (n) { return n < 10 ? "0" + n : "" + n; };

var formatTime = function (sec) {
  if (sec == null || isNaN(sec)) return "--:--";
  var total = Math.floor(sec);
  var h = Math.floor(total / 3600);
  var m = Math.floor((total % 3600) / 60);
  var s = total % 60;
  if (h > 0) return h + ":" + pad2(m) + ":" + pad2(s);
  return m + ":" + pad2(s);
};

var truncate = function (str, max) {
  if (!str) return "";
  return str.length > max ? str.slice(0, max) + "\u2026" : str;
};

// Parse "HH:MM:SS", "MM:SS", or a bare float to seconds. Returns null on failure.
var parseTimestamp = function (str) {
  str = (str == null ? "" : String(str)).trim();
  if (!str) return null;
  var parts = str.split(":");
  if (parts.length === 3) {
    var h = parseFloat(parts[0]), m = parseFloat(parts[1]), s = parseFloat(parts[2]);
    if (isNaN(h) || isNaN(m) || isNaN(s)) return null;
    return h * 3600 + m * 60 + s;
  }
  if (parts.length === 2) {
    var m2 = parseFloat(parts[0]), s2 = parseFloat(parts[1]);
    if (isNaN(m2) || isNaN(s2)) return null;
    return m2 * 60 + s2;
  }
  var n = parseFloat(str);
  return isNaN(n) ? null : n;
};

// Parse a timestamp string in clock semantics: 2-part is HH:MM (not MM:SS).
// Mirrors Python utils._clock_to_seconds. Returns null on failure.
var parseClockTimestamp = function (str) {
  str = (str == null ? "" : String(str)).trim();
  if (!str) return null;
  var parts = str.split(":");
  if (parts.length === 3) {
    var h = parseFloat(parts[0]), m = parseFloat(parts[1]), s = parseFloat(parts[2]);
    if (isNaN(h) || isNaN(m) || isNaN(s)) return null;
    return h * 3600 + m * 60 + s;
  }
  if (parts.length === 2) {
    var h2 = parseFloat(parts[0]), m2 = parseFloat(parts[1]);
    if (isNaN(h2) || isNaN(m2)) return null;
    return h2 * 3600 + m2 * 60;
  }
  return null;
};

// Parse a Sheet cell's timestamp tokens into [{startSeconds, duration}].
// When baselineSeconds > 0, tokens are treated as absolute clock times
// (2-part = HH:MM) and the baseline is subtracted from both ends of each
// range. Mirrors files.prepare_clip + utils.convert_clock_pairs_to_relative.
// Pairs that resolve to negative or zero-length intervals are skipped.
var parseClipSegmentsForCell = function (raw, baselineSeconds, defaultDuration) {
  var DEFAULT_DUR = defaultDuration || 60;
  var hasBaseline = baselineSeconds && baselineSeconds > 0;
  var tsParse = hasBaseline ? parseClockTimestamp : parseTimestamp;
  var cleaned = String(raw || "").toLowerCase().replace(/!key/g, "").replace(/[+;,]/g, " ");
  var tokens = cleaned.split(/\s+/).filter(function (t) { return t && t !== "x"; });
  var segments = [];
  for (var i = 0; i < tokens.length; i++) {
    var tok = tokens[i].replace(/\.$/, "").replace(/\./g, ":");
    var dashIdx = -1;
    for (var d = 1; d < tok.length; d++) {
      if (tok[d] === "-" && tok[d - 1] >= "0" && tok[d - 1] <= "9") { dashIdx = d; break; }
    }
    if (dashIdx > 0) {
      var s = tsParse(tok.substring(0, dashIdx));
      var e = tsParse(tok.substring(dashIdx + 1));
      if (s === null || e === null) continue;
      if (hasBaseline) {
        s -= baselineSeconds;
        e -= baselineSeconds;
        if (s < 0 || e <= 0 || e <= s) continue;
      }
      segments.push({ startSeconds: Math.floor(s), duration: Math.max(0, e - s) });
    } else if (tok.indexOf(":") > 0) {
      var sec = tsParse(tok);
      if (sec === null) continue;
      if (hasBaseline) {
        sec -= baselineSeconds;
        if (sec < 0) continue;
      }
      segments.push({ startSeconds: Math.floor(sec), duration: DEFAULT_DUR });
    }
  }
  return segments;
};

// ---- Math ----

var clamp = function (val, min, max) {
  return Math.max(min, Math.min(max, val));
};

// ---- Tooltip positioning ----

// Position a tooltip element centered above an anchor rect, flipping below
// if there's no room above and clamping to the viewport horizontally.
var positionTooltipAnchored = function (tooltipEl, anchorRect) {
  var ttW = tooltipEl.offsetWidth;
  var ttH = tooltipEl.offsetHeight;
  var left = anchorRect.left + anchorRect.width / 2 - ttW / 2;
  var top = anchorRect.top - ttH - 6;
  if (top < 4) top = anchorRect.bottom + 6;
  if (left < 4) left = 4;
  if (left + ttW > window.innerWidth - 4) left = window.innerWidth - ttW - 4;
  tooltipEl.style.left = left + "px";
  tooltipEl.style.top = top + "px";
};

// ---- Debounce / escaping / color / toast (shared across Studio tabs + web UIs) ----

var debounce = function (fn, ms) {
  var timer;
  return function () {
    var ctx = this, args = arguments;
    clearTimeout(timer);
    timer = setTimeout(function () { fn.apply(ctx, args); }, ms);
  };
};

var escapeHtml = function (str) {
  var div = document.createElement("div");
  div.appendChild(document.createTextNode(str));
  return div.innerHTML;
};

var hexToRgba = function (hex, alpha) {
  var r = parseInt(hex.slice(1, 3), 16);
  var g = parseInt(hex.slice(3, 5), 16);
  var b = parseInt(hex.slice(5, 7), 16);
  return "rgba(" + r + "," + g + "," + b + "," + alpha + ")";
};

var SHOW_TOAST_DEFAULT_MS = 3000;

var showToast = function (msg, opts) {
  var durationMs = SHOW_TOAST_DEFAULT_MS;
  if (opts != null && opts.durationMs != null) durationMs = opts.durationMs;
  var toastEl = qs("#toast");
  if (!toastEl) return;
  toastEl.textContent = msg;
  toastEl.classList.remove("hidden");
  clearTimeout(toastEl._timer);
  toastEl._timer = setTimeout(function () {
    toastEl.classList.add("hidden");
  }, durationMs);
};

// ---- Severity ----

var severityClass = function (raw) {
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
};

// ---- API helpers (always check r.ok) ----

var apiGet = function (path) {
  return fetch(path).then(function (r) {
    if (!r.ok) throw new Error("Server error " + r.status);
    return r.json();
  });
};

var apiPost = function (path, body) {
  return fetch(path, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  }).then(function (r) {
    if (!r.ok) throw new Error("Server error " + r.status);
    return r.json();
  });
};

var apiPut = function (path, body) {
  return fetch(path, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  }).then(function (r) {
    if (!r.ok) throw new Error("Server error " + r.status);
    return r.json();
  });
};

var apiDelete = function (path) {
  return fetch(path, { method: "DELETE" }).then(function (r) {
    if (!r.ok) throw new Error("Server error " + r.status);
    return r.json();
  });
};

// ---- Polling ----

var POLL_INTERVAL = 3000;

// ---- Mark categories ----
// Hardcoded fallback that mirrors config.MARK_CATEGORIES defaults; the live
// values are repopulated in place by setMarkCategories() once the page fetches
// settings from the server. Existing references to MARK_CATEGORIES keep
// working because we mutate this object rather than replace it.
// Verified against config.MARK_CATEGORIES by tests/test_shared_constants.py.

var MARK_CATEGORIES = {
  pain_point: { label: "Pain Point", color: "#dc2626" },
  delight:    { label: "Delight",    color: "#16a34a" },
  quote:      { label: "Quote",      color: "#2563eb" },
  insight:    { label: "Insight",    color: "#f97316" },
  task:       { label: "Task Issue", color: "#8b5cf6" },
  bookmark:   { label: "Bookmark",   color: "#0891b2" },
};

function setMarkCategories(next) {
  if (!next || typeof next !== "object") return;
  for (var k in MARK_CATEGORIES) {
    if (Object.prototype.hasOwnProperty.call(MARK_CATEGORIES, k)) {
      delete MARK_CATEGORIES[k];
    }
  }
  for (var key in next) {
    if (!Object.prototype.hasOwnProperty.call(next, key)) continue;
    var entry = next[key];
    if (entry && typeof entry === "object") {
      MARK_CATEGORIES[key] = {
        label: entry.label || key,
        color: entry.color || "#888888",
      };
    }
  }
}

// ---- Cross-reference badge metadata ----
// Icon names reference files in assets/icons/; rendered via CSS mask-image.

var XREF_BADGES = {
  screenspace: { icon: "squares-2x2", color: "rgba(52, 152, 219, 0.85)" },
  transcript:  { icon: "chat-bubble-bottom-center-text", color: "rgba(16, 163, 74, 0.85)" },
  sheet:       { icon: "table-cells", color: "rgba(234, 179, 8, 0.85)" },
};

// ---- Detector colors ----
// Read from CSS custom properties (tokens.css). Mutable so the same object
// reference can be re-populated after a theme toggle without breaking callers.
// Hardcoded fallback supports exported viewers that may lack tokens.css.

var DETECTOR_COLORS = {};
var _DETECTOR_TYPES = [
  "multitool", "color", "change", "similarity", "text",
  "numbers", "timelapse", "template", "flow", "scene", "inactivity",
];
var _DETECTOR_FALLBACK = {
  multitool: "#2563eb", color: "#8b5cf6", change: "#f97316",
  similarity: "#0ea5e9", text: "#10b981", numbers: "#eab308",
  timelapse: "#ec4899", template: "#f43f5e", flow: "#6366f1",
  scene: "#14b8a6", inactivity: "#78716c",
};

function refreshDetectorColors() {
  try {
    var style = getComputedStyle(document.documentElement);
    _DETECTOR_TYPES.forEach(function (t) {
      var val = style.getPropertyValue("--color-task-" + t).trim();
      DETECTOR_COLORS[t] = val || _DETECTOR_FALLBACK[t] || "#888";
    });
  } catch (_) {
    _DETECTOR_TYPES.forEach(function (t) {
      DETECTOR_COLORS[t] = _DETECTOR_FALLBACK[t] || "#888";
    });
  }
}

refreshDetectorColors();

// ---- Shared settings (localStorage) ----

var THEME_STORAGE_KEY = "clipgen-theme";
var TOOLTIP_STORAGE_KEY = "clipgen-tooltips";

var applyStoredThemePreference = function () {
  var stored = null;
  try { stored = window.localStorage.getItem(THEME_STORAGE_KEY); } catch (_) {}
  var root = document.documentElement;
  if (stored === "light" || stored === "dark") {
    root.setAttribute("data-theme", stored);
  } else {
    root.removeAttribute("data-theme");
  }
  updateThemeToggleButton(stored);
};

var toggleThemePreference = function () {
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
  refreshDetectorColors();
};

var updateThemeToggleButton = function (explicitTheme) {
  var btn = qs("#themeToggle");
  if (!btn) return;
  var effective = explicitTheme;
  if (effective !== "light" && effective !== "dark") {
    var prefersDark = false;
    try {
      prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
    } catch (_) {}
    effective = prefersDark ? "dark" : "light";
  }
  btn.setAttribute("data-theme", effective);
  btn.setAttribute("aria-pressed", effective === "dark" ? "true" : "false");
};

var initThemeToggle = function (onToggle) {
  applyStoredThemePreference();
  var btn = qs("#themeToggle");
  if (!btn) return;
  btn.addEventListener("click", function () {
    toggleThemePreference();
    if (onToggle) onToggle();
  });
};

// ---- Frontend switcher (shared across Studio / Screenspace / Transcripts / Insights) ----

var initFrontendSwitcher = function () {
  var root = qs(".frontend-switcher");
  if (!root) return;
  var trigger = root.querySelector(".frontend-switcher-trigger");
  var panel = root.querySelector(".frontend-switcher-panel");
  if (!trigger || !panel) return;
  var closeTimer = null;

  function open() {
    clearTimeout(closeTimer);
    root.classList.add("open");
    trigger.setAttribute("aria-expanded", "true");
    panel.setAttribute("aria-hidden", "false");
  }
  function close() {
    root.classList.remove("open");
    trigger.setAttribute("aria-expanded", "false");
    panel.setAttribute("aria-hidden", "true");
  }
  function scheduleClose() {
    clearTimeout(closeTimer);
    closeTimer = setTimeout(close, 120);
  }

  root.addEventListener("mouseenter", open);
  root.addEventListener("mouseleave", scheduleClose);
  trigger.addEventListener("click", function (e) {
    e.preventDefault();
    if (root.classList.contains("open")) close();
    else open();
  });
  trigger.addEventListener("keydown", function (e) {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      open();
      var first = panel.querySelector(".frontend-switcher-item");
      if (first) first.focus();
    } else if (e.key === "Escape") {
      close();
    }
  });
  document.addEventListener("keydown", function (e) {
    if (e.key === "Escape" && root.classList.contains("open")) {
      close();
      trigger.focus();
    }
  });
  document.addEventListener("click", function (e) {
    if (!root.contains(e.target)) close();
  });

  fetch("/api/status")
    .then(function (r) { return r.json(); })
    .then(function (status) {
      var items = panel.querySelectorAll(".frontend-switcher-item");
      items.forEach(function (item) {
        var key = item.dataset.frontend;
        if (key && status[key] === false) item.classList.add("hidden");
      });
    })
    .catch(function () {});
};

var getStoredTooltipPref = function () {
  try {
    var v = window.localStorage.getItem(TOOLTIP_STORAGE_KEY);
    if (v === "false") return false;
  } catch (_) {}
  return true;
};

var setStoredTooltipPref = function (enabled) {
  try {
    window.localStorage.setItem(TOOLTIP_STORAGE_KEY, enabled ? "true" : "false");
  } catch (_) {}
};

// ---- Per-page UI state (localStorage) ----

var UI_STATE_STORAGE_KEY = "clipgen-ui-state";

var getStoredUIState = function (page) {
  try {
    var raw = window.localStorage.getItem(UI_STATE_STORAGE_KEY);
    if (!raw) return {};
    var all = JSON.parse(raw);
    return (all && typeof all[page] === "object" && all[page]) ? all[page] : {};
  } catch (_) { return {}; }
};

var setStoredUIStateField = function (page, field, value) {
  try {
    var raw = window.localStorage.getItem(UI_STATE_STORAGE_KEY);
    var all = {};
    if (raw) {
      try { all = JSON.parse(raw) || {}; } catch (_) { all = {}; }
    }
    if (!all[page] || typeof all[page] !== "object") all[page] = {};
    if (value === null || value === undefined) delete all[page][field];
    else all[page][field] = value;
    window.localStorage.setItem(UI_STATE_STORAGE_KEY, JSON.stringify(all));
  } catch (_) {}
};
