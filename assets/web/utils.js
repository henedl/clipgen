/* clipgen shared utilities – utils.js
 *
 * Common helpers extracted from individual page scripts.
 * Loaded before the page-specific JS in both Flask-served and
 * inlined/exported viewers.
 *
 * All declarations are global vars (no IIFE) so page scripts
 * running inside their own IIFEs can access them via scope chain.
 */

// ---- DOM helpers ----

var qs = function (sel) { return document.querySelector(sel); };
var qsa = function (sel) { return document.querySelectorAll(sel); };

var el = function (tag, cls, text) {
  var e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
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

// ---- Mark categories (mirrored from transcripts.py MARK_CATEGORIES) ----
// Kept in sync by tests/test_shared_constants.py — update both together.

var MARK_CATEGORIES = {
  pain_point: { label: "Pain Point", color: "#dc2626" },
  delight:    { label: "Delight",    color: "#16a34a" },
  quote:      { label: "Quote",      color: "#2563eb" },
  insight:    { label: "Insight",    color: "#f97316" },
  task:       { label: "Task Issue", color: "#8b5cf6" },
  bookmark:   { label: "Bookmark",   color: "#0891b2" },
};

// ---- Cross-reference badge metadata ----
// Icon names reference files in assets/icons/; rendered via CSS mask-image.

var XREF_BADGES = {
  screenspace: { icon: "squares-2x2", color: "rgba(52, 152, 219, 0.85)" },
  transcript:  { icon: "chat-bubble-bottom-center-text", color: "rgba(16, 163, 74, 0.85)" },
  sheet:       { icon: "table-cells", color: "rgba(234, 179, 8, 0.85)" },
};

// ---- Detector colors ----
// Read from CSS custom properties (tokens.css) with hardcoded fallback
// for exported viewers that may lack tokens.css at parse time.

var DETECTOR_COLORS = (function () {
  var types = [
    "multitool", "color", "change", "similarity", "text",
    "numbers", "timelapse", "template", "flow", "scene", "inactivity",
  ];
  var fallback = {
    multitool: "#2563eb", color: "#8b5cf6", change: "#f97316",
    similarity: "#0ea5e9", text: "#10b981", numbers: "#eab308",
    timelapse: "#ec4899", template: "#f43f5e", flow: "#6366f1",
    scene: "#14b8a6", inactivity: "#78716c",
  };
  try {
    var style = getComputedStyle(document.documentElement);
    var map = {};
    types.forEach(function (t) {
      var val = style.getPropertyValue("--color-task-" + t).trim();
      map[t] = val || fallback[t] || "#888";
    });
    return map;
  } catch (_) {
    return fallback;
  }
})();

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
