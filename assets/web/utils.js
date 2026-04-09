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
