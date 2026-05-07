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

// Live token-tweak debug widget (dev-token-tweak.js). When false, the widget
// script bails on load and never mounts. Flip to false during a build, or
// when not iterating on the redesign. The widget never ships in exports
// either way — viewer.py strips data-dev-only tags during inlining.
var CLIPGEN_DEV_TOKEN_TWEAK = false;

// ---- Canonical config (mirror of config.py via utils.get_frontend_config)
//
// Source of truth: every API response (server.py /api/sheet-data) and every
// exported viewer payload (viewer.py finalize_*) embeds a `config` field.
// Pages call
// clipgenApplyConfig(payload) to overlay the live values onto these defaults.
// The hardcoded defaults below cover purely-offline contexts (re-opened
// older exported viewers); tests/test_shared_constants.py asserts they
// match config.py.

var CLIPGEN_CONFIG = {
  defaultDuration: 60,
  severity: [
    { label: "Critical",      rank: -4, cssClass: "sev-critical" },
    { label: "High",          rank: -3, cssClass: "sev-high" },
    { label: "Medium",        rank: -2, cssClass: "sev-medium" },
    { label: "Low",           rank: -1, cssClass: "sev-low" },
    { label: "N/A",           rank:  0, cssClass: "sev-na" },
    { label: "Positive",      rank:  1, cssClass: "sev-positive" },
    { label: "Very Positive", rank:  2, cssClass: "sev-very-positive" },
  ],
  annotationKeyphrases: ["!key"],
  ignoredTimestampTokens: ["x"],
};

var clipgenApplyConfig = function (payload) {
  if (!payload || typeof payload !== "object") return;
  if (typeof payload.defaultDuration === "number") {
    CLIPGEN_CONFIG.defaultDuration = payload.defaultDuration;
  }
  if (Array.isArray(payload.severity) && payload.severity.length) {
    CLIPGEN_CONFIG.severity = payload.severity;
  }
  if (Array.isArray(payload.annotationKeyphrases)) {
    CLIPGEN_CONFIG.annotationKeyphrases = payload.annotationKeyphrases;
  }
  if (Array.isArray(payload.ignoredTimestampTokens)) {
    CLIPGEN_CONFIG.ignoredTimestampTokens = payload.ignoredTimestampTokens;
  }
};

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
// than <img>. Used by the gallery and viewer so they agree which extensions
// are looping video. Keep as the single source of truth.
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

// ---- Brand mark hydration ----
//
// Fetches assets/logos/favicon.svg and injects it as inline <svg> into every
// .brand-mark element so the three F-paths can be animated via CSS
// stroke-dashoffset (see .brand-mark.is-animated rules in tokens.css).
//
// The draw-on cascade plays only the first time a browser session sees the
// mark (sessionStorage key BRAND_MARK_PLAYED_KEY). Navigating between Studio,
// Screenspace, and Transcripts within the same tab re-injects the SVG but
// skips the animation — closing the tab/browser clears the flag and the
// next visit replays.
//
// On fetch failure (e.g. file:// open with no server, blocked by CSP) the
// existing mask-image fallback in tokens.css renders the static mark — no
// flicker, no broken state.
//
// Loading from the file rather than inlining means future drop-in replacement
// of assets/logos/favicon.svg propagates without code edits.
var BRAND_MARK_PLAYED_KEY = "clipgen.brand-mark.played";

var clipgenInitBrandMark = function () {
  var marks = document.querySelectorAll(".brand-mark");
  if (!marks.length) return;
  var played = false;
  try { played = window.sessionStorage.getItem(BRAND_MARK_PLAYED_KEY) === "1"; } catch (_) {}
  fetch("logos/favicon.svg")
    .then(function (r) { return r.ok ? r.text() : null; })
    .then(function (text) {
      if (!text) return;
      var doc = new DOMParser().parseFromString(text, "image/svg+xml");
      var src = doc.documentElement;
      if (!src || src.tagName.toLowerCase() !== "svg") return;
      var paths = src.querySelectorAll("path");
      if (paths.length !== 3) return;
      paths[0].setAttribute("class", "brand-mark__line brand-mark__line--1");
      paths[1].setAttribute("class", "brand-mark__line brand-mark__line--2");
      paths[2].setAttribute("class", "brand-mark__line brand-mark__line--3");
      src.setAttribute("aria-hidden", "true");
      src.setAttribute("focusable", "false");
      Array.prototype.forEach.call(marks, function (mark) {
        mark.replaceChildren(src.cloneNode(true));
        mark.classList.add("is-hydrated");
        if (!played) mark.classList.add("is-animated");
      });
      if (!played) {
        try { window.sessionStorage.setItem(BRAND_MARK_PLAYED_KEY, "1"); } catch (_) {}
      }
    })
    .catch(function () { /* fall back to mask-image in tokens.css */ });
};

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", clipgenInitBrandMark);
} else {
  clipgenInitBrandMark();
}

// ---- Mask-image icon helpers ----
//
// Apply both `mask-image` and `-webkit-mask-image` to a DOM element. `urlValue`
// is a CSS url(...) string such as 'url("icons/check.svg")' or
// 'url("/screenspace/icons/eye.svg")'. The element's `mask-size`,
// `mask-repeat`, and `background-color: currentColor` should come from a CSS
// class on the element (see .xref-badge-icon, .cg-icon, .ss-task-icon, etc).

var applyMaskIcon = function (el, urlValue) {
  el.style.maskImage = urlValue;
  el.style.webkitMaskImage = urlValue;
};

// Inline-style equivalent for embedding in HTML strings:
//   "mask-image: url(...); -webkit-mask-image: url(...);"
var maskIconStyle = function (urlValue) {
  return "mask-image:" + urlValue + ";-webkit-mask-image:" + urlValue + ";";
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

// `options.decimals` adds fractional-second precision (e.g. { decimals: 1 } → "m:ss.s").
var formatTime = function (sec, options) {
  options = options || {};
  if (sec == null || isNaN(sec) || !isFinite(sec)) return "--:--";
  if (sec < 0) sec = 0;
  var decimals = options.decimals || 0;
  var totalInt = Math.floor(sec);
  var h = Math.floor(totalInt / 3600);
  var m = Math.floor((totalInt % 3600) / 60);
  var s = totalInt % 60;
  var sStr;
  if (decimals > 0) {
    var sValue = s + (sec - totalInt);
    sStr = (sValue < 10 ? "0" : "") + sValue.toFixed(decimals);
  } else {
    sStr = pad2(s);
  }
  if (h > 0) return h + ":" + pad2(m) + ":" + sStr;
  return m + ":" + sStr;
};

// Like formatTime but rounds rather than floors. Used where the value is a
// duration (e.g. clip length, ruler ticks) and rounding to the nearest second
// reads more naturally than truncating.
var formatDuration = function (sec) {
  if (sec == null || isNaN(sec)) return "--:--";
  var total = Math.round(sec);
  var h = Math.floor(total / 3600);
  var m = Math.floor((total % 3600) / 60);
  var s = total % 60;
  if (h > 0) return h + ":" + pad2(m) + ":" + pad2(s);
  return m + ":" + pad2(s);
};

var artifactDurationSec = function (a) {
  var s = Number(a.start);
  var e = Number(a.end);
  if (isNaN(s)) s = 0;
  if (isNaN(e)) e = isNaN(s) ? 0 : s;
  var d = e - s;
  if (isNaN(d) || d < 0) return 0;
  return d;
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
// defaultDuration: required (per-clip duration when only a start time is
// given). Callers must pass CLIPGEN_CONFIG.defaultDuration; a missing value
// is a contract bug.
var parseClipSegmentsForCell = function (raw, baselineSeconds, defaultDuration) {
  var DEFAULT_DUR = defaultDuration;
  var hasBaseline = baselineSeconds && baselineSeconds > 0;
  var tsParse = hasBaseline ? parseClockTimestamp : parseTimestamp;
  var cleaned = String(raw || "").toLowerCase();
  for (var ki = 0; ki < CLIPGEN_CONFIG.annotationKeyphrases.length; ki++) {
    var phrase = CLIPGEN_CONFIG.annotationKeyphrases[ki];
    cleaned = cleaned.split(phrase).join("");
  }
  cleaned = cleaned.replace(/[+;,]/g, " ");
  var ignored = CLIPGEN_CONFIG.ignoredTimestampTokens;
  var tokens = cleaned.split(/\s+/).filter(function (t) {
    return t && ignored.indexOf(t) === -1;
  });
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

// Median of a numeric array. Returns 0 for empty input.
var median = function (arr) {
  if (!arr.length) return 0;
  var sorted = arr.slice().sort(function (a, b) { return a - b; });
  var mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
};

// Population standard deviation. Returns 0 for arrays shorter than 2.
var stddev = function (nums) {
  if (nums.length < 2) return 0;
  var sum = 0;
  for (var i = 0; i < nums.length; i++) sum += nums[i];
  var mean = sum / nums.length;
  var sq = 0;
  for (var j = 0; j < nums.length; j++) sq += (nums[j] - mean) * (nums[j] - mean);
  return Math.sqrt(sq / nums.length);
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

// Position a popover element bottom-left aligned with an anchor rect, flipping
// above if there's no room below, and clamping to the viewport with 4px
// margins. Reads the popover's actual offsetWidth/Height — the popover must
// be in the DOM and visible (not display:none) before calling.
var positionPopoverAnchored = function (popoverEl, anchorRect) {
  var w = popoverEl.offsetWidth;
  var h = popoverEl.offsetHeight;
  var top = anchorRect.bottom + 4;
  if (top + h > window.innerHeight - 4) top = anchorRect.top - h - 4;
  if (top < 4) top = 4;
  var left = anchorRect.left;
  if (left + w > window.innerWidth - 4) left = window.innerWidth - w - 4;
  if (left < 4) left = 4;
  popoverEl.style.left = left + "px";
  popoverEl.style.top = top + "px";
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

// Read a CSS custom property from :root, returning fallback when unset/empty.
// Pages use this for theme-aware values (--color-accent, --color-heatmap, ...).
var getCSSVar = function (name, fallback) {
  try {
    var v = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    return v || fallback;
  } catch (_) {
    return fallback;
  }
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
//
// Read severity metadata from CLIPGEN_CONFIG.severity (kept in sync with
// config.py SEVERITY_NUMERIC_TO_LABEL via tests/test_shared_constants.py).

var severityClass = function (raw) {
  if (!raw || !String(raw).trim()) return "";
  var k = String(raw).trim().toLowerCase();
  for (var i = 0; i < CLIPGEN_CONFIG.severity.length; i++) {
    if (CLIPGEN_CONFIG.severity[i].label.toLowerCase() === k) {
      return CLIPGEN_CONFIG.severity[i].cssClass;
    }
  }
  return "sev-unknown";
};

// Numeric rank (most-severe = lowest, e.g. Critical = -4) used by sorters
// and Studio's severity filter. Returns null for empty or unrecognized input
// — callers decide how to treat unknown (Studio filters them out; viewer
// sorts them last).
var severityRank = function (raw) {
  if (!raw || !String(raw).trim()) return null;
  var k = String(raw).trim().toLowerCase();
  for (var i = 0; i < CLIPGEN_CONFIG.severity.length; i++) {
    if (CLIPGEN_CONFIG.severity[i].label.toLowerCase() === k) {
      return CLIPGEN_CONFIG.severity[i].rank;
    }
  }
  return null;
};

// Same as severityRank but keyed by CSS class (sev-critical, sev-high, ...).
// Useful when the caller has already resolved the class via severityClass().
// Returns null for unknown class so sort callers can fall back to 999.
var severityRankByClass = function (cssClass) {
  if (!cssClass) return null;
  for (var i = 0; i < CLIPGEN_CONFIG.severity.length; i++) {
    if (CLIPGEN_CONFIG.severity[i].cssClass === cssClass) {
      return CLIPGEN_CONFIG.severity[i].rank;
    }
  }
  return null;
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
// Colors use color-mix on the canonical `--stream-*` tokens in tokens.css so
// any theme change or token tweak propagates automatically (color-mix yields a
// CSS color string usable both in inline style and JS .style assignment).

var XREF_BADGES = {
  screenspace: { icon: "squares-2x2", color: "color-mix(in srgb, var(--stream-screenspace) 85%, transparent)" },
  transcript:  { icon: "chat-bubble-bottom-center-text", color: "color-mix(in srgb, var(--stream-transcript) 85%, transparent)" },
  sheet:       { icon: "table-cells", color: "color-mix(in srgb, var(--stream-sheet) 85%, transparent)" },
};

// ---- Filter helpers (artifact grids in viewer) ----

// Sorted unique non-empty values of `field` across `items`.
// opts.trim: trim string values before deduping (default false).
var uniqueFieldValues = function (items, field, opts) {
  var trim = opts && opts.trim;
  var seen = {};
  var out = [];
  for (var i = 0; i < items.length; i++) {
    var v = items[i][field];
    if (v == null) continue;
    if (trim) v = String(v).trim();
    if (v && !seen[v]) { seen[v] = true; out.push(v); }
  }
  return out.sort();
};

// Replace a <select>'s options with [allLabel, ...values]. The "all" option
// has empty value "" so callers can detect it via select.value === "".
var populateSelect = function (selectEl, values, allLabel) {
  if (!selectEl) return;
  selectEl.innerHTML = "";
  var frag = document.createDocumentFragment();
  var allOpt = document.createElement("option");
  allOpt.value = "";
  allOpt.textContent = allLabel;
  frag.appendChild(allOpt);
  for (var i = 0; i < values.length; i++) {
    var o = document.createElement("option");
    o.value = values[i];
    o.textContent = values[i];
    frag.appendChild(o);
  }
  selectEl.appendChild(frag);
};

// ---- Detector colors ----
//
// Single source of truth: `--color-task-{type}` tokens in `tokens.css`.
// Every place that paints a detector — Screenspace workflow tabs, Screenspace
// result rows, Studio Screenspace-Intake (filter chips / density bars / card
// labels), and exported viewers — pulls from those tokens via either CSS
// (`var(--color-task-X)` directly) or the JS helpers below.
//
// `_DETECTOR_FALLBACK` is the offline-export safety net for HTML files that
// somehow ship without `tokens.css`; values here MUST stay aligned with the
// dark-theme `--color-task-*` block in `tokens.css`. `CATEGORY_HUES` stays
// the path for non-detector labels (transcript intake categories, mark
// categories) — adding a new detector without updating tokens.css is caught
// by `tests/test_shared_constants.py`.

var DETECTOR_COLORS = {};
var _DETECTOR_TYPES = [
  "multitool", "color", "change", "similarity", "text",
  "numbers", "timelapse", "template", "flow", "scene", "inactivity",
];
// Values mirrored from the dark-theme `--color-task-*` block in tokens.css.
// Update this map and tokens.css together when changing a detector palette.
var _DETECTOR_FALLBACK = {
  multitool: "#60a5fa", color: "#a78bfa", change: "#fb923c",
  similarity: "#22d3ee", text: "#34d399", numbers: "#facc15",
  timelapse: "#f472b6", template: "#fb7185", flow: "#818cf8",
  scene: "#2dd4bf", inactivity: "#94a3b8",
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

// Return a CSS color string for a known Screenspace detector label, sourced
// from the canonical `--color-task-{type}` token in `tokens.css` so chips /
// dots / bars / labels in Studio's Screenspace Intake match the live
// Screenspace surface exactly. Returns `null` for unknown labels so callers
// can fall back to the oklch / `categoryHue` path for non-detector tints
// (transcript categories, mark categories, ad-hoc labels).
//
//   alpha == null | >= 1  →  raw `var(--color-task-X)`
//   alpha < 1             →  `color-mix(in oklch, var(--color-task-X) <pct>%, transparent)`
function detectorColor(label, alpha) {
  if (!label) return null;
  var key = String(label).toLowerCase().trim();
  if (_DETECTOR_TYPES.indexOf(key) === -1) return null;
  var v = "var(--color-task-" + key + ")";
  if (alpha == null || alpha >= 1) return v;
  var pct = Math.max(0, Math.min(100, Math.round(alpha * 100)));
  return "color-mix(in oklch, " + v + " " + pct + "%, transparent)";
}

// ---- Category hue palette (redesign) ----
// Stable hue per non-detector label (mark categories / transcript intake
// categories), rendered at runtime via oklch(). Detector colors do NOT live
// here — see `detectorColor()` above and the `--color-task-*` tokens for the
// canonical detector palette. The detector entries below remain so legacy
// callers that pass a detector label keep working, but new code colouring
// a detector should prefer `detectorColor(label)` over `categoryColor(label)`
// to stay aligned with Screenspace.
var CATEGORY_HUES = {
  multitool: 220, color: 280, change: 30, similarity: 200,
  text: 170, numbers: 330, timelapse: 350, template: 18,
  flow: 145, scene: 155, inactivity: 210,
  "pain point": 350, "pain-point": 350,
  delight: 140, quote: 40, insight: 220,
  "task issue": 30, task: 30, bookmark: 280,
  onboarding: 220, behavior: 145, layout: 280,
  copy: 40, performance: 30, uncategorized: 220,
};

function categoryHue(label) {
  if (!label) return 220;
  var key = String(label).toLowerCase().trim();
  if (Object.prototype.hasOwnProperty.call(CATEGORY_HUES, key)) {
    return CATEGORY_HUES[key];
  }
  // Stable fallback hash so unknown labels still get a consistent color.
  var h = 0;
  for (var i = 0; i < key.length; i++) {
    h = (h * 31 + key.charCodeAt(i)) | 0;
  }
  return ((h % 360) + 360) % 360;
}

function categoryColor(label, alpha) {
  var hue = categoryHue(label);
  if (alpha == null || alpha >= 1) {
    return "oklch(0.7 0.16 " + hue + ")";
  }
  return "oklch(0.7 0.16 " + hue + " / " + alpha + ")";
}

// ---- Shared settings (localStorage) ----

var THEME_STORAGE_KEY = "clipgen-theme";
var TOOLTIP_STORAGE_KEY = "clipgen-tooltips";

var applyStoredThemePreference = function () {
  var stored = null;
  try { stored = window.localStorage.getItem(THEME_STORAGE_KEY); } catch (_) {}
  var theme = stored === "light" ? "light" : "dark";
  document.documentElement.setAttribute("data-theme", theme);
  updateThemeToggleButton(theme);
};

var toggleThemePreference = function () {
  var root = document.documentElement;
  var current = root.getAttribute("data-theme") || "dark";
  var next = current === "light" ? "dark" : "light";
  root.setAttribute("data-theme", next);
  try { window.localStorage.setItem(THEME_STORAGE_KEY, next); } catch (_) {}
  updateThemeToggleButton(next);
  refreshDetectorColors();
};

var updateThemeToggleButton = function (explicitTheme) {
  var btn = qs("#themeToggle");
  if (!btn) return;
  var effective = explicitTheme === "light" ? "light" : "dark";
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

// ---- Frontend switcher (shared across Studio / Screenspace / Transcripts) ----

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

// ---- Canvas helpers (timeline overlays) ----

// Draw stacked per-series amplitude bands inside a canvas rect.
//
// Pure: no DOM lookups, no global state. Each timeline page (screenspace,
// viewer, transcripts, convergence) builds a `series` array from its own
// event shape and calls this helper from inside its renderTimeline().
//
// opts:
//   x, y, w, h         band rect on the canvas (pixels, integer)
//   visStart, visEnd   visible time window (seconds)
//   series             [{ key, color, timestamps: number[] }, ...]
//                      `key` is used for dim comparison; `color` is "#rrggbb";
//                      `timestamps` are seconds (already filtered to visible scope
//                      is fine but not required — out-of-window samples are skipped)
//   binPx              column width in pixels for binning (default 2)
//   dimKey             optional series key to keep at full opacity; all other series
//                      render at reduced alpha. Pass null/undefined for no dimming.
//
// Each series is normalized against its OWN peak so quiet types still show
// shape. Curves are drawn back-to-front so dimKey paints last when set.
var drawAmplitudeBands = function (ctx, opts) {
  var x = opts.x, y = opts.y, w = opts.w, h = opts.h;
  var visStart = opts.visStart, visEnd = opts.visEnd;
  var series = opts.series || [];
  var binPx = opts.binPx || 2;
  var dimKey = opts.dimKey;

  if (w <= 0 || h <= 0 || series.length === 0) return;
  var visLen = visEnd - visStart;
  if (!(visLen > 0)) return;

  var numBins = Math.max(1, Math.ceil(w / binPx));
  var binSec = visLen / numBins;

  // Bin each series and remember its own max
  var binned = [];
  for (var s = 0; s < series.length; s++) {
    var ts = series[s].timestamps || [];
    var bins = new Array(numBins);
    for (var b = 0; b < numBins; b++) bins[b] = 0;
    var maxCount = 0;
    for (var i = 0; i < ts.length; i++) {
      var t = ts[i];
      if (t < visStart || t >= visEnd) continue;
      var idx = Math.floor((t - visStart) / binSec);
      if (idx < 0) idx = 0;
      else if (idx >= numBins) idx = numBins - 1;
      var c = bins[idx] + 1;
      bins[idx] = c;
      if (c > maxCount) maxCount = c;
    }
    binned.push({ key: series[s].key, color: series[s].color, bins: bins, max: maxCount });
  }

  // Order: dimmed series first, focused series last (paints on top)
  var order = [];
  for (var k = 0; k < binned.length; k++) {
    if (dimKey && binned[k].key !== dimKey) order.push(k);
  }
  for (var k2 = 0; k2 < binned.length; k2++) {
    if (!dimKey || binned[k2].key === dimKey) order.push(k2);
  }

  var baselineY = y + h;
  for (var oi = 0; oi < order.length; oi++) {
    var ser = binned[order[oi]];
    if (ser.max <= 0) continue;
    var dimmed = dimKey && ser.key !== dimKey;
    var fillAlpha = dimmed ? 0.05 : 0.18;
    var strokeAlpha = dimmed ? 0.25 : 1.0;

    // Build the area path along bin tops
    ctx.beginPath();
    ctx.moveTo(x, baselineY);
    for (var bi = 0; bi < numBins; bi++) {
      var norm = ser.bins[bi] / ser.max;
      var py = baselineY - norm * h;
      var px = x + bi * binPx;
      ctx.lineTo(px, py);
      ctx.lineTo(px + binPx, py);
    }
    ctx.lineTo(x + numBins * binPx, baselineY);
    ctx.closePath();
    ctx.fillStyle = hexToRgba(ser.color, fillAlpha);
    ctx.fill();

    ctx.beginPath();
    var started = false;
    for (var bi2 = 0; bi2 < numBins; bi2++) {
      var n2 = ser.bins[bi2] / ser.max;
      var py2 = baselineY - n2 * h;
      var px2 = x + bi2 * binPx;
      if (!started) { ctx.moveTo(px2, py2); started = true; }
      else ctx.lineTo(px2, py2);
      ctx.lineTo(px2 + binPx, py2);
    }
    ctx.lineWidth = 1.5;
    ctx.strokeStyle = strokeAlpha === 1.0 ? ser.color : hexToRgba(ser.color, strokeAlpha);
    ctx.stroke();
  }
};
