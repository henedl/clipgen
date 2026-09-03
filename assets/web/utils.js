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

// Gates dev-token-tweak.js; viewer.py strips data-dev-only tags from exports regardless.
var CLIPGEN_DEV_TOKEN_TWEAK = false;

// ---- Canonical config ----
// Offline defaults mirroring config.py; clipgenApplyConfig overlays live payloads. tests/test_shared_constants.py checks.

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
  annotations: [
    { id: "key", token: "!key" },
  ],
  ignoredTimestampTokens: ["x"],
  screenspaceOcrMinConfidence: 0.6,
  screenspaceOcrFuzzyThreshold: 0.75,
  screenspaceMultitoolMaxOffset: 30,
  screenspaceMaskFallbackTools: ["similarity", "inactivity", "boundary", "timelapse", "attention"],
  frictionCategories: [
    { key: "hesitation",      label: "Hesitation" },
    { key: "confusion",       label: "Confusion" },
    { key: "frustration",     label: "Frustration" },
    { key: "surprise",        label: "Surprise" },
    { key: "self_correction", label: "Self-correction" },
    { key: "help_seeking",    label: "Help-seeking" },
  ],
  frictionColorToken: "--color-friction",
  frictionMomentLimit: 5,
  convergenceSources: ["sheet", "screenspace", "transcript", "composer"],
  cardScrubberSpriteCols: 5,
  cardScrubberSpriteRows: 5,
  clipFormat: ".mp4",
  screenshotFormat: ".png",
  gifFormat: ".gif",
  composerAnnotationColor: "#f05a3c",
  composerAnnotationColorSecondary: "#f8fafc",
  composerAnnotationStrokeWidth: 0.004,
  composerAnnotationStrokeStyle: "solid",
  composerAnnotationFontSize: 0.035,
  composerAnnotationSpanSeconds: 10.0,
  composerScrubMaxAudioSeconds: 180.0,
  composerDoubleClickCuts: true,
  crossReferences: true,
  mediaContainerWarning: true,
  // Mirrors video.SUBTITLE_CODEC_BY_CONTAINER / SUBTITLE_ALWAYS_DEFAULT_CONTAINERS;
  // the mp4 muxer ignores -disposition:s:0.
  subtitleContainers: {
    supported: [".m4v", ".mkv", ".mov", ".mp4", ".webm"],
    alwaysDefault: [".m4v", ".mov", ".mp4"],
  },
  hotkeyOverrides: {},
  profiling: false,
};

var clipgenApplyConfig = function (payload) {
  if (!payload || typeof payload !== "object") return;
  if (typeof payload.defaultDuration === "number") {
    CLIPGEN_CONFIG.defaultDuration = payload.defaultDuration;
  }
  if (Array.isArray(payload.severity)) {
    CLIPGEN_CONFIG.severity = payload.severity;
  }
  if (Array.isArray(payload.annotationKeyphrases)) {
    CLIPGEN_CONFIG.annotationKeyphrases = payload.annotationKeyphrases;
  }
  if (Array.isArray(payload.annotations)) {
    CLIPGEN_CONFIG.annotations = payload.annotations;
  }
  if (Array.isArray(payload.ignoredTimestampTokens)) {
    CLIPGEN_CONFIG.ignoredTimestampTokens = payload.ignoredTimestampTokens;
  }
  if (typeof payload.screenspaceOcrMinConfidence === "number") {
    CLIPGEN_CONFIG.screenspaceOcrMinConfidence = payload.screenspaceOcrMinConfidence;
  }
  if (typeof payload.screenspaceOcrFuzzyThreshold === "number") {
    CLIPGEN_CONFIG.screenspaceOcrFuzzyThreshold = payload.screenspaceOcrFuzzyThreshold;
  }
  if (typeof payload.screenspaceMultitoolMaxOffset === "number") {
    CLIPGEN_CONFIG.screenspaceMultitoolMaxOffset = payload.screenspaceMultitoolMaxOffset;
  }
  if (Array.isArray(payload.screenspaceMaskFallbackTools)) {
    CLIPGEN_CONFIG.screenspaceMaskFallbackTools = payload.screenspaceMaskFallbackTools;
  }
  if (Array.isArray(payload.frictionCategories)) {
    CLIPGEN_CONFIG.frictionCategories = payload.frictionCategories;
  }
  if (typeof payload.frictionColorToken === "string") {
    CLIPGEN_CONFIG.frictionColorToken = payload.frictionColorToken;
  }
  if (typeof payload.frictionMomentLimit === "number") {
    CLIPGEN_CONFIG.frictionMomentLimit = payload.frictionMomentLimit;
  }
  if (Array.isArray(payload.convergenceSources)) {
    CLIPGEN_CONFIG.convergenceSources = payload.convergenceSources;
  }
  if (typeof payload.cardScrubberSpriteCols === "number") {
    CLIPGEN_CONFIG.cardScrubberSpriteCols = payload.cardScrubberSpriteCols;
  }
  if (typeof payload.cardScrubberSpriteRows === "number") {
    CLIPGEN_CONFIG.cardScrubberSpriteRows = payload.cardScrubberSpriteRows;
  }
  if (typeof payload.clipFormat === "string") {
    CLIPGEN_CONFIG.clipFormat = payload.clipFormat;
  }
  if (typeof payload.screenshotFormat === "string") {
    CLIPGEN_CONFIG.screenshotFormat = payload.screenshotFormat;
  }
  if (typeof payload.gifFormat === "string") {
    CLIPGEN_CONFIG.gifFormat = payload.gifFormat;
  }
  if (typeof payload.composerAnnotationColor === "string") {
    CLIPGEN_CONFIG.composerAnnotationColor = payload.composerAnnotationColor;
  }
  if (typeof payload.composerAnnotationColorSecondary === "string") {
    CLIPGEN_CONFIG.composerAnnotationColorSecondary = payload.composerAnnotationColorSecondary;
  }
  if (typeof payload.composerAnnotationStrokeWidth === "number") {
    CLIPGEN_CONFIG.composerAnnotationStrokeWidth = payload.composerAnnotationStrokeWidth;
  }
  if (typeof payload.composerAnnotationStrokeStyle === "string") {
    CLIPGEN_CONFIG.composerAnnotationStrokeStyle = payload.composerAnnotationStrokeStyle;
  }
  if (typeof payload.composerAnnotationFontSize === "number") {
    CLIPGEN_CONFIG.composerAnnotationFontSize = payload.composerAnnotationFontSize;
  }
  if (typeof payload.composerAnnotationSpanSeconds === "number") {
    CLIPGEN_CONFIG.composerAnnotationSpanSeconds = payload.composerAnnotationSpanSeconds;
  }
  if (typeof payload.composerScrubMaxAudioSeconds === "number") {
    CLIPGEN_CONFIG.composerScrubMaxAudioSeconds = payload.composerScrubMaxAudioSeconds;
  }
  if (typeof payload.composerDoubleClickCuts === "boolean") {
    CLIPGEN_CONFIG.composerDoubleClickCuts = payload.composerDoubleClickCuts;
  }
  if (typeof payload.crossReferences === "boolean") {
    CLIPGEN_CONFIG.crossReferences = payload.crossReferences;
  }
  if (typeof payload.mediaContainerWarning === "boolean") {
    CLIPGEN_CONFIG.mediaContainerWarning = payload.mediaContainerWarning;
  }
  if (payload.subtitleContainers && typeof payload.subtitleContainers === "object") {
    CLIPGEN_CONFIG.subtitleContainers = payload.subtitleContainers;
  }
  if (payload.hotkeyOverrides && typeof payload.hotkeyOverrides === "object") {
    CLIPGEN_CONFIG.hotkeyOverrides = payload.hotkeyOverrides;
    if (window.ClipgenHotkeys) {
      window.ClipgenHotkeys.applyOverrides(payload.hotkeyOverrides);
    }
  }
  if (typeof payload.profiling === "boolean") {
    CLIPGEN_CONFIG.profiling = payload.profiling;
    if (payload.profiling && window.clipgenPerf) {
      window.clipgenPerf.observe();
    }
  }
};

// ---- Performance instrumentation ----
// No-op unless CLIPGEN_CONFIG.profiling; read via `tests/ui/shot.py --perf`. See agents/skills/profile/SKILL.md.
var clipgenPerf = (function () {
  var acc = {
    measures: {},
    longtasks: { count: 0, totalMs: 0, maxMs: 0 },
  };
  window.__clipgenPerf = acc;
  var observer = null;

  function record(label, ms) {
    var m = acc.measures[label];
    if (!m) { m = acc.measures[label] = { totalMs: 0, n: 0, maxMs: 0 }; }
    m.totalMs += ms;
    m.n += 1;
    if (ms > m.maxMs) m.maxMs = ms;
  }

  function begin(label) {
    if (!CLIPGEN_CONFIG.profiling) return;
    try { performance.mark("cg:" + label + ":start"); } catch (_) {}
  }

  function end(label) {
    if (!CLIPGEN_CONFIG.profiling) return;
    try {
      var entry = performance.measure("cg:" + label, "cg:" + label + ":start");
      record(label, entry.duration);
      performance.clearMarks("cg:" + label + ":start");
      performance.clearMeasures("cg:" + label);
    } catch (_) {} // begin() never ran for this label
  }

  // Time a synchronous function; returns its result.
  function span(label, fn) {
    if (!CLIPGEN_CONFIG.profiling) return fn();
    begin(label);
    try {
      return fn();
    } finally {
      end(label);
    }
  }

  // Records wall time per call, including a returned promise's async tail.
  function wrap(label, fn) {
    return function () {
      if (!CLIPGEN_CONFIG.profiling) return fn.apply(this, arguments);
      var t0 = performance.now();
      var result;
      try {
        result = fn.apply(this, arguments);
      } catch (e) {
        record(label, performance.now() - t0);
        throw e;
      }
      if (result && typeof result.then === "function") {
        var settle = function () { record(label, performance.now() - t0); };
        result.then(settle, settle);
      } else {
        record(label, performance.now() - t0);
      }
      return result;
    };
  }

  // Longtask observer: main-thread stalls >50ms, the browser's own signal.
  // Feature-detected; unsupported builds degrade to measures-only.
  function observe() {
    if (observer || !CLIPGEN_CONFIG.profiling) return;
    if (typeof PerformanceObserver === "undefined") return;
    try {
      observer = new PerformanceObserver(function (list) {
        var entries = list.getEntries();
        for (var i = 0; i < entries.length; i++) {
          var d = entries[i].duration;
          acc.longtasks.count += 1;
          acc.longtasks.totalMs += d;
          if (d > acc.longtasks.maxMs) acc.longtasks.maxMs = d;
        }
      });
      observer.observe({ entryTypes: ["longtask"] });
      window.addEventListener("pagehide", function () {
        if (observer) { observer.disconnect(); observer = null; }
      });
    } catch (_) {
      observer = null;
    }
  }

  // Plain data only: Playwright's serializer drops PerformanceEntry prototype getters.
  function snapshot() {
    return JSON.parse(JSON.stringify(acc));
  }

  return {
    begin: begin,
    end: end,
    span: span,
    wrap: wrap,
    record: record,
    observe: observe,
    snapshot: snapshot,
  };
})();
window.clipgenPerf = clipgenPerf;

// ---- DOM helpers ----

var qs = function (sel) { return document.querySelector(sel); };
var qsa = function (sel) { return document.querySelectorAll(sel); };

var el = function (tag, cls, text) {
  var e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};

// English-only count + noun phrase (1 → singular, else plural).
var clipgenPluralUnit = function (n, singular, plural) {
  return n + " " + (n === 1 ? singular : plural);
};

// Appends `cols` header cells plus rows×cols body cells into a .skeleton-grid target (primitives.css).
var buildSkeletonGrid = function (target, cols, rows) {
  if (!target) return;
  var frag = document.createDocumentFragment();
  for (var i = 0; i < cols; i++) {
    frag.appendChild(el("div", "skeleton skeleton-cell is-header"));
  }
  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < cols; c++) {
      frag.appendChild(el("div", "skeleton skeleton-cell"));
    }
  }
  target.appendChild(frag);
};

// Single source of truth for which animated artifacts render as <video>.
var isVideoLoop = function (filename) {
  return /\.webm$/i.test(filename || "");
};

// Looping, silent, autoplay <video> for .webm artifacts; keeps the attribute set uniform.
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

// ---- Export attribution (version + repo from CLIPGEN_DATA.meta) ----
var clipgenRenderFooter = function (meta) {
  var el = document.getElementById("footerCredit");
  if (!el || !meta) return;
  var version = meta.clipgenVersion ? " v" + meta.clipgenVersion : "";
  el.textContent = "Generated by clipgen" + version;
  if (!meta.repoUrl) return;
  el.appendChild(document.createTextNode(" \u00b7 "));
  var link = document.createElement("a");
  link.href = meta.repoUrl;
  link.target = "_blank";
  link.rel = "noopener noreferrer";
  link.textContent = meta.repoUrl.replace(/^https?:\/\//, "");
  el.appendChild(link);
};

// ---- Brand mark hydration ----
// Inlines logos/favicon.svg for tokens.css's stroke animation; plays once per session.
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
// mask-size, mask-repeat and currentColor fill come from the element's class.

var applyMaskIcon = function (el, urlValue) {
  el.style.maskImage = urlValue;
  el.style.webkitMaskImage = urlValue;
};

// Inline-style equivalent for embedding in HTML strings:
//   "mask-image: url(...); -webkit-mask-image: url(...);"
var maskIconStyle = function (urlValue) {
  return "mask-image:" + urlValue + ";-webkit-mask-image:" + urlValue + ";";
};

// ---- Icon masks by basename ----
// `basePath` defaults to "icons/"; other routes override it.

// Single-quoted so it can sit inside a double-quoted style attribute.
var iconMaskUrl = function (name, basePath) {
  return "url('" + (basePath || "icons/") + name + ".svg')";
};

// Inline-style string for embedding in HTML strings.
var iconMaskStyle = function (name, basePath) {
  return maskIconStyle(iconMaskUrl(name, basePath));
};

// Apply mask-image to an existing element from an icon basename.
var applyIconMask = function (el, name, basePath) {
  applyMaskIcon(el, iconMaskUrl(name, basePath));
};

// Build a <span> with the mask applied. opts = { className, basePath }.
var iconMaskSpan = function (name, opts) {
  opts = opts || {};
  var span = el("span", opts.className || "");
  applyMaskIcon(span, iconMaskUrl(name, opts.basePath));
  return span;
};

// opts = { selector, basePath }; default selector "[data-icon]".
var applyIconMasksIn = function (scope, opts) {
  opts = opts || {};
  var root = scope || document;
  var nodes = root.querySelectorAll(opts.selector || "[data-icon]");
  for (var i = 0; i < nodes.length; i++) {
    var name = nodes[i].getAttribute("data-icon");
    if (!name) continue;
    applyIconMask(nodes[i], name, opts.basePath);
  }
};

// ---- Segmented capsule track ----
// --seg-index moves the thumb in CSS; safe to build detached.
var createSegTrack = function (opts) {
  opts = opts || {};
  var options = opts.options || [];
  var track = el("div", "cg-segtrack" + (opts.size === "sm" ? " cg-segtrack--sm" : ""));
  // Set before the first style computation so the thumb never slides in.
  track.style.setProperty("--seg-count", String(options.length));
  var hidden = document.createElement("input");
  hidden.type = "hidden";
  if (opts.id) hidden.id = opts.id;
  track.appendChild(hidden);
  track.appendChild(el("span", "cg-segtrack-thumb"));
  options.forEach(function (spec) {
    var btn = el("button", "cg-segtrack-btn");
    btn.type = "button";
    btn.setAttribute("data-value", spec.value);
    if (spec.desc) btn.setAttribute("data-desc", spec.desc);
    // data-tooltip carries no accessible name; icon-only segments need aria-label.
    if (spec.title) {
      btn.setAttribute("data-tooltip", spec.title);
      if (!spec.label) btn.setAttribute("aria-label", spec.title);
    }
    if (spec.hotkey) btn.setAttribute("data-hotkey", spec.hotkey);
    if (spec.icon) {
      btn.appendChild(iconMaskSpan(spec.icon, { className: "cg-segtrack-icon", basePath: opts.basePath }));
    }
    if (spec.label) btn.appendChild(el("span", "cg-segtrack-label", spec.label));
    btn.addEventListener("click", function (e) {
      e.preventDefault();
      if (spec.value === hidden.value) return; // nothing to slide
      segTrackSetValue(track, spec.value);
      if (opts.onChange) opts.onChange(spec.value);
      // Bubbles so container-level input listeners (e.g. Screenspace's
      // addParamRow live-preview handler) fire, mirroring a checkbox change.
      hidden.dispatchEvent(new Event("input", { bubbles: true }));
    });
    track.appendChild(btn);
  });
  segTrackSetValue(track, opts.value);
  return track;
};

// Fires no events and no onChange; unknown values are a no-op.
var segTrackSetValue = function (trackEl, value) {
  var btns = trackEl.querySelectorAll(".cg-segtrack-btn");
  var index = -1;
  for (var i = 0; i < btns.length; i++) {
    if (btns[i].getAttribute("data-value") === value) { index = i; break; }
  }
  if (index < 0) return;
  for (var j = 0; j < btns.length; j++) {
    var active = j === index;
    btns[j].classList.toggle("active", active);
    btns[j].setAttribute("aria-pressed", active ? "true" : "false");
  }
  var hidden = trackEl.querySelector("input[type=hidden]");
  if (hidden) hidden.value = value;
  trackEl.style.setProperty("--seg-index", String(index));
};

// ---- Hover tooltips (dark pill, pairs with .cg-tooltip in tokens.css) ----

var createTooltip = function (opts) {
  opts = opts || {};
  var cls = "cg-tooltip";
  if (opts.multiline) cls += " cg-tooltip--multiline";
  var tip = document.createElement("div");
  tip.className = cls;
  // Separate icon node lets show() hang a glyph off controls too small to carry one.
  var iconEl = document.createElement("span");
  iconEl.className = "cg-tooltip-icon";
  var textEl = document.createElement("span");
  textEl.className = "cg-tooltip-text";
  tip.appendChild(iconEl);
  tip.appendChild(textEl);
  document.body.appendChild(tip);
  return {
    el: tip,
    // `icon` is an applyIconMask basename ("check", "octicon/dependabot-16"); omit for text-only.
    show: function (anchor, text, icon) {
      textEl.textContent = text || "";
      if (icon) {
        applyIconMask(iconEl, icon, opts.iconBasePath);
        tip.classList.add("cg-tooltip--has-icon");
      } else {
        tip.classList.remove("cg-tooltip--has-icon");
      }
      // Set content before measuring so offsetHeight reflects final size.
      positionTooltipAnchored(tip, anchor.getBoundingClientRect());
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

// Singleton [data-tooltip] tooltip; tokens.css's pointer-events:auto keeps disabled buttons firing mouseover.
var clipgenInitDataTooltips = function () {
  var tip = null;
  var current = null;
  var ensureTip = function () {
    if (!tip) tip = createTooltip({ multiline: true });
    return tip;
  };
  var shouldSuppress = function (el) {
    // Studio hides disabled-button tooltips while a generation is running.
    if (document.body.classList.contains("studio-generating")) {
      if (el.matches && el.matches(".btn[data-tooltip]:disabled")) return true;
    }
    return false;
  };
  // data-tooltip-icon rides along with data-tooltip; alone it shows nothing.
  var showFor = function (el) {
    if (shouldSuppress(el)) return;
    var text = el.getAttribute("data-tooltip");
    if (!text) return;
    current = el;
    ensureTip().show(el, text, el.getAttribute("data-tooltip-icon"));
  };
  var hide = function () {
    if (!current) return;
    if (tip) tip.hide();
    current = null;
  };
  document.addEventListener("mouseover", function (e) {
    var el = e.target.closest && e.target.closest("[data-tooltip]");
    if (!el || el === current) return;
    if (current && !current.contains(el)) hide();
    showFor(el);
  });
  document.addEventListener("mouseout", function (e) {
    if (!current) return;
    // mouseout fires when crossing into children; stay while relatedTarget is inside.
    var to = e.relatedTarget;
    if (to && current.contains(to)) return;
    hide();
  });
  window.addEventListener("scroll", hide, true);
  window.addEventListener("resize", hide);
};

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", clipgenInitDataTooltips);
} else {
  clipgenInitDataTooltips();
}

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

// Sheet cross-reference helpers bound to a hub's state (Studio and Overview).
var createSheetXrefHelpers = function (getState) {
  function parseClipTimestamps(raw, participantId) {
    var state = getState();
    var baselineSeconds = 0;
    if (participantId && state.convergenceBaselines) {
      baselineSeconds = state.convergenceBaselines[participantId] || 0;
    }
    return parseClipSegmentsForCell(raw, baselineSeconds, CLIPGEN_CONFIG.defaultDuration);
  }

  var ROW_FUNCTIONS = {
    Count: function (row, participants) {
      var total = 0;
      for (var j = 0; j < participants.length; j++) {
        var c = row.cells[participants[j]];
        if (c && c.valid) total += parseClipTimestamps(c.value, participants[j]).length;
      }
      return total;
    },
    Unique: function (row, participants) {
      var count = 0;
      for (var j = 0; j < participants.length; j++) {
        var c = row.cells[participants[j]];
        if (c && c.valid) count++;
      }
      return count;
    },
  };

  // Overlapping data from sibling sources for one participant + time range.
  function findOverlappingData(participant, start, end) {
    var state = getState();
    var result = { transcriptSnippets: [], screenspaceEvents: [], sheetObservations: [] };

    // Projection: consumers expect `text` already resolved (text || label).
    for (var i = 0; i < state.trIntakeClusters.length; i++) {
      var tc = state.trIntakeClusters[i];
      if (tc.participant === participant && tc.start < end && tc.end > start) {
        result.transcriptSnippets.push({ text: tc.text || tc.label || "", category: tc.category, start: tc.start, end: tc.end });
      }
    }

    // Pass the cluster through; consumers read only detector / event_type.
    for (var j = 0; j < state.intakeClusters.length; j++) {
      var sc = state.intakeClusters[j];
      if (sc.participant === participant && sc.start < end && sc.end > start) {
        result.screenspaceEvents.push(sc);
      }
    }

    if (state.sheetData && state.sheetData.rows) {
      for (var k = 0; k < state.sheetData.rows.length; k++) {
        var row = state.sheetData.rows[k];
        var cell = row.cells[participant];
        if (!cell || !cell.valid) continue;
        var segs = parseClipTimestamps(cell.value, participant);
        for (var s = 0; s < segs.length; s++) {
          var segEnd = segs[s].startSeconds + segs[s].duration;
          if (segs[s].startSeconds < end && segEnd > start) {
            result.sheetObservations.push(row);
            break;
          }
        }
      }
    }

    return result;
  }

  return {
    parseClipTimestamps: parseClipTimestamps,
    ROW_FUNCTIONS: ROW_FUNCTIONS,
    findOverlappingData: findOverlappingData,
  };
};

// Rounds rather than floors; use for durations (clip length, ruler ticks).
var formatDuration = function (sec) {
  if (sec == null || isNaN(sec)) return "--:--";
  var total = Math.round(sec);
  var h = Math.floor(total / 3600);
  var m = Math.floor((total % 3600) / 60);
  var s = total % 60;
  if (h > 0) return h + ":" + pad2(m) + ":" + pad2(s);
  return m + ":" + pad2(s);
};

// ---- Elapsed-time / ETA estimation for long-running operations ----

// Null unless 0 < progress < 1 and elapsed > 0; callers show elapsed only.
var estimateRemainingSec = function (elapsedSec, progress) {
  if (progress == null || !isFinite(progress) || progress <= 0 || progress >= 1) return null;
  if (!isFinite(elapsedSec) || elapsedSec <= 0) return null;
  return (elapsedSec * (1 - progress)) / progress;
};

// Elapsed plus EMA-smoothed remaining estimate; omit progress for indeterminate jobs. pause()/resume() exclude paused spans.
var createEtaTracker = function (opts) {
  opts = opts || {};
  var emaAlpha = opts.emaAlpha != null ? opts.emaAlpha : 0.3;
  var minProgress = opts.minProgress != null ? opts.minProgress : 0.05;
  var minElapsed = opts.minElapsed != null ? opts.minElapsed : 3;
  var startMs = null;
  var ema = null;
  var pausedMs = 0; // total ms spent paused across completed pause spans
  var pausedAt = null; // epoch (ms) the current pause began, else null
  return {
    // Idempotent; an explicit epoch (ms) seeds from a known start (reattach).
    start: function (nowMs) {
      if (startMs == null) startMs = nowMs != null ? nowMs : Date.now();
    },
    // Freeze elapsed at the pause instant. Idempotent — safe to call every tick.
    pause: function (nowMs) {
      if (pausedAt == null) pausedAt = nowMs != null ? nowMs : Date.now();
    },
    // Resume ticking, folding the just-ended pause span into pausedMs. No-op when
    // not paused.
    resume: function (nowMs) {
      if (pausedAt != null) {
        pausedMs += (nowMs != null ? nowMs : Date.now()) - pausedAt;
        pausedAt = null;
      }
    },
    // remainingSec stays null until the progress/elapsed gates open, then EMA-smoothed.
    update: function (progress) {
      if (startMs == null) startMs = Date.now();
      // Paused: "now" holds at pausedAt, so elapsed freezes until resume().
      var now = pausedAt != null ? pausedAt : Date.now();
      var elapsedSec = (now - startMs - pausedMs) / 1000;
      var raw = estimateRemainingSec(elapsedSec, progress);
      var remainingSec = null;
      if (raw != null && progress >= minProgress && elapsedSec >= minElapsed) {
        // First real estimate seeds the EMA directly; later ones blend to damp jitter.
        ema = ema == null ? raw : emaAlpha * raw + (1 - emaAlpha) * ema;
        remainingSec = Math.round(ema);
      }
      return { elapsedSec: elapsedSec, remainingSec: remainingSec };
    },
    reset: function () {
      startMs = null;
      ema = null;
      pausedMs = 0;
      pausedAt = null;
    },
  };
};

// "~1:20 left"; bucketed so the text doesn't twitch each tick. "" without estimate.
var formatEtaLabel = function (remainingSec) {
  if (remainingSec == null || !isFinite(remainingSec) || remainingSec < 0) return "";
  var bucketed;
  if (remainingSec < 60) bucketed = Math.round(remainingSec / 5) * 5;
  else if (remainingSec < 300) bucketed = Math.round(remainingSec / 10) * 10;
  else bucketed = Math.round(remainingSec / 30) * 30;
  if (bucketed <= 0) bucketed = 5;
  return "~" + formatDuration(bucketed) + " left";
};

// Fixed-interval ticker; self-stops when isActive() is false or (gateHidden) the tab hides.
var createIntervalTicker = function (tickFn, opts) {
  opts = opts || {};
  var intervalMs = opts.intervalMs != null ? opts.intervalMs : 1000;
  var gateHidden = !!opts.gateHidden;
  var isActive = opts.isActive || null;
  var handle = null;
  function stop() {
    if (handle) {
      clearInterval(handle);
      handle = null;
    }
  }
  function run() {
    if (gateHidden && typeof document !== "undefined" && document.hidden) {
      stop();
      return;
    }
    if (isActive && !isActive()) {
      stop();
      return;
    }
    tickFn();
  }
  return {
    ensure: function () {
      if (!handle) handle = setInterval(run, intervalMs);
    },
    stop: stop,
  };
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

// Clock semantics: 2-part is HH:MM. Mirrors Python utils._clock_to_seconds.
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

// Mirrors files.prepare_clip + utils.convert_clock_pairs_to_relative. A baseline makes tokens clock times; defaultDuration is required.
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

// parseFloat with a fallback for NaN/empty/garbage input.
var numberOrDefault = function (value, fallback) {
  var n = parseFloat(value);
  return isNaN(n) ? fallback : n;
};

// parseInt (base 10) with a fallback for NaN/empty/garbage input.
var intOrDefault = function (value, fallback) {
  var n = parseInt(value, 10);
  return isNaN(n) ? fallback : n;
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

// Centered above the anchor; flips below when cramped, clamps horizontally.
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

// Below the anchor, flipping above when cramped; popover must be visible to measure.
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

// ---- Debounce / escaping / color / toast (shared across pages) ----

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

// Escapes, then converts inline `code`, **bold**, hugging *italic*. No underscore emphasis: snake_case would mangle.
var clipgenRenderInlineMarkdown = function (str) {
  var html = escapeHtml(str == null ? "" : String(str));
  html = html.replace(/`([^`]+)`/g, "<code>$1</code>");
  html = html.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
  html = html.replace(/(^|[\s(])\*([^*\s](?:[^*]*[^*\s])?)\*(?=$|[\s).,;:!?])/g, "$1<em>$2</em>");
  return html;
};

var hexToRgba = function (hex, alpha) {
  var r = parseInt(hex.slice(1, 3), 16);
  var g = parseInt(hex.slice(3, 5), 16);
  var b = parseInt(hex.slice(5, 7), 16);
  return "rgba(" + r + "," + g + "," + b + "," + alpha + ")";
};

// hexToRgb takes "#rgb"/"#rrggbb" (hash optional), else null; rgbToHex clamps and rounds float channels.
var hexToRgb = function (hex) {
  var h = String(hex == null ? "" : hex).trim().replace(/^#/, "");
  if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  if (!/^[0-9a-fA-F]{6}$/.test(h)) return null;
  var n = parseInt(h, 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
};

var rgbToHex = function (r, g, b) {
  function c(x) {
    x = Math.max(0, Math.min(255, Math.round(x)));
    var s = x.toString(16);
    return s.length < 2 ? "0" + s : s;
  }
  return "#" + c(r) + c(g) + c(b);
};

// Read a :root custom property, falling back when unset or empty.
var getCSSVar = function (name, fallback) {
  try {
    var v = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    return v || fallback;
  } catch (_) {
    return fallback;
  }
};

// Per-frame canvas renderers read this cache; the theme toggle invalidates it.
var _canvasThemeColorsCache = null;

var getCanvasThemeColors = function () {
  if (_canvasThemeColorsCache) return _canvasThemeColorsCache;
  var cs = getComputedStyle(document.documentElement);
  function v(name, fallback) {
    var x = cs.getPropertyValue(name).trim();
    return x || fallback;
  }
  _canvasThemeColorsCache = {
    fg:         v("--fg",                "#ffffff"),
    bg:         v("--bg",                "#0d0e10"),
    surfaceAlt: v("--color-surface-alt", "#f1ece4"),
    border:     v("--color-border",      "#e0ddd7"),
    textDim:    v("--color-text-dim",    "#6b7280"),
    accent:     v("--color-accent",      "#1d4f72"),
    heatmap:    v("--color-heatmap",     "#3b82f6"),
    positive:   v("--cell-data-positive-fg", "#16a34a"),
    fontMono:   v("--font-mono",         "monospace"),
  };
  return _canvasThemeColorsCache;
};

var invalidateCanvasThemeColors = function () {
  _canvasThemeColorsCache = null;
};

var SHOW_TOAST_DEFAULT_MS = 3000;

var showToast = function (msg, opts) {
  var durationMs = SHOW_TOAST_DEFAULT_MS;
  if (opts != null && opts.durationMs != null) durationMs = opts.durationMs;
  var toastEl = qs("#toast");
  if (!toastEl) return;
  toastEl.textContent = msg;
  toastEl.classList.remove("hidden");
  // Reused element: a stale fade-out must not hide a fresh toast.
  var gen = (toastEl._toastGen = (toastEl._toastGen || 0) + 1);
  // Fade in supersedes a stale exit fill; motion.js only loads on some pages.
  if (window.ClipgenMotion) ClipgenMotion.animateIn(toastEl, "fade");
  clearTimeout(toastEl._timer);
  toastEl._timer = setTimeout(function () {
    var hide = function () {
      if (toastEl._toastGen === gen) toastEl.classList.add("hidden");
    };
    if (window.ClipgenMotion) ClipgenMotion.animateOut(toastEl, "fade").then(hide);
    else hide();
  }, durationMs);
};

// ---- Severity ----
// CLIPGEN_CONFIG.severity mirrors config.py SEVERITY_NUMERIC_TO_LABEL (tests/test_shared_constants.py).

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

// Lowest = most severe (Critical -4). Null for unknown; callers decide how to treat it.
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

// Vertical wheel scrolls an overflowing strip horizontally; passive:false so preventDefault works.
var clipgenWheelToHorizontal = function (el) {
  el.addEventListener(
    "wheel",
    function (e) {
      if (el.scrollWidth > el.clientWidth) {
        e.preventDefault();
        el.scrollLeft += e.deltaY;
      }
    },
    { passive: false }
  );
};

// ---- API helpers (always check r.ok) ----

// Rejects with the server's envelope error; .status for branching, .serverMessage empty for generic failures.
var _apiJson = function (r) {
  if (!r.ok) {
    var mkError = function (message) {
      var e = new Error(message || "Server error " + r.status);
      e.status = r.status;
      e.serverMessage = message || "";
      return e;
    };
    return r.json().then(
      function (data) {
        throw mkError(data && data.error);
      },
      function () {
        throw mkError(null);
      }
    );
  }
  return r.json();
};

var apiGet = function (path) {
  return fetch(path).then(_apiJson);
};

var apiPost = function (path, body) {
  return fetch(path, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  }).then(_apiJson);
};

var apiPut = function (path, body) {
  return fetch(path, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  }).then(_apiJson);
};

var apiPatch = function (path, body) {
  return fetch(path, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  }).then(_apiJson);
};

var apiDelete = function (path) {
  return fetch(path, { method: "DELETE" }).then(_apiJson);
};

// Blob variants for image/media routes (frame thumbnails, preview renders).
var apiGetBlob = function (path) {
  return fetch(path).then(function (r) {
    if (!r.ok) throw new Error("Server error " + r.status);
    return r.blob();
  });
};

var apiPostBlob = function (path, body) {
  return fetch(path, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  }).then(function (r) {
    if (!r.ok) throw new Error("Server error " + r.status);
    return r.blob();
  });
};

// .catch handler that toasts the failure; for user-initiated calls, not background polling.
var toastError = function (prefix) {
  return function (err) {
    showToast(prefix + (err && err.message ? ": " + err.message : ""));
  };
};

// ---- Polling ----

var POLL_INTERVAL = 3000;

// Live-refresh loop with hidden-tab pause. maxIntervalMs > intervalMs enables backoff; fn's truthy result means "active".
var createPoller = function (fn, intervalMs, opts) {
  opts = opts || {};
  if (opts.label) {
    fn = clipgenPerf.wrap("poll." + opts.label, fn);
  }
  var pauseWhenHidden = opts.pauseWhenHidden !== false;
  var runImmediately = opts.runImmediately !== false;
  var maxIntervalMs = opts.maxIntervalMs != null ? opts.maxIntervalMs : intervalMs;
  var backoffAfter = opts.backoffAfter != null ? opts.backoffAfter : 3;
  var adaptive = maxIntervalMs > intervalMs;
  var timer = null;
  var visListener = null;
  var wantRunning = false;
  var currentDelay = intervalMs;
  var quiet = 0;
  var inFlight = false;
  var pendingWake = false;
  // Settles only after any pendingWake follow-up, so wake() callers' spinners stop on time.
  var currentRun = null;

  function safeFn() {
    try { fn(); } catch (_) {}
  }
  function hidden() {
    return pauseWhenHidden && typeof document !== "undefined" && document.hidden;
  }
  // --- backoff (self-rescheduling) path ---
  function applySignal(active) {
    if (active) {
      quiet = 0;
      currentDelay = intervalMs;
    } else {
      quiet += 1;
      if (quiet >= backoffAfter && currentDelay < maxIntervalMs) {
        currentDelay = Math.min(currentDelay * 2, maxIntervalMs);
      }
    }
  }
  function schedule() {
    timer = setTimeout(runAdaptive, currentDelay);
  }
  // Resolves once this poll and any wake()-queued follow-up complete.
  function runAdaptive() {
    timer = null;
    inFlight = true;
    var result;
    try { result = fn(); } catch (_) { result = false; }
    currentRun = Promise.resolve(result).then(
      function (active) { applySignal(!!active); },
      function () { applySignal(false); }
    ).then(function () {
      inFlight = false;
      if (!wantRunning || hidden()) { pendingWake = false; return; }
      // A mid-flight wake() may predate the user's mutation; poll again now.
      if (pendingWake) {
        pendingWake = false;
        currentDelay = intervalMs;
        quiet = 0;
        return runAdaptive();
      }
      schedule();
    });
    return currentRun;
  }
  function arm() {
    if (timer != null || inFlight) return;
    if (hidden()) return;
    if (adaptive) {
      if (runImmediately) runAdaptive(); else schedule();
    } else {
      if (runImmediately) safeFn();
      timer = setInterval(safeFn, intervalMs);
    }
  }
  function disarm() {
    if (timer != null) {
      if (adaptive) clearTimeout(timer); else clearInterval(timer);
      timer = null;
    }
  }
  function onVisibility() {
    if (!wantRunning) return;
    if (document.hidden) {
      disarm();
    } else {
      if (adaptive) { currentDelay = intervalMs; quiet = 0; }
      arm();
    }
  }
  return {
    start: function () {
      if (wantRunning) return;
      wantRunning = true;
      if (adaptive) { currentDelay = intervalMs; quiet = 0; }
      arm();
      if (pauseWhenHidden && !visListener) {
        visListener = onVisibility;
        document.addEventListener("visibilitychange", visListener);
      }
    },
    stop: function () {
      wantRunning = false;
      disarm();
      if (visListener) {
        document.removeEventListener("visibilitychange", visListener);
        visListener = null;
      }
    },
    // Resolves when the triggered refresh lands; already resolved when nothing new runs.
    wake: function () {
      if (!wantRunning || hidden()) return Promise.resolve();
      if (adaptive) {
        currentDelay = intervalMs;
        quiet = 0;
        // The in-flight poll may predate this action; queue a follow-up and return the chained run.
        if (inFlight) { pendingWake = true; return currentRun || Promise.resolve(); }
        disarm();
        return runAdaptive();
      }
      disarm();
      safeFn();
      timer = setInterval(safeFn, intervalMs);
      return Promise.resolve();
    },
  };
};

// Lazily-built createPoller behind an idempotent start/stop pair.
var createManagedPoller = function (fn, intervalMs, opts) {
  var poller = null;
  return {
    start: function () {
      if (poller) return;
      poller = createPoller(fn, intervalMs, opts);
      poller.start();
    },
    stop: function () {
      if (!poller) return;
      poller.stop();
      poller = null;
    },
  };
};

// ---- SSE stream ----
// Closes itself before onError so the caller's polling fallback can start.
var createSSEStream = function (url, opts) {
  opts = opts || {};
  if (!window.EventSource) {
    if (opts.onUnsupported) opts.onUnsupported();
    return null;
  }
  var es = new EventSource(url);
  if (opts.onOpen) es.onopen = function () { opts.onOpen(); };
  es.onmessage = function (e) {
    var data;
    try { data = JSON.parse(e.data); } catch (_) { return; }
    if (opts.onMessage) opts.onMessage(data);
  };
  es.onerror = function () {
    es.close();
    if (opts.onError) opts.onError();
  };
  return es;
};

// ---- NDJSON streaming reader ----
// onLine(trimmedLine) per non-empty line; rejects when the body cannot stream.
var readNDJSONStream = function (response, onLine) {
  if (!response.body || typeof response.body.getReader !== "function") {
    return Promise.reject(new Error("Streaming response not supported"));
  }
  var reader = response.body.getReader();
  var decoder = new TextDecoder();
  var buffer = "";
  function pump() {
    return reader.read().then(function (result) {
      if (result.done) {
        if (buffer.trim()) onLine(buffer.trim());
        return;
      }
      buffer += decoder.decode(result.value, { stream: true });
      var lines = buffer.split("\n");
      buffer = lines.pop();
      for (var i = 0; i < lines.length; i++) {
        if (lines[i].trim()) onLine(lines[i].trim());
      }
      return pump();
    });
  }
  return pump();
};

// Streaming POST; !r.ok rejects with .status and .bodyText for branching, aborts pass through as AbortError.
var apiPostNDJSON = function (path, body, opts) {
  opts = opts || {};
  return fetch(path, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
    signal: opts.signal,
  }).then(function (r) {
    if (!r.ok) {
      var mkError = function (txt) {
        var e = new Error("Server error " + r.status);
        e.status = r.status;
        e.bodyText = txt || "";
        return e;
      };
      return r.text().then(
        function (txt) {
          throw mkError(txt);
        },
        function () {
          throw mkError("");
        }
      );
    }
    return readNDJSONStream(r, opts.onLine || function () {});
  });
};

// ---- File downloads ----
// WKWebView drops every download, so desktop.py's save_file writes the bytes instead.
var clipgenSaveFile = function (filename, content, mime, onDone) {
  var api = window.pywebview && window.pywebview.api;
  if (api && typeof api.save_file === "function") {
    api.save_file(filename, content).then(
      function (path) { if (onDone) onDone(path || null, null); },
      function (err) { if (onDone) onDone(null, err); }
    );
    return;
  }
  var url = URL.createObjectURL(new Blob([content], { type: mime || "application/octet-stream" }));
  var a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  if (onDone) onDone(filename, null);
};

// For server-built payloads; navigating to the URL cannot download inside the desktop window.
var clipgenSaveFromUrl = function (url, filename, onDone) {
  fetch(url)
    .then(function (r) {
      if (!r.ok) throw new Error("Server error " + r.status);
      return r.text();
    })
    .then(function (text) {
      var mime = /\.csv$/i.test(filename) ? "text/csv" : "application/json";
      clipgenSaveFile(filename, text, mime, onDone);
    })
    .catch(function (err) { if (onDone) onDone(null, err); });
};

// ---- Video seek coalescer ----
// Defers seeks until loadedmetadata; RAF-coalesces bursts to one per frame.
var createSeekCoalescer = function (getVideo, onDeferred, applySeek) {
  var pendingTime = null;
  var raf = 0;
  var listener = null;
  return {
    cancel: function () {
      var video = getVideo();
      pendingTime = null;
      cancelAnimationFrame(raf);
      raf = 0;
      if (listener) {
        if (video) video.removeEventListener("loadedmetadata", listener);
        listener = null;
      }
    },
    seek: function (time) {
      var video = getVideo();
      if (!video || !video.src) return;
      // Remove any previous deferred-seek listener.
      if (listener) {
        video.removeEventListener("loadedmetadata", listener);
        listener = null;
      }
      // Metadata not loaded yet: defer the seek until it is.
      if (video.readyState < 1) {
        pendingTime = time;
        listener = function () {
          video.removeEventListener("loadedmetadata", listener);
          listener = null;
          var t = pendingTime;
          pendingTime = null;
          if (t !== null) onDeferred(t);
        };
        video.addEventListener("loadedmetadata", listener);
        return;
      }
      // Coalesce rapid seeks into one write per animation frame.
      pendingTime = time;
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(function () {
        var t = pendingTime;
        pendingTime = null;
        raf = 0;
        if (t === null) return;
        applySeek(video, t);
      });
    },
  };
};

// ---- Blocking modal lifecycle (Escape / backdrop / focus trap) ----
// Singleton; release() is idempotent.
var _TRAP_FOCUSABLE =
  'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])';
var _activeBlockingModal = null;

// hotkeys.js scopes Alt-hold hints to this root. Self-managed modals (settings-modal.js) set it explicitly.
var _activeModalRoot = null;
var setActiveModalRoot = function (el) { _activeModalRoot = el || null; };
var getActiveModalRoot = function () { return _activeModalRoot; };

var openBlockingModal = function (overlayEl, opts) {
  opts = opts || {};
  if (!overlayEl) return null;
  if (_activeBlockingModal && _activeBlockingModal.el === overlayEl) {
    _activeBlockingModal.opts = opts;
    return _activeBlockingModal;
  }
  if (_activeBlockingModal) _activeBlockingModal.release();
  var prevFocus = opts.restoreFocus ? document.activeElement : null;

  function visibleFocusable() {
    return Array.prototype.slice
      .call(overlayEl.querySelectorAll(_TRAP_FOCUSABLE))
      .filter(function (n) { return !n.disabled && n.offsetParent !== null; });
  }
  function onKey(ev) {
    if (ev.key === "Escape") {
      ev.preventDefault();
      if (trap.opts.onEscape) trap.opts.onEscape();
      return;
    }
    if (!trap.opts.trapFocus || ev.key !== "Tab") return;
    // Re-query each Tab: button visibility changes between overlay phases.
    var f = visibleFocusable();
    if (f.length === 0) { ev.preventDefault(); return; }
    var first = f[0];
    var last = f[f.length - 1];
    if (ev.shiftKey && document.activeElement === first) {
      ev.preventDefault();
      last.focus();
    } else if (!ev.shiftKey && document.activeElement === last) {
      ev.preventDefault();
      first.focus();
    } else if (f.indexOf(document.activeElement) === -1) {
      // Focus parked outside the cycle (tabindex="-1" initialFocus): pull it back in.
      ev.preventDefault();
      (ev.shiftKey ? last : first).focus();
    }
  }
  function onClick(ev) {
    if (ev.target === overlayEl && trap.opts.onBackdropClick) {
      trap.opts.onBackdropClick();
    }
  }
  function release() {
    if (_activeBlockingModal !== trap) return;
    document.removeEventListener("keydown", onKey, true);
    overlayEl.removeEventListener("click", onClick);
    _activeBlockingModal = null;
    if (_activeModalRoot === overlayEl) _activeModalRoot = null;
    if (prevFocus && prevFocus.focus) prevFocus.focus();
  }

  var trap = { el: overlayEl, opts: opts, release: release };
  document.addEventListener("keydown", onKey, true);
  if (opts.onBackdropClick) overlayEl.addEventListener("click", onClick);
  if (opts.trapFocus) {
    var initial = visibleFocusable();
    var target = opts.initialFocus || initial[0];
    if (target && target.focus) target.focus();
  }
  _activeBlockingModal = trap;
  _activeModalRoot = overlayEl;
  return trap;
};

var closeBlockingModal = function (overlayEl) {
  if (_activeBlockingModal && _activeBlockingModal.el === overlayEl) {
    _activeBlockingModal.release();
  }
};

// hotkeys.js mutes page hotkeys on this; the palette refuses to steal an open modal's trap.
var isBlockingModalOpen = function () {
  return _activeBlockingModal !== null;
};

// ---- Modal reveal / dismiss ----
// Visual half of openBlockingModal's lifecycle; veil rules: tokens.css .cg-modal-veil.

var _cgVeilMs = function (overlayEl) {
  var raw = getComputedStyle(overlayEl).getPropertyValue("--duration-veil");
  var ms = parseFloat(raw);
  return isFinite(ms) && ms > 0 ? ms : 360;
};

var popModalIn = function (overlayEl, cardEl) {
  if (!overlayEl) return;
  var gen = (overlayEl._cgModalGen = (overlayEl._cgModalGen || 0) + 1);
  var wasHidden = overlayEl.classList.contains("hidden");
  var wasExiting = !!overlayEl._cgModalExiting;
  overlayEl._cgModalExiting = false;
  overlayEl.classList.remove("hidden");
  if (overlayEl.classList.contains("cg-modal-veil")) {
    // Next frame: display:none needs a painted start value; `gen` guards a same-frame dismiss.
    requestAnimationFrame(function () {
      if (overlayEl._cgModalGen === gen) overlayEl.classList.add("is-veiled");
    });
  }
  // Re-pop only after hidden or a cancelled exit (which leaves the card filled invisible).
  if ((!wasHidden && !wasExiting) || !window.ClipgenMotion) return;
  if (cardEl) ClipgenMotion.animateIn(cardEl, "pop");
};

// `commit` does the visual hide only; logical cleanup stays synchronous at the call site.
var popModalOut = function (overlayEl, cardEl, commit) {
  if (!overlayEl) return;
  var gen = (overlayEl._cgModalGen = (overlayEl._cgModalGen || 0) + 1);
  var veiled = overlayEl.classList.contains("cg-modal-veil");
  overlayEl.classList.remove("is-veiled");
  if (!window.ClipgenMotion) {
    commit();
    return;
  }
  overlayEl._cgModalExiting = true;
  var done = function () {
    if (overlayEl._cgModalGen !== gen) return; // superseded by a re-open
    overlayEl._cgModalExiting = false;
    commit();
  };
  var cardExit = cardEl
    ? ClipgenMotion.animateOut(cardEl, "pop")
    : Promise.resolve();
  // The veil outlasts the card, so wait on it; reduced motion disables its transition.
  if (veiled && !ClipgenMotion.isReduced()) setTimeout(done, _cgVeilMs(overlayEl));
  else cardExit.then(done);
};

// Pop a modal open; traps focus when `onEscape` is given.
var openPopModal = function (overlayEl, cardEl, opts) {
  opts = opts || {};
  popModalIn(overlayEl, cardEl);
  if (opts.modalOpen) document.body.classList.add("modal-open");
  if (opts.onEscape) {
    openBlockingModal(overlayEl, {
      onEscape: opts.onEscape,
      trapFocus: true,
      restoreFocus: true,
    });
  }
};

// Reverse of openPopModal; `commit` runs with the visual hide.
var closePopModal = function (overlayEl, cardEl, opts, commit) {
  opts = opts || {};
  if (opts.releaseTrapNow) closeBlockingModal(overlayEl);
  popModalOut(overlayEl, cardEl, function () {
    if (!opts.releaseTrapNow) closeBlockingModal(overlayEl);
    overlayEl.classList.add("hidden");
    if (opts.modalOpen) document.body.classList.remove("modal-open");
    if (commit) commit();
  });
};

// ---- Mark categories ----
// Fallback mirroring config.MARK_CATEGORIES (tests/test_shared_constants.py); setMarkCategories() mutates it in place.

var MARK_CATEGORIES = {
  pain_point: { label: "Pain Point", color: "#dc2626" },
  delight:    { label: "Delight",    color: "#16a34a" },
  quote:      { label: "Quote",      color: "#2563eb" },
  insight:    { label: "Insight",    color: "#f97316" },
  task:       { label: "Task Issue", color: "#8b5cf6" },
  bookmark:   { label: "Bookmark",   color: "#0891b2" },
  friction:   { label: "Friction",   color: "#ea580c" },
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
// Colors color-mix the `--stream-*` tokens so theme changes propagate.

var XREF_BADGES = {
  screenspace: { icon: "squares-2x2", color: "color-mix(in srgb, var(--stream-screenspace) 85%, transparent)" },
  transcript:  { icon: "chat-bubble-bottom-center-text", color: "color-mix(in srgb, var(--stream-transcript) 85%, transparent)" },
  sheet:       { icon: "table-cells", color: "color-mix(in srgb, var(--stream-sheet) 85%, transparent)" },
  composer:    { icon: "scissors", color: "color-mix(in srgb, var(--color-accent) 85%, transparent)" },
  mindnode:    { icon: "share", color: "color-mix(in srgb, var(--stream-mindnode) 85%, transparent)" },
};

// Works from any /prefix/ page: each has a sibling /screenspace/icons/ route.
var XREF_ICON_BASE = "../screenspace/icons/";

var xrefBadgeIcon = function (iconName) {
  return iconMaskSpan(iconName, { className: "xref-badge-icon", basePath: XREF_ICON_BASE });
};

// Stacked source badges for a findOverlappingData() result; selfBadge { icon, color, title } goes first.
var buildXrefBadges = function (xref, selfSource, selfBadge) {
  if (!CLIPGEN_CONFIG.crossReferences) return null;
  var badges = [];
  if (selfBadge) badges.push(selfBadge);
  if (selfSource !== "screenspace" && xref.screenspaceEvents.length > 0) {
    var types = [];
    var seen = {};
    for (var i = 0; i < xref.screenspaceEvents.length; i++) {
      var et = xref.screenspaceEvents[i].event_type || xref.screenspaceEvents[i].detector;
      if (!seen[et]) { seen[et] = true; types.push(et); }
    }
    badges.push({ icon: XREF_BADGES.screenspace.icon, color: XREF_BADGES.screenspace.color, title: types.join(", ") });
  }
  if (selfSource !== "transcript" && xref.transcriptSnippets.length > 0) {
    var trTexts = [];
    for (var j = 0; j < xref.transcriptSnippets.length && j < 3; j++) {
      var t = xref.transcriptSnippets[j].text;
      trTexts.push(t.length > 80 ? t.substring(0, 80) + "…" : t);
    }
    badges.push({ icon: XREF_BADGES.transcript.icon, color: XREF_BADGES.transcript.color, title: trTexts.join("\n") });
  }
  if (xref.sheetObservations.length > 0) {
    var obsTexts = [];
    for (var k = 0; k < xref.sheetObservations.length && k < 3; k++) {
      obsTexts.push(xref.sheetObservations[k].observation);
    }
    badges.push({ icon: XREF_BADGES.sheet.icon, color: XREF_BADGES.sheet.color, title: obsTexts.join("\n") });
  }
  if (badges.length === 0) return null;
  var container = el("span", "xref-badge-stack");
  for (var b = 0; b < badges.length; b++) {
    var badge = el("span", "xref-badge");
    badge.style.background = badges[b].color;
    badge.style.zIndex = badges.length - b;
    badge.appendChild(xrefBadgeIcon(badges[b].icon));
    badge.title = badges[b].title;
    container.appendChild(badge);
  }
  return container;
};

// ---- Filter helpers (artifact grids in viewer) ----

// Sorted unique non-empty values of `field`; opts.trim trims strings first.
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

// The "all" option has value "" so callers can detect it.
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
// Source of truth: `--color-task-{type}` tokens in tokens.css; tests/test_shared_constants.py catches drift.

var DETECTOR_COLORS = {};
var _DETECTOR_TYPES = [
  "multitool", "color", "change", "similarity", "text",
  "numbers", "timelapse", "template", "shape", "flow", "scene", "inactivity",
  "boundary", "attention",
];
// Mirrors the dark-theme `--color-task-*` block in tokens.css; update both together.
var _DETECTOR_FALLBACK = {
  multitool: "#60a5fa", color: "#a78bfa", change: "#fb923c",
  similarity: "#22d3ee", text: "#34d399", numbers: "#facc15",
  timelapse: "#f472b6", template: "#fb7185", shape: "#f87171",
  flow: "#818cf8", scene: "#2dd4bf", inactivity: "#94a3b8",
  boundary: "#e879f9", attention: "#a3e635",
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

// Detector label → its `--color-task-*` token (color-mix when alpha < 1); null for unknown labels.
function detectorColor(label, alpha) {
  if (!label) return null;
  var key = String(label).toLowerCase().trim();
  if (_DETECTOR_TYPES.indexOf(key) === -1) return null;
  var v = "var(--color-task-" + key + ")";
  if (alpha == null || alpha >= 1) return v;
  var pct = Math.max(0, Math.min(100, Math.round(alpha * 100)));
  return "color-mix(in oklch, " + v + " " + pct + "%, transparent)";
}

// ---- Category hue palette ----
// Hue per non-detector label; detector entries are legacy, prefer detectorColor().
var CATEGORY_HUES = {
  multitool: 220, color: 280, change: 30, similarity: 200,
  text: 170, numbers: 330, timelapse: 350, template: 18,
  flow: 145, scene: 155, inactivity: 210, boundary: 300,
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

// AppKit fills resize-exposed frame from the window appearance; Light flashes white on dark pages.
var syncDesktopAppearance = function (theme) {
  if (!document.documentElement.dataset.desktopChrome) return;
  var send = function () {
    var api = window.pywebview && window.pywebview.api;
    if (api && typeof api.set_window_appearance === "function") {
      api.set_window_appearance(theme);
    }
  };
  // On first load the bridge arrives after this runs; pywebviewready signals it.
  if (window.pywebview && window.pywebview.api) send();
  else window.addEventListener("pywebviewready", send, { once: true });
};

var applyStoredThemePreference = function () {
  var stored = null;
  try { stored = window.localStorage.getItem(THEME_STORAGE_KEY); } catch (_) {}
  var theme = stored === "light" ? "light" : "dark";
  document.documentElement.setAttribute("data-theme", theme);
  updateThemeToggleButton(theme);
  syncDesktopAppearance(theme);
};

var toggleThemePreference = function () {
  var root = document.documentElement;
  var current = root.getAttribute("data-theme") || "dark";
  var next = current === "light" ? "dark" : "light";
  root.setAttribute("data-theme", next);
  try { window.localStorage.setItem(THEME_STORAGE_KEY, next); } catch (_) {}
  updateThemeToggleButton(next);
  syncDesktopAppearance(next);
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
    invalidateCanvasThemeColors();
    if (onToggle) onToggle();
  });
};

// ---- Shared /api/status ----
// Memoized: every hit scans directories server-side. force=true refetches.
var _clipgenStatusPromise = null;
var clipgenStatus = function (force) {
  if (force || !_clipgenStatusPromise) {
    _clipgenStatusPromise = fetch("/api/status").then(function (r) {
      if (!r.ok) throw new Error("HTTP " + r.status);
      return r.json();
    });
    // Drop a failed fetch so the next caller retries; also prevents an unhandled rejection.
    _clipgenStatusPromise.catch(function () {
      _clipgenStatusPromise = null;
    });
  }
  return _clipgenStatusPromise;
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

  clipgenStatus()
    .then(function (status) {
      var items = panel.querySelectorAll(".frontend-switcher-item");
      items.forEach(function (item) {
        var key = item.dataset.frontend;
        if (key && status[key] === false) item.classList.add("hidden");
      });
    })
    .catch(function () {});
};

// Reads a /api/settings save or reset payload; returns whether the value moved.
var applyCrossRefSetting = function (applied, settings) {
  var value;
  if (applied && applied.CROSS_REFERENCES_ENABLED !== undefined) {
    value = applied.CROSS_REFERENCES_ENABLED;
  } else if (settings) {
    for (var i = 0; i < settings.length; i++) {
      if (settings[i].name === "CROSS_REFERENCES_ENABLED") {
        value = settings[i].value;
        break;
      }
    }
  }
  if (value === undefined) return false;
  var changed = CLIPGEN_CONFIG.crossReferences !== !!value;
  CLIPGEN_CONFIG.crossReferences = !!value;
  return changed;
};

// Settings live at the combined-app root, not under the page prefix.
var setCrossReferences = function (enabled) {
  return apiPut("/api/settings", { settings: { CROSS_REFERENCES_ENABLED: enabled } })
    .then(function () {
      CLIPGEN_CONFIG.crossReferences = enabled;
      if (typeof window.clipgenRerenderCrossRefs === "function") {
        window.clipgenRerenderCrossRefs();
      }
    })
    .catch(function () {});
};

// ---- Participant deep links (location.hash) ----

// Part owning global second g; startKey: "cumulativeStart" (default) or "offset" (composer). Clamps past end.
var clipgenPartForGlobal = function (parts, g, startKey) {
  var k = startKey || "cumulativeStart";
  for (var i = 0; i < parts.length; i++) {
    if (g >= parts[i][k] && g < parts[i][k] + parts[i].duration) return i;
  }
  return Math.max(0, parts.length - 1);
};

// First of hashPid / currentId / storedId present in participants, else "".
var clipgenPickParticipant = function (participants, opts) {
  opts = opts || {};
  function present(id) {
    if (!id) return "";
    for (var i = 0; i < participants.length; i++) {
      if (participants[i].id === id) return id;
    }
    return "";
  }
  return present(opts.hashPid) || present(opts.currentId) || present(opts.storedId) || "";
};

// /transcripts/#P07 pre-selects a participant; any token passes so prefix rules stay in config.py; pages validate.
var clipgenHashParticipant = function () {
  var raw = (window.location.hash || "").replace(/^#/, "");
  if (!raw) return "";
  try { raw = decodeURIComponent(raw); } catch (_) {}
  return /^[A-Za-z][\w-]*$/.test(raw) ? raw : "";
};

// /overview/#tab=metadata deep links. The "=" never matches the participant hash pattern.
var clipgenHashTab = function () {
  var m = /^#tab=([\w-]+)$/.exec(window.location.hash || "");
  return m ? m[1] : "";
};

// ---- Local AI availability ----

// Classify /api/models' `llm` block: "ok", "missing" or "stopped". An unknown payload is "ok": never block.
var clipgenLlmStatus = function (llm) {
  var baseUrl = (llm && llm.base_url) || "localhost";
  var hint = (llm && llm.install_hint) || [];
  var base = { hint: hint, baseUrl: baseUrl };
  if (!llm) return { state: "ok", message: "", hint: hint, baseUrl: baseUrl };
  // Messages name the problem but prescribe no control; the three surfaces differ in affordances.
  if (llm.installed === false) {
    base.state = "missing";
    base.message = "The local AI runtime is not installed — summaries, citations and reports need it.";
    return base;
  }
  if (llm.available === false) {
    base.state = "stopped";
    base.message = "The AI server is not running at " + baseUrl + ".";
    return base;
  }
  base.state = "ok";
  base.message = "";
  return base;
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

// One key inside a map-valued stored field (videoTimeByParticipant, tabByParticipant, ...).
var getStoredUIMapEntry = function (page, field, key, fallback) {
  var map = getStoredUIState(page)[field];
  return map && typeof map === "object" && Object.prototype.hasOwnProperty.call(map, key)
    ? map[key]
    : fallback;
};

var setStoredUIMapEntry = function (page, field, key, value) {
  var st = getStoredUIState(page);
  var map = st[field] && typeof st[field] === "object" ? st[field] : {};
  map[key] = value;
  setStoredUIStateField(page, field, map);
};

// ---- Canvas helpers (timeline overlays) ----

// Stacked per-series bands normalized to their own peaks; dimKey paints last; colors are #rrggbb.
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

// ---- Timeline ruler core (shared by canvas timeline surfaces) ----

// "Nice" tick intervals (seconds) for a timeline ruler, coarse → fine.
var TIMELINE_TICK_STEPS = [1, 2, 5, 10, 15, 30, 60, 120, 300, 600, 900, 1800, 3600];

// maxTicks: largest step with <= N ticks; targetTicks: smallest step giving ~N ticks (default 8).
var niceTimeInterval = function (visibleSeconds, opts) {
  opts = opts || {};
  var steps = TIMELINE_TICK_STEPS;
  var i;
  if (opts.maxTicks) {
    for (i = 0; i < steps.length; i++) {
      if (visibleSeconds / steps[i] <= opts.maxTicks) return steps[i];
    }
    return steps[steps.length - 1];
  }
  var target = visibleSeconds / (opts.targetTicks || 8);
  for (i = 0; i < steps.length; i++) {
    if (steps[i] >= target) return steps[i];
  }
  return steps[steps.length - 1];
};

// Ruler ticks and labels only; markers, bands and playheads stay per-surface.
var drawTimelineRuler = function (ctx, opts) {
  var interval = opts.interval;
  if (!(interval > 0)) return;
  var visStart = opts.visStart;
  var visEnd = opts.visEnd;
  var timeToX = opts.timeToX;
  var fmt = opts.format || formatTime;
  var tickH = opts.tickHeight != null ? opts.tickHeight : 6;
  var labelY = opts.labelY != null ? opts.labelY : 16;
  var c = opts.colors || {};

  ctx.strokeStyle = c.border;
  ctx.fillStyle = c.textDim;
  ctx.font = "10px " + (c.fontMono || "monospace");
  ctx.textAlign = "center";
  ctx.lineWidth = 1;
  var firstTick = Math.ceil(visStart / interval) * interval;
  for (var t = firstTick; t <= visEnd; t += interval) {
    var x = timeToX(t);
    ctx.beginPath();
    ctx.moveTo(x, 0);
    ctx.lineTo(x, tickH);
    ctx.stroke();
    ctx.fillText(fmt(t), x, labelY);
  }
  ctx.textAlign = "start";
};

// ---- Video helpers ----

// Hidden tabs drop paused <video> frames; snapshot to canvas until repaint. Positioned parent required.
var clipgenInstallPausedFrameOverlay = function (video) {
  if (!video || video._clipgenPausedOverlay) return;
  var parent = video.parentNode;
  if (!parent) return;

  var canvas = document.createElement("canvas");
  canvas.className = "video-paused-overlay";
  // Inline styles so the helper works without page-specific CSS.
  canvas.style.position = "absolute";
  canvas.style.inset = "0";
  canvas.style.width = "100%";
  canvas.style.height = "100%";
  canvas.style.objectFit = "contain";
  canvas.style.pointerEvents = "none";
  canvas.style.display = "none";
  parent.appendChild(canvas);
  video._clipgenPausedOverlay = canvas;

  var hide = function () { canvas.style.display = "none"; };

  var snapshot = function () {
    if (!video.src || !video.paused) return;
    var w = video.videoWidth, h = video.videoHeight;
    // videoWidth/Height are zero until the first frame decodes.
    if (!w || !h) return;
    canvas.width = w;
    canvas.height = h;
    try {
      canvas.getContext("2d").drawImage(video, 0, 0, w, h);
      canvas.style.display = "";
    } catch (_) {
      // Cross-origin or other draw failure: leave the overlay hidden.
    }
  };

  // The live video reasserts itself: drop the snapshot.
  video.addEventListener("play", hide);
  video.addEventListener("seeked", hide);
  video.addEventListener("emptied", hide);
  video.addEventListener("loadedmetadata", hide);

  document.addEventListener("visibilitychange", function () {
    if (document.hidden) {
      snapshot();
    } else if (video.paused && video.src) {
      // Nudge currentTime so `seeked` hides the snapshot; same-value assignment may be optimized away.
      var t = video.currentTime;
      video.currentTime = t > 0.001 ? t - 0.001 : 0.001;
    }
  });
};


// ---- Drag-to-resize handles ----

// rAF-throttled mouse/touch drag along `axis`; cfg.onStart() may return false to refuse.
function initDragHandle(handle, axis, cfg) {
  if (!handle) return;
  var dragging = false;
  var start = 0;
  var rafPending = false;

  function coord(e) {
    var touch = e.touches && e.touches[0];
    if (axis === "x") return e.clientX || (touch && touch.clientX) || 0;
    return e.clientY || (touch && touch.clientY) || 0;
  }

  function onDown(e) {
    if (cfg.onStart && cfg.onStart() === false) return;
    e.preventDefault();
    if (cfg.stopPropagation) e.stopPropagation();
    dragging = true;
    start = coord(e);
    handle.classList.add("active");
    document.body.style.cursor = cfg.cursor || (axis === "x" ? "col-resize" : "row-resize");
    document.body.style.userSelect = "none";
  }

  function onMove(e) {
    if (!dragging || rafPending) return;
    rafPending = true;
    var now = coord(e);
    requestAnimationFrame(function () {
      cfg.onDelta(now - start);
      rafPending = false;
    });
  }

  function onUp() {
    if (!dragging) return;
    dragging = false;
    handle.classList.remove("active");
    document.body.style.cursor = "";
    document.body.style.userSelect = "";
    if (cfg.onEnd) cfg.onEnd();
  }

  handle.addEventListener("mousedown", onDown);
  handle.addEventListener("touchstart", onDown, { passive: false });
  document.addEventListener("mousemove", onMove);
  document.addEventListener("touchmove", onMove, { passive: false });
  document.addEventListener("mouseup", onUp);
  document.addEventListener("touchend", onUp);

  if (cfg.onToggle) {
    handle.addEventListener("dblclick", function (e) {
      e.preventDefault();
      if (cfg.stopPropagation) e.stopPropagation();
      cfg.onToggle();
    });
  }
}

// #panelDivider drag + dblclick for Studio and Screenspace; page specifics arrive as cfg callbacks.
function initPanelDivider(cfg) {
  var startHeight = 0;
  var bounds = { min: 0, max: 0 };
  initDragHandle(document.querySelector("#panelDivider"), "y", {
    onStart: function () {
      if (cfg.isCollapsed()) return false;
      startHeight = cfg.getHeight();
      bounds = cfg.getBounds();
      if (cfg.onDragStart) cfg.onDragStart();
      return true;
    },
    // Dragging up (negative delta) grows the bottom panel.
    onDelta: function (delta) {
      cfg.setHeight(Math.max(bounds.min, Math.min(bounds.max, startHeight - delta)));
    },
    onEnd: function () {
      if (cfg.onDragEnd) cfg.onDragEnd();
      if (cfg.persist) cfg.persist();
    },
    onToggle: cfg.onToggle,
  });
}
