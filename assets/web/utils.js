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
  annotations: [
    { id: "key", token: "!key" },
  ],
  ignoredTimestampTokens: ["x"],
  screenspaceOcrMinConfidence: 0.7,
  screenspaceMultitoolMaxOffset: 30,
  screenspaceMaskFallbackTools: ["similarity", "inactivity", "boundary", "timelapse"],
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
  composerAnnotationStrokeWidth: 0.004,
  composerAnnotationFontSize: 0.035,
  composerAnnotationSpanSeconds: 10.0,
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

// English-only count + noun phrase (1 → singular, else plural).
var clipgenPluralUnit = function (n, singular, plural) {
  return n + " " + (n === 1 ? singular : plural);
};

// Populate a container with a skeleton table: `cols` header cells plus
// `rows × cols` body cells. Pairs with the `.skeleton-grid` / `.skeleton-cell`
// CSS in primitives.css (and the shimmer in tokens.css). Target should
// already have class `skeleton-grid`; the helper appends cells via a single
// DocumentFragment.
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
// is a CSS url(...) string such as "url('icons/check.svg')" or
// "url('/screenspace/icons/eye.svg')". The element's `mask-size`,
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

// ---- Unified icon-mask helper family ----
//
// All build on applyMaskIcon / maskIconStyle above. `name` is the basename of a
// file in an icons/ directory (no extension); `basePath` defaults to "icons/"
// (the common relative case). Pages with a different icon route pass their own
// (e.g. "/screenspace/icons/", "../screenspace/icons/").

// Full CSS url(...) string for an icon basename. Single-quoted so the result is
// safe to embed inside a double-quoted HTML style attribute (e.g.
// `style="..."`); a double-quoted url() would terminate the attribute early.
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

// Apply mask-image to every matching element within `scope`.
// opts = { selector, basePath }; selector defaults to "[data-icon]" and the
// icon name is read from each node's data-icon attribute.
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

// ---- Hover tooltips (dark pill, pairs with .cg-tooltip in tokens.css) ----

var createTooltip = function (opts) {
  opts = opts || {};
  var cls = "cg-tooltip";
  if (opts.multiline) cls += " cg-tooltip--multiline";
  var tip = document.createElement("div");
  tip.className = cls;
  document.body.appendChild(tip);
  return {
    el: tip,
    show: function (anchor, text) {
      tip.textContent = text || "";
      // Set text before measuring so offsetHeight reflects final size.
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

// Singleton tooltip driven by [data-tooltip] attributes anywhere on the page.
// Replaces the per-page CSS ::after pseudo-element variants so positioning
// goes through positionTooltipAnchored and stays inside the viewport. The
// pointer-events:auto override in tokens.css ensures Chrome/Safari still
// dispatch mouseover events for disabled <button data-tooltip>.
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
  var showFor = function (el) {
    if (shouldSuppress(el)) return;
    var text = el.getAttribute("data-tooltip");
    if (!text) return;
    current = el;
    ensureTip().show(el, text);
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
    // relatedTarget is what the cursor moved onto. If still inside the
    // anchor, keep the tooltip — mouseout fires when crossing into children.
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

// ---- Elapsed-time / ETA estimation for long-running operations ----

// Linearly extrapolate remaining seconds from elapsed time and a 0..1 progress
// fraction. Returns null when no honest estimate is possible (progress not inside
// the open interval (0, 1), or no time elapsed yet) — callers render elapsed only.
var estimateRemainingSec = function (elapsedSec, progress) {
  if (progress == null || !isFinite(progress) || progress <= 0 || progress >= 1) return null;
  if (!isFinite(elapsedSec) || elapsedSec <= 0) return null;
  return (elapsedSec * (1 - progress)) / progress;
};

// Track elapsed time and a smoothed remaining-time estimate for one long-running
// operation. Frontend-only; fed a 0..1 progress fraction on each update. For
// indeterminate operations (no progress signal) callers omit progress and read
// elapsedSec only — remainingSec stays null. Returns a plain object (not a class,
// per project convention) closing over the start timestamp and an EMA of the raw
// estimate to damp jitter.
//
// pause()/resume() are opt-in: elapsed freezes while paused and continues afterward,
// with the paused span excluded from elapsed. Callers that never call pause() see the
// original continuous wall-clock behavior (so e.g. the Transcripts timers are
// unaffected).
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
    // Idempotent: repeated calls keep the original start so elapsed never resets.
    // Pass an explicit epoch (ms) to seed from a known start (e.g. reattach).
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
    // Returns { elapsedSec, remainingSec } — remainingSec is null until the
    // progress/elapsed gates open, then an EMA-smoothed, rounded estimate.
    update: function (progress) {
      if (startMs == null) startMs = Date.now();
      // While paused, "now" holds at pausedAt (the current pause isn't in pausedMs
      // yet), so elapsed freezes; after resume() it continues seamlessly.
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

// Friendly "~1:20 left" label for a remaining-seconds estimate. Buckets the value
// before formatting so the text doesn't twitch every tick. Returns "" when no
// estimate is available.
var formatEtaLabel = function (remainingSec) {
  if (remainingSec == null || !isFinite(remainingSec) || remainingSec < 0) return "";
  var bucketed;
  if (remainingSec < 60) bucketed = Math.round(remainingSec / 5) * 5;
  else if (remainingSec < 300) bucketed = Math.round(remainingSec / 10) * 10;
  else bucketed = Math.round(remainingSec / 30) * 30;
  if (bucketed <= 0) bucketed = 5;
  return "~" + formatDuration(bucketed) + " left";
};

// Drive *tickFn* on a fixed interval (default 1s) while a job is active, with
// optional pause-when-hidden. ensure() starts the timer (idempotent); each tick
// first checks the guards and self-stops when isActive() returns false (or, with
// gateHidden, when the tab is hidden), so pages don't re-implement the
// setInterval/clearInterval lifecycle. Returns a plain object (not a class, per
// project convention).
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

// Canonical hex <-> rgb helpers (previously duplicated in screenspace-utils.js
// and color-picker.js). hexToRgb accepts "#rgb"/"#rrggbb" (with or without the
// hash) and returns null on anything else; rgbToHex clamps + rounds so it is
// safe on float channel values from HSV sliders.
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

// Cached snapshot of the theme tokens that canvas renderers (Screenspace
// timeline, Transcripts ruler, Studio heatmap) read on every frame. The cache
// is invalidated automatically when initThemeToggle()'s click handler fires;
// callers can also invalidate manually via invalidateCanvasThemeColors().
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
  // Generation token: the toast is a reused element, so a fade-out that started
  // before a re-show must not add `.hidden` when it settles onto the fresh toast.
  var gen = (toastEl._toastGen = (toastEl._toastGen || 0) + 1);
  // Fade in on show; the newer entry animation supersedes any stale exit fill.
  // Guarded because motion.js only loads on some pages (elsewhere the toast snaps).
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

// Returns a .catch handler that surfaces the failure as a toast. For
// user-initiated loads and mutations whose silent failure would leave the
// UI wrong; background polling/preloading keeps its silent catches.
// Usage: apiGet("api/regions").then(...).catch(toastError("Failed to load regions"));
var toastError = function (prefix) {
  return function (err) {
    showToast(prefix + (err && err.message ? ": " + err.message : ""));
  };
};

// ---- Polling ----

var POLL_INTERVAL = 3000;

// Generic poller. Encapsulates the recurring `setInterval` + `visibilitychange`
// + `document.hidden` dance used by Studio intake counters, Screenspace task
// status, and similar live-refresh loops.
//
//   var poller = createPoller(fn, ms, opts);
//   poller.start();   // arm; pauses automatically when tab hidden
//   poller.stop();    // disarm; safe to call multiple times
//
// Options:
//   pauseWhenHidden (default true)  — pause when document.hidden, resume on
//                                     visibilitychange and run fn() once to
//                                     catch up.
//   runImmediately  (default true)  — run fn() once on start (and again on
//                                     resume) before the next interval tick.
//   maxIntervalMs   (default = intervalMs) — enable idle backoff. When set
//                                     above intervalMs the loop self-reschedules
//                                     with setTimeout instead of setInterval:
//                                     fn's resolved value is an "active this
//                                     tick" signal — truthy resets to the base
//                                     interval, falsy backs the delay off toward
//                                     maxIntervalMs after backoffAfter quiet
//                                     ticks (e.g. 5s → 10 → 20 → 30, capped).
//   backoffAfter    (default 3)     — consecutive quiet ticks before backing off.
//
// In backoff mode fn may return a value or a Promise; the loop waits for it
// before scheduling the next tick (so slow polls never overlap). The returned
// object also gains wake(): snap back to the base cadence and refresh now —
// call it after a user action so polling both refreshes and speeds back up.
//
// fn exceptions are swallowed so a transient error does not kill the loop.
var createPoller = function (fn, intervalMs, opts) {
  opts = opts || {};
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
  function runAdaptive() {
    timer = null;
    inFlight = true;
    var result;
    try { result = fn(); } catch (_) { result = false; }
    Promise.resolve(result).then(
      function (active) { applySignal(!!active); },
      function () { applySignal(false); }
    ).then(function () {
      inFlight = false;
      if (!wantRunning || hidden()) { pendingWake = false; return; }
      // A wake() landed mid-flight: that response may predate the user's
      // mutation, so run one fresh poll now instead of waiting a full tick.
      if (pendingWake) {
        pendingWake = false;
        currentDelay = intervalMs;
        quiet = 0;
        runAdaptive();
        return;
      }
      schedule();
    });
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
    wake: function () {
      if (!wantRunning || hidden()) return;
      if (adaptive) {
        currentDelay = intervalMs;
        quiet = 0;
        // A poll is already running but may predate this action's server-side
        // effect — queue one fresh poll for when it completes.
        if (inFlight) { pendingWake = true; return; }
        disarm();
        runAdaptive();
      } else {
        disarm();
        safeFn();
        timer = setInterval(safeFn, intervalMs);
      }
    },
  };
};

// ---- SSE stream with standard parse + fallback hook ----
// Open an EventSource with the project's standard JSON-parse onmessage wrapper
// and auto-close-on-error. The caller owns its polling fallback — each subscriber
// stores its own stream/poller and reacts to a drop differently — so onError
// fires AFTER the stream is closed. Returns the EventSource (or null if the
// browser lacks EventSource, in which case onUnsupported runs).
//
// Options:
//   onMessage(data)   — called with the parsed JSON of each message
//   onOpen()          — called when the connection (re)opens
//   onError()         — called once the dropped stream has been closed
//   onUnsupported()   — called instead of opening when window.EventSource is absent
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

// ---- Blocking modal lifecycle (Escape / backdrop / optional focus trap) ----
// Shared lifecycle for blocking overlays: closes on Escape, optionally on
// backdrop click, optionally traps Tab/Shift+Tab inside the overlay and restores
// focus to the trigger on release. Singleton — opening a new modal releases the
// previous one (callers never stack blocking overlays). The returned trap's
// release() is idempotent cleanup-only, so any dismiss path (button, backdrop,
// Escape) can call closeBlockingModal safely.
//
// Options:
//   onEscape()        — fired on Escape
//   onBackdropClick() — fired when the overlay element itself is clicked
//   trapFocus         — keep Tab/Shift+Tab inside the overlay and focus its first
//                       control on open
//   restoreFocus      — restore focus to the prior element on release
var _TRAP_FOCUSABLE =
  'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])';
var _activeBlockingModal = null;

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
    // Re-query each Tab — overlay button visibility can change between phases
    // (e.g. a status overlay's in-progress vs result state).
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
    if (prevFocus && prevFocus.focus) prevFocus.focus();
  }

  var trap = { el: overlayEl, opts: opts, release: release };
  document.addEventListener("keydown", onKey, true);
  if (opts.onBackdropClick) overlayEl.addEventListener("click", onClick);
  if (opts.trapFocus) {
    var initial = visibleFocusable();
    if (initial.length) initial[0].focus();
  }
  _activeBlockingModal = trap;
  return trap;
};

var closeBlockingModal = function (overlayEl) {
  if (_activeBlockingModal && _activeBlockingModal.el === overlayEl) {
    _activeBlockingModal.release();
  }
};

// Whether any blocking-modal trap is active. openBlockingModal is a
// singleton, so anything that opens on a chord (the command palette) must
// check this first — opening would release the existing overlay's trap and
// leave it visible with Escape/backdrop dismiss dead.
var hasBlockingModal = function () {
  return _activeBlockingModal !== null;
};

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
// Icon names reference files in assets/icons/; rendered via CSS mask-image.
// Colors use color-mix on the canonical `--stream-*` tokens in tokens.css so
// any theme change or token tweak propagates automatically (color-mix yields a
// CSS color string usable both in inline style and JS .style assignment).

var XREF_BADGES = {
  screenspace: { icon: "squares-2x2", color: "color-mix(in srgb, var(--stream-screenspace) 85%, transparent)" },
  transcript:  { icon: "chat-bubble-bottom-center-text", color: "color-mix(in srgb, var(--stream-transcript) 85%, transparent)" },
  sheet:       { icon: "table-cells", color: "color-mix(in srgb, var(--stream-sheet) 85%, transparent)" },
  composer:    { icon: "scissors", color: "color-mix(in srgb, var(--color-accent) 85%, transparent)" },
};

// Relative icon base works from any /prefix/ page — every served page has a
// sibling /screenspace/icons/ route.
var XREF_ICON_BASE = "../screenspace/icons/";

var xrefBadgeIcon = function (iconName) {
  return iconMaskSpan(iconName, { className: "xref-badge-icon", basePath: XREF_ICON_BASE });
};

// Stacked source badges for a findOverlappingData() result. Used by Studio's
// intake cards and the Overview page (Convergence detail rows, Map drill-down).
// selfBadge: optional { icon, color, title } to prepend as the "self" source badge.
var buildXrefBadges = function (xref, selfSource, selfBadge) {
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
  "boundary",
];
// Values mirrored from the dark-theme `--color-task-*` block in tokens.css.
// Update this map and tokens.css together when changing a detector palette.
var _DETECTOR_FALLBACK = {
  multitool: "#60a5fa", color: "#a78bfa", change: "#fb923c",
  similarity: "#22d3ee", text: "#34d399", numbers: "#facc15",
  timelapse: "#f472b6", template: "#fb7185", flow: "#818cf8",
  scene: "#2dd4bf", inactivity: "#94a3b8", boundary: "#e879f9",
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
    invalidateCanvasThemeColors();
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

// ---- Participant deep links (location.hash) ----

// /transcripts/#P07 and /screenspace/#P07 pre-select that participant on
// load (set by the Overview Map's explain-panel links). Accepts any simple
// token so participant-prefix rules stay in config.py alone; the consuming
// page validates the id against its actual participant list.
var clipgenHashParticipant = function () {
  var raw = (window.location.hash || "").replace(/^#/, "");
  if (!raw) return "";
  try { raw = decodeURIComponent(raw); } catch (_) {}
  return /^[A-Za-z][\w-]*$/.test(raw) ? raw : "";
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

// ---- Timeline ruler core (shared by canvas timeline surfaces) ----

// "Nice" tick intervals (seconds) for a timeline ruler, coarse → fine.
var TIMELINE_TICK_STEPS = [1, 2, 5, 10, 15, 30, 60, 120, 300, 600, 900, 1800, 3600];

// Pick a tick interval (seconds) for a ruler spanning `visibleSeconds`. Two
// strategies, matching the two canvas surfaces' historical behavior:
//   { maxTicks: N }    largest step that keeps the tick count at/under N
//                      (Screenspace's zoomable ruler — was `<= 20`).
//   { targetTicks: N } smallest step giving roughly N ticks, i.e. step >=
//                      visibleSeconds / N (Transcripts' fixed-extent ruler — was N=8).
// Falls back to the coarsest step (3600s) when nothing fits.
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

// Draw the tick marks + time labels of a timeline ruler onto a 2D canvas.
// Markers, bands, playheads, and shading stay per-surface; this is only the ruler.
//
// opts:
//   visStart, visEnd   visible time window (seconds); ticks are drawn from the
//                      first multiple of `interval` at/after visStart up to visEnd
//   interval           tick spacing in seconds (see niceTimeInterval)
//   timeToX            fn(seconds) -> x pixel
//   colors             { border, textDim, fontMono }
//   tickHeight         tick line length in px (default 6)
//   labelY             baseline y for the label text (default 16)
//   format             fn(seconds) -> label string (default formatTime)
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

// Bridge the "paused <video> goes blank when the tab is hidden" gap.
//
// Chrome / Safari (and Edge, same engine as Chrome) will release the
// decoded frame buffer — and sometimes the entire compositor texture — of
// paused <video> elements while the tab is hidden, to free GPU memory. On
// return, the element shows nothing until the next play or seek manages to
// re-establish the surface. A plain `currentTime` nudge is unreliable
// because the seek can land before the compositor has re-attached, and
// because `readyState` may have dropped below HAVE_CURRENT_DATA while
// hidden.
//
// Workaround: snapshot the current frame onto a canvas overlay just before
// the browser has a chance to discard it (i.e. on visibilitychange→hidden
// while paused), then keep the overlay visible until the live video proves
// it has a fresh frame again (via the `seeked` / `play` / `loadedmetadata`
// events). Same architectural trick Screenspace uses for its paused state,
// except the pixels come from the live video instead of a server PNG.
//
// Install once per <video>. The helper attaches its own listeners (visibility,
// play, seeked, emptied, loadedmetadata) and is idempotent — repeated calls
// no-op.
//
// Requirements: the video's parent must be a positioned container (the
// overlay is absolutely positioned inside it). `#videoFrame` in transcripts
// already is `position: relative`.
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
    // videoWidth/Height are only non-zero once the first frame has decoded;
    // skip if we have nothing to paint.
    if (!w || !h) return;
    canvas.width = w;
    canvas.height = h;
    try {
      canvas.getContext("2d").drawImage(video, 0, 0, w, h);
      canvas.style.display = "";
    } catch (_) {
      // Cross-origin or other draw failure — leave overlay hidden, the page
      // is no worse off than it was before this helper existed.
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
      // Nudge currentTime to coax the video into producing a fresh frame;
      // the `seeked` listener above will then hide the snapshot. We use a
      // ~1 ms back-step (well under one frame at any reasonable framerate)
      // because same-value assignment is sometimes optimized away.
      var t = video.currentTime;
      video.currentTime = t > 0.001 ? t - 0.001 : 0.001;
    }
  });
};


// ---- Bottom-panel drag-to-resize divider ----

// Shared drag + dblclick wiring for the #panelDivider handle, used by Studio
// and Screenspace. The pages differ only in how they read/apply panel height,
// compute bounds, and persist — supplied as callbacks in `cfg`:
//   isCollapsed() -> bool   getHeight() -> px   setHeight(px)
//   getBounds() -> { min, max }   onToggle()   [onDragStart] [onDragEnd] [persist]
function initPanelDivider(cfg) {
  var handle = document.querySelector("#panelDivider");
  if (!handle) return;
  var dragging = false;
  var startY = 0;
  var startHeight = 0;
  var minH = 0;
  var maxH = 0;
  var rafPending = false;

  function onDown(e) {
    if (cfg.isCollapsed()) return;
    e.preventDefault();
    dragging = true;
    startY = e.clientY || (e.touches && e.touches[0].clientY) || 0;
    startHeight = cfg.getHeight();
    var bounds = cfg.getBounds();
    minH = bounds.min;
    maxH = bounds.max;
    handle.classList.add("active");
    if (cfg.onDragStart) cfg.onDragStart();
    document.body.style.cursor = "row-resize";
    document.body.style.userSelect = "none";
  }

  function onMove(e) {
    if (!dragging || rafPending) return;
    rafPending = true;
    var clientY = e.clientY || (e.touches && e.touches[0].clientY) || 0;
    requestAnimationFrame(function () {
      var delta = startY - clientY;
      cfg.setHeight(Math.max(minH, Math.min(maxH, startHeight + delta)));
      rafPending = false;
    });
  }

  function onUp() {
    if (!dragging) return;
    dragging = false;
    handle.classList.remove("active");
    if (cfg.onDragEnd) cfg.onDragEnd();
    document.body.style.cursor = "";
    document.body.style.userSelect = "";
    if (cfg.persist) cfg.persist();
  }

  handle.addEventListener("mousedown", onDown);
  handle.addEventListener("touchstart", onDown, { passive: false });
  document.addEventListener("mousemove", onMove);
  document.addEventListener("touchmove", onMove, { passive: false });
  document.addEventListener("mouseup", onUp);
  document.addEventListener("touchend", onUp);

  handle.addEventListener("dblclick", function (e) {
    e.preventDefault();
    cfg.onToggle();
  });
}
