/* dev-token-tweak.js — Pass 0 of FRONTEND-PLAN.md
 *
 * Floating dev-only widget for tweaking design tokens (CSS custom properties)
 * live in the browser. Persists overrides in localStorage, applies them via
 * inline style on <html>, exports a paste-ready :root snippet.
 *
 * Mirrors the theme-toggle precedent in utils.js:567-626 for storage and
 * hook-into-theme-changes patterns.
 *
 * Excluded from exports: viewer.py:_generate_viewer_html() strips any
 * <script data-dev-only> tag during inlining, so this never ships in
 * standalone exported HTML.
 */
(function () {
  "use strict";

  // Gated by the CLIPGEN_DEV_TOKEN_TWEAK feature flag in utils.js. Flip that
  // flag to false to disable the widget without removing the script tag.
  if (typeof window !== "undefined" && window.CLIPGEN_DEV_TOKEN_TWEAK === false) return;

  var STORAGE_KEY = "clipgen-token-overrides";
  var COLLAPSED_KEY = "clipgen-token-tweak-collapsed";

  // Token discovery scope — only redesign-tunable categories. Categorical
  // tokens (severity, content type, task colors, regions, intake categories)
  // are intentionally excluded.
  var INCLUDED_PREFIXES = [
    "--space-",
    "--text-",
    "--radius-",
    "--shadow-",
    "--duration-",
    "--layout-",
    "--sidebar-",
    "--bp-",
    "--button-",
    "--card-width-",
    "--bottom-panel-",
    "--icon-size-",
  ];
  var INCLUDED_THEME_TOKENS = {
    "--color-bg": 1,
    "--color-surface": 1,
    "--color-surface-alt": 1,
    "--color-text": 1,
    "--color-text-dim": 1,
    "--color-accent": 1,
    "--color-accent-hover": 1,
    "--color-accent-highlight": 1,
    "--color-accent-strong-hover": 1,
    "--color-border": 1,
    "--color-selected": 1,
    "--color-grid": 1,
    "--color-panel-bg": 1,
    "--color-panel-border": 1,
    "--color-panel-shadow": 1,
    "--font-mono": 1,
  };

  var GROUPS = [
    { id: "layout", label: "Layout", prefixes: ["--layout-", "--sidebar-", "--card-width-", "--bottom-panel-", "--bp-"] },
    { id: "spacing", label: "Spacing", prefixes: ["--space-"] },
    { id: "typography", label: "Typography", prefixes: ["--text-", "--font-mono"] },
    { id: "density", label: "Density", prefixes: ["--button-", "--icon-size-"] },
    { id: "radius", label: "Radius", prefixes: ["--radius-"] },
    { id: "shadow", label: "Shadow", prefixes: ["--shadow-"] },
    { id: "duration", label: "Duration", prefixes: ["--duration-"] },
    { id: "theme", label: "Theme colors", themeOnly: true },
  ];

  // ---- Persistence ----
  function loadOverrides() {
    try {
      var raw = window.localStorage.getItem(STORAGE_KEY);
      if (!raw) return {};
      var parsed = JSON.parse(raw);
      return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
    } catch (_) {
      return {};
    }
  }

  function saveOverrides(map) {
    try {
      if (Object.keys(map).length === 0) {
        window.localStorage.removeItem(STORAGE_KEY);
      } else {
        window.localStorage.setItem(STORAGE_KEY, JSON.stringify(map));
      }
    } catch (_) {}
  }

  function loadCollapsed() {
    try { return window.localStorage.getItem(COLLAPSED_KEY) === "1"; } catch (_) { return false; }
  }

  function saveCollapsed(collapsed) {
    try { window.localStorage.setItem(COLLAPSED_KEY, collapsed ? "1" : "0"); } catch (_) {}
  }

  // ---- Apply overrides ----
  function applyOverride(name, value) {
    document.documentElement.style.setProperty(name, value);
  }

  function clearOverride(name) {
    document.documentElement.style.removeProperty(name);
  }

  function applyAll(map) {
    Object.keys(map).forEach(function (name) { applyOverride(name, map[name]); });
  }

  // Apply stored overrides synchronously at script load so that page JS reading
  // tokens at init time sees the override values.
  var state = {
    overrides: loadOverrides(),
    tokens: [],
    referenced: {},
    panel: null,
    collapsed: loadCollapsed(),
  };
  applyAll(state.overrides);

  // ---- Token discovery ----
  function isIncluded(name) {
    if (INCLUDED_THEME_TOKENS[name]) return true;
    for (var i = 0; i < INCLUDED_PREFIXES.length; i++) {
      if (name.indexOf(INCLUDED_PREFIXES[i]) === 0) return true;
    }
    return false;
  }

  function tokenGroupId(name) {
    if (INCLUDED_THEME_TOKENS[name]) return "theme";
    for (var g = 0; g < GROUPS.length; g++) {
      var grp = GROUPS[g];
      if (!grp.prefixes) continue;
      for (var p = 0; p < grp.prefixes.length; p++) {
        if (name.indexOf(grp.prefixes[p]) === 0) return grp.id;
      }
    }
    return "other";
  }

  function discoverTokens() {
    var seen = {};
    var sheets = document.styleSheets;
    for (var s = 0; s < sheets.length; s++) {
      var rules;
      try { rules = sheets[s].cssRules; } catch (_) { continue; }
      if (!rules) continue;
      for (var r = 0; r < rules.length; r++) {
        var rule = rules[r];
        if (!rule.style || !rule.selectorText) continue;
        var sel = rule.selectorText;
        var isRoot = sel === ":root" || sel.indexOf(":root") === 0 || sel.indexOf("html[data-theme") === 0;
        if (!isRoot) continue;
        for (var p = 0; p < rule.style.length; p++) {
          var prop = rule.style[p];
          if (prop.indexOf("--") !== 0) continue;
          if (!isIncluded(prop)) continue;
          seen[prop] = true;
        }
      }
    }
    return Object.keys(seen).sort();
  }

  // Walk every rule (including nested @media/@supports) and collect every
  // var(--token-name) reference. Used to flag tokens with no consumers in
  // the page CSS — those are the redesign-seeded tokens whose Pass-N
  // migration hasn't landed yet, so tweaking them does nothing visible.
  function findReferencedTokens() {
    var referenced = {};
    var sheets = document.styleSheets;
    for (var s = 0; s < sheets.length; s++) {
      var rules;
      try { rules = sheets[s].cssRules; } catch (_) { continue; }
      if (!rules) continue;
      scanRules(rules, referenced);
    }
    return referenced;
  }

  function scanRules(rules, referenced) {
    for (var r = 0; r < rules.length; r++) {
      var rule = rules[r];
      if (rule.cssRules) scanRules(rule.cssRules, referenced);
      if (!rule.style) continue;
      for (var p = 0; p < rule.style.length; p++) {
        var prop = rule.style[p];
        var val = rule.style.getPropertyValue(prop);
        if (!val || val.indexOf("var(") === -1) continue;
        var re = /var\(\s*(--[a-zA-Z0-9_-]+)/g;
        var m;
        while ((m = re.exec(val)) !== null) referenced[m[1]] = true;
      }
    }
  }

  function getDefaultValue(name) {
    // Read computed value with this widget's inline override temporarily
    // removed so we always know the *underlying* token value.
    var elt = document.documentElement;
    var inlineSaved = elt.style.getPropertyValue(name);
    var inlinePrioSaved = elt.style.getPropertyPriority(name);
    if (inlineSaved !== "") elt.style.removeProperty(name);
    var v = getComputedStyle(elt).getPropertyValue(name).trim();
    if (inlineSaved !== "") elt.style.setProperty(name, inlineSaved, inlinePrioSaved);
    return v;
  }

  // ---- Control inference ----
  function inferControlType(value) {
    if (!value) return "text";
    var v = value.trim();
    if (/^#[0-9a-fA-F]{3,8}$/.test(v)) return "color";
    if (/^rgb\(/i.test(v) || /^hsl\(/i.test(v)) return "color";
    if (/^rgba\(/i.test(v) || /^hsla\(/i.test(v)) return "color-text";
    if (/^(-?[0-9]*\.?[0-9]+)(px|rem|em)$/.test(v)) return "length";
    if (/^(-?[0-9]*\.?[0-9]+)ms$/.test(v)) return "duration";
    if (/^-?[0-9]*\.?[0-9]+$/.test(v)) return "number";
    return "text";
  }

  function parseLength(v) {
    var m = v.match(/^(-?[0-9]*\.?[0-9]+)(px|rem|em)$/);
    if (!m) return null;
    return { num: parseFloat(m[1]), unit: m[2] };
  }

  function colorToHex(v) {
    if (!v) return "#000000";
    var s = v.trim();
    if (/^#[0-9a-fA-F]{6}$/.test(s)) return s.toLowerCase();
    if (/^#[0-9a-fA-F]{3}$/.test(s)) {
      return "#" + s[1] + s[1] + s[2] + s[2] + s[3] + s[3];
    }
    if (/^#[0-9a-fA-F]{8}$/.test(s)) return s.substring(0, 7).toLowerCase();
    var m = s.match(/^rgba?\(([^)]+)\)/i);
    if (m) {
      var parts = m[1].split(",").map(function (p) { return p.trim(); });
      if (parts.length >= 3) {
        var r = clampInt(parts[0]);
        var g = clampInt(parts[1]);
        var b = clampInt(parts[2]);
        return "#" + hex2(r) + hex2(g) + hex2(b);
      }
    }
    return "#000000";
  }

  function clampInt(s) {
    var n = parseFloat(s);
    if (isNaN(n)) return 0;
    n = Math.round(n);
    return n < 0 ? 0 : n > 255 ? 255 : n;
  }

  function hex2(n) {
    var h = n.toString(16);
    return h.length === 1 ? "0" + h : h;
  }

  // ---- UI ----
  function injectStyles() {
    var css = [
      "#cgTokenTweak{position:fixed;top:8px;right:8px;width:340px;max-height:calc(100vh - 16px);",
      "background:var(--color-panel-bg,rgba(255,255,255,0.97));color:var(--color-text,#111827);",
      "border:1px solid var(--color-panel-border,rgba(148,163,184,0.22));border-radius:8px;",
      "box-shadow:0 8px 24px var(--color-panel-shadow,rgba(15,23,42,0.12));",
      "font:12px/1.4 system-ui,sans-serif;z-index:99999;display:flex;flex-direction:column;}",
      "#cgTokenTweak .cg-tt-header{display:flex;align-items:center;gap:6px;padding:6px 8px;",
      "border-bottom:1px solid var(--color-panel-border,rgba(148,163,184,0.22));user-select:none;}",
      "#cgTokenTweak .cg-tt-title{flex:1;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:0.04em;opacity:0.7;}",
      "#cgTokenTweak .cg-tt-btn{appearance:none;background:transparent;border:1px solid var(--color-border,#e0ddd7);",
      "color:inherit;padding:2px 6px;font:inherit;font-size:11px;border-radius:4px;cursor:pointer;}",
      "#cgTokenTweak .cg-tt-btn:hover{background:var(--color-selected,rgba(29,79,114,0.08));}",
      "#cgTokenTweak.cg-tt-collapsed .cg-tt-body{display:none;}",
      "#cgTokenTweak .cg-tt-body{overflow-y:auto;padding:4px 0;flex:1;}",
      "#cgTokenTweak .cg-tt-group-header{display:flex;align-items:center;gap:6px;padding:4px 8px;font-weight:600;",
      "font-size:11px;text-transform:uppercase;letter-spacing:0.04em;opacity:0.6;cursor:pointer;}",
      "#cgTokenTweak .cg-tt-group-header:hover{opacity:1;}",
      "#cgTokenTweak .cg-tt-group-toggle{display:inline-block;width:8px;text-align:center;font-size:9px;}",
      "#cgTokenTweak .cg-tt-group.cg-tt-group-collapsed .cg-tt-row{display:none;}",
      "#cgTokenTweak .cg-tt-row{display:grid;grid-template-columns:1fr auto auto;align-items:center;gap:6px;",
      "padding:3px 8px;}",
      "#cgTokenTweak .cg-tt-row.cg-tt-row-overridden{background:var(--color-accent-hover,rgba(29,79,114,0.06));}",
      "#cgTokenTweak .cg-tt-row-unused{opacity:0.42;}",
      "#cgTokenTweak .cg-tt-row-unused .cg-tt-name::after{content:' (unused)';font-style:italic;opacity:0.7;font-size:9px;}",
      "#cgTokenTweak .cg-tt-name{font:11px/1.3 var(--font-mono,monospace);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;",
      "min-width:0;}",
      "#cgTokenTweak .cg-tt-control{display:flex;align-items:center;gap:4px;}",
      "#cgTokenTweak input[type=range]{width:100px;}",
      "#cgTokenTweak input[type=number],#cgTokenTweak input[type=text]{width:70px;font:11px/1 var(--font-mono,monospace);",
      "padding:2px 4px;border:1px solid var(--color-border,#e0ddd7);border-radius:3px;",
      "background:var(--color-surface,#fff);color:inherit;}",
      "#cgTokenTweak input[type=color]{width:28px;height:20px;padding:0;border:1px solid var(--color-border,#e0ddd7);",
      "border-radius:3px;background:transparent;cursor:pointer;}",
      "#cgTokenTweak .cg-tt-reset{appearance:none;background:transparent;border:none;color:inherit;cursor:pointer;",
      "width:18px;height:18px;line-height:1;font-size:14px;opacity:0;padding:0;}",
      "#cgTokenTweak .cg-tt-row-overridden .cg-tt-reset{opacity:0.6;}",
      "#cgTokenTweak .cg-tt-reset:hover{opacity:1;}",
      "#cgTokenTweak .cg-tt-empty{padding:12px;opacity:0.6;font-style:italic;text-align:center;}",
      "#cgTokenTweak .cg-tt-toast{position:absolute;bottom:8px;left:50%;transform:translateX(-50%);",
      "background:var(--color-text,#111);color:var(--color-bg,#fff);padding:4px 10px;border-radius:4px;",
      "font-size:11px;opacity:0;transition:opacity 200ms;pointer-events:none;}",
      "#cgTokenTweak .cg-tt-toast.cg-tt-toast-visible{opacity:1;}",
    ].join("");
    var style = document.createElement("style");
    style.id = "cgTokenTweakStyles";
    style.textContent = css;
    document.head.appendChild(style);
  }

  function buildLengthControl(name, current, defaultV) {
    var wrap = document.createElement("div");
    wrap.className = "cg-tt-control";
    var parsed = parseLength(current) || parseLength(defaultV);
    if (!parsed) return buildTextControl(name, current);
    var unit = parsed.unit;
    var step = unit === "px" ? 1 : 0.0625;
    var max = Math.max(parsed.num * 4, unit === "px" ? 100 : 6);
    var slider = document.createElement("input");
    slider.type = "range";
    slider.min = "0";
    slider.max = String(max);
    slider.step = String(step);
    slider.value = String(parsed.num);
    var num = document.createElement("input");
    num.type = "number";
    num.min = "0";
    num.step = String(step);
    num.value = String(parsed.num);
    var unitLabel = document.createElement("span");
    unitLabel.textContent = unit;
    unitLabel.style.opacity = "0.6";
    unitLabel.style.fontSize = "10px";
    function commit(value) {
      var v = parseFloat(value);
      if (isNaN(v)) return;
      var formatted = v + unit;
      slider.value = String(v);
      num.value = String(v);
      setOverride(name, formatted);
    }
    slider.addEventListener("input", function () { commit(slider.value); });
    num.addEventListener("input", function () { commit(num.value); });
    wrap.appendChild(slider);
    wrap.appendChild(num);
    wrap.appendChild(unitLabel);
    return wrap;
  }

  function buildColorControl(name, current) {
    var wrap = document.createElement("div");
    wrap.className = "cg-tt-control";
    var picker = document.createElement("input");
    picker.type = "color";
    picker.value = colorToHex(current);
    var text = document.createElement("input");
    text.type = "text";
    text.value = current;
    text.autocomplete = "off";
    text.spellcheck = false;
    picker.addEventListener("input", function () {
      text.value = picker.value;
      setOverride(name, picker.value);
    });
    text.addEventListener("change", function () {
      var v = text.value.trim();
      if (!v) return;
      setOverride(name, v);
      picker.value = colorToHex(v);
    });
    wrap.appendChild(picker);
    wrap.appendChild(text);
    return wrap;
  }

  function buildColorTextControl(name, current) {
    // For rgba()/hsla()/named colors that <input type="color"> can't represent.
    var wrap = document.createElement("div");
    wrap.className = "cg-tt-control";
    var text = document.createElement("input");
    text.type = "text";
    text.value = current;
    text.autocomplete = "off";
    text.spellcheck = false;
    text.style.width = "150px";
    text.addEventListener("change", function () {
      var v = text.value.trim();
      if (!v) return;
      setOverride(name, v);
    });
    wrap.appendChild(text);
    return wrap;
  }

  function buildDurationControl(name, current) {
    var wrap = document.createElement("div");
    wrap.className = "cg-tt-control";
    var num = parseFloat(current) || 0;
    var input = document.createElement("input");
    input.type = "number";
    input.min = "0";
    input.step = "10";
    input.value = String(num);
    var unitLabel = document.createElement("span");
    unitLabel.textContent = "ms";
    unitLabel.style.opacity = "0.6";
    unitLabel.style.fontSize = "10px";
    input.addEventListener("input", function () {
      var v = parseFloat(input.value);
      if (isNaN(v)) return;
      setOverride(name, v + "ms");
    });
    wrap.appendChild(input);
    wrap.appendChild(unitLabel);
    return wrap;
  }

  function buildNumberControl(name, current) {
    var wrap = document.createElement("div");
    wrap.className = "cg-tt-control";
    var input = document.createElement("input");
    input.type = "number";
    input.value = String(parseFloat(current) || 0);
    input.addEventListener("input", function () {
      if (input.value === "" || isNaN(parseFloat(input.value))) return;
      setOverride(name, input.value);
    });
    wrap.appendChild(input);
    return wrap;
  }

  function buildTextControl(name, current) {
    var wrap = document.createElement("div");
    wrap.className = "cg-tt-control";
    var input = document.createElement("input");
    input.type = "text";
    input.value = current;
    input.autocomplete = "off";
    input.spellcheck = false;
    input.style.width = "180px";
    input.addEventListener("change", function () {
      setOverride(name, input.value);
    });
    wrap.appendChild(input);
    return wrap;
  }

  function buildControl(name) {
    var current = state.overrides[name] || getDefaultValue(name);
    var defaultV = getDefaultValue(name);
    var type = inferControlType(current || defaultV);
    if (type === "color") return buildColorControl(name, current);
    if (type === "color-text") return buildColorTextControl(name, current);
    if (type === "length") return buildLengthControl(name, current, defaultV);
    if (type === "duration") return buildDurationControl(name, current);
    if (type === "number") return buildNumberControl(name, current);
    return buildTextControl(name, current);
  }

  function buildRow(name) {
    var row = document.createElement("div");
    row.className = "cg-tt-row";
    row.setAttribute("data-token", name);
    if (state.overrides[name] != null) row.classList.add("cg-tt-row-overridden");
    var unused = !state.referenced[name];
    if (unused) row.classList.add("cg-tt-row-unused");

    var label = document.createElement("div");
    label.className = "cg-tt-name";
    label.textContent = name;
    label.title = unused
      ? name + "\n\nNo consumers in current page CSS — tweaking this is a no-op until a Pass-N migration wires it up."
      : name;
    row.appendChild(label);

    var control = buildControl(name);
    row.appendChild(control);

    var resetBtn = document.createElement("button");
    resetBtn.type = "button";
    resetBtn.className = "cg-tt-reset";
    resetBtn.textContent = "✕";
    resetBtn.title = "Reset to default";
    resetBtn.addEventListener("click", function () { resetOne(name); });
    row.appendChild(resetBtn);

    return row;
  }

  function buildGroup(group, tokens) {
    if (tokens.length === 0) return null;
    var groupEl = document.createElement("div");
    groupEl.className = "cg-tt-group";
    groupEl.setAttribute("data-group", group.id);

    var header = document.createElement("div");
    header.className = "cg-tt-group-header";
    var toggle = document.createElement("span");
    toggle.className = "cg-tt-group-toggle";
    toggle.textContent = "▾";
    var title = document.createElement("span");
    title.textContent = group.label;
    var count = document.createElement("span");
    count.textContent = "(" + tokens.length + ")";
    count.style.opacity = "0.5";
    count.style.marginLeft = "auto";
    header.appendChild(toggle);
    header.appendChild(title);
    header.appendChild(count);
    header.addEventListener("click", function () {
      groupEl.classList.toggle("cg-tt-group-collapsed");
      toggle.textContent = groupEl.classList.contains("cg-tt-group-collapsed") ? "▸" : "▾";
    });
    groupEl.appendChild(header);

    tokens.forEach(function (name) { groupEl.appendChild(buildRow(name)); });
    return groupEl;
  }

  function buildPanel() {
    var panel = document.createElement("div");
    panel.id = "cgTokenTweak";
    if (state.collapsed) panel.classList.add("cg-tt-collapsed");

    var header = document.createElement("div");
    header.className = "cg-tt-header";
    var title = document.createElement("div");
    title.className = "cg-tt-title";
    title.textContent = "Tokens";
    var copyBtn = document.createElement("button");
    copyBtn.type = "button";
    copyBtn.className = "cg-tt-btn";
    copyBtn.textContent = "Copy snippet";
    copyBtn.title = "Copy overridden tokens as a :root { ... } CSS block";
    copyBtn.addEventListener("click", copySnippet);
    var resetAllBtn = document.createElement("button");
    resetAllBtn.type = "button";
    resetAllBtn.className = "cg-tt-btn";
    resetAllBtn.textContent = "Reset all";
    resetAllBtn.addEventListener("click", resetAll);
    var collapseBtn = document.createElement("button");
    collapseBtn.type = "button";
    collapseBtn.className = "cg-tt-btn";
    collapseBtn.textContent = state.collapsed ? "+" : "–";
    collapseBtn.title = "Collapse panel";
    collapseBtn.addEventListener("click", function () {
      var nowCollapsed = !panel.classList.contains("cg-tt-collapsed");
      panel.classList.toggle("cg-tt-collapsed", nowCollapsed);
      collapseBtn.textContent = nowCollapsed ? "+" : "–";
      state.collapsed = nowCollapsed;
      saveCollapsed(nowCollapsed);
    });
    header.appendChild(title);
    header.appendChild(copyBtn);
    header.appendChild(resetAllBtn);
    header.appendChild(collapseBtn);
    panel.appendChild(header);

    var body = document.createElement("div");
    body.className = "cg-tt-body";

    if (state.tokens.length === 0) {
      var empty = document.createElement("div");
      empty.className = "cg-tt-empty";
      empty.textContent = "No tunable tokens found.";
      body.appendChild(empty);
    } else {
      var grouped = {};
      state.tokens.forEach(function (name) {
        var gid = tokenGroupId(name);
        if (!grouped[gid]) grouped[gid] = [];
        grouped[gid].push(name);
      });
      GROUPS.forEach(function (group) {
        var tokens = grouped[group.id] || [];
        var groupEl = buildGroup(group, tokens);
        if (groupEl) body.appendChild(groupEl);
      });
    }

    panel.appendChild(body);

    var toast = document.createElement("div");
    toast.className = "cg-tt-toast";
    panel.appendChild(toast);
    state.toast = toast;

    return panel;
  }

  // ---- Mutations ----
  function setOverride(name, value) {
    state.overrides[name] = value;
    applyOverride(name, value);
    saveOverrides(state.overrides);
    var row = state.panel && state.panel.querySelector('.cg-tt-row[data-token="' + cssEscape(name) + '"]');
    if (row) row.classList.add("cg-tt-row-overridden");
    if (typeof window.refreshDetectorColors === "function") {
      try { window.refreshDetectorColors(); } catch (_) {}
    }
  }

  function resetOne(name) {
    delete state.overrides[name];
    clearOverride(name);
    saveOverrides(state.overrides);
    rerender();
    if (typeof window.refreshDetectorColors === "function") {
      try { window.refreshDetectorColors(); } catch (_) {}
    }
  }

  function resetAll() {
    Object.keys(state.overrides).forEach(clearOverride);
    state.overrides = {};
    saveOverrides(state.overrides);
    rerender();
    if (typeof window.refreshDetectorColors === "function") {
      try { window.refreshDetectorColors(); } catch (_) {}
    }
  }

  function copySnippet() {
    var keys = Object.keys(state.overrides).sort();
    if (keys.length === 0) {
      flashToast("No overrides to copy");
      return;
    }
    var lines = [":root {"];
    keys.forEach(function (name) {
      lines.push("  " + name + ": " + state.overrides[name] + ";");
    });
    lines.push("}");
    var text = lines.join("\n");
    var ok = false;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(
        function () { flashToast("Snippet copied"); },
        function () { fallbackCopy(text); }
      );
      return;
    }
    ok = fallbackCopy(text);
    flashToast(ok ? "Snippet copied" : "Copy failed");
  }

  function fallbackCopy(text) {
    var ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    var ok = false;
    try { ok = document.execCommand("copy"); } catch (_) {}
    document.body.removeChild(ta);
    return ok;
  }

  function flashToast(msg) {
    if (!state.toast) return;
    state.toast.textContent = msg;
    state.toast.classList.add("cg-tt-toast-visible");
    clearTimeout(state.toastTimer);
    state.toastTimer = setTimeout(function () {
      state.toast.classList.remove("cg-tt-toast-visible");
    }, 1400);
  }

  function rerender() {
    if (!state.panel) return;
    var parent = state.panel.parentNode;
    var nextSibling = state.panel.nextSibling;
    var newPanel = buildPanel();
    parent.removeChild(state.panel);
    if (nextSibling) parent.insertBefore(newPanel, nextSibling);
    else parent.appendChild(newPanel);
    state.panel = newPanel;
  }

  function cssEscape(s) {
    if (window.CSS && typeof window.CSS.escape === "function") return window.CSS.escape(s);
    return s.replace(/[^a-zA-Z0-9_-]/g, "\\$&");
  }

  // ---- Theme awareness ----
  function observeTheme() {
    if (!window.MutationObserver) return;
    var obs = new MutationObserver(function (records) {
      for (var i = 0; i < records.length; i++) {
        if (records[i].attributeName === "data-theme") {
          rerender();
          return;
        }
      }
    });
    obs.observe(document.documentElement, { attributes: true, attributeFilter: ["data-theme"] });
  }

  // ---- Init ----
  function mount() {
    if (!document.body) return;
    if (document.getElementById("cgTokenTweak")) return;
    state.tokens = discoverTokens();
    // Scan reference usage *before* injecting widget styles, so the widget's
    // own var() references don't count as page-CSS consumers.
    state.referenced = findReferencedTokens();
    injectStyles();
    state.panel = buildPanel();
    document.body.appendChild(state.panel);
    observeTheme();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", mount);
  } else {
    mount();
  }

  // Expose for manual debugging via the console.
  window.cgTokenTweak = {
    state: state,
    rerender: rerender,
    resetAll: resetAll,
    snippet: function () {
      var keys = Object.keys(state.overrides).sort();
      var lines = [":root {"];
      keys.forEach(function (name) { lines.push("  " + name + ": " + state.overrides[name] + ";"); });
      lines.push("}");
      return lines.join("\n");
    },
  };
})();
