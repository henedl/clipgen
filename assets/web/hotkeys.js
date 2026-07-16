/* clipgen shared hotkey registry — hotkeys.js
 *
 * One document-level keydown/keyup dispatcher shared by every web frontend
 * (and inlined into exported viewers by viewer.py). Pages register actions
 * against the catalog below; user overrides arrive via CLIPGEN_CONFIG
 * (config.HOTKEY_OVERRIDES -> utils.get_frontend_config -> clipgenApplyConfig)
 * and are edited in Settings -> Hotkeys (settings-modal.js).
 *
 * Contracts:
 *  - HOTKEY_SECTIONS / HOTKEY_CATALOG must stay JSON-parseable literals; the
 *    guard test (tests/test_hotkeys_frontend_source.py) extracts them.
 *  - Escape and Tab are reserved (modal / focus semantics) and are never
 *    dispatchable or recordable. Pages hand their Escape cascade to
 *    registerEscape() instead of adding their own document listeners.
 *  - A handler may return false to decline an event (no preventDefault), so
 *    e.g. Mod+V can fall through to native paste when nothing is selected.
 *  - Combo grammar: "[Mod+][Ctrl+][Alt+][Shift+]Key"; aliases separated by a
 *    space ("X Backspace"). Mod = Cmd on macOS, Ctrl elsewhere. Shift-produced
 *    punctuation keeps the produced character ("?", "{"), dropping Shift.
 */

(function () {
  "use strict";

  var IS_MAC = /Mac|iPhone|iPad/.test(navigator.platform || "");

  // ---- Catalog (the keymap as data) ----
  // Sections group actions for the cheatsheet and scope conflict detection:
  // two bindings only conflict when their sections' page sets intersect
  // (pages: null = every page).

  var HOTKEY_SECTIONS = [
    { id: "global",      label: "Everywhere",     pages: null },
    { id: "transport",   label: "Video playback", pages: ["composer", "screenspace", "transcripts"] },
    { id: "nav",         label: "Navigation",     pages: ["composer", "transcripts", "viewer", "studio"] },
    { id: "edit",        label: "Editing",        pages: ["composer", "workflows"] },
    { id: "studio",      label: "Studio",         pages: ["studio"] },
    { id: "composer",    label: "Composer",       pages: ["composer"] },
    { id: "screenspace", label: "Screenspace",    pages: ["screenspace"] },
    { id: "transcripts", label: "Transcripts",    pages: ["transcripts"] },
    { id: "workflows",   label: "Workflows",      pages: ["workflows"] },
    { id: "overview",    label: "Overview",       pages: ["overview"] },
    { id: "viewer",      label: "Viewer",         pages: ["viewer"] },
    { id: "gallery",     label: "Gallery",        pages: ["gallery"] }
  ];

  var HOTKEY_CATALOG = [
    { id: "global.cheatsheet", section: "global", group: "", label: "Keyboard shortcuts", combos: ["?"] },
    { id: "global.palette",   section: "global", group: "", label: "Command palette", combos: ["Mod+Shift+P", "Mod+K"] },
    { id: "global.primary",   section: "global", group: "", label: "Primary action (generate / run)", combos: ["G"] },
    { id: "global.refresh",   section: "global", group: "", label: "Refresh data", combos: ["R"] },
    { id: "global.search",    section: "global", group: "", label: "Focus search", combos: ["S"] },

    { id: "transport.playPause", section: "transport", group: "", label: "Play / pause", combos: ["Space"] },
    { id: "transport.seekBack",  section: "transport", group: "", label: "Seek back 5 s", combos: ["ArrowLeft"] },
    { id: "transport.seekFwd",   section: "transport", group: "", label: "Seek forward 5 s", combos: ["ArrowRight"] },
    { id: "transport.stepBack",  section: "transport", group: "", label: "Fine step back", combos: [","] },
    { id: "transport.stepFwd",   section: "transport", group: "", label: "Fine step forward", combos: ["."] },

    { id: "nav.next", section: "nav", group: "", label: "Next item", combos: ["J"] },
    { id: "nav.prev", section: "nav", group: "", label: "Previous item", combos: ["K"] },

    { id: "edit.undo", section: "edit", group: "", label: "Undo", combos: ["Mod+Z"] },
    { id: "edit.redo", section: "edit", group: "", label: "Redo", combos: ["Mod+Shift+Z", "Mod+Y"] },

    { id: "studio.buildReel",       section: "studio", group: "", label: "Build reel", combos: ["B"] },
    { id: "studio.buildHighlights", section: "studio", group: "", label: "Build highlights", combos: ["H"] },
    { id: "studio.moveLeft",        section: "studio", group: "Selection", label: "Move selection left", combos: ["ArrowLeft"] },
    { id: "studio.moveRight",       section: "studio", group: "Selection", label: "Move selection right", combos: ["ArrowRight"] },
    { id: "studio.moveUp",          section: "studio", group: "Selection", label: "Move selection up", combos: ["ArrowUp"] },
    { id: "studio.moveDown",        section: "studio", group: "Selection", label: "Move selection down", combos: ["ArrowDown"] },
    { id: "studio.sendArtifacts",   section: "studio", group: "Selection", label: "Add / remove selection in work area", combos: ["Enter"] },
    { id: "studio.sendReel",        section: "studio", group: "Selection", label: "Add / remove selection in reel", combos: ["Shift+Enter"] },

    { id: "composer.setIn",            section: "composer", group: "Cuts", label: "Set in point", combos: ["I"] },
    { id: "composer.setOut",           section: "composer", group: "Cuts", label: "Set out point", combos: ["O"] },
    { id: "composer.nudgeLeft",        section: "composer", group: "Cuts", label: "Nudge cut edge left (fine)", combos: ["["] },
    { id: "composer.nudgeRight",       section: "composer", group: "Cuts", label: "Nudge cut edge right (fine)", combos: ["]"] },
    { id: "composer.nudgeLeftBig",     section: "composer", group: "Cuts", label: "Nudge cut edge left (1 s)", combos: ["{"] },
    { id: "composer.nudgeRightBig",    section: "composer", group: "Cuts", label: "Nudge cut edge right (1 s)", combos: ["}"] },
    { id: "composer.deleteSelection",  section: "composer", group: "Cuts", label: "Delete selected annotation / cut", combos: ["X", "Backspace"] },
    { id: "composer.toolSelect",       section: "composer", group: "Annotate", label: "Select tool", combos: ["V"] },
    { id: "composer.toolText",         section: "composer", group: "Annotate", label: "Text tool", combos: ["T"] },
    { id: "composer.toolDraw",         section: "composer", group: "Annotate", label: "Draw tool", combos: ["D"] },
    { id: "composer.toolErase",        section: "composer", group: "Annotate", label: "Erase tool", combos: ["E"] },
    { id: "composer.toggleSource",     section: "composer", group: "Marker lanes", label: "Toggle marker lane 1–3", combos: ["1", "2", "3"], rebindable: false, displayKeys: "1–3" },
    { id: "composer.toggleAllSources", section: "composer", group: "Marker lanes", label: "Toggle all marker lanes", combos: ["0"] },
    { id: "composer.note.zoomTimeline", section: "composer", group: "Timeline", label: "Zoom / pan timeline", note: "scroll · drag to pan" },

    { id: "screenspace.blink", section: "screenspace", group: "", label: "Blink region overlay (hold)", combos: ["B"] },

    { id: "transcripts.mark",         section: "transcripts", group: "Marks", label: "Mark active segment", combos: ["M"] },
    { id: "transcripts.nextMarked",   section: "transcripts", group: "Marks", label: "Next marked segment", combos: ["N"] },
    { id: "transcripts.prevMarked",   section: "transcripts", group: "Marks", label: "Previous marked segment", combos: ["P"] },
    { id: "transcripts.markCategory", section: "transcripts", group: "Marks", label: "Mark with category 1–9", combos: ["1", "2", "3", "4", "5", "6", "7", "8", "9"], rebindable: false, displayKeys: "1–9" },

    { id: "workflows.copy",            section: "workflows", group: "Clipboard", label: "Copy selected nodes", combos: ["Mod+C"] },
    { id: "workflows.paste",           section: "workflows", group: "Clipboard", label: "Paste nodes", combos: ["Mod+V"] },
    { id: "workflows.duplicate",       section: "workflows", group: "Clipboard", label: "Duplicate selection", combos: ["Mod+D"] },
    { id: "workflows.fitView",         section: "workflows", group: "Canvas", label: "Fit graph to view", combos: ["F"] },
    { id: "workflows.deleteSelection", section: "workflows", group: "Canvas", label: "Delete selected wire / nodes", combos: ["Delete", "Backspace"] },
    { id: "workflows.note.pan",        section: "workflows", group: "Canvas", label: "Pan canvas", note: "middle-drag" },
    { id: "workflows.note.zoom",       section: "workflows", group: "Canvas", label: "Zoom canvas", note: "scroll wheel" },
    { id: "workflows.note.select",     section: "workflows", group: "Canvas", label: "Select nodes", note: "drag / shift-click" },

    { id: "overview.tabMap",         section: "overview", group: "Tabs", label: "Show Map tab", combos: ["1"] },
    { id: "overview.tabConvergence", section: "overview", group: "Tabs", label: "Show Convergence tab", combos: ["2"] },
    { id: "overview.tabMetadata",    section: "overview", group: "Tabs", label: "Show Metadata tab", combos: ["3"] },
    { id: "overview.zonePrev",       section: "overview", group: "Convergence", label: "Previous convergence zone", combos: ["ArrowLeft"] },
    { id: "overview.zoneNext",       section: "overview", group: "Convergence", label: "Next convergence zone", combos: ["ArrowRight"] },

    { id: "gallery.prev", section: "gallery", group: "", label: "Previous image", combos: ["ArrowLeft"] },
    { id: "gallery.next", section: "gallery", group: "", label: "Next image", combos: ["ArrowRight"] }
  ];

  // ---- Internal state ----

  var ACTIONS_BY_ID = {};
  var SECTIONS_BY_ID = {};
  var i;
  for (i = 0; i < HOTKEY_CATALOG.length; i++) ACTIONS_BY_ID[HOTKEY_CATALOG[i].id] = HOTKEY_CATALOG[i];
  for (i = 0; i < HOTKEY_SECTIONS.length; i++) SECTIONS_BY_ID[HOTKEY_SECTIONS[i].id] = HOTKEY_SECTIONS[i];

  var _overrides = {};        // action id -> combo string ("" = disabled)
  var _attachments = {};      // action id -> [{handler, when, onRelease, repeat, allowInInput}]
  var _attachOrder = [];      // action ids in first-registration order
  var _comboIndex = {};       // combo string -> [action id] (attach order)
  var _escapeHandlers = [];   // [fn(e) -> true if consumed]
  var _held = [];             // [{attachment, baseKey}] pending onRelease

  // ---- Combo helpers ----

  function normalizeEvent(e) {
    var key = e.key;
    if (key === undefined || key === "Escape" || key === "Tab") return null;
    if (key === "Meta" || key === "Control" || key === "Alt" || key === "Shift") return null;
    if (key === "Dead") return null; // dead key (e.g. ` or ´ on ISO layouts)
    var mod = IS_MAC ? e.metaKey : e.ctrlKey;
    var ctrl = IS_MAC && e.ctrlKey;
    var alt = e.altKey;
    var name;
    var shift = false;
    if (key === " " || key === "Spacebar") {
      name = "Space";
      shift = e.shiftKey;
    } else if (key.length === 1) {
      if (key.toUpperCase() !== key.toLowerCase()) {
        // A letter: uppercase base + explicit Shift.
        name = key.toUpperCase();
        shift = e.shiftKey;
      } else {
        // Digit or punctuation: keep the produced character, drop Shift
        // ("?" and "{" are combos of their own, not Shift+/).
        name = key;
        if (!/[0-9]/.test(key)) {
          // Punctuation often *requires* Option/AltGr on ISO layouts
          // (Option+8 = "[" on a German/Swedish Mac; AltGr arrives as
          // Ctrl+Alt on Windows), so the producing modifiers are not part
          // of the combo — the produced character is. Digits keep their
          // modifiers; they never need AltGr.
          alt = false;
          if (!IS_MAC && e.altKey) mod = false;
        }
      }
    } else {
      name = key; // ArrowLeft, Backspace, Enter, Delete, F1, ...
      shift = e.shiftKey;
    }
    var parts = [];
    if (mod) parts.push("Mod");
    if (ctrl) parts.push("Ctrl");
    if (alt) parts.push("Alt");
    if (shift) parts.push("Shift");
    parts.push(name);
    return parts.join("+");
  }

  var _MAC_MODS = { Mod: "⌘", Ctrl: "⌃", Alt: "⌥", Shift: "⇧" };
  var _PC_MODS = { Mod: "Ctrl", Ctrl: "Ctrl", Alt: "Alt", Shift: "Shift" };
  var _KEY_GLYPHS = {
    ArrowLeft: "←", ArrowRight: "→", ArrowUp: "↑", ArrowDown: "↓",
    Backspace: "⌫", Delete: "Del", Enter: "↩", Space: "Space"
  };

  function formatCombo(combo) {
    var tokens = combo.split("+");
    // "Shift+Z" splits fine, but a literal "+" key would arrive as trailing
    // empty tokens; re-join those into "+".
    var cleaned = [];
    for (var t = 0; t < tokens.length; t++) {
      if (tokens[t] === "" ) {
        if (cleaned.length && cleaned[cleaned.length - 1] === "") continue;
        cleaned.push("+");
      } else cleaned.push(tokens[t]);
    }
    var out = [];
    for (var k = 0; k < cleaned.length; k++) {
      var tok = cleaned[k];
      var isMod = tok === "Mod" || tok === "Ctrl" || tok === "Alt" || tok === "Shift";
      if (isMod) out.push(IS_MAC ? _MAC_MODS[tok] : _PC_MODS[tok]);
      else out.push(_KEY_GLYPHS[tok] || tok);
    }
    return out.join(IS_MAC ? "" : "+");
  }

  function resolvedCombos(id) {
    var entry = ACTIONS_BY_ID[id];
    if (!entry || entry.note) return [];
    if (Object.prototype.hasOwnProperty.call(_overrides, id)) {
      var ov = _overrides[id];
      if (typeof ov !== "string" || ov.replace(/\s+/g, "") === "") return [];
      return ov.replace(/^\s+|\s+$/g, "").split(/\s+/);
    }
    return entry.combos ? entry.combos.slice() : [];
  }

  function rebuildIndex() {
    _comboIndex = {};
    for (var a = 0; a < _attachOrder.length; a++) {
      var id = _attachOrder[a];
      var entry = ACTIONS_BY_ID[id];
      if (!entry || entry.note) continue;
      var combos = resolvedCombos(id);
      for (var c = 0; c < combos.length; c++) {
        var combo = combos[c];
        if (!_comboIndex[combo]) _comboIndex[combo] = [];
        if (_comboIndex[combo].indexOf(id) === -1) _comboIndex[combo].push(id);
      }
    }
  }

  function pagesIntersect(sectionA, sectionB) {
    var a = SECTIONS_BY_ID[sectionA];
    var b = SECTIONS_BY_ID[sectionB];
    if (!a || !b) return true; // unknown -> be conservative
    if (a.pages === null || b.pages === null) return true;
    for (var p = 0; p < a.pages.length; p++) {
      if (b.pages.indexOf(a.pages[p]) !== -1) return true;
    }
    return false;
  }

  // All catalog actions (any page) whose resolved combos include `combo` and
  // whose section page-scope intersects `id`'s section. Used by the settings
  // recorder for conflict warnings.
  function comboConflicts(combo, id) {
    var self = ACTIONS_BY_ID[id];
    var out = [];
    for (var c = 0; c < HOTKEY_CATALOG.length; c++) {
      var entry = HOTKEY_CATALOG[c];
      if (entry.id === id || entry.note) continue;
      if (resolvedCombos(entry.id).indexOf(combo) === -1) continue;
      if (self && !pagesIntersect(entry.section, self.section)) continue;
      out.push({ id: entry.id, label: entry.label, section: entry.section });
    }
    return out;
  }

  // ---- Registration ----

  function register(entries) {
    for (var n = 0; n < entries.length; n++) {
      var spec = entries[n];
      var entry = ACTIONS_BY_ID[spec.id];
      if (!entry) {
        if (window.console && console.error) console.error("hotkeys: unknown action id " + spec.id);
        continue;
      }
      if (!_attachments[spec.id]) {
        _attachments[spec.id] = [];
        _attachOrder.push(spec.id);
      }
      if (!entry.note) {
        _attachments[spec.id].push({
          handler: spec.handler,
          when: spec.when || null,
          onRelease: spec.onRelease || null,
          repeat: spec.repeat !== false,
          allowInInput: !!spec.allowInInput
        });
      }
    }
    rebuildIndex();
  }

  function registerEscape(fn) {
    _escapeHandlers.push(fn);
  }

  function applyOverrides(map) {
    _overrides = {};
    if (map && typeof map === "object") {
      for (var key in map) {
        if (Object.prototype.hasOwnProperty.call(map, key) && typeof map[key] === "string") {
          _overrides[key] = map[key];
        }
      }
    }
    rebuildIndex();
  }

  // ---- Dispatch ----

  function isTypingTarget(t) {
    if (!t || !t.matches) return false;
    return t.matches("input, textarea, select") || t.isContentEditable === true;
  }

  function blockingModalOpen() {
    if (typeof isBlockingModalOpen === "function" && isBlockingModalOpen()) return true;
    return document.body && document.body.classList.contains("modal-open");
  }

  function onDocKeydown(e) {
    if (e.defaultPrevented) return;
    if (e.key === "Escape") {
      // A blocking modal's own capture-phase trap owns Escape.
      if (blockingModalOpen()) return;
      for (var h = 0; h < _escapeHandlers.length; h++) {
        if (_escapeHandlers[h](e) === true) {
          e.preventDefault();
          return;
        }
      }
      return;
    }
    if (blockingModalOpen()) {
      // The cheatsheet is our own blocking modal: its toggle combo passes
      // back through so a second "?" press closes it (like the old per-page
      // popovers). Everything else stays suppressed.
      if (_sheetEl && !_sheetEl.classList.contains("hidden") && !isTypingTarget(e.target)) {
        var sheetCombo = normalizeEvent(e);
        if (sheetCombo && resolvedCombos("global.cheatsheet").indexOf(sheetCombo) !== -1) {
          e.preventDefault();
          closeCheatsheet();
        }
      }
      return;
    }
    var combo = normalizeEvent(e);
    if (!combo) return;
    var ids = _comboIndex[combo];
    if (!ids) return;
    var typing = isTypingTarget(e.target);
    for (var n = 0; n < ids.length; n++) {
      var atts = _attachments[ids[n]] || [];
      for (var a = 0; a < atts.length; a++) {
        var att = atts[a];
        if (typing && !att.allowInInput) continue;
        if (e.repeat && !att.repeat) continue;
        if (att.when && !att.when()) continue;
        if (att.handler && att.handler(e, combo) === false) continue;
        e.preventDefault();
        if (att.onRelease) {
          var toks = combo.split("+");
          _held.push({ attachment: att, baseKey: toks[toks.length - 1].toUpperCase() });
        }
        return;
      }
    }
  }

  function onDocKeyup(e) {
    if (!_held.length) return;
    var key = (e.key || "").length === 1 ? e.key.toUpperCase() : e.key;
    var remaining = [];
    for (var n = 0; n < _held.length; n++) {
      if (_held[n].baseKey === key) _held[n].attachment.onRelease(e);
      else remaining.push(_held[n]);
    }
    _held = remaining;
  }

  function onWindowBlur() {
    // Alt-tab etc. can swallow keyup; treat blur as release-all so hold-to-act
    // actions (Screenspace blink) never get stuck on.
    for (var n = 0; n < _held.length; n++) _held[n].attachment.onRelease(null);
    _held = [];
  }

  document.addEventListener("keydown", onDocKeydown);
  document.addEventListener("keyup", onDocKeyup);
  window.addEventListener("blur", onWindowBlur);

  // ---- Cheatsheet overlay ----

  var _sheetEl = null;

  function buildSheet() {
    var overlay = el("div", "hk-overlay hidden");
    var panel = el("div", "hk-panel");
    var header = el("div", "hk-header");
    header.appendChild(el("h2", "", "Keyboard shortcuts"));
    var close = el("button", "hk-close", "×");
    close.type = "button";
    close.setAttribute("aria-label", "Close");
    close.addEventListener("click", closeCheatsheet);
    header.appendChild(close);
    panel.appendChild(header);
    panel.appendChild(el("div", "hk-body"));
    var footer = el("div", "hk-footer");
    var esc = el("span", "hk-footer-esc");
    var escKbd = el("kbd", "", "Esc");
    esc.appendChild(escKbd);
    esc.appendChild(document.createTextNode(" closes dialogs and backs out of tools."));
    footer.appendChild(esc);
    if (window.openSettingsModal) {
      footer.appendChild(el("span", "hk-footer-hint", "Customize in Settings → Hotkeys."));
    }
    panel.appendChild(footer);
    overlay.appendChild(panel);
    document.body.appendChild(overlay);
    return overlay;
  }

  function renderSheet() {
    var body = _sheetEl.querySelector(".hk-body");
    body.innerHTML = "";
    var frag = document.createDocumentFragment();
    for (var s = 0; s < HOTKEY_SECTIONS.length; s++) {
      var section = HOTKEY_SECTIONS[s];
      var rows = [];
      for (var c = 0; c < HOTKEY_CATALOG.length; c++) {
        var entry = HOTKEY_CATALOG[c];
        if (entry.section !== section.id) continue;
        if (!_attachments[entry.id]) continue;
        if (!entry.note && resolvedCombos(entry.id).length === 0) continue; // disabled
        rows.push(entry);
      }
      if (!rows.length) continue;
      var secEl = el("div", "hk-section");
      secEl.appendChild(el("h3", "", section.label));
      for (var r = 0; r < rows.length; r++) {
        var row = el("div", "hk-row");
        var keys = el("span", "hk-keys");
        if (rows[r].note) {
          keys.appendChild(el("span", "hk-note", rows[r].note));
        } else if (rows[r].displayKeys) {
          keys.appendChild(el("kbd", "", rows[r].displayKeys));
        } else {
          var combos = resolvedCombos(rows[r].id);
          for (var k = 0; k < combos.length; k++) {
            if (k > 0) keys.appendChild(el("span", "hk-or", "or"));
            keys.appendChild(el("kbd", "", formatCombo(combos[k])));
          }
        }
        row.appendChild(keys);
        row.appendChild(el("span", "hk-label", rows[r].label));
        secEl.appendChild(row);
      }
      frag.appendChild(secEl);
    }
    body.appendChild(frag);
  }

  function openCheatsheet() {
    if (!_sheetEl) _sheetEl = buildSheet();
    renderSheet();
    _sheetEl.classList.remove("hidden");
    if (typeof openBlockingModal === "function") {
      openBlockingModal(_sheetEl, {
        onEscape: closeCheatsheet,
        onBackdropClick: closeCheatsheet,
        restoreFocus: true
      });
    }
  }

  function closeCheatsheet() {
    if (!_sheetEl || _sheetEl.classList.contains("hidden")) return;
    _sheetEl.classList.add("hidden");
    if (typeof closeBlockingModal === "function") closeBlockingModal(_sheetEl);
  }

  function toggleCheatsheet() {
    if (_sheetEl && !_sheetEl.classList.contains("hidden")) closeCheatsheet();
    else openCheatsheet();
  }

  function catalog() {
    return { sections: HOTKEY_SECTIONS.slice(), actions: HOTKEY_CATALOG.slice() };
  }

  register([{ id: "global.cheatsheet", handler: toggleCheatsheet }]);

  window.ClipgenHotkeys = {
    register: register,
    registerEscape: registerEscape,
    catalog: catalog,
    resolvedCombos: resolvedCombos,
    applyOverrides: applyOverrides,
    normalizeEvent: normalizeEvent,
    formatCombo: formatCombo,
    comboConflicts: comboConflicts,
    toggleCheatsheet: toggleCheatsheet,
    closeCheatsheet: closeCheatsheet
  };
})();
