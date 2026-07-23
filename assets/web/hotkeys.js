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
    { id: "gallery",     label: "Gallery",        pages: ["gallery"] },
    { id: "start",       label: "Start launcher", pages: [] },
    { id: "settings",    label: "Settings",       pages: [] }
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
    { id: "studio.togglePanel",     section: "studio", group: "Panels", label: "Collapse / expand artifact & reel panel", combos: ["V"] },
    { id: "studio.toggleSidebar",   section: "studio", group: "Panels", label: "Collapse / expand filter sidebar", combos: ["F"] },
    { id: "studio.stashArtifacts",  section: "studio", group: "Queue", label: "Stash artifacts", combos: ["A"] },
    { id: "studio.stashReel",       section: "studio", group: "Queue", label: "Stash reel", combos: ["Shift+A"] },
    { id: "studio.clearArtifacts",  section: "studio", group: "Queue", label: "Clear artifacts", combos: ["C"] },
    { id: "studio.clearReel",       section: "studio", group: "Queue", label: "Clear reel", combos: ["Shift+C"] },
    { id: "studio.focusFilter",        section: "studio", group: "Selection", label: "Select filter list", combos: ["1"] },
    { id: "studio.focusArtifacts",     section: "studio", group: "Selection", label: "Select artifact queue", combos: ["2"] },
    { id: "studio.focusReel",          section: "studio", group: "Selection", label: "Select reel queue", combos: ["3"] },
    { id: "studio.focusArtifactStash", section: "studio", group: "Selection", label: "Select stashed artifacts", combos: ["4"] },
    { id: "studio.focusReelStash",     section: "studio", group: "Selection", label: "Select stashed reels", combos: ["5"] },
    { id: "studio.selectTab",          section: "studio", group: "Tabs", label: "Switch preview tab by number", combos: ["Shift+1", "Shift+2", "Shift+3", "Shift+4"], rebindable: false, displayKeys: "⇧1–4" },

    { id: "composer.seekBackMid",      section: "composer", group: "Playhead", label: "Seek back 2.5 s", combos: ["Shift+ArrowLeft"] },
    { id: "composer.seekFwdMid",       section: "composer", group: "Playhead", label: "Seek forward 2.5 s", combos: ["Shift+ArrowRight"] },
    { id: "composer.stepBackFrame",    section: "composer", group: "Playhead", label: "Frame step back", combos: ["Shift+,"] },
    { id: "composer.stepFwdFrame",     section: "composer", group: "Playhead", label: "Frame step forward", combos: ["Shift+."] },
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
    { id: "composer.toolRect",         section: "composer", group: "Annotate", label: "Rectangle tool", combos: ["U"] },
    { id: "composer.toolEllipse",      section: "composer", group: "Annotate", label: "Circle tool", combos: ["C"] },
    { id: "composer.holdHideAnnotations", section: "composer", group: "Annotate", label: "Hide annotations (hold or tap)", combos: ["B"] },
    { id: "composer.toggleSource",     section: "composer", group: "Marker lanes", label: "Toggle marker lane 1–3", combos: ["1", "2", "3"], rebindable: false, displayKeys: "1–3" },
    { id: "composer.toggleAllSources", section: "composer", group: "Marker lanes", label: "Toggle all marker lanes", combos: ["0"] },
    { id: "composer.toggleThumbs",     section: "composer", group: "Marker lanes", label: "Toggle thumbnail strips", combos: ["S"] },
    { id: "composer.toggleScrubAudio", section: "composer", group: "Marker lanes", label: "Toggle hover audio scrub", combos: ["W"] },
    { id: "composer.toggleSidebar",    section: "composer", group: "Panels", label: "Collapse / expand timelines sidebar", combos: ["F"] },
    { id: "composer.note.zoomTimeline", section: "composer", group: "Timeline", label: "Zoom / pan timeline", note: "scroll · drag to pan" },

    { id: "screenspace.blink", section: "screenspace", group: "", label: "Blink region overlay (hold)", combos: ["B"] },
    { id: "screenspace.stepBackFine", section: "screenspace", group: "Transport", label: "Fine step back 1 s", combos: ["Shift+ArrowLeft"] },
    { id: "screenspace.stepFwdFine",  section: "screenspace", group: "Transport", label: "Fine step forward 1 s", combos: ["Shift+ArrowRight"] },
    { id: "screenspace.setIn",  section: "screenspace", group: "Marks", label: "Set in marker", combos: ["I"] },
    { id: "screenspace.setOut", section: "screenspace", group: "Marks", label: "Set out marker", combos: ["O"] },
    { id: "screenspace.togglePanel",   section: "screenspace", group: "Panels", label: "Collapse / expand bottom panel", combos: ["V"] },
    { id: "screenspace.toggleInfoPanel", section: "screenspace", group: "Panels", label: "Collapse / expand participant details", combos: ["F"] },
    { id: "screenspace.cycleToolPrev", section: "screenspace", group: "Tools", label: "Previous tool tab", combos: ["Z"] },
    { id: "screenspace.cycleToolNext", section: "screenspace", group: "Tools", label: "Next tool tab", combos: ["X"] },
    { id: "screenspace.selectTool",    section: "screenspace", group: "Tools", label: "Select tool / category by number", combos: ["1", "2", "3", "4", "5", "6", "7", "8", "9"], rebindable: false, displayKeys: "1–9" },

    { id: "transcripts.mark",         section: "transcripts", group: "Marks", label: "Mark active segment", combos: ["M"] },
    { id: "transcripts.nextMarked",   section: "transcripts", group: "Marks", label: "Next marked segment", combos: ["N"] },
    { id: "transcripts.prevMarked",   section: "transcripts", group: "Marks", label: "Previous marked segment", combos: ["P"] },
    { id: "transcripts.markCategory", section: "transcripts", group: "Marks", label: "Mark with category 1–9", combos: ["1", "2", "3", "4", "5", "6", "7", "8", "9"], rebindable: false, displayKeys: "1–9" },
    { id: "transcripts.cyclePartPrev",  section: "transcripts", group: "Participants", label: "Previous participant", combos: ["Z"] },
    { id: "transcripts.cyclePartNext",  section: "transcripts", group: "Participants", label: "Next participant", combos: ["X"] },
    { id: "transcripts.pillMenu",       section: "transcripts", group: "Participants", label: "Open participant options (then 1–4)", combos: ["O"] },
    { id: "transcripts.toggleCaptions", section: "transcripts", group: "Playback", label: "Toggle captions", combos: ["C"] },
    { id: "transcripts.speedDown",      section: "transcripts", group: "Playback", label: "Slower playback", combos: ["Shift+,"] },
    { id: "transcripts.speedUp",        section: "transcripts", group: "Playback", label: "Faster playback", combos: ["Shift+."] },
    { id: "transcripts.fullscreen",     section: "transcripts", group: "Playback", label: "Toggle fullscreen video", combos: ["F"] },

    { id: "workflows.copy",            section: "workflows", group: "Clipboard", label: "Copy selected nodes", combos: ["Mod+C"] },
    { id: "workflows.paste",           section: "workflows", group: "Clipboard", label: "Paste nodes", combos: ["Mod+V"] },
    { id: "workflows.duplicate",       section: "workflows", group: "Clipboard", label: "Duplicate selection", combos: ["Mod+D"] },
    { id: "workflows.fitView",         section: "workflows", group: "Canvas", label: "Fit graph to view", combos: ["F"] },
    { id: "workflows.deleteSelection", section: "workflows", group: "Canvas", label: "Delete selected wire / nodes", combos: ["Delete", "Backspace"] },
    { id: "workflows.selectAll",       section: "workflows", group: "Canvas", label: "Select all nodes", combos: ["Mod+A"] },
    { id: "workflows.panMode",         section: "workflows", group: "Canvas", label: "Pan canvas (hold + drag)", combos: ["Space"] },
    { id: "workflows.nudge",           section: "workflows", group: "Canvas", label: "Nudge selection (Shift = grid step)", combos: ["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown", "Shift+ArrowLeft", "Shift+ArrowRight", "Shift+ArrowUp", "Shift+ArrowDown"], rebindable: false, displayKeys: "←↑↓→" },
    { id: "workflows.zoomReset",       section: "workflows", group: "Canvas", label: "Reset zoom to 100%", combos: ["0"] },
    { id: "workflows.toggleSnap",      section: "workflows", group: "Canvas", label: "Toggle snap to grid", combos: ["Shift+G"] },
    { id: "workflows.addNote",         section: "workflows", group: "Canvas", label: "Add sticky note", combos: ["N"] },
    { id: "workflows.note.pan",        section: "workflows", group: "Canvas", label: "Pan canvas", note: "space-drag · middle-drag · two-finger scroll" },
    { id: "workflows.note.zoom",       section: "workflows", group: "Canvas", label: "Zoom canvas", note: "pinch · Ctrl/⌘ + scroll" },
    { id: "workflows.note.select",     section: "workflows", group: "Canvas", label: "Select nodes", note: "drag / shift-click" },
    { id: "workflows.cleanUp",         section: "workflows", group: "Blueprint", label: "Clean up layout", combos: ["L"] },
    { id: "workflows.stash",           section: "workflows", group: "Blueprint", label: "Stash selection", combos: ["Shift+S"] },
    { id: "workflows.newBlueprint",    section: "workflows", group: "Blueprint", label: "New blueprint", combos: ["Shift+N"] },
    { id: "workflows.renameBlueprint", section: "workflows", group: "Blueprint", label: "Rename blueprint", combos: ["Shift+E"] },
    { id: "workflows.focusSelector",   section: "workflows", group: "Blueprint", label: "Focus blueprint selector", combos: ["Shift+B"] },
    { id: "workflows.deleteBlueprint", section: "workflows", group: "Blueprint", label: "Delete blueprint", combos: ["Mod+Shift+Backspace"] },

    { id: "overview.tabMap",         section: "overview", group: "Tabs", label: "Show Map tab", combos: ["1"] },
    { id: "overview.tabConvergence", section: "overview", group: "Tabs", label: "Show Convergence tab", combos: ["2"] },
    { id: "overview.tabMetadata",    section: "overview", group: "Tabs", label: "Show Metadata tab", combos: ["3"] },
    { id: "overview.zonePrev",       section: "overview", group: "Convergence", label: "Previous convergence zone", combos: ["ArrowLeft"] },
    { id: "overview.zoneNext",       section: "overview", group: "Convergence", label: "Next convergence zone", combos: ["ArrowRight"] },

    { id: "gallery.prev", section: "gallery", group: "", label: "Previous image", combos: ["ArrowLeft"] },
    { id: "gallery.next", section: "gallery", group: "", label: "Next image", combos: ["ArrowRight"] },

    { id: "start.tabGoogle",    section: "start", group: "Spreadsheet", label: "Use Google Sheets",     combos: ["G"] },
    { id: "start.tabExcel",     section: "start", group: "Spreadsheet", label: "Use Excel file",        combos: ["E"] },
    { id: "start.tabNone",      section: "start", group: "Spreadsheet", label: "No spreadsheet",         combos: ["N"] },
    { id: "start.browseInput",  section: "start", group: "Folders",     label: "Browse input folder",    combos: ["I"] },
    { id: "start.browseOutput", section: "start", group: "Folders",     label: "Browse output folder",   combos: ["O"] },
    { id: "start.confirm",      section: "start", group: "",            label: "Open workspace",         combos: ["Mod+Enter"] },

    { id: "settings.tab",       section: "settings", group: "", label: "Switch settings tab by number", combos: ["1", "2", "3", "4", "5", "6", "7", "8", "9"], rebindable: false, displayKeys: "1–9" },
    { id: "settings.cyclePrev", section: "settings", group: "", label: "Previous settings tab", combos: ["Z"] },
    { id: "settings.cycleNext", section: "settings", group: "", label: "Next settings tab", combos: ["X"] },
    { id: "settings.resetTab",  section: "settings", group: "", label: "Reset this tab to defaults", combos: ["R"] },
    { id: "settings.resetAll",  section: "settings", group: "", label: "Reset all settings to defaults", combos: ["Shift+R"] }
  ];

  // ---- Internal state ----

  var ACTIONS_BY_ID = {};
  var SECTIONS_BY_ID = {};
  var i;
  for (i = 0; i < HOTKEY_CATALOG.length; i++) ACTIONS_BY_ID[HOTKEY_CATALOG[i].id] = HOTKEY_CATALOG[i];
  for (i = 0; i < HOTKEY_SECTIONS.length; i++) SECTIONS_BY_ID[HOTKEY_SECTIONS[i].id] = HOTKEY_SECTIONS[i];

  var _overrides = {};        // action id -> combo string ("" = disabled)
  var _attachments = {};      // action id -> [{handler, when, onRelease, repeat, allowInInput, inModal}]
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
      if (e.shiftKey && (e.code === "Comma" || e.code === "Period")) {
        // Shifted , / . resolve to layout-dependent characters ("<" on US,
        // ";" on Swedish), so "Shift+," would never match below. The physical
        // Comma/Period keys are layout-stable — map these two by e.code so
        // frame-step combos work everywhere.
        name = e.code === "Comma" ? "," : ".";
        shift = true;
      } else if (e.shiftKey && /^Digit[0-9]$/.test(e.code)) {
        // Shifted digits resolve to layout-dependent symbols ("!" on US,
        // "&" on French), so "Shift+1" would never match below. The physical
        // DigitN keys are layout-stable — map by e.code so numbered combos
        // (e.g. Studio's Shift+1…4 tab switch) work everywhere.
        name = e.code.charAt(5);
        shift = true;
      } else if (key.toUpperCase() !== key.toLowerCase()) {
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

  // The symbol glyphs (mac modifier keys + arrows / enter / backspace) that
  // render visually smaller than Latin letters/digits at the same point size,
  // so they get bumped up via .hk-glyph in fillKeycap.
  var _GLYPH_CHARS = { "⌘": 1, "⌃": 1, "⌥": 1, "⇧": 1, "←": 1, "→": 1, "↑": 1, "↓": 1, "⌫": 1, "↩": 1 };

  // Combo string → array of display tokens, e.g. "Shift+A" → ["⇧", "A"] on
  // macOS, ["Shift", "A"] on PC. "+" is the token separator; a literal "+" key
  // arrives as empty split tokens and is re-joined into a "+" token.
  function comboTokens(combo) {
    var tokens = combo.split("+");
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
    return out;
  }

  function formatCombo(combo) {
    return comboTokens(combo).join(IS_MAC ? "" : "+");
  }

  // Render a combo into a key-cap element as DOM nodes, wrapping the small
  // symbol glyphs in .hk-glyph so they read at a legible size beside letters.
  // Used by the cheatsheet key-caps and the Alt-hold hint chips (the plain
  // string formatCombo stays for the settings rebinder and elsewhere).
  function fillKeycap(node, combo) {
    var out = comboTokens(combo);
    var sep = IS_MAC ? "" : "+";
    for (var i = 0; i < out.length; i++) {
      if (i > 0 && sep) node.appendChild(document.createTextNode(sep));
      if (_GLYPH_CHARS[out[i]]) node.appendChild(el("span", "hk-glyph", out[i]));
      else node.appendChild(document.createTextNode(out[i]));
    }
    return node;
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
          allowInInput: !!spec.allowInInput,
          inModal: !!spec.inModal
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
    if (e.key === "Alt") {
      // Bare Alt (no other modifiers) arms the discoverability hints. The
      // !ctrlKey gate keeps Windows AltGr (which arrives as Ctrl+Alt) from
      // arming them.
      if (!e.repeat && !e.ctrlKey && !e.metaKey && !e.shiftKey) armHints();
    } else if (_hintTimer !== null || _hintsShown) {
      // Any other key while armed/shown means a real chord (Alt+Tab,
      // Option-typing, an Alt+X hotkey) — not a discoverability hold.
      disarmHints();
    }
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
    var modalOpen = blockingModalOpen();
    if (modalOpen && _sheetEl && !_sheetEl.classList.contains("hidden")) {
      // The cheatsheet is our own blocking modal: its toggle combo passes
      // back through so a second "?" press closes it (like the old per-page
      // popovers). Everything else stays suppressed while it owns the keyboard.
      if (!isTypingTarget(e.target)) {
        var sheetCombo = normalizeEvent(e);
        if (sheetCombo && resolvedCombos("global.cheatsheet").indexOf(sheetCombo) !== -1) {
          e.preventDefault();
          closeCheatsheet();
        }
      }
      return;
    }
    // A non-cheatsheet blocking modal (start launcher / settings) is open: fall
    // through, but the loop below only fires attachments flagged inModal — the
    // modal's own keyboard nav. Background page hotkeys stay dead.
    var combo = normalizeEvent(e);
    if (!combo) return;
    var ids = _comboIndex[combo];
    if (!ids) return;
    var typing = isTypingTarget(e.target);
    for (var n = 0; n < ids.length; n++) {
      var atts = _attachments[ids[n]] || [];
      for (var a = 0; a < atts.length; a++) {
        var att = atts[a];
        if (modalOpen && !att.inModal) continue;
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
    if (e.key === "Alt" && (_hintTimer !== null || _hintsShown)) {
      // Swallow the keyup when hints were shown so the Alt hold doesn't
      // also focus the browser menubar (Firefox/Windows). A quick tap
      // (released before the show delay) keeps native behavior.
      if (_hintsShown) e.preventDefault();
      disarmHints();
    }
    if (!_held.length) return;
    // Normalize to the combo token the keydown stored as baseKey — Space's
    // e.key is " ", which would never match "SPACE" without this.
    var key =
      e.key === " " || e.key === "Spacebar"
        ? "SPACE"
        : (e.key || "").length === 1
          ? e.key.toUpperCase()
          : e.key;
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
    disarmHints();
  }

  document.addEventListener("keydown", onDocKeydown);
  document.addEventListener("keyup", onDocKeyup);
  window.addEventListener("blur", onWindowBlur);

  // ---- Alt-hold hint chips ----
  //
  // Hold Alt for HINT_DELAY_MS and a small combo chip appears next to every
  // visible control tagged data-hotkey="<catalog id>" (optionally
  // data-hotkey-combo="<n>" to pick the Nth resolved combo). Release Alt,
  // press any other key, scroll, resize, or blur the window and they vanish.

  var HINT_DELAY_MS = 200;
  var _hintTimer = null;   // pending show timeout id
  var _hintLayer = null;   // fixed .hk-hints container, kept for reuse
  var _hintsShown = false;

  // Context action-hint providers: a page with a keyboard cursor registers a
  // function returning {anchor, entries: [{id, label}]} (or null while no
  // cursor is active). On Alt-hold, showHints() stacks one labeled chip per
  // entry vertically to the right of the anchor (e.g. Studio's cell browser:
  // "↩ Send to Artifacts" over "⇧↩ Send to Reel").
  var _actionHintProviders = [];

  function registerActionHints(fn) {
    if (typeof fn === "function") _actionHintProviders.push(fn);
  }

  function armHints() {
    if (_hintTimer !== null || _hintsShown) return;
    _hintTimer = window.setTimeout(showHints, HINT_DELAY_MS);
  }

  function disarmHints() {
    if (_hintTimer !== null) {
      window.clearTimeout(_hintTimer);
      _hintTimer = null;
    }
    hideHints();
  }

  // True when another element covers the control's center — so a hint chip would
  // otherwise float over whatever now sits on top (e.g. a dragged-up bottom panel
  // covering the video controls). The .hk-hints layer is pointer-events:none, so
  // it never registers as the occluder. elementFromPoint wants viewport coords,
  // so clamp the sample point into the visible box.
  function isOccluded(node, rect) {
    if (!document.elementFromPoint) return false;
    var cx = Math.min(Math.max(rect.left + rect.width / 2, 1), window.innerWidth - 1);
    var cy = Math.min(Math.max(rect.top + rect.height / 2, 1), window.innerHeight - 1);
    var hit = document.elementFromPoint(cx, cy);
    if (!hit) return false;
    return hit !== node && !node.contains(hit) && !hit.contains(node);
  }

  function showHints() {
    _hintTimer = null;
    if (_hintsShown || isTypingTarget(document.activeElement)) return;
    // When a modal owns the keyboard, scope hints to its own [data-hotkey]
    // controls (so background-page chips never float over the modal). A modal
    // we can't scope (bare body.modal-open with no registered root) suppresses
    // hints entirely, exactly as before.
    var modalRoot = (typeof getActiveModalRoot === "function") ? getActiveModalRoot() : null;
    if (blockingModalOpen() && !modalRoot) return;
    var scope = modalRoot || document;
    var nodes = scope.querySelectorAll("[data-hotkey]");
    // Read pass: measure every candidate before any DOM writes.
    var targets = [];
    for (var n = 0; n < nodes.length; n++) {
      var node = nodes[n];
      // A disabled control still gets a chip, dimmed, so the shortcut stays
      // discoverable while showing it is currently inert (e.g. "Stash" with an
      // empty queue).
      var dim = node.disabled === true;
      var combos = resolvedCombos(node.getAttribute("data-hotkey"));
      if (!combos.length) continue; // unknown id or user-disabled binding
      var combo = combos[parseInt(node.getAttribute("data-hotkey-combo") || "0", 10)] || combos[0];
      var rect = node.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) continue; // hidden (offsetParent is null inside fixed subheaders)
      if (rect.bottom < 0 || rect.top > window.innerHeight || rect.right < 0 || rect.left > window.innerWidth) continue;
      if (isOccluded(node, rect)) continue; // covered by another panel (e.g. a dragged-up bottom panel over the video controls)
      targets.push({ rect: rect, combo: combo, dim: dim });
    }
    // Context action-hint providers anchor to background-page controls, so skip
    // them while a modal owns the keyboard.
    for (var p = 0; !modalRoot && p < _actionHintProviders.length; p++) {
      var ctx = _actionHintProviders[p]();
      if (!ctx || !ctx.anchor || !ctx.entries) continue;
      var arect = ctx.anchor.getBoundingClientRect();
      if (arect.width === 0 && arect.height === 0) continue;
      if (arect.bottom < 0 || arect.top > window.innerHeight || arect.right < 0 || arect.left > window.innerWidth) continue;
      var stack = 0;
      for (var q = 0; q < ctx.entries.length; q++) {
        var ccombos = resolvedCombos(ctx.entries[q].id);
        if (!ccombos.length) continue; // user-disabled binding
        targets.push({ rect: arect, combo: ccombos[0], label: ctx.entries[q].label, stack: stack });
        stack++;
      }
    }
    if (!targets.length) return;
    if (!_hintLayer) {
      _hintLayer = el("div", "hk-hints hidden");
      document.body.appendChild(_hintLayer);
    }
    // Write pass: batch all chips into one fragment.
    var frag = document.createDocumentFragment();
    var chips = [];
    for (var c = 0; c < targets.length; c++) {
      var chip;
      if (targets[c].label) {
        chip = el("span", "hk-hint");
        chip.appendChild(fillKeycap(el("kbd", ""), targets[c].combo));
        chip.appendChild(el("span", "hk-hint-label", targets[c].label));
      } else {
        chip = el("span", "hk-hint");
        fillKeycap(chip, targets[c].combo);
      }
      if (targets[c].dim) chip.classList.add("hk-hint-dim");
      chips.push(frag.appendChild(chip));
    }
    _hintLayer.appendChild(frag);
    _hintLayer.classList.remove("hidden");
    // Second read pass (chip sizes), then one write pass placing each chip
    // badge-style on its target's top-right corner, clamped into the viewport.
    var sizes = [];
    for (var m = 0; m < chips.length; m++) {
      sizes.push({ w: chips[m].offsetWidth, h: chips[m].offsetHeight });
    }
    // Corner badges anchor to the anchor's vertical CENTER (minus a fixed
    // half-height) rather than its top edge, so controls of different heights
    // sharing a centered row — e.g. a <select> among icon buttons — still get
    // level chips. For a standard ~24 px control this still straddles the top
    // edge as before.
    var BADGE_RAISE = 12;
    for (var k = 0; k < chips.length; k++) {
      var r = targets[k].rect;
      var left, top;
      if (targets[k].label) {
        // Labeled action chip: to the right of the anchor, stacked downward;
        // flips to the anchor's left side when the viewport lacks room.
        left = r.right + 6;
        if (left + sizes[k].w > window.innerWidth - 4) left = r.left - sizes[k].w - 6;
        top = r.top + targets[k].stack * (sizes[k].h + 4);
      } else {
        left = r.right - sizes[k].w + 6;
        top = r.top + r.height / 2 - BADGE_RAISE - sizes[k].h / 2;
      }
      left = Math.max(4, Math.min(left, window.innerWidth - sizes[k].w - 4));
      top = Math.max(4, Math.min(top, window.innerHeight - sizes[k].h - 4));
      chips[k].style.left = left + "px";
      chips[k].style.top = top + "px";
    }
    // De-overlap corner badges: adjacent narrow buttons (e.g. undo ⌘Z / redo
    // ⌘⇧Z) yield chips wider than the buttons, so their badges collide. Spread
    // any colliding badge rightward to sit just past the previous one, so a
    // tight cluster reads as a row of chips. Labeled action chips already stack
    // downward, so they're excluded.
    var order = [];
    for (var bi = 0; bi < chips.length; bi++) {
      if (!targets[bi].label) order.push(bi);
    }
    order.sort(function (a, b2) {
      return parseFloat(chips[a].style.left) - parseFloat(chips[b2].style.left);
    });
    for (var s2 = 1; s2 < order.length; s2++) {
      var cur = order[s2];
      var prev = order[s2 - 1];
      var pRight = parseFloat(chips[prev].style.left) + sizes[prev].w;
      var pTop = parseFloat(chips[prev].style.top);
      var cLeft = parseFloat(chips[cur].style.left);
      var cTop = parseFloat(chips[cur].style.top);
      var sameRow = cTop < pTop + sizes[prev].h && cTop + sizes[cur].h > pTop;
      if (sameRow && cLeft < pRight + 4) {
        var nl = Math.min(pRight + 4, window.innerWidth - sizes[cur].w - 4);
        if (nl > cLeft) chips[cur].style.left = nl + "px";
      }
    }
    _hintsShown = true;
    // Chips don't track their anchors; any viewport change just hides them.
    window.addEventListener("scroll", disarmHints, true);
    window.addEventListener("resize", disarmHints);
  }

  function hideHints() {
    if (!_hintsShown) return;
    _hintsShown = false;
    window.removeEventListener("scroll", disarmHints, true);
    window.removeEventListener("resize", disarmHints);
    _hintLayer.classList.add("hidden");
    _hintLayer.textContent = "";
  }

  // ---- Shared "?" help button ----
  // Every page carries a [data-hotkeys-help] button (see hotkeys.css for the
  // shared look). Delegation instead of a per-page init: no load-order
  // assumptions, and JS-created buttons work automatically.

  document.addEventListener("click", function (e) {
    var t = e.target;
    var btn = t && t.closest ? t.closest("[data-hotkeys-help]") : null;
    if (btn) toggleCheatsheet();
  });

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
            keys.appendChild(fillKeycap(el("kbd", ""), combos[k]));
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
    closeCheatsheet: closeCheatsheet,
    registerActionHints: registerActionHints
  };
})();
