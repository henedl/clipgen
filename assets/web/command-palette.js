/* clipgen global command palette — command-palette.js
 *
 * Spotlight-style palette summoned with Cmd/Ctrl+Shift+P (or Cmd/Ctrl+K —
 * Firefox reserves Ctrl+Shift+P for private windows). Loaded on every hub
 * page after utils.js/topnav.js/settings-modal.js and before the page hub.
 *
 * Exposes exactly one global:
 *   window.ClipgenCommandPalette = { register, open, close, toggle }
 *
 * register(sourceId, providerOrArray) — pages contribute commands. A provider
 * is a function returning an array of commands, called on every open so
 * dynamic entries (participants, gated actions) stay fresh; a plain array is
 * wrapped in a constant provider. Re-registering a sourceId replaces it.
 *
 * Command shape:
 *   {
 *     id: "studio.generate",     // stable, required — keys the recents list
 *     title: "Generate clips",   // required
 *     subtitle: "…",             // optional second line
 *     icon: "play",              // heroicon stem in assets/icons/, optional
 *     keywords: "render build",  // optional extra match terms
 *     section: "Studio",         // group header when browsing
 *     hint: "G",                 // optional kbd hint (display only)
 *     run: function () {},       // required
 *     enabled: fn|bool,          // false = grayed row, Enter ignored
 *     visible: fn|bool,          // false = omitted entirely
 *   }
 *
 * Built-in providers: page navigation (read from the rendered topnav, so it
 * inherits the enabled-surface filtering), global chrome actions (settings /
 * theme / start panel / tooltips), and the page's TopNav quick actions via
 * ClipgenTopNav.getQuickActions({ refresh: true }).
 *
 * The overlay lifecycle rides openBlockingModal (utils.js): Escape, Tab trap,
 * backdrop click, restore-focus. Because openBlockingModal is a singleton,
 * open() bails while the settings modal is up (body.modal-open).
 */

(function () {
  "use strict";

  var RECENTS_PAGE = "commandPalette";
  var RECENTS_MAX = 8;

  var NAV_ICONS = {
    studio: "film",
    screenspace: "window",
    transcripts: "chat-bubble-left",
    workflows: "puzzle-piece",
    composer: "scissors",
    overview: "chart-bar",
  };

  var providers = []; // [{ id, fn }] in registration order
  var els = null;     // { overlay, panel, input, list, empty }
  var isOpen = false;
  var commands = [];  // prepared visible commands for the current open
  var rendered = [];  // commands currently in the list, in DOM order
  var selectedIndex = -1;

  // ---- Registry ----

  function register(sourceId, providerOrArray) {
    if (!sourceId) return;
    var fn = typeof providerOrArray === "function"
      ? providerOrArray
      : function () { return providerOrArray || []; };
    for (var i = 0; i < providers.length; i++) {
      if (providers[i].id === sourceId) {
        providers[i].fn = fn;
        return;
      }
    }
    providers.push({ id: sourceId, fn: fn });
  }

  // ---- Built-in providers ----

  function navCommands() {
    var out = [];
    var tabs = document.querySelectorAll(".topnav-tab");
    for (var i = 0; i < tabs.length; i++) {
      var tab = tabs[i];
      if (tab.getAttribute("aria-current") === "page") continue;
      var frontend = tab.getAttribute("data-frontend") || "";
      out.push({
        id: "nav." + frontend,
        title: "Go to " + tab.textContent,
        icon: NAV_ICONS[frontend] || "arrow-right",
        keywords: "navigate open page switch",
        section: "Navigate",
        run: (function (href) {
          return function () { location.href = href; };
        })(tab.href),
      });
    }
    return out;
  }

  function globalCommands() {
    return [
      {
        id: "global.settings",
        title: "Open Settings",
        icon: "cog-6-tooth",
        keywords: "preferences configure options",
        section: "Global",
        visible: function () { return typeof window.openSettingsModal === "function"; },
        run: function () { window.openSettingsModal({}); },
      },
      {
        id: "global.theme",
        title: "Toggle theme",
        icon: "moon",
        keywords: "dark light mode appearance",
        section: "Global",
        visible: function () { return !!document.getElementById("themeToggle"); },
        run: function () { document.getElementById("themeToggle").click(); },
      },
      {
        id: "global.start",
        title: "Open Start panel",
        icon: "home",
        keywords: "study picker first run",
        section: "Global",
        visible: function () {
          return !!(window.ClipgenStartOverlay && window.ClipgenStartOverlay.open);
        },
        run: function () { window.ClipgenStartOverlay.open(); },
      },
      {
        id: "global.tooltips",
        title: "Toggle cross-reference tooltips",
        icon: "chat-bubble-left-ellipsis",
        keywords: "xref hover badges",
        section: "Global",
        visible: function () { return !!document.getElementById("tooltipToggle"); },
        run: function () { document.getElementById("tooltipToggle").click(); },
      },
    ];
  }

  function quickActionCommands() {
    if (!window.ClipgenTopNav || !window.ClipgenTopNav.getQuickActions) return [];
    var items = window.ClipgenTopNav.getQuickActions({ refresh: true });
    var out = [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      if (item.divider || item.header || typeof item.action !== "function") continue;
      out.push({
        id: "qa." + (item.label || i),
        title: item.label || "",
        subtitle: item.title || "",
        icon: item.icon || "",
        section: "Quick actions",
        enabled: !item.disabled,
        run: item.action,
      });
    }
    return out;
  }

  // ---- Collection + matching ----

  function truthy(value, fallback) {
    if (value === undefined || value === null) return fallback;
    if (typeof value === "function") {
      try { return !!value(); } catch (_) { return fallback; }
    }
    return !!value;
  }

  function collectCommands() {
    var out = [];
    var sources = [
      { id: "nav", fn: navCommands },
      { id: "global", fn: globalCommands },
      { id: "quick-actions", fn: quickActionCommands },
    ].concat(providers);
    for (var s = 0; s < sources.length; s++) {
      var list = [];
      try { list = sources[s].fn() || []; } catch (e) {
        console.error("Command palette provider error (" + sources[s].id + "):", e);
      }
      for (var i = 0; i < list.length; i++) {
        var cmd = list[i];
        if (!cmd || !cmd.id || !cmd.title || typeof cmd.run !== "function") continue;
        if (!truthy(cmd.visible, true)) continue;
        cmd._index = out.length;
        cmd._enabled = truthy(cmd.enabled, true);
        cmd._titleLower = String(cmd.title).toLowerCase();
        cmd._titleWords = cmd._titleLower.split(/\s+/);
        cmd._kwLower = cmd.keywords ? String(cmd.keywords).toLowerCase() : "";
        cmd._metaLower =
          ((cmd.subtitle || "") + " " + (cmd.section || "")).toLowerCase();
        out.push(cmd);
      }
    }
    return out;
  }

  function scoreToken(cmd, token) {
    if (cmd._titleLower.indexOf(token) === 0) return 100;
    for (var i = 0; i < cmd._titleWords.length; i++) {
      if (cmd._titleWords[i].indexOf(token) === 0) return 80;
    }
    if (cmd._titleLower.indexOf(token) >= 0) return 60;
    if (cmd._kwLower && cmd._kwLower.indexOf(token) >= 0) return 40;
    if (cmd._metaLower && cmd._metaLower.indexOf(token) >= 0) return 25;
    return 0;
  }

  function scoreCommand(cmd, tokens) {
    var total = 0;
    for (var i = 0; i < tokens.length; i++) {
      var s = scoreToken(cmd, tokens[i]);
      if (s === 0) return 0; // every token must match somewhere
      total += s;
    }
    return total;
  }

  // ---- Recents ----

  function getRecents() {
    var stored = getStoredUIState(RECENTS_PAGE);
    return Array.isArray(stored.recents) ? stored.recents : [];
  }

  function pushRecent(id) {
    var next = [id];
    var prev = getRecents();
    for (var i = 0; i < prev.length && next.length < RECENTS_MAX; i++) {
      if (prev[i] !== id) next.push(prev[i]);
    }
    setStoredUIStateField(RECENTS_PAGE, "recents", next);
  }

  // ---- DOM ----

  function buildDom() {
    var overlay = document.createElement("div");
    overlay.className = "cmdp-overlay hidden";

    var panel = document.createElement("div");
    panel.className = "cmdp-panel";
    panel.setAttribute("role", "dialog");
    panel.setAttribute("aria-modal", "true");
    panel.setAttribute("aria-label", "Command palette");

    var inputRow = document.createElement("div");
    inputRow.className = "cmdp-input-row";
    var searchIcon = document.createElement("span");
    searchIcon.className = "cmdp-search-icon";
    searchIcon.setAttribute("aria-hidden", "true");
    applyIconMask(searchIcon, "magnifying-glass");
    var input = document.createElement("input");
    input.className = "cmdp-input";
    input.type = "text";
    input.placeholder = "Type a command…";
    input.autocomplete = "off";
    input.spellcheck = false;
    input.setAttribute("role", "combobox");
    input.setAttribute("aria-expanded", "true");
    input.setAttribute("aria-controls", "cmdpList");
    inputRow.appendChild(searchIcon);
    inputRow.appendChild(input);

    var list = document.createElement("div");
    list.id = "cmdpList";
    list.className = "cmdp-list";
    list.setAttribute("role", "listbox");

    var empty = document.createElement("div");
    empty.className = "cmdp-empty hidden";
    empty.textContent = "No matching commands";

    var footer = document.createElement("div");
    footer.className = "cmdp-footer";
    footer.innerHTML =
      "<span><kbd>↑</kbd><kbd>↓</kbd> navigate</span>" +
      "<span><kbd>↵</kbd> run</span>" +
      "<span><kbd>esc</kbd> close</span>";

    panel.appendChild(inputRow);
    panel.appendChild(list);
    panel.appendChild(empty);
    panel.appendChild(footer);
    overlay.appendChild(panel);
    document.body.appendChild(overlay);

    input.addEventListener("input", render);
    input.addEventListener("keydown", onInputKeydown);
    list.addEventListener("mousemove", onListMousemove);
    list.addEventListener("click", onListClick);

    els = { overlay: overlay, panel: panel, input: input, list: list, empty: empty };
  }

  function buildItem(cmd, renderedIdx) {
    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "cmdp-item";
    btn.setAttribute("role", "option");
    btn.setAttribute("data-idx", String(renderedIdx));
    if (!cmd._enabled) btn.disabled = true;

    var icon = document.createElement("span");
    icon.className = "cmdp-item-icon";
    icon.setAttribute("aria-hidden", "true");
    if (cmd.icon) applyIconMask(icon, cmd.icon);
    else icon.style.visibility = "hidden";
    btn.appendChild(icon);

    var text = document.createElement("span");
    text.className = "cmdp-item-text";
    var title = document.createElement("span");
    title.className = "cmdp-item-title";
    title.textContent = cmd.title;
    text.appendChild(title);
    if (cmd.subtitle) {
      var subtitle = document.createElement("span");
      subtitle.className = "cmdp-item-subtitle";
      subtitle.textContent = cmd.subtitle;
      text.appendChild(subtitle);
    }
    btn.appendChild(text);

    if (cmd.hint) {
      var hint = document.createElement("kbd");
      hint.className = "cmdp-item-hint";
      hint.textContent = cmd.hint;
      btn.appendChild(hint);
    }
    return btn;
  }

  function buildSectionLabel(label) {
    var div = document.createElement("div");
    div.className = "cmdp-section";
    div.textContent = label;
    return div;
  }

  // ---- Rendering ----

  function render() {
    var query = els.input.value.replace(/^\s+|\s+$/g, "").toLowerCase();
    var frag = document.createDocumentFragment();
    rendered = [];

    if (!query) {
      // Recents first, then everything grouped by section in provider order.
      var recentIds = getRecents();
      var byId = {};
      var i;
      for (i = 0; i < commands.length; i++) byId[commands[i].id] = commands[i];
      var recents = [];
      for (i = 0; i < recentIds.length; i++) {
        if (byId[recentIds[i]]) recents.push(byId[recentIds[i]]);
      }
      if (recents.length) {
        frag.appendChild(buildSectionLabel("Recent"));
        appendItems(frag, recents);
      }
      var currentSection = null;
      for (i = 0; i < commands.length; i++) {
        var section = commands[i].section || "Commands";
        if (section !== currentSection) {
          frag.appendChild(buildSectionLabel(section));
          currentSection = section;
        }
        appendItems(frag, [commands[i]]);
      }
    } else {
      var tokens = query.split(/\s+/);
      var scored = [];
      for (var c = 0; c < commands.length; c++) {
        var score = scoreCommand(commands[c], tokens);
        if (score > 0) scored.push({ cmd: commands[c], score: score });
      }
      scored.sort(function (a, b) {
        if (b.score !== a.score) return b.score - a.score;
        return a.cmd._index - b.cmd._index; // stable by insertion order
      });
      var matches = [];
      for (var m = 0; m < scored.length; m++) matches.push(scored[m].cmd);
      appendItems(frag, matches);
    }

    els.list.innerHTML = "";
    els.list.appendChild(frag);
    els.empty.classList.toggle("hidden", rendered.length > 0);
    selectedIndex = -1;
    moveSelection(1); // land on the first enabled row
  }

  function appendItems(frag, cmds) {
    for (var i = 0; i < cmds.length; i++) {
      frag.appendChild(buildItem(cmds[i], rendered.length));
      rendered.push(cmds[i]);
    }
  }

  function itemEls() {
    return els.list.querySelectorAll(".cmdp-item");
  }

  function setSelection(idx, scroll) {
    var items = itemEls();
    if (selectedIndex >= 0 && items[selectedIndex]) {
      items[selectedIndex].classList.remove("is-selected");
      items[selectedIndex].setAttribute("aria-selected", "false");
    }
    selectedIndex = idx;
    if (idx >= 0 && items[idx]) {
      items[idx].classList.add("is-selected");
      items[idx].setAttribute("aria-selected", "true");
      if (scroll && items[idx].scrollIntoView) {
        items[idx].scrollIntoView({ block: "nearest" });
      }
    }
  }

  function moveSelection(dir) {
    if (!rendered.length) { setSelection(-1, false); return; }
    // Step to the next enabled row, wrapping; give up after a full lap.
    var idx = selectedIndex;
    for (var hops = 0; hops < rendered.length; hops++) {
      idx = (idx + dir + rendered.length) % rendered.length;
      if (rendered[idx]._enabled) { setSelection(idx, true); return; }
    }
    setSelection(-1, false);
  }

  // ---- Events ----

  function onInputKeydown(e) {
    var handled = true;
    if (e.key === "ArrowDown") moveSelection(1);
    else if (e.key === "ArrowUp") moveSelection(-1);
    else if (e.key === "Home") { selectedIndex = -1; moveSelection(1); }
    else if (e.key === "End") { selectedIndex = 0; moveSelection(-1); }
    else if (e.key === "Enter") {
      if (selectedIndex >= 0 && rendered[selectedIndex]) {
        runCommand(rendered[selectedIndex]);
      }
    } else handled = false;
    if (handled) {
      e.preventDefault();
      e.stopPropagation();
    }
  }

  function onListMousemove(e) {
    // mousemove, not mouseover: hover must not steal the selection while the
    // list scrolls under a keyboard-driven cursor.
    var btn = e.target.closest ? e.target.closest(".cmdp-item") : null;
    if (!btn || btn.disabled) return;
    var idx = parseInt(btn.getAttribute("data-idx"), 10);
    if (idx !== selectedIndex) setSelection(idx, false);
  }

  function onListClick(e) {
    var btn = e.target.closest ? e.target.closest(".cmdp-item") : null;
    if (!btn || btn.disabled) return;
    var idx = parseInt(btn.getAttribute("data-idx"), 10);
    if (rendered[idx]) runCommand(rendered[idx]);
  }

  function runCommand(cmd) {
    if (!cmd._enabled) return;
    pushRecent(cmd.id);
    close();
    // Run after close so restore-focus can't clobber commands that set focus
    // themselves (e.g. "Focus transcript search").
    setTimeout(function () {
      try { cmd.run(); } catch (e) { console.error("Command palette action error:", e); }
    }, 0);
  }

  // ---- Open/close ----

  function open() {
    if (isOpen) return;
    // openBlockingModal is a singleton; don't steal the settings modal's trap.
    if (document.body.classList.contains("modal-open")) return;
    if (!els) buildDom();
    commands = collectCommands();
    els.input.value = "";
    isOpen = true;
    els.overlay.classList.remove("hidden");
    render();
    openBlockingModal(els.overlay, {
      trapFocus: true,
      restoreFocus: true,
      onEscape: close,
      onBackdropClick: close,
    });
    requestAnimationFrame(function () {
      els.panel.classList.add("is-in");
    });
  }

  function close() {
    if (!isOpen) return;
    isOpen = false;
    closeBlockingModal(els.overlay);
    els.panel.classList.remove("is-in");
    els.overlay.classList.add("hidden");
    els.input.value = "";
  }

  function toggle() {
    if (isOpen) close();
    else open();
  }

  // ---- Summon chord ----

  document.addEventListener("keydown", function (e) {
    var k = (e.key || "").toLowerCase();
    var mod = e.metaKey || e.ctrlKey;
    var primary = mod && e.shiftKey && !e.altKey && k === "p";
    var secondary = mod && !e.shiftKey && !e.altKey && k === "k";
    if (!primary && !secondary) return;
    e.preventDefault();
    e.stopPropagation();
    toggle();
  }, true);

  window.ClipgenCommandPalette = {
    register: register,
    open: open,
    close: close,
    toggle: toggle,
  };
})();
