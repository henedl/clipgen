/* clipgen global command palette — command-palette.js
 *
 * Spotlight-style palette summoned via the shared hotkey registry: catalog
 * action "global.palette" in hotkeys.js, default Mod+Shift+P or Mod+K
 * (Firefox reserves Ctrl+Shift+P for private windows), rebindable in
 * Settings → Hotkeys. Loaded on every hub page after
 * utils.js/hotkeys.js/topnav.js/settings-modal.js and before the page hub.
 *
 * Command ids use ":" separators ("studio:generate") on purpose — dotted
 * "a.b" ids are reserved for hotkey catalog actions and guarded by
 * tests/test_hotkeys_frontend_source.py's registered-id scan.
 *
 * Exposes exactly one global:
 *   window.ClipgenCommandPalette = { register, setParticipants, open, close, toggle }
 *
 * register(sourceId, providerOrArray) — pages contribute commands. A provider
 * is a function returning an array of commands, called on every open so
 * dynamic entries (participants, gated actions) stay fresh; a plain array is
 * wrapped in a constant provider. Re-registering a sourceId replaces it.
 *
 * setParticipants(fn) — pages hand over their participant-id list; the
 * palette turns it into cross-page "Open Pxx in <Page>" commands for every
 * #Pxx-hash destination except the current page. Wording convention:
 * "Jump to … in <here>" stays on the page, "Open … in <there>" navigates,
 * as do the "Go to <Page>" / "Open <Tab> in <Page>" nav commands.
 *
 * Command shape:
 *   {
 *     id: "studio:generate",     // stable, required — keys the recents list
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
 * Built-in providers: page navigation incl. cross-page tab deep links (read
 * from the rendered topnav, so it inherits the enabled-surface filtering),
 * global chrome actions (settings / theme / start panel / tooltips), the
 * page's TopNav quick actions via ClipgenTopNav.getQuickActions({ refresh:
 * true }), and the cross-page participant jumps from setParticipants.
 *
 * The overlay lifecycle rides openBlockingModal (utils.js): Escape, Tab trap,
 * backdrop click, restore-focus. Because openBlockingModal is a singleton,
 * open() bails while another overlay is up (body.modal-open for the settings
 * modal, isBlockingModalOpen() for everything that holds the shared trap).
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

  // Cross-page tab deep links, consumed as /PAGE/#tab=KEY by clipgenHashTab()
  // on the receiving page (Studio + Overview route it through their stored
  // active-tab restore; Transcripts clicks the panel-tab button).
  var NAV_TABS = {
    studio: [
      { key: "intake", label: "Screenspace Intake" },
      { key: "transcript-intake", label: "Transcript Intake" },
      { key: "composer-intake", label: "Composer Intake" },
      { key: "mindnode-intake", label: "MindNode Intake" },
    ],
    transcripts: [
      { key: "summary", label: "Summary" },
      { key: "friction", label: "Friction" },
    ],
    overview: [
      { key: "metadata", label: "Metadata" },
      { key: "convergence", label: "Convergence" },
    ],
  };

  // Pages that pre-select a participant from a /PAGE/#Pxx hash
  // (clipgenHashParticipant in utils.js).
  var PARTICIPANT_PAGES = [
    { id: "transcripts", label: "Transcripts" },
    { id: "screenspace", label: "Screenspace" },
    { id: "composer", label: "Composer" },
  ];

  var providers = []; // [{ id, fn }] in registration order
  var participantSource = null; // page-set fn returning participant id strings
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

  // setParticipants(fn) — each hub hands the palette its participant-id list
  // (a function returning an array of id strings, read on every open). Feeds
  // the built-in cross-page "Open Pxx in <Page>" commands so participant
  // jumps exist on every page, not just the ones that hold participants.
  function setParticipants(fn) {
    participantSource = typeof fn === "function" ? fn : null;
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
        id: "nav:" + frontend,
        title: "Go to " + tab.textContent,
        icon: NAV_ICONS[frontend] || "arrow-right",
        keywords: "navigate open page switch",
        section: "Navigate",
        run: (function (href) {
          return function () { location.href = href; };
        })(tab.href),
      });
      var pageTabs = NAV_TABS[frontend] || [];
      for (var t = 0; t < pageTabs.length; t++) {
        out.push({
          id: "nav:" + frontend + ":tab-" + pageTabs[t].key,
          title: "Open " + pageTabs[t].label + " in " + tab.textContent,
          icon: NAV_ICONS[frontend] || "arrow-right",
          keywords: "navigate tab " + frontend,
          section: "Navigate",
          run: (function (href, key) {
            return function () { location.href = href + "#tab=" + key; };
          })(tab.href, pageTabs[t].key),
        });
      }
    }
    return out;
  }

  // Cross-page participant jumps ("Open P07 in Transcripts"). Same-page jumps
  // ("Jump to P07 in <Page>") stay page-registered — they select in place
  // instead of navigating.
  function participantNavCommands() {
    if (!participantSource) return [];
    var pids = [];
    try { pids = participantSource() || []; } catch (_) {}
    var current = currentFrontend();
    var out = [];
    for (var i = 0; i < pids.length; i++) {
      for (var d = 0; d < PARTICIPANT_PAGES.length; d++) {
        var dest = PARTICIPANT_PAGES[d];
        if (dest.id === current) continue;
        out.push({
          id: "nav:p:" + dest.id + ":" + pids[i],
          title: "Open " + pids[i] + " in " + dest.label,
          icon: "user-circle",
          keywords: "participant navigate " + dest.id,
          section: "Participants",
          run: (function (destId, pid) {
            return function () { location.href = "/" + destId + "/#" + pid; };
          })(dest.id, pids[i]),
        });
      }
    }
    return out;
  }

  function currentFrontend() {
    var active = document.querySelector('.topnav-tab[aria-current="page"]');
    return active ? active.getAttribute("data-frontend") || "" : "";
  }

  function globalCommands() {
    return [
      {
        id: "global:settings",
        title: "Open Settings",
        icon: "cog-6-tooth",
        keywords: "preferences configure options",
        section: "Global",
        visible: function () { return typeof window.openSettingsModal === "function"; },
        run: function () { window.openSettingsModal({}); },
      },
      {
        id: "global:theme",
        title: "Toggle theme",
        icon: "moon",
        keywords: "dark light mode appearance",
        section: "Global",
        visible: function () { return !!document.getElementById("themeToggle"); },
        run: function () { document.getElementById("themeToggle").click(); },
      },
      {
        id: "global:start",
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
        id: "global:crossrefs",
        title: "Toggle cross-references",
        icon: "chat-bubble-left-ellipsis",
        keywords: "xref hover badges overlap tooltips",
        section: "Global",
        visible: function () { return typeof setCrossReferences === "function"; },
        run: function () { setCrossReferences(!CLIPGEN_CONFIG.crossReferences); },
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
        id: "qa:" + (item.label || i),
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
    // Participant nav runs last so its "Participants" section lands directly
    // after the page's own participant jumps (the page provider lists them
    // last) — the browse view then shows one contiguous Participants group.
    var sources = [
      { id: "nav", fn: navCommands },
      { id: "global", fn: globalCommands },
      { id: "quick-actions", fn: quickActionCommands },
    ].concat(providers, [{ id: "nav-participants", fn: participantNavCommands }]);
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
    if (isToggleChord(e)) {
      e.preventDefault();
      e.stopPropagation();
      close();
      return;
    }
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
    // openBlockingModal is a singleton; never steal an existing trap. Two
    // checks because the overlays split: the settings modal sets
    // body.modal-open but doesn't use openBlockingModal, while Studio's
    // gallery/status/confirm and Transcripts' install dialog do the reverse.
    if (document.body.classList.contains("modal-open")) return;
    if (isBlockingModalOpen()) return;
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
  //
  // Lives in the shared hotkey registry (catalog id "global.palette", default
  // Mod+Shift+P / Mod+K, rebindable). allowInInput keeps the Spotlight behavior:
  // the chord is deliberate, so it fires even while typing. The dispatcher
  // suppresses all combos while a blocking modal is open — the palette included —
  // so toggling *closed* falls to the palette input's own keydown handler
  // (isToggleChord), as the hotkeys cheatsheet does with its toggle.

  function isToggleChord(e) {
    if (!window.ClipgenHotkeys) return false;
    var combo = window.ClipgenHotkeys.normalizeEvent(e);
    if (!combo) return false;
    return window.ClipgenHotkeys.resolvedCombos("global.palette").indexOf(combo) !== -1;
  }

  if (window.ClipgenHotkeys) {
    window.ClipgenHotkeys.register([
      { id: "global.palette", handler: toggle, allowInInput: true },
    ]);
  }

  window.ClipgenCommandPalette = {
    register: register,
    setParticipants: setParticipants,
    open: open,
    close: close,
    toggle: toggle,
  };
})();
