/* Start overlay — Direction B redesign.
 *
 * Two-column launcher mounted by Studio / Screenspace / Transcripts. Drives:
 *   • brand-mark + wordmark intro (reuses window.clipgenInitBrandMark; cascade
 *     plays once per browser session, gated by the existing sessionStorage flag)
 *   • section cascade-in (220ms base, 80ms stagger)
 *   • backdrop blur via --host-blur / --veil-alpha on the overlay root
 *   • folder + spreadsheet picker (Google / Excel / No spreadsheet)
 *   • Extras tabs: tools / changelog / about
 *   • persistence via the existing /api/start-settings endpoint
 *
 * Public API on window.ClipgenStartOverlay:
 *   open()    — show the overlay
 *   close()   — hide the overlay
 *   isOpen()  — boolean
 */

(function () {
  "use strict";

  var DISMISSED_KEY = "clipgen.startOverlayDismissed";
  var INTRO_BLUR_PX = 14;
  var INTRO_VEIL_ALPHA = 0.55;
  var CASCADE_BASE_MS = 220;
  var CASCADE_STEP_MS = 80;
  // Google-auth poll lifecycle: the server's /api/spreadsheets/google/auth
  // launches an OAuth flow in a daemon thread; we poll /api/spreadsheets/google
  // until it reports `authenticated`, an `auth_error`, or we exceed the budget.
  var GOOGLE_POLL_TIMEOUT_MS = 90 * 1000;
  var GOOGLE_POLL_INTERVAL_MS = 1500;
  var GOOGLE_POLL_RETRY_MS = 2500;

  function sessionDismissed() {
    try { return sessionStorage.getItem(DISMISSED_KEY) === "1"; } catch (_) { return false; }
  }
  function markDismissed() {
    try { sessionStorage.setItem(DISMISSED_KEY, "1"); } catch (_) { /* ignore */ }
  }

  function shouldAutoOpen(status) {
    if (sessionDismissed()) return false;
    if (status && status.sheet_loaded) return false;
    var path = (window.location.pathname || "").toLowerCase();
    var isVideoTool =
      path.indexOf("/screenspace/") === 0 ||
      path.indexOf("/transcripts/") === 0 ||
      path.indexOf("/workflows/") === 0 ||
      path.indexOf("/composer/") === 0;
    if (isVideoTool && status && (status.videos_in_input || 0) > 0) return false;
    return true;
  }

  var state = {
    mounted: false,
    open: false,
    sheetLoaded: false,
    selection: null,        // { type, id_or_path, label } | null
    persistEnabled: true,
    googlePollTimer: null,
    googlePollDeadline: 0,
    confirmInFlight: false,
    activeTab: "google",    // 'google' | 'excel' | 'none'
    extrasTab: "tools",     // 'tools' | 'updates' | 'about'
    changelogLoaded: false,
    aboutLoaded: false,
    googleSheets: [],
    excelFiles: [],
    statusData: null,
    recentProjects: [],
    // Baseline = the (input, output, sheet selection, tab) snapshot at the
    // moment the overlay opened (or when a recent project was clicked).
    // Drives the .is-loaded / .is-dirty glow on the path-input and
    // sheet-card chrome.
    baseline: {
      input: "",
      output: "",
      sheetTab: "google",
      sheetKey: "",        // type|id_or_path or empty if no sheet
    },
  };

  var root = null;
  var els = {};

  function on(target, evt, fn) {
    if (target) target.addEventListener(evt, fn);
  }

  function show(node, visible) {
    if (!node) return;
    if (visible) node.classList.remove("hidden");
    else node.classList.add("hidden");
  }

  function setHidden(node, hidden) {
    if (!node) return;
    if (hidden) node.setAttribute("hidden", "");
    else node.removeAttribute("hidden");
  }

  function applyIcons(scope) {
    applyIconMasksIn(scope || root, { selector: ".so-icon[data-icon]" });
  }

  function mount() {
    if (state.mounted) return Promise.resolve();
    var slot = document.querySelector("start-overlay-mount");
    if (!slot) return Promise.resolve();
    return fetch("start-overlay.html")
      .then(function (r) {
        if (!r.ok) throw new Error("Server error " + r.status);
        return r.text();
      })
      .then(function (html) {
        var wrap = document.createElement("div");
        wrap.innerHTML = html;
        var node = wrap.firstElementChild;
        slot.replaceWith(node);
        root = node;
        cacheEls();
        applyIcons();
        bind();
        // The page-level clipgenInitBrandMark ran on DOMContentLoaded before
        // the overlay was inserted, so re-run it now to hydrate the rail mark.
        // It will skip the cascade animation if the session flag was already
        // set by an earlier hydration of a topnav brand-mark on this page.
        if (typeof window.clipgenInitBrandMark === "function") {
          window.clipgenInitBrandMark();
        }
        state.mounted = true;
      })
      .catch(function (err) {
        console.error("Start overlay: failed to mount", err);
      });
  }

  function cacheEls() {
    els.panel = root.querySelector('[data-role="panel"]');
    els.backdrop = root.querySelector('[data-role="backdrop"]');
    els.closeBtn = root.querySelector('[data-role="close"]');
    els.wordmark = root.querySelector('[data-role="wordmark"]');
    els.mark = root.querySelector(".rail__mark");

    els.cascades = root.querySelectorAll(".cascade-in");

    els.railRecents = root.querySelector('[data-role="rail-recents"]');

    els.inputField = root.querySelector('[data-role="input-field"]');
    els.outputField = root.querySelector('[data-role="output-field"]');
    els.inputDir = root.querySelector("#startInputDir");
    els.outputDir = root.querySelector("#startOutputDir");
    els.inputRecentsBtn = root.querySelector('[data-role="input-recents-toggle"]');
    els.outputRecentsBtn = root.querySelector('[data-role="output-recents-toggle"]');
    els.inputRecentsList = root.querySelector('[data-role="input-recents-list"]');
    els.outputRecentsList = root.querySelector('[data-role="output-recents-list"]');
    els.inputBrowseBtn = root.querySelector('[data-role="input-browse"]');
    els.outputBrowseBtn = root.querySelector('[data-role="output-browse"]');

    els.tabs = root.querySelectorAll(".sheet-card__tab");
    els.googlePanel = root.querySelector('[data-tabpanel="google"]');
    els.excelPanel = root.querySelector('[data-tabpanel="excel"]');
    els.nonePanel = root.querySelector('[data-tabpanel="none"]');
    els.googleStatus = root.querySelector('[data-role="google-status"]');
    els.googlePicker = root.querySelector('[data-role="google-picker"]');
    els.googlePickerTrigger = root.querySelector('[data-role="google-picker-trigger"]');
    els.googlePickerLabel = root.querySelector('[data-role="google-picker-label"]');
    els.googlePickerMenu = root.querySelector('[data-role="google-picker-menu"]');
    els.googlePaste = root.querySelector("#startGooglePaste");
    els.excelStatus = root.querySelector('[data-role="excel-status"]');
    els.excelPicker = root.querySelector('[data-role="excel-picker"]');
    els.excelPickerTrigger = root.querySelector('[data-role="excel-picker-trigger"]');
    els.excelPickerLabel = root.querySelector('[data-role="excel-picker-label"]');
    els.excelPickerMenu = root.querySelector('[data-role="excel-picker-menu"]');
    els.excelPaste = root.querySelector("#startExcelPaste");

    els.extrasTabs = root.querySelectorAll(".extras-tabs__tab");
    els.extrasPanels = {
      tools: root.querySelector('[data-extras-panel="tools"]'),
      updates: root.querySelector('[data-extras-panel="updates"]'),
      about: root.querySelector('[data-extras-panel="about"]'),
    };
    els.updatesBadge = root.querySelector('[data-role="updates-badge"]');
    els.changelogList = root.querySelector('[data-role="changelog-list"]');
    els.aboutGrid = root.querySelector('[data-role="about-grid"]');

    els.persist = root.querySelector("#startPersistEnabled");
    els.confirmBtn = root.querySelector('[data-role="confirm"]');
  }

  function bind() {
    on(els.closeBtn, "click", close);
    on(els.confirmBtn, "click", confirm);
    on(els.backdrop, "click", close);

    // Sheet-card tab strip
    Array.prototype.forEach.call(els.tabs, function (tab) {
      on(tab, "click", function () { setTab(tab.getAttribute("data-tab")); });
    });

    // Extras tab strip
    Array.prototype.forEach.call(els.extrasTabs, function (tab) {
      on(tab, "click", function () { setExtrasTab(tab.getAttribute("data-extras-tab")); });
    });

    on(els.inputRecentsBtn, "click", function (e) {
      e.stopPropagation();
      toggleRecents("input");
    });
    on(els.outputRecentsBtn, "click", function (e) {
      e.stopPropagation();
      toggleRecents("output");
    });
    on(els.inputBrowseBtn, "click", function () { browseFolder("input"); });
    on(els.outputBrowseBtn, "click", function () { browseFolder("output"); });

    on(els.inputDir, "input", function () {
      clearFieldError(els.inputField);
      applyFieldStates();
    });
    on(els.outputDir, "input", function () {
      clearFieldError(els.outputField);
      applyFieldStates();
    });

    on(els.googlePickerTrigger, "click", function (e) {
      e.stopPropagation();
      togglePicker("google");
    });
    on(els.excelPickerTrigger, "click", function (e) {
      e.stopPropagation();
      togglePicker("excel");
    });

    on(document, "click", function (e) {
      // Bail when the overlay isn't open — the listener stays bound for the
      // page's lifetime (mount() is once-per-page; close() just hides),
      // so guard explicitly to avoid pointless work on every page click.
      if (!root || !state.open) return;
      closeRecentsIfOutside(e);
      closePickersIfOutside(e);
    });

    on(els.persist, "change", function () {
      state.persistEnabled = !!els.persist.checked;
      apiPost("/api/start-settings", { persist_enabled: state.persistEnabled })
        .catch(function (err) { console.error("Persist toggle failed", err); });
    });

    on(els.googlePaste, "input", function () {
      var v = (els.googlePaste.value || "").trim();
      if (v) {
        setSelection({ type: "google", id_or_path: v, label: parseLabelFromGoogleInput(v) });
      } else if (state.selection && state.selection.type === "google") {
        renderGoogleList(state.googleSheets || []);
      }
    });

    on(els.excelPaste, "input", function () {
      var v = (els.excelPaste.value || "").trim();
      if (v) {
        setSelection({ type: "excel", id_or_path: v, label: v.split("/").pop() || v });
      } else if (state.selection && state.selection.type === "excel") {
        renderExcelList(state.excelFiles || []);
      }
    });

    // Keyboard: Escape (close) and Tab containment are owned by the shared
    // blocking-modal trap (see open()); the rest are shared-registry hotkeys so
    // holding Alt reveals them and they list in the "?" cheatsheet. Gated by
    // inModal + `when` so they fire only while the launcher owns the keyboard.
    if (window.ClipgenHotkeys) {
      var isOpen = function () { return state.open; };
      ClipgenHotkeys.register([
        { id: "start.tabGoogle",    inModal: true, when: isOpen, handler: function () { setTab("google"); } },
        { id: "start.tabExcel",     inModal: true, when: isOpen, handler: function () { setTab("excel"); } },
        { id: "start.tabNone",      inModal: true, when: isOpen, handler: function () { setTab("none"); } },
        { id: "start.browseInput",  inModal: true, when: isOpen, handler: function () { browseFolder("input"); } },
        { id: "start.browseOutput", inModal: true, when: isOpen, handler: function () { browseFolder("output"); } },
        { id: "start.confirm",      inModal: true, allowInInput: true, when: isOpen, handler: function () { confirm(); } }
      ]);
    }
  }

  function parseLabelFromGoogleInput(s) {
    if (/^https?:\/\//i.test(s)) return "Google Sheets URL";
    return s;
  }

  function closeRecentsIfOutside(e) {
    if (els.inputRecentsList && !els.inputRecentsList.classList.contains("hidden")) {
      if (!els.inputRecentsList.contains(e.target) && e.target !== els.inputRecentsBtn && !els.inputRecentsBtn.contains(e.target)) {
        hideRecents("input");
      }
    }
    if (els.outputRecentsList && !els.outputRecentsList.classList.contains("hidden")) {
      if (!els.outputRecentsList.contains(e.target) && e.target !== els.outputRecentsBtn && !els.outputRecentsBtn.contains(e.target)) {
        hideRecents("output");
      }
    }
  }

  function toggleRecents(kind) {
    var list = kind === "input" ? els.inputRecentsList : els.outputRecentsList;
    if (!list) return;
    if (list.classList.contains("hidden")) {
      // Hide the other before showing this one.
      hideRecents(kind === "input" ? "output" : "input");
      showRecents(kind);
    } else {
      hideRecents(kind);
    }
  }

  function showRecents(kind) {
    var list = kind === "input" ? els.inputRecentsList : els.outputRecentsList;
    var field = kind === "input" ? els.inputField : els.outputField;
    if (!list) return;
    list.classList.remove("hidden");
    if (field) field.classList.add("has-popover");
  }

  function hideRecents(kind) {
    var list = kind === "input" ? els.inputRecentsList : els.outputRecentsList;
    var field = kind === "input" ? els.inputField : els.outputField;
    if (!list) return;
    list.classList.add("hidden");
    if (field) field.classList.remove("has-popover");
  }

  // ---- Native folder browse ----

  function browseFolder(kind) {
    var input = kind === "input" ? els.inputDir : els.outputDir;
    var btn = kind === "input" ? els.inputBrowseBtn : els.outputBrowseBtn;
    if (!input) return;
    if (btn) btn.disabled = true;
    var payload = { initial: (input.value || "").trim() };
    apiPost("/api/folder-picker", payload)
      .then(function (r) {
        if (r && r.ok && r.path) {
          input.value = r.path;
          // Fire an input event so the dirty/loaded glow + clear-error logic
          // run as if the user typed the path.
          input.dispatchEvent(new Event("input", { bubbles: true }));
          input.focus();
        } else {
          // No path returned — either the user cancelled or the platform has
          // no native dialog. Falling back to focusing the field is the same
          // behaviour as before this endpoint existed.
          input.focus();
        }
      })
      .catch(function (err) {
        console.warn("Folder picker failed", err);
        input.focus();
      })
      .finally(function () {
        if (btn) btn.disabled = false;
      });
  }

  // ---- Sheet picker (Google + Excel dropdowns) ----

  function togglePicker(kind) {
    var menu = kind === "google" ? els.googlePickerMenu : els.excelPickerMenu;
    if (!menu) return;
    if (menu.classList.contains("hidden")) {
      // Close the other picker before opening this one.
      closePicker(kind === "google" ? "excel" : "google");
      openPicker(kind);
    } else {
      closePicker(kind);
    }
  }

  function openPicker(kind) {
    var menu = kind === "google" ? els.googlePickerMenu : els.excelPickerMenu;
    var trigger = kind === "google" ? els.googlePickerTrigger : els.excelPickerTrigger;
    if (!menu || !trigger) return;
    menu.classList.remove("hidden");
    trigger.setAttribute("aria-expanded", "true");
    liftSectionForPicker(trigger, true);
  }

  function closePicker(kind) {
    var menu = kind === "google" ? els.googlePickerMenu : els.excelPickerMenu;
    var trigger = kind === "google" ? els.googlePickerTrigger : els.excelPickerTrigger;
    if (!menu || !trigger) return;
    menu.classList.add("hidden");
    trigger.setAttribute("aria-expanded", "false");
    liftSectionForPicker(trigger, false);
  }

  function liftSectionForPicker(trigger, on) {
    var section = trigger && trigger.closest(".start-section");
    if (!section) return;
    section.classList.toggle("has-picker-open", !!on);
  }

  function closePickersIfOutside(e) {
    if (els.googlePicker && !els.googlePicker.contains(e.target) &&
        els.googlePickerMenu && !els.googlePickerMenu.classList.contains("hidden")) {
      closePicker("google");
    }
    if (els.excelPicker && !els.excelPicker.contains(e.target) &&
        els.excelPickerMenu && !els.excelPickerMenu.classList.contains("hidden")) {
      closePicker("excel");
    }
  }

  function updatePickerLabel(kind, label) {
    var labelEl = kind === "google" ? els.googlePickerLabel : els.excelPickerLabel;
    var trigger = kind === "google" ? els.googlePickerTrigger : els.excelPickerTrigger;
    if (!labelEl || !trigger) return;
    if (label) {
      labelEl.textContent = label;
      trigger.classList.remove("is-placeholder");
    } else {
      labelEl.textContent = kind === "google" ? "Select a Google Sheet…" : "Select an Excel file…";
      trigger.classList.add("is-placeholder");
    }
  }

  // ---- Tabs ----

  function setTab(name) {
    state.activeTab = name;
    Array.prototype.forEach.call(els.tabs, function (tab) {
      var active = tab.getAttribute("data-tab") === name;
      tab.classList.toggle("is-active", active);
      tab.setAttribute("aria-selected", active ? "true" : "false");
    });
    setHidden(els.googlePanel, name !== "google");
    setHidden(els.excelPanel, name !== "excel");
    setHidden(els.nonePanel, name !== "none");
    if (name === "none") {
      // Clear any sheet selection — user is opting out.
      state.selection = null;
    }
    clearSheetError();
    applyFieldStates();
  }

  function setExtrasTab(name) {
    state.extrasTab = name;
    Array.prototype.forEach.call(els.extrasTabs, function (tab) {
      var active = tab.getAttribute("data-extras-tab") === name;
      tab.classList.toggle("is-active", active);
      tab.setAttribute("aria-selected", active ? "true" : "false");
    });
    setHidden(els.extrasPanels.tools, name !== "tools");
    setHidden(els.extrasPanels.updates, name !== "updates");
    setHidden(els.extrasPanels.about, name !== "about");
    if (name === "updates" && !state.changelogLoaded) {
      loadChangelog();
    }
    if (name === "about" && !state.aboutLoaded) {
      renderAbout();
    }
  }

  function setSelection(sel) {
    state.selection = sel;
    if (sel && sel.type === "google") {
      highlightGoogleSelection(sel.id_or_path);
      updatePickerLabel("google", sel.label || sel.id_or_path);
    } else if (sel && sel.type === "excel") {
      highlightExcelSelection(sel.id_or_path);
      updatePickerLabel("excel", sel.label || sel.id_or_path);
    }
    clearSheetError();
    applyFieldStates();
  }

  // ---- Data loading ----

  function loadStatus() {
    return apiGet("/api/status").then(function (s) {
      state.statusData = s;
      state.sheetLoaded = !!s.sheet_loaded;
      return s;
    });
  }

  function loadDirs() {
    return apiGet("/api/dirs").then(function (d) {
      if (!d || !d.ok) return;
      els.inputDir.value = d.input || "";
      els.outputDir.value = d.output || "";
      renderFolderRecents(els.inputRecentsList, d.recent_inputs || [], els.inputDir, "input");
      renderFolderRecents(els.outputRecentsList, d.recent_outputs || [], els.outputDir, "output");
    });
  }

  function renderFolderRecents(container, items, targetInput, kind) {
    if (!container) return;
    container.innerHTML = "";
    if (!items.length) {
      var empty = el("div", "recent-pop__empty", "No recent folders");
      container.appendChild(empty);
      return;
    }
    items.forEach(function (path) {
      var btn = el("button", "recent-pop__item", path);
      btn.type = "button";
      btn.setAttribute("role", "menuitem");
      btn.addEventListener("click", function () {
        targetInput.value = path;
        hideRecents(kind);
      });
      container.appendChild(btn);
    });
  }

  function renderRailRecents(projects) {
    if (!els.railRecents) return;
    els.railRecents.innerHTML = "";
    if (!projects.length) {
      els.railRecents.appendChild(el("div", "rail-recent--empty", "No recent projects yet"));
      return;
    }
    var currentKey = currentSessionKey();
    projects.slice(0, 4).forEach(function (project) {
      var btn = el("button", "rail-recent");
      btn.type = "button";
      var key = projectKey(project);
      if (currentKey && key === currentKey) {
        btn.classList.add("is-current");
        btn.title = "Currently loaded";
      }
      var paths = el("span", "rail-recent__paths");
      paths.appendChild(el("span", "rail-recent__path mono", project.input || ""));
      if (project.output && project.output !== project.input) {
        paths.appendChild(el("span", "rail-recent__path rail-recent__path--secondary mono", project.output));
      }
      btn.appendChild(paths);
      var when = formatWhen(project.last_opened);
      if (when) btn.appendChild(el("span", "rail-recent__when", when));
      btn.addEventListener("click", function () {
        restoreProject(project);
      });
      els.railRecents.appendChild(btn);
    });
  }

  function projectKey(project) {
    if (!project) return "";
    var sheet = project.spreadsheet || null;
    var sheetKey = sheet ? (sheet.type + "|" + sheet.id_or_path) : "";
    return (project.input || "") + "::" + (project.output || "") + "::" + sheetKey;
  }

  function currentSessionKey() {
    var s = state.statusData || {};
    if (!s.input_dir && !s.output_dir) return "";
    var sheetKey = s.sheet_loaded && s.spreadsheet_type && s.spreadsheet_id_or_path
      ? s.spreadsheet_type + "|" + s.spreadsheet_id_or_path
      : "";
    return (s.input_dir || "") + "::" + (s.output_dir || "") + "::" + sheetKey;
  }

  function formatWhen(iso) {
    if (!iso) return "";
    var then = new Date(iso);
    if (isNaN(then.getTime())) return "";
    var diffMs = Date.now() - then.getTime();
    if (diffMs < 0) return "";
    var minutes = Math.floor(diffMs / 60000);
    if (minutes < 1) return "just now";
    if (minutes < 60) return minutes + "m ago";
    var hours = Math.floor(minutes / 60);
    if (hours < 24) return hours + "h ago";
    var days = Math.floor(hours / 24);
    if (days < 7) return days + "d ago";
    if (days < 30) return Math.floor(days / 7) + "w ago";
    var months = Math.floor(days / 30);
    if (months < 12) return months + "mo ago";
    return Math.floor(days / 365) + "y ago";
  }

  function restoreProject(project) {
    if (!project) return;
    els.inputDir.value = project.input || "";
    els.outputDir.value = project.output || "";
    var sheet = project.spreadsheet;
    if (sheet && sheet.type && sheet.id_or_path) {
      setTab(sheet.type);
      setSelection({
        type: sheet.type,
        id_or_path: sheet.id_or_path,
        label: sheet.label || sheet.id_or_path,
      });
      if (sheet.type === "google" && els.googlePaste) {
        els.googlePaste.value = sheet.id_or_path;
      } else if (sheet.type === "excel" && els.excelPaste) {
        els.excelPaste.value = sheet.id_or_path;
      }
    } else {
      setTab("none");
      state.selection = null;
    }
    // Restoring a project resets the baseline so the field glow reads as
    // "loaded" rather than "dirty" until the user edits again.
    state.baseline = baselineFromInputs();
    applyFieldStates();
    clearFieldError(els.inputField);
    clearFieldError(els.outputField);
  }

  function baselineFromInputs() {
    return {
      input: els.inputDir.value || "",
      output: els.outputDir.value || "",
      sheetTab: state.activeTab,
      sheetKey: state.selection
        ? state.selection.type + "|" + state.selection.id_or_path
        : "",
    };
  }

  function loadStartSettings() {
    return apiGet("/api/start-settings").then(function (r) {
      if (!r || !r.settings) return;
      state.persistEnabled = r.settings.persist_enabled !== false;
      if (els.persist) els.persist.checked = state.persistEnabled;
      state.recentProjects = r.settings.recent_projects || [];
      renderRailRecents(state.recentProjects);
    });
  }

  // ---- Field state glow (.is-loaded / .is-dirty / .is-error) ----

  function applyFieldStates() {
    setFieldState(els.inputField, fieldState("input", els.inputDir.value || ""));
    setFieldState(els.outputField, fieldState("output", els.outputDir.value || ""));
    setSheetCardState(sheetState());
  }

  function fieldState(kind, currentValue) {
    var node = kind === "input" ? els.inputField : els.outputField;
    if (node && node.classList.contains("has-error")) return "error";
    var baseline = kind === "input" ? state.baseline.input : state.baseline.output;
    if (!currentValue) return "";
    if (currentValue === baseline) return "loaded";
    return "dirty";
  }

  function sheetState() {
    var card = root && root.querySelector(".sheet-card");
    if (card && card.classList.contains("has-error")) return "error";
    var currentKey = state.selection
      ? state.selection.type + "|" + state.selection.id_or_path
      : "";
    var sameKey = currentKey === state.baseline.sheetKey;
    var sameTab = state.activeTab === state.baseline.sheetTab;
    if (sameKey && sameTab) {
      // Show "loaded" only when there is something concrete to mark.
      return currentKey ? "loaded" : "";
    }
    return "dirty";
  }

  function setFieldState(node, stateName) {
    if (!node) return;
    var input = node.querySelector(".path-input");
    if (!input) return;
    input.classList.remove("is-loaded", "is-dirty", "is-error");
    if (stateName) input.classList.add("is-" + stateName);
  }

  function setSheetCardState(stateName) {
    var card = root && root.querySelector(".sheet-card");
    if (!card) return;
    card.classList.remove("is-loaded", "is-dirty", "is-error");
    if (stateName) card.classList.add("is-" + stateName);
  }

  function markFieldError(node, message) {
    if (!node) return;
    node.classList.add("has-error");
    var input = node.querySelector(".path-input");
    if (input) {
      input.classList.remove("is-loaded", "is-dirty");
      input.classList.add("is-error");
    }
    if (message && typeof showToast === "function") showToast(message);
  }

  function clearFieldError(node) {
    if (!node) return;
    if (!node.classList.contains("has-error")) return;
    node.classList.remove("has-error");
    var input = node.querySelector(".path-input");
    if (input) input.classList.remove("is-error");
  }

  function markSheetError(message) {
    var card = root && root.querySelector(".sheet-card");
    if (!card) return;
    card.classList.add("has-error", "is-error");
    card.classList.remove("is-loaded", "is-dirty");
    if (message && typeof showToast === "function") showToast(message);
  }

  function clearSheetError() {
    var card = root && root.querySelector(".sheet-card");
    if (!card) return;
    card.classList.remove("has-error", "is-error");
  }

  function loadGoogleSheets() {
    if (!els.googleStatus) return Promise.resolve();
    els.googleStatus.textContent = "Checking authentication…";
    setHidden(els.googlePicker, true);
    return apiGet("/api/spreadsheets/google").then(function (g) {
      if (!g.authenticated) {
        renderGoogleConnectCTA(g.auth_in_flight, g.auth_error);
        return;
      }
      if (g.auth_error) {
        els.googleStatus.textContent = "Google: " + g.auth_error;
        return;
      }
      state.googleSheets = g.sheets || [];
      els.googleStatus.textContent = state.googleSheets.length
        ? state.googleSheets.length + " spreadsheets available"
        : "No spreadsheets found in your account";
      renderGoogleList(state.googleSheets);
      setHidden(els.googlePicker, false);
    });
  }

  function renderGoogleConnectCTA(inFlight, errorMsg) {
    els.googleStatus.innerHTML = "";
    setHidden(els.googlePicker, true);
    if (errorMsg) {
      els.googleStatus.appendChild(el("div", "sheet-panel__status-error", errorMsg));
    }
    var row = el("div", "sheet-panel__connect-row");
    if (!inFlight) {
      var btn = el("button", "sheet-panel__connect", "Connect Google");
      btn.type = "button";
      btn.addEventListener("click", connectGoogle);
      row.appendChild(btn);
    }
    var msg = el(
      "div",
      "sheet-panel__connect-msg",
      inFlight
        ? "Waiting for Google sign-in… (a browser tab should have opened)"
        : "Sign in to list your Google Sheets."
    );
    row.appendChild(msg);
    els.googleStatus.appendChild(row);
  }

  function connectGoogle() {
    els.googleStatus.textContent = "Starting Google sign-in…";
    apiPost("/api/spreadsheets/google/auth", {})
      .then(function () { pollGoogleAuth(); })
      .catch(function (err) {
        els.googleStatus.textContent = "Could not start sign-in: " + err.message;
      });
  }

  function stopGooglePoll() {
    if (state.googlePollTimer) {
      clearTimeout(state.googlePollTimer);
      state.googlePollTimer = null;
    }
  }

  function onGooglePollSuccess(sheets) {
    state.googleSheets = sheets || [];
    els.googleStatus.textContent = state.googleSheets.length
      ? state.googleSheets.length + " spreadsheets available"
      : "No spreadsheets found in your account";
    renderGoogleList(state.googleSheets);
    setHidden(els.googlePicker, false);
  }

  function pollGoogleAuth() {
    stopGooglePoll();
    state.googlePollDeadline = Date.now() + GOOGLE_POLL_TIMEOUT_MS;
    pollGoogleAuthOnce();
  }

  function pollGoogleAuthOnce() {
    apiGet("/api/spreadsheets/google").then(function (g) {
      if (g.authenticated && !g.auth_error) {
        state.googlePollTimer = null;
        onGooglePollSuccess(g.sheets);
        return;
      }
      if (g.auth_error) {
        state.googlePollTimer = null;
        renderGoogleConnectCTA(false, g.auth_error);
        return;
      }
      if (Date.now() > state.googlePollDeadline) {
        state.googlePollTimer = null;
        renderGoogleConnectCTA(false, "Sign-in timed out. Try again.");
        return;
      }
      state.googlePollTimer = setTimeout(pollGoogleAuthOnce, GOOGLE_POLL_INTERVAL_MS);
    }).catch(function () {
      state.googlePollTimer = setTimeout(pollGoogleAuthOnce, GOOGLE_POLL_RETRY_MS);
    });
  }

  function renderGoogleList(sheets) {
    if (!els.googlePickerMenu) return;
    els.googlePickerMenu.innerHTML = "";
    if (!sheets.length) {
      els.googlePickerMenu.appendChild(
        el("div", "sheet-picker__empty", "No spreadsheets in your account")
      );
      return;
    }
    sheets.forEach(function (sheet) {
      var option = el("button", "sheet-picker__option");
      option.type = "button";
      option.setAttribute("role", "option");
      option.setAttribute("data-id", sheet.name);
      option.appendChild(el("span", "sheet-picker__option-main", sheet.name));
      option.addEventListener("click", function () {
        if (els.googlePaste) els.googlePaste.value = "";
        setSelection({ type: "google", id_or_path: sheet.name, label: sheet.name });
        closePicker("google");
      });
      els.googlePickerMenu.appendChild(option);
    });
    if (state.selection && state.selection.type === "google") {
      highlightGoogleSelection(state.selection.id_or_path);
    }
  }

  function highlightGoogleSelection(idOrPath) {
    if (!els.googlePickerMenu) return;
    var items = els.googlePickerMenu.querySelectorAll(".sheet-picker__option");
    Array.prototype.forEach.call(items, function (item) {
      item.classList.toggle("is-selected", item.getAttribute("data-id") === idOrPath);
    });
  }

  function loadExcelFiles() {
    if (!els.excelStatus) return Promise.resolve();
    els.excelStatus.textContent = "Scanning input folder…";
    return apiGet("/api/spreadsheets/excel").then(function (r) {
      state.excelFiles = r.files || [];
      els.excelStatus.textContent = state.excelFiles.length
        ? state.excelFiles.length + " .xlsx in " + r.input_dir
        : "No .xlsx files in " + r.input_dir;
      renderExcelList(state.excelFiles);
    });
  }

  function renderExcelList(files) {
    if (!els.excelPickerMenu) return;
    els.excelPickerMenu.innerHTML = "";
    if (!files.length) {
      els.excelPickerMenu.appendChild(
        el("div", "sheet-picker__empty", "No .xlsx files in the input folder")
      );
      return;
    }
    files.forEach(function (f) {
      var option = el("button", "sheet-picker__option");
      option.type = "button";
      option.setAttribute("role", "option");
      option.setAttribute("data-path", f.path);
      option.appendChild(el("span", "sheet-picker__option-main", f.name));
      option.appendChild(el("span", "sheet-picker__option-sub", f.path));
      option.addEventListener("click", function () {
        if (els.excelPaste) els.excelPaste.value = "";
        setSelection({ type: "excel", id_or_path: f.path, label: f.name });
        closePicker("excel");
      });
      els.excelPickerMenu.appendChild(option);
    });
    if (state.selection && state.selection.type === "excel") {
      highlightExcelSelection(state.selection.id_or_path);
    }
  }

  function highlightExcelSelection(path) {
    if (!els.excelPickerMenu) return;
    var items = els.excelPickerMenu.querySelectorAll(".sheet-picker__option");
    Array.prototype.forEach.call(items, function (item) {
      item.classList.toggle("is-selected", item.getAttribute("data-path") === path);
    });
  }

  function loadChangelog() {
    return apiGet("/api/changelog").then(function (r) {
      state.changelogLoaded = true;
      var entries = (r && r.entries) || [];
      renderChangelog(entries);
      if (entries.length > 0 && els.updatesBadge) {
        els.updatesBadge.textContent = String(entries.length);
        setHidden(els.updatesBadge, false);
      }
    }).catch(function () {
      state.changelogLoaded = true;
      renderChangelog([]);
    });
  }

  function renderChangelog(entries) {
    if (!els.changelogList) return;
    els.changelogList.innerHTML = "";
    if (!entries.length) {
      var empty = el("div", "changelog__empty", "No changelog entries yet.");
      els.changelogList.appendChild(empty);
      return;
    }
    entries.forEach(function (entry) {
      var item = el("li", "changelog__entry");
      var head = el("div", "changelog__head");
      head.appendChild(el("span", "changelog__version", entry.version));
      var tag = el("span", "changelog__tag");
      tag.setAttribute("data-tool", entry.tool || "Core");
      tag.appendChild(el("span", "changelog__tag-dot"));
      tag.appendChild(document.createTextNode(entry.tool || "Core"));
      head.appendChild(tag);
      head.appendChild(el("span", "changelog__when", entry.date || ""));
      item.appendChild(head);
      if (entry.title) item.appendChild(el("div", "changelog__title", entry.title));
      if (entry.body) item.appendChild(el("div", "changelog__body", entry.body));
      els.changelogList.appendChild(item);
    });
  }

  function renderAbout() {
    if (!els.aboutGrid) return;
    state.aboutLoaded = true;
    var s = state.statusData || {};
    els.aboutGrid.innerHTML = "";

    aboutRow("Project", function (val) {
      val.appendChild(el("span", "about__title", "clipgen"));
      val.appendChild(el("span", "about__sub", "Generate clips, screenshots, and transcripts from playtest videos."));
    });
    aboutRow("Version", function (val) {
      val.classList.add("mono");
      val.textContent = "v" + (s.version || "0.0.0");
    });
    aboutRow("Author", function (val) {
      val.textContent = s.author || "Henrik Edlund";
    });
    aboutRow("Repository", function (val) {
      var repo = s.repo_url || "https://github.com/henedl/clipgen";
      var label = repo.replace(/^https?:\/\//, "");
      var link = document.createElement("a");
      link.className = "about__link mono";
      link.href = repo;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.textContent = label;
      var icon = document.createElement("span");
      icon.className = "so-icon so-icon--xs";
      icon.setAttribute("data-icon", "arrow-up-right");
      applyIconMask(icon, "arrow-up-right");
      link.appendChild(icon);
      val.appendChild(link);
    });
    aboutRow("License", function (val) {
      val.appendChild(el("span", "about__pill", s.license || "MIT"));
      val.appendChild(el("span", "about__sub", "© 2017–2026 " + (s.author || "Henrik Edlund")));
    });
  }

  function aboutRow(label, build) {
    var row = el("div", "about__row");
    row.appendChild(el("div", "about__label", label));
    var val = el("div", "about__value");
    build(val);
    row.appendChild(val);
    els.aboutGrid.appendChild(row);
  }

  // ---- Open / dismiss flows ----

  function confirm() {
    // Guard against re-entry: the click handler disables the button, but the
    // Cmd/Ctrl+Enter shortcut bypasses that path. Two simultaneous flights
    // would race the server-side _swap_worksheet.
    if (state.confirmInFlight) return;
    state.confirmInFlight = true;

    var inputVal = (els.inputDir.value || "").trim();
    var outputVal = (els.outputDir.value || "").trim();

    clearFieldError(els.inputField);
    clearFieldError(els.outputField);
    clearSheetError();

    var dirsPayload = {};
    if (inputVal) dirsPayload.input = inputVal;
    if (outputVal) dirsPayload.output = outputVal;

    var dirsPromise = Object.keys(dirsPayload).length
      ? fetch("/api/dirs", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(dirsPayload),
        }).then(function (r) {
          return r.json().then(function (j) { return { ok: r.ok, body: j }; });
        })
      : Promise.resolve({ ok: true, body: {} });

    function releaseConfirm() {
      state.confirmInFlight = false;
      els.confirmBtn.disabled = false;
    }

    els.confirmBtn.disabled = true;
    dirsPromise.then(function (res) {
      if (!res.ok) {
        var errors = (res.body && res.body.errors) || {};
        if (errors.input) markFieldError(els.inputField, errors.input);
        if (errors.output) markFieldError(els.outputField, errors.output);
        if (!errors.input && !errors.output && typeof showToast === "function") {
          showToast("Folder error");
        }
        releaseConfirm();
        return;
      }
      var skipSpreadsheet = state.activeTab === "none" || !state.selection;
      if (skipSpreadsheet) {
        recordSession(inputVal, outputVal, null).finally(function () {
          releaseConfirm();
          close();
        });
        return;
      }
      fetch("/api/spreadsheets/open", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(state.selection),
      })
        .then(function (r) {
          return r.json().then(function (j) { return { ok: r.ok, body: j }; });
        })
        .then(function (res2) {
          if (!res2.ok || !res2.body.ok) {
            releaseConfirm();
            markSheetError((res2.body && res2.body.error) || "Could not open spreadsheet");
            return;
          }
          // Hard reload so all three frontends re-fetch their data.
          // Don't release the in-flight flag — the page is about to unmount.
          window.location.reload();
        })
        .catch(function (err) {
          releaseConfirm();
          markSheetError("Open failed: " + err.message);
        });
    }).catch(function (err) {
      releaseConfirm();
      console.error("Confirm dirs failed", err);
    });
  }

  function recordSession(input, output, spreadsheet) {
    // Fire-and-forget on the no-sheet path. The sheet-open path records on
    // the server during /api/spreadsheets/open, so this is the only call
    // site that needs to invoke /api/sessions/record explicitly.
    var payload = { input: input, output: output };
    if (spreadsheet) payload.spreadsheet = spreadsheet;
    return apiPost("/api/sessions/record", payload).catch(function (err) {
      console.warn("Session record failed", err);
    });
  }

  // ---- Open / close ----

  function open() {
    if (!state.mounted) {
      mount().then(open);
      return;
    }
    if (state.open) return;
    state.open = true;
    show(root, true);
    // Register as the shared blocking modal: Escape closes, Tab is trapped
    // inside, and hotkeys.js scopes Alt-hold hints to this overlay's controls.
    // initialFocus is the panel (tabindex="-1") so focus lands on a non-typing
    // element — the reveal-hints and letter shortcuts work immediately, and the
    // trap pulls Tab back into the cycle from there. restoreFocus returns focus
    // to the launch trigger on dismiss (else it stays stuck on the hidden panel).
    if (typeof openBlockingModal === "function") {
      openBlockingModal(root, {
        onEscape: close,
        trapFocus: true,
        initialFocus: els.panel,
        restoreFocus: true
      });
    }
    runIntro();
    refresh();
  }

  function close() {
    if (!root) return;
    state.open = false;
    markDismissed();
    stopGooglePoll();
    // Animate the panel and backdrop out, then hide.
    if (els.panel) els.panel.classList.remove("is-in");
    root.style.setProperty("--host-blur", "0px");
    root.style.setProperty("--veil-alpha", "0");
    setTimeout(function () {
      // Guard against a re-open during the fade (open() flips state.open back).
      if (state.open) return;
      show(root, false);
      // Release the blocking modal only once the overlay is actually hidden.
      // The start overlay never sets body.modal-open, so blockingModalOpen()
      // tracks this trap; releasing it early would re-activate page hotkeys and
      // background Alt-hint chips while the launcher is still fading out.
      if (typeof closeBlockingModal === "function") closeBlockingModal(root);
    }, 460);
  }

  // Replay the rail brand-mark stroke-draw on every overlay open. The shared
  // per-session gate in clipgenInitBrandMark (utils.js) draws only the first
  // .brand-mark hydrated on the page — usually the topnav logo — so the launcher
  // forces its own replay here, independent of that flag. Retries across a few
  // animation frames because hydration (an async favicon.svg fetch) may not have
  // landed yet on the first open; it no-ops once attempts run out.
  function replayBrandMark(attempts) {
    var mark = els.mark;
    if (!mark) return;
    if (mark.classList.contains("is-hydrated")) {
      mark.classList.remove("is-animated");
      void mark.offsetWidth; // force reflow so the draw-on restarts
      mark.classList.add("is-animated");
      return;
    }
    if (attempts > 0) {
      requestAnimationFrame(function () { replayBrandMark(attempts - 1); });
    }
  }

  function runIntro() {
    if (!root) return;
    // Reset cascade-in elements so the animation can replay.
    Array.prototype.forEach.call(els.cascades, function (node) {
      node.classList.remove("is-in");
    });
    if (els.panel) els.panel.classList.remove("is-in");

    // Backdrop builds in 100ms after mount so the panel is in place first.
    root.style.setProperty("--host-blur", "0px");
    root.style.setProperty("--veil-alpha", "0");
    setTimeout(function () {
      root.style.setProperty("--host-blur", INTRO_BLUR_PX + "px");
      root.style.setProperty("--veil-alpha", String(INTRO_VEIL_ALPHA));
    }, 100);

    // Panel slide-in.
    requestAnimationFrame(function () {
      if (els.panel) els.panel.classList.add("is-in");
    });

    // Replay the wordmark by re-triggering the keyframe.
    if (els.wordmark) {
      els.wordmark.style.animation = "none";
      // Force reflow so the next assignment restarts the animation.
      void els.wordmark.offsetWidth;
      els.wordmark.style.animation = "";
    }

    // Replay the brand-mark stroke-draw alongside the wordmark.
    replayBrandMark(30);

    // Section cascade.
    Array.prototype.forEach.call(els.cascades, function (node) {
      var idx = parseInt(node.getAttribute("data-cascade") || "0", 10);
      setTimeout(function () { node.classList.add("is-in"); }, CASCADE_BASE_MS + idx * CASCADE_STEP_MS);
    });
  }

  function refresh() {
    loadStatus()
      .then(loadDirs)
      .then(loadStartSettings)
      .then(loadGoogleSheets)
      .then(loadExcelFiles)
      .then(applyCurrentSessionPrefill)
      .catch(function (err) {
        console.error("Start overlay refresh failed", err);
      });
  }

  function applyCurrentSessionPrefill() {
    var s = state.statusData || {};
    if (s.sheet_loaded && s.spreadsheet_type && s.spreadsheet_id_or_path) {
      // Seed the picker so the current session shows as "loaded" without the
      // user needing to re-select. Activate the correct tab + selection.
      setTab(s.spreadsheet_type);
      setSelection({
        type: s.spreadsheet_type,
        id_or_path: s.spreadsheet_id_or_path,
        label: s.spreadsheet_label || s.spreadsheet_id_or_path,
      });
      if (s.spreadsheet_type === "google" && els.googlePaste) {
        // Match against the open Google sheet list if possible; otherwise show
        // the identifier in the paste field so the user can confirm the value.
        var matches = (state.googleSheets || []).some(function (sheet) {
          return sheet.name === s.spreadsheet_id_or_path;
        });
        if (!matches) els.googlePaste.value = s.spreadsheet_id_or_path;
      } else if (s.spreadsheet_type === "excel" && els.excelPaste) {
        var inList = (state.excelFiles || []).some(function (f) {
          return f.path === s.spreadsheet_id_or_path;
        });
        if (!inList) els.excelPaste.value = s.spreadsheet_id_or_path;
      }
    } else if (s.sheet_loaded) {
      // Sheet loaded via CLI without matching meta — leave the picker on its
      // default tab but show a "current" baseline as no-sheet so the user can
      // change things without the dirty glow misfiring.
      state.baseline.sheetTab = state.activeTab;
      state.baseline.sheetKey = "";
    } else {
      // No sheet loaded: the session's current state is "no spreadsheet".
      setTab("none");
    }
    state.baseline = baselineFromInputs();
    applyFieldStates();
    // Re-render rail recents now that we know the current session key.
    renderRailRecents(state.recentProjects);
  }

  // ---- Boot ----

  function boot() {
    mount().then(function () {
      apiGet("/api/status").then(function (s) {
        state.statusData = s;
        state.sheetLoaded = !!s.sheet_loaded;
        if (shouldAutoOpen(s)) open();
      }).catch(function () { /* offline / dev */ });
    });

    if (window.ClipgenTopNav && typeof window.ClipgenTopNav.onReady === "function") {
      window.ClipgenTopNav.onReady(function () {
        var btn = document.getElementById("startBtn");
        if (btn) {
          btn.addEventListener("click", function () { open(); });
        }
      });
    }
  }

  window.ClipgenStartOverlay = {
    open: open,
    close: close,
    isOpen: function () { return state.open; },
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
