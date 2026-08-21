/* Start overlay — Direction B redesign.
 *
 * Two-column launcher mounted by Studio / Screenspace / Transcripts. Drives:
 *   • brand-mark + wordmark intro (reuses window.clipgenInitBrandMark; cascade
 *     plays once per browser session, gated by the existing sessionStorage flag)
 *   • section cascade-in (220ms base, 80ms stagger)
 *   • backdrop blur via the shared .cg-veil / .cg-veil-layer (tokens.css)
 *   • optional project name, stored on the recent-projects record and shown as
 *     that entry's title (there is no server-side "current project name")
 *   • the rail's Recently-opened list: RAIL_RECENTS_VISIBLE rows plus a
 *     fold-out that overlays the brand block behind a blurred scrim
 *   • the right column's three top-level tabs: Open (name + folders +
 *     spreadsheet + the confirm footer), About (tool tiles + about rows),
 *     Recent updates (changelog)
 *   • folder + spreadsheet picker (Google / Excel / No spreadsheet)
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
  var CASCADE_BASE_MS = 220;
  var CASCADE_STEP_MS = 80;
  // Google-auth poll lifecycle: the server's /api/spreadsheets/google/auth
  // launches an OAuth flow in a daemon thread; we poll /api/spreadsheets/google
  // until it reports `authenticated`, an `auth_error`, or we exceed the budget.
  var GOOGLE_POLL_TIMEOUT_MS = 90 * 1000;
  var GOOGLE_POLL_INTERVAL_MS = 1500;
  var GOOGLE_POLL_RETRY_MS = 2500;
  // Wait for typing to settle before fetching worksheets for a pasted URL/name.
  var WORKSHEET_PASTE_DEBOUNCE_MS = 600;
  // Recents shown in the rail before the fold-out; the rest overlay the brand
  // block on demand. The server keeps start_settings.RECENTS_CAP of them.
  var RAIL_RECENTS_VISIBLE = 3;

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
    rememberWindow: true,
    googlePollTimer: null,
    googlePollDeadline: 0,
    confirmInFlight: false,
    activeTab: "google",    // 'google' | 'excel' | 'mindnode' | 'none'
    startTab: "open",       // right column: 'open' | 'about' | 'updates'
    changelogLoaded: false,
    // The notice file never changes at runtime, so the fetch is latched. The
    // rows still live in state because renderAttribution() re-runs on every
    // About activation alongside renderAbout().
    licensesLoaded: false,
    licenseComponents: [],
    googleSheets: [],
    // Last unauthenticated /api/spreadsheets/google payload: the credentials
    // filename, searched paths and setup link behind the "Don't have
    // credentials.json?" disclosure. Cached because the auth poll re-renders
    // the CTA without one, and these facts don't change mid-sign-in.
    googleCreds: null,
    excelFiles: [],
    mindnodeFiles: [],
    mindnodePreviewReqVer: 0,   // rejects stale mind-map summary fetches
    statusData: null,
    recentProjects: [],
    projectName: "",          // the optional label typed in the right column
    // The name field has no server-side source of truth — it is filled in from
    // the matching recent-projects entry at the tail of refresh(). These two
    // flags say whether that value is trustworthy yet, so a fast Cmd+Enter
    // can't post an empty name that clears a stored label, and a late-landing
    // prefill can't stomp what the user was typing meanwhile.
    projectNamePrefilled: false,  // applyCurrentSessionPrefill has run at least once
    projectNameAuthored: false,   // the user typed it, or picked a recent project
    recentsExpanded: false,   // rail fold-out revealing projects past RAIL_RECENTS_VISIBLE
    worksheetsCache: {},      // "type|id_or_path" -> { worksheets, recommended }
    worksheetReqVer: 0,       // rejects stale worksheet fetches
    worksheetLoading: false,  // a worksheet fetch is in flight (gates Confirm)
    wsPasteTimer: null,       // debounce for paste-driven worksheet loads
    // Source-video filename preview, keyed "type|id_or_path|worksheet|inputDir"
    // — the input dir is part of the key because it decides found/missing.
    previewCache: {},
    previewReqVer: 0,         // rejects stale preview fetches
    previewDirTimer: null,    // debounce for input-folder-driven refreshes
    // Set while a preview row's filename is being edited inline. The overlay's
    // Escape is owned by a capture-phase modal trap that no listener on the
    // input can pre-empt, so onEscape consults this first.
    cancelPreviewEdit: null,
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

  // Spin a panel's Refresh button while its reload runs. Shared by the Google
  // and Excel buttons; both loaders recover from their own failures, so this
  // only owns the button state.
  function runPanelRefresh(btn, load) {
    if (!btn || btn.disabled) return;
    btn.disabled = true;
    btn.classList.add("is-spinning");
    load()
      .catch(function (err) { console.error("Panel refresh failed", err); })
      .then(function () {
        // Final link, after .catch(), so the spinner always stops.
        btn.disabled = false;
        btn.classList.remove("is-spinning");
      });
  }

  // Sweeping "still working" status text. The carrier gets its own hugging span:
  // .sheet-panel__status is a stretched flex column that also hosts CTA rows, and
  // .cg-shimmer's transparent text fill inherits onto anything inside it. Every
  // result/error write below is a plain textContent, which replaces the span and
  // ends the sweep on its own.
  function setStatusShimmer(node, text) {
    if (!node) return;
    node.innerHTML = "";
    node.appendChild(el("span", "cg-shimmer", text));
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

    els.railRecentsSection = root.querySelector('[data-role="rail-recents-section"]');
    els.railRecents = root.querySelector('[data-role="rail-recents"]');
    els.railRecentsOverflow = root.querySelector('[data-role="rail-recents-overflow"]');
    els.railRecentsMore = root.querySelector('[data-role="rail-recents-more"]');
    els.railRecentsMoreLabel = root.querySelector('[data-role="rail-recents-more-label"]');
    els.railRecentsScrim = root.querySelector('[data-role="rail-recents-scrim"]');

    els.projectName = root.querySelector("#startProjectName");
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
    els.googleRefresh = root.querySelector('[data-role="google-refresh"]');
    els.googlePicker = root.querySelector('[data-role="google-picker"]');
    els.googlePickerTrigger = root.querySelector('[data-role="google-picker-trigger"]');
    els.googlePickerLabel = root.querySelector('[data-role="google-picker-label"]');
    els.googlePickerMenu = root.querySelector('[data-role="google-picker-menu"]');
    els.googlePaste = root.querySelector("#startGooglePaste");
    els.excelStatus = root.querySelector('[data-role="excel-status"]');
    els.excelRefresh = root.querySelector('[data-role="excel-refresh"]');
    els.excelPicker = root.querySelector('[data-role="excel-picker"]');
    els.excelPickerTrigger = root.querySelector('[data-role="excel-picker-trigger"]');
    els.excelPickerLabel = root.querySelector('[data-role="excel-picker-label"]');
    els.excelPickerMenu = root.querySelector('[data-role="excel-picker-menu"]');
    els.excelPaste = root.querySelector("#startExcelPaste");
    els.mindnodePanel = root.querySelector('[data-tabpanel="mindnode"]');
    els.mindnodeStatus = root.querySelector('[data-role="mindnode-status"]');
    els.mindnodeRefresh = root.querySelector('[data-role="mindnode-refresh"]');
    els.mindnodePicker = root.querySelector('[data-role="mindnode-picker"]');
    els.mindnodePickerTrigger = root.querySelector('[data-role="mindnode-picker-trigger"]');
    els.mindnodePickerLabel = root.querySelector('[data-role="mindnode-picker-label"]');
    els.mindnodePickerMenu = root.querySelector('[data-role="mindnode-picker-menu"]');
    els.mindnodePaste = root.querySelector("#startMindnodePaste");
    els.mindnodePreview = root.querySelector('[data-role="mindnode-preview"]');
    els.mindnodePreviewThumb = root.querySelector('[data-role="mindnode-preview-thumb"]');
    els.mindnodePreviewSummary = root.querySelector('[data-role="mindnode-preview-summary"]');
    els.mindnodeSources = root.querySelector('[data-role="mindnode-sources"]');
    els.mindnodeSourcesSummary = root.querySelector('[data-role="mindnode-sources-summary"]');
    els.mindnodeSourcesList = root.querySelector('[data-role="mindnode-sources-list"]');
    els.worksheetSection = root.querySelector('[data-role="worksheet-section"]');
    els.worksheetLoading = root.querySelector('[data-role="worksheet-loading"]');
    els.worksheetPicker = root.querySelector('[data-role="worksheet-picker"]');
    els.worksheetPickerTrigger = root.querySelector('[data-role="worksheet-picker-trigger"]');
    els.worksheetPickerLabel = root.querySelector('[data-role="worksheet-picker-label"]');
    els.worksheetPickerMenu = root.querySelector('[data-role="worksheet-picker-menu"]');
    els.sourcePreview = root.querySelector('[data-role="source-preview"]');
    els.sourcePreviewSpinner = root.querySelector('[data-role="source-preview-spinner"]');
    els.sourcePreviewIcon = root.querySelector('[data-role="source-preview-icon"]');
    els.sourcePreviewSummary = root.querySelector('[data-role="source-preview-summary"]');
    els.sourcePreviewList = root.querySelector('[data-role="source-preview-list"]');
    els.sourceFiles = root.querySelector("#startSourceFiles");

    els.startTabs = root.querySelectorAll(".start-tab");
    els.startPanels = {
      open: root.querySelector('[data-start-panel="open"]'),
      about: root.querySelector('[data-start-panel="about"]'),
      updates: root.querySelector('[data-start-panel="updates"]'),
    };
    els.updatesBadge = root.querySelector('[data-role="updates-badge"]');
    els.changelogList = root.querySelector('[data-role="changelog-list"]');
    els.aboutGrid = root.querySelector('[data-role="about-grid"]');
    els.attributionList = root.querySelector('[data-role="attribution-list"]');

    els.persist = root.querySelector("#startPersistEnabled");
    els.rememberWindow = root.querySelector("#startRememberWindow");
    els.rememberWindowRow = root.querySelector("#startRememberWindowRow");
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

    // Right column's top-level tab strip
    Array.prototype.forEach.call(els.startTabs, function (tab) {
      on(tab, "click", function () { setStartTab(tab.getAttribute("data-start-tab")); });
    });

    on(els.projectName, "input", function () {
      state.projectNameAuthored = true;
      setProjectName(els.projectName.value || "");
    });

    on(els.railRecentsMore, "click", function (e) {
      e.stopPropagation();
      setRecentsExpanded(!state.recentsExpanded);
    });
    on(els.railRecentsScrim, "click", function () {
      setRecentsExpanded(false);
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
      // The folder decides which expected videos count as found, so re-resolve
      // the preview once typing settles (cached per folder, so re-typing a
      // folder already seen costs no request).
      if (state.previewDirTimer) clearTimeout(state.previewDirTimer);
      state.previewDirTimer = setTimeout(function () {
        state.previewDirTimer = null;
        if (!state.selection) return;
        if (state.selection.type === "mindnode") loadMindnodePreview(state.selection);
        else loadSourcePreview();
      }, WORKSHEET_PASTE_DEBOUNCE_MS);
    });
    on(els.outputDir, "input", function () {
      clearFieldError(els.outputField);
      applyFieldStates();
    });

    on(els.googlePickerTrigger, "click", function (e) {
      e.stopPropagation();
      togglePicker("google");
    });
    on(els.googleRefresh, "click", function () {
      // The server caches Drive's listing for 5 minutes; this is the escape
      // hatch for a spreadsheet created mid-session.
      runPanelRefresh(els.googleRefresh, function () {
        return loadGoogleSheets(true);
      });
    });
    on(els.excelRefresh, "click", function () {
      runPanelRefresh(els.excelRefresh, loadExcelFiles);
    });
    on(els.excelPickerTrigger, "click", function (e) {
      e.stopPropagation();
      togglePicker("excel");
    });
    on(els.mindnodeRefresh, "click", function () {
      runPanelRefresh(els.mindnodeRefresh, loadMindnodeFiles);
    });
    on(els.mindnodePickerTrigger, "click", function (e) {
      e.stopPropagation();
      togglePicker("mindnode");
    });
    on(els.worksheetPickerTrigger, "click", function (e) {
      e.stopPropagation();
      togglePicker("worksheet");
    });

    on(document, "click", function (e) {
      // Bail when the overlay isn't open — the listener stays bound for the
      // page's lifetime (mount() is once-per-page; close() just hides),
      // so guard explicitly to avoid pointless work on every page click.
      if (!root || !state.open) return;
      closeRecentsIfOutside(e);
      closePickersIfOutside(e);
      if (state.recentsExpanded && els.railRecentsSection &&
          !els.railRecentsSection.contains(e.target)) {
        setRecentsExpanded(false);
      }
    });

    on(els.persist, "change", function () {
      state.persistEnabled = !!els.persist.checked;
      apiPost("/api/start-settings", { persist_enabled: state.persistEnabled })
        .catch(function (err) { console.error("Persist toggle failed", err); });
    });

    on(els.rememberWindow, "change", function () {
      state.rememberWindow = !!els.rememberWindow.checked;
      apiPost("/api/start-settings", { remember_window: state.rememberWindow })
        .catch(function (err) { console.error("Window toggle failed", err); });
    });

    on(els.googlePaste, "input", function () {
      var v = (els.googlePaste.value || "").trim();
      if (v) {
        var sel = { type: "google", id_or_path: v, label: parseLabelFromGoogleInput(v) };
        setSelection(sel);
        scheduleWorksheetLoad(sel);
      } else if (state.selection && state.selection.type === "google") {
        renderGoogleList(state.googleSheets || []);
        hideWorksheetSection();
      }
    });

    on(els.excelPaste, "input", function () {
      var v = (els.excelPaste.value || "").trim();
      if (v) {
        var sel = { type: "excel", id_or_path: v, label: v.split("/").pop() || v };
        setSelection(sel);
        scheduleWorksheetLoad(sel);
      } else if (state.selection && state.selection.type === "excel") {
        renderExcelList(state.excelFiles || []);
        hideWorksheetSection();
      }
    });

    on(els.mindnodePaste, "input", function () {
      var v = (els.mindnodePaste.value || "").trim();
      if (v) {
        // A mind map has no worksheets, so there is nothing to schedule —
        // setSelection alone settles the selection and Confirm is free.
        setSelection({ type: "mindnode", id_or_path: v, label: v.split("/").pop() || v });
      } else if (state.selection && state.selection.type === "mindnode") {
        renderMindnodeList(state.mindnodeFiles || []);
      }
    });

    // Keyboard: Escape (close) and Tab containment are owned by the shared
    // blocking-modal trap (see open()); the rest are shared-registry hotkeys so
    // holding Alt reveals them and they list in the "?" cheatsheet. Gated by
    // inModal + `when` so they fire only while the launcher owns the keyboard.
    if (window.ClipgenHotkeys) {
      var isOpen = function () { return state.open; };
      // Spreadsheet/folder/confirm shortcuts target the Open pane. Gating only
      // on state.open left G/E/I/Cmd+Enter live on About and Recent updates,
      // where they mutated a hidden form or confirmed the workspace.
      var isOpenForm = function () { return state.open && state.startTab === "open"; };
      ClipgenHotkeys.register([
        { id: "start.tabOpen",      inModal: true, when: isOpen, handler: function () { setStartTab("open"); } },
        { id: "start.tabAbout",     inModal: true, when: isOpen, handler: function () { setStartTab("about"); } },
        { id: "start.tabUpdates",   inModal: true, when: isOpen, handler: function () { setStartTab("updates"); } },
        { id: "start.tabGoogle",    inModal: true, when: isOpenForm, handler: function () { setTab("google"); } },
        { id: "start.tabExcel",     inModal: true, when: isOpenForm, handler: function () { setTab("excel"); } },
        { id: "start.tabMindnode",  inModal: true, when: isOpenForm, handler: function () { setTab("mindnode"); } },
        { id: "start.tabNone",      inModal: true, when: isOpenForm, handler: function () { setTab("none"); } },
        { id: "start.browseInput",  inModal: true, when: isOpenForm, handler: function () { browseFolder("input"); } },
        { id: "start.browseOutput", inModal: true, when: isOpenForm, handler: function () { browseFolder("output"); } },
        { id: "start.confirm",      inModal: true, allowInInput: true, when: isOpenForm, handler: function () { confirm(); } }
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

  // ---- Sheet picker (Google + Excel + MindNode dropdowns) ----

  var PICKER_KINDS = ["google", "excel", "mindnode", "worksheet"];

  function pickerRefs(kind) {
    if (kind === "google") {
      return {
        menu: els.googlePickerMenu,
        trigger: els.googlePickerTrigger,
        label: els.googlePickerLabel,
        placeholder: "Select a Google Sheet…",
      };
    }
    if (kind === "excel") {
      return {
        menu: els.excelPickerMenu,
        trigger: els.excelPickerTrigger,
        label: els.excelPickerLabel,
        placeholder: "Select an Excel file…",
      };
    }
    if (kind === "mindnode") {
      return {
        menu: els.mindnodePickerMenu,
        trigger: els.mindnodePickerTrigger,
        label: els.mindnodePickerLabel,
        placeholder: "Select a MindNode document…",
      };
    }
    return {
      menu: els.worksheetPickerMenu,
      trigger: els.worksheetPickerTrigger,
      label: els.worksheetPickerLabel,
      placeholder: "Select a worksheet…",
    };
  }

  function togglePicker(kind) {
    var refs = pickerRefs(kind);
    if (!refs.menu) return;
    if (refs.menu.classList.contains("hidden")) {
      // Close the sibling pickers before opening this one.
      PICKER_KINDS.forEach(function (k) { if (k !== kind) closePicker(k); });
      openPicker(kind);
    } else {
      closePicker(kind);
    }
  }

  function openPicker(kind) {
    var refs = pickerRefs(kind);
    if (!refs.menu || !refs.trigger) return;
    refs.menu.classList.remove("hidden");
    refs.trigger.setAttribute("aria-expanded", "true");
    liftSectionForPicker(refs.trigger, true);
  }

  function closePicker(kind) {
    var refs = pickerRefs(kind);
    if (!refs.menu || !refs.trigger) return;
    refs.menu.classList.add("hidden");
    refs.trigger.setAttribute("aria-expanded", "false");
    liftSectionForPicker(refs.trigger, false);
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
    if (els.mindnodePicker && !els.mindnodePicker.contains(e.target) &&
        els.mindnodePickerMenu && !els.mindnodePickerMenu.classList.contains("hidden")) {
      closePicker("mindnode");
    }
    if (els.worksheetPicker && !els.worksheetPicker.contains(e.target) &&
        els.worksheetPickerMenu && !els.worksheetPickerMenu.classList.contains("hidden")) {
      closePicker("worksheet");
    }
  }

  function updatePickerLabel(kind, label) {
    var refs = pickerRefs(kind);
    if (!refs.label || !refs.trigger) return;
    if (label) {
      refs.label.textContent = label;
      refs.trigger.classList.remove("is-placeholder");
    } else {
      refs.label.textContent = refs.placeholder;
      refs.trigger.classList.add("is-placeholder");
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
    setHidden(els.mindnodePanel, name !== "mindnode");
    setHidden(els.nonePanel, name !== "none");
    if (name === "none") {
      // Clear any sheet selection — user is opting out.
      state.selection = null;
      hideWorksheetSection();
    } else if (state.selection && state.selection.type === name) {
      // Re-show the worksheet dropdown for a selection already on this tab.
      loadWorksheets(state.selection);
    } else {
      hideWorksheetSection();
    }
    clearSheetError();
    applyFieldStates();
  }

  function setStartTab(name) {
    // Only a real switch animates: re-selecting the active tab shouldn't
    // replay, and the reset to "open" on every overlay open would otherwise
    // fight the section cascade in runIntro().
    var changed = state.startTab !== name;
    state.startTab = name;
    Array.prototype.forEach.call(els.startTabs, function (tab) {
      var active = tab.getAttribute("data-start-tab") === name;
      tab.classList.toggle("is-active", active);
      tab.setAttribute("aria-selected", active ? "true" : "false");
    });
    setHidden(els.startPanels.open, name !== "open");
    setHidden(els.startPanels.about, name !== "about");
    setHidden(els.startPanels.updates, name !== "updates");
    if (changed) playTabEnter(els.startPanels[name]);
    if (name === "updates" && !state.changelogLoaded) {
      loadChangelog();
    }
    // Re-rendered on every activation rather than latched: it reads
    // state.statusData, which /api/status may not have filled in yet the first
    // time the tab is opened — a latch would freeze the panel on v0.0.0.
    if (name === "about") {
      renderAbout();
      renderAttribution();
      if (!state.licensesLoaded) loadLicenses();
    }
  }

  // One-shot enter animation on the panel being revealed. The class has to come
  // off again or the animation won't replay on the next switch.
  function playTabEnter(panel) {
    if (!panel) return;
    var scroll = panel.querySelector(".start-tabpanel__scroll");
    if (!scroll) return;
    panel.classList.remove("is-entering");
    void panel.offsetWidth; // force reflow so the animation restarts
    panel.classList.add("is-entering");
    var done = function () {
      panel.classList.remove("is-entering");
      scroll.removeEventListener("animationend", done);
    };
    scroll.addEventListener("animationend", done);
  }

  function setSelection(sel) {
    state.selection = sel;
    if (sel && sel.type === "google") {
      highlightGoogleSelection(sel.id_or_path);
      updatePickerLabel("google", sel.label || sel.id_or_path);
    } else if (sel && sel.type === "excel") {
      highlightExcelSelection(sel.id_or_path);
      updatePickerLabel("excel", sel.label || sel.id_or_path);
    } else if (sel && sel.type === "mindnode") {
      highlightMindnodeSelection(sel.id_or_path);
      updatePickerLabel("mindnode", sel.label || sel.id_or_path);
    }
    loadMindnodePreview(sel);
    // A fresh spreadsheet identity resets the worksheet choice; bump the
    // request version so any in-flight worksheet fetch for the prior selection
    // is ignored, then reset the dropdown (loadWorksheets re-shows it for 2+
    // tabs).
    state.worksheetReqVer++;
    hideWorksheetSection();
    clearSheetError();
    applyFieldStates();
  }

  // Pick a spreadsheet from a dropdown (or recent restore) and populate its
  // worksheet dropdown.
  function selectSpreadsheet(sel) {
    setSelection(sel);
    loadWorksheets(sel);
  }

  // ---- Worksheet dropdown ----

  function scheduleWorksheetLoad(sel) {
    if (state.wsPasteTimer) clearTimeout(state.wsPasteTimer);
    state.wsPasteTimer = setTimeout(function () {
      state.wsPasteTimer = null;
      // Only load if this is still the active selection.
      if (state.selection && state.selection.type === sel.type &&
          state.selection.id_or_path === sel.id_or_path) {
        loadWorksheets(state.selection);
      }
    }, WORKSHEET_PASTE_DEBOUNCE_MS);
  }

  function loadWorksheets(sel) {
    if (!sel || (sel.type !== "google" && sel.type !== "excel") || !sel.id_or_path) {
      hideWorksheetSection();
      return;
    }
    var key = sel.type + "|" + sel.id_or_path;
    var reqVer = ++state.worksheetReqVer;
    var cached = state.worksheetsCache[key];
    if (cached) {
      applyWorksheets(sel, cached, reqVer);
      return;
    }
    // Fetching a spreadsheet's tabs can be slow (Google opens it server-side).
    // Show a spinner and hold the Confirm button until we know the tab list, so
    // the user can't Open before choosing (or seeing) a worksheet.
    showWorksheetLoading();
    apiGet("/api/spreadsheets/worksheets?type=" + encodeURIComponent(sel.type) +
           "&id_or_path=" + encodeURIComponent(sel.id_or_path))
      .then(function (r) {
        var data = {
          worksheets: (r && r.worksheets) || [],
          recommended: (r && r.recommended) || "",
        };
        state.worksheetsCache[key] = data;
        applyWorksheets(sel, data, reqVer);
      })
      .catch(function () {
        // On failure just hide the section; the server auto-picks on open.
        if (reqVer === state.worksheetReqVer) hideWorksheetSection();
      });
  }

  function applyWorksheets(sel, data, reqVer) {
    if (reqVer !== state.worksheetReqVer) return;               // stale fetch
    if (!state.selection || state.selection.type !== sel.type ||
        state.selection.id_or_path !== sel.id_or_path) return;   // selection moved
    if (state.activeTab !== sel.type) return;                    // tab switched away
    var titles = data.worksheets || [];
    var preferred = data.recommended || (titles.length ? titles[0] : "");
    // Keep a remembered worksheet (recent-project / current-session restore)
    // when still valid.
    var want = (sel.worksheet && titles.indexOf(sel.worksheet) >= 0)
      ? sel.worksheet : preferred;
    if (titles.length > 1) {
      state.selection.worksheet = want;
      renderWorksheetList(titles);
      showWorksheetList(want);
    } else {
      // 0-1 tabs: nothing to choose — record the single tab (if any) so the
      // sheet-card dirty state matches what will open, then hide the section.
      state.selection.worksheet = titles.length === 1 ? titles[0] : "";
      hideWorksheetSection();
    }
    applyFieldStates();
    // The worksheet is settled either way, so the filename preview can resolve
    // against the tab that will actually open.
    loadSourcePreview();
  }

  function renderWorksheetList(titles) {
    if (!els.worksheetPickerMenu) return;
    els.worksheetPickerMenu.innerHTML = "";
    titles.forEach(function (title) {
      var option = el("button", "sheet-picker__option");
      option.type = "button";
      option.setAttribute("role", "option");
      option.setAttribute("data-worksheet", title);
      option.appendChild(el("span", "sheet-picker__option-main", title));
      option.addEventListener("click", function () {
        if (state.selection) state.selection.worksheet = title;
        setWorksheetSelection(title);
        closePicker("worksheet");
        // Reflect the tab change in the sheet-card dirty glow.
        applyFieldStates();
        loadSourcePreview();
      });
      els.worksheetPickerMenu.appendChild(option);
    });
  }

  function setWorksheetSelection(title) {
    updatePickerLabel("worksheet", title);
    if (!els.worksheetPickerMenu) return;
    var items = els.worksheetPickerMenu.querySelectorAll(".sheet-picker__option");
    Array.prototype.forEach.call(items, function (item) {
      item.classList.toggle("is-selected", item.getAttribute("data-worksheet") === title);
    });
  }

  function showWorksheetLoading() {
    state.worksheetLoading = true;
    if (els.worksheetSection) setHidden(els.worksheetSection, false);
    if (els.worksheetLoading) setHidden(els.worksheetLoading, false);
    if (els.worksheetPicker) setHidden(els.worksheetPicker, true);
    updateConfirmEnabled();
  }

  function showWorksheetList(selected) {
    state.worksheetLoading = false;
    if (els.worksheetLoading) setHidden(els.worksheetLoading, true);
    if (els.worksheetPicker) setHidden(els.worksheetPicker, false);
    if (els.worksheetSection) setHidden(els.worksheetSection, false);
    setWorksheetSelection(selected);
    updateConfirmEnabled();
  }

  function hideWorksheetSection() {
    state.worksheetLoading = false;
    if (els.worksheetSection) setHidden(els.worksheetSection, true);
    if (els.worksheetLoading) setHidden(els.worksheetLoading, true);
    if (els.worksheetPicker) setHidden(els.worksheetPicker, false);
    closePicker("worksheet");
    hideSourcePreview();
    updateConfirmEnabled();
  }

  // ---- Source video preview ----
  //
  // clipgen resolves a participant's footage as {study}_{participant}.mp4 (or
  // whatever the sheet's Filename row overrides it to). Getting that wrong used
  // to surface only after the workspace opened, as clips with no source — so
  // once a worksheet is settled we ask the server what it will look for and
  // whether it is already in the input folder.

  function hideSourcePreview() {
    state.previewReqVer++;   // drop anything in flight
    // An inline edit dies with the rows below; its cancel closure would
    // otherwise outlive them and fire against a detached row.
    state.cancelPreviewEdit = null;
    setHidden(els.sourcePreview, true);
    stopSourcePreviewLoading();
    // Drop the rows too: the next thing to show here belongs to a different
    // spreadsheet, and stale participants must not flash under the spinner.
    if (els.sourcePreviewList) els.sourcePreviewList.innerHTML = "";
  }

  // Reveal the block in its loading state. The head row swaps its film icon for
  // a spinner and the list (empty on a fresh selection, the previous rows on an
  // input-folder recheck) stays put, so the box appears once and grows into the
  // result rather than popping in fully formed.
  function showSourcePreviewLoading() {
    if (!els.sourcePreview) return;
    els.sourcePreview.classList.remove("is-error");
    setHidden(els.sourcePreviewSpinner, false);
    setHidden(els.sourcePreviewIcon, true);
    setStatusShimmer(els.sourcePreviewSummary, "Checking source videos…");
    setHidden(els.sourcePreview, false);
  }

  function stopSourcePreviewLoading() {
    setHidden(els.sourcePreviewSpinner, true);
    setHidden(els.sourcePreviewIcon, false);
  }

  function loadSourcePreview() {
    var sel = state.selection;
    if (!sel || (sel.type !== "google" && sel.type !== "excel") || !sel.id_or_path) {
      hideSourcePreview();
      return;
    }
    var worksheet = sel.worksheet || "";
    var inputDir = ((els.inputDir && els.inputDir.value) || "").trim();
    var key = sel.type + "|" + sel.id_or_path + "|" + worksheet + "|" + inputDir;
    var reqVer = ++state.previewReqVer;
    var cached = state.previewCache[key];
    if (cached) {
      applySourcePreview(sel, cached, reqVer);
      return;
    }
    // Only on a real fetch — a cache hit renders instantly and would just flash
    // the spinner.
    showSourcePreviewLoading();
    // Deliberately not apiGet(): this route answers a malformed spreadsheet with
    // a 400 carrying user-facing guidance, which apiGet would collapse into a
    // bare "Server error 400".
    fetch("/api/spreadsheets/preview?type=" + encodeURIComponent(sel.type) +
          "&id_or_path=" + encodeURIComponent(sel.id_or_path) +
          "&worksheet=" + encodeURIComponent(worksheet) +
          "&input_dir=" + encodeURIComponent(inputDir))
      .then(function (r) { return r.json(); })
      .then(function (data) {
        state.previewCache[key] = data;
        applySourcePreview(sel, data, reqVer);
      })
      .catch(function () {
        // Network/parse failure: stay quiet rather than blaming the sheet, but
        // never leave the spinner running.
        if (reqVer !== state.previewReqVer) return;
        stopSourcePreviewLoading();
        setHidden(els.sourcePreview, true);
      });
  }

  function applySourcePreview(sel, data, reqVer) {
    if (reqVer !== state.previewReqVer) return;                  // stale fetch
    if (!state.selection || state.selection.type !== sel.type ||
        state.selection.id_or_path !== sel.id_or_path) return;   // selection moved
    if (state.activeTab !== sel.type) return;                    // tab switched away
    if (!els.sourcePreview) return;

    stopSourcePreviewLoading();
    els.sourcePreviewList.innerHTML = "";
    if (!data || data.ok !== true) {
      els.sourcePreview.classList.add("is-error");
      els.sourcePreviewSummary.textContent =
        (data && data.error) || "Could not read this worksheet.";
      setHidden(els.sourcePreview, false);
      return;
    }
    els.sourcePreview.classList.remove("is-error");
    var rows = data.participants || [];
    if (!rows.length) {
      setHidden(els.sourcePreview, true);
      return;
    }

    renderPreviewRows(els.sourcePreviewList, els.sourcePreviewSummary, {
      type: sel.type,
      id_or_path: sel.id_or_path,
      // The worksheet the server actually resolved (an auto-picked tab has no
      // client-side name yet), because that is what an override is keyed on.
      worksheet: data.worksheet || sel.worksheet || "",
    }, data.study || "", rows, data.unmatched || []);
    setHidden(els.sourcePreview, false);
  }

  // ---- Editable preview rows ----
  //
  // One line per expected source video, shared by the spreadsheet and mind-map
  // previews. Each line can be pointed at a different file: clipgen cannot
  // write a spreadsheet back, so before this the only fix for a naming mismatch
  // was to leave the app and edit the sheet's Filename row (and a mind map has
  // no such row at all). An override is stored per user against the source's
  // identity and beats the sheet's own row; see start_settings.py.

  function renderPreviewRows(listEl, summaryEl, source, study, rows, unmatched) {
    if (!listEl) return;
    listEl.innerHTML = "";
    fillSourceDatalist(unmatched);
    var ctx = { source: source, study: study, rows: rows, summaryEl: summaryEl };
    // Missing first: with the list capped to a few rows, the problem should not
    // be the part you have to scroll for. Only on a full render — a row that
    // just got fixed stays put, rather than jumping away under the cursor.
    var ordered = rows.filter(function (r) { return !r.found; })
      .concat(rows.filter(function (r) { return r.found; }));
    var frag = document.createDocumentFragment();
    ordered.forEach(function (row) { frag.appendChild(buildPreviewRow(row, ctx)); });
    listEl.appendChild(frag);
    applyIcons(listEl);
    updatePreviewSummary(summaryEl, rows);
  }

  function updatePreviewSummary(summaryEl, rows) {
    if (!summaryEl) return;
    var missing = 0;
    rows.forEach(function (row) { if (!row.found) missing++; });
    summaryEl.textContent = missing === 0
      ? "All " + rows.length + " source videos found in the input folder"
      : (rows.length - missing) + " of " + rows.length +
        " source videos found — " + missing + " missing";
  }

  // The videos sitting in the input folder that no participant claims — i.e.
  // exactly the files an override is likely to want. One datalist for both
  // previews; only one of them is ever on screen.
  function fillSourceDatalist(unmatched) {
    if (!els.sourceFiles) return;
    els.sourceFiles.innerHTML = "";
    (unmatched || []).forEach(function (name) {
      var option = document.createElement("option");
      option.value = name;
      els.sourceFiles.appendChild(option);
    });
  }

  function previewIcon(name, cls) {
    var icon = el("span", "so-icon so-icon--xs " + cls);
    icon.setAttribute("data-icon", name);
    return icon;
  }

  function buildPreviewRow(row, ctx) {
    var item = el("div", "source-preview__row");
    item.appendChild(previewIcon("check-circle", "source-preview__status"));
    item.appendChild(el("span", "source-preview__pid", row.id));
    item.appendChild(el("span", "source-preview__name"));

    var edit = el("button", "source-preview__btn");
    edit.type = "button";
    edit.setAttribute("data-tooltip", "Use a different filename for this participant");
    edit.appendChild(previewIcon("pencil-square", "source-preview__btn-icon"));
    on(edit, "click", function () { beginPreviewEdit(item, row, ctx); });

    var reset = el("button", "source-preview__btn source-preview__reset");
    reset.type = "button";
    reset.setAttribute("data-tooltip", "Restore the default filename");
    reset.appendChild(previewIcon("arrow-uturn-left", "source-preview__btn-icon"));
    on(reset, "click", function () { commitPreviewOverride(item, row, ctx, ""); });

    item.appendChild(edit);
    item.appendChild(reset);
    paintPreviewRow(item, row);
    return item;
  }

  // Repaint one row from its (mutated in place) data. Rows are patched rather
  // than re-rendered so the list neither reorders nor loses scroll position
  // while the user works down it.
  function paintPreviewRow(item, row) {
    item.classList.toggle("is-missing", !row.found);
    item.classList.toggle("is-override", !!row.override_value);
    var status = item.querySelector(".source-preview__status");
    if (status) {
      status.setAttribute("data-icon", row.found ? "check-circle" : "exclamation-circle");
    }
    var name = item.querySelector(".source-preview__name");
    if (name) name.textContent = (row.filenames || []).join("  +  ");
    var reset = item.querySelector(".source-preview__reset");
    if (reset) reset.disabled = !row.override_value;
    applyIcons(item);
  }

  function beginPreviewEdit(item, row, ctx) {
    if (item.querySelector(".source-preview__input")) return;
    var name = item.querySelector(".source-preview__name");
    if (!name) return;
    var input = document.createElement("input");
    input.type = "text";
    input.className = "source-preview__input";
    input.autocomplete = "off";
    input.spellcheck = false;
    input.setAttribute("list", "startSourceFiles");
    input.placeholder = "filename.mp4";
    // Prefill with what the row resolves to today, so a small correction is a
    // small edit. Several files for one participant are separated by "+".
    input.value = row.override_value || (row.filenames || []).join(" + ");
    item.replaceChild(input, name);
    input.focus();
    input.select();

    // Compared against on commit, and deliberately not row.override_value: with
    // no override the field is prefilled with the *resolved* default, so
    // opening the editor and clicking away would otherwise pin that default as
    // a user override — silently overriding whatever the sheet's Filename row
    // says next week.
    var initial = input.value;
    var done = false;
    function finish(commit) {
      if (done) return;
      done = true;
      state.cancelPreviewEdit = null;
      var value = (input.value || "").trim();
      item.replaceChild(name, input);
      if (commit && value !== initial.trim()) {
        commitPreviewOverride(item, row, ctx, value);
      } else {
        paintPreviewRow(item, row);
      }
    }
    // The overlay's Escape belongs to a capture-phase modal trap, so it can't
    // be intercepted from here; open() consults this instead.
    state.cancelPreviewEdit = function () { finish(false); };
    on(input, "keydown", function (e) {
      if (e.key === "Enter") {
        e.preventDefault();
        finish(true);
      }
    });
    on(input, "blur", function () { finish(true); });
  }

  function commitPreviewOverride(item, row, ctx, value) {
    item.classList.add("is-saving");
    apiPost("/api/spreadsheets/preview/override", {
      type: ctx.source.type,
      id_or_path: ctx.source.id_or_path,
      worksheet: ctx.source.worksheet || "",
      participant: row.id,
      filename: value,
      study: ctx.study || "",
      sheet_value: row.sheet_value || "",
      input_dir: ((els.inputDir && els.inputDir.value) || "").trim(),
    })
      .then(function (r) {
        var updated = r && r.row;
        if (updated) {
          // Mutated in place: these row objects are the ones inside
          // state.previewCache, so the edit survives a re-render from cache.
          row.filenames = updated.filenames;
          row.found = updated.found;
          row.override = updated.override;
          row.override_value = updated.override_value;
        }
      })
      .catch(function (err) {
        console.error("Start overlay: could not save the filename override", err);
      })
      .then(function () {
        item.classList.remove("is-saving");
        paintPreviewRow(item, row);
        updatePreviewSummary(ctx.summaryEl, ctx.rows);
      });
  }

  // The Confirm ("Open workspace") button waits while worksheets are being
  // fetched so a fast click can't open the recommended tab before the user
  // sees (or picks) one. confirm() manages the in-flight open flag itself.
  function updateConfirmEnabled() {
    if (!els.confirmBtn) return;
    els.confirmBtn.disabled = state.confirmInFlight || state.worksheetLoading;
    // Set-or-remove: an empty data-tooltip would leave the button a hover anchor.
    if (state.worksheetLoading) {
      els.confirmBtn.setAttribute("data-tooltip", "Checking worksheets…");
    } else {
      els.confirmBtn.removeAttribute("data-tooltip");
    }
  }

  // ---- Data loading ----

  function loadStatus() {
    return apiGet("/api/status").then(function (s) {
      state.statusData = s;
      state.sheetLoaded = !!s.sheet_loaded;
      if (s.startup_notice && !state.startupNoticeShown) {
        // A window-first `-s` launch could not open its spreadsheet on the
        // boot-build thread; explain why the session is sheetless. One-shot:
        // refresh() re-runs on every overlay open and must not re-toast.
        state.startupNoticeShown = true;
        markSheetError(s.startup_notice);
      }
      // setStartTab("about") renders from statusData; if About is already
      // visible when this lands, the panel would otherwise stay on v0.0.0.
      if (state.startTab === "about") renderAbout();
      return s;
    }).catch(function (err) {
      // Same contract as loadGoogleSheets: log and keep refresh()'s boot
      // chain going — one failed link must not strand every later panel
      // on its static "Scanning…" shimmer.
      console.error("Status load failed", err);
    });
  }

  function loadDirs() {
    return apiGet("/api/dirs").then(function (d) {
      if (!d || !d.ok) return;
      if (els.inputDir) els.inputDir.value = d.input || "";
      if (els.outputDir) els.outputDir.value = d.output || "";
      renderFolderRecents(els.inputRecentsList, d.recent_inputs || [], els.inputDir, "input");
      renderFolderRecents(els.outputRecentsList, d.recent_outputs || [], els.outputDir, "output");
    }).catch(function (err) {
      console.error("Directory load failed", err);
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

  function basename(path) {
    if (!path) return "";
    var parts = String(path).replace(/[\\/]+$/, "").split(/[\\/]/);
    return parts[parts.length - 1] || path;
  }

  // One chip = one fact (input folder, output folder, spreadsheet). The icon
  // rides on data-icon so applyIcons() can resolve it against whichever host
  // page's /icons/ route mounted the overlay.
  function recentChip(icon, text, mono) {
    var chip = el("span", "rail-recent__chip");
    var glyph = el("span", "so-icon so-icon--xs");
    glyph.setAttribute("data-icon", icon);
    chip.appendChild(glyph);
    chip.appendChild(el("span", "rail-recent__chip-text" + (mono ? " mono" : ""), text));
    return chip;
  }

  // Two lines: title + relative time, then the folder/spreadsheet chips. The
  // full paths live in a single row-level data-tooltip — never one per chip,
  // since clipgenInitDataTooltips resolves via closest() and a nested
  // [data-tooltip] would shadow the row's.
  function buildRecentRow(project, currentKey) {
    var btn = el("button", "rail-recent");
    btn.type = "button";
    var isCurrent = currentKey && projectKey(project) === currentKey;
    if (isCurrent) btn.classList.add("is-current");

    var sheet = project.spreadsheet || null;
    var sheetLabel = sheet ? (sheet.label || sheet.id_or_path || "") : "";
    var name = (project.name || "").trim();
    var title = name || sheetLabel || basename(project.input) || project.input || "Untitled";

    var head = el("span", "rail-recent__head");
    head.appendChild(el("span", "rail-recent__title", title));
    var when = formatWhen(project.last_opened);
    if (when) head.appendChild(el("span", "rail-recent__when", when));
    btn.appendChild(head);

    var meta = el("span", "rail-recent__meta");
    if (project.input) meta.appendChild(recentChip("folder", basename(project.input), true));
    if (project.output && project.output !== project.input) {
      meta.appendChild(recentChip("folder-arrow-down", basename(project.output), true));
    }
    // Skip the sheet chip when the title already *is* the sheet label.
    if (sheetLabel && title !== sheetLabel) {
      meta.appendChild(recentChip("table-cells", sheetLabel, false));
    }
    btn.appendChild(meta);

    var tip = [];
    if (isCurrent) tip.push("Currently loaded");
    if (project.input) tip.push("Input: " + project.input);
    if (project.output) tip.push("Output: " + project.output);
    if (sheetLabel) {
      tip.push("Sheet: " + sheetLabel + (sheet.worksheet ? " · " + sheet.worksheet : ""));
    }
    if (tip.length) btn.setAttribute("data-tooltip", tip.join("\n"));

    btn.addEventListener("click", function () {
      setRecentsExpanded(false);
      restoreProject(project);
    });
    return btn;
  }

  function renderRailRecents(projects) {
    if (!els.railRecents) return;
    // A re-render invalidates the fold-out's contents; never leave it open
    // over a list that no longer matches.
    setRecentsExpanded(false);
    els.railRecents.innerHTML = "";
    if (els.railRecentsOverflow) els.railRecentsOverflow.innerHTML = "";
    setHidden(els.railRecentsMore, true);
    if (!projects.length) {
      els.railRecents.appendChild(el("div", "rail-recent--empty", "No recent projects yet"));
      return;
    }
    var currentKey = currentSessionKey();
    var visible = document.createDocumentFragment();
    projects.slice(0, RAIL_RECENTS_VISIBLE).forEach(function (project) {
      visible.appendChild(buildRecentRow(project, currentKey));
    });
    els.railRecents.appendChild(visible);
    applyIcons(els.railRecents);

    var rest = projects.slice(RAIL_RECENTS_VISIBLE);
    if (!rest.length || !els.railRecentsOverflow) return;
    var hidden = document.createDocumentFragment();
    // Newest-first reads top-down in the visible list, but the fold-out grows
    // upward from it — reverse so the whole stack stays chronological.
    rest.slice().reverse().forEach(function (project) {
      hidden.appendChild(buildRecentRow(project, currentKey));
    });
    els.railRecentsOverflow.appendChild(hidden);
    applyIcons(els.railRecentsOverflow);
    setHidden(els.railRecentsMore, false);
    setRecentsMoreLabel();
  }

  function setRecentsMoreLabel() {
    if (!els.railRecentsMoreLabel) return;
    if (state.recentsExpanded) {
      els.railRecentsMoreLabel.textContent = "Show less";
      return;
    }
    var rest = Math.max(0, state.recentProjects.length - RAIL_RECENTS_VISIBLE);
    els.railRecentsMoreLabel.textContent =
      clipgenPluralUnit(rest, "more project", "more projects");
  }

  function setRecentsExpanded(expanded) {
    var next = !!expanded;
    if (state.recentsExpanded === next) return;
    state.recentsExpanded = next;
    if (els.railRecentsSection) {
      els.railRecentsSection.classList.toggle("is-expanded", next);
    }
    setHidden(els.railRecentsOverflow, !next);
    if (els.railRecentsMore) {
      els.railRecentsMore.setAttribute("aria-expanded", next ? "true" : "false");
    }
    setRecentsMoreLabel();
  }

  // The input is a view of state.projectName, never the store — the guard
  // keeps a programmatic prefill from stomping the caret mid-typing.
  function setProjectName(value) {
    state.projectName = value || "";
    if (els.projectName && els.projectName.value !== state.projectName) {
      els.projectName.value = state.projectName;
    }
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
    // Keyed off the source the server actually recorded, so this mirrors
    // projectKey() for every source type. Deriving it from the spreadsheet
    // fields alone left a mind-map session keyed "input::output::", which
    // matched no stored project — no current-session highlight, and no
    // restored project name. Falls back to the sheet fields for a session
    // predating active_source (e.g. a CLI-loaded sheet with no recorded open).
    var src = s.active_source;
    var sourceKey = src && src.type && src.id_or_path
      ? src.type + "|" + src.id_or_path
      : s.sheet_loaded && s.spreadsheet_type && s.spreadsheet_id_or_path
        ? s.spreadsheet_type + "|" + s.spreadsheet_id_or_path
        : "";
    return (s.input_dir || "") + "::" + (s.output_dir || "") + "::" + sourceKey;
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

  // Accepts an ISO-8601 string (Google modifiedTime) or epoch-seconds number
  // (Excel st_mtime) and returns an "Edited …" label, or "" when unavailable.
  function formatEdited(value) {
    var when;
    if (typeof value === "number") {
      if (!value) return "";
      when = formatWhen(new Date(value * 1000).toISOString());
    } else {
      when = formatWhen(value);
    }
    return when ? "Edited " + when : "";
  }

  function restoreProject(project) {
    if (!project) return;
    // The left rail is visible on every top-level tab. Filling Open-pane
    // fields while About/Updates is showing reads as a dead click; switch
    // first so the restored values are actually on screen.
    setStartTab("open");
    // Picking a recent project is an explicit authoring act — confirming after
    // it should write that project's name, not fall back to "leave alone".
    state.projectNameAuthored = true;
    setProjectName(project.name || "");
    els.inputDir.value = project.input || "";
    els.outputDir.value = project.output || "";
    var sheet = project.spreadsheet;
    if (sheet && sheet.type && sheet.id_or_path) {
      setTab(sheet.type);
      selectSpreadsheet({
        type: sheet.type,
        id_or_path: sheet.id_or_path,
        label: sheet.label || sheet.id_or_path,
        worksheet: sheet.worksheet || "",
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
      sheetKey: selectionKey(state.selection),
    };
  }

  function loadStartSettings() {
    return apiGet("/api/start-settings").then(function (r) {
      if (!r || !r.settings) return;
      state.persistEnabled = r.settings.persist_enabled !== false;
      if (els.persist) els.persist.checked = state.persistEnabled;
      state.rememberWindow = r.settings.remember_window !== false;
      if (els.rememberWindow) els.rememberWindow.checked = state.rememberWindow;
      // A browser tab has no window for clipgen to place, so the toggle only
      // appears in a desktop launch.
      if (els.rememberWindowRow) els.rememberWindowRow.hidden = !r.desktop;
      state.recentProjects = r.settings.recent_projects || [];
      renderRailRecents(state.recentProjects);
    }).catch(function (err) {
      // Log-and-continue like the loaders below — and still render the
      // recents rail so it shows its empty state instead of staying blank.
      console.error("Start settings load failed", err);
      renderRailRecents(state.recentProjects || []);
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

  // Identity of a selection for dirty-state tracking: includes the worksheet
  // so switching tabs on the same spreadsheet reads as "dirty" (a different
  // tab would open).
  function selectionKey(sel) {
    if (!sel) return "";
    return sel.type + "|" + sel.id_or_path + "|" + (sel.worksheet || "");
  }

  function sheetState() {
    var card = root && root.querySelector(".sheet-card");
    if (card && card.classList.contains("has-error")) return "error";
    var currentKey = selectionKey(state.selection);
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

  // force=true asks the server to re-list from Drive instead of serving its
  // 5-minute cache — the Refresh button's path.
  function loadGoogleSheets(force) {
    if (!els.googleStatus) return Promise.resolve();
    setStatusShimmer(els.googleStatus, "Checking authentication…");
    setHidden(els.googlePicker, true);
    var url = "/api/spreadsheets/google" + (force ? "?refresh=true" : "");
    return apiGet(url).then(function (g) {
      if (!g.authenticated) {
        renderGoogleConnectCTA(g.auth_in_flight, g.auth_error, g);
        return;
      }
      if (g.auth_error) {
        // Signed in, but Drive didn't answer (rate limit, network).
        keepPreviousGoogleList("Google: " + g.auth_error);
        return;
      }
      state.googleSheets = g.sheets || [];
      els.googleStatus.textContent = state.googleSheets.length
        ? state.googleSheets.length + " spreadsheets available"
        : "No spreadsheets found in your account";
      renderGoogleList(state.googleSheets);
      setHidden(els.googlePicker, false);
      setHidden(els.googleRefresh, false);
    }).catch(function (err) {
      // Recovers rather than re-throwing, for two reasons: the panel is
      // mid-load (status replaced, picker hidden) and nothing downstream would
      // put it back, and on boot this sits in refresh()'s chain ahead of
      // loadExcelFiles + applyCurrentSessionPrefill, which must still run.
      console.error("Google sheet list failed", err);
      keepPreviousGoogleList("Couldn't reach clipgen.");
    });
  }

  // Leave a failed listing recoverable: whatever we already had stays
  // selectable, and Refresh stays reachable so the user can retry in place.
  function keepPreviousGoogleList(message) {
    var had = (state.googleSheets || []).length;
    els.googleStatus.textContent = had
      ? message + " Showing the last list."
      : message;
    if (had) setHidden(els.googlePicker, false);
    setHidden(els.googleRefresh, false);
  }

  // `creds` is the unauthenticated /api/spreadsheets/google payload. It is
  // absent on the poll's re-renders, so the last one seen is cached on state
  // rather than re-fetched — the setup facts don't change mid-sign-in.
  function renderGoogleConnectCTA(inFlight, errorMsg, creds) {
    if (creds && creds.credentials_paths) state.googleCreds = creds;
    els.googleStatus.innerHTML = "";
    setHidden(els.googlePicker, true);
    setHidden(els.googleRefresh, true);
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
    if (!inFlight) renderGoogleCredentialsHelp(els.googleStatus);
  }

  // Everything a first-time user needs to get past "Connect Google" and has so
  // far had no way to learn: what the file is, which three folders clipgen
  // looks in, and where Google documents making one. The server has always
  // known all of it — it just printed it to a stdout a windowed launch has no
  // console for. Collapsed by default so it stays out of the way of the
  // already-configured user.
  function renderGoogleCredentialsHelp(container) {
    var creds = state.googleCreds;
    if (!creds || !creds.credentials_paths || !creds.credentials_paths.length) return;
    var filename = creds.credentials_filename || "credentials.json";
    var found = creds.credentials_found;

    var details = el("details", "sheet-panel__creds");
    var summary = el(
      "summary",
      "sheet-panel__creds-summary",
      found ? "Using " + filename + " from " + found : "Don't have " + filename + "?"
    );
    details.appendChild(summary);

    details.appendChild(el(
      "p",
      "sheet-panel__creds-text",
      found
        ? "Sign-in uses this file. If it is the wrong project, replace it and connect again."
        : "Google Sheets access needs an OAuth client file from Google Cloud. " +
          "Save it as " + filename + " in any of these folders, then Connect:"
    ));

    var list = el("ul", "sheet-panel__creds-paths");
    for (var i = 0; i < creds.credentials_paths.length; i++) {
      var path = creds.credentials_paths[i];
      var item = el("li", "sheet-panel__creds-path", path);
      if (found && path === found) item.classList.add("is-found");
      list.appendChild(item);
    }
    details.appendChild(list);

    if (creds.credentials_guide_url) {
      var link = el("a", "sheet-panel__creds-link", "How to create " + filename + " ↗");
      link.href = creds.credentials_guide_url;
      link.target = "_blank";
      link.rel = "noopener";
      details.appendChild(link);
    }
    container.appendChild(details);
  }

  function connectGoogle() {
    setStatusShimmer(els.googleStatus, "Starting Google sign-in…");
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
    setHidden(els.googleRefresh, false);
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
      var edited = formatEdited(sheet.modifiedTime);
      if (edited) option.appendChild(el("span", "sheet-picker__option-meta", edited));
      option.addEventListener("click", function () {
        if (els.googlePaste) els.googlePaste.value = "";
        selectSpreadsheet({ type: "google", id_or_path: sheet.name, label: sheet.name });
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

  // No server-side cache to bust here — the route re-globs the input folder on
  // every call, so the Refresh button just re-runs this.
  function loadExcelFiles() {
    if (!els.excelStatus) return Promise.resolve();
    setStatusShimmer(els.excelStatus, "Scanning input folder…");
    return apiGet("/api/spreadsheets/excel").then(function (r) {
      state.excelFiles = r.files || [];
      els.excelStatus.textContent = state.excelFiles.length
        ? state.excelFiles.length + " .xlsx in " + r.input_dir
        : "No .xlsx files in " + r.input_dir;
      renderExcelList(state.excelFiles);
    }).catch(function (err) {
      // Same contract as loadGoogleSheets: never strand the panel on
      // "Scanning…", and never break the rest of refresh()'s boot chain.
      console.error("Excel scan failed", err);
      els.excelStatus.textContent = (state.excelFiles || []).length
        ? "Couldn't re-scan the input folder. Showing the last list."
        : "Could not scan the input folder.";
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
      var edited = formatEdited(f.modified);
      if (edited) option.appendChild(el("span", "sheet-picker__option-meta", edited));
      option.addEventListener("click", function () {
        if (els.excelPaste) els.excelPaste.value = "";
        selectSpreadsheet({ type: "excel", id_or_path: f.path, label: f.name });
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

  // Mind maps live in the input folder alongside the videos, so this mirrors
  // loadExcelFiles: the route re-globs on every call and Refresh just re-runs it.
  function loadMindnodeFiles() {
    if (!els.mindnodeStatus) return Promise.resolve();
    setStatusShimmer(els.mindnodeStatus, "Scanning input folder…");
    return apiGet("/api/spreadsheets/mindnode").then(function (r) {
      state.mindnodeFiles = r.files || [];
      els.mindnodeStatus.textContent = state.mindnodeFiles.length
        ? state.mindnodeFiles.length +
          (state.mindnodeFiles.length === 1 ? " mind map in " : " mind maps in ") +
          r.input_dir
        : "No .mindnode documents in " + r.input_dir;
      renderMindnodeList(state.mindnodeFiles);
    }).catch(function (err) {
      console.error("MindNode scan failed", err);
      els.mindnodeStatus.textContent = (state.mindnodeFiles || []).length
        ? "Couldn't re-scan the input folder. Showing the last list."
        : "Could not scan the input folder.";
    });
  }

  function renderMindnodeList(files) {
    if (!els.mindnodePickerMenu) return;
    els.mindnodePickerMenu.innerHTML = "";
    if (!files.length) {
      els.mindnodePickerMenu.appendChild(
        el("div", "sheet-picker__empty", "No .mindnode documents in the input folder")
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
      var edited = formatEdited(f.modified);
      if (edited) option.appendChild(el("span", "sheet-picker__option-meta", edited));
      option.addEventListener("click", function () {
        if (els.mindnodePaste) els.mindnodePaste.value = "";
        // No worksheets to fetch, so setSelection is the whole flow.
        setSelection({ type: "mindnode", id_or_path: f.path, label: f.name });
        closePicker("mindnode");
      });
      els.mindnodePickerMenu.appendChild(option);
    });
    if (state.selection && state.selection.type === "mindnode") {
      highlightMindnodeSelection(state.selection.id_or_path);
    }
  }

  function highlightMindnodeSelection(path) {
    if (!els.mindnodePickerMenu) return;
    var items = els.mindnodePickerMenu.querySelectorAll(".sheet-picker__option");
    Array.prototype.forEach.call(items, function (item) {
      item.classList.toggle("is-selected", item.getAttribute("data-path") === path);
    });
  }

  // Summarize the chosen map before it is opened — the equivalent of the
  // source-video preview the spreadsheet tabs get, plus the bundle's own
  // QuickLook render so the researcher can confirm it is the right document.
  function loadMindnodePreview(sel) {
    if (!els.mindnodePreview) return;
    var reqVer = ++state.mindnodePreviewReqVer;
    if (!sel || sel.type !== "mindnode" || !sel.id_or_path) {
      setHidden(els.mindnodePreview, true);
      return;
    }
    var inputDir = ((els.inputDir && els.inputDir.value) || "").trim();
    apiGet("/api/spreadsheets/mindnode/preview?path=" +
           encodeURIComponent(sel.id_or_path) +
           "&input_dir=" + encodeURIComponent(inputDir))
      .then(function (r) {
        if (reqVer !== state.mindnodePreviewReqVer) return;   // stale fetch
        applyMindnodePreview(sel, r);
      })
      .catch(function (err) {
        if (reqVer !== state.mindnodePreviewReqVer) return;
        setHidden(els.mindnodePreviewThumb, true);
        els.mindnodePreviewSummary.textContent =
          (err && err.message) || "Could not read this mind map.";
        setHidden(els.mindnodePreview, false);
      });
  }

  function applyMindnodePreview(sel, data) {
    var parts = [data.notes + (data.notes === 1 ? " note" : " notes")];
    if (data.without_times) {
      // Said up front: these are the notes that cannot become clips.
      parts.push(data.without_times + " without a timestamp");
    }
    parts.push(
      data.participants.length +
        (data.participants.length === 1 ? " participant" : " participants") +
        (data.participants.length ? " (" + data.participants.join(", ") + ")" : "")
    );
    els.mindnodePreviewSummary.textContent =
      "Study “" + data.study + "” — " + parts.join(", ");
    if (data.has_preview && els.mindnodePreviewThumb) {
      els.mindnodePreviewThumb.src =
        "/api/spreadsheets/mindnode/thumb?path=" + encodeURIComponent(sel.id_or_path);
      setHidden(els.mindnodePreviewThumb, false);
    } else {
      setHidden(els.mindnodePreviewThumb, true);
    }
    // Same editable file list the spreadsheet tabs get. A mind map has no
    // Filename row, so an override set here is the only way to point a
    // participant at differently-named footage.
    var sources = data.sources || [];
    if (sources.length) {
      renderPreviewRows(els.mindnodeSourcesList, els.mindnodeSourcesSummary, {
        type: "mindnode",
        id_or_path: sel.id_or_path,
        worksheet: "",
      }, data.study || "", sources, data.unmatched || []);
    }
    setHidden(els.mindnodeSources, !sources.length);
    setHidden(els.mindnodePreview, false);
  }

  function loadLicenses() {
    return apiGet("/api/licenses").then(function (r) {
      state.licensesLoaded = true;
      state.licenseComponents = (r && r.components) || [];
      renderAttribution();
    }).catch(function () {
      // A checkout without build/THIRD-PARTY-LICENSES still shows the section
      // heading and the link out; only the generated list stays empty.
      state.licensesLoaded = true;
      state.licenseComponents = [];
      renderAttribution();
    });
  }

  function loadChangelog() {
    return apiGet("/api/changelog").then(function (r) {
      state.changelogLoaded = true;
      var entries = (r && r.entries) || [];
      renderChangelog(entries);
      if (entries.length > 0 && els.updatesBadge) {
        // Count changes, not releases: a release holding four of them should
        // not read as "1 update".
        var changes = entries.reduce(function (n, e) {
          return n + ((e.changes && e.changes.length) || 0);
        }, 0);
        els.updatesBadge.textContent = String(changes);
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
      head.appendChild(el("span", "changelog__when", entry.date || ""));
      item.appendChild(head);
      (entry.changes || []).forEach(function (change) {
        var row = el("div", "changelog__change");
        var tag = el("span", "changelog__tag");
        tag.setAttribute("data-tool", change.tool || "Core");
        tag.appendChild(el("span", "changelog__tag-dot"));
        tag.appendChild(document.createTextNode(change.tool || "Core"));
        row.appendChild(tag);
        var text = el("span", "changelog__body");
        if (change.kind) {
          text.appendChild(el("span", "changelog__kind", change.kind + ":"));
        }
        text.appendChild(document.createTextNode(change.text || ""));
        row.appendChild(text);
        item.appendChild(row);
      });
      els.changelogList.appendChild(item);
    });
  }

  function renderAbout() {
    if (!els.aboutGrid) return;
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

  // Rows arrive in THIRD-PARTY-LICENSES order, which already clusters components
  // by license, so a heading is opened whenever the group changes rather than by
  // re-sorting. Nested rows (the cv2 FFmpeg DLL, the PP-OCR models) belong to the
  // component above them and never open a group of their own.
  function renderAttribution() {
    if (!els.attributionList) return;
    var rows = state.licenseComponents || [];
    els.attributionList.innerHTML = "";
    if (!rows.length) return;

    var frag = document.createDocumentFragment();
    var currentGroup = null;
    rows.forEach(function (item) {
      if (!item.nested && item.group && item.group !== currentGroup) {
        currentGroup = item.group;
        frag.appendChild(el("div", "attribution__group", currentGroup));
      }
      var row = el("div", "attribution__row");
      if (item.nested) row.classList.add("attribution__row--nested");
      row.appendChild(el("div", "attribution__name", item.component || ""));
      row.appendChild(el("div", "attribution__version", item.version || ""));
      // The full cell, not the group: "MIT (macOS only)" and the cv2 DLL's
      // platform caveat are the part a reader actually needs.
      row.appendChild(el("div", "attribution__license", item.license || ""));
      frag.appendChild(row);
    });
    els.attributionList.appendChild(frag);
  }

  // ---- Open / dismiss flows ----

  function confirm() {
    // Guard against re-entry: the click handler disables the button, but the
    // Cmd/Ctrl+Enter shortcut bypasses that path. Two simultaneous flights
    // would race the server-side _swap_worksheet.
    if (state.confirmInFlight) return;
    // Hold off while the worksheet list is loading so we never post a
    // selection whose worksheet field isn't settled yet (which would open the
    // priority tab instead of the dropdown default). The button is disabled in
    // this state; this also blocks the Cmd/Ctrl+Enter shortcut.
    if (state.worksheetLoading) return;
    state.confirmInFlight = true;

    var inputVal = (els.inputDir.value || "").trim();
    var outputVal = (els.outputDir.value || "").trim();
    // null omits the field, which the server reads as "keep the stored name".
    // Sending "" is a deliberate clear, so only do it once we actually know
    // what was stored — a Cmd+Enter that beats refresh() must not wipe a label
    // the user never saw, let alone edited.
    var nameVal = (state.projectNameAuthored || state.projectNamePrefilled)
      ? state.projectName.trim()
      : null;

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
      updateConfirmEnabled();
    }

    updateConfirmEnabled();
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
        // "No spreadsheet" has to actually close whatever is open, not just
        // record the session. Nothing else in the UI ever posted to
        // /api/spreadsheets/close, so a source opened earlier lived for the
        // rest of the process — a mind map especially, since it is never
        // cleared by opening a sheet either.
        var st = state.statusData || {};
        var needsClose = !!(st.sheet_loaded || st.mindnode_loaded);
        // One call per source: the route drops the mind map for
        // {type: "mindnode"} and the worksheet otherwise, and the two coexist,
        // so closing both takes both calls. Each is checked for r.ok like the
        // open path below — the route refuses with 409 while a generation is
        // running, and swallowing that would record a no-spreadsheet session
        // and reload with the source still very much loaded.
        function postClose(body) {
          return fetch("/api/spreadsheets/close", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
          }).then(function (r) {
            return r.json().then(
              function (j) { return { ok: r.ok && j && j.ok !== false, body: j }; },
              function () { return { ok: false, body: null }; }
            );
          });
        }
        var closeStep = Promise.resolve({ ok: true, body: null });
        function chainClose(prev, payload) {
          return prev.then(function (res) {
            if (!res.ok) return res; // first failure wins; don't keep closing
            return postClose(payload);
          });
        }
        if (st.mindnode_loaded) {
          closeStep = chainClose(closeStep, { type: "mindnode" });
        }
        if (st.sheet_loaded) {
          closeStep = chainClose(closeStep, {});
        }
        closeStep
          .then(function (res) {
            if (!res.ok) {
              releaseConfirm();
              markSheetError(
                (res.body && res.body.error) || "Could not close the current source"
              );
              return null;
            }
            return recordSession(inputVal, outputVal, null, nameVal).finally(
              function () {
                releaseConfirm();
                // Reload only when something was actually unloaded — the other
                // frontends are holding data for a source that is now gone.
                if (needsClose) {
                  window.location.reload();
                  return;
                }
                close();
              }
            );
          })
          .catch(function (err) {
            releaseConfirm();
            markSheetError("Close failed: " + (err && err.message));
          });
        return;
      }
      // Built explicitly rather than posting state.selection verbatim so the
      // name rides along with the open that records the session server-side.
      var openPayload = {
        type: state.selection.type,
        id_or_path: state.selection.id_or_path,
        label: state.selection.label,
        worksheet: state.selection.worksheet || "",
      };
      if (nameVal !== null) openPayload.project_name = nameVal;
      fetch("/api/spreadsheets/open", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(openPayload),
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

  function recordSession(input, output, spreadsheet, name) {
    // Fire-and-forget on the no-sheet path. The sheet-open path records on
    // the server during /api/spreadsheets/open, so this is the only call
    // site that needs to invoke /api/sessions/record explicitly.
    var payload = { input: input, output: output };
    // Omitted (not "") when the stored name isn't known yet — see confirm().
    if (name !== null && name !== undefined) payload.name = name;
    if (spreadsheet) payload.spreadsheet = spreadsheet;
    return apiPost("/api/sessions/record", payload).catch(function (err) {
      console.warn("Session record failed", err);
    });
  }

  // ---- Open / close ----

  function open(tab) {
    if (!state.mounted) {
      // Not mount().then(open): then() would pass mount's resolved value as tab.
      // Re-enter only if the mount actually succeeded — mount() resolves (never
      // rejects) on a missing slot or a failed template fetch, and recursing on
      // that would spin forever re-fetching start-overlay.html.
      mount().then(function () { if (state.mounted) open(tab); });
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
        // Escape cancels an inline filename edit first, then peels the recents
        // fold-out, then dismisses the overlay.
        onEscape: function () {
          if (state.cancelPreviewEdit) {
            state.cancelPreviewEdit();
            return;
          }
          if (state.recentsExpanded) {
            setRecentsExpanded(false);
            return;
          }
          close();
        },
        trapFocus: true,
        initialFocus: els.panel,
        restoreFocus: true
      });
    }
    // Always land on the workspace form, wherever the user left the tabs —
    // unless the caller asked for a specific one (the desktop Help menu's
    // "What's New…" opens straight onto the changelog).
    setStartTab(tab || "open");
    runIntro();
    refresh();
  }

  function close() {
    if (!root) return;
    state.open = false;
    setRecentsExpanded(false);
    markDismissed();
    stopGooglePoll();
    // Don't let a debounced preview fetch fire into a closed overlay.
    if (state.previewDirTimer) {
      clearTimeout(state.previewDirTimer);
      state.previewDirTimer = null;
    }
    // Animate the panel and backdrop out, then hide.
    if (els.panel) els.panel.classList.remove("is-in");
    root.classList.remove("is-veiled");
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
    root.classList.remove("is-veiled");
    setTimeout(function () {
      root.classList.add("is-veiled");
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
      // Wrapped, not passed by reference: loadGoogleSheets(force) would
      // otherwise receive the previous link's resolved value as `force` and
      // re-list from Drive on every overlay open.
      .then(function () { return loadGoogleSheets(); })
      .then(loadExcelFiles)
      .then(loadMindnodeFiles)
      .then(applyCurrentSessionPrefill)
      // Reads CHANGELOG.md off disk, so it is cheap enough to do up front —
      // and the Recent updates count badge only exists if we do.
      .then(loadChangelog)
      .catch(function (err) {
        console.error("Start overlay refresh failed", err);
      });
  }

  function applyCurrentSessionPrefill() {
    var s = state.statusData || {};
    // A mind map is an independent source with no worksheet, so it branches ahead
    // of the spreadsheet cases. Gate on active_source, NOT mindnode_loaded: the
    // two coexist by design, so a bare mindnode_loaded check let one opened mind
    // map hijack the prefill for the rest of the session — recents would highlight
    // the spreadsheet opened afterwards while this panel switched to the Mind map
    // tab, and confirming re-opened the map. currentSessionKey() already keys on
    // active_source and the two must agree. An empty active_source keeps the
    // original precedence, so a mind-map-only launch still prefills.
    var activeType = (s.active_source && s.active_source.type) || "";
    var mindnodeIsActive = activeType === "mindnode" || !activeType;
    if (mindnodeIsActive && s.mindnode_loaded && s.mindnode_path) {
      setTab("mindnode");
      setSelection({
        type: "mindnode",
        id_or_path: s.mindnode_path,
        label: s.mindnode_label || s.mindnode_path,
      });
      var known = (state.mindnodeFiles || []).some(function (f) {
        return f.path === s.mindnode_path;
      });
      if (!known && els.mindnodePaste) els.mindnodePaste.value = s.mindnode_path;
      state.baseline = baselineFromInputs();
      applyFieldStates();
      return;
    }
    if (s.sheet_loaded && s.spreadsheet_type && s.spreadsheet_id_or_path) {
      // Seed the picker so the current session shows as "loaded" without the
      // user needing to re-select. Activate the correct tab + selection, and
      // restore the loaded worksheet so re-confirming can't silently switch to
      // the priority tab (and the dropdown reflects the active tab).
      setTab(s.spreadsheet_type);
      selectSpreadsheet({
        type: s.spreadsheet_type,
        id_or_path: s.spreadsheet_id_or_path,
        label: s.spreadsheet_label || s.spreadsheet_id_or_path,
        worksheet: s.spreadsheet_worksheet || "",
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
    // The name isn't a server-side field — it lives on the recent-projects
    // record, so the current session's label is whichever stored project
    // matches the session key. refresh() is fire-and-forget, so this can land
    // after the user has already started typing; their value wins.
    if (!state.projectNameAuthored) {
      var currentKey = currentSessionKey();
      var current = currentKey && state.recentProjects.filter(function (p) {
        return projectKey(p) === currentKey;
      })[0];
      setProjectName(current ? (current.name || "") : "");
    }
    state.projectNamePrefilled = true;
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
