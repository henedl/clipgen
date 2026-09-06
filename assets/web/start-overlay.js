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
  // Google OAuth runs server-side; poll /api/spreadsheets/google until authenticated, auth_error, or timeout.
  var GOOGLE_POLL_TIMEOUT_MS = 90 * 1000;
  var GOOGLE_POLL_INTERVAL_MS = 1500;
  var GOOGLE_POLL_RETRY_MS = 2500;
  // Wait for typing to settle before fetching worksheets for a pasted URL/name.
  var WORKSHEET_PASTE_DEBOUNCE_MS = 600;
  // Rail shows this many recents; the fold-out holds the rest (server caps at RECENTS_CAP).
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
    changelogEntries: [],
    // Last /api/update/status snapshot; null until the desktop app answers.
    update: null,
    updatePoller: null,
    // Fetched once; kept in state because renderAttribution() re-runs on every About activation.
    licensesLoaded: false,
    licenseComponents: [],
    googleSheets: [],
    // Last unauthenticated Google payload (credentials paths, setup link); the auth poll re-renders without one.
    googleCreds: null,
    excelFiles: [],
    mindnodeFiles: [],
    mindnodePreviewReqVer: 0,   // rejects stale mind-map summary fetches
    statusData: null,
    recentProjects: [],
    projectName: "",          // the optional label typed in the right column
    // Name prefills from recents; these flags stop early Cmd+Enter or late prefill clobbering it.
    projectNamePrefilled: false,  // applyCurrentSessionPrefill has run at least once
    startupNoticeShown: false,  // the boot -s failure toast fired (highlight persists)
    projectNameAuthored: false,   // the user typed it, or picked a recent project
    recentsExpanded: false,   // rail fold-out revealing projects past RAIL_RECENTS_VISIBLE
    worksheetsCache: {},      // "type|id_or_path" -> { worksheets, recommended }
    worksheetReqVer: 0,       // rejects stale worksheet fetches
    worksheetLoading: false,  // a worksheet fetch is in flight (gates Confirm)
    wsPasteTimer: null,       // debounce for paste-driven worksheet loads
    // Source-video preview, keyed "type|id_or_path|worksheet|inputDir"; the folder decides found/missing.
    previewCache: {},
    previewReqVer: 0,         // rejects stale preview fetches
    previewDirTimer: null,    // debounce for input-folder-driven refreshes
    // Set during an inline filename edit; the capture-phase modal trap's onEscape checks it first.
    cancelPreviewEdit: null,
    // Snapshot at open (or recent-project click); drives the .is-loaded / .is-dirty glow.
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

  // Spin a Refresh button while its loader runs; loaders handle their own failures.
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

  // Shimmer text in its own span: .cg-shimmer's transparent fill would leak onto CTA rows.
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
        // Page-level clipgenInitBrandMark ran before the overlay existed; re-run to hydrate the rail mark.
        if (typeof window.clipgenInitBrandMark === "function") {
          window.clipgenInitBrandMark();
        }
        state.mounted = true;
        // A snapshot may have arrived before the tabs existed.
        syncUpdateBadges();
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
      // The folder decides found/missing, so re-resolve the preview once typing settles.
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
      // Bypass the server's 5-minute Drive cache for a spreadsheet created mid-session.
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
      // Bound for the page's lifetime (close() only hides), so bail when not open.
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
        // No worksheets on a mind map: setSelection alone settles it and frees Confirm.
        setSelection({ type: "mindnode", id_or_path: v, label: v.split("/").pop() || v });
      } else if (state.selection && state.selection.type === "mindnode") {
        renderMindnodeList(state.mindnodeFiles || []);
      }
    });

    // Escape/Tab belong to the blocking-modal trap (see open()); the rest are registry hotkeys.
    if (window.ClipgenHotkeys) {
      var isOpen = function () { return state.open; };
      // Form shortcuts need the Open pane, else they mutate a hidden form from About/Updates.
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
          // Fire input so the glow and clear-error logic run as if typed.
          input.dispatchEvent(new Event("input", { bubbles: true }));
          input.focus();
        } else {
          // Cancelled, or no native dialog on this platform: just focus the field.
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
    // Only a real switch animates; re-selecting or the open-time reset would fight runIntro()'s cascade.
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
    // Re-render every activation: statusData may be empty at first; latching would freeze v0.0.0.
    if (name === "about") {
      renderAbout();
      renderAttribution();
      if (!state.licensesLoaded) loadLicenses();
    }
  }

  // One-shot enter animation; the class must come off so the next switch replays.
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
    // New spreadsheet identity: invalidate in-flight worksheet fetches and reset the dropdown.
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
    // Tab listing can be slow (Google opens server-side); hold Confirm until tabs are known.
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
      // 0-1 tabs: record the single tab so dirty state matches what opens, then hide.
      state.selection.worksheet = titles.length === 1 ? titles[0] : "";
      hideWorksheetSection();
    }
    applyFieldStates();
    // Worksheet settled either way; the filename preview can resolve against it.
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
  // Catches {study}_{participant}.mp4 naming mismatches before the workspace opens.

  function hideSourcePreview() {
    state.previewReqVer++;   // drop anything in flight
    // Rows are going away; a live cancel closure would fire against a detached row.
    state.cancelPreviewEdit = null;
    setHidden(els.sourcePreview, true);
    stopSourcePreviewLoading();
    // Drop the rows: stale participants must not flash under the next spinner.
    if (els.sourcePreviewList) els.sourcePreviewList.innerHTML = "";
  }

  // Reveal in loading state; existing rows stay so the box grows into the result.
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
    // Cache hits render instantly; the spinner would only flash.
    showSourcePreviewLoading();
    // Not apiGet(): a 400 here carries user-facing guidance apiGet would collapse.
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
      // Overrides key on the server-resolved worksheet; an auto-picked tab has no client name.
      worksheet: data.worksheet || sel.worksheet || "",
    }, data.study || "", rows, data.unmatched || []);
    setHidden(els.sourcePreview, false);
  }

  // ---- Editable preview rows ----
  // Per-user overrides beat the sheet's Filename row; see start_settings.py.

  function renderPreviewRows(listEl, summaryEl, source, study, rows, unmatched) {
    if (!listEl) return;
    listEl.innerHTML = "";
    fillSourceDatalist(unmatched);
    var ctx = { source: source, study: study, rows: rows, summaryEl: summaryEl };
    // Missing first, but only on a full render so a fixed row stays put.
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

  // Unclaimed input-folder videos: the likely override targets. One datalist serves both previews.
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

  // Patch in place so the list neither reorders nor loses scroll position.
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
    // Prefill with today's resolution; multiple files join with "+".
    input.value = row.override_value || (row.filenames || []).join(" + ");
    item.replaceChild(input, name);
    input.focus();
    input.select();

    // Not row.override_value: clicking away must not pin the resolved default as an override.
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
    // Escape belongs to the modal trap; open()'s onEscape calls this instead.
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
          // Mutated in place: these rows live in state.previewCache, so cached re-renders keep the edit.
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

  // Hold Confirm until worksheets list, so a fast click can't open an unseen tab.
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

  function loadStatus(force) {
    // Shared memoized fetch (utils.js); force=true bypasses the page-load snapshot.
    return clipgenStatus(force).then(function (s) {
      state.statusData = s;
      state.sheetLoaded = !!s.sheet_loaded;
      // startup_notice is handled in applyCurrentSessionPrefill, whose setTab would wipe a highlight set here.
      if (state.startTab === "about") renderAbout();
      // The installed-version highlight reads statusData too.
      if (state.changelogLoaded) renderChangelog(state.changelogEntries);
      return s;
    }).catch(function (err) {
      // Log and continue so refresh()'s boot chain reaches every panel.
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

  // One chip per fact; data-icon lets applyIcons() resolve against the host page's /icons/ route.
  function recentChip(icon, text, mono) {
    var chip = el("span", "rail-recent__chip");
    var glyph = el("span", "so-icon so-icon--xs");
    glyph.setAttribute("data-icon", icon);
    chip.appendChild(glyph);
    chip.appendChild(el("span", "rail-recent__chip-text" + (mono ? " mono" : ""), text));
    return chip;
  }

  // One row-level data-tooltip for full paths; a nested one would shadow it via closest().
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
    // A re-render invalidates the fold-out; never leave it open over a changed list.
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
    // The fold-out grows upward, so reverse to keep the stack chronological.
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

  // The input mirrors state.projectName; the guard keeps a prefill from moving the caret mid-typing.
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
    // Mirror projectKey() using active_source, else a mind-map session keys as "input::output::" and matches nothing.
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

  // value: ISO-8601 string (Google) or epoch seconds (Excel st_mtime). Returns "Edited …" or "".
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
    // The rail shows on every tab; switch to Open so the restored fields are visible.
    setStartTab("open");
    // Picking a recent is authoring: confirming should write its name, not leave it alone.
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
    // Reset the baseline so the glow reads "loaded", not "dirty".
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
      // Only a desktop launch has a window to place.
      if (els.rememberWindowRow) els.rememberWindowRow.hidden = !r.desktop;
      state.recentProjects = r.settings.recent_projects || [];
      renderRailRecents(state.recentProjects);
    }).catch(function (err) {
      // Log and continue; still render the rail so it shows its empty state.
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

  // Includes the worksheet: a different tab on the same spreadsheet reads as dirty.
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

  // force=true re-lists from Drive instead of the server's 5-minute cache (Refresh button).
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
      // Recover, don't rethrow: nothing downstream restores the panel, and Promise.all must reach applyCurrentSessionPrefill.
      console.error("Google sheet list failed", err);
      keepPreviousGoogleList("Couldn't reach clipgen.");
    });
  }

  // Keep the last list selectable and Refresh reachable so the user can retry.
  function keepPreviousGoogleList(message) {
    var had = (state.googleSheets || []).length;
    els.googleStatus.textContent = had
      ? message + " Showing the last list."
      : message;
    if (had) setHidden(els.googlePicker, false);
    setHidden(els.googleRefresh, false);
  }

  // creds is the unauthenticated payload; poll re-renders omit it, so cache the last one.
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

  // Credentials setup help a windowed launch has no stdout for. Collapsed by default.
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

  // No server cache: the route re-globs each call, so Refresh just re-runs this.
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
      // Log and continue: never strand the panel on "Scanning…" or break refresh()'s chain.
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

  // Mirrors loadExcelFiles: the route re-globs each call, so Refresh just re-runs this.
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

  // Mind-map counterpart of the source-video preview, plus the bundle's QuickLook render.
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
    // Same editable list as spreadsheets; a mind map has no Filename row to edit.
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
      // Without build/THIRD-PARTY-LICENSES the heading and link still show; only the list stays empty.
      state.licensesLoaded = true;
      state.licenseComponents = [];
      renderAttribution();
    });
  }

  function loadChangelog() {
    return apiGet("/api/changelog").then(function (r) {
      state.changelogLoaded = true;
      state.changelogEntries = (r && r.entries) || [];
      renderChangelog(state.changelogEntries);
    }).catch(function () {
      state.changelogLoaded = true;
      state.changelogEntries = [];
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
    var installed = String((state.statusData || {}).version || "").replace(/^v/, "");
    entries.forEach(function (entry) {
      var item = el("li", "changelog__entry");
      var head = el("div", "changelog__head");
      var version = el("span", "changelog__version", entry.version);
      var isCurrent = !!installed && String(entry.version || "").replace(/^v/, "") === installed;
      if (isCurrent) {
        item.classList.add("changelog__entry--current");
        version.classList.add("changelog__version--current");
      }
      head.appendChild(version);
      if (isCurrent) head.appendChild(el("span", "changelog__current", "installed"));
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
    var u = state.update;
    aboutRow("Version", function (val) {
      val.appendChild(el("span", "mono", "v" + (s.version || "0.0.0")));
      if (u && u.supported) val.appendChild(buildAutoCheckToggle(u));
    });
    if (u && u.supported) {
      aboutRow("Update", function (val) { buildUpdateRow(val, u); });
    }
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

  // ---- Self-update (frozen desktop app only) ----

  function formatMegabytes(bytes) {
    return Math.round((bytes || 0) / 1048576) + " MB";
  }

  function updateAction(path) {
    return apiPost(path, {}).then(applyUpdateSnapshot).catch(function () {});
  }

  function buildUpdateRow(val, u) {
    var phase = u.phase;
    var line = el("div", "about__update");
    var actions = el("div", "about__update-actions");
    function button(label, primary, onClick) {
      var b = document.createElement("button");
      b.type = "button";
      b.className = "so-btn so-btn--sm" + (primary ? " so-btn--primary" : "");
      b.textContent = label;
      b.addEventListener("click", onClick);
      actions.appendChild(b);
    }
    function releaseLink(label) {
      if (!u.release_url) return;
      var link = document.createElement("a");
      link.className = "about__link";
      link.href = u.release_url;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.textContent = label;
      actions.appendChild(link);
    }
    if (phase === "checking") {
      line.textContent = "Checking for updates…";
    } else if (phase === "available") {
      line.textContent = u.version + " is available.";
      button("Download", true, function () { updateAction("/api/update/download"); });
      button("Skip this version", false, function () { updateAction("/api/update/skip"); });
      releaseLink("Release notes");
    } else if (phase === "downloading") {
      line.textContent = "Downloading " + u.version + "… " + formatMegabytes(u.completed) + " of " + formatMegabytes(u.total);
      var track = el("div", "about__progress");
      var fill = el("div", "about__progress-fill");
      fill.style.width = (u.total ? Math.round(100 * u.completed / u.total) : 0) + "%";
      track.appendChild(fill);
      val.appendChild(line);
      val.appendChild(track);
      return;
    } else if (phase === "ready" && u.error) {
      line.textContent = u.error + ". Install " + u.version + " by hand.";
      button("Show download", false, function () { updateAction("/api/update/reveal"); });
      releaseLink("Open Releases page");
    } else if (phase === "ready") {
      line.textContent = u.version + " is downloaded.";
      button("Restart to update", true, function () { updateAction("/api/update/apply"); });
      button("Skip this version", false, function () { updateAction("/api/update/skip"); });
      releaseLink("Release notes");
    } else if (phase === "applying") {
      line.textContent = "Installing " + u.version + "… clipgen will restart.";
    } else if (phase === "error") {
      line.textContent = u.error || "Update failed.";
      button("Retry", false, function () { checkForUpdates(true); });
      releaseLink("Open Releases page");
    } else if (u.skipped) {
      line.textContent = u.skipped + " skipped.";
      button("Check again", false, function () { checkForUpdates(true); });
      releaseLink("Release notes");
    } else {
      line.textContent = u.checked ? "You're on the latest version." : "";
      button(u.checked ? "Check again" : "Check for updates", false, function () { checkForUpdates(true); });
    }
    val.appendChild(line);
    if (actions.childNodes.length) val.appendChild(actions);
    if (u.last_error) {
      val.appendChild(el("div", "about__sub", "Last update failed: " + u.last_error));
    }
  }

  // Beside the version: the UPDATE_CHECK_ON_LAUNCH setting via the shared settings route.
  function buildAutoCheckToggle(u) {
    var label = el("label", "persist-check about__auto");
    var input = document.createElement("input");
    input.type = "checkbox";
    input.checked = !!u.auto_check;
    // Hotkeys treat a focused checkbox as a typing target; keep focus off it.
    input.tabIndex = -1;
    input.addEventListener("change", function () {
      input.blur();
      apiPut("/api/settings", { settings: { UPDATE_CHECK_ON_LAUNCH: input.checked } })
        .then(function () { return apiGet("/api/update/status"); })
        .then(applyUpdateSnapshot)
        .catch(function () {});
    });
    label.appendChild(input);
    label.appendChild(el("span", null, "Check for updates automatically"));
    return label;
  }

  // The Start button's dot says "something is waiting"; the About tab's dot says where.
  function syncUpdateBadges() {
    var u = state.update;
    var pending = !!(u && u.supported && (u.phase === "available" || u.phase === "ready"));
    if (window.ClipgenTopNav && typeof window.ClipgenTopNav.setStartBadge === "function") {
      window.ClipgenTopNav.setStartBadge(pending);
    }
    var aboutTab = root ? root.querySelector('[data-start-tab="about"]') : null;
    if (aboutTab) aboutTab.classList.toggle("has-badge", pending);
  }

  function applyUpdateSnapshot(u) {
    if (!u || !u.ok) return;
    state.update = u;
    var busy = u.phase === "checking" || u.phase === "downloading" || u.phase === "applying";
    if (busy) startUpdatePoll(); else stopUpdatePoll();
    syncUpdateBadges();
    if (state.mounted && state.open && state.startTab === "about") renderAbout();
  }

  function checkForUpdates(force) {
    return apiPost("/api/update/check", { force: !!force }).then(applyUpdateSnapshot).catch(function () {});
  }

  function startUpdatePoll() {
    if (state.updatePoller) return;
    var misses = 0;
    state.updatePoller = createPoller(function () {
      return apiGet("/api/update/status").then(function (u) {
        misses = 0;
        applyUpdateSnapshot(u);
      }).catch(function () {
        // The server vanishing mid-apply is the restart, not a failure.
        if (++misses >= 5) stopUpdatePoll();
      });
    }, 1000, { runImmediately: false, label: "start.update" });
    state.updatePoller.start();
  }

  function stopUpdatePoll() {
    if (!state.updatePoller) return;
    state.updatePoller.stop();
    state.updatePoller = null;
  }

  function aboutRow(label, build) {
    var row = el("div", "about__row");
    row.appendChild(el("div", "about__label", label));
    var val = el("div", "about__value");
    build(val);
    row.appendChild(val);
    els.aboutGrid.appendChild(row);
  }

  // Rows arrive grouped by license: heading on group change; nested rows never start one.
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
      // The full cell, not the group: platform caveats like "MIT (macOS only)" matter.
      row.appendChild(el("div", "attribution__license", item.license || ""));
      frag.appendChild(row);
    });
    els.attributionList.appendChild(frag);
  }

  // ---- Open / dismiss flows ----

  function confirm() {
    // Re-entry guard: Cmd/Ctrl+Enter bypasses the disabled button and would race _swap_worksheet.
    if (state.confirmInFlight) return;
    // An unsettled worksheet field would open the priority tab instead of the dropdown default.
    if (state.worksheetLoading) return;
    state.confirmInFlight = true;

    var inputVal = (els.inputDir.value || "").trim();
    var outputVal = (els.outputDir.value || "").trim();
    // null keeps the stored name; "" clears it, so only send "" once known.
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
        // "No spreadsheet" must close the open source too; nothing else ever posts /api/spreadsheets/close.
        var st = state.statusData || {};
        var needsClose = !!(st.sheet_loaded || st.mindnode_loaded);
        // One call per coexisting source; check r.ok because the route 409s mid-generation.
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
                // Reload only if something unloaded; other frontends hold data for the gone source.
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
      // Built explicitly so the name rides along with the session-recording open.
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
          // Hard reload so every frontend re-fetches; the flag stays set since the page unmounts.
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
    // Only the no-sheet path records explicitly; /api/spreadsheets/open records server-side.
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
      // Not mount().then(open): that passes the resolved value as tab, and mount() resolves on failure.
      mount().then(function () { if (state.mounted) open(tab); });
      return;
    }
    if (state.open) return;
    state.open = true;
    show(root, true);
    // Blocking modal. initialFocus is the non-typing panel so letter hotkeys work at once.
    if (typeof openBlockingModal === "function") {
      openBlockingModal(root, {
        // Escape: cancel an inline edit, then fold the recents, then dismiss.
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
    // Default to the form; the desktop Help menu's "What's New…" passes "updates".
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
      // Release only once hidden: an early release re-arms page hotkeys during the fade-out.
      if (typeof closeBlockingModal === "function") closeBlockingModal(root);
    }, 460);
  }

  // Replay the rail mark's draw every open, bypassing clipgenInitBrandMark's session gate; retry until hydrated.
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
    // Concurrent: the Drive call must not queue the local panels; only the prefill waits.
    Promise.all([
      loadStatus(true),
      loadDirs(),
      loadStartSettings(),
      loadGoogleSheets(),
      loadExcelFiles(),
      loadMindnodeFiles(),
    ])
      .then(applyCurrentSessionPrefill)
      .catch(function (err) {
        console.error("Start overlay refresh failed", err);
      });
  }

  function applyCurrentSessionPrefill() {
    var s = state.statusData || {};
    // Gate on active_source, not mindnode_loaded: the two coexist, and currentSessionKey() must agree.
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
      // Seed the picker as "loaded"; restoring the worksheet stops re-confirm switching tabs.
      setTab(s.spreadsheet_type);
      selectSpreadsheet({
        type: s.spreadsheet_type,
        id_or_path: s.spreadsheet_id_or_path,
        label: s.spreadsheet_label || s.spreadsheet_id_or_path,
        worksheet: s.spreadsheet_worksheet || "",
      });
      if (s.spreadsheet_type === "google" && els.googlePaste) {
        // Not in the listed sheets: show the identifier in the paste field instead.
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
      // CLI-loaded sheet without meta: baseline as no-sheet so the dirty glow doesn't misfire.
      state.baseline.sheetTab = state.activeTab;
      state.baseline.sheetKey = "";
    } else if (s.startup_notice) {
      // Boot -s failed: never "none" (hides the CTA); highlight after setTab clears it.
      setTab(s.startup_notice_source === "excel" ? "excel" : "google");
      markSheetError(state.startupNoticeShown ? "" : s.startup_notice);
      state.startupNoticeShown = true;
    } else {
      // No sheet loaded: the session's current state is "no spreadsheet".
      setTab("none");
    }
    state.baseline = baselineFromInputs();
    applyFieldStates();
    // The name lives on the recents record; typed input beats this late prefill.
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
      clipgenStatus().then(function (s) {
        state.statusData = s;
        state.sheetLoaded = !!s.sheet_loaded;
        if (shouldAutoOpen(s)) open();
        // The server decides whether this launch is updatable and honours the cooldown.
        checkForUpdates(false);
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
    // macOS "Check for Updates…" menu item: show the About tab and force a check.
    checkForUpdates: function () {
      if (state.open) setStartTab("about"); else open("about");
      checkForUpdates(true);
    },
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
