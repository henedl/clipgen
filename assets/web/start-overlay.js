/* Start overlay — first-run / spreadsheet picker / folders / user guide.
 *
 * Loads start-overlay.html into <start-overlay-mount>, wires picker tabs +
 * recents dropdowns, calls the /api endpoints, and reloads the page after a
 * successful sheet open so all three frontends pick up the new context.
 *
 * Public API exposed on window.ClipgenStartOverlay:
 *   open()     — show the overlay
 *   close()    — hide the overlay
 *   isOpen()   — boolean
 */

(function () {
  "use strict";

  var DISMISSED_KEY = "clipgen.startOverlayDismissed";

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
      path.indexOf("/screenspace/") === 0 || path.indexOf("/transcripts/") === 0;
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
    activeTab: "google",
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

  function mount() {
    if (state.mounted) return Promise.resolve();
    var slot = document.querySelector("start-overlay-mount");
    if (!slot) return Promise.resolve();
    return fetch("start-overlay.html")
      .then(function (r) { return r.text(); })
      .then(function (html) {
        var wrap = document.createElement("div");
        wrap.innerHTML = html;
        var node = wrap.firstElementChild;
        slot.replaceWith(node);
        root = node;
        cacheEls();
        bind();
        state.mounted = true;
      })
      .catch(function (err) {
        console.error("Start overlay: failed to mount", err);
      });
  }

  function cacheEls() {
    els.statusLine = root.querySelector('[data-role="status-line"]');
    els.closeBtn = root.querySelector('[data-role="close"]');
    els.inputDir = root.querySelector("#startInputDir");
    els.outputDir = root.querySelector("#startOutputDir");
    els.inputRecentsBtn = root.querySelector('[data-role="input-recents"]');
    els.outputRecentsBtn = root.querySelector('[data-role="output-recents"]');
    els.inputRecentsList = root.querySelector('[data-role="input-recents-list"]');
    els.outputRecentsList = root.querySelector('[data-role="output-recents-list"]');
    els.tabs = root.querySelectorAll(".start-overlay-tab");
    els.googlePanel = root.querySelector('[data-tabpanel="google"]');
    els.excelPanel = root.querySelector('[data-tabpanel="excel"]');
    els.googleStatus = root.querySelector('[data-role="google-status"]');
    els.googleList = root.querySelector('[data-role="google-list"]');
    els.googlePaste = root.querySelector("#startGooglePaste");
    els.excelStatus = root.querySelector('[data-role="excel-status"]');
    els.excelList = root.querySelector('[data-role="excel-list"]');
    els.excelPaste = root.querySelector("#startExcelPaste");
    els.guideToggle = root.querySelector('[data-role="guide-toggle"]');
    els.guide = root.querySelector('[data-role="guide"]');
    els.persist = root.querySelector("#startPersistEnabled");
    els.confirmBtn = root.querySelector('[data-role="confirm"]');
    els.dismissBtn = root.querySelector('[data-role="dismiss"]');
    els.backdrop = root.querySelector('[data-role="backdrop"]');
  }

  function bind() {
    on(els.closeBtn, "click", close);
    on(els.dismissBtn, "click", close);
    on(els.confirmBtn, "click", confirm);
    on(els.backdrop, "click", function () {
      if (state.sheetLoaded) close();
    });

    Array.prototype.forEach.call(els.tabs, function (tab) {
      on(tab, "click", function () { setTab(tab.getAttribute("data-tab")); });
    });

    on(els.inputRecentsBtn, "click", function (e) {
      e.stopPropagation();
      toggleRecents("input");
    });
    on(els.outputRecentsBtn, "click", function (e) {
      e.stopPropagation();
      toggleRecents("output");
    });
    on(document, "click", function (e) {
      if (!root) return;
      if (els.inputRecentsList && !els.inputRecentsList.classList.contains("hidden")) {
        if (!els.inputRecentsList.contains(e.target) && e.target !== els.inputRecentsBtn) {
          els.inputRecentsList.classList.add("hidden");
        }
      }
      if (els.outputRecentsList && !els.outputRecentsList.classList.contains("hidden")) {
        if (!els.outputRecentsList.contains(e.target) && e.target !== els.outputRecentsBtn) {
          els.outputRecentsList.classList.add("hidden");
        }
      }
    });

    on(els.guideToggle, "click", function () {
      var open = els.guideToggle.getAttribute("aria-expanded") === "true";
      els.guideToggle.setAttribute("aria-expanded", open ? "false" : "true");
      if (open) els.guide.classList.add("hidden");
      else els.guide.classList.remove("hidden");
    });

    on(els.persist, "change", function () {
      state.persistEnabled = !!els.persist.checked;
      apiPost("/api/start-settings", { persist_enabled: state.persistEnabled })
        .catch(function (err) { console.error("Persist toggle failed", err); });
    });

    on(els.googlePaste, "input", function () {
      var v = (els.googlePaste.value || "").trim();
      if (v) {
        var label = parseLabelFromGoogleInput(v);
        setSelection({ type: "google", id_or_path: v, label: label });
      } else if (state.selection && state.selection.type === "google") {
        // Re-render to clear highlighted state if any
        renderGoogleList(state.googleSheets || []);
      }
    });

    on(els.excelPaste, "input", function () {
      var v = (els.excelPaste.value || "").trim();
      if (v) {
        setSelection({
          type: "excel",
          id_or_path: v,
          label: v.split("/").pop() || v,
        });
      } else if (state.selection && state.selection.type === "excel") {
        renderExcelList(state.excelFiles || []);
      }
    });

    on(document, "keydown", function (e) {
      if (e.key === "Escape" && state.open && state.sheetLoaded) close();
    });
  }

  function parseLabelFromGoogleInput(s) {
    if (/^https?:\/\//i.test(s)) return "Google Sheets URL";
    return s;
  }

  function toggleRecents(kind) {
    var list = kind === "input" ? els.inputRecentsList : els.outputRecentsList;
    if (!list) return;
    var hidden = list.classList.contains("hidden");
    if (els.inputRecentsList) els.inputRecentsList.classList.add("hidden");
    if (els.outputRecentsList) els.outputRecentsList.classList.add("hidden");
    if (hidden) list.classList.remove("hidden");
  }

  function setTab(name) {
    state.activeTab = name;
    Array.prototype.forEach.call(els.tabs, function (tab) {
      var active = tab.getAttribute("data-tab") === name;
      tab.classList.toggle("is-active", active);
      tab.setAttribute("aria-selected", active ? "true" : "false");
    });
    show(els.googlePanel, name === "google");
    show(els.excelPanel, name === "excel");
  }

  function setSelection(sel) {
    state.selection = sel;
    updateConfirmButton();
    if (sel && sel.type === "google") {
      highlightGoogleSelection(sel.id_or_path);
    } else if (sel && sel.type === "excel") {
      highlightExcelSelection(sel.id_or_path);
    }
  }

  function updateConfirmButton() {
    if (!els.confirmBtn) return;
    var hasSelection = !!state.selection;
    els.confirmBtn.textContent = hasSelection ? "Open" : "Save folders";
  }

  function setStatusLine(loaded, label, inputDir, outputDir) {
    if (!els.statusLine) return;
    if (loaded && label) {
      els.statusLine.textContent = "Current spreadsheet: " + label;
    } else if (loaded) {
      els.statusLine.textContent = "Spreadsheet loaded.";
    } else {
      els.statusLine.textContent = "No spreadsheet loaded.";
    }
  }

  // ---- Data loading ----

  function loadStatus() {
    return apiGet("/api/status").then(function (s) {
      state.sheetLoaded = !!s.sheet_loaded;
      setStatusLine(state.sheetLoaded, s.spreadsheet_label, s.input_dir, s.output_dir);
      setHidden(els.closeBtn, !state.sheetLoaded);
      return s;
    });
  }

  function loadDirs() {
    return apiGet("/api/dirs").then(function (d) {
      if (!d || !d.ok) return;
      els.inputDir.value = d.input || "";
      els.outputDir.value = d.output || "";
      renderRecents(els.inputRecentsList, d.recent_inputs || [], els.inputDir);
      renderRecents(els.outputRecentsList, d.recent_outputs || [], els.outputDir);
    });
  }

  function renderRecents(container, items, targetInput) {
    if (!container) return;
    container.innerHTML = "";
    if (!items.length) {
      var empty = el("div", "start-overlay-recents-empty", "No recent folders");
      container.appendChild(empty);
      return;
    }
    items.forEach(function (path) {
      var btn = el("button", "start-overlay-recents-item", path);
      btn.type = "button";
      btn.setAttribute("role", "menuitem");
      btn.addEventListener("click", function () {
        targetInput.value = path;
        container.classList.add("hidden");
      });
      container.appendChild(btn);
    });
  }

  function loadStartSettings() {
    return apiGet("/api/start-settings").then(function (r) {
      if (!r || !r.settings) return;
      state.persistEnabled = r.settings.persist_enabled !== false;
      if (els.persist) els.persist.checked = state.persistEnabled;
    });
  }

  function loadGoogleSheets() {
    if (!els.googleStatus) return Promise.resolve();
    els.googleStatus.textContent = "Checking authentication…";
    setHidden(els.googleList, true);
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
    });
  }

  function renderGoogleConnectCTA(inFlight, errorMsg) {
    els.googleStatus.innerHTML = "";
    setHidden(els.googleList, true);
    if (errorMsg) {
      var err = el("div", "start-overlay-list-empty", errorMsg);
      els.googleStatus.appendChild(err);
    }
    var msg = el(
      "div",
      "",
      inFlight
        ? "Waiting for Google sign-in… (a browser tab should have opened)"
        : "Sign in to list your Google Sheets."
    );
    msg.style.marginBottom = "8px";
    els.googleStatus.appendChild(msg);

    if (!inFlight) {
      var btn = el("button", "cg-btn cg-btn-solid cg-btn-sm", "Connect Google");
      btn.type = "button";
      btn.addEventListener("click", connectGoogle);
      els.googleStatus.appendChild(btn);
    }
  }

  function connectGoogle() {
    els.googleStatus.textContent = "Starting Google sign-in…";
    apiPost("/api/spreadsheets/google/auth", {})
      .then(function () {
        pollGoogleAuth();
      })
      .catch(function (err) {
        els.googleStatus.textContent = "Could not start sign-in: " + err.message;
      });
  }

  function pollGoogleAuth() {
    state.googlePollDeadline = Date.now() + 90 * 1000;
    function tick() {
      apiGet("/api/spreadsheets/google").then(function (g) {
        if (g.authenticated && !g.auth_error) {
          state.googleSheets = g.sheets || [];
          els.googleStatus.textContent = state.googleSheets.length
            ? state.googleSheets.length + " spreadsheets available"
            : "No spreadsheets found in your account";
          renderGoogleList(state.googleSheets);
          state.googlePollTimer = null;
          return;
        }
        if (g.auth_error) {
          renderGoogleConnectCTA(false, g.auth_error);
          state.googlePollTimer = null;
          return;
        }
        if (Date.now() > state.googlePollDeadline) {
          renderGoogleConnectCTA(false, "Sign-in timed out — try again.");
          state.googlePollTimer = null;
          return;
        }
        state.googlePollTimer = setTimeout(tick, 1500);
      }).catch(function () {
        state.googlePollTimer = setTimeout(tick, 2500);
      });
    }
    tick();
  }

  function renderGoogleList(sheets) {
    setHidden(els.googleList, sheets.length === 0);
    els.googleList.innerHTML = "";
    sheets.forEach(function (sheet) {
      var item = el("button", "start-overlay-list-item", sheet.name);
      item.type = "button";
      item.setAttribute("data-id", sheet.id);
      item.addEventListener("click", function () {
        els.googlePaste.value = "";
        setSelection({ type: "google", id_or_path: sheet.name, label: sheet.name });
      });
      els.googleList.appendChild(item);
    });
    if (state.selection && state.selection.type === "google") {
      highlightGoogleSelection(state.selection.id_or_path);
    }
  }

  function highlightGoogleSelection(idOrPath) {
    if (!els.googleList) return;
    var items = els.googleList.querySelectorAll(".start-overlay-list-item");
    Array.prototype.forEach.call(items, function (item) {
      item.classList.toggle("is-selected", item.getAttribute("data-id") === idOrPath);
    });
  }

  function loadExcelFiles() {
    if (!els.excelStatus) return Promise.resolve();
    els.excelStatus.textContent = "Scanning input folder…";
    els.excelList.innerHTML = "";
    return apiGet("/api/spreadsheets/excel").then(function (r) {
      state.excelFiles = r.files || [];
      els.excelStatus.textContent = state.excelFiles.length
        ? state.excelFiles.length + " .xlsx in " + r.input_dir
        : "No .xlsx files in " + r.input_dir;
      renderExcelList(state.excelFiles);
    });
  }

  function renderExcelList(files) {
    els.excelList.innerHTML = "";
    if (!files.length) {
      els.excelList.appendChild(
        el("div", "start-overlay-list-empty", "Paste a path below to load any .xlsx.")
      );
      return;
    }
    files.forEach(function (f) {
      var item = el("button", "start-overlay-list-item", "");
      item.type = "button";
      item.setAttribute("data-path", f.path);
      var name = el("span", "", f.name);
      var sub = el("span", "start-overlay-list-item-sub", f.path);
      item.appendChild(name);
      item.appendChild(sub);
      item.addEventListener("click", function () {
        els.excelPaste.value = "";
        setSelection({ type: "excel", id_or_path: f.path, label: f.name });
      });
      els.excelList.appendChild(item);
    });
    if (state.selection && state.selection.type === "excel") {
      highlightExcelSelection(state.selection.id_or_path);
    }
  }

  function highlightExcelSelection(path) {
    if (!els.excelList) return;
    var items = els.excelList.querySelectorAll(".start-overlay-list-item");
    Array.prototype.forEach.call(items, function (item) {
      item.classList.toggle("is-selected", item.getAttribute("data-path") === path);
    });
  }

  // ---- Open / dismiss flows ----

  function confirm() {
    var inputVal = (els.inputDir.value || "").trim();
    var outputVal = (els.outputDir.value || "").trim();

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

    els.confirmBtn.disabled = true;
    dirsPromise.then(function (res) {
      if (!res.ok) {
        var msg = "Folder error";
        if (res.body && res.body.errors) {
          var parts = [];
          if (res.body.errors.input) parts.push(res.body.errors.input);
          if (res.body.errors.output) parts.push(res.body.errors.output);
          msg = parts.join(" / ") || msg;
        }
        showToast(msg);
        els.confirmBtn.disabled = false;
        return;
      }
      if (!state.selection) {
        // Only dirs were changed — just close.
        showToast("Folders saved.");
        els.confirmBtn.disabled = false;
        close();
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
        .then(function (res) {
          els.confirmBtn.disabled = false;
          if (!res.ok || !res.body.ok) {
            showToast((res.body && res.body.error) || "Could not open spreadsheet");
            return;
          }
          // Hard reload so all three frontends re-fetch their data.
          window.location.reload();
        })
        .catch(function (err) {
          els.confirmBtn.disabled = false;
          showToast("Open failed: " + err.message);
        });
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
    refresh();
  }

  function close() {
    if (!root) return;
    state.open = false;
    show(root, false);
    markDismissed();
    if (state.googlePollTimer) {
      clearTimeout(state.googlePollTimer);
      state.googlePollTimer = null;
    }
  }

  function refresh() {
    loadStatus()
      .then(loadDirs)
      .then(loadStartSettings)
      .then(loadGoogleSheets)
      .then(loadExcelFiles)
      .then(updateConfirmButton)
      .catch(function (err) {
        console.error("Start overlay refresh failed", err);
      });
  }

  // ---- Boot ----

  function boot() {
    mount().then(function () {
      // After mount, check status once to decide auto-open.
      apiGet("/api/status").then(function (s) {
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
