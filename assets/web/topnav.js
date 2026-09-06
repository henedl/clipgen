/* Unified top navigation — shared chrome across Studio, Screenspace, Transcripts.
 *
 * Mounts on DOMContentLoaded into a <topnav-mount data-frontend="..."> element.
 * Reads page-specific Quick Actions from window.CLIPGEN_QUICK_ACTIONS, an
 * array of { icon, label, action, disabled?, title? } | { divider: true } |
 * { header } items. Items with disabled=true render grayed and ignore clicks;
 * the title field becomes a hover tooltip explaining why (rendered through the
 * [data-tooltip] singleton in utils.js, not the native attribute).
 * Pages may also call ClipgenTopNav.setQuickActions(items) post-mount to
 * update the menu as state changes, and ClipgenTopNav.onBeforeOpen(cb) to
 * refresh state right before the menu opens.
 *
 * Wires existing #themeToggle / #logBtn / #settingsBtn IDs inside the new
 * cluster so page setup code that addEventListener's to those IDs continues
 * to work — the IDs just live inside the TopNav DOM now.
 */

(function () {
  "use strict";

  var SURFACES = [
    { id: "studio", label: "Studio", href: "/studio/" },
    { id: "screenspace", label: "Screenspace", href: "/screenspace/" },
    { id: "transcripts", label: "Transcripts", href: "/transcripts/" },
    { id: "workflows", label: "Workflows", href: "/workflows/" },
    { id: "composer", label: "Composer", href: "/composer/" },
    { id: "overview", label: "Overview", href: "/overview/" },
  ];

  var state = {
    quickActions: [],
    quickActionsOpen: false,
    activeFrontend: null,
  };

  var els = {};
  var readyCallbacks = [];
  var beforeOpenCallbacks = [];
  var isReady = false;

  function build() {
    var mount = document.querySelector("topnav-mount");
    if (!mount) return;

    state.activeFrontend = mount.getAttribute("data-frontend") || null;

    var nav = document.createElement("nav");
    nav.className = "topnav";
    nav.setAttribute("aria-label", "Main");

    nav.appendChild(buildLeft());
    nav.appendChild(buildCenter());
    nav.appendChild(buildRight());

    // Desktop chrome hides the title bar; pywebview drags .pywebview-drag-region direct targets. See desktop.py.
    if (document.documentElement.dataset.desktopChrome) {
      nav.classList.add("pywebview-drag-region");
      var columns = nav.children;
      for (var c = 0; c < columns.length; c++) {
        columns[c].classList.add("pywebview-drag-region");
      }
      // Restore the title bar's double-click, gated on the same direct-target rule as dragging.
      nav.addEventListener("dblclick", function (e) {
        var cls = e.target && e.target.classList;
        if (!cls || !cls.contains("pywebview-drag-region")) return;
        var api = window.pywebview && window.pywebview.api;
        if (api && typeof api.titlebar_double_click === "function") {
          api.titlebar_double_click();
        }
      });
    }

    mount.replaceWith(nav);
    els.nav = nav;
    els.tabs = nav.querySelectorAll(".topnav-tab");
    els.qaWrap = nav.querySelector(".topnav-qa");
    els.qaTrigger = nav.querySelector(".topnav-qa-trigger");
    els.qaPanel = nav.querySelector(".topnav-qa-panel");

    bindEvents();
    setQuickActions(window.CLIPGEN_QUICK_ACTIONS || []);

    isReady = true;
    for (var i = 0; i < readyCallbacks.length; i++) {
      try { readyCallbacks[i](); } catch (_) {}
    }
    readyCallbacks = [];
  }

  function buildLeft() {
    var left = document.createElement("div");
    left.className = "topnav-left";
    var brand = document.createElement("a");
    brand.className = "topnav-brand";
    brand.href = "/studio/";
    brand.innerHTML = '<span class="brand-mark" aria-hidden="true"></span><span>clipgen</span>';
    left.appendChild(brand);
    return left;
  }

  function buildCenter() {
    var center = document.createElement("div");
    center.className = "topnav-center";
    SURFACES.forEach(function (s) {
      var tab = document.createElement("a");
      tab.className = "topnav-tab" + (s.id === state.activeFrontend ? " is-active" : "");
      tab.href = s.href;
      tab.textContent = s.label;
      tab.setAttribute("data-frontend", s.id);
      if (s.id === state.activeFrontend) {
        tab.setAttribute("aria-current", "page");
      }
      center.appendChild(tab);
    });
    return center;
  }

  function buildRight() {
    var right = document.createElement("div");
    right.className = "topnav-right";

    // Quick Actions
    var qaWrap = document.createElement("div");
    qaWrap.className = "topnav-qa";
    var qaTrigger = document.createElement("button");
    qaTrigger.type = "button";
    qaTrigger.className = "topnav-qa-trigger";
    qaTrigger.setAttribute("aria-haspopup", "menu");
    qaTrigger.setAttribute("aria-expanded", "false");
    qaTrigger.innerHTML = '<span>Quick actions</span><span class="topnav-icon topnav-qa-caret" aria-hidden="true"></span>';
    var qaPanel = document.createElement("div");
    qaPanel.className = "topnav-qa-panel cg-menu";
    qaPanel.setAttribute("role", "menu");
    qaWrap.appendChild(qaTrigger);
    qaWrap.appendChild(qaPanel);
    right.appendChild(qaWrap);

    right.appendChild(makeDivider());

    // Start overlay opener — keeps the first-run picker reachable after dismissal.
    var startBtn = document.createElement("button");
    startBtn.type = "button";
    startBtn.id = "startBtn";
    startBtn.className = "topnav-icon-btn";
    startBtn.setAttribute("data-tooltip", "Start panel");
    startBtn.setAttribute("aria-label", "Start panel");
    var startIcon = document.createElement("span");
    startIcon.className = "topnav-icon";
    startIcon.style.cssText = iconMaskStyle("home");
    startBtn.appendChild(startIcon);
    right.appendChild(startBtn);

    // Log button only where an artifact log exists (studio.js #logOverlay, composer.js log panel).
    if (state.activeFrontend === "studio" || state.activeFrontend === "composer") {
      var logBtn = document.createElement("button");
      logBtn.type = "button";
      logBtn.id = "logBtn";
      logBtn.className = "topnav-icon-btn";
      logBtn.setAttribute("data-tooltip", "Artifact Log");
      logBtn.setAttribute("aria-label", "Artifact Log");
      var logIcon = document.createElement("span");
      logIcon.className = "topnav-icon";
      logIcon.style.cssText = iconMaskStyle("list-bullet");
      logBtn.appendChild(logIcon);
      right.appendChild(logBtn);
    }

    // Settings button — keeps existing #settingsBtn id.
    var settingsBtn = document.createElement("button");
    settingsBtn.type = "button";
    settingsBtn.id = "settingsBtn";
    settingsBtn.className = "topnav-icon-btn";
    settingsBtn.setAttribute("data-tooltip", "Settings");
    settingsBtn.setAttribute("aria-label", "Settings");
    var settingsIcon = document.createElement("span");
    settingsIcon.className = "topnav-icon";
    settingsIcon.style.cssText = iconMaskStyle("cog-6-tooth");
    settingsBtn.appendChild(settingsIcon);
    right.appendChild(settingsBtn);

    // Theme toggle. Keeps #themeToggle and .theme-toggle-icon for initThemeToggle() in utils.js.
    var themeBtn = document.createElement("button");
    themeBtn.type = "button";
    themeBtn.id = "themeToggle";
    themeBtn.setAttribute("aria-label", "Toggle dark mode");
    themeBtn.setAttribute("aria-pressed", "false");
    themeBtn.innerHTML = '<span class="theme-toggle-icon theme-icon-sun"></span><span class="theme-toggle-icon theme-icon-moon"></span>';
    right.appendChild(themeBtn);

    return right;
  }

  function makeDivider() {
    var d = document.createElement("span");
    d.className = "topnav-divider";
    d.setAttribute("aria-hidden", "true");
    return d;
  }

  function bindEvents() {
    els.qaTrigger.addEventListener("click", function (e) {
      e.stopPropagation();
      toggleQuickActions();
    });

    document.addEventListener("click", function (e) {
      if (!state.quickActionsOpen) return;
      if (els.qaWrap.contains(e.target)) return;
      closeQuickActions();
    });

    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && state.quickActionsOpen) {
        closeQuickActions();
      }
    });
  }

  function toggleQuickActions() {
    if (state.quickActionsOpen) closeQuickActions();
    else openQuickActions();
  }

  function openQuickActions() {
    for (var i = 0; i < beforeOpenCallbacks.length; i++) {
      try { beforeOpenCallbacks[i](); } catch (_) {}
    }
    state.quickActionsOpen = true;
    els.qaTrigger.classList.add("is-open");
    els.qaTrigger.setAttribute("aria-expanded", "true");
    els.qaPanel.classList.add("is-open");
  }

  function closeQuickActions() {
    state.quickActionsOpen = false;
    els.qaTrigger.classList.remove("is-open");
    els.qaTrigger.setAttribute("aria-expanded", "false");
    els.qaPanel.classList.remove("is-open");
  }

  function getQuickActions(opts) {
    // refresh:true re-runs the menu's open-time gating so the palette sees the same snapshot.
    if (opts && opts.refresh) {
      for (var i = 0; i < beforeOpenCallbacks.length; i++) {
        try { beforeOpenCallbacks[i](); } catch (_) {}
      }
    }
    return state.quickActions.slice();
  }

  function setQuickActions(items) {
    state.quickActions = Array.isArray(items) ? items : [];
    if (!els.qaPanel) return;
    els.qaPanel.innerHTML = "";
    state.quickActions.forEach(function (item) {
      els.qaPanel.appendChild(buildQuickActionItem(item));
    });
    // Hide the trigger when the menu is empty.
    els.qaWrap.style.display = state.quickActions.length === 0 ? "none" : "";
  }

  function buildQuickActionItem(item) {
    if (item.divider) {
      var d = document.createElement("div");
      d.className = "topnav-qa-divider";
      d.setAttribute("role", "separator");
      return d;
    }
    if (item.header) {
      var h = document.createElement("div");
      h.className = "topnav-qa-header";
      h.textContent = item.header;
      return h;
    }
    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "topnav-qa-item";
    btn.setAttribute("role", "menuitem");
    if (item.disabled) {
      btn.disabled = true;
      btn.classList.add("is-disabled");
      btn.setAttribute("aria-disabled", "true");
    }
    if (item.title) {
      btn.setAttribute("data-tooltip", item.title);
    }
    if (item.icon) {
      var ic = document.createElement("span");
      ic.className = "topnav-icon";
      ic.style.cssText = iconMaskStyle(item.icon);
      btn.appendChild(ic);
    }
    var label = document.createElement("span");
    label.textContent = item.label || "";
    btn.appendChild(label);
    btn.addEventListener("click", function () {
      if (item.disabled) return;
      closeQuickActions();
      if (typeof item.action === "function") {
        try { item.action(); } catch (e) { console.error("Quick action error:", e); }
      }
    });
    return btn;
  }

  function onReady(cb) {
    if (typeof cb !== "function") return;
    if (isReady) {
      try { cb(); } catch (_) {}
    } else {
      readyCallbacks.push(cb);
    }
  }

  function onBeforeOpen(cb) {
    if (typeof cb !== "function") return;
    beforeOpenCallbacks.push(cb);
  }

  // Build quick actions now, on export-status flips, and (opt-in) on open.
  function installQuickActions(build, opts) {
    opts = opts || {};
    var exportActions = window.ClipgenExportActions;
    function rebuild() {
      setQuickActions(build());
    }
    rebuild();
    if (exportActions) exportActions.refreshExportStatus(rebuild);
    onBeforeOpen(function () {
      if (opts.rebuildOnOpen) rebuild();
      if (exportActions) exportActions.refreshExportStatus(rebuild);
    });
    return rebuild;
  }

  // "Update available" dot on the Start button; start-overlay.js drives it.
  function setStartBadge(on) {
    var btn = document.getElementById("startBtn");
    if (!btn) return;
    btn.classList.toggle("has-badge", !!on);
    btn.setAttribute("data-tooltip", on ? "Update available" : "Start panel");
    btn.setAttribute("aria-label", on ? "Start panel, update available" : "Start panel");
  }

  window.ClipgenTopNav = {
    setStartBadge: setStartBadge,
    setQuickActions: setQuickActions,
    getQuickActions: getQuickActions,
    installQuickActions: installQuickActions,
    onReady: onReady,
    onBeforeOpen: onBeforeOpen,
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", build);
  } else {
    build();
  }
})();
