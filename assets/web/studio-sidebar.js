/* clipgen Studio sidebar satellite — studio-sidebar.js
 *
 * The sheet sidebar carved out of studio.js: participant / category /
 * keyword / severity filters, the view switcher, the collapse toggle, and
 * their persistence. Loads right after the hub and destructures its grid
 * helpers at load; the hub keeps same-named delegators for its own call
 * sites. Function bodies are unchanged from the hub.
 */
(function () {
  "use strict";

  var STUDIO = window.ClipgenStudio;
  var state = STUDIO.state;
  var applyGridFilters = STUDIO.applyGridFilters,
    hasSeverityData = STUDIO.hasSeverityData,
    renderGrid = STUDIO.renderGrid;


  var SIDEBAR_VIEW_KEY = "clipgen-studio-sidebar-open";
  // VIEWS derive severity allowlists from CLIPGEN_CONFIG rank (<= -2, >= 1), not labels.
  var _severityLabelsWhere = function (predicate) {
    return CLIPGEN_CONFIG.severity
      .filter(function (s) { return predicate(s.rank); })
      .map(function (s) { return s.label; });
  };
  var SIDEBAR_VIEWS = [
    { id: "all", label: "All" },
    {
      id: "highlights",
      label: "Highlights",
      severities: _severityLabelsWhere(function (r) { return r <= -2; }),
    },
    {
      id: "positive",
      label: "Positive",
      severities: _severityLabelsWhere(function (r) { return r >= 1; }),
    },
  ];

  function readPersistedSidebarOpen() {
    try {
      var stored = localStorage.getItem(SIDEBAR_VIEW_KEY);
      if (stored !== null) state.sidebarOpen = (stored !== "false");
    } catch (_) {}
    // Set data-open before first paint, or the collapse animates on entry (.tx-ready gates it).
    var sidebar = document.getElementById("studioSidebar");
    if (!sidebar) return;
    sidebar.setAttribute("data-open", state.sidebarOpen ? "true" : "false");
    requestAnimationFrame(function () {
      requestAnimationFrame(function () {
        sidebar.classList.add("tx-ready");
      });
    });
  }

  // Runs at script load: DOMContentLoaded can land after first paint.
  readPersistedSidebarOpen();

  function persistSidebarOpen() {
    try { localStorage.setItem(SIDEBAR_VIEW_KEY, state.sidebarOpen ? "true" : "false"); } catch (_) {}
  }

  function applySidebarView(viewId) {
    var view = null;
    for (var i = 0; i < SIDEBAR_VIEWS.length; i++) {
      if (SIDEBAR_VIEWS[i].id === viewId) { view = SIDEBAR_VIEWS[i]; break; }
    }
    state.filters.severities = (view && view.severities) ? view.severities.slice() : [];
  }

  function applySidebarCategories() {
    state.filters.categories = Object.keys(state.sidebarCategories).filter(function (k) {
      return !!state.sidebarCategories[k];
    });
  }

  function applySidebarKeywords() {
    state.filters.keywords = Object.keys(state.sidebarKeywords).filter(function (k) {
      return !!state.sidebarKeywords[k];
    });
  }

  // Raw maps round-trip; filters.categories/keywords re-derive on restore via applySidebar*.
  function persistSidebarFilters() {
    setStoredUIStateField("studio", "filters", {
      severities: state.filters.severities,
      categories: state.sidebarCategories,
      keywords: state.sidebarKeywords,
      participants: state.sidebarParticipants,
      activeFunction: state.activeFunction,
      fnMin: state.filters.fnMin,
      fnMax: state.filters.fnMax,
    });
  }

  function restoreSidebarFilters() {
    var stored = getStoredUIState("studio").filters;
    if (!stored) return;
    if (stored.severities) state.filters.severities = stored.severities.slice();
    if (stored.categories) state.sidebarCategories = stored.categories;
    if (stored.keywords) state.sidebarKeywords = stored.keywords;
    if (stored.participants) state.sidebarParticipants = stored.participants;
    if (stored.activeFunction) state.activeFunction = stored.activeFunction;
    if (stored.fnMin != null) state.filters.fnMin = stored.fnMin;
    if (stored.fnMax != null) state.filters.fnMax = stored.fnMax;
    applySidebarCategories();
    applySidebarKeywords();
  }

  function keywordLabel(annotationId) {
    if (!annotationId) return "";
    return annotationId.charAt(0).toUpperCase() + annotationId.slice(1);
  }

  function countSidebarSelectedParticipants() {
    var n = 0;
    for (var k in state.sidebarParticipants) {
      if (Object.prototype.hasOwnProperty.call(state.sidebarParticipants, k) && state.sidebarParticipants[k]) n++;
    }
    return n;
  }

  function renderSidebar() {
    var sidebar = qs("#studioSidebar");
    if (!sidebar) return;
    sidebar.setAttribute("data-open", state.sidebarOpen ? "true" : "false");

    if (!state.sheetData) return;
    var d = state.sheetData;

    // Counts per severity / category / participant / keyword
    var sevCounts = { all: d.rows.length };
    var catCounts = {};
    var kwCounts = {};
    var partCounts = {};
    var participants = d.participants || [];
    for (var p = 0; p < participants.length; p++) partCounts[participants[p]] = 0;
    for (var i = 0; i < d.rows.length; i++) {
      var row = d.rows[i];
      var sev = (row.severity || "").trim();
      sevCounts[sev] = (sevCounts[sev] || 0) + 1;
      if (row.category) catCounts[row.category] = (catCounts[row.category] || 0) + 1;
      if (row.keywords) {
        for (var rk = 0; rk < row.keywords.length; rk++) {
          var kid = row.keywords[rk];
          kwCounts[kid] = (kwCounts[kid] || 0) + 1;
        }
      }
      for (var j = 0; j < participants.length; j++) {
        var c = row.cells[participants[j]];
        if (c && c.valid) partCounts[participants[j]] += 1;
      }
    }
    function countRowsBySeverities(severities) {
      if (!severities || severities.length === 0) return d.rows.length;
      var n = 0;
      for (var k = 0; k < d.rows.length; k++) {
        var sev = (d.rows[k].severity || "").trim();
        if (sev && severities.indexOf(sev) >= 0) n++;
      }
      return n;
    }

    function severitiesEqual(a, b) {
      if (a.length !== b.length) return false;
      for (var i = 0; i < a.length; i++) {
        if (b.indexOf(a[i]) < 0) return false;
      }
      return true;
    }

    // VIEWS rows; active derives from state.filters.severities so pills re-highlight views.
    var viewsBody = sidebar.querySelector('[data-target="views"]');
    if (viewsBody) {
      viewsBody.innerHTML = "";
      SIDEBAR_VIEWS.forEach(function (view) {
        var viewSevs = view.severities || [];
        var count = view.id === "all" ? d.rows.length : countRowsBySeverities(viewSevs);
        viewsBody.appendChild(createSidebarRow({
          label: view.label,
          count: count,
          active: severitiesEqual(viewSevs, state.filters.severities),
          onClick: function () {
            applySidebarView(view.id);
            persistSidebarFilters();
            renderSidebar();
            renderGrid();
          },
        }));
      });
    }

    // CATEGORIES — vertical-list rows with counts and category-hue dots.
    var catsBody = sidebar.querySelector('[data-target="categories"]');
    if (catsBody) {
      catsBody.innerHTML = "";
      var cats = Object.keys(catCounts).sort();
      cats.forEach(function (cat) {
        catsBody.appendChild(createSidebarRow({
          label: cat,
          count: catCounts[cat],
          active: !!state.sidebarCategories[cat],
          dotColor: "oklch(0.7 0.16 " + categoryHue(cat) + ")",
          onClick: function () {
            state.sidebarCategories[cat] = !state.sidebarCategories[cat];
            applySidebarCategories();
            persistSidebarFilters();
            renderSidebar();
            renderGrid();
          },
        }));
      });
      if (cats.length === 0) {
        var empty = el("span", "studio-sidebar-row-label", "(no categories)");
        empty.style.padding = "6px 16px";
        empty.style.color = "var(--fg-faint)";
        catsBody.appendChild(empty);
      }
    }

    // SEVERITY pills: "Any severity" clears; each pill toggles its label.
    var sevBody = sidebar.querySelector('[data-target="severity"]');
    if (sevBody) {
      sevBody.innerHTML = "";
      if (!hasSeverityData(d.rows)) {
        sevBody.appendChild(makeSidebarEmpty("(no severity data)"));
      } else {
        sevBody.appendChild(createSidebarRow({
          label: "Any severity",
          count: d.rows.length,
          active: state.filters.severities.length === 0,
          onClick: function () {
            state.filters.severities = [];
            persistSidebarFilters();
            renderSidebar();
            renderGrid();
          },
        }));
        for (var si = 0; si < CLIPGEN_CONFIG.severity.length; si++) {
          var sevLabel = CLIPGEN_CONFIG.severity[si].label;
          var sevCount = sevCounts[sevLabel] || 0;
          if (sevCount === 0) continue;
          (function (label) {
            sevBody.appendChild(createSidebarRow({
              label: label,
              count: sevCount,
              active: state.filters.severities.indexOf(label) >= 0,
              dotClass: severityClass(label),
              onClick: function () {
                var arr = state.filters.severities.slice();
                var idx = arr.indexOf(label);
                if (idx >= 0) arr.splice(idx, 1); else arr.push(label);
                state.filters.severities = arr;
                persistSidebarFilters();
                renderSidebar();
                renderGrid();
              },
            }));
          })(sevLabel);
        }
      }
    }

    // KEYWORDS: filter is row-level; cell-level emphasis happens in grid render.
    var kwBody = sidebar.querySelector('[data-target="keywords"]');
    if (kwBody) {
      kwBody.innerHTML = "";
      var annotations = (CLIPGEN_CONFIG && CLIPGEN_CONFIG.annotations) || [];
      var anyKw = false;
      for (var ak = 0; ak < annotations.length; ak++) {
        if (kwCounts[annotations[ak].id]) { anyKw = true; break; }
      }
      if (annotations.length === 0 || !anyKw) {
        kwBody.appendChild(makeSidebarEmpty("(no keywords)"));
      } else {
        annotations.forEach(function (ann) {
          var count = kwCounts[ann.id] || 0;
          if (count === 0) return;
          kwBody.appendChild(createSidebarRow({
            label: keywordLabel(ann.id),
            count: count,
            active: !!state.sidebarKeywords[ann.id],
            onClick: function () {
              state.sidebarKeywords[ann.id] = !state.sidebarKeywords[ann.id];
              applySidebarKeywords();
              persistSidebarFilters();
              renderSidebar();
              renderGrid();
            },
          }));
        });
      }
    }

    // FUNCTION — min/max numeric inputs gated on the activeFunction picker
    // in the table header.
    var fnBody = sidebar.querySelector('[data-target="function"]');
    if (fnBody) {
      fnBody.innerHTML = "";
      fnBody.appendChild(buildSidebarFunctionRange());
    }

    // PARTICIPANTS — compact 6-col grid of mono pills.
    var partsBody = sidebar.querySelector('[data-target="participants"]');
    if (partsBody) {
      partsBody.innerHTML = "";
      participants.forEach(function (pid) {
        var pill = el("button", "studio-sidebar-pill cg-mono", pid);
        pill.type = "button";
        if (state.sidebarParticipants[pid]) pill.classList.add("is-active");
        pill.addEventListener("click", function () {
          state.sidebarParticipants[pid] = !state.sidebarParticipants[pid];
          persistSidebarFilters();
          renderSidebar();
          renderGrid();
        });
        partsBody.appendChild(pill);
      });
    }
  }

  function buildSidebarFunctionRange() {
    var row = el("div", "studio-sidebar-range");
    var minIn = document.createElement("input");
    minIn.type = "number";
    minIn.id = "sidebarFnMin";
    minIn.placeholder = "Min";
    minIn.autocomplete = "off";
    minIn.disabled = !state.activeFunction;
    if (state.filters.fnMin !== null) minIn.value = String(state.filters.fnMin);

    var maxIn = document.createElement("input");
    maxIn.type = "number";
    maxIn.id = "sidebarFnMax";
    maxIn.placeholder = "Max";
    maxIn.autocomplete = "off";
    maxIn.disabled = !state.activeFunction;
    if (state.filters.fnMax !== null) maxIn.value = String(state.filters.fnMax);

    function onChange() {
      var mn = minIn.value.trim();
      var mx = maxIn.value.trim();
      state.filters.fnMin = mn !== "" ? parseFloat(mn) : null;
      state.filters.fnMax = mx !== "" ? parseFloat(mx) : null;
      persistSidebarFilters();
      applyGridFilters();
    }
    minIn.addEventListener("input", onChange);
    maxIn.addEventListener("input", onChange);
    row.appendChild(minIn);
    row.appendChild(el("span", "studio-sidebar-range-sep", "to"));
    row.appendChild(maxIn);
    return row;
  }

  function makeSidebarEmpty(text) {
    var span = el("span", "studio-sidebar-row-label", text);
    span.style.padding = "6px 16px";
    span.style.color = "var(--fg-faint)";
    return span;
  }

  function createSidebarRow(opts) {
    var row = el("button", "studio-sidebar-row");
    row.type = "button";
    if (opts.active) row.classList.add("is-active");
    if (opts.dotColor || opts.dotClass) {
      var dot = el("span", "studio-sidebar-row-dot");
      if (opts.dotClass) dot.classList.add(opts.dotClass);
      if (opts.dotColor) dot.style.background = opts.dotColor;
      row.appendChild(dot);
    }
    var label = el("span", "studio-sidebar-row-label");
    label.textContent = opts.label || "";
    label.title = opts.label || "";
    row.appendChild(label);
    if (opts.count != null) {
      var count = el("span", "studio-sidebar-row-count cg-mono");
      count.textContent = String(opts.count);
      row.appendChild(count);
    }
    if (typeof opts.onClick === "function") {
      row.addEventListener("click", opts.onClick);
    }
    return row;
  }

  function toggleSidebar() {
    state.sidebarOpen = !state.sidebarOpen;
    persistSidebarOpen();
    renderSidebar();
  }

  function bindSidebarToggle() {
    var btn = qs("#studioSidebarToggle");
    if (!btn) return;
    btn.addEventListener("click", toggleSidebar);
  }

  function isParticipantHidden(pid) {
    if (countSidebarSelectedParticipants() === 0) return false;
    return !state.sidebarParticipants[pid];
  }

  function setActiveTabAttr(tab) {
    document.body.setAttribute("data-active-tab", tab);
  }

  function syncFilterFnDisabled() {
    var fnMin = qs("#sidebarFnMin");
    var fnMax = qs("#sidebarFnMax");
    var enabled = !!state.activeFunction;
    if (fnMin) fnMin.disabled = !enabled;
    if (fnMax) fnMax.disabled = !enabled;
    if (!enabled && (state.filters.fnMin !== null || state.filters.fnMax !== null)) {
      state.filters.fnMin = null;
      state.filters.fnMax = null;
      if (fnMin) fnMin.value = "";
      if (fnMax) fnMax.value = "";
    }
  }

  STUDIO.applySidebarView = applySidebarView;
  STUDIO.bindSidebarToggle = bindSidebarToggle;
  STUDIO.isParticipantHidden = isParticipantHidden;
  STUDIO.persistSidebarFilters = persistSidebarFilters;
  STUDIO.renderSidebar = renderSidebar;
  STUDIO.restoreSidebarFilters = restoreSidebarFilters;
  STUDIO.setActiveTabAttr = setActiveTabAttr;
  STUDIO.syncFilterFnDisabled = syncFilterFnDisabled;
  STUDIO.toggleSidebar = toggleSidebar;
})();
