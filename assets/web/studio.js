/* clipgen Studio page.
 *
 * Main hub: spreadsheet grid on top, artifact + reel queues + generated
 * outputs below, with intake panels that surface Screenspace and Transcript
 * activity for cross-referencing.
 *
 * Key shapes on `state`:
 *   sheetData                  — rows/cells from the loaded spreadsheet.
 *   artifactQueue / reelQueue  — pending generation work; persisted to
 *                                sessionStorage under QUEUE_STORAGE_KEY.
 *   generatedArtifacts         — completed outputs surfaced to the UI.
 *   cellResults                — per-cell success/error status overlaid
 *                                onto the sheet grid (keyed by cellKey()).
 *   intakeEvents / intakeClusters / intakeSeenIds — Screenspace polling
 *                                snapshot; trIntakeMarks/Clusters mirror
 *                                this for Transcripts.
 */

(function () {
  "use strict";

  var QUEUE_STORAGE_KEY = "clipgen-studio-queues";

  // Build a mask-image icon span as an HTML string. Sizing comes from a
  // parent rule (e.g. .cg-btn-icon) or from extraClass. See
  // .cg-icon family in studio.css.
  function iconHTML(name, extraClass) {
    return '<span class="cg-icon cg-icon--' + name + (extraClass ? " " + extraClass : "") + '"></span>';
  }

  var state = {
    sheetData: null,
    artifactQueue: [],
    reelQueue: [],
    generatedArtifacts: [],
    generating: false,
    cellResults: {},
    stashes: [],
    artifactStashes: [],
    settingsData: null,
    bottomCollapsed: false,
    activeFunction: "",
    cellExpandHover: true,
    filters: {
      categories: [],
      severities: [],
      keywords: [],
      fnMin: null,
      fnMax: null,
    },
    intakeEvents: [],
    intakeClusters: [],
    intakeSeenIds: {},
    intakePollTimer: null,
    intakeFilterText: "",
    intakeFilterDetector: "",
    intakeFilterNew: false,
    intakeFilterParticipants: [],
    intakeHoveredIdx: -1,
    activePreviewTab: "sheet",
    trIntakeMarks: [],
    trIntakeClusters: [],
    trIntakePollTimer: null,
    intakeStatusTimer: null,
    trIntakeStatusTimer: null,
    trIntakeFilterCategory: "",
    trIntakeFilterParticipants: [],
    trIntakeFilterText: "",
    trIntakeShowAll: false,
    trIntakeHoveredIdx: -1,
    trIntakeTooltipsEnabled: true,
    convergenceBaselines: {},
    convergenceDataVersion: 0,
    convergenceStale: false,
    sidebarOpen: true,
    sidebarCategories: {},
    sidebarKeywords: {},
    sidebarParticipants: {},
  };

  function isIntakeSource(source) {
    return source === "screenspace" || source === "transcript";
  }

  var ROW_FUNCTIONS = {
    Count: function (row, participants) {
      var total = 0;
      for (var j = 0; j < participants.length; j++) {
        var c = row.cells[participants[j]];
        if (c && c.valid) total += parseClipTimestamps(c.value, participants[j]).length;
      }
      return total;
    },
    Unique: function (row, participants) {
      var count = 0;
      for (var j = 0; j < participants.length; j++) {
        var c = row.cells[participants[j]];
        if (c && c.valid) count++;
      }
      return count;
    },
  };

  // ---- Helpers ----

  function cellKey(participant, rowNum) {
    return participant + "." + rowNum;
  }

  function findInQueue(queue, participant, rowNum) {
    var key = cellKey(participant, rowNum);
    for (var i = 0; i < queue.length; i++) {
      if (cellKey(queue[i].participant, queue[i].row) === key) return i;
    }
    return -1;
  }

  function hasSegmentInQueue(queue, participant, rowNum, segIdx) {
    var key = cellKey(participant, rowNum);
    for (var i = 0; i < queue.length; i++) {
      if (cellKey(queue[i].participant, queue[i].row) === key && queue[i].segIdx === segIdx) return true;
    }
    return false;
  }

  function removeAllCellEntries(queue, participant, rowNum) {
    var key = cellKey(participant, rowNum);
    for (var i = queue.length - 1; i >= 0; i--) {
      if (cellKey(queue[i].participant, queue[i].row) === key) queue.splice(i, 1);
    }
  }

  function removeSegmentEntry(queue, participant, rowNum, segIdx) {
    var key = cellKey(participant, rowNum);
    for (var i = 0; i < queue.length; i++) {
      if (cellKey(queue[i].participant, queue[i].row) === key && queue[i].segIdx === segIdx) {
        queue.splice(i, 1);
        return;
      }
    }
  }

  function expandCellToSegments(info) {
    var segments = parseClipTimestamps(info.timestamp, info.participant);
    var entries = [];
    for (var i = 0; i < segments.length; i++) {
      entries.push({
        participant: info.participant,
        row: info.row,
        desc: info.desc,
        timestamp: info.timestamp,
        segIdx: i,
        start: segments[i].startSeconds,
        end: segments[i].startSeconds + segments[i].duration,
        segTotal: segments.length,
      });
    }
    return entries;
  }

  function clearGridHighlights() {
    var els = qsa(".header-highlight");
    for (var i = 0; i < els.length; i++) els[i].classList.remove("header-highlight");
  }

  function highlightGridHeaders(participant, row) {
    clearGridHighlights();
    var th = qs('#sheetGrid thead th[data-participant="' + participant + '"]');
    if (th) th.classList.add("header-highlight");
    var td = qs('#sheetGrid tbody td[data-select-row="' + row + '"]');
    if (td) td.classList.add("header-highlight");
  }

  function parseClipTimestamps(raw, participantId) {
    var DEFAULT_DUR = CLIPGEN_CONFIG.defaultDuration;
    var baselineSeconds = 0;
    if (participantId && state.convergenceBaselines) {
      baselineSeconds = state.convergenceBaselines[participantId] || 0;
    }
    var segments = parseClipSegmentsForCell(raw, baselineSeconds, DEFAULT_DUR);
    if (segments.length === 0) {
      segments.push({ startSeconds: 0, duration: DEFAULT_DUR });
    }
    return segments;
  }

  // Cross-referencing: find overlapping data from other sources for a given
  // participant + time range. Used by both Screenspace and Transcript intake
  // card renderers to surface context from sibling data sources.
  function findOverlappingData(participant, start, end) {
    var result = { transcriptSnippets: [], screenspaceEvents: [], sheetObservations: [] };

    // Transcript marks/clusters — keep a small projection because `text` has
    // fallback logic (text || label) that consumers expect already resolved.
    for (var i = 0; i < state.trIntakeClusters.length; i++) {
      var tc = state.trIntakeClusters[i];
      if (tc.participant === participant && tc.start < end && tc.end > start) {
        result.transcriptSnippets.push({ text: tc.text || tc.label || "", category: tc.category, start: tc.start, end: tc.end });
      }
    }

    // Screenspace event clusters — pass through the original object; consumers
    // only read detector / event_type and the extra fields are harmless.
    for (var j = 0; j < state.intakeClusters.length; j++) {
      var sc = state.intakeClusters[j];
      if (sc.participant === participant && sc.start < end && sc.end > start) {
        result.screenspaceEvents.push(sc);
      }
    }

    // Sheet observations — pass through the row directly.
    if (state.sheetData && state.sheetData.rows) {
      for (var k = 0; k < state.sheetData.rows.length; k++) {
        var row = state.sheetData.rows[k];
        var cell = row.cells[participant];
        if (!cell || !cell.valid) continue;
        var segs = parseClipTimestamps(cell.value, participant);
        for (var s = 0; s < segs.length; s++) {
          var segEnd = segs[s].startSeconds + segs[s].duration;
          if (segs[s].startSeconds < end && segEnd > start) {
            result.sheetObservations.push(row);
            break;
          }
        }
      }
    }

    return result;
  }

  // Cached once — `--radius` lives on :root in tokens.css and isn't expected to
  // change at runtime. Avoids a getComputedStyle call inside setCardDragImage,
  // which is on the dragstart hot path.
  var _cardDragImageRadius = null;
  function getCardDragImageRadius() {
    if (_cardDragImageRadius !== null) return _cardDragImageRadius;
    var raw = getComputedStyle(document.documentElement).getPropertyValue("--radius").trim();
    _cardDragImageRadius = raw ? "calc(" + raw + " - 2px)" : "4px";
    return _cardDragImageRadius;
  }

  function setCardDragImage(ev, card) {
    var clone = card.cloneNode(true);
    var rect = card.getBoundingClientRect();
    clone.style.position = "absolute";
    clone.style.top = "-9999px";
    clone.style.left = "-9999px";
    clone.style.width = rect.width + "px";
    clone.style.zIndex = "-1";
    // Some Chromium versions render the drag-image bitmap without honoring the
    // class-driven border-radius / overflow clip, leaving square corners on
    // an otherwise rounded card. Pin them inline on the clone so the snapshot
    // captures the rounding.
    clone.style.borderRadius = getCardDragImageRadius();
    clone.style.overflow = "hidden";
    document.body.appendChild(clone);
    ev.dataTransfer.setDragImage(clone, ev.clientX - rect.left, ev.clientY - rect.top);
    setTimeout(function () { document.body.removeChild(clone); }, 0);
  }

  // Transparent 1×1 image used to suppress the browser's default drag preview
  // when we want a custom DOM-based ghost (see bindDragFromGrid). Cached so
  // the same Image instance is reused across drags.
  var _TRANSPARENT_DRAG_IMAGE = null;
  function getTransparentDragImage() {
    if (_TRANSPARENT_DRAG_IMAGE) return _TRANSPARENT_DRAG_IMAGE;
    var img = new Image(1, 1);
    img.src = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
    _TRANSPARENT_DRAG_IMAGE = img;
    return img;
  }

  // Single capture-phase gate: while a drag is in flight, `body.dragging` is
  // set so CSS can suspend expensive effects (backdrop-filter on the floating
  // nav, drop-target transitions, hover paint, etc.). Centralized here instead
  // of patched into every dragstart handler — see studio.css / topnav.css /
  // tokens.css for the matching rules.
  function bindDragGate() {
    function clear() { document.body.classList.remove("dragging"); }
    document.addEventListener("dragstart", function () {
      document.body.classList.add("dragging");
    }, true);
    document.addEventListener("dragend", clear, true);
    document.addEventListener("drop", clear, true);
    window.addEventListener("blur", clear);
    document.addEventListener("visibilitychange", function () {
      if (document.hidden) clear();
    });
  }

  // ---- Filtering ----

  function hasActiveFilters() {
    var f = state.filters;
    return (
      f.categories.length > 0 ||
      f.severities.length > 0 ||
      f.keywords.length > 0 ||
      f.fnMin !== null ||
      f.fnMax !== null
    );
  }

  function getFilteredRows(rows) {
    var f = state.filters;
    if (!hasActiveFilters()) return rows;

    var fnActive = state.activeFunction && ROW_FUNCTIONS[state.activeFunction];
    var participants = state.sheetData ? state.sheetData.participants : [];

    return rows.filter(function (row) {
      if (f.categories.length > 0) {
        if (!row.category || f.categories.indexOf(row.category) < 0) return false;
      }
      if (f.severities.length > 0) {
        if (!row.severity || f.severities.indexOf(row.severity) < 0) return false;
      }
      if (f.keywords.length > 0) {
        var rowKw = row.keywords || [];
        var kwMatch = false;
        for (var ki = 0; ki < f.keywords.length; ki++) {
          if (rowKw.indexOf(f.keywords[ki]) >= 0) { kwMatch = true; break; }
        }
        if (!kwMatch) return false;
      }
      if (f.fnMin !== null || f.fnMax !== null) {
        if (!fnActive) return false;
        var val = fnActive(row, participants);
        if (f.fnMin !== null && val < f.fnMin) return false;
        if (f.fnMax !== null && val > f.fnMax) return false;
      }
      return true;
    });
  }

  function clearAllFilters() {
    state.filters.categories = [];
    state.filters.severities = [];
    state.filters.keywords = [];
    state.filters.fnMin = null;
    state.filters.fnMax = null;
  }

  // ---- Sheet sidebar ----

  var SIDEBAR_VIEW_KEY = "clipgen-studio-sidebar-open";
  // VIEWS section: each entry overrides state.filters.severities. "all"
  // clears the selection; the rest pin a literal allowlist of severities.
  var SIDEBAR_VIEWS = [
    { id: "all",        label: "All" },
    { id: "highlights", label: "Highlights", severities: ["Critical", "High", "Medium"] },
    { id: "positive",   label: "Positive",   severities: ["Positive", "Very Positive"] },
  ];

  function readPersistedSidebarOpen() {
    try {
      var stored = localStorage.getItem(SIDEBAR_VIEW_KEY);
      if (stored !== null) state.sidebarOpen = (stored !== "false");
    } catch (_) {}
  }

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

    // VIEWS — vertical-list rows with counts. Active is derived from the
    // current state.filters.severities so picking severity pills below
    // re-highlights the matching view.
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

    // SEVERITY — multi-select pills. "Any severity" clears the selection;
    // each pill below toggles its severity in/out of state.filters.severities.
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
                renderSidebar();
                renderGrid();
              },
            }));
          })(sevLabel);
        }
      }
    }

    // KEYWORDS — multi-select pills for cell-level annotation tokens
    // (e.g. "!key" → "Key"). Filter is row-level (any cell in the row carries
    // the annotation) but cell-level emphasis is applied during grid render.
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
    minIn.disabled = !state.activeFunction;
    if (state.filters.fnMin !== null) minIn.value = String(state.filters.fnMin);

    var maxIn = document.createElement("input");
    maxIn.type = "number";
    maxIn.id = "sidebarFnMax";
    maxIn.placeholder = "Max";
    maxIn.disabled = !state.activeFunction;
    if (state.filters.fnMax !== null) maxIn.value = String(state.filters.fnMax);

    function onChange() {
      var mn = minIn.value.trim();
      var mx = maxIn.value.trim();
      state.filters.fnMin = mn !== "" ? parseFloat(mn) : null;
      state.filters.fnMax = mx !== "" ? parseFloat(mx) : null;
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

  function bindSidebarToggle() {
    var btn = qs("#studioSidebarToggle");
    if (!btn) return;
    btn.addEventListener("click", function () {
      state.sidebarOpen = !state.sidebarOpen;
      persistSidebarOpen();
      renderSidebar();
    });
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

  // ---- Preview tabs ----

  function initPreviewTabs() {
    var tabs = qsa(".preview-tab");
    for (var i = 0; i < tabs.length; i++) {
      tabs[i].addEventListener("click", function () {
        var target = this.dataset.tab;
        if (target === state.activePreviewTab) return;
        state.activePreviewTab = target;
        setStoredUIStateField("studio", "activeTab", target);
        var allTabs = qsa(".preview-tab");
        for (var j = 0; j < allTabs.length; j++) allTabs[j].classList.remove("active");
        this.classList.add("active");
        syncPreviewTab();
      });
    }
    restoreStoredPreviewTab();
  }

  var _previewTabRestored = false;
  function restoreStoredPreviewTab() {
    if (_previewTabRestored) return;
    var stored = getStoredUIState("studio").activeTab;
    if (!stored) return;
    var match = null;
    var tabs = qsa(".preview-tab");
    for (var i = 0; i < tabs.length; i++) {
      if (tabs[i].dataset.tab === stored && !tabs[i].classList.contains("hidden")) {
        match = tabs[i];
        break;
      }
    }
    if (!match) return;
    _previewTabRestored = true;
    if (stored === state.activePreviewTab) return;
    state.activePreviewTab = stored;
    for (var k = 0; k < tabs.length; k++) tabs[k].classList.remove("active");
    match.classList.add("active");
    syncPreviewTab();
  }

  function syncPreviewTab() {
    setActiveTabAttr(state.activePreviewTab);
    var grid = qs("#sheetGrid");
    var refreshBtn = qs("#refreshSheet");
    var intakePanel = qs("#intakePanel");
    var trIntakePanel = qs("#trIntakePanel");
    var convergencePanel = qs("#convergencePanel");
    var metadataPanel = qs("#metadataPanel");

    // Hide everything first
    grid.classList.add("hidden");
    intakePanel.classList.add("hidden");
    if (trIntakePanel) trIntakePanel.classList.add("hidden");
    if (convergencePanel) convergencePanel.classList.add("hidden");
    if (metadataPanel) metadataPanel.classList.add("hidden");
    if (refreshBtn) refreshBtn.classList.add("hidden");

    // Intake poll timers are started at DOMContentLoaded and kept running
    // across all tabs so the sub-tab counter badges stay live; switching away
    // from an intake tab no longer tears them down.
    if (window.convergenceDeactivate) window.convergenceDeactivate();
    if (window.metadataDeactivate) window.metadataDeactivate();

    var activePanel = null;
    if (state.activePreviewTab === "sheet") {
      grid.classList.remove("hidden");
      activePanel = grid;
      if (refreshBtn) refreshBtn.classList.remove("hidden");
    } else if (state.activePreviewTab === "intake") {
      intakePanel.classList.remove("hidden");
      activePanel = intakePanel;
    } else if (state.activePreviewTab === "transcript-intake") {
      if (trIntakePanel) trIntakePanel.classList.remove("hidden");
      activePanel = trIntakePanel;
    } else if (state.activePreviewTab === "convergence") {
      if (convergencePanel) convergencePanel.classList.remove("hidden");
      activePanel = convergencePanel;
      if (window.convergenceActivate) window.convergenceActivate();
    } else if (state.activePreviewTab === "metadata") {
      if (metadataPanel) metadataPanel.classList.remove("hidden");
      activePanel = metadataPanel;
      if (window.metadataActivate) window.metadataActivate();
    }
    if (activePanel) {
      activePanel.classList.add("tab-slide-enter");
      requestAnimationFrame(function () {
        requestAnimationFrame(function () {
          activePanel.classList.remove("tab-slide-enter");
        });
      });
    }
  }

  // ---- Queue persistence (sessionStorage) ----

  function saveQueues() {
    try {
      sessionStorage.setItem(QUEUE_STORAGE_KEY, JSON.stringify({
        artifactQueue: state.artifactQueue,
        reelQueue: state.reelQueue,
      }));
    } catch (_) { /* ignore quota errors */ }
  }

  function restoreQueues() {
    try {
      var raw = sessionStorage.getItem(QUEUE_STORAGE_KEY);
      if (!raw) return;
      var saved = JSON.parse(raw);
      if (saved.artifactQueue) state.artifactQueue = saved.artifactQueue;
      if (saved.reelQueue) state.reelQueue = saved.reelQueue;
    } catch (_) { /* ignore parse errors */ }
  }

  // ---- Data loading ----

  function loadSheetData() {
    apiGet("api/sheet")
      .then(function (data) {
        if (!data.ok) {
          qs("#sheetLoading").textContent =
            "Error: " + (data.error || "Unknown error");
          return;
        }
        state.sheetData = data;
        clipgenApplyConfig(data.config);
        renderHeader();
        renderSidebar();
        renderGrid();
        // Load per-participant baselines so the grid color-codes durations
        // and segment metadata in the video-relative frame (matches Python
        // prepare_clip behavior) instead of raw clock-time spans. Re-render
        // once they arrive so cell intensities reflect baselined durations.
        apiGet("api/sheet/baseline")
          .then(function (bdata) {
            state.convergenceBaselines = (bdata.ok && bdata.baselines) ? bdata.baselines : {};
            if (Object.keys(state.convergenceBaselines).length > 0) renderGrid();
          })
          .catch(function () { state.convergenceBaselines = {}; });
        populateGalleryParticipants(data.participants || []);
        var durInput = qs("#highlightsDuration");
        if (durInput && data.highlightsDuration) {
          durInput.value = data.highlightsDuration;
        }
        var tcCheckbox = qs("#titlecardEnabled");
        var tcDurInput = qs("#titlecardDuration");
        if (tcCheckbox && data.titlecardsEnabled !== undefined) {
          tcCheckbox.checked = data.titlecardsEnabled;
        }
        if (tcDurInput && data.titlecardDuration) {
          tcDurInput.value = data.titlecardDuration;
        }
        if (data.cellExpandHover !== undefined) {
          state.cellExpandHover = data.cellExpandHover;
        }
        var tcGroup = qs("#titlecardGroup");
        if (tcGroup && qs("#artifactFormat").value === "clip") {
          tcGroup.classList.remove("hidden");
        }
        restoreQueues();
        if (state.artifactQueue.length > 0 || state.reelQueue.length > 0) {
          renderArtifactQueue();
          renderReelQueue();
          updateCellClasses();
        }
        loadManifestState();
        checkConvergenceTabVisibility();
      })
      .catch(function (err) {
        qs("#sheetLoading").textContent = "Failed to load sheet: " + err;
      });
  }

  function refreshSheetData() {
    var btn = qs("#refreshSheet");
    if (!btn || btn.disabled) return;
    btn.disabled = true;
    var svg = btn.querySelector("svg");
    if (svg) svg.style.animation = "spin 0.7s linear infinite";

    apiPost("api/sheet/refresh")
      .then(function (data) {
        if (!data.ok) {
          showResult(null, "Refresh failed: " + (data.error || "Unknown error"));
          return;
        }
        loadSheetData();
      })
      .catch(function (err) {
        showResult(null, "Refresh failed: " + err);
      })
      .then(function () {
        btn.disabled = false;
        if (svg) svg.style.animation = "";
      });
  }

  // On page load, reconcile the persisted manifest against the live sheet:
  // for each manifest artifact, find its (participant, row) in sheetData,
  // mark that cell green, and re-enqueue any artifact whose cell still has
  // a valid timestamp but isn't already queued. The `seen` dedup guards
  // against multiple artifacts mapping to the same cell (e.g. two clips
  // generated from the same timestamp).
  function loadManifestState() {
    apiGet("api/manifest")
      .then(function (data) {
        if (!data.ok || !state.sheetData) return;
        var artifacts = data.artifacts || [];
        if (artifacts.length === 0) return;

        var seen = {};
        for (var i = 0; i < artifacts.length; i++) {
          var a = artifacts[i];
          if (!a.participant || !a.cellRow || a.type === "transcript") continue;
          var key = cellKey(a.participant, a.cellRow);
          if (seen[key]) continue;
          seen[key] = true;

          var matchedRow = null;
          for (var j = 0; j < state.sheetData.rows.length; j++) {
            if (state.sheetData.rows[j].rowNum === a.cellRow) {
              matchedRow = state.sheetData.rows[j];
              break;
            }
          }
          if (!matchedRow) continue;

          var cellData = matchedRow.cells[a.participant];
          if (!cellData || !cellData.valid) continue;

          state.cellResults[key] = "success";
          if (findInQueue(state.artifactQueue, a.participant, a.cellRow) >= 0) continue;

          var info = {
            participant: a.participant,
            row: a.cellRow,
            desc: matchedRow.observation || "",
            timestamp: cellData.value || "",
          };
          var entries = expandCellToSegments(info);
          for (var ei = 0; ei < entries.length; ei++) state.artifactQueue.push(entries[ei]);
        }

        state.generatedArtifacts = state.generatedArtifacts.concat(
          artifacts.filter(function (a) { return a.type !== "transcript" && a.file; })
        );
        renderArtifactQueue();
        updateCellClasses();
        updateViewerButton();
      })
      .catch(function () {});
  }

  // ---- Header rendering ----

  function renderHeader() {
    var d = state.sheetData;
    qs("#studyName").textContent = d.study || "Unknown study";
    qs("#versionInfo").textContent = d.participants.length + " participants";
  }

  // ---- Grid rendering ----

  function isRowEmpty(row, participants) {
    for (var j = 0; j < participants.length; j++) {
      var c = row.cells[participants[j]];
      if (c && c.hasText) return false;
    }
    return true;
  }

  function hasSeverityData(rows) {
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].severity) return true;
    }
    return false;
  }

  var _gridEventsBound = false;

  function renderGrid() {
    var d = state.sheetData;
    var grid = qs("#sheetGrid");
    grid.innerHTML = "";

    var showSeverity = hasSeverityData(d.rows);
    var metaCols = showSeverity ? 5 : 4;
    var visibleParticipants = d.participants.filter(function (pid) {
      return !isParticipantHidden(pid);
    });
    var totalCols = metaCols + visibleParticipants.length;
    state._gridTotalCols = totalCols;
    state._gridShowSeverity = showSeverity;
    state._gridVisibleParticipants = visibleParticipants;
    var table = el("table");

    // Colgroup for fixed column widths — participant columns share equal width
    var colgroup = document.createElement("colgroup");
    var colRowNum = document.createElement("col");
    colRowNum.style.width = "2.5rem";
    colgroup.appendChild(colRowNum);
    var colFn = document.createElement("col");
    colFn.style.width = "2.25rem";
    colgroup.appendChild(colFn);
    var colObs = document.createElement("col");
    // Explicit width keeps the table size predictable under table-layout: fixed
    // — `auto` collapses to 0 when other cols already exceed the table width.
    // The td inside has overflow:hidden + ellipsis to clamp the long observation.
    colObs.style.width = "18rem";
    colgroup.appendChild(colObs);
    var colCat = document.createElement("col");
    colCat.style.width = "7.5rem";
    colgroup.appendChild(colCat);
    if (showSeverity) {
      var colSev = document.createElement("col");
      colSev.style.width = "6.875rem";
      colgroup.appendChild(colSev);
    }
    for (var c = 0; c < visibleParticipants.length; c++) {
      var colP = document.createElement("col");
      colP.className = "col-participant-col";
      colgroup.appendChild(colP);
    }
    table.appendChild(colgroup);

    var thead = el("thead");
    var hrow = el("tr");

    var batchTh = el("th", "col-row-num col-row-num-header", "#");
    batchTh.title = "Select all cells";
    hrow.appendChild(batchTh);

    var fnTh = el("th", "col-function");
    var fnWrap = el("div", "fn-header-wrap");
    var fnSelect = document.createElement("select");
    fnSelect.className = "fn-select";
    fnSelect.title = "Row function";
    var defaultOpt = document.createElement("option");
    defaultOpt.value = "";
    defaultOpt.textContent = "\u0192";
    defaultOpt.selected = !state.activeFunction;
    fnSelect.appendChild(defaultOpt);
    var fnNames = Object.keys(ROW_FUNCTIONS);
    for (var fi = 0; fi < fnNames.length; fi++) {
      var fnOpt = document.createElement("option");
      fnOpt.value = fnNames[fi];
      fnOpt.textContent = fnNames[fi];
      if (state.activeFunction === fnNames[fi]) fnOpt.selected = true;
      fnSelect.appendChild(fnOpt);
    }
    var fnClear = el("button", "fn-clear", "\u00d7");
    fnClear.title = "Clear function";
    fnClear.type = "button";
    if (!state.activeFunction) fnClear.style.display = "none";
    fnSelect.addEventListener("change", function () {
      state.activeFunction = this.value;
      fnClear.style.display = this.value ? "" : "none";
      syncFilterFnDisabled();
      updateFunctionColumn();
    });
    fnClear.addEventListener("click", function () {
      state.activeFunction = "";
      fnSelect.value = "";
      this.style.display = "none";
      syncFilterFnDisabled();
      updateFunctionColumn();
    });
    fnWrap.appendChild(fnSelect);
    fnWrap.appendChild(fnClear);
    fnTh.appendChild(fnWrap);
    hrow.appendChild(fnTh);

    hrow.appendChild(el("th", "col-observation", "Observation"));
    hrow.appendChild(el("th", "col-category", "Category"));
    if (showSeverity) {
      hrow.appendChild(el("th", "col-severity", "Severity"));
    }

    for (var p = 0; p < visibleParticipants.length; p++) {
      var pTh = el("th", "col-participant", visibleParticipants[p]);
      pTh.setAttribute("data-participant", visibleParticipants[p]);
      pTh.title = "Select all " + visibleParticipants[p] + " cells";
      hrow.appendChild(pTh);
    }
    thead.appendChild(hrow);
    table.appendChild(thead);

    var tbody = el("tbody");
    tbody.id = "gridTbody";
    table.appendChild(tbody);
    grid.appendChild(table);

    applyGridFilters();

    if (!_gridEventsBound) {
      bindGridEvents();
      bindDragFromGrid();
      _gridEventsBound = true;
    }
    if (state.activeFunction) updateFunctionColumn();
  }

  function applyGridFilters() {
    var d = state.sheetData;
    if (!d) return;
    var tbody = qs("#gridTbody");
    if (!tbody) return;
    var totalCols = state._gridTotalCols;
    var showSeverity = state._gridShowSeverity;
    var visibleParticipants = state._gridVisibleParticipants || d.participants;

    var filteredRows = getFilteredRows(d.rows);
    var frag = document.createDocumentFragment();
    var i = 0;
    while (i < filteredRows.length) {
      var row = filteredRows[i];
      if (isRowEmpty(row, visibleParticipants)) {
        var emptyStart = i;
        while (i < filteredRows.length && isRowEmpty(filteredRows[i], visibleParticipants)) {
          i++;
        }
        var emptyCount = i - emptyStart;
        var sepTr = el("tr", "empty-rows-separator");
        var sepTd = el(
          "td",
          "",
          emptyCount === 1 ? "1 empty row" : emptyCount + " empty rows"
        );
        sepTd.setAttribute("colspan", String(totalCols));
        sepTr.appendChild(sepTd);
        frag.appendChild(sepTr);
      } else {
        frag.appendChild(renderDataRow(row, visibleParticipants, showSeverity));
        i++;
      }
    }
    tbody.innerHTML = "";
    tbody.appendChild(frag);
    updateCellClasses();
    if (state.activeFunction) updateFunctionColumn();
  }

  // ---- Panel divider (resizable split between sheet preview and bottom panel) ----
  //
  // Layout model: #sheetPreview is `flex: 1 1 auto` and #bottomPanel has an
  // explicit pixel `height` set from state.bottomH. The drag updates that
  // pixel height directly; the upper pane absorbs the remainder via flex.
  // state.bottomH is clamped to [BOTTOM_STRIP_MIN, BOTTOM_STRIP_MAX].
  var BOTTOM_STRIP_MIN = 60;
  var BOTTOM_STRIP_MAX = 560;
  var BOTTOM_STRIP_DEFAULT = 380;
  var BOTTOM_STORAGE_KEY = "clipgen-studio-bottom-h";

  function applyBottomHeight() {
    var bottom = qs("#bottomPanel");
    if (!bottom) return;
    if (state.bottomCollapsed) {
      bottom.style.height = "";
    } else {
      bottom.style.height = state.bottomH + "px";
    }
  }

  function loadStoredBottomHeight() {
    state.bottomH = BOTTOM_STRIP_DEFAULT;
    try {
      var raw = window.localStorage.getItem(BOTTOM_STORAGE_KEY);
      if (raw) {
        var parsed = JSON.parse(raw);
        if (parsed && typeof parsed.bottomH === "number") {
          state.bottomH = Math.max(BOTTOM_STRIP_MIN, Math.min(BOTTOM_STRIP_MAX, parsed.bottomH));
        }
        if (parsed && parsed.collapsed) {
          state.bottomCollapsed = true;
          document.body.classList.add("bottom-collapsed");
        }
      }
    } catch (_) {}
    applyBottomHeight();
  }

  function persistBottomHeight() {
    try {
      window.localStorage.setItem(BOTTOM_STORAGE_KEY, JSON.stringify({
        bottomH: state.bottomH,
        collapsed: !!state.bottomCollapsed,
      }));
    } catch (_) {}
  }

  function initPanelDivider() {
    var handle = qs("#panelDivider");
    if (!handle) return;
    var dragging = false;
    var startY = 0;
    var startBottomH = 0;
    var dragMaxBottom = BOTTOM_STRIP_MAX;

    function onDown(e) {
      if (state.bottomCollapsed) return;
      e.preventDefault();
      dragging = true;
      startY = e.clientY || (e.touches && e.touches[0].clientY) || 0;
      startBottomH = state.bottomH;

      // Reserve at least MIN_UPPER for the sheet pane.
      var header = qs("#studioSubheader");
      var divider = qs("#panelDivider");
      if (header && divider) {
        var headerRect = header.getBoundingClientRect();
        var available = window.innerHeight - headerRect.top - headerRect.height - divider.offsetHeight;
        var MIN_UPPER = 120;
        dragMaxBottom = Math.max(BOTTOM_STRIP_MIN, Math.min(BOTTOM_STRIP_MAX, available - MIN_UPPER));
      } else {
        dragMaxBottom = BOTTOM_STRIP_MAX;
      }

      handle.classList.add("active");
      document.body.style.cursor = "row-resize";
      document.body.style.userSelect = "none";
    }

    handle.addEventListener("mousedown", onDown);
    handle.addEventListener("touchstart", onDown, { passive: false });

    var rafPending = false;
    document.addEventListener("mousemove", onMove);
    document.addEventListener("touchmove", onMove, { passive: false });

    function onMove(e) {
      if (!dragging || rafPending) return;
      rafPending = true;
      var clientY = e.clientY || (e.touches && e.touches[0].clientY) || 0;
      requestAnimationFrame(function () {
        // Drag up (clientY decreases) → bottom grows. Drag down → bottom shrinks.
        var delta = startY - clientY;
        var nextH = startBottomH + delta;
        state.bottomH = Math.max(BOTTOM_STRIP_MIN, Math.min(dragMaxBottom, nextH));
        applyBottomHeight();
        rafPending = false;
      });
    }

    function onUp() {
      if (!dragging) return;
      dragging = false;
      handle.classList.remove("active");
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      persistBottomHeight();
    }

    document.addEventListener("mouseup", onUp);
    document.addEventListener("touchend", onUp);

    handle.addEventListener("dblclick", function (e) {
      e.preventDefault();
      toggleBottomPanel();
    });
  }

  function toggleBottomPanel() {
    var bottom = qs("#bottomPanel");
    if (!bottom || bottom._transitioning) return;
    bottom._transitioning = true;

    if (state.bottomCollapsed) {
      // --- Restore ---
      state.bottomCollapsed = false;
      document.body.classList.add("bottom-animating");
      // Animate maxHeight from 0 to the persisted bottomH; afterwards drop the
      // inline maxHeight so the explicit `height` style takes over.
      bottom.style.maxHeight = "0px";
      bottom.offsetHeight; // reflow at 0
      document.body.classList.remove("bottom-collapsed");
      bottom.style.maxHeight = state.bottomH + "px";
      bottom.style.height = state.bottomH + "px";

      onCollapseTransitionEnd(bottom, function () {
        document.body.classList.remove("bottom-animating");
        bottom.style.maxHeight = "";
        bottom._transitioning = false;
        persistBottomHeight();
      });
    } else {
      // --- Collapse ---
      state.bottomCollapsed = true;
      document.body.classList.add("bottom-animating");
      var currentH = bottom.offsetHeight;
      bottom.style.maxHeight = currentH + "px";
      document.body.classList.add("bottom-collapsed");
      bottom.offsetHeight; // reflow
      bottom.style.maxHeight = "0px";
      bottom.style.height = "";

      onCollapseTransitionEnd(bottom, function () {
        bottom._transitioning = false;
        document.body.classList.remove("bottom-animating");
        persistBottomHeight();
      });
    }
  }

  function onCollapseTransitionEnd(el, cb) {
    var fired = false;
    function done() {
      if (fired) return;
      fired = true;
      el.removeEventListener("transitionend", handler);
      cb();
    }
    function handler(e) {
      if (e.target === el && e.propertyName === "max-height") done();
    }
    el.addEventListener("transitionend", handler);
    setTimeout(done, 400);
  }

  function renderDataRow(row, participants, showSeverity) {
    var tr = el("tr");

    var rowTd = el("td", "col-row-num col-row-num-clickable", String(row.rowNum));
    rowTd.setAttribute("data-select-row", row.rowNum);
    rowTd.title = "Select row " + row.rowNum;
    tr.appendChild(rowTd);

    var fnTd = el("td", "col-function");
    fnTd.setAttribute("data-fn-row", row.rowNum);
    if (state.activeFunction && ROW_FUNCTIONS[state.activeFunction]) {
      fnTd.textContent = ROW_FUNCTIONS[state.activeFunction](row, participants);
    }
    tr.appendChild(fnTd);

    var obsTd = el("td", "col-observation");
    obsTd.textContent = row.observation || "";
    obsTd.title = row.observation || "";
    tr.appendChild(obsTd);
    tr.appendChild(el("td", "col-category", row.category || ""));
    if (showSeverity) {
      var sevCls = "col-severity";
      var sevTd = el("td", sevCls);
      if (row.severity) {
        var sevSlug = severityClass(row.severity);
        var pill = document.createElement("span");
        pill.className = "sev-pill" + (sevSlug ? " " + sevSlug : "");
        var dot = document.createElement("span");
        dot.className = "sev-pill-dot";
        dot.setAttribute("aria-hidden", "true");
        pill.appendChild(dot);
        var lbl = document.createElement("span");
        lbl.className = "sev-pill-label";
        lbl.textContent = row.severity;
        pill.appendChild(lbl);
        sevTd.appendChild(pill);
      }
      tr.appendChild(sevTd);
    }

    var activeKeywords = state.filters.keywords;
    for (var j = 0; j < participants.length; j++) {
      var pid = participants[j];
      var cellData = row.cells[pid] || {};
      var td = el("td", "ts-cell");
      td.setAttribute("data-row", row.rowNum);
      td.setAttribute("data-participant", pid);
      td.setAttribute("data-observation", row.observation || "");
      td.setAttribute("data-category", row.category || "");
      var sevSlug = severityClass(row.severity);
      if (sevSlug) td.setAttribute("data-severity", sevSlug.replace(/^sev-/, ""));

      if (cellData.hasText) {
        var chip = document.createElement("span");
        chip.className = "ts-chip cg-mono";
        chip.textContent = cellData.value;
        td.appendChild(chip);
        if (cellData.valid) {
          td.classList.add("valid-ts");
        } else {
          td.classList.add("has-text");
        }
        td.setAttribute("draggable", "true");
      } else {
        td.classList.add("empty");
      }

      if (activeKeywords.length > 0) {
        var cellKw = cellData.keywords || [];
        var matched = false;
        for (var ck = 0; ck < activeKeywords.length; ck++) {
          if (cellKw.indexOf(activeKeywords[ck]) >= 0) { matched = true; break; }
        }
        td.classList.add(matched ? "cell-keyword-match" : "cell-keyword-dim");
      }

      tr.appendChild(td);
    }
    return tr;
  }

  function updateFunctionColumn() {
    var d = state.sheetData;
    if (!d) return;
    var fn = state.activeFunction && ROW_FUNCTIONS[state.activeFunction];
    var cells = qsa("[data-fn-row]");

    // First pass: compute values and find range
    var values = [];
    for (var i = 0; i < cells.length; i++) {
      var rowNum = parseInt(cells[i].getAttribute("data-fn-row"), 10);
      var row = null;
      for (var j = 0; j < d.rows.length; j++) {
        if (d.rows[j].rowNum === rowNum) { row = d.rows[j]; break; }
      }
      var val = (fn && row) ? fn(row, d.participants) : null;
      values.push(val);
      cells[i].textContent = val !== null ? val : "";
    }

    // Second pass: apply conditional background
    var max = 0;
    for (var k = 0; k < values.length; k++) {
      if (values[k] !== null && values[k] > max) max = values[k];
    }
    var hm = getCSSVar("--color-heatmap", "168, 130, 214");
    for (var m = 0; m < cells.length; m++) {
      if (values[m] !== null && max > 0) {
        var t = values[m] / max;
        cells[m].style.backgroundColor = "rgba(" + hm + ", " + (t * 0.45).toFixed(3) + ")";
      } else {
        cells[m].style.backgroundColor = "";
      }
    }
  }

  // ---- Cell selection ----

  function updateCellClasses() {
    var cells = qsa(".ts-cell");
    for (var i = 0; i < cells.length; i++) {
      var td = cells[i];
      var p = td.getAttribute("data-participant");
      var r = parseInt(td.getAttribute("data-row"), 10);
      var inArt = findInQueue(state.artifactQueue, p, r) >= 0;
      var inReel = findInQueue(state.reelQueue, p, r) >= 0;
      if (inArt || inReel) {
        td.classList.add("selected");
      } else {
        td.classList.remove("selected");
      }
    }
  }

  function updateSingleCellClass(participant, row) {
    var td = qs('.ts-cell[data-participant="' + participant + '"][data-row="' + row + '"]');
    if (!td) return;
    var inArt = findInQueue(state.artifactQueue, participant, row) >= 0;
    var inReel = findInQueue(state.reelQueue, participant, row) >= 0;
    td.classList.toggle("selected", inArt || inReel);
  }

  function getCellInfo(td) {
    return {
      participant: td.getAttribute("data-participant"),
      row: parseInt(td.getAttribute("data-row"), 10),
      desc: td.getAttribute("data-observation") || "",
      timestamp: td.textContent || "",
    };
  }

  function toggleArtifactCell(info) {
    if (findInQueue(state.artifactQueue, info.participant, info.row) >= 0) {
      removeAllCellEntries(state.artifactQueue, info.participant, info.row);
    } else {
      var entries = expandCellToSegments(info);
      for (var i = 0; i < entries.length; i++) state.artifactQueue.push(entries[i]);
    }
    renderArtifactQueue();
    updateSingleCellClass(info.participant, info.row);
  }

  function toggleReelCell(info) {
    if (findInQueue(state.reelQueue, info.participant, info.row) >= 0) {
      removeAllCellEntries(state.reelQueue, info.participant, info.row);
    } else {
      var entries = expandCellToSegments(info);
      for (var i = 0; i < entries.length; i++) state.reelQueue.push(entries[i]);
    }
    renderReelQueue();
    updateSingleCellClass(info.participant, info.row);
  }

  function addToQueue(targetQueue, info, renderFn) {
    var added = false;
    if (info.segIdx !== undefined) {
      if (!hasSegmentInQueue(targetQueue, info.participant, info.row, info.segIdx)) {
        targetQueue.push(info);
        added = true;
      }
    } else if (findInQueue(targetQueue, info.participant, info.row) < 0) {
      var entries = expandCellToSegments(info);
      for (var i = 0; i < entries.length; i++) targetQueue.push(entries[i]);
      added = true;
    }
    if (added) {
      renderFn();
      updateSingleCellClass(info.participant, info.row);
    }
  }

  // Collect all non-empty cell infos matching a filter
  function collectCellInfos(filterFn) {
    var infos = [];
    var cells = qsa(".ts-cell");
    for (var i = 0; i < cells.length; i++) {
      var td = cells[i];
      if (td.classList.contains("empty")) continue;
      var info = getCellInfo(td);
      if (filterFn(info)) infos.push(info);
    }
    return infos;
  }

  function toggleBatchInQueue(queue, infos, renderFn) {
    // If all cells are already in the queue, remove them; otherwise add missing ones
    var allPresent = true;
    for (var i = 0; i < infos.length; i++) {
      if (findInQueue(queue, infos[i].participant, infos[i].row) < 0) {
        allPresent = false;
        break;
      }
    }
    if (allPresent) {
      for (var j = 0; j < infos.length; j++) {
        removeAllCellEntries(queue, infos[j].participant, infos[j].row);
      }
    } else {
      for (var k = 0; k < infos.length; k++) {
        if (findInQueue(queue, infos[k].participant, infos[k].row) < 0) {
          var entries = expandCellToSegments(infos[k]);
          for (var m = 0; m < entries.length; m++) queue.push(entries[m]);
        }
      }
    }
    renderFn();
    updateCellClasses();
  }

  // ---- Grid events ----

  function bindGridEvents() {
    var grid = qs("#sheetGrid");

    grid.addEventListener("selectstart", function (ev) {
      ev.preventDefault();
    });

    grid.addEventListener("click", function (ev) {
      // Batch select: click # header
      var batchTh = ev.target.closest(".col-row-num-header");
      if (batchTh) {
        var allInfos = collectCellInfos(function () { return true; });
        if (allInfos.length > 0) {
          var targetQueue = ev.shiftKey ? state.reelQueue : state.artifactQueue;
          var renderFn = ev.shiftKey ? renderReelQueue : renderArtifactQueue;
          toggleBatchInQueue(targetQueue, allInfos, renderFn);
        }
        return;
      }

      // Column select: click participant header
      var pTh = ev.target.closest(".col-participant");
      if (pTh && pTh.tagName === "TH") {
        var pid = pTh.getAttribute("data-participant");
        if (pid) {
          var colInfos = collectCellInfos(function (info) {
            return info.participant === pid;
          });
          if (colInfos.length > 0) {
            var tq = ev.shiftKey ? state.reelQueue : state.artifactQueue;
            var rf = ev.shiftKey ? renderReelQueue : renderArtifactQueue;
            toggleBatchInQueue(tq, colInfos, rf);
          }
        }
        return;
      }

      // Row select: click row number
      var rowTd = ev.target.closest("[data-select-row]");
      if (rowTd) {
        var rowNum = parseInt(rowTd.getAttribute("data-select-row"), 10);
        var rowInfos = collectCellInfos(function (info) {
          return info.row === rowNum;
        });
        if (rowInfos.length > 0) {
          var tq2 = ev.shiftKey ? state.reelQueue : state.artifactQueue;
          var rf2 = ev.shiftKey ? renderReelQueue : renderArtifactQueue;
          toggleBatchInQueue(tq2, rowInfos, rf2);
        }
        return;
      }

      // Single cell select
      var td = ev.target.closest(".ts-cell");
      if (!td || td.classList.contains("empty")) return;
      var info = getCellInfo(td);

      if (ev.shiftKey) {
        toggleReelCell(info);
      } else {
        toggleArtifactCell(info);
      }
    });

    grid.addEventListener("contextmenu", function (ev) {
      var td = ev.target.closest(".ts-cell");
      if (!td || td.classList.contains("empty")) return;
      ev.preventDefault();
      var info = getCellInfo(td);
      toggleReelCell(info);
    });

    // Floating expanded cell for overflowing timestamp cells
    var cellFloat = document.createElement("div");
    cellFloat.className = "ts-cell-float";
    cellFloat.style.display = "none";
    document.body.appendChild(cellFloat);
    var floatCell = null;

    function showFloat(td) {
      if (!state.cellExpandHover) return;
      // Anchor the float to the chip's box (not the td) so the float reads as
      // the same chip widening rather than a tooltip popping in.
      var chip = td.querySelector(".ts-chip");
      if (!chip) return;
      if (chip.scrollWidth <= chip.clientWidth + 1) return;
      floatCell = td;
      var rect = chip.getBoundingClientRect();
      cellFloat.textContent = chip.textContent;
      cellFloat.classList.toggle("has-text", td.classList.contains("has-text"));
      cellFloat.setAttribute("data-severity", td.getAttribute("data-severity") || "");
      cellFloat.style.top = rect.top + "px";
      cellFloat.style.left = rect.left + "px";
      cellFloat.style.height = rect.height + "px";
      cellFloat.style.display = "inline-flex";
      cellFloat.offsetWidth; // force reflow
      cellFloat.style.opacity = "1";
    }

    function hideFloat() {
      cellFloat.style.opacity = "0";
      floatCell = null;
      setTimeout(function () {
        if (!floatCell) cellFloat.style.display = "none";
      }, 70);
    }

    grid.addEventListener("mouseover", function (ev) {
      var td = ev.target.closest(".ts-cell");
      if (!td || td.classList.contains("empty")) { if (floatCell) hideFloat(); return; }
      if (td !== floatCell) {
        if (floatCell) hideFloat();
        showFloat(td);
      }
    });

    grid.addEventListener("mouseleave", function () {
      if (floatCell) hideFloat();
    });

    grid.addEventListener("scroll", function () {
      if (floatCell) hideFloat();
    }, { passive: true });
  }

  // ---- Drag from grid ----

  // Defensive click-vs-drag threshold layered on top of the browser's own
  // dragstart heuristic. The native ghost is suppressed at dragstart and we
  // wait until the cursor has moved this many pixels before mounting our
  // cascade preview, so a small click-with-jitter never flashes a ghost.
  var _CELL_DRAG_THRESHOLD_PX = 6;
  var _CELL_GHOST_OFFSET_X = 14;
  var _CELL_GHOST_OFFSET_Y = 10;

  // Build a fixed-position overlay holding one queue-style card per parsed
  // segment, matching the look of cards in the Artifact/Reel queues. Cards
  // stack down-right via the --i custom property (see studio.css). The
  // .queue-card-thumb's surface-alt background acts as a skeleton state
  // until the eagerly-loaded thumbnail resolves.
  function buildCellDragGhost(info, segments) {
    var ghost = el("div", "cell-drag-ghost");
    var n = segments.length;
    for (var i = 0; i < n; i++) {
      var seg = segments[i];
      var card = el("div", "queue-card clip-card size-md cell-drag-ghost-card");
      card.style.setProperty("--i", i);

      var thumb = el("div", "queue-card-thumb");
      var img = document.createElement("img");
      img.alt = "";
      img.draggable = false;
      img.src = "api/thumbnail/" + encodeURIComponent(info.participant) + "/" + seg.start;
      thumb.appendChild(img);
      // Mirror the queue-card error fallback (red X, queue-card-error class)
      // so a missing video / invalid thumbnail looks the same here as it does
      // once dropped into the Artifact or Reel queue.
      (function (cardEl, thumbEl) {
        img.addEventListener("error", function () {
          this.remove();
          thumbEl.appendChild(el("span", "", "✕"));
          cardEl.classList.add("queue-card-error");
        });
      })(card, thumb);
      thumb.appendChild(el("span", "queue-card-duration", formatDuration(seg.end - seg.start)));
      card.appendChild(thumb);

      var meta = el("div", "queue-card-meta");
      var refText = info.participant + "." + info.row;
      if (n > 1) refText += " (" + (i + 1) + "/" + n + ")";
      meta.appendChild(el("span", "queue-card-ref", refText));
      card.appendChild(meta);

      ghost.appendChild(card);
    }
    document.body.appendChild(ghost);
    return ghost;
  }

  function bindDragFromGrid() {
    var grid = qs("#sheetGrid");
    var pointerOrigin = null;   // last pointerdown position on a .ts-cell
    var pendingDrag = null;     // {info, originX, originY} until ghost mounts
    var ghost = null;           // active overlay element
    var rafPending = 0;
    var cursorX = 0;
    var cursorY = 0;

    grid.addEventListener("pointerdown", function (ev) {
      var td = ev.target.closest(".ts-cell");
      if (!td || td.classList.contains("empty")) {
        pointerOrigin = null;
        return;
      }
      pointerOrigin = { x: ev.clientX, y: ev.clientY };
    });

    function clearPointerOrigin() { pointerOrigin = null; }
    document.addEventListener("pointerup", clearPointerOrigin, true);
    document.addEventListener("pointercancel", clearPointerOrigin, true);

    grid.addEventListener("dragstart", function (ev) {
      var td = ev.target.closest(".ts-cell");
      if (!td || td.classList.contains("empty")) return;

      var info = getCellInfo(td);
      ev.dataTransfer.setData("application/json", JSON.stringify(info));
      ev.dataTransfer.effectAllowed = "copy";

      // Suppress the browser's snapshot — we render a custom cascade overlay.
      try { ev.dataTransfer.setDragImage(getTransparentDragImage(), 0, 0); } catch (_) {}

      pendingDrag = {
        info: info,
        originX: pointerOrigin ? pointerOrigin.x : ev.clientX,
        originY: pointerOrigin ? pointerOrigin.y : ev.clientY,
      };
      cursorX = ev.clientX;
      cursorY = ev.clientY;
    });

    function positionGhost() {
      if (!ghost) return;
      ghost.style.transform = "translate("
        + (cursorX + _CELL_GHOST_OFFSET_X) + "px, "
        + (cursorY + _CELL_GHOST_OFFSET_Y) + "px)";
    }

    function ensureGhostBuilt() {
      if (ghost || !pendingDrag) return;
      var dx = cursorX - pendingDrag.originX;
      var dy = cursorY - pendingDrag.originY;
      if (Math.hypot(dx, dy) < _CELL_DRAG_THRESHOLD_PX) return;
      var info = pendingDrag.info;
      var segments = expandCellToSegments(info);
      ghost = buildCellDragGhost(info, segments);
      pendingDrag = null;
      positionGhost();
      // Flip on the .in class one frame later so the entrance transition runs.
      requestAnimationFrame(function () {
        if (ghost) ghost.classList.add("in");
      });
    }

    function onDragOver(ev) {
      if (!pendingDrag && !ghost) return;
      cursorX = ev.clientX;
      cursorY = ev.clientY;
      ensureGhostBuilt();
      if (!ghost || rafPending) return;
      rafPending = requestAnimationFrame(function () {
        rafPending = 0;
        positionGhost();
      });
    }

    function cleanup() {
      pendingDrag = null;
      if (rafPending) {
        cancelAnimationFrame(rafPending);
        rafPending = 0;
      }
      if (ghost) {
        var node = ghost;
        ghost = null;
        node.classList.remove("in");
        node.classList.add("out");
        setTimeout(function () {
          if (node.parentNode) node.parentNode.removeChild(node);
        }, 140);
      }
      pointerOrigin = null;
    }

    document.addEventListener("dragover", onDragOver, true);
    document.addEventListener("dragend", cleanup, true);
    document.addEventListener("drop", cleanup, true);
    // mouseup fires immediately on release regardless of drop target. dragend
    // is delayed up to ~1s by the browser's snap-back animation when a drop
    // is rejected (e.g. dropped on the sheet, not on a queue), so without
    // this the ghost lingers visibly. cleanup() is idempotent.
    document.addEventListener("mouseup", cleanup, true);
    window.addEventListener("blur", cleanup);
    document.addEventListener("visibilitychange", function () {
      if (document.hidden) cleanup();
    });
  }

  // ---- Drop targets ----

  function removeFromQueue(queue, info) {
    if (info.segIdx !== undefined) {
      removeSegmentEntry(queue, info.participant, info.row, info.segIdx);
    } else {
      removeAllCellEntries(queue, info.participant, info.row);
    }
  }

  function initDropTargets() {
    setupDropTarget(qs("#artifactsList"), function (info) {
      if (info.source === "reel-stash" || info.source === "artifact-stash") {
        for (var i = 0; i < info.items.length; i++)
          addToQueue(state.artifactQueue, info.items[i], renderArtifactQueue);
        return;
      }
      if (info.source === "screenspace" || info.source === "transcript") {
        state.artifactQueue.push(info);
        renderArtifactQueue();
        return;
      }
      if (info.source === "reel") {
        removeFromQueue(state.reelQueue, info);
        renderReelQueue();
      }
      addToQueue(state.artifactQueue, info, renderArtifactQueue);
    });
    setupDropTarget(qs("#reelList"), function (info) {
      if (info.source === "reel-stash" || info.source === "artifact-stash") {
        for (var i = 0; i < info.items.length; i++)
          addToQueue(state.reelQueue, info.items[i], renderReelQueue);
        return;
      }
      if (info.source === "screenspace" || info.source === "transcript") {
        state.reelQueue.push(info);
        renderReelQueue();
        return;
      }
      if (info.source === "artifact") {
        removeFromQueue(state.artifactQueue, info);
        renderArtifactQueue();
      }
      addToQueue(state.reelQueue, info, renderReelQueue);
    });

    setupDropTarget(qs("#stashedReelsList"), function (info) {
      if (info.source === "reel-stash" || info.source === "artifact-stash") {
        createStashViaAPI("api/stashes", info.items, function (stash) {
          state.stashes.push(stash);
          renderStashedReels();
        });
      }
    });
    setupDropTarget(qs("#stashedArtifactsList"), function (info) {
      if (info.source === "reel-stash" || info.source === "artifact-stash") {
        createStashViaAPI("api/artifact-stashes", info.items, function (stash) {
          state.artifactStashes.push(stash);
          renderStashedArtifacts();
        });
      }
    });
  }

  function setupDropTarget(target, onDrop) {
    target.addEventListener("dragover", function (ev) {
      ev.preventDefault();
      ev.dataTransfer.dropEffect = "copy";
      // dragover fires ~60Hz; skip the no-op write when the class is already
      // applied so we don't churn the attribute / invalidate style.
      if (!target.classList.contains("drag-over")) target.classList.add("drag-over");
    });
    target.addEventListener("dragleave", function (ev) {
      if (!target.contains(ev.relatedTarget)) {
        target.classList.remove("drag-over");
      }
    });
    target.addEventListener("drop", function (ev) {
      ev.preventDefault();
      target.classList.remove("drag-over");
      try {
        var info = JSON.parse(ev.dataTransfer.getData("application/json"));
        if (info && (info.participant || info.items)) {
          onDrop(info);
        }
      } catch (_) {}
    });
  }

  // ---- Wheel-to-horizontal scroll for card queues ----

  function initWheelScroll() {
    ["#artifactsList", "#reelList"].forEach(function (sel) {
      var el = qs(sel);
      el.addEventListener(
        "wheel",
        function (e) {
          if (el.scrollWidth > el.clientWidth) {
            e.preventDefault();
            el.scrollLeft += e.deltaY;
          }
        },
        { passive: false },
      );
    });
  }

  // ---- Reel reordering ----

  var _reelDragIdx = null;

  function bindReelReorder() {
    var list = qs("#reelList");

    list.addEventListener("dragstart", function (ev) {
      var card = ev.target.closest(".reel-card[data-reel-idx]");
      if (!card) return;
      _reelDragIdx = parseInt(card.getAttribute("data-reel-idx"), 10);
      ev.dataTransfer.effectAllowed = "copyMove";
      ev.dataTransfer.setData("text/plain", String(_reelDragIdx));
      var reelItem = state.reelQueue[_reelDragIdx];
      if (reelItem) {
        var data = {
          participant: reelItem.participant,
          row: reelItem.row,
          desc: reelItem.desc,
          timestamp: reelItem.timestamp,
          segIdx: reelItem.segIdx,
          start: reelItem.start,
          end: reelItem.end,
          segTotal: reelItem.segTotal,
          source: "reel",
        };
        ev.dataTransfer.setData("application/json", JSON.stringify(data));
      }
    });

    list.addEventListener("dragover", function (ev) {
      var item = ev.target.closest(".reel-card[data-reel-idx]");
      if (item && _reelDragIdx !== null) {
        ev.preventDefault();
        ev.dataTransfer.dropEffect = "move";
      }
    });

    list.addEventListener("drop", function (ev) {
      var item = ev.target.closest(".reel-card[data-reel-idx]");
      if (!item || _reelDragIdx === null) return;
      ev.preventDefault();
      var toIdx = parseInt(item.getAttribute("data-reel-idx"), 10);
      if (_reelDragIdx !== toIdx) {
        var moved = state.reelQueue.splice(_reelDragIdx, 1)[0];
        state.reelQueue.splice(toIdx, 0, moved);
        renderReelQueue();
      }
      _reelDragIdx = null;
    });

    list.addEventListener("dragend", function () {
      _reelDragIdx = null;
    });
  }

  // ---- Queue rendering ----

  function renderArtifactQueue() {
    clearGridHighlights();
    var list = qs("#artifactsList");
    var n = state.artifactQueue.length;
    qs("#artifactsCount").textContent = "(" + n + ")";
    qs("#generateBtn").disabled = n === 0;
    if (n === 0) qs("#generateBtn").setAttribute("data-tooltip", "Add cells to the work area first");
    qs("#addToReelBtn").disabled = n === 0;
    if (n === 0) qs("#addToReelBtn").setAttribute("data-tooltip", "Add cells to the work area first");
    qs("#stashArtifactsBtn").disabled = n === 0;
    if (n === 0) qs("#stashArtifactsBtn").setAttribute("data-tooltip", "Add cells to the work area first");
    list.innerHTML = "";
    saveQueues();

    if (n === 0) {
      list.appendChild(
        el("div", "queue-card-ghost", "Click or drag cells here to queue for generation")
      );
      return;
    }

    for (var i = 0; i < n; i++) {
      var item = state.artifactQueue[i];
      var isIntake = isIntakeSource(item.source);
      var segTotal = item.segTotal || 1;
      var segIdx = item.segIdx || 0;

      var card = el(
        "div",
        "queue-card clip-card size-md" + (isIntake ? " queue-card-intake" : ""),
      );
      card.setAttribute("data-participant", item.participant);
      card.setAttribute("data-row", isIntake ? "" : item.row);
      if (isIntake) card.setAttribute("data-source", item.source);
      card.setAttribute("data-seg-idx", segIdx);
      card.setAttribute("draggable", "true");
      (function (itm, isI) {
        card.addEventListener("dragstart", function (ev) {
          var data = {
            participant: itm.participant,
            desc: itm.desc,
            start: itm.start,
            end: itm.end,
            source: isI ? itm.source : "artifact",
          };
          if (!isI) {
            data.row = itm.row;
            data.timestamp = itm.timestamp;
            data.segIdx = itm.segIdx;
            data.segTotal = itm.segTotal;
          } else {
            data.event_type = itm.event_type;
            data.event_ids = itm.event_ids;
            data.mark_ids = itm.mark_ids;
          }
          ev.dataTransfer.setData("application/json", JSON.stringify(data));
          ev.dataTransfer.effectAllowed = "copyMove";
          setCardDragImage(ev, this);
        });
      })(item, isIntake);

      var thumb = el("div", "queue-card-thumb");
      var img = document.createElement("img");
      img.alt = "";
      img.draggable = false;
      thumb.appendChild(img);
      if (isIntake) {
        ssObserveThumb(card, img, thumb, item.participant, item.start);
      } else {
        img.src = "api/thumbnail/" + encodeURIComponent(item.participant) + "/" + item.start;
        img.loading = "lazy";
        (function (cardEl, thumbEl) {
          img.addEventListener("error", function () {
            this.remove();
            thumbEl.appendChild(el("span", "", "\u2715"));
            cardEl.classList.add("queue-card-error");
          });
        })(card, thumb);
      }
      thumb.appendChild(el("span", "queue-card-duration", formatDuration(item.end - item.start)));
      if (isIntake) {
        var ssBadge = el("span", "queue-card-source-badge");
        if (item.source === "transcript") {
          ssBadge.innerHTML = iconHTML("bars-3");
        } else {
          ssBadge.innerHTML = iconHTML("squares-2x2");
        }
        thumb.appendChild(ssBadge);
      }
      card.appendChild(thumb);

      var meta = el("div", "queue-card-meta");
      var refText;
      if (isIntake) {
        refText = item.participant + " \u00b7 " + (item.event_type || item.desc || "intake");
      } else {
        refText = item.participant + "." + item.row;
        if (segTotal > 1) refText += " (" + (segIdx + 1) + "/" + segTotal + ")";
      }
      meta.appendChild(el("span", "queue-card-ref", refText));
      card.appendChild(meta);

      var removeBtn = el("button", "queue-card-remove");
      removeBtn.innerHTML = iconHTML("x-mark");
      removeBtn.title = "Remove";
      (function (idx) {
        removeBtn.addEventListener("click", function (ev) {
          ev.stopPropagation();
          var removed = state.artifactQueue.splice(idx, 1)[0];
          if (removed.row) delete state.cellResults[cellKey(removed.participant, removed.row)];
          renderArtifactQueue();
          updateCellClasses();
        });
      })(i);
      card.appendChild(removeBtn);

      if (!isIntake) {
        (function (p, r) {
          card.addEventListener("mouseenter", function () { highlightGridHeaders(p, r); });
          card.addEventListener("mouseleave", clearGridHighlights);
        })(item.participant, item.row);
      }

      list.appendChild(card);
    }
    applyCardStates(list);
  }

  function renderReelQueue() {
    clearGridHighlights();
    var list = qs("#reelList");
    var n = state.reelQueue.length;
    qs("#reelCount").textContent = "(" + n + ")";
    qs("#buildReelBtn").disabled = n === 0;
    if (n === 0) qs("#buildReelBtn").setAttribute("data-tooltip", "Add clips to the reel first");
    qs("#stashReelBtn").disabled = n === 0;
    if (n === 0) qs("#stashReelBtn").setAttribute("data-tooltip", "Add clips to the reel first");
    list.innerHTML = "";
    saveQueues();

    if (n === 0) {
      list.appendChild(
        el("div", "queue-card-ghost", "Shift+click or drag cells here to build a reel")
      );
      qs("#reelDuration").textContent = "";
      return;
    }

    var totalDur = 0;
    for (var i = 0; i < n; i++) {
      var item = state.reelQueue[i];
      var isIntake = isIntakeSource(item.source);
      var segDuration = item.end - item.start;
      var segTotal = item.segTotal || 1;
      var segIdx = item.segIdx || 0;
      totalDur += segDuration;

      var card = el(
        "div",
        "queue-card reel-card clip-card size-md" + (isIntake ? " queue-card-intake" : ""),
      );
      card.setAttribute("data-reel-idx", i);
      card.setAttribute("data-participant", item.participant);
      card.setAttribute("data-row", isIntake ? "" : item.row);
      if (isIntake) card.setAttribute("data-source", item.source);
      card.setAttribute("data-seg-idx", segIdx);
      card.setAttribute("draggable", "true");

      var thumb = el("div", "queue-card-thumb");
      var img = document.createElement("img");
      img.alt = "";
      img.draggable = false;
      thumb.appendChild(img);
      if (isIntake) {
        ssObserveThumb(card, img, thumb, item.participant, item.start);
      } else {
        img.src = "api/thumbnail/" + encodeURIComponent(item.participant) + "/" + item.start;
        img.loading = "lazy";
        (function (cardEl, thumbEl) {
          img.addEventListener("error", function () {
            this.remove();
            thumbEl.appendChild(el("span", "", "\u2715"));
            cardEl.classList.add("queue-card-error");
          });
        })(card, thumb);
      }
      thumb.appendChild(el("span", "queue-card-duration", formatDuration(segDuration)));
      if (isIntake) {
        var ssBadge = el("span", "queue-card-source-badge");
        if (item.source === "transcript") {
          ssBadge.innerHTML = iconHTML("bars-3");
        } else {
          ssBadge.innerHTML = iconHTML("squares-2x2");
        }
        thumb.appendChild(ssBadge);
      }
      card.appendChild(thumb);

      var meta = el("div", "queue-card-meta");
      meta.appendChild(el("span", "reel-card-order", String(i + 1)));
      var refText;
      if (isIntake) {
        refText = item.participant + " \u00b7 " + (item.event_type || item.desc || "intake");
      } else {
        refText = item.participant + "." + item.row;
        if (segTotal > 1) refText += " (" + (segIdx + 1) + "/" + segTotal + ")";
      }
      meta.appendChild(el("span", "queue-card-ref", refText));
      card.appendChild(meta);

      var removeBtn = el("button", "queue-card-remove");
      removeBtn.innerHTML = iconHTML("x-mark");
      removeBtn.title = "Remove";
      (function (idx) {
        removeBtn.addEventListener("click", function (ev) {
          ev.stopPropagation();
          var removed = state.reelQueue.splice(idx, 1)[0];
          if (removed.row) delete state.cellResults[cellKey(removed.participant, removed.row)];
          renderReelQueue();
          updateCellClasses();
        });
      })(i);
      card.appendChild(removeBtn);

      if (!isIntake) {
        (function (p, r) {
          card.addEventListener("mouseenter", function () { highlightGridHeaders(p, r); });
          card.addEventListener("mouseleave", clearGridHighlights);
        })(item.participant, item.row);
      }

      list.appendChild(card);
    }
    qs("#reelDuration").textContent = formatDuration(totalDur);
    applyCardStates(list);
  }

  // ---- Stashed reels ----

  function loadStashes() {
    apiGet("api/stashes")
      .then(function (data) {
        if (data.ok) {
          state.stashes = data.stashes || [];
          renderStashedReels();
        }
      })
      .catch(function () {});
  }

  function renderStashedReels() {
    var area = qs("#stashedReelsArea");
    var list = qs("#stashedReelsList");
    var n = state.stashes.length;
    qs("#stashedReelsCount").textContent = "(" + n + ")";
    area.classList.remove("stash-drop-reveal");
    list.innerHTML = "";

    if (n === 0) {
      list.appendChild(el("div", "stash-empty-hint", "Stash reels to set them aside for later."));
      return;
    }

    for (var i = 0; i < n; i++) {
      list.appendChild(buildStashCard(state.stashes[i], "api/stashes", state.stashes, renderStashedReels, "reel-stash", recallStash));
    }
  }

  function makeStashFolderIcon(stash) {
    var icon = el("span", "stash-card-icon");
    var hue = categoryHue((stash && stash.id) || "uncategorized");
    // Hue-tinted backing remains as a fallback if every thumbnail fails.
    icon.style.background = "oklch(0.32 0.06 " + hue + ")";

    var items = (stash && stash.items) || [];
    // Pull the first 3 distinct thumbnails to fake a stacked-folder look.
    var picks = [];
    var seen = {};
    for (var i = 0; i < items.length && picks.length < 3; i++) {
      var item = items[i];
      var key = item && item.participant != null && item.start != null
        ? item.participant + ":" + item.start
        : null;
      if (!key || seen[key]) continue;
      seen[key] = true;
      picks.push(item);
    }

    picks.forEach(function (item, idx) {
      var img = document.createElement("img");
      img.className = "stash-card-icon-img";
      img.alt = "";
      img.draggable = false;
      img.style.zIndex = String(picks.length - idx);
      img.style.transform = "translate(" + (idx * 2) + "px, " + (-idx * 2) + "px)";
      img.src = ssThumbUrl(item.participant, item.start);
      img.addEventListener("error", function () {
        if (img.parentNode) img.parentNode.removeChild(img);
      });
      icon.appendChild(img);
    });

    return icon;
  }

  function buildStashCard(stash, apiPath, listRef, rerender, dragSource, onRecall) {
    var card = el("div", "stash-card");
    card.setAttribute("data-stash-id", stash.id);
    card.setAttribute("draggable", "true");
    card.addEventListener("dragstart", function (ev) {
      ev.dataTransfer.setData("application/json", JSON.stringify({
        stashId: stash.id,
        items: stash.items,
        source: dragSource,
      }));
      ev.dataTransfer.effectAllowed = "copy";
    });

    card.appendChild(makeStashFolderIcon(stash));

    var nameEl = el("span", "stash-card-name", truncate(stash.name, 18));
    nameEl.title = stash.name;
    nameEl.addEventListener("click", function (ev) {
      ev.stopPropagation();
      startStashRename(stash, nameEl, apiPath);
    });
    card.appendChild(nameEl);

    var info = el("span", "stash-card-info");
    info.appendChild(el("span", "", String(stash.count)));
    info.appendChild(el("span", "", formatDuration(stash.totalDuration)));
    card.appendChild(info);

    var removeBtn = el("button", "stash-card-remove", "\u00D7");
    removeBtn.title = "Delete stash";
    removeBtn.addEventListener("click", function (ev) {
      ev.stopPropagation();
      deleteStash(stash.id, apiPath, listRef, rerender);
    });
    card.appendChild(removeBtn);

    card.addEventListener("click", function () {
      if (typeof onRecall === "function") onRecall(stash);
    });
    return card;
  }

  function computeReelDuration(items) {
    var total = 0;
    for (var i = 0; i < items.length; i++) {
      total += Math.max(0, items[i].end - items[i].start);
    }
    return total;
  }

  function stashCurrentReel() {
    if (state.reelQueue.length === 0) return;

    var items = state.reelQueue.slice();
    var totalDuration = computeReelDuration(items);
    apiPost("api/stashes", { action: "create", items: items, name: "", totalDuration: totalDuration })
      .then(function (data) {
        if (data.ok) {
          state.stashes.push(data.stash);
          for (var i = 0; i < state.reelQueue.length; i++) {
            var item = state.reelQueue[i];
            delete state.cellResults[cellKey(item.participant, item.row)];
          }
          state.reelQueue = [];
          renderReelQueue();
          renderStashedReels();
          updateCellClasses();
        }
      })
      .catch(function () {});
  }

  function recallStash(stash) {
    state.reelQueue = stash.items.slice();
    renderReelQueue();
    updateCellClasses();
  }

  function deleteStash(stashId, endpoint, stateArray, renderFn) {
    apiPost(endpoint, { action: "delete", id: stashId })
      .then(function (data) {
        if (data.ok) {
          for (var i = 0; i < stateArray.length; i++) {
            if (stateArray[i].id === stashId) {
              stateArray.splice(i, 1);
              break;
            }
          }
          renderFn();
        }
      })
      .catch(function () {});
  }

  function startStashRename(stash, nameNode, endpoint) {
    var parent = nameNode.parentNode;
    var input = document.createElement("input");
    input.className = "stash-card-name-input";
    input.type = "text";
    input.autocomplete = "off";
    input.value = stash.name;

    function commit() {
      var newName = input.value.trim() || stash.name;
      stash.name = newName;
      var span = el("span", "stash-card-name", truncate(newName, 20));
      span.title = newName;
      span.addEventListener("click", function (ev) {
        ev.stopPropagation();
        startStashRename(stash, span, endpoint);
      });
      parent.replaceChild(span, input);

      apiPost(endpoint, { action: "update", id: stash.id, name: newName }).catch(function () {});
    }

    input.addEventListener("blur", commit);
    input.addEventListener("keydown", function (ev) {
      if (ev.key === "Enter") { ev.preventDefault(); input.blur(); }
      if (ev.key === "Escape") { input.value = stash.name; input.blur(); }
    });
    input.addEventListener("click", function (ev) { ev.stopPropagation(); });

    parent.replaceChild(input, nameNode);
    input.focus();
    input.select();
  }

  function createStashViaAPI(endpoint, items, onSuccess) {
    var totalDuration = computeReelDuration(items);
    apiPost(endpoint, { action: "create", items: items, name: "", totalDuration: totalDuration })
      .then(function (data) {
        if (data.ok) onSuccess(data.stash);
      })
      .catch(function () {});
  }

  // ---- Stashed artifacts ----

  function loadArtifactStashes() {
    apiGet("api/artifact-stashes")
      .then(function (data) {
        if (data.ok) {
          state.artifactStashes = data.stashes || [];
          renderStashedArtifacts();
        }
      })
      .catch(function () {});
  }

  function renderStashedArtifacts() {
    var area = qs("#stashedArtifactsArea");
    var list = qs("#stashedArtifactsList");
    var n = state.artifactStashes.length;
    qs("#stashedArtifactsCount").textContent = "(" + n + ")";
    area.classList.remove("stash-drop-reveal");
    list.innerHTML = "";

    if (n === 0) {
      list.appendChild(el("div", "stash-empty-hint", "Stash artifacts to keep them aside — drag, or use the Stash button."));
      return;
    }

    for (var i = 0; i < n; i++) {
      list.appendChild(buildStashCard(state.artifactStashes[i], "api/artifact-stashes", state.artifactStashes, renderStashedArtifacts, "artifact-stash", recallArtifactStash));
    }
  }

  function stashCurrentArtifacts() {
    if (state.artifactQueue.length === 0) return;

    var items = state.artifactQueue.slice();
    var totalDuration = computeReelDuration(items);
    apiPost("api/artifact-stashes", { action: "create", items: items, name: "", totalDuration: totalDuration })
      .then(function (data) {
        if (data.ok) {
          state.artifactStashes.push(data.stash);
          for (var i = 0; i < state.artifactQueue.length; i++) {
            var item = state.artifactQueue[i];
            delete state.cellResults[cellKey(item.participant, item.row)];
          }
          state.artifactQueue = [];
          renderArtifactQueue();
          renderStashedArtifacts();
          updateCellClasses();
        }
      })
      .catch(function () {});
  }

  function recallArtifactStash(stash) {
    state.artifactQueue = stash.items.slice();
    renderArtifactQueue();
    updateCellClasses();
  }

  // ---- Stash drag-reveal ----

  function revealEmptyStashAreas() {
    var artArea = qs("#stashedArtifactsArea");
    var reelArea = qs("#stashedReelsArea");
    if (state.artifactStashes.length === 0) artArea.classList.add("stash-drop-reveal");
    if (state.stashes.length === 0) reelArea.classList.add("stash-drop-reveal");
  }

  function hideEmptyStashAreas() {
    var artArea = qs("#stashedArtifactsArea");
    var reelArea = qs("#stashedReelsArea");
    artArea.classList.remove("stash-drop-reveal");
    reelArea.classList.remove("stash-drop-reveal");
  }

  // ---- Buttons ----

  function bindButtons() {
    qs("#clearArtifactsBtn").addEventListener("click", function () {
      for (var i = 0; i < state.artifactQueue.length; i++) {
        var item = state.artifactQueue[i];
        delete state.cellResults[cellKey(item.participant, item.row)];
      }
      state.artifactQueue = [];
      renderArtifactQueue();
      updateCellClasses();
    });

    qs("#addToReelBtn").addEventListener("click", function () {
      for (var i = 0; i < state.artifactQueue.length; i++) {
        addToQueue(state.reelQueue, state.artifactQueue[i], renderReelQueue);
      }
      updateCellClasses();
    });

    qs("#clearReelBtn").addEventListener("click", function () {
      for (var i = 0; i < state.reelQueue.length; i++) {
        var item = state.reelQueue[i];
        delete state.cellResults[cellKey(item.participant, item.row)];
      }
      state.reelQueue = [];
      renderReelQueue();
      updateCellClasses();
    });

    qs("#stashReelBtn").addEventListener("click", stashCurrentReel);
    qs("#stashArtifactsBtn").addEventListener("click", stashCurrentArtifacts);
    qs("#generateBtn").addEventListener("click", onGenerate);
    qs("#cancelGenerateBtn").addEventListener("click", onCancelGenerate);
    qs("#buildReelBtn").addEventListener("click", onBuildReel);
    qs("#cancelReelBtn").addEventListener("click", onCancelReel);
    qs("#buildHighlightsBtn").addEventListener("click", onBuildHighlights);
    bindGalleryDialog();

    qs("#artifactFormat").addEventListener("change", function () {
      var tcGroup = qs("#titlecardGroup");
      if (tcGroup) {
        if (this.value === "clip") tcGroup.classList.remove("hidden");
        else tcGroup.classList.add("hidden");
      }
    });

    qs("#titlecardEnabled").addEventListener("change", persistTitlecardSettings);
    qs("#titlecardDuration").addEventListener("change", persistTitlecardSettings);

    qs("#refreshSheet").addEventListener("click", refreshSheetData);
    qs("#settingsBtn").addEventListener("click", function () {
      openSettingsModal({
        initialTab: "General",
        version: state.sheetData ? state.sheetData.version : "",
        onSave: function (_applied, full) {
          state.settingsData = full;
          syncInlineControls();
          _syncMarkCategoriesFromSettings(full);
        },
        onReset: function (_scope, full) {
          state.settingsData = full;
          syncInlineControls();
          _syncMarkCategoriesFromSettings(full);
        },
      });
    });

    qs("#logBtn").addEventListener("click", openLog);
    qs("#logClose").addEventListener("click", closeLog);
    qs("#logOverlay").addEventListener("click", function (e) {
      if (e.target === qs("#logOverlay")) closeLog();
    });

    qs("#statusDismiss").addEventListener("click", hideOverlay);
    qs("#statusOpen").addEventListener("click", function () {
      if (_lastViewerFile) {
        apiPost("api/open-viewer", { file: _lastViewerFile }).catch(function () {});
      }
      hideOverlay();
    });

    qs("#confirmOverlay").addEventListener("click", function (e) {
      if (e.target === qs("#confirmOverlay")) hideConfirm();
    });
  }

  // ---- Generation progress helpers ----

  function setTitleSpinner(id, active) {
    var s = qs("#" + id);
    if (s) s.classList.toggle("active", active);
  }

  function createPulserOverlay() {
    var overlay = el("div", "card-gen-overlay");
    overlay.innerHTML =
      '<svg width="26" height="10" viewBox="0 0 26 10">' +
      '<circle cx="5" cy="7" r="3"/>' +
      '<circle cx="13" cy="7" r="3"/>' +
      '<circle cx="21" cy="7" r="3"/>' +
      '</svg>';
    return overlay;
  }

  function createResultBadge(success) {
    var badge = el("div", "card-gen-badge " + (success ? "card-gen-badge-ok" : "card-gen-badge-fail"));
    badge.innerHTML = iconHTML(success ? "check" : "x-mark");
    return badge;
  }

  function setCardQueued(card) {
    card.classList.add("queue-card-queued");
    var thumb = card.querySelector(".queue-card-thumb");
    if (thumb) thumb.appendChild(createPulserOverlay());
    var p = card.getAttribute("data-participant");
    var r = card.getAttribute("data-row");
    if (p && r) delete state.cellResults[cellKey(p, parseInt(r, 10))];
  }

  function clearCardStatus(card) {
    card.classList.remove("queue-card-queued", "queue-card-success", "queue-card-fail");
    var overlay = card.querySelector(".card-gen-overlay");
    if (overlay) overlay.remove();
    var badge = card.querySelector(".card-result-badge");
    if (badge) badge.remove();
  }

  function setCardResult(card, success) {
    card.classList.remove("queue-card-queued");
    card.classList.add(success ? "queue-card-success" : "queue-card-fail");
    var overlay = card.querySelector(".card-gen-overlay");
    if (overlay) overlay.remove();
    var thumb = card.querySelector(".queue-card-thumb");
    if (thumb) thumb.appendChild(createResultBadge(success));
    var p = card.getAttribute("data-participant");
    var r = card.getAttribute("data-row");
    if (p && r) state.cellResults[cellKey(p, parseInt(r, 10))] = success ? "success" : "fail";
  }

  function applyCardStates(listEl) {
    var cards = listEl.querySelectorAll(".queue-card");
    for (var i = 0; i < cards.length; i++) {
      var card = cards[i];
      var p = card.getAttribute("data-participant");
      var r = card.getAttribute("data-row");
      var result = state.cellResults[cellKey(p, parseInt(r, 10))];
      if (result === "success" || result === "fail") {
        setCardResult(card, result === "success");
      }
    }
  }

  function setGeneratingLock(locked) {
    var ids = [
      "#generateBtn", "#buildReelBtn", "#clearArtifactsBtn",
      "#clearReelBtn", "#addToReelBtn", "#buildHighlightsBtn",
      "#stashReelBtn", "#stashArtifactsBtn"
    ];
    for (var i = 0; i < ids.length; i++) {
      var b = qs(ids[i]);
      if (b) b.disabled = locked;
    }
    document.body.classList.toggle("studio-generating", locked);
  }

  // ---- API calls ----

  function onGenerate() {
    if (state.generating || state.artifactQueue.length === 0) return;
    state.generating = true;
    setGeneratingLock(true);
    setTitleSpinner("artifactsSpinner", true);
    qs("#cancelGenerateBtn").classList.remove("hidden");

    var format = qs("#artifactFormat").value;
    var list = qs("#artifactsList");
    var items = state.artifactQueue.slice();

    // Separate spreadsheet and intake items
    var sheetItems = [];
    var intakeItems = [];
    for (var ci = 0; ci < items.length; ci++) {
      if (isIntakeSource(items[ci].source)) {
        intakeItems.push(items[ci]);
      } else {
        sheetItems.push(items[ci]);
      }
    }

    var allCards = list.querySelectorAll(".queue-card");
    for (var i = 0; i < allCards.length; i++) {
      setCardQueued(allCards[i]);
    }

    var totalSuccess = 0;
    var totalFail = 0;
    var allArtifacts = [];
    var cancelled = false;
    var pending = (sheetItems.length > 0 ? 1 : 0) + (intakeItems.length > 0 ? 1 : 0);

    function finishBranch() {
      if (--pending > 0) return;
      setTitleSpinner("artifactsSpinner", false);
      state.generating = false;
      setGeneratingLock(false);
      qs("#cancelGenerateBtn").classList.add("hidden");
      updateViewerButton();
      var msg;
      var err = null;
      if (cancelled) {
        msg = totalSuccess > 0
          ? "Cancelled after " + totalSuccess + " artifact(s)"
          : null;
        err = totalSuccess > 0 ? null : "Generation cancelled";
      } else {
        msg = "Generated " + totalSuccess + " artifact(s)";
        if (totalFail > 0) msg += ", " + totalFail + " failed";
        if (totalSuccess === 0 && totalFail > 0) {
          msg = null;
          err = "All generations failed";
        }
      }
      showResult(msg, err);
      qs("#statusOverlay").classList.remove("hidden");
    }

    // Handle spreadsheet items via streaming api/generate
    if (sheetItems.length > 0) {
      var cellsSeen = {};
      var cells = [];
      for (var si = 0; si < sheetItems.length; si++) {
        var ck = sheetItems[si].participant + "." + sheetItems[si].row;
        if (!cellsSeen[ck]) { cellsSeen[ck] = true; cells.push(ck); }
      }

      function handleLine(line) {
        var data;
        try { data = JSON.parse(line); } catch (_) { return; }
        if (!data) return;
        if (data.cancelled) {
          cancelled = true;
          var queuedCards = list.querySelectorAll(".queue-card.queued");
          for (var qi = 0; qi < queuedCards.length; qi++) {
            clearCardStatus(queuedCards[qi]);
          }
          return;
        }
        if (!data.cell) return;
        var dot = data.cell.lastIndexOf(".");
        var participant = data.cell.substring(0, dot);
        var row = data.cell.substring(dot + 1);
        var cards = list.querySelectorAll(
          '[data-participant="' + participant + '"][data-row="' + row + '"]'
        );
        if (data.ok) {
          for (var ci = 0; ci < cards.length; ci++) setCardResult(cards[ci], true);
          totalSuccess += (data.generated || 1);
          if (data.artifacts) {
            allArtifacts = allArtifacts.concat(data.artifacts);
            state.generatedArtifacts = state.generatedArtifacts.concat(data.artifacts);
            updateViewerButton();
          }
        } else {
          for (ci = 0; ci < cards.length; ci++) setCardResult(cards[ci], false);
          totalFail++;
        }
      }

      var genBody = { cells: cells, format: format };
      if (format === "clip") {
        var tcCb = qs("#titlecardEnabled");
        var tcDur = qs("#titlecardDuration");
        if (tcCb) genBody.titlecards_enabled = tcCb.checked;
        if (tcDur) genBody.titlecard_duration = parseInt(tcDur.value, 10) || 2;
      }

      // TODO: streaming NDJSON response — not a JSON body. apiPost doesn't apply;
      // manual fetch is required to get a reader and parse line-delimited progress events.
      fetch("api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(genBody),
      })
        .then(function (response) {
          if (!response.ok) throw new Error("Server error " + response.status);
          var reader = response.body.getReader();
          var decoder = new TextDecoder();
          var buffer = "";

          function read() {
            return reader.read().then(function (result) {
              if (result.done) {
                if (buffer.trim()) handleLine(buffer.trim());
                finishBranch();
                return;
              }
              buffer += decoder.decode(result.value, { stream: true });
              var lines = buffer.split("\n");
              buffer = lines.pop();
              for (var li = 0; li < lines.length; li++) {
                if (lines[li].trim()) handleLine(lines[li].trim());
              }
              return read();
            });
          }

          return read();
        })
        .catch(function () {
          finishBranch();
        });
    }

    // Handle intake items via api/generate-intake
    if (intakeItems.length > 0) {
      var intakePayload = intakeItems.map(function (itm) {
        return {
          participant: itm.participant,
          start: itm.start,
          end: itm.end,
          event_type: itm.event_type || itm.desc || "",
          event_ids: itm.event_ids || [],
          source: itm.source || "screenspace",
          mark_ids: itm.mark_ids || [],
        };
      });

      apiPost("api/generate-intake", { items: intakePayload, format: format })
        .then(function (data) {
          var intakeCards = list.querySelectorAll('[data-source="screenspace"], [data-source="transcript"]');
          if (data.ok && data.results) {
            for (var ri = 0; ri < data.results.length; ri++) {
              var res = data.results[ri];
              if (res.ok) {
                totalSuccess++;
                if (res.artifact) {
                  allArtifacts.push(res.artifact);
                  state.generatedArtifacts.push(res.artifact);
                  updateViewerButton();
                }
                if (intakeCards[ri]) setCardResult(intakeCards[ri], true);
              } else {
                totalFail++;
                if (intakeCards[ri]) setCardResult(intakeCards[ri], false);
              }
            }
          }
          finishBranch();
        })
        .catch(function () {
          var intakeCards = list.querySelectorAll('[data-source="screenspace"], [data-source="transcript"]');
          for (var j = 0; j < intakeCards.length; j++) setCardResult(intakeCards[j], false);
          totalFail += intakeItems.length;
          finishBranch();
        });
    }
  }

  function onCancelReel() {
    qs("#cancelReelBtn").classList.add("hidden");
    apiPost("api/reel/cancel").catch(function () {});
  }

  function onCancelGenerate() {
    qs("#cancelGenerateBtn").classList.add("hidden");
    apiPost("api/generate/cancel").catch(function () {});
  }

  function onBuildReel() {
    if (state.generating || state.reelQueue.length === 0) return;
    state.generating = true;
    setGeneratingLock(true);
    setTitleSpinner("reelSpinner", true);
    qs("#cancelReelBtn").classList.remove("hidden");

    // Determine if we have intake items — use direct endpoint for mixed/intake reels
    var hasIntake = false;
    for (var ci = 0; ci < state.reelQueue.length; ci++) {
      if (isIntakeSource(state.reelQueue[ci].source)) { hasIntake = true; break; }
    }

    var reelBody;
    var endpoint;

    if (hasIntake) {
      var segments = [];
      for (var si = 0; si < state.reelQueue.length; si++) {
        var item = state.reelQueue[si];
        segments.push({
          participant: item.participant,
          start: item.start,
          end: item.end,
          source: item.source || "screenspace",
        });
      }
      reelBody = { segments: segments };
      endpoint = "api/reel-direct";
    } else {
      var cellsSeen = {};
      var cells = [];
      for (var ci2 = 0; ci2 < state.reelQueue.length; ci2++) {
        var ck = state.reelQueue[ci2].participant + "." + state.reelQueue[ci2].row;
        if (!cellsSeen[ck]) { cellsSeen[ck] = true; cells.push(ck); }
      }
      reelBody = { cells: cells };
      endpoint = "api/reel";
    }

    var tcCb = qs("#titlecardEnabled");
    var tcDur = qs("#titlecardDuration");
    if (tcCb) reelBody.titlecards_enabled = tcCb.checked;
    if (tcDur) reelBody.titlecard_duration = parseInt(tcDur.value, 10) || 2;

    var list = qs("#reelList");
    var reelCards = list.querySelectorAll(".queue-card");
    for (var i = 0; i < reelCards.length; i++) {
      setCardQueued(reelCards[i]);
    }

    apiPost(endpoint, reelBody)
      .then(function (data) {
        state.generating = false;
        setTitleSpinner("reelSpinner", false);
        setGeneratingLock(false);
        qs("#cancelReelBtn").classList.add("hidden");

        var cancelled = !!data.cancelled;
        var cards = list.querySelectorAll(".queue-card");
        for (var j = 0; j < cards.length; j++) {
          if (cancelled) {
            clearCardStatus(cards[j]);
          } else {
            setCardResult(cards[j], !!data.ok);
          }
        }

        if (cancelled) {
          showResult(null, "Reel generation cancelled");
        } else if (data.ok) {
          showResult("Reel built successfully", null);
        } else {
          showResult(null, data.error || "Reel build failed");
        }
        qs("#statusOverlay").classList.remove("hidden");
      })
      .catch(function (err) {
        state.generating = false;
        setTitleSpinner("reelSpinner", false);
        setGeneratingLock(false);
        qs("#cancelReelBtn").classList.add("hidden");

        var cards = list.querySelectorAll(".queue-card");
        for (var j = 0; j < cards.length; j++) {
          setCardResult(cards[j], false);
        }

        showResult(null, "Request failed: " + err);
        qs("#statusOverlay").classList.remove("hidden");
      });
  }

  function onBuildViewer() {
    if (state.generating || state.generatedArtifacts.length === 0) return;
    state.generating = true;

    showOverlay("Building timeline viewer...");

    apiPost("api/viewer", {})
      .then(function (data) {
        state.generating = false;
        if (data.ok) {
          showResult("Viewer created: " + (data.file || ""), null, data.file);
        } else {
          showResult(null, data.error || "Viewer build failed");
        }
      })
      .catch(function (err) {
        state.generating = false;
        showResult(null, "Request failed: " + err);
      });
  }

  function onBuildTimelineViewer() {
    if (state.generating) return;

    if (state.intakeClusters && state.intakeClusters.length > 0) {
      var n = state.intakeClusters.length;
      showConfirm(
        "Include Intake Events?",
        n + " Intake event group" + (n === 1 ? "" : "s") +
          " detected from Screenspace. Include them as clips in the timeline viewer?",
        function () { startTimelineViewerBuild(true); },
        function () { startTimelineViewerBuild(false); }
      );
    } else {
      startTimelineViewerBuild(false);
    }
  }

  function startTimelineViewerBuild(includeIntake) {
    state.generating = true;
    var body = {};

    if (includeIntake && state.intakeClusters && state.intakeClusters.length > 0) {
      showOverlay("Building timeline viewer with intake events\u2026");
      body.include_intake = true;
      body.intake_items = state.intakeClusters.map(function (c) {
        return {
          participant: c.participant,
          start: c.start,
          end: c.end,
          event_type: c.event_type,
          event_ids: c.events.map(function (e) { return e.id; }),
        };
      });
    } else {
      showOverlay("Building timeline viewer\u2026");
    }

    apiPost("api/timeline-viewer", body)
      .then(function (data) {
        state.generating = false;
        if (data.ok) {
          var msg = "Timeline viewer created: " + (data.file || "");
          if (data.generated) msg = "Generated " + data.generated + " clip(s). " + msg;
          updateViewerButton();
          showResult(msg, null, data.file);
        } else {
          showResult(null, data.error || "Timeline viewer build failed");
        }
      })
      .catch(function (err) {
        state.generating = false;
        showResult(null, "Request failed: " + err);
      });
  }

  var _highlightsBtnOrigHTML = "";

  function onBuildHighlights() {
    if (state.generating) return;

    var drawer = qs("#highlightsDurationDrawer");
    var btn = qs("#buildHighlightsBtn");
    var isOpen = drawer.classList.contains("open");

    var checkHTML = iconHTML("check", "cg-icon--confirm");

    if (!isOpen) {
      _highlightsBtnOrigHTML = btn.innerHTML;
      drawer.classList.add("open");
      var w = btn.offsetWidth;
      btn.style.minWidth = w + "px";
      btn.innerHTML = checkHTML + "Confirm";
      return;
    }

    var duration = parseInt(qs("#highlightsDuration").value, 10);
    if (!duration || duration < 1) duration = 180;

    drawer.classList.remove("open");
    btn.style.minWidth = "";
    btn.innerHTML = _highlightsBtnOrigHTML;

    state.generating = true;
    showOverlay("Finding best clips (" + duration + "s budget)...");

    apiPost("api/highlights-preview", { highlights_duration: duration })
      .then(function (data) {
        state.generating = false;
        if (data.ok && data.clips && data.clips.length > 0) {
          state.reelQueue = [];
          for (var i = 0; i < data.clips.length; i++) {
            var entries = expandCellToSegments(data.clips[i]);
            for (var ei = 0; ei < entries.length; ei++) {
              state.reelQueue.push(entries[ei]);
            }
          }
          renderReelQueue();
          updateCellClasses();
          showResult(
            "Added " + data.clips.length + " clip(s) to reel queue",
            null
          );
        } else {
          showResult(
            null,
            data.error || "No clips found for highlights selection"
          );
        }
      })
      .catch(function (err) {
        state.generating = false;
        showResult(null, "Request failed: " + err);
      });
  }

  function populateGalleryParticipants(participants) {
    var sel = qs("#galleryParticipant");
    if (!sel) return;
    sel.innerHTML = "";
    for (var i = 0; i < participants.length; i++) {
      var opt = el("option");
      opt.value = participants[i];
      opt.textContent = participants[i];
      sel.appendChild(opt);
    }
  }

  function openGalleryDialog() {
    if (state.generating) return;
    var overlay = qs("#galleryOverlay");
    if (!overlay) return;
    overlay.classList.remove("hidden");
    var sel = qs("#galleryParticipant");
    if (sel) sel.focus();
  }

  function closeGalleryDialog() {
    var overlay = qs("#galleryOverlay");
    if (overlay) overlay.classList.add("hidden");
  }

  function submitGalleryDialog() {
    if (state.generating) return;

    var participant = qs("#galleryParticipant").value;
    var format = qs("#galleryFormat").value;
    var interval = parseInt(qs("#galleryInterval").value, 10);
    if (!interval || interval < 1) interval = 10;
    var bundle = qs("#galleryBundle").checked;

    if (!participant) {
      showResult(null, "No participant selected for gallery");
      return;
    }

    closeGalleryDialog();
    state.generating = true;
    showOverlay("Generating gallery viewer for " + participant + "...");

    apiPost("api/gallery", { participant: participant, format: format, interval: interval, bundle: bundle })
      .then(function (data) {
        state.generating = false;
        if (data.ok) {
          showResult("Gallery viewer created: " + (data.file || ""), null, data.file);
        } else {
          showResult(null, data.error || "Gallery build failed");
        }
      })
      .catch(function (err) {
        state.generating = false;
        showResult(null, "Request failed: " + err);
      });
  }

  function bindGalleryDialog() {
    var overlay = qs("#galleryOverlay");
    var cancel = qs("#galleryDialogCancel");
    var confirm = qs("#galleryDialogConfirm");
    if (overlay) {
      overlay.addEventListener("click", function (ev) {
        if (ev.target === overlay) closeGalleryDialog();
      });
    }
    if (cancel) cancel.addEventListener("click", closeGalleryDialog);
    if (confirm) confirm.addEventListener("click", submitGalleryDialog);
    document.addEventListener("keydown", function (ev) {
      if (ev.key === "Escape" && overlay && !overlay.classList.contains("hidden")) {
        closeGalleryDialog();
      }
    });
  }

  function updateViewerButton() {
    var count = qs("#viewerArtifactCount");
    var n = state.generatedArtifacts.length;
    count.textContent = n > 0 ? n + " artifact(s) ready" : "";
  }

  // ---- Status overlay ----

  var _lastViewerFile = "";

  function showOverlay(message) {
    qs("#statusSpinner").style.display = "";
    qs("#statusTitle").textContent = message;
    qs("#statusMessage").textContent = "";
    qs("#statusMessage").className = "";
    qs("#statusDismiss").classList.add("hidden");
    qs("#statusOpen").classList.add("hidden");
    _lastViewerFile = "";
    qs("#statusOverlay").classList.remove("hidden");
  }

  function showResult(successMsg, errorMsg, filePath) {
    qs("#statusSpinner").style.display = "none";
    if (errorMsg) {
      qs("#statusTitle").textContent = "Error";
      qs("#statusMessage").textContent = errorMsg;
      qs("#statusMessage").className = "error-text";
      qs("#statusOpen").classList.add("hidden");
    } else {
      qs("#statusTitle").textContent = "Done";
      qs("#statusMessage").textContent = successMsg || "";
      qs("#statusMessage").className = "";
      if (filePath) {
        _lastViewerFile = filePath;
        qs("#statusOpen").classList.remove("hidden");
      } else {
        qs("#statusOpen").classList.add("hidden");
      }
    }
    qs("#statusDismiss").classList.remove("hidden");
  }

  function hideOverlay() {
    qs("#statusOverlay").classList.add("hidden");
  }

  // ---- Confirm overlay ----

  var _confirmCleanup = null;

  function showConfirm(title, message, onYes, onNo) {
    qs("#confirmTitle").textContent = title;
    qs("#confirmMessage").textContent = message;
    qs("#confirmOverlay").classList.remove("hidden");

    var yesBtn = qs("#confirmYes");
    var noBtn = qs("#confirmNo");

    function cleanup() {
      qs("#confirmOverlay").classList.add("hidden");
      yesBtn.removeEventListener("click", handleYes);
      noBtn.removeEventListener("click", handleNo);
      _confirmCleanup = null;
    }

    function handleYes() { cleanup(); onYes(); }
    function handleNo() { cleanup(); onNo(); }

    yesBtn.addEventListener("click", handleYes);
    noBtn.addEventListener("click", handleNo);
    _confirmCleanup = cleanup;
  }

  function hideConfirm() {
    if (_confirmCleanup) _confirmCleanup();
    else qs("#confirmOverlay").classList.add("hidden");
  }

  // ---- Artifact log ----

  function openLog() {
    qs("#logOverlay").classList.remove("hidden");
    renderLog();
  }

  function closeLog() {
    qs("#logOverlay").classList.add("hidden");
  }

  function renderLog() {
    var container = qs("#logContent");
    var countEl = qs("#logCount");
    var items = state.generatedArtifacts;

    if (items.length === 0) {
      container.innerHTML = "";
      container.appendChild(el("div", "log-empty", "No artifacts generated yet."));
      countEl.textContent = "";
      return;
    }

    container.innerHTML = "";
    for (var i = items.length - 1; i >= 0; i--) {
      var a = items[i];
      var row = el("div", "log-entry");

      var badge = el("span", "log-type-badge", a.type || "clip");
      badge.setAttribute("data-type", a.type || "clip");
      row.appendChild(badge);

      row.appendChild(el("span", "log-entry-file", a.file || ""));

      var meta = [];
      if (a.participant) meta.push(a.participant);
      if (a.description) {
        var desc = a.description.length > 40 ? a.description.slice(0, 40) + "\u2026" : a.description;
        meta.push(desc);
      }
      if (meta.length > 0) {
        row.appendChild(el("span", "log-entry-meta", meta.join(" \u00B7 ")));
      }

      container.appendChild(row);
    }

    var n = items.length;
    countEl.textContent = n + " artifact" + (n !== 1 ? "s" : "");
  }

  // ---- Settings (shared modal lives in settings-modal.js) ----

  function _findSetting(name) {
    if (!state.settingsData) return null;
    for (var i = 0; i < state.settingsData.length; i++) {
      if (state.settingsData[i].name === name) return state.settingsData[i];
    }
    return null;
  }

  function syncInlineControls() {
    if (!state.settingsData) return;
    var tcEnabled = _findSetting("TITLECARDS_ENABLED");
    var tcDuration = _findSetting("TITLECARD_DURATION_SECONDS");
    var hlDuration = _findSetting("HIGHLIGHTS_REEL_DURATION_SECONDS");

    if (tcEnabled) {
      var cb = qs("#titlecardEnabled");
      if (cb) cb.checked = !!tcEnabled.value;
    }
    if (tcDuration) {
      var dur = qs("#titlecardDuration");
      if (dur) dur.value = tcDuration.value;
    }
    if (hlDuration) {
      var hl = qs("#highlightsDuration");
      if (hl) hl.value = hlDuration.value;
    }
    var cellExpand = _findSetting("STUDIO_CELL_EXPAND_HOVER");
    if (cellExpand) {
      state.cellExpandHover = !!cellExpand.value;
    }
  }

  function persistTitlecardSettings() {
    var cb = qs("#titlecardEnabled");
    var dur = qs("#titlecardDuration");
    if (!cb || !dur) return;

    var payload = {
      TITLECARDS_ENABLED: cb.checked,
      TITLECARD_DURATION_SECONDS: parseInt(dur.value, 10) || 2,
    };

    apiPut("api/settings", { settings: payload })
      .then(function (data) {
        if (data.ok && state.settingsData) {
          var tcE = _findSetting("TITLECARDS_ENABLED");
          var tcD = _findSetting("TITLECARD_DURATION_SECONDS");
          if (tcE) tcE.value = cb.checked;
          if (tcD) tcD.value = parseInt(dur.value, 10) || 2;
        }
      })
      .catch(function () {});
  }

  // ---- Init ----

  function checkConvergenceTabVisibility() {
    var tab = qs('.preview-tab[data-tab="convergence"]');
    if (!tab) return;
    var multipleParticipants = false;
    if (state.sheetData && state.sheetData.participants && state.sheetData.participants.length > 1) {
      multipleParticipants = true;
    }
    if (!multipleParticipants && state.intakeEvents.length > 0) {
      var seenSS = {};
      for (var i = 0; i < state.intakeEvents.length; i++) {
        seenSS[state.intakeEvents[i].participant] = true;
      }
      if (Object.keys(seenSS).length > 1) multipleParticipants = true;
    }
    if (!multipleParticipants && state.trIntakeMarks.length > 0) {
      var seenTr = {};
      for (var j = 0; j < state.trIntakeMarks.length; j++) {
        seenTr[state.trIntakeMarks[j].participant] = true;
      }
      if (Object.keys(seenTr).length > 1) multipleParticipants = true;
    }
    if (multipleParticipants) {
      tab.classList.remove("hidden");
    } else {
      tab.classList.add("hidden");
    }
    restoreStoredPreviewTab();
  }

  function checkNavLinks() {
    apiGet("../api/status")
      .then(function (data) {
        if (data.screenspace) {
          var intakeTab = qs('.preview-tab[data-tab="intake"]');
          if (intakeTab) intakeTab.classList.remove("hidden");
        }
        if (data.transcripts) {
          var trIntakeTab = qs('.preview-tab[data-tab="transcript-intake"]');
          if (trIntakeTab) trIntakeTab.classList.remove("hidden");
        }
        restoreStoredPreviewTab();
      })
      .catch(function () {});
  }

  // ---- Screenspace Intake ----

  var INTAKE_DETECTOR_COLORS = DETECTOR_COLORS;
  var INTAKE_DETECTORS = [
    "multitool", "color", "change", "similarity", "text",
    "numbers", "timelapse", "template", "flow", "scene", "inactivity",
  ];

  // ---- Screenspace thumbnail queue (throttled + cached) ----

  var _ssThumbQueue = [];
  var _ssThumbActive = 0;
  var _SS_THUMB_MAX = 3;
  var _ssThumbCache = {}; // url -> objectURL | "error"

  function ssThumbUrl(participant, timestamp) {
    return "../screenspace/api/video/frame/" + encodeURIComponent(participant) + "/" + timestamp + "?w=200";
  }

  // Thumb element selector covers the legacy `.queue-card-thumb` (artifact /
  // reel cards in the bottom strip) and the new primitive cards
  // (`.clip-card-thumb`, `.transcript-card-thumb`) used by Studio Intake.
  var SS_THUMB_SELECTOR = ".queue-card-thumb, .clip-card-thumb, .transcript-card-thumb";

  function ssProcessQueue() {
    while (_ssThumbActive < _SS_THUMB_MAX && _ssThumbQueue.length) {
      var item = _ssThumbQueue.shift();
      if (!item.img.parentNode) continue;
      _ssThumbActive++;
      (function (entry) {
        // TODO: returns a blob (image), not JSON. apiGet doesn't cover blob responses.
        fetch(entry.url)
          .then(function (r) {
            if (!r.ok) throw new Error("status " + r.status);
            return r.blob();
          })
          .then(function (blob) {
            var objUrl = URL.createObjectURL(blob);
            _ssThumbCache[entry.url] = objUrl;
            if (entry.img.parentNode) entry.img.src = objUrl;
          })
          .catch(function () {
            _ssThumbCache[entry.url] = "error";
            if (entry.img.parentNode) {
              entry.img.remove();
              entry.thumbEl.appendChild(el("span", "", "\u2715"));
              entry.cardEl.classList.add("queue-card-error");
            }
          })
          .then(function () {
            _ssThumbActive--;
            ssProcessQueue();
          });
      })(item);
    }
  }

  function ssEnqueueThumb(img, cardEl, thumbEl, participant, timestamp) {
    var url = ssThumbUrl(participant, timestamp);
    var cached = _ssThumbCache[url];
    if (cached && cached !== "error") { img.src = cached; return; }
    if (cached === "error") {
      img.remove();
      thumbEl.appendChild(el("span", "", "\u2715"));
      cardEl.classList.add("queue-card-error");
      return;
    }
    _ssThumbQueue.push({ img: img, cardEl: cardEl, thumbEl: thumbEl, url: url });
    ssProcessQueue();
  }

  var _ssThumbObservers = {};

  function ssGetObserver(root) {
    var key = root ? (root.id || "anon") : "viewport";
    if (_ssThumbObservers[key]) return _ssThumbObservers[key];
    var obs = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (!entry.isIntersecting) return;
        obs.unobserve(entry.target);
        var d = entry.target.dataset;
        var tEl = entry.target.querySelector(SS_THUMB_SELECTOR);
        var imgEl = tEl ? tEl.querySelector("img") : null;
        if (imgEl && tEl) ssEnqueueThumb(imgEl, entry.target, tEl, d.ssThumbPid, d.ssThumbTs);
      });
    }, { root: root || null, rootMargin: "200px 0px" });
    _ssThumbObservers[key] = obs;
    return obs;
  }

  function ssObserveThumb(card, img, thumbEl, participant, timestamp) {
    card.dataset.ssThumbPid = participant;
    card.dataset.ssThumbTs = timestamp;
    var scrollParent = card.closest(".intake-cards-grid") || card.closest("#artifactsList") || card.closest("#reelList");
    ssGetObserver(scrollParent).observe(card);
  }

  function ssClearPending() {
    _ssThumbQueue = [];
  }

  function clusterIntakeEvents(events, thresholdSec) {
    if (!events.length) return [];
    var sorted = events.slice().sort(function (a, b) {
      if (a.participant !== b.participant) return a.participant < b.participant ? -1 : 1;
      if (a.event_type !== b.event_type) return a.event_type < b.event_type ? -1 : 1;
      return a.time_in - b.time_in;
    });
    var clusters = [];
    var cur = null;
    for (var i = 0; i < sorted.length; i++) {
      var ev = sorted[i];
      if (
        !cur ||
        ev.participant !== cur.participant ||
        ev.event_type !== cur.event_type ||
        ev.time_in - cur.end > thresholdSec
      ) {
        if (cur) clusters.push(cur);
        cur = {
          participant: ev.participant,
          source_video: ev.source_video,
          start: ev.time_in,
          end: ev.time_out,
          event_type: ev.event_type,
          detector: ev.detector,
          region: ev.region,
          events: [ev],
          confidence_avg: ev.confidence,
        };
      } else {
        cur.end = Math.max(cur.end, ev.time_out);
        cur.events.push(ev);
        var sum = 0;
        for (var j = 0; j < cur.events.length; j++) sum += cur.events[j].confidence;
        cur.confidence_avg = sum / cur.events.length;
      }
    }
    if (cur) clusters.push(cur);
    for (var k = 0; k < clusters.length; k++) {
      var c = clusters[k];
      if (c.start === c.end) {
        c.start = Math.max(0, c.start - 5);
        c.end = c.end + 5;
      }
    }
    return clusters;
  }

  function setTabDot(elId, on) {
    var el = document.getElementById(elId);
    if (!el) return;
    if (on) el.classList.remove("hidden");
    else el.classList.add("hidden");
  }

  function pollIntakeStatus() {
    apiGet("../screenspace/api/tasks")
      .then(function (data) {
        if (!data || !data.ok) {
          setTabDot("intakeTabDot", false);
          return;
        }
        var tasks = data.tasks || [];
        var running = false;
        var hasQueued = false;
        for (var i = 0; i < tasks.length; i++) {
          if (tasks[i].status === "running") { running = true; break; }
          if (tasks[i].status === "queued") hasQueued = true;
        }
        setTabDot("intakeTabDot", running || (data.worker_alive && hasQueued));
      })
      .catch(function () {
        setTabDot("intakeTabDot", false);
      });
  }

  function pollTrIntakeStatus() {
    var statusP = apiGet("../transcripts/api/transcribe/status").catch(function () { return null; });
    var modelP = apiGet("../transcripts/api/transcribe/model-status").catch(function () { return null; });
    var partsP = apiGet("../transcripts/api/participants").catch(function () { return null; });
    Promise.all([statusP, modelP, partsP]).then(function (results) {
      var status = results[0];
      var model = results[1];
      var parts = results[2];
      var running = false;
      if (status && status.ok && Array.isArray(status.tasks)) {
        for (var i = 0; i < status.tasks.length; i++) {
          if (status.tasks[i].status === "running") { running = true; break; }
        }
      }
      if (!running && model && model.ok && model.warming) running = true;
      if (!running && parts && parts.ok && Array.isArray(parts.participants)) {
        for (var j = 0; j < parts.participants.length; j++) {
          var agents = parts.participants[j].agents || {};
          if (agents.summary === "running" || agents.citations === "running") {
            running = true;
            break;
          }
        }
      }
      setTabDot("trIntakeTabDot", running);
    });
  }

  function pollIntakeEvents() {
    apiGet("../screenspace/api/events?excluded=false")
      .then(function (data) {
        if (!data.ok) return;
        var events = data.events || [];
        var hasNew = false;
        events.forEach(function (ev) {
          if (!state.intakeSeenIds[ev.id]) {
            state.intakeSeenIds[ev.id] = "new";
            hasNew = true;
          }
        });
        state.intakeEvents = events;
        var threshold = parseInt((qs("#intakeClusterThreshold") || {}).value) || 10;
        state.intakeClusters = clusterIntakeEvents(events, threshold);
        renderIntake(hasNew);
        checkConvergenceTabVisibility();
      })
      .catch(function (err) {
        console.warn("[Intake] poll failed:", err);
      });
  }

  var XREF_ICON_BASE = "../screenspace/icons/";

  function xrefBadgeIcon(iconName) {
    var span = el("span", "xref-badge-icon");
    applyMaskIcon(span, 'url("' + XREF_ICON_BASE + iconName + '.svg")');
    return span;
  }

  // selfBadge: optional { icon, color, title } to prepend as the "self" source badge
  function buildXrefBadges(xref, selfSource, selfBadge) {
    var badges = [];
    if (selfBadge) badges.push(selfBadge);
    if (selfSource !== "screenspace" && xref.screenspaceEvents.length > 0) {
      var types = [];
      var seen = {};
      for (var i = 0; i < xref.screenspaceEvents.length; i++) {
        var et = xref.screenspaceEvents[i].event_type || xref.screenspaceEvents[i].detector;
        if (!seen[et]) { seen[et] = true; types.push(et); }
      }
      badges.push({ icon: XREF_BADGES.screenspace.icon, color: XREF_BADGES.screenspace.color, title: types.join(", ") });
    }
    if (selfSource !== "transcript" && xref.transcriptSnippets.length > 0) {
      var trTexts = [];
      for (var j = 0; j < xref.transcriptSnippets.length && j < 3; j++) {
        var t = xref.transcriptSnippets[j].text;
        trTexts.push(t.length > 80 ? t.substring(0, 80) + "\u2026" : t);
      }
      badges.push({ icon: XREF_BADGES.transcript.icon, color: XREF_BADGES.transcript.color, title: trTexts.join("\n") });
    }
    if (xref.sheetObservations.length > 0) {
      var obsTexts = [];
      for (var k = 0; k < xref.sheetObservations.length && k < 3; k++) {
        obsTexts.push(xref.sheetObservations[k].observation);
      }
      badges.push({ icon: XREF_BADGES.sheet.icon, color: XREF_BADGES.sheet.color, title: obsTexts.join("\n") });
    }
    if (badges.length === 0) return null;
    var container = el("span", "xref-badge-stack");
    for (var b = 0; b < badges.length; b++) {
      var badge = el("span", "xref-badge");
      badge.style.background = badges[b].color;
      badge.style.zIndex = badges.length - b;
      badge.appendChild(xrefBadgeIcon(badges[b].icon));
      badge.title = badges[b].title;
      container.appendChild(badge);
    }
    return container;
  }

  function highlightIntakeCard(idx) {
    var cards = qsa(".intake-queue-card");
    for (var i = 0; i < cards.length; i++) {
      if (i === idx) {
        cards[i].classList.add("intake-highlight");
      } else {
        cards[i].classList.remove("intake-highlight");
      }
    }
  }

  function buildIntakeDetectorPills() {
    var container = qs("#intakeDetectorPills");
    if (!container) return;
    var counts = {};
    for (var i = 0; i < state.intakeClusters.length; i++) {
      var d = state.intakeClusters[i].detector;
      if (d) counts[d] = (counts[d] || 0) + 1;
    }
    container.innerHTML = "";
    INTAKE_DETECTORS.forEach(function (det) {
      if (!counts[det]) return;
      var chip = ClipgenPrimitives.createFilterChip({
        label: det,
        active: state.intakeFilterDetector === det,
        count: counts[det],
        hue: categoryHue(det),
        // Pin detector chips to the canonical `--color-task-*` token so the
        // dot matches Screenspace's workflow tab + result row exactly.
        color: detectorColor(det),
        onClick: function () {
          state.intakeFilterDetector = state.intakeFilterDetector === det ? "" : det;
          renderIntake(false);
        },
      });
      container.appendChild(chip);
    });
  }

  function buildIntakeParticipantPills() {
    var container = qs("#intakeFilterParticipants");
    if (!container) return;
    var seen = {};
    var participants = [];
    for (var i = 0; i < state.intakeClusters.length; i++) {
      var p = state.intakeClusters[i].participant;
      if (p && !seen[p]) { seen[p] = true; participants.push(p); }
    }
    participants.sort();
    state.intakeFilterParticipants = state.intakeFilterParticipants.filter(function (p) { return seen[p]; });
    container.innerHTML = "";
    participants.forEach(function (p) {
      var pill = ClipgenPrimitives.createParticipantPill({
        id: p,
        active: state.intakeFilterParticipants.indexOf(p) !== -1,
        onClick: function () {
          var idx = state.intakeFilterParticipants.indexOf(p);
          if (idx === -1) state.intakeFilterParticipants.push(p);
          else state.intakeFilterParticipants.splice(idx, 1);
          renderIntake(false);
        },
      });
      container.appendChild(pill);
    });
  }

  var _intakeDensityEl = null;

  function buildIntakeDensityTimeline(clusters) {
    var host = qs("#intakeTimeline");
    if (!host) return;
    host.innerHTML = "";
    _intakeDensityEl = null;
    if (!clusters.length) return;
    var maxEnd = 0;
    for (var i = 0; i < clusters.length; i++) {
      if (clusters[i].end > maxEnd) maxEnd = clusters[i].end;
    }
    var duration = Math.max(maxEnd * 1.05, 60);
    var events = clusters.map(function (c) {
      return {
        t: duration > 0 ? c.start / duration : 0,
        count: c.events ? c.events.length : 1,
        hue: categoryHue(c.detector),
        // Density bars match Screenspace's `--color-task-*` exactly when the
        // cluster's detector is one of the known types; falls back to the
        // hue-based oklch path otherwise.
        color: detectorColor(c.detector),
      };
    });
    var dt = ClipgenPrimitives.createDensityTimeline({
      events: events,
      durationSec: duration,
      tickCount: 6,
      onBarMouseEnter: function (idx) {
        state.intakeHoveredIdx = idx;
        highlightIntakeCard(idx);
        if (_intakeDensityEl) _intakeDensityEl.setHovered(idx);
      },
      onBarMouseLeave: function () {
        state.intakeHoveredIdx = -1;
        highlightIntakeCard(-1);
        if (_intakeDensityEl) _intakeDensityEl.setHovered(-1);
      },
      onBarClick: function (idx, ev) {
        var cluster = filteredIntakeClusters()[idx];
        if (!cluster) return;
        if (ev && ev.shiftKey) intakeAddToReel(cluster);
        else intakeAddToArtifacts(cluster);
      },
    });
    _intakeDensityEl = dt;
    host.appendChild(dt);
  }

  function renderIntake(_hasNew) {
    ssClearPending();
    var container = qs("#intakeCards");
    var addAllBtn = qs("#intakeAddAllBtn");
    var reelAllBtn = qs("#intakeReelAllBtn");
    var tabBadge = qs("#intakeTabBadge");

    buildIntakeDetectorPills();
    buildIntakeParticipantPills();

    if (!state.intakeClusters.length) {
      if (tabBadge) tabBadge.classList.add("hidden");
      container.innerHTML = "";
      container.appendChild(el("div", "drop-target-empty", "Screenspace events will appear here"));
      addAllBtn.disabled = true;
      reelAllBtn.disabled = true;
      buildIntakeDensityTimeline([]);
      return;
    }
    if (tabBadge) {
      tabBadge.textContent = state.intakeClusters.length;
      tabBadge.classList.remove("hidden");
    }
    var clusters = filteredIntakeClusters();
    addAllBtn.disabled = clusters.length === 0;
    reelAllBtn.disabled = clusters.length === 0;

    buildIntakeDensityTimeline(clusters);

    container.innerHTML = "";
    if (clusters.length === 0) {
      container.appendChild(el("div", "drop-target-empty", "No events match the current filters"));
      return;
    }
    clusters.forEach(function (c, idx) {
      var segDuration = c.end - c.start;
      var hueLabel = c.event_type || c.detector || "uncategorized";

      var card = ClipgenPrimitives.createClipCard({
        participant: c.participant,
        duration: formatDuration(segDuration),
        label: c.event_type || c.detector || "intake",
        hue: categoryHue(hueLabel),
        // Pin the hue dot + label to the canonical `--color-task-*` token
        // when the event is a known detector — otherwise keep the oklch
        // fallback so non-detector labels still get a stable colour.
        color: detectorColor(c.detector),
        size: "lg",
        dataset: { intakeIdx: idx },
        onDragStart: function (ev) {
          ev.dataTransfer.setData("application/json", JSON.stringify({
            participant: c.participant,
            desc: c.event_type,
            start: c.start,
            end: c.end,
            source: "screenspace",
            event_type: c.event_type,
            event_ids: c.events.map(function (e) { return e.id; }),
          }));
          ev.dataTransfer.effectAllowed = "copyMove";
          setCardDragImage(ev, this);
        },
      });
      card.classList.add("intake-queue-card");

      // Lazy-loaded source-frame thumbnail (existing observer integration).
      var thumb = card.querySelector(".clip-card-thumb");
      var img = document.createElement("img");
      img.alt = "";
      img.draggable = false;
      thumb.insertBefore(img, thumb.firstChild);
      ssObserveThumb(card, img, thumb, c.participant, c.start);

      // Cross-reference badges layered over the thumb.
      var xref = findOverlappingData(c.participant, c.start, c.end);
      var ssSelf = { icon: XREF_BADGES.screenspace.icon, color: XREF_BADGES.screenspace.color, title: c.event_type || "Screenspace" };
      var badgeStack = buildXrefBadges(xref, "screenspace", ssSelf);
      if (badgeStack) thumb.appendChild(badgeStack);
      if (xref.transcriptSnippets.length > 0) {
        card.dataset.transcriptContext = xref.transcriptSnippets.map(function (s) { return s.text; }).join("\n");
      }

      container.appendChild(card);
    });
  }

  function intakeAddToArtifacts(cluster) {
    state.artifactQueue.push({
      participant: cluster.participant,
      start: cluster.start,
      end: cluster.end,
      desc: cluster.event_type,
      source: "screenspace",
      event_type: cluster.event_type,
      event_ids: cluster.events.map(function (e) { return e.id; }),
    });
    renderArtifactQueue();
  }

  function intakeDismissCluster(cluster) {
    var ids = cluster.events.map(function (e) { return e.id; });
    apiPut("../screenspace/api/events/bulk-exclude", { ids: ids })
      .then(function () { pollIntakeEvents(); })
      .catch(function () {});
  }

  function intakeAddToReel(cluster) {
    state.reelQueue.push({
      participant: cluster.participant,
      start: cluster.start,
      end: cluster.end,
      desc: cluster.event_type,
      source: "screenspace",
      event_type: cluster.event_type,
      event_ids: cluster.events.map(function (e) { return e.id; }),
    });
    renderReelQueue();
  }

  function filteredIntakeClusters() {
    var clusters = state.intakeClusters;
    var text = state.intakeFilterText.toLowerCase();
    var det = state.intakeFilterDetector;
    var onlyNew = state.intakeFilterNew;
    var parts = state.intakeFilterParticipants;
    if (!text && !det && !onlyNew && !parts.length) return clusters;
    return clusters.filter(function (c) {
      if (parts.length && parts.indexOf(c.participant) === -1) return false;
      if (det && c.detector !== det) return false;
      if (text && (c.event_type || "").toLowerCase().indexOf(text) === -1
          && (c.region || "").toLowerCase().indexOf(text) === -1
          && (c.participant || "").toLowerCase().indexOf(text) === -1) return false;
      if (onlyNew) {
        var hasNew = false;
        for (var i = 0; i < c.events.length; i++) {
          if (state.intakeSeenIds[c.events[i].id] === "new") { hasNew = true; break; }
        }
        if (!hasNew) return false;
      }
      return true;
    });
  }

  function initIntake() {
    var intakeCards = qs("#intakeCards");

    // Click: normal = Artifacts, shift = Reel
    intakeCards.addEventListener("click", function (e) {
      var card = e.target.closest(".intake-queue-card");
      if (!card) return;
      var idx = parseInt(card.dataset.intakeIdx);
      var cluster = filteredIntakeClusters()[idx];
      if (!cluster) return;
      if (e.shiftKey) intakeAddToReel(cluster);
      else intakeAddToArtifacts(cluster);
    });

    // Right-click to dismiss
    intakeCards.addEventListener("contextmenu", function (e) {
      var card = e.target.closest(".intake-queue-card");
      if (!card) return;
      e.preventDefault();
      var idx = parseInt(card.dataset.intakeIdx);
      var cluster = filteredIntakeClusters()[idx];
      if (cluster) intakeDismissCluster(cluster);
    });

    // Card hover → highlight + transcript tooltip + timeline marker
    intakeCards.addEventListener("mouseover", function (e) {
      var card = e.target.closest(".intake-queue-card");
      if (!card) return;
      var idx = parseInt(card.dataset.intakeIdx);
      if (state.intakeHoveredIdx !== idx) {
        state.intakeHoveredIdx = idx;
        highlightIntakeCard(idx);
        if (_intakeDensityEl) _intakeDensityEl.setHovered(idx);
      }
      var trTooltip = qs("#trIntakeTooltip");
      if (trTooltip && state.trIntakeTooltipsEnabled) {
        var tooltipText = card.dataset.transcriptContext || "";
        if (tooltipText) {
          trTooltip.textContent = tooltipText;
          trTooltip.classList.remove("hidden");
          positionTooltipAnchored(trTooltip, card.getBoundingClientRect());
        } else {
          trTooltip.classList.add("hidden");
        }
      }
    });
    intakeCards.addEventListener("mouseleave", function () {
      if (state.intakeHoveredIdx !== -1) {
        state.intakeHoveredIdx = -1;
        highlightIntakeCard(-1);
        if (_intakeDensityEl) _intakeDensityEl.setHovered(-1);
      }
      var trTooltip = qs("#trIntakeTooltip");
      if (trTooltip) trTooltip.classList.add("hidden");
    });

    qs("#intakeAddAllBtn").addEventListener("click", function () {
      filteredIntakeClusters().forEach(function (c) { intakeAddToArtifacts(c); });
    });
    qs("#intakeReelAllBtn").addEventListener("click", function () {
      filteredIntakeClusters().forEach(function (c) { intakeAddToReel(c); });
    });
    qs("#intakeClusterThreshold").addEventListener("change", function () {
      var threshold = parseInt(this.value) || 5;
      state.intakeClusters = clusterIntakeEvents(state.intakeEvents, threshold);
      renderIntake(false);
    });

    // Search + "New only" toggle
    var searchEl = qs("#intakeFilterSearch");
    var _intakeSearchTimer = 0;
    if (searchEl) {
      searchEl.addEventListener("input", function () {
        state.intakeFilterText = this.value;
        clearTimeout(_intakeSearchTimer);
        _intakeSearchTimer = setTimeout(function () { renderIntake(false); }, 250);
      });
    }
    var newToggle = qs("#intakeFilterNew");
    if (newToggle) {
      newToggle.addEventListener("change", function () {
        state.intakeFilterNew = this.checked;
        renderIntake(false);
      });
    }

    // Polling runs whenever the page is visible so the intake-tab counter
    // badge stays live regardless of which sub-tab the user is currently on.
    document.addEventListener("visibilitychange", function () {
      if (document.hidden) {
        if (state.intakePollTimer) { clearInterval(state.intakePollTimer); state.intakePollTimer = null; }
      } else {
        pollIntakeEvents();
        if (!state.intakePollTimer) state.intakePollTimer = setInterval(pollIntakeEvents, 10000);
      }
    });
  }

  // ---- Transcript Intake ----

  // Colors resolve to CSS tokens (see tokens.css `--cat-*`) so dark mode tracks the theme.
  var TR_INTAKE_CATEGORIES = {
    pain_point: { label: "Pain Point", token: "--cat-pain-point" },
    delight:    { label: "Delight",    token: "--cat-delight" },
    quote:      { label: "Quote",      token: "--cat-quote" },
    insight:    { label: "Insight",    token: "--cat-insight" },
    task:       { label: "Task Issue", token: "--cat-task" },
    bookmark:   { label: "Bookmark",   token: "--cat-bookmark" },
  };

  function trIntakeCategoryColor(key) {
    var entry = TR_INTAKE_CATEGORIES[key] || TR_INTAKE_CATEGORIES.bookmark;
    return getCSSVar(entry.token, "");
  }

  function _syncMarkCategoriesFromSettings(settings) {
    if (!settings) return;
    for (var i = 0; i < settings.length; i++) {
      if (settings[i].name === "MARK_CATEGORIES" && settings[i].value) {
        setMarkCategories(settings[i].value);
        return;
      }
    }
  }

  function pollTranscriptIntakeMarks() {
    apiGet("../transcripts/api/marks")
      .then(function (data) {
        if (!data.ok) return;
        if (data.categories) setMarkCategories(data.categories);
        state.trIntakeMarks = data.marks.filter(function (m) { return m.valid; });
        var threshold = parseInt((qs("#trIntakeClusterThreshold") || {}).value) || 10;
        state.trIntakeClusters = clusterTranscriptMarks(state.trIntakeMarks, threshold);

        // If "Show all" is enabled, also fetch all segments as unmark items
        if (state.trIntakeShowAll) {
          apiGet("../transcripts/api/participants")
            .then(function (pData) {
              if (!pData.ok) return;
              var transcribed = pData.participants.filter(function (p) { return p.has_transcript; });
              var promises = transcribed.map(function (p) {
                return apiGet("../transcripts/api/transcript/" + p.id);
              });
              Promise.all(promises).then(function (results) {
                var markedIds = {};
                for (var i = 0; i < state.trIntakeMarks.length; i++) markedIds[state.trIntakeMarks[i].segment_id] = true;
                var allItems = state.trIntakeMarks.slice();
                for (var j = 0; j < results.length; j++) {
                  if (!results[j].ok) continue;
                  var pid = results[j].participant;
                  var segs = results[j].segments;
                  for (var k = 0; k < segs.length; k++) {
                    if (!markedIds[segs[k].id]) {
                      allItems.push({
                        id: null,
                        segment_id: segs[k].id,
                        category: null,
                        label: null,
                        valid: true,
                        participant: pid,
                        start: segs[k].start,
                        end: segs[k].end,
                        text: segs[k].text,
                      });
                    }
                  }
                }
                state.trIntakeClusters = clusterTranscriptMarks(allItems, threshold);
                renderTranscriptIntake();
                checkConvergenceTabVisibility();
              });
            })
            .catch(function () {});
        } else {
          renderTranscriptIntake();
          checkConvergenceTabVisibility();
        }
      })
      .catch(function () {});
  }

  function clusterTranscriptMarks(marks, thresholdSec) {
    if (!marks.length) return [];
    var sorted = marks.slice().sort(function (a, b) {
      if (a.participant !== b.participant) return a.participant < b.participant ? -1 : 1;
      return a.start - b.start;
    });
    var clusters = [];
    var cur = null;
    for (var i = 0; i < sorted.length; i++) {
      var m = sorted[i];
      if (!cur || m.participant !== cur.participant || m.start - cur.end > thresholdSec) {
        if (cur) clusters.push(cur);
        cur = {
          participant: m.participant,
          start: m.start,
          end: m.end,
          marks: [m],
          category: m.category || "bookmark",
          label: m.label || "",
          text: m.text || "",
        };
      } else {
        cur.end = Math.max(cur.end, m.end);
        cur.marks.push(m);
        if (m.text) cur.text += " " + m.text;
        if (m.label && !cur.label) cur.label = m.label;
      }
    }
    if (cur) clusters.push(cur);
    return clusters;
  }

  function filteredTranscriptIntakeClusters() {
    var clusters = state.trIntakeClusters;
    var cat = state.trIntakeFilterCategory;
    var parts = state.trIntakeFilterParticipants;
    var text = state.trIntakeFilterText.toLowerCase();
    if (!cat && !parts.length && !text) return clusters;
    return clusters.filter(function (c) {
      if (parts.length && parts.indexOf(c.participant) === -1) return false;
      if (cat && c.category !== cat) return false;
      if (text && (c.text || "").toLowerCase().indexOf(text) === -1
          && (c.label || "").toLowerCase().indexOf(text) === -1
          && (c.participant || "").toLowerCase().indexOf(text) === -1) return false;
      return true;
    });
  }

  function renderTranscriptIntake() {
    ssClearPending();
    var filtered = filteredTranscriptIntakeClusters();
    var container = qs("#trIntakeCards");
    var addAllBtn = qs("#trIntakeAddAllBtn");
    var reelAllBtn = qs("#trIntakeReelAllBtn");
    var badge = qs("#trIntakeTabBadge");

    if (badge) {
      if (state.trIntakeMarks.length > 0) {
        badge.textContent = state.trIntakeMarks.length;
        badge.classList.remove("hidden");
      } else {
        badge.classList.add("hidden");
      }
    }

    if (addAllBtn) addAllBtn.disabled = filtered.length === 0;
    if (reelAllBtn) reelAllBtn.disabled = filtered.length === 0;

    buildTrIntakeCategoryPills();
    buildTrIntakeParticipantPills();
    buildTrIntakeDensityTimeline(filtered);

    if (filtered.length === 0) {
      container.innerHTML = '<div class="drop-target-empty">Transcript marks will appear here</div>';
      return;
    }

    container.innerHTML = "";
    filtered.forEach(function (c, i) {
      var segDuration = Math.max(0, c.end - c.start);
      var labelKey = c.category || "bookmark";
      var labelText = (TR_INTAKE_CATEGORIES[labelKey] || TR_INTAKE_CATEGORIES.bookmark).label;
      var rangeStr = formatDuration(c.start) + "\u2013" + formatDuration(c.end);
      var snippet = c.label || c.text || "";

      var card = ClipgenPrimitives.createTranscriptCard({
        participant: c.participant,
        timeRange: rangeStr,
        duration: formatDuration(segDuration),
        label: labelText,
        text: snippet,
        hue: categoryHue(labelKey),
        dataset: { trIntakeIdx: i },
        onDragStart: function (ev) {
          ev.dataTransfer.setData("application/json", JSON.stringify({
            participant: c.participant,
            desc: c.category || "transcript",
            start: c.start,
            end: c.end,
            source: "transcript",
            mark_ids: c.marks.map(function (m) { return m.id; }),
          }));
          ev.dataTransfer.effectAllowed = "copyMove";
          setCardDragImage(ev, this);
        },
      });
      card.classList.add("intake-queue-card", "tr-intake-queue-card");

      // Lazy-loaded source-frame thumbnail
      var thumb = card.querySelector(".transcript-card-thumb");
      var img = document.createElement("img");
      img.alt = "";
      img.draggable = false;
      thumb.insertBefore(img, thumb.firstChild);
      ssObserveThumb(card, img, thumb, c.participant, c.start);

      // Cross-reference badges
      var xref = findOverlappingData(c.participant, c.start, c.end);
      var trSelf = { icon: XREF_BADGES.transcript.icon, color: XREF_BADGES.transcript.color, title: c.label || c.category || "Transcript" };
      var badgeStack = buildXrefBadges(xref, "transcript", trSelf);
      if (badgeStack) thumb.appendChild(badgeStack);

      container.appendChild(card);
    });
  }

  function buildTrIntakeCategoryPills() {
    var container = qs("#trIntakeCategoryPills");
    if (!container) return;
    var cats = Object.keys(TR_INTAKE_CATEGORIES);
    var counts = {};
    state.trIntakeClusters.forEach(function (c) {
      var k = c.category || "bookmark";
      counts[k] = (counts[k] || 0) + 1;
    });
    container.innerHTML = "";
    cats.forEach(function (key) {
      var cat = TR_INTAKE_CATEGORIES[key];
      var chip = ClipgenPrimitives.createFilterChip({
        label: cat.label,
        active: state.trIntakeFilterCategory === key,
        count: counts[key] || 0,
        hue: categoryHue(key),
        onClick: function () {
          state.trIntakeFilterCategory = state.trIntakeFilterCategory === key ? "" : key;
          renderTranscriptIntake();
        },
      });
      container.appendChild(chip);
    });
  }

  var _trIntakeDensityEl = null;

  function buildTrIntakeDensityTimeline(filtered) {
    var host = qs("#trIntakeTimeline");
    if (!host) return;
    host.innerHTML = "";
    _trIntakeDensityEl = null;
    if (!filtered.length) return;
    var maxEnd = 0;
    for (var i = 0; i < filtered.length; i++) {
      if (filtered[i].end > maxEnd) maxEnd = filtered[i].end;
    }
    var duration = Math.max(maxEnd * 1.05, 60);
    var events = filtered.map(function (c) {
      return {
        t: duration > 0 ? c.start / duration : 0,
        count: c.marks ? c.marks.length : 1,
        hue: categoryHue(c.category || "bookmark"),
      };
    });
    var dt = ClipgenPrimitives.createDensityTimeline({
      events: events,
      durationSec: duration,
      tickCount: 6,
      onBarMouseEnter: function (idx) {
        state.trIntakeHoveredIdx = idx;
        highlightTrIntakeCard(idx);
        if (_trIntakeDensityEl) _trIntakeDensityEl.setHovered(idx);
      },
      onBarMouseLeave: function () {
        state.trIntakeHoveredIdx = -1;
        highlightTrIntakeCard(-1);
        if (_trIntakeDensityEl) _trIntakeDensityEl.setHovered(-1);
      },
      onBarClick: function (idx, ev) {
        var cluster = filteredTranscriptIntakeClusters()[idx];
        if (!cluster) return;
        if (ev && ev.shiftKey) trIntakeAddToReel(cluster);
        else trIntakeAddToArtifacts(cluster);
      },
    });
    _trIntakeDensityEl = dt;
    host.appendChild(dt);
  }

  function trIntakeAddToArtifacts(cluster) {
    state.artifactQueue.push({
      participant: cluster.participant,
      start: cluster.start,
      end: cluster.end,
      desc: cluster.category || "transcript",
      source: "transcript",
      mark_ids: cluster.marks.map(function (m) { return m.id; }),
    });
    renderArtifactQueue();
  }

  function trIntakeAddToReel(cluster) {
    state.reelQueue.push({
      participant: cluster.participant,
      start: cluster.start,
      end: cluster.end,
      desc: cluster.category || "transcript",
      source: "transcript",
      mark_ids: cluster.marks.map(function (m) { return m.id; }),
    });
    renderReelQueue();
  }

  function buildTrIntakeParticipantPills() {
    var container = qs("#trIntakeFilterParticipants");
    if (!container) return;
    var pids = {};
    for (var i = 0; i < state.trIntakeClusters.length; i++) {
      pids[state.trIntakeClusters[i].participant] = true;
    }
    var sorted = Object.keys(pids).sort();
    state.trIntakeFilterParticipants = state.trIntakeFilterParticipants.filter(function (p) {
      return pids[p];
    });
    container.innerHTML = "";
    sorted.forEach(function (p) {
      var pill = ClipgenPrimitives.createParticipantPill({
        id: p,
        active: state.trIntakeFilterParticipants.indexOf(p) !== -1,
        onClick: function () {
          var idx = state.trIntakeFilterParticipants.indexOf(p);
          if (idx === -1) state.trIntakeFilterParticipants.push(p);
          else state.trIntakeFilterParticipants.splice(idx, 1);
          renderTranscriptIntake();
        },
      });
      container.appendChild(pill);
    });
  }

  function highlightTrIntakeCard(idx) {
    var cards = qsa(".tr-intake-queue-card");
    for (var i = 0; i < cards.length; i++) {
      if (i === idx) {
        cards[i].classList.add("intake-highlight");
      } else {
        cards[i].classList.remove("intake-highlight");
      }
    }
  }

  function initTranscriptIntake() {
    var trIntakeCards = qs("#trIntakeCards");
    if (!trIntakeCards) return;
    var trTooltip = qs("#trIntakeTooltip");

    // Click: normal = Artifacts, shift = Reel
    trIntakeCards.addEventListener("click", function (e) {
      var card = e.target.closest(".tr-intake-queue-card");
      if (!card) return;
      var idx = parseInt(card.dataset.trIntakeIdx);
      var cluster = filteredTranscriptIntakeClusters()[idx];
      if (!cluster) return;
      if (e.shiftKey) trIntakeAddToReel(cluster);
      else trIntakeAddToArtifacts(cluster);
    });

    // Right-click: dismiss — remove marks
    trIntakeCards.addEventListener("contextmenu", function (e) {
      var card = e.target.closest(".tr-intake-queue-card");
      if (!card) return;
      e.preventDefault();
      var idx = parseInt(card.dataset.trIntakeIdx);
      var cluster = filteredTranscriptIntakeClusters()[idx];
      if (!cluster) return;
      var ids = cluster.marks.map(function (m) { return m.id; }).filter(Boolean);
      if (!ids.length) return;
      // TODO: DELETE with a JSON body — apiDelete takes no body, so this custom fetch stays.
      fetch("../transcripts/api/marks/" + ids[0], {
        method: "DELETE",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ids: ids }),
      })
        .then(function () { pollTranscriptIntakeMarks(); })
        .catch(function () {});
    });

    // Card hover → highlight + tooltip + timeline marker
    trIntakeCards.addEventListener("mouseover", function (e) {
      var card = e.target.closest(".tr-intake-queue-card");
      if (!card) return;
      var idx = parseInt(card.dataset.trIntakeIdx);
      if (state.trIntakeHoveredIdx !== idx) {
        state.trIntakeHoveredIdx = idx;
        highlightTrIntakeCard(idx);
        if (_trIntakeDensityEl) _trIntakeDensityEl.setHovered(idx);
      }
      if (trTooltip && state.trIntakeTooltipsEnabled) {
        var cluster = filteredTranscriptIntakeClusters()[idx];
        var fullText = cluster ? (cluster.text || cluster.label || "") : "";
        if (fullText) {
          trTooltip.textContent = fullText;
          trTooltip.classList.remove("hidden");
          positionTooltipAnchored(trTooltip, card.getBoundingClientRect());
        } else {
          trTooltip.classList.add("hidden");
        }
      }
    });
    trIntakeCards.addEventListener("mouseleave", function () {
      if (state.trIntakeHoveredIdx !== -1) {
        state.trIntakeHoveredIdx = -1;
        highlightTrIntakeCard(-1);
        if (_trIntakeDensityEl) _trIntakeDensityEl.setHovered(-1);
      }
      if (trTooltip) trTooltip.classList.add("hidden");
    });

    // Cluster threshold change
    var thresholdInput = qs("#trIntakeClusterThreshold");
    if (thresholdInput) {
      thresholdInput.addEventListener("change", function () {
        pollTranscriptIntakeMarks();
      });
    }

    // "Show all" toggle
    var showAllToggle = qs("#trIntakeShowAll");
    if (showAllToggle) {
      showAllToggle.addEventListener("change", function () {
        state.trIntakeShowAll = this.checked;
        pollTranscriptIntakeMarks();
      });
    }

    // Text search filter
    var trSearchEl = qs("#trIntakeFilterSearch");
    var _trIntakeSearchTimer = 0;
    if (trSearchEl) {
      trSearchEl.addEventListener("input", function () {
        state.trIntakeFilterText = this.value;
        clearTimeout(_trIntakeSearchTimer);
        _trIntakeSearchTimer = setTimeout(function () { renderTranscriptIntake(); }, 250);
      });
    }

    // Add All buttons
    var addAllBtn = qs("#trIntakeAddAllBtn");
    if (addAllBtn) {
      addAllBtn.addEventListener("click", function () {
        var filtered = filteredTranscriptIntakeClusters();
        for (var i = 0; i < filtered.length; i++) trIntakeAddToArtifacts(filtered[i]);
      });
    }
    var reelAllBtn = qs("#trIntakeReelAllBtn");
    if (reelAllBtn) {
      reelAllBtn.addEventListener("click", function () {
        var filtered = filteredTranscriptIntakeClusters();
        for (var i = 0; i < filtered.length; i++) trIntakeAddToReel(filtered[i]);
      });
    }

    // Visibility change — pause/resume polling. Runs whenever the page is
    // visible so the transcript-intake counter badge stays live across tabs.
    document.addEventListener("visibilitychange", function () {
      if (document.hidden) {
        if (state.trIntakePollTimer) { clearInterval(state.trIntakePollTimer); state.trIntakePollTimer = null; }
      } else {
        pollTranscriptIntakeMarks();
        if (!state.trIntakePollTimer) state.trIntakePollTimer = setInterval(pollTranscriptIntakeMarks, 10000);
      }
    });
  }

  function initTooltipToggle() {
    state.trIntakeTooltipsEnabled = getStoredTooltipPref();
    var btn = qs("#tooltipToggle");
    if (!btn) return;
    btn.setAttribute("aria-pressed", state.trIntakeTooltipsEnabled ? "true" : "false");
    btn.addEventListener("click", function () {
      state.trIntakeTooltipsEnabled = !state.trIntakeTooltipsEnabled;
      btn.setAttribute("aria-pressed", state.trIntakeTooltipsEnabled ? "true" : "false");
      setStoredTooltipPref(state.trIntakeTooltipsEnabled);
      if (!state.trIntakeTooltipsEnabled) {
        var tt = qs("#trIntakeTooltip");
        if (tt) tt.classList.add("hidden");
      }
    });
  }

  var _exportEnabled = false;

  function _exportItem() {
    return {
      icon: "arrow-down-tray",
      label: "Export",
      action: runExport,
      disabled: !_exportEnabled,
      title: _exportEnabled
        ? "Write JSON+CSV exports of Screenspace and Transcripts manifests"
        : "Run a Screenspace task or transcribe a video first to enable Export.",
    };
  }

  function initTopNavActions() {
    if (!window.ClipgenTopNav) return;
    function clickIfExists(id) {
      var el = document.getElementById(id);
      if (el) el.click();
    }
    function rebuild() {
      window.ClipgenTopNav.setQuickActions([
        { icon: "eye",        label: "Build Viewer",  action: onBuildViewer },
        { icon: "film",       label: "Open Timeline", action: onBuildTimelineViewer },
        { icon: "photo",      label: "Open Gallery",  action: openGalleryDialog },
        _exportItem(),
        { icon: "arrow-path", label: "Refresh sheet", action: function () { clickIfExists("refreshSheet"); } },
      ]);
    }
    function refreshExportStatus() {
      fetch("/api/export/status")
        .then(function (r) { return r.json(); })
        .then(function (j) {
          var next = !!(j && j.any);
          if (next === _exportEnabled) return;
          _exportEnabled = next;
          rebuild();
        })
        .catch(function () {});
    }
    rebuild();
    refreshExportStatus();
    window.ClipgenTopNav.onBeforeOpen(refreshExportStatus);
  }

  function runExport() {
    fetch("/api/export", { method: "POST" })
      .then(function (r) { return r.json().then(function (j) { return { ok: r.ok, j: j }; }); })
      .then(function (res) {
        if (res.ok && res.j && res.j.ok) {
          var n = (res.j.written || []).length;
          showToast("Exported " + n + " file(s) to " + res.j.output_dir);
        } else {
          showToast((res.j && res.j.error) || "Export failed");
        }
      })
      .catch(function (err) { showToast("Export failed: " + err.message); });
  }

  // Apply mask-image to every static [data-icon] element (mirrors what
  // createBtn does for primitives — needed for the bottom-strip toolbar
  // buttons that are written as static HTML).
  function applyDataIconMasks() {
    var nodes = qsa("[data-icon]");
    for (var i = 0; i < nodes.length; i++) {
      var name = nodes[i].dataset.icon;
      if (!name) continue;
      applyMaskIcon(nodes[i], 'url("icons/' + name + '.svg")');
    }
  }

  document.addEventListener("DOMContentLoaded", function () {
    applyDataIconMasks();
    readPersistedSidebarOpen();
    setActiveTabAttr(state.activePreviewTab);
    bindSidebarToggle();
    initThemeToggle();
    initTooltipToggle();
    initPreviewTabs();
    initDropTargets();
    initWheelScroll();
    bindDragGate();
    bindReelReorder();
    bindButtons();
    loadStoredBottomHeight();
    initPanelDivider();
    loadSheetData();
    loadStashes();
    loadArtifactStashes();
    updateViewerButton();
    checkNavLinks();
    initFrontendSwitcher();
    initTopNavActions();
    initIntake();
    initTranscriptIntake();
    pollIntakeStatus();
    pollTrIntakeStatus();
    state.intakeStatusTimer = setInterval(pollIntakeStatus, 5000);
    state.trIntakeStatusTimer = setInterval(pollTrIntakeStatus, 5000);
    // Start intake event polls immediately so the sub-tab counter badges
    // populate on page load and update live regardless of which sub-tab the
    // user is currently viewing. Visibility-change handlers in
    // initIntake / initTranscriptIntake pause + resume them with the page.
    if (!document.hidden) {
      pollIntakeEvents();
      state.intakePollTimer = setInterval(pollIntakeEvents, 10000);
      pollTranscriptIntakeMarks();
      state.trIntakePollTimer = setInterval(pollTranscriptIntakeMarks, 10000);
    }
    window.addEventListener("resize", function () {
      if (window.convergenceResize) window.convergenceResize();
      if (window.metadataResize) window.metadataResize();
    });

    document.addEventListener("dragstart", function (ev) {
      if (!ev.target.closest(".stash-card")) return;
      requestAnimationFrame(revealEmptyStashAreas);
    });
    document.addEventListener("dragend", function () {
      hideEmptyStashAreas();
    });
  });

  window._studioState = state;
  window._studioParseClipTimestamps = parseClipTimestamps;
  window._studioFormatDuration = formatDuration;
  window._studioFindOverlappingData = findOverlappingData;
  window._studioBuildXrefBadges = buildXrefBadges;
  window._studioRenderArtifactQueue = renderArtifactQueue;
  window._studioRenderReelQueue = renderReelQueue;
  window._studioClusterIntakeEvents = clusterIntakeEvents;
  window._studioClusterTranscriptMarks = clusterTranscriptMarks;
  window._studioROW_FUNCTIONS = ROW_FUNCTIONS;
  window._studioSyncPreviewTab = syncPreviewTab;
})();
