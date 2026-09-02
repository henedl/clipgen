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
 *   generatedArtifacts         — completed clip/screen/gif outputs.
 *   generatedReels             — completed reels (hydrated from manifest + appended on build).
 *   generatedViewers           — viewers built this session (not persisted).
 *   cellResults                — per-cell success/error status overlaid
 *                                onto the sheet grid (keyed by cellKey()).
 *   intakeEvents / intakeClusters / intakeSeenIds — Screenspace polling
 *                                snapshot; trIntakeMarks/Clusters mirror
 *                                this for Transcripts.
 */

(function () {
  "use strict";

  var QUEUE_STORAGE_KEY = "clipgen-studio-queues";

  // Clustering lives in intake-cluster.js so Overview can share it.
  var clusterIntakeEvents = window.ClipgenIntakeCluster.clusterIntakeEvents;
  var clusterTranscriptMarks = window.ClipgenIntakeCluster.clusterTranscriptMarks;

  // Satellite namespace: hub publishes state at the tail, satellites publish entry points back.
  var STUDIO = (window.ClipgenStudio = window.ClipgenStudio || {});

  // Mask-image icon span; sizing comes from the parent rule or extraClass (studio.css .cg-icon).
  function iconHTML(name, extraClass) {
    return '<span class="cg-icon cg-icon--' + name + (extraClass ? " " + extraClass : "") + '"></span>';
  }

  var state = {
    sheetData: null,
    desktop: false, // native window: Artifact Log rows get "Show on disk"
    artifactQueue: [],
    reelQueue: [],
    generatedArtifacts: [],
    generatedReels: [],
    generatedViewers: [],
    _logSeq: 0,
    artifactGenerating: false,
    reelGenerating: false,
    overlayJobRunning: false,
    timelineViewerCancelledByUser: false,
    galleryCancelledByUser: false,
    cellResults: {},
    stashes: [],
    artifactStashes: [],
    settingsData: null,
    bottomCollapsed: false,
    activeFunction: "",
    sortColumn: "", // "" | "row" | "category" | "severity" | "function"
    sortDir: "asc", // "asc" | "desc"
    cellExpandHover: true,
    cardScrubberEnabled: false,
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
    intakeFilterText: "",
    intakeFilterDetector: "",
    intakeFilterParticipants: [],
    intakeHoveredIdx: -1,
    activePreviewTab: "sheet",
    trIntakeMarks: [],
    trIntakeClusters: [],
    trIntakeFilterCategory: "",
    trIntakeFilterParticipants: [],
    trIntakeFilterSeverities: [],
    trIntakeFilterText: "",
    trIntakeShowAll: false,
    trIntakeHoveredIdx: -1,
    coIntakeItems: [],
    coTrimCardKeys: {},
    coIntakeFilterParticipants: [],
    coIntakeFilterTypes: [],
    coIntakeFilterText: "",
    coIntakeHoveredIdx: -1,
    _coIntakeFp: null,
    // MindNode intake: one item per timestamp pair; skipped holds timestamp-less notes.
    mnIntakeItems: [],
    mnIntakeSkipped: [],
    mnIntakeFilterParticipants: [],
    mnIntakeFilterCategories: [],
    mnIntakeFilterText: "",
    mnIntakeHoveredIdx: -1,
    _mnIntakeFp: null,
    convergenceBaselines: {},
    convergenceDataVersion: 0,
    convergenceStale: false,
    sidebarOpen: true,
    sidebarCategories: {},
    sidebarKeywords: {},
    sidebarParticipants: {},
  };

  function isIntakeSource(source) {
    return source === "screenspace" || source === "transcript" ||
      source === "composer" || source === "mindnode";
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

  function stampLog(entry) {
    entry._seq = state._logSeq++;
    return entry;
  }

  function pathBasename(p) {
    if (!p) return "";
    var idx = Math.max(p.lastIndexOf("/"), p.lastIndexOf("\\"));
    return idx >= 0 ? p.slice(idx + 1) : p;
  }

  function findInQueue(queue, participant, rowNum) {
    var key = cellKey(participant, rowNum);
    for (var i = 0; i < queue.length; i++) {
      if (cellKey(queue[i].participant, queue[i].row) === key) return i;
    }
    return -1;
  }

  // Cluster spans drift (poll merges, threshold edits); match on shared ids, else exact span.
  function intakeIds(item) {
    if (!item) return [];
    var raw =
      item.source === "screenspace" || item.source === "composer" ||
      item.source === "mindnode"
        ? item.event_ids
        : item.source === "transcript"
          ? item.mark_ids
          : null;
    return raw ? raw.filter(Boolean) : [];
  }

  function intakeItemsOverlap(a, b) {
    if (!a || !b) return false;
    if (a.source !== b.source || a.participant !== b.participant) return false;
    if (!isIntakeSource(a.source)) return false;
    var aIds = intakeIds(a);
    var bIds = intakeIds(b);
    if (aIds.length && bIds.length) {
      for (var i = 0; i < aIds.length; i++) {
        for (var j = 0; j < bIds.length; j++) {
          if (aIds[i] === bIds[j]) return true;
        }
      }
      return false;
    }
    return a.start === b.start && a.end === b.end;
  }

  function findIntakeInQueue(queue, item) {
    if (!item || !isIntakeSource(item.source)) return -1;
    for (var i = 0; i < queue.length; i++) {
      if (intakeItemsOverlap(queue[i], item)) return i;
    }
    return -1;
  }

  // Remove every overlapping entry: a drifted cluster can subsume several (mirrors removeAllCellEntries).
  function removeIntakeFromQueue(queue, item) {
    for (var i = queue.length - 1; i >= 0; i--) {
      if (intakeItemsOverlap(queue[i], item)) queue.splice(i, 1);
    }
  }

  // Idempotent batch add renders once; per-item adds rebuilt the queue every push.
  function intakeAddItems(queue, items, renderFn) {
    var locked = queue === state.artifactQueue ? isArtifactQueueLocked() : isReelQueueLocked();
    if (locked) return;
    var added = false;
    for (var i = 0; i < items.length; i++) {
      if (!items[i] || findIntakeInQueue(queue, items[i]) >= 0) continue;
      queue.push(items[i]);
      added = true;
    }
    if (added) renderFn();
  }

  function intakeAddItem(queue, item, renderFn) {
    intakeAddItems(queue, [item], renderFn);
  }

  function intakeToggleItem(queue, item, renderFn) {
    var locked = queue === state.artifactQueue ? isArtifactQueueLocked() : isReelQueueLocked();
    if (locked) return;
    if (findIntakeInQueue(queue, item) >= 0) removeIntakeFromQueue(queue, item);
    else queue.push(item);
    renderFn();
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
        severity: info.severity || "",
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
    return parseClipSegmentsForCell(raw, baselineSeconds, DEFAULT_DUR);
  }

  // Overlapping data from sibling sources for one participant + time range.
  function findOverlappingData(participant, start, end) {
    var result = { transcriptSnippets: [], screenspaceEvents: [], sheetObservations: [] };

    // Projection: consumers expect `text` already resolved (text || label).
    for (var i = 0; i < state.trIntakeClusters.length; i++) {
      var tc = state.trIntakeClusters[i];
      if (tc.participant === participant && tc.start < end && tc.end > start) {
        result.transcriptSnippets.push({ text: tc.text || tc.label || "", category: tc.category, start: tc.start, end: tc.end });
      }
    }

    // Pass the cluster through; consumers read only detector / event_type.
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

  // Cached: --radius is static on :root, and setCardDragImage is on the dragstart hot path.
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
    // Some Chromium builds ignore class-driven rounding in the drag bitmap; pin it inline.
    clone.style.borderRadius = getCardDragImageRadius();
    clone.style.overflow = "hidden";
    document.body.appendChild(clone);
    ev.dataTransfer.setDragImage(clone, ev.clientX - rect.left, ev.clientY - rect.top);
    requestAnimationFrame(function () {
      if (clone.parentNode) clone.parentNode.removeChild(clone);
    });
  }

  // Must be a rendered element: WebKit aborts the drag on an `Image` or hidden box.
  function createBlankDragImage() {
    var blank = el("div", "drag-image-blank");
    document.body.appendChild(blank);
    return blank;
  }

  // body.dragging lets CSS suspend expensive effects during drags (studio.css, topnav.css, tokens.css).
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

  // ---- Sheet preview sorting ----
  // Cycles asc/desc/off; empties sink, ties keep source order.
  function compareByColumn(a, b, column, participants, asc) {
    if (column === "row") {
      var dr = a.rowNum - b.rowNum;
      return asc ? dr : -dr;
    }
    if (column === "category") {
      var ac = (a.category || "").trim();
      var bc = (b.category || "").trim();
      if (!ac && !bc) return 0;
      if (!ac) return 1;
      if (!bc) return -1;
      var dc = ac.localeCompare(bc, undefined, { sensitivity: "base" });
      return asc ? dc : -dc;
    }
    if (column === "severity") {
      var ar = severityRank(a.severity);
      var br = severityRank(b.severity);
      if (ar === null && br === null) return 0;
      if (ar === null) return 1;
      if (br === null) return -1;
      var ds = ar - br; // rank -4 (Critical) .. 2 (Very Positive): asc = most severe first
      return asc ? ds : -ds;
    }
    if (column === "function") {
      var fn = state.activeFunction && ROW_FUNCTIONS[state.activeFunction];
      if (!fn) return 0;
      var df = fn(a, participants) - fn(b, participants);
      return asc ? df : -df;
    }
    return 0;
  }

  function sortRows(rows) {
    if (!state.sortColumn) return rows;
    var participants = state.sheetData ? state.sheetData.participants : [];
    var asc = state.sortDir !== "desc";
    var col = state.sortColumn;
    return rows.slice().sort(function (a, b) {
      var c = compareByColumn(a, b, col, participants, asc);
      return c !== 0 ? c : a.rowNum - b.rowNum;
    });
  }

  function cycleSort(column) {
    if (state.sortColumn !== column) {
      state.sortColumn = column;
      state.sortDir = "asc";
    } else if (state.sortDir === "asc") {
      state.sortDir = "desc";
    } else {
      state.sortColumn = "";
      state.sortDir = "asc";
    }
    renderGrid();
  }

  // Rebuilt on every renderGrid, so the listener attaches fresh (like fnSelect/fnClear).
  function buildSortButton(column) {
    var active = state.sortColumn === column;
    var iconName = !active
      ? "chevron-up-down"
      : state.sortDir === "asc"
        ? "bars-arrow-up"
        : "bars-arrow-down";
    var btn = el("button", "sort-btn" + (active ? " sort-btn-active" : ""));
    btn.type = "button";
    btn.title = active
      ? state.sortDir === "asc"
        ? "Sorted ascending. Click for descending"
        : "Sorted descending. Click to clear"
      : "Sort by this column";
    btn.innerHTML = '<span class="cg-icon cg-icon--' + iconName + '"></span>';
    btn.addEventListener("click", function (ev) {
      ev.stopPropagation(); // don't trigger the header's batch-select handler
      cycleSort(column);
    });
    return btn;
  }

  // <th> with a centered label + sort button. Used for #, Category, Severity.
  function sortableHeaderTh(thClass, label, column) {
    var th = el("th", thClass);
    var inner = el("div", "col-header-inner");
    inner.appendChild(el("span", "", label));
    inner.appendChild(buildSortButton(column));
    th.appendChild(inner);
    return th;
  }

  function clearAllFilters() {
    // Reset the checkbox source maps too, or sidebar rows stay active after clearing.
    state.sidebarCategories = {};
    state.sidebarKeywords = {};
    state.filters.categories = [];
    state.filters.severities = [];
    state.filters.keywords = [];
    state.filters.fnMin = null;
    state.filters.fnMax = null;
  }

  // ---- Sheet sidebar ----

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

  // ---- Preview tabs ----

  function initPreviewTabs() {
    var tabs = qsa(".preview-tab");
    for (var i = 0; i < tabs.length; i++) {
      tabs[i].addEventListener("click", function () {
        var target = this.dataset.tab;
        if (target === state.activePreviewTab) return;
        // Queues/stashes span tabs; a parked cursor would leave a stale outline.
        if (_kbRegion) kbClearCursor();
        state.activePreviewTab = target;
        setStoredUIStateField("studio", "activeTab", target);
        var allTabs = qsa(".preview-tab");
        for (var j = 0; j < allTabs.length; j++) allTabs[j].classList.remove("active");
        this.classList.add("active");
        syncPreviewTab(true);
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
    syncPreviewTab(false);
  }

  function syncPreviewTab(animate) {
    setActiveTabAttr(state.activePreviewTab);
    var grid = qs("#sheetGrid");
    var refreshBtn = qs("#studioRefresh");
    var intakePanel = qs("#intakePanel");
    var trIntakePanel = qs("#trIntakePanel");
    var coIntakePanel = qs("#coIntakePanel");
    var mnIntakePanel = qs("#mnIntakePanel");

    // Hide everything first
    grid.classList.add("hidden");
    intakePanel.classList.add("hidden");
    if (trIntakePanel) trIntakePanel.classList.add("hidden");
    if (coIntakePanel) coIntakePanel.classList.add("hidden");
    if (mnIntakePanel) mnIntakePanel.classList.add("hidden");

    // The Refresh button stays on every tab; only what it refreshes changes.
    var refreshTitle = "Refresh sheet data from source";

    var activePanel = null;
    if (state.activePreviewTab === "sheet") {
      grid.classList.remove("hidden");
      activePanel = grid;
    } else if (state.activePreviewTab === "intake") {
      intakePanel.classList.remove("hidden");
      activePanel = intakePanel;
      refreshTitle = "Refresh Screenspace intake";
    } else if (state.activePreviewTab === "transcript-intake") {
      if (trIntakePanel) trIntakePanel.classList.remove("hidden");
      activePanel = trIntakePanel;
      refreshTitle = "Refresh transcript intake";
    } else if (state.activePreviewTab === "composer-intake") {
      if (coIntakePanel) coIntakePanel.classList.remove("hidden");
      activePanel = coIntakePanel;
      refreshTitle = "Refresh Composer intake";
    } else if (state.activePreviewTab === "mindnode-intake") {
      if (mnIntakePanel) mnIntakePanel.classList.remove("hidden");
      activePanel = mnIntakePanel;
      refreshTitle = "Re-read the mind map";
    }
    if (refreshBtn) refreshBtn.title = refreshTitle;
    if (activePanel && animate) {
      activePanel.classList.add("tab-slide-enter");
      requestAnimationFrame(function () {
        requestAnimationFrame(function () {
          activePanel.classList.remove("tab-slide-enter");
        });
      });
    }
    // Cursor is per-surface; repaint drops paint on the inactive tab.
    kbPaintCursor();
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

  function populateSheetSkeleton() {
    buildSkeletonGrid(qs("#sheetLoading .skeleton-grid"), 9, 4);
    var loading = qs("#sheetLoading");
    if (loading && !loading.querySelector(".sheet-loading-caption")) {
      var cap = el("div", "sheet-loading-caption cg-shimmer", "Loading sheet…");
      cap.setAttribute("aria-live", "polite");
      loading.insertBefore(cap, loading.firstChild);
    }
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
        if (data.sheet_loaded === false) {
          var loading = qs("#sheetLoading");
          if (loading) {
            var lcap = loading.querySelector(".sheet-loading-caption");
            if (lcap) loading.removeChild(lcap);
            loading.classList.add("is-empty");
            var caption = document.createElement("div");
            caption.className = "sheet-empty-caption";
            var prefix = document.createTextNode("No spreadsheet loaded. Click ");
            var icon = document.createElement("span");
            icon.className = "sheet-empty-caption-icon";
            icon.setAttribute("aria-hidden", "true");
            var suffix = document.createTextNode(" in the top bar to pick one.");
            caption.appendChild(prefix);
            caption.appendChild(icon);
            caption.appendChild(suffix);
            loading.appendChild(caption);
          }
          clipgenApplyConfig(data.config);
          updateArtifactFormatLabels();
          state.sheetData = data;
          // Mind-map-only sessions have a study; guard keeps empty sessions' subheader blank.
          if (data.study) renderHeader();
          return;
        }
        state.sheetData = data;
        clipgenApplyConfig(data.config);
        updateArtifactFormatLabels();
        // Restore filters before first render (activeFunction first keeps the range enabled).
        restoreSidebarFilters();
        renderHeader();
        renderSidebar();
        renderGrid();
        // Baselines put durations in the video-relative frame (matches prepare_clip); re-render on arrival.
        apiGet("api/sheet/baseline")
          .then(function (bdata) {
            state.convergenceBaselines = (bdata.ok && bdata.baselines) ? bdata.baselines : {};
            if (Object.keys(state.convergenceBaselines).length > 0) renderGrid();
          })
          .catch(function () {
            state.convergenceBaselines = {};
          });
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
        if (data.cardScrubberEnabled !== undefined) {
          state.cardScrubberEnabled = data.cardScrubberEnabled;
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
      })
      .catch(function (err) {
        qs("#sheetLoading").textContent = "Failed to load sheet: " + err;
      });
  }

  function refreshSheetData() {
    return apiPost("api/sheet/refresh")
      .then(function (data) {
        if (!data.ok) {
          showResult(null, "Refresh failed: " + (data.error || "Unknown error"));
          return;
        }
        loadSheetData();
      })
      .catch(function (err) {
        showResult(null, "Refresh failed: " + err);
      });
  }

  // Refresh acts on the visible tab: Sheet re-reads the spreadsheet, intakes wake their poller.
  function refreshActiveTab() {
    var btn = qs("#studioRefresh");
    if (!btn || btn.disabled) return;
    btn.disabled = true;
    // Spin the masked icon span — .cg-btn has no inline <svg> to animate.
    var icon = btn.querySelector(".cg-btn-icon");
    if (icon) icon.style.animation = "spin 0.7s linear infinite";

    var work;
    if (state.activePreviewTab === "intake") {
      work = refreshScreenspaceIntake();
    } else if (state.activePreviewTab === "transcript-intake") {
      work = refreshTranscriptIntake();
    } else if (state.activePreviewTab === "composer-intake") {
      work = refreshComposerIntake();
    } else if (state.activePreviewTab === "mindnode-intake") {
      work = refreshMindnodeIntake();
    } else {
      work = refreshSheetData();
    }

    Promise.resolve(work).catch(function () {}).then(function () {
      btn.disabled = false;
      if (icon) icon.style.animation = "";
    });
  }

  // Reconcile manifest against the sheet: mark cells green, re-enqueue valid ones; `seen` dedupes cells.
  function loadManifestState() {
    apiGet("api/manifest")
      .then(function (data) {
        if (!data.ok || !state.sheetData) return;
        var artifacts = data.artifacts || [];
        var reels = data.reels || [];
        if (artifacts.length === 0 && reels.length === 0) return;

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

        // Dedupe by id (fallback file) so a job-status reattach doesn't double-list.
        var seenArtifact = {};
        for (var ai = 0; ai < state.generatedArtifacts.length; ai++) {
          var prev = state.generatedArtifacts[ai];
          var k = prev.id || prev.file;
          if (k) seenArtifact[k] = true;
        }
        var keep = artifacts.filter(function (a) { return a.type !== "transcript" && a.file; });
        for (var ki = 0; ki < keep.length; ki++) {
          var entry = keep[ki];
          var ek = entry.id || entry.file;
          if (ek && seenArtifact[ek]) continue;
          if (ek) seenArtifact[ek] = true;
          state.generatedArtifacts.push(stampLog(entry));
        }

        var seenReel = {};
        for (var rj = 0; rj < state.generatedReels.length; rj++) {
          var pr = state.generatedReels[rj];
          var rk = pr.id || pr.file;
          if (rk) seenReel[rk] = true;
        }
        for (var ri = 0; ri < reels.length; ri++) {
          var reel = reels[ri];
          var rk2 = reel.id || reel.file;
          if (rk2 && seenReel[rk2]) continue;
          if (rk2) seenReel[rk2] = true;
          state.generatedReels.push(stampLog(reel));
        }

        renderArtifactQueue();
        updateCellClasses();
        renderLog();
      })
      .catch(function () {});
  }

  // Re-attach to a background build after navigating away; otherwise Cancel is unreachable.
  function applyJobStatus(status) {
    if (!status) return;
    var reel = status.reel || {};
    var gen = status.generate || {};

    // ---- Reel side ----
    if (reel.in_progress) {
      if (!state.reelGenerating) setReelGenerating(true);
      qs("#cancelReelBtn").classList.remove("hidden");
      var totalClips = reel.total_clips || 0;
      var clipsDone = reel.clips_done || 0;
      var concatFraction = typeof reel.concat_progress === "number" ? reel.concat_progress : 0;
      var clipFraction = totalClips > 0 ? Math.min(clipsDone / totalClips, 1) : 0;
      // Same 0.7/0.3 weighting as the live stream handler in onBuildReel().
      setButtonProgress("buildReelBtn", clipFraction * 0.7 + concatFraction * 0.3);
      // Seed from the server start time; idempotent start() leaves a live clock alone.
      _reelEtaTracker.start(reel.started_at ? reel.started_at * 1000 : undefined);
      _studioEtaTicker.ensure();
      _paintReelElapsed();
    } else if (state._jobStatusReelWasInProgress) {
      // Busy → idle while polling: a background build finished. Reload the manifest.
      setReelGenerating(false);
      qs("#cancelReelBtn").classList.add("hidden");
      setButtonProgress("buildReelBtn", null);
      _reelEtaTracker.reset();
      _paintReelElapsed();
      loadManifestState();
    }
    state._jobStatusReelWasInProgress = !!reel.in_progress;

    // ---- Generate side: sheet and intake streams share one button, so combine their counts.
    var intake = status.intake || {};
    var genActive = !!gen.in_progress || !!intake.in_progress;
    if (genActive) {
      if (!state.artifactGenerating) setArtifactGenerating(true);
      qs("#cancelGenerateBtn").classList.remove("hidden");
      var combinedTotal = (gen.total || 0) + (intake.total || 0);
      var combinedDone = (gen.done || 0) + (intake.done || 0);
      if (combinedTotal > 0) {
        setButtonProgress("generateBtn", Math.min(combinedDone / combinedTotal, 1));
      }
      // Seed elapsed from the earlier start; idempotent start() leaves a live clock alone.
      var genStartedAt =
        gen.started_at && intake.started_at
          ? Math.min(gen.started_at, intake.started_at)
          : gen.started_at || intake.started_at;
      _generateEtaTracker.start(genStartedAt ? genStartedAt * 1000 : undefined);
      // Both sides count artifacts, so the totals sum and intake-only runs keep a readout.
      updateGenerateProgress(combinedDone, combinedTotal);
      _studioEtaTicker.ensure();
    } else if (state._jobStatusGenerateWasInProgress) {
      setArtifactGenerating(false);
      qs("#cancelGenerateBtn").classList.add("hidden");
      setButtonProgress("generateBtn", null);
      _generateEtaTracker.reset();
      updateGenerateProgress(0, 0);
      loadManifestState();
    }
    state._jobStatusGenerateWasInProgress = genActive;
  }

  function pollJobStatus() {
    return apiGet("api/job-status")
      .then(function (data) {
        if (!data || !data.ok) return;
        applyJobStatus(data);
        // Poll only while a job is in flight.
        var stillBusy =
          (data.reel && data.reel.in_progress) ||
          (data.generate && data.generate.in_progress) ||
          (data.intake && data.intake.in_progress);
        if (stillBusy) {
          startJobStatusPoll();
        } else {
          stopJobStatusPoll();
        }
      })
      .catch(function () { /* transient errors — keep timer running */ });
  }

  function startJobStatusPoll() {
    if (state.jobStatusPoller) return;
    // runImmediately is false: the first poll fires after one interval (1s).
    state.jobStatusPoller = createPoller(pollJobStatus, 1000, {
      runImmediately: false,
      label: "studio.jobStatus",
    });
    state.jobStatusPoller.start();
  }

  function stopJobStatusPoll() {
    if (state.jobStatusPoller) {
      state.jobStatusPoller.stop();
      state.jobStatusPoller = null;
    }
  }

  // ---- Header rendering ----

  function renderHeader() {
    var d = state.sheetData;
    // Without a sheet the cohort is the mind map's; `participants` means sheet columns.
    var people = (d.participants && d.participants.length)
      ? d.participants
      : (d.mindnodeParticipants || []);
    qs("#studyName").textContent = d.study || "Unknown study";
    qs("#versionInfo").textContent = people.length + " participants";
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
    return clipgenPerf.span("studio.renderGrid", renderGridImpl);
  }

  function renderGridImpl() {
    var d = state.sheetData;
    var grid = qs("#sheetGrid");
    var prevScrollTop = grid.scrollTop;
    var prevScrollLeft = grid.scrollLeft;
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
    colRowNum.style.width = "3.5rem"; // wide enough for "#" label + sort button
    colgroup.appendChild(colRowNum);
    var colFn = document.createElement("col");
    colFn.style.width = "3.25rem"; // fits the fn select + clear + optional sort button
    colgroup.appendChild(colFn);
    var colObs = document.createElement("col");
    // Explicit width: under table-layout fixed, `auto` collapses to 0 when others overflow.
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

    var batchTh = sortableHeaderTh("col-row-num col-row-num-header", "#", "row");
    batchTh.title = "Select all cells";
    hrow.appendChild(batchTh);

    var fnTh = el("th", "col-function");
    var fnWrap = el("div", "fn-header-wrap");
    var fnSelect = document.createElement("select");
    // No caret: .col-function is 3.5rem and also holds .fn-clear.
    fnSelect.className = "fn-select cg-select-nocaret";
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
      if (!this.value && state.sortColumn === "function") state.sortColumn = "";
      // syncFilterFnDisabled may clear fnMin/fnMax — persist after so the
      // stored range matches what's actually applied.
      syncFilterFnDisabled();
      persistSidebarFilters();
      // Full re-render so the function-column sort button appears/disappears.
      renderGrid();
    });
    fnClear.addEventListener("click", function () {
      state.activeFunction = "";
      fnSelect.value = "";
      this.style.display = "none";
      if (state.sortColumn === "function") state.sortColumn = "";
      syncFilterFnDisabled();
      persistSidebarFilters();
      renderGrid();
    });
    fnWrap.appendChild(fnSelect);
    fnWrap.appendChild(fnClear);
    // Sort button only makes sense once a function populates the column.
    if (state.activeFunction) fnWrap.appendChild(buildSortButton("function"));
    fnTh.appendChild(fnWrap);
    hrow.appendChild(fnTh);

    hrow.appendChild(el("th", "col-observation", "Observation"));
    hrow.appendChild(sortableHeaderTh("col-category", "Category", "category"));
    if (showSeverity) {
      hrow.appendChild(sortableHeaderTh("col-severity", "Severity", "severity"));
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
    grid.scrollTop = prevScrollTop;
    grid.scrollLeft = prevScrollLeft;

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

    var filteredRows = sortRows(getFilteredRows(d.rows));
    var frag = document.createDocumentFragment();
    var i = 0;
    while (i < filteredRows.length) {
      var row = filteredRows[i];
      if (isRowEmpty(row, visibleParticipants)) {
        // Sorted views drop empty rows; the spreadsheet grouping is meaningless there.
        if (state.sortColumn) {
          i++;
          continue;
        }
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
    kbPaintCursor();
  }

  // ---- Keyboard cursor ----
  // One logical cursor per surface; re-renders repaint it via kbPaintCursor().

  var _kbCursor = null; // {surface: "sheet", participant, row} | {surface: <intake tab | list surface>, idx}
  var _kbRegion = null; // explicit list-surface override set by the jump hotkeys (kbJumpTo)

  // Index-addressed list surfaces the cursor can jump into; sheet and intake cards stay special-cased.
  function kbActivateClick(el) { if (el) { el.click(); return true; } return false; }
  function kbActivateRemove(el) {
    var btn = el && el.querySelector(".queue-card-remove");
    if (btn) { btn.click(); return true; }
    return false;
  }
  function ensureSidebarOpen() { if (!state.sidebarOpen) toggleSidebar(); }
  function ensureBottomOpen() { if (state.bottomCollapsed) toggleBottomPanel(); }

  var KB_LIST_SURFACES = {
    filter: { itemsSel: "#studioSidebar .studio-sidebar-row", ensure: ensureSidebarOpen, activate: kbActivateClick, verb: "Toggle filter", sheetOnly: true },
    "artifact-queue": { itemsSel: "#artifactsList .queue-card", ghostSel: "#artifactsList .queue-card-ghost", ensure: ensureBottomOpen, activate: kbActivateRemove, verb: "Remove" },
    "reel-queue": { itemsSel: "#reelList .queue-card", ghostSel: "#reelList .queue-card-ghost", ensure: ensureBottomOpen, activate: kbActivateRemove, verb: "Remove" },
    "artifact-stash": { itemsSel: "#stashedArtifactsList .stash-card", ensure: ensureBottomOpen, activate: kbActivateClick, verb: "Recall to queue" },
    "reel-stash": { itemsSel: "#stashedReelsList .stash-card", ensure: ensureBottomOpen, activate: kbActivateClick, verb: "Recall to queue" },
  };

  function kbListItems(surface) {
    var cfg = KB_LIST_SURFACES[surface];
    return cfg ? qsa(cfg.itemsSel) : [];
  }

  function kbSurface() {
    return _kbRegion || state.activePreviewTab || "sheet";
  }

  // Double-pulse an empty queue's ghost; remove-then-reflow lets a repeat replay it.
  function pulseGhost(sel) {
    var ghost = qs(sel);
    if (!ghost) return;
    ghost.classList.remove("queue-card-ghost-pulse");
    void ghost.offsetWidth;
    ghost.classList.add("queue-card-ghost-pulse");
  }

  // Jump to a list surface's first item; declines the key when off-tab or empty.
  function kbJumpTo(region) {
    var cfg = KB_LIST_SURFACES[region];
    if (!cfg) return false;
    if (cfg.sheetOnly && state.activePreviewTab !== "sheet") return false;
    if (cfg.ensure) cfg.ensure();
    if (!kbListItems(region).length) {
      if (cfg.ghostSel) pulseGhost(cfg.ghostSel);
      return false;
    }
    // Drop stray native focus so only one focus indicator shows.
    if (window.ClipgenHotkeys && window.ClipgenHotkeys.blurStrayFocus) {
      window.ClipgenHotkeys.blurStrayFocus();
    }
    _kbRegion = region;
    _kbCursor = { surface: region, idx: 0 };
    kbPaintCursor();
    return true;
  }

  function kbSheetCells() {
    return qsa("#sheetGrid .ts-cell.valid-ts");
  }

  function kbCursorEl() {
    if (!_kbCursor || _kbCursor.surface !== kbSurface()) return null;
    if (_kbCursor.surface === "sheet") {
      return qs('#sheetGrid .ts-cell.valid-ts[data-participant="' + CSS.escape(_kbCursor.participant) +
                '"][data-row="' + CSS.escape(String(_kbCursor.row)) + '"]');
    }
    if (KB_LIST_SURFACES[_kbCursor.surface]) {
      return kbListItems(_kbCursor.surface)[_kbCursor.idx] || null;
    }
    return STUDIO.intakeCardAt ? STUDIO.intakeCardAt(_kbCursor.surface, _kbCursor.idx) : null;
  }

  function kbPaintCursor() {
    var old = qsa(".kb-cursor");
    for (var i = 0; i < old.length; i++) old[i].classList.remove("kb-cursor");
    var cur = kbCursorEl();
    if (cur) {
      cur.classList.add("kb-cursor");
      if (cur.scrollIntoView) cur.scrollIntoView({ block: "nearest", inline: "nearest" });
    }
  }

  // Alt-hold hints: sheet cells flip Send/Remove by queue membership; intake cards stay generic.
  function kbActionHintEntries() {
    if (_kbCursor && KB_LIST_SURFACES[_kbCursor.surface]) {
      // Single Enter chip labeled with the list surface's primary action.
      return [{ id: "studio.sendArtifacts", label: KB_LIST_SURFACES[_kbCursor.surface].verb }];
    }
    var artLabel = "Send to Artifacts";
    var reelLabel = "Send to Reel";
    if (_kbCursor && _kbCursor.surface === "sheet") {
      if (findInQueue(state.artifactQueue, _kbCursor.participant, _kbCursor.row) >= 0) {
        artLabel = "Remove from Artifacts";
      }
      if (findInQueue(state.reelQueue, _kbCursor.participant, _kbCursor.row) >= 0) {
        reelLabel = "Remove from Reel";
      }
    }
    return [
      { id: "studio.sendArtifacts", label: artLabel },
      { id: "studio.sendReel", label: reelLabel },
    ];
  }

  function kbClearCursor() {
    if (!_kbCursor && !_kbRegion) return false;
    _kbCursor = null;
    _kbRegion = null;
    kbPaintCursor();
    return true;
  }

  // Row-major step across cells, index step across cards; false when nothing to select.
  function kbStep(delta) {
    var surface = kbSurface();
    if (surface === "sheet") {
      var cells = kbSheetCells();
      if (!cells.length) return false;
      var idx = -1;
      var cur = kbCursorEl();
      for (var i = 0; i < cells.length; i++) {
        if (cells[i] === cur) { idx = i; break; }
      }
      var next = idx < 0 ? (delta > 0 ? 0 : cells.length - 1) : idx + delta;
      if (next < 0) next = 0;
      if (next > cells.length - 1) next = cells.length - 1;
      var td = cells[next];
      _kbCursor = {
        surface: "sheet",
        participant: td.getAttribute("data-participant"),
        row: parseInt(td.getAttribute("data-row"), 10),
      };
    } else if (KB_LIST_SURFACES[surface]) {
      var items = kbListItems(surface);
      if (!items.length) return false;
      var listIdx = _kbCursor && _kbCursor.surface === surface ? _kbCursor.idx + delta : (delta > 0 ? 0 : items.length - 1);
      if (listIdx < 0) listIdx = 0;
      if (listIdx > items.length - 1) listIdx = items.length - 1;
      _kbCursor = { surface: surface, idx: listIdx };
    } else {
      var count = STUDIO.intakeCardCount ? STUDIO.intakeCardCount(surface) : 0;
      if (!count) return false;
      var cardIdx = _kbCursor && _kbCursor.surface === surface ? _kbCursor.idx + delta : (delta > 0 ? 0 : count - 1);
      if (cardIdx < 0) cardIdx = 0;
      if (cardIdx > count - 1) cardIdx = count - 1;
      _kbCursor = { surface: surface, idx: cardIdx };
    }
    kbPaintCursor();
    return true;
  }

  // Vertical: nearest valid cell in an adjacent row, same participant preferred; cards step linearly.
  function kbStepVertical(dir) {
    var surface = kbSurface();
    if (surface !== "sheet") return kbStep(dir);
    var cells = kbSheetCells();
    if (!cells.length) return false;
    var cur = kbCursorEl();
    if (!cur) return kbStep(dir);
    var curRow = parseInt(cur.getAttribute("data-row"), 10);
    var curP = cur.getAttribute("data-participant");
    var idx = -1;
    for (var i = 0; i < cells.length; i++) {
      if (cells[i] === cur) { idx = i; break; }
    }
    var fallback = null;
    var target = null;
    for (var j = idx + dir; j >= 0 && j < cells.length; j += dir) {
      var rowNum = parseInt(cells[j].getAttribute("data-row"), 10);
      if (dir > 0 ? rowNum <= curRow : rowNum >= curRow) continue;
      if (!fallback) fallback = cells[j];
      if (cells[j].getAttribute("data-participant") === curP) { target = cells[j]; break; }
    }
    var td = target || fallback;
    if (!td) return false;
    _kbCursor = {
      surface: "sheet",
      participant: td.getAttribute("data-participant"),
      row: parseInt(td.getAttribute("data-row"), 10),
    };
    kbPaintCursor();
    return true;
  }

  // Enter must not hijack a focused control (button/link activation).
  function kbEnterSafe() {
    var ae = document.activeElement;
    if (!ae || ae === document.body) return true;
    return !(ae.tagName === "BUTTON" || ae.tagName === "A" || ae.tabIndex >= 0);
  }

  function kbSend(reel) {
    if (!kbEnterSafe()) return false;
    var surface = kbSurface();
    if (!_kbCursor || _kbCursor.surface !== surface) return false;
    if (surface === "sheet") {
      var td = kbCursorEl();
      if (!isSelectableTimestampCell(td)) return false;
      var info = getCellInfo(td);
      if (reel) toggleReelCell(info);
      else toggleArtifactCell(info);
      return true;
    }
    if (KB_LIST_SURFACES[surface]) {
      // Lists have one primary action; repaint because it may rebuild or shrink the list.
      var acted = KB_LIST_SURFACES[surface].activate(kbCursorEl(), reel);
      if (acted) kbPaintCursor();
      return acted;
    }
    return !!(STUDIO.intakeToggleAt && STUDIO.intakeToggleAt(surface, _kbCursor.idx, reel));
  }

  // Backspace/Delete remove the focused queue card; false elsewhere keeps browser-back.
  function kbRemoveCard() {
    var surface = kbSurface();
    if (surface !== "artifact-queue" && surface !== "reel-queue") return false;
    if (!_kbCursor || _kbCursor.surface !== surface) return false;
    var el = kbCursorEl();
    if (el) {
      KB_LIST_SURFACES[surface].activate(el, false);
      kbPaintCursor();
    }
    return true;
  }

  // ---- Panel divider ----
  // #bottomPanel gets a pixel height from state.bottomH; #sheetPreview flexes.
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

  function initBottomPanelDivider() {
    initPanelDivider({
      isCollapsed: function () {
        return state.bottomCollapsed;
      },
      getHeight: function () {
        return state.bottomH;
      },
      setHeight: function (h) {
        state.bottomH = h;
        applyBottomHeight();
      },
      // Suppress the height transition during drag, or the panel lags the cursor (see Screenspace).
      onDragStart: function () {
        document.body.classList.add("panel-dragging");
      },
      onDragEnd: function () {
        document.body.classList.remove("panel-dragging");
      },
      getBounds: function () {
        // Reserve at least MIN_UPPER for the sheet pane.
        var header = qs("#studioSubheader");
        var divider = qs("#panelDivider");
        if (header && divider) {
          var headerRect = header.getBoundingClientRect();
          var available =
            window.innerHeight - headerRect.top - headerRect.height - divider.offsetHeight;
          var MIN_UPPER = 120;
          return {
            min: BOTTOM_STRIP_MIN,
            max: Math.max(BOTTOM_STRIP_MIN, Math.min(BOTTOM_STRIP_MAX, available - MIN_UPPER)),
          };
        }
        return { min: BOTTOM_STRIP_MIN, max: BOTTOM_STRIP_MAX };
      },
      onToggle: toggleBottomPanel,
      persist: persistBottomHeight,
    });
  }

  function toggleBottomPanel() {
    var bottom = qs("#bottomPanel");
    if (!bottom || bottom._transitioning) return;
    bottom._transitioning = true;

    if (state.bottomCollapsed) {
      // --- Restore ---
      // Animate between pixel endpoints; never `auto` mid-flight.
      state.bottomCollapsed = false;
      document.body.classList.add("bottom-animating");
      document.body.classList.remove("bottom-collapsed");
      bottom.style.height = "0px";
      bottom.offsetHeight; // reflow — pin the start frame
      bottom.style.height = state.bottomH + "px";

      onCollapseTransitionEnd(bottom, function () {
        document.body.classList.remove("bottom-animating");
        bottom._transitioning = false;
        persistBottomHeight();
      });
    } else {
      // --- Collapse ---
      // Pin the pixel height, then animate to 0; `auto` would hitch.
      state.bottomCollapsed = true;
      var currentH = bottom.offsetHeight;
      document.body.classList.add("bottom-animating");
      bottom.style.height = currentH + "px";
      bottom.offsetHeight; // reflow — pin the start frame
      document.body.classList.add("bottom-collapsed");
      bottom.style.height = "0px";

      onCollapseTransitionEnd(bottom, function () {
        bottom._transitioning = false;
        document.body.classList.remove("bottom-animating");
        // Drop the inline height so the collapsed CSS (height: 0) governs the
        // rest state.
        bottom.style.height = "";
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
      if (e.target === el && e.propertyName === "height") done();
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
        // Native tooltip for clipped chips; cheaper than scrollWidth reads in the render loop.
        chip.title = cellData.value;
        td.appendChild(chip);
        if (cellData.valid) {
          td.classList.add("valid-ts");
          td.setAttribute("draggable", "true");
        } else {
          td.classList.add("has-text");
        }
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
    var selected = {};
    for (var i = 0; i < state.artifactQueue.length; i++) {
      var a = state.artifactQueue[i];
      if (a.row) selected[cellKey(a.participant, a.row)] = true;
    }
    for (var j = 0; j < state.reelQueue.length; j++) {
      var r = state.reelQueue[j];
      if (r.row) selected[cellKey(r.participant, r.row)] = true;
    }
    var cells = qsa(".ts-cell");
    for (var k = 0; k < cells.length; k++) {
      var td = cells[k];
      var key = cellKey(
        td.getAttribute("data-participant"),
        parseInt(td.getAttribute("data-row"), 10),
      );
      td.classList.toggle("selected", !!selected[key]);
    }
  }

  function updateSingleCellClass(participant, row) {
    var td = qs('.ts-cell[data-participant="' + CSS.escape(participant) +
                '"][data-row="' + CSS.escape(row) + '"]');
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
      severity: td.getAttribute("data-severity") || "",
    };
  }

  function isSelectableTimestampCell(td) {
    return !!(td && td.classList.contains("valid-ts"));
  }

  function isAnyStudioJobRunning() {
    return state.artifactGenerating || state.reelGenerating || state.overlayJobRunning;
  }

  function isArtifactQueueLocked() {
    return state.artifactGenerating;
  }

  function isReelQueueLocked() {
    return state.reelGenerating;
  }

  // "3 clip" -> "3 clips" ("GIF" -> "GIFs"). Presentation copy for tooltips.
  function plural(n, word) {
    return n + " " + word + (n === 1 ? "" : "s");
  }

  // Singular noun for the currently selected artifact output format.
  function artifactNoun() {
    var sel = qs("#artifactFormat");
    var v = sel ? sel.value : "";
    if (v === "screen") return "screenshot";
    if (v === "gif") return "GIF";
    if (v === "clip") return "clip";
    return "artifact";
  }

  function updateArtifactActions() {
    var n = state.artifactQueue.length;
    var artLocked = isArtifactQueueLocked();
    var reelLocked = isReelQueueLocked();
    var genBtn = qs("#generateBtn");
    if (genBtn) {
      genBtn.disabled = artLocked || n === 0;
      genBtn.setAttribute(
        "data-tooltip",
        artLocked
          ? "Generating…"
          : n === 0
            ? "Add cells to the work area first"
            : "Generate " + plural(n, artifactNoun()),
      );
    }
    var clearBtn = qs("#clearArtifactsBtn");
    if (clearBtn) clearBtn.disabled = artLocked || n === 0;
    var stashBtn = qs("#stashArtifactsBtn");
    if (stashBtn) {
      stashBtn.disabled = artLocked || n === 0;
      stashBtn.setAttribute(
        "data-tooltip",
        artLocked
          ? "Finish generating first"
          : n === 0
            ? "Add cells to the work area first"
            : "Stash " + plural(n, "artifact") + " to reuse later",
      );
    }
    var addToReelBtn = qs("#addToReelBtn");
    if (addToReelBtn) {
      addToReelBtn.disabled = reelLocked || n === 0;
      addToReelBtn.setAttribute(
        "data-tooltip",
        reelLocked
          ? "Finish generating first"
          : n === 0
            ? "Add cells to the work area first"
            : "Add " + plural(n, "clip") + " to the reel",
      );
    }
  }

  function updateReelActions() {
    var n = state.reelQueue.length;
    var reelLocked = isReelQueueLocked();
    var buildBtn = qs("#buildReelBtn");
    if (buildBtn) {
      buildBtn.disabled = reelLocked || n === 0;
      buildBtn.setAttribute(
        "data-tooltip",
        reelLocked
          ? "Building…"
          : n === 0
            ? "Add clips to the reel first"
            : "Build a reel from " + plural(n, "clip"),
      );
    }
    var clearBtn = qs("#clearReelBtn");
    if (clearBtn) clearBtn.disabled = reelLocked || n === 0;
    var stashBtn = qs("#stashReelBtn");
    if (stashBtn) {
      stashBtn.disabled = reelLocked || n === 0;
      stashBtn.setAttribute(
        "data-tooltip",
        reelLocked
          ? "Finish generating first"
          : n === 0
            ? "Add clips to the reel first"
            : "Stash this reel to reuse later",
      );
    }
    var highlightsBtn = qs("#buildHighlightsBtn");
    if (highlightsBtn) highlightsBtn.disabled = reelLocked;
  }

  function setArtifactGenerating(active) {
    state.artifactGenerating = active;
    setTitleSpinner("artifactsSpinner", active);
    updateArtifactActions();
    updateReelActions();
  }

  function setReelGenerating(active) {
    state.reelGenerating = active;
    setTitleSpinner("reelSpinner", active);
    updateReelActions();
    updateArtifactActions();
  }

  function toggleArtifactCell(info) {
    if (isArtifactQueueLocked()) return;
    if (findInQueue(state.artifactQueue, info.participant, info.row) >= 0) {
      removeAllCellEntries(state.artifactQueue, info.participant, info.row);
      renderArtifactQueue();
      updateSingleCellClass(info.participant, info.row);
      return;
    }
    var entries = expandCellToSegments(info);
    if (entries.length === 0) return;
    for (var i = 0; i < entries.length; i++) state.artifactQueue.push(entries[i]);
    renderArtifactQueue();
    updateSingleCellClass(info.participant, info.row);
  }

  function toggleReelCell(info) {
    if (isReelQueueLocked()) return;
    if (findInQueue(state.reelQueue, info.participant, info.row) >= 0) {
      removeAllCellEntries(state.reelQueue, info.participant, info.row);
      renderReelQueue();
      updateSingleCellClass(info.participant, info.row);
      return;
    }
    var entries = expandCellToSegments(info);
    if (entries.length === 0) return;
    for (var i = 0; i < entries.length; i++) state.reelQueue.push(entries[i]);
    renderReelQueue();
    updateSingleCellClass(info.participant, info.row);
  }

  function addToQueue(targetQueue, info, renderFn) {
    if (targetQueue === state.artifactQueue && isArtifactQueueLocked()) return;
    if (targetQueue === state.reelQueue && isReelQueueLocked()) return;
    var added = false;
    if (isIntakeSource(info.source)) {
      if (findIntakeInQueue(targetQueue, info) < 0) {
        targetQueue.push(info);
        added = true;
      }
    } else if (info.segIdx !== undefined) {
      if (!hasSegmentInQueue(targetQueue, info.participant, info.row, info.segIdx)) {
        targetQueue.push(info);
        added = true;
      }
    } else if (findInQueue(targetQueue, info.participant, info.row) < 0) {
      var entries = expandCellToSegments(info);
      if (entries.length > 0) {
        for (var i = 0; i < entries.length; i++) targetQueue.push(entries[i]);
        added = true;
      }
    }
    if (added) {
      if (renderFn) renderFn();
      if (info.row) updateSingleCellClass(info.participant, info.row);
    }
  }

  // Collect selectable timestamp cell infos matching a filter
  function collectCellInfos(filterFn) {
    var infos = [];
    var cells = qsa(".ts-cell");
    for (var i = 0; i < cells.length; i++) {
      var td = cells[i];
      if (!isSelectableTimestampCell(td)) continue;
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
          if (entries.length === 0) continue;
          for (var m = 0; m < entries.length; m++) queue.push(entries[m]);
        }
      }
    }
    renderFn();
    for (var n = 0; n < infos.length; n++) {
      updateSingleCellClass(infos[n].participant, infos[n].row);
    }
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
          if (ev.shiftKey) {
            if (!isReelQueueLocked()) toggleBatchInQueue(state.reelQueue, allInfos, renderReelQueue);
          } else if (!isArtifactQueueLocked()) {
            toggleBatchInQueue(state.artifactQueue, allInfos, renderArtifactQueue);
          }
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
            if (ev.shiftKey) {
              if (!isReelQueueLocked()) toggleBatchInQueue(state.reelQueue, colInfos, renderReelQueue);
            } else if (!isArtifactQueueLocked()) {
              toggleBatchInQueue(state.artifactQueue, colInfos, renderArtifactQueue);
            }
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
          if (ev.shiftKey) {
            if (!isReelQueueLocked()) toggleBatchInQueue(state.reelQueue, rowInfos, renderReelQueue);
          } else if (!isArtifactQueueLocked()) {
            toggleBatchInQueue(state.artifactQueue, rowInfos, renderArtifactQueue);
          }
        }
        return;
      }

      // Single cell select
      var td = ev.target.closest(".ts-cell");
      if (!isSelectableTimestampCell(td)) return;
      var info = getCellInfo(td);

      if (ev.shiftKey) {
        toggleReelCell(info);
      } else {
        toggleArtifactCell(info);
      }
    });

    grid.addEventListener("contextmenu", function (ev) {
      var td = ev.target.closest(".ts-cell");
      if (!isSelectableTimestampCell(td)) return;
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
      // Anchor to the chip, not the td, so it reads as the chip widening.
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

  // Click-vs-drag threshold: the ghost mounts only after this much movement.
  var _CELL_DRAG_THRESHOLD_PX = 6;
  var _CELL_GHOST_OFFSET_X = 14;
  var _CELL_GHOST_OFFSET_Y = 10;

  // dragover silence that counts as release. Measured worst Safari gap 52ms; WebKit's dragend lags ~560ms.
  var _CELL_RELEASE_WATCHDOG_MS = 120;

  // Shared .queue-card-thumb. observe: lazy via ssObserveThumb; nativeLazy false for off-DOM ghosts; editItem enables trim.
  function buildQueueCardThumb(card, opts) {
    var thumb = el("div", "queue-card-thumb");
    // Window coordinates for the optional hover card scrubber (sprite + audio).
    thumb.dataset.participant = opts.participant;
    thumb.dataset.start = opts.start;
    thumb.dataset.end = Number(opts.start) + Number(opts.duration);
    var img = document.createElement("img");
    img.decoding = "async";
    img.alt = "";
    img.draggable = false;
    thumb.appendChild(img);
    if (opts.observe) {
      ssObserveThumb(card, img, thumb, opts.participant, opts.start);
    } else {
      img.src = "api/thumbnail/" + encodeURIComponent(opts.participant) + "/" + opts.start;
      if (opts.nativeLazy !== false) img.loading = "lazy";
      img.addEventListener("error", function () {
        this.remove();
        thumb.appendChild(el("span", "", "✕"));
        card.classList.add("queue-card-error");
      });
    }
    if (opts.editItem) {
      appendDurationBadge(thumb, opts.editItem, opts.renderFn);
    } else {
      thumb.appendChild(el("span", "queue-card-duration", formatDuration(opts.duration)));
    }
    card.appendChild(thumb);
    return thumb;
  }

  // ---- Card scrubber (opt-in: STUDIO_CARD_SCRUBBER) — studio-scrubber.js ----
  // Hub delegators; the implementations live in studio-scrubber.js.
  function attachQueueScrubbers() {
    return STUDIO.attachQueueScrubbers && STUDIO.attachQueueScrubbers.apply(null, arguments);
  }
  function resetScrubberPrefetch() {
    return STUDIO.resetScrubberPrefetch && STUDIO.resetScrubberPrefetch.apply(null, arguments);
  }

  // Source-origin badge (Screenspace / Transcript / Composer) layered over an
  // intake thumb.
  function buildSourceBadge(source) {
    var badge = el("span", "queue-card-source-badge" +
      (source === "composer" ? " source-composer" : "") +
      (source === "mindnode" ? " source-mindnode" : ""));
    var icon = source === "transcript" ? "bars-3"
      : source === "composer" ? "scissors"
      : source === "mindnode" ? "share"
      : "squares-2x2";
    badge.innerHTML = iconHTML(icon);
    return badge;
  }

  // Fixed overlay of one queue-style card per segment; cards cascade via --i (studio.css).
  function buildCellDragGhost(info, segments) {
    var ghost = el("div", "cell-drag-ghost");
    var n = segments.length;
    for (var i = 0; i < n; i++) {
      var seg = segments[i];
      var card = el("div", "queue-card cell-drag-ghost-card");
      card.style.setProperty("--i", i);
      if (info.severity) card.setAttribute("data-severity", info.severity);

      // Off-DOM drag image: must load now, so no loading="lazy".
      buildQueueCardThumb(card, {
        participant: info.participant,
        start: seg.start,
        duration: seg.end - seg.start,
        observe: false,
        nativeLazy: false,
      });

      var meta = el("div", "queue-card-meta");
      var refText = info.participant + "." + info.row;
      if (n > 1) refText += " (" + (i + 1) + "/" + n + ")";
      var metaRow = el("div", "queue-card-meta-row");
      metaRow.appendChild(el("span", "queue-card-ref", refText));
      meta.appendChild(metaRow);
      meta.appendChild(el("span", "queue-card-time", formatDuration(seg.start) + "–" + formatDuration(seg.end)));
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
    var releaseTimer = 0;      // dragover-gone deadline, see the constant above
    var blankDragImage = createBlankDragImage();

    grid.addEventListener("pointerdown", function (ev) {
      var td = ev.target.closest(".ts-cell");
      if (!isSelectableTimestampCell(td)) {
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
      if (!isSelectableTimestampCell(td)) return;

      var info = getCellInfo(td);
      ev.dataTransfer.setData("application/json", JSON.stringify(info));
      ev.dataTransfer.effectAllowed = "copy";

      // Suppress the browser's snapshot — we render a custom cascade overlay.
      try { ev.dataTransfer.setDragImage(blankDragImage, 0, 0); } catch (_) {}

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
      // pendingDrag outlives the mount so a watchdog misfire re-mounts on the next dragover.
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
      // Each dragover pushes the release deadline back.
      clearTimeout(releaseTimer);
      releaseTimer = setTimeout(hideGhost, _CELL_RELEASE_WATCHDOG_MS);
      ensureGhostBuilt();
      if (!ghost || rafPending) return;
      rafPending = requestAnimationFrame(function () {
        rafPending = 0;
        positionGhost();
      });
    }

    // Fade out and drop the node; pendingDrag stays so a misfire self-heals.
    function hideGhost() {
      if (rafPending) {
        cancelAnimationFrame(rafPending);
        rafPending = 0;
      }
      if (!ghost) return;
      var node = ghost;
      ghost = null;
      node.classList.remove("in");
      node.classList.add("out");
      setTimeout(function () {
        if (node.parentNode) node.parentNode.removeChild(node);
      }, 140);
    }

    function cleanup() {
      clearTimeout(releaseTimer);
      releaseTimer = 0;
      pendingDrag = null;
      hideGhost();
      pointerOrigin = null;
    }

    document.addEventListener("dragover", onDragOver, true);
    // dragend is the authoritative reset. No mouseup listener: neither engine fires it mid-drag.
    document.addEventListener("dragend", cleanup, true);
    document.addEventListener("drop", cleanup, true);
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
      if (isArtifactQueueLocked()) return;
      if (info.source === "reel-stash" || info.source === "artifact-stash") {
        for (var i = 0; i < info.items.length; i++)
          addToQueue(state.artifactQueue, info.items[i], null);
        renderArtifactQueue();
        return;
      }
      if (isIntakeSource(info.source)) {
        addToQueue(state.artifactQueue, info, renderArtifactQueue);
        return;
      }
      if (info.source === "reel") {
        removeFromQueue(state.reelQueue, info);
        renderReelQueue();
      }
      addToQueue(state.artifactQueue, info, renderArtifactQueue);
    });
    setupDropTarget(qs("#reelList"), function (info) {
      if (isReelQueueLocked()) return;
      if (info.source === "reel-stash" || info.source === "artifact-stash") {
        for (var i = 0; i < info.items.length; i++)
          addToQueue(state.reelQueue, info.items[i], null);
        renderReelQueue();
        return;
      }
      if (isIntakeSource(info.source)) {
        addToQueue(state.reelQueue, info, renderReelQueue);
        return;
      }
      if (info.source === "artifact") {
        removeFromQueue(state.artifactQueue, info);
        renderArtifactQueue();
      }
      addToQueue(state.reelQueue, info, renderReelQueue);
    });

    setupDropTarget(qs("#stashedReelsList"), function (info) {
      if (STUDIO.stashDropReel) STUDIO.stashDropReel(info);
    });
    setupDropTarget(qs("#stashedArtifactsList"), function (info) {
      if (STUDIO.stashDropArtifacts) STUDIO.stashDropArtifacts(info);
    });
  }

  function setupDropTarget(target, onDrop) {
    target.addEventListener("dragover", function (ev) {
      ev.preventDefault();
      ev.dataTransfer.dropEffect = "copy";
      // dragover fires ~60Hz; skip the no-op class write.
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
      clipgenWheelToHorizontal(qs(sel));
    });
  }

  // ---- Reel reordering ----

  var _reelDragIdx = null;

  function bindReelReorder() {
    var list = qs("#reelList");

    list.addEventListener("dragstart", function (ev) {
      if (isReelQueueLocked()) {
        ev.preventDefault();
        return;
      }
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
          severity: reelItem.severity,
          segIdx: reelItem.segIdx,
          start: reelItem.start,
          end: reelItem.end,
          segTotal: reelItem.segTotal,
          source: "reel",
        };
        // Keep intake identity so a drop back into the queue keeps its linkage.
        if (reelItem.event_type) data.event_type = reelItem.event_type;
        if (reelItem.event_ids) data.event_ids = reelItem.event_ids;
        if (reelItem.mark_ids) data.mark_ids = reelItem.mark_ids;
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
      if (isReelQueueLocked()) return;
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

  // ---- Trim pop-over — studio-trim.js ----
  // Guarded delegators; the satellite reaches saveQueues/isIntakeSource via STUDIO.
  function appendDurationBadge() {
    return STUDIO.appendDurationBadge && STUDIO.appendDurationBadge.apply(null, arguments);
  }
  function buildCellOverrides() {
    return STUDIO.buildCellOverrides && STUDIO.buildCellOverrides.apply(null, arguments);
  }

  // ---- Queue rendering ----

  var ARTIFACT_QUEUE = {
    listSel: "#artifactsList",
    countSel: "#artifactsCount",
    queueKey: "artifactQueue",
    isLocked: isArtifactQueueLocked,
    emptyGhost: "Click or drag cells here to queue for generation",
    updateActions: updateArtifactActions,
    isReel: false,
    attachDragstart: true,
    durationSel: null,
    countNoun: "artifact",
    countNounPlural: "artifacts",
  };
  var REEL_QUEUE = {
    listSel: "#reelList",
    countSel: "#reelCount",
    queueKey: "reelQueue",
    isLocked: isReelQueueLocked,
    emptyGhost: "Shift+click or drag cells here to build a reel",
    updateActions: updateReelActions,
    isReel: true,
    attachDragstart: false,
    durationSel: "#reelDuration",
    countNoun: "clip",
    countNounPlural: "clips",
  };

  // Composer-trim key or null; sheet keys match composer-markers.js (row:participant:segIdx), intakes key on ids.
  function queueItemTrimKey(item) {
    // Gate on carded keys so the deep-link has a card to land on.
    var cardKeys = state.coTrimCardKeys || {};
    if (!isIntakeSource(item.source) && item.row) {
      var key = "sheet:" + item.row + ":" + item.participant + ":" + (item.segIdx || 0);
      return cardKeys[key] ? key : null;
    }
    var ids = item.source === "screenspace" ? item.event_ids
      : item.source === "transcript" ? item.mark_ids
      : null;
    if (!ids) return null;
    var prefix = item.source === "screenspace" ? "screenspace:" : "transcript-mark:";
    for (var i = 0; i < ids.length; i++) {
      if (ids[i] && cardKeys[prefix + ids[i]]) return prefix + ids[i];
    }
    return null;
  }

  function buildQueueCard(item, idx, cfg, ctx) {
    var isIntake = isIntakeSource(item.source);
    var segTotal = item.segTotal || 1;
    var segIdx = item.segIdx || 0;

    var card = el(
      "div",
      "queue-card" + (cfg.isReel ? " reel-card" : "") + (isIntake ? " queue-card-intake" : ""),
    );
    if (cfg.isReel) card.setAttribute("data-reel-idx", idx);
    card.setAttribute("data-queue-idx", idx);
    card.setAttribute("data-participant", item.participant);
    card.setAttribute("data-row", isIntake ? "" : item.row);
    if (isIntake) card.setAttribute("data-source", item.source);
    if (!isIntake && item.severity) card.setAttribute("data-severity", item.severity);
    card.setAttribute("data-seg-idx", segIdx);
    if (!ctx.locked) card.setAttribute("draggable", "true");

    var thumb = buildQueueCardThumb(card, {
      participant: item.participant,
      start: item.start,
      duration: item.end - item.start,
      observe: isIntake,
      editItem: item,
      renderFn: ctx.render,
    });
    if (isIntake) thumb.appendChild(buildSourceBadge(item.source));

    // Trim badge: the timestamp has a Composer trim; click jumps to its Intake card.
    var trimKey = queueItemTrimKey(item);
    if (trimKey) {
      var trimBadge = el("button", "intake-trim-badge");
      trimBadge.innerHTML = iconHTML("scissors");
      trimBadge.type = "button";
      trimBadge.title = "Trimmed in Composer. Click to view the trimmed version";
      trimBadge.setAttribute("aria-label", "Show trimmed version in Composer Intake");
      trimBadge.addEventListener("click", function (ev) {
        ev.stopPropagation();
        focusComposerIntakeItem(trimKey);
      });
      thumb.appendChild(trimBadge);
    }

    var meta = el("div", "queue-card-meta");
    var refText;
    if (isIntake) {
      refText = item.participant + " \u00b7 " + (item.event_type || item.desc || "intake");
    } else {
      refText = item.participant + "." + item.row;
      if (segTotal > 1) refText += " (" + (segIdx + 1) + "/" + segTotal + ")";
    }
    var metaRow = el("div", "queue-card-meta-row");
    if (cfg.isReel) metaRow.appendChild(el("span", "reel-card-order", String(idx + 1)));
    metaRow.appendChild(el("span", "queue-card-ref", refText));
    meta.appendChild(metaRow);
    meta.appendChild(el("span", "queue-card-time", formatDuration(item.start) + "\u2013" + formatDuration(item.end)));
    card.appendChild(meta);

    var removeBtn = el("button", "queue-card-remove");
    removeBtn.innerHTML = iconHTML("x-mark");
    removeBtn.title = "Remove";
    card.appendChild(removeBtn);

    return card;
  }

  function renderQueue(cfg) {
    return clipgenPerf.span("studio.renderQueue", function () {
      renderQueueImpl(cfg);
    });
  }

  // A card is one segment; spell out cells vs cards so the counts agree.
  function queueCountTooltip(q, noun, nounPlural) {
    var cellsSeen = {};
    var cellCount = 0;
    for (var i = 0; i < q.length; i++) {
      if (isIntakeSource(q[i].source)) continue;
      var key = q[i].participant + "." + q[i].row;
      if (!cellsSeen[key]) { cellsSeen[key] = true; cellCount++; }
    }
    var phrase = clipgenPluralUnit(q.length, noun, nounPlural);
    if (cellCount > 0 && cellCount !== q.length) {
      return phrase + " from " + clipgenPluralUnit(cellCount, "cell", "cells");
    }
    return phrase + " queued";
  }

  function renderQueueImpl(cfg) {
    clearGridHighlights();
    var list = qs(cfg.listSel);
    var q = state[cfg.queueKey];
    var n = q.length;
    var countEl = qs(cfg.countSel);
    countEl.textContent = "(" + n + ")";
    countEl.setAttribute(
      "data-tooltip",
      queueCountTooltip(q, cfg.countNoun, cfg.countNounPlural)
    );
    list.innerHTML = "";
    saveQueues();
    refreshIntakeCardStates();

    if (n === 0) {
      list.appendChild(el("div", "queue-card-ghost", cfg.emptyGhost));
      if (cfg.durationSel) qs(cfg.durationSel).textContent = "";
      cfg.updateActions();
      return;
    }

    var locked = cfg.isLocked();
    var render = function () {
      renderQueue(cfg);
    };
    var totalDur = 0;
    var frag = document.createDocumentFragment();
    for (var i = 0; i < n; i++) {
      totalDur += q[i].end - q[i].start;
      frag.appendChild(buildQueueCard(q[i], i, cfg, { locked: locked, render: render }));
    }
    list.appendChild(frag);
    if (cfg.durationSel) qs(cfg.durationSel).textContent = formatDuration(totalDur);
    applyCardStates(list);
    attachQueueScrubbers(list);
    cfg.updateActions();
  }

  function renderArtifactQueue() {
    renderQueue(ARTIFACT_QUEUE);
  }

  function renderReelQueue() {
    renderQueue(REEL_QUEUE);
  }

  // One listener set per list, bound once at boot, never from renderQueue (CODE-REVIEW.md).
  function bindQueueList(cfg) {
    var list = qs(cfg.listSel);
    if (!list) return;

    if (cfg.attachDragstart) {
      list.addEventListener("dragstart", function (ev) {
        if (cfg.isLocked()) {
          ev.preventDefault();
          return;
        }
        var card = ev.target.closest(".queue-card");
        if (!card || !list.contains(card)) return;
        var idx = parseInt(card.getAttribute("data-queue-idx"), 10);
        var item = state[cfg.queueKey][idx];
        if (!item) return;
        var isIntake = isIntakeSource(item.source);
        var data = {
          participant: item.participant,
          desc: item.desc,
          start: item.start,
          end: item.end,
          source: isIntake ? item.source : "artifact",
        };
        if (!isIntake) {
          data.row = item.row;
          data.timestamp = item.timestamp;
          data.severity = item.severity;
          data.segIdx = item.segIdx;
          data.segTotal = item.segTotal;
        } else {
          data.event_type = item.event_type;
          data.event_ids = item.event_ids;
          data.mark_ids = item.mark_ids;
        }
        ev.dataTransfer.setData("application/json", JSON.stringify(data));
        ev.dataTransfer.effectAllowed = "copyMove";
        setCardDragImage(ev, card);
      });
    }

    list.addEventListener("click", function (ev) {
      var btn = ev.target.closest(".queue-card-remove");
      if (!btn || !list.contains(btn)) return;
      if (cfg.isLocked()) return;
      ev.stopPropagation();
      var card = btn.closest(".queue-card");
      if (!card) return;
      var idx = parseInt(card.getAttribute("data-queue-idx"), 10);
      var item = state[cfg.queueKey][idx];
      if (!item) return;
      var commit = function () {
        // Resolve by identity: an earlier removal can shift indices during the exit animation.
        var q = state[cfg.queueKey];
        var ix = q.indexOf(item);
        if (ix < 0) return;
        var removed = q.splice(ix, 1)[0];
        if (removed.row) delete state.cellResults[cellKey(removed.participant, removed.row)];
        renderQueue(cfg);
        if (removed.row) updateSingleCellClass(removed.participant, removed.row);
      };
      if (window.ClipgenMotion) ClipgenMotion.animateOut(card, "delete").then(commit);
      else commit();
    });

    list.addEventListener("mouseover", function (ev) {
      var card = ev.target.closest(".queue-card");
      if (!card || !list.contains(card)) return;
      if (card.classList.contains("queue-card-intake") || !card.getAttribute("data-row")) {
        clearGridHighlights();
        return;
      }
      highlightGridHeaders(card.getAttribute("data-participant"), parseInt(card.getAttribute("data-row"), 10));
    });
    list.addEventListener("mouseleave", clearGridHighlights);
  }

  // ---- Stashes — studio-stash.js ----
  // Guarded delegators; drop targets reach STUDIO.stashDropReel/.stashDropArtifacts late-bound.
  function loadStashes() { return STUDIO.loadStashes && STUDIO.loadStashes.apply(null, arguments); }
  function loadArtifactStashes() { return STUDIO.loadArtifactStashes && STUDIO.loadArtifactStashes.apply(null, arguments); }
  function stashCurrentReel() { return STUDIO.stashCurrentReel && STUDIO.stashCurrentReel.apply(null, arguments); }
  function stashCurrentArtifacts() { return STUDIO.stashCurrentArtifacts && STUDIO.stashCurrentArtifacts.apply(null, arguments); }
  function revealEmptyStashAreas() { return STUDIO.revealEmptyStashAreas && STUDIO.revealEmptyStashAreas.apply(null, arguments); }
  function hideEmptyStashAreas() { return STUDIO.hideEmptyStashAreas && STUDIO.hideEmptyStashAreas.apply(null, arguments); }

  // ---- Buttons ----

  // Empty the queue (also the Clear hotkeys); cards animate out before the commit.
  function clearArtifacts() {
    if (isArtifactQueueLocked()) return;
    var cards = qsa("#artifactsList .queue-card");
    var commit = function () {
      var cleared = state.artifactQueue.slice();
      for (var i = 0; i < cleared.length; i++) {
        delete state.cellResults[cellKey(cleared[i].participant, cleared[i].row)];
      }
      state.artifactQueue = [];
      renderArtifactQueue();
      for (var u = 0; u < cleared.length; u++) {
        if (cleared[u].row) updateSingleCellClass(cleared[u].participant, cleared[u].row);
      }
    };
    if (cards.length && window.ClipgenMotion) ClipgenMotion.animateOutAll(cards, "delete").then(commit);
    else commit();
  }

  function clearReel() {
    if (isReelQueueLocked()) return;
    var cards = qsa("#reelList .queue-card");
    var commit = function () {
      var cleared = state.reelQueue.slice();
      for (var i = 0; i < cleared.length; i++) {
        delete state.cellResults[cellKey(cleared[i].participant, cleared[i].row)];
      }
      state.reelQueue = [];
      renderReelQueue();
      for (var u = 0; u < cleared.length; u++) {
        if (cleared[u].row) updateSingleCellClass(cleared[u].participant, cleared[u].row);
      }
    };
    if (cards.length && window.ClipgenMotion) ClipgenMotion.animateOutAll(cards, "delete").then(commit);
    else commit();
  }

  function bindButtons() {
    qs("#clearArtifactsBtn").addEventListener("click", clearArtifacts);

    qs("#addToReelBtn").addEventListener("click", function () {
      for (var i = 0; i < state.artifactQueue.length; i++) {
        addToQueue(state.reelQueue, state.artifactQueue[i], null);
      }
      renderReelQueue();
    });

    qs("#clearReelBtn").addEventListener("click", clearReel);

    qs("#stashReelBtn").addEventListener("click", stashCurrentReel);
    qs("#stashArtifactsBtn").addEventListener("click", stashCurrentArtifacts);
    qs("#generateBtn").addEventListener("click", onGenerate);
    qs("#cancelGenerateBtn").addEventListener("click", onCancelGenerate);
    qs("#buildReelBtn").addEventListener("click", onBuildReel);
    qs("#cancelReelBtn").addEventListener("click", onCancelReel);
    qs("#buildHighlightsBtn").addEventListener("click", onBuildHighlights);
    bindGalleryDialog();

    // Hotkeys gate on the toolbar button so queue-lock logic stays in one place.
    function hotkeyBtnEnabled(sel) {
      var b = qs(sel);
      return !!(b && !b.disabled);
    }
    window.ClipgenHotkeys.register([
      {
        id: "global.primary",
        when: function () { return hotkeyBtnEnabled("#generateBtn"); },
        handler: function () { onGenerate(); },
      },
      {
        id: "studio.buildReel",
        when: function () { return hotkeyBtnEnabled("#buildReelBtn"); },
        handler: function () { onBuildReel(); },
      },
      {
        id: "studio.buildHighlights",
        when: function () { return hotkeyBtnEnabled("#buildHighlightsBtn"); },
        handler: function () { onBuildHighlights(); },
      },
      // Route through the button so the busy guard and tab dispatch apply.
      {
        id: "global.refresh",
        handler: function () {
          var btn = qs("#studioRefresh");
          if (btn) btn.click();
        },
      },
      { id: "nav.next", handler: function () { return kbStep(1); } },
      { id: "nav.prev", handler: function () { return kbStep(-1); } },
      { id: "studio.moveRight", handler: function () { return kbStep(1); } },
      { id: "studio.moveLeft", handler: function () { return kbStep(-1); } },
      { id: "studio.moveDown", handler: function () { return kbStepVertical(1); } },
      { id: "studio.moveUp", handler: function () { return kbStepVertical(-1); } },
      { id: "studio.sendArtifacts", handler: function () { return kbSend(false); } },
      { id: "studio.sendReel", handler: function () { return kbSend(true); } },
      { id: "studio.removeCard", repeat: false, handler: function () { return kbRemoveCard(); } },
      { id: "studio.togglePanel", handler: function () { toggleBottomPanel(); } },
      {
        id: "studio.toggleSidebar",
        when: function () { return state.activePreviewTab === "sheet"; },
        handler: function () { toggleSidebar(); },
      },
      {
        id: "studio.stashArtifacts",
        when: function () { return hotkeyBtnEnabled("#stashArtifactsBtn"); },
        handler: function () { stashCurrentArtifacts(); },
      },
      {
        id: "studio.stashReel",
        when: function () { return hotkeyBtnEnabled("#stashReelBtn"); },
        handler: function () { stashCurrentReel(); },
      },
      {
        id: "studio.clearArtifacts",
        when: function () { return hotkeyBtnEnabled("#clearArtifactsBtn"); },
        handler: function () { clearArtifacts(); },
      },
      {
        id: "studio.clearReel",
        when: function () { return hotkeyBtnEnabled("#clearReelBtn"); },
        handler: function () { clearReel(); },
      },
      {
        id: "studio.focusFilter",
        when: function () { return state.activePreviewTab === "sheet"; },
        handler: function () { return kbJumpTo("filter"); },
      },
      { id: "studio.focusArtifacts", handler: function () { return kbJumpTo("artifact-queue"); } },
      { id: "studio.focusReel", handler: function () { return kbJumpTo("reel-queue"); } },
      { id: "studio.focusArtifactStash", handler: function () { return kbJumpTo("artifact-stash"); } },
      { id: "studio.focusReelStash", handler: function () { return kbJumpTo("reel-stash"); } },
      {
        id: "studio.selectTab",
        handler: function (e, combo) {
          // 1…4 → Nth preview tab; hidden tabs no-op. .click() reuses tab persistence.
          var n = parseInt(combo, 10);
          if (isNaN(n)) return;
          var tab = qsa(".preview-tab")[n - 1];
          if (tab && !tab.classList.contains("hidden")) tab.click();
        },
      },
    ]);

    // Alt-hold hint chips for the send actions, beside the selected cell/card.
    window.ClipgenHotkeys.registerActionHints(function () {
      var cur = kbCursorEl();
      if (!cur) return null;
      return { anchor: cur, entries: kbActionHintEntries() };
    });

    // Escape cascade: cancel an open highlights-parameter drawer first, then
    // clear the keyboard cursor.
    window.ClipgenHotkeys.registerEscape(function () {
      if (cancelHighlightsDrawer()) return true;
      return kbClearCursor();
    });

    qs("#artifactFormat").addEventListener("change", function () {
      var tcGroup = qs("#titlecardGroup");
      if (tcGroup) {
        if (this.value === "clip") tcGroup.classList.remove("hidden");
        else tcGroup.classList.add("hidden");
      }
      // Refresh the Generate tooltip so its noun tracks the selected format.
      updateArtifactActions();
    });

    qs("#titlecardEnabled").addEventListener("change", persistTitlecardSettings);
    qs("#titlecardDuration").addEventListener("change", persistTitlecardSettings);

    qs("#studioRefresh").addEventListener("click", refreshActiveTab);
    if (window.wireSettingsButton) {
      window.wireSettingsButton({
        initialTab: "General",
        // Click-time function: the sheet version loads after boot.
        version: function () { return state.sheetData ? state.sheetData.version : ""; },
        onApply: function (_applied, full) {
          state.settingsData = full;
          syncInlineControls();
          _syncMarkCategoriesFromSettings(full);
        },
      });
    }

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

    qs("#buildStatusDismiss").addEventListener("click", hideBuildStatus);
    qs("#buildStatusOpen").addEventListener("click", function () {
      if (_buildStatusFile) {
        apiPost("api/open-viewer", { file: _buildStatusFile }).catch(function () {});
      }
      hideBuildStatus();
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

  var setButtonProgress = ClipgenPrimitives.setButtonProgress;

  function createPulserOverlay() {
    var overlay = el("div", "card-gen-overlay");
    // Intentional inline SVG (icon convention exception): a three-dot "generating" pulse animation, not an icon.
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
    var badge = card.querySelector(".card-gen-badge");
    if (badge) badge.remove();
  }

  function setCardResult(card, success, reason) {
    card.classList.remove("queue-card-queued");
    card.classList.add(success ? "queue-card-success" : "queue-card-fail");
    // Surface the per-item failure reason as a native tooltip on the card.
    if (!success && reason) card.title = reason;
    else if (success) card.removeAttribute("title");
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

  // ---- API calls ----

  // ---- Artifact generation — studio-generate.js ----
  // Delegators below; shared painters/trackers publish at the tail.

  // ---- Elapsed-time tracking ----
  // Elapsed only, no ETA; idempotent start() survives a job-status reattach.
  var _reelEtaTracker = createEtaTracker();
  var _generateEtaTracker = createEtaTracker();
  var _buildEtaTracker = createEtaTracker();
  var _studioEtaTicker = createIntervalTicker(_tickStudioEta, {
    isActive: isAnyStudioJobRunning,
  });
  var _genLastDone = 0;
  var _genLastTotal = 0;

  function _paintReelElapsed() {
    var el = qs("#reelEtaText");
    if (!el) return;
    if (!state.reelGenerating) {
      el.classList.add("hidden");
      el.textContent = "";
      return;
    }
    el.classList.remove("hidden");
    el.textContent = formatDuration(_reelEtaTracker.update().elapsedSec);
  }

  function _paintBuildElapsed() {
    var el = qs("#buildElapsed");
    if (!el) return;
    if (!state.overlayJobRunning) {
      el.textContent = "";
      return;
    }
    el.textContent = formatDuration(_buildEtaTracker.update().elapsedSec);
  }

  function _paintGenerateProgress() {
    var el = qs("#generateProgress");
    if (!el) return;
    var hasCount = _genLastTotal > 0;
    // No countable work: elapsed-only while generating, hidden once idle.
    if (!state.artifactGenerating && !hasCount) {
      el.classList.add("hidden");
      el.textContent = "";
      return;
    }
    el.classList.remove("hidden");
    // Shimmer only while live; the final tally lingers and should not move.
    el.classList.toggle("cg-shimmer", state.artifactGenerating);
    var parts = [];
    // Artifacts, not cells, to agree with the card-count badge beside it.
    if (hasCount) {
      parts.push(
        _genLastDone + " / " + clipgenPluralUnit(_genLastTotal, "artifact", "artifacts")
      );
    }
    if (state.artifactGenerating) {
      parts.push(formatDuration(_generateEtaTracker.update().elapsedSec));
    }
    el.textContent = parts.join(" · ");
  }

  function _tickStudioEta() {
    // The ticker's isActive guard stops it once no job is live.
    _paintReelElapsed();
    _paintGenerateProgress();
    _paintBuildElapsed();
  }

  function updateGenerateProgress(done, total) {
    _genLastDone = done;
    _genLastTotal = total;
    _paintGenerateProgress();
  }

  function onGenerate() { return STUDIO.onGenerate && STUDIO.onGenerate.apply(null, arguments); }
  function onCancelGenerate() { return STUDIO.onCancelGenerate && STUDIO.onCancelGenerate.apply(null, arguments); }

  function onCancelReel() {
    qs("#cancelReelBtn").classList.add("hidden");
    apiPost("api/reel/cancel").catch(toastError("Cancel failed"));
  }

  // ---- API: reel + standalone viewers (timeline / HTML viewer) ----

  function onBuildReel() {
    if (state.reelGenerating || state.reelQueue.length === 0) return;
    setReelGenerating(true);
    qs("#cancelReelBtn").classList.remove("hidden");
    _reelEtaTracker.reset();
    _reelEtaTracker.start();
    _studioEtaTicker.ensure();
    _paintReelElapsed();

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
      var reelOverrides = buildCellOverrides(state.reelQueue);
      if (Object.keys(reelOverrides).length > 0) reelBody.overrides = reelOverrides;
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

    // Two phases: clips 0.7, concat 0.3; clips finish first, so the bar stays monotonic.
    var totalClips = 0;
    var clipsDone = 0;
    var concatFraction = 0;
    var finalPayload = null;
    var cancelled = false;
    var finished = false;

    function updateProgress() {
      var clipFraction = totalClips > 0 ? Math.min(clipsDone / totalClips, 1) : 0;
      var overall = clipFraction * 0.7 + concatFraction * 0.3;
      setButtonProgress("buildReelBtn", overall);
    }

    function finish() {
      if (finished) return;
      finished = true;
      setReelGenerating(false);
      qs("#cancelReelBtn").classList.add("hidden");
      setButtonProgress("buildReelBtn", null);
      _reelEtaTracker.reset();
      _paintReelElapsed();

      var data = finalPayload || {};
      var isCancelled = cancelled || !!data.cancelled;

      var cards = list.querySelectorAll(".queue-card");
      for (var j = 0; j < cards.length; j++) {
        if (isCancelled) {
          clearCardStatus(cards[j]);
        } else {
          setCardResult(cards[j], !!data.ok);
        }
      }

      if (isCancelled) {
        showResult(null, "Reel generation cancelled");
      } else if (data.ok) {
        showResult("Reel built successfully", null);
      } else {
        showResult(null, data.error || "Reel build failed");
      }
      revealStatusOverlay();
    }

    function handleLine(line) {
      var data;
      try { data = JSON.parse(line); } catch (e) { return; }
      if (!data) return;
      if (data.cancelled) cancelled = true;
      if (data.phase === "start") {
        totalClips = data.total_clips || 0;
        updateProgress();
      } else if (data.phase === "clip_done") {
        clipsDone += 1;
        updateProgress();
      } else if (data.phase === "concat") {
        concatFraction = typeof data.progress === "number" ? data.progress : 0;
        updateProgress();
      } else if (data.phase === "done") {
        // Stream-copy concat emits no progress; fill to 100% before the final line.
        concatFraction = 1;
        updateProgress();
      } else if (data.ok !== undefined || data.error !== undefined) {
        finalPayload = data;
        if (data.ok && Array.isArray(data.reels)) {
          for (var ri = 0; ri < data.reels.length; ri++) {
            state.generatedReels.push(stampLog(data.reels[ri]));
          }
        }
      }
    }

    apiPostNDJSON(endpoint, reelBody, { onLine: handleLine })
      .then(finish)
      .catch(function (err) {
        // 4xx/5xx (e.g. 409 in-progress) arrive as JSON in err.bodyText; parse into the payload.
        if (err && err.status >= 400) {
          try {
            finalPayload = JSON.parse(err.bodyText);
          } catch (_) {
            finalPayload = {
              ok: false,
              error: err.bodyText || ("HTTP " + err.status),
            };
          }
        } else {
          finalPayload = { ok: false, error: "Request failed: " + err };
        }
        finish();
      });
  }

  function onBuildViewer() {
    if (isAnyStudioJobRunning() || state.generatedArtifacts.length === 0) return;
    state.overlayJobRunning = true;

    showBuildStatus("Building timeline viewer…", null);

    apiPost("api/viewer", {})
      .then(function (data) {
        state.overlayJobRunning = false;
        if (data.ok) {
          state.generatedViewers.push(stampLog({
            type: "viewer",
            subtype: "viewer",
            file: pathBasename(data.file),
            description: "Timeline viewer",
          }));
          showBuildResult("Viewer created: " + (data.file || ""), null, data.file);
        } else {
          showBuildResult(null, data.error || "Viewer build failed");
        }
      })
      .catch(function (err) {
        state.overlayJobRunning = false;
        showBuildResult(null, "Request failed: " + err);
      });
  }

  function onBuildTimelineViewer() {
    if (isAnyStudioJobRunning()) return;

    var ssCount = (state.intakeClusters || []).length;
    var trCount = (state.trIntakeClusters || []).length;
    if (ssCount === 0 && trCount === 0) {
      startTimelineViewerBuild(false);
      return;
    }

    var parts = [];
    if (ssCount > 0) {
      parts.push(ssCount + " Screenspace event group" + (ssCount === 1 ? "" : "s"));
    }
    if (trCount > 0) {
      parts.push(trCount + " Transcript mark group" + (trCount === 1 ? "" : "s"));
    }
    var msg = parts.join(" and ") + " detected. Include them as clips in the timeline viewer?";

    showConfirm(
      "Include Intake Events?",
      msg,
      function () { startTimelineViewerBuild(true); },
      function () { startTimelineViewerBuild(false); }
    );
  }

  function startTimelineViewerBuild(includeIntake) {
    state.overlayJobRunning = true;
    state.timelineViewerCancelledByUser = false;
    var body = {};

    var ssClusters = state.intakeClusters || [];
    var trClusters = state.trIntakeClusters || [];
    var hasIntake = includeIntake && (ssClusters.length > 0 || trClusters.length > 0);

    if (hasIntake) {
      showBuildStatus(
        "Building timeline viewer with intake events\u2026",
        onCancelTimelineViewer
      );
      body.include_intake = true;
      var items = ssClusters.map(function (c) {
        return {
          participant: c.participant,
          start: c.start,
          end: c.end,
          event_type: c.event_type,
          event_ids: c.events.map(function (e) { return e.id; }),
        };
      });
      for (var i = 0; i < trClusters.length; i++) {
        var c = trClusters[i];
        items.push({
          participant: c.participant,
          start: c.start,
          end: c.end,
          event_type: c.category || "transcript",
          source: "transcript",
          mark_ids: c.marks.map(function (m) { return m.id; }),
          text: c.text || "",
          label: c.label || "",
        });
      }
      body.intake_items = items;
    } else {
      showBuildStatus("Building timeline viewer\u2026", onCancelTimelineViewer);
    }

    apiPost("api/timeline-viewer", body)
      .then(function (data) {
        state.overlayJobRunning = false;
        if (data.cancelled || state.timelineViewerCancelledByUser) {
          state.timelineViewerCancelledByUser = false;
          hideBuildStatus();
          showToast("Build cancelled");
          return;
        }
        if (data.ok) {
          state.generatedViewers.push(stampLog({
            type: "viewer",
            subtype: "timeline-viewer",
            file: pathBasename(data.file),
            description: "Timeline viewer (full sheet)",
          }));
          var msg = "Timeline viewer created: " + (data.file || "");
          if (data.generated) {
            msg = "Generated " + clipgenPluralUnit(data.generated, "clip", "clips") + ". " + msg;
          }
          showBuildResult(msg, null, data.file);
        } else {
          showBuildResult(null, data.error || "Timeline viewer build failed");
        }
      })
      .catch(function (err) {
        state.overlayJobRunning = false;
        if (state.timelineViewerCancelledByUser) {
          state.timelineViewerCancelledByUser = false;
          hideBuildStatus();
          showToast("Build cancelled");
          return;
        }
        showBuildResult(null, "Request failed: " + err);
      });
  }

  function onCancelTimelineViewer() {
    state.timelineViewerCancelledByUser = true;
    apiPost("api/timeline-viewer/cancel").catch(toastError("Cancel failed"));
  }

  var _highlightsBtnOrigHTML = "";

  // Close the highlights drawer without running (Escape); returns whether one was open.
  function cancelHighlightsDrawer() {
    var drawer = qs("#highlightsDurationDrawer");
    if (!drawer || !drawer.classList.contains("open")) return false;
    if (document.activeElement && drawer.contains(document.activeElement)) {
      document.activeElement.blur();
    }
    drawer.classList.remove("open");
    var btn = qs("#buildHighlightsBtn");
    if (btn) {
      btn.style.minWidth = "";
      if (_highlightsBtnOrigHTML) btn.innerHTML = _highlightsBtnOrigHTML;
    }
    return true;
  }

  function onBuildHighlights() {
    if (isAnyStudioJobRunning()) return;

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
    if (!Number.isFinite(duration) || duration < 1) duration = 180;

    drawer.classList.remove("open");
    btn.style.minWidth = "";
    btn.innerHTML = _highlightsBtnOrigHTML;

    setReelGenerating(true);
    showOverlay("Finding best clips (" + duration + "s budget)...");

    apiPost("api/highlights-preview", { highlights_duration: duration })
      .then(function (data) {
        setReelGenerating(false);
        if (data.ok && data.clips && data.clips.length > 0) {
          var prev = state.reelQueue.slice();
          state.reelQueue = [];
          for (var i = 0; i < data.clips.length; i++) {
            var entries = expandCellToSegments(data.clips[i]);
            for (var ei = 0; ei < entries.length; ei++) {
              state.reelQueue.push(entries[ei]);
            }
          }
          renderReelQueue();
          var touchedKeys = {};
          for (var p = 0; p < prev.length; p++) {
            if (prev[p].row) touchedKeys[cellKey(prev[p].participant, prev[p].row)] = prev[p];
          }
          for (var q = 0; q < state.reelQueue.length; q++) {
            var rq = state.reelQueue[q];
            if (rq.row) touchedKeys[cellKey(rq.participant, rq.row)] = rq;
          }
          for (var key in touchedKeys) {
            updateSingleCellClass(touchedKeys[key].participant, touchedKeys[key].row);
          }
          showResult(
            "Added " + clipgenPluralUnit(data.clips.length, "clip", "clips") + " to reel queue",
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
        setReelGenerating(false);
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
    if (isAnyStudioJobRunning()) return;
    var overlay = qs("#galleryOverlay");
    if (!overlay) return;
    popModalIn(overlay, qs(".gallery-card"));
    openModalTrap(overlay, closeGalleryDialog);
    var sel = qs("#galleryParticipant");
    if (sel) sel.focus();
  }

  function closeGalleryDialog() {
    var overlay = qs("#galleryOverlay");
    if (!overlay) return;
    // Trap released now; the visual hide trails the fade.
    closeModalTrap(overlay);
    popModalOut(overlay, qs(".gallery-card"), function () {
      overlay.classList.add("hidden");
    });
  }

  function submitGalleryDialog() {
    if (isAnyStudioJobRunning()) return;

    var participant = qs("#galleryParticipant").value;
    var format = qs("#galleryFormat").value;
    var interval = parseInt(qs("#galleryInterval").value, 10);
    if (!interval || interval < 1) interval = 10;
    var bundle = qs("#galleryBundle").checked;

    if (!participant) {
      showToast("No participant selected for gallery");
      return;
    }

    closeGalleryDialog();
    state.overlayJobRunning = true;
    state.galleryCancelledByUser = false;
    showBuildStatus(
      "Generating gallery viewer for " + participant + "…",
      onCancelGallery
    );

    apiPost("api/gallery", { participant: participant, format: format, interval: interval, bundle: bundle })
      .then(function (data) {
        state.overlayJobRunning = false;
        if (data.cancelled || state.galleryCancelledByUser) {
          state.galleryCancelledByUser = false;
          hideBuildStatus();
          showToast("Build cancelled");
          return;
        }
        if (data.ok) {
          state.generatedViewers.push(stampLog({
            type: "viewer",
            subtype: "gallery",
            file: pathBasename(data.file),
            participant: participant,
            description: "Gallery viewer (" + format + ", " + interval + "s)",
          }));
          showBuildResult("Gallery viewer created: " + (data.file || ""), null, data.file);
        } else {
          showBuildResult(null, data.error || "Gallery build failed");
        }
      })
      .catch(function (err) {
        state.overlayJobRunning = false;
        if (state.galleryCancelledByUser) {
          state.galleryCancelledByUser = false;
          hideBuildStatus();
          showToast("Build cancelled");
          return;
        }
        showBuildResult(null, "Request failed: " + err);
      });
  }

  function onCancelGallery() {
    state.galleryCancelledByUser = true;
    apiPost("api/gallery/cancel").catch(toastError("Cancel failed"));
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
    // Escape is handled by the modal focus trap opened in openGalleryDialog.
  }

  // ---- Modal focus trap ----
  // openBlockingModal delegators; Studio never stacks modals, release() is idempotent.
  function openModalTrap(overlayEl, onEscape) {
    return openBlockingModal(overlayEl, {
      onEscape: onEscape,
      trapFocus: true,
      restoreFocus: true,
    });
  }

  function closeModalTrap(overlayEl) {
    closeBlockingModal(overlayEl);
  }


  // ---- Status overlay ----

  var _lastViewerFile = "";

  function revealStatusOverlay() {
    var overlay = qs("#statusOverlay");
    popModalIn(overlay, qs(".status-card"));
    openModalTrap(overlay, hideOverlay);
  }

  function showOverlay(message) {
    qs("#statusSpinner").style.display = "";
    qs("#statusTitle").classList.add("cg-shimmer");
    qs("#statusTitle").textContent = message;
    qs("#statusMessage").textContent = "";
    qs("#statusMessage").className = "";
    qs("#statusDismiss").classList.add("hidden");
    qs("#statusOpen").classList.add("hidden");
    _lastViewerFile = "";
    revealStatusOverlay();
  }

  function showResult(successMsg, errorMsg, filePath) {
    qs("#statusSpinner").style.display = "none";
    qs("#statusTitle").classList.remove("cg-shimmer"); // Error / Done are resting states
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
    var overlay = qs("#statusOverlay");
    // Release the trap immediately — only the visual hide waits for the fade.
    closeModalTrap(overlay);
    popModalOut(overlay, qs(".status-card"), function () {
      overlay.classList.add("hidden");
    });
  }

  // ---- Build status corner card ----
  // Cleanup ref stops repeated builds stacking Cancel listeners.

  var _buildStatusFile = "";
  var _buildStatusCancelCleanup = null;

  function showBuildStatus(message, onCancel) {
    if (_buildStatusCancelCleanup) _buildStatusCancelCleanup();
    qs("#buildStatusSpinner").style.display = "";
    qs("#buildStatusMessage").textContent = message;
    // showBuildResult() rewrites className, so the shimmer clears itself.
    qs("#buildStatusMessage").className = "build-status-msg cg-shimmer";
    // Elapsed clock — idempotent start so a multi-message build keeps one clock.
    _buildEtaTracker.start();
    _studioEtaTicker.ensure();
    _paintBuildElapsed();
    qs("#buildStatusOpen").classList.add("hidden");
    qs("#buildStatusDismiss").classList.add("hidden");
    var cancelBtn = qs("#buildStatusCancel");
    var cancelLabel = cancelBtn.querySelector("span:last-child");
    if (onCancel) {
      cancelBtn.classList.remove("hidden");
      cancelBtn.disabled = false;
      cancelLabel.textContent = "Cancel";
      var handler = function () {
        cancelBtn.disabled = true;
        cancelLabel.textContent = "Cancelling…";
        onCancel();
      };
      cancelBtn.addEventListener("click", handler);
      _buildStatusCancelCleanup = function () {
        cancelBtn.removeEventListener("click", handler);
        _buildStatusCancelCleanup = null;
      };
    } else {
      cancelBtn.classList.add("hidden");
      _buildStatusCancelCleanup = null;
    }
    popModalIn(qs("#buildStatus"), qs(".build-status-card"));
  }

  function showBuildResult(successMsg, errorMsg, filePath) {
    if (_buildStatusCancelCleanup) _buildStatusCancelCleanup();
    _buildEtaTracker.reset();
    qs("#buildElapsed").textContent = "";
    qs("#buildStatusSpinner").style.display = "none";
    qs("#buildStatusCancel").classList.add("hidden");
    if (errorMsg) {
      qs("#buildStatusMessage").textContent = errorMsg;
      qs("#buildStatusMessage").className = "build-status-msg error-text";
      qs("#buildStatusOpen").classList.add("hidden");
    } else {
      qs("#buildStatusMessage").textContent = successMsg || "";
      qs("#buildStatusMessage").className = "build-status-msg";
      _buildStatusFile = filePath || "";
      qs("#buildStatusOpen").classList.toggle("hidden", !filePath);
    }
    qs("#buildStatusDismiss").classList.remove("hidden");
    popModalIn(qs("#buildStatus"), qs(".build-status-card"));
  }

  function hideBuildStatus() {
    if (_buildStatusCancelCleanup) _buildStatusCancelCleanup();
    _buildEtaTracker.reset();
    var buildEl = qs("#buildStatus");
    // Blank the clock with the hide, not before, so it doesn't vanish mid-fade.
    popModalOut(buildEl, qs(".build-status-card"), function () {
      qs("#buildElapsed").textContent = "";
      buildEl.classList.add("hidden");
    });
  }

  // ---- Confirm overlay ----

  var _confirmCleanup = null;

  function showConfirm(title, message, onYes, onNo) {
    if (_confirmCleanup) _confirmCleanup();
    qs("#confirmTitle").textContent = title;
    qs("#confirmMessage").textContent = message;
    var confirmEl = qs("#confirmOverlay");
    popModalIn(confirmEl, qs(".confirm-card"));

    var yesBtn = qs("#confirmYes");
    var noBtn = qs("#confirmNo");

    // Unwind synchronously so onYes() can open another overlay; only the hide trails.
    function cleanup() {
      closeModalTrap(confirmEl);
      yesBtn.removeEventListener("click", handleYes);
      noBtn.removeEventListener("click", handleNo);
      _confirmCleanup = null;
      popModalOut(confirmEl, qs(".confirm-card"), function () {
        confirmEl.classList.add("hidden");
      });
    }

    function handleYes() { cleanup(); onYes(); }
    function handleNo() { cleanup(); onNo(); }

    yesBtn.addEventListener("click", handleYes);
    noBtn.addEventListener("click", handleNo);
    _confirmCleanup = cleanup;
    // Escape cancels (same as No / backdrop click).
    openModalTrap(qs("#confirmOverlay"), handleNo);
  }

  function hideConfirm() {
    if (_confirmCleanup) {
      _confirmCleanup();
      return;
    }
    var overlay = qs("#confirmOverlay");
    popModalOut(overlay, qs(".confirm-card"), function () {
      overlay.classList.add("hidden");
    });
  }

  // ---- Artifact log ----

  // Same path as the dialogs; the veil is the shared .cg-modal-veil.
  function openLog() {
    var overlay = qs("#logOverlay");
    popModalIn(overlay, qs(".log-panel"));
    document.body.classList.add("modal-open");
    openModalTrap(overlay, closeLog);
    renderLog();
  }

  function closeLog() {
    var overlay = qs("#logOverlay");
    // Release the trap with the hide, not before; popModalOut's generation guard makes it safe.
    popModalOut(overlay, qs(".log-panel"), function () {
      closeModalTrap(overlay);
      overlay.classList.add("hidden");
      document.body.classList.remove("modal-open");
    });
  }

  // Desktop only: a native window has no other way to reach the file.
  function buildRevealButton(file) {
    var btn = el("button", "log-entry-reveal");
    btn.type = "button";
    btn.title = "Show on disk";
    btn.setAttribute("aria-label", "Show on disk");
    btn.innerHTML = iconHTML("folder-open");
    btn.addEventListener("click", function () {
      apiPost("api/reveal-artifact", { file: file }).catch(function (e) {
        showToast((e && e.message) || "Could not open the folder");
      });
    });
    return btn;
  }

  function renderLog() {
    var container = qs("#logContent");
    var countEl = qs("#logCount");

    var items = [];
    for (var ai = 0; ai < state.generatedArtifacts.length; ai++) items.push(state.generatedArtifacts[ai]);
    for (var ri = 0; ri < state.generatedReels.length; ri++) items.push(state.generatedReels[ri]);
    for (var vi = 0; vi < state.generatedViewers.length; vi++) items.push(state.generatedViewers[vi]);

    if (items.length === 0) {
      container.innerHTML = "";
      container.appendChild(el("div", "log-empty", "No artifacts, reels, or viewers generated yet."));
      countEl.textContent = "";
      return;
    }

    items.sort(function (a, b) { return (a._seq || 0) - (b._seq || 0); });

    container.innerHTML = "";
    var frag = document.createDocumentFragment();
    for (var k = items.length - 1; k >= 0; k--) {
      var a = items[k];
      var row = el("div", "log-entry");

      // Manifest reels carry no "type": detect via components or "reel:" id.
      var badgeType;
      if (a.type === "viewer") badgeType = "viewer";
      else if (a.components || (a.id || "").indexOf("reel:") === 0) badgeType = "reel";
      else badgeType = a.type || "clip";

      var badge = el("span", "log-type-badge", badgeType);
      badge.setAttribute("data-type", badgeType);
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
      if (state.desktop && a.file) row.appendChild(buildRevealButton(a.file));

      frag.appendChild(row);
    }
    container.appendChild(frag);

    var n = items.length;
    countEl.textContent = n + " item" + (n !== 1 ? "s" : "");
  }

  // ---- Settings (shared modal lives in settings-modal.js) ----

  function _findSetting(name) {
    if (!state.settingsData) return null;
    for (var i = 0; i < state.settingsData.length; i++) {
      if (state.settingsData[i].name === name) return state.settingsData[i];
    }
    return null;
  }

  // Option "(.ext)" labels follow the output-format settings; settingsData first, CLIPGEN_CONFIG fallback.
  function _formatExt(name, fallback) {
    var s = _findSetting(name);
    return s && s.value ? s.value : fallback;
  }

  function updateArtifactFormatLabels() {
    var sel = qs("#artifactFormat");
    if (!sel) return;
    var bases = { clip: "Clip", screen: "Screenshot", gif: "GIF" };
    var exts = {
      clip: _formatExt("FILEFORMAT", CLIPGEN_CONFIG.clipFormat),
      screen: _formatExt("SCREENSHOT_FORMAT", CLIPGEN_CONFIG.screenshotFormat),
      gif: _formatExt("GIF_FORMAT", CLIPGEN_CONFIG.gifFormat),
    };
    for (var i = 0; i < sel.options.length; i++) {
      var opt = sel.options[i];
      var base = bases[opt.value];
      if (base && exts[opt.value]) {
        opt.textContent = base + " (" + exts[opt.value] + ")";
      }
    }
  }

  function syncInlineControls() {
    updateArtifactFormatLabels();
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
    var scrubber = _findSetting("STUDIO_CARD_SCRUBBER");
    if (scrubber) {
      var wasOn = state.cardScrubberEnabled;
      state.cardScrubberEnabled = !!scrubber.value;
      if (wasOn !== state.cardScrubberEnabled) {
        // Tear down current attachments, then re-render so cards (re)wire (or
        // shed) their scrubbers via attachQueueScrubbers.
        if (window.clipgenCardScrubber) window.clipgenCardScrubber.detachAll();
        resetScrubberPrefetch();
        renderArtifactQueue();
        renderReelQueue();
        renderIntake(false);
      }
    }
    if (applyCrossRefSetting(null, state.settingsData)) rerenderCrossRefs();
  }

  // Also used by the palette; a hovered tooltip survives the rebuild, so hide it.
  function rerenderCrossRefs() {
    var tooltip = qs("#trIntakeTooltip");
    if (tooltip) tooltip.classList.add("hidden");
    renderIntake(false);
  }
  window.clipgenRerenderCrossRefs = rerenderCrossRefs;

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

  function checkNavLinks() {
    // Shared memoized fetch (utils.js): one /api/status per page load.
    clipgenStatus()
      .then(function (data) {
        state.desktop = !!data.desktop;
        if (state.desktop) renderLog();
        if (data.screenspace) {
          var intakeTab = qs('.preview-tab[data-tab="intake"]');
          if (intakeTab) intakeTab.classList.remove("hidden");
        }
        if (data.transcripts) {
          var trIntakeTab = qs('.preview-tab[data-tab="transcript-intake"]');
          if (trIntakeTab) trIntakeTab.classList.remove("hidden");
        }
        if (data.composer) {
          var coIntakeTab = qs('.preview-tab[data-tab="composer-intake"]');
          if (coIntakeTab) coIntakeTab.classList.remove("hidden");
        }
        // This tab only exists once a mind map has been opened.
        if (data.mindnode_loaded) {
          var mnIntakeTab = qs('.preview-tab[data-tab="mindnode-intake"]');
          if (mnIntakeTab) mnIntakeTab.classList.remove("hidden");
        }
        restoreStoredPreviewTab();
      })
      .catch(function () {});
  }

  // ---- Screenspace thumbnail queue (throttled + cached) ----

  var _ssThumbQueue = [];
  var _ssThumbActive = 0;
  var _SS_THUMB_MAX = 3;
  var _ssThumbCache = {}; // url -> objectURL | "error"

  // Revoke on pagehide or every lazily loaded thumb stays pinned (mirrors viewer.js).
  window.addEventListener("pagehide", function () {
    Object.keys(_ssThumbCache).forEach(function (key) {
      var url = _ssThumbCache[key];
      if (url && url !== "error" && url !== "loading") {
        try {
          URL.revokeObjectURL(url);
        } catch (_) {}
      }
      delete _ssThumbCache[key];
    });
  });

  function ssThumbUrl(participant, timestamp) {
    return "../screenspace/api/video/frame/" + encodeURIComponent(participant) + "/" + timestamp + "?w=200";
  }

  // Every card shares .queue-card-thumb for lazy source-frame loading.
  var SS_THUMB_SELECTOR = ".queue-card-thumb";

  function ssProcessQueue() {
    while (_ssThumbActive < _SS_THUMB_MAX && _ssThumbQueue.length) {
      var item = _ssThumbQueue.shift();
      if (!item.img.parentNode) continue;
      _ssThumbActive++;
      (function (entry) {
        apiGetBlob(entry.url)
          .then(function (blob) {
            var objUrl = URL.createObjectURL(blob);
            var prev = _ssThumbCache[entry.url];
            if (prev && prev !== "error" && prev !== "loading") {
              try { URL.revokeObjectURL(prev); } catch (_) {}
            }
            _ssThumbCache[entry.url] = objUrl;
            if (entry.img.parentNode) entry.img.src = objUrl;
          })
          .catch(function () {
            _ssThumbCache[entry.url] = "error";
            if (!entry.img.parentNode) return;
            // Custom onError (e.g. stash-folder icon) beats the queue-card fallback.
            if (entry.onError) { entry.onError(entry); return; }
            entry.img.remove();
            entry.thumbEl.appendChild(el("span", "", "\u2715"));
            entry.cardEl.classList.add("queue-card-error");
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

  // Enqueue with a custom error handler; shares ssEnqueueThumb's throttle and cache.
  function ssEnqueueThumbCustom(img, participant, timestamp, onError) {
    var url = ssThumbUrl(participant, timestamp);
    var cached = _ssThumbCache[url];
    if (cached && cached !== "error") { img.src = cached; return; }
    if (cached === "error") { if (onError) onError({ img: img }); return; }
    _ssThumbQueue.push({ img: img, url: url, onError: onError });
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

  // ---- Intake delegators — studio-intake.js ----
  // Guarded wrappers; bare hub call sites stay unchanged.
  function initIntake() { return STUDIO.initIntake && STUDIO.initIntake.apply(null, arguments); }
  function pollScreenspaceIntake() { return STUDIO.pollScreenspaceIntake && STUDIO.pollScreenspaceIntake.apply(null, arguments); }
  function pollTranscriptIntake() { return STUDIO.pollTranscriptIntake && STUDIO.pollTranscriptIntake.apply(null, arguments); }
  function pollComposerIntake() { return STUDIO.pollComposerIntake && STUDIO.pollComposerIntake.apply(null, arguments); }
  function pollMindnodeIntake() { return STUDIO.pollMindnodeIntake && STUDIO.pollMindnodeIntake.apply(null, arguments); }
  function refreshScreenspaceIntake() { return STUDIO.refreshScreenspaceIntake && STUDIO.refreshScreenspaceIntake.apply(null, arguments); }
  function refreshTranscriptIntake() { return STUDIO.refreshTranscriptIntake && STUDIO.refreshTranscriptIntake.apply(null, arguments); }
  function refreshComposerIntake() { return STUDIO.refreshComposerIntake && STUDIO.refreshComposerIntake.apply(null, arguments); }
  function refreshMindnodeIntake() { return STUDIO.refreshMindnodeIntake && STUDIO.refreshMindnodeIntake.apply(null, arguments); }
  function focusComposerIntakeItem() { return STUDIO.focusComposerIntakeItem && STUDIO.focusComposerIntakeItem.apply(null, arguments); }
  function refreshIntakeCardStates() { return STUDIO.refreshIntakeCardStates && STUDIO.refreshIntakeCardStates.apply(null, arguments); }
  function renderIntake() { return STUDIO.renderIntake && STUDIO.renderIntake.apply(null, arguments); }
  function _syncMarkCategoriesFromSettings() { return STUDIO._syncMarkCategoriesFromSettings && STUDIO._syncMarkCategoriesFromSettings.apply(null, arguments); }

  // buildXrefBadges / xrefBadgeIcon / XREF_ICON_BASE live in utils.js (shared with Overview).

  function initTopNavActions() {
    if (!window.ClipgenTopNav) return;
    function rebuild() {
      window.ClipgenTopNav.setQuickActions([
        {
          icon: "eye",
          label: "Build Viewer",
          action: onBuildViewer,
          title: "Build a self-contained timeline viewer HTML from the clips generated in this session",
        },
        {
          icon: "film",
          label: "Build Timeline",
          action: onBuildTimelineViewer,
          title: "Cut every clip in the sheet, then build and open a full timeline viewer HTML in the output folder",
        },
        {
          icon: "photo",
          label: "Build Gallery",
          action: openGalleryDialog,
          title: "Build a gallery viewer of screenshots or GIFs sampled from one participant's source video",
        },
        window.ClipgenExportActions.exportQuickAction(),
      ]);
    }
    rebuild();
    window.ClipgenExportActions.refreshExportStatus(rebuild);
    window.ClipgenTopNav.onBeforeOpen(function () {
      window.ClipgenExportActions.refreshExportStatus(rebuild);
    });
  }

  // Command palette additions beyond the auto-ingested quick actions.
  function initCommandPalette() {
    if (!window.ClipgenCommandPalette) return;
    window.ClipgenCommandPalette.setParticipants(function () {
      return (state.sheetData && state.sheetData.participants) || [];
    });
    function tabCommand(tabKey, title, icon) {
      function tabEl() {
        return document.querySelector('.preview-tab[data-tab="' + tabKey + '"]');
      }
      return {
        id: "studio:tab-" + tabKey,
        title: title,
        icon: icon,
        keywords: "tab show switch",
        section: "Studio",
        visible: function () {
          var tab = tabEl();
          return !!tab && !tab.classList.contains("hidden");
        },
        run: function () { tabEl().click(); },
      };
    }
    // Mirror the sidebar row's mutate → persist → re-render sequence.
    function applyFilterChange() {
      persistSidebarFilters();
      renderSidebar();
      renderGrid();
    }
    function viewCommand(viewId, title, icon) {
      return {
        id: "studio:view-" + viewId,
        title: title,
        icon: icon,
        keywords: "view filter severity rows sidebar",
        section: "Studio",
        visible: function () {
          return !!(state.sheetData && state.sheetData.rows && state.sheetData.rows.length);
        },
        run: function () { applySidebarView(viewId); applyFilterChange(); },
      };
    }
    window.ClipgenCommandPalette.register("studio", function () {
      return [
        {
          id: "studio:generate",
          title: "Generate clips",
          icon: "play",
          keywords: "render build artifacts",
          section: "Studio",
          enabled: function () {
            var btn = document.getElementById("generateBtn");
            return !!btn && !btn.disabled;
          },
          run: function () { document.getElementById("generateBtn").click(); },
        },
        tabCommand("sheet", "Show Sheet tab", "table-cells"),
        tabCommand("intake", "Show Screenspace Intake tab", "rectangle-stack"),
        tabCommand("transcript-intake", "Show Transcript Intake tab", "rectangle-stack"),
        tabCommand("composer-intake", "Show Composer Intake tab", "rectangle-stack"),
        {
          id: "studio:refresh",
          title: "Refresh current tab",
          icon: "arrow-path",
          keywords: "reload fetch update sheet intake",
          section: "Studio",
          visible: function () { return !!document.getElementById("studioRefresh"); },
          run: function () { document.getElementById("studioRefresh").click(); },
        },
        {
          id: "studio:clear-filters",
          title: "Clear all filters",
          icon: "x-mark",
          keywords: "reset remove sidebar category severity keyword",
          section: "Studio",
          enabled: hasActiveFilters,
          run: function () { clearAllFilters(); applyFilterChange(); },
        },
        viewCommand("all", "Show all rows", "bars-3"),
        viewCommand("highlights", "Highlights only", "funnel"),
        viewCommand("positive", "Positive only", "funnel"),
        {
          id: "studio:toggle-sidebar",
          title: "Toggle filters sidebar",
          icon: "adjustments-horizontal",
          keywords: "show hide panel drawer collapse",
          section: "Studio",
          visible: function () { return !!document.getElementById("studioSidebarToggle"); },
          run: function () { document.getElementById("studioSidebarToggle").click(); },
        },
        {
          id: "studio:artifact-log",
          title: "Open Artifact Log",
          icon: "list-bullet",
          keywords: "history builds",
          section: "Studio",
          visible: function () { return !!document.getElementById("logBtn"); },
          run: function () { document.getElementById("logBtn").click(); },
        },
      ];
    });
  }

  document.addEventListener("DOMContentLoaded", function () {
    // Mask static [data-icon] elements (the bottom-strip toolbar is static HTML).
    applyIconMasksIn(document);
    // #tab= deep links seed the stored tab so restoreStoredPreviewTab (and its retry) applies it.
    var hashTab = clipgenHashTab();
    if (hashTab) setStoredUIStateField("studio", "activeTab", hashTab);
    setActiveTabAttr(state.activePreviewTab);
    bindSidebarToggle();
    initThemeToggle();
    initPreviewTabs();
    initDropTargets();
    initWheelScroll();
    bindDragGate();
    bindReelReorder();
    bindQueueList(ARTIFACT_QUEUE);
    bindQueueList(REEL_QUEUE);
    bindButtons();
    updateArtifactActions();
    updateReelActions();
    loadStoredBottomHeight();
    initBottomPanelDivider();
    populateSheetSkeleton();
    loadSheetData();
    loadStashes();
    loadArtifactStashes();
    checkNavLinks();
    initFrontendSwitcher();
    initTopNavActions();
    initCommandPalette();
    initIntake();
    // Pollers back off 5s → 30s when quiet; handles live on state for wake().
    state.ssIntakePoller = createPoller(pollScreenspaceIntake, 5000, { maxIntervalMs: 30000, label: "studio.ssIntake" });
    state.trIntakePoller = createPoller(pollTranscriptIntake, 5000, { maxIntervalMs: 30000, label: "studio.trIntake" });
    state.coIntakePoller = createPoller(pollComposerIntake, 5000, { maxIntervalMs: 30000, label: "studio.coIntake" });
    state.mnIntakePoller = createPoller(pollMindnodeIntake, 5000, { maxIntervalMs: 30000, label: "studio.mnIntake" });
    state.ssIntakePoller.start();
    state.trIntakePoller.start();
    state.coIntakePoller.start();
    state.mnIntakePoller.start();
    // One-shot job-status fetch re-attaches to a background build; it self-schedules if busy.
    pollJobStatus();
    document.addEventListener("visibilitychange", function () {
      if (document.hidden) {
        stopJobStatusPoll();
        _studioEtaTicker.stop();
      } else {
        pollJobStatus();
        if (isAnyStudioJobRunning()) _studioEtaTicker.ensure();
      }
    });
    document.addEventListener("dragstart", function (ev) {
      if (!ev.target.closest(".stash-card")) return;
      requestAnimationFrame(revealEmptyStashAreas);
    });
    document.addEventListener("dragend", function () {
      hideEmptyStashAreas();
    });
  });

  // Hub → studio-intake.js. utils.js globals reach it via scope; only hub-locals need publishing.
  STUDIO.state = state;
  STUDIO.buildQueueCardThumb = buildQueueCardThumb;
  STUDIO.buildXrefBadges = buildXrefBadges;
  STUDIO.findIntakeInQueue = findIntakeInQueue;
  STUDIO.findOverlappingData = findOverlappingData;
  STUDIO.intakeAddItem = intakeAddItem;
  STUDIO.intakeAddItems = intakeAddItems;
  STUDIO.intakeToggleItem = intakeToggleItem;
  STUDIO.isIntakeSource = isIntakeSource;
  STUDIO.isArtifactQueueLocked = isArtifactQueueLocked;
  STUDIO.isReelQueueLocked = isReelQueueLocked;
  STUDIO.cellKey = cellKey;
  STUDIO.updateSingleCellClass = updateSingleCellClass;
  STUDIO.ssEnqueueThumbCustom = ssEnqueueThumbCustom;
  STUDIO.renderArtifactQueue = renderArtifactQueue;
  STUDIO.renderReelQueue = renderReelQueue;
  STUDIO.saveQueues = saveQueues;
  STUDIO.setCardDragImage = setCardDragImage;
  STUDIO.ssClearPending = ssClearPending;
  STUDIO.kbPaintCursor = kbPaintCursor;

  // Hub → studio-generate.js: card painters, status/result helpers, elapsed trackers.
  STUDIO.setArtifactGenerating = setArtifactGenerating;
  STUDIO.showResult = showResult;
  STUDIO.revealStatusOverlay = revealStatusOverlay;
  STUDIO.setCardQueued = setCardQueued;
  STUDIO.clearCardStatus = clearCardStatus;
  STUDIO.setCardResult = setCardResult;
  STUDIO.updateGenerateProgress = updateGenerateProgress;
  STUDIO.stampLog = stampLog;
  STUDIO._generateEtaTracker = _generateEtaTracker;
  STUDIO._studioEtaTicker = _studioEtaTicker;
})();
