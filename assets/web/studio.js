/* clipgen Studio */

(function () {
  "use strict";

  var QUEUE_STORAGE_KEY = "clipgen-studio-queues";

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
    dividerOffset: 0,
    bottomCollapsed: false,
    dividerOffsetBeforeCollapse: 0,
    activeFunction: "",
    cellExpandHover: true,
    cellColorCoding: true,
    filtersVisible: false,
    filters: {
      categories: [],
      sevMin: "",
      sevMax: "",
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
    trIntakeFilterCategory: "",
    trIntakeFilterParticipants: [],
    trIntakeFilterText: "",
    trIntakeShowAll: false,
    trIntakeHoveredIdx: -1,
    trIntakeTooltipsEnabled: true,
    convergenceBaselines: {},
    convergenceDataVersion: 0,
    convergenceStale: false,
  };

  function isIntakeSource(source) {
    return source === "screenspace" || source === "transcript";
  }

  var SEVERITY_ORDER = [
    { label: "Critical", rank: -4 },
    { label: "High", rank: -3 },
    { label: "Medium", rank: -2 },
    { label: "Low", rank: -1 },
    { label: "N/A", rank: 0 },
    { label: "Positive", rank: 1 },
    { label: "Very Positive", rank: 2 },
  ];

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
        segStart: segments[i].startSeconds,
        segDuration: segments[i].duration,
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
    var DEFAULT_DUR = (state.sheetData && state.sheetData.defaultDuration) || 60;
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

  function formatDuration(secs) {
    secs = Math.round(secs);
    if (secs >= 3600) {
      var h = Math.floor(secs / 3600);
      var m = Math.floor((secs % 3600) / 60);
      var s = secs % 60;
      return h + ":" + (m < 10 ? "0" : "") + m + ":" + (s < 10 ? "0" : "") + s;
    }
    var m2 = Math.floor(secs / 60);
    var s2 = secs % 60;
    return m2 + ":" + (s2 < 10 ? "0" : "") + s2;
  }

  // Cross-referencing: find overlapping data from other sources for a given
  // participant + time range. Used by both Screenspace and Transcript intake
  // card renderers to surface context from sibling data sources.
  function findOverlappingData(participant, start, end) {
    var result = { transcriptSnippets: [], screenspaceEvents: [], sheetObservations: [] };

    // Transcript marks/clusters
    for (var i = 0; i < state.trIntakeClusters.length; i++) {
      var tc = state.trIntakeClusters[i];
      if (tc.participant === participant && tc.start < end && tc.end > start) {
        result.transcriptSnippets.push({ text: tc.text || tc.label || "", category: tc.category, start: tc.start, end: tc.end });
      }
    }

    // Screenspace event clusters
    for (var j = 0; j < state.intakeClusters.length; j++) {
      var sc = state.intakeClusters[j];
      if (sc.participant === participant && sc.start < end && sc.end > start) {
        result.screenspaceEvents.push({ detector: sc.detector, event_type: sc.event_type, start: sc.start, end: sc.end });
      }
    }

    // Sheet observations
    if (state.sheetData && state.sheetData.rows) {
      for (var k = 0; k < state.sheetData.rows.length; k++) {
        var row = state.sheetData.rows[k];
        var cell = row.cells[participant];
        if (!cell || !cell.valid) continue;
        var segs = parseClipTimestamps(cell.value, participant);
        for (var s = 0; s < segs.length; s++) {
          var segEnd = segs[s].startSeconds + segs[s].duration;
          if (segs[s].startSeconds < end && segEnd > start) {
            result.sheetObservations.push({ observation: row.observation, category: row.category, severity: row.severity });
            break;
          }
        }
      }
    }

    return result;
  }

  function intakeComputeTickInterval(visLen) {
    var candidates = [1, 2, 5, 10, 15, 30, 60, 120, 300, 600, 900, 1800, 3600];
    for (var i = 0; i < candidates.length; i++) {
      if (visLen / candidates[i] <= 12) return candidates[i];
    }
    return 3600;
  }

  function setCardDragImage(ev, card) {
    var clone = card.cloneNode(true);
    clone.style.position = "absolute";
    clone.style.top = "-9999px";
    clone.style.left = "-9999px";
    clone.style.width = card.offsetWidth + "px";
    clone.style.zIndex = "-1";
    document.body.appendChild(clone);
    var rect = card.getBoundingClientRect();
    ev.dataTransfer.setDragImage(clone, ev.clientX - rect.left, ev.clientY - rect.top);
    setTimeout(function () { document.body.removeChild(clone); }, 0);
  }

  // ---- Filtering ----

  function severityRank(label) {
    if (!label) return null;
    var k = label.trim().toLowerCase();
    for (var i = 0; i < SEVERITY_ORDER.length; i++) {
      if (SEVERITY_ORDER[i].label.toLowerCase() === k) return SEVERITY_ORDER[i].rank;
    }
    return null;
  }

  function hasActiveFilters() {
    var f = state.filters;
    return (
      f.categories.length > 0 ||
      f.sevMin !== "" ||
      f.sevMax !== "" ||
      f.fnMin !== null ||
      f.fnMax !== null
    );
  }

  function getFilteredRows(rows) {
    var f = state.filters;
    if (!hasActiveFilters()) return rows;

    var sevMinRank = f.sevMin ? severityRank(f.sevMin) : null;
    var sevMaxRank = f.sevMax ? severityRank(f.sevMax) : null;
    var fnActive = state.activeFunction && ROW_FUNCTIONS[state.activeFunction];
    var participants = state.sheetData ? state.sheetData.participants : [];

    return rows.filter(function (row) {
      if (f.categories.length > 0) {
        if (!row.category || f.categories.indexOf(row.category) < 0) return false;
      }
      if (sevMinRank !== null || sevMaxRank !== null) {
        var r = severityRank(row.severity);
        if (r === null) return false;
        if (sevMinRank !== null && r < sevMinRank) return false;
        if (sevMaxRank !== null && r > sevMaxRank) return false;
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
    state.filters.sevMin = "";
    state.filters.sevMax = "";
    state.filters.fnMin = null;
    state.filters.fnMax = null;
  }

  function renderFilterBar() {
    var bar = qs("#filterBar");
    if (!bar || !state.sheetData) return;
    bar.innerHTML = "";

    var d = state.sheetData;

    // --- Category filter ---
    var uniqueCats = [];
    for (var i = 0; i < d.rows.length; i++) {
      var cat = d.rows[i].category;
      if (cat && uniqueCats.indexOf(cat) < 0) uniqueCats.push(cat);
    }
    uniqueCats.sort();

    if (uniqueCats.length > 0) {
      var catGroup = el("div", "filter-group");
      catGroup.appendChild(el("span", "filter-group-label", "Category"));

      var catWrap = el("div", "filter-cat-wrap");
      var catBtn = el("button", "filter-cat-btn");
      catBtn.type = "button";
      catBtn.innerHTML = 'All <span class="chevron">\u25BE</span>';
      var catPanel = el("div", "filter-cat-panel hidden");

      for (var ci = 0; ci < uniqueCats.length; ci++) {
        var lbl = document.createElement("label");
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.value = uniqueCats[ci];
        cb.setAttribute("data-filter-cat", uniqueCats[ci]);
        lbl.appendChild(cb);
        lbl.appendChild(document.createTextNode(uniqueCats[ci]));
        catPanel.appendChild(lbl);
      }

      catBtn.addEventListener("click", function (ev) {
        ev.stopPropagation();
        catBtn.classList.toggle("open");
        catPanel.classList.toggle("hidden");
      });

      catPanel.addEventListener("change", function () {
        var checked = catPanel.querySelectorAll("input:checked");
        state.filters.categories = [];
        for (var j = 0; j < checked.length; j++) {
          state.filters.categories.push(checked[j].value);
        }
        var count = state.filters.categories.length;
        catBtn.innerHTML =
          (count === 0 ? "All" : count + " selected") +
          ' <span class="chevron">\u25BE</span>';
        applyGridFilters();
        computeGridMaxHeight();
      });

      catWrap.appendChild(catBtn);
      catWrap.appendChild(catPanel);
      catGroup.appendChild(catWrap);

      var catClear = document.createElement("button");
      catClear.className = "filter-clear";
      catClear.type = "button";
      catClear.title = "Clear category filter";
      catClear.innerHTML =
        '<svg viewBox="0 0 16 16" fill="currentColor"><path d="M5.28 4.22a.75.75 0 0 0-1.06 1.06L6.94 8l-2.72 2.72a.75.75 0 1 0 1.06 1.06L8 9.06l2.72 2.72a.75.75 0 1 0 1.06-1.06L9.06 8l2.72-2.72a.75.75 0 0 0-1.06-1.06L8 6.94 5.28 4.22Z"/></svg>';
      catClear.addEventListener("click", function () {
        state.filters.categories = [];
        var cbs = catPanel.querySelectorAll("input[type=checkbox]");
        for (var j = 0; j < cbs.length; j++) cbs[j].checked = false;
        catBtn.innerHTML = 'All <span class="chevron">\u25BE</span>';
        applyGridFilters();
        computeGridMaxHeight();
      });
      catGroup.appendChild(catClear);

      bar.appendChild(catGroup);
    }

    // --- Severity filter ---
    var showSeverity = hasSeverityData(d.rows);
    if (showSeverity) {
      var sevGroup = el("div", "filter-group");
      sevGroup.appendChild(el("span", "filter-group-label", "Severity"));

      var sevMin = document.createElement("select");
      sevMin.className = "filter-select";
      sevMin.id = "filterSevMin";
      var sevMinDefault = document.createElement("option");
      sevMinDefault.value = "";
      sevMinDefault.textContent = "Any";
      sevMin.appendChild(sevMinDefault);
      for (var si = 0; si < SEVERITY_ORDER.length; si++) {
        var opt = document.createElement("option");
        opt.value = SEVERITY_ORDER[si].label;
        opt.textContent = SEVERITY_ORDER[si].label;
        sevMin.appendChild(opt);
      }

      var sevMax = document.createElement("select");
      sevMax.className = "filter-select";
      sevMax.id = "filterSevMax";
      var sevMaxDefault = document.createElement("option");
      sevMaxDefault.value = "";
      sevMaxDefault.textContent = "Any";
      sevMax.appendChild(sevMaxDefault);
      for (var sj = 0; sj < SEVERITY_ORDER.length; sj++) {
        var opt2 = document.createElement("option");
        opt2.value = SEVERITY_ORDER[sj].label;
        opt2.textContent = SEVERITY_ORDER[sj].label;
        sevMax.appendChild(opt2);
      }

      function onSevChange() {
        state.filters.sevMin = sevMin.value;
        state.filters.sevMax = sevMax.value;
        applyGridFilters();
        computeGridMaxHeight();
      }
      sevMin.addEventListener("change", onSevChange);
      sevMax.addEventListener("change", onSevChange);

      sevGroup.appendChild(sevMin);
      sevGroup.appendChild(el("span", "filter-range-sep", "to"));
      sevGroup.appendChild(sevMax);

      var sevClear = document.createElement("button");
      sevClear.className = "filter-clear";
      sevClear.type = "button";
      sevClear.title = "Clear severity filter";
      sevClear.innerHTML =
        '<svg viewBox="0 0 16 16" fill="currentColor"><path d="M5.28 4.22a.75.75 0 0 0-1.06 1.06L6.94 8l-2.72 2.72a.75.75 0 1 0 1.06 1.06L8 9.06l2.72 2.72a.75.75 0 1 0 1.06-1.06L9.06 8l2.72-2.72a.75.75 0 0 0-1.06-1.06L8 6.94 5.28 4.22Z"/></svg>';
      sevClear.addEventListener("click", function () {
        state.filters.sevMin = "";
        state.filters.sevMax = "";
        sevMin.value = "";
        sevMax.value = "";
        applyGridFilters();
        computeGridMaxHeight();
      });
      sevGroup.appendChild(sevClear);

      bar.appendChild(sevGroup);
    }

    // --- Function filter ---
    var fnGroup = el("div", "filter-group");
    fnGroup.appendChild(el("span", "filter-group-label", "Function"));

    var fnMin = document.createElement("input");
    fnMin.type = "number";
    fnMin.className = "filter-number";
    fnMin.id = "filterFnMin";
    fnMin.placeholder = "Min";
    fnMin.disabled = !state.activeFunction;

    var fnMax = document.createElement("input");
    fnMax.type = "number";
    fnMax.className = "filter-number";
    fnMax.id = "filterFnMax";
    fnMax.placeholder = "Max";
    fnMax.disabled = !state.activeFunction;

    function onFnChange() {
      var minVal = fnMin.value.trim();
      var maxVal = fnMax.value.trim();
      state.filters.fnMin = minVal !== "" ? parseFloat(minVal) : null;
      state.filters.fnMax = maxVal !== "" ? parseFloat(maxVal) : null;
      applyGridFilters();
      computeGridMaxHeight();
    }
    fnMin.addEventListener("input", onFnChange);
    fnMax.addEventListener("input", onFnChange);

    fnGroup.appendChild(fnMin);
    fnGroup.appendChild(el("span", "filter-range-sep", "to"));
    fnGroup.appendChild(fnMax);

    var fnClear = document.createElement("button");
    fnClear.className = "filter-clear";
    fnClear.type = "button";
    fnClear.title = "Clear function filter";
    fnClear.innerHTML =
      '<svg viewBox="0 0 16 16" fill="currentColor"><path d="M5.28 4.22a.75.75 0 0 0-1.06 1.06L6.94 8l-2.72 2.72a.75.75 0 1 0 1.06 1.06L8 9.06l2.72 2.72a.75.75 0 1 0 1.06-1.06L9.06 8l2.72-2.72a.75.75 0 0 0-1.06-1.06L8 6.94 5.28 4.22Z"/></svg>';
    fnClear.addEventListener("click", function () {
      state.filters.fnMin = null;
      state.filters.fnMax = null;
      fnMin.value = "";
      fnMax.value = "";
      applyGridFilters();
      computeGridMaxHeight();
    });
    fnGroup.appendChild(fnClear);

    bar.appendChild(fnGroup);
  }

  function initFilterToggle() {
    var btn = qs("#filterToggle");
    if (!btn) return;
    btn.addEventListener("click", function () {
      state.filtersVisible = !state.filtersVisible;
      var bar = qs("#filterBar");
      if (state.filtersVisible) {
        bar.classList.remove("hidden");
      } else {
        bar.classList.add("hidden");
        if (hasActiveFilters()) {
          clearAllFilters();
          applyGridFilters();
        }
      }
      computeGridMaxHeight();
    });
  }

  function syncFilterFnDisabled() {
    var fnMin = qs("#filterFnMin");
    var fnMax = qs("#filterFnMax");
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
        var allTabs = qsa(".preview-tab");
        for (var j = 0; j < allTabs.length; j++) allTabs[j].classList.remove("active");
        this.classList.add("active");
        syncPreviewTab();
      });
    }
  }

  function syncPreviewTab() {
    var grid = qs("#sheetGrid");
    var filterBar = qs("#filterBar");
    var filterToggle = qs("#filterToggle");
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
    if (filterBar) filterBar.classList.add("hidden");
    if (filterToggle) filterToggle.classList.add("hidden");
    if (refreshBtn) refreshBtn.classList.add("hidden");

    // Stop both intake poll timers
    if (state.intakePollTimer) { clearInterval(state.intakePollTimer); state.intakePollTimer = null; }
    if (state.trIntakePollTimer) { clearInterval(state.trIntakePollTimer); state.trIntakePollTimer = null; }
    if (window.convergenceDeactivate) window.convergenceDeactivate();
    if (window.metadataDeactivate) window.metadataDeactivate();

    if (state.activePreviewTab === "sheet") {
      grid.classList.remove("hidden");
      if (filterToggle) filterToggle.classList.remove("hidden");
      if (refreshBtn) refreshBtn.classList.remove("hidden");
      if (state.filtersVisible && filterBar) filterBar.classList.remove("hidden");
    } else if (state.activePreviewTab === "intake") {
      intakePanel.classList.remove("hidden");
      if (!document.hidden) {
        pollIntakeEvents();
        state.intakePollTimer = setInterval(pollIntakeEvents, 10000);
      }
      setTimeout(sizeIntakeCanvas, 0);
    } else if (state.activePreviewTab === "transcript-intake") {
      if (trIntakePanel) trIntakePanel.classList.remove("hidden");
      if (!document.hidden) {
        pollTranscriptIntakeMarks();
        state.trIntakePollTimer = setInterval(pollTranscriptIntakeMarks, 10000);
      }
      setTimeout(sizeTrIntakeCanvas, 0);
    } else if (state.activePreviewTab === "convergence") {
      if (convergencePanel) convergencePanel.classList.remove("hidden");
      if (window.convergenceActivate) window.convergenceActivate();
    } else if (state.activePreviewTab === "metadata") {
      if (metadataPanel) metadataPanel.classList.remove("hidden");
      if (window.metadataActivate) window.metadataActivate();
    }
    computeGridMaxHeight();
  }

  // ---- Queue persistence (sessionStorage) ----

  function saveQueues() {
    try {
      sessionStorage.setItem(QUEUE_STORAGE_KEY, JSON.stringify({
        artifactQueue: state.artifactQueue,
        reelQueue: state.reelQueue,
      }));
    } catch (e) { /* ignore quota errors */ }
  }

  function restoreQueues() {
    try {
      var raw = sessionStorage.getItem(QUEUE_STORAGE_KEY);
      if (!raw) return;
      var saved = JSON.parse(raw);
      if (saved.artifactQueue) state.artifactQueue = saved.artifactQueue;
      if (saved.reelQueue) state.reelQueue = saved.reelQueue;
    } catch (e) { /* ignore parse errors */ }
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
        renderHeader();
        renderFilterBar();
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
        computeGridMaxHeight();
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
        if (data.cellColorCoding !== undefined) {
          state.cellColorCoding = data.cellColorCoding;
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

    // TODO: no r.ok check; migrating to apiPost would add throw on HTTP error.
    fetch("api/sheet/refresh", { method: "POST" })
      .then(function (r) { return r.json(); })
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

  function loadManifestState() {
    // TODO: no r.ok check; migrating to apiGet would add throw on HTTP error.
    fetch("api/manifest")
      .then(function (r) { return r.json(); })
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
    qs("#versionInfo").textContent =
      "v" + (d.version || "") + " \u00B7 " + d.participants.length + " participants";
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
    var totalCols = metaCols + d.participants.length;
    state._gridTotalCols = totalCols;
    state._gridShowSeverity = showSeverity;
    var table = el("table");

    // Colgroup for fixed column widths — participant columns share equal width
    var colgroup = document.createElement("colgroup");
    var colRowNum = document.createElement("col");
    colRowNum.style.width = "3rem";
    colgroup.appendChild(colRowNum);
    var colFn = document.createElement("col");
    colFn.style.width = "3.5rem";
    colgroup.appendChild(colFn);
    var colObs = document.createElement("col");
    colObs.style.width = "auto";
    colgroup.appendChild(colObs);
    var colCat = document.createElement("col");
    colCat.style.width = "7rem";
    colgroup.appendChild(colCat);
    if (showSeverity) {
      var colSev = document.createElement("col");
      colSev.style.width = "5.5rem";
      colgroup.appendChild(colSev);
    }
    for (var c = 0; c < d.participants.length; c++) {
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

    for (var p = 0; p < d.participants.length; p++) {
      var pTh = el("th", "col-participant", d.participants[p]);
      pTh.setAttribute("data-participant", d.participants[p]);
      pTh.title = "Select all " + d.participants[p] + " cells";
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

    var filteredRows = getFilteredRows(d.rows);
    var frag = document.createDocumentFragment();
    var i = 0;
    while (i < filteredRows.length) {
      var row = filteredRows[i];
      if (isRowEmpty(row, d.participants)) {
        var emptyStart = i;
        while (i < filteredRows.length && isRowEmpty(filteredRows[i], d.participants)) {
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
        frag.appendChild(renderDataRow(row, d.participants, showSeverity));
        i++;
      }
    }
    tbody.innerHTML = "";
    tbody.appendChild(frag);
    updateCellClasses();
    if (state.activeFunction) updateFunctionColumn();
  }

  // ---- Panel divider (resizable split between sheet preview and bottom panel) ----

  // Tab panels whose height tracks the divider. Keep in sync with the markup in studio.html;
  // every preview tab that shares the upper pane must appear here so the drag/collapse/resize
  // code paths apply the same maxHeight to each.
  var UPPER_PANE_PANELS = [
    "#sheetGrid",
    "#intakePanel",
    "#trIntakePanel",
    "#convergencePanel",
    "#metadataPanel",
  ];

  function applyUpperPaneMaxHeight(value) {
    for (var i = 0; i < UPPER_PANE_PANELS.length; i++) {
      var p = qs(UPPER_PANE_PANELS[i]);
      if (p) p.style.maxHeight = value;
    }
  }

  function computeGridMaxHeight(bottomHeightOverride) {
    var header = qs("#studioHeader");
    var preview = qs("#sheetPreview");
    var divider = qs("#panelDivider");
    var bottom = qs("#bottomPanel");
    var grid = qs("#sheetGrid");
    if (!header || !preview || !divider || !bottom || !grid) return;

    var previewHeader = preview.querySelector(".sheet-preview-header");
    var filterBar = qs("#filterBar");
    var previewStyle = getComputedStyle(preview);
    var previewPadTop = parseFloat(previewStyle.paddingTop) || 0;
    var previewPadBot = parseFloat(previewStyle.paddingBottom) || 0;
    var phHeight = previewHeader ? previewHeader.offsetHeight : 0;
    var phStyle = previewHeader ? getComputedStyle(previewHeader) : null;
    var phMargin = phStyle
      ? (parseFloat(phStyle.marginTop) || 0) + (parseFloat(phStyle.marginBottom) || 0)
      : 0;
    var fbHeight = filterBar && !filterBar.classList.contains("hidden") ? filterBar.offsetHeight : 0;
    var fbStyle = filterBar && !filterBar.classList.contains("hidden") ? getComputedStyle(filterBar) : null;
    var fbMargin = fbStyle
      ? (parseFloat(fbStyle.marginTop) || 0) + (parseFloat(fbStyle.marginBottom) || 0)
      : 0;

    var headerRect = header.getBoundingClientRect();
    var sheetChrome = previewPadTop + phHeight + phMargin + fbHeight + fbMargin + previewPadBot;
    var bottomH = bottomHeightOverride !== undefined ? bottomHeightOverride : bottom.offsetHeight;
    var available = window.innerHeight - headerRect.top - headerRect.height
      - sheetChrome - divider.offsetHeight - bottomH;

    var MIN_GRID = 100;
    var maxAllowed = Math.max(0, available - MIN_GRID);
    state.dividerOffset = Math.min(state.dividerOffset, maxAllowed);

    var maxH = Math.max(MIN_GRID, available - state.dividerOffset) + "px";
    applyUpperPaneMaxHeight(maxH);
  }

  function initPanelDivider() {
    var handle = qs("#panelDivider");
    if (!handle) return;
    var dragging = false;
    var startY = 0;
    var startOffset = 0;
    var dragAvailable = 0; // stable available-space snapshot for the drag
    var dragMaxOff = 0;    // stable upper bound for dividerOffset

    function onDown(e) {
      if (state.bottomCollapsed) return;
      e.preventDefault();
      dragging = true;
      startY = e.clientY || (e.touches && e.touches[0].clientY) || 0;
      startOffset = state.dividerOffset;

      // Snapshot layout values once so they stay stable for the whole drag.
      // Re-reading bottom.offsetHeight each frame causes oscillation when
      // the upper panel shrinks and the bottom panel's visible area grows.
      var header = qs("#studioHeader");
      var preview = qs("#sheetPreview");
      var divider = qs("#panelDivider");
      var bottom = qs("#bottomPanel");
      if (header && preview && divider && bottom) {
        var previewHeader = preview.querySelector(".sheet-preview-header");
        var filterBar = qs("#filterBar");
        var previewStyle = getComputedStyle(preview);
        var previewPadTop = parseFloat(previewStyle.paddingTop) || 0;
        var previewPadBot = parseFloat(previewStyle.paddingBottom) || 0;
        var phHeight = previewHeader ? previewHeader.offsetHeight : 0;
        var phStyle = previewHeader ? getComputedStyle(previewHeader) : null;
        var phMargin = phStyle
          ? (parseFloat(phStyle.marginTop) || 0) + (parseFloat(phStyle.marginBottom) || 0)
          : 0;
        var fbHeight = filterBar && !filterBar.classList.contains("hidden") ? filterBar.offsetHeight : 0;
        var fbStyle = filterBar && !filterBar.classList.contains("hidden") ? getComputedStyle(filterBar) : null;
        var fbMargin = fbStyle
          ? (parseFloat(fbStyle.marginTop) || 0) + (parseFloat(fbStyle.marginBottom) || 0)
          : 0;
        var headerRect = header.getBoundingClientRect();
        var sheetChrome = previewPadTop + phHeight + phMargin + fbHeight + fbMargin + previewPadBot;
        dragAvailable = window.innerHeight - headerRect.top - headerRect.height
          - sheetChrome - divider.offsetHeight - bottom.offsetHeight;
      }

      var MIN_GRID = 100;
      dragMaxOff = Math.max(0, dragAvailable - MIN_GRID);

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
        var delta = startY - clientY;
        state.dividerOffset = Math.max(0, Math.min(dragMaxOff, startOffset + delta));

        // Apply maxHeight directly using the stable snapshot
        var MIN_GRID = 100;
        var maxH = Math.max(MIN_GRID, dragAvailable - state.dividerOffset) + "px";
        applyUpperPaneMaxHeight(maxH);

        rafPending = false;
      });
    }

    function onUp() {
      if (!dragging) return;
      dragging = false;
      handle.classList.remove("active");
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      computeGridMaxHeight(); // finalize with fresh layout values
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
    var divider = qs("#panelDivider");

    if (state.bottomCollapsed) {
      // --- Restore ---
      state.bottomCollapsed = false;
      state.dividerOffset = 0;

      // Measure target bottom height (temporarily lift the inline constraint)
      bottom.style.transition = "none";
      bottom.style.maxHeight = "none";
      var targetH = bottom.offsetHeight;

      // Transition divider margin from 20px → 0 alongside the bottom panel grow
      if (divider) divider.style.marginBottom = "0";

      document.body.classList.add("bottom-animating");

      // Keep bottom-collapsed during animation so flex layout smoothly adjusts
      bottom.style.maxHeight = "0px";
      bottom.offsetHeight; // reflow — margin transition starts, bottom at 0
      bottom.style.transition = "";
      bottom.style.maxHeight = targetH + "px";

      onCollapseTransitionEnd(bottom, function () {
        document.body.classList.remove("bottom-collapsed");
        document.body.classList.remove("bottom-animating");
        if (divider) divider.style.marginBottom = "";
        bottom.style.maxHeight = "";
        bottom._transitioning = false;
        computeGridMaxHeight();
      });
    } else {
      // --- Collapse ---
      state.bottomCollapsed = true;
      state.dividerOffsetBeforeCollapse = state.dividerOffset;
      state.dividerOffset = 0;

      // Clear grid maxHeight — collapsed CSS flex rules fill the space instead
      applyUpperPaneMaxHeight("");

      document.body.classList.add("bottom-animating");
      var currentH = bottom.offsetHeight;
      bottom.style.maxHeight = currentH + "px";
      document.body.classList.add("bottom-collapsed");
      bottom.offsetHeight; // reflow
      bottom.style.maxHeight = "0px";

      onCollapseTransitionEnd(bottom, function () {
        bottom._transitioning = false;
        document.body.classList.remove("bottom-animating");
        computeGridMaxHeight(0);
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
    obsTd.textContent = truncate(row.observation, 50);
    obsTd.title = row.observation;
    tr.appendChild(obsTd);
    tr.appendChild(el("td", "col-category", row.category || ""));
    if (showSeverity) {
      var sevCls = "col-severity";
      if (row.severity) sevCls += " " + severityClass(row.severity);
      tr.appendChild(el("td", sevCls, row.severity || ""));
    }

    for (var j = 0; j < participants.length; j++) {
      var pid = participants[j];
      var cellData = row.cells[pid] || {};
      var td = el("td", "ts-cell");
      td.setAttribute("data-row", row.rowNum);
      td.setAttribute("data-participant", pid);
      td.setAttribute("data-observation", row.observation || "");
      td.setAttribute("data-category", row.category || "");

      if (cellData.hasText) {
        td.textContent = cellData.value;
        if (cellData.valid) {
          td.classList.add("valid-ts");
          if (state.cellColorCoding) {
            if (row.severity) {
              var sevCellCls = severityClass(row.severity);
              if (sevCellCls) td.classList.add(sevCellCls);
            }
            var segs = parseClipTimestamps(cellData.value, pid);
            var totalDur = 0;
            for (var k = 0; k < segs.length; k++) totalDur += segs[k].duration;
            var intensity;
            if (totalDur < 15) intensity = 0.55;
            else if (totalDur < 45) intensity = 0.7;
            else if (totalDur < 90) intensity = 0.85;
            else intensity = 1;
            td.style.setProperty("--ts-intensity", intensity);
          }
        } else {
          td.classList.add("has-text");
        }
        td.setAttribute("draggable", "true");
      } else {
        td.classList.add("empty");
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
    var hm = getComputedStyle(document.documentElement).getPropertyValue("--color-heatmap").trim() || "168, 130, 214";
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
      if (td.scrollWidth <= td.clientWidth) return;
      floatCell = td;
      var rect = td.getBoundingClientRect();
      cellFloat.textContent = td.textContent;
      cellFloat.style.backgroundColor = getComputedStyle(td).backgroundColor;
      cellFloat.style.top = (rect.top + 1) + "px";
      cellFloat.style.left = rect.left + "px";
      cellFloat.style.height = (rect.height - 1) + "px";
      cellFloat.style.display = "block";
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

  function bindDragFromGrid() {
    var grid = qs("#sheetGrid");

    grid.addEventListener("dragstart", function (ev) {
      var td = ev.target.closest(".ts-cell");
      if (!td || td.classList.contains("empty")) return;
      var info = getCellInfo(td);
      ev.dataTransfer.setData("application/json", JSON.stringify(info));
      ev.dataTransfer.effectAllowed = "copy";
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
      target.classList.add("drag-over");
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
          segStart: reelItem.segStart,
          segDuration: reelItem.segDuration,
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
        el("div", "drop-target-empty", "Click or drag cells here to queue for generation")
      );
      return;
    }

    for (var i = 0; i < n; i++) {
      var item = state.artifactQueue[i];
      var isIntake = isIntakeSource(item.source);
      var segStart, segDuration;
      if (item.segStart !== undefined && item.segDuration !== undefined) {
        segStart = item.segStart;
        segDuration = item.segDuration;
      } else {
        var parsed = parseClipTimestamps(item.timestamp, item.participant)[0];
        segStart = parsed.startSeconds;
        segDuration = parsed.duration;
      }
      var segTotal = item.segTotal || 1;
      var segIdx = item.segIdx || 0;

      var card = el("div", "queue-card" + (isIntake ? " queue-card-intake" : ""));
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
            segStart: itm.segStart,
            segDuration: itm.segDuration,
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
        ssObserveThumb(card, img, thumb, item.participant, segStart);
      } else {
        img.src = "api/thumbnail/" + encodeURIComponent(item.participant) + "/" + segStart;
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
          ssBadge.innerHTML = '<svg viewBox="0 0 16 16" fill="currentColor"><path d="M3 3h10v1.5H3zM3 7h8v1.5H3zM3 11h10v1.5H3z"/></svg>';
        } else {
          ssBadge.innerHTML = '<svg viewBox="0 0 16 16" fill="currentColor"><path d="M3.5 2C2.67157 2 2 2.67157 2 3.5V5.5C2 6.32843 2.67157 7 3.5 7H5.5C6.32843 7 7 6.32843 7 5.5V3.5C7 2.67157 6.32843 2 5.5 2H3.5Z"/><path d="M3.5 9C2.67157 9 2 9.67157 2 10.5V12.5C2 13.3284 2.67157 14 3.5 14H5.5C6.32843 14 7 13.3284 7 12.5V10.5C7 9.67157 6.32843 9 5.5 9H3.5Z"/><path d="M9 3.5C9 2.67157 9.67157 2 10.5 2H12.5C13.3284 2 14 2.67157 14 3.5V5.5C14 6.32843 13.3284 7 12.5 7H10.5C9.67157 7 9 6.32843 9 5.5V3.5Z"/><path d="M10.5 9C9.67157 9 9 9.67157 9 10.5V12.5C9 13.3284 9.67157 14 10.5 14H12.5C13.3284 14 14 13.3284 14 12.5V10.5C14 9.67157 13.3284 9 12.5 9H10.5Z"/></svg>';
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
      removeBtn.innerHTML = '<svg width="10" height="10" viewBox="0 0 10 10"><path d="M2.5 2.5l5 5M7.5 2.5l-5 5" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>';
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
        el("div", "drop-target-empty", "Shift+click or drag cells here to build a reel")
      );
      qs("#reelDuration").textContent = "";
      return;
    }

    var totalDur = 0;
    for (var i = 0; i < n; i++) {
      var item = state.reelQueue[i];
      var isIntake = isIntakeSource(item.source);
      var segStart, segDuration;
      if (item.segStart !== undefined && item.segDuration !== undefined) {
        segStart = item.segStart;
        segDuration = item.segDuration;
      } else {
        var parsed = parseClipTimestamps(item.timestamp, item.participant)[0];
        segStart = parsed.startSeconds;
        segDuration = parsed.duration;
      }
      var segTotal = item.segTotal || 1;
      var segIdx = item.segIdx || 0;
      totalDur += segDuration;

      var card = el("div", "queue-card reel-card" + (isIntake ? " queue-card-intake" : ""));
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
        ssObserveThumb(card, img, thumb, item.participant, segStart);
      } else {
        img.src = "api/thumbnail/" + encodeURIComponent(item.participant) + "/" + segStart;
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
          ssBadge.innerHTML = '<svg viewBox="0 0 16 16" fill="currentColor"><path d="M3 3h10v1.5H3zM3 7h8v1.5H3zM3 11h10v1.5H3z"/></svg>';
        } else {
          ssBadge.innerHTML = '<svg viewBox="0 0 16 16" fill="currentColor"><path d="M3.5 2C2.67157 2 2 2.67157 2 3.5V5.5C2 6.32843 2.67157 7 3.5 7H5.5C6.32843 7 7 6.32843 7 5.5V3.5C7 2.67157 6.32843 2 5.5 2H3.5Z"/><path d="M3.5 9C2.67157 9 2 9.67157 2 10.5V12.5C2 13.3284 2.67157 14 3.5 14H5.5C6.32843 14 7 13.3284 7 12.5V10.5C7 9.67157 6.32843 9 5.5 9H3.5Z"/><path d="M9 3.5C9 2.67157 9.67157 2 10.5 2H12.5C13.3284 2 14 2.67157 14 3.5V5.5C14 6.32843 13.3284 7 12.5 7H10.5C9.67157 7 9 6.32843 9 5.5V3.5Z"/><path d="M10.5 9C9.67157 9 9 9.67157 9 10.5V12.5C9 13.3284 9.67157 14 10.5 14H12.5C13.3284 14 14 13.3284 14 12.5V10.5C14 9.67157 13.3284 9 12.5 9H10.5Z"/></svg>';
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
      removeBtn.innerHTML = '<svg width="10" height="10" viewBox="0 0 10 10"><path d="M2.5 2.5l5 5M7.5 2.5l-5 5" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>';
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
    // TODO: no r.ok check; migrating to apiGet would surface HTTP errors via the catch.
    fetch("api/stashes")
      .then(function (r) { return r.json(); })
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

    if (n === 0) {
      if (!area.classList.contains("stash-drop-reveal")) area.classList.add("hidden");
      list.innerHTML = "";
      return;
    }
    area.classList.remove("hidden");
    area.classList.remove("stash-drop-reveal");
    list.innerHTML = "";

    for (var i = 0; i < n; i++) {
      var stash = state.stashes[i];
      var card = el("div", "stash-card");
      card.setAttribute("data-stash-id", stash.id);
      card.setAttribute("draggable", "true");
      (function (stashRef) {
        card.addEventListener("dragstart", function (ev) {
          ev.dataTransfer.setData("application/json", JSON.stringify({
            stashId: stashRef.id,
            items: stashRef.items,
            source: "reel-stash",
          }));
          ev.dataTransfer.effectAllowed = "copy";
        });
      })(stash);

      var nameEl = el("span", "stash-card-name", truncate(stash.name, 20));
      nameEl.title = stash.name;
      (function (stashRef, nameNode) {
        nameNode.addEventListener("click", function (ev) {
          ev.stopPropagation();
          startStashRename(stashRef, nameNode, "api/stashes");
        });
      })(stash, nameEl);
      card.appendChild(nameEl);

      var info = el("div", "stash-card-info");
      info.appendChild(el("span", "", stash.count + " clips"));
      info.appendChild(el("span", "", formatDuration(stash.totalDuration)));
      card.appendChild(info);

      var removeBtn = el("button", "stash-card-remove", "\u00D7");
      removeBtn.title = "Delete stash";
      (function (stashId) {
        removeBtn.addEventListener("click", function (ev) {
          ev.stopPropagation();
          deleteStash(stashId, "api/stashes", state.stashes, renderStashedReels);
        });
      })(stash.id);
      card.appendChild(removeBtn);

      (function (stashRef) {
        card.addEventListener("click", function () {
          recallStash(stashRef);
        });
      })(stash);

      list.appendChild(card);
    }
    computeGridMaxHeight();
  }

  function computeReelDuration(items) {
    var total = 0;
    for (var i = 0; i < items.length; i++) {
      var dur = items[i].segDuration;
      if (dur === undefined || dur === null) {
        var segs = parseClipTimestamps(items[i].timestamp, items[i].participant);
        dur = segs.length > 0 ? segs[0].duration : 0;
      }
      total += dur || 0;
    }
    return total;
  }

  function stashCurrentReel() {
    if (state.reelQueue.length === 0) return;

    var items = state.reelQueue.slice();
    var totalDuration = computeReelDuration(items);
    // TODO: no r.ok check; migrating to apiPost would surface HTTP errors via the catch.
    fetch("api/stashes", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: "create", items: items, name: "", totalDuration: totalDuration }),
    })
      .then(function (r) { return r.json(); })
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
    // TODO: no r.ok check; migrating to apiPost would surface HTTP errors via the catch.
    fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: "delete", id: stashId }),
    })
      .then(function (r) { return r.json(); })
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

      // TODO: fire-and-forget; needs custom shape, no migration to apiPost.
      fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "update", id: stash.id, name: newName }),
      }).catch(function () {});
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
    // TODO: no r.ok check; migrating to apiPost would surface HTTP errors via the catch.
    fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: "create", items: items, name: "", totalDuration: totalDuration }),
    })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data.ok) onSuccess(data.stash);
      })
      .catch(function () {});
  }

  // ---- Stashed artifacts ----

  function loadArtifactStashes() {
    // TODO: no r.ok check; migrating to apiGet would surface HTTP errors via the catch.
    fetch("api/artifact-stashes")
      .then(function (r) { return r.json(); })
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

    if (n === 0) {
      if (!area.classList.contains("stash-drop-reveal")) area.classList.add("hidden");
      list.innerHTML = "";
      return;
    }
    area.classList.remove("hidden");
    area.classList.remove("stash-drop-reveal");
    list.innerHTML = "";

    for (var i = 0; i < n; i++) {
      var stash = state.artifactStashes[i];
      var card = el("div", "stash-card");
      card.setAttribute("data-stash-id", stash.id);
      card.setAttribute("draggable", "true");
      (function (stashRef) {
        card.addEventListener("dragstart", function (ev) {
          ev.dataTransfer.setData("application/json", JSON.stringify({
            stashId: stashRef.id,
            items: stashRef.items,
            source: "artifact-stash",
          }));
          ev.dataTransfer.effectAllowed = "copy";
        });
      })(stash);

      var nameEl = el("span", "stash-card-name", truncate(stash.name, 20));
      nameEl.title = stash.name;
      (function (stashRef, nameNode) {
        nameNode.addEventListener("click", function (ev) {
          ev.stopPropagation();
          startStashRename(stashRef, nameNode, "api/artifact-stashes");
        });
      })(stash, nameEl);
      card.appendChild(nameEl);

      var info = el("div", "stash-card-info");
      info.appendChild(el("span", "", stash.count + " clips"));
      info.appendChild(el("span", "", formatDuration(stash.totalDuration)));
      card.appendChild(info);

      var removeBtn = el("button", "stash-card-remove", "\u00D7");
      removeBtn.title = "Delete stash";
      (function (stashId) {
        removeBtn.addEventListener("click", function (ev) {
          ev.stopPropagation();
          deleteStash(stashId, "api/artifact-stashes", state.artifactStashes, renderStashedArtifacts);
        });
      })(stash.id);
      card.appendChild(removeBtn);

      (function (stashRef) {
        card.addEventListener("click", function () {
          recallArtifactStash(stashRef);
        });
      })(stash);

      list.appendChild(card);
    }
    computeGridMaxHeight();
  }

  function stashCurrentArtifacts() {
    if (state.artifactQueue.length === 0) return;

    var items = state.artifactQueue.slice();
    var totalDuration = computeReelDuration(items);
    // TODO: no r.ok check; migrating to apiPost would surface HTTP errors via the catch.
    fetch("api/artifact-stashes", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: "create", items: items, name: "", totalDuration: totalDuration }),
    })
      .then(function (r) { return r.json(); })
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
    if (state.artifactStashes.length === 0) {
      artArea.classList.remove("hidden");
      artArea.classList.add("stash-drop-reveal");
    }
    if (state.stashes.length === 0) {
      reelArea.classList.remove("hidden");
      reelArea.classList.add("stash-drop-reveal");
    }
    computeGridMaxHeight();
  }

  function hideEmptyStashAreas() {
    var artArea = qs("#stashedArtifactsArea");
    var reelArea = qs("#stashedReelsArea");
    if (state.artifactStashes.length === 0) {
      artArea.classList.add("hidden");
      artArea.classList.remove("stash-drop-reveal");
    }
    if (state.stashes.length === 0) {
      reelArea.classList.add("hidden");
      reelArea.classList.remove("stash-drop-reveal");
    }
    computeGridMaxHeight();
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
    qs("#buildReelBtn").addEventListener("click", onBuildReel);
    qs("#cancelReelBtn").addEventListener("click", onCancelReel);
    qs("#buildViewerBtn").addEventListener("click", onBuildViewer);
    qs("#buildTimelineViewerBtn").addEventListener("click", onBuildTimelineViewer);
    qs("#buildHighlightsBtn").addEventListener("click", onBuildHighlights);
    qs("#galleryBtn").addEventListener("click", onGallery);

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
        onSave: function (_applied, full) {
          state.settingsData = full;
          syncInlineControls();
        },
        onReset: function (_scope, full) {
          state.settingsData = full;
          syncInlineControls();
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
        // TODO: fire-and-forget with no response handling; no clean migration to apiPost.
        fetch("api/open-viewer", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ file: _lastViewerFile }),
        });
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
    badge.innerHTML = success
      ? '<svg width="12" height="12" viewBox="0 0 12 12"><path d="M2.5 6.5l2.5 2.5 4.5-5" fill="none" stroke="#fff" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>'
      : '<svg width="12" height="12" viewBox="0 0 12 12"><path d="M3 3l6 6M9 3l-6 6" fill="none" stroke="#fff" stroke-width="1.5" stroke-linecap="round"/></svg>';
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
      "#buildViewerBtn", "#buildTimelineViewerBtn", "#galleryBtn",
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
    var pending = (sheetItems.length > 0 ? 1 : 0) + (intakeItems.length > 0 ? 1 : 0);

    function finishBranch() {
      if (--pending > 0) return;
      setTitleSpinner("artifactsSpinner", false);
      state.generating = false;
      setGeneratingLock(false);
      updateViewerButton();
      var msg = "Generated " + totalSuccess + " artifact(s)";
      if (totalFail > 0) msg += ", " + totalFail + " failed";
      showResult(
        totalSuccess > 0 ? msg : null,
        totalSuccess === 0 && totalFail > 0 ? "All generations failed" : null
      );
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
        if (!data || !data.cell) return;
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
          for (var ci = 0; ci < cards.length; ci++) setCardResult(cards[ci], false);
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
          start: itm.segStart,
          end: itm.segStart + itm.segDuration,
          event_type: itm.event_type || itm.desc || "",
          event_ids: itm.event_ids || [],
          source: itm.source || "screenspace",
          mark_ids: itm.mark_ids || [],
        };
      });

      // TODO: skips r.ok check — success is signaled by data.ok from server. Migrating to apiPost
      // would add an r.ok throw that this handler doesn't currently trigger.
      fetch("api/generate-intake", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ items: intakePayload, format: format }),
      })
        .then(function (r) { return r.json(); })
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
    // TODO: fire-and-forget POST with no body; apiPost would send JSON.stringify(undefined).
    fetch("api/reel/cancel", { method: "POST" });
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
        var segStart, segDuration;
        if (item.segStart !== undefined && item.segDuration !== undefined) {
          segStart = item.segStart;
          segDuration = item.segDuration;
        } else {
          var parsed = parseClipTimestamps(item.timestamp, item.participant)[0];
          segStart = parsed.startSeconds;
          segDuration = parsed.duration;
        }
        segments.push({
          participant: item.participant,
          start: segStart,
          end: segStart + segDuration,
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
  var _galleryBtnOrigHTML = "";

  function onBuildHighlights() {
    if (state.generating) return;

    var drawer = qs("#highlightsDurationDrawer");
    var btn = qs("#buildHighlightsBtn");
    var isOpen = drawer.classList.contains("open");

    var checkHTML =
      '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
      '<path d="M3 8.5l3.5 3.5 6.5-8"/>' +
      "</svg>";

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
            state.reelQueue.push(data.clips[i]);
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

  function onGallery() {
    if (state.generating) return;

    var drawer = qs("#galleryDrawer");
    var btn = qs("#galleryBtn");
    var isOpen = drawer.classList.contains("open");

    var checkHTML =
      '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
      '<path d="M3 8.5l3.5 3.5 6.5-8"/>' +
      "</svg>";

    if (!isOpen) {
      _galleryBtnOrigHTML = btn.innerHTML;
      drawer.classList.add("open");
      var w = btn.offsetWidth;
      btn.style.minWidth = w + "px";
      btn.innerHTML = checkHTML + "Confirm";
      return;
    }

    var participant = qs("#galleryParticipant").value;
    var format = qs("#galleryFormat").value;
    var interval = parseInt(qs("#galleryInterval").value, 10);
    if (!interval || interval < 1) interval = 10;
    var bundle = qs("#galleryBundle").checked;

    drawer.classList.remove("open");
    btn.style.minWidth = "";
    btn.innerHTML = _galleryBtnOrigHTML;

    if (!participant) {
      showResult(null, "No participant selected for gallery");
      return;
    }

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

  function updateViewerButton() {
    var n = state.generatedArtifacts.length;
    qs("#buildViewerBtn").disabled = n === 0;
    if (n === 0) qs("#buildViewerBtn").setAttribute("data-tooltip", "Generate artifacts first");
    var count = qs("#viewerArtifactCount");
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
    var cellColor = _findSetting("STUDIO_SHEET_CELL_COLOR_CODING");
    if (cellColor) {
      var newVal = !!cellColor.value;
      if (newVal !== state.cellColorCoding) {
        state.cellColorCoding = newVal;
        if (state.sheetData) renderGrid();
      }
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

    // TODO: skips r.ok; only reacts to data.ok. Silent on HTTP failure — apiPut would expose it via .catch.
    fetch("api/settings", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ settings: payload }),
    })
      .then(function (r) { return r.json(); })
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
  }

  function checkNavLinks() {
    // TODO: skips r.ok; a failed request silently leaves tabs hidden. apiGet would catch but
    // the current code already swallows errors via the .catch below.
    fetch("../api/status")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data.screenspace) {
          var intakeTab = qs('.preview-tab[data-tab="intake"]');
          if (intakeTab) intakeTab.classList.remove("hidden");
        }
        if (data.transcripts) {
          var trIntakeTab = qs('.preview-tab[data-tab="transcript-intake"]');
          if (trIntakeTab) trIntakeTab.classList.remove("hidden");
        }
      })
      .catch(function () {});
  }

  // ---- Screenspace Intake ----

  var INTAKE_DETECTOR_COLORS = DETECTOR_COLORS;

  var INTAKE_DETECTOR_ICON_FILES = {
    multitool: "link",
    color: "eye-dropper",
    change: "bolt",
    similarity: "photo",
    text: "language",
    numbers: "hashtag",
    timelapse: "film",
    template: "viewfinder-circle",
    flow: "arrows-right-left",
    scene: "squares-2x2",
    inactivity: "pause-circle",
  };

  // ---- Screenspace thumbnail queue (throttled + cached) ----

  var _ssThumbQueue = [];
  var _ssThumbActive = 0;
  var _SS_THUMB_MAX = 3;
  var _ssThumbCache = {}; // url -> objectURL | "error"

  function ssThumbUrl(participant, timestamp) {
    return "../screenspace/api/video/frame/" + encodeURIComponent(participant) + "/" + timestamp + "?w=200";
  }

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
        var imgEl = entry.target.querySelector(".queue-card-thumb img");
        var tEl = entry.target.querySelector(".queue-card-thumb");
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
        var threshold = parseInt((qs("#intakeClusterThreshold") || {}).value) || 5;
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
    var url = 'url("' + XREF_ICON_BASE + iconName + '.svg")';
    span.style.maskImage = url;
    span.style.webkitMaskImage = url;
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

  var _intakeHitRects = [];
  var _trIntakeHitRects = [];

  function sizeIntakeCanvas() {
    var canvas = qs("#intakeTimeline");
    if (!canvas) return;
    var w = canvas.clientWidth;
    if (w <= 0) return;
    var dpr = window.devicePixelRatio || 1;
    canvas.width = w * dpr;
    canvas.height = 48 * dpr;
    var ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    renderIntakeTimeline();
  }

  function renderIntakeTimeline() {
    var canvas = qs("#intakeTimeline");
    if (!canvas) return;
    var ctx = canvas.getContext("2d");
    var dpr = window.devicePixelRatio || 1;
    var w = canvas.width / dpr;
    var h = canvas.height / dpr;
    if (w <= 0 || h <= 0) return;

    var cs = getComputedStyle(document.documentElement);
    var surfaceAlt = cs.getPropertyValue("--color-surface-alt").trim() || "#f1ece4";
    var borderColor = cs.getPropertyValue("--color-border").trim() || "#e0ddd7";
    var textDim = cs.getPropertyValue("--color-text-dim").trim() || "#6b7280";
    var fontMono = cs.getPropertyValue("--font-mono").trim() || "monospace";

    ctx.clearRect(0, 0, w, h);
    ctx.fillStyle = surfaceAlt;
    ctx.fillRect(0, 0, w, h);

    var clusters = filteredIntakeClusters();
    if (!clusters.length) {
      ctx.fillStyle = textDim;
      ctx.font = "12px -apple-system, sans-serif";
      ctx.textAlign = "center";
      ctx.fillText("No events", w / 2, h / 2 + 4);
      ctx.textAlign = "start";
      _intakeHitRects = [];
      return;
    }

    var maxEnd = 0;
    for (var i = 0; i < clusters.length; i++) {
      if (clusters[i].end > maxEnd) maxEnd = clusters[i].end;
    }
    var duration = Math.max(maxEnd * 1.05, 60);

    function timeToX(t) {
      return (t / duration) * w;
    }

    // Time ruler ticks
    var tickInterval = intakeComputeTickInterval(duration);
    var firstTick = Math.ceil(0 / tickInterval) * tickInterval;
    ctx.strokeStyle = borderColor;
    ctx.fillStyle = textDim;
    ctx.font = "10px " + fontMono;
    ctx.textAlign = "center";
    ctx.lineWidth = 1;
    for (var t = firstTick; t <= duration; t += tickInterval) {
      var x = timeToX(t);
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, 8);
      ctx.stroke();
      ctx.fillText(formatDuration(t), x, 18);
    }
    ctx.textAlign = "start";

    // Cluster markers
    var markerY = 22;
    var markerH = h - markerY - 4;
    _intakeHitRects = [];

    for (var ci = 0; ci < clusters.length; ci++) {
      var c = clusters[ci];
      var color = INTAKE_DETECTOR_COLORS[c.detector] || "#888";
      var highlighted = state.intakeHoveredIdx === ci;
      var dimmed = state.intakeHoveredIdx !== -1 && !highlighted;
      var alpha = highlighted ? 0.85 : (dimmed ? 0.15 : 0.5);

      var x1 = timeToX(c.start);
      var x2 = timeToX(c.end);
      var mw = Math.max(x2 - x1, 3);

      ctx.fillStyle = hexToRgba(color, alpha);
      ctx.fillRect(x1, markerY, mw, markerH);

      _intakeHitRects.push({ x1: x1, x2: x1 + mw, y: markerY, h: markerH, clusterIdx: ci });
    }
  }

  function intakeHitTest(mx, my) {
    for (var i = _intakeHitRects.length - 1; i >= 0; i--) {
      var hr = _intakeHitRects[i];
      if (mx >= hr.x1 && mx <= hr.x2 && my >= hr.y && my <= hr.y + hr.h) return hr;
    }
    return null;
  }

  function highlightIntakeCard(idx) {
    var cards = qsa(".intake-queue-card");
    for (var i = 0; i < cards.length; i++) {
      if (i === idx) {
        cards[i].classList.add("intake-highlight");
        cards[i].scrollIntoView({ block: "nearest", behavior: "smooth" });
      } else {
        cards[i].classList.remove("intake-highlight");
      }
    }
  }

  function buildParticipantPills() {
    var container = qs("#intakeFilterParticipants");
    if (!container) return;
    var seen = {};
    var participants = [];
    for (var i = 0; i < state.intakeClusters.length; i++) {
      var p = state.intakeClusters[i].participant;
      if (p && !seen[p]) { seen[p] = true; participants.push(p); }
    }
    participants.sort();
    var key = participants.join(",");
    if (container.dataset.participants === key) {
      syncParticipantPillStates();
      return;
    }
    container.dataset.participants = key;
    state.intakeFilterParticipants = state.intakeFilterParticipants.filter(function (p) { return seen[p]; });
    container.innerHTML = "";
    for (var j = 0; j < participants.length; j++) {
      var btn = el("button", "intake-filter-participant", participants[j]);
      btn.dataset.participant = participants[j];
      if (state.intakeFilterParticipants.indexOf(participants[j]) !== -1) btn.classList.add("active");
      container.appendChild(btn);
    }
  }

  function syncParticipantPillStates() {
    var btns = qsa(".intake-filter-participant");
    for (var i = 0; i < btns.length; i++) {
      var p = btns[i].dataset.participant;
      if (state.intakeFilterParticipants.indexOf(p) !== -1) btns[i].classList.add("active");
      else btns[i].classList.remove("active");
    }
  }

  function renderIntake(hasNew) {
    ssClearPending();
    var container = qs("#intakeCards");
    var addAllBtn = qs("#intakeAddAllBtn");
    var reelAllBtn = qs("#intakeReelAllBtn");
    var tabBadge = qs("#intakeTabBadge");

    buildParticipantPills();

    if (!state.intakeClusters.length) {
      if (tabBadge) tabBadge.classList.add("hidden");
      container.innerHTML = "";
      container.appendChild(el("div", "drop-target-empty", "Screenspace events will appear here"));
      addAllBtn.disabled = true;
      reelAllBtn.disabled = true;
      sizeIntakeCanvas();
      return;
    }
    if (tabBadge) {
      tabBadge.textContent = state.intakeClusters.length;
      tabBadge.classList.remove("hidden");
    }
    var clusters = filteredIntakeClusters();
    addAllBtn.disabled = clusters.length === 0;
    reelAllBtn.disabled = clusters.length === 0;

    sizeIntakeCanvas();

    container.innerHTML = "";
    if (clusters.length === 0) {
      container.appendChild(el("div", "drop-target-empty", "No events match the current filters"));
      return;
    }
    clusters.forEach(function (c, idx) {
      var segDuration = c.end - c.start;
      var card = el("div", "queue-card intake-queue-card");
      card.dataset.intakeIdx = idx;
      card.setAttribute("draggable", "true");

      var thumb = el("div", "queue-card-thumb");
      var img = document.createElement("img");
      img.alt = "";
      img.draggable = false;
      thumb.appendChild(img);
      ssObserveThumb(card, img, thumb, c.participant, c.start);
      thumb.appendChild(el("span", "queue-card-duration", formatDuration(segDuration)));

      var detDot = el("span", "intake-card-det-dot");
      detDot.style.background = INTAKE_DETECTOR_COLORS[c.detector] || "#888";
      thumb.appendChild(detDot);

      // Source + cross-reference badges (SS self-badge leads the stack)
      var xref = findOverlappingData(c.participant, c.start, c.end);
      var ssSelf = { icon: XREF_BADGES.screenspace.icon, color: XREF_BADGES.screenspace.color, title: c.event_type || "Screenspace" };
      var badgeStack = buildXrefBadges(xref, "screenspace", ssSelf);
      if (badgeStack) thumb.appendChild(badgeStack);
      if (xref.transcriptSnippets.length > 0) {
        card.dataset.transcriptContext = xref.transcriptSnippets.map(function (s) { return s.text; }).join("\n");
      }

      card.appendChild(thumb);

      var meta = el("div", "queue-card-meta");
      meta.appendChild(el("span", "queue-card-ref",
        c.participant + " \u00b7 " + (c.event_type || "intake")));
      card.appendChild(meta);

      // Drag support
      (function (cluster) {
        card.addEventListener("dragstart", function (ev) {
          ev.dataTransfer.setData("application/json", JSON.stringify({
            participant: cluster.participant,
            desc: cluster.event_type,
            segStart: cluster.start,
            segDuration: cluster.end - cluster.start,
            source: "screenspace",
            event_type: cluster.event_type,
            event_ids: cluster.events.map(function (e) { return e.id; }),
          }));
          ev.dataTransfer.effectAllowed = "copyMove";
          setCardDragImage(ev, this);
        });
      })(c);

      container.appendChild(card);
    });
  }

  function intakeAddToArtifacts(cluster) {
    state.artifactQueue.push({
      participant: cluster.participant,
      segStart: cluster.start,
      segDuration: cluster.end - cluster.start,
      desc: cluster.event_type,
      source: "screenspace",
      event_type: cluster.event_type,
      event_ids: cluster.events.map(function (e) { return e.id; }),
    });
    renderArtifactQueue();
  }

  function intakeDismissCluster(cluster) {
    var ids = cluster.events.map(function (e) { return e.id; });
    // TODO: response body ignored; apiPut would parse JSON unnecessarily. Refactor if a
    // fire-and-forget apiPutVoid helper is added.
    fetch("../screenspace/api/events/bulk-exclude", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ids: ids }),
    })
      .then(function () { pollIntakeEvents(); })
      .catch(function () {});
  }

  function intakeAddToReel(cluster) {
    state.reelQueue.push({
      participant: cluster.participant,
      segStart: cluster.start,
      segDuration: cluster.end - cluster.start,
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
    // Set mask-image on filter pill icons
    var iconSpans = qsa(".intake-det-icon");
    for (var k = 0; k < iconSpans.length; k++) {
      var iconName = iconSpans[k].dataset.icon;
      if (iconName) {
        var iconUrl = 'url("../screenspace/icons/' + iconName + '.svg")';
        iconSpans[k].style.maskImage = iconUrl;
        iconSpans[k].style.webkitMaskImage = iconUrl;
      }
    }

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

    // Card hover → highlight timeline marker + transcript tooltip
    intakeCards.addEventListener("mouseover", function (e) {
      var card = e.target.closest(".intake-queue-card");
      if (!card) return;
      var idx = parseInt(card.dataset.intakeIdx);
      if (state.intakeHoveredIdx !== idx) {
        state.intakeHoveredIdx = idx;
        renderIntakeTimeline();
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
        renderIntakeTimeline();
      }
      var trTooltip = qs("#trIntakeTooltip");
      if (trTooltip) trTooltip.classList.add("hidden");
    });

    // Timeline canvas interactions
    var intakeCanvas = qs("#intakeTimeline");
    intakeCanvas.addEventListener("mousemove", function (e) {
      var rect = intakeCanvas.getBoundingClientRect();
      var dpr = window.devicePixelRatio || 1;
      var mx = (e.clientX - rect.left);
      var my = (e.clientY - rect.top);
      var hit = intakeHitTest(mx, my);
      var idx = hit ? hit.clusterIdx : -1;
      if (state.intakeHoveredIdx !== idx) {
        state.intakeHoveredIdx = idx;
        renderIntakeTimeline();
        highlightIntakeCard(idx);
      }
    });
    intakeCanvas.addEventListener("mouseleave", function () {
      if (state.intakeHoveredIdx !== -1) {
        state.intakeHoveredIdx = -1;
        renderIntakeTimeline();
        highlightIntakeCard(-1);
      }
    });
    intakeCanvas.addEventListener("click", function (e) {
      var rect = intakeCanvas.getBoundingClientRect();
      var mx = e.clientX - rect.left;
      var my = e.clientY - rect.top;
      var hit = intakeHitTest(mx, my);
      if (!hit) return;
      var cluster = filteredIntakeClusters()[hit.clusterIdx];
      if (!cluster) return;
      if (e.shiftKey) intakeAddToReel(cluster);
      else intakeAddToArtifacts(cluster);
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

    // Filter controls
    var searchEl = qs("#intakeFilterSearch");
    var _intakeSearchTimer = 0;
    if (searchEl) {
      searchEl.addEventListener("input", function () {
        state.intakeFilterText = this.value;
        clearTimeout(_intakeSearchTimer);
        _intakeSearchTimer = setTimeout(function () { renderIntake(false); }, 250);
      });
    }
    var detBtns = qsa(".intake-filter-det");
    for (var i = 0; i < detBtns.length; i++) {
      detBtns[i].addEventListener("click", function () {
        var val = this.dataset.detector || "";
        state.intakeFilterDetector = state.intakeFilterDetector === val ? "" : val;
        var all = qsa(".intake-filter-det");
        for (var j = 0; j < all.length; j++) all[j].classList.remove("active");
        if (state.intakeFilterDetector) this.classList.add("active");
        renderIntake(false);
      });
    }
    var partContainer = qs("#intakeFilterParticipants");
    if (partContainer) {
      partContainer.addEventListener("click", function (e) {
        var btn = e.target.closest(".intake-filter-participant");
        if (!btn) return;
        var val = btn.dataset.participant;
        var idx = state.intakeFilterParticipants.indexOf(val);
        if (idx === -1) state.intakeFilterParticipants.push(val);
        else state.intakeFilterParticipants.splice(idx, 1);
        syncParticipantPillStates();
        renderIntake(false);
      });
    }
    var newToggle = qs("#intakeFilterNew");
    if (newToggle) {
      newToggle.addEventListener("change", function () {
        state.intakeFilterNew = this.checked;
        renderIntake(false);
      });
    }

    // Start polling immediately — silently fails if Screenspace is unavailable
    pollIntakeEvents();
    state.intakePollTimer = setInterval(pollIntakeEvents, 10000);

    // Pause polling when browser tab or intake panel is not visible
    document.addEventListener("visibilitychange", function () {
      if (document.hidden) {
        if (state.intakePollTimer) { clearInterval(state.intakePollTimer); state.intakePollTimer = null; }
      } else if (state.activePreviewTab === "intake") {
        pollIntakeEvents();
        if (!state.intakePollTimer) state.intakePollTimer = setInterval(pollIntakeEvents, 10000);
      }
    });
  }

  // ---- Transcript Intake ----

  var TR_INTAKE_CATEGORIES = {
    pain_point: { label: "Pain Point", color: "#dc2626" },
    delight:    { label: "Delight",    color: "#16a34a" },
    quote:      { label: "Quote",      color: "#2563eb" },
    insight:    { label: "Insight",    color: "#f97316" },
    task:       { label: "Task Issue", color: "#8b5cf6" },
    bookmark:   { label: "Bookmark",   color: "#0891b2" },
  };

  function pollTranscriptIntakeMarks() {
    // TODO: skips r.ok; migrating to apiGet would catch HTTP errors currently silently swallowed.
    fetch("../transcripts/api/marks")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (!data.ok) return;
        state.trIntakeMarks = data.marks.filter(function (m) { return m.valid; });
        var threshold = parseInt((qs("#trIntakeClusterThreshold") || {}).value) || 5;
        state.trIntakeClusters = clusterTranscriptMarks(state.trIntakeMarks, threshold);

        // If "Show all" is enabled, also fetch all segments as unmark items
        if (state.trIntakeShowAll) {
          // TODO: skips r.ok (same pattern).
          fetch("../transcripts/api/participants")
            .then(function (r2) { return r2.json(); })
            .then(function (pData) {
              if (!pData.ok) return;
              var transcribed = pData.participants.filter(function (p) { return p.has_transcript; });
              var promises = transcribed.map(function (p) {
                // TODO: skips r.ok (same pattern).
                return fetch("../transcripts/api/transcript/" + p.id).then(function (r3) { return r3.json(); });
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
            });
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

    buildTrIntakeParticipantPills();

    if (filtered.length === 0) {
      container.innerHTML = '<div class="drop-target-empty">Transcript marks will appear here</div>';
      renderTrIntakeTimeline();
      return;
    }

    container.innerHTML = "";
    for (var i = 0; i < filtered.length; i++) {
      var c = filtered[i];
      var card = document.createElement("div");
      card.className = "queue-card intake-queue-card tr-intake-queue-card";
      card.draggable = true;
      card.dataset.trIntakeIdx = i;

      // Thumbnail area with frame, category dot, and duration
      var thumb = document.createElement("div");
      thumb.className = "queue-card-thumb";

      var segDuration = Math.max(0, c.end - c.start);
      var img = document.createElement("img");
      img.alt = "";
      img.draggable = false;
      thumb.appendChild(img);
      ssObserveThumb(card, img, thumb, c.participant, c.start);

      var catColor = (TR_INTAKE_CATEGORIES[c.category] || TR_INTAKE_CATEGORIES.bookmark).color;
      var dot = document.createElement("span");
      dot.className = "tr-intake-card-category-dot";
      dot.style.background = catColor;
      thumb.appendChild(dot);

      var dur = document.createElement("span");
      dur.className = "queue-card-duration";
      dur.textContent = formatDuration(segDuration);
      thumb.appendChild(dur);

      // Source + cross-reference badges (TR self-badge leads the stack)
      var xref = findOverlappingData(c.participant, c.start, c.end);
      var trSelf = { icon: XREF_BADGES.transcript.icon, color: XREF_BADGES.transcript.color, title: c.label || c.category || "Transcript" };
      var badgeStack = buildXrefBadges(xref, "transcript", trSelf);
      if (badgeStack) thumb.appendChild(badgeStack);

      card.appendChild(thumb);

      // Metadata
      var meta = document.createElement("div");
      meta.className = "queue-card-meta";
      var ref = document.createElement("span");
      ref.className = "queue-card-ref";
      ref.textContent = c.participant + " \u00b7 " + formatDuration(c.start) + "\u2013" + formatDuration(c.end);
      meta.appendChild(ref);

      // Text snippet
      var textSnippet = document.createElement("span");
      textSnippet.className = "tr-intake-card-text";
      var txt = c.label || c.text || "";
      textSnippet.textContent = txt.length > 80 ? txt.substring(0, 80) + "\u2026" : txt;
      meta.appendChild(textSnippet);
      card.appendChild(meta);

      // Drag support
      (function (cluster) {
        card.addEventListener("dragstart", function (ev) {
          ev.dataTransfer.setData("application/json", JSON.stringify({
            participant: cluster.participant,
            desc: cluster.category || "transcript",
            segStart: cluster.start,
            segDuration: cluster.end - cluster.start,
            source: "transcript",
            mark_ids: cluster.marks.map(function (m) { return m.id; }),
          }));
          ev.dataTransfer.effectAllowed = "copyMove";
          setCardDragImage(ev, this);
        });
      })(c);

      container.appendChild(card);
    }

    renderTrIntakeTimeline();
  }

  function trIntakeAddToArtifacts(cluster) {
    state.artifactQueue.push({
      participant: cluster.participant,
      segStart: cluster.start,
      segDuration: cluster.end - cluster.start,
      desc: cluster.category || "transcript",
      source: "transcript",
      mark_ids: cluster.marks.map(function (m) { return m.id; }),
    });
    renderArtifactQueue();
  }

  function trIntakeAddToReel(cluster) {
    state.reelQueue.push({
      participant: cluster.participant,
      segStart: cluster.start,
      segDuration: cluster.end - cluster.start,
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
    container.innerHTML = "";
    for (var j = 0; j < sorted.length; j++) {
      var btn = document.createElement("button");
      btn.className = "intake-filter-participant" + (state.trIntakeFilterParticipants.indexOf(sorted[j]) >= 0 ? " active" : "");
      btn.textContent = sorted[j];
      btn.dataset.participant = sorted[j];
      container.appendChild(btn);
    }
  }

  function sizeTrIntakeCanvas() {
    var canvas = qs("#trIntakeTimeline");
    if (!canvas) return;
    var w = canvas.clientWidth;
    if (w <= 0) return;
    var dpr = window.devicePixelRatio || 1;
    canvas.width = w * dpr;
    canvas.height = 48 * dpr;
    var ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    renderTrIntakeTimeline();
  }

  function renderTrIntakeTimeline() {
    var canvas = qs("#trIntakeTimeline");
    if (!canvas) return;
    var ctx = canvas.getContext("2d");
    var dpr = window.devicePixelRatio || 1;
    var w = canvas.width / dpr;
    var h = canvas.height / dpr;
    if (w <= 0 || h <= 0) return;

    var cs = getComputedStyle(document.documentElement);
    var surfaceAlt = cs.getPropertyValue("--color-surface-alt").trim() || "#f1ece4";
    var borderColor = cs.getPropertyValue("--color-border").trim() || "#e0ddd7";
    var textDim = cs.getPropertyValue("--color-text-dim").trim() || "#888";
    var fontMono = cs.getPropertyValue("--font-mono").trim() || "monospace";

    ctx.clearRect(0, 0, w, h);
    ctx.fillStyle = surfaceAlt;
    ctx.fillRect(0, 0, w, h);

    var filtered = filteredTranscriptIntakeClusters();
    if (!filtered.length) {
      ctx.fillStyle = textDim;
      ctx.font = "12px -apple-system, sans-serif";
      ctx.textAlign = "center";
      ctx.fillText("No marks", w / 2, h / 2 + 4);
      ctx.textAlign = "start";
      _trIntakeHitRects = [];
      return;
    }

    var maxEnd = 0;
    for (var i = 0; i < filtered.length; i++) {
      if (filtered[i].end > maxEnd) maxEnd = filtered[i].end;
    }
    var duration = Math.max(maxEnd * 1.05, 60);

    function timeToX(t) {
      return (t / duration) * w;
    }

    // Time ruler ticks
    var tickInterval = intakeComputeTickInterval(duration);
    var firstTick = Math.ceil(0 / tickInterval) * tickInterval;
    ctx.strokeStyle = borderColor;
    ctx.fillStyle = textDim;
    ctx.font = "10px " + fontMono;
    ctx.textAlign = "center";
    ctx.lineWidth = 1;
    for (var t = firstTick; t <= duration; t += tickInterval) {
      var x = timeToX(t);
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, 8);
      ctx.stroke();
      ctx.fillText(formatDuration(t), x, 18);
    }
    ctx.textAlign = "start";

    // Cluster markers
    var markerY = 22;
    var markerH = h - markerY - 4;
    _trIntakeHitRects = [];

    for (var ci = 0; ci < filtered.length; ci++) {
      var c = filtered[ci];
      var color = (TR_INTAKE_CATEGORIES[c.category] || TR_INTAKE_CATEGORIES.bookmark).color;
      var highlighted = state.trIntakeHoveredIdx === ci;
      var dimmed = state.trIntakeHoveredIdx !== -1 && !highlighted;
      var alpha = highlighted ? 0.85 : (dimmed ? 0.15 : 0.5);

      var x1 = timeToX(c.start);
      var x2 = timeToX(c.end);
      var mw = Math.max(x2 - x1, 3);

      ctx.fillStyle = hexToRgba(color, alpha);
      ctx.fillRect(x1, markerY, mw, markerH);

      _trIntakeHitRects.push({ x1: x1, x2: x1 + mw, y: markerY, h: markerH, clusterIdx: ci });
    }
  }

  function trIntakeHitTest(mx, my) {
    for (var i = _trIntakeHitRects.length - 1; i >= 0; i--) {
      var hr = _trIntakeHitRects[i];
      if (mx >= hr.x1 && mx <= hr.x2 && my >= hr.y && my <= hr.y + hr.h) return hr;
    }
    return null;
  }

  function highlightTrIntakeCard(idx) {
    var cards = qsa(".tr-intake-queue-card");
    for (var i = 0; i < cards.length; i++) {
      if (i === idx) {
        cards[i].classList.add("intake-highlight");
        cards[i].scrollIntoView({ block: "nearest", behavior: "smooth" });
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

    // Card hover → highlight timeline marker + tooltip
    trIntakeCards.addEventListener("mouseover", function (e) {
      var card = e.target.closest(".tr-intake-queue-card");
      if (!card) return;
      var idx = parseInt(card.dataset.trIntakeIdx);
      if (state.trIntakeHoveredIdx !== idx) {
        state.trIntakeHoveredIdx = idx;
        renderTrIntakeTimeline();
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
        renderTrIntakeTimeline();
      }
      if (trTooltip) trTooltip.classList.add("hidden");
    });

    // Timeline canvas interactions
    var trIntakeCanvas = qs("#trIntakeTimeline");
    if (trIntakeCanvas) {
      trIntakeCanvas.addEventListener("mousemove", function (e) {
        var rect = trIntakeCanvas.getBoundingClientRect();
        var mx = e.clientX - rect.left;
        var my = e.clientY - rect.top;
        var hit = trIntakeHitTest(mx, my);
        var idx = hit ? hit.clusterIdx : -1;
        if (state.trIntakeHoveredIdx !== idx) {
          state.trIntakeHoveredIdx = idx;
          renderTrIntakeTimeline();
          highlightTrIntakeCard(idx);
        }
      });
      trIntakeCanvas.addEventListener("mouseleave", function () {
        if (state.trIntakeHoveredIdx !== -1) {
          state.trIntakeHoveredIdx = -1;
          renderTrIntakeTimeline();
          highlightTrIntakeCard(-1);
        }
      });
      trIntakeCanvas.addEventListener("click", function (e) {
        var rect = trIntakeCanvas.getBoundingClientRect();
        var mx = e.clientX - rect.left;
        var my = e.clientY - rect.top;
        var hit = trIntakeHitTest(mx, my);
        if (!hit) return;
        var cluster = filteredTranscriptIntakeClusters()[hit.clusterIdx];
        if (!cluster) return;
        if (e.shiftKey) trIntakeAddToReel(cluster);
        else trIntakeAddToArtifacts(cluster);
      });
    }

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

    // Category pills
    var catPills = qs("#trIntakeCategoryPills");
    if (catPills) {
      var cats = Object.keys(TR_INTAKE_CATEGORIES);
      for (var i = 0; i < cats.length; i++) {
        (function (key) {
          var cat = TR_INTAKE_CATEGORIES[key];
          var btn = document.createElement("button");
          btn.className = "intake-filter-det tr-intake-filter-cat";
          btn.style.setProperty("--det-color", cat.color);
          btn.textContent = cat.label;
          btn.dataset.category = key;
          btn.addEventListener("click", function () {
            if (state.trIntakeFilterCategory === key) {
              state.trIntakeFilterCategory = "";
              btn.classList.remove("active");
            } else {
              state.trIntakeFilterCategory = key;
              var all = catPills.querySelectorAll(".tr-intake-filter-cat");
              for (var j = 0; j < all.length; j++) all[j].classList.remove("active");
              btn.classList.add("active");
            }
            renderTranscriptIntake();
          });
          catPills.appendChild(btn);
        })(cats[i]);
      }
    }

    // Participant pills (delegated)
    var partPills = qs("#trIntakeFilterParticipants");
    if (partPills) {
      partPills.addEventListener("click", function (e) {
        var btn = e.target.closest(".intake-filter-participant");
        if (!btn) return;
        var pid = btn.dataset.participant;
        var idx = state.trIntakeFilterParticipants.indexOf(pid);
        if (idx >= 0) {
          state.trIntakeFilterParticipants.splice(idx, 1);
          btn.classList.remove("active");
        } else {
          state.trIntakeFilterParticipants.push(pid);
          btn.classList.add("active");
        }
        renderTranscriptIntake();
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

    // Visibility change — pause/resume polling
    document.addEventListener("visibilitychange", function () {
      if (document.hidden) {
        if (state.trIntakePollTimer) { clearInterval(state.trIntakePollTimer); state.trIntakePollTimer = null; }
      } else if (state.activePreviewTab === "transcript-intake") {
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

  document.addEventListener("click", function (ev) {
    var wrap = qs(".filter-cat-wrap");
    if (wrap && !wrap.contains(ev.target)) {
      var btn = wrap.querySelector(".filter-cat-btn");
      var panel = wrap.querySelector(".filter-cat-panel");
      if (btn) btn.classList.remove("open");
      if (panel) panel.classList.add("hidden");
    }
  });

  document.addEventListener("DOMContentLoaded", function () {
    initThemeToggle();
    initTooltipToggle();
    initFilterToggle();
    initPreviewTabs();
    initDropTargets();
    initWheelScroll();
    bindReelReorder();
    bindButtons();
    initPanelDivider();
    loadSheetData();
    loadStashes();
    loadArtifactStashes();
    updateViewerButton();
    checkNavLinks();
    initFrontendSwitcher();
    initIntake();
    initTranscriptIntake();
    window.addEventListener("resize", function () {
      computeGridMaxHeight();
      sizeIntakeCanvas();
      sizeTrIntakeCanvas();
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
  window._studioHexToRgba = hexToRgba;
  window._studioFormatDuration = formatDuration;
  window._studioFindOverlappingData = findOverlappingData;
  window._studioBuildXrefBadges = buildXrefBadges;
  window._studioRenderArtifactQueue = renderArtifactQueue;
  window._studioRenderReelQueue = renderReelQueue;
  window._studioClusterIntakeEvents = clusterIntakeEvents;
  window._studioClusterTranscriptMarks = clusterTranscriptMarks;
  window._studioROW_FUNCTIONS = ROW_FUNCTIONS;
  window._studioSEVERITY_ORDER = SEVERITY_ORDER;
  window._studioSeverityRank = severityRank;
  window._studioSyncPreviewTab = syncPreviewTab;
})();
