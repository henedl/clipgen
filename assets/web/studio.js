/* clipgen Studio */

(function () {
  "use strict";

  var THEME_STORAGE_KEY = "clipgen-studio-theme";
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
    activeFunction: "",
    cellExpandHover: true,
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
  };

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
        if (c && c.valid) total += parseClipTimestamps(c.value).length;
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

  function qs(sel) {
    return document.querySelector(sel);
  }
  function qsa(sel) {
    return document.querySelectorAll(sel);
  }

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
  }

  function truncate(str, max) {
    if (!str) return "";
    return str.length > max ? str.slice(0, max) + "\u2026" : str;
  }

  function severityClass(raw) {
    if (!raw) return "";
    var k = raw.trim().toLowerCase();
    var map = {
      critical: "sev-critical",
      high: "sev-high",
      medium: "sev-medium",
      low: "sev-low",
      "n/a": "sev-na",
      positive: "sev-positive",
      "very positive": "sev-very-positive",
    };
    return map[k] || "sev-unknown";
  }

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
    var segments = parseClipTimestamps(info.timestamp);
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

  function parseTimestampToSeconds(ts) {
    var parts = ts.split(":");
    if (parts.length === 3)
      return (
        parseInt(parts[0], 10) * 3600 +
        parseInt(parts[1], 10) * 60 +
        parseInt(parts[2], 10)
      );
    if (parts.length === 2)
      return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
    return NaN;
  }

  function parseClipTimestamps(raw) {
    var DEFAULT_DUR = 60;
    var cleaned = raw
      .toLowerCase()
      .replace(/!key/g, "")
      .replace(/[+;,]/g, " ");
    var tokens = cleaned.split(/\s+/).filter(function (t) {
      return t && t !== "x";
    });
    var segments = [];
    for (var i = 0; i < tokens.length; i++) {
      var tok = tokens[i].replace(/\.$/, "").replace(/\./g, ":");
      var dashIdx = -1;
      for (var d = 1; d < tok.length; d++) {
        if (tok[d] === "-" && tok[d - 1] >= "0" && tok[d - 1] <= "9") {
          dashIdx = d;
          break;
        }
      }
      if (dashIdx > 0) {
        var s = parseTimestampToSeconds(tok.substring(0, dashIdx));
        var e = parseTimestampToSeconds(tok.substring(dashIdx + 1));
        if (!isNaN(s) && !isNaN(e)) {
          segments.push({ startSeconds: Math.floor(s), duration: Math.max(0, e - s) });
        }
      } else if (tok.indexOf(":") > 0) {
        var sec = parseTimestampToSeconds(tok);
        if (!isNaN(sec)) {
          segments.push({ startSeconds: Math.floor(sec), duration: DEFAULT_DUR });
        }
      }
    }
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
        renderGrid();
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
        renderGrid();
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
        renderGrid();
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
        renderGrid();
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
      renderGrid();
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
      renderGrid();
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
          renderGrid();
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

  // ---- Theme ----

  function initThemeToggle() {
    applyStoredThemePreference();
    var btn = qs("#themeToggle");
    if (!btn) return;
    btn.addEventListener("click", function () {
      toggleThemePreference();
    });
  }

  function applyStoredThemePreference() {
    var stored = null;
    try {
      stored = window.localStorage.getItem(THEME_STORAGE_KEY);
    } catch (_) {}
    var root = document.documentElement;
    if (stored === "light" || stored === "dark") {
      root.setAttribute("data-theme", stored);
    } else {
      root.removeAttribute("data-theme");
    }
    updateThemeToggleButton(stored);
  }

  function toggleThemePreference() {
    var root = document.documentElement;
    var current = root.getAttribute("data-theme");
    var next;
    if (current === "dark") {
      next = "light";
    } else if (current === "light") {
      next = "dark";
    } else {
      var prefersDark = false;
      try {
        prefersDark =
          window.matchMedia &&
          window.matchMedia("(prefers-color-scheme: dark)").matches;
      } catch (_) {}
      next = prefersDark ? "light" : "dark";
    }
    root.setAttribute("data-theme", next);
    try {
      window.localStorage.setItem(THEME_STORAGE_KEY, next);
    } catch (_) {}
    updateThemeToggleButton(next);
  }

  function updateThemeToggleButton(theme) {
    var btn = qs("#themeToggle");
    if (!btn) return;
    btn.setAttribute("data-theme", theme || "");
    btn.setAttribute("aria-pressed", theme === "dark" ? "true" : "false");
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
    fetch("api/sheet")
      .then(function (r) {
        if (!r.ok) throw new Error("Server error " + r.status);
        return r.json();
      })
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

  function loadManifestState() {
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

  function renderGrid() {
    var d = state.sheetData;
    var grid = qs("#sheetGrid");
    grid.innerHTML = "";

    var showSeverity = hasSeverityData(d.rows);
    var metaCols = showSeverity ? 5 : 4;
    var totalCols = metaCols + d.participants.length;
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

    // Group consecutive empty rows into separators
    var filteredRows = getFilteredRows(d.rows);
    var tbody = el("tbody");
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
        tbody.appendChild(sepTr);
      } else {
        tbody.appendChild(renderDataRow(row, d.participants, showSeverity));
        i++;
      }
    }
    table.appendChild(tbody);
    grid.appendChild(table);

    bindGridEvents();
    bindDragFromGrid();
    if (state.activeFunction) updateFunctionColumn();
  }

  // ---- Panel divider (resizable split between sheet preview and bottom panel) ----

  function computeGridMaxHeight() {
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
    var available = window.innerHeight - headerRect.top - headerRect.height
      - sheetChrome - divider.offsetHeight - bottom.offsetHeight;

    var MIN_GRID = 100;
    var maxAllowed = Math.max(0, available - MIN_GRID);
    state.dividerOffset = Math.min(state.dividerOffset, maxAllowed);

    grid.style.maxHeight = Math.max(MIN_GRID, available - state.dividerOffset) + "px";
  }

  function initPanelDivider() {
    var handle = qs("#panelDivider");
    if (!handle) return;
    var dragging = false;
    var startY = 0;
    var startOffset = 0;

    function onDown(e) {
      e.preventDefault();
      dragging = true;
      startY = e.clientY || (e.touches && e.touches[0].clientY) || 0;
      startOffset = state.dividerOffset;
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
        var grid = qs("#sheetGrid");
        var table = grid ? grid.querySelector("table") : null;
        var tableH = table ? table.offsetHeight : 0;
        var maxOff = Math.max(0, tableH - 100);
        state.dividerOffset = Math.max(0, Math.min(maxOff, startOffset + delta));
        computeGridMaxHeight();
        rafPending = false;
      });
    }

    function onUp() {
      if (!dragging) return;
      dragging = false;
      handle.classList.remove("active");
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    }

    document.addEventListener("mouseup", onUp);
    document.addEventListener("touchend", onUp);
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
    updateCellClasses();
  }

  function toggleReelCell(info) {
    if (findInQueue(state.reelQueue, info.participant, info.row) >= 0) {
      removeAllCellEntries(state.reelQueue, info.participant, info.row);
    } else {
      var entries = expandCellToSegments(info);
      for (var i = 0; i < entries.length; i++) state.reelQueue.push(entries[i]);
    }
    renderReelQueue();
    updateCellClasses();
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
      updateCellClasses();
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
      var isIntake = item.source === "screenspace";
      var segStart, segDuration;
      if (item.segStart !== undefined && item.segDuration !== undefined) {
        segStart = item.segStart;
        segDuration = item.segDuration;
      } else {
        var parsed = parseClipTimestamps(item.timestamp)[0];
        segStart = parsed.startSeconds;
        segDuration = parsed.duration;
      }
      var segTotal = item.segTotal || 1;
      var segIdx = item.segIdx || 0;

      var card = el("div", "queue-card" + (isIntake ? " queue-card-intake" : ""));
      card.setAttribute("data-participant", item.participant);
      card.setAttribute("data-row", isIntake ? "" : item.row);
      if (isIntake) card.setAttribute("data-source", "screenspace");
      card.setAttribute("data-seg-idx", segIdx);
      card.setAttribute("draggable", "true");
      (function (itm, isI) {
        card.addEventListener("dragstart", function (ev) {
          var data = {
            participant: itm.participant,
            desc: itm.desc,
            segStart: itm.segStart,
            segDuration: itm.segDuration,
            source: isI ? "screenspace" : "artifact",
          };
          if (!isI) {
            data.row = itm.row;
            data.timestamp = itm.timestamp;
            data.segIdx = itm.segIdx;
            data.segTotal = itm.segTotal;
          } else {
            data.event_type = itm.event_type;
            data.event_ids = itm.event_ids;
          }
          ev.dataTransfer.setData("application/json", JSON.stringify(data));
          ev.dataTransfer.effectAllowed = "copyMove";
        });
      })(item, isIntake);

      var thumb = el("div", "queue-card-thumb");
      var img = document.createElement("img");
      img.src = isIntake
        ? ("../screenspace/api/video/frame/" + encodeURIComponent(item.participant) + "/" + segStart)
        : ("api/thumbnail/" + encodeURIComponent(item.participant) + "/" + segStart);
      img.loading = "lazy";
      img.alt = "";
      img.draggable = false;
      (function (cardEl, thumbEl) {
        img.addEventListener("error", function () {
          this.remove();
          thumbEl.appendChild(el("span", "", "\u2715"));
          cardEl.classList.add("queue-card-error");
        });
      })(card, thumb);
      thumb.appendChild(img);
      thumb.appendChild(el("span", "queue-card-duration", formatDuration(segDuration)));
      if (isIntake) {
        var ssBadge = el("span", "queue-card-source-badge");
        ssBadge.innerHTML = '<svg viewBox="0 0 16 16" fill="currentColor"><path d="M3.5 2C2.67157 2 2 2.67157 2 3.5V5.5C2 6.32843 2.67157 7 3.5 7H5.5C6.32843 7 7 6.32843 7 5.5V3.5C7 2.67157 6.32843 2 5.5 2H3.5Z"/><path d="M3.5 9C2.67157 9 2 9.67157 2 10.5V12.5C2 13.3284 2.67157 14 3.5 14H5.5C6.32843 14 7 13.3284 7 12.5V10.5C7 9.67157 6.32843 9 5.5 9H3.5Z"/><path d="M9 3.5C9 2.67157 9.67157 2 10.5 2H12.5C13.3284 2 14 2.67157 14 3.5V5.5C14 6.32843 13.3284 7 12.5 7H10.5C9.67157 7 9 6.32843 9 5.5V3.5Z"/><path d="M10.5 9C9.67157 9 9 9.67157 9 10.5V12.5C9 13.3284 9.67157 14 10.5 14H12.5C13.3284 14 14 13.3284 14 12.5V10.5C14 9.67157 13.3284 9 12.5 9H10.5Z"/></svg>';
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
      var isIntake = item.source === "screenspace";
      var segStart, segDuration;
      if (item.segStart !== undefined && item.segDuration !== undefined) {
        segStart = item.segStart;
        segDuration = item.segDuration;
      } else {
        var parsed = parseClipTimestamps(item.timestamp)[0];
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
      if (isIntake) card.setAttribute("data-source", "screenspace");
      card.setAttribute("data-seg-idx", segIdx);
      card.setAttribute("draggable", "true");

      var thumb = el("div", "queue-card-thumb");
      var img = document.createElement("img");
      img.src = isIntake
        ? ("../screenspace/api/video/frame/" + encodeURIComponent(item.participant) + "/" + segStart)
        : ("api/thumbnail/" + encodeURIComponent(item.participant) + "/" + segStart);
      img.loading = "lazy";
      img.alt = "";
      img.draggable = false;
      (function (cardEl, thumbEl) {
        img.addEventListener("error", function () {
          this.remove();
          thumbEl.appendChild(el("span", "", "\u2715"));
          cardEl.classList.add("queue-card-error");
        });
      })(card, thumb);
      thumb.appendChild(img);
      thumb.appendChild(el("span", "queue-card-duration", formatDuration(segDuration)));
      if (isIntake) {
        var ssBadge = el("span", "queue-card-source-badge");
        ssBadge.innerHTML = '<svg viewBox="0 0 16 16" fill="currentColor"><path d="M3.5 2C2.67157 2 2 2.67157 2 3.5V5.5C2 6.32843 2.67157 7 3.5 7H5.5C6.32843 7 7 6.32843 7 5.5V3.5C7 2.67157 6.32843 2 5.5 2H3.5Z"/><path d="M3.5 9C2.67157 9 2 9.67157 2 10.5V12.5C2 13.3284 2.67157 14 3.5 14H5.5C6.32843 14 7 13.3284 7 12.5V10.5C7 9.67157 6.32843 9 5.5 9H3.5Z"/><path d="M9 3.5C9 2.67157 9.67157 2 10.5 2H12.5C13.3284 2 14 2.67157 14 3.5V5.5C14 6.32843 13.3284 7 12.5 7H10.5C9.67157 7 9 6.32843 9 5.5V3.5Z"/><path d="M10.5 9C9.67157 9 9 9.67157 9 10.5V12.5C9 13.3284 9.67157 14 10.5 14H12.5C13.3284 14 14 13.3284 14 12.5V10.5C14 9.67157 13.3284 9 12.5 9H10.5Z"/></svg>';
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
        var segs = parseClipTimestamps(items[i].timestamp);
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

    qs("#settingsBtn").addEventListener("click", openSettings);
    qs("#settingsClose").addEventListener("click", closeSettings);
    qs("#settingsResetBtn").addEventListener("click", resetSettings);
    qs("#settingsOverlay").addEventListener("click", function (e) {
      if (e.target === qs("#settingsOverlay")) closeSettings();
    });

    qs("#statusDismiss").addEventListener("click", hideOverlay);
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
      if (items[ci].source === "screenspace") {
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
      if (allArtifacts.length > 0) {
        state.generatedArtifacts = state.generatedArtifacts.concat(allArtifacts);
      }
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
          if (data.artifacts) allArtifacts = allArtifacts.concat(data.artifacts);
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
        };
      });

      fetch("api/generate-intake", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ items: intakePayload, format: format }),
      })
        .then(function (r) { return r.json(); })
        .then(function (data) {
          var intakeCards = list.querySelectorAll('[data-source="screenspace"]');
          if (data.ok && data.results) {
            for (var ri = 0; ri < data.results.length; ri++) {
              var res = data.results[ri];
              if (res.ok) {
                totalSuccess++;
                if (res.artifact) allArtifacts.push(res.artifact);
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
          var intakeCards = list.querySelectorAll('[data-source="screenspace"]');
          for (var j = 0; j < intakeCards.length; j++) setCardResult(intakeCards[j], false);
          totalFail += intakeItems.length;
          finishBranch();
        });
    }
  }

  function onBuildReel() {
    if (state.generating || state.reelQueue.length === 0) return;
    state.generating = true;
    setGeneratingLock(true);
    setTitleSpinner("reelSpinner", true);

    // Determine if we have intake items — use direct endpoint for mixed/intake reels
    var hasIntake = false;
    for (var ci = 0; ci < state.reelQueue.length; ci++) {
      if (state.reelQueue[ci].source === "screenspace") { hasIntake = true; break; }
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
          var parsed = parseClipTimestamps(item.timestamp)[0];
          segStart = parsed.startSeconds;
          segDuration = parsed.duration;
        }
        segments.push({
          participant: item.participant,
          start: segStart,
          end: segStart + segDuration,
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

    fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(reelBody),
    })
      .then(function (r) {
        if (!r.ok) throw new Error("Server error " + r.status);
        return r.json();
      })
      .then(function (data) {
        state.generating = false;
        setTitleSpinner("reelSpinner", false);
        setGeneratingLock(false);

        var cards = list.querySelectorAll(".queue-card");
        for (var j = 0; j < cards.length; j++) {
          setCardResult(cards[j], !!data.ok);
        }

        if (data.ok) {
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

    fetch("api/viewer", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({}),
    })
      .then(function (r) {
        if (!r.ok) throw new Error("Server error " + r.status);
        return r.json();
      })
      .then(function (data) {
        state.generating = false;
        if (data.ok) {
          showResult("Viewer created: " + (data.file || ""), null);
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
    state.generating = true;

    showOverlay("Building timeline viewer...");

    fetch("api/timeline-viewer", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({}),
    })
      .then(function (r) {
        if (!r.ok) throw new Error("Server error " + r.status);
        return r.json();
      })
      .then(function (data) {
        state.generating = false;
        if (data.ok) {
          var msg = "Timeline viewer created: " + (data.file || "");
          if (data.generated) msg = "Generated " + data.generated + " clip(s). " + msg;
          updateViewerButton();
          showResult(msg, null);
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

    fetch("api/highlights-preview", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ highlights_duration: duration }),
    })
      .then(function (r) {
        if (!r.ok) throw new Error("Server error " + r.status);
        return r.json();
      })
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

    drawer.classList.remove("open");
    btn.style.minWidth = "";
    btn.innerHTML = _galleryBtnOrigHTML;

    if (!participant) {
      showResult(null, "No participant selected for gallery");
      return;
    }

    state.generating = true;
    showOverlay("Generating gallery viewer for " + participant + "...");

    fetch("api/gallery", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ participant: participant, format: format, interval: interval }),
    })
      .then(function (r) {
        if (!r.ok) throw new Error("Server error " + r.status);
        return r.json();
      })
      .then(function (data) {
        state.generating = false;
        if (data.ok) {
          showResult("Gallery viewer created: " + (data.file || ""), null);
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

  function showOverlay(message) {
    qs("#statusSpinner").style.display = "";
    qs("#statusTitle").textContent = message;
    qs("#statusMessage").textContent = "";
    qs("#statusMessage").className = "";
    qs("#statusDismiss").classList.add("hidden");
    qs("#statusOverlay").classList.remove("hidden");
  }

  function showResult(successMsg, errorMsg) {
    qs("#statusSpinner").style.display = "none";
    if (errorMsg) {
      qs("#statusTitle").textContent = "Error";
      qs("#statusMessage").textContent = errorMsg;
      qs("#statusMessage").className = "error-text";
    } else {
      qs("#statusTitle").textContent = "Done";
      qs("#statusMessage").textContent = successMsg || "";
      qs("#statusMessage").className = "";
    }
    qs("#statusDismiss").classList.remove("hidden");
  }

  function hideOverlay() {
    qs("#statusOverlay").classList.add("hidden");
  }

  // ---- Settings ----

  var _settingsSaveTimer = null;

  function openSettings() {
    qs("#settingsOverlay").classList.remove("hidden");
    loadSettings();
  }

  function closeSettings() {
    qs("#settingsOverlay").classList.add("hidden");
  }

  function loadSettings() {
    qs("#settingsContent").textContent = "Loading settings\u2026";
    fetch("api/settings")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (!data.ok) {
          qs("#settingsContent").textContent = "Failed to load settings.";
          return;
        }
        state.settingsData = data.settings;
        renderSettings();
      })
      .catch(function () {
        qs("#settingsContent").textContent = "Failed to load settings.";
      });
  }

  function renderSettings() {
    var container = qs("#settingsContent");
    container.innerHTML = "";

    var groups = {};
    var groupOrder = [];
    for (var i = 0; i < state.settingsData.length; i++) {
      var s = state.settingsData[i];
      if (!groups[s.group]) {
        groups[s.group] = [];
        groupOrder.push(s.group);
      }
      groups[s.group].push(s);
    }

    for (var gi = 0; gi < groupOrder.length; gi++) {
      var groupName = groupOrder[gi];
      container.appendChild(el("div", "settings-group-label", groupName));
      var items = groups[groupName];
      for (var si = 0; si < items.length; si++) {
        container.appendChild(buildSettingRow(items[si]));
      }
    }
  }

  function buildSettingRow(s) {
    var row = el("div", "settings-row");
    if (s.value !== s.default) row.classList.add("settings-changed");
    row.setAttribute("data-setting", s.name);

    var labelDiv = el("div", "settings-label");
    var friendlyName = s.name
      .replace(/_/g, " ").toLowerCase()
      .replace(/\b\w/g, function (c) { return c.toUpperCase(); })
      .replace(/Mb$/i, "(MB)").replace(/Seconds$/i, "(s)");
    labelDiv.appendChild(el("div", "settings-label-name", friendlyName));
    labelDiv.appendChild(el("div", "settings-label-desc", s.description));

    var controlDiv = el("div", "settings-control");
    var settingName = s.name;

    if (s.type === "bool") {
      var toggle = document.createElement("input");
      toggle.type = "checkbox";
      toggle.className = "settings-toggle";
      toggle.checked = !!s.value;
      toggle.addEventListener("change", function () {
        var setting = findSetting(settingName);
        if (setting) setting.value = this.checked;
        updateSettingChanged(settingName);
        scheduleSaveSettings();
      });
      controlDiv.appendChild(toggle);
    } else if (s.type === "select" && s.options) {
      var sel = document.createElement("select");
      for (var oi = 0; oi < s.options.length; oi++) {
        var opt = document.createElement("option");
        opt.value = s.options[oi];
        opt.textContent = s.options[oi];
        if (s.options[oi] === s.value) opt.selected = true;
        sel.appendChild(opt);
      }
      sel.addEventListener("change", function () {
        var setting = findSetting(settingName);
        if (setting) setting.value = this.value;
        updateSettingChanged(settingName);
        scheduleSaveSettings();
      });
      controlDiv.appendChild(sel);
    } else {
      var input = document.createElement("input");
      input.type = "number";
      if (s.min !== undefined && s.min !== null) input.min = s.min;
      if (s.step !== undefined && s.step !== null) input.step = s.step;
      input.value = s.value;
      input.placeholder = String(s.default);
      input.addEventListener("change", function () {
        var setting = findSetting(settingName);
        if (setting) setting.value = parseInt(this.value, 10) || 0;
        updateSettingChanged(settingName);
        scheduleSaveSettings();
      });
      controlDiv.appendChild(input);
    }

    row.appendChild(labelDiv);
    row.appendChild(controlDiv);
    return row;
  }

  function findSetting(name) {
    if (!state.settingsData) return null;
    for (var i = 0; i < state.settingsData.length; i++) {
      if (state.settingsData[i].name === name) return state.settingsData[i];
    }
    return null;
  }

  function updateSettingChanged(name) {
    var setting = findSetting(name);
    if (!setting) return;
    var row = qs('.settings-row[data-setting="' + name + '"]');
    if (!row) return;
    if (setting.value !== setting.default) {
      row.classList.add("settings-changed");
    } else {
      row.classList.remove("settings-changed");
    }
  }

  function scheduleSaveSettings() {
    if (_settingsSaveTimer) clearTimeout(_settingsSaveTimer);
    _settingsSaveTimer = setTimeout(saveSettings, 400);
  }

  function saveSettings() {
    if (!state.settingsData) return;
    var payload = {};
    for (var i = 0; i < state.settingsData.length; i++) {
      var s = state.settingsData[i];
      payload[s.name] = s.value;
    }

    var statusEl = qs("#settingsSaveStatus");
    if (statusEl) statusEl.textContent = "Saving\u2026";

    fetch("api/settings", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ settings: payload }),
    })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (statusEl) {
          statusEl.textContent = data.ok ? "Saved" : "Save failed";
          setTimeout(function () { statusEl.textContent = ""; }, 2000);
        }
        if (data.ok) syncInlineControls();
      })
      .catch(function () {
        if (statusEl) statusEl.textContent = "Save failed";
      });
  }

  function resetSettings() {
    if (!state.settingsData) return;
    var payload = {};
    for (var i = 0; i < state.settingsData.length; i++) {
      var s = state.settingsData[i];
      s.value = s.default;
      payload[s.name] = s.default;
    }
    renderSettings();

    fetch("api/settings", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ settings: payload }),
    })
      .then(function (r) { return r.json(); })
      .then(function () {
        var statusEl = qs("#settingsSaveStatus");
        if (statusEl) {
          statusEl.textContent = "Reset to defaults";
          setTimeout(function () { statusEl.textContent = ""; }, 2000);
        }
        syncInlineControls();
      })
      .catch(function () {});
  }

  function syncInlineControls() {
    if (!state.settingsData) return;
    var tcEnabled = findSetting("TITLECARDS_ENABLED");
    var tcDuration = findSetting("TITLECARD_DURATION_SECONDS");
    var hlDuration = findSetting("HIGHLIGHTS_REEL_DURATION_SECONDS");

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
    var cellExpand = findSetting("STUDIO_CELL_EXPAND_HOVER");
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

    fetch("api/settings", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ settings: payload }),
    })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data.ok && state.settingsData) {
          var tcE = findSetting("TITLECARDS_ENABLED");
          var tcD = findSetting("TITLECARD_DURATION_SECONDS");
          if (tcE) tcE.value = cb.checked;
          if (tcD) tcD.value = parseInt(dur.value, 10) || 2;
        }
      })
      .catch(function () {});
  }

  // ---- Init ----

  function checkNavLinks() {
    fetch("../api/status")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data.insights) {
          var link = qs("#insightsLink");
          if (link) link.classList.remove("hidden");
        }
        if (data.screenspace) {
          var link = qs("#screenspaceLink");
          if (link) link.classList.remove("hidden");
        }
      })
      .catch(function () {});
  }

  // ---- Screenspace Intake ----

  var INTAKE_DETECTOR_COLORS = {
    color: "#8b5cf6",
    change: "#f97316",
    similarity: "#0ea5e9",
    text: "#10b981",
    numbers: "#eab308",
  };

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
    fetch("../screenspace/api/events?excluded=false")
      .then(function (r) {
        if (!r.ok) throw new Error("status " + r.status);
        return r.json();
      })
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
      })
      .catch(function (err) {
        console.warn("[Intake] poll failed:", err);
      });
  }

  function renderIntake(hasNew) {
    var area = qs("#intakeArea");
    var list = qs("#intakeList");
    var countEl = qs("#intakeCount");
    var addAllBtn = qs("#intakeAddAllBtn");
    var reelAllBtn = qs("#intakeReelAllBtn");
    var badge = qs("#intakeNewBadge");

    if (!state.intakeClusters.length) {
      if (!area.classList.contains("hidden")) {
        area.classList.add("hidden");
        computeGridMaxHeight();
      }
      return;
    }
    var wasHidden = area.classList.contains("hidden");
    area.classList.remove("hidden");
    if (wasHidden) computeGridMaxHeight();
    var clusters = filteredIntakeClusters();
    countEl.textContent = "(" + clusters.length + "/" + state.intakeClusters.length + ")";
    addAllBtn.disabled = clusters.length === 0;
    reelAllBtn.disabled = clusters.length === 0;

    if (hasNew && badge) badge.classList.remove("hidden");

    list.innerHTML = "";
    if (clusters.length === 0) {
      list.appendChild(el("div", "drop-target-empty", "No events match the current filters"));
      return;
    }
    clusters.forEach(function (c, idx) {
      var card = el("div", "intake-card");
      card.dataset.intakeIdx = idx;

      var header = el("div", "intake-card-header");
      var pid = el("span", "intake-participant", c.participant);
      header.appendChild(pid);
      var detBadge = el("span", "intake-detector-badge", c.detector);
      detBadge.style.background = INTAKE_DETECTOR_COLORS[c.detector] || "#888";
      header.appendChild(detBadge);
      card.appendChild(header);

      var info = el("div", "intake-card-info");
      info.appendChild(el("span", "intake-time", formatDuration(c.start) + " \u2013 " + formatDuration(c.end)));
      info.appendChild(el("span", "intake-label", c.event_type));
      if (c.region) info.appendChild(el("span", "intake-region", c.region));
      info.appendChild(el("span", "intake-count", c.events.length + " event" + (c.events.length !== 1 ? "s" : "")));
      card.appendChild(info);

      var confBar = el("div", "intake-conf-bar");
      var confFill = el("div", "intake-conf-fill");
      confFill.style.width = Math.round(c.confidence_avg * 100) + "%";
      confBar.appendChild(confFill);
      card.appendChild(confBar);

      var actions = el("div", "intake-card-actions");
      var addBtn = el("button", "btn btn-small", "Add to Artifacts");
      addBtn.dataset.action = "add-artifact";
      actions.appendChild(addBtn);
      var reelBtn = el("button", "btn btn-small", "Add to Reel");
      reelBtn.dataset.action = "add-reel";
      actions.appendChild(reelBtn);
      var dismissBtn = el("button", "btn btn-small btn-ghost", "\u00d7");
      dismissBtn.dataset.action = "dismiss";
      dismissBtn.title = "Dismiss (exclude events)";
      actions.appendChild(dismissBtn);
      card.appendChild(actions);

      list.appendChild(card);
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
    if (!text && !det && !onlyNew) return clusters;
    return clusters.filter(function (c) {
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
    var area = qs("#intakeArea");
    if (localStorage.getItem("studio-intake-collapsed")) {
      area.classList.add("intake-collapsed");
    }
    qs("#intakeHeader").addEventListener("click", function () {
      area.classList.toggle("intake-collapsed");
      if (area.classList.contains("intake-collapsed")) {
        localStorage.setItem("studio-intake-collapsed", "1");
      } else {
        localStorage.removeItem("studio-intake-collapsed");
      }
      computeGridMaxHeight();
    });

    var list = qs("#intakeList");
    list.addEventListener("click", function (e) {
      var btn = e.target.closest("[data-action]");
      if (!btn) return;
      var card = btn.closest(".intake-card");
      if (!card) return;
      var idx = parseInt(card.dataset.intakeIdx);
      var visible = filteredIntakeClusters();
      var cluster = visible[idx];
      if (!cluster) return;
      var action = btn.dataset.action;
      if (action === "add-artifact") intakeAddToArtifacts(cluster);
      else if (action === "add-reel") intakeAddToReel(cluster);
      else if (action === "dismiss") intakeDismissCluster(cluster);
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
    if (searchEl) {
      searchEl.addEventListener("input", function () {
        state.intakeFilterText = this.value;
        renderIntake(false);
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
    initFilterToggle();
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
    initIntake();
    window.addEventListener("resize", computeGridMaxHeight);

    document.addEventListener("dragstart", function (ev) {
      if (!ev.target.closest(".stash-card")) return;
      requestAnimationFrame(revealEmptyStashAreas);
    });
    document.addEventListener("dragend", function () {
      hideEmptyStashAreas();
    });
  });
})();
