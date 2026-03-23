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
        renderGrid();
        var durInput = qs("#highlightsDuration");
        if (durInput && data.highlightsDuration) {
          durInput.value = data.highlightsDuration;
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
          artifacts.filter(function (a) { return a.type !== "transcript"; })
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
    var metaCols = showSeverity ? 4 : 3;
    var totalCols = metaCols + d.participants.length;
    var table = el("table");

    // Colgroup for fixed column widths — participant columns share equal width
    var colgroup = document.createElement("colgroup");
    var colRowNum = document.createElement("col");
    colRowNum.style.width = "3rem";
    colgroup.appendChild(colRowNum);
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
    var tbody = el("tbody");
    var i = 0;
    while (i < d.rows.length) {
      var row = d.rows[i];
      if (isRowEmpty(row, d.participants)) {
        var emptyStart = i;
        while (i < d.rows.length && isRowEmpty(d.rows[i], d.participants)) {
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
  }

  function renderDataRow(row, participants, showSeverity) {
    var tr = el("tr");

    var rowTd = el("td", "col-row-num col-row-num-clickable", String(row.rowNum));
    rowTd.setAttribute("data-select-row", row.rowNum);
    rowTd.title = "Select row " + row.rowNum;
    tr.appendChild(rowTd);
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
        td.title = cellData.value;
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
      if (info.source === "reel") {
        removeFromQueue(state.reelQueue, info);
        renderReelQueue();
      }
      addToQueue(state.artifactQueue, info, renderArtifactQueue);
    });
    setupDropTarget(qs("#reelList"), function (info) {
      if (info.source === "artifact") {
        removeFromQueue(state.artifactQueue, info);
        renderArtifactQueue();
      }
      addToQueue(state.reelQueue, info, renderReelQueue);
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
        if (info && info.participant && info.row) {
          onDrop(info);
        }
      } catch (_) {}
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
    qs("#addToReelBtn").disabled = n === 0;
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
      var segStart = item.segStart !== undefined ? item.segStart : parseClipTimestamps(item.timestamp)[0].startSeconds;
      var segDuration = item.segDuration !== undefined ? item.segDuration : parseClipTimestamps(item.timestamp)[0].duration;
      var segTotal = item.segTotal || 1;
      var segIdx = item.segIdx || 0;

      var card = el("div", "artifact-card");
      card.setAttribute("data-participant", item.participant);
      card.setAttribute("data-row", item.row);
      card.setAttribute("data-seg-idx", segIdx);
      card.setAttribute("draggable", "true");
      (function (itm) {
        card.addEventListener("dragstart", function (ev) {
          var data = {
            participant: itm.participant,
            row: itm.row,
            desc: itm.desc,
            timestamp: itm.timestamp,
            segIdx: itm.segIdx,
            segStart: itm.segStart,
            segDuration: itm.segDuration,
            segTotal: itm.segTotal,
            source: "artifact",
          };
          ev.dataTransfer.setData("application/json", JSON.stringify(data));
          ev.dataTransfer.effectAllowed = "copyMove";
        });
      })(item);

      var thumb = el("div", "artifact-card-thumb");
      var img = document.createElement("img");
      img.src = "api/thumbnail/" + encodeURIComponent(item.participant) + "/" + segStart;
      img.loading = "lazy";
      img.alt = "";
      img.draggable = false;
      (function (cardEl, thumbEl) {
        img.addEventListener("error", function () {
          this.remove();
          thumbEl.appendChild(el("span", "", "\u2715"));
          cardEl.classList.add("artifact-card-error");
        });
      })(card, thumb);
      thumb.appendChild(img);
      thumb.appendChild(el("span", "artifact-card-duration", formatDuration(segDuration)));
      card.appendChild(thumb);

      var meta = el("div", "artifact-card-meta");
      var refText = item.participant + "." + item.row;
      if (segTotal > 1) refText += " (" + (segIdx + 1) + "/" + segTotal + ")";
      meta.appendChild(el("span", "artifact-card-ref", refText));
      card.appendChild(meta);

      var removeBtn = el("button", "artifact-card-remove", "\u00D7");
      removeBtn.title = "Remove";
      (function (idx) {
        removeBtn.addEventListener("click", function (ev) {
          ev.stopPropagation();
          var removed = state.artifactQueue.splice(idx, 1)[0];
          delete state.cellResults[cellKey(removed.participant, removed.row)];
          renderArtifactQueue();
          updateCellClasses();
        });
      })(i);
      card.appendChild(removeBtn);

      (function (p, r) {
        card.addEventListener("mouseenter", function () { highlightGridHeaders(p, r); });
        card.addEventListener("mouseleave", clearGridHighlights);
      })(item.participant, item.row);

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
      var segStart = item.segStart !== undefined ? item.segStart : parseClipTimestamps(item.timestamp)[0].startSeconds;
      var segDuration = item.segDuration !== undefined ? item.segDuration : parseClipTimestamps(item.timestamp)[0].duration;
      var segTotal = item.segTotal || 1;
      var segIdx = item.segIdx || 0;
      totalDur += segDuration;

      var card = el("div", "reel-card");
      card.setAttribute("data-reel-idx", i);
      card.setAttribute("data-participant", item.participant);
      card.setAttribute("data-row", item.row);
      card.setAttribute("data-seg-idx", segIdx);
      card.setAttribute("draggable", "true");

      var thumb = el("div", "reel-card-thumb");
      var img = document.createElement("img");
      img.src = "api/thumbnail/" + encodeURIComponent(item.participant) + "/" + segStart;
      img.loading = "lazy";
      img.alt = "";
      img.draggable = false;
      (function (cardEl, thumbEl) {
        img.addEventListener("error", function () {
          this.remove();
          thumbEl.appendChild(el("span", "", "\u2715"));
          cardEl.classList.add("reel-card-error");
        });
      })(card, thumb);
      thumb.appendChild(img);
      thumb.appendChild(el("span", "reel-card-duration", formatDuration(segDuration)));
      card.appendChild(thumb);

      var meta = el("div", "reel-card-meta");
      meta.appendChild(el("span", "reel-card-order", String(i + 1)));
      var refText = item.participant + "." + item.row;
      if (segTotal > 1) refText += " (" + (segIdx + 1) + "/" + segTotal + ")";
      meta.appendChild(el("span", "reel-card-ref", refText));
      card.appendChild(meta);

      var removeBtn = el("button", "reel-card-remove", "\u00D7");
      removeBtn.title = "Remove";
      (function (idx) {
        removeBtn.addEventListener("click", function (ev) {
          ev.stopPropagation();
          var removed = state.reelQueue.splice(idx, 1)[0];
          delete state.cellResults[cellKey(removed.participant, removed.row)];
          renderReelQueue();
          updateCellClasses();
        });
      })(i);
      card.appendChild(removeBtn);

      (function (p, r) {
        card.addEventListener("mouseenter", function () { highlightGridHeaders(p, r); });
        card.addEventListener("mouseleave", clearGridHighlights);
      })(item.participant, item.row);

      list.appendChild(card);
    }
    qs("#reelDuration").textContent = formatDuration(totalDur);
    applyCardStates(list);
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

    qs("#generateBtn").addEventListener("click", onGenerate);
    qs("#buildReelBtn").addEventListener("click", onBuildReel);
    qs("#buildViewerBtn").addEventListener("click", onBuildViewer);
    qs("#buildTimelineViewerBtn").addEventListener("click", onBuildTimelineViewer);
    qs("#buildHighlightsBtn").addEventListener("click", onBuildHighlights);

    qs("#statusDismiss").addEventListener("click", hideOverlay);
  }

  // ---- Generation progress helpers ----

  function setTitleSpinner(id, active) {
    var s = qs("#" + id);
    if (s) s.classList.toggle("active", active);
  }

  function findCard(listEl, participant, row) {
    return listEl.querySelector(
      '[data-participant="' + participant + '"][data-row="' + row + '"]'
    );
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
    card.classList.add(
      card.classList.contains("reel-card") ? "reel-card-queued" : "artifact-card-queued"
    );
    var thumb = card.querySelector(".artifact-card-thumb, .reel-card-thumb");
    if (thumb) thumb.appendChild(createPulserOverlay());
    var p = card.getAttribute("data-participant");
    var r = card.getAttribute("data-row");
    if (p && r) delete state.cellResults[cellKey(p, parseInt(r, 10))];
  }

  function setCardResult(card, success) {
    var isReel = card.classList.contains("reel-card");
    card.classList.remove(isReel ? "reel-card-queued" : "artifact-card-queued");
    card.classList.add(
      isReel
        ? (success ? "reel-card-success" : "reel-card-fail")
        : (success ? "artifact-card-success" : "artifact-card-fail")
    );
    var overlay = card.querySelector(".card-gen-overlay");
    if (overlay) overlay.remove();
    var thumb = card.querySelector(".artifact-card-thumb, .reel-card-thumb");
    if (thumb) thumb.appendChild(createResultBadge(success));
    var p = card.getAttribute("data-participant");
    var r = card.getAttribute("data-row");
    if (p && r) state.cellResults[cellKey(p, parseInt(r, 10))] = success ? "success" : "fail";
  }

  function applyCardStates(listEl) {
    var cards = listEl.querySelectorAll(".artifact-card, .reel-card");
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

  function clearCardStates(listEl) {
    var cards = listEl.querySelectorAll(
      ".artifact-card-queued, .artifact-card-success, .artifact-card-fail," +
      ".reel-card-queued, .reel-card-success, .reel-card-fail"
    );
    for (var i = 0; i < cards.length; i++) {
      cards[i].classList.remove(
        "artifact-card-queued", "artifact-card-success", "artifact-card-fail",
        "reel-card-queued", "reel-card-success", "reel-card-fail"
      );
      var overlay = cards[i].querySelector(".card-gen-overlay");
      if (overlay) overlay.remove();
      var badge = cards[i].querySelector(".card-gen-badge");
      if (badge) badge.remove();
    }
  }

  function setGeneratingLock(locked) {
    var ids = [
      "#generateBtn", "#buildReelBtn", "#clearArtifactsBtn",
      "#clearReelBtn", "#addToReelBtn", "#buildHighlightsBtn",
      "#buildViewerBtn", "#buildTimelineViewerBtn"
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
    var cellsSeen = {};
    var cells = [];
    for (var ci = 0; ci < items.length; ci++) {
      var ck = items[ci].participant + "." + items[ci].row;
      if (!cellsSeen[ck]) { cellsSeen[ck] = true; cells.push(ck); }
    }

    var allCards = list.querySelectorAll(".artifact-card");
    for (var i = 0; i < allCards.length; i++) {
      setCardQueued(allCards[i]);
    }

    var totalSuccess = 0;
    var totalFail = 0;
    var allArtifacts = [];

    function finishGenerate() {
      setTitleSpinner("artifactsSpinner", false);
      state.generating = false;
      setGeneratingLock(false);
      if (allArtifacts.length > 0) {
        state.generatedArtifacts = state.generatedArtifacts.concat(allArtifacts);
        updateViewerButton();
      }
      var msg = "Generated " + totalSuccess + " artifact(s)";
      if (totalFail > 0) msg += ", " + totalFail + " failed";
      showResult(
        totalSuccess > 0 ? msg : null,
        totalSuccess === 0 && totalFail > 0 ? "All generations failed" : null
      );
      qs("#statusOverlay").classList.remove("hidden");
    }

    function handleLine(line) {
      var data = JSON.parse(line);
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

    fetch("api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ cells: cells, format: format }),
    })
      .then(function (response) {
        var reader = response.body.getReader();
        var decoder = new TextDecoder();
        var buffer = "";

        function read() {
          return reader.read().then(function (result) {
            if (result.done) {
              if (buffer.trim()) handleLine(buffer.trim());
              finishGenerate();
              return;
            }
            buffer += decoder.decode(result.value, { stream: true });
            var lines = buffer.split("\n");
            buffer = lines.pop();
            for (var i = 0; i < lines.length; i++) {
              if (lines[i].trim()) handleLine(lines[i].trim());
            }
            return read();
          });
        }

        return read();
      })
      .catch(function (err) {
        finishGenerate();
      });
  }

  function onBuildReel() {
    if (state.generating || state.reelQueue.length === 0) return;
    state.generating = true;
    setGeneratingLock(true);
    setTitleSpinner("reelSpinner", true);

    var cellsSeen = {};
    var cells = [];
    for (var ci = 0; ci < state.reelQueue.length; ci++) {
      var ck = state.reelQueue[ci].participant + "." + state.reelQueue[ci].row;
      if (!cellsSeen[ck]) { cellsSeen[ck] = true; cells.push(ck); }
    }

    var list = qs("#reelList");
    var reelCards = list.querySelectorAll(".reel-card");
    for (var i = 0; i < reelCards.length; i++) {
      setCardQueued(reelCards[i]);
    }

    fetch("api/reel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ cells: cells }),
    })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        state.generating = false;
        setTitleSpinner("reelSpinner", false);
        setGeneratingLock(false);

        var cards = list.querySelectorAll(".reel-card");
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

        var cards = list.querySelectorAll(".reel-card");
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

  function onBuildHighlights() {
    if (state.generating) return;

    var drawer = qs("#highlightsDurationDrawer");
    var btn = qs("#buildHighlightsBtn");
    var isOpen = drawer.classList.contains("open");

    var sparklesHTML =
      '<svg class="sparkles-icon" width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">' +
      '<path d="M8 1l1.1 3.4L12.5 5.5l-3.4 1.1L8 10l-1.1-3.4L3.5 5.5l3.4-1.1z"/>' +
      '<path d="M12.5 10l.6 1.7 1.7.6-1.7.6-.6 1.7-.6-1.7-1.7-.6 1.7-.6z"/>' +
      '<path d="M3 11l.4 1.1 1.1.4-1.1.4L3 14l-.4-1.1L1.5 12.5l1.1-.4z"/>' +
      "</svg>";
    var checkHTML =
      '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
      '<path d="M3 8.5l3.5 3.5 6.5-8"/>' +
      "</svg>";

    if (!isOpen) {
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
    btn.innerHTML = sparklesHTML + "Find Highlights";

    state.generating = true;
    showOverlay("Finding best clips (" + duration + "s budget)...");

    fetch("api/highlights-preview", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ highlights_duration: duration }),
    })
      .then(function (r) {
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

  function updateViewerButton() {
    var n = state.generatedArtifacts.length;
    qs("#buildViewerBtn").disabled = n === 0;
    var count = qs("#viewerArtifactCount");
    count.textContent = n > 0 ? n + " artifact(s) ready" : "";
  }

  // ---- Status overlay ----

  function showOverlay(message) {
    qs("#statusSpinner").style.display = "";
    qs("#statusTitle").textContent = message;
    qs("#statusMessage").textContent = "";
    qs("#statusMessage").className = "";
    qs("#statusDismiss").style.display = "none";
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
    qs("#statusDismiss").style.display = "";
  }

  function hideOverlay() {
    qs("#statusOverlay").classList.add("hidden");
  }

  // ---- Init ----

  function checkNavLinks() {
    fetch("/api/status")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data.insights) {
          var link = qs("#insightsLink");
          if (link) link.classList.remove("hidden");
        }
      })
      .catch(function () {});
  }

  document.addEventListener("DOMContentLoaded", function () {
    initThemeToggle();
    initDropTargets();
    bindReelReorder();
    bindButtons();
    loadSheetData();
    updateViewerButton();
    checkNavLinks();
  });
})();
