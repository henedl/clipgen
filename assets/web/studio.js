/* clipgen Studio */

(function () {
  "use strict";

  var THEME_STORAGE_KEY = "clipgen-studio-theme";

  var state = {
    sheetData: null,
    artifactQueue: [],
    reelQueue: [],
    generatedArtifacts: [],
    generating: false,
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

  // ---- Data loading ----

  function loadSheetData() {
    fetch("/api/sheet")
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
      })
      .catch(function (err) {
        qs("#sheetLoading").textContent = "Failed to load sheet: " + err;
      });
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
    var idx = findInQueue(state.artifactQueue, info.participant, info.row);
    if (idx >= 0) {
      state.artifactQueue.splice(idx, 1);
    } else {
      state.artifactQueue.push(info);
    }
    renderArtifactQueue();
    updateCellClasses();
  }

  function toggleReelCell(info) {
    var idx = findInQueue(state.reelQueue, info.participant, info.row);
    if (idx >= 0) {
      state.reelQueue.splice(idx, 1);
    } else {
      state.reelQueue.push(info);
    }
    renderReelQueue();
    updateCellClasses();
  }

  function addToQueue(targetQueue, info, renderFn) {
    if (findInQueue(targetQueue, info.participant, info.row) < 0) {
      targetQueue.push(info);
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
    // If all are already in the queue, remove them; otherwise add missing ones
    var allPresent = true;
    for (var i = 0; i < infos.length; i++) {
      if (findInQueue(queue, infos[i].participant, infos[i].row) < 0) {
        allPresent = false;
        break;
      }
    }
    if (allPresent) {
      for (var j = 0; j < infos.length; j++) {
        var idx = findInQueue(queue, infos[j].participant, infos[j].row);
        if (idx >= 0) queue.splice(idx, 1);
      }
    } else {
      for (var k = 0; k < infos.length; k++) {
        if (findInQueue(queue, infos[k].participant, infos[k].row) < 0) {
          queue.push(infos[k]);
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

  function initDropTargets() {
    setupDropTarget(qs("#artifactsList"), function (info) {
      addToQueue(state.artifactQueue, info, renderArtifactQueue);
    });
    setupDropTarget(qs("#reelList"), function (info) {
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
      var item = ev.target.closest(".queue-item[data-reel-idx]");
      if (!item) return;
      _reelDragIdx = parseInt(item.getAttribute("data-reel-idx"), 10);
      ev.dataTransfer.effectAllowed = "move";
      ev.dataTransfer.setData("text/plain", String(_reelDragIdx));
    });

    list.addEventListener("dragover", function (ev) {
      var item = ev.target.closest(".queue-item[data-reel-idx]");
      if (item && _reelDragIdx !== null) {
        ev.preventDefault();
        ev.dataTransfer.dropEffect = "move";
      }
    });

    list.addEventListener("drop", function (ev) {
      var item = ev.target.closest(".queue-item[data-reel-idx]");
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
    var list = qs("#artifactsList");
    qs("#artifactsCount").textContent = "(" + state.artifactQueue.length + ")";
    qs("#generateBtn").disabled = state.artifactQueue.length === 0;
    list.innerHTML = "";

    if (state.artifactQueue.length === 0) {
      list.appendChild(
        el("div", "drop-target-empty", "Click or drag cells here to queue for generation")
      );
      return;
    }

    for (var i = 0; i < state.artifactQueue.length; i++) {
      var item = state.artifactQueue[i];
      list.appendChild(makeQueueItem(item, i, "artifact"));
    }
  }

  function renderReelQueue() {
    var list = qs("#reelList");
    qs("#reelCount").textContent = "(" + state.reelQueue.length + ")";
    qs("#buildReelBtn").disabled = state.reelQueue.length === 0;
    list.innerHTML = "";

    if (state.reelQueue.length === 0) {
      list.appendChild(
        el("div", "drop-target-empty", "Shift+click or drag cells here to build a reel")
      );
      return;
    }

    for (var i = 0; i < state.reelQueue.length; i++) {
      var item = state.reelQueue[i];
      var row = makeQueueItem(item, i, "reel");
      row.setAttribute("data-reel-idx", i);
      row.setAttribute("draggable", "true");

      var handle = el("span", "reel-handle", "\u2261");
      var order = el("span", "reel-order", String(i + 1));
      row.insertBefore(order, row.firstChild);
      row.insertBefore(handle, row.firstChild);
      list.appendChild(row);
    }
  }

  function makeQueueItem(item, idx, type) {
    var div = el("div", "queue-item");
    div.appendChild(
      el("span", "queue-item-ref", item.participant + "." + item.row)
    );
    div.appendChild(el("span", "queue-item-desc", truncate(item.desc, 40)));
    div.appendChild(el("span", "queue-item-ts", item.timestamp));

    var removeBtn = el("button", "queue-item-remove", "\u00D7");
    removeBtn.title = "Remove";
    removeBtn.addEventListener("click", function (ev) {
      ev.stopPropagation();
      if (type === "artifact") {
        state.artifactQueue.splice(idx, 1);
        renderArtifactQueue();
      } else {
        state.reelQueue.splice(idx, 1);
        renderReelQueue();
      }
      updateCellClasses();
    });
    div.appendChild(removeBtn);

    return div;
  }

  // ---- Buttons ----

  function bindButtons() {
    qs("#clearArtifactsBtn").addEventListener("click", function () {
      state.artifactQueue = [];
      renderArtifactQueue();
      updateCellClasses();
    });

    qs("#clearReelBtn").addEventListener("click", function () {
      state.reelQueue = [];
      renderReelQueue();
      updateCellClasses();
    });

    qs("#generateBtn").addEventListener("click", onGenerate);
    qs("#buildReelBtn").addEventListener("click", onBuildReel);
    qs("#buildViewerBtn").addEventListener("click", onBuildViewer);
    qs("#buildTimelineViewerBtn").addEventListener("click", onBuildTimelineViewer);

    qs("#statusDismiss").addEventListener("click", hideOverlay);
  }

  // ---- API calls ----

  function onGenerate() {
    if (state.generating || state.artifactQueue.length === 0) return;
    state.generating = true;

    var cells = state.artifactQueue.map(function (item) {
      return item.participant + "." + item.row;
    });
    var format = qs("#artifactFormat").value;

    showOverlay("Generating " + cells.length + " " + format + "(s)...");

    fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ cells: cells, format: format }),
    })
      .then(function (r) {
        return r.json();
      })
      .then(function (data) {
        state.generating = false;
        if (data.ok) {
          if (data.artifacts) {
            state.generatedArtifacts = state.generatedArtifacts.concat(
              data.artifacts
            );
          }
          updateViewerButton();
          showResult(
            "Generated " + (data.generated || 0) + " artifact(s)",
            null
          );
        } else {
          showResult(null, data.error || "Generation failed");
        }
      })
      .catch(function (err) {
        state.generating = false;
        showResult(null, "Request failed: " + err);
      });
  }

  function onBuildReel() {
    if (state.generating || state.reelQueue.length === 0) return;
    state.generating = true;

    var cells = state.reelQueue.map(function (item) {
      return item.participant + "." + item.row;
    });

    showOverlay("Building reel from " + cells.length + " clips...");

    fetch("/api/reel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ cells: cells }),
    })
      .then(function (r) {
        return r.json();
      })
      .then(function (data) {
        state.generating = false;
        if (data.ok) {
          showResult("Reel built successfully", null);
        } else {
          showResult(null, data.error || "Reel build failed");
        }
      })
      .catch(function (err) {
        state.generating = false;
        showResult(null, "Request failed: " + err);
      });
  }

  function onBuildViewer() {
    if (state.generating || state.generatedArtifacts.length === 0) return;
    state.generating = true;

    showOverlay("Building timeline viewer...");

    fetch("/api/viewer", {
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

    fetch("/api/timeline-viewer", {
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

  document.addEventListener("DOMContentLoaded", function () {
    initThemeToggle();
    initDropTargets();
    bindReelReorder();
    bindButtons();
    loadSheetData();
    updateViewerButton();
  });
})();
