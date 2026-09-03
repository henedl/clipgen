/* clipgen Screenspace tasks satellite — screenspace-tasks.js
 *
 * The task lifecycle surface carved out of screenspace.js: queue sort/filter,
 * right-pane tabs + results switcher, the draggable task list, task -> workflow
 * restore, the pause control, SSE/polling intake, and per-task ETA ticking.
 * Loaded after screenspace.js (and before the other satellites, which destructure
 * findTask / restoreTaskToWorkflow / setInputValue / syncValueDisplays off SS at
 * load time). Reads the hub's shared state + helpers through
 * window.ClipgenScreenspace and publishes its entry points back onto it; the hub
 * keeps same-named thin delegators so its own call sites are unchanged. Function
 * bodies are unchanged from when they lived inline in screenspace.js — the locals
 * below stand in for the closure. renderResults / loadAndShowResults (results
 * surface) and setTargetColor (color satellite) are reached via SS.* late-binding
 * because their owners load/register after this file.
 */
(function () {
  "use strict";

  var SS = window.ClipgenScreenspace;
  var state = SS.state;
  var _colorMode = SS._colorMode,
    _updateMinAreaReadout = SS._updateMinAreaReadout,
    activeRegionRef = SS.activeRegionRef,
    applyColorMode = SS.applyColorMode,
    applyNormalizeMode = SS.applyNormalizeMode,
    iconSpan = SS.iconSpan,
    normalizeRegionRef = SS.normalizeRegionRef,
    refreshCalibration = SS.refreshCalibration,
    renderOverlay = SS.renderOverlay,
    renderRegionChips = SS.renderRegionChips,
    renderRunRegionPicker = SS.renderRunRegionPicker,
    renderTimeline = SS.renderTimeline,
    renderWorkflowParams = SS.renderWorkflowParams,
    selectParticipant = SS.selectParticipant,
    taskTypeColor = SS.taskTypeColor,
    updateMarkerInfo = SS.updateMarkerInfo,
    updateParamResetButtons = SS.updateParamResetButtons,
    updateRegionButtons = SS.updateRegionButtons,
    updateRunButton = SS.updateRunButton;

  // ---- ETA tracking + poll fingerprint ----
  // Structural fingerprint (id:status:heatmap); handleTaskData owns the rebuild gate.
  var _lastTaskFp = "";
  // Per-task ETA trackers by id; progress is linear in scanned duration.
  var _etaTrackers = {};
  var _etaTicker = createIntervalTicker(tickEtas, {
    gateHidden: true,
    isActive: function () {
      return state.tasks.some(taskIsActive);
    },
  });

  var TASK_TYPE_ICON_FILES = {
    multitool: "link",
    color: "eye-dropper",
    change: "bolt",
    similarity: "photo",
    text: "language",
    numbers: "hashtag",
    timelapse: "forward",
    template: "viewfinder-circle",
    shape: "star",
    flow: "arrows-right-left",
    scene: "squares-2x2",
    inactivity: "pause-circle",
    boundary: "flag",
    attention: "eye",
  };

  function sortTasks() {
    // completed/failed at top (oldest first), then running, then queued (by priority), cancelled last
    var statusOrder = { completed: 0, failed: 1, running: 2, paused: 3, queued: 4, cancelled: 5 };
    state.tasks.sort(function (a, b) {
      var sa = statusOrder[a.status] !== undefined ? statusOrder[a.status] : 5;
      var sb = statusOrder[b.status] !== undefined ? statusOrder[b.status] : 5;
      if (sa !== sb) return sa - sb;
      if (a.status === "queued" && b.status === "queued") {
        return (a.priority || 100) - (b.priority || 100);
      }
      return (a.created_at || "").localeCompare(b.created_at || "");
    });
  }

  var TOOL_LABELS = {
    multitool: "Multitool",
    color: "Color",
    change: "Change",
    similarity: "Similarity",
    text: "Text",
    numbers: "Numbers",
    timelapse: "Timelapse",
    template: "Template",
    shape: "Shape",
    flow: "Flow",
    scene: "Scene",
    inactivity: "Inactivity",
    boundary: "Boundary",
    attention: "Attention",
  };

  function selectableTasks() {
    return state.tasks.filter(function (t) {
      return t.status === "completed" || t.status === "paused" || t.status === "running";
    });
  }

  function setRightPaneTab(tab) {
    state.rightPaneTab = tab;
    setStoredUIStateField("screenspace", "rightPaneTab", tab);
    qsa("#rightPaneTabs .rp-tab").forEach(function (b) {
      b.classList.toggle("active", b.dataset.tab === tab);
    });
    qsa("#rightPaneBody .rp-tab-panel").forEach(function (p) {
      p.classList.toggle("hidden", p.dataset.tab !== tab);
    });
    qsa("#rightPaneTabs .rp-tab-actions").forEach(function (a) {
      a.classList.toggle("hidden", a.dataset.for !== tab);
    });
    closeResultsSwitcher();
    // Hidden preview tabs skip work; entering refetches. calRefresh is late-bound (calibration loads later).
    if (tab === "preview") {
      SS.refreshModelView();
      if (SS.calRefresh) SS.calRefresh();
    }
  }

  function updateResultsCrumb() {
    var crumbEl = qs("#resultsTabCrumb");
    if (!crumbEl) return;
    var task = state.selectedTaskId ? findTask(state.selectedTaskId) : null;
    if (!task) {
      crumbEl.textContent = "";
      crumbEl.style.color = "";
      return;
    }
    var label = (task.parameters || {}).event_label || task.name;
    if (!label) {
      // Unnamed tasks from old manifests fall back to "Color 3" style labels.
      var sameType = state.tasks
        .filter(function (t) {
          return t.type === task.type &&
            (t.status === "completed" || t.status === "paused" || t.status === "running");
        });
      var idx = -1;
      for (var i = 0; i < sameType.length; i++) {
        if (sameType[i].id === task.id) { idx = i; break; }
      }
      label = (TOOL_LABELS[task.type] || task.type) + " " + (idx >= 0 ? idx + 1 : 1);
    }
    var participant = task.participant || "";
    crumbEl.textContent = ": " + label + (participant ? " \u00b7 " + participant : "");
    crumbEl.style.color = taskTypeColor(task.type);
  }

  function openResultsSwitcher() {
    var panel = qs("#resultsSwitcherPanel");
    if (!panel) return;
    panel.innerHTML = "";
    var tasks = selectableTasks();
    if (tasks.length === 0) {
      var empty = el("div", "rp-switcher-empty", "No completed tasks yet.");
      panel.appendChild(empty);
    } else {
      var frag = document.createDocumentFragment();
      tasks.forEach(function (t) {
        var item = el("button", "rp-switcher-item");
        item.type = "button";
        item.dataset.taskId = t.id;
        if (t.id === state.selectedTaskId) item.classList.add("active");
        var badge = el("span", "rp-switcher-item-badge");
        badge.style.background = taskTypeColor(t.type);
        item.appendChild(badge);
        var label = (t.parameters || {}).event_label || t.name || TOOL_LABELS[t.type] || t.type;
        var primary = el("span", null, label + " \u00b7 " + (t.participant || ""));
        item.appendChild(primary);
        if (t.region) {
          var meta = el("span", "rp-switcher-item-meta", t.region);
          item.appendChild(meta);
        }
        item.addEventListener("click", function (e) {
          e.stopPropagation();
          var taskId = t.id;
          closeResultsSwitcher();
          var task = findTask(taskId);
          if (!task) return;
          if (task.participant && task.participant !== state.selectedParticipant) {
            selectParticipant(task.participant);
          }
          state.selectedTaskId = taskId;
          setRightPaneTab("results");
          SS.loadAndShowResults(taskId);
          renderTaskList();
        });
        frag.appendChild(item);
      });
      panel.appendChild(frag);
    }
    var tab = qs('.rp-tab[data-tab="results"]');
    var tabsEl = qs("#rightPaneTabs");
    if (tab && tabsEl) {
      var tabRect = tab.getBoundingClientRect();
      var tabsRect = tabsEl.getBoundingClientRect();
      var tabCenter = tabRect.left - tabsRect.left + tabRect.width / 2;
      panel.style.left = tabCenter + "px";
      panel.style.transform = "translateX(-50%)";
    }
    panel.classList.remove("hidden");
    state.resultsSwitcherOpen = true;
    if (tab) tab.classList.add("switcher-open");
  }

  function closeResultsSwitcher() {
    var panel = qs("#resultsSwitcherPanel");
    if (panel) panel.classList.add("hidden");
    state.resultsSwitcherOpen = false;
    var tab = qs('.rp-tab[data-tab="results"]');
    if (tab) tab.classList.remove("switcher-open");
  }

  function initRightPaneTabs() {
    qsa("#rightPaneTabs .rp-tab").forEach(function (btn) {
      btn.addEventListener("click", function (e) {
        e.stopPropagation();
        var tab = btn.dataset.tab;
        if (tab === "queue" || tab === "preview") {
          setRightPaneTab(tab);
          return;
        }
        if (tab === "results") {
          if (state.rightPaneTab === "results" && state.selectedTaskId) {
            if (state.resultsSwitcherOpen) closeResultsSwitcher();
            else openResultsSwitcher();
          } else {
            setRightPaneTab("results");
          }
        }
      });
    });
    document.addEventListener("click", function (e) {
      if (!state.resultsSwitcherOpen) return;
      if (e.target.closest("#resultsSwitcherPanel")) return;
      if (e.target.closest('.rp-tab[data-tab="results"]')) return;
      closeResultsSwitcher();
    });
  }

  function initTaskQueue() {
    var taskListEl = qs("#taskList");

    // Click handler delegated on taskList
    taskListEl.addEventListener("click", function (e) {
      var card = e.target.closest(".task-card");
      if (!card) return;
      var taskId = card.dataset.taskId;

      // Dismiss button
      if (e.target.closest(".task-card-dismiss")) {
        apiDelete("api/tasks/" + taskId + "?dismiss=true")
          .then(function (data) {
            if (data.ok) {
              state.tasks = state.tasks.filter(function (t) { return t.id !== taskId; });
              if (state.hoveredTaskId === taskId) {
                state.hoveredTaskId = null;
              }
              if (state.selectedTaskId === taskId) {
                state.selectedTaskId = null;
                state.selectedTaskResults = null;
                SS.renderResults();
                setRightPaneTab("queue");
              }
              renderTaskList();
              renderTimeline();
              showToast("Task dismissed");
            }
          })
          .catch(function () { showToast("Failed to dismiss task"); });
        return;
      }

      // Edit button
      if (e.target.closest(".task-card-edit")) {
        var task = findTask(taskId);
        if (task) restoreTaskToWorkflow(task);
        return;
      }

      // Select completed/paused/running task to view results; click again to deselect
      task = findTask(taskId);
      if (task && (task.status === "completed" || task.status === "paused" || task.status === "running")) {
        if (state.selectedTaskId === taskId) {
          state.resultsRequestVersion += 1;
          state.selectedTaskId = null;
          state.selectedTaskResults = null;
          SS.renderResults();
          renderTaskList();
          renderTimeline();
          updateResultsCrumb();
          setRightPaneTab("queue");
        } else {
          if (task.participant && task.participant !== state.selectedParticipant) {
            selectParticipant(task.participant);
          }
          state.selectedTaskId = taskId;
          setRightPaneTab("results");
          SS.loadAndShowResults(taskId);
          renderTaskList();
        }
      }
    });

    // Hover handler for task focus (dim non-hovered timeline markers)
    taskListEl.addEventListener("mouseover", function (e) {
      var card = e.target.closest(".task-card");
      if (!card) return;
      var task = findTask(card.dataset.taskId);
      if (task && (task.status === "completed" || task.status === "running")) {
        if (state.hoveredTaskId !== task.id) {
          state.hoveredTaskId = task.id;
          renderTimeline();
        }
      } else if (state.hoveredTaskId) {
        state.hoveredTaskId = null;
        renderTimeline();
      }
    });

    taskListEl.addEventListener("mouseleave", function () {
      if (state.hoveredTaskId) {
        state.hoveredTaskId = null;
        renderTimeline();
      }
    });

    // Drag-and-drop: only initiate drag from the handle
    taskListEl.addEventListener("dragstart", function (e) {
      var card = e.target.closest(".task-card");
      if (!card) { e.preventDefault(); return; }
      var task = findTask(card.dataset.taskId);
      if (!task) { e.preventDefault(); return; }
      var allowed = task.status === "queued" || task.status === "completed" || task.status === "failed";
      if (!allowed) { e.preventDefault(); return; }
      card.classList.add("dragging");
      e.dataTransfer.setData("text/plain", card.dataset.taskId);
      e.dataTransfer.setData("application/x-task-status", task.status);
      e.dataTransfer.setData("application/x-task-id", card.dataset.taskId);
      e.dataTransfer.effectAllowed = "move";
      _cacheTaskDragMidpoints(taskListEl);
    });

    taskListEl.addEventListener("dragend", function (e) {
      var card = e.target.closest(".task-card");
      if (card) {
        card.classList.remove("dragging");
        card.removeAttribute("draggable");
      }
      if (_taskListDragOverRaf != null) {
        cancelAnimationFrame(_taskListDragOverRaf);
        _taskListDragOverRaf = null;
      }
      _taskListPendingDragOver = null;
      clearDragIndicators(taskListEl);
      _taskDragCache = null;
    });

    taskListEl.addEventListener("dragover", function (e) {
      if (e.dataTransfer.types.indexOf("text/plain") < 0) return;

      var cards = taskListEl.querySelectorAll(".task-card:not(.dragging)");
      var insertIdx = getDropIndex(taskListEl, e.clientY);

      // Determine boundary between finished (completed/failed) and queued zones
      var finishedCount = 0;
      for (var i = 0; i < cards.length; i++) {
        var t = findTask(cards[i].dataset.taskId);
        if (t && (t.status === "completed" || t.status === "failed")) finishedCount++;
        else break;
      }

      // dataTransfer payloads are unreadable during dragover; read status off the card.
      var draggingCard = taskListEl.querySelector(".task-card.dragging");
      var draggingTask = draggingCard ? findTask(draggingCard.dataset.taskId) : null;
      var isQueuedDrag = draggingTask && draggingTask.status === "queued";
      var isFinishedDrag = draggingTask && (draggingTask.status === "completed" || draggingTask.status === "failed");

      // Queued tasks can't go above finished tasks
      if (isQueuedDrag && insertIdx < finishedCount) return;
      // Finished tasks can't go below into queued zone
      if (isFinishedDrag && insertIdx > finishedCount) return;

      e.preventDefault();
      e.dataTransfer.dropEffect = "move";

      _taskListPendingDragOver = { insertIdx: insertIdx };
      if (_taskListDragOverRaf != null) return;
      _taskListDragOverRaf = requestAnimationFrame(function () {
        _taskListDragOverRaf = null;
        var pending = _taskListPendingDragOver;
        if (!pending) return;
        var idx = pending.insertIdx;
        var cardsNow = taskListEl.querySelectorAll(".task-card:not(.dragging)");
        clearDragIndicators(taskListEl);
        if (idx < cardsNow.length) {
          cardsNow[idx].classList.add("drag-over");
        } else {
          taskListEl.classList.add("drag-over-append");
        }
      });
    });

    taskListEl.addEventListener("dragleave", function (e) {
      var card = e.target.closest(".task-card");
      if (card) card.classList.remove("drag-over");
      if (!taskListEl.contains(e.relatedTarget)) {
        taskListEl.classList.remove("drag-over-append");
      }
    });

    taskListEl.addEventListener("drop", function (e) {
      e.preventDefault();
      clearDragIndicators(taskListEl);
      var draggedId = e.dataTransfer.getData("text/plain");
      if (!draggedId) return;

      var draggedTask = findTask(draggedId);
      if (!draggedTask) return;
      var isQueued = draggedTask.status === "queued";

      if (isQueued) {
        // Reorder among queued tasks
        var queuedIds = [];
        state.tasks.forEach(function (t) {
          if (t.status === "queued") queuedIds.push(t.id);
        });
        var fromIdx = queuedIds.indexOf(draggedId);
        if (fromIdx < 0) return;
        queuedIds.splice(fromIdx, 1);
        var toIdx = getDropIndexAmongStatus(taskListEl, e.clientY, "queued");
        queuedIds.splice(toIdx, 0, draggedId);

        apiPut("api/tasks/reorder", { task_ids: queuedIds }).catch(function () {
          showToast("Failed to reorder tasks");
        });
        for (var i = 0; i < queuedIds.length; i++) {
          var t = findTask(queuedIds[i]);
          if (t) t.priority = i + 1;
        }
      } else {
        // Reorder finished tasks visually via created_at swapping
        var finishedTasks = [];
        state.tasks.forEach(function (t) {
          if (t.status === "completed" || t.status === "failed") finishedTasks.push(t);
        });
        var fromIdx2 = -1;
        for (var j = 0; j < finishedTasks.length; j++) {
          if (finishedTasks[j].id === draggedId) { fromIdx2 = j; break; }
        }
        if (fromIdx2 < 0) return;
        finishedTasks.splice(fromIdx2, 1);
        var toIdx2 = getDropIndexAmongStatus(taskListEl, e.clientY, "finished");
        finishedTasks.splice(toIdx2, 0, draggedTask);
        // Reassign created_at to maintain the visual order across polls
        var timestamps = finishedTasks.map(function (t) { return t.created_at; });
        timestamps.sort();
        for (var k = 0; k < finishedTasks.length; k++) {
          finishedTasks[k].created_at = timestamps[k];
        }
      }

      sortTasks();
      renderTaskList();
    });
  }

  // Cached at dragstart: { all: number[], statusGrouped: { queued: number[], finished: number[] } }
  var _taskDragCache = null;
  var _taskListDragOverRaf = null;
  var _taskListPendingDragOver = null;

  function _cacheTaskDragMidpoints(container) {
    var cards = container.querySelectorAll(".task-card:not(.dragging)");
    var all = new Array(cards.length);
    var queued = [];
    var finished = [];
    for (var i = 0; i < cards.length; i++) {
      var r = cards[i].getBoundingClientRect();
      var mid = r.top + r.height / 2;
      all[i] = mid;
      var t = findTask(cards[i].dataset.taskId);
      if (!t) continue;
      if (t.status === "queued") queued.push(mid);
      else if (t.status === "completed" || t.status === "failed") finished.push(mid);
    }
    _taskDragCache = { all: all, queued: queued, finished: finished };
  }

  function getDropIndex(container, clientY) {
    if (!_taskDragCache) _cacheTaskDragMidpoints(container);
    var mids = _taskDragCache.all;
    for (var i = 0; i < mids.length; i++) {
      if (clientY < mids[i]) return i;
    }
    return mids.length;
  }

  function getDropIndexAmongStatus(container, clientY, group) {
    if (!_taskDragCache) _cacheTaskDragMidpoints(container);
    var mids = group === "queued" ? _taskDragCache.queued : _taskDragCache.finished;
    for (var i = 0; i < mids.length; i++) {
      if (clientY < mids[i]) return i;
    }
    return mids.length;
  }

  function clearDragIndicators(container) {
    var cards = container.querySelectorAll(".task-card.drag-over");
    for (var i = 0; i < cards.length; i++) cards[i].classList.remove("drag-over");
    container.classList.remove("drag-over-append");
  }

  // Sets value without dispatching input; callers needing listeners must dispatch it themselves.
  function setInputValue(selector, value) {
    var inp = qs(selector);
    if (inp) inp.value = value;
  }

  function syncValueDisplays() {
    var inputs = qsa(".param-control input[type='range']");
    for (var i = 0; i < inputs.length; i++) {
      var valSpan = inputs[i].parentNode.querySelector(".param-value");
      // _updateMinAreaReadout owns the worded "Min area %" readout.
      if (valSpan && !valSpan.classList.contains("param-value--minarea")) {
        valSpan.textContent = inputs[i].value;
      }
    }
  }

  function restoreTaskToWorkflow(task) {
    // Switch workflow tab
    state.activeWorkflow = task.type;
    qsa(".wf-tab").forEach(function (t) { t.classList.remove("active"); });
    var targetTab = qs('.wf-tab[data-type="' + task.type + '"]');
    if (targetTab) targetTab.classList.add("active");
    // No .click() here, so mirror the tab into the category nav (late-bound).
    if (SS.syncToolCategoryNav) SS.syncToolCategoryNav();

    // Select participant
    if (task.participant) {
      state.selectedParticipant = task.participant;
      var sel = qs("#participantSelect");
      if (sel) sel.value = task.participant;
    }

    // Select region
    if (task.region_ref) {
      var restoredRef = normalizeRegionRef(task.region_ref);
      state.runRegions = restoredRef ? [restoredRef] : [];
      // Explicit region choice; keep the chip-follow seed from replacing it (renderRunRegionPicker).
      state.runRegionsSeeded = false;
      state.pendingRegion = null;
      if (restoredRef && restoredRef.source === "active" && state.regions[restoredRef.name]) {
        state.activeRegion = restoredRef.name;
      } else {
        state.activeRegion = null;
      }
      renderRegionChips();
      renderRunRegionPicker();
      renderOverlay();
      updateRegionButtons();
    } else if (task.region && state.regions[task.region]) {
      state.activeRegion = task.region;
      state.pendingRegion = null;
      state.runRegions = [activeRegionRef(task.region)];
      state.runRegionsSeeded = false;
      renderRegionChips();
      renderRunRegionPicker();
      renderOverlay();
      updateRegionButtons();
    }

    // For similarity, restore reference timestamp before rendering params
    if (task.type === "similarity") {
      var params = task.parameters || {};
      if (params.reference_timestamp !== undefined) {
        state.referenceTimestamp = params.reference_timestamp;
      } else {
        showToast("Reference frame must be recaptured");
      }
    }

    // For scene, restore references into state before rendering so the list shows them.
    if (task.type === "scene") {
      var sceneParams = task.parameters || {};
      state.sceneReferences = (sceneParams.scene_references || []).map(function (ref) {
        return { name: ref.name, timestamp: ref.timestamp, threshold: numberOrDefault(ref.threshold, 0.75) };
      });
    }

    // Rebuild multitool steps first; step._initial lets _mtRender* set values at creation.
    if (task.type === "multitool") {
      var mtParams = task.parameters || {};
      state.multitoolFocus = 0;
      state.multitoolSteps = (mtParams.steps || []).map(function (s) {
        var step = { type: s.type, collapsed: true };
        step.logic = (s.logic || "AND").toUpperCase();
        if (s.offset && typeof s.offset.min === "number" && typeof s.offset.max === "number") {
          step.offset = { min: s.offset.min, max: s.offset.max };
        }
        if (s.region) step.region = s.region;
        if (s.region_ref) step.region_ref = s.region_ref;
        if (s.reference_timestamp !== undefined) step._refTs = s.reference_timestamp;
        if (s.scene_references) step._scenes = s.scene_references.map(function (ref) {
          return { name: ref.name, timestamp: ref.timestamp, threshold: numberOrDefault(ref.threshold, 0.75) };
        });
        step._initial = s;
        return step;
      });
    }

    // Calibration waits for refreshCalibration() below; {defaults: true} so absent params read absent.
    state.suppressCalibrationRefresh = true;
    renderWorkflowParams({ defaults: true });
    state.suppressCalibrationRefresh = false;

    params = task.parameters || {};
    if (task.type === "multitool") {
      setInputValue("#paramMultitoolInterval", numberOrDefault(params.interval, 1.0));
    } else if (task.type === "color") {
      var tc = params.target_color || {};
      var ch = numberOrDefault(tc.h, 90);
      var cs = numberOrDefault(tc.s, 200);
      var cv = numberOrDefault(tc.v, 200);
      var savedTol = params.tolerance ? Math.round(params.tolerance.h * 100 / 90) : 30;
      setInputValue("#paramColorTol", savedTol);
      setInputValue("#paramColorInterval", numberOrDefault(params.interval, 1.0));
      var savedColorMode = _colorMode(params.color_mode);
      applyColorMode("paramColorMode", savedColorMode);
      // Server drops min_coverage at 0 ("any presence"); restore 0, not the default.
      if (savedColorMode === "presence") {
        setInputValue("#paramColorMinArea", params.min_coverage != null ? params.min_coverage * 100 : 0);
      }
      _updateMinAreaReadout("");
      // setTargetColor writes hidden h/s/v + hex input + preview + palette + brightness strip.
      SS.setTargetColor(ch, cs, cv);
    } else if (task.type === "change") {
      setInputValue("#paramChangeThresh", numberOrDefault(params.threshold, 0.03));
      setInputValue("#paramChangeNoise", intOrDefault(params.noise_threshold, 30));
      setInputValue("#paramChangeInterval", numberOrDefault(params.interval, 1.0));
      setInputValue("#paramChangeConsecutive", intOrDefault(params.require_consecutive, 1));
    } else if (task.type === "similarity") {
      setInputValue("#paramSimThresh", numberOrDefault(params.threshold, 0.90));
      setInputValue("#paramSimInterval", numberOrDefault(params.interval, 1.0));
    } else if (task.type === "text") {
      setInputValue("#paramTextSearch", params.search_string || "");
      setInputValue("#paramTextFuzzy", numberOrDefault(params.fuzzy_threshold, CLIPGEN_CONFIG.screenspaceOcrFuzzyThreshold));
      setInputValue("#paramTextOcrConf", numberOrDefault(params.ocr_confidence_threshold, CLIPGEN_CONFIG.screenspaceOcrMinConfidence));
      var textPpEl = qs("#paramTextOcrPreprocess");
      if (textPpEl) textPpEl.checked = !!params.ocr_preprocess;
      applyNormalizeMode("paramTextOcrNormalize", params.ocr_normalize || "off");
      setInputValue("#paramTextInterval", numberOrDefault(params.interval, 2.0));
      if (params.languages && params.languages[0]) {
        setInputValue("#paramTextLang", params.languages[0]);
      }
      setInputValue("#paramTextConsecutive", intOrDefault(params.require_consecutive, 1));
    } else if (task.type === "numbers") {
      setInputValue("#paramNumOperator", params.operator || "gt");
      // Fire change so range/target row visibility tracks the restored operator
      // (listener attached in renderNumbersParams).
      var opSel = qs("#paramNumOperator");
      if (opSel) opSel.dispatchEvent(new Event("change"));
      if (params.operator === "range") {
        setInputValue("#paramNumMin", numberOrDefault(params.range_min, 0));
        setInputValue("#paramNumMax", numberOrDefault(params.range_max, 100));
      } else {
        setInputValue("#paramNumTarget", numberOrDefault(params.target_value, 100));
      }
      setInputValue("#paramNumOcrConf", numberOrDefault(params.ocr_confidence_threshold, CLIPGEN_CONFIG.screenspaceOcrMinConfidence));
      var numPpEl = qs("#paramNumOcrPreprocess");
      if (numPpEl) numPpEl.checked = !!params.ocr_preprocess;
      var numIoEl = qs("#paramNumIntegersOnly");
      if (numIoEl) numIoEl.checked = !!params.integers_only;
      setInputValue("#paramNumInterval", numberOrDefault(params.interval, 2.0));
      setInputValue("#paramNumConsecutive", intOrDefault(params.require_consecutive, 1));
    } else if (task.type === "timelapse") {
      setInputValue("#paramTlSpeed", numberOrDefault(params.speedup_factor, 10));
      setInputValue("#paramTlFormat", params.output_format || "mp4");
      if (params.sample_interval !== undefined) {
        setInputValue("#paramTlSampleInterval", params.sample_interval);
      }
    } else if (task.type === "template") {
      if (params.reference_timestamp !== undefined) {
        state.referenceTimestamp = params.reference_timestamp;
        // Re-arm the capture region (name only) so a re-run samples the same region.
        state.capturedRefPreview = params.reference_region
          ? { region: params.reference_region, ts: params.reference_timestamp, dataUrl: null }
          : null;
      }
      setInputValue("#paramTemplateThresh", numberOrDefault(params.threshold, 0.70));
      setInputValue("#paramTemplateInterval", numberOrDefault(params.interval, 1.0));
      if (params.template_scale) {
        setInputValue("#paramTemplateScale", Math.round(params.template_scale * 100));
      }
    } else if (task.type === "shape") {
      if (params.reference_timestamp !== undefined) {
        state.referenceTimestamp = params.reference_timestamp;
        // Re-arm the capture region (name only) so a re-run samples the same region.
        state.capturedRefPreview = params.reference_region
          ? { region: params.reference_region, ts: params.reference_timestamp, dataUrl: null }
          : null;
      }
      setInputValue("#paramShapeThresh", numberOrDefault(params.threshold, 0.55));
      if (params.scale_min) {
        setInputValue("#paramShapeScaleMin", Math.round(params.scale_min * 100));
      }
      if (params.scale_max) {
        setInputValue("#paramShapeScaleMax", Math.round(params.scale_max * 100));
      }
      setInputValue("#paramShapeSteps", intOrDefault(params.scale_steps, 7));
      if (params.scale_y_min || params.scale_y_max) {
        var linkEl = qs("#paramShapeLinkAxes");
        if (linkEl) {
          linkEl.checked = false;
          // The change handler owns the V-row visibility sync.
          linkEl.dispatchEvent(new Event("change", { bubbles: true }));
        }
        if (params.scale_y_min) {
          setInputValue("#paramShapeScaleYMin", Math.round(params.scale_y_min * 100));
        }
        if (params.scale_y_max) {
          setInputValue("#paramShapeScaleYMax", Math.round(params.scale_y_max * 100));
        }
        setInputValue("#paramShapeStepsY", intOrDefault(params.scale_y_steps, 3));
      }
      setInputValue("#paramShapeInterval", numberOrDefault(params.interval, 1.0));
    } else if (task.type === "flow") {
      setInputValue("#paramFlowMag", numberOrDefault(params.magnitude_threshold, 2.0));
      setInputValue("#paramFlowInterval", numberOrDefault(params.interval, 1.0));
      setInputValue("#paramFlowConsecutive", intOrDefault(params.require_consecutive, 1));
    } else if (task.type === "scene") {
      setInputValue("#paramSceneInterval", numberOrDefault(params.interval, 1.0));
    } else if (task.type === "inactivity") {
      setInputValue("#paramInactThresh", intOrDefault(params.threshold, 10));
      setInputValue("#paramInactMinDur", numberOrDefault(params.min_duration, 2.0));
      setInputValue("#paramInactInterval", numberOrDefault(params.interval, 1.0));
    } else if (task.type === "boundary") {
      setInputValue("#paramBoundaryMetric", params.metric || "");
      setInputValue("#paramBoundaryThresh", intOrDefault(params.threshold, 14));
      setInputValue("#paramBoundaryMinGap", numberOrDefault(params.min_gap, 3.0));
      setInputValue("#paramBoundaryInterval", numberOrDefault(params.interval, 1.0));
    }

    // Shared by every non-timelapse type; gatherWorkflowParams is the save path.
    if (task.type !== "timelapse") {
      if (params.event_label) setInputValue("#paramEventLabel", params.event_label);
      if (params.detect_first) {
        var dfEl = qs("#paramDetectFirst");
        if (dfEl) dfEl.checked = true;
      }
    }

    if (state.restoreMarkersOnEdit) {
      var hasIn = params.start_seconds !== undefined && params.start_seconds !== null;
      var hasOut = params.end_seconds !== undefined && params.end_seconds !== null;
      if (hasIn || hasOut) {
        state.inMarker = hasIn ? params.start_seconds : null;
        state.outMarker = hasOut ? params.end_seconds : null;
        // Markers persist per participant; write the restored range through too.
        if (SS.persistMarkers) SS.persistMarkers();
        updateMarkerInfo();
        renderTimeline();
      }
    }

    syncValueDisplays();
    // setInputValue dispatches nothing, so re-derive reset buttons against the task's values.
    updateParamResetButtons();
    updateRunButton();
    // Re-evaluate pins so the strip and Run hint reflect the restored params.
    refreshCalibration();
    showToast("Restored " + task.type + " task parameters");
  }

  function findTask(id) {
    for (var i = 0; i < state.tasks.length; i++) {
      if (state.tasks[i].id === id) return state.tasks[i];
    }
    return null;
  }

  function focusedTaskId() {
    if (state.hoveredTaskId) {
      var ht = findTask(state.hoveredTaskId);
      if (ht && ht.status === "completed") return state.hoveredTaskId;
    }
    return state.selectedTaskId;
  }

  function updatePauseButton() {
    var btn = qs("#taskQueuePauseBtn");
    if (!btn) return;
    btn.innerHTML = "";
    if (state.queuePaused) {
      btn.appendChild(iconSpan("play"));
      btn.title = "Resume queue";
    } else {
      btn.appendChild(iconSpan("pause"));
      btn.title = "Pause queue";
    }
  }

  function initPauseButton() {
    var btn = qs("#taskQueuePauseBtn");
    if (!btn) return;
    updatePauseButton();
    btn.addEventListener("click", function () {
      var endpoint = state.queuePaused ? "api/tasks/resume" : "api/tasks/pause";
      apiPost(endpoint)
        .then(function (data) {
          if (data.ok) {
            state.queuePaused = data.paused;
            updatePauseButton();
          }
        })
        .catch(function (err) { showToast("Error: " + err.message); });
    });
  }

  // ---- Task list: filter chips + card DOM (drag reorder uses _taskDragCache) ----

  function initTaskFilters() {
    var doneBtn = qs("#taskFilterDoneBtn");
    var failedBtn = qs("#taskFilterFailedBtn");
    if (doneBtn) {
      doneBtn.appendChild(iconSpan("check"));
      doneBtn.addEventListener("click", function () { toggleTaskFilter("completed"); });
    }
    if (failedBtn) {
      failedBtn.appendChild(iconSpan("x-mark"));
      failedBtn.addEventListener("click", function () { toggleTaskFilter("failed"); });
    }
  }

  function toggleTaskFilter(status) {
    state.taskFilter = state.taskFilter === status ? null : status;
    updateTaskFilterButtons();
    renderTaskList();
  }

  function updateTaskFilterButtons() {
    var doneBtn = qs("#taskFilterDoneBtn");
    var failedBtn = qs("#taskFilterFailedBtn");
    if (doneBtn) doneBtn.classList.toggle("active", state.taskFilter === "completed");
    if (failedBtn) failedBtn.classList.toggle("active", state.taskFilter === "failed");
  }

  function renderTaskList() {
    sortTasks();
    var container = qs("#taskList");
    var count = qs("#taskCount");
    var filtered = state.taskFilter
      ? state.tasks.filter(function (t) { return t.status === state.taskFilter; })
      : state.tasks;
    count.textContent = "(" + filtered.length + ")";
    updateTaskFilterButtons();
    if (state.tasks.length === 0) {
      container.innerHTML = "";
      container.appendChild(el("div", "panel-empty", "No tasks yet. Configure a workflow and click Run."));
      return;
    }
    if (filtered.length === 0) {
      container.innerHTML = "";
      container.appendChild(el("div", "panel-empty", "No " + state.taskFilter + " tasks."));
      return;
    }
    var frag = document.createDocumentFragment();
    filtered.forEach(function (task) {
      var card = el("div", "task-card task-card-" + task.status);
      card.dataset.taskId = task.id;
      if (task.id === state.selectedTaskId) card.classList.add("selected");

      // Drag handle for reorderable tasks (completed, failed, queued)
      var isDraggable = task.status === "queued" || task.status === "completed" || task.status === "failed";
      if (isDraggable) {
        var handle = el("span", "task-card-drag-handle");
        handle.appendChild(iconSpan("bars-2"));
        handle.addEventListener("mousedown", function () { card.setAttribute("draggable", "true"); });
        handle.addEventListener("mouseup", function () { card.removeAttribute("draggable"); });
        card.appendChild(handle);
      } else if (task.status === "running") {
        card.appendChild(el("span", "task-card-spinner"));
      } else if (task.status === "paused") {
        var pauseIcon = el("span", "task-card-pause-icon");
        pauseIcon.appendChild(iconSpan("pause", "ss-icon--xs"));
        card.appendChild(pauseIcon);
      }

      // Type badge
      var badge = el("span", "task-card-type");
      badge.style.color = taskTypeColor(task.type);
      badge.title = task.type;
      var typeIconEl = el("span", "task-card-type-icon");
      var iconFile = TASK_TYPE_ICON_FILES[task.type] || "squares-2x2";
      applyIconMask(typeIconEl, iconFile, "/screenspace/icons/");
      badge.appendChild(typeIconEl);
      card.appendChild(badge);

      // Fast scan badge
      if ((task.parameters || {}).scan_mode === "fast") {
        var fb = el("span", "task-fast-badge");
        var bi = el("span", "task-fast-badge-icon");
        applyIconMask(bi, "chevron-double-right", "/screenspace/icons/");
        fb.appendChild(bi);
        fb.appendChild(document.createTextNode("Fast"));
        card.appendChild(fb);
      }

      // Info
      var info = el("div", "task-card-info");
      var meta = el("span", "task-card-meta");
      var eventLabel = (task.parameters || {}).event_label;
      var title = eventLabel || task.name;
      if (title) {
        meta.classList.add("task-card-meta--title");
        meta.textContent = title;
        info.appendChild(meta);
        var sub = el("span", "task-card-meta");
        sub.textContent = task.participant + " \u00b7 " + (task.region || "");
        info.appendChild(sub);
      } else {
        meta.textContent = task.participant + " \u00b7 " + (task.region || "");
        info.appendChild(meta);
      }

      if (task.status === "running" || task.status === "paused") {
        var prog = el("div", "task-card-progress");
        var fill = el("div", "task-card-progress-fill");
        fill.dataset.taskProgress = task.id;
        fill.style.width = Math.round((task.progress || 0) * 100) + "%";
        prog.appendChild(fill);
        info.appendChild(prog);

        // Live elapsed / ETA line, refreshed every second by the eta ticker.
        var etaEl = el("span", "task-card-eta");
        etaEl.dataset.taskEta = task.id;
        etaEl.textContent = taskEtaLabel(task);
        info.appendChild(etaEl);
      }
      card.appendChild(info);

      // Status text; tickTaskProgress refreshes running/paused text in place.
      if (task.status === "failed" && task.error) card.title = task.error;
      // Only running shimmers; a status transition rebuilds the card, so no clearing.
      var statusSpan = el(
        "span",
        "task-card-status" + (task.status === "running" ? " cg-shimmer" : ""),
        taskStatusText(task)
      );
      statusSpan.dataset.taskStatus = task.id;
      if (task.status === "failed" && task.error) {
        // Truncated error text; click toasts the full error. stopPropagation avoids selecting the card.
        statusSpan.classList.add("task-card-status-error");
        statusSpan.title = "Click to view the full error";
        (function (err) {
          statusSpan.addEventListener("click", function (e) {
            e.stopPropagation();
            showToast(err);
          });
        })(task.error);
      }
      card.appendChild(statusSpan);

      // Edit button
      var editBtn = el("button", "task-card-edit");
      editBtn.title = "Edit";
      editBtn.appendChild(iconSpan("pencil-square"));
      card.appendChild(editBtn);

      // Dismiss button
      var dismissBtn = el("button", "task-card-dismiss");
      dismissBtn.title = "Dismiss";
      dismissBtn.appendChild(iconSpan("x-mark"));
      card.appendChild(dismissBtn);

      frag.appendChild(card);
    });
    var prevScrollTop = container.scrollTop;
    container.innerHTML = "";
    container.appendChild(frag);
    container.scrollTop = prevScrollTop;
    // The rebuild (and re-sort) drops the keyboard cursor; re-anchor it by task id.
    if (SS.ssRefreshNav) SS.ssRefreshNav();
    updateResultsCrumb();
  }

  // ---- Elapsed / ETA ticker ----

  function taskIsActive(task) {
    return task.status === "running" || task.status === "paused";
  }

  function taskIsTerminal(task) {
    return (
      task.status === "completed" ||
      task.status === "failed" ||
      task.status === "cancelled"
    );
  }

  // Pure; both the card build and tickTaskProgress call it (progress skips the rebuild gate).
  function taskStatusText(task) {
    // Ticks carry result_count (not the result list — see #521); use it here.
    var rLen = task.result_count || 0;
    if (task.status === "running") {
      var rPct = Math.round((task.progress || 0) * 100);
      // "scanning…" reads as in-progress, not hung, before the first hit.
      return rPct + "%" + (rLen ? " · " + rLen + " result" + (rLen !== 1 ? "s" : "") : " · scanning…");
    }
    if (task.status === "paused") {
      var pPct = Math.round((task.progress || 0) * 100);
      return "paused " + pPct + "%" + (rLen ? " · " + rLen + " result" + (rLen !== 1 ? "s" : "") : "");
    }
    if (task.status === "failed" && task.error) return task.error;
    if (task.status === "completed" && rLen) {
      return rLen + " result" + (rLen !== 1 ? "s" : "");
    }
    return task.status;
  }

  // In-place progress fill + status refresh; transitions still rebuild via renderTaskList.
  function tickTaskProgress() {
    var fills = document.querySelectorAll("#taskList [data-task-progress]");
    for (var i = 0; i < fills.length; i++) {
      var t = findTask(fills[i].dataset.taskProgress);
      if (t && taskIsActive(t)) {
        fills[i].style.width = Math.round((t.progress || 0) * 100) + "%";
      }
    }
    var spans = document.querySelectorAll("#taskList [data-task-status]");
    for (var j = 0; j < spans.length; j++) {
      var st = findTask(spans[j].dataset.taskStatus);
      if (st && taskIsActive(st)) spans[j].textContent = taskStatusText(st);
    }
  }

  // "0:42 · ~1:20 left" label. Seeded from created_at (includes queue wait); paused shows elapsed only.
  function taskEtaLabel(task) {
    var t = _etaTrackers[task.id];
    if (!t) {
      t = createEtaTracker();
      var seed = task.created_at ? Date.parse(task.created_at) : NaN;
      t.start(isNaN(seed) ? undefined : seed);
      _etaTrackers[task.id] = t;
    }
    // Freeze the elapsed clock while paused; resume folds the paused span out.
    if (task.status === "paused") t.pause();
    else t.resume();
    var prog = task.status === "paused" ? 0 : task.progress;
    var e = t.update(prog);
    var label = formatDuration(e.elapsedSec);
    var eta = formatEtaLabel(e.remainingSec);
    if (eta) label += " · " + eta;
    return label;
  }

  function tickEtas() {
    // Prune only terminal/gone tasks: a resumed task passes through "queued" and keeps its paused-time accumulator.
    var keepIds = {};
    state.tasks.forEach(function (t) {
      if (!taskIsTerminal(t)) keepIds[t.id] = true;
    });
    Object.keys(_etaTrackers).forEach(function (id) {
      if (!keepIds[id]) delete _etaTrackers[id];
    });
    // Refresh the visible eta spans in place (no list re-render).
    var spans = document.querySelectorAll("#taskList [data-task-eta]");
    for (var i = 0; i < spans.length; i++) {
      var task = findTask(spans[i].dataset.taskEta);
      if (task && taskIsActive(task)) spans[i].textContent = taskEtaLabel(task);
    }
  }

  function ensureEtaTicker() {
    // Start only while a task is active and the tab is visible; otherwise stop.
    if (state.tasks.some(taskIsActive) && !document.hidden) _etaTicker.ensure();
    else _etaTicker.stop();
  }

  // ---- SSE (Server-Sent Events) with polling fallback ----

  // Heatmap filenames land after completion (outside the worker lock); signature catches late arrival.
  function _heatmapSig(t) {
    if (!t) return "";
    return (t.heatmap || "") + "|" + (t.heatmap_gif || "") + "|" + (t.heatmap_rolling_gif || "");
  }

  // Results are append-only during a scan, so result_count is a safe tail cursor for _syncTaskResults.
  var _resultFetching = {};
  var _taskStatusSeen = {};

  // Drop caches no longer append-consistent: completion merges hits into spans; pause/resume shrinks result_count.
  function _reconcileResultCache(tasks) {
    var seenNow = {};
    tasks.forEach(function (t) {
      seenNow[t.id] = t.status;
      var prev = _taskStatusSeen[t.id];
      var terminal = t.status === "completed" || t.status === "failed" || t.status === "cancelled";
      if (prev && prev !== t.status && terminal) {
        delete state.taskResults[t.id];
      }
      var cached = state.taskResults[t.id];
      if (Array.isArray(cached) && (t.result_count || 0) < cached.length) {
        delete state.taskResults[t.id];
      }
    });
    // Prune cache/fetch state for tasks that no longer exist (dismissed).
    Object.keys(state.taskResults).forEach(function (id) {
      if (!(id in seenNow)) delete state.taskResults[id];
    });
    Object.keys(_resultFetching).forEach(function (id) {
      if (!(id in seenNow)) delete _resultFetching[id];
    });
    _taskStatusSeen = seenNow;
  }

  function _syncTaskResults() {
    if (!Array.isArray(state.tasks)) return;
    state.tasks.forEach(function (t) {
      if (t.status === "cancelled") return;
      if (t.participant && t.participant !== state.selectedParticipant) return;
      var count = t.result_count || 0;
      var cached = state.taskResults[t.id];
      var have = Array.isArray(cached) ? cached.length : 0;
      if (count <= have || _resultFetching[t.id]) return;
      _resultFetching[t.id] = true;
      var reqId = t.id;
      var since = have;
      apiGet("api/tasks/" + reqId + "/results?since=" + since).then(
        function (data) {
          _resultFetching[reqId] = false;
          if (!data || !data.ok) return;
          var base = Array.isArray(state.taskResults[reqId]) ? state.taskResults[reqId] : [];
          // Reject if the cursor moved under us; the next tick reconciles.
          if (since !== base.length) return;
          if (data.results && data.results.length) {
            state.taskResults[reqId] = base.concat(data.results);
          } else if (!Array.isArray(state.taskResults[reqId])) {
            state.taskResults[reqId] = base;
          }
          if (reqId === state.selectedTaskId) {
            state.selectedTaskResults = state.taskResults[reqId];
            SS.renderResults();
          }
          renderTimeline();
        },
        function () {
          _resultFetching[reqId] = false;
        }
      );
    });
  }

  function handleTaskData(data) {
    if (!data.ok) return;
    var oldSelected = state.selectedTaskId;
    var oldTask = oldSelected ? findTask(oldSelected) : null;
    var wasRunning = oldTask && (oldTask.status === "queued" || oldTask.status === "running");
    var oldHeatmapSig = _heatmapSig(oldTask);
    state.tasks = data.tasks;
    _reconcileResultCache(data.tasks);
    // Rebuild the pause icon only on flips; every push (~2/s) would refetch its svg.
    if (data.paused !== undefined && data.paused !== state.queuePaused) {
      state.queuePaused = data.paused;
      updatePauseButton();
    }
    // Rebuild only on status/heatmap transitions; progress and result_count update in place (tickTaskProgress, _syncTaskResults).
    var taskFp = JSON.stringify(data.tasks.map(function (t) {
      return t.id + ":" + t.status + ":" + _heatmapSig(t);
    }));
    var taskChanged = taskFp !== _lastTaskFp;
    _lastTaskFp = taskFp;
    if (taskChanged) {
      renderTaskList();
      renderTimeline();
    } else {
      tickTaskProgress(); // in-place bar + status refresh when no rebuild
    }
    // Fetch result tails; this redraws the timeline as hits land mid-scan.
    _syncTaskResults();
    // Auto-load results when selected task completes
    if (wasRunning && oldSelected) {
      var newTask = findTask(oldSelected);
      if (newTask && newTask.status === "completed") {
        SS.loadAndShowResults(oldSelected);
      }
    } else if (oldSelected) {
      // Late heatmap arrival on a completed task: re-render so the section appears.
      var curTask = findTask(oldSelected);
      if (curTask && curTask.status === "completed" && _heatmapSig(curTask) !== oldHeatmapSig) {
        if (state.selectedTaskResults) SS.renderResults();
        else SS.loadAndShowResults(oldSelected);
      }
    }
    ensureEtaTicker();
  }

  function startSSE() {
    if (state.eventSource) return;
    state.eventSource = createSSEStream("api/tasks/stream", {
      // Retire the poller or drop-and-recover leaves both transports running; re-arm the drop toast.
      onOpen: function () {
        stopPolling();
        state.sseFellBack = false;
      },
      onUnsupported: startPolling,
      onMessage: handleTaskData,
      onError: function () {
        // Fall back to polling; onError repeats, so toast once per drop.
        state.eventSource = null;
        if (!state.sseFellBack) {
          state.sseFellBack = true;
          showToast("Live updates interrupted. Falling back to polling");
        }
        startPolling();
      },
    });
  }

  // ---- Polling (fallback) ----

  function startPolling() {
    if (state.poller) return;
    // createPoller handles hidden tabs itself; runImmediately false matches the old setInterval.
    state.poller = createPoller(pollTasks, POLL_INTERVAL, {
      runImmediately: false,
      label: "screenspace.tasks",
    });
    state.poller.start();
  }

  function stopPolling() {
    if (state.poller) {
      state.poller.stop();
      state.poller = null;
    }
  }

  function pollTasks() {
    var hasActive = state.tasks.some(function (t) {
      return t.status === "queued" || t.status === "running" || t.status === "paused";
    });
    if (!hasActive) {
      stopPolling();
      return;
    }

    apiGet("api/tasks")
      .then(function (data) { handleTaskData(data); })
      .catch(function () {});
  }

  function stopSSE() {
    if (state.eventSource) {
      state.eventSource.close();
      state.eventSource = null;
    }
  }

  document.addEventListener("dragend", function () {
    if (SS.cancelMultitoolDrag) SS.cancelMultitoolDrag();
    if (_taskListDragOverRaf != null) {
      cancelAnimationFrame(_taskListDragOverRaf);
      _taskListDragOverRaf = null;
    }
    _taskListPendingDragOver = null;
    var stepsDiv = document.querySelector(".multitool-steps");
    if (stepsDiv) SS.clearMultitoolDragIndicators(stepsDiv);
    var taskListEl = qs("#taskList");
    if (taskListEl) clearDragIndicators(taskListEl);
  });

  document.addEventListener("visibilitychange", function () {
    if (document.hidden) {
      stopSSE();
      stopPolling();
      _etaTicker.stop();
      return;
    }
    var hasActive = state.tasks.some(function (t) {
      return t.status === "queued" || t.status === "running" || t.status === "paused";
    });
    if (hasActive) startSSE();
    ensureEtaTicker();
  });

  // ---- Satellite interface (window.ClipgenScreenspace) ----
  // Consumed by hub delegators and later-loading satellites.
  SS.findTask = findTask;
  SS.focusedTaskId = focusedTaskId;
  SS.renderTaskList = renderTaskList;
  SS.startSSE = startSSE;
  // Boot and participant switches call this so the timeline has markers before a tick.
  SS.syncTaskResults = _syncTaskResults;
  SS.reconcileResultCache = _reconcileResultCache;
  SS.setRightPaneTab = setRightPaneTab;
  SS.updateResultsCrumb = updateResultsCrumb;
  SS.initRightPaneTabs = initRightPaneTabs;
  SS.initPauseButton = initPauseButton;
  SS.initTaskQueue = initTaskQueue;
  SS.initTaskFilters = initTaskFilters;
  SS.restoreTaskToWorkflow = restoreTaskToWorkflow;
  SS.setInputValue = setInputValue;
  SS.syncValueDisplays = syncValueDisplays;
})();
