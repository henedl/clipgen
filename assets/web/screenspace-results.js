/* clipgen Screenspace results satellite — screenspace-results.js
 *
 * The Results panel carved out of screenspace.js: the confidence histogram +
 * per-result rows (with the certainty-cutoff exclude filtering and per-tool
 * specials — timelapse, colour span, scene name), and the per-result heatmap
 * overlay. Loaded last (after screenspace-tasks.js, which publishes findTask /
 * renderTaskList / startSSE / updateResultsCrumb that this file destructures).
 * Reads the hub's shared state + helpers through window.ClipgenScreenspace and
 * publishes its three entry points back; the hub keeps thin delegators for the
 * ones its own code calls (initResultsPanel / renderResults). Function bodies are
 * unchanged from when they lived inline in screenspace.js.
 */
(function () {
  "use strict";

  var SS = window.ClipgenScreenspace;
  var state = SS.state;
  var buildTypeIcon = SS.buildTypeIcon,
    findTask = SS.findTask,
    iconSpan = SS.iconSpan,
    loadFrame = SS.loadFrame,
    renderOverlay = SS.renderOverlay,
    renderTaskList = SS.renderTaskList,
    renderTimeline = SS.renderTimeline,
    selectParticipant = SS.selectParticipant,
    startSSE = SS.startSSE,
    taskRegionPixels = SS.taskRegionPixels,
    taskTypeColor = SS.taskTypeColor,
    updateResultsCrumb = SS.updateResultsCrumb;

  // Confidence bar mirrors the prototype's ConfBar: 4 px tall, hue-tinted fill,
  // opacity ramps from 0.4 (low) to 1.0 (full) so high-confidence rows feel
  // saturated while low ones recede.
  function buildConfBar(value, type) {
    var v = Math.max(0, Math.min(1, Number(value) || 0));
    var bar = el("div", "result-bar");
    var fill = el("div", "result-bar-fill");
    fill.style.width = Math.round(v * 100) + "%";
    fill.style.background = taskTypeColor(type);
    fill.style.opacity = (0.4 + v * 0.6).toFixed(2);
    bar.appendChild(fill);
    return bar;
  }

  // Confidence-distribution histogram for the Results panel. Buckets every
  // event's confidence (0–1, uniform across tools) into 10 fixed bins and draws
  // bottom-aligned bars scaled to the tallest bin, tinted with the tool color
  // and using the same opacity ramp as buildConfBar. An absolutely-positioned
  // marker shows the current certainty cutoff. Rebuilt wholesale on each
  // renderResults() call, so it tracks the slider and stays color-consistent
  // with the result rows at render time.
  function renderConfidenceHistogram(events, taskType) {
    var host = qs("#confHistogram");
    if (!host) return;
    host.innerHTML = "";

    var BINS = 10;
    var counts = [];
    var i;
    for (i = 0; i < BINS; i++) counts[i] = 0;
    for (i = 0; i < events.length; i++) {
      var c = Number(events[i].confidence);
      if (isNaN(c)) c = 0;
      else if (c < 0) c = 0;
      else if (c > 1) c = 1;
      counts[Math.min(BINS - 1, Math.floor(c * BINS))]++; // conf 1.0 -> bin 9
    }
    var maxCount = 0;
    for (i = 0; i < BINS; i++) if (counts[i] > maxCount) maxCount = counts[i];

    var track = el("div", "conf-hist-track");
    var color = taskTypeColor(taskType);
    for (i = 0; i < BINS; i++) {
      var n = counts[i];
      var pct = maxCount > 0 ? (n / maxCount) * 100 : 0;
      if (n > 0 && pct < 6) pct = 6; // floor so a 1-event bin stays visible
      var bar = el("div", "conf-hist-bar");
      var fill = el("div", "conf-hist-bar-fill");
      fill.style.height = pct.toFixed(1) + "%";
      fill.style.background = color;
      fill.style.opacity = (0.4 + (i / (BINS - 1)) * 0.6).toFixed(2); // mirrors buildConfBar
      bar.appendChild(fill);
      (function (lo, hi, count) {
        attachHoverTooltip(bar, function () {
          return Math.round(lo * 100) + "–" + Math.round(hi * 100) + "%: " +
            clipgenPluralUnit(count, "event", "events");
        }, { align: "center" });
      })(i / BINS, (i + 1) / BINS, n);
      track.appendChild(bar);
    }
    var marker = el("div", "conf-hist-marker");
    marker.style.left = (state.certaintyCutoff * 100) + "%";
    track.appendChild(marker);
    host.appendChild(track);
  }

  function initResultsPanel() {
    qs("#resultsList").addEventListener("click", function (e) {
      // Handle exclude toggle
      var btn = e.target.closest(".result-exclude-btn");
      if (btn && btn.dataset.eventId) {
        var evId = btn.dataset.eventId;
        var isExcluded = btn.dataset.excluded === "true";
        var endpoint = isExcluded ? "api/events/" + evId + "/include" : "api/events/" + evId + "/exclude";
        // Optimistic flip, reverted on failure so the row never lies about the
        // server's actual exclude state.
        var setExcluded = function (val) {
          var evts = state.taskEvents[state.selectedTaskId] || [];
          for (var i = 0; i < evts.length; i++) {
            if (evts[i].id === evId) { evts[i].excluded = val; break; }
          }
        };
        setExcluded(!isExcluded);
        renderResults();
        apiPut(endpoint).catch(function () {
          setExcluded(isExcluded);
          renderResults();
          showToast("Failed to update result");
        });
        return;
      }
      var row = e.target.closest(".result-row");
      if (!row || !row.dataset.timestamp) return;
      var ts = parseFloat(row.dataset.timestamp);
      if (isNaN(ts)) return;
      var task = state.selectedTaskId ? findTask(state.selectedTaskId) : null;
      if (task && task.participant && task.participant !== state.selectedParticipant) {
        selectParticipant(task.participant, ts);
        return;
      }
      // Set result overlay for spatial visualization
      var ri = parseInt(row.dataset.resultIndex, 10);
      var rData = (!isNaN(ri) && state.selectedTaskResults) ? state.selectedTaskResults[ri] : null;
      if (task && rData && task.type === "template" && rData.matches) {
        state.resultOverlay = { type: "template", data: rData };
      } else if (task && rData && task.type === "flow" && rData.flow_grid) {
        state.resultOverlay = { type: "flow", data: rData, region: taskRegionPixels(task) };
      } else {
        state.resultOverlay = null;
      }
      loadFrame(ts);
    });

    // Hover result rows to highlight matching scene markers on timeline
    qs("#resultsList").addEventListener("mouseover", function (e) {
      var row = e.target.closest(".result-row");
      if (!row) return;
      var task = state.selectedTaskId ? findTask(state.selectedTaskId) : null;
      if (!task || task.type !== "scene") return;
      var ri = parseInt(row.dataset.resultIndex, 10);
      var results = state.selectedTaskResults || [];
      if (isNaN(ri) || ri >= results.length) return;
      var sceneName = results[ri].scene_name;
      if (sceneName !== state.hoveredResultSceneName) {
        state.hoveredResultSceneName = sceneName;
        renderTimeline();
      }
    });

    qs("#resultsList").addEventListener("mouseleave", function () {
      if (state.hoveredResultSceneName !== null) {
        state.hoveredResultSceneName = null;
        renderTimeline();
      }
    });

    var showExcludedBtn = qs("#showExcludedBtn");
    function updateShowExcludedIcon() {
      var iconSpan = showExcludedBtn.querySelector(".rp-icon-btn-icon");
      iconSpan.classList.toggle("rp-icon-eye", state.showExcluded);
      iconSpan.classList.toggle("rp-icon-eye-slash", !state.showExcluded);
      showExcludedBtn.classList.toggle("active", state.showExcluded);
    }
    updateShowExcludedIcon();
    showExcludedBtn.addEventListener("click", function () {
      state.showExcluded = !state.showExcluded;
      updateShowExcludedIcon();
      renderResults();
    });
    attachHoverTooltip(showExcludedBtn, function () {
      return state.showExcluded ? "Hiding excluded results is off" : "Hiding excluded results is on";
    }, { align: "center" });

    function downloadEventsExport(format) {
      var url = "api/export/events?format=" + encodeURIComponent(format);
      if (!state.showExcluded) url += "&excluded=false";
      var a = document.createElement("a");
      a.href = url;
      var ext = format === "csv" ? "csv" : "json";
      a.download = "screenspace_events." + ext;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }

    var exportBtn = qs("#exportEventsBtn");
    var exportMenu = qs("#exportEventsMenu");
    if (exportBtn && exportMenu) {
      attachHoverTooltip(exportBtn, "Export events", { align: "center" });
      var closeExportMenu = function () {
        exportMenu.classList.add("hidden");
        exportBtn.setAttribute("aria-expanded", "false");
      };
      exportBtn.addEventListener("click", function (e) {
        e.stopPropagation();
        var open = exportMenu.classList.toggle("hidden");
        exportBtn.setAttribute("aria-expanded", open ? "false" : "true");
      });
      exportMenu.addEventListener("click", function (e) {
        var item = e.target.closest(".rp-export-item");
        if (!item) return;
        e.stopPropagation();
        downloadEventsExport(item.dataset.format);
        closeExportMenu();
      });
      document.addEventListener("click", function (e) {
        if (exportMenu.classList.contains("hidden")) return;
        if (e.target.closest("#exportEventsWrap")) return;
        closeExportMenu();
      });
    }

    var certaintySlider = qs("#certaintyCutoff");
    certaintySlider.addEventListener("input", function () {
      state.certaintyCutoff = parseInt(this.value, 10) / 100;
      renderResults();
    });
    attachHoverTooltip(certaintySlider, function () {
      return "Certainty threshold: " + certaintySlider.value + "%";
    }, { align: "center" });

    var exclBtn = qs("#excludeNonVisibleBtn");
    attachHoverTooltip(exclBtn, "Exclude results below the certainty threshold", { align: "center" });

    qs("#excludeNonVisibleBtn").addEventListener("click", function () {
      if (state.certaintyCutoff <= 0) {
        showToast("Set a certainty threshold first");
        return;
      }
      var task = state.selectedTaskId ? findTask(state.selectedTaskId) : null;
      if (!task) return; // button is hidden without a selected task

      function applyExclusion(events) {
        var idsToExclude = [];
        events.forEach(function (ev) {
          if (!ev.excluded && ev.confidence < state.certaintyCutoff) {
            idsToExclude.push(ev.id);
          }
        });
        if (idsToExclude.length === 0) {
          showToast("No events below threshold");
          return;
        }
        apiPut("api/events/bulk-exclude", { ids: idsToExclude }).then(function () {
          idsToExclude.forEach(function (id) {
            for (var i = 0; i < events.length; i++) {
              if (events[i].id === id) { events[i].excluded = true; break; }
            }
          });
          renderResults();
          renderTimeline();
          showToast("Excluded " + clipgenPluralUnit(idsToExclude.length, "event", "events"));
        });
      }

      // renderResults() draws from state.selectedTaskResults, which is populated
      // eagerly (live while running, and on completion). The events list, however,
      // is only fetched into state.taskEvents by loadAndShowResults — so when the
      // results panel is showing a streamed/just-completed task, the slider filters
      // the visible rows yet taskEvents is still empty. Fetch on demand so the cut
      // works without first navigating away from and back to the results page.
      var loaded = state.taskEvents[state.selectedTaskId];
      if (loaded && loaded.length) {
        applyExclusion(loaded);
        return;
      }
      // Capture the task id at request time and bail if the user has switched away
      // before the fetch resolves — otherwise the callback would write this task's
      // events into (and bulk-exclude against) whatever task is selected later.
      var fetchTaskId = state.selectedTaskId;
      apiGet("api/events?task_id=" + fetchTaskId).then(function (evData) {
        if (state.selectedTaskId !== fetchTaskId) return;
        var evts = (evData && evData.events) || [];
        state.taskEvents[fetchTaskId] = evts;
        renderResults(); // surface freshly-loaded events (histogram, excluded toggle, rows)
        if (!evts.length) {
          showToast(task.status === "running"
            ? "No events to exclude yet. Analysis still running"
            : "No events to exclude");
          return;
        }
        applyExclusion(evts);
      }).catch(function () { showToast("Failed to load events"); });
    });
  }

  function loadAndShowResults(taskId) {
    var resultsRequestVersion = ++state.resultsRequestVersion;
    var selectedTaskId = taskId;
    state.heatmapOverlayRequestVersion += 1;
    state.resultOverlay = null;
    state.heatmapOverlay = null;
    state.hoveredResultSceneName = null;
    state.certaintyCutoff = 0;
    var slider = qs("#certaintyCutoff");
    if (slider) slider.value = "0";
    // Surface a loading state immediately so the panel reads "Loading…" rather
    // than the idle "Click a task…" while the two fetches are in flight.
    state.resultsLoading = true;
    renderResults();
    apiGet("api/tasks/" + taskId + "/results")
      .then(function (data) {
        if (resultsRequestVersion !== state.resultsRequestVersion || state.selectedTaskId !== selectedTaskId) return null;
        state.selectedTaskResults = data.results;
        // Seed the shared per-task cache so the timeline draws these results and
        // _syncTaskResults appends further tails from here rather than re-fetching.
        // Redraw now: when the cache was empty (fresh completed task) nothing else
        // fires renderTimeline, so the markers would otherwise never paint.
        if (Array.isArray(data.results)) {
          state.taskResults[selectedTaskId] = data.results;
          renderTimeline();
        }
        return apiGet("api/events?task_id=" + selectedTaskId);
      })
      .then(function (evData) {
        if (!evData) return;
        if (resultsRequestVersion !== state.resultsRequestVersion || state.selectedTaskId !== selectedTaskId) return;
        state.taskEvents[selectedTaskId] = evData.events || [];
        state.resultsLoading = false;
        renderResults();
        renderTaskList();
        // Repaint the timeline now that excluded events are known, so excluded
        // markers get their dashed/faded styling.
        renderTimeline();
        updateResultsCrumb();
      })
      .catch(function () {
        // Only the current request clears the flag; a superseded load leaves the
        // newer one's spinner intact.
        if (resultsRequestVersion === state.resultsRequestVersion && state.selectedTaskId === selectedTaskId) {
          state.resultsLoading = false;
          renderResults();
        }
        showToast("Failed to load results");
      });
  }

  // Build a single result row element. Extracted from renderResults so the row
  // list can be rendered in lazy chunks (see RESULTS_RENDER_ALL / renderChunk):
  // the per-result event matching + certainty/excluded filtering stays in the
  // data pass, this just turns one (already-matched, already-kept) result into DOM.
  function buildResultRow(r, rIdx, matchedEvent, isExcluded, task) {
    var row = el("div", "result-row" + (isExcluded ? " excluded" : ""));
    row.dataset.resultIndex = rIdx;

    if (task.type === "color") {
      row.dataset.timestamp = r.start;
      row.appendChild(el("span", "result-timestamp", formatTime(r.start, { decimals: 1 }) + " \u2013 " + formatTime(r.end, { decimals: 1 })));
      row.appendChild(el("span", "result-detail", r.duration.toFixed(1) + "s"));
    } else if (task.type === "inactivity") {
      row.dataset.timestamp = r.start;
      row.appendChild(el("span", "result-timestamp", formatTime(r.start, { decimals: 1 }) + " \u2013 " + formatTime(r.end, { decimals: 1 })));
      row.appendChild(el("span", "result-detail", r.duration.toFixed(1) + "s"));
      row.appendChild(el("span", "result-score", "d:" + (r.avg_distance !== undefined ? r.avg_distance : "?")));
    } else if (task.type === "change") {
      row.dataset.timestamp = r.timestamp;
      row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
      row.appendChild(buildConfBar(Math.min(r.magnitude, 1), task.type));
      row.appendChild(el("span", "result-score", (r.magnitude * 100).toFixed(1) + "%"));
    } else if (task.type === "similarity") {
      row.dataset.timestamp = r.timestamp;
      row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
      row.appendChild(buildConfBar(r.score, task.type));
      row.appendChild(el("span", "result-score", (r.score * 100).toFixed(1) + "%"));
    } else if (task.type === "text") {
      row.dataset.timestamp = r.timestamp;
      row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
      row.appendChild(el("span", "result-detail", r.text_found || ""));
      row.appendChild(buildConfBar(r.confidence, task.type));
      row.appendChild(el("span", "result-score", (r.confidence * 100).toFixed(0) + "%"));
    } else if (task.type === "numbers") {
      row.dataset.timestamp = r.timestamp;
      row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
      row.appendChild(el("span", "result-detail", String(r.number_found)));
      if (r.confidence !== undefined) {
        row.appendChild(buildConfBar(r.confidence, task.type));
        row.appendChild(el("span", "result-score", (r.confidence * 100).toFixed(0) + "%"));
      }
    } else if (task.type === "template") {
      row.dataset.timestamp = r.timestamp;
      row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
      row.appendChild(buildConfBar(r.best_score, task.type));
      row.appendChild(el("span", "result-score", (r.best_score * 100).toFixed(1) + "%"));
      row.appendChild(el("span", "result-detail", r.match_count + " match" + (r.match_count !== 1 ? "es" : "")));
    } else if (task.type === "flow") {
      row.dataset.timestamp = r.timestamp;
      row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
      row.appendChild(buildConfBar(Math.min(r.magnitude / 20, 1), task.type));
      row.appendChild(el("span", "result-score", r.magnitude.toFixed(2)));
    } else if (task.type === "scene") {
      row.dataset.timestamp = r.timestamp;
      row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
      row.appendChild(el("span", "result-detail", r.scene_name));
      row.appendChild(buildConfBar(r.score, task.type));
      row.appendChild(el("span", "result-score", (r.score * 100).toFixed(1) + "%"));
    } else if (task.type === "boundary") {
      row.dataset.timestamp = r.timestamp;
      row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
      row.appendChild(buildConfBar(r._confidence !== undefined ? r._confidence : 0, task.type));
      row.appendChild(el("span", "result-score", "d:" + (r.distance !== undefined ? r.distance : "?")));
      // Scene label (Scene A/B/… — recurrence-aware for scene/hybrid metrics).
      if (r.scene_label) row.appendChild(el("span", "result-scene", r.scene_label));
    } else if (task.type === "multitool") {
      row.dataset.timestamp = r.timestamp;
      row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
      var badges = el("span", "result-detail multitool-badges");
      var stepDefs = (task.parameters && task.parameters.steps) || [];
      var types = r.tool_types || stepDefs.map(function (s) { return s.type; });
      types.forEach(function (t, i) {
        var step = stepDefs[i] || { type: t };
        if (i > 0) {
          var logic = (step.logic || "AND").toUpperCase();
          var sep = el("span", "multitool-step-logic" + (logic === "NOT" ? " logic-not" : ""), logic);
          badges.appendChild(sep);
        }
        var badge = el("span", "multitool-type-badge");
        badge.style.color = taskTypeColor(t);
        var paramStr = formatMultitoolStepParams(step);
        badge.title = t + (paramStr ? ": " + paramStr : "");
        var icon = buildTypeIcon(t);
        if (icon) badge.appendChild(icon);
        if (paramStr) badge.appendChild(el("span", "multitool-step-params", paramStr));
        badges.appendChild(badge);
      });
      row.appendChild(badges);
      row.appendChild(el("span", "result-score", ((r.min_confidence || 0) * 100).toFixed(1) + "%"));
    }

    if (matchedEvent) {
      var btn = el("button", "result-exclude-btn");
      btn.dataset.eventId = matchedEvent.id;
      btn.dataset.excluded = isExcluded ? "true" : "false";
      btn.title = isExcluded ? "Include event" : "Exclude event";
      var exIcon = isExcluded ? iconSpan("x-mark", "ss-icon--sm") : iconSpan("check", "ss-icon--sm");
      btn.appendChild(exIcon);
      row.appendChild(btn);
    }
    return row;
  }

  // Task types whose detections carry a confidence score (drive the certainty
  // slider + histogram). Shared by the full render's filter and the streaming
  // append fast path so the two never drift.
  var CONF_TASK_TYPES = {
    change: 1, similarity: 1, text: 1, numbers: 1, template: 1,
    scene: 1, flow: 1, multitool: 1, inactivity: 1, boundary: 1,
  };

  function taskHasConfidence(task) {
    return !!(task && CONF_TASK_TYPES[task.type]);
  }

  // Confidence value for a result under its task's detector (null when the
  // detector has none). Only meaningful when taskHasConfidence(task).
  function resultConfidence(r, task) {
    if (task.type === "change") return r.magnitude;
    if (task.type === "similarity") return r.score;
    if (task.type === "text") return r.confidence;
    if (task.type === "numbers") return r.confidence;
    if (task.type === "template") return r.best_score;
    if (task.type === "flow") return Math.min(r.magnitude / 10, 1);
    if (task.type === "scene") return r.score;
    if (task.type === "multitool") return r.min_confidence;
    if (task.type === "inactivity") return Math.min((r.duration || 0) / 30, 1);
    if (task.type === "boundary") return r._confidence;
    return null;
  }

  // Signature of what renderResults last painted, so a streaming push that only
  // grew the tail can append the new rows instead of a full rebuild (see the
  // fast path in renderResults). null after any non-array/early-return render.
  var _lastResultsSig = null;

  function resultsSignature(results, task) {
    return {
      taskId: state.selectedTaskId,
      cutoff: state.certaintyCutoff,
      showExcluded: state.showExcluded,
      eventsLen: (state.taskEvents[state.selectedTaskId] || []).length,
      heatmapSig: (task.heatmap || "") + "|" + (task.heatmap_gif || "") + "|" + (task.heatmap_rolling_gif || ""),
      fastScan: (task.parameters || {}).scan_mode === "fast",
      rawLen: Array.isArray(results) ? results.length : -1,
    };
  }

  // Render all rows inline below this many; above it, rows stream in chunks via
  // an IntersectionObserver so a 500+ result task doesn't build thousands of DOM
  // nodes (and reflow) in one synchronous pass. Variable row heights (wrapping
  // text / multitool badges) rule out fixed-height windowing, so we grow the
  // list incrementally instead.
  var RESULTS_RENDER_ALL = 150;
  var RESULTS_CHUNK = 120;

  function renderResults() {
    var container = qs("#resultsList");
    var prevResultsScrollTop = container.scrollTop;
    // Tear down any lazy-load observer from a previous render — its sentinel is
    // about to be wiped with container.innerHTML, and a stale observer would
    // keep a detached node alive.
    if (state.resultsLazyObserver) {
      state.resultsLazyObserver.disconnect();
      state.resultsLazyObserver = null;
    }
    var countEl = qs("#resultCount") || { textContent: "" };
    var actionsEl = qs("#resultsActions");
    var results = state.selectedTaskResults;
    var task = state.selectedTaskId ? findTask(state.selectedTaskId) : null;

    // Streaming-append fast path: while the selected task is still running, each
    // push only grows state.selectedTaskResults at the tail (events aren't fetched
    // until completion, so there's no per-timestamp event matching / exclude
    // buttons to reconcile). Append just the new rows instead of wiping + rebuilding
    // the whole list and re-creating the IntersectionObserver every ~500ms — the
    // latter is O(n) per push, O(n²) over a scan. Any change to the task, filters,
    // events, or heatmap breaks the signature match and falls through to the full
    // render (which, while running, renders eagerly with no observer — see below).
    if (!state.resultsLoading && task && task.status === "running" && Array.isArray(results)) {
      var sig = resultsSignature(results, task);
      var prev = _lastResultsSig;
      if (
        prev
        && prev.taskId === sig.taskId
        && prev.cutoff === sig.cutoff
        && prev.showExcluded === sig.showExcluded
        && prev.fastScan === sig.fastScan
        && prev.heatmapSig === sig.heatmapSig
        && prev.eventsLen === 0 && sig.eventsLen === 0
        && sig.rawLen >= prev.rawLen
      ) {
        if (sig.rawLen > prev.rawLen) {
          var hasConfFast = taskHasConfidence(task);
          var appendFrag = document.createDocumentFragment();
          for (var ai = prev.rawLen; ai < results.length; ai++) {
            var ar = results[ai];
            if (hasConfFast && state.certaintyCutoff > 0) {
              var acv = resultConfidence(ar, task);
              if (acv !== null && acv < state.certaintyCutoff) continue;
            }
            appendFrag.appendChild(buildResultRow(ar, ai, null, false, task));
          }
          container.appendChild(appendFrag);
          countEl.textContent = "(" + results.length + ")";
        }
        _lastResultsSig = sig;
        return;
      }
    }
    // Falling through to a full render — clear the signature so the early-return
    // branches below (loading/timelapse/non-array) leave it null; the successful
    // array render re-stamps it.
    _lastResultsSig = null;

    // Manage fast scan label — between panel-header and resultsList
    var fastLabel = qs("#fastScanLabel");
    if (!fastLabel) {
      fastLabel = el("div", "fast-scan-label hidden");
      fastLabel.id = "fastScanLabel";
      container.parentNode.insertBefore(fastLabel, container);
    }

    // While a results fetch is in flight, show a loading state even if stale
    // results from the previous task are still in state.selectedTaskResults.
    if (state.resultsLoading || !results || !task) {
      var emptyMsg = state.resultsLoading ? "Loading results…" : "Click a task to view results.";
      container.innerHTML = '<div class="panel-empty">' + emptyMsg + "</div>";
      countEl.textContent = "";
      actionsEl.classList.add("hidden");
      fastLabel.classList.add("hidden");
      var emptyHist = qs("#confHistogram");
      if (emptyHist) emptyHist.classList.add("hidden");
      return;
    }

    actionsEl.classList.remove("hidden");

    if ((task.parameters || {}).scan_mode === "fast") {
      fastLabel.classList.remove("hidden");
      fastLabel.innerHTML = "";
      var fIcon = el("span", "fast-scan-label-icon");
      applyIconMask(fIcon, "chevron-double-right", "/screenspace/icons/");
      fastLabel.appendChild(fIcon);
      fastLabel.appendChild(document.createTextNode("Fast scan results"));
      var rerunBtn = el("button", "btn btn-small fast-scan-rerun-btn", "Re-Run Normal");
      (function (t) {
        rerunBtn.addEventListener("click", function () {
          var params = {};
          Object.keys(t.parameters || {}).forEach(function (k) { params[k] = t.parameters[k]; });
          delete params.scan_mode;
          var body = {
            type: t.type,
            participant: t.participant,
            region: t.region || "",
            parameters: params,
          };
          if (t.region_ref) body.region_ref = t.region_ref;
          apiPost("api/tasks", body).then(function (data) {
            if (data.ok) {
              state.tasks.push(data.task);
              renderTaskList();
              startSSE();
              showToast("Re-queued in Normal mode");
            } else {
              showToast(data.error || "Failed to re-queue task");
            }
          });
        });
      })(task);
      fastLabel.appendChild(rerunBtn);
    } else {
      fastLabel.classList.add("hidden");
    }

    // Show/hide certainty controls based on whether the tool has confidence scores
    var hasConf = taskHasConfidence(task);
    var certWrap = qs("#certaintyCutoffWrap");
    var exclBtn = qs("#excludeNonVisibleBtn");
    if (certWrap) certWrap.classList.toggle("hidden", !hasConf);
    if (exclBtn) exclBtn.classList.toggle("hidden", !hasConf);

    // Confidence histogram — shown exactly when the cutoff slider is (same
    // hasConf gate), and only when there are events to bucket.
    var histHost = qs("#confHistogram");
    var histEvents = state.taskEvents[state.selectedTaskId] || [];
    var showHist = state.showConfidenceHistogram && !!hasConf && histEvents.length > 0;
    if (histHost) histHost.classList.toggle("hidden", !showHist);
    if (showHist) renderConfidenceHistogram(histEvents, task.type);

    // Timelapse: single file result
    if (task.type === "timelapse" && typeof results === "string") {
      countEl.textContent = "";
      container.innerHTML = "";
      var wrapper = el("div", "timelapse-result");
      var ext = results.split(".").pop().toLowerCase();
      var filename = results.split("/").pop();
      if (ext === "gif") {
        var img = document.createElement("img");
        img.decoding = "async";
        img.src = "media/" + filename;
        wrapper.appendChild(img);
      } else {
        var vid = document.createElement("video");
        vid.src = "media/" + filename;
        vid.controls = true;
        vid.muted = true;
        wrapper.appendChild(vid);
      }
      container.innerHTML = "";
      container.appendChild(wrapper);
      return;
    }

    if (!Array.isArray(results)) {
      container.innerHTML = '<div class="panel-empty">No results.</div>';
      countEl.textContent = "";
      return;
    }

    var events = state.taskEvents[state.selectedTaskId] || [];
    var eventsByTs = {};
    events.forEach(function (ev) {
      var key = ev.time_in.toFixed(2);
      if (!eventsByTs[key]) eventsByTs[key] = [];
      eventsByTs[key].push(ev);
    });

    // For color results (spans), build a consumed-index tracker per timestamp
    var eventTsIndex = {};

    var showToggle = qs("#showExcludedBtn");
    if (showToggle) showToggle.classList.toggle("hidden", events.length === 0);

    countEl.textContent = "(" + results.length + ")";
    container.innerHTML = "";

    // Heatmap artifact display (template, flow, change)
    if (task.heatmap && (task.type === "template" || task.type === "flow" || task.type === "change")) {
      var heatmapSection = el("div", "heatmap-result");
      var heatmapLabel = el("div", "heatmap-label");
      // Clickable title collapses the section to cut visual noise (state persists
      // across results re-renders via state.heatmapCollapsed).
      var collapseToggle = el("button", "heatmap-collapse-toggle");
      collapseToggle.appendChild(el("span", "heatmap-collapse-chevron"));
      collapseToggle.appendChild(document.createTextNode("Detection Heatmap"));
      collapseToggle.addEventListener("click", function () {
        state.heatmapCollapsed = !state.heatmapCollapsed;
        heatmapSection.classList.toggle("collapsed", !!state.heatmapCollapsed);
      });
      heatmapLabel.appendChild(collapseToggle);
      var overlayBtn = el("button", "btn btn-small", state.heatmapOverlay ? "Hide Overlay" : "Overlay on Frame");
      overlayBtn.addEventListener("click", function () {
        if (state.heatmapOverlay) {
          state.heatmapOverlayRequestVersion += 1;
          state.heatmapOverlay = null;
          overlayBtn.textContent = "Overlay on Frame";
          renderOverlay();
        } else {
          // Overlay always uses the static heatmap image, never the animations.
          var overlaySrc = "media/" + task.heatmap;
          var overlayRequestVersion = ++state.heatmapOverlayRequestVersion;
          state.heatmapOverlay = {
            src: overlaySrc,
            type: task.type,
            region_coords: taskRegionPixels(task),
          };
          overlayBtn.textContent = "Hide Overlay";
          var hmImg = new Image();
          hmImg.onload = function () {
            if (
              overlayRequestVersion === state.heatmapOverlayRequestVersion
              && state.heatmapOverlay
              && state.heatmapOverlay.src === overlaySrc
            ) {
              state.heatmapOverlay._img = hmImg;
              renderOverlay();
            }
          };
          hmImg.src = overlaySrc;
        }
      });
      heatmapLabel.appendChild(overlayBtn);
      heatmapSection.appendChild(heatmapLabel);

      // Show every generated mode side by side as a small thumbnail strip.
      var heatmapViews = [
        { label: "Static", src: task.heatmap, alt: "Detection heatmap", animated: false },
      ];
      if (task.heatmap_gif) {
        heatmapViews.push({
          label: "Accumulation",
          src: task.heatmap_gif,
          alt: "Heatmap accumulation animation",
          animated: true,
        });
      }
      if (task.heatmap_rolling_gif) {
        heatmapViews.push({
          label: "Rolling Window",
          src: task.heatmap_rolling_gif,
          alt: "Rolling-window heatmap animation",
          animated: true,
        });
      }

      var heatmapStrip = el("div", "heatmap-strip");
      heatmapViews.forEach(function (view) {
        var thumb = el("div", "heatmap-thumb");
        var media = el("div", "heatmap-thumb-media");
        var img = document.createElement("img");
        img.decoding = "async";
        img.src = "media/" + view.src;
        img.alt = view.alt;
        media.appendChild(img);
        if (view.animated) {
          // GIFs can't be paused natively: freeze the current frame onto a
          // canvas to pause, restore the gif src to resume from the start.
          media.classList.add("heatmap-thumb-animated");
          media.appendChild(el("span", "heatmap-thumb-glyph"));
          var gifSrc = "media/" + view.src;
          var frozen = null;
          media.addEventListener("click", function () {
            if (!frozen) {
              var canvas = el("canvas", "heatmap-thumb-frozen");
              canvas.width = img.naturalWidth || img.clientWidth || 1;
              canvas.height = img.naturalHeight || img.clientHeight || 1;
              try {
                canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
              } catch (e) {
                return; // frame not ready yet — leave it playing
              }
              img.classList.add("hidden");
              media.insertBefore(canvas, img);
              frozen = canvas;
              media.classList.add("paused");
            } else {
              media.removeChild(frozen);
              frozen = null;
              img.classList.remove("hidden");
              img.src = gifSrc; // restart the animation
              media.classList.remove("paused");
            }
          });
        }
        thumb.appendChild(media);
        thumb.appendChild(el("div", "heatmap-thumb-label", view.label));
        heatmapStrip.appendChild(thumb);
      });
      heatmapSection.appendChild(heatmapStrip);

      // Restore persisted collapsed state for this render.
      heatmapSection.classList.toggle("collapsed", !!state.heatmapCollapsed);

      container.appendChild(heatmapSection);
    }

    // Data pass: pair each result to its event (sequential per-timestamp
    // consumption), apply certainty + excluded filtering, and collect the kept
    // rows. Cheap (no DOM) so it stays a single pass over every result.
    var visibleRows = [];
    results.forEach(function (r, rIdx) {
      // Find matching event for this result
      var ts = r.timestamp !== undefined ? r.timestamp : r.start;
      var tsKey = ts !== undefined ? ts.toFixed(2) : null;
      var matchedEvent = null;
      if (tsKey && eventsByTs[tsKey]) {
        var idx = eventTsIndex[tsKey] || 0;
        if (idx < eventsByTs[tsKey].length) {
          matchedEvent = eventsByTs[tsKey][idx];
          eventTsIndex[tsKey] = idx + 1;
        }
      }

      // Certainty filtering
      if (hasConf && state.certaintyCutoff > 0) {
        var confValue = resultConfidence(r, task);
        if (confValue !== null && confValue < state.certaintyCutoff) return;
      }

      var isExcluded = matchedEvent && matchedEvent.excluded;
      if (isExcluded && !state.showExcluded) return;

      visibleRows.push({ r: r, rIdx: rIdx, matchedEvent: matchedEvent, isExcluded: isExcluded });
    });

    // Render pass: append rows a chunk at a time. renderChunk returns true while
    // more rows remain, so callers can keep pulling.
    var rendered = 0;
    function renderChunk() {
      var endIdx = Math.min(rendered + RESULTS_CHUNK, visibleRows.length);
      var frag = document.createDocumentFragment();
      for (var i = rendered; i < endIdx; i++) {
        var d = visibleRows[i];
        frag.appendChild(buildResultRow(d.r, d.rIdx, d.matchedEvent, d.isExcluded, task));
      }
      container.appendChild(frag);
      rendered = endIdx;
      return rendered < visibleRows.length;
    }

    // Stamp what this render covers so a subsequent streaming push can append the
    // tail instead of rebuilding (see the fast path at the top of renderResults).
    _lastResultsSig = resultsSignature(results, task);

    // Small lists (and any running task) render in full with no observer \u2014 the
    // observer's sentinel/closure can't coexist with the fast path's tail appends,
    // and a live scan grows the DOM a few rows per push rather than all at once.
    if (visibleRows.length <= RESULTS_RENDER_ALL || task.status === "running") {
      while (renderChunk()) { /* render everything */ }
      container.scrollTop = prevResultsScrollTop;
      return;
    }

    // Large list: render the first chunk, then enough more to cover the prior
    // scroll position (so re-renders during live streaming don't jump to the
    // top), then lazy-load the rest as the user scrolls toward the bottom.
    renderChunk();
    var coverTarget = prevResultsScrollTop + container.clientHeight;
    while (rendered < visibleRows.length && container.scrollHeight < coverTarget) renderChunk();
    container.scrollTop = prevResultsScrollTop;

    if (rendered < visibleRows.length) {
      var sentinel = el("div", "results-lazy-sentinel");
      container.appendChild(sentinel);
      var io = new IntersectionObserver(function (entries) {
        if (!entries[0].isIntersecting) return;
        var more = renderChunk();
        if (more) {
          container.appendChild(sentinel); // keep the sentinel below freshly added rows
        } else {
          io.disconnect();
          if (sentinel.parentNode) sentinel.parentNode.removeChild(sentinel);
          state.resultsLazyObserver = null;
        }
      }, { root: container, rootMargin: "300px" });
      io.observe(sentinel);
      state.resultsLazyObserver = io;
    }
  }

  // ---- Satellite interface (published back to window.ClipgenScreenspace) ----
  // The hub keeps same-named thin delegators for the entry points it calls
  // itself (initResultsPanel from init, renderResults on histogram-toggle); the
  // tasks satellite reaches renderResults / loadAndShowResults via these handles.
  SS.initResultsPanel = initResultsPanel;
  SS.loadAndShowResults = loadAndShowResults;
  SS.renderResults = renderResults;
})();
