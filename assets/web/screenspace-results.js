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

  // Confidence bar; opacity ramps 0.4 to 1.0 with confidence.
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

  // 10-bin confidence histogram with a cutoff marker; rebuilt on every renderResults() call.
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
        // Optimistic flip, reverted on failure.
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
      if (task && rData && (task.type === "template" || task.type === "shape") && rData.matches) {
        state.resultOverlay = { type: task.type, data: rData };
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
      var ext = format === "csv" ? "csv" : "json";
      clipgenSaveFromUrl(url, "screenspace_events." + ext, function (path, err) {
        if (err) showToast("Export failed: " + err.message);
      });
    }

    var exportBtn = qs("#exportEventsBtn");
    var exportMenu = qs("#exportEventsMenu");
    if (exportBtn && exportMenu) {
      attachHoverTooltip(exportBtn, "Download the detected events as a CSV or JSON file", { align: "center" });
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

      // loadAndShowResults alone fills taskEvents; a streamed task may have none yet, so fetch here.
      var loaded = state.taskEvents[state.selectedTaskId];
      if (loaded && loaded.length) {
        applyExclusion(loaded);
        return;
      }
      // Bail if the user switched tasks before the fetch resolved.
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
    // Show "Loading…" immediately while the two fetches run.
    state.resultsLoading = true;
    renderResults();
    apiGet("api/tasks/" + taskId + "/results")
      .then(function (data) {
        if (resultsRequestVersion !== state.resultsRequestVersion || state.selectedTaskId !== selectedTaskId) return null;
        state.selectedTaskResults = data.results;
        // Seed the shared cache for the timeline and _syncTaskResults; redraw now, nothing else will.
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
        // Repaint so excluded markers get their dashed/faded styling.
        renderTimeline();
        updateResultsCrumb();
      })
      .catch(function () {
        // Only the current request clears the flag; superseded loads leave the spinner.
        if (resultsRequestVersion === state.resultsRequestVersion && state.selectedTaskId === selectedTaskId) {
          state.resultsLoading = false;
          renderResults();
        }
        showToast("Failed to load results");
      });
  }

  // One result row as DOM; matching and filtering happen in the caller's data pass.
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
    } else if (task.type === "template" || task.type === "shape") {
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
    } else if (task.type === "attention") {
      // Confirmed focus shifts; Δ is the normalized jump of the attention peak.
      row.dataset.timestamp = r.timestamp;
      row.appendChild(el("span", "result-timestamp", formatTime(r.timestamp, { decimals: 1 })));
      row.appendChild(buildConfBar(r._confidence !== undefined ? r._confidence : 0, task.type));
      row.appendChild(el("span", "result-score", "Δ" + (r.shift_distance !== undefined ? r.shift_distance.toFixed(2) : "?")));
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

  // Types with confidence scores; shared by the full render and the append fast path.
  var CONF_TASK_TYPES = {
    change: 1, similarity: 1, text: 1, numbers: 1, template: 1, shape: 1,
    scene: 1, flow: 1, multitool: 1, inactivity: 1, boundary: 1,
    attention: 1,
  };

  function taskHasConfidence(task) {
    return !!(task && CONF_TASK_TYPES[task.type]);
  }

  // Per-detector confidence; null when the detector has none.
  function resultConfidence(r, task) {
    if (task.type === "change") return r.magnitude;
    if (task.type === "similarity") return r.score;
    if (task.type === "text") return r.confidence;
    if (task.type === "numbers") return r.confidence;
    if (task.type === "template") return r.best_score;
    if (task.type === "shape") return r.best_score;
    if (task.type === "flow") return Math.min(r.magnitude / 10, 1);
    if (task.type === "scene") return r.score;
    if (task.type === "multitool") return r.min_confidence;
    if (task.type === "inactivity") return Math.min((r.duration || 0) / 30, 1);
    if (task.type === "boundary") return r._confidence;
    if (task.type === "attention") return r._confidence;
    return null;
  }

  // Last painted signature so a streaming push can append; null after early returns.
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

  // Placeholder for a 404'd heatmap image; the warn names the dead URL.
  function markHeatmapMissing(media, url) {
    if (typeof console !== "undefined" && console.warn) {
      console.warn("clipgen: heatmap image failed to load: " + url);
    }
    media.innerHTML = "";
    media.style.backgroundImage = "";
    // Clear the sprite aspect ratio too, or the box stays tall and empty.
    media.style.aspectRatio = "";
    media.classList.remove("heatmap-thumb-animated", "paused");
    media.classList.add("heatmap-thumb-missing");
    media.appendChild(el("span", "heatmap-thumb-missing-text", "Image unavailable"));
  }

  // Plain <img>: static PNG, playing GIF, or sprite-less fallback. Load errors here are final.
  function fillHeatmapImage(media, view) {
    var img = document.createElement("img");
    img.decoding = "async";
    img.src = "media/" + view.src;
    img.alt = view.alt;
    img.addEventListener("error", function () {
      markHeatmapMissing(media, img.src);
    });
    media.appendChild(img);
    return img;
  }

  // Animated views rest on a sprite sheet; play loads the GIF. Sprite-less views just loop.
  function buildHeatmapMedia(view) {
    var media = el("div", "heatmap-thumb-media");
    var sprite = view.sprite;
    if (!view.key || !sprite || !sprite.frames) {
      fillHeatmapImage(media, view);
      return media;
    }

    media.classList.add("heatmap-thumb-animated");
    // Pin the aspect ratio from sprite geometry so the sprite/GIF swap can't resize the box.
    media.style.aspectRatio = sprite.w + " / " + sprite.h;
    var glyph = el("span", "heatmap-thumb-glyph");
    var progress = el("div", "heatmap-scrub-progress");
    var progressFill = el("div", "heatmap-scrub-progress-fill");
    progress.appendChild(progressFill);
    var detach = null;
    var spriteBroken = false;
    // Sprite sheets are rendered on demand from the GIF, not stored beside it.
    var spriteUrl = "api/heatmap-sprite/" + encodeURIComponent(view.src) +
      "?cols=" + sprite.cols;

    function showPaused() {
      if (detach) return;
      // A broken sprite would paint an empty box; fall back instead.
      if (spriteBroken) {
        fallBackToPlainGif();
        return;
      }
      media.innerHTML = "";
      media.classList.add("paused");
      media.style.backgroundImage = 'url("' + spriteUrl + '")';
      media.appendChild(glyph);
      media.appendChild(progress);
      progressFill.style.width = "100%";
      detach = window.clipgenCardScrubber.attach(media, {
        spriteData: {
          cols: sprite.cols,
          rows: sprite.rows,
          frameCount: sprite.frames,
          interval: 0,
        },
        // Rest on the finished accumulation rather than the empty first frame.
        restFrame: sprite.frames - 1,
        onScrub: function (frac) {
          progressFill.style.width = (frac === null ? 1 : frac) * 100 + "%";
        },
      });
    }

    function showPlaying() {
      if (detach) {
        detach();
        detach = null;
      }
      media.innerHTML = "";
      media.classList.remove("paused");
      media.style.backgroundImage = "";
      fillHeatmapImage(media, view);
      media.appendChild(glyph);
    }

    // Plain GIF without scrub; a broken sprite must never blank a thumb whose GIF works.
    function fallBackToPlainGif() {
      if (detach) {
        detach();
        detach = null;
      }
      media.removeEventListener("click", onToggle);
      media.classList.remove("paused");
      media.style.backgroundImage = "";
      media.innerHTML = "";
      fillHeatmapImage(media, view);
    }

    // Keyed by task too: same-task re-renders keep the play state; a new task starts paused.
    var playKey = state.selectedTaskId + "|" + view.key;
    function onToggle() {
      var playing = !state.heatmapPlaying[playKey];
      state.heatmapPlaying[playKey] = playing;
      if (playing) showPlaying();
      else showPaused();
    }
    media.addEventListener("click", onToggle);

    // A failing background-image fires no error event, so probe it separately.
    var probe = new Image();
    probe.addEventListener("error", function () {
      if (typeof console !== "undefined" && console.warn) {
        console.warn("clipgen: heatmap sprite failed to load: " + spriteUrl);
      }
      spriteBroken = true;
      // The GIF may already be playing; leave it and let the next pause fall back.
      if (media.classList.contains("paused")) fallBackToPlainGif();
    });
    probe.src = spriteUrl;

    if (state.heatmapPlaying[playKey]) showPlaying();
    else showPaused();
    return media;
  }

  // Above this, rows stream in IntersectionObserver chunks; variable row heights rule out windowing.
  var RESULTS_RENDER_ALL = 150;
  var RESULTS_CHUNK = 120;

  function renderResults() {
    return clipgenPerf.span("screenspace.renderResults", renderResultsImpl);
  }

  function renderResultsImpl() {
    var container = qs("#resultsList");
    var prevResultsScrollTop = container.scrollTop;
    // Disconnect the old lazy observer; its sentinel is about to be wiped.
    if (state.resultsLazyObserver) {
      state.resultsLazyObserver.disconnect();
      state.resultsLazyObserver = null;
    }
    var countEl = qs("#resultCount") || { textContent: "" };
    var actionsEl = qs("#resultsActions");
    var results = state.selectedTaskResults;
    var task = state.selectedTaskId ? findTask(state.selectedTaskId) : null;

    // Fast path: a running task only grows the tail; append instead of rebuilding every push.
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
    // Full render: clear so early returns leave it null; the array render re-stamps it.
    _lastResultsSig = null;

    // Manage fast scan label — between panel-header and resultsList
    var fastLabel = qs("#fastScanLabel");
    if (!fastLabel) {
      fastLabel = el("div", "fast-scan-label hidden");
      fastLabel.id = "fastScanLabel";
      container.parentNode.insertBefore(fastLabel, container);
    }

    // Loading wins over stale results left from the previous task.
    if (state.resultsLoading || !results || !task) {
      var emptyMsg = state.resultsLoading ? "Loading results…" : "Click a task to view results.";
      var emptyCls = "panel-empty" + (state.resultsLoading ? " cg-shimmer" : "");
      container.innerHTML = '<div class="' + emptyCls + '">' + emptyMsg + "</div>";
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

    // Histogram follows the slider's hasConf gate and needs events to bucket.
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

    // Heatmap artifact display (template, shape, flow, change, attention)
    if (task.heatmap && (task.type === "template" || task.type === "shape" || task.type === "flow" || task.type === "change" || task.type === "attention")) {
      var heatmapSection = el("div", "heatmap-result");
      var heatmapLabel = el("div", "heatmap-label");
      // Collapsible title; state.heatmapCollapsed persists across re-renders.
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
        { label: "Static", src: task.heatmap, alt: "Detection heatmap", key: null },
      ];
      if (task.heatmap_gif) {
        heatmapViews.push({
          label: "Accumulation",
          src: task.heatmap_gif,
          alt: "Heatmap accumulation animation",
          key: "heatmap_gif",
          sprite: task.heatmap_gif_sprite,
        });
      }
      if (task.heatmap_rolling_gif) {
        heatmapViews.push({
          label: "Rolling Window",
          src: task.heatmap_rolling_gif,
          alt: "Rolling-window heatmap animation",
          key: "heatmap_rolling_gif",
          sprite: task.heatmap_rolling_gif_sprite,
        });
      }

      // The rebuild orphans attached thumbs; drop dead entries first.
      if (window.clipgenCardScrubber) window.clipgenCardScrubber.detachStale();

      var heatmapStrip = el("div", "heatmap-strip");
      heatmapViews.forEach(function (view) {
        var thumb = el("div", "heatmap-thumb");
        thumb.appendChild(buildHeatmapMedia(view));
        thumb.appendChild(el("div", "heatmap-thumb-label", view.label));
        heatmapStrip.appendChild(thumb);
      });
      heatmapSection.appendChild(heatmapStrip);

      // Restore persisted collapsed state for this render.
      heatmapSection.classList.toggle("collapsed", !!state.heatmapCollapsed);

      container.appendChild(heatmapSection);
    }

    // Data pass (no DOM): match events per timestamp, apply filters, collect kept rows.
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

    // Render pass; renderChunk returns true while rows remain.
    var rendered = 0;
    function renderChunk() {
      return clipgenPerf.span("screenspace.renderChunk", function () {
        var endIdx = Math.min(rendered + RESULTS_CHUNK, visibleRows.length);
        var frag = document.createDocumentFragment();
        for (var i = rendered; i < endIdx; i++) {
          var d = visibleRows[i];
          frag.appendChild(buildResultRow(d.r, d.rIdx, d.matchedEvent, d.isExcluded, task));
        }
        container.appendChild(frag);
        rendered = endIdx;
        return rendered < visibleRows.length;
      });
    }

    // Stamp this render so a streaming push can append the tail (fast path above).
    _lastResultsSig = resultsSignature(results, task);

    // Small lists and running tasks render fully; the observer's sentinel can't coexist with tail appends.
    if (visibleRows.length <= RESULTS_RENDER_ALL || task.status === "running") {
      while (renderChunk()) { /* render everything */ }
      container.scrollTop = prevResultsScrollTop;
      return;
    }

    // Large list: first chunk, then enough to cover the prior scroll position, then lazy-load.
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

  // ---- Satellite interface (window.ClipgenScreenspace) ----
  // Consumed by hub delegators and the tasks satellite.
  SS.initResultsPanel = initResultsPanel;
  SS.loadAndShowResults = loadAndShowResults;
  SS.renderResults = renderResults;
})();
