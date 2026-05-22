/* clipgen Transcripts page.
 *
 * Editor for per-participant Whisper transcripts, plus cross-references into
 * Screenspace events and the source spreadsheet. Two long-running side flows
 * piggyback on the same `state`:
 *
 *   - Transcription warmup: a single `tryPostTranscriptionWarmup()` post that
 *     asks the backend to preload the Whisper model. `_transcriptionWarmupPosted`
 *     guards it so we never double-post per page load.
 *   - Summary / citations: Ollama-generated; `_summaryPollTimer` and
 *     `_citationsPollTimer` poll the backend until the result lands or the
 *     user navigates away.
 */

(function () {
  "use strict";

  var SEARCH_DEBOUNCE = 300;

  // Speed cycle for the custom video controls. Transcript-friendly steps so
  // users can slow review or skim faster without large jumps.
  var VIDEO_SPEEDS = [0.75, 1, 1.25, 1.5, 2];

  var state = {
    participants: [],
    selectedParticipant: null,
    segments: [],
    corrections: [],
    tasks: [],
    searchQuery: "",
    searchResults: null,
    activeSegmentIndex: -1,
    editingTextEl: null,
    pollTimer: null,
    lastMarkCategory: "bookmark",
    streamingParticipant: null,
    ssEvents: [],
    ssEventsLoaded: false,
    sheetRows: [],
    sheetParticipants: [],
    sheetLoaded: false,
    xrefPollTimer: null,
    xrefEligible: false,
    xrefIndex: { eventsByParticipant: {}, sheetByParticipant: {} },
    tooltipsEnabled: true,
    summaryCollapsed: false,
    summaryEditing: false,
    summaryText: "",
    summaryCitations: null,
    citationsGenerating: false,
    transcribePrewarm: "queue_open",
    modelStatus: null,
    modelFailSince: 0,
    videoPlaying: false,
    videoMuted: false,
    videoPlaybackRate: 1,
    ccEnabled: false,
    pipActive: false,
    pipEnabled: true,
    videoCollapsed: false,
  };

  var _transcriptionWarmupPosted = false;
  var _modelHintPollTimer = null;
  var _hadActiveTranscriptionLastPoll = false;
  var MODEL_FAIL_GRACE_MS = 10000;

  // Bumped on every participant switch so per-participant fetches that resolve
  // late (`loadTranscript`, `loadSummary`, summary/citations polls) can detect
  // they're stale and bail before clobbering the active participant's UI.
  var _participantReqVer = 0;

  // ---- Helpers (showToast: 2500ms hide; shared default in utils.js is 3000ms) ----

  var _utilsShowToast = window.showToast;
  function showToast(msg) {
    _utilsShowToast(msg, { durationMs: 2500 });
  }

  // ---- Nav links ----

  function checkNavLinks() {
    // TODO: skips r.ok; silent on HTTP errors.
    fetch("../api/status").then(function (r) { return r.json(); }).then(function (data) {
      if (data.screenspace || data.studio) {
        state.xrefEligible = true;
        startXrefPolling();
      }
    }).catch(function () {});
  }

  // ---- Cross-reference data ----

  function startXrefPolling() {
    if (!state.xrefEligible || state.xrefPollTimer) return;
    loadCrossRefData();
    state.xrefPollTimer = setInterval(loadCrossRefData, 30000);
  }

  function stopXrefPolling() {
    if (state.xrefPollTimer) {
      clearInterval(state.xrefPollTimer);
      state.xrefPollTimer = null;
    }
  }

  function loadCrossRefData() {
    // TODO: skips r.ok; silent on HTTP errors (polling every 30s, so caller tolerates it).
    fetch("../screenspace/api/events?excluded=false")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data.ok) {
          state.ssEvents = data.events || [];
          state.ssEventsLoaded = true;
          _buildEventsIndex();
          if (state.searchResults) renderSearchResults(state.searchResults);
        }
      })
      .catch(function () {});

    // TODO: skips r.ok (same pattern).
    fetch("../studio/api/sheet")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data.ok) {
          clipgenApplyConfig(data.config);
          state.sheetRows = data.rows || [];
          state.sheetParticipants = data.participants || [];
          state.sheetLoaded = true;
          _buildSheetIndex();
          if (state.searchResults) renderSearchResults(state.searchResults);
        }
      })
      .catch(function () {});
  }

  function parseSheetTimestamps(raw) {
    // Cross-reference search; baselines are not applied here so timestamps
    // are interpreted in their relative form (MM:SS for 2-part).
    var segs = parseClipSegmentsForCell(raw, 0, CLIPGEN_CONFIG.defaultDuration);
    return segs.map(function (s) {
      return { start: s.startSeconds, duration: s.duration };
    });
  }

  // Per-participant indexes keyed on start time. Rebuilt when loadCrossRefData
  // receives fresh data, so findOverlapsForSearch can binary-search instead of
  // linearly scanning ssEvents / sheetRows (and re-parsing sheet timestamps)
  // on every rendered segment.

  function _buildEventsIndex() {
    var byP = {};
    for (var i = 0; i < state.ssEvents.length; i++) {
      var ev = state.ssEvents[i];
      var p = ev.participant;
      if (!byP[p]) byP[p] = [];
      byP[p].push({ in: ev.time_in, out: ev.time_out, ev: ev });
    }
    for (var k in byP) {
      byP[k].sort(function (a, b) { return a.in - b.in; });
    }
    state.xrefIndex.eventsByParticipant = byP;
  }

  function _buildSheetIndex() {
    var byP = {};
    for (var j = 0; j < state.sheetRows.length; j++) {
      var row = state.sheetRows[j];
      if (!row.cells) continue;
      for (var p in row.cells) {
        var cell = row.cells[p];
        if (!cell || !cell.valid) continue;
        var segs = parseSheetTimestamps(cell.value);
        if (segs.length === 0) continue;
        if (!byP[p]) byP[p] = [];
        for (var k = 0; k < segs.length; k++) {
          byP[p].push({
            start: segs[k].start,
            end: segs[k].start + segs[k].duration,
            observation: row.observation,
            category: row.category,
            rowIdx: j,
          });
        }
      }
    }
    for (var pp in byP) {
      byP[pp].sort(function (a, b) { return a.start - b.start; });
    }
    state.xrefIndex.sheetByParticipant = byP;
  }

  // Binary search: return the first index where arr[i].key >= value, using the
  // provided key function.
  function _lowerBound(arr, value, keyFn) {
    var lo = 0;
    var hi = arr.length;
    while (lo < hi) {
      var mid = (lo + hi) >>> 1;
      if (keyFn(arr[mid]) < value) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  function findOverlapsForSearch(participant, start, end) {
    var result = { screenspaceEvents: [], sheetObservations: [] };

    var events = state.xrefIndex.eventsByParticipant[participant];
    if (events && events.length > 0) {
      // Candidates are entries where `in < end`. Since events are sorted by
      // `in`, lower-bound on `end` gives us the exclusive upper cursor.
      var upper = _lowerBound(events, end, function (e) { return e.in; });
      for (var i = 0; i < upper; i++) {
        if (events[i].out > start) result.screenspaceEvents.push(events[i].ev);
      }
    }

    var segs = state.xrefIndex.sheetByParticipant[participant];
    if (segs && segs.length > 0) {
      var upper2 = _lowerBound(segs, end, function (s) { return s.start; });
      // Preserve the original "first observation per row" behavior by de-duping
      // on rowIdx. Row-iteration order matched sheet row order before; replicate
      // that by collecting matches and sorting by rowIdx.
      var seenRow = {};
      var matches = [];
      for (var m = 0; m < upper2; m++) {
        if (segs[m].end <= start) continue;
        var rid = segs[m].rowIdx;
        if (seenRow[rid]) continue;
        seenRow[rid] = true;
        matches.push(segs[m]);
      }
      matches.sort(function (a, b) { return a.rowIdx - b.rowIdx; });
      for (var n = 0; n < matches.length; n++) {
        result.sheetObservations.push({
          observation: matches[n].observation,
          category: matches[n].category,
        });
      }
    }

    return result;
  }

  // ---- Participants ----

  function needsTranscription() {
    for (var i = 0; i < state.participants.length; i++) {
      var p = state.participants[i];
      if (p.has_video && !p.has_transcript) return true;
    }
    return false;
  }

  function applyTranscriptionModelHint(data) {
    if (!data || !data.ok) return;
    state.modelStatus = data;
    // Track how long we've been in an apparent "failed to load" state, so
    // the indicator doesn't flash red on cold load before the first response.
    var looksFailed = !data.loaded && !data.warming && data.prewarm !== "off";
    if (looksFailed) {
      if (state.modelFailSince === 0) state.modelFailSince = Date.now();
    } else {
      state.modelFailSince = 0;
    }
    updateStatusIndicator();
  }

  // ---- Status indicator ----

  function _taskForSelectedParticipant() {
    var pid = state.selectedParticipant;
    if (!pid) return null;
    var latest = null;
    for (var i = 0; i < state.tasks.length; i++) {
      var t = state.tasks[i];
      if (t.participant !== pid) continue;
      // Priority: running > queued > failed > completed/cancelled > stale
      if (!latest) { latest = t; continue; }
      var order = { running: 5, queued: 4, failed: 3, completed: 2, cancelled: 1 };
      if ((order[t.status] || 0) > (order[latest.status] || 0)) latest = t;
    }
    return latest;
  }

  function _selectedParticipantRow() {
    var pid = state.selectedParticipant;
    if (!pid) return null;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) return state.participants[i];
    }
    return null;
  }

  function _modelLine() {
    var ms = state.modelStatus;
    if (!ms) return "Model: loading\u2026";
    if (ms.loaded) return "Model: ready" + (ms.model ? " \u00B7 " + ms.model : "");
    if (ms.warming) return "Model: loading\u2026";
    if (ms.prewarm === "off") return "Model: idle (loads on demand)";
    if (state.modelFailSince && Date.now() - state.modelFailSince >= MODEL_FAIL_GRACE_MS) {
      return "Model: failed to load";
    }
    return "Model: not loaded";
  }

  function _modelHasFailed() {
    var ms = state.modelStatus;
    if (!ms) return false;
    if (ms.loaded || ms.warming) return false;
    if (ms.prewarm === "off") return false;
    return state.modelFailSince > 0 && Date.now() - state.modelFailSince >= MODEL_FAIL_GRACE_MS;
  }

  function computeIndicatorState() {
    var pid = state.selectedParticipant;
    var task = _taskForSelectedParticipant();
    var row = _selectedParticipantRow();
    var cls = "status-indicator--ready";
    var taskLine;

    if (!pid) {
      taskLine = "No participant selected";
    } else if (task && task.status === "running") {
      cls = "status-indicator--working";
      var pct = Math.round((task.progress || 0) * 100);
      taskLine = pid + ": transcribing\u2026 " + pct + "%";
    } else if (task && task.status === "queued") {
      cls = "status-indicator--working";
      taskLine = pid + ": queued";
    } else if (task && task.status === "failed") {
      cls = "status-indicator--error";
      taskLine = pid + ": transcription failed" + (task.error ? " (" + task.error + ")" : "");
    } else if (row && row.has_transcript) {
      taskLine = pid + ": " + (row.segment_count || 0) + " segments";
      if (row.has_stale_artifacts) taskLine += " \u00B7 artifacts outdated";
    } else if (row && row.has_video) {
      taskLine = pid + ": not transcribed";
    } else {
      taskLine = pid + ": no source video";
    }

    // Model failure overrides non-task-active states (keep working class if a
    // task is currently running — the user can see the task is progressing).
    if (cls !== "status-indicator--working" && _modelHasFailed()) {
      cls = "status-indicator--error";
    }

    var lines = [_modelLine(), taskLine];
    var ariaLabel;
    if (cls === "status-indicator--working") ariaLabel = taskLine;
    else if (cls === "status-indicator--error") ariaLabel = lines.join(" \u2014 ");
    else ariaLabel = "Ready \u2014 " + taskLine;

    return { cls: cls, ariaLabel: ariaLabel, lines: lines };
  }

  function updateStatusIndicator() {
    var indicator = qs("#trStatusIndicator");
    if (!indicator) return;
    var s = computeIndicatorState();
    indicator.classList.remove(
      "status-indicator--ready",
      "status-indicator--error",
      "status-indicator--working"
    );
    indicator.classList.add(s.cls);
    indicator.setAttribute("aria-label", s.ariaLabel);
    var sr = qs("#trStatusSr");
    if (sr) sr.textContent = s.lines.join(" \u2014 ");
  }

  function initStatusIndicatorTooltip() {
    var indicator = qs("#trStatusIndicator");
    if (!indicator) return;
    attachHoverTooltip(indicator, function () {
      return computeIndicatorState().lines.join("\n");
    }, { multiline: true, align: "center" });
    updateStatusIndicator();
  }

  function refreshTranscriptionModelHintOnce() {
    apiGet("api/transcribe/model-status")
      .then(function (data) {
        applyTranscriptionModelHint(data);
      })
      .catch(function () {});
  }

  function stopModelHintPoll() {
    if (_modelHintPollTimer) {
      clearInterval(_modelHintPollTimer);
      _modelHintPollTimer = null;
    }
  }

  function startModelHintPoll() {
    stopModelHintPoll();
    var ticks = 0;
    var poll = function () {
      ticks++;
      if (ticks > 120) {
        stopModelHintPoll();
        return;
      }
      apiGet("api/transcribe/model-status")
        .then(function (data) {
          if (!data.ok) return;
          applyTranscriptionModelHint(data);
          if (data.loaded) {
            stopModelHintPoll();
            return;
          }
          if (!data.warming) {
            stopModelHintPoll();
          }
        })
        .catch(function () {});
    };
    poll();
    _modelHintPollTimer = setInterval(poll, 1500);
  }

  function maybeWarmOnPillHover(p, s) {
    if (state.transcribePrewarm !== "queue_open") return;
    if (!s || s.status === "completed") return;
    tryPostTranscriptionWarmup();
  }

  // Ask the backend to preload the Whisper model. Idempotent per page load
  // via `_transcriptionWarmupPosted`; we reset the flag whenever the post
  // didn't actually lead to a load (skipped / error / no-op response) so a
  // later trigger (e.g. pill hover) can retry. Five response branches:
  //   skipped         — backend declined (e.g. no GPU policy met); keep flag
  //                     down so a later hover can re-prompt.
  //   already_loaded  — model is in memory; nothing to poll.
  //   started / warming — kick off the model-hint poller until it finishes.
  //   !ok / catch     — treat as transient; clear flag for retry.
  function tryPostTranscriptionWarmup() {
    if (_transcriptionWarmupPosted) return;
    if (state.transcribePrewarm === "off") return;
    if (!needsTranscription()) return;
    _transcriptionWarmupPosted = true;
    apiPost("api/transcribe/warmup", {})
      .then(function (data) {
        if (!data.ok) {
          _transcriptionWarmupPosted = false;
          return;
        }
        if (data.skipped) {
          _transcriptionWarmupPosted = false;
          refreshTranscriptionModelHintOnce();
          return;
        }
        if (data.already_loaded) {
          refreshTranscriptionModelHintOnce();
          return;
        }
        if (data.started || data.already_warming) {
          startModelHintPoll();
          return;
        }
        _transcriptionWarmupPosted = false;
      })
      .catch(function () {
        _transcriptionWarmupPosted = false;
      });
  }

  function loadParticipants() {
    return apiGet("api/participants").then(function (data) {
      if (!data.ok) return;
      state.participants = data.participants;
      state.transcribePrewarm = data.transcribe_prewarm || "queue_open";
      renderPills();
      refreshTopNavActions();

      if (!needsTranscription()) {
        _transcriptionWarmupPosted = false;
      } else if (state.transcribePrewarm === "page_load") {
        tryPostTranscriptionWarmup();
      }
      refreshTranscriptionModelHintOnce();

      // Preserve current in-memory selection if still valid (soft refresh)
      if (state.selectedParticipant) {
        for (var i = 0; i < state.participants.length; i++) {
          if (state.participants[i].id === state.selectedParticipant) return;
        }
      }

      // Restore from localStorage if present and still valid (fresh page load)
      var storedPid = getStoredUIState("transcripts").selectedParticipant;
      if (storedPid) {
        for (var j = 0; j < state.participants.length; j++) {
          if (state.participants[j].id === storedPid) {
            selectParticipant(storedPid);
            return;
          }
        }
      }

      // Auto-select first participant with a transcript, or just the first
      var first = null;
      for (var k = 0; k < state.participants.length; k++) {
        if (state.participants[k].has_transcript) { first = state.participants[k]; break; }
      }
      if (!first && state.participants.length > 0) first = state.participants[0];
      if (first) selectParticipant(first.id);
      else renderEmptyState();
    });
  }


  function selectParticipant(pid) {
    _participantReqVer++;
    _stopSummaryPoll();
    _stopCitationsPoll();
    hideMarkPopover();
    state.selectedParticipant = pid;
    setStoredUIStateField("transcripts", "selectedParticipant", pid);
    renderPills();
    refreshTopNavActions();

    // Find participant info
    var p = null;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) { p = state.participants[i]; break; }
    }
    if (!p) return;

    // Clean up previous video state
    var video = qs("#videoPlayer");
    var videoEmpty = qs("#videoEmpty");
    video.pause();
    _pendingSeekTime = null;
    cancelAnimationFrame(_seekRaf);
    _seekRaf = 0;
    if (_pendingSeekListener) {
      video.removeEventListener("loadedmetadata", _pendingSeekListener);
      _pendingSeekListener = null;
    }

    // Set video source
    if (p.has_video) {
      video.src = "media/" + p.video_filename;
      video.classList.remove("hidden");
      videoEmpty.classList.add("hidden");

      // Set VTT track. Browsers reset textTracks[0].mode when the track src
      // changes, so re-apply the user's preference now and again on the
      // track's load event in initVideoPlayer.
      var track = qs("#subtitleTrack");
      track.src = "api/vtt/" + pid;
      applyCaptionMode();

      // Restore the saved playback offset for this participant if we have one;
      // otherwise seek to 0.001s so the first frame renders without waiting
      // for play/scrub. `preload="metadata"` decodes duration but not pixels,
      // so without this nudge the viewer shows a blank gray box on load.
      var storedMap = getStoredUIState("transcripts").videoTimeByParticipant;
      var savedTime =
        storedMap && typeof storedMap[pid] === "number" ? storedMap[pid] : 0.001;
      var restoreTime = function () {
        video.removeEventListener("loadedmetadata", restoreTime);
        if (state.selectedParticipant === pid) video.currentTime = savedTime;
      };
      video.addEventListener("loadedmetadata", restoreTime);
    } else {
      video.removeAttribute("src");
      video.load();
      video.classList.add("hidden");
      videoEmpty.classList.remove("hidden");
    }

    var taskForPid = null;
    state.tasks.forEach(function (t) {
      if (t.participant === pid && (t.status === "running" || t.status === "queued")) {
        taskForPid = t;
      }
    });

    updateStatusIndicator();

    // Load transcript
    if (p.has_transcript) {
      state.streamingParticipant = null;
      loadTranscript(pid);
      loadSummary(pid);
    } else if (taskForPid && taskForPid.status === "running" && taskForPid.partial_segments && taskForPid.partial_segments.length > 0) {
      renderPartialSegments(taskForPid.partial_segments, taskForPid.progress);
      state.streamingParticipant = pid;
      clearSummary();
    } else {
      state.segments = [];
      state.streamingParticipant = null;
      renderSegments();
      renderTimeline();
      clearSummary();
    }
  }

  function renderEmptyState() {
    qs("#videoPlayer").classList.add("hidden");
    qs("#videoEmpty").classList.remove("hidden");
    qs("#segmentList").innerHTML = "";
    qs("#transcriptEmpty").classList.remove("hidden");
    clearSummary();
    _markerHitRects = [];
    renderTimeline();
  }

  // ---- Transcript loading ----

  function loadTranscript(pid) {
    var ver = _participantReqVer;
    return apiGet("api/transcript/" + pid).then(function (data) {
      if (ver !== _participantReqVer) return;
      if (!data.ok) {
        state.segments = [];
        renderSegments();
        renderTimeline();
        return;
      }
      state.segments = data.segments;
      state.activeSegmentIndex = -1;
      renderSegments();
      renderTimeline();
    });
  }

  // ---- AI Summary ----
  //
  // Two cooperating pollers:
  //   _summaryPollTimer   — runs while the backend is still generating the
  //                         summary. Stops as soon as a summary lands, or
  //                         when the user switches participant.
  //   _citationsPollTimer — runs after the summary arrives if citations are
  //                         still being computed (citations depend on summary).
  // Both are cleared by their own _stop*Poll() helpers; either also stops if
  // `state.selectedParticipant` no longer matches the participant the timer
  // was started for.

  var _summaryPollTimer = null;

  function loadSummary(pid) {
    var section = qs("#summarySection");
    var ver = _participantReqVer;

    apiGet("api/summary/" + pid).then(function (data) {
      if (ver !== _participantReqVer) return;
      if (data.ok && data.summary) {
        _stopSummaryPoll();
        renderSummary(data.summary);
        // Handle citations
        if (data.citations && data.citations.length > 0) {
          state.summaryCitations = data.citations;
          state.citationsGenerating = false;
          renderCitations();
        } else if (data.citations_generating) {
          state.summaryCitations = null;
          state.citationsGenerating = true;
          renderCitationsStatus();
          _startCitationsPoll(pid);
          _refreshAgentStateNow();
        }
      } else if (data.generating) {
        renderSummaryGenerating();
        _startSummaryPoll(pid);
        _refreshAgentStateNow();
      } else {
        section.classList.add("hidden");
      }
    }).catch(function () {
      if (ver !== _participantReqVer) return;
      // Ollama unavailable or no summary — stay hidden
      section.classList.add("hidden");
    });
  }

  function _startSummaryPoll(pid) {
    _stopSummaryPoll();
    var ver = _participantReqVer;
    _summaryPollTimer = setInterval(function () {
      if (ver !== _participantReqVer || state.selectedParticipant !== pid) {
        _stopSummaryPoll();
        return;
      }
      apiGet("api/summary/" + pid).then(function (data) {
        if (ver !== _participantReqVer) return;
        if (data.ok && data.summary) {
          _stopSummaryPoll();
          renderSummary(data.summary);
          // Summary just arrived — check citation status
          if (data.citations && data.citations.length > 0) {
            state.summaryCitations = data.citations;
            state.citationsGenerating = false;
            renderCitations();
          } else if (data.citations_generating) {
            state.summaryCitations = null;
            state.citationsGenerating = true;
            renderCitationsStatus();
            _startCitationsPoll(pid);
          }
        } else if (!data.generating) {
          // Generation finished without result — stop polling
          _stopSummaryPoll();
          qs("#summarySection").classList.add("hidden");
        }
      }).catch(function () {
        if (ver !== _participantReqVer) return;
        _stopSummaryPoll();
        qs("#summarySection").classList.add("hidden");
      });
    }, 3000);
  }

  function _stopSummaryPoll() {
    if (_summaryPollTimer) {
      clearInterval(_summaryPollTimer);
      _summaryPollTimer = null;
    }
  }

  function renderSummaryGenerating() {
    var section = qs("#summarySection");
    var content = qs("#summaryContent");
    content.innerHTML = '<p class="summary-generating">Generating summary\u2026</p>';
    section.classList.remove("hidden");
    section.classList.toggle("collapsed", state.summaryCollapsed);
    qs("#summaryToggle").setAttribute("aria-expanded", state.summaryCollapsed ? "false" : "true");
    qs("#summaryActions").classList.add("hidden");
    state.summaryEditing = false;
    state.summaryText = "";
  }

  function renderSummary(text) {
    state.summaryText = text;
    state.summaryEditing = false;
    var section = qs("#summarySection");
    var content = qs("#summaryContent");
    var lines = text.split("\n");
    var paragraphSentences = [];
    var bullets = [];
    var inBullets = false;

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;
      if (line.indexOf("- ") === 0 || line.indexOf("* ") === 0) {
        inBullets = true;
        bullets.push(escapeHtml(line.substring(2)));
      } else if (!inBullets) {
        // Split paragraph into individual sentences for citation targeting
        var parts = line.split(/(?<=[.!?])\s+/);
        for (var k = 0; k < parts.length; k++) {
          var part = parts[k].trim();
          if (part) paragraphSentences.push(escapeHtml(part));
        }
      } else {
        bullets.push(escapeHtml(line));
      }
    }

    // Build HTML with data-cite-index on each sentence/bullet
    var citeIdx = 0;
    var html = "";
    if (paragraphSentences.length > 0) {
      html += "<p>";
      for (var si = 0; si < paragraphSentences.length; si++) {
        if (si > 0) html += " ";
        html += '<span data-cite-index="' + citeIdx + '">' + paragraphSentences[si] + "</span>";
        citeIdx++;
      }
      html += "</p>";
    }
    if (bullets.length > 0) {
      html += "<ul>";
      for (var j = 0; j < bullets.length; j++) {
        html += '<li data-cite-index="' + citeIdx + '">' + bullets[j] + "</li>";
        citeIdx++;
      }
      html += "</ul>";
    }

    content.innerHTML = html;
    section.classList.remove("hidden");
    section.classList.toggle("collapsed", state.summaryCollapsed);
    qs("#summaryToggle").setAttribute("aria-expanded", state.summaryCollapsed ? "false" : "true");
    qs("#summaryActions").classList.remove("hidden");
    _setSummaryEditMode(false);

    // Re-apply citations if already loaded
    if (state.summaryCitations) {
      renderCitations();
    }
  }

  function clearSummary() {
    _stopSummaryPoll();
    _stopCitationsPoll();
    qs("#summarySection").classList.add("hidden");
    qs("#summaryContent").innerHTML = "";
    qs("#summaryActions").classList.add("hidden");
    state.summaryEditing = false;
    state.summaryText = "";
    state.summaryCitations = null;
    state.citationsGenerating = false;
  }

  // ---- Citation rendering (Pass 2) ----

  var _citationsPollTimer = null;

  function renderCitations() {
    // Remove any existing status text
    var existing = qs("#summaryContent .citations-status");
    if (existing) existing.remove();

    // Remove any previously rendered citation links
    var oldLinks = qs("#summaryContent").querySelectorAll(".citation-link");
    for (var r = 0; r < oldLinks.length; r++) oldLinks[r].remove();

    if (!state.summaryCitations) return;

    var refNum = 1;
    for (var i = 0; i < state.summaryCitations.length; i++) {
      var cite = state.summaryCitations[i];
      if (!cite.refs || cite.refs.length === 0) continue;
      var el = qs('#summaryContent [data-cite-index="' + i + '"]');
      if (!el) continue;
      for (var j = 0; j < cite.refs.length; j++) {
        var ref = cite.refs[j];
        var sup = document.createElement("sup");
        sup.className = "citation-link";
        sup.dataset.start = String(ref.start);
        sup.title = formatTime(ref.start);
        sup.textContent = "[" + refNum + "]";
        (function (startTime) {
          sup.addEventListener("click", function (e) {
            e.stopPropagation();
            seekVideo(startTime);
          });
        })(ref.start);
        el.appendChild(sup);
        refNum++;
      }
    }
  }

  function renderCitationsStatus() {
    // Remove any existing status
    var existing = qs("#summaryContent .citations-status");
    if (existing) existing.remove();

    var p = document.createElement("p");
    p.className = "citations-status";
    p.textContent = "Finding sources\u2026";
    qs("#summaryContent").appendChild(p);
  }

  // Hard cap on the citations poll. The previous 90 s value was shorter than
  // some real Ollama runs on long transcripts, so the result would land in
  // the manifest after we'd given up — and the UI only picked it up on a
  // full page reload. Five minutes covers realistic completion times; the
  // server-side `generating: false` signal stops the poll earlier when the
  // agent finishes (or fails) sooner.
  var _CITATIONS_POLL_TIMEOUT = 300000;

  function _startCitationsPoll(pid) {
    _stopCitationsPoll();
    var started = Date.now();
    var ver = _participantReqVer;
    _citationsPollTimer = setInterval(function () {
      if (ver !== _participantReqVer ||
          state.selectedParticipant !== pid ||
          Date.now() - started > _CITATIONS_POLL_TIMEOUT) {
        _stopCitationsPoll();
        state.citationsGenerating = false;
        var status = qs("#summaryContent .citations-status");
        if (status) status.remove();
        return;
      }
      apiGet("api/citations/" + pid).then(function (data) {
        if (ver !== _participantReqVer) return;
        if (data.ok && data.citations) {
          _stopCitationsPoll();
          state.summaryCitations = data.citations;
          state.citationsGenerating = false;
          renderCitations();
        } else if (!data.generating) {
          _stopCitationsPoll();
          state.citationsGenerating = false;
          var status = qs("#summaryContent .citations-status");
          if (status) status.remove();
        }
      }).catch(function () {
        if (ver !== _participantReqVer) return;
        _stopCitationsPoll();
        state.citationsGenerating = false;
        var status = qs("#summaryContent .citations-status");
        if (status) status.remove();
      });
    }, 3000);
  }

  function _stopCitationsPoll() {
    if (_citationsPollTimer) {
      clearInterval(_citationsPollTimer);
      _citationsPollTimer = null;
    }
  }

  function initSummaryToggle() {
    qs("#summaryHeader").addEventListener("click", function () {
      var section = qs("#summarySection");
      state.summaryCollapsed = !state.summaryCollapsed;
      section.classList.toggle("collapsed", state.summaryCollapsed);
      qs("#summaryToggle").setAttribute("aria-expanded", state.summaryCollapsed ? "false" : "true");
    });
  }

  function _setSummaryEditMode(editing) {
    var btn = qs("#summaryEdit");
    var icon = btn.querySelector(".summary-action-icon");
    if (editing) {
      icon.classList.remove("summary-action-edit");
      icon.classList.add("summary-action-save");
      btn.title = "Save summary";
      btn.setAttribute("aria-label", "Save summary");
    } else {
      icon.classList.remove("summary-action-save");
      icon.classList.add("summary-action-edit");
      btn.title = "Edit summary";
      btn.setAttribute("aria-label", "Edit summary");
    }
  }

  function initSummaryActions() {
    qs("#summaryRegenerate").addEventListener("click", function (e) {
      e.stopPropagation();
      var pid = state.selectedParticipant;
      if (!pid) return;
      var previousText = state.summaryText;
      state.summaryCitations = null;
      state.citationsGenerating = false;
      _stopCitationsPoll();
      renderSummaryGenerating();
      apiPost("api/summary/" + pid + "/regenerate", {}).then(function (data) {
        if (data.ok && data.generating) {
          _startSummaryPoll(pid);
        }
      }).catch(function () {
        showToast("Failed to regenerate summary");
        if (previousText) {
          renderSummary(previousText);
        }
      });
    });

    qs("#citationsRegenerate").addEventListener("click", function (e) {
      e.stopPropagation();
      var pid = state.selectedParticipant;
      if (!pid) return;
      state.summaryCitations = null;
      state.citationsGenerating = true;
      renderCitations(); // clear existing links
      renderCitationsStatus();
      apiPost("api/citations/" + pid + "/regenerate", {}).then(function (data) {
        if (data.ok && data.generating) {
          _startCitationsPoll(pid);
        }
      }).catch(function () {
        showToast("Failed to regenerate citations");
        state.citationsGenerating = false;
        var status = qs("#summaryContent .citations-status");
        if (status) status.remove();
      });
    });

    qs("#summaryCopy").addEventListener("click", function (e) {
      e.stopPropagation();
      var text = state.summaryEditing
        ? qs("#summaryContent textarea").value
        : state.summaryText;
      if (!text) return;
      navigator.clipboard.writeText(text).then(function () {
        showToast("Summary copied");
      });
    });

    qs("#summaryEdit").addEventListener("click", function (e) {
      e.stopPropagation();
      var pid = state.selectedParticipant;
      if (!pid) return;
      if (!state.summaryEditing) {
        var ta = document.createElement("textarea");
        ta.className = "summary-edit-textarea";
        ta.value = state.summaryText;
        ta.autocomplete = "off";
        ta.addEventListener("click", function (ev) { ev.stopPropagation(); });
        qs("#summaryContent").innerHTML = "";
        qs("#summaryContent").appendChild(ta);
        ta.focus();
        state.summaryEditing = true;
        _setSummaryEditMode(true);
      } else {
        var newText = qs("#summaryContent textarea").value.trim();
        if (!newText) {
          showToast("Summary cannot be empty");
          return;
        }
        apiPut("api/summary/" + pid, { summary: newText }).then(function () {
          state.summaryCitations = null;
          state.citationsGenerating = false;
          _stopCitationsPoll();
          renderSummary(newText);
          showToast("Summary saved");
        }).catch(function () {
          showToast("Failed to save summary");
        });
      }
    });
  }

  // ---- Segment rendering ----

  var _cachedSegmentRows = null;

  function renderSegments() {
    var container = qs("#segmentList");
    var empty = qs("#transcriptEmpty");
    state.editingTextEl = null;
    _cachedSegmentRows = null;

    if (state.segments.length === 0) {
      container.innerHTML = "";
      empty.classList.remove("hidden");
      return;
    }
    empty.classList.add("hidden");

    var html = "";
    for (var i = 0; i < state.segments.length; i++) {
      var seg = state.segments[i];
      var activeClass = i === state.activeSegmentIndex ? " active" : "";
      var correctedClass = seg.corrected ? " segment-corrected" : "";
      var markObj = seg.marks && seg.marks.length > 0 ? seg.marks[0] : null;
      var markCat = markObj ? (MARK_CATEGORIES[markObj.category] || MARK_CATEGORIES.bookmark) : null;
      var markColor = markCat ? markCat.color : null;
      var markClass = markObj ? "segment-mark marked" : "segment-mark";
      var markStyle = markColor ? ' style="background:' + markColor + '"' : "";
      var markLabel = markObj && markObj.label ? ' title="' + escapeHtml(markObj.label) + '"' : "";
      var annoBadgeHtml = "";
      if (markObj && markObj.label && markColor) {
        var bgMix = "color-mix(in oklch, " + markColor + " 18%, transparent)";
        var borderMix = "color-mix(in oklch, " + markColor + " 50%, transparent)";
        var badgeStyle = "--anno-badge-fg:" + markColor + ";--anno-badge-bg:" + bgMix + ";--anno-badge-border:" + borderMix;
        annoBadgeHtml = '<span class="segment-anno-badge" style="' + badgeStyle + '">' + escapeHtml(markObj.label) + '</span>';
      }

      html += '<div class="segment-row' + activeClass + correctedClass + '" data-index="' + i + '" data-start="' + seg.start + '">';
      html += '<span class="' + markClass + '" data-segment-id="' + escapeHtml(seg.id) + '"' + markStyle + markLabel + '></span>';
      html += '<span class="segment-timestamp">' + formatTime(seg.start);
      // Cross-reference badges in gutter (inside timestamp, positioned at right edge)
      if (state.tooltipsEnabled) {
        var xref = findOverlapsForSearch(state.selectedParticipant, seg.start, seg.end);
        if (xref.screenspaceEvents.length > 0 || xref.sheetObservations.length > 0) {
          html += '<span class="segment-xref-badges">';
          if (xref.screenspaceEvents.length > 0) {
            var evTypes = [];
            var evSeen = {};
            for (var ei = 0; ei < xref.screenspaceEvents.length; ei++) {
              var et = xref.screenspaceEvents[ei].event_type || xref.screenspaceEvents[ei].detector;
              if (!evSeen[et]) { evSeen[et] = true; evTypes.push(et); }
            }
            html += '<span class="segment-xref-badge" style="background:' + XREF_BADGES.screenspace.color + '" title="' + escapeHtml(evTypes.join(", ")) + '"><span class="xref-badge-icon" style="' + maskIconStyle("url(icons/" + XREF_BADGES.screenspace.icon + ".svg)") + '"></span></span>';
          }
          if (xref.sheetObservations.length > 0) {
            var obsTitle = xref.sheetObservations[0].observation;
            html += '<span class="segment-xref-badge" style="background:' + XREF_BADGES.sheet.color + '" title="' + escapeHtml(obsTitle) + '"><span class="xref-badge-icon" style="' + maskIconStyle("url(icons/" + XREF_BADGES.sheet.icon + ".svg)") + '"></span></span>';
          }
          html += '</span>';
        }
      }
      html += '</span>';
      // Split text into word spans
      var tokens = seg.text.split(/(\s+)/);
      var wordHtml = "";
      for (var w = 0; w < tokens.length; w++) {
        if (/^\s+$/.test(tokens[w])) {
          wordHtml += tokens[w];
        } else if (tokens[w]) {
          wordHtml += '<span class="segment-word">' + escapeHtml(tokens[w]) + '</span>';
        }
      }
      html += '<span class="segment-text" data-id="' + escapeHtml(seg.id) + '">' + annoBadgeHtml + wordHtml + '</span>';
      html += '<span class="segment-copy" title="Copy text"><span class="segment-copy-icon"></span></span>';
      html += '</div>';
    }
    container.innerHTML = html;

    _ensureSegmentListDelegation();
    _partialRender.count = 0;
    _partialRender.pid = null;
    _partialRender.segments = null;
    _partialRender.marksVersion = _streamingMarksVersion;
  }

  // Append-only state for renderPartialSegments. Each streaming poll appends new
  // trailing segments to #segmentList instead of rebuilding the entire list. A
  // full rebuild is only performed when the participant changes, when the
  // segment count drops (restart), or when the in-memory marks cache changes.

  var _partialRender = {
    pid: null,
    count: 0,
    segments: null,
    marksVersion: 0,
  };

  function _renderPartialSegmentRow(seg, i, pid) {
    var segId = pid + ":" + i;
    var cachedMark = _streamingMarks[segId];
    var cachedColor = cachedMark ? cachedMark.color : null;
    var markClass = "segment-mark" + (cachedColor ? " marked" : "");
    var markStyle = cachedColor ? ' style="background:' + cachedColor + '"' : "";
    var html = '<div class="segment-row segment-streaming" data-index="' + i + '" data-start="' + seg.start + '">';
    html += '<span class="' + markClass + '" data-segment-id="' + escapeHtml(segId) + '"' + markStyle + '></span>';
    html += '<span class="segment-timestamp">' + formatTime(seg.start) + '</span>';
    html += '<span class="segment-text">' + escapeHtml(seg.text) + '</span>';
    html += '<span class="segment-copy" title="Copy text"><span class="segment-copy-icon"></span></span>';
    html += '</div>';
    return html;
  }

  function _streamingIndicatorHtml(progress) {
    return '<div class="streaming-indicator">' +
      '<span class="streaming-dot"></span>' +
      'Transcribing\u2026 ' + Math.round(progress * 100) + '%' +
      '</div>';
  }

  var _streamIndicatorRaf = null;
  var _streamIndicatorPending = null;

  function _updateStreamingIndicator(container, progress) {
    _streamIndicatorPending = { container: container, progress: progress };
    if (_streamIndicatorRaf) return;
    _streamIndicatorRaf = requestAnimationFrame(function () {
      _streamIndicatorRaf = null;
      var pending = _streamIndicatorPending;
      _streamIndicatorPending = null;
      if (!pending || !pending.container || !pending.container.isConnected) return;
      var ind = pending.container.querySelector(".streaming-indicator");
      if (ind) ind.parentNode.removeChild(ind);
      pending.container.insertAdjacentHTML(
        "beforeend",
        _streamingIndicatorHtml(pending.progress)
      );
    });
  }

  function renderPartialSegments(segments, progress) {
    var container = qs("#segmentList");
    var empty = qs("#transcriptEmpty");
    var pid = state.streamingParticipant || state.selectedParticipant;
    empty.classList.add("hidden");

    // If user is actively editing a segment, skip DOM mutation to preserve edit state
    if (state.editingTextEl && state.editingTextEl.isConnected) return;

    // Row list changes shape on both append and rebuild paths.
    _cachedSegmentRows = null;

    // Pass 6 floating-nav scroll-under: scroll lives on #trMain, not on
    // #segmentList. Probe the actual scroll container so the
    // auto-follow-streaming-tail behaviour keeps working.
    var trMain = qs("#trMain");
    var scrollHost = trMain || container;
    var nearBottom = scrollHost.scrollHeight - scrollHost.scrollTop - scrollHost.clientHeight < 100;

    var canAppend =
      _partialRender.pid === pid &&
      _partialRender.marksVersion === _streamingMarksVersion &&
      segments.length >= _partialRender.count &&
      container.querySelector(".segment-streaming") !== null;
    var indicatorOnly =
      canAppend && segments.length === _partialRender.count;

    if (canAppend && segments.length > _partialRender.count) {
      // Append only the new trailing rows, then move/refresh the indicator.
      var indicator = container.querySelector(".streaming-indicator");
      if (indicator) indicator.parentNode.removeChild(indicator);
      var tmp = document.createElement("div");
      var fragHtml = "";
      for (var a = _partialRender.count; a < segments.length; a++) {
        fragHtml += _renderPartialSegmentRow(segments[a], a, pid);
      }
      tmp.innerHTML = fragHtml;
      var frag = document.createDocumentFragment();
      while (tmp.firstChild) frag.appendChild(tmp.firstChild);
      container.appendChild(frag);
      container.insertAdjacentHTML("beforeend", _streamingIndicatorHtml(progress));
    } else if (canAppend && segments.length === _partialRender.count) {
      // Same segment count - only refresh the progress indicator (coalesced).
      _updateStreamingIndicator(container, progress);
    } else {
      // Full rebuild: pid change, restart, or marks cache changed.
      var html = "";
      for (var i = 0; i < segments.length; i++) {
        html += _renderPartialSegmentRow(segments[i], i, pid);
      }
      html += _streamingIndicatorHtml(progress);
      container.innerHTML = html;
      _ensureSegmentListDelegation();
    }

    _partialRender.pid = pid;
    _partialRender.count = segments.length;
    _partialRender.segments = segments;
    _partialRender.marksVersion = _streamingMarksVersion;

    if (nearBottom) {
      scrollHost.scrollTop = scrollHost.scrollHeight;
    }

    // Mirror partial segments into state.segments so the marker timeline can
    // resolve marker positions during streaming. Each row's id is the
    // composed "<pid>:<index>" string, matching what _streamingMarks is keyed
    // on via _renderPartialSegmentRow above.
    var mirrored = [];
    for (var mi = 0; mi < segments.length; mi++) {
      var s = segments[mi];
      mirrored.push({
        id: pid + ":" + mi,
        start: s.start,
        end: s.end,
        text: s.text,
        marks: [],
      });
    }
    state.segments = mirrored;
    if (!indicatorOnly) {
      renderTimeline();
    }
  }

  // ---- Segment list event delegation ----

  var _segmentListDelegated = false;

  function _ensureSegmentListDelegation() {
    if (_segmentListDelegated) return;
    var container = qs("#segmentList");
    if (!container) return;
    _segmentListDelegated = true;

    container.addEventListener("click", function (e) {
      var row = e.target.closest(".segment-row");
      if (!row) return;
      var isStreaming = row.classList.contains("segment-streaming");
      var idx = parseInt(row.getAttribute("data-index"), 10);
      var start = parseFloat(row.getAttribute("data-start"));

      var markEl = e.target.closest(".segment-mark");
      if (markEl && row.contains(markEl)) {
        e.stopPropagation();
        var segId = markEl.getAttribute("data-segment-id");
        if (isStreaming) {
          var existing = _streamingMarks[segId];
          if (existing) showMarkPopover(markEl, segId, existing);
          else toggleMarkStreaming(segId, markEl);
        } else {
          var seg = state.segments[idx];
          var mark = seg && seg.marks && seg.marks.length > 0 ? seg.marks[0] : null;
          if (mark) showMarkPopover(markEl, segId, mark);
          else toggleMark(segId);
        }
        return;
      }

      var copyEl = e.target.closest(".segment-copy");
      if (copyEl && row.contains(copyEl)) {
        e.stopPropagation();
        var src = isStreaming ? (_partialRender.segments || []) : state.segments;
        var segCopy = src[idx];
        if (!segCopy) return;
        navigator.clipboard.writeText(segCopy.text).then(function () {
          showToast("Copied to clipboard");
        });
        return;
      }

      var tsEl = e.target.closest(".segment-timestamp");
      if (tsEl && row.contains(tsEl)) {
        e.stopPropagation();
        seekVideo(start);
        return;
      }

      var textEl = e.target.closest(".segment-text");
      if (textEl && row.contains(textEl)) {
        e.stopPropagation();
        if (state.editingTextEl === textEl) return;
        seekVideo(start);
        return;
      }
    });

    container.addEventListener("dblclick", function (e) {
      var row = e.target.closest(".segment-row");
      if (!row) return;
      var textEl = e.target.closest(".segment-text");
      if (!textEl || !row.contains(textEl)) return;
      e.stopPropagation();
      startSegmentEditing(textEl);
    });
  }

  // Cache marks made during streaming so they survive DOM rebuilds.
  // Each entry: { color, id, category, label }. `version` is bumped on any
  // write to invalidate renderPartialSegments' append-only fast path.
  var _streamingMarks = {};
  var _streamingMarksVersion = 0;
  var _streamingMarksLoaded = false;

  function _bumpStreamingMarksVersion() {
    _streamingMarksVersion++;
    renderTimeline();
  }

  function _loadStreamingMarks(pid) {
    if (_streamingMarksLoaded) return;
    _streamingMarksLoaded = true;
    apiGet("api/marks").then(function (data) {
      if (!data.ok) return;
      if (data.categories) setMarkCategories(data.categories);
      var added = false;
      data.marks.forEach(function (m) {
        if (!m.valid || m.participant !== pid) return;
        if (_streamingMarks[m.segment_id]) return; // don't overwrite fresh marks
        var cat = MARK_CATEGORIES[m.category] || MARK_CATEGORIES.bookmark;
        _streamingMarks[m.segment_id] = {
          color: cat.color,
          id: m.id,
          category: m.category,
          label: m.label || "",
        };
        added = true;
      });
      if (added) _bumpStreamingMarksVersion();
    });
  }

  // ---- Custom video player ----
  // Custom DOM controls on #videoPlayer (no native controls=""). Timeline
  // canvas + mark hit-testing live after timelineXToTime().

  var _markerHitRects = [];
  var _playheadRaf = 0;
  var _timelineTooltipRaf = 0;
  var _lastTimelineHit = null;
  var _timelineResizeObs = null;

  function setIconClass(span, klass) {
    if (!span) return;
    span.className = "player-btn-icon " + klass;
  }

  function updatePlayerButtons() {
    var playBtn = qs("#videoPlayBtn");
    var muteBtn = qs("#videoMuteBtn");
    var speedBtn = qs("#videoSpeedBtn");
    var ccBtn = qs("#videoCcBtn");
    var pipBtn = qs("#videoPipBtn");
    if (playBtn) {
      var pIcon = playBtn.querySelector(".player-btn-icon");
      setIconClass(pIcon, state.videoPlaying ? "player-icon-pause" : "player-icon-play");
      playBtn.title = state.videoPlaying ? "Pause (Space)" : "Play/Pause (Space)";
    }
    if (muteBtn) {
      var mIcon = muteBtn.querySelector(".player-btn-icon");
      setIconClass(mIcon, state.videoMuted ? "player-icon-mute-off" : "player-icon-mute");
      muteBtn.classList.toggle("active", !state.videoMuted);
    }
    if (speedBtn) {
      speedBtn.textContent = state.videoPlaybackRate + "x";
      speedBtn.classList.toggle("active", state.videoPlaybackRate !== 1);
    }
    if (ccBtn) {
      ccBtn.classList.toggle("active", state.ccEnabled);
      ccBtn.setAttribute("aria-pressed", state.ccEnabled ? "true" : "false");
    }
    if (pipBtn) {
      pipBtn.classList.toggle("active", state.pipEnabled);
      pipBtn.setAttribute("aria-pressed", state.pipEnabled ? "true" : "false");
    }
    var collapseBtn = qs("#videoCollapseBtn");
    if (collapseBtn) {
      collapseBtn.title = state.videoCollapsed ? "Show video" : "Hide video";
      collapseBtn.setAttribute("aria-expanded", state.videoCollapsed ? "false" : "true");
    }
  }

  function applyPlaybackRate() {
    var v = qs("#videoPlayer");
    if (!v) return;
    v.defaultPlaybackRate = state.videoPlaybackRate;
    v.playbackRate = state.videoPlaybackRate;
    // Audio pitch will rise at high speeds, but the time-stretch filter is CPU
    // heavy and causes judder at >=2x.
    v.preservesPitch = false;
    v.mozPreservesPitch = false;
    v.webkitPreservesPitch = false;
  }

  function updateTimeLabel() {
    var v = qs("#videoPlayer");
    var label = qs("#videoTime");
    if (!v || !label) return;
    var dur = isFinite(v.duration) ? v.duration : 0;
    label.textContent = formatTime(v.currentTime || 0) + " / " + formatTime(dur);
  }

  function applyCaptionMode() {
    var v = qs("#videoPlayer");
    if (!v || !v.textTracks || !v.textTracks.length) return;
    v.textTracks[0].mode = state.ccEnabled ? "showing" : "hidden";
  }

  function sizeTimelineCanvas() {
    var wrap = qs("#timelineCanvasWrapper");
    var c1 = qs("#timelineCanvas");
    var c2 = qs("#playheadCanvas");
    if (!wrap || !c1 || !c2) return;
    var rect = wrap.getBoundingClientRect();
    if (rect.width === 0) return;
    var dpr = window.devicePixelRatio || 1;
    [c1, c2].forEach(function (c) {
      var cssH = c === c2 ? c.offsetHeight || 48 : c.offsetHeight || 48;
      var cssW = rect.width;
      c.width = Math.round(cssW * dpr);
      c.height = Math.round(cssH * dpr);
      var ctx = c.getContext("2d");
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    });
  }

  function computeTickInterval(visLen) {
    var candidates = [1, 2, 5, 10, 15, 30, 60, 120, 300, 600, 900, 1800, 3600];
    var target = visLen / 8;
    for (var i = 0; i < candidates.length; i++) {
      if (candidates[i] >= target) return candidates[i];
    }
    return candidates[candidates.length - 1];
  }

  function getMarkForSegment(seg) {
    if (seg.marks && seg.marks.length > 0) return seg.marks[0];
    var streaming = _streamingMarks[seg.id];
    return streaming || null;
  }

  function renderTimeline() {
    var canvas = qs("#timelineCanvas");
    if (!canvas) return;
    var ctx = canvas.getContext("2d");
    var v = qs("#videoPlayer");
    var cssW = canvas.offsetWidth;
    var cssH = canvas.offsetHeight;
    ctx.clearRect(0, 0, cssW, cssH);

    var isDark = (document.documentElement.getAttribute("data-theme") || "").toLowerCase() === "dark";
    var surfaceAlt = getCSSVar("--color-surface-alt", isDark ? "#1f2937" : "#f1ece4");
    var border = getCSSVar("--color-border", isDark ? "#374151" : "#e0ddd7");
    var textDim = getCSSVar("--color-text-dim", isDark ? "#9ca3af" : "#6b7280");

    ctx.fillStyle = surfaceAlt;
    ctx.fillRect(0, 0, cssW, cssH);

    var dur = v && isFinite(v.duration) ? v.duration : 0;
    if (dur <= 0) {
      _markerHitRects = [];
      ctx.fillStyle = textDim;
      ctx.font = "11px -apple-system, sans-serif";
      ctx.textAlign = "center";
      ctx.fillText(state.selectedParticipant ? "Loading…" : "No video", cssW / 2, cssH / 2 + 4);
      ctx.textAlign = "start";
      renderPlayhead();
      return;
    }

    function timeToX(t) { return (t / dur) * cssW; }

    var tickInterval = computeTickInterval(dur);
    var firstTick = Math.ceil(0 / tickInterval) * tickInterval;
    ctx.strokeStyle = border;
    ctx.fillStyle = textDim;
    ctx.font = "10px " + getCSSVar("--font-mono", "monospace");
    ctx.textAlign = "center";
    ctx.lineWidth = 1;
    for (var t = firstTick; t <= dur; t += tickInterval) {
      var x = timeToX(t);
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, 6);
      ctx.stroke();
      ctx.fillText(formatTime(t), x, 16);
    }
    ctx.textAlign = "start";

    var markerY = 22;
    var markerH = cssH - markerY - 4;
    _markerHitRects = [];

    for (var i = 0; i < state.segments.length; i++) {
      var seg = state.segments[i];
      var mark = getMarkForSegment(seg);
      if (!mark) continue;
      var cat = MARK_CATEGORIES[mark.category] || MARK_CATEGORIES.bookmark || { color: "#0891b2" };
      var color = mark.color || cat.color;
      var startX = timeToX(seg.start);
      var endX = seg.end ? timeToX(seg.end) : startX + 2;
      var barX = Math.max(0, startX - 1);
      var barW = Math.max(2, Math.min(endX - startX, 6));
      ctx.fillStyle = color;
      ctx.fillRect(barX, markerY, barW, markerH);
      _markerHitRects.push({
        x1: barX - 3,
        x2: barX + barW + 3,
        y: markerY,
        h: markerH,
        segIndex: i,
      });
    }

    renderPlayhead();
  }

  function renderPlayhead() {
    var canvas = qs("#playheadCanvas");
    var v = qs("#videoPlayer");
    if (!canvas || !v) return;
    var cssW = canvas.offsetWidth;
    var cssH = canvas.offsetHeight;
    var ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, cssW, cssH);
    var dur = isFinite(v.duration) ? v.duration : 0;
    if (dur <= 0) return;
    var px = (v.currentTime / dur) * cssW;
    var accent = getCSSVar("--color-accent", "#1d4f72");
    ctx.strokeStyle = accent;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(px, 0);
    ctx.lineTo(px, cssH);
    ctx.stroke();
    ctx.fillStyle = accent;
    ctx.beginPath();
    ctx.moveTo(px - 4, 0);
    ctx.lineTo(px + 4, 0);
    ctx.lineTo(px, 5);
    ctx.closePath();
    ctx.fill();
  }

  function timelineXToTime(event) {
    var canvas = qs("#timelineCanvas");
    var v = qs("#videoPlayer");
    if (!canvas || !v) return null;
    var dur = isFinite(v.duration) ? v.duration : 0;
    if (dur <= 0) return null;
    var rect = canvas.getBoundingClientRect();
    var frac = (event.clientX - rect.left) / rect.width;
    if (frac < 0) frac = 0; else if (frac > 1) frac = 1;
    return frac * dur;
  }

  // ---- Transcript timeline canvas (marks, playhead, hover hit-testing) ----

  function hitTestTimeline(clientX, clientY) {
    var canvas = qs("#timelineCanvas");
    if (!canvas) return null;
    var rect = canvas.getBoundingClientRect();
    var mx = clientX - rect.left;
    var my = clientY - rect.top;
    for (var i = _markerHitRects.length - 1; i >= 0; i--) {
      var hr = _markerHitRects[i];
      if (mx >= hr.x1 && mx <= hr.x2 && my >= hr.y && my <= hr.y + hr.h) return hr;
    }
    return null;
  }

  function showTimelineTooltip(hit, clientX, clientY) {
    var tip = qs("#trTooltip");
    if (!tip) return;
    var seg = state.segments[hit.segIndex];
    if (!seg) return;
    var mark = getMarkForSegment(seg);
    var cat = (mark && MARK_CATEGORIES[mark.category]) || MARK_CATEGORIES.bookmark || { label: "Mark", color: "#888" };
    var snippet = (seg.text || "").trim().slice(0, 80);
    if ((seg.text || "").length > 80) snippet += "…";
    var extraCount = (seg.marks && seg.marks.length > 1) ? (seg.marks.length - 1) : 0;
    var label = mark && mark.label ? " — " + mark.label : "";
    tip.textContent = "";
    var catSpan = el("span", "tr-tooltip-cat", cat.label);
    // Set color via property API rather than string-interpolating into a
    // style attribute — mark.color comes from a stash/manifest file and a
    // crafted value (e.g. `red" onmouseover=...`) would otherwise break out.
    catSpan.style.color = (mark && mark.color) || cat.color || "";
    tip.appendChild(catSpan);
    tip.appendChild(document.createTextNode(formatTime(seg.start) + label));
    if (extraCount > 0) {
      tip.appendChild(document.createTextNode(" "));
      var extraSpan = el("span", "", "+" + extraCount + " more");
      extraSpan.style.opacity = ".7";
      tip.appendChild(extraSpan);
    }
    tip.appendChild(document.createElement("br"));
    tip.appendChild(document.createTextNode(snippet));
    tip.classList.remove("hidden");
    var tipRect = tip.getBoundingClientRect();
    var x = clientX + 12;
    var y = clientY - tipRect.height - 12;
    if (x + tipRect.width > window.innerWidth - 8) x = window.innerWidth - tipRect.width - 8;
    if (y < 8) y = clientY + 16;
    tip.style.left = x + "px";
    tip.style.top = y + "px";
  }

  function hideTimelineTooltip() {
    var tip = qs("#trTooltip");
    if (tip) tip.classList.add("hidden");
  }

  function onMarkerClick(hit) {
    var seg = state.segments[hit.segIndex];
    if (!seg) return;
    seekVideo(seg.start);
    if (!_cachedSegmentRows) {
      _cachedSegmentRows = qs("#segmentList").querySelectorAll(".segment-row");
    }
    var row = _cachedSegmentRows[hit.segIndex];
    if (row) scrollToSegment(row);
  }

  function initVideoPlayer() {
    var video = qs("#videoPlayer");
    if (!video) return;

    // Restore CC preference. We restore via getStoredUIState rather than a
    // module-level constant so that switching browsers / clearing storage
    // gives the user the documented default (off).
    var stored = getStoredUIState("transcripts");
    state.ccEnabled = !!(stored && stored.ccEnabled);
    state.videoCollapsed = !!(stored && stored.videoCollapsed);
    var section = qs("#videoSection");
    if (section) section.classList.toggle("video-collapsed", state.videoCollapsed);

    qs("#videoPlayBtn").addEventListener("click", function () {
      if (video.paused) video.play();
      else video.pause();
    });
    qs("#videoMuteBtn").addEventListener("click", function () {
      state.videoMuted = !state.videoMuted;
      video.muted = state.videoMuted;
      updatePlayerButtons();
    });
    qs("#videoSpeedBtn").addEventListener("click", function () {
      var idx = VIDEO_SPEEDS.indexOf(state.videoPlaybackRate);
      state.videoPlaybackRate = VIDEO_SPEEDS[(idx + 1) % VIDEO_SPEEDS.length];
      applyPlaybackRate();
      updatePlayerButtons();
    });
    qs("#videoCcBtn").addEventListener("click", function () {
      state.ccEnabled = !state.ccEnabled;
      applyCaptionMode();
      setStoredUIStateField("transcripts", "ccEnabled", state.ccEnabled);
      updatePlayerButtons();
    });
    qs("#videoPipBtn").addEventListener("click", function () {
      state.pipEnabled = !state.pipEnabled;
      // When the user disables PiP while the player is detached, drop it
      // back into flow immediately. _setPipActive is provided by initPipScroll
      // and undefined until then; guard so the button still toggles state
      // even if scroll wiring failed for some reason.
      if (!state.pipEnabled && state.pipActive && typeof _setPipActive === "function") {
        _setPipActive(false);
      }
      updatePlayerButtons();
    });
    qs("#videoCollapseBtn").addEventListener("click", function () {
      state.videoCollapsed = !state.videoCollapsed;
      var sec = qs("#videoSection");
      if (sec) sec.classList.toggle("video-collapsed", state.videoCollapsed);
      setStoredUIStateField("transcripts", "videoCollapsed", state.videoCollapsed);
      updatePlayerButtons();
    });

    video.addEventListener("play", function () {
      state.videoPlaying = true;
      updatePlayerButtons();
    });
    video.addEventListener("pause", function () {
      state.videoPlaying = false;
      updatePlayerButtons();
    });
    video.addEventListener("ended", function () {
      state.videoPlaying = false;
      updatePlayerButtons();
    });
    video.addEventListener("loadedmetadata", function () {
      sizeTimelineCanvas();
      applyCaptionMode();
      applyPlaybackRate();
      updateTimeLabel();
      renderTimeline();
    });

    // The <track> can finish loading after the video metadata; re-apply the
    // caption mode once cues are parsed or some browsers render nothing.
    var track = qs("#subtitleTrack");
    if (track) {
      track.addEventListener("load", applyCaptionMode);
    }
    video.addEventListener("durationchange", function () {
      updateTimeLabel();
      renderTimeline();
    });
    video.addEventListener("timeupdate", function () {
      updateTimeLabel();
      if (_playheadRaf) return;
      _playheadRaf = requestAnimationFrame(function () {
        _playheadRaf = 0;
        renderPlayhead();
      });
    });

    // Keep the paused frame visible across tab switches. See utils.js.
    clipgenInstallPausedFrameOverlay(video);

    updatePlayerButtons();
  }

  function initTimelineCanvas() {
    var canvas = qs("#timelineCanvas");
    if (!canvas) return;
    sizeTimelineCanvas();

    if (typeof ResizeObserver === "function") {
      _timelineResizeObs = new ResizeObserver(function () {
        sizeTimelineCanvas();
        renderTimeline();
      });
      _timelineResizeObs.observe(qs("#timelineCanvasWrapper"));
      window.addEventListener("pagehide", function () {
        if (_timelineResizeObs) { _timelineResizeObs.disconnect(); _timelineResizeObs = null; }
      });
    } else {
      window.addEventListener("resize", function () {
        sizeTimelineCanvas();
        renderTimeline();
      });
    }

    canvas.addEventListener("click", function (e) {
      var hit = hitTestTimeline(e.clientX, e.clientY);
      if (hit) {
        onMarkerClick(hit);
        return;
      }
      var t = timelineXToTime(e);
      if (t !== null) seekVideo(t);
    });

    canvas.addEventListener("mousemove", function (e) {
      if (_timelineTooltipRaf) return;
      var cx = e.clientX;
      var cy = e.clientY;
      _timelineTooltipRaf = requestAnimationFrame(function () {
        _timelineTooltipRaf = 0;
        var hit = hitTestTimeline(cx, cy);
        if (hit) {
          _lastTimelineHit = hit;
          showTimelineTooltip(hit, cx, cy);
          canvas.style.cursor = "pointer";
        } else if (_lastTimelineHit) {
          _lastTimelineHit = null;
          hideTimelineTooltip();
          canvas.style.cursor = "pointer";
        }
      });
    });
    canvas.addEventListener("mouseleave", function () {
      _lastTimelineHit = null;
      hideTimelineTooltip();
    });
  }

  // ---- PiP scroll behaviour ----

  // Hoisted so initVideoPlayer's PiP-toggle handler can drop the player back
  // into flow when the user disables PiP. Assigned in initPipScroll.
  var _setPipActive = null;

  // Chrome strip height — TopNav (48) + subheader (44) + pill bar (56). Mirrored
  // in transcripts.css `#trMain { padding-top: 148px }`; PiP compensation has
  // to add the videoSection height on top of this so layout doesn't jump.
  var TR_CHROME_TOP = 148;

  function initPipScroll() {
    var section = qs("#videoSection");
    // #trMain is the scroll container (pass 6 floating-nav scroll-under) so
    // scrolled content slides under the chrome strip. PiP triggers when the
    // user scrolls past a commit; only returning to the very top dismisses
    // it. Asymmetric thresholds avoid the bounce caused by switching the
    // video section to position:fixed mid-scroll.
    var scroller = qs("#trMain");
    if (!section || !scroller) return;

    var ENTER_THRESHOLD = 140;
    var scrollRaf = 0;

    function setPipActive(active) {
      if (active === state.pipActive) return;
      if (active) {
        // Reserve the section's natural height on top of the chrome inset so
        // transcript content does not jump up when the player detaches from
        // flow. We restore the scroll position on the next frame because the
        // browser may rebase scrollTop relative to the new content size when
        // padding is added.
        var h = Math.round(section.getBoundingClientRect().height);
        if (h > 0) scroller.style.paddingTop = TR_CHROME_TOP + h + "px";
        var keepTop = scroller.scrollTop;
        state.pipActive = true;
        section.classList.add("pip");
        requestAnimationFrame(function () {
          scroller.scrollTop = keepTop;
          sizeTimelineCanvas();
          renderTimeline();
        });
      } else {
        var keepTop2 = scroller.scrollTop;
        state.pipActive = false;
        section.classList.remove("pip");
        // Empty inline override falls back to the CSS default (148px).
        scroller.style.paddingTop = "";
        requestAnimationFrame(function () {
          scroller.scrollTop = keepTop2;
          sizeTimelineCanvas();
          renderTimeline();
        });
      }
    }
    _setPipActive = setPipActive;

    function evaluatePip() {
      if (!state.pipEnabled) return;
      var top = scroller.scrollTop;
      if (state.pipActive) {
        // Only release PiP once the user is back at the very top, so the user
        // makes a deliberate "scroll up" or "click PiP" gesture rather than
        // having the player drop out as soon as they ease back.
        if (top <= 0) setPipActive(false);
      } else {
        if (top > ENTER_THRESHOLD) setPipActive(true);
      }
    }

    scroller.addEventListener("scroll", function () {
      if (scrollRaf) return;
      scrollRaf = requestAnimationFrame(function () {
        scrollRaf = 0;
        evaluatePip();
      });
    }, { passive: true });

    section.addEventListener("click", function (e) {
      if (!state.pipActive) return;
      if (e.target.closest(".player-btn") || e.target.closest("#timelineCanvasWrapper")) return;
      scroller.scrollTo({ top: 0, behavior: "smooth" });
    });
  }

  // ---- Keyboard ----

  function initPlayerKeyboard() {
    document.addEventListener("keydown", function (e) {
      if (e.code !== "Space" && e.key !== " ") return;
      if (e.metaKey || e.ctrlKey || e.altKey) return;
      var t = e.target;
      if (t && t.matches && t.matches("input, textarea, select, [contenteditable=true]")) return;
      if (state.editingTextEl) return;
      var modal = qs("#correctionsModal");
      if (modal && !modal.classList.contains("hidden")) return;
      var pop = qs("#markPopover");
      if (pop && !pop.classList.contains("hidden")) return;
      var v = qs("#videoPlayer");
      if (!v || !v.src) return;
      e.preventDefault();
      if (v.paused) v.play();
      else v.pause();
    });
  }

  var _pendingSeekTime = null;
  var _seekRaf = 0;
  var _pendingSeekListener = null;

  function seekVideo(time) {
    var video = qs("#videoPlayer");
    if (!video || !video.src) return;

    // Remove any previous deferred-seek listener
    if (_pendingSeekListener) {
      video.removeEventListener("loadedmetadata", _pendingSeekListener);
      _pendingSeekListener = null;
    }

    // If metadata hasn't loaded yet, defer the seek
    if (video.readyState < 1) {
      _pendingSeekTime = time;
      _pendingSeekListener = function () {
        video.removeEventListener("loadedmetadata", _pendingSeekListener);
        _pendingSeekListener = null;
        var t = _pendingSeekTime;
        _pendingSeekTime = null;
        if (t !== null) seekVideo(t);
      };
      video.addEventListener("loadedmetadata", _pendingSeekListener);
      return;
    }

    // Coalesce rapid seeks into one per animation frame
    _pendingSeekTime = time;
    cancelAnimationFrame(_seekRaf);
    _seekRaf = requestAnimationFrame(function () {
      var t = _pendingSeekTime;
      _pendingSeekTime = null;
      _seekRaf = 0;
      if (t === null) return;
      video.currentTime = t;
      if (video.paused) video.play();
    });
  }

  // ---- Video sync ----

  var _syncRaf = 0;

  function initVideoSync() {
    var video = qs("#videoPlayer");
    var save = function () { persistVideoTime(video.currentTime); };
    video.addEventListener("timeupdate", function () {
      save();
      if (_syncRaf) return;
      _syncRaf = requestAnimationFrame(function () {
        _syncRaf = 0;
        highlightActiveSegment();
      });
    });
    // Native scrubbing on a paused video doesn't always fire timeupdate.
    video.addEventListener("seeked", save);
  }

  function persistVideoTime(t) {
    if (!state.selectedParticipant || !isFinite(t)) return;
    var stored = getStoredUIState("transcripts");
    var map = (stored.videoTimeByParticipant && typeof stored.videoTimeByParticipant === "object")
      ? stored.videoTimeByParticipant : {};
    map[state.selectedParticipant] = t;
    setStoredUIStateField("transcripts", "videoTimeByParticipant", map);
  }

  function highlightActiveSegment() {
    var video = qs("#videoPlayer");
    if (!video || !video.src) return;
    var t = video.currentTime;

    // Binary search for active segment (sorted, non-overlapping)
    var lo = 0, hi = state.segments.length - 1, newIndex = -1;
    while (lo <= hi) {
      var mid = (lo + hi) >> 1;
      if (state.segments[mid].start <= t) { newIndex = mid; lo = mid + 1; }
      else { hi = mid - 1; }
    }
    if (newIndex >= 0 && t >= state.segments[newIndex].end) newIndex = -1;

    if (newIndex === state.activeSegmentIndex) return;

    // Cache segment row elements
    if (!_cachedSegmentRows) {
      _cachedSegmentRows = qs("#segmentList").querySelectorAll(".segment-row");
    }
    var rows = _cachedSegmentRows;

    // Remove old active
    if (state.activeSegmentIndex >= 0 && state.activeSegmentIndex < rows.length) {
      rows[state.activeSegmentIndex].classList.remove("active");
    }

    // Set new active
    state.activeSegmentIndex = newIndex;
    if (newIndex >= 0 && newIndex < rows.length) {
      rows[newIndex].classList.add("active");
      // Only chase the playhead when PiP is active. With the embedded player
      // visible at the top, the user is driving via the timeline and doesn't
      // want the transcript pulled away from what they're reading.
      if (state.pipActive) scrollToSegment(rows[newIndex]);
    }
  }

  function scrollToSegment(row) {
    // Pass 6 floating-nav scroll-under: #trMain is the scroll container; the
    // top 148px of its viewport sits *under* the fixed chrome strip, so the
    // visible top edge is at scroller.top + TR_CHROME_TOP, not at scroller.top.
    var scroller = qs("#trMain");
    if (!scroller) return;
    var rowRect = row.getBoundingClientRect();
    var scRect = scroller.getBoundingClientRect();
    var rowTopInScroll = rowRect.top - scRect.top + scroller.scrollTop;
    var rowBottomInScroll = rowTopInScroll + rowRect.height;
    var visibleTop = scroller.scrollTop + TR_CHROME_TOP;
    var visibleBottom = scroller.scrollTop + scroller.clientHeight;

    if (rowTopInScroll < visibleTop + 40) {
      scroller.scrollTop = rowTopInScroll - TR_CHROME_TOP - 40;
    } else if (rowBottomInScroll > visibleBottom - 40) {
      scroller.scrollTop = rowBottomInScroll - scroller.clientHeight + 40;
    }
  }

  // ---- Inline segment editing ----

  function startSegmentEditing(textEl) {
    if (state.editingTextEl === textEl) return;
    if (state.editingTextEl) finishSegmentEditing(state.editingTextEl, false);

    state.editingTextEl = textEl;
    textEl.setAttribute("data-original-text", textEl.textContent);
    textEl.setAttribute("contenteditable", "true");
    textEl.setAttribute("spellcheck", "false");
    textEl.classList.add("segment-text-editing");

    function onBlur() {
      textEl.removeEventListener("blur", onBlur);
      textEl.removeEventListener("keydown", onKeydown);
      textEl.removeEventListener("paste", onPaste);
      finishSegmentEditing(textEl, false);
    }

    function onKeydown(e) {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        textEl.blur();
      }
      if (e.key === "Escape") {
        e.preventDefault();
        textEl.removeEventListener("blur", onBlur);
        textEl.removeEventListener("keydown", onKeydown);
        textEl.removeEventListener("paste", onPaste);
        finishSegmentEditing(textEl, true);
      }
    }

    function onPaste(e) {
      e.preventDefault();
      var text = (e.clipboardData || window.clipboardData).getData("text/plain");
      text = text.replace(/[\r\n]+/g, " ").trim();
      document.execCommand("insertText", false, text);
    }

    textEl.addEventListener("blur", onBlur);
    textEl.addEventListener("keydown", onKeydown);
    textEl.addEventListener("paste", onPaste);
  }

  function finishSegmentEditing(textEl, cancel) {
    var originalText = textEl.getAttribute("data-original-text") || "";
    var newText = textEl.textContent.trim();

    textEl.removeAttribute("contenteditable");
    textEl.removeAttribute("data-original-text");
    textEl.classList.remove("segment-text-editing");
    if (state.editingTextEl === textEl) state.editingTextEl = null;

    if (cancel || !newText || newText === originalText) {
      // During streaming, skip reload — next poll will re-render
      if (state.streamingParticipant) return;
      // Reload to restore clean word spans
      var pid = state.selectedParticipant;
      if (pid) loadTranscript(pid);
      return;
    }

    var corrections = extractCorrections(originalText, newText);
    if (corrections.length === 0) return;

    saveCorrections(corrections);
  }

  function extractCorrections(oldText, newText) {
    var oldWords = oldText.trim().split(/\s+/).filter(Boolean);
    var newWords = newText.trim().split(/\s+/).filter(Boolean);
    if (oldWords.join(" ") === newWords.join(" ")) return [];

    // LCS table for word-level alignment
    var m = oldWords.length, n = newWords.length;
    var dp = [];
    for (var i = 0; i <= m; i++) {
      dp[i] = [];
      for (var j = 0; j <= n; j++) {
        if (i === 0 || j === 0) dp[i][j] = 0;
        else if (oldWords[i - 1] === newWords[j - 1]) dp[i][j] = dp[i - 1][j - 1] + 1;
        else dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }

    // Backtrack to get edit operations
    var ops = [];
    i = m; j = n;
    while (i > 0 || j > 0) {
      if (i > 0 && j > 0 && oldWords[i - 1] === newWords[j - 1]) {
        ops.unshift({ type: "eq" });
        i--; j--;
      } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
        ops.unshift({ type: "ins", word: newWords[j - 1] });
        j--;
      } else {
        ops.unshift({ type: "del", word: oldWords[i - 1] });
        i--;
      }
    }

    // Group consecutive non-equal ops into from→to correction pairs
    var corrections = [];
    var k = 0;
    while (k < ops.length) {
      if (ops[k].type !== "eq") {
        var fromParts = [];
        var toParts = [];
        while (k < ops.length && ops[k].type !== "eq") {
          if (ops[k].type === "del") fromParts.push(ops[k].word);
          else toParts.push(ops[k].word);
          k++;
        }
        if (fromParts.length > 0 && toParts.length > 0) {
          corrections.push({ from: fromParts.join(" "), to: toParts.join(" ") });
        }
      } else {
        k++;
      }
    }
    return corrections;
  }

  function saveCorrections(corrections) {
    var updated = 0, removed = 0;
    var chain = Promise.resolve();
    corrections.forEach(function (c) {
      chain = chain.then(function () {
        return apiPost("api/corrections", { from: c.from, to: c.to }).then(function (data) {
          if (data.ok) {
            if (data.removed) removed++;
            else if (data.correction) updated++;  // covers both new and updated
          }
        });
      });
    });
    chain.then(function () {
      var parts = [];
      if (updated) parts.push(updated === 1 ? "1 correction saved" : updated + " corrections saved");
      if (removed) parts.push(removed === 1 ? "1 reverted" : removed + " reverted");
      showToast(parts.join(", ") || "No changes");
      // During streaming, skip reload — corrections are persisted and will apply on completion
      if (state.streamingParticipant) return;
      var pid = state.selectedParticipant;
      if (pid) {
        loadTranscript(pid);
        loadCorrections();
      }
    }).catch(function () {
      showToast("Failed to save correction");
    });
  }

  // ---- Marks ----

  function toggleMark(segmentId) {
    apiPost("api/marks", {
      segment_ids: [segmentId],
      category: state.lastMarkCategory,
    }).then(function (data) {
      if (data.ok) {
        showToast("Marked");
        if (state.selectedParticipant) loadTranscript(state.selectedParticipant);
      }
    });
  }

  function toggleMarkStreaming(segmentId, markEl) {
    var cat = MARK_CATEGORIES[state.lastMarkCategory] || MARK_CATEGORIES.bookmark;
    apiPost("api/marks", {
      segment_ids: [segmentId],
      category: state.lastMarkCategory,
    }).then(function (data) {
      if (data.ok && data.marks && data.marks.length > 0) {
        var m = data.marks[0];
        showToast("Marked");
        markEl.classList.add("marked");
        markEl.style.background = cat.color;
        _streamingMarks[segmentId] = {
          color: cat.color,
          id: m.id,
          category: m.category,
          label: m.label || "",
        };
        _bumpStreamingMarksVersion();
      }
    });
  }

  function removeMark(markId) {
    apiDelete("api/marks/" + markId).then(function (data) {
      if (data.ok) {
        showToast("Mark removed");
        hideMarkPopover();
        if (state.streamingParticipant) {
          for (var key in _streamingMarks) {
            if (_streamingMarks[key].id === markId) {
              delete _streamingMarks[key];
              _bumpStreamingMarksVersion();
              break;
            }
          }
          pollTaskStatus();
        } else if (state.selectedParticipant) {
          loadTranscript(state.selectedParticipant);
        }
      }
    });
  }

  function updateMarkCategory(markId, category) {
    state.lastMarkCategory = category;
    apiPut("api/marks/" + markId, { category: category }).then(function (data) {
      if (data.ok) {
        hideMarkPopover();
        if (state.streamingParticipant) {
          var cat = MARK_CATEGORIES[category] || MARK_CATEGORIES.bookmark;
          for (var key in _streamingMarks) {
            if (_streamingMarks[key].id === markId) {
              _streamingMarks[key].category = category;
              _streamingMarks[key].color = cat.color;
              _bumpStreamingMarksVersion();
              break;
            }
          }
          pollTaskStatus();
        } else if (state.selectedParticipant) {
          loadTranscript(state.selectedParticipant);
        }
      }
    });
  }

  function updateMarkLabel(markId, label) {
    apiPut("api/marks/" + markId, { label: label || null });
    if (state.streamingParticipant) {
      for (var key in _streamingMarks) {
        if (_streamingMarks[key].id === markId) {
          _streamingMarks[key].label = label || "";
          break;
        }
      }
    }
  }

  function showMarkPopover(anchorEl, segmentId, markObj) {
    var popover = qs("#markPopover");
    hideMarkPopover();

    // Build category pills
    var catContainer = popover.querySelector(".mark-popover-categories");
    catContainer.innerHTML = "";
    var cats = Object.keys(MARK_CATEGORIES);
    for (var i = 0; i < cats.length; i++) {
      (function (key) {
        var cat = MARK_CATEGORIES[key];
        var pill = document.createElement("button");
        pill.className = "mark-cat-pill" + (markObj.category === key ? " active" : "");
        pill.style.background = cat.color;
        pill.title = cat.label;
        pill.addEventListener("click", function (e) {
          e.stopPropagation();
          updateMarkCategory(markObj.id, key);
        });
        catContainer.appendChild(pill);
      })(cats[i]);
    }

    // Label input
    var labelInput = popover.querySelector(".mark-popover-label");
    labelInput.value = markObj.label || "";
    labelInput._markId = markObj.id;
    labelInput.onblur = function () {
      var val = labelInput.value.trim();
      if (val !== (markObj.label || "")) {
        updateMarkLabel(markObj.id, val);
      }
    };
    labelInput.onkeydown = function (e) {
      if (e.key === "Enter") { e.preventDefault(); labelInput.blur(); hideMarkPopover(); }
      if (e.key === "Escape") { e.preventDefault(); hideMarkPopover(); }
    };

    // Remove button
    var removeBtn = popover.querySelector(".mark-popover-remove");
    removeBtn.onclick = function (e) {
      e.stopPropagation();
      removeMark(markObj.id);
    };

    // Position below anchor
    var rect = anchorEl.getBoundingClientRect();
    popover.style.top = (rect.bottom + window.scrollY + 4) + "px";
    popover.style.left = (rect.left + window.scrollX - 4) + "px";
    popover.classList.remove("hidden");

    // Close on outside click (deferred so this click doesn't trigger it).
    // Track the timeout so a fast hideMarkPopover (e.g. Esc within the same
    // tick) cancels the pending attach instead of leaving a permanently
    // attached listener after the deferred .addEventListener fires.
    if (_popoverAttachTimer) clearTimeout(_popoverAttachTimer);
    document.removeEventListener("click", _popoverOutsideClick);
    _popoverAttachTimer = setTimeout(function () {
      _popoverAttachTimer = null;
      document.addEventListener("click", _popoverOutsideClick);
    }, 0);
  }

  function _popoverOutsideClick(e) {
    var popover = qs("#markPopover");
    if (popover && !popover.contains(e.target)) {
      hideMarkPopover();
    }
  }

  var _popoverAttachTimer = null;

  function hideMarkPopover() {
    var popover = qs("#markPopover");
    if (popover) popover.classList.add("hidden");
    if (_popoverAttachTimer) {
      clearTimeout(_popoverAttachTimer);
      _popoverAttachTimer = null;
    }
    document.removeEventListener("click", _popoverOutsideClick);
  }

  function markAllSearchResults() {
    if (!state.searchResults || !state.searchResults.results) return;
    var ids = [];
    state.searchResults.results.forEach(function (r) {
      if (r.segment_id) ids.push(r.segment_id);
    });
    if (ids.length === 0) return;
    apiPost("api/marks", {
      segment_ids: ids,
      category: state.lastMarkCategory,
    }).then(function (data) {
      if (data.ok) {
        showToast("Marked " + clipgenPluralUnit(data.marks.length, "segment", "segments"));
        if (state.selectedParticipant) loadTranscript(state.selectedParticipant);
      }
    });
  }

  // ---- Search ----

  var _searchTimer = null;

  function initSearch() {
    var input = qs("#searchInput");
    input.addEventListener("input", function () {
      clearTimeout(_searchTimer);
      var q = input.value.trim();
      if (q.length < 2) {
        hideSearchResults();
        return;
      }
      _searchTimer = setTimeout(function () { doSearch(q); }, SEARCH_DEBOUNCE);
    });

    input.addEventListener("keydown", function (e) {
      if (e.key === "Escape") {
        input.value = "";
        hideSearchResults();
      }
    });

    // Close search results when clicking outside
    document.addEventListener("click", function (e) {
      var searchArea = qs("#headerSearch");
      if (searchArea && !searchArea.contains(e.target)) {
        hideSearchResults();
      }
    });
  }

  function doSearch(query) {
    state.searchQuery = query;
    apiGet("api/search?q=" + encodeURIComponent(query)).then(function (data) {
      if (!data.ok) return;
      // Merge client-side search of partial segments for the streaming participant
      if (state.streamingParticipant) {
        var partials = _searchPartialSegments(query, state.streamingParticipant);
        if (partials.length > 0) {
          data.results = data.results.concat(partials);
          data.total_count += partials.length;
          data.counts_by_participant[state.streamingParticipant] =
            (data.counts_by_participant[state.streamingParticipant] || 0) + partials.length;
        }
      }
      state.searchResults = data;
      renderSearchResults(data);
    });
  }

  function _searchPartialSegments(query, pid) {
    var results = [];
    var lowerQ = query.toLowerCase();
    var task = null;
    state.tasks.forEach(function (t) {
      if (t.participant === pid && t.status === "running" && t.partial_segments) {
        task = t;
      }
    });
    if (!task) return results;
    for (var i = 0; i < task.partial_segments.length; i++) {
      var seg = task.partial_segments[i];
      if (seg.text.toLowerCase().indexOf(lowerQ) >= 0) {
        results.push({
          participant: pid,
          segment_id: pid + ":" + i,
          start: seg.start,
          end: seg.end,
          text: seg.text,
          count: 1,
        });
      }
    }
    return results;
  }

  function renderSearchResults(data) {
    var container = qs("#searchResults");
    var countEl = qs("#searchCount");

    if (data.total_count === 0) {
      countEl.textContent = "0 results";
      container.innerHTML = '<div class="search-result-row" style="justify-content:center;color:var(--color-text-dim)">No matches found</div>';
      container.classList.remove("hidden");
      return;
    }

    countEl.textContent = data.total_count + " match" + (data.total_count === 1 ? "" : "es");

    // Add "Mark All" button next to count
    var markAllBtn = qs("#searchMarkAllBtn");
    if (!markAllBtn) {
      markAllBtn = document.createElement("button");
      markAllBtn.id = "searchMarkAllBtn";
      markAllBtn.className = "btn btn-small";
      markAllBtn.textContent = "Mark All";
      countEl.parentNode.insertBefore(markAllBtn, countEl.nextSibling);
    }
    markAllBtn.classList.remove("hidden");
    markAllBtn.onclick = function () { markAllSearchResults(); };

    // Group results by participant
    var groups = {};
    var order = [];
    data.results.forEach(function (r) {
      if (!groups[r.participant]) {
        groups[r.participant] = [];
        order.push(r.participant);
      }
      groups[r.participant].push(r);
    });

    var html = "";
    order.forEach(function (pid) {
      var count = data.counts_by_participant[pid] || 0;
      html += '<div class="search-group-header">' + escapeHtml(pid) + ' (' + count + ')</div>';
      groups[pid].forEach(function (r) {
        var xref = findOverlapsForSearch(r.participant, r.start, r.end);
        html += '<div class="search-result-row" data-participant="' + escapeHtml(r.participant) + '" data-start="' + r.start + '">';
        html += '<span class="search-result-time">' + formatTime(r.start) + '</span>';
        html += '<span class="search-result-text">' + highlightQuery(r.text, state.searchQuery) + '</span>';
        if (state.tooltipsEnabled && xref.screenspaceEvents.length > 0) {
          var seen = {};
          html += '<span class="search-xref-events">';
          for (var ei = 0; ei < xref.screenspaceEvents.length; ei++) {
            var det = xref.screenspaceEvents[ei].detector;
            if (seen[det]) continue;
            seen[det] = true;
            html += '<span class="search-xref-dot" style="background:var(--color-task-' + det + ', #888)" title="' + escapeHtml(xref.screenspaceEvents[ei].event_type || det) + '"></span>';
          }
          html += '</span>';
        }
        if (state.tooltipsEnabled && xref.sheetObservations.length > 0) {
          var obsText = xref.sheetObservations[0].observation;
          var truncObs = obsText.length > 50 ? obsText.substring(0, 50) + "\u2026" : obsText;
          html += '<span class="search-xref-sheet" title="' + escapeHtml(obsText) + '">' + escapeHtml(truncObs) + '</span>';
        }
        html += '</div>';
      });
    });

    var prevScrollTop = container.scrollTop;
    container.innerHTML = html;
    container.scrollTop = prevScrollTop;
    container.classList.remove("hidden");

    // Attach click handlers
    var rows = container.querySelectorAll(".search-result-row[data-participant]");
    for (var i = 0; i < rows.length; i++) {
      rows[i].addEventListener("click", function () {
        var pid = this.getAttribute("data-participant");
        var start = parseFloat(this.getAttribute("data-start"));
        jumpToResult(pid, start);
      });
    }
  }

  function highlightQuery(text, query) {
    if (!query) return escapeHtml(text);
    var escaped = escapeHtml(text);
    var queryEscaped = escapeHtml(query);
    var regex = new RegExp("(" + queryEscaped.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")", "gi");
    return escaped.replace(regex, '<span class="search-highlight">$1</span>');
  }

  function hideSearchResults() {
    qs("#searchResults").classList.add("hidden");
    qs("#searchCount").textContent = "";
    var markAllBtn = qs("#searchMarkAllBtn");
    if (markAllBtn) markAllBtn.classList.add("hidden");
    state.searchResults = null;
  }

  function jumpToResult(pid, start) {
    hideSearchResults();
    if (pid !== state.selectedParticipant) selectParticipant(pid);
    seekVideo(start);
  }

  // ---- Participant pills ----
  // Per-participant transcribe control; chevron opens a body-mounted popover
  // (togglePillOptions → buildPillOptions).

  // Icon shown on the trigger by default (left), and on hover (right).
  // The trigger click always invokes the appropriate action for the status.
  var PILL_TRIGGER = {
    idle: { rest: "icons/microphone.svg", hover: "icons/microphone.svg", label: "Transcribe", action: "transcribe" },
    failed: { rest: "icons/exclamation-triangle.svg", hover: "icons/microphone.svg", label: "Retry transcription", action: "transcribe" },
    queued: { rest: "icons/clock.svg", hover: "icons/stop-circle.svg", label: "Cancel", action: "cancel" },
    running: { rest: "icons/arrow-path.svg", hover: "icons/stop-circle.svg", label: "Cancel transcription", action: "cancel" },
    completed: { rest: "icons/check-circle.svg", hover: "icons/arrow-path.svg", label: "Re-transcribe", action: "retranscribe" }
  };

  var PILL_LANGUAGES = [
    { code: "", label: "Auto-detect" },
    { code: "en", label: "English" },
    { code: "sv", label: "Swedish" },
    { code: "es", label: "Spanish" },
    { code: "fr", label: "French" },
    { code: "de", label: "German" },
    { code: "it", label: "Italian" },
    { code: "pt", label: "Portuguese" },
    { code: "nl", label: "Dutch" },
    { code: "no", label: "Norwegian" },
    { code: "da", label: "Danish" },
    { code: "fi", label: "Finnish" }
  ];

  // UI-only state for pills (reset on reload)
  state.pillOverrides = {};    // pid → { model, language }
  state.pillOptionsOpen = null; // pid of the pill whose options pane is open

  function _dotStateTranscription(p, task) {
    if (task && (task.status === "running" || task.status === "queued")) return "running";
    if (task && task.status === "failed") return "failed";
    if (p.has_transcript) return "done";
    return "idle";
  }

  function pillState(p, taskByPid) {
    var task = taskByPid[p.id];
    var status = "idle";
    var progress = 0;
    var taskId = null;

    if (task && (task.status === "running" || task.status === "queued")) {
      status = task.status;
      taskId = task.id;
      if (task.status === "running") progress = Math.round((task.progress || 0) * 100);
    } else if (task && task.status === "failed") {
      status = "failed";
      taskId = task.id;
    } else if (p.has_transcript) {
      status = "completed";
      progress = 100;
    }
    var agents = {
      transcription: _dotStateTranscription(p, task),
      summary: (p.agents && p.agents.summary) || "idle",
      citations: (p.agents && p.agents.citations) || "idle",
    };
    return { status: status, progress: progress, taskId: taskId, agents: agents };
  }

  function renderPills() {
    var container = qs("#participantPills");
    if (!container) return;

    if (state.participants.length === 0) {
      container.innerHTML = '<span class="pill-row-empty">No participants</span>';
      return;
    }

    var taskByPid = {};
    state.tasks.forEach(function (t) {
      if (!taskByPid[t.participant] || t.created_at > taskByPid[t.participant].created_at) {
        taskByPid[t.participant] = t;
      }
    });

    // In-place patch when structure unchanged (avoids layout thrash + preserves
    // the open options pane across polling ticks)
    var existing = container.querySelectorAll(".pill-wrap[data-pid]");
    if (existing.length === state.participants.length) {
      var canPatch = true;
      for (var k = 0; k < state.participants.length; k++) {
        var p0 = state.participants[k];
        var s0 = pillState(p0, taskByPid);
        var agentsAttr = s0.agents.transcription + "," + s0.agents.summary + "," + s0.agents.citations;
        if (existing[k].getAttribute("data-pid") !== p0.id ||
            existing[k].getAttribute("data-status") !== s0.status ||
            existing[k].getAttribute("data-active") !== (state.selectedParticipant === p0.id ? "1" : "0") ||
            existing[k].getAttribute("data-agents") !== agentsAttr) {
          canPatch = false; break;
        }
      }
      if (canPatch) {
        for (k = 0; k < state.participants.length; k++) {
          var wrap = existing[k];
          p0 = state.participants[k];
          s0 = pillState(p0, taskByPid);
          var bar = wrap.querySelector(".pill-progress");
          if (bar) bar.style.width = s0.progress + "%";
        }
        return;
      }
    }

    // Full rebuild
    var openPid = state.pillOptionsOpen;
    var frag = document.createDocumentFragment();
    state.participants.forEach(function (p) {
      frag.appendChild(buildPillWrap(p, taskByPid));
    });
    container.innerHTML = "";
    container.appendChild(frag);
    // The pane is mounted on <body>, not inside the rebuilt wrap — reposition
    // it to the new wrap's rect and re-render its contents so the agent rows
    // reflect the latest state. Tear it down if the pill disappeared.
    if (openPid !== null) {
      var newWrap = _findPillWrap(openPid);
      var floating = document.querySelector("body > .pill-options");
      if (newWrap && floating) {
        _refreshPillOptionsContent(openPid, taskByPid);
        var refreshed = document.querySelector("body > .pill-options[data-pid='" + openPid + "']");
        if (refreshed) _positionPillOptions(refreshed, newWrap);
      } else {
        closePillOptions();
      }
    }
  }

  function _refreshPillOptionsContent(pid, taskByPid) {
    var floating = document.querySelector("body > .pill-options[data-pid='" + pid + "']");
    if (!floating) return;
    var p = null;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) { p = state.participants[i]; break; }
    }
    if (!p) return;
    var s = pillState(p, taskByPid);
    var fresh = buildPillOptions(p, s);
    fresh.setAttribute("data-pid", pid);
    floating.parentNode.replaceChild(fresh, floating);
  }

  function buildPillWrap(p, taskByPid) {
    var s = pillState(p, taskByPid);
    var wrap = document.createElement("div");
    var isActive = state.selectedParticipant === p.id;
    wrap.className = "pill-wrap";
    wrap.setAttribute("data-pid", p.id);
    wrap.setAttribute("data-status", s.status);
    wrap.setAttribute("data-active", isActive ? "1" : "0");
    wrap.setAttribute("data-agents", s.agents.transcription + "," + s.agents.summary + "," + s.agents.citations);
    wrap.appendChild(buildPill(p, s, isActive));
    wrap.appendChild(buildPillDots(p, s));
    if (state.pillOptionsOpen === p.id) {
      wrap.classList.add("pill-wrap--options-open");
    }
    wrap.addEventListener("mouseenter", function () {
      maybeWarmOnPillHover(p, s);
    });
    return wrap;
  }

  function _dotStateLabel(st) {
    if (st === "done") return "done";
    if (st === "running") return "in progress\u2026";
    if (st === "failed") return "failed";
    return "not started";
  }

  function buildPillDots(p, s) {
    var ag = s.agents;
    var labels = ["Transcription", "Summary", "Citations"];
    var keys = ["transcription", "summary", "citations"];
    var anyActive = false;
    for (var i = 0; i < keys.length; i++) {
      if (ag[keys[i]] !== "idle") { anyActive = true; break; }
    }
    if (!anyActive) {
      var empty = document.createElement("div");
      empty.className = "pill-dots pill-dots--empty";
      return empty;
    }
    var row = document.createElement("div");
    row.className = "pill-dots";
    for (var j = 0; j < keys.length; j++) {
      var dot = document.createElement("span");
      dot.className = "pill-dot pill-dot--" + ag[keys[j]];
      row.appendChild(dot);
    }
    attachHoverTooltip(row, function () {
      var lines = [];
      for (var k = 0; k < keys.length; k++) {
        lines.push(labels[k] + ": " + _dotStateLabel(ag[keys[k]]));
      }
      return lines.join("\n");
    }, { multiline: true, align: "center" });
    return row;
  }

  function buildPill(p, s, isActive) {
    var pill = document.createElement("div");
    var classes = ["pill", "pill--" + s.status];
    if (isActive) classes.push("pill--active");
    pill.className = classes.join(" ");
    pill.setAttribute("role", "tab");
    pill.setAttribute("aria-selected", isActive ? "true" : "false");

    // Progress fill (background layer)
    var prog = document.createElement("div");
    prog.className = "pill-progress";
    prog.style.width = s.progress + "%";
    pill.appendChild(prog);

    // Trigger — status icon doubles as action button (hover swaps the glyph)
    pill.appendChild(buildPillTrigger(p, s));

    var idSpan = document.createElement("span");
    idSpan.className = "pill-id";
    idSpan.textContent = p.id;
    pill.appendChild(idSpan);

    // Stale badge (inline)
    if (p.has_stale_artifacts) {
      var stale = document.createElement("span");
      stale.className = "pill-stale-badge";
      stale.textContent = "stale";
      stale.title = "Artifacts built from an older transcript";
      pill.appendChild(stale);
    }

    // Chevron for options pane
    var chevBtn = document.createElement("button");
    chevBtn.type = "button";
    chevBtn.className = "pill-chevron-btn";
    chevBtn.setAttribute("aria-label", "Transcription options");
    chevBtn.setAttribute("aria-expanded", state.pillOptionsOpen === p.id ? "true" : "false");
    var chev = document.createElement("span");
    chev.className = "pill-chevron";
    chevBtn.appendChild(chev);
    chevBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      togglePillOptions(p.id);
    });
    pill.appendChild(chevBtn);

    // Clicking the pill body (not its buttons) selects the participant
    pill.addEventListener("click", function () {
      if (p.id !== state.selectedParticipant) selectParticipant(p.id);
    });

    return pill;
  }

  function buildPillTrigger(p, s) {
    var cfg = PILL_TRIGGER[s.status] || PILL_TRIGGER.idle;
    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "pill-trigger pill-trigger--" + s.status;
    btn.setAttribute("aria-label", cfg.label);
    btn.setAttribute("title", cfg.label);

    // Two stacked icons; CSS hides one and shows the other on hover/focus.
    var rest = document.createElement("span");
    rest.className = "pill-trigger-icon pill-trigger-icon--rest";
    applyMaskIcon(rest, "url(" + cfg.rest + ")");
    btn.appendChild(rest);

    var hover = document.createElement("span");
    hover.className = "pill-trigger-icon pill-trigger-icon--hover";
    applyMaskIcon(hover, "url(" + cfg.hover + ")");
    btn.appendChild(hover);

    btn.addEventListener("click", function (e) {
      e.stopPropagation();
      if (cfg.action === "cancel") {
        if (s.taskId) {
          apiDelete("api/transcribe/" + s.taskId).then(function () { pollTaskStatus(); });
        }
      } else if (cfg.action === "retranscribe") {
        startTranscribe(p.id, true);
      } else {
        startTranscribe(p.id, false);
      }
    });
    return btn;
  }

  // ---- Participant pill popover (model/lang + agent dependency rows) ----

  function buildPillOptions(p, s) {
    var pane = document.createElement("div");
    pane.className = "pill-options";
    pane.addEventListener("click", function (e) { e.stopPropagation(); });

    var ov = state.pillOverrides[p.id] || {};

    // Model row
    var modelRow = document.createElement("div");
    modelRow.className = "pill-options-row";
    var modelLabel = document.createElement("label");
    modelLabel.textContent = "Model";
    var modelSelect = document.createElement("select");
    modelSelect.innerHTML = '<option value="">Loading…</option>';
    _trFetchModels().then(function (data) {
      var models = (data && data.whisper && data.whisper.models) || [];
      if (models.length === 0) {
        modelSelect.innerHTML = '<option value="">Default</option>';
        return;
      }
      var defaultName = "";
      var opts = '<option value="">Default</option>';
      models.forEach(function (m) {
        if (m.selected) defaultName = m.name;
        opts += '<option value="' + escapeHtml(m.name) + '">' + escapeHtml(m.name) + '</option>';
      });
      modelSelect.innerHTML = opts;
      modelSelect.options[0].textContent = "Default (" + (defaultName || "base") + ")";
      modelSelect.value = ov.model || "";
    });
    modelSelect.addEventListener("change", function () {
      _setOverride(p.id, "model", this.value);
    });
    modelRow.appendChild(modelLabel);
    modelRow.appendChild(modelSelect);
    pane.appendChild(modelRow);

    // Language row
    var langRow = document.createElement("div");
    langRow.className = "pill-options-row";
    var langLabel = document.createElement("label");
    langLabel.textContent = "Language";
    var langSelect = document.createElement("select");
    var lh = "";
    for (var i = 0; i < PILL_LANGUAGES.length; i++) {
      var L = PILL_LANGUAGES[i];
      lh += '<option value="' + escapeHtml(L.code) + '">' + escapeHtml(L.label) + '</option>';
    }
    langSelect.innerHTML = lh;
    langSelect.value = ov.language || "";
    langSelect.addEventListener("change", function () {
      _setOverride(p.id, "language", this.value);
    });
    langRow.appendChild(langLabel);
    langRow.appendChild(langSelect);
    pane.appendChild(langRow);

    // Agent rows — manual run / re-run / stop controls with dependency gating.
    // Order: Transcription → Summary → Citations. Summary requires
    // transcription; citations requires summary. Re-running summary cascades
    // to citations server-side (see transcripts_server.py).
    pane.appendChild(buildPillAgentsSection(p, s));

    return pane;
  }

  function buildPillAgentsSection(p, s) {
    var section = document.createElement("div");
    section.className = "pill-options-agents";

    // 1. Transcription
    section.appendChild(buildAgentRow({
      pid: p.id,
      label: "Transcription",
      agent: "transcription",
      depMet: true,
      agentState: s.agents.transcription,
      hasResult: !!p.has_transcript,
      cascadeWarning: !!(p.agents && (p.agents.summary === "done" || p.agents.citations === "done")),
      onStart: function () { startTranscribe(p.id, !!p.has_transcript); },
      onStop: function () {
        if (s.taskId) {
          apiDelete("api/transcribe/" + s.taskId).then(function () { pollTaskStatus(); });
        }
      },
    }));

    // 2. Summary
    section.appendChild(buildAgentRow({
      pid: p.id,
      label: "Summary",
      agent: "summary",
      depLabel: "transcription",
      depMet: s.agents.transcription === "done",
      agentState: s.agents.summary,
      hasResult: !!(p.agents && p.agents.summary === "done"),
      cascadeWarning: !!(p.agents && p.agents.citations === "done"),
      onStart: function () {
        apiPost("api/summary/" + p.id + "/regenerate", {}).then(function () {
          _refreshAgentStateNow();
        }).catch(function () {
          showToast("Failed to start summary");
        });
      },
      onStop: function () {
        apiPost("api/summary/" + p.id + "/stop", {}).then(function () {
          _refreshAgentStateNow();
        }).catch(function () {
          showToast("Failed to stop summary");
        });
      },
    }));

    // 3. Citations
    section.appendChild(buildAgentRow({
      pid: p.id,
      label: "Citations",
      agent: "citations",
      depLabel: "summary",
      depMet: s.agents.summary === "done",
      agentState: s.agents.citations,
      hasResult: !!(p.agents && p.agents.citations === "done"),
      cascadeWarning: false,
      onStart: function () {
        apiPost("api/citations/" + p.id + "/regenerate", {}).then(function () {
          _refreshAgentStateNow();
        }).catch(function () {
          showToast("Failed to start citations");
        });
      },
      onStop: function () {
        apiPost("api/citations/" + p.id + "/stop", {}).then(function () {
          _refreshAgentStateNow();
        }).catch(function () {
          showToast("Failed to stop citations");
        });
      },
    }));

    return section;
  }

  function buildAgentRow(opts) {
    var row = document.createElement("div");
    row.className = "pill-options-row pill-options-agent-row";
    row.setAttribute("data-agent", opts.agent);

    var label = document.createElement("label");
    label.textContent = opts.label;
    row.appendChild(label);

    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn btn-small pill-agent-btn";

    var running = opts.agentState === "running";
    var mode = "start"; // start | stop | disabled
    var title = "";

    if (running) {
      btn.textContent = "Stop";
      btn.classList.add("pill-agent-btn--stop");
      mode = "stop";
    } else if (!opts.depMet) {
      btn.textContent = "Run";
      mode = "disabled";
      title = "Requires " + opts.depLabel + " to finish first";
    } else if (opts.hasResult) {
      btn.textContent = "Re-run";
      if (opts.cascadeWarning) {
        title = opts.agent === "transcription"
          ? "Re-transcribing invalidates Summary and Citations"
          : "Re-running will also re-run Citations";
      }
    } else {
      btn.textContent = "Run";
    }

    if (mode === "disabled") btn.setAttribute("disabled", "disabled");
    if (title) btn.title = title;

    btn.addEventListener("click", function (e) {
      e.stopPropagation();
      if (btn.hasAttribute("disabled")) return;
      // Optimistic UI swap; the next poll re-renders via _refreshPillOptionsContent.
      if (mode === "stop") {
        btn.textContent = "Stopping\u2026";
        btn.setAttribute("disabled", "disabled");
        opts.onStop();
      } else {
        btn.textContent = "Starting\u2026";
        btn.setAttribute("disabled", "disabled");
        opts.onStart();
      }
    });

    row.appendChild(btn);
    return row;
  }

  function _setOverride(pid, key, value) {
    if (!state.pillOverrides[pid]) state.pillOverrides[pid] = {};
    if (value) state.pillOverrides[pid][key] = value;
    else delete state.pillOverrides[pid][key];
  }

  function _findPillWrap(pid) {
    var container = qs("#participantPills");
    if (!container) return null;
    var wraps = container.querySelectorAll(".pill-wrap[data-pid]");
    for (var i = 0; i < wraps.length; i++) {
      if (wraps[i].getAttribute("data-pid") === pid) return wraps[i];
    }
    return null;
  }

  function _positionPillOptions(pane, wrap) {
    positionPopoverAnchored(pane, wrap.getBoundingClientRect());
  }

  function _closeOpenPaneInDom() {
    if (state.pillOptionsOpen === null) return;
    var wrap = _findPillWrap(state.pillOptionsOpen);
    if (wrap) {
      wrap.classList.remove("pill-wrap--options-open");
      var chev = wrap.querySelector(".pill-chevron-btn");
      if (chev) chev.setAttribute("aria-expanded", "false");
    }
    var floating = document.querySelector("body > .pill-options");
    if (floating) floating.remove();
    if (state.pillOptionsReposition) {
      window.removeEventListener("resize", state.pillOptionsReposition);
      window.removeEventListener("scroll", state.pillOptionsReposition, true);
      state.pillOptionsReposition = null;
    }
    state.pillOptionsOpen = null;
  }

  function togglePillOptions(pid) {
    // Close whichever pane is currently open (possibly same pill → toggle off)
    var wasOpen = state.pillOptionsOpen === pid;
    _closeOpenPaneInDom();
    if (wasOpen) return;

    var wrap = _findPillWrap(pid);
    if (!wrap) return;
    var p = null;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) { p = state.participants[i]; break; }
    }
    if (!p) return;
    var taskByPid = {};
    state.tasks.forEach(function (t) {
      if (!taskByPid[t.participant] || t.created_at > taskByPid[t.participant].created_at) {
        taskByPid[t.participant] = t;
      }
    });
    var s = pillState(p, taskByPid);

    // Mount the pane on <body> (not inside #participantPills) so it escapes the
    // pill row's overflow clipping and renders above the search bar.
    var pane = buildPillOptions(p, s);
    pane.setAttribute("data-pid", pid);
    document.body.appendChild(pane);
    _positionPillOptions(pane, wrap);

    wrap.classList.add("pill-wrap--options-open");
    var chev = wrap.querySelector(".pill-chevron-btn");
    if (chev) chev.setAttribute("aria-expanded", "true");
    state.pillOptionsOpen = pid;

    var reposition = function () {
      var w = _findPillWrap(pid);
      var floating = document.querySelector("body > .pill-options[data-pid='" + pid + "']");
      if (w && floating) _positionPillOptions(floating, w);
    };
    state.pillOptionsReposition = reposition;
    window.addEventListener("resize", reposition);
    window.addEventListener("scroll", reposition, true);

    if (state.transcribePrewarm === "queue_open") {
      tryPostTranscriptionWarmup();
    }
  }

  function closePillOptions() {
    _closeOpenPaneInDom();
  }

  function initPillOutsideClick() {
    document.addEventListener("click", function (e) {
      if (state.pillOptionsOpen === null) return;
      var wrap = _findPillWrap(state.pillOptionsOpen);
      if (wrap && wrap.contains(e.target)) return;
      var floating = document.querySelector("body > .pill-options");
      if (floating && floating.contains(e.target)) return;
      closePillOptions();
    });
    window.addEventListener("pagehide", closePillOptions);
  }

  function initPillWheelScroll() {
    var el = qs("#participantPills");
    if (!el) return;
    el.addEventListener("wheel", function (e) {
      if (el.scrollWidth > el.clientWidth) {
        e.preventDefault();
        el.scrollLeft += e.deltaY;
      }
    }, { passive: false });
  }

  function startTranscribe(pid, force) {
    var overrides = {};
    var ov = state.pillOverrides[pid];
    if (ov && (ov.model || ov.language)) {
      overrides[pid] = {};
      if (ov.model) overrides[pid].model = ov.model;
      if (ov.language) overrides[pid].language = ov.language;
    }
    transcribeParticipants([pid], force, overrides);
  }

  function transcribeParticipants(pids, force, overrides) {
    var body = { participants: pids, force: force };
    if (overrides && Object.keys(overrides).length > 0) body.overrides = overrides;
    apiPost("api/transcribe", body).then(function (data) {
      if (!data.ok) {
        showToast("Failed to enqueue transcription");
        return;
      }
      showToast("Enqueued " + clipgenPluralUnit(data.tasks.length, "transcription", "transcriptions"));
      startPolling();
      pollTaskStatus();
    });
  }

  // ---- Task polling ----
  // state.pollTimer → pollTaskStatus for Whisper jobs; summary/citations use
  // separate timers (see file header).

  function startPolling() {
    if (state.pollTimer) return;
    state.pollTimer = setInterval(pollTaskStatus, POLL_INTERVAL);
  }

  function stopPolling() {
    if (state.pollTimer) {
      clearInterval(state.pollTimer);
      state.pollTimer = null;
    }
  }

  var _refreshedCompletedPids = {};

  function _anyAgentActive() {
    for (var i = 0; i < state.participants.length; i++) {
      var ag = state.participants[i].agents;
      if (ag && (ag.summary === "running" || ag.citations === "running")) return true;
    }
    return false;
  }

  // Called from pill agent onStart/onStop after the API POST resolves.
  // Reloads participants so pill state reflects the new running/idle state
  // immediately, and starts the poll loop so the pill keeps updating without
  // waiting for the next external trigger (poll loop is otherwise gated on
  // `_anyAgentActive()` which reads stale state right after a manual run).
  function _refreshAgentStateNow() {
    loadParticipants().then(function () {
      if (_anyAgentActive()) startPolling();
    });
  }

  function pollTaskStatus() {
    apiGet("api/transcribe/status").then(function (data) {
      if (!data.ok) return;
      state.tasks = data.tasks;

      // Re-render the status circle immediately so completed tasks reflect
      // before the async loadParticipants()/loadTranscript() chain resolves.
      // This is what keeps the indicator from freezing at "95%".
      updateStatusIndicator();

      // Stream partial segments for the selected participant's running task
      var selectedRunningTask = null;
      if (state.selectedParticipant) {
        data.tasks.forEach(function (t) {
          if (t.participant === state.selectedParticipant && t.status === "running" && t.partial_segments) {
            selectedRunningTask = t;
          }
        });
      }
      if (selectedRunningTask && selectedRunningTask.partial_segments.length > 0) {
        renderPartialSegments(selectedRunningTask.partial_segments, selectedRunningTask.progress);
        state.streamingParticipant = state.selectedParticipant;
        _loadStreamingMarks(state.selectedParticipant);
      } else if (state.streamingParticipant) {
        state.streamingParticipant = null;
      }

      var hasActive = false;
      var newlyCompleted = [];
      data.tasks.forEach(function (t) {
        if (t.status === "queued" || t.status === "running") hasActive = true;
        if (t.status === "completed" && !_refreshedCompletedPids[t.participant]) {
          newlyCompleted.push(t.participant);
          _refreshedCompletedPids[t.participant] = true;
        }
      });

      // Refresh participants and transcript as each task completes.
      // Thinking-agents (summary → citations) are spawned on whisper completion
      // and on server startup, so we always refresh after anything completes
      // or if any agent is currently running on any pill.
      var needsRefresh = newlyCompleted.length > 0 || _anyAgentActive();
      if (newlyCompleted.length > 0) {
        _streamingMarks = {};
        _streamingMarksLoaded = false;
        _bumpStreamingMarksVersion();
      }
      if (needsRefresh) {
        loadParticipants().then(function () {
          if (newlyCompleted.length > 0 && state.selectedParticipant &&
              newlyCompleted.indexOf(state.selectedParticipant) >= 0) {
            state.streamingParticipant = null;
            loadTranscript(state.selectedParticipant);
            loadSummary(state.selectedParticipant);
          }
          updateStatusIndicator();
          // Agents typically kick in right after whisper completes; keep the
          // poll alive so dot transitions (running → done → next) are seen.
          if (_anyAgentActive()) startPolling();
          else if (!hasActive) {
            stopPolling();
            _refreshedCompletedPids = {};
          }
        });
      }

      if (hasActive || _anyAgentActive()) {
        startPolling();
      } else if (!needsRefresh) {
        stopPolling();
        _refreshedCompletedPids = {};
      }

      if (!hasActive && _hadActiveTranscriptionLastPoll) {
        refreshTranscriptionModelHintOnce();
      }
      _hadActiveTranscriptionLastPoll = hasActive;

      renderPills();
    });
  }

  // ---- Corrections modal ----

  function initCorrectionsModal() {
    qs("#correctionsBtn").addEventListener("click", function () {
      qs("#correctionsModal").classList.remove("hidden");
      loadCorrections();
    });

    qs("#closeCorrectionsBtn").addEventListener("click", function () {
      qs("#correctionsModal").classList.add("hidden");
    });

    qs("#correctionsModal").addEventListener("click", function (e) {
      if (e.target === this) this.classList.add("hidden");
    });

    qs("#addCorrectionBtn").addEventListener("click", function () {
      addCorrection();
    });

    // Enter key in correction form
    qs("#correctionTo").addEventListener("keydown", function (e) {
      if (e.key === "Enter") {
        e.preventDefault();
        addCorrection();
      }
    });
  }

  function loadCorrections() {
    apiGet("api/corrections").then(function (data) {
      if (!data.ok) return;
      state.corrections = data.corrections;
      renderCorrections();
    });
  }

  function renderCorrections() {
    var container = qs("#correctionsList");
    if (state.corrections.length === 0) {
      container.innerHTML = '<div style="color:var(--color-text-dim);font-size:var(--text-sm);padding:var(--space-2) 0">No corrections yet</div>';
      return;
    }

    var html = "";
    state.corrections.forEach(function (c) {
      html += '<div class="correction-row">';
      html += '<span class="correction-from">' + escapeHtml(c.from) + '</span>';
      html += '<span class="correction-arrow">&rarr;</span>';
      html += '<span class="correction-to">' + escapeHtml(c.to) + '</span>';
      html += '<button class="correction-delete" data-id="' + escapeHtml(c.id) + '">Remove</button>';
      html += '</div>';
    });
    container.innerHTML = html;

    // Attach delete handlers
    var btns = container.querySelectorAll(".correction-delete");
    for (var i = 0; i < btns.length; i++) {
      btns[i].addEventListener("click", function () {
        deleteCorrection(this.getAttribute("data-id"));
      });
    }
  }

  function addCorrection() {
    var fromInput = qs("#correctionFrom");
    var toInput = qs("#correctionTo");
    var fromText = fromInput.value.trim();
    var toText = toInput.value.trim();
    if (!fromText || !toText) return;

    apiPost("api/corrections", { from: fromText, to: toText }).then(function (data) {
      if (data.ok) {
        fromInput.value = "";
        toInput.value = "";
        showToast("Correction added");
        loadCorrections();
        // Reload transcript to apply new correction
        if (state.selectedParticipant) loadTranscript(state.selectedParticipant);
      }
    });
  }

  function deleteCorrection(id) {
    apiDelete("api/corrections/" + id).then(function (data) {
      if (data.ok) {
        showToast("Correction removed");
        loadCorrections();
        // Reload transcript to unapply correction
        if (state.selectedParticipant) loadTranscript(state.selectedParticipant);
      }
    });
  }

  function initTooltipToggle() {
    state.tooltipsEnabled = getStoredTooltipPref();
    var btn = qs("#tooltipToggle");
    if (!btn) return;
    btn.setAttribute("aria-pressed", state.tooltipsEnabled ? "true" : "false");
    btn.addEventListener("click", function () {
      state.tooltipsEnabled = !state.tooltipsEnabled;
      btn.setAttribute("aria-pressed", state.tooltipsEnabled ? "true" : "false");
      setStoredTooltipPref(state.tooltipsEnabled);
      if (state.searchResults) renderSearchResults(state.searchResults);
      if (state.segments.length > 0) renderSegments();
    });
  }

  // ---- Settings (shared modal lives in settings-modal.js) ----

  // Models are also fetched for per-pill model overrides; keep a tiny
  // cached fetcher here. The shared modal maintains its own cache.
  var _trModelsCache = null;
  var _trModelsCachePromise = null;

  function _trFetchModels() {
    if (_trModelsCache) return Promise.resolve(_trModelsCache);
    if (_trModelsCachePromise) return _trModelsCachePromise;
    _trModelsCachePromise = fetch("/api/models")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data && data.ok) _trModelsCache = data;
        return data;
      })
      .catch(function () { return null; });
    return _trModelsCachePromise;
  }

  function _applySettingsSnapshot(applied, settings) {
    var nextCats = null;
    if (applied && applied.MARK_CATEGORIES) {
      nextCats = applied.MARK_CATEGORIES;
    } else if (settings) {
      for (var i = 0; i < settings.length; i++) {
        if (settings[i].name === "MARK_CATEGORIES") {
          nextCats = settings[i].value;
          break;
        }
      }
    }
    if (nextCats) {
      setMarkCategories(nextCats);
      if (!MARK_CATEGORIES[state.lastMarkCategory]) {
        var firstKey = Object.keys(MARK_CATEGORIES)[0];
        state.lastMarkCategory = firstKey || "bookmark";
      }
      // Refresh streaming mark colors and re-render the visible transcript.
      for (var sid in _streamingMarks) {
        var sm = _streamingMarks[sid];
        var cat = MARK_CATEGORIES[sm.category];
        if (cat) sm.color = cat.color;
      }
      _bumpStreamingMarksVersion();
      hideMarkPopover();
      if (state.selectedParticipant) loadTranscript(state.selectedParticipant);
    }
  }

  function initTranscriptSettings() {
    var btn = qs("#settingsBtn");
    if (!btn) return;
    btn.addEventListener("click", function () {
      openSettingsModal({
        initialTab: "Transcription",
        onSave: function (applied, settings) {
          _trModelsCache = null;
          _trModelsCachePromise = null;
          _applySettingsSnapshot(applied, settings);
        },
        onReset: function (scope, settings) {
          _trModelsCache = null;
          _trModelsCachePromise = null;
          _applySettingsSnapshot(null, settings);
        },
      });
    });
  }

  // ---- Boot ----

  function runEmbedSubtitle() {
    var pid = state.selectedParticipant;
    if (!pid) return;
    showToast("Embedding subtitle…");
    apiPost("api/embed-subtitle/" + pid, {})
      .then(function (data) {
        if (data && data.ok) showToast("Wrote " + data.output_filename);
        else showToast((data && data.error) || "Embed failed");
      })
      .catch(function (err) { showToast("Embed failed: " + err.message); });
  }

  function runEmbedAllSubtitles() {
    showToast("Embedding subtitles for all transcripts…");
    apiPost("api/embed-all-subtitles", {})
      .then(function (data) {
        if (!data || !data.ok) {
          showToast((data && data.error) || "Embed failed");
          return;
        }
        var results = data.results || [];
        var okCount = 0;
        for (var i = 0; i < results.length; i++) if (results[i].ok) okCount++;
        showToast("Embedded " + okCount + "/" + clipgenPluralUnit(results.length, "video", "videos") + " to " + data.output_dir);
      })
      .catch(function (err) { showToast("Embed failed: " + err.message); });
  }

  var _rebuildTopNavActions = function () {};

  function _currentParticipantHasTranscript() {
    var pid = state.selectedParticipant;
    if (!pid || !state.participants) return false;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) return !!state.participants[i].has_transcript;
    }
    return false;
  }

  function _anyTranscriptExists() {
    var ps = state.participants || [];
    for (var i = 0; i < ps.length; i++) {
      if (ps[i].has_transcript) return true;
    }
    return false;
  }

  function refreshTopNavActions() {
    _rebuildTopNavActions();
  }

  function initTopNavActions() {
    if (!window.ClipgenTopNav) return;
    function rebuild() {
      var hasOne = _currentParticipantHasTranscript();
      var hasAny = _anyTranscriptExists();
      window.ClipgenTopNav.setQuickActions([
        {
          icon: "language",
          label: "Embed Subtitle in Video",
          action: runEmbedSubtitle,
          disabled: !hasOne,
          title: hasOne
            ? "Mux this participant's transcript into a copy of their source video"
            : "Select a participant with a transcript to enable this.",
        },
        {
          icon: "film",
          label: "Embed all Subtitles",
          action: runEmbedAllSubtitles,
          disabled: !hasAny,
          title: hasAny
            ? "Mux every participant's transcript into a subtitled copy of their video"
            : "Transcribe at least one video to enable this.",
        },
        window.ClipgenExportActions.exportQuickAction(),
      ]);
    }
    _rebuildTopNavActions = rebuild;
    rebuild();
    window.ClipgenExportActions.refreshExportStatus(rebuild);
    // Always rebuild on menu open so Embed-Subtitle items pick up
    // participant-selection changes; also refresh export-enabled state.
    window.ClipgenTopNav.onBeforeOpen(function () {
      rebuild();
      window.ClipgenExportActions.refreshExportStatus(rebuild);
    });
  }

  document.addEventListener("DOMContentLoaded", function () {
    initThemeToggle();
    initTooltipToggle();
    initStatusIndicatorTooltip();
    checkNavLinks();
    initFrontendSwitcher();
    initSearch();
    initPillOutsideClick();
    initPillWheelScroll();
    initCorrectionsModal();
    initVideoPlayer();
    initVideoSync();
    initTimelineCanvas();
    initPipScroll();
    initPlayerKeyboard();
    initSummaryToggle();
    initSummaryActions();
    initTranscriptSettings();
    initTopNavActions();

    // Pause every poller when tab is hidden; resume what was active on focus.
    // Without this, summary/citations/model-hint pollers (1.5–3 s cadence)
    // keep hammering the backend from background tabs.
    document.addEventListener("visibilitychange", function () {
      if (document.hidden) {
        stopPolling();
        stopXrefPolling();
        stopModelHintPoll();
        _stopSummaryPoll();
        _stopCitationsPoll();
      } else {
        pollTaskStatus();
        startXrefPolling();
        // Re-check summary + citations on tab refocus. Background-running
        // Ollama agents finish without notifying the frontend; if the citations
        // poll already gave up (or summary completed after we stopped polling)
        // the manifest result would only surface on a full page reload. This
        // catches the common "user goes to another tab, comes back" case.
        // loadSummary re-arms the summary/citations polls if generation is
        // still in flight, and the transcribe-status poll re-arms the
        // model-hint poll on the next active task transition.
        if (state.selectedParticipant) loadSummary(state.selectedParticipant);
      }
    });

    // Load initial data
    loadParticipants();

    // Fetch live mark categories so the popover/pill renders match overrides.
    apiGet("api/marks").then(function (data) {
      if (data && data.ok && data.categories) {
        setMarkCategories(data.categories);
        if (!MARK_CATEGORIES[state.lastMarkCategory]) {
          var firstKey = Object.keys(MARK_CATEGORIES)[0];
          state.lastMarkCategory = firstKey || "bookmark";
        }
      }
    });

    // Check for active tasks on load
    pollTaskStatus();
  });

})();
