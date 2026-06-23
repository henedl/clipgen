/* clipgen Transcripts page.
 *
 * Editor for per-participant Whisper transcripts, plus cross-references into
 * Screenspace events and the source spreadsheet. Two long-running side flows
 * piggyback on the same `state`:
 *
 *   - Transcription warmup: a single `tryPostTranscriptionWarmup()` post that
 *     asks the backend to preload the Whisper model. `_transcriptionWarmupPosted`
 *     guards it so we never double-post per page load.
 *   - Summary / citations: Ollama-generated; `_summaryPoller` and
 *     `_citationsPoller` (createPoller handles) poll the backend until the
 *     result lands or the user navigates away.
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
    pollPoller: null,
    lastMarkCategory: "bookmark",
    streamingParticipant: null,
    ssEvents: [],
    ssEventsLoaded: false,
    sheetRows: [],
    sheetParticipants: [],
    sheetLoaded: false,
    xrefPoller: null,
    xrefEligible: false,
    xrefIndex: { eventsByParticipant: {}, sheetByParticipant: {} },
    tooltipsEnabled: true,
    summaryEditing: false,
    summaryText: "",
    summaryCitations: null,
    citationsGenerating: false,
    activeTab: "summary",
    frictionData: null,
    frictionBySegId: {},
    frictionGenerating: false,
    // Server-recorded friction run start (epoch ms) so the elapsed clock
    // survives page navigation; null while idle or for a just-clicked run.
    frictionStartedAt: null,
    frictionHeatmapEnabled: false,
    frictionThreshold: 0.5,
    frictionCategoryFilter: null,
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
  // Prewarm never downloads silently: when the model isn't cached we confirm
  // the download with the user. _prewarmDownloadPrompting guards against
  // double-prompting; _prewarmDeclinedModel records the specific model the user
  // declined so we stop re-asking for it — but switching to a different model
  // (or changing TRANSCRIBE_MODEL in settings) still gets its own prompt. The
  // model still loads on demand, with confirmation, at transcribe time.
  var _prewarmDownloadPrompting = false;
  var _prewarmDeclinedModel = null;
  // Last-known TRANSCRIBE_MODEL, so a settings change can reset the prewarm
  // guards for the new model. Seeded once from the model-status hint.
  var _lastTranscribeModel = null;
  var _modelHintPoller = null;
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
    if (!state.xrefEligible || state.xrefPoller) return;
    // createPoller runs loadCrossRefData once immediately (runImmediately default),
    // then every 30s — matching the old explicit call + setInterval.
    state.xrefPoller = createPoller(loadCrossRefData, 30000);
    state.xrefPoller.start();
  }

  function stopXrefPolling() {
    if (state.xrefPoller) {
      state.xrefPoller.stop();
      state.xrefPoller = null;
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
    // Seed the change-detector once; thereafter only settings saves update it,
    // so a poll landing mid-save can't mask a model change.
    if (data.model && _lastTranscribeModel === null) _lastTranscribeModel = data.model;
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
      taskLine = pid + ": transcribing\u2026 " + pct + "%" + _txEtaSuffix(pid, task);
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
    if (_modelHintPoller) {
      _modelHintPoller.stop();
      _modelHintPoller = null;
    }
  }

  // Forget which Whisper downloads the user has agreed to and refresh the
  // models cache. Call when a download attempt has concluded (transcription
  // finished, or a warmup load loaded/failed) so the next gate re-reads real
  // cache state — re-prompting for a model whose download failed, staying
  // quiet for one that succeeded. Never call mid-download (it would re-prompt).
  function _forgetWhisperDownloadAgreements() {
    _whisperDownloadConfirmed = {};
    _trModelsCache = null;
    _trModelsCachePromise = null;
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
            _forgetWhisperDownloadAgreements();
            return;
          }
          if (!data.warming) {
            // Warmup ended without loading (failed/idle) — let the next attempt
            // re-confirm rather than silently re-downloading.
            stopModelHintPoll();
            _forgetWhisperDownloadAgreements();
          }
        })
        .catch(function () {});
    };
    // createPoller runs poll() once immediately (runImmediately default), then
    // every 1.5s — matching the old explicit poll() + setInterval.
    _modelHintPoller = createPoller(poll, 1500);
    _modelHintPoller.start();
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
          if (data.reason === "model_not_cached") {
            _confirmPrewarmDownload(data);
            return;
          }
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

  // Prewarm wanted to load a model that isn't downloaded yet. We never
  // download silently — confirm with the user first (once per session). On
  // confirm, re-post warmup with force=true so the backend proceeds; on
  // decline, leave the warmup flag set so we stop re-posting/re-prompting
  // (the model still loads, with confirmation, at transcribe time).
  function _confirmPrewarmDownload(data) {
    if (_prewarmDownloadPrompting) return;
    if (_prewarmDeclinedModel === data.model) {
      _transcriptionWarmupPosted = true;
      refreshTranscriptionModelHintOnce();
      return;
    }
    _prewarmDownloadPrompting = true;
    confirmModelInstall({
      kind: "whisper",
      prewarm: true,
      model: data.model,
      sizeMb: data.size_mb,
    }).then(function (ok) {
      _prewarmDownloadPrompting = false;
      if (!ok) {
        _prewarmDeclinedModel = data.model;
        _transcriptionWarmupPosted = true;
        refreshTranscriptionModelHintOnce();
        return;
      }
      // Agreed once — don't re-prompt at transcribe time before the download
      // completes (the models cache still reports it as not cached).
      if (data.model) _whisperDownloadConfirmed[data.model] = true;
      apiPost("api/transcribe/warmup", { force: true })
        .then(function (d) {
          if (!d.ok) {
            _transcriptionWarmupPosted = false;
            return;
          }
          if (d.started || d.already_warming) {
            startModelHintPoll();
            return;
          }
          if (d.already_loaded) {
            refreshTranscriptionModelHintOnce();
            return;
          }
          _transcriptionWarmupPosted = false;
        })
        .catch(function () {
          _transcriptionWarmupPosted = false;
        });
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
    _stopFrictionPoll();
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

    // Set video source. ?v=<mtime_ns> mirrors the screenspace cache-bust so
    // a re-encoded or replaced source file invalidates the browser HTTP cache
    // instead of relying on send_from_directory's Last-Modified revalidation.
    if (p.has_video) {
      // Multi-video participants carry a per-part timeline; play part-by-part
      // with client-side source switching (see the timeline helpers).
      state.videoTimeline = p.timeline && p.timeline.length > 1 ? p.timeline : null;
      state.videoVersion = p.video_version != null ? p.video_version : null;
      state.videoActivePart = 0;
      state.videoOffset = 0;
      video.classList.remove("hidden");
      videoEmpty.classList.add("hidden");

      // Set VTT track. Browsers reset textTracks[0].mode when the track src
      // changes, so re-apply the user's preference now and again on the
      // track's load event in initVideoPlayer. The native overlay track carries
      // GLOBAL cue times, which only align with a single continuous file — for
      // multi-video participants it is disabled (the in-app transcript list still
      // highlights the active segment via global time).
      var track = qs("#subtitleTrack");
      if (state.videoTimeline) {
        track.removeAttribute("src");
      } else {
        track.src = "api/vtt/" + pid;
      }
      applyCaptionMode();

      // Restore the saved playback offset (GLOBAL) for this participant if we
      // have one; otherwise seek to 0.001s so the first frame renders without
      // waiting for play/scrub. `preload="metadata"` decodes duration but not
      // pixels, so without this nudge the viewer shows a blank gray box on load.
      var storedMap = getStoredUIState("transcripts").videoTimeByParticipant;
      var savedTime =
        storedMap && typeof storedMap[pid] === "number" ? storedMap[pid] : 0.001;
      if (state.videoTimeline) {
        var pi = _partForGlobal(state.videoTimeline, savedTime);
        state.videoActivePart = pi;
        state.videoOffset = state.videoTimeline[pi].cumulativeStart;
        var localStart = savedTime - state.videoOffset;
        video.src = _partMediaUrl(pi);
      } else {
        var mediaUrl = "media/" + p.video_filename;
        if (p.video_version != null) {
          mediaUrl += "?v=" + encodeURIComponent(p.video_version);
        }
        video.src = mediaUrl;
      }
      var restoreTime = function () {
        video.removeEventListener("loadedmetadata", restoreTime);
        if (state.selectedParticipant !== pid) return;
        video.currentTime = state.videoTimeline ? localStart : savedTime;
      };
      video.addEventListener("loadedmetadata", restoreTime);
    } else {
      state.videoTimeline = null;
      state.videoOffset = 0;
      state.videoActivePart = 0;
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
      _setAnalysisPanelVisible(true);
      _restoreActiveTab(pid);
      loadTranscript(pid);
      loadSummary(pid);
      loadFriction(pid);
    } else if (taskForPid && taskForPid.status === "running" && taskForPid.partial_segments && taskForPid.partial_segments.length > 0) {
      renderPartialSegments(taskForPid.partial_segments, taskForPid.progress);
      state.streamingParticipant = pid;
      clearAnalysisPanel();
    } else {
      state.segments = [];
      state.streamingParticipant = null;
      renderSegments();
      renderTimeline();
      clearAnalysisPanel();
    }
  }

  function renderEmptyState() {
    qs("#videoPlayer").classList.add("hidden");
    qs("#videoEmpty").classList.remove("hidden");
    qs("#segmentList").innerHTML = "";
    qs("#transcriptEmpty").classList.remove("hidden");
    clearAnalysisPanel();
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
  // Two cooperating pollers (createPoller handles):
  //   _summaryPoller   — runs while the backend is still generating the
  //                      summary. Stops as soon as a summary lands, or
  //                      when the user switches participant.
  //   _citationsPoller — runs after the summary arrives if citations are
  //                      still being computed (citations depend on summary).
  // Both are stopped by their own _stop*Poll() helpers; either also stops if
  // `state.selectedParticipant` no longer matches the participant the poll
  // was started for.

  var _summaryPoller = null;

  function loadSummary(pid) {
    var ver = _participantReqVer;

    // The analysis panel must be visible whenever we surface agent state for the
    // selected, transcribed participant. renderSummaryGenerating/renderSummary
    // (and the friction equivalents) don't toggle #summarySection themselves, so
    // a summary that registers *after* the transcript finalized would otherwise
    // paint its "Generating…" box into a hidden panel — visible only on reload.
    if (pid === state.selectedParticipant && _currentParticipantHasTranscript()) {
      _setAnalysisPanelVisible(true);
    }

    apiGet("api/summary/" + pid).then(function (data) {
      if (ver !== _participantReqVer) return;
      if (data.ok && data.summary) {
        _stopSummaryPoll();
        // Clear any citation state carried over from a previous participant
        // before rendering — renderSummary() reapplies state.summaryCitations,
        // so stale superscripts would otherwise leak onto this summary.
        state.summaryCitations = null;
        state.citationsGenerating = false;
        renderSummary(data.summary);
        // Handle citations
        if (data.citations && data.citations.length > 0) {
          state.summaryCitations = data.citations;
          state.citationsGenerating = false;
          renderCitations();
        } else if (data.citations_generating) {
          state.summaryCitations = null;
          state.citationsGenerating = true;
          renderCitationsStatus(
            data.citations_started_at ? data.citations_started_at * 1000 : undefined
          );
          _startCitationsPoll(pid);
          _refreshAgentStateNow();
        }
      } else if (data.generating) {
        renderSummaryGenerating(data.started_at ? data.started_at * 1000 : undefined);
        _startSummaryPoll(pid);
        _refreshAgentStateNow();
      } else {
        renderSummaryEmpty();
      }
    }).catch(function () {
      if (ver !== _participantReqVer) return;
      // Ollama unavailable or no summary — show the empty-state CTA.
      renderSummaryEmpty();
    });
  }

  function _startSummaryPoll(pid) {
    _stopSummaryPoll();
    var ver = _participantReqVer;
    // runImmediately is false to match the previous setInterval (first poll after 3s).
    _summaryPoller = createPoller(function () {
      if (ver !== _participantReqVer || state.selectedParticipant !== pid) {
        _stopSummaryPoll();
        return;
      }
      apiGet("api/summary/" + pid).then(function (data) {
        if (ver !== _participantReqVer) return;
        if (data.ok && data.summary) {
          _stopSummaryPoll();
          // Clear stale citation state before render (see loadSummary).
          state.summaryCitations = null;
          state.citationsGenerating = false;
          renderSummary(data.summary);
          // Summary just arrived — check citation status
          if (data.citations && data.citations.length > 0) {
            state.summaryCitations = data.citations;
            state.citationsGenerating = false;
            renderCitations();
          } else if (data.citations_generating) {
            state.summaryCitations = null;
            state.citationsGenerating = true;
            renderCitationsStatus(
              data.citations_started_at ? data.citations_started_at * 1000 : undefined
            );
            _startCitationsPoll(pid);
          }
        } else if (!data.generating) {
          // Generation finished without result — stop polling
          _stopSummaryPoll();
          renderSummaryEmpty();
        }
      }).catch(function () {
        if (ver !== _participantReqVer) return;
        _stopSummaryPoll();
        renderSummaryEmpty();
      });
    }, 3000, { runImmediately: false });
    _summaryPoller.start();
  }

  function _stopSummaryPoll() {
    if (_summaryPoller) {
      _summaryPoller.stop();
      _summaryPoller = null;
    }
  }

  // startedAtMs (optional): server-recorded run start in epoch ms. Seeds the
  // elapsed clock so navigating away and back resumes from the true elapsed
  // time instead of zero; omit it for a just-clicked manual run (starts now).
  function renderSummaryGenerating(startedAtMs) {
    var content = qs("#summaryContent");
    content.innerHTML =
      '<p class="summary-generating">Generating summary\u2026' +
      '<span class="agent-elapsed" id="summaryElapsed"></span>' +
      '<button type="button" class="agent-cancel-btn" id="summaryCancel">Cancel</button></p>';
    qs("#summaryEmpty").classList.add("hidden");
    qs("#summaryBody").classList.remove("hidden");
    qs("#summaryActions").classList.add("hidden");
    qs("#summaryCancel").addEventListener("click", function (e) {
      e.stopPropagation();
      _stopSummaryRun();
    });
    state.summaryEditing = false;
    state.summaryText = "";
    _summaryEtaTracker.start(startedAtMs || undefined);
    _updateAgentElapsed("summaryElapsed", _summaryEtaTracker);
    _txEtaTicker.ensure();
  }

  function renderSummaryEmpty() {
    _stopSummaryPoll();
    _stopCitationsPoll();
    _summaryEtaTracker.reset();
    qs("#summaryContent").innerHTML = "";
    qs("#summaryBody").classList.add("hidden");
    qs("#summaryActions").classList.add("hidden");
    qs("#summaryEmpty").classList.remove("hidden");
    state.summaryEditing = false;
    state.summaryText = "";
    state.summaryCitations = null;
    state.citationsGenerating = false;
  }

  function renderSummary(text) {
    _summaryEtaTracker.reset();
    state.summaryText = text;
    state.summaryEditing = false;
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
    qs("#summaryEmpty").classList.add("hidden");
    qs("#summaryBody").classList.remove("hidden");
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
    qs("#summaryContent").innerHTML = "";
    qs("#summaryActions").classList.add("hidden");
    qs("#summaryBody").classList.add("hidden");
    qs("#summaryEmpty").classList.add("hidden");
    state.summaryEditing = false;
    state.summaryText = "";
    state.summaryCitations = null;
    state.citationsGenerating = false;
  }

  // ---- Analysis panel (tabbed shell: Summary + Friction) ----

  function _setAnalysisPanelVisible(show) {
    qs("#summarySection").classList.toggle("hidden", !show);
  }

  function clearAnalysisPanel() {
    _setAnalysisPanelVisible(false);
    clearSummary();
    clearFriction();
  }

  function selectTab(name) {
    state.activeTab = name;
    var pid = state.selectedParticipant;
    if (pid) setStoredUIStateField("transcripts", "tabByParticipant", _withKey(getStoredUIState("transcripts").tabByParticipant, pid, name));
    var isSummary = name === "summary";
    qs("#tabBtnSummary").classList.toggle("active", isSummary);
    qs("#tabBtnSummary").setAttribute("aria-selected", isSummary ? "true" : "false");
    qs("#tabBtnFriction").classList.toggle("active", !isSummary);
    qs("#tabBtnFriction").setAttribute("aria-selected", !isSummary ? "true" : "false");
    qs("#summaryTab").classList.toggle("hidden", !isSummary);
    qs("#frictionTab").classList.toggle("hidden", isSummary);
  }

  function _withKey(obj, key, value) {
    var next = obj && typeof obj === "object" ? obj : {};
    next[key] = value;
    return next;
  }

  function _restoreActiveTab(pid) {
    var map = getStoredUIState("transcripts").tabByParticipant;
    var saved = map && map[pid] === "friction" ? "friction" : "summary";
    selectTab(saved);
  }

  function initPanelTabs() {
    qs("#tabBtnSummary").addEventListener("click", function () { selectTab("summary"); });
    qs("#tabBtnFriction").addEventListener("click", function () { selectTab("friction"); });
    qs("#summaryRunCta").addEventListener("click", function () { _startSummaryRun(); });
  }

  // ---- Citation rendering (Pass 2) ----

  var _citationsPoller = null;

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

  // startedAtMs (optional): server-recorded run start in epoch ms \u2014 seeds the
  // elapsed clock so it survives navigation; omit for a just-clicked manual run.
  function renderCitationsStatus(startedAtMs) {
    // Remove any existing status
    var existing = qs("#summaryContent .citations-status");
    if (existing) existing.remove();

    var p = document.createElement("p");
    p.className = "citations-status";
    p.textContent = "Finding sources\u2026";
    var sp = document.createElement("span");
    sp.className = "agent-elapsed";
    sp.id = "citationsElapsed";
    p.appendChild(sp);
    var cancel = document.createElement("button");
    cancel.type = "button";
    cancel.className = "agent-cancel-btn";
    cancel.textContent = "Cancel";
    cancel.addEventListener("click", function (e) {
      e.stopPropagation();
      _stopCitationsRun();
    });
    p.appendChild(cancel);
    qs("#summaryContent").appendChild(p);
    // Reset before seeding so a re-run adopts the new (server) start instead of
    // the idempotent tracker clinging to a prior run's start. Teardown
    // (_stopCitationsPoll) is timer-only — mirroring summary/friction — so the
    // _startCitationsPoll() that follows this call can't wipe this seed.
    _citationsEtaTracker.reset();
    _citationsEtaTracker.start(startedAtMs || undefined);
    _updateAgentElapsed("citationsElapsed", _citationsEtaTracker);
    _txEtaTicker.ensure();
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
    // runImmediately is false to match the previous setInterval (first poll after 3s).
    _citationsPoller = createPoller(function () {
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
    }, 3000, { runImmediately: false });
    _citationsPoller.start();
  }

  // Poller teardown only — does NOT reset _citationsEtaTracker (renderCitationsStatus
  // resets-then-seeds), mirroring _stopSummaryPoll / _stopFrictionPoll. Resetting
  // here would wipe the seed when _startCitationsPoll restarts the poll.
  function _stopCitationsPoll() {
    if (_citationsPoller) {
      _citationsPoller.stop();
      _citationsPoller = null;
    }
  }

  function _stopCitationsRun() {
    var pid = state.selectedParticipant;
    if (!pid) return;
    // Citations run after the summary exists, so keep the summary visible and
    // only remove the "Finding sources…" status line.
    _stopCitationsPoll();
    state.citationsGenerating = false;
    var status = qs("#summaryContent .citations-status");
    if (status) status.remove();
    apiPost("api/citations/" + pid + "/stop", {}).then(function () {
      _refreshAgentStateNow();
    }).catch(function () {});
  }

  function _startSummaryRun() {
    var pid = state.selectedParticipant;
    if (!pid) return;
    ensureAgentModelInstalled("summary").then(function (ok) {
      if (!ok) return;
      var previousText = state.summaryText;
      state.summaryCitations = null;
      state.citationsGenerating = false;
      _stopCitationsPoll();
      renderSummaryGenerating();
      apiPost("api/summary/" + pid + "/regenerate", {}).then(function (data) {
        if (data.ok && data.generating) {
          _startSummaryPoll(pid);
          _refreshAgentStateNow();
        }
      }).catch(function () {
        showToast("Failed to regenerate summary");
        if (previousText) renderSummary(previousText);
        else renderSummaryEmpty();
      });
    });
  }

  function _stopSummaryRun() {
    var pid = state.selectedParticipant;
    if (!pid) return;
    _stopSummaryPoll();
    _stopCitationsPoll();
    state.citationsGenerating = false;
    // Regenerate runs summary → citations as a chain, so summary may have
    // already finished and citations started by the time Cancel is clicked
    // (3s poll gap). Stop both — each call is a no-op if that pass isn't
    // running. Re-sync only after both stops are acknowledged, otherwise the
    // follow-up GET can still see citations in-flight and restart its poll.
    Promise.all([
      apiPost("api/summary/" + pid + "/stop", {}),
      apiPost("api/citations/" + pid + "/stop", {}),
    ]).then(function () {
      _refreshAgentStateNow();
      loadSummary(pid); // re-sync with backend (mirrors friction's loadFriction)
    }).catch(function () {});
    renderSummaryEmpty();
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
      _startSummaryRun();
    });

    qs("#citationsRegenerate").addEventListener("click", function (e) {
      e.stopPropagation();
      var pid = state.selectedParticipant;
      if (!pid) return;
      ensureAgentModelInstalled("citations").then(function (ok) {
        if (!ok) return;
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

  // ---- Friction detection (Pass 3) ----
  //
  // Programmatic scores + LLM moments land together in the manifest's
  // `friction` field. The tab renders stats, a score/category filter, and the
  // top moments; the timeline heatmap + segment tints read the per-segment
  // scores. Generation mirrors summary/citations (poll until done; manual
  // run/cancel bypass the global flag).

  var _frictionPoller = null;

  function _currentParticipant() {
    var pid = state.selectedParticipant;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) return state.participants[i];
    }
    return null;
  }

  function _frictionDepMet() {
    var p = _currentParticipant();
    if (!p) return false;
    if (p.has_summary) return true;
    return !!(p.agents && p.agents.summary === "done");
  }

  function _frictionCatLabel(key) {
    var cats = CLIPGEN_CONFIG.frictionCategories || [];
    for (var i = 0; i < cats.length; i++) {
      if (cats[i].key === key) return cats[i].label;
    }
    return key || "—";
  }

  function _friendlyTimeAgo(iso) {
    if (!iso) return "just now";
    var then = Date.parse(iso);
    if (isNaN(then)) return "recently";
    var secs = Math.max(0, Math.round((Date.now() - then) / 1000));
    if (secs < 60) return "just now";
    var mins = Math.round(secs / 60);
    if (mins < 60) return mins + (mins === 1 ? " minute ago" : " minutes ago");
    var hrs = Math.round(mins / 60);
    if (hrs < 24) return hrs + (hrs === 1 ? " hour ago" : " hours ago");
    var days = Math.round(hrs / 24);
    return days + (days === 1 ? " day ago" : " days ago");
  }

  function loadFriction(pid) {
    var ver = _participantReqVer;
    // Reveal the analysis panel for the selected, transcribed participant (see
    // loadSummary) so a friction run that registers after finalize is visible.
    if (pid === state.selectedParticipant && _currentParticipantHasTranscript()) {
      _setAnalysisPanelVisible(true);
    }
    state.frictionData = null;
    state.frictionBySegId = {};
    state.frictionGenerating = false;
    apiGet("api/friction/" + pid).then(function (data) {
      if (ver !== _participantReqVer) return;
      if (data.ok && data.friction) {
        _setFrictionData(data.friction);
      } else if (data.generating) {
        state.frictionGenerating = true;
        state.frictionStartedAt = data.started_at ? data.started_at * 1000 : null;
        renderFrictionGenerating();
        _startFrictionPoll(pid);
        _refreshAgentStateNow();
      } else {
        renderFrictionEmpty();
      }
      _refreshHeatmapAffordances();
    }).catch(function () {
      if (ver !== _participantReqVer) return;
      renderFrictionEmpty();
      _refreshHeatmapAffordances();
    });
  }

  function _setFrictionData(friction) {
    state.frictionData = friction;
    state.frictionGenerating = false;
    var byId = {};
    var segs = (friction && friction.segments) || [];
    for (var i = 0; i < segs.length; i++) {
      if (segs[i] && segs[i].id) byId[segs[i].id] = segs[i];
    }
    state.frictionBySegId = byId;
    renderFriction();
    updateFrictionStaleDot();
    _refreshHeatmapAffordances();
    if (state.frictionHeatmapEnabled) {
      renderSegments();
      renderTimeline();
    }
  }

  function clearFriction() {
    _stopFrictionPoll();
    state.frictionData = null;
    state.frictionBySegId = {};
    state.frictionGenerating = false;
    qs("#frictionContent").classList.add("hidden");
    qs("#frictionGenerating").classList.add("hidden");
    qs("#frictionEmpty").classList.add("hidden");
    updateFrictionStaleDot();
    _refreshHeatmapAffordances();
  }

  function _renderFrictionHeader() {
    var statusEl = qs("#frictionStatus");
    var rerun = qs("#frictionRerun");
    var cancel = qs("#frictionCancel");
    if (state.frictionGenerating) {
      statusEl.textContent = "Analyzing friction…";
      var sp = document.createElement("span");
      sp.className = "agent-elapsed";
      sp.id = "frictionElapsed";
      statusEl.appendChild(sp);
      statusEl.classList.remove("friction-status--stale");
      rerun.classList.add("hidden");
      cancel.classList.remove("hidden");
      _frictionEtaTracker.start(state.frictionStartedAt || undefined);
      _updateAgentElapsed("frictionElapsed", _frictionEtaTracker);
      _txEtaTicker.ensure();
      return;
    }
    _frictionEtaTracker.reset();
    cancel.classList.add("hidden");
    rerun.classList.remove("hidden");
    rerun.textContent = state.frictionData ? "Re-run friction" : "Run friction analysis";
    var depMet = _frictionDepMet();
    if (depMet) {
      rerun.removeAttribute("disabled");
      rerun.title = "";
    } else {
      rerun.setAttribute("disabled", "disabled");
      rerun.title = "Requires a summary first";
    }
    if (state.frictionData) {
      var fd = state.frictionData;
      var llmFailed = fd.llm_ok === false;
      if (llmFailed) {
        statusEl.textContent =
          "Moment detection failed — model unavailable" +
          (fd.model ? " (tried " + fd.model + ")" : "") +
          ". Showing programmatic scores; re-run with an installed model.";
      } else if (fd.stale) {
        statusEl.textContent = "Stale — segments edited since last run" +
          (fd.model ? " · " + fd.model : "");
      } else {
        statusEl.textContent = "Computed " + _friendlyTimeAgo(fd.computed_at) +
          (fd.model ? " · " + fd.model : "");
      }
      statusEl.classList.toggle("friction-status--stale", !!fd.stale && !llmFailed);
      statusEl.classList.toggle("friction-status--error", llmFailed);
    } else {
      statusEl.textContent = depMet ? "" : "Requires a summary first.";
      statusEl.classList.remove("friction-status--stale");
      statusEl.classList.remove("friction-status--error");
    }
  }

  function renderFriction() {
    _renderFrictionHeader();
    if (state.frictionGenerating) {
      renderFrictionGenerating();
      return;
    }
    if (state.frictionData) {
      qs("#frictionEmpty").classList.add("hidden");
      qs("#frictionGenerating").classList.add("hidden");
      qs("#frictionContent").classList.remove("hidden");
      renderFrictionStats();
      renderFrictionFilterControls();
      renderFrictionMoments();
    } else {
      renderFrictionEmpty();
    }
  }

  function renderFrictionEmpty() {
    state.frictionGenerating = false;
    qs("#frictionContent").classList.add("hidden");
    qs("#frictionGenerating").classList.add("hidden");
    qs("#frictionEmpty").classList.remove("hidden");
    qs("#frictionEmptyHint").textContent = _frictionDepMet()
      ? "Run the analysis to surface moments of likely friction."
      : "Requires a summary first — run Summary, then friction.";
    _renderFrictionHeader();
    updateFrictionStaleDot();
  }

  function renderFrictionGenerating() {
    qs("#frictionContent").classList.add("hidden");
    qs("#frictionEmpty").classList.add("hidden");
    qs("#frictionGenerating").classList.remove("hidden");
    _renderFrictionHeader();
  }

  function updateFrictionStaleDot() {
    var dot = qs("#frictionStaleDot");
    if (!dot) return;
    dot.classList.toggle("hidden", !(state.frictionData && state.frictionData.stale));
  }

  function _startFrictionPoll(pid) {
    _stopFrictionPoll();
    var started = Date.now();
    var ver = _participantReqVer;
    // runImmediately is false to match the previous setInterval (first poll after 3s).
    _frictionPoller = createPoller(function () {
      if (ver !== _participantReqVer ||
          state.selectedParticipant !== pid ||
          Date.now() - started > _CITATIONS_POLL_TIMEOUT) {
        _stopFrictionPoll();
        state.frictionGenerating = false;
        renderFriction();
        return;
      }
      apiGet("api/friction/" + pid).then(function (data) {
        if (ver !== _participantReqVer) return;
        if (data.ok && data.friction) {
          _stopFrictionPoll();
          _setFrictionData(data.friction);
        } else if (!data.generating) {
          _stopFrictionPoll();
          state.frictionGenerating = false;
          renderFrictionEmpty();
        }
      }).catch(function () {
        if (ver !== _participantReqVer) return;
        _stopFrictionPoll();
        state.frictionGenerating = false;
        renderFrictionEmpty();
      });
    }, 3000, { runImmediately: false });
    _frictionPoller.start();
  }

  function _stopFrictionPoll() {
    if (_frictionPoller) {
      _frictionPoller.stop();
      _frictionPoller = null;
    }
  }

  function _startFrictionRun() {
    var pid = state.selectedParticipant;
    if (!pid || !_frictionDepMet()) return;
    ensureAgentModelInstalled("friction").then(function (ok) {
      if (!ok) return;
      state.frictionGenerating = true;
      state.frictionStartedAt = null;
      renderFrictionGenerating();
      apiPost("api/friction/" + pid + "/regenerate", {}).then(function (data) {
        if (data.ok && data.generating) {
          _startFrictionPoll(pid);
          _refreshAgentStateNow();
        } else {
          state.frictionGenerating = false;
          renderFriction();
        }
      }).catch(function () {
        showToast("Failed to start friction analysis");
        state.frictionGenerating = false;
        renderFriction();
      });
    });
  }

  function _stopFrictionRun() {
    var pid = state.selectedParticipant;
    if (!pid) return;
    _stopFrictionPoll();
    state.frictionGenerating = false;
    apiPost("api/friction/" + pid + "/stop", {}).then(function () {
      _refreshAgentStateNow();
      loadFriction(pid);
    }).catch(function () {});
    renderFriction();
  }

  function renderFrictionStats() {
    var el2 = qs("#frictionStats");
    el2.innerHTML = "";
    var fd = state.frictionData;
    if (!fd || !fd.stats) return;
    var byCat = fd.stats.by_category || {};
    var cats = CLIPGEN_CONFIG.frictionCategories || [];
    var chips = document.createElement("div");
    chips.className = "friction-stat-chips";
    cats.forEach(function (c) {
      var chip = document.createElement("span");
      chip.className = "friction-chip";
      var lab = el("span", "friction-chip-label", c.label);
      var cnt = el("span", "friction-chip-count", String(byCat[c.key] || 0));
      chip.appendChild(lab);
      chip.appendChild(cnt);
      chips.appendChild(chip);
    });
    el2.appendChild(chips);
    var line = document.createElement("div");
    line.className = "friction-stat-line";
    var mpm = fd.stats.markers_per_minute != null ? fd.stats.markers_per_minute : 0;
    var total = fd.stats.total_markers != null ? fd.stats.total_markers : 0;
    line.textContent = mpm + " markers/min · " + total + " total";
    el2.appendChild(line);
  }

  function _ensureFrictionFilter() {
    if (state.frictionCategoryFilter) return;
    var f = {};
    var cats = CLIPGEN_CONFIG.frictionCategories || [];
    for (var i = 0; i < cats.length; i++) f[cats[i].key] = true;
    state.frictionCategoryFilter = f;
  }

  function renderFrictionFilterControls() {
    _ensureFrictionFilter();
    var wrap = qs("#frictionCategoryToggles");
    wrap.innerHTML = "";
    var cats = CLIPGEN_CONFIG.frictionCategories || [];
    cats.forEach(function (c) {
      var lab = document.createElement("label");
      lab.className = "friction-cat-toggle";
      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = state.frictionCategoryFilter[c.key] !== false;
      cb.addEventListener("change", function () {
        state.frictionCategoryFilter[c.key] = cb.checked;
        renderFrictionMoments();
      });
      lab.appendChild(cb);
      lab.appendChild(document.createTextNode(" " + c.label));
      wrap.appendChild(lab);
    });
    var slider = qs("#frictionThreshold");
    slider.value = String(state.frictionThreshold);
    qs("#frictionThresholdVal").textContent = state.frictionThreshold.toFixed(2);
  }

  function _frictionMomentMatches(m) {
    if ((m.score || 0) < state.frictionThreshold) return false;
    var f = state.frictionCategoryFilter;
    if (f && m.category && f[m.category] === false) return false;
    return true;
  }

  function _segmentIndexById(id) {
    for (var i = 0; i < state.segments.length; i++) {
      if (state.segments[i].id === id) return i;
    }
    return -1;
  }

  function _seekToSegmentIndex(idx) {
    var seg = state.segments[idx];
    if (!seg) return;
    seekVideo(seg.start);
    if (!_cachedSegmentRows) {
      _cachedSegmentRows = qs("#segmentList").querySelectorAll(".segment-row");
    }
    var row = _cachedSegmentRows[idx];
    if (row) scrollToSegment(row);
  }

  // Resolved segment indices a moment cites, in order, valid only.
  function _momentSegmentIndices(m) {
    var idxs = [];
    var ids = m.segment_ids || [];
    for (var i = 0; i < ids.length; i++) {
      var idx = _segmentIndexById(ids[i]);
      if (idx >= 0) idxs.push(idx);
    }
    return idxs;
  }

  function _buildMomentRow(m) {
    var idxs = _momentSegmentIndices(m);
    // An unsourced moment can't be quoted or seeked to — skip it entirely.
    if (idxs.length === 0) return null;
    var firstIdx = idxs[0];

    var row = document.createElement("div");
    row.className = "friction-moment friction-moment--seekable";

    var head = document.createElement("div");
    head.className = "friction-moment-head";
    head.appendChild(el("span", "friction-cat-badge", _frictionCatLabel(m.category)));
    head.appendChild(el("span", "friction-moment-score", (m.score != null ? m.score : 0).toFixed(2)));
    head.appendChild(el("span", "friction-moment-time", formatTime(state.segments[firstIdx].start)));
    row.appendChild(head);

    // Quote the transcript line(s) the moment was detected on.
    var quote = idxs
      .map(function (i) { return (state.segments[i].text || "").trim(); })
      .filter(Boolean)
      .join(" ");
    if (quote) row.appendChild(el("blockquote", "friction-moment-quote", quote));

    if (m.rationale) row.appendChild(el("div", "friction-moment-rationale", m.rationale));

    row.addEventListener("click", function () { _seekToSegmentIndex(firstIdx); });
    return row;
  }

  function renderFrictionMoments() {
    _ensureFrictionFilter();
    var el2 = qs("#frictionMoments");
    el2.innerHTML = "";
    var fd = state.frictionData;
    if (!fd) return;
    var moments = (fd.moments || []).filter(_frictionMomentMatches).filter(function (m) {
      return _momentSegmentIndices(m).length > 0;
    });
    if (moments.length === 0) {
      var msg;
      if (fd.llm_ok === false) {
        msg = "Moment detection failed — re-run with an installed Ollama model.";
      } else if (fd.moments && fd.moments.length) {
        msg = "No moments match the current filter.";
      } else {
        msg = "No moments detected.";
      }
      el2.appendChild(el("p", "friction-moments-empty", msg));
      return;
    }
    var frag = document.createDocumentFragment();
    moments.forEach(function (m) {
      var rowEl = _buildMomentRow(m);
      if (rowEl) frag.appendChild(rowEl);
    });
    el2.appendChild(frag);
  }

  function _primaryCategory(frow) {
    var counts = frow.counts || {};
    var cats = frow.categories || [];
    var best = cats[0] || null;
    var bestN = best ? (counts[best] || 0) : 0;
    for (var i = 1; i < cats.length; i++) {
      var n = counts[cats[i]] || 0;
      if (n > bestN) { best = cats[i]; bestN = n; }
    }
    return best;
  }

  function _frictionMarkAll() {
    var fd = state.frictionData;
    if (!fd || !fd.segments) return;
    _ensureFrictionFilter();
    var groups = {};
    fd.segments.forEach(function (frow) {
      if ((frow.score || 0) < state.frictionThreshold) return;
      var cats = frow.categories || [];
      var matched = [];
      for (var i = 0; i < cats.length; i++) {
        if (state.frictionCategoryFilter[cats[i]] !== false) matched.push(cats[i]);
      }
      if (matched.length === 0) return;
      var primary = _primaryCategory(frow);
      if (!primary || state.frictionCategoryFilter[primary] === false) primary = matched[0];
      if (!groups[primary]) groups[primary] = [];
      groups[primary].push(frow.id);
    });
    var keys = Object.keys(groups);
    if (keys.length === 0) {
      showToast("No segments match the filter");
      return;
    }
    var pending = keys.length;
    var total = 0;
    keys.forEach(function (cat) {
      var ids = groups[cat];
      total += ids.length;
      apiPost("api/marks", {
        segment_ids: ids,
        category: "friction",
        label: "Friction · " + cat,
      }).then(done).catch(done);
    });
    function done() {
      pending--;
      if (pending <= 0) {
        showToast("Marked " + clipgenPluralUnit(total, "segment", "segments"));
        if (state.selectedParticipant) loadTranscript(state.selectedParticipant);
      }
    }
  }

  function _refreshHeatmapAffordances() {
    var btn = qs("#frictionHeatmapBtn");
    if (!btn) return;
    var has = !!(state.frictionData && state.frictionData.segments && state.frictionData.segments.length);
    btn.disabled = !has;
    // Only show the pressed/active styling when the toggle is both enabled and
    // actionable — otherwise a stored-enabled toggle renders active-but-greyed
    // on load before friction data arrives.
    var showActive = has && state.frictionHeatmapEnabled;
    btn.classList.toggle("active", showActive);
    btn.setAttribute("aria-pressed", showActive ? "true" : "false");
  }

  function initFriction() {
    qs("#frictionRerun").addEventListener("click", function () { _startFrictionRun(); });
    qs("#frictionCancel").addEventListener("click", function () { _stopFrictionRun(); });
    qs("#frictionMarkAll").addEventListener("click", function () { _frictionMarkAll(); });
    var slider = qs("#frictionThreshold");
    slider.addEventListener("input", function () {
      state.frictionThreshold = parseFloat(slider.value);
      qs("#frictionThresholdVal").textContent = state.frictionThreshold.toFixed(2);
      renderFrictionMoments();
    });
  }

  function initFrictionHeatmapToggle() {
    var btn = qs("#frictionHeatmapBtn");
    if (!btn) return;
    var stored = getStoredUIState("transcripts");
    state.frictionHeatmapEnabled = !!(stored && stored.frictionHeatmapEnabled);
    // Defer the active/aria-pressed visual to _refreshHeatmapAffordances so it
    // only lights up once friction data is actually present.
    _refreshHeatmapAffordances();
    btn.addEventListener("click", function () {
      state.frictionHeatmapEnabled = !state.frictionHeatmapEnabled;
      setStoredUIStateField("transcripts", "frictionHeatmapEnabled", state.frictionHeatmapEnabled);
      _refreshHeatmapAffordances();
      renderSegments();
      renderTimeline();
    });
  }

  // Friction tooltip on hot segments (reuses the shared #trTooltip element).
  // _frictionTooltipShown lets hideTimelineTooltip yield while a friction
  // tooltip owns #trTooltip (mirror of _hideFrictionTooltip's _lastTimelineHit
  // guard); _segTooltipRaf coalesces the segment-list mousemove like the canvas.
  var _frictionTooltipShown = false;
  var _segTooltipRaf = 0;

  function _showFrictionTooltip(frow, clientX, clientY) {
    var tip = qs("#trTooltip");
    if (!tip) return;
    tip.textContent = "";
    var cats = frow.categories || [];
    if (cats.length) {
      var badges = document.createElement("div");
      badges.className = "tr-tooltip-friction-cats";
      cats.forEach(function (c) {
        badges.appendChild(el("span", "friction-cat-badge friction-cat-badge--sm", _frictionCatLabel(c)));
      });
      tip.appendChild(badges);
    }
    var markers = frow.markers || [];
    if (markers.length) {
      var shown = markers.slice(0, 5).join(", ");
      if (markers.length > 5) shown += " +" + (markers.length - 5) + " more";
      tip.appendChild(document.createTextNode(shown));
      tip.appendChild(document.createElement("br"));
    }
    tip.appendChild(el("span", "tr-tooltip-friction-score", "Score " + (frow.score || 0).toFixed(2)));
    tip.classList.remove("hidden");
    var tipRect = tip.getBoundingClientRect();
    var x = clientX + 12;
    var y = clientY - tipRect.height - 12;
    if (x + tipRect.width > window.innerWidth - 8) x = window.innerWidth - tipRect.width - 8;
    if (y < 8) y = clientY + 16;
    tip.style.left = x + "px";
    tip.style.top = y + "px";
    _frictionTooltipShown = true;
  }

  function _hideFrictionTooltip() {
    _frictionTooltipShown = false;
    var tip = qs("#trTooltip");
    if (tip && !_lastTimelineHit) tip.classList.add("hidden");
  }

  // Draw the smoothed friction density band across the timeline ruler.
  // Per-pixel averaging of overlapping segment scores (mirrors the Screenspace
  // amplitude graph's binning) gives a continuous band without a separate
  // smoothing constant; alpha scales with score.
  function _drawFrictionBand(ctx, timeToX, bandY, bandH, cssW) {
    if (!state.frictionHeatmapEnabled) return;
    if (!state.segments.length) return;
    var fcolor = getCSSVar("--color-friction", "#ea580c");
    if (fcolor.charAt(0) !== "#") fcolor = "#ea580c";
    var numBins = Math.max(1, Math.floor(cssW));
    var sums = new Array(numBins);
    var counts = new Array(numBins);
    for (var b = 0; b < numBins; b++) { sums[b] = 0; counts[b] = 0; }
    var any = false;
    for (var i = 0; i < state.segments.length; i++) {
      var seg = state.segments[i];
      var frow = state.frictionBySegId[seg.id];
      var sc = frow ? (frow.score || 0) : 0;
      if (sc <= 0) continue;
      any = true;
      var x0 = Math.max(0, Math.floor(timeToX(seg.start)));
      var x1 = Math.min(numBins - 1, Math.floor(timeToX(seg.end || seg.start)));
      if (x1 < x0) x1 = x0;
      for (var x = x0; x <= x1; x++) { sums[x] += sc; counts[x] += 1; }
    }
    if (!any) return;
    for (var px = 0; px < numBins; px++) {
      if (!counts[px]) continue;
      var v = sums[px] / counts[px];
      if (v <= 0) continue;
      ctx.fillStyle = hexToRgba(fcolor, Math.min(0.85, 0.15 + v * 0.7));
      ctx.fillRect(px, bandY, 1, bandH);
    }
  }

  // ---- Segment rendering ----

  var _cachedSegmentRows = null;

  function renderSegments() {
    var container = qs("#segmentList");
    var empty = qs("#transcriptEmpty");
    state.editingTextEl = null;
    _cachedSegmentRows = null;
    // We're rendering the finalized transcript — drop any queued streaming
    // indicator so a paused-tab RAF can't re-insert it over the real segments.
    _cancelStreamingIndicator();

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

      var frictionClass = "";
      var frictionStyle = "";
      if (state.frictionHeatmapEnabled) {
        var frow = state.frictionBySegId[seg.id];
        var fScore = frow ? (frow.score || 0) : 0;
        if (fScore > 0) {
          frictionClass = " segment-friction";
          frictionStyle = ' style="--seg-friction-alpha:' + fScore + '"';
        }
      }
      html += '<div class="segment-row' + activeClass + correctedClass + frictionClass + '" data-index="' + i + '" data-start="' + seg.start + '"' + frictionStyle + '>';
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
            html += '<span class="segment-xref-badge" style="background:' + XREF_BADGES.screenspace.color + '" title="' + escapeHtml(evTypes.join(", ")) + '"><span class="xref-badge-icon" style="' + iconMaskStyle(XREF_BADGES.screenspace.icon) + '"></span></span>';
          }
          if (xref.sheetObservations.length > 0) {
            var obsTitle = xref.sheetObservations[0].observation;
            html += '<span class="segment-xref-badge" style="background:' + XREF_BADGES.sheet.color + '" title="' + escapeHtml(obsTitle) + '"><span class="xref-badge-icon" style="' + iconMaskStyle(XREF_BADGES.sheet.icon) + '"></span></span>';
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

  // ---- Elapsed / ETA tracking ----
  // Transcription progress is a linear fraction of media duration, so its ETA
  // extrapolation is meaningful (per-participant trackers). The thinking agents
  // expose no progress fraction, so they show elapsed only (single trackers).
  var _txEtaTrackers = {};
  var _summaryEtaTracker = createEtaTracker();
  var _citationsEtaTracker = createEtaTracker();
  var _frictionEtaTracker = createEtaTracker();
  var _txEtaTicker = createIntervalTicker(_tickTxEta, {
    isActive: _anyTxEtaActive,
  });

  // " \u00b7 0:42 \u00b7 ~1:20 left" suffix for a participant's running transcription, or
  // "" when not running. Each entry is keyed by the task's created_at so a re-run
  // of the same participant seeds a fresh tracker from the new task rather than
  // continuing the prior run's elapsed (created_at includes any queue wait, so
  // elapsed may slightly overstate). Stale entries are pruned in _tickTxEta.
  function _txEtaSuffix(pid, task) {
    if (!pid || !task || task.status !== "running") return "";
    var entry = _txEtaTrackers[pid];
    if (!entry || entry.createdAt !== task.created_at) {
      var t = createEtaTracker();
      var seed = task.created_at ? Date.parse(task.created_at) : NaN;
      t.start(isNaN(seed) ? undefined : seed);
      entry = { tracker: t, createdAt: task.created_at };
      _txEtaTrackers[pid] = entry;
    }
    var e = entry.tracker.update(task.progress);
    var s = " \u00b7 " + formatDuration(e.elapsedSec);
    var eta = formatEtaLabel(e.remainingSec);
    if (eta) s += " \u00b7 " + eta;
    return s;
  }

  function _streamingTextStr(progress) {
    var pid = state.streamingParticipant || state.selectedParticipant;
    var task = _taskForSelectedParticipant();
    return "Transcribing\u2026 " + Math.round(progress * 100) + "%" + _txEtaSuffix(pid, task);
  }

  // Paint a thinking-agent's elapsed clock (elapsed only \u2014 no progress signal).
  function _updateAgentElapsed(spanId, tracker) {
    var sp = document.getElementById(spanId);
    if (!sp) return;
    var e = tracker.update();
    sp.textContent = formatDuration(e.elapsedSec);
  }

  function _anyTxEtaActive() {
    if (_anyAgentActive()) return true;
    for (var i = 0; i < state.tasks.length; i++) {
      if (state.tasks[i].status === "running") return true;
    }
    return false;
  }

  function _tickTxEta() {
    // The ticker's isActive guard (_anyTxEtaActive) self-stops it; this body
    // only runs while transcription or a thinking agent is active.
    // Drop trackers for participants with no running transcription so memory
    // stays bounded and a later re-run starts fresh.
    var runningPids = {};
    for (var r = 0; r < state.tasks.length; r++) {
      if (state.tasks[r].status === "running") runningPids[state.tasks[r].participant] = true;
    }
    Object.keys(_txEtaTrackers).forEach(function (p) {
      if (!runningPids[p]) delete _txEtaTrackers[p];
    });
    // Transcription: status-indicator tooltip + the streaming "Transcribing\u2026" line.
    updateStatusIndicator();
    var txt = document.querySelector("#segmentList .streaming-text");
    if (txt) {
      var task = _taskForSelectedParticipant();
      if (task && task.status === "running") txt.textContent = _streamingTextStr(task.progress || 0);
    }
    // Thinking agents: elapsed-only clocks (spans present only while generating).
    _updateAgentElapsed("summaryElapsed", _summaryEtaTracker);
    _updateAgentElapsed("citationsElapsed", _citationsEtaTracker);
    _updateAgentElapsed("frictionElapsed", _frictionEtaTracker);
  }

  function _streamingIndicatorHtml(progress) {
    return '<div class="streaming-indicator">' +
      '<span class="streaming-dot"></span>' +
      '<span class="streaming-text">' + _streamingTextStr(progress) + '</span>' +
      '</div>';
  }

  var _streamIndicatorRaf = null;
  var _streamIndicatorPending = null;

  // Cancel a queued streaming-indicator insert. requestAnimationFrame callbacks
  // are paused while the tab is backgrounded, so a RAF scheduled during the last
  // streaming poll can outlive the transcript being finalized and re-insert a
  // stale "Transcribing… X%" row when the user returns. Call this whenever the
  // finalized transcript replaces the streaming view.
  function _cancelStreamingIndicator() {
    if (_streamIndicatorRaf) {
      cancelAnimationFrame(_streamIndicatorRaf);
      _streamIndicatorRaf = null;
    }
    _streamIndicatorPending = null;
  }

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

    // Friction tooltip on hot segments (only while the heatmap is on).
    // RAF-coalesced like the timeline canvas so getBoundingClientRect isn't
    // called on every mousemove event.
    container.addEventListener("mousemove", function (e) {
      if (_segTooltipRaf) return;
      var cx = e.clientX, cy = e.clientY, tgt = e.target;
      _segTooltipRaf = requestAnimationFrame(function () {
        _segTooltipRaf = 0;
        if (!state.frictionHeatmapEnabled) { _hideFrictionTooltip(); return; }
        var row = tgt.closest && tgt.closest(".segment-row");
        if (!row) { _hideFrictionTooltip(); return; }
        var idx = parseInt(row.getAttribute("data-index"), 10);
        var seg = state.segments[idx];
        var frow = seg ? state.frictionBySegId[seg.id] : null;
        if (!frow || !(frow.score > 0)) { _hideFrictionTooltip(); return; }
        _showFrictionTooltip(frow, cx, cy);
      });
    });
    container.addEventListener("mouseleave", function () { _hideFrictionTooltip(); });
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
    window.ClipgenVideoControls.applyPlaybackRate(qs("#videoPlayer"), state.videoPlaybackRate);
  }

  // ---- Multi-video timeline (client-side source switching) ----
  // For a participant whose recording spans several files, p.timeline carries
  // [{filename, duration, cumulativeStart}]. The <video> plays one part at a
  // time; these helpers present a single GLOBAL timeline to the controls so the
  // playhead, labels, and segment sync all use global time. Single-video
  // participants have state.videoTimeline === null and take the original path.
  function _timelineTotal(tl) {
    if (!tl || !tl.length) return 0;
    var last = tl[tl.length - 1];
    return last.cumulativeStart + last.duration;
  }
  function _partForGlobal(tl, g) {
    for (var i = 0; i < tl.length; i++) {
      if (g >= tl[i].cumulativeStart && g < tl[i].cumulativeStart + tl[i].duration) {
        return i;
      }
    }
    return tl.length - 1;
  }
  function _partMediaUrl(i) {
    var url = "media/" + state.videoTimeline[i].filename;
    if (state.videoVersion != null) url += "?v=" + encodeURIComponent(state.videoVersion);
    return url;
  }
  function videoDisplayDuration() {
    var v = qs("#videoPlayer");
    if (state.videoTimeline) return _timelineTotal(state.videoTimeline);
    return v && isFinite(v.duration) ? v.duration : 0;
  }
  function videoGlobalTime() {
    var v = qs("#videoPlayer");
    if (!v) return 0;
    return (v.currentTime || 0) + (state.videoTimeline ? state.videoOffset : 0);
  }
  function _switchToPart(i, localTime, autoplay) {
    var v = qs("#videoPlayer");
    state.videoActivePart = i;
    state.videoOffset = state.videoTimeline[i].cumulativeStart;
    v.src = _partMediaUrl(i);
    var onMeta = function () {
      v.removeEventListener("loadedmetadata", onMeta);
      v.currentTime = localTime;
      if (autoplay) v.play();
    };
    v.addEventListener("loadedmetadata", onMeta);
  }

  function updateTimeLabel() {
    var v = qs("#videoPlayer");
    var label = qs("#videoTime");
    if (!v || !label) return;
    var dur = videoDisplayDuration();
    label.textContent = formatTime(videoGlobalTime()) + " / " + formatTime(dur);
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

    var theme = getCanvasThemeColors();

    ctx.fillStyle = theme.surfaceAlt;
    ctx.fillRect(0, 0, cssW, cssH);

    var dur = videoDisplayDuration();
    if (dur <= 0) {
      _markerHitRects = [];
      ctx.fillStyle = theme.textDim;
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
    ctx.strokeStyle = theme.border;
    ctx.fillStyle = theme.textDim;
    ctx.font = "10px " + theme.fontMono;
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

    // Friction heatmap band (behind marks).
    _drawFrictionBand(ctx, timeToX, markerY, markerH, cssW);

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
    var dur = videoDisplayDuration();
    if (dur <= 0) return;
    var px = (videoGlobalTime() / dur) * cssW;
    ctx.strokeStyle = getCanvasThemeColors().accent;
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
    var dur = videoDisplayDuration();
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
    // Yield if a friction (hot-segment) tooltip currently owns the shared
    // element, so a canvas mouseleave doesn't clobber it mid-display.
    if (_frictionTooltipShown) return;
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
      state.videoPlaybackRate = window.ClipgenVideoControls.nextSpeed(VIDEO_SPEEDS, state.videoPlaybackRate);
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
      // Multi-video: hand off to the next part as playback nears the boundary so
      // continuous playback spans the whole recording.
      if (state.videoTimeline) {
        var tl = state.videoTimeline;
        var i = state.videoActivePart;
        if (i < tl.length - 1 && video.currentTime >= tl[i].duration - 0.05) {
          _switchToPart(i + 1, 0.001, !video.paused);
        }
      }
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
    // *time* is GLOBAL. For a multi-video participant, switch the <video> source
    // to the part that owns it and seek the local offset; single-video falls
    // straight through to the original local seek.
    if (state.videoTimeline) {
      var tl = state.videoTimeline;
      var g = time < 0 ? 0 : Math.min(time, _timelineTotal(tl));
      var i = _partForGlobal(tl, g);
      var local = g - tl[i].cumulativeStart;
      if (i !== state.videoActivePart) {
        _switchToPart(i, local, true);
      } else {
        _seekLocal(local);
      }
      return;
    }
    _seekLocal(time);
  }

  function _seekLocal(time) {
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
    var save = function () { persistVideoTime(videoGlobalTime()); };
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
    var t = videoGlobalTime();

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
    if (task && task.status === "queued") return "queued";
    if (task && task.status === "running") return "running";
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
      friction: (p.agents && p.agents.friction) || "idle",
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
        var agentsAttr = s0.agents.transcription + "," + s0.agents.summary + "," + s0.agents.citations + "," + s0.agents.friction;
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
    wrap.setAttribute("data-agents", s.agents.transcription + "," + s.agents.summary + "," + s.agents.citations + "," + s.agents.friction);
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
    if (st === "queued") return "queued";
    if (st === "failed") return "failed";
    return "not started";
  }

  function buildPillDots(p, s) {
    var ag = s.agents;
    var labels = ["Transcription", "Summary", "Citations", "Friction"];
    var keys = ["transcription", "summary", "citations", "friction"];
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
        var label = m.name + (m.cached === false ? " — not downloaded" : "");
        opts += '<option value="' + escapeHtml(m.name) + '">' + escapeHtml(label) + '</option>';
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
        ensureAgentModelInstalled("summary").then(function (ok) {
          if (!ok) return;
          apiPost("api/summary/" + p.id + "/regenerate", {}).then(function () {
            _refreshAgentStateNow();
          }).catch(function () {
            showToast("Failed to start summary");
          });
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
        ensureAgentModelInstalled("citations").then(function (ok) {
          if (!ok) return;
          apiPost("api/citations/" + p.id + "/regenerate", {}).then(function () {
            _refreshAgentStateNow();
          }).catch(function () {
            showToast("Failed to start citations");
          });
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

    // 4. Friction — depends on summary only (independent of citations).
    section.appendChild(buildAgentRow({
      pid: p.id,
      label: "Friction",
      agent: "friction",
      depLabel: "summary",
      depMet: s.agents.summary === "done",
      agentState: s.agents.friction,
      hasResult: !!(p.agents && p.agents.friction === "done"),
      cascadeWarning: false,
      onStart: function () {
        ensureAgentModelInstalled("friction").then(function (ok) {
          if (!ok) return;
          apiPost("api/friction/" + p.id + "/regenerate", {}).then(function () {
            _refreshAgentStateNow();
            if (state.selectedParticipant === p.id) loadFriction(p.id);
          }).catch(function () {
            showToast("Failed to start friction");
          });
        });
      },
      onStop: function () {
        apiPost("api/friction/" + p.id + "/stop", {}).then(function () {
          _refreshAgentStateNow();
          if (state.selectedParticipant === p.id) loadFriction(p.id);
        }).catch(function () {
          showToast("Failed to stop friction");
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
    _postTranscribe(pids, force, overrides, false);
  }

  // POST to the transcribe endpoint. The server is the authority on whether a
  // Whisper model is cached: when it rejects with reason "model_not_cached", we
  // confirm each uncached download with the user and retry with allow_download.
  function _postTranscribe(pids, force, overrides, allowDownload) {
    var body = { participants: pids, force: force };
    if (overrides && Object.keys(overrides).length > 0) body.overrides = overrides;
    if (allowDownload) body.allow_download = true;
    apiPost("api/transcribe", body).then(function (data) {
      if (!data.ok) {
        if (data.reason === "model_not_cached" && data.uncached && data.uncached.length) {
          _confirmUncachedWhisperModels(data.uncached).then(function (ok) {
            if (ok) _postTranscribe(pids, force, overrides, true);
          });
          return;
        }
        showToast("Failed to enqueue transcription");
        return;
      }
      showToast("Enqueued " + clipgenPluralUnit(data.tasks.length, "transcription", "transcriptions"));
      startPolling();
      pollTaskStatus();
    });
  }

  // Confirm each distinct non-cached Whisper model in turn; any cancel aborts.
  function _confirmUncachedWhisperModels(uncached) {
    return uncached.reduce(function (chain, m) {
      return chain.then(function (okSoFar) {
        if (!okSoFar) return false;
        return confirmModelInstall({
          kind: "whisper",
          model: m.model,
          sizeMb: m.size_mb,
        }).then(function (ok) {
          if (ok) {
            _whisperDownloadConfirmed[m.model] = true;
            _trModelsCache = null;
            _trModelsCachePromise = null;
          }
          return ok;
        });
      });
    }, Promise.resolve(true));
  }

  // ---- Task polling ----
  // state.pollPoller → pollTaskStatus for Whisper jobs; summary/citations use
  // separate pollers (see file header).

  function startPolling() {
    if (state.pollPoller) return;
    // runImmediately is false to match the previous setInterval (first poll after POLL_INTERVAL).
    state.pollPoller = createPoller(pollTaskStatus, POLL_INTERVAL, { runImmediately: false });
    state.pollPoller.start();
  }

  function stopPolling() {
    if (state.pollPoller) {
      state.pollPoller.stop();
      state.pollPoller = null;
    }
  }

  // Keyed by task id (NOT participant) so each task's completion is handled
  // exactly once. A participant can have several completed tasks over a session
  // (re-transcription creates a new task while old completed tasks linger in the
  // worker and are restored from the manifest), so a participant key would let a
  // stale completed task suppress the new run's completion transition.
  var _refreshedCompletedTaskIds = {};
  // After a whisper task flips to "completed", keep the main poll loop alive
  // for a few cycles so the summary "Generating…" state surfaces even if the
  // next /api/participants or /api/summary response races the in-flight slot.
  var _postCompletionGrace = 0;
  var POST_COMPLETION_GRACE_CYCLES = 4; // ~12s at POLL_INTERVAL=3000ms

  function _anyAgentActive() {
    for (var i = 0; i < state.participants.length; i++) {
      var ag = state.participants[i].agents;
      if (ag && (ag.summary === "running" || ag.citations === "running" || ag.friction === "running")) return true;
    }
    return false;
  }

  // During the post-completion grace window an agent may register as running a
  // cycle or two after the whisper task completes. Re-load the selected
  // participant's summary so renderSummaryGenerating()/_startSummaryPoll fire
  // once the backend reports it generating; also refresh friction when its dep
  // is met. Guarded so it won't stomp an already-armed/rendered panel.
  function _rearmSelectedAgentPanels() {
    var pid = state.selectedParticipant;
    if (!pid) return;
    var p = _currentParticipant();
    var summaryRunning = !!(p && p.agents && p.agents.summary === "running");
    if (summaryRunning && !_summaryPoller && !state.summaryText) loadSummary(pid);
    if (!state.frictionData && !state.frictionGenerating && _frictionDepMet()) loadFriction(pid);
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

  // Reconcile a participant that is still showing the streaming view against the
  // backend: if its whisper task has finished (no running/queued task remains
  // and a completed one exists) and its transcript is ready, swap the streaming
  // "Transcribing… X%" footer for the finalized transcript and reveal the
  // analysis panel — mirroring selectParticipant's has_transcript path. Called
  // every poll, so it self-heals: state.streamingParticipant is only cleared
  // once the finalized transcript has actually rendered, so a transient API
  // failure just retries on the next poll instead of freezing the footer.
  function _finalizeStreamingIfComplete(pid) {
    var running = false;
    var completed = false;
    for (var i = 0; i < state.tasks.length; i++) {
      var t = state.tasks[i];
      if (t.participant !== pid) continue;
      if (t.status === "running" || t.status === "queued") running = true;
      else if (t.status === "completed") completed = true;
    }
    // Still transcribing → leave the streaming view up; a later poll re-enters.
    if (running) return;
    // No running task and nothing completed (failed/cancelled/dismissed) → drop
    // the streaming flag so the normal empty/failed rendering can take over and
    // the poll loop can wind down (matches the prior behaviour here).
    if (!completed) {
      state.streamingParticipant = null;
      return;
    }
    var ver = _participantReqVer;
    apiGet("api/transcript/" + pid).then(function (data) {
      if (ver !== _participantReqVer || state.selectedParticipant !== pid) return;
      // Transcript not merged yet (rare race) → keep streamingParticipant set so
      // the next poll retries rather than leaving a frozen footer.
      if (!(data.ok && data.segments && data.segments.length > 0)) return;
      state.streamingParticipant = null;
      state.segments = data.segments;
      state.activeSegmentIndex = -1;
      renderSegments();
      renderTimeline();
      _setAnalysisPanelVisible(true);
      _restoreActiveTab(pid);
      loadSummary(pid);
      loadFriction(pid);
    });
  }

  function pollTaskStatus() {
    apiGet("api/transcribe/status").then(function (data) {
      if (!data.ok) return;
      state.tasks = data.tasks;
      if (_anyTxEtaActive()) _txEtaTicker.ensure();

      // Snapshot before _finalizeStreamingIfComplete can clear streamingParticipant
      // in its async /api/transcript callback; the newlyCompleted refresh must not
      // double-fire loadTranscript/loadSummary/loadFriction for that transition.
      var wasStreamingSelected =
        state.streamingParticipant === state.selectedParticipant;

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
        // The selected participant's streaming view is up but it has no running
        // task. Finalize in place once its transcript is ready. This is
        // state-driven and retried every poll (not a one-shot tied to the
        // newlyCompleted de-dup), so a transient /api/participants or
        // /api/transcript failure can't leave the "Transcribing… X%" footer
        // frozen forever. It owns clearing state.streamingParticipant.
        _finalizeStreamingIfComplete(state.streamingParticipant);
      }

      var hasActive = false;
      var newlyCompleted = [];
      data.tasks.forEach(function (t) {
        if (t.status === "queued" || t.status === "running") hasActive = true;
        if (t.status === "completed" && !_refreshedCompletedTaskIds[t.id]) {
          newlyCompleted.push(t.participant);
          _refreshedCompletedTaskIds[t.id] = true;
        }
      });

      // Refresh participants and transcript as each task completes.
      // Thinking-agents (summary → citations) are spawned on whisper completion
      // and on server startup, so we always refresh after anything completes
      // or if any agent is currently running on any pill.
      // Refresh during the grace window too, so participants are re-fetched each
      // cycle until the summary agent registers as running (or grace expires).
      var needsRefresh =
        newlyCompleted.length > 0 || _anyAgentActive() || _postCompletionGrace > 0;
      if (newlyCompleted.length > 0) {
        // Re-arm on every fresh completion so a multi-participant queue keeps
        // extending the grace window.
        _postCompletionGrace = POST_COMPLETION_GRACE_CYCLES;
        _streamingMarks = {};
        _streamingMarksLoaded = false;
        _bumpStreamingMarksVersion();
      }
      if (needsRefresh) {
        loadParticipants().then(function () {
          if (newlyCompleted.length > 0 && state.selectedParticipant &&
              newlyCompleted.indexOf(state.selectedParticipant) >= 0 &&
              !wasStreamingSelected) {
            // The selected participant finished but was NOT mid-stream in this
            // view (e.g. it was queued/idle when it completed), so reveal the
            // analysis panel and load the finalized transcript here, mirroring
            // selectParticipant's has_transcript path. The streaming→done case
            // is owned by _finalizeStreamingIfComplete (state-driven + retried),
            // so skip when wasStreamingSelected to avoid double-firing after it
            // clears streamingParticipant and loads summary/friction first.
            _setAnalysisPanelVisible(true);
            _restoreActiveTab(state.selectedParticipant);
            loadTranscript(state.selectedParticipant);
            loadSummary(state.selectedParticipant);
            loadFriction(state.selectedParticipant);
          } else if (state.selectedParticipant) {
            // Summary/citations/friction may be auto-chaining server-side after
            // an earlier completion, or registering during the grace window;
            // re-arm the selected participant's panels so the running state
            // surfaces without a manual reload.
            _rearmSelectedAgentPanels();
          }
          updateStatusIndicator();
          // Agents typically kick in right after whisper completes; keep the
          // poll alive across the grace window so dot transitions (running →
          // done → next) are seen even before the agent registers as running.
          // Also stay alive while a streaming view still needs finalizing, so
          // _finalizeStreamingIfComplete keeps retrying until it renders.
          if (_anyAgentActive() || _postCompletionGrace > 0 || state.streamingParticipant) startPolling();
          else if (!hasActive) {
            stopPolling();
            _refreshedCompletedTaskIds = {};
          }
        });
      }

      if (hasActive || _anyAgentActive() || _postCompletionGrace > 0 || state.streamingParticipant) {
        startPolling();
      } else if (!needsRefresh) {
        stopPolling();
        _refreshedCompletedTaskIds = {};
      }

      if (!hasActive && _hadActiveTranscriptionLastPoll) {
        refreshTranscriptionModelHintOnce();
        // Transcription just finished — re-validate downloads against real
        // cache state (downloaded → no prompt, failed → prompt again).
        _forgetWhisperDownloadAgreements();
      }
      _hadActiveTranscriptionLastPoll = hasActive;

      // Count down the grace window at the end so the current cycle still
      // counts as "in grace" for the decisions above.
      if (_postCompletionGrace > 0) _postCompletionGrace--;

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
  // Whisper models the user has already agreed to download this session. The
  // models response is cached and a just-confirmed model won't read back as
  // cached until its background download finishes, so without this we'd
  // re-prompt on every transcription. Keyed by model name.
  var _whisperDownloadConfirmed = {};
  // Serializes confirmModelInstall() calls: there is one shared modal element,
  // so overlapping callers (e.g. prewarm + an agent run) must take turns.
  var _modelInstallChain = Promise.resolve();

  function _trFetchModels() {
    if (_trModelsCache) return Promise.resolve(_trModelsCache);
    if (_trModelsCachePromise) return _trModelsCachePromise;
    _trModelsCachePromise = fetch("/api/models")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        // Don't pin a result where Ollama wasn't reachable — otherwise the
        // agent gate and pickers stay blind to installed models for the whole
        // session. Reset so the next call re-fetches once the server is up.
        if (data && data.ok && !(data.ollama && data.ollama.available === false)) {
          _trModelsCache = data;
        } else {
          _trModelsCachePromise = null;
        }
        return data;
      })
      .catch(function () { _trModelsCachePromise = null; return null; });
    return _trModelsCachePromise;
  }

  // ---- Local-model install confirmation ----
  // Both whisper transcription models and Ollama agent models are "local
  // models" that get installed on demand. We never download one silently:
  // confirmModelInstall() gates every install behind an explicit dialog.

  function _trFormatModelSize(mb) {
    if (!mb || mb <= 0) return "";
    if (mb >= 1024) return (mb / 1024).toFixed(1) + " GB";
    return Math.round(mb) + " MB";
  }

  // Stream an Ollama pull, reporting progress dicts via onProgress. Resolves
  // true on success, false on failure/cancellation. isCancelled() is polled
  // each tick so a dismissed dialog stops the poll (the server-side pull keeps
  // running and resumes from cached layers). The poll also self-terminates
  // after a run of unanswered status checks so it can never leak forever.
  function installOllamaModel(model, onProgress, isCancelled) {
    return apiPost("api/models/ollama/pull", { model: model }).then(function (data) {
      if (!data || !data.ok) return false;
      return new Promise(function (resolve) {
        var misses = 0;
        var poll = setInterval(function () {
          if (isCancelled && isCancelled()) {
            clearInterval(poll);
            resolve(false);
            return;
          }
          apiGet("api/models/ollama/pull-status?model=" + encodeURIComponent(model))
            .then(function (st) {
              if (!st || !st.ok || !st.found) {
                if (++misses >= 20) { clearInterval(poll); resolve(false); }
                return;
              }
              misses = 0;
              if (onProgress) onProgress(st);
              if (st.done) {
                clearInterval(poll);
                resolve(!!st.succeeded);
              }
            })
            .catch(function () {
              if (++misses >= 20) { clearInterval(poll); resolve(false); }
            });
        }, 1000);
      });
    }).catch(function () { return false; });
  }

  // Show the confirm/install dialog. Resolves true when the model is available
  // to use (whisper: user agreed to the download; ollama: pull succeeded),
  // false when the user cancels or the install fails. Calls are serialized
  // (see _modelInstallChain) so concurrent callers never share the one modal.
  function confirmModelInstall(opts) {
    var run = function () { return _confirmModelInstallNow(opts); };
    var result = _modelInstallChain.then(run, run);
    // Advance the chain when this dialog settles, swallowing its outcome so a
    // cancelled/failed dialog doesn't break the queue for the next caller.
    _modelInstallChain = result.then(function () {}, function () {});
    return result;
  }

  function _confirmModelInstallNow(opts) {
    return new Promise(function (resolve) {
      var cancelled = false;
      var modal = qs("#modelInstallModal");
      var titleEl = qs("#modelInstallTitle");
      var msgEl = qs("#modelInstallMessage");
      var progress = qs("#modelInstallProgress");
      var barFill = qs("#modelInstallBarFill");
      var progressText = qs("#modelInstallProgressText");
      var cancelBtn = qs("#modelInstallCancel");
      var confirmBtn = qs("#modelInstallConfirm");

      progress.classList.add("hidden");
      barFill.style.width = "0%";
      progressText.textContent = "";
      cancelBtn.disabled = false;
      cancelBtn.textContent = "Cancel";
      confirmBtn.disabled = false;
      confirmBtn.classList.remove("hidden");

      if (opts.kind === "whisper") {
        titleEl.textContent = "Download transcription model?";
        var size = opts.sizeMb ? " (~" + _trFormatModelSize(opts.sizeMb) + ")" : "";
        if (opts.prewarm) {
          msgEl.textContent = 'The "' + opts.model + '" transcription model' + size +
            " isn't downloaded yet. Download it now so transcription is ready to start? It will be stored locally.";
        } else {
          msgEl.textContent = 'The "' + opts.model + '" transcription model' + size +
            " isn't downloaded yet. It will be downloaded and stored locally before transcription begins.";
        }
        confirmBtn.textContent = "Download";
      } else {
        titleEl.textContent = "Install AI model?";
        msgEl.textContent = 'The Ollama model "' + opts.model + '" used by the ' +
          (opts.agentKey || "analysis") +
          " agent isn't installed. Install it now? This downloads the model locally and may take several minutes.";
        confirmBtn.textContent = "Install";
      }

      function cleanup() {
        cancelBtn.removeEventListener("click", onCancel);
        confirmBtn.removeEventListener("click", onConfirm);
        modal.removeEventListener("click", onBackdrop);
        document.removeEventListener("keydown", onKey);
      }
      function close(result) {
        cancelled = true; // stop any in-flight pull poll and its late callbacks
        cleanup();
        modal.classList.add("hidden");
        resolve(result);
      }
      function onCancel() { close(false); }
      function onBackdrop(e) { if (e.target === modal) close(false); }
      function onKey(e) { if (e.key === "Escape") close(false); }
      function onConfirm() {
        if (opts.kind === "whisper") { close(true); return; }
        // Ollama: kick off the pull and show progress in place. Cancel stays
        // enabled so the user can dismiss while it runs.
        confirmBtn.classList.add("hidden");
        progress.classList.remove("hidden");
        progressText.textContent = "Starting…";
        installOllamaModel(opts.model, function (st) {
          if (st.total > 0) {
            var pct = Math.max(0, Math.min(100, Math.round((st.completed / st.total) * 100)));
            barFill.style.width = pct + "%";
            progressText.textContent = (st.status || "Downloading") + " — " + pct + "%";
          } else {
            progressText.textContent = st.status || "Working…";
          }
        }, function () { return cancelled; }).then(function (ok) {
          if (cancelled) return; // dialog dismissed mid-pull — no toast, no re-close
          if (ok) {
            _trModelsCache = null;
            _trModelsCachePromise = null;
            showToast("Model installed");
            close(true);
          } else {
            // Leave the dialog open so the user can read the failure and
            // dismiss it; Cancel now resolves false.
            progressText.textContent = "Installation failed. Check that Ollama is running.";
            cancelBtn.textContent = "Close";
          }
        });
      }

      cancelBtn.addEventListener("click", onCancel);
      confirmBtn.addEventListener("click", onConfirm);
      modal.addEventListener("click", onBackdrop);
      document.addEventListener("keydown", onKey);
      modal.classList.remove("hidden");
    });
  }

  // Gate an agent run on its Ollama model being installed. Resolves true to
  // proceed, false to abort. We only block when Ollama is reachable AND the
  // model is positively missing — otherwise the existing "model unavailable"
  // error path handles it.
  function ensureAgentModelInstalled(agentKey) {
    return _trFetchModels().then(function (data) {
      var oll = data && data.ollama;
      if (!oll || !oll.available) return true;
      var agents = oll.agents || [];
      var info = null;
      for (var i = 0; i < agents.length; i++) {
        if (agents[i].key === agentKey) { info = agents[i]; break; }
      }
      if (!info || info.installed || !info.model) return true;
      return confirmModelInstall({ kind: "ollama", agentKey: agentKey, model: info.model });
    }).catch(function () { return true; });
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

  // A changed transcription model invalidates the prewarm guards: the new
  // model may be uncached and must get its own download confirmation, so we
  // clear the "already posted/declined" state and let prewarm re-offer it.
  // Clearing the declined model also means re-selecting a previously-declined
  // model gets a fresh prompt (the user explicitly chose it again).
  function _onTranscribeModelMaybeChanged(newModel) {
    if (newModel === undefined || newModel === null) return;
    if (newModel !== _lastTranscribeModel) {
      _transcriptionWarmupPosted = false;
      _prewarmDownloadPrompting = false;
      _prewarmDeclinedModel = null;
    }
    _lastTranscribeModel = newModel;
  }

  function _settingValueFromRecords(settings, name) {
    if (!settings) return undefined;
    for (var i = 0; i < settings.length; i++) {
      if (settings[i].name === name) return settings[i].value;
    }
    return undefined;
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
          _onTranscribeModelMaybeChanged(
            (applied && applied.TRANSCRIBE_MODEL) !== undefined
              ? applied.TRANSCRIBE_MODEL
              : _settingValueFromRecords(settings, "TRANSCRIBE_MODEL")
          );
          _applySettingsSnapshot(applied, settings);
        },
        onReset: function (scope, settings) {
          _trModelsCache = null;
          _trModelsCachePromise = null;
          _onTranscribeModelMaybeChanged(
            _settingValueFromRecords(settings, "TRANSCRIBE_MODEL")
          );
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
    initPanelTabs();
    initSummaryActions();
    initFriction();
    initFrictionHeatmapToggle();
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
        _stopFrictionPoll();
        _txEtaTicker.stop();
      } else {
        pollTaskStatus();
        startXrefPolling();
        if (_anyTxEtaActive()) _txEtaTicker.ensure();
        // Re-check summary + citations on tab refocus. Background-running
        // Ollama agents finish without notifying the frontend; if the citations
        // poll already gave up (or summary completed after we stopped polling)
        // the manifest result would only surface on a full page reload. This
        // catches the common "user goes to another tab, comes back" case.
        // loadSummary re-arms the summary/citations polls if generation is
        // still in flight, and the transcribe-status poll re-arms the
        // model-hint poll on the next active task transition. loadFriction does
        // the same for the friction pass.
        if (state.selectedParticipant) {
          loadSummary(state.selectedParticipant);
          loadFriction(state.selectedParticipant);
        }
      }
    });

    // Window focus is a separate signal from tab visibility: switching to
    // another window/app (Cmd-Tab) leaves the tab "visible" (document.hidden
    // stays false, so the visibilitychange handler above never fires) yet
    // browsers — Safari most aggressively — still pause/throttle setInterval for
    // the unfocused window. That freezes the streaming "Transcribing… X%"
    // progress until the next throttled tick or a manual reload. Re-poll
    // immediately on focus so the transcription progress resyncs without a
    // reload (the per-agent summary/friction pollers self-resume on focus).
    window.addEventListener("focus", function () {
      if (document.hidden) return;
      pollTaskStatus();
      if (_anyTxEtaActive()) _txEtaTicker.ensure();
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
