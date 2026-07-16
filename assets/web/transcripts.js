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
    // Lazily-built cache of .segment-row elements (invalidated on re-render);
    // shared by the segments editor (hub) and the video satellite's playhead
    // highlight / marker-click. frictionTooltipShown coordinates the shared
    // #trTooltip between the timeline-canvas hover (video) and the hot-segment
    // friction tooltip (agents), so neither clobbers the other.
    cachedSegmentRows: null,
    frictionTooltipShown: false,
    // Bumped on every participant switch so per-participant fetches that resolve
    // late (loadTranscript, loadSummary, summary/citations/friction polls — now
    // in the agents satellite) can detect they're stale and bail before
    // clobbering the active participant's UI.
    participantReqVer: 0,
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

  // ---- Helpers (showToast: 2500ms hide; shared default in utils.js is 3000ms) ----

  var _utilsShowToast = window.showToast;
  function showToast(msg) {
    _utilsShowToast(msg, { durationMs: 2500 });
  }

  // ---- Nav links ----

  function checkNavLinks() {
    apiGet("../api/status").then(function (data) {
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

  // A genuinely unreachable sibling tool (network error or non-2xx) flips a
  // per-source flag so the status tooltip can say cross-references are missing,
  // rather than the badges silently never appearing. A successful poll clears it.
  function _markXrefSource(source, failed) {
    var prev = !!(state.xrefErrors && (state.xrefErrors.screenspace || state.xrefErrors.studio));
    if (!state.xrefErrors) state.xrefErrors = { screenspace: false, studio: false };
    state.xrefErrors[source] = failed;
    var now = state.xrefErrors.screenspace || state.xrefErrors.studio;
    if (now && !prev && !state.xrefErrorToastShown) {
      state.xrefErrorToastShown = true;
      showToast("Cross-references unavailable. Is Screenspace/Studio running?");
    }
    if (!now) state.xrefErrorToastShown = false;
    updateStatusIndicator();
  }

  function loadCrossRefData() {
    fetch("../screenspace/api/events?excluded=false")
      .then(function (r) {
        if (!r.ok) throw new Error("HTTP " + r.status);
        return r.json();
      })
      .then(function (data) {
        _markXrefSource("screenspace", false);
        if (data.ok) {
          state.ssEvents = data.events || [];
          state.ssEventsLoaded = true;
          _buildEventsIndex();
        }
      })
      .catch(function () { _markXrefSource("screenspace", true); });

    fetch("../studio/api/sheet")
      .then(function (r) {
        if (!r.ok) throw new Error("HTTP " + r.status);
        return r.json();
      })
      .then(function (data) {
        _markXrefSource("studio", false);
        if (data.ok) {
          clipgenApplyConfig(data.config);
          state.sheetRows = data.rows || [];
          state.sheetParticipants = data.participants || [];
          state.sheetLoaded = true;
          _buildSheetIndex();
        }
      })
      .catch(function () { _markXrefSource("studio", true); });
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
    if (state.xrefErrors && (state.xrefErrors.screenspace || state.xrefErrors.studio)) {
      var down = [];
      if (state.xrefErrors.screenspace) down.push("Screenspace");
      if (state.xrefErrors.studio) down.push("Studio");
      lines.push("Cross-references unavailable (" + down.join(", ") + ")");
    }
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

  // One-shot #P07 deep-link guard: the hash must not re-hijack selection on
  // later participant-list refreshes after the user has moved on.
  var _hashPidApplied = false;

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

      // Deep link (#P07, from the Overview Map's explain panel) wins once,
      // on the first load that has the participant list.
      var hashPid = _hashPidApplied ? "" : clipgenHashParticipant();
      if (hashPid && state.participants.length) {
        _hashPidApplied = true;
        for (var h = 0; h < state.participants.length; h++) {
          if (state.participants[h].id === hashPid) {
            selectParticipant(hashPid);
            return;
          }
        }
      }

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
    state.participantReqVer++;
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
    cancelPendingSeek();

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
    } else if (taskForPid && taskForPid.status === "running" && taskForPid.partial_count > 0) {
      state.streamingParticipant = pid;
      clearAnalysisPanel();
      _syncStreamSegs(taskForPid, function (segs) {
        if (segs.length > 0) renderPartialSegments(segs, taskForPid.progress);
      });
    } else {
      state.segments = [];
      state.streamingParticipant = null;
      renderSegments();
      renderTimeline();
      clearAnalysisPanel();
    }

    // Reflect the newly-selected participant's transcription progress on the
    // timeline immediately (draws the band if it's mid-transcription, clears it
    // otherwise) rather than waiting up to a full poll interval.
    updateTranscribeFill();
  }

  function renderEmptyState() {
    qs("#videoPlayer").classList.add("hidden");
    qs("#videoEmpty").classList.remove("hidden");
    qs("#segmentList").innerHTML = "";
    qs("#transcriptEmpty").classList.remove("hidden");
    clearAnalysisPanel();
    clearTimelineMarkers();
    updateTranscribeFill(); // clear any lingering transcribe band
    renderTimeline();
  }

  // ---- Transcript loading ----

  function loadTranscript(pid) {
    var ver = state.participantReqVer;
    return apiGet("api/transcript/" + pid).then(function (data) {
      if (ver !== state.participantReqVer) return;
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

  // ---- Analysis panel: summary + citations + friction (impl in transcripts-agents.js) ----
  // Thin hub delegators forward to the satellite. selectParticipant + the poller +
  // the visibility/focus handlers drive load*/clear*/stop*/_restoreActiveTab; boot
  // wires the init*; the segment-list hover calls _show/_hideFrictionTooltip; the
  // poller's re-arm asks _currentParticipant/_frictionDepMet.
  function loadSummary() { return TS.loadSummary && TS.loadSummary.apply(null, arguments); }
  function loadFriction() { return TS.loadFriction && TS.loadFriction.apply(null, arguments); }
  function clearAnalysisPanel() { return TS.clearAnalysisPanel && TS.clearAnalysisPanel(); }
  function _setAnalysisPanelVisible() { return TS._setAnalysisPanelVisible && TS._setAnalysisPanelVisible.apply(null, arguments); }
  function _restoreActiveTab() { return TS._restoreActiveTab && TS._restoreActiveTab.apply(null, arguments); }
  function initPanelTabs() { return TS.initPanelTabs && TS.initPanelTabs(); }
  function initSummaryActions() { return TS.initSummaryActions && TS.initSummaryActions(); }
  function initFriction() { return TS.initFriction && TS.initFriction(); }
  function initFrictionHeatmapToggle() { return TS.initFrictionHeatmapToggle && TS.initFrictionHeatmapToggle(); }
  function _stopSummaryPoll() { return TS._stopSummaryPoll && TS._stopSummaryPoll(); }
  function _stopCitationsPoll() { return TS._stopCitationsPoll && TS._stopCitationsPoll(); }
  function _stopFrictionPoll() { return TS._stopFrictionPoll && TS._stopFrictionPoll(); }
  function _currentParticipant() { return TS._currentParticipant && TS._currentParticipant(); }
  function _frictionDepMet() { return TS._frictionDepMet && TS._frictionDepMet(); }
  function _showFrictionTooltip() { return TS._showFrictionTooltip && TS._showFrictionTooltip.apply(null, arguments); }
  function _hideFrictionTooltip() { return TS._hideFrictionTooltip && TS._hideFrictionTooltip(); }

  // ---- Segment rendering ----

  function renderSegments() {
    var container = qs("#segmentList");
    var empty = qs("#transcriptEmpty");
    state.editingTextEl = null;
    state.cachedSegmentRows = null;
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
      var sevDotHtml = "";
      if (markObj && markObj.severity) {
        sevDotHtml = '<span class="segment-sev-dot ' + severityClass(markObj.severity) + '" title="' + escapeHtml(markObj.severity) + '"></span>';
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
      html += sevDotHtml;
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

  // Client-side accumulator for streaming partial segments. The status poll only
  // carries partial_count now (not the full array), so we fetch the tail beyond
  // our cursor and append. partial_segments is append-only server-side, so the
  // count is a safe cursor and the request payload stays flat as the transcript
  // grows. Resets when the streamed task changes.
  var _streamSegs = { taskId: null, segments: [], fetching: false };

  function _syncStreamSegs(task, cb) {
    if (_streamSegs.taskId !== task.id) {
      _streamSegs.taskId = task.id;
      _streamSegs.segments = [];
      _streamSegs.fetching = false;
    }
    var count = task.partial_count || 0;
    var have = _streamSegs.segments.length;
    if (count <= have || _streamSegs.fetching) {
      cb(_streamSegs.segments);
      return;
    }
    _streamSegs.fetching = true;
    var reqTaskId = task.id;
    var since = have;
    apiGet(
      "api/transcribe/" + encodeURIComponent(task.id) + "/segments?since=" + since
    ).then(
      function (data) {
        _streamSegs.fetching = false;
        // Reject if the streamed task changed or our cursor moved under us.
        if (!data.ok || _streamSegs.taskId !== reqTaskId) return;
        if (since === _streamSegs.segments.length && data.segments && data.segments.length) {
          _streamSegs.segments = _streamSegs.segments.concat(data.segments);
        }
        cb(_streamSegs.segments);
      },
      function () {
        _streamSegs.fetching = false;
      }
    );
  }

  function _renderPartialSegmentRow(seg, i, pid) {
    var segId = pid + ":" + i;
    var cachedMark = _streamingMarks[segId];
    var cachedColor = cachedMark ? cachedMark.color : null;
    var markClass = "segment-mark" + (cachedColor ? " marked" : "");
    var markStyle = cachedColor ? ' style="background:' + cachedColor + '"' : "";
    var sevDotHtml = "";
    if (cachedMark && cachedMark.severity) {
      sevDotHtml = '<span class="segment-sev-dot ' + severityClass(cachedMark.severity) + '" title="' + escapeHtml(cachedMark.severity) + '"></span>';
    }
    var html = '<div class="segment-row segment-streaming" data-index="' + i + '" data-start="' + seg.start + '">';
    html += '<span class="' + markClass + '" data-segment-id="' + escapeHtml(segId) + '"' + markStyle + '></span>';
    html += sevDotHtml;
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
    state.cachedSegmentRows = null;

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
  // Coalesces the segment-list mousemove (hot-segment friction tooltip) the same
  // way the timeline canvas does. Lives here with its only user — the delegation
  // handler below — even though the tooltip render itself is in the agents
  // satellite (reached via TS._show/_hideFrictionTooltip).
  var _segTooltipRaf = 0;

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
  // Each entry: { color, id, category, label, severity }. `version` is bumped on
  // any write to invalidate renderPartialSegments' append-only fast path.
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
          severity: m.severity || "",
        };
        added = true;
      });
      if (added) _bumpStreamingMarksVersion();
    });
  }

  // getMarkForSegment resolves a segment's active mark — the persisted
  // seg.marks[0], falling back to the streaming-marks cache. Kept in the hub
  // (reads _streamingMarks); the video satellite reads it via TS.getMarkForSegment
  // for timeline marker rendering + tooltips.
  function getMarkForSegment(seg) {
    if (seg.marks && seg.marks.length > 0) return seg.marks[0];
    var streaming = _streamingMarks[seg.id];
    return streaming || null;
  }

  // ---- Video player + timeline (impl in transcripts-video.js) ----
  // Thin hub delegators forward to the satellite; selectParticipant, the segment
  // list, renderEmptyState/loadTranscript, the agents panel, search, and boot keep
  // calling these bare names. cancelPendingSeek / clearTimelineMarkers /
  // hasTimelineHover encapsulate the video-internal state the hub used to poke.
  function initVideoPlayer() { return TS.initVideoPlayer && TS.initVideoPlayer(); }
  function initTimelineCanvas() { return TS.initTimelineCanvas && TS.initTimelineCanvas(); }
  function initPipScroll() { return TS.initPipScroll && TS.initPipScroll(); }
  function initVideoSync() { return TS.initVideoSync && TS.initVideoSync(); }
  function initPlayerKeyboard() { return TS.initPlayerKeyboard && TS.initPlayerKeyboard(); }
  function renderTimeline() { return TS.renderTimeline && TS.renderTimeline(); }
  function updateTranscribeFill() { return TS.updateTranscribeFill && TS.updateTranscribeFill(); }
  function seekVideo() { return TS.seekVideo && TS.seekVideo.apply(null, arguments); }
  function scrollToSegment() { return TS.scrollToSegment && TS.scrollToSegment.apply(null, arguments); }
  function applyCaptionMode() { return TS.applyCaptionMode && TS.applyCaptionMode(); }
  function _partForGlobal() { return TS._partForGlobal && TS._partForGlobal.apply(null, arguments); }
  function _partMediaUrl() { return TS._partMediaUrl && TS._partMediaUrl.apply(null, arguments); }
  function cancelPendingSeek() { return TS.cancelPendingSeek && TS.cancelPendingSeek(); }
  function clearTimelineMarkers() { return TS.clearTimelineMarkers && TS.clearTimelineMarkers(); }

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

  // Find the loaded (non-streaming) segment index for an id; -1 if absent.
  function _segmentIndexById(segmentId) {
    for (var i = 0; i < state.segments.length; i++) {
      if (state.segments[i].id === segmentId) return i;
    }
    return -1;
  }

  // Find the loaded segment carrying a mark id; null if absent.
  function _findSegmentByMarkId(markId) {
    for (var i = 0; i < state.segments.length; i++) {
      var marks = state.segments[i].marks || [];
      for (var j = 0; j < marks.length; j++) {
        if (marks[j].id === markId) return { idx: i, seg: state.segments[i], mark: marks[j] };
      }
    }
    return null;
  }

  // Repaint a single segment's mark dot + annotation badge in place (mirrors the
  // template in renderSegments) so a mark add/remove/recolor shows instantly
  // without a full loadTranscript round-trip. markObj null clears the mark.
  function _paintSegmentMark(idx, markObj) {
    var list = qs("#segmentList");
    if (!list) return;
    var row = list.querySelector('.segment-row[data-index="' + idx + '"]');
    if (!row) return;
    var dot = row.querySelector(".segment-mark");
    var textEl = row.querySelector(".segment-text");
    var oldBadge = textEl ? textEl.querySelector(".segment-anno-badge") : null;
    if (oldBadge) oldBadge.remove();
    var oldSevDot = row.querySelector(".segment-sev-dot");
    if (oldSevDot) oldSevDot.remove();
    if (!dot) return;
    if (markObj) {
      var cat = MARK_CATEGORIES[markObj.category] || MARK_CATEGORIES.bookmark;
      dot.classList.add("marked");
      dot.style.background = cat.color;
      if (markObj.severity) {
        var sevDot = document.createElement("span");
        sevDot.className = "segment-sev-dot " + severityClass(markObj.severity);
        sevDot.title = markObj.severity;
        dot.insertAdjacentElement("afterend", sevDot);
      }
      if (markObj.label) {
        dot.title = markObj.label;
        if (textEl) {
          var bgMix = "color-mix(in oklch, " + cat.color + " 18%, transparent)";
          var borderMix = "color-mix(in oklch, " + cat.color + " 50%, transparent)";
          var badge = document.createElement("span");
          badge.className = "segment-anno-badge";
          badge.style.cssText =
            "--anno-badge-fg:" + cat.color + ";--anno-badge-bg:" + bgMix + ";--anno-badge-border:" + borderMix;
          badge.textContent = markObj.label;
          textEl.insertBefore(badge, textEl.firstChild);
        }
      } else {
        dot.removeAttribute("title");
      }
    } else {
      dot.classList.remove("marked");
      dot.style.background = "";
      dot.removeAttribute("title");
    }
  }

  function toggleMark(segmentId) {
    var idx = _segmentIndexById(segmentId);
    var seg = idx >= 0 ? state.segments[idx] : null;
    if (!seg) {
      // No loaded row (shouldn't happen for the persisted path) — fall back to a
      // reload so the new mark still appears.
      apiPost("api/marks", { segment_ids: [segmentId], category: state.lastMarkCategory }).then(function (data) {
        if (data.ok && state.selectedParticipant) loadTranscript(state.selectedParticipant);
      });
      return;
    }
    // Optimistic: fill the dot now, reconcile the real id on success, revert on failure.
    var prevMarks = seg.marks || [];
    var provisional = { id: null, category: state.lastMarkCategory, label: "" };
    seg.marks = [provisional];
    _paintSegmentMark(idx, provisional);
    function revert() {
      seg.marks = prevMarks;
      _paintSegmentMark(idx, prevMarks[0] || null);
      showToast("Failed to mark");
    }
    apiPost("api/marks", { segment_ids: [segmentId], category: state.lastMarkCategory })
      .then(function (data) {
        if (data.ok && data.marks && data.marks.length > 0) {
          seg.marks = [data.marks[0]];
          showToast("Marked");
        } else {
          revert();
        }
      })
      .catch(revert);
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
          severity: m.severity || "",
        };
        _bumpStreamingMarksVersion();
      }
    });
  }

  function removeMark(markId) {
    hideMarkPopover();
    // Streaming participant: keep the existing reload-on-success path.
    if (state.streamingParticipant) {
      apiDelete("api/marks/" + markId).then(function (data) {
        if (data.ok) {
          showToast("Mark removed");
          for (var key in _streamingMarks) {
            if (_streamingMarks[key].id === markId) {
              delete _streamingMarks[key];
              _bumpStreamingMarksVersion();
              break;
            }
          }
          pollTaskStatus();
        }
      });
      return;
    }
    // Optimistic: clear the dot now, restore on failure.
    var found = _findSegmentByMarkId(markId);
    if (!found) {
      apiDelete("api/marks/" + markId).then(function (data) {
        if (data.ok && state.selectedParticipant) loadTranscript(state.selectedParticipant);
      });
      return;
    }
    var prevMarks = found.seg.marks;
    found.seg.marks = [];
    _paintSegmentMark(found.idx, null);
    apiDelete("api/marks/" + markId)
      .then(function (data) {
        if (data.ok) {
          showToast("Mark removed");
        } else {
          found.seg.marks = prevMarks;
          _paintSegmentMark(found.idx, prevMarks[0] || null);
          showToast("Failed to remove mark");
        }
      })
      .catch(function () {
        found.seg.marks = prevMarks;
        _paintSegmentMark(found.idx, prevMarks[0] || null);
        showToast("Failed to remove mark");
      });
  }

  function updateMarkCategory(markId, category) {
    state.lastMarkCategory = category;
    hideMarkPopover();
    // Streaming participant: keep the existing reload-on-success path.
    if (state.streamingParticipant) {
      apiPut("api/marks/" + markId, { category: category }).then(function (data) {
        if (data.ok) {
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
        }
      });
      return;
    }
    // Optimistic: recolor the dot now, restore on failure.
    var found = _findSegmentByMarkId(markId);
    if (!found) {
      apiPut("api/marks/" + markId, { category: category }).then(function (data) {
        if (data.ok && state.selectedParticipant) loadTranscript(state.selectedParticipant);
      });
      return;
    }
    var prevCategory = found.mark.category;
    found.mark.category = category;
    _paintSegmentMark(found.idx, found.mark);
    apiPut("api/marks/" + markId, { category: category }).catch(function () {
      found.mark.category = prevCategory;
      _paintSegmentMark(found.idx, found.mark);
      showToast("Failed to update mark");
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

  function updateMarkSeverity(markId, severity) {
    var sev = severity || null;
    // Streaming participant: mirror updateMarkCategory — update the cache and
    // re-render on confirmed success. Severity changes the visible segment dot,
    // so a version bump alone isn't enough; pollTaskStatus re-renders the list
    // now instead of leaving a stale dot until the next scheduled poll.
    if (state.streamingParticipant) {
      apiPut("api/marks/" + markId, { severity: sev }).then(function (data) {
        if (data.ok) {
          for (var key in _streamingMarks) {
            if (_streamingMarks[key].id === markId) {
              _streamingMarks[key].severity = severity || "";
              _bumpStreamingMarksVersion();
              break;
            }
          }
          pollTaskStatus();
        }
      });
      return;
    }
    // Loaded participant: optimistically repaint the dot, restore on failure
    // (mirrors updateMarkCategory).
    var found = _findSegmentByMarkId(markId);
    if (!found) {
      apiPut("api/marks/" + markId, { severity: sev }).then(function (data) {
        if (data.ok && state.selectedParticipant) loadTranscript(state.selectedParticipant);
      });
      return;
    }
    var prevSeverity = found.mark.severity;
    found.mark.severity = sev;
    _paintSegmentMark(found.idx, found.mark);
    apiPut("api/marks/" + markId, { severity: sev }).catch(function () {
      found.mark.severity = prevSeverity;
      _paintSegmentMark(found.idx, found.mark);
      showToast("Failed to update mark");
    });
  }

  function showMarkPopover(anchorEl, segmentId, markObj) {
    var popover = qs("#markPopover");
    hideMarkPopover();

    // Build category pills
    var catContainer = popover.querySelector(".mark-popover-categories");
    catContainer.innerHTML = "";
    var cats = Object.keys(MARK_CATEGORIES);
    var pills = [];
    var activeIdx = 0;
    for (var i = 0; i < cats.length; i++) {
      (function (key, idx) {
        var cat = MARK_CATEGORIES[key];
        var pill = document.createElement("button");
        pill.type = "button";
        pill.className = "mark-cat-pill" + (markObj.category === key ? " active" : "");
        if (markObj.category === key) activeIdx = idx;
        pill.style.background = cat.color;
        pill.title = cat.label;
        pill.setAttribute("aria-label", cat.label);
        pill.addEventListener("click", function (e) {
          e.stopPropagation();
          updateMarkCategory(markObj.id, key);
        });
        pills.push(pill);
        catContainer.appendChild(pill);
      })(cats[i], i);
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

    // Severity dropdown (optional). Options come from CLIPGEN_CONFIG.severity
    // (mirrored from config.py SEVERITY_NUMERIC_TO_LABEL); a blank first option
    // means "no severity".
    var sevSelect = popover.querySelector(".mark-popover-severity");
    if (!sevSelect.options.length) {
      var blankOpt = document.createElement("option");
      blankOpt.value = "";
      blankOpt.textContent = "No severity";
      sevSelect.appendChild(blankOpt);
      var levels = CLIPGEN_CONFIG.severity || [];
      for (var s = 0; s < levels.length; s++) {
        var opt = document.createElement("option");
        opt.value = levels[s].label;
        opt.textContent = levels[s].label;
        sevSelect.appendChild(opt);
      }
    }
    sevSelect.value = markObj.severity || "";
    sevSelect.onchange = function () {
      // updateMarkSeverity owns the mark-state mutation (like updateMarkCategory)
      // so it can capture the previous value and roll back on failure. markObj is
      // the live seg.marks[0]/streaming-cache ref, so it stays in sync.
      updateMarkSeverity(markObj.id, sevSelect.value);
    };

    // Remove button
    var removeBtn = popover.querySelector(".mark-popover-remove");
    removeBtn.onclick = function (e) {
      e.stopPropagation();
      removeMark(markObj.id);
    };

    // Keyboard: arrows rove between category pills, Enter applies the focused
    // category (the pill's own click), Esc dismisses. Typing in the label input
    // keeps its own Enter/Esc handling and is skipped for arrow roving.
    popover.onkeydown = function (e) {
      if (e.key === "Escape") { e.preventDefault(); hideMarkPopover(); return; }
      if (document.activeElement === labelInput || document.activeElement === sevSelect) return;
      var idx = pills.indexOf(document.activeElement);
      if (e.key === "ArrowRight" || e.key === "ArrowDown") {
        e.preventDefault();
        pills[idx < 0 ? 0 : (idx + 1) % pills.length].focus();
      } else if (e.key === "ArrowLeft" || e.key === "ArrowUp") {
        e.preventDefault();
        pills[idx < 0 ? pills.length - 1 : (idx - 1 + pills.length) % pills.length].focus();
      }
    };

    // Position below anchor
    var rect = anchorEl.getBoundingClientRect();
    popover.style.top = (rect.bottom + window.scrollY + 4) + "px";
    popover.style.left = (rect.left + window.scrollX - 4) + "px";
    popover.classList.remove("hidden");

    // Focus the current category so arrows/Enter work immediately. preventScroll
    // keeps the segment list from jumping when the popover opens via the keyboard.
    if (pills.length) {
      try { pills[activeIdx].focus({ preventScroll: true }); }
      catch (err) { pills[activeIdx].focus(); }
    }

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

  // ---- Search (impl in transcripts-search.js) ----
  // Thin hub delegator: boot wires initSearch. Search renders its own results
  // internally, so no renderSearchResults delegator is needed.
  function initSearch() {
    return TS.initSearch && TS.initSearch();
  }

  // ---- Participant pills (impl in transcripts-pills.js) ----
  // Thin hub delegators: loadParticipants / selectParticipant / pollTaskStatus
  // call renderPills; boot wires initPillOutsideClick / initPillWheelScroll.
  function renderPills() { return TS.renderPills && TS.renderPills(); }
  function initPillOutsideClick() { return TS.initPillOutsideClick && TS.initPillOutsideClick(); }
  function initPillWheelScroll() { return TS.initPillWheelScroll && TS.initPillWheelScroll(); }

  // _confirmUncachedWhisperModels stays in the hub (model-install state:
  // _trModelsCache / _trModelsCachePromise / _whisperDownloadConfirmed +
  // confirmModelInstall). The pills satellite's _postTranscribe reaches it via
  // TS._confirmUncachedWhisperModels.

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
    // _summaryPoller lives in the agents satellite; ask it whether the summary
    // poll is already armed (TS.isSummaryPolling) rather than reading the var.
    if (summaryRunning && !(TS.isSummaryPolling && TS.isSummaryPolling()) && !state.summaryText) loadSummary(pid);
    // Reload friction when the panel is empty, or when the deterministic-only
    // placeholder is showing and the summary-gated agent has now started/finished
    // — otherwise the deterministic blob keeps this guard shut and
    // the AI moments would never replace it (see agents/CODE-REVIEW.md poll gates).
    var frictionActive = !!(p && p.agents && (p.agents.friction === "running" || p.agents.friction === "done"));
    var showingDeterministic = !!(state.frictionData && state.frictionData.deterministic);
    if (!state.frictionGenerating && _frictionDepMet() &&
        (!state.frictionData || (showingDeterministic && frictionActive))) {
      loadFriction(pid);
    }
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
    var ver = state.participantReqVer;
    apiGet("api/transcript/" + pid).then(function (data) {
      if (ver !== state.participantReqVer || state.selectedParticipant !== pid) return;
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

      // Fill the timeline in sync with the selected participant's transcription
      // progress (starts/updates/clears the band purely from state.tasks).
      updateTranscribeFill();

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
          if (t.participant === state.selectedParticipant && t.status === "running" && t.partial_count) {
            selectedRunningTask = t;
          }
        });
      }
      if (selectedRunningTask) {
        state.streamingParticipant = state.selectedParticipant;
        _loadStreamingMarks(state.selectedParticipant);
        _syncStreamSegs(selectedRunningTask, function (segs) {
          if (segs.length > 0) renderPartialSegments(segs, selectedRunningTask.progress);
        });
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

  // ---- Corrections modal (impl in transcripts-corrections.js) ----
  // Thin hub delegators forward to the satellite; the inline-edit
  // saveCorrections() flow (loadCorrections) and boot (initCorrectionsModal)
  // keep calling these bare names.
  function initCorrectionsModal() {
    return TS.initCorrectionsModal && TS.initCorrectionsModal();
  }
  function loadCorrections() {
    return TS.loadCorrections && TS.loadCorrections();
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
    _trModelsCachePromise = apiGet("/api/models")
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
        closeBlockingModal(modal);
      }
      function close(result) {
        cancelled = true; // stop any in-flight pull poll and its late callbacks
        cleanup();
        modal.classList.add("hidden");
        resolve(result);
      }
      function onCancel() { close(false); }
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
            progressText.textContent = (st.status || "Downloading") + ": " + pct + "%";
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
      modal.classList.remove("hidden");
      // Escape and backdrop click both cancel (no focus trap — matches prior
      // behavior for this lightweight progress dialog).
      openBlockingModal(modal, { onEscape: onCancel, onBackdropClick: onCancel });
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

  // Command palette (command-palette.js): additions beyond the auto-ingested
  // quick actions — search focus, cheatsheet, panel tabs, participant jumps
  // (the provider runs on every palette open, so it tracks state.participants).
  function initCommandPalette() {
    if (!window.ClipgenCommandPalette) return;
    function clickCommand(id, title, icon, keywords, elId) {
      return {
        id: id,
        title: title,
        icon: icon,
        keywords: keywords,
        section: "Transcripts",
        visible: function () { return !!document.getElementById(elId); },
        run: function () { document.getElementById(elId).click(); },
      };
    }
    window.ClipgenCommandPalette.setParticipants(function () {
      return (state.participants || []).map(function (p) { return p.id; });
    });
    window.ClipgenCommandPalette.register("transcripts", function () {
      var cmds = [
        {
          id: "transcripts.search",
          title: "Focus transcript search",
          icon: "magnifying-glass",
          keywords: "find text query",
          section: "Transcripts",
          visible: function () { return !!document.getElementById("searchInput"); },
          // Runs after the palette closes and restores focus, so this focus
          // call wins.
          run: function () { document.getElementById("searchInput").focus(); },
        },
        clickCommand("transcripts.shortcuts", "Keyboard shortcuts", "command-line",
          "cheatsheet keys help", "shortcutsBtn"),
        clickCommand("transcripts.tab-summary", "Show Summary tab", "table-cells",
          "panel agents", "tabBtnSummary"),
        clickCommand("transcripts.tab-friction", "Show Friction tab", "table-cells",
          "panel analysis moments", "tabBtnFriction"),
      ];
      // "Jump to … in Transcripts" = stays here and selects in place; the
      // palette's built-in provider adds the cross-page "Open … in <Page>".
      (state.participants || []).forEach(function (p) {
        cmds.push({
          id: "transcripts.p." + p.id,
          title: "Jump to " + p.id + " in Transcripts",
          icon: "user",
          keywords: "participant select transcript",
          section: "Participants",
          run: function () { selectParticipant(p.id); },
        });
      });
      return cmds;
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
    // /transcripts/#tab=friction style deep links (command palette). The
    // participant hash form (#P07) is handled separately in loadParticipants.
    var hashTab = clipgenHashTab();
    if (hashTab === "summary" || hashTab === "friction") {
      var hashTabBtn = qs(hashTab === "friction" ? "#tabBtnFriction" : "#tabBtnSummary");
      if (hashTabBtn) hashTabBtn.click();
    }
    initSummaryActions();
    initFriction();
    initFrictionHeatmapToggle();
    initTranscriptSettings();
    initTopNavActions();
    initCommandPalette();

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

  // ---- Satellite interface (window.ClipgenTranscripts) ----
  // Published for the transcripts-*.js satellite files (corrections, search,
  // video, pills, agents) that load after this script. They read the hub's
  // shared `state` + helpers through this object and attach their own published
  // functions back onto it — mirrors screenspace.js / window.ClipgenScreenspace.
  // Assigned synchronously here (during the hub script's load) so the object is
  // fully populated before any satellite IIFE runs; the DOMContentLoaded init
  // above and all user-event handlers fire later still, by which point
  // satellites have registered their functions.
  //
  // Hub helpers the satellites call outward are published below as they are
  // needed by each carved satellite. Functions the hub calls that live in a
  // satellite are reached through thin guarded delegators (see the delegators
  // section) so the hub degrades to a no-op if a satellite fails to load.
  var TS = (window.ClipgenTranscripts = window.ClipgenTranscripts || {});
  TS.state = state;
  TS.showToast = showToast;
  // Hub helpers the satellites call outward.
  TS.loadTranscript = loadTranscript; // corrections, search, agents
  TS.findOverlapsForSearch = findOverlapsForSearch; // search
  TS.selectParticipant = selectParticipant; // search, pills
  // Accumulated streaming segments for the currently-streamed participant. The
  // status poll no longer carries partial_segments, so search reads them here.
  TS.streamingSegmentsFor = function (pid) {
    return pid && pid === state.streamingParticipant ? _streamSegs.segments : [];
  }; // search
  TS.getMarkForSegment = getMarkForSegment; // video (timeline markers + tooltip)
  TS.toggleMark = toggleMark; // video (keyboard marking)
  TS.updateMarkCategory = updateMarkCategory; // video (keyboard category set)
  TS.showMarkPopover = showMarkPopover; // video (keyboard marking on already-marked segment)
  TS.maybeWarmOnPillHover = maybeWarmOnPillHover; // pills
  TS.tryPostTranscriptionWarmup = tryPostTranscriptionWarmup; // pills
  TS.pollTaskStatus = pollTaskStatus; // pills
  TS.startPolling = startPolling; // pills
  TS._refreshAgentStateNow = _refreshAgentStateNow; // pills
  TS._trFetchModels = _trFetchModels; // pills
  TS.ensureAgentModelInstalled = ensureAgentModelInstalled; // pills, agents
  TS._confirmUncachedWhisperModels = _confirmUncachedWhisperModels; // pills (model-install kept in hub)
  // Hub helpers the agents satellite calls outward (loadFriction is owned by the
  // agents satellite now and reached through the delegator above).
  TS.renderSegments = renderSegments; // agents (heatmap toggle, friction mark-all)
  TS._txEtaTicker = _txEtaTicker; // agents (summary/citations/friction elapsed)
  TS._summaryEtaTracker = _summaryEtaTracker; // agents
  TS._citationsEtaTracker = _citationsEtaTracker; // agents
  TS._frictionEtaTracker = _frictionEtaTracker; // agents
  TS._updateAgentElapsed = _updateAgentElapsed; // agents
  TS._currentParticipantHasTranscript = _currentParticipantHasTranscript; // agents (panel-visible guard) + hub topnav

})();
