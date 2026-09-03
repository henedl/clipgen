/* clipgen Transcripts page.
 *
 * Editor for per-participant Whisper transcripts, plus cross-references into
 * Screenspace events and the source spreadsheet. Two long-running side flows
 * piggyback on the same `state`:
 *
 *   - Transcription warmup: a single `tryPostTranscriptionWarmup()` post that
 *     asks the backend to preload the Whisper model. `_transcriptionWarmupPosted`
 *     guards it so we never double-post per page load.
 *   - Summary / citations: LLM-generated; `_summaryPoller` and
 *     `_citationsPoller` (createPoller handles) poll the backend until the
 *     result lands or the user navigates away.
 */

(function () {
  "use strict";

  var state = {
    participants: [],
    // "<pid>/<agent>" -> the failure reason already toasted for that run.
    agentErrorsSeen: {},
    selectedParticipant: null,
    segments: [],
    corrections: [],
    knownTerms: [],
    tasks: [],
    // Task id -> { at, progress } for cancels in flight; server still reports "running".
    cancellingTasks: {},
    // Transcribe-range markers in global-timeline seconds; null = unset. Video satellite owns.
    inMarker: null,
    outMarker: null,
    searchQuery: "",
    searchResults: null,
    activeSegmentIndex: -1,
    editingTextEl: null,
    lastMarkCategory: "bookmark",
    streamingParticipant: null,
    ssEvents: [],
    ssEventsLoaded: false,
    sheetRows: [],
    sheetParticipants: [],
    sheetLoaded: false,
    // Whether a spreadsheet is loaded at all — gates the per-pill off-sheet badge.
    hasSheet: false,
    xrefPoller: null,
    xrefEligible: false,
    xrefIndex: { eventsByParticipant: {}, sheetByParticipant: {} },
    summaryEditing: false,
    summaryText: "",
    summaryCitations: null,
    citationsGenerating: false,
    activeTab: "summary",
    frictionData: null,
    // Owner of frictionData; loadFriction blanks the pane only when this changes.
    frictionPid: null,
    frictionBySegId: {},
    frictionGenerating: false,
    // Server-recorded run start (epoch ms); survives navigation. Null while idle.
    frictionStartedAt: null,
    // "off" | "highlight" | "isolate". Persisted; owned by transcripts-agents.js.
    frictionMode: "off",
    // Score band the filter keeps; starts fully open so the histogram shows everything.
    frictionMin: 0,
    frictionMax: 1,
    // Separate filters per evidence source (scorer vs agent); both persist.
    frictionCategoryFilter: null,
    frictionMomentFilter: null,
    // Outputs of _recomputeFrictionMatches; every friction consumer reads these.
    frictionMatchBySegId: {},
    // Union of both sources, strongest score per line; the timeline band draws this.
    frictionBandBySegId: {},
    frictionVisibleMoments: [],
    frictionCitedBySegId: {},
    // Filtered moments citing segment ids no longer in state.segments.
    frictionUnsourcedMoments: 0,
    frictionMomentIndex: -1,
    transcribePrewarm: "queue_open",
    modelStatus: null,
    modelFailSince: 0,
    videoPlaying: false,
    videoMuted: false,
    // From /api/audio-info; feeds the volume popover's track caption and mixer.
    audioTracks: [],
    audioPanel: null, // ClipgenVideoControls audio-popover controller
    videoPlaybackRate: 1,
    ccEnabled: false,
    pipActive: false,
    pipEnabled: true,
    videoCollapsed: false,
    // Lazy .segment-row cache, invalidated on re-render. frictionTooltipShown arbitrates #trTooltip.
    cachedSegmentRows: null,
    frictionTooltipShown: false,
    // Bumped per participant switch; late-resolving fetches compare and bail.
    participantReqVer: 0,
    // Keyboard cursor into the participant-options dropdown (-1 = none); see pillNav*.
    pillOptionsCursor: -1,
  };

  var _transcriptionWarmupPosted = false;
  // Prewarm confirms downloads; a declined model is not re-asked until it changes.
  var _prewarmDownloadPrompting = false;
  var _prewarmDeclinedModel = null;
  // Last-known TRANSCRIBE_MODEL; a change resets the prewarm guards.
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
    // Re-arm the sheet leg; a sheet may have opened from another tab since.
    _sheetXrefIdle = false;
    state.xrefPoller = createPoller(loadCrossRefData, 30000, { label: "transcripts.xref" });
    state.xrefPoller.start();
  }

  function stopXrefPolling() {
    if (state.xrefPoller) {
      state.xrefPoller.stop();
      state.xrefPoller = null;
    }
  }

  // Per-source unreachable flag for the status tooltip; a good poll clears it.
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

  // Set after sheet_loaded: false; stops polling the sheet until tab focus re-arms.
  var _sheetXrefIdle = false;

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

    if (_sheetXrefIdle) return;
    fetch("../studio/api/sheet")
      .then(function (r) {
        if (!r.ok) throw new Error("HTTP " + r.status);
        return r.json();
      })
      .then(function (data) {
        _markXrefSource("studio", false);
        if (data.ok && data.sheet_loaded === false) _sheetXrefIdle = true;
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
    // No baseline applied; timestamps stay relative (MM:SS for 2-part).
    var segs = parseClipSegmentsForCell(raw, 0, CLIPGEN_CONFIG.defaultDuration);
    return segs.map(function (s) {
      return { start: s.startSeconds, duration: s.duration };
    });
  }

  // Per-participant indexes sorted by start so findOverlapsForSearch can binary-search.

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

  // First index where keyFn(arr[i]) >= value.
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
      // Sorted by `in`, so lower-bound on `end` is the exclusive upper cursor.
      var upper = _lowerBound(events, end, function (e) { return e.in; });
      for (var i = 0; i < upper; i++) {
        if (events[i].out > start) result.screenspaceEvents.push(events[i].ev);
      }
    }

    var segs = state.xrefIndex.sheetByParticipant[participant];
    if (segs && segs.length > 0) {
      var upper2 = _lowerBound(segs, end, function (s) { return s.start; });
      // One observation per row, in sheet row order.
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
    // Seed once; only settings saves update it afterwards.
    if (data.model && _lastTranscribeModel === null) _lastTranscribeModel = data.model;
    // Time the failed-looking state so cold load doesn't flash red.
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

  // Every surface that words task progress asks this, so they agree.
  function _cancelPending(task) {
    return !!(task && state.cancellingTasks[task.id]);
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
    } else if (_cancelPending(task)) {
      // Still --working: the worker is winding down.
      cls = "status-indicator--working";
      taskLine = pid + ": cancelling…";
    } else if (task && task.status === "running" && task.phase === "loading_model") {
      cls = "status-indicator--working";
      taskLine = pid + ": loading transcription model\u2026";
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

    // Model failure overrides every state except an active task.
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

  // Repaint task wording now, without waiting for the poll. Keep out of _txEtaTicker.
  function refreshTranscribeWording() {
    updateStatusIndicator();
    var task = _taskForSelectedParticipant();
    _setTranscriptEmptyText(task);
    var txt = document.querySelector("#segmentList .streaming-text");
    if (txt && task) txt.textContent = _streamingTextStr(task.progress || 0);
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

  // Call once a download attempt has concluded; never mid-download (would re-prompt).
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
            // Warmup ended without loading; the next attempt re-confirms.
            stopModelHintPoll();
            _forgetWhisperDownloadAgreements();
          }
        })
        .catch(function () {});
    };
    _modelHintPoller = createPoller(poll, 1500, { label: "transcripts.modelHint" });
    _modelHintPoller.start();
  }

  function maybeWarmOnPillHover(p, s) {
    if (state.transcribePrewarm !== "queue_open") return;
    if (!s || s.status === "completed") return;
    tryPostTranscriptionWarmup();
  }

  // Preload the Whisper model once per page load; the flag resets unless a load began.
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

  // Confirm before downloading; confirm re-posts with force=true, decline stops re-asking.
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
      // Skip the transcribe-time prompt while this download is still running.
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

  // The #P07 deep link applies once; later refreshes must not re-hijack selection.
  var _hashPidApplied = false;

  // A failed boot fetch must replace the skeleton placeholders with real empty states.
  function _clearBootPlaceholders() {
    if (state.participants.length) return;
    renderPills();
    renderEmptyState();
  }

  // Toast failed AI runs for every participant, deduped per pid/agent until cleared.
  function _reportAgentErrors(participants) {
    var live = {};
    for (var i = 0; i < participants.length; i++) {
      var pid = participants[i].id;
      var errors = participants[i].agent_errors || {};
      for (var agent in errors) {
        if (!Object.prototype.hasOwnProperty.call(errors, agent)) continue;
        live[pid + "/" + agent] = true;
        reportAgentError(pid, agent, errors[agent]);
      }
    }
    var seen = state.agentErrorsSeen;
    for (var known in seen) {
      if (!Object.prototype.hasOwnProperty.call(live, known)) delete seen[known];
    }
  }

  // Shared with the per-agent panel polls; first caller wins, the rest dedupe.
  function reportAgentError(pid, agentKey, message) {
    if (!message) return;
    var key = pid + "/" + agentKey;
    if (state.agentErrorsSeen[key] === message) return;
    state.agentErrorsSeen[key] = message;
    showToast(pid + " " + agentKey + ": " + message);
  }

  function loadParticipants() {
    return apiGet("api/participants").then(function (data) {
      if (!data.ok) {
        _clearBootPlaceholders();
        return;
      }
      // Primary config channel; the xref sheet leg only refreshes it later.
      if (data.config) clipgenApplyConfig(data.config);
      state.participants = data.participants;
      _reportAgentErrors(data.participants);
      state.hasSheet = !!data.has_sheet;
      state.transcribePrewarm = data.transcribe_prewarm || "queue_open";
      renderPills();
      refreshTopNavActions();

      if (!needsTranscription()) {
        _transcriptionWarmupPosted = false;
      } else if (state.transcribePrewarm === "page_load") {
        tryPostTranscriptionWarmup();
      }
      refreshTranscriptionModelHintOnce();

      // Precedence: #P07 deep link (once) > in-memory selection > localStorage.
      var hashPid = _hashPidApplied ? "" : clipgenHashParticipant();
      if (hashPid && state.participants.length) _hashPidApplied = true;
      var pick = clipgenPickParticipant(state.participants, {
        hashPid: hashPid,
        currentId: state.selectedParticipant,
        storedId: getStoredUIState("transcripts").selectedParticipant,
      });
      if (pick) {
        if (pick !== state.selectedParticipant) selectParticipant(pick);
        return;
      }

      // Auto-select first participant with a transcript, or just the first
      var first = null;
      for (var k = 0; k < state.participants.length; k++) {
        if (state.participants[k].has_transcript) { first = state.participants[k]; break; }
      }
      if (!first && state.participants.length > 0) first = state.participants[0];
      if (first) selectParticipant(first.id);
      else renderEmptyState();
    }).catch(function () {
      _clearBootPlaceholders();
    });
  }


  // Previous/next participant, wrapping. Bound to Z / X in transcripts-video.js.
  function cycleParticipant(delta) {
    var list = state.participants;
    if (!list || list.length < 2) return;
    var idx = -1;
    for (var i = 0; i < list.length; i++) {
      if (list[i].id === state.selectedParticipant) { idx = i; break; }
    }
    if (idx === -1) idx = 0;
    var next = list[(idx + delta + list.length) % list.length];
    if (next && next.id !== state.selectedParticipant) selectParticipant(next.id);
  }

  function selectParticipant(pid) {
    state.participantReqVer++;
    _stopSummaryPoll();
    _stopCitationsPoll();
    _stopFrictionPoll();
    hideMarkPopover();
    state.selectedParticipant = pid;
    setStoredUIStateField("transcripts", "selectedParticipant", pid);
    // Restore transcribe-range markers; loadedmetadata clamps them later.
    if (TS.restoreMarkers) {
      TS.restoreMarkers(pid);
      TS.updateMarkerInfo();
      renderTimeline();
    }
    renderPills();
    refreshTopNavActions();

    // Find participant info
    var p = null;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) { p = state.participants[i]; break; }
    }
    if (!p) return;

    // Fragmented-MP4 warning + remux action, above the player.
    if (window.clipgenMediaBanner) {
      window.clipgenMediaBanner.show(qs("#videoSection"), p);
    }

    // Clean up previous video state
    var video = qs("#videoPlayer");
    var videoEmpty = qs("#videoEmpty");
    video.pause();
    cancelPendingSeek();

    // Reset the mix even without video, or orphaned <audio> elements keep playing.
    state.audioTracks = [];
    if (state.audioPanel) state.audioPanel.refresh();

    if (p.has_video) {
      // Probe audio layout for the volume popover; participantReqVer drops stale replies.
      (function (ver) {
        _trFetchAudioInfo(pid, p.video_version)
          .then(function (info) {
            if (ver !== state.participantReqVer) return;
            if (info) state.audioTracks = info.tracks;
            if (state.audioPanel) state.audioPanel.refresh();
          })
          .catch(function () {});
      })(state.participantReqVer);

      // Multi-part participants play part-by-part with client-side source switching.
      state.videoTimeline = p.timeline && p.timeline.length > 1 ? p.timeline : null;
      state.videoVersion = p.video_version != null ? p.video_version : null;
      state.videoActivePart = 0;
      state.videoOffset = 0;
      video.classList.remove("hidden");
      videoEmpty.classList.add("hidden");

      // VTT cues use global time, so multi-part participants get no native track.
      var track = qs("#subtitleTrack");
      if (state.videoTimeline) {
        track.removeAttribute("src");
      } else {
        track.src = "api/vtt/" + pid;
      }
      applyCaptionMode();

      // Restore saved global time, else 0.001s so preload="metadata" paints a frame.
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
      // Not taskForPid: the pane must match the indicator's pick among duplicate tasks.
      _setTranscriptEmptyText(_taskForSelectedParticipant());
      renderTimeline();
      clearAnalysisPanel();
    }

    // Draw or clear the transcribe band now, not at the next poll.
    updateTranscribeFill();
  }

  // #transcriptEmpty copy for the selected participant's task; null restores defaults.
  function _setTranscriptEmptyText(task) {
    var empty = qs("#transcriptEmpty");
    if (!empty) return;
    var main = empty.querySelector("p");
    var hint = empty.querySelector(".empty-hint");
    // Only in-flight states shimmer.
    var waiting = !!(task && (task.status === "running" || task.status === "queued"));
    main.classList.toggle("cg-shimmer", waiting);
    if (_cancelPending(task)) {
      main.textContent = "Cancelling…";
      // Not "finishing the current segment": a cancel may wait on the model load.
      hint.textContent = "Waiting for the transcription worker to stop";
    } else if (task && task.status === "running" && task.phase === "loading_model") {
      main.textContent = "Loading transcription model…";
      hint.textContent = "The first transcription after a restart takes a few extra seconds";
    } else if (task && task.status === "running") {
      main.textContent = "Starting transcription…";
      hint.textContent = "Lines appear here as they are transcribed";
    } else if (task && task.status === "queued") {
      main.textContent = "Queued for transcription…";
      hint.textContent = "Waiting for the current task to finish";
    } else {
      main.textContent = "No transcript available";
      hint.textContent = "Use the Queue panel to transcribe this participant's video";
    }
  }

  function renderEmptyState() {
    qs("#videoPlayer").classList.add("hidden");
    qs("#videoEmpty").classList.remove("hidden");
    qs("#segmentList").innerHTML = "";
    _setTranscriptEmptyText(null);
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

  // ---- Analysis panel delegators; implementation in transcripts-agents.js ----
  function loadSummary() { return TS.loadSummary && TS.loadSummary.apply(null, arguments); }
  function loadFriction() { return TS.loadFriction && TS.loadFriction.apply(null, arguments); }
  function clearAnalysisPanel() { return TS.clearAnalysisPanel && TS.clearAnalysisPanel(); }
  function _setAnalysisPanelVisible() { return TS._setAnalysisPanelVisible && TS._setAnalysisPanelVisible.apply(null, arguments); }
  function _restoreActiveTab() { return TS._restoreActiveTab && TS._restoreActiveTab.apply(null, arguments); }
  function initPanelTabs() { return TS.initPanelTabs && TS.initPanelTabs(); }
  function initSummaryActions() { return TS.initSummaryActions && TS.initSummaryActions(); }
  function initFriction() { return TS.initFriction && TS.initFriction(); }
  function initFrictionMode() { return TS.initFrictionMode && TS.initFrictionMode(); }
  function applyFrictionDecorations() { return TS.applyFrictionDecorations && TS.applyFrictionDecorations(); }
  function cycleFrictionMode() { return TS.cycleFrictionMode && TS.cycleFrictionMode(); }
  function _stopSummaryPoll() { return TS._stopSummaryPoll && TS._stopSummaryPoll(); }
  function _stopCitationsPoll() { return TS._stopCitationsPoll && TS._stopCitationsPoll(); }
  function _stopFrictionPoll() { return TS._stopFrictionPoll && TS._stopFrictionPoll(); }
  function _currentParticipant() { return TS._currentParticipant && TS._currentParticipant(); }
  function _frictionDepMet() { return TS._frictionDepMet && TS._frictionDepMet(); }
  function _showFrictionTooltip() { return TS._showFrictionTooltip && TS._showFrictionTooltip.apply(null, arguments); }
  function _hideFrictionTooltip() { return TS._hideFrictionTooltip && TS._hideFrictionTooltip(); }

  // ---- Segment rendering ----

  // Perf span decides whether >2000-segment lists ever need virtualizing.
  function renderSegments() {
    return clipgenPerf.span("transcripts.renderSegments", renderSegmentsImpl);
  }

  function renderSegmentsImpl() {
    var container = qs("#segmentList");
    var empty = qs("#transcriptEmpty");
    state.editingTextEl = null;
    state.cachedSegmentRows = null;
    // Drop any queued streaming indicator before the final render.
    _cancelStreamingIndicator();

    // Scroll lives on #trMain. Same-participant rebuilds keep the offset; switches reset.
    var scrollHost = qs("#trMain") || container;
    var samePid = _renderedSegmentsPid === state.selectedParticipant;
    var restoreTop = samePid ? scrollHost.scrollTop : 0;
    _renderedSegmentsPid = state.selectedParticipant;

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
      var markLabel = markObj && markObj.label ? ' data-tooltip="' + escapeHtml(markObj.label) + '"' : "";
      var annoBadgeHtml = "";
      if (markObj && markObj.label && markColor) {
        var bgMix = "color-mix(in oklch, " + markColor + " 18%, transparent)";
        var borderMix = "color-mix(in oklch, " + markColor + " 50%, transparent)";
        var badgeStyle = "--anno-badge-fg:" + markColor + ";--anno-badge-bg:" + bgMix + ";--anno-badge-border:" + borderMix;
        annoBadgeHtml = '<span class="segment-anno-badge" style="' + badgeStyle + '">' + escapeHtml(markObj.label) + '</span>';
      }
      var sevDotHtml = "";
      if (markObj && markObj.severity) {
        sevDotHtml = '<span class="segment-sev-dot ' + severityClass(markObj.severity) + '" data-tooltip="' + escapeHtml(markObj.severity) + '"></span>';
      }

      // No friction markup here; applyFrictionDecorations() owns it all.
      html += '<div class="segment-row' + activeClass + correctedClass + '" data-index="' + i + '" data-start="' + seg.start + '">';
      html += '<span class="' + markClass + '" data-segment-id="' + escapeHtml(seg.id) + '"' + markStyle + markLabel + '></span>';
      html += sevDotHtml;
      html += '<span class="segment-timestamp">' + formatTime(seg.start);
      // Cross-reference badges in gutter (inside timestamp, positioned at right edge)
      if (CLIPGEN_CONFIG.crossReferences) {
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
            html += '<span class="segment-xref-badge" style="background:' + XREF_BADGES.screenspace.color + '" data-tooltip="' + escapeHtml(evTypes.join(", ")) + '"><span class="xref-badge-icon" style="' + iconMaskStyle(XREF_BADGES.screenspace.icon) + '"></span></span>';
          }
          if (xref.sheetObservations.length > 0) {
            var obsTitle = xref.sheetObservations[0].observation;
            html += '<span class="segment-xref-badge" style="background:' + XREF_BADGES.sheet.color + '" data-tooltip="' + escapeHtml(obsTitle) + '"><span class="xref-badge-icon" style="' + iconMaskStyle(XREF_BADGES.sheet.icon) + '"></span></span>';
          }
          html += '</span>';
        }
      }
      html += '</span>';
      // Word spans carry data-ws/data-we for the karaoke sweep; count mismatch leaves them untimed.
      var tokens = seg.text.split(/(\s+)/);
      var wordCount = 0;
      for (var w = 0; w < tokens.length; w++) {
        if (tokens[w] && !/^\s+$/.test(tokens[w])) wordCount++;
      }
      var segWords = (!seg.corrected && seg.words && seg.words.length === wordCount) ? seg.words : null;
      var wordHtml = "";
      var wi = 0;
      for (var w = 0; w < tokens.length; w++) {
        if (/^\s+$/.test(tokens[w])) {
          wordHtml += tokens[w];
        } else if (tokens[w]) {
          var timing = segWords ? ' data-ws="' + segWords[wi].start + '" data-we="' + segWords[wi].end + '"' : "";
          wordHtml += '<span class="segment-word"' + timing + '>' + escapeHtml(tokens[w]) + '</span>';
          wi++;
        }
      }
      html += '<span class="segment-text" data-id="' + escapeHtml(seg.id) + '">' + annoBadgeHtml + wordHtml + '</span>';
      html += '<span class="segment-copy" data-tooltip="Copy text"><span class="segment-copy-icon"></span></span>';
      html += '</div>';
    }
    container.innerHTML = html;
    // Decorate before the scroll restore; isolate mode changes scrollHeight.
    applyFrictionDecorations();
    // Always write, even 0, or the outgoing offset survives a participant switch.
    ignoreNextScroll();
    scrollHost.scrollTop = restoreTop;

    _ensureSegmentListDelegation();
    _partialRender.count = 0;
    _partialRender.pid = null;
    _partialRender.segments = null;
    _partialRender.marksVersion = _streamingMarksVersion;
  }

  // Participant #segmentList shows; renderPartialSegments keeps it current too.
  var _renderedSegmentsPid = null;

  // Append cursor for renderPartialSegments; full rebuild on pid change, count drop, or marks change.

  var _partialRender = {
    pid: null,
    count: 0,
    segments: null,
    marksVersion: 0,
  };

  // Streaming accumulator; partial_segments is append-only server-side, so count is the cursor.
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
      sevDotHtml = '<span class="segment-sev-dot ' + severityClass(cachedMark.severity) + '" data-tooltip="' + escapeHtml(cachedMark.severity) + '"></span>';
    }
    var html = '<div class="segment-row segment-streaming" data-index="' + i + '" data-start="' + seg.start + '">';
    html += '<span class="' + markClass + '" data-segment-id="' + escapeHtml(segId) + '"' + markStyle + '></span>';
    html += sevDotHtml;
    html += '<span class="segment-timestamp">' + formatTime(seg.start) + '</span>';
    html += '<span class="segment-text">' + escapeHtml(seg.text) + '</span>';
    html += '<span class="segment-copy" data-tooltip="Copy text"><span class="segment-copy-icon"></span></span>';
    html += '</div>';
    return html;
  }

  // ---- Elapsed / ETA tracking ----
  // Transcription gets an ETA; thinking agents show elapsed only.
  var _txEtaTrackers = {};
  var _summaryEtaTracker = createEtaTracker();
  var _citationsEtaTracker = createEtaTracker();
  var _frictionEtaTracker = createEtaTracker();
  var _txEtaTicker = createIntervalTicker(_tickTxEta, {
    isActive: _anyTxEtaActive,
  });

  // " \u00b7 0:42 \u00b7 ~1:20 left" or "". Keyed by created_at; seeded from transcribe_started_at.
  function _txEtaSuffix(pid, task) {
    if (!pid || !task || task.status !== "running") return "";
    if (task.phase === "loading_model") return "";
    var entry = _txEtaTrackers[pid];
    if (!entry || entry.createdAt !== task.created_at) {
      var t = createEtaTracker();
      var seedIso = task.transcribe_started_at || task.created_at;
      var seed = seedIso ? Date.parse(seedIso) : NaN;
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
    // Every footer writer renders through here, so one branch covers cancel.
    if (_cancelPending(task)) return "Cancelling\u2026";
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
    // Runs only while something is active (_anyTxEtaActive). Prune idle trackers.
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
      '<span class="streaming-text cg-shimmer">' + _streamingTextStr(progress) + '</span>' +
      '</div>';
  }

  var _streamIndicatorRaf = null;
  var _streamIndicatorPending = null;

  // A RAF paused in a background tab could re-insert a stale indicator after finalize.
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

  // Fires per poll tick while streaming; the likelier jank source.
  function renderPartialSegments(segments, progress) {
    return clipgenPerf.span("transcripts.renderPartialSegments", function () {
      return renderPartialSegmentsImpl(segments, progress);
    });
  }

  function renderPartialSegmentsImpl(segments, progress) {
    var container = qs("#segmentList");
    var empty = qs("#transcriptEmpty");
    var pid = state.streamingParticipant || state.selectedParticipant;
    empty.classList.add("hidden");

    // If user is actively editing a segment, skip DOM mutation to preserve edit state
    if (state.editingTextEl && state.editingTextEl.isConnected) return;

    // Row list changes shape on both append and rebuild paths.
    state.cachedSegmentRows = null;

    // Scroll lives on #trMain, not #segmentList.
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
    // So the finalized render counts as a same-participant rebuild and keeps scroll.
    _renderedSegmentsPid = pid;

    if (nearBottom) {
      scrollHost.scrollTop = scrollHost.scrollHeight;
    }

    // Mirror into state.segments for the marker timeline; ids are "<pid>:<index>".
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
  // RAF handle coalescing the friction-tooltip mousemove below.
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

    // Friction tooltip on hot segments; RAF-coalesced like the timeline canvas.
    container.addEventListener("mousemove", function (e) {
      if (_segTooltipRaf) return;
      var cx = e.clientX, cy = e.clientY, tgt = e.target;
      _segTooltipRaf = requestAnimationFrame(function () {
        _segTooltipRaf = 0;
        if (state.frictionMode === "off") { _hideFrictionTooltip(); return; }
        var row = tgt.closest && tgt.closest(".segment-row");
        if (!row) { _hideFrictionTooltip(); return; }
        var idx = parseInt(row.getAttribute("data-index"), 10);
        var seg = state.segments[idx];
        var frow = seg ? state.frictionBySegId[seg.id] : null;
        if (!frow || !(frow.score > 0)) { _hideFrictionTooltip(); return; }
        _showFrictionTooltip(frow, seg, cx, cy);
      });
    });
    container.addEventListener("mouseleave", function () { _hideFrictionTooltip(); });
  }

  // Streaming marks: id -> { color, id, category, label, severity }; version bumps on write.
  var _streamingMarks = {};
  var _streamingMarksVersion = 0;
  // Per participant, or the first stream's load would swallow later ones.
  var _streamingMarksLoadedByPid = {};

  function _bumpStreamingMarksVersion() {
    _streamingMarksVersion++;
    renderTimeline();
  }

  function _loadStreamingMarks(pid) {
    if (_streamingMarksLoadedByPid[pid]) return;
    _streamingMarksLoadedByPid[pid] = true;
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

  // Persisted seg.marks[0], else the streaming cache. Video satellite reads via TS.
  function getMarkForSegment(seg) {
    if (seg.marks && seg.marks.length > 0) return seg.marks[0];
    var streaming = _streamingMarks[seg.id];
    return streaming || null;
  }

  // ---- Video player + timeline delegators; implementation in transcripts-video.js ----
  function initVideoPlayer() { return TS.initVideoPlayer && TS.initVideoPlayer(); }
  function initTimelineCanvas() { return TS.initTimelineCanvas && TS.initTimelineCanvas(); }
  function initPipScroll() { return TS.initPipScroll && TS.initPipScroll(); }
  function initVideoSync() { return TS.initVideoSync && TS.initVideoSync(); }
  function initPlayerKeyboard() { return TS.initPlayerKeyboard && TS.initPlayerKeyboard(); }
  function renderTimeline() { return TS.renderTimeline && TS.renderTimeline(); }
  function updateTranscribeFill() { return TS.updateTranscribeFill && TS.updateTranscribeFill(); }
  function seekVideo() { return TS.seekVideo && TS.seekVideo.apply(null, arguments); }
  function scrollToSegment() { return TS.scrollToSegment && TS.scrollToSegment.apply(null, arguments); }
  function ignoreNextScroll() { return TS.ignoreNextScroll && TS.ignoreNextScroll(); }
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
      // No reload during streaming; corrections apply on completion.
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

  // Repaint one row's mark dot and badge in place; mirrors renderSegments. Null clears.
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
        sevDot.setAttribute("data-tooltip", markObj.severity);
        dot.insertAdjacentElement("afterend", sevDot);
      }
      if (markObj.label) {
        dot.setAttribute("data-tooltip", markObj.label);
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
        dot.removeAttribute("data-tooltip");
      }
    } else {
      dot.classList.remove("marked");
      dot.style.background = "";
      dot.removeAttribute("data-tooltip");
    }
  }

  function toggleMark(segmentId) {
    var idx = _segmentIndexById(segmentId);
    var seg = idx >= 0 ? state.segments[idx] : null;
    if (!seg) {
      // No loaded row; reload so the new mark still appears.
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
    if (state.streamingParticipant) {
      apiPut("api/marks/" + markId, { label: label || null }).catch(function () {
        showToast("Failed to update mark");
      });
      for (var key in _streamingMarks) {
        if (_streamingMarks[key].id === markId) {
          _streamingMarks[key].label = label || "";
          break;
        }
      }
      return;
    }
    // Write state and repaint optimistically, like updateMarkCategory; restore on failure.
    var found = _findSegmentByMarkId(markId);
    if (!found) {
      apiPut("api/marks/" + markId, { label: label || null }).catch(function () {
        showToast("Failed to update mark");
      });
      return;
    }
    var prevLabel = found.mark.label;
    found.mark.label = label || "";
    _paintSegmentMark(found.idx, found.mark);
    apiPut("api/marks/" + markId, { label: label || null }).catch(function () {
      found.mark.label = prevLabel;
      _paintSegmentMark(found.idx, found.mark);
      showToast("Failed to update mark");
    });
  }

  function updateMarkSeverity(markId, severity) {
    var sev = severity || null;
    // Streaming: update the cache, then pollTaskStatus re-renders the dot now.
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
    // Loaded: optimistic repaint, restore on failure (mirrors updateMarkCategory).
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
    // A provisional mark has no id yet; actions would hit api/marks/null.
    if (markObj && markObj.id == null) return;
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
        pill.setAttribute("data-tooltip", cat.label);
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

    // Severity options from CLIPGEN_CONFIG.severity; blank first option = none.
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
      // updateMarkSeverity owns the mutation so it can roll back on failure.
      updateMarkSeverity(markObj.id, sevSelect.value);
    };

    // Remove button
    var removeBtn = popover.querySelector(".mark-popover-remove");
    removeBtn.onclick = function (e) {
      e.stopPropagation();
      removeMark(markObj.id);
    };

    // Arrows rove the pills, Enter clicks, Esc dismisses; inputs keep their own keys.
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

    // Focus the current category; preventScroll keeps the list from jumping.
    if (pills.length) {
      try { pills[activeIdx].focus({ preventScroll: true }); }
      catch (err) { pills[activeIdx].focus(); }
    }

    // Deferred outside-click close; the timer lets hideMarkPopover cancel the attach.
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

  // ---- Search delegator; implementation in transcripts-search.js ----
  function initSearch() {
    return TS.initSearch && TS.initSearch();
  }

  // ---- Participant pill delegators; implementation in transcripts-pills.js ----
  function renderPills() { return TS.renderPills && TS.renderPills(); }
  function transcribeParticipants() { return TS.transcribeParticipants && TS.transcribeParticipants.apply(null, arguments); }
  function initPillOutsideClick() { return TS.initPillOutsideClick && TS.initPillOutsideClick(); }
  function initPillWheelScroll() { return TS.initPillWheelScroll && TS.initPillWheelScroll(); }

  // Confirm each uncached Whisper model in turn; any cancel aborts. Pills reach via TS.
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
  // Whisper jobs only; summary/citations have their own pollers (see file header).

  // First poll after POLL_INTERVAL, not immediately.
  var _taskPoller = createManagedPoller(pollTaskStatus, POLL_INTERVAL, {
    runImmediately: false,
    label: "transcripts.tasks",
  });

  function startPolling() {
    _taskPoller.start();
  }

  function stopPolling() {
    _taskPoller.stop();
  }

  // Keyed by task id, not participant: old completed tasks linger and would suppress new runs.
  var _refreshedCompletedTaskIds = {};
  // Poll cycles kept alive after completion so the agent "Generating…" state surfaces.
  var _postCompletionGrace = 0;
  var POST_COMPLETION_GRACE_CYCLES = 4; // ~12s at POLL_INTERVAL=3000ms

  // "Cancelling…" ceiling; only a wedged worker gets near it. Hands the stop button back.
  var CANCEL_PENDING_MAX_MS = 30000;

  // Drop pending-cancel flags for tasks no longer active or past the ceiling.
  function _sweepCancellingTasks() {
    var active = {};
    for (var i = 0; i < state.tasks.length; i++) {
      var t = state.tasks[i];
      if (t.status === "running" || t.status === "queued") active[t.id] = true;
    }
    var now = Date.now();
    for (var id in state.cancellingTasks) {
      if (!active[id] || now - state.cancellingTasks[id].at > CANCEL_PENDING_MAX_MS) {
        delete state.cancellingTasks[id];
      }
    }
  }

  function _anyAgentActive() {
    for (var i = 0; i < state.participants.length; i++) {
      var ag = state.participants[i].agents;
      if (ag && (ag.summary === "running" || ag.citations === "running" || ag.friction === "running")) return true;
    }
    return false;
  }

  // Agents may register as running a cycle after completion; reload panels without stomping armed ones.
  function _rearmSelectedAgentPanels() {
    var pid = state.selectedParticipant;
    if (!pid) return;
    var p = _currentParticipant();
    var summaryRunning = !!(p && p.agents && p.agents.summary === "running");
    // _summaryPoller lives in the agents satellite; ask via TS.isSummaryPolling.
    if (summaryRunning && !(TS.isSummaryPolling && TS.isSummaryPolling()) && !state.summaryText) loadSummary(pid);
    // Also reload when only deterministic scores show and the agent has since run.
    var frictionActive = !!(p && p.agents && (p.agents.friction === "running" || p.agents.friction === "done"));
    var showingDeterministic = !!(state.frictionData && state.frictionData.deterministic);
    if (!state.frictionGenerating && _frictionDepMet() &&
        (!state.frictionData || (showingDeterministic && frictionActive))) {
      loadFriction(pid);
    }
  }

  // After a pill agent start/stop: reload pills now and kick the poll loop.
  function _refreshAgentStateNow() {
    loadParticipants().then(function () {
      if (_anyAgentActive()) startPolling();
    });
  }

  // Swap the streaming view for the final transcript; retried each poll until rendered.
  function _finalizeStreamingIfComplete(pid) {
    var running = false;
    var completed = false;
    for (var i = 0; i < state.tasks.length; i++) {
      var t = state.tasks[i];
      if (t.participant !== pid) continue;
      if (t.status === "running" || t.status === "queued") running = true;
      else if (t.status === "completed") completed = true;
    }
    if (running) return;
    // Failed/cancelled/dismissed: drop the flag so normal rendering takes over.
    if (!completed) {
      state.streamingParticipant = null;
      // Keep the partial rows, drop the footer; cancel the queued RAF insert first.
      _cancelStreamingIndicator();
      var ind = document.querySelector("#segmentList .streaming-indicator");
      if (ind) ind.parentNode.removeChild(ind);
      return;
    }
    var ver = state.participantReqVer;
    apiGet("api/transcript/" + pid).then(function (data) {
      if (ver !== state.participantReqVer || state.selectedParticipant !== pid) return;
      // Not merged yet: keep the flag so the next poll retries.
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
      // Sweep first; the fill, indicator and pills below all read the flag.
      _sweepCancellingTasks();
      if (_anyTxEtaActive()) _txEtaTicker.ensure();

      updateTranscribeFill();

      // Snapshot before _finalizeStreamingIfComplete clears it asynchronously.
      var wasStreamingSelected =
        state.streamingParticipant === state.selectedParticipant;

      // Now, not after the async chain, or the indicator freezes at "95%".
      updateStatusIndicator();

      // Stream partial segments for the selected participant's running task
      var selectedRunningTask = null;
      if (state.selectedParticipant) {
        data.tasks.forEach(function (t) {
          if (t.participant === state.selectedParticipant && t.status === "running" && t.partial_count) {
            selectedRunningTask = t;
          }
        });
        // Keep the wait text current; same task pick as the status indicator.
        _setTranscriptEmptyText(_taskForSelectedParticipant());
      }
      if (selectedRunningTask) {
        state.streamingParticipant = state.selectedParticipant;
        _loadStreamingMarks(state.selectedParticipant);
        _syncStreamSegs(selectedRunningTask, function (segs) {
          if (segs.length > 0) renderPartialSegments(segs, selectedRunningTask.progress);
        });
      } else if (state.streamingParticipant) {
        // Streaming view up with no running task; finalize once the transcript is ready.
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

      // Refresh on completion, while any agent runs, and through the grace window.
      var needsRefresh =
        newlyCompleted.length > 0 || _anyAgentActive() || _postCompletionGrace > 0;
      if (newlyCompleted.length > 0) {
        // Re-arm per completion so a queue keeps extending the window.
        _postCompletionGrace = POST_COMPLETION_GRACE_CYCLES;
        _streamingMarks = {};
        _streamingMarksLoadedByPid = {};
        _bumpStreamingMarksVersion();
      }
      if (needsRefresh) {
        loadParticipants().then(function () {
          if (newlyCompleted.length > 0 && state.selectedParticipant &&
              newlyCompleted.indexOf(state.selectedParticipant) >= 0 &&
              !wasStreamingSelected) {
            // Completed while not streaming; the streaming case belongs to _finalizeStreamingIfComplete.
            _setAnalysisPanelVisible(true);
            _restoreActiveTab(state.selectedParticipant);
            loadTranscript(state.selectedParticipant);
            loadSummary(state.selectedParticipant);
            loadFriction(state.selectedParticipant);
          } else if (state.selectedParticipant) {
            // Agents may be chaining server-side; surface the running state.
            _rearmSelectedAgentPanels();
          }
          updateStatusIndicator();
          // Stay alive through the grace window and while a stream awaits finalizing.
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
        // Transcription finished; re-read real cache state before the next gate.
        _forgetWhisperDownloadAgreements();
      }
      _hadActiveTranscriptionLastPoll = hasActive;

      // Count down last so this cycle still counts as in grace.
      if (_postCompletionGrace > 0) _postCompletionGrace--;

      renderPills();
    });
  }

  // ---- Corrections modal delegators; implementation in transcripts-corrections.js ----
  function initCorrectionsModal() {
    return TS.initCorrectionsModal && TS.initCorrectionsModal();
  }
  function loadCorrections() {
    return TS.loadCorrections && TS.loadCorrections();
  }

  // ---- Settings (shared modal lives in settings-modal.js) ----

  // Cached models fetch for per-pill overrides; the shared modal has its own cache.
  var _trModelsCache = null;
  var _trModelsCachePromise = null;
  // Model name -> agreed this session; a downloading model still reads as uncached.
  var _whisperDownloadConfirmed = {};
  // Serializes confirmModelInstall(); there is one shared modal element.
  var _modelInstallChain = Promise.resolve();

  function _trFetchModels() {
    if (_trModelsCache) return Promise.resolve(_trModelsCache);
    if (_trModelsCachePromise) return _trModelsCachePromise;
    _trModelsCachePromise = apiGet("/api/models")
      .then(function (data) {
        // Never cache an unreachable-AI-server result; re-fetch next call.
        if (data && data.ok && !(data.llm && data.llm.available === false)) {
          _trModelsCache = data;
        } else {
          _trModelsCachePromise = null;
        }
        return data;
      })
      .catch(function () { _trModelsCachePromise = null; return null; });
    return _trModelsCachePromise;
  }

  // One /api/audio-info cache for mixer and pill picker; keyed by pid + video_version.
  var _trAudioInfoCache = {};
  var _trAudioInfoPromises = {};

  function _trAudioInfoKey(pid, videoVersion) {
    return pid + ":" + (videoVersion == null ? "" : videoVersion);
  }

  // Synchronous peek so the pill popover's track row doesn't strobe per poll.
  function audioInfoCached(pid, videoVersion) {
    return _trAudioInfoCache[_trAudioInfoKey(pid, videoVersion)] || null;
  }

  function _trFetchAudioInfo(pid, videoVersion) {
    var key = _trAudioInfoKey(pid, videoVersion);
    if (_trAudioInfoCache[key]) return Promise.resolve(_trAudioInfoCache[key]);
    if (_trAudioInfoPromises[key]) return _trAudioInfoPromises[key];
    _trAudioInfoPromises[key] = apiGet("api/audio-info/" + encodeURIComponent(pid))
      .then(function (data) {
        if (!data || !data.ok) {
          delete _trAudioInfoPromises[key];
          return null;
        }
        var info = {
          tracks: data.audio_tracks || [],
          count: data.audio_track_count || 0,
          auto: data.auto_index || 0
        };
        _trAudioInfoCache[key] = info;
        return info;
      })
      .catch(function () { delete _trAudioInfoPromises[key]; return null; });
    return _trAudioInfoPromises[key];
  }

  // ---- Local-model install confirmation ----
  // confirmModelInstall() gates every Whisper and GGUF download behind a dialog.

  function _trFormatModelSize(mb) {
    if (!mb || mb <= 0) return "";
    if (mb >= 1024) return (mb / 1024).toFixed(1) + " GB";
    return Math.round(mb) + " MB";
  }

  // Resolves true on success. Dismissal stops the poll only; the server download continues.
  function downloadLlmModel(model, onProgress, isCancelled) {
    return apiPost("api/models/llm/download", { model: model }).then(function (data) {
      if (!data || !data.ok) return false;
      return new Promise(function (resolve) {
        // createPoller pauses in a backgrounded tab; a raw setInterval would not.
        var misses = 0;
        var poller = createPoller(function () {
          if (isCancelled && isCancelled()) {
            poller.stop();
            resolve(false);
            return;
          }
          apiGet("api/models/llm/download-status?model=" + encodeURIComponent(model))
            .then(function (st) {
              if (!st || !st.ok || !st.found) {
                if (++misses >= 20) { poller.stop(); resolve(false); }
                return;
              }
              misses = 0;
              if (onProgress) onProgress(st);
              if (st.done) {
                poller.stop();
                resolve(!!st.succeeded);
              }
            })
            .catch(function () {
              if (++misses >= 20) { poller.stop(); resolve(false); }
            });
        }, 1000, { runImmediately: true, label: "transcripts.llmDownload" });
        poller.start();
      });
    }).catch(function () { return false; });
  }

  // Resolves true when usable (whisper: agreed; llm: downloaded). Serialized via _modelInstallChain.
  function confirmModelInstall(opts) {
    var run = function () { return _confirmModelInstallNow(opts); };
    var result = _modelInstallChain.then(run, run);
    // Swallow the outcome so a failed dialog doesn't break the queue.
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
      var hintEl = qs("#modelInstallHint");
      var licenseEl = qs("#modelInstallLicense");
      var licenseLink = qs("#modelInstallLicenseLink");

      // Shimmer only while moving; pass working=false for terminal states.
      function setProgressText(text, working) {
        progressText.classList.toggle("cg-shimmer", working !== false);
        progressText.textContent = text;
      }

      progress.classList.add("hidden");
      barFill.style.width = "0%";
      setProgressText("", false);
      hintEl.classList.add("hidden");
      hintEl.textContent = "";
      licenseEl.classList.add("hidden");
      licenseLink.removeAttribute("href");
      cancelBtn.disabled = false;
      cancelBtn.textContent = "Cancel";
      confirmBtn.disabled = false;
      confirmBtn.classList.remove("hidden");

      if (opts.kind === "llm-runtime") {
        // Source-tree runs only (frozen builds bundle llama-server). Not status.message: its "then Refresh" misleads here.
        titleEl.textContent = "AI runtime isn't installed";
        msgEl.textContent = "clipgen couldn't find llama-server on this machine. " +
          "The AI summaries, citations and reports need it — everything else works without it.";
        confirmBtn.textContent = "I've installed it — retry";
        if (opts.hint && opts.hint.length) {
          hintEl.textContent = opts.hint.join("\n");
          hintEl.classList.remove("hidden");
        }
      } else if (opts.kind === "whisper") {
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
        titleEl.textContent = "Download AI model?";
        msgEl.textContent = 'The AI model "' + (opts.label || opts.model) +
          '" used by the ' + (opts.agentKey || "analysis") +
          " agent isn't downloaded. Download it now? The model is stored locally and may take several minutes.";
        // Curated models only; shows whose terms a multi-GB download accepts.
        if (opts.modelUrl) {
          licenseLink.href = opts.modelUrl;
          licenseEl.classList.remove("hidden");
        }
        confirmBtn.textContent = "Download";
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

      // Re-fetch rather than trust the user; a runtime still off PATH fails.
      function onRuntimeConfirm() {
        confirmBtn.disabled = true;
        progress.classList.remove("hidden");
        setProgressText("Checking…");
        // Must run on every ending, or the dialog sticks mid-check with dead buttons.
        function stillUnavailable(text) {
          if (cancelled) return;
          setProgressText(text, false);
          confirmBtn.disabled = false;
          cancelBtn.textContent = "Close";
        }

        _trModelsCache = null;
        _trModelsCachePromise = null;
        _trFetchModels().then(function (data) {
          if (cancelled) return;
          // clipgenLlmStatus(null) reads "ok"; here a failed check is not a yes.
          if (!data || !data.ok) {
            stillUnavailable("Couldn't check — clipgen didn't answer. Try again.");
            return;
          }
          // "stopped" counts; a fresh install is never running yet.
          if (clipgenLlmStatus(data.llm).state !== "missing") {
            showToast("AI runtime found");
            close(true);
            return;
          }
          stillUnavailable(
            "Still not finding it. Open a new terminal and check `llama-server --version`.");
        }).catch(function () {
          stillUnavailable("Couldn't check — clipgen didn't answer. Try again.");
        });
      }

      function onConfirm() {
        if (opts.kind === "llm-runtime") {
          onRuntimeConfirm();
          return;
        }
        if (opts.kind === "whisper") { close(true); return; }
        // llm: download with progress in place; Cancel stays enabled.
        confirmBtn.classList.add("hidden");
        progress.classList.remove("hidden");
        setProgressText("Starting…");
        downloadLlmModel(opts.model, function (st) {
          if (st.total > 0) {
            var pct = Math.max(0, Math.min(100, Math.round((st.completed / st.total) * 100)));
            barFill.style.width = pct + "%";
            setProgressText((st.status || "Downloading") + ": " + pct + "%");
          } else {
            setProgressText(st.status || "Working…");
          }
        }, function () { return cancelled; }).then(function (ok) {
          if (cancelled) return; // dismissed mid-download: no toast, no re-close
          if (ok) {
            _trModelsCache = null;
            _trModelsCachePromise = null;
            showToast("Model downloaded");
            close(true);
          } else {
            // Stay open so the failure is readable; Cancel now resolves false.
            setProgressText("Download failed. Check the model name and connection.", false);
            cancelBtn.textContent = "Close";
          }
        });
      }

      cancelBtn.addEventListener("click", onCancel);
      confirmBtn.addEventListener("click", onConfirm);
      modal.classList.remove("hidden");
      // Escape and backdrop click both cancel; no focus trap.
      openBlockingModal(modal, { onEscape: onCancel, onBackdropClick: onCancel });
    });
  }

  // Gate an agent run: start a stopped runtime, ask about a missing one.
  function ensureAgentModelInstalled(agentKey) {
    return _trFetchModels().then(function (data) {
      var llm = data && data.llm;
      if (!llm) return true;
      var status = clipgenLlmStatus(llm);
      if (status.state === "stopped") {
        return _startAiServer().then(function (fresh) {
          if (!fresh) return false;
          // Server up says nothing about the agent's model yet.
          return _ensureModelFromPayload(fresh, agentKey);
        });
      }
      if (status.state === "missing") {
        return confirmModelInstall({
          kind: "llm-runtime",
          baseUrl: status.baseUrl,
          hint: status.hint,
        }).then(function (recovered) {
          if (!recovered) return false;
          // Installed now, but never already running — same path as stopped.
          return _startAiServer().then(function (fresh) {
            if (!fresh) return false;
            return _ensureModelFromPayload(fresh, agentKey);
          });
        });
      }
      return _ensureModelFromPayload(data, agentKey);
    }).catch(function () { return true; });
  }

  // Resolves the refreshed /api/models payload, or null. The toast covers the ~2 s boot.
  function _startAiServer() {
    showToast("Starting AI server…");
    return apiPost("api/models/llm/start", {})
      .then(function () {
        _trModelsCache = null;
        _trModelsCachePromise = null;
        return _trFetchModels();
      })
      .then(function (fresh) {
        if (fresh && fresh.ok && clipgenLlmStatus(fresh.llm).state === "ok") {
          return fresh;
        }
        showToast("The AI server did not start");
        return null;
      })
      .catch(function (e) {
        showToast((e && (e.serverMessage || e.message)) || "The AI server did not start");
        return null;
      });
  }

  // Model-downloaded half of the gate, against a fetched /api/models payload.
  function _ensureModelFromPayload(data, agentKey) {
    var llm = data && data.llm;
    if (!llm) return true;
    var agents = llm.agents || [];
    var info = null;
    for (var i = 0; i < agents.length; i++) {
      if (agents[i].key === agentKey) { info = agents[i]; break; }
    }
    if (!info || info.installed || !info.model) return true;
    return confirmModelInstall({
      kind: "llm",
      agentKey: agentKey,
      model: info.model,
      label: info.label,
      modelUrl: info.model_url,
    });
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
    if (applyCrossRefSetting(applied, settings)) rerenderCrossRefs();
  }

  // Shared by the settings modal and the command palette's cross-ref command.
  function rerenderCrossRefs() {
    if (state.segments.length > 0) renderSegments();
  }
  window.clipgenRerenderCrossRefs = rerenderCrossRefs;

  // A model change resets the prewarm guards so the new model gets its own prompt.
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
    if (!window.wireSettingsButton) return;
    window.wireSettingsButton({
      initialTab: "Transcription",
      onApply: function (applied, settings) {
        _trModelsCache = null;
        _trModelsCachePromise = null;
        _onTranscribeModelMaybeChanged(
          applied && applied.TRANSCRIBE_MODEL !== undefined
            ? applied.TRANSCRIBE_MODEL
            : _settingValueFromRecords(settings, "TRANSCRIBE_MODEL")
        );
        _applySettingsSnapshot(applied, settings);
      },
    });
  }

  // ---- Clip marked lines ----
  // One clip per mark cluster via Studio's generate-intake.

  // Mirror Studio's #trIntakeClusterThreshold and pad-0 so both pages cut identical spans.
  var CLIP_MARKS_DEFAULT_GAP_SECONDS = 10;
  var CLIP_MARKS_DEFAULT_PAD_SECONDS = 0;

  // { done, failed, total, abort } while streaming, null idle; outlives a dismissed modal.
  var _clipMarksRun = null;
  // Valid resolved marks, refetched every time the modal opens.
  var _clipMarksMarks = [];

  function _clipMarksScopedMarks() {
    var scope = (qs("#clipMarksScope") || {}).value;
    if (scope !== "current") return _clipMarksMarks;
    var pid = state.selectedParticipant;
    return _clipMarksMarks.filter(function (m) { return m.participant === pid; });
  }

  function _clipMarksNumber(sel, fallback, min, max) {
    var raw = parseFloat((qs(sel) || {}).value);
    if (isNaN(raw)) return fallback;
    return Math.min(max, Math.max(min, raw));
  }

  // Preview and payload share this, so the shown count is the clip count.
  function _clipMarksClusters() {
    var marks = _clipMarksScopedMarks();
    if (!marks.length) return [];
    var gap = _clipMarksNumber("#clipMarksGap", CLIP_MARKS_DEFAULT_GAP_SECONDS, 0, 120);
    return window.ClipgenIntakeCluster.clusterTranscriptMarks(marks, gap);
  }

  function renderClipMarksSummary() {
    var summaryEl = qs("#clipMarksSummary");
    var confirmBtn = qs("#clipMarksConfirm");
    if (!summaryEl || !confirmBtn) return;
    if (_clipMarksRun) return; // progress block owns the copy while a run streams
    summaryEl.classList.remove("cg-shimmer"); // the "Loading marks…" fetch landed
    var marks = _clipMarksScopedMarks();
    if (!marks.length) {
      var pid = state.selectedParticipant;
      var scope = (qs("#clipMarksScope") || {}).value;
      summaryEl.textContent =
        scope === "current" && pid
          ? "No marked lines in " + pid + " yet."
          : "No marked lines yet — mark a line with M or the gutter dot.";
      confirmBtn.disabled = true;
      return;
    }
    var clusters = _clipMarksClusters();
    summaryEl.textContent =
      clipgenPluralUnit(marks.length, "marked line", "marked lines") +
      " → " +
      clipgenPluralUnit(clusters.length, "clip", "clips");
    confirmBtn.disabled = false;
  }

  // Option labels carry live mark counts.
  function _renderClipMarksScopeOptions() {
    var sel = qs("#clipMarksScope");
    if (!sel || sel.options.length < 2) return;
    var pid = state.selectedParticipant;
    var mine = pid
      ? _clipMarksMarks.filter(function (m) { return m.participant === pid; }).length
      : 0;
    sel.options[0].textContent =
      (pid ? "Current participant (" + pid + ")" : "Current participant") +
      " — " + clipgenPluralUnit(mine, "mark", "marks");
    sel.options[0].disabled = !pid;
    sel.options[1].textContent =
      "All participants — " + clipgenPluralUnit(_clipMarksMarks.length, "mark", "marks");
    if (!pid) sel.value = "all";
  }

  function _renderClipMarksProgress() {
    var wrap = qs("#clipMarksProgress");
    var fill = qs("#clipMarksBarFill");
    var text = qs("#clipMarksProgressText");
    var confirmBtn = qs("#clipMarksConfirm");
    var cancelBtn = qs("#clipMarksCancel");
    if (!wrap || !fill || !text || !confirmBtn || !cancelBtn) return;
    var run = _clipMarksRun;
    wrap.classList.toggle("hidden", !run);
    confirmBtn.classList.toggle("hidden", !!run);
    cancelBtn.textContent = run ? "Stop" : "Cancel";
    if (!run) return;
    var pct = run.total ? Math.round((run.done / run.total) * 100) : 0;
    fill.style.width = pct + "%";
    text.textContent =
      "Clipping… " + run.done + "/" + run.total +
      (run.failed ? " (" + run.failed + " failed)" : "");
  }

  function openClipMarksModal() {
    var modal = qs("#clipMarksModal");
    if (!modal) return;
    modal.classList.remove("hidden");
    openBlockingModal(modal, {
      onEscape: closeClipMarksModal,
      onBackdropClick: closeClipMarksModal,
    });
    _renderClipMarksProgress();
    // A run in flight owns the dialog's copy; only refresh the pickers when idle.
    if (_clipMarksRun) return;
    qs("#clipMarksSummary").classList.add("cg-shimmer");
    qs("#clipMarksSummary").textContent = "Loading marks…";
    qs("#clipMarksConfirm").disabled = true;
    apiGet("api/marks")
      .then(function (data) {
        _clipMarksMarks = data.ok
          ? (data.marks || []).filter(function (m) { return m.valid; })
          : [];
        _renderClipMarksScopeOptions();
        renderClipMarksSummary();
      })
      .catch(function () {
        qs("#clipMarksSummary").classList.remove("cg-shimmer");
        qs("#clipMarksSummary").textContent = "Could not load marks.";
      });
  }

  function closeClipMarksModal() {
    var modal = qs("#clipMarksModal");
    if (!modal) return;
    closeBlockingModal(modal);
    modal.classList.add("hidden");
  }

  function submitClipMarks() {
    if (_clipMarksRun) return;
    var clusters = _clipMarksClusters();
    if (!clusters.length) return;
    var pad = _clipMarksNumber("#clipMarksPad", CLIP_MARKS_DEFAULT_PAD_SECONDS, 0, 10);
    // Only the start needs clamping; ffmpeg stops at EOF.
    var items = clusters.map(function (c) {
      return {
        participant: c.participant,
        start: Math.max(0, c.start - pad),
        end: c.end + pad,
        event_type: c.category || "transcript",
        event_ids: [],
        source: "transcript",
        mark_ids: c.marks.map(function (m) { return m.id; }),
        text: c.text || "",
        label: c.label || "",
      };
    });

    var run = { done: 0, failed: 0, total: items.length, abort: new AbortController() };
    _clipMarksRun = run;
    _renderClipMarksProgress();

    function handleLine(line) {
      var data;
      try { data = JSON.parse(line); } catch (_) { return; }
      // The trailing {"cancelled": true} line has no index and stays out of the tally.
      if (!data || typeof data.index !== "number") return;
      run.done++;
      if (!data.ok) run.failed++;
      _renderClipMarksProgress();
    }

    function finish(message) {
      _clipMarksRun = null;
      _renderClipMarksProgress();
      closeClipMarksModal();
      showToast(message);
    }

    apiPostNDJSON(
      "../studio/api/generate-intake",
      { items: items, format: "clip" },
      { signal: run.abort.signal, onLine: handleLine }
    )
      .then(function () {
        var made = run.done - run.failed;
        finish(
          run.failed
            ? clipgenPluralUnit(made, "clip", "clips") + " generated, " + run.failed + " failed"
            : clipgenPluralUnit(made, "clip", "clips") + " generated — open Studio to review"
        );
      })
      .catch(function (err) {
        var aborted = err && (err.name === "AbortError" || err.code === 20);
        finish(aborted ? "Clip generation cancelled" : "Clip generation failed: " + (err && err.message));
      });
  }

  function onClipMarksCancel() {
    if (!_clipMarksRun) {
      closeClipMarksModal();
      return;
    }
    apiPost("../studio/api/generate-intake/cancel", {}).catch(function () {});
    _clipMarksRun.abort.abort();
  }

  function initClipMarksModal() {
    qs("#clipMarksCancel").addEventListener("click", onClipMarksCancel);
    qs("#clipMarksConfirm").addEventListener("click", submitClipMarks);
    qs("#clipMarksScope").addEventListener("change", renderClipMarksSummary);
    qs("#clipMarksGap").addEventListener("input", renderClipMarksSummary);
  }

  // ---- Embed subtitles ----
  // Soft-mux transcripts into video copies; twin of Clip Marked Lines.

  // { done, failed, total, abort } while streaming, null idle; outlives a dismissed modal.
  var _embedSubsRun = null;

  // Multi-part participants are filtered client-side; the server would refuse them anyway.
  function _embedSubsCandidates() {
    var ps = state.participants || [];
    var out = [];
    for (var i = 0; i < ps.length; i++) {
      if (ps[i].has_transcript) out.push(ps[i]);
    }
    return out;
  }

  function _embedSubsIsMultiPart(p) {
    return !!(p.video_paths && p.video_paths.length > 1);
  }

  function _embedSubsScoped() {
    var all = _embedSubsCandidates();
    if ((qs("#embedSubsScope") || {}).value !== "current") return all;
    var pid = state.selectedParticipant;
    return all.filter(function (p) { return p.id === pid; });
  }

  // Container extension of a participant's first source file, lowercased.
  function _embedSubsExt(p) {
    var path = (p.video_paths && p.video_paths[0]) || "";
    var dot = path.lastIndexOf(".");
    return dot === -1 ? "" : path.slice(dot).toLowerCase();
  }

  function _embedSubsContainers() {
    return CLIPGEN_CONFIG.subtitleContainers || { supported: [], alwaysDefault: [] };
  }

  // Mirrors mux_subtitles' `codec is None` guard so the summary count is honest.
  function _embedSubsIsUnsupported(p) {
    var supported = _embedSubsContainers().supported || [];
    return supported.indexOf(_embedSubsExt(p)) === -1;
  }

  function _embedSubsTargets() {
    return _embedSubsScoped().filter(function (p) {
      return !_embedSubsIsMultiPart(p) && !_embedSubsIsUnsupported(p);
    });
  }

  // ISOBMFF ignores -disposition:s:0 (measured, ffmpeg 8.1.2); only .mkv/.webm can be off.
  function _embedSubsAlwaysDefault(targets) {
    var always = _embedSubsContainers().alwaysDefault || [];
    return targets.filter(function (p) {
      return always.indexOf(_embedSubsExt(p)) !== -1;
    });
  }

  function renderEmbedSubsSummary() {
    var summaryEl = qs("#embedSubsSummary");
    var confirmBtn = qs("#embedSubsConfirm");
    if (!summaryEl || !confirmBtn) return;
    if (_embedSubsRun) return; // progress block owns the copy while a run streams
    var targets = _embedSubsTargets();
    var scoped = _embedSubsScoped();
    var skipped = scoped.filter(_embedSubsIsMultiPart);
    var unsupported = scoped.filter(function (p) {
      return !_embedSubsIsMultiPart(p) && _embedSubsIsUnsupported(p);
    });
    var skipNote = skipped.length
      ? " " + skipped.map(function (p) { return p.id; }).join(", ") +
        " skipped (multi-part " + (skipped.length === 1 ? "recording" : "recordings") + ")."
      : "";
    if (unsupported.length) {
      skipNote += " " + unsupported.map(function (p) { return p.id; }).join(", ") +
        " skipped (subtitles cannot be muxed into " +
        _embedSubsExt(unsupported[0]) + ").";
    }
    if (!targets.length) {
      var scope = (qs("#embedSubsScope") || {}).value;
      // Name the real blocker; only "no transcript" is fixed by transcribing.
      if (unsupported.length && !skipped.length) {
        summaryEl.textContent =
          (unsupported.length === 1
            ? unsupported[0].id + "'s recording is " + _embedSubsExt(unsupported[0])
            : "These recordings are in a container") +
          ", which cannot carry an embedded subtitle track. Supported: " +
          (_embedSubsContainers().supported || []).join(", ") + ".";
      } else if (skipped.length) {
        summaryEl.textContent =
          (skipped.length === 1
            ? skipped[0].id + "'s transcript spans several video files"
            : "Every transcript here spans several video files") +
          ", which cannot be muxed back into one subtitled copy.";
      } else {
        summaryEl.textContent =
          scope === "current" && state.selectedParticipant
            ? "No transcript for " + state.selectedParticipant + " yet."
            : "No transcripts yet — transcribe a video first.";
      }
      confirmBtn.disabled = true;
      return;
    }
    // Only when unticked; ticked agrees with the mp4 muxer anyway.
    var stuckOn = (qs("#embedSubsDefault") || {}).checked
      ? []
      : _embedSubsAlwaysDefault(targets);
    var stuckNote = stuckOn.length
      ? " " + (stuckOn.length === targets.length ? "The track" : stuckOn.length + " of these")
        + " will still be on by default — .mp4/.mov cannot carry a subtitle track that is off."
      : "";
    summaryEl.textContent =
      clipgenPluralUnit(targets.length, "transcript", "transcripts") +
      " → " +
      clipgenPluralUnit(targets.length, "subtitled video", "subtitled videos") +
      "." + skipNote + stuckNote;
    confirmBtn.disabled = false;
  }

  // Option labels carry live counts (mirrors the clip-marks picker).
  function _renderEmbedSubsScopeOptions() {
    var sel = qs("#embedSubsScope");
    if (!sel || sel.options.length < 2) return;
    var pid = state.selectedParticipant;
    var all = _embedSubsCandidates();
    var mine = pid
      ? all.filter(function (p) { return p.id === pid; }).length
      : 0;
    sel.options[0].textContent =
      (pid ? "Current participant (" + pid + ")" : "Current participant") +
      " — " + clipgenPluralUnit(mine, "transcript", "transcripts");
    sel.options[0].disabled = !mine;
    sel.options[1].textContent =
      "All participants — " + clipgenPluralUnit(all.length, "transcript", "transcripts");
    if (!mine) sel.value = "all";
  }

  function _renderEmbedSubsProgress() {
    var wrap = qs("#embedSubsProgress");
    var fill = qs("#embedSubsBarFill");
    var text = qs("#embedSubsProgressText");
    var confirmBtn = qs("#embedSubsConfirm");
    var cancelBtn = qs("#embedSubsCancel");
    if (!wrap || !fill || !text || !confirmBtn || !cancelBtn) return;
    var run = _embedSubsRun;
    wrap.classList.toggle("hidden", !run);
    confirmBtn.classList.toggle("hidden", !!run);
    cancelBtn.textContent = run ? "Stop" : "Cancel";
    if (!run) return;
    var pct = run.total ? Math.round((run.done / run.total) * 100) : 0;
    fill.style.width = pct + "%";
    text.textContent =
      "Embedding… " + run.done + "/" + run.total +
      (run.failed ? " (" + run.failed + " failed)" : "");
  }

  function openEmbedSubsModal() {
    var modal = qs("#embedSubsModal");
    if (!modal) return;
    modal.classList.remove("hidden");
    openBlockingModal(modal, {
      onEscape: closeEmbedSubsModal,
      onBackdropClick: closeEmbedSubsModal,
    });
    _renderEmbedSubsProgress();
    // A run in flight owns the dialog's copy; only refresh the pickers when idle.
    if (_embedSubsRun) return;
    _renderEmbedSubsScopeOptions();
    renderEmbedSubsSummary();
  }

  function closeEmbedSubsModal() {
    var modal = qs("#embedSubsModal");
    if (!modal) return;
    closeBlockingModal(modal);
    modal.classList.add("hidden");
  }

  function submitEmbedSubs() {
    if (_embedSubsRun) return;
    var targets = _embedSubsTargets();
    if (!targets.length) return;
    var pids = targets.map(function (p) { return p.id; });
    var defaultTrack = !!(qs("#embedSubsDefault") || {}).checked;

    var run = { done: 0, failed: 0, total: pids.length, abort: new AbortController() };
    _embedSubsRun = run;
    _renderEmbedSubsProgress();
    var outputDir = "";

    var sawDone = false;

    function handleLine(line) {
      var data;
      try { data = JSON.parse(line); } catch (_) { return; }
      if (!data) return;
      // Header and trailing sentinel lines have no index; they stay out of the tally.
      if (data.output_dir) outputDir = data.output_dir;
      if (data.token) run.token = data.token; // echoed by Stop to scope the cancel
      if (data.done) sawDone = true;
      if (typeof data.index !== "number") return;
      run.done++;
      if (!data.ok) run.failed++;
      _renderEmbedSubsProgress();
    }

    function finish(message) {
      _embedSubsRun = null;
      _renderEmbedSubsProgress();
      closeEmbedSubsModal();
      showToast(message);
    }

    apiPostNDJSON(
      "api/embed-subtitles",
      { participants: pids, default_track: defaultTrack },
      { signal: run.abort.signal, onLine: handleLine }
    )
      .then(function () {
        var made = run.done - run.failed;
        var where = outputDir ? " to " + outputDir : "";
        // No sentinel means the server died partway; the reader can't tell otherwise.
        if (!sawDone) {
          finish(
            "Subtitle embedding stopped early — " +
              clipgenPluralUnit(made, "video was", "videos were") + " written" + where +
              " of " + run.total + ". Check the clipgen log."
          );
          return;
        }
        finish(
          run.failed
            ? clipgenPluralUnit(made, "subtitled video", "subtitled videos") + " written" + where + ", " + run.failed + " failed"
            : clipgenPluralUnit(made, "subtitled video", "subtitled videos") + " written" + where
        );
      })
      .catch(function (err) {
        var aborted = err && (err.name === "AbortError" || err.code === 20);
        finish(aborted ? "Subtitle embedding cancelled" : "Subtitle embedding failed: " + (err && err.message));
      });
  }

  function onEmbedSubsCancel() {
    if (!_embedSubsRun) {
      closeEmbedSubsModal();
      return;
    }
    // The server stops between files; the abort just drops our end.
    apiPost("api/embed-subtitles/cancel", { token: _embedSubsRun.token || null }).catch(function () {});
    _embedSubsRun.abort.abort();
  }

  function initEmbedSubsModal() {
    qs("#embedSubsCancel").addEventListener("click", onEmbedSubsCancel);
    qs("#embedSubsConfirm").addEventListener("click", submitEmbedSubs);
    qs("#embedSubsScope").addEventListener("change", renderEmbedSubsSummary);
    // The checkbox decides whether the .mp4 caveat shows.
    qs("#embedSubsDefault").addEventListener("change", renderEmbedSubsSummary);
  }

  // ---- Normalize audio ----
  // Loudnorm videos in place, keeping .orig like remux.

  // { done, failed, total, abort } while streaming, null idle; outlives a dismissed modal.
  var _normAudioRun = null;

  // pid -> parts with a kept .orig. Only fully-kept participants are excluded.
  var _normAudioKept = {};

  // Pin which participant the async-built track checkboxes belong to.
  var _normAudioTrackPid = null;
  var _normAudioTrackInfo = null;

  function _normAudioCandidates() {
    var ps = state.participants || [];
    var out = [];
    for (var i = 0; i < ps.length; i++) {
      // A transcript is not required — normalization reads only the media.
      if (ps[i].has_video) out.push(ps[i]);
    }
    return out;
  }

  function _normAudioScoped() {
    var all = _normAudioCandidates();
    if ((qs("#normAudioScope") || {}).value !== "current") return all;
    var pid = state.selectedParticipant;
    return all.filter(function (p) { return p.id === pid; });
  }

  function _normAudioKeptCount(p) {
    return _normAudioKept[p.id] || 0;
  }

  function _normAudioIsFullyKept(p) {
    var parts = (p.video_paths || []).length || 1;
    return _normAudioKeptCount(p) >= parts;
  }

  function _normAudioTargets() {
    return _normAudioScoped().filter(function (p) {
      return !_normAudioIsFullyKept(p);
    });
  }

  // Files a run would actually rewrite: parts whose backup slot is free.
  function _normAudioFileCount(targets) {
    var n = 0;
    for (var i = 0; i < targets.length; i++) {
      var parts = (targets[i].video_paths || []).length || 1;
      n += Math.max(0, parts - _normAudioKeptCount(targets[i]));
    }
    return n;
  }

  // Mode word for scope=all, checked indices for multi-track current, else "auto".
  function _normAudioTracksSpec() {
    if ((qs("#normAudioScope") || {}).value !== "current") {
      return (qs("#normAudioTrackMode") || {}).value || "auto";
    }
    var info = _normAudioTrackInfo;
    if (!info || info.count <= 1) return "auto";
    var boxes = document.querySelectorAll("#normAudioTrackList input");
    var picked = [];
    for (var i = 0; i < boxes.length; i++) {
      if (boxes[i].checked) picked.push(parseInt(boxes[i].value, 10));
    }
    return picked;
  }

  function renderNormAudioSummary() {
    var summaryEl = qs("#normAudioSummary");
    var confirmBtn = qs("#normAudioConfirm");
    if (!summaryEl || !confirmBtn) return;
    if (_normAudioRun) return; // progress block owns the copy while a run streams
    var scope = (qs("#normAudioScope") || {}).value;
    var scoped = _normAudioScoped();
    var targets = _normAudioTargets();
    var fullyKept = scoped.filter(_normAudioIsFullyKept);
    var resuming = targets.filter(function (p) { return _normAudioKeptCount(p) > 0; });
    var keptNote = fullyKept.length
      ? " " + fullyKept.map(function (p) { return p.id; }).join(", ") +
        " skipped (an earlier original is still kept — delete or restore it first)."
      : "";
    if (resuming.length) {
      // The server skips parts whose backup slot is occupied.
      keptNote += " " + resuming.map(function (p) { return p.id; }).join(", ") +
        (resuming.length === 1
          ? " resumes where it stopped"
          : " resume where they stopped") +
        " — already-rewritten parts are skipped.";
    }
    if (!targets.length) {
      summaryEl.textContent = fullyKept.length
        ? "Nothing to normalize —" + keptNote
        : scope === "current" && state.selectedParticipant
          ? "No source video for " + state.selectedParticipant + "."
          : "No source videos yet.";
      confirmBtn.disabled = true;
      return;
    }
    if (scope === "current" && _normAudioTrackPid && !_normAudioTrackInfo) {
      summaryEl.textContent = "Checking audio tracks…";
      confirmBtn.disabled = true;
      return;
    }
    var spec = _normAudioTracksSpec();
    if (Object.prototype.toString.call(spec) === "[object Array]" && !spec.length) {
      summaryEl.textContent = "Select at least one track to normalize." + keptNote;
      confirmBtn.disabled = true;
      return;
    }
    var files = _normAudioFileCount(targets);
    var trackNote =
      scope === "current" && _normAudioTrackInfo && _normAudioTrackInfo.count === 1
        ? " (1 audio track)"
        : "";
    summaryEl.textContent =
      clipgenPluralUnit(files, "video", "videos") + trackNote +
      " → rewritten in place; " +
      (files === 1 ? "the original is" : "originals are") +
      " kept beside " + (files === 1 ? "it" : "them") + "." + keptNote;
    confirmBtn.disabled = false;
  }

  // Option labels carry live counts (mirrors the embed picker).
  function _renderNormAudioScopeOptions() {
    var sel = qs("#normAudioScope");
    if (!sel || sel.options.length < 2) return;
    var pid = state.selectedParticipant;
    var all = _normAudioCandidates();
    var mine = pid
      ? all.filter(function (p) { return p.id === pid; }).length
      : 0;
    sel.options[0].textContent =
      (pid ? "Current participant (" + pid + ")" : "Current participant") +
      " — " + clipgenPluralUnit(mine, "video", "videos");
    sel.options[0].disabled = !mine;
    sel.options[1].textContent =
      "All participants — " + clipgenPluralUnit(all.length, "video", "videos");
    if (!mine) sel.value = "all";
  }

  // Mode select for scope=all; per-track checkboxes for a multi-track current participant.
  function _renderNormAudioTrackField() {
    var modeLabel = qs("#normAudioTrackModeLabel");
    var list = qs("#normAudioTrackList");
    if (!modeLabel || !list) return;
    var scope = (qs("#normAudioScope") || {}).value;
    var pid = scope === "current" ? state.selectedParticipant : null;
    var p = pid
      ? _normAudioCandidates().filter(function (c) { return c.id === pid; })[0]
      : null;
    if (!p) {
      _normAudioTrackPid = null;
      _normAudioTrackInfo = null;
      modeLabel.classList.toggle("hidden", scope === "current");
      list.classList.add("hidden");
      list.innerHTML = "";
      renderNormAudioSummary();
      return;
    }
    modeLabel.classList.add("hidden");
    list.classList.add("hidden");
    list.innerHTML = "";
    _normAudioTrackPid = pid;
    _normAudioTrackInfo = null;
    renderNormAudioSummary(); // "Checking audio tracks…" while the probe runs
    _trFetchAudioInfo(pid, p.video_version).then(function (info) {
      // A scope flip or participant change while the probe ran owns the field.
      if (_normAudioTrackPid !== pid) return;
      _normAudioTrackInfo = info || { tracks: [], count: 1, auto: 0 };
      if (_normAudioTrackInfo.count > 1) {
        for (var i = 0; i < _normAudioTrackInfo.tracks.length; i++) {
          var row = document.createElement("label");
          row.className = "param-modal-label";
          var text = document.createElement("span");
          // Late-bound; transcripts-pills.js publishes the helper after the hub loads.
          text.textContent = TS.trackOptionLabel
            ? TS.trackOptionLabel(_normAudioTrackInfo.tracks[i], i)
            : "Track " + (i + 1);
          var box = document.createElement("input");
          box.type = "checkbox";
          box.className = "param-modal-checkbox";
          box.value = String(i);
          box.checked = i === _normAudioTrackInfo.auto;
          row.appendChild(text);
          row.appendChild(box);
          list.appendChild(row);
        }
        list.classList.remove("hidden");
      }
      renderNormAudioSummary();
    });
  }

  function _renderNormAudioProgress() {
    var wrap = qs("#normAudioProgress");
    var fill = qs("#normAudioBarFill");
    var text = qs("#normAudioProgressText");
    var confirmBtn = qs("#normAudioConfirm");
    var cancelBtn = qs("#normAudioCancel");
    if (!wrap || !fill || !text || !confirmBtn || !cancelBtn) return;
    var run = _normAudioRun;
    wrap.classList.toggle("hidden", !run);
    confirmBtn.classList.toggle("hidden", !!run);
    cancelBtn.textContent = run ? "Stop" : "Cancel";
    if (!run) return;
    var pct = run.total ? Math.round((run.done / run.total) * 100) : 0;
    fill.style.width = pct + "%";
    text.textContent =
      "Normalizing… " + run.done + "/" + run.total +
      (run.failed ? " (" + run.failed + " failed)" : "");
  }

  function openNormalizeAudioModal() {
    var modal = qs("#normAudioModal");
    if (!modal) return;
    modal.classList.remove("hidden");
    openBlockingModal(modal, {
      onEscape: closeNormalizeAudioModal,
      onBackdropClick: closeNormalizeAudioModal,
    });
    _renderNormAudioProgress();
    // A run in flight owns the dialog's copy; only refresh the pickers when idle.
    if (_normAudioRun) return;
    _normAudioKept = {};
    _renderNormAudioScopeOptions();
    _renderNormAudioTrackField();
    // Kept-original state lives on disk; remux/status re-probes on every call.
    apiGet("api/remux/status")
      .then(function (data) {
        if (!data || !data.ok || _normAudioRun) return;
        var kept = data.kept || {};
        var map = {};
        for (var pid in kept) {
          if (kept[pid] && kept[pid].length) map[pid] = kept[pid].length;
        }
        _normAudioKept = map;
        renderNormAudioSummary();
      })
      .catch(function () {});
  }

  function closeNormalizeAudioModal() {
    var modal = qs("#normAudioModal");
    if (!modal) return;
    closeBlockingModal(modal);
    modal.classList.add("hidden");
  }

  function submitNormalizeAudio() {
    if (_normAudioRun) return;
    var targets = _normAudioTargets();
    if (!targets.length) return;
    var pids = targets.map(function (p) { return p.id; });
    var spec = _normAudioTracksSpec();

    var run = {
      done: 0,
      failed: 0,
      changed: 0,
      total: pids.length,
      abort: new AbortController(),
    };
    _normAudioRun = run;
    _renderNormAudioProgress();

    var sawDone = false;

    function handleLine(line) {
      var data;
      try { data = JSON.parse(line); } catch (_) { return; }
      if (!data) return;
      // Header and trailing sentinel lines have no index; they stay out of the tally.
      if (data.token) run.token = data.token; // echoed by Stop to scope the cancel
      if (data.done) sawDone = true;
      if (typeof data.index !== "number") return;
      run.done++;
      if (!data.ok) run.failed++;
      // Files swapped; can be non-zero on an ok=false multi-part line.
      if (typeof data.parts_done === "number") run.changed += data.parts_done;
      _renderNormAudioProgress();
    }

    function finish(message) {
      _normAudioRun = null;
      _renderNormAudioProgress();
      closeNormalizeAudioModal();
      showToast(message);
      // Swapped files invalidate the <video> stream; reload like media-banner.js, after the toast.
      if (run.changed > 0) {
        setTimeout(function () { window.location.reload(); }, 1500);
      }
    }

    apiPostNDJSON(
      "api/normalize-audio",
      { participants: pids, tracks: spec },
      { signal: run.abort.signal, onLine: handleLine }
    )
      .then(function () {
        var made = run.done - run.failed;
        // No sentinel means the server died partway; the reader can't tell otherwise.
        if (!sawDone) {
          finish(
            "Audio normalization stopped early — " +
              clipgenPluralUnit(made, "video was", "videos were") + " rewritten of " +
              run.total + ". Check the clipgen log."
          );
          return;
        }
        finish(
          run.failed
            ? clipgenPluralUnit(made, "video", "videos") + " normalized, " + run.failed + " failed"
            : clipgenPluralUnit(made, "video", "videos") + " normalized; originals kept beside the sources"
        );
      })
      .catch(function (err) {
        var aborted = err && (err.name === "AbortError" || err.code === 20);
        finish(aborted ? "Audio normalization cancelled" : "Audio normalization failed: " + (err && err.message));
      });
  }

  function onNormalizeAudioCancel() {
    if (!_normAudioRun) {
      closeNormalizeAudioModal();
      return;
    }
    // Unlike embed, Stop interrupts ffmpeg mid-encode; the abort just drops our end.
    apiPost("api/normalize-audio/cancel", { token: _normAudioRun.token || null }).catch(function () {});
    _normAudioRun.abort.abort();
  }

  function initNormalizeAudioModal() {
    qs("#normAudioCancel").addEventListener("click", onNormalizeAudioCancel);
    qs("#normAudioConfirm").addEventListener("click", submitNormalizeAudio);
    // Scope picks the track control; the field re-renders the summary itself.
    qs("#normAudioScope").addEventListener("change", _renderNormAudioTrackField);
    qs("#normAudioTrackMode").addEventListener("change", renderNormAudioSummary);
    // Delegated: the checkbox rows are rebuilt per participant.
    qs("#normAudioTrackList").addEventListener("change", renderNormAudioSummary);
  }

  // ---- Boot ----

  // Video, no transcript, nothing queued/running: /api/transcribe has no in-flight guard.
  function _untranscribedParticipants() {
    var ps = state.participants || [];
    var tasks = state.tasks || [];
    var busy = {};
    for (var t = 0; t < tasks.length; t++) {
      if (tasks[t].status === "queued" || tasks[t].status === "running") busy[tasks[t].participant] = true;
    }
    var pids = [];
    for (var i = 0; i < ps.length; i++) {
      if (ps[i].has_video && !ps[i].has_transcript && !busy[ps[i].id]) pids.push(ps[i].id);
    }
    return pids;
  }

  // One POST for the list; _postTranscribe owns toast, poll restart and model confirm.
  function runTranscribeAll() {
    var pids = _untranscribedParticipants();
    if (!pids.length) return;
    transcribeParticipants(pids, false);
  }

  var _rebuildTopNavActions = function () {};

  // Published on TS; transcripts-agents.js gates panel refreshes on it.
  function _currentParticipantHasTranscript() {
    var pid = state.selectedParticipant;
    if (!pid || !state.participants) return false;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) return !!state.participants[i].has_transcript;
    }
    return false;
  }

  function refreshTopNavActions() {
    _rebuildTopNavActions();
  }

  // Rebuilt on open so Transcribe All re-counts against in-flight tasks.
  function initTopNavActions() {
    if (!window.ClipgenTopNav) return;
    _rebuildTopNavActions = window.ClipgenTopNav.installQuickActions(function () {
      var pending = _untranscribedParticipants().length;
      return [
        {
          icon: "microphone",
          label: "Transcribe All",
          action: runTranscribeAll,
          disabled: pending === 0,
          title: pending
            ? "Queue transcription for the " + clipgenPluralUnit(pending, "participant", "participants") + " without a transcript"
            : "Every participant with a video already has a transcript or is queued.",
        },
        {
          icon: "language",
          label: "Embed Subtitles…",
          action: openEmbedSubsModal,
          // Never gated: the modal reports counts and disables its own button.
          title: "Write a subtitled copy of each source video into the output folder; the originals are never modified",
        },
        {
          icon: "speaker-wave",
          label: "Normalize Audio…",
          action: openNormalizeAudioModal,
          // Never gated, same as the neighbours.
          title: "Rewrite source videos in place with loudness-normalized audio; the original is kept beside each file until you delete it",
        },
        {
          icon: "scissors",
          label: "Clip Marked Lines…",
          action: openClipMarksModal,
          // Never gated; keeps the menu free of an async mark fetch.
          title: "Cut a clip for every manually marked line",
        },
        window.ClipgenExportActions.exportQuickAction(),
      ];
    }, { rebuildOnOpen: true });
  }

  // Palette commands beyond the auto-ingested quick actions; provider runs per open.
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
          id: "transcripts:search",
          title: "Focus transcript search",
          icon: "magnifying-glass",
          keywords: "find text query",
          section: "Transcripts",
          visible: function () { return !!document.getElementById("searchInput"); },
          // Runs after the palette restores focus, so this call wins.
          run: function () { document.getElementById("searchInput").focus(); },
        },
        clickCommand("transcripts:shortcuts", "Keyboard shortcuts", "command-line",
          "cheatsheet keys help", "shortcutsBtn"),
        clickCommand("transcripts:tab-summary", "Show Summary tab", "table-cells",
          "panel agents", "tabBtnSummary"),
        clickCommand("transcripts:tab-friction", "Show Friction tab", "table-cells",
          "panel analysis moments", "tabBtnFriction"),
        clickCommand("transcripts:toggle-video", "Toggle video panel", "video-camera",
          "hide show collapse drawer player", "videoCollapseBtn"),
        clickCommand("transcripts:toggle-captions", "Toggle captions", "language",
          "subtitles cc text", "videoCcBtn"),
        {
          id: "transcripts:cycle-friction-mode",
          title: "Cycle friction mode (off / highlight / isolate)",
          icon: "fire",
          keywords: "friction analysis overlay heatmap filter isolate highlight timeline",
          section: "Transcripts",
          // Gated on data: the mode means nothing without friction scores.
          enabled: function () {
            var fd = state.frictionData;
            return !!(fd && fd.segments && fd.segments.length);
          },
          run: function () { cycleFrictionMode(); },
        },
      ];
      // Selects in place; the palette's built-in provider adds cross-page "Open …".
      (state.participants || []).forEach(function (p) {
        cmds.push({
          id: "transcripts:p:" + p.id,
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
    initStatusIndicatorTooltip();
    checkNavLinks();
    initFrontendSwitcher();
    initSearch();
    initPillOutsideClick();
    initPillWheelScroll();
    initCorrectionsModal();
    initClipMarksModal();
    initEmbedSubsModal();
    initNormalizeAudioModal();
    initVideoPlayer();
    initVideoSync();
    initTimelineCanvas();
    initPipScroll();
    initPlayerKeyboard();
    initPanelTabs();
    // #tab=friction deep links; the #P07 form is handled in loadParticipants.
    var hashTab = clipgenHashTab();
    if (hashTab === "summary" || hashTab === "friction") {
      var hashTabBtn = qs(hashTab === "friction" ? "#tabBtnFriction" : "#tabBtnSummary");
      if (hashTabBtn) hashTabBtn.click();
    }
    initSummaryActions();
    initFriction();
    initFrictionMode();
    initTranscriptSettings();
    initTopNavActions();
    initCommandPalette();

    // Pause every poller while hidden; resume what was active on focus.
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
        // Agents finish silently in the background; reload and re-arm their polls.
        if (state.selectedParticipant) {
          loadSummary(state.selectedParticipant);
          loadFriction(state.selectedParticipant);
        }
      }
    });

    // Cmd-Tab keeps document.hidden false yet Safari throttles timers; re-poll on focus.
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
  // Assigned synchronously so satellites see it; see agents/skills/carve-satellite/SKILL.md.
  var TS = (window.ClipgenTranscripts = window.ClipgenTranscripts || {});
  TS.state = state;
  TS.showToast = showToast;
  TS.reportAgentError = reportAgentError;
  // Hub helpers the satellites call outward.
  TS.loadTranscript = loadTranscript; // corrections, search, agents
  TS.findOverlapsForSearch = findOverlapsForSearch; // search
  TS.selectParticipant = selectParticipant; // search, pills
  TS.cycleParticipant = cycleParticipant; // video (Z/X participant cycle)
  // Streaming segments for the streamed participant; search reads them here.
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
  TS.refreshTranscribeWording = refreshTranscribeWording; // pills (cancel repaints before the next poll)
  TS.startPolling = startPolling; // pills
  TS._refreshAgentStateNow = _refreshAgentStateNow; // pills
  TS._trFetchModels = _trFetchModels; // pills
  TS._trFetchAudioInfo = _trFetchAudioInfo; // pills (audio-track picker)
  TS.audioInfoCached = audioInfoCached; // pills (synchronous warm-cache peek)
  TS.ensureAgentModelInstalled = ensureAgentModelInstalled; // pills, agents
  TS.confirmModelInstall = confirmModelInstall; // shot.py state probing; future satellites
  TS._confirmUncachedWhisperModels = _confirmUncachedWhisperModels; // pills (model-install kept in hub)
  // Hub helpers the agents satellite calls outward.
  TS.renderSegments = renderSegments; // agents (heatmap toggle, friction mark-all)
  TS._txEtaTicker = _txEtaTicker; // agents (summary/citations/friction elapsed)
  TS._summaryEtaTracker = _summaryEtaTracker; // agents
  TS._citationsEtaTracker = _citationsEtaTracker; // agents
  TS._frictionEtaTracker = _frictionEtaTracker; // agents
  TS._updateAgentElapsed = _updateAgentElapsed; // agents
  TS._currentParticipantHasTranscript = _currentParticipantHasTranscript; // agents (panel-visible guard)

})();
