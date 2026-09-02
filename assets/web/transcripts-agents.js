/* clipgen Transcripts analysis-panel satellite — transcripts-agents.js
 *
 * The tabbed analysis panel: the AI summary (+ inline edit + citations) and the
 * friction pass (mode switch, score histogram, category chips, moment jump strip,
 * and the decorations they drive on the transcript below), plus the panel tab
 * switching. These are the local-LLM "thinking agent" results
 * surfaced per participant. Loaded LAST (after transcripts.js + the other
 * satellites); reads the hub's shared state + helpers through
 * window.ClipgenTranscripts (TS) and publishes its load/clear/stop/init entry
 * points back so selectParticipant, the task poller, the visibility/focus
 * handlers, boot, and the segment-list hover can reach them. The pills satellite
 * also reaches loadFriction via TS.loadFriction.
 *
 * The ETA trackers + ticker and _updateAgentElapsed live in the hub (shared with
 * the transcription ETA path) and are used here by reference. isSummaryPolling
 * lets the hub poller re-arm the summary poll without reaching into the
 * summary descriptor's poller directly. Plain utils.js globals (qs/el/escapeHtml/apiGet/apiPost/formatTime/
 * createPoller/MARK_CATEGORIES/CLIPGEN_CONFIG/...) are reached via the scope chain.
 */
(function () {
  "use strict";

  var TS = window.ClipgenTranscripts;
  var state = TS.state;
  var showToast = TS.showToast,
    reportAgentError = TS.reportAgentError,
    renderTimeline = TS.renderTimeline,
    seekVideo = TS.seekVideo,
    scrollToSegment = TS.scrollToSegment,
    loadTranscript = TS.loadTranscript,
    ensureAgentModelInstalled = TS.ensureAgentModelInstalled,
    _trFetchModels = TS._trFetchModels,
    _refreshAgentStateNow = TS._refreshAgentStateNow,
    _txEtaTicker = TS._txEtaTicker,
    _summaryEtaTracker = TS._summaryEtaTracker,
    _citationsEtaTracker = TS._citationsEtaTracker,
    _frictionEtaTracker = TS._frictionEtaTracker,
    _updateAgentElapsed = TS._updateAgentElapsed,
    _currentParticipantHasTranscript = TS._currentParticipantHasTranscript;

  // ---- Thinking-agent plumbing (shared poll factory) ----
  // A new agent is one descriptor plus hooks.

  // Cap for citations + friction polls; long LLM runs outlived shorter caps. Summary streams uncapped.
  var _AGENT_POLL_TIMEOUT = 300000;

  var AGENT_DESCRIPTORS = {
    summary: {
      key: "summary",
      urlBase: "api/agent/summary",
      interval: 1200,
      timeout: null,
      _poller: null,
      getResult: function (d) { return d.summary; },
      onResult: function (pid, d) { _onSummaryResult(pid, d); },
      onGenerating: function (pid, d) {
        // Rebuild the generating box if a re-render dropped it, then push partial text.
        if (!qs("#summaryStream")) {
          renderSummaryGenerating(d.started_at ? d.started_at * 1000 : undefined);
        }
        if (d.partial) _updateSummaryStream(d.partial);
      },
      onEmpty: function () { renderSummaryEmpty(); },
      // Participant switch: stop silently; the panel now belongs to another participant.
      onStale: function () {},
    },
    citations: {
      key: "citations",
      urlBase: "api/agent/citations",
      interval: 3000,
      timeout: _AGENT_POLL_TIMEOUT,
      _poller: null,
      getResult: function (d) { return d.citations; },
      onResult: function (pid, d) {
        state.summaryCitations = d.citations;
        state.citationsGenerating = false;
        renderCitations();
      },
      onEmpty: function () {
        // Cancel cleared the flag but a GET was already in flight; not a failure.
        if (!state.citationsGenerating) return;
        // Run ended without persisting (route 404s). Say so; the cause is a suggestion.
        state.citationsGenerating = false;
        _renderCitationsNote("Couldn't find sources. Check that the AI server is running, then re-run citations.");
      },
      onStale: function () { _clearCitationsStatus(); },
    },
    friction: {
      key: "friction",
      urlBase: "api/agent/friction",
      interval: 3000,
      timeout: _AGENT_POLL_TIMEOUT,
      _poller: null,
      // The deterministic placeholder is not a finished run; let it fall through to onEmpty.
      getResult: function (d) { return d.friction && !d.friction.deterministic ? d.friction : null; },
      onResult: function (pid, d) { _setFrictionData(d.friction); },
      onEmpty: function (pid, d) {
        state.frictionGenerating = false;
        if (d && d.friction) _setFrictionData(d.friction);
        else renderFrictionEmpty();
      },
      onStale: function () { state.frictionGenerating = false; renderFriction(); },
    },
  };

  // runImmediately:false — the initial render already painted the box. Hooks carry per-agent behavior.
  function _makeAgentPoll(desc) {
    return function (pid) {
      _stopAgentPoll(desc);
      var started = Date.now();
      var ver = state.participantReqVer;
      desc._failStreak = 0;
      desc._poller = createPoller(function () {
        if (ver !== state.participantReqVer ||
            state.selectedParticipant !== pid ||
            (desc.timeout && Date.now() - started > desc.timeout)) {
          _stopAgentPoll(desc);
          desc.onStale(pid);
          return;
        }
        apiGet(desc.urlBase + "/" + pid).then(function (data) {
          if (ver !== state.participantReqVer) return;
          desc._failStreak = 0;
          if (data.error) reportAgentError(pid, desc.key, data.error);
          if (data.ok && desc.getResult(data)) {
            _stopAgentPoll(desc);
            desc.onResult(pid, data);
          } else if (data.generating) {
            if (desc.onGenerating) desc.onGenerating(pid, data);
          } else {
            _stopAgentPoll(desc);
            desc.onEmpty(pid, data);
          }
        }).catch(function (err) {
          if (ver !== state.participantReqVer) return;
          if (err && err.status === 404) {
            // 404: run ended with nothing persisted. serverMessage carries the failure reason.
            if (err.serverMessage) reportAgentError(pid, desc.key, err.serverMessage);
            _stopAgentPoll(desc);
            desc.onEmpty(pid);
            return;
          }
          // Transient failure: keep polling and the panel intact. Only a streak gives up, via onStale.
          desc._failStreak = (desc._failStreak || 0) + 1;
          if (desc._failStreak >= 3) {
            _stopAgentPoll(desc);
            desc.onStale(pid);
          }
        });
      }, desc.interval, { runImmediately: false, label: "transcripts.agent." + desc.key });
      desc._poller.start();
    };
  }

  // Timer-only. Resetting ETA trackers here would wipe the seed render*Generating just set.
  function _stopAgentPoll(desc) {
    if (desc._poller) {
      desc._poller.stop();
      desc._poller = null;
    }
  }

  var _startSummaryPoll = _makeAgentPoll(AGENT_DESCRIPTORS.summary);
  var _startCitationsPoll = _makeAgentPoll(AGENT_DESCRIPTORS.citations);
  var _startFrictionPoll = _makeAgentPoll(AGENT_DESCRIPTORS.friction);

  // ---- AI Summary ----
  // Summary polls while generating; citations poll once the summary lands.

  // SSE token stream is the primary transport; the summary poll is the fallback.
  var _summaryStream = null;
  // Painted participant. Drives clear-on-switch below; clearSummary itself never resets it.
  var _summaryPaintedPid = null;

  function loadSummary(pid) {
    var ver = state.participantReqVer;

    // Blank only on a real switch: the catch below keeps painted panes on transient failures.
    if (_summaryPaintedPid !== pid) {
      clearSummary();
      _summaryPaintedPid = pid;
    }

    // Reveal the panel here: render* never toggles #summarySection, so post-finalize runs would paint hidden.
    if (pid === state.selectedParticipant && _currentParticipantHasTranscript()) {
      _setAnalysisPanelVisible(true);
    }

    apiGet(AGENT_DESCRIPTORS.summary.urlBase + "/" + pid).then(function (data) {
      if (ver !== state.participantReqVer) return;
      if (data.ok && data.summary) {
        _stopSummaryPoll();
        _onSummaryResult(pid, data);
      } else if (data.generating) {
        renderSummaryGenerating(data.started_at ? data.started_at * 1000 : undefined);
        if (data.partial) _updateSummaryStream(data.partial);
        _startSummaryStream(pid);
        _refreshAgentStateNow();
      } else {
        renderSummaryEmpty();
      }
    }).catch(function (err) {
      if (ver !== state.participantReqVer) return;
      // Only a 404 blanks the panel; a transport blip leaves painted state alone.
      if (err && err.status === 404) renderSummaryEmpty();
    });
  }

  // Summary landed: render, then attach citations (poll if generating, else fetch stored).
  function _onSummaryResult(pid, data) {
    // Drop the previous participant's citations first; renderSummary reapplies state.summaryCitations.
    state.summaryCitations = null;
    state.citationsGenerating = false;
    renderSummary(data.summary);
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
    } else {
      _loadStoredCitations(pid);
    }
  }

  // The summary GET carries citations' status, not their payload; fetch them separately.
  function _loadStoredCitations(pid) {
    var ver = state.participantReqVer;
    apiGet(AGENT_DESCRIPTORS.citations.urlBase + "/" + pid).then(function (data) {
      if (ver !== state.participantReqVer || state.selectedParticipant !== pid) return;
      // A run may have started meanwhile; don't replace its live status with stale superscripts.
      if (state.citationsGenerating) return;
      if (!data.ok || !data.citations) return;
      state.summaryCitations = data.citations;
      renderCitations();
    }).catch(function () {
      // 404 = citations never ran. The poll's onEmpty owns the run-ended-empty message.
    });
  }

  // SSE token stream; falls back to the GET poll. {done} re-runs loadSummary.
  function _startSummaryStream(pid) {
    _stopSummaryPoll(); // clear any prior poller/stream before (re)starting
    var ver = state.participantReqVer;
    _summaryStream = createSSEStream(AGENT_DESCRIPTORS.summary.urlBase + "/" + pid + "/stream", {
      onMessage: function (data) {
        if (ver !== state.participantReqVer || state.selectedParticipant !== pid) {
          _stopSummaryStream();
          return;
        }
        if (data.done) {
          _stopSummaryStream();
          loadSummary(pid); // finalized summary + citation status in one GET
          return;
        }
        if (data.partial != null) {
          // Rebuild the generating box if a re-render dropped it, then stream in.
          if (!qs("#summaryStream")) renderSummaryGenerating();
          _updateSummaryStream(data.partial);
        }
      },
      onError: function () {
        // Transport dropped mid-run → degrade to the GET poll (unless we've
        // since navigated away).
        _summaryStream = null;
        if (ver === state.participantReqVer && state.selectedParticipant === pid) {
          _startSummaryPoll(pid);
        }
      },
      onUnsupported: function () {
        _startSummaryPoll(pid);
      },
    });
  }

  function _stopSummaryStream() {
    if (_summaryStream) {
      _summaryStream.close();
      _summaryStream = null;
    }
  }

  // bfcache would keep the EventSource open; the other teardown sites never fire on pagehide.
  window.addEventListener("pagehide", _stopSummaryStream);

  // Stops both transports so teardown sites need not know which is active.
  function _stopSummaryPoll() {
    _stopAgentPoll(AGENT_DESCRIPTORS.summary);
    _stopSummaryStream();
  }

  // startedAtMs: server run start (epoch ms) seeds the elapsed clock; omit for manual runs.
  function renderSummaryGenerating(startedAtMs) {
    var content = qs("#summaryContent");
    // #summaryStream is separate so partial text updates leave the clock / Cancel wiring alone.
    content.innerHTML =
      '<div class="summary-stream" id="summaryStream"></div>' +
      // Own span: .cg-shimmer's transparent text fill would erase the clock and Cancel.
      '<p class="summary-generating"><span class="cg-shimmer">Generating summary\u2026</span>' +
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
    // Friction's dependency is unmet while the summary regenerates; refresh so Re-run disables.
    _renderFrictionHeader();
  }

  // Plain textContent into #summaryStream; citation anchors arrive with renderSummary.
  function _updateSummaryStream(text) {
    var stream = qs("#summaryStream");
    if (stream) stream.textContent = text;
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
    _updateSummaryEmptyHint();
  }

  // Explain why Run would do nothing. Fire-and-forget: a slow /api/models must never delay or throw.
  function _updateSummaryEmptyHint() {
    var hintEl = qs("#summaryEmptyHint");
    if (!hintEl) return;
    hintEl.classList.add("hidden");
    hintEl.textContent = "";
    _trFetchModels().then(function (data) {
      // A stopped server is not worth a hint: running the summary starts it.
      var status = clipgenLlmStatus(data && data.llm);
      if (status.state !== "missing") return;
      // Running the summary raises the install dialog, so say just that.
      var extra = status.canInstall
        ? " clipgen can download it for you when you run the summary."
        : (status.hint.length ? " " + status.hint[0] : "");
      hintEl.textContent = status.message + extra;
      hintEl.classList.remove("hidden");
    }).catch(function () { /* leave the plain empty state alone */ });
  }

  function renderSummary(text) {
    _summaryEtaTracker.reset();
    state.summaryText = text;
    state.summaryEditing = false;
    var content = qs("#summaryContent");
    var lines = text.split("\n");

    // Must match thinking_agents._split_summary_sentences exactly: data-cite-index is the backend's claim numbering.
    var blocks = [];
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;
      var isBullet = line.indexOf("- ") === 0 || line.indexOf("* ") === 0;
      var kind = isBullet ? "ul" : "p";
      if (!blocks.length || blocks[blocks.length - 1].kind !== kind) {
        blocks.push({ kind: kind, items: [] });
      }
      var items = blocks[blocks.length - 1].items;
      // Render **bold** / `code` instead of literal asterisks (same helper as Overview).
      if (isBullet) {
        items.push(clipgenRenderInlineMarkdown(line.substring(2).trim()));
      } else {
        // Split prose into individual sentences for citation targeting.
        var parts = line.split(/(?<=[.!?])\s+/);
        for (var k = 0; k < parts.length; k++) {
          var part = parts[k].trim();
          if (part) items.push(clipgenRenderInlineMarkdown(part));
        }
      }
    }

    // Build HTML with data-cite-index on each sentence/bullet
    var citeIdx = 0;
    var html = "";
    for (var b = 0; b < blocks.length; b++) {
      var block = blocks[b];
      if (!block.items.length) continue;
      if (block.kind === "p") {
        html += "<p>";
        for (var si = 0; si < block.items.length; si++) {
          if (si > 0) html += " ";
          html += '<span data-cite-index="' + citeIdx + '">' + block.items[si] + "</span>";
          citeIdx++;
        }
        html += "</p>";
      } else {
        html += "<ul>";
        for (var j = 0; j < block.items.length; j++) {
          html += '<li data-cite-index="' + citeIdx + '">' + block.items[j] + "</li>";
          citeIdx++;
        }
        html += "</ul>";
      }
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
    // Summary landed, so the friction dependency is met; refresh even when loadFriction is skipped.
    _renderFrictionHeader();
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
    if (pid) setStoredUIMapEntry("transcripts", "tabByParticipant", pid, name);
    var isSummary = name === "summary";
    qs("#tabBtnSummary").classList.toggle("active", isSummary);
    qs("#tabBtnSummary").setAttribute("aria-selected", isSummary ? "true" : "false");
    qs("#tabBtnFriction").classList.toggle("active", !isSummary);
    qs("#tabBtnFriction").setAttribute("aria-selected", !isSummary ? "true" : "false");
    qs("#summaryTab").classList.toggle("hidden", !isSummary);
    qs("#frictionTab").classList.toggle("hidden", isSummary);
  }

  function _restoreActiveTab(pid) {
    var saved =
      getStoredUIMapEntry("transcripts", "tabByParticipant", pid) === "friction"
        ? "friction"
        : "summary";
    selectTab(saved);
  }

  function initPanelTabs() {
    qs("#tabBtnSummary").addEventListener("click", function () { selectTab("summary"); });
    qs("#tabBtnFriction").addEventListener("click", function () { selectTab("friction"); });
    qs("#summaryRunCta").addEventListener("click", function () { _startSummaryRun(); });
  }

  // ---- Citation rendering (Pass 2) ----

  // Shared by onEmpty/onStale, cancel and regenerate-error: drop the flag and status line.
  function _clearCitationsStatus() {
    state.citationsGenerating = false;
    var status = qs("#summaryContent .citations-status");
    if (status) status.remove();
  }

  function renderCitations() {
    // Remove any existing status text
    var existing = qs("#summaryContent .citations-status");
    if (existing) existing.remove();

    // Remove any previously rendered citation links
    var oldLinks = qs("#summaryContent").querySelectorAll(".citation-link");
    for (var r = 0; r < oldLinks.length; r++) oldLinks[r].remove();

    if (!state.summaryCitations) return;

    var dataRefs = 0;
    var refNum = 1;
    for (var i = 0; i < state.summaryCitations.length; i++) {
      var cite = state.summaryCitations[i];
      dataRefs += (cite.refs || []).length;
      if (!cite.refs || cite.refs.length === 0) continue;
      var el = qs('#summaryContent [data-cite-index="' + i + '"]');
      if (!el) continue;
      for (var j = 0; j < cite.refs.length; j++) {
        var ref = cite.refs[j];
        var sup = document.createElement("sup");
        sup.className = "citation-link";
        sup.dataset.start = String(ref.start);
        sup.setAttribute("data-tooltip", formatTime(ref.start));
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

    // Nothing rendered is a result. dataRefs > 0 means the two sentence splits disagreed.
    if (refNum === 1) {
      _renderCitationsNote(dataRefs === 0
        ? "No supporting segments found for this summary."
        : "Sources were found but couldn't be matched to the summary text. Re-run citations.");
    }
  }

  // Terminal status line; shares .citations-status with the in-flight line, minus clock and Cancel.
  function _renderCitationsNote(text) {
    var existing = qs("#summaryContent .citations-status");
    if (existing) existing.remove();
    var p = document.createElement("p");
    p.className = "citations-status";
    p.textContent = text;
    qs("#summaryContent").appendChild(p);
  }

  // startedAtMs: server run start (epoch ms) seeds the elapsed clock; omit for manual runs.
  function renderCitationsStatus(startedAtMs) {
    // Remove any existing status
    var existing = qs("#summaryContent .citations-status");
    if (existing) existing.remove();

    var p = document.createElement("p");
    p.className = "citations-status";
    // Own span: the shimmer's transparent fill would blank the clock and Cancel.
    var label = document.createElement("span");
    label.className = "cg-shimmer";
    label.textContent = "Finding sources\u2026";
    p.appendChild(label);
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
    // Reset before seeding so a re-run adopts the new start; _stopCitationsPoll is timer-only.
    _citationsEtaTracker.reset();
    _citationsEtaTracker.start(startedAtMs || undefined);
    _updateAgentElapsed("citationsElapsed", _citationsEtaTracker);
    _txEtaTicker.ensure();
  }

  function _stopCitationsPoll() {
    _stopAgentPoll(AGENT_DESCRIPTORS.citations);
  }

  function _stopCitationsRun() {
    var pid = state.selectedParticipant;
    if (!pid) return;
    // Keep the summary visible; only remove the status line.
    _stopCitationsPoll();
    _clearCitationsStatus();
    apiPost(AGENT_DESCRIPTORS.citations.urlBase + "/" + pid + "/stop", {}).then(function () {
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
      apiPost(AGENT_DESCRIPTORS.summary.urlBase + "/" + pid + "/regenerate", {}).then(function (data) {
        if (data.ok && data.generating) {
          _startSummaryStream(pid);
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
    // The chain may already be on citations: stop both, then re-sync only after both acknowledge.
    Promise.all([
      apiPost(AGENT_DESCRIPTORS.summary.urlBase + "/" + pid + "/stop", {}),
      apiPost(AGENT_DESCRIPTORS.citations.urlBase + "/" + pid + "/stop", {}),
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
      btn.setAttribute("data-tooltip", "Save summary");
      btn.setAttribute("aria-label", "Save summary");
    } else {
      icon.classList.remove("summary-action-save");
      icon.classList.add("summary-action-edit");
      btn.setAttribute("data-tooltip", "Edit summary");
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
        apiPost(AGENT_DESCRIPTORS.citations.urlBase + "/" + pid + "/regenerate", {}).then(function (data) {
          if (data.ok && data.generating) {
            _startCitationsPoll(pid);
          }
        }).catch(function () {
          showToast("Failed to regenerate citations");
          _clearCitationsStatus();
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
        apiPut(AGENT_DESCRIPTORS.summary.urlBase + "/" + pid, { summary: newText }).then(function () {
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

  // ---- Friction detection ----
  // A control surface over #segmentList; everything reads one derived map (_recomputeFrictionMatches).

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
    // state.summaryText lands before /api/participants catches up, so trust it too.
    if (state.summaryText) return true;
    if (p.has_summary) return true;
    return !!(p.agents && p.agents.summary === "done");
  }

  function _frictionCatLabel(key) {
    var cats = CLIPGEN_CONFIG.frictionCategories || [];
    for (var i = 0; i < cats.length; i++) {
      if (cats[i].key === key) return cats[i].label;
    }
    if (key === "other") return "Other";
    // A model-invented category: show its own wording (the evidence table groups it under Other).
    if (!key) return "—";
    return key.charAt(0).toUpperCase() + key.slice(1).replace(/_/g, " ");
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
    var ver = state.participantReqVer;
    // Reveal the panel (see loadSummary) so a post-finalize friction run is visible.
    if (pid === state.selectedParticipant && _currentParticipantHasTranscript()) {
      _setAnalysisPanelVisible(true);
    }
    // Blank DOM and state only on a real switch: same-participant refetches must keep programmatic scores.
    if (state.frictionPid !== pid) {
      clearFriction();
    }
    state.frictionPid = pid;
    state.frictionGenerating = false;
    apiGet(AGENT_DESCRIPTORS.friction.urlBase + "/" + pid).then(function (data) {
      if (ver !== state.participantReqVer) return;
      if (data.ok && data.friction) {
        _setFrictionData(data.friction);
      } else if (data.generating) {
        // Mid-run the server sends deterministic scores alongside `generating`; adopt them before flipping the flag.
        if (data.friction) _setFrictionData(data.friction);
        state.frictionGenerating = true;
        state.frictionStartedAt = data.started_at ? data.started_at * 1000 : null;
        renderFrictionGenerating();
        _startFrictionPoll(pid);
        _refreshAgentStateNow();
      } else {
        renderFrictionEmpty();
      }
    }).catch(function (err) {
      if (ver !== state.participantReqVer) return;
      // Same contract as loadSummary: only a 404 blanks the panel.
      if (err && err.status === 404) renderFrictionEmpty();
    });
  }

  function _setFrictionData(friction) {
    state.frictionData = friction;
    state.frictionGenerating = false;
    state.frictionMomentIndex = -1;
    var byId = {};
    var segs = (friction && friction.segments) || [];
    for (var i = 0; i < segs.length; i++) {
      if (segs[i] && segs[i].id) byId[segs[i].id] = segs[i];
    }
    state.frictionBySegId = byId;
    renderFriction();
    updateFrictionStaleDot();
    // renderFriction already re-decorated rows in place; only the canvas band needs redrawing.
    renderTimeline();
  }

  function clearFriction() {
    _stopFrictionPoll();
    state.frictionData = null;
    state.frictionBySegId = {};
    state.frictionGenerating = false;
    state.frictionMomentIndex = -1;
    state.frictionPid = null;
    qs("#frictionContent").classList.add("hidden");
    qs("#frictionGenerating").classList.add("hidden");
    qs("#frictionEmpty").classList.add("hidden");
    updateFrictionStaleDot();
    // Strip tints / hidden rows / callouts left over from the previous participant.
    applyFrictionDecorations();
    renderTimeline();
  }

  function _renderFrictionHeader() {
    var statusEl = qs("#frictionStatus");
    var rerun = qs("#frictionRerun");
    var cancel = qs("#frictionCancel");
    if (state.frictionGenerating) {
      // Own span: the shimmer's transparent fill would blank the elapsed clock too.
      statusEl.textContent = "";
      var label = document.createElement("span");
      label.className = "cg-shimmer";
      label.textContent = "Analyzing friction…";
      statusEl.appendChild(label);
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
    // Deterministic-only reads "Run". Write the label span: textContent on the button deletes .ai-agent-badge.
    var isDeterministic = !!(state.frictionData && state.frictionData.deterministic);
    rerun.querySelector(".agent-run-label").textContent =
      state.frictionData && !isDeterministic ? "Re-run friction" : "Run friction analysis";
    var depMet = _frictionDepMet();
    if (depMet) {
      rerun.removeAttribute("disabled");
      rerun.setAttribute("data-tooltip", "Runs a local AI thinking agent on this transcript");
    } else {
      rerun.setAttribute("disabled", "disabled");
      rerun.setAttribute("data-tooltip", "Requires a summary first");
    }
    if (state.frictionData) {
      var fd = state.frictionData;
      var llmFailed = fd.llm_ok === false;
      if (isDeterministic) {
        // depMet picks the step the user can take: run Summary, or run friction.
        statusEl.textContent = depMet
          ? "Programmatic scores shown. Run friction analysis for AI-refined moments."
          : "Programmatic scores shown. Run Summary for AI-refined moments.";
      } else if (llmFailed) {
        statusEl.textContent =
          "Moment detection failed: model unavailable" +
          (fd.model ? " (tried " + fd.model + ")" : "") +
          ". Showing programmatic scores; re-run with an installed model.";
      } else if (fd.stale) {
        statusEl.textContent = "Stale: segments edited since last run" +
          (fd.model ? " · " + fd.model : "");
      } else if (!(fd.moments && fd.moments.length)) {
        // A completed run that found nothing; otherwise the header reads like any success.
        statusEl.textContent = "No friction moments found · computed " +
          _friendlyTimeAgo(fd.computed_at) + (fd.model ? " · " + fd.model : "");
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
      var banner = qs("#frictionStaleBanner");
      if (banner) {
        var fd = state.frictionData;
        // Banner only for edit-staleness; the header already explains LLM failure.
        banner.classList.toggle("hidden", !(fd.stale && fd.llm_ok !== false));
      }
      // Bins change only with the data; everything else follows from applyFrictionDecorations.
      renderFrictionHistogram();
      applyFrictionDecorations();
    } else {
      renderFrictionEmpty();
    }
  }

  function renderFrictionEmpty() {
    state.frictionGenerating = false;
    qs("#frictionContent").classList.add("hidden");
    qs("#frictionGenerating").classList.add("hidden");
    qs("#frictionEmpty").classList.remove("hidden");
    applyFrictionDecorations();
    qs("#frictionEmptyHint").textContent = _frictionDepMet()
      ? "Run the analysis to surface moments of likely friction."
      : "Requires a summary first. Run Summary, then friction.";
    _renderFrictionHeader();
    updateFrictionStaleDot();
  }

  function renderFrictionGenerating() {
    // Keep programmatic scores on screen during the run; the agent only adds moments.
    var hasData = !!state.frictionData;
    qs("#frictionContent").classList.toggle("hidden", !hasData);
    qs("#frictionEmpty").classList.add("hidden");
    qs("#frictionGenerating").classList.toggle("hidden", hasData);
    if (hasData) {
      renderFrictionHistogram();
      applyFrictionDecorations();
    }
    _renderFrictionHeader();
  }

  function updateFrictionStaleDot() {
    var dot = qs("#frictionStaleDot");
    if (!dot) return;
    dot.classList.toggle("hidden", !(state.frictionData && state.frictionData.stale));
  }

  function _stopFrictionPoll() {
    _stopAgentPoll(AGENT_DESCRIPTORS.friction);
  }

  function _startFrictionRun() {
    var pid = state.selectedParticipant;
    if (!pid || !_frictionDepMet()) return;
    ensureAgentModelInstalled("friction").then(function (ok) {
      if (!ok) return;
      state.frictionGenerating = true;
      state.frictionStartedAt = null;
      renderFrictionGenerating();
      apiPost(AGENT_DESCRIPTORS.friction.urlBase + "/" + pid + "/regenerate", {}).then(function (data) {
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
    apiPost(AGENT_DESCRIPTORS.friction.urlBase + "/" + pid + "/stop", {}).then(function () {
      _refreshAgentStateNow();
      loadFriction(pid);
    }).catch(function () {});
    renderFriction();
  }

  // Two filters, one per source; one shared dict hid categories on both sides.
  function _ensureFrictionFilter() {
    var keys = _frictionCatKeys();
    var prog = state.frictionCategoryFilter || {};
    var ai = state.frictionMomentFilter || {};
    // Fill in, don't replace: persisted filters keep user choices and pick up new categories.
    for (var i = 0; i < keys.length; i++) {
      if (prog[keys[i]] === undefined) prog[keys[i]] = true;
      if (ai[keys[i]] === undefined) ai[keys[i]] = true;
    }
    if (ai[FRICTION_OTHER] === undefined) ai[FRICTION_OTHER] = true;
    state.frictionCategoryFilter = prog;
    state.frictionMomentFilter = ai;
  }

  // Both bounds are user-controlled, so every score test goes through here.
  function _frictionScoreInBand(score) {
    return score >= state.frictionMin && score <= state.frictionMax;
  }

  // ---- The one derived filter product ----
  // Every consumer reads these; pane and transcript agree.
  function _recomputeFrictionMatches() {
    _ensureFrictionFilter();
    var fd = state.frictionData;
    var segs = (fd && fd.segments) || [];
    var filter = state.frictionCategoryFilter;
    var matches = {};
    var i, j;
    for (i = 0; i < segs.length; i++) {
      var frow = segs[i];
      var score = frow.score || 0;
      // Exclude score <= 0 so a lower bound of 0 still means "flagged segments only".
      if (score <= 0 || !_frictionScoreInBand(score)) continue;
      var cats = frow.categories || [];
      for (j = 0; j < cats.length; j++) {
        if (filter[cats[j]] !== false) { matches[frow.id] = score; break; }
      }
    }
    state.frictionMatchBySegId = matches;

    // Resolve moment indices once so strip numbering and callouts never drift apart.
    var visible = [];
    var cited = {};
    var unsourced = 0;
    var all = (fd && fd.moments) || [];
    for (i = 0; i < all.length; i++) {
      if (!_frictionMomentMatches(all[i])) continue;
      var idxs = _momentSegmentIndices(all[i]);
      // Unsourced moment: counted so the empty strip can name the cause.
      if (idxs.length === 0) { unsourced++; continue; }
      visible.push({ moment: all[i], idxs: idxs });
      for (j = 0; j < idxs.length; j++) {
        var seg = state.segments[idxs[j]];
        if (seg) cited[seg.id] = visible.length; // 1-based strip number
      }
    }
    state.frictionVisibleMoments = visible;
    state.frictionCitedBySegId = cited;
    state.frictionUnsourcedMoments = unsourced;

    // Band = union of both sources; without AI-only lines, isolate would hide cited rows.
    var band = {};
    for (var id in matches) {
      if (Object.prototype.hasOwnProperty.call(matches, id)) band[id] = matches[id];
    }
    for (i = 0; i < visible.length; i++) {
      var mscore = visible[i].moment.score || 0;
      for (j = 0; j < visible[i].idxs.length; j++) {
        var cseg = state.segments[visible[i].idxs[j]];
        if (!cseg) continue;
        if (!(band[cseg.id] > mscore)) band[cseg.id] = mscore;
      }
    }
    state.frictionBandBySegId = band;
  }

  // The one entry point for filter changes. Per-frame cheap: writes classes only, never rebuilds rows.
  function applyFrictionDecorations() {
    _recomputeFrictionMatches();
    renderFrictionEvidence();
    _updateFrictionBounds(null);
    renderFrictionJumpStrip();
    _decorateSegmentList();
  }

  // ---- "Why was this selected" ----
  // One builder so bin, row and band hovers agree.

  var _FRICTION_TOOLTIP_SEGMENTS = 4; // per histogram bin, before "+N more"

  function _frictionQuote(seg, maxChars) {
    var text = (seg && seg.text ? seg.text : "").trim();
    if (!text) return "";
    if (text.length > maxChars) text = text.slice(0, maxChars - 1).replace(/\s+\S*$/, "") + "…";
    return "“" + text + "”";
  }

  // One compact line: 0:12 · 0.45 · Confusion · “where is the export button”
  function _frictionWhyLine(frow, seg) {
    var parts = [formatTime(seg.start), (frow.score || 0).toFixed(2)];
    var cats = (frow.categories || []).map(_frictionCatLabel);
    if (cats.length) parts.push(cats.join(", "));
    var quote = _frictionQuote(seg, 60);
    if (quote) parts.push(quote);
    return parts.join(" · ");
  }

  // Scored segments whose score falls in [lo, hi) — the top bin takes 1.0 too.
  function _frictionSegmentsInBin(lo, hi) {
    var rows = (state.frictionData && state.frictionData.segments) || [];
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      var sc = rows[i].score || 0;
      if (sc <= 0) continue;
      if (sc < lo) continue;
      if (hi < 1 ? sc >= hi : sc > hi) continue;
      out.push(rows[i]);
    }
    return out;
  }

  function _frictionBinTooltip(lo, hi) {
    var rows = _frictionSegmentsInBin(lo, hi);
    var head = "Score " + lo.toFixed(2) + "–" + hi.toFixed(2) + " · " +
      clipgenPluralUnit(rows.length, "segment", "segments");
    if (!rows.length) return head;
    var lines = [head];
    // Resolve text only for the handful shown.
    for (var i = 0; i < rows.length && i < _FRICTION_TOOLTIP_SEGMENTS; i++) {
      var idx = _segmentIndexById(rows[i].id);
      var seg = idx >= 0 ? state.segments[idx] : null;
      lines.push(seg ? _frictionWhyLine(rows[i], seg)
        : formatTime(0) + " · " + (rows[i].score || 0).toFixed(2));
    }
    if (rows.length > _FRICTION_TOOLTIP_SEGMENTS) {
      lines.push("+" + (rows.length - _FRICTION_TOOLTIP_SEGMENTS) + " more");
    }
    return lines.join("\n");
  }

  // ---- Score histogram (the score-band control) ----

  var _FRICTION_HIST_BINS = 10;

  function renderFrictionHistogram() {
    var host = qs("#frictionHistogram");
    if (!host) return;
    host.innerHTML = "";
    var segs = (state.frictionData && state.frictionData.segments) || [];
    var counts = [];
    var i;
    for (i = 0; i < _FRICTION_HIST_BINS; i++) counts[i] = 0;
    var scored = 0;
    for (i = 0; i < segs.length; i++) {
      var s = Number(segs[i].score) || 0;
      // Only scored segments are binned; most score exactly 0 and would spike bin 0.
      if (s <= 0) continue;
      if (s > 1) s = 1;
      scored++;
      counts[Math.min(_FRICTION_HIST_BINS - 1, Math.floor(s * _FRICTION_HIST_BINS))]++;
    }
    var maxCount = 0;
    for (i = 0; i < _FRICTION_HIST_BINS; i++) if (counts[i] > maxCount) maxCount = counts[i];
    for (i = 0; i < _FRICTION_HIST_BINS; i++) {
      var n = counts[i];
      var pct = maxCount > 0 ? (n / maxCount) * 100 : 0;
      if (n > 0 && pct < 6) pct = 6; // floor so a single-segment bin stays visible
      var bar = el("div", "friction-hist-bar");
      var fill = el("div", "friction-hist-bar-fill");
      fill.style.height = pct.toFixed(1) + "%";
      bar.appendChild(fill);
      (function (lo, hi) {
        // Built lazily on hover: the transcript text may not have landed yet.
        attachHoverTooltip(bar, function () { return _frictionBinTooltip(lo, hi); },
          { multiline: true });
      })(i / _FRICTION_HIST_BINS, (i + 1) / _FRICTION_HIST_BINS);
      host.appendChild(bar);
    }
    host.appendChild(el("span", "friction-hist-label",
      clipgenPluralUnit(scored, "scored segment", "scored segments")));
    host.appendChild(_buildFrictionHandle("min"));
    host.appendChild(_buildFrictionHandle("max"));

    var bounds = qs("#frictionBounds");
    if (bounds) {
      bounds.innerHTML = "";
      bounds.appendChild(_buildFrictionBoundLabel("min"));
      bounds.appendChild(_buildFrictionBoundLabel("max"));
    }
  }

  function _buildFrictionHandle(which) {
    var m = el("div", "friction-hist-marker");
    m.setAttribute("data-bound", which);
    return m;
  }

  function _buildFrictionBoundLabel(which) {
    var s = el("span", "friction-hist-bound");
    s.setAttribute("data-bound", which);
    return s;
  }

  // Split from renderFrictionHistogram: rewriting innerHTML mid-drag would drop the pointer capture.
  function _updateFrictionBounds(activeBound) {
    var host = qs("#frictionHistogram");
    if (!host) return;
    var lo = state.frictionMin;
    var hi = state.frictionMax;
    var markers = host.querySelectorAll(".friction-hist-marker");
    for (var m = 0; m < markers.length; m++) {
      var which = markers[m].getAttribute("data-bound");
      markers[m].style.left = ((which === "min" ? lo : hi) * 100) + "%";
      markers[m].classList.toggle("is-active", which === activeBound);
    }
    var bars = host.querySelectorAll(".friction-hist-bar");
    for (var i = 0; i < bars.length; i++) {
      // A bar is outside when its whole range falls beyond either bound.
      var barLo = i / _FRICTION_HIST_BINS;
      var barHi = (i + 1) / _FRICTION_HIST_BINS;
      bars[i].classList.toggle("is-outside", barHi <= lo || barLo >= hi);
    }

    // Nudge near-coincident labels apart so both numbers stay readable.
    var loPct = lo * 100;
    var hiPct = hi * 100;
    if (hiPct - loPct < 8) {
      var mid = (loPct + hiPct) / 2;
      loPct = Math.max(0, mid - 4);
      hiPct = Math.min(100, mid + 4);
    }
    var labels = (qs("#frictionBounds") || host).querySelectorAll(".friction-hist-bound");
    for (var j = 0; j < labels.length; j++) {
      var isMin = labels[j].getAttribute("data-bound") === "min";
      labels[j].style.left = (isMin ? loPct : hiPct) + "%";
      labels[j].textContent = (isMin ? lo : hi).toFixed(2);
      labels[j].classList.toggle("is-active", labels[j].getAttribute("data-bound") === activeBound);
    }
  }

  function _initFrictionHistogramDrag() {
    var host = qs("#frictionHistogram");
    if (!host) return;
    var raf = 0;
    var pending = null;
    var grabbed = null; // "min" | "max" — which handle this gesture moves

    function scoreAt(clientX) {
      var r = host.getBoundingClientRect();
      if (!r.width) return null;
      // 0.05 steps, matching the range input this replaced.
      return Math.max(0, Math.min(1, Math.round(((clientX - r.left) / r.width) * 20) / 20));
    }

    function commit(v) {
      if (v === null || !grabbed) return;
      // Bounds clamp against each other; dragging past collapses the band, never inverts it.
      if (grabbed === "min") v = Math.min(v, state.frictionMax);
      else v = Math.max(v, state.frictionMin);
      var key = grabbed === "min" ? "frictionMin" : "frictionMax";
      if (state[key] === v) {
        _updateFrictionBounds(grabbed); // still reflect the grabbed handle
        return;
      }
      state[key] = v;
      applyFrictionDecorations();
      _updateFrictionBounds(grabbed);
      renderTimeline();
    }

    // Capture is optional (WKWebView may refuse), so gate on `grabbed`; buttons===0 heals a missed pointerup.
    host.addEventListener("pointerdown", function (e) {
      var v = scoreAt(e.clientX);
      if (v === null) return;
      // Outside the band, grab that side: a collapsed band (min === max) must stay recoverable.
      if (v > state.frictionMax) grabbed = "max";
      else if (v < state.frictionMin) grabbed = "min";
      else grabbed = Math.abs(v - state.frictionMin) <= Math.abs(v - state.frictionMax) ? "min" : "max";
      try { host.setPointerCapture(e.pointerId); } catch (err) { /* capture is optional */ }
      pending = v;
      commit(v);
      e.preventDefault();
    });
    host.addEventListener("pointermove", function (e) {
      if (!grabbed) return;
      if (e.buttons === 0) { endDrag(e); return; }
      pending = scoreAt(e.clientX);
      if (raf) return;
      raf = requestAnimationFrame(function () { raf = 0; commit(pending); });
    });
    function endDrag(e) {
      if (!grabbed) return;
      try { host.releasePointerCapture(e.pointerId); } catch (err) { /* never captured */ }
      if (raf) { cancelAnimationFrame(raf); raf = 0; }
      commit(pending);
      grabbed = null;
      _updateFrictionBounds(null);
      // Persist once per gesture; the stored UI state re-serializes the whole page blob.
      setStoredUIStateField("transcripts", "frictionMin", state.frictionMin);
      setStoredUIStateField("transcripts", "frictionMax", state.frictionMax);
    }
    host.addEventListener("pointerup", endDrag);
    host.addEventListener("pointercancel", endDrag);
  }

  // ---- Evidence table ----
  // Keyword scorer labels segments, the agent labels moments; count them apart.
  var FRICTION_SOURCES = ["prog", "ai"];
  // Bucket for model categories outside the six; chips and callouts keep the model's wording.
  var FRICTION_OTHER = "other";

  function _frictionCatKeys() {
    var cats = CLIPGEN_CONFIG.frictionCategories || [];
    return cats.map(function (c) { return c.key; });
  }

  // The model may emit any string; bucket unknowns so they don't escape every filter.
  function _frictionMomentCategory(m) {
    var raw = (m && m.category) || "";
    return _frictionCatKeys().indexOf(raw) === -1 ? FRICTION_OTHER : raw;
  }

  // Counted here, not from stats.by_category: that counts marker hits and ignores the band.
  function _frictionEvidenceCounts() {
    var prog = {};
    var ai = {};
    var segs = (state.frictionData && state.frictionData.segments) || [];
    var i, j;
    for (i = 0; i < segs.length; i++) {
      var sc = segs[i].score || 0;
      if (sc <= 0 || !_frictionScoreInBand(sc)) continue;
      var cats = segs[i].categories || [];
      for (j = 0; j < cats.length; j++) prog[cats[j]] = (prog[cats[j]] || 0) + 1;
    }
    var moments = (state.frictionData && state.frictionData.moments) || [];
    for (i = 0; i < moments.length; i++) {
      if (!_frictionScoreInBand(moments[i].score || 0)) continue;
      var key = _frictionMomentCategory(moments[i]);
      ai[key] = (ai[key] || 0) + 1;
    }
    return { prog: prog, ai: ai };
  }

  // Band-independent on purpose: rows must not jump mid-drag, and banded-to-0 cells stay clickable.
  function _frictionEvidenceRows() {
    var seen = {};
    var segs = (state.frictionData && state.frictionData.segments) || [];
    var i, j;
    for (i = 0; i < segs.length; i++) {
      if ((segs[i].score || 0) <= 0) continue;
      var cats = segs[i].categories || [];
      for (j = 0; j < cats.length; j++) {
        if (!seen[cats[j]]) seen[cats[j]] = { prog: 0, ai: 0 };
        seen[cats[j]].prog++;
      }
    }
    var moments = (state.frictionData && state.frictionData.moments) || [];
    for (i = 0; i < moments.length; i++) {
      var key = _frictionMomentCategory(moments[i]);
      if (!seen[key]) seen[key] = { prog: 0, ai: 0 };
      seen[key].ai++;
    }
    // Config order first so the table reads consistently, then Other last.
    var out = [];
    var keys = _frictionCatKeys();
    for (i = 0; i < keys.length; i++) {
      if (seen[keys[i]]) out.push({ key: keys[i], totals: seen[keys[i]] });
    }
    if (seen[FRICTION_OTHER]) {
      out.push({ key: FRICTION_OTHER, totals: seen[FRICTION_OTHER] });
    }
    return out;
  }

  function _frictionSourceFilter(source) {
    return source === "ai" ? state.frictionMomentFilter : state.frictionCategoryFilter;
  }

  function _buildFrictionEvidenceCell(catKey, source) {
    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "friction-ev-cell";
    btn.setAttribute("data-cat", catKey);
    btn.setAttribute("data-src", source);
    btn.addEventListener("click", function () {
      if (btn.classList.contains("is-empty")) return;
      _ensureFrictionFilter();
      var f = _frictionSourceFilter(source);
      f[catKey] = f[catKey] === false;
      setStoredUIStateField(
        "transcripts",
        source === "ai" ? "frictionMomentFilter" : "frictionCategoryFilter",
        f
      );
      applyFrictionDecorations();
      renderTimeline();
    });
    return btn;
  }

  var FRICTION_SOURCE_META = {
    prog: {
      label: "Keyword",
      tooltip:
        "Segments the keyword scorer matched. Scored 0–1 by marker density, and " +
        "filtered by the histogram band above.",
    },
    ai: {
      label: "AI",
      tooltip:
        "Moments the local AI flagged. Scored 0–1 by the model's own confidence — " +
        "a different scale from the keyword score, filtered by the same band.",
    },
  };

  // Sources as rows, categories as columns: the sources are what's compared.
  function _buildFrictionEvidenceRow(source, cats) {
    var row = el("div", "friction-ev-row");
    row.setAttribute("data-src", source);
    var meta = FRICTION_SOURCE_META[source];
    var label = el("span", "friction-ev-src", meta.label);
    label.setAttribute("data-tooltip", meta.tooltip);
    row.appendChild(label);
    for (var i = 0; i < cats.length; i++) {
      row.appendChild(_buildFrictionEvidenceCell(cats[i], source));
    }
    return row;
  }

  function _buildFrictionEvidenceHead(cats) {
    var head = el("div", "friction-ev-row friction-ev-head");
    head.appendChild(el("span", "friction-ev-src", ""));
    for (var i = 0; i < cats.length; i++) {
      var label = _frictionCatLabel(cats[i]);
      var col = el("span", "friction-ev-col", label);
      // The header truncates on narrow panes; the tooltip keeps the full name.
      col.setAttribute("data-tooltip", label);
      head.appendChild(col);
    }
    return head;
  }

  function renderFrictionEvidence() {
    var host = qs("#frictionEvidence");
    if (!host) return;
    _ensureFrictionFilter();
    var rows = _frictionEvidenceRows();
    var cats = rows.map(function (r) { return r.key; });
    // Rebuild only when the category set changes; drags just rewrite counts and classes.
    var signature = cats.join(",");
    if (host.getAttribute("data-cats") !== signature) {
      host.setAttribute("data-cats", signature);
      host.innerHTML = "";
      if (cats.length === 0) {
        host.appendChild(el("span", "friction-ev-empty", "No findings in this analysis."));
        return;
      }
      var frag = document.createDocumentFragment();
      frag.appendChild(_buildFrictionEvidenceHead(cats));
      for (var i = 0; i < FRICTION_SOURCES.length; i++) {
        frag.appendChild(_buildFrictionEvidenceRow(FRICTION_SOURCES[i], cats));
      }
      host.appendChild(frag);
    }
    if (cats.length === 0) return;

    var counts = _frictionEvidenceCounts();
    var byKey = {};
    for (var r = 0; r < rows.length; r++) byKey[rows[r].key] = rows[r].totals;
    var cells = host.querySelectorAll(".friction-ev-cell");
    for (var c = 0; c < cells.length; c++) {
      var cell = cells[c];
      var key = cell.getAttribute("data-cat");
      var source = cell.getAttribute("data-src");
      var totals = byKey[key] || { prog: 0, ai: 0 };
      // Inert only with no findings at all; a banded-to-0 cell must stay clickable.
      var everyAny = totals[source] > 0;
      var n = (counts[source][key] || 0);
      var on = _frictionSourceFilter(source)[key] !== false;
      cell.textContent = everyAny ? String(n) : "—";
      cell.classList.toggle("is-empty", !everyAny);
      cell.classList.toggle("is-on", everyAny && on);
      cell.classList.toggle("is-muted", everyAny && !on);
      cell.setAttribute("aria-pressed", everyAny && on ? "true" : "false");
      cell.setAttribute(
        "aria-label",
        _frictionCatLabel(key) + " " + (source === "ai" ? "AI" : "keyword") + " findings"
      );
    }
  }

  function _frictionMomentMatches(m) {
    if (!_frictionScoreInBand(m.score || 0)) return false;
    var f = state.frictionMomentFilter;
    return !(f && f[_frictionMomentCategory(m)] === false);
  }

  // id->index map, rebuilt when segments are replaced; the per-frame drag recompute needs dict hits.
  var _segIndexMap = null;
  var _segIndexMapFor = null;

  function _segmentIndexById(id) {
    if (_segIndexMapFor !== state.segments) {
      _segIndexMap = {};
      for (var i = 0; i < state.segments.length; i++) {
        // First occurrence wins, matching the scan this replaced.
        if (!(state.segments[i].id in _segIndexMap)) {
          _segIndexMap[state.segments[i].id] = i;
        }
      }
      _segIndexMapFor = state.segments;
    }
    var idx = _segIndexMap[id];
    return idx === undefined ? -1 : idx;
  }

  function _seekToSegmentIndex(idx) {
    var seg = state.segments[idx];
    if (!seg) return;
    seekVideo(seg.start);
    if (!state.cachedSegmentRows) {
      state.cachedSegmentRows = qs("#segmentList").querySelectorAll(".segment-row");
    }
    var row = state.cachedSegmentRows[idx];
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

  // ---- Moment jump strip ----
  // Moments are navigation: one chip each; rationales become callouts.

  function _frictionJumpEmptyText() {
    var fd = state.frictionData;
    // The pane stays up mid-run, so the strip must cover the in-flight case.
    if (state.frictionGenerating) return "Analyzing friction…";
    if (!fd) return "No moments detected.";
    if (fd.deterministic) {
      return _frictionDepMet()
        ? "Run friction analysis to surface AI-refined moments."
        : "Run Summary to surface AI-refined friction moments.";
    }
    if (fd.llm_ok === false) return "Moment detection failed. Re-run with a downloaded AI model.";
    if (!(fd.moments && fd.moments.length)) return "No friction moments found in this transcript.";
    // Moments exist but none reached the strip: filter, or cited segments gone. Different fixes.
    if (state.frictionUnsourcedMoments) {
      return "Moments found, but the segments they cite are no longer in the transcript. Re-run friction.";
    }
    return "No moments match the current filter.";
  }

  function renderFrictionJumpStrip() {
    var strip = qs("#frictionJumpStrip");
    if (!strip) return;
    strip.innerHTML = "";
    var moments = state.frictionVisibleMoments || [];
    var prev = qs("#frictionJumpPrev");
    var next = qs("#frictionJumpNext");
    if (prev) prev.disabled = moments.length === 0;
    if (next) next.disabled = moments.length === 0;
    if (moments.length === 0) {
      state.frictionMomentIndex = -1;
      strip.appendChild(el(
        "span",
        "friction-jump-empty" + (state.frictionGenerating ? " cg-shimmer" : ""),
        _frictionJumpEmptyText()
      ));
      return;
    }
    // A filter change can shrink the list out from under the current selection.
    if (state.frictionMomentIndex >= moments.length) state.frictionMomentIndex = -1;
    var frag = document.createDocumentFragment();
    moments.forEach(function (entry, i) {
      var seg = state.segments[entry.idxs[0]];
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "friction-jump-chip" + (i === state.frictionMomentIndex ? " is-current" : "");
      if (entry.moment.rationale) {
        btn.setAttribute("data-tooltip", entry.moment.rationale);
      }
      btn.appendChild(el("span", "friction-jump-chip-num", String(i + 1)));
      btn.appendChild(el("span", "", _frictionCatLabel(entry.moment.category)));
      btn.appendChild(el("span", "friction-jump-chip-time", formatTime(seg.start)));
      btn.addEventListener("click", function () { _goToFrictionMoment(i); });
      frag.appendChild(btn);
    });
    strip.appendChild(frag);
  }

  function _goToFrictionMoment(i) {
    var moments = state.frictionVisibleMoments || [];
    if (moments.length === 0) return;
    if (i < 0) i = moments.length - 1;
    if (i >= moments.length) i = 0;
    state.frictionMomentIndex = i;
    renderFrictionJumpStrip();
    // Seek to the FIRST cited segment; the callout closes the passage below.
    _seekToSegmentIndex(moments[i].idxs[0]);
  }

  function _stepFrictionMoment(dir) {
    var moments = state.frictionVisibleMoments || [];
    if (moments.length === 0) return;
    var cur = state.frictionMomentIndex;
    _goToFrictionMoment(cur < 0 ? (dir > 0 ? 0 : moments.length - 1) : cur + dir);
  }

  // ---- Decorating the transcript itself ----

  function _buildFrictionCallout(entry, number) {
    var box = el("div", "friction-callout");
    var head = el("div", "friction-callout-head");
    head.appendChild(el("span", "friction-callout-num", String(number)));
    head.appendChild(el("span", "friction-cat-badge friction-cat-badge--sm",
      _frictionCatLabel(entry.moment.category)));
    head.appendChild(el("span", "friction-callout-score",
      (entry.moment.score != null ? entry.moment.score : 0).toFixed(2)));
    box.appendChild(head);
    if (entry.moment.rationale) {
      box.appendChild(el("div", "friction-callout-rationale", entry.moment.rationale));
    }
    return box;
  }

  // The ONLY place friction touches #segmentList; renderSegments emits no friction markup.
  function _decorateSegmentList() {
    var list = qs("#segmentList");
    if (!list) return;
    // renderPartialSegments appends by ordinal while streaming; a mid-list callout would break it.
    if (state.streamingParticipant) return;

    var stale = list.querySelectorAll(".friction-callout");
    for (var s = 0; s < stale.length; s++) stale[s].parentNode.removeChild(stale[s]);

    var on = state.frictionMode !== "off";
    var isolate = state.frictionMode === "isolate";
    var matches = state.frictionMatchBySegId || {};
    var cited = state.frictionCitedBySegId || {};
    var band = state.frictionBandBySegId || {};
    var rows = list.querySelectorAll(".segment-row");
    var matched = 0;
    for (var i = 0; i < rows.length; i++) {
      var seg = state.segments[i];
      var score = seg ? matches[seg.id] : undefined;
      var isCited = !!(seg && cited[seg.id]);
      // flagged reads the union map; score drives tint alpha, isCited the left rail.
      var flagged = !!seg && band[seg.id] !== undefined;
      if (flagged) matched++;
      var row = rows[i];
      if (on && score !== undefined) {
        row.classList.add("segment-friction");
        row.style.setProperty("--seg-friction-alpha", score);
      } else {
        row.classList.remove("segment-friction");
        row.style.removeProperty("--seg-friction-alpha");
      }
      row.classList.toggle("segment-cited", on && isCited);
      // Hidden, never removed: state.cachedSegmentRows is indexed positionally against state.segments.
      row.classList.toggle("segment-hidden", isolate && !flagged);
    }

    if (on) {
      var moments = state.frictionVisibleMoments || [];
      for (var k = 0; k < moments.length; k++) {
        // Anchor on the LAST cited row so the callout closes the quoted passage.
        var anchor = rows[moments[k].idxs[moments[k].idxs.length - 1]];
        if (!anchor) continue;
        list.insertBefore(_buildFrictionCallout(moments[k], k + 1), anchor.nextSibling);
      }
    }

    var counter = qs("#frictionCounter");
    if (counter) {
      var total = state.segments.length;
      var text = matched + " of " + clipgenPluralUnit(total, "segment", "segments");
      var stats = state.frictionData && state.frictionData.stats;
      if (stats && stats.markers_per_minute != null) {
        text += " · " + stats.markers_per_minute + " markers/min";
      }
      counter.textContent = text;
    }
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
    // Read the shared map rather than re-deriving "matching" here.
    _recomputeFrictionMatches();
    var matches = state.frictionMatchBySegId;
    var groups = {};
    var claimed = {};
    function claim(segId, cat) {
      if (claimed[segId]) return;
      claimed[segId] = true;
      if (!groups[cat]) groups[cat] = [];
      groups[cat].push(segId);
    }
    fd.segments.forEach(function (frow) {
      if (matches[frow.id] === undefined) return;
      var primary = _primaryCategory(frow);
      // Prefer the dominant category, but never one the user filtered out.
      if (!primary || state.frictionCategoryFilter[primary] === false) {
        primary = null;
        var cats = frow.categories || [];
        for (var i = 0; i < cats.length; i++) {
          if (state.frictionCategoryFilter[cats[i]] !== false) { primary = cats[i]; break; }
        }
      }
      if (!primary) return;
      claim(frow.id, primary);
    });
    // AI-only lines are absent from the match map; claim them under the moment's category.
    var visible = state.frictionVisibleMoments || [];
    visible.forEach(function (entry) {
      var cat = _frictionMomentCategory(entry.moment);
      entry.idxs.forEach(function (idx) {
        var seg = state.segments[idx];
        if (seg) claim(seg.id, cat);
      });
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

  function initFriction() {
    qs("#frictionRerun").addEventListener("click", function () { _startFrictionRun(); });
    qs("#frictionCancel").addEventListener("click", function () { _stopFrictionRun(); });
    var staleRerun = qs("#frictionStaleRerun");
    if (staleRerun) staleRerun.addEventListener("click", function () { _startFrictionRun(); });
    qs("#frictionMarkAll").addEventListener("click", function () { _frictionMarkAll(); });
    qs("#frictionJumpPrev").addEventListener("click", function () { _stepFrictionMoment(-1); });
    qs("#frictionJumpNext").addEventListener("click", function () { _stepFrictionMoment(1); });
  }

  // ---- Friction mode (Off / Highlight / Isolate) ----
  // Highlight tints; Isolate also hides non-matches.
  var FRICTION_MODES = [
    { value: "off", icon: "no-symbol", title: "Friction off" },
    { value: "highlight", icon: "fire", title: "Highlight matching segments in the transcript" },
    { value: "isolate", icon: "funnel", title: "Show only matching segments" },
  ];
  var FRICTION_MODE_ORDER = ["off", "highlight", "isolate"];

  function _frictionMode(v) {
    return v === "highlight" || v === "isolate" ? v : "off";
  }

  function initFrictionMode() {
    var mount = qs("#frictionModeMount");
    if (!mount) return;
    var stored = getStoredUIState("transcripts") || {};
    state.frictionMode = _frictionMode(stored.frictionMode);
    if (typeof stored.frictionMin === "number") state.frictionMin = stored.frictionMin;
    if (typeof stored.frictionMax === "number") state.frictionMax = stored.frictionMax;
    if (stored.frictionCategoryFilter) {
      state.frictionCategoryFilter = stored.frictionCategoryFilter;
    }
    if (stored.frictionMomentFilter) {
      state.frictionMomentFilter = stored.frictionMomentFilter;
    }
    mount.appendChild(createSegTrack({
      id: "frictionModeInput",
      value: state.frictionMode,
      options: FRICTION_MODES,
      size: "sm",
      onChange: _setFrictionMode,
    }));
    _initFrictionHistogramDrag();
  }

  function _setFrictionMode(mode) {
    state.frictionMode = _frictionMode(mode);
    setStoredUIStateField("transcripts", "frictionMode", state.frictionMode);
    applyFrictionDecorations();
    renderTimeline();
  }

  // Keyboard/palette entry point: the segmented control has no single toggle.
  function cycleFrictionMode() {
    var next = FRICTION_MODE_ORDER[
      (FRICTION_MODE_ORDER.indexOf(state.frictionMode) + 1) % FRICTION_MODE_ORDER.length
    ];
    var hidden = qs("#frictionModeInput");
    if (hidden && hidden.parentNode) segTrackSetValue(hidden.parentNode, next);
    _setFrictionMode(next);
  }

  // Friction tooltips share #trTooltip; state.frictionTooltipShown lets the video satellite's hideTimelineTooltip yield.
  function _visibleMomentsCiting(seg) {
    var out = [];
    if (!seg) return out;
    var visible = state.frictionVisibleMoments || [];
    for (var i = 0; i < visible.length; i++) {
      for (var j = 0; j < visible[i].idxs.length; j++) {
        var s = state.segments[visible[i].idxs[j]];
        if (s && s.id === seg.id) { out.push(visible[i]); break; }
      }
    }
    return out;
  }

  function _showFrictionTooltip(frow, seg, clientX, clientY) {
    var tip = qs("#trTooltip");
    if (!tip) return;
    tip.textContent = "";
    // Fold in moments: AI-cited lines score 0 with keywords and would read as a bug.
    var moments = _visibleMomentsCiting(seg);
    var cats = frow.categories || [];
    if (cats.length || moments.length) {
      var badges = document.createElement("div");
      badges.className = "tr-tooltip-friction-cats";
      cats.forEach(function (c) {
        badges.appendChild(el("span", "friction-cat-badge friction-cat-badge--sm", _frictionCatLabel(c)));
      });
      moments.forEach(function (entry) {
        badges.appendChild(el("span", "friction-cat-badge friction-cat-badge--sm friction-cat-badge--ai",
          _frictionCatLabel(entry.moment.category)));
      });
      tip.appendChild(badges);
    }
    // Quote the line: on the timeline band there is no text nearby.
    var quote = _frictionQuote(seg, 90);
    if (quote) {
      tip.appendChild(el("span", "tr-tooltip-friction-quote", quote));
      tip.appendChild(document.createElement("br"));
    }
    var markers = frow.markers || [];
    if (markers.length) {
      var shown = markers.slice(0, 5).join(", ");
      if (markers.length > 5) shown += " +" + (markers.length - 5) + " more";
      tip.appendChild(el("span", "tr-tooltip-friction-markers", "matched: " + shown));
      tip.appendChild(document.createElement("br"));
    }
    // Name each score's source; keyword density and model confidence are different scales.
    var scoreParts = [];
    if ((frow.score || 0) > 0) scoreParts.push("keyword " + frow.score.toFixed(2));
    for (var mi = 0; mi < moments.length; mi++) {
      scoreParts.push("AI " + (moments[mi].moment.score || 0).toFixed(2));
    }
    tip.appendChild(el("span", "tr-tooltip-friction-score",
      (seg ? formatTime(seg.start) + " · " : "") +
      (scoreParts.length ? scoreParts.join(" · ") : "score 0.00")));
    tip.classList.remove("hidden");
    var tipRect = tip.getBoundingClientRect();
    var x = clientX + 12;
    var y = clientY - tipRect.height - 12;
    if (x + tipRect.width > window.innerWidth - 8) x = window.innerWidth - tipRect.width - 8;
    if (y < 8) y = clientY + 16;
    if (y + tipRect.height > window.innerHeight - 8) y = Math.max(8, window.innerHeight - tipRect.height - 8);
    tip.style.left = x + "px";
    tip.style.top = y + "px";
    state.frictionTooltipShown = true;
  }

  function _hideFrictionTooltip() {
    state.frictionTooltipShown = false;
    var tip = qs("#trTooltip");
    // Yield to the timeline-canvas hover (video satellite owns _lastTimelineHit).
    if (tip && !(TS.hasTimelineHover && TS.hasTimelineHover())) tip.classList.add("hidden");
  }

  // ---- Published back to the hub ----
  // Reached by boot, selectParticipant, the poller, renderSegments, pills.
  TS.loadSummary = loadSummary;
  TS.loadFriction = loadFriction;
  TS.clearAnalysisPanel = clearAnalysisPanel;
  TS._setAnalysisPanelVisible = _setAnalysisPanelVisible;
  TS._restoreActiveTab = _restoreActiveTab;
  TS.initPanelTabs = initPanelTabs;
  TS.initSummaryActions = initSummaryActions;
  TS.initFriction = initFriction;
  TS.initFrictionMode = initFrictionMode;
  TS.applyFrictionDecorations = applyFrictionDecorations;
  TS.cycleFrictionMode = cycleFrictionMode;
  TS._stopSummaryPoll = _stopSummaryPoll;
  TS._stopCitationsPoll = _stopCitationsPoll;
  TS._stopFrictionPoll = _stopFrictionPoll;
  TS._currentParticipant = _currentParticipant;
  TS._frictionDepMet = _frictionDepMet;
  TS._showFrictionTooltip = _showFrictionTooltip;
  TS._hideFrictionTooltip = _hideFrictionTooltip;
  // True for either transport, so the hub's re-arm never restarts a live stream.
  TS.isSummaryPolling = function () {
    return !!(AGENT_DESCRIPTORS.summary._poller || _summaryStream);
  };
})();

