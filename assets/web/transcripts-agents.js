/* clipgen Transcripts analysis-panel satellite — transcripts-agents.js
 *
 * The tabbed analysis panel: the AI summary (+ inline edit + citations) and the
 * friction pass (mode switch, score histogram, category chips, moment jump strip,
 * and the decorations they drive on the transcript below), plus the panel tab
 * switching. These are the Ollama "thinking agent" results
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
  //
  // summary / citations / friction all poll the same generic /api/agent/<key>/
  // endpoints; only the URL base, cadence, and result-handling hooks differ, so
  // one _makeAgentPoll scaffold drives all three (the "one JS poll factory").
  // Adding an agent is a descriptor entry + its render hooks — no new poll/stop
  // plumbing. Summary stays richest (SSE token stream + citation chaining +
  // inline edit); its extra behavior rides in its hooks and the bespoke run
  // helpers further down.

  // Hard cap on agent polls (citations + friction). Long Ollama runs on big
  // transcripts once outlived a shorter timeout, so the result landed in the
  // manifest after we'd given up and only surfaced on a full reload. Five
  // minutes covers realistic completion; the server's `generating: false`
  // stops the poll earlier when the agent finishes or fails sooner. Summary has
  // no cap — its SSE stream (or 1.2s fallback poll) runs until done.
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
        // Still generating — stream in whatever tokens have arrived. Rebuild the
        // generating box if a re-render dropped it, then push the partial text.
        if (!qs("#summaryStream")) {
          renderSummaryGenerating(d.started_at ? d.started_at * 1000 : undefined);
        }
        if (d.partial) _updateSummaryStream(d.partial);
      },
      onEmpty: function () { renderSummaryEmpty(); },
      // Participant switch: just stop — don't paint an empty box into the panel,
      // which now belongs to a different participant.
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
        // Cancel clears the flag (via _clearCitationsStatus) but cannot recall a
        // GET already on the wire, so that response still lands here — stay
        // silent, a deliberate abort is not a failure.
        if (!state.citationsGenerating) return;
        // Otherwise the run is genuinely over with nothing to show: the route
        // 404s (which apiGet rejects) once it ends without persisting — after
        // find_citations' failure return, the Ollama-unavailable path. Say so
        // instead of just dropping "Finding sources…". The cause is a
        // suggestion, not a claim: a transport blip lands here too, and the fix
        // is the same. Participant switch and poll timeout go to onStale.
        state.citationsGenerating = false;
        _renderCitationsNote("Couldn't find sources. Check that Ollama is running, then re-run citations.");
      },
      onStale: function () { _clearCitationsStatus(); },
    },
    friction: {
      key: "friction",
      urlBase: "api/agent/friction",
      interval: 3000,
      timeout: _AGENT_POLL_TIMEOUT,
      _poller: null,
      // The deterministic placeholder (no persisted LLM run) is a display fallback,
      // not a completed agent result — exclude it here so a run that ends without
      // persisting (cancel/exception/None) falls through to onEmpty instead of being
      // treated as "done". onEmpty still surfaces those programmatic scores.
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

  // The one poll scaffold. The version/participant staleness guard + optional
  // timeout are identical across agents; per-result behavior is the descriptor's
  // hooks. runImmediately:false so the first poll waits one interval (the
  // initial render already painted the box). createPoller auto-pauses on hidden
  // tabs and the endpoints are cheap in-memory reads.
  function _makeAgentPoll(desc) {
    return function (pid) {
      _stopAgentPoll(desc);
      var started = Date.now();
      var ver = state.participantReqVer;
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
          if (data.ok && desc.getResult(data)) {
            _stopAgentPoll(desc);
            desc.onResult(pid, data);
          } else if (data.generating) {
            if (desc.onGenerating) desc.onGenerating(pid, data);
          } else {
            _stopAgentPoll(desc);
            desc.onEmpty(pid, data);
          }
        }).catch(function () {
          if (ver !== state.participantReqVer) return;
          _stopAgentPoll(desc);
          desc.onEmpty(pid);
        });
      }, desc.interval, { runImmediately: false });
      desc._poller.start();
    };
  }

  // Timer-only teardown (mirrors the old per-agent _stop*Poll): does NOT reset
  // ETA trackers — render*Status / render*Generating reset-then-seed, so resetting here
  // would wipe the seed when a poll restarts.
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
  //
  // Two cooperating pollers, both built by _makeAgentPoll above:
  //   AGENT_DESCRIPTORS.summary._poller   — runs while the backend is still
  //                      generating the summary. Stops as soon as a summary
  //                      lands, or when the user switches participant.
  //   AGENT_DESCRIPTORS.citations._poller — runs after the summary arrives if
  //                      citations are still computing (citations depend on
  //                      summary).
  // Both are stopped by their own _stop*Poll() helpers; either also stops if
  // `state.selectedParticipant` no longer matches the participant the poll
  // was started for.

  // SSE token stream: the primary live-update transport while a summary
  // generates — pushes each token as the model emits it (true word-by-word).
  // The summary poll (AGENT_DESCRIPTORS.summary._poller) is the fallback used
  // when EventSource is unsupported or the stream drops mid-run.
  var _summaryStream = null;

  function loadSummary(pid) {
    var ver = state.participantReqVer;

    // The analysis panel must be visible whenever we surface agent state for the
    // selected, transcribed participant. renderSummaryGenerating/renderSummary
    // (and the friction equivalents) don't toggle #summarySection themselves, so
    // a summary that registers *after* the transcript finalized would otherwise
    // paint its "Generating…" box into a hidden panel — visible only on reload.
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
    }).catch(function () {
      if (ver !== state.participantReqVer) return;
      // Ollama unavailable or no summary — show the empty-state CTA.
      renderSummaryEmpty();
    });
  }

  // Summary landed (via initial load OR the fallback poll's onResult hook):
  // render it, then attach citations — still generating (surface the status and
  // poll), or already stored (fetch them). Shared so the loader and the poll
  // stay in lockstep (the summary poll's descriptor onResult delegates here).
  function _onSummaryResult(pid, data) {
    // Clear any citation state carried over from a previous participant before
    // rendering — renderSummary() reapplies state.summaryCitations, so stale
    // superscripts would otherwise leak onto this summary.
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

  // The summary response carries citations' *status* but not their payload (the
  // generic agent GET returns only its own manifest field, and inlining the
  // dependents would put the whole friction blob on the 1.2s summary poll). So
  // a settled load has to ask for them separately, or every loadSummary —
  // participant switch, tab refocus, post-cancel re-sync — would render the
  // summary with its superscripts permanently stripped.
  function _loadStoredCitations(pid) {
    var ver = state.participantReqVer;
    apiGet(AGENT_DESCRIPTORS.citations.urlBase + "/" + pid).then(function (data) {
      if (ver !== state.participantReqVer || state.selectedParticipant !== pid) return;
      // A run may have started while this GET was in flight (Regenerate, or the
      // chain reaching citations). Restoring the old result now would replace
      // the live "Finding sources…" line — renderCitations() removes it — with
      // superscripts the run is about to supersede.
      if (state.citationsGenerating) return;
      if (!data.ok || !data.citations) return;
      state.summaryCitations = data.citations;
      renderCitations();
    }).catch(function () {
      // 404 = citations never ran (or are disabled). Not a failed run, so stay
      // quiet — the poll's onEmpty owns the "a run just ended empty" message.
    });
  }

  // Open the SSE token stream for a generating summary. Falls back to the GET
  // poll if EventSource is unavailable or the stream drops. onMessage carries
  // either {partial} (text so far) or {done} (run finished → render the
  // finalized summary via loadSummary, which also kicks the citations chain).
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

  // Stops all live summary updates — the poller AND the SSE stream — so every
  // existing teardown site (participant switch, clear, cancel, completion)
  // covers both transports without needing to know which one is active.
  function _stopSummaryPoll() {
    _stopAgentPoll(AGENT_DESCRIPTORS.summary);
    _stopSummaryStream();
  }

  // startedAtMs (optional): server-recorded run start in epoch ms. Seeds the
  // elapsed clock so navigating away and back resumes from the true elapsed
  // time instead of zero; omit it for a just-clicked manual run (starts now).
  function renderSummaryGenerating(startedAtMs) {
    var content = qs("#summaryContent");
    // #summaryStream is a separate node so streamed partial text can be updated
    // by later polls (via _updateSummaryStream) without disturbing the elapsed
    // clock / Cancel wiring below, which is built once here.
    content.innerHTML =
      '<div class="summary-stream" id="summaryStream"></div>' +
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
    // Summary is (re)generating → its dependents' gate is momentarily unmet;
    // refresh the friction header so Re-run disables now and re-enables in
    // renderSummary() once the summary lands.
    _renderFrictionHeader();
  }

  // Push streamed partial summary text into the #summaryStream node built by
  // renderSummaryGenerating, without touching the elapsed clock / Cancel footer.
  // Plain textContent (CSS white-space: pre-wrap handles newlines); citation
  // anchors are added only when the finished summary lands via renderSummary.
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

  // "No summary yet." plus a Run button that silently does nothing is the most
  // common first contact with the AI features, and it never said why. Fill in
  // the reason when there is one. Fire-and-forget: the empty state is already
  // on screen and correct without it, so a slow or failed /api/models must
  // never delay the render or throw.
  function _updateSummaryEmptyHint() {
    var hintEl = qs("#summaryEmptyHint");
    if (!hintEl) return;
    hintEl.classList.add("hidden");
    hintEl.textContent = "";
    _trFetchModels().then(function (data) {
      var status = clipgenOllamaStatus(data && data.ollama);
      if (status.state === "ok") return;
      var extra = "";
      if (status.state === "missing") {
        // Running the summary raises the install dialog, so the shortest true
        // instruction here is "just run it".
        extra = status.canInstall
          ? " clipgen can download it for you when you run the summary."
          : (status.hint.length ? " " + status.hint[0] : "");
      }
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
    var paragraphSentences = [];
    var bullets = [];
    var inBullets = false;

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;
      if (line.indexOf("- ") === 0 || line.indexOf("* ") === 0) {
        inBullets = true;
        bullets.push(clipgenRenderInlineMarkdown(line.substring(2)));
      } else if (!inBullets) {
        // Split paragraph into individual sentences for citation targeting
        var parts = line.split(/(?<=[.!?])\s+/);
        for (var k = 0; k < parts.length; k++) {
          var part = parts[k].trim();
          // Models emit **bold** / `code` emphasis; render it instead of
          // showing literal asterisks (same helper as the Overview report).
          if (part) paragraphSentences.push(clipgenRenderInlineMarkdown(part));
        }
      } else {
        bullets.push(clipgenRenderInlineMarkdown(line));
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
    // The friction Re-run button is gated on the summary dependency. The summary
    // just landed (dep now met), so refresh the friction header even when friction
    // data already exists — loadFriction is skipped in that case, so otherwise the
    // Re-run button stays disabled/unclickable until a page reload.
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

  // Citations cleanup shared by the poll's onEmpty/onStale hooks, the cancel
  // path, and the regenerate error path: drop the generating flag and remove
  // the "Finding sources…" status line from the summary panel.
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

    // Rendering nothing is itself a result: say so, rather than leaving the
    // panel looking like citations never ran. dataRefs separates "the model
    // found no supporting segments" from "it found some, but renderSummary's
    // sentence split didn't line up with the backend's" \u2014 the two splits are
    // independent implementations, and that mismatch is otherwise silent.
    if (refNum === 1) {
      _renderCitationsNote(dataRefs === 0
        ? "No supporting segments found for this summary."
        : "Sources were found but couldn't be matched to the summary text. Re-run citations.");
    }
  }

  // Terminal one-line status in the summary panel (nothing found / run failed).
  // Shares .citations-status with the in-flight "Finding sources\u2026" line, minus
  // its elapsed clock and Cancel button.
  function _renderCitationsNote(text) {
    var existing = qs("#summaryContent .citations-status");
    if (existing) existing.remove();
    var p = document.createElement("p");
    p.className = "citations-status";
    p.textContent = text;
    qs("#summaryContent").appendChild(p);
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

  function _stopCitationsPoll() {
    _stopAgentPoll(AGENT_DESCRIPTORS.citations);
  }

  function _stopCitationsRun() {
    var pid = state.selectedParticipant;
    if (!pid) return;
    // Citations run after the summary exists, so keep the summary visible and
    // only remove the "Finding sources…" status line.
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
    // Regenerate runs summary → citations as a chain, so summary may have
    // already finished and citations started by the time Cancel is clicked
    // (3s poll gap). Stop both — each call is a no-op if that pass isn't
    // running. Re-sync only after both stops are acknowledged, otherwise the
    // follow-up GET can still see citations in-flight and restart its poll.
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

  // ---- Friction detection (Pass 3) ----
  //
  // Programmatic scores + LLM moments land together in the manifest's
  // `friction` field. The tab is a CONTROL SURFACE over #segmentList rather than
  // a results list: a mode switch, a score histogram and category chips filter
  // the transcript below (tinting it, or isolating it down to the matches), and
  // the moments are a jump strip whose rationales render as inline callouts under
  // the segments they quote. Everything downstream reads one derived map — see
  // _recomputeFrictionMatches. Generation mirrors summary/citations (poll until
  // done; manual run/cancel bypass the global flag).

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
    // Trust the client-side summary too: renderSummary sets state.summaryText for
    // the selected participant the moment a summary lands, before /api/participants
    // catches up to summary === "done". Without this the Re-run gate lags a poll
    // (or never updates until reload) after a summary regenerate completes.
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
    // A category the model invented. Show its own wording rather than dropping
    // it — the evidence table groups these under Other, but the jump chip and
    // the callout should still say what the model actually called it.
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
    // Reveal the analysis panel for the selected, transcribed participant (see
    // loadSummary) so a friction run that registers after finalize is visible.
    if (pid === state.selectedParticipant && _currentParticipantHasTranscript()) {
      _setAnalysisPanelVisible(true);
    }
    // Blank only when the participant actually changed. A same-participant
    // refetch — tab refocus, the mid-run re-arm in transcripts.js, a re-select —
    // must keep the programmatic scores on screen: they come from the
    // deterministic scorer and owe nothing to the LLM. Wiping them here is what
    // blanked the histogram, chips, tinting and timeline band for a whole run.
    if (state.frictionPid !== pid) {
      state.frictionData = null;
      state.frictionBySegId = {};
      state.frictionMomentIndex = -1;
    }
    state.frictionPid = pid;
    state.frictionGenerating = false;
    apiGet(AGENT_DESCRIPTORS.friction.urlBase + "/" + pid).then(function (data) {
      if (ver !== state.participantReqVer) return;
      if (data.ok && data.friction) {
        _setFrictionData(data.friction);
      } else if (data.generating) {
        // Regenerate pops the stored result, so mid-run the server answers with
        // the deterministic scores alongside `generating` — adopt them before
        // flipping the flag, or this branch shows the empty "Analyzing…" box.
        if (data.friction) _setFrictionData(data.friction);
        state.frictionGenerating = true;
        state.frictionStartedAt = data.started_at ? data.started_at * 1000 : null;
        renderFrictionGenerating();
        _startFrictionPoll(pid);
        _refreshAgentStateNow();
      } else {
        renderFrictionEmpty();
      }
    }).catch(function () {
      if (ver !== state.participantReqVer) return;
      renderFrictionEmpty();
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
    // renderFriction's applyFrictionDecorations already re-decorated the existing
    // rows in place, so no segment-list rebuild is needed — only the canvas band.
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
    // The deterministic-only placeholder has no AI result yet, so it reads "Run",
    // not "Re-run". Write the label span, not the button: the button also holds
    // the .ai-agent-badge, which a textContent write would delete.
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
        // depMet distinguishes "no summary yet" (run Summary, which auto-chains
        // friction) from "summary done, friction not run" (the friction button is
        // enabled) so the copy points at the step the user can actually take.
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
        // A completed run that surfaced nothing. Without this the header reads
        // like any successful run and the only hint is the jump strip's one
        // italic line, well below the fold of the eye.
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
        // Show the banner only for edit-staleness, not for an LLM failure (which
        // the header already explains with its own re-run guidance).
        banner.classList.toggle("hidden", !(fd.stale && fd.llm_ok !== false));
      }
      // Bins change only when the data does; the marker, chips, strip and the
      // transcript decorations all follow from applyFrictionDecorations.
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
    // Keep the programmatic scores on screen while the agent runs. They come
    // from the deterministic scorer, are already computed, and are independent
    // of the LLM — blanking the pane costs the user the histogram, chips and
    // transcript tinting for the whole run, and the agent only ever *adds*
    // moments on top. The standalone box is for when there is nothing yet.
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

  // Two filters, one per evidence source: the keyword scorer's segment
  // categories and the agent's moment categories are labelled independently, so
  // one shared dict meant hiding a category on one side silently hid it on the
  // other. The moment side also carries the "other" bucket.
  function _ensureFrictionFilter() {
    var keys = _frictionCatKeys();
    var prog = state.frictionCategoryFilter || {};
    var ai = state.frictionMomentFilter || {};
    // Fill in rather than replace: a persisted filter from an older category set
    // keeps whatever the user chose and picks up new categories enabled.
    for (var i = 0; i < keys.length; i++) {
      if (prog[keys[i]] === undefined) prog[keys[i]] = true;
      if (ai[keys[i]] === undefined) ai[keys[i]] = true;
    }
    if (ai[FRICTION_OTHER] === undefined) ai[FRICTION_OTHER] = true;
    state.frictionCategoryFilter = prog;
    state.frictionMomentFilter = ai;
  }

  // Both ends of the score band are user-controlled (the histogram's two
  // handles), so every score test goes through here rather than comparing
  // against a lone threshold.
  function _frictionScoreInBand(score) {
    return score >= state.frictionMin && score <= state.frictionMax;
  }

  // ---- The one derived filter product ----
  //
  // (threshold, category filter, friction data, segments) -> the three shared
  // fields every consumer reads: state.frictionMatchBySegId (segment id -> score),
  // state.frictionVisibleMoments and state.frictionCitedBySegId. The segment
  // tints, isolate hiding, the timeline density band, "Mark all matching", the
  // counter and the isolate-aware keyboard nav all read these, so the pane and
  // the transcript below it can never disagree about what counts as friction.
  //
  // A segment matches when its programmatic score clears the threshold and at
  // least one of its categories is still enabled, OR when a visible moment cites
  // it. That second clause is not redundant: the segment score comes from the
  // regex scorer (friction.py) while the moment score is the model's, so a moment
  // can clear the threshold while the line it quotes scores 0. Without it,
  // isolate mode would hide the very row the jump strip seeks to.
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
      // score <= 0 is excluded outright so a lower bound of 0 still means "every
      // segment the scorer flagged", not "every segment in the transcript".
      if (score <= 0 || !_frictionScoreInBand(score)) continue;
      var cats = frow.categories || [];
      for (j = 0; j < cats.length; j++) {
        if (filter[cats[j]] !== false) { matches[frow.id] = score; break; }
      }
    }
    state.frictionMatchBySegId = matches;

    // Resolve moments to segment indices exactly once, so the jump strip's
    // numbering and the set of inline callouts can't drift apart.
    var visible = [];
    var cited = {};
    var unsourced = 0;
    var all = (fd && fd.moments) || [];
    for (i = 0; i < all.length; i++) {
      if (!_frictionMomentMatches(all[i])) continue;
      var idxs = _momentSegmentIndices(all[i]);
      // Unsourced moment: nothing to quote or seek to. Counted, not just
      // dropped, so the empty strip can name this cause instead of the filter.
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

    // What the timeline density band draws: the union of both sources, keyed to
    // the strongest evidence on each line. Built here rather than at the band so
    // the canvas can never disagree with the pane about what is flagged. An
    // AI-only line scores 0 with the keyword scorer, so reading the keyword map
    // alone left the moments the jump strip is built around off the band
    // entirely; the two scales differ, but the band's alpha only ever meant
    // "how strong is the evidence here".
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

  // The single entry point for "the filter changed": recompute, then reflect it
  // everywhere. Called by the hub from renderSegments (before the scroll restore)
  // and by every control in the pane. Cheap enough to run per animation frame
  // during a threshold drag — it only writes classes and inline custom
  // properties, and never rebuilds the segment list.
  function applyFrictionDecorations() {
    _recomputeFrictionMatches();
    renderFrictionEvidence();
    _updateFrictionBounds(null);
    renderFrictionJumpStrip();
    _decorateSegmentList();
  }

  // ---- "Why was this selected" ----
  //
  // The scores are opaque on their own, so every friction hover surface — a
  // histogram bin, a hot segment row, the timeline density band — answers the
  // same question with the same words: what the segment scored, which categories
  // fired, the phrases that matched, and the line itself. One builder, so the
  // three can't end up explaining the same segment differently.

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
    // Resolve text only for the handful shown — a busy bin can hold hundreds of
    // rows and _segmentIndexById is a linear scan.
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
      // Only scored segments are binned. score_segments emits a row for every
      // segment and most score exactly 0, so including them makes bin 0 a spike
      // that flattens every other bin against the 6% floor.
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
        // Built lazily on hover, not captured here: the histogram can be rendered
        // before the transcript lands, and the explanation needs segment text.
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

  // Handle positions, per-bar dimming and the bound readouts. Split from
  // renderFrictionHistogram so a drag never wipes the track's innerHTML
  // mid-gesture (which would drop the pointer capture the drag depends on).
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

    // Labels sit under their own handle. When the two bounds nearly coincide the
    // labels would overlap, so nudge them apart — the positions stay indicative,
    // and the numbers stay readable, which is the point of the readout.
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
      // Bounds clamp against each other rather than swapping: dragging one end
      // past the other collapses the band instead of silently inverting it.
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

    // Pointer capture keeps the gesture on the track while the pointer wanders
    // off it, so nothing is bound at document scope and there is nothing to tear
    // down on pagehide. The handlers gate on `grabbed`, not on hasPointerCapture:
    // capture is a nicety some engines refuse (and WKWebView hosts this app in
    // the desktop bundle), and a drag that silently stops tracking is worse than
    // one that merely stops following the pointer past the edges. A move with no
    // button held means we missed the pointerup, so the gesture self-heals.
    host.addEventListener("pointerdown", function (e) {
      var v = scoreAt(e.clientX);
      if (v === null) return;
      // A press outside the band always grabs the bound on that side, so it
      // widens toward the press. Nearest-handle only applies inside. Without the
      // outside case, a band dragged shut (min === max) would be unrecoverable:
      // every press would tie, ties would pick min, and min can never exceed max.
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
      // Persist once per gesture, not per frame — the stored UI state is a
      // parse/stringify of the whole page blob.
      setStoredUIStateField("transcripts", "frictionMin", state.frictionMin);
      setStoredUIStateField("transcripts", "frictionMax", state.frictionMax);
    }
    host.addEventListener("pointerup", endDrag);
    host.addEventListener("pointercancel", endDrag);
  }

  // ---- Evidence table ----
  //
  // Two independent systems produce findings, and the old single chip row only
  // ever counted one of them. The keyword scorer (friction.py) labels SEGMENTS
  // from regex hits; the agent labels MOMENTS with a category string of its own
  // choosing, which is never reconciled against what the scorer found on the
  // line it quotes. So a category could read "Confusion 0" while a Confusion
  // moment sat in the jump strip — and because a 0-count chip was inert, that
  // was the one category you could not filter by.
  //
  // The table therefore counts and filters the two apart: a row per category,
  // a cell per source. Counts are computed here rather than read from
  // stats.by_category, which counts marker hits (not segments) and ignores the
  // score band, so it would never move with the histogram.
  var FRICTION_SOURCES = ["prog", "ai"];
  // Where a model category that is not one of the six lands. The jump chip and
  // the callout still show the model's own wording; only the row groups them.
  var FRICTION_OTHER = "other";

  function _frictionCatKeys() {
    var cats = CLIPGEN_CONFIG.frictionCategories || [];
    return cats.map(function (c) { return c.key; });
  }

  // The model is free to emit any string (thinking_agents.py only lowercases
  // and underscores it), so bucket anything unrecognized rather than letting it
  // fall through every filter and every row.
  function _frictionMomentCategory(m) {
    var raw = (m && m.category) || "";
    return _frictionCatKeys().indexOf(raw) === -1 ? FRICTION_OTHER : raw;
  }

  // Counts inside the current score band — what each cell displays.
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

  // Which rows exist at all, ignoring the band and the filters. Deliberately
  // band-independent: rows appearing and vanishing mid-drag would make the
  // control block jump under the pointer, and a cell whose count the band has
  // driven to 0 must stay clickable — gating inertness on the banded count is
  // exactly the bug this table replaces.
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

  // One row per source, one column per category: the two sources are the thing
  // being compared, so they read better as adjacent rows than as two columns
  // scanned down. Categories head the columns.
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
      // The header truncates on a narrow pane, so the full name has to survive
      // somewhere — and "Self-correction" is the first to go.
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
    // Rebuild only when the category set itself changed; a threshold drag then
    // just rewrites counts and state classes per animation frame.
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
      // Inert only when this source has no finding of this kind at all — not
      // merely when the band has hidden them, or the band would lock the user
      // out of the control that widens it.
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
  //
  // Moments are navigation, not content: one chip per moment, one line, always.
  // The evidence they used to carry as a blockquote is the transcript itself, and
  // their rationale renders as a .friction-callout under the quoted passage.

  function _frictionJumpEmptyText() {
    var fd = state.frictionData;
    // The pane now stays up during a run (renderFrictionGenerating), so the
    // strip has to speak for the in-flight case too — otherwise it advertises
    // the Run button that is currently mid-run.
    if (state.frictionGenerating) return "Analyzing friction…";
    if (!fd) return "No moments detected.";
    if (fd.deterministic) {
      return _frictionDepMet()
        ? "Run friction analysis to surface AI-refined moments."
        : "Run Summary to surface AI-refined friction moments.";
    }
    if (fd.llm_ok === false) return "Moment detection failed. Re-run with an installed Ollama model.";
    if (!(fd.moments && fd.moments.length)) return "No friction moments found in this transcript.";
    // Moments exist but none reached the strip. Either the filter excluded them
    // all, or they cite segments that are gone — different causes, different
    // fixes, so don't blame the filter for both.
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
      strip.appendChild(el("span", "friction-jump-empty", _frictionJumpEmptyText()));
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
    // Seek to the FIRST cited segment so the reader lands at the start of the
    // quoted passage and reads down into the callout that closes it.
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

  // Tints, isolate hiding, inline callouts and the counter. The ONLY place
  // friction touches #segmentList — renderSegments deliberately emits no friction
  // markup, so the full-rebuild path and this update path cannot diverge.
  function _decorateSegmentList() {
    var list = qs("#segmentList");
    if (!list) return;
    // renderPartialSegments appends by ordinal while a transcript streams in; a
    // callout inserted mid-list would corrupt that fast path.
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
      // "Flagged by either source" is the derived union map, not a third
      // hand-rolled OR. score and isCited stay separate below because they drive
      // different things: the tint alpha is the keyword score, the left rail is
      // the citation.
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
      // Hidden, never removed: state.cachedSegmentRows is indexed positionally
      // against state.segments, so removing a row would misalign every consumer.
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
    // Same predicate as the tints, the band and the counter — read the shared map
    // rather than re-deriving "matching" a second time here.
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
      // The row matched on some enabled category; prefer the dominant one, but
      // never label a group with a category the user filtered out.
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
    // Segments only the agent flagged. They score 0 with the keyword scorer, so
    // frictionMatchBySegId never holds them — and this action used to skip the
    // very lines the jump strip exists to surface. Marked under the moment's
    // own category, and never twice (the keyword pass claims a line first).
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
  //
  // Replaces the old player-bar heatmap button. "Highlight" is what that toggle
  // did (tint hot segments + draw the timeline density band); "Isolate" adds
  // hiding every row that doesn't match the filter, turning the transcript into
  // the result list. The control lives with the filters it belongs to.
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

  // Friction tooltip on hot segments (reuses the shared #trTooltip element).
  // state.frictionTooltipShown lets the video satellite's hideTimelineTooltip
  // yield while a friction tooltip owns #trTooltip (mirror of
  // _hideFrictionTooltip's TS.hasTimelineHover() guard). The segment-list
  // mousemove that calls _showFrictionTooltip — and its coalescing _segTooltipRaf
  // — live in the hub's segment-list delegation, not here.
  // Visible moments quoting this segment. Reads the already-resolved indices so
  // the hover can never name a moment the jump strip has filtered out.
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
    // The band now draws AI-cited lines too, and those score 0 with the keyword
    // scorer — without the moments folded in, hovering one of those stripes read
    // "score 0.00" with no categories, i.e. as if the stripe were a bug.
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
    // The line itself, so the hover answers "why is this flagged" without
    // needing to look back at the transcript (it is the whole point on the
    // timeline band, where there is no text nearby at all).
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
    // Name the source of each score rather than printing a bare number: they
    // are on different scales (marker density vs the model's own confidence).
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
  // Boot wires initPanelTabs/initSummaryActions/initFriction/initFrictionMode;
  // selectParticipant + the poller + the visibility/focus handlers drive
  // load*/clear*/stop*/_restoreActiveTab; renderSegments calls
  // applyFrictionDecorations; the command palette calls cycleFrictionMode; the
  // segment-list hover calls the friction tooltips; the poller asks
  // isSummaryPolling. loadFriction is also reached by the pills satellite's
  // friction run/stop rows.
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
  // True while a summary is being live-tracked by EITHER transport, so the hub's
  // grace-window re-arm doesn't restart the stream out from under itself.
  TS.isSummaryPolling = function () {
    return !!(AGENT_DESCRIPTORS.summary._poller || _summaryStream);
  };
})();

