/* clipgen Transcripts analysis-panel satellite — transcripts-agents.js
 *
 * The tabbed analysis panel: the AI summary (+ inline edit + citations) and the
 * friction pass (scores, moments, filters, heatmap toggle, hot-segment tooltip),
 * plus the panel tab switching. These are the Ollama "thinking agent" results
 * surfaced per participant. Loaded LAST (after transcripts.js + the other
 * satellites); reads the hub's shared state + helpers through
 * window.ClipgenTranscripts (TS) and publishes its load/clear/stop/init entry
 * points back so selectParticipant, the task poller, the visibility/focus
 * handlers, boot, and the segment-list hover can reach them. The pills satellite
 * also reaches loadFriction via TS.loadFriction.
 *
 * The ETA trackers + ticker and _updateAgentElapsed live in the hub (shared with
 * the transcription ETA path) and are used here by reference. isSummaryPolling
 * lets the hub poller re-arm the summary poll without reading _summaryPoller
 * directly. Plain utils.js globals (qs/el/escapeHtml/apiGet/apiPost/formatTime/
 * createPoller/MARK_CATEGORIES/CLIPGEN_CONFIG/...) are reached via the scope chain.
 */
(function () {
  "use strict";

  var TS = window.ClipgenTranscripts;
  var state = TS.state;
  var showToast = TS.showToast,
    renderSegments = TS.renderSegments,
    renderTimeline = TS.renderTimeline,
    seekVideo = TS.seekVideo,
    scrollToSegment = TS.scrollToSegment,
    loadTranscript = TS.loadTranscript,
    ensureAgentModelInstalled = TS.ensureAgentModelInstalled,
    _refreshAgentStateNow = TS._refreshAgentStateNow,
    _txEtaTicker = TS._txEtaTicker,
    _summaryEtaTracker = TS._summaryEtaTracker,
    _citationsEtaTracker = TS._citationsEtaTracker,
    _frictionEtaTracker = TS._frictionEtaTracker,
    _updateAgentElapsed = TS._updateAgentElapsed,
    _currentParticipantHasTranscript = TS._currentParticipantHasTranscript;

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
    var ver = state.participantReqVer;

    // The analysis panel must be visible whenever we surface agent state for the
    // selected, transcribed participant. renderSummaryGenerating/renderSummary
    // (and the friction equivalents) don't toggle #summarySection themselves, so
    // a summary that registers *after* the transcript finalized would otherwise
    // paint its "Generating…" box into a hidden panel — visible only on reload.
    if (pid === state.selectedParticipant && _currentParticipantHasTranscript()) {
      _setAnalysisPanelVisible(true);
    }

    apiGet("api/summary/" + pid).then(function (data) {
      if (ver !== state.participantReqVer) return;
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
      if (ver !== state.participantReqVer) return;
      // Ollama unavailable or no summary — show the empty-state CTA.
      renderSummaryEmpty();
    });
  }

  function _startSummaryPoll(pid) {
    _stopSummaryPoll();
    var ver = state.participantReqVer;
    // runImmediately is false to match the previous setInterval (first poll after 3s).
    _summaryPoller = createPoller(function () {
      if (ver !== state.participantReqVer || state.selectedParticipant !== pid) {
        _stopSummaryPoll();
        return;
      }
      apiGet("api/summary/" + pid).then(function (data) {
        if (ver !== state.participantReqVer) return;
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
        if (ver !== state.participantReqVer) return;
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
    // Summary is (re)generating → its dependents' gate is momentarily unmet;
    // refresh the friction header so Re-run disables now and re-enables in
    // renderSummary() once the summary lands.
    _renderFrictionHeader();
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

  // Hard cap on agent polls (citations and friction). The previous 90 s value
  // was shorter than some real Ollama runs on long transcripts, so the result
  // would land in the manifest after we'd given up — and the UI only picked it
  // up on a full page reload. Five minutes covers realistic completion times;
  // the server-side `generating: false` signal stops the poll earlier when the
  // agent finishes (or fails) sooner.
  var _AGENT_POLL_TIMEOUT = 300000;

  function _startCitationsPoll(pid) {
    _stopCitationsPoll();
    var started = Date.now();
    var ver = state.participantReqVer;
    // runImmediately is false to match the previous setInterval (first poll after 3s).
    _citationsPoller = createPoller(function () {
      if (ver !== state.participantReqVer ||
          state.selectedParticipant !== pid ||
          Date.now() - started > _AGENT_POLL_TIMEOUT) {
        _stopCitationsPoll();
        state.citationsGenerating = false;
        var status = qs("#summaryContent .citations-status");
        if (status) status.remove();
        return;
      }
      apiGet("api/citations/" + pid).then(function (data) {
        if (ver !== state.participantReqVer) return;
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
        if (ver !== state.participantReqVer) return;
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
    var ver = state.participantReqVer;
    // Reveal the analysis panel for the selected, transcribed participant (see
    // loadSummary) so a friction run that registers after finalize is visible.
    if (pid === state.selectedParticipant && _currentParticipantHasTranscript()) {
      _setAnalysisPanelVisible(true);
    }
    state.frictionData = null;
    state.frictionBySegId = {};
    state.frictionGenerating = false;
    apiGet("api/friction/" + pid).then(function (data) {
      if (ver !== state.participantReqVer) return;
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
      if (ver !== state.participantReqVer) return;
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
    var ver = state.participantReqVer;
    // runImmediately is false to match the previous setInterval (first poll after 3s).
    _frictionPoller = createPoller(function () {
      if (ver !== state.participantReqVer ||
          state.selectedParticipant !== pid ||
          Date.now() - started > _AGENT_POLL_TIMEOUT) {
        _stopFrictionPoll();
        state.frictionGenerating = false;
        renderFriction();
        return;
      }
      apiGet("api/friction/" + pid).then(function (data) {
        if (ver !== state.participantReqVer) return;
        if (data.ok && data.friction) {
          _stopFrictionPoll();
          _setFrictionData(data.friction);
        } else if (!data.generating) {
          _stopFrictionPoll();
          state.frictionGenerating = false;
          renderFrictionEmpty();
        }
      }).catch(function () {
        if (ver !== state.participantReqVer) return;
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
  // state.frictionTooltipShown lets the video satellite's hideTimelineTooltip
  // yield while a friction tooltip owns #trTooltip (mirror of
  // _hideFrictionTooltip's TS.hasTimelineHover() guard). The segment-list
  // mousemove that calls _showFrictionTooltip — and its coalescing _segTooltipRaf
  // — live in the hub's segment-list delegation, not here.
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
  // Boot wires initPanelTabs/initSummaryActions/initFriction/initFrictionHeatmapToggle;
  // selectParticipant + the poller + the visibility/focus handlers drive
  // load*/clear*/stop*/_restoreActiveTab; the segment-list hover calls the
  // friction tooltips; the poller asks isSummaryPolling. loadFriction is also
  // reached by the pills satellite's friction run/stop rows.
  TS.loadSummary = loadSummary;
  TS.loadFriction = loadFriction;
  TS.clearAnalysisPanel = clearAnalysisPanel;
  TS._setAnalysisPanelVisible = _setAnalysisPanelVisible;
  TS._restoreActiveTab = _restoreActiveTab;
  TS.initPanelTabs = initPanelTabs;
  TS.initSummaryActions = initSummaryActions;
  TS.initFriction = initFriction;
  TS.initFrictionHeatmapToggle = initFrictionHeatmapToggle;
  TS._stopSummaryPoll = _stopSummaryPoll;
  TS._stopCitationsPoll = _stopCitationsPoll;
  TS._stopFrictionPoll = _stopFrictionPoll;
  TS._currentParticipant = _currentParticipant;
  TS._frictionDepMet = _frictionDepMet;
  TS._showFrictionTooltip = _showFrictionTooltip;
  TS._hideFrictionTooltip = _hideFrictionTooltip;
  TS.isSummaryPolling = function () { return !!_summaryPoller; };
})();

