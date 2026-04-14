(function () {
  "use strict";

  var SEARCH_DEBOUNCE = 300;

  var fmtTime = formatTime;
  var SS_DETECTOR_COLORS = DETECTOR_COLORS;

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
    sheetDefaultDuration: 60,
    sheetLoaded: false,
    xrefPollTimer: null,
    tooltipsEnabled: true,
    summaryCollapsed: false,
    summaryEditing: false,
    summaryText: "",
    summaryCitations: null,
    citationsGenerating: false,
  };

  // ---- Helpers ----

  function escapeHtml(str) {
    var div = document.createElement("div");
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
  }

  function showToast(msg) {
    var el = qs("#toast");
    if (!el) return;
    el.textContent = msg;
    el.classList.remove("hidden");
    clearTimeout(el._timer);
    el._timer = setTimeout(function () { el.classList.add("hidden"); }, 2500);
  }

  // ---- Nav links ----

  function checkNavLinks() {
    fetch("../api/status").then(function (r) { return r.json(); }).then(function (data) {
      if (data.studio) qs("#studioLink").classList.remove("hidden");
      if (data.insights) qs("#insightsLink").classList.remove("hidden");
      if (data.screenspace) qs("#screenspaceLink").classList.remove("hidden");
      if (data.screenspace || data.studio) {
        loadCrossRefData();
        state.xrefPollTimer = setInterval(loadCrossRefData, 30000);
      }
    }).catch(function () {});
  }

  // ---- Cross-reference data ----

  function loadCrossRefData() {
    fetch("../screenspace/api/events?excluded=false")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data.ok) {
          state.ssEvents = data.events || [];
          state.ssEventsLoaded = true;
          if (state.searchResults) renderSearchResults(state.searchResults);
        }
      })
      .catch(function () {});

    fetch("../studio/api/sheet")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data.ok) {
          state.sheetRows = data.rows || [];
          state.sheetParticipants = data.participants || [];
          state.sheetDefaultDuration = data.defaultDuration || 60;
          state.sheetLoaded = true;
          if (state.searchResults) renderSearchResults(state.searchResults);
        }
      })
      .catch(function () {});
  }

  function parseTS(str) {
    var parts = str.split(":");
    if (parts.length === 3) return (+parts[0]) * 3600 + (+parts[1]) * 60 + (+parts[2]);
    if (parts.length === 2) return (+parts[0]) * 60 + (+parts[1]);
    return NaN;
  }

  function parseSheetTimestamps(raw) {
    var DEFAULT_DUR = state.sheetDefaultDuration || 60;
    var cleaned = raw.toLowerCase().replace(/!key/g, "").replace(/[+;,]/g, " ");
    var tokens = cleaned.split(/\s+/).filter(function (t) { return t && t !== "x"; });
    var segments = [];
    for (var i = 0; i < tokens.length; i++) {
      var tok = tokens[i].replace(/\.$/, "").replace(/\./g, ":");
      var dashIdx = -1;
      for (var d = 1; d < tok.length; d++) {
        if (tok[d] === "-" && tok[d - 1] >= "0" && tok[d - 1] <= "9") { dashIdx = d; break; }
      }
      if (dashIdx > 0) {
        var s = parseTS(tok.substring(0, dashIdx));
        var e = parseTS(tok.substring(dashIdx + 1));
        if (!isNaN(s) && !isNaN(e)) segments.push({ start: Math.floor(s), duration: Math.max(0, e - s) });
      } else if (tok.indexOf(":") > 0) {
        var sec = parseTS(tok);
        if (!isNaN(sec)) segments.push({ start: Math.floor(sec), duration: DEFAULT_DUR });
      }
    }
    return segments;
  }

  function findOverlapsForSearch(participant, start, end) {
    var result = { screenspaceEvents: [], sheetObservations: [] };

    for (var i = 0; i < state.ssEvents.length; i++) {
      var ev = state.ssEvents[i];
      if (ev.participant === participant && ev.time_in < end && ev.time_out > start) {
        result.screenspaceEvents.push(ev);
      }
    }

    for (var j = 0; j < state.sheetRows.length; j++) {
      var row = state.sheetRows[j];
      var cell = row.cells[participant];
      if (!cell || !cell.valid) continue;
      var segs = parseSheetTimestamps(cell.value);
      for (var k = 0; k < segs.length; k++) {
        var segEnd = segs[k].start + segs[k].duration;
        if (segs[k].start < end && segEnd > start) {
          result.sheetObservations.push({ observation: row.observation, category: row.category });
          break;
        }
      }
    }

    return result;
  }

  // ---- Participants ----

  function loadParticipants() {
    return apiGet("api/participants").then(function (data) {
      if (!data.ok) return;
      state.participants = data.participants;
      renderParticipantSelect();

      // Preserve current selection if still valid
      if (state.selectedParticipant) {
        for (var i = 0; i < state.participants.length; i++) {
          if (state.participants[i].id === state.selectedParticipant) return;
        }
      }

      // Auto-select first participant with a transcript, or just the first
      var first = null;
      for (var i = 0; i < state.participants.length; i++) {
        if (state.participants[i].has_transcript) { first = state.participants[i]; break; }
      }
      if (!first && state.participants.length > 0) first = state.participants[0];
      if (first) selectParticipant(first.id);
      else renderEmptyState();
    });
  }

  function renderParticipantSelect() {
    var sel = qs("#participantSelect");
    sel.innerHTML = "";
    if (state.participants.length === 0) {
      sel.innerHTML = '<option value="">No participants</option>';
      return;
    }
    state.participants.forEach(function (p) {
      var opt = document.createElement("option");
      opt.value = p.id;
      opt.textContent = p.id + (p.has_transcript ? " \u2713" : "") + (p.has_stale_artifacts ? " \u26a0" : "");
      sel.appendChild(opt);
    });
    if (state.selectedParticipant) {
      sel.value = state.selectedParticipant;
    }
  }

  function selectParticipant(pid) {
    state.selectedParticipant = pid;
    var sel = qs("#participantSelect");
    sel.value = pid;

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

      // Set VTT track
      var track = qs("#subtitleTrack");
      track.src = "api/vtt/" + pid;
    } else {
      video.removeAttribute("src");
      video.load();
      video.classList.add("hidden");
      videoEmpty.classList.remove("hidden");
    }

    // Update status
    var statusEl = qs("#transcriptStatus");
    var taskForPid = null;
    state.tasks.forEach(function (t) {
      if (t.participant === pid && (t.status === "running" || t.status === "queued")) {
        taskForPid = t;
      }
    });
    if (p.has_transcript) {
      statusEl.textContent = p.segment_count + " segments";
      if (p.has_stale_artifacts) {
        statusEl.textContent += " \u2022 artifacts outdated";
        statusEl.classList.add("transcript-stale");
      } else {
        statusEl.classList.remove("transcript-stale");
      }
    } else if (taskForPid && taskForPid.status === "running") {
      statusEl.textContent = "transcribing\u2026 " + Math.round((taskForPid.progress || 0) * 100) + "%";
      statusEl.classList.remove("transcript-stale");
    } else if (taskForPid && taskForPid.status === "queued") {
      statusEl.textContent = "queued";
      statusEl.classList.remove("transcript-stale");
    } else {
      statusEl.textContent = "not transcribed";
      statusEl.classList.remove("transcript-stale");
    }

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
      clearSummary();
    }
  }

  function renderEmptyState() {
    qs("#videoPlayer").classList.add("hidden");
    qs("#videoEmpty").classList.remove("hidden");
    qs("#segmentList").innerHTML = "";
    qs("#transcriptEmpty").classList.remove("hidden");
    clearSummary();
  }

  // ---- Transcript loading ----

  function loadTranscript(pid) {
    return apiGet("api/transcript/" + pid).then(function (data) {
      if (!data.ok) {
        state.segments = [];
        renderSegments();
        return;
      }
      state.segments = data.segments;
      state.activeSegmentIndex = -1;
      renderSegments();
    });
  }

  // ---- AI Summary ----

  var _summaryPollTimer = null;

  function loadSummary(pid) {
    var section = qs("#summarySection");

    apiGet("api/summary/" + pid).then(function (data) {
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
        }
      } else if (data.generating) {
        renderSummaryGenerating();
        _startSummaryPoll(pid);
      } else {
        section.classList.add("hidden");
      }
    }).catch(function () {
      // Ollama unavailable or no summary — stay hidden
      section.classList.add("hidden");
    });
  }

  function _startSummaryPoll(pid) {
    _stopSummaryPoll();
    _summaryPollTimer = setInterval(function () {
      if (state.selectedParticipant !== pid) {
        _stopSummaryPoll();
        return;
      }
      apiGet("api/summary/" + pid).then(function (data) {
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
        sup.title = fmtTime(ref.start);
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

  var _CITATIONS_POLL_TIMEOUT = 90000; // stop polling after 90 seconds

  function _startCitationsPoll(pid) {
    _stopCitationsPoll();
    var started = Date.now();
    _citationsPollTimer = setInterval(function () {
      if (state.selectedParticipant !== pid || Date.now() - started > _CITATIONS_POLL_TIMEOUT) {
        _stopCitationsPoll();
        state.citationsGenerating = false;
        var status = qs("#summaryContent .citations-status");
        if (status) status.remove();
        return;
      }
      apiGet("api/citations/" + pid).then(function (data) {
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
      var markClass = markObj ? "segment-mark marked" : "segment-mark";
      var markStyle = markObj ? ' style="background:' + (MARK_CATEGORIES[markObj.category] || MARK_CATEGORIES.bookmark).color + '"' : "";
      var markLabel = markObj && markObj.label ? ' title="' + escapeHtml(markObj.label) + '"' : "";

      html += '<div class="segment-row' + activeClass + correctedClass + '" data-index="' + i + '" data-start="' + seg.start + '">';
      html += '<span class="' + markClass + '" data-segment-id="' + escapeHtml(seg.id) + '"' + markStyle + markLabel + '></span>';
      html += '<span class="segment-timestamp">' + fmtTime(seg.start);
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
            html += '<span class="segment-xref-badge" style="background:' + XREF_BADGES.screenspace.color + '" title="' + escapeHtml(evTypes.join(", ")) + '"><span class="xref-badge-icon" style="mask-image:url(icons/' + XREF_BADGES.screenspace.icon + '.svg);-webkit-mask-image:url(icons/' + XREF_BADGES.screenspace.icon + '.svg)"></span></span>';
          }
          if (xref.sheetObservations.length > 0) {
            var obsTitle = xref.sheetObservations[0].observation;
            html += '<span class="segment-xref-badge" style="background:' + XREF_BADGES.sheet.color + '" title="' + escapeHtml(obsTitle) + '"><span class="xref-badge-icon" style="mask-image:url(icons/' + XREF_BADGES.sheet.icon + '.svg);-webkit-mask-image:url(icons/' + XREF_BADGES.sheet.icon + '.svg)"></span></span>';
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
          wordHtml += '<span class="segment-word" data-original="' + escapeHtml(tokens[w]) + '">' + escapeHtml(tokens[w]) + '</span>';
        }
      }
      html += '<span class="segment-text" data-id="' + escapeHtml(seg.id) + '">' + wordHtml + '</span>';
      html += '<span class="segment-copy" title="Copy text"><span class="segment-copy-icon"></span></span>';
      html += '</div>';
    }
    container.innerHTML = html;

    // Attach event listeners
    var rows = container.querySelectorAll(".segment-row");
    for (var j = 0; j < rows.length; j++) {
      (function (row) {
        var textEl = row.querySelector(".segment-text");
        var markEl = row.querySelector(".segment-mark");
        row.querySelector(".segment-timestamp").addEventListener("click", function (e) {
          e.stopPropagation();
          var start = parseFloat(row.getAttribute("data-start"));
          seekVideo(start);
        });
        textEl.addEventListener("click", function (e) {
          e.stopPropagation();
          if (state.editingTextEl === textEl) return;
          var start = parseFloat(row.getAttribute("data-start"));
          seekVideo(start);
        });
        textEl.addEventListener("dblclick", function (e) {
          e.stopPropagation();
          startSegmentEditing(textEl);
        });
        markEl.addEventListener("click", function (e) {
          e.stopPropagation();
          var segId = markEl.getAttribute("data-segment-id");
          var idx = parseInt(row.getAttribute("data-index"), 10);
          var seg = state.segments[idx];
          var mark = seg && seg.marks && seg.marks.length > 0 ? seg.marks[0] : null;
          if (mark) {
            showMarkPopover(markEl, segId, mark);
          } else {
            toggleMark(segId);
          }
        });
        row.querySelector(".segment-copy").addEventListener("click", function (e) {
          e.stopPropagation();
          var idx = parseInt(row.getAttribute("data-index"), 10);
          var seg = state.segments[idx];
          if (!seg) return;
          navigator.clipboard.writeText(seg.text).then(function () {
            showToast("Copied to clipboard");
          });
        });
      })(rows[j]);
    }
  }

  function renderPartialSegments(segments, progress) {
    var container = qs("#segmentList");
    var empty = qs("#transcriptEmpty");
    var pid = state.streamingParticipant || state.selectedParticipant;
    empty.classList.add("hidden");

    // If user is actively editing a segment, skip DOM rebuild to preserve edit state
    if (state.editingTextEl && state.editingTextEl.isConnected) return;

    // Invalidate cached segment rows since we're rebuilding DOM
    _cachedSegmentRows = null;

    // Only auto-scroll if user is near the bottom
    var nearBottom = container.scrollHeight - container.scrollTop - container.clientHeight < 100;

    var html = "";
    for (var i = 0; i < segments.length; i++) {
      var seg = segments[i];
      var segId = pid + ":" + i;
      var cachedMark = _streamingMarks[segId];
      var cachedColor = cachedMark ? cachedMark.color : null;
      var markClass = "segment-mark" + (cachedColor ? " marked" : "");
      var markStyle = cachedColor ? ' style="background:' + cachedColor + '"' : "";
      html += '<div class="segment-row segment-streaming" data-index="' + i + '" data-start="' + seg.start + '">';
      html += '<span class="' + markClass + '" data-segment-id="' + escapeHtml(segId) + '"' + markStyle + '></span>';
      html += '<span class="segment-timestamp">' + fmtTime(seg.start) + '</span>';
      html += '<span class="segment-text">' + escapeHtml(seg.text) + '</span>';
      html += '<span class="segment-copy" title="Copy text"><span class="segment-copy-icon"></span></span>';
      html += '</div>';
    }
    html += '<div class="streaming-indicator">';
    html += '<span class="streaming-dot"></span>';
    html += 'Transcribing\u2026 ' + Math.round(progress * 100) + '%';
    html += '</div>';

    container.innerHTML = html;

    // Click-to-seek on timestamps/text, dblclick-to-edit, and mark clicks
    var rows = container.querySelectorAll(".segment-streaming");
    for (var j = 0; j < rows.length; j++) {
      (function (row) {
        var textEl = row.querySelector(".segment-text");
        row.querySelector(".segment-timestamp").addEventListener("click", function (e) {
          e.stopPropagation();
          seekVideo(parseFloat(row.getAttribute("data-start")));
        });
        textEl.addEventListener("click", function (e) {
          e.stopPropagation();
          if (state.editingTextEl === textEl) return;
          seekVideo(parseFloat(row.getAttribute("data-start")));
        });
        textEl.addEventListener("dblclick", function (e) {
          e.stopPropagation();
          startSegmentEditing(textEl);
        });
        row.querySelector(".segment-mark").addEventListener("click", function (e) {
          e.stopPropagation();
          var segId = this.getAttribute("data-segment-id");
          var existing = _streamingMarks[segId];
          if (existing) {
            showMarkPopover(this, segId, existing);
          } else {
            toggleMarkStreaming(segId, this);
          }
        });
        row.querySelector(".segment-copy").addEventListener("click", function (e) {
          e.stopPropagation();
          var idx = parseInt(row.getAttribute("data-index"), 10);
          var seg = segments[idx];
          if (!seg) return;
          navigator.clipboard.writeText(seg.text).then(function () {
            showToast("Copied to clipboard");
          });
        });
      })(rows[j]);
    }

    if (nearBottom) {
      container.scrollTop = container.scrollHeight;
    }
  }

  // Cache marks made during streaming so they survive DOM rebuilds.
  // Each entry: { color, id, category, label }
  var _streamingMarks = {};
  var _streamingMarksLoaded = false;

  function _loadStreamingMarks(pid) {
    if (_streamingMarksLoaded) return;
    _streamingMarksLoaded = true;
    apiGet("api/marks").then(function (data) {
      if (!data.ok) return;
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
      });
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
    video.addEventListener("timeupdate", function () {
      if (_syncRaf) return;
      _syncRaf = requestAnimationFrame(function () {
        _syncRaf = 0;
        highlightActiveSegment();
      });
    });
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
      scrollToSegment(rows[newIndex]);
    }
  }

  function scrollToSegment(row) {
    var section = qs("#transcriptSection");
    var rowTop = row.offsetTop;
    var rowBottom = rowTop + row.offsetHeight;
    var scrollTop = section.scrollTop;
    var viewHeight = section.clientHeight;

    if (rowTop < scrollTop + 40) {
      section.scrollTop = rowTop - 40;
    } else if (rowBottom > scrollTop + viewHeight - 40) {
      section.scrollTop = rowBottom - viewHeight + 40;
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
    var i = m, j = n;
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
    var created = 0, updated = 0, removed = 0;
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
            if (_streamingMarks[key].id === markId) { delete _streamingMarks[key]; break; }
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

    // Close on outside click (deferred so this click doesn't trigger it)
    setTimeout(function () {
      document.addEventListener("click", _popoverOutsideClick);
    }, 0);
  }

  function _popoverOutsideClick(e) {
    var popover = qs("#markPopover");
    if (popover && !popover.contains(e.target)) {
      hideMarkPopover();
    }
  }

  function hideMarkPopover() {
    var popover = qs("#markPopover");
    if (popover) popover.classList.add("hidden");
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
        showToast("Marked " + data.marks.length + " segment" + (data.marks.length === 1 ? "" : "s"));
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
        html += '<span class="search-result-time">' + fmtTime(r.start) + '</span>';
        html += '<span class="search-result-text">' + highlightQuery(r.text, state.searchQuery) + '</span>';
        if (state.tooltipsEnabled && xref.screenspaceEvents.length > 0) {
          var seen = {};
          html += '<span class="search-xref-events">';
          for (var ei = 0; ei < xref.screenspaceEvents.length; ei++) {
            var det = xref.screenspaceEvents[ei].detector;
            if (seen[det]) continue;
            seen[det] = true;
            var evColor = SS_DETECTOR_COLORS[det] || "#888";
            html += '<span class="search-xref-dot" style="background:' + evColor + '" title="' + escapeHtml(xref.screenspaceEvents[ei].event_type || det) + '"></span>';
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

    container.innerHTML = html;
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

  // ---- Queue panel ----

  function initQueuePanel() {
    qs("#queueBtn").addEventListener("click", function () {
      var panel = qs("#queuePanel");
      panel.classList.toggle("hidden");
      if (!panel.classList.contains("hidden")) {
        renderQueue();
        pollTaskStatus();
      }
    });

    qs("#closeQueueBtn").addEventListener("click", function () {
      qs("#queuePanel").classList.add("hidden");
    });

    qs("#transcribeAllBtn").addEventListener("click", function () {
      transcribeAll();
    });
  }

  function queueItemState(p, taskByPid) {
    var task = taskByPid[p.id];
    var status = "not started";
    var statusClass = "";
    var progress = 0;
    var showProgress = false;
    var taskId = null;

    if (p.has_transcript && (!task || task.status === "completed")) {
      status = "completed";
      statusClass = "status-completed";
      progress = 100;
    } else if (task) {
      status = task.status;
      taskId = task.id;
      if (task.status === "running") {
        statusClass = "status-running";
        progress = Math.round(task.progress * 100);
        showProgress = true;
      } else if (task.status === "queued") {
        statusClass = "";
      } else if (task.status === "failed") {
        statusClass = "status-failed";
        status = "failed";
      } else if (task.status === "cancelled") {
        statusClass = "status-cancelled";
      } else if (task.status === "completed") {
        statusClass = "status-completed";
        progress = 100;
      }
    }
    return { status: status, statusClass: statusClass, progress: progress, showProgress: showProgress, taskId: taskId };
  }

  function renderQueue() {
    var container = qs("#queueList");
    if (!container) return;

    var taskByPid = {};
    state.tasks.forEach(function (t) {
      if (!taskByPid[t.participant] || t.created_at > taskByPid[t.participant].created_at) {
        taskByPid[t.participant] = t;
      }
    });

    // Try in-place update when item list and status categories match (avoids layout thrash)
    var existing = container.querySelectorAll(".queue-item[data-pid]");
    if (existing.length === state.participants.length) {
      var canPatch = true;
      for (var k = 0; k < state.participants.length; k++) {
        var s = queueItemState(state.participants[k], taskByPid);
        if (existing[k].getAttribute("data-pid") !== state.participants[k].id ||
            existing[k].getAttribute("data-status") !== s.status) {
          canPatch = false; break;
        }
      }
      if (canPatch) {
        for (var k = 0; k < state.participants.length; k++) {
          var item = existing[k];
          var s = queueItemState(state.participants[k], taskByPid);
          var statusEl = item.querySelector(".queue-item-status");
          statusEl.className = "queue-item-status" + (s.statusClass ? " " + s.statusClass : "");
          statusEl.textContent = s.status + (s.showProgress ? " " + s.progress + "%" : "");
          var bar = item.querySelector(".queue-progress-bar");
          if (bar) bar.style.width = s.progress + "%";
        }
        return;
      }
    }

    // Full rebuild when structure changes
    var html = "";
    state.participants.forEach(function (p) {
      var s = queueItemState(p, taskByPid);
      html += '<div class="queue-item" data-pid="' + escapeHtml(p.id) + '" data-status="' + s.status + '">';
      html += '<div class="queue-item-header">';
      html += '<span class="queue-item-id">' + escapeHtml(p.id) + '</span>';
      html += '<span class="queue-item-status ' + s.statusClass + '">' + s.status + (s.showProgress ? " " + s.progress + "%" : "") + '</span>';
      if (s.taskId && (s.status === "running" || s.status === "queued")) {
        html += '<button class="queue-item-cancel" data-cancel-task="' + escapeHtml(s.taskId) + '" title="Cancel">&times;</button>';
      }
      html += '</div>';

      if (s.showProgress || s.progress === 100) {
        html += '<div class="queue-progress"><div class="queue-progress-bar" style="width:' + s.progress + '%"></div></div>';
      }

      if (!p.has_transcript && (!taskByPid[p.id] || taskByPid[p.id].status === "failed" || taskByPid[p.id].status === "cancelled")) {
        html += '<div class="queue-item-action"><button class="btn btn-small" data-transcribe="' + escapeHtml(p.id) + '">Transcribe</button></div>';
      } else if (p.has_transcript && (!taskByPid[p.id] || taskByPid[p.id].status === "completed")) {
        html += '<div class="queue-item-action"><button class="btn btn-small" data-retranscribe="' + escapeHtml(p.id) + '">Re-transcribe</button>';
        if (p.has_stale_artifacts) {
          html += '<span class="queue-stale-badge">artifacts outdated</span>';
        }
        html += '</div>';
      }

      html += '</div>';
    });

    container.innerHTML = html;

    // Attach event listeners
    var transcribeBtns = container.querySelectorAll("[data-transcribe]");
    for (var i = 0; i < transcribeBtns.length; i++) {
      transcribeBtns[i].addEventListener("click", function () {
        var pid = this.getAttribute("data-transcribe");
        transcribeParticipants([pid], false);
      });
    }

    var retranscribeBtns = container.querySelectorAll("[data-retranscribe]");
    for (var j = 0; j < retranscribeBtns.length; j++) {
      retranscribeBtns[j].addEventListener("click", function () {
        var pid = this.getAttribute("data-retranscribe");
        transcribeParticipants([pid], true);
      });
    }

    var cancelBtns = container.querySelectorAll("[data-cancel-task]");
    for (var c = 0; c < cancelBtns.length; c++) {
      cancelBtns[c].addEventListener("click", function () {
        var taskId = this.getAttribute("data-cancel-task");
        apiDelete("api/transcribe/" + taskId).then(function () {
          pollTaskStatus();
        });
      });
    }
  }

  function transcribeAll() {
    var pids = [];
    state.participants.forEach(function (p) {
      if (p.has_video && !p.has_transcript) pids.push(p.id);
    });
    if (pids.length === 0) {
      showToast("All participants already transcribed");
      return;
    }
    transcribeParticipants(pids, false);
  }

  function transcribeParticipants(pids, force) {
    apiPost("api/transcribe", { participants: pids, force: force }).then(function (data) {
      if (!data.ok) {
        showToast("Failed to enqueue transcription");
        return;
      }
      showToast("Enqueued " + data.tasks.length + " transcription(s)");
      startPolling();
      pollTaskStatus();
    });
  }

  // ---- Task polling ----

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

  function pollTaskStatus() {
    apiGet("api/transcribe/status").then(function (data) {
      if (!data.ok) return;
      state.tasks = data.tasks;

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
        // Update status text
        var statusEl = qs("#transcriptStatus");
        if (statusEl) statusEl.textContent = "transcribing\u2026 " + Math.round(selectedRunningTask.progress * 100) + "%";
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

      // Refresh participants and transcript as each task completes
      if (newlyCompleted.length > 0) {
        _streamingMarks = {};
        _streamingMarksLoaded = false;
        loadParticipants().then(function () {
          if (state.selectedParticipant && newlyCompleted.indexOf(state.selectedParticipant) >= 0) {
            state.streamingParticipant = null;
            loadTranscript(state.selectedParticipant);
            loadSummary(state.selectedParticipant);
          }
        });
      }

      if (hasActive) {
        startPolling();
      } else {
        stopPolling();
        _refreshedCompletedPids = {};
      }

      // Update queue panel if visible
      if (!qs("#queuePanel").classList.contains("hidden")) {
        renderQueue();
      }
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

  // ---- Settings Popover ----

  var _trModelsCache = null;
  var _trModelsCachePromise = null;
  var _trSettingsData = null;
  var _trSettingsSaveTimer = null;

  var TR_SETTINGS_GROUPS = ["Transcription", "AI Summary"];

  function _trFetchModels() {
    if (_trModelsCache) return Promise.resolve(_trModelsCache);
    if (_trModelsCachePromise) return _trModelsCachePromise;
    _trModelsCachePromise = fetch("../api/models")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data.ok) _trModelsCache = data;
        return data;
      })
      .catch(function () { return null; });
    return _trModelsCachePromise;
  }

  function _trFormatSize(mb) {
    if (mb >= 1024) return (mb / 1024).toFixed(1) + " GB";
    return mb + " MB";
  }

  function _trLoadModelsForSelect(sel, provider, currentValue) {
    _trFetchModels().then(function (data) {
      if (!data || !data.ok) { sel.disabled = false; return; }

      var models = [];
      if (provider === "whisper") {
        models = (data.whisper && data.whisper.models) || [];
      } else if (provider === "ollama") {
        models = (data.ollama && data.ollama.models) || [];
      }

      sel.innerHTML = "";
      var hasCurrentValue = false;
      for (var i = 0; i < models.length; i++) {
        var m = models[i];
        var opt = document.createElement("option");
        opt.value = m.name;
        var label = m.name;
        if (m.size_mb) label += " (" + _trFormatSize(m.size_mb) + ")";
        if (m.parameter_size) label += " \u00B7 " + m.parameter_size;
        if (m.description) label += " \u2014 " + m.description;
        opt.textContent = label;
        if (m.name === currentValue) {
          opt.selected = true;
          hasCurrentValue = true;
        }
        sel.appendChild(opt);
      }
      if (!hasCurrentValue && currentValue) {
        var custom = document.createElement("option");
        custom.value = currentValue;
        custom.textContent = currentValue + " (current)";
        custom.selected = true;
        sel.insertBefore(custom, sel.firstChild);
      }
      sel.disabled = false;
    });
  }

  function _trFindSetting(name) {
    if (!_trSettingsData) return null;
    for (var i = 0; i < _trSettingsData.length; i++) {
      if (_trSettingsData[i].name === name) return _trSettingsData[i];
    }
    return null;
  }

  function _trScheduleSave() {
    if (_trSettingsSaveTimer) clearTimeout(_trSettingsSaveTimer);
    _trSettingsSaveTimer = setTimeout(_trSaveSettings, 400);
  }

  function _trSaveSettings() {
    if (!_trSettingsData) return;
    var payload = {};
    for (var i = 0; i < _trSettingsData.length; i++) {
      var s = _trSettingsData[i];
      if (s.value !== s.default) payload[s.name] = s.value;
    }
    var statusEl = qs("#trSettingsSaveStatus");
    fetch("../api/settings", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ settings: payload }),
    })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (statusEl) {
          statusEl.textContent = data.ok ? "Saved" : "Error";
          setTimeout(function () { statusEl.textContent = ""; }, 1500);
        }
      })
      .catch(function () {
        if (statusEl) statusEl.textContent = "Error";
      });
  }

  function _trUpdateChanged(settingName) {
    var row = qs('.tr-settings-row[data-setting="' + settingName + '"]');
    if (!row) return;
    var s = _trFindSetting(settingName);
    if (!s) return;
    if (s.value !== s.default) {
      row.classList.add("settings-changed");
    } else {
      row.classList.remove("settings-changed");
    }
  }

  function _trBuildRow(s) {
    var row = document.createElement("div");
    row.className = "tr-settings-row";
    if (s.value !== s.default) row.classList.add("settings-changed");
    row.setAttribute("data-setting", s.name);

    var labelDiv = document.createElement("div");
    labelDiv.className = "tr-settings-label";
    var nameEl = document.createElement("div");
    nameEl.className = "tr-settings-label-name";
    nameEl.textContent = s.name
      .replace(/_/g, " ").toLowerCase()
      .replace(/\b\w/g, function (c) { return c.toUpperCase(); })
      .replace(/Mb$/i, "(MB)").replace(/Seconds$/i, "(s)");
    var descEl = document.createElement("div");
    descEl.className = "tr-settings-label-desc";
    descEl.textContent = s.description;
    labelDiv.appendChild(nameEl);
    labelDiv.appendChild(descEl);

    var controlDiv = document.createElement("div");
    controlDiv.className = "tr-settings-control";
    var settingName = s.name;

    if (s.type === "bool") {
      var toggle = document.createElement("input");
      toggle.type = "checkbox";
      toggle.checked = !!s.value;
      toggle.addEventListener("change", function () {
        var setting = _trFindSetting(settingName);
        if (setting) setting.value = this.checked;
        _trUpdateChanged(settingName);
        _trScheduleSave();
      });
      controlDiv.appendChild(toggle);
    } else if (s.type === "select" && s.options) {
      var sel = document.createElement("select");
      for (var oi = 0; oi < s.options.length; oi++) {
        var opt = document.createElement("option");
        opt.value = s.options[oi];
        opt.textContent = s.options[oi];
        if (s.options[oi] === s.value) opt.selected = true;
        sel.appendChild(opt);
      }
      sel.addEventListener("change", function () {
        var setting = _trFindSetting(settingName);
        if (setting) setting.value = this.value;
        _trUpdateChanged(settingName);
        _trScheduleSave();
      });
      controlDiv.appendChild(sel);
    } else if (s.type === "model_select") {
      var msel = document.createElement("select");
      var curOpt = document.createElement("option");
      curOpt.value = s.value;
      curOpt.textContent = s.value;
      curOpt.selected = true;
      msel.appendChild(curOpt);
      msel.disabled = true;
      msel.addEventListener("change", function () {
        var setting = _trFindSetting(settingName);
        if (setting) setting.value = this.value;
        _trUpdateChanged(settingName);
        _trScheduleSave();
      });
      controlDiv.appendChild(msel);
      _trLoadModelsForSelect(msel, s.provider, s.value);
    } else if (s.type === "str") {
      var txtInput = document.createElement("input");
      txtInput.type = "text";
      txtInput.autocomplete = "off";
      txtInput.value = s.value || "";
      txtInput.placeholder = String(s.default || "");
      txtInput.addEventListener("change", function () {
        var setting = _trFindSetting(settingName);
        if (setting) setting.value = this.value;
        _trUpdateChanged(settingName);
        _trScheduleSave();
      });
      controlDiv.appendChild(txtInput);
    }

    row.appendChild(labelDiv);
    row.appendChild(controlDiv);
    return row;
  }

  function _trRenderSettings() {
    var container = qs("#trSettingsContent");
    if (!container || !_trSettingsData) return;
    container.innerHTML = "";

    var groups = {};
    for (var i = 0; i < _trSettingsData.length; i++) {
      var s = _trSettingsData[i];
      if (TR_SETTINGS_GROUPS.indexOf(s.group) === -1) continue;
      if (!groups[s.group]) groups[s.group] = [];
      groups[s.group].push(s);
    }

    for (var gi = 0; gi < TR_SETTINGS_GROUPS.length; gi++) {
      var groupName = TR_SETTINGS_GROUPS[gi];
      if (!groups[groupName]) continue;
      var header = document.createElement("div");
      header.className = "tr-settings-group-label";
      header.textContent = groupName;
      container.appendChild(header);
      var items = groups[groupName];
      for (var si = 0; si < items.length; si++) {
        container.appendChild(_trBuildRow(items[si]));
      }
    }
  }

  function _trLoadSettings() {
    fetch("../api/settings")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (!data.ok) return;
        _trSettingsData = data.settings;
        _trRenderSettings();
      })
      .catch(function () {});
  }

  function initTranscriptSettings() {
    var btn = qs("#trSettingsBtn");
    var popover = qs("#trSettingsPopover");
    var closeBtn = qs("#trSettingsClose");
    if (!btn || !popover) return;

    btn.addEventListener("click", function () {
      var isOpen = !popover.classList.contains("hidden");
      if (isOpen) {
        popover.classList.add("hidden");
      } else {
        popover.classList.remove("hidden");
        _trModelsCache = null;
        _trModelsCachePromise = null;
        _trLoadSettings();
      }
    });

    if (closeBtn) {
      closeBtn.addEventListener("click", function () {
        popover.classList.add("hidden");
      });
    }

    document.addEventListener("click", function (e) {
      if (popover.classList.contains("hidden")) return;
      if (popover.contains(e.target) || btn.contains(e.target)) return;
      popover.classList.add("hidden");
    });
  }

  // ---- Boot ----

  document.addEventListener("DOMContentLoaded", function () {
    initThemeToggle();
    initTooltipToggle();
    checkNavLinks();
    initSearch();
    initQueuePanel();
    initCorrectionsModal();
    initVideoSync();
    initSummaryToggle();
    initSummaryActions();
    initTranscriptSettings();

    // Pause polling when tab is hidden; resume when visible
    document.addEventListener("visibilitychange", function () {
      if (document.hidden) {
        stopPolling();
      } else {
        pollTaskStatus();
      }
    });

    // Participant select handler
    qs("#participantSelect").addEventListener("change", function () {
      if (this.value) selectParticipant(this.value);
    });

    // Load initial data
    loadParticipants();

    // Check for active tasks on load
    pollTaskStatus();
  });

})();
