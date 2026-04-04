(function () {
  "use strict";

  var THEME_STORAGE_KEY = "clipgen-transcripts-theme";
  var POLL_INTERVAL = 3000;
  var SEARCH_DEBOUNCE = 300;

  // ---- State ----

  var state = {
    participants: [],
    selectedParticipant: null,
    segments: [],
    corrections: [],
    tasks: [],
    searchQuery: "",
    searchResults: null,
    activeSegmentIndex: -1,
    editingSegmentId: null,
    pollTimer: null,
  };

  // ---- Helpers ----

  function qs(sel) { return document.querySelector(sel); }

  function apiGet(path) {
    return fetch(path).then(function (r) { return r.json(); });
  }

  function apiPost(path, body) {
    return fetch(path, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    }).then(function (r) { return r.json(); });
  }

  function apiPut(path, body) {
    return fetch(path, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    }).then(function (r) { return r.json(); });
  }

  function apiDelete(path) {
    return fetch(path, { method: "DELETE" }).then(function (r) { return r.json(); });
  }

  function fmtTime(seconds) {
    var total = Math.floor(seconds);
    var h = Math.floor(total / 3600);
    var m = Math.floor((total % 3600) / 60);
    var s = total % 60;
    if (h > 0) return h + ":" + pad2(m) + ":" + pad2(s);
    return m + ":" + pad2(s);
  }

  function pad2(n) { return n < 10 ? "0" + n : "" + n; }

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

  // ---- Theme toggle ----

  function initThemeToggle() {
    applyStoredThemePreference();
    var btn = qs("#themeToggle");
    if (!btn) return;
    btn.addEventListener("click", function () {
      toggleThemePreference();
    });
  }

  function applyStoredThemePreference() {
    var stored = null;
    try { stored = window.localStorage.getItem(THEME_STORAGE_KEY); } catch (_) {}
    var root = document.documentElement;
    if (stored === "light" || stored === "dark") {
      root.setAttribute("data-theme", stored);
    } else {
      root.removeAttribute("data-theme");
    }
    updateThemeToggleButton(stored);
  }

  function toggleThemePreference() {
    var root = document.documentElement;
    var current = root.getAttribute("data-theme");
    var next;
    if (current === "dark") {
      next = "light";
    } else if (current === "light") {
      next = "dark";
    } else {
      var prefersDark = false;
      try {
        prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
      } catch (_) {}
      next = prefersDark ? "light" : "dark";
    }
    root.setAttribute("data-theme", next);
    try { window.localStorage.setItem(THEME_STORAGE_KEY, next); } catch (_) {}
    updateThemeToggleButton(next);
  }

  function updateThemeToggleButton(theme) {
    var btn = qs("#themeToggle");
    if (!btn) return;
    btn.setAttribute("data-theme", theme || "");
    btn.setAttribute("aria-pressed", theme === "dark" ? "true" : "false");
  }

  // ---- Nav links ----

  function checkNavLinks() {
    fetch("../api/status").then(function (r) { return r.json(); }).then(function (data) {
      if (data.studio) qs("#studioLink").classList.remove("hidden");
      if (data.insights) qs("#insightsLink").classList.remove("hidden");
      if (data.screenspace) qs("#screenspaceLink").classList.remove("hidden");
    }).catch(function () {});
  }

  // ---- Participants ----

  function loadParticipants() {
    return apiGet("api/participants").then(function (data) {
      if (!data.ok) return;
      state.participants = data.participants;
      renderParticipantSelect();
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
      opt.textContent = p.id + (p.has_transcript ? " \u2713" : "");
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

    // Set video source
    var video = qs("#videoPlayer");
    var videoEmpty = qs("#videoEmpty");
    if (p.has_video) {
      video.src = "media/" + p.video_filename;
      video.classList.remove("hidden");
      videoEmpty.classList.add("hidden");

      // Set VTT track
      var track = qs("#subtitleTrack");
      track.src = "api/vtt/" + pid;
    } else {
      video.removeAttribute("src");
      video.classList.add("hidden");
      videoEmpty.classList.remove("hidden");
    }

    // Update status
    var statusEl = qs("#transcriptStatus");
    if (p.has_transcript) {
      statusEl.textContent = p.segment_count + " segments";
    } else {
      statusEl.textContent = "not transcribed";
    }

    // Load transcript
    if (p.has_transcript) {
      loadTranscript(pid);
    } else {
      state.segments = [];
      renderSegments();
    }
  }

  function renderEmptyState() {
    qs("#videoPlayer").classList.add("hidden");
    qs("#videoEmpty").classList.remove("hidden");
    qs("#segmentList").innerHTML = "";
    qs("#transcriptEmpty").classList.remove("hidden");
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

  // ---- Segment rendering ----

  function renderSegments() {
    var container = qs("#segmentList");
    var empty = qs("#transcriptEmpty");

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
      html += '<div class="segment-row' + activeClass + correctedClass + '" data-index="' + i + '" data-start="' + seg.start + '">';
      html += '<span class="segment-timestamp">' + fmtTime(seg.start) + '</span>';
      html += '<span class="segment-text" data-id="' + escapeHtml(seg.id) + '">' + escapeHtml(seg.text) + '</span>';
      html += '</div>';
    }
    container.innerHTML = html;

    // Attach event listeners
    var rows = container.querySelectorAll(".segment-row");
    for (var j = 0; j < rows.length; j++) {
      (function (row) {
        row.querySelector(".segment-timestamp").addEventListener("click", function (e) {
          e.stopPropagation();
          var start = parseFloat(row.getAttribute("data-start"));
          seekVideo(start);
        });
        row.querySelector(".segment-text").addEventListener("click", function (e) {
          e.stopPropagation();
          var start = parseFloat(row.getAttribute("data-start"));
          seekVideo(start);
        });
        row.querySelector(".segment-text").addEventListener("dblclick", function (e) {
          e.stopPropagation();
          var segId = this.getAttribute("data-id");
          startEditing(segId, this);
        });
      })(rows[j]);
    }
  }

  function seekVideo(time) {
    var video = qs("#videoPlayer");
    if (video && video.src) {
      video.currentTime = time;
      if (video.paused) video.play();
    }
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

    // Find active segment
    var newIndex = -1;
    for (var i = 0; i < state.segments.length; i++) {
      var seg = state.segments[i];
      if (t >= seg.start && t < seg.end) {
        newIndex = i;
        break;
      }
    }

    if (newIndex === state.activeSegmentIndex) return;

    // Remove old active
    var container = qs("#segmentList");
    var rows = container.querySelectorAll(".segment-row");
    if (state.activeSegmentIndex >= 0 && state.activeSegmentIndex < rows.length) {
      rows[state.activeSegmentIndex].classList.remove("active");
    }

    // Set new active
    state.activeSegmentIndex = newIndex;
    if (newIndex >= 0 && newIndex < rows.length) {
      rows[newIndex].classList.add("active");
      // Auto-scroll to keep active segment visible
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

  // ---- Inline editing ----

  function startEditing(segmentId, textEl) {
    if (state.editingSegmentId === segmentId) return;
    state.editingSegmentId = segmentId;

    var currentText = textEl.textContent;
    var input = document.createElement("textarea");
    input.className = "segment-text-input";
    input.value = currentText;
    input.autocomplete = "off";
    input.rows = 2;
    textEl.innerHTML = "";
    textEl.appendChild(input);
    input.focus();
    input.select();

    function save() {
      var newText = input.value.trim();
      state.editingSegmentId = null;
      if (!newText || newText === currentText) {
        textEl.textContent = currentText;
        return;
      }
      textEl.textContent = newText;
      saveSegmentEdit(segmentId, newText);
    }

    input.addEventListener("blur", save);
    input.addEventListener("keydown", function (e) {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        input.blur();
      }
      if (e.key === "Escape") {
        state.editingSegmentId = null;
        textEl.textContent = currentText;
      }
    });
  }

  function saveSegmentEdit(segmentId, newText) {
    var pid = state.selectedParticipant;
    if (!pid) return;

    apiPut("api/transcript/" + pid + "/segment", {
      segment_id: segmentId,
      text: newText,
    }).then(function (data) {
      if (data.ok && data.correction) {
        showToast("Correction created");
        // Reload transcript to show updated corrections
        loadTranscript(pid);
        loadCorrections();
      }
    }).catch(function () {
      showToast("Failed to save edit");
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
      state.searchResults = data;
      renderSearchResults(data);
    });
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
        html += '<div class="search-result-row" data-participant="' + escapeHtml(r.participant) + '" data-start="' + r.start + '">';
        html += '<span class="search-result-time">' + fmtTime(r.start) + '</span>';
        html += '<span class="search-result-text">' + highlightQuery(r.text, state.searchQuery) + '</span>';
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
    state.searchResults = null;
  }

  function jumpToResult(pid, start) {
    hideSearchResults();
    if (pid !== state.selectedParticipant) {
      selectParticipant(pid);
      // Wait for transcript to load, then seek
      setTimeout(function () { seekVideo(start); }, 300);
    } else {
      seekVideo(start);
    }
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

  function renderQueue() {
    var container = qs("#queueList");
    if (!container) return;

    // Build lookup of task status by participant
    var taskByPid = {};
    state.tasks.forEach(function (t) {
      // Keep the most recent task per participant
      if (!taskByPid[t.participant] || t.created_at > taskByPid[t.participant].created_at) {
        taskByPid[t.participant] = t;
      }
    });

    var html = "";
    state.participants.forEach(function (p) {
      var task = taskByPid[p.id];
      var status = "not started";
      var statusClass = "";
      var progress = 0;
      var showProgress = false;

      if (p.has_transcript && (!task || task.status === "completed")) {
        status = "completed";
        statusClass = "status-completed";
        progress = 100;
      } else if (task) {
        status = task.status;
        if (task.status === "running") {
          statusClass = "status-running";
          progress = Math.round(task.progress * 100);
          showProgress = true;
        } else if (task.status === "queued") {
          statusClass = "";
        } else if (task.status === "failed") {
          statusClass = "status-failed";
          status = "failed";
        } else if (task.status === "completed") {
          statusClass = "status-completed";
          progress = 100;
        }
      }

      html += '<div class="queue-item">';
      html += '<div class="queue-item-header">';
      html += '<span class="queue-item-id">' + escapeHtml(p.id) + '</span>';
      html += '<span class="queue-item-status ' + statusClass + '">' + status + (showProgress ? " " + progress + "%" : "") + '</span>';
      html += '</div>';

      if (showProgress || progress === 100) {
        html += '<div class="queue-progress"><div class="queue-progress-bar" style="width:' + progress + '%"></div></div>';
      }

      if (!p.has_transcript && (!task || task.status === "failed" || task.status === "cancelled")) {
        html += '<div class="queue-item-action"><button class="btn btn-small" data-transcribe="' + escapeHtml(p.id) + '">Transcribe</button></div>';
      } else if (p.has_transcript && (!task || task.status === "completed")) {
        html += '<div class="queue-item-action"><button class="btn btn-small" data-retranscribe="' + escapeHtml(p.id) + '">Re-transcribe</button></div>';
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

  function pollTaskStatus() {
    apiGet("api/transcribe/status").then(function (data) {
      if (!data.ok) return;
      state.tasks = data.tasks;

      // Check if any tasks are still active
      var hasActive = false;
      var justCompleted = [];
      data.tasks.forEach(function (t) {
        if (t.status === "queued" || t.status === "running") hasActive = true;
        if (t.status === "completed") justCompleted.push(t.participant);
      });

      if (!hasActive) {
        stopPolling();
        // Refresh participants to pick up new transcript data
        if (justCompleted.length > 0) {
          loadParticipants().then(function () {
            // Reload current transcript if it was just completed
            if (state.selectedParticipant && justCompleted.indexOf(state.selectedParticipant) >= 0) {
              loadTranscript(state.selectedParticipant);
            }
          });
        }
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

  // ---- Boot ----

  document.addEventListener("DOMContentLoaded", function () {
    initThemeToggle();
    checkNavLinks();
    initSearch();
    initQueuePanel();
    initCorrectionsModal();
    initVideoSync();

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
