(function () {
  "use strict";

  var THEME_STORAGE_KEY = "clipgen-transcripts-theme";
  var POLL_INTERVAL = 3000;
  var SEARCH_DEBOUNCE = 300;

  // Mark categories — mirrored from transcripts.py MARK_CATEGORIES
  var MARK_CATEGORIES = {
    pain_point: { label: "Pain Point", color: "#dc2626" },
    delight:    { label: "Delight",    color: "#16a34a" },
    quote:      { label: "Quote",      color: "#2563eb" },
    insight:    { label: "Insight",    color: "#f97316" },
    task:       { label: "Task Issue", color: "#8b5cf6" },
    bookmark:   { label: "Bookmark",   color: "#0891b2" },
  };

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
    editingTextEl: null,
    pollTimer: null,
    lastMarkCategory: "bookmark",
    streamingParticipant: null,
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
    } else if (taskForPid && taskForPid.status === "running") {
      statusEl.textContent = "transcribing\u2026 " + Math.round((taskForPid.progress || 0) * 100) + "%";
    } else if (taskForPid && taskForPid.status === "queued") {
      statusEl.textContent = "queued";
    } else {
      statusEl.textContent = "not transcribed";
    }

    // Load transcript
    if (p.has_transcript) {
      state.streamingParticipant = null;
      loadTranscript(pid);
    } else if (taskForPid && taskForPid.status === "running" && taskForPid.partial_segments && taskForPid.partial_segments.length > 0) {
      renderPartialSegments(taskForPid.partial_segments, taskForPid.progress);
      state.streamingParticipant = pid;
    } else {
      state.segments = [];
      state.streamingParticipant = null;
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
      html += '<span class="segment-timestamp">' + fmtTime(seg.start) + '</span>';
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
      })(rows[j]);
    }
  }

  function renderPartialSegments(segments, progress) {
    var container = qs("#segmentList");
    var empty = qs("#transcriptEmpty");
    var pid = state.streamingParticipant || state.selectedParticipant;
    empty.classList.add("hidden");

    // Only auto-scroll if user is near the bottom
    var nearBottom = container.scrollHeight - container.scrollTop - container.clientHeight < 100;

    var html = "";
    for (var i = 0; i < segments.length; i++) {
      var seg = segments[i];
      var segId = pid + ":" + i;
      html += '<div class="segment-row segment-streaming" data-index="' + i + '" data-start="' + seg.start + '">';
      html += '<span class="segment-mark" data-segment-id="' + escapeHtml(segId) + '"></span>';
      html += '<span class="segment-timestamp">' + fmtTime(seg.start) + '</span>';
      html += '<span class="segment-text">' + escapeHtml(seg.text) + '</span>';
      html += '</div>';
    }
    html += '<div class="streaming-indicator">';
    html += '<span class="streaming-dot"></span>';
    html += 'Transcribing\u2026 ' + Math.round(progress * 100) + '%';
    html += '</div>';

    container.innerHTML = html;

    // Click-to-seek on timestamps and mark clicks
    var rows = container.querySelectorAll(".segment-streaming");
    for (var j = 0; j < rows.length; j++) {
      (function (row) {
        row.querySelector(".segment-timestamp").addEventListener("click", function (e) {
          e.stopPropagation();
          seekVideo(parseFloat(row.getAttribute("data-start")));
        });
        row.querySelector(".segment-mark").addEventListener("click", function (e) {
          e.stopPropagation();
          var segId = this.getAttribute("data-segment-id");
          toggleMarkStreaming(segId, this);
        });
      })(rows[j]);
    }

    if (nearBottom) {
      container.scrollTop = container.scrollHeight;
    }
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
      if (data.ok) {
        showToast("Marked");
        markEl.classList.add("marked");
        markEl.style.background = cat.color;
      }
    });
  }

  function removeMark(markId) {
    apiDelete("api/marks/" + markId).then(function (data) {
      if (data.ok) {
        showToast("Mark removed");
        hideMarkPopover();
        if (state.selectedParticipant) loadTranscript(state.selectedParticipant);
      }
    });
  }

  function updateMarkCategory(markId, category) {
    state.lastMarkCategory = category;
    apiPut("api/marks/" + markId, { category: category }).then(function (data) {
      if (data.ok) {
        hideMarkPopover();
        if (state.selectedParticipant) loadTranscript(state.selectedParticipant);
      }
    });
  }

  function updateMarkLabel(markId, label) {
    apiPut("api/marks/" + markId, { label: label || null });
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
    return { status: status, statusClass: statusClass, progress: progress, showProgress: showProgress };
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
      html += '</div>';

      if (s.showProgress || s.progress === 100) {
        html += '<div class="queue-progress"><div class="queue-progress-bar" style="width:' + s.progress + '%"></div></div>';
      }

      if (!p.has_transcript && (!taskByPid[p.id] || taskByPid[p.id].status === "failed" || taskByPid[p.id].status === "cancelled")) {
        html += '<div class="queue-item-action"><button class="btn btn-small" data-transcribe="' + escapeHtml(p.id) + '">Transcribe</button></div>';
      } else if (p.has_transcript && (!taskByPid[p.id] || taskByPid[p.id].status === "completed")) {
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
        loadParticipants().then(function () {
          if (state.selectedParticipant && newlyCompleted.indexOf(state.selectedParticipant) >= 0) {
            state.streamingParticipant = null;
            loadTranscript(state.selectedParticipant);
          }
        });
      }

      if (!hasActive) {
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
