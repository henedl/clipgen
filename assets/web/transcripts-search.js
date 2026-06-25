/* clipgen Transcripts search satellite — transcripts-search.js
 *
 * Header full-text search across all transcripts: debounced query, client-side
 * merge of the streaming participant's partial segments, grouped results with
 * cross-reference badges, "Mark All", and jump-to-result. Loaded after
 * transcripts.js; reads the hub's shared state + helpers through
 * window.ClipgenTranscripts (TS) and publishes initSearch / renderSearchResults
 * back (boot wires initSearch; the xref data-load and tooltip toggle re-render
 * open results). seekVideo lives in the video satellite, so jumpToResult reaches
 * it late-bound via TS.seekVideo. Plain utils.js globals (qs/apiGet/apiPost/
 * escapeHtml/formatTime/clipgenPluralUnit) are reached via the scope chain.
 */
(function () {
  "use strict";

  var TS = window.ClipgenTranscripts;
  var state = TS.state;
  var showToast = TS.showToast,
    loadTranscript = TS.loadTranscript,
    findOverlapsForSearch = TS.findOverlapsForSearch,
    selectParticipant = TS.selectParticipant;

  var SEARCH_DEBOUNCE = 300;
  var _searchTimer = null;

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
      markAllBtn.addEventListener("click", function () { markAllSearchResults(); });
    }
    markAllBtn.classList.remove("hidden");

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
          var truncObs = obsText.length > 50 ? obsText.substring(0, 50) + "…" : obsText;
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
    // seekVideo lives in the video satellite; late-bind through TS.
    if (TS.seekVideo) TS.seekVideo(start);
  }

  // ---- Published back to the hub (boot wires initSearch; xref/tooltip re-render) ----
  TS.initSearch = initSearch;
  TS.renderSearchResults = renderSearchResults;
})();
