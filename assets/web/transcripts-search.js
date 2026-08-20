/* clipgen Transcripts search satellite — transcripts-search.js
 *
 * Header full-text search across all transcripts: debounced query, client-side
 * merge of the streaming participant's partial segments, grouped results,
 * "Mark All", and jump-to-result. Result rows show only time + highlighted text
 * (no cross-reference badges — they smush the narrow dropdown). Loaded after
 * transcripts.js; reads the hub's shared state + helpers through
 * window.ClipgenTranscripts (TS) and publishes initSearch back (boot wires it).
 * seekVideo lives in the video satellite, so
 * jumpToResult reaches it late-bound via TS.seekVideo. Plain utils.js globals
 * (qs/apiGet/apiPost/escapeHtml/formatTime/clipgenPluralUnit) are reached via
 * the scope chain.
 */
(function () {
  "use strict";

  var TS = window.ClipgenTranscripts;
  var state = TS.state;
  var showToast = TS.showToast,
    loadTranscript = TS.loadTranscript,
    selectParticipant = TS.selectParticipant;

  var SEARCH_DEBOUNCE = 300;
  var _searchTimer = null;
  // Request generation: the debounce limits how many requests fire, not their
  // ordering. A slow response for an old query landing after a fast one for
  // the current query must not overwrite state.searchResults — "Mark All"
  // reads it and would persist marks for a query the user already replaced.
  var _searchVer = 0;

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

    // Static markup now, so it is wired once here rather than re-created on
    // every render (which is what the old inject-on-first-result path did).
    qs("#searchMarkAllBtn").addEventListener("click", function () {
      markAllSearchResults();
    });

    input.addEventListener("input", function () {
      clearTimeout(_searchTimer);
      var q = input.value.trim();
      if (q.length < 2) {
        _searchVer++; // invalidate any in-flight response so it can't re-show
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

  function _renderSearchError() {
    var container = qs("#searchResults");
    var list = qs("#searchResultsList");
    if (!container || !list) return;
    qs("#searchResultsHeader").classList.add("hidden");
    list.innerHTML =
      '<div class="search-result-row" style="justify-content:center;color:var(--sev-high)">Search failed — try again</div>';
    container.classList.remove("hidden");
  }

  function doSearch(query) {
    state.searchQuery = query;
    var reqVer = ++_searchVer;
    apiGet("api/search?q=" + encodeURIComponent(query)).then(function (data) {
      if (reqVer !== _searchVer) return; // a newer query superseded this one
      if (!data || !data.ok) { _renderSearchError(); return; }
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
    }).catch(function () {
      if (reqVer !== _searchVer) return;
      _renderSearchError();
    });
  }

  function _searchPartialSegments(query, pid) {
    var results = [];
    var lowerQ = query.toLowerCase();
    // Status polls carry partial_count, not the segment array; the streaming
    // participant's accumulated segments live in the hub (fetched via the tail
    // cursor). This is only called for state.streamingParticipant.
    var segments = TS.streamingSegmentsFor(pid);
    for (var i = 0; i < segments.length; i++) {
      var seg = segments[i];
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
    var list = qs("#searchResultsList");
    var header = qs("#searchResultsHeader");

    // No header on an empty result set: the "No matches found" row already says
    // it, and a Mark All button with nothing to mark is a dead control.
    if (data.total_count === 0) {
      header.classList.add("hidden");
      list.innerHTML = '<div class="search-result-row" style="justify-content:center;color:var(--color-text-dim)">No matches found</div>';
      container.classList.remove("hidden");
      return;
    }

    qs("#searchCount").textContent = clipgenPluralUnit(data.total_count, "match", "matches");
    header.classList.remove("hidden");

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
        html += '<span class="search-result-time">' + formatTime(r.start) + '</span>';
        html += '<span class="search-result-text">' + highlightQuery(r.text, state.searchQuery) + '</span>';
        html += '</div>';
      });
    });

    var prevScrollTop = list.scrollTop;
    list.innerHTML = html;
    list.scrollTop = prevScrollTop;
    container.classList.remove("hidden");

    // Attach click handlers
    var rows = list.querySelectorAll(".search-result-row[data-participant]");
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
    // Match on the RAW text and escape each piece separately: matching the
    // escaped text let a query like "amp" or "lt" land inside an entity
    // escapeHtml had just produced and split it with the span, rendering the
    // entity literally.
    var regex = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
    var out = "";
    var last = 0;
    var m;
    while ((m = regex.exec(text)) !== null) {
      out += escapeHtml(text.slice(last, m.index));
      out += '<span class="search-highlight">' + escapeHtml(m[0]) + "</span>";
      last = m.index + m[0].length;
    }
    out += escapeHtml(text.slice(last));
    return out;
  }

  function hideSearchResults() {
    qs("#searchResults").classList.add("hidden");
    qs("#searchResultsHeader").classList.add("hidden");
    qs("#searchCount").textContent = "";
    state.searchResults = null;
  }

  function jumpToResult(pid, start) {
    hideSearchResults();
    if (pid !== state.selectedParticipant) selectParticipant(pid);
    // seekVideo lives in the video satellite; late-bind through TS.
    if (TS.seekVideo) TS.seekVideo(start);
  }

  // ---- Published back to the hub (boot wires initSearch) ----
  TS.initSearch = initSearch;
})();
