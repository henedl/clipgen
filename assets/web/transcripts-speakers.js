/* Transcripts satellite: speaker attribution.
 *
 * Owns the speaker chip markup, the per-participant switch calls, the
 * "Detect speakers" run/stop calls, and the rename popover. Shares state with
 * the hub through window.ClipgenTranscripts (TS); loads right after
 * transcripts.js so search, video and pills can import from it.
 */
(function () {
  "use strict";

  var TS = window.ClipgenTranscripts;
  var state = TS.state;
  var showToast = TS.showToast,
    loadTranscript = TS.loadTranscript,
    loadParticipants = TS.loadParticipants,
    renderSegments = TS.renderSegments,
    pollTaskStatus = TS.pollTaskStatus,
    startPolling = TS.startPolling;

  // Matches --region-color-1..8 in tokens.css; ids past 8 wrap.
  var SPEAKER_PALETTE_SIZE = 8;

  function speakersOn() {
    return !!(state.speakers && state.speakers.enabled);
  }

  function speakerName(id) {
    var labels = state.speakers && state.speakers.labels;
    return (labels && labels[id]) || "Speaker " + id;
  }

  function speakerClass(id) {
    var n = parseInt(id, 10);
    if (isNaN(n) || n < 1) n = 1;
    return "spk-" + (((n - 1) % SPEAKER_PALETTE_SIZE) + 1);
  }

  // Chip markup. opts.repeat dims a same-speaker run; opts.inert drops the rename hook.
  function speakerChipHtml(id, opts) {
    opts = opts || {};
    if (!id) return '<span class="segment-speaker segment-speaker--none"></span>';
    var cls = "segment-speaker " + speakerClass(id);
    if (opts.repeat) cls += " segment-speaker--repeat";
    var name = opts.name || speakerName(id);
    var attrs = "";
    if (!opts.inert) {
      attrs = ' data-speaker="' + escapeHtml(id) + '" data-tooltip="Rename speaker"';
    }
    return '<span class="' + cls + '"' + attrs + ">" + escapeHtml(name) + "</span>";
  }

  // Null = no per-participant choice yet; follow the global default.
  function speakersEnabledFor(p) {
    var sp = p.speakers || {};
    if (sp.enabled === null || sp.enabled === undefined) {
      return !!CLIPGEN_CONFIG.transcribeSpeakers;
    }
    return !!sp.enabled;
  }

  function _refreshAfterSpeakerChange(pid) {
    return loadParticipants().then(function () {
      if (state.selectedParticipant === pid) loadTranscript(pid);
      startPolling();
      pollTaskStatus();
    });
  }

  function setSpeakersEnabled(pid, enabled) {
    return apiPut("api/speakers/" + pid, { enabled: enabled })
      .then(function (data) {
        if (!data.ok) throw new Error(data.error || "speakers");
        return _refreshAfterSpeakerChange(pid);
      })
      .catch(function (err) {
        showToast(err && err.message ? err.message : "Failed to update speakers");
        if (TS.renderPills) TS.renderPills();
      });
  }

  function regenerateSpeakers(pid) {
    return apiPost("api/speakers/" + pid + "/regenerate", {})
      .then(function (data) {
        if (!data.ok) throw new Error(data.error || "speakers");
        return _refreshAfterSpeakerChange(pid);
      })
      .catch(function (err) {
        showToast(err && err.message ? err.message : "Failed to start speakers");
        if (TS.renderPills) TS.renderPills();
      });
  }

  function stopSpeakers(pid) {
    return apiPost("api/speakers/" + pid + "/stop", {})
      .then(function () {
        pollTaskStatus();
      })
      .catch(function () {
        showToast("Failed to stop speakers");
      });
  }

  // ---- Rename popover ----

  var _spkOpen = null;          // { pid, id, segIndex, ver } while the popover is up
  var _spkCancelled = false;    // Escape sets it so the trailing blur does not save
  var _spkAttachTimer = null;

  function _spkOutsideClick(e) {
    var pop = qs("#speakerPopover");
    if (pop && pop.contains(e.target)) return;
    _spkCancelled = true;
    hideSpeakerPopover();
  }

  // Chips for every known speaker plus one "new" slot; the line's own is marked.
  function _buildAssignChips(container, currentId) {
    container.innerHTML = "";
    var count = Math.max(state.speakers && state.speakers.count || 0, parseInt(currentId, 10) || 0);
    for (var i = 1; i <= count + 1; i++) {
      var id = String(i);
      var chip = document.createElement("button");
      chip.type = "button";
      var isNew = i === count + 1;
      chip.className = "segment-speaker speaker-popover-chip " +
        (isNew ? "speaker-popover-chip--new" : speakerClass(id)) +
        (id === currentId ? " active" : "");
      chip.textContent = isNew ? "+ New" : speakerName(id);
      chip.setAttribute("data-assign", id);
      chip.setAttribute("data-tooltip", isNew ? "Assign this line to a new speaker" : "Assign this line to " + speakerName(id));
      container.appendChild(chip);
    }
  }

  function _assignThisLine(speaker) {
    var open = _spkOpen;
    _spkCancelled = true;
    hideSpeakerPopover();
    if (!open || open.segIndex === undefined) return;
    var seg = state.segments[open.segIndex];
    if (!seg || !seg.id || seg.speaker === speaker) return;
    apiPut("api/speakers/" + open.pid + "/segment", { segment_id: seg.id, speaker: speaker })
      .then(function (data) {
        if (!data.ok || open.ver !== state.participantReqVer) return;
        seg.speaker = data.speaker;
        if (state.speakers && data.speakers) state.speakers.count = data.speakers.count;
        renderSegments();
      })
      .catch(function () {
        showToast("Failed to reassign line");
      });
  }

  function showSpeakerPopover(anchorEl, id, segIndex) {
    if (TS.hideMarkPopover) TS.hideMarkPopover();
    hideSpeakerPopover();
    var pop = qs("#speakerPopover");
    if (!pop) return;
    var input = pop.querySelector(".speaker-popover-input");
    var labels = (state.speakers && state.speakers.labels) || {};
    input.value = labels[id] || "";
    input.placeholder = "Speaker " + id;
    _spkOpen = { pid: state.selectedParticipant, id: id, segIndex: segIndex, ver: state.participantReqVer };
    _spkCancelled = false;
    _buildAssignChips(pop.querySelector(".speaker-popover-assign"), id);
    // Buttons keep focus on the input so their click, not a blur, decides.
    pop.onmousedown = function (e) {
      if (e.target !== input) e.preventDefault();
    };
    pop.onclick = function (e) {
      e.stopPropagation();
      var chip = e.target.closest(".speaker-popover-chip");
      if (chip) { _assignThisLine(chip.getAttribute("data-assign")); return; }
      if (e.target.closest(".speaker-popover-reset")) {
        input.value = "";
        _commitSpeakerLabel("");
      }
    };
    input.onkeydown = function (e) {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      if (e.key === "Escape") { e.preventDefault(); _spkCancelled = true; hideSpeakerPopover(); }
      e.stopPropagation();
    };
    input.onblur = function () {
      if (!_spkCancelled && _spkOpen) _commitSpeakerLabel(input.value);
    };
    pop.classList.remove("hidden");
    positionPopoverAnchored(pop, anchorEl.getBoundingClientRect());
    input.focus();
    input.select();
    // Deferred outside-click, same shape as the mark popover.
    if (_spkAttachTimer) clearTimeout(_spkAttachTimer);
    _spkAttachTimer = setTimeout(function () {
      _spkAttachTimer = null;
      document.addEventListener("click", _spkOutsideClick);
    }, 0);
  }

  function hideSpeakerPopover() {
    var pop = qs("#speakerPopover");
    if (pop) pop.classList.add("hidden");
    if (_spkAttachTimer) { clearTimeout(_spkAttachTimer); _spkAttachTimer = null; }
    document.removeEventListener("click", _spkOutsideClick);
    _spkOpen = null;
  }

  function _commitSpeakerLabel(raw) {
    var open = _spkOpen;
    hideSpeakerPopover();
    if (!open) return;
    var value = raw.trim();
    var labels = (state.speakers && state.speakers.labels) || {};
    var current = labels[open.id] || "";
    if (value === current) return;
    var body = {};
    body[open.id] = value; // "" resets server-side
    apiPut("api/speakers/" + open.pid + "/labels", body)
      .then(function (data) {
        if (!data.ok || open.ver !== state.participantReqVer || !state.speakers) return;
        if (!state.speakers.labels) state.speakers.labels = {};
        if (value) state.speakers.labels[open.id] = value;
        else delete state.speakers.labels[open.id];
        renderSegments();
      })
      .catch(function () {
        showToast("Failed to rename speaker");
      });
  }

  function initSpeakers() {
    if (window.ClipgenHotkeys && window.ClipgenHotkeys.registerEscape) {
      window.ClipgenHotkeys.registerEscape(function () {
        if (!_spkOpen) return false;
        _spkCancelled = true;
        hideSpeakerPopover();
        return true;
      });
    }
  }

  TS.speakersOn = speakersOn;
  TS.speakerName = speakerName;
  TS.speakerClass = speakerClass;
  TS.speakerChipHtml = speakerChipHtml;
  TS.speakersEnabledFor = speakersEnabledFor;
  TS.setSpeakersEnabled = setSpeakersEnabled;
  TS.regenerateSpeakers = regenerateSpeakers;
  TS.stopSpeakers = stopSpeakers;
  TS.showSpeakerPopover = showSpeakerPopover;
  TS.hideSpeakerPopover = hideSpeakerPopover;
  TS.initSpeakers = initSpeakers;
})();
