/* Transcripts satellite: PII redaction.
 *
 * Owns the placeholder chip markup, the per-participant switch calls and the
 * Redact run/stop calls. Spans arrive numbered from the server with UTF-16
 * offsets; this file only maps them onto the whitespace tokens the karaoke
 * sweep times. Loads right after transcripts-speakers.js.
 */
(function () {
  "use strict";

  var TS = window.ClipgenTranscripts;
  var state = TS.state;
  var showToast = TS.showToast,
    loadTranscript = TS.loadTranscript,
    loadParticipants = TS.loadParticipants,
    pollTaskStatus = TS.pollTaskStatus,
    startPolling = TS.startPolling,
    _tokenRuns = TS._tokenRuns,
    segWordTiming = TS.segWordTiming,
    enabledFor = TS.enabledFor,
    latestTask = TS.latestTask,
    _selectedParticipantRow = TS._selectedParticipantRow;

  function redactOn() {
    return !!(state.redaction && state.redaction.enabled && state.redaction.detected);
  }

  function redactEnabledFor(p) {
    return enabledFor(p.redaction, CLIPGEN_CONFIG.transcribeRedact);
  }

  // "GIVEN_NAME" -> "Given name"; no table to keep in sync with labels.json.
  function labelTitle(label) {
    var words = String(label || "").toLowerCase().split("_");
    if (!words[0]) return "";
    words[0] = words[0].charAt(0).toUpperCase() + words[0].slice(1);
    return words.join(" ");
  }

  function _spanFor(spans, run) {
    if (!run) return null;
    for (var i = 0; i < spans.length; i++) {
      var sp = spans[i];
      if (sp.excluded) continue;
      if (Math.max(sp.start, run.start) < Math.min(sp.end, run.end)) return sp;
    }
    return null;
  }

  // Same word spans as the hub; a covered word becomes a chip.
  function redactedTextHtml(seg) {
    var spans = seg.pii || [];
    var runs = _tokenRuns(seg.text);
    var segWords = segWordTiming(seg, runs);
    var html = "";
    var wi = 0;
    var open = null; // {span, ws, we}
    function flush() {
      if (!open) return;
      var timing = open.ws !== null ? ' data-ws="' + open.ws + '" data-we="' + open.we + '"' : "";
      var tip = labelTitle(open.span.label) + " · " + Math.round((open.span.score || 0) * 100) + "%";
      html += '<span class="segment-word segment-pii"' + timing + ' data-tooltip="' + escapeHtml(tip) + '">' +
        escapeHtml(open.span.placeholder || "[PII]") + "</span>";
      open = null;
    }
    for (var r = 0; r < runs.length; r++) {
      var run = runs[r];
      if (run.space) {
        // Whitespace stays unless the same span continues on the next word.
        if (open && _spanFor(spans, runs[r + 1]) === open.span) continue;
        flush();
        html += run.text;
        continue;
      }
      var sp = _spanFor(spans, run);
      var ws = segWords ? segWords[wi].start : null;
      var we = segWords ? segWords[wi].end : null;
      wi++;
      if (!sp) {
        flush();
        html += '<span class="segment-word"' +
          (ws !== null ? ' data-ws="' + ws + '" data-we="' + we + '"' : "") + ">" +
          escapeHtml(run.text) + "</span>";
        continue;
      }
      if (open && open.span !== sp) flush();
      var head = run.text.slice(0, Math.max(0, sp.start - run.start));
      var tail = run.text.slice(Math.min(run.text.length, sp.end - run.start));
      if (head) { flush(); html += escapeHtml(head); }
      if (!open) open = { span: sp, ws: ws, we: we };
      else open.we = we !== null ? we : open.we;
      if (tail) { flush(); html += escapeHtml(tail); }
    }
    flush();
    return html;
  }

  function redactedPlainText(seg) {
    var spans = (seg.pii || []).slice().sort(function (a, b) { return b.start - a.start; });
    var out = seg.text;
    for (var i = 0; i < spans.length; i++) {
      if (spans[i].excluded) continue;
      out = out.slice(0, spans[i].start) + (spans[i].placeholder || "[PII]") + out.slice(spans[i].end);
    }
    return out;
  }

  // What any tooltip or quote may show: placeholders while redaction is applied.
  function displayText(seg) {
    if (!seg) return "";
    if (redactOn() && seg.pii && seg.pii.length) return redactedPlainText(seg);
    return seg.text || "";
  }

  function _adoptTask(task) {
    if (!task) return;
    state.tasks = state.tasks.filter(function (t) { return t.id !== task.id; }).concat([task]);
    startPolling();
    pollTaskStatus();
  }

  function _redactFailToast(data, fallback) {
    showToast(data.reason === "model_missing" ? "Download the Redact model in Settings" : (data.error || fallback));
  }

  function setRedactEnabled(pid, enabled) {
    return apiPut("api/redact/" + pid, { enabled: enabled }).then(function (data) {
      if (!data.ok) {
        _redactFailToast(data, "Failed to update redaction");
        return loadParticipants();
      }
      _adoptTask(data.task);
      return loadParticipants().then(function () {
        if (state.selectedParticipant === pid) loadTranscript(pid);
      });
    }).catch(function () {
      showToast("Failed to update redaction");
      return loadParticipants();
    });
  }

  function regenerateRedact(pid) {
    return apiPost("api/redact/" + pid + "/regenerate", {}).then(function (data) {
      if (!data.ok) {
        _redactFailToast(data, "Failed to start redaction");
        return;
      }
      _adoptTask(data.task);
      loadParticipants();
    }).catch(function () {
      showToast("Failed to start redaction");
    });
  }

  function stopRedact(pid) {
    return apiPost("api/redact/" + pid + "/stop", {}).then(function () {
      pollTaskStatus();
    }).catch(function () {
      showToast("Failed to stop redaction");
    });
  }

  // ---- Redact tab ----

  // One row per placeholder: label, surface, occurrences, weakest score, first segment.
  function _items() {
    var byKey = {};
    var order = [];
    for (var i = 0; i < state.segments.length; i++) {
      var seg = state.segments[i];
      var spans = seg.pii || [];
      for (var j = 0; j < spans.length; j++) {
        var sp = spans[j];
        var key = sp.placeholder || (sp.label + ":" + j);
        var it = byKey[key];
        if (!it) {
          it = { placeholder: sp.placeholder, label: sp.label, surface: seg.text.slice(sp.start, sp.end),
                 count: 0, score: 1, segIndex: i, excluded: !!sp.excluded };
          byKey[key] = it;
          order.push(it);
        }
        it.count++;
        if (typeof sp.score === "number" && sp.score < it.score) it.score = sp.score;
      }
    }
    return order;
  }

  function renderRedactPanel() {
    var tab = document.getElementById("redactTab");
    if (!tab) return;
    var row = _selectedParticipantRow();
    var task = latestTask(state.selectedParticipant, "redact");
    var live = !!task && (task.status === "running" || task.status === "queued");
    var enabled = row ? redactEnabledFor(row) : false;
    var modelOk = state.redactModel !== false;
    var rd = state.redaction || {};
    var items = enabled ? _items() : [];

    var sw = document.getElementById("redactSwitch");
    sw.checked = enabled;
    sw.disabled = !modelOk || !row;

    var status = "";
    if (live) status = task.status === "running"
      ? "Redacting\u2026 " + Math.round((task.progress || 0) * 100) + "%"
      : "Redaction queued";
    else if (task && task.status === "failed" && enabled) status = "Redaction failed" + (task.error ? " (" + task.error + ")" : "");
    else if (!enabled) status = "Off";
    else if (rd.error) status = rd.error;
    else if (!rd.detected) status = modelOk ? "Not run yet" : "Model not downloaded";
    else status = items.length + " item" + (items.length === 1 ? "" : "s") + ", " + (rd.count || 0) + " mention" + (rd.count === 1 ? "" : "s");
    var statusEl = document.getElementById("redactStatus");
    statusEl.textContent = status;
    statusEl.classList.toggle("friction-status--error", /failed/.test(status) || !!rd.error);
    statusEl.classList.toggle("cg-shimmer", live);
    // Like Summary, the empty state speaks for itself.
    statusEl.classList.toggle("hidden", !live && !(enabled && rd.detected) && modelOk);

    // Header controls appear once a pass ran; the empty-state CTA starts it.
    var ran = enabled && !!rd.detected;
    document.querySelector("#redactTab .redact-switch").classList.toggle("hidden", !ran && !live);
    var rerun = document.getElementById("redactRerun");
    rerun.classList.toggle("hidden", !ran || !modelOk || live);
    document.getElementById("redactCancel").classList.toggle("hidden", !live);

    document.getElementById("redactModel").classList.toggle("hidden", modelOk);
    var showEmpty = !live && (!ran || items.length === 0);
    var empty = document.getElementById("redactEmpty");
    empty.classList.toggle("hidden", !showEmpty);
    var cta = document.getElementById("redactRunCta");
    if (showEmpty) {
      document.getElementById("redactEmptyText").textContent = ran
        ? "No personal data found."
        : "No redaction yet.";
      cta.classList.toggle("hidden", ran || !modelOk);
      cta.disabled = !row;
    }

    var list = document.getElementById("redactList");
    list.classList.toggle("hidden", !(ran && items.length > 0));
    list.innerHTML = "";
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var el = document.createElement("div");
      el.className = "redact-item" + (it.excluded ? " redact-item--restored" : "");
      el.setAttribute("data-seg", String(it.segIndex));
      el.setAttribute("data-label", it.label);
      el.setAttribute("data-text", it.surface);
      el.setAttribute("data-excluded", it.excluded ? "1" : "0");
      el.innerHTML =
        '<span class="segment-word segment-pii">' + escapeHtml(it.placeholder || "[PII]") + "</span>" +
        '<span class="redact-item-label">' + escapeHtml(labelTitle(it.label)) + "</span>" +
        '<span class="redact-item-surface">' + escapeHtml(it.surface) + "</span>" +
        '<span class="redact-item-meta">' + it.count + "\u00D7 \u00B7 " + Math.round(it.score * 100) + "%</span>" +
        '<button type="button" class="redact-item-undo" data-tooltip="' +
          (it.excluded ? "Redact again" : "Restore the original text") + '" aria-label="' +
          (it.excluded ? "Redact again" : "Restore") + '"><span class="redact-item-undo-icon' +
          (it.excluded ? " redact-item-undo-icon--redo" : "") + '"></span></button>';
      list.appendChild(el);
    }
  }

  function setExcluded(pid, label, text, excluded) {
    return apiPut("api/redact/" + pid + "/exclude", { label: label, text: text, excluded: excluded })
      .then(function (data) {
        if (!data.ok) { showToast(data.error || "Failed to update redaction"); return; }
        if (state.selectedParticipant === pid) loadTranscript(pid);
      })
      .catch(function () { showToast("Failed to update redaction"); });
  }

  function _setDownloadUi(st) {
    var bar = document.getElementById("redactModelBar");
    var fill = document.getElementById("redactModelFill");
    var note = document.getElementById("redactModelNote");
    var btn = document.getElementById("redactDownload");
    if (st.done) {
      bar.classList.add("hidden");
      btn.disabled = false;
      note.textContent = st.succeeded ? "" : (st.error || "Download failed");
      if (st.succeeded) loadParticipants();
      return;
    }
    btn.disabled = true;
    bar.classList.remove("hidden");
    if (st.total > 0) {
      var pct = Math.max(0, Math.min(100, Math.round((st.completed / st.total) * 100)));
      fill.style.width = pct + "%";
      note.textContent = pct + "%";
    }
  }

  var _downloading = false;
  function _pollDownload() {
    if (_downloading) return;
    _downloading = true;
    pollDownloadStatus("/api/models/redact/download-status", _setDownloadUi, { label: "transcripts.redactDownload" })
      .then(function (st) {
        _downloading = false;
        if (!st) _setDownloadUi({ done: true, succeeded: false });
      });
  }

  function startModelDownload() {
    apiPost("/api/models/redact/download", {}).then(function (data) {
      if (!data || !data.ok) { _setDownloadUi({ done: true, succeeded: false, error: data && data.error }); return; }
      if (data.installed) { _setDownloadUi({ done: true, succeeded: true }); return; }
      _pollDownload();
    }).catch(function () { _setDownloadUi({ done: true, succeeded: false }); });
  }

  function initRedact() {
    var sw = document.getElementById("redactSwitch");
    if (!sw) return;
    sw.addEventListener("change", function () {
      sw.blur();
      if (!state.selectedParticipant) return;
      sw.disabled = true;
      setRedactEnabled(state.selectedParticipant, sw.checked);
    });
    document.getElementById("redactRerun").addEventListener("click", function () {
      if (state.selectedParticipant) regenerateRedact(state.selectedParticipant);
    });
    document.getElementById("redactCancel").addEventListener("click", function () {
      if (state.selectedParticipant) stopRedact(state.selectedParticipant);
    });
    document.getElementById("redactDownload").addEventListener("click", startModelDownload);
    document.getElementById("redactRunCta").addEventListener("click", function () {
      if (!state.selectedParticipant) return;
      document.getElementById("redactRunCta").disabled = true;
      setRedactEnabled(state.selectedParticipant, true);
    });
    document.getElementById("redactList").addEventListener("click", function (e) {
      var item = e.target.closest(".redact-item");
      if (!item) return;
      var undo = e.target.closest(".redact-item-undo");
      if (undo && state.selectedParticipant) {
        undo.disabled = true;
        setExcluded(state.selectedParticipant, item.getAttribute("data-label"),
          item.getAttribute("data-text"), item.getAttribute("data-excluded") !== "1");
        return;
      }
      var seg = state.segments[parseInt(item.getAttribute("data-seg"), 10)];
      var row = seg && document.querySelector('#segmentList .segment-row[data-index="' + item.getAttribute("data-seg") + '"]');
      if (row && TS.scrollToSegment) TS.scrollToSegment(row);
      if (seg && TS.seekVideo) TS.seekVideo(seg.start);
    });
    // Size and license come from the models catalog; a download in flight joins its poll.
    apiGet("/api/models").then(function (data) {
      var rd = data && data.redact;
      if (!rd) return;
      document.getElementById("redactModelSize").textContent = "(" + formatModelSize(rd.size_mb || 0) + ")";
      var lic = document.getElementById("redactLicenseLink");
      lic.textContent = rd.license || "model license";
      lic.href = rd.license_url || "#";
      // Set here, not in the HTML: live pages may reference no remote origin.
      if (rd.vendor_url) document.getElementById("redactVendorLink").href = rd.vendor_url;
    }).catch(function () {});
    apiGet("/api/models/redact/download-status").then(function (st) {
      if (st && st.ok && st.found && !st.done) { _setDownloadUi(st); _pollDownload(); }
    }).catch(function () {});
  }

  TS.redactOn = redactOn;
  TS.renderRedactPanel = renderRedactPanel;
  TS.initRedact = initRedact;
  TS.redactedTextHtml = redactedTextHtml;
  TS.redactedPlainText = redactedPlainText;
  TS.displayText = displayText;
})();
