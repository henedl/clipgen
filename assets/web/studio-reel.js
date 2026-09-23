/* clipgen Studio reel satellite — studio-reel.js
 *
 * The reel and standalone-viewer builds carved out of studio.js: Build Reel
 * (streaming /api/reel and /api/reel-direct), Find Highlights, the timeline
 * viewer and HTML viewer builds, and the gallery dialog. Loads last and
 * destructures the hub's status / progress helpers at load; the hub keeps
 * same-named delegators for its button wiring. Function bodies are unchanged
 * from the hub.
 */
(function () {
  "use strict";

  var STUDIO = window.ClipgenStudio;
  var state = STUDIO.state;
  var _paintReelElapsed = STUDIO._paintReelElapsed,
    _reelEtaTracker = STUDIO._reelEtaTracker,
    _studioEtaTicker = STUDIO._studioEtaTicker,
    buildCellOverrides = STUDIO.buildCellOverrides,
    cellKey = STUDIO.cellKey,
    clearCardStatus = STUDIO.clearCardStatus,
    expandCellToSegments = STUDIO.expandCellToSegments,
    hideBuildStatus = STUDIO.hideBuildStatus,
    iconHTML = STUDIO.iconHTML,
    isAnyStudioJobRunning = STUDIO.isAnyStudioJobRunning,
    isIntakeSource = STUDIO.isIntakeSource,
    pathBasename = STUDIO.pathBasename,
    renderReelQueue = STUDIO.renderReelQueue,
    revealStatusOverlay = STUDIO.revealStatusOverlay,
    setButtonProgress = STUDIO.setButtonProgress,
    setCardQueued = STUDIO.setCardQueued,
    setCardResult = STUDIO.setCardResult,
    setReelGenerating = STUDIO.setReelGenerating,
    showBuildResult = STUDIO.showBuildResult,
    showBuildStatus = STUDIO.showBuildStatus,
    showConfirm = STUDIO.showConfirm,
    showOverlay = STUDIO.showOverlay,
    showResult = STUDIO.showResult,
    stampLog = STUDIO.stampLog,
    updateSingleCellClass = STUDIO.updateSingleCellClass;


  function onBuildReel() {
    if (state.reelGenerating || state.reelQueue.length === 0) return;
    setReelGenerating(true);
    qs("#cancelReelBtn").classList.remove("hidden");
    _reelEtaTracker.reset();
    _reelEtaTracker.start();
    _studioEtaTicker.ensure();
    _paintReelElapsed();

    // Determine if we have intake items — use direct endpoint for mixed/intake reels
    var hasIntake = false;
    for (var ci = 0; ci < state.reelQueue.length; ci++) {
      if (isIntakeSource(state.reelQueue[ci].source)) { hasIntake = true; break; }
    }

    var reelBody;
    var endpoint;

    if (hasIntake) {
      var segments = [];
      for (var si = 0; si < state.reelQueue.length; si++) {
        var item = state.reelQueue[si];
        segments.push({
          participant: item.participant,
          start: item.start,
          end: item.end,
          source: item.source || "screenspace",
        });
      }
      reelBody = { segments: segments };
      endpoint = "api/reel-direct";
    } else {
      var cellsSeen = {};
      var cells = [];
      for (var ci2 = 0; ci2 < state.reelQueue.length; ci2++) {
        var ck = state.reelQueue[ci2].participant + "." + state.reelQueue[ci2].row;
        if (!cellsSeen[ck]) { cellsSeen[ck] = true; cells.push(ck); }
      }
      reelBody = { cells: cells };
      var reelOverrides = buildCellOverrides(state.reelQueue);
      if (Object.keys(reelOverrides).length > 0) reelBody.overrides = reelOverrides;
      endpoint = "api/reel";
    }

    var tcCb = qs("#titlecardEnabled");
    var tcDur = qs("#titlecardDuration");
    if (tcCb) reelBody.titlecards_enabled = tcCb.checked;
    if (tcDur) reelBody.titlecard_duration = parseInt(tcDur.value, 10) || 2;

    var list = qs("#reelList");
    var reelCards = list.querySelectorAll(".queue-card");
    for (var i = 0; i < reelCards.length; i++) {
      setCardQueued(reelCards[i]);
    }

    // Two phases: clips 0.7, concat 0.3; clips finish first, so the bar stays monotonic.
    var totalClips = 0;
    var clipsDone = 0;
    var concatFraction = 0;
    var finalPayload = null;
    var cancelled = false;
    var finished = false;

    function updateProgress() {
      var clipFraction = totalClips > 0 ? Math.min(clipsDone / totalClips, 1) : 0;
      var overall = clipFraction * 0.7 + concatFraction * 0.3;
      setButtonProgress("buildReelBtn", overall);
    }

    function finish() {
      if (finished) return;
      finished = true;
      setReelGenerating(false);
      qs("#cancelReelBtn").classList.add("hidden");
      setButtonProgress("buildReelBtn", null);
      _reelEtaTracker.reset();
      _paintReelElapsed();

      var data = finalPayload || {};
      var isCancelled = cancelled || !!data.cancelled;

      var cards = list.querySelectorAll(".queue-card");
      for (var j = 0; j < cards.length; j++) {
        if (isCancelled) {
          clearCardStatus(cards[j]);
        } else {
          setCardResult(cards[j], !!data.ok);
        }
      }

      if (isCancelled) {
        showResult(null, "Reel generation cancelled");
      } else if (data.ok) {
        showResult("Reel built successfully", null);
      } else {
        showResult(null, data.error || "Reel build failed");
      }
      revealStatusOverlay();
    }

    function handleLine(line) {
      var data;
      try { data = JSON.parse(line); } catch (e) { return; }
      if (!data) return;
      if (data.cancelled) cancelled = true;
      if (data.phase === "start") {
        totalClips = data.total_clips || 0;
        updateProgress();
      } else if (data.phase === "clip_done") {
        clipsDone += 1;
        updateProgress();
      } else if (data.phase === "concat") {
        concatFraction = typeof data.progress === "number" ? data.progress : 0;
        updateProgress();
      } else if (data.phase === "done") {
        // Stream-copy concat emits no progress; fill to 100% before the final line.
        concatFraction = 1;
        updateProgress();
      } else if (data.ok !== undefined || data.error !== undefined) {
        finalPayload = data;
        if (data.ok && Array.isArray(data.reels)) {
          for (var ri = 0; ri < data.reels.length; ri++) {
            state.generatedReels.push(stampLog(data.reels[ri]));
          }
        }
      }
    }

    apiPostNDJSON(endpoint, reelBody, { onLine: handleLine })
      .then(finish)
      .catch(function (err) {
        // 4xx/5xx (e.g. 409 in-progress) arrive as JSON in err.bodyText; parse into the payload.
        if (err && err.status >= 400) {
          try {
            finalPayload = JSON.parse(err.bodyText);
          } catch (_) {
            finalPayload = {
              ok: false,
              error: err.bodyText || ("HTTP " + err.status),
            };
          }
        } else {
          finalPayload = { ok: false, error: "Request failed: " + err };
        }
        finish();
      });
  }

  function onBuildViewer() {
    if (isAnyStudioJobRunning() || state.generatedArtifacts.length === 0) return;
    state.overlayJobRunning = true;

    showBuildStatus("Building timeline viewer…", null);

    apiPost("api/viewer", {})
      .then(function (data) {
        state.overlayJobRunning = false;
        if (data.ok) {
          state.generatedViewers.push(stampLog({
            type: "viewer",
            subtype: "viewer",
            file: pathBasename(data.file),
            description: "Timeline viewer",
          }));
          showBuildResult("Viewer created: " + (data.file || ""), null, data.file);
        } else {
          showBuildResult(null, data.error || "Viewer build failed");
        }
      })
      .catch(function (err) {
        state.overlayJobRunning = false;
        showBuildResult(null, "Request failed: " + err);
      });
  }

  function onBuildTimelineViewer() {
    if (isAnyStudioJobRunning()) return;

    var ssCount = (state.intakeClusters || []).length;
    var trCount = (state.trIntakeClusters || []).length;
    if (ssCount === 0 && trCount === 0) {
      startTimelineViewerBuild(false);
      return;
    }

    var parts = [];
    if (ssCount > 0) {
      parts.push(ssCount + " Screenspace event group" + (ssCount === 1 ? "" : "s"));
    }
    if (trCount > 0) {
      parts.push(trCount + " Transcript mark group" + (trCount === 1 ? "" : "s"));
    }
    var msg = parts.join(" and ") + " detected. Include them as clips in the timeline viewer?";

    showConfirm(
      "Include Intake Events?",
      msg,
      function () { startTimelineViewerBuild(true); },
      function () { startTimelineViewerBuild(false); }
    );
  }

  function startTimelineViewerBuild(includeIntake) {
    state.overlayJobRunning = true;
    state.timelineViewerCancelledByUser = false;
    var body = {};

    var ssClusters = state.intakeClusters || [];
    var trClusters = state.trIntakeClusters || [];
    var hasIntake = includeIntake && (ssClusters.length > 0 || trClusters.length > 0);

    if (hasIntake) {
      showBuildStatus(
        "Building timeline viewer with intake events\u2026",
        onCancelTimelineViewer
      );
      body.include_intake = true;
      var items = ssClusters.map(function (c) {
        return {
          participant: c.participant,
          start: c.start,
          end: c.end,
          event_type: c.event_type,
          event_ids: c.events.map(function (e) { return e.id; }),
        };
      });
      for (var i = 0; i < trClusters.length; i++) {
        var c = trClusters[i];
        items.push({
          participant: c.participant,
          start: c.start,
          end: c.end,
          event_type: c.category || "transcript",
          source: "transcript",
          mark_ids: c.marks.map(function (m) { return m.id; }),
          text: c.text || "",
          label: c.label || "",
        });
      }
      body.intake_items = items;
    } else {
      showBuildStatus("Building timeline viewer\u2026", onCancelTimelineViewer);
    }

    apiPost("api/timeline-viewer", body)
      .then(function (data) {
        state.overlayJobRunning = false;
        if (data.cancelled || state.timelineViewerCancelledByUser) {
          state.timelineViewerCancelledByUser = false;
          hideBuildStatus();
          showToast("Build cancelled");
          return;
        }
        if (data.ok) {
          state.generatedViewers.push(stampLog({
            type: "viewer",
            subtype: "timeline-viewer",
            file: pathBasename(data.file),
            description: "Timeline viewer (full sheet)",
          }));
          var msg = "Timeline viewer created: " + (data.file || "");
          if (data.generated) {
            msg = "Generated " + clipgenPluralUnit(data.generated, "clip", "clips") + ". " + msg;
          }
          showBuildResult(msg, null, data.file);
        } else {
          showBuildResult(null, data.error || "Timeline viewer build failed");
        }
      })
      .catch(function (err) {
        state.overlayJobRunning = false;
        if (state.timelineViewerCancelledByUser) {
          state.timelineViewerCancelledByUser = false;
          hideBuildStatus();
          showToast("Build cancelled");
          return;
        }
        showBuildResult(null, "Request failed: " + err);
      });
  }

  function onCancelTimelineViewer() {
    state.timelineViewerCancelledByUser = true;
    apiPost("api/timeline-viewer/cancel").catch(toastError("Cancel failed"));
  }

  var _highlightsBtnOrigHTML = "";

  // Close the highlights drawer without running (Escape); returns whether one was open.
  function cancelHighlightsDrawer() {
    var drawer = qs("#highlightsDurationDrawer");
    if (!drawer || !drawer.classList.contains("open")) return false;
    if (document.activeElement && drawer.contains(document.activeElement)) {
      document.activeElement.blur();
    }
    drawer.classList.remove("open");
    var btn = qs("#buildHighlightsBtn");
    if (btn) {
      btn.style.minWidth = "";
      if (_highlightsBtnOrigHTML) btn.innerHTML = _highlightsBtnOrigHTML;
    }
    return true;
  }

  function onBuildHighlights() {
    if (isAnyStudioJobRunning()) return;

    var drawer = qs("#highlightsDurationDrawer");
    var btn = qs("#buildHighlightsBtn");
    var isOpen = drawer.classList.contains("open");

    var checkHTML = iconHTML("check", "cg-icon--confirm");

    if (!isOpen) {
      _highlightsBtnOrigHTML = btn.innerHTML;
      drawer.classList.add("open");
      var w = btn.offsetWidth;
      btn.style.minWidth = w + "px";
      btn.innerHTML = checkHTML + "Confirm";
      return;
    }

    var duration = parseInt(qs("#highlightsDuration").value, 10);
    if (!Number.isFinite(duration) || duration < 1) duration = 180;

    drawer.classList.remove("open");
    btn.style.minWidth = "";
    btn.innerHTML = _highlightsBtnOrigHTML;

    setReelGenerating(true);
    showOverlay("Finding best clips (" + duration + "s budget)...");

    apiPost("api/highlights-preview", { highlights_duration: duration })
      .then(function (data) {
        setReelGenerating(false);
        if (data.ok && data.clips && data.clips.length > 0) {
          var prev = state.reelQueue.slice();
          state.reelQueue = [];
          for (var i = 0; i < data.clips.length; i++) {
            var entries = expandCellToSegments(data.clips[i]);
            for (var ei = 0; ei < entries.length; ei++) {
              state.reelQueue.push(entries[ei]);
            }
          }
          renderReelQueue();
          var touchedKeys = {};
          for (var p = 0; p < prev.length; p++) {
            if (prev[p].row) touchedKeys[cellKey(prev[p].participant, prev[p].row)] = prev[p];
          }
          for (var q = 0; q < state.reelQueue.length; q++) {
            var rq = state.reelQueue[q];
            if (rq.row) touchedKeys[cellKey(rq.participant, rq.row)] = rq;
          }
          for (var key in touchedKeys) {
            updateSingleCellClass(touchedKeys[key].participant, touchedKeys[key].row);
          }
          showResult(
            "Added " + clipgenPluralUnit(data.clips.length, "clip", "clips") + " to reel queue",
            null
          );
        } else {
          showResult(
            null,
            data.error || "No clips found for highlights selection"
          );
        }
      })
      .catch(function (err) {
        setReelGenerating(false);
        showResult(null, "Request failed: " + err);
      });
  }

  function populateGalleryParticipants(participants) {
    var sel = qs("#galleryParticipant");
    if (!sel) return;
    sel.innerHTML = "";
    for (var i = 0; i < participants.length; i++) {
      var opt = el("option");
      opt.value = participants[i];
      opt.textContent = participants[i];
      sel.appendChild(opt);
    }
  }

  function openGalleryDialog() {
    if (isAnyStudioJobRunning()) return;
    var overlay = qs("#galleryOverlay");
    if (!overlay) return;
    openPopModal(overlay, qs(".gallery-card"), { onEscape: closeGalleryDialog });
    var sel = qs("#galleryParticipant");
    if (sel) sel.focus();
  }

  function closeGalleryDialog() {
    var overlay = qs("#galleryOverlay");
    if (!overlay) return;
    // Trap released now; the visual hide trails the fade.
    closePopModal(overlay, qs(".gallery-card"), { releaseTrapNow: true });
  }

  function submitGalleryDialog() {
    if (isAnyStudioJobRunning()) return;

    var participant = qs("#galleryParticipant").value;
    var format = qs("#galleryFormat").value;
    var interval = parseInt(qs("#galleryInterval").value, 10);
    if (!interval || interval < 1) interval = 10;
    var bundle = qs("#galleryBundle").checked;

    if (!participant) {
      showToast("No participant selected for gallery");
      return;
    }

    closeGalleryDialog();
    state.overlayJobRunning = true;
    state.galleryCancelledByUser = false;
    showBuildStatus(
      "Generating gallery viewer for " + participant + "…",
      onCancelGallery
    );

    apiPost("api/gallery", { participant: participant, format: format, interval: interval, bundle: bundle })
      .then(function (data) {
        state.overlayJobRunning = false;
        if (data.cancelled || state.galleryCancelledByUser) {
          state.galleryCancelledByUser = false;
          hideBuildStatus();
          showToast("Build cancelled");
          return;
        }
        if (data.ok) {
          state.generatedViewers.push(stampLog({
            type: "viewer",
            subtype: "gallery",
            file: pathBasename(data.file),
            participant: participant,
            description: "Gallery viewer (" + format + ", " + interval + "s)",
          }));
          showBuildResult("Gallery viewer created: " + (data.file || ""), null, data.file);
        } else {
          showBuildResult(null, data.error || "Gallery build failed");
        }
      })
      .catch(function (err) {
        state.overlayJobRunning = false;
        if (state.galleryCancelledByUser) {
          state.galleryCancelledByUser = false;
          hideBuildStatus();
          showToast("Build cancelled");
          return;
        }
        showBuildResult(null, "Request failed: " + err);
      });
  }

  function onCancelGallery() {
    state.galleryCancelledByUser = true;
    apiPost("api/gallery/cancel").catch(toastError("Cancel failed"));
  }

  function bindGalleryDialog() {
    var overlay = qs("#galleryOverlay");
    var cancel = qs("#galleryDialogCancel");
    var confirm = qs("#galleryDialogConfirm");
    if (overlay) {
      overlay.addEventListener("click", function (ev) {
        if (ev.target === overlay) closeGalleryDialog();
      });
    }
    if (cancel) cancel.addEventListener("click", closeGalleryDialog);
    if (confirm) confirm.addEventListener("click", submitGalleryDialog);
    // Escape is handled by the modal focus trap opened in openGalleryDialog.
  }

  STUDIO.bindGalleryDialog = bindGalleryDialog;
  STUDIO.cancelHighlightsDrawer = cancelHighlightsDrawer;
  STUDIO.onBuildHighlights = onBuildHighlights;
  STUDIO.onBuildReel = onBuildReel;
  STUDIO.onBuildTimelineViewer = onBuildTimelineViewer;
  STUDIO.onBuildViewer = onBuildViewer;
  STUDIO.openGalleryDialog = openGalleryDialog;
  STUDIO.populateGalleryParticipants = populateGalleryParticipants;
})();
