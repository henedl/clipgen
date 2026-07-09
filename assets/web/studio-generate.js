/* clipgen Studio — artifact-generation satellite.
 *
 * Carved out of studio.js (the hub) following the hub+satellite convention
 * (see studio-scrubber/trim/intake). Owns the streaming api/generate +
 * api/generate-intake flow: the Generate button handler (onGenerate), its cancel
 * handler (onCancelGenerate), and the per-cell card index helper.
 *
 * Reaches hub-owned pieces through window.ClipgenStudio (STUDIO): the card-state
 * painters (setCardQueued/clearCardStatus/setCardResult) and readNDJSONStream are
 * shared with the reel/build path and stay in the hub; setArtifactGenerating /
 * showResult / revealStatusOverlay / updateGenerateProgress / stampLog likewise;
 * the _generateEtaTracker + _studioEtaTicker objects are shared elapsed-time
 * infrastructure (also driven by job-status polling and the reel/build flows);
 * buildCellOverrides lives in studio-trim.js (reached via STUDIO). isIntakeSource
 * is a hub helper. apiPost / qs / setButtonProgress / clipgenPluralUnit are
 * ambient utils.js / primitives.js globals (scope chain).
 *
 * The hub keeps same-named onGenerate/onCancelGenerate delegators for the button
 * wiring. Loaded by studio.html after studio.js and studio-trim.js (whose
 * STUDIO.buildCellOverrides this file uses), before studio-intake.js.
 */

(function () {
  "use strict";

  var STUDIO = window.ClipgenStudio;
  var state = STUDIO.state;
  // setButtonProgress is a ClipgenPrimitives namespace fn (primitives.js global),
  // aliased here the same way the hub aliases it.
  var setButtonProgress = ClipgenPrimitives.setButtonProgress;
  // Hub-owned helpers + shared elapsed-time trackers, published during the hub's
  // load (and buildCellOverrides during studio-trim.js's), before this file runs.
  var setArtifactGenerating = STUDIO.setArtifactGenerating,
    showResult = STUDIO.showResult,
    revealStatusOverlay = STUDIO.revealStatusOverlay,
    readNDJSONStream = STUDIO.readNDJSONStream,
    setCardQueued = STUDIO.setCardQueued,
    clearCardStatus = STUDIO.clearCardStatus,
    setCardResult = STUDIO.setCardResult,
    updateGenerateProgress = STUDIO.updateGenerateProgress,
    stampLog = STUDIO.stampLog,
    isIntakeSource = STUDIO.isIntakeSource,
    buildCellOverrides = STUDIO.buildCellOverrides,
    _generateEtaTracker = STUDIO._generateEtaTracker,
    _studioEtaTicker = STUDIO._studioEtaTicker;

  function buildGenerateCardIndex(listEl) {
    var map = {};
    var cards = listEl.querySelectorAll(".queue-card");
    for (var i = 0; i < cards.length; i++) {
      var card = cards[i];
      var participant = card.getAttribute("data-participant");
      var row = card.getAttribute("data-row");
      if (!participant || row == null) continue;
      var key = participant + "." + row;
      if (!map[key]) map[key] = [];
      map[key].push(card);
    }
    return map;
  }

  function isGenerateFetchAborted(err) {
    if (state.generateCancelledByUser) return true;
    return !!(err && err.name === "AbortError");
  }

  function onGenerate() {
    if (state.artifactGenerating || state.artifactQueue.length === 0) return;
    state.generateCancelledByUser = false;
    setArtifactGenerating(true);
    qs("#cancelGenerateBtn").classList.remove("hidden");
    _generateEtaTracker.reset();
    _generateEtaTracker.start();
    _studioEtaTicker.ensure();

    // Per-branch AbortControllers let onCancelGenerate stop the network
    // fetches immediately; the server-side cancel endpoints also trip the
    // cancel events so in-flight ffmpeg subprocesses get terminated.
    var sheetAbort = new AbortController();
    var intakeAbort = new AbortController();
    state.activeGenerateAborts = [sheetAbort, intakeAbort];

    var format = qs("#artifactFormat").value;
    var list = qs("#artifactsList");
    var items = state.artifactQueue.slice();

    // Capture the queue cards before any async work so per-item result
    // markers don't drift onto the wrong card if the queue re-renders mid
    // request. allCards is in DOM order, which matches state.artifactQueue.
    var allCards = list.querySelectorAll(".queue-card");
    for (var i = 0; i < allCards.length; i++) {
      setCardQueued(allCards[i]);
    }

    // Separate spreadsheet and intake items, keeping each split's card
    // element parallel to its item array so the resolve handler can match
    // by index against the captured card list (immune to later re-renders).
    var sheetItems = [];
    var sheetCardEls = [];
    var intakeItems = [];
    var intakeCardEls = [];
    for (var ci = 0; ci < items.length; ci++) {
      if (isIntakeSource(items[ci].source)) {
        intakeItems.push(items[ci]);
        intakeCardEls.push(allCards[ci]);
      } else {
        sheetItems.push(items[ci]);
        sheetCardEls.push(allCards[ci]);
      }
    }

    var totalSuccess = 0;
    var totalFail = 0;
    var allArtifacts = [];
    var failReasons = [];
    var cancelled = false;
    var pending = (sheetItems.length > 0 ? 1 : 0) + (intakeItems.length > 0 ? 1 : 0);
    var sheetCellTotal = 0;
    var sheetCellsDone = 0;
    var intakeDone = 0;
    var intakeTotal = intakeItems.length;
    var generateCardIndex = null;

    function updateGenerateButtonProgress() {
      var total = sheetCellTotal + intakeTotal;
      if (total <= 0) return;
      setButtonProgress("generateBtn", (sheetCellsDone + intakeDone) / total);
    }

    function finishBranch() {
      if (--pending > 0) return;
      setButtonProgress("generateBtn", null);
      setArtifactGenerating(false);
      _generateEtaTracker.reset();
      // Hide after artifactGenerating is false so the elapsed-only fallback in
      // _paintGenerateProgress doesn't keep the readout visible.
      updateGenerateProgress(0, 0);
      qs("#cancelGenerateBtn").classList.add("hidden");
      var msg;
      var err = null;
      if (cancelled) {
        msg = totalSuccess > 0
          ? "Cancelled after " + clipgenPluralUnit(totalSuccess, "artifact", "artifacts")
          : null;
        err = totalSuccess > 0 ? null : "Generation cancelled";
      } else if (totalSuccess === 0 && totalFail === 0) {
        // Stream ended without any per-item results — treat as an error
        // rather than silently reporting "Generated 0 artifacts".
        msg = null;
        err = "No artifacts were generated";
      } else {
        msg = "Generated " + clipgenPluralUnit(totalSuccess, "artifact", "artifacts");
        if (totalFail > 0) msg += ", " + totalFail + " failed";
        if (totalSuccess === 0 && totalFail > 0) {
          msg = null;
          err = "All generations failed";
        }
      }
      // Append up to 3 distinct failure reasons so the user can act on them
      // instead of just seeing a count (full reason is also on each card title).
      if (totalFail > 0 && failReasons.length) {
        var seenReason = {};
        var uniqReasons = [];
        for (var fr = 0; fr < failReasons.length; fr++) {
          var rsn = failReasons[fr];
          if (!rsn || seenReason[rsn]) continue;
          seenReason[rsn] = true;
          uniqReasons.push(rsn);
          if (uniqReasons.length >= 3) break;
        }
        if (uniqReasons.length) {
          var suffix = " (" + uniqReasons.join("; ") + ")";
          if (err) err += suffix;
          else if (msg) msg += suffix;
        }
      }
      showResult(msg, err);
      revealStatusOverlay();
    }

    // Handle spreadsheet items via streaming api/generate
    if (sheetItems.length > 0) {
      var cellsSeen = {};
      var cells = [];
      for (var si = 0; si < sheetItems.length; si++) {
        var ck = sheetItems[si].participant + "." + sheetItems[si].row;
        if (!cellsSeen[ck]) { cellsSeen[ck] = true; cells.push(ck); }
      }
      sheetCellTotal = cells.length;
      generateCardIndex = buildGenerateCardIndex(list);
      updateGenerateProgress(0, sheetCellTotal);

      function handleLine(line) {
        var data;
        try { data = JSON.parse(line); } catch (_) { return; }
        if (!data) return;
        if (data.cancelled) {
          cancelled = true;
          var queuedCards = list.querySelectorAll(".queue-card-queued");
          for (var qi = 0; qi < queuedCards.length; qi++) {
            clearCardStatus(queuedCards[qi]);
          }
          return;
        }
        if (!data.cell) return;
        sheetCellsDone++;
        updateGenerateProgress(sheetCellsDone, sheetCellTotal);
        updateGenerateButtonProgress();
        var cards = generateCardIndex[data.cell] || [];
        if (data.ok) {
          for (var ci = 0; ci < cards.length; ci++) setCardResult(cards[ci], true);
          totalSuccess += (data.generated || 1);
          if (data.artifacts) {
            allArtifacts = allArtifacts.concat(data.artifacts);
            for (var gi = 0; gi < data.artifacts.length; gi++) {
              state.generatedArtifacts.push(stampLog(data.artifacts[gi]));
            }
          }
        } else {
          // The server omits `error` when a clean run simply produced nothing.
          var reason = data.error || "No artifact produced";
          for (ci = 0; ci < cards.length; ci++) setCardResult(cards[ci], false, reason);
          failReasons.push(reason);
          totalFail++;
        }
      }

      var genBody = { cells: cells, format: format };
      var genOverrides = buildCellOverrides(sheetItems);
      if (Object.keys(genOverrides).length > 0) genBody.overrides = genOverrides;
      if (format === "clip") {
        var tcCb = qs("#titlecardEnabled");
        var tcDur = qs("#titlecardDuration");
        if (tcCb) genBody.titlecards_enabled = tcCb.checked;
        if (tcDur) genBody.titlecard_duration = parseInt(tcDur.value, 10) || 2;
      }

      // Streaming NDJSON response — needs the raw Response (reader + AbortSignal),
      // so this intentionally stays a manual fetch rather than an api* helper.
      fetch("api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(genBody),
        signal: sheetAbort.signal,
      })
        .then(function (response) {
          if (!response.ok) throw new Error("Server error " + response.status);
          return readNDJSONStream(response, handleLine).then(finishBranch);
        })
        .catch(function (err) {
          if (isGenerateFetchAborted(err)) {
            cancelled = true;
            for (var sq = 0; sq < sheetCardEls.length; sq++) {
              var sc = sheetCardEls[sq];
              if (sc && sc.classList.contains("queue-card-queued")) clearCardStatus(sc);
            }
            finishBranch();
            return;
          }
          // Mark every captured sheet card as failed so they don't stay
          // visually queued; finishBranch reports the failure tally.
          for (var j = 0; j < sheetCardEls.length; j++) {
            if (sheetCardEls[j]) setCardResult(sheetCardEls[j], false);
          }
          totalFail += sheetItems.length;
          finishBranch();
        });
    }

    // Handle intake items via api/generate-intake
    if (intakeItems.length > 0) {
      var intakePayload = intakeItems.map(function (itm) {
        return {
          participant: itm.participant,
          start: itm.start,
          end: itm.end,
          event_type: itm.event_type || itm.desc || "",
          event_ids: itm.event_ids || [],
          source: itm.source || "screenspace",
          mark_ids: itm.mark_ids || [],
        };
      });

      function handleIntakeLine(line) {
        var data;
        try { data = JSON.parse(line); } catch (_) { return; }
        if (!data) return;
        if (data.cancelled) {
          cancelled = true;
          // Clear queued state from any intake card that hasn't received a
          // per-item result yet, so the cards don't stay visually queued
          // after the server short-circuits on cancel.
          for (var qi = 0; qi < intakeCardEls.length; qi++) {
            var qcard = intakeCardEls[qi];
            if (qcard && qcard.classList.contains("queue-card-queued")) {
              clearCardStatus(qcard);
            }
          }
          return;
        }
        if (typeof data.index !== "number") return;
        var card = intakeCardEls[data.index];
        if (data.ok) {
          totalSuccess++;
          if (data.artifact) {
            allArtifacts.push(data.artifact);
            state.generatedArtifacts.push(stampLog(data.artifact));
          }
          if (card) setCardResult(card, true);
        } else {
          var reason = data.error || "Generation failed";
          totalFail++;
          failReasons.push(reason);
          if (card) setCardResult(card, false, reason);
        }
        intakeDone++;
        updateGenerateButtonProgress();
      }

      // Streaming NDJSON response — manual fetch is required to get a reader
      // and parse line-delimited per-item events as ffmpeg finishes each cut.
      fetch("api/generate-intake", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ items: intakePayload, format: format }),
        signal: intakeAbort.signal,
      })
        .then(function (response) {
          if (!response.ok) throw new Error("Server error " + response.status);
          return readNDJSONStream(response, handleIntakeLine).then(finishBranch);
        })
        .catch(function (err) {
          if (isGenerateFetchAborted(err)) {
            cancelled = true;
            for (var iq = 0; iq < intakeCardEls.length; iq++) {
              var ic = intakeCardEls[iq];
              if (ic && ic.classList.contains("queue-card-queued")) clearCardStatus(ic);
            }
            finishBranch();
            return;
          }
          for (var j = 0; j < intakeCardEls.length; j++) {
            if (intakeCardEls[j]) setCardResult(intakeCardEls[j], false);
          }
          totalFail += intakeItems.length;
          finishBranch();
        });
    }
  }

  function onCancelGenerate() {
    state.generateCancelledByUser = true;
    qs("#cancelGenerateBtn").classList.add("hidden");
    var aborts = state.activeGenerateAborts || [];
    for (var i = 0; i < aborts.length; i++) {
      try { aborts[i].abort(); } catch (_) {}
    }
    state.activeGenerateAborts = [];
    apiPost("api/generate/cancel").catch(function () {});
    apiPost("api/generate-intake/cancel").catch(function () {});
  }

  STUDIO.onGenerate = onGenerate;
  STUDIO.onCancelGenerate = onCancelGenerate;
})();
