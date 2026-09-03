/* clipgen Studio — artifact-generation satellite.
 *
 * Carved out of studio.js (the hub) following the hub+satellite convention
 * (see studio-scrubber/trim/intake). Owns the streaming api/generate +
 * api/generate-intake flow: the Generate button handler (onGenerate), its cancel
 * handler (onCancelGenerate), and the per-cell card index helper.
 *
 * Reaches hub-owned pieces through window.ClipgenStudio (STUDIO): the card-state
 * painters (setCardQueued/clearCardStatus/setCardResult) are shared with the
 * reel/build path and stay in the hub; setArtifactGenerating / showResult /
 * revealStatusOverlay / updateGenerateProgress / stampLog likewise; the
 * _generateEtaTracker + _studioEtaTicker objects are shared elapsed-time
 * infrastructure (also driven by job-status polling and the reel/build flows);
 * buildCellOverrides lives in studio-trim.js (reached via STUDIO). isIntakeSource
 * is a hub helper. apiPost / qs / apiPostNDJSON / setButtonProgress /
 * clipgenPluralUnit are ambient utils.js / primitives.js globals (scope chain).
 *
 * The hub keeps same-named onGenerate/onCancelGenerate delegators for the button
 * wiring. Loaded by studio.html after studio.js and studio-trim.js (whose
 * STUDIO.buildCellOverrides this file uses), before studio-intake.js.
 */

(function () {
  "use strict";

  var STUDIO = window.ClipgenStudio;
  var state = STUDIO.state;
  // primitives.js global, aliased as in the hub.
  var setButtonProgress = ClipgenPrimitives.setButtonProgress;
  // Published by the hub (buildCellOverrides by studio-trim.js) before this file loads.
  var setArtifactGenerating = STUDIO.setArtifactGenerating,
    showResult = STUDIO.showResult,
    revealStatusOverlay = STUDIO.revealStatusOverlay,
    setCardQueued = STUDIO.setCardQueued,
    clearCardStatus = STUDIO.clearCardStatus,
    setCardResult = STUDIO.setCardResult,
    updateGenerateProgress = STUDIO.updateGenerateProgress,
    stampLog = STUDIO.stampLog,
    isIntakeSource = STUDIO.isIntakeSource,
    buildCellOverrides = STUDIO.buildCellOverrides,
    _generateEtaTracker = STUDIO._generateEtaTracker,
    _studioEtaTicker = STUDIO._studioEtaTicker;

  // Result lines echo the sheet header's casing; every cell-ref map keys on this.
  function generateCellKey(ref) {
    return String(ref).toLowerCase();
  }

  function buildGenerateCardIndex(listEl) {
    var map = {};
    var cards = listEl.querySelectorAll(".queue-card");
    for (var i = 0; i < cards.length; i++) {
      var card = cards[i];
      var participant = card.getAttribute("data-participant");
      var row = card.getAttribute("data-row");
      if (!participant || row == null) continue;
      var key = generateCellKey(participant + "." + row);
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

    // onCancelGenerate aborts these; the server cancel endpoints kill ffmpeg.
    var sheetAbort = new AbortController();
    var intakeAbort = new AbortController();
    state.activeGenerateAborts = [sheetAbort, intakeAbort];

    var format = qs("#artifactFormat").value;
    var list = qs("#artifactsList");
    var items = state.artifactQueue.slice();

    // Capture cards before async work; a mid-request re-render must not move markers.
    var allCards = list.querySelectorAll(".queue-card");
    for (var i = 0; i < allCards.length; i++) {
      setCardQueued(allCards[i]);
    }

    // Each split keeps its card list parallel to its item list.
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
    // Counts are artifacts (one per card), never cells; a cell can hold several pairs.
    var sheetArtifactTotal = sheetItems.length;
    var sheetArtifactsDone = 0;
    var intakeDone = 0;
    var intakeTotal = intakeItems.length;
    var generateCardIndex = null;
    updateGenerateProgress(0, sheetArtifactTotal + intakeTotal);

    // One readout for both branches; /api/job-status reattaches to the same unit.
    function updateGenerateButtonProgress() {
      var total = sheetArtifactTotal + intakeTotal;
      if (total <= 0) return;
      var done = sheetArtifactsDone + intakeDone;
      setButtonProgress("generateBtn", done / total);
      updateGenerateProgress(done, total);
    }

    function finishBranch() {
      if (--pending > 0) return;
      setButtonProgress("generateBtn", null);
      setArtifactGenerating(false);
      _generateEtaTracker.reset();
      // After setArtifactGenerating(false), or _paintGenerateProgress keeps the readout up.
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
        // No per-item results at all is an error, not "Generated 0 artifacts".
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
      // Up to 3 distinct reasons; each card title holds the full one.
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
      // One ref per cell; cardsPerCell lets a result line advance by its card count.
      var cellsSeen = {};
      var cells = [];
      var cardsPerCell = {};
      for (var si = 0; si < sheetItems.length; si++) {
        // POST the ref as written; tally on the folded key.
        var ref = sheetItems[si].participant + "." + sheetItems[si].row;
        var ck = generateCellKey(ref);
        if (!cellsSeen[ck]) { cellsSeen[ck] = true; cells.push(ref); }
        cardsPerCell[ck] = (cardsPerCell[ck] || 0) + 1;
      }
      var cellCounted = {};
      generateCardIndex = buildGenerateCardIndex(list);

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
        // One advance per cell; a result line and "No clip found" may both arrive.
        var cellKey = generateCellKey(data.cell);
        if (!cellCounted[cellKey]) {
          cellCounted[cellKey] = true;
          sheetArtifactsDone += cardsPerCell[cellKey] || 1;
          updateGenerateButtonProgress();
        }
        var cards = generateCardIndex[cellKey] || [];
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

      apiPostNDJSON("api/generate", genBody, {
        signal: sheetAbort.signal,
        onLine: handleLine,
      })
        .then(finishBranch)
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
          // Fail every captured sheet card; finishBranch reports the tally.
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
          text: itm.text || "",
          label: itm.label || "",
        };
      });

      function handleIntakeLine(line) {
        var data;
        try { data = JSON.parse(line); } catch (_) { return; }
        if (!data) return;
        if (data.cancelled) {
          cancelled = true;
          // Cancel short-circuits the server; clear cards still marked queued.
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

      apiPostNDJSON(
        "api/generate-intake",
        { items: intakePayload, format: format },
        { signal: intakeAbort.signal, onLine: handleIntakeLine }
      )
        .then(finishBranch)
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
