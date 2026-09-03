/* clipgen Studio — reel/artifact stash satellite.
 *
 * Carved out of studio.js (the hub) following the hub+satellite convention used
 * by studio-scrubber.js / studio-trim.js / studio-generate.js / studio-intake.js.
 * Owns both stash panels (stashed reels + stashed artifacts): load/render, the
 * stacked-folder cards with lazy thumbnails, stash/recall/delete/rename, the
 * drop-target handlers for stash-to-stash drags, and the empty-area drag reveal.
 *
 * Cross-file contract: the satellite reaches hub state and helpers through the
 * window.ClipgenStudio (STUDIO) namespace — state, renderReelQueue /
 * renderArtifactQueue, isReelQueueLocked / isArtifactQueueLocked, cellKey,
 * updateSingleCellClass, and the hub-retained lazy-thumb queue's
 * ssEnqueueThumbCustom; all destructured at load time, which is safe because
 * this file loads after the hub IIFE has published them. apiGet / apiPost /
 * toastError / qs / qsa / el / truncate / formatDuration / categoryHue are
 * ambient utils.js globals reached via the scope chain. The hub calls back in
 * via same-named guarded delegators (loadStashes, loadArtifactStashes,
 * stashCurrentReel, stashCurrentArtifacts, revealEmptyStashAreas,
 * hideEmptyStashAreas) plus the late-bound drop handlers (stashDropReel,
 * stashDropArtifacts). Loaded by studio.html after studio.js; order relative to
 * the other satellites is free (no cross-destructuring).
 */

(function () {
  "use strict";

  var STUDIO = window.ClipgenStudio;
  var state = STUDIO.state,
    renderReelQueue = STUDIO.renderReelQueue,
    renderArtifactQueue = STUDIO.renderArtifactQueue,
    isReelQueueLocked = STUDIO.isReelQueueLocked,
    isArtifactQueueLocked = STUDIO.isArtifactQueueLocked,
    cellKey = STUDIO.cellKey,
    updateSingleCellClass = STUDIO.updateSingleCellClass,
    ssEnqueueThumbCustom = STUDIO.ssEnqueueThumbCustom;

  // ---- Stashed reels ----

  var REEL_STASH = {
    stateKey: "stashes",
    apiPath: "api/stashes",
    countSel: "#stashedReelsCount",
    areaSel: "#stashedReelsArea",
    listSel: "#stashedReelsList",
    dragSource: "reel-stash",
    emptyHint: "Stash reels to set them aside for later.",
    queueKey: "reelQueue",
    isLocked: isReelQueueLocked,
    renderQueue: renderReelQueue,
  };
  var ARTIFACT_STASH = {
    stateKey: "artifactStashes",
    apiPath: "api/artifact-stashes",
    countSel: "#stashedArtifactsCount",
    areaSel: "#stashedArtifactsArea",
    listSel: "#stashedArtifactsList",
    dragSource: "artifact-stash",
    emptyHint: "Stash artifacts to keep them aside. Drag them here or use the Stash button.",
    queueKey: "artifactQueue",
    isLocked: isArtifactQueueLocked,
    renderQueue: renderArtifactQueue,
  };

  // One-shot: the matching card animates in once, then nulls this.
  var _justStashedId = null;

  function loadStashes() {
    apiGet("api/stashes")
      .then(function (data) {
        if (data.ok) {
          state.stashes = data.stashes || [];
          renderStashedReels();
        }
      })
      .catch(toastError("Failed to load stashes"));
  }

  function renderStashes(cfg) {
    var area = qs(cfg.areaSel);
    var list = qs(cfg.listSel);
    var arr = state[cfg.stateKey];
    var n = arr.length;
    qs(cfg.countSel).textContent = "(" + n + ")";
    area.classList.remove("stash-drop-reveal");
    list.innerHTML = "";

    if (n === 0) {
      list.appendChild(el("div", "stash-empty-hint", cfg.emptyHint));
      return;
    }

    var rerender = function () { renderStashes(cfg); };
    var onRecall = function (s) { recallStashItem(cfg, s); };
    for (var i = 0; i < n; i++) {
      list.appendChild(buildStashCard(arr[i], cfg.apiPath, arr, rerender, cfg.dragSource, onRecall));
    }
  }

  function renderStashedReels() {
    renderStashes(REEL_STASH);
  }

  function makeStashFolderIcon(stash) {
    var icon = el("span", "stash-card-icon");
    var hue = categoryHue((stash && stash.id) || "uncategorized");
    // Hue-tinted backing remains as a fallback if every thumbnail fails.
    icon.style.background = "oklch(0.32 0.06 " + hue + ")";

    var items = (stash && stash.items) || [];
    // Pull the first 3 distinct thumbnails to fake a stacked-folder look.
    var picks = [];
    var seen = {};
    for (var i = 0; i < items.length && picks.length < 3; i++) {
      var item = items[i];
      var key = item && item.participant != null && item.start != null
        ? item.participant + ":" + item.start
        : null;
      if (!key || seen[key]) continue;
      seen[key] = true;
      picks.push(item);
    }

    picks.forEach(function (item, idx) {
      var img = document.createElement("img");
      img.decoding = "async";
      img.className = "stash-card-icon-img";
      img.alt = "";
      img.draggable = false;
      img.style.zIndex = String(picks.length - idx);
      img.style.transform = "translate(" + (idx * 2) + "px, " + (-idx * 2) + "px)";
      // Append first: ssProcessQueue runs synchronously and skips detached imgs.
      icon.appendChild(img);
      ssEnqueueThumbCustom(img, item.participant, item.start, function () {
        if (img.parentNode) img.parentNode.removeChild(img);
      });
    });

    return icon;
  }

  function buildStashCard(stash, apiPath, listRef, rerender, dragSource, onRecall) {
    var card = el("div", "stash-card");
    card.setAttribute("data-stash-id", stash.id);
    if (stash.id === _justStashedId) {
      if (window.ClipgenMotion) ClipgenMotion.animateIn(card, "stashLand");
      _justStashedId = null;
    }
    card.setAttribute("draggable", "true");
    card.addEventListener("dragstart", function (ev) {
      ev.dataTransfer.setData("application/json", JSON.stringify({
        stashId: stash.id,
        items: stash.items,
        source: dragSource,
      }));
      ev.dataTransfer.effectAllowed = "copy";
    });

    card.appendChild(makeStashFolderIcon(stash));

    var nameEl = el("span", "stash-card-name", truncate(stash.name, 18));
    nameEl.title = stash.name;
    nameEl.addEventListener("click", function (ev) {
      ev.stopPropagation();
      startStashRename(stash, nameEl, apiPath);
    });
    card.appendChild(nameEl);

    var info = el("span", "stash-card-info");
    info.appendChild(el("span", "", String(stash.count)));
    info.appendChild(el("span", "", formatDuration(stash.totalDuration)));
    card.appendChild(info);

    var removeBtn = el("button", "stash-card-remove", "\u00D7");
    removeBtn.title = "Delete stash";
    removeBtn.addEventListener("click", function (ev) {
      ev.stopPropagation();
      deleteStash(stash.id, apiPath, listRef, rerender);
    });
    card.appendChild(removeBtn);

    card.addEventListener("click", function () {
      if (typeof onRecall === "function") onRecall(stash);
    });
    return card;
  }

  function computeReelDuration(items) {
    var total = 0;
    for (var i = 0; i < items.length; i++) {
      total += Math.max(0, items[i].end - items[i].start);
    }
    return total;
  }

  function stashCurrent(cfg) {
    if (cfg.isLocked() || state[cfg.queueKey].length === 0) return;

    var items = state[cfg.queueKey].slice();
    // cfg.listSel points at the *stashed* list; the source cards live in the queue.
    var sourceSel = cfg.queueKey === "reelQueue" ? "#reelList" : "#artifactsList";
    var cards = qsa(sourceSel + " .queue-card");
    createStashViaAPI(cfg.apiPath, items, function (stash) {
      var commit = function () {
        state[cfg.stateKey].push(stash);
        _justStashedId = stash.id;
        var q = state[cfg.queueKey];
        for (var i = 0; i < q.length; i++) {
          var item = q[i];
          delete state.cellResults[cellKey(item.participant, item.row)];
        }
        state[cfg.queueKey] = [];
        cfg.renderQueue();
        renderStashes(cfg);
        for (var u = 0; u < items.length; u++) {
          if (items[u].row) updateSingleCellClass(items[u].participant, items[u].row);
        }
      };
      // Queue cards stash out, then the new stash card lands (renderStashes).
      if (cards.length && window.ClipgenMotion) ClipgenMotion.animateOutAll(cards, "stash").then(commit);
      else commit();
    });
  }

  function stashCurrentReel() {
    stashCurrent(REEL_STASH);
  }

  function recallStashItem(cfg, stash) {
    if (cfg.isLocked()) return;
    // Deep copy: the trim pop-over edits queue items in place.
    state[cfg.queueKey] = JSON.parse(JSON.stringify(stash.items));
    cfg.renderQueue();
    var q = state[cfg.queueKey];
    for (var i = 0; i < q.length; i++) {
      var it = q[i];
      if (it.row) updateSingleCellClass(it.participant, it.row);
    }
  }

  function recallStash(stash) {
    recallStashItem(REEL_STASH, stash);
  }

  function deleteStash(stashId, endpoint, stateArray, renderFn) {
    apiPost(endpoint, { action: "delete", id: stashId })
      .then(function (data) {
        if (data.ok) {
          for (var i = 0; i < stateArray.length; i++) {
            if (stateArray[i].id === stashId) {
              stateArray.splice(i, 1);
              break;
            }
          }
          renderFn();
        }
      })
      .catch(toastError("Failed to delete stash"));
  }

  function startStashRename(stash, nameNode, endpoint) {
    var parent = nameNode.parentNode;
    var input = document.createElement("input");
    input.className = "stash-card-name-input";
    input.type = "text";
    input.autocomplete = "off";
    input.value = stash.name;

    function commit() {
      var newName = input.value.trim() || stash.name;
      stash.name = newName;
      var span = el("span", "stash-card-name", truncate(newName, 20));
      span.title = newName;
      span.addEventListener("click", function (ev) {
        ev.stopPropagation();
        startStashRename(stash, span, endpoint);
      });
      parent.replaceChild(span, input);

      apiPost(endpoint, { action: "update", id: stash.id, name: newName }).catch(toastError("Failed to rename stash"));
    }

    input.addEventListener("blur", commit);
    input.addEventListener("keydown", function (ev) {
      if (ev.key === "Enter") { ev.preventDefault(); input.blur(); }
      if (ev.key === "Escape") { input.value = stash.name; input.blur(); }
    });
    input.addEventListener("click", function (ev) { ev.stopPropagation(); });

    parent.replaceChild(input, nameNode);
    input.focus();
    input.select();
  }

  function createStashViaAPI(endpoint, items, onSuccess) {
    var totalDuration = computeReelDuration(items);
    apiPost(endpoint, { action: "create", items: items, name: "", totalDuration: totalDuration })
      .then(function (data) {
        if (data.ok) onSuccess(data.stash);
      })
      .catch(toastError("Failed to save stash"));
  }

  // ---- Stashed artifacts ----

  function loadArtifactStashes() {
    apiGet("api/artifact-stashes")
      .then(function (data) {
        if (data.ok) {
          state.artifactStashes = data.stashes || [];
          renderStashedArtifacts();
        }
      })
      .catch(toastError("Failed to load stashes"));
  }

  function renderStashedArtifacts() {
    renderStashes(ARTIFACT_STASH);
  }

  function stashCurrentArtifacts() {
    stashCurrent(ARTIFACT_STASH);
  }

  function recallArtifactStash(stash) {
    recallStashItem(ARTIFACT_STASH, stash);
  }

  // ---- Stash drag-reveal ----

  function revealEmptyStashAreas() {
    var artArea = qs("#stashedArtifactsArea");
    var reelArea = qs("#stashedReelsArea");
    if (state.artifactStashes.length === 0) artArea.classList.add("stash-drop-reveal");
    if (state.stashes.length === 0) reelArea.classList.add("stash-drop-reveal");
  }

  function hideEmptyStashAreas() {
    var artArea = qs("#stashedArtifactsArea");
    var reelArea = qs("#stashedReelsArea");
    artArea.classList.remove("stash-drop-reveal");
    reelArea.classList.remove("stash-drop-reveal");
  }

  // ---- Drop targets (stash-to-stash drags) ----
  // Called late-bound from the hub's setupDropTarget wiring.

  function stashDroppedItems(cfg, info) {
    if (info.source !== "reel-stash" && info.source !== "artifact-stash") return;
    createStashViaAPI(cfg.apiPath, info.items, function (stash) {
      state[cfg.stateKey].push(stash);
      _justStashedId = stash.id;
      renderStashes(cfg);
    });
  }

  function stashDropReel(info) {
    stashDroppedItems(REEL_STASH, info);
  }

  function stashDropArtifacts(info) {
    stashDroppedItems(ARTIFACT_STASH, info);
  }

  STUDIO.loadStashes = loadStashes;
  STUDIO.loadArtifactStashes = loadArtifactStashes;
  STUDIO.stashCurrentReel = stashCurrentReel;
  STUDIO.stashCurrentArtifacts = stashCurrentArtifacts;
  STUDIO.revealEmptyStashAreas = revealEmptyStashAreas;
  STUDIO.hideEmptyStashAreas = hideEmptyStashAreas;
  STUDIO.stashDropReel = stashDropReel;
  STUDIO.stashDropArtifacts = stashDropArtifacts;
})();
