/* clipgen Screenspace regions satellite — screenspace-regions.js
 *
 * Region stashing (stash cards, restore / rename / dismiss, drop targets) and
 * the region-chip drag state machine, carved out of screenspace.js. Loads
 * before screenspace-overlay-interaction.js, which destructures stashRegions
 * at load time; renderRegionChips / updateRegionButtons are therefore reached
 * late-bound through SS. Function bodies are unchanged from the hub.
 */
(function () {
  "use strict";

  var SS = window.ClipgenScreenspace;
  var state = SS.state;
  var iconSpan = SS.iconSpan,
    regionColorForIndex = SS.regionColorForIndex,
    regionToPixels = SS.regionToPixels,
    renderRunRegionPicker = SS.renderRunRegionPicker,
    updateRunButton = SS.updateRunButton;


  // Stash id whose card gets the landing animation; consumed on first render.
  var _justStashedStashId = null;

  function stashRegions() {
    var chips = qsa("#regionChips .region-chip");
    apiPost("api/stashes", {}).then(function (data) {
      if (!data.ok) return;
      var commit = function () {
        state.stashes.push(data.stash);
        _justStashedStashId = data.stash.id;
        state.regions = {};
        state.activeRegion = null;
        state.pendingRegion = null;
        SS.renderRegionChips();
        SS.renderOverlay();
        SS.updateRegionButtons();
        updateRunButton();
        renderStashCards();
        showToast("Regions stashed");
      };
      // Pills stash out (jump + wiggle + dissolve), then the stash card lands.
      if (chips.length && window.ClipgenMotion) ClipgenMotion.animateOutAll(chips, "stash").then(commit);
      else commit();
    }).catch(toastError("Could not create stash"));
  }

  function dismissStash(stashId) {
    apiDelete("api/stashes/" + stashId).then(function (data) {
      if (!data.ok) return;
      state.stashes = state.stashes.filter(function (s) { return s.id !== stashId; });
      // Remove any run-selected regions that belonged to this stash
      renderRunRegionPicker();
      renderStashCards();
      showToast("Stash dismissed");
    }).catch(toastError("Could not delete stash"));
  }

  function restoreStash(stashId) {
    apiPost("api/stashes/" + stashId + "/restore", {}).then(function (data) {
      if (!data.ok) return;
      state.regions = data.regions || {};
      state.activeRegion = null;
      state.pendingRegion = null;
      SS.renderRegionChips();
      SS.renderOverlay();
      SS.updateRegionButtons();
      updateRunButton();
      renderStashCards();
      showToast("Regions restored");
    }).catch(toastError("Could not restore stash"));
  }

  function copyRegionToStash(name, stashId) {
    if (!(name in state.regions)) return;
    apiPost("api/stashes/" + stashId + "/regions", { name: name })
      .then(function (data) {
        if (!data.ok) return;
        for (var i = 0; i < state.stashes.length; i++) {
          if (state.stashes[i].id === stashId) {
            state.stashes[i] = data.stash;
            break;
          }
        }
        renderStashCards(); // updated count + dots
        renderRunRegionPicker(); // stash folder now lists the new region
        showToast("Added “" + name + "” to " + data.stash.name);
      })
      .catch(function () {
        showToast("Failed to add region to stash");
      });
  }

  function renameStash(stashId, newName) {
    apiPut("api/stashes/" + stashId, { name: newName }).then(function (data) {
      if (!data.ok) return;
      for (var i = 0; i < state.stashes.length; i++) {
        if (state.stashes[i].id === stashId) {
          state.stashes[i].name = data.stash.name;
          break;
        }
      }
      renderRunRegionPicker();
    }).catch(toastError("Could not rename stash"));
  }

  function renderStashCards() {
    var existing = qs("#stashArea");
    if (state.stashes.length === 0) {
      if (existing) existing.remove();
      return;
    }
    var area = existing || el("div");
    area.id = "stashArea";
    area.innerHTML = "";

    var MAX_DOTS = 5;

    state.stashes.forEach(function (stash) {
      var card = el("div", "stash-card");
      card.dataset.stashId = stash.id;
      if (stash.id === _justStashedStashId && window.ClipgenMotion) {
        ClipgenMotion.animateIn(card, "stashLand");
        _justStashedStashId = null;
      }
      var regionNames = Object.keys(stash.regions);

      // Editable name
      var nameEl = el("span", "stash-card-name", stash.name);
      nameEl.setAttribute("contenteditable", "true");
      nameEl.setAttribute("spellcheck", "false");
      nameEl.addEventListener("blur", function () {
        var trimmed = nameEl.textContent.trim();
        if (trimmed && trimmed !== stash.name) {
          renameStash(stash.id, trimmed);
        } else {
          nameEl.textContent = stash.name;
        }
      });
      nameEl.addEventListener("keydown", function (e) {
        if (e.key === "Enter") { e.preventDefault(); nameEl.blur(); }
      });
      card.appendChild(nameEl);

      // Separator dot and count
      card.appendChild(el("span", "stash-card-sep", "\u00b7"));
      card.appendChild(el("span", "stash-card-count", regionNames.length + " region" + (regionNames.length !== 1 ? "s" : "")));

      // Colored dots (max 5, with fade-out on 6th)
      var dots = el("span", "stash-card-dots");
      var showCount = Math.min(regionNames.length, MAX_DOTS);
      for (var i = 0; i < showCount; i++) {
        var dot = el("span", "region-chip-dot");
        dot.style.background = regionColorForIndex(i);
        dot.title = regionNames[i];
        dots.appendChild(dot);
      }
      if (regionNames.length > MAX_DOTS) {
        var fadeDot = el("span", "region-chip-dot stash-dot-fade");
        fadeDot.style.background = regionColorForIndex(MAX_DOTS);
        dots.appendChild(fadeDot);
      }
      card.appendChild(dots);

      // Action buttons
      var actions = el("span", "stash-card-actions");

      var restoreBtn = el("button", "stash-card-action-btn");
      restoreBtn.title = "Restore regions";
      restoreBtn.appendChild(iconSpan("arrow-up-tray"));
      restoreBtn.addEventListener("click", function () { restoreStash(stash.id); });
      actions.appendChild(restoreBtn);

      var dismissBtn = el("button", "stash-card-action-btn");
      dismissBtn.title = "Dismiss stash";
      dismissBtn.appendChild(iconSpan("x-mark"));
      dismissBtn.addEventListener("click", function () { dismissStash(stash.id); });
      actions.appendChild(dismissBtn);

      card.appendChild(actions);

      // Hover preview: show stashed regions on the overlay
      card.addEventListener("mouseenter", function () {
        state.previewRegions = stash.regions;
        SS.renderOverlay();
      });
      card.addEventListener("mouseleave", function () {
        state.previewRegions = null;
        SS.renderOverlay();
      });

      area.appendChild(card);
    });

    if (!existing) {
      bindStashDrop(area);
      var viewerSection = qs("#viewerSection");
      viewerSection.parentNode.insertBefore(area, viewerSection.nextSibling);
    }
  }

  // Drop handlers for copying a chip into a stash, bound once on the persistent
  // #stashArea.
  var _stashDragOverCard = null;

  function clearStashDragIndicators() {
    if (_stashDragOverCard) {
      _stashDragOverCard.classList.remove("drag-over");
      _stashDragOverCard = null;
    }
    qsa(".stash-card.drag-over").forEach(function (card) {
      card.classList.remove("drag-over");
    });
  }

  function bindStashDrop(area) {
    area.addEventListener("dragover", function (e) {
      if (!hasRegionDragPayload(e)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "copy";
      var card = e.target.closest(".stash-card");
      if (_stashDragOverCard && _stashDragOverCard !== card) {
        _stashDragOverCard.classList.remove("drag-over");
      }
      if (card) {
        card.classList.add("drag-over");
        _stashDragOverCard = card;
      }
    });
    area.addEventListener("dragleave", function (e) {
      var card = e.target.closest(".stash-card");
      if (card && !card.contains(e.relatedTarget)) {
        card.classList.remove("drag-over");
        if (_stashDragOverCard === card) _stashDragOverCard = null;
      }
    });
    area.addEventListener("drop", function (e) {
      if (!hasRegionDragPayload(e)) return;
      var card = e.target.closest(".stash-card");
      if (!card) return;
      e.preventDefault();
      clearStashDragIndicators();
      _regionDragDropped = true;
      var regionName = getDraggedRegionName(e);
      if (regionName) copyRegionToStash(regionName, card.dataset.stashId);
    });
  }

  // ---- Region chip drag ----
  // Mirrors the vertical task drag; excluding .dragging keeps indices aligned.
  var REGION_DRAG_MIME = "application/x-region-name";
  var _regionDragMidpoints = null;
  var _regionDragOverRaf = null;
  var _regionPendingDragOverX = null;
  var _regionDragActive = false;
  var _regionDragMoved = false;
  var _regionDragDropped = false;

  function dataTransferHasType(dt, type) {
    if (!dt || !dt.types) return false;
    if (typeof dt.types.indexOf === "function") return dt.types.indexOf(type) >= 0;
    if (typeof dt.types.contains === "function") return dt.types.contains(type);
    for (var i = 0; i < dt.types.length; i++) {
      if (dt.types[i] === type) return true;
    }
    return false;
  }

  function hasRegionDragPayload(e) {
    // Firefox/Safari may hide custom MIME types on dragover; the local flag is reliable.
    return _regionDragActive || dataTransferHasType(e.dataTransfer, REGION_DRAG_MIME);
  }

  function setRegionDragData(dt, idx, name) {
    if (!dt) return;
    dt.setData("text/plain", String(idx));
    try {
      dt.setData(REGION_DRAG_MIME, name);
    } catch (_) {
      // Some engines reject custom types; text/plain + local state covers us.
    }
  }

  function getDraggedRegionName(e) {
    var name = "";
    try {
      name = e.dataTransfer.getData(REGION_DRAG_MIME);
    } catch (_) {
      name = "";
    }
    if (name && Object.prototype.hasOwnProperty.call(state.regions, name)) return name;
    var fromIdx = parseInt(e.dataTransfer.getData("text/plain"), 10);
    if (isNaN(fromIdx)) return "";
    return Object.keys(state.regions)[fromIdx] || "";
  }

  function getDraggedRegionIndex(e) {
    var fromIdx = parseInt(e.dataTransfer.getData("text/plain"), 10);
    if (!isNaN(fromIdx)) return fromIdx;
    var name = getDraggedRegionName(e);
    return name ? Object.keys(state.regions).indexOf(name) : -1;
  }

  function _cacheRegionDragMidpoints(container) {
    var chips = container.querySelectorAll(".region-chip:not(.dragging)");
    var mids = new Array(chips.length);
    for (var i = 0; i < chips.length; i++) {
      var r = chips[i].getBoundingClientRect();
      mids[i] = r.left + r.width / 2;
    }
    _regionDragMidpoints = mids;
  }

  function getRegionDropIndex(container, clientX) {
    var mids = _regionDragMidpoints;
    if (!mids) {
      _cacheRegionDragMidpoints(container);
      mids = _regionDragMidpoints;
    }
    for (var i = 0; i < mids.length; i++) {
      if (clientX < mids[i]) return i;
    }
    return mids.length;
  }

  function clearRegionDragIndicators(container) {
    var chips = container.querySelectorAll(".region-chip.drag-over");
    for (var i = 0; i < chips.length; i++) chips[i].classList.remove("drag-over");
    container.classList.remove("drag-over-append");
  }

  function initRegionDrag() {
    var chips = qs("#regionChips");

    chips.addEventListener("dragstart", function (e) {
      var chip = e.target.closest(".region-chip");
      if (!chip) {
        e.preventDefault();
        return;
      }
      _regionDragActive = true;
      _regionDragMoved = false;
      _regionDragDropped = false;
      chip.classList.add("dragging");
      setRegionDragData(e.dataTransfer, chip.dataset.regionIdx, chip.dataset.regionName);
      e.dataTransfer.effectAllowed = "copyMove";
      _cacheRegionDragMidpoints(chips);
    });

    chips.addEventListener("dragend", function (e) {
      var chip = e.target.closest(".region-chip");
      if (chip) chip.classList.remove("dragging");
      if (_regionDragOverRaf != null) {
        cancelAnimationFrame(_regionDragOverRaf);
        _regionDragOverRaf = null;
      }
      _regionPendingDragOverX = null;
      clearRegionDragIndicators(chips);
      clearStashDragIndicators();
      _regionDragMidpoints = null;
      if (_regionDragMoved || _regionDragDropped) {
        state.regionSuppressNextClick = true;
        setTimeout(function () { state.regionSuppressNextClick = false; }, 250);
      }
      _regionDragActive = false;
      _regionDragMoved = false;
      _regionDragDropped = false;
    });

    chips.addEventListener("dragover", function (e) {
      if (!hasRegionDragPayload(e)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      _regionDragMoved = true;
      _regionPendingDragOverX = e.clientX;
      if (_regionDragOverRaf != null) return; // RAF-debounce ~60Hz dragover
      _regionDragOverRaf = requestAnimationFrame(function () {
        _regionDragOverRaf = null;
        if (_regionPendingDragOverX == null) return;
        clearRegionDragIndicators(chips);
        var visible = chips.querySelectorAll(".region-chip:not(.dragging)");
        var idx = getRegionDropIndex(chips, _regionPendingDragOverX);
        if (idx < visible.length) visible[idx].classList.add("drag-over");
        else chips.classList.add("drag-over-append");
      });
    });

    chips.addEventListener("dragleave", function (e) {
      var chip = e.target.closest(".region-chip");
      if (chip) chip.classList.remove("drag-over");
      if (!chips.contains(e.relatedTarget)) chips.classList.remove("drag-over-append");
    });

    chips.addEventListener("drop", function (e) {
      if (!hasRegionDragPayload(e)) return;
      e.preventDefault();
      _regionDragDropped = true;
      clearRegionDragIndicators(chips);
      var fromIdx = getDraggedRegionIndex(e);
      if (fromIdx < 0) return;
      var toIdx = getRegionDropIndex(chips, e.clientX);
      if (fromIdx === toIdx) return;
      var names = Object.keys(state.regions);
      var previousRegions = state.regions;
      var moved = names.splice(fromIdx, 1)[0];
      if (!moved) return;
      names.splice(toIdx, 0, moved);
      // Rebuild state.regions in the new order.
      var reordered = {};
      names.forEach(function (n) {
        reordered[n] = state.regions[n];
      });
      state.regions = reordered;
      // Colors are position-based, so reordering recolors regions on purpose; repaint
      // both.
      SS.renderRegionChips();
      SS.renderOverlay();
      apiPut("api/regions/reorder", { names: names }).then(function (data) {
        if (!data || !data.ok) throw new Error((data && data.error) || "reorder failed");
      }).catch(function () {
        state.regions = previousRegions;
        SS.renderRegionChips();
        SS.renderOverlay();
        showToast("Failed to save region order");
      });
    });

    document.addEventListener("dragend", function () {
      clearStashDragIndicators();
    });
    document.addEventListener("drop", function () {
      clearStashDragIndicators();
    });
  }

  SS.initRegionDrag = initRegionDrag;
  SS.renderStashCards = renderStashCards;
  SS.stashRegions = stashRegions;
})();
