/* clipgen Studio — intake satellite (Screenspace + Transcript intake panels).
 *
 * Carved out of studio.js (the hub) following the hub+satellite convention used
 * by transcripts.js / screenspace.js. Shares the hub's mutable `state` and the
 * helper functions it needs through the `window.ClipgenStudio` (STUDIO)
 * namespace; the hub calls back in via same-named guarded delegators. Loaded by
 * studio.html immediately after studio.js.
 *
 * Kept in the hub (NOT moved here; reached via STUDIO): `buildXrefBadges` (also
 * the legacy `window._studioBuildXrefBadges` that convergence.js reads) and the
 * `ss*Thumb*` lazy-thumbnail cluster (`ssClearPending`), which the hub's queue
 * cards share.
 */

(function () {
  "use strict";

  var STUDIO = window.ClipgenStudio;
  var state = STUDIO.state;
  var attachQueueScrubbers = STUDIO.attachQueueScrubbers,
    buildQueueCardThumb = STUDIO.buildQueueCardThumb,
    buildXrefBadges = STUDIO.buildXrefBadges,
    checkConvergenceTabVisibility = STUDIO.checkConvergenceTabVisibility,
    findIntakeInQueue = STUDIO.findIntakeInQueue,
    findOverlappingData = STUDIO.findOverlappingData,
    intakeAddItem = STUDIO.intakeAddItem,
    intakeToggleItem = STUDIO.intakeToggleItem,
    refreshMetadataIfActive = STUDIO.refreshMetadataIfActive,
    renderArtifactQueue = STUDIO.renderArtifactQueue,
    renderReelQueue = STUDIO.renderReelQueue,
    setCardDragImage = STUDIO.setCardDragImage,
    ssClearPending = STUDIO.ssClearPending;
  // Pure intake clustering (aliased the same way in the hub at load).
  var clusterIntakeEvents = window.ClipgenIntakeCluster.clusterIntakeEvents;
  var clusterTranscriptMarks = window.ClipgenIntakeCluster.clusterTranscriptMarks;

  // ---- Screenspace Intake ----

  var INTAKE_DETECTOR_COLORS = DETECTOR_COLORS;
  // Ordered detector list for the intake filter chips — copied from utils.js's
  // canonical `_DETECTOR_TYPES` (includes `boundary`) so a new detector automatically
  // gets a chip instead of being silently un-filterable.
  var INTAKE_DETECTORS = _DETECTOR_TYPES.slice();

  function setTabDot(elId, on) {
    var el = document.getElementById(elId);
    if (!el) return;
    if (on) el.classList.remove("hidden");
    else el.classList.add("hidden");
  }

  // ---- Screenspace intake: poll Screenspace/Transcripts + cluster for Studio ----

  function setIntakeDot(on) {
    if (state._intakeTabDotOn === on) return;
    state._intakeTabDotOn = on;
    setTabDot("intakeTabDot", on);
  }

  function setTrIntakeDot(on) {
    if (state._trIntakeTabDotOn === on) return;
    state._trIntakeTabDotOn = on;
    setTabDot("trIntakeTabDot", on);
  }

  // Screenspace intake poll: one combined request (../screenspace/api/intake-poll)
  // for the task-status dot + curation events. Returns a Promise<boolean> — truthy
  // ("active this tick") keeps createPoller at the fast cadence; falsy lets it back
  // off. Replaces the former pollIntakeStatus + pollIntakeEvents pair.
  function pollScreenspaceIntake() {
    return apiGet("../screenspace/api/intake-poll?excluded=false")
      .then(function (data) {
        if (!data || !data.ok) {
          setIntakeDot(false);
          return false;
        }
        var status = data.status || {};
        var dotOn = !!status.running || (status.worker_alive && !!status.queued);
        setIntakeDot(dotOn);
        var eventsChanged = applyIntakeEvents(data.events || []);
        return dotOn || eventsChanged;
      })
      .catch(function () {
        setIntakeDot(false);
        return false;
      });
  }

  // Transcript intake poll: one combined request (../transcripts/api/intake-poll)
  // for the running-state dot + resolved marks. Returns a Promise<boolean> as above.
  // Replaces the former pollTrIntakeStatus (3-request fan-out) + pollTranscriptIntakeMarks.
  function pollTranscriptIntake() {
    return apiGet("../transcripts/api/intake-poll")
      .then(function (data) {
        if (!data || !data.ok) {
          setTrIntakeDot(false);
          return false;
        }
        var status = data.status || {};
        var running = !!status.tasks_running || !!status.model_warming || !!status.agents_running;
        setTrIntakeDot(running);
        var marksChanged = applyTranscriptMarks(data.marks || {});
        return running || marksChanged;
      })
      .catch(function () {
        setTrIntakeDot(false);
        return false;
      });
  }

  // On-demand refresh after a user action: wake the adaptive poller so it
  // refetches now AND resets its idle backoff to the fast cadence. Falls back to
  // a direct poll if the poller hasn't been created yet (pre-boot).
  function refreshScreenspaceIntake() {
    if (state.ssIntakePoller) state.ssIntakePoller.wake();
    else pollScreenspaceIntake();
  }
  function refreshTranscriptIntake() {
    if (state.trIntakePoller) state.trIntakePoller.wake();
    else pollTranscriptIntake();
  }

  // Events fed into the intake clustering surface. Navigational (boundary)
  // events are orientation scaffolding, not clip candidates — hidden by
  // default so they don't flood the curation queue. state.intakeEvents still
  // holds ALL events (Metadata's boundary count reads from it).
  function intakeClusterSource() {
    if (state.intakeShowNavigational) return state.intakeEvents;
    return state.intakeEvents.filter(function (ev) { return !ev.navigational; });
  }

  function reclusterIntake() {
    var threshold = parseInt((qs("#intakeClusterThreshold") || {}).value, 10) || 10;
    state.intakeClusters = clusterIntakeEvents(intakeClusterSource(), threshold);
    renderIntake(false);
  }

  // One-click boundary detection: enqueue a full-frame boundary task per
  // participant that has a source video. Progress surfaces through the existing
  // task poll (intake tab dot) and new boundary events arrive via pollScreenspaceIntake.
  function intakeDetectBoundaries(btn) {
    var labelEl = btn ? btn.querySelector("span") : null;
    var origLabel = labelEl ? labelEl.textContent : "";
    function restore() {
      if (btn) btn.disabled = false;
      if (labelEl) labelEl.textContent = origLabel;
    }
    apiGet("../screenspace/api/participants")
      .then(function (data) {
        if (!data.ok) return restore();
        var withVideo = (data.participants || []).filter(function (p) { return p.has_video; });
        if (!withVideo.length) {
          if (labelEl) labelEl.textContent = "No videos";
          setTimeout(restore, 1500);
          return;
        }
        if (btn) btn.disabled = true;
        if (labelEl) labelEl.textContent = "Queued " + withVideo.length + "…";
        return Promise.all(withVideo.map(function (p) {
          return apiPost("../screenspace/api/tasks", { type: "boundary", participant: p.id })
            .catch(function () { return null; });
        })).then(restore);
      })
      .catch(restore);
  }

  // Apply the events slice from a Screenspace intake poll. Dirty-checks against
  // the last raw payload; returns true when it changed (and re-rendered), false
  // when nothing changed (so the poller can back off).
  function applyIntakeEvents(events) {
    var raw = JSON.stringify(events);
    if (raw === state._intakeEventsPollRaw) return false;
    state._intakeEventsPollRaw = raw;
    var hasNew = false;
    events.forEach(function (ev) {
      if (!state.intakeSeenIds[ev.id]) {
        state.intakeSeenIds[ev.id] = "new";
        hasNew = true;
      }
    });
    state.intakeEvents = events;
    var threshold = parseInt((qs("#intakeClusterThreshold") || {}).value, 10) || 10;
    state.intakeClusters = clusterIntakeEvents(intakeClusterSource(), threshold);
    renderIntake(hasNew);
    checkConvergenceTabVisibility();
    refreshMetadataIfActive();
    return true;
  }

  function highlightIntakeCard(idx) {
    // Scope to the Screenspace intake panel: transcript cards also carry
    // .intake-queue-card, so an unscoped query would index across both panels.
    var cards = qsa("#intakeCards .intake-queue-card");
    for (var i = 0; i < cards.length; i++) {
      if (i === idx) {
        cards[i].classList.add("intake-highlight");
      } else {
        cards[i].classList.remove("intake-highlight");
      }
    }
  }

  function buildIntakeDetectorPills() {
    var container = qs("#intakeDetectorPills");
    if (!container) return;
    var counts = {};
    for (var i = 0; i < state.intakeClusters.length; i++) {
      var d = state.intakeClusters[i].detector;
      if (d) counts[d] = (counts[d] || 0) + 1;
    }
    container.innerHTML = "";
    INTAKE_DETECTORS.forEach(function (det) {
      if (!counts[det]) return;
      var chip = ClipgenPrimitives.createFilterChip({
        label: det,
        active: state.intakeFilterDetector === det,
        count: counts[det],
        hue: categoryHue(det),
        // Pin detector chips to the canonical `--color-task-*` token so the
        // dot matches Screenspace's workflow tab + result row exactly.
        color: detectorColor(det),
        onClick: function () {
          state.intakeFilterDetector = state.intakeFilterDetector === det ? "" : det;
          renderIntake(false);
        },
      });
      container.appendChild(chip);
    });
  }

  // Shared participant-pill builder for both intake panels. The cluster source,
  // selected-participant list, container, and re-render are all per-type (cfg).
  function buildIntakeParticipantPills(cfg) {
    var container = qs(cfg.participantsSel);
    if (!container) return;
    var seen = {};
    var participants = [];
    var clusters = state[cfg.clustersKey];
    for (var i = 0; i < clusters.length; i++) {
      var p = clusters[i].participant;
      if (p && !seen[p]) { seen[p] = true; participants.push(p); }
    }
    participants.sort();
    state[cfg.filterParticipantsKey] = state[cfg.filterParticipantsKey].filter(function (p) { return seen[p]; });
    container.innerHTML = "";
    participants.forEach(function (p) {
      var pill = ClipgenPrimitives.createParticipantPill({
        id: p,
        active: state[cfg.filterParticipantsKey].indexOf(p) !== -1,
        onClick: function () {
          var idx = state[cfg.filterParticipantsKey].indexOf(p);
          if (idx === -1) state[cfg.filterParticipantsKey].push(p);
          else state[cfg.filterParticipantsKey].splice(idx, 1);
          cfg.rerender();
        },
      });
      container.appendChild(pill);
    });
  }

  // Density-timeline component refs for both intake panels. Declared together
  // (before the SS_INTAKE/TR_INTAKE configs) so the get/set hooks close over a
  // stable binding; the shared builder/handlers always read the live value at
  // call-time so a rebuild can never leave a stale node behind .setHovered().
  var _intakeDensityEl = null;
  var _trIntakeDensityEl = null;

  // Shared density timeline for both intake panels. Per-type bits (timeline
  // host, bar count/color, hover state, highlight, toggle, density-el get/set)
  // come from the SS_INTAKE / TR_INTAKE config.
  function buildIntakeDensityTimeline(cfg, filtered) {
    var host = qs(cfg.timelineSel);
    if (!host) return;
    host.innerHTML = "";
    cfg.setDensityEl(null);
    if (!filtered.length) return;
    var maxEnd = 0;
    for (var i = 0; i < filtered.length; i++) {
      if (filtered[i].end > maxEnd) maxEnd = filtered[i].end;
    }
    var duration = Math.max(maxEnd * 1.05, 60);
    var events = filtered.map(function (c) {
      var event = {
        t: duration > 0 ? c.start / duration : 0,
        tEnd: duration > 0 ? c.end / duration : 0,
        count: cfg.barCount(c),
      };
      var color = cfg.barColor(c);
      for (var key in color) {
        if (Object.prototype.hasOwnProperty.call(color, key)) event[key] = color[key];
      }
      return event;
    });
    var dt = ClipgenPrimitives.createDensityTimeline({
      events: events,
      durationSec: duration,
      tickCount: 6,
      onBarMouseEnter: function (idx) {
        state[cfg.hoveredIdxKey] = idx;
        cfg.highlightCard(idx);
        var densityEl = cfg.getDensityEl();
        if (densityEl) densityEl.setHovered(idx);
      },
      onBarMouseLeave: function () {
        state[cfg.hoveredIdxKey] = -1;
        cfg.highlightCard(-1);
        var densityEl = cfg.getDensityEl();
        if (densityEl) densityEl.setHovered(-1);
      },
      onBarClick: function (idx, ev) {
        var cluster = cfg.filtered()[idx];
        if (!cluster) return;
        if (ev && ev.shiftKey) cfg.toggleReel(cluster);
        else cfg.toggleArtifacts(cluster);
      },
    });
    cfg.setDensityEl(dt);
    host.appendChild(dt);
  }

  // Shared intake card rendering. Screenspace and Transcript intake build the
  // same card skeleton (thumb + xref badges + participant/type/time meta); the
  // per-type bits are supplied via SS_INTAKE / TR_INTAKE config objects.
  var SS_INTAKE = {
    cardsSel: "#intakeCards",
    cardSel: ".intake-queue-card",
    cardClass: "",
    idxAttr: "intakeIdx",
    selfSource: "screenspace",
    setTranscriptContext: true,
    snippet: null,
    clusterToItem: screenspaceClusterToItem,
    cardHue: function (c) {
      return detectorColor(c.detector) || categoryColor(c.event_type || c.detector || "uncategorized");
    },
    typeText: function (c) {
      return c.event_type || c.detector || "intake";
    },
    selfBadge: function (c) {
      return {
        icon: XREF_BADGES.screenspace.icon,
        color: XREF_BADGES.screenspace.color,
        title: c.event_type || "Screenspace",
      };
    },
    // --- participant pills + density timeline ---
    participantsSel: "#intakeFilterParticipants",
    timelineSel: "#intakeTimeline",
    clustersKey: "intakeClusters",
    filterParticipantsKey: "intakeFilterParticipants",
    hoveredIdxKey: "intakeHoveredIdx",
    filterTextKey: "intakeFilterText",
    filtered: filteredIntakeClusters,
    highlightCard: highlightIntakeCard,
    rerender: function () { renderIntake(false); },
    getDensityEl: function () { return _intakeDensityEl; },
    setDensityEl: function (dt) { _intakeDensityEl = dt; },
    barCount: function (c) { return c.events ? c.events.length : 1; },
    barColor: function (c) {
      // Density bars match Screenspace's `--color-task-*` exactly when the
      // cluster's detector is a known type; falls back to the hue-based oklch
      // path otherwise.
      return { hue: categoryHue(c.detector), color: detectorColor(c.detector) };
    },
    // --- init / delegated listeners ---
    addAllBtnSel: "#intakeAddAllBtn",
    reelAllBtnSel: "#intakeReelAllBtn",
    thresholdSel: "#intakeClusterThreshold",
    searchSel: "#intakeFilterSearch",
    toggleArtifacts: intakeToggleArtifacts,
    toggleReel: intakeToggleReel,
    addToArtifacts: intakeAddToArtifacts,
    addToReel: intakeAddToReel,
    onDismiss: intakeDismissCluster,
    onThresholdChange: function (threshold) {
      state.intakeClusters = clusterIntakeEvents(intakeClusterSource(), threshold);
      renderIntake(false);
    },
    onCardHover: function (card) {
      return card.dataset.transcriptContext || "";
    },
    extraControl: function (cfg) {
      var newToggle = qs("#intakeFilterNew");
      if (newToggle) {
        newToggle.addEventListener("change", function () {
          state.intakeFilterNew = this.checked;
          cfg.rerender();
        });
      }
      var navToggle = qs("#intakeShowNavigational");
      if (navToggle) {
        navToggle.addEventListener("change", function () {
          state.intakeShowNavigational = this.checked;
          // Navigational events change the clustering input, so re-cluster
          // rather than just re-render the existing clusters.
          reclusterIntake();
        });
      }
      var detectBtn = qs("#intakeDetectBoundariesBtn");
      if (detectBtn) {
        detectBtn.addEventListener("click", function () {
          intakeDetectBoundaries(detectBtn);
        });
      }
    },
  };
  var TR_INTAKE = {
    cardsSel: "#trIntakeCards",
    cardSel: ".tr-intake-queue-card",
    cardClass: "tr-intake-queue-card",
    idxAttr: "trIntakeIdx",
    selfSource: "transcript",
    setTranscriptContext: false,
    snippet: function (c) {
      return c.label || c.text || "";
    },
    clusterToItem: transcriptClusterToItem,
    cardHue: function (c) {
      return categoryColor(c.category || "bookmark");
    },
    typeText: function (c) {
      return (TR_INTAKE_CATEGORIES[c.category || "bookmark"] || TR_INTAKE_CATEGORIES.bookmark).label;
    },
    selfBadge: function (c) {
      return {
        icon: XREF_BADGES.transcript.icon,
        color: XREF_BADGES.transcript.color,
        title: c.label || c.category || "Transcript",
      };
    },
    // --- participant pills + density timeline ---
    participantsSel: "#trIntakeFilterParticipants",
    timelineSel: "#trIntakeTimeline",
    clustersKey: "trIntakeClusters",
    filterParticipantsKey: "trIntakeFilterParticipants",
    hoveredIdxKey: "trIntakeHoveredIdx",
    filterTextKey: "trIntakeFilterText",
    filtered: filteredTranscriptIntakeClusters,
    highlightCard: highlightTrIntakeCard,
    rerender: renderTranscriptIntake,
    getDensityEl: function () { return _trIntakeDensityEl; },
    setDensityEl: function (dt) { _trIntakeDensityEl = dt; },
    barCount: function (c) { return c.marks ? c.marks.length : 1; },
    barColor: function (c) {
      return { hue: categoryHue(c.category || "bookmark") };
    },
    // --- init / delegated listeners ---
    addAllBtnSel: "#trIntakeAddAllBtn",
    reelAllBtnSel: "#trIntakeReelAllBtn",
    thresholdSel: "#trIntakeClusterThreshold",
    searchSel: "#trIntakeFilterSearch",
    toggleArtifacts: trIntakeToggleArtifacts,
    toggleReel: trIntakeToggleReel,
    addToArtifacts: trIntakeAddToArtifacts,
    addToReel: trIntakeAddToReel,
    onDismiss: trIntakeDismissCluster,
    onThresholdChange: function () {
      // TR re-polls (the poll re-reads the threshold input itself); ignores arg.
      refreshTranscriptIntake();
    },
    onCardHover: function (card, idx) {
      var cluster = filteredTranscriptIntakeClusters()[idx];
      return cluster ? (cluster.text || cluster.label || "") : "";
    },
    extraControl: function (_cfg) {
      var showAllToggle = qs("#trIntakeShowAll");
      if (!showAllToggle) return;
      showAllToggle.addEventListener("change", function () {
        state.trIntakeShowAll = this.checked;
        state._trIntakeMarksPollFp = null;
        refreshTranscriptIntake();
      });
    },
  };

  function renderIntakeCards(cfg, items) {
    var container = qs(cfg.cardsSel);
    container.innerHTML = "";
    items.forEach(function (c, idx) {
      var card = el("div", "queue-card intake-queue-card" + (cfg.cardClass ? " " + cfg.cardClass : ""));
      card.style.setProperty("--cg-card-hue", cfg.cardHue(c));
      card.dataset[cfg.idxAttr] = idx;
      card.setAttribute("draggable", "true");
      card.addEventListener("dragstart", function (ev) {
        ev.dataTransfer.setData("application/json", JSON.stringify(cfg.clusterToItem(c)));
        ev.dataTransfer.effectAllowed = "copyMove";
        setCardDragImage(ev, this);
      });

      var thumb = buildQueueCardThumb(card, {
        participant: c.participant,
        start: c.start,
        duration: Math.max(0, c.end - c.start),
        observe: true,
      });

      var xref = findOverlappingData(c.participant, c.start, c.end);
      var badgeStack = buildXrefBadges(xref, cfg.selfSource, cfg.selfBadge(c));
      if (badgeStack) thumb.appendChild(badgeStack);

      var meta = el("div", "queue-card-meta");
      var row = el("div", "queue-card-meta-row");
      row.appendChild(el("span", "queue-card-participant", c.participant));
      row.appendChild(el("span", "queue-card-type", cfg.typeText(c)));
      meta.appendChild(row);
      // Navigational boundaries are points (unpadded) — show a single time
      // rather than a "0:12–0:12" range.
      var timeText = (c.navigational && c.start === c.end)
        ? formatDuration(c.start)
        : formatDuration(c.start) + "–" + formatDuration(c.end);
      meta.appendChild(el("span", "queue-card-time", timeText));
      if (cfg.snippet) {
        var snippet = cfg.snippet(c);
        if (snippet) {
          var textEl = el("span", "queue-card-text", snippet);
          textEl.title = snippet;
          meta.appendChild(textEl);
        }
      }
      card.appendChild(meta);

      if (cfg.setTranscriptContext && xref.transcriptSnippets.length > 0) {
        card.dataset.transcriptContext = xref.transcriptSnippets
          .map(function (s) {
            return s.text;
          })
          .join("\n");
      }

      container.appendChild(card);
    });
    attachQueueScrubbers(container);
    refreshIntakeCardStates();
  }

  // ---- Screenspace intake: render cards, filters, and density timeline ----

  function renderIntake(_hasNew) {
    ssClearPending();
    var container = qs("#intakeCards");
    var addAllBtn = qs("#intakeAddAllBtn");
    var reelAllBtn = qs("#intakeReelAllBtn");
    var tabBadge = qs("#intakeTabBadge");

    buildIntakeDetectorPills();
    buildIntakeParticipantPills(SS_INTAKE);

    if (!state.intakeClusters.length) {
      if (tabBadge) tabBadge.classList.add("hidden");
      container.innerHTML = "";
      container.appendChild(el("div", "drop-target-empty", "Screenspace events will appear here"));
      addAllBtn.disabled = true;
      reelAllBtn.disabled = true;
      buildIntakeDensityTimeline(SS_INTAKE, []);
      return;
    }
    if (tabBadge) {
      tabBadge.textContent = state.intakeClusters.length;
      tabBadge.classList.remove("hidden");
    }
    var clusters = filteredIntakeClusters();
    addAllBtn.disabled = clusters.length === 0;
    reelAllBtn.disabled = clusters.length === 0;

    buildIntakeDensityTimeline(SS_INTAKE, clusters);

    if (clusters.length === 0) {
      container.innerHTML = "";
      container.appendChild(el("div", "drop-target-empty", "No events match the current filters"));
      return;
    }
    renderIntakeCards(SS_INTAKE, clusters);
  }

  function screenspaceClusterToItem(cluster) {
    var start = cluster.start;
    var end = cluster.end;
    // Navigational boundaries are points (no intake padding), so a clip would be
    // zero-length. Give one a forward default-duration window so "Add to
    // Artifacts/Reel" still produces a real clip starting at the boundary.
    if (cluster.navigational && start === end) {
      end = start + (CLIPGEN_CONFIG.defaultDuration || 5);
    }
    return {
      participant: cluster.participant,
      start: start,
      end: end,
      desc: cluster.event_type,
      source: "screenspace",
      event_type: cluster.event_type,
      event_ids: cluster.events.map(function (e) { return e.id; }),
    };
  }

  function intakeAddToArtifacts(cluster) {
    intakeAddItem(state.artifactQueue, screenspaceClusterToItem(cluster), renderArtifactQueue);
  }

  function intakeToggleArtifacts(cluster) {
    intakeToggleItem(state.artifactQueue, screenspaceClusterToItem(cluster), renderArtifactQueue);
  }

  function intakeDismissCluster(cluster) {
    var ids = cluster.events.map(function (e) { return e.id; });
    apiPut("../screenspace/api/events/bulk-exclude", { ids: ids })
      .then(function () { refreshScreenspaceIntake(); })
      .catch(function () {});
  }

  function intakeAddToReel(cluster) {
    intakeAddItem(state.reelQueue, screenspaceClusterToItem(cluster), renderReelQueue);
  }

  function intakeToggleReel(cluster) {
    intakeToggleItem(state.reelQueue, screenspaceClusterToItem(cluster), renderReelQueue);
  }

  // Mark intake cards whose cluster is in either queue, mirroring how
  // updateCellClasses highlights queued spreadsheet cells. Driven by the render
  // queue functions (so every mutation re-syncs) and the intake render functions
  // (so the highlight survives the poll that rebuilds cards).
  function refreshIntakeCardStates() {
    var ssClusters = filteredIntakeClusters();
    qsa("#intakeCards .intake-queue-card").forEach(function (card) {
      var c = ssClusters[parseInt(card.dataset.intakeIdx, 10)];
      if (!c) return;
      var item = screenspaceClusterToItem(c);
      card.classList.toggle(
        "in-queue",
        findIntakeInQueue(state.artifactQueue, item) >= 0 ||
          findIntakeInQueue(state.reelQueue, item) >= 0,
      );
    });
    var trClusters = filteredTranscriptIntakeClusters();
    qsa("#trIntakeCards .tr-intake-queue-card").forEach(function (card) {
      var c = trClusters[parseInt(card.dataset.trIntakeIdx, 10)];
      if (!c) return;
      var item = transcriptClusterToItem(c);
      card.classList.toggle(
        "in-queue",
        findIntakeInQueue(state.artifactQueue, item) >= 0 ||
          findIntakeInQueue(state.reelQueue, item) >= 0,
      );
    });
  }

  function filteredIntakeClusters() {
    var clusters = state.intakeClusters;
    var text = state.intakeFilterText.toLowerCase();
    var det = state.intakeFilterDetector;
    var onlyNew = state.intakeFilterNew;
    var parts = state.intakeFilterParticipants;
    if (!text && !det && !onlyNew && !parts.length) return clusters;
    return clusters.filter(function (c) {
      if (parts.length && parts.indexOf(c.participant) === -1) return false;
      if (det && c.detector !== det) return false;
      if (text && (c.event_type || "").toLowerCase().indexOf(text) === -1
          && (c.region || "").toLowerCase().indexOf(text) === -1
          && (c.participant || "").toLowerCase().indexOf(text) === -1) return false;
      if (onlyNew) {
        var hasNew = false;
        for (var i = 0; i < c.events.length; i++) {
          if (state.intakeSeenIds[c.events[i].id] === "new") { hasNew = true; break; }
        }
        if (!hasNew) return false;
      }
      return true;
    });
  }

  // Shared init for both intake panels. Attaches the delegated card listeners
  // (click/contextmenu/mouseover/mouseleave) ONCE on the cards container, plus
  // the Add-All/Reel-All buttons, threshold, search, and the per-type extra
  // control. Must be called exactly once per panel at startup — never from a
  // render path (CODE-REVIEW.md listener-cleanup rule). Per-type behavior is
  // supplied via cfg; the shared `#trIntakeTooltip` host and the
  // `trIntakeTooltipsEnabled` gate are owned here so both panels behave alike.
  function initIntakePanel(cfg) {
    var cards = qs(cfg.cardsSel);
    if (!cards) return;

    // Click: normal = Artifacts, shift = Reel
    cards.addEventListener("click", function (e) {
      var card = e.target.closest(cfg.cardSel);
      if (!card) return;
      var cluster = cfg.filtered()[parseInt(card.dataset[cfg.idxAttr], 10)];
      if (!cluster) return;
      if (e.shiftKey) cfg.toggleReel(cluster);
      else cfg.toggleArtifacts(cluster);
    });

    // Right-click to dismiss
    cards.addEventListener("contextmenu", function (e) {
      var card = e.target.closest(cfg.cardSel);
      if (!card) return;
      e.preventDefault();
      var cluster = cfg.filtered()[parseInt(card.dataset[cfg.idxAttr], 10)];
      if (cluster) cfg.onDismiss(cluster);
    });

    // Card hover → highlight + tooltip + timeline marker
    cards.addEventListener("mouseover", function (e) {
      var card = e.target.closest(cfg.cardSel);
      if (!card) return;
      var idx = parseInt(card.dataset[cfg.idxAttr], 10);
      if (state[cfg.hoveredIdxKey] !== idx) {
        state[cfg.hoveredIdxKey] = idx;
        cfg.highlightCard(idx);
        var densityEl = cfg.getDensityEl();
        if (densityEl) densityEl.setHovered(idx);
      }
      var tooltip = qs("#trIntakeTooltip");
      if (tooltip && state.trIntakeTooltipsEnabled) {
        var tooltipText = cfg.onCardHover(card, idx) || "";
        if (tooltipText) {
          tooltip.textContent = tooltipText;
          tooltip.classList.remove("hidden");
          positionTooltipAnchored(tooltip, card.getBoundingClientRect());
        } else {
          tooltip.classList.add("hidden");
        }
      }
    });
    cards.addEventListener("mouseleave", function () {
      if (state[cfg.hoveredIdxKey] !== -1) {
        state[cfg.hoveredIdxKey] = -1;
        cfg.highlightCard(-1);
        var densityEl = cfg.getDensityEl();
        if (densityEl) densityEl.setHovered(-1);
      }
      var tooltip = qs("#trIntakeTooltip");
      if (tooltip) tooltip.classList.add("hidden");
    });

    var addAllBtn = qs(cfg.addAllBtnSel);
    if (addAllBtn) {
      addAllBtn.addEventListener("click", function () {
        cfg.filtered().forEach(function (c) { cfg.addToArtifacts(c); });
      });
    }
    var reelAllBtn = qs(cfg.reelAllBtnSel);
    if (reelAllBtn) {
      reelAllBtn.addEventListener("click", function () {
        cfg.filtered().forEach(function (c) { cfg.addToReel(c); });
      });
    }
    var thresholdInput = qs(cfg.thresholdSel);
    if (thresholdInput) {
      thresholdInput.addEventListener("change", function () {
        cfg.onThresholdChange(parseInt(this.value, 10) || 5);
      });
    }

    // Text search filter (debounced)
    var searchEl = qs(cfg.searchSel);
    var _searchTimer = 0;
    if (searchEl) {
      searchEl.addEventListener("input", function () {
        state[cfg.filterTextKey] = this.value;
        clearTimeout(_searchTimer);
        _searchTimer = setTimeout(function () { cfg.rerender(); }, 250);
      });
    }

    // Per-type extra control (SS "New only" / TR "Show all")
    cfg.extraControl(cfg);
  }

  // ---- Transcript Intake ----

  // Colors resolve to CSS tokens (see tokens.css `--cat-*`) so dark mode tracks the theme.
  var TR_INTAKE_CATEGORIES = {
    pain_point: { label: "Pain Point", token: "--cat-pain-point" },
    delight:    { label: "Delight",    token: "--cat-delight" },
    quote:      { label: "Quote",      token: "--cat-quote" },
    insight:    { label: "Insight",    token: "--cat-insight" },
    task:       { label: "Task Issue", token: "--cat-task" },
    bookmark:   { label: "Bookmark",   token: "--cat-bookmark" },
  };

  function trIntakeCategoryColor(key) {
    var entry = TR_INTAKE_CATEGORIES[key] || TR_INTAKE_CATEGORIES.bookmark;
    return getCSSVar(entry.token, "");
  }

  function _syncMarkCategoriesFromSettings(settings) {
    if (!settings) return;
    for (var i = 0; i < settings.length; i++) {
      if (settings[i].name === "MARK_CATEGORIES" && settings[i].value) {
        setMarkCategories(settings[i].value);
        return;
      }
    }
  }

  // Apply the marks block ({marks, categories}) from a transcript intake poll.
  // In default mode a fingerprint dirty-check skips redundant re-renders and
  // returns false so the poller can back off; returns true when it re-rendered.
  // In "Show all" mode it always refreshes (via applyTranscriptShowAll, its own
  // participants+transcripts fan-out) and returns false — backoff there is
  // governed by the status flags, so the heavy fan-out slows when nothing runs.
  function applyTranscriptMarks(block) {
    var marks = block.marks || [];
    var categories = block.categories;
    var threshold = parseInt((qs("#trIntakeClusterThreshold") || {}).value, 10) || 10;
    if (!state.trIntakeShowAll) {
      var fp =
        String(threshold) +
        "\0" +
        (categories ? JSON.stringify(categories) : "") +
        "\0" +
        JSON.stringify(marks);
      if (fp === state._trIntakeMarksPollFp) return false;
      state._trIntakeMarksPollFp = fp;
    }
    if (categories) setMarkCategories(categories);
    state.trIntakeMarks = marks.filter(function (m) { return m.valid; });
    state.trIntakeClusters = clusterTranscriptMarks(state.trIntakeMarks, threshold);
    if (state.trIntakeShowAll) {
      applyTranscriptShowAll(threshold);
      return false;
    }
    renderTranscriptIntake();
    checkConvergenceTabVisibility();
    refreshMetadataIfActive();
    return true;
  }

  // "Show all": append every unmarked segment as a queue item. Its own fan-out
  // (participants + one transcript fetch each) — user-toggled, not always-on.
  function applyTranscriptShowAll(threshold) {
    apiGet("../transcripts/api/participants")
      .then(function (pData) {
        if (!pData.ok) return;
        var transcribed = pData.participants.filter(function (p) { return p.has_transcript; });
        var promises = transcribed.map(function (p) {
          return apiGet("../transcripts/api/transcript/" + p.id);
        });
        Promise.all(promises).then(function (results) {
          var markedIds = {};
          for (var i = 0; i < state.trIntakeMarks.length; i++) markedIds[state.trIntakeMarks[i].segment_id] = true;
          var allItems = state.trIntakeMarks.slice();
          for (var j = 0; j < results.length; j++) {
            if (!results[j].ok) continue;
            var pid = results[j].participant;
            var segs = results[j].segments;
            for (var k = 0; k < segs.length; k++) {
              if (!markedIds[segs[k].id]) {
                allItems.push({
                  id: null,
                  segment_id: segs[k].id,
                  category: null,
                  label: null,
                  valid: true,
                  participant: pid,
                  start: segs[k].start,
                  end: segs[k].end,
                  text: segs[k].text,
                });
              }
            }
          }
          state.trIntakeClusters = clusterTranscriptMarks(allItems, threshold);
          renderTranscriptIntake();
          checkConvergenceTabVisibility();
          refreshMetadataIfActive();
        });
      })
      .catch(function () {});
  }

  function filteredTranscriptIntakeClusters() {
    var clusters = state.trIntakeClusters;
    var cat = state.trIntakeFilterCategory;
    var parts = state.trIntakeFilterParticipants;
    var text = state.trIntakeFilterText.toLowerCase();
    if (!cat && !parts.length && !text) return clusters;
    return clusters.filter(function (c) {
      if (parts.length && parts.indexOf(c.participant) === -1) return false;
      if (cat && c.category !== cat) return false;
      if (text && (c.text || "").toLowerCase().indexOf(text) === -1
          && (c.label || "").toLowerCase().indexOf(text) === -1
          && (c.participant || "").toLowerCase().indexOf(text) === -1) return false;
      return true;
    });
  }

  function renderTranscriptIntake() {
    ssClearPending();
    var filtered = filteredTranscriptIntakeClusters();
    var container = qs("#trIntakeCards");
    var addAllBtn = qs("#trIntakeAddAllBtn");
    var reelAllBtn = qs("#trIntakeReelAllBtn");
    var badge = qs("#trIntakeTabBadge");

    if (badge) {
      if (state.trIntakeMarks.length > 0) {
        badge.textContent = state.trIntakeMarks.length;
        badge.classList.remove("hidden");
      } else {
        badge.classList.add("hidden");
      }
    }

    if (addAllBtn) addAllBtn.disabled = filtered.length === 0;
    if (reelAllBtn) reelAllBtn.disabled = filtered.length === 0;

    buildTrIntakeCategoryPills();
    buildIntakeParticipantPills(TR_INTAKE);
    buildIntakeDensityTimeline(TR_INTAKE, filtered);

    if (filtered.length === 0) {
      container.innerHTML = '<div class="drop-target-empty">Transcript marks will appear here</div>';
      return;
    }

    renderIntakeCards(TR_INTAKE, filtered);
  }

  function buildTrIntakeCategoryPills() {
    var container = qs("#trIntakeCategoryPills");
    if (!container) return;
    var cats = Object.keys(TR_INTAKE_CATEGORIES);
    var counts = {};
    state.trIntakeClusters.forEach(function (c) {
      var k = c.category || "bookmark";
      counts[k] = (counts[k] || 0) + 1;
    });
    container.innerHTML = "";
    cats.forEach(function (key) {
      var cat = TR_INTAKE_CATEGORIES[key];
      var chip = ClipgenPrimitives.createFilterChip({
        label: cat.label,
        active: state.trIntakeFilterCategory === key,
        count: counts[key] || 0,
        hue: categoryHue(key),
        onClick: function () {
          state.trIntakeFilterCategory = state.trIntakeFilterCategory === key ? "" : key;
          renderTranscriptIntake();
        },
      });
      container.appendChild(chip);
    });
  }

  function transcriptClusterToItem(cluster) {
    return {
      participant: cluster.participant,
      start: cluster.start,
      end: cluster.end,
      desc: cluster.category || "transcript",
      source: "transcript",
      mark_ids: cluster.marks.map(function (m) { return m.id; }),
    };
  }

  function trIntakeAddToArtifacts(cluster) {
    intakeAddItem(state.artifactQueue, transcriptClusterToItem(cluster), renderArtifactQueue);
  }

  function trIntakeToggleArtifacts(cluster) {
    intakeToggleItem(state.artifactQueue, transcriptClusterToItem(cluster), renderArtifactQueue);
  }

  function trIntakeAddToReel(cluster) {
    intakeAddItem(state.reelQueue, transcriptClusterToItem(cluster), renderReelQueue);
  }

  function trIntakeToggleReel(cluster) {
    intakeToggleItem(state.reelQueue, transcriptClusterToItem(cluster), renderReelQueue);
  }

  function trIntakeDismissCluster(cluster) {
    var ids = cluster.marks.map(function (m) { return m.id; }).filter(Boolean);
    if (!ids.length) return;
    // DELETE with a JSON body — apiDelete takes no body, so this custom fetch stays.
    fetch("../transcripts/api/marks/" + ids[0], {
      method: "DELETE",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ids: ids }),
    })
      .then(function () { refreshTranscriptIntake(); })
      .catch(function () {});
  }

  function highlightTrIntakeCard(idx) {
    var cards = qsa(".tr-intake-queue-card");
    for (var i = 0; i < cards.length; i++) {
      if (i === idx) {
        cards[i].classList.add("intake-highlight");
      } else {
        cards[i].classList.remove("intake-highlight");
      }
    }
  }

  function initTooltipToggle() {
    state.trIntakeTooltipsEnabled = getStoredTooltipPref();
    var btn = qs("#tooltipToggle");
    if (!btn) return;
    btn.setAttribute("aria-pressed", state.trIntakeTooltipsEnabled ? "true" : "false");
    btn.addEventListener("click", function () {
      state.trIntakeTooltipsEnabled = !state.trIntakeTooltipsEnabled;
      btn.setAttribute("aria-pressed", state.trIntakeTooltipsEnabled ? "true" : "false");
      setStoredTooltipPref(state.trIntakeTooltipsEnabled);
      if (!state.trIntakeTooltipsEnabled) {
        var tt = qs("#trIntakeTooltip");
        if (tt) tt.classList.add("hidden");
      }
    });
  }

  // Init both intake panels. Folded from the two hub-boot initIntakePanel calls
  // so the SS_INTAKE/TR_INTAKE configs never need to leave this file.
  function initIntake() {
    initIntakePanel(SS_INTAKE);
    initIntakePanel(TR_INTAKE);
  }

  // ---- Published to the hub (window.ClipgenStudio) ----
  STUDIO.initIntake = initIntake;
  STUDIO.pollScreenspaceIntake = pollScreenspaceIntake;
  STUDIO.pollTranscriptIntake = pollTranscriptIntake;
  STUDIO.initTooltipToggle = initTooltipToggle;
  STUDIO.refreshIntakeCardStates = refreshIntakeCardStates;
  STUDIO.renderIntake = renderIntake;
  STUDIO._syncMarkCategoriesFromSettings = _syncMarkCategoriesFromSettings;
})();
