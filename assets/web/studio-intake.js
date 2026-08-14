/* clipgen Studio — intake satellite.
 *
 * Panels: Screenspace, Transcript, Composer and MindNode intake.
 *
 * Carved out of studio.js (the hub) following the hub+satellite convention used
 * by transcripts.js / screenspace.js. Shares the hub's mutable `state` and the
 * helper functions it needs through the `window.ClipgenStudio` (STUDIO)
 * namespace; the hub calls back in via same-named guarded delegators. Loaded by
 * studio.html immediately after studio.js.
 *
 * Kept in the hub (NOT moved here; reached via STUDIO): the `ss*Thumb*`
 * lazy-thumbnail cluster (`ssClearPending`), which the hub's queue cards
 * share. `buildXrefBadges` is a utils.js ambient global (shared with the
 * Overview page).
 */

(function () {
  "use strict";

  var STUDIO = window.ClipgenStudio;
  var state = STUDIO.state;
  var attachQueueScrubbers = STUDIO.attachQueueScrubbers,
    buildQueueCardThumb = STUDIO.buildQueueCardThumb,
    buildXrefBadges = STUDIO.buildXrefBadges,
    findIntakeInQueue = STUDIO.findIntakeInQueue,
    findOverlappingData = STUDIO.findOverlappingData,
    intakeAddItem = STUDIO.intakeAddItem,
    intakeToggleItem = STUDIO.intakeToggleItem,
    renderArtifactQueue = STUDIO.renderArtifactQueue,
    renderReelQueue = STUDIO.renderReelQueue,
    setCardDragImage = STUDIO.setCardDragImage,
    ssClearPending = STUDIO.ssClearPending;
  // Pure intake clustering (aliased the same way in the hub at load).
  var clusterIntakeEvents = window.ClipgenIntakeCluster.clusterIntakeEvents;
  var clusterTranscriptMarks = window.ClipgenIntakeCluster.clusterTranscriptMarks;

  // ---- Screenspace Intake ----

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
    // Echo the last-seen events version so the server can skip the events
    // payload (deep-copy + sanitize + ship) when nothing changed this tick.
    var url = "../screenspace/api/intake-poll?excluded=false";
    if (state._intakeEventsVersion != null) {
      url += "&events_version=" + state._intakeEventsVersion;
    }
    return apiGet(url)
      .then(function (data) {
        if (!data || !data.ok) {
          setIntakeDot(false);
          return false;
        }
        var status = data.status || {};
        var dotOn = !!status.running || (status.worker_alive && !!status.queued);
        setIntakeDot(dotOn);
        if (data.events_version != null) {
          state._intakeEventsVersion = data.events_version;
        }
        // events omitted => unchanged since our cursor; keep current view.
        var eventsChanged = data.events_unchanged
          ? false
          : applyIntakeEvents(data.events || []);
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

  // On-demand refresh after a user action (or the subheader Refresh button):
  // wake the adaptive poller so it refetches now AND resets its idle backoff to
  // the fast cadence. Falls back to a direct poll if the poller hasn't been
  // created yet (pre-boot). Each returns a promise settling when the refetch
  // has landed, so the Refresh button can spin until then.
  function refreshScreenspaceIntake() {
    if (state.ssIntakePoller) return state.ssIntakePoller.wake();
    return pollScreenspaceIntake();
  }
  function refreshTranscriptIntake() {
    if (state.trIntakePoller) return state.trIntakePoller.wake();
    return pollTranscriptIntake();
  }
  function refreshComposerIntake() {
    if (state.coIntakePoller) return state.coIntakePoller.wake();
    return pollComposerIntake();
  }

  // Events fed into the intake clustering surface. Navigational (boundary)
  // events are orientation scaffolding, not clip candidates, so they never
  // reach the curation queue. state.intakeEvents still holds ALL events
  // (Metadata's boundary count reads from it).
  function intakeClusterSource() {
    return state.intakeEvents.filter(function (ev) { return !ev.navigational; });
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
    // Composer-trim cross-link: first event in the cluster with a carded trim.
    trimBadgeKey: function (c) {
      var events = c.events || [];
      for (var i = 0; i < events.length; i++) {
        if (state.coTrimCardKeys["screenspace:" + events[i].id]) {
          return "screenspace:" + events[i].id;
        }
      }
      return null;
    },
    extraControl: function () {},
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
    // Composer-trim cross-link: first mark in the cluster with a carded trim.
    trimBadgeKey: function (c) {
      var marks = c.marks || [];
      for (var i = 0; i < marks.length; i++) {
        if (marks[i].id && state.coTrimCardKeys["transcript-mark:" + marks[i].id]) {
          return "transcript-mark:" + marks[i].id;
        }
      }
      return null;
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

      // Asterisk badge: this cluster has a trimmed counterpart in Composer —
      // click jumps to the Composer Intake tab and highlights it.
      if (cfg.trimBadgeKey) {
        var trimKey = cfg.trimBadgeKey(c);
        if (trimKey) {
          // iconHTML lives in the studio.js hub IIFE (not reachable here), so
          // inline the sanctioned mask-icon span directly.
          var trimBadge = el("button", "intake-trim-badge");
          trimBadge.innerHTML = '<span class="cg-icon cg-icon--scissors"></span>';
          trimBadge.type = "button";
          trimBadge.title = "Trimmed in Composer. Click to view the trimmed version";
          trimBadge.setAttribute("aria-label", "Show trimmed version in Composer Intake");
          trimBadge.addEventListener("click", function (ev) {
            ev.stopPropagation();
            focusComposerIntakeItem(trimKey);
          });
          thumb.appendChild(trimBadge);
        }
      }

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
    // Re-renders wipe the hub's keyboard-cursor outline; repaint it.
    if (STUDIO.kbPaintCursor) STUDIO.kbPaintCursor();
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

  // Mark intake cards whose cluster is in either queue, as updateCellClasses does
  // for queued spreadsheet cells. Driven by the render-queue functions (so every
  // mutation re-syncs) and the intake render functions (so the highlight survives
  // the poll that rebuilds cards).
  //
  // Covers *every* intake panel: MindNode and Composer cards carry .in-queue too
  // but were once never visited here, so a queued card showed no "already queued"
  // state and clicking it again silently toggled it out. Driven off each panel's
  // own config, so a new panel cannot be forgotten the same way.
  function intakeCardPanels() {
    return [
      {
        cardsSel: "#intakeCards",
        cardSel: ".intake-queue-card",
        idxAttr: "intakeIdx",
        filtered: filteredIntakeClusters,
        clusterToItem: screenspaceClusterToItem,
      },
      {
        cardsSel: "#trIntakeCards",
        cardSel: ".tr-intake-queue-card",
        idxAttr: "trIntakeIdx",
        filtered: filteredTranscriptIntakeClusters,
        clusterToItem: transcriptClusterToItem,
      },
      CO_INTAKE,
      MN_INTAKE,
    ];
  }

  function refreshIntakeCardStates() {
    intakeCardPanels().forEach(function (panel) {
      var items = panel.filtered();
      qsa(panel.cardsSel + " " + panel.cardSel).forEach(function (card) {
        var c = items[parseInt(card.dataset[panel.idxAttr], 10)];
        if (!c) return;
        var item = panel.clusterToItem(c);
        card.classList.toggle(
          "in-queue",
          findIntakeInQueue(state.artifactQueue, item) >= 0 ||
            findIntakeInQueue(state.reelQueue, item) >= 0,
        );
      });
    });
  }

  function filteredIntakeClusters() {
    var clusters = state.intakeClusters;
    var text = state.intakeFilterText.toLowerCase();
    var det = state.intakeFilterDetector;
    var parts = state.intakeFilterParticipants;
    if (!text && !det && !parts.length) return clusters;
    return clusters.filter(function (c) {
      if (parts.length && parts.indexOf(c.participant) === -1) return false;
      if (det && c.detector !== det) return false;
      if (text && (c.event_type || "").toLowerCase().indexOf(text) === -1
          && (c.region || "").toLowerCase().indexOf(text) === -1
          && (c.participant || "").toLowerCase().indexOf(text) === -1) return false;
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

    // Per-type extra control (TR "Show all"; SS and CO have none)
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
        });
      })
      .catch(function () {});
  }

  function filteredTranscriptIntakeClusters() {
    var clusters = state.trIntakeClusters;
    var cat = state.trIntakeFilterCategory;
    var parts = state.trIntakeFilterParticipants;
    var sevs = state.trIntakeFilterSeverities;
    var text = state.trIntakeFilterText.toLowerCase();
    if (!cat && !parts.length && !sevs.length && !text) return clusters;
    return clusters.filter(function (c) {
      if (parts.length && parts.indexOf(c.participant) === -1) return false;
      if (cat && c.category !== cat) return false;
      if (sevs.length && sevs.indexOf(c.severity) === -1) return false;
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
    buildTrIntakeSeverityPills();
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

  // Severity filter chips — multi-select (mirrors the Sheet-grid severity
  // filter): each chip toggles its label in/out of state.trIntakeFilterSeverities;
  // only severities present in the data render. Cluster severity is the most-
  // severe of its marks (hoisted in clusterTranscriptMarks).
  function buildTrIntakeSeverityPills() {
    var container = qs("#trIntakeSeverityPills");
    if (!container) return;
    var counts = {};
    state.trIntakeClusters.forEach(function (c) {
      if (c.severity) counts[c.severity] = (counts[c.severity] || 0) + 1;
    });
    container.innerHTML = "";
    for (var i = 0; i < CLIPGEN_CONFIG.severity.length; i++) {
      (function (label) {
        if (!counts[label]) return;
        var chip = ClipgenPrimitives.createFilterChip({
          label: label,
          active: state.trIntakeFilterSeverities.indexOf(label) >= 0,
          count: counts[label],
          color: "var(--" + severityClass(label) + ")",
          onClick: function () {
            var arr = state.trIntakeFilterSeverities.slice();
            var idx = arr.indexOf(label);
            if (idx >= 0) arr.splice(idx, 1);
            else arr.push(label);
            state.trIntakeFilterSeverities = arr;
            renderTranscriptIntake();
          },
        });
        container.appendChild(chip);
      })(CLIPGEN_CONFIG.severity[i].label);
    }
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

  // ---- Composer Intake ----
  //
  // The simplest of the three panels: Composer cuts and trims are already
  // curated in/out pairs (no clustering, no dismiss — both are managed on
  // the Composer page), so this is a fingerprint-polled list of cards
  // sharing the generic panel machinery via a CO_INTAKE config. Trimmed
  // source markers appear alongside cuts (item.isTrim), and the SS/TR
  // panels' asterisk badges deep-link here via focusComposerIntakeItem.

  var _coIntakeDensityEl = null;

  function composerCutToItem(cut) {
    var name = cut.label || (cut.isTrim ? "trimmed" : "composer");
    return {
      participant: cut.participant,
      start: cut.start,
      end: cut.end,
      desc: name,
      source: "composer",
      event_type: name,
      event_ids: [cut.id],
    };
  }

  function coIntakeAddToArtifacts(cut) {
    intakeAddItem(state.artifactQueue, composerCutToItem(cut), renderArtifactQueue);
  }

  function coIntakeToggleArtifacts(cut) {
    intakeToggleItem(state.artifactQueue, composerCutToItem(cut), renderArtifactQueue);
  }

  function coIntakeAddToReel(cut) {
    intakeAddItem(state.reelQueue, composerCutToItem(cut), renderReelQueue);
  }

  function coIntakeToggleReel(cut) {
    intakeToggleItem(state.reelQueue, composerCutToItem(cut), renderReelQueue);
  }

  // The four kinds of item the Composer feeds into intake: its own in/out cut
  // pairs, plus the non-destructive trims of each read-only marker lane. Order
  // here is the pill order.
  var CO_INTAKE_TYPES = [
    { key: "cut", label: "Cuts", color: "var(--color-accent)" },
    { key: "transcript", label: "Transcript edit", color: "var(--stream-transcript)" },
    { key: "screenspace", label: "Screenspace edit", color: "var(--stream-screenspace)" },
    { key: "sheet", label: "Sheet edit", color: "var(--stream-sheet)" },
  ];

  // Which lane a trim came from. Named cutType, not source: the queue items
  // composerCutToItem() builds already carry source "composer". The key prefix
  // is load-bearing rather than the stored `source`, because trims written
  // before trims carried metadata have source "".
  function composerTrimType(key, stored) {
    if (key.indexOf("sheet:") === 0) return "sheet";
    if (key.indexOf("screenspace:") === 0) return "screenspace";
    if (key.indexOf("transcript-mark:") === 0) return "transcript";
    return stored || "";
  }

  function filteredComposerIntakeCuts() {
    var items = state.coIntakeItems;
    var parts = state.coIntakeFilterParticipants;
    var types = state.coIntakeFilterTypes;
    var text = state.coIntakeFilterText.toLowerCase();
    if (!parts.length && !types.length && !text) return items;
    return items.filter(function (c) {
      if (parts.length && parts.indexOf(c.participant) === -1) return false;
      if (types.length && types.indexOf(c.cutType) === -1) return false;
      if (text && (c.label || "").toLowerCase().indexOf(text) === -1
          && (c.participant || "").toLowerCase().indexOf(text) === -1) return false;
      return true;
    });
  }

  // Type filter chips — multi-select, and all four always render (with a 0
  // count) so the row reads as a legend even on an empty panel.
  function buildCoIntakeTypePills() {
    var container = qs("#coIntakeTypePills");
    if (!container) return;
    var counts = {};
    state.coIntakeItems.forEach(function (c) {
      if (c.cutType) counts[c.cutType] = (counts[c.cutType] || 0) + 1;
    });
    container.innerHTML = "";
    CO_INTAKE_TYPES.forEach(function (type) {
      var chip = ClipgenPrimitives.createFilterChip({
        label: type.label,
        active: state.coIntakeFilterTypes.indexOf(type.key) >= 0,
        count: counts[type.key] || 0,
        color: type.color,
        onClick: function () {
          var arr = state.coIntakeFilterTypes.slice();
          var idx = arr.indexOf(type.key);
          if (idx >= 0) arr.splice(idx, 1);
          else arr.push(type.key);
          state.coIntakeFilterTypes = arr;
          renderComposerIntake();
        },
      });
      container.appendChild(chip);
    });
  }

  function highlightCoIntakeCard(idx) {
    var cards = qsa(".co-intake-queue-card");
    for (var i = 0; i < cards.length; i++) {
      if (i === idx) cards[i].classList.add("intake-highlight");
      else cards[i].classList.remove("intake-highlight");
    }
  }

  var CO_INTAKE = {
    cardsSel: "#coIntakeCards",
    cardSel: ".co-intake-queue-card",
    cardClass: "co-intake-queue-card",
    idxAttr: "coIntakeIdx",
    selfSource: "composer",
    setTranscriptContext: false,
    snippet: function (c) { return c.label || ""; },
    clusterToItem: composerCutToItem,
    cardHue: function (c) {
      return categoryColor(c.label || "composer");
    },
    typeText: function (c) {
      if (c.isTrim) return (c.label || "marker") + " (trimmed)";
      return c.label || "composer cut";
    },
    selfBadge: function (c) {
      return {
        icon: XREF_BADGES.composer.icon,
        color: XREF_BADGES.composer.color,
        title: c.label || "Composer cut",
      };
    },
    // --- participant pills + density timeline ---
    participantsSel: "#coIntakeFilterParticipants",
    timelineSel: "#coIntakeTimeline",
    clustersKey: "coIntakeItems",
    filterParticipantsKey: "coIntakeFilterParticipants",
    hoveredIdxKey: "coIntakeHoveredIdx",
    filterTextKey: "coIntakeFilterText",
    filtered: filteredComposerIntakeCuts,
    highlightCard: highlightCoIntakeCard,
    rerender: renderComposerIntake,
    getDensityEl: function () { return _coIntakeDensityEl; },
    setDensityEl: function (dt) { _coIntakeDensityEl = dt; },
    barCount: function () { return 1; },
    barColor: function (c) {
      return { hue: categoryHue(c.label || "composer") };
    },
    // --- init / delegated listeners ---
    addAllBtnSel: "#coIntakeAddAllBtn",
    reelAllBtnSel: "#coIntakeReelAllBtn",
    thresholdSel: "#coIntakeClusterThreshold", // absent — cuts never cluster
    searchSel: "#coIntakeFilterSearch",
    toggleArtifacts: coIntakeToggleArtifacts,
    toggleReel: coIntakeToggleReel,
    addToArtifacts: coIntakeAddToArtifacts,
    addToReel: coIntakeAddToReel,
    onDismiss: function () {}, // cuts are deleted on the Composer page, not here
    onThresholdChange: function () {},
    onCardHover: function (card, idx) {
      var cut = filteredComposerIntakeCuts()[idx];
      return cut ? (cut.label || "") : "";
    },
    extraControl: function () {},
  };

  // Composer intake poll: the cuts/trims lists are tiny, so fetch the whole
  // composer manifest and fingerprint the spans — re-render only on change.
  // Trims also repaint the SS/TR panels (their asterisk badges read coTrims).
  function pollComposerIntake() {
    return apiGet("../composer/api/manifest")
      .then(function (data) {
        if (!data || !data.ok || !data.manifest) return false;
        var cuts = data.manifest.cuts || [];
        var trims = data.manifest.trims || {};
        var items = cuts.map(function (cut) {
          cut.cutType = "cut";
          return cut;
        });
        // Keys that actually produce a card below — the asterisk badges on
        // sheet/SS/TR/queue cards gate on this so a deep-link never lands on
        // an empty panel.
        var cardKeys = {};
        Object.keys(trims).forEach(function (key) {
          var t = trims[key];
          if (!t) return;
          var participant = t.participant || "";
          var label = t.label || "";
          if (!participant && key.indexOf("sheet:") === 0) {
            // Sheet keys embed the row + participant ("sheet:<row>:<pid>:<seg>"),
            // so metadata-less trims (written before trims carried metadata)
            // can still be carded.
            var bits = key.split(":");
            participant = bits[2] || "";
            if (!label && bits[1]) label = participant + "." + bits[1];
          }
          if (!participant) return; // uncardable (stale SS/TR trim) — no badge either
          cardKeys[key] = true;
          items.push({
            id: key,
            key: key,
            participant: participant,
            start: t.start,
            end: t.end,
            label: label,
            isTrim: true,
            cutType: composerTrimType(key, t.source),
          });
        });
        var fp = JSON.stringify(items.map(function (c) {
          return [c.id, c.participant, c.start, c.end, c.label, !!c.isTrim, c.cutType];
        }));
        if (fp === state._coIntakeFp) return false;
        state._coIntakeFp = fp;
        state.coIntakeItems = items;
        state.coTrims = trims;
        state.coTrimCardKeys = cardKeys;
        renderComposerIntake();
        renderIntake(false);
        renderTranscriptIntake();
        // Queue cards carry trim-asterisk badges too — repaint on change.
        renderArtifactQueue();
        renderReelQueue();
        return true;
      })
      .catch(function () { return false; });
  }

  // Deep link from an SS/TR card's asterisk badge: switch to the Composer
  // Intake tab and highlight the trimmed counterpart (clearing any filters
  // that would hide it).
  function focusComposerIntakeItem(key) {
    var tab = qs('.preview-tab[data-tab="composer-intake"]');
    if (tab) tab.click();

    function indexOfKey() {
      var filtered = filteredComposerIntakeCuts();
      for (var i = 0; i < filtered.length; i++) {
        if (filtered[i].isTrim && filtered[i].key === key) return i;
      }
      return -1;
    }
    var idx = indexOfKey();
    if (idx === -1) {
      state.coIntakeFilterText = "";
      state.coIntakeFilterParticipants = [];
      state.coIntakeFilterTypes = [];
      var searchEl = qs("#coIntakeFilterSearch");
      if (searchEl) searchEl.value = "";
      renderComposerIntake();
      idx = indexOfKey();
    }
    if (idx === -1) return;
    state.coIntakeHoveredIdx = idx;
    highlightCoIntakeCard(idx);
    if (_coIntakeDensityEl) _coIntakeDensityEl.setHovered(idx);
    var card = qsa(".co-intake-queue-card")[idx];
    if (card) card.scrollIntoView({ block: "nearest" });
  }

  function renderComposerIntake() {
    ssClearPending();
    var container = qs("#coIntakeCards");
    if (!container) return;
    var filtered = filteredComposerIntakeCuts();
    var addAllBtn = qs("#coIntakeAddAllBtn");
    var reelAllBtn = qs("#coIntakeReelAllBtn");
    var badge = qs("#coIntakeTabBadge");

    if (badge) {
      if (state.coIntakeItems.length > 0) {
        badge.textContent = state.coIntakeItems.length;
        badge.classList.remove("hidden");
      } else {
        badge.classList.add("hidden");
      }
    }

    if (addAllBtn) addAllBtn.disabled = filtered.length === 0;
    if (reelAllBtn) reelAllBtn.disabled = filtered.length === 0;

    // Both pill rows build before the empty-panel bail so the type legend and
    // its counts stay visible when nothing matches.
    buildCoIntakeTypePills();
    buildIntakeParticipantPills(CO_INTAKE);
    buildIntakeDensityTimeline(CO_INTAKE, filtered);

    if (filtered.length === 0) {
      container.innerHTML = '<div class="drop-target-empty">Composer cuts will appear here. Set in/out pairs on the Composer page</div>';
      return;
    }

    renderIntakeCards(CO_INTAKE, filtered);
  }

  // ---- MindNode Intake ----

  // Teams that work in mind maps rather than spreadsheets get their notes here.
  // Like Composer's panel this is a fingerprint-polled list with no clustering:
  // a note is already a discrete observation. Two things are specific to it —
  // one note can carry several timestamp pairs (each becomes its own card, as a
  // sheet cell expands to segments), and notes carrying no timestamp at all are
  // kept and shown disabled rather than dropped, so the researcher can see what
  // is not clippable instead of silently losing it.

  var _mnIntakeDensityEl = null;

  function mindnodeNoteToItem(note) {
    return {
      participant: note.participant,
      start: note.start,
      end: note.end,
      // A bare-timestamp node has no text; the server falls back to a generic
      // "MindNode intake" label rather than inventing a name.
      desc: note.desc,
      event_type: note.desc,
      category: note.category,
      study: note.study,
      source: "mindnode",
      event_ids: [note.id],
    };
  }

  function mnIntakeAddToArtifacts(note) {
    intakeAddItem(state.artifactQueue, mindnodeNoteToItem(note), renderArtifactQueue);
  }

  function mnIntakeToggleArtifacts(note) {
    intakeToggleItem(state.artifactQueue, mindnodeNoteToItem(note), renderArtifactQueue);
  }

  function mnIntakeAddToReel(note) {
    intakeAddItem(state.reelQueue, mindnodeNoteToItem(note), renderReelQueue);
  }

  function mnIntakeToggleReel(note) {
    intakeToggleItem(state.reelQueue, mindnodeNoteToItem(note), renderReelQueue);
  }

  function filteredMindnodeNotes() {
    var items = state.mnIntakeItems;
    var parts = state.mnIntakeFilterParticipants;
    var cats = state.mnIntakeFilterCategories;
    var text = state.mnIntakeFilterText.toLowerCase();
    if (!parts.length && !cats.length && !text) return items;
    return items.filter(function (n) {
      if (parts.length && parts.indexOf(n.participant) === -1) return false;
      if (cats.length && cats.indexOf(n.category) === -1) return false;
      if (text && (n.desc || "").toLowerCase().indexOf(text) === -1
          && (n.category || "").toLowerCase().indexOf(text) === -1
          && (n.participant || "").toLowerCase().indexOf(text) === -1) return false;
      return true;
    });
  }

  // Category chips are the map's own question branches, so unlike the fixed
  // Composer/detector lists they are derived from the loaded document.
  function buildMnIntakeCategoryPills() {
    var container = qs("#mnIntakeCategoryPills");
    if (!container) return;
    var counts = {};
    var order = [];
    state.mnIntakeItems.forEach(function (n) {
      if (!n.category) return;
      if (counts[n.category] === undefined) order.push(n.category);
      counts[n.category] = (counts[n.category] || 0) + 1;
    });
    container.innerHTML = "";
    order.forEach(function (cat) {
      var chip = ClipgenPrimitives.createFilterChip({
        label: cat,
        active: state.mnIntakeFilterCategories.indexOf(cat) >= 0,
        count: counts[cat],
        color: categoryColor(cat),
        onClick: function () {
          var arr = state.mnIntakeFilterCategories.slice();
          var idx = arr.indexOf(cat);
          if (idx >= 0) arr.splice(idx, 1);
          else arr.push(cat);
          state.mnIntakeFilterCategories = arr;
          renderMindnodeIntake();
        },
      });
      container.appendChild(chip);
    });
  }

  function highlightMnIntakeCard(idx) {
    var cards = qsa(".mn-intake-queue-card");
    for (var i = 0; i < cards.length; i++) {
      if (i === idx) cards[i].classList.add("intake-highlight");
      else cards[i].classList.remove("intake-highlight");
    }
  }

  var MN_INTAKE = {
    cardsSel: "#mnIntakeCards",
    cardSel: ".mn-intake-queue-card",
    cardClass: "mn-intake-queue-card",
    idxAttr: "mnIntakeIdx",
    selfSource: "mindnode",
    setTranscriptContext: false,
    snippet: function (n) { return n.category || ""; },
    clusterToItem: mindnodeNoteToItem,
    cardHue: function (n) { return categoryColor(n.category || "mindnode"); },
    typeText: function (n) { return n.desc || "untitled note"; },
    selfBadge: function (n) {
      return {
        icon: XREF_BADGES.mindnode.icon,
        color: XREF_BADGES.mindnode.color,
        title: n.category ? n.category + " — " + (n.desc || "note") : "MindNode note",
      };
    },
    // --- participant pills + density timeline ---
    participantsSel: "#mnIntakeFilterParticipants",
    timelineSel: "#mnIntakeTimeline",
    clustersKey: "mnIntakeItems",
    filterParticipantsKey: "mnIntakeFilterParticipants",
    hoveredIdxKey: "mnIntakeHoveredIdx",
    filterTextKey: "mnIntakeFilterText",
    filtered: filteredMindnodeNotes,
    highlightCard: highlightMnIntakeCard,
    rerender: renderMindnodeIntake,
    getDensityEl: function () { return _mnIntakeDensityEl; },
    setDensityEl: function (dt) { _mnIntakeDensityEl = dt; },
    barCount: function () { return 1; },
    barColor: function (n) { return { hue: categoryHue(n.category || "mindnode") }; },
    // --- init / delegated listeners ---
    addAllBtnSel: "#mnIntakeAddAllBtn",
    reelAllBtnSel: "#mnIntakeReelAllBtn",
    thresholdSel: "#mnIntakeClusterThreshold", // absent — notes never cluster
    searchSel: "#mnIntakeFilterSearch",
    toggleArtifacts: mnIntakeToggleArtifacts,
    toggleReel: mnIntakeToggleReel,
    addToArtifacts: mnIntakeAddToArtifacts,
    addToReel: mnIntakeAddToReel,
    onDismiss: function () {}, // notes are edited in MindNode, not here
    onThresholdChange: function () {},
    onCardHover: function (card, idx) {
      var note = filteredMindnodeNotes()[idx];
      return note ? (note.text || "") : "";
    },
    extraControl: function () {},
  };

  // Poll the parsed document and fingerprint it — the server re-parses the
  // bundle on each call, so editing the map in MindNode shows up here without
  // reopening the workspace.
  function pollMindnodeIntake() {
    return apiGet("api/mindnode")
      .then(function (data) {
        if (!data || !data.ok || !data.mindnode_loaded || !data.document) return false;
        var items = [];
        var skipped = [];
        (data.document.notes || []).forEach(function (note) {
          if (!note.spans || !note.spans.length) {
            skipped.push(note);
            return;
          }
          // One card per timestamp pair, the way a sheet cell expands to
          // segments — the note id alone would collide across its own spans.
          note.spans.forEach(function (span, segIdx) {
            items.push({
              id: note.id + "#" + segIdx,
              participant: note.participant,
              category: note.category,
              desc: note.desc,
              text: note.text,
              study: note.study,
              start: span[0],
              end: span[1],
            });
          });
        });
        var fp = JSON.stringify([
          items.map(function (n) {
            return [n.id, n.participant, n.category, n.desc, n.start, n.end];
          }),
          skipped.map(function (n) { return [n.id, n.participant, n.desc]; }),
        ]);
        if (fp === state._mnIntakeFp) return false;
        state._mnIntakeFp = fp;
        state.mnIntakeItems = items;
        state.mnIntakeSkipped = skipped;
        renderMindnodeIntake();
        return true;
      })
      .catch(function () { return false; });
  }

  function refreshMindnodeIntake() {
    if (state.mnIntakePoller) return state.mnIntakePoller.wake();
    return pollMindnodeIntake();
  }

  // The untimestamped notes, rendered disabled below the grid. They cannot be
  // cut, but they are real observations the researcher wrote, so hiding them
  // would read as clipgen having lost them.
  function renderMnIntakeSkipped() {
    var host = qs("#mnIntakeSkipped");
    if (!host) return;
    host.innerHTML = "";
    var skipped = state.mnIntakeSkipped || [];
    if (!skipped.length) {
      host.classList.add("hidden");
      return;
    }
    host.classList.remove("hidden");
    host.appendChild(el(
      "div",
      "mn-skipped-head",
      skipped.length + (skipped.length === 1 ? " note has" : " notes have") +
        " no timestamp — add one in MindNode to make them clippable"
    ));
    var list = el("div", "mn-skipped-list");
    skipped.forEach(function (note) {
      var row = el("div", "mn-skipped-row");
      row.appendChild(el("span", "mn-skipped-participant", note.participant));
      row.appendChild(el("span", "mn-skipped-category", note.category || ""));
      row.appendChild(el("span", "mn-skipped-desc", note.desc || note.text || "untitled note"));
      list.appendChild(row);
    });
    host.appendChild(list);
  }

  function renderMindnodeIntake() {
    ssClearPending();
    var container = qs("#mnIntakeCards");
    if (!container) return;
    var filtered = filteredMindnodeNotes();
    var addAllBtn = qs("#mnIntakeAddAllBtn");
    var reelAllBtn = qs("#mnIntakeReelAllBtn");
    var badge = qs("#mnIntakeTabBadge");

    if (badge) {
      if (state.mnIntakeItems.length > 0) {
        badge.textContent = state.mnIntakeItems.length;
        badge.classList.remove("hidden");
      } else {
        badge.classList.add("hidden");
      }
    }

    if (addAllBtn) addAllBtn.disabled = filtered.length === 0;
    if (reelAllBtn) reelAllBtn.disabled = filtered.length === 0;

    buildMnIntakeCategoryPills();
    buildIntakeParticipantPills(MN_INTAKE);
    buildIntakeDensityTimeline(MN_INTAKE, filtered);
    renderMnIntakeSkipped();

    if (filtered.length === 0) {
      container.innerHTML = '<div class="drop-target-empty">Timestamped notes from the mind map appear here. Open a .mindnode document from the Start overlay</div>';
      return;
    }

    renderIntakeCards(MN_INTAKE, filtered);
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

  // Init the intake panels. Folded from the hub-boot initIntakePanel calls
  // so the SS_INTAKE/TR_INTAKE/CO_INTAKE configs never need to leave this file.
  function initIntake() {
    initIntakePanel(SS_INTAKE);
    initIntakePanel(TR_INTAKE);
    initIntakePanel(CO_INTAKE);
    initIntakePanel(MN_INTAKE);
  }

  // ---- Keyboard-cursor access (hub's kbStep/kbSend, keyed by preview tab) ----

  function intakeCfgForTab(tab) {
    if (tab === "intake") return SS_INTAKE;
    if (tab === "transcript-intake") return TR_INTAKE;
    if (tab === "composer-intake") return CO_INTAKE;
    if (tab === "mindnode-intake") return MN_INTAKE;
    return null;
  }

  function intakeCardCount(tab) {
    var cfg = intakeCfgForTab(tab);
    return cfg ? cfg.filtered().length : 0;
  }

  function intakeCardAt(tab, idx) {
    var cfg = intakeCfgForTab(tab);
    if (!cfg) return null;
    var container = qs(cfg.cardsSel);
    return container ? container.children[idx] || null : null;
  }

  function intakeToggleAt(tab, idx, reel) {
    var cfg = intakeCfgForTab(tab);
    var cluster = cfg ? cfg.filtered()[idx] : null;
    if (!cluster) return false;
    if (reel) cfg.toggleReel(cluster);
    else cfg.toggleArtifacts(cluster);
    return true;
  }

  // ---- Published to the hub (window.ClipgenStudio) ----
  STUDIO.initIntake = initIntake;
  STUDIO.intakeCardCount = intakeCardCount;
  STUDIO.intakeCardAt = intakeCardAt;
  STUDIO.intakeToggleAt = intakeToggleAt;
  STUDIO.pollScreenspaceIntake = pollScreenspaceIntake;
  STUDIO.pollTranscriptIntake = pollTranscriptIntake;
  STUDIO.pollComposerIntake = pollComposerIntake;
  STUDIO.pollMindnodeIntake = pollMindnodeIntake;
  STUDIO.refreshScreenspaceIntake = refreshScreenspaceIntake;
  STUDIO.refreshTranscriptIntake = refreshTranscriptIntake;
  STUDIO.refreshComposerIntake = refreshComposerIntake;
  STUDIO.refreshMindnodeIntake = refreshMindnodeIntake;
  STUDIO.focusComposerIntakeItem = focusComposerIntakeItem;
  STUDIO.initTooltipToggle = initTooltipToggle;
  STUDIO.refreshIntakeCardStates = refreshIntakeCardStates;
  STUDIO.renderIntake = renderIntake;
  STUDIO._syncMarkCategoriesFromSettings = _syncMarkCategoriesFromSettings;
})();
