/* Overview hub — shared state, data loading, and tab switching for /overview/.
 *
 * The Overview page hosts three cohort-level tabs as satellites: Map
 * (overview-map.js, the 3D similarity space), Convergence
 * (overview-convergence.js), and Metadata (overview-metadata.js). They share
 * this hub's state through the window.ClipgenOverview (OV) namespace — the
 * same hub-plus-satellite pattern Studio/Screenspace/Transcripts use.
 *
 * Data loading is bootstrap-fetch, not polling: Overview is an exploration
 * surface, so ensureData() memoizes one parallel round of cross-prefix
 * fetches (../studio/api/sheet, ../studio/api/sheet/baseline,
 * ../screenspace/api/events, ../transcripts/api/marks) and a Refresh action
 * re-runs it. All four degrade gracefully when their blueprint has no data
 * (no spreadsheet loaded, no scans yet, no transcripts yet).
 *
 * Clustering mirrors studio-intake.js: state.intakeClusters holds only
 * non-navigational events (boundary events are orientation scaffolding);
 * state.intakeEvents keeps ALL events (Metadata's boundary count reads it).
 */

(function () {
  "use strict";

  var OV = {};
  window.ClipgenOverview = OV;

  // Same default the Studio intake threshold input uses.
  var CLUSTER_THRESHOLD_SEC = 10;

  var state = {
    sheetData: null,          // ../studio/api/sheet payload (null until loaded)
    convergenceBaselines: {}, // participant -> baseline seconds
    metadataClusterScreenspace: true,
    intakeEvents: [],         // all screenspace events (incl. navigational)
    intakeClusters: [],       // clustered, non-navigational
    trIntakeMarks: [],        // valid transcript marks
    trIntakeClusters: [],
    composerCuts: [],         // ../composer/api/manifest cuts (each carries participant)
    frictionMoments: [],      // LLM friction moments with resolved times
    activeTab: "metadata",
    // Bumped after every completed loadAll(). The tabs' staleness snapshots
    // compare against this — data can only "change" via an actual refetch,
    // never via length-heuristic false positives.
    dataVersion: 0,
  };
  OV.state = state;

  // ---- Shared helpers (lifted from studio.js; bind this page's state) ----

  function parseClipTimestamps(raw, participantId) {
    var DEFAULT_DUR = CLIPGEN_CONFIG.defaultDuration;
    var baselineSeconds = 0;
    if (participantId && state.convergenceBaselines) {
      baselineSeconds = state.convergenceBaselines[participantId] || 0;
    }
    return parseClipSegmentsForCell(raw, baselineSeconds, DEFAULT_DUR);
  }

  var ROW_FUNCTIONS = {
    Count: function (row, participants) {
      var total = 0;
      for (var j = 0; j < participants.length; j++) {
        var c = row.cells[participants[j]];
        if (c && c.valid) total += parseClipTimestamps(c.value, participants[j]).length;
      }
      return total;
    },
    Unique: function (row, participants) {
      var count = 0;
      for (var j = 0; j < participants.length; j++) {
        var c = row.cells[participants[j]];
        if (c && c.valid) count++;
      }
      return count;
    },
  };

  // Cross-referencing: overlapping data from the other sources for a given
  // participant + time range (same contract as Studio's copy — consumers are
  // the moved Convergence detail rows and the Map drill-down).
  function findOverlappingData(participant, start, end) {
    var result = { transcriptSnippets: [], screenspaceEvents: [], sheetObservations: [] };

    for (var i = 0; i < state.trIntakeClusters.length; i++) {
      var tc = state.trIntakeClusters[i];
      if (tc.participant === participant && tc.start < end && tc.end > start) {
        result.transcriptSnippets.push({ text: tc.text || tc.label || "", category: tc.category, start: tc.start, end: tc.end });
      }
    }

    for (var j = 0; j < state.intakeClusters.length; j++) {
      var sc = state.intakeClusters[j];
      if (sc.participant === participant && sc.start < end && sc.end > start) {
        result.screenspaceEvents.push(sc);
      }
    }

    if (state.sheetData && state.sheetData.rows) {
      for (var k = 0; k < state.sheetData.rows.length; k++) {
        var row = state.sheetData.rows[k];
        var cell = row.cells[participant];
        if (!cell || !cell.valid) continue;
        var segs = parseClipTimestamps(cell.value, participant);
        for (var s = 0; s < segs.length; s++) {
          var segEnd = segs[s].startSeconds + segs[s].duration;
          if (segs[s].startSeconds < end && segEnd > start) {
            result.sheetObservations.push(row);
            break;
          }
        }
      }
    }

    return result;
  }

  OV.parseClipTimestamps = parseClipTimestamps;
  OV.ROW_FUNCTIONS = ROW_FUNCTIONS;
  OV.findOverlappingData = findOverlappingData;

  // ---- Data loading ----

  var _dataPromise = null;

  function buildClusters() {
    var cluster = window.ClipgenIntakeCluster;
    if (!cluster) return;
    var nonNav = state.intakeEvents.filter(function (ev) { return !ev.navigational; });
    state.intakeClusters = cluster.clusterIntakeEvents(nonNav, CLUSTER_THRESHOLD_SEC);
    state.trIntakeClusters = cluster.clusterTranscriptMarks(state.trIntakeMarks, CLUSTER_THRESHOLD_SEC);
  }

  function loadAll() {
    var sheetP = apiGet("../studio/api/sheet")
      .then(function (data) {
        if (data && data.config) clipgenApplyConfig(data.config);
        state.sheetData = data || null;
        if (data && data.metadataClusterScreenspace !== undefined) {
          state.metadataClusterScreenspace = data.metadataClusterScreenspace;
        }
        setStudyName();
      })
      .catch(function () { state.sheetData = null; });

    var baselineP = apiGet("../studio/api/sheet/baseline")
      .then(function (data) {
        state.convergenceBaselines = (data && data.baselines) || {};
      })
      .catch(function () { state.convergenceBaselines = {}; });

    var eventsP = apiGet("../screenspace/api/events?excluded=false")
      .then(function (data) {
        state.intakeEvents = (data && data.events) || [];
      })
      .catch(function () { state.intakeEvents = []; });

    var marksP = apiGet("../transcripts/api/marks")
      .then(function (data) {
        var marks = (data && data.marks) || [];
        state.trIntakeMarks = marks.filter(function (m) { return m.valid; });
      })
      .catch(function () { state.trIntakeMarks = []; });

    var frictionP = apiGet("api/friction-moments")
      .then(function (data) {
        state.frictionMoments = (data && data.moments) || [];
      })
      .catch(function () { state.frictionMoments = []; });

    // Composer cuts are one flat array (each cut carries its own participant),
    // so a single fetch covers every participant's 4th Convergence lane.
    var composerP = apiGet("../composer/api/manifest")
      .then(function (data) {
        state.composerCuts = (data && data.manifest && data.manifest.cuts) || [];
      })
      .catch(function () { state.composerCuts = []; });

    return Promise.all([sheetP, baselineP, eventsP, marksP, frictionP, composerP]).then(function () {
      buildClusters();
      state.dataVersion++;
      return state;
    });
  }

  function ensureData() {
    if (!_dataPromise) _dataPromise = loadAll();
    return _dataPromise;
  }

  // Drop the memo and refetch; satellites re-render from their activate paths.
  function refreshData() {
    _dataPromise = null;
    return ensureData();
  }

  OV.ensureData = ensureData;
  OV.refreshData = refreshData;
  OV.buildClusters = buildClusters;

  // ---- Tabs ----

  function setStudyName() {
    var elName = qs("#ovStudyName");
    if (!elName) return;
    var study = state.sheetData && state.sheetData.study;
    elName.textContent = study || "";
    var elMeta = qs("#ovStudyMeta");
    if (elMeta) {
      var participants = state.sheetData && state.sheetData.participants;
      elMeta.textContent = participants ? participants.length + " participants" : "";
    }
  }

  function initTabs() {
    var tabs = qsa(".preview-tab");
    for (var i = 0; i < tabs.length; i++) {
      tabs[i].addEventListener("click", function () {
        var target = this.dataset.tab;
        if (target === state.activeTab) return;
        state.activeTab = target;
        setStoredUIStateField("overview", "activeTab", target);
        var allTabs = qsa(".preview-tab");
        for (var j = 0; j < allTabs.length; j++) allTabs[j].classList.remove("active");
        this.classList.add("active");
        syncTab(true);
      });
    }
    restoreStoredTab();
    syncTab(false);
  }

  function restoreStoredTab() {
    var stored = getStoredUIState("overview").activeTab;
    if (!stored || stored === state.activeTab) return;
    var tabs = qsa(".preview-tab");
    for (var i = 0; i < tabs.length; i++) {
      if (tabs[i].dataset.tab === stored) {
        state.activeTab = stored;
        for (var k = 0; k < tabs.length; k++) tabs[k].classList.remove("active");
        tabs[i].classList.add("active");
        return;
      }
    }
  }

  function syncTab(animate) {
    var panels = {
      map: qs("#mapPanel"),
      convergence: qs("#convergencePanel"),
      metadata: qs("#metadataPanel"),
    };
    var key;
    for (key in panels) {
      if (panels[key]) panels[key].classList.add("hidden");
    }

    if (OV.convergenceDeactivate) OV.convergenceDeactivate();
    if (OV.metadataDeactivate) OV.metadataDeactivate();
    if (OV.mapDeactivate) OV.mapDeactivate();

    var activePanel = panels[state.activeTab];
    if (activePanel) activePanel.classList.remove("hidden");
    if (state.activeTab === "map" && OV.mapActivate) OV.mapActivate();
    if (state.activeTab === "convergence" && OV.convergenceActivate) OV.convergenceActivate();
    if (state.activeTab === "metadata" && OV.metadataActivate) OV.metadataActivate();

    if (activePanel && animate) {
      activePanel.classList.add("tab-slide-enter");
      requestAnimationFrame(function () {
        requestAnimationFrame(function () {
          activePanel.classList.remove("tab-slide-enter");
        });
      });
    }
  }

  OV.syncTab = syncTab;

  // ---- Boot ----

  document.addEventListener("DOMContentLoaded", function () {
    // TopNav renders #themeToggle / #settingsBtn synchronously before this
    // hub loads; wire them here as the other surfaces do (utils.js owns the
    // theme logic, settings-modal.js the shared modal).
    initThemeToggle();
    var settingsBtn = qs("#settingsBtn");
    if (settingsBtn && typeof window.openSettingsModal === "function") {
      settingsBtn.addEventListener("click", function () {
        window.openSettingsModal({});
      });
    }

    var refreshBtn = qs("#ovRefresh");
    if (refreshBtn) {
      refreshBtn.addEventListener("click", function () {
        refreshData().then(function () {
          // Re-activate the current tab so it re-renders from fresh state.
          syncTab(false);
        });
      });
    }

    window.addEventListener("resize", function () {
      if (state.activeTab === "map" && OV.mapResize) OV.mapResize();
      if (state.activeTab === "convergence" && OV.convergenceResize) OV.convergenceResize();
      if (state.activeTab === "metadata" && OV.metadataResize) OV.metadataResize();
    });

    ensureData();
    initTabs();
  });
})();
