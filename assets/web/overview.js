/* Overview hub — shared state, data loading, and tab switching for /overview/.
 *
 * The Overview page hosts three cohort-level tabs as satellites: Metadata
 * (overview-metadata.js), Convergence (overview-convergence.js), and Reports
 * (overview-reports.js, the per-participant mini-report). They share
 * this hub's state through the window.ClipgenOverview (OV) namespace — the
 * same hub-plus-satellite pattern Studio/Screenspace/Transcripts use.
 *
 * Data loading is bootstrap-fetch, not polling: Overview is an exploration
 * surface, so ensureData() memoizes one parallel round of cross-prefix
 * fetches (../studio/api/sheet, ../studio/api/sheet/baseline,
 * ../screenspace/api/events, ../transcripts/api/marks) and a Refresh action
 * re-runs it. All of them degrade gracefully when their blueprint has no data
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
    activeTab: "metadata",
    // Bumped after every completed loadAll(); tab staleness snapshots compare against it.
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

  // Overlapping data from other sources for a participant + time range (Convergence detail rows).
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

    // Cuts are one flat array carrying participants, so one fetch covers every lane.
    var composerP = apiGet("../composer/api/manifest")
      .then(function (data) {
        state.composerCuts = (data && data.manifest && data.manifest.cuts) || [];
      })
      .catch(function () { state.composerCuts = []; });

    return Promise.all([sheetP, baselineP, eventsP, marksP, composerP]).then(function () {
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

  // Staleness paint: tabs report a stale dataVersion, the hub's single Refresh button shows it.
  function setRefreshStale(stale) {
    var btn = qs("#ovRefresh");
    if (!btn) return;
    btn.classList.toggle("is-stale", stale);
    btn.title = stale
      ? "New upstream data available — click to refresh"
      : "Re-fetch sheet, Screenspace, and transcript data";
  }

  OV.setRefreshStale = setRefreshStale;

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
    // #tab=KEY deep links (command palette) beat the stored tab; the hash persists across reloads.
    var stored = clipgenHashTab() || getStoredUIState("overview").activeTab;
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

  var TAB_KEYS = ["convergence", "metadata", "reports"];

  // Call a tab satellite's lifecycle hook (e.g. OV.reportsActivate) if the
  // satellite published it.
  function tabHook(tab, phase) {
    var fn = OV[tab + phase];
    if (fn) fn();
  }

  // Settings modal + palette hook. syncTab would close the open zone detail without rebuilding badges.
  function rerenderCrossRefs() {
    if (OV.convergenceRenderCrossRefs) OV.convergenceRenderCrossRefs();
  }
  window.clipgenRerenderCrossRefs = rerenderCrossRefs;

  function syncTab(animate) {
    var panels = {
      convergence: qs("#convergencePanel"),
      metadata: qs("#metadataPanel"),
      reports: qs("#reportsPanel"),
    };
    var key;
    for (key in panels) {
      if (panels[key]) panels[key].classList.add("hidden");
    }

    TAB_KEYS.forEach(function (tab) { tabHook(tab, "Deactivate"); });

    // Drop the outgoing tab's paint; the incoming one re-asserts from activate.
    setRefreshStale(false);

    var activePanel = panels[state.activeTab];
    if (activePanel) activePanel.classList.remove("hidden");
    tabHook(state.activeTab, "Activate");

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

  // Overview has no TopNav quick actions, so the palette gets tab switchers and refresh here.
  function initCommandPalette() {
    if (!window.ClipgenCommandPalette) return;
    window.ClipgenCommandPalette.setParticipants(function () {
      return (state.sheetData && state.sheetData.participants) || [];
    });
    function tabCommand(tabKey, title, icon) {
      function tabEl() {
        return qs('.preview-tab[data-tab="' + tabKey + '"]');
      }
      return {
        id: "overview:tab-" + tabKey,
        title: title,
        icon: icon,
        keywords: "tab show switch",
        section: "Overview",
        visible: function () { return !!tabEl(); },
        run: function () { tabEl().click(); },
      };
    }
    window.ClipgenCommandPalette.register("overview", [
      tabCommand("metadata", "Show Metadata tab", "table-cells"),
      tabCommand("convergence", "Show Convergence tab", "arrows-pointing-in"),
      tabCommand("reports", "Show Reports tab", "document-text"),
      {
        id: "overview:refresh",
        title: "Refresh Overview data",
        icon: "arrow-path",
        keywords: "reload fetch update",
        section: "Overview",
        visible: function () { return !!qs("#ovRefresh"); },
        run: function () { qs("#ovRefresh").click(); },
      },
      {
        id: "overview:reset-offsets",
        title: "Reset convergence offsets",
        icon: "arrow-uturn-left",
        keywords: "alignment convergence clear restore per-participant",
        section: "Overview",
        // The button only exists unhidden on Convergence once offsets are set.
        visible: function () {
          var b = qs("#cvResetOffsetsBtn");
          return !!b && !b.classList.contains("hidden");
        },
        run: function () { qs("#cvResetOffsetsBtn").click(); },
      },
    ]);
  }

  // ---- Boot ----

  document.addEventListener("DOMContentLoaded", function () {
    // Static [data-icon] elements need mask resolution; createBtn handles JS-built ones.
    applyIconMasksIn(document);
    // TopNav renders these buttons before this hub loads; wire them like the other surfaces.
    if (typeof initThemeToggle === "function") {
      initThemeToggle();
    }
    if (window.wireSettingsButton) {
      window.wireSettingsButton({
        onApply: function (applied, settings) {
          if (applyCrossRefSetting(applied, settings)) rerenderCrossRefs();
        },
      });
    }

    var refreshBtn = qs("#ovRefresh");
    if (refreshBtn) {
      refreshBtn.addEventListener("click", function () {
        if (refreshBtn.disabled) return;
        refreshBtn.disabled = true;
        // Spin the masked icon span — there is no inline <svg> to animate.
        var icon = refreshBtn.querySelector(".cg-btn-icon");
        if (icon) icon.style.animation = "spin 0.7s linear infinite";
        refreshData().then(function () {
          // Re-activate the current tab so it re-renders from fresh state.
          syncTab(false);
        }).catch(function () {
          // Never leave the button stuck disabled + spinning.
        }).then(function () {
          refreshBtn.disabled = false;
          if (icon) icon.style.animation = "";
        });
      });
    }

    window.addEventListener("resize", function () {
      if (state.activeTab === "convergence" && OV.convergenceResize) OV.convergenceResize();
      if (state.activeTab === "metadata" && OV.metadataResize) OV.metadataResize();
      if (state.activeTab === "reports" && OV.reportsResize) OV.reportsResize();
    });

    ensureData();
    initTabs();
    initCommandPalette();
    initHotkeys();
  });

  // ---- Hotkeys ----

  function switchToTab(name) {
    var tabs = qsa(".preview-tab");
    for (var i = 0; i < tabs.length; i++) {
      if (tabs[i].dataset.tab === name) {
        tabs[i].click();
        return;
      }
    }
  }

  function initHotkeys() {
    window.ClipgenHotkeys.register([
      { id: "overview.tabConvergence", handler: function () { switchToTab("convergence"); } },
      { id: "overview.tabMetadata", handler: function () { switchToTab("metadata"); } },
      { id: "overview.tabReports", handler: function () { switchToTab("reports"); } },
      {
        id: "global.refresh",
        handler: function () {
          var btn = qs("#ovRefresh");
          if (btn) btn.click();
        },
      },
      {
        id: "global.search",
        when: function () { return state.activeTab === "metadata"; },
        handler: function () {
          var input = qs("#mdSearchInput");
          if (!input) return false;
          input.focus();
          input.select();
        },
      },
    ]);
  }
})();
