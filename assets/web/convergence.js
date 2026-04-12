/* Convergence Browser – Phase 1: Tab shell and filter controls */

(function () {
  "use strict";

  var cvState = {
    active: false,
    initialized: false,
    baselines: null,
    events: [],
    filteredEvents: [],
    convergenceZones: [],
    selection: null,
    filters: {
      streams: [],
      eventTypes: [],
      minParticipants: 2,
      windowSec: 5,
      timeRange: null,
    },
    dataVersion: 0,
    duration: 0,
    participants: [],
  };

  // --- Utilities ---

  function debounce(fn, ms) {
    var timer;
    return function () {
      var ctx = this, args = arguments;
      clearTimeout(timer);
      timer = setTimeout(function () { fn.apply(ctx, args); }, ms);
    };
  }

  // --- Filter Controls ---

  var STREAM_DEFS = [
    { key: "all", label: "All streams", color: "var(--color-accent)" },
    { key: "sheet", label: "Sheet", color: XREF_BADGES.sheet.color },
    { key: "screenspace", label: "Screenspace", color: XREF_BADGES.screenspace.color },
    { key: "transcript", label: "Transcript", color: XREF_BADGES.transcript.color },
  ];

  function buildFilterControls() {
    var controls = qs("#convergenceControls");
    var filters = qs("#convergenceFilters");
    if (!controls || !filters) return;

    // --- Controls bar: min participants + window ---
    var minLabel = el("label", "intake-cluster-label");
    minLabel.textContent = "Min participants ";
    var minInput = document.createElement("input");
    minInput.type = "number";
    minInput.min = "2";
    minInput.value = String(cvState.filters.minParticipants);
    minInput.className = "intake-cluster-input";
    minInput.autocomplete = "off";
    minInput.addEventListener("input", debouncedFilterChange);
    minLabel.appendChild(minInput);

    var winLabel = el("label", "intake-cluster-label");
    winLabel.textContent = "Window ";
    var winInput = document.createElement("input");
    winInput.type = "number";
    winInput.min = "1";
    winInput.max = "60";
    winInput.value = String(cvState.filters.windowSec);
    winInput.className = "intake-cluster-input";
    winInput.autocomplete = "off";
    winInput.addEventListener("input", debouncedFilterChange);
    var winSuffix = document.createTextNode("\u00B1s");
    winLabel.appendChild(winInput);
    winLabel.appendChild(winSuffix);

    controls.appendChild(minLabel);
    controls.appendChild(winLabel);

    // --- Filters bar: stream toggles + event type pills ---
    for (var i = 0; i < STREAM_DEFS.length; i++) {
      var def = STREAM_DEFS[i];
      var btn = document.createElement("button");
      btn.className = "cv-stream-toggle" + (def.key === "all" ? " active" : "");
      btn.textContent = def.label;
      btn.dataset.stream = def.key;
      btn.style.setProperty("--det-color", def.color);
      btn.addEventListener("click", (function (key) {
        return function () { onStreamToggle(key); };
      })(def.key));
      filters.appendChild(btn);
    }

    var typePills = document.createElement("div");
    typePills.id = "cvEventTypePills";
    filters.appendChild(typePills);
  }

  function onStreamToggle(stream) {
    if (stream === "all") {
      cvState.filters.streams = [];
    } else {
      var idx = cvState.filters.streams.indexOf(stream);
      if (idx >= 0) {
        cvState.filters.streams.splice(idx, 1);
      } else {
        cvState.filters.streams.push(stream);
      }
    }

    // Update pill active states
    var pills = qsa(".cv-stream-toggle");
    var isAll = cvState.filters.streams.length === 0;
    for (var i = 0; i < pills.length; i++) {
      var key = pills[i].dataset.stream;
      if (key === "all") {
        pills[i].classList.toggle("active", isAll);
      } else {
        pills[i].classList.toggle("active", cvState.filters.streams.indexOf(key) >= 0);
      }
    }

    onFilterChange();
  }

  function syncFilterInputs() {
    var controls = qs("#convergenceControls");
    if (!controls) return;
    var inputs = controls.querySelectorAll("input[type=number]");
    if (inputs[0]) cvState.filters.minParticipants = Math.max(2, parseInt(inputs[0].value, 10) || 2);
    if (inputs[1]) cvState.filters.windowSec = Math.max(1, parseInt(inputs[1].value, 10) || 5);
  }

  function onFilterChange() {
    syncFilterInputs();
    // Phase 2: recalculate convergence zones and re-render
  }

  var debouncedFilterChange = debounce(onFilterChange, 250);

  // --- Lifecycle ---

  function activate() {
    cvState.active = true;
    if (!cvState.initialized) {
      buildFilterControls();
      cvState.initialized = true;
    }
  }

  function deactivate() {
    cvState.active = false;
  }

  function init() {
    // Phase 2: fetch baselines, collect events
  }

  // --- Window exports ---
  window.convergenceActivate = activate;
  window.convergenceDeactivate = deactivate;
  window.convergenceInit = init;
  window.convergenceResize = function () { /* Phase 2: resize canvases */ };

  init();
})();
