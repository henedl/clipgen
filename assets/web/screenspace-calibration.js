/* clipgen Screenspace calibration satellite — screenspace-calibration.js
 *
 * The calibration strip: scores the participant's pins against the active
 * tool + params and plots them on a normalized matchiness axis with a
 * draggable threshold line. Carved out of screenspace.js to shrink the page
 * script; loaded after it. Reads hub state + helpers via
 * window.ClipgenScreenspace and registers its entry points as SS.cal*; the
 * hub keeps thin same-named delegators so its ~30 call sites are unchanged.
 * Function bodies are unchanged from when they lived inline in screenspace.js.
 */
(function () {
  "use strict";

  var SS = window.ClipgenScreenspace;
  var state = SS.state;
  var _previewRegionRef = SS._previewRegionRef,
    gatherWorkflowParams = SS.gatherWorkflowParams,
    loadFrame = SS.loadFrame,
    regionRefPayload = SS.regionRefPayload,
    renderWorkflowParams = SS.renderWorkflowParams,
    restoreTaskToWorkflow = SS.restoreTaskToWorkflow,
    setInputValue = SS.setInputValue,
    syncValueDisplays = SS.syncValueDisplays,
    updateRunButton = SS.updateRunButton;

  // ---- Calibration strip ----
  // Scores pinned frames only (never re-runs on seek); temporal params unvalidated.

  var _calibrationGen = 0;
  var _calibrationTimer = 0;
  // Pin verdict per threshold slider id; rebuilt by every renderCalibration, read by updateCalibrationSliderMarks.
  var _calBands = {};

  // Per-tool axis: sliderId gives the range; `compare` is the pass test, `invert` only flips display.
  var CAL_AXIS = {
    change: { sliderId: "paramChangeThresh", invert: false, drawLine: true, compare: "ge" },
    similarity: { sliderId: "paramSimThresh", invert: false, drawLine: true, compare: "ge" },
    text: { sliderId: "paramTextFuzzy", invert: false, drawLine: true, compare: "ge" },
    template: { sliderId: "paramTemplateThresh", invert: false, drawLine: true, compare: "ge" },
    shape: { sliderId: "paramShapeThresh", invert: false, drawLine: true, compare: "ge" },
    flow: { sliderId: "paramFlowMag", invert: false, drawLine: true, compare: "ge" },
    numbers: { sliderId: "paramNumOcrConf", invert: false, drawLine: true, compare: "ge" },
    inactivity: { sliderId: "paramInactThresh", invert: true, drawLine: true, compare: "le" },
    color: {
      sliderId: null, rangeMin: 0, rangeMax: 1, invert: false, drawLine: false,
      rowNote: "Tolerance ≠ confidence. Read the gap between green and red dots.",
    },
    scene: {
      sliderId: null, rangeMin: 0, rangeMax: 1, invert: false, drawLine: false,
      rowNote: "Per-scene thresholds. Dot colour is pass/fail; hover for the matched scene.",
    },
  };

  // Presence mode: score is pixel coverage, Min-area is the cutoff; scoreScale maps 0–1 to percent.
  function _calColorAxis(sfx) {
    var modeEl = qs("#paramColorMode" + (sfx || ""));
    if (modeEl && modeEl.value === "presence") {
      return {
        sliderId: "paramColorMinArea", invert: false, drawLine: true,
        compare: "ge", scoreScale: 100,
      };
    }
    return CAL_AXIS.color;
  }

  function _calIsCalibratable(tool) {
    return tool === "multitool" || !!CAL_AXIS[tool];
  }

  function _calTitle(type) {
    return type ? type.charAt(0).toUpperCase() + type.slice(1) : "";
  }

  // Range from the live slider, else the descriptor (color / scene have none).
  function _calAxisRange(axis, sliderId) {
    var slider = sliderId ? qs("#" + sliderId) : null;
    if (slider) {
      var mn = parseFloat(slider.min);
      var mx = parseFloat(slider.max);
      if (isFinite(mn) && isFinite(mx) && mx > mn) {
        return { min: mn, max: mx, value: parseFloat(slider.value) };
      }
    }
    return {
      min: axis && axis.rangeMin != null ? axis.rangeMin : 0,
      max: axis && axis.rangeMax != null ? axis.rangeMax : 1,
      value: null,
    };
  }

  function _calPos(value, range, invert) {
    if (!(range.max > range.min)) return 0;
    var t = (value - range.min) / (range.max - range.min);
    if (t < 0) t = 0; else if (t > 1) t = 1;
    if (invert) t = 1 - t;
    return t * 100;
  }

  // A pin's pass/fail contradicts its polarity when a positive doesn't fire or a
  // negative does.
  function _calContradicts(polarity, passed) {
    if (passed == null) return false;
    return polarity === "positive" ? !passed : !!passed;
  }

  // Reduce one pin to pass / fail / null (indeterminate or not-evaluable).
  function _calPinPass(tool, e) {
    if (tool === "multitool") {
      if (!e || e.status === "not_evaluable" || !e.steps) return null;
      return e.passed == null ? null : !!e.passed;
    }
    if (!e || e.status !== "ok") return null;
    return !!e.passed;
  }

  function _calDotTooltip(tool, sc, timestamp, scoreScale) {
    var lines = [formatTime(timestamp, { decimals: 1 })];
    if (!sc || sc.status === "not_evaluable" || sc.score == null) {
      lines.push("not evaluable");
      return lines.join("\n");
    }
    var d = sc.detail || {};
    var s = Number(sc.score);
    if (tool === "color" && scoreScale && scoreScale !== 1) {
      lines.push("coverage " + (s * scoreScale).toFixed(1) + "%");
    } else if (tool === "text") {
      lines.push("fuzzy " + s.toFixed(2));
      if (d.text_found) lines.push("“" + d.text_found + "”");
    } else if (tool === "numbers") {
      lines.push("OCR conf " + s.toFixed(2));
      if (d.number_found != null) lines.push("read " + d.number_found);
    } else if (tool === "scene") {
      lines.push("similarity " + s.toFixed(2));
      if (d.scene_name) lines.push(d.scene_name);
    } else if (tool === "inactivity") {
      lines.push("distance " + Math.round(s));
    } else if (tool === "flow") {
      lines.push("magnitude " + s.toFixed(2));
    } else {
      lines.push("score " + s.toFixed(3));
    }
    lines.push(sc.passed ? "✓ matches" : "✗ no match");
    return lines.join("\n");
  }

  // "Nice" 1/2/5 × 10ⁿ ticks across [min, max], about five intervals.
  function _calGridTicks(min, max) {
    if (!(max > min)) return { ticks: [], step: 0 };
    var rough = (max - min) / 5;
    var mag = Math.pow(10, Math.floor(Math.log(rough) / Math.LN10));
    var norm = rough / mag;
    var mult = norm < 1.5 ? 1 : (norm < 3 ? 2 : (norm < 7 ? 5 : 10));
    var step = mult * mag;
    var ticks = [];
    var start = Math.ceil(min / step) * step;
    for (var v = start; v <= max + step * 1e-6; v += step) {
      ticks.push(Math.round(v / step) * step);
    }
    return { ticks: ticks, step: step };
  }

  // Value grid behind the dots, using the same value→percent mapping as them.
  function _calBuildGrid(ax, range, invert) {
    var info = _calGridTicks(range.min, range.max);
    info.ticks.forEach(function (v) {
      var pos = _calPos(v, range, invert);
      var line = el("div", "cal-grid-line");
      line.style.left = pos + "%";
      ax.appendChild(line);
      var label = el("span", "cal-grid-label", _calFmtVal(v, info.step));
      label.style.left = pos + "%";
      ax.appendChild(label);
    });
  }

  // One track (axis, dots, threshold line). rows: [{polarity, timestamp, sc, stale}].
  function _calBuildTrack(rows, tool, axis, sliderId, label, suggest) {
    var track = el("div", "cal-track");
    if (label) track.appendChild(label);
    // Backend score → slider units (100 for color presence coverage).
    var scoreScale = axis.scoreScale || 1;
    var range = _calAxisRange(axis, sliderId);
    var ax = el("div", "cal-axis");
    _calBuildGrid(ax, range, axis.invert);
    if (axis.drawLine && range.value != null) {
      var line = el("div", "cal-threshold");
      line.style.left = _calPos(range.value, range, axis.invert) + "%";
      line.setAttribute("data-cal-slider", sliderId || "");
      line.setAttribute("data-cal-invert", axis.invert ? "1" : "0");
      ax.appendChild(line);
    }
    // Suggested cutoff: mid-gap with both polarities, else hugging the pinned cluster's edge.
    var suggestion = (suggest && axis.drawLine && sliderId) ? _calSuggest(rows, axis.compare, scoreScale) : null;
    var applyBadge = null;
    var narrowGap = false;
    if (suggestion && suggestion.separated) {
      var slider = qs("#" + sliderId);
      var step = slider ? parseFloat(slider.step) : 0;
      if (!isFinite(step)) step = 0;
      // Step-aligned; a raw target can snap onto a boundary and leak a pin.
      var applyVal = _calApplyValue(suggestion.lo, suggestion.hi, axis.compare, step, range.min, range.max);
      if (applyVal == null) {
        narrowGap = true; // valid interval exists but no step-aligned value lands in it
      } else {
        // Recorded only when a reachable cutoff exists, so slider marks match the Apply badge.
        _calBands[sliderId] = {
          lo: suggestion.lo, hi: suggestion.hi, compare: axis.compare,
          applyVal: applyVal, min: range.min, max: range.max, step: step,
        };
        // Marker sits where Apply lands, so the threshold line meets it exactly.
        var marker = el("div", "cal-suggestion");
        marker.style.left = _calPos(applyVal, range, axis.invert) + "%";
        ax.appendChild(marker);
        applyBadge = el("button", "cal-suggest-apply", "Apply " + _calFmtVal(applyVal, step));
        applyBadge.type = "button";
        var applyTip;
        if (suggestion.mode === "gap") {
          applyTip = "Set the threshold midway between your positives and negatives (gap " + _calFmtVal(suggestion.margin, step) + ").";
        } else if (suggestion.mode === "positives") {
          applyTip = "Set the threshold at the edge of your positives. Pin a negative to widen the margin.";
        } else {
          applyTip = "Set the threshold just past your negatives. Pin a positive to widen the margin.";
        }
        applyBadge.setAttribute("data-tooltip", applyTip);
        (function (sid, val, sl) {
          applyBadge.addEventListener("click", function () {
            setInputValue("#" + sid, val);
            if (sl) {
              // Bubbles to #workflowParams → glides the line + re-evaluates scores.
              sl.dispatchEvent(new Event("input", { bubbles: true }));
              sl.dispatchEvent(new Event("change", { bubbles: true }));
            }
            syncValueDisplays();
          });
        })(sliderId, applyVal, slider);
      }
    }
    var stack = {}; // rounded position bucket -> count, for collision fan-out
    rows.forEach(function (r) {
      var sc = r.sc;
      var dot = el("button", "cal-dot cal-dot--" + (r.polarity === "negative" ? "negative" : "positive"));
      dot.type = "button";
      var evaluable = sc && sc.status === "ok" && sc.score != null && isFinite(sc.score);
      var pos = evaluable ? _calPos(sc.score * scoreScale, range, axis.invert) : 0;
      dot.style.left = pos + "%";
      if (!evaluable) dot.classList.add("cal-dot--hollow");
      if (r.stale) dot.classList.add("cal-dot--hollow", "cal-dot--stale");
      if (evaluable && _calContradicts(r.polarity, sc.passed)) dot.classList.add("cal-dot--fail");
      // Fan coincident dots upward so each stays hoverable; cap so stacks don't overflow.
      var key = Math.round(pos / 3);
      var n = stack[key] || 0;
      stack[key] = n + 1;
      if (n > 0) dot.style.bottom = "calc(var(--space-2) * " + Math.min(n, 4) + ")";
      var tip = _calDotTooltip(tool, sc, r.timestamp, scoreScale);
      dot.setAttribute("data-tooltip", tip);
      dot.setAttribute("aria-label", tip.replace(/\n/g, ", "));
      (function (ts) {
        dot.addEventListener("click", function () { loadFrame(ts); });
      })(r.timestamp);
      ax.appendChild(dot);
    });
    track.appendChild(ax);
    if (applyBadge) track.appendChild(applyBadge);
    var note = (!axis.drawLine && axis.rowNote) ? axis.rowNote : null;
    if (suggestion && !suggestion.separated) {
      note = "Positive and negative scores overlap. No clean threshold (check the region or tool).";
    } else if (narrowGap) {
      note = "No step-aligned threshold separates these pins (gap narrower than the slider step, or outside its range).";
    }
    return { track: track, note: note };
  }

  function _calSummary(result) {
    var tool = result.tool;
    var posTotal = 0, posPass = 0, negTotal = 0, negPass = 0, na = 0;
    result.pins.forEach(function (e) {
      var p = _calPinPass(tool, e);
      if (p == null) { na++; return; }
      if (e.polarity === "negative") { negTotal++; if (!p) negPass++; }
      else { posTotal++; if (p) posPass++; }
    });
    var text = posPass + "/" + posTotal + " positives pass · "
      + negPass + "/" + negTotal + " negatives pass";
    if (na) text += " · " + na + " not evaluable";
    var evaluableTotal = posTotal + negTotal;
    var pass = evaluableTotal > 0 && na === 0 && posPass === posTotal && negPass === negTotal;
    return { text: text, pass: pass };
  }

  // Valid threshold interval [lo, hi] over the scored pins; null = open, overlap → separated:false.
  function _calSuggest(rows, compare, scoreScale) {
    if (!compare) return null;
    var scale = scoreScale || 1;
    var pos = [], neg = [];
    rows.forEach(function (r) {
      var sc = r.sc;
      if (!sc || sc.status !== "ok" || sc.score == null || !isFinite(sc.score)) return;
      if (r.polarity === "negative") neg.push(Number(sc.score) * scale);
      else pos.push(Number(sc.score) * scale);
    });
    if (!pos.length && !neg.length) return null; // no scored pins to satisfy
    var lo = null, hi = null;
    if (compare === "le") {
      if (pos.length) lo = Math.max.apply(null, pos);
      if (neg.length) hi = Math.min.apply(null, neg);
    } else {
      if (neg.length) lo = Math.max.apply(null, neg);
      if (pos.length) hi = Math.min.apply(null, pos);
    }
    var bothPolarities = pos.length > 0 && neg.length > 0;
    if (bothPolarities && !(hi > lo)) return { separated: false };
    var mode = bothPolarities ? "gap" : (pos.length ? "positives" : "negatives");
    return {
      separated: true,
      mode: mode,
      lo: lo,
      hi: hi,
      margin: (lo != null && hi != null) ? hi - lo : null,
    };
  }

  // Decimals follow the slider step (0.83 for 0.01 steps, 13 for integers).
  function _calFmtVal(v, step) {
    var decimals = step >= 1 ? 0 : (step >= 0.1 ? 1 : 2);
    return v.toFixed(decimals);
  }

  // Inclusivity mirrors the backend comparison; `!= null` because 0 is a legal bound.
  function _calValueSatisfies(t, lo, hi, compare) {
    if (!isFinite(t)) return false;
    if (compare === "le") {
      return (lo == null || t >= lo) && (hi == null || t < hi);
    }
    return (lo == null || t > lo) && (hi == null || t <= hi);
  }

  // Nearest step-aligned value to the midpoint (or the lone finite bound); null when none fits.
  function _calApplyValue(lo, hi, compare, step, rmin, rmax) {
    function valid(t) {
      if (t < rmin || t > rmax) return false;
      return _calValueSatisfies(t, lo, hi, compare);
    }
    var target;
    if (lo != null && hi != null) target = (lo + hi) / 2;
    else if (hi != null) target = hi;
    else if (lo != null) target = lo;
    else return null; // unbounded both sides ⇒ no pins (guarded upstream)
    if (!(step > 0)) {
      // Continuous slider: clamp the target into range and verify.
      var c = Math.min(Math.max(target, rmin), rmax);
      return valid(c) ? c : null;
    }
    // Range inputs align to min + k*step; keep the valid step nearest the target.
    var best = null, bestDist = Infinity;
    var kLo = Math.floor((Math.max(lo != null ? lo : rmin, rmin) - rmin) / step) - 1;
    var kHi = Math.ceil((Math.min(hi != null ? hi : rmax, rmax) - rmin) / step) + 1;
    for (var k = kLo; k <= kHi; k++) {
      var t = parseFloat((rmin + k * step).toFixed(6));
      if (!valid(t)) continue;
      var dist = Math.abs(t - target);
      if (dist < bestDist) { bestDist = dist; best = t; }
    }
    return best;
  }

  function _calCoverageNote() {
    var consecutiveIds = ["paramChangeConsecutive", "paramTextConsecutive", "paramNumConsecutive", "paramFlowConsecutive"];
    var hasConsecutive = consecutiveIds.some(function (id) {
      var c = qs("#" + id);
      return c && parseInt(c.value, 10) > 1;
    });
    var df = qs("#paramDetectFirst");
    if (hasConsecutive || (df && df.checked)) {
      return "Consecutive / detect-first settings are not validated by calibration.";
    }
    return null;
  }

  function _calSetNote(notes) {
    var noteEl = qs("#calibrationNote");
    if (!noteEl) return;
    if (!notes.length) { noteEl.textContent = ""; noteEl.classList.add("hidden"); return; }
    noteEl.textContent = notes.join("  •  ");
    noteEl.classList.remove("hidden");
  }

  function renderCalibration() {
    var result = state.calibrationResult;
    var summaryEl = qs("#calibrationSummary");
    var stripsEl = qs("#calibrationStrips");
    if (!stripsEl) return;
    stripsEl.innerHTML = "";
    // Cleared first so stale slider marks vanish when pins or participant change.
    _calBands = {};
    if (!result || !result.pins || !result.pins.length) {
      if (summaryEl) summaryEl.textContent = "";
      _calSetNote([]);
      state.calibrationGreen = false;
      updateCalibrationSliderMarks();
      updateRunButton();
      return;
    }
    var tool = result.tool;
    var notes = [];
    var staleById = {};
    (state.pins || []).forEach(function (p) { staleById[p.id] = !!p.stale; });

    var sum = _calSummary(result);
    state.calibrationGreen = sum.pass;
    if (summaryEl) {
      summaryEl.innerHTML = "";
      summaryEl.appendChild(el("span", "cal-chip" + (sum.pass ? " cal-chip--pass" : ""), sum.text));
    }

    var frag = document.createDocumentFragment();
    if (tool === "multitool") {
      var stepCount = 0;
      result.pins.forEach(function (e) {
        if (e.steps && e.steps.length > stepCount) stepCount = e.steps.length;
      });
      for (var k = 0; k < stepCount; k++) {
        (function (k) {
          var stepType = null, logic = null;
          var rows = result.pins.map(function (e) {
            var sc = e.steps && e.steps[k] ? e.steps[k] : null;
            if (sc) { stepType = sc.type; if (k > 0) logic = sc.logic; }
            return { polarity: e.polarity, timestamp: e.timestamp, sc: sc, stale: staleById[e.pin_id] };
          });
          if (!stepType) {
            var def = state.multitoolSteps[k];
            stepType = def ? def.type : "change";
            if (k > 0 && def) logic = (def.logic || "AND").toUpperCase();
          }
          var axis = stepType === "color"
            ? _calColorAxis("_mt" + k)
            : (CAL_AXIS[stepType] || { sliderId: null, rangeMin: 0, rangeMax: 1, invert: false, drawLine: false });
          var sliderId = axis.sliderId ? axis.sliderId + "_mt" + k : null;
          var label = el("div", "cal-track-label");
          label.appendChild(el("span", null, (k + 1) + ". " + _calTitle(stepType)));
          if (k > 0 && logic) {
            label.appendChild(el("span", "cal-track-logic" + (logic === "NOT" ? " is-not" : ""), logic));
          }
          var built = _calBuildTrack(rows, stepType, axis, sliderId, label, !!sliderId);
          frag.appendChild(built.track);
          if (built.note) notes.push((k + 1) + ". " + built.note);
        })(k);
      }
    } else {
      var axis2 = tool === "color" ? _calColorAxis("") : CAL_AXIS[tool];
      if (axis2) {
        var rows2 = result.pins.map(function (e) {
          return { polarity: e.polarity, timestamp: e.timestamp, sc: e, stale: staleById[e.pin_id] };
        });
        var built2 = _calBuildTrack(rows2, tool, axis2, axis2.sliderId || null, null, true);
        frag.appendChild(built2.track);
        if (built2.note) notes.push(built2.note);
      }
    }
    stripsEl.appendChild(frag);
    var cov = _calCoverageNote();
    if (cov) notes.push(cov);
    _calSetNote(notes);
    updateCalibrationSliderMarks();
    updateRunButton();
  }

  // Glide threshold lines with the slider, no refetch; only while Preview shows them.
  function updateCalibrationThresholdLine() {
    if (state.rightPaneTab !== "preview") return;
    var lines = document.querySelectorAll("#calibrationStrips .cal-threshold[data-cal-slider]");
    Array.prototype.forEach.call(lines, function (line) {
      var sid = line.getAttribute("data-cal-slider");
      if (!sid) return;
      var slider = qs("#" + sid);
      if (!slider) return;
      var mn = parseFloat(slider.min), mx = parseFloat(slider.max), val = parseFloat(slider.value);
      if (!(mx > mn) || !isFinite(val)) return;
      var invert = line.getAttribute("data-cal-invert") === "1";
      line.style.left = _calPos(val, { min: mn, max: mx }, invert) + "%";
    });
  }

  // Readout tooltip text; a bound is often open (positives-only pins).
  function _calBandLabel(band) {
    var lo = band.lo != null ? _calFmtVal(band.lo, band.step) : null;
    var hi = band.hi != null ? _calFmtVal(band.hi, band.step) : null;
    if (lo != null && hi != null) return lo + "–" + hi;
    if (hi != null) return "up to " + hi;
    return "from " + lo;
  }

  // Hairline + pass/fail tint on each threshold control. Sweep first: stale marks outlive their cause.
  function updateCalibrationSliderMarks() {
    // Find marked nodes in the DOM; #workflowParams innerHTML rebuilds would orphan retained refs.
    qsa("#workflowParams .cal-mark, #workflowParams .cal-pass, #workflowParams .cal-fail")
      .forEach(function (node) {
        node.classList.remove("cal-mark", "cal-pass", "cal-fail");
        node.style.removeProperty("--cal-mark-frac");
        node.removeAttribute("data-tooltip");
      });
    Object.keys(_calBands).forEach(function (sliderId) {
      var band = _calBands[sliderId];
      var ctrl = qs("#" + sliderId);
      if (!ctrl) return;
      if (ctrl.type === "range" && band.max > band.min) {
        var frac = (band.applyVal - band.min) / (band.max - band.min);
        ctrl.style.setProperty("--cal-mark-frac", String(Math.min(Math.max(frac, 0), 1)));
        ctrl.classList.add("cal-mark");
      }
      // Same sibling lookup as syncValueDisplays(); spinner steps have no readout.
      var target = ctrl.parentNode && ctrl.parentNode.querySelector(".param-value");
      if (!target) target = ctrl;
      var ok = _calValueSatisfies(parseFloat(ctrl.value), band.lo, band.hi, band.compare);
      target.classList.add(ok ? "cal-pass" : "cal-fail");
      // "scored": _calSuggest skips not-evaluable pins, so green here can pair with a neutral chip.
      target.setAttribute("data-tooltip", ok
        ? "Satisfies every scored pin on this axis (valid " + _calBandLabel(band) + ")."
        : "Lets a pinned frame through — valid " + _calBandLabel(band) + ".");
    });
  }

  function _calStatus(msg, kind) {
    var statusEl = qs("#calibrationStatus");
    if (!statusEl) return;
    statusEl.classList.remove("cal-status--loading", "cal-status--error", "cg-shimmer");
    if (!msg) { statusEl.textContent = ""; statusEl.classList.add("hidden"); return; }
    statusEl.textContent = msg;
    if (kind === "loading") statusEl.classList.add("cal-status--loading", "cg-shimmer");
    else if (kind === "error") statusEl.classList.add("cal-status--error");
    statusEl.classList.remove("hidden");
  }

  // Request body via the Run path's param + region builders; {skip: reason} when not ready.
  function _calBuildBody() {
    var tool = state.activeWorkflow;
    if (!_calIsCalibratable(tool)) return { skip: "Calibration is not available for this tool." };
    var params = gatherWorkflowParams(tool, { silent: true });
    if (params === null) return { skip: "Add the missing parameters above to calibrate." };
    var body = { participant: state.selectedParticipant, tool: tool, parameters: params };
    if (tool === "multitool") {
      body.region = (params.steps && params.steps[0]) ? (params.steps[0].region || "") : "";
    } else {
      var norm = _previewRegionRef();
      if (!norm && !((tool === "template" || tool === "shape") && state.uploadedTemplate)) {
        return { skip: "Select a region to calibrate." };
      }
      body.region = norm ? norm.name : "";
      if (norm) body.region_ref = regionRefPayload(norm);
    }
    return { body: body };
  }

  function _doRefreshCalibration() {
    var gen = ++_calibrationGen;
    var pid = state.selectedParticipant;
    if (!pid || !state.pins || !state.pins.length) {
      state.calibrationResult = null;
      renderCalibration();
      _calStatus("");
      return;
    }
    var built = _calBuildBody();
    if (built.skip) {
      state.calibrationResult = null;
      renderCalibration();
      _calStatus(built.skip);
      return;
    }
    var tool = built.body.tool;
    var needsOcr = tool === "text" || tool === "numbers"
      || (tool === "multitool" && (built.body.parameters.steps || []).some(function (s) {
        return s.type === "text" || s.type === "numbers";
      }));
    // Flag cold-start OCR; otherwise "Evaluating…" only on first load, so re-evals don't flicker.
    if (needsOcr && !state.calibrationOcrWarmed) {
      _calStatus("Preparing OCR… (first run only)", "loading");
    } else if (!state.calibrationResult) {
      _calStatus("Evaluating…", "loading");
    }
    apiPost("api/calibrate", built.body)
      .then(function (data) {
        // Drop stale responses: superseded gen or a participant switched away from.
        if (gen !== _calibrationGen || pid !== state.selectedParticipant) return;
        if (!data || !data.ok) {
          state.calibrationResult = null;
          renderCalibration();
          _calStatus("Calibration unavailable.", "error");
          return;
        }
        if (needsOcr) state.calibrationOcrWarmed = true;
        state.calibrationResult = data;
        renderCalibration();
        _calStatus("");
      })
      .catch(function () {
        if (gen !== _calibrationGen || pid !== state.selectedParticipant) return;
        // Clear the now-stale dots so an error can't be read as current scores.
        state.calibrationResult = null;
        renderCalibration();
        _calStatus("Calibration unavailable.", "error");
      });
  }

  function refreshCalibration(opts) {
    // Not gated on the Preview tab: the Run hint and slider hairlines need it everywhere.
    if (_calibrationTimer) { clearTimeout(_calibrationTimer); _calibrationTimer = 0; }
    if (opts && opts.debounce) {
      _calibrationTimer = setTimeout(_doRefreshCalibration, 150);
    } else {
      _doRefreshCalibration();
    }
  }

  // No pins: show the "pin some frames" prompt; the section header stays.
  function updateCalibrationVisibility() {
    var body = qs("#calibrationBody");
    var empty = qs("#calibrationEmpty");
    var hasPins = !!(state.pins && state.pins.length);
    if (body) body.classList.toggle("hidden", !hasPins);
    if (empty) empty.classList.toggle("hidden", hasPins);
  }

  function initCalibration() {
    // Delegated listeners on the stable containers catch every param input, not just sliders.
    ["workflowParams", "workflowIntervalSlot"].forEach(function (id) {
      var container = qs("#" + id);
      if (!container) return;
      var handler = function () {
        updateCalibrationThresholdLine();
        // Re-tint from the cached interval; scores are threshold-independent, so no refetch needed.
        updateCalibrationSliderMarks();
        refreshCalibration({ debounce: true });
      };
      container.addEventListener("input", handler);
      container.addEventListener("change", handler);
    });
  }

  // ---- Published to the hub (its delegators forward here) ----
  SS.calRefresh = refreshCalibration;
  SS.calUpdateThresholdLine = updateCalibrationThresholdLine;
  SS.calRender = renderCalibration;
  SS.calVisibility = updateCalibrationVisibility;
  SS.calInit = initCalibration;
  // The hub's selectParticipant bumps this to invalidate in-flight responses.
  SS.calBumpGen = function () { _calibrationGen += 1; };
})();
