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
  //
  // Scores the participant's pins against the active tool + parameters, plotting
  // each as a dot on a normalized 0–1 "matchiness" axis (green = positive pin, red
  // = negative), with the threshold drawn as a vertical line so the researcher can
  // place the cutoff in the gap between populations. Evaluation is per-frame only
  // — temporal params are not validated, and the coverage note says so. Unlike the
  // model view it scores the PINNED timestamps, so it never re-runs on seek.

  var _calibrationGen = 0;
  var _calibrationTimer = 0;
  // Per-slider pin verdict, keyed by the threshold control's element id:
  // {lo, hi, compare, applyVal, min, max, step}. Written by _calBuildTrack (only
  // when a step-aligned cutoff exists), read by updateCalibrationSliderMarks to
  // annotate the control itself. Rebuilt from scratch on every renderCalibration
  // — which every param-panel rebuild funnels through — so an entry can never
  // outlive the slider it describes.
  var _calBands = {};
  // Set while restoreTaskToWorkflow() rebuilds the param panel: the
  // renderWorkflowParams() call there fires before the saved values are
  // written, so its calibration re-eval would POST default params only to be
  // immediately superseded. Suppress it; restore runs its own refresh at the
  // end with the real values.

  // Per-tool axis metadata. The range comes from the tool's threshold slider where
  // one exists in matching units; color and scene have no single clean cutoff, so
  // they draw no line and rely on per-pin pass/fail plus the population gap.
  // inactivity's Sensitivity slider is already in phash-distance units, so its line
  // is drawn inverted (low distance = more inactive = right). Per-pin scalars come
  // from the backend `score`; this map only positions them.
  //
  // `compare` is the *pass comparison* ("ge" passes when score ≥ threshold), used
  // by _calSuggest to bisect the gap. Distinct from `invert`, which only flips the
  // axis *display* — they coincide today but mean different things.
  var CAL_AXIS = {
    change: { sliderId: "paramChangeThresh", invert: false, drawLine: true, compare: "ge" },
    similarity: { sliderId: "paramSimThresh", invert: false, drawLine: true, compare: "ge" },
    text: { sliderId: "paramTextFuzzy", invert: false, drawLine: true, compare: "ge" },
    template: { sliderId: "paramTemplateThresh", invert: false, drawLine: true, compare: "ge" },
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

  // The color axis is mode-dependent. In presence mode the score IS the matching
  // pixel coverage and the Min-area slider IS the cutoff, so it becomes a normal
  // draw-line + suggest axis. `scoreScale` (100) maps coverage (0–1) onto the
  // slider's percent units (0–100) so dots, line and Apply all share one scale.
  // `sfx` is "" (single tool) or "_mt{k}"; the base sliderId gets the suffix
  // appended by the caller, like every other tool. Average mode keeps the static
  // no-line descriptor.
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

  // Axis range from the live slider DOM (the control the user is moving) or the
  // descriptor fallback when there is no matching-units slider (color / scene).
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

  // "Nice" value ticks (1/2/5 × 10ⁿ) across [min, max], aiming for ~5 intervals.
  // Returns { ticks, step }; empty when the range is degenerate.
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

  // Faint value grid behind a track's dots: a hairline + value label per tick,
  // positioned with the same value→percent mapping (and invert) as the dots, so
  // absolute spacing is legible (not just relative).
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

  // Build one track (axis + dots + optional threshold line). `rows` is
  // [{polarity, timestamp, sc, stale}] where `sc` is the score object for this
  // track (the pin entry for a single tool, or entry.steps[k] for multitool).
  // `suggest` enables the dashed midpoint marker + "Apply" badge — every track
  // whose tool has a threshold slider, single or multitool step.
  function _calBuildTrack(rows, tool, axis, sliderId, label, suggest) {
    var track = el("div", "cal-track");
    if (label) track.appendChild(label);
    // Factor mapping the backend score into the axis/slider units (1 for tools
    // whose score already matches the slider; 100 for color presence coverage).
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
    // Suggested cutoff (step-aligned): midway through the gap when both
    // polarities are pinned, otherwise hugging the edge of the pinned cluster
    // (positives-only is a primary workflow). Drawn behind the dots (appended
    // before them) and re-derived on every render, so it stays put while the
    // threshold slider is dragged (these tools' scores are threshold-independent).
    var suggestion = (suggest && axis.drawLine && sliderId) ? _calSuggest(rows, axis.compare, scoreScale) : null;
    var applyBadge = null;
    var narrowGap = false;
    if (suggestion && suggestion.separated) {
      var slider = qs("#" + sliderId);
      var step = slider ? parseFloat(slider.step) : 0;
      if (!isFinite(step)) step = 0;
      // Step-aligned cutoff that actually satisfies the pins (a raw target can
      // snap onto a boundary and still let a pin through).
      var applyVal = _calApplyValue(suggestion.lo, suggestion.hi, axis.compare, step, range.min, range.max);
      if (applyVal == null) {
        narrowGap = true; // valid interval exists but no step-aligned value lands in it
      } else {
        // Same gate as the Apply badge, so the mark drawn on the slider itself
        // (updateCalibrationSliderMarks) and the panel always agree: a pin
        // verdict is recorded only when a reachable cutoff exists.
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
      // Fan dots that land on the same spot upward so each stays hoverable;
      // cap the stack so a tight cluster (e.g. several pins at SSIM 1.0) doesn't
      // overflow the axis into the row above.
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

  // Suggest a cutoff that satisfies whatever pins exist — pinning only expected
  // positives is a primary workflow, so this does NOT require negatives. It
  // returns the *valid threshold interval* as bounds [lo, hi] (either may be null
  // = unbounded), which the populations constrain:
  //   ge (pass when score >= T): positives need T <= min(pos) (hi, inclusive);
  //      negatives need T > max(neg) (lo, exclusive).
  //   le (pass when score <= T): positives need T >= max(pos) (lo, inclusive);
  //      negatives need T < min(neg) (hi, exclusive).
  // `mode` records what the caller is working with so it can word the Apply hint.
  // With both polarities the interval can be empty → {separated:false} (the
  // use-case-#2 "overlap / wrong region or tool" signal). `compare` is the pass
  // comparison, independent of axis rendering.
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

  // Decimals follow the slider step so suggested values read cleanly (0.83 for
  // 0.01-step tools, 13 for inactivity's integer step).
  function _calFmtVal(v, step) {
    var decimals = step >= 1 ? 0 : (step >= 0.1 ? 1 : 2);
    return v.toFixed(decimals);
  }

  // Does threshold `t` fall inside the valid interval, i.e. satisfy every scored
  // pin? Either bound may be null (unbounded). The inclusivity is asymmetric by
  // direction and mirrors the backend's own comparison, so the slider tint and
  // the per-pin dots can never disagree. `!= null` throughout: 0 is a legitimate
  // bound (inactivity and color min-area both start there).
  function _calValueSatisfies(t, lo, hi, compare) {
    if (!isFinite(t)) return false;
    if (compare === "le") {
      return (lo == null || t >= lo) && (hi == null || t < hi);
    }
    return (lo == null || t > lo) && (hi == null || t <= hi);
  }

  // Pick a step-aligned threshold inside the valid interval [lo, hi] (either
  // bound may be null = unbounded). The interval is asymmetric by direction —
  // ge needs lo < T <= hi, le needs lo <= T < hi — so the plain midpoint can
  // round onto a boundary and still let a pin through; we verify alignment.
  // Target: the midpoint when both bounds are finite (a true gap); the finite
  // bound itself when only one polarity is pinned, so the cutoff hugs that
  // cluster's edge (loosest threshold that still satisfies every pin). Returns
  // null when no aligned value fits — gap narrower than the step, or the whole
  // interval sits outside the slider range — so the caller shows a note instead
  // of an unsafe "Apply".
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
    // Range inputs align to min + k*step; scan the steps spanning the interval
    // (bounded by the slider range when a side is open) and keep the valid one
    // nearest the target.
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
    // Re-derived below by _calBuildTrack; clearing first is what un-marks the
    // sliders when the pins, the participant, or the fetch outcome change.
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

  // Glide the threshold line(s) with the slider without refetching scores.
  function updateCalibrationThresholdLine() {
    if (!state.calibrationOpen) return;
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

  // Human-readable valid interval for the readout tooltip. Either bound may be
  // open, which is the common case when only positives are pinned.
  function _calBandLabel(band) {
    var lo = band.lo != null ? _calFmtVal(band.lo, band.step) : null;
    var hi = band.hi != null ? _calFmtVal(band.hi, band.step) : null;
    if (lo != null && hi != null) return lo + "–" + hi;
    if (hi != null) return "up to " + hi;
    return "from " + lo;
  }

  // Annotate each threshold control with what the pins say: a hairline on the
  // slider track at the suggested cutoff, and a green/red tint on the value
  // readout for whether the *current* value satisfies every scored pin. Runs
  // whether or not the strip is expanded — the panel ships collapsed, so this is
  // the only calibration signal most users see while tuning.
  //
  // Sweep-then-apply, because the paths that invalidate a mark leave the slider
  // DOM in place: participant switch, clear-pins, a failed fetch, or a pin that
  // flips a clean gap into an overlap. Only a param-panel rebuild disposes of
  // the controls for us — and multitool's step ids are positional, so after a
  // delete-and-reindex the same id can belong to a different step.
  function updateCalibrationSliderMarks() {
    // Previously-marked nodes are found in the DOM rather than tracked in a JS
    // list: #workflowParams is rebuilt by innerHTML at arbitrary times, which
    // would leave any retained element references pointing at detached nodes.
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
      // Same sibling lookup as syncValueDisplays(); multitool's number-spinner
      // steps have no readout, so the tint lands on the input's own text.
      var target = ctrl.parentNode && ctrl.parentNode.querySelector(".param-value");
      if (!target) target = ctrl;
      var ok = _calValueSatisfies(parseFloat(ctrl.value), band.lo, band.hi, band.compare);
      target.classList.add(ok ? "cal-pass" : "cal-fail");
      // "scored" is load-bearing: _calSuggest ignores not-evaluable pins, so a
      // green readout here can coexist with a neutral summary chip (which
      // requires every pin to be evaluable).
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

  // Build the calibrate request body, reusing the Run path's param + region
  // construction. Returns {skip: reason} when the tool isn't ready (shown
  // inline instead of toasting on every keystroke).
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
      if (!norm && !(tool === "template" && state.uploadedTemplate)) {
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
    // Cold-start OCR is the only slow case worth flagging; otherwise show a
    // brief "Evaluating…" only on the first load so debounced re-evals (which
    // keep the prior dots visible) don't flicker the status line.
    if (needsOcr && !state.calibrationOcrWarmed) {
      _calStatus("Preparing OCR… (first run only)", "loading");
    } else if (!state.calibrationResult) {
      _calStatus("Evaluating…", "loading");
    }
    apiPost("api/calibrate", built.body)
      .then(function (data) {
        // Reject stale responses: a superseded refresh (gen) or a response for
        // a participant we've since switched away from must not overwrite the
        // strip.
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
    // Not gated on the panel being open: the Run "Calibrated" hint must reflect
    // pin agreement whenever pins exist, even while the strip is collapsed.
    // _doRefreshCalibration() no-ops cheaply (no POST) when there are no pins,
    // so there's nothing to skip when the participant has none.
    if (_calibrationTimer) { clearTimeout(_calibrationTimer); _calibrationTimer = 0; }
    if (opts && opts.debounce) {
      _calibrationTimer = setTimeout(_doRefreshCalibration, 150);
    } else {
      _doRefreshCalibration();
    }
  }

  // Hide the whole panel when the participant has no pins (distinct from the
  // collapsed open/close state).
  function updateCalibrationVisibility() {
    var panel = qs("#calibrationPanel");
    if (!panel) return;
    panel.classList.toggle("hidden", !(state.pins && state.pins.length));
  }

  function toggleCalibration() {
    state.calibrationOpen = !state.calibrationOpen;
    var panel = qs("#calibrationPanel");
    var body = qs("#calibrationBody");
    var btn = qs("#calibrationToggle");
    if (state.calibrationOpen) {
      panel.classList.remove("collapsed");
      body.classList.remove("hidden");
      btn.setAttribute("aria-expanded", "true");
      refreshCalibration();
    } else {
      panel.classList.add("collapsed");
      body.classList.add("hidden");
      btn.setAttribute("aria-expanded", "false");
    }
  }

  function initCalibration() {
    var btn = qs("#calibrationToggle");
    if (btn) btn.addEventListener("click", toggleCalibration);
    // Delegated listeners on the stable param containers: any control change
    // glides the threshold line (sync) and re-evaluates scores (debounced).
    // This catches every param input — sliders, text, selects, checkboxes —
    // including ones (search string, operator) that the model view ignores.
    ["workflowParams", "workflowIntervalSlot"].forEach(function (id) {
      var container = qs("#" + id);
      if (!container) return;
      var handler = function () {
        updateCalibrationThresholdLine();
        // Re-tint from the cached interval so the readout flips the instant the
        // value crosses the boundary, rather than after the debounced refetch.
        // The scores these tools produce are threshold-independent, so the
        // client-side verdict is exact, not an approximation.
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
  // selectParticipant (hub) bumps the generation counter to invalidate any
  // in-flight calibration response when the participant changes.
  SS.calBumpGen = function () { _calibrationGen += 1; };
})();
