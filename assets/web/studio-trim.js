/* clipgen Studio — duration-badge trim satellite.
 *
 * Carved out of studio.js (the hub) following the hub+satellite convention used
 * by studio-scrubber.js / studio-intake.js. Owns the queue-card duration badge
 * and its three-row trim pop-over (drag / ±30s / type-to-edit the clip in-out
 * points), plus buildCellOverrides() which serializes user edits into the
 * per-cell override payload the generate/reel routes send.
 *
 * Cross-file contract: the satellite reaches two hub helpers through the
 * window.ClipgenStudio (STUDIO) namespace — saveQueues (persist queues after an
 * edit) and isIntakeSource (filter intake items out of cell overrides); both are
 * destructured at load time, which is safe because this file loads after the hub
 * IIFE has published them. el / formatDuration / formatTime / parseTimestamp are
 * ambient utils.js globals reached via the scope chain. The hub calls back in via
 * same-named guarded delegators (appendDurationBadge, buildCellOverrides). Loaded
 * by studio.html after studio.js; order relative to the other satellites is free
 * (no cross-destructuring).
 */

(function () {
  "use strict";

  var STUDIO = window.ClipgenStudio;
  var saveQueues = STUDIO.saveQueues,
    isIntakeSource = STUDIO.isIntakeSource;

  // ---- Duration-badge trim pop-over -------------------------------------
  // Clicking a queue card's duration badge opens a small three-row pop-over
  // (same dark "badge" styling) for adjusting the clip's in/out points:
  //   Row 1  total duration — drag horizontally to grow/shrink symmetrically
  //   Row 2  in / out times — drag each marker independently; click to type
  //   Row 3  ±30s quick buttons for the front and back
  // Edits mutate the queue item's start/end (seconds) and set item.edited, so
  // generation sends them as overrides (spreadsheet cells) or directly (intake).
  var TRIM_SECONDS_PER_PX = 0.2; // drag sensitivity (seconds per pixel)
  var TRIM_MIN_CLIP = 1; // shortest allowed clip, seconds
  var TRIM_STEP = 30; // ±30s quick-button step
  var TRIM_DRAG_THRESHOLD = 3; // px of movement before a press counts as a drag
  var activeTrim = null;

  function closeTrimPopover() {
    if (!activeTrim) return;
    var t = activeTrim;
    activeTrim = null;
    document.removeEventListener("pointerdown", t.onDocDown, true);
    document.removeEventListener("keydown", t.onKey, true);
    window.removeEventListener("scroll", t.onDismiss, true);
    window.removeEventListener("resize", t.onDismiss, true);
    if (t.popover && t.popover.parentNode) {
      t.popover.parentNode.removeChild(t.popover);
    }
    // Re-render the queue so derived totals (e.g. the reel duration in the
    // toolbar) reflect the new in/out points — but only when something actually
    // changed, so merely opening and dismissing the pop-over is cheap.
    if (t.dirty && t.renderFn) t.renderFn();
  }

  function positionTrimPopover(popover, anchorRect) {
    // Right-align to the badge and grow up/left, clamped to the viewport.
    var w = popover.offsetWidth;
    var h = popover.offsetHeight;
    var left = anchorRect.right - w;
    if (left + w > window.innerWidth - 4) left = window.innerWidth - w - 4;
    if (left < 4) left = 4;
    var top = anchorRect.bottom - h;
    if (top + h > window.innerHeight - 4) top = window.innerHeight - h - 4;
    if (top < 4) top = 4;
    popover.style.left = left + "px";
    popover.style.top = top + "px";
  }

  // Horizontal drag-to-adjust. handlers: { onStart(): base, onDelta(sec, base),
  // onClick() }. A press that never crosses TRIM_DRAG_THRESHOLD is treated as a
  // click (so the in/out values can switch to manual numeric entry).
  function bindTrimDrag(target, handlers) {
    target.addEventListener("pointerdown", function (ev) {
      if (ev.button !== 0) return;
      ev.preventDefault();
      ev.stopPropagation();
      var originX = ev.clientX;
      var base = handlers.onStart ? handlers.onStart() : null;
      var dragged = false;
      var rafPending = false;
      var lastDelta = 0;
      try {
        target.setPointerCapture(ev.pointerId);
      } catch (e) {
        /* pointer capture is best-effort */
      }

      function onMove(e) {
        var dx = e.clientX - originX;
        if (!dragged && Math.abs(dx) < TRIM_DRAG_THRESHOLD) return;
        dragged = true;
        lastDelta = dx * TRIM_SECONDS_PER_PX;
        if (rafPending) return;
        rafPending = true;
        requestAnimationFrame(function () {
          rafPending = false;
          if (handlers.onDelta) handlers.onDelta(lastDelta, base);
        });
      }
      function onUp() {
        target.removeEventListener("pointermove", onMove);
        target.removeEventListener("pointerup", onUp);
        target.removeEventListener("pointercancel", onUp);
        if (!dragged) {
          if (handlers.onClick) handlers.onClick();
        } else {
          saveQueues();
        }
      }
      target.addEventListener("pointermove", onMove);
      target.addEventListener("pointerup", onUp);
      target.addEventListener("pointercancel", onUp);
    });
  }

  function makeTrimButton(label, title, onClick) {
    var b = el("button", "trim-add-btn", label);
    b.type = "button";
    b.title = title;
    b.addEventListener("pointerdown", function (ev) {
      ev.stopPropagation();
    });
    b.addEventListener("click", function (ev) {
      ev.preventDefault();
      ev.stopPropagation();
      onClick();
    });
    return b;
  }

  function openTrimPopover(badge, item, renderFn) {
    closeTrimPopover();

    var pop = el("div", "trim-popover");

    // Row 1 — total duration (drag to resize both ends symmetrically).
    var rowDur = el("div", "trim-row trim-row-duration");
    rowDur.appendChild(el("span", "trim-row-label", "Length"));
    var durVal = el("span", "trim-duration-value", formatDuration(item.end - item.start));
    durVal.title = "Drag to lengthen / shorten";
    rowDur.appendChild(durVal);

    // Row 2 — in / out points (drag each; click to type).
    var rowInOut = el("div", "trim-row trim-row-inout");
    var inVal = el("span", "trim-time trim-time-in", formatTime(item.start));
    var outVal = el("span", "trim-time trim-time-out", formatTime(item.end));
    inVal.title = "Drag to move the in-point · click to type";
    outVal.title = "Drag to move the out-point · click to type";
    rowInOut.appendChild(inVal);
    rowInOut.appendChild(el("span", "trim-inout-sep", "→"));
    rowInOut.appendChild(outVal);

    // Row 3 — ±30s quick buttons for the front and back.
    var rowBtns = el("div", "trim-row trim-row-buttons");
    var frontGroup = el("div", "trim-btn-group");
    frontGroup.appendChild(
      makeTrimButton("−" + TRIM_STEP, "Trim " + TRIM_STEP + "s off the front", function () {
        setTimes(item.start + TRIM_STEP, item.end, false);
      })
    );
    frontGroup.appendChild(el("span", "trim-btn-group-label", "front"));
    frontGroup.appendChild(
      makeTrimButton("+" + TRIM_STEP, "Add " + TRIM_STEP + "s to the front", function () {
        setTimes(item.start - TRIM_STEP, item.end, false);
      })
    );
    var backGroup = el("div", "trim-btn-group");
    backGroup.appendChild(
      makeTrimButton("−" + TRIM_STEP, "Trim " + TRIM_STEP + "s off the back", function () {
        setTimes(item.start, item.end - TRIM_STEP, false);
      })
    );
    backGroup.appendChild(el("span", "trim-btn-group-label", "back"));
    backGroup.appendChild(
      makeTrimButton("+" + TRIM_STEP, "Add " + TRIM_STEP + "s to the back", function () {
        setTimes(item.start, item.end + TRIM_STEP, false);
      })
    );
    rowBtns.appendChild(frontGroup);
    rowBtns.appendChild(backGroup);

    pop.appendChild(rowDur);
    pop.appendChild(rowInOut);
    pop.appendChild(rowBtns);
    document.body.appendChild(pop);
    positionTrimPopover(pop, badge.getBoundingClientRect());

    function refreshTexts() {
      var d = formatDuration(item.end - item.start);
      durVal.textContent = d;
      inVal.textContent = formatTime(item.start);
      outVal.textContent = formatTime(item.end);
      badge.textContent = d;
    }

    // Clamp + apply new in/out points. skipSave defers the sessionStorage write
    // to the drag's pointerup so we don't write on every animation frame.
    function setTimes(newStart, newEnd, skipSave) {
      newStart = Math.max(0, Math.round(newStart));
      newEnd = Math.round(newEnd);
      if (newEnd < newStart + TRIM_MIN_CLIP) newEnd = newStart + TRIM_MIN_CLIP;
      item.start = newStart;
      item.end = newEnd;
      item.edited = true;
      if (activeTrim) activeTrim.dirty = true;
      refreshTexts();
      if (!skipSave) saveQueues();
    }

    function startNumericEntry(span, which) {
      var input = document.createElement("input");
      input.type = "text";
      input.className = "trim-time-input";
      input.autocomplete = "off";
      input.value = formatTime(which === "in" ? item.start : item.end);
      var done = false;
      function commit() {
        if (done) return;
        done = true;
        var sec = parseTimestamp(input.value);
        if (sec != null && isFinite(sec)) {
          if (which === "in") {
            var ns = Math.min(Math.max(0, sec), item.end - TRIM_MIN_CLIP);
            setTimes(ns, item.end, false);
          } else {
            setTimes(item.start, Math.max(item.start + TRIM_MIN_CLIP, sec), false);
          }
        }
        if (input.parentNode) input.parentNode.replaceChild(span, input);
        refreshTexts();
      }
      input.addEventListener("blur", commit);
      input.addEventListener("keydown", function (ev) {
        ev.stopPropagation();
        if (ev.key === "Enter") {
          ev.preventDefault();
          input.blur();
        } else if (ev.key === "Escape") {
          ev.preventDefault();
          done = true;
          if (input.parentNode) input.parentNode.replaceChild(span, input);
        }
      });
      input.addEventListener("pointerdown", function (ev) {
        ev.stopPropagation();
      });
      span.parentNode.replaceChild(input, span);
      input.focus();
      input.select();
    }

    bindTrimDrag(durVal, {
      onStart: function () {
        return { start: item.start, end: item.end };
      },
      onDelta: function (deltaSec, base) {
        // Grow/shrink equally; clamp the front at 0 (the back keeps moving).
        var half = deltaSec / 2;
        var ns = base.start - half;
        if (ns < 0) ns = 0;
        setTimes(ns, base.end + half, true);
      },
    });
    bindTrimDrag(inVal, {
      onStart: function () {
        return item.start;
      },
      onDelta: function (deltaSec, base) {
        setTimes(base + deltaSec, item.end, true);
      },
      onClick: function () {
        startNumericEntry(inVal, "in");
      },
    });
    bindTrimDrag(outVal, {
      onStart: function () {
        return item.end;
      },
      onDelta: function (deltaSec, base) {
        setTimes(item.start, base + deltaSec, true);
      },
      onClick: function () {
        startNumericEntry(outVal, "out");
      },
    });

    var onDocDown = function (ev) {
      if (pop.contains(ev.target)) return;
      closeTrimPopover();
    };
    var onKey = function (ev) {
      if (ev.key === "Escape") {
        ev.preventDefault();
        closeTrimPopover();
      }
    };
    var onDismiss = function () {
      closeTrimPopover();
    };
    document.addEventListener("pointerdown", onDocDown, true);
    document.addEventListener("keydown", onKey, true);
    window.addEventListener("scroll", onDismiss, true);
    window.addEventListener("resize", onDismiss, true);

    activeTrim = {
      popover: pop,
      renderFn: renderFn,
      onDocDown: onDocDown,
      onKey: onKey,
      onDismiss: onDismiss,
    };
  }

  // Build the duration badge as an editable trigger for the trim pop-over.
  // Used by both the artifact and reel queue renderers.
  function appendDurationBadge(thumb, item, renderFn) {
    var badge = el(
      "span",
      "queue-card-duration queue-card-duration--editable",
      formatDuration(item.end - item.start)
    );
    badge.title = "Adjust clip length";
    badge.addEventListener("pointerdown", function (ev) {
      ev.stopPropagation();
    });
    badge.addEventListener("dragstart", function (ev) {
      ev.preventDefault();
      ev.stopPropagation();
    });
    badge.addEventListener("click", function (ev) {
      ev.preventDefault();
      ev.stopPropagation();
      openTrimPopover(badge, item, renderFn);
    });
    thumb.appendChild(badge);
  }

  // Collect per-cell time overrides for spreadsheet clips the user trimmed on
  // the duration badge or pruned from the queue. The backend replaces a cell's
  // whole time list, so whenever a cell has an edited segment OR has had cards
  // removed (fewer queued segments than its original segTotal) we send every
  // remaining segment (segIdx-ordered) as [startSec, endSec] pairs — that
  // becomes the cell's complete output. Returns {} when a cell is untouched
  // (no edits, no removals) so artifact caching still applies. Intake items
  // carry their own start/end and are skipped here.
  function buildCellOverrides(items) {
    var byCell = {};
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      if (isIntakeSource(it.source)) continue;
      var key = it.participant + "." + it.row;
      if (!byCell[key]) byCell[key] = [];
      byCell[key].push(it);
    }
    var overrides = {};
    Object.keys(byCell).forEach(function (key) {
      var segs = byCell[key];
      var anyEdited = false;
      for (var s = 0; s < segs.length; s++) {
        if (segs[s].edited) {
          anyEdited = true;
          break;
        }
      }
      var segTotal = segs[0].segTotal || segs.length;
      var removed = segs.length < segTotal;
      if (!anyEdited && !removed) return;
      segs.sort(function (a, b) {
        return (a.segIdx || 0) - (b.segIdx || 0);
      });
      overrides[key] = segs.map(function (seg) {
        return [Math.round(seg.start), Math.round(seg.end)];
      });
    });
    return overrides;
  }

  STUDIO.appendDurationBadge = appendDurationBadge;
  STUDIO.buildCellOverrides = buildCellOverrides;
})();
