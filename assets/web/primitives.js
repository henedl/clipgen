/* clipgen design primitives — factory helpers shared across Studio,
 * Screenspace, and Transcripts. Each factory returns a DOM element
 * (not an HTML string) so callers can wire event handlers cleanly and
 * appending hundreds of cards in a frame stays cheap.
 *
 * Hue resolution falls back to categoryHue(label) from utils.js when
 * the caller omits an explicit `hue`.
 *
 * Surface (window.ClipgenPrimitives):
 *   createFilterChip, createParticipantPill, createDensityTimeline,
 *   createSparkBars, createClipCard, createTranscriptCard, createBtn
 */

(function (global) {
  function setDataset(el, dataset) {
    if (!dataset) return;
    Object.keys(dataset).forEach(function (k) {
      var v = dataset[k];
      if (v == null) return;
      el.dataset[k] = String(v);
    });
  }

  function resolveHue(label, hue) {
    if (typeof hue === "number") return hue;
    if (typeof global.categoryHue === "function") {
      return global.categoryHue(label || "");
    }
    return 220;
  }

  function fmtHue(hue, lightness, chroma, alpha) {
    var l = lightness != null ? lightness : 0.7;
    var c = chroma != null ? chroma : 0.16;
    if (alpha == null || alpha >= 1) {
      return "oklch(" + l + " " + c + " " + hue + ")";
    }
    return "oklch(" + l + " " + c + " " + hue + " / " + alpha + ")";
  }

  // ---- FilterChip ----

  function createFilterChip(opts) {
    opts = opts || {};
    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "filter-chip";
    var hue = resolveHue(opts.label, opts.hue);

    if (opts.dot !== false) {
      var dot = document.createElement("span");
      dot.className = "filter-chip-dot";
      dot.style.color = fmtHue(hue);
      btn.appendChild(dot);
    }

    var labelEl = document.createElement("span");
    labelEl.className = "filter-chip-label";
    labelEl.textContent = opts.label || "";
    btn.appendChild(labelEl);

    if (opts.count != null) {
      var countEl = document.createElement("span");
      countEl.className = "filter-chip-count";
      countEl.textContent = String(opts.count);
      btn.appendChild(countEl);
    }

    if (opts.active) {
      btn.classList.add("is-active");
      btn.style.setProperty("--cg-chip-fg", fmtHue(hue));
      btn.style.setProperty("--cg-chip-bg", fmtHue(hue, 0.7, 0.16, 0.12));
      btn.style.setProperty("--cg-chip-border", fmtHue(hue, 0.7, 0.16, 0.45));
    }

    setDataset(btn, opts.dataset);
    if (typeof opts.onClick === "function") {
      btn.addEventListener("click", opts.onClick);
    }
    return btn;
  }

  // ---- ParticipantPill ----

  function createParticipantPill(opts) {
    opts = opts || {};
    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "participant-pill cg-mono";
    btn.textContent = opts.id || "";
    if (opts.active) btn.classList.add("is-active");
    setDataset(btn, opts.dataset);
    if (typeof opts.onClick === "function") {
      btn.addEventListener("click", opts.onClick);
    }
    return btn;
  }

  // ---- DensityTimeline ----
  //
  // events: [{ t: 0..1, count, hue?, label? }, ...]
  // marker: 0..1 (optional), tickCount: number of tick labels (default 6),
  // durationSec: number — used to format tick labels as M:SS / H:MM:SS.

  function fmtTick(sec) {
    var s = Math.max(0, Math.floor(sec));
    var h = Math.floor(s / 3600);
    var m = Math.floor((s - h * 3600) / 60);
    var rs = s - h * 3600 - m * 60;
    if (h > 0) {
      return h + ":" + (m < 10 ? "0" + m : m) + ":" + (rs < 10 ? "0" + rs : rs);
    }
    return m + ":" + (rs < 10 ? "0" + rs : rs);
  }

  function createDensityTimeline(opts) {
    opts = opts || {};
    var wrap = document.createElement("div");
    wrap.className = "density-timeline";
    if (opts.height) wrap.style.setProperty("--cg-density-h", opts.height + "px");

    var ticks = document.createElement("div");
    ticks.className = "density-timeline-ticks cg-mono";
    wrap.appendChild(ticks);

    var track = document.createElement("div");
    track.className = "density-timeline-track";
    wrap.appendChild(track);

    function renderTicks(durationSec, tickCount) {
      ticks.textContent = "";
      var n = Math.max(2, tickCount || 6);
      for (var i = 0; i < n; i++) {
        var s = (durationSec || 0) * (i / (n - 1));
        var t = document.createElement("span");
        t.textContent = fmtTick(s);
        ticks.appendChild(t);
      }
    }

    function renderBars(events, marker) {
      track.textContent = "";
      var max = 1;
      (events || []).forEach(function (e) {
        if (e && typeof e.count === "number" && e.count > max) max = e.count;
      });
      (events || []).forEach(function (e, idx) {
        if (!e || typeof e.t !== "number") return;
        var bar = document.createElement("div");
        bar.className = "density-timeline-bar";
        var hue = resolveHue(e.label, e.hue);
        var alpha = Math.min(1, 0.35 + 0.6 * ((e.count || 1) / max));
        bar.style.left = "calc(" + (e.t * 100) + "% - 3px)";
        bar.style.background = fmtHue(hue, 0.7, 0.16, alpha);
        bar.dataset.idx = idx;
        if (typeof opts.onBarMouseEnter === "function") {
          bar.addEventListener("mouseenter", function () { opts.onBarMouseEnter(idx); });
        }
        if (typeof opts.onBarMouseLeave === "function") {
          bar.addEventListener("mouseleave", function () { opts.onBarMouseLeave(idx); });
        }
        if (typeof opts.onBarClick === "function") {
          bar.addEventListener("click", function (ev) { opts.onBarClick(idx, ev); });
        }
        track.appendChild(bar);
      });
      if (marker != null) {
        var mk = document.createElement("div");
        mk.className = "density-timeline-marker";
        mk.style.left = "calc(" + (marker * 100) + "% - 1px)";
        track.appendChild(mk);
      }
    }

    wrap.update = function (events, marker, durationSec, tickCount) {
      if (durationSec != null || tickCount != null) {
        renderTicks(durationSec != null ? durationSec : opts.durationSec, tickCount != null ? tickCount : opts.tickCount);
      }
      renderBars(events, marker);
    };

    wrap.setHovered = function (idx) {
      var bars = track.querySelectorAll(".density-timeline-bar");
      for (var i = 0; i < bars.length; i++) {
        var b = bars[i];
        if (idx == null || idx === -1) {
          b.classList.remove("is-hover", "is-dim");
        } else if (parseInt(b.dataset.idx, 10) === idx) {
          b.classList.add("is-hover");
          b.classList.remove("is-dim");
        } else {
          b.classList.add("is-dim");
          b.classList.remove("is-hover");
        }
      }
    };

    renderTicks(opts.durationSec, opts.tickCount);
    renderBars(opts.events, opts.marker);
    return wrap;
  }

  // ---- SparkBars ----

  function createSparkBars(opts) {
    opts = opts || {};
    var wrap = document.createElement("div");
    wrap.className = "spark-bars";
    if (opts.height) wrap.style.height = opts.height + "px";
    var data = opts.data || [];
    var hue = resolveHue(null, opts.hue);
    var max = 1;
    data.forEach(function (v) { if (v > max) max = v; });
    data.forEach(function (v) {
      var bar = document.createElement("span");
      var ratio = max > 0 ? (v / max) : 0;
      bar.style.height = Math.max(2, ratio * 100) + "%";
      var alpha = 0.45 + ratio * 0.5;
      bar.style.background = fmtHue(hue, 0.6, 0.14, Math.min(0.95, alpha));
      wrap.appendChild(bar);
    });
    return wrap;
  }

  // ---- Card primitives ----

  function applyThumbHue(thumb, hue) {
    thumb.style.setProperty(
      "--cg-card-thumb-bg",
      "radial-gradient(circle at 30% 40%, " + fmtHue(hue, 0.4, 0.12) + ", " + fmtHue(hue, 0.18, 0.05) + ")"
    );
  }

  function setThumbImage(thumb, src) {
    if (!src) return null;
    var img = document.createElement("img");
    img.src = src;
    img.alt = "";
    img.draggable = false;
    img.addEventListener("error", function () {
      // Drop the broken <img> so the gradient bg shows through cleanly.
      if (img.parentNode) img.parentNode.removeChild(img);
    });
    thumb.appendChild(img);
    return img;
  }

  function makePill(text, side) {
    var pill = document.createElement("span");
    pill.className = "clip-card-pill-" + side;
    pill.textContent = text;
    return pill;
  }

  function buildCardBase(opts, kind) {
    opts = opts || {};
    var card = document.createElement("div");
    card.className = kind === "transcript" ? "transcript-card" : "clip-card";
    if (kind === "clip" && opts.size) card.classList.add("size-" + opts.size);
    card.setAttribute("draggable", "true");

    var hue = resolveHue(opts.label, opts.hue);
    card.style.setProperty("--cg-card-hue", fmtHue(hue));

    var thumb = document.createElement("div");
    thumb.className = kind === "transcript" ? "transcript-card-thumb" : "clip-card-thumb";
    applyThumbHue(thumb, hue);
    card.appendChild(thumb);
    if (opts.thumbSrc) setThumbImage(thumb, opts.thumbSrc);

    var dot = document.createElement("span");
    dot.className = "clip-card-dot";
    thumb.appendChild(dot);

    if (opts.participant) thumb.appendChild(makePill(opts.participant, "tl"));
    if (opts.duration) thumb.appendChild(makePill(opts.duration, "br"));

    setDataset(card, opts.dataset);

    if (typeof opts.onDragStart === "function") {
      card.addEventListener("dragstart", opts.onDragStart);
    }
    if (typeof opts.onClick === "function") {
      card.addEventListener("click", opts.onClick);
    }
    if (typeof opts.onMouseEnter === "function") {
      card.addEventListener("mouseenter", opts.onMouseEnter);
    }
    if (typeof opts.onMouseLeave === "function") {
      card.addEventListener("mouseleave", opts.onMouseLeave);
    }

    return { card: card, thumb: thumb };
  }

  function createClipCard(opts) {
    var built = buildCardBase(opts, "clip");
    var caption = document.createElement("div");
    caption.className = "clip-card-caption";

    if (opts.label) {
      var labelEl = document.createElement("span");
      labelEl.className = "clip-card-label";
      labelEl.textContent = opts.label;
      caption.appendChild(labelEl);
    }
    if (opts.caption) {
      var textEl = document.createElement("span");
      textEl.className = "clip-card-text";
      textEl.textContent = opts.caption;
      textEl.title = opts.caption;
      caption.appendChild(textEl);
    }
    built.card.appendChild(caption);
    return built.card;
  }

  function createTranscriptCard(opts) {
    var built = buildCardBase(opts, "transcript");
    var caption = document.createElement("div");
    caption.className = "transcript-card-caption";

    if (opts.timeRange) {
      var range = document.createElement("span");
      range.className = "transcript-card-range cg-mono";
      range.textContent = opts.timeRange;
      caption.appendChild(range);
    }
    if (opts.text) {
      var text = document.createElement("span");
      text.className = "transcript-card-text";
      text.textContent = opts.text;
      caption.appendChild(text);
    }
    built.card.appendChild(caption);
    return built.card;
  }

  // ---- Btn ----

  function createBtn(opts) {
    opts = opts || {};
    var btn = document.createElement("button");
    btn.type = "button";
    var classes = ["cg-btn"];
    var variant = opts.variant || "ghost";
    if (variant !== "ghost") classes.push("cg-btn-" + variant);
    var size = opts.size || "md";
    if (size !== "md") classes.push("cg-btn-" + size);
    btn.className = classes.join(" ");

    if (opts.icon) {
      var ic = document.createElement("span");
      ic.className = "cg-btn-icon";
      ic.style.maskImage = 'url("icons/' + opts.icon + '.svg")';
      ic.style.webkitMaskImage = 'url("icons/' + opts.icon + '.svg")';
      btn.appendChild(ic);
    }
    if (opts.label) {
      var label = document.createElement("span");
      label.textContent = opts.label;
      btn.appendChild(label);
    }
    if (opts.disabled) btn.disabled = true;
    setDataset(btn, opts.dataset);
    if (typeof opts.onClick === "function") {
      btn.addEventListener("click", opts.onClick);
    }
    return btn;
  }

  global.ClipgenPrimitives = {
    createFilterChip: createFilterChip,
    createParticipantPill: createParticipantPill,
    createDensityTimeline: createDensityTimeline,
    createSparkBars: createSparkBars,
    createClipCard: createClipCard,
    createTranscriptCard: createTranscriptCard,
    createBtn: createBtn,
  };
})(window);
