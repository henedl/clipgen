/* clipgen animated dot-grid background – grid-bg.js
 *
 * Draws a dot grid to a fixed full-viewport canvas behind all page content.
 * Adds a subtle ambient shimmer and supports radial "pulse" waves that
 * ripple outward from a point. Exposes a global API so any button press or
 * completed process can trigger a pulse.
 *
 * Global API:
 *   window.clipgenGridBg.pulse(x, y, opts?) – start a ripple at viewport (x, y)
 *   window.clipgenGridBg.setShimmer(level)  – multiplier on ambient shimmer
 *   window.clipgenGridBg.destroy()          – tear down (dev/testing)
 */

(function () {
  var GRID_SIZE = 40;
  var DOT_RADIUS = 1.1;   // corner dot radius
  var BEAD_RADIUS = 0.55; // smaller dots between corners
  var BEADS = 3;          // beads between each pair of adjacent corners
  var MAX_PULSES = 24;    // higher cap so a drag wake can keep emitting

  // Default pulse parameters (viewport px / sec)
  var PULSE_DEFAULTS = {
    amplitude: 2.2,   // max dot displacement (px)
    speed: 640,       // wave radius growth (px/sec)
    width: 170,        // ring thickness (px std-dev of gaussian)
    decay: 0.85,       // exponential decay time constant (sec)
    brightness: 0.6,  // opacity multiplier at the wave crest (subtle)
  };

  // Drag wake: emit a small pulse every N px of pointer travel while dragging.
  var DRAG_EMIT_INTERVAL_PX = 36;
  var DRAG_PULSE_OPTS = {
    amplitude: 1.6,
    speed: 520,
    width: 50,
    decay: 0.7,
    brightness: 0.45,
  };

  var canvas = null;
  var ctx = null;
  var dpr = 1;
  var w = 0;
  var h = 0;
  var pulses = [];
  var shimmerLevel = 1;
  var reducedMotion = false;
  var gridColorRGB = "0,0,0";
  var gridColorAlpha = 0.03;
  var rafId = null;
  var startTime = 0;
  var resizeObs = null;
  var themeObs = null;
  var darkMq = null;
  var mousedownHandler = null;
  var mousemoveHandler = null;
  var mouseupHandler = null;
  var visibilityHandler = null;
  var reducedMotionMq = null;
  var dragging = false;
  var dragLastX = 0;
  var dragLastY = 0;

  function parseGridColor(raw) {
    if (!raw) return;
    var s = raw.trim();
    var m;
    if (s.indexOf("rgba") === 0) {
      m = s.match(/rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([\d.]+)\s*\)/);
      if (m) {
        gridColorRGB = m[1] + "," + m[2] + "," + m[3];
        gridColorAlpha = parseFloat(m[4]);
        return;
      }
    }
    if (s.indexOf("rgb") === 0) {
      m = s.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
      if (m) {
        gridColorRGB = m[1] + "," + m[2] + "," + m[3];
        gridColorAlpha = 1;
        return;
      }
    }
    if (s.charAt(0) === "#") {
      var hex = s.slice(1);
      if (hex.length === 3) hex = hex.replace(/(.)/g, "$1$1");
      if (hex.length === 6) {
        gridColorRGB =
          parseInt(hex.slice(0, 2), 16) + "," +
          parseInt(hex.slice(2, 4), 16) + "," +
          parseInt(hex.slice(4, 6), 16);
        gridColorAlpha = 1;
      }
    }
  }

  function readThemeColor() {
    var raw = getComputedStyle(document.documentElement)
      .getPropertyValue("--color-grid");
    parseGridColor(raw);
    // When the CSS var is very faint (e.g. 0.03) the shimmer/pulse boosts
    // won't be visible. Use a perceptual floor so dots are drawable.
    if (gridColorAlpha < 0.08) gridColorAlpha = 0.22;
  }

  function resize() {
    if (!canvas) return;
    dpr = window.devicePixelRatio || 1;
    w = window.innerWidth;
    h = window.innerHeight;
    canvas.width = Math.round(w * dpr);
    canvas.height = Math.round(h * dpr);
    canvas.style.width = w + "px";
    canvas.style.height = h + "px";
  }

  // Cheap deterministic hash → [0, 1).
  function hash2(ix, iy) {
    var h = ix * 374761393 + iy * 668265263;
    h = (h ^ (h >>> 13)) * 1274126177;
    h = h ^ (h >>> 16);
    return ((h >>> 0) % 100000) / 100000;
  }

  // 2D value noise in [-1, 1] with bilinear interpolation between integer grid
  // hashes. Domain is arbitrary (callers pick a scale).
  function valueNoise(x, y) {
    var xi = Math.floor(x);
    var yi = Math.floor(y);
    var xf = x - xi;
    var yf = y - yi;
    var u = xf * xf * (3 - 2 * xf);
    var v = yf * yf * (3 - 2 * yf);
    var a = hash2(xi, yi);
    var b = hash2(xi + 1, yi);
    var c = hash2(xi, yi + 1);
    var d = hash2(xi + 1, yi + 1);
    var ab = a + (b - a) * u;
    var cd = c + (d - c) * u;
    return (ab + (cd - ab) * v) * 2 - 1;
  }

  function drawSample(gx, gy, fx, fy, t, shimmerAmp, baseAlpha, isCorner) {
    // Per-point noise — slow-moving, low frequency. Drives organic variation
    // in shimmer phase and in pulse wave strength so ripples don't look like
    // perfect uniform rings.
    var n = valueNoise(fx * 0.18 + t * 0.08, fy * 0.18 - t * 0.05);

    // Ambient shimmer: phase varies by (fractional) grid position + time,
    // with noise adding a slow, per-region phase wander.
    var shimmer = 0;
    if (shimmerAmp > 0) {
      shimmer = Math.sin(t * 1.6 + fx * 0.55 + fy * 0.41 + n * 2.4) * shimmerAmp;
      shimmer += n * 0.25 * shimmerAmp; // small DC offset so some regions stay brighter
    }

    // Sum per-pulse contributions.
    var dx = 0;
    var dy = 0;
    var boost = 0;
    for (var k = 0; k < pulses.length; k++) {
      var pp = pulses[k];
      if (reducedMotion) break;
      var ageP = t - pp.t0;
      if (ageP < 0) continue;
      var rx = gx - pp.x;
      var ry = gy - pp.y;
      var r = Math.sqrt(rx * rx + ry * ry);
      var frontR = pp.speed * ageP;
      var ring = Math.exp(-((r - frontR) * (r - frontR)) / (2 * pp.width * pp.width));
      var env = Math.exp(-ageP / pp.decay);
      // Noise breaks up the ring so the wave feels organic, not mechanical.
      var wave = ring * env * (1 + 0.55 * n);
      if (wave < 0.005) continue;
      if (r > 0.01) {
        dx += (rx / r) * pp.amplitude * wave;
        dy += (ry / r) * pp.amplitude * wave;
      }
      boost += pp.brightness * wave;
    }

    // Beads sit slightly dimmer than corners so the lattice still reads.
    var sampleAlpha = isCorner ? baseAlpha : baseAlpha * 0.75;
    var alpha = Math.max(0, Math.min(1, sampleAlpha * (1 + shimmer + boost)));
    if (alpha <= 0.002) return;
    var baseRad = isCorner ? DOT_RADIUS : BEAD_RADIUS;
    var rad = baseRad + Math.min(0.5, boost * 0.35);
    ctx.fillStyle = "rgba(" + gridColorRGB + "," + alpha.toFixed(4) + ")";
    ctx.beginPath();
    ctx.arc(gx + dx, gy + dy, rad, 0, Math.PI * 2);
    ctx.fill();
  }

  function drawFrame(nowMs) {
    rafId = null;
    if (!ctx) return;
    var t = (nowMs - startTime) / 1000;

    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, w, h);

    // Prune dead pulses.
    var live = [];
    for (var i = 0; i < pulses.length; i++) {
      var p = pulses[i];
      var age = t - p.t0;
      if (age < 0) continue;
      // Stop once envelope has decayed to ~1% and wave is past the viewport.
      var envelope = Math.exp(-age / p.decay);
      var maxR = Math.sqrt(w * w + h * h);
      if (envelope > 0.01 || p.speed * age < maxR + p.width * 3) {
        live.push(p);
      }
    }
    pulses = live;

    var shimmerAmp = reducedMotion ? 0 : 0.35 * shimmerLevel;
    var cols = Math.ceil(w / GRID_SIZE) + 1;
    var rows = Math.ceil(h / GRID_SIZE) + 1;
    var baseAlpha = gridColorAlpha;
    var step = GRID_SIZE / (BEADS + 1); // pixel gap between adjacent sample points

    // Iterate over corners; emit the corner plus BEADS beads to its right and
    // BEADS beads below it. That covers the full beaded grid without overlap.
    for (var iy = 0; iy < rows; iy++) {
      var gy = iy * GRID_SIZE;
      for (var ix = 0; ix < cols; ix++) {
        var gx = ix * GRID_SIZE;

        // --- corner ---
        drawSample(gx, gy, ix, iy, t, shimmerAmp, baseAlpha, true);

        // --- horizontal beads between (gx, gy) and (gx + GRID_SIZE, gy) ---
        for (var b = 1; b <= BEADS; b++) {
          var bx = gx + b * step;
          // Fractional grid coord keeps shimmer phase continuous along the line.
          drawSample(bx, gy, ix + b / (BEADS + 1), iy, t, shimmerAmp, baseAlpha, false);
        }

        // --- vertical beads between (gx, gy) and (gx, gy + GRID_SIZE) ---
        for (var b2 = 1; b2 <= BEADS; b2++) {
          var by = gy + b2 * step;
          drawSample(gx, by, ix, iy + b2 / (BEADS + 1), t, shimmerAmp, baseAlpha, false);
        }
      }
    }

    // Keep looping while there is motion to draw.
    if (!document.hidden && (shimmerAmp > 0 || pulses.length > 0)) {
      rafId = requestAnimationFrame(drawFrame);
    }
  }

  function requestDraw() {
    if (rafId != null) return;
    if (document.hidden) return;
    rafId = requestAnimationFrame(drawFrame);
  }

  function pulse(x, y, opts) {
    if (reducedMotion) return;
    opts = opts || {};
    var now = performance.now();
    var t = (now - startTime) / 1000;
    pulses.push({
      x: x,
      y: y,
      t0: t,
      amplitude: opts.amplitude != null ? opts.amplitude : PULSE_DEFAULTS.amplitude,
      speed: opts.speed != null ? opts.speed : PULSE_DEFAULTS.speed,
      width: opts.width != null ? opts.width : PULSE_DEFAULTS.width,
      decay: opts.decay != null ? opts.decay : PULSE_DEFAULTS.decay,
      brightness: opts.brightness != null ? opts.brightness : PULSE_DEFAULTS.brightness,
    });
    if (pulses.length > MAX_PULSES) pulses.shift();
    requestDraw();
  }

  function setShimmer(level) {
    shimmerLevel = Math.max(0, level);
    requestDraw();
  }

  function onBackgroundMousedown(ev) {
    // Only trigger when the press landed on the body itself (i.e. background),
    // not on any interactive element.
    if (ev.target !== document.body) return;
    if (ev.button !== 0) return; // left button only
    dragging = true;
    dragLastX = ev.clientX;
    dragLastY = ev.clientY;
    pulse(ev.clientX, ev.clientY);
  }

  function onDragMousemove(ev) {
    if (!dragging) return;
    var dx = ev.clientX - dragLastX;
    var dy = ev.clientY - dragLastY;
    if (dx * dx + dy * dy < DRAG_EMIT_INTERVAL_PX * DRAG_EMIT_INTERVAL_PX) return;
    dragLastX = ev.clientX;
    dragLastY = ev.clientY;
    pulse(ev.clientX, ev.clientY, DRAG_PULSE_OPTS);
  }

  function onDragMouseup() {
    dragging = false;
  }

  function onVisibility() {
    if (!document.hidden) requestDraw();
  }

  function onThemeChange() {
    readThemeColor();
    requestDraw();
  }

  function init() {
    if (canvas) return;
    // Feature flag (declared in utils.js). If disabled, leave the static CSS
    // grid in place and don't mount the canvas.
    if (typeof CLIPGEN_ANIMATED_BG !== "undefined" && !CLIPGEN_ANIMATED_BG) return;
    document.body.classList.add("cg-grid-bg-on");
    reducedMotionMq = window.matchMedia("(prefers-reduced-motion: reduce)");
    reducedMotion = reducedMotionMq.matches;
    if (reducedMotionMq.addEventListener) {
      reducedMotionMq.addEventListener("change", function (e) {
        reducedMotion = e.matches;
        requestDraw();
      });
    }

    canvas = document.createElement("canvas");
    canvas.id = "cg-grid-bg";
    canvas.setAttribute("aria-hidden", "true");
    if (document.body.firstChild) {
      document.body.insertBefore(canvas, document.body.firstChild);
    } else {
      document.body.appendChild(canvas);
    }
    ctx = canvas.getContext("2d");

    readThemeColor();
    resize();
    startTime = performance.now();

    if (typeof ResizeObserver !== "undefined") {
      resizeObs = new ResizeObserver(function () {
        resize();
        requestDraw();
      });
      resizeObs.observe(document.documentElement);
    } else {
      window.addEventListener("resize", function () {
        resize();
        requestDraw();
      });
    }

    themeObs = new MutationObserver(onThemeChange);
    themeObs.observe(document.documentElement, {
      attributes: true,
      attributeFilter: ["data-theme", "class"],
    });
    darkMq = window.matchMedia("(prefers-color-scheme: dark)");
    if (darkMq.addEventListener) darkMq.addEventListener("change", onThemeChange);

    visibilityHandler = onVisibility;
    document.addEventListener("visibilitychange", visibilityHandler);

    mousedownHandler = onBackgroundMousedown;
    mousemoveHandler = onDragMousemove;
    mouseupHandler = onDragMouseup;
    document.addEventListener("mousedown", mousedownHandler);
    document.addEventListener("mousemove", mousemoveHandler);
    document.addEventListener("mouseup", mouseupHandler);

    requestDraw();
  }

  function destroy() {
    if (rafId != null) cancelAnimationFrame(rafId);
    rafId = null;
    if (resizeObs) resizeObs.disconnect();
    if (themeObs) themeObs.disconnect();
    if (darkMq && darkMq.removeEventListener) darkMq.removeEventListener("change", onThemeChange);
    if (visibilityHandler) document.removeEventListener("visibilitychange", visibilityHandler);
    if (mousedownHandler) document.removeEventListener("mousedown", mousedownHandler);
    if (mousemoveHandler) document.removeEventListener("mousemove", mousemoveHandler);
    if (mouseupHandler) document.removeEventListener("mouseup", mouseupHandler);
    if (canvas && canvas.parentNode) canvas.parentNode.removeChild(canvas);
    canvas = null;
    ctx = null;
    pulses = [];
    document.body.classList.remove("cg-grid-bg-on");
  }

  window.clipgenGridBg = {
    pulse: pulse,
    setShimmer: setShimmer,
    destroy: destroy,
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
