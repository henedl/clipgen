/* clipgen Screenspace color-picker satellite — screenspace-color.js
 *
 * The HSV color target picker (preview swatch, hue/sat palette, brightness
 * strip, "from region" sampler). Carved out of screenspace.js to shrink the
 * page script; loaded after it. Reads hub state + helpers via
 * window.ClipgenScreenspace and registers its render functions as SS.*; the
 * hub keeps thin same-named delegators so existing call sites (and the
 * sampleColorFromRegion click handler) are unchanged. Bodies moved verbatim.
 */
(function () {
  "use strict";

  var SS = window.ClipgenScreenspace;
  var state = SS.state;
  var regionToPixels = SS.regionToPixels,
    refreshCalibration = SS.refreshCalibration,
    updateCalibrationThresholdLine = SS.updateCalibrationThresholdLine;

  // The hidden H/S/V/hex <input>s the single-tool color panel builds. Owned
  // here so the render math reads them directly; the hub's renderColorParams
  // populates them via SS.setColorHiddenInputs and reads back via getter.
  var _colorHiddenInputs = null;

  function updateColorPreview() {
    var preview = qs("#colorPreview");
    var c = _colorHiddenInputs;
    if (!preview || !c) return;
    var rgb = hsvToRgb(numberOrDefault(c.h.value, 0), numberOrDefault(c.s.value, 0), numberOrDefault(c.v.value, 0));
    preview.style.background = rgbToHex(rgb.r, rgb.g, rgb.b);
  }

  function setTargetColor(h, s, v) {
    h = clamp(Math.round(h), 0, 180);
    s = clamp(Math.round(s), 0, 255);
    v = clamp(Math.round(v), 0, 255);
    var c = _colorHiddenInputs;
    if (c) {
      c.h.value = h; c.s.value = s; c.v.value = v;
      var rgb = hsvToRgb(h, s, v);
      c.hex.value = rgbToHex(rgb.r, rgb.g, rgb.b);
    }
    updateColorPreview();
    renderColorPalette();
    renderBrightnessStrip();
    // The H/S/V hidden inputs are set programmatically (no DOM input event), so
    // the #workflowParams delegated listener never fires — nudge calibration
    // directly. The color target drives the score, so palette / pipette / "From
    // Region" must re-evaluate. refreshCalibration self-guards on panel state.
    updateCalibrationThresholdLine();
    refreshCalibration({ debounce: true });
  }

  function renderColorPalette() {
    var canvas = qs("#colorPalette");
    if (!canvas || !canvas.getBoundingClientRect().width) return;
    var size = sizeCanvasToDisplay(canvas);
    var w = size.w, h = size.h, dpr = size.dpr;
    var ctx = canvas.getContext("2d");

    // Hue spectrum (horizontal)
    var hueGrad = ctx.createLinearGradient(0, 0, w, 0);
    var stops = ["#ff0000", "#ffff00", "#00ff00", "#00ffff", "#0000ff", "#ff00ff", "#ff0000"];
    for (var i = 0; i < stops.length; i++) hueGrad.addColorStop(i / (stops.length - 1), stops[i]);
    ctx.fillStyle = hueGrad;
    ctx.fillRect(0, 0, w, h);

    // White-to-transparent overlay (bottom = white = low saturation)
    var satGrad = ctx.createLinearGradient(0, 0, 0, h);
    satGrad.addColorStop(0, "rgba(255,255,255,0)");
    satGrad.addColorStop(1, "rgba(255,255,255,1)");
    ctx.fillStyle = satGrad;
    ctx.fillRect(0, 0, w, h);

    // Black overlay for brightness
    var c = _colorHiddenInputs;
    var curH = c ? numberOrDefault(c.h.value, 0) : 0;
    var curS = c ? numberOrDefault(c.s.value, 0) : 0;
    var curV = c ? numberOrDefault(c.v.value, 0) : 0;
    var darkness = 1 - curV / 255;
    if (darkness > 0) {
      ctx.fillStyle = "rgba(0,0,0," + darkness + ")";
      ctx.fillRect(0, 0, w, h);
    }

    // Current position
    var cx = (curH / 180) * w;
    var cy = (1 - curS / 255) * h;

    // Tolerance range visualization
    var tol = numberOrDefault((qs("#paramColorTol") || {}).value, 0);
    if (tol > 0) {
      var tolH = tol * 90 / 100;
      var tolS = tol * 128 / 100;
      var rx = (tolH / 180) * w;
      var ry = (tolS / 255) * h;
      ctx.fillStyle = "rgba(255,255,255,0.18)";
      ctx.strokeStyle = "rgba(255,255,255,0.4)";
      ctx.lineWidth = 1 * dpr;
      ctx.beginPath();
      ctx.rect(cx - rx, cy - ry, rx * 2, ry * 2);
      ctx.fill();
      ctx.stroke();
    }

    // Crosshair indicator
    var r = 5 * dpr;
    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, Math.PI * 2);
    ctx.strokeStyle = "#fff";
    ctx.lineWidth = 2 * dpr;
    ctx.stroke();
    ctx.beginPath();
    ctx.arc(cx, cy, r + 1 * dpr, 0, Math.PI * 2);
    ctx.strokeStyle = "#000";
    ctx.lineWidth = 1 * dpr;
    ctx.stroke();
  }

  function renderBrightnessStrip() {
    var canvas = qs("#colorBrightness");
    if (!canvas || !canvas.getBoundingClientRect().width) return;
    var size = sizeCanvasToDisplay(canvas);
    var w = size.w, h = size.h, dpr = size.dpr;
    var ctx = canvas.getContext("2d");
    var c = _colorHiddenInputs;
    var curH = c ? numberOrDefault(c.h.value, 0) : 0;
    var curS = c ? numberOrDefault(c.s.value, 0) : 0;
    var curV = c ? numberOrDefault(c.v.value, 0) : 0;

    // Gradient from black (left) to fully saturated color (right)
    var fullRgb = hsvToRgb(curH, curS, 255);
    var grad = ctx.createLinearGradient(0, 0, w, 0);
    grad.addColorStop(0, "#000000");
    grad.addColorStop(1, rgbToHex(fullRgb.r, fullRgb.g, fullRgb.b));
    ctx.fillStyle = grad;
    ctx.fillRect(0, 0, w, h);

    // Position indicator
    var ix = (curV / 255) * w;
    var r = 4 * dpr;
    ctx.beginPath();
    ctx.arc(clamp(ix, r, w - r), h / 2, r, 0, Math.PI * 2);
    ctx.fillStyle = "#fff";
    ctx.fill();
    ctx.strokeStyle = "#000";
    ctx.lineWidth = 1 * dpr;
    ctx.stroke();
  }

  // The sampler reads the highlighted chip, not the run-picker selection, so the
  // button says which region that is. Called on every renderRegionChips() (the
  // chokepoint for state.activeRegion changes) and once when the panel renders.
  function updateColorSampleBtnLabel() {
    var btn = qs("#colorSampleBtn");
    if (!btn) return;
    btn.textContent = state.activeRegion
      ? 'From region "' + state.activeRegion + '"'
      : "From region";
  }

  function sampleColorFromRegion() {
    if (!state.frameImage || !state.activeRegion) {
      showToast("Select a saved region first");
      return;
    }
    var r = regionToPixels(state.regions[state.activeRegion]);
    var ctx = qs("#frameCanvas").getContext("2d");
    var imgData = ctx.getImageData(r.x, r.y, r.w, r.h);
    var data = imgData.data;
    var totalR = 0, totalG = 0, totalB = 0;
    var count = data.length / 4;
    for (var i = 0; i < data.length; i += 4) {
      totalR += data[i];
      totalG += data[i + 1];
      totalB += data[i + 2];
    }
    var hsv = rgbToHsv(Math.round(totalR / count), Math.round(totalG / count), Math.round(totalB / count));
    setTargetColor(hsv.h, hsv.s, hsv.v);
    showToast("Sampled color from " + state.activeRegion);
  }

  // ---- Published to the hub (delegators / renderColorParams forward here) ----
  SS.updateColorPreview = updateColorPreview;
  SS.setTargetColor = setTargetColor;
  SS.renderColorPalette = renderColorPalette;
  SS.renderBrightnessStrip = renderBrightnessStrip;
  SS.sampleColorFromRegion = sampleColorFromRegion;
  SS.updateColorSampleBtnLabel = updateColorSampleBtnLabel;
  SS.setColorHiddenInputs = function (v) { _colorHiddenInputs = v; };
  SS.getColorHiddenInputs = function () { return _colorHiddenInputs; };
})();
