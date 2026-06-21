/* clipgen Screenspace leaf helpers — screenspace-utils.js
 *
 * Pure, state-free helpers extracted from screenspace.js so the page script
 * (and its satellite files) shrink and can share them. Loaded after utils.js
 * and before screenspace.js; like utils.js these are global declarations (no
 * IIFE) so the page scripts running inside their own IIFEs reach them via the
 * scope chain. Nothing here touches Screenspace `state` or the DOM beyond the
 * element it is handed/creates.
 */

// ---- Color conversion ----
// HSV here uses OpenCV-style ranges to match screenspace.py: h 0–180, s/v 0–255.

function rgbToHsv(r, g, b) {
  var rn = r / 255, gn = g / 255, bn = b / 255;
  var max = Math.max(rn, gn, bn), min = Math.min(rn, gn, bn);
  var d = max - min, hue = 0;
  if (d > 0) {
    if (max === rn) hue = ((gn - bn) / d) % 6;
    else if (max === gn) hue = (bn - rn) / d + 2;
    else hue = (rn - gn) / d + 4;
    hue = Math.round(hue * 30);
    if (hue < 0) hue += 180;
  }
  return { h: hue, s: max > 0 ? Math.round((d / max) * 255) : 0, v: Math.round(max * 255) };
}

function hsvToRgb(h, s, v) {
  var hDeg = h * 2, sn = s / 255, vn = v / 255;
  var c = vn * sn, x = c * (1 - Math.abs((hDeg / 60) % 2 - 1)), m = vn - c;
  var r1 = 0, g1 = 0, b1 = 0;
  if (hDeg < 60) { r1 = c; g1 = x; }
  else if (hDeg < 120) { r1 = x; g1 = c; }
  else if (hDeg < 180) { g1 = c; b1 = x; }
  else if (hDeg < 240) { g1 = x; b1 = c; }
  else if (hDeg < 300) { r1 = x; b1 = c; }
  else { r1 = c; b1 = x; }
  return { r: Math.round((r1 + m) * 255), g: Math.round((g1 + m) * 255), b: Math.round((b1 + m) * 255) };
}

function rgbToHex(r, g, b) {
  return "#" + ((1 << 24) | (r << 16) | (g << 8) | b).toString(16).slice(1);
}

function hexToRgb(hex) {
  hex = hex.replace(/^#/, "");
  if (hex.length === 3) hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return null;
  var n = parseInt(hex, 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}

// ---- Form input builders ----

function rangeInput(id, min, max, value, step) {
  var inp = document.createElement("input");
  inp.type = "range";
  inp.id = id;
  inp.min = min;
  inp.max = max;
  inp.value = value;
  if (step) inp.step = step;
  return inp;
}

function numberInput(id, min, max, value, step) {
  var inp = document.createElement("input");
  inp.type = "number";
  inp.id = id;
  inp.min = min;
  inp.max = max;
  inp.value = value;
  if (step) inp.step = step;
  return inp;
}

function textInput(id, placeholder) {
  var inp = document.createElement("input");
  inp.type = "text";
  inp.autocomplete = "off";
  inp.id = id;
  inp.placeholder = placeholder || "";
  return inp;
}

// ---- Canvas ----

// Size a canvas's backing store to its CSS box × devicePixelRatio. Returns the
// device-pixel { w, h, dpr } so callers can scale their drawing context.
function sizeCanvasToDisplay(canvas) {
  var rect = canvas.getBoundingClientRect();
  var dpr = window.devicePixelRatio || 1;
  var w = Math.round(rect.width * dpr);
  var h = Math.round(rect.height * dpr);
  if (canvas.width !== w || canvas.height !== h) {
    canvas.width = w;
    canvas.height = h;
  }
  return { w: w, h: h, dpr: dpr };
}
