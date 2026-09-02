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
// OpenCV HSV ranges (h 0–180, s/v 0–255), matching screenspace_primitives.py.

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

// hexToRgb / rgbToHex are utils.js globals; the HSV pair stays here for the OpenCV ranges.

// ---- Form input builders ----

function rangeInput(id, min, max, value, step) {
  var inp = document.createElement("input");
  inp.type = "range";
  inp.id = id;
  inp.min = min;
  inp.max = max;
  // Step before value: a range input snaps value to its step, which defaults to 1.
  if (step) inp.step = step;
  inp.value = value;
  return inp;
}

function numberInput(id, min, max, value, step) {
  var inp = document.createElement("input");
  inp.type = "number";
  inp.id = id;
  inp.min = min;
  inp.max = max;
  if (step) inp.step = step;
  inp.value = value;
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

// Backing store = CSS box × devicePixelRatio. Returns device-pixel { w, h, dpr }.
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

// ---- Geometry ----

// Two corners in any order → top-left { x, y, w, h }.
function normalizeRect(x1, y1, x2, y2) {
  return {
    x: Math.min(x1, x2),
    y: Math.min(y1, y2),
    w: Math.abs(x2 - x1),
    h: Math.abs(y2 - y1),
  };
}

// ---- Formatting ----

// One-line step summary for step chips and restored-task rows, e.g. "H120° S200 V255 · presence".
function formatMultitoolStepParams(step) {
  if (!step) return "";
  var t = step.type;
  if (t === "color") {
    var tc = step.target_color || {};
    var swatch = "H" + (tc.h || 0) + "° S" + (tc.s || 0) + " V" + (tc.v || 0);
    return step.color_mode === "presence" ? swatch + " · presence" : swatch;
  }
  if (t === "change") return ">" + ((step.threshold || 0) * 100).toFixed(0) + "%";
  if (t === "similarity") return "≥" + ((step.threshold || 0) * 100).toFixed(0) + "%";
  if (t === "text") return "“" + (step.search_string || "") + "”";
  if (t === "numbers") {
    var opSym = { gt: ">", lt: "<", eq: "=", gte: "≥", lte: "≤" }[step.operator] || step.operator || "";
    return (opSym + " " + step.target_value).trim();
  }
  if (t === "template") return "≥" + ((step.threshold || 0) * 100).toFixed(0) + "%";
  if (t === "flow") return ">" + (step.magnitude_threshold || 0).toFixed(1);
  if (t === "scene") {
    var refs = step.scene_references || [];
    if (refs.length === 1) return refs[0].name || "1 ref";
    return refs.length + " refs";
  }
  if (t === "inactivity") return "≥" + (step.threshold || 0) + "s";
  return "";
}

// Min-area readout: percentage plus approximate pixel count when the region area is known.
function _formatMinAreaReadout(pct, area) {
  if (!(pct > 0)) return "Any presence (no minimum size)";
  var txt = pct + "%";
  if (area && area > 0) {
    var px = Math.max(1, Math.round((pct / 100) * area));
    txt += " · ~" + px.toLocaleString() + " px";
  }
  return txt;
}

// ---- Shaped-region (lasso / magic wand) geometry ----
// Points are [x, y] pairs; no DOM.

// Axis-aligned bounding box of a point list.
function polygonBounds(points) {
  var minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  for (var i = 0; i < points.length; i++) {
    var p = points[i];
    if (p[0] < minX) minX = p[0];
    if (p[0] > maxX) maxX = p[0];
    if (p[1] < minY) minY = p[1];
    if (p[1] > maxY) maxY = p[1];
  }
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
}

// Absolute shoelace area of a closed polygon.
function polygonArea(points) {
  var area = 0;
  for (var i = 0, j = points.length - 1; i < points.length; j = i++) {
    area += points[j][0] * points[i][1] - points[i][0] * points[j][1];
  }
  return Math.abs(area) / 2;
}

// Ray-cast point-in-polygon (implicitly closed). Mirrors the Python
// point_in_mask_points in screenspace_primitives.py.
function pointInPolygon(x, y, points) {
  var inside = false;
  for (var i = 0, j = points.length - 1; i < points.length; j = i++) {
    var yi = points[i][1], yj = points[j][1];
    if ((yi > y) !== (yj > y)) {
      var crossX = (points[j][0] - points[i][0]) * (y - yi) / (yj - yi) + points[i][0];
      if (x < crossX) inside = !inside;
    }
  }
  return inside;
}

// Douglas-Peucker (iterative): keeps endpoints, drops points within epsilon of the chord.
function simplifyPolygon(points, epsilon) {
  if (points.length < 3 || epsilon <= 0) return points.slice();
  var keep = new Array(points.length);
  keep[0] = keep[points.length - 1] = true;
  var stack = [[0, points.length - 1]];
  while (stack.length) {
    var range = stack.pop();
    var a = points[range[0]], b = points[range[1]];
    var dx = b[0] - a[0], dy = b[1] - a[1];
    var chordLen = Math.sqrt(dx * dx + dy * dy) || 1e-9;
    var maxDist = 0, maxIdx = -1;
    for (var i = range[0] + 1; i < range[1]; i++) {
      var p = points[i];
      var dist = Math.abs(dy * p[0] - dx * p[1] + b[0] * a[1] - b[1] * a[0]) / chordLen;
      if (dist > maxDist) { maxDist = dist; maxIdx = i; }
    }
    if (maxDist > epsilon && maxIdx > 0) {
      keep[maxIdx] = true;
      stack.push([range[0], maxIdx], [maxIdx, range[1]]);
    }
  }
  var out = [];
  for (var k = 0; k < points.length; k++) if (keep[k]) out.push(points[k]);
  return out;
}

// Scanline flood fill within a per-channel RGB tolerance; returns a w*h Uint8Array or null.
function floodFillMask(data, w, h, sx, sy, tolerance) {
  sx = Math.round(sx); sy = Math.round(sy);
  if (sx < 0 || sy < 0 || sx >= w || sy >= h) return null;
  var seedIdx = (sy * w + sx) * 4;
  var sr = data[seedIdx], sg = data[seedIdx + 1], sb = data[seedIdx + 2];
  var mask = new Uint8Array(w * h);
  var stack = [sx, sy];
  var matches = function (px, py) {
    var i = (py * w + px) * 4;
    return Math.abs(data[i] - sr) <= tolerance &&
           Math.abs(data[i + 1] - sg) <= tolerance &&
           Math.abs(data[i + 2] - sb) <= tolerance;
  };
  while (stack.length) {
    var y = stack.pop();
    var x = stack.pop();
    // Walk to the left edge of this run.
    while (x > 0 && !mask[y * w + x - 1] && matches(x - 1, y)) x--;
    var spanUp = false, spanDown = false;
    while (x < w && !mask[y * w + x] && matches(x, y)) {
      mask[y * w + x] = 1;
      if (y > 0) {
        var up = !mask[(y - 1) * w + x] && matches(x, y - 1);
        if (up && !spanUp) { stack.push(x, y - 1); spanUp = true; }
        else if (!up) spanUp = false;
      }
      if (y < h - 1) {
        var down = !mask[(y + 1) * w + x] && matches(x, y + 1);
        if (down && !spanDown) { stack.push(x, y + 1); spanDown = true; }
        else if (!down) spanDown = false;
      }
      x++;
    }
  }
  return mask;
}

// Over a list of contours; saved shaped regions store one list per part.
function contoursBounds(contours) {
  var all = [];
  for (var i = 0; i < contours.length; i++) all = all.concat(contours[i]);
  return polygonBounds(all);
}

function contoursArea(contours) {
  var area = 0;
  for (var i = 0; i < contours.length; i++) area += polygonArea(contours[i]);
  return area;
}

function contoursTotalPoints(contours) {
  var n = 0;
  for (var i = 0; i < contours.length; i++) n += contours[i].length;
  return n;
}

// Rasterize {rect} / {contours} shapes into a w*h Uint8Array; per-contour fill mirrors cv2.fillPoly.
function rasterizeShapesMask(shapes, w, h) {
  var cv = document.createElement("canvas");
  cv.width = w;
  cv.height = h;
  var ctx = cv.getContext("2d");
  ctx.fillStyle = "#fff";
  shapes.forEach(function (shape) {
    if (shape.contours) {
      shape.contours.forEach(function (c) {
        if (c.length < 3) return;
        ctx.beginPath();
        ctx.moveTo(c[0][0], c[0][1]);
        for (var i = 1; i < c.length; i++) ctx.lineTo(c[i][0], c[i][1]);
        ctx.closePath();
        ctx.fill();
      });
    } else if (shape.rect) {
      ctx.fillRect(shape.rect.x, shape.rect.y, shape.rect.w, shape.rect.h);
    }
  });
  var data = ctx.getImageData(0, 0, w, h).data;
  var mask = new Uint8Array(w * h);
  // Canvas fills antialias — threshold the alpha channel at half coverage.
  for (var p = 0, a = 3; p < mask.length; p++, a += 4) {
    if (data[a] >= 128) mask[p] = 1;
  }
  return mask;
}

// Combine masks in place on `base` (shift = add, alt = subtract, shift+alt = intersect).
function combineShapeMasks(base, other, op) {
  for (var i = 0; i < base.length; i++) {
    if (op === "add") {
      if (other[i]) base[i] = 1;
    } else if (op === "subtract") {
      if (other[i]) base[i] = 0;
    } else if (op === "intersect") {
      if (!other[i]) base[i] = 0;
    }
  }
  return base;
}

// Outer contour per connected component (4-connected label + Moore trace); drops specks under minArea.
function maskToContours(mask, w, h, minArea) {
  var visited = new Uint8Array(w * h);
  var contours = [];
  var comp = new Uint8Array(w * h);
  for (var start = 0; start < mask.length; start++) {
    if (!mask[start] || visited[start]) continue;
    // Flood this component into `comp` (cleared per component).
    comp.fill(0);
    var stack = [start];
    visited[start] = 1;
    while (stack.length) {
      var idx = stack.pop();
      comp[idx] = 1;
      var x = idx % w, y = (idx / w) | 0;
      if (x > 0 && mask[idx - 1] && !visited[idx - 1]) { visited[idx - 1] = 1; stack.push(idx - 1); }
      if (x < w - 1 && mask[idx + 1] && !visited[idx + 1]) { visited[idx + 1] = 1; stack.push(idx + 1); }
      if (y > 0 && mask[idx - w] && !visited[idx - w]) { visited[idx - w] = 1; stack.push(idx - w); }
      if (y < h - 1 && mask[idx + w] && !visited[idx + w]) { visited[idx + w] = 1; stack.push(idx + w); }
    }
    var contour = traceMaskContour(comp, w, h);
    if (contour.length >= 3 && polygonArea(contour) >= (minArea || 0)) {
      contours.push(contour);
    }
    if (contours.length >= 32) break; // matches the server's contour cap
  }
  return contours;
}

// Douglas-Peucker each contour, growing epsilon until under maxTotal (server cap 400).
function simplifyContours(contours, epsilon, maxTotal) {
  var out = contours.map(function (c) { return simplifyPolygon(c, epsilon); });
  var eps = epsilon;
  while (contoursTotalPoints(out) > maxTotal) {
    eps *= 1.5;
    out = out.map(function (c) { return simplifyPolygon(c, eps); });
  }
  return out.filter(function (c) { return c.length >= 3; });
}

// Single axis-aligned quad (within tol) → {x,y,w,h} so rectangular results save as rects.
function contoursToAxisRect(contours, tol) {
  if (contours.length !== 1) return null;
  var pts = contours[0];
  // Collapse collinear runs (cross product ≈ 0 with both neighbors).
  var kept = [];
  for (var i = 0; i < pts.length; i++) {
    var prev = pts[(i + pts.length - 1) % pts.length];
    var next = pts[(i + 1) % pts.length];
    var cross = (pts[i][0] - prev[0]) * (next[1] - prev[1]) - (pts[i][1] - prev[1]) * (next[0] - prev[0]);
    if (Math.abs(cross) > tol) kept.push(pts[i]);
  }
  if (kept.length !== 4) return null;
  for (var k = 0; k < 4; k++) {
    var a = kept[k], b = kept[(k + 1) % 4];
    if (Math.abs(a[0] - b[0]) > tol && Math.abs(a[1] - b[1]) > tol) return null;
  }
  var bounds = polygonBounds(kept);
  return { x: Math.round(bounds.x), y: Math.round(bounds.y), w: Math.round(bounds.w), h: Math.round(bounds.h) };
}

// Moore-neighbor trace of the outer contour; holes are ignored (the fill closes them).
function traceMaskContour(mask, w, h) {
  var startX = -1, startY = -1;
  for (var i = 0; i < mask.length; i++) {
    if (mask[i]) { startX = i % w; startY = (i / w) | 0; break; }
  }
  if (startX < 0) return [];
  // Moore neighborhood in clockwise order starting from west.
  var DIRS = [[-1, 0], [-1, -1], [0, -1], [1, -1], [1, 0], [1, 1], [0, 1], [-1, 1]];
  var filled = function (x, y) {
    return x >= 0 && y >= 0 && x < w && y < h && mask[y * w + x] === 1;
  };
  var contour = [[startX, startY]];
  var cx = startX, cy = startY;
  var backtrack = 0; // direction index pointing at the previous (empty) cell
  var maxSteps = w * h * 4;
  for (var step = 0; step < maxSteps; step++) {
    var found = false;
    for (var d = 0; d < 8; d++) {
      var dirIdx = (backtrack + d) % 8;
      var nx = cx + DIRS[dirIdx][0], ny = cy + DIRS[dirIdx][1];
      if (filled(nx, ny)) {
        // Next backtrack: the direction just before the one that hit.
        backtrack = (dirIdx + 6) % 8;
        cx = nx; cy = ny;
        found = true;
        break;
      }
    }
    if (!found) break; // isolated pixel
    if (cx === startX && cy === startY) break;
    contour.push([cx, cy]);
  }
  return contour;
}
