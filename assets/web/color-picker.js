/* clipgen reusable color picker — color-picker.js
 *
 * A small, dependency-light popover color picker for any clipgen web UI.
 * Vanilla JS, no frameworks. Keeps its own HSV math (h in [0,360), s/v in
 * [0,1]) but reuses hexToRgb/rgbToHex and positionPopoverAnchored from utils.js,
 * which every page loads before this file.
 *
 * Usage:
 *   window.ClipgenColorPicker.open({
 *     anchor: el,                  // element to position the popover near
 *     value: "#ff8800",            // initial color (#rrggbb)
 *     swatches: ["#000000", ...],  // optional preset row (defaults provided)
 *     onInput: function (hex) {},  // fires live while dragging / typing
 *     onChange: function (hex) {}, // fires on commit (pointer up, Enter, swatch, close)
 *     onClose: function (hex) {},  // fires when the popover closes
 *   });
 *   -> returns { close: fn }
 *
 * Only one picker is open at a time; opening a new one closes the previous.
 * Closes on outside pointerdown or Escape, and cleans up its document listeners.
 *
 * Positioning reuses the global positionPopoverAnchored() from utils.js.
 */

(function () {
  var DEFAULT_SWATCHES = [
    "#000000", "#ffffff", "#ef4444", "#f59e0b", "#eab308",
    "#22c55e", "#06b6d4", "#3b82f6", "#8b5cf6", "#ec4899",
  ];

  var _active = null; // { root, cleanup, opts, h, s, v }

  // ---- color math (h in [0,360), s/v in [0,1]) ----

  function _clamp01(x) { return x < 0 ? 0 : x > 1 ? 1 : x; }

  // hexToRgb / rgbToHex come from utils.js (loaded before this file).

  function rgbToHsv(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    var mx = Math.max(r, g, b), mn = Math.min(r, g, b), d = mx - mn;
    var h = 0;
    if (d) {
      if (mx === r) h = ((g - b) / d) % 6;
      else if (mx === g) h = (b - r) / d + 2;
      else h = (r - g) / d + 4;
      h *= 60;
      if (h < 0) h += 360;
    }
    return { h: h, s: mx === 0 ? 0 : d / mx, v: mx };
  }

  function hsvToRgb(h, s, v) {
    h = ((h % 360) + 360) % 360;
    var c = v * s, x = c * (1 - Math.abs(((h / 60) % 2) - 1)), m = v - c;
    var r, g, b;
    if (h < 60) { r = c; g = x; b = 0; }
    else if (h < 120) { r = x; g = c; b = 0; }
    else if (h < 180) { r = 0; g = c; b = x; }
    else if (h < 240) { r = 0; g = x; b = c; }
    else if (h < 300) { r = x; g = 0; b = c; }
    else { r = c; g = 0; b = x; }
    return { r: (r + m) * 255, g: (g + m) * 255, b: (b + m) * 255 };
  }

  function _currentHex(st) {
    var rgb = hsvToRgb(st.h, st.s, st.v);
    return rgbToHex(rgb.r, rgb.g, rgb.b);
  }

  // ---- drag helper: maps pointer position within rect to [0,1] x/y ----

  function _dragRegion(elem, onMove) {
    function fromEvent(ev) {
      var rect = elem.getBoundingClientRect();
      var x = _clamp01((ev.clientX - rect.left) / rect.width);
      var y = _clamp01((ev.clientY - rect.top) / rect.height);
      onMove(x, y);
    }
    elem.addEventListener("pointerdown", function (ev) {
      ev.preventDefault();
      fromEvent(ev);
      function move(e) { fromEvent(e); }
      function up() {
        document.removeEventListener("pointermove", move);
        document.removeEventListener("pointerup", up);
        if (_active && _active.opts.onChange) _active.opts.onChange(_currentHex(_active));
      }
      document.addEventListener("pointermove", move);
      document.addEventListener("pointerup", up);
    });
  }

  function close() {
    if (!_active) return;
    var st = _active;
    _active = null;
    st.cleanup();
    if (st.root.parentNode) st.root.parentNode.removeChild(st.root);
    var hex = _currentHex(st);
    // Commit on close so consumers relying only on onChange get the final value.
    if (st.opts.onChange) st.opts.onChange(hex);
    if (st.opts.onClose) st.opts.onClose(hex);
  }

  function open(opts) {
    opts = opts || {};
    close(); // single active instance

    var rgb = hexToRgb(opts.value) || { r: 0, g: 0, b: 0 };
    var hsv = rgbToHsv(rgb.r, rgb.g, rgb.b);

    var root = document.createElement("div");
    root.className = "cgcp-popover";

    var sv = document.createElement("div");
    sv.className = "cgcp-sv";
    var svThumb = document.createElement("div");
    svThumb.className = "cgcp-sv-thumb";
    sv.appendChild(svThumb);

    var hue = document.createElement("div");
    hue.className = "cgcp-hue";
    var hueThumb = document.createElement("div");
    hueThumb.className = "cgcp-hue-thumb";
    hue.appendChild(hueThumb);

    var area = document.createElement("div");
    area.className = "cgcp-area";
    area.appendChild(sv);
    area.appendChild(hue);
    root.appendChild(area);

    var row = document.createElement("div");
    row.className = "cgcp-row";
    var preview = document.createElement("span");
    preview.className = "cgcp-preview";
    var hexInput = document.createElement("input");
    hexInput.type = "text";
    hexInput.className = "cgcp-hex";
    hexInput.autocomplete = "off";
    hexInput.spellcheck = false;
    hexInput.maxLength = 7;
    row.appendChild(preview);
    row.appendChild(hexInput);
    root.appendChild(row);

    var swatchWrap = document.createElement("div");
    swatchWrap.className = "cgcp-swatches";
    var swatches = opts.swatches || DEFAULT_SWATCHES;
    root.appendChild(swatchWrap);

    var st = { root: root, opts: opts, h: hsv.h, s: hsv.s, v: hsv.v, cleanup: function () {} };

    function emitInput() {
      if (opts.onInput) opts.onInput(_currentHex(st));
    }

    function renderFromState(updateHexField) {
      sv.style.backgroundColor = rgbToHexHue(st.h);
      svThumb.style.left = st.s * 100 + "%";
      svThumb.style.top = (1 - st.v) * 100 + "%";
      hueThumb.style.top = (st.h / 360) * 100 + "%";
      var hex = _currentHex(st);
      preview.style.background = hex;
      svThumb.style.background = hex;
      if (updateHexField) hexInput.value = hex;
    }

    function rgbToHexHue(h) {
      var c = hsvToRgb(h, 1, 1);
      return rgbToHex(c.r, c.g, c.b);
    }

    _dragRegion(sv, function (x, y) {
      st.s = x;
      st.v = 1 - y;
      renderFromState(true);
      emitInput();
    });
    _dragRegion(hue, function (_x, y) {
      st.h = y * 360;
      renderFromState(true);
      emitInput();
    });

    hexInput.addEventListener("input", function () {
      var c = hexToRgb(hexInput.value);
      if (!c) return;
      var hsv2 = rgbToHsv(c.r, c.g, c.b);
      st.h = hsv2.h; st.s = hsv2.s; st.v = hsv2.v;
      renderFromState(false);
      emitInput();
    });
    hexInput.addEventListener("change", function () {
      if (opts.onChange) opts.onChange(_currentHex(st));
    });
    hexInput.addEventListener("keydown", function (ev) {
      if (ev.key === "Enter") close();
    });

    for (var i = 0; i < swatches.length; i++) {
      (function (hex) {
        var sw = document.createElement("button");
        sw.type = "button";
        sw.className = "cgcp-swatch";
        sw.style.background = hex;
        sw.title = hex;
        sw.addEventListener("click", function () {
          var c = hexToRgb(hex);
          if (!c) return;
          var hsv2 = rgbToHsv(c.r, c.g, c.b);
          st.h = hsv2.h; st.s = hsv2.s; st.v = hsv2.v;
          renderFromState(true);
          emitInput();
          if (opts.onChange) opts.onChange(_currentHex(st));
        });
        swatchWrap.appendChild(sw);
      })(swatches[i]);
    }

    document.body.appendChild(root);
    renderFromState(true);
    if (typeof positionPopoverAnchored === "function" && opts.anchor) {
      positionPopoverAnchored(root, opts.anchor.getBoundingClientRect());
    }

    // Outside-click + Escape close, with listener cleanup on close.
    function onDocDown(ev) {
      if (root.contains(ev.target)) return;
      if (opts.anchor && opts.anchor.contains(ev.target)) return;
      close();
    }
    function onKey(ev) {
      if (ev.key === "Escape") close();
    }
    // Defer attaching so the opening click doesn't immediately close it.
    var attachTimer = setTimeout(function () {
      document.addEventListener("pointerdown", onDocDown, true);
      document.addEventListener("keydown", onKey, true);
    }, 0);
    st.cleanup = function () {
      clearTimeout(attachTimer);
      document.removeEventListener("pointerdown", onDocDown, true);
      document.removeEventListener("keydown", onKey, true);
    };

    _active = st;
    return { close: close };
  }

  window.ClipgenColorPicker = { open: open, close: close };
})();
