/* clipgen Composer — annotate satellite.
 *
 * The visual annotation layer: a canvas positioned over the <video>'s content
 * box (letterbox-aware) that live-renders text labels, freehand strokes, and
 * rotatable rect/ellipse shapes whose visibility span contains the playhead.
 * Owns the tool state machine (select / text / draw / rect / ellipse), pointer
 * capture for drawing, text placement, shape drag-create and corner/rotation
 * handles, the positioned text <input>, the color swatches + picker, and
 * screen↔normalized coordinate mapping (geometry is normalized 0..1 to the
 * frame so the browser preview matches the server's PIL burn-in at any
 * resolution; shape rotation is degrees clockwise, applied in pixel space).
 *
 * CRUD + undo/redo live in the hub (CO.createAnnotation / deleteAnnotation /
 * commitAnnotationField / selectAnnotation); this file only draws and
 * captures gestures. CLIPGEN_CONFIG / qs / el / clamp are ambient utils.js
 * globals.
 */
(function () {
  "use strict";

  var CO = window.ClipgenComposer;
  var state = CO.state;

  var SWATCH_COLORS = [
    "#f05a3c", "#f0b429", "#3ecf8e", "#38bdf8", "#a78bfa", "#f8fafc",
  ];

  var _hitBoxes = [];  // screen-space bboxes from the last render (topmost last)
  var _drawing = null; // {points: [[nx, ny], ...]} while a stroke is captured
  var _dragging = null; // {ann, startNx, startNy, origGeometry, moved}
  var _pendingText = null; // {x, y} normalized, while the text input is open
  var _erasing = null; // ids already deleted this eraser gesture (dedupe)
  var _shaping = null; // {x0, y0, x1, y1} normalized, while drag-creating a shape
  var _shapeEdit = null; // {ann, mode:"resize"|"rotate", corner, orig} handle drag

  var SHAPE_HANDLE_R = 4;    // corner/rotation handle radius, canvas px
  var SHAPE_ROT_OFFSET = 18; // rotation handle stem length above the top edge

  function canvasEl() { return qs("#coAnnotateCanvas"); }

  // ---- Content-box math (object-fit: contain letterboxing) ----

  // Position + size the overlay canvas exactly over the video's displayed
  // content area, so canvas pixels map linearly onto normalized frame coords.
  function syncCanvasToVideo() {
    var video = qs("#coVideo");
    var frame = qs("#coVideoFrame");
    var canvas = canvasEl();
    if (!video || !video.videoWidth || !video.videoHeight) return;
    var frameRect = frame.getBoundingClientRect();
    var videoRect = video.getBoundingClientRect();
    var scale = Math.min(
      videoRect.width / video.videoWidth,
      videoRect.height / video.videoHeight
    );
    var w = Math.round(video.videoWidth * scale);
    var h = Math.round(video.videoHeight * scale);
    canvas.width = w;
    canvas.height = h;
    canvas.style.width = w + "px";
    canvas.style.height = h + "px";
    canvas.style.left = Math.round(
      videoRect.left - frameRect.left + (videoRect.width - w) / 2) + "px";
    canvas.style.top = Math.round(
      videoRect.top - frameRect.top + (videoRect.height - h) / 2) + "px";
    renderAnnotations();
  }

  function eventToNormalized(e) {
    var rect = canvasEl().getBoundingClientRect();
    if (!rect.width || !rect.height) return null;
    return {
      x: clamp((e.clientX - rect.left) / rect.width, 0, 1),
      y: clamp((e.clientY - rect.top) / rect.height, 0, 1),
    };
  }

  // ---- Rendering ----

  // Span edges round-trip through the server's round(…, 3): a span started
  // at the un-rounded playhead can come back up to 0.5 ms LATER than it, so a
  // strict comparison made a just-created annotation vanish on mouse-up about
  // half the time. The tolerance also absorbs the <video> settling a frame
  // shy of a requested seek when a span was snapped to a cut edge.
  var SPAN_EPS = 0.005;

  function spanContainsPlayhead(a) {
    return a.span.start - SPAN_EPS <= state.playhead &&
      state.playhead <= a.span.end + SPAN_EPS;
  }

  function visibleAnnotations() {
    return CO.participantAnnotations().filter(spanContainsPlayhead);
  }

  // Rotated-frame math for shape annotations, shared by rendering, handles,
  // and hit tests. All outputs in canvas px; rotation is degrees clockwise
  // (the canvas y-down ctx.rotate convention — the server mirrors it).
  function shapeFrame(ann, w, h) {
    var g = ann.geometry;
    var rad = ((g.rotation || 0) * Math.PI) / 180;
    var cos = Math.cos(rad);
    var sin = Math.sin(rad);
    var cx = g.x * w;
    var cy = g.y * h;
    var hw = (g.w * w) / 2;
    var hh = (g.h * h) / 2;
    function pt(dx, dy) {
      return [cx + dx * cos - dy * sin, cy + dx * sin + dy * cos];
    }
    return {
      cx: cx, cy: cy, hw: hw, hh: hh, rad: rad, cos: cos, sin: sin,
      corners: [pt(-hw, -hh), pt(hw, -hh), pt(hw, hh), pt(-hw, hh)],
      topMid: pt(0, -hh),
      rotHandle: pt(0, -hh - SHAPE_ROT_OFFSET),
    };
  }

  function drawAnnotation(ctx, ann, w, h, selected) {
    var style = ann.style || {};
    var color = style.color || CLIPGEN_CONFIG.composerAnnotationColor;
    if (ann.type === "shape") {
      var frame = shapeFrame(ann, w, h);
      var shapeStroke = Math.max(1,
        (style.strokeWidth || CLIPGEN_CONFIG.composerAnnotationStrokeWidth) * w);
      ctx.save();
      ctx.translate(frame.cx, frame.cy);
      ctx.rotate(frame.rad);
      ctx.strokeStyle = color;
      ctx.lineWidth = shapeStroke;
      if (ann.geometry.shape === "ellipse") {
        ctx.beginPath();
        ctx.ellipse(0, 0, frame.hw, frame.hh, 0, 0, Math.PI * 2);
        ctx.stroke();
      } else {
        ctx.strokeRect(-frame.hw, -frame.hh, frame.hw * 2, frame.hh * 2);
      }
      ctx.restore();
      var cxs = frame.corners.map(function (p) { return p[0]; });
      var cys = frame.corners.map(function (p) { return p[1]; });
      return {
        x1: Math.min.apply(null, cxs) - shapeStroke,
        y1: Math.min.apply(null, cys) - shapeStroke,
        x2: Math.max.apply(null, cxs) + shapeStroke,
        y2: Math.max.apply(null, cys) + shapeStroke,
      };
    }
    if (ann.type === "freehand") {
      var points = (ann.geometry.points || []).map(function (p) {
        return [p[0] * w, p[1] * h];
      });
      if (!points.length) return null;
      var stroke = Math.max(1,
        (style.strokeWidth || CLIPGEN_CONFIG.composerAnnotationStrokeWidth) * w);
      ctx.strokeStyle = color;
      ctx.fillStyle = color;
      ctx.lineWidth = stroke;
      ctx.lineJoin = "round";
      ctx.lineCap = "round";
      if (points.length === 1) {
        ctx.beginPath();
        ctx.arc(points[0][0], points[0][1], Math.max(stroke, 2), 0, Math.PI * 2);
        ctx.fill();
      } else {
        ctx.beginPath();
        ctx.moveTo(points[0][0], points[0][1]);
        for (var i = 1; i < points.length; i++) ctx.lineTo(points[i][0], points[i][1]);
        ctx.stroke();
      }
      var xs = points.map(function (p) { return p[0]; });
      var ys = points.map(function (p) { return p[1]; });
      return {
        x1: Math.min.apply(null, xs) - stroke,
        y1: Math.min.apply(null, ys) - stroke,
        x2: Math.max.apply(null, xs) + stroke,
        y2: Math.max.apply(null, ys) + stroke,
      };
    }
    // text — mirrors the server's PIL render: dark backing box + colored text.
    var text = ann.geometry.text || "";
    if (!text) return null;
    var size = Math.max(8,
      (style.fontSize || CLIPGEN_CONFIG.composerAnnotationFontSize) * h);
    ctx.font = size + "px " + getCSSVar("--font-sans", "sans-serif");
    ctx.textBaseline = "top";
    var x = ann.geometry.x * w;
    var y = ann.geometry.y * h;
    var metrics = ctx.measureText(text);
    var pad = Math.max(2, size * 0.25);
    var box = {
      x1: x - pad, y1: y - pad,
      x2: x + metrics.width + pad, y2: y + size * 1.2 + pad,
    };
    ctx.fillStyle = "rgba(0, 0, 0, 0.43)";
    ctx.fillRect(box.x1, box.y1, box.x2 - box.x1, box.y2 - box.y1);
    ctx.fillStyle = color;
    ctx.fillText(text, x, y);
    return box;
  }

  function renderAnnotations() {
    var canvas = canvasEl();
    if (!canvas || !canvas.width || !canvas.height) return;
    var ctx = canvas.getContext("2d");
    var w = canvas.width;
    var h = canvas.height;
    ctx.clearRect(0, 0, w, h);
    _hitBoxes = [];
    // Hidden layer: nothing drawn, nothing hit-testable (select/erase find
    // nothing), but an in-flight stroke preview still renders below so the
    // draw tool keeps working.
    if (!state.annHidden) visibleAnnotations().forEach(function (ann) {
      var selected = ann.id === state.selectedAnnotationId;
      var box = drawAnnotation(ctx, ann, w, h, selected);
      if (!box) return;
      if (selected) {
        ctx.strokeStyle = getCSSVar("--color-accent", "#1d4f72");
        ctx.lineWidth = 1.5;
        ctx.setLineDash([4, 3]);
        ctx.strokeRect(box.x1, box.y1, box.x2 - box.x1, box.y2 - box.y1);
        ctx.setLineDash([]);
        if (ann.type === "shape" && state.annTool === "select") {
          drawShapeHandles(ctx, ann, w, h);
        }
      }
      _hitBoxes.push({ box: box, ann: ann });
    });
    // In-flight stroke preview.
    if (_drawing && _drawing.points.length) {
      var stroke = Math.max(1, CLIPGEN_CONFIG.composerAnnotationStrokeWidth * w);
      ctx.strokeStyle = state.annColor;
      ctx.lineWidth = stroke;
      ctx.lineJoin = "round";
      ctx.lineCap = "round";
      ctx.beginPath();
      ctx.moveTo(_drawing.points[0][0] * w, _drawing.points[0][1] * h);
      for (var i = 1; i < _drawing.points.length; i++) {
        ctx.lineTo(_drawing.points[i][0] * w, _drawing.points[i][1] * h);
      }
      ctx.stroke();
    }
    // In-flight shape-draft preview.
    if (_shaping) {
      var sx = Math.min(_shaping.x0, _shaping.x1) * w;
      var sy = Math.min(_shaping.y0, _shaping.y1) * h;
      var sw = Math.abs(_shaping.x1 - _shaping.x0) * w;
      var sh = Math.abs(_shaping.y1 - _shaping.y0) * h;
      ctx.strokeStyle = state.annColor;
      ctx.lineWidth = Math.max(1, CLIPGEN_CONFIG.composerAnnotationStrokeWidth * w);
      if (state.annTool === "ellipse") {
        ctx.beginPath();
        ctx.ellipse(sx + sw / 2, sy + sh / 2, sw / 2, sh / 2, 0, 0, Math.PI * 2);
        ctx.stroke();
      } else {
        ctx.strokeRect(sx, sy, sw, sh);
      }
    }
  }

  // Corner squares + a stemmed rotation knob for the selected shape.
  function drawShapeHandles(ctx, ann, w, h) {
    var frame = shapeFrame(ann, w, h);
    var accent = getCSSVar("--color-accent", "#1d4f72");
    var bg = getCSSVar("--bg", "#0d0e10");
    var r = SHAPE_HANDLE_R;
    ctx.strokeStyle = accent;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(frame.topMid[0], frame.topMid[1]);
    ctx.lineTo(frame.rotHandle[0], frame.rotHandle[1]);
    ctx.stroke();
    frame.corners.forEach(function (p) {
      ctx.fillStyle = bg;
      ctx.fillRect(p[0] - r, p[1] - r, r * 2, r * 2);
      ctx.strokeRect(p[0] - r + 0.5, p[1] - r + 0.5, r * 2 - 1, r * 2 - 1);
    });
    ctx.fillStyle = bg;
    ctx.beginPath();
    ctx.arc(frame.rotHandle[0], frame.rotHandle[1], r, 0, Math.PI * 2);
    ctx.fill();
    ctx.stroke();
  }

  // Handle hit test for the SELECTED shape only (handles render only there).
  function hitTestShapeHandle(nx, ny) {
    var ann = state.selectedAnnotationId &&
      CO.findAnnotation(state.selectedAnnotationId);
    if (!ann || ann.type !== "shape" || state.annHidden) return null;
    // Out-of-span shapes render nothing — their handles must not grab either.
    if (!spanContainsPlayhead(ann)) return null;
    var canvas = canvasEl();
    var px = nx * canvas.width;
    var py = ny * canvas.height;
    var frame = shapeFrame(ann, canvas.width, canvas.height);
    var slop = SHAPE_HANDLE_R + 3;
    if (Math.abs(px - frame.rotHandle[0]) <= slop &&
        Math.abs(py - frame.rotHandle[1]) <= slop) {
      return { ann: ann, mode: "rotate" };
    }
    for (var i = 0; i < 4; i++) {
      if (Math.abs(px - frame.corners[i][0]) <= slop &&
          Math.abs(py - frame.corners[i][1]) <= slop) {
        return { ann: ann, mode: "resize", corner: i };
      }
    }
    return null;
  }

  // Apply a handle drag to the shape's geometry (mutates in place; the orig
  // snapshot backs the undoable commit on pointer-up).
  function updateShapeEdit(pos, e) {
    var canvas = canvasEl();
    var w = canvas.width;
    var h = canvas.height;
    var g = _shapeEdit.ann.geometry;
    var orig = _shapeEdit.orig;
    var px = pos.x * w;
    var py = pos.y * h;
    var cx = orig.x * w;
    var cy = orig.y * h;
    if (_shapeEdit.mode === "rotate") {
      // The knob sits at local -y, so the pointer angle needs a +90° bias.
      var deg = (Math.atan2(py - cy, px - cx) * 180) / Math.PI + 90;
      if (e.shiftKey) deg = Math.round(deg / 15) * 15;
      g.rotation = ((deg % 360) + 360) % 360;
      return;
    }
    // Corner resize in the shape's local (rotated) frame; the corner opposite
    // the grabbed one stays fixed.
    var rad = ((orig.rotation || 0) * Math.PI) / 180;
    var cos = Math.cos(rad);
    var sin = Math.sin(rad);
    var dx = px - cx;
    var dy = py - cy;
    var lx = dx * cos + dy * sin;  // R(-θ)·(pointer − center)
    var ly = -dx * sin + dy * cos;
    var signs = [[-1, -1], [1, -1], [1, 1], [-1, 1]][_shapeEdit.corner];
    var fx = -signs[0] * (orig.w * w) / 2;
    var fy = -signs[1] * (orig.h * h) / 2;
    var newW = Math.max(Math.abs(lx - fx), 8);
    var newH = Math.max(Math.abs(ly - fy), 8);
    var mx = (lx + fx) / 2;
    var my = (ly + fy) / 2;
    g.x = clamp((cx + mx * cos - my * sin) / w, 0, 1);
    g.y = clamp((cy + mx * sin + my * cos) / h, 0, 1);
    g.w = Math.min(newW / w, 1);
    g.h = Math.min(newH / h, 1);
  }

  // Delete whatever sits under the eraser, once per annotation per gesture
  // (the hit boxes rebuild async after each delete; the dedupe map keeps a
  // slow response from double-deleting → 404 toasts).
  function eraseAt(pos) {
    var ann = hitTestAnnotation(pos.x, pos.y);
    if (!ann || _erasing[ann.id]) return;
    _erasing[ann.id] = true;
    CO.deleteAnnotation(ann.id);
  }

  function hitTestAnnotation(nx, ny) {
    var canvas = canvasEl();
    var px = nx * canvas.width;
    var py = ny * canvas.height;
    for (var i = _hitBoxes.length - 1; i >= 0; i--) {
      var hb = _hitBoxes[i].box;
      if (px >= hb.x1 && px <= hb.x2 && py >= hb.y1 && py <= hb.y2) {
        return _hitBoxes[i].ann;
      }
    }
    return null;
  }

  // ---- Tools ----

  function defaultSpan() {
    // Creation-only: an annotation placed while the playhead sits inside a cut
    // adopts that cut's span (the selected cut wins over an earlier overlap),
    // so it travels with the clip. Span edits afterwards are free-form.
    var cuts = (CO.participantCuts ? CO.participantCuts() : []).filter(function (c) {
      return c.start <= state.playhead && state.playhead <= c.end;
    });
    if (cuts.length) {
      var cut = null;
      cuts.forEach(function (c) {
        if (c.id === state.selectedCutId) cut = c;
      });
      if (!cut) {
        cuts.sort(function (a, b) { return a.start - b.start; });
        cut = cuts[0];
      }
      return { start: cut.start, end: cut.end };
    }
    var span = CLIPGEN_CONFIG.composerAnnotationSpanSeconds;
    // Back the start off by a few frames: the <video> can settle a frame
    // before a requested seek and the server rounds span edges, so opening
    // the span exactly at the playhead risks it starting just out of view.
    var start = Math.max(0, state.playhead - 0.1);
    var end = Math.min(
      state.duration || start + span, start + span);
    return { start: start, end: Math.max(end, start + 0.5) };
  }

  function setAnnotateTool(tool) {
    if (!state.participant) return;
    state.annTool = tool;
    var canvas = canvasEl();
    canvas.classList.toggle("co-tool-text", tool === "text");
    canvas.classList.toggle("co-tool-draw", tool === "draw");
    canvas.classList.toggle("co-tool-erase", tool === "erase");
    canvas.classList.toggle("co-tool-shape", tool === "rect" || tool === "ellipse");
    qsa(".co-tool-btn[data-tool]").forEach(function (btn) {
      btn.classList.toggle("active", btn.getAttribute("data-tool") === tool);
    });
    if (tool !== "text") hideTextInput();
  }

  // ---- Text input flow ----

  function showTextInput(pos) {
    var input = qs("#coAnnotateText");
    var canvas = canvasEl();
    _pendingText = pos;
    input.style.left = (canvas.offsetLeft + pos.x * canvas.width) + "px";
    input.style.top = (canvas.offsetTop + pos.y * canvas.height) + "px";
    input.value = "";
    input.classList.remove("hidden");
    // Focus on the next frame — belt-and-braces with the caller's
    // preventDefault against the default mousedown focus steal.
    requestAnimationFrame(function () { input.focus(); });
  }

  function hideTextInput() {
    var input = qs("#coAnnotateText");
    input.classList.add("hidden");
    _pendingText = null;
  }

  function commitTextInput() {
    var input = qs("#coAnnotateText");
    var text = input.value.trim();
    var pos = _pendingText;
    hideTextInput();
    if (!text || !pos) return;
    CO.createAnnotation({
      participant: state.participant,
      type: "text",
      span: defaultSpan(),
      geometry: { x: pos.x, y: pos.y, text: text },
      style: { color: state.annColor },
    });
    setAnnotateTool("select");
  }

  // ---- Init ----

  function initAnnotate() {
    var canvas = canvasEl();
    var video = qs("#coVideo");

    // Color swatches + custom-color picker. A picked color that matches no
    // preset leaves every preset inactive; the picker button always shows the
    // current color.
    var swatchHost = qs("#coAnnotateColors");

    function applyAnnColor(color) {
      state.annColor = color;
      qsa(".co-color-swatch").forEach(function (s) {
        s.classList.toggle("active", s.getAttribute("data-color") === color);
      });
      var custom = qs(".co-color-custom");
      if (custom) custom.style.setProperty("--co-swatch-color", color);
      // Recolor the selected annotation in place.
      var ann = state.selectedAnnotationId &&
        CO.findAnnotation(state.selectedAnnotationId);
      if (ann && ann.style.color !== color) {
        var before = JSON.parse(JSON.stringify(ann.style));
        ann.style.color = color;
        CO.commitAnnotationField(ann, "style", before);
      }
    }

    SWATCH_COLORS.forEach(function (color, idx) {
      var swatch = el("button", "co-color-swatch" + (idx === 0 ? " active" : ""));
      swatch.type = "button";
      swatch.title = "Annotation color " + color;
      swatch.setAttribute("aria-label", swatch.title);
      swatch.setAttribute("data-color", color);
      swatch.style.setProperty("--co-swatch-color", color);
      swatch.addEventListener("click", function () { applyAnnColor(color); });
      swatchHost.appendChild(swatch);
    });

    var custom = el("button", "co-color-swatch co-color-custom");
    custom.type = "button";
    custom.title = "Custom color…";
    custom.setAttribute("aria-label", custom.title);
    custom.style.setProperty("--co-swatch-color", state.annColor || SWATCH_COLORS[0]);
    custom.appendChild(el("span", "co-btn-icon co-icon-eye-dropper"));
    custom.addEventListener("click", function () {
      window.ClipgenColorPicker.open({
        anchor: custom,
        value: state.annColor || SWATCH_COLORS[0],
        swatches: SWATCH_COLORS,
        onChange: applyAnnColor,
      });
    });
    swatchHost.appendChild(custom);

    // Tool buttons ([data-tool] excludes the independent #coToolHide toggle).
    qsa(".co-tool-btn[data-tool]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        setAnnotateTool(btn.getAttribute("data-tool"));
      });
    });

    // Keep the overlay glued to the video content box.
    video.addEventListener("loadedmetadata", syncCanvasToVideo);
    if (typeof ResizeObserver === "function") {
      var obs = new ResizeObserver(syncCanvasToVideo);
      obs.observe(qs("#coVideoFrame"));
      window.addEventListener("pagehide", function () { obs.disconnect(); });
    } else {
      window.addEventListener("resize", syncCanvasToVideo);
    }

    // Text input commit/cancel.
    var input = qs("#coAnnotateText");
    input.addEventListener("keydown", function (e) {
      e.stopPropagation();
      if (e.key === "Enter") commitTextInput();
      else if (e.key === "Escape") hideTextInput();
    });
    input.addEventListener("blur", commitTextInput);

    // Pointer gestures: draw a stroke, place text, erase, or select/move.
    canvas.addEventListener("pointerdown", function (e) {
      if (!state.participant) return;
      var pos = eventToNormalized(e);
      if (!pos) return;
      if (state.annTool === "rect" || state.annTool === "ellipse") {
        _shaping = { x0: pos.x, y0: pos.y, x1: pos.x, y1: pos.y };
        canvas.setPointerCapture(e.pointerId);
      } else if (state.annTool === "draw") {
        _drawing = { points: [[pos.x, pos.y]] };
        canvas.setPointerCapture(e.pointerId);
      } else if (state.annTool === "text") {
        // preventDefault: the browser's default mousedown action would move
        // focus off the just-focused input, blur-committing it empty before
        // the user can type.
        e.preventDefault();
        showTextInput(pos);
      } else if (state.annTool === "erase") {
        _erasing = {};
        canvas.setPointerCapture(e.pointerId);
        eraseAt(pos);
      } else {
        // Shape handles win over body hits — they extend past the outline.
        var handle = hitTestShapeHandle(pos.x, pos.y);
        if (handle) {
          _shapeEdit = {
            ann: handle.ann,
            mode: handle.mode,
            corner: handle.corner,
            orig: JSON.parse(JSON.stringify(handle.ann.geometry)),
          };
          canvas.setPointerCapture(e.pointerId);
          return;
        }
        var ann = hitTestAnnotation(pos.x, pos.y);
        CO.selectAnnotation(ann ? ann.id : null);
        if (ann) {
          _dragging = {
            ann: ann,
            startX: pos.x,
            startY: pos.y,
            origGeometry: JSON.parse(JSON.stringify(ann.geometry)),
            moved: false,
          };
          canvas.setPointerCapture(e.pointerId);
        }
      }
    });

    canvas.addEventListener("pointermove", function (e) {
      var pos = eventToNormalized(e);
      if (!pos) return;
      if (_shaping) {
        _shaping.x1 = pos.x;
        _shaping.y1 = pos.y;
        renderAnnotations();
      } else if (_shapeEdit) {
        updateShapeEdit(pos, e);
        renderAnnotations();
      } else if (_drawing) {
        var last = _drawing.points[_drawing.points.length - 1];
        // Thin out near-duplicate samples; keeps stored strokes compact.
        if (Math.abs(pos.x - last[0]) + Math.abs(pos.y - last[1]) > 0.003) {
          _drawing.points.push([pos.x, pos.y]);
          renderAnnotations();
        }
      } else if (_erasing) {
        eraseAt(pos);
      } else if (_dragging) {
        var dx = pos.x - _dragging.startX;
        var dy = pos.y - _dragging.startY;
        if (Math.abs(dx) + Math.abs(dy) > 0.004) _dragging.moved = true;
        if (!_dragging.moved) return;
        var geometry = _dragging.ann.geometry;
        var orig = _dragging.origGeometry;
        if (_dragging.ann.type === "text" || _dragging.ann.type === "shape") {
          geometry.x = clamp(orig.x + dx, 0, 1);
          geometry.y = clamp(orig.y + dy, 0, 1);
        } else {
          geometry.points = orig.points.map(function (p) {
            return [clamp(p[0] + dx, 0, 1), clamp(p[1] + dy, 0, 1)];
          });
        }
        renderAnnotations();
      }
    });

    function endGesture(e) {
      if (canvas.hasPointerCapture && canvas.hasPointerCapture(e.pointerId)) {
        canvas.releasePointerCapture(e.pointerId);
      }
      if (_erasing) {
        _erasing = null;
        return;
      }
      if (_shaping) {
        var draft = _shaping;
        _shaping = null;
        // Reject accidental clicks: the drag must span a visible size.
        if (Math.abs(draft.x1 - draft.x0) * canvas.width >= 6 &&
            Math.abs(draft.y1 - draft.y0) * canvas.height >= 6) {
          CO.createAnnotation({
            participant: state.participant,
            type: "shape",
            span: defaultSpan(),
            geometry: {
              shape: state.annTool === "ellipse" ? "ellipse" : "rect",
              x: (draft.x0 + draft.x1) / 2,
              y: (draft.y0 + draft.y1) / 2,
              w: Math.abs(draft.x1 - draft.x0),
              h: Math.abs(draft.y1 - draft.y0),
              rotation: 0,
            },
            style: { color: state.annColor },
          });
          setAnnotateTool("select");
        }
        renderAnnotations(); // clear the draft preview
        return;
      }
      if (_shapeEdit) {
        var edit = _shapeEdit;
        _shapeEdit = null;
        if (JSON.stringify(edit.ann.geometry) !== JSON.stringify(edit.orig)) {
          CO.commitAnnotationField(edit.ann, "geometry", edit.orig);
        }
        return;
      }
      if (_drawing) {
        var points = _drawing.points;
        _drawing = null;
        if (points.length) {
          CO.createAnnotation({
            participant: state.participant,
            type: "freehand",
            span: defaultSpan(),
            geometry: { points: points },
            style: { color: state.annColor },
          });
        }
      } else if (_dragging) {
        var d = _dragging;
        _dragging = null;
        if (d.moved) {
          CO.commitAnnotationField(d.ann, "geometry", d.origGeometry);
        }
      }
    }
    canvas.addEventListener("pointerup", endGesture);
    canvas.addEventListener("pointercancel", endGesture);
  }

  CO.initAnnotate = initAnnotate;
  CO.renderAnnotations = renderAnnotations;
  CO.setAnnotateTool = setAnnotateTool;
  CO.syncAnnotateCanvas = syncCanvasToVideo;
})();
