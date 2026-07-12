/* Overview Map tab — 3D similarity space over the study's participants.
 *
 * The server (/overview/api/data, overview.py) ships a RAW participant x feature
 * matrix; everything after that happens here so the weight sliders re-layout
 * instantly with no round-trips:
 *
 *   z-score per column (clamp |z| <= 4, zero-variance floor)
 *   -> per-group weights (w / sqrt(group size) so big groups don't outvote)
 *   -> deterministic 3-component PCA (power iteration + deflation; fixed
 *      init vector + sign convention, so the same data + weights always
 *      produce the identical layout)
 *   -> outlier score = RMS of the available weighted z's — distance from
 *      the cohort centroid in FULL feature space, never the projection.
 *
 * Missing sources: a participant absent from a manifest has null cells for
 * that group. Nulls impute to 0 (= the cohort mean) for the layout so they
 * don't push the dot anywhere, and are excluded from the outlier score.
 *
 * Rendering is vendored Three.js r147 (global THREE — see vendor/README.md)
 * with hand-rolled orbit controls and HTML overlay labels, drawn on demand
 * (dirty flag), not a free-running rAF loop.
 */

(function () {
  "use strict";

  var Z_CLAMP = 4;          // matches the explain panel's bar scale
  var STD_FLOOR = 1e-9;     // zero-variance columns contribute nothing
  var PCA_ITERATIONS = 50;  // plenty at <=30 x ~70
  var WORLD_RADIUS = 10;    // projected cloud is scaled to fill this radius
  var HALO_COUNT = 3;       // top-k outliers get a halo ring
  var TOP_FEATURES = 8;     // explain panel rows
  var LERP_MS = 250;        // dot travel time after a re-layout

  var state = {
    data: null,        // /overview/api/data payload
    weights: {},       // group key -> 0..2
    groupSizes: {},    // group key -> column count
    stats: null,       // per-column {mean, std} over non-null cells
    zRaw: null,        // [n][d] clamped z, or null where the group is missing
    weighted: null,    // [n][d] weighted z, nulls imputed to 0 (PCA input)
    coords: null,      // [n][3] projected, world-scaled
    components: null,  // [3][d] PCA loadings
    variance: [],      // per-component share of total variance
    scores: [],        // [n] outlier scores
    order: [],         // participant indices, most unusual first
    selected: -1,
    hovered: -1,
  };

  var three = {
    renderer: null,
    scene: null,
    camera: null,
    raycaster: null,
    dots: [],       // one sphere mesh per participant
    halos: [],      // top-k translucent rings
    labels: [],     // one overlay div per participant
    colors: null,   // {base, hot} THREE.Color from tokens
  };

  // Spherical orbit: hand-rolled (rotate + dolly is all a 30-dot scatter
  // needs); vendoring OrbitControls would double the third-party surface.
  var orbit = {
    theta: Math.PI / 4,
    phi: Math.PI / 3,
    radius: 30,
    dragging: false,
    moved: 0,
    lastX: 0,
    lastY: 0,
  };

  var renderPending = false;
  var lerp = { active: false, from: null, to: null, start: 0 };

  var els = {};

  // ---- Math layer ---------------------------------------------------------

  function computeStats(matrix, d) {
    var stats = [];
    var j, i;
    for (j = 0; j < d; j++) {
      var sum = 0, count = 0;
      for (i = 0; i < matrix.length; i++) {
        if (matrix[i][j] != null) { sum += matrix[i][j]; count++; }
      }
      var mean = count ? sum / count : 0;
      var varSum = 0;
      for (i = 0; i < matrix.length; i++) {
        if (matrix[i][j] != null) {
          var dv = matrix[i][j] - mean;
          varSum += dv * dv;
        }
      }
      var std = count ? Math.sqrt(varSum / count) : 0;
      stats.push({ mean: mean, std: std });
    }
    return stats;
  }

  function computeZ(matrix, stats) {
    var z = [];
    var i, j;
    for (i = 0; i < matrix.length; i++) {
      var row = [];
      for (j = 0; j < stats.length; j++) {
        var v = matrix[i][j];
        if (v == null || stats[j].std < STD_FLOOR) {
          row.push(v == null ? null : 0);
        } else {
          var zv = (v - stats[j].mean) / stats[j].std;
          row.push(Math.max(-Z_CLAMP, Math.min(Z_CLAMP, zv)));
        }
      }
      z.push(row);
    }
    return z;
  }

  function columnWeight(col) {
    var w = state.weights[col.group];
    if (w == null) w = 1;
    var size = state.groupSizes[col.group] || 1;
    return w / Math.sqrt(size);
  }

  function applyWeights(zRaw, columns) {
    var weighted = [];
    var i, j;
    for (i = 0; i < zRaw.length; i++) {
      var row = [];
      for (j = 0; j < columns.length; j++) {
        var z = zRaw[i][j];
        row.push(z == null ? 0 : z * columnWeight(columns[j]));
      }
      weighted.push(row);
    }
    return weighted;
  }

  // One principal component by power iteration on X^T X (never materialized).
  // Deterministic: fixed uniform init + fixed iteration count + the sign
  // convention that the largest-|loading| entry is positive.
  function pcaComponent(X, d) {
    var v = [];
    var i, j, k, it;
    for (j = 0; j < d; j++) v.push(1 / Math.sqrt(d));
    for (it = 0; it < PCA_ITERATIONS; it++) {
      var w = [];
      for (j = 0; j < d; j++) w.push(0);
      for (i = 0; i < X.length; i++) {
        var s = 0;
        for (k = 0; k < d; k++) s += X[i][k] * v[k];
        for (k = 0; k < d; k++) w[k] += X[i][k] * s;
      }
      var norm = 0;
      for (j = 0; j < d; j++) norm += w[j] * w[j];
      norm = Math.sqrt(norm);
      if (norm < 1e-12) break;
      for (j = 0; j < d; j++) v[j] = w[j] / norm;
    }
    var mx = 0;
    for (j = 1; j < d; j++) if (Math.abs(v[j]) > Math.abs(v[mx])) mx = j;
    if (v[mx] < 0) for (j = 0; j < d; j++) v[j] = -v[j];
    return v;
  }

  // Project X (weighted z matrix, column means are 0 by construction) onto
  // its top 3 components. Returns {coords, components, variance}.
  function pcaProject(X, d) {
    var n = X.length;
    var work = [];
    var i, j, c;
    var totalVar = 0;
    for (i = 0; i < n; i++) {
      work.push(X[i].slice());
      for (j = 0; j < d; j++) totalVar += X[i][j] * X[i][j];
    }
    var components = [];
    var coords = [];
    var variance = [];
    for (i = 0; i < n; i++) coords.push([0, 0, 0]);
    for (c = 0; c < 3; c++) {
      var v = pcaComponent(work, d);
      components.push(v);
      var compVar = 0;
      for (i = 0; i < n; i++) {
        var s = 0;
        for (j = 0; j < d; j++) s += work[i][j] * v[j];
        coords[i][c] = s;
        compVar += s * s;
        for (j = 0; j < d; j++) work[i][j] -= s * v[j]; // deflate
      }
      variance.push(totalVar > 0 ? compVar / totalVar : 0);
    }
    return { coords: coords, components: components, variance: variance };
  }

  // Outlier score: RMS of the participant's available weighted z's. Mean
  // (not sum) over available features keeps missing-source participants
  // comparable to fully-covered ones.
  function computeOutlierScores(zRaw, columns) {
    var scores = [];
    var i, j;
    for (i = 0; i < zRaw.length; i++) {
      var sq = 0, count = 0;
      for (j = 0; j < columns.length; j++) {
        if (zRaw[i][j] == null) continue;
        var wz = zRaw[i][j] * columnWeight(columns[j]);
        sq += wz * wz;
        count++;
      }
      scores.push(count ? Math.sqrt(sq / count) : 0);
    }
    return scores;
  }

  function scaleToWorld(coords) {
    var maxR = 0;
    var i, c;
    for (i = 0; i < coords.length; i++) {
      for (c = 0; c < 3; c++) maxR = Math.max(maxR, Math.abs(coords[i][c]));
    }
    if (maxR <= 0) return coords;
    var s = WORLD_RADIUS / maxR;
    var out = [];
    for (i = 0; i < coords.length; i++) {
      out.push([coords[i][0] * s, coords[i][1] * s, coords[i][2] * s]);
    }
    return out;
  }

  function recompute() {
    var data = state.data;
    var columns = data.columns;
    state.zRaw = computeZ(data.matrix, state.stats);
    state.weighted = applyWeights(state.zRaw, columns);
    var pca = pcaProject(state.weighted, columns.length);
    state.components = pca.components;
    state.variance = pca.variance;
    var newCoords = scaleToWorld(pca.coords);
    state.scores = computeOutlierScores(state.zRaw, columns);
    state.order = data.participants
      .map(function (_, i) { return i; })
      .sort(function (a, b) { return state.scores[b] - state.scores[a]; });

    if (state.coords && three.renderer) {
      startLerp(state.coords, newCoords);
    }
    state.coords = newCoords;

    renderOutliers();
    renderAxisLegend();
    if (three.renderer) {
      styleDots();
      if (!lerp.active) positionDots(state.coords);
      requestRender();
    }
    if (state.selected >= 0) renderExplain(state.selected);
  }

  // ---- Data load ----------------------------------------------------------

  function loadData() {
    apiGet("api/data")
      .then(function (data) {
        if (!data || !data.ok) throw new Error((data && data.error) || "bad payload");
        init(data);
      })
      .catch(function (err) {
        showNotice("Could not load study data: " + err.message);
      });
  }

  function init(data) {
    state.data = data;
    if (window.clipgenApplyConfig) window.clipgenApplyConfig(data.config);

    var i;
    state.groupSizes = {};
    for (i = 0; i < data.columns.length; i++) {
      var g = data.columns[i].group;
      state.groupSizes[g] = (state.groupSizes[g] || 0) + 1;
    }
    for (i = 0; i < data.groups.length; i++) state.weights[data.groups[i].key] = 1;

    if (!data.participants.length) {
      els.empty.classList.remove("hidden");
      return;
    }
    if (!window.THREE) {
      showNotice("Three.js failed to load — the map cannot render.");
      return;
    }
    if (data.participants.length < 3) {
      showNotice("Only " + data.participants.length + " participant" +
        (data.participants.length === 1 ? "" : "s") +
        " so far — 3+ make the similarity layout meaningful.");
    }

    state.stats = computeStats(data.matrix, data.columns.length);
    initScene();
    buildDots();
    renderWeights();
    recompute();
  }

  // ---- Three.js scene -----------------------------------------------------

  function tokenColor(name, fallback) {
    var raw = getComputedStyle(document.documentElement)
      .getPropertyValue(name).trim();
    var color = new THREE.Color(fallback);
    try { if (raw) color.set(raw); } catch (e) { /* keep fallback */ }
    return color;
  }

  function initScene() {
    var wrap = els.canvasWrap;
    three.colors = {
      base: tokenColor("--accent", "#2D6BFF"),
      hot: tokenColor("--severity-high", "#f87171"),
      axis: tokenColor("--border-strong", "#2e3138"),
      bg: tokenColor("--bg", "#0a0a0b"),
    };

    three.scene = new THREE.Scene();
    three.scene.background = three.colors.bg;
    three.camera = new THREE.PerspectiveCamera(
      50, wrap.clientWidth / wrap.clientHeight, 0.1, 500);
    three.renderer = new THREE.WebGLRenderer({ antialias: true });
    three.renderer.setPixelRatio(window.devicePixelRatio || 1);
    three.renderer.setSize(wrap.clientWidth, wrap.clientHeight);
    els.canvas.appendChild(three.renderer.domElement);
    three.raycaster = new THREE.Raycaster();

    // Faint axis tripod so rotation reads spatially; the axis legend in the
    // sidebar explains what each direction means.
    var axisMat = new THREE.LineBasicMaterial({
      color: three.colors.axis, transparent: true, opacity: 0.6,
    });
    var r = WORLD_RADIUS * 1.2;
    var pts = [
      -r, 0, 0, r, 0, 0,
      0, -r, 0, 0, r, 0,
      0, 0, -r, 0, 0, r,
    ];
    var axisGeo = new THREE.BufferGeometry();
    axisGeo.setAttribute("position", new THREE.Float32BufferAttribute(pts, 3));
    three.scene.add(new THREE.LineSegments(axisGeo, axisMat));

    bindOrbit();
    window.addEventListener("resize", onResize);
    applyCamera();
  }

  function buildDots() {
    var participants = state.data.participants;
    var geo = new THREE.SphereGeometry(0.45, 24, 16);
    var haloGeo = new THREE.SphereGeometry(0.8, 24, 16);
    var labelFrag = document.createDocumentFragment();
    var i;
    for (i = 0; i < participants.length; i++) {
      var mat = new THREE.MeshBasicMaterial({ color: three.colors.base });
      var dot = new THREE.Mesh(geo, mat);
      dot.userData.index = i;
      three.scene.add(dot);
      three.dots.push(dot);

      var label = document.createElement("div");
      label.className = "map-label";
      label.textContent = participants[i];
      if (isPartial(i)) label.classList.add("is-partial");
      labelFrag.appendChild(label);
      three.labels.push(label);
    }
    els.labels.appendChild(labelFrag);

    var h;
    for (h = 0; h < HALO_COUNT; h++) {
      var haloMat = new THREE.MeshBasicMaterial({
        color: three.colors.hot, transparent: true, opacity: 0.18,
      });
      var halo = new THREE.Mesh(haloGeo, haloMat);
      halo.visible = false;
      three.scene.add(halo);
      three.halos.push(halo);
    }
  }

  function isPartial(i) {
    var pid = state.data.participants[i];
    var avail = state.data.availability[pid] || {};
    var g;
    for (g in avail) if (avail.hasOwnProperty(g) && !avail[g]) return true;
    return false;
  }

  function styleDots() {
    var maxScore = 0;
    var i;
    for (i = 0; i < state.scores.length; i++) {
      maxScore = Math.max(maxScore, state.scores[i]);
    }
    for (i = 0; i < three.dots.length; i++) {
      var t = maxScore > 0 ? state.scores[i] / maxScore : 0;
      three.dots[i].material.color
        .copy(three.colors.base).lerp(three.colors.hot, t);
      var scale = (i === state.selected ? 1.4 : 1) * (0.85 + t * 0.5);
      three.dots[i].scale.setScalar(scale);
      three.labels[i].classList.toggle("is-selected", i === state.selected);
    }
  }

  function positionDots(coords) {
    var i;
    for (i = 0; i < three.dots.length; i++) {
      three.dots[i].position.set(coords[i][0], coords[i][1], coords[i][2]);
    }
    var h;
    for (h = 0; h < three.halos.length; h++) {
      var idx = state.order[h];
      var show = idx != null && state.scores[idx] > 0 &&
        state.data.participants.length >= 3;
      three.halos[h].visible = !!show;
      if (show) {
        three.halos[h].position.copy(three.dots[idx].position);
        three.halos[h].scale.copy(three.dots[idx].scale);
      }
    }
  }

  // Short position lerp after a re-layout so dots travel instead of teleport
  // — continuity makes "what moved when I changed the lens" followable.
  function startLerp(from, to) {
    lerp.active = true;
    lerp.from = from;
    lerp.to = to;
    lerp.start = 0;
    requestAnimationFrame(stepLerp);
  }

  function stepLerp(now) {
    if (!lerp.start) lerp.start = now;
    var t = Math.min((now - lerp.start) / LERP_MS, 1);
    var ease = t * (2 - t); // easeOutQuad
    var mixed = [];
    var i;
    for (i = 0; i < lerp.to.length; i++) {
      var f = lerp.from[i] || lerp.to[i];
      mixed.push([
        f[0] + (lerp.to[i][0] - f[0]) * ease,
        f[1] + (lerp.to[i][1] - f[1]) * ease,
        f[2] + (lerp.to[i][2] - f[2]) * ease,
      ]);
    }
    positionDots(mixed);
    requestRender();
    if (t < 1) {
      requestAnimationFrame(stepLerp);
    } else {
      lerp.active = false;
      positionDots(lerp.to);
      requestRender();
    }
  }

  // ---- Orbit + picking ----------------------------------------------------

  function applyCamera() {
    var sinPhi = Math.sin(orbit.phi);
    three.camera.position.set(
      orbit.radius * sinPhi * Math.cos(orbit.theta),
      orbit.radius * Math.cos(orbit.phi),
      orbit.radius * sinPhi * Math.sin(orbit.theta));
    three.camera.lookAt(0, 0, 0);
  }

  function bindOrbit() {
    var dom = three.renderer.domElement;

    dom.addEventListener("mousedown", function (e) {
      if (e.button !== 0) return;
      orbit.dragging = true;
      orbit.moved = 0;
      orbit.lastX = e.clientX;
      orbit.lastY = e.clientY;
    });

    window.addEventListener("mousemove", function (e) {
      if (orbit.dragging) {
        var dx = e.clientX - orbit.lastX;
        var dy = e.clientY - orbit.lastY;
        orbit.lastX = e.clientX;
        orbit.lastY = e.clientY;
        orbit.moved += Math.abs(dx) + Math.abs(dy);
        orbit.theta += dx * 0.008;
        orbit.phi = Math.max(0.15, Math.min(Math.PI - 0.15, orbit.phi - dy * 0.008));
        applyCamera();
        requestRender();
      } else {
        updateHover(e);
      }
    });

    window.addEventListener("mouseup", function (e) {
      if (!orbit.dragging) return;
      orbit.dragging = false;
      if (orbit.moved < 4) handleClick(e);
    });

    dom.addEventListener("wheel", function (e) {
      e.preventDefault();
      orbit.radius = Math.max(8, Math.min(90, orbit.radius * (1 + e.deltaY * 0.0015)));
      applyCamera();
      requestRender();
    }, { passive: false });
  }

  function pick(e) {
    var rect = three.renderer.domElement.getBoundingClientRect();
    if (e.clientX < rect.left || e.clientX > rect.right ||
        e.clientY < rect.top || e.clientY > rect.bottom) return -1;
    var ndc = new THREE.Vector2(
      ((e.clientX - rect.left) / rect.width) * 2 - 1,
      -((e.clientY - rect.top) / rect.height) * 2 + 1);
    three.raycaster.setFromCamera(ndc, three.camera);
    var hits = three.raycaster.intersectObjects(three.dots, false);
    return hits.length ? hits[0].object.userData.index : -1;
  }

  function updateHover(e) {
    if (!three.renderer) return;
    var idx = pick(e);
    if (idx === state.hovered) {
      if (idx >= 0) moveTooltip(e);
      return;
    }
    state.hovered = idx;
    three.renderer.domElement.style.cursor = idx >= 0 ? "pointer" : "";
    if (idx < 0) {
      els.tooltip.classList.add("hidden");
      return;
    }
    var pid = state.data.participants[idx];
    var rank = state.order.indexOf(idx) + 1;
    els.tooltip.textContent = pid + " — unusualness " +
      state.scores[idx].toFixed(2) + " (#" + rank + ")";
    els.tooltip.classList.remove("hidden");
    moveTooltip(e);
  }

  function moveTooltip(e) {
    var rect = els.canvasWrap.getBoundingClientRect();
    els.tooltip.style.left = (e.clientX - rect.left + 14) + "px";
    els.tooltip.style.top = (e.clientY - rect.top + 14) + "px";
  }

  function handleClick(e) {
    var idx = pick(e);
    selectParticipant(idx === state.selected ? -1 : idx);
  }

  function selectParticipant(idx) {
    state.selected = idx;
    if (idx < 0) {
      els.explain.classList.add("hidden");
    } else {
      renderExplain(idx);
      els.explain.classList.remove("hidden");
    }
    renderOutliers();
    styleDots();
    requestRender();
    onResize(); // panel show/hide changes the canvas width
  }

  // ---- Render loop (on demand) --------------------------------------------

  function requestRender() {
    if (renderPending || !three.renderer) return;
    renderPending = true;
    requestAnimationFrame(function () {
      renderPending = false;
      three.renderer.render(three.scene, three.camera);
      updateLabels();
    });
  }

  function updateLabels() {
    var w = els.canvasWrap.clientWidth;
    var h = els.canvasWrap.clientHeight;
    var v = new THREE.Vector3();
    var i;
    for (i = 0; i < three.dots.length; i++) {
      v.copy(three.dots[i].position).project(three.camera);
      if (v.z > 1) {
        three.labels[i].style.display = "none";
        continue;
      }
      three.labels[i].style.display = "";
      three.labels[i].style.left = ((v.x + 1) / 2 * w) + "px";
      three.labels[i].style.top = ((1 - (v.y + 1) / 2) * h) + "px";
    }
  }

  function onResize() {
    if (!three.renderer) return;
    var wrap = els.canvasWrap;
    three.camera.aspect = wrap.clientWidth / wrap.clientHeight;
    three.camera.updateProjectionMatrix();
    three.renderer.setSize(wrap.clientWidth, wrap.clientHeight);
    requestRender();
  }

  // ---- Panels ---------------------------------------------------------------

  function showNotice(text) {
    els.notice.textContent = text;
    els.notice.classList.remove("hidden");
  }

  function fmtNum(v) {
    if (v == null) return "—";
    var a = Math.abs(v);
    if (a >= 100) return String(Math.round(v));
    if (a >= 10) return v.toFixed(1);
    return v.toFixed(2);
  }

  function renderWeights() {
    var frag = document.createDocumentFragment();
    state.data.groups.forEach(function (group) {
      var row = document.createElement("div");
      row.className = "map-weight-row";

      var label = document.createElement("div");
      label.className = "map-weight-label";
      var name = document.createElement("span");
      name.textContent = group.label;
      var value = document.createElement("span");
      value.className = "map-weight-value";
      value.textContent = "1.0";
      label.appendChild(name);
      label.appendChild(value);

      var slider = document.createElement("input");
      slider.type = "range";
      slider.min = "0";
      slider.max = "2";
      slider.step = "0.1";
      slider.value = "1";
      slider.setAttribute("aria-label", group.label + " weight");
      slider.addEventListener("input", function () {
        state.weights[group.key] = parseFloat(slider.value);
        value.textContent = parseFloat(slider.value).toFixed(1);
        recompute();
      });

      row.appendChild(label);
      row.appendChild(slider);
      frag.appendChild(row);
    });
    els.weights.innerHTML = "";
    els.weights.appendChild(frag);
  }

  function renderOutliers() {
    var frag = document.createDocumentFragment();
    var suppress = state.data.participants.length < 3;
    var maxScore = state.order.length ? state.scores[state.order[0]] : 0;
    if (!suppress) {
      state.order.forEach(function (idx) {
        var li = document.createElement("li");
        li.className = "map-outlier-item" +
          (idx === state.selected ? " is-selected" : "");

        var pid = document.createElement("span");
        pid.className = "map-outlier-pid";
        pid.textContent = state.data.participants[idx];

        var track = document.createElement("span");
        track.className = "map-outlier-bar-track";
        var bar = document.createElement("span");
        bar.className = "map-outlier-bar";
        bar.style.width = (maxScore > 0 ? (state.scores[idx] / maxScore) * 100 : 0) + "%";
        track.appendChild(bar);

        var score = document.createElement("span");
        score.className = "map-outlier-score";
        score.textContent = state.scores[idx].toFixed(2);

        li.appendChild(pid);
        li.appendChild(track);
        li.appendChild(score);
        li.addEventListener("click", function () {
          selectParticipant(idx === state.selected ? -1 : idx);
        });
        frag.appendChild(li);
      });
    }
    els.outliers.innerHTML = "";
    els.outliers.appendChild(frag);
  }

  function renderAxisLegend() {
    var frag = document.createDocumentFragment();
    var axisNames = ["X", "Y", "Z"];
    var dl = document.createElement("dl");
    state.components.forEach(function (comp, c) {
      var loadings = comp
        .map(function (v, j) { return { j: j, v: v }; })
        .sort(function (a, b) { return Math.abs(b.v) - Math.abs(a.v); })
        .slice(0, 3);
      var dt = document.createElement("dt");
      dt.textContent = axisNames[c] + " — " +
        Math.round(state.variance[c] * 100) + "% of variance";
      var dd = document.createElement("dd");
      dd.textContent = loadings
        .map(function (l) {
          return (l.v >= 0 ? "+" : "−") + " " + state.data.columns[l.j].label;
        })
        .join(", ");
      dl.appendChild(dt);
      dl.appendChild(dd);
    });
    frag.appendChild(dl);
    els.axisList.innerHTML = "";
    els.axisList.appendChild(frag);
  }

  function renderExplain(idx) {
    var data = state.data;
    var pid = data.participants[idx];
    var avail = data.availability[pid] || {};
    els.explainTitle.textContent = pid;

    var frag = document.createDocumentFragment();

    // Availability chips + outlier rank.
    var chips = document.createElement("div");
    chips.className = "map-chip-row";
    data.groups.forEach(function (group) {
      var chip = document.createElement("span");
      chip.className = "map-chip" + (avail[group.key] ? "" : " is-missing");
      chip.textContent = group.label;
      if (!avail[group.key]) chip.title = "No " + group.label + " data for " + pid;
      chips.appendChild(chip);
    });
    var rankChip = document.createElement("span");
    rankChip.className = "map-chip";
    rankChip.textContent = "unusualness " + state.scores[idx].toFixed(2) +
      " · #" + (state.order.indexOf(idx) + 1) + " of " + data.participants.length;
    chips.appendChild(rankChip);
    frag.appendChild(chips);

    // Per-group share of the squared distance: which lens makes them unusual.
    var contributions = {};
    var total = 0;
    var j;
    for (j = 0; j < data.columns.length; j++) {
      if (state.zRaw[idx][j] == null) continue;
      var wz = state.zRaw[idx][j] * columnWeight(data.columns[j]);
      contributions[data.columns[j].group] =
        (contributions[data.columns[j].group] || 0) + wz * wz;
      total += wz * wz;
    }
    var groupHead = document.createElement("div");
    groupHead.className = "map-explain-section";
    groupHead.textContent = "Distance by signal group";
    frag.appendChild(groupHead);
    var groupBars = document.createElement("div");
    groupBars.className = "map-group-bars";
    data.groups.forEach(function (group) {
      var share = total > 0 ? (contributions[group.key] || 0) / total : 0;
      var row = document.createElement("div");
      row.className = "map-feature-row";
      var label = document.createElement("div");
      label.className = "map-feature-label";
      var name = document.createElement("span");
      name.textContent = group.label;
      var pct = document.createElement("span");
      pct.className = "map-feature-z";
      pct.textContent = avail[group.key] ? Math.round(share * 100) + "%" : "no data";
      label.appendChild(name);
      label.appendChild(pct);
      var track = document.createElement("div");
      track.className = "map-outlier-bar-track";
      var bar = document.createElement("span");
      bar.className = "map-outlier-bar";
      bar.style.display = "block";
      bar.style.width = (share * 100) + "%";
      track.appendChild(bar);
      row.appendChild(label);
      row.appendChild(track);
      groupBars.appendChild(row);
    });
    frag.appendChild(groupBars);

    // Top features by |weighted z| — what actually pushes this dot away.
    var ranked = [];
    for (j = 0; j < data.columns.length; j++) {
      if (state.zRaw[idx][j] == null) continue;
      var wzv = state.zRaw[idx][j] * columnWeight(data.columns[j]);
      if (Math.abs(wzv) < 1e-9) continue;
      ranked.push({ j: j, wz: wzv, z: state.zRaw[idx][j] });
    }
    ranked.sort(function (a, b) { return Math.abs(b.wz) - Math.abs(a.wz); });

    var featHead = document.createElement("div");
    featHead.className = "map-explain-section";
    featHead.textContent = "What sets " + pid + " apart";
    frag.appendChild(featHead);

    ranked.slice(0, TOP_FEATURES).forEach(function (item) {
      var col = data.columns[item.j];
      var row = document.createElement("div");
      row.className = "map-feature-row";

      var label = document.createElement("div");
      label.className = "map-feature-label";
      var name = document.createElement("span");
      name.textContent = col.label;
      var zEl = document.createElement("span");
      zEl.className = "map-feature-z";
      zEl.textContent = (item.z >= 0 ? "+" : "−") +
        Math.abs(item.z).toFixed(1) + "σ";
      label.appendChild(name);
      label.appendChild(zEl);

      var detail = document.createElement("div");
      detail.className = "map-feature-detail";
      detail.textContent = fmtNum(data.matrix[idx][item.j]) +
        " vs cohort mean " + fmtNum(state.stats[item.j].mean);

      // Signed bar from the center line; half-width = the clamp bound.
      var track = document.createElement("div");
      track.className = "map-feature-bar-track";
      var bar = document.createElement("div");
      bar.className = "map-feature-bar" + (item.z < 0 ? " is-negative" : "");
      var half = Math.min(Math.abs(item.z) / Z_CLAMP, 1) * 50;
      if (item.z >= 0) {
        bar.style.left = "50%";
        bar.style.width = half + "%";
      } else {
        bar.style.left = (50 - half) + "%";
        bar.style.width = half + "%";
      }
      track.appendChild(bar);

      row.appendChild(label);
      row.appendChild(detail);
      row.appendChild(track);
      frag.appendChild(row);
    });

    if (!ranked.length) {
      var none = document.createElement("p");
      none.className = "map-feature-detail";
      none.textContent = "Sits at the cohort mean on every available signal.";
      frag.appendChild(none);
    }

    // Plain links out; hash deep-links are a later enhancement.
    var links = document.createElement("div");
    links.className = "map-explain-links";
    [
      ["Transcripts", "/transcripts/"],
      ["Screenspace", "/screenspace/"],
      ["Studio", "/studio/"],
    ].forEach(function (pair) {
      var a = document.createElement("a");
      a.textContent = pair[0] + " →";
      a.href = pair[1];
      links.appendChild(a);
    });
    frag.appendChild(links);

    els.explainBody.innerHTML = "";
    els.explainBody.appendChild(frag);
  }

  // ---- Boot -----------------------------------------------------------------

  function initDom() {
    els.canvasWrap = document.getElementById("mapCanvasWrap");
    els.canvas = document.getElementById("mapCanvas");
    els.labels = document.getElementById("mapLabels");
    els.tooltip = document.getElementById("mapTooltip");
    els.empty = document.getElementById("mapEmpty");
    els.weights = document.getElementById("mapWeights");
    els.outliers = document.getElementById("mapOutliers");
    els.axisList = document.getElementById("mapAxisList");
    els.notice = document.getElementById("mapNotice");
    els.explain = document.getElementById("mapExplain");
    els.explainTitle = document.getElementById("mapExplainTitle");
    els.explainBody = document.getElementById("mapExplainBody");
    document.getElementById("mapExplainClose")
      .addEventListener("click", function () { selectParticipant(-1); });
    loadData();
  }

  initDom();
})();
