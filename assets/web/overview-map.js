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
  var SIM_EDGE_K = 3;       // each participant links to its k nearest peers
  var BURST_CAP = 40;       // max in-scene satellites per burst (drawer shows all)
  var GOLDEN_ANGLE = 2.39996; // radians; spreads burst items on a sphere shell

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
    showSimEdges: true,   // similarity-link layer toggle
    showAnchors: false,   // shared-anchor layer toggle (busier; default off)
    showAllMoments: false, // every participant's items as a point cloud
    mutedFeatures: {},    // column key -> true; muted = weight 0 everywhere
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
    simEdges: null, // {mesh, edges: [{a, b, sim}]} similarity-link layer
    anchors: null,  // {features, meshes, labels, links} shared-anchor layer
    burst: null,    // {owner, meshes, items, offsets} drill-down satellites
    moments: null,  // {mesh, owners, offsets, norms, baseColors} point cloud
    axisLabels: [], // [{el, pos}] X/Y/Z tip labels tied to the axis legend
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
    if (state.mutedFeatures[col.key]) return 0;
    var w = state.weights[col.group];
    if (w == null) w = 1;
    var size = state.groupSizes[col.group] || 1;
    return w / Math.sqrt(size);
  }

  function setFeatureMuted(key, muted) {
    if (muted) state.mutedFeatures[key] = true;
    else delete state.mutedFeatures[key];
    renderMutedChips();
    recompute();
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
      rebuildLayers();
      requestRender();
    }
    if (state.selected >= 0) renderExplain(state.selected);
  }

  // ---- Similarity-link layer ----------------------------------------------
  //
  // Lines between each participant and its k most-similar peers, computed in
  // the FULL weighted feature space (same matrix the layout projects from),
  // so a link means "alike overall", not "happen to project nearby". The
  // edge set is the union of every node's k nearest (non-mutual — mutual-only
  // would starve exactly the outliers you want context for). Line topology
  // rebuilds on every recompute; per-frame position updates follow the dots
  // (including through the re-layout lerp) via syncEdgePositions().

  function computeSimilarityEdges() {
    var W = state.weighted;
    var n = W.length;
    var edges = [];
    if (n < 2) return edges;

    var dist = [];
    var i, j, k;
    for (i = 0; i < n; i++) dist.push(new Array(n));
    for (i = 0; i < n; i++) {
      dist[i][i] = 0;
      for (j = i + 1; j < n; j++) {
        var s = 0;
        for (k = 0; k < W[i].length; k++) {
          var d = W[i][k] - W[j][k];
          s += d * d;
        }
        var dd = Math.sqrt(s);
        dist[i][j] = dd;
        dist[j][i] = dd;
      }
    }

    var seen = {};
    for (i = 0; i < n; i++) {
      var order = [];
      for (j = 0; j < n; j++) if (j !== i) order.push(j);
      // Deterministic tie-break on index keeps the layout comparable.
      order.sort(function (a, b) { return dist[i][a] - dist[i][b] || a - b; });
      for (k = 0; k < Math.min(SIM_EDGE_K, order.length); k++) {
        var a = Math.min(i, order[k]);
        var b = Math.max(i, order[k]);
        var key = a + "-" + b;
        if (seen[key]) continue;
        seen[key] = true;
        edges.push({ a: a, b: b, sim: 1 / (1 + dist[a][b]) });
      }
    }

    // Min-max normalize similarity into [0.15, 1] so the faintest edge is
    // still visible while the strongest reads clearly brighter.
    var lo = Infinity, hi = -Infinity;
    for (i = 0; i < edges.length; i++) {
      lo = Math.min(lo, edges[i].sim);
      hi = Math.max(hi, edges[i].sim);
    }
    var span = hi - lo;
    for (i = 0; i < edges.length; i++) {
      edges[i].sim = span > 1e-12 ? 0.15 + 0.85 * ((edges[i].sim - lo) / span) : 1;
    }
    return edges;
  }

  function disposeLayer(entry) {
    if (!entry || !entry.mesh) return;
    three.scene.remove(entry.mesh);
    entry.mesh.geometry.dispose();
    entry.mesh.material.dispose();
  }

  // Rebuild every derived scene layer after a re-layout. Weight-slider drags
  // fire this rapidly, so old geometries/materials are disposed each time.
  function rebuildLayers() {
    rebuildSimEdges();
    rebuildAnchors();
  }

  function rebuildSimEdges() {
    disposeLayer(three.simEdges);
    three.simEdges = null;
    var edges = computeSimilarityEdges();
    if (!edges.length) return;

    var positions = new Float32Array(edges.length * 6);
    var colors = new Float32Array(edges.length * 6);
    var c = new THREE.Color();
    // LineBasicMaterial has no per-vertex alpha: fake per-edge opacity by
    // lerping the edge color toward the scene background by (1 - sim).
    for (var e = 0; e < edges.length; e++) {
      c.copy(three.colors.base).lerp(three.colors.bg, 1 - edges[e].sim);
      for (var v = 0; v < 2; v++) {
        colors[e * 6 + v * 3] = c.r;
        colors[e * 6 + v * 3 + 1] = c.g;
        colors[e * 6 + v * 3 + 2] = c.b;
      }
    }

    var geo = new THREE.BufferGeometry();
    geo.setAttribute("position", new THREE.BufferAttribute(positions, 3));
    geo.setAttribute("color", new THREE.BufferAttribute(colors, 3));
    var mat = new THREE.LineBasicMaterial({
      vertexColors: true, transparent: true, opacity: 0.85,
    });
    var mesh = new THREE.LineSegments(geo, mat);
    mesh.visible = !!state.showSimEdges;
    three.scene.add(mesh);
    three.simEdges = { mesh: mesh, edges: edges };
    syncEdgePositions();
  }

  function syncEdgePositions() {
    if (!three.simEdges) return;
    var arr = three.simEdges.mesh.geometry.attributes.position.array;
    var edges = three.simEdges.edges;
    for (var e = 0; e < edges.length; e++) {
      var pa = three.dots[edges[e].a].position;
      var pb = three.dots[edges[e].b].position;
      arr[e * 6] = pa.x;
      arr[e * 6 + 1] = pa.y;
      arr[e * 6 + 2] = pa.z;
      arr[e * 6 + 3] = pb.x;
      arr[e * 6 + 4] = pb.y;
      arr[e * 6 + 5] = pb.z;
    }
    three.simEdges.mesh.geometry.attributes.position.needsUpdate = true;
  }

  // ---- Shared-anchor layer --------------------------------------------------
  //
  // Small anchor nodes for the interpretable dynamic features (observation
  // categories + detector rates), each placed at the value-weighted centroid
  // of the participants exhibiting it, with faint participant->anchor lines.
  // Anchors follow the layout deterministically and answer WHY dots cluster
  // ("these three sit together because they share the nav category"). A
  // feature needs >= 2 exhibitors to earn an anchor — a one-exhibitor anchor
  // would just shadow that participant's dot.

  function anchorShortName(col) {
    return col.label
      .replace(" share of observations", "")
      .replace(" events/min", "");
  }

  function anchorFeatures() {
    var data = state.data;
    var out = [];
    for (var j = 0; j < data.columns.length; j++) {
      var key = data.columns[j].key;
      if (key.indexOf("obs_cat_") !== 0 && key.indexOf("ss_rate_") !== 0) continue;
      if (state.mutedFeatures[key]) continue; // muted features earn no anchor
      var exhibitors = 0;
      for (var i = 0; i < data.matrix.length; i++) {
        var v = data.matrix[i][j];
        if (v != null && v > 0) exhibitors++;
      }
      if (exhibitors >= 2) {
        out.push({
          j: j,
          group: data.columns[j].group,
          name: anchorShortName(data.columns[j]),
        });
      }
    }
    return out;
  }

  function disposeAnchors() {
    if (!three.anchors) return;
    var a = three.anchors;
    var i;
    for (i = 0; i < a.meshes.length; i++) {
      three.scene.remove(a.meshes[i]);
      a.meshes[i].geometry.dispose();
      a.meshes[i].material.dispose();
    }
    disposeLayer(a.links);
    for (i = 0; i < a.labels.length; i++) {
      if (a.labels[i].parentNode) a.labels[i].parentNode.removeChild(a.labels[i]);
    }
    three.anchors = null;
  }

  function rebuildAnchors() {
    disposeAnchors();
    var features = anchorFeatures();
    if (!features.length) return;
    var data = state.data;
    var visible = !!state.showAnchors;

    var meshes = [];
    var labels = [];
    var pairs = [];
    var labelFrag = document.createDocumentFragment();
    var geo = new THREE.OctahedronGeometry(0.22);
    var f, i;
    for (f = 0; f < features.length; f++) {
      var color = features[f].group === "observations"
        ? three.colors.obsAnchor : three.colors.ssAnchor;
      var mesh = new THREE.Mesh(
        geo, new THREE.MeshBasicMaterial({ color: color })
      );
      mesh.visible = visible;
      three.scene.add(mesh);
      meshes.push(mesh);

      var label = document.createElement("div");
      label.className = "map-label map-anchor-label";
      label.textContent = features[f].name;
      labelFrag.appendChild(label);
      labels.push(label);

      for (i = 0; i < data.matrix.length; i++) {
        var v = data.matrix[i][features[f].j];
        if (v != null && v > 0) pairs.push({ p: i, a: f });
      }
    }
    els.labels.appendChild(labelFrag);

    var linkGeo = new THREE.BufferGeometry();
    linkGeo.setAttribute(
      "position", new THREE.BufferAttribute(new Float32Array(pairs.length * 6), 3)
    );
    var linkMesh = new THREE.LineSegments(
      linkGeo,
      new THREE.LineBasicMaterial({
        color: three.colors.axis, transparent: true, opacity: 0.35,
      })
    );
    linkMesh.visible = visible;
    three.scene.add(linkMesh);

    three.anchors = {
      features: features, meshes: meshes, labels: labels,
      links: { mesh: linkMesh, pairs: pairs },
    };
    placeAnchors();
  }

  // Anchor position = value-weighted centroid of its exhibitors' CURRENT dot
  // positions — recomputed per frame so anchors travel with the lerp. Cheap:
  // <= ~25 anchors x ~30 participants.
  function placeAnchors() {
    if (!three.anchors) return;
    var a = three.anchors;
    var data = state.data;
    var f, i;
    for (f = 0; f < a.features.length; f++) {
      var j = a.features[f].j;
      var sx = 0, sy = 0, sz = 0, sw = 0;
      for (i = 0; i < data.matrix.length; i++) {
        var v = data.matrix[i][j];
        if (v == null || v <= 0) continue;
        var p = three.dots[i].position;
        sx += v * p.x;
        sy += v * p.y;
        sz += v * p.z;
        sw += v;
      }
      if (sw > 0) a.meshes[f].position.set(sx / sw, sy / sw, sz / sw);
    }

    var arr = a.links.mesh.geometry.attributes.position.array;
    for (i = 0; i < a.links.pairs.length; i++) {
      var pp = three.dots[a.links.pairs[i].p].position;
      var pa = a.meshes[a.links.pairs[i].a].position;
      arr[i * 6] = pp.x;
      arr[i * 6 + 1] = pp.y;
      arr[i * 6 + 2] = pp.z;
      arr[i * 6 + 3] = pa.x;
      arr[i * 6 + 4] = pa.y;
      arr[i * 6 + 5] = pa.z;
    }
    a.links.mesh.geometry.attributes.position.needsUpdate = true;
  }

  function setAnchorsVisible(on) {
    if (!three.anchors) return;
    var a = three.anchors;
    for (var i = 0; i < a.meshes.length; i++) a.meshes[i].visible = on;
    a.links.mesh.visible = on;
    // Label visibility is applied by updateLabels() on the next render.
  }

  // ---- All-moments point cloud ----------------------------------------------
  //
  // Every participant's items (sampled to BURST_CAP each) as one THREE.Points
  // cloud on tight golden-angle shells around their dots — the whole study's
  // timestamps visible at once, colored by source. Independent of weights
  // (items don't change on recompute), so it rebuilds only on toggle-on and
  // tab re-activation; positions follow the dots per frame. Per-point
  // normalized session times + base colors are kept for the replay glow.

  function rebuildMoments() {
    disposeLayer(three.moments);
    three.moments = null;
    if (!state.data || !state.data.participants.length) return;

    var owners = [];
    var offsets = [];
    var colors = [];
    var norms = [];
    var i, k;
    for (i = 0; i < state.data.participants.length; i++) {
      var items = buildParticipantItems(state.data.participants[i]);
      if (!items.length) continue;
      var shown = items;
      if (items.length > BURST_CAP) {
        shown = [];
        var step = items.length / BURST_CAP;
        for (k = 0; k < BURST_CAP; k++) shown.push(items[Math.floor(k * step)]);
      }
      var maxEnd = 0;
      for (k = 0; k < items.length; k++) maxEnd = Math.max(maxEnd, items[k].end);
      for (k = 0; k < shown.length; k++) {
        var y = 1 - 2 * (k + 0.5) / shown.length;
        var ring = Math.sqrt(Math.max(0, 1 - y * y));
        var theta = k * GOLDEN_ANGLE;
        var r = 0.7 + 0.35 * ((k * 0.41) % 1);
        offsets.push(new THREE.Vector3(
          r * ring * Math.cos(theta), r * y, r * ring * Math.sin(theta)
        ));
        owners.push(i);
        norms.push(maxEnd > 0 ? shown[k].start / maxEnd : 0);
        var c = burstColor(shown[k].source);
        colors.push(c.r, c.g, c.b);
      }
    }
    if (!owners.length) return;

    var geo = new THREE.BufferGeometry();
    geo.setAttribute(
      "position", new THREE.BufferAttribute(new Float32Array(owners.length * 3), 3)
    );
    geo.setAttribute("color", new THREE.BufferAttribute(new Float32Array(colors), 3));
    var mat = new THREE.PointsMaterial({
      size: 0.16, vertexColors: true, transparent: true, opacity: 0.8,
    });
    var mesh = new THREE.Points(geo, mat);
    mesh.visible = !!state.showAllMoments;
    three.scene.add(mesh);
    three.moments = {
      mesh: mesh,
      owners: owners,
      offsets: offsets,
      norms: norms,
      baseColors: new Float32Array(colors),
    };
    syncMomentsPositions();
  }

  function syncMomentsPositions() {
    if (!three.moments) return;
    var arr = three.moments.mesh.geometry.attributes.position.array;
    for (var k = 0; k < three.moments.owners.length; k++) {
      var d = three.dots[three.moments.owners[k]].position;
      var o = three.moments.offsets[k];
      arr[k * 3] = d.x + o.x;
      arr[k * 3 + 1] = d.y + o.y;
      arr[k * 3 + 2] = d.z + o.z;
    }
    three.moments.mesh.geometry.attributes.position.needsUpdate = true;
  }

  // ---- Session replay --------------------------------------------------------
  //
  // The mindwalk idea: sweep a playhead over normalized session time (each
  // participant normalized by their own session length) and let the map glow
  // where activity clusters — dots brighten and pulse by how many of their
  // items sit near the playhead, and the all-moments cloud (when on) dims
  // everything except the moments around it. Playback free-runs a rAF loop
  // (this is an animation; the render loop stays on-demand otherwise).

  var REPLAY_WINDOW = 0.05;        // kernel half-width, normalized time
  var REPLAY_DURATION_MS = 30000;  // full session sweep duration

  var replay = {
    playing: false,
    t: 0,
    norms: null, // [participant idx] -> array of normalized item start times
    lastTs: 0,
  };

  function ensureReplayData() {
    if (replay.norms) return Promise.resolve();
    return window.ClipgenOverview.ensureData().then(function () {
      if (replay.norms) return;
      var norms = [];
      for (var i = 0; i < state.data.participants.length; i++) {
        var items = buildParticipantItems(state.data.participants[i]);
        var maxEnd = 0;
        var k;
        for (k = 0; k < items.length; k++) maxEnd = Math.max(maxEnd, items[k].end);
        var arr = [];
        for (k = 0; k < items.length; k++) {
          arr.push(maxEnd > 0 ? items[k].start / maxEnd : 0);
        }
        norms.push(arr);
      }
      replay.norms = norms;
    });
  }

  function replayIntensity(i, t) {
    var arr = (replay.norms && replay.norms[i]) || [];
    var count = 0;
    for (var k = 0; k < arr.length; k++) {
      if (Math.abs(arr[k] - t) <= REPLAY_WINDOW) count++;
    }
    return Math.min(count / 3, 1);
  }

  var _replayWhite = null;

  function applyReplayGlow() {
    if (!replay.norms || !three.renderer) return;
    if (!_replayWhite) _replayWhite = new THREE.Color("#ffffff");
    var t = replay.t;
    var maxScore = 0;
    var i;
    for (i = 0; i < state.scores.length; i++) {
      maxScore = Math.max(maxScore, state.scores[i]);
    }
    for (i = 0; i < three.dots.length; i++) {
      var tScore = maxScore > 0 ? state.scores[i] / maxScore : 0;
      var glow = replayIntensity(i, t);
      var c = three.dots[i].material.color;
      c.copy(three.colors.base).lerp(three.colors.hot, tScore);
      if (glow > 0) c.lerp(_replayWhite, glow * 0.7);
      var scale = (i === state.selected ? 1.4 : 1) *
        (0.85 + tScore * 0.5) * (1 + glow * 0.4);
      three.dots[i].scale.setScalar(scale);
    }
    if (three.moments && state.showAllMoments) {
      var colors = three.moments.mesh.geometry.attributes.color.array;
      var base = three.moments.baseColors;
      for (var k = 0; k < three.moments.norms.length; k++) {
        var f = Math.abs(three.moments.norms[k] - t) <= REPLAY_WINDOW ? 1 : 0.2;
        colors[k * 3] = base[k * 3] * f;
        colors[k * 3 + 1] = base[k * 3 + 1] * f;
        colors[k * 3 + 2] = base[k * 3 + 2] * f;
      }
      three.moments.mesh.geometry.attributes.color.needsUpdate = true;
    }
    if (!lerp.active) positionDots(state.coords); // re-sync halo scale
    requestRender();
  }

  // Back to the idle look: outlier ramp colors, no pulse, full-bright cloud.
  function clearReplayGlow() {
    styleDots();
    if (three.moments) {
      three.moments.mesh.geometry.attributes.color.array.set(three.moments.baseColors);
      three.moments.mesh.geometry.attributes.color.needsUpdate = true;
    }
    if (!lerp.active && state.coords) positionDots(state.coords);
    requestRender();
  }

  function syncReplayUI() {
    if (els.replayScrub) els.replayScrub.value = String(replay.t);
    if (els.replayPos) els.replayPos.textContent = Math.round(replay.t * 100) + "%";
    if (els.replayPlay) els.replayPlay.textContent = replay.playing ? "Pause" : "Play";
  }

  function replayStep(ts) {
    if (!replay.playing) return;
    if (!replay.lastTs) replay.lastTs = ts;
    replay.t = Math.min(replay.t + (ts - replay.lastTs) / REPLAY_DURATION_MS, 1);
    replay.lastTs = ts;
    syncReplayUI();
    applyReplayGlow();
    if (replay.t >= 1) {
      setReplayPlaying(false);
      return;
    }
    requestAnimationFrame(replayStep);
  }

  function setReplayPlaying(playing) {
    replay.playing = playing;
    replay.lastTs = 0;
    syncReplayUI();
    if (!playing) return;
    if (replay.t >= 1) replay.t = 0; // replay from the top after a full sweep
    ensureReplayData().then(function () {
      if (replay.playing) requestAnimationFrame(replayStep);
    });
  }

  function initReplayControls() {
    els.replayPlay = document.getElementById("mapReplayPlay");
    els.replayScrub = document.getElementById("mapReplayScrub");
    els.replayPos = document.getElementById("mapReplayPos");
    if (!els.replayPlay) return;
    els.replayPlay.addEventListener("click", function () {
      setReplayPlaying(!replay.playing);
    });
    els.replayScrub.addEventListener("input", function () {
      replay.playing = false;
      replay.t = parseFloat(els.replayScrub.value) || 0;
      syncReplayUI();
      if (replay.t <= 0) {
        clearReplayGlow();
        return;
      }
      ensureReplayData().then(applyReplayGlow);
    });
  }

  // ---- Participant drill-down: items, satellite burst, timeline drawer -----
  //
  // The underlying timestamps and notes come from the Overview hub's state —
  // the same streams Convergence/Metadata read: sheet observations (rows +
  // OV.parseClipTimestamps for baselined times), clustered screenspace
  // events, clustered transcript marks, and LLM friction moments (the hub
  // fetches api/friction-moments, which resolves segment ids to times).

  function buildParticipantItems(pid) {
    var OV = window.ClipgenOverview;
    var hub = OV && OV.state;
    var items = [];
    if (!hub) return items;
    var i;

    if (hub.sheetData && hub.sheetData.rows) {
      for (i = 0; i < hub.sheetData.rows.length; i++) {
        var row = hub.sheetData.rows[i];
        var cell = row.cells && row.cells[pid];
        if (!cell || !cell.valid) continue;
        var segs = OV.parseClipTimestamps(cell.value, pid);
        for (var s = 0; s < segs.length; s++) {
          items.push({
            source: "sheet",
            text: row.observation || "",
            category: row.category || "",
            severity: row.severity || "",
            start: segs[s].startSeconds,
            end: segs[s].startSeconds + segs[s].duration,
          });
        }
      }
    }

    for (i = 0; i < hub.intakeClusters.length; i++) {
      var sc = hub.intakeClusters[i];
      if (sc.participant !== pid) continue;
      items.push({
        source: "screenspace",
        text: sc.label || sc.event_type || sc.detector || "",
        category: sc.detector || "",
        severity: "",
        start: sc.start,
        end: sc.end,
      });
    }

    for (i = 0; i < hub.trIntakeClusters.length; i++) {
      var tc = hub.trIntakeClusters[i];
      if (tc.participant !== pid) continue;
      items.push({
        source: "transcript",
        text: tc.text || tc.label || "",
        category: tc.category || "",
        severity: tc.severity || "",
        start: tc.start,
        end: tc.end,
      });
    }

    for (i = 0; i < hub.frictionMoments.length; i++) {
      var fm = hub.frictionMoments[i];
      if (fm.participant !== pid) continue;
      items.push({
        source: "friction",
        text: fm.rationale || "",
        category: fm.category || "",
        severity: "",
        start: fm.start,
        end: fm.end,
      });
    }

    items.sort(function (a, b) { return a.start - b.start || a.end - b.end; });
    return items;
  }

  function burstColor(source) {
    if (source === "sheet") return three.colors.streamSheet;
    if (source === "screenspace") return three.colors.streamScreenspace;
    if (source === "friction") return three.colors.friction;
    return three.colors.streamTranscript;
  }

  function clearBurst() {
    if (!three.burst) return;
    for (var i = 0; i < three.burst.meshes.length; i++) {
      three.scene.remove(three.burst.meshes[i]);
      three.burst.meshes[i].geometry.dispose();
      three.burst.meshes[i].material.dispose();
    }
    three.burst = null;
  }

  // Golden-angle spiral on a sphere shell around the selected dot. Items are
  // capped for the scene (evenly sampled across session time); the drawer
  // below always shows the full list.
  function showBurst(idx, items) {
    clearBurst();
    if (!items.length) return;

    var shown = items;
    if (items.length > BURST_CAP) {
      shown = [];
      var step = items.length / BURST_CAP;
      for (var k = 0; k < BURST_CAP; k++) {
        shown.push(items[Math.floor(k * step)]);
      }
    }

    var geo = new THREE.SphereGeometry(0.12, 12, 8);
    var meshes = [];
    var offsets = [];
    var n = shown.length;
    for (var i = 0; i < n; i++) {
      var y = 1 - 2 * (i + 0.5) / n;
      var ring = Math.sqrt(Math.max(0, 1 - y * y));
      var theta = i * GOLDEN_ANGLE;
      // Vary the shell radius deterministically so items never sit coplanar.
      var r = 1.1 + 0.5 * ((i * 0.37) % 1);
      var offset = new THREE.Vector3(
        r * ring * Math.cos(theta), r * y, r * ring * Math.sin(theta)
      );
      var mesh = new THREE.Mesh(
        geo, new THREE.MeshBasicMaterial({ color: burstColor(shown[i].source) })
      );
      mesh.userData.burstItem = shown[i];
      mesh.position.copy(three.dots[idx].position).add(offset);
      three.scene.add(mesh);
      meshes.push(mesh);
      offsets.push(offset);
    }
    three.burst = { owner: idx, meshes: meshes, items: shown, offsets: offsets };
  }

  function syncBurstPositions() {
    if (!three.burst) return;
    var dot = three.dots[three.burst.owner].position;
    for (var i = 0; i < three.burst.meshes.length; i++) {
      three.burst.meshes[i].position.copy(dot).add(three.burst.offsets[i]);
    }
  }

  // ---- Session-timeline drawer ---------------------------------------------

  var DRAWER_SOURCES = [
    { key: "sheet", label: "Sheet" },
    { key: "screenspace", label: "Screenspace" },
    { key: "transcript", label: "Transcript" },
    { key: "friction", label: "Friction" },
  ];

  function drawerCssColor(source) {
    if (source === "friction") return "var(--color-friction)";
    return "var(--stream-" + source + ")";
  }

  function showDrawer(pid, items) {
    els.drawerTitle.textContent = pid;
    var capped = items.length > BURST_CAP
      ? " (scene shows " + BURST_CAP + " of them)" : "";
    els.drawerCount.textContent =
      items.length + " moment" + (items.length === 1 ? "" : "s") + capped;
    els.drawerNote.classList.add("hidden");

    var duration = 0;
    var i;
    for (i = 0; i < items.length; i++) duration = Math.max(duration, items[i].end);

    var frag = document.createDocumentFragment();
    if (!items.length) {
      var none = document.createElement("p");
      none.className = "map-feature-detail";
      none.textContent = "No timestamps, events, or marks for this participant yet.";
      frag.appendChild(none);
    }
    DRAWER_SOURCES.forEach(function (src) {
      var laneItems = items.filter(function (it) { return it.source === src.key; });
      if (!laneItems.length) return;
      var lane = document.createElement("div");
      lane.className = "map-drawer-lane";
      var label = document.createElement("span");
      label.className = "map-drawer-lane-label";
      label.textContent = src.label;
      var track = document.createElement("div");
      track.className = "map-drawer-track";
      laneItems.forEach(function (item) {
        var chip = document.createElement("button");
        chip.type = "button";
        chip.className = "map-drawer-item";
        chip.style.background = drawerCssColor(src.key);
        var left = duration > 0 ? (item.start / duration) * 100 : 0;
        var width = duration > 0 ? ((item.end - item.start) / duration) * 100 : 0;
        chip.style.left = Math.min(left, 99) + "%";
        chip.style.width = Math.max(width, 0.8) + "%";
        chip.title = formatTime(item.start) + " · " + (item.text || item.category);
        chip.addEventListener("click", function () { renderDrawerNote(item); });
        track.appendChild(chip);
      });
      lane.appendChild(label);
      lane.appendChild(track);
      frag.appendChild(lane);
    });
    els.drawerLanes.innerHTML = "";
    els.drawerLanes.appendChild(frag);
    els.drawer.classList.remove("hidden");
  }

  function renderDrawerNote(item) {
    var host = els.drawerNote;
    host.innerHTML = "";
    var head = document.createElement("div");
    head.className = "map-drawer-note-head";
    var when = document.createElement("span");
    when.className = "map-drawer-note-time";
    when.textContent = formatTime(item.start) +
      (item.end > item.start ? " – " + formatTime(item.end) : "");
    head.appendChild(when);
    var meta = [item.source, item.category, item.severity]
      .filter(function (x) { return x; }).join(" · ");
    var metaEl = document.createElement("span");
    metaEl.className = "map-drawer-note-meta";
    metaEl.textContent = meta;
    head.appendChild(metaEl);
    var body = document.createElement("div");
    body.className = "map-drawer-note-text";
    body.textContent = item.text || "(no note text)";
    host.appendChild(head);
    host.appendChild(body);
    host.classList.remove("hidden");
  }

  function hideDrawer() {
    els.drawer.classList.add("hidden");
    els.drawerNote.classList.add("hidden");
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
    renderLayerToggles();
    initReplayControls();
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
      obsAnchor: tokenColor("--severity-positive", "#4ade80"),
      ssAnchor: tokenColor("--cell-data-bright", "#2bc8c8"),
      streamSheet: tokenColor("--stream-sheet", "#eab308"),
      streamScreenspace: tokenColor("--stream-screenspace", "#3498db"),
      streamTranscript: tokenColor("--stream-transcript", "#10a34a"),
      friction: tokenColor("--color-friction", "#ea580c"),
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

    // X/Y/Z tip labels — same names the sidebar axis legend explains.
    var axisNames = ["X", "Y", "Z"];
    var axisTips = [
      new THREE.Vector3(r, 0, 0),
      new THREE.Vector3(0, r, 0),
      new THREE.Vector3(0, 0, r),
    ];
    for (var ai = 0; ai < 3; ai++) {
      var axLabel = document.createElement("div");
      axLabel.className = "map-label map-axis-tip";
      axLabel.textContent = axisNames[ai];
      els.labels.appendChild(axLabel);
      three.axisLabels.push({ el: axLabel, pos: axisTips[ai] });
    }

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
    syncEdgePositions();
    placeAnchors();
    syncBurstPositions();
    syncMomentsPositions();
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

  // Burst satellites are raycast before participant dots (they are smaller
  // and sit in front of their owner). Returns {type: "dot"|"burst", ...}.
  function pickTarget(e) {
    var rect = three.renderer.domElement.getBoundingClientRect();
    if (e.clientX < rect.left || e.clientX > rect.right ||
        e.clientY < rect.top || e.clientY > rect.bottom) return null;
    var ndc = new THREE.Vector2(
      ((e.clientX - rect.left) / rect.width) * 2 - 1,
      -((e.clientY - rect.top) / rect.height) * 2 + 1);
    three.raycaster.setFromCamera(ndc, three.camera);
    if (three.burst) {
      var bursts = three.raycaster.intersectObjects(three.burst.meshes, false);
      if (bursts.length) {
        return { type: "burst", item: bursts[0].object.userData.burstItem };
      }
    }
    var hits = three.raycaster.intersectObjects(three.dots, false);
    if (hits.length) return { type: "dot", index: hits[0].object.userData.index };
    return null;
  }

  function pick(e) {
    var target = pickTarget(e);
    return target && target.type === "dot" ? target.index : -1;
  }

  function updateHover(e) {
    if (!three.renderer) return;
    var target = pickTarget(e);

    if (target && target.type === "burst") {
      state.hovered = -1;
      three.renderer.domElement.style.cursor = "pointer";
      var item = target.item;
      var text = item.text || item.category || item.source;
      if (text.length > 90) text = text.substring(0, 90) + "…";
      els.tooltip.textContent = formatTime(item.start) + " · " + text;
      els.tooltip.classList.remove("hidden");
      moveTooltip(e);
      return;
    }

    var idx = target && target.type === "dot" ? target.index : -1;
    if (idx === state.hovered) {
      if (idx >= 0) moveTooltip(e);
      else els.tooltip.classList.add("hidden");
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
    var target = pickTarget(e);
    if (target && target.type === "burst") {
      renderDrawerNote(target.item);
      return;
    }
    var idx = target && target.type === "dot" ? target.index : -1;
    selectParticipant(idx === state.selected ? -1 : idx);
  }

  function selectParticipant(idx) {
    state.selected = idx;
    if (idx < 0) {
      els.explain.classList.add("hidden");
      clearBurst();
      hideDrawer();
    } else {
      renderExplain(idx);
      els.explain.classList.remove("hidden");
      // The hub's data is memoized (bootstrapped at page load); the .then is
      // for the first-click race only.
      var pid = state.data.participants[idx];
      window.ClipgenOverview.ensureData().then(function () {
        if (state.selected !== idx) return; // stale selection
        var items = buildParticipantItems(pid);
        showBurst(idx, items);
        showDrawer(pid, items);
        requestRender();
      });
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

  // Labels are projected each render, then decluttered in screen space:
  // higher priority first (selected > participants > anchors > axis tips),
  // nearer-to-camera breaking ties; a label whose estimated rect overlaps an
  // already-placed one is hidden. Width is estimated from text length —
  // measuring real rects per frame (getBoundingClientRect) is a hot-loop
  // cost the render loop must not pay.
  var LABEL_EST_CHAR_PX = 6.5;
  var LABEL_EST_HEIGHT_PX = 15;

  function labelsOverlap(a, b) {
    return a.x1 < b.x2 && a.x2 > b.x1 && a.y1 < b.y2 && a.y2 > b.y1;
  }

  function updateLabels() {
    var w = els.canvasWrap.clientWidth;
    var h = els.canvasWrap.clientHeight;
    var v = new THREE.Vector3();
    var entries = [];
    var i;

    for (i = 0; i < three.dots.length; i++) {
      entries.push({
        el: three.labels[i],
        pos: three.dots[i].position,
        priority: i === state.selected ? 0 : 1,
      });
    }
    if (three.anchors) {
      for (i = 0; i < three.anchors.labels.length; i++) {
        if (!state.showAnchors) {
          three.anchors.labels[i].style.display = "none";
          continue;
        }
        entries.push({
          el: three.anchors.labels[i],
          pos: three.anchors.meshes[i].position,
          priority: 2,
        });
      }
    }
    for (i = 0; i < three.axisLabels.length; i++) {
      entries.push({
        el: three.axisLabels[i].el,
        pos: three.axisLabels[i].pos,
        priority: 3,
      });
    }

    for (i = 0; i < entries.length; i++) {
      var e = entries[i];
      v.copy(e.pos).project(three.camera);
      e.visible = v.z <= 1;
      e.x = (v.x + 1) / 2 * w;
      e.y = (1 - (v.y + 1) / 2) * h;
      e.depth = v.z;
    }

    var order = entries.slice().sort(function (a, b) {
      return a.priority - b.priority || a.depth - b.depth;
    });
    var placed = [];
    for (i = 0; i < order.length; i++) {
      var it = order[i];
      if (!it.visible) continue;
      // Rect mirrors the CSS transform: centered on x, sitting ~8px above y.
      var halfW = ((it.el.textContent || "").length * LABEL_EST_CHAR_PX + 8) / 2;
      var rect = {
        x1: it.x - halfW,
        x2: it.x + halfW,
        y1: it.y - LABEL_EST_HEIGHT_PX - 8,
        y2: it.y - 8,
      };
      for (var p = 0; p < placed.length; p++) {
        if (labelsOverlap(rect, placed[p])) { it.visible = false; break; }
      }
      if (it.visible) placed.push(rect);
    }

    for (i = 0; i < entries.length; i++) {
      var out = entries[i];
      if (!out.visible) {
        out.el.style.display = "none";
        continue;
      }
      out.el.style.display = "";
      out.el.style.left = out.x + "px";
      out.el.style.top = out.y + "px";
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

  // Scene-layer toggles. Layers stay built; toggling only flips visibility
  // (no recompute) via each def's apply(on).
  function layerToggleDefs() {
    return [
      {
        key: "showSimEdges",
        label: "Similarity links",
        hint: "Each participant links to its " + SIM_EDGE_K + " most-similar peers; brighter = more alike",
        apply: function (on) {
          if (three.simEdges) three.simEdges.mesh.visible = on;
        },
      },
      {
        key: "showAnchors",
        label: "Shared anchors",
        hint: "Observation categories and detectors as anchor nodes, placed amid the participants exhibiting them — shows why dots cluster",
        apply: setAnchorsVisible,
      },
      {
        key: "showAllMoments",
        label: "All moments",
        hint: "Every participant's timestamps as a point cloud around their dot, colored by source",
        apply: function (on) {
          if (three.moments) {
            three.moments.mesh.visible = on;
          } else if (on) {
            window.ClipgenOverview.ensureData().then(function () {
              rebuildMoments();
              requestRender();
            });
          }
        },
      },
    ];
  }

  function renderLayerToggles() {
    var host = els.layers;
    if (!host) return;
    var frag = document.createDocumentFragment();
    layerToggleDefs().forEach(function (def) {
      var label = document.createElement("label");
      label.className = "map-layer-row";
      label.title = def.hint;
      var box = document.createElement("input");
      box.type = "checkbox";
      box.checked = !!state[def.key];
      box.addEventListener("change", function () {
        state[def.key] = box.checked;
        def.apply(box.checked);
        requestRender();
      });
      var text = document.createElement("span");
      text.textContent = def.label;
      label.appendChild(box);
      label.appendChild(text);
      frag.appendChild(label);
    });
    host.innerHTML = "";
    host.appendChild(frag);
  }

  // Chips for individually muted features; click restores. Lives in the
  // sidebar because muted features drop out of the explain panel's ranking
  // (their weighted z is 0), so the panel can't offer the unmute itself.
  function renderMutedChips() {
    var section = document.getElementById("mapMutedSection");
    var host = document.getElementById("mapMuted");
    if (!section || !host) return;
    var keys = Object.keys(state.mutedFeatures);
    section.classList.toggle("hidden", keys.length === 0);
    host.innerHTML = "";
    if (!keys.length) return;
    var byKey = {};
    state.data.columns.forEach(function (col) { byKey[col.key] = col; });
    var frag = document.createDocumentFragment();
    keys.sort().forEach(function (key) {
      var chip = document.createElement("button");
      chip.type = "button";
      chip.className = "map-muted-chip";
      chip.title = "Restore this feature";
      chip.textContent = (byKey[key] ? byKey[key].label : key) + " ✕";
      chip.addEventListener("click", function () { setFeatureMuted(key, false); });
      frag.appendChild(chip);
    });
    host.appendChild(frag);
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
      var muteBtn = document.createElement("button");
      muteBtn.type = "button";
      muteBtn.className = "map-feature-mute";
      muteBtn.textContent = "✕";
      muteBtn.title = "Mute this feature — exclude it from similarity, " +
        "outlier scores, and the layout (restore from the sidebar)";
      muteBtn.addEventListener("click", (function (key) {
        return function () { setFeatureMuted(key, true); };
      })(col.key));
      label.appendChild(name);
      label.appendChild(zEl);
      label.appendChild(muteBtn);

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

    // Deep links: Transcripts/Screenspace pre-select this participant via the
    // #pid hash (clipgenHashParticipant in utils.js). Studio has no hash
    // support, so its link stays plain.
    var links = document.createElement("div");
    links.className = "map-explain-links";
    var hash = "#" + encodeURIComponent(pid);
    [
      ["Transcripts", "/transcripts/" + hash],
      ["Screenspace", "/screenspace/" + hash],
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
    els.layers = document.getElementById("mapLayers");
    els.drawer = document.getElementById("mapDrawer");
    els.drawerTitle = document.getElementById("mapDrawerTitle");
    els.drawerCount = document.getElementById("mapDrawerCount");
    els.drawerLanes = document.getElementById("mapDrawerLanes");
    els.drawerNote = document.getElementById("mapDrawerNote");
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

  // ---- Tab-satellite contract (called by the overview.js hub) ----

  var _mapBooted = false;

  function mapActivate() {
    if (!_mapBooted) {
      // First activation happens with #mapPanel already visible, so the
      // canvas measures its real size.
      _mapBooted = true;
      initDom();
      return;
    }
    // Re-activation (tab switch / hub Refresh): hub data may have changed,
    // so data-derived layers rebuild here rather than on every recompute.
    if (state.showAllMoments && three.renderer) rebuildMoments();
    onResize();
    requestRender();
  }

  window.ClipgenOverview.mapActivate = mapActivate;
  window.ClipgenOverview.mapDeactivate = function () {};
  window.ClipgenOverview.mapResize = onResize;
})();
