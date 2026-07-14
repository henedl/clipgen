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
    compareWith: -1,   // second participant of the compare pair (A = selected)
    colorBy: null,     // column key driving the dot choropleth (null = outlier ramp)
    layoutMode: "pca", // "pca" (similarity layout) | "manual" (direct axes)
    axisFeatures: { x: null, y: null, z: null, size: null }, // manual-mode column keys
    imputedAxis: [],   // [n] true where a manual-axis value was imputed to the mean
    worldScale: 1,     // scaleToWorld factor (projection units -> world units)
    showSimEdges: true,   // similarity-link layer toggle
    showAnchors: false,   // shared-anchor layer toggle (busier; default off)
    showAllMoments: false, // every participant's items as a point cloud
    showTrajectories: false, // per-participant session paths (bolder; default off)
    showClusterHulls: false, // k-means ellipsoid shells (default off)
    clusterK: 0,          // manual cluster count; 0 = auto (best silhouette)
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
    compare: null,  // {a, b, line, beads} pairwise diff arc
    trajectories: null, // {lines, comets, cometGeo} session-trajectory paths
    clusterHulls: null, // {geo, meshes, wires, labels, clusters} k-means shells
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
  var renderRaf = 0;
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

  // Pairwise weighted-z difference between two participants — the same
  // distance form the explain panel's contribution loop uses, but between two
  // dots instead of dot-vs-centroid. Only columns where BOTH sides have data
  // count; a group where either side is fully null reports count 0 so the
  // panel can say "no shared data" instead of a misleading zero distance.
  function computePairDiff(a, b) {
    var data = state.data;
    var groups = {};
    var ranked = [];
    var total = 0;
    for (var j = 0; j < data.columns.length; j++) {
      var col = data.columns[j];
      if (!groups[col.group]) groups[col.group] = { dist: 0, count: 0 };
      var zA = state.zRaw[a][j];
      var zB = state.zRaw[b][j];
      if (zA == null || zB == null) continue;
      var dz = (zA - zB) * columnWeight(col);
      groups[col.group].dist += dz * dz;
      groups[col.group].count++;
      total += dz * dz;
      if (Math.abs(dz) > 1e-9) ranked.push({ j: j, dz: dz, zA: zA, zB: zB });
    }
    Object.keys(groups).forEach(function (g) {
      groups[g].dist = Math.sqrt(groups[g].dist);
    });
    ranked.sort(function (x, y) { return Math.abs(y.dz) - Math.abs(x.dz); });
    return { groups: groups, total: Math.sqrt(total), ranked: ranked };
  }

  function scaleToWorld(coords) {
    var maxR = 0;
    var i, c;
    for (i = 0; i < coords.length; i++) {
      for (c = 0; c < 3; c++) maxR = Math.max(maxR, Math.abs(coords[i][c]));
    }
    if (maxR <= 0) { state.worldScale = 1; return coords; }
    var s = WORLD_RADIUS / maxR;
    state.worldScale = s; // kept so other layers can project into dot space
    var out = [];
    for (i = 0; i < coords.length; i++) {
      out.push([coords[i][0] * s, coords[i][1] * s, coords[i][2] * s]);
    }
    return out;
  }

  function axisFeatureIndex(dim) {
    var key = state.axisFeatures[dim];
    if (!key || !state.data) return -1;
    var cols = state.data.columns;
    for (var j = 0; j < cols.length; j++) if (cols[j].key === key) return j;
    return -1;
  }

  // Layout coordinates for the current mode. PCA mode projects the weighted
  // matrix; manual mode maps the three hand-picked features straight onto the
  // axes from state.zRaw — clamped z, UNWEIGHTED, because the group weights
  // are a similarity lens and must not warp axes the user chose explicitly.
  // Missing values impute to z = 0 (the cohort mean IS the axis midpoint in
  // z-space); state.imputedAxis marks those dots so styleDots can fade them.
  // PCA components are computed either way: the outlier list, sim-edges, and
  // anchors all live off the full weighted feature space regardless of mode.
  function computeLayoutCoords() {
    var columns = state.data.columns;
    var n = state.zRaw.length;
    var i;
    state.imputedAxis = [];
    for (i = 0; i < n; i++) state.imputedAxis.push(false);
    var pca = pcaProject(state.weighted, columns.length);
    state.components = pca.components;
    state.variance = pca.variance;
    if (state.layoutMode !== "manual") return scaleToWorld(pca.coords);

    var idx = [axisFeatureIndex("x"), axisFeatureIndex("y"), axisFeatureIndex("z")];
    var coords = [];
    for (i = 0; i < n; i++) {
      var row = [0, 0, 0];
      for (var c = 0; c < 3; c++) {
        if (idx[c] < 0) continue;
        var z = state.zRaw[i][idx[c]];
        if (z == null) { state.imputedAxis[i] = true; continue; }
        row[c] = z;
      }
      coords.push(row);
    }
    return scaleToWorld(coords);
  }

  function recompute() {
    var data = state.data;
    var columns = data.columns;
    state.zRaw = computeZ(data.matrix, state.stats);
    state.weighted = applyWeights(state.zRaw, columns);
    var newCoords = computeLayoutCoords();
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
    updateAxisTipText();
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
    rebuildCompare();
    rebuildTrajectories(); // no-op (dispose only) while the toggle is off
    rebuildClusterHulls(); // ditto; reclusters on weight/mute changes
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
      .replace(" share of timestamps", "")
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

  // ---- Pairwise compare arc --------------------------------------------------
  //
  // Shift-click (or arm the Compare button and click) a second dot while one
  // is selected: a raised arc connects the pair, with one "difference bead"
  // per feature group along it — bead size = that group's weighted-z distance
  // between the two. Beads are deliberately NOT raycast targets (they sit
  // between dots and would fight selection); numbers live in the panel.

  var COMPARE_ARC_SEGMENTS = 32;
  var _comparePickArmed = false;

  // Feature-group colors reuse the page-wide stream tokens (observations ARE
  // the sheet stream, screenspace/transcript likewise); session_shape has no
  // stream of its own so it takes the neutral accent.
  function groupColor(groupKey) {
    if (groupKey === "observations") return three.colors.streamSheet;
    if (groupKey === "screenspace") return three.colors.streamScreenspace;
    if (groupKey === "transcript") return three.colors.streamTranscript;
    return three.colors.base;
  }

  function groupCssColor(groupKey) {
    if (groupKey === "observations") return "var(--stream-sheet)";
    if (groupKey === "screenspace") return "var(--stream-screenspace)";
    if (groupKey === "transcript") return "var(--stream-transcript)";
    return "var(--accent)";
  }

  function disposeCompare() {
    if (!three.compare) return;
    disposeLayer({ mesh: three.compare.line });
    for (var i = 0; i < three.compare.beads.length; i++) {
      three.scene.remove(three.compare.beads[i].mesh);
      three.compare.beads[i].mesh.geometry.dispose();
      three.compare.beads[i].mesh.material.dispose();
    }
    three.compare = null;
  }

  // Arc between the pair's CURRENT dot positions: control point at the
  // midpoint, pushed radially away from the origin so the arc rises above
  // the cloud instead of cutting through it.
  function compareCurve(a, b) {
    var pa = three.dots[a].position;
    var pb = three.dots[b].position;
    var mid = new THREE.Vector3().addVectors(pa, pb).multiplyScalar(0.5);
    var chord = pa.distanceTo(pb);
    var lift = mid.length() > 1e-6
      ? mid.clone().normalize() : new THREE.Vector3(0, 1, 0);
    mid.add(lift.multiplyScalar(Math.max(chord * 0.35, 1.5)));
    return new THREE.QuadraticBezierCurve3(pa.clone(), mid, pb.clone());
  }

  function rebuildCompare() {
    disposeCompare();
    if (state.selected < 0 || state.compareWith < 0) return;
    var a = state.selected;
    var b = state.compareWith;
    var diff = computePairDiff(a, b);

    var geo = new THREE.BufferGeometry();
    geo.setAttribute("position", new THREE.BufferAttribute(
      new Float32Array((COMPARE_ARC_SEGMENTS + 1) * 3), 3));
    var line = new THREE.Line(geo, new THREE.LineBasicMaterial({
      color: three.colors.base, transparent: true, opacity: 0.8,
    }));
    three.scene.add(line);

    var groups = state.data.groups;
    var maxDist = 0;
    var g;
    for (g = 0; g < groups.length; g++) {
      var entry = diff.groups[groups[g].key];
      if (entry) maxDist = Math.max(maxDist, entry.dist);
    }
    var beadGeo = new THREE.SphereGeometry(1, 16, 12);
    var beads = [];
    for (g = 0; g < groups.length; g++) {
      var e = diff.groups[groups[g].key] || { dist: 0, count: 0 };
      var mesh = new THREE.Mesh(beadGeo, new THREE.MeshBasicMaterial({
        color: groupColor(groups[g].key),
        // A no-shared-data bead stays visible but ghosted.
        transparent: e.count === 0, opacity: e.count === 0 ? 0.25 : 1,
      }));
      mesh.scale.setScalar(0.15 + (maxDist > 0 ? 0.45 * (e.dist / maxDist) : 0));
      three.scene.add(mesh);
      beads.push({ mesh: mesh, t: (g + 1) / (groups.length + 1) });
    }
    three.compare = { a: a, b: b, line: line, beads: beads };
    syncComparePositions();
  }

  function syncComparePositions() {
    if (!three.compare) return;
    var curve = compareCurve(three.compare.a, three.compare.b);
    var arr = three.compare.line.geometry.attributes.position.array;
    for (var s = 0; s <= COMPARE_ARC_SEGMENTS; s++) {
      var p = curve.getPoint(s / COMPARE_ARC_SEGMENTS);
      arr[s * 3] = p.x;
      arr[s * 3 + 1] = p.y;
      arr[s * 3 + 2] = p.z;
    }
    three.compare.line.geometry.attributes.position.needsUpdate = true;
    for (var i = 0; i < three.compare.beads.length; i++) {
      three.compare.beads[i].mesh.position.copy(
        curve.getPoint(three.compare.beads[i].t));
    }
  }

  function syncCompareBtn() {
    var btn = document.getElementById("mapCompareBtn");
    if (!btn) return;
    btn.classList.toggle("is-armed", _comparePickArmed || state.compareWith >= 0);
    btn.textContent = state.compareWith >= 0 ? "End compare"
      : (_comparePickArmed ? "Click a dot…" : "Compare…");
  }

  function setCompare(idx) {
    state.compareWith = idx;
    _comparePickArmed = false;
    rebuildCompare();
    syncCompareBtn();
    if (state.selected >= 0) renderExplain(state.selected);
    requestRender();
  }

  // ---- Cluster hulls (k-means) ------------------------------------------------
  //
  // Deterministic k-means over the FULL weighted feature space (the same
  // matrix the layout projects from, never the projection): farthest-point
  // init seeded from the top outlier — no randomness, so the same data +
  // weights always produce identical clusters — then Lloyd iterations to
  // convergence. Auto-k picks the best mean silhouette over k = 2..min(5,
  // floor(n/2)). Each cluster renders as a translucent ellipsoid around its
  // member dots (recomputed from live positions, so hulls ride the lerp)
  // with an auto-label naming its top distinguishing features.

  var KMEANS_MAX_ITER = 50;
  var CLUSTER_MIN_PARTICIPANTS = 4;

  function _clusterDistSq(vec, center) {
    var s = 0;
    for (var j = 0; j < vec.length; j++) {
      var d = vec[j] - center[j];
      s += d * d;
    }
    return s;
  }

  function runKmeans(k) {
    var W = state.weighted;
    var n = W.length;
    var i, c, j;
    // Farthest-point init: seed = the most unusual participant, then
    // repeatedly the point farthest from every chosen center.
    var centers = [W[state.order[0]].slice()];
    while (centers.length < k) {
      var far = 0;
      var farDist = -1;
      for (i = 0; i < n; i++) {
        var nearest = Infinity;
        for (c = 0; c < centers.length; c++) {
          nearest = Math.min(nearest, _clusterDistSq(W[i], centers[c]));
        }
        if (nearest > farDist) { farDist = nearest; far = i; }
      }
      centers.push(W[far].slice());
    }
    var assign = new Array(n);
    for (var it = 0; it < KMEANS_MAX_ITER; it++) {
      var changed = false;
      for (i = 0; i < n; i++) {
        var bc = 0;
        var bd = Infinity;
        for (c = 0; c < k; c++) {
          var dsq = _clusterDistSq(W[i], centers[c]);
          if (dsq < bd) { bd = dsq; bc = c; }
        }
        if (assign[i] !== bc) { assign[i] = bc; changed = true; }
      }
      if (!changed) break;
      for (c = 0; c < k; c++) {
        var sum = null;
        var count = 0;
        for (i = 0; i < n; i++) {
          if (assign[i] !== c) continue;
          if (!sum) { sum = W[i].slice(); count = 1; continue; }
          for (j = 0; j < sum.length; j++) sum[j] += W[i][j];
          count++;
        }
        if (sum) {
          for (j = 0; j < sum.length; j++) sum[j] /= count;
          centers[c] = sum; // an empty cluster keeps its old center
        }
      }
    }
    return assign;
  }

  function meanSilhouette(assign, k) {
    var W = state.weighted;
    var n = W.length;
    var total = 0;
    var counted = 0;
    for (var i = 0; i < n; i++) {
      var sums = new Array(k);
      var counts = new Array(k);
      var c;
      for (c = 0; c < k; c++) { sums[c] = 0; counts[c] = 0; }
      for (var j = 0; j < n; j++) {
        if (j === i) continue;
        var dist = Math.sqrt(_clusterDistSq(W[i], W[j]));
        sums[assign[j]] += dist;
        counts[assign[j]]++;
      }
      if (!counts[assign[i]]) continue; // singleton: no silhouette
      var a = sums[assign[i]] / counts[assign[i]];
      var b = Infinity;
      for (c = 0; c < k; c++) {
        if (c === assign[i] || !counts[c]) continue;
        b = Math.min(b, sums[c] / counts[c]);
      }
      if (b === Infinity) continue;
      total += (b - a) / Math.max(a, b, 1e-12);
      counted++;
    }
    return counted ? total / counted : -1;
  }

  function computeKmeans() {
    var n = state.weighted ? state.weighted.length : 0;
    if (n < CLUSTER_MIN_PARTICIPANTS) return null;
    var k = state.clusterK;
    if (k >= 2) {
      k = Math.max(2, Math.min(k, Math.floor(n / 2)));
      return { assign: runKmeans(k), k: k };
    }
    var kMax = Math.min(5, Math.floor(n / 2));
    var best = null;
    for (var kk = 2; kk <= kMax; kk++) {
      var assign = runKmeans(kk);
      var sil = meanSilhouette(assign, kk);
      if (!best || sil > best.sil + 1e-9) best = { assign: assign, k: kk, sil: sil };
    }
    return best;
  }

  // Auto-label: top features by |cluster mean z| (the cohort mean z is 0 by
  // construction, so a cluster's mean z IS its deviation), muted excluded.
  function clusterLabel(members) {
    var cols = state.data.columns;
    var diffs = [];
    for (var j = 0; j < cols.length; j++) {
      if (state.mutedFeatures[cols[j].key]) continue;
      var sum = 0;
      var count = 0;
      for (var m = 0; m < members.length; m++) {
        var z = state.zRaw[members[m]][j];
        if (z == null) continue;
        sum += z;
        count++;
      }
      if (count) diffs.push({ j: j, mean: sum / count });
    }
    diffs.sort(function (a, b) { return Math.abs(b.mean) - Math.abs(a.mean); });
    var parts = [];
    for (var t = 0; t < Math.min(2, diffs.length); t++) {
      if (Math.abs(diffs[t].mean) < 0.3) break; // near the mean: not distinguishing
      parts.push((diffs[t].mean > 0 ? "high " : "low ") + cols[diffs[t].j].label);
    }
    return parts.length ? parts.join(", ") : "near cohort mean";
  }

  function renderClusterNotice(result) {
    var info = document.getElementById("mapClusterInfo");
    if (!info) return;
    if (!state.showClusterHulls) { info.textContent = ""; return; }
    if (!result) {
      info.textContent = "Needs " + CLUSTER_MIN_PARTICIPANTS + "+ participants";
      return;
    }
    info.textContent = result.k + " clusters" + (state.clusterK >= 2 ? "" : " (auto)");
  }

  function disposeClusterHulls() {
    if (!three.clusterHulls) return;
    var h = three.clusterHulls;
    var i;
    for (i = 0; i < h.meshes.length; i++) {
      three.scene.remove(h.meshes[i]);
      h.meshes[i].material.dispose();
    }
    for (i = 0; i < h.wires.length; i++) {
      three.scene.remove(h.wires[i]);
      h.wires[i].material.dispose();
    }
    if (h.geo) h.geo.dispose();
    for (i = 0; i < h.labels.length; i++) {
      if (h.labels[i].parentNode) h.labels[i].parentNode.removeChild(h.labels[i]);
    }
    three.clusterHulls = null;
  }

  function rebuildClusterHulls() {
    disposeClusterHulls();
    if (!state.showClusterHulls) { renderClusterNotice(null); return; }
    var result = computeKmeans();
    renderClusterNotice(result);
    if (!result) return;

    var clusters = [];
    var c, i;
    for (c = 0; c < result.k; c++) clusters.push([]);
    for (i = 0; i < result.assign.length; i++) {
      clusters[result.assign[i]].push(i);
    }
    clusters = clusters.filter(function (m) { return m.length > 0; });
    // Size-desc ordering keeps color assignment stable and meaningful.
    clusters.sort(function (a, b) { return b.length - a.length || a[0] - b[0]; });

    // Cluster colors cycle the group/stream palette (k <= 5, so at most one
    // reuse); hulls are translucent so a repeat still reads.
    var hullColors = [
      three.colors.streamSheet,
      three.colors.streamScreenspace,
      three.colors.streamTranscript,
      three.colors.base,
      three.colors.friction,
    ];

    var geo = new THREE.SphereGeometry(1, 24, 16);
    var meshes = [];
    var wires = [];
    var labels = [];
    var labelFrag = document.createDocumentFragment();
    for (c = 0; c < clusters.length; c++) {
      var color = hullColors[c % hullColors.length];
      var mesh = new THREE.Mesh(geo, new THREE.MeshBasicMaterial({
        color: color, transparent: true, opacity: 0.08, depthWrite: false,
      }));
      var wire = new THREE.Mesh(geo, new THREE.MeshBasicMaterial({
        color: color, transparent: true, opacity: 0.18,
        wireframe: true, depthWrite: false,
      }));
      three.scene.add(mesh);
      three.scene.add(wire);
      meshes.push(mesh);
      wires.push(wire);

      var label = document.createElement("div");
      label.className = "map-label map-cluster-label";
      label.textContent = "C" + (c + 1) + " · " + clusterLabel(clusters[c]) +
        " (" + clusters[c].length + ")";
      labelFrag.appendChild(label);
      labels.push(label);
    }
    els.labels.appendChild(labelFrag);
    three.clusterHulls = {
      geo: geo, meshes: meshes, wires: wires, labels: labels, clusters: clusters,
    };
    syncClusterHullPositions();
  }

  // Ellipsoid = member-dot centroid, per-axis 1.6 x std (floored so tight
  // clusters stay visible) — from LIVE dot positions, so hulls ride the lerp.
  function syncClusterHullPositions() {
    if (!three.clusterHulls) return;
    var h = three.clusterHulls;
    for (var c = 0; c < h.clusters.length; c++) {
      var members = h.clusters[c];
      var cx = 0, cy = 0, cz = 0;
      var m, p;
      for (m = 0; m < members.length; m++) {
        p = three.dots[members[m]].position;
        cx += p.x;
        cy += p.y;
        cz += p.z;
      }
      cx /= members.length;
      cy /= members.length;
      cz /= members.length;
      var vx = 0, vy = 0, vz = 0;
      for (m = 0; m < members.length; m++) {
        p = three.dots[members[m]].position;
        vx += (p.x - cx) * (p.x - cx);
        vy += (p.y - cy) * (p.y - cy);
        vz += (p.z - cz) * (p.z - cz);
      }
      var sx = Math.max(1.6 * Math.sqrt(vx / members.length), 0.9);
      var sy = Math.max(1.6 * Math.sqrt(vy / members.length), 0.9);
      var sz = Math.max(1.6 * Math.sqrt(vz / members.length), 0.9);
      h.meshes[c].position.set(cx, cy, cz);
      h.meshes[c].scale.set(sx, sy, sz);
      h.wires[c].position.set(cx, cy, cz);
      h.wires[c].scale.set(sx, sy, sz);
    }
  }

  // ---- Session trajectories ---------------------------------------------------
  //
  // Each participant's session split into K temporal windows (server-built,
  // payload.windows) becomes a polyline through the SAME space the dots live
  // in: every window vector is z-scored with the whole-session stats,
  // weighted, and projected onto the whole-session PCA basis
  // (state.components x state.worldScale) — the frame the researcher tuned
  // with the sliders. A pooled-window basis would either move the dots or
  // put paths in a different space than the dots; both would be incoherent.
  // In manual mode the window vectors map through the same hand-picked axes
  // instead. The path's final vertex is appended AT the dot (the whole
  // session), so every path visibly terminates at the participant's
  // canonical position and all other layers stay anchored to one spot.

  // [participant] -> [window] -> [x,y,z] or null (window fully null).
  function computeWindowCoords() {
    var data = state.data;
    if (!data || !data.windows || !data.windows.matrices) return null;
    var mats = data.windows.matrices;
    var columns = data.columns;
    var manual = state.layoutMode === "manual";
    var axIdx = manual
      ? [axisFeatureIndex("x"), axisFeatureIndex("y"), axisFeatureIndex("z")]
      : null;
    var out = [];
    var i, w, j, c;
    for (i = 0; i < data.participants.length; i++) out.push([]);
    for (w = 0; w < mats.length; w++) {
      var z = computeZ(mats[w], state.stats);
      for (i = 0; i < z.length; i++) {
        var row = z[i];
        var allNull = true;
        for (j = 0; j < row.length; j++) {
          if (row[j] != null) { allNull = false; break; }
        }
        if (allNull) { out[i].push(null); continue; }
        var v = [0, 0, 0];
        if (manual) {
          for (c = 0; c < 3; c++) {
            if (axIdx[c] >= 0 && row[axIdx[c]] != null) v[c] = row[axIdx[c]];
          }
        } else {
          for (j = 0; j < columns.length; j++) {
            var wz = row[j] == null ? 0 : row[j] * columnWeight(columns[j]);
            if (wz === 0) continue;
            v[0] += wz * state.components[0][j];
            v[1] += wz * state.components[1][j];
            v[2] += wz * state.components[2][j];
          }
        }
        out[i].push([
          v[0] * state.worldScale,
          v[1] * state.worldScale,
          v[2] * state.worldScale,
        ]);
      }
    }
    return out;
  }

  function disposeTrajectories() {
    if (!three.trajectories) return;
    var t = three.trajectories;
    var i;
    for (i = 0; i < t.lines.length; i++) {
      three.scene.remove(t.lines[i].mesh);
      t.lines[i].mesh.geometry.dispose();
      t.lines[i].mesh.material.dispose();
    }
    for (i = 0; i < t.comets.length; i++) {
      three.scene.remove(t.comets[i].mesh);
      t.comets[i].mesh.material.dispose();
    }
    if (t.cometGeo) t.cometGeo.dispose();
    three.trajectories = null;
  }

  function rebuildTrajectories() {
    disposeTrajectories();
    if (!state.showTrajectories) return;
    var coords = computeWindowCoords();
    if (!coords) return;
    var lines = [];
    var comets = [];
    var cometGeo = new THREE.SphereGeometry(0.18, 12, 8);
    var base = new THREE.Color();
    for (var i = 0; i < coords.length; i++) {
      var verts = [];
      for (var w = 0; w < coords[i].length; w++) {
        if (coords[i][w]) verts.push(coords[i][w]);
      }
      if (!verts.length) continue; // no windowable source -> no path
      var vCount = verts.length + 1; // +1: the dot itself, synced per frame
      var positions = new Float32Array(vCount * 3);
      var colors = new Float32Array(vCount * 3);
      dotBaseColor(i, base);
      for (var v = 0; v < vCount; v++) {
        // Older vertices fade toward the background (LineBasicMaterial has
        // no per-vertex alpha in r147 — same trick as the sim edges).
        var mix = 1 - (v + 1) / vCount;
        var cc = base.clone().lerp(three.colors.bg, mix * 0.75);
        colors[v * 3] = cc.r;
        colors[v * 3 + 1] = cc.g;
        colors[v * 3 + 2] = cc.b;
        if (v < verts.length) {
          positions[v * 3] = verts[v][0];
          positions[v * 3 + 1] = verts[v][1];
          positions[v * 3 + 2] = verts[v][2];
        }
      }
      var geo = new THREE.BufferGeometry();
      geo.setAttribute("position", new THREE.BufferAttribute(positions, 3));
      geo.setAttribute("color", new THREE.BufferAttribute(colors, 3));
      var mesh = new THREE.Line(geo, new THREE.LineBasicMaterial({
        vertexColors: true, transparent: true, opacity: 0.9,
      }));
      three.scene.add(mesh);
      lines.push({ mesh: mesh, owner: i, count: vCount });

      // One replay comet per path, positioned by syncComets during replay.
      var comet = new THREE.Mesh(
        cometGeo, new THREE.MeshBasicMaterial({ color: base.clone() })
      );
      comet.visible = false;
      three.scene.add(comet);
      comets.push({ mesh: comet, owner: i });
    }
    three.trajectories = { lines: lines, comets: comets, cometGeo: cometGeo };
    syncTrajectoryTails();
  }

  // The appended last vertex tracks the (possibly lerping) dot each frame.
  function syncTrajectoryTails() {
    if (!three.trajectories) return;
    var lines = three.trajectories.lines;
    for (var i = 0; i < lines.length; i++) {
      var arr = lines[i].mesh.geometry.attributes.position.array;
      var p = three.dots[lines[i].owner].position;
      var last = (lines[i].count - 1) * 3;
      arr[last] = p.x;
      arr[last + 1] = p.y;
      arr[last + 2] = p.z;
      lines[i].mesh.geometry.attributes.position.needsUpdate = true;
    }
  }

  // Replay comet: a bright marker interpolated along each path at the
  // playhead's normalized session time.
  function syncComets(t) {
    if (!three.trajectories) return;
    var comets = three.trajectories.comets;
    var lines = three.trajectories.lines;
    var i;
    if (!state.showTrajectories || t <= 0) {
      for (i = 0; i < comets.length; i++) comets[i].mesh.visible = false;
      return;
    }
    if (!_replayWhite) _replayWhite = new THREE.Color("#ffffff");
    for (i = 0; i < lines.length; i++) {
      var arr = lines[i].mesh.geometry.attributes.position.array;
      var segs = lines[i].count - 1;
      var f = Math.min(t, 1) * segs;
      var v0 = Math.min(Math.floor(f), segs - 1);
      var frac = f - v0;
      var comet = comets[i].mesh;
      comet.position.set(
        arr[v0 * 3] + (arr[(v0 + 1) * 3] - arr[v0 * 3]) * frac,
        arr[v0 * 3 + 1] + (arr[(v0 + 1) * 3 + 1] - arr[v0 * 3 + 1]) * frac,
        arr[v0 * 3 + 2] + (arr[(v0 + 1) * 3 + 2] - arr[v0 * 3 + 2]) * frac
      );
      dotBaseColor(lines[i].owner, comet.material.color);
      comet.material.color.lerp(_replayWhite, 0.5);
      comet.visible = true;
    }
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
    var i;
    for (i = 0; i < three.dots.length; i++) {
      var glow = replayIntensity(i, t);
      var c = three.dots[i].material.color;
      dotBaseColor(i, c); // idle look from the single authority
      if (glow > 0) c.lerp(_replayWhite, glow * 0.7);
      three.dots[i].scale.setScalar(dotBaseScale(i) * (1 + glow * 0.4));
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
    syncComets(t);
    if (!lerp.active) positionDots(state.coords); // re-sync halo scale
    requestRender();
  }

  // Back to the idle look: base dot colors, no pulse, full-bright cloud.
  function clearReplayGlow() {
    styleDots();
    syncComets(0);
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
  //
  // The feature matrix (api/data) is tied to the hub's dataVersion, the same
  // staleness contract the Convergence/Metadata tabs use: _matrixVersion
  // records the hub version the current matrix was fetched at, and a tab
  // re-activation with a newer version re-fetches and rebuilds — so the
  // header Refresh updates the layout, not just the drill-down items.

  var _matrixVersion = -1;

  function hubDataVersion() {
    var OV = window.ClipgenOverview;
    return OV && OV.state ? OV.state.dataVersion : 0;
  }

  // True only while the Map tab is the active one. The orbit/hover listeners are
  // window-level (so a drag can leave the canvas), so gate the per-move raycast
  // on this to avoid work while Convergence/Metadata is showing.
  function mapIsActive() {
    var OV = window.ClipgenOverview;
    return !!(OV && OV.state && OV.state.activeTab === "map");
  }

  function loadData() {
    // Await the hub bootstrap alongside our own fetch so _matrixVersion is
    // captured after the version settles (ensureData never rejects).
    Promise.all([apiGet("api/data"), window.ClipgenOverview.ensureData()])
      .then(function (results) {
        var data = results[0];
        if (!data || !data.ok) throw new Error((data && data.error) || "bad payload");
        _matrixVersion = hubDataVersion();
        init(data);
      })
      .catch(function (err) {
        showNotice("Could not load study data: " + err.message);
      });
  }

  function reloadData() {
    apiGet("api/data")
      .then(function (data) {
        if (!data || !data.ok) throw new Error((data && data.error) || "bad payload");
        _matrixVersion = hubDataVersion();
        applyData(data);
      })
      .catch(function (err) {
        showNotice("Could not refresh study data: " + err.message);
      });
  }

  // Ingest a fresh matrix into a live scene: rebuild the participant objects
  // (the cohort can grow/shrink), keep the user's lens (weights, muted
  // features, layer toggles), and carry the selection over by participant id.
  function applyData(data) {
    if (window.clipgenApplyConfig) window.clipgenApplyConfig(data.config);

    if (!data.participants.length) {
      state.data = data;
      els.empty.classList.remove("hidden");
      if (three.renderer) {
        state.scores = [];
        state.order = [];
        clearBurst();
        disposeLayer(three.simEdges);
        three.simEdges = null;
        disposeAnchors();
        disposeLayer(three.moments);
        three.moments = null;
        disposeTrajectories();
        disposeClusterHulls();
        disposeDots();
        selectParticipant(-1);
      }
      return;
    }
    els.empty.classList.add("hidden");

    if (!three.renderer) {
      // The study appeared after an empty boot — run the full first-boot path.
      init(data);
      return;
    }

    var prevSelected = state.selected >= 0 && state.data
      ? state.data.participants[state.selected] : null;

    state.data = data;
    state.groupSizes = {};
    var i;
    for (i = 0; i < data.columns.length; i++) {
      var g = data.columns[i].group;
      state.groupSizes[g] = (state.groupSizes[g] || 0) + 1;
    }
    for (i = 0; i < data.groups.length; i++) {
      if (state.weights[data.groups[i].key] == null) {
        state.weights[data.groups[i].key] = 1;
      }
    }
    var knownKeys = {};
    for (i = 0; i < data.columns.length; i++) knownKeys[data.columns[i].key] = true;
    Object.keys(state.mutedFeatures).forEach(function (key) {
      if (!knownKeys[key]) delete state.mutedFeatures[key];
    });
    if (state.colorBy && !knownKeys[state.colorBy]) state.colorBy = null;
    ["x", "y", "z", "size"].forEach(function (d) {
      if (state.axisFeatures[d] && !knownKeys[state.axisFeatures[d]]) {
        state.axisFeatures[d] = null;
      }
    });
    if (state.layoutMode === "manual") ensureAxisDefaults();

    if (data.participants.length < 3) {
      showNotice("Only " + data.participants.length + " participant" +
        (data.participants.length === 1 ? "" : "s") +
        " so far — 3+ make the similarity layout meaningful.");
    } else {
      els.notice.classList.add("hidden");
    }

    state.stats = computeStats(data.matrix, data.columns.length);
    state.coords = null; // cohort indices changed; don't lerp across datasets
    replay.norms = null; // items changed; replay densities rebuild lazily
    state.compareWith = -1; // compare pair is index-based; indices just moved
    _comparePickArmed = false;
    syncCompareBtn();

    clearBurst();
    disposeLayer(three.moments);
    three.moments = null;
    disposeTrajectories(); // rebuilt by recompute's rebuildLayers if toggled on
    disposeClusterHulls();
    disposeDots();
    buildDots();

    state.selected = prevSelected ? data.participants.indexOf(prevSelected) : -1;
    renderWeights();
    renderMutedChips();
    renderColorByChip();
    recompute();
    if (state.selected >= 0) {
      selectParticipant(state.selected);
    } else {
      els.explain.classList.add("hidden");
      hideDrawer();
    }
    if (state.showAllMoments) rebuildMoments();
    requestRender();
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
    els.empty.classList.add("hidden");
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
    renderLayoutControls();
    renderAxisPickers();
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
    // Resize is driven by the hub's tab-aware forwarder (OV.mapResize) and by
    // mapActivate()'s onResize() on (re)activation; a direct window listener here
    // would also fire on inactive tabs and double-fire on the Map.
    applyCamera();
  }

  function disposeDots() {
    var i;
    for (i = 0; i < three.dots.length; i++) {
      three.scene.remove(three.dots[i]);
      three.dots[i].material.dispose();
    }
    for (i = 0; i < three.halos.length; i++) {
      three.scene.remove(three.halos[i]);
      three.halos[i].material.dispose();
    }
    for (i = 0; i < three.labels.length; i++) {
      if (three.labels[i].parentNode) {
        three.labels[i].parentNode.removeChild(three.labels[i]);
      }
    }
    if (three._dotGeo) three._dotGeo.dispose();
    if (three._haloGeo) three._haloGeo.dispose();
    three._dotGeo = null;
    three._haloGeo = null;
    three.dots = [];
    three.halos = [];
    three.labels = [];
  }

  function buildDots() {
    var participants = state.data.participants;
    var geo = new THREE.SphereGeometry(0.45, 24, 16);
    var haloGeo = new THREE.SphereGeometry(0.8, 24, 16);
    three._dotGeo = geo;
    three._haloGeo = haloGeo;
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

  // ---- Dot styling (single authority) ---------------------------------------
  //
  // dotBaseColor/dotBaseScale are the ONE place a dot's idle look is computed;
  // styleDots() and applyReplayGlow() both build on them so the choropleth,
  // the outlier ramp, selection, and the replay glow compose instead of each
  // maintaining a divergent copy of the formula.

  function colorByIndex() {
    if (!state.colorBy || !state.data) return -1;
    var cols = state.data.columns;
    for (var j = 0; j < cols.length; j++) {
      if (cols[j].key === state.colorBy) return j;
    }
    return -1;
  }

  function maxOutlierScore() {
    var maxScore = 0;
    for (var i = 0; i < state.scores.length; i++) {
      maxScore = Math.max(maxScore, state.scores[i]);
    }
    return maxScore;
  }

  // Writes dot i's idle color into `out`: with a "color by" feature active,
  // that column's z ramp (a null cell fades toward the background — "no
  // data", not "low"); otherwise the outlier-score ramp.
  function dotBaseColor(i, out) {
    var cj = colorByIndex();
    if (cj >= 0) {
      var z = state.zRaw ? state.zRaw[i][cj] : null;
      if (z == null) {
        return out.copy(three.colors.base).lerp(three.colors.bg, 0.65);
      }
      return out.copy(three.colors.base)
        .lerp(three.colors.hot, (z + Z_CLAMP) / (2 * Z_CLAMP));
    }
    var maxScore = maxOutlierScore();
    var t = maxScore > 0 ? state.scores[i] / maxScore : 0;
    return out.copy(three.colors.base).lerp(three.colors.hot, t);
  }

  function dotBaseScale(i) {
    var sizeFactor;
    var sj = state.layoutMode === "manual" ? axisFeatureIndex("size") : -1;
    if (sj >= 0) {
      // Manual size dimension: z ramp; null sits at the midpoint so missing
      // data doesn't read as "small".
      var z = state.zRaw ? state.zRaw[i][sj] : null;
      var tz = z == null ? 0.5 : (z + Z_CLAMP) / (2 * Z_CLAMP);
      sizeFactor = 0.6 + 0.9 * tz;
    } else {
      var maxScore = maxOutlierScore();
      var t = maxScore > 0 ? state.scores[i] / maxScore : 0;
      sizeFactor = 0.85 + t * 0.5;
    }
    return (i === state.selected ? 1.4 : 1) * sizeFactor;
  }

  function styleDots() {
    for (var i = 0; i < three.dots.length; i++) {
      dotBaseColor(i, three.dots[i].material.color);
      // Fade dots whose manual-axis value was imputed to the cohort mean.
      // transparent is set per-dot, only where needed — cohort-wide
      // transparency would invite sorting artifacts.
      var imputed = state.layoutMode === "manual" && !!state.imputedAxis[i];
      three.dots[i].material.transparent = imputed;
      three.dots[i].material.opacity = imputed ? 0.35 : 1;
      three.dots[i].scale.setScalar(dotBaseScale(i));
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
    syncComparePositions();
    syncTrajectoryTails();
    syncClusterHullPositions();
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
      } else if (mapIsActive()) {
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
    var tip = pid + " — unusualness " +
      state.scores[idx].toFixed(2) + " (#" + rank + ")";
    var cj = colorByIndex();
    if (cj >= 0) {
      tip += " · " + state.data.columns[cj].label + ": " +
        fmtNum(state.data.matrix[idx][cj]);
    }
    els.tooltip.textContent = tip;
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
    // Shift-click (or an armed Compare button) on a second dot diffs the
    // pair instead of moving the selection; repeating it ends the compare.
    if (idx >= 0 && state.selected >= 0 && idx !== state.selected &&
        (e.shiftKey || _comparePickArmed)) {
      setCompare(idx === state.compareWith ? -1 : idx);
      return;
    }
    selectParticipant(idx === state.selected ? -1 : idx);
  }

  function selectParticipant(idx) {
    if (state.compareWith >= 0 && idx !== state.selected) {
      // Any selection change ends the compare (B was relative to old A).
      state.compareWith = -1;
      _comparePickArmed = false;
      disposeCompare();
      syncCompareBtn();
    }
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
    renderRaf = requestAnimationFrame(function () {
      renderPending = false;
      renderRaf = 0;
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
    if (three.clusterHulls) {
      for (i = 0; i < three.clusterHulls.labels.length; i++) {
        entries.push({
          el: three.clusterHulls.labels[i],
          pos: three.clusterHulls.meshes[i].position,
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
      // Honor the current weight — after a data refresh the sliders rebuild
      // and must not silently reset the user's lens.
      var current = state.weights[group.key] != null ? state.weights[group.key] : 1;
      var value = document.createElement("span");
      value.className = "map-weight-value";
      value.textContent = current.toFixed(1);
      label.appendChild(name);
      label.appendChild(value);

      var slider = document.createElement("input");
      slider.type = "range";
      slider.min = "0";
      slider.max = "2";
      slider.step = "0.1";
      slider.value = String(current);
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
        key: "showTrajectories",
        label: "Session trajectories",
        hint: "Each participant's path through behavior space across " +
          (state.data && state.data.windows ? state.data.windows.count : 5) +
          " session phases; the dot is the whole session",
        apply: function (on) {
          if (on) rebuildTrajectories();
          else disposeTrajectories();
        },
      },
      {
        key: "showClusterHulls",
        label: "Cluster hulls",
        hint: "Deterministic k-means in the weighted feature space; " +
          "translucent shells labeled by each cluster's top distinguishing features",
        apply: function (on) {
          var controls = document.getElementById("mapClusterControls");
          if (controls) controls.classList.toggle("hidden", !on);
          if (on) rebuildClusterHulls();
          else {
            disposeClusterHulls();
            renderClusterNotice(null);
          }
        },
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

  // ---- Layout mode (PCA vs manual axes) -------------------------------------

  // Manual mode never boots blank: X/Y/Z default to the first column of the
  // first three non-empty groups (then any remaining columns), skipping keys
  // already assigned to another axis.
  function ensureAxisDefaults() {
    var cols = state.data.columns;
    if (!cols.length) return;
    var used = {};
    ["x", "y", "z"].forEach(function (d) {
      if (axisFeatureIndex(d) >= 0) used[state.axisFeatures[d]] = true;
    });
    var firstByGroup = [];
    var seenGroup = {};
    var j;
    for (j = 0; j < cols.length; j++) {
      if (!seenGroup[cols[j].group]) {
        seenGroup[cols[j].group] = true;
        firstByGroup.push(cols[j].key);
      }
    }
    var pool = firstByGroup.concat(cols.map(function (c) { return c.key; }));
    ["x", "y", "z"].forEach(function (d) {
      if (axisFeatureIndex(d) >= 0) return;
      for (var p = 0; p < pool.length; p++) {
        if (!used[pool[p]]) {
          state.axisFeatures[d] = pool[p];
          used[pool[p]] = true;
          break;
        }
      }
    });
  }

  function renderLayoutControls() {
    var host = document.getElementById("mapLayoutMode");
    if (!host) return;
    var frag = document.createDocumentFragment();
    [
      { value: "pca", label: "Similarity (PCA)" },
      { value: "manual", label: "Manual axes" },
    ].forEach(function (opt) {
      var label = document.createElement("label");
      label.className = "map-layer-row";
      var radio = document.createElement("input");
      radio.type = "radio";
      radio.name = "mapLayoutMode";
      radio.value = opt.value;
      radio.checked = state.layoutMode === opt.value;
      radio.addEventListener("change", function () {
        if (!radio.checked) return;
        state.layoutMode = opt.value;
        if (opt.value === "manual") ensureAxisDefaults();
        var pickers = document.getElementById("mapAxisPickers");
        if (pickers) pickers.classList.toggle("hidden", opt.value !== "manual");
        renderAxisPickers();
        recompute(); // startLerp animates the mode switch for free
      });
      var text = document.createElement("span");
      text.textContent = opt.label;
      label.appendChild(radio);
      label.appendChild(text);
      frag.appendChild(label);
    });
    host.innerHTML = "";
    host.appendChild(frag);
  }

  // Five pickers: X/Y/Z position, Size (4th dim), Color (5th dim — the same
  // state.colorBy the explain-panel choropleth uses, so the two stay in
  // lockstep by construction).
  var AXIS_PICKER_DIMS = [
    { dim: "x", label: "X" },
    { dim: "y", label: "Y" },
    { dim: "z", label: "Z" },
    { dim: "size", label: "Size" },
    { dim: "color", label: "Color" },
  ];

  function renderAxisPickers() {
    var host = document.getElementById("mapAxisPickers");
    if (!host || !state.data) return;
    var frag = document.createDocumentFragment();
    AXIS_PICKER_DIMS.forEach(function (def) {
      var row = document.createElement("label");
      row.className = "map-axis-picker-row";
      var name = document.createElement("span");
      name.className = "map-axis-picker-label";
      name.textContent = def.label;
      var select = document.createElement("select");
      select.setAttribute("aria-label", def.label + " feature");
      if (def.dim === "size" || def.dim === "color") {
        var none = document.createElement("option");
        none.value = "";
        none.textContent = "(none)";
        select.appendChild(none);
      }
      var byGroup = {};
      state.data.groups.forEach(function (group) {
        var og = document.createElement("optgroup");
        og.label = group.label;
        byGroup[group.key] = og;
      });
      state.data.columns.forEach(function (col) {
        var opt = document.createElement("option");
        opt.value = col.key;
        opt.textContent = col.label;
        (byGroup[col.group] || select).appendChild(opt);
      });
      state.data.groups.forEach(function (group) {
        if (byGroup[group.key].children.length) {
          select.appendChild(byGroup[group.key]);
        }
      });
      var current = def.dim === "color"
        ? state.colorBy : state.axisFeatures[def.dim];
      select.value = current || "";
      select.addEventListener("change", function () {
        var key = select.value || null;
        if (def.dim === "color") {
          setColorBy(key); // never a recompute — color doesn't move dots
          return;
        }
        state.axisFeatures[def.dim] = key;
        recompute();
      });
      row.appendChild(name);
      row.appendChild(select);
      frag.appendChild(row);
    });
    host.innerHTML = "";
    host.appendChild(frag);
  }

  // Axis tripod tip labels: feature names in manual mode, plain X/Y/Z in PCA
  // (the sidebar legend explains the loadings there).
  function updateAxisTipText() {
    if (!three.axisLabels.length) return;
    var names = ["X", "Y", "Z"];
    var dims = ["x", "y", "z"];
    for (var c = 0; c < 3; c++) {
      var text = names[c];
      if (state.layoutMode === "manual") {
        var j = axisFeatureIndex(dims[c]);
        if (j >= 0) {
          var label = state.data.columns[j].label;
          // The declutterer estimates width from text length; keep tips short.
          text = label.length > 14 ? label.substring(0, 13) + "…" : label;
        }
      }
      three.axisLabels[c].el.textContent = text;
    }
  }

  // "Color by" choropleth: one feature key drives the dot color ramp instead
  // of the outlier score. Set from the explain panel / axis legend feature
  // names; cleared from the sidebar chip. Never a recompute — color doesn't
  // move the layout.
  function setColorBy(key) {
    state.colorBy = key || null;
    renderColorByChip();
    if (three.renderer) {
      styleDots();
      requestRender();
    }
  }

  function renderColorByChip() {
    var section = document.getElementById("mapColorBySection");
    var chip = document.getElementById("mapColorByChip");
    if (!section || !chip) return;
    var j = colorByIndex();
    section.classList.toggle("hidden", j < 0);
    renderAxisPickers(); // keep the manual-mode Color select in lockstep
    if (j < 0) return;
    chip.textContent = state.data.columns[j].label + " ✕";
    chip.title = "Stop coloring dots by this feature";
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
    if (state.layoutMode === "manual") {
      // Manual axes: each axis IS a feature; no variance shares to explain.
      ["x", "y", "z"].forEach(function (dim, c) {
        var j = axisFeatureIndex(dim);
        var dt = document.createElement("dt");
        dt.textContent = axisNames[c] + " — " +
          (j >= 0 ? state.data.columns[j].label : "(unset)");
        dl.appendChild(dt);
      });
      var note = document.createElement("dd");
      note.textContent = "z-scored per feature; the center is the cohort mean";
      dl.appendChild(note);
      frag.appendChild(dl);
      els.axisList.innerHTML = "";
      els.axisList.appendChild(frag);
      return;
    }
    state.components.forEach(function (comp, c) {
      var loadings = comp
        .map(function (v, j) { return { j: j, v: v }; })
        .sort(function (a, b) { return Math.abs(b.v) - Math.abs(a.v); })
        .slice(0, 3);
      var dt = document.createElement("dt");
      dt.textContent = axisNames[c] + " — " +
        Math.round(state.variance[c] * 100) + "% of variance";
      var dd = document.createElement("dd");
      loadings.forEach(function (l, li) {
        if (li) dd.appendChild(document.createTextNode(", "));
        var btn = document.createElement("button");
        btn.type = "button";
        btn.className = "map-feature-name";
        btn.textContent =
          (l.v >= 0 ? "+" : "−") + " " + state.data.columns[l.j].label;
        btn.title = "Color all dots by this feature";
        btn.addEventListener("click", (function (key) {
          return function () { setColorBy(key); };
        })(state.data.columns[l.j].key));
        dd.appendChild(btn);
      });
      dl.appendChild(dt);
      dl.appendChild(dd);
    });
    frag.appendChild(dl);
    els.axisList.innerHTML = "";
    els.axisList.appendChild(frag);
  }

  function renderExplain(idx) {
    // While a compare is active the panel belongs to the pair, not to A.
    if (state.compareWith >= 0 && state.compareWith !== idx) {
      renderCompare(idx, state.compareWith);
      return;
    }
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

    // Cluster membership (when the hull layer is on).
    if (three.clusterHulls) {
      for (var ci = 0; ci < three.clusterHulls.clusters.length; ci++) {
        if (three.clusterHulls.clusters[ci].indexOf(idx) < 0) continue;
        var clusterLine = document.createElement("p");
        clusterLine.className = "map-feature-detail";
        clusterLine.textContent =
          "Cluster: " + three.clusterHulls.labels[ci].textContent;
        frag.appendChild(clusterLine);
        break;
      }
    }

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
      var name = document.createElement("button");
      name.type = "button";
      name.className = "map-feature-name";
      name.textContent = col.label;
      name.title = "Color all dots by this feature";
      name.addEventListener("click", (function (key) {
        return function () { setColorBy(key); };
      })(col.key));
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

  // A-vs-B panel, shown in the explain pane while a compare is active:
  // per-group weighted-z distance bars (bead sizes in numbers) and the top
  // differing features with both raw values.
  function renderCompare(a, b) {
    var data = state.data;
    var pidA = data.participants[a];
    var pidB = data.participants[b];
    els.explainTitle.textContent = pidA + " vs " + pidB;
    var diff = computePairDiff(a, b);
    var frag = document.createDocumentFragment();

    var head = document.createElement("div");
    head.className = "map-explain-section";
    head.textContent = "Difference by signal group";
    frag.appendChild(head);

    var maxDist = 0;
    data.groups.forEach(function (g) {
      var e = diff.groups[g.key];
      if (e) maxDist = Math.max(maxDist, e.dist);
    });
    var bars = document.createElement("div");
    bars.className = "map-group-bars";
    data.groups.forEach(function (group) {
      var e = diff.groups[group.key] || { dist: 0, count: 0 };
      var row = document.createElement("div");
      row.className = "map-feature-row";
      var label = document.createElement("div");
      label.className = "map-feature-label";
      var name = document.createElement("span");
      name.className = "map-compare-name";
      var swatch = document.createElement("span");
      swatch.className = "map-compare-swatch";
      swatch.style.background = groupCssColor(group.key);
      name.appendChild(swatch);
      name.appendChild(document.createTextNode(group.label));
      var val = document.createElement("span");
      val.className = "map-feature-z";
      val.textContent = e.count ? e.dist.toFixed(2) : "no shared data";
      label.appendChild(name);
      label.appendChild(val);
      var track = document.createElement("div");
      track.className = "map-outlier-bar-track";
      var bar = document.createElement("span");
      bar.className = "map-outlier-bar";
      bar.style.display = "block";
      bar.style.width = (maxDist > 0 ? (e.dist / maxDist) * 100 : 0) + "%";
      track.appendChild(bar);
      row.appendChild(label);
      row.appendChild(track);
      bars.appendChild(row);
    });
    frag.appendChild(bars);

    var featHead = document.createElement("div");
    featHead.className = "map-explain-section";
    featHead.textContent = "Where they differ most";
    frag.appendChild(featHead);

    diff.ranked.slice(0, TOP_FEATURES).forEach(function (item) {
      var col = data.columns[item.j];
      var dzRaw = item.zA - item.zB; // unweighted σ gap, same sign as dz
      var row = document.createElement("div");
      row.className = "map-feature-row";

      var label = document.createElement("div");
      label.className = "map-feature-label";
      var name = document.createElement("span");
      name.textContent = col.label;
      var zEl = document.createElement("span");
      zEl.className = "map-feature-z";
      zEl.textContent = (dzRaw >= 0 ? pidA : pidB) + " +" +
        Math.abs(dzRaw).toFixed(1) + "σ";
      label.appendChild(name);
      label.appendChild(zEl);

      var detail = document.createElement("div");
      detail.className = "map-feature-detail";
      detail.textContent = pidA + ": " + fmtNum(data.matrix[a][item.j]) +
        " · " + pidB + ": " + fmtNum(data.matrix[b][item.j]);

      // Signed bar off the center line: right = A higher, left = B higher.
      var track = document.createElement("div");
      track.className = "map-feature-bar-track";
      var bar = document.createElement("div");
      bar.className = "map-feature-bar" + (dzRaw < 0 ? " is-negative" : "");
      var half = Math.min(Math.abs(dzRaw) / (2 * Z_CLAMP), 1) * 50;
      if (dzRaw >= 0) {
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

    if (!diff.ranked.length) {
      var none = document.createElement("p");
      none.className = "map-feature-detail";
      none.textContent = "No shared features to diff.";
      frag.appendChild(none);
    }

    var exit = document.createElement("button");
    exit.type = "button";
    exit.className = "btn btn-small";
    exit.textContent = "End compare";
    exit.addEventListener("click", function () { setCompare(-1); });
    frag.appendChild(exit);

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
    document.getElementById("mapColorByChip")
      .addEventListener("click", function () { setColorBy(null); });
    document.getElementById("mapCompareBtn")
      .addEventListener("click", function () {
        if (state.compareWith >= 0) { setCompare(-1); return; }
        _comparePickArmed = !_comparePickArmed;
        syncCompareBtn();
      });
    var clusterK = document.getElementById("mapClusterK");
    if (clusterK) {
      clusterK.addEventListener("change", function () {
        state.clusterK = parseInt(clusterK.value, 10) || 0;
        if (state.showClusterHulls) {
          rebuildClusterHulls();
          requestRender();
        }
      });
    }
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
    onResize();
    // Re-activation (tab switch / hub Refresh): when the hub refetched since
    // this matrix was loaded, re-fetch api/data and rebuild the layout — the
    // same dataVersion staleness contract Convergence/Metadata use. Otherwise
    // only the hub-data-derived layers need a rebuild.
    if (hubDataVersion() !== _matrixVersion) {
      reloadData();
      return;
    }
    if (state.showAllMoments && three.renderer) rebuildMoments();
    requestRender();
  }

  function mapDeactivate() {
    // Pause the free-running replay sweep. Its rAF loop guards on replay.playing,
    // and setReplayPlaying(false) also resyncs the Play/Pause button + scrub UI.
    if (replay.playing) setReplayPlaying(false);
    // Cancel any queued on-demand render — no point painting a hidden canvas.
    if (renderRaf) {
      cancelAnimationFrame(renderRaf);
      renderRaf = 0;
      renderPending = false;
    }
    // Clear the transient hover tooltip; the drill-down drawer + selection persist
    // so returning to the tab keeps the user's context.
    if (els.tooltip) els.tooltip.classList.add("hidden");
    state.hovered = -1;
  }

  window.ClipgenOverview.mapActivate = mapActivate;
  window.ClipgenOverview.mapDeactivate = mapDeactivate;
  window.ClipgenOverview.mapResize = onResize;
})();
