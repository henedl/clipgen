/* motion.js — ClipgenMotion: shared micro-animation engine.
 *
 * FIRST Web Animations API (WAAPI, `element.animate`) usage in this codebase.
 * Everything else animates via CSS class toggles; this module is deliberately
 * different because the requirements — developer-tweakable parameters and smooth
 * performance with *dozens* of elements animating at once — are best met by WAAPI:
 *   - every tunable lives in one JS object (ClipgenMotion.PARAMS), so a developer
 *     can retune the feel from the console (e.g. `ClipgenMotion.PARAMS.stash.wiggleDeg = 9`);
 *   - `animation.finished` is a Promise, giving a clean "animate, then commit" hook
 *     that matches the house `.then()`-chaining style at every call site;
 *   - keyframes touch ONLY `transform` + `opacity`, so the browser runs them on the
 *     compositor thread (no layout/paint per frame) — this is what makes "clear 40
 *     cards" cheap. (See studio.css `.queue-card`: an always-on `will-change` once
 *     saturated the compositor, so we promote layers only transiently, below.)
 *   - kinds flagged `sizeAware` scale their intensity to the element's measured
 *     size, so ONE definition reads right on both small pills and large cards (the
 *     same tilt/travel looks far more dramatic on a big object).
 *
 * Shared by Screenspace region pills and Studio queue/stash cards; loaded right
 * after utils.js on both pages. Extensible to any surface — load motion.js and call
 * the same API. All functions ALWAYS resolve (never reject) so the caller's commit
 * logic runs on every path (no WAAPI, reduced motion, detached node, thrown animate).
 *
 * Public surface (window.ClipgenMotion):
 *   PARAMS                              developer-tweakable config (see below)
 *   animateOut(el, kind, opts)  -> Promise   exit: "stash" | "delete" | "pop" | "fade"
 *   animateOutAll(els, kind, opts) -> Promise  staggered exit for whole-list clears
 *   animateIn(el, kind, opts)   -> Promise   entry: "stashLand" | "pop" | "fade"
 *   flyTo(el, targetEl, opts)   -> Promise   FUTURE seam (ghost fly-to-target)
 */
(function (global) {
  "use strict";

  // Feature-detect once. If WAAPI is absent we skip straight to the commit.
  var HAS_WAAPI =
    typeof Element !== "undefined" &&
    Element.prototype &&
    typeof Element.prototype.animate === "function";

  // One cached MediaQueryList; read `.matches` per call, never create one per animation.
  var REDUCED =
    typeof global.matchMedia === "function"
      ? global.matchMedia("(prefers-reduced-motion: reduce)")
      : { matches: false };

  // ---- Tweakable parameters: durations ms, angles deg, distances px ----
  var PARAMS = {
    // Stash exit: little jump up, joyous wiggle at the top, then slide down + dissolve.
    stash: {
      duration: 430,
      easing: "cubic-bezier(0.2, 0.7, 0.3, 1)",
      jumpPx: 0, // initial move: + hops up, − crouches down (spring-load)
      wiggleDeg: 3, // peak rotation of the chime (gentle)
      wiggleCycles: 1, // number of chime swings (0 = no wiggle at all)
      wiggleSpan: 0.40, // fraction of the total duration the chime occupies (its "speed")
      hopPx: 12, // post-chime hop: + springs up, − drops
      slideDownPx: 16, // how far it slides down as it fades
      sizeAware: true, // scale the motion down on larger objects (cards vs pills)
    },
    // Delete/clear exit: falls downward, tilts lopsided, fades out fast.
    delete: {
      duration: 260,
      easing: "cubic-bezier(0.4, 0, 1, 1)", // ease-in ≈ gravity
      fallPx: 22,
      tiltDeg: 8,
      sizeAware: false, // reads well at any size already — leave it untouched
    },
    // Stash-card landing (entry) — the freshly-saved stash card rises + scales in.
    stashLand: {
      duration: 190,
      easing: "cubic-bezier(0.2, 0.7, 0.3, 1)",
      risePx: 6,
      scaleFrom: 0.96,
    },
    // Overlay pop (enter + exit), mirrors studio.css cg-overlay-pop; pair animateIn/animateOut, see showToast.
    pop: {
      duration: 150, // == tokens.css --duration-fast
      easing: "cubic-bezier(0.2, 0.7, 0.3, 1)",
      risePx: 8,
      scaleFrom: 0.98,
    },
    // Plain opacity fade; every kind collapses to this under reduced motion / no WAAPI.
    fade: {
      duration: 150,
      easing: "ease",
    },
    // *All stagger: item i waits min(i*step, maxTotal) ms; clamp bounds long lists.
    stagger: { step: 18, maxTotal: 180 },
    // Collapsed animation used under prefers-reduced-motion / no WAAPI.
    reduced: { duration: 80 },
    // Extra headroom for the belt-and-suspenders finish fallback timer.
    slackMs: 120,
    // Size awareness: `reference` diagonal px = full intensity; larger objects scale down, run longer.
    size: { reference: 100, minScale: 0.4, maxScale: 1.1, durationGain: 0.4 },
  };

  // ---- Keyframe builders: each `easing` governs the segment it starts; offsets ascend in [0,1].

  function buildStashKeyframes(p, scale) {
    scale = scale || 1;
    // jumpPx/hopPx are signed: + up, − down; negative jumpPx crouches to spring-load.
    var up = -p.jumpPx * scale; // initial move (negative Y = up)
    var wig = Math.abs(p.wiggleDeg) * scale;
    var cycles = Math.max(0, p.wiggleCycles | 0); // 0 = no wiggle at all
    var hop = -p.hopPx * scale; // apex of the post-chime hop
    // Anticipation: a small wind-up opposite the hop before release.
    var windup = Math.abs(up) * 0.3;
    var dip = (up + (hop < up ? 1 : -1) * windup).toFixed(2);
    var down = Math.abs(p.slideDownPx) * scale;

    // Offsets are duration fractions: rise, chime (`wiggleSpan` of total), then hop + drop.
    var rise = 0.16;
    var minTail = 0.06; // always reserve a little time for the hop + drop
    var wSpan = clamp(p.wiggleSpan != null ? p.wiggleSpan : 0.26, 0, 1 - rise - minTail);
    var wEnd = rise + wSpan;
    var rem = 1 - wEnd;
    var settleOff = wEnd + rem * 0.09;
    var dipOff = wEnd + rem * 0.28;
    var hopOff = wEnd + rem * 0.52;

    var frames = [
      // smooth ease-out hop up (no springy overshoot — keeps it gentle)
      { offset: 0, transform: "translateY(0px) rotate(0deg)", opacity: 1, easing: "cubic-bezier(0.22, 0.61, 0.36, 1)" },
      // reach the top
      { offset: rise, transform: "translateY(" + up + "px) rotate(0deg)", opacity: 1, easing: "ease-in-out" },
    ];

    // Damped swings around the top; cycles === 0 or wSpan === 0 skips the wiggle.
    var n = wSpan > 0.001 ? cycles * 2 : 0;
    for (var i = 1; i <= n; i++) {
      var t = i / (n + 1);
      var off = rise + wSpan * t;
      var dir = i % 2 === 1 ? 1 : -1;
      var mag = wig * (1 - 0.55 * t); // strong damping → each swing much smaller
      frames.push({
        offset: +off.toFixed(4),
        transform: "translateY(" + up + "px) rotate(" + (dir * mag).toFixed(2) + "deg)",
        opacity: 1,
        easing: "ease-in-out",
      });
    }

    // Settle upright, dip (anticipation), then spring up past the rest height...
    frames.push({ offset: +settleOff.toFixed(4), transform: "translateY(" + up + "px) rotate(0deg)", opacity: 1, easing: "ease-in-out" });
    frames.push({ offset: +dipOff.toFixed(4), transform: "translateY(" + dip + "px) rotate(0deg)", opacity: 1, easing: "cubic-bezier(0.2, 0.7, 0.35, 1)" });
    frames.push({ offset: +hopOff.toFixed(4), transform: "translateY(" + hop + "px) rotate(0deg)", opacity: 1, easing: "cubic-bezier(0.4, 0, 1, 1)" });
    // ...and release into the drop-down fadeout.
    frames.push({ offset: 1, transform: "translateY(" + down + "px) rotate(0deg)", opacity: 0 });
    return frames;
  }

  function buildDeleteKeyframes(p, scale) {
    scale = scale || 1;
    var fall = Math.abs(p.fallPx) * scale;
    var tilt = p.tiltDeg * scale;
    return [
      { offset: 0, transform: "translateY(0px) rotate(0deg)", opacity: 1, easing: "cubic-bezier(0.4, 0, 1, 1)" },
      { offset: 0.3, transform: "translateY(2px) rotate(" + (tilt * 0.5).toFixed(2) + "deg)", opacity: 0.9 },
      { offset: 1, transform: "translateY(" + fall + "px) rotate(" + tilt + "deg)", opacity: 0 },
    ];
  }

  function buildStashLandKeyframes(p) {
    var rise = Math.abs(p.risePx);
    var s = p.scaleFrom;
    return [
      { offset: 0, transform: "translateY(" + -rise + "px) scale(" + s + ")", opacity: 0 },
      { offset: 1, transform: "translateY(0px) scale(1)", opacity: 1 },
    ];
  }

  // Opacity-only fade; shared by the `fade` kind and reducedFade.
  function buildFadeKeyframes(isEntry) {
    return isEntry ? [{ opacity: 0 }, { opacity: 1 }] : [{ opacity: 1 }, { opacity: 0 }];
  }

  // Overlay pop: translateY + scale + fade; exit reverses. Mirrors studio.css cg-overlay-pop (starts below).
  function buildPopKeyframes(p, isEntry) {
    var rise = Math.abs(p.risePx);
    var s = p.scaleFrom;
    var shown = { transform: "translateY(0px) scale(1)", opacity: 1 };
    var hidden = { transform: "translateY(" + rise + "px) scale(" + s + ")", opacity: 0 };
    return isEntry ? [hidden, shown] : [shown, hidden];
  }

  // ---- Size awareness: measureSizeScale() gives ≈1 for pills, less for cards; size-aware kinds multiply in.
  function clamp(v, lo, hi) {
    return v < lo ? lo : v > hi ? hi : v;
  }
  function measureSizeScale(el) {
    if (!el || !el.getBoundingClientRect) return 1;
    var r = el.getBoundingClientRect();
    var diag = Math.sqrt(r.width * r.width + r.height * r.height);
    if (!diag) return 1; // detached / unmeasurable → neutral
    return clamp(PARAMS.size.reference / diag, PARAMS.size.minScale, PARAMS.size.maxScale);
  }

  // ---- Core runner. Resolves once; `.finished` may reject or never settle, hence the timer.
  function runOne(el, keyframes, timing, lockPointer) {
    return new Promise(function (resolve) {
      if (!el || !HAS_WAAPI) {
        resolve();
        return;
      }
      var settled = false;
      var done = function () {
        if (settled) return;
        settled = true;
        try {
          el.style.willChange = "";
          if (lockPointer) el.style.pointerEvents = "";
        } catch (e) {
          /* element may be detached; ignore */
        }
        resolve();
      };

      var anim;
      try {
        // Transiently promote to a composited layer for the animation only.
        el.style.willChange = "transform, opacity";
        if (lockPointer) el.style.pointerEvents = "none";
        anim = el.animate(keyframes, timing);
      } catch (e) {
        done(); // animate() threw — commit anyway so state isn't stranded
        return;
      }

      if (anim && anim.finished) {
        Promise.resolve(anim.finished).then(done, done);
      }
      var total = (timing.duration || 0) + (timing.delay || 0) + PARAMS.slackMs;
      global.setTimeout(done, total);
    });
  }

  // ---- Public API --------------------------------------------------------------

  function reducedFade(el, isEntry, lockPointer) {
    var kf = buildFadeKeyframes(isEntry);
    return runOne(
      el,
      kf,
      { duration: PARAMS.reduced.duration, delay: 0, easing: "linear", fill: isEntry ? "backwards" : "forwards" },
      lockPointer
    );
  }

  // Exit. fill:"forwards" holds the invisible end state until the caller re-renders or hides.
  function animateOut(el, kind, opts) {
    opts = opts || {};
    if (REDUCED.matches || !HAS_WAAPI) return reducedFade(el, false, true);
    var p = PARAMS[kind] || PARAMS.delete;
    // Size-aware kinds: smaller motion and slightly longer duration on larger elements.
    var scale = 1;
    var dur = p.duration;
    if (p.sizeAware) {
      scale = opts.sizeScale != null ? opts.sizeScale : measureSizeScale(el);
      dur = Math.round(p.duration * (1 + PARAMS.size.durationGain * Math.max(0, 1 - scale)));
    }
    var kf;
    if (kind === "stash") kf = buildStashKeyframes(p, scale);
    else if (kind === "pop") kf = buildPopKeyframes(p, false);
    else if (kind === "fade") kf = buildFadeKeyframes(false);
    else kf = buildDeleteKeyframes(p, scale);
    return runOne(el, kf, { duration: dur, delay: opts.delay || 0, easing: p.easing || "ease", fill: "forwards" }, true);
  }

  // Staggered exit for a whole list; resolves when the last element settles.
  function animateOutAll(els, kind, opts) {
    opts = opts || {};
    var arr = Array.prototype.slice.call(els || []);
    if (!arr.length) return Promise.resolve();
    var reduced = REDUCED.matches || !HAS_WAAPI;
    var step = opts.staggerStep != null ? opts.staggerStep : PARAMS.stagger.step;
    var maxT = opts.staggerMaxTotal != null ? opts.staggerMaxTotal : PARAMS.stagger.maxTotal;
    var p = PARAMS[kind] || PARAMS.delete;
    // Read pass first: measure all sizes before any .animate() to avoid layout thrash.
    var scales = arr.map(function (el) {
      return !reduced && p.sizeAware ? measureSizeScale(el) : 1;
    });
    // Write pass: start every animation with its pre-measured scale.
    return Promise.all(
      arr.map(function (el, i) {
        var delay = reduced ? 0 : Math.min(i * step, maxT);
        return animateOut(el, kind, { delay: delay, sizeScale: scales[i] });
      })
    );
  }

  // Entry. fill:"both" hides the element from creation, even before it is connected.
  function animateIn(el, kind, opts) {
    opts = opts || {};
    // Force layout first: an element shown this same tick would otherwise skip its opening frames.
    if (el) void el.offsetWidth;
    if (REDUCED.matches || !HAS_WAAPI) return reducedFade(el, true, false);
    var p = PARAMS[kind] || PARAMS.stashLand;
    var kf;
    if (kind === "pop") kf = buildPopKeyframes(p, true);
    else if (kind === "fade") kf = buildFadeKeyframes(true);
    else kf = buildStashLandKeyframes(p);
    return runOne(el, kf, { duration: p.duration, delay: opts.delay || 0, easing: p.easing || "ease", fill: "both" }, false);
  }

  // FUTURE SEAM: ghost fly-to-target (clone into a fixed layer). Degrades to the stash exit today.
  function flyTo(el, targetEl, opts) {
    opts = opts || {};
    return animateOut(el, opts.fallbackKind || "stash", opts);
  }

  // Shared reduced-motion answer; callers (utils.js popModalOut) must not re-query matchMedia.
  function isReduced() {
    return !!REDUCED.matches || !HAS_WAAPI;
  }

  global.ClipgenMotion = {
    PARAMS: PARAMS,
    animateOut: animateOut,
    animateOutAll: animateOutAll,
    animateIn: animateIn,
    flyTo: flyTo,
    isReduced: isReduced,
  };
})(window);
