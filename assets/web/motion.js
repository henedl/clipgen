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

  // One cached MediaQueryList — read `.matches` per call (cheap); never create a
  // new MQL per animation. This is the repo's first JS-side reduced-motion check,
  // complementing the CSS `@media (prefers-reduced-motion: reduce)` blocks.
  var REDUCED =
    typeof global.matchMedia === "function"
      ? global.matchMedia("(prefers-reduced-motion: reduce)")
      : { matches: false };

  // ---- Tweakable parameters (developer knobs) ----------------------------------
  // Durations are ms. All angles are degrees, distances px. Keep defaults short.
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
    // Overlay/modal "pop" (enter + exit) — mirrors the studio.css cg-overlay-pop
    // keyframe (translateY + scale + opacity). Meant for reused surfaces that toggle
    // display rather than being removed, so callers pair animateIn on show with a
    // guarded animateOut on hide (the newer entry animation cleanly supersedes a
    // stale forwards-filled exit; see showToast).
    pop: {
      duration: 150, // == tokens.css --duration-fast
      easing: "cubic-bezier(0.2, 0.7, 0.3, 1)",
      risePx: 8,
      scaleFrom: 0.98,
    },
    // Plain opacity fade (enter + exit) — the simplest kind, and what every other
    // kind collapses to under reduced motion / no WAAPI.
    fade: {
      duration: 150,
      easing: "ease",
    },
    // Stagger for *All variants: item i waits min(i*step, maxTotal) ms. The clamp
    // keeps "clear dozens of cards" finishing within ~maxTotal + duration.
    stagger: { step: 18, maxTotal: 180 },
    // Collapsed animation used under prefers-reduced-motion / no WAAPI.
    reduced: { duration: 80 },
    // Extra headroom for the belt-and-suspenders finish fallback timer.
    slackMs: 120,
    // Size awareness (used by kinds with sizeAware:true). `reference` is the object
    // diagonal (px) that maps to full intensity (≈ a region pill); larger objects
    // get a smaller scale (down to minScale) and, via durationGain, a slightly
    // longer, weightier animation so the motion never feels sudden.
    size: { reference: 100, minScale: 0.4, maxScale: 1.1, durationGain: 0.4 },
  };

  // ---- Keyframe builders (read PARAMS[kind]) -----------------------------------
  // Each keyframe carries its own `easing`, which governs the segment that STARTS
  // at that keyframe (the last keyframe's easing is ignored). Offsets are strictly
  // increasing in [0, 1].

  function buildStashKeyframes(p, scale) {
    scale = scale || 1;
    // jumpPx/hopPx are SIGNED (+ up, − down): a negative jumpPx crouches down to
    // spring-load, and a positive hopPx then springs up out of it.
    var up = -p.jumpPx * scale; // initial move (negative Y = up)
    var wig = Math.abs(p.wiggleDeg) * scale;
    var cycles = Math.max(0, p.wiggleCycles | 0); // 0 = no wiggle at all
    var hop = -p.hopPx * scale; // apex of the post-chime hop
    // Anticipation before the hop: a small wind-up opposite to the hop's travel —
    // for a spring-load this loads the crouch a touch deeper before the release.
    var windup = Math.abs(up) * 0.3;
    var dip = (up + (hop < up ? 1 : -1) * windup).toFixed(2);
    var down = Math.abs(p.slideDownPx) * scale;

    // Timeline as fractions of the total duration: a short rise, then the chime,
    // then the hop + drop reflow into whatever time the chime leaves. `wiggleSpan`
    // is how much of the animation the chime occupies — shrink it (or add cycles)
    // to speed the wiggle up, grow it to slow it down.
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

    // Damped "bell" ring: smooth, quickly-shrinking swings around the top. With
    // cycles === 0 (or wSpan === 0) the loop is skipped and the object simply holds
    // at the top — no wiggle — before hopping.
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

    // Settle upright, then a little playful hop — a small dip (anticipation), then
    // a quick spring up to a touch higher than before...
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

  // Opacity-only fade. Entry 0→1, exit 1→0. Shared by the `fade` kind and the
  // reduced-motion fallback (reducedFade).
  function buildFadeKeyframes(isEntry) {
    return isEntry ? [{ opacity: 0 }, { opacity: 1 }] : [{ opacity: 1 }, { opacity: 0 }];
  }

  // Overlay "pop": a small translateY + scale paired with a fade. Entry rises from
  // (risePx, scaleFrom, 0) to (0, 1, 1); exit is the reverse — settles down, shrinks,
  // fades. Mirrors studio.css cg-overlay-pop (translateY starts BELOW, +risePx).
  function buildPopKeyframes(p, isEntry) {
    var rise = Math.abs(p.risePx);
    var s = p.scaleFrom;
    var shown = { transform: "translateY(0px) scale(1)", opacity: 1 };
    var hidden = { transform: "translateY(" + rise + "px) scale(" + s + ")", opacity: 0 };
    return isEntry ? [hidden, shown] : [shown, hidden];
  }

  // ---- Size awareness ----------------------------------------------------------
  // The same tilt/travel looks far stronger on a big element than a small one.
  // measureSizeScale() returns a factor (≈1 for a pill-sized object, smaller for
  // larger ones) that size-aware kinds multiply into their rotation/translation so
  // pills stay lively while cards stay calm — from a single animation definition.
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

  // ---- Core runner -------------------------------------------------------------
  // Drives one element through one keyframe set, resolving exactly once when the
  // animation finishes. Robust against WAAPI quirks:
  //   - `.finished` can reject with AbortError if the animation is cancelled, so we
  //     funnel both fulfil and reject into the same `done`;
  //   - some engines have historically not settled `.finished` for detached nodes,
  //     so a `setTimeout` fallback guarantees `done` fires; a `settled` flag makes
  //     it idempotent.
  // `lockPointer` disables pointer events for the animation's lifetime (exits only —
  // prevents a mid-flight re-click on the element being removed).
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

  // Exit animation. `kind` ∈ {"stash","delete","pop","fade"}. opts.delay offsets the
  // start. fill:"forwards" holds the end state (invisible) so there's no flash before
  // the caller re-renders the list away (or, for reused surfaces, hides + re-shows).
  function animateOut(el, kind, opts) {
    opts = opts || {};
    if (REDUCED.matches || !HAS_WAAPI) return reducedFade(el, false, true);
    var p = PARAMS[kind] || PARAMS.delete;
    // Size-aware kinds calm down on larger objects: scale the motion by the
    // element's measured size and lengthen it slightly (weightier, less sudden).
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

  // Staggered exit for a whole list (clear-all / stash-all). Always resolves once
  // the LAST (max-delay) element settles. Under reduced motion the stagger is 0.
  function animateOutAll(els, kind, opts) {
    opts = opts || {};
    var arr = Array.prototype.slice.call(els || []);
    if (!arr.length) return Promise.resolve();
    var reduced = REDUCED.matches || !HAS_WAAPI;
    var step = opts.staggerStep != null ? opts.staggerStep : PARAMS.stagger.step;
    var maxT = opts.staggerMaxTotal != null ? opts.staggerMaxTotal : PARAMS.stagger.maxTotal;
    var p = PARAMS[kind] || PARAMS.delete;
    // Read pass first: measure every element's size up front, so we never interleave
    // getBoundingClientRect (layout read) with .animate() (write) across the batch.
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

  // Entry animation (e.g. a stash card landing). fill:"both" keeps the element at
  // the start (invisible) from the moment the animation is created — including the
  // brief window before it's connected to the DOM — so there's no first-frame flash.
  function animateIn(el, kind, opts) {
    opts = opts || {};
    if (REDUCED.matches || !HAS_WAAPI) return reducedFade(el, true, false);
    var p = PARAMS[kind] || PARAMS.stashLand;
    var kf;
    if (kind === "pop") kf = buildPopKeyframes(p, true);
    else if (kind === "fade") kf = buildFadeKeyframes(true);
    else kf = buildStashLandKeyframes(p);
    return runOne(el, kf, { duration: p.duration, delay: opts.delay || 0, easing: p.easing || "ease", fill: "both" }, false);
  }

  // FUTURE SEAM — "the card flies to the stash object".
  // Planned implementation (not built yet): clone `el` into a single fixed-position
  // `#cg-motion-layer` overlay appended lazily to <body>; read `el` and `targetEl`
  // bounding rects in ONE pass (never interleave getBoundingClientRect with
  // .animate() — that thrashes layout), animate the clone's transform to the delta
  // + shrink, fade the source, resolve on settle. For now it degrades to the stash
  // exit so callers can adopt the API today.
  function flyTo(el, targetEl, opts) {
    opts = opts || {};
    return animateOut(el, opts.fallbackKind || "stash", opts);
  }

  global.ClipgenMotion = {
    PARAMS: PARAMS,
    animateOut: animateOut,
    animateOutAll: animateOutAll,
    animateIn: animateIn,
    flyTo: flyTo,
  };
})(window);
