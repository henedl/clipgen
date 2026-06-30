# LOC-reduction opportunities — plan

Status: **in progress** (2026-06-30). Items **1 (1a + 1b)** and **2** landed (see check-marks
below); items 3–5 still open. Investigation targets for genuinely *reducing* total
lines of code (not relocating them). Each item is sized as one or more focused `refactor:`
commits. Check items off and add a "Done" note as they land (per AGENTS.md plan-maintenance rule).

## Framing

Hub→satellite carves (see [FRONTEND-REFACTORS-PLAN.md](FRONTEND-REFACTORS-PLAN.md)) **relocate**
lines — the hub shrinks but the total is unchanged. Real reduction comes from **deduplication
and deleting redundancy**. Everything below targets repeated shapes, ordered by leverage.

Baseline (2026-06-30): Python ~40.2k lines, JS ~38.8k, CSS ~17.9k. The four Flask server
files alone are ~9.8k lines.

## Do NOT chase

- **The test suite** (~26k lines) — it's guard rails, not fat. Don't compress it.
- **Comments / docstrings** — the house style is deliberately heavy inline documentation
  (AGENTS.md). Not a target.
- **More carves** — they don't change the total; they live in the other plan.

---

## 1. Flask route boilerplate — highest leverage

The four server files (`server.py`, `screenspace_server.py`, `transcripts_server.py`,
`workflows_server.py`) span ~9.8k lines across **149 routes**, with **416 `jsonify(` calls**,
**281 `{"ok": False, "error": …}` returns**, and **113 try/excepts**. Sampling
(`screenspace_server.py` pins routes ~461–540) confirms near-identical per-route scaffolding:

```python
return jsonify({"ok": False, "error": "timestamp must be a number"}), 400
...
return jsonify({"ok": True, "pin": pin})
```

- [x] **1a. Response helpers.** **Done** — `ok(**fields)` / `err(msg, code=400)` live in the new
  Flask-only `server_utils.py` (not `utils.py`, which is deliberately Flask-free). Applied across all
  four blueprints: workflows, transcripts, screenspace, server. Non-standard envelopes left raw on
  purpose (`{"ok": False}` with no `error`; `{"ok": False, "generating": True, ...}` poll/cancelled
  payloads; `{"errors": [...]}` plural; the 202 google-auth envelope; bare `jsonify(var)`; the two
  `api_sheet` payloads `test_shared_constants` guards by source text; binary/SSE/VTT routes).
- [x] **1b. `@json_endpoint` decorator** — **Done** (added to `server_utils.py`). Catches **only**
  `ApiError` (not bare `Exception`) so it never swallows real 500s or fights routes with their own
  cleanup/`finally`. Applied conservatively: most server try/excepts catch *specific* exceptions or do
  resource cleanup (`_release_busy`), so they were left intact; the decorator currently backs the one
  `parse_number_arg` site (screenspace `api_pins_create`). It's in place for future raise-based routes.
- **Risk:** low. **Approach:** one blueprint at a time; every route already has test coverage, so
  convert + run that blueprint's tests before the next. Watch for routes with non-standard
  envelopes (not all use `{"ok": …}`) — leave those or special-case, don't force-fit.
- **Payoff (realized):** net **~−310 lines** across the four files (workflows −15, transcripts −38,
  screenspace −120, server −138) + far more scannable routes. Full suite green (1745 passed).

## 2. Request-arg parsing/validation (rides on #1)

`CODE-REVIEW.md` mandates manual `float()/int()` parsing of string route params (Flask's `<float:>`
converter 404s when JS sends ints). The "parse → on failure return error dict" block is copy-pasted
dozens of times.

- [x] **2. `parse_number_arg(raw, name, *, int_only=False, min_=None, max_=None, finite=False)`** —
  **Done** (in `server_utils.py`). Parses one numeric value and raises `ApiError(400)` (caught by 1b's
  decorator) on bad value / out-of-bounds / non-finite. Applied where the clean *error-on-failure* shape
  existed (screenspace `api_pins_create`: an isinstance + `float()` + finite/bounds block → one line).
  **Note:** most numeric parsing in these files is *silent-fallback* (`except: x = default`) or
  *returns `None`*, which deliberately must NOT raise — those sites were left as-is, so #2's reach is
  smaller than the "dozens" first estimated. The helper is in place for future error-on-failure sites.

## 3. Cross-cutting JS dedup (these *delete* lines)

Shared with [FRONTEND-REFACTORS-PLAN.md](FRONTEND-REFACTORS-PLAN.md) Theme B — listed here because,
unlike the carves, they remove duplication rather than move it.

- [ ] **3a. `createSSEStream(url, {onMessage, onError, poll})`** in `utils.js` — collapses 4
  near-identical `EventSource → onerror → createPoller` blocks (`screenspace-tasks.js:1064`,
  `workflows-runs.js:77/142/1019`).
- [ ] **3b. Generalize the modal focus-trap** — `studio.js openModalTrap` (~`4627`) vs
  `transcripts.js confirmModelInstall` (~`2180`) are two implementations of the same lifecycle;
  promote one `openBlockingModal({onEscape, onBackdropClick})` to `utils.js`. Leave singleton
  pickers (`color-picker.js`, `settings-modal.js`) owning their own lifecycle.
- [ ] **3c. Form-input factories** (`rangeInput`/`numberInput`/`textInput`) and color-conversion —
  consolidate the duplicated copies (`screenspace-utils.js` vs `screenspace-multitool-params.js`;
  `color-picker.js` vs `screenspace-utils.js`). **Caution:** color conversion uses two
  *incompatible* HSV ranges (standard [0,1] vs OpenCV 0–180/0–255) — clarify names, don't blind-merge.
- **Est. payoff:** ~100–200 lines.

## 4. CSS shared-component promotion

Continues today's pass (radius-full, `--color-backdrop`, theme-toggle dedup landed in `0f1012b`).
The CSS audit flagged more repeated blocks across the page stylesheets:

- [ ] **4. Promote duplicated component blocks** to `primitives.css`/`tokens.css`: webkit
  **scrollbar** styling (~studio + start-overlay), the floating **popover/dropdown container**
  (position+shadow+border+padding, ~screenspace + studio), **icon-sizing** classes (`.ss-icon`/
  `.cg-icon`/`.so-icon` variants with raw px), and the **badge/pill** micro-label pattern.
- **Est. payoff:** ~20–30 rule blocks. **Risk:** medium (visual) — needs a browser check, and
  page CSS that loads after `primitives.css` can override, so verify specificity.

## 5. Manifest load/save patterns (investigation first)

`PERFORMANCE.md` notes the "load-on-startup, save-after-mutation" JSON-manifest shape recurs for
clipgen / screenspace / transcripts / workflows / stashes / settings.

- [ ] **5. Audit whether the manifest implementations are truly parallel.** If so, a small
  load/save helper could dedup them. **Confirm they're not subtly different first** — manifests have
  different schemas and mutation semantics; only the I/O envelope is a candidate, and forcing a
  shared abstraction over diverging shapes would be net-negative.

## Suggested order

1. **1a** (response helpers — biggest, safest win). 2. **1b** + **2** (decorator + arg parsing,
   building on 1a). 3. **3a/3b/3c** (JS dedup, independent). 4. **4** (CSS, needs browser check).
5. **5** only after the investigation confirms it's worthwhile.

Quantify before committing to a tier: a duplication scan (jscpd for JS/CSS, a `pylint`-style
duplicate-code pass for Python) would put hard block counts behind each item.
