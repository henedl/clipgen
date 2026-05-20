# Bugs — clipgen

## Context

A May 2026 bug hunt surfaced a small set of real, verified issues in clipgen. The "safe to fix" items (defensive serialization, KeyError hardening, NaN guards, CLI mode validation) shipped as one bundled fix. The two items in this document are **moderate-risk** — they touch state-machine code or persisted identifiers, and need a deliberate design pass before changing.

The full hunt write-up — including the false-alarm patterns to skip in future passes — lives at `~/.claude/plans/system-instruction-you-are-working-cozy-teacup.md`.

---

## H1 — AgentOrchestrator stop-vs-restart race

### Where

`transcripts_server.py`:
- `AgentOrchestrator.stop`, lines 1232–1246
- `AgentOrchestrator.run_agent` finally block, lines 1342–1346
- `AgentOrchestrator.run_agent` claim block, lines 1289–1297

### What's wrong

`stop(agent_key, participant)` removes `participant` from `_in_flight[agent_key]` and sets the cancel event, **but does not pop the cancel event from `_cancel_events`**. The intent is to let the daemon thread's `finally` block do that cleanup.

If a new `run_agent` for the same `(agent_key, participant)` claims the slot between `stop()` releasing it and the old daemon thread reaching its `finally`, the old `finally` clobbers the new run's bookkeeping:

```
T0  run_agent #1 claims in_flight[k] and stores cancel_event_1
T1  user clicks Stop → stop() removes participant from in_flight, sets cancel_event_1
T2  user clicks Regenerate → run_agent #2 claims in_flight[k], overwrites cancel_event_1 → cancel_event_2 in _cancel_events
T3  daemon thread #1 reaches finally → discards participant from in_flight (removes #2's claim!)
                                      → pops _cancel_events[k][participant] (removes #2's event!)
T4  thread #2 is still running but:
      - is_generating() returns False (UI shows idle)
      - stop() returns False (uncancellable)
      - its own finally will be a no-op on already-empty dicts
```

### How to reproduce

User-facing: open Transcripts UI, start a summary, click Stop, immediately click Regenerate within a second or two. Sometimes the second run becomes uncancellable and the UI shows it as idle even while it streams tokens.

Programmatic repro idea: spawn run_agent, call stop() but block the daemon thread from completing (monkeypatch the agent's run callable to take a `threading.Event`), call run_agent again, then unblock the first thread and assert that `is_generating` still reports True and stop() can cancel the second run.

### Why this is moderate risk to fix

The orchestrator's docstring at lines 1206–1213 explicitly calls out the "finish-vs-cancel race" the current design is meant to handle. Any fix must preserve those invariants:

- The cancel event must still cause the Ollama HTTP read loop to unblock promptly.
- The double-cancel check (line 1317 + line 1326) must still drop in-flight results from a cancelled run.
- `_threads[agent_key].discard(t)` cleanup must still happen when the thread truly exits.

It's easy to fix one race and introduce another. The bug hunt didn't enumerate every interleaving — only the one above.

### Suggested approaches

**Option A: per-run identity.** Generate a `run_id` (e.g. `uuid4()`) inside `run_agent`, store it as `_cancel_events[agent_key][participant] = (run_id, event)`, and have the finally pop only if the slot still holds *this* `run_id`:

```python
finally:
    with self._lock:
        slot = self._cancel_events.get(agent_key, {}).get(participant)
        if slot is not None and slot[0] == run_id:
            self._in_flight[agent_key].discard(participant)
            self._cancel_events[agent_key].pop(participant, None)
        self._threads[agent_key].discard(t)
```

Stop also matches on `run_id` if multiple runs could be live (not strictly needed today — only one slot per `(agent, participant)`).

**Option B: stop fully cleans up.** Have `stop()` pop both `_in_flight` *and* `_cancel_events`, and have the daemon thread's finally do nothing when its cancel event is set (defer cleanup entirely to whoever called stop). This is simpler but loses the "finally always cleans up" invariant — if the daemon crashes after stop, leak detection is harder.

Option A is the better match for the existing comment style.

### Verification plan

- Add a unit test in `tests/` that drives the interleaving above using `threading.Event` to block the agent's `run` callable. Assert `is_generating` and `stop` behave correctly across both runs.
- Manually verify in the Transcripts UI: Stop-then-Regenerate ten times in a row, check the UI never gets stuck.

---

## M1 — Synthetic artifact ID collision when `cell` is missing

### Where

`utils.py:868` (in `build_artifact_record`):

```python
return {
    "id": f"a{cell_row or 0}c{cell_col or 0}s{seg_idx}",
    ...
}
```

### What's wrong

`cell_row` / `cell_col` come from `getattr(cell, "row", None)` and `getattr(cell, "col", None)`. When either is `None` (or zero), the `or 0` fallback kicks in and the id collapses to `a0c0s{seg_idx}`.

Real synthetic clips from `--ss-clips` and `--transcript-clips` are namespaced correctly:

- `cli.py:2018-2028` builds `SimpleNamespace(value="", row=-(cluster_idx + 1), col=cell_col)`.
- Negative `cell_row` defeats the `or 0` fallback (it's truthy), so ids look like `a-1c1s0`, `a-2c1s0`, etc.
- `_SS_CLIPS_CELL_COL` and `_TRANSCRIPT_CLIPS_CELL_COL` differ, so cross-mode collisions are avoided.

The bug is **latent**: any *future* caller of `build_artifact_record` that forgets the synthetic-cell convention — passing `cell=None`, or a cell-like object without `.row`/`.col` — will silently mint ids of the form `a0c0s{seg_idx}`. Two such records for the same `seg_idx` deduplicate in the manifest (which is keyed by `id` in `viewer.save_manifest`), one will overwrite the other.

### How to reproduce

Today: nothing in the call graph reaches `build_artifact_record` with `cell=None`. The collision is a regression-vector, not a current crash.

Programmatic repro idea: call `build_artifact_record` directly in a test with a clip whose `cell` is `None`, twice with different segment data and the same `seg_idx`. Assert the resulting ids are unique (they aren't, today).

### Why this is moderate risk to fix

The artifact id is the **dedup key** for the manifest. Anything that changes the id shape will:

- Cause artifacts written by older clipgen versions to **not** dedup against artifacts written by newer clipgen (the project rule is "no compat layers, just re-run" — so this is acceptable, but worth being explicit).
- Affect any frontend caches keyed by `artifact.id` (thumbnail caches, etc.) — those need to invalidate cleanly on the first run after upgrade.
- Affect any reel components whose source artifact id is referenced.

The bug hunt did not enumerate every consumer of `artifact.id`. A pre-fix sweep would catch surprises.

### Suggested approaches

**Option A: explicit guard.** Assert `cell_row is not None and cell_col is not None` at the top of `build_artifact_record`, document the synthetic-cell contract in the docstring, and let any new caller crash loudly rather than collide silently.

```python
if cell_row is None or cell_col is None:
    raise ValueError(
        "build_artifact_record requires a cell with row and col; "
        "synthetic records must use negative row to namespace ids "
        "(see _make_synthetic_clip_record)"
    )
```

**Option B: collision-proof id.** Mix in a stable but-distinct field when row/col are missing, e.g. `study`, `participant`, `source_filename`, or a hash. This avoids forcing callers to set synthetic rows but makes ids longer and breaks the current convention.

Option A matches the project preference for loud failures over silent fallbacks.

### Verification plan

- Add a unit test that calls `build_artifact_record` with `cell=None` and asserts it raises (Option A) or produces unique ids (Option B).
- Grep for all callers of `build_artifact_record` and confirm each path provides a real or synthetic cell.
- Run an end-to-end `--ss-clips` and `--transcript-clips` pass and confirm the manifest still dedups correctly within each mode.

---

## Out of scope

- Refactors and style improvements — not addressed in either tier.
- "No schema version on manifests" — explicit project rule (AGENTS.md hard rules).
- Test coverage gaps — out of scope for a bug hunt.
