# Composer bug hunt

Status: fixed (2026-09-18).

Scope: Composer mutation routes and undo/redo history. Four verified findings,
all fixed with API / Node regressions in `tests/test_composer_server.py` and
`tests/test_js_units.py`.

## 1. Rejected annotation updates mutate live state

Priority: medium. Location: `source/composer_server.py:550–569`, `api_annotation_update`.

The route assigns a valid incoming span directly to the stored annotation before
validating geometry. Invalid geometry then returns HTTP 400 without restoring the
span or persisting it. The live manifest differs from disk; a later successful
mutation can persist the rejected change.

Reproduction:

1. POST `/composer/api/annotations` with:
   `{"participant":"P01","type":"text","span":{"start":1,"end":2},"geometry":{"text":"hello"}}`.
2. PATCH `/composer/api/annotations/<returned id>` with:
   `{"span":{"start":5,"end":6},"geometry":{"text":""}}`.
3. Observe HTTP 400 (`invalid geometry for type text`).
4. Inspect the manifest: the in-memory span is now `{start:5,end:6}`;
   the last persisted snapshot still has `{start:1,end:2}`.

Expected: a rejected update leaves both live and persisted state unchanged.

Implementation plan:

- [x] Validate supplied fields before mutating the stored annotation, under the existing lock.
- [x] Add an API regression in `tests/test_composer_server.py` covering the rejected combined update.
- [x] Assert live state and disk retain the original annotation after rejection and a subsequent successful mutation.
- [x] Verify valid combined updates still persist every requested field.

## 2. Marker trims accept invalid spans

Priority: medium. Location: `source/composer_server.py:327–358`, `api_trim_put`.

The route checks ordering before clamping a negative start to zero. It also
accepts non-finite floats. Both paths save invalid marker times and return success.

Reproduction:

1. PUT `/composer/api/trims/sheet:P01:1` with `{"start":-2,"end":-1}`.
   Actual: HTTP 200; stored trim has `start:0.0,end:-1.0`.
2. PUT the same endpoint with `{"start":0,"end":"NaN"}`.
   Actual: HTTP 200; response contains the literal `"end":NaN`.
   This is invalid JSON for browser `JSON.parse`, despite being accepted by Python.

Expected: accepted trims have finite times and satisfy
`0 <= start < end` with the configured minimum duration after normalization.
Invalid requests must not mutate or persist state.

Implementation plan:

- [x] Reject non-finite times using the existing numeric-validation pattern.
- [x] Normalize start before checking the final span's minimum duration.
- [x] Add API regressions for negative spans, NaN, and positive/negative infinity.
- [x] Verify rejected requests preserve prior trims and return valid JSON.
- [x] Retain valid trim metadata and time-only undo/redo behavior.

## 3. Restoring deleted objects breaks earlier undo entries

Priority: medium. Locations: `assets/web/composer.js:561–613` (`applyOp`,
`shiftHistory`), cut history recording at `643–673`, annotation history recording
at `750–775`; server creation at `source/composer_server.py:262–283` and `520–546`.

Undoing deletion recreates the object through POST, which assigns a new ID.
`applyOp` updates only the current deletion operation's snapshot. Earlier edit
operations still carry the deleted ID. Their next undo receives a missing-object
error and stays at the top of the undo stack, blocking further progress.
The same stale references occur in creation records and annotation group members.

Reproduction:

1. Create a cut, change its end time, then delete it.
2. Undo the deletion. The cut returns with a fresh server ID.
3. Undo the timing edit. Its PATCH still targets the original ID and fails.
4. Further Undo attempts retry the same failed operation.

The annotation branches have the same structure: create, edit a field, delete,
undo deletion, then undo the field edit.

Verified with the actual `applyOp` and `shiftHistory` functions evaluated in Node,
using stub API appliers that return a fresh ID on creation and reject the old ID
on edit. Output: restored `cut_new`, prior edit targets `cut_old`, then
`No cut cut_old`; the failed entry remains on the undo stack. Server creation
routes independently show unconditional UUID allocation. This was not a browser
interaction test.

Expected: undo/redo continues across deletion and recreation of the same object.

Implementation plan:

- [x] Remap object references across both history stacks when recreation changes an ID.
- [x] Include create/delete snapshots, edit IDs, and nested annotation group operations.
- [x] Cover cut and annotation create/edit/delete, full undo, and full redo sequences.
- [x] Verify grouped annotation history and repeated delete/restore cycles.
  Out of scope: an `ann-group` still applies sub-ops with `Promise.all`, so a
  sub-op failing mid-group leaves the others applied.
- [x] Exercise the sequence in a browser with existing UI tooling (`shot.py composer --eval-file`: edit, delete, undo ×2, redo ×2 with no failure toast).

## 4. Concurrent cut updates silently overwrite newer times

Priority: medium. Location: `source/composer_server.py:288–315`, `api_cut_update`.

The route reads both times under `_manifest_lock`, releases it for clamping,
then unconditionally writes both captured times back. Re-finding the cut guards
against deletion but does not guard against another update. Even a label-only
PATCH rewrites times omitted from its payload.

Reproduction, deterministically coordinated with thread events:

1. Seed `cut_test` with `{start:1,end:2,label:"old"}`.
2. Start PATCH `{"label":"new"}`; pause it inside `_clamp_span` after the snapshot.
3. Complete another PATCH `{"end":4}`. The stored end is now `4.0`.
4. Resume the rename. Both requests return HTTP 200, but the final cut is
   `{start:1,end:2,label:"new"}`. The successful timing edit has disappeared.

This can affect overlapping requests from separate tabs or clients. Clamping can
probe media on a cold cache, widening the interval between the two lock sections.

Verified against the actual Flask route in two threads, replacing `_clamp_span`
with an event-controlled identity function and persistence with a no-op. This
isolates the lock interleaving without ffprobe or filesystem writes.

Expected: a label-only update preserves current times; independent endpoint
updates do not silently restore stale fields omitted from the request.

Implementation plan:

- [x] Keep duration probing outside the lock, but merge and clamp against current state atomically.
- [x] Avoid rewriting times for label-only updates.
- [x] Add event-coordinated API regressions for rename versus timing edits.
- [x] Cover overlapping start-only/end-only edits and deletion during duration lookup.
- [x] Retain span validation and missing-cut errors without holding locks during media I/O.

## Verification and handoff

Backend findings reproduced on 2026-09-18 using `uv run --no-sync python` and a Flask
test client against the actual Composer blueprint. The reproduction initialized
an empty in-memory manifest and replaced `_persist_locked` with a snapshot
collector or no-op; it performed no media processing or project-state writes. Disk behavior
should be covered by the existing temporary-directory fixture when adding tests.

After implementation, run:

```sh
uv run --extra dev pytest -c tests/pytest.ini tests/test_composer_server.py
```

Follow `agents/skills/check/SKILL.md` before committing fixes and the UI check
procedure for history changes. This hunt did not run the full test suite, launch
a browser, or review other subsystems.
