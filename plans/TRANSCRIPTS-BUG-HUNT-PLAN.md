# Transcripts bug hunt

Status: two findings verified; fixes pending.

Scope: inline transcript editing and study-local correction rules. No production
code or test changes. Separate from the Composer and Workflows hunts.

## 1. Inline edits silently lose inserted or deleted words

Priority: medium. Locations:

- `assets/web/transcripts.js:1561`, `finishSegmentEditing`.
- `assets/web/transcripts.js:1585`, `extractCorrections`.
- `assets/web/transcripts.js:1640`, `saveCorrections`.

The word-diff helper emits a correction only when an edit group contains both
deleted and inserted words. Pure insertions and deletions disappear from its
output. When every group is discarded, `finishSegmentEditing` returns without
saving or reloading, leaving the edited DOM visible until a later refresh.
Mixed edits can save a replacement while silently discarding another change.

Verified helper reproductions:

| Original | Edited | Generated corrections |
| --- | --- | --- |
| `I like it` | `I really like it` | `[]` |
| `I really like it` | `I like it` | `[]` |
| `I like it` | `I love it` | `like → love` |
| `I like it today` | `I really love it` | `like → really love` only |

For the mixed example, applying the saved correction leaves `today` in the
transcript, contrary to the submitted edit. For the first two, no save request
is made by the caller. The caller behavior was inspected in source; a full browser
interaction was not performed.

Expected: accepted text edits survive reload in full. Unsupported edit forms must
not appear saved or partially succeed without feedback.

Implementation plan:

- [ ] Represent insertion and deletion edits without silently dropping diff groups.
- [ ] Choose contextual replacements or persisted segment edits that preserve the whole submitted text.
- [ ] Retain intentional study-wide correction behavior and avoid empty-pattern global substitutions.
- [ ] Add behavioral regressions for insert-only, delete-only, and mixed edits.
- [ ] Verify the final visible text equals the requested text after a reload.
- [ ] Exercise editing and reload through existing UI tooling; verify failed saves restore or explain state.

Coordinate representation changes with the backend: correction creation requires
nonempty `from` and `to`, and `transcripts.apply_corrections` skips empty values.
Emitting empty correction halves from JavaScript alone will not fix this.

## 2. Chaining updates only one matching correction

Priority: medium. Locations:

- `source/transcripts_server.py:1520`, `api_corrections_add`.
- `source/transcripts.py:1098`, `apply_corrections`.
- `assets/web/transcripts.js:1640`, inline edits submit these global correction pairs.

When a new correction's source matches an existing correction's destination,
the route rewrites the first matching rule and breaks. It neither updates other
matching rules nor adds the submitted rule. Results depend on insertion order,
and some occurrences of the edited word remain unchanged.

Verified API reproduction, starting with an empty corrections list:

1. POST `/transcripts/api/corrections` with `{"from":"teh","to":"the"}`.
2. POST with `{"from":"hte","to":"the"}`.
3. POST with `{"from":"the","to":"they"}`.
4. All three requests return HTTP 200.
5. Stored pairs are only `teh → they` and `hte → the`.
6. Applying the rules to `teh hte the` produces `they the the`.

This also affects editing the displayed correction of `hte`: the UI submits
`the → they`, but the server changes the unrelated first rule, leaving that
displayed occurrence unchanged on reload.

Expected: a correction submitted for a visible word must update that occurrence
consistently, without selecting an arbitrary earlier rule. Preserve the intended
global-correction semantics for other matching occurrences as well.

Implementation plan:

- [ ] Resolve every applicable correction chain deterministically instead of stopping at the first match.
- [ ] Preserve the submitted replacement's intended effect on raw matching text.
- [ ] Cover two misspellings converging on one destination, in both insertion orders.
- [ ] Cover editing the second corrected occurrence and reloading the transcript.
- [ ] Preserve chain reversal, case-insensitive matching, and literal replacement behavior.
- [ ] Verify response handling and persistence when more than one rule changes.

## Verification and handoff

Reproduced on 2026-09-18. The frontend reproduction evaluated the actual
`extractCorrections` function from the checked-out JavaScript in Node. The backend
reproduction used `uv run --no-sync python`, a Flask test client with the real
Transcripts blueprint, and the real `transcripts.apply_corrections` function.
Persistence scheduling was replaced with a no-op and the manifest was seeded
in memory; no project-state writes, transcription, AI calls, or media processing
occurred.

After implementing regressions, run:

```sh
uv run --extra dev pytest -c tests/pytest.ini tests/test_transcripts.py tests/test_transcripts_api.py
```

Follow the existing UI-check procedure for frontend changes and
`agents/skills/check/SKILL.md` before committing. This hunt did not run the full
suite or launch a browser.
