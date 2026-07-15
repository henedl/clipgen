# clipgen-new-thinking-agent — Add an Ollama thinking agent

Both the orchestrator **and** the HTTP routes in `transcripts_server.py` auto-pick up new agents from the `AGENTS` list by key. No orchestrator or route edits needed. The frontend touch is a descriptor entry, not new plumbing.

## Checklist

1. **Config** (`config.py`)
   - Add `OLLAMA_{NAME}_ENABLED: bool` toggle
   - Add `OLLAMA_{NAME}_MODEL: str` default model name

2. **Agent implementation** (`thinking_agents.py`)
   - Write a `run(transcript_entry: dict) -> value | None` callable
     - Returns the value to store in the manifest field, or `None` to skip
     - `transcript_entry` is the dict for one participant from `transcripts_manifest.json`
   - Append an `Agent(...)` entry to the `AGENTS` list
     - Respect topological order: dependencies must appear before dependents
     - Set `depends_on` to the `manifest_field` names of agents this one needs
     - Set `manifest_field` to the key that will be written into the transcript entry
     - Set `on_upstream_change` to `"clear"` (drop the result when an upstream
       dependency regenerates, the default) or `"stale"` (keep it but flag for a
       prompted re-run; only the friction shape carries a `stale` flag today, so a
       new `"stale"` agent also needs a `stale`-flagging path)

3. **Backend routes** (`transcripts_server.py`), **usually nothing**
   - The generic `/api/agent/<key>/<participant>` routes (poll / `regenerate` /
     `stop`) cover every agent in `AGENTS` by key, so a new agent needs **no**
     route edits. The generic poll returns the result under the agent's
     `manifest_field`, plus a `<dep_field>_generating`/`_started_at` block for
     any dependents, driven off `depends_on`.
   - Only add a route for genuinely-unique behavior. Summary is the sole example:
     its SSE token **stream** (`/api/agent/summary/<pid>/stream`) and user-**edit**
     PUT are hand-written because only summary streams tokens / is user-editable.
     A second streaming agent would generalize the `/stream` seam via an `Agent`
     `streams` flag rather than cloning the route.

4. **UI surfacing** (`assets/web/transcripts-agents.js`)
   - Add an `AGENT_DESCRIPTORS` entry (URL base `api/agent/<key>`, poll interval,
     optional timeout) plus the agent's render hooks (`onResult`/`onGenerating`/
     `onEmpty`/`onStale`). The shared `_makeAgentPoll` factory handles the
     poll/staleness/timeout scaffolding. No new poll/stop plumbing.
   - Agent run/stop rows in `transcripts-pills.js` build URLs from the agent key
     (`api/agent/<key>/...`), so a new pill row is data-only too.

5. **Tests** (`tests/test_thinking_agents.py`)
   - Test the `run()` callable with a mock transcript entry
   - Test that the agent appears in `AGENTS` with correct metadata

## No edits needed

- `transcripts_server.py` routes (the generic `/api/agent/<key>/...` routes and
  orchestrator both key off `AGENTS`; the chain auto-advances)
- `ollama_client.py` (pure transport layer)
