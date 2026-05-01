# clipgen-new-thinking-agent — Add an Ollama thinking agent

The orchestrator in `transcripts_server.py` auto-picks up new agents from the `AGENTS` list — no orchestrator edits needed.

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

3. **UI surfacing (if needed)** (`transcripts_server.py`)
   - Extend an existing endpoint's response shape, or add a new endpoint
   - The in-flight check pattern: `_is_generating(participant, agent_key)`

4. **Tests** (`tests/test_thinking_agents.py`)
   - Test the `run()` callable with a mock transcript entry
   - Test that the agent appears in `AGENTS` with correct metadata

## No edits needed

- `transcripts_server.py` orchestrator logic (chain auto-advances)
- `ollama_client.py` (pure transport layer)
