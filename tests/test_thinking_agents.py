"""Tests for thinking_agents module.

Covers the built-in summary and citation agents plus registry invariants.
Transport-layer tests live in tests/test_ollama_client.py.
"""

import threading
from unittest.mock import patch

import thinking_agents


class TestRegistry:
    def test_every_agent_has_required_fields(self):
        required = {
            "key",
            "enabled_config_key",
            "manifest_field",
            "depends_on",
            "thread_name_prefix",
            "run",
        }
        for agent in thinking_agents.AGENTS:
            assert required.issubset(agent.keys()), (
                f"Agent {agent.get('key')} missing fields: {required - agent.keys()}"
            )

    def test_dependencies_resolve_to_registered_agents(self):
        keys = {a["key"] for a in thinking_agents.AGENTS}
        for agent in thinking_agents.AGENTS:
            for dep in agent["depends_on"]:
                assert dep in keys, (
                    f"Agent {agent['key']} depends on unknown agent {dep!r}"
                )

    def test_dependencies_listed_before_dependents(self):
        """AGENTS must be topologically ordered so the orchestrator's
        linear scan produces a valid run order."""
        seen: set[str] = set()
        for agent in thinking_agents.AGENTS:
            for dep in agent["depends_on"]:
                assert dep in seen, (
                    f"Agent {agent['key']} appears before its dependency {dep!r}"
                )
            seen.add(agent["key"])

    def test_get_agent_returns_registered_agent(self):
        assert thinking_agents.get_agent("summary") is not None
        assert thinking_agents.get_agent("citations") is not None

    def test_get_agent_returns_none_for_unknown(self):
        assert thinking_agents.get_agent("nonexistent") is None

    def test_builtin_agents_present(self):
        keys = {a["key"] for a in thinking_agents.AGENTS}
        assert "summary" in keys
        assert "citations" in keys

    def test_resolve_model_uses_agent_model(self, monkeypatch):
        import config

        monkeypatch.setattr(config, "OLLAMA_SUMMARY_MODEL", "qwen3.5:9b")
        monkeypatch.setattr(config, "OLLAMA_FRICTION_MODEL", "gemma4:latest")
        summary = thinking_agents.get_agent("summary")
        friction = thinking_agents.get_agent("friction")
        assert summary is not None and friction is not None
        assert thinking_agents.resolve_model(summary) == "qwen3.5:9b"
        assert thinking_agents.resolve_model(friction) == "gemma4:latest"

    def test_resolve_model_blank_friction_inherits_summary(self, monkeypatch):
        import config

        monkeypatch.setattr(config, "OLLAMA_SUMMARY_MODEL", "llama3.1:8b")
        monkeypatch.setattr(config, "OLLAMA_FRICTION_MODEL", "")
        friction = thinking_agents.get_agent("friction")
        assert friction is not None
        assert thinking_agents.resolve_model(friction) == "llama3.1:8b"


class TestSummarizeTranscript:
    def test_returns_none_for_empty_segments(self):
        result = thinking_agents.summarize_transcript([])
        assert result is None

    def test_returns_none_for_very_short_text(self):
        result = thinking_agents.summarize_transcript([{"text": "Hi"}])
        assert result is None

    @patch("thinking_agents.ollama_client.generate")
    def test_concatenates_segment_text(self, mock_generate):
        mock_generate.return_value = "A summary."
        segments = [
            {"text": "First segment text here."},
            {"text": "Second segment text here."},
            {"text": "Third segment more text."},
        ]
        thinking_agents.summarize_transcript(segments)

        prompt = mock_generate.call_args[0][0]
        assert "First segment text here." in prompt
        assert "Second segment text here." in prompt
        assert "Third segment more text." in prompt

    def test_uses_configurable_prompt(self, monkeypatch):
        """The summary prompt is read from config at call time, so an edit in
        Settings → Summaries takes effect on the next run."""
        import config

        monkeypatch.setattr(config, "OLLAMA_SUMMARY_PROMPT", "CUSTOM-MARKER\n{text}")
        with patch("thinking_agents.ollama_client.generate") as mock_generate:
            mock_generate.return_value = "ok"
            thinking_agents.summarize_transcript(
                [{"text": "A sufficiently long segment of text for the length check."}]
            )
        prompt = mock_generate.call_args[0][0]
        assert prompt.startswith("CUSTOM-MARKER")
        assert "A sufficiently long segment of text" in prompt

    @patch("thinking_agents.ollama_client.generate")
    def test_returns_generate_result(self, mock_generate):
        mock_generate.return_value = "This is a summary.\n- Point one\n- Point two"
        segments = [
            {
                "text": "A sufficiently long segment of text for the minimum length check."
            },
        ]
        result = thinking_agents.summarize_transcript(segments)
        assert result == "This is a summary.\n- Point one\n- Point two"

    @patch("thinking_agents.ollama_client.generate")
    def test_returns_none_when_generate_fails(self, mock_generate):
        mock_generate.return_value = None
        segments = [
            {
                "text": "A sufficiently long segment of text for the minimum length check."
            },
        ]
        result = thinking_agents.summarize_transcript(segments)
        assert result is None

    @patch("thinking_agents.ollama_client.generate")
    def test_strips_think_block_from_result(self, mock_generate):
        # Transport now returns raw text; the summary agent owns <think> stripping.
        mock_generate.return_value = "<think>reasoning</think>\n\nActual summary."
        segments = [
            {
                "text": "A sufficiently long segment of text for the minimum length check."
            },
        ]
        result = thinking_agents.summarize_transcript(segments)
        assert result == "Actual summary."

    @patch("thinking_agents.ollama_client.generate")
    def test_returns_none_when_only_think_block(self, mock_generate):
        mock_generate.return_value = "<think>Thinking but producing nothing.</think>"
        segments = [
            {
                "text": "A sufficiently long segment of text for the minimum length check."
            },
        ]
        result = thinking_agents.summarize_transcript(segments)
        assert result is None

    @patch("thinking_agents.ollama_client.generate")
    def test_passes_model_override(self, mock_generate):
        mock_generate.return_value = "ok"
        segments = [
            {
                "text": "A sufficiently long segment of text for the minimum length check."
            },
        ]
        thinking_agents.summarize_transcript(segments, model="llama3.1:8b")
        assert mock_generate.call_args[1]["model"] == "llama3.1:8b"

    @patch("thinking_agents.ollama_client.generate")
    def test_uses_default_model_when_no_override(self, mock_generate):
        mock_generate.return_value = "ok"
        segments = [
            {
                "text": "A sufficiently long segment of text for the minimum length check."
            },
        ]
        thinking_agents.summarize_transcript(segments)
        assert mock_generate.call_args[1]["model"] == "qwen3.5:9b"

    @patch("thinking_agents.ollama_client.generate")
    def test_passes_cancel_event_to_generate(self, mock_generate):
        mock_generate.return_value = "ok"
        evt = threading.Event()
        segments = [
            {
                "text": "A sufficiently long segment of text for the minimum length check."
            },
        ]
        thinking_agents.summarize_transcript(segments, cancel_event=evt)
        assert mock_generate.call_args[1]["cancel_event"] is evt

    @patch("thinking_agents.ollama_client.generate")
    def test_disables_thinking(self, mock_generate):
        # Summaries run with think=False (like citations/friction): a reasoning
        # model would otherwise stall in a silent think phase with no response
        # text to stream and pure added latency.
        mock_generate.return_value = "ok"
        segments = [
            {
                "text": "A sufficiently long segment of text for the minimum length check."
            },
        ]
        thinking_agents.summarize_transcript(segments)
        assert mock_generate.call_args[1]["think"] is False

    @patch("thinking_agents.ollama_client.generate")
    def test_forwards_on_token_to_generate(self, mock_generate):
        mock_generate.return_value = "ok"
        sink = lambda _tok: None
        segments = [
            {
                "text": "A sufficiently long segment of text for the minimum length check."
            },
        ]
        thinking_agents.summarize_transcript(segments, on_token=sink)
        assert mock_generate.call_args[1]["on_token"] is sink

    @patch("thinking_agents.ollama_client.generate")
    def test_run_summary_forwards_on_token(self, mock_generate):
        mock_generate.return_value = "ok"
        sink = lambda _tok: None
        entry = {
            "segments": [
                {"text": "A sufficiently long segment of text for the length check."}
            ]
        }
        thinking_agents._run_summary(entry, None, sink)
        assert mock_generate.call_args[1]["on_token"] is sink


class TestSplitSummarySentences:
    def test_splits_paragraph_on_sentence_boundaries(self):
        text = "First sentence. Second sentence. Third sentence."
        result = thinking_agents._split_summary_sentences(text)
        assert result == ["First sentence.", "Second sentence.", "Third sentence."]

    def test_treats_bullets_as_individual_sentences(self):
        text = "Overview.\n- Bullet one\n- Bullet two\n* Star bullet"
        result = thinking_agents._split_summary_sentences(text)
        assert result == ["Overview.", "Bullet one", "Bullet two", "Star bullet"]

    def test_skips_empty_lines(self):
        text = "First.\n\n- Bullet\n\n"
        result = thinking_agents._split_summary_sentences(text)
        assert result == ["First.", "Bullet"]

    def test_handles_exclamation_and_question_marks(self):
        text = "Was it good? It was great! Done."
        result = thinking_agents._split_summary_sentences(text)
        assert result == ["Was it good?", "It was great!", "Done."]

    def test_returns_empty_for_empty_input(self):
        assert thinking_agents._split_summary_sentences("") == []
        assert thinking_agents._split_summary_sentences("  \n\n  ") == []


class TestFormatSegmentChunk:
    def test_formats_segments_as_timestamped_lines(self):
        segments = [
            {"start": 0, "end": 5, "text": "Hello"},
            {"start": 65, "end": 70, "text": "World"},
        ]
        result = thinking_agents._format_segment_chunk(segments)
        assert result == "[0:00] Hello\n[1:05] World"

    def test_skips_segments_with_empty_text(self):
        segments = [
            {"start": 0, "end": 5, "text": "Hello"},
            {"start": 10, "end": 15, "text": ""},
            {"start": 20, "end": 25, "text": "End"},
        ]
        result = thinking_agents._format_segment_chunk(segments)
        assert result == "[0:00] Hello\n[0:20] End"


class TestParseCitationResponse:
    def _make_segments(self, starts):
        return [{"start": s, "end": s + 5, "text": f"seg at {s}"} for s in starts]

    def test_parses_basic_format(self):
        segments = self._make_segments([45, 62, 120])
        response = "1: 0:45, 2:00\n2: NONE\n3: 0:45"
        result = thinking_agents._parse_citation_response(response, segments)
        # Claim 1 (index 0) matches segments at 45 and 120
        assert 0 in result
        assert len(result[0]) == 2
        assert result[0][0]["segment_index"] == 0  # start=45 (0:45)
        assert result[0][1]["segment_index"] == 2  # start=120 (2:00)
        # Claim 2 (index 1) is NONE
        assert 1 not in result
        # Claim 3 (index 2) matches segment at 45
        assert 2 in result
        assert result[2][0]["segment_index"] == 0

    def test_handles_bracketed_timestamps(self):
        segments = self._make_segments([45])
        response = "1: [0:45]"
        result = thinking_agents._parse_citation_response(response, segments)
        assert 0 in result
        assert result[0][0]["start"] == 45

    def test_handles_hms_format(self):
        segments = self._make_segments([3661])
        response = "1: 1:01:01"
        result = thinking_agents._parse_citation_response(response, segments)
        assert 0 in result
        assert result[0][0]["segment_index"] == 0

    def test_ignores_unparseable_lines(self):
        segments = self._make_segments([10])
        response = "This is gibberish\n1: 0:10\nmore nonsense"
        result = thinking_agents._parse_citation_response(response, segments)
        assert 0 in result
        assert result[0][0]["start"] == 10

    def test_returns_empty_for_all_none(self):
        segments = self._make_segments([10, 20])
        response = "1: NONE\n2: NONE"
        result = thinking_agents._parse_citation_response(response, segments)
        assert result == {}

    def test_rejects_timestamps_beyond_tolerance(self):
        segments = self._make_segments([100])
        response = "1: 0:10"  # 10s is 90s away from 100s
        result = thinking_agents._parse_citation_response(response, segments)
        assert result == {}

    def test_rejects_timestamps_with_seconds_over_59(self):
        # Timestamps route through the strict utils.timestamp_to_seconds, so a
        # malformed seconds field (>= 60) is dropped rather than parsed as 75s.
        segments = self._make_segments([75])
        result = thinking_agents._parse_citation_response("1: 0:75", segments)
        assert result == {}

    def test_handles_unsorted_segments(self):
        # Segments supplied out of start order; the binary search must still
        # resolve to the correct *original* segment index.
        segments = self._make_segments([120, 45, 62])
        result = thinking_agents._parse_citation_response("1: 0:45", segments)
        assert 0 in result
        assert result[0][0]["segment_index"] == 1  # original index of start=45
        assert result[0][0]["start"] == 45

    def test_picks_closest_neighbour(self):
        segments = self._make_segments([10, 20])
        # 0:12 is nearer the 10s segment; 0:18 is nearer the 20s segment.
        near_low = thinking_agents._parse_citation_response("1: 0:12", segments)
        near_high = thinking_agents._parse_citation_response("1: 0:18", segments)
        assert near_low[0][0]["segment_index"] == 0
        assert near_high[0][0]["segment_index"] == 1


class TestFindCitations:
    def test_returns_none_for_empty_summary(self):
        segments = [{"start": 0, "end": 5, "text": "Hello"}]
        assert thinking_agents.find_citations("", segments) is None

    def test_returns_none_for_empty_segments(self):
        assert thinking_agents.find_citations("A summary.", []) is None

    @patch("thinking_agents.ollama_client.generate")
    def test_uses_default_model(self, mock_generate):
        mock_generate.return_value = "1: NONE"
        segments = [{"start": 0, "end": 5, "text": "Some text here"}]
        thinking_agents.find_citations("A summary sentence.", segments)
        assert mock_generate.call_args[1]["model"] == "qwen3.5:9b"

    @patch("thinking_agents.ollama_client.generate")
    def test_returns_citations_on_success(self, mock_generate):
        mock_generate.return_value = "1: 0:00\n2: 0:10"
        segments = [
            {"start": 0, "end": 5, "text": "First part"},
            {"start": 10, "end": 15, "text": "Second part"},
        ]
        summary = "First claim.\n- Bullet claim"
        result = thinking_agents.find_citations(summary, segments)
        assert result is not None
        assert len(result) == 2
        assert result[0]["sentence"] == "First claim."
        assert result[0]["refs"][0]["start"] == 0
        assert result[1]["sentence"] == "Bullet claim"
        assert result[1]["refs"][0]["start"] == 10

    @patch("thinking_agents.ollama_client.generate")
    def test_returns_empty_refs_when_generate_fails(self, mock_generate):
        mock_generate.return_value = None
        segments = [{"start": 0, "end": 5, "text": "Some text"}]
        result = thinking_agents.find_citations("A sentence.", segments)
        # Returns citations list but with empty refs
        assert result is not None
        assert len(result) == 1
        assert result[0]["refs"] == []

    @patch("thinking_agents.ollama_client.generate")
    def test_multiple_refs_per_claim(self, mock_generate):
        mock_generate.return_value = "1: 0:00, 0:10"
        segments = [
            {"start": 0, "end": 5, "text": "Part A"},
            {"start": 10, "end": 15, "text": "Part B"},
        ]
        result = thinking_agents.find_citations("A claim.", segments)
        assert result is not None
        assert len(result[0]["refs"]) == 2
        assert result[0]["refs"][0]["start"] == 0
        assert result[0]["refs"][1]["start"] == 10

    @patch("thinking_agents.ollama_client.generate")
    def test_passes_cancel_event_to_generate(self, mock_generate):
        mock_generate.return_value = "1: NONE"
        evt = threading.Event()
        segments = [{"start": 0, "end": 5, "text": "Some text here"}]
        thinking_agents.find_citations("A claim.", segments, cancel_event=evt)
        assert mock_generate.call_args[1]["cancel_event"] is evt


class TestAgentRunCallables:
    """Smoke tests that the AGENTS registry's run callables dispatch to the
    right underlying helpers."""

    @patch("thinking_agents.summarize_transcript")
    def test_summary_agent_run_invokes_summarize(self, mock_sum):
        mock_sum.return_value = "A summary"
        agent = thinking_agents.get_agent("summary")
        assert agent is not None
        entry = {"segments": [{"start": 0, "end": 1, "text": "hi"}]}
        result = agent["run"](entry, None)
        assert result == "A summary"
        mock_sum.assert_called_once()
        assert mock_sum.call_args[0][0] == entry["segments"]

    @patch("thinking_agents.summarize_transcript")
    def test_summary_agent_forwards_cancel_event(self, mock_sum):
        mock_sum.return_value = "A summary"
        agent = thinking_agents.get_agent("summary")
        assert agent is not None
        evt = threading.Event()
        entry = {"segments": [{"start": 0, "end": 1, "text": "hi"}]}
        agent["run"](entry, evt)
        assert mock_sum.call_args[1]["cancel_event"] is evt

    @patch("thinking_agents.find_citations")
    def test_citations_agent_run_invokes_find_citations(self, mock_find):
        mock_find.return_value = [{"sentence": "s", "refs": []}]
        agent = thinking_agents.get_agent("citations")
        assert agent is not None
        entry = {
            "summary": "A summary.",
            "segments": [{"start": 0, "end": 1, "text": "hi"}],
        }
        result = agent["run"](entry, None)
        assert result == [{"sentence": "s", "refs": []}]
        mock_find.assert_called_once_with(
            entry["summary"], entry["segments"], cancel_event=None
        )

    @patch("thinking_agents.find_citations")
    def test_citations_agent_forwards_cancel_event(self, mock_find):
        mock_find.return_value = [{"sentence": "s", "refs": []}]
        agent = thinking_agents.get_agent("citations")
        assert agent is not None
        evt = threading.Event()
        entry = {
            "summary": "A summary.",
            "segments": [{"start": 0, "end": 1, "text": "hi"}],
        }
        agent["run"](entry, evt)
        assert mock_find.call_args[1]["cancel_event"] is evt
