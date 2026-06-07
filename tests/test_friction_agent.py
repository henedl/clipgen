# -*- coding: utf-8 -*-
"""Tests for the friction thinking agent (thinking_agents friction helpers).

Covers defensive JSON extraction/parsing, candidate formatting, the Ollama-
backed moment finder, and the _run_friction assembler. Transport is mocked.
"""

import threading
from unittest.mock import patch

import config
import thinking_agents


class TestExtractJsonArray:
    def test_plain_array(self):
        out = thinking_agents._extract_json_array('[{"a": 1}]')
        assert out == [{"a": 1}]

    def test_prose_wrapped(self):
        text = 'Here are the moments:\n[{"category": "x"}]\nThanks!'
        assert thinking_agents._extract_json_array(text) == [{"category": "x"}]

    def test_markdown_fenced(self):
        text = '```json\n[{"category": "frustration"}]\n```'
        assert thinking_agents._extract_json_array(text) == [
            {"category": "frustration"}
        ]

    def test_strips_think_block(self):
        text = '<think>let me reason about this</think>\n[{"ok": true}]'
        assert thinking_agents._extract_json_array(text) == [{"ok": True}]

    def test_returns_empty_on_garbage(self):
        assert thinking_agents._extract_json_array("no json here") == []
        assert thinking_agents._extract_json_array("") == []

    def test_ignores_non_array_json(self):
        # An object, not an array → keep scanning, find the array later.
        text = '{"not": "an array"} then [{"yes": 1}]'
        assert thinking_agents._extract_json_array(text) == [{"yes": 1}]


class TestParseFrictionResponse:
    def test_normalizes_and_ignores_extra_fields(self):
        response = (
            '[{"segment_ids": ["P01:7", "P01:8"], "category": "frustration", '
            '"rationale": "tried repeatedly", "score": 0.85, "extra": "ignored"}]'
        )
        moments = thinking_agents._parse_friction_response(response)
        assert len(moments) == 1
        m = moments[0]
        assert m["segment_ids"] == ["P01:7", "P01:8"]
        assert m["category"] == "frustration"
        assert m["rationale"] == "tried repeatedly"
        assert m["score"] == 0.85
        assert "extra" not in m

    def test_coerces_single_segment_id_string(self):
        moments = thinking_agents._parse_friction_response(
            '[{"segment_ids": "P01:3", "category": "confusion", "score": 0.4}]'
        )
        assert moments[0]["segment_ids"] == ["P01:3"]

    def test_clamps_score(self):
        moments = thinking_agents._parse_friction_response(
            '[{"segment_ids": ["a"], "score": 5.0}, '
            '{"segment_ids": ["b"], "score": -2.0}, '
            '{"segment_ids": ["c"], "score": "bad"}]'
        )
        assert moments[0]["score"] == 1.0
        assert moments[1]["score"] == 0.0
        assert moments[2]["score"] == 0.0

    def test_normalizes_category_key(self):
        moments = thinking_agents._parse_friction_response(
            '[{"segment_ids": ["a"], "category": "Self-Correction"}]'
        )
        assert moments[0]["category"] == "self_correction"

    def test_drops_moments_without_segment_ids(self):
        moments = thinking_agents._parse_friction_response(
            '[{"category": "confusion"}, {"segment_ids": [], "category": "x"}, '
            '{"segment_ids": ["a"], "category": "confusion"}]'
        )
        assert len(moments) == 1
        assert moments[0]["segment_ids"] == ["a"]


class TestFormatFrictionCandidates:
    def _segments(self):
        return [
            {"id": "P01:0", "start": 0, "text": "intro words"},
            {"id": "P01:1", "start": 30, "text": "um where is it"},
            {"id": "P01:2", "start": 60, "text": "found it"},
            {"id": "P01:3", "start": 90, "text": "unrelated"},
        ]

    def test_includes_candidate_with_context(self):
        block = thinking_agents._format_friction_candidates(
            self._segments(), [{"id": "P01:1"}]
        )
        # Candidate plus ±1 neighbor, ordered, with timestamps.
        assert "[P01:0]" in block
        assert "[P01:1]" in block
        assert "[P01:2]" in block
        assert "[P01:3]" not in block
        assert "(0:30)" in block

    def test_merges_adjacent_candidates_without_duplicates(self):
        block = thinking_agents._format_friction_candidates(
            self._segments(), [{"id": "P01:1"}, {"id": "P01:2"}]
        )
        assert block.count("[P01:1]") == 1
        assert block.count("[P01:2]") == 1


class TestFindFrictionMoments:
    def test_empty_candidates_skips_model(self):
        with patch("thinking_agents.ollama_client.generate") as mock_gen:
            result = thinking_agents.find_friction_moments("summary", [], [])
        assert result == []
        mock_gen.assert_not_called()

    @patch("thinking_agents.ollama_client.generate")
    def test_builds_prompt_from_summary_and_candidates(self, mock_gen):
        mock_gen.return_value = "[]"
        segments = [
            {"id": "P01:0", "start": 0, "text": "before"},
            {"id": "P01:1", "start": 5, "text": "um where is it"},
        ]
        thinking_agents.find_friction_moments(
            "The user struggled.", segments, [{"id": "P01:1"}]
        )
        prompt = mock_gen.call_args[0][0]
        assert "The user struggled." in prompt
        assert "[P01:1]" in prompt
        assert mock_gen.call_args[1]["think"] is False

    @patch("thinking_agents.ollama_client.generate")
    def test_parses_and_caps_moments(self, mock_gen, monkeypatch):
        monkeypatch.setattr(config, "FRICTION_MOMENT_LIMIT", 2)
        mock_gen.return_value = (
            '[{"segment_ids": ["P01:1"], "category": "confusion", "score": 0.5},'
            ' {"segment_ids": ["P01:1"], "category": "hesitation", "score": 0.4},'
            ' {"segment_ids": ["P01:1"], "category": "surprise", "score": 0.3}]'
        )
        segments = [{"id": "P01:1", "start": 5, "text": "um where is it"}]
        result = thinking_agents.find_friction_moments(
            "summary", segments, [{"id": "P01:1"}]
        )
        assert result is not None
        assert len(result) == 2

    @patch("thinking_agents.ollama_client.generate")
    def test_returns_none_when_generate_fails(self, mock_gen):
        # None (model failure) is distinct from [] (ran, no moments) so the
        # caller can surface the failure instead of pretending it computed.
        mock_gen.return_value = None
        segments = [{"id": "P01:1", "start": 5, "text": "um where is it"}]
        result = thinking_agents.find_friction_moments(
            "summary", segments, [{"id": "P01:1"}]
        )
        assert result is None

    @patch("thinking_agents.ollama_client.generate")
    def test_returns_empty_list_when_model_runs_but_finds_nothing(self, mock_gen):
        mock_gen.return_value = "[]"
        segments = [{"id": "P01:1", "start": 5, "text": "um where is it"}]
        result = thinking_agents.find_friction_moments(
            "summary", segments, [{"id": "P01:1"}]
        )
        assert result == []

    def test_no_candidates_returns_empty_not_none(self):
        assert thinking_agents.find_friction_moments("summary", [], []) == []

    def test_friction_model_defaults_to_summary_model(self, monkeypatch):
        monkeypatch.setattr(config, "OLLAMA_FRICTION_MODEL", "")
        monkeypatch.setattr(config, "OLLAMA_SUMMARY_MODEL", "llama3.2")
        assert thinking_agents.friction_model() == "llama3.2"
        monkeypatch.setattr(config, "OLLAMA_FRICTION_MODEL", "tiny:1b")
        assert thinking_agents.friction_model() == "tiny:1b"

    @patch("thinking_agents.ollama_client.generate")
    def test_passes_cancel_event_and_model(self, mock_gen):
        mock_gen.return_value = "[]"
        evt = threading.Event()
        segments = [{"id": "P01:1", "start": 5, "text": "um where is it"}]
        thinking_agents.find_friction_moments(
            "summary", segments, [{"id": "P01:1"}], model="tiny:1b", cancel_event=evt
        )
        assert mock_gen.call_args[1]["cancel_event"] is evt
        assert mock_gen.call_args[1]["model"] == "tiny:1b"


class TestRunFriction:
    def _entry(self):
        return {
            "summary": "User struggled to find the save button.",
            "segments": [
                {
                    "id": "P01:0",
                    "start": 0,
                    "end": 5,
                    "text": "um where is the save button",
                },
                {
                    "id": "P01:1",
                    "start": 5,
                    "end": 10,
                    "text": "this is broken why won't it work",
                },
                {"id": "P01:2", "start": 10, "end": 15, "text": "okay that works now"},
            ],
        }

    def test_returns_none_without_segments(self):
        assert thinking_agents._run_friction({"summary": "s"}, None) is None

    def test_returns_none_without_summary(self):
        entry = {"segments": [{"id": "P01:0", "text": "um"}]}
        assert thinking_agents._run_friction(entry, None) is None

    @patch("thinking_agents.ollama_client.generate")
    def test_cancel_before_llm_returns_none(self, mock_gen):
        evt = threading.Event()
        evt.set()
        result = thinking_agents._run_friction(self._entry(), evt)
        assert result is None
        mock_gen.assert_not_called()

    @patch("thinking_agents.ollama_client.generate")
    def test_success_returns_full_dict(self, mock_gen):
        mock_gen.return_value = (
            '[{"segment_ids": ["P01:1"], "category": "frustration", '
            '"rationale": "save button broken", "score": 0.8}]'
        )
        result = thinking_agents._run_friction(self._entry(), None)
        assert result is not None
        assert {
            "segments",
            "moments",
            "stats",
            "computed_at",
            "model",
            "llm_ok",
            "stale",
        } <= set(result)
        assert result["stale"] is False
        assert result["llm_ok"] is True
        assert result["model"] == thinking_agents.friction_model()
        assert len(result["segments"]) == 3
        assert result["moments"][0]["category"] == "frustration"
        assert "by_category" in result["stats"]

    @patch("thinking_agents.ollama_client.generate")
    def test_llm_failure_persists_scores_with_llm_ok_false(self, mock_gen):
        # Model unavailable (e.g. wrong model → HTTP 404) returns None from
        # generate; the friction dict must still carry programmatic scores/stats
        # but flag llm_ok=False so the UI shows the failure honestly.
        mock_gen.return_value = None
        result = thinking_agents._run_friction(self._entry(), None)
        assert result is not None
        assert result["llm_ok"] is False
        assert result["moments"] == []
        assert len(result["segments"]) == 3
        assert "by_category" in result["stats"]

    @patch("thinking_agents.ollama_client.generate")
    def test_run_friction_uses_resolved_model(self, mock_gen, monkeypatch):
        monkeypatch.setattr(config, "OLLAMA_FRICTION_MODEL", "")
        monkeypatch.setattr(config, "OLLAMA_SUMMARY_MODEL", "llama3.2")
        mock_gen.return_value = "[]"
        result = thinking_agents._run_friction(self._entry(), None)
        assert result is not None
        assert result["model"] == "llama3.2"
        assert mock_gen.call_args[1]["model"] == "llama3.2"

    @patch("thinking_agents.ollama_client.generate")
    def test_registry_run_callable_dispatches(self, mock_gen):
        mock_gen.return_value = "[]"
        agent = thinking_agents.get_agent("friction")
        assert agent is not None
        assert agent["depends_on"] == ["summary"]
        result = agent["run"](self._entry(), None)
        assert result is not None
        assert result["moments"] == []
