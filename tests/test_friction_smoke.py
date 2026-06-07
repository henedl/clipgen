# -*- coding: utf-8 -*-
"""End-to-end smoke test for the friction pipeline.

Stub Whisper output (segments) flows through the programmatic scorer into the
friction agent's _run callable, with the Ollama call mocked. Verifies the whole
chain produces scored segments, aggregated stats, and parsed moments in the
expected manifest shape.
"""

from unittest.mock import patch

import config
import thinking_agents


_STUB_SEGMENTS = [
    {"id": "P01:0", "start": 0.0, "end": 6.0, "text": "okay let me try clicking here"},
    {"id": "P01:1", "start": 6.0, "end": 12.0, "text": "um where is the settings menu"},
    {
        "id": "P01:2",
        "start": 12.0,
        "end": 18.0,
        "text": "this is broken why won't it open",
    },
    {"id": "P01:3", "start": 18.0, "end": 24.0, "text": "oh wait there it is"},
    {"id": "P01:4", "start": 24.0, "end": 30.0, "text": "can you help me here"},
]

_STUB_MOMENTS = (
    '[{"segment_ids": ["P01:2"], "category": "frustration", '
    '"rationale": "Could not open settings", "score": 0.9},'
    ' {"segment_ids": ["P01:1"], "category": "confusion", '
    '"rationale": "Hunting for the settings menu", "score": 0.6}]'
)


@patch("thinking_agents.ollama_client.generate")
def test_friction_pipeline_end_to_end(mock_generate):
    mock_generate.return_value = _STUB_MOMENTS

    entry = {
        "summary": "Participant explored settings and hit a blocker opening the menu.",
        "segments": _STUB_SEGMENTS,
    }

    agent = thinking_agents.get_agent("friction")
    assert agent is not None
    result = agent["run"](entry, None)

    assert result is not None
    # Programmatic scoring covers every segment, in order.
    assert [s["id"] for s in result["segments"]] == [s["id"] for s in _STUB_SEGMENTS]
    # The hot segments scored above zero; the model received candidates.
    assert any(s["score"] > 0 for s in result["segments"])
    assert mock_generate.called

    # Stats aggregate across all six categories.
    assert set(result["stats"]["by_category"]) == set(config.FRICTION_CATEGORIES)
    assert result["stats"]["total_markers"] > 0
    assert result["stats"]["markers_per_minute"] > 0

    # Moments parsed from the (mocked) model response.
    categories = {m["category"] for m in result["moments"]}
    assert "frustration" in categories
    assert result["stale"] is False
    assert result["model"] == config.OLLAMA_FRICTION_MODEL


@patch("thinking_agents.ollama_client.generate")
def test_friction_pipeline_survives_unparseable_model_output(mock_generate):
    # The programmatic layer must still produce scores/stats even when the LLM
    # response yields no usable moments.
    mock_generate.return_value = "I could not find any friction, sorry."

    entry = {"summary": "A session summary.", "segments": _STUB_SEGMENTS}
    agent = thinking_agents.get_agent("friction")
    assert agent is not None
    result = agent["run"](entry, None)

    assert result is not None
    assert result["moments"] == []
    assert len(result["segments"]) == len(_STUB_SEGMENTS)
    assert result["stats"]["total_markers"] > 0
