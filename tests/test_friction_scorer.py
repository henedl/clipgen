"""Tests for the programmatic friction scorer (friction.py).

Pure deterministic engine — no Ollama, no I/O. Covers phrase matching per
category, the score formula, candidate selection, and stats aggregation.
"""

import config
import friction


class TestCategoryConsistency:
    def test_pattern_weight_and_config_keys_match(self):
        """The scorer's pattern/weight tables and config.FRICTION_CATEGORIES
        must share the exact same category key set (single source of truth)."""
        assert set(friction.FRICTION_PATTERNS) == set(friction.CATEGORY_WEIGHTS)
        assert set(friction.FRICTION_PATTERNS) == set(config.FRICTION_CATEGORIES)

    def test_category_order_matches_pattern_order(self):
        assert friction.CATEGORY_ORDER == tuple(friction.FRICTION_PATTERNS.keys())


class TestScoreSegments:
    def test_clean_segment_scores_zero(self):
        scored = friction.score_segments(
            [{"id": "P01:0", "text": "the quick brown fox jumps over"}]
        )
        assert scored[0]["score"] == 0.0
        assert scored[0]["categories"] == []
        assert scored[0]["markers"] == []
        assert scored[0]["counts"] == {}

    def test_hesitation_marker_matches(self):
        scored = friction.score_segments([{"id": "P01:0", "text": "um yes"}])
        assert "hesitation" in scored[0]["categories"]
        assert "um" in scored[0]["markers"]
        assert scored[0]["counts"]["hesitation"] == 1

    def test_confusion_phrase_matches(self):
        scored = friction.score_segments(
            [{"id": "P01:0", "text": "where is the save button"}]
        )
        assert "confusion" in scored[0]["categories"]
        assert "where is" in scored[0]["markers"]

    def test_frustration_phrase_matches(self):
        scored = friction.score_segments(
            [{"id": "P01:0", "text": "ugh this is broken"}]
        )
        # "ugh" + "this is broken" → two frustration markers.
        assert "frustration" in scored[0]["categories"]
        assert scored[0]["counts"]["frustration"] == 2

    def test_word_doubling_matches_a_real_double(self):
        scored = friction.score_segments([{"id": "P01:0", "text": "I want want it"}])
        assert "hesitation" in scored[0]["categories"]
        assert "want want" in scored[0]["markers"]

    def test_word_doubling_no_false_positive(self):
        # No adjacent repeated word, and no other hesitation phrase present.
        scored = friction.score_segments(
            [{"id": "P01:0", "text": "the cat sat on a mat"}]
        )
        assert scored[0]["score"] == 0.0

    def test_score_formula(self):
        # "um, where is the menu?" → 5 words.
        # hesitation "um" (w=1.0) + confusion "where is" (w=1.5) = 2.5 / 5 = 0.5
        scored = friction.score_segments(
            [{"id": "P01:0", "text": "um, where is the menu?"}]
        )
        assert scored[0]["score"] == 0.5
        assert scored[0]["categories"] == ["hesitation", "confusion"]

    def test_score_clamped_to_one(self):
        # A one-word frustration interjection would score 2.0 raw → clamped.
        scored = friction.score_segments([{"id": "P01:0", "text": "ugh"}])
        assert scored[0]["score"] == 1.0

    def test_id_falls_back_to_index(self):
        scored = friction.score_segments([{"text": "um"}])
        assert scored[0]["id"] == "0"

    def test_categories_ordered_by_config_order(self):
        # frustration appears before help_seeking in CATEGORY_ORDER regardless of
        # text order.
        scored = friction.score_segments([{"id": "P01:0", "text": "can you help, ugh"}])
        cats = scored[0]["categories"]
        assert cats.index("frustration") < cats.index("help_seeking")


class TestSelectCandidates:
    def test_returns_top_n_by_score(self):
        scored = [
            {"id": "a", "score": 0.1, "counts": {"hesitation": 1}},
            {"id": "b", "score": 0.9, "counts": {"frustration": 1}},
            {"id": "c", "score": 0.0, "counts": {}},
            {"id": "d", "score": 0.5, "counts": {"confusion": 2}},
        ]
        result = friction.select_candidates(scored, n=2)
        assert [r["id"] for r in result] == ["b", "d"]

    def test_excludes_zero_score(self):
        scored = [
            {"id": "a", "score": 0.0, "counts": {}},
            {"id": "b", "score": 0.3, "counts": {"hesitation": 1}},
        ]
        result = friction.select_candidates(scored, n=10)
        assert [r["id"] for r in result] == ["b"]

    def test_ties_broken_by_marker_count(self):
        scored = [
            {"id": "a", "score": 0.5, "counts": {"hesitation": 1}},
            {"id": "b", "score": 0.5, "counts": {"hesitation": 3}},
        ]
        result = friction.select_candidates(scored, n=1)
        assert result[0]["id"] == "b"


class TestComputeStats:
    def test_aggregates_counts_and_rate(self):
        scored = [
            {"id": "a", "score": 0.5, "counts": {"hesitation": 2}},
            {"id": "b", "score": 0.3, "counts": {"confusion": 1, "frustration": 3}},
        ]
        stats = friction.compute_stats(scored, duration_seconds=120.0)
        assert stats["total_markers"] == 6
        assert stats["markers_per_minute"] == 3.0
        assert stats["by_category"]["hesitation"] == 2
        assert stats["by_category"]["confusion"] == 1
        assert stats["by_category"]["frustration"] == 3
        # All categories present, zeros included.
        assert stats["by_category"]["help_seeking"] == 0
        assert set(stats["by_category"]) == set(friction.CATEGORY_ORDER)

    def test_zero_duration_gives_zero_rate(self):
        scored = [{"id": "a", "score": 0.5, "counts": {"hesitation": 2}}]
        stats = friction.compute_stats(scored, duration_seconds=0.0)
        assert stats["markers_per_minute"] == 0.0
        assert stats["total_markers"] == 2
