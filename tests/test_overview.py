"""Tests for overview: feature builders and the /overview/ blueprint.

Builder tests use synthetic manifest dicts (mirroring tests/test_data_export.py);
route smokes mount the blueprint on a bare Flask app with a tmp output dir
(mirroring tests/test_workflows_api.py).
"""

import json
import math

import pytest

import config
import friction
import overview


# ---- Fixtures -------------------------------------------------------------


def _obs_row(participant, category="nav", severity="High", timestamps=1):
    return {
        "participant": participant,
        "category": category,
        "severity": severity,
        "timestamps": timestamps,
    }


def _ss_row(participant="P01", detector="change", time_in=10.0, time_out=20.0, **kw):
    row = {
        "id": f"ev_{participant}_{detector}_{time_in}",
        "participant": participant,
        "detector": detector,
        "event_type": detector,
        "time_in": time_in,
        "time_out": time_out,
        "duration": time_out - time_in,
        "confidence": 0.8,
        "excluded": False,
        "navigational": False,
        "task_id": "ss_x",
    }
    row.update(kw)
    return row


def _segments(participant, texts, seconds_each=30.0):
    segs = []
    for i, text in enumerate(texts):
        segs.append(
            {
                "id": f"{participant}:{i}",
                "start": i * seconds_each,
                "end": (i + 1) * seconds_each,
                "text": text,
            }
        )
    return segs


# ---- Observation features ---------------------------------------------------


def test_observation_shares_sum_to_one():
    # Timestamp-weighted: a two-timestamp nav cell counts twice.
    rows = [
        _obs_row("P01", category="nav", timestamps=2),
        _obs_row("P01", category="search", severity="Low"),
        _obs_row("P02", category="search"),
    ]
    columns, values = overview.build_observation_features(rows)
    keys = {c["key"] for c in columns}
    assert "obs_cat_nav" in keys and "obs_cat_search" in keys
    assert "obs_sev_high" in keys and "obs_sev_low" in keys

    p01 = values["P01"]
    cat_shares = [p01[k] for k in p01 if k.startswith("obs_cat_")]
    assert sum(cat_shares) == pytest.approx(1.0, abs=1e-3)
    assert p01["obs_cat_nav"] == pytest.approx(2 / 3, abs=1e-3)
    assert p01["obs_total"] == 3.0
    # P02 never saw "nav" but still gets the column (cohort union), as 0.
    assert values["P02"]["obs_cat_nav"] == 0.0


def test_observation_empty_category_becomes_uncategorized():
    columns, values = overview.build_observation_features(
        [_obs_row("P01", category="", severity="")]
    )
    assert values["P01"]["obs_cat_uncategorized"] == 1.0
    # No severity labels anywhere -> no severity columns at all.
    assert not [c for c in columns if c["key"].startswith("obs_sev_")]


def test_observation_rows_without_timestamps_are_ignored():
    _, values = overview.build_observation_features([_obs_row("P01", timestamps=0)])
    assert values == {}


# ---- Screenspace features ----------------------------------------------------


def test_screenspace_rates_use_last_event_time():
    # Session proxy = last time_out (120 s = 2 min); 4 change events -> 2/min.
    rows = [
        _ss_row(detector="change", time_in=t, time_out=t + 5) for t in (0, 30, 60, 115)
    ]
    rows.append(_ss_row(detector="text", time_in=50, time_out=55, navigational=True))
    columns, values = overview.build_screenspace_features(rows)
    p01 = values["P01"]
    assert p01["ss_rate_change"] == pytest.approx(2.0)
    assert p01["ss_total_rate"] == pytest.approx(2.5)
    assert p01["ss_nav_share"] == pytest.approx(1 / 5)
    assert p01["ss_conf_change"] == pytest.approx(0.8)


# ---- Transcript features -------------------------------------------------------


def test_transcript_features_fallback_scores_without_friction_field():
    # No "friction" key on the entry: the deterministic scorer must run.
    source = {
        "P01": {
            "segments": _segments(
                "P01", ["um let me think", "this is broken", "plain sailing here"]
            )
        }
    }
    columns, values = overview.build_transcript_features(source)
    p01 = values["P01"]
    assert {c["key"] for c in columns} >= {
        f"tr_fric_{c}" for c in friction.CATEGORY_ORDER
    }
    assert p01["tr_fric_hesitation"] > 0
    assert p01["tr_fric_frustration"] > 0
    assert p01["tr_segment_count"] == 3.0
    assert p01["tr_duration_min"] == pytest.approx(1.5)
    assert p01["tr_words_per_min"] > 0


def test_transcript_features_prefer_manifest_friction_stats():
    # Stats present in the manifest win over recomputation (2-min session).
    source = {
        "P01": {
            "segments": _segments("P01", ["a", "b", "c", "d"]),
            "friction": {
                "segments": [],
                "stats": {
                    "by_category": {"confusion": 6},
                    "markers_per_minute": 3.0,
                    "total_markers": 6,
                },
            },
        }
    }
    _, values = overview.build_transcript_features(source)
    assert values["P01"]["tr_fric_confusion"] == pytest.approx(3.0)
    assert values["P01"]["tr_markers_per_min"] == 3.0


def test_transcript_mark_features():
    # Marks (researcher annotations) contribute marks/min + category shares;
    # participant comes from the mark's segment_id prefix when absent.
    source = {"P01": {"segments": _segments("P01", ["a", "b"], seconds_each=60.0)}}
    marks = [
        {"segment_id": "P01:0", "category": "pain_point"},
        {"segment_id": "P01:1", "category": "pain_point"},
        {"segment_id": "P01:1", "category": "quote"},
    ]
    columns, values = overview.build_transcript_features(source, marks)
    keys = {c["key"] for c in columns}
    assert "tr_marks_per_min" in keys
    assert "tr_mark_pain_point" in keys and "tr_mark_quote" in keys
    p01 = values["P01"]
    assert p01["tr_marks_per_min"] == pytest.approx(1.5)  # 3 marks / 2 min
    assert p01["tr_mark_pain_point"] == pytest.approx(2 / 3, abs=1e-3)


def test_transcript_entry_without_segments_is_skipped():
    _, values = overview.build_transcript_features({"P01": {"segments": []}})
    assert values == {}


# ---- Session shape --------------------------------------------------------------


def test_session_shape_bins_place_events_and_sum_to_one():
    # Events at the very start and very end of a 100 s session.
    rows = [
        _ss_row(time_in=0.0, time_out=2.0),
        _ss_row(time_in=1.0, time_out=3.0),
        _ss_row(time_in=98.0, time_out=100.0),
    ]
    columns, values = overview.build_session_shape_features(rows, {}, bins=4)
    p01 = values["P01"]
    ss_bins = [p01[f"shape_ss_bin{i}"] for i in range(4)]
    assert sum(ss_bins) == pytest.approx(1.0, abs=1e-3)
    assert ss_bins[0] == pytest.approx(2 / 3, abs=1e-3)
    assert ss_bins[3] == pytest.approx(1 / 3, abs=1e-3)
    # No transcript -> friction bins exist but are all zero.
    assert all(p01[f"shape_fric_bin{i}"] == 0.0 for i in range(4))


def test_session_shape_friction_bins_from_segments():
    source = {
        "P01": {
            "segments": _segments(
                "P01", ["um well um", "fine", "fine", "this is broken argh"]
            )
        }
    }
    _, values = overview.build_session_shape_features([], source, bins=4)
    p01 = values["P01"]
    fric_bins = [p01[f"shape_fric_bin{i}"] for i in range(4)]
    assert sum(fric_bins) == pytest.approx(1.0, abs=1e-3)
    assert fric_bins[0] > 0 and fric_bins[3] > 0
    assert fric_bins[1] == 0.0 and fric_bins[2] == 0.0


# ---- Assembly ----------------------------------------------------------------


@pytest.fixture
def seeded_output_dir(tmp_path, monkeypatch):
    """Point the manifest loaders at a tmp output dir; return it for seeding."""
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path), raising=False)
    return tmp_path


def _write_manifests(output_dir, *, screenspace=None, transcripts=None, clips=None):
    if screenspace is not None:
        path = output_dir / config.SCREENSPACE_MANIFEST_FILENAME
        path.write_text(json.dumps(screenspace), encoding="utf-8")
    if transcripts is not None:
        path = output_dir / config.TRANSCRIPTS_MANIFEST_FILENAME
        path.write_text(json.dumps(transcripts), encoding="utf-8")
    if clips is not None:
        path = output_dir / config.MANIFEST_FILENAME
        path.write_text(json.dumps(clips), encoding="utf-8")


def _ss_manifest(events):
    return {"regions": {}, "tasks": [], "events": events, "stashes": [], "pins": {}}


def test_feature_matrix_empty_output_dir(seeded_output_dir):
    payload = overview.build_feature_matrix()
    assert payload["participants"] == []
    assert payload["matrix"] == []
    assert payload["availability"] == {}
    # Fixed columns (transcript categories, shape bins, totals) exist even with
    # no data; dynamic ones (categories, detectors) don't.
    assert all(isinstance(c["key"], str) for c in payload["columns"])
    assert [g["key"] for g in payload["groups"]] == list(overview.GROUP_KEYS)
    assert "config" in payload


def test_feature_matrix_missing_groups_are_none(seeded_output_dir):
    _write_manifests(
        seeded_output_dir,
        screenspace=_ss_manifest([_ss_row(participant="P01")]),
        transcripts={
            "source_transcripts": {"P02": {"segments": _segments("P02", ["um okay"])}},
            "corrections": [],
            "marks": [],
        },
    )
    payload = overview.build_feature_matrix()
    assert payload["participants"] == ["P01", "P02"]
    assert payload["availability"]["P01"] == {
        "observations": False,
        "screenspace": True,
        "transcript": False,
        "session_shape": True,
    }
    assert payload["availability"]["P02"]["screenspace"] is False
    assert payload["availability"]["P02"]["transcript"] is True

    cols = payload["columns"]
    p01_row = payload["matrix"][0]
    p02_row = payload["matrix"][1]
    for j, col in enumerate(cols):
        if col["group"] == "transcript":
            assert p01_row[j] is None
        if col["group"] == "screenspace":
            assert p02_row[j] is None
            assert p01_row[j] is not None


def test_feature_matrix_scrubs_nonfinite_floats(seeded_output_dir):
    _write_manifests(
        seeded_output_dir,
        screenspace=_ss_manifest([_ss_row(participant="P01", confidence=math.nan)]),
    )
    payload = overview.build_feature_matrix()
    text = json.dumps(payload)  # must not raise / emit bare NaN
    assert "NaN" not in text


def test_feature_matrix_is_deterministic(seeded_output_dir):
    _write_manifests(
        seeded_output_dir,
        screenspace=_ss_manifest(
            [_ss_row(participant="P01"), _ss_row(participant="P02", detector="text")]
        ),
    )
    assert overview.build_feature_matrix() == overview.build_feature_matrix()


# ---- Route smokes ----------------------------------------------------------


@pytest.fixture
def overview_client(seeded_output_dir):
    Flask = pytest.importorskip("flask").Flask
    app = Flask(__name__)
    app.register_blueprint(overview.overview_bp, url_prefix="/overview")
    with app.test_client() as c:
        yield c


def test_overview_page_serves(overview_client):
    resp = overview_client.get("/overview/")
    assert resp.status_code == 200
    assert "text/html" in resp.content_type
    assert 'data-frontend="overview"' in resp.get_data(as_text=True)


def test_overview_api_data_is_json_not_shadowed_by_static_route(overview_client):
    resp = overview_client.get("/overview/api/data")
    assert resp.status_code == 200
    assert "application/json" in resp.content_type
    body = resp.get_json()
    assert body["ok"] is True
    assert body["participants"] == []


def test_overview_api_data_with_seeded_manifest(overview_client, seeded_output_dir):
    _write_manifests(
        seeded_output_dir,
        screenspace=_ss_manifest([_ss_row(participant="P01")]),
    )
    body = overview_client.get("/overview/api/data").get_json()
    assert body["participants"] == ["P01"]
    assert body["availability"]["P01"]["screenspace"] is True
    assert body["availability"]["P01"]["observations"] is False


def test_vendored_three_js_served_by_static_route(overview_client):
    resp = overview_client.get("/overview/vendor/three.min.js")
    assert resp.status_code == 200
    assert b"THREE" in resp.data[:1000]


# ---- Convergence offsets routes (moved here with the Convergence tab) ----


def test_api_convergence_offsets_get_empty(overview_client, seeded_output_dir):
    resp = overview_client.get("/overview/api/convergence/offsets")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["offsets"] == {}


def test_api_convergence_offsets_put_persists(overview_client, seeded_output_dir):
    # Nested per-lane shape: the transcript: 0 lane is dropped on cleaning.
    payload = {
        "offsets": {
            "P01": {"sheet": 12.5, "screenspace": 12.5, "transcript": 0},
            "P03": {"sheet": -7.0},
        }
    }
    expected = {"P01": {"sheet": 12.5, "screenspace": 12.5}, "P03": {"sheet": -7.0}}
    resp = overview_client.put("/overview/api/convergence/offsets", json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["offsets"] == expected

    saved = json.loads(
        (seeded_output_dir / config.CONVERGENCE_OFFSETS_FILENAME).read_text()
    )
    assert saved == {"offsets": expected}

    resp2 = overview_client.get("/overview/api/convergence/offsets")
    assert resp2.get_json()["offsets"] == expected


def test_api_convergence_offsets_put_strips_zeros_and_garbage(
    overview_client, seeded_output_dir
):
    payload = {
        "offsets": {
            # Surviving participant: only the non-zero, known-source lane stays.
            "P01": {"sheet": 5.0, "screenspace": 0, "transcript": "nope"},
            # All lanes zero/garbage -> participant dropped entirely.
            "P02": {"sheet": 0, "screenspace": float("inf")},
            # Unknown source key dropped -> participant has no lanes -> dropped.
            "P03": {"audio": 5.0},
            # Non-dict participant value -> dropped (not a 400).
            "P04": 5,
            # Empty participant id -> dropped.
            "": {"sheet": 9.0},
        }
    }
    resp = overview_client.put("/overview/api/convergence/offsets", json=payload)
    assert resp.status_code == 200
    assert resp.get_json()["offsets"] == {"P01": {"sheet": 5.0}}


def test_api_convergence_offsets_round_trip_per_lane(
    overview_client, seeded_output_dir
):
    """A participant with one non-zero and one zero lane round-trips to just the
    non-zero lane through PUT then GET, keeping the get/put cleaners in sync."""
    payload = {"offsets": {"P02": {"sheet": 0, "transcript": -3.5}}}
    put = overview_client.put("/overview/api/convergence/offsets", json=payload)
    assert put.status_code == 200
    assert put.get_json()["offsets"] == {"P02": {"transcript": -3.5}}

    get = overview_client.get("/overview/api/convergence/offsets")
    assert get.get_json()["offsets"] == {"P02": {"transcript": -3.5}}


def test_api_convergence_offsets_put_empty_removes_file(
    overview_client, seeded_output_dir
):
    # Seed a manifest file via an earlier put.
    overview_client.put(
        "/overview/api/convergence/offsets",
        json={"offsets": {"P01": {"sheet": 3.0}}},
    )
    settings_file = seeded_output_dir / config.CONVERGENCE_OFFSETS_FILENAME
    assert settings_file.is_file()

    resp = overview_client.put(
        "/overview/api/convergence/offsets",
        json={"offsets": {"P01": {"sheet": 0}}},
    )
    assert resp.status_code == 200
    assert resp.get_json()["offsets"] == {}
    assert not settings_file.is_file()


def test_api_convergence_offsets_put_rejects_non_dict(
    overview_client, seeded_output_dir
):
    resp = overview_client.put(
        "/overview/api/convergence/offsets", json={"offsets": "nope"}
    )
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False


# ---- Injected sheet observation rows ---------------------------------------


def test_feature_matrix_uses_injected_observation_rows(seeded_output_dir, monkeypatch):
    """With the server-injected getter wired, sheet timestamps alone populate
    the observations group — no clip artifacts involved."""
    monkeypatch.setattr(
        overview,
        "_observation_rows_getter",
        lambda: [_obs_row("P05", category="nav", timestamps=3)],
    )
    payload = overview.build_feature_matrix()
    assert payload["participants"] == ["P05"]
    assert payload["availability"]["P05"]["observations"] is True
    cols = {c["key"]: j for j, c in enumerate(payload["columns"])}
    assert payload["matrix"][0][cols["obs_total"]] == 3.0


def test_api_friction_moments_resolves_times(overview_client, seeded_output_dir):
    _write_manifests(
        seeded_output_dir,
        transcripts={
            "source_transcripts": {
                "P01": {
                    "segments": _segments("P01", ["a", "b", "c"]),
                    "friction": {
                        "segments": [],
                        "stats": {
                            "by_category": {},
                            "markers_per_minute": 0,
                            "total_markers": 0,
                        },
                        "moments": [
                            {
                                "segment_ids": ["P01:1", "P01:2"],
                                "category": "confusion",
                                "rationale": "lost in nav",
                                "score": 0.8,
                            },
                            # Dangling segment ids -> dropped, not crashed.
                            {
                                "segment_ids": ["P01:99"],
                                "category": "x",
                                "rationale": "y",
                                "score": 0.1,
                            },
                        ],
                    },
                }
            },
            "corrections": [],
            "marks": [],
        },
    )
    body = overview_client.get("/overview/api/friction-moments").get_json()
    assert body["ok"] is True
    assert len(body["moments"]) == 1
    m = body["moments"][0]
    assert m["participant"] == "P01"
    assert m["start"] == 30.0 and m["end"] == 90.0
    assert m["rationale"] == "lost in nav"
