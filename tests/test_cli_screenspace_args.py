"""Tests for Screenspace CLI flags and dispatch helpers."""

from argparse import Namespace

import pytest

import cli


def _ss_args(**overrides):
    """Return a Namespace pre-populated with all Screenspace + thinking-agent CLI keys."""
    defaults = {
        "batch": False,
        "lines": None,
        "range": None,
        "category": None,
        "cell": None,
        "participant": None,
        "keyword": False,
        "severity": None,
        "mixed": None,
        "reel": None,
        "chronologic": None,
        "highlights": None,
        "screen": False,
        "gif": False,
        "no_input": False,
        "verbose": False,
        "spreadsheet": None,
        "viewer": False,
        "manifest": False,
        "regenerate": False,
        "studio": False,
        "screenspace": False,
        "transcripts": False,
        "timeline_viewer": False,
        "gallery": None,
        "interval": None,
        "bundle": False,
        "input": None,
        "output": None,
        "titlecards": None,
        "filmstrip": None,
        "transcribe": False,
        "transcript_format": None,
        "pre_transcribe": None,
        "whisper_model": None,
        "llm_model": None,
        "summarize": None,
        "citations": None,
        "ss_task": None,
        "ss_run_task": None,
        "ss_list_regions": False,
        "ss_list_stashes": False,
        "ss_list_tasks": None,
        "ss_target_color": None,
        "ss_tolerance": None,
        "ss_color_mode": "average",
        "ss_min_area": None,
        "ss_threshold": None,
        "ss_reference_timestamp": None,
        "ss_scene_ref": None,
        "ss_text": None,
        "ss_fuzzy_threshold": None,
        "ss_operator": None,
        "ss_target_value": None,
        "ss_range_min": None,
        "ss_range_max": None,
        "ss_speedup": None,
        "ss_output_format": None,
        "ss_start": None,
        "ss_end": None,
        "ss_interval": None,
        "ss_event_label": None,
        "ss_clips": False,
        "transcript_clips": False,
        "ss_clips_detector": None,
        "ss_clips_region": None,
        "ss_clips_participant": None,
        "ss_clips_min_confidence": None,
        "ss_clips_event_type": None,
        "transcript_clips_participant": None,
        "transcript_clips_mark": None,
        "transcript_clips_text": None,
        "transcript_mark": None,
        "transcript_mark_category": None,
        "transcript_mark_participant": None,
        "transcript_mark_label": None,
        "cluster_gap": 5.0,
        "clip_pre": 5.0,
        "clip_post": 5.0,
        "max_clip_duration": 0.0,
        "export": False,
    }
    defaults.update(overrides)
    return Namespace(**defaults)


# ---- Argparse parsing ----


# Each argv case is a single space-split string (no argument here contains a
# space); `expected` maps parsed-Namespace attrs to their values.
@pytest.mark.parametrize(
    "argv,expected",
    [
        pytest.param(
            "--ss-task color P01 button --ss-target-color #FF0000"
            " --ss-tolerance 20,30,30 --ss-threshold 0.85",
            {
                "ss_task": ["color", "P01", "button"],
                "ss_target_color": "#FF0000",
                "ss_tolerance": "20,30,30",
                "ss_threshold": 0.85,
                "ss_color_mode": "average",  # asserts the default
            },
            id="color-minimal",
        ),
        pytest.param(
            # REGION is optional — `--ss-task TYPE PARTICIPANT` parses to a
            # 2-element list.
            "--ss-task color P01 --ss-target-color #FF0000"
            " --ss-tolerance 20,30,30 --ss-threshold 0.85",
            {"ss_task": ["color", "P01"]},
            id="region-optional",
        ),
        pytest.param(
            "--ss-task color P01 button --ss-target-color #8B0000"
            " --ss-tolerance 20,80,80 --ss-color-mode presence --ss-min-area 1",
            {"ss_color_mode": "presence", "ss_min_area": 1.0},
            id="color-presence",
        ),
        pytest.param(
            "--ss-list-tasks completed",
            {"ss_list_tasks": "completed"},
            id="list-tasks-with-status",
        ),
        pytest.param(
            "--ss-list-tasks",
            {"ss_list_tasks": ""},  # const="" when no value
            id="list-tasks-no-status",
        ),
        pytest.param(
            "--ss-run-task ss_abc123", {"ss_run_task": "ss_abc123"}, id="run-task"
        ),
        pytest.param(
            "--ss-task scene P01 btn --ss-scene-ref menu:12.5 --ss-scene-ref game:30:0.8",
            {
                "ss_task": ["scene", "P01", "btn"],
                "ss_scene_ref": ["menu:12.5", "game:30:0.8"],
            },
            id="scene-ref-repeatable",
        ),
    ],
)
def test_parse_ss_flags(monkeypatch, argv, expected):
    monkeypatch.setattr("sys.argv", ["clipgen.py"] + argv.split())
    args = cli.parse_arguments()
    for attr, value in expected.items():
        assert getattr(args, attr) == value


@pytest.mark.parametrize(
    "argv",
    [
        pytest.param(
            "--ss-list-regions --ss-list-stashes", id="list-modes-mutually-exclusive"
        ),
        pytest.param(
            "--ss-run-task ss_abc --ss-task color P01 btn",
            id="run-task-and-task-mutually-exclusive",
        ),
    ],
)
def test_parse_ss_conflicting_flags_exit(monkeypatch, argv):
    monkeypatch.setattr("sys.argv", ["clipgen.py"] + argv.split())
    with pytest.raises(SystemExit):
        cli.parse_arguments()


def test_ss_task_omitted_region_defaults_to_full_frame(monkeypatch, capsys):
    # With no REGION, resolution must succeed (full_frame) and fall through to the
    # next step — here the missing video — rather than erroring on the region.
    fake_manifest = {"regions": {}, "stashes": [], "tasks": [], "events": []}
    import screenspace

    monkeypatch.setattr(screenspace, "load_screenspace_manifest", lambda: fake_manifest)
    monkeypatch.setattr(cli, "_ss_resolve_videos_for_participant", lambda pid: [])
    args = _ss_args(
        ss_task=["color", "P01"],
        ss_target_color="#FF0000",
        ss_tolerance="20,30,30",
        ss_threshold=0.85,
    )
    with pytest.raises(SystemExit) as exc:
        cli._run_ss_task(args)
    assert exc.value.code == 1
    # Reached the video check (region resolved as full_frame), not a region error.
    assert "P01" in capsys.readouterr().out


def test_ss_build_params_color_presence():
    region = {"x": 0, "y": 0, "w": 100, "h": 100}
    args = _ss_args(
        ss_target_color="#8B0000",
        ss_tolerance="20,80,80",
        ss_color_mode="presence",
        ss_min_area=1.0,
    )
    params = cli._ss_build_params(args, "color", region, lambda _ts: None)
    assert params["color_mode"] == "presence"
    assert params["min_coverage"] == pytest.approx(0.01)


def test_ss_build_params_color_average_omits_mode():
    region = {"x": 0, "y": 0, "w": 100, "h": 100}
    args = _ss_args(ss_target_color="#FF0000", ss_tolerance="20,30,30")
    params = cli._ss_build_params(args, "color", region, lambda _ts: None)
    assert "color_mode" not in params
    assert "min_coverage" not in params


# ---- Conflict validation ----


@pytest.mark.parametrize(
    "overrides",
    [
        pytest.param(
            {
                "ss_task": ["color", "P01", "button"],
                "ss_target_color": "#FF0000",
                "ss_tolerance": "20,30,30",
                "ss_threshold": 0.85,
                "studio": True,
            },
            id="ss-task-vs-studio",
        ),
        pytest.param(
            {"summarize": [], "pre_transcribe": ["P01"]},
            id="summarize-vs-pre-transcribe",
        ),
        pytest.param(
            {"ss_run_task": "ss_abc", "studio": True}, id="ss-run-task-vs-studio"
        ),
    ],
)
def test_mode_conflicts_exit(overrides):
    with pytest.raises(SystemExit):
        cli._validate_mode_conflicts(_ss_args(**overrides))


def test_ss_run_task_marks_cli_mode():
    modes = cli._validate_mode_conflicts(_ss_args(ss_run_task="ss_abc"))
    assert modes["ss_run_task"] is True


def test_validate_returns_dict_shape():
    args = _ss_args(ss_list_regions=True)
    modes = cli._validate_mode_conflicts(args)
    assert isinstance(modes, dict)
    assert modes["ss_list_regions"] is True
    assert modes["studio"] is False
    assert modes["gallery_arg"] is None


# ---- Helper conversions ----


def test_ss_hex_to_hsv_red_only():
    hsv = cli._ss_hex_to_hsv("#FF0000")
    # OpenCV HSV: red is hue 0, saturation 255, value 255
    assert hsv["s"] == 255
    assert hsv["v"] == 255


def test_ss_parse_tolerance_valid():
    tol = cli._ss_parse_tolerance("20,30,30")
    assert tol == {"h": 20, "s": 30, "v": 30}


@pytest.mark.parametrize(
    "fn,raw",
    [
        pytest.param(cli._ss_hex_to_hsv, "#ABC", id="hex-invalid-length"),
        pytest.param(cli._ss_parse_tolerance, "20,30", id="tolerance-wrong-count"),
    ],
)
def test_ss_conversion_invalid_input_raises(fn, raw):
    with pytest.raises(ValueError):
        fn(raw)


@pytest.mark.parametrize(
    "raw,expected",
    [
        pytest.param(
            "menu:12.5", {"name": "menu", "timestamp": 12.5}, id="name-timestamp"
        ),
        pytest.param(
            "game:30:0.8",
            {"name": "game", "timestamp": 30.0, "threshold": 0.8},
            id="with-threshold",
        ),
    ],
)
def test_ss_parse_scene_ref_valid(raw, expected):
    assert cli._ss_parse_scene_ref(raw) == expected


@pytest.mark.parametrize(
    "raw,match",
    [
        pytest.param("menu", None, id="missing-timestamp"),
        pytest.param("menu:soon", None, id="nonnumeric-timestamp"),
        pytest.param("menu:12.5:1.5", "between 0 and 1", id="threshold-too-high"),
        pytest.param("menu:12.5:-0.1", "between 0 and 1", id="threshold-negative"),
    ],
)
def test_ss_parse_scene_ref_invalid_raises(raw, match):
    with pytest.raises(ValueError, match=match):
        cli._ss_parse_scene_ref(raw)


# ---- Listing helpers ----


def test_ss_list_regions_outputs_names(capsys, monkeypatch):
    fake_manifest = {
        "regions": {
            "btn": {
                "x": 0.1,
                "y": 0.2,
                "w": 0.3,
                "h": 0.4,
                "source_width": 1920,
                "source_height": 1080,
            },
            "viewport": {
                "x": 0.0,
                "y": 0.0,
                "w": 1.0,
                "h": 1.0,
                "source_width": 1920,
                "source_height": 1080,
            },
        },
        "stashes": [],
        "tasks": [],
        "events": [],
    }
    import screenspace

    monkeypatch.setattr(screenspace, "load_screenspace_manifest", lambda: fake_manifest)
    cli._run_ss_list_regions(_ss_args(ss_list_regions=True))
    out = capsys.readouterr().out
    assert "btn" in out
    assert "viewport" in out


def test_ss_list_stashes_outputs_names(capsys, monkeypatch):
    fake_manifest = {
        "regions": {},
        "stashes": [
            {
                "id": "stash_abc",
                "name": "saved_v1",
                "regions": {"btn": {"x": 0.1, "y": 0.2, "w": 0.3, "h": 0.4}},
            }
        ],
        "tasks": [],
        "events": [],
    }
    import screenspace

    monkeypatch.setattr(screenspace, "load_screenspace_manifest", lambda: fake_manifest)
    cli._run_ss_list_stashes(_ss_args(ss_list_stashes=True))
    out = capsys.readouterr().out
    assert "saved_v1" in out
    assert "btn" in out


def test_ss_list_tasks_filters_by_status(capsys, monkeypatch):
    fake_manifest = {
        "regions": {},
        "stashes": [],
        "tasks": [
            {
                "id": "ss_1",
                "type": "color",
                "participant": "P01",
                "region": "btn",
                "status": "completed",
                "result": [{"timestamp": 1.0}],
            },
            {
                "id": "ss_2",
                "type": "change",
                "participant": "P02",
                "region": "btn",
                "status": "failed",
                "result": None,
            },
        ],
        "events": [],
    }
    import screenspace

    monkeypatch.setattr(screenspace, "load_screenspace_manifest", lambda: fake_manifest)
    cli._run_ss_list_tasks(_ss_args(ss_list_tasks="completed"))
    out = capsys.readouterr().out
    assert "ss_1" in out
    assert "ss_2" not in out


# ---- ss_task dispatcher errors ----


def test_ss_task_unknown_region_errors(capsys, monkeypatch):
    fake_manifest = {
        "regions": {"existing": {"x": 0, "y": 0, "w": 0.5, "h": 0.5}},
        "stashes": [],
        "tasks": [],
        "events": [],
    }
    import screenspace

    monkeypatch.setattr(screenspace, "load_screenspace_manifest", lambda: fake_manifest)
    args = _ss_args(
        ss_task=["color", "P01", "missing"],
        ss_target_color="#FF0000",
        ss_tolerance="20,30,30",
        ss_threshold=0.85,
    )
    with pytest.raises(SystemExit) as exc:
        cli._run_ss_task(args)
    assert exc.value.code == 1
    err = capsys.readouterr().out
    assert "missing" in err
    assert "existing" in err


def test_ss_task_unknown_type_errors(monkeypatch, capsys):
    args = _ss_args(ss_task=["bogus_tool", "P01", "btn"])
    with pytest.raises(SystemExit) as exc:
        cli._run_ss_task(args)
    assert exc.value.code == 1
    err = capsys.readouterr().out
    assert "bogus_tool" in err


def test_ss_task_no_video_errors(monkeypatch, capsys):
    fake_manifest = {
        "regions": {
            "btn": {
                "x": 0.1,
                "y": 0.2,
                "w": 0.3,
                "h": 0.4,
                "source_width": 1920,
                "source_height": 1080,
            }
        },
        "stashes": [],
        "tasks": [],
        "events": [],
    }
    import screenspace

    monkeypatch.setattr(screenspace, "load_screenspace_manifest", lambda: fake_manifest)
    monkeypatch.setattr(cli, "_ss_resolve_videos_for_participant", lambda pid: [])

    args = _ss_args(
        ss_task=["color", "P01", "btn"],
        ss_target_color="#FF0000",
        ss_tolerance="20,30,30",
        ss_threshold=0.85,
    )
    with pytest.raises(SystemExit) as exc:
        cli._run_ss_task(args)
    assert exc.value.code == 1
    err = capsys.readouterr().out
    assert "P01" in err


def test_ss_task_color_missing_target_color_errors(monkeypatch, capsys):
    fake_manifest = {
        "regions": {
            "btn": {
                "x": 0.1,
                "y": 0.2,
                "w": 0.3,
                "h": 0.4,
                "source_width": 1920,
                "source_height": 1080,
            }
        },
        "stashes": [],
        "tasks": [],
        "events": [],
    }
    import screenspace
    import video as video_mod

    monkeypatch.setattr(screenspace, "load_screenspace_manifest", lambda: fake_manifest)
    monkeypatch.setattr(
        cli, "_ss_resolve_videos_for_participant", lambda pid: ["/tmp/fake.mp4"]
    )
    monkeypatch.setattr(
        video_mod,
        "probe_video_properties",
        lambda p: {"width": 1920, "height": 1080},
    )

    args = _ss_args(ss_task=["color", "P01", "btn"])  # missing color params
    with pytest.raises(SystemExit) as exc:
        cli._run_ss_task(args)
    assert exc.value.code == 1
    err = capsys.readouterr().out
    assert "target-color" in err.lower() or "target_color" in err.lower()


# ---- ss_task happy path with stubbed worker ----


def test_ss_task_color_dispatches_and_persists(monkeypatch):
    """Smoke: ss_task builds a task, enqueues, polls to completion, persists."""
    fake_manifest = {
        "regions": {
            "btn": {
                "x": 0.0,
                "y": 0.0,
                "w": 0.5,
                "h": 0.5,
                "source_width": 100,
                "source_height": 100,
            }
        },
        "stashes": [],
        "tasks": [],
        "events": [],
    }

    import screenspace
    import video as video_mod

    monkeypatch.setattr(screenspace, "load_screenspace_manifest", lambda: fake_manifest)
    monkeypatch.setattr(
        cli, "_ss_resolve_videos_for_participant", lambda pid: ["/tmp/fake.mp4"]
    )
    monkeypatch.setattr(
        video_mod,
        "probe_video_properties",
        lambda p: {"width": 100, "height": 100},
    )

    saved_tasks: list[dict] = []

    def fake_save(
        regions, tasks, events, stashes=None, per_participant=None, pins=None
    ):
        saved_tasks.extend(tasks)

    monkeypatch.setattr(screenspace, "save_screenspace_manifest", fake_save)

    class FakeWorker:
        def __init__(self):
            self.task_id: str = ""
            self.task_dict: dict = {}

        def restore_tasks(self, tasks):
            pass

        def start(self):
            pass

        def stop(self):
            pass

        def enqueue(self, task):
            self.task_id = task["id"]
            task["status"] = "completed"
            task["progress"] = 1.0
            task["result"] = [{"timestamp": 1.0}]
            self.task_dict = task
            return task["id"]

        def get_task(self, tid):
            return dict(self.task_dict) if tid == self.task_id else None

        def get_all_tasks(self):
            return [dict(self.task_dict)] if self.task_dict else []

        def drain_new_events(self):
            return []

    monkeypatch.setattr(screenspace, "ScreenspaceWorker", FakeWorker)

    args = _ss_args(
        ss_task=["color", "P01", "btn"],
        ss_target_color="#FF0000",
        ss_tolerance="20,30,30",
        ss_threshold=0.85,
    )
    cli._run_ss_task(args)

    assert saved_tasks, "Manifest should have been persisted with the new task"
    persisted = saved_tasks[0]
    assert persisted["type"] == "color"
    assert persisted["participant"] == "P01"
    assert persisted["region"] == "btn"
    assert persisted["status"] == "completed"


def test_parse_ss_task_attention(monkeypatch):
    monkeypatch.setattr(
        "sys.argv",
        ["clipgen.py", "--ss-task", "attention", "P01"],
    )
    args = cli.parse_arguments()
    assert args.ss_task == ["attention", "P01"]


def test_ss_build_params_attention_threshold_optional():
    region = {"x": 0, "y": 0, "w": 100, "h": 100}
    args = _ss_args(ss_threshold=0.2)
    params = cli._ss_build_params(args, "attention", region, lambda ts: None)
    assert params["shift_threshold"] == 0.2

    args = _ss_args()
    params = cli._ss_build_params(args, "attention", region, lambda ts: None)
    assert "shift_threshold" not in params  # config default applies at scan time


def test_ss_task_attention_forces_full_frame(monkeypatch, capsys):
    """A named region on an attention task is ignored with a warning: the task
    persists as full_frame (mirrors the server's forced rewrite)."""
    fake_manifest = {
        "regions": {
            "hud": {
                "x": 0.0,
                "y": 0.0,
                "w": 0.5,
                "h": 0.5,
                "source_width": 100,
                "source_height": 100,
            }
        },
        "stashes": [],
        "tasks": [],
        "events": [],
    }

    import screenspace
    import video as video_mod

    monkeypatch.setattr(screenspace, "load_screenspace_manifest", lambda: fake_manifest)
    monkeypatch.setattr(
        cli, "_ss_resolve_videos_for_participant", lambda pid: ["/tmp/fake.mp4"]
    )
    monkeypatch.setattr(
        video_mod,
        "probe_video_properties",
        lambda p: {"width": 100, "height": 100},
    )

    saved_tasks: list[dict] = []

    def fake_save(
        regions, tasks, events, stashes=None, per_participant=None, pins=None
    ):
        saved_tasks.extend(tasks)

    monkeypatch.setattr(screenspace, "save_screenspace_manifest", fake_save)
    monkeypatch.setattr(screenspace, "ScreenspaceWorker", _FakeWorker)

    args = _ss_args(ss_task=["attention", "P01", "hud"], ss_threshold=0.2)
    cli._run_ss_task(args)

    assert saved_tasks
    persisted = saved_tasks[0]
    assert persisted["type"] == "attention"
    assert persisted["region"] == "full_frame"
    assert persisted["parameters"]["shift_threshold"] == 0.2
    assert "full-frame only" in capsys.readouterr().out


# ---- scene flag path + manifest re-run (--ss-run-task) ----


class _FakeWorker:
    """Stub ScreenspaceWorker that completes the enqueued task immediately."""

    def __init__(self):
        self.task_id: str = ""
        self.task_dict: dict = {}

    def restore_tasks(self, tasks):
        pass

    def start(self):
        pass

    def stop(self):
        pass

    def enqueue(self, task):
        self.task_id = task["id"]
        task["status"] = "completed"
        task["progress"] = 1.0
        task["result"] = [{"timestamp": 1.0}]
        self.task_dict = task
        return task["id"]

    def get_task(self, tid):
        return dict(self.task_dict) if tid == self.task_id else None

    def get_all_tasks(self):
        return [dict(self.task_dict)] if self.task_dict else []

    def drain_new_events(self):
        return []


def _install_ss_stubs(monkeypatch, fake_manifest):
    """Wire up the common Screenspace stubs (manifest, video, media, worker, save)."""
    import screenspace
    import video as video_mod

    monkeypatch.setattr(screenspace, "load_screenspace_manifest", lambda: fake_manifest)
    monkeypatch.setattr(
        cli, "_ss_resolve_videos_for_participant", lambda pid: ["/tmp/fake.mp4"]
    )
    monkeypatch.setattr(
        video_mod, "probe_video_properties", lambda p: {"width": 100, "height": 100}
    )
    monkeypatch.setattr(video_mod, "extract_frame_at_timestamp", lambda p, ts: [[0]])
    monkeypatch.setattr(screenspace, "extract_region", lambda frame, coords: [[1]])
    monkeypatch.setattr(screenspace, "ScreenspaceWorker", _FakeWorker)

    saved_tasks: list[dict] = []

    def fake_save(
        regions, tasks, events, stashes=None, per_participant=None, pins=None
    ):
        saved_tasks.extend(tasks)

    monkeypatch.setattr(screenspace, "save_screenspace_manifest", fake_save)
    return saved_tasks


def test_ss_task_scene_dispatches_and_persists(monkeypatch):
    """Scene flag path: --ss-scene-ref entries become reference_scenes + scene_references."""
    fake_manifest = {
        "regions": {"btn": {"x": 0.0, "y": 0.0, "w": 0.5, "h": 0.5}},
        "stashes": [],
        "tasks": [],
        "events": [],
    }
    saved_tasks = _install_ss_stubs(monkeypatch, fake_manifest)

    args = _ss_args(
        ss_task=["scene", "P01", "btn"],
        ss_scene_ref=["menu:12.5", "game:30:0.8"],
        ss_threshold=0.9,
    )
    cli._run_ss_task(args)

    assert saved_tasks
    persisted = saved_tasks[0]
    assert persisted["type"] == "scene"
    params = persisted["parameters"]
    assert len(params["scene_references"]) == 2
    assert len(params["reference_scenes"]) == 2
    assert params["scene_references"][1]["threshold"] == 0.8
    assert params["threshold"] == 0.9


def test_ss_task_region_resolves_active_over_stash(monkeypatch):
    """A region name present in both active regions and a stash resolves to the
    active one — regression for the flattened (last-write-wins) lookup that let a
    stashed region with the same name silently shadow the active geometry."""
    import screenspace

    active_btn = {"x": 0.1, "y": 0.1, "w": 0.2, "h": 0.2}
    stash_btn = {"x": 0.5, "y": 0.5, "w": 0.4, "h": 0.4}
    fake_manifest = {
        "regions": {"btn": dict(active_btn)},
        "stashes": [{"id": "s1", "regions": {"btn": dict(stash_btn)}}],
        "tasks": [],
        "events": [],
    }
    saved_tasks = _install_ss_stubs(monkeypatch, fake_manifest)

    args = _ss_args(
        ss_task=["color", "P01", "btn"],
        ss_target_color="#FF0000",
        ss_tolerance="20,30,30",
        ss_threshold=0.85,
    )
    cli._run_ss_task(args)

    assert saved_tasks
    # Active geometry wins; the stash copy must not shadow it.
    assert saved_tasks[0]["region_coords"] == screenspace.denormalize_region(
        active_btn, 100, 100
    )


def test_ss_task_scene_missing_refs_errors(monkeypatch, capsys):
    fake_manifest = {
        "regions": {"btn": {"x": 0.0, "y": 0.0, "w": 0.5, "h": 0.5}},
        "stashes": [],
        "tasks": [],
        "events": [],
    }
    _install_ss_stubs(monkeypatch, fake_manifest)
    args = _ss_args(ss_task=["scene", "P01", "btn"])  # no --ss-scene-ref
    with pytest.raises(SystemExit) as exc:
        cli._run_ss_task(args)
    assert exc.value.code == 1
    assert "scene-ref" in capsys.readouterr().out.lower()


def test_ss_run_task_unknown_id_errors(monkeypatch, capsys):
    fake_manifest = {"regions": {}, "stashes": [], "tasks": [], "events": []}
    import screenspace

    monkeypatch.setattr(screenspace, "load_screenspace_manifest", lambda: fake_manifest)
    with pytest.raises(SystemExit) as exc:
        cli._run_ss_rerun_task(_ss_args(ss_run_task="ss_nope"))
    assert exc.value.code == 1
    assert "ss_nope" in capsys.readouterr().out


def test_ss_run_task_scene_rerun(monkeypatch):
    """Re-run a saved scene task: frames re-extracted from saved scene_references."""
    fake_manifest = {
        "regions": {"btn": {"x": 0.0, "y": 0.0, "w": 0.5, "h": 0.5}},
        "stashes": [],
        "tasks": [
            {
                "id": "ss_scene01",
                "type": "scene",
                "participant": "P01",
                "region": "btn",
                "region_coords": {"x": 0, "y": 0, "w": 50, "h": 50},
                "parameters": {
                    "scene_references": [{"name": "menu", "timestamp": 12.5}],
                    "threshold": 0.9,
                },
                "status": "completed",
            }
        ],
        "events": [],
    }
    saved_tasks = _install_ss_stubs(monkeypatch, fake_manifest)

    cli._run_ss_rerun_task(_ss_args(ss_run_task="ss_scene01"))

    assert saved_tasks
    persisted = saved_tasks[0]
    assert persisted["type"] == "scene"
    assert persisted["id"] != "ss_scene01"  # fresh run, original preserved
    assert len(persisted["parameters"]["reference_scenes"]) == 1


def test_ss_run_task_multitool_rerun(monkeypatch):
    """Re-run a saved multitool task: per-step regions resolved, scene step re-extracted."""
    fake_manifest = {
        "regions": {"btn": {"x": 0.0, "y": 0.0, "w": 0.5, "h": 0.5}},
        "stashes": [],
        "tasks": [
            {
                "id": "ss_mt01",
                "type": "multitool",
                "participant": "P01",
                "region": "btn",
                "region_coords": {"x": 0, "y": 0, "w": 50, "h": 50},
                "parameters": {
                    "steps": [
                        {
                            "type": "color",
                            "region": "btn",
                            "target_color": {"h": 0, "s": 255, "v": 255},
                            "tolerance": {"h": 10, "s": 40, "v": 40},
                        },
                        {
                            "type": "scene",
                            "region": "btn",
                            "offset": {"min": 0, "max": 3},
                            "scene_references": [{"name": "menu", "timestamp": 5.0}],
                        },
                    ]
                },
                "status": "completed",
            }
        ],
        "events": [],
    }
    saved_tasks = _install_ss_stubs(monkeypatch, fake_manifest)

    cli._run_ss_rerun_task(_ss_args(ss_run_task="ss_mt01"))

    assert saved_tasks
    persisted = saved_tasks[0]
    assert persisted["type"] == "multitool"
    steps = persisted["parameters"]["steps"]
    assert len(steps) == 2
    assert "region_coords" in steps[0]  # resolved during rehydration
    assert len(steps[1]["reference_scenes"]) == 1


def test_ss_run_task_uploaded_template_step_errors(monkeypatch, capsys):
    """A multitool template step with no reference_timestamp cannot be re-run."""
    fake_manifest = {
        "regions": {"btn": {"x": 0.0, "y": 0.0, "w": 0.5, "h": 0.5}},
        "stashes": [],
        "tasks": [
            {
                "id": "ss_mt02",
                "type": "multitool",
                "participant": "P01",
                "region": "btn",
                "region_coords": {"x": 0, "y": 0, "w": 50, "h": 50},
                "parameters": {
                    "steps": [
                        {"type": "change", "region": "btn", "threshold": 0.2},
                        {"type": "template", "region": "btn", "threshold": 0.8},
                    ]
                },
                "status": "completed",
            }
        ],
        "events": [],
    }
    _install_ss_stubs(monkeypatch, fake_manifest)
    with pytest.raises(SystemExit) as exc:
        cli._run_ss_rerun_task(_ss_args(ss_run_task="ss_mt02"))
    assert exc.value.code == 1
    assert "cannot be re-run" in capsys.readouterr().out.lower()


def test_ss_run_task_multitool_full_frame_step(monkeypatch):
    """A full_frame step resolves to the whole frame, not the parent task's region."""
    fake_manifest = {
        "regions": {"btn": {"x": 0.0, "y": 0.0, "w": 0.5, "h": 0.5}},
        "stashes": [],
        "tasks": [
            {
                "id": "ss_ff01",
                "type": "multitool",
                "participant": "P01",
                "region": "btn",
                "region_coords": {"x": 0, "y": 0, "w": 50, "h": 50},
                "parameters": {
                    "steps": [
                        {
                            "type": "color",
                            "region": "full_frame",
                            "region_ref": {"source": "full_frame"},
                            "target_color": {"h": 0, "s": 255, "v": 255},
                            "tolerance": {"h": 10, "s": 40, "v": 40},
                        }
                    ]
                },
                "status": "completed",
            }
        ],
        "events": [],
    }
    saved_tasks = _install_ss_stubs(monkeypatch, fake_manifest)

    cli._run_ss_rerun_task(_ss_args(ss_run_task="ss_ff01"))

    assert saved_tasks
    step = saved_tasks[0]["parameters"]["steps"][0]
    # Probe stub reports 100x100, so full frame is {0,0,100,100}, not parent's {0,0,50,50}.
    assert step["region_coords"] == {"x": 0, "y": 0, "w": 100, "h": 100}
    assert step["region"] == "full_frame"


def test_ss_run_task_multitool_stash_step_disambiguates(monkeypatch):
    """A stash-backed step honors stash_id instead of a last-write-wins name merge."""
    fake_manifest = {
        "regions": {"btn": {"x": 0.0, "y": 0.0, "w": 0.5, "h": 0.5}},
        "stashes": [
            {
                "id": "stash_a",
                "regions": {"hud": {"x": 0.1, "y": 0.1, "w": 0.1, "h": 0.1}},
            },
            {
                "id": "stash_b",
                "regions": {"hud": {"x": 0.2, "y": 0.2, "w": 0.2, "h": 0.2}},
            },
        ],
        "tasks": [
            {
                "id": "ss_stash01",
                "type": "multitool",
                "participant": "P01",
                "region": "btn",
                "region_coords": {"x": 0, "y": 0, "w": 50, "h": 50},
                "parameters": {
                    "steps": [
                        {
                            "type": "color",
                            "region": "hud",
                            # Points at stash_a (the non-last stash); a last-write-wins
                            # merge would wrongly pick stash_b's "hud".
                            "region_ref": {
                                "source": "stash",
                                "stash_id": "stash_a",
                                "name": "hud",
                            },
                            "target_color": {"h": 0, "s": 255, "v": 255},
                            "tolerance": {"h": 10, "s": 40, "v": 40},
                        }
                    ]
                },
                "status": "completed",
            }
        ],
        "events": [],
    }
    saved_tasks = _install_ss_stubs(monkeypatch, fake_manifest)

    cli._run_ss_rerun_task(_ss_args(ss_run_task="ss_stash01"))

    assert saved_tasks
    step = saved_tasks[0]["parameters"]["steps"][0]
    # stash_a's hud {0.1,0.1,0.1,0.1} at 100x100 -> {10,10,10,10}, not stash_b's {20,...}.
    assert step["region_coords"] == {"x": 10, "y": 10, "w": 10, "h": 10}


def test_ss_run_task_multitool_step_missing_region_errors(monkeypatch, capsys):
    """A step whose region is gone fails loudly instead of silently using parent coords."""
    fake_manifest = {
        "regions": {"btn": {"x": 0.0, "y": 0.0, "w": 0.5, "h": 0.5}},
        "stashes": [],
        "tasks": [
            {
                "id": "ss_gone01",
                "type": "multitool",
                "participant": "P01",
                "region": "btn",
                "region_coords": {"x": 0, "y": 0, "w": 50, "h": 50},
                "parameters": {
                    "steps": [
                        {
                            "type": "color",
                            "region": "ghost",
                            "target_color": {"h": 0, "s": 255, "v": 255},
                            "tolerance": {"h": 10, "s": 40, "v": 40},
                        }
                    ]
                },
                "status": "completed",
            }
        ],
        "events": [],
    }
    _install_ss_stubs(monkeypatch, fake_manifest)
    with pytest.raises(SystemExit) as exc:
        cli._run_ss_rerun_task(_ss_args(ss_run_task="ss_gone01"))
    assert exc.value.code == 1
    assert "not found" in capsys.readouterr().out.lower()


def test_ss_run_task_parent_stash_disambiguates(monkeypatch):
    """The parent task honors a saved region_ref instead of a last-write-wins merge."""
    fake_manifest = {
        "regions": {},
        "stashes": [
            {
                "id": "stash_a",
                "regions": {"hud": {"x": 0.1, "y": 0.1, "w": 0.1, "h": 0.1}},
            },
            {
                "id": "stash_b",
                "regions": {"hud": {"x": 0.2, "y": 0.2, "w": 0.2, "h": 0.2}},
            },
        ],
        "tasks": [
            {
                "id": "ss_parent01",
                "type": "color",
                "participant": "P01",
                "region": "hud",
                "region_ref": {"source": "stash", "stash_id": "stash_a", "name": "hud"},
                "region_coords": {"x": 20, "y": 20, "w": 20, "h": 20},
                "parameters": {
                    "target_color": {"h": 0, "s": 255, "v": 255},
                    "tolerance": {"h": 10, "s": 40, "v": 40},
                },
                "status": "completed",
            }
        ],
        "events": [],
    }
    saved_tasks = _install_ss_stubs(monkeypatch, fake_manifest)

    cli._run_ss_rerun_task(_ss_args(ss_run_task="ss_parent01"))

    assert saved_tasks
    # stash_a's hud -> {10,10,10,10}, not stash_b's {20,...} nor the stale saved coords.
    assert saved_tasks[0]["region_coords"] == {"x": 10, "y": 10, "w": 10, "h": 10}
