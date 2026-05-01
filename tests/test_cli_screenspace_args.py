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
        "yes": False,
        "verbose": False,
        "spreadsheet": None,
        "viewer": False,
        "manifest": False,
        "regenerate": False,
        "studio": False,
        "insights": False,
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
        "ollama_model": None,
        "summarize": None,
        "citations": None,
        "ss_task": None,
        "ss_list_regions": False,
        "ss_list_stashes": False,
        "ss_list_tasks": None,
        "ss_target_color": None,
        "ss_tolerance": None,
        "ss_threshold": None,
        "ss_reference_timestamp": None,
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


def test_parse_ss_task_color_minimal(monkeypatch):
    monkeypatch.setattr(
        "sys.argv",
        [
            "clipgen.py",
            "--ss-task",
            "color",
            "P01",
            "button",
            "--ss-target-color",
            "#FF0000",
            "--ss-tolerance",
            "20,30,30",
            "--ss-threshold",
            "0.85",
        ],
    )
    args = cli.parse_arguments()
    assert args.ss_task == ["color", "P01", "button"]
    assert args.ss_target_color == "#FF0000"
    assert args.ss_tolerance == "20,30,30"
    assert args.ss_threshold == 0.85


def test_parse_ss_list_tasks_with_status(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--ss-list-tasks", "completed"])
    args = cli.parse_arguments()
    assert args.ss_list_tasks == "completed"


def test_parse_ss_list_tasks_no_status(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--ss-list-tasks"])
    args = cli.parse_arguments()
    assert args.ss_list_tasks == ""  # const="" when no value


def test_ss_list_modes_are_mutually_exclusive(monkeypatch):
    monkeypatch.setattr(
        "sys.argv", ["clipgen.py", "--ss-list-regions", "--ss-list-stashes"]
    )
    with pytest.raises(SystemExit):
        cli.parse_arguments()


# ---- Conflict validation ----


def test_ss_task_conflicts_with_studio():
    args = _ss_args(
        ss_task=["color", "P01", "button"],
        ss_target_color="#FF0000",
        ss_tolerance="20,30,30",
        ss_threshold=0.85,
        studio=True,
    )
    with pytest.raises(SystemExit):
        cli._validate_mode_conflicts(args)


def test_summarize_conflicts_with_pre_transcribe():
    args = _ss_args(summarize=[], pre_transcribe=["P01"])
    with pytest.raises(SystemExit):
        cli._validate_mode_conflicts(args)


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


def test_ss_hex_to_hsv_invalid_length_raises():
    with pytest.raises(ValueError):
        cli._ss_hex_to_hsv("#ABC")


def test_ss_parse_tolerance_valid():
    tol = cli._ss_parse_tolerance("20,30,30")
    assert tol == {"h": 20, "s": 30, "v": 30}


def test_ss_parse_tolerance_wrong_count_raises():
    with pytest.raises(ValueError):
        cli._ss_parse_tolerance("20,30")


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
    monkeypatch.setattr(cli, "_ss_resolve_video_for_participant", lambda pid: None)

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
        cli, "_ss_resolve_video_for_participant", lambda pid: "/tmp/fake.mp4"
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
        cli, "_ss_resolve_video_for_participant", lambda pid: "/tmp/fake.mp4"
    )
    monkeypatch.setattr(
        video_mod,
        "probe_video_properties",
        lambda p: {"width": 100, "height": 100},
    )

    saved_tasks: list[dict] = []

    def fake_save(regions, tasks, events, stashes=None):
        saved_tasks.extend(tasks)
        return None

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
