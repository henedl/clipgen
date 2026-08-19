"""Tests for the Attention tool: saliency primitives, scan, heatmaps, events."""

import cv2
import numpy as np
import pytest

import config
import screenspace
import screenspace_frames
import screenspace_heatmap
import screenspace_primitives
import screenspace_scans


def _bright_patch_frame(w=160, h=120, px=120, py=80, size=20):
    """Flat dark frame with one bright high-contrast square patch."""
    frame = np.full((h, w, 3), 30, dtype=np.uint8)
    frame[py : py + size, px : px + size] = 245
    return frame


class TestSpectralResidual:
    def test_shape_range_and_repeatability(self):
        gray = np.random.RandomState(3).randint(0, 255, (90, 120), dtype=np.uint8)
        sal_a = screenspace.compute_spectral_residual(gray)
        sal_b = screenspace.compute_spectral_residual(gray.copy())
        assert sal_a.shape == gray.shape
        assert sal_a.dtype == np.float32
        assert float(sal_a.min()) >= 0.0 and float(sal_a.max()) <= 1.0
        # Repeatable to float32 noise rather than bit-identical: cv2.dft is not
        # bit-reproducible across calls on every OpenCV build (see
        # compute_spectral_residual). Anything that actually broke the transform
        # would move the map by orders of magnitude more than this.
        assert np.allclose(sal_a, sal_b, rtol=0, atol=1e-6)

    def test_flat_frame_does_not_crash(self):
        gray = np.full((60, 80), 128, dtype=np.uint8)
        sal = screenspace.compute_spectral_residual(gray)
        assert sal.shape == gray.shape
        assert np.all(np.isfinite(sal))


def _numpy_spectral_residual(gray):
    """The complex128 numpy transform, as the cv2.dft path's oracle."""
    fft = np.fft.fft2(gray.astype(np.float32))
    log_amp = np.log1p(np.abs(fft)).astype(np.float32)
    phase = np.angle(fft)
    residual = log_amp - cv2.blur(log_amp, (3, 3))
    sal = np.abs(np.fft.ifft2(np.exp(residual) * np.exp(1j * phase))) ** 2
    sal = cv2.GaussianBlur(sal.astype(np.float32), (9, 9), 2.5)
    peak = float(sal.max())
    return sal / peak if peak > 0 else sal


class TestSpectralResidualTransforms:
    """cv2.dft is a faster transform of the same math, not different math."""

    @pytest.mark.parametrize("shape", [(144, 256), (90, 120), (64, 64)])
    def test_dft_path_agrees_with_numpy(self, shape):
        assert screenspace_primitives._dft_friendly(shape)
        gray = np.random.RandomState(7).randint(0, 256, shape, dtype=np.uint8)
        assert np.allclose(
            screenspace.compute_spectral_residual(gray),
            _numpy_spectral_residual(gray),
            atol=1e-5,
        )

    def test_uniform_frame_agrees_with_numpy(self):
        # Every AC coefficient is exactly zero here, so the dft path's
        # "no direction" fallback (angle(0) == 0) is the whole result.
        for fill in (0, 7, 255):
            gray = np.full((64, 64), fill, dtype=np.uint8)
            assert np.allclose(
                screenspace.compute_spectral_residual(gray),
                _numpy_spectral_residual(gray),
                atol=1e-5,
            )

    def test_prime_dimensions_take_the_numpy_fallback(self):
        # cv2.dft is several times *slower* than numpy on these, so the gate
        # must reject them; the fallback is then bit-identical by definition.
        assert not screenspace_primitives._dft_friendly((151, 257))
        gray = np.random.RandomState(9).randint(0, 256, (151, 257), dtype=np.uint8)
        assert np.array_equal(
            screenspace.compute_spectral_residual(gray),
            _numpy_spectral_residual(gray),
        )

    def test_gate_accepts_only_2_3_5_smooth_sizes(self):
        assert screenspace_primitives._dft_friendly((120, 160))
        assert not screenspace_primitives._dft_friendly((139, 256))


class TestChannels:
    def test_color_contrast_surround_is_half_scale(self):
        # The surround blur is deliberately computed at half resolution (see
        # compute_color_contrast); pin it so it is not "corrected" back to the
        # full-scale kernel, which cost 5x more for a <=0.012 map difference.
        frame = _bright_patch_frame()
        lab = cv2.cvtColor(frame, cv2.COLOR_BGR2LAB).astype(np.float32)
        height, width = lab.shape[:2]
        sigma = max(3.0, max(height, width) / 8.0)
        small = cv2.resize(lab, (width // 2, height // 2), interpolation=cv2.INTER_AREA)
        surround = cv2.resize(
            cv2.GaussianBlur(small, (0, 0), sigma / 2.0),
            (width, height),
            interpolation=cv2.INTER_LINEAR,
        )
        diff = cv2.absdiff(lab, surround)
        expected = diff[:, :, 0] + diff[:, :, 1] + diff[:, :, 2]
        expected /= float(expected.max())
        assert np.array_equal(screenspace.compute_color_contrast(frame), expected)

    def test_color_contrast_handles_single_pixel_frame(self):
        contrast = screenspace.compute_color_contrast(
            np.full((1, 1, 3), 200, dtype=np.uint8)
        )
        assert contrast.shape == (1, 1)
        assert np.all(np.isfinite(contrast))

    def test_color_contrast_highlights_patch(self):
        frame = _bright_patch_frame()
        contrast = screenspace.compute_color_contrast(frame)
        assert contrast.shape == frame.shape[:2]
        patch_mean = float(contrast[80:100, 120:140].mean())
        background_mean = float(contrast[0:40, 0:40].mean())
        assert patch_mean > background_mean

    def test_motion_zero_without_prev(self):
        gray = np.full((60, 80), 128, dtype=np.uint8)
        motion = screenspace.compute_motion_saliency(None, gray)
        assert motion.shape == gray.shape
        assert float(motion.max()) == 0.0

    def test_motion_zero_on_shape_change(self):
        prev = np.full((50, 70), 128, dtype=np.uint8)
        curr = np.full((60, 80), 128, dtype=np.uint8)
        motion = screenspace.compute_motion_saliency(prev, curr)
        assert motion.shape == curr.shape
        assert float(motion.max()) == 0.0

    def test_motion_highlights_changed_area(self):
        prev = np.full((60, 80), 30, dtype=np.uint8)
        curr = prev.copy()
        curr[20:40, 40:60] = 240
        motion = screenspace.compute_motion_saliency(prev, curr)
        assert float(motion[25:35, 45:55].mean()) > float(motion[0:10, 0:10].mean())

    def test_face_saliency_zeros_without_faces(self):
        gray = np.full((60, 80), 128, dtype=np.uint8)
        blobs = screenspace.compute_face_saliency(gray)
        assert blobs.shape == gray.shape
        assert float(blobs.max()) == 0.0

    def test_face_saliency_degrades_without_cascade_api(self, monkeypatch):
        # opencv-python 5.x wheels removed the legacy CascadeClassifier; the
        # channel must degrade to a zeros map instead of raising.
        import cv2

        import screenspace_primitives

        monkeypatch.setattr(screenspace_primitives, "_face_cascade", None)
        monkeypatch.delattr(cv2, "CascadeClassifier", raising=False)
        gray = np.full((60, 80), 128, dtype=np.uint8)
        assert screenspace.face_detection_available() is False
        blobs = screenspace.compute_face_saliency(gray)
        assert blobs.shape == gray.shape
        assert float(blobs.max()) == 0.0

    def test_face_weight_left_out_of_mix_without_cascade_api(self, monkeypatch):
        # With include_face forced on but no cascade support, the face weight
        # must not join the denominator (a zeros-only channel would dim the
        # whole map).
        import cv2

        import screenspace_primitives

        frame = _bright_patch_frame()
        with_face_off, _ = screenspace.compute_saliency_map(
            frame, None, center_bias=0.0, include_face=False
        )
        monkeypatch.setattr(screenspace_primitives, "_face_cascade", None)
        monkeypatch.delattr(cv2, "CascadeClassifier", raising=False)
        with_face_on, _ = screenspace.compute_saliency_map(
            frame, None, center_bias=0.0, include_face=True
        )
        # Closeness, not bit-equality: cv2.dft is not bit-reproducible call to
        # call on every OpenCV build (Linux CI produced 1-ulp differences at
        # float32 where macOS produced none), and the two maps here are two
        # independent computations of the same math. The regression this guards
        # -- a zeros-only face channel joining the denominator -- scales the
        # whole map by 3/4: measured at 1.25e-1, five orders of magnitude above
        # this bound, against an observed wobble of 6.0e-8.
        assert np.allclose(with_face_off, with_face_on, rtol=0, atol=1e-6)


class TestSaliencyMap:
    def test_peak_lands_on_bright_patch(self):
        frame = _bright_patch_frame()
        sal, curr_gray = screenspace.compute_saliency_map(
            frame, None, center_bias=0.0, include_face=False
        )
        assert sal.shape == frame.shape[:2]
        assert curr_gray.shape == frame.shape[:2]
        px, py, pv = screenspace.saliency_peak(sal)
        # Patch spans x 120-140 of 160 (0.75-0.875), y 80-100 of 120 (0.67-0.83)
        assert 0.70 <= px <= 0.92
        assert 0.60 <= py <= 0.90
        assert pv > 0.0

    def test_center_bias_pulls_toward_center(self):
        frame = np.full((120, 160, 3), 30, dtype=np.uint8)
        # Two identical patches: one centered, one cornered
        frame[50:70, 70:90] = 245
        frame[0:20, 0:20] = 245
        sal, _ = screenspace.compute_saliency_map(
            frame, None, center_bias=0.9, include_face=False
        )
        px, py, _ = screenspace.saliency_peak(sal)
        assert 0.3 <= px <= 0.7
        assert 0.3 <= py <= 0.7

    def test_map_in_unit_range(self):
        frame = np.random.RandomState(5).randint(0, 255, (90, 120, 3), dtype=np.uint8)
        sal, _ = screenspace.compute_saliency_map(frame, None, include_face=False)
        assert float(sal.min()) >= 0.0
        assert float(sal.max()) <= 1.0


class TestSaliencyGrid:
    def test_grid_cells_shape_and_threshold(self):
        frame = _bright_patch_frame()
        sal, _ = screenspace.compute_saliency_map(
            frame, None, center_bias=0.0, include_face=False
        )
        grid = screenspace.saliency_grid_from_map(sal, 16, 0.15)
        assert grid
        assert len(grid) <= 16 * 16
        for cell in grid:
            assert set(cell.keys()) == {"x", "y", "mag"}
            assert 0.0 <= cell["x"] <= 1.0
            assert 0.0 <= cell["y"] <= 1.0
            assert 0.15 <= cell["mag"] <= 1.0

    def test_sparse_grid_cells_bit_identical_to_inline_rounding(self):
        """The memoized-center fast path must emit byte-identical cells.

        Reference is the formula it replaced: Python round() applied inline
        per cell (hot-path rule: exact equality, not tolerance).
        """
        rng = np.random.default_rng(7)
        for grid_n in (5, 16, 24):
            cells = rng.random((grid_n, grid_n), dtype=np.float32)
            got = screenspace.sparse_grid_cells(cells, 0.3)
            ys, xs = np.nonzero(cells >= 0.3)
            expected = [
                {
                    "x": round((int(x) + 0.5) / grid_n, 3),
                    "y": round((int(y) + 0.5) / grid_n, 3),
                    "mag": round(float(cells[y, x]), 3),
                }
                for y, x in zip(ys, xs)
            ]
            assert got == expected

    def test_grid_empty_on_zero_map(self):
        sal = np.zeros((60, 80), dtype=np.float32)
        assert screenspace.saliency_grid_from_map(sal, 16, 0.15) == []

    def test_grid_max_cell_near_patch(self):
        frame = _bright_patch_frame()
        sal, _ = screenspace.compute_saliency_map(
            frame, None, center_bias=0.0, include_face=False
        )
        grid = screenspace.saliency_grid_from_map(sal, 16, 0.15)
        best = max(grid, key=lambda c: c["mag"])
        assert best["mag"] == 1.0
        assert 0.70 <= best["x"] <= 0.92
        assert 0.60 <= best["y"] <= 0.90


class TestSaliencyKwargsFromParams:
    def test_empty_params_defer_to_config(self):
        assert screenspace.saliency_kwargs_from_params({}) == {}

    def test_partial_weights_fill_from_config(self):
        kwargs = screenspace.saliency_kwargs_from_params({"weight_motion": 2.0})
        assert kwargs["weights"]["motion"] == 2.0
        assert kwargs["weights"]["spectral"] == (
            config.SCREENSPACE_ATTENTION_WEIGHT_SPECTRAL
        )
        assert kwargs["weights"]["contrast"] == (
            config.SCREENSPACE_ATTENTION_WEIGHT_CONTRAST
        )
        assert "include_face" not in kwargs
        assert "center_bias" not in kwargs

    def test_face_weight_zero_disables_channel(self):
        kwargs = screenspace.saliency_kwargs_from_params({"weight_face": 0})
        assert kwargs["include_face"] is False
        kwargs = screenspace.saliency_kwargs_from_params({"weight_face": 0.8})
        assert kwargs["include_face"] is True

    def test_center_bias_passthrough(self):
        kwargs = screenspace.saliency_kwargs_from_params({"center_bias": 0.5})
        assert kwargs == {"center_bias": 0.5}

    def test_weights_reach_saliency_map(self):
        # Motion-only weights: a static bright patch scores ~zero without a
        # prev frame, proving the override actually reaches the channel mix.
        frame = _bright_patch_frame()
        kwargs = screenspace.saliency_kwargs_from_params(
            {
                "weight_spectral": 0,
                "weight_contrast": 0,
                "weight_motion": 1.0,
                "weight_face": 0,
                "center_bias": 0,
            }
        )
        sal, _ = screenspace.compute_saliency_map(frame, None, **kwargs)
        assert float(sal.max()) == 0.0


def _run_scan(monkeypatch, frames, **kwargs):
    """Drive scan_attention over a synthetic (ts, frame) sequence."""

    def fake_scan(video_path, interval, callback, **_kw):
        for ts, frame in frames:
            if callback(ts, frame) is False:
                break

    monkeypatch.setattr(screenspace_scans, "scan_video_full_frames", fake_scan)
    monkeypatch.setattr(
        screenspace_frames, "_probe_video_meta", lambda p: (30.0, float(len(frames)))
    )
    streamed = []
    results = screenspace.scan_attention(
        "/fake.mp4",
        None,
        interval_seconds=1.0,
        on_result=streamed.append,
        **kwargs,
    )
    return results, streamed


class TestScanAttention:
    def _patch_at(self, px, py):
        return _bright_patch_frame(w=160, h=120, px=px, py=py, size=20)

    def test_jumping_patch_emits_one_confirmed_shift(self, monkeypatch):
        frames = [(float(i), self._patch_at(10, 10)) for i in range(3)]
        frames += [(float(i), self._patch_at(130, 90)) for i in range(3, 7)]
        results, streamed = _run_scan(monkeypatch, frames, ema_alpha=1.0)

        assert len(results) == len(frames)
        for r in results:
            assert "saliency_grid" in r
            assert 0.0 <= r["peak_x"] <= 1.0
            assert 0.0 <= r["peak_y"] <= 1.0

        shifts = [r for r in results if r.get("shift")]
        assert len(shifts) == 1
        assert len(streamed) == 1
        shift = shifts[0]
        # Emitted on the frame where the jump began (confirmation back-stamps)
        assert shift["timestamp"] == 3.0
        # From the top-left patch to the bottom-right one
        assert shift["from_x"] < 0.35 and shift["from_y"] < 0.45
        assert shift["to_x"] > 0.65 and shift["to_y"] > 0.55
        assert shift["shift_distance"] > 0.5
        assert 0.05 <= shift["_confidence"] <= 1.0

    def test_streamed_payload_omits_grid(self, monkeypatch):
        frames = [(float(i), self._patch_at(10, 10)) for i in range(3)]
        frames += [(float(i), self._patch_at(130, 90)) for i in range(3, 7)]
        results, streamed = _run_scan(monkeypatch, frames, ema_alpha=1.0)
        assert streamed
        assert "saliency_grid" not in streamed[0]
        assert streamed[0]["shift"] is True
        # The returned result for the same frame keeps its grid
        match = [r for r in results if r["timestamp"] == streamed[0]["timestamp"]]
        assert match and "saliency_grid" in match[0]

    def test_static_sequence_emits_no_shifts(self, monkeypatch):
        frames = [(float(i), self._patch_at(70, 50)) for i in range(6)]
        results, streamed = _run_scan(monkeypatch, frames)
        assert len(results) == len(frames)
        assert streamed == []
        assert not any(r.get("shift") for r in results)

    def test_single_sample_jump_is_not_confirmed(self, monkeypatch):
        # Peak visits the far corner for one sample only, then returns: the
        # confirm counter (default 2) must reject the blip.
        frames = [(float(i), self._patch_at(10, 10)) for i in range(3)]
        frames.append((3.0, self._patch_at(130, 90)))
        frames += [(float(i), self._patch_at(10, 10)) for i in range(4, 7)]
        _, streamed = _run_scan(monkeypatch, frames, ema_alpha=1.0)
        assert streamed == []

    def test_cancel_stops_scan(self, monkeypatch):
        frames = [(float(i), self._patch_at(10, 10)) for i in range(10)]

        def fake_scan(video_path, interval, callback, **_kw):
            for ts, frame in frames:
                if callback(ts, frame) is False:
                    break

        monkeypatch.setattr(screenspace_scans, "scan_video_full_frames", fake_scan)
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 10.0)
        )
        calls = [0]

        def cancel():
            calls[0] += 1
            return calls[0] > 3

        results = screenspace.scan_attention("/fake.mp4", cancel_flag=cancel)
        assert len(results) < len(frames)


def _grid_results(n=8):
    """Synthetic per-frame results with a saliency_grid hot spot."""
    return [
        {
            "timestamp": float(i),
            "saliency_grid": [
                {"x": 0.25, "y": 0.25, "mag": 1.0},
                {"x": 0.75, "y": 0.75, "mag": 0.4},
            ],
            "peak_x": 0.25,
            "peak_y": 0.25,
            "peak_value": 0.8,
        }
        for i in range(n)
    ]


def _varied_grid_results(n):
    """Per-frame results whose grids overlap and *differ* frame to frame.

    ``_grid_results`` repeats one grid, so every draw order produces the same
    pixels — useless for proving that bucketed layers replay in the right order.
    These cells move and change magnitude, so an out-of-order fold shows up.
    """
    rng = np.random.RandomState(11)
    results = []
    for i in range(n):
        cells = [
            {
                "x": round(float((gx + 0.5) / 8), 3),
                "y": round(float((gy + 0.5) / 8), 3),
                "mag": round(float(rng.rand()), 3),
            }
            for gy in range(8)
            for gx in range(8)
            if rng.rand() > 0.3
        ]
        results.append({"timestamp": float(i), "saliency_grid": cells})
    return results


class TestAttentionHeatmap:
    def test_static_png_generated(self, tmp_path):
        out = str(tmp_path / "heatmap.png")
        path = screenspace.generate_attention_heatmap(_grid_results(), 320, 240, out)
        assert path == out
        import cv2

        img = cv2.imread(out)
        assert img is not None
        assert img.shape[:2] == (240, 320)

    def test_empty_grids_yield_none(self, tmp_path):
        out = str(tmp_path / "heatmap.png")
        results = [{"timestamp": 0.0, "saliency_grid": []}]
        assert screenspace.generate_attention_heatmap(results, 320, 240, out) is None

    def test_cumulative_and_rolling_gifs(self, tmp_path):
        cum = str(tmp_path / "cum.gif")
        roll = str(tmp_path / "roll.gif")
        cum_info = screenspace.generate_heatmap_gif(
            _grid_results(), 160, 120, cum, heatmap_type="attention"
        )
        roll_info = screenspace.generate_rolling_heatmap_gif(
            _grid_results(), 160, 120, roll, heatmap_type="attention"
        )
        assert cum_info is not None and cum_info["path"] == cum
        assert roll_info is not None and roll_info["path"] == roll


class TestSharedGridLayers:
    """The prebuilt-layers fast path must be byte-identical to replaying results.

    The worker draws each GIF bucket's grid cells once and hands the layers to
    the PNG and both GIFs instead of letting all three replay the results
    (measured 2.84 s → 1.23 s of post-scan work on a 962-result attention scan).
    That is only sound because ``cv2.circle`` sets rather than accumulates, so
    equality here is exact — a tolerance would hide exactly the drift this
    guards.
    """

    def _artifacts(self, tmp_path, results, tag, layers):
        import screenspace_heatmap

        png = tmp_path / f"{tag}.png"
        cum = tmp_path / f"{tag}_cum.gif"
        roll = tmp_path / f"{tag}_roll.gif"
        screenspace.generate_attention_heatmap(
            results, 320, 240, str(png), layers=layers
        )
        screenspace_heatmap.generate_heatmap_gif(
            results, 320, 240, str(cum), heatmap_type="attention", layers=layers
        )
        screenspace_heatmap.generate_rolling_heatmap_gif(
            results,
            320,
            240,
            str(roll),
            heatmap_type="attention",
            window_frames=3,
            layers=layers,
        )
        return [p.read_bytes() if p.exists() else None for p in (png, cum, roll)]

    @pytest.mark.parametrize("count", [1, 2, 8, 37])
    def test_layers_path_is_byte_identical_to_replay(self, tmp_path, count):
        results = _varied_grid_results(count)
        layers = screenspace.build_grid_layers(
            results, "attention", screenspace.grid_layer_count(results)
        )
        replayed = self._artifacts(tmp_path, results, "replay", None)
        shared = self._artifacts(tmp_path, results, "shared", layers)
        assert replayed[0] is not None  # the PNG always lands
        assert shared == replayed

    def test_layer_count_matches_gif_buckets(self):
        assert (
            screenspace.grid_layer_count(_grid_results(100)) == screenspace.GIF_FRAMES
        )
        assert screenspace.grid_layer_count(_grid_results(5)) == 5

    def test_layers_tile_every_result(self):
        """Folding all buckets equals a full replay — the PNG depends on it."""
        results = _varied_grid_results(37)
        layers = screenspace.build_grid_layers(
            results, "attention", screenspace.grid_layer_count(results)
        )
        assert layers is not None
        drawn = np.zeros((256, 256), dtype=bool)
        for _vals, mask in layers:
            drawn |= mask
        replay = np.zeros((256, 256), dtype=np.float32)
        for r in results:
            screenspace_heatmap._accumulate_heatmap_result(replay, r, "attention")
        assert np.array_equal(
            screenspace_heatmap._fold_grid_layers(layers, len(layers) - 1), replay
        )
        assert drawn.any()

    def test_non_grid_type_has_no_layers(self):
        """Template accumulates additively and frame-native — never bucketed."""
        assert screenspace.build_grid_layers([{"matches": []}], "template", 4) is None

    def test_stale_layers_are_rejected_not_trusted(self, tmp_path):
        """A layer set bucketed for a different result count must not be used."""
        results = _varied_grid_results(24)
        wrong = screenspace.build_grid_layers(results[:6], "attention", 6)
        out = tmp_path / "cum.gif"
        ref = tmp_path / "ref.gif"
        screenspace_heatmap.generate_heatmap_gif(
            results, 320, 240, str(out), heatmap_type="attention", layers=wrong
        )
        screenspace_heatmap.generate_heatmap_gif(
            results, 320, 240, str(ref), heatmap_type="attention"
        )
        assert out.read_bytes() == ref.read_bytes()


class TestAttentionEvents:
    def _task(self):
        return {
            "id": "ss_attn0001",
            "type": "attention",
            "participant": "P01",
            "region": "full_frame",
            "parameters": {},
        }

    def _shift_result(self):
        return {
            "timestamp": 4.0,
            "peak_x": 0.8,
            "peak_y": 0.7,
            "peak_value": 0.9,
            "shift": True,
            "shift_distance": 0.5,
            "_confidence": 0.77,
            "from_x": 0.2,
            "from_y": 0.3,
            "to_x": 0.8,
            "to_y": 0.7,
        }

    def test_shift_metadata_and_confidence(self):
        events = screenspace.generate_events_from_results(
            self._task(), [self._shift_result()]
        )
        assert len(events) == 1
        ev = events[0]
        assert ev["detector"] == "attention"
        assert ev["confidence"] == 0.77
        assert ev["metadata"]["shift_distance"] == 0.5
        assert ev["metadata"]["from_x"] == 0.2
        assert ev["metadata"]["to_x"] == 0.8
        assert ev["metadata"]["peak_value"] == 0.9

    def test_non_shift_samples_filtered(self):
        # Regeneration paths can hand the full per-sample list; only shift
        # results may become events.
        raw = _grid_results(5) + [self._shift_result()]
        events = screenspace.generate_events_from_results(self._task(), raw)
        assert len(events) == 1
        assert events[0]["time_in"] == 4.0

    def test_manifest_strips_saliency_grid(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        tasks = [
            {
                "id": "ss_attn1234",
                "type": "attention",
                "participant": "P01",
                "status": "completed",
                "result": [
                    {
                        "timestamp": 1.0,
                        "peak_x": 0.5,
                        "peak_y": 0.5,
                        "peak_value": 0.6,
                        "saliency_grid": [{"x": 0.5, "y": 0.5, "mag": 1.0}],
                    }
                ],
            }
        ]
        path = screenspace.save_screenspace_manifest({}, tasks)
        assert path is not None
        loaded = screenspace.load_screenspace_manifest()
        result = loaded["tasks"][0]["result"][0]
        assert "saliency_grid" not in result
        assert result["peak_value"] == 0.6

    def test_describe_shows_threshold(self):
        name = screenspace.describe_task("attention", "full_frame", {})
        assert "Attention" in name
        assert "Δ≥0.2" in screenspace.describe_task(
            "attention", "full_frame", {"shift_threshold": 0.2}
        )


class TestAttentionToolParamMapping:
    def test_scan_forwards_weight_params(self, monkeypatch):
        import screenspace_tools

        captured = {}

        def fake_scan(video_path, region, **kwargs):
            captured.update(kwargs)
            return []

        monkeypatch.setattr(screenspace_tools, "scan_attention", fake_scan)
        screenspace.TOOLS["attention"].scan(
            "/v.mp4",
            {"x": 0, "y": 0, "w": 0, "h": 0},
            {
                "shift_threshold": 0.2,
                "weight_motion": 2.0,
                "weight_face": 0,
                "center_bias": 0.1,
            },
            task_id="ss_x",
            scan_mode="full",
            on_progress=lambda _p: None,
            cancel_flag=lambda: False,
            on_result=None,
            fast_opts=None,
        )
        assert captured["shift_threshold"] == 0.2
        assert captured["weights"]["motion"] == 2.0
        assert captured["include_face"] is False
        assert captured["center_bias"] == 0.1

    def test_scan_omits_saliency_kwargs_when_untuned(self, monkeypatch):
        import screenspace_tools

        captured = {}

        def fake_scan(video_path, region, **kwargs):
            captured.update(kwargs)
            return []

        monkeypatch.setattr(screenspace_tools, "scan_attention", fake_scan)
        screenspace.TOOLS["attention"].scan(
            "/v.mp4",
            {"x": 0, "y": 0, "w": 0, "h": 0},
            {},
            task_id="ss_x",
            scan_mode="full",
            on_progress=lambda _p: None,
            cancel_flag=lambda: False,
            on_result=None,
            fast_opts=None,
        )
        assert "weights" not in captured
        assert "include_face" not in captured
        assert "center_bias" not in captured


class TestAttentionWorker:
    def test_paused_task_stores_only_shifts(self, monkeypatch):
        # Pausing must not expose the full per-sample stream: the paused
        # t["result"] is what the results/timeline API serves (and what resume
        # re-seeds _partial_results from), so it follows the shift-only
        # contract the live on_result stream follows.
        samples = _grid_results(6)
        samples[3].update({"shift": True, "shift_distance": 0.5, "_confidence": 0.7})

        def pausing_dispatch(self, task, on_progress, cancel_flag, on_result=None):
            with self._lock:
                self._tasks[task["id"]]["_paused_flag"] = True
            return [dict(s) for s in samples]

        monkeypatch.setattr(
            screenspace.ScreenspaceWorker, "_dispatch", pausing_dispatch
        )
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "attention",
            "P01",
            "s_P01.mp4",
            ["/v.mp4"],
            "full_frame",
            {"x": 0, "y": 0, "w": 0, "h": 0},
        )
        worker.enqueue(task)
        worker.start()
        try:
            import time

            for _ in range(100):
                t = worker.get_task(task["id"])
                if t and t["status"] == screenspace.TASK_STATUS_PAUSED:
                    break
                time.sleep(0.05)
            with worker._lock:
                t = worker._tasks[task["id"]]
                assert t["status"] == screenspace.TASK_STATUS_PAUSED
                assert len(t["result"]) == 1
                assert t["result"][0]["shift"] is True
        finally:
            worker.stop()

    def test_completion_is_shift_only_even_with_heatmaps_disabled(
        self, tmp_path, monkeypatch
    ):
        # The shift filter runs at result assignment (inside the completion
        # lock), not as a post-heatmap afterthought — so the completed→heatmap
        # window and the heatmaps-off path never expose per-sample rows.
        import time

        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        monkeypatch.setattr(config, "SCREENSPACE_GENERATE_ATTENTION_HEATMAP", False)

        samples = _grid_results(6)
        samples[4].update({"shift": True, "shift_distance": 0.6, "_confidence": 0.8})

        def fake_dispatch(self, task, on_progress, cancel_flag, on_result=None):
            return [dict(s) for s in samples]

        monkeypatch.setattr(screenspace.ScreenspaceWorker, "_dispatch", fake_dispatch)
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "attention",
            "P01",
            "s_P01.mp4",
            ["/v.mp4"],
            "full_frame",
            {"x": 0, "y": 0, "w": 0, "h": 0},
        )
        worker.enqueue(task)
        worker.start()
        try:
            for _ in range(100):
                t = worker.get_task(task["id"])
                if t and t["status"] == "completed":
                    break
                time.sleep(0.05)
            t = worker.get_task(task["id"])
            assert t is not None
            assert t["status"] == "completed"
            assert len(t["result"]) == 1
            assert t["result"][0]["shift"] is True
            assert not t.get("heatmap")
        finally:
            worker.stop()

    def test_completion_filters_results_to_shifts_and_attaches_heatmaps(
        self, tmp_path, monkeypatch
    ):
        import time

        import video

        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        monkeypatch.setattr(
            video,
            "probe_video_properties",
            lambda p: {"width": 320, "height": 240},
        )

        samples = _grid_results(6)
        shift = {
            "shift": True,
            "shift_distance": 0.6,
            "_confidence": 0.8,
            "from_x": 0.25,
            "from_y": 0.25,
            "to_x": 0.75,
            "to_y": 0.75,
        }
        samples[4].update(shift)

        def fake_dispatch(self, task, on_progress, cancel_flag, on_result=None):
            if on_result:
                stream = {k: v for k, v in samples[4].items() if k != "saliency_grid"}
                on_result(stream)
            return [dict(s) for s in samples]

        monkeypatch.setattr(screenspace.ScreenspaceWorker, "_dispatch", fake_dispatch)
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "attention",
            "P01",
            "s_P01.mp4",
            ["/v.mp4"],
            "full_frame",
            {"x": 0, "y": 0, "w": 0, "h": 0},
        )
        worker.enqueue(task)
        worker.start()
        try:
            for _ in range(100):
                t = worker.get_task(task["id"])
                if t and t["status"] == "completed" and t.get("heatmap"):
                    break
                time.sleep(0.05)
            t = worker.get_task(task["id"])
            assert t is not None
            assert t["status"] == "completed"
            # Post-heatmap: visible results are the confirmed shifts only
            assert len(t["result"]) == 1
            assert t["result"][0]["shift"] is True
            assert "saliency_grid" not in t["result"][0]
            # Heatmap artifacts were generated from the full sample stream
            assert t.get("heatmap")
            assert (tmp_path / t["heatmap"]).exists()
            assert t.get("heatmap_gif")
            assert t.get("heatmap_rolling_gif")
            # Events derive from the shift-only on_result stream
            events = t.get("_generated_events", [])
            assert len(events) == 1
            assert events[0]["detector"] == "attention"
        finally:
            worker.stop()
