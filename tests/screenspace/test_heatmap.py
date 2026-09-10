"""Unit tests for the pure grid-heatmap helpers in screenspace_heatmap."""

import numpy as np

import config
import screenspace_heatmap as hm


def test_sprite_grid_caps_columns_and_rounds_rows_up(monkeypatch):
    monkeypatch.setattr(config, "SCREENSPACE_HEATMAP_SPRITE_COLS", 4)
    assert hm.sprite_grid(1) == (1, 1)
    assert hm.sprite_grid(4) == (4, 1)
    assert hm.sprite_grid(9) == (4, 3)


def test_grid_layer_count_is_bounded_by_gif_frames():
    assert hm.grid_layer_count([]) == 0
    assert hm.grid_layer_count([{}] * (hm.GIF_FRAMES + 5)) == hm.GIF_FRAMES


def test_build_grid_layers_draws_only_buckets_with_cells():
    heatmap_type, key = next(iter(hm._GRID_KEYS.items()))
    results = [
        {key: [{"x": 0.5, "y": 0.5, "mag": 0.9}]},
        {key: []},
    ]
    layers = hm.build_grid_layers(results, heatmap_type, num_frames=2)
    assert layers is not None and len(layers) == 2
    vals, mask = layers[0]
    assert mask.any() and float(vals[mask].max()) == np.float32(0.9)
    assert not layers[1][1].any()
    assert hm.build_grid_layers(results, "template", 2) is None
