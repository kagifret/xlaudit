"""Tests for xlaudit analysis — complexity scoring."""

import pytest

from xlaudit.analysis import classify_band, compute_complexity


class TestComputeComplexity:
    def test_zero_inputs(self):
        score = compute_complexity(
            external_links=0,
            volatile_functions=0,
            cross_sheet_refs=0,
            formula_density=0.0,
            sheet_count=0,
            named_ranges=0,
        )
        assert score == 0.0

    def test_known_weights(self):
        score = compute_complexity(
            external_links=1,    # 1 × 5.0 = 5.0
            volatile_functions=1, # 1 × 2.0 = 2.0
            cross_sheet_refs=10, # 10 × 0.1 = 1.0
            formula_density=0.5, # 0.5 × 15.0 = 7.5
            sheet_count=2,       # 2 × 0.5 = 1.0
            named_ranges=5,      # 5 × 0.3 = 1.5
        )
        assert score == 18.0

    def test_high_external_links(self):
        score = compute_complexity(
            external_links=10,
            volatile_functions=0,
            cross_sheet_refs=0,
            formula_density=0.0,
            sheet_count=1,
            named_ranges=0,
        )
        # 10 × 5.0 + 1 × 0.5 = 50.5
        assert score == 50.5


class TestClassifyBand:
    def test_low(self):
        assert classify_band(0) == "LOW"
        assert classify_band(9.9) == "LOW"

    def test_med(self):
        assert classify_band(10) == "MED"
        assert classify_band(25) == "MED"

    def test_high(self):
        assert classify_band(25.1) == "HIGH"
        assert classify_band(100) == "HIGH"
