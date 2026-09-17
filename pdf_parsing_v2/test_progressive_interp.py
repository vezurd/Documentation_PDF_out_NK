"""Unit tests for progressive interpolation helpers in grid_matcher.

Run:  python -m pytest pdf_parsing_v2/test_progressive_interp.py -v
"""

from __future__ import annotations

import pytest

from pdf_parsing_v2_engine.grid_matcher import _progressive_interp, _insert_anchor


# ---------------------------------------------------------------------------
# _progressive_interp
# ---------------------------------------------------------------------------

class TestProgressiveInterp:
    """Tests for piecewise-linear interpolation."""

    def test_uniform_scale(self):
        """Anchors with uniform 2x scale: interp should be exact."""
        anchors = [(0.0, 0.0), (100.0, 200.0)]
        assert _progressive_interp(50.0, anchors) == pytest.approx(100.0)
        assert _progressive_interp(25.0, anchors) == pytest.approx(50.0)
        assert _progressive_interp(75.0, anchors) == pytest.approx(150.0)

    def test_non_uniform_stretch(self):
        """Three anchors: first half compressed, second half stretched."""
        anchors = [(0.0, 0.0), (50.0, 30.0), (100.0, 100.0)]
        # Between anchors[0] and anchors[1]: scale = 30/50 = 0.6
        assert _progressive_interp(25.0, anchors) == pytest.approx(15.0)
        # Between anchors[1] and anchors[2]: scale = 70/50 = 1.4
        assert _progressive_interp(75.0, anchors) == pytest.approx(65.0)
        # Exactly at anchor
        assert _progressive_interp(50.0, anchors) == pytest.approx(30.0)

    def test_extrapolation_below(self):
        """pos_t below all anchors: extrapolate from first two."""
        anchors = [(10.0, 20.0), (20.0, 40.0)]
        # scale = 20/10 = 2.0, pos_t=5 -> 20 + (5-10)*2 = 10
        assert _progressive_interp(5.0, anchors) == pytest.approx(10.0)

    def test_extrapolation_above(self):
        """pos_t above all anchors: extrapolate from last two."""
        anchors = [(10.0, 20.0), (20.0, 40.0)]
        # scale = 2.0, pos_t=25 -> 20 + (25-10)*2 = 50
        assert _progressive_interp(25.0, anchors) == pytest.approx(50.0)

    def test_single_anchor(self):
        """One anchor: translate by offset."""
        anchors = [(100.0, 110.0)]
        assert _progressive_interp(50.0, anchors) == pytest.approx(60.0)
        assert _progressive_interp(150.0, anchors) == pytest.approx(160.0)

    def test_no_anchors(self):
        """No anchors: identity."""
        assert _progressive_interp(42.0, []) == pytest.approx(42.0)

    def test_many_anchors_picks_bracket(self):
        """With 5 anchors, interp at pos_t=35 should use bracket (30, 40)."""
        anchors = [
            (10.0, 10.0),
            (20.0, 22.0),
            (30.0, 36.0),
            (40.0, 52.0),
            (50.0, 70.0),
        ]
        # Between (30, 36) and (40, 52): scale = 16/10 = 1.6
        # pos_t=35 -> 36 + (35-30)*1.6 = 36 + 8 = 44
        assert _progressive_interp(35.0, anchors) == pytest.approx(44.0)

    def test_exactly_at_anchor(self):
        """pos_t exactly at an anchor returns that anchor's snapped pos."""
        anchors = [(10.0, 100.0), (20.0, 200.0), (30.0, 300.0)]
        assert _progressive_interp(20.0, anchors) == pytest.approx(200.0)

    def test_boundary_only_gives_global_affine(self):
        """Two boundary anchors = equivalent to old _local_interpolate."""
        anchors = [(500.0, 510.0), (800.0, 830.0)]
        # scale = 320/300 = 1.0667, offset
        for pos_t in [550.0, 600.0, 650.0, 700.0, 750.0]:
            expected = 510.0 + (pos_t - 500.0) * (320.0 / 300.0)
            assert _progressive_interp(pos_t, anchors) == pytest.approx(expected)


# ---------------------------------------------------------------------------
# _insert_anchor
# ---------------------------------------------------------------------------

class TestInsertAnchor:
    """Tests for anchor insertion with monotonicity guard."""

    def test_insert_between(self):
        anchors = [(0.0, 0.0), (100.0, 200.0)]
        ok = _insert_anchor(anchors, (50.0, 100.0))
        assert ok is True
        assert len(anchors) == 3
        assert anchors[1] == (50.0, 100.0)

    def test_reject_non_monotonic_snapped_below(self):
        """New snapped <= left neighbor's snapped: reject."""
        anchors = [(0.0, 50.0), (100.0, 200.0)]
        ok = _insert_anchor(anchors, (50.0, 40.0))
        assert ok is False
        assert len(anchors) == 2

    def test_reject_non_monotonic_snapped_above(self):
        """New snapped >= right neighbor's snapped: reject."""
        anchors = [(0.0, 50.0), (100.0, 200.0)]
        ok = _insert_anchor(anchors, (50.0, 210.0))
        assert ok is False
        assert len(anchors) == 2

    def test_insert_at_start(self):
        anchors = [(10.0, 20.0), (20.0, 40.0)]
        ok = _insert_anchor(anchors, (5.0, 10.0))
        assert ok is True
        assert anchors[0] == (5.0, 10.0)

    def test_insert_at_end(self):
        anchors = [(10.0, 20.0), (20.0, 40.0)]
        ok = _insert_anchor(anchors, (30.0, 60.0))
        assert ok is True
        assert anchors[-1] == (30.0, 60.0)

    def test_sequential_inserts_maintain_monotonicity(self):
        """Simulate a walk: sequential insertions all succeed."""
        anchors = [(0.0, 100.0), (100.0, 300.0)]
        positions = [(20.0, 140.0), (40.0, 170.0), (60.0, 210.0), (80.0, 260.0)]
        for t, s in positions:
            ok = _insert_anchor(anchors, (t, s))
            assert ok is True
        assert len(anchors) == 6
        snapped_vals = [a[1] for a in anchors]
        assert snapped_vals == sorted(snapped_vals)

    def test_equal_snapped_rejected(self):
        """Exactly equal snapped to neighbor: reject (>= check)."""
        anchors = [(0.0, 50.0), (100.0, 200.0)]
        ok = _insert_anchor(anchors, (50.0, 50.0))
        assert ok is False


# ---------------------------------------------------------------------------
# Integration: progressive walk simulation
# ---------------------------------------------------------------------------

class TestProgressiveWalkSimulation:
    """Simulate the progressive walk on synthetic data matching the log scenarios."""

    def test_non_uniform_redistribution(self):
        """Template has non-uniform rows, PDF has uniform rows.

        This mimics log 172457: template rows are [10, 10, 15, 10, 25, 10, ...]
        but PDF rows are all ~14.2.
        """
        # Template line positions (non-uniform gaps)
        tpl_positions = [0.0, 10.0, 20.0, 35.0, 45.0, 70.0, 80.0, 100.0]

        # Detected lines (uniform gaps in a proportionally scaled range)
        det_uniform = [0.0, 14.2, 28.4, 42.6, 56.8, 71.0, 85.2, 100.0]

        # Start with boundary anchors only (first and last)
        anchors = [(0.0, 0.0), (100.0, 100.0)]

        matched_count = 0
        max_delta = 0.0

        for i in range(1, len(tpl_positions) - 1):
            pos_t = tpl_positions[i]
            interp = _progressive_interp(pos_t, anchors)

            # Find closest detected line (simple nearest-neighbor)
            best_det = min(det_uniform, key=lambda d: abs(d - interp))
            delta = abs(best_det - interp)
            max_delta = max(max_delta, delta)

            if delta < 20.0:  # generous threshold
                _insert_anchor(anchors, (pos_t, best_det))
                matched_count += 1

        assert matched_count == 6  # all interior lines matched
        assert max_delta < 15.0  # progressive keeps deltas manageable

    def test_consecutive_no_match(self):
        """Three lines with no detected counterpart — anchors unchanged."""
        anchors = [(0.0, 0.0), (100.0, 200.0)]
        initial_len = len(anchors)

        # Three lines in a row with no match: just compute interp, don't insert
        for pos_t in [30.0, 40.0, 50.0]:
            interp = _progressive_interp(pos_t, anchors)
            # No match found — do NOT insert anchor
            assert interp == pytest.approx(pos_t * 2.0)  # still from boundary affine

        assert len(anchors) == initial_len  # no new anchors

        # Next line DOES match — progressive still works from boundaries
        interp = _progressive_interp(60.0, anchors)
        assert interp == pytest.approx(120.0)
        _insert_anchor(anchors, (60.0, 125.0))  # slight drift
        assert len(anchors) == 3

        # After the new anchor, subsequent interp adapts
        interp_70 = _progressive_interp(70.0, anchors)
        # Between (60, 125) and (100, 200): scale = 75/40 = 1.875
        expected = 125.0 + (70.0 - 60.0) * 1.875
        assert interp_70 == pytest.approx(expected)
