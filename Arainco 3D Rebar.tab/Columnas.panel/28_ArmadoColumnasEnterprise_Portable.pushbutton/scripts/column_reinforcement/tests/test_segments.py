# -*- coding: utf-8 -*-

import os
import sys
import unittest


_SCRIPTS_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..")
)
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

from column_reinforcement.geometry.schemes import (  # noqa: E402
    is_a_split_scheme,
    is_b_split_scheme,
    is_lap_extension_scheme,
)
from column_reinforcement.geometry.segments import (  # noqa: E402
    apply_embedment_after_split,
    extend_intermediate_segments,
    split_z_span_by_cut_values,
)
from column_reinforcement.models.segments import SegmentZ  # noqa: E402


class SegmentGeometryTests(unittest.TestCase):
    def test_split_uses_core_span_only(self):
        segments = split_z_span_by_cut_values(0.0, 10.0, [-1.0, 4.0, 12.0], 0.01)
        self.assertEqual([s.as_tuple() for s in segments], [(0.0, 4.0), (4.0, 6.0)])

    def test_intermediate_lap_skips_last_segment(self):
        original = [SegmentZ(0.0, 4.0), SegmentZ(4.0, 6.0)]
        out = extend_intermediate_segments(original, 1.5, eligible=True)
        self.assertEqual([s.as_tuple() for s in out], [(0.0, 5.5), (4.0, 6.0)])

    def test_intermediate_lap_requires_eligible_scheme(self):
        original = [SegmentZ(0.0, 4.0), SegmentZ(4.0, 6.0)]
        out = extend_intermediate_segments(original, 1.5, eligible=False)
        self.assertEqual([s.as_tuple() for s in out], [(0.0, 4.0), (4.0, 6.0)])

    def test_embedment_is_applied_after_split_to_first_and_last(self):
        original = [SegmentZ(0.0, 4.0), SegmentZ(4.0, 6.0)]
        out = apply_embedment_after_split(
            original,
            top_lap_length=1.0,
            keep_top_embed=True,
            top_revoke_delta=0.0,
            keep_bottom_embed=True,
            bottom_revoke_delta=0.0,
            min_length=0.1,
        )
        self.assertEqual([s.as_tuple() for s in out], [(-1.0, 5.0), (4.0, 7.0)])

    def test_scheme_helpers_match_current_policy(self):
        self.assertTrue(is_lap_extension_scheme("A"))
        self.assertTrue(is_lap_extension_scheme("IA"))
        self.assertTrue(is_lap_extension_scheme("B"))
        self.assertTrue(is_lap_extension_scheme("IB"))
        self.assertTrue(is_a_split_scheme("A"))
        self.assertTrue(is_a_split_scheme("IA"))
        self.assertFalse(is_a_split_scheme("B"))
        self.assertFalse(is_a_split_scheme("IB"))
        self.assertTrue(is_b_split_scheme("B"))
        self.assertTrue(is_b_split_scheme("IB"))
        self.assertFalse(is_b_split_scheme("A"))


if __name__ == "__main__":
    unittest.main()
