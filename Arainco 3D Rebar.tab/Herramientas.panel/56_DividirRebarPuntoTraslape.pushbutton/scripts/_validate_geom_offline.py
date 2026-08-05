# -*- coding: utf-8 -*-
"""Validación offline (sin Revit) — geometría de corte + tabla de traslapes."""

from __future__ import print_function

import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from dividir_rebar_punto_geom import (
    build_plan_polyline_mm,
    build_spans_mm,
    fit_polyline_to_canvas,
    mm_to_internal,
    nearest_arc_length_px,
    overlap_length,
    piece_intervals_with_lap,
    point_at_arc_length_uv,
    set_span_length_mm,
    split_distances_with_lap,
    validate_cut_with_lap,
    validate_cuts_with_lap,
)

try:
    from dividir_rebar_punto_shapes import target_shape_names_for_pieces
except Exception:
    target_shape_names_for_pieces = None
from bimtools_rebar_hook_lengths import traslape_mm_from_nominal_diameter_mm


def _assert(cond, msg):
    if not cond:
        raise AssertionError(msg)


def test_table_base_and_grades():
    # Tabla base = H30 tracción superior
    _assert(traslape_mm_from_nominal_diameter_mm(16) == 1140.0, u"base ø16")
    _assert(traslape_mm_from_nominal_diameter_mm(25) == 2230.0, u"base ø25")
    _assert(traslape_mm_from_nominal_diameter_mm(16, u"G25") == 1110.0, u"G25 ø16")
    _assert(traslape_mm_from_nominal_diameter_mm(16, u"G35") == 940.0, u"G35 ø16")
    _assert(traslape_mm_from_nominal_diameter_mm(16, u"G45") == 820.0, u"G45 ø16")
    # redondeo ~8.2 mm → 8
    _assert(traslape_mm_from_nominal_diameter_mm(8.2) == 570.0, u"round 8.2")


def test_split_center_overlap():
    # Barra 12 m, lap 1.14 m (ø16 base), corte al centro
    L = 12000.0
    lap = 1140.0
    c = 6000.0
    ok, msg, half, la, lb = validate_cut_with_lap(L, c, lap, 100.0)
    _assert(ok, msg or u"validate center")
    _assert(abs(half - 570.0) < 1e-9, u"half")
    a, b, _ = split_distances_with_lap(L, c, lap)
    ov = overlap_length(a, b)
    _assert(abs(ov - lap) < 1e-6, u"overlap == lap, got {0}".format(ov))
    _assert(abs(la - (c + half)) < 1e-9, u"len_a")
    _assert(abs(lb - (L - c + half)) < 1e-9, u"len_b")


def test_split_near_end_rejects():
    L = 5000.0
    lap = 2230.0  # ø25
    half = lap / 2.0
    ok, msg, _, _, _ = validate_cut_with_lap(L, half - 1.0, lap, 100.0)
    _assert(not ok, u"should reject near start")
    ok2, _, _, _, _ = validate_cut_with_lap(L, L - half + 1.0, lap, 100.0)
    _assert(not ok2, u"should reject near end")


def test_units_mm_to_ft_roundtrip():
    lap_mm = 1140.0
    lap_ft = mm_to_internal(lap_mm)
    _assert(abs(lap_ft * 304.8 - lap_mm) < 1e-6, u"mm↔ft")


def test_multipoint_spans_and_pieces():
    L = 6318.0
    lap = 860.0
    ok, msg, cuts = validate_cuts_with_lap(L, [2100.0, 4200.0], lap, 100.0)
    _assert(ok, msg or u"two cuts")
    spans = build_spans_mm(L, cuts)
    _assert(len(spans) == 3, u"3 vanos")
    ssum = sum(s[u"length_mm"] for s in spans)
    _assert(abs(ssum - L) < 1e-6, u"vanos suman L")
    ints = piece_intervals_with_lap(L, cuts, lap)
    _assert(len(ints) == 3, u"3 tramos")
    ok2, msg2, cuts2 = set_span_length_mm(L, cuts, 0, 2000.0, lap, 100.0)
    _assert(ok2, msg2 or u"edit vano")
    _assert(abs(cuts2[0] - 2000.0) < 1e-6, u"cut moved")


def test_l_shape_plan_polyline():
    # Forma L en plano XZ (patas 3000 + 2000), normal = +Y (muro)
    pts = [
        (0.0, 0.0, 0.0),
        (0.0, 0.0, 3000.0),
        (2000.0, 0.0, 3000.0),
    ]
    plan = build_plan_polyline_mm(pts, normal=(0.0, 1.0, 0.0))
    _assert(plan is not None, u"plan")
    _assert(plan[u"plane"] == u"normal", u"plane normal, got {0}".format(plan[u"plane"]))
    _assert(abs(plan[u"total_mm"] - 5000.0) < 1e-6, u"L shape length")
    _assert(len(plan[u"points_uv"]) == 3, u"3 verts")
    # V debe crecer con Z: el segundo punto más alto que el primero
    _assert(plan[u"points_uv"][1][1] > plan[u"points_uv"][0][1], u"Z up on V")
    px, scale = fit_polyline_to_canvas(
        plan[u"points_uv"], 420, 480, 40, swap_uv=False, flip_v=True
    )
    _assert(len(px) == 3, u"px verts")
    _assert(scale > 0, u"scale")
    # Tras flip_v, el punto más alto en Z queda más arriba en pantalla (menor Y)
    _assert(px[1][1] < px[0][1], u"screen Y inverted for Z-up")
    mid = point_at_arc_length_uv(px, plan[u"arc_mm"], 1500.0)
    _assert(mid is not None, u"mid point")
    s, dist, _x, _y = nearest_arc_length_px(px, plan[u"arc_mm"], mid[0], mid[1])
    _assert(abs(s - 1500.0) < 1.0, u"roundtrip s={0}".format(s))
    _assert(dist < 1.0, u"on path")


def test_shape_rule_03_ends_middle():
    if target_shape_names_for_pieces is None:
        print(u"(omitido test shape: RevitAPI no disponible offline)")
        return
    _assert(
        target_shape_names_for_pieces(u"03", 2) == [u"02", u"02"],
        u"2 tramos",
    )
    _assert(
        target_shape_names_for_pieces(u"Shape 03", 3) == [u"02", u"01", u"02"],
        u"3 tramos",
    )
    _assert(
        target_shape_names_for_pieces(u"3", 4) == [u"02", u"01", u"01", u"02"],
        u"4 tramos",
    )
    _assert(target_shape_names_for_pieces(u"99", 3) is None, u"sin regla")


def test_endpoint_modes():
    L = 12000.0
    lap = 1140.0
    c = 6000.0
    ok, msg, cuts = validate_cuts_with_lap(
        L, [c], lap, 100.0, lap_mode=u"endpoint_prev"
    )
    _assert(ok, msg or u"endpoint_prev")
    ints = piece_intervals_with_lap(L, cuts, lap, lap_mode=u"endpoint_prev")
    _assert(abs(ints[0][1] - (c + lap)) < 1e-6, u"prev A end")
    _assert(abs(ints[1][0] - c) < 1e-6, u"prev B start")
    _assert(abs(overlap_length(ints[0], ints[1]) - lap) < 1e-6, u"prev overlap")
    ok2, msg2, cuts2 = validate_cuts_with_lap(
        L, [c], lap, 100.0, lap_mode=u"endpoint_next"
    )
    _assert(ok2, msg2 or u"endpoint_next")
    ints2 = piece_intervals_with_lap(L, cuts2, lap, lap_mode=u"endpoint_next")
    _assert(abs(ints2[0][1] - c) < 1e-6, u"next A end")
    _assert(abs(ints2[1][0] - (c - lap)) < 1e-6, u"next B start")
    _assert(abs(overlap_length(ints2[0], ints2[1]) - lap) < 1e-6, u"next overlap")
    ok3, _, _ = validate_cuts_with_lap(
        L, [lap - 1.0], lap, 100.0, lap_mode=u"endpoint_next"
    )
    _assert(not ok3, u"next cerca inicio debe fallar")
    ok4, _, _ = validate_cuts_with_lap(
        L, [L - lap + 1.0], lap, 100.0, lap_mode=u"endpoint_prev"
    )
    _assert(not ok4, u"prev cerca final debe fallar")


def main():
    test_table_base_and_grades()
    test_split_center_overlap()
    test_split_near_end_rejects()
    test_units_mm_to_ft_roundtrip()
    test_multipoint_spans_and_pieces()
    test_endpoint_modes()
    test_l_shape_plan_polyline()
    test_shape_rule_03_ends_middle()
    print(u"OK: geometría + tabla traslapes validadas offline.")


if __name__ == "__main__":
    main()
