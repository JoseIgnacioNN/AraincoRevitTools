# -*- coding: utf-8 -*-
"""
Geometría pura — Barras de retorno de malla (sin Revit).

Longitudinales en espesor de muro (≥2), Z fondo fundación o pie de muro,
extremos Pata 90º / Empotramiento, empalme A-B-A como Coronamiento.
"""

from __future__ import print_function

import math

MAX_BARRA_COMERCIAL_MM = 12000.0
COVER_WALL_MM = 25.0
COVER_FUND_BOT_MM = 50.0
COVER_WALL_BOT_MM = 50.0
MARGIN_END_MM = 50.0
CUT_HIT_MM = 180.0
STAGGER_MARGIN_MM = 300.0
LAYER_CENTERLINE_SPACING_MM = 50.0
PATA_MIN_MM = 100.0

DIAMS_MM = (8, 10, 12, 16, 18, 20)
N_BARS_OPTS = (2, 3, 4)
CAPAS_OPTS = (1, 2, 3)

DOSIFICACION_HORMIGON_OPCIONES = (u"G25", u"G35", u"G45")
DOSIFICACION_HORMIGON_DEFAULT = u"G25"

END_PATA = u"pata_l"
END_EMPOTRO = u"empotramiento"
END_OPTS = (
    (END_PATA, u"Pata 90º"),
    (END_EMPOTRO, u"Empotramiento"),
)

LAP_MODE_SYMMETRIC = u"symmetric"
LAP_MODE_ENDPOINT_PREV = u"endpoint_prev"
LAP_MODE_ENDPOINT_NEXT = u"endpoint_next"

LAP_MODE_LABELS = (
    (LAP_MODE_SYMMETRIC, u"Simétrico ± L/2"),
    (LAP_MODE_ENDPOINT_PREV, u"Endpoint · estira tramo anterior (+L)"),
    (LAP_MODE_ENDPOINT_NEXT, u"Endpoint · estira tramo siguiente (−L)"),
)

_TRASLAPE_BY_DIAM = {
    8: 570.0,
    10: 710.0,
    12: 860.0,
    16: 1140.0,
    18: 1290.0,
    20: 1430.0,
    22: 1960.0,
    25: 2230.0,
}


def ceil10_mm(mm):
    try:
        v = float(mm)
    except Exception:
        return 0.0
    if v <= 1e-9:
        return 0.0
    return float(int(math.ceil(v / 10.0) * 10))


def clamp_n_bars(n):
    try:
        v = int(round(float(n)))
    except Exception:
        v = 2
    return max(2, min(4, v))


def clamp_diam_mm(d):
    try:
        v = int(round(float(d)))
    except Exception:
        v = 10
    if v in DIAMS_MM:
        return v
    return 10


def clamp_n_capas(n):
    try:
        v = int(round(float(n)))
    except Exception:
        v = 1
    return max(1, min(3, v))


def normalize_concrete_grade(concrete_grade):
    try:
        s = u"{0}".format(concrete_grade or u"").strip().upper()
    except Exception:
        s = u""
    if s in DOSIFICACION_HORMIGON_OPCIONES:
        return s
    return DOSIFICACION_HORMIGON_DEFAULT


def normalize_end_condition(end):
    try:
        s = u"{0}".format(end or u"").strip().lower()
    except Exception:
        s = u""
    if s in (u"pata", u"pata_l", u"pata90", u"pata_90", u"hook", u"l"):
        return END_PATA
    if s in (u"empotramiento", u"empotrado", u"embed", u"empotro"):
        return END_EMPOTRO
    return END_EMPOTRO


def end_label(end):
    e = normalize_end_condition(end)
    for key, lab in END_OPTS:
        if key == e:
            return lab
    return u"Empotramiento"


def traslape_mm_from_diam(diam_mm, concrete_grade=None):
    d = clamp_diam_mm(diam_mm)
    g = None
    if concrete_grade is not None:
        g = normalize_concrete_grade(concrete_grade)
    try:
        from bimtools_rebar_hook_lengths import traslape_mm_from_nominal_diameter_mm

        t = traslape_mm_from_nominal_diameter_mm(d, g)
        if t is not None and float(t) > 0:
            return float(t)
    except Exception:
        pass
    if d in _TRASLAPE_BY_DIAM:
        return float(_TRASLAPE_BY_DIAM[d])
    return 710.0


def empotramiento_mm_from_diam(diam_mm, concrete_grade=None):
    return traslape_mm_from_diam(diam_mm, concrete_grade)


def pata_mm_from_diam(diam_mm, concrete_grade=None, bar_type=None):
    """
    Longitud de **eje** para pata 90º (tabla BIMTools − Ø/2).

    Usa el nominal del ``RebarBarType`` cuando se entrega y la dosificación G25/G35/G45.
    Sin compensación de eje, Revit suele reportar ~tabla + Ø/2 (p. ej. 330 en vez de 320).
    """
    d = clamp_diam_mm(diam_mm)
    grade = normalize_concrete_grade(concrete_grade) if concrete_grade is not None else None
    d_nom = float(d)
    if bar_type is not None:
        try:
            import armado_muros_coronamiento as cor

            resolved = cor._nominal_diameter_bar_type_mm(bar_type, d)
            if resolved is not None and float(resolved) > 0.1:
                d_nom = float(resolved)
        except Exception:
            pass
    try:
        import armado_muros_cabezal as cabezal

        leje = cabezal._pata_l_eje_sketch_mm_desde_diametro(d_nom, grade)
        if leje is not None and float(leje) > 0.1:
            return max(PATA_MIN_MM, float(leje))
    except Exception:
        pass
    try:
        from bimtools_rebar_hook_lengths import (
            hook_length_mm_from_nominal_diameter_mm,
            pata_eje_curve_loop_mm_desde_tabla_mm,
        )

        tab = hook_length_mm_from_nominal_diameter_mm(d_nom, grade)
        if tab is not None and float(tab) > 0:
            eje = pata_eje_curve_loop_mm_desde_tabla_mm(float(tab), float(d_nom))
            if eje is not None and float(eje) > 0:
                return max(PATA_MIN_MM, float(eje))
    except Exception:
        pass
    return max(PATA_MIN_MM, 12.0 * float(d_nom))


def long_bar_length_mm(wall_length_mm, margin_end_mm=None):
    """Longitud del tramo longitudinal entre holguras de extremo."""
    m = float(MARGIN_END_MM if margin_end_mm is None else margin_end_mm)
    try:
        L = float(wall_length_mm)
    except Exception:
        L = 0.0
    return max(0.0, ceil10_mm(L - 2.0 * m))


def layer_bar_offsets_mm(thickness_mm, n_bars, cover_mm=None):
    """
    Offsets transversales (mm desde cara −Orientation) dentro del espesor del muro.
    n ≥ 2; extremos a COVER_WALL_MM de cada cara (banda usable = e − 2·cover).
    No usar ancho de fundación / voladizo.
    """
    cover = float(COVER_WALL_MM if cover_mm is None else cover_mm)
    try:
        t = float(thickness_mm)
    except Exception:
        t = 200.0
    cover = min(cover, t * 0.5)
    x0 = cover
    x1 = t - cover
    span = max(0.0, x1 - x0)
    mid = t * 0.5
    n = clamp_n_bars(n_bars)
    out = []
    for i in range(n):
        if n <= 1:
            out.append(mid)
        else:
            tt = float(i) / float(n - 1)
            out.append(x0 + tt * span if span > 0 else mid)
    return out


def cover_axis_offset_mm_for_layer(layers, layer_index, base_cover_mm=None):
    """
    Offset vertical desde cara de referencia (fondo) hasta eje de capa.

    Capa 0: base_cover + Ø0/2.
    Capas siguientes: + LAYER_CENTERLINE_SPACING_MM entre ejes.
    """
    base = float(COVER_FUND_BOT_MM if base_cover_mm is None else base_cover_mm)
    if not layers:
        return base + 5.0
    li = max(0, min(int(layer_index), len(layers) - 1))
    d0 = float(clamp_diam_mm(layers[0].get(u"diam_mm", 10)))
    z = base + 0.5 * d0
    for _i in range(1, li + 1):
        z += float(LAYER_CENTERLINE_SPACING_MM)
    return float(z)


def sync_layers(prev_layers, n_capas):
    n = clamp_n_capas(n_capas)
    out = []
    prev = list(prev_layers or [])
    for i in range(n):
        if i < len(prev) and isinstance(prev[i], dict):
            out.append(
                {
                    u"n_bars": clamp_n_bars(prev[i].get(u"n_bars", 2)),
                    u"diam_mm": clamp_diam_mm(prev[i].get(u"diam_mm", 10)),
                }
            )
        else:
            out.append({u"n_bars": 2, u"diam_mm": 10})
    return out


def normalize_lap_mode_ui(mode):
    try:
        key = u"{0}".format(mode or u"").strip().lower()
    except Exception:
        key = u""
    aliases = {
        u"symmetric": LAP_MODE_SYMMETRIC,
        u"simetrico": LAP_MODE_SYMMETRIC,
        u"simétrico": LAP_MODE_SYMMETRIC,
        u"endpoint_prev": LAP_MODE_ENDPOINT_PREV,
        u"prev": LAP_MODE_ENDPOINT_PREV,
        u"anterior": LAP_MODE_ENDPOINT_PREV,
        u"endpoint_next": LAP_MODE_ENDPOINT_NEXT,
        u"next": LAP_MODE_ENDPOINT_NEXT,
        u"siguiente": LAP_MODE_ENDPOINT_NEXT,
    }
    return aliases.get(key, LAP_MODE_SYMMETRIC)


def to_dividir_lap_mode(lap_mode_ui):
    return normalize_lap_mode_ui(lap_mode_ui)


def stagger_cuts_for_layer(cuts_ref_mm, layer_index, main_mm, lap_mm):
    """
    Alternancia A-B-A (mismo criterio que Coronamiento).

    Capas pares: mismos cortes. Capas impares: desfase ±L_traslape hacia midspan.
    """
    cuts = []
    for c in cuts_ref_mm or []:
        try:
            cuts.append(float(c))
        except Exception:
            continue
    cuts = sorted(set(cuts))
    if not cuts:
        return []
    if int(layer_index) % 2 == 0:
        return list(cuts)

    L = float(main_mm)
    if L <= 2.0:
        return list(cuts)
    lap = float(lap_mm or 0.0)
    delta = float(lap) if lap >= 1.0 else float(STAGGER_MARGIN_MM)
    lo = max(1.0, float(STAGGER_MARGIN_MM))
    hi = max(lo, L - lo)
    mid = 0.5 * L
    out = []
    for c in cuts:
        sign = 1.0 if c < mid else -1.0
        cand = c + sign * delta
        if cand < lo or cand > hi:
            cand = c - sign * delta
        cand = max(lo, min(hi, cand))
        out.append(ceil10_mm(cand))
    return sorted(set(out))


def toggle_cut_at_mm(cuts_mm, click_mm, main_mm):
    try:
        x = float(click_mm)
        L = float(main_mm)
    except Exception:
        return list(cuts_mm or [])
    if L <= 1.0:
        return list(cuts_mm or [])
    x = max(1.0, min(L - 1.0, x))
    cuts = []
    for c in cuts_mm or []:
        try:
            cuts.append(float(c))
        except Exception:
            continue
    for i, c in enumerate(cuts):
        if abs(c - x) <= CUT_HIT_MM:
            return [cuts[j] for j in range(len(cuts)) if j != i]
    cuts.append(ceil10_mm(x))
    return sorted(set(cuts))


def format_mm_es(mm):
    try:
        v = int(round(float(mm)))
    except Exception:
        v = 0
    try:
        return u"{0:,}".format(v).replace(u",", u".")
    except Exception:
        return u"{0}".format(v)


def lap_zone_around_cut(cut, lap, mode):
    m = normalize_lap_mode_ui(mode)
    c = float(cut)
    L = float(lap)
    if m == LAP_MODE_ENDPOINT_PREV:
        return c, c + L
    if m == LAP_MODE_ENDPOINT_NEXT:
        return c - L, c
    return c - 0.5 * L, c + 0.5 * L
