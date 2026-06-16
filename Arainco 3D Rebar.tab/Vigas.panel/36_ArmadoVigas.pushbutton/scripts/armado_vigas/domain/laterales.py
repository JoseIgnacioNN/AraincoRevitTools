# -*- coding: utf-8 -*-
"""Barras laterales del alma — cantidad, separación y diámetro de confinamiento."""

from __future__ import division

import math

from armado_vigas.domain.layers import beam_n_capas_inf, beam_n_capas_sup
from armado_vigas.domain.stirrups import compute_stirrup_zones, section_height_mm

LATERAL_CLEAR_FLEX_MM = 100.0
LATERAL_COUNT_STEP_MM = 200.0
LAYER_OFFSET_MM = 50.0
LATERALES_COUNT_MIN = 1
LATERALES_COUNT_MAX = 99
LATERALES_DIAM_DEFAULT = 16


def suggest_n_laterales(h_mm):
    """
    Cantidad sugerida: ``ceil(h_mm / 200) - 1``, mínimo 1.
    Ej.: 720 mm → ceil(3.6) − 1 = 3.
    """
    if h_mm is None or h_mm <= 0:
        return LATERALES_COUNT_MIN
    n = int(math.ceil(float(h_mm) / LATERAL_COUNT_STEP_MM)) - 1
    return max(LATERALES_COUNT_MIN, min(LATERALES_COUNT_MAX, n))


def suggest_n_laterales_from_beams(domain_beams):
    """Mayor altura de sección del lote → cantidad sugerida."""
    h_max = None
    for beam in domain_beams or []:
        h = section_height_mm(beam.get("type"))
        if h > 0:
            h_max = h if h_max is None else max(h_max, h)
    if h_max is None:
        return LATERALES_COUNT_MIN
    return suggest_n_laterales(h_max)


def lateral_clear_mm(beam):
    """
    Hueco vertical (mm) entre fibras flexión y zona de laterales, además del recubrimiento
    que ya aplica ``armadura_vigas_capas``. Incluye 100 mm fijos y desplazamiento de capas
    (+(nCapas−1)·50 mm hacia el centro).
    """
    ensure = beam or {}
    n_sup = max(1, int(beam_n_capas_sup(ensure)))
    n_inf = max(1, int(beam_n_capas_inf(ensure)))
    layer_extra = max(n_sup - 1, n_inf - 1) * LAYER_OFFSET_MM
    return float(LATERAL_CLEAR_FLEX_MM) + float(layer_extra)


def lateral_clear_mm_for_chain(domain_beams_by_id, chain_elems):
    """Máximo ``lateral_clear_mm`` entre vigas de una cadena colineal."""
    clear = LATERAL_CLEAR_FLEX_MM
    for el in chain_elems or []:
        try:
            eid = int(el.Id.IntegerValue)
        except Exception:
            continue
        beam = (domain_beams_by_id or {}).get(eid)
        if beam is not None:
            clear = max(clear, lateral_clear_mm(beam))
    return clear


def conf_diam_mm(beam):
    """
    Diámetro de estribo/confinamiento (mm) para inset en cara del alma.
    Zona central si hay tramo Ext+Cent; si no, el único lote.
    """
    if not beam:
        return 8
    plan = compute_stirrup_zones(beam)
    if plan.get("mode") == "single":
        z = (plan.get("zones") or [{}])[0]
        role = z.get("role") or u"cent"
        if role == u"ext":
            return int(beam.get("estExtDiam") or 10)
        return int(beam.get("estCentDiam") or beam.get("estExtDiam") or 8)
    return int(beam.get("estCentDiam") or beam.get("estExtDiam") or 8)
