# -*- coding: utf-8 -*-
"""Suple inferior — capa n_capas_inf + 1, longitud central 80 % de cada viga."""

from armado_vigas.domain.constants import BAR_COUNT_MIN
from armado_vigas.domain.layers import beam_n_capas_inf, clamp_bar_count

SUPLE_TRIM_PCT_EACH_END = 0.10
SUPLE_SPAN_PCT = 0.80
DEFAULT_DIAM_SUPLE_INF_MM = 16
DEFAULT_N_SUPLE_INF = 2


def ensure_beam_suple_inferior(beam):
    """Inicializa campos de suple inferior en el dict de viga."""
    if beam.get("supleInfEnabled") is None:
        beam["supleInfEnabled"] = False
    if beam.get("diamSupleInf") is None:
        beam["diamSupleInf"] = DEFAULT_DIAM_SUPLE_INF_MM
    if beam.get("nSupleInf") is None:
        beam["nSupleInf"] = DEFAULT_N_SUPLE_INF
    beam["nSupleInf"] = clamp_bar_count(beam["nSupleInf"])
    return beam


def beam_suple_inf_enabled(beam):
    ensure_beam_suple_inferior(beam)
    return bool(beam.get("supleInfEnabled"))


def beam_suple_layer_index(beam):
    """Índice 1-based: inmediatamente después de la última capa inferior modelada."""
    return beam_n_capas_inf(beam) + 1


def suple_metrics_mm(length_mm):
    """Offsets y longitud central del suple para la longitud de una viga (mm)."""
    L = max(0, int(round(float(length_mm or 0))))
    offset = int(round(L * SUPLE_TRIM_PCT_EACH_END))
    span = int(round(L * SUPLE_SPAN_PCT))
    return {
        "Lmm": L,
        "offsetMm": offset,
        "spanMm": span,
        "trimPct": int(SUPLE_TRIM_PCT_EACH_END * 100),
        "spanPct": int(SUPLE_SPAN_PCT * 100),
    }


def trim_line_central_portion(line, pct_each_end=SUPLE_TRIM_PCT_EACH_END):
    """
    Recorta ``line`` dejando el tramo central ``(1 - 2·pct)`` de su longitud.

    Returns:
        Línea Revit recortada o ``None`` si inválida.
    """
    if line is None:
        return None
    try:
        from Autodesk.Revit.DB import Line

        L = float(line.Length)
        if L < 1e-9:
            return None
        trim = L * float(pct_each_end)
        if trim * 2.0 >= L - 1e-6:
            return None
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        du = (p1 - p0).Normalize()
        pa = p0 + du.Multiply(trim)
        pb = p1 - du.Multiply(trim)
        if pa.DistanceTo(pb) < 1e-6:
            return None
        return Line.CreateBound(pa, pb)
    except Exception:
        return None
