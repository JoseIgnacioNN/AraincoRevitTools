# -*- coding: utf-8 -*-
"""Suple superior — capa n_capas_sup + 1, 25 % en cada extremo; fusión en vigas consecutivas.

Los tipos ``start`` / ``end`` / ``merged`` siguen el orden del canvas (izquierda→derecha).
La colocación Revit traduce ini/fin de LocationCurve con ``axisReversed``.
"""

from armado_vigas.domain.constants import BAR_COUNT_MIN
from armado_vigas.domain.layers import beam_n_capas_sup, clamp_bar_count
from armado_vigas.domain.tramos import beams_share_bar_run_section

SUPLE_END_PCT = 0.25
DEFAULT_DIAM_SUPLE_SUP_MM = 16
DEFAULT_N_SUPLE_SUP = 2


def ensure_beam_suple_superior(beam):
    """Inicializa campos de suple superior en el dict de viga."""
    if beam.get("supleSupEnabled") is None:
        beam["supleSupEnabled"] = False
    if beam.get("supleSupStartEnabled") is None:
        beam["supleSupStartEnabled"] = True
    if beam.get("supleSupEndEnabled") is None:
        beam["supleSupEndEnabled"] = True
    if beam.get("diamSupleSup") is None:
        beam["diamSupleSup"] = DEFAULT_DIAM_SUPLE_SUP_MM
    if beam.get("nSupleSup") is None:
        beam["nSupleSup"] = DEFAULT_N_SUPLE_SUP
    beam["nSupleSup"] = clamp_bar_count(beam["nSupleSup"])
    return beam


def beam_suple_sup_enabled(beam):
    ensure_beam_suple_superior(beam)
    return bool(beam.get("supleSupEnabled"))


def beam_suple_sup_side_enabled(beam, view_side):
    """True si el suple superior está activo en el extremo canvas ``start`` o ``end``."""
    ensure_beam_suple_superior(beam)
    if not beam_suple_sup_enabled(beam):
        return False
    if view_side == "start":
        return bool(beam.get("supleSupStartEnabled"))
    if view_side == "end":
        return bool(beam.get("supleSupEndEnabled"))
    return False


def beam_suple_sup_layer_index(beam):
    """Índice 1-based: inmediatamente debajo de la última capa superior modelada."""
    return beam_n_capas_sup(beam) + 1


def suple_sup_metrics_mm(length_mm):
    """Longitud del tramo suple (25 % de L_viga) en mm."""
    L = max(0, int(round(float(length_mm or 0))))
    span = int(round(L * SUPLE_END_PCT))
    return {
        "Lmm": L,
        "spanMm": span,
        "endPct": int(SUPLE_END_PCT * 100),
    }


def beams_consecutive_for_suple(prev, cur):
    """Mismo criterio que corrida Tn superior en ``tramos.py``."""
    return beams_share_bar_run_section(prev, cur, es_cara_inferior=False)


def consecutive_pair_merges_suple(prev, cur):
    return (
        beams_consecutive_for_suple(prev, cur)
        and beam_suple_sup_enabled(prev)
        and beam_suple_sup_enabled(cur)
        and beam_suple_sup_side_enabled(prev, "end")
        and beam_suple_sup_side_enabled(cur, "start")
    )


def beam_axis_reversed(beam):
    """True si LocationCurve 0 queda a la derecha del 1 en la vista activa."""
    return bool(beam and beam.get("axisReversed"))


def suple_sup_trim_from_curve_start(beam, view_side):
    """
    Mapea extremo en canvas (``start``=izquierda, ``end``=derecha) al extremo
    0/1 de LocationCurve para :func:`trim_line_end_portion`.

    Vigas con flecha hacia la derecha (``axisReversed`` False): regla estándar.
    Vigas con flecha hacia la izquierda: invierte ini/fin respecto al canvas.
    """
    at_curve_start = view_side == "start"
    if beam_axis_reversed(beam):
        at_curve_start = not at_curve_start
    return at_curve_start


def suple_sup_resolver_at_view_side(beam, view_side):
    """
    Emp/gancho en el extremo libre según lado de canvas (no según índice de curva).

    Returns:
        ``(resolver_inicio, resolver_fin)`` — solo uno True.
    """
    rev = beam_axis_reversed(beam)
    if view_side == "start":
        return (not rev, rev)
    if view_side == "end":
        return (rev, not rev)
    return (False, False)


def merged_suple_sup_trim_sides(beam_a, beam_b):
    """
    Junta consecutiva (A izquierda, B derecha en canvas): extremos de curva
    que forman el tramo fusionado.

    Returns:
        ``(from_start_a, from_start_b)`` para :func:`trim_line_end_portion`.
    """
    return (
        suple_sup_trim_from_curve_start(beam_a, "end"),
        suple_sup_trim_from_curve_start(beam_b, "start"),
    )


def trim_line_view_end_portion(line, beam, view_side, pct=SUPLE_END_PCT):
    """Recorta ``pct`` desde el extremo izquierdo (``start``) o derecho (``end``) del canvas."""
    return trim_line_end_portion(
        line,
        from_start=suple_sup_trim_from_curve_start(beam, view_side),
        pct=pct,
    )


def compute_suple_sup_segment_specs(sorted_beams):
    """
    Especificaciones de tramos suple superior en orden de vigas (``u``).

    ``start`` / ``end`` son extremos izquierdo y derecho del canvas, no de LocationCurve.

    Returns:
        Lista de ``{"type": "start"|"end"|"merged", "indices": [i, ...]}``.
    """
    beams = list(sorted_beams or [])
    n = len(beams)
    specs = []

    for i in range(n):
        beam = beams[i]
        if not beam_suple_sup_enabled(beam):
            continue
        prev_merge = (
            i > 0 and consecutive_pair_merges_suple(beams[i - 1], beam)
        )
        next_merge = (
            i < n - 1 and consecutive_pair_merges_suple(beam, beams[i + 1])
        )
        if not prev_merge and beam_suple_sup_side_enabled(beam, "start"):
            specs.append({"type": "start", "indices": [i]})
        if not next_merge and beam_suple_sup_side_enabled(beam, "end"):
            specs.append({"type": "end", "indices": [i]})

    for i in range(n - 1):
        if consecutive_pair_merges_suple(beams[i], beams[i + 1]):
            specs.append({"type": "merged", "indices": [i, i + 1]})

    return specs


def trim_line_fraction(line, frac_start, frac_end):
    """
    Recorta ``line`` entre fracciones normalizadas ``[0, 1]`` a lo largo del eje.

    Returns:
        Línea Revit recortada o ``None`` si inválida.
    """
    if line is None:
        return None
    try:
        from Autodesk.Revit.DB import Line

        fs = max(0.0, min(1.0, float(frac_start)))
        fe = max(0.0, min(1.0, float(frac_end)))
        if fe <= fs + 1e-9:
            return None
        L = float(line.Length)
        if L < 1e-9:
            return None
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        du = (p1 - p0).Normalize()
        pa = p0 + du.Multiply(L * fs)
        pb = p0 + du.Multiply(L * fe)
        if pa.DistanceTo(pb) < 1e-6:
            return None
        return Line.CreateBound(pa, pb)
    except Exception:
        return None


def trim_line_end_portion(line, from_start=False, pct=SUPLE_END_PCT):
    """Conserva ``pct`` de la longitud desde el extremo start o end."""
    if from_start:
        return trim_line_fraction(line, 0.0, pct)
    return trim_line_fraction(line, 1.0 - pct, 1.0)


def suple_sup_segments_layout_px(sorted_beams, layouts, content_w, pct_to_px_fn):
    """
    Segmentos suple superior en px de alzado (para canvas WPF / mockup).

    ``pct_to_px_fn(pct, content_w)`` → posición horizontal en px.
    """
    if not beam_suple_sup_enabled_any(sorted_beams):
        return []

    specs = compute_suple_sup_segment_specs(sorted_beams)
    segs = []
    for spec in specs:
        typ = spec.get("type")
        idxs = spec.get("indices") or []
        if typ == "merged" and len(idxs) >= 2:
            i, j = idxs[0], idxs[1]
            if i >= len(layouts) or j >= len(layouts):
                continue
            lay_a = layouts[i]
            lay_b = layouts[j]
            left_a = pct_to_px_fn(lay_a["leftPct"], content_w)
            width_a = pct_to_px_fn(lay_a["widthPct"], content_w)
            left_b = pct_to_px_fn(lay_b["leftPct"], content_w)
            width_b = pct_to_px_fn(lay_b["widthPct"], content_w)
            span_a = width_a * SUPLE_END_PCT
            span_b = width_b * SUPLE_END_PCT
            segs.append({
                "type": "merged",
                "indices": idxs,
                "x0": left_a + width_a - span_a,
                "x1": left_b + span_b,
                "junctionX": left_a + width_a,
                "merged": True,
            })
        elif typ in ("start", "end") and idxs:
            i = idxs[0]
            if i >= len(layouts):
                continue
            lay_i = layouts[i]
            left = pct_to_px_fn(lay_i["leftPct"], content_w)
            width = pct_to_px_fn(lay_i["widthPct"], content_w)
            span_w = width * SUPLE_END_PCT
            if typ == "start":
                x0, x1 = left, left + span_w
            else:
                x0, x1 = left + width - span_w, left + width
            segs.append({
                "type": typ,
                "indices": idxs,
                "x0": x0,
                "x1": x1,
                "merged": False,
            })
    return sorted(segs, key=lambda s: s.get("x0", 0))


def beam_suple_sup_enabled_any(beams):
    for beam in beams or []:
        if (
            beam_suple_sup_side_enabled(beam, "start")
            or beam_suple_sup_side_enabled(beam, "end")
        ):
            return True
    return False
