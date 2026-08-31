# -*- coding: utf-8 -*-
"""Barras laterales del alma — cantidad (ceil H/200 − 1) y clear de cara H/segs."""

from __future__ import division

import math

from armado_vigas.domain.stirrups import compute_stirrup_zones, section_height_mm

# Paso de segmentación del canto (mm).
LATERAL_COUNT_STEP_MM = 200.0
# Recubrimiento lateral alineado con armadura_vigas_capas (_COVER_MM_FIXED).
LATERAL_COVER_MM = 25.0
LATERALES_COUNT_MIN = 0
LATERALES_COUNT_MAX = 99
LATERALES_DIAM_DEFAULT = 10

# Residuo mínimo del span útil (mm) para que SetLayoutAsFixedNumber no colapse.
_LATERAL_MIN_STRIP_MM = 4.0

# --- Legacy (comentarios / fallback si no hay H) ---
LATERAL_CLEAR_FROM_LAST_LAYER_MM = 50.0
LAYER_OFFSET_MM = 50.0
LATERAL_CLEAR_FLEX_MM = LATERAL_CLEAR_FROM_LAST_LAYER_MM


def lateral_segments(h_mm):
    """
    Nº de segmentos: ``ceil(H / 200)`` (entero más próximo hacia arriba).

    Ej.: 700 → ceil(3.5) = 4.
    """
    if h_mm is None or h_mm <= 0:
        return 0
    return int(math.ceil(float(h_mm) / LATERAL_COUNT_STEP_MM))


def lateral_face_clear_mm(h_mm):
    """
    Distancia (mm) cara de hormigón → lateral extremo del set:
    ``H / max(1, ceil(H/200))``.

    Ej.: 700 → 700/4 = 175 mm (simétrico sup e inf).
    """
    if h_mm is None or h_mm <= 0:
        return 0.0
    segs = lateral_segments(h_mm)
    if segs < 1:
        segs = 1
    return float(h_mm) / float(segs)


def suggest_n_laterales(h_mm):
    """
    Cantidad: ``ceil(H / 200) − 1`` (mín. 0).

    Ej.: 700 → 4 − 1 = 3.  H ≤ 200 → segs=1 → n=0.
    """
    if h_mm is None or h_mm <= 0:
        return LATERALES_COUNT_MIN
    n = lateral_segments(h_mm) - 1
    return max(LATERALES_COUNT_MIN, min(LATERALES_COUNT_MAX, int(n)))


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


def beam_section_height_mm(beam):
    """Alto de sección (mm) del tipo de viga del dict de dominio."""
    if not beam:
        return 0.0
    try:
        return float(section_height_mm(beam.get("type")))
    except Exception:
        return 0.0


def lateral_ys_from_face_mm(h_mm, n_lat):
    """
    Coordenadas Y (mm desde cara superior) de cada lateral del set.

    Clears simétricos ``face_clear``; reparto FixedNumber entre extremos.
    """
    try:
        n = int(n_lat or 0)
    except Exception:
        n = 0
    if n < 1 or h_mm is None or h_mm <= 0:
        return []
    h = float(h_mm)
    clear = lateral_face_clear_mm(h)
    span = h - 2.0 * clear
    if span <= 0.5:
        return [h * 0.5]
    if n <= 1:
        return [clear + span * 0.5]
    step = span / float(n - 1)
    return [clear + i * step for i in range(n)]


def lateral_clear_mm(beam, bar_diam_mm=None):
    """
    Extra vertical (mm) para ``armadura_vigas_capas`` (tras recubrimiento+½ø).

    El modelo aplica ``posición_desde_cara = cover + ½ø + clear``.
    Queremos ``posición = face_clear = H/segs``, entonces:

        clear = face_clear − cover − ½ø  (clamped, con span mín. residual)

    """
    h = beam_section_height_mm(beam)
    if h <= 0:
        return float(LATERAL_CLEAR_FROM_LAST_LAYER_MM)
    return lateral_api_clear_mm(h, bar_diam_mm=bar_diam_mm)


def lateral_api_clear_mm(h_mm, bar_diam_mm=None, cover_mm=None):
    """Clear extra (mm) que, sumado a cover+½ø, sitúa el extremo del set a face_clear."""
    if h_mm is None or h_mm <= 0:
        return float(LATERAL_CLEAR_FROM_LAST_LAYER_MM)
    h = float(h_mm)
    face = lateral_face_clear_mm(h)
    cov = float(cover_mm if cover_mm is not None else LATERAL_COVER_MM)
    try:
        d = float(bar_diam_mm if bar_diam_mm is not None else LATERALES_DIAM_DEFAULT)
    except Exception:
        d = float(LATERALES_DIAM_DEFAULT)
    half_d = 0.5 * max(d, 0.0)
    # Posición deseada desde cara = face → extra = face − cover − ½ø
    extra = face - cov - half_d
    # Dejar un hueco mínimo entre n_bot y n_top (layout 1 barra o set).
    max_extra = 0.5 * h - cov - half_d - 0.5 * float(_LATERAL_MIN_STRIP_MM)
    if max_extra < 0.0:
        max_extra = 0.0
    if extra < 0.0:
        extra = 0.0
    if extra > max_extra:
        extra = max_extra
    return float(extra)


def lateral_clear_mm_for_chain(domain_beams_by_id, chain_elems, bar_diam_mm=None):
    """Máximo clear API entre vigas de una cadena colineal (misma franja cantil)."""
    clear = 0.0
    any_beam = False
    for el in chain_elems or []:
        try:
            eid = int(el.Id.IntegerValue)
        except Exception:
            continue
        beam = (domain_beams_by_id or {}).get(eid)
        if beam is not None:
            any_beam = True
            clear = max(clear, float(lateral_clear_mm(beam, bar_diam_mm=bar_diam_mm)))
    if not any_beam:
        # Fallback sin domain: paso 200 clásico poco útil; cover-linked legacy.
        return float(LATERAL_CLEAR_FROM_LAST_LAYER_MM)
    return clear


def session_n_laterales(session, default=0):
    """Lee ``session.nLaterales`` sin que ``0`` se convierta por ``or`` en otro valor."""
    if session is None:
        return int(default)
    try:
        v = getattr(session, "nLaterales", None)
        if v is None:
            return int(default)
        return max(LATERALES_COUNT_MIN, min(LATERALES_COUNT_MAX, int(v)))
    except Exception:
        return int(default)


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
