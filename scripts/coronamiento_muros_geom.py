# -*- coding: utf-8 -*-
"""
Geometría pura — Coronamiento muros (sin Revit).

Cotas ↑10 mm, plan de capas, alternancia A-B-A de empalmes,
mapeo de escenarios de traslape (conjunto global).
"""

from __future__ import print_function

import math

MAX_BARRA_COMERCIAL_MM = 12000.0
COVER_SUPERIOR_MM = 25.0
PATA_RESTA_MM = 50.0
PATA_MIN_MM = 100.0
CUT_HIT_MM = 180.0
STAGGER_MARGIN_MM = 300.0  # margen extremo al clampar cortes desfasados
# Desfase A-B (capas impares): delta = L_traslape (solo traslape; no espejo L−c)
# Separación fija entre ejes (centerlines) de capas sucesivas en el tramo mayor.
LAYER_CENTERLINE_SPACING_MM = 50.0

DIAMS_MM = (8, 10, 12, 16, 18, 20, 22, 25)
N_BARS_OPTS = (2, 3, 4)
CAPAS_OPTS = (1, 2, 3)

# Dosificación hormigón (mismas claves que Wall Foundation / bimtools_rebar_hook_lengths)
DOSIFICACION_HORMIGON_OPCIONES = (u"G25", u"G35", u"G45")
DOSIFICACION_HORMIGON_DEFAULT = u"G25"

# Misma clave que 56_DividirRebarPuntoTraslape (dividir_rebar_punto_geom)
LAP_MODE_SYMMETRIC = u"symmetric"
LAP_MODE_ENDPOINT_PREV = u"endpoint_prev"
LAP_MODE_ENDPOINT_NEXT = u"endpoint_next"

LAP_MODE_LABELS = (
    (LAP_MODE_SYMMETRIC, u"Simétrico ± L/2"),
    (LAP_MODE_ENDPOINT_PREV, u"Endpoint · estira tramo anterior (+L)"),
    (LAP_MODE_ENDPOINT_NEXT, u"Endpoint · estira tramo siguiente (−L)"),
)

# Traslape base BIMTools (mm) — espejo bimtools_rebar_hook_lengths
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
        v = 16
    if v in DIAMS_MM:
        return v
    return 16


def clamp_n_capas(n):
    try:
        v = int(round(float(n)))
    except Exception:
        v = 1
    return max(1, min(3, v))


def normalize_concrete_grade(concrete_grade):
    """``G25`` / ``G35`` / ``G45``; otro valor → default G25."""
    try:
        s = u"{0}".format(concrete_grade or u"").strip().upper()
    except Exception:
        s = u""
    if s in DOSIFICACION_HORMIGON_OPCIONES:
        return s
    return DOSIFICACION_HORMIGON_DEFAULT


def traslape_mm_from_diam(diam_mm, concrete_grade=None):
    """
    Traslape / empotro (mm) por Ø y dosificación.

    Usa tablas de ``bimtools_rebar_hook_lengths`` (G25/G35/G45 o legacy si grade
    no reconocido internamente). Respaldo: tabla local base.
    """
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
    return 1140.0


def pata_mm_from_espesor(espesor_mm):
    e = ceil10_mm(espesor_mm)
    return max(PATA_MIN_MM, e - PATA_RESTA_MM)


def estimate_bar_lengths_mm(largo_muro_mm, espesor_mm):
    """
    Largo vano principal y desarrollado estimado (U libre: 2 patas).

    Returns:
        dict: main_mm, pata_mm, developed_mm, exceeds_12m, embed_mm, geom_mode
    """
    main_mm = ceil10_mm(largo_muro_mm)
    pata = pata_mm_from_espesor(espesor_mm)
    developed = main_mm + 2.0 * pata
    return {
        u"main_mm": main_mm,
        u"pata_mm": pata,
        u"developed_mm": developed,
        u"exceeds_12m": developed > MAX_BARRA_COMERCIAL_MM,
        u"embed_mm": 0.0,
        u"overhang_mm": 0.0,
        u"geom_mode": u"u_libre",
    }


def empotramiento_mm_from_diam(diam_mm, concrete_grade=None):
    """Empotramiento voladizo V3 ≈ tabla de traslape por Ø y dosificación."""
    return traslape_mm_from_diam(diam_mm, concrete_grade)


def estimate_empotrado_bar_lengths_mm(
    overhang_mm, espesor_mm, diam_mm=16, concrete_grade=None
):
    """
    Desarrollado Empotrado (V3 voladizo): empotro(Ø, grado) + voladizo libre + 1 pata L.

    Returns:
        dict: main_mm (horiz), pata_mm, developed_mm, exceeds_12m,
              embed_mm, overhang_mm, geom_mode
    """
    overhang = ceil10_mm(overhang_mm)
    embed = ceil10_mm(empotramiento_mm_from_diam(diam_mm, concrete_grade))
    pata = pata_mm_from_espesor(espesor_mm)
    main_mm = overhang + embed
    developed = main_mm + pata
    return {
        u"main_mm": main_mm,
        u"pata_mm": pata,
        u"developed_mm": developed,
        u"exceeds_12m": developed > MAX_BARRA_COMERCIAL_MM,
        u"embed_mm": embed,
        u"overhang_mm": overhang,
        u"geom_mode": u"empotrado",
    }


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
        u"forward": LAP_MODE_ENDPOINT_PREV,
        u"endpoint_next": LAP_MODE_ENDPOINT_NEXT,
        u"next": LAP_MODE_ENDPOINT_NEXT,
        u"siguiente": LAP_MODE_ENDPOINT_NEXT,
        u"backward": LAP_MODE_ENDPOINT_NEXT,
    }
    return aliases.get(key, LAP_MODE_SYMMETRIC)


def to_dividir_lap_mode(lap_mode_ui):
    """UI lap mode → symmetric|endpoint_prev|endpoint_next (API 56)."""
    return normalize_lap_mode_ui(lap_mode_ui)


def to_dividir_splice_mode(lap_mode_ui):
    """
    Compat: UI → symmetric|forward|backward (dividir_barra_traslape_punto).

    Preferir ``to_dividir_lap_mode`` para Dividir Rebar Punto (56).
    """
    m = normalize_lap_mode_ui(lap_mode_ui)
    if m == LAP_MODE_ENDPOINT_PREV:
        return u"forward"
    if m == LAP_MODE_ENDPOINT_NEXT:
        return u"backward"
    return u"symmetric"


def sync_layers(prev_layers, n_capas):
    """Conserva configs; rellena 2Ø16 en capas nuevas."""
    n = clamp_n_capas(n_capas)
    out = []
    prev = list(prev_layers or [])
    for i in range(n):
        if i < len(prev) and isinstance(prev[i], dict):
            out.append(
                {
                    u"n_bars": clamp_n_bars(prev[i].get(u"n_bars", 2)),
                    u"diam_mm": clamp_diam_mm(prev[i].get(u"diam_mm", 16)),
                }
            )
        else:
            out.append({u"n_bars": 2, u"diam_mm": 16})
    return out


def cover_axis_offset_mm_for_layer(layers, layer_index):
    """
    Offset desde cara superior hasta el eje de la capa ``layer_index``.

    Capa 0: cover + Ø0/2.
    Capas sucesivas: + ``LAYER_CENTERLINE_SPACING_MM`` (50 mm) entre ejes,
    independiente del diámetro.
    """
    if not layers:
        return COVER_SUPERIOR_MM + 8.0
    li = max(0, min(int(layer_index), len(layers) - 1))
    d0 = float(clamp_diam_mm(layers[0].get(u"diam_mm", 16)))
    z = COVER_SUPERIOR_MM + 0.5 * d0
    for i in range(1, li + 1):
        z += float(LAYER_CENTERLINE_SPACING_MM)
    return float(z)


def stagger_cuts_for_layer(cuts_ref_mm, layer_index, main_mm, lap_mm):
    """
    Alternancia A-B-A de estación de empalme.

    Capas pares (0, 2, …): mismos cortes que la referencia.
    Capas impares (1, 3, …): cada corte se desplaza delta = L_traslape
    hacia el centro del muro si cabe; no espejo L−c.
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
    delta = float(lap)
    if delta < 1.0:
        delta = STAGGER_MARGIN_MM
    lo = max(1.0, float(STAGGER_MARGIN_MM))
    hi = max(lo, L - lo)
    mid = 0.5 * L
    out = []
    for c in cuts:
        # Mismo |delta| en todos los cortes; sentido hacia midspan (o el opuesto si sale).
        sign = 1.0 if c < mid else -1.0
        cand = c + sign * delta
        if cand < lo or cand > hi:
            cand = c - sign * delta
        cand = max(lo, min(hi, cand))
        out.append(ceil10_mm(cand))
    return sorted(set(out))


def toggle_cut_at_mm(cuts_mm, click_mm, main_mm):
    """
    Clic cerca de un corte → lo quita; si no, añade (↑10 mm).

    Returns:
        nueva lista de cortes
    """
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
