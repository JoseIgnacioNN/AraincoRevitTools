# -*- coding: utf-8 -*-
"""Zonas de estribos Ext/Cent por viga (regla 2·h + override UI)."""

from armado_vigas.domain.constants import (
    ESTRIBO_INSET_MM,
    ESTRIBO_SPACING_DEFAULT_CENT,
    ESTRIBO_SPACING_DEFAULT_EXT,
    ESTRIBO_SPACING_MIN,
)

# Condición de zonificación del estribado (por viga).
# - auto: solo legado / cálculo interno (regla 2·h) → se resuelve a ext_cent|unico
# - ext_cent: extremos + centro (opción UI)
# - unico: un solo tramo en todo el vano (opción UI)
STIRRUP_ZONE_AUTO = u"auto"
STIRRUP_ZONE_EXT_CENT = u"ext_cent"
STIRRUP_ZONE_UNICO = u"unico"

# Opciones expuestas en UI CONF (Auto no se muestra: se asigna y luego se puede forzar).
STIRRUP_ZONE_MODE_OPTIONS = (
    (STIRRUP_ZONE_EXT_CENT, u"Extremos + Centro"),
    (STIRRUP_ZONE_UNICO, u"Único"),
)

_STIRRUP_ZONE_MODE_SET = frozenset(
    (STIRRUP_ZONE_AUTO, STIRRUP_ZONE_EXT_CENT, STIRRUP_ZONE_UNICO)
)
_STIRRUP_ZONE_LABEL_BY_MODE = dict(STIRRUP_ZONE_MODE_OPTIONS)
_STIRRUP_ZONE_LABEL_BY_MODE[STIRRUP_ZONE_AUTO] = u"Auto (regla 2·h)"  # legado
_STIRRUP_ZONE_MODE_BY_LABEL = dict((lab, mode) for mode, lab in STIRRUP_ZONE_MODE_OPTIONS)
_STIRRUP_ZONE_MODE_BY_LABEL[u"Auto (regla 2·h)"] = STIRRUP_ZONE_AUTO
_STIRRUP_ZONE_MODE_BY_LABEL[u"Auto"] = STIRRUP_ZONE_AUTO


def parse_beam_section(type_str):
    import re
    m = re.match(
        r"(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)",
        unicode(type_str or u"30×60"),
    )
    if m:
        return float(m.group(1)), float(m.group(2))
    return 30.0, 60.0


def section_height_mm(type_str):
    _, h = parse_beam_section(type_str)
    return int(round(h * 10.0))


def beam_array_length_mm(beam):
    return max(0, int(round(float(beam.get("len", 0.0)) * 1000.0) - 2 * ESTRIBO_INSET_MM))


def normalize_stirrup_zone_mode(raw):
    """Devuelve auto | ext_cent | unico (auto solo legado hasta ensure)."""
    if raw is None:
        return STIRRUP_ZONE_AUTO
    s = unicode(raw).strip()
    if not s:
        return STIRRUP_ZONE_AUTO
    low = s.lower().replace(u" ", u"_").replace(u"-", u"_")
    if low in _STIRRUP_ZONE_MODE_SET:
        return low
    if low in (u"triple", u"ext", u"extremos", u"extremos_centro", u"ext_cent"):
        return STIRRUP_ZONE_EXT_CENT
    if low in (u"single", u"uni", u"unico", u"único", u"merge", u"cent"):
        return STIRRUP_ZONE_UNICO
    if s in _STIRRUP_ZONE_MODE_BY_LABEL:
        return _STIRRUP_ZONE_MODE_BY_LABEL[s]
    return STIRRUP_ZONE_AUTO


def stirrup_zone_mode_label(mode):
    m = normalize_stirrup_zone_mode(mode)
    if m == STIRRUP_ZONE_AUTO:
        # No mostrar “Auto” en UI: se etiqueta como el recomendado o Ext+Cent.
        return _STIRRUP_ZONE_LABEL_BY_MODE.get(
            STIRRUP_ZONE_EXT_CENT, u"Extremos + Centro"
        )
    return _STIRRUP_ZONE_LABEL_BY_MODE.get(
        m, _STIRRUP_ZONE_LABEL_BY_MODE[STIRRUP_ZONE_EXT_CENT]
    )


def stirrup_zone_mode_labels():
    """Etiquetas del combo CONF (sin Auto)."""
    return [lab for _mode, lab in STIRRUP_ZONE_MODE_OPTIONS]


def recommended_stirrup_zone_mode(beam, l_arr_mm=None, h_mm=None):
    """Modo que asigna la regla 2·h (ext_cent | unico), sin opción Auto en UI."""
    g = _geometry_metrics(beam, l_arr_mm=l_arr_mm, h_mm=h_mm)
    return effective_stirrup_zone_kind(_plan_auto_from_metrics(g))


def ensure_beam_stirrup_zone_mode(beam):
    """Garantiza estZonasMode = ext_cent | unico.

    Si falta o era «auto», escribe la evaluación geométrica (regla 2·h).
    El usuario puede cambiar después entre Extremos+Centro y Único.
    """
    if beam is None:
        return STIRRUP_ZONE_UNICO
    mode = normalize_stirrup_zone_mode(beam.get("estZonasMode"))
    if mode == STIRRUP_ZONE_AUTO:
        mode = recommended_stirrup_zone_mode(beam)
    beam["estZonasMode"] = mode
    return mode


def _zone_fracs(lw, l_arr, l_ext_each, l_cent, triple):
    s0 = float(ESTRIBO_INSET_MM)
    if lw <= 0:
        lw = 1.0
    if not triple:
        return [{
            "role": "uni",
            "lenMm": int(round(l_arr)),
            "fracStart": s0 / float(lw),
            "fracLen": l_arr / float(lw) if l_arr else 1.0,
        }]
    return [
        {
            "role": "ext",
            "lenMm": int(round(l_ext_each)),
            "fracStart": s0 / float(lw),
            "fracLen": l_ext_each / float(lw) if l_ext_each else 0.0,
        },
        {
            "role": "cent",
            "lenMm": int(round(l_cent)),
            "fracStart": (s0 + l_ext_each) / float(lw),
            "fracLen": l_cent / float(lw) if l_cent else 0.0,
        },
        {
            "role": "ext",
            "lenMm": int(round(l_ext_each)),
            "fracStart": (s0 + l_ext_each + l_cent) / float(lw),
            "fracLen": l_ext_each / float(lw) if l_ext_each else 0.0,
        },
    ]


def _geometry_metrics(beam, l_arr_mm=None, h_mm=None):
    """Métricas de zonificación; permite longitud de array y canto físicos (modelado)."""
    if h_mm is None:
        h_mm = section_height_mm(beam.get("type") if beam else None)
    else:
        h_mm = float(h_mm)
    if l_arr_mm is None:
        l_arr = float(beam_array_length_mm(beam)) if beam else 0.0
    else:
        l_arr = max(0.0, float(l_arr_mm))
    sp_e = ESTRIBO_SPACING_DEFAULT_EXT
    sp_c = ESTRIBO_SPACING_DEFAULT_CENT
    if beam is not None:
        sp_e = max(ESTRIBO_SPACING_MIN, int(beam.get("estExtSpacing") or ESTRIBO_SPACING_DEFAULT_EXT))
        sp_c = max(ESTRIBO_SPACING_MIN, int(beam.get("estCentSpacing") or ESTRIBO_SPACING_DEFAULT_CENT))
    min_len = max(sp_e, sp_c) * 0.2
    min_edge = 50.0
    if beam is not None:
        lw = int(round(float(beam.get("len", 0.0)) * 1000.0))
    else:
        lw = int(round(l_arr + 2.0 * ESTRIBO_INSET_MM))
    if l_arr_mm is not None:
        # Array físico: long. total de referencia ≈ array + recortes extremos.
        lw = max(lw, int(round(l_arr + 2.0 * ESTRIBO_INSET_MM)))
    two_h = 2.0 * float(h_mm)
    return {
        "h_mm": h_mm,
        "l_arr": l_arr,
        "lw": lw if lw > 0 else 1,
        "two_h": two_h,
        "min_len": min_len,
        "min_edge": min_edge,
        "sp_e": sp_e,
        "sp_c": sp_c,
    }


def _plan_auto_from_metrics(g):
    """Cálculo por defecto: Ext+Cent si L_arr ≥ 2·h y el centro es viable; si no, único."""
    l_arr = g["l_arr"]
    lw = g["lw"]
    two_h = g["two_h"]
    min_len = g["min_len"]
    min_edge = g["min_edge"]

    if l_arr < two_h - 1e-6:
        zones = _zone_fracs(lw, l_arr, 0.0, l_arr, triple=False)
        return {
            "mode": "single",
            "singleKind": "cent",
            "L_ext_each": 0,
            "L_cent": int(round(l_arr)),
            "zones": zones,
        }

    l_ext_tgt = two_h
    l_half = 0.5 * l_arr
    l_ext_each = min(l_ext_tgt, l_half)
    l_cent = max(0.0, l_arr - 2.0 * l_ext_each)
    if l_cent < min_len + min_edge:
        zones = _zone_fracs(lw, l_arr, l_ext_each, l_cent, triple=False)
        return {
            "mode": "single",
            "singleKind": "merge",
            "L_ext_each": int(round(l_ext_each)),
            "L_cent": int(round(l_cent)),
            "zones": zones,
        }

    zones = _zone_fracs(lw, l_arr, l_ext_each, l_cent, triple=True)
    return {
        "mode": "triple",
        "L_ext_each": int(round(l_ext_each)),
        "L_cent": int(round(l_cent)),
        "zones": zones,
    }


def _plan_force_ext_cent_from_metrics(g):
    l_arr = g["l_arr"]
    lw = g["lw"]
    two_h = g["two_h"]
    if l_arr <= 0:
        zones = _zone_fracs(lw, 0.0, 0.0, 0.0, triple=False)
        return {
            "mode": "single",
            "singleKind": "cent",
            "L_ext_each": 0,
            "L_cent": 0,
            "zones": zones,
        }
    l_half = 0.5 * l_arr
    l_ext_each = min(two_h, l_half)
    l_cent = max(0.0, l_arr - 2.0 * l_ext_each)
    if l_cent < 1.0 and l_arr < two_h - 1e-6:
        zones = _zone_fracs(lw, l_arr, 0.0, l_arr, triple=False)
        return {
            "mode": "single",
            "singleKind": "cent",
            "L_ext_each": int(round(l_ext_each)),
            "L_cent": int(round(l_arr)),
            "zones": zones,
            "forceCollapsed": True,
        }
    zones = _zone_fracs(lw, l_arr, l_ext_each, l_cent, triple=True)
    return {
        "mode": "triple",
        "L_ext_each": int(round(l_ext_each)),
        "L_cent": int(round(l_cent)),
        "zones": zones,
    }


def _plan_force_unico_from_metrics(g):
    l_arr = g["l_arr"]
    lw = g["lw"]
    zones = _zone_fracs(lw, l_arr, 0.0, l_arr, triple=False)
    return {
        "mode": "single",
        "singleKind": "cent",
        "L_ext_each": 0,
        "L_cent": int(round(l_arr)),
        "zones": zones,
    }


def _plan_auto(beam):
    return _plan_auto_from_metrics(_geometry_metrics(beam))


def _plan_force_ext_cent(beam):
    return _plan_force_ext_cent_from_metrics(_geometry_metrics(beam))


def _plan_force_unico(beam):
    return _plan_force_unico_from_metrics(_geometry_metrics(beam))


def effective_stirrup_zone_kind(plan):
    """'ext_cent' o 'unico' según plan ya resuelto."""
    if not plan:
        return STIRRUP_ZONE_UNICO
    if plan.get("mode") == "triple":
        return STIRRUP_ZONE_EXT_CENT
    return STIRRUP_ZONE_UNICO


def compute_stirrup_zones(beam, l_arr_mm=None, h_mm=None):
    """
    Returns dict: mode ('triple'|'single'), zones list, singleKind optional.

    Respeta beam['estZonasMode'] (ext_cent | unico). «auto» se resuelve al
    arranque en ensure_beam_stirrup_zone_mode con la regla 2·h.

    ``l_arr_mm`` / ``h_mm``: opcionales; se usan al modelar sobre la longitud
    real del array/host para alinear preview y Rebar.
    """
    ensure_beam_stirrup_zone_mode(beam)
    user_mode = normalize_stirrup_zone_mode(beam.get("estZonasMode") if beam else None)
    if user_mode == STIRRUP_ZONE_AUTO:
        # Seguridad: no debería quedarse auto tras ensure.
        user_mode = recommended_stirrup_zone_mode(beam, l_arr_mm=l_arr_mm, h_mm=h_mm)
        if beam is not None:
            beam["estZonasMode"] = user_mode
    g = _geometry_metrics(beam, l_arr_mm=l_arr_mm, h_mm=h_mm)
    auto_plan = _plan_auto_from_metrics(g)
    auto_kind = effective_stirrup_zone_kind(auto_plan)

    if user_mode == STIRRUP_ZONE_EXT_CENT:
        plan = _plan_force_ext_cent_from_metrics(g)
    elif user_mode == STIRRUP_ZONE_UNICO:
        plan = _plan_force_unico_from_metrics(g)
    else:
        plan = dict(auto_plan)

    plan["userMode"] = user_mode
    plan["autoMode"] = auto_kind
    plan["effectiveMode"] = effective_stirrup_zone_kind(plan)
    # Override puntual si el usuario eligió distinto a la evaluación 2·h.
    plan["isOverride"] = user_mode != auto_kind
    return plan


def stirrup_zone_placement_segments(beam, l_arr_mm, h_mm=None):
    """
    Segmentos (role, len_mm) para el reparto Rebar MaximumSpacing.

    Orden a lo largo del array: ext → cent → ext | uni.
    Longitudes en mm coherentes con ``l_arr_mm`` (array físico).
    """
    plan = compute_stirrup_zones(beam, l_arr_mm=l_arr_mm, h_mm=h_mm)
    out = []
    if plan.get("mode") == "triple":
        le = float(plan.get("L_ext_each") or 0.0)
        lc = float(plan.get("L_cent") or 0.0)
        # Ajuste residual por redondeo para cubrir L_arr exacta.
        total = 2.0 * le + lc
        target = float(l_arr_mm)
        if total > 0 and abs(total - target) > 0.5:
            scale = target / total
            le *= scale
            lc *= scale
        out.append((u"ext", le))
        out.append((u"cent", lc))
        out.append((u"ext", le))
    else:
        out.append((u"uni", float(l_arr_mm)))
    return out, plan
