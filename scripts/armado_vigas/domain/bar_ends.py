# -*- coding: utf-8 -*-
"""Modo de extremos libres (empotramiento / pata L / auto) por cara SUP/INF."""

from armado_vigas.domain.constants import (
    BAR_END_MODE_AUTO,
    BAR_END_MODE_DEFAULT,
    BAR_END_MODE_EMP,
    BAR_END_MODE_LABELS,
    BAR_END_MODE_PATA_L,
    BAR_END_MODES,
    normalize_bar_end_mode,
)

# side: "start" = izquierda vista · "end" = derecha vista
_SESSION_KEYS = {
    (u"sup", u"start"): u"barEndStartSup",
    (u"sup", u"end"): u"barEndEndSup",
    (u"inf", u"start"): u"barEndStartInf",
    (u"inf", u"end"): u"barEndEndInf",
}


def session_end_mode_key(face, side):
    face = u"sup" if face != u"inf" else u"inf"
    side = u"end" if side == u"end" else u"start"
    return _SESSION_KEYS[(face, side)]


def get_bar_end_mode(session, face, side):
    """Lee modo extremo (canvas izq/der) de la sesión."""
    if session is None:
        return BAR_END_MODE_DEFAULT
    key = session_end_mode_key(face, side)
    return normalize_bar_end_mode(getattr(session, key, None))


def set_bar_end_mode(session, face, side, mode):
    if session is None:
        return BAR_END_MODE_DEFAULT
    key = session_end_mode_key(face, side)
    m = normalize_bar_end_mode(mode)
    setattr(session, key, m)
    return m


def ensure_session_bar_end_modes(session):
    """Inicializa claves de extremos si faltan (conserva valores previos)."""
    if session is None:
        return
    for key in _SESSION_KEYS.values():
        cur = getattr(session, key, None)
        setattr(session, key, normalize_bar_end_mode(cur))


def bar_end_mode_options_for_combo():
    """``(value, label)`` para ComboBox."""
    return [(m, BAR_END_MODE_LABELS.get(m, m)) for m in BAR_END_MODES]


def bar_end_mode_label(mode):
    m = normalize_bar_end_mode(mode)
    return BAR_END_MODE_LABELS.get(m, m)


def mode_from_label(label):
    """Etiqueta UI → valor de modo."""
    if label is None:
        return BAR_END_MODE_DEFAULT
    try:
        s = unicode(label).strip()
    except Exception:
        try:
            s = str(label).strip()
        except Exception:
            return BAR_END_MODE_DEFAULT
    for mode, lab in BAR_END_MODE_LABELS.items():
        if s == lab or s.lower() == mode:
            return mode
    return normalize_bar_end_mode(s)


# Tabla desarrollo/empotramiento (mm) por Ø — misma de ``enfierrado_shaft_hashtag``.
# Usada si no se puede importar la API BIMTools (p. ej. preview sin bootstrap Revit).
_EMP_FALLBACK_MM_BY_DIAM = (
    (8, 570),
    (10, 710),
    (12, 860),
    (16, 1140),
    (18, 1290),
    (22, 1960),
    (25, 2230),
    (28, 2500),
    (32, 2850),
    (36, 3210),
)


def _empotramiento_mm_fallback(diam_mm):
    """Interpolación por filas de respaldo (mm de empotramiento según Ø)."""
    table = _EMP_FALLBACK_MM_BY_DIAM
    try:
        d = float(diam_mm)
    except Exception:
        d = 16.0
    if d <= 1e-9:
        d = 16.0
    d_min, e_min = table[0]
    d_max, e_max = table[-1]
    if d <= d_min:
        return float(e_min), u"Ø≤{0} → {1} mm (tabla emp.)".format(int(d_min), int(e_min))
    if d >= d_max:
        return float(e_max), u"Ø≥{0} → {1} mm (tabla emp.)".format(int(d_max), int(e_max))
    for i in range(len(table) - 1):
        d_lo, e_lo = table[i]
        d_hi, e_hi = table[i + 1]
        if d_lo <= d <= d_hi:
            if abs(d - d_lo) < 1e-9:
                return float(e_lo), u"Ø{0} → {1} mm".format(int(d_lo), int(e_lo))
            if abs(d - d_hi) < 1e-9:
                return float(e_hi), u"Ø{0} → {1} mm".format(int(d_hi), int(e_hi))
            span = float(d_hi) - float(d_lo)
            if span <= 1e-12:
                return float(e_lo), u"Ø{0} → {1} mm".format(int(d_lo), int(e_lo))
            t = (d - float(d_lo)) / span
            ext = float(e_lo) + t * (float(e_hi) - float(e_lo))
            return float(ext), u"Ø{0:.0f} → {1:.0f} mm".format(float(d), float(ext))
    return float(e_min), u"Ø? → {0} mm".format(int(e_min))


def empotramiento_mm_for_diam(diam_mm, concrete_grade=None):
    """
    Longitud de empotramiento (mm) según Ø nominal y dosificación.

    Prefiere tablas G25/G35/G45 (``bimtools_rebar_hook_lengths`` vía
    ``domain.concrete_lengths``); si falla, API geometría o tabla local.
    """
    try:
        d = float(diam_mm)
    except Exception:
        d = 16.0
    if d <= 1e-9:
        d = 16.0

    try:
        from armado_vigas.domain.concrete_lengths import empotramiento_mm_for_diameter

        emp, desc = empotramiento_mm_for_diameter(d, concrete_grade=concrete_grade)
        if emp is not None and float(emp) > 1e-6:
            return float(emp), desc or u""
    except Exception:
        pass

    try:
        from geometria_empotramiento_extremos import _empotramiento_mm_desde_diametro

        emp, desc = _empotramiento_mm_desde_diametro(d, concrete_grade=concrete_grade)
        if emp is not None and float(emp) > 1e-6:
            return float(emp), desc or u""
    except TypeError:
        try:
            from geometria_empotramiento_extremos import _empotramiento_mm_desde_diametro

            emp, desc = _empotramiento_mm_desde_diametro(d)
            if emp is not None and float(emp) > 1e-6:
                return float(emp), desc or u""
        except Exception:
            pass
    except Exception:
        pass

    return _empotramiento_mm_fallback(d)


def curve_end_modes_from_view(mode_view_start, mode_view_end, beam_first=None, beam_last=None):
    """
    Traduce modos de canvas (izq/der) a extremos 0/1 de la curva fusionada.

    Usa ``axisReversed`` de la viga en cada extremo (misma regla que suple SUP).
    """
    m_left = normalize_bar_end_mode(mode_view_start)
    m_right = normalize_bar_end_mode(mode_view_end)
    rev0 = bool(beam_first and beam_first.get(u"axisReversed"))
    rev1 = bool(beam_last and beam_last.get(u"axisReversed"))
    # curve0 ≈ extremo físico de LocationCurve 0 del arranque de la cadena
    mode_c0 = m_right if rev0 else m_left
    mode_c1 = m_left if rev1 else m_right
    return mode_c0, mode_c1


def session_curve_end_modes(session, face, beam_first=None, beam_last=None):
    """``(end_mode_curve0, end_mode_curve1)`` desde sesión + vigas de extremos."""
    return curve_end_modes_from_view(
        get_bar_end_mode(session, face, u"start"),
        get_bar_end_mode(session, face, u"end"),
        beam_first=beam_first,
        beam_last=beam_last if beam_last is not None else beam_first,
    )


__all__ = [
    u"BAR_END_MODE_AUTO",
    u"BAR_END_MODE_EMP",
    u"BAR_END_MODE_PATA_L",
    u"BAR_END_MODES",
    u"BAR_END_MODE_DEFAULT",
    u"get_bar_end_mode",
    u"set_bar_end_mode",
    u"ensure_session_bar_end_modes",
    u"bar_end_mode_options_for_combo",
    u"bar_end_mode_label",
    u"mode_from_label",
    u"empotramiento_mm_for_diam",
    u"curve_end_modes_from_view",
    u"session_curve_end_modes",
    u"session_end_mode_key",
    u"normalize_bar_end_mode",
]
