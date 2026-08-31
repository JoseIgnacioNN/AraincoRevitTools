# -*- coding: utf-8 -*-
"""Constantes de dominio (alineadas al mockup canvas)."""

MAX_BAR_MM = 12000
BAR_COUNT_MIN = 2
BAR_COUNT_MAX = 8
CAPAS_MIN = 1
CAPAS_MAX = 3
CAPAS_DEFAULT = 1
LONG_DIAM_OPTS = (12, 16, 20, 25, 28, 32)
ESTRIBO_INSET_MM = 50
ESTRIBO_SPACING_MIN = 50
ESTRIBO_SPACING_MAX = 400
ESTRIBO_SPACING_DEFAULT_EXT = 100
ESTRIBO_SPACING_DEFAULT_CENT = 200
ESTRIBO_SPACING_OPTS = tuple(range(ESTRIBO_SPACING_MIN, 425, 25))
UMBRAL_EMPALME_MM = 12000

# Dosificación hormigón global (tablas traslape / empalme — misma clave BIMTools).
CONCRETE_GRADES = (u"G25", u"G35", u"G45")
CONCRETE_GRADE_DEFAULT = u"G25"

# Extremos libres de la corrida longitudinal (vista izq → der).
# auto = resolución por sonda; emp = forzar empotramiento; pata_l = forzar pata L.
BAR_END_MODE_AUTO = u"auto"
BAR_END_MODE_EMP = u"emp"
BAR_END_MODE_PATA_L = u"pata_l"
BAR_END_MODES = (BAR_END_MODE_AUTO, BAR_END_MODE_EMP, BAR_END_MODE_PATA_L)
BAR_END_MODE_DEFAULT = BAR_END_MODE_AUTO
BAR_END_MODE_LABELS = {
    BAR_END_MODE_AUTO: u"Auto",
    BAR_END_MODE_EMP: u"Empotramiento",
    BAR_END_MODE_PATA_L: u"Pata L",
}

# Alternancia empalme por capa (canvas + modelado):
# capas impares → corte en centro de viga; pares → + (k/2)·lap total
# (k=2 ⇒ un solape completo, sin superponer nudos de 1.ª/3.ª).
EMPALME_LAYER_ALT_LAP_K = 2.0


def normalize_bar_end_mode(mode):
    """``auto`` | ``emp`` | ``pata_l``."""
    if mode is None:
        return BAR_END_MODE_DEFAULT
    try:
        s = unicode(mode).strip().lower()
    except Exception:
        try:
            s = str(mode).strip().lower()
        except Exception:
            return BAR_END_MODE_DEFAULT
    aliases = {
        u"auto": BAR_END_MODE_AUTO,
        u"automatico": BAR_END_MODE_AUTO,
        u"automático": BAR_END_MODE_AUTO,
        u"emp": BAR_END_MODE_EMP,
        u"empotramiento": BAR_END_MODE_EMP,
        u"embed": BAR_END_MODE_EMP,
        u"pata_l": BAR_END_MODE_PATA_L,
        u"pata-l": BAR_END_MODE_PATA_L,
        u"patal": BAR_END_MODE_PATA_L,
        u"gancho": BAR_END_MODE_PATA_L,
        u"l": BAR_END_MODE_PATA_L,
    }
    if s in aliases:
        return aliases[s]
    if s in BAR_END_MODES:
        return s
    return BAR_END_MODE_DEFAULT


def normalize_concrete_grade(grade):
    """``G25`` / ``G35`` / ``G45``; inválido o vacío → G25."""
    if grade is None:
        return CONCRETE_GRADE_DEFAULT
    try:
        s = unicode(grade).strip().upper()
    except Exception:
        try:
            s = str(grade).strip().upper()
        except Exception:
            return CONCRETE_GRADE_DEFAULT
    if s in CONCRETE_GRADES:
        return s
    return CONCRETE_GRADE_DEFAULT
