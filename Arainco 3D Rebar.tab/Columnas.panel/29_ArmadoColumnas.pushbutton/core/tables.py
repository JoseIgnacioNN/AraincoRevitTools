# -*- coding: utf-8 -*-
"""
Tablas BIMTools de traslapes, ganchos y patas por diámetro nominal.

MÓDULO PURO: cero imports de Revit API ni de IronPython.
Portado literalmente desde scripts/bimtools_rebar_hook_lengths.py para
hacer el .pushbutton 100% self-contained.

Diámetros intermedios: interpolación lineal entre nodos; fuera del rango
se usa el extremo más cercano. El diámetro se redondea al entero mm más
cercano antes de interpolar.
"""

# ---------------------------------------------------------------------------
# Tablas de datos
# ---------------------------------------------------------------------------

_BIMTOOLS_HOOK_LENGTH_BY_DIAMETER_MM = (
    (8,  160), (10, 200), (12, 240), (16, 320),
    (18, 360), (22, 440), (25, 500), (28, 570), (32, 650), (36, 720),
)
BIMTOOLS_REBAR_HOOK_LENGTH_MM_TABLE = _BIMTOOLS_HOOK_LENGTH_BY_DIAMETER_MM

_BIMTOOLS_TRASLAPE_BY_DIAMETER_MM = (
    (8,  570), (10,  710), (12,  860), (16, 1140),
    (18, 1290), (22, 1960), (25, 2230), (28, 2500), (32, 2850), (36, 3210),
)
BIMTOOLS_TRASLAPE_LENGTH_MM_TABLE = _BIMTOOLS_TRASLAPE_BY_DIAMETER_MM

_G25_TRASLAPE_BY_DIAMETER_MM = (
    (8,  560), (10,  690), (12,  840), (16, 1110),
    (18, 1240), (22, 1890), (25, 2150), (28, 2410), (32, 2750), (36, 3090),
)

_G35_HOOK_LENGTH_BY_DIAMETER_MM = (
    (8,  150), (10, 170), (12, 210), (16, 270),
    (18, 310), (22, 380), (25, 430), (28, 480), (32, 540), (36, 610),
)
_G35_TRASLAPE_BY_DIAMETER_MM = (
    (8,  470), (10,  590), (12,  710), (16,  940),
    (18, 1060), (22, 1600), (25, 1810), (28, 2030), (32, 2320), (36, 2620),
)

_G45_HOOK_LENGTH_BY_DIAMETER_MM = (
    (8,  150), (10, 150), (12, 180), (16, 240),
    (18, 270), (22, 330), (25, 380), (28, 420), (32, 480), (36, 540),
)
_G45_TRASLAPE_BY_DIAMETER_MM = (
    (8,  420), (10,  520), (12,  630), (16,  820),
    (18,  930), (22, 1410), (25, 1600), (28, 1800), (32, 2050), (36, 2310),
)

_MIN_PATA_TRAMO_EJE_MM = 40.0

# ---------------------------------------------------------------------------
# Helpers internos
# ---------------------------------------------------------------------------

def _normalize_concrete_grade(concrete_grade):
    if concrete_grade is None:
        return None
    try:
        s = str(concrete_grade).strip().upper()
    except Exception:
        return None
    if s in (u"G25", u"G35", u"G45"):
        return s
    return None


def _interpolate_length_mm_from_table(d_int, tbl):
    d = float(d_int)
    d0, L0 = tbl[0]
    if d <= float(d0):
        return float(L0)
    d_last, L_last = tbl[-1]
    if d >= float(d_last):
        return float(L_last)
    for i in range(len(tbl) - 1):
        da, La = tbl[i]
        db, Lb = tbl[i + 1]
        da, db = float(da), float(db)
        if da <= d <= db:
            if abs(db - da) < 1e-9:
                return float(La)
            t = (d - da) / (db - da)
            return float(La) + t * (float(Lb) - float(La))
    return float(L_last)


def _traslape_table_for_grade(grade_norm):
    if grade_norm == u"G25":
        return _G25_TRASLAPE_BY_DIAMETER_MM
    if grade_norm == u"G35":
        return _G35_TRASLAPE_BY_DIAMETER_MM
    if grade_norm == u"G45":
        return _G45_TRASLAPE_BY_DIAMETER_MM
    return _BIMTOOLS_TRASLAPE_BY_DIAMETER_MM


def _hook_table_for_grade(grade_norm):
    if grade_norm == u"G35":
        return _G35_HOOK_LENGTH_BY_DIAMETER_MM
    if grade_norm == u"G45":
        return _G45_HOOK_LENGTH_BY_DIAMETER_MM
    return _BIMTOOLS_HOOK_LENGTH_BY_DIAMETER_MM

# ---------------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------------

def traslape_mm_from_nominal_diameter_mm(diameter_mm, concrete_grade=None):
    """Longitud de traslape/empalme (mm) según Ø nominal y grado de hormigón."""
    try:
        d = float(diameter_mm)
    except Exception:
        return None
    if d <= 0.0 or d != d:
        return None
    d = float(int(round(d)))
    g = _normalize_concrete_grade(concrete_grade)
    return _interpolate_length_mm_from_table(d, _traslape_table_for_grade(g))


def hook_length_mm_from_nominal_diameter_mm(diameter_mm, concrete_grade=None):
    """Largo de gancho/pata (mm) por extremo según Ø nominal."""
    try:
        d = float(diameter_mm)
    except Exception:
        d = None
    if d is None or d <= 0.0:
        d = 12.0
    d = float(int(round(d)))
    g = _normalize_concrete_grade(concrete_grade)
    return _interpolate_length_mm_from_table(d, _hook_table_for_grade(g))


def pata_eje_curve_loop_mm_desde_tabla_mm(tabla_pata_mm, diameter_nominal_mm):
    """
    Tramo recto del eje de la pata (mm) para que los parámetros de forma
    de Revit coincidan con la tabla. Resta medio Ø nominal al valor de tabla.
    """
    try:
        Ltab = float(tabla_pata_mm)
    except Exception:
        return None
    try:
        d = float(int(round(float(diameter_nominal_mm))))
    except Exception:
        d = 0.0
    Leje = Ltab - 0.5 * d if d > 1e-6 else Ltab
    return float(max(Leje, _MIN_PATA_TRAMO_EJE_MM))
