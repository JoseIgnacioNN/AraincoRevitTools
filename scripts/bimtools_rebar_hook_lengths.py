# -*- coding: utf-8 -*-
"""
Tabla fija BIMTools: largo de **pata** / gancho (mm) por diámetro nominal (mm).

Valores según especificación de proyecto (tabla ø → largo de pata):

  8→160, 10→200, 12→240, 16→320, 18→360, 22→440, 25→500, 28→570, 32→650, 36→720

**Traslape / empalme** y **ganchos**: :func:`traslape_mm_from_nominal_diameter_mm` /
:func:`hook_length_mm_from_nominal_diameter_mm` según ``concrete_grade`` (**G25**, **G35**,
**G45**). Grado no reconocido o ``None``: tablas base BIMTools (legacy).

Para polilíneas / ``CreateFromCurves``, usar :func:`pata_eje_curve_loop_mm_desde_tabla_mm` a
partir del valor de tabla para alinear A/C de forma con el criterio del proyecto.

El diámetro de entrada se redondea al **entero mm** más cercano antes de interpolar:
así un nominal de 8 mm no pasa a ~164 mm si la API devuelve ~8,2 mm (interpolación 8–10).

Diámetros intermedios (p. ej. 14 mm): interpolación lineal entre nodos consecutivos; fuera del
rango se usa el extremo más cercano.
"""

# Pares (diámetro_nominal_mm, largo_pata_mm), ordenados por diámetro.
_BIMTOOLS_HOOK_LENGTH_BY_DIAMETER_MM = (
    (8, 160),
    (10, 200),
    (12, 240),
    (16, 320),
    (18, 360),
    (22, 440),
    (25, 500),
    (28, 570),
    (32, 650),
    (36, 720),
)

# Referencia de solo lectura para otros módulos (no modificar la tupla en runtime).
BIMTOOLS_REBAR_HOOK_LENGTH_MM_TABLE = _BIMTOOLS_HOOK_LENGTH_BY_DIAMETER_MM

# Pares (diámetro_nominal_mm, traslape_mm); tabla base solo si grado no G25/G35/G45.
_BIMTOOLS_TRASLAPE_BY_DIAMETER_MM = (
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
BIMTOOLS_TRASLAPE_LENGTH_MM_TABLE = _BIMTOOLS_TRASLAPE_BY_DIAMETER_MM

# Hormigón G25 — traslapes y empotramientos (mm); ganchos = misma escala cm→mm que la tabla de patas.
_G25_TRASLAPE_BY_DIAMETER_MM = (
    (8, 560),
    (10, 690),
    (12, 840),
    (16, 1110),
    (18, 1240),
    (22, 1890),
    (25, 2150),
    (28, 2410),
    (32, 2750),
    (36, 3090),
)

# G35 — ganchos (mm, tabla proyecto en cm ×10), ø 8…36 mm.
_G35_HOOK_LENGTH_BY_DIAMETER_MM = (
    (8, 150),
    (10, 170),
    (12, 210),
    (16, 270),
    (18, 310),
    (22, 380),
    (25, 430),
    (28, 480),
    (32, 540),
    (36, 610),
)

# G35 — traslapes / empotramientos (mm, tabla proyecto en cm ×10).
_G35_TRASLAPE_BY_DIAMETER_MM = (
    (8, 470),
    (10, 590),
    (12, 710),
    (16, 940),
    (18, 1060),
    (22, 1600),
    (25, 1810),
    (28, 2030),
    (32, 2320),
    (36, 2620),
)

# G45 — ganchos (mm, tabla proyecto en cm ×10), ø 8…36 mm.
_G45_HOOK_LENGTH_BY_DIAMETER_MM = (
    (8, 150),
    (10, 150),
    (12, 180),
    (16, 240),
    (18, 270),
    (22, 330),
    (25, 380),
    (28, 420),
    (32, 480),
    (36, 540),
)

# G45 — traslapes / empotramientos (mm, tabla proyecto en cm ×10).
_G45_TRASLAPE_BY_DIAMETER_MM = (
    (8, 420),
    (10, 520),
    (12, 630),
    (16, 820),
    (18, 930),
    (22, 1410),
    (25, 1600),
    (28, 1800),
    (32, 2050),
    (36, 2310),
)

# Mínimo de tramo recto (eje) al compensar frente a parámetros de forma en Revit.
_MIN_PATA_TRAMO_EJE_MM = 40.0


def _normalize_concrete_grade(concrete_grade):
    if concrete_grade is None:
        return None
    try:
        s = unicode(concrete_grade).strip().upper()
    except Exception:
        return None
    if s in (u"G25", u"G35", u"G45"):
        return s
    return None


def _interpolate_length_mm_from_table(d_int, tbl):
    """d_int: diámetro ya redondeado a mm entero. tbl: secuencia de (d_mm, L_mm)."""
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
    # G25 y None: ganchos tabla legacy (= G25 proyecto en cm×10).
    return _BIMTOOLS_HOOK_LENGTH_BY_DIAMETER_MM


def pata_eje_curve_loop_mm_desde_tabla_mm(tabla_pata_mm, diameter_nominal_mm):
    """
    Longitud de cada **tramo recto** de la polilínea (eje de barra) para que los parámetros de
    forma que Revit muestra (p. ej. **A**, **C** en forma «03») coincidan con ``tabla_pata_mm``.

    Si se modela el eje a la misma cota que la tabla, Revit suele reportar ~«tabla + ø/2»;
    aquí se resta **medio diámetro nominal** (redondeado a mm, igual que la tabla).
    """
    try:
        Ltab = float(tabla_pata_mm)
    except Exception:
        return None
    try:
        d = float(diameter_nominal_mm)
    except Exception:
        d = 0.0
    d = float(int(round(d)))
    if d > 1e-6:
        Leje = Ltab - 0.5 * d
    else:
        Leje = Ltab
    if Leje < float(_MIN_PATA_TRAMO_EJE_MM):
        Leje = float(_MIN_PATA_TRAMO_EJE_MM)
    return float(Leje)


def traslape_mm_from_nominal_diameter_mm(diameter_mm, concrete_grade=None):
    """
    Longitud de **traslape / empalme** (mm) según ø nominal de la **barra longitudinal**.

    Args:
        diameter_mm: ø nominal (mm).
        concrete_grade: ``'G25'`` / ``'G35'`` / ``'G45'`` → tablas de proyecto; ``None``
            → tabla base BIMTools.

    Returns:
        float o ``None`` si el diámetro no es válido.
    """
    try:
        d = float(diameter_mm)
    except Exception:
        return None
    if d <= 0.0 or d != d:
        return None
    d = float(int(round(d)))
    g = _normalize_concrete_grade(concrete_grade)
    tbl = _traslape_table_for_grade(g)
    return _interpolate_length_mm_from_table(d, tbl)


def hook_length_mm_from_nominal_diameter_mm(diameter_mm, concrete_grade=None):
    """
    Largo de gancho / pata (mm) por extremo según ø nominal (mm).

    Args:
        diameter_mm: diámetro nominal en mm (> 0). Si es ``None`` o inválido,
            se usa **12 mm** como respaldo.
        concrete_grade: ``'G35'`` / ``'G45'`` tablas de proyecto; ``'G25'`` y ``None`` → legacy.

    Returns:
        float: largo en mm (>= 0).
    """
    try:
        d = float(diameter_mm)
    except Exception:
        d = None
    if d is None or d <= 0.0:
        d = 12.0
    d = float(int(round(d)))
    g = _normalize_concrete_grade(concrete_grade)
    tbl = _hook_table_for_grade(g)
    return _interpolate_length_mm_from_table(d, tbl)
