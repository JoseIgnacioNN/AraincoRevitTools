# -*- coding: utf-8 -*-
"""Longitudes de traslape / empotramiento / pata según dosificación (G25/G35/G45)."""

from armado_vigas.domain.constants import CONCRETE_GRADE_DEFAULT, normalize_concrete_grade


def session_concrete_grade(session=None):
    """Grado de hormigón de la sesión Armado vigas (o por defecto G25)."""
    if session is not None:
        return normalize_concrete_grade(getattr(session, u"concreteGrade", None))
    try:
        from armado_vigas.revit.session import SESSION

        return normalize_concrete_grade(getattr(SESSION, u"concreteGrade", None))
    except Exception:
        return CONCRETE_GRADE_DEFAULT


def lap_mm_for_diameter(diam_mm, concrete_grade=None):
    """
    Traslape / empalme total (mm) según Ø y dosificación.

    Usa tablas BIMTools G25/G35/G45. Si falla, ``None``.
    """
    try:
        d = float(diam_mm)
    except Exception:
        return None
    if d <= 1e-9:
        return None
    g = normalize_concrete_grade(
        concrete_grade if concrete_grade is not None else session_concrete_grade()
    )
    try:
        from bimtools_rebar_hook_lengths import traslape_mm_from_nominal_diameter_mm

        L = traslape_mm_from_nominal_diameter_mm(d, g)
        if L is not None and float(L) > 1e-6:
            return float(L)
    except Exception:
        pass
    return None


def lap_mm_for_diameter_pair(diam_a_mm, diam_b_mm, concrete_grade=None):
    """
    Traslape (mm) entre tramos de distintos Ø: tabla del **mayor** diámetro.

    Si solo uno es válido, usa ese. Si ninguno, ``None``.
    """
    ds = []
    for d in (diam_a_mm, diam_b_mm):
        try:
            v = float(d)
            if v > 1e-9:
                ds.append(v)
        except Exception:
            pass
    if not ds:
        return None
    return lap_mm_for_diameter(max(ds), concrete_grade=concrete_grade)


def hook_mm_for_diameter(diam_mm, concrete_grade=None):
    """Largo de pata / gancho (mm) según Ø y dosificación."""
    try:
        d = float(diam_mm)
    except Exception:
        d = 16.0
    if d <= 1e-9:
        d = 16.0
    g = normalize_concrete_grade(
        concrete_grade if concrete_grade is not None else session_concrete_grade()
    )
    try:
        from bimtools_rebar_hook_lengths import hook_length_mm_from_nominal_diameter_mm

        L = hook_length_mm_from_nominal_diameter_mm(d, g)
        if L is not None and float(L) > 1e-6:
            return float(L)
    except Exception:
        pass
    return max(6.0 * d, 150.0)


def empotramiento_mm_for_diameter(diam_mm, concrete_grade=None):
    """
    Empotramiento / desarrollo (mm).

    En tablas de proyecto coincide con la fila de traslape por dosificación.
    """
    L = lap_mm_for_diameter(diam_mm, concrete_grade=concrete_grade)
    if L is not None and float(L) > 1e-6:
        g = normalize_concrete_grade(
            concrete_grade if concrete_grade is not None else session_concrete_grade()
        )
        try:
            d = int(round(float(diam_mm)))
        except Exception:
            d = 16
        return float(L), u"{0} · Ø{1} → emp./trasl. {2:.0f} mm".format(g, d, float(L))
    return None, u""
