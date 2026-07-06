# -*- coding: utf-8 -*-
"""Resolución de tipos Revit (RebarBarType) para armado vigas."""

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB.Structure import RebarBarType

try:
    from rebar_fundacion_cara_inferior import HOOK_GANCHO_90_STANDARD_NAME
except Exception:
    HOOK_GANCHO_90_STANDARD_NAME = u"Standard - 90 deg."


def _bar_diameter_mm_from_type(bar_type):
    if bar_type is None:
        return 0
    try:
        d_ft = float(bar_type.BarNominalDiameter)
        if d_ft <= 0:
            d_ft = float(getattr(bar_type, "BarModelDiameter", 0) or 0)
        return int(round(d_ft * 304.8)) if d_ft > 0 else 0
    except Exception:
        return 0


def list_bar_diameters_mm(document, fallback=None):
    """Diámetros únicos (mm) de ``RebarBarType`` en el documento, ordenados."""
    if fallback is None:
        from armado_vigas.domain.constants import LONG_DIAM_OPTS

        fallback = LONG_DIAM_OPTS
    if document is None:
        return tuple(fallback)
    seen = set()
    result = []
    try:
        types = FilteredElementCollector(document).OfClass(RebarBarType)
    except Exception:
        return tuple(fallback)
    for bt in types:
        dmm = _bar_diameter_mm_from_type(bt)
        if dmm <= 0 or dmm in seen:
            continue
        seen.add(dmm)
        result.append(dmm)
    if not result:
        return tuple(fallback)
    result.sort()
    return tuple(result)


def resolve_bar_type_mm(document, diam_mm, default_mm=16):
    if document is None:
        return None
    try:
        target = int(round(float(diam_mm or default_mm)))
    except Exception:
        target = int(default_mm)
    best = None
    best_diff = None
    try:
        types = FilteredElementCollector(document).OfClass(RebarBarType)
    except Exception:
        return None
    for bt in types:
        dmm = _bar_diameter_mm_from_type(bt)
        if dmm <= 0:
            continue
        diff = abs(dmm - target)
        if best_diff is None or diff < best_diff:
            best_diff = diff
            best = bt
    return best
