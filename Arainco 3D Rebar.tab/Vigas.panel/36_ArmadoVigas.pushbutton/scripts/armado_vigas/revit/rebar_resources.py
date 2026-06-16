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
        if bt is None:
            continue
        try:
            d_ft = float(bt.BarNominalDiameter)
            if d_ft <= 0:
                d_ft = float(getattr(bt, "BarModelDiameter", 0) or 0)
            dmm = int(round(d_ft * 304.8)) if d_ft > 0 else 0
        except Exception:
            continue
        if dmm <= 0:
            continue
        diff = abs(dmm - target)
        if best_diff is None or diff < best_diff:
            best_diff = diff
            best = bt
    return best
