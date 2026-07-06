# -*- coding: utf-8 -*-
"""Adaptadores Revit → dominio."""

from __future__ import division

import clr

from armado_vigas.domain.constants import CAPAS_DEFAULT, ESTRIBO_SPACING_DEFAULT_CENT, ESTRIBO_SPACING_DEFAULT_EXT

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import BuiltInCategory, FamilyInstance, LocationCurve

_FRAMING_CAT = int(BuiltInCategory.OST_StructuralFraming)
_COL_CAT = int(BuiltInCategory.OST_StructuralColumns)
_WALL_CAT = int(BuiltInCategory.OST_Walls)


def elements_from_refs(document, refs_or_elements):
    out = []
    for item in refs_or_elements or []:
        el = None
        try:
            if hasattr(item, "ElementId"):
                el = document.GetElement(item.ElementId)
            elif hasattr(item, "Id"):
                el = item
            else:
                el = document.GetElement(item)
        except Exception:
            el = None
        if el is not None and el.IsValidObject:
            out.append(el)
    return out


def framing_from_elements(elements):
    out = []
    for el in elements or []:
        try:
            if el.Category and int(el.Category.Id.IntegerValue) == _FRAMING_CAT:
                if isinstance(el, FamilyInstance):
                    out.append(el)
        except Exception:
            pass
    return out


def _read_width_depth_ft(document, elem, curve):
    try:
        from armadura_vigas_capas import _read_width_depth_ft
        return _read_width_depth_ft(document, elem, curve)
    except Exception:
        pass
    return 0.3, 0.6


def _beam_type_label(document, elem, curve):
    try:
        w_ft, d_ft = _read_width_depth_ft(document, elem, curve)
        w_cm = int(round(float(w_ft) * 304.8 / 10.0))
        h_cm = int(round(float(d_ft) * 304.8 / 10.0))
        return u"{0}×{1}".format(w_cm, h_cm)
    except Exception:
        return u"30×60"


def _beam_length_m(elem):
    try:
        loc = elem.Location
        if isinstance(loc, LocationCurve):
            return float(loc.Curve.Length) * 304.8 / 1000.0
    except Exception:
        pass
    return 0.0


def _element_id_int(el):
    try:
        return int(el.Id.IntegerValue)
    except Exception:
        return None


def _element_label(el, prefix):
    try:
        p = el.LookupParameter(u"Mark")
        if p and p.HasValue and p.AsString():
            return p.AsString().strip()
    except Exception:
        pass
    eid = _element_id_int(el)
    return u"{0}-{1}".format(prefix, eid if eid is not None else u"?")


def _apoyo_span_mm_in_view(el, view):
    """Ancho del apoyo en el plano de la vista activa (mm)."""
    if view is None:
        return None
    try:
        bbox = el.get_BoundingBox(view)
        if bbox is None:
            return None
        dx = abs(float(bbox.Max.X) - float(bbox.Min.X))
        dy = abs(float(bbox.Max.Y) - float(bbox.Min.Y))
        span_ft = min(dx, dy) if dx > 1e-9 and dy > 1e-9 else max(dx, dy)
        if span_ft < 1e-9:
            return None
        return int(round(span_ft * 304.8))
    except Exception:
        return None


def apoyos_from_elements(elements, document=None, view=None):
    """Lista ordenada de apoyos (columnas/muros) con id legible."""
    apoyos = []
    for el in elements or []:
        try:
            cid = int(el.Category.Id.IntegerValue)
        except Exception:
            continue
        span_mm = _apoyo_span_mm_in_view(el, view)
        if cid == _COL_CAT:
            entry = {"id": _element_label(el, u"C"), "kind": "column", "element": el}
            if span_mm:
                entry["widthMm"] = span_mm
            apoyos.append(entry)
        elif cid == _WALL_CAT:
            entry = {"id": _element_label(el, u"M"), "kind": "wall", "element": el}
            if span_mm:
                entry["thicknessMm"] = span_mm
            apoyos.append(entry)
    return apoyos


def domain_beams_from_framing(document, framing_elements, apoyos=None):
    apoyos = apoyos or []
    beams = []
    for i, el in enumerate(framing_elements or []):
        curve = None
        try:
            if isinstance(el.Location, LocationCurve):
                curve = el.Location.Curve
        except Exception:
            curve = None
        name = u"V-{0}".format(_element_id_int(el) or u"?")
        try:
            p = el.LookupParameter(u"Mark")
            if p and p.HasValue and p.AsString():
                name = p.AsString().strip() or name
        except Exception:
            pass
        beams.append({
            "id": name,
            "elementIdInt": _element_id_int(el),
            "element": el,
            "type": _beam_type_label(document, el, curve),
            "len": _beam_length_m(el),
            "nCapasSup": CAPAS_DEFAULT,
            "nCapasInf": CAPAS_DEFAULT,
            "nSup": 2,
            "nInf": 2,
            "diamSup": 16,
            "diamInf": 16,
            "estExtDiam": 10,
            "estExtSpacing": ESTRIBO_SPACING_DEFAULT_EXT,
            "estCentDiam": 8,
            "estCentSpacing": ESTRIBO_SPACING_DEFAULT_CENT,
            "estConfin": u"Perimetral",
            "supleInfEnabled": False,
            "diamSupleInf": 16,
            "nSupleInf": 2,
            "supleSupEnabled": False,
            "supleSupStartEnabled": True,
            "supleSupEndEnabled": True,
            "diamSupleSup": 16,
            "nSupleSup": 2,
            "colStart": u"",
            "colEnd": u"",
        })
    return beams
