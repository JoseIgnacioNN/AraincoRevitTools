# -*- coding: utf-8 -*-
"""Adaptadores Revit → dominio de elevación / sección."""

from __future__ import division

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter

from armado_columnas_v2.revit.view_geometry import (
    elevation_rect_from_element,
    ft_to_cm,
    ft_to_m,
)

_COL_CAT = int(BuiltInCategory.OST_StructuralColumns)
_FND_CAT = int(BuiltInCategory.OST_StructuralFoundation)
_FRM_CAT = int(BuiltInCategory.OST_StructuralFraming)
_FLR_CAT = int(BuiltInCategory.OST_Floors)

# kind → prefijo de id/marca
_KIND_PREFIX = {
    u"column": u"C",
    u"foundation": u"F",
    u"beam": u"V",
    u"floor": u"L",
}

_KIND_ORDER = (u"floor", u"foundation", u"beam", u"column")

_WIDTH_NAMES = (
    u"b", u"B", u"Width", u"Ancho", u"width",
    u"b1", u"B1",
)
_DEPTH_NAMES = (
    u"h", u"H", u"Depth", u"Profundidad", u"depth",
    u"h1", u"H1",
)


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


def _kind_from_element(el):
    try:
        cid = int(el.Category.Id.IntegerValue)
    except Exception:
        return None
    if cid == _COL_CAT:
        return u"column"
    if cid == _FND_CAT:
        return u"foundation"
    if cid == _FLR_CAT:
        return u"floor"
    if cid == _FRM_CAT:
        return u"beam"
    return None


def classify_elements(elements):
    """Agrupa por kind: column / foundation / beam / floor / other."""
    groups = {
        u"column": [],
        u"foundation": [],
        u"beam": [],
        u"floor": [],
        u"other": [],
    }
    for el in elements or []:
        kind = _kind_from_element(el)
        if kind is None:
            groups[u"other"].append(el)
        else:
            groups[kind].append(el)
    return groups


def iter_classified(groups):
    """Itera elementos permitidos (sin other) en orden de dibujo de fondo."""
    for kind in _KIND_ORDER:
        for el in groups.get(kind) or []:
            yield el


def _element_id_int(el):
    try:
        return int(el.Id.IntegerValue)
    except Exception:
        return None


def _mark_or_fallback(el, prefix):
    try:
        p = el.LookupParameter(u"Mark")
        if p and p.HasValue and p.AsString():
            s = p.AsString().strip()
            if s:
                return s
    except Exception:
        pass
    eid = _element_id_int(el)
    return u"{0}-{1}".format(prefix, eid if eid is not None else u"?")


def _param_double_ft(target, names):
    if target is None:
        return None
    for n in names:
        try:
            p = target.LookupParameter(n)
            if p is not None and p.HasValue:
                return float(p.AsDouble())
        except Exception:
            pass
    return None


def _lookup_section_width_depth_ft(elem, document=None):
    """Ancho × profundidad de sección en pies (instancia → tipo)."""
    if elem is None:
        return None, None
    doc = document or getattr(elem, "Document", None)
    et = None
    try:
        tid = elem.GetTypeId()
        if doc is not None and tid is not None:
            et = doc.GetElement(tid)
    except Exception:
        et = None
    w = d = None
    for target in (elem, et):
        if w is None:
            w = _param_double_ft(target, _WIDTH_NAMES)
        if d is None:
            d = _param_double_ft(target, _DEPTH_NAMES)
        if w is not None and d is not None:
            break
    return w, d


def _floor_thickness_ft(elem, document=None):
    """Espesor de losa en pies."""
    if elem is None:
        return None
    try:
        p = elem.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
        if p is not None and p.HasValue:
            return float(p.AsDouble())
    except Exception:
        pass
    doc = document or getattr(elem, "Document", None)
    try:
        tid = elem.GetTypeId()
        if doc is not None and tid is not None:
            et = doc.GetElement(tid)
            if et is not None:
                p = et.get_Parameter(BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM)
                if p is not None and p.HasValue:
                    return float(p.AsDouble())
                p = et.LookupParameter(u"Thickness")
                if p is not None and p.HasValue:
                    return float(p.AsDouble())
                p = et.LookupParameter(u"Espesor")
                if p is not None and p.HasValue:
                    return float(p.AsDouble())
    except Exception:
        pass
    return None


def _level_name(document, level_id):
    if document is None or level_id is None:
        return None
    try:
        from Autodesk.Revit.DB import ElementId

        if level_id == ElementId.InvalidElementId:
            return None
        lv = document.GetElement(level_id)
        if lv is not None:
            return lv.Name
    except Exception:
        pass
    return None


def _column_levels(document, elem):
    base = top = None
    try:
        p = elem.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM)
        if p and p.HasValue:
            base = _level_name(document, p.AsElementId())
    except Exception:
        pass
    try:
        p = elem.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
        if p and p.HasValue:
            top = _level_name(document, p.AsElementId())
    except Exception:
        pass
    return base, top


def _type_name(document, elem):
    try:
        tid = elem.GetTypeId()
        if document is not None and tid is not None:
            et = document.GetElement(tid)
            if et is not None:
                return et.Name
    except Exception:
        pass
    try:
        return elem.Name
    except Exception:
        return u""


def _section_label_cm(w_cm, d_cm):
    try:
        return u"{0:.0f}×{1:.0f}".format(float(w_cm), float(d_cm))
    except Exception:
        return u""


def _plan_xy_ft(elem):
    """Huella en planta (bbox modelo, pies): (dx, dy)."""
    if elem is None:
        return None, None
    try:
        bb = elem.get_BoundingBox(None)
        if bb is None:
            return None, None
        dx = abs(float(bb.Max.X) - float(bb.Min.X))
        dy = abs(float(bb.Max.Y) - float(bb.Min.Y))
        if dx < 1e-9 and dy < 1e-9:
            return None, None
        return dx, dy
    except Exception:
        return None, None


def _section_dims_for_kind(elem, kind, document, elev):
    """(w_ft, d_ft) para preview de sección / metadatos."""
    span_u = float(elev["spanU_ft"])
    span_v = float(elev["spanV_ft"])
    w_ft, d_ft = _lookup_section_width_depth_ft(elem, document)

    if kind == u"foundation":
        px, py = _plan_xy_ft(elem)
        if w_ft is None or d_ft is None:
            if px is not None and py is not None:
                if w_ft is None:
                    w_ft = max(px, py)
                if d_ft is None:
                    d_ft = min(px, py) if min(px, py) > 1e-9 else max(px, py)
        if w_ft is None:
            w_ft = span_u
        if d_ft is None:
            d_ft = span_u
        return w_ft, d_ft

    if kind == u"floor":
        # Sección de losa: espesor × profundidad vista (o huella)
        th = _floor_thickness_ft(elem, document)
        if th is None:
            th = span_v
        # Ancho en preview = dimensión en planta hacia depth de vista
        if w_ft is None:
            w_ft = max(span_u, span_v)
        d_ft = th
        return w_ft, d_ft

    if kind == u"beam":
        # w = ancho, d = canto (como vigas)
        if w_ft is None:
            w_ft = min(span_u, span_v) if min(span_u, span_v) > 1e-9 else span_u
        if d_ft is None:
            d_ft = max(span_u, span_v)
        # En alzado típico span_v es el canto y span_u la longitud en vista
        if span_v > 1e-9 and (d_ft is None or d_ft < span_v * 0.5):
            # preferir canto visto si el param falta
            if d_ft is None:
                d_ft = span_v
        return w_ft, d_ft

    # column
    if w_ft is None:
        w_ft = span_u
    if d_ft is None:
        d_ft = w_ft
    return w_ft, d_ft


def _base_domain_dict(elem, kind, document, view, elev):
    prefix = _KIND_PREFIX.get(kind, u"E")
    eid = _element_id_int(elem)
    label = _mark_or_fallback(elem, prefix)
    span_u = float(elev["spanU_ft"])
    span_v = float(elev["spanV_ft"])
    w_ft, d_ft = _section_dims_for_kind(elem, kind, document, elev)
    w_cm = ft_to_cm(w_ft)
    d_cm = ft_to_cm(d_ft)
    type_lbl = _section_label_cm(w_cm, d_cm) or _type_name(document, elem)

    base_lv = top_lv = None
    if kind == u"column":
        base_lv, top_lv = _column_levels(document, elem)
    elif kind == u"floor":
        try:
            p = elem.get_Parameter(BuiltInParameter.LEVEL_PARAM)
            if p and p.HasValue:
                base_lv = _level_name(document, p.AsElementId())
        except Exception:
            pass

    is_column = kind == u"column"
    return {
        "id": u"{0}:{1}".format(prefix, eid if eid is not None else label),
        "elementIdInt": eid,
        "element": elem,
        "kind": kind,
        "label": label,
        "typeName": type_lbl,
        "widthCm": w_cm,
        "depthCm": d_cm,
        "heightM": ft_to_m(span_v),
        "levelBase": base_lv or u"",
        "levelTop": top_lv or u"",
        # Coordenadas de vista (pies) — fuente de verdad del canvas
        "uMin": float(elev["uMin"]),
        "uMax": float(elev["uMax"]),
        "vMin": float(elev["vMin"]),
        "vMax": float(elev["vMax"]),
        "uMid": float(elev["uMid"]),
        "vMid": float(elev["vMid"]),
        "spanU_ft": span_u,
        "spanV_ft": span_v,
        "nBarsX": 4 if is_column else (2 if kind == u"beam" else 0),
        "nBarsY": 4 if is_column else (2 if kind == u"beam" else 0),
        "diamLong": 22 if is_column else 16,
        "diamEstribo": 10 if is_column else 8,
        "coverMm": 40.0,
    }


def domain_members_from_selection(document, elements, view):
    """
    Convierte la selección en miembros de dominio con rect de elevación.

    Orden de lista: izquierda→derecha (uMid). El canvas pinta floors/fnd/beams
    antes que columnas para z-order de lectura.
    """
    groups = classify_elements(elements)
    members = []
    for kind in _KIND_ORDER:
        for el in groups.get(kind) or []:
            elev = elevation_rect_from_element(el, view)
            if elev is None:
                continue
            members.append(_base_domain_dict(el, kind, document, view, elev))

    members.sort(
        key=lambda m: (
            float(m.get("uMid") or 0.0),
            float(m.get("vMin") or 0.0),
            m.get("id") or u"",
        )
    )
    return members
