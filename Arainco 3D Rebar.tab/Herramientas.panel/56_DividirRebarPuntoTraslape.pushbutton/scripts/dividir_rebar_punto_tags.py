# -*- coding: utf-8 -*-
"""
Etiquetado de barras divididas: familia de la etiqueta original + tipo según RebarShape.

Revit 2024+ | IronPython
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    ElementId,
    FilteredElementCollector,
    IndependentTag,
    Reference,
    TagMode,
    TagOrientation,
    XYZ,
)

from rebar_tag_shape_sync_core import (
    lookup_tag_type_id,
    rebar_shape_name_candidates,
    symbol_map_from_family,
    symbol_map_from_family_names,
    tag_rebar_int_if_match,
)


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _element_id_int(eid):
    if eid is None:
        return None
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid.Value)
        except Exception:
            return None


def _rebar_midpoint_xyz(rebar, bar_index=0):
    """Punto medio aproximado de la centerline (posición de cabeza de etiqueta)."""
    return _rebar_tag_anchor_xyz(rebar, bar_index)


def _primary_tag_bar_index(rebar, view):
    """Índice de barra representada en la vista (Show Middle, FirstLast, …)."""
    if rebar is None:
        return 0
    try:
        from dividir_rebar_punto_lap_detail import represented_bar_indices

        idxs = represented_bar_indices(rebar, view)
        if idxs:
            return int(idxs[0])
    except Exception:
        pass
    return 0


def _tag_bar_indices_for_view(rebar, view):
    """
    Índices donde colocar etiqueta(s) según PresentationMode.

    Middle/All → uno; FirstLast/Select → uno por barra representada.
    """
    if rebar is None:
        return [0]
    try:
        from dividir_rebar_punto_lap_detail import represented_bar_indices

        idxs = represented_bar_indices(rebar, view)
        if idxs:
            return [int(i) for i in idxs]
    except Exception:
        pass
    return [0]


def _iter_rebar_centerline_curves(rebar, bar_index=0):
    if rebar is None:
        return []
    try:
        from Autodesk.Revit.DB.Structure import MultiplanarOption

        curves = rebar.GetCenterlineCurves(
            False,
            True,
            True,
            MultiplanarOption.IncludeAllMultiplanarCurves,
            int(bar_index),
        )
        if curves is None:
            return []
        out = []
        try:
            for c in curves:
                if c is not None:
                    out.append(c)
        except Exception:
            try:
                n = int(curves.Count)
            except Exception:
                n = 0
            for i in range(n):
                try:
                    c = curves.get_Item(i)
                except Exception:
                    try:
                        c = curves[i]
                    except Exception:
                        c = None
                if c is not None:
                    out.append(c)
        return out
    except Exception:
        return []


def _curve_midpoint(crv):
    if crv is None:
        return None
    try:
        return crv.Evaluate(0.5, True)
    except Exception:
        pass
    try:
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
        return XYZ(
            (p0.X + p1.X) * 0.5,
            (p0.Y + p1.Y) * 0.5,
            (p0.Z + p1.Z) * 0.5,
        )
    except Exception:
        return None


def _rebar_tag_anchor_xyz(rebar, bar_index=0):
    """
    Ancla para posicionar etiqueta: punto medio del segmento más largo
    (segmento mayor / vano). Fallback: primer tramo o bbox.
    """
    curves = _iter_rebar_centerline_curves(rebar, bar_index)
    if curves:
        best = None
        best_len = -1.0
        for c in curves:
            try:
                ln = float(c.Length)
            except Exception:
                ln = 0.0
            if ln > best_len:
                best_len = ln
                best = c
        mid = _curve_midpoint(best)
        if mid is not None:
            return mid
    try:
        bb = rebar.get_BoundingBox(None)
        if bb is not None and bb.Min is not None and bb.Max is not None:
            return XYZ(
                (bb.Min.X + bb.Max.X) * 0.5,
                (bb.Min.Y + bb.Max.Y) * 0.5,
                (bb.Min.Z + bb.Max.Z) * 0.5,
            )
    except Exception:
        pass
    return None


def _xyz_sub(a, b):
    if a is None or b is None:
        return None
    try:
        return XYZ(float(a.X) - float(b.X), float(a.Y) - float(b.Y), float(a.Z) - float(b.Z))
    except Exception:
        return None


def _xyz_add(a, b):
    if a is None:
        return b
    if b is None:
        return a
    try:
        return XYZ(float(a.X) + float(b.X), float(a.Y) + float(b.Y), float(a.Z) + float(b.Z))
    except Exception:
        return a


def capture_rebar_tag_infos(doc, rebar_id):
    """
    IndependentTags que etiquetan ``rebar_id``.

    Cada info: type_id, view_id, head, orient, leader, family_name,
    anchor, bar_index (barra representada en esa vista), head_offset, rotation.
    """
    rid = _element_id_int(rebar_id)
    if rid is None or doc is None:
        return []
    rebar = None
    try:
        rebar = doc.GetElement(rebar_id)
    except Exception:
        rebar = None
    rebar_set = {rid}
    invalid = ElementId.InvalidElementId
    out = []
    try:
        coll = (
            FilteredElementCollector(doc)
            .OfClass(IndependentTag)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return out
    for tag in coll:
        if tag_rebar_int_if_match(tag, rebar_set, invalid) is None:
            continue
        info = {}
        try:
            info[u"type_id"] = tag.GetTypeId()
        except Exception:
            info[u"type_id"] = None
        try:
            info[u"view_id"] = tag.OwnerViewId
        except Exception:
            info[u"view_id"] = None
        try:
            info[u"head"] = tag.TagHeadPosition
        except Exception:
            info[u"head"] = None
        try:
            info[u"orient"] = tag.TagOrientation
        except Exception:
            info[u"orient"] = TagOrientation.Horizontal
        try:
            info[u"leader"] = bool(tag.HasLeader)
        except Exception:
            info[u"leader"] = True
        try:
            info[u"rotation"] = float(tag.RotationAngle)
        except Exception:
            info[u"rotation"] = None
        view = None
        try:
            if info.get(u"view_id") is not None:
                view = doc.GetElement(info[u"view_id"])
        except Exception:
            view = None
        bar_idx = _primary_tag_bar_index(rebar, view) if rebar is not None else 0
        anchor = (
            _rebar_tag_anchor_xyz(rebar, bar_idx) if rebar is not None else None
        )
        info[u"bar_index"] = int(bar_idx)
        info[u"anchor"] = anchor
        info[u"head_offset"] = _xyz_sub(info.get(u"head"), anchor)
        try:
            sym = doc.GetElement(info[u"type_id"])
            fam = sym.Family if sym is not None else None
            info[u"family_name"] = fam.Name if fam is not None else u""
            info[u"family_id"] = fam.Id if fam is not None else None
        except Exception:
            info[u"family_name"] = u""
            info[u"family_id"] = None
        if info.get(u"view_id") is None or info.get(u"type_id") is None:
            continue
        out.append(info)
    return out


def _symbol_map_for_tag_info(doc, info, cache):
    fam_id = info.get(u"family_id")
    fam_name = info.get(u"family_name") or u""
    cache_key = None
    if fam_id is not None:
        try:
            cache_key = int(fam_id.IntegerValue)
        except Exception:
            cache_key = None
    if cache_key is None and fam_name:
        cache_key = u"n:" + _as_unicode(fam_name).lower()
    if cache is not None and cache_key is not None and cache_key in cache:
        return cache[cache_key]

    sm = {}
    if fam_id is not None:
        try:
            fam = doc.GetElement(fam_id)
            if fam is not None:
                sm = symbol_map_from_family(doc, fam)
        except Exception:
            sm = {}
    if not sm and fam_name:
        sm = symbol_map_from_family_names(doc, [fam_name])
    if cache is not None and cache_key is not None:
        cache[cache_key] = sm
    return sm


def resolve_tag_type_id_for_rebar(doc, rebar, tag_info, symbol_map_cache=None):
    """
    Tipo de etiqueta: homónimo del RebarShape dentro de la familia de la etiqueta
    original; si no hay match, el type_id capturado.
    """
    fallback = tag_info.get(u"type_id") if tag_info else None
    if doc is None or rebar is None or not tag_info:
        return fallback
    sm = _symbol_map_for_tag_info(doc, tag_info, symbol_map_cache)
    if sm:
        for shape in rebar_shape_name_candidates(doc, rebar):
            tid = lookup_tag_type_id(sm, shape)
            if tid is not None:
                return tid
    return fallback


def _referencias_tag_rebar(doc, rebar, preferred_bar_index=None):
    refs = []
    seen = set()

    def _add(r):
        if r is None:
            return
        try:
            key = r.ConvertToStableRepresentation(doc)
        except Exception:
            key = id(r)
        if key in seen:
            return
        seen.add(key)
        refs.append(r)

    def _ref_at(idx):
        try:
            if hasattr(rebar, u"GetReferenceToBarPosition"):
                return rebar.GetReferenceToBarPosition(int(idx))
        except Exception:
            pass
        try:
            if hasattr(rebar, u"GetReferenceForBarPosition"):
                return rebar.GetReferenceForBarPosition(int(idx))
        except Exception:
            pass
        return None

    # Preferir la barra representada (Show Middle / FirstLast / …).
    if preferred_bar_index is not None:
        _add(_ref_at(preferred_bar_index))

    try:
        subs = rebar.GetSubelements() if hasattr(rebar, u"GetSubelements") else None
    except Exception:
        subs = None
    if subs:
        for sub in subs:
            if sub is None:
                continue
            try:
                sref = sub.GetReference() if hasattr(sub, u"GetReference") else None
                if sref is not None:
                    _add(sref)
            except Exception:
                pass
    try:
        npos = int(rebar.NumberOfBarPositions)
    except Exception:
        npos = 0
    if npos > 0:
        for idx in (0, max(0, npos - 1)):
            if preferred_bar_index is not None and int(idx) == int(preferred_bar_index):
                continue
            _add(_ref_at(idx))
    try:
        _add(Reference(rebar))
    except Exception:
        pass
    return refs


def create_rebar_independent_tag(
    doc,
    view,
    rebar,
    type_id,
    head,
    orient,
    add_leader,
    rotation=None,
    bar_index=None,
):
    if view is None or type_id is None or rebar is None:
        return None
    if bar_index is None:
        bar_index = _primary_tag_bar_index(rebar, view)
    bi = int(bar_index)
    if head is None:
        head = _rebar_tag_anchor_xyz(rebar, bi)
    if head is None:
        return None
    try:
        sym = doc.GetElement(type_id)
        if sym is not None and not sym.IsActive:
            sym.Activate()
    except Exception:
        pass
    if orient is None:
        orient = TagOrientation.Horizontal
    refs = _referencias_tag_rebar(doc, rebar, preferred_bar_index=bi)
    if not refs:
        return None
    for ref in refs:
        tag = None
        try:
            tag = IndependentTag.Create(
                doc, type_id, view.Id, ref, bool(add_leader), orient, head
            )
        except Exception:
            tag = None
        if tag is None:
            try:
                tag = IndependentTag.Create(
                    doc,
                    view.Id,
                    ref,
                    bool(add_leader),
                    TagMode.TM_ADDBY_CATEGORY,
                    orient,
                    head,
                )
                if tag is not None:
                    try:
                        tag.ChangeTypeId(type_id)
                    except Exception:
                        try:
                            tag.SetTypeId(type_id)
                        except Exception:
                            pass
            except Exception:
                tag = None
        if tag is None:
            continue
        # Forzar alineamiento capturado (Create a veces lo reescribe).
        try:
            tag.TagOrientation = orient
        except Exception:
            pass
        try:
            tag.TagHeadPosition = head
        except Exception:
            pass
        if rotation is not None:
            try:
                tag.RotationAngle = float(rotation)
            except Exception:
                pass
        return tag
    return None


def _head_for_divided_rebar(rebar, tag_info, view=None, bar_index=None):
    """
    Cabeza de etiqueta: ancla del tramo (barra representada) + offset relativo
    de la etiqueta original (mismo lado / distancia / alineamiento).
    """
    if bar_index is None:
        if view is not None:
            bar_index = _primary_tag_bar_index(rebar, view)
        else:
            try:
                bar_index = int(tag_info.get(u"bar_index"))
            except Exception:
                bar_index = 0
    anchor = _rebar_tag_anchor_xyz(rebar, int(bar_index))
    off = tag_info.get(u"head_offset") if tag_info else None
    if anchor is not None and off is not None:
        return _xyz_add(anchor, off)
    if tag_info and tag_info.get(u"head") is not None and anchor is not None:
        # Sin offset capturable: mantener orientación relativa nula (sobre ancla)
        return anchor
    return (tag_info or {}).get(u"head") or anchor


def tag_divided_rebars(doc, tag_infos, new_rebars):
    """
    Crea IndependentTags para cada barra nueva en las vistas de las etiquetas
    originales, eligiendo el tipo por shape dentro de la misma familia.

    Posiciona en barras representadas (Show Middle / FirstLast / Select).
    Respeta ``TagOrientation``, rotación y offset cabeza↔barra de la etiqueta
    original.

    Returns:
        int: etiquetas creadas.
    """
    if doc is None or not tag_infos or not new_rebars:
        return 0
    creadas = 0
    cache = {}
    for rb in new_rebars:
        if rb is None:
            continue
        try:
            rb = doc.GetElement(rb.Id)
        except Exception:
            pass
        if rb is None:
            continue
        # Una pasada por vista capturada (dedupe por view_id)
        seen_views = set()
        for info in tag_infos:
            vid = info.get(u"view_id")
            try:
                vkey = int(vid.IntegerValue)
            except Exception:
                vkey = None
            if vkey is not None and vkey in seen_views:
                continue
            if vkey is not None:
                seen_views.add(vkey)
            view = doc.GetElement(vid) if vid is not None else None
            if view is None:
                continue
            type_id = resolve_tag_type_id_for_rebar(doc, rb, info, cache)
            if type_id is None:
                continue
            bar_indices = _tag_bar_indices_for_view(rb, view)
            for bar_idx in bar_indices:
                head = _head_for_divided_rebar(rb, info, view, bar_idx)
                tag = create_rebar_independent_tag(
                    doc,
                    view,
                    rb,
                    type_id,
                    head,
                    info.get(u"orient", TagOrientation.Horizontal),
                    info.get(u"leader", True),
                    rotation=info.get(u"rotation"),
                    bar_index=bar_idx,
                )
                if tag is not None:
                    creadas += 1
    return creadas
