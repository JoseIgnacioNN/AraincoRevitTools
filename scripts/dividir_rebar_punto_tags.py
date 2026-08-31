# -*- coding: utf-8 -*-
"""
Etiquetado de barras divididas: familia de la etiqueta original + tipo según RebarShape.

Tras dividir: IndependentTag (familia EST_A si no había etiqueta) + MRA
«Recorrido Barras» en la vista activa para marcar el recorrido.

Revit 2024+ | IronPython
"""

from __future__ import print_function

import os
import sys

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

# Misma convención que zapata de muro / armadura de losa.
_DEFAULT_TAG_FAMILY_NAMES = (
    u"EST_A_STRUCTURAL REBAR TAG",
    u"EST_A_STRUCTURAL REBAR TAG_FLOOR",
)
_FALLBACK_TAG_TYPE = u"01"
_MRA_TYPE_NAME_RECORRIDO = u"Recorrido Barras"


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
    rid = _element_id_int(rebar.Id) if rebar is not None else None
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
    for i_ref, ref in enumerate(refs):
        tag = None
        try:
            tag = IndependentTag.Create(
                doc, type_id, view.Id, ref, bool(add_leader), orient, head
            )
        except Exception as ex:
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
            except Exception as ex:
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
        rid = _element_id_int(rb.Id)
        shapes = []
        try:
            shapes = list(rebar_shape_name_candidates(doc, rb) or [])
        except Exception:
            shapes = []
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


def _find_extension_scripts_dir():
    """Localiza ``BIMTools.extension/scripts`` (helpers compartidos MRA)."""
    cursor = os.path.dirname(os.path.abspath(__file__))
    for _ in range(24):
        marker = os.path.join(cursor, u"geometria_estribos_viga.py")
        if os.path.isfile(marker):
            return cursor
        nested = os.path.join(cursor, u"scripts", u"geometria_estribos_viga.py")
        if os.path.isfile(nested):
            return os.path.join(cursor, u"scripts")
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


def _ensure_extension_scripts_on_path():
    ext = _find_extension_scripts_dir()
    if ext and ext not in sys.path:
        try:
            sys.path.append(ext)
        except Exception:
            pass
    return ext


def _view_ok_for_annotation(view):
    if view is None:
        return False
    try:
        if bool(view.IsTemplate):
            return False
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import View3D, ViewSheet, ViewSchedule

        if isinstance(view, (View3D, ViewSheet, ViewSchedule)):
            return False
    except Exception:
        pass
    return True


def _default_tag_type_id_for_rebar(doc, rebar, symbol_map_cache=None):
    """
    Tipo EST_A homónimo al RebarShape; fallback «01».
    """
    if doc is None or rebar is None:
        return None
    cache = symbol_map_cache if symbol_map_cache is not None else {}
    cache_key = u"__default_est_a__"
    sm = cache.get(cache_key)
    if sm is None:
        sm = symbol_map_from_family_names(doc, list(_DEFAULT_TAG_FAMILY_NAMES)) or {}
        cache[cache_key] = sm
    if not sm:
        return None
    shapes = []
    try:
        shapes = list(rebar_shape_name_candidates(doc, rebar) or [])
    except Exception:
        shapes = []
    for shape in shapes:
        tid = lookup_tag_type_id(sm, shape)
        if tid is not None:
            return tid
    fb = lookup_tag_type_id(sm, _FALLBACK_TAG_TYPE)
    return fb


def tag_rebars_with_default_family(doc, view, new_rebars):
    """
    Una IndependentTag por barra (y por posición representada) con familia EST_A.
    Sin leader; cabecera en ancla de la barra.
    """
    if doc is None or view is None or not new_rebars:
        return 0
    if not _view_ok_for_annotation(view):
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
        type_id = _default_tag_type_id_for_rebar(doc, rb, cache)
        if type_id is None:
            continue
        bar_indices = _tag_bar_indices_for_view(rb, view)
        for bar_idx in bar_indices:
            head = _rebar_tag_anchor_xyz(rb, bar_idx)
            tag = create_rebar_independent_tag(
                doc,
                view,
                rb,
                type_id,
                head,
                TagOrientation.Horizontal,
                False,
                rotation=None,
                bar_index=bar_idx,
            )
            if tag is not None:
                try:
                    tag.HasLeader = False
                except Exception:
                    pass
                try:
                    if head is not None:
                        tag.TagHeadPosition = head
                except Exception:
                    pass
                creadas += 1
    return creadas


def _apply_mra_local_fallback(doc, view, rebars, avisos):
    """
    MRA «Recorrido Barras» sin depender de geometria_estribos_viga
    (mismo criterio básico: centro bbox + offset según UpDirection).
    """
    if avisos is None:
        avisos = []
    try:
        from Autodesk.Revit.DB import (
            BuiltInParameter,
            DimensionStyleType,
            MultiReferenceAnnotation,
            MultiReferenceAnnotationOptions,
            MultiReferenceAnnotationType,
            StorageType,
        )
        from System.Collections.Generic import List
    except Exception as ex:
        avisos.append(u"MRA: imports fallaron ({0}).".format(_as_unicode(ex)))
        return 0

    def _norm(s):
        try:
            return u" ".join(_as_unicode(s).replace(u"\u00A0", u" ").split())
        except Exception:
            return u""

    mrat = None
    type_names = []
    try:
        col = FilteredElementCollector(doc).OfClass(MultiReferenceAnnotationType)
        tgt = _norm(_MRA_TYPE_NAME_RECORRIDO).lower()
        for t in col:
            names = []
            try:
                names.append(_norm(getattr(t, u"Name", None)))
            except Exception:
                pass
            for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
                try:
                    bip = getattr(BuiltInParameter, bip_name, None)
                    if bip is None:
                        continue
                    p = t.get_Parameter(bip)
                    if p is None or not p.HasValue:
                        continue
                    if p.StorageType == StorageType.String:
                        names.append(_norm(p.AsString()))
                except Exception:
                    continue
            for n in names:
                if n and n not in type_names:
                    type_names.append(n)
                if n and n.lower() == tgt:
                    mrat = t
                    break
            if mrat is not None:
                break
    except Exception as ex:
        avisos.append(u"MRA: búsqueda de tipo ({0}).".format(_as_unicode(ex)))
        return 0
    if mrat is None:
        avisos.append(
            u"Multi-Rebar Annotation: no existe el tipo «{0}» en el proyecto.".format(
                _MRA_TYPE_NAME_RECORRIDO
            )
        )
        return 0

    try:
        vd = view.ViewDirection.Normalize()
        rd = view.RightDirection.Normalize()
        v_up = view.UpDirection
        if v_up is None or v_up.GetLength() < 1e-12:
            v_up = XYZ.BasisZ
        else:
            v_up = v_up.Normalize()
    except Exception:
        avisos.append(u"MRA: dirección de vista inválida.")
        return 0

    # ~300 mm hacia −Up (fuera del conjunto).
    try:
        from Autodesk.Revit.DB import UnitTypeId, UnitUtils

        off_ft = float(
            UnitUtils.ConvertToInternalUnits(300.0, UnitTypeId.Millimeters)
        )
    except Exception:
        off_ft = 300.0 / 304.8

    n_ok = 0
    for rb in rebars or []:
        if rb is None:
            continue
        try:
            rb = doc.GetElement(rb.Id)
        except Exception:
            pass
        if rb is None:
            continue
        rid = _element_id_int(rb.Id)
        p_mid = None
        try:
            bb = rb.get_BoundingBox(view)
            if bb is not None:
                p_mid = (bb.Min + bb.Max) * 0.5
        except Exception:
            p_mid = None
        if p_mid is None:
            try:
                bb0 = rb.get_BoundingBox(None)
                if bb0 is not None:
                    p_mid = (bb0.Min + bb0.Max) * 0.5
            except Exception:
                p_mid = None
        if p_mid is None:
            continue
        try:
            p_line = p_mid - v_up.Multiply(float(off_ft))
        except Exception:
            p_line = p_mid
        try:
            opts = MultiReferenceAnnotationOptions(mrat)
        except Exception:
            try:
                opts = MultiReferenceAnnotationOptions()
                opts.MultiReferenceAnnotationType = mrat.Id
            except Exception as ex:
                continue
        try:
            opts.DimensionStyleType = DimensionStyleType.Linear
        except Exception:
            pass
        try:
            opts.DimensionPlaneNormal = vd
            opts.DimensionLineDirection = rd
            opts.DimensionLineOrigin = p_line
            opts.TagHeadPosition = p_line
            opts.TagHasLeader = False
        except Exception as ex:
            continue
        ids = List[ElementId]()
        ids.Add(rb.Id)
        try:
            opts.SetElementsToDimension(ids)
        except Exception as ex:
            continue
        try:
            if hasattr(opts, u"ElementsMatchReferenceCategory"):
                if not opts.ElementsMatchReferenceCategory(doc):
                    continue
        except Exception:
            pass
        try:
            mra = MultiReferenceAnnotation.Create(doc, view.Id, opts)
            if mra is not None:
                n_ok += 1
        except Exception as ex:
            avisos.append(
                u"MRA Id {0}: {1}".format(rid, _as_unicode(ex))
            )
    return int(n_ok)


def _list_mra_type_names(doc):
    """Nombres de tipos MultiReferenceAnnotation en el proyecto."""
    names = []
    try:
        from Autodesk.Revit.DB import (
            BuiltInParameter,
            MultiReferenceAnnotationType,
            StorageType,
        )
    except Exception:
        return names

    def _norm(s):
        try:
            return u" ".join(_as_unicode(s).replace(u"\u00A0", u" ").split())
        except Exception:
            return u""

    try:
        col = FilteredElementCollector(doc).OfClass(MultiReferenceAnnotationType)
        for t in col:
            try:
                n = _norm(getattr(t, u"Name", None))
                if n and n not in names:
                    names.append(n)
            except Exception:
                pass
            for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
                try:
                    bip = getattr(BuiltInParameter, bip_name, None)
                    if bip is None:
                        continue
                    p = t.get_Parameter(bip)
                    if p is None or not p.HasValue:
                        continue
                    if p.StorageType == StorageType.String:
                        n = _norm(p.AsString())
                        if n and n not in names:
                            names.append(n)
                except Exception:
                    continue
    except Exception:
        pass
    return names


def apply_mra_recorrido_barras(doc, view, new_rebars, avisos=None):
    """
    Una MultiReferenceAnnotation «Recorrido Barras» por cada Rebar nuevo.
    Preferir helper compartido; si no, fallback local.
    """
    if avisos is None:
        avisos = []
    if doc is None or view is None or not new_rebars:
        return 0
    if not _view_ok_for_annotation(view):
        avisos.append(
            u"MRA «{0}»: use planta/alzado/sección (no plantilla ni 3D).".format(
                _MRA_TYPE_NAME_RECORRIDO
            )
        )
        return 0

    type_names = _list_mra_type_names(doc)

    _ensure_extension_scripts_on_path()
    try:
        from geometria_estribos_viga import (
            crear_multi_rebar_annotations_por_nombre_tipo,
        )

        n = int(
            crear_multi_rebar_annotations_por_nombre_tipo(
                doc,
                view,
                list(new_rebars),
                avisos,
                _MRA_TYPE_NAME_RECORRIDO,
            )
            or 0
        )
        return n
    except Exception as ex:
        return _apply_mra_local_fallback(doc, view, new_rebars, avisos)


def annotate_divided_rebars(doc, view, new_rebars, tag_infos=None):
    """
    Etiqueta los tramos nuevos y aplica MRA de recorrido.

    1. Si había IndependentTag en el original → recrear en esas vistas.
    2. Si no se creó ninguna → etiqueta EST_A en la vista activa.
    3. Siempre intenta MRA «Recorrido Barras» en la vista activa.

    Returns:
        dict: n_tags, n_mra, avisos, used_default_tags
    """
    result = {
        u"n_tags": 0,
        u"n_mra": 0,
        u"avisos": [],
        u"used_default_tags": False,
    }
    if doc is None or not new_rebars:
        return result

    rebars = []
    for rb in new_rebars:
        if rb is None:
            continue
        try:
            rebars.append(doc.GetElement(rb.Id))
        except Exception:
            rebars.append(rb)
    rebars = [r for r in rebars if r is not None]
    if not rebars:
        return result

    n_tags = 0
    if tag_infos:
        try:
            n_tags = int(
                tag_divided_rebars(doc, tag_infos, rebars) or 0
            )
        except Exception as ex:
            result[u"avisos"].append(
                u"Etiquetas (recreación): {0}".format(_as_unicode(ex))
            )
            n_tags = 0

    if n_tags <= 0 and view is not None:
        try:
            n_def = int(
                tag_rebars_with_default_family(doc, view, rebars) or 0
            )
        except Exception as ex:
            result[u"avisos"].append(
                u"Etiquetas (EST_A): {0}".format(_as_unicode(ex))
            )
            n_def = 0
        if n_def > 0:
            n_tags = n_def
            result[u"used_default_tags"] = True
        elif not tag_infos:
            result[u"avisos"].append(
                u"No se crearon etiquetas (falta familia «{0}» o vista no válida).".format(
                    _DEFAULT_TAG_FAMILY_NAMES[0]
                )
            )

    result[u"n_tags"] = int(n_tags)

    avisos_mra = []
    try:
        n_mra = int(
            apply_mra_recorrido_barras(doc, view, rebars, avisos_mra)
            or 0
        )
    except Exception as ex:
        n_mra = 0
        avisos_mra.append(u"MRA: {0}".format(_as_unicode(ex)))
    result[u"n_mra"] = int(n_mra)
    for av in avisos_mra:
        if av:
            result[u"avisos"].append(av)
    return result