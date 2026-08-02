# -*- coding: utf-8 -*-
"""
Helpers de recolección / caché para Elevación Eje (contorno + etiquetas).

Un pase por vista: hormigón visible, muros/vigas paralelos, índice de tags.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    FamilyInstance,
    FilteredElementCollector,
    IndependentTag,
    Plane,
    Wall,
    XYZ,
)

from contorno_material_concrete import (
    _CATS_ESCANEO_MATERIAL_ESTRUCTURAL,
    material_estructural_es_concrete,
)

_TOL_MURO_PARALELO = 0.92
_TOL_VIGA_EJE_VS_VIEWDIR = 0.35


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _vector_unitario(v):
    if v is None:
        return None
    try:
        ln = float(v.GetLength())
        if ln < 1e-12:
            return None
        return XYZ(v.X / ln, v.Y / ln, v.Z / ln)
    except Exception:
        return None


def _eid_int(eid):
    if eid is None:
        return None
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid.Value)
        except Exception:
            return None


def plano_desde_vista(view):
    """Plano de la vista (ViewDirection + Origin)."""
    if view is None:
        return None
    try:
        return Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin)
    except Exception:
        return None


class MaterialConcreteCache(object):
    """Caché de ``material_estructural_es_concrete`` por ElementId."""

    def __init__(self):
        self._cache = {}

    def es_concrete(self, elem):
        if elem is None:
            return False
        key = _eid_int(getattr(elem, u"Id", None))
        if key is not None and key in self._cache:
            return self._cache[key]
        ok = bool(material_estructural_es_concrete(elem))
        if key is not None:
            self._cache[key] = ok
        return ok


def muro_paralelo_a_vista(wall, view, tol=_TOL_MURO_PARALELO):
    if wall is None or view is None:
        return False
    try:
        ori = _vector_unitario(wall.Orientation)
        vd = _vector_unitario(view.ViewDirection)
    except Exception:
        return False
    if ori is None or vd is None:
        return False
    try:
        return abs(float(ori.DotProduct(vd))) >= float(tol)
    except Exception:
        return False


def _direccion_eje_viga(beam):
    try:
        loc = beam.Location
        curve = loc.Curve if loc is not None else None
        if curve is None or not curve.IsBound:
            return None
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        return _vector_unitario(p1.Subtract(p0))
    except Exception:
        return None


def _es_viga(el):
    if el is None or not isinstance(el, FamilyInstance):
        return False
    try:
        from Autodesk.Revit.DB.Structure import StructuralType

        st = el.StructuralType
        if st == StructuralType.Beam:
            return True
        if st == StructuralType.Brace or st == StructuralType.Column:
            return False
    except Exception:
        pass
    try:
        return el.Category.Id.IntegerValue == int(
            BuiltInCategory.OST_StructuralFraming
        )
    except Exception:
        return False


def viga_paralela_a_vista(beam, view, tol=_TOL_VIGA_EJE_VS_VIEWDIR):
    if beam is None or view is None:
        return False
    eje = _direccion_eje_viga(beam)
    try:
        vd = _vector_unitario(view.ViewDirection)
    except Exception:
        vd = None
    if eje is None or vd is None:
        return False
    try:
        return abs(float(eje.DotProduct(vd))) <= float(tol)
    except Exception:
        return False


def recoger_concrete_en_vista(document, view, material_cache=None):
    """
    Un pase: hosts Concrete + muros/vigas paralelos al plano de la vista.

    Returns:
        dict con ``hosts``, ``muros``, ``vigas``.
    """
    out = {u"hosts": [], u"muros": [], u"vigas": []}
    if document is None or view is None:
        return out
    cache = material_cache if material_cache is not None else MaterialConcreteCache()
    seen = set()

    for cat in _CATS_ESCANEO_MATERIAL_ESTRUCTURAL:
        try:
            col = (
                FilteredElementCollector(document, view.Id)
                .OfCategory(cat)
                .WhereElementIsNotElementType()
            )
        except Exception:
            continue
        for el in col:
            if el is None:
                continue
            try:
                eid = _eid_int(el.Id)
            except Exception:
                eid = None
            if eid is not None and eid in seen:
                continue
            if not cache.es_concrete(el):
                continue
            if eid is not None:
                seen.add(eid)
            out[u"hosts"].append(el)
            try:
                if isinstance(el, Wall) and muro_paralelo_a_vista(el, view):
                    out[u"muros"].append(el)
            except Exception:
                pass
            try:
                if _es_viga(el) and viga_paralela_a_vista(el, view):
                    out[u"vigas"].append(el)
            except Exception:
                pass
    return out


def _tagged_host_ids(tag):
    ids = []
    if tag is None:
        return ids
    for meth in (u"GetTaggedLocalElementIds", u"GetTaggedElementIds"):
        try:
            coll = getattr(tag, meth)()
        except Exception:
            coll = None
        if coll is None:
            continue
        try:
            for obj in coll:
                k = _eid_int(obj)
                if k is None:
                    try:
                        k = _eid_int(getattr(obj, u"HostElementId", None))
                    except Exception:
                        k = None
                if k is not None:
                    ids.append(k)
        except Exception:
            pass
        if ids:
            return ids
    return ids


def indexar_tags_por_host(document, view):
    """
    Set de ElementId (int) de hosts ya etiquetados en la vista
    (Wall Tags + Structural Framing Tags).
    """
    tagged = set()
    if document is None or view is None:
        return tagged
    try:
        view_id = view.Id
    except Exception:
        return tagged

    cats = (
        BuiltInCategory.OST_WallTags,
        BuiltInCategory.OST_StructuralFramingTags,
    )
    for cat in cats:
        try:
            col = (
                FilteredElementCollector(document, view_id)
                .OfClass(IndependentTag)
                .OfCategory(cat)
            )
        except Exception:
            try:
                col = FilteredElementCollector(document, view_id).OfClass(
                    IndependentTag
                )
            except Exception:
                continue
        for tag in col:
            if tag is None:
                continue
            for hid in _tagged_host_ids(tag):
                tagged.add(hid)
    return tagged
