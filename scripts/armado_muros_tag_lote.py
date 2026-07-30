# -*- coding: utf-8 -*-
"""
Lote de etiquetas por muro: wall tag (mono-host) + malla V/H en la misma vista.

Usado por el DMU de co-movimiento y (futuro) selector/move manual.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
    IndependentTag,
    XYZ,
)
from Autodesk.Revit.DB.Structure import Rebar

try:
    unicode
except NameError:
    unicode = str

try:
    from armado_muros_malla_rebar_tags import MALLA_REBAR_TAG_FAMILY_NAME
except Exception:
    MALLA_REBAR_TAG_FAMILY_NAME = u"EST_A_STRUCTURAL REBAR TAG_MALLA"

_EPS = 1e-9


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


def _norm_name(s):
    try:
        t = unicode(s).strip().lower()
    except Exception:
        try:
            t = str(s or u"").strip().lower()
        except Exception:
            return u""
    for ch in (u"\xa0", u"\u200b", u"\ufeff"):
        t = t.replace(ch, u"")
    return u" ".join(t.split())


def _tagged_ids(tag):
    out = []
    if tag is None:
        return out
    for getter in (
        lambda: tag.GetTaggedLocalElementIds(),
        lambda: tag.GetTaggedElementIds(),
    ):
        try:
            ids = getter()
        except Exception:
            ids = None
        if ids is None:
            continue
        try:
            for eid in ids:
                if eid is not None and eid != ElementId.InvalidElementId:
                    out.append(eid)
        except Exception:
            try:
                n = int(ids.Count)
                for i in range(n):
                    eid = ids[i]
                    if eid is not None and eid != ElementId.InvalidElementId:
                        out.append(eid)
            except Exception:
                pass
        if out:
            break
    return out


def _cat_int(el):
    try:
        cat = el.Category
        if cat is None:
            return None
        return int(cat.Id.IntegerValue)
    except Exception:
        return None


def _is_wall_tag(tag):
    try:
        return _cat_int(tag) == int(BuiltInCategory.OST_WallTags)
    except Exception:
        return False


def _is_rebar_tag(tag):
    try:
        return _cat_int(tag) == int(BuiltInCategory.OST_RebarTags)
    except Exception:
        return False


def _tag_family_name(doc, tag):
    if doc is None or tag is None:
        return u""
    try:
        tid = tag.GetTypeId()
        sym = doc.GetElement(tid)
        if sym is None:
            return u""
        fam = getattr(sym, "Family", None)
        if fam is None:
            return u""
        return unicode(fam.Name or u"")
    except Exception:
        return u""


def is_malla_rebar_tag(doc, tag):
    if not _is_rebar_tag(tag):
        return False
    fam = _norm_name(_tag_family_name(doc, tag))
    want = _norm_name(MALLA_REBAR_TAG_FAMILY_NAME)
    if not fam or not want:
        return False
    return fam == want or want in fam or fam in want


def is_lote_relevant_tag(doc, tag):
    """Wall tag o etiqueta de malla (familia MALLA)."""
    if tag is None or not isinstance(tag, IndependentTag):
        return False
    if _is_wall_tag(tag):
        return True
    return is_malla_rebar_tag(doc, tag)


def _owner_view_id(tag):
    try:
        return tag.OwnerViewId
    except Exception:
        return None


def _head(tag):
    try:
        return tag.TagHeadPosition
    except Exception:
        return None


def _set_head(tag, xyz):
    if tag is None or xyz is None:
        return False
    try:
        tag.TagHeadPosition = xyz
        return True
    except Exception:
        return False


def _xyz_delta(a, b):
    if a is None or b is None:
        return None
    try:
        return XYZ(b.X - a.X, b.Y - a.Y, b.Z - a.Z)
    except Exception:
        return None


def _xyz_add(p, d):
    if p is None or d is None:
        return None
    try:
        return XYZ(p.X + d.X, p.Y + d.Y, p.Z + d.Z)
    except Exception:
        return None


def _xyz_len2(d):
    if d is None:
        return 0.0
    try:
        return float(d.X * d.X + d.Y * d.Y + d.Z * d.Z)
    except Exception:
        return 0.0


def delta_is_significant(d, tol=1e-6):
    return _xyz_len2(d) > (tol * tol)


def _rebar_host_wall_id(doc, rebar):
    if doc is None or rebar is None:
        return None
    try:
        if not isinstance(rebar, Rebar):
            return None
    except Exception:
        return None
    try:
        hid = rebar.GetHostId()
    except Exception:
        return None
    host = None
    try:
        host = doc.GetElement(hid)
    except Exception:
        host = None
    if host is None:
        return None
    try:
        from Autodesk.Revit.DB import Wall
        if isinstance(host, Wall):
            return _eid_int(host.Id)
    except Exception:
        pass
    # AreaReinforcement u otro: subir un nivel de host
    try:
        hid2 = host.GetHostId()
        host2 = doc.GetElement(hid2)
        from Autodesk.Revit.DB import Wall
        if host2 is not None and isinstance(host2, Wall):
            return _eid_int(host2.Id)
    except Exception:
        pass
    return None


def _wall_id_from_wall_tag(doc, tag):
    tagged = _tagged_ids(tag)
    ints = []
    for eid in tagged:
        k = _eid_int(eid)
        if k is None:
            continue
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        try:
            from Autodesk.Revit.DB import Wall
            if el is not None and isinstance(el, Wall):
                ints.append(k)
        except Exception:
            pass
    uniq = list(set(ints))
    if len(uniq) == 1:
        return uniq[0], False
    if len(uniq) > 1:
        return None, True  # multi-host
    return None, False


def _wall_id_from_malla_tag(doc, tag):
    for eid in _tagged_ids(tag):
        try:
            el = doc.GetElement(eid)
        except Exception:
            continue
        wid = _rebar_host_wall_id(doc, el)
        if wid is not None:
            return wid
    return None


def wall_id_for_lote_tag(doc, tag):
    """
    Id del muro del lote, o ``None``.

    Retorna ``(wall_id_int, is_multihost_wall_tag)``.
    """
    if doc is None or tag is None:
        return None, False
    if _is_wall_tag(tag):
        return _wall_id_from_wall_tag(doc, tag)
    if is_malla_rebar_tag(doc, tag):
        return _wall_id_from_malla_tag(doc, tag), False
    return None, False


def _collect_tags_in_view(doc, view_id):
    out = []
    if doc is None or view_id is None:
        return out
    try:
        col = (
            FilteredElementCollector(doc, view_id)
            .OfClass(IndependentTag)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return out
    for tag in col:
        if tag is None or not isinstance(tag, IndependentTag):
            continue
        try:
            if tag.OwnerViewId != view_id:
                continue
        except Exception:
            pass
        out.append(tag)
    return out


def collect_malla_tags_for_wall(doc, view_id, wall_id):
    """IndependentTag malla en la vista cuyos rebars tienen host = wall_id."""
    hits = []
    if wall_id is None:
        return hits
    for tag in _collect_tags_in_view(doc, view_id):
        if not is_malla_rebar_tag(doc, tag):
            continue
        wid = _wall_id_from_malla_tag(doc, tag)
        if wid == wall_id:
            hits.append(tag)
    return hits


def collect_mono_wall_tags_for_wall(doc, view_id, wall_id):
    hits = []
    if wall_id is None:
        return hits
    for tag in _collect_tags_in_view(doc, view_id):
        if not _is_wall_tag(tag):
            continue
        wid, is_mh = _wall_id_from_wall_tag(doc, tag)
        if is_mh or wid != wall_id:
            continue
        hits.append(tag)
    return hits


def resolve_tag_lote(doc, tag):
    """
    Lista de ``IndependentTag`` del lote (muro mono-host + mallas) en la misma vista.

    Vacía si no se puede resolver muro, o wall tag multi-host.
    """
    if doc is None or tag is None or not isinstance(tag, IndependentTag):
        return []
    if not is_lote_relevant_tag(doc, tag):
        return []
    view_id = _owner_view_id(tag)
    if view_id is None or view_id == ElementId.InvalidElementId:
        return []
    wall_id, is_mh = wall_id_for_lote_tag(doc, tag)
    if is_mh:
        return []
    if wall_id is None:
        return []
    lote = []
    seen = set()
    for t in collect_mono_wall_tags_for_wall(doc, view_id, wall_id):
        k = _eid_int(t.Id)
        if k is None or k in seen:
            continue
        seen.add(k)
        lote.append(t)
    for t in collect_malla_tags_for_wall(doc, view_id, wall_id):
        k = _eid_int(t.Id)
        if k is None or k in seen:
            continue
        seen.add(k)
        lote.append(t)
    return lote


def resolve_tag_lote_ids(doc, tag):
    return [_eid_int(t.Id) for t in resolve_tag_lote(doc, tag) if _eid_int(t.Id) is not None]


def move_tag_lote_by_delta(doc, tags, delta, skip_ids=None):
    """
    Desplaza ``TagHeadPosition`` de cada tag en ``tags`` (excepto ``skip_ids``).

    :returns: nº de etiquetas movidas.
    """
    if doc is None or not tags or delta is None:
        return 0
    if not delta_is_significant(delta):
        return 0
    skip = set(skip_ids or [])
    n = 0
    for tag in tags:
        if tag is None:
            continue
        k = _eid_int(tag.Id)
        if k is not None and k in skip:
            continue
        h = _head(tag)
        nh = _xyz_add(h, delta)
        if nh is None:
            continue
        if _set_head(tag, nh):
            n += 1
    return n
