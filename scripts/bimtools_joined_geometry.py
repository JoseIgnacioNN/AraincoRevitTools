# -*- coding: utf-8 -*-
"""
Lectura robusta de **Unir geometría** (JoinGeometryUtils.GetJoinedElements).

``ICollection<ElementId>`` a veces no itera bien con ``for x in col`` en IronPython/CPython
de pyRevit; se usan ``Count`` + ``get_Item``/índice como respaldo.
"""

from Autodesk.Revit.DB import ElementId, JoinGeometryUtils


def _coerce_to_element_id_list(raw):
    if raw is None:
        return []
    out = []
    try:
        for jid in raw:
            if jid is not None and jid != ElementId.InvalidElementId:
                out.append(jid)
    except (TypeError, SystemError, AttributeError):
        pass
    if out:
        return out
    try:
        n = int(raw.Count)
    except Exception:
        n = 0
    for i in range(n):
        jid = None
        try:
            jid = raw[i]
        except Exception:
            try:
                jid = raw.get_Item(i)
            except Exception:
                jid = None
        if jid is not None and jid != ElementId.InvalidElementId:
            out.append(jid)
    return out


def get_joined_element_ids(doc, element):
    """
    Ids de elementos unidos al dado, en el orden que devuelve Revit.
    Prueba ``GetJoinedElements(doc, element)`` y ``(doc, element.Id)``.
    """
    if doc is None or element is None:
        return []
    raw = None
    for get_args in (
        lambda: JoinGeometryUtils.GetJoinedElements(doc, element),
        lambda: JoinGeometryUtils.GetJoinedElements(doc, element.Id),
    ):
        try:
            raw = get_args()
        except Exception:
            raw = None
        if raw is not None:
            break
    return _coerce_to_element_id_list(raw)
