# -*- coding: utf-8 -*-
"""
Lectura y recolección de elementos por ``Armadura_Conjunto_GUID``.

Vendor mínimo de ``armado_muros_rebar_params`` (Armado Muros) para uso portable
sin depender de ``34_ArmadoMuros.pushbutton``.
"""

from Autodesk.Revit.DB import (
    BuiltInCategory,
    DetailCurve,
    FamilyInstance,
    FilledRegion,
    FilteredElementCollector,
    TextNote,
)
from Autodesk.Revit.DB.Structure import Rebar

ARMADURA_CONJUNTO_GUID_PARAM = u"Armadura_Conjunto_GUID"


def _norm_param_def_name(name):
    if name is None:
        return u""
    try:
        t = unicode(name).replace(u"\u00A0", u" ").strip()
    except Exception:
        try:
            t = str(name).strip()
        except Exception:
            return u""
    return t


def _iter_element_parameters(element):
    if element is None:
        return
    try:
        for p in element.Parameters:
            yield p
    except Exception:
        pass


def _find_element_parameter(element, param_name):
    if element is None or not param_name:
        return None
    target = _norm_param_def_name(param_name).lower()
    if not target:
        return None
    try:
        p = element.LookupParameter(param_name)
        if p is not None:
            return p
    except Exception:
        pass
    try:
        for p in _iter_element_parameters(element):
            if p is None:
                continue
            try:
                dn = _norm_param_def_name(p.Definition.Name).lower()
            except Exception:
                continue
            if dn == target:
                return p
    except Exception:
        pass
    return None


def get_armadura_conjunto_guid(element):
    """Lee ``Armadura_Conjunto_GUID`` de un elemento o ``None``."""
    if element is None:
        return None
    p = _find_element_parameter(element, ARMADURA_CONJUNTO_GUID_PARAM)
    if p is None:
        return None
    val = None
    try:
        val = p.AsString()
    except Exception:
        pass
    if not val:
        try:
            val = p.AsValueString()
        except Exception:
            pass
    if not val:
        return None
    try:
        t = unicode(val).strip()
    except Exception:
        try:
            t = str(val or u"").strip()
        except Exception:
            return None
    return t or None


def _normalize_conjunto_guid_target(conjunto_guid):
    if not conjunto_guid:
        return None
    try:
        target = unicode(conjunto_guid).strip()
    except Exception:
        try:
            target = str(conjunto_guid or u"").strip()
        except Exception:
            return None
    return target or None


def collect_rebars_por_conjunto_guid(doc, conjunto_guid):
    target = _normalize_conjunto_guid_target(conjunto_guid)
    if doc is None or not target:
        return []

    ids = []
    try:
        rebars = (
            FilteredElementCollector(doc)
            .OfClass(Rebar)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return []

    for rebar in rebars:
        try:
            gid = get_armadura_conjunto_guid(rebar)
        except Exception:
            continue
        if gid == target:
            try:
                ids.append(rebar.Id)
            except Exception:
                pass
    return ids


def collect_empalmes_por_conjunto_guid(doc, conjunto_guid):
    target = _normalize_conjunto_guid_target(conjunto_guid)
    if doc is None or not target:
        return []

    ids = []
    try:
        details = (
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_DetailComponents)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return []

    for el in details:
        if not isinstance(el, FamilyInstance):
            continue
        try:
            gid = get_armadura_conjunto_guid(el)
        except Exception:
            continue
        if gid == target:
            try:
                ids.append(el.Id)
            except Exception:
                pass
    return ids


def _collect_elements_por_conjunto_guid(doc, conjunto_guid, element_class):
    target = _normalize_conjunto_guid_target(conjunto_guid)
    if doc is None or not target or element_class is None:
        return []

    ids = []
    try:
        elements = (
            FilteredElementCollector(doc)
            .OfClass(element_class)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return []

    for el in elements:
        try:
            gid = get_armadura_conjunto_guid(el)
        except Exception:
            continue
        if gid == target:
            try:
                ids.append(el.Id)
            except Exception:
                pass
    return ids


def collect_lienzo_por_conjunto_guid(doc, conjunto_guid):
    """
    Croquis de despiece en vista (curvas de detalle, textos, lienzo masking)
    con el mismo GUID de corrida.
    """
    target = _normalize_conjunto_guid_target(conjunto_guid)
    if doc is None or not target:
        return []

    ids = []
    seen = set()
    for cls in (DetailCurve, TextNote, FilledRegion):
        for eid in _collect_elements_por_conjunto_guid(doc, target, cls):
            try:
                key = int(eid.IntegerValue)
            except Exception:
                key = None
            if key is not None and key not in seen:
                ids.append(eid)
                seen.add(key)
    return ids


def collect_corrida_por_conjunto_guid(doc, conjunto_guid):
    """
    Barras + empalmes + lienzo de despiece con el mismo GUID de corrida.

    Retorna ``dict`` con claves ``rebar_ids``, ``empalme_ids``, ``lienzo_ids``,
    ``all_ids``.
    """
    rebar_ids = collect_rebars_por_conjunto_guid(doc, conjunto_guid)
    empalme_ids = collect_empalmes_por_conjunto_guid(doc, conjunto_guid)
    lienzo_ids = collect_lienzo_por_conjunto_guid(doc, conjunto_guid)
    all_ids = list(rebar_ids)
    seen = set()
    for eid in rebar_ids:
        try:
            seen.add(int(eid.IntegerValue))
        except Exception:
            pass
    for eid in empalme_ids + lienzo_ids:
        try:
            key = int(eid.IntegerValue)
        except Exception:
            key = None
        if key is not None and key not in seen:
            all_ids.append(eid)
            seen.add(key)
    return {
        u"rebar_ids": rebar_ids,
        u"empalme_ids": empalme_ids,
        u"lienzo_ids": lienzo_ids,
        u"all_ids": all_ids,
    }
