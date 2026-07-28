# -*- coding: utf-8 -*-
"""
Arainco: seleccionar capa de armadura por GUID + Armadura_Capa.

1. Filtrar Rebar / Detail Items con el mismo Armadura_Conjunto_GUID y
   Armadura_Capa que el elemento de referencia.
2. Incluir IndependentTag (OST_RebarTags) de esas barras en la vista activa.
3. Seleccionar el conjunto resultante en el modelo.

Revit 2024–2026 · IronPython (pyRevit).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    IndependentTag,
)
from Autodesk.Revit.DB.Structure import Rebar
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

ARMADURA_CONJUNTO_GUID_PARAM = u"Armadura_Conjunto_GUID"
ARMADURA_CAPA_PARAM = u"Armadura_Capa"
_DIALOG_BASE = u"Arainco: Seleccionar capa"
_PICK_PROMPT = (
    u"Selecciona una barra (Rebar) o Detail Item con «Armadura_Conjunto_GUID» "
    u"y «Armadura_Capa» para seleccionar el resto de la misma capa."
)


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _guid_snippet(guid, max_len=72):
    if not guid:
        return u""
    s = _as_unicode(guid).strip()
    if len(s) <= max_len:
        return s
    return s[: max_len - 1] + u"…"


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
        for p in element.Parameters:
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


def _param_as_text(element, param_name):
    p = _find_element_parameter(element, param_name)
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
    t = _as_unicode(val).strip()
    return t or None


def get_armadura_conjunto_guid(element):
    return _param_as_text(element, ARMADURA_CONJUNTO_GUID_PARAM)


def get_armadura_capa(element):
    return _param_as_text(element, ARMADURA_CAPA_PARAM)


def _normalize_target(value):
    t = _as_unicode(value).strip()
    return t or None


def _rebar_ids_int(rebar_ids):
    out = set()
    for eid in rebar_ids or []:
        try:
            out.add(int(eid.IntegerValue))
        except Exception:
            pass
    return out


def _tag_referencia_rebar_ids(tag, rebar_ids_int):
    if tag is None or not rebar_ids_int:
        return False
    invalid = ElementId.InvalidElementId
    try:
        for tid in tag.GetTaggedLocalElementIds():
            try:
                if int(tid.IntegerValue) in rebar_ids_int:
                    return True
            except Exception:
                continue
    except Exception:
        pass
    try:
        for leid in tag.GetTaggedElementIds():
            try:
                link_inst = leid.LinkInstanceId
                if (
                    link_inst is not None
                    and link_inst != invalid
                    and int(link_inst.IntegerValue) >= 0
                ):
                    continue
            except Exception:
                pass
            for attr in (u"LinkedElementId", u"HostElementId"):
                try:
                    eid = getattr(leid, attr, None)
                    if eid is None:
                        continue
                    if int(eid.IntegerValue) in rebar_ids_int:
                        return True
                except Exception:
                    continue
    except Exception:
        pass
    try:
        tid = tag.TaggedLocalElementId
        if tid is not None and int(tid.IntegerValue) in rebar_ids_int:
            return True
    except Exception:
        pass
    return False


def collect_rebars_por_conjunto_guid_y_capa(doc, conjunto_guid, capa):
    target_guid = _normalize_target(conjunto_guid)
    target_capa = _normalize_target(capa)
    if doc is None or not target_guid or not target_capa:
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
            if get_armadura_conjunto_guid(rebar) != target_guid:
                continue
            if get_armadura_capa(rebar) != target_capa:
                continue
            ids.append(rebar.Id)
        except Exception:
            continue
    return ids


def collect_empalmes_por_conjunto_guid_y_capa(doc, conjunto_guid, capa):
    target_guid = _normalize_target(conjunto_guid)
    target_capa = _normalize_target(capa)
    if doc is None or not target_guid or not target_capa:
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
            if get_armadura_conjunto_guid(el) != target_guid:
                continue
            if get_armadura_capa(el) != target_capa:
                continue
            ids.append(el.Id)
        except Exception:
            continue
    return ids


def collect_etiquetas_rebar_en_vista(doc, rebar_ids, view):
    rebar_ids_int = _rebar_ids_int(rebar_ids)
    if doc is None or not rebar_ids_int or view is None:
        return []
    ids = []
    seen = set()
    try:
        tags = (
            FilteredElementCollector(doc, view.Id)
            .OfClass(IndependentTag)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return []
    rebar_tag_cat = int(BuiltInCategory.OST_RebarTags)
    for tag in tags or []:
        if tag is None or not isinstance(tag, IndependentTag):
            continue
        try:
            cat = tag.Category
            if cat is None or int(cat.Id.IntegerValue) != rebar_tag_cat:
                continue
        except Exception:
            continue
        try:
            if tag.OwnerViewId != view.Id:
                continue
        except Exception:
            pass
        if not _tag_referencia_rebar_ids(tag, rebar_ids_int):
            continue
        try:
            key = int(tag.Id.IntegerValue)
        except Exception:
            key = None
        if key is not None and key in seen:
            continue
        ids.append(tag.Id)
        if key is not None:
            seen.add(key)
    return ids


def collect_capa_por_conjunto_guid(doc, conjunto_guid, capa, view=None):
    rebar_ids = collect_rebars_por_conjunto_guid_y_capa(
        doc, conjunto_guid, capa,
    )
    empalme_ids = collect_empalmes_por_conjunto_guid_y_capa(
        doc, conjunto_guid, capa,
    )
    tag_ids = collect_etiquetas_rebar_en_vista(doc, rebar_ids, view)
    all_ids = list(rebar_ids)
    seen = set()
    for eid in rebar_ids:
        try:
            seen.add(int(eid.IntegerValue))
        except Exception:
            pass
    for eid in empalme_ids + tag_ids:
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
        u"tag_ids": tag_ids,
        u"all_ids": all_ids,
        u"capa": _normalize_target(capa),
    }


def _es_detail_item(elem):
    if not isinstance(elem, FamilyInstance):
        return False
    try:
        cat = elem.Category
        if cat is None:
            return False
        return int(cat.Id.IntegerValue) == int(
            BuiltInCategory.OST_DetailComponents,
        )
    except Exception:
        return False


class _FiltroCapaReferencia(ISelectionFilter):
    def AllowElement(self, elem):
        if isinstance(elem, Rebar):
            return True
        return _es_detail_item(elem)

    def AllowReference(self, reference, position):
        return False


def _element_id_list(ids):
    sel = List[ElementId]()
    for eid in ids or []:
        if eid is None:
            continue
        if isinstance(eid, ElementId):
            sel.Add(eid)
        else:
            try:
                sel.Add(ElementId(int(eid)))
            except Exception:
                pass
    return sel


def pick_capa_referencia(uidoc):
    if uidoc is None:
        TaskDialog.Show(_DIALOG_BASE, u"No hay documento activo.")
        return None
    doc = uidoc.Document
    view = uidoc.ActiveView
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _FiltroCapaReferencia(),
            _PICK_PROMPT,
        )
    except OperationCanceledException:
        return None
    except Exception as ex:
        TaskDialog.Show(_DIALOG_BASE, u"Error al seleccionar:\n{0}".format(ex))
        return None

    if ref is None:
        return None

    elem = doc.GetElement(ref.ElementId)
    guid = get_armadura_conjunto_guid(elem)
    if not guid:
        TaskDialog.Show(
            _DIALOG_BASE,
            u"El elemento elegido no tiene valor en «{0}».\n\n"
            u"Solo aplica a barras o Detail Items creados por herramientas "
            u"Arainco que estampan ese parámetro.".format(
                ARMADURA_CONJUNTO_GUID_PARAM,
            ),
        )
        return None

    capa = get_armadura_capa(elem)
    if not capa:
        TaskDialog.Show(
            _DIALOG_BASE,
            u"El elemento elegido no tiene valor en «{0}».\n\n"
            u"Solo aplica a barras / empalmes por capa "
            u"(p. ej. Armado Muros o Armado Vigas).".format(ARMADURA_CAPA_PARAM),
        )
        return None

    capa_corrida = collect_capa_por_conjunto_guid(
        doc, guid, capa, view=view,
    )
    if not (capa_corrida.get(u"all_ids") or []):
        TaskDialog.Show(
            _DIALOG_BASE,
            u"No se encontraron elementos con GUID «{0}» y capa «{1}».".format(
                _guid_snippet(guid), capa,
            ),
        )
        return None

    return doc, guid, capa, capa_corrida


def select_capa_en_modelo(uidoc, capa_corrida):
    ids = (capa_corrida or {}).get(u"all_ids") or []
    sel = _element_id_list(ids)
    if sel.Count < 1:
        return 0
    try:
        uidoc.Selection.SetElementIds(sel)
        uidoc.ShowElements(sel)
        return int(sel.Count)
    except Exception as ex:
        TaskDialog.Show(_DIALOG_BASE, u"Error al seleccionar:\n{0}".format(ex))
        return 0


def run(uiapp):
    """Entrada pyRevit: selecciona capa por GUID + Armadura_Capa."""
    uidoc = uiapp.ActiveUIDocument if uiapp is not None else None
    picked = pick_capa_referencia(uidoc)
    if not picked:
        return
    _doc, guid, capa, capa_corrida = picked
    n = select_capa_en_modelo(uidoc, capa_corrida)
    if n < 1:
        return
    n_rebar = len(capa_corrida.get(u"rebar_ids") or [])
    n_emp = len(capa_corrida.get(u"empalme_ids") or [])
    n_tag = len(capa_corrida.get(u"tag_ids") or [])
    TaskDialog.Show(
        _DIALOG_BASE,
        u"Capa seleccionada ({0} elemento(s)).\n\n"
        u"Capa: {1}\n"
        u"GUID: {2}\n"
        u"  · Barras: {3}\n"
        u"  · Detail Items: {4}\n"
        u"  · Etiquetas (vista activa): {5}".format(
            n, capa, _guid_snippet(guid), n_rebar, n_emp, n_tag,
        ),
    )
