# -*- coding: utf-8 -*-
"""
DMU: al cambiar Rebar, las IndependentTag asociadas pasan al tipo de familia
configurada cuyo nombre coincide con el RebarShape.

La lógica de mapa / transacciones vive en rebar_tag_shape_sync_core.py.
Familia(s): REBAR_TAG_SYNC_DEFAULT_FAMILY_NAMES en ese módulo (o tupla importada abajo).

El trabajo se difiere con ExternalEvent (no dentro de IUpdater.Execute).
"""

from __future__ import print_function

from collections import defaultdict

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    ChangePriority,
    Element,
    ElementClassFilter,
    ElementId,
    FilteredElementCollector,
    FamilySymbol,
    IUpdater,
    UpdaterId,
    UpdaterRegistry,
)
from Autodesk.Revit.DB.Structure import Rebar
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System import Guid

from rebar_tag_shape_sync_core import (
    REBAR_TAG_SYNC_DEFAULT_FAMILY_NAMES,
    apply_tag_sync_for_rebar_ints,
    comparison_keys,
    family_names_match,
    normalize_label,
    symbol_map_from_family_names,
    to_python_str,
)

UPDATER_GUID = Guid("e4b2a91c-6f03-4b5d-9e81-2c7d4a0b1f36")

# Primer nombre de la tupla por compatibilidad con mensajes y diagnósticos antiguos.
REBAR_TAG_FAMILY_NAME = (
    REBAR_TAG_SYNC_DEFAULT_FAMILY_NAMES[0]
    if REBAR_TAG_SYNC_DEFAULT_FAMILY_NAMES
    else u""
)

_TXN_REFRESH = u"BIMTools: tipo etiqueta según RebarShape (DMU)"

_pending_rebar_by_doc = {}
_tag_refresh_event = None


def _ensure_tag_refresh_event():
    global _tag_refresh_event
    if _tag_refresh_event is None:
        _tag_refresh_event = ExternalEvent.Create(_TagRefreshExternalHandler())
    return _tag_refresh_event


def _enqueue_rebar_refresh(doc, rebar_element_ids):
    global _pending_rebar_by_doc
    key = id(doc)
    new_ints = set()
    for eid in rebar_element_ids:
        try:
            new_ints.add(int(eid.IntegerValue))
        except Exception:
            continue
    if not new_ints:
        return
    if key in _pending_rebar_by_doc:
        _doc_ref, existing = _pending_rebar_by_doc[key]
        existing |= new_ints
    else:
        _pending_rebar_by_doc[key] = (doc, new_ints)
    try:
        _ensure_tag_refresh_event().Raise()
    except Exception:
        pass


def _drain_pending():
    global _pending_rebar_by_doc
    out = list(_pending_rebar_by_doc.values())
    _pending_rebar_by_doc = {}
    return out


class _TagRefreshExternalHandler(IExternalEventHandler):
    def GetName(self):
        return u"BIMTools — DMU tipo etiqueta = RebarShape (diferido)"

    def Execute(self, uiapp):
        pending = _drain_pending()
        for doc, rebar_ints in pending:
            try:
                if doc is None or not doc.IsValidObject or doc.IsLinked:
                    continue
            except Exception:
                continue
            if not rebar_ints:
                continue
            sm = symbol_map_from_family_names(doc, REBAR_TAG_SYNC_DEFAULT_FAMILY_NAMES)
            if not sm:
                continue
            apply_tag_sync_for_rebar_ints(doc, rebar_ints, sm, _TXN_REFRESH)


def _is_rebar_category(el, bic):
    try:
        from Autodesk.Revit.DB import BuiltInCategory

        if el is None or el.Category is None:
            return False
        return int(el.Category.Id.IntegerValue) == int(bic.OST_Rebar)
    except Exception:
        return False


class RebarShapeTagRefresherUpdater(IUpdater):
    def __init__(self, addin_id):
        self._Element = Element
        self._ElementId = ElementId
        self._updater_id = UpdaterId(addin_id, UPDATER_GUID)

    def GetUpdaterId(self):
        return self._updater_id

    def GetUpdaterName(self):
        return u"BIMTools — Refrescar etiquetas al cambiar RebarShape"

    def GetAdditionalInformation(self):
        return (
            u"Encola cambio de tipo de IndependentTag al nombre del RebarShape "
            u"(familias: %s)." % u", ".join(REBAR_TAG_SYNC_DEFAULT_FAMILY_NAMES)
        )

    def GetChangePriority(self):
        return ChangePriority.Rebar

    def _shape_change_type(self):
        return self._Element.GetChangeTypeAny()

    def Execute(self, data):
        doc = data.GetDocument()
        if doc is None or doc.IsLinked:
            return
        modified = list(data.GetModifiedElementIds())
        if not modified:
            return
        from Autodesk.Revit.DB import BuiltInCategory

        bic = BuiltInCategory
        rebar_ids = []
        for eid in modified:
            try:
                el = doc.GetElement(eid)
            except Exception:
                continue
            if not _is_rebar_category(el, bic):
                continue
            rebar_ids.append(eid)
        if not rebar_ids:
            return
        _enqueue_rebar_refresh(doc, rebar_ids)


def register_rebar_shape_tag_updater(addin_id, doc=None):
    _ensure_tag_refresh_event()
    updater = RebarShapeTagRefresherUpdater(addin_id)
    uid = updater.GetUpdaterId()
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        try:
            UpdaterRegistry.UnregisterUpdater(uid)
        except Exception:
            pass
    UpdaterRegistry.RegisterUpdater(updater)
    flt = ElementClassFilter(Rebar)
    change_type = updater._shape_change_type()
    if doc is None:
        UpdaterRegistry.AddTrigger(uid, flt, change_type)
    else:
        UpdaterRegistry.AddTrigger(uid, doc, flt, change_type)


def unregister_rebar_shape_tag_updater(addin_id):
    uid = UpdaterId(addin_id, UPDATER_GUID)
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        UpdaterRegistry.UnregisterUpdater(uid)


def _addin_id_pyrevit_or_none():
    try:
        from pyrevit import HOST_APP

        return HOST_APP.addin_id
    except Exception:
        return None


def is_rebar_shape_tag_dmu_registered(addin_id=None):
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
        if addin_id is None:
            return None
    uid = UpdaterId(addin_id, UPDATER_GUID)
    return UpdaterRegistry.IsUpdaterRegistered(uid)


def print_rebar_shape_tag_dmu_status(addin_id=None):
    print(u"--- Estado DMU etiquetas / RebarShape ---")
    print(u"GUID updater: {}".format(UPDATER_GUID))
    print(u"Nombre: BIMTools — Refrescar etiquetas al cambiar RebarShape")
    print(
        u"Familias (core): {}".format(
            u", ".join(REBAR_TAG_SYNC_DEFAULT_FAMILY_NAMES)
        )
    )
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
    if addin_id is None:
        print(
            u"AddInId: (no disponible — ejecuta esto desde un script **pyRevit**, "
            u"no desde RPS, o pasa addin_id explícito)."
        )
        return
    print(u"AddInId (pyRevit): {!r}".format(addin_id))
    uid = UpdaterId(addin_id, UPDATER_GUID)
    reg = UpdaterRegistry.IsUpdaterRegistered(uid)
    print(u"Registrado en UpdaterRegistry: {}".format(reg))
    global _tag_refresh_event
    print(u"ExternalEvent creado: {}".format(_tag_refresh_event is not None))


def _sym_family_in_default_list(fam_name):
    for n in REBAR_TAG_SYNC_DEFAULT_FAMILY_NAMES:
        if family_names_match(fam_name, n):
            return True
    return False


def diagnostic_rebar_shape_vs_tag_types(doc):
    from Autodesk.Revit.DB.Structure import RebarShape

    fec = FilteredElementCollector
    shape_labels = set()
    for sh in fec(doc).OfClass(RebarShape):
        try:
            shape_labels.add(normalize_label(to_python_str(sh.Name)))
        except Exception:
            continue

    tag_labels = set()
    for sym in fec(doc).OfClass(FamilySymbol):
        try:
            fam = sym.Family
            if fam is None or not _sym_family_in_default_list(fam.Name):
                continue
            tag_labels.add(normalize_label(to_python_str(sym.Name)))
        except Exception:
            continue

    def _expand_keys(label_set):
        keys = set()
        for lb in label_set:
            for k in comparison_keys(lb):
                keys.add(k)
        return keys

    shape_keys = _expand_keys(shape_labels)
    tag_keys = _expand_keys(tag_labels)

    shapes_resolved = set()
    for lb in shape_labels:
        for k in comparison_keys(lb):
            if k in tag_keys:
                shapes_resolved.add(lb)
                break
    tags_resolved = set()
    for lb in tag_labels:
        for k in comparison_keys(lb):
            if k in shape_keys:
                tags_resolved.add(lb)
                break

    exact_match = shape_labels & tag_labels
    only_sh = sorted(shape_labels - shapes_resolved)
    only_tg = sorted(tag_labels - tags_resolved)

    return {
        u"rebar_shape_names": sorted(shape_labels),
        u"tag_type_names": sorted(tag_labels),
        u"exact_normalized_name_match": sorted(exact_match),
        u"shapes_linked_to_a_tag_type_by_keys": sorted(shapes_resolved),
        u"tag_types_linked_to_a_shape_by_keys": sorted(tags_resolved),
        u"shapes_without_matching_tag_type": only_sh,
        u"tag_types_without_matching_shape": only_tg,
    }


def print_diagnostic_rebar_shape_vs_tag_types(doc):
    r = diagnostic_rebar_shape_vs_tag_types(doc)
    order = (
        u"rebar_shape_names",
        u"tag_type_names",
        u"exact_normalized_name_match",
        u"shapes_linked_to_a_tag_type_by_keys",
        u"tag_types_linked_to_a_shape_by_keys",
        u"shapes_without_matching_tag_type",
        u"tag_types_without_matching_shape",
    )
    for k in order:
        print(u"--- {} ---".format(k))
        for line in r.get(k, []):
            print(u"  {}".format(line))
        print(u"  (total: {})".format(len(r.get(k, []))))

    miss_sh = r.get(u"shapes_without_matching_tag_type", [])
    miss_tg = r.get(u"tag_types_without_matching_shape", [])
    print(u"--- resumen ---")
    if not miss_sh and not miss_tg:
        print(u"  Todo enlazado (convención de claves).")
    if miss_sh:
        print(
            u"  Shapes sin tipo en familias {}: {}.".format(
                REBAR_TAG_SYNC_DEFAULT_FAMILY_NAMES,
                u", ".join(miss_sh),
            )
        )
    if miss_tg:
        print(
            u"  Tipos de etiqueta sin RebarShape homónimo: {}.".format(
                u", ".join(miss_tg),
            )
        )


def list_rebar_tag_families_and_types(doc):
    from Autodesk.Revit.DB import BuiltInCategory

    bic = BuiltInCategory
    fec = FilteredElementCollector
    by_fam = defaultdict(list)
    for sym in fec(doc).OfClass(FamilySymbol):
        try:
            if sym.Category is None:
                continue
            if int(sym.Category.Id.IntegerValue) != int(bic.OST_RebarTags):
                continue
            raw_fam = to_python_str(sym.Family.Name)
            raw_typ = to_python_str(sym.Name)
            by_fam[raw_fam].append(raw_typ)
        except Exception:
            continue
    return dict((k, sorted(set(v))) for k, v in by_fam.items())


def print_rebar_tag_families_inventory(doc):
    print(u"--- Inventario OST_RebarTags (todas las familias del proyecto) ---")
    inv = list_rebar_tag_families_and_types(doc)
    if not inv:
        print(u"  (ningún FamilySymbol en OST_RebarTags)")
        return
    for raw_fam in sorted(inv.keys(), key=lambda x: normalize_label(x).lower()):
        types = inv[raw_fam]
        is_dmu = _sym_family_in_default_list(raw_fam)
        tag = u"  <<< DMU / core (REBAR_TAG_SYNC_DEFAULT_FAMILY_NAMES)" if is_dmu else u""
        print(u"Familia {!r}  ({} tipos){}".format(raw_fam, len(types), tag))
        for raw_typ in types:
            print(u"  tipo {!r}".format(raw_typ))
