# -*- coding: utf-8 -*-
"""
DMU: si un Structural Rebar supera 12 m (12000 mm), colorea de rojo la barra
en la vista activa (OverrideGraphicSettings).

Se excluyen vistas 3D y plantillas de vista. El trabajo se difiere con
ExternalEvent (no dentro de IUpdater.Execute).

Revit 2024–2026 / IronPython (pyRevit).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    ChangePriority,
    ChangeType,
    Color,
    Element,
    ElementClassFilter,
    ElementId,
    FillPatternElement,
    FillPatternTarget,
    FilteredElementCollector,
    IUpdater,
    OverrideGraphicSettings,
    Transaction,
    TransactionStatus,
    UpdaterId,
    UpdaterRegistry,
    ViewType,
)
from Autodesk.Revit.DB.Structure import Rebar
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System import Guid

from armado_muros_rebar_params import rebar_total_length_mm

UPDATER_GUID = Guid("a8c3f1e2-5b7d-4a9e-8c2f-6d1e0b4a9f3c")

MAX_BARRA_COMERCIAL_MM = 12000.0
_COLOR_ROJO = Color(255, 0, 0)
_TXN = u"Arainco: Color barras largo >12 m (DMU)"

_pending_rebar_by_doc = {}
_color_event = None
# Elementos a los que este DMU aplicó override rojo: (doc_id, view_id, rebar_id).
_colored_keys = set()


def _ensure_color_event():
    global _color_event
    if _color_event is None:
        _color_event = ExternalEvent.Create(_RebarLargoExcesoColorHandler())
    return _color_event


def _enqueue_rebar_ids(doc, rebar_element_ids):
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
        _ensure_color_event().Raise()
    except Exception:
        pass


def _drain_pending():
    global _pending_rebar_by_doc
    out = list(_pending_rebar_by_doc.values())
    _pending_rebar_by_doc = {}
    return out


def _is_rebar_category(el, bic):
    try:
        if el is None or el.Category is None:
            return False
        return int(el.Category.Id.IntegerValue) == int(bic.OST_Rebar)
    except Exception:
        return False


def _vista_aplicable(view):
    """Vista activa válida para overrides; excluye 3D y plantillas."""
    if view is None:
        return False
    try:
        if getattr(view, "IsTemplate", False):
            return False
    except Exception:
        pass
    try:
        if view.ViewType == ViewType.ThreeD:
            return False
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import View3D

        if isinstance(view, View3D):
            return False
    except Exception:
        pass
    return True


def _solid_fill_pattern_id(doc):
    invalid = ElementId.InvalidElementId
    if doc is None:
        return invalid
    try:
        fp = FillPatternElement.GetFillPatternElementByName(
            doc, FillPatternTarget.Drafting, u"<Solid fill>",
        )
        if fp is not None:
            return fp.Id
    except Exception:
        pass
    try:
        for el in FilteredElementCollector(doc).OfClass(FillPatternElement):
            if el is None:
                continue
            try:
                pat = el.GetFillPattern()
                if pat is not None and pat.IsSolidFill:
                    return el.Id
            except Exception:
                continue
    except Exception:
        pass
    return invalid


def _aplicar_override_rojo(view, element_id, solid_fill_id=None):
    if view is None or element_id is None:
        return False
    try:
        ogs = OverrideGraphicSettings()
        ogs.SetProjectionLineColor(_COLOR_ROJO)
        try:
            ogs.SetCutLineColor(_COLOR_ROJO)
        except Exception:
            pass
        try:
            ogs.SetSurfaceForegroundPatternColor(_COLOR_ROJO)
        except Exception:
            pass
        try:
            ogs.SetSurfaceBackgroundPatternColor(_COLOR_ROJO)
        except Exception:
            pass
        if solid_fill_id is not None:
            try:
                if solid_fill_id != ElementId.InvalidElementId:
                    ogs.SetSurfaceForegroundPatternId(solid_fill_id)
                    try:
                        ogs.SetSurfaceBackgroundPatternId(solid_fill_id)
                    except Exception:
                        pass
            except Exception:
                pass
        try:
            ogs.SetProjectionLineWeight(6)
        except Exception:
            pass
        view.SetElementOverrides(element_id, ogs)
        return True
    except Exception:
        return False


def _limpiar_override(view, element_id):
    if view is None or element_id is None:
        return False
    try:
        view.SetElementOverrides(element_id, OverrideGraphicSettings())
        return True
    except Exception:
        return False


def _color_key(doc, view, rebar_id):
    try:
        return (id(doc), int(view.Id.IntegerValue), int(rebar_id.IntegerValue))
    except Exception:
        return None


def _aplicar_colores_en_vista(doc, view, rebar_ints):
    """
    Colorea de rojo barras >12 m; quita el override solo si este DMU lo había puesto.
    Devuelve (n_rojo, n_limpiados, n_revisados).
    """
    if doc is None or view is None or not rebar_ints:
        return 0, 0, 0
    if not _vista_aplicable(view):
        return 0, 0, 0

    solid_id = _solid_fill_pattern_id(doc)
    lim = float(MAX_BARRA_COMERCIAL_MM)
    n_rojo = 0
    n_limpio = 0
    n_rev = 0

    for iv in sorted(rebar_ints):
        try:
            eid = ElementId(int(iv))
            el = doc.GetElement(eid)
        except Exception:
            continue
        if el is None or not isinstance(el, Rebar):
            continue
        n_rev += 1
        try:
            L_mm = rebar_total_length_mm(el)
        except Exception:
            L_mm = None

        ck = _color_key(doc, view, eid)
        exceso = L_mm is not None and float(L_mm) > lim + 1e-6

        if exceso:
            if _aplicar_override_rojo(view, eid, solid_id):
                n_rojo += 1
                if ck is not None:
                    _colored_keys.add(ck)
        elif ck is not None and ck in _colored_keys:
            if _limpiar_override(view, eid):
                n_limpio += 1
                _colored_keys.discard(ck)

    return n_rojo, n_limpio, n_rev


class _RebarLargoExcesoColorHandler(IExternalEventHandler):
    def GetName(self):
        return u"Arainco: DMU color barras >12 m (diferido)"

    def Execute(self, uiapp):
        pending = _drain_pending()
        if not pending:
            return

        uidoc = None
        view = None
        active_doc = None
        try:
            uidoc = uiapp.ActiveUIDocument
            if uidoc is not None:
                active_doc = uidoc.Document
                view = uidoc.ActiveView
        except Exception:
            uidoc = None
            view = None
            active_doc = None

        if active_doc is None or not _vista_aplicable(view):
            return

        for doc, rebar_ints in pending:
            try:
                if doc is None or not doc.IsValidObject or doc.IsLinked:
                    continue
            except Exception:
                continue
            if not rebar_ints:
                continue
            # Solo la vista activa del documento activo (no otros docs abiertos).
            try:
                if doc.Equals(active_doc) is False:
                    continue
            except Exception:
                try:
                    if doc != active_doc:
                        continue
                except Exception:
                    continue

            txn = Transaction(doc, _TXN)
            try:
                if txn.Start() != TransactionStatus.Started:
                    continue
            except Exception:
                continue
            try:
                _aplicar_colores_en_vista(doc, view, rebar_ints)
            finally:
                try:
                    if txn.GetStatus() == TransactionStatus.Started:
                        txn.Commit()
                except Exception:
                    try:
                        if txn.GetStatus() == TransactionStatus.Started:
                            txn.RollBack()
                    except Exception:
                        pass


def _rebar_modification_change_type():
    """Geometría + Any: cambios de largo suelen llegar como geometría."""
    try:
        return ChangeType.Concatenate(
            Element.GetChangeTypeGeometry(),
            Element.GetChangeTypeAny(),
        )
    except Exception:
        try:
            return Element.GetChangeTypeGeometry()
        except Exception:
            return Element.GetChangeTypeAny()


class RebarLargoExcesoColorUpdater(IUpdater):
    def __init__(self, addin_id):
        self._Element = Element
        self._updater_id = UpdaterId(addin_id, UPDATER_GUID)

    def GetUpdaterId(self):
        return self._updater_id

    def GetUpdaterName(self):
        return u"Arainco: Color rojo barras largo >12 m"

    def GetAdditionalInformation(self):
        return (
            u"Colorea de rojo en la vista activa (no 3D) los Rebar cuyo largo "
            u"total supera 12 m (12000 mm)."
        )

    def GetChangePriority(self):
        return ChangePriority.Rebar

    def Execute(self, data):
        doc = data.GetDocument()
        if doc is None or doc.IsLinked:
            return
        from Autodesk.Revit.DB import BuiltInCategory

        bic = BuiltInCategory
        rebar_ids = []
        seen = set()

        for getter in (data.GetModifiedElementIds, data.GetAddedElementIds):
            try:
                eids = list(getter())
            except Exception:
                eids = []
            for eid in eids:
                try:
                    iv = int(eid.IntegerValue)
                except Exception:
                    continue
                if iv in seen:
                    continue
                try:
                    el = doc.GetElement(eid)
                except Exception:
                    el = None
                if not _is_rebar_category(el, bic):
                    continue
                seen.add(iv)
                rebar_ids.append(eid)

        if not rebar_ids:
            return
        _enqueue_rebar_ids(doc, rebar_ids)


def register_rebar_largo_exceso_color_updater(addin_id, doc=None):
    _ensure_color_event()
    updater = RebarLargoExcesoColorUpdater(addin_id)
    uid = updater.GetUpdaterId()
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        try:
            UpdaterRegistry.UnregisterUpdater(uid)
        except Exception:
            pass
    UpdaterRegistry.RegisterUpdater(updater)
    flt = ElementClassFilter(Rebar)

    try:
        ct_mod = _rebar_modification_change_type()
        if doc is None:
            UpdaterRegistry.AddTrigger(uid, flt, ct_mod)
        else:
            UpdaterRegistry.AddTrigger(uid, doc, flt, ct_mod)
    except Exception:
        pass

    try:
        ct_add = Element.GetChangeTypeElementAddition()
        if doc is None:
            UpdaterRegistry.AddTrigger(uid, flt, ct_add)
        else:
            UpdaterRegistry.AddTrigger(uid, doc, flt, ct_add)
    except Exception:
        pass


def unregister_rebar_largo_exceso_color_updater(addin_id):
    uid = UpdaterId(addin_id, UPDATER_GUID)
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        UpdaterRegistry.UnregisterUpdater(uid)


def _addin_id_pyrevit_or_none():
    try:
        from pyrevit import HOST_APP

        return HOST_APP.addin_id
    except Exception:
        return None


def is_rebar_largo_exceso_color_dmu_registered(addin_id=None):
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
        if addin_id is None:
            return None
    uid = UpdaterId(addin_id, UPDATER_GUID)
    return UpdaterRegistry.IsUpdaterRegistered(uid)
