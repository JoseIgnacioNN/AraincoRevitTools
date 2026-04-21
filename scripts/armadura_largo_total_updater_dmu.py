# -*- coding: utf-8 -*-
"""
DMU: al modificar un Structural Rebar (p. ej. tramos de forma A, B, C…), actualiza el
parámetro compartido ``Armadura_Largo Total`` con la suma de segmentos (misma lógica que
``_apply_armadura_largo_total_to_rebars`` en enfierrado_shaft_hashtag).

El trabajo se difiere con ExternalEvent (no dentro de IUpdater.Execute).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    ChangePriority,
    Element,
    ElementClassFilter,
    ElementId,
    IUpdater,
    Transaction,
    TransactionStatus,
    UpdaterId,
    UpdaterRegistry,
)
from Autodesk.Revit.DB.Structure import Rebar
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System import Guid

from enfierrado_shaft_hashtag import _apply_armadura_largo_total_to_rebars

UPDATER_GUID = Guid("7f2c9e1a-4d8b-4f6e-9c0a-1b3d5e7f9a2b")

_TXN = u"BIMTools: Armadura_Largo Total (DMU)"

_pending_rebar_by_doc = {}
_largo_total_event = None


def _ensure_largo_total_event():
    global _largo_total_event
    if _largo_total_event is None:
        _largo_total_event = ExternalEvent.Create(_ArmaduraLargoTotalExternalHandler())
    return _largo_total_event


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
        _ensure_largo_total_event().Raise()
    except Exception:
        pass


def _drain_pending():
    global _pending_rebar_by_doc
    out = list(_pending_rebar_by_doc.values())
    _pending_rebar_by_doc = {}
    return out


class _ArmaduraLargoTotalExternalHandler(IExternalEventHandler):
    def GetName(self):
        return u"BIMTools — DMU Armadura_Largo Total (diferido)"

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
            ids = [ElementId(i) for i in sorted(rebar_ints)]
            txn = Transaction(doc, _TXN)
            try:
                if txn.Start() != TransactionStatus.Started:
                    continue
            except Exception:
                continue
            try:
                _apply_armadura_largo_total_to_rebars(doc, ids, [])
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


def _is_rebar_category(el, bic):
    try:
        if el is None or el.Category is None:
            return False
        return int(el.Category.Id.IntegerValue) == int(bic.OST_Rebar)
    except Exception:
        return False


class ArmaduraLargoTotalUpdater(IUpdater):
    def __init__(self, addin_id):
        self._Element = Element
        self._updater_id = UpdaterId(addin_id, UPDATER_GUID)

    def GetUpdaterId(self):
        return self._updater_id

    def GetUpdaterName(self):
        return u"BIMTools — Actualizar Armadura_Largo Total al cambiar Rebar"

    def GetAdditionalInformation(self):
        return (
            u"Sincroniza el parámetro compartido Armadura_Largo Total con la suma de "
            u"tramos de forma (A+B+C…), en mm."
        )

    def GetChangePriority(self):
        return ChangePriority.Rebar

    def _rebar_change_type(self):
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
        _enqueue_rebar_ids(doc, rebar_ids)


def register_armadura_largo_total_updater(addin_id, doc=None):
    _ensure_largo_total_event()
    updater = ArmaduraLargoTotalUpdater(addin_id)
    uid = updater.GetUpdaterId()
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        try:
            UpdaterRegistry.UnregisterUpdater(uid)
        except Exception:
            pass
    UpdaterRegistry.RegisterUpdater(updater)
    flt = ElementClassFilter(Rebar)
    change_type = updater._rebar_change_type()
    if doc is None:
        UpdaterRegistry.AddTrigger(uid, flt, change_type)
    else:
        UpdaterRegistry.AddTrigger(uid, doc, flt, change_type)


def unregister_armadura_largo_total_updater(addin_id):
    uid = UpdaterId(addin_id, UPDATER_GUID)
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        UpdaterRegistry.UnregisterUpdater(uid)


def _addin_id_pyrevit_or_none():
    try:
        from pyrevit import HOST_APP

        return HOST_APP.addin_id
    except Exception:
        return None


def is_armadura_largo_total_dmu_registered(addin_id=None):
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
        if addin_id is None:
            return None
    uid = UpdaterId(addin_id, UPDATER_GUID)
    return UpdaterRegistry.IsUpdaterRegistered(uid)
