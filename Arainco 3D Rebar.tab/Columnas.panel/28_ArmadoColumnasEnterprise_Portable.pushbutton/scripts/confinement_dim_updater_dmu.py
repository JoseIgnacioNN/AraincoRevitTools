# -*- coding: utf-8 -*-
"""
DMU: al borrar una cota de confinamiento (columnas), elimina los DetailCurve
marcadores ligados vía ``confinement_dim_link_schema``.

Solo reacciona a **borrado** de ``Dimension`` (no a cambio de tipo ni geometría).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    ChangePriority,
    Dimension,
    Element,
    ElementClassFilter,
    ElementId,
    IUpdater,
    Transaction,
    TransactionStatus,
    UpdaterId,
    UpdaterRegistry,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System import Guid

from confinement_dim_link_schema import (
    element_id_to_int,
    find_confinement_dim_markers_for_dim_ids,
)

UPDATER_GUID = Guid("e5f6a7b8-c9d0-4123-e456-789abcdef012")

_TXN = u"Arainco: borrar marcadores cota confinamiento (DMU)"

_pending_dim_ids_by_doc = {}
_confinement_dim_event = None


def _ensure_event():
    global _confinement_dim_event
    if _confinement_dim_event is None:
        _confinement_dim_event = ExternalEvent.Create(_ConfinementDimExternalHandler())
    return _confinement_dim_event


def _enqueue(doc, dim_id_ints):
    global _pending_dim_ids_by_doc
    if doc is None or not dim_id_ints:
        return
    key = id(doc)
    new_ints = set(int(x) for x in dim_id_ints)
    if not new_ints:
        return
    if key in _pending_dim_ids_by_doc:
        _doc_ref, existing = _pending_dim_ids_by_doc[key]
        existing |= new_ints
    else:
        _pending_dim_ids_by_doc[key] = (doc, new_ints)
    try:
        _ensure_event().Raise()
    except Exception:
        pass


def _drain_pending():
    global _pending_dim_ids_by_doc
    out = list(_pending_dim_ids_by_doc.values())
    _pending_dim_ids_by_doc = {}
    return out


class _ConfinementDimExternalHandler(IExternalEventHandler):
    def GetName(self):
        return u"Arainco: DMU marcadores cota confinamiento (diferido)"

    def Execute(self, uiapp):
        pending = _drain_pending()
        for doc, dim_ints in pending:
            try:
                if doc is None or not doc.IsValidObject or doc.IsLinked:
                    continue
            except Exception:
                continue
            if not dim_ints:
                continue
            markers = find_confinement_dim_markers_for_dim_ids(doc, dim_ints)
            if not markers:
                continue
            txn = Transaction(doc, _TXN)
            try:
                if txn.Start() != TransactionStatus.Started:
                    continue
            except Exception:
                continue
            try:
                for dc in markers:
                    try:
                        if dc is not None and dc.IsValidObject:
                            doc.Delete(dc.Id)
                    except Exception:
                        pass
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


class ConfinementDimLinkUpdater(IUpdater):
    def __init__(self, addin_id):
        self._Element = Element
        self._updater_id = UpdaterId(addin_id, UPDATER_GUID)

    def GetUpdaterId(self):
        return self._updater_id

    def GetUpdaterName(self):
        return u"Arainco: Borrar marcadores al eliminar cota confinamiento"

    def GetAdditionalInformation(self):
        return (
            u"Elimina las líneas de detalle marcadoras ligadas a cotas "
            u"«Linear - Confinamiento» de columnas cuando se borra la cota."
        )

    def GetChangePriority(self):
        return ChangePriority.Annotations

    def Execute(self, data):
        doc = data.GetDocument()
        if doc is None or doc.IsLinked:
            return
        touch = set()
        try:
            for eid in data.GetDeletedElementIds():
                ei = element_id_to_int(eid)
                if ei is not None:
                    touch.add(ei)
        except Exception:
            pass
        if not touch:
            return
        _enqueue(doc, touch)


def register_confinement_dim_link_updater(addin_id, doc=None):
    _ensure_event()
    updater = ConfinementDimLinkUpdater(addin_id)
    uid = updater.GetUpdaterId()
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        try:
            UpdaterRegistry.UnregisterUpdater(uid)
        except Exception:
            pass
    UpdaterRegistry.RegisterUpdater(updater)
    flt = ElementClassFilter(Dimension)
    ct_del = Element.GetChangeTypeElementDeletion()
    if doc is None:
        UpdaterRegistry.AddTrigger(uid, flt, ct_del)
    else:
        UpdaterRegistry.AddTrigger(uid, doc, flt, ct_del)


def unregister_confinement_dim_link_updater(addin_id):
    uid = UpdaterId(addin_id, UPDATER_GUID)
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        UpdaterRegistry.UnregisterUpdater(uid)


def _addin_id_pyrevit_or_none():
    try:
        from pyrevit import HOST_APP

        return HOST_APP.addin_id
    except Exception:
        return None


def is_confinement_dim_link_updater_registered(addin_id=None):
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
        if addin_id is None:
            return None
    uid = UpdaterId(addin_id, UPDATER_GUID)
    return UpdaterRegistry.IsUpdaterRegistered(uid)


def ensure_confinement_dim_link_updater_registered():
    """
    Registro idempotente (startup de extensión o primera cota creada en sesión).
    """
    addin_id = _addin_id_pyrevit_or_none()
    if addin_id is None:
        return False
    uid = UpdaterId(addin_id, UPDATER_GUID)
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        return True
    try:
        register_confinement_dim_link_updater(addin_id, doc=None)
        return True
    except Exception:
        return False
