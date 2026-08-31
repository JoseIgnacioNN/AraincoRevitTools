# -*- coding: utf-8 -*-
"""
DMU: al cambiar el diámetro (RebarBarType) de un Structural Rebar con pata L,
ajusta el largo de la pata según la tabla BIMTools (``bimtools_rebar_hook_lengths``).

El trabajo se difiere con ExternalEvent (no dentro de IUpdater.Execute).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ChangePriority,
    ChangeType,
    Element,
    ElementClassFilter,
    ElementId,
    IUpdater,
    UpdaterId,
    UpdaterRegistry,
)
from Autodesk.Revit.DB.Structure import Rebar
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System import Guid

from rebar_pata_l_diameter_change import (
    get_document_concrete_grade_override,
    process_rebars_for_diameter_pata_l,
    rebar_bar_type_change_detected,
    remember_bar_type_snapshot_for_rebar,
    seed_all_rebar_bar_types_in_document,
    seed_bar_type_cache_if_unknown,
    set_document_concrete_grade_override,
)

UPDATER_GUID = Guid("c1d2e3f4-a5b6-4789-c012-3456789abcde")

_pending_rebar_by_doc = {}
_pata_l_event = None
_document_opened_hooked = False


def _on_document_opened_seed(sender, args):
    try:
        doc = args.Document
    except Exception:
        doc = None
    if doc is None or doc.IsLinked:
        return
    try:
        seed_all_rebar_bar_types_in_document(doc)
    except Exception:
        pass


def _ensure_document_opened_seed():
    global _document_opened_hooked
    if _document_opened_hooked:
        return
    try:
        from pyrevit import HOST_APP

        app = HOST_APP.app
        if app is not None:
            app.DocumentOpened += _on_document_opened_seed
            _document_opened_hooked = True
    except Exception:
        pass


def _ensure_pata_l_event():
    global _pata_l_event
    if _pata_l_event is None:
        _pata_l_event = ExternalEvent.Create(_RebarPataLDiameterExternalHandler())
    return _pata_l_event


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
        _ensure_pata_l_event().Raise()
    except Exception:
        pass


def _drain_pending():
    global _pending_rebar_by_doc
    out = list(_pending_rebar_by_doc.values())
    _pending_rebar_by_doc = {}
    return out


class _RebarPataLDiameterExternalHandler(IExternalEventHandler):
    def GetName(self):
        return u"Arainco: DMU pata L por diámetro (diferido)"

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
            try:
                process_rebars_for_diameter_pata_l(doc, ids)
            except Exception:
                pass


def _is_rebar_category(el, bic):
    try:
        if el is None or el.Category is None:
            return False
        return int(el.Category.Id.IntegerValue) == int(bic.OST_Rebar)
    except Exception:
        return False


class RebarPataLDiameterUpdater(IUpdater):
    def __init__(self, addin_id):
        self._Element = Element
        self._updater_id = UpdaterId(addin_id, UPDATER_GUID)

    def GetUpdaterId(self):
        return self._updater_id

    def GetUpdaterName(self):
        return u"Arainco: Ajustar pata L al cambiar diámetro de Rebar"

    def GetAdditionalInformation(self):
        return (
            u"En barras con RebarShape «02» o «03», al cambiar el diámetro "
            u"actualiza el parámetro A (y C en «03») según la tabla BIMTools "
            u"por dosificación G25/G35/G45."
        )

    def GetChangePriority(self):
        return ChangePriority.Rebar

    def _rebar_change_type(self):
        try:
            return ChangeType.Concatenate(
                self._Element.GetChangeTypeGeometry(),
                self._Element.GetChangeTypeAny(),
                self._Element.GetChangeTypeElementType(),
            )
        except Exception:
            try:
                return ChangeType.Concatenate(
                    self._Element.GetChangeTypeGeometry(),
                    self._Element.GetChangeTypeAny(),
                )
            except Exception:
                return self._Element.GetChangeTypeAny()

    def Execute(self, data):
        doc = data.GetDocument()
        if doc is None or doc.IsLinked:
            return
        modified = list(data.GetModifiedElementIds())
        if not modified:
            return
        bic = BuiltInCategory
        rebar_ids = []
        for eid in modified:
            try:
                el = doc.GetElement(eid)
            except Exception:
                continue
            if not _is_rebar_category(el, bic):
                continue
            if not isinstance(el, Rebar):
                continue
            if rebar_bar_type_change_detected(
                doc, el, updater_data=data, element_id=eid
            ):
                rebar_ids.append(eid)
            else:
                seed_bar_type_cache_if_unknown(doc, el)
        if not rebar_ids:
            return
        _enqueue_rebar_ids(doc, rebar_ids)


def register_rebar_pata_l_diameter_updater(addin_id, doc=None):
    _ensure_pata_l_event()
    _ensure_document_opened_seed()
    updater = RebarPataLDiameterUpdater(addin_id)
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
    try:
        if doc is not None:
            seed_all_rebar_bar_types_in_document(doc)
        else:
            from pyrevit import HOST_APP

            uiapp = getattr(HOST_APP, "uiapp", None)
            if uiapp is not None and uiapp.ActiveUIDocument is not None:
                seed_all_rebar_bar_types_in_document(uiapp.ActiveUIDocument.Document)
    except Exception:
        pass
    return True


def unregister_rebar_pata_l_diameter_updater(addin_id):
    uid = UpdaterId(addin_id, UPDATER_GUID)
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        UpdaterRegistry.UnregisterUpdater(uid)


def _addin_id_pyrevit_or_none():
    try:
        from pyrevit import HOST_APP

        return HOST_APP.addin_id
    except Exception:
        return None


def is_rebar_pata_l_diameter_dmu_registered(addin_id=None):
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
        if addin_id is None:
            return None
    uid = UpdaterId(addin_id, UPDATER_GUID)
    return UpdaterRegistry.IsUpdaterRegistered(uid)


def toggle_rebar_pata_l_diameter_dmu(addin_id=None, doc=None):
    """Alterna registro. Retorna ``(registered_now, message)``."""
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
    if addin_id is None:
        return None, u"No hay AddInId (ejecutar desde pyRevit)."
    if is_rebar_pata_l_diameter_dmu_registered(addin_id):
        unregister_rebar_pata_l_diameter_updater(addin_id)
        if doc is not None:
            set_document_concrete_grade_override(doc, None)
        return False, u"DMU pata L por diámetro: DESACTIVADO."
    ok = register_rebar_pata_l_diameter_updater(addin_id, doc=doc)
    if not ok:
        return False, u"No se pudo registrar el updater."
    if doc is not None:
        set_document_concrete_grade_override(doc, None)
    return (
        True,
        u"DMU pata L por diámetro: ACTIVADO (auto G25/G35/G45).\n"
        u"Detecta dosificación en parámetros del Rebar, host, proyecto o material.\n"
        u"Vuelva a pulsar el botón para forzar G25 → G35 → G45 → desactivar.",
    )


def cycle_rebar_pata_l_diameter_dmu(addin_id=None, doc=None):
    """
    Ciclo: off → auto → G25 → G35 → G45 → off.
    Retorna ``(registered_now, message)``.
    """
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
    if addin_id is None:
        return None, u"No hay AddInId (ejecutar desde pyRevit)."
    if not is_rebar_pata_l_diameter_dmu_registered(addin_id):
        ok = register_rebar_pata_l_diameter_updater(addin_id, doc=doc)
        if not ok:
            return False, u"No se pudo registrar el updater."
        if doc is not None:
            set_document_concrete_grade_override(doc, None)
        return (
            True,
            u"DMU pata L por diámetro: ACTIVADO (auto G25/G35/G45).\n"
            u"Busca dosificación en Rebar, host, proyecto o material estructural.",
        )
    override = get_document_concrete_grade_override(doc) if doc is not None else None
    if override is None:
        if doc is not None:
            set_document_concrete_grade_override(doc, u"G25")
        return True, u"DMU activo — tabla G25 (forzada)."
    if override == u"G25":
        if doc is not None:
            set_document_concrete_grade_override(doc, u"G35")
        return True, u"DMU activo — tabla G35 (forzada)."
    if override == u"G35":
        if doc is not None:
            set_document_concrete_grade_override(doc, u"G45")
        return True, u"DMU activo — tabla G45 (forzada)."
    unregister_rebar_pata_l_diameter_updater(addin_id)
    if doc is not None:
        set_document_concrete_grade_override(doc, None)
    return False, u"DMU pata L por diámetro: DESACTIVADO."


def run(__revit__):
    """Entrada pushbutton: toggle on/off + TaskDialog."""
    from Autodesk.Revit.UI import TaskDialog

    addin_id = _addin_id_pyrevit_or_none()
    doc = None
    try:
        uidoc = __revit__.ActiveUIDocument
        if uidoc is not None:
            doc = uidoc.Document
    except Exception:
        doc = None
    _on, msg = cycle_rebar_pata_l_diameter_dmu(addin_id, doc=doc)
    title = u"Arainco: Pata L por diámetro (DMU)"
    try:
        TaskDialog.Show(title, msg or u"?")
    except Exception:
        print(msg)
