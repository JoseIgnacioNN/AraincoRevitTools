# -*- coding: utf-8 -*-
"""
DMU: al insertar una vista Detail en una lámina (Viewport), renombra el
View Name a ``[Sheet Number]_[Detail Number]``.

También actualiza el View Name si cambian:
  - Detail Number de la vista / viewport
  - Sheet Number de la vista (p. ej. al moverla de lámina)
  - Sheet Number de la lámina (renumerar el plano)

El trabajo se difiere con ExternalEvent (no dentro de IUpdater.Execute).
Respaldo: Application.DocumentChanged (Revit 2024 a veces no dispara IUpdater).

Revit 2024–2026 / IronPython (pyRevit).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ChangePriority,
    Element,
    ElementClassFilter,
    ElementId,
    FilteredElementCollector,
    IUpdater,
    Transaction,
    TransactionStatus,
    UpdaterId,
    UpdaterRegistry,
    View,
    ViewFamily,
    ViewFamilyType,
    ViewSection,
    ViewSheet,
    ViewType,
    Viewport,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System import Guid

UPDATER_GUID = Guid("9e4c7b12-8a3d-4f61-b2e5-7c18d90a4f33")

_TXN = u"Arainco: Nombre Detail = lámina_detalle"
_PLACEHOLDERS = frozenset(
    (
        u"",
        u"-",
        u"--",
        u"---",
        u"not on sheets",
        u"not on sheet",
        u"no en láminas",
        u"no en laminas",
    )
)

_pending_by_doc = {}
_rename_event = None
_enabled = False
_doc_changed_subscribed = False
_idling_for_doc_changed_scheduled = False


def _ensure_rename_event():
    global _rename_event
    if _rename_event is None:
        _rename_event = ExternalEvent.Create(_DetailViewSheetRenameHandler())
    return _rename_event


def _as_unicode(value):
    if value is None:
        return u""
    try:
        if isinstance(value, unicode):
            return value
    except NameError:
        if isinstance(value, str):
            return value
    try:
        return unicode(value)
    except NameError:
        return str(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""


def _eid_int(eid):
    if eid is None:
        return None
    try:
        return int(eid.IntegerValue)
    except Exception:
        return None


def _enqueue_element_ids(doc, element_ids):
    global _pending_by_doc
    if doc is None:
        return
    new_ints = set()
    for eid in element_ids or ():
        iv = _eid_int(eid)
        if iv is not None:
            new_ints.add(iv)
    if not new_ints:
        return
    key = id(doc)
    if key in _pending_by_doc:
        _doc_ref, existing = _pending_by_doc[key]
        existing |= new_ints
    else:
        _pending_by_doc[key] = (doc, new_ints)
    try:
        _ensure_rename_event().Raise()
    except Exception:
        pass


def _drain_pending():
    global _pending_by_doc
    out = list(_pending_by_doc.values())
    _pending_by_doc = {}
    return out


def _param_string(el, bip):
    if el is None:
        return u""
    try:
        p = el.get_Parameter(bip)
    except Exception:
        p = None
    if p is None:
        return u""
    try:
        raw = p.AsString()
    except Exception:
        raw = None
    if raw is None:
        try:
            raw = p.AsValueString()
        except Exception:
            raw = None
    return _as_unicode(raw).strip()


def _is_usable_token(text):
    s = _as_unicode(text).strip()
    if not s:
        return False
    return s.lower() not in _PLACEHOLDERS


def _es_detail_view(doc, view):
    if view is None:
        return False
    try:
        if getattr(view, "IsTemplate", False):
            return False
    except Exception:
        pass
    try:
        if view.ViewType == ViewType.Detail:
            return True
    except Exception:
        pass
    try:
        if doc is None:
            return False
        vft = doc.GetElement(view.GetTypeId())
        if vft is not None and isinstance(vft, ViewFamilyType):
            return vft.ViewFamily == ViewFamily.Detail
    except Exception:
        pass
    return False


def _viewport_ids_on_sheet(doc, sheet):
    ids = []
    if sheet is None:
        return ids
    try:
        ids = list(sheet.GetAllViewports())
    except Exception:
        ids = []
    if ids:
        return ids
    if doc is None:
        return ids
    try:
        for vp in FilteredElementCollector(doc).OfClass(Viewport):
            try:
                if vp.SheetId == sheet.Id:
                    ids.append(vp.Id)
            except Exception:
                continue
    except Exception:
        pass
    return ids


def _first_viewport_of_view(doc, view):
    if doc is None or view is None:
        return None
    try:
        for vp in FilteredElementCollector(doc).OfClass(Viewport):
            try:
                if vp.ViewId == view.Id:
                    return vp
            except Exception:
                continue
    except Exception:
        pass
    return None


def _sheet_number_from_viewport(doc, viewport):
    if viewport is None:
        return u""
    sheet = _param_string(viewport, BuiltInParameter.VIEWPORT_SHEET_NUMBER)
    if _is_usable_token(sheet):
        return sheet
    if doc is None:
        return u""
    try:
        sh = doc.GetElement(viewport.SheetId)
    except Exception:
        sh = None
    if sh is None:
        return u""
    sheet = _param_string(sh, BuiltInParameter.SHEET_NUMBER)
    if _is_usable_token(sheet):
        return sheet
    try:
        return _as_unicode(sh.SheetNumber).strip()
    except Exception:
        return u""


def _nombre_objetivo(view, viewport=None, doc=None):
    sheet = _param_string(view, BuiltInParameter.VIEWER_SHEET_NUMBER)
    detail = _param_string(view, BuiltInParameter.VIEWER_DETAIL_NUMBER)
    if viewport is None and doc is not None and (
        not _is_usable_token(sheet) or not _is_usable_token(detail)
    ):
        viewport = _first_viewport_of_view(doc, view)
    if viewport is not None:
        if not _is_usable_token(sheet):
            sheet = _sheet_number_from_viewport(doc, viewport)
        if not _is_usable_token(detail):
            detail = _param_string(viewport, BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
    if not _is_usable_token(sheet) or not _is_usable_token(detail):
        return None
    return u"{0}_{1}".format(sheet.strip(), detail.strip())


def _view_name_actual(view):
    if view is None:
        return u""
    try:
        return _as_unicode(view.Name).strip()
    except Exception:
        return _param_string(view, BuiltInParameter.VIEW_NAME)


def _set_view_name(view, name):
    if view is None or not name:
        return False
    try:
        view.Name = name
        return True
    except Exception:
        pass
    try:
        p = view.get_Parameter(BuiltInParameter.VIEW_NAME)
        if p is not None and (not p.IsReadOnly):
            p.Set(name)
            return True
    except Exception:
        pass
    return False


def _viewport_view(doc, viewport):
    if doc is None or viewport is None:
        return None
    try:
        view_id = viewport.ViewId
    except Exception:
        view_id = None
    if view_id is None:
        return None
    try:
        if view_id == ElementId.InvalidElementId:
            return None
    except Exception:
        pass
    try:
        return doc.GetElement(view_id)
    except Exception:
        return None


def _resolve_view_and_viewport(doc, el):
    if el is None:
        return None, None
    if isinstance(el, Viewport):
        return _viewport_view(doc, el), el
    if isinstance(el, ViewSheet):
        return None, None
    if isinstance(el, View):
        return el, None
    return None, None


def _expand_related_ids(doc, element_ids):
    """Viewport / vista Detail / lámina → ids de Viewport y View a evaluar."""
    out_eids = []
    seen = set()

    def _add(eid):
        iv = _eid_int(eid)
        if iv is None or iv in seen:
            return
        seen.add(iv)
        out_eids.append(eid)

    for eid in element_ids or ():
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        if isinstance(el, ViewSheet):
            for vp_id in _viewport_ids_on_sheet(doc, el):
                _add(vp_id)
                try:
                    vp = doc.GetElement(vp_id)
                except Exception:
                    vp = None
                view = _viewport_view(doc, vp)
                if view is not None:
                    _add(view.Id)
            continue
        if isinstance(el, Viewport):
            _add(eid)
            view = _viewport_view(doc, el)
            if view is not None:
                _add(view.Id)
            continue
        if isinstance(el, View) and _es_detail_view(doc, el):
            _add(eid)
    return out_eids


def _plan_renombre(doc, element_ints):
    """Lista de (view, nuevo_nombre) pendientes; no abre transacción."""
    planned = []
    seen_view = set()
    used_targets = set()
    eids = [ElementId(iv) for iv in (element_ints or ())]
    for eid in _expand_related_ids(doc, eids):
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        view, vp = _resolve_view_and_viewport(doc, el)
        if not _es_detail_view(doc, view):
            continue
        view_int = _eid_int(view.Id)
        if view_int in seen_view:
            continue
        seen_view.add(view_int)
        target = _nombre_objetivo(view, vp, doc)
        if not target:
            continue
        current = _view_name_actual(view)
        if current == target:
            continue
        target_key = target.lower()
        if target_key in used_targets:
            continue
        used_targets.add(target_key)
        planned.append((view, target))
    return planned


def _aplicar_renombres(doc, element_ints):
    if doc is None:
        return 0, 0, u"Documento no válido."
    try:
        if not doc.IsValidObject or doc.IsLinked:
            return 0, 0, None
    except Exception:
        return 0, 0, None
    try:
        if getattr(doc, "IsFamilyDocument", False):
            return 0, 0, None
    except Exception:
        pass

    planned = _plan_renombre(doc, element_ints)
    if not planned:
        return 0, 0, None

    n_ok = 0
    n_skip = 0
    t = Transaction(doc, _TXN)
    try:
        t.Start()
        for view, target in planned:
            try:
                if _view_name_actual(view) == target:
                    continue
                if _set_view_name(view, target):
                    n_ok += 1
                else:
                    n_skip += 1
            except Exception:
                n_skip += 1
        status = t.Commit()
        if status != TransactionStatus.Committed:
            try:
                t.RollBack()
            except Exception:
                pass
            return 0, n_skip, u"La transacción no se confirmó."
    except Exception as ex:
        try:
            if t.HasStarted():
                t.RollBack()
        except Exception:
            pass
        return 0, n_skip, _as_unicode(ex)
    return n_ok, n_skip, None


def _collect_candidate_ids(doc, eids):
    """Filtra a Viewport, ViewSection/Detail y ViewSheet."""
    ids = []
    seen = set()
    for eid in eids or ():
        iv = _eid_int(eid)
        if iv is None or iv in seen:
            continue
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        if isinstance(el, Viewport) or isinstance(el, ViewSheet):
            seen.add(iv)
            ids.append(eid)
            continue
        if isinstance(el, View) and _es_detail_view(doc, el):
            seen.add(iv)
            ids.append(eid)
    return ids


class _DetailViewSheetRenameHandler(IExternalEventHandler):
    def GetName(self):
        return u"Arainco: DMU nombre Detail en lámina (diferido)"

    def Execute(self, uiapp):
        pending = _drain_pending()
        if not pending:
            return
        for doc, element_ints in pending:
            try:
                if doc is None or not doc.IsValidObject or doc.IsLinked:
                    continue
            except Exception:
                continue
            _aplicar_renombres(doc, element_ints)


class DetailViewSheetRenameUpdater(IUpdater):
    def __init__(self, addin_id):
        self._updater_id = UpdaterId(addin_id, UPDATER_GUID)

    def GetUpdaterId(self):
        return self._updater_id

    def GetUpdaterName(self):
        return u"Arainco: Nombre Detail = lámina_detalle"

    def GetAdditionalInformation(self):
        return (
            u"Al insertar una vista Detail en una lámina, o al cambiar Detail "
            u"Number / Sheet Number, renombra el View Name a "
            u"Sheet Number + '_' + Detail Number."
        )

    def GetChangePriority(self):
        return ChangePriority.Views

    def Execute(self, data):
        if not _enabled:
            return
        doc = data.GetDocument()
        if doc is None:
            return
        try:
            if doc.IsLinked:
                return
        except Exception:
            pass

        raw = []
        for getter in (data.GetAddedElementIds, data.GetModifiedElementIds):
            try:
                raw.extend(list(getter()))
            except Exception:
                pass
        ids = _collect_candidate_ids(doc, raw)
        if not ids:
            return
        _enqueue_element_ids(doc, ids)


def _add_trigger(uid, doc, flt, change_type):
    if change_type is None:
        return
    try:
        if doc is None:
            UpdaterRegistry.AddTrigger(uid, flt, change_type)
        else:
            UpdaterRegistry.AddTrigger(uid, doc, flt, change_type)
    except Exception:
        pass


def _any_or_add(uid, doc, flt, also_addition=False):
    try:
        ct_any = Element.GetChangeTypeAny()
    except Exception:
        ct_any = None
    _add_trigger(uid, doc, flt, ct_any)
    if also_addition:
        try:
            _add_trigger(uid, doc, flt, Element.GetChangeTypeElementAddition())
        except Exception:
            pass


def _subscribe_document_changed_if_needed(uiapp):
    global _doc_changed_subscribed
    if _doc_changed_subscribed or uiapp is None:
        return
    try:
        uiapp.Application.DocumentChanged += _on_application_document_changed
        _doc_changed_subscribed = True
    except Exception:
        pass


def _schedule_document_changed_subscription():
    global _idling_for_doc_changed_scheduled
    if _idling_for_doc_changed_scheduled or _doc_changed_subscribed:
        return
    _idling_for_doc_changed_scheduled = True
    try:
        from pyrevit import HOST_APP

        uiapp = getattr(HOST_APP, "uiapp", None) or getattr(HOST_APP, "app", None)
        if uiapp is None:
            _idling_for_doc_changed_scheduled = False
            return

        def _idling_once(sender, args):
            try:
                sender.Idling -= _idling_once
            except Exception:
                pass
            _subscribe_document_changed_if_needed(sender)

        uiapp.Idling += _idling_once
    except Exception:
        _idling_for_doc_changed_scheduled = False


def _on_application_document_changed(sender, args):
    if not _enabled:
        return
    try:
        names = list(args.GetTransactionNames())
        if _TXN in names:
            return
    except Exception:
        pass
    try:
        doc = args.GetDocument()
    except Exception:
        return
    if doc is None:
        return
    try:
        if doc.IsLinked:
            return
    except Exception:
        return
    raw = []
    try:
        raw.extend(list(args.GetAddedElementIds()))
    except Exception:
        pass
    try:
        raw.extend(list(args.GetModifiedElementIds()))
    except Exception:
        pass
    ids = _collect_candidate_ids(doc, raw)
    if ids:
        _enqueue_element_ids(doc, ids)


def register_detail_view_sheet_rename_updater(addin_id, doc=None):
    global _enabled
    _ensure_rename_event()
    updater = DetailViewSheetRenameUpdater(addin_id)
    uid = updater.GetUpdaterId()
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        try:
            UpdaterRegistry.UnregisterUpdater(uid)
        except Exception:
            pass
    UpdaterRegistry.RegisterUpdater(updater)
    # Viewport: inserción y cualquier cambio (Detail Number, lámina).
    _any_or_add(uid, doc, ElementClassFilter(Viewport), also_addition=True)
    # Vista Detail (ViewSection): cambio de Detail Number / Sheet Number en la vista.
    _any_or_add(uid, doc, ElementClassFilter(ViewSection), also_addition=False)
    # Lámina: renumerar Sheet Number actualiza todas las Detail colocadas.
    _any_or_add(uid, doc, ElementClassFilter(ViewSheet), also_addition=False)
    _enabled = True
    _schedule_document_changed_subscription()


def unregister_detail_view_sheet_rename_updater(addin_id):
    global _enabled
    _enabled = False
    uid = UpdaterId(addin_id, UPDATER_GUID)
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        UpdaterRegistry.UnregisterUpdater(uid)


def _addin_id_pyrevit_or_none():
    try:
        from pyrevit import HOST_APP

        return HOST_APP.addin_id
    except Exception:
        return None


def is_detail_view_sheet_rename_dmu_registered(addin_id=None):
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
        if addin_id is None:
            return None
    uid = UpdaterId(addin_id, UPDATER_GUID)
    return UpdaterRegistry.IsUpdaterRegistered(uid)


def toggle_detail_view_sheet_rename_dmu(addin_id=None, doc=None):
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
    if addin_id is None:
        return None, u"No hay AddInId (ejecutar desde pyRevit)."
    if is_detail_view_sheet_rename_dmu_registered(addin_id):
        unregister_detail_view_sheet_rename_updater(addin_id)
        return False, (
            u"DMU nombre Detail en lámina: DESACTIVADO.\n"
            u"Las vistas Detail conservarán el nombre que asigne Revit."
        )
    register_detail_view_sheet_rename_updater(addin_id, doc=doc)
    return True, (
        u"DMU nombre Detail en lámina: ACTIVADO.\n"
        u"Al insertar una Detail en lámina, o al cambiar Detail Number o "
        u"Sheet Number, el View Name pasa a [Sheet Number]_[Detail Number] "
        u"(p. ej. E-101_A)."
    )


def _mostrar_aviso(uiapp, instruction, content=u""):
    title = u"Arainco: Nombre Detail en lámina (DMU)"
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = None
        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
        if show_message_dialog(
            title,
            instruction=instruction,
            content=content,
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        ):
            return
    except Exception:
        pass
    from Autodesk.Revit.UI import TaskDialog

    text = instruction
    extra = _as_unicode(content).strip()
    if extra:
        text = instruction + u"\n\n" + extra
    TaskDialog.Show(title, text or u"?")


def run(__revit__):
    addin_id = _addin_id_pyrevit_or_none()
    doc = None
    try:
        uidoc = __revit__.ActiveUIDocument
        if uidoc is not None:
            doc = uidoc.Document
    except Exception:
        doc = None
    _on, msg = toggle_detail_view_sheet_rename_dmu(addin_id, doc=doc)
    lines = _as_unicode(msg or u"?").split(u"\n", 1)
    instruction = lines[0]
    content = lines[1].strip() if len(lines) > 1 else u""
    _mostrar_aviso(__revit__, instruction, content)
