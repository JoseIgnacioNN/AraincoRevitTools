# -*- coding: utf-8 -*-
"""
DMU: al insertar una vista Detail en una lámina (Viewport), asigna el tipo
de viewport ``Seccion``.

El trabajo se difiere con ExternalEvent (no dentro de IUpdater.Execute).

Revit 2024–2026 / IronPython (pyRevit).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ChangePriority,
    Element,
    ElementClassFilter,
    ElementId,
    ElementType,
    FilteredElementCollector,
    IUpdater,
    Transaction,
    TransactionStatus,
    UpdaterId,
    UpdaterRegistry,
    ViewFamily,
    ViewFamilyType,
    ViewType,
    Viewport,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System import Guid

UPDATER_GUID = Guid("3f8a1c62-7e4b-4d91-b5a0-2c9e6f4d8a17")

_TXN = u"Arainco: Viewport Detail = Seccion"
_VIEWPORT_TYPE_NAMES = (u"Seccion", u"Sección")

_pending_by_doc = {}
_type_event = None
_type_id_cache = {}


def _ensure_type_event():
    global _type_event
    if _type_event is None:
        _type_event = ExternalEvent.Create(_DetailViewViewportTypeHandler())
    return _type_event


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
        _ensure_type_event().Raise()
    except Exception:
        pass


def _drain_pending():
    global _pending_by_doc
    out = list(_pending_by_doc.values())
    _pending_by_doc = {}
    return out


def _element_type_name(el):
    if el is None:
        return u""
    for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = el.get_Parameter(bip)
            if p is None:
                continue
            raw = p.AsString()
            if not raw:
                raw = p.AsValueString()
            s = _as_unicode(raw).strip()
            if s:
                return s
        except Exception:
            continue
    try:
        s = _as_unicode(Element.Name.__get__(el)).strip()
        if s:
            return s
    except Exception:
        pass
    try:
        return _as_unicode(el.Name).strip()
    except Exception:
        return u""


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


def _wanted_type_names():
    wanted = []
    seen = set()
    for name in _VIEWPORT_TYPE_NAMES:
        key = _as_unicode(name).strip()
        if not key:
            continue
        low = key.lower()
        if low in seen:
            continue
        seen.add(low)
        wanted.append(key)
    return wanted


def _iter_viewport_types(doc):
    if doc is None:
        return
    try:
        col = (
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Viewports)
            .WhereElementIsElementType()
        )
        for t in col:
            yield t
        return
    except Exception:
        pass
    try:
        col = (
            FilteredElementCollector(doc)
            .OfClass(ElementType)
            .OfCategory(BuiltInCategory.OST_Viewports)
        )
        for t in col:
            yield t
    except Exception:
        return


def _pick_wanted_type(candidates):
    by_lower = {}
    for t in candidates or ():
        if t is None:
            continue
        name = _as_unicode(_element_type_name(t)).strip()
        if not name:
            continue
        low = name.lower()
        if low not in by_lower:
            by_lower[low] = t
    for wanted in _wanted_type_names():
        found = by_lower.get(wanted.lower())
        if found is not None:
            return found
    return None


def _types_from_viewport(doc, viewport):
    out = []
    if doc is None or viewport is None:
        return out
    try:
        ids = viewport.GetValidTypes()
    except Exception:
        return out
    for tid in ids or ():
        try:
            t = doc.GetElement(tid)
        except Exception:
            t = None
        if t is not None:
            out.append(t)
    return out


def _find_viewport_type(doc, viewport=None):
    """ElementType Viewport cuyo nombre es Seccion (o Sección)."""
    global _type_id_cache
    if doc is None:
        return None
    key = id(doc)
    cached = _type_id_cache.get(key)
    if cached is not None:
        try:
            el = doc.GetElement(ElementId(int(cached)))
        except Exception:
            el = None
        if el is not None:
            return el
        _type_id_cache.pop(key, None)

    found = _pick_wanted_type(_types_from_viewport(doc, viewport))
    if found is None:
        found = _pick_wanted_type(_iter_viewport_types(doc))
    if found is None:
        return None
    iv = _eid_int(found.Id)
    if iv is not None:
        _type_id_cache[key] = iv
    return found


def _viewport_already_target(viewport, type_id):
    if viewport is None or type_id is None:
        return False
    try:
        cur = viewport.GetTypeId()
        return (
            cur is not None
            and int(cur.IntegerValue) == int(type_id.IntegerValue)
        )
    except Exception:
        return False


def _change_viewport_type(viewport, type_id):
    if viewport is None or type_id is None:
        return False
    try:
        if hasattr(viewport, "IsValidType") and (not viewport.IsValidType(type_id)):
            return False
    except Exception:
        pass
    try:
        viewport.ChangeTypeId(type_id)
        return True
    except Exception:
        return False


def _aplicar_tipos(doc, element_ints):
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

    planned = []
    seen = set()
    for iv in element_ints or ():
        if iv in seen:
            continue
        seen.add(iv)
        try:
            el = doc.GetElement(ElementId(iv))
        except Exception:
            el = None
        if el is None or not isinstance(el, Viewport):
            continue
        view = _viewport_view(doc, el)
        if not _es_detail_view(doc, view):
            continue
        planned.append(el)

    if not planned:
        return 0, 0, None

    type_el = _find_viewport_type(doc, viewport=planned[0])
    if type_el is None:
        return 0, 0, None
    type_id = type_el.Id
    planned = [vp for vp in planned if not _viewport_already_target(vp, type_id)]
    if not planned:
        return 0, 0, None

    n_ok = 0
    n_skip = 0
    t = Transaction(doc, _TXN)
    try:
        t.Start()
        for vp in planned:
            try:
                if _viewport_already_target(vp, type_id):
                    continue
                if _change_viewport_type(vp, type_id):
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


class _DetailViewViewportTypeHandler(IExternalEventHandler):
    def GetName(self):
        return u"Arainco: DMU viewport Detail Seccion (diferido)"

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
            _aplicar_tipos(doc, element_ints)


class DetailViewViewportTypeUpdater(IUpdater):
    def __init__(self, addin_id):
        self._updater_id = UpdaterId(addin_id, UPDATER_GUID)

    def GetUpdaterId(self):
        return self._updater_id

    def GetUpdaterName(self):
        return u"Arainco: Viewport Detail = Seccion"

    def GetAdditionalInformation(self):
        return (
            u"Al insertar una vista Detail en una lámina, asigna el tipo de "
            u"viewport «Seccion»."
        )

    def GetChangePriority(self):
        return ChangePriority.Views

    def Execute(self, data):
        doc = data.GetDocument()
        if doc is None:
            return
        try:
            if doc.IsLinked:
                return
        except Exception:
            pass

        try:
            eids = list(data.GetAddedElementIds())
        except Exception:
            eids = []

        ids = []
        seen = set()
        for eid in eids:
            iv = _eid_int(eid)
            if iv is None or iv in seen:
                continue
            try:
                el = doc.GetElement(eid)
            except Exception:
                el = None
            if el is None or not isinstance(el, Viewport):
                continue
            view = _viewport_view(doc, el)
            if not _es_detail_view(doc, view):
                continue
            seen.add(iv)
            ids.append(eid)

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


def register_detail_view_viewport_type_updater(addin_id, doc=None):
    _ensure_type_event()
    updater = DetailViewViewportTypeUpdater(addin_id)
    uid = updater.GetUpdaterId()
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        try:
            UpdaterRegistry.UnregisterUpdater(uid)
        except Exception:
            pass
    UpdaterRegistry.RegisterUpdater(updater)
    flt_vp = ElementClassFilter(Viewport)
    _add_trigger(uid, doc, flt_vp, Element.GetChangeTypeElementAddition())


def unregister_detail_view_viewport_type_updater(addin_id):
    uid = UpdaterId(addin_id, UPDATER_GUID)
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        UpdaterRegistry.UnregisterUpdater(uid)


def _addin_id_pyrevit_or_none():
    try:
        from pyrevit import HOST_APP

        return HOST_APP.addin_id
    except Exception:
        return None


def is_detail_view_viewport_type_dmu_registered(addin_id=None):
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
        if addin_id is None:
            return None
    uid = UpdaterId(addin_id, UPDATER_GUID)
    return UpdaterRegistry.IsUpdaterRegistered(uid)


def _aviso_tipo_faltante(doc):
    if doc is None:
        return u""
    try:
        if _find_viewport_type(doc) is not None:
            return u""
    except Exception:
        return u""
    names = u" / ".join(_wanted_type_names())
    return (
        u"En este proyecto no se encontró el tipo de viewport «{0}». "
        u"El DMU quedará activo y aplicará el tipo cuando exista en el documento."
    ).format(names)


def toggle_detail_view_viewport_type_dmu(addin_id=None, doc=None):
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
    if addin_id is None:
        return None, u"No hay AddInId (ejecutar desde pyRevit)."
    if is_detail_view_viewport_type_dmu_registered(addin_id):
        unregister_detail_view_viewport_type_updater(addin_id)
        return False, (
            u"DMU viewport Detail: DESACTIVADO.\n"
            u"Los viewports de vistas Detail nuevas conservarán el tipo que asigne Revit."
        )
    register_detail_view_viewport_type_updater(addin_id, doc=doc)
    extra = _aviso_tipo_faltante(doc)
    msg = (
        u"DMU viewport Detail: ACTIVADO.\n"
        u"Al insertar una vista Detail en una lámina, el viewport pasa al tipo Seccion."
    )
    if extra:
        msg = msg + u"\n\n" + extra
    return True, msg


def _mostrar_aviso(uiapp, instruction, content=u""):
    title = u"Arainco: Viewport Detail Seccion (DMU)"
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
    _on, msg = toggle_detail_view_viewport_type_dmu(addin_id, doc=doc)
    lines = _as_unicode(msg or u"?").split(u"\n", 1)
    instruction = lines[0]
    content = lines[1].strip() if len(lines) > 1 else u""
    _mostrar_aviso(__revit__, instruction, content)
