# -*- coding: utf-8 -*-
"""
DMU: al usar el comando nativo Sección desde una Building Section, asigna a la
sección nueva el ``ViewFamilyType`` Detail cuyo nombre contiene el
«Section Filter» de la vista origen (p. ej. ``02_MA`` → ``Detail (02_MA)``).

Si ``ChangeTypeId`` a Detail no es posible (familia distinta), usa como
respaldo el Building Section homónimo (``Building Section (02_MA)``).

Flujo:
  1. Hook ``command-before-exec[ID_SECTION]`` (y variantes) cachea origen + tipos.
  2. ``IUpdater`` en ``ElementAddition`` de ``ViewSection``.
  3. Trabajo diferido con ``ExternalEvent`` (sin Transaction en Execute).
  4. Respaldo: si no hay caché del hook, lee ActiveView en el momento de la adición.

Revit 2024–2026 / IronPython (pyRevit).
"""

from __future__ import print_function

import time

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
    ViewFamily,
    ViewFamilyType,
    ViewSection,
    ViewType,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System import Guid

UPDATER_GUID = Guid("d4e5f6a7-b8c9-4d0e-9f1a-2b3c4d5e6f70")

_TXN = u"Arainco: Tipo sección según Section Filter"
_CACHE_TTL_SEC = 180.0

# doc_key -> cache dict
_origin_cache_by_doc = {}
_pending_by_doc = {}
_type_event = None


def _ensure_type_event():
    global _type_event
    if _type_event is None:
        _type_event = ExternalEvent.Create(_SectionTypeFromFilterHandler())
    return _type_event


def _doc_key(doc):
    if doc is None:
        return None
    try:
        return id(doc)
    except Exception:
        return None


def _eid_int(eid):
    if eid is None:
        return None
    try:
        return int(eid.IntegerValue)
    except Exception:
        return None


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


def _es_building_section_view(doc, view):
    if view is None:
        return False
    try:
        from filtro_armadura_eje import es_vista_building_section

        return bool(es_vista_building_section(view))
    except Exception:
        pass
    try:
        if getattr(view, "IsTemplate", False):
            return False
    except Exception:
        pass
    try:
        if not isinstance(view, ViewSection):
            return False
    except Exception:
        return False
    try:
        return view.ViewType == ViewType.Section
    except Exception:
        return False


def _es_detail_view(doc, view):
    if view is None:
        return False
    try:
        if getattr(view, "IsTemplate", False):
            return False
    except Exception:
        pass
    try:
        if not isinstance(view, ViewSection):
            return False
    except Exception:
        return False
    try:
        if view.ViewType == ViewType.Detail:
            return True
    except Exception:
        pass
    try:
        vft = doc.GetElement(view.GetTypeId()) if doc is not None else None
        if vft is not None and isinstance(vft, ViewFamilyType):
            return vft.ViewFamily == ViewFamily.Detail
    except Exception:
        pass
    return False


def _es_seccion_creada_por_comando(doc, view):
    """Vista nueva candidata: Section o Detail (no plantilla)."""
    if view is None:
        return False
    try:
        if getattr(view, "IsTemplate", False):
            return False
    except Exception:
        pass
    try:
        if not isinstance(view, ViewSection):
            return False
    except Exception:
        return False
    if _es_detail_view(doc, view):
        return True
    try:
        return view.ViewType == ViewType.Section
    except Exception:
        return False


def _resolve_types_from_origin(doc, view):
    """
    Lee Section Filter de ``view`` y resuelve VFT Detail + Building Section.

    Returns:
        dict con section_filter, detail_vft_id, section_vft_id (ints o None),
        o None si no hay filtro usable.
    """
    if doc is None or view is None:
        return None

    sf_text = None
    detail_vft_id = None
    section_vft_id = None
    find_detail = None

    try:
        from seccion_detalle_extremo_muro import (
            leer_section_filter_texto,
            find_view_family_type_detail_by_name,
        )

        find_detail = find_view_family_type_detail_by_name
        sf_text, _err = leer_section_filter_texto(doc, view)
    except Exception:
        sf_text = None

    if not sf_text:
        return None

    sf_text = _as_unicode(sf_text).strip()
    if not sf_text:
        return None

    if find_detail is not None:
        try:
            vft_d, _e = find_detail(doc, sf_text)
            if vft_d is not None:
                detail_vft_id = _eid_int(vft_d.Id)
        except Exception:
            pass

    try:
        from elevacion_eje import find_building_section_type_by_section_filter

        vft_s, _e = find_building_section_type_by_section_filter(doc, sf_text)
        if vft_s is not None:
            section_vft_id = _eid_int(vft_s.Id)
    except Exception:
        pass

    if detail_vft_id is None and section_vft_id is None:
        return None

    return {
        u"section_filter": sf_text,
        u"detail_vft_id": detail_vft_id,
        u"section_vft_id": section_vft_id,
    }


def cache_section_command_origin(doc, view):
    """
    Guarda la vista activa al lanzar el comando Sección.

    Requiere Building Section origen con «Section Filter» y al menos un VFT
    Detail o Building Section coincidente.
    """
    key = _doc_key(doc)
    if key is None or view is None:
        return False
    try:
        if getattr(view, "IsTemplate", False):
            return False
    except Exception:
        pass
    if not _es_building_section_view(doc, view):
        return False

    resolved = _resolve_types_from_origin(doc, view)
    if not resolved:
        return False

    view_id = _eid_int(getattr(view, "Id", None))
    if view_id is None:
        return False

    entry = dict(resolved)
    entry[u"origin_view_id"] = view_id
    entry[u"ts"] = time.time()
    _origin_cache_by_doc[key] = entry
    return True


def _cache_from_active_view_fallback(doc, added_ids):
    """Si el hook no cacheó, intenta ActiveView (padre al dibujar la sección)."""
    try:
        from pyrevit import HOST_APP

        uiapp = getattr(HOST_APP, "uiapp", None)
        if uiapp is None:
            return None
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None or uidoc.Document is None:
            return None
        if id(uidoc.Document) != id(doc):
            return None
        view = uidoc.ActiveView
    except Exception:
        return None

    if view is None or not _es_building_section_view(doc, view):
        return None

    avid = _eid_int(getattr(view, "Id", None))
    if avid is None:
        return None
    added_ints = set()
    for eid in added_ids or []:
        iv = _eid_int(eid)
        if iv is not None:
            added_ints.add(iv)
    if avid in added_ints:
        return None

    resolved = _resolve_types_from_origin(doc, view)
    if not resolved:
        return None
    entry = dict(resolved)
    entry[u"origin_view_id"] = avid
    entry[u"ts"] = time.time()
    return entry


def _get_valid_cache(doc, peek_only=False):
    key = _doc_key(doc)
    if key is None:
        return None
    entry = _origin_cache_by_doc.get(key)
    if not entry:
        return None
    try:
        age = time.time() - float(entry.get(u"ts") or 0)
    except Exception:
        age = _CACHE_TTL_SEC + 1.0
    if age > _CACHE_TTL_SEC:
        try:
            del _origin_cache_by_doc[key]
        except Exception:
            pass
        return None
    if not peek_only:
        try:
            del _origin_cache_by_doc[key]
        except Exception:
            pass
    return entry


def clear_section_command_origin(doc=None):
    if doc is None:
        _origin_cache_by_doc.clear()
        return
    key = _doc_key(doc)
    if key is not None and key in _origin_cache_by_doc:
        try:
            del _origin_cache_by_doc[key]
        except Exception:
            pass


def _enqueue_section_ids(doc, section_ids, cache_entry):
    global _pending_by_doc
    key = _doc_key(doc)
    if key is None:
        return
    new_ints = set()
    for eid in section_ids:
        iv = _eid_int(eid)
        if iv is not None:
            new_ints.add(iv)
    if not new_ints:
        return
    if key in _pending_by_doc:
        _doc_ref, existing, _old = _pending_by_doc[key]
        existing |= new_ints
        _pending_by_doc[key] = (_doc_ref, existing, cache_entry or _old)
    else:
        _pending_by_doc[key] = (doc, new_ints, cache_entry)
    try:
        _ensure_type_event().Raise()
    except Exception:
        pass


def _drain_pending():
    global _pending_by_doc
    out = list(_pending_by_doc.values())
    _pending_by_doc = {}
    return out


def _vft_from_cache_id(doc, vft_id):
    if vft_id is None or doc is None:
        return None
    try:
        el = doc.GetElement(ElementId(int(vft_id)))
    except Exception:
        return None
    if el is None or not isinstance(el, ViewFamilyType):
        return None
    return el


def _target_vft_for_view(doc, view, cache_entry):
    """
    Elige el VFT a aplicar:
      - Vista Detail → Detail VFT (prioridad).
      - Vista Section → Detail VFT primero; si no hay, Building Section VFT.
    """
    if not cache_entry:
        return None

    detail_vft = _vft_from_cache_id(doc, cache_entry.get(u"detail_vft_id"))
    section_vft = _vft_from_cache_id(doc, cache_entry.get(u"section_vft_id"))
    sf_text = cache_entry.get(u"section_filter")

    if detail_vft is None and section_vft is None and sf_text:
        origin = None
        origin_id = cache_entry.get(u"origin_view_id")
        if origin_id is not None:
            try:
                origin = doc.GetElement(ElementId(int(origin_id)))
            except Exception:
                origin = None
        if origin is not None:
            resolved = _resolve_types_from_origin(doc, origin)
            if resolved:
                detail_vft = _vft_from_cache_id(doc, resolved.get(u"detail_vft_id"))
                section_vft = _vft_from_cache_id(doc, resolved.get(u"section_vft_id"))

    if _es_detail_view(doc, view):
        return detail_vft or section_vft

    # Section creada por comando Sección: preferir Detail (pedido de negocio).
    return detail_vft or section_vft


def _aplicar_tipo_a_secciones(doc, section_ints, cache_entry):
    if doc is None or not section_ints or not cache_entry:
        return 0, 0, None

    origin_view_id = cache_entry.get(u"origin_view_id")
    n_ok = 0
    n_skip = 0

    t = Transaction(doc, _TXN)
    try:
        t.Start()
    except Exception as ex:
        return 0, 0, u"No se pudo abrir transacción: {0}".format(_as_unicode(ex))

    try:
        for iv in sorted(section_ints):
            try:
                view = doc.GetElement(ElementId(int(iv)))
            except Exception:
                view = None
            if view is None or not isinstance(view, ViewSection):
                n_skip += 1
                continue
            if origin_view_id is not None and iv == int(origin_view_id):
                n_skip += 1
                continue
            if not _es_seccion_creada_por_comando(doc, view):
                n_skip += 1
                continue

            detail_vft = _vft_from_cache_id(doc, cache_entry.get(u"detail_vft_id"))
            section_vft = _vft_from_cache_id(doc, cache_entry.get(u"section_vft_id"))
            if detail_vft is None and section_vft is None:
                vft = _target_vft_for_view(doc, view, cache_entry)
                if vft is None:
                    n_skip += 1
                    continue
                try:
                    if vft.ViewFamily == ViewFamily.Detail:
                        detail_vft = vft
                    else:
                        section_vft = vft
                except Exception:
                    section_vft = vft

            applied = False

            # El comando Sección crea ViewFamily.Section: aplicar primero el
            # Building Section del filtro (ChangeTypeId válido). Detail solo
            # si la vista nueva ya es Detail, o como intento secundario.
            candidates = []
            if _es_detail_view(doc, view):
                if detail_vft is not None:
                    candidates.append(detail_vft)
                if section_vft is not None:
                    candidates.append(section_vft)
            else:
                if section_vft is not None:
                    candidates.append(section_vft)
                if detail_vft is not None:
                    candidates.append(detail_vft)

            for vft in candidates:
                try:
                    cur = view.GetTypeId()
                    if (
                        cur is not None
                        and int(cur.IntegerValue) == int(vft.Id.IntegerValue)
                    ):
                        applied = True
                        break
                    view.ChangeTypeId(vft.Id)
                    applied = True
                    break
                except Exception:
                    continue

            if applied:
                n_ok += 1
            else:
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


class _SectionTypeFromFilterHandler(IExternalEventHandler):
    def GetName(self):
        return u"Arainco: DMU tipo sección Section Filter (diferido)"

    def Execute(self, uiapp):
        pending = _drain_pending()
        if not pending:
            return
        for doc, section_ints, cache_entry in pending:
            try:
                if doc is None or not doc.IsValidObject or doc.IsLinked:
                    continue
            except Exception:
                continue
            _aplicar_tipo_a_secciones(doc, section_ints, cache_entry)


class SectionTypeFromFilterUpdater(IUpdater):
    def __init__(self, addin_id):
        self._updater_id = UpdaterId(addin_id, UPDATER_GUID)

    def GetUpdaterId(self):
        return self._updater_id

    def GetUpdaterName(self):
        return u"Arainco: Tipo sección según Section Filter"

    def GetAdditionalInformation(self):
        return (
            u"Al crear una sección con el comando Sección desde una Building "
            u"Section, asigna el tipo Detail (o Building Section) cuyo nombre "
            u"contiene el «Section Filter» de la vista origen."
        )

    def GetChangePriority(self):
        return ChangePriority.Views

    def Execute(self, data):
        doc = data.GetDocument()
        if doc is None or doc.IsLinked:
            return

        added = []
        try:
            added = list(data.GetAddedElementIds())
        except Exception:
            added = []
        if not added:
            return

        section_ids = []
        for eid in added:
            try:
                el = doc.GetElement(eid)
            except Exception:
                el = None
            if el is None or not isinstance(el, ViewSection):
                continue
            if not _es_seccion_creada_por_comando(doc, el):
                continue
            section_ids.append(eid)

        if not section_ids:
            return

        entry = _get_valid_cache(doc, peek_only=True)
        if entry is None:
            entry = _cache_from_active_view_fallback(doc, section_ids)

        if entry is None:
            return

        # Consumir caché del hook (si existía) al encolar.
        _get_valid_cache(doc, peek_only=False)
        _enqueue_section_ids(doc, section_ids, entry)


def register_section_type_from_filter_updater(addin_id, doc=None):
    _ensure_type_event()
    updater = SectionTypeFromFilterUpdater(addin_id)
    uid = updater.GetUpdaterId()
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        try:
            UpdaterRegistry.UnregisterUpdater(uid)
        except Exception:
            pass
    UpdaterRegistry.RegisterUpdater(updater)
    flt = ElementClassFilter(ViewSection)
    try:
        ct_add = Element.GetChangeTypeElementAddition()
        if doc is None:
            UpdaterRegistry.AddTrigger(uid, flt, ct_add)
        else:
            UpdaterRegistry.AddTrigger(uid, doc, flt, ct_add)
    except Exception:
        pass


def unregister_section_type_from_filter_updater(addin_id):
    uid = UpdaterId(addin_id, UPDATER_GUID)
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        UpdaterRegistry.UnregisterUpdater(uid)
    clear_section_command_origin(None)


def _addin_id_pyrevit_or_none():
    try:
        from pyrevit import HOST_APP

        return HOST_APP.addin_id
    except Exception:
        return None


def is_section_type_from_filter_dmu_registered(addin_id=None):
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
        if addin_id is None:
            return None
    uid = UpdaterId(addin_id, UPDATER_GUID)
    return UpdaterRegistry.IsUpdaterRegistered(uid)


def toggle_section_type_from_filter_dmu(addin_id=None, doc=None):
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
    if addin_id is None:
        return None, u"No hay AddInId (ejecutar desde pyRevit)."
    if is_section_type_from_filter_dmu_registered(addin_id):
        unregister_section_type_from_filter_updater(addin_id)
        return False, (
            u"DMU tipo sección según Section Filter: DESACTIVADO.\n"
            u"Las secciones nuevas conservarán el tipo por defecto de Revit."
        )
    register_section_type_from_filter_updater(addin_id, doc=doc)
    return True, (
        u"DMU tipo sección según Section Filter: ACTIVADO.\n"
        u"Al usar el comando Sección desde una Building Section, se asigna el "
        u"tipo Detail (o Building Section) cuyo nombre contiene el "
        u"«Section Filter» (p. ej. 02_MA → Detail (02_MA))."
    )


def handle_section_command_before_executed(uiapp, eventargs):
    """Hook: no cancela el comando; cachea Building Section + Section Filter."""
    if uiapp is None:
        return
    try:
        uidoc = uiapp.ActiveUIDocument
    except Exception:
        uidoc = None
    if uidoc is None:
        return
    try:
        doc = uidoc.Document
        view = uidoc.ActiveView
    except Exception:
        return
    cache_section_command_origin(doc, view)


def run(__revit__):
    from Autodesk.Revit.UI import TaskDialog

    addin_id = _addin_id_pyrevit_or_none()
    doc = None
    try:
        uidoc = __revit__.ActiveUIDocument
        if uidoc is not None:
            doc = uidoc.Document
    except Exception:
        doc = None
    _on, msg = toggle_section_type_from_filter_dmu(addin_id, doc=doc)
    title = u"Arainco: Tipo sección según Section Filter (DMU)"
    try:
        TaskDialog.Show(title, msg or u"?")
    except Exception:
        print(msg)
