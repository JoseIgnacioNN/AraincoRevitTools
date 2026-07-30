# -*- coding: utf-8 -*-
"""
DMU: co-mover lote de etiquetas muro + malla (V/H) en la misma vista.

Si el usuario mueve **una** etiqueta del lote, las hermanas reciben el mismo Δ
en ``TagHeadPosition``. Si mueve **todas** a la vez (selección múltiple), solo
se actualiza la caché (Revit ya las movió).

Trabajo diferido con ExternalEvent (nunca Transaction dentro de IUpdater.Execute).
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
    IndependentTag,
    IUpdater,
    Transaction,
    UpdaterId,
    UpdaterRegistry,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System import Guid

from armado_muros_tag_lote import (
    _eid_int,
    _head,
    _xyz_delta,
    delta_is_significant,
    is_lote_relevant_tag,
    move_tag_lote_by_delta,
    resolve_tag_lote,
)

# GUID propio (no reutilizar otros DMU).
UPDATER_GUID = Guid("7c8e9f0a-1b2c-4d3e-9f4a-5b6c7d8e9f01")

_TXN = u"Arainco: Co-mover lote etiquetas muro/malla (DMU)"

# Caché de cabezas: (doc_hash, tag_id_int) → XYZ
_head_cache = {}
# Ids que estamos escribiendo (anti-reentrada)
_writing_ids = set()
# Cola: doc_id → (doc, list of move jobs)
# job = {u"driver": int, u"delta": XYZ, u"lote_ids": set[int]}
_pending_by_doc = {}
_event = None


def _doc_key(doc):
    try:
        return id(doc)
    except Exception:
        return None


def _cache_key(doc, tag_id_int):
    return (_doc_key(doc), int(tag_id_int))


def _cache_get(doc, tag_id_int):
    return _head_cache.get(_cache_key(doc, tag_id_int))


def _cache_set(doc, tag_id_int, xyz):
    if xyz is None or tag_id_int is None:
        return
    _head_cache[_cache_key(doc, tag_id_int)] = xyz


def _cache_clear_doc(doc):
    dk = _doc_key(doc)
    doomed = [k for k in _head_cache if k[0] == dk]
    for k in doomed:
        _head_cache.pop(k, None)


def _ensure_event():
    global _event
    if _event is None:
        _event = ExternalEvent.Create(_TagLoteCoMoveHandler())
    return _event


def _enqueue_job(doc, job):
    global _pending_by_doc
    if doc is None or not job:
        return
    key = _doc_key(doc)
    if key is None:
        return
    if key in _pending_by_doc:
        _doc_ref, jobs = _pending_by_doc[key]
        # Coalescer por conjunto de lote
        lote = frozenset(job.get(u"lote_ids") or [])
        merged = False
        for existing in jobs:
            if frozenset(existing.get(u"lote_ids") or []) == lote:
                # Conservar el Δ del driver más reciente
                existing[u"driver"] = job.get(u"driver")
                existing[u"delta"] = job.get(u"delta")
                merged = True
                break
        if not merged:
            jobs.append(job)
    else:
        _pending_by_doc[key] = (doc, [job])
    try:
        _ensure_event().Raise()
    except Exception:
        pass


def _drain_pending():
    global _pending_by_doc
    out = list(_pending_by_doc.values())
    _pending_by_doc = {}
    return out


class _TagLoteCoMoveHandler(IExternalEventHandler):
    def GetName(self):
        return u"Arainco: DMU co-mover lote etiquetas muro/malla"

    def Execute(self, uiapp):
        pending = _drain_pending()
        for doc, jobs in pending:
            try:
                if doc is None or not doc.IsValidObject or doc.IsLinked:
                    continue
            except Exception:
                continue
            if not jobs:
                continue
            _apply_jobs(doc, jobs)


def _apply_jobs(doc, jobs):
    global _writing_ids
    to_write = []
    for job in jobs:
        driver = job.get(u"driver")
        delta = job.get(u"delta")
        lote_ids = set(job.get(u"lote_ids") or [])
        if driver is None or delta is None or not lote_ids:
            continue
        if not delta_is_significant(delta):
            continue
        tags = []
        for tid in lote_ids:
            try:
                el = doc.GetElement(ElementId(int(tid)))
            except Exception:
                el = None
            if el is not None and isinstance(el, IndependentTag):
                tags.append(el)
        if len(tags) < 2:
            # Solo actualizar caché del driver
            try:
                el = doc.GetElement(ElementId(int(driver)))
                _cache_set(doc, driver, _head(el))
            except Exception:
                pass
            continue
        to_write.append((driver, delta, tags, lote_ids))

    if not to_write:
        return

    write_ints = set()
    for driver, _d, tags, lote_ids in to_write:
        for tid in lote_ids:
            if tid != driver:
                write_ints.add(int(tid))
    _writing_ids |= write_ints

    t = Transaction(doc, _TXN)
    try:
        t.Start()
    except Exception:
        _writing_ids -= write_ints
        return
    try:
        for driver, delta, tags, lote_ids in to_write:
            move_tag_lote_by_delta(
                doc, tags, delta, skip_ids=set([int(driver)]),
            )
            # Actualizar caché de todo el lote tras el movimiento
            for tag in tags:
                k = _eid_int(tag.Id)
                if k is None:
                    continue
                _cache_set(doc, k, _head(tag))
        t.Commit()
    except Exception:
        try:
            if t.HasStarted():
                t.RollBack()
        except Exception:
            pass
    finally:
        _writing_ids -= write_ints


def _change_type_tag_move():
    """Geometry + Any: TagHeadPosition a veces llega solo como Any."""
    try:
        from Autodesk.Revit.DB import ChangeType
        return ChangeType.Concatenate(
            Element.GetChangeTypeGeometry(),
            Element.GetChangeTypeAny(),
        )
    except Exception:
        try:
            return Element.GetChangeTypeGeometry()
        except Exception:
            try:
                return Element.GetChangeTypeAny()
            except Exception:
                return None


class TagLoteCoMoveUpdater(IUpdater):
    def __init__(self, addin_id):
        self._updater_id = UpdaterId(addin_id, UPDATER_GUID)

    def GetUpdaterId(self):
        return self._updater_id

    def GetUpdaterName(self):
        return u"Arainco: Co-mover lote etiquetas muro/malla"

    def GetAdditionalInformation(self):
        return (
            u"Si se mueve una etiqueta de muro o malla del lote, desplaza las "
            u"hermanas el mismo delta (vista activa / OwnerView)."
        )

    def GetChangePriority(self):
        return ChangePriority.Annotations

    def Execute(self, data):
        doc = data.GetDocument()
        if doc is None or doc.IsLinked:
            return
        try:
            modified = list(data.GetModifiedElementIds() or [])
        except Exception:
            modified = []
        if not modified:
            return

        modified_ints = set()
        for eid in modified:
            k = _eid_int(eid)
            if k is not None:
                modified_ints.add(k)

        for eid in modified:
            k = _eid_int(eid)
            if k is None:
                continue
            if k in _writing_ids:
                # Actualizar caché tras nuestro propio write
                try:
                    el = doc.GetElement(eid)
                    _cache_set(doc, k, _head(el))
                except Exception:
                    pass
                continue
            try:
                tag = doc.GetElement(eid)
            except Exception:
                continue
            if not isinstance(tag, IndependentTag):
                continue
            if not is_lote_relevant_tag(doc, tag):
                continue

            new_h = _head(tag)
            if new_h is None:
                continue
            old_h = _cache_get(doc, k)
            if old_h is None:
                # Primera vez: sembrar caché del lote completo (sin mover).
                _cache_set(doc, k, new_h)
                try:
                    for t in resolve_tag_lote(doc, tag):
                        tid = _eid_int(t.Id)
                        if tid is not None and _cache_get(doc, tid) is None:
                            _cache_set(doc, tid, _head(t))
                except Exception:
                    pass
                continue

            delta = _xyz_delta(old_h, new_h)
            if not delta_is_significant(delta):
                _cache_set(doc, k, new_h)
                continue

            lote = resolve_tag_lote(doc, tag)
            if len(lote) < 2:
                _cache_set(doc, k, new_h)
                continue

            lote_ids = set()
            for t in lote:
                tid = _eid_int(t.Id)
                if tid is not None:
                    lote_ids.add(tid)

            # Todo el lote ya vino en modified → Revit movió el bloque; solo caché
            if lote_ids and lote_ids.issubset(modified_ints):
                for t in lote:
                    tid = _eid_int(t.Id)
                    if tid is not None:
                        _cache_set(doc, tid, _head(t))
                continue

            # Solo el driver (o un subconjunto) se movió → co-mover hermanas
            _cache_set(doc, k, new_h)  # driver ya en posición nueva
            _enqueue_job(
                doc,
                {
                    u"driver": k,
                    u"delta": delta,
                    u"lote_ids": lote_ids,
                },
            )


def _addin_id_pyrevit_or_none():
    try:
        from pyrevit import HOST_APP
        return HOST_APP.addin_id
    except Exception:
        return None


def register_tag_lote_comove_updater(addin_id, doc=None):
    """Registra el updater. ``doc`` opcional para trigger acotado al documento."""
    _ensure_event()
    updater = TagLoteCoMoveUpdater(addin_id)
    uid = updater.GetUpdaterId()
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        try:
            UpdaterRegistry.UnregisterUpdater(uid)
        except Exception:
            pass
    UpdaterRegistry.RegisterUpdater(updater)
    flt = ElementClassFilter(IndependentTag)
    change_type = _change_type_tag_move()
    if change_type is None:
        return False
    if doc is None:
        UpdaterRegistry.AddTrigger(uid, flt, change_type)
    else:
        UpdaterRegistry.AddTrigger(uid, doc, flt, change_type)
    return True


def unregister_tag_lote_comove_updater(addin_id, doc=None):
    uid = UpdaterId(addin_id, UPDATER_GUID)
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        try:
            UpdaterRegistry.UnregisterUpdater(uid)
        except Exception:
            pass
    if doc is not None:
        _cache_clear_doc(doc)
    global _pending_by_doc, _writing_ids
    _pending_by_doc = {}
    _writing_ids = set()


def is_tag_lote_comove_registered(addin_id=None):
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
        if addin_id is None:
            return None
    uid = UpdaterId(addin_id, UPDATER_GUID)
    return UpdaterRegistry.IsUpdaterRegistered(uid)


def toggle_tag_lote_comove(addin_id=None, doc=None):
    """
    Alterna registro. Retorna ``(registered_now, message)``.
    """
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
    if addin_id is None:
        return None, u"No hay AddInId (ejecutar desde pyRevit)."
    if is_tag_lote_comove_registered(addin_id):
        unregister_tag_lote_comove_updater(addin_id, doc=doc)
        return False, (
            u"DMU co-mover lote etiquetas muro/malla: DESACTIVADO."
        )
    ok = register_tag_lote_comove_updater(addin_id, doc=doc)
    if not ok:
        return False, u"No se pudo registrar el updater (ChangeType)."
    return True, (
        u"DMU co-mover lote etiquetas muro/malla: ACTIVADO.\n"
        u"Mueve una etiqueta (muro o malla); las del mismo muro en la vista "
        u"siguen el mismo desplazamiento."
    )


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
    _on, msg = toggle_tag_lote_comove(addin_id, doc=doc)
    title = u"Arainco: Lote etiquetas muro/malla (DMU)"
    try:
        TaskDialog.Show(title, msg or u"?")
    except Exception:
        print(msg)
