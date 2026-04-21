# -*- coding: utf-8 -*-
"""
DMU: sincroniza anotaciones ligadas a Rebar al cambiar o borrar barras:

- Detalles de empalme/traslapo (familia line-based): schema **genérico** (shaft/borde losa);
  schema **zapata de muro** (``lap_detail_link_wall_foundation_schema``: sin recolocar el tramo
  del símbolo al cambiar el armado; solo borrar detalle/cota si falta una barra); schema **vigas**
  (``lap_detail_link_vigas_schema``) con ``compute_lap_segment_endpoints_vigas`` si aplica.
- Cotas de empotramiento: DetailCurve marcador + cota (línea de cota alineada al tramo de barra).

El trabajo pesado va en ExternalEvent (no dentro de IUpdater.Execute).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    ChangePriority,
    ChangeType,
    Element,
    ElementClassFilter,
    ElementId,
    IUpdater,
    Line,
    LocationCurve,
    Transaction,
    TransactionStatus,
    UpdaterId,
    UpdaterRegistry,
)
from Autodesk.Revit.DB.Structure import Rebar
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System import Guid

from embed_anchorage_link_schema import (
    element_id_to_int,
    find_embed_anchorage_touching_rebar_ids,
)
from embed_anchorage_refresh import refresh_embed_anchorage_marker
from lap_detail_link_schema import find_lap_details_touching_rebar_ids
from lap_detail_link_vigas_schema import find_vigas_lap_details_touching_rebar_ids
from lap_detail_link_wall_foundation_schema import (
    find_wall_foundation_lap_details_touching_rebar_ids,
)
from lap_detail_overlap_geom import (
    compute_lap_segment_endpoints,
    compute_lap_segment_endpoints_vigas,
)

UPDATER_GUID = Guid("b2c3d4e5-f6a7-4890-b123-456789abcdef")

_TXN = u"BIMTools: sincronizar anotaciones armadura (DMU)"

_pending_ids_by_doc = {}
_lap_detail_event = None
_doc_changed_subscribed = False
_idling_for_doc_changed_scheduled = False


def _ensure_event():
    global _lap_detail_event
    if _lap_detail_event is None:
        _lap_detail_event = ExternalEvent.Create(_LapDetailExternalHandler())
    return _lap_detail_event


def _enqueue(doc, element_ids_ints):
    global _pending_ids_by_doc
    if doc is None or not element_ids_ints:
        return
    key = id(doc)
    new_ints = set(int(x) for x in element_ids_ints)
    if not new_ints:
        return
    if key in _pending_ids_by_doc:
        _doc_ref, existing = _pending_ids_by_doc[key]
        existing |= new_ints
    else:
        _pending_ids_by_doc[key] = (doc, new_ints)
    try:
        _ensure_event().Raise()
    except Exception as ex:
        try:
            print(u"[BIMTools] ExternalEvent.Raise falló: {0}".format(ex))
        except Exception:
            pass


def _drain_pending():
    global _pending_ids_by_doc
    out = list(_pending_ids_by_doc.values())
    _pending_ids_by_doc = {}
    return out


class _LapDetailExternalHandler(IExternalEventHandler):
    def GetName(self):
        return u"BIMTools — DMU anotaciones armadura (diferido)"

    def Execute(self, uiapp):
        _subscribe_document_changed_if_needed(uiapp)
        pending = _drain_pending()
        for doc, id_ints in pending:
            try:
                if doc is None or not doc.IsValidObject or doc.IsLinked:
                    continue
            except Exception:
                continue
            if not id_ints:
                continue
            lap_pairs = find_lap_details_touching_rebar_ids(doc, id_ints)
            wf_pairs = find_wall_foundation_lap_details_touching_rebar_ids(doc, id_ints)
            vigas_pairs = find_vigas_lap_details_touching_rebar_ids(doc, id_ints)
            embed_pairs = find_embed_anchorage_touching_rebar_ids(doc, id_ints)
            if not lap_pairs and not wf_pairs and not vigas_pairs and not embed_pairs:
                # Sin anotaciones BIMTools ligadas a esos ids: no hacer nada. No usar print()
                # aquí: pyRevit abre ventana con la salida y, con muchos ids (p. ej. tras otras
                # herramientas), resulta molesto.
                continue
            # Asegurar geometría actual del modelo antes de leer curvas de Rebar (evita marcador
            # recreado en la posición anterior por datos obsoletos).
            try:
                doc.Regenerate()
            except Exception:
                pass
            txn = Transaction(doc, _TXN)
            try:
                if txn.Start() != TransactionStatus.Started:
                    continue
            except Exception:
                continue
            try:
                if wf_pairs:
                    self._process_pairs(doc, wf_pairs, update_curve=False)
                if lap_pairs:
                    self._process_pairs(doc, lap_pairs)
                if vigas_pairs:
                    self._process_pairs(
                        doc,
                        vigas_pairs,
                        lap_endpoint_fn=compute_lap_segment_endpoints_vigas,
                    )
                if embed_pairs:
                    self._process_embed_anchorage(doc, embed_pairs)
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

    def _process_pairs(self, doc, pairs, lap_endpoint_fn=None, update_curve=True):
        """
        ``lap_endpoint_fn(ra, rb, view) -> (p0,p1)``; por defecto ``compute_lap_segment_endpoints``.
        Los detail de **vigas** pasan ``compute_lap_segment_endpoints_vigas`` (personalizable).

        Si ``update_curve`` es False (empalmes zapata de muro), solo se valida que existan las
        dos barras; no se modifica ``Location.Curve`` para conservar la colocación original.
        """
        if lap_endpoint_fn is None:
            lap_endpoint_fn = compute_lap_segment_endpoints
        seen_detail = set()
        for inst, link in pairs:
            iid = element_id_to_int(inst.Id)
            if iid is None:
                continue
            if iid in seen_detail:
                continue
            seen_detail.add(iid)
            try:
                ra = doc.GetElement(link["ra"])
                rb = doc.GetElement(link["rb"])
            except Exception:
                ra, rb = None, None
            dim_id = link.get("dim")

            if ra is None or rb is None:
                self._delete_detail_and_dim(doc, inst, dim_id)
                continue
            if not isinstance(ra, Rebar) or not isinstance(rb, Rebar):
                self._delete_detail_and_dim(doc, inst, dim_id)
                continue

            if not update_curve:
                continue

            try:
                view = doc.GetElement(inst.OwnerViewId)
            except Exception:
                view = None
            if view is None:
                continue

            p0, p1 = lap_endpoint_fn(ra, rb, view)
            if p0 is None or p1 is None:
                self._delete_detail_and_dim(doc, inst, dim_id)
                continue

            try:
                ln = Line.CreateBound(p0, p1)
            except Exception:
                self._delete_detail_and_dim(doc, inst, dim_id)
                continue

            loc = getattr(inst, "Location", None)
            if not isinstance(loc, LocationCurve):
                continue
            try:
                loc.Curve = ln
            except Exception:
                try:
                    self._delete_detail_and_dim(doc, inst, dim_id)
                except Exception:
                    pass

    def _process_embed_anchorage(self, doc, pairs):
        seen = set()
        for marker, link in pairs:
            mid = element_id_to_int(marker.Id)
            if mid is None:
                continue
            if mid in seen:
                continue
            seen.add(mid)
            try:
                refresh_embed_anchorage_marker(doc, marker, link)
            except Exception as ex:
                try:
                    print(u"[BIMTools DMU empotramiento] {0}".format(ex))
                except Exception:
                    pass

    def _delete_detail_and_dim(self, doc, detail_inst, dim_id):
        if dim_id is not None and dim_id != ElementId.InvalidElementId:
            try:
                de = doc.GetElement(dim_id)
                if de is not None:
                    doc.Delete(dim_id)
            except Exception:
                pass
        try:
            if detail_inst is not None:
                doc.Delete(detail_inst.Id)
        except Exception:
            pass


def _is_rebar_category(el, bic):
    try:
        if el is None or el.Category is None:
            return False
        cid = element_id_to_int(el.Category.Id)
        if cid is None:
            return False
        return cid == int(bic.OST_Rebar)
    except Exception:
        return False


def _subscribe_document_changed_if_needed(uiapp):
    """Respaldo frente a IUpdater: Application.DocumentChanged suele disparar en 2024."""
    global _doc_changed_subscribed
    if _doc_changed_subscribed or uiapp is None:
        return
    try:
        uiapp.Application.DocumentChanged += _on_application_document_changed
        _doc_changed_subscribed = True
    except Exception:
        # No usar print() aquí: en carga de extensión pyRevit abre la ventana de salida.
        pass


def _schedule_document_changed_subscription():
    """Primera oportunidad: Idling tras cargar la extensión (UIApplication disponible)."""
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
    from Autodesk.Revit.DB import BuiltInCategory

    bic = BuiltInCategory
    touch = set()
    try:
        for eid in args.GetModifiedElementIds():
            try:
                el = doc.GetElement(eid)
            except Exception:
                el = None
            if el is not None and _is_rebar_category(el, bic):
                ei = element_id_to_int(eid)
                if ei is not None:
                    touch.add(ei)
    except Exception:
        pass
    # No incluir GetDeletedElementIds() aquí: incluye cualquier borrado (Spot Elevation, muros…).
    # El IUpdater LapDetailLinkUpdater ya tiene trigger de borrado sobre Rebar; bastaba con eso
    # para barras y evita encolar ids que no son armadura (p. ej. al borrar cotas de nivel).
    if touch:
        _enqueue(doc, touch)


class LapDetailLinkUpdater(IUpdater):
    def __init__(self, addin_id):
        self._Element = Element
        self._updater_id = UpdaterId(addin_id, UPDATER_GUID)

    def GetUpdaterId(self):
        return self._updater_id

    def GetUpdaterName(self):
        return u"BIMTools — Sincronizar anotaciones de armadura (empalme / empotramiento)"

    def GetAdditionalInformation(self):
        return (
            u"Actualiza detalles de traslapo y cotas de empotramiento (marcador + cota) "
            u"ligados a Rebar cuando cambia la geometría o se elimina una barra."
        )

    def GetChangePriority(self):
        return ChangePriority.Rebar

    def Execute(self, data):
        doc = data.GetDocument()
        if doc is None or doc.IsLinked:
            return
        from Autodesk.Revit.DB import BuiltInCategory

        bic = BuiltInCategory
        touch = set()
        try:
            for eid in data.GetModifiedElementIds():
                try:
                    el = doc.GetElement(eid)
                except Exception:
                    el = None
                if el is not None and _is_rebar_category(el, bic):
                    ei = element_id_to_int(eid)
                    if ei is not None:
                        touch.add(ei)
        except Exception:
            pass
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


def register_lap_detail_link_updater(addin_id, doc=None):
    _ensure_event()
    updater = LapDetailLinkUpdater(addin_id)
    uid = updater.GetUpdaterId()
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        try:
            UpdaterRegistry.UnregisterUpdater(uid)
        except Exception:
            pass
    UpdaterRegistry.RegisterUpdater(updater)
    flt = ElementClassFilter(Rebar)
    # Cambios en geometría / parámetros (incl. largo por estirar o por forma)
    try:
        ct_mod = _rebar_modification_change_type()
        if doc is None:
            UpdaterRegistry.AddTrigger(uid, flt, ct_mod)
        else:
            UpdaterRegistry.AddTrigger(uid, doc, flt, ct_mod)
    except Exception:
        pass
    # Borrado de barras
    try:
        ct_del = Element.GetChangeTypeElementDeletion()
        if doc is None:
            UpdaterRegistry.AddTrigger(uid, flt, ct_del)
        else:
            UpdaterRegistry.AddTrigger(uid, doc, flt, ct_del)
    except Exception:
        pass
    _schedule_document_changed_subscription()


def _rebar_modification_change_type():
    """
    Unión geometría + cualquier otro cambio: al editar largo de armado a veces Revit
    notifica sobre todo como cambio de geometría; ``GetChangeTypeAny()`` solo puede
    no bastar según versión / tipo de edición.
    """
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


def unregister_lap_detail_link_updater(addin_id):
    uid = UpdaterId(addin_id, UPDATER_GUID)
    if UpdaterRegistry.IsUpdaterRegistered(uid):
        UpdaterRegistry.UnregisterUpdater(uid)
