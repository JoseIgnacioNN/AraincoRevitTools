# -*- coding: utf-8 -*-
"""
DMU: al cambiar el largo de un Structural Rebar, evalúa el umbral comercial de
12 m (12000 mm):

- Si supera 12 m: colorea de rojo la barra en la vista activa.
- Si deja de superar 12 m: quita el override rojo de largo exceso en todas las
  vistas aplicables del documento (reset ``OverrideGraphicSettings``), solo
  cuando el override coincide con el patrón rojo de esta herramienta.

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
    View,
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


def _color_es_rojo_largo_exceso(color):
    try:
        if color is None:
            return False
        if hasattr(color, "IsValid") and not color.IsValid:
            return False
        return (
            int(color.Red) == 255
            and int(color.Green) == 0
            and int(color.Blue) == 0
        )
    except Exception:
        return False


def _override_es_largo_exceso_rojo(view, element_id):
    """
    True si la vista tiene el patrón rojo de largo >12 m (línea roja peso 6
    y/o relleno sólido rojo), coherente con Armado Muros y este DMU.
    """
    if view is None or element_id is None:
        return False
    try:
        ogs = view.GetElementOverrides(element_id)
    except Exception:
        return False
    if ogs is None:
        return False
    try:
        if ogs.IsValidObject is False:
            return False
    except Exception:
        pass

    rojo_linea = _color_es_rojo_largo_exceso(ogs.ProjectionLineColor)
    if not rojo_linea:
        try:
            rojo_linea = _color_es_rojo_largo_exceso(ogs.CutLineColor)
        except Exception:
            rojo_linea = False
    if not rojo_linea:
        return False

    try:
        if int(ogs.ProjectionLineWeight) == 6:
            return True
    except Exception:
        pass

    try:
        if _color_es_rojo_largo_exceso(ogs.SurfaceForegroundPatternColor):
            return True
    except Exception:
        pass
    try:
        if _color_es_rojo_largo_exceso(ogs.SurfaceBackgroundPatternColor):
            return True
    except Exception:
        pass

    return False


def _iter_vistas_aplicables(doc):
    if doc is None:
        return
    try:
        views = FilteredElementCollector(doc).OfClass(View).ToElements()
    except Exception:
        views = []
    for view in views or []:
        if _vista_aplicable(view):
            yield view


def _reset_largo_exceso_overrides_en_documento(doc, rebar_id):
    """Quita override rojo de largo >12 m en todas las vistas aplicables."""
    if doc is None or rebar_id is None:
        return 0
    n_limpio = 0
    for view in _iter_vistas_aplicables(doc):
        if not _override_es_largo_exceso_rojo(view, rebar_id):
            continue
        if _limpiar_override(view, rebar_id):
            n_limpio += 1
            ck = _color_key(doc, view, rebar_id)
            if ck is not None:
                _colored_keys.discard(ck)
    return n_limpio


def _color_key(doc, view, rebar_id):
    try:
        return (id(doc), int(view.Id.IntegerValue), int(rebar_id.IntegerValue))
    except Exception:
        return None


def _aplicar_colores_en_vista(doc, view, rebar_ints):
    """
    Colorea de rojo barras >12 m en la vista activa (si aplica); quita el
    override rojo de largo exceso en todo el documento si la barra ya no
    supera 12 m. Devuelve (n_rojo, n_limpiados, n_revisados).
    """
    if doc is None or not rebar_ints:
        return 0, 0, 0

    vista_pinta = view if _vista_aplicable(view) else None
    solid_id = _solid_fill_pattern_id(doc) if vista_pinta is not None else None
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

        ck = _color_key(doc, vista_pinta, eid) if vista_pinta is not None else None
        exceso = L_mm is not None and float(L_mm) > lim + 1e-6

        if exceso:
            if vista_pinta is not None and _aplicar_override_rojo(
                vista_pinta, eid, solid_id,
            ):
                n_rojo += 1
                if ck is not None:
                    _colored_keys.add(ck)
        else:
            n_limpio += _reset_largo_exceso_overrides_en_documento(doc, eid)

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

        if active_doc is None:
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
            u"Al cambiar el largo de un Rebar: colorea de rojo en la vista activa "
            u"si supera 12 m (12000 mm) y quita el override rojo de largo exceso "
            u"en todas las vistas cuando vuelve a medir 12 m o menos."
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


def toggle_rebar_largo_exceso_color_dmu(addin_id=None, doc=None):
    """Alterna registro del DMU. Retorna ``(registered_now, message)``."""
    if addin_id is None:
        addin_id = _addin_id_pyrevit_or_none()
    if addin_id is None:
        return None, u"No hay AddInId (ejecutar desde pyRevit)."
    if is_rebar_largo_exceso_color_dmu_registered(addin_id):
        unregister_rebar_largo_exceso_color_updater(addin_id)
        return False, (
            u"DMU color / reset barras >12 m: DESACTIVADO.\n"
            u"No se colorearán ni despintarán barras automáticamente."
        )
    register_rebar_largo_exceso_color_updater(addin_id, doc=doc)
    return True, (
        u"DMU color / reset barras >12 m: ACTIVADO.\n"
        u"Colorea en rojo en la vista activa si supera 12 m y quita el override "
        u"rojo en todas las vistas cuando el largo baja a 12 m o menos."
    )


def run(__revit__):
    """Entrada pushbutton: alternar registro del DMU."""
    from Autodesk.Revit.UI import TaskDialog

    addin_id = _addin_id_pyrevit_or_none()
    doc = None
    try:
        uidoc = __revit__.ActiveUIDocument
        if uidoc is not None:
            doc = uidoc.Document
    except Exception:
        doc = None
    _on, msg = toggle_rebar_largo_exceso_color_dmu(addin_id, doc=doc)
    title = u"Arainco: Color / reset barras >12 m (DMU)"
    try:
        TaskDialog.Show(title, msg or u"?")
    except Exception:
        print(msg)
