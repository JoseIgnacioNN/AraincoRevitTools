# -*- coding: utf-8 -*-
"""
Selector para enfierrado en pasada/shaft:

  1) Primera selección: el elemento que hosteará las barras (p. ej. la losa).
  2) Siguientes: caras de ese mismo elemento que definen la posición del refuerzo
     alrededor del hueco (solo caras del host).

Flujo recomendado en pyRevit: botón «Refuerzo borde losa» (scripts/barras_bordes_losa_gancho_empotramiento):
primero el formulario BIMTools; host y caras se eligen con el botón de la ventana (ExternalEvent).

Para solo seleccionar y guardar: run_pyrevit en este módulo (RPS / pruebas).

Tras selección exitosa quedan:
  HOST_CARAS_SELECCION, REFERENCIAS_CARAS_SELECCION

Revit 2024–2026 | pyRevit (IronPython 2.7). ElementId: .Value (2026+) o .IntegerValue.
"""

import sys
import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    Face,
    Options,
    PlanarFace,
    ViewDetailLevel,
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType


def element_id_to_int(element_id):
    """Entero estable para ElementId (2026+: Value; 2024–25: IntegerValue)."""
    if element_id is None:
        return None
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


HOST_CARAS_SELECCION = None
REFERENCIAS_CARAS_SELECCION = None


class _FiltroCarasDelHost(ISelectionFilter):
    def __init__(self, host_id):
        self._host_id = host_id

    def AllowElement(self, element):
        if element is None:
            return False
        return element.Id == self._host_id

    def AllowReference(self, reference, position):
        if reference is None:
            return False
        return reference.ElementId == self._host_id


def _tipo_cara(obj):
    if obj is None:
        return u"(null)"
    if isinstance(obj, PlanarFace):
        return u"PlanarFace"
    if isinstance(obj, Face):
        return type(obj).__name__
    return type(obj).__name__


def _resolver_geometria_host(host, reference):
    if host is None or reference is None:
        return None
    opts = Options()
    opts.ComputeReferences = True
    opts.DetailLevel = ViewDetailLevel.Fine
    try:
        return host.GetGeometryObjectFromReference(reference)
    except Exception:
        return None


def _referencias_unicas(document, refs):
    from Autodesk.Revit.DB import ElementId as _EId

    _inv = element_id_to_int(_EId.InvalidElementId)

    if not refs:
        return []
    seen = set()
    out = []
    for r in refs:
        key = None
        try:
            key = r.ConvertToStableRepresentation(document)
        except Exception:
            try:
                lid = getattr(r, "LinkedElementId", None)
                lv = element_id_to_int(lid) if lid is not None else _inv
                key = (element_id_to_int(r.ElementId), lv)
            except Exception:
                key = id(r)
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def _asignar_resultado_modulo(host, refs_tuple):
    m = sys.modules[__name__]
    m.HOST_CARAS_SELECCION = host
    m.REFERENCIAS_CARAS_SELECCION = refs_tuple


def seleccionar_host_y_caras(uidoc, doc, guardar_en_modulo=True, mostrar_errores=True):
    """
    1) Pick elemento host. 2) Pick caras (solo de ese host).

    Returns:
        (host, tuple_refs) si OK; (None, None) si cancelación o error.
    """
    if uidoc is None or doc is None:
        if mostrar_errores:
            TaskDialog.Show(u"Selección de caras", u"No hay documento activo.")
        return None, None

    try:
        ref_host = uidoc.Selection.PickObject(
            ObjectType.Element,
            u"1/2 — Selecciona el elemento host de las barras (p. ej. la losa).",
        )
    except OperationCanceledException:
        return None, None
    except Exception:
        return None, None

    if ref_host is None:
        return None, None

    host = doc.GetElement(ref_host.ElementId)
    if host is None:
        if mostrar_errores:
            TaskDialog.Show(u"Selección de caras", u"No se pudo obtener el elemento host.")
        return None, None

    filtro = _FiltroCarasDelHost(host.Id)
    try:
        refs_caras = list(
            uidoc.Selection.PickObjects(
                ObjectType.Face,
                filtro,
                u"2/2 — Caras del refuerzo alrededor de la pasada (solo este elemento). "
                u"Finaliza con clic derecho o Enter.",
            )
        )
    except OperationCanceledException:
        return None, None
    except Exception as ex:
        if mostrar_errores:
            TaskDialog.Show(
                u"Selección de caras",
                u"Error al seleccionar caras:\n{}".format(ex),
            )
        return None, None

    refs_todas = _referencias_unicas(doc, refs_caras)
    if not refs_todas:
        if mostrar_errores:
            TaskDialog.Show(
                u"Selección de caras",
                u"Selecciona al menos una cara del host para la posición de las barras.",
            )
        return None, None

    refs_tuple = tuple(refs_todas)
    if guardar_en_modulo:
        _asignar_resultado_modulo(host, refs_tuple)

    return host, refs_tuple


def run_pyrevit(revit):
    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(u"Selección de caras", u"No hay documento activo.")
        return
    doc = uidoc.Document
    host, refs_todas = seleccionar_host_y_caras(uidoc, doc, guardar_en_modulo=True, mostrar_errores=True)
    if host is None:
        return

    lineas = []
    for i, r in enumerate(refs_todas):
        go = _resolver_geometria_host(host, r)
        tipo = _tipo_cara(go)
        lineas.append(u"  [{0}] Reference — tipo: {1}".format(i + 1, tipo))

    nombre = u""
    try:
        nombre = host.Name or u""
    except Exception:
        pass
    categoria = u""
    try:
        if host.Category is not None:
            categoria = host.Category.Name or u""
    except Exception:
        pass

    msg = (
        u"Host (elemento barras) — Id {0}\n".format(element_id_to_int(host.Id))
        + (u"Nombre: {0}\n".format(nombre) if nombre else u"")
        + (u"Categoría: {0}\n".format(categoria) if categoria else u"")
        + u"\nCaras seleccionadas: {0}\n".format(len(refs_todas))
        + u"\n".join(lineas)
        + u"\n\nHost y referencias guardados en el módulo:\n"
        u"  seleccion_caras_elemento.HOST_CARAS_SELECCION\n"
        u"  seleccion_caras_elemento.REFERENCIAS_CARAS_SELECCION"
    )
    TaskDialog.Show(u"Selección de caras — resultado", msg)


def run(document, uidocument):
    class _RevitShim(object):
        def __init__(self, d, u):
            self.ActiveUIDocument = u

    run_pyrevit(_RevitShim(document, uidocument))
