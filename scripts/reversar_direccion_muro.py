# -*- coding: utf-8 -*-
"""
Reversar dirección de muro — invierte la LocationCurve de muros seleccionados.

Revit 2024+ | pyRevit | IronPython 3.4

Paquete portable autocontenido (``21_ReversarDireccionMuro.pushbutton/scripts/``).
Respaldo de desarrollo en ``BIMTools.extension/scripts/`` — sincronice tras editar.

Flujo:
  1. Muros preseleccionados o selección interactiva múltiple (PickObjects).
  2. Por cada muro, invierte la LocationCurve (Line o Arc acotada).
  3. Resumen de muros invertidos y omitidos.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Arc, Line, LocationCurve, Transaction, Wall
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

try:
    from Autodesk.Revit.Exceptions import OperationCanceledException
except Exception:
    OperationCanceledException = Exception

from pyrevit import forms

_TITULO = u"Arainco: Reversar dirección de muro"
_PICK_PROMPT = (
    u"Seleccione uno o más muros. "
    u"Finalizar en la cinta o Esc para cancelar."
)


class _FiltroWall(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return elem is not None and isinstance(elem, Wall)
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


def _pick_walls(uidoc):
    """Modo selección interactiva: uno o más muros."""
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            _FiltroWall(),
            _PICK_PROMPT,
        )
    except OperationCanceledException:
        return []
    except Exception:
        return []
    if not refs:
        return []

    doc = uidoc.Document
    walls = []
    vistos = set()
    for ref in refs:
        try:
            el = doc.GetElement(ref.ElementId)
        except Exception:
            continue
        if el is None or not isinstance(el, Wall):
            continue
        try:
            eid = int(el.Id.IntegerValue)
        except Exception:
            eid = None
        if eid is not None and eid in vistos:
            continue
        if eid is not None:
            vistos.add(eid)
        walls.append(el)
    return walls


def _obtener_muros(uidoc, doc):
    """Muros preseleccionados o, si no hay ninguno, abre el picker múltiple."""
    walls = []
    vistos = set()
    try:
        for eid in uidoc.Selection.GetElementIds():
            el = doc.GetElement(eid)
            if el is None or not isinstance(el, Wall):
                continue
            try:
                key = int(el.Id.IntegerValue)
            except Exception:
                key = None
            if key is not None and key in vistos:
                continue
            if key is not None:
                vistos.add(key)
            walls.append(el)
    except Exception:
        pass

    if walls:
        return walls

    return _pick_walls(uidoc)


def _reverse_wall_curve(wall):
    """
    Invierte la curva de ubicación del muro.

    Returns:
        (True, None) si se aplicó la inversión.
        (False, motivo) si no se pudo procesar.
    """
    loc_curve = wall.Location
    if not isinstance(loc_curve, LocationCurve):
        return False, u"ubicación no basada en curva"

    curve = loc_curve.Curve
    if curve is None or not curve.IsBound:
        return False, u"curva no acotada"

    start_pt = curve.GetEndPoint(0)
    end_pt = curve.GetEndPoint(1)

    if isinstance(curve, Line):
        reversed_curve = Line.CreateBound(end_pt, start_pt)
    elif isinstance(curve, Arc):
        mid_pt = curve.Evaluate(0.5, True)
        reversed_curve = Arc.Create(end_pt, start_pt, mid_pt)
    else:
        return False, u"tipo de curva no soportado"

    loc_curve.Curve = reversed_curve
    return True, None


def _ejecutar(uidoc):
    doc = uidoc.Document

    walls = _obtener_muros(uidoc, doc)
    if not walls:
        return

    success_count = 0
    skipped = []

    t = Transaction(doc, _TITULO)
    t.Start()
    try:
        for wall in walls:
            ok, reason = _reverse_wall_curve(wall)
            if ok:
                success_count += 1
            else:
                skipped.append((wall.Id.IntegerValue, reason))
        t.Commit()
    except Exception as ex:
        t.RollBack()
        forms.alert(u"Error: {0}".format(ex), title=_TITULO)
        raise

    if success_count == 0:
        forms.alert(
            u"No se pudo invertir ningún muro.\n"
            u"Solo se admiten muros con LocationCurve Line o Arc acotada.",
            title=_TITULO,
        )
        return

    msg = u"Muros invertidos: {0} de {1}.".format(success_count, len(walls))
    if skipped:
        msg += u"\n\nOmitidos ({0}):".format(len(skipped))
        for wall_id, reason in skipped[:8]:
            msg += u"\n  • Id {0}: {1}".format(wall_id, reason)
        if len(skipped) > 8:
            msg += u"\n  … y {0} más.".format(len(skipped) - 8)

    forms.alert(msg, title=_TITULO)


def run(revit):
    """Entrada pyRevit: selección e inversión de muros."""
    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        forms.alert(u"No hay documento activo.", title=_TITULO)
        return
    _ejecutar(uidoc)
