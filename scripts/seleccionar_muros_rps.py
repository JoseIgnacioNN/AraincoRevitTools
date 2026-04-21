# -*- coding: utf-8 -*-
"""
RevitPythonShell (RPS): seleccionar muros (Wall) en el modelo activo.

Flujo API: 1) Filtro ISelectionFilter (solo Wall). 2) PickObjects en la vista activa.
3) Opcionalmente actualizar uidoc.Selection con los ElementId obtenidos.

- Revit 2021+ (API según instalación) | IronPython 2.7 típico en RPS.
- Ejecutar: File > Run script (no pegar línea a línea en la consola interactiva).

Configuración (editar antes de ejecutar):
  REPLACE_SELECTION  True  → la selección de Revit pasa a ser solo estos muros.
                     False → se añaden a la selección actual (sin duplicar por Id).
  PROMPT             Texto del mensaje en la barra de estado durante el pick.

Tras ejecutar, los ids quedan en la variable de módulo SELECTED_WALL_IDS.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Collections.Generic import List
from Autodesk.Revit.DB import ElementId, Wall
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

# --- Configuración ---
REPLACE_SELECTION = True
PROMPT = u"Selecciona uno o más muros. Termina con Finish (barra de opciones) o Esc para cancelar."

SELECTED_WALL_IDS = []


class WallOnlyFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem, Wall)
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return False


def _hide_rps_window_if_available():
    """
    Patrón RPS: ocultar __window__ antes de PickObjects para que Revit reciba el foco.
    """
    w = globals().get("__window__", None)
    if w is None:
        def _noop():
            return

        return _noop

    prev_topmost = None
    try:
        prev_topmost = bool(getattr(w, "Topmost", False))
    except Exception:
        prev_topmost = None

    try:
        w.Topmost = False
    except Exception:
        pass
    try:
        w.Hide()
    except Exception:
        pass

    def _restore():
        try:
            w.Show()
        except Exception:
            pass
        try:
            if prev_topmost is not None:
                w.Topmost = prev_topmost
        except Exception:
            pass
        try:
            w.Activate()
        except Exception:
            pass

    return _restore


def _merge_element_ids(uidoc, nuevos):
    """nuevos: iterable de ElementId. Devuelve List[ElementId] sin duplicar IntegerValue."""
    seen = set()
    out = List[ElementId]()
    for eid in uidoc.Selection.GetElementIds():
        try:
            iv = eid.IntegerValue
        except Exception:
            continue
        if iv not in seen:
            seen.add(iv)
            out.Add(eid)
    for eid in nuevos:
        try:
            iv = eid.IntegerValue
        except Exception:
            continue
        if iv not in seen:
            seen.add(iv)
            out.Add(eid)
    return out


def ejecutar(uidoc, doc):
    global SELECTED_WALL_IDS

    SELECTED_WALL_IDS = []

    if uidoc is None or doc is None:
        TaskDialog.Show(u"Seleccionar muros", u"No hay documento activo.")
        return

    restore = _hide_rps_window_if_available()
    try:
        refs = list(
            uidoc.Selection.PickObjects(
                ObjectType.Element,
                WallOnlyFilter(),
                PROMPT,
            )
        )
    except OperationCanceledException:
        return
    except Exception as ex:
        TaskDialog.Show(
            u"Seleccionar muros",
            u"Selección cancelada o error:\n{}".format(ex),
        )
        return
    finally:
        restore()

    if not refs:
        TaskDialog.Show(u"Seleccionar muros", u"No se seleccionó ningún muro.")
        return

    wall_ids = []
    for r in refs:
        elem = doc.GetElement(r.ElementId)
        if elem is not None and isinstance(elem, Wall):
            wall_ids.append(r.ElementId)

    if not wall_ids:
        TaskDialog.Show(u"Seleccionar muros", u"Ningún elemento válido es muro (Wall).")
        return

    SELECTED_WALL_IDS = list(wall_ids)

    if REPLACE_SELECTION:
        ids = List[ElementId]()
        for eid in wall_ids:
            ids.Add(eid)
        uidoc.Selection.SetElementIds(ids)
    else:
        uidoc.Selection.SetElementIds(_merge_element_ids(uidoc, wall_ids))

    msg = u"{} muro(s) seleccionado(s).".format(len(wall_ids))
    if len(refs) > len(wall_ids):
        msg += u" ({} elemento(s) ignorados: no son Wall).".format(len(refs) - len(wall_ids))

    TaskDialog.Show(u"Seleccionar muros", msg)
    print(msg)
    print(u"Ids (IntegerValue):", [e.IntegerValue for e in wall_ids])


def _main():
    try:
        uidoc = __revit__.ActiveUIDocument
        doc = uidoc.Document
    except NameError:
        TaskDialog.Show(
            u"Seleccionar muros",
            u"No está definido __revit__. Ejecuta desde RPS/pyRevit o llama ejecutar(uidoc, doc).",
        )
        return
    ejecutar(uidoc, doc)


_main()
