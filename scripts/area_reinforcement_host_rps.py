# -*- coding: utf-8 -*-
"""
RPS / pyRevit — Revit 2024+: seleccionar Area Reinforcement, obtener su host y
escribir el nombre del nivel del host en el parámetro "Closest Level".

Flujo API:
  1. Tomar Area Reinforcement de la selección actual o pedir pick interactivo.
  2. Leer host con AreaReinforcement.GetHostId().
  3. Obtener el nivel del host y escribir su nombre en "Closest Level".

Uso RPS: File > Run script (no pegar línea a línea en la consola interactiva).

Configuración (editar antes de ejecutar):
  USE_CURRENT_SELECTION   True  → usar selección previa si hay Area Reinforcement.
  PICK_IF_NO_SELECTION    True  → si no hay selección válida, abrir PickObjects.
  WRITE_CLOSEST_LEVEL     True  → escribir nombre de nivel en "Closest Level".
  SELECT_HOST_AFTER       True  → dejar seleccionado el host en Revit al finalizar.
  PROMPT                  Texto en la barra de estado durante el pick.

Tras ejecutar, el último resultado queda en AREA_REIN_HOST_RESULT (lista de dicts).
"""

from __future__ import print_function

import os
import sys

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Collections.Generic import List
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB.Structure import AreaReinforcement
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from area_rein_closest_level import (
    DIALOG_TITLE,
    apply_closest_level_to_targets,
    build_summary,
    format_result_line,
    set_element_selection,
)

# --- Configuración ---
USE_CURRENT_SELECTION = True
PICK_IF_NO_SELECTION = True
WRITE_CLOSEST_LEVEL = True
SELECT_HOST_AFTER = True
PROMPT = (
    u"Selecciona uno o más Area Reinforcement. "
    u"Termina con Finish (barra de opciones) o Esc para cancelar."
)

AREA_REIN_HOST_RESULT = []


class AreaReinforcementOnlyFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem, AreaReinforcement)
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return False


def _get_doc_uidoc():
    try:
        return doc, uidoc
    except NameError:
        u = __revit__.ActiveUIDocument
        return u.Document, u


def _element_id_int(eid):
    if eid is None:
        return None
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return None


def _hide_rps_window_if_available():
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


def _area_reinforcements_from_selection(document, uidoc_):
    out = []
    seen = set()
    for eid in uidoc_.Selection.GetElementIds():
        el = document.GetElement(eid)
        if el is None or not isinstance(el, AreaReinforcement):
            continue
        key = _element_id_int(el.Id)
        if key is None or key in seen:
            continue
        seen.add(key)
        out.append(el)
    return out


def _pick_area_reinforcements(document, uidoc_):
    restore = _hide_rps_window_if_available()
    try:
        refs = list(
            uidoc_.Selection.PickObjects(
                ObjectType.Element,
                AreaReinforcementOnlyFilter(),
                PROMPT,
            )
        )
    except OperationCanceledException:
        return []
    except Exception as ex:
        TaskDialog.Show(
            DIALOG_TITLE,
            u"Selección cancelada o error:\n{}".format(ex),
        )
        return []
    finally:
        restore()

    out = []
    seen = set()
    for ref in refs:
        el = document.GetElement(ref.ElementId)
        if el is None or not isinstance(el, AreaReinforcement):
            continue
        key = _element_id_int(el.Id)
        if key is None or key in seen:
            continue
        seen.add(key)
        out.append(el)
    return out


def main():
    global AREA_REIN_HOST_RESULT

    document, uidoc_ = _get_doc_uidoc()
    if document is None or uidoc_ is None:
        msg = u"No hay documento activo."
        print(u"Error:", msg)
        try:
            TaskDialog.Show(DIALOG_TITLE, msg)
        except Exception:
            pass
        return

    targets = []
    if USE_CURRENT_SELECTION:
        targets = _area_reinforcements_from_selection(document, uidoc_)

    if not targets and PICK_IF_NO_SELECTION:
        targets = _pick_area_reinforcements(document, uidoc_)

    if not targets:
        msg = (
            u"No hay Area Reinforcement seleccionados. "
            u"Selecciona refuerzos de área o activa PICK_IF_NO_SELECTION."
        )
        print(u"Error:", msg)
        try:
            TaskDialog.Show(DIALOG_TITLE, msg)
        except Exception:
            pass
        return

    try:
        results, updated, write_failed, missing_host, missing_level = apply_closest_level_to_targets(
            document,
            targets,
            write_closest_level=WRITE_CLOSEST_LEVEL,
        )
    except Exception as ex:
        msg = u"No se pudo escribir Closest Level: {}".format(str(ex))
        print(u"Error:", msg)
        try:
            TaskDialog.Show(DIALOG_TITLE, msg)
        except Exception:
            pass
        return

    AREA_REIN_HOST_RESULT = results
    lines = [format_result_line(info) for info in results]

    print(u"\n=== Host de Area Reinforcement ===")
    for line in lines:
        print(line)

    summary = build_summary(results, updated, write_failed, missing_host, missing_level)
    if SELECT_HOST_AFTER:
        host_ids = [
            ElementId(info.get(u"host_id"))
            for info in results
            if info.get(u"host_id") is not None
        ]
        if host_ids:
            set_element_selection(uidoc_, host_ids)
            summary += u"\nSelección actualizada al/los host(s)."

    detail = u"\n".join(lines)
    if len(detail) > 1800:
        detail = detail[:1800] + u"\n… (ver consola RPS para el listado completo)"

    try:
        TaskDialog.Show(DIALOG_TITLE, summary + u"\n\n" + detail)
    except Exception:
        pass


main()
