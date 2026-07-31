# -*- coding: utf-8 -*-
"""
Entrada — Arainco: Barras de retorno de malla.

Pick muro(s) → detectar fundación / unidos (interno) → UI → Colocar.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Wall
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

_DIALOG_TITLE = u"Arainco: Barras de retorno de malla"


class _WallSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, Wall)

    def AllowReference(self, reference, point):
        return True


def _mostrar(uiapp, instruction, content=u""):
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        show_message_dialog(
            _DIALOG_TITLE,
            instruction=instruction,
            content=content or None,
            ok_text=u"Entendido",
            hwnd_revit=revit_main_hwnd(uiapp) if uiapp else None,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        TaskDialog.Show(
            _DIALOG_TITLE,
            u"{0}\n{1}".format(instruction, content or u"").strip(),
        )
    except Exception:
        pass


def pick_walls(uidoc):
    """Multi-select de muros. Cancelar / 0 → lista vacía."""
    if uidoc is None:
        return []
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            _WallSelectionFilter(),
            u"Seleccione muro(s) para barras de retorno. Finish confirma; Esc cancela.",
        )
    except OperationCanceledException:
        return []
    except Exception:
        return []
    if not refs:
        return []
    doc = uidoc.Document
    walls = []
    seen = set()
    for ref in refs:
        if ref is None:
            continue
        try:
            el = doc.GetElement(ref)
        except Exception:
            el = None
        if not isinstance(el, Wall):
            continue
        try:
            wid = int(el.Id.IntegerValue)
        except Exception:
            continue
        if wid in seen:
            continue
        seen.add(wid)
        walls.append(el)
    return walls


def run(uiapp):
    """Entrada pyRevit: ``run(__revit__)``."""
    if uiapp is None:
        return
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        _mostrar(uiapp, u"No hay documento activo.")
        return
    doc = uidoc.Document
    if doc is None:
        _mostrar(uiapp, u"No hay documento activo.")
        return

    try:
        walls = pick_walls(uidoc)
    except OperationCanceledException:
        return
    except Exception as ex:
        try:
            msg = unicode(ex)
        except Exception:
            msg = str(ex)
        _mostrar(uiapp, u"No se pudo seleccionar el muro.", msg)
        return

    if not walls:
        return

    from barras_retorno_malla_ui import show_barras_retorno_window

    show_barras_retorno_window(uiapp, uidoc, doc, walls)


def run_pyrevit(revit_app):
    run(revit_app)
