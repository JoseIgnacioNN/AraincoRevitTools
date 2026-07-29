# -*- coding: utf-8 -*-
"""
Entrada — Arainco: Coronamiento muros.

Pick muro → UI (elevación + stack) → Colocar → re-pick sin cerrar.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Wall
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

_DIALOG_TITLE = u"Arainco: Coronamiento muros"


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


def pick_wall(uidoc, allow_cancel=False):
    """PickObject de un muro. Si allow_cancel y el usuario cancela → None."""
    if uidoc is None:
        return None
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _WallSelectionFilter(),
            u"Seleccione un muro para coronamiento",
        )
    except OperationCanceledException:
        if allow_cancel:
            return None
        raise
    except Exception:
        if allow_cancel:
            return None
        raise
    if ref is None:
        return None
    el = uidoc.Document.GetElement(ref)
    if isinstance(el, Wall):
        return el
    return None


def run(uiapp):
    """Entrada pyRevit / RPS: ``run(__revit__)``."""
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
        wall = pick_wall(uidoc, allow_cancel=False)
    except OperationCanceledException:
        return
    except Exception as ex:
        try:
            msg = unicode(ex)
        except Exception:
            msg = str(ex)
        _mostrar(uiapp, u"No se pudo seleccionar el muro.", msg)
        return
    if wall is None:
        _mostrar(uiapp, u"Seleccione un muro válido.")
        return

    from coronamiento_muros_ui import show_coronamiento_window

    show_coronamiento_window(uiapp, uidoc, doc, wall)


def run_pyrevit(revit_app):
    run(revit_app)
