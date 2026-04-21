# -*- coding: utf-8 -*-
"""Enfierrado vigas — UI y selección; lógica de barras pendiente."""

import os
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_TOOL_DIALOG_TITLE = u"Enfierrado vigas"


def _find_module(start_dir):
    cursor = start_dir
    for _ in range(10):
        candidate = os.path.join(cursor, "scripts", "enfierrado_vigas.py")
        if os.path.isfile(candidate):
            return candidate
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_module_path = _find_module(_pushbutton_dir)

if not _module_path:
    TaskDialog.Show(
        _TOOL_DIALOG_TITLE,
        u"No se encontró scripts/enfierrado_vigas.py",
    )
    raise Exception(u"No se encontró scripts/enfierrado_vigas.py")

try:
    _mod = imp.load_source("enfierrado_vigas", _module_path)
    _mod.run_pyrevit(__revit__)
except Exception as ex:
    TaskDialog.Show(
        _TOOL_DIALOG_TITLE,
        u"Error ejecutando la rutina:\n{0}".format(ex),
    )
