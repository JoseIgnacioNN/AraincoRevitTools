# -*- coding: utf-8 -*-
"""Boton: borde losa gancho y empotramiento."""

import os
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_TOOL_DIALOG_TITLE = u"Refuerzo borde losa"


def _find_module(start_dir):
    cursor = start_dir
    for _ in range(10):
        candidate = os.path.join(cursor, "scripts", "barras_bordes_losa_gancho_empotramiento.py")
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
        u"No se encontro scripts/barras_bordes_losa_gancho_empotramiento.py",
    )
    raise Exception(u"No se encontro scripts/barras_bordes_losa_gancho_empotramiento.py")

import sys as _sys

_scripts_dir = os.path.dirname(_module_path)
if _scripts_dir not in _sys.path:
    _sys.path.insert(0, _scripts_dir)
import bimtools_paths

bimtools_paths.set_pushbutton_dir(_pushbutton_dir)
_sys.modules.pop("enfierrado_shaft_hashtag", None)
_sys.modules.pop("barras_bordes_losa_gancho_empotramiento", None)

try:
    _mod = imp.load_source("barras_bordes_losa_gancho_empotramiento", _module_path)
    _mod.run_pyrevit(__revit__)
except Exception as ex:
    TaskDialog.Show(
        _TOOL_DIALOG_TITLE,
        u"Error ejecutando la rutina:\n{0}".format(ex),
    )
