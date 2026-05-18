# -*- coding: utf-8 -*-
"""Entrada pyRevit delgada para Armado Columnas enterprise."""

import os
import sys
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_TOOL_DIALOG_TITLE = u"Arainco: Armado Columnas"


def _find_scripts_dir(start_dir):
    cursor = start_dir
    for _ in range(10):
        candidate = os.path.join(cursor, "scripts")
        runner = os.path.join(candidate, "column_reinforcement", "runner.py")
        if os.path.isfile(runner):
            return candidate
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_scripts_dir = _find_scripts_dir(_pushbutton_dir)

if not _scripts_dir:
    TaskDialog.Show(
        _TOOL_DIALOG_TITLE,
        u"No se encontró scripts/column_reinforcement/runner.py",
    )
    raise Exception(u"No se encontró scripts/column_reinforcement/runner.py")

if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

try:
    import bimtools_paths

    bimtools_paths.set_pushbutton_dir(_pushbutton_dir)
except Exception:
    pass

try:
    _runner_path = os.path.join(_scripts_dir, "column_reinforcement", "runner.py")
    _runner = imp.load_source("column_reinforcement.runner", _runner_path)

    _runner.run_pyrevit(__revit__)
except Exception as ex:
    TaskDialog.Show(
        _TOOL_DIALOG_TITLE,
        u"Error ejecutando la rutina:\n{0}".format(ex),
    )
