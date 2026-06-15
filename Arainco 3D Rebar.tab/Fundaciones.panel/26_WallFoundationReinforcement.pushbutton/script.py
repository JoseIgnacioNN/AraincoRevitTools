# -*- coding: utf-8 -*-
"""Wall Foundation Reinforcement — entrada pyRevit."""

__title__ = u"Fundacion\nCorrida"

import os
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_TOOL_DIALOG_TITLE = u"Wall Foundation Reinforcement"


def _find_module(start_dir):
    cursor = start_dir
    for _ in range(10):
        candidate = os.path.join(cursor, "scripts", "enfierrado_wall_foundation.py")
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
        u"No se encontró scripts/enfierrado_wall_foundation.py",
    )
    raise Exception(u"No se encontró scripts/enfierrado_wall_foundation.py")

import sys

_scripts_dir = os.path.dirname(_module_path)
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)
import bimtools_paths

bimtools_paths.set_pushbutton_dir(_pushbutton_dir)

# --- Validación acceso corporativo (RECURSOS COMPARTIDOS) ---
import os as _os_ac
import sys as _sys_ac
_tab_ac = _os_ac.path.dirname(_os_ac.path.abspath(__file__))
for _iac in range(16):
    if _os_ac.path.basename(_tab_ac).endswith(u".tab"):
        break
    _parent_ac = _os_ac.path.dirname(_tab_ac)
    if _parent_ac == _tab_ac:
        _tab_ac = None
        break
    _tab_ac = _parent_ac
if _tab_ac and _tab_ac not in _sys_ac.path:
    _sys_ac.path.insert(0, _tab_ac)
import bimtools_access_bootstrap as _bimtools_access
if _bimtools_access.require_tool_access(__file__, __revit__, __title__):
    try:
        _mod = imp.load_source("enfierrado_wall_foundation", _module_path)
        _mod.run_pyrevit(__revit__)
    except Exception as ex:
        TaskDialog.Show(
            _TOOL_DIALOG_TITLE,
            u"Error ejecutando la rutina:\n{0}".format(ex),
        )
