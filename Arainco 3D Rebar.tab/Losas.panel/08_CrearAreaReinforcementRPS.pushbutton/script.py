# -*- coding: utf-8 -*-
"""Crear Area Reinforcement — entrada pyRevit; lógica en scripts/area_reinforcement_losa.py."""

__title__ = "Crear Area\nReinf. RPS"
__author__ = "BIMTools"
__doc__ = (
    "Abre la interfaz de Area Reinforcement Losa para crear mallas en losas "
    "y losas de cimentación."
)

import os
import sys
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_TOOL_DIALOG_TITLE = u"Arainco: Malla en Losa"
_MAIN_MODULE = "area_reinforcement_losa.py"
_MAIN_MODULE_ID = "area_reinforcement_losa"


def _find_module(start_dir):
    cursor = start_dir
    for _ in range(10):
        candidate = os.path.join(cursor, "scripts", _MAIN_MODULE)
        if os.path.isfile(candidate):
            return candidate
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


def _pin_scripts_first(scripts_dir):
    """Deja scripts/ de la extensión delante de copias locales de otros pushbuttons."""
    if not scripts_dir:
        return
    try:
        while scripts_dir in sys.path:
            sys.path.remove(scripts_dir)
    except Exception:
        pass
    sys.path.insert(0, scripts_dir)


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_module_path = _find_module(_pushbutton_dir)

if not _module_path:
    TaskDialog.Show(
        _TOOL_DIALOG_TITLE,
        u"No se encontró scripts/{0}".format(_MAIN_MODULE),
    )
    raise Exception(u"No se encontró scripts/{0}".format(_MAIN_MODULE))

_scripts_dir = os.path.dirname(_module_path)
_pin_scripts_first(_scripts_dir)
import bimtools_paths

bimtools_paths.set_pushbutton_dir(_pushbutton_dir)

# --- Validacion acceso corporativo (prod: bootstrap junto al boton) ---
# === BEGIN BIZARDS_PROD_PORTABLE_BOOTSTRAP (prod_builder) ===
import os as _os_ac
import sys as _sys_ac

_pb_ac = _os_ac.path.dirname(_os_ac.path.abspath(__file__))
if _pb_ac and _pb_ac not in _sys_ac.path:
    _sys_ac.path.insert(0, _pb_ac)
import bimtools_access_bootstrap as _bimtools_access
# === END BIZARDS_PROD_PORTABLE_BOOTSTRAP (prod_builder) ===
if _bimtools_access.require_tool_access(__file__, __revit__, __title__):
    _pin_scripts_first(_scripts_dir)
    try:
        _mod = imp.load_source(_MAIN_MODULE_ID, _module_path)
        _mod.run(__revit__, close_on_finish=True)
    except Exception as ex:
        TaskDialog.Show(
            _TOOL_DIALOG_TITLE,
            u"Error ejecutando la rutina:\n{0}".format(ex),
        )
        raise
