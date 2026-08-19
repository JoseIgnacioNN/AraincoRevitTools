# -*- coding: utf-8 -*-
"""Elevación Eje — entrada pyRevit; lógica en scripts/elevacion_eje_ui.py."""

__title__ = u"Arainco: Elevación\nEje"
__author__ = u"BIMTools"
__doc__ = (
    u"Solo en vistas de planta. Lista los ejes (Grids), permite seleccionar "
    u"uno o varios y crea una sección Building Section por cada eje marcado, "
    u"con la escala elegida. Hereda los filtros de la plantilla y aplica un "
    u"filtro Armadura_Eje a cada vista; si la plantilla rige los filtros, "
    u"la edita para permitir un filtro distinto por vista. Fija Detail Level "
    u"Coarse en la plantilla de las vistas creadas."
)

import os
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_TOOL_DIALOG_TITLE = u"Arainco: Elevación Eje"
_MAIN_MODULE = "elevacion_eje_ui.py"
_MAIN_MODULE_ID = "elevacion_eje_ui"


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


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_module_path = _find_module(_pushbutton_dir)

if not _module_path:
    TaskDialog.Show(
        _TOOL_DIALOG_TITLE,
        u"No se encontró scripts/{0}".format(_MAIN_MODULE),
    )
    raise Exception(u"No se encontró scripts/{0}".format(_MAIN_MODULE))

import sys

_scripts_dir = os.path.dirname(_module_path)
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)
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
    try:
        _mod = imp.load_source(_MAIN_MODULE_ID, _module_path)
        _mod.run(__revit__)
    except Exception as ex:
        TaskDialog.Show(
            _TOOL_DIALOG_TITLE,
            u"Error ejecutando la rutina:\n{0}".format(ex),
        )
