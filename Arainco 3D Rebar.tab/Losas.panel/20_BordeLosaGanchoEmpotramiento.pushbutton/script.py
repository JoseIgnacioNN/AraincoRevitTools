# -*- coding: utf-8 -*-
"""Refuerzo borde losa — entrada pyRevit (copia portable).

Resuelve ``<pushbutton>/scripts/`` en primer lugar; si no existe, sube directorios
hasta encontrar ``scripts/barras_bordes_losa_gancho_empotramiento.py`` (extensión BIMTools).
"""

import os
import sys
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_TOOL_DIALOG_TITLE = u"Arainco: Refuerzo Borde Losa"
_MAIN_MODULE = "barras_bordes_losa_gancho_empotramiento.py"


def _find_scripts_dir(pushbutton_dir):
    """
    1) ``<pushbutton>/scripts/barras_bordes_losa_gancho_empotramiento.py`` (portable).
    2) ``.../scripts/`` subiendo hasta 10 niveles (extensión clásica).
    """
    local = os.path.join(pushbutton_dir, "scripts")
    if os.path.isfile(os.path.join(local, _MAIN_MODULE)):
        return os.path.abspath(local)

    cursor = pushbutton_dir
    for _ in range(10):
        candidate = os.path.join(cursor, "scripts")
        if os.path.isfile(os.path.join(candidate, _MAIN_MODULE)):
            return os.path.abspath(candidate)
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


def _ensure_scripts_on_path(scripts_dir):
    if scripts_dir and scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_scripts_dir = _find_scripts_dir(_pushbutton_dir)

if not _scripts_dir:
    TaskDialog.Show(
        _TOOL_DIALOG_TITLE,
        u"No se encontró scripts/{0}".format(_MAIN_MODULE),
    )
    raise Exception(u"No se encontró scripts/{0}".format(_MAIN_MODULE))

_ensure_scripts_on_path(_scripts_dir)

try:
    import bimtools_paths

    bimtools_paths.set_pushbutton_dir(_pushbutton_dir)
except Exception:
    pass

for _mod_name in (
    "enfierrado_shaft_hashtag",
    "barras_bordes_losa_gancho_empotramiento",
    "seleccion_caras_elemento",
    "lap_detail_link_schema",
    "embed_anchorage_link_schema",
):
    sys.modules.pop(_mod_name, None)

try:
    _module_path = os.path.join(_scripts_dir, _MAIN_MODULE)
    _mod = imp.load_source("barras_bordes_losa_gancho_empotramiento", _module_path)
    _mod.run_pyrevit(__revit__)
except Exception as ex:
    TaskDialog.Show(
        _TOOL_DIALOG_TITLE,
        u"Error ejecutando la rutina:\n{0}".format(ex),
    )
