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
_REQUIRED_MODULES = (
    _MAIN_MODULE,
    "bimtools_paths.py",
    "bimtools_wpf_dark_theme.py",
    "revit_wpf_window_position.py",
    "bimtools_rebar_hook_lengths.py",
    "enfierrado_shaft_hashtag.py",
    "seleccion_caras_elemento.py",
    "lap_detail_link_schema.py",
    "embed_anchorage_link_schema.py",
)


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


def _missing_modules(scripts_dir):
    missing = []
    for name in _REQUIRED_MODULES:
        if not os.path.isfile(os.path.join(scripts_dir, name)):
            missing.append(name)
    return missing


def _pin_scripts_first(scripts_dir):
    """Prioriza ``scripts/`` del botón sobre ``BIMTools.extension/scripts/``."""
    if not scripts_dir:
        return
    try:
        while scripts_dir in sys.path:
            sys.path.remove(scripts_dir)
    except Exception:
        pass
    sys.path.insert(0, scripts_dir)


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

_missing = _missing_modules(_scripts_dir)
if _missing:
    TaskDialog.Show(
        _TOOL_DIALOG_TITLE,
        u"Paquete portable incompleto. Faltan en scripts/:\n\n- {0}".format(
            u"\n- ".join(_missing)
        ),
    )
    raise Exception(u"Paquete portable incompleto: {0}".format(u", ".join(_missing)))

if _pushbutton_dir not in sys.path:
    sys.path.insert(0, _pushbutton_dir)

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

# --- Validación acceso corporativo (RECURSOS COMPARTIDOS) ---
import bimtools_access_bootstrap as _bimtools_access

if _bimtools_access.require_tool_access(__file__, __revit__, __title__):
    _pin_scripts_first(_scripts_dir)
    try:
        _module_path = os.path.join(_scripts_dir, _MAIN_MODULE)
        _mod = imp.load_source("barras_bordes_losa_gancho_empotramiento", _module_path)
        _mod.run_pyrevit(__revit__)
    except Exception as ex:
        TaskDialog.Show(
            _TOOL_DIALOG_TITLE,
            u"Error ejecutando la rutina:\n{0}".format(ex),
        )
