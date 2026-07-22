# -*- coding: utf-8 -*-
"""Entrada pyRevit delgada para Armado Columnas enterprise (copia portable).

Flujo: resuelve ``scripts/`` → ``bimtools_paths`` (opcional) → ejecuta
``column_reinforcement.runner.run_pyrevit`` (fachada única; ver ARCHITECTURE.py).
"""

__title__ = u"Armado\nColumnas"

import os
import sys
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_TOOL_DIALOG_TITLE = u"Arainco: Armado Columnas"


def _find_scripts_dir(pushbutton_dir):
    """
    1) ``<pushbutton>/scripts/column_reinforcement/runner.py`` (portable).
    2) Misma resolución que la extensión clásica: sube directorios hasta 10 niveles.
    """
    local = os.path.join(pushbutton_dir, "scripts")
    if os.path.isfile(os.path.join(local, "column_reinforcement", "runner.py")):
        return os.path.abspath(local)

    cursor = pushbutton_dir
    for _ in range(10):
        candidate = os.path.join(cursor, "scripts")
        runner = os.path.join(candidate, "column_reinforcement", "runner.py")
        if os.path.isfile(runner):
            return os.path.abspath(candidate)
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


def _ensure_scripts_on_path():
    _pin_scripts_first(_scripts_dir)


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


def _purge_portable_modules():
    for name in (
        "column_reinforcement_layout_rps",
        "column_stirrup_creator",
        "column_stirrup_tags",
        "confinement_dim_link_schema",
        "confinement_dim_updater_dmu",
        "armado_muros_cabezal_tags",
    ):
        sys.modules.pop(name, None)
    for key in list(sys.modules.keys()):
        if key == "column_reinforcement" or key.startswith("column_reinforcement."):
            sys.modules.pop(key, None)
        if key == "armado_vigas" or key.startswith("armado_vigas."):
            sys.modules.pop(key, None)


def _active_view_is_section_or_elevation(uiapp):
    _ensure_scripts_on_path()
    try:
        from column_reinforcement.revit.api.context import (
            is_section_or_elevation_uiapp,
        )

        return is_section_or_elevation_uiapp(uiapp)
    except Exception:
        return False


def __selfinit__(script_cmp_key, uiapp_api, __revit__):
    """Habilita el botón solo en vistas de sección o alzado."""
    return _active_view_is_section_or_elevation(uiapp_api)


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_scripts_dir = _find_scripts_dir(_pushbutton_dir)

if not _scripts_dir:
    TaskDialog.Show(
        _TOOL_DIALOG_TITLE,
        u"No se encontró scripts/column_reinforcement/runner.py",
    )
    raise Exception(u"No se encontró scripts/column_reinforcement/runner.py")

if _pushbutton_dir not in sys.path:
    sys.path.insert(0, _pushbutton_dir)

_ensure_scripts_on_path()
_purge_portable_modules()

try:
    import bimtools_paths

    bimtools_paths.set_pushbutton_dir(_pushbutton_dir)
except Exception:
    pass


# --- Validación acceso corporativo (RECURSOS COMPARTIDOS) ---
import bimtools_access_bootstrap as _bimtools_access

if _bimtools_access.require_tool_access(__file__, __revit__, __title__):
    _pin_scripts_first(_scripts_dir)
    _purge_portable_modules()
    try:
        _runner_path = os.path.join(_scripts_dir, "column_reinforcement", "runner.py")
        _runner = imp.load_source("column_reinforcement.runner", _runner_path)

        _runner.run_pyrevit(__revit__)
    except Exception as ex:
        TaskDialog.Show(
            _TOOL_DIALOG_TITLE,
            u"Error ejecutando la rutina:\n{0}".format(ex),
        )
