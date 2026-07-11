# -*- coding: utf-8 -*-
"""
Cuantificación de armadura en losas — entrada pyRevit (copia portable).

Prioriza ``<pushbutton>/scripts/run.py``.
"""

__title__ = u"Cuantif.\nLosa"
__author__ = u"BIMTools"
__doc__ = (
    u"Abre un formulario para seleccionar niveles (checkbox) y crear/actualizar "
    u"cuadros de armadura (Rebar): Malla/Sin Malla × Superior (F')/Inferior (F). "
    u"Filtros: Host Category = Floor, Armadura_Nivel, Armadura_Malla y "
    u"Armadura_Ubicacion."
)

import os
import sys
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Cuantificación losa por nivel"
_MAIN_MODULES = (
    ("run.py", "run"),
)
_MAIN_MODULE_ID = u"cuantificacion_losa_run"
_RELOAD_PREFIXES = (
    _MAIN_MODULE_ID,
    "run",
    "cuantificacion_losa_ui",
    "lib",
    "lib.rebar_schedule_cuantificacion_losa",
    "lib.cuantificacion_progress",
    "bimtools_ui_tokens",
    "bimtools_wpf_shell",
    "bimtools_wpf_dark_theme",
    "bimtools_instruction_dialog",
    "revit_wpf_window_position",
)


def _find_scripts_dir(pushbutton_dir):
    local = os.path.join(pushbutton_dir, "scripts")
    for module_file, _ in _MAIN_MODULES:
        if os.path.isfile(os.path.join(local, module_file)):
            return os.path.abspath(local), module_file

    cursor = pushbutton_dir
    for _ in range(10):
        candidate = os.path.join(cursor, "scripts")
        for module_file, _ in _MAIN_MODULES:
            if os.path.isfile(os.path.join(candidate, module_file)):
                return os.path.abspath(candidate), module_file
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None, None


def _pin_scripts_first(scripts_dir):
    if not scripts_dir:
        return
    try:
        while scripts_dir in sys.path:
            sys.path.remove(scripts_dir)
    except Exception:
        pass
    sys.path.insert(0, scripts_dir)


def _purge_modules():
    for name in _RELOAD_PREFIXES:
        sys.modules.pop(name, None)
    for key in list(sys.modules.keys()):
        if key.startswith("lib."):
            sys.modules.pop(key, None)


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_scripts_dir, _main_module = _find_scripts_dir(_pushbutton_dir)

if not _scripts_dir:
    TaskDialog.Show(
        _DIALOG_TITLE,
        u"No se encontró scripts/run.py",
    )
    raise Exception(u"No se encontró módulo principal de cuantificación losa")

_pin_scripts_first(_scripts_dir)
_purge_modules()

# --- Validación acceso corporativo (RECURSOS COMPARTIDOS) ---
import os as _os_ac
import sys as _sys_ac

_tab_ac = _os_ac.path.dirname(_os_ac.path.abspath(__file__))
for _iac in range(16):
    if _os_ac.path.basename(_tab_ac) == u"BIMTools.tab":
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
    _pin_scripts_first(_scripts_dir)
    _purge_modules()
    try:
        _module_path = os.path.join(_scripts_dir, _main_module)
        _mod_name = (
            _MAIN_MODULE_ID
            if _main_module == "run.py"
            else os.path.splitext(_main_module)[0]
        )
        _mod = imp.load_source(_mod_name, _module_path)
        _mod.run(__revit__)
    except Exception as ex:
        try:
            msg = unicode(ex)
        except NameError:
            msg = str(ex)
        TaskDialog.Show(
            _DIALOG_TITLE,
            u"Error al ejecutar la herramienta:\n\n{0}".format(msg),
        )
        raise
