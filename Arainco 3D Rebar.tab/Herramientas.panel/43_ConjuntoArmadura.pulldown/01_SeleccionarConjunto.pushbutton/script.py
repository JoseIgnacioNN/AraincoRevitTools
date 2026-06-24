# -*- coding: utf-8 -*-
"""
Arainco: Seleccionar conjunto de armadura por GUID de creación.

Prioriza ``<pushbutton>/scripts/run.py`` (paquete portable con capas).
"""

__title__ = u"Seleccionar\nConjunto"
__author__ = u"BIMTools"
__doc__ = (
    u"Selecciona en el modelo un conjunto de armadura identificado por "
    u"«Armadura_Conjunto_GUID» a partir de una barra, empalme o croquis."
)

import os
import sys
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Seleccionar conjunto"
_MAIN_MODULE = u"run.py"
_MAIN_MODULE_ID = u"conjunto_armadura_seleccionar_run"
_RELOAD_PREFIXES = (
    _MAIN_MODULE_ID,
    "run",
    "lib",
    "lib.corrida_guid",
    "lib.conjunto_actions",
)


def _find_scripts_dir(pushbutton_dir):
    local = os.path.join(pushbutton_dir, u"scripts")
    if os.path.isfile(os.path.join(local, _MAIN_MODULE)):
        return os.path.abspath(local)
    return None


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
        if key.startswith(u"lib."):
            sys.modules.pop(key, None)


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_scripts_dir = _find_scripts_dir(_pushbutton_dir)

if not _scripts_dir:
    TaskDialog.Show(
        _DIALOG_TITLE,
        u"No se encontró scripts/run.py en la carpeta del botón.",
    )
    raise Exception(u"No se encontró módulo principal de Seleccionar Conjunto")

_pin_scripts_first(_scripts_dir)
_purge_modules()

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
        _module_path = os.path.join(_scripts_dir, _MAIN_MODULE)
        _mod = imp.load_source(_MAIN_MODULE_ID, _module_path)
        _mod.run(__revit__)
    except NameError:
        pass
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
