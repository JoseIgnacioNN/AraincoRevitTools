# -*- coding: utf-8 -*-
"""
Arainco: Corrida GUID — seleccionar, resaltar y eliminar barras por corrida.

Prioriza ``<pushbutton>/scripts/run.py`` (paquete portable con capas).
"""

__title__ = u"Seleccionar\ncorrida GUID"
__author__ = u"BIMTools"
__doc__ = (
    u"Identifica una corrida de Armado Muros por GUID, resalta sus barras "
    u"en el modelo o elimínalas."
)

import os
import sys
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Corrida GUID"
_MAIN_MODULE = u"run.py"
# Nombre único: evita colisión con otros pushbuttons que usan imp.load_source("run", ...)
_MAIN_MODULE_ID = u"corrida_guid_run"
_RELOAD_PREFIXES = (
    _MAIN_MODULE_ID,
    "run",
    "bimtools_wpf_dark_theme",
    "revit_wpf_window_position",
    "lib",
    "ui",
    "lib.corrida_guid",
    "ui.corrida_guid_window",
    "ui.ok_cancel_dialog",
)


def _find_scripts_dir(pushbutton_dir):
    """``<pushbutton>/scripts/run.py`` (portable con capas)."""
    local = os.path.join(pushbutton_dir, "scripts")
    if os.path.isfile(os.path.join(local, _MAIN_MODULE)):
        return os.path.abspath(local)
    return None


def _pin_scripts_first(scripts_dir):
    """Prioriza ``scripts/`` del botón para resolver ``lib.*`` y ``ui.*`` locales."""
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
        if key.startswith("lib.") or key.startswith("ui."):
            sys.modules.pop(key, None)


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_scripts_dir = _find_scripts_dir(_pushbutton_dir)

if not _scripts_dir:
    TaskDialog.Show(
        _DIALOG_TITLE,
        u"No se encontró scripts/run.py en la carpilla del botón.",
    )
    raise Exception(u"No se encontró módulo principal de Corrida GUID")

_pin_scripts_first(_scripts_dir)
_purge_modules()

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
            u"Error al abrir la herramienta:\n\n{0}".format(msg),
        )
        raise
