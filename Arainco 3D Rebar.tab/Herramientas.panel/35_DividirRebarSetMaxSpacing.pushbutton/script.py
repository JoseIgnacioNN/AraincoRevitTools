# -*- coding: utf-8 -*-
"""
Dividir rebar set (Maximum Spacing) — entrada pyRevit (copia portable).

Todo el código vive en ``<pushbutton>/scripts/``. Copie solo esta carpeta
.pushbutton; no depende de BIMTools.extension ni de rutas externas.
"""

__title__ = "Dividir rebar\nset (Max Sp.)"
__author__ = "BIMTools"
__doc__ = (
    "Divide un conjunto de barras con regla Maximum Spacing en dos subconjuntos "
    "en la barra que indiques."
)

import os
import sys
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Dividir rebar set Maximum Spacing"
_MAIN_MODULE = "dividir_rebar_set_max_spacing.py"
_REQUIRED_MODULES = (
    _MAIN_MODULE,
    "dividir_rebar_set_instruction_dialog.py",
    "bimtools_wpf_dark_theme.py",
    "revit_wpf_window_position.py",
)
_MODULES_TO_PURGE = (
    "dividir_rebar_set_max_spacing",
    "dividir_rebar_set_instruction_dialog",
    "bimtools_wpf_dark_theme",
    "revit_wpf_window_position",
)


def _scripts_dir(pushbutton_dir):
    """Única fuente: ``<pushbutton>/scripts/`` (sin fallback externo)."""
    return os.path.abspath(os.path.join(pushbutton_dir, "scripts"))


def _missing_modules(scripts_dir):
    missing = []
    for name in _REQUIRED_MODULES:
        if not os.path.isfile(os.path.join(scripts_dir, name)):
            missing.append(name)
    return missing


def _ensure_scripts_on_path(scripts_dir):
    if scripts_dir and scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)


def _purge_modules():
    for mod_name in _MODULES_TO_PURGE:
        try:
            if mod_name in sys.modules:
                del sys.modules[mod_name]
        except Exception:
            pass


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_scripts_dir = _scripts_dir(_pushbutton_dir)
_missing = _missing_modules(_scripts_dir)

if _missing:
    TaskDialog.Show(
        _DIALOG_TITLE,
        u"Paquete portable incompleto. Faltan en scripts/:\n\n- {0}".format(
            u"\n- ".join(_missing)
        ),
    )
    raise Exception(u"Paquete portable incompleto: {0}".format(u", ".join(_missing)))

_ensure_scripts_on_path(_scripts_dir)
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
    try:
        _module_path = os.path.join(_scripts_dir, _MAIN_MODULE)
        _mod = imp.load_source("dividir_rebar_set_max_spacing", _module_path)
        _mod.run_pyrevit(__revit__)
    except Exception as ex:
        try:
            msg = unicode(ex)
        except NameError:
            msg = str(ex)
        TaskDialog.Show(
            _DIALOG_TITLE,
            u"Error al ejecutar la herramienta:\n\n{}".format(msg),
        )
        raise
