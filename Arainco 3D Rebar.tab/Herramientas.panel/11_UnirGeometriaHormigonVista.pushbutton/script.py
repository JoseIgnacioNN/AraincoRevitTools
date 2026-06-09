# -*- coding: utf-8 -*-
"""
Unir geometría — hormigón (vista activa). Entrada pyRevit (copia portable).

Todo el código vive en ``<pushbutton>/scripts/``. Copie solo esta carpeta
.pushbutton; no depende de BIMTools.extension ni de rutas externas.
"""

__title__ = u"Unir\nGeom. Hormigón"
__author__ = u"BIMTools"
__doc__ = (
    u"Unir Join Geometry entre elementos de material estructural Concrete, "
    u"acotado a la vista activa."
)

import os
import sys
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Unir geometría (hormigón, vista)"
_MAIN_MODULE = u"join_geometry_concrete_vista.py"
_REQUIRED_MODULES = (
    _MAIN_MODULE,
    u"join_geometry_instruction_dialog.py",
    u"join_geometry_material_concrete.py",
    u"bimtools_wpf_dark_theme.py",
    u"revit_wpf_window_position.py",
)
_MODULES_TO_PURGE = (
    u"join_geometry_concrete_vista",
    u"join_geometry_instruction_dialog",
    u"join_geometry_material_concrete",
    u"bimtools_wpf_dark_theme",
    u"revit_wpf_window_position",
)


def _scripts_dir(pushbutton_dir):
    """Única fuente: ``<pushbutton>/scripts/`` (sin fallback externo)."""
    return os.path.abspath(os.path.join(pushbutton_dir, u"scripts"))


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

try:
    _module_path = os.path.join(_scripts_dir, _MAIN_MODULE)
    _mod = imp.load_source(u"join_geometry_concrete_vista", _module_path)
    _mod.run(__revit__)
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
