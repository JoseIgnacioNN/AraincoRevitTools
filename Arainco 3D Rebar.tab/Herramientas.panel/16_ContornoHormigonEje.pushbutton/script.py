# -*- coding: utf-8 -*-
"""
Contorno de hormigón por eje — entrada pyRevit (copia portable).

Todo el código vive en ``<pushbutton>/scripts/``. Copie solo esta carpeta
.pushbutton; no depende de BIMTools.extension ni de rutas externas.
"""

__title__ = u"Contorno\nHormigón"
__author__ = u"BIMTools"
__doc__ = (
    u"Selecciona el hormigón (Concrete) de la vista activa, lo une en un sólido, "
    u"corta con el plano del eje elegido y crea detail lines agrupadas."
)

import os
import sys
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Contorno hormigón por eje"
_MAIN_MODULE = u"contorno_hormigon_eje.py"
_MAIN_MODULE_ID = u"contorno_hormigon_eje"
_REQUIRED_MODULES = (
    _MAIN_MODULE,
    u"contorno_hormigon_eje_ui.py",
    u"contorno_material_concrete.py",
    u"bimtools_wpf_dark_theme.py",
    u"revit_wpf_window_position.py",
)
_MODULES_TO_PURGE = (
    _MAIN_MODULE_ID,
    u"contorno_hormigon_eje",
    u"contorno_hormigon_eje_ui",
    u"contorno_material_concrete",
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
            u"\n- ".join(_missing),
        ),
    )
    raise Exception(u"Paquete portable incompleto: {0}".format(u", ".join(_missing)))

if _pushbutton_dir not in sys.path:
    sys.path.insert(0, _pushbutton_dir)

_pin_scripts_first(_scripts_dir)
_purge_modules()

import bimtools_access_bootstrap as _bimtools_access

if _bimtools_access.require_tool_access(__file__, __revit__, __title__):
    _pin_scripts_first(_scripts_dir)
    _purge_modules()
    try:
        _module_path = os.path.join(_scripts_dir, _MAIN_MODULE)
        _mod = imp.load_source(_MAIN_MODULE_ID, _module_path)
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
