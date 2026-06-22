# -*- coding: utf-8 -*-
"""
Arainco: super selector de elementos por categoría de modelado en la vista activa.

Versión portable: dependencias en ``<pushbutton>/scripts/``.
Ver ESTRUCTURA_PORTABLE.txt para despliegue.
"""

__title__ = u"Super\nSelector"
__author__ = u"BIMTools"
__doc__ = (
    u"Lista las categorías de modelado con elementos en la vista activa "
    u"y permite seleccionarlos marcando checkboxes."
)

import os
import sys
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Super selector por categoría"
_MAIN_MODULE = u"run.py"
_MAIN_MODULE_ID = u"super_selector_categoria_run"
_REQUIRED_SCRIPTS = (
    u"run.py",
    u"bimtools_wpf_dark_theme.py",
    u"revit_wpf_window_position.py",
    u"lib/elements_by_category.py",
    u"ui/selector_window.py",
)
_RELOAD_PREFIXES = (
    _MAIN_MODULE_ID,
    u"run",
    u"bimtools_wpf_dark_theme",
    u"revit_wpf_window_position",
    u"lib",
    u"ui",
    u"lib.elements_by_category",
    u"ui.selector_window",
)


def _missing_modules(scripts_dir):
    missing = []
    for rel in _REQUIRED_SCRIPTS:
        if not os.path.isfile(os.path.join(scripts_dir, rel)):
            missing.append(rel)
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
    for name in _RELOAD_PREFIXES:
        sys.modules.pop(name, None)
    for key in list(sys.modules.keys()):
        if key.startswith(u"lib.") or key.startswith(u"ui."):
            sys.modules.pop(key, None)


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_scripts_dir = os.path.join(_pushbutton_dir, u"scripts")

if not os.path.isfile(os.path.join(_scripts_dir, _MAIN_MODULE)):
    TaskDialog.Show(
        _DIALOG_TITLE,
        u"No se encontró scripts/run.py en la carpeta del botón.",
    )
    raise Exception(u"No se encontró módulo principal de Super Selector")

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

# Acceso corporativo: ``bimtools_access_bootstrap.py`` en esta carpilla (portable).
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
            u"Error al abrir la herramienta:\n\n{0}".format(msg),
        )
        raise
