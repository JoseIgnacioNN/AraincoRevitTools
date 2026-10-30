# -*- coding: utf-8 -*-
"""Arainco: Suples en Losa — entrada pyRevit; lógica en scripts/."""

__title__ = u"Suples\nen Losa"
__author__ = u"BIMTools"
__doc__ = (
    u"Selecciona una losa y abre un canvas de planta para dibujar "
    u"el armado superior sobre la losa y su contexto."
)

import os
import imp
import sys

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog


_TOOL_DIALOG_TITLE = u"Arainco: Suples en Losa"
_MAIN_MODULE = u"suples_en_losa.py"
_MAIN_MODULE_ID = u"suples_en_losa"


def _find_module(start_dir):
    cursor = start_dir
    for _ in range(24):
        candidate = os.path.join(cursor, u"scripts", _MAIN_MODULE)
        if os.path.isfile(candidate):
            return candidate
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


def _aviso_arranque(mensaje):
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        show_message_dialog(
            _TOOL_DIALOG_TITLE,
            instruction=mensaje,
            ok_text=u"Entendido",
            hwnd_revit=revit_main_hwnd(__revit__),
            uiapp=__revit__,
        )
        return
    except Exception:
        pass
    TaskDialog.Show(_TOOL_DIALOG_TITLE, mensaje)


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_module_path = _find_module(_pushbutton_dir)

if not _module_path:
    _aviso_arranque(u"No se encontró scripts/{0}".format(_MAIN_MODULE))
    raise Exception(u"No se encontró scripts/{0}".format(_MAIN_MODULE))

_scripts_dir = os.path.dirname(_module_path)
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

import bimtools_paths

bimtools_paths.set_pushbutton_dir(_pushbutton_dir)

# --- Validacion acceso corporativo (prod: bootstrap junto al boton) ---
# === BEGIN BIZARDS_PROD_PORTABLE_BOOTSTRAP (prod_builder) ===
import os as _os_ac
import sys as _sys_ac

_pb_ac = _os_ac.path.dirname(_os_ac.path.abspath(__file__))
if _pb_ac and _pb_ac not in _sys_ac.path:
    _sys_ac.path.insert(0, _pb_ac)
import bimtools_access_bootstrap as _bimtools_access
# === END BIZARDS_PROD_PORTABLE_BOOTSTRAP (prod_builder) ===
if _bimtools_access.require_tool_access(__file__, __revit__, __title__):
    try:
        _mod = imp.load_source(_MAIN_MODULE_ID, _module_path)
        _mod.run(__revit__)
    except Exception as ex:
        try:
            msg = unicode(ex)
        except NameError:
            msg = str(ex)
        _aviso_arranque(u"Error al ejecutar la herramienta:\n\n{0}".format(msg))
        raise
