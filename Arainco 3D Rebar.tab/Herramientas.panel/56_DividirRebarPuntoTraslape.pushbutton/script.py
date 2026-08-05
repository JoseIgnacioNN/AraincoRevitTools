# -*- coding: utf-8 -*-
"""
Dividir rebar en punto con traslape — entrada pyRevit (copia portable).

Todo el código vive en ``<pushbutton>/scripts/``.
"""

__title__ = u"Dividir barra\ntraslape"
__author__ = u"BIMTools"
__doc__ = (
    u"Selecciona una Structural Rebar y abre la UI multipunto: marque cortes "
    u"sobre la barra en la vista de Revit (marca temporal; Finalizar en la "
    u"barra de opciones), afine vanos en mm y aplique divisiones con "
    u"traslape según diámetro (tabla BIMTools G25/G35/G45; default G25). "
    u"Admite layout Single, Fixed Number y Maximum Spacing."
)

import os
import sys
import imp

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Dividir barra con traslape"
_MAIN_MODULE = "dividir_rebar_punto.py"
_MAIN_MODULE_ID = "dividir_rebar_punto"
_REQUIRED_MODULES = (
    _MAIN_MODULE,
    "dividir_rebar_punto_ui.py",
    "dividir_rebar_punto_core.py",
    "dividir_rebar_punto_geom.py",
    "dividir_rebar_punto_lap_detail.py",
    "bimtools_rebar_hook_lengths.py",
    "bimtools_rebar_3d_visibility.py",
    "rebar_tag_shape_sync_core.py",
    "dividir_rebar_punto_tags.py",
    "dividir_rebar_punto_shapes.py",
    "bimtools_instruction_dialog.py",
    "bimtools_ui_tokens.py",
    "bimtools_wpf_shell.py",
    "bimtools_wpf_dark_theme.py",
    "revit_wpf_window_position.py",
    "corporate_access.py",
    "bimtools_script_guard.py",
)
# Solo purgar módulos de la herramienta. Mantener theme/tokens/revit_wpf/
# corporate en caché: el purge completo forzaba reparse XAML (~40 KB) +
# reimport CLR en cada clic y colgaba Revit.
_MODULES_TO_PURGE = (
    _MAIN_MODULE_ID,
    "dividir_rebar_punto",
    "dividir_rebar_punto_ui",
    "dividir_rebar_punto_core",
    "dividir_rebar_punto_geom",
    "dividir_rebar_punto_lap_detail",
    "dividir_rebar_punto_tags",
    "dividir_rebar_punto_shapes",
    "bimtools_rebar_hook_lengths",
    "bimtools_rebar_3d_visibility",
    "rebar_tag_shape_sync_core",
)


def _scripts_dir(pushbutton_dir):
    return os.path.abspath(os.path.join(pushbutton_dir, "scripts"))


def _missing_modules(scripts_dir):
    missing = []
    for name in _REQUIRED_MODULES:
        if not os.path.isfile(os.path.join(scripts_dir, name)):
            missing.append(name)
    return missing


def _show_error_dialog(message):
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = None
        try:
            hwnd = revit_main_hwnd(__revit__)
        except Exception:
            hwnd = None
        show_message_dialog(
            _DIALOG_TITLE,
            instruction=u"Error al ejecutar la herramienta.",
            content=message,
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=__revit__,
        )
        return
    except Exception:
        pass
    TaskDialog.Show(
        _DIALOG_TITLE,
        u"Error al ejecutar la herramienta:\n\n{0}".format(message),
    )


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
            u"\n- ".join(_missing)
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
    # Un solo purge tras acceso (no dos): evita reimportar el paquete dos veces.
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
        _show_error_dialog(msg)
        raise
