# -*- coding: utf-8 -*-
"""
Dividir y Traslapar — entrada pyRevit (botón ligero).

Lógica en ``BIMTools.extension/scripts/dividir_rebar_punto*.py``.
"""

__title__ = u"Dividir y\nTraslapar"
__author__ = u"BIMTools"
__doc__ = (
    u"Selecciona una Structural Rebar y abre la UI multipunto: marque cortes "
    u"en el esquema de elevación, afine vanos en mm y aplique divisiones con "
    u"traslape. Etiqueta cada tramo y aplica MRA «Recorrido Barras»."
)

import os
import sys

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Dividir y Traslapar"
_MAIN_MODULE = u"dividir_rebar_punto.py"


def _find_scripts_dir(start_dir):
    cursor = start_dir
    for _ in range(16):
        candidate = os.path.join(cursor, u"scripts", _MAIN_MODULE)
        if os.path.isfile(candidate):
            return os.path.dirname(candidate)
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_scripts_dir = _find_scripts_dir(_pushbutton_dir)

if not _scripts_dir:
    TaskDialog.Show(
        _DIALOG_TITLE,
        u"No se encontró scripts/{0}".format(_MAIN_MODULE),
    )
    raise Exception(u"No se encontró scripts/{0}".format(_MAIN_MODULE))

if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

try:
    import bimtools_paths

    bimtools_paths.set_pushbutton_dir(_pushbutton_dir)
except Exception:
    pass

# Acceso corporativo: walk-up hasta cualquier *.tab
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
        from dividir_rebar_punto import run

        run(__revit__)
    except Exception as ex:
        try:
            msg = unicode(ex)
        except NameError:
            msg = str(ex)
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
                content=msg,
                ok_text=u"Entendido",
                hwnd_revit=hwnd,
                uiapp=__revit__,
            )
        except Exception:
            TaskDialog.Show(
                _DIALOG_TITLE,
                u"Error al ejecutar la herramienta:\n\n{0}".format(msg),
            )
        raise
