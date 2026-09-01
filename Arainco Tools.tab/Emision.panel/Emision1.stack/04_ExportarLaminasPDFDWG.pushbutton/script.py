# -*- coding: utf-8 -*-
"""Exportar Láminas — entrada pyRevit; lógica en scripts/exportar_laminas/."""

__title__ = "Exportar\nLáminas"
__author__ = "BIMTools"
__doc__ = (
    "Selecciona y exporta láminas (PDF/DWG). Nombre Personalizado: encabezado «Nombre de archivo». "
    "Opcional: listado Excel de las seleccionadas (plantilla TemplateListado en la carpeta del botón). "
    "Ruta de entrega completa y editable; «Examinar…» la completa con YYYY.MM.DD_ENTREGA (hoy). "
    "Subcarpetas PDF y DWG."
)

import os
import sys

_TOOL_DIALOG_TITLE = u"Arainco: Exportar Láminas"
_MAIN_REL = os.path.join("exportar_laminas", "run.py")


def _show_entry_error(revit, text):
    """Aviso de arranque: WPF estándar; TaskDialog solo si no hay scripts o falla WPF."""
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        uiapp = None
        try:
            uiapp = revit.Application
        except Exception:
            uiapp = revit
        if show_message_dialog(
            _TOOL_DIALOG_TITLE,
            text,
            u"",
            ok_text=u"Entendido",
            hwnd_revit=revit_main_hwnd(uiapp) if uiapp is not None else None,
            uiapp=uiapp,
        ):
            return
    except Exception:
        pass
    try:
        import clr

        clr.AddReference("RevitAPIUI")
        from Autodesk.Revit.UI import TaskDialog

        TaskDialog.Show(_TOOL_DIALOG_TITLE, text)
    except Exception:
        pass


def _find_scripts_dir(start_dir):
    cursor = start_dir
    for _ in range(10):
        candidate = os.path.join(cursor, "scripts", _MAIN_REL)
        if os.path.isfile(candidate):
            return os.path.join(cursor, "scripts")
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_scripts_dir = _find_scripts_dir(_pushbutton_dir)

if not _scripts_dir:
    try:
        import clr

        clr.AddReference("RevitAPIUI")
        from Autodesk.Revit.UI import TaskDialog

        TaskDialog.Show(
            _TOOL_DIALOG_TITLE,
            u"No se encontró scripts/exportar_laminas/run.py",
        )
    except Exception:
        pass
    raise Exception(u"No se encontró scripts/exportar_laminas/run.py")

if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)
import bimtools_paths

bimtools_paths.set_pushbutton_dir(_pushbutton_dir)

for _key in list(sys.modules.keys()):
    if _key == u"exportar_laminas" or _key.startswith(u"exportar_laminas."):
        try:
            del sys.modules[_key]
        except Exception:
            pass

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
        from exportar_laminas.run import run

        run(__revit__)
    except Exception as ex:
        _show_entry_error(
            __revit__,
            u"Error ejecutando la rutina:\n{0}".format(ex),
        )
