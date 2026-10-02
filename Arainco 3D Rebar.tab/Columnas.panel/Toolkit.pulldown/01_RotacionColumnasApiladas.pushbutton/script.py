# -*- coding: utf-8 -*-
"""Verifica la rotación de columnas de hormigón apiladas."""

__title__ = u"Rotación\ncolumnas"
__author__ = u"BIMTools"
__doc__ = (
    u"Colorea columnas de hormigón según su rotación "
    u"y dibuja la leyenda en la vista activa."
)

import os
import sys

_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_MAIN_MODULE = u"rotacion_columnas_apiladas.py"
_MAIN_MODULE_ID = u"rotacion_columnas_apiladas"


def _find_scripts_dir(start_dir):
    cursor = start_dir
    for _ in range(16):
        candidate = os.path.join(cursor, "scripts", _MAIN_MODULE)
        if os.path.isfile(candidate):
            return os.path.join(cursor, "scripts")
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


_scripts_dir = _find_scripts_dir(_pushbutton_dir)
if _scripts_dir and _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

try:
    import bimtools_paths

    bimtools_paths.set_pushbutton_dir(_pushbutton_dir)
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
    if _scripts_dir is None:
        from Autodesk.Revit.UI import TaskDialog

        TaskDialog.Show(
            u"Arainco: Rotación columnas",
            u"No se encontró scripts/rotacion_columnas_apiladas.py.",
        )
    else:
        if _MAIN_MODULE_ID in sys.modules:
            try:
                del sys.modules[_MAIN_MODULE_ID]
            except Exception:
                pass
        try:
            from rotacion_columnas_apiladas import run

            run(__revit__)
        except Exception as ex:
            import traceback

            from Autodesk.Revit.UI import TaskDialog

            TaskDialog.Show(
                u"Arainco: Rotación columnas",
                u"Error al iniciar la herramienta:\n\n{0}\n\n{1}".format(
                    ex, traceback.format_exc()
                ),
            )
