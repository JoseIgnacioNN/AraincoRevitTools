# -*- coding: utf-8 -*-
"""
Pushbutton: Exportar cuadros Excel.
Lógica en scripts/exportar_cuadros_excel_ui.py
"""

__title__ = u"Exportar\ncuadros Excel"
__author__ = u"BIMTools"
__doc__ = (
    u"Selecciona tablas (cuadros de cantidades) del proyecto y las exporta "
    u"a un archivo Excel (.xlsx), una hoja por tabla."
)

import os
import sys

_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_MAIN_MODULE = u"exportar_cuadros_excel_ui.py"
_MAIN_MODULE_ID = u"exportar_cuadros_excel_ui"


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
            u"Arainco: Exportar cuadros Excel",
            u"No se encontró scripts/exportar_cuadros_excel_ui.py.",
        )
    else:
        for _mod_id in list(sys.modules.keys()):
            if (
                _mod_id == _MAIN_MODULE_ID
                or _mod_id == u"exportar_cuadros_excel"
                or _mod_id.startswith(u"exportar_cuadros_excel.")
            ):
                try:
                    del sys.modules[_mod_id]
                except Exception:
                    pass
        try:
            from exportar_cuadros_excel_ui import run

            run(__revit__)
        except Exception as ex:
            import traceback

            from Autodesk.Revit.UI import TaskDialog

            TaskDialog.Show(
                u"Arainco: Exportar cuadros Excel",
                u"Error al iniciar la herramienta:\n\n{0}\n\n{1}".format(
                    ex, traceback.format_exc()
                ),
            )
