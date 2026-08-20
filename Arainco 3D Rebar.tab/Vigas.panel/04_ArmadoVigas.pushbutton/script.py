# -*- coding: utf-8 -*-
"""Seleccione vigas, columnas y muros

Estructura clásica: este script es solo entrada. Inteligencia en scripts/armado_vigas/__init__.py.
"""

__title__ = u"Armado\nVigas"
__author__ = "José Ignacio Núñez"
__doc__ = (
    u"Armado de vigas: selección de lote (vigas/columnas/muros), alzado + sección, "
    u"pestañas SUP/INF/LAT/CONF y SUPLE. Coloca Rebar longitudinal (SUP/INF), "
    u"estribos y laterales según toggles del rail."
)

import os
import sys
import traceback

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Bizards Tool: Armado vigas"
_MAIN_FILE = u"armado_vigas/__init__.py"
_ENTRY = u"run_pyrevit"
_SPECIAL = u"armado_vigas"

_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))


def _find_extension_scripts(start_dir):
    cursor = start_dir
    for _ in range(24):
        scripts = os.path.join(cursor, u"scripts")
        if _SPECIAL == u"armado_vigas":
            marker = os.path.join(scripts, u"armado_vigas", u"__init__.py")
        else:
            marker = os.path.join(scripts, _MAIN_FILE.replace(u"/", os.sep))
        if os.path.isfile(marker):
            return os.path.abspath(scripts)
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


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
    _scripts_dir = _find_extension_scripts(_pushbutton_dir)
    if not _scripts_dir:
        TaskDialog.Show(
            _DIALOG_TITLE,
            u"No se encontró scripts/{0} en la extensión BizardsTool.".format(_MAIN_FILE),
        )
        raise Exception(u"No se encontró scripts/{0}".format(_MAIN_FILE))

    if _scripts_dir not in sys.path:
        sys.path.insert(0, _scripts_dir)

    try:
        import bimtools_paths

        bimtools_paths.set_pushbutton_dir(_pushbutton_dir)
    except Exception:
        pass

    try:
        if _SPECIAL == u"armado_vigas":
            try:
                os.environ["ARAINCO_ARMADO_VIGAS_PB_DIR"] = _pushbutton_dir
            except Exception:
                pass
            try:
                from bootstrap import (
                    prepare_runtime,
                    purge_armado_vigas_modules,
                    setup_armado_vigas_paths,
                )

                setup_armado_vigas_paths()
                purge_armado_vigas_modules()
                prepare_runtime(_pushbutton_dir)
            except Exception:
                pass
            from armado_vigas.revit.run import run_pyrevit

            run_pyrevit(__revit__)
        else:
            mod_name = _MAIN_FILE.replace(u"\\", u"/").split(u"/")[-1]
            if mod_name.endswith(u".py"):
                mod_name = mod_name[:-3]
            if mod_name in sys.modules:
                try:
                    del sys.modules[mod_name]
                except Exception:
                    pass
            mod = __import__(mod_name)
            fn = getattr(mod, _ENTRY)
            fn(__revit__)
    except Exception:
        _err = traceback.format_exc()
        try:
            print(_err)
        except Exception:
            pass
        try:
            TaskDialog.Show(
                _DIALOG_TITLE,
                u"Error al iniciar la herramienta:\n\n{0}".format(_err[-1800:]),
            )
        except Exception:
            pass
        raise
