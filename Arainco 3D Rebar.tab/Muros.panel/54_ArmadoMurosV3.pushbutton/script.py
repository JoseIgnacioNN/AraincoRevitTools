# -*- coding: utf-8 -*-
"""Arainco: Armado Muros v3 — entrada pyRevit fina.

Lógica en ``BIMTools.extension/scripts/`` (armado_muros_run / lineales / preview_ui).
"""

__title__ = "Armado\nMuros v3"
__author__ = "BIMTools"
__doc__ = (
    "Armado Muros v3: UI estilo Machones (Inicio, Término). "
    "Misma lógica que v2. Solo muro tradicional. "
    "Mallas: herramienta dedicada «Mallas en muros»."
)

import os
import sys
import traceback

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Armado Muros v3"
_MAIN_FILE = u"armado_muros_run.py"

_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))


def _find_extension_scripts(start_dir):
    cursor = start_dir
    for _ in range(24):
        candidate = os.path.join(cursor, u"scripts", _MAIN_FILE)
        if os.path.isfile(candidate):
            return os.path.dirname(candidate)
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
            u"No se encontró scripts/{0} en la extensión.".format(_MAIN_FILE),
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
        # Forzar recarga del bootstrap: si queda en sys.modules, el fingerprint
        # viejo no ve cambios en scripts/.
        if u"armado_muros_run" in sys.modules:
            try:
                del sys.modules[u"armado_muros_run"]
            except Exception:
                pass
        from armado_muros_run import run_pyrevit

        run_pyrevit(__revit__)
    except Exception:
        _err = traceback.format_exc()
        try:
            print(_err)
        except Exception:
            pass
        try:
            TaskDialog.Show(
                _DIALOG_TITLE,
                u"Error al iniciar Armado Muros v3:\n\n{0}".format(_err[-1800:]),
            )
        except Exception:
            pass
        raise
