# -*- coding: utf-8 -*-
"""Arainco: Mallas en muros — entrada pyRevit fina.

Lógica en ``BIMTools.extension/scripts/`` (mallas_en_muros_run).
Motor de creación: ``armado_muros_lineales`` / preview_ui modo mallas (V3).
"""

__title__ = "Mallas\nen muros"
__author__ = "BIMTools"
__doc__ = (
    "Area Reinforcement ext.+int. en muros seleccionados. "
    "Elevación apilada estilo Armado Muros v3; sin fases Inicio/Término."
)

import os
import sys
import traceback

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Mallas en muros"
_MAIN_FILE = u"mallas_en_muros_run.py"

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


import os as _os_ac
import sys as _sys_ac

_tab_ac = _os_ac.path.dirname(_os_ac.path.abspath(__file__))
for _iac in range(16):
    if _os_ac.path.basename(_tab_ac).endswith(u".tab"):
        break
    _parent_ac = _os_ac.path.dirname(_tab_ac)
    if _parent_ac == _tab_ac:
        _tab_ac = None
        break
    _tab_ac = _parent_ac
if _tab_ac and _tab_ac not in _sys_ac.path:
    _sys_ac.path.insert(0, _tab_ac)
import bimtools_access_bootstrap as _bimtools_access

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
        if u"mallas_en_muros_run" in sys.modules:
            try:
                del sys.modules[u"mallas_en_muros_run"]
            except Exception:
                pass
        from mallas_en_muros_run import run_pyrevit

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
                u"Error al iniciar Mallas en muros:\n\n{0}".format(_err[-1800:]),
            )
        except Exception:
            pass
        raise
