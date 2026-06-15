# -*- coding: utf-8 -*-
"""
Arainco: Armado Muros — cabezal + mallas (solo muro tradicional).

Versión portable: dependencias en ``<pushbutton>/scripts/``.
Orden: longitudinales, confinamiento, mallas; etiquetas long., conf., malla.
"""

__title__ = "Armado\nMuros"
__author__ = "BIMTools"
__doc__ = (
    "Muros tradicionales: cabezal ini/fin y malla ext.+int. en un asistente. "
    "Crea armadura y etiqueta al final. No incluye muro de contención."
)

import os
import sys

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Armado Muros"
_MAIN_FILE = u"armado_muros_lineales.py"

_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_scripts_dir = os.path.join(_pushbutton_dir, u"scripts")

if not os.path.isfile(os.path.join(_scripts_dir, _MAIN_FILE)):
    TaskDialog.Show(
        _DIALOG_TITLE,
        u"No se encontró scripts/{0} en la carpilla del botón.".format(_MAIN_FILE),
    )
    raise Exception(u"No se encontró el módulo principal de Armado Muros")

if _pushbutton_dir not in sys.path:
    sys.path.insert(0, _pushbutton_dir)

from armado_muros_run import purge_armado_muros_modules, setup_armado_muros_paths

setup_armado_muros_paths()
purge_armado_muros_modules()

from armado_muros_lineales import run_unificado

# --- Validación acceso corporativo (RECURSOS COMPARTIDOS) ---
import os as _os_ac
import sys as _sys_ac
_tab_ac = _os_ac.path.dirname(_os_ac.path.abspath(__file__))
for _iac in range(16):
    if _os_ac.path.basename(_tab_ac) == u"BIMTools.tab":
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
    setup_armado_muros_paths()
    purge_armado_muros_modules()
    run_unificado(__revit__)
