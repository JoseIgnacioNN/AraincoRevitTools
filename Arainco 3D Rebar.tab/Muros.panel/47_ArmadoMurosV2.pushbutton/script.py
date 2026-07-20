# -*- coding: utf-8 -*-
"""
Arainco: Armado Muros v2 — cabezal + mallas (solo muro tradicional).

Copia experimental de 34_ArmadoMuros con optimizaciones de UI.
Versión portable: dependencias en ``<pushbutton>/scripts/``.
Orden: longitudinales, confinamiento, mallas; etiquetas long., conf., malla.
"""

__title__ = "Armado\nMuros v2"
__author__ = "BIMTools"
__doc__ = (
    "Experimental: Armado Muros con UI más rápida. "
    "Muros tradicionales: cabezal ini/fin y malla ext.+int. "
    "No incluye muro de contención."
)

import os
import sys
import traceback

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG_TITLE = u"Arainco: Armado Muros v2"
_MAIN_FILE = u"armado_muros_lineales.py"

_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_scripts_dir = os.path.join(_pushbutton_dir, u"scripts")

if not os.path.isfile(os.path.join(_scripts_dir, _MAIN_FILE)):
    TaskDialog.Show(
        _DIALOG_TITLE,
        u"No se encontró scripts/{0} en la carpilla del botón.".format(_MAIN_FILE),
    )
    raise Exception(u"No se encontró el módulo principal de Armado Muros v2")

if _pushbutton_dir not in sys.path:
    sys.path.insert(0, _pushbutton_dir)

try:
    from armado_muros_run import (
        ensure_armado_muros_modules_fresh,
        setup_armado_muros_paths,
    )

    setup_armado_muros_paths()
    ensure_armado_muros_modules_fresh()

    from armado_muros_lineales import run_unificado

    # Acceso corporativo: ``bimtools_access_bootstrap.py`` en esta carpilla (portable).
    import bimtools_access_bootstrap as _bimtools_access

    if _bimtools_access.require_tool_access(__file__, __revit__, __title__):
        setup_armado_muros_paths()
        ensure_armado_muros_modules_fresh()
        run_unificado(__revit__)
except Exception:
    _err = traceback.format_exc()
    try:
        print(_err)
    except Exception:
        pass
    try:
        TaskDialog.Show(
            _DIALOG_TITLE,
            u"Error al iniciar Armado Muros v2:\n\n{0}".format(_err[-1800:]),
        )
    except Exception:
        pass
    raise
