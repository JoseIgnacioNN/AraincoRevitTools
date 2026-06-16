# -*- coding: utf-8 -*-
"""
Arainco: Armado vigas — entrada pyRevit (portable).

Prioriza ``<pushbutton>/scripts/``; en desarrollo puede resolver
``BIMTools.extension/scripts/`` subiendo directorios.
"""

__title__ = "Armado\nvigas"
__author__ = "BIMTools"
__doc__ = (
    "Armado longitudinal y estribos en vigas. Selección de lote (vigas, columnas, muros); "
    "guías con sonda 50 mm y extremos empotrado/gancho; coloca Rebar longitudinales y "
    "estribos Ext/Cent con confinamiento E (Perimetral, pares, trabas)."
)

import os
import sys

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG = u"Arainco: Armado vigas"
_PKG_MARKER = os.path.join("armado_vigas", "__init__.py")

_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_local_scripts = os.path.join(_pushbutton_dir, "scripts")

# Resolver scripts/ (local o extensión canónica en desarrollo)
_scripts_dir = _local_scripts
if not os.path.isfile(os.path.join(_local_scripts, _PKG_MARKER)):
    _cursor = _pushbutton_dir
    _scripts_dir = None
    for _ in range(12):
        _candidate = os.path.join(_cursor, "scripts")
        if os.path.isfile(os.path.join(_candidate, _PKG_MARKER)):
            _scripts_dir = os.path.abspath(_candidate)
            break
        _parent = os.path.dirname(_cursor)
        if _parent == _cursor:
            break
        _cursor = _parent

if not _scripts_dir or not os.path.isfile(os.path.join(_scripts_dir, _PKG_MARKER)):
    TaskDialog.Show(
        _DIALOG,
        u"No se encontró scripts/armado_vigas/.\n\n"
        u"Ejecute sync_portable_scripts.py o copie el paquete scripts/ completo.",
    )
    raise Exception(u"No se encontró scripts/armado_vigas/")

if _pushbutton_dir not in sys.path:
    sys.path.insert(0, _pushbutton_dir)
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

from bootstrap import (
    prepare_runtime,
    purge_armado_vigas_modules,
    setup_armado_vigas_paths,
)

setup_armado_vigas_paths()
purge_armado_vigas_modules()

import bimtools_access_bootstrap as _bimtools_access

if _bimtools_access.require_tool_access(__file__, __revit__, __title__):
    try:
        os.environ["ARAINCO_ARMADO_VIGAS_PB_DIR"] = _pushbutton_dir
    except Exception:
        pass

    try:
        prepare_runtime(_pushbutton_dir)
        from armado_vigas.revit.run import run_pyrevit

        run_pyrevit(__revit__)
    except Exception as ex:
        try:
            msg = unicode(ex)
        except NameError:
            msg = str(ex)
        TaskDialog.Show(_DIALOG, u"Error:\n\n{0}".format(msg))
        raise
