# -*- coding: utf-8 -*-
"""
Arainco: Exportar Láminas — PDF, DWG y listado Excel.

Versión portable: dependencias en ``<pushbutton>/scripts/`` (capas lib, mvvm, ui, infra).
Ver ESTRUCTURA_PORTABLE.txt para despliegue.
"""

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
import imp

_PUSHBUTTON_DIR = os.path.dirname(os.path.abspath(__file__))
_SCRIPTS_DIR = os.path.join(_PUSHBUTTON_DIR, "scripts")
_MAIN_MODULE = "run.py"

if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

from bootstrap import purge_export_laminas_modules, setup_export_laminas_paths

setup_export_laminas_paths()
purge_export_laminas_modules()

_module_path = os.path.join(_SCRIPTS_DIR, _MAIN_MODULE)
_mod = imp.load_source("run", _module_path)
_mod.main(__revit__)
