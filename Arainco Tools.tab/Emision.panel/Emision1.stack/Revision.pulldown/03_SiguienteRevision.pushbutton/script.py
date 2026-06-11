# -*- coding: utf-8 -*-
"""
Arainco: Revisiones — emisión de revisión en láminas.

Versión portable: dependencias en ``<pushbutton>/scripts/`` (siguiente_revision, lib, ui, infra).
Ver ESTRUCTURA_PORTABLE.txt para despliegue.
"""

__title__ = u"Arainco: Revisiones"
__author__ = "BIMTools"
__doc__ = (
    "Crea una nueva entrada de revisión por lámina a partir del último correlativo "
    "en el índice de la lámina y actualiza datos de revisión/nubes configurados."
)

import os
import sys
import imp

_PUSHBUTTON_DIR = os.path.dirname(os.path.abspath(__file__))
_SCRIPTS_DIR = os.path.join(_PUSHBUTTON_DIR, "scripts")
_MAIN_MODULE = "run.py"

if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

from bootstrap import purge_siguiente_revision_modules, setup_siguiente_revision_paths

setup_siguiente_revision_paths()
purge_siguiente_revision_modules()

_module_path = os.path.join(_SCRIPTS_DIR, _MAIN_MODULE)
_mod = imp.load_source("run", _module_path)
_mod.main(__revit__)
