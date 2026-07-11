# -*- coding: utf-8 -*-
"""
Numerar Fundaciones â Numera fundaciones aisladas segÃºn sus dimensiones.
Interfaz WPF alineada con el resto de herramientas de armadura (tema oscuro BIMTools).
"""

__title__ = "Numerar\nFundaciones"
__author__ = "pyRevit"
__doc__ = "Numera las fundaciones aisladas del proyecto segÃºn sus dimensiones (Length x Width)."


import os
import sys

_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_ext_root = os.path.dirname(os.path.dirname(os.path.dirname(_pushbutton_dir)))
_scripts_dir = os.path.join(_ext_root, "scripts")
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

import bimtools_paths

bimtools_paths.set_pushbutton_dir(_pushbutton_dir)

from numerar_fundaciones_ui import run


if _pushbutton_dir not in sys.path:
    sys.path.insert(0, _pushbutton_dir)
import bimtools_access_bootstrap as _bimtools_access
if _bimtools_access.require_tool_access(__file__, __revit__, __title__):
    run(__revit__)
