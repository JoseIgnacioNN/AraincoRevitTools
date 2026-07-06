# -*- coding: utf-8 -*-
"""
Visibilidad de barras estructurales por Host Category en la vista activa.

Revit 2024+ | pyRevit / IronPython.
Desarrollo: importado desde el pushbutton o ejecutado vía scripts del botón 40.
"""

from __future__ import print_function

import os
import sys

_PUSHBUTTON_SCRIPTS = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    u"..",
    u"BIMTools.tab",
    u"Armadura.panel",
    u"40_RebarVisibilidadHost.pushbutton",
    u"scripts",
)
_PUSHBUTTON_SCRIPTS = os.path.abspath(_PUSHBUTTON_SCRIPTS)
if os.path.isdir(_PUSHBUTTON_SCRIPTS) and _PUSHBUTTON_SCRIPTS not in sys.path:
    sys.path.insert(0, _PUSHBUTTON_SCRIPTS)

from run import run

__all__ = (u"run",)
