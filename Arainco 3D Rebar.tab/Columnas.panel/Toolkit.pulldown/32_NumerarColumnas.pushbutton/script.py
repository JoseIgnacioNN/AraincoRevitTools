# -*- coding: utf-8 -*-
"""
Numerar Columnas — copia portable (todo el código en esta carpeta).

Copiar ``32_NumerarColumnas.pushbutton`` a otra extensión o panel pyRevit;
no depende de ``scripts/`` de BIMTools.
"""

__title__ = u"Arainco: Numerar\nColumnas"
__author__ = "pyRevit"
__doc__ = (
    u"Numera columnas estructurales apiladas según torre y fundación. "
    u"Un esquema por lote; escribe «Numeracion Columna» al confirmar."
)

import os
import sys

_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
if _pushbutton_dir not in sys.path:
    sys.path.insert(0, _pushbutton_dir)

from numerar_columnas_ui import run


import bimtools_access_bootstrap as _bimtools_access
if _bimtools_access.require_tool_access(__file__, __revit__, __title__):
    run(__revit__)
