# -*- coding: utf-8 -*-
"""
Punto de entrada de desarrollo para cuadro de armadura por lámina.

El pushbutton portable carga ``scripts/run.py`` local; este módulo permite
ejecutar la misma lógica desde ``scripts/`` en el árbol de la extensión.
"""

from __future__ import print_function

import os
import sys

_PUSHBUTTON_SCRIPTS = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    u"BIMTools.tab",
    u"Armadura.panel",
    u"41_TablaArmaduraLamina.pushbutton",
    u"scripts",
)

if _PUSHBUTTON_SCRIPTS not in sys.path:
    sys.path.insert(0, _PUSHBUTTON_SCRIPTS)

from run import run  # noqa: E402
