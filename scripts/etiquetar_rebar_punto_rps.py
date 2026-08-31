# -*- coding: utf-8 -*-
# RPS: File > Run script — wrapper de scripts/rebar_section_tag.py
"""
Wrapper RPS de Rebar Section Tag (misma lógica que el pushbutton Detallamiento).
"""

import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
if _here not in sys.path:
    sys.path.insert(0, _here)

from rebar_section_tag import run

try:
    _uiapp = __revit__
except NameError:
    _uiapp = None

run(_uiapp)
