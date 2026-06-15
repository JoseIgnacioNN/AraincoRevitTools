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


# --- Validación acceso corporativo (RECURSOS COMPARTIDOS) ---
import os as _os_ac
import sys as _sys_ac
_tab_ac = _os_ac.path.dirname(_os_ac.path.abspath(__file__))
for _iac in range(16):
    if _os_ac.path.basename(_tab_ac).endswith(u".tab"):
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
    run(__revit__)
