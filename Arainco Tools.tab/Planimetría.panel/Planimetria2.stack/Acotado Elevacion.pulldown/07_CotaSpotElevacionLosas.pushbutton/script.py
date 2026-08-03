# -*- coding: utf-8 -*-
"""
Pushbutton: cota alineada + Spot Elevations en losas (sección/alzado).
Lógica en scripts/cota_spot_elevacion_losas.py.
"""

__title__ = u"Arainco: Cota Spot\nLosas"
__author__ = "BIMTools"
__doc__ = (
    u"En Sección o Alzado: seleccione dos o más losas, indique la posición de la "
    u"cota y genera una cota alineada entre caras superiores horizontales más "
    u"Spot Elevations (tipo Survey Point_Nivel Tope de Losa si existe). "
    u"Los elementos quedan visibles solo en la vista activa del árbol de dependientes."
)

import os
import sys

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
    _scripts_dir = None
    _d = os.path.dirname(os.path.abspath(__file__))
    for _ in range(24):
        _sp = os.path.join(_d, "scripts")
        if os.path.isfile(os.path.join(_sp, "cota_spot_elevacion_losas.py")):
            _scripts_dir = _sp
            break
        _p = os.path.dirname(_d)
        if _p == _d:
            break
        _d = _p

    if _scripts_dir is None:
        from Autodesk.Revit.UI import TaskDialog

        TaskDialog.Show(
            u"Arainco: Cota y Spot Losas",
            u"No se encontró scripts/cota_spot_elevacion_losas.py en la extensión.",
        )
    else:
        if _scripts_dir not in sys.path:
            sys.path.insert(0, _scripts_dir)
        try:
            import bimtools_paths

            bimtools_paths.set_pushbutton_dir(os.path.dirname(os.path.abspath(__file__)))
        except Exception:
            pass

        for _key in list(sys.modules.keys()):
            if _key == "cota_spot_elevacion_losas":
                try:
                    del sys.modules[_key]
                except Exception:
                    pass

        from cota_spot_elevacion_losas import run

        run(__revit__)
