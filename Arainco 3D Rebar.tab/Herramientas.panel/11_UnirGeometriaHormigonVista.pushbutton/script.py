# -*- coding: utf-8 -*-
"""
Unir geometría — hormigón (vista activa).

Trigger pyRevit: la lógica vive en ``scripts/join_geometry_concrete_vista.py``.
"""

__title__ = u"Unir\nGeom. Hormigón"
__author__ = u"BIMTools"
__doc__ = (
    u"Unir Join Geometry entre elementos de material estructural Concrete, "
    u"acotado a la vista activa."
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
        if os.path.isfile(os.path.join(_sp, "join_geometry_concrete_vista.py")):
            _scripts_dir = _sp
            break
        _p = os.path.dirname(_d)
        if _p == _d:
            break
        _d = _p

    if _scripts_dir is None:
        from pyrevit import forms

        forms.alert(
            u"No se encontro scripts/join_geometry_concrete_vista.py en la extension.",
            title=__title__,
        )
        sys.exit(1)

    if _scripts_dir not in sys.path:
        sys.path.insert(0, _scripts_dir)

    import join_geometry_concrete_vista as _mod

    reload(_mod)
    _mod.run(__revit__)
