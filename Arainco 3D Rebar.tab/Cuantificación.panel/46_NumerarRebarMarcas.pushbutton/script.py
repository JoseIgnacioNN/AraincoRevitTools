# -*- coding: utf-8 -*-
"""
Arainco: numerar marcas de Rebar en el documento completo.
"""

__title__ = u"Numerar\nmarcas Rebar"
__author__ = u"BIMTools"
__doc__ = (
    u"Recorre todas las Rebar del documento y las RebarInSystem de Area "
    u"Reinforcement, las agrupa por forma, tipo, segmentos (roundup 10 mm), "
    u"ganchos y end treatments, y escribe en Armadura_Marca la marca "
    u"{diámetro}{número} (ej. 815). Numeración incremental: conserva marcas "
    u"existentes y solo crea índices nuevos para grupos sin marca previa."
)

import os
import sys

_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
_ext_root = os.path.dirname(os.path.dirname(os.path.dirname(_pushbutton_dir)))
_scripts_dir = os.path.join(_ext_root, "scripts")
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

import bimtools_paths

bimtools_paths.set_pushbutton_dir(_pushbutton_dir)

from numerar_rebar_marcas import run

import os as _os_ac
import sys as _sys_ac

_tab_ac = _os_ac.path.dirname(_os_ac.path.abspath(__file__))
for _iac in range(16):
    if _os_ac.path.basename(_tab_ac) == u"BIMTools.tab":
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
