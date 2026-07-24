# -*- coding: utf-8 -*-
"""
Arainco: Vistas por categoría — conjunto 01_ENTREGABLE por categoría/zona.

Versión portable: dependencias en ``<pushbutton>/scripts/`` (vistas_por_categoria, infra).
Helpers compartidos en ``BIMTools.extension/scripts/crear_vistas_revision_estructural.py``.
"""

__title__ = u"Arainco: Vistas por categoría"
__author__ = "BIMTools"
__doc__ = (
    u"Crea plantas Cielo/Piso por nivel, plantillas y tipos Detail/Sección "
    u"para la categoría y zona seleccionadas (clasificación 01_ENTREGABLE)."
)

import os
import sys

_PUSHBUTTON_DIR = os.path.dirname(os.path.abspath(__file__))
_SCRIPTS_DIR = os.path.join(_PUSHBUTTON_DIR, "scripts")

if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

# --- Validación acceso corporativo (RECURSOS COMPARTIDOS) ---
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
    from bootstrap import purge_vistas_por_categoria_modules, setup_vistas_por_categoria_paths

    setup_vistas_por_categoria_paths()
    purge_vistas_por_categoria_modules()

    from run import main

    main(__revit__)
