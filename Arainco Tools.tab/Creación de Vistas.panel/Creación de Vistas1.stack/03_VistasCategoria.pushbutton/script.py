# -*- coding: utf-8 -*-
"""
Pushbutton: vistas por categoría (conjunto 01_ENTREGABLE por categoría/zona).
Lógica en scripts/vistas_por_categoria_ui.py y scripts/vistas_por_categoria/.
"""

__title__ = u"Arainco: Vistas por categoría"
__author__ = "BIMTools"
__doc__ = (
    u"Crea plantas Cielo/Piso por nivel, plantillas y tipos Detail/Sección "
    u"para la categoría y zona seleccionadas (clasificación 01_ENTREGABLE)."
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
        if os.path.isfile(os.path.join(_sp, "vistas_por_categoria_ui.py")):
            _scripts_dir = _sp
            break
        _p = os.path.dirname(_d)
        if _p == _d:
            break
        _d = _p

    if _scripts_dir is None:
        from pyrevit import forms

        forms.alert(
            u"No se encontró scripts/vistas_por_categoria_ui.py en la extensión.",
            title=u"Arainco: Vistas por categoría",
        )
        sys.exit(1)

    if _scripts_dir not in sys.path:
        sys.path.insert(0, _scripts_dir)

    # Hot-reload: lógica de la herramienta (no purgar tema WPF / helpers compartidos)
    for _key in list(sys.modules.keys()):
        if (
            _key == "vistas_por_categoria_ui"
            or _key == "vistas_por_categoria"
            or _key.startswith("vistas_por_categoria.")
        ):
            try:
                del sys.modules[_key]
            except Exception:
                pass

    from vistas_por_categoria_ui import run

    run(__revit__)
