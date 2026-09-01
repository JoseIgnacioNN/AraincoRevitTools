# -*- coding: utf-8 -*-
"""
Pushbutton: láminas por categoría (correlativo PG-001, LO-014, …).
Lógica en scripts/laminas_por_categoria_ui.py y scripts/laminas_por_categoria/.
"""

__title__ = u"Arainco: Láminas por categoría"
__author__ = "BIMTools"
__doc__ = (
    u"Crea láminas CONTENIDO LAMINA por categoría, con correlativo "
    u"CAT-NNN, cajetín y datos de firma/fecha."
)

import os
import sys

# --- Validacion acceso corporativo (prod: bootstrap junto al boton) ---
# === BEGIN BIZARDS_PROD_PORTABLE_BOOTSTRAP (prod_builder) ===
import os as _os_ac
import sys as _sys_ac

_pb_ac = _os_ac.path.dirname(_os_ac.path.abspath(__file__))
if _pb_ac and _pb_ac not in _sys_ac.path:
    _sys_ac.path.insert(0, _pb_ac)
import bimtools_access_bootstrap as _bimtools_access
# === END BIZARDS_PROD_PORTABLE_BOOTSTRAP (prod_builder) ===
if _bimtools_access.require_tool_access(__file__, __revit__, __title__):
    _scripts_dir = None
    _d = os.path.dirname(os.path.abspath(__file__))
    for _ in range(24):
        _sp = os.path.join(_d, "scripts")
        if os.path.isfile(os.path.join(_sp, "laminas_por_categoria_ui.py")):
            _scripts_dir = _sp
            break
        _p = os.path.dirname(_d)
        if _p == _d:
            break
        _d = _p

    if _scripts_dir is None:
        from pyrevit import forms

        forms.alert(
            u"No se encontró scripts/laminas_por_categoria_ui.py en la extensión.",
            title=u"Arainco: Láminas por categoría",
        )
        sys.exit(1)

    if _scripts_dir not in sys.path:
        sys.path.insert(0, _scripts_dir)

    import bimtools_paths

    bimtools_paths.set_pushbutton_dir(os.path.dirname(os.path.abspath(__file__)))

    for _key in list(sys.modules.keys()):
        if (
            _key == "laminas_por_categoria_ui"
            or _key == "laminas_por_categoria"
            or _key.startswith("laminas_por_categoria.")
        ):
            try:
                del sys.modules[_key]
            except Exception:
                pass

    from laminas_por_categoria_ui import run

    run(__revit__)
