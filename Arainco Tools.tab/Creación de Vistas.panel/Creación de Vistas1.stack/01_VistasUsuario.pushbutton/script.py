# -*- coding: utf-8 -*-
"""
Pushbutton: vistas por usuario (conjunto 02_TRABAJO por modelador).
Lógica en scripts/vistas_por_usuario_ui.py y scripts/vistas_por_usuario/.
"""

__title__ = u"Arainco: Vistas por usuario"
__author__ = "BIMTools"
__doc__ = (
    u"Crea plantas Cielo/Piso por nivel, plantillas y tipos Detail/Sección "
    u"para el modelador seleccionado desde personas.json (clasificación 02_TRABAJO)."
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
        if os.path.isfile(os.path.join(_sp, "vistas_por_usuario_ui.py")):
            _scripts_dir = _sp
            break
        _p = os.path.dirname(_d)
        if _p == _d:
            break
        _d = _p

    if _scripts_dir is None:
        from pyrevit import forms

        forms.alert(
            u"No se encontró scripts/vistas_por_usuario_ui.py en la extensión.",
            title=u"Arainco: Vistas por usuario",
        )
        sys.exit(1)

    if _scripts_dir not in sys.path:
        sys.path.insert(0, _scripts_dir)

    # Hot-reload: lógica de la herramienta (no purgar tema WPF / helpers compartidos)
    for _key in list(sys.modules.keys()):
        if (
            _key == "vistas_por_usuario_ui"
            or _key == "vistas_por_usuario"
            or _key.startswith("vistas_por_usuario.")
            or _key == "crear_vistas_revision_estructural"
        ):
            try:
                del sys.modules[_key]
            except Exception:
                pass

    from vistas_por_usuario_ui import run

    run(__revit__)
