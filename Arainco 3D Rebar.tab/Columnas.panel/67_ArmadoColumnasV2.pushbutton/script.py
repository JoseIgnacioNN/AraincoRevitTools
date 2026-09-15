# -*- coding: utf-8 -*-
"""
Arainco: Armado columnas V2 — entrada pyRevit (botón ligero).

UI shell basada en Armado vigas: canvas de elevación + rail lateral de sección.
Lógica de colocación y selección de modelo se añadirá en iteraciones siguientes.
No modifica Armado Columnas Enterprise ni column_reinforcement_v2.
"""

__title__ = u"Armado\nColumnas V2"
__author__ = u"BIMTools"
__doc__ = (
    u"Armado de columnas V2. Selección de columnas, fundaciones, vigas y losas "
    u"en la vista activa; elevación a escala (dimensiones, posición y relación) "
    u"y preview de sección en rail lateral."
)

import os
import sys

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

_DIALOG = u"Arainco: Armado columnas V2"
_PKG_MARKER = os.path.join("armado_columnas_v2", "__init__.py")

_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))

# Resolver BIMTools.extension/scripts/
_cursor = _pushbutton_dir
_scripts_dir = None
for _ in range(16):
    _candidate = os.path.join(_cursor, "scripts")
    if os.path.isfile(os.path.join(_candidate, _PKG_MARKER)):
        _scripts_dir = os.path.abspath(_candidate)
        break
    _parent = os.path.dirname(_cursor)
    if _parent == _cursor:
        break
    _cursor = _parent

if not _scripts_dir:
    TaskDialog.Show(
        _DIALOG,
        u"No se encontró scripts/armado_columnas_v2/.\n"
        u"Compruebe que la extensión BIMTools está completa.",
    )
    raise Exception(u"No se encontró scripts/armado_columnas_v2/")

if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

try:
    import bimtools_paths

    bimtools_paths.set_pushbutton_dir(_pushbutton_dir)
except Exception:
    pass

# Acceso corporativo (walk-up hasta *.tab)
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
    try:
        # Forzar recarga de módulos en diseño (pyRevit cache)
        _stale = [k for k in list(sys.modules.keys()) if k.startswith("armado_columnas_v2")]
        for _k in _stale:
            del sys.modules[_k]
        from armado_columnas_v2.run import run

        run(__revit__, pushbutton_dir=_pushbutton_dir)
    except Exception as ex:
        try:
            msg = unicode(ex)
        except NameError:
            msg = str(ex)
        TaskDialog.Show(_DIALOG, u"Error:\n\n{0}".format(msg))
        raise
