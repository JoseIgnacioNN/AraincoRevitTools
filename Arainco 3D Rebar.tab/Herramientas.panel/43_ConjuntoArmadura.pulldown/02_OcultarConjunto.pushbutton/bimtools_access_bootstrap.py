# -*- coding: utf-8 -*-
"""
Bootstrap de acceso corporativo — copia portable (Conjunto armadura).

Sube directorios desde ``script.py`` buscando ``<extensión>/scripts/corporate_access.py``.
Si no existe (despliegue en otra extensión sin guardia), permite la ejecución.
"""

from __future__ import print_function

import os
import sys

_MARKER = u"corporate_access.py"


def _find_extension_scripts_dir(from_file):
    cursor = os.path.dirname(os.path.abspath(from_file))
    for _ in range(24):
        scripts = os.path.join(cursor, "scripts")
        if os.path.isfile(os.path.join(scripts, _MARKER)):
            return scripts
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


def require_tool_access(script_file, uiapp, button_title):
    """
    Instala ``scripts/`` de la extensión en sys.path y valida acceso corporativo.

    Devuelve True si la herramienta puede ejecutarse. Sin ``corporate_access.py``
    en la jerarquía de la extensión destino, se omite la validación (modo portable).
    """
    scripts = _find_extension_scripts_dir(script_file)
    if scripts and scripts not in sys.path:
        sys.path.insert(0, scripts)
    if not scripts:
        return True

    from bimtools_script_guard import guard_tool

    return guard_tool(uiapp, button_title=button_title)
