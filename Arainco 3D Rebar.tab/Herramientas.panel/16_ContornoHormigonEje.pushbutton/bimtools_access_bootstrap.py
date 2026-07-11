# -*- coding: utf-8 -*-
"""
Bootstrap de acceso corporativo — portabilidad total.

Solo usa ``<pushbutton>/scripts/`` local. No sube carpetas padre ni depende
de ``BIMTools.extension/scripts/``.
"""

from __future__ import print_function

import os
import sys

_MARKER = u"corporate_access.py"


def _local_scripts_dir(script_file):
    pushbutton_dir = os.path.dirname(os.path.abspath(script_file))
    return os.path.join(pushbutton_dir, u"scripts")


def require_tool_access(script_file, uiapp, button_title):
    """
    Valida acceso corporativo desde el paquete local del pushbutton.

    Devuelve True si la herramienta puede ejecutarse. Si falta
    ``corporate_access.py`` en ``scripts/``, se omite la validación.
    """
    scripts = _local_scripts_dir(script_file)
    if scripts not in sys.path:
        sys.path.insert(0, scripts)
    if not os.path.isfile(os.path.join(scripts, _MARKER)):
        return True

    from bimtools_script_guard import guard_tool

    return guard_tool(uiapp, button_title=button_title)
