# -*- coding: utf-8 -*-
"""
Bootstrap de validación corporativa para script.py de pushbuttons.

Uso típico al inicio del flujo principal (después de resolver sys.path):

    from bimtools_script_guard import ensure_extension_scripts_on_path, guard_tool
    ensure_extension_scripts_on_path(__file__)
    if not guard_tool(__revit__, __title__):
        pass
    else:
        ... lógica de la herramienta ...
"""

from __future__ import print_function

import os
import sys

_MARKER = u"corporate_access.py"


def ensure_extension_scripts_on_path(from_file):
    """Inserta ``.../extension/scripts`` en sys.path si existe corporate_access.py."""
    cursor = os.path.dirname(os.path.abspath(from_file))
    for _ in range(24):
        scripts = os.path.join(cursor, "scripts")
        marker = os.path.join(scripts, _MARKER)
        if os.path.isfile(marker):
            if scripts not in sys.path:
                sys.path.insert(0, scripts)
            return scripts
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


def dialog_title_from_button_title(button_title):
    text = button_title.replace("\n", " ").strip()
    try:
        text = unicode(text)
    except NameError:
        text = str(text)
    lower = text.lower()
    if lower.startswith(u"arainco:") or lower.startswith(u"arainco |"):
        return text
    return u"Arainco: " + text


def guard_tool(uiapp, button_title=None, dialog_title=None):
    """
    Valida acceso corporativo. Devuelve True si la herramienta puede continuar.
    """
    import corporate_access

    title = dialog_title or dialog_title_from_button_title(button_title or u"BIMTools")
    return corporate_access.ensure_corporate_access(dialog_title=title, uiapp=uiapp)
