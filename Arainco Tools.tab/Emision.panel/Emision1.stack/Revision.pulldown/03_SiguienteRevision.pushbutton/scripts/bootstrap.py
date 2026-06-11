# -*- coding: utf-8 -*-
"""Bootstrap de rutas e imports para pushbutton portable 03_SiguienteRevision."""

from __future__ import print_function

import os
import sys

_MODULES_TO_PURGE = (
    "run",
    "bootstrap",
    "bootstrap_path",
    "gestionar_personas_wpf",
    "sheet_revision_display",
    "bimtools_paths",
    "bimtools_wpf_dark_theme",
    "revit_wpf_window_position",
    "revit_window_blocker",
    "join_geometry_concrete_vista",
    "exportar_laminas_pdf_dwg",
)

_PACKAGE_PREFIXES = (
    "siguiente_revision",
    "siguiente_revision.",
    "infra.",
    "ui.",
    "lib.",
)


def setup_siguiente_revision_paths():
    """Inserta ``<pushbutton>/scripts/`` al frente de ``sys.path``."""
    try:
        scripts_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        scripts_dir = os.getcwd()
    if scripts_dir and os.path.isdir(scripts_dir) and scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    return scripts_dir


def purge_siguiente_revision_modules():
    for name in _MODULES_TO_PURGE:
        try:
            if name in sys.modules:
                del sys.modules[name]
        except Exception:
            pass
    for key in list(sys.modules.keys()):
        for prefix in _PACKAGE_PREFIXES:
            if key == prefix.rstrip(".") or key.startswith(prefix):
                try:
                    del sys.modules[key]
                except Exception:
                    pass
                break
