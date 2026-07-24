# -*- coding: utf-8 -*-
"""Bootstrap de rutas para pushbutton portable «Vistas Usuario»."""

from __future__ import print_function

import os
import sys

# Solo lógica de la herramienta (hot-reload). NO purgar tema WPF ni helpers
# compartidos: re-parsear ~40 KB de estilos + reimportar crear_vistas_*
# en cada clic retrasa el despliegue de la UI.
_MODULES_TO_PURGE = (
    "run",
    "bootstrap",
)

_PACKAGE_PREFIXES = (
    "vistas_por_usuario",
    "vistas_por_usuario.",
)


def _find_extension_scripts_dir(from_dir):
    cursor = os.path.abspath(from_dir)
    for _ in range(24):
        candidate = os.path.join(cursor, "scripts")
        marker = os.path.join(candidate, "crear_vistas_revision_estructural.py")
        if os.path.isfile(marker):
            return candidate
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


def setup_vistas_por_usuario_paths():
    """
    Prioriza ``<pushbutton>/scripts/``; añade ``.../extension/scripts/`` para helpers compartidos.
    """
    try:
        scripts_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        scripts_dir = os.getcwd()

    if scripts_dir and os.path.isdir(scripts_dir):
        try:
            while scripts_dir in sys.path:
                sys.path.remove(scripts_dir)
        except Exception:
            pass
        sys.path.insert(0, scripts_dir)

    ext_scripts = _find_extension_scripts_dir(scripts_dir)
    if ext_scripts and ext_scripts not in sys.path:
        sys.path.append(ext_scripts)

    return scripts_dir


def purge_vistas_por_usuario_modules():
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
