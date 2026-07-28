# -*- coding: utf-8 -*-
"""Prioriza ``BIMTools.extension/scripts/`` (este directorio) en ``sys.path``.

Usado por Armado Muros y módulos que viven en la carpeta canónica de la extensión.
Si ``bimtools_paths`` tiene un pushbutton registrado, también lo deja en posición 1
(para logos / assets locales del botón).
"""

from __future__ import print_function

import os
import sys

_SCRIPTS_DIR = None


def local_scripts_dir():
    """Directorio ``…/extension/scripts`` donde vive este módulo."""
    global _SCRIPTS_DIR
    if _SCRIPTS_DIR is None:
        try:
            _SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            _SCRIPTS_DIR = os.getcwd()
    return _SCRIPTS_DIR


def pushbutton_dir():
    """Pushbutton registrado vía ``bimtools_paths``, si existe."""
    try:
        import bimtools_paths

        pb = bimtools_paths.get_pushbutton_dir()
        if pb and os.path.isdir(pb):
            return pb
    except Exception:
        pass
    return None


def _pin_dir_first(path):
    if not path or not os.path.isdir(path):
        return
    try:
        while path in sys.path:
            sys.path.remove(path)
    except Exception:
        pass
    sys.path.insert(0, path)


def pin_local_scripts_first():
    """Antepone ``extension/scripts/``; opcionalmente el pushbutton en posición 1."""
    sd = local_scripts_dir()
    _pin_dir_first(sd)
    pb = pushbutton_dir()
    if pb and pb != sd:
        if pb in sys.path:
            try:
                sys.path.remove(pb)
            except Exception:
                pass
        sys.path.insert(1, pb)
    return sd


pin_local_scripts_first()
