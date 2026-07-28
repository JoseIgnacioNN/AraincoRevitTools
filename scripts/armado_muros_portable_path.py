# -*- coding: utf-8 -*-
"""Resolución de rutas para Armado Muros (scripts canónicos de la extensión)."""

from __future__ import print_function

import os
import sys

_SCRIPTS_DIR = None


def scripts_dir():
    global _SCRIPTS_DIR
    if _SCRIPTS_DIR is None:
        try:
            _SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            _SCRIPTS_DIR = os.getcwd()
    return _SCRIPTS_DIR


def pushbutton_dir():
    try:
        import bimtools_paths

        pb = bimtools_paths.get_pushbutton_dir()
        if pb and os.path.isdir(pb):
            return pb
    except Exception:
        pass
    return None


def ensure_pushbutton_on_path():
    """Prioriza ``extension/scripts/`` sobre otras entradas de ``sys.path``."""
    sd = scripts_dir()
    if sd and os.path.isdir(sd):
        try:
            import bootstrap_paths

            bootstrap_paths.pin_local_scripts_first()
            return bootstrap_paths.local_scripts_dir()
        except Exception:
            pass
        try:
            while sd in sys.path:
                sys.path.remove(sd)
        except Exception:
            pass
        sys.path.insert(0, sd)
    pb = pushbutton_dir()
    if pb and pb != sd and pb not in sys.path:
        sys.path.insert(1, pb)
    return sd
