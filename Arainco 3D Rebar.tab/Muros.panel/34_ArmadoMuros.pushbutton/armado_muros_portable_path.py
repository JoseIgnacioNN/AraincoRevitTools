# -*- coding: utf-8 -*-
"""Bootstrap de rutas: pushbutton + ``scripts/`` (34_ArmadoMuros portable)."""

from __future__ import print_function

import os
import sys

_PB_DIR = None
_SCRIPTS_DIR = None


def pushbutton_dir():
    global _PB_DIR
    if _PB_DIR is None:
        try:
            _PB_DIR = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            _PB_DIR = os.getcwd()
    return _PB_DIR


def scripts_dir():
    global _SCRIPTS_DIR
    if _SCRIPTS_DIR is None:
        _SCRIPTS_DIR = os.path.join(pushbutton_dir(), u"scripts")
    return _SCRIPTS_DIR


def ensure_pushbutton_on_path():
    """Prioriza ``scripts/`` empaquetado del botón sobre librerías globales."""
    sd = scripts_dir()
    pb = pushbutton_dir()
    if sd and os.path.isdir(sd):
        if sd not in sys.path:
            sys.path.insert(0, sd)
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
    if pb and pb != sd and pb not in sys.path:
        sys.path.insert(1, pb)
    return sd
