# -*- coding: utf-8 -*-
"""Prioriza ``<pushbutton>/scripts/`` sobre ``BIMTools.extension/scripts/``."""

from __future__ import print_function

import os
import sys

_LOCAL_SCRIPTS_DIR = None
_PUSHBUTTON_DIR = None


def local_scripts_dir():
    global _LOCAL_SCRIPTS_DIR
    if _LOCAL_SCRIPTS_DIR is None:
        try:
            _LOCAL_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            _LOCAL_SCRIPTS_DIR = os.getcwd()
    return _LOCAL_SCRIPTS_DIR


def pushbutton_dir():
    global _PUSHBUTTON_DIR
    if _PUSHBUTTON_DIR is None:
        _PUSHBUTTON_DIR = os.path.dirname(local_scripts_dir())
    return _PUSHBUTTON_DIR


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
    """Antepone ``scripts/`` del botón y deja la raíz en posición 1."""
    sd = local_scripts_dir()
    pb = pushbutton_dir()
    _pin_dir_first(sd)
    if pb and pb != sd:
        if pb in sys.path:
            try:
                sys.path.remove(pb)
            except Exception:
                pass
        sys.path.insert(1, pb)
    return sd


pin_local_scripts_first()
