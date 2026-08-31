# -*- coding: utf-8 -*-
"""Antepone ``scripts/`` local y subcarpetas del layout portable."""

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


def _portable_roots(scripts_root):
    pb = pushbutton_dir()
    if pb and pb not in sys.path:
        sys.path.insert(0, pb)
    try:
        from portable_layout import is_portable_layout, portable_import_roots

        if is_portable_layout(scripts_root):
            return portable_import_roots(scripts_root)
    except Exception:
        pass
    return []


def pin_local_scripts_first():
    """Antepone subcarpetas portable, ``scripts/`` y raíz del pushbutton."""
    sd = local_scripts_dir()
    pb = pushbutton_dir()
    ordered = []
    ordered.extend(_portable_roots(sd))
    ordered.append(sd)
    if pb and pb != sd:
        ordered.append(pb)
    for d in reversed(ordered):
        _pin_dir_first(d)
    return sd
