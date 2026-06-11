# -*- coding: utf-8 -*-
"""Resolución de la carpeta raíz del pushbutton (padre de ``scripts/``)."""

from __future__ import print_function

import os

_PB_DIR = None


def pushbutton_dir():
    """``<pushbutton>/`` — un nivel arriba de ``scripts/``."""
    global _PB_DIR
    if _PB_DIR is None:
        try:
            _PB_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        except NameError:
            _PB_DIR = os.getcwd()
    return _PB_DIR
