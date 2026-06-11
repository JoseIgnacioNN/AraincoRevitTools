# -*- coding: utf-8 -*-
"""
Bloqueo de la ventana principal de Revit durante ProgressBar (autocontenido en el pushbutton).
"""

from __future__ import print_function


def _hwnd_to_int(hwnd):
    if hwnd is None:
        return None
    try:
        if hasattr(hwnd, u"ToInt32"):
            return int(hwnd.ToInt32())
    except Exception:
        pass
    try:
        if hasattr(hwnd, u"ToInt64"):
            return int(hwnd.ToInt64())
    except Exception:
        pass
    try:
        return int(hwnd)
    except Exception:
        return None


def _revit_main_window_set_enabled(revit, enable):
    if revit is None:
        return
    try:
        from infra.revit_wpf_window_position import revit_main_hwnd
    except Exception:
        return
    hwnd = revit_main_hwnd(revit)
    h = _hwnd_to_int(hwnd)
    if h is None or h == 0:
        return
    try:
        import ctypes
    except Exception:
        return
    try:
        ctypes.windll.user32.EnableWindow(h, 1 if enable else 0)
    except Exception:
        pass


class BloquearComandosRevit(object):
    """Deshabilita la ventana principal de Revit durante el bloque with."""

    def __init__(self, revit):
        self._revit = revit
        self._touched = False

    def __enter__(self):
        _revit_main_window_set_enabled(self._revit, False)
        self._touched = True
        return self

    def __exit__(self, _exc_type, _exc, _tb):
        if self._touched:
            _revit_main_window_set_enabled(self._revit, True)
        self._touched = False
        return False
