# -*- coding: utf-8 -*-
"""
Control de instancia única de ventana WPF vía AppDomain.

Garantiza que la herramienta Revisiones no pueda abrirse dos veces
simultáneamente, y activa/restaura la ventana existente si se intenta.
"""

from __future__ import print_function

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import AppDomain
from System.Windows import WindowState

from siguiente_revision.constants import APP_DOMAIN_KEY


def try_activate_existing():
    """
    Intenta activar la ventana ya abierta.

    Returns:
        True si existe una ventana activa y fue activada.
        False si no hay ventana activa (la herramienta puede abrirse).
    """
    try:
        o = AppDomain.CurrentDomain.GetData(APP_DOMAIN_KEY)
        if o is None:
            return False
        win = o.Target if hasattr(o, "Target") else o
        if win is None:
            AppDomain.CurrentDomain.SetData(APP_DOMAIN_KEY, None)
            return False
        try:
            if not getattr(win, "IsVisible", False):
                AppDomain.CurrentDomain.SetData(APP_DOMAIN_KEY, None)
                return False
            win.WindowState = WindowState.Normal
            win.Activate()
            return True
        except Exception:
            AppDomain.CurrentDomain.SetData(APP_DOMAIN_KEY, None)
            return False
    except Exception:
        return False


def register(win):
    """Registra la ventana activa en AppDomain."""
    try:
        AppDomain.CurrentDomain.SetData(APP_DOMAIN_KEY, win)
    except Exception:
        pass


def clear():
    """Libera la referencia para permitir una nueva apertura."""
    try:
        AppDomain.CurrentDomain.SetData(APP_DOMAIN_KEY, None)
    except Exception:
        pass
