# -*- coding: utf-8 -*-
"""Control de instancia única — Vistas por Categoría."""

from __future__ import print_function

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import AppDomain
from System.Windows import WindowState

from vistas_por_categoria.constants import APP_DOMAIN_KEY


def try_activate_existing():
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
    AppDomain.CurrentDomain.SetData(APP_DOMAIN_KEY, win)


def clear():
    try:
        AppDomain.CurrentDomain.SetData(APP_DOMAIN_KEY, None)
    except Exception:
        pass
