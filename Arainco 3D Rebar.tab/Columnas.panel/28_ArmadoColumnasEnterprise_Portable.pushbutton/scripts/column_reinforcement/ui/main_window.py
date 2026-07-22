# -*- coding: utf-8 -*-
"""Carga WPF compatible con IronPython/RPS y pyRevit."""

import os

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import AppDomain
from System.IO import File
from System.Windows.Markup import XamlReader

from column_reinforcement.viewmodels.main_viewmodel import ColumnReinforcementViewModel

try:
    from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
except Exception:
    BIMTOOLS_DARK_STYLES_XML = u""

try:
    clr.AddReference("RevitAPIUI")
    from Autodesk.Revit.UI import TaskDialog
except Exception:
    TaskDialog = None


_SINGLETON_KEY = "Arainco.column_reinforcement.MainWindowSingleton"


def _load_xaml_text():
    here = os.path.dirname(os.path.abspath(__file__))
    xaml_path = os.path.join(here, "main_window.xaml")
    txt = File.ReadAllText(xaml_path)
    return txt.replace("__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML)


class ColumnReinforcementWindowController(object):
    """Controlador mínimo de la vista; la lógica de negocio vive fuera de WPF."""

    def __init__(self, viewmodel=None):
        self.viewmodel = viewmodel or ColumnReinforcementViewModel()
        self.request = None
        self.window = XamlReader.Parse(_load_xaml_text())
        self.window.DataContext = self.viewmodel
        self._wire_events()

    def _wire_events(self):
        run_btn = self.window.FindName("RunButton")
        cancel_btn = self.window.FindName("CancelButton")
        if run_btn is not None:
            run_btn.Click += self._on_run
        if cancel_btn is not None:
            cancel_btn.Click += self._on_cancel
        self.window.Closed += self._on_closed

    def _on_run(self, sender, args):
        self.request = self.viewmodel.to_request()
        self.window.DialogResult = True
        self.window.Close()

    def _on_cancel(self, sender, args):
        self.request = None
        self.window.DialogResult = False
        self.window.Close()

    def _on_closed(self, sender, args):
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
        except Exception:
            pass

    def show_dialog(self):
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, self.window)
        except Exception:
            pass
        ok = self.window.ShowDialog()
        if ok:
            return self.request or self.viewmodel.to_request()
        return None


def show_singleton_dialog():
    """Muestra la ventana o enfoca la existente si ya está abierta."""
    try:
        existing = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
        if existing is not None:
            try:
                existing.Activate()
                existing.Focus()
            except Exception:
                pass
            if TaskDialog is not None:
                try:
                    TaskDialog.Show(
                        "Arainco: Armado Columnas",
                        "La herramienta ya esta en ejecucion.",
                    )
                except Exception:
                    pass
            return None
    except Exception:
        pass
    return ColumnReinforcementWindowController().show_dialog()
