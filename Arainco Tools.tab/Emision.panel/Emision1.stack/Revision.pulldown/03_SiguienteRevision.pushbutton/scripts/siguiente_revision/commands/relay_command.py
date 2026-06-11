# -*- coding: utf-8 -*-
"""
RelayCommand — implementación de ICommand para IronPython + WPF.

Permite exponer acciones como comandos WPF desde el ViewModel,
con soporte opcional de CanExecute para habilitar/deshabilitar controles.
"""

from __future__ import print_function

import clr
clr.AddReference("WindowsBase")

from System.Windows.Input import ICommand


class RelayCommand(ICommand):
    """
    Implementación minimalista de ICommand para IronPython.

    IronPython mantiene referencias débiles a métodos instanciados; se guardan
    las funciones en atributos de instancia para evitar que el GC las libere.

    Uso::

        self.ok_command     = RelayCommand(self._on_ok)
        self.cancel_command = RelayCommand(self._on_cancel, lambda _: self._can_cancel())
    """

    def __init__(self, execute_fn, can_execute_fn=None):
        self._execute_fn = execute_fn
        self._can_execute_fn = can_execute_fn
        self._can_execute_changed_handlers = []

    # --- ICommand interface ---

    def add_CanExecuteChanged(self, handler):
        self._can_execute_changed_handlers.append(handler)

    def remove_CanExecuteChanged(self, handler):
        try:
            self._can_execute_changed_handlers.remove(handler)
        except ValueError:
            pass

    def CanExecute(self, parameter):
        if self._can_execute_fn is None:
            return True
        try:
            return bool(self._can_execute_fn(parameter))
        except Exception:
            return True

    def Execute(self, parameter):
        try:
            self._execute_fn(parameter)
        except Exception:
            pass

    def raise_can_execute_changed(self):
        """Notifica a WPF que CanExecute debe reevaluarse."""
        for handler in list(self._can_execute_changed_handlers):
            try:
                handler(self, None)
            except Exception:
                pass
