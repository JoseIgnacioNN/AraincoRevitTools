# -*- coding: utf-8 -*-
"""
Base ViewModel para IronPython + WPF.

Implementa INotifyPropertyChanged del framework .NET para permitir que WPF
reaccione a cambios de propiedades Python mediante bindings declarativos.
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import clr
clr.AddReference("System")
clr.AddReference("WindowsBase")

from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs


class ObservableObject(INotifyPropertyChanged):
    """
    Clase base para ViewModels con soporte de notificación de cambios (INotifyPropertyChanged).

    IronPython implementa interfaces .NET mediante herencia directa.
    Los event handlers se mantienen en una lista Python para evitar que
    el GC de IronPython libere referencias a bound methods.

    Uso::

        class MiVM(ObservableObject):
            def __init__(self):
                super(MiVM, self).__init__()
                self._valor = u""

            @property
            def valor(self):
                return self._valor

            @valor.setter
            def valor(self, v):
                self.set_property(u"_valor", v, u"valor")
    """

    def __init__(self):
        self._property_changed_handlers = []

    # --- INotifyPropertyChanged interface ---

    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        try:
            self._property_changed_handlers.remove(handler)
        except ValueError:
            pass

    def notify(self, prop_name):
        """Dispara PropertyChanged para el nombre de propiedad dado."""
        args = PropertyChangedEventArgs(prop_name)
        for handler in list(self._property_changed_handlers):
            try:
                handler(self, args)
            except Exception:
                pass

    def set_property(self, attr_name, new_value, prop_name):
        """
        Asigna new_value a self.attr_name y notifica si el valor cambia.

        Args:
            attr_name:  nombre del atributo privado (ej. u"_titulo").
            new_value:  valor nuevo.
            prop_name:  nombre de la propiedad pública (para PropertyChanged).
        """
        if getattr(self, attr_name, None) != new_value:
            setattr(self, attr_name, new_value)
            self.notify(prop_name)
