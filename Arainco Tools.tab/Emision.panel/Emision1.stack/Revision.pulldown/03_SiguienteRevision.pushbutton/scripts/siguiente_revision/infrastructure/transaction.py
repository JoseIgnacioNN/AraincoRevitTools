# -*- coding: utf-8 -*-
"""
Transaction Wrapper para Revit API.

Proporciona un context manager que abre, confirma y revierte automáticamente
transacciones de Revit, simplificando el manejo de errores en la capa de servicios.
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

from siguiente_revision.constants import TX_NAME


class RevitTransaction(object):
    """
    Context manager de transacción Revit con rollback automático ante excepciones.

    Uso::

        with RevitTransaction(doc) as tx:
            # operaciones Revit aquí
            pass   # Commit automático al salir sin excepción

        with RevitTransaction(doc, u"Arainco: Operación") as tx:
            element.param.Set(value)
    """

    def __init__(self, doc, name=None):
        self._doc = doc
        self._name = name or TX_NAME
        self._tx = None

    def __enter__(self):
        from Autodesk.Revit.DB import Transaction
        self._tx = Transaction(self._doc, self._name)
        self._tx.Start()
        return self._tx

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._tx is None:
            return False
        try:
            if exc_type is not None:
                if self._tx.HasStarted() and not self._tx.HasEnded():
                    self._tx.RollBack()
                return False
            if self._tx.HasStarted() and not self._tx.HasEnded():
                self._tx.Commit()
        except Exception:
            try:
                if self._tx.HasStarted() and not self._tx.HasEnded():
                    self._tx.RollBack()
            except Exception:
                pass
        return False

    @property
    def is_open(self):
        try:
            return self._tx is not None and self._tx.HasStarted() and not self._tx.HasEnded()
        except Exception:
            return False

    def rollback(self):
        """Rollback explícito fuera del context manager."""
        try:
            if self.is_open:
                self._tx.RollBack()
        except Exception:
            pass

    def commit(self):
        """Commit explícito fuera del context manager."""
        try:
            if self.is_open:
                self._tx.Commit()
        except Exception:
            pass
