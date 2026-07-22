# -*- coding: utf-8 -*-
"""Adapter de transacciones Revit.

Los nombres deben conservar el prefijo obligatorio `Arainco: `.

Stubs / preparación: sin referencias runtime actuales. Ver ``DEPENDENCIES.md``.
"""


class TransactionRunner(object):
    """Punto de extensión para mover transacciones fuera del script legado."""

    def __init__(self, doc, transaction_cls):
        self.doc = doc
        self.transaction_cls = transaction_cls

    def run(self, name, action):
        tx_name = name if str(name).startswith("Arainco: ") else "Arainco: " + str(name)
        txn = self.transaction_cls(self.doc, tx_name)
        txn.Start()
        try:
            result = action()
            txn.Commit()
            return result
        except Exception:
            if txn.HasStarted():
                txn.RollBack()
            raise
