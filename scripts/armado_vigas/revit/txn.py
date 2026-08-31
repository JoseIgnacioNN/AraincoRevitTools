# -*- coding: utf-8 -*-
"""
Transacciones pyRevit 2025+ para Armado vigas.

- ``transaction_group_scope`` → ``revit.TransactionGroup`` (Assimilate = un Undo).
- ``transaction_scope`` → ``revit.Transaction`` (rollback limpio ante excepción).

Si pyRevit no está disponible, cae a ``Autodesk.Revit.DB.TransactionGroup`` /
``Transaction`` con el mismo contrato.
"""

from __future__ import print_function

from Autodesk.Revit.DB import Transaction, TransactionGroup

from armado_vigas.revit.rebar_failures import attach_rebar_outside_host_swallower


def _pyrevit_revit():
    try:
        from pyrevit import revit as _rv

        return _rv
    except Exception:
        return None


def _underlying_autodesk_txn(pyrevit_txn):
    """``DB.Transaction`` bajo el wrapper pyRevit, o el propio objeto si ya lo es."""
    if pyrevit_txn is None:
        return None
    if isinstance(pyrevit_txn, Transaction):
        return pyrevit_txn
    for attr in (u"_rvtxn", u"Transaction", u"transaction"):
        try:
            obj = getattr(pyrevit_txn, attr, None)
        except Exception:
            obj = None
        if obj is not None and isinstance(obj, Transaction):
            return obj
    return None


def attach_swallower_to_txn(txn_or_wrapper):
    """Adjunta el silenciador de rebar-fuera-de-host a la Transaction real."""
    return attach_rebar_outside_host_swallower(
        _underlying_autodesk_txn(txn_or_wrapper) or txn_or_wrapper
    )


class transaction_group_scope(object):
    """
    ``with revit.TransactionGroup(name, doc, assimilate=True)``.

    Excepción → ``RollBack`` del grupo (documento no queda bloqueado).
    Sin excepción + ``assimilate`` → un solo paso Deshacer.
    """

    def __init__(self, doc, name, assimilate=True):
        self._doc = doc
        self._name = name
        self._assimilate = bool(assimilate)
        self._cm = None
        self._tg = None
        self._owns_fallback = False

    def __enter__(self):
        _rv = _pyrevit_revit()
        if _rv is not None:
            self._cm = _rv.TransactionGroup(
                self._name, self._doc, assimilate=self._assimilate,
            )
            return self._cm.__enter__()
        self._tg = TransactionGroup(self._doc, self._name)
        self._tg.Start()
        self._owns_fallback = True
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._cm is not None:
            return self._cm.__exit__(exc_type, exc, tb)
        if not self._owns_fallback or self._tg is None:
            return False
        try:
            if exc_type is not None:
                try:
                    self._tg.RollBack()
                except Exception:
                    pass
            else:
                try:
                    if self._assimilate:
                        self._tg.Assimilate()
                    else:
                        self._tg.Commit()
                except Exception:
                    try:
                        self._tg.RollBack()
                    except Exception:
                        pass
                    raise
        finally:
            try:
                self._tg.Dispose()
            except Exception:
                pass
            self._tg = None
            self._owns_fallback = False
        return False


class transaction_scope(object):
    """
    ``with revit.Transaction(name, doc)`` + swallower opcional.

    Excepción → ``RollBack`` (pyRevit / DB). Fallo de Commit en respaldo →
    ``RollBack`` y se re-lanza. No deja el documento bloqueado.
    """

    def __init__(self, doc, name, swallow_outside_host=True):
        self._doc = doc
        self._name = name
        self._swallow = bool(swallow_outside_host)
        self._cm = None
        self._txn = None
        self._owns_fallback = False

    def __enter__(self):
        _rv = _pyrevit_revit()
        if _rv is not None:
            self._cm = _rv.Transaction(self._name, self._doc)
            entered = self._cm.__enter__()
            if self._swallow:
                try:
                    attach_swallower_to_txn(entered)
                except Exception:
                    pass
            return entered
        self._txn = Transaction(self._doc, self._name)
        if self._swallow:
            try:
                attach_rebar_outside_host_swallower(self._txn)
            except Exception:
                pass
        self._txn.Start()
        self._owns_fallback = True
        return self._txn

    def __exit__(self, exc_type, exc, tb):
        if self._cm is not None:
            return self._cm.__exit__(exc_type, exc, tb)
        if not self._owns_fallback or self._txn is None:
            return False
        try:
            if exc_type is not None:
                try:
                    if self._txn.HasStarted():
                        self._txn.RollBack()
                except Exception:
                    try:
                        self._txn.RollBack()
                    except Exception:
                        pass
            else:
                try:
                    self._txn.Commit()
                except Exception:
                    try:
                        if self._txn.HasStarted():
                            self._txn.RollBack()
                    except Exception:
                        pass
                    raise
        finally:
            try:
                self._txn.Dispose()
            except Exception:
                pass
            self._txn = None
            self._owns_fallback = False
        return False
