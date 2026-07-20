# -*- coding: utf-8 -*-
"""
Transacciones anidables para Armado Muros v2.

Si el documento ya está en una Transaction abierta, abre ``SubTransaction``
(rollback local barato). Si no, abre ``Transaction`` completa.

Uso típico: envolver un lote en una Transaction y dejar que las operaciones
por rebar usen estas helpers → un solo commit de documento.

También adjunta un ``IFailuresPreprocessor`` que silencia warnings conocidos
de rebar completamente fuera del host (evita el diálogo de Revit al Commit).
"""

from __future__ import print_function

from Autodesk.Revit.DB import (
    BuiltInFailures,
    FailureProcessingResult,
    FailureSeverity,
    IFailuresPreprocessor,
    SubTransaction,
    Transaction,
)

_KIND_TXN = u"txn"
_KIND_SUB = u"sub"

_OUTSIDE_HOST_IDS = None


def _outside_host_failure_ids():
    """FailureDefinitionId de rebar/container fuera del host (según versión API)."""
    global _OUTSIDE_HOST_IDS
    if _OUTSIDE_HOST_IDS is not None:
        return _OUTSIDE_HOST_IDS
    ids = []
    try:
        rf = BuiltInFailures.RebarFailures
    except Exception:
        rf = None
    if rf is not None:
        for attr in (
            u"OutSideOfHost",
            u"RebarOutSideOfHost",
            u"RebarContainerOutSideOfHostWarning",
        ):
            try:
                fid = getattr(rf, attr, None)
                if fid is not None:
                    ids.append(fid)
            except Exception:
                pass
    _OUTSIDE_HOST_IDS = ids
    return _OUTSIDE_HOST_IDS


def _failure_message_looks_outside_host(fmsg):
    """Respaldo por texto si el FailureDefinitionId no está en la API."""
    try:
        desc = fmsg.GetDescriptionText() or u""
    except Exception:
        return False
    try:
        low = unicode(desc).lower()
    except Exception:
        try:
            low = str(desc).lower()
        except Exception:
            return False
    # ES / EN típicos de Revit
    markers = (
        u"completamente fuera",
        u"completely outside",
        u"outside of its host",
        u"fuera de su anfitrión",
        u"fuera de su host",
    )
    for m in markers:
        if m in low:
            return True
    return False


class RebarOutsideHostWarningSwallower(IFailuresPreprocessor):
    """
    Elimina warnings de rebar (o contenedor) colocado completamente fuera del host.
    No toca errores ni otros warnings.
    """

    def _iter_failure_msgs(self, failures_accessor):
        if failures_accessor is None:
            return
        try:
            fmsgs = failures_accessor.GetFailureMessages()
        except Exception:
            return
        if fmsgs is None:
            return
        try:
            n = int(fmsgs.Count)
        except Exception:
            n = 0
        for i in range(n):
            f = None
            try:
                f = fmsgs.get_Item(i)
            except Exception:
                try:
                    f = fmsgs[i]
                except Exception:
                    f = None
            if f is not None:
                yield f

    def PreprocessFailures(self, failures_accessor):
        if failures_accessor is None:
            return FailureProcessingResult.Continue
        known = _outside_host_failure_ids()
        for f in self._iter_failure_msgs(failures_accessor):
            try:
                if f.GetSeverity() != FailureSeverity.Warning:
                    continue
            except Exception:
                continue
            delete = False
            try:
                fid = f.GetFailureDefinitionId()
            except Exception:
                fid = None
            if fid is not None and known:
                for kid in known:
                    try:
                        if fid == kid:
                            delete = True
                            break
                    except Exception:
                        pass
            if not delete:
                delete = _failure_message_looks_outside_host(f)
            if delete:
                try:
                    failures_accessor.DeleteWarning(f)
                except Exception:
                    pass
        return FailureProcessingResult.Continue


def attach_rebar_outside_host_swallower(txn):
    """
    Adjunta el preprocessor a una ``Transaction`` (no aplica a SubTransaction).
    Retorna True si se pudo configurar.
    """
    if txn is None or not isinstance(txn, Transaction):
        return False
    try:
        opts = txn.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(RebarOutsideHostWarningSwallower())
        txn.SetFailureHandlingOptions(opts)
        return True
    except Exception:
        return False


def doc_is_modifiable(doc):
    try:
        return bool(doc.IsModifiable)
    except Exception:
        return False


def start_transaction(doc, name):
    """
    Inicia Transaction o SubTransaction.

    :returns: handle ``(kind, obj)`` o ``None`` si no se pudo iniciar.
    """
    if doc is None:
        return None
    if doc_is_modifiable(doc):
        try:
            st = SubTransaction(doc)
            st.Start()
            return (_KIND_SUB, st)
        except Exception:
            return None
    try:
        t = Transaction(doc, name)
        attach_rebar_outside_host_swallower(t)
        t.Start()
        return (_KIND_TXN, t)
    except Exception:
        return None


def commit_transaction(handle):
    if handle is None:
        return
    try:
        kind, obj = handle
    except Exception:
        return
    try:
        if kind == _KIND_SUB:
            obj.Commit()
        else:
            obj.Commit()
    except Exception:
        rollback_transaction(handle)
        raise


def rollback_transaction(handle):
    if handle is None:
        return
    try:
        kind, obj = handle
    except Exception:
        return
    try:
        if kind == _KIND_SUB:
            obj.RollBack()
        else:
            if obj.HasStarted():
                obj.RollBack()
    except Exception:
        pass


def run_in_transaction(doc, name, fn):
    """
    Ejecuta ``fn()`` dentro de una Transaction de documento (no SubTransaction).
    Si ya hay txn abierta, ejecuta ``fn`` sin abrir otra (caller es dueño).
    """
    if doc is None:
        return fn()
    if doc_is_modifiable(doc):
        return fn()
    t = Transaction(doc, name)
    attach_rebar_outside_host_swallower(t)
    t.Start()
    try:
        result = fn()
        t.Commit()
        return result
    except Exception:
        if t.HasStarted():
            try:
                t.RollBack()
            except Exception:
                pass
        raise


class TxnScope(object):
    """Ámbito Transaction/SubTransaction con commit/rollback explícitos."""

    def __init__(self, doc, name):
        self.handle = start_transaction(doc, name)

    def commit(self):
        commit_transaction(self.handle)
        self.handle = None

    def rollback(self):
        rollback_transaction(self.handle)
        self.handle = None
