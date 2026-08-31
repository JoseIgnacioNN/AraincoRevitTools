# -*- coding: utf-8 -*-
"""
Silencia warnings de Revit: rebar colocado **completamente fuera del host**.

Aparece con estirón post-fusión (muro/viga/columna) y patas L que salen del
sólido de la viga. Mismo criterio que ``armado_muros_txn``.
"""

from __future__ import print_function

from Autodesk.Revit.DB import (
    BuiltInFailures,
    FailureProcessingResult,
    FailureSeverity,
    IFailuresPreprocessor,
    Transaction,
)

_OUTSIDE_HOST_IDS = None
_SWALLOWER_SINGLETON = None


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
    markers = (
        u"completamente fuera",
        u"completely outside",
        u"outside of its host",
        u"outside of the host",
        u"outside of host",
        u"fuera de su anfitrión",
        u"fuera de su host",
        u"fuera del host",
        u"fuera del anfitrión",
    )
    for m in markers:
        if m in low:
            return True
    return False


class RebarOutsideHostWarningSwallower(IFailuresPreprocessor):
    """Elimina solo warnings «rebar outside of its host»; no toca errores."""

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
        to_delete = []
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
                to_delete.append(f)
        for f in to_delete:
            try:
                failures_accessor.DeleteWarning(f)
            except Exception:
                pass
        return FailureProcessingResult.Continue


def _as_db_transaction(txn):
    """Acepta ``DB.Transaction`` o wrapper pyRevit (``_rvtxn`` / similares)."""
    if txn is None:
        return None
    if isinstance(txn, Transaction):
        return txn
    for attr in (u"_rvtxn", u"Transaction", u"transaction"):
        try:
            obj = getattr(txn, attr, None)
        except Exception:
            obj = None
        if obj is not None and isinstance(obj, Transaction):
            return obj
    return None


def attach_rebar_outside_host_swallower(txn):
    """
    Adjunta el preprocessor a una ``Transaction`` (o wrapper pyRevit).

    Usa un singleton CLR: recrear el wrapper IronPython por cada corrida
    puede tumbar Revit en la 2.ª ejecución.
    """
    global _SWALLOWER_SINGLETON
    db_txn = _as_db_transaction(txn)
    if db_txn is None:
        return False
    # Preferir util compartida si está cargada (misma instancia de silenciador).
    try:
        from armado_muros_txn import attach_rebar_outside_host_swallower as _shared

        return bool(_shared(db_txn))
    except Exception:
        pass
    try:
        if _SWALLOWER_SINGLETON is None:
            _SWALLOWER_SINGLETON = RebarOutsideHostWarningSwallower()
        opts = db_txn.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(_SWALLOWER_SINGLETON)
        db_txn.SetFailureHandlingOptions(opts)
        return True
    except Exception:
        return False
