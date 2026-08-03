# -*- coding: utf-8 -*-
"""
Transacciones anidables para Armado Muros v3 (Revit 2025+ / pyRevit).

Si el documento ya está en una Transaction abierta, abre ``SubTransaction``
(rollback local barato). Si no, abre ``Transaction`` completa.

Preferir los context managers nativos de pyRevit::

    with transaction_group_scope(doc, u"Arainco: …", assimilate=True):
        with transaction_scope(doc, u"Arainco: …"):
            ...

``transaction_group_scope`` → ``revit.TransactionGroup`` (Assimilate = un Undo).
``transaction_scope`` → ``revit.Transaction`` (SubTransaction si ya hay txn;
excepción → RollBack limpio).

``batch_mutation_scope`` agrupa muchas mutaciones en una sola SubTransaction/
Transaction; mientras está activo, ``TxnScope`` anidados son no-op (evita
O(rebars) SubTransactions en el post-proceso de mallas).

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
    TransactionGroup,
)

_KIND_TXN = u"txn"
_KIND_SUB = u"sub"
_KIND_NOOP = u"noop"

# Profundidad de ``batch_mutation_scope``: si > 0, ``TxnScope`` / ``start_transaction``
# no abren SubTransaction anidada (mutaciones van al batch abierto).
_BATCH_MUTATION_DEPTH = 0

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
    # ES / EN típicos de Revit
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
        # Recolectar primero: DeleteWarning mientras se itera GetFailureMessages
        # puede tumbar Revit (sobre todo con muchos warnings de capas nuevas).
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


def attach_rebar_outside_host_swallower(txn):
    """
    Adjunta el preprocessor a una ``Transaction`` (no aplica a SubTransaction).
    Retorna True si se pudo configurar.

    Reutiliza una sola instancia CLR (crear N wrappers IronPython del
    ``IFailuresPreprocessor`` y adjuntarlos en sucesivas ejecuciones puede
    tumbar Revit en la 2ª apertura de la herramienta).
    """
    global _SWALLOWER_SINGLETON
    if txn is None or not isinstance(txn, Transaction):
        return False
    try:
        if _SWALLOWER_SINGLETON is None:
            _SWALLOWER_SINGLETON = RebarOutsideHostWarningSwallower()
        opts = txn.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(_SWALLOWER_SINGLETON)
        txn.SetFailureHandlingOptions(opts)
        return True
    except Exception:
        return False


def _pyrevit_revit():
    """Módulo ``pyrevit.revit`` (Transaction / TransactionGroup nativos)."""
    from pyrevit import revit as _rv

    return _rv


def _underlying_autodesk_txn(pyrevit_txn):
    """``DB.Transaction`` / ``DB.SubTransaction`` bajo el wrapper pyRevit."""
    if pyrevit_txn is None:
        return None
    for attr in (u"_rvtxn", u"Transaction", u"transaction"):
        try:
            obj = getattr(pyrevit_txn, attr, None)
        except Exception:
            obj = None
        if obj is not None:
            return obj
    return None


def attach_swallower_to_pyrevit_txn(pyrevit_txn):
    """Adjunta el swallower al ``DB.Transaction`` del context manager pyRevit."""
    return attach_rebar_outside_host_swallower(
        _underlying_autodesk_txn(pyrevit_txn),
    )


def doc_is_modifiable(doc):
    try:
        return bool(doc.IsModifiable)
    except Exception:
        return False


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
        _rv = None
        try:
            _rv = _pyrevit_revit()
        except Exception:
            _rv = None
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
    ``with revit.Transaction(name, doc)`` + swallower de rebar fuera de host.

    Si el documento ya es modificable → ``SubTransaction`` (pyRevit nested).
    Excepción → ``RollBack`` (pyRevit / respaldo Autodesk). Fallo de Commit
    en el respaldo → ``RollBack`` y se re-lanza.
    """

    def __init__(self, doc, name, swallow_outside_host=True):
        self._doc = doc
        self._name = name
        self._swallow = bool(swallow_outside_host)
        self._cm = None
        self._fallback = None
        self._owns_fallback = False

    def __enter__(self):
        _rv = None
        try:
            _rv = _pyrevit_revit()
        except Exception:
            _rv = None
        if _rv is not None:
            self._cm = _rv.Transaction(self._name, self._doc)
            entered = self._cm.__enter__()
            if self._swallow:
                try:
                    attach_swallower_to_pyrevit_txn(entered)
                except Exception:
                    pass
            return entered
        handle = start_transaction(self._doc, self._name)
        if handle is None:
            raise Exception(
                u"No se pudo abrir transacción: {0}".format(self._name),
            )
        self._fallback = handle
        self._owns_fallback = True
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._cm is not None:
            return self._cm.__exit__(exc_type, exc, tb)
        if not self._owns_fallback:
            return False
        try:
            if exc_type is not None:
                rollback_transaction(self._fallback)
            else:
                try:
                    commit_transaction(self._fallback)
                except Exception:
                    rollback_transaction(self._fallback)
                    raise
        finally:
            self._fallback = None
            self._owns_fallback = False
        return False


def start_transaction(doc, name):
    """
    Inicia Transaction o SubTransaction.

    Bajo ``batch_mutation_scope`` (doc ya modificable): retorna handle no-op
    para no anidar SubTransactions por mutación.

    :returns: handle ``(kind, obj)`` o ``None`` si no se pudo iniciar.
    """
    if doc is None:
        return None
    if _BATCH_MUTATION_DEPTH > 0 and doc_is_modifiable(doc):
        return (_KIND_NOOP, None)
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
    if kind == _KIND_NOOP:
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
    if kind == _KIND_NOOP:
        return
    try:
        if kind == _KIND_SUB:
            obj.RollBack()
        else:
            if obj.HasStarted():
                obj.RollBack()
    except Exception:
        pass


class _BatchMutationScope(object):
    """
    Una sola SubTransaction/Transaction para un bloque de mutaciones.

    Mientras está activo, ``TxnScope`` anidados son no-op (sin SubTxn por barra).
    """

    def __init__(self, doc, name):
        self._doc = doc
        self._name = name
        self._handle = None
        self._owns_handle = False
        self._entered = False

    def __enter__(self):
        global _BATCH_MUTATION_DEPTH
        if _BATCH_MUTATION_DEPTH > 0:
            # Ya hay un batch externo: solo anidar contador.
            _BATCH_MUTATION_DEPTH += 1
            self._entered = True
            return self
        # Abrir batch real con profundidad 0 (start_transaction no ve batch aún).
        self._handle = start_transaction(self._doc, self._name)
        self._owns_handle = self._handle is not None
        _BATCH_MUTATION_DEPTH += 1
        self._entered = True
        return self

    def __exit__(self, exc_type, exc, tb):
        global _BATCH_MUTATION_DEPTH
        if not self._entered:
            return False
        try:
            if self._owns_handle:
                if exc_type is not None:
                    rollback_transaction(self._handle)
                else:
                    try:
                        commit_transaction(self._handle)
                    except Exception:
                        rollback_transaction(self._handle)
                        raise
        finally:
            if _BATCH_MUTATION_DEPTH > 0:
                _BATCH_MUTATION_DEPTH -= 1
            self._handle = None
            self._owns_handle = False
            self._entered = False
        return False


def batch_mutation_scope(doc, name):
    """
    Context manager: una SubTransaction (o Transaction) para muchas mutaciones.

    Uso::

        with batch_mutation_scope(doc, u\"Arainco: …\"):
            # TxnScope internos no abren SubTxn adicionales
            ...
    """
    return _BatchMutationScope(doc, name)


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
    """
    Ámbito Transaction/SubTransaction con commit/rollback explícitos.

    Si el documento ya es modificable (txn padre abierta), abre SubTransaction;
    si no, abre Transaction de documento. Compatible con el patrón API
    ``Commit`` / ``RollBack`` / ``HasStarted`` de ``Transaction``.

    Bajo ``batch_mutation_scope``, el handle es no-op (sin SubTxn anidada).

    También usable como context manager (commit / rollback automático)::

        with TxnScope(doc, u\"Arainco: …\") as t:
            ...
    """

    def __init__(self, doc, name):
        self.handle = start_transaction(doc, name)

    def __enter__(self):
        if self.handle is None:
            raise Exception(u"No se pudo abrir transacción.")
        return self

    def __exit__(self, exc_type, exc, tb):
        if self.handle is None:
            return False
        if exc_type is not None:
            self.rollback()
        else:
            try:
                self.commit()
            except Exception:
                self.rollback()
                raise
        return False

    def has_started(self):
        return self.handle is not None

    def has_ended(self):
        return self.handle is None

    def commit(self):
        commit_transaction(self.handle)
        self.handle = None

    def rollback(self):
        rollback_transaction(self.handle)
        self.handle = None

    # Alias estilo Autodesk.Revit.DB.Transaction (drop-in en call sites).
    def HasStarted(self):
        return self.has_started()

    def HasEnded(self):
        return self.has_ended()

    def Commit(self):
        return self.commit()

    def RollBack(self):
        return self.rollback()


def open_txn(doc, name):
    """
    Abre Transaction o SubTransaction (ya iniciada).

    Sustituto de ``Transaction(doc, name)`` + ``Start()`` + swallower.
    """
    return TxnScope(doc, name)
