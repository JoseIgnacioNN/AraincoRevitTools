# -*- coding: utf-8 -*-
"""
Intercepta el comando nativo «Duplicate as a Dependent» del Project Browser.

Flujo recomendado (Revit 2024/2025 + pyRevit):
  hooks/command-before-exec[ID_CREATE_DEPENDENT_VIEW].py
  → cancela el comando nativo y pide cuántas vistas dependientes crear.

También puede registrarse manualmente desde startup.py (menos fiable tras reload).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Transaction, TransactionStatus, View, ViewDuplicateOption
from Autodesk.Revit.UI import (
    RevitCommandId,
    TaskDialog,
    TaskDialogCommandLinkId,
    TaskDialogCommonButtons,
    TaskDialogResult,
    ExternalEvent,
    IExternalEventHandler,
)
from System import EventHandler
from Autodesk.Revit.UI.Events import BeforeExecutedEventArgs

_COMMAND_ID = "ID_CREATE_DEPENDENT_VIEW"
_TXN = u"Arainco: Duplicar vistas dependientes"
_MAX_COUNT = 50
_DIALOG_TITLE = u"Arainco: Duplicar vistas dependientes"

_binding = None
_external_event = None
_pending_request = None
_uiapp = None


def _ask_dependent_view_count():
    """TaskDialog con enlaces 1–3 y opción para otra cantidad."""
    td = TaskDialog(_DIALOG_TITLE)
    td.MainInstruction = u"¿Cuántas vistas dependientes desea crear?"
    td.MainContent = (
        u"Se crearán desde la vista madre seleccionada en el Navegador de proyecto."
    )
    td.CommonButtons = TaskDialogCommonButtons.Cancel
    td.AddCommandLink(
        TaskDialogCommandLinkId.CommandLink1,
        u"1 vista dependiente",
        u"Crear una sola vista dependiente.",
    )
    td.AddCommandLink(
        TaskDialogCommandLinkId.CommandLink2,
        u"2 vistas dependientes",
        u"Crear dos vistas dependientes.",
    )
    td.AddCommandLink(
        TaskDialogCommandLinkId.CommandLink3,
        u"3 vistas dependientes",
        u"Crear tres vistas dependientes.",
    )
    td.AddCommandLink(
        TaskDialogCommandLinkId.CommandLink4,
        u"Otra cantidad…",
        u"Indicar un número personalizado (máx. {}).".format(_MAX_COUNT),
    )

    result = td.Show()
    if result == TaskDialogResult.Cancel:
        return None
    if result == TaskDialogResult.CommandLink1:
        return 1
    if result == TaskDialogResult.CommandLink2:
        return 2
    if result == TaskDialogResult.CommandLink3:
        return 3
    if result == TaskDialogResult.CommandLink4:
        return _ask_custom_count()
    return None


def _ask_custom_count():
    try:
        from pyrevit import forms
    except Exception:
        TaskDialog.Show(
            _DIALOG_TITLE,
            u"No se pudo abrir el cuadro de cantidad personalizada.",
        )
        return None

    raw = forms.ask_for_string(
        default="5",
        prompt=u"Cantidad de vistas dependientes (1–{}).".format(_MAX_COUNT),
        title=_DIALOG_TITLE,
    )
    if raw is None:
        return None
    raw = unicode(raw).strip()
    if not raw:
        return None
    try:
        count = int(raw)
    except Exception:
        TaskDialog.Show(_DIALOG_TITLE, u"La cantidad debe ser un número entero.")
        return None
    if count < 1 or count > _MAX_COUNT:
        TaskDialog.Show(
            _DIALOG_TITLE,
            u"Indique un valor entre 1 y {}.".format(_MAX_COUNT),
        )
        return None
    return count


def _resolve_parent_view(doc, element_id):
    if doc is None or element_id is None:
        return None
    try:
        eid = int(element_id.IntegerValue)
    except Exception:
        try:
            eid = int(element_id.Value)
        except Exception:
            return None
    if eid == 0:
        return None
    elem = doc.GetElement(element_id)
    if elem is None or not isinstance(elem, View):
        return None
    if elem.IsTemplate:
        return None
    return elem


def _create_dependent_views(doc, parent_view, count):
    if doc is None or parent_view is None or count < 1:
        return 0, u"No hay vista válida para duplicar."

    if not parent_view.CanViewBeDuplicated(ViewDuplicateOption.AsDependent):
        return 0, (
            u"La vista «{}» no admite duplicación como dependiente.\n\n"
            u"Seleccione una vista madre (no una vista dependiente ni un tipo "
            u"no compatible)."
        ).format(parent_view.Name)

    created = 0
    txn = Transaction(doc, _TXN)
    txn.Start()
    try:
        for _ in range(int(count)):
            new_id = parent_view.Duplicate(ViewDuplicateOption.AsDependent)
            if new_id is None:
                raise Exception(u"Revit no devolvió un Id de vista nueva.")
            try:
                ok = int(new_id.IntegerValue) != 0
            except Exception:
                ok = int(new_id.Value) != 0
            if not ok:
                raise Exception(u"Revit no devolvió un Id de vista nueva.")
            created += 1
        if txn.Commit() != TransactionStatus.Committed:
            return 0, u"No se pudo confirmar la transacción."
    except Exception as ex:
        if txn.HasStarted() and not txn.HasEnded():
            txn.RollBack()
        return 0, unicode(ex)

    return created, None


class _DependentViewDuplicateHandler(IExternalEventHandler):
    def GetName(self):
        return u"Arainco: duplicar vistas dependientes (intercept)"

    def Execute(self, uiapp):
        global _pending_request
        request = _pending_request
        _pending_request = None
        if request is None:
            return

        uidoc = None
        if uiapp is not None:
            try:
                uidoc = uiapp.ActiveUIDocument
            except Exception:
                uidoc = None
        if uidoc is None:
            TaskDialog.Show(_DIALOG_TITLE, u"No hay documento activo.")
            return

        doc = uidoc.Document
        view_id = request.get("view_id")
        count = request.get("count")
        parent_view = _resolve_parent_view(doc, view_id)
        if parent_view is None:
            TaskDialog.Show(
                _DIALOG_TITLE,
                u"No se encontró la vista madre seleccionada.",
            )
            return

        created, err = _create_dependent_views(doc, parent_view, count)
        if err:
            TaskDialog.Show(_DIALOG_TITLE, err)
            return

        TaskDialog.Show(
            _DIALOG_TITLE,
            u"Se crearon {} vista(s) dependiente(s) desde «{}».".format(
                created,
                parent_view.Name,
            ),
        )


def _ensure_external_event():
    global _external_event
    if _external_event is None:
        _external_event = ExternalEvent.Create(_DependentViewDuplicateHandler())
    return _external_event


def _get_uiapp(fallback=None):
    global _uiapp
    if _uiapp is not None:
        return _uiapp
    if fallback is not None:
        try:
            # UIApplication tiene ActiveUIDocument
            _ = fallback.ActiveUIDocument
            _uiapp = fallback
            return _uiapp
        except Exception:
            pass
    try:
        from pyrevit import HOST_APP

        uiapp = getattr(HOST_APP, "uiapp", None)
        if uiapp is None:
            uiapp = getattr(HOST_APP, "app", None)
        if uiapp is not None:
            _uiapp = uiapp
            return _uiapp
    except Exception:
        pass
    return None


def _resolve_uidoc_and_doc(uiapp_hint=None, args=None):
    uiapp = _get_uiapp(uiapp_hint)
    uidoc = None
    if uiapp is not None:
        try:
            uidoc = uiapp.ActiveUIDocument
        except Exception:
            uidoc = None

    doc = None
    if uidoc is not None:
        doc = uidoc.Document
    if doc is None and args is not None:
        try:
            doc = args.ActiveDocument
        except Exception:
            doc = None
    return uiapp, uidoc, doc


def handle_command_before_executed(uiapp, args):
    """
    Entrada del hook pyRevit command-before-exec.

    Cancela siempre el comando nativo y, si hay vista válida, pide cantidad
    y crea las dependientes vía ExternalEvent (transacciones seguras).
    """
    global _pending_request, _uiapp

    # Cancelar el duplicado nativo (1 sola vista) siempre que el hook corra.
    try:
        args.Cancel = True
    except Exception:
        pass

    try:
        if uiapp is not None:
            _uiapp = uiapp

        uiapp, uidoc, doc = _resolve_uidoc_and_doc(uiapp, args)
        if doc is None or uidoc is None:
            TaskDialog.Show(_DIALOG_TITLE, u"No hay documento activo.")
            return

        sel_ids = list(uidoc.Selection.GetElementIds())
        if len(sel_ids) != 1:
            TaskDialog.Show(
                _DIALOG_TITLE,
                u"Seleccione exactamente una vista madre en el Navegador de proyecto.",
            )
            return

        parent_view = _resolve_parent_view(doc, sel_ids[0])
        if parent_view is None:
            TaskDialog.Show(
                _DIALOG_TITLE,
                u"La selección no es una vista válida para duplicar como dependiente.",
            )
            return

        count = _ask_dependent_view_count()
        if not count:
            return

        _pending_request = {
            "view_id": parent_view.Id,
            "count": int(count),
        }
        _ensure_external_event().Raise()
    except Exception as ex:
        _pending_request = None
        TaskDialog.Show(
            _DIALOG_TITLE,
            u"Error al preparar la duplicación:\n{}".format(unicode(ex)),
        )


def _on_duplicate_dependent_executed(sender, args):
    """Respaldo si se registra binding Executed desde startup."""
    handle_command_before_executed(_get_uiapp(sender), args)


def register_dependent_view_duplicate_intercept(uiapp):
    """
    Registro manual opcional (startup). Preferir el hook pyRevit.

    Usa BeforeExecuted + Cancel (no Executed) para no pelear con hooks.
    """
    global _binding, _uiapp

    if uiapp is None:
        return False

    _uiapp = uiapp

    if _binding is not None:
        return True

    cmd_id = RevitCommandId.LookupCommandId(_COMMAND_ID)
    if cmd_id is None:
        print("BIMTools: no se encontró el comando {}".format(_COMMAND_ID))
        return False

    try:
        uiapp.RemoveAddInCommandBinding(cmd_id)
    except Exception:
        pass

    try:
        _binding = uiapp.CreateAddInCommandBinding(cmd_id)
    except Exception as ex:
        # Ya enlazado por hook pyRevit u otro add-in: OK, el hook cubre el caso.
        print(
            "BIMTools: binding {} no creado (puede estar el hook): {}".format(
                _COMMAND_ID,
                unicode(ex),
            )
        )
        return False

    _binding.BeforeExecuted += EventHandler[BeforeExecutedEventArgs](
        lambda s, a: handle_command_before_executed(_uiapp, a)
    )
    return True


def unregister_dependent_view_duplicate_intercept(uiapp):
    """Quita el binding manual (no elimina el hook de pyRevit)."""
    global _binding, _pending_request, _external_event, _uiapp

    if uiapp is None:
        return

    cmd_id = RevitCommandId.LookupCommandId(_COMMAND_ID)
    if cmd_id is not None:
        try:
            uiapp.RemoveAddInCommandBinding(cmd_id)
        except Exception:
            pass

    _binding = None
    _pending_request = None
    _external_event = None
    _uiapp = None
