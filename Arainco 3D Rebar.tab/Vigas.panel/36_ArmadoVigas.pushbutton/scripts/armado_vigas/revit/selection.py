# -*- coding: utf-8 -*-
"""Selección inicial: vigas, columnas y muros."""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from armado_vigas.ui.instruction_dialog import show_ok_cancel_dialog

_DIALOG = u"Arainco: Armado vigas"
_PICK_PROMPT = (
    u"Seleccione vigas a armar y apoyos relacionados (columnas, muros) · "
    u"Finalizar en la cinta o Esc para cancelar"
)

_ALLOWED = frozenset([
    int(BuiltInCategory.OST_StructuralFraming),
    int(BuiltInCategory.OST_StructuralColumns),
    int(BuiltInCategory.OST_Walls),
])


class ArmadoVigasSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        if elem is None:
            return False
        try:
            return int(elem.Category.Id.IntegerValue) in _ALLOWED
        except Exception:
            return False

    def AllowReference(self, ref, point):
        return False


def show_selection_instructions(uiapp=None):
    """
    Instrucciones previas a la selección en modelo (al ejecutar el botón).
    Devuelve ``True`` si el usuario pulsa Aceptar; ``False`` si cancela.
    """
    hwnd = None
    try:
        from revit_wpf_window_position import revit_main_hwnd

        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    return show_ok_cancel_dialog(
        _DIALOG,
        u"Seleccione las vigas a armar y los elementos de apoyo relacionados "
        u"(columnas y/o muros).",
        u"Pulse Aceptar para iniciar la selección en el modelo. "
        u"Finalice con la cinta (Finalizar) o cancela con Esc.",
        ok_text=u"Aceptar",
        cancel_text=u"Cancelar",
        hwnd_revit=hwnd,
    )


def pick_lote_inicial(uidoc):
    """
    Selección síncrona antes de abrir la UI (mismo filtro que el botón en ventana).
    Devuelve lista de ``Reference`` o ``None`` si el usuario cancela.
    """
    if uidoc is None:
        return None
    try:
        refs = list(
            uidoc.Selection.PickObjects(
                ObjectType.Element,
                ArmadoVigasSelectionFilter(),
                _PICK_PROMPT,
            )
        )
    except OperationCanceledException:
        return None
    except Exception:
        return None
    return refs


def validate_initial_selection(document, refs):
    """
    Comprueba que haya al menos una viga estructural en el lote.
    Devuelve ``(ok, message)``.
    """
    from armado_vigas.revit.adapters import elements_from_refs, framing_from_elements

    elems = elements_from_refs(document, refs)
    beams = framing_from_elements(elems)
    if not elems:
        return False, u"No seleccionó ningún elemento."
    if not beams:
        return False, (
            u"El lote no incluye vigas estructurales (Structural Framing — Beam). "
            u"Incluya las vigas a armar junto con sus apoyos."
        )
    return True, u""
