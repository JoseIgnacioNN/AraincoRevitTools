# -*- coding: utf-8 -*-
"""Selección inicial: vigas, columnas y muros de hormigón."""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from armado_vigas.ui.instruction_dialog import show_ok_cancel_dialog

_DIALOG = u"Arainco: Armado vigas"
_PICK_PROMPT = (
    u"Seleccione vigas de hormigón paralelas a la vista activa y apoyos de hormigón "
    u"(columnas, muros) · Finalizar en la cinta o Esc para cancelar"
)

_ALLOWED = frozenset([
    int(BuiltInCategory.OST_StructuralFraming),
    int(BuiltInCategory.OST_StructuralColumns),
    int(BuiltInCategory.OST_Walls),
])
_FRAMING_CAT = int(BuiltInCategory.OST_StructuralFraming)


def _categoria_permitida(elem):
    try:
        return int(elem.Category.Id.IntegerValue) in _ALLOWED
    except Exception:
        return False


def _es_framing(elem):
    try:
        return int(elem.Category.Id.IntegerValue) == _FRAMING_CAT
    except Exception:
        return False


def _es_hormigon(elem):
    """Material for Model Behavior = Concrete (``StructuralMaterialType`` + respaldos)."""
    try:
        from geometria_colision_vigas import material_estructural_es_concrete

        return material_estructural_es_concrete(elem)
    except Exception:
        return False


def _viga_paralela_a_vista(elem, view):
    if not _es_framing(elem):
        return True
    try:
        from armado_vigas.revit.view_order import beam_axis_parallel_to_view_plane

        return beam_axis_parallel_to_view_plane(elem, view)
    except Exception:
        return False


class ArmadoVigasSelectionFilter(ISelectionFilter):
    def __init__(self, view=None):
        self._view = view

    def AllowElement(self, elem):
        if elem is None:
            return False
        if not _categoria_permitida(elem):
            return False
        if not _es_hormigon(elem):
            return False
        return _viga_paralela_a_vista(elem, self._view)

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
        u"Seleccione las vigas de hormigón a armar y los apoyos de hormigón "
        u"relacionados (columnas y/o muros).",
        u"Solo se permiten elementos con Material for Model Behavior = Concrete.\n"
        u"Las vigas (Structural Framing) deben tener su eje paralelo al plano de "
        u"la vista activa.\n\n"
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
    view = uidoc.ActiveView
    try:
        refs = list(
            uidoc.Selection.PickObjects(
                ObjectType.Element,
                ArmadoVigasSelectionFilter(view),
                _PICK_PROMPT,
            )
        )
    except OperationCanceledException:
        return None
    except Exception:
        return None
    return refs


def validate_initial_selection(document, refs, view=None):
    """
    Comprueba que haya al menos una viga estructural de hormigón en el lote.
    Devuelve ``(ok, message)``.
    """
    from armado_vigas.revit.adapters import elements_from_refs, framing_from_elements

    elems = elements_from_refs(document, refs)
    if not elems:
        return False, u"No seleccionó ningún elemento."
    for el in elems:
        if not _es_hormigon(el):
            return False, (
                u"El lote incluye elementos que no son de hormigón "
                u"(Material for Model Behavior ≠ Concrete). "
                u"Solo se permiten vigas, columnas y muros de hormigón."
            )
        if not _viga_paralela_a_vista(el, view):
            return False, (
                u"El lote incluye vigas cuyo eje no es paralelo al plano de la vista activa.\n\n"
                u"Solo se pueden armar vigas visibles en la vista actual "
                u"(eje paralelo al corte, alzado o planta). "
                u"Las vigas en punta o no alineadas con la vista quedan excluidas."
            )
    beams = framing_from_elements(elems)
    if not beams:
        return False, (
            u"El lote no incluye vigas estructurales de hormigón "
            u"(Structural Framing — Beam). "
            u"Incluya las vigas a armar junto con sus apoyos de hormigón."
        )
    return True, u""
