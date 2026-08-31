# -*- coding: utf-8 -*-
"""Selección de columnas, fundaciones, vigas y losas desde la vista activa."""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

_DIALOG = u"Arainco: Armado columnas V2"
_PICK_PROMPT = (
    u"Seleccione columnas, fundaciones, vigas y losas de hormigón · "
    u"Finalizar en la cinta o Esc para cancelar"
)

_COL_CAT = int(BuiltInCategory.OST_StructuralColumns)
_FND_CAT = int(BuiltInCategory.OST_StructuralFoundation)
_FRM_CAT = int(BuiltInCategory.OST_StructuralFraming)
_FLR_CAT = int(BuiltInCategory.OST_Floors)

_ALLOWED = frozenset([_COL_CAT, _FND_CAT, _FRM_CAT, _FLR_CAT])


def _categoria_permitida(elem):
    try:
        return int(elem.Category.Id.IntegerValue) in _ALLOWED
    except Exception:
        return False


def _es_hormigon(elem):
    """Material for Model Behavior = Concrete. Si no se puede comprobar → False."""
    if elem is None:
        return False
    try:
        from contorno_material_concrete import material_estructural_es_concrete

        return bool(material_estructural_es_concrete(elem))
    except Exception:
        pass
    try:
        from geometria_colision_vigas import material_estructural_es_concrete

        return bool(material_estructural_es_concrete(elem))
    except Exception:
        return False


class ArmadoColumnasSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        if elem is None:
            return False
        if not _categoria_permitida(elem):
            return False
        return _es_hormigon(elem)

    def AllowReference(self, ref, point):
        return False


def show_selection_instructions(uiapp=None):
    """Instrucciones previas. ``True`` = Aceptar; ``False`` = cancelar."""
    hwnd = None
    try:
        from revit_wpf_window_position import revit_main_hwnd

        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass

    try:
        from armado_columnas_instruction_dialog import show_ok_cancel_dialog

        return show_ok_cancel_dialog(
            _DIALOG,
            u"Seleccione en la vista activa solo elementos de hormigón: "
            u"columnas, fundaciones, vigas y losas.",
            u"Solo se permiten elementos con Material for Model Behavior = Concrete.\n"
            u"Se dibujan con la misma forma, tamaño y posición que en la vista "
            u"(RightDirection → horizontal, UpDirection → vertical).\n\n"
            u"Pulse Aceptar para iniciar la selección. Finalice con la cinta "
            u"(Finalizar) o cancele con Esc.",
            ok_text=u"Aceptar",
            cancel_text=u"Cancelar",
            hwnd_revit=hwnd,
        )
    except Exception:
        from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult

        r = TaskDialog.Show(
            _DIALOG,
            u"Seleccione columnas, fundaciones, vigas y losas de hormigón.\n"
            u"Aceptar para seleccionar · Esc en el modelo para cancelar.",
            TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel,
        )
        return r == TaskDialogResult.Ok


def pick_lote_inicial(uidoc):
    """Lista de ``Reference`` o ``None`` si cancela."""
    if uidoc is None:
        return None
    try:
        refs = list(
            uidoc.Selection.PickObjects(
                ObjectType.Element,
                ArmadoColumnasSelectionFilter(),
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
    Al menos un elemento de hormigón permitido con proyección válida.
    Devuelve ``(ok, message)``.
    """
    from armado_columnas_v2.revit.adapters import (
        elements_from_refs,
        classify_elements,
        iter_classified,
    )
    from armado_columnas_v2.revit.view_geometry import elevation_rect_from_element

    elems = elements_from_refs(document, refs)
    if not elems:
        return False, u"No seleccionó ningún elemento."

    for el in elems:
        if not _es_hormigon(el):
            return False, (
                u"El lote incluye elementos que no son de hormigón "
                u"(Material for Model Behavior ≠ Concrete).\n\n"
                u"Solo se permiten columnas, fundaciones, vigas y losas de hormigón."
            )

    groups = classify_elements(elems)
    if groups.get("other"):
        return False, (
            u"El lote incluye elementos no permitidos.\n"
            u"Solo: columnas, fundaciones, vigas y losas de hormigón."
        )
    allowed = list(iter_classified(groups))
    if not allowed:
        return False, (
            u"Seleccione al menos una columna, fundación, viga o losa de hormigón."
        )

    if view is not None:
        ok_proj = 0
        for el in allowed:
            if elevation_rect_from_element(el, view) is not None:
                ok_proj += 1
        if ok_proj == 0:
            return False, (
                u"No se pudo proyectar ningún elemento en la vista activa.\n"
                u"Use un alzado o sección donde los elementos sean visibles."
            )

    return True, u""
