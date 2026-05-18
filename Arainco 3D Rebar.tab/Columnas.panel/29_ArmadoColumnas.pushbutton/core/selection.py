# -*- coding: utf-8 -*-
"""
Selección interactiva de columnas estructurales y utilidades de ElementId.

REGLA DE CAPA: este módulo puede leer de Revit (UIDocument.Selection)
pero NO escribe en el modelo ni abre transacciones.
"""
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType


# ---------------------------------------------------------------------------
# Filtro de selección
# ---------------------------------------------------------------------------

class StructuralColumnFilter(ISelectionFilter):
    """Permite seleccionar solo columnas estructurales en el canvas de Revit."""

    _CAT_IV = int(BuiltInCategory.OST_StructuralColumns)

    def AllowElement(self, elem):
        try:
            return int(elem.Category.Id.IntegerValue) == self._CAT_IV
        except Exception:
            return False

    def AllowReference(self, ref, point):
        return False


# ---------------------------------------------------------------------------
# Selección interactiva
# ---------------------------------------------------------------------------

def pick_structural_columns(uidoc):
    """
    Abre el PickObjects de Revit con filtro de columnas estructurales.
    Devuelve lista de References (puede estar vacía; OperationCanceledException
    se propaga hacia runner.py para un abort limpio).
    """
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        StructuralColumnFilter(),
        u"Selecciona columnas estructurales y finaliza (Enter).",
    )
    return list(refs) if refs else []


def pick_structural_columns_optional(uidoc):
    """
    Igual que pick_structural_columns pero Esc devuelve lista vacía en lugar
    de propagar OperationCanceledException.
    """
    from Autodesk.Revit.Exceptions import OperationCanceledException
    try:
        return pick_structural_columns(uidoc)
    except OperationCanceledException:
        return []


# ---------------------------------------------------------------------------
# Resolución de elementos
# ---------------------------------------------------------------------------

def element_id_iv(elem):
    """
    Entero estable del Id de un elemento. Compatible con Revit 2024+
    (Id.Value) y versiones anteriores (Id.IntegerValue).
    Devuelve -1 si no puede resolverse.
    """
    if elem is None:
        return -1
    try:
        return int(elem.Id.Value)
    except AttributeError:
        pass
    try:
        return int(elem.Id.IntegerValue)
    except Exception:
        return -1


def build_column_elements_ordered(doc, refs):
    """
    Dado el listado de References de PickObjects, resuelve y deduplica los
    elementos Revit correspondientes, manteniendo el orden de selección.

    Lanza Exception si la lista queda vacía tras deduplicar.
    """
    seen, out = set(), []
    for ref in refs or []:
        if ref is None:
            continue
        try:
            iv = int(ref.ElementId.IntegerValue)
        except AttributeError:
            try:
                iv = int(ref.ElementId.Value)
            except Exception:
                continue
        if iv in seen:
            continue
        elem = doc.GetElement(ref.ElementId)
        if elem is not None:
            seen.add(iv)
            out.append(elem)
    if not out:
        raise Exception(
            u"No quedaron columnas estructurales tras deduplicar la selección."
        )
    return out
