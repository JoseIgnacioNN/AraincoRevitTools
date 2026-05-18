# -*- coding: utf-8 -*-
"""Fachada de ejecución pyRevit para Armado Columnas Wizard v2.

Flujo:
  1. Solicitar selección de columnas estructurales al usuario.
  2. Agrupar columnas por sección geométrica (ColumnGroupingService).
  3. Abrir wizard WPF pre-poblado (show_singleton_wizard).
  4. Si el usuario confirma → ejecutar RebarCreationService dentro de Transaction.
  5. Mostrar resultado.
"""

import os
import sys

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_SCRIPTS_DIR = os.path.dirname(_THIS_DIR)
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

import clr
clr.AddReference("RevitAPIUI")
clr.AddReference("RevitAPI")

from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.DB import BuiltInCategory


class StructuralColumnFilter(ISelectionFilter):
    """Filtra la selección a sólo columnas estructurales."""

    _CAT_ID = int(BuiltInCategory.OST_StructuralColumns)

    def AllowElement(self, elem):
        try:
            return elem.Category.Id.IntegerValue == self._CAT_ID
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return False


def _pick_columns(uidoc):
    """Solicita la selección múltiple de columnas; devuelve lista de elementos."""
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            StructuralColumnFilter(),
            u"Selecciona las columnas a armar (contiguas en altura) → Enter para confirmar",
        )
        return [uidoc.Document.GetElement(r) for r in refs]
    except Exception:
        # El usuario canceló con Esc
        return []


def run_pyrevit(revit_app):
    """Entrada desde pushbutton pyRevit."""
    uidoc = revit_app.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(u"Arainco: Armado Columnas Wizard", u"No hay documento activo.")
        return

    doc = uidoc.Document

    # ── Step 1 fuera del wizard: selección en Revit ── #
    elements = _pick_columns(uidoc)
    if not elements:
        return  # cancelado por el usuario

    # Forzar recarga de módulos v2 para evitar cache stale de pyRevit
    _stale = [k for k in sys.modules if "column_reinforcement_v2" in k]
    for _k in _stale:
        del sys.modules[_k]

    # ── Mostrar wizard ── #
    from column_reinforcement_v2.ui.main_window import show_singleton_wizard
    wizard_request = show_singleton_wizard(elements)

    if wizard_request is None:
        return  # cancelado en el wizard

    # ── Ejecutar creación de barras ── #
    try:
        from column_reinforcement_v2.services.rebar_creation_service import RebarCreationService
    except Exception as imp_ex:
        TaskDialog.Show(u"Arainco: Armado Columnas Wizard", u"Error importando servicio:\n{0}".format(imp_ex))
        return

    from column_reinforcement_v2.revit.adapters import create_version_adapter

    version_adapter = create_version_adapter(
        revit_app.Application if hasattr(revit_app, "Application") else None
    )
    service = RebarCreationService(doc, version_adapter)
    result  = service.execute(wizard_request)

    # ── Resultado ── #
    if result.success:
        TaskDialog.Show(
            u"Arainco: Armado Columnas Wizard",
            u"Armado generado correctamente.\n{0}".format(result.message),
        )
    else:
        TaskDialog.Show(
            u"Arainco: Armado Columnas Wizard",
            u"Error durante la generación:\n{0}".format(result.message),
        )
