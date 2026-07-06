# -*- coding: utf-8 -*-
"""Fachada de ejecución pyRevit/RPS para armado de columnas.

Punto oficial pyRevit (pushbutton): ``script.py`` carga este módulo con ``imp.load_source``
y llama ``run_pyrevit`` con ``__revit__``.

Alternativa (script legado como ``__main__``): ``column_reinforcement_layout_rps.run_pyrevit``
delega aquí: reload del módulo legado, inyección doc/uidoc/__revit__, ejecución de ``main``
vía ``LegacyColumnReinforcementService``.

RPS con ``main`` ya resuelto: ``run_rps`` sin recargar el módulo legado.
"""

import importlib
import os
import sys

from column_reinforcement.models.requests import ColumnReinforcementRequest
from column_reinforcement.revit.api.context import RevitExecutionContext
from column_reinforcement.revit.versioning.adapters import create_version_adapter
from column_reinforcement.services.command import ColumnReinforcementCommand
from column_reinforcement.services.legacy_engine import LegacyColumnReinforcementService


def _ensure_scripts_dir():
    scripts_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    return scripts_dir


def _configure_legacy_module(legacy_module, revit_app):
    """Inyecta doc/uidoc en el módulo legado cargado por `importlib`/pyRevit."""
    try:
        uidoc = revit_app.ActiveUIDocument
        legacy_module.uidoc = uidoc
        legacy_module.doc = uidoc.Document if uidoc is not None else None
        legacy_module.__revit__ = revit_app
    except Exception:
        pass


def _default_request():
    """Opciones que antes expone la UI; siempre activas sin formulario."""
    return ColumnReinforcementRequest(
        use_legacy_engine=True,
        enable_split_planes=True,
        enable_embedment=True,
        source="direct",
    )


def _run_with_legacy_main(revit_app, legacy_main, show_wpf=False):
    version_adapter = create_version_adapter(
        revit_app.Application if hasattr(revit_app, "Application") else None
    )
    context = RevitExecutionContext.from_pyrevit(revit_app, version_adapter)
    service = LegacyColumnReinforcementService(legacy_main)
    if show_wpf:
        command = ColumnReinforcementCommand(service)
        return command.execute(context)
    return service.execute(context, _default_request())


def run_pyrevit(revit_app):
    """Entrada desde pushbutton pyRevit."""
    _ensure_scripts_dir()
    from column_reinforcement.revit.api.context import is_section_or_elevation_uiapp

    if not is_section_or_elevation_uiapp(revit_app):
        try:
            from Autodesk.Revit.UI import TaskDialog

            TaskDialog.Show(
                u"Arainco: Armado Columnas",
                u"Esta herramienta solo está disponible en vistas de sección o alzado.\n"
                u"Abra o active una de esas vistas e intente de nuevo.",
            )
        except Exception:
            pass
        return None
    legacy = importlib.import_module("column_reinforcement_layout_rps")
    try:
        legacy = importlib.reload(legacy)
    except AttributeError:
        legacy = reload(legacy)  # noqa: F821 - IronPython 2 / pyRevit
    _configure_legacy_module(legacy, revit_app)
    return _run_with_legacy_main(revit_app, legacy.main, show_wpf=False)


def run_rps(revit_app, legacy_main):
    """Entrada desde el script RPS legado; mantiene el wrapper temporal."""
    return _run_with_legacy_main(revit_app, legacy_main, show_wpf=False)
