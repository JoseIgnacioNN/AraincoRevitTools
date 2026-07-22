# -*- coding: utf-8 -*-
"""Contexto Revit aislado de la capa de aplicación."""


def is_section_or_elevation_view(view):
    """True si la vista es sección o alzado (no planta, 3D, drafting, etc.)."""
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import ViewDrafting, ViewSection, ViewType

        try:
            if isinstance(view, ViewDrafting):
                return False
        except Exception:
            pass
        try:
            if isinstance(view, ViewSection):
                return True
        except Exception:
            pass
        try:
            vt = view.ViewType
            return vt == ViewType.Section or vt == ViewType.Elevation
        except Exception:
            return False
    except Exception:
        return False


def is_section_or_elevation_uiapp(uiapp):
    try:
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            return False
        return is_section_or_elevation_view(uidoc.ActiveView)
    except Exception:
        return False


class RevitExecutionContext(object):
    """Agrupa objetos host sin que las capas puras importen Revit API."""

    def __init__(self, revit_app=None, uidoc=None, doc=None, version_adapter=None):
        self.revit_app = revit_app
        self.uidoc = uidoc
        self.doc = doc
        self.version_adapter = version_adapter

    @classmethod
    def from_pyrevit(cls, revit_app, version_adapter=None):
        uidoc = revit_app.ActiveUIDocument
        doc = uidoc.Document if uidoc is not None else None
        return cls(revit_app, uidoc, doc, version_adapter)
