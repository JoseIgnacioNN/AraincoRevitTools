# -*- coding: utf-8 -*-
"""Contexto Revit aislado de la capa de aplicación."""


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
