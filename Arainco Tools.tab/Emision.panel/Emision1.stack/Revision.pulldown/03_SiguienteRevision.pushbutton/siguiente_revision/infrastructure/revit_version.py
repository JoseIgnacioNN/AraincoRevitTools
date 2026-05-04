# -*- coding: utf-8 -*-
"""
Revit Version Adapter.

Detecta la versión de Revit en tiempo de ejecución y adapta el comportamiento
a cambios de API entre versiones (especialmente 2024+ donde ElementId es 64-bit).
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str


class RevitVersionAdapter(object):
    """
    Encapsula la detección de versión de Revit y adaptaciones de API.

    Instanciado una vez por ejecución de la herramienta y pasado a los
    servicios que necesitan comportamiento condicional por versión.
    """

    def __init__(self, app):
        try:
            self.year = int(unicode(app.VersionNumber).strip())
        except Exception:
            self.year = 2024

    @property
    def uses_64bit_element_id(self):
        """Revit 2024+ usa ElementId con Value como Int64."""
        return self.year >= 2024

    @property
    def version_string(self):
        return u"Revit {}".format(self.year)

    def element_id_integer(self, eid):
        """Extrae el valor entero de un ElementId de forma segura en cualquier versión."""
        try:
            return int(eid.Value)
        except Exception:
            try:
                return int(eid.IntegerValue)
            except Exception:
                return 0

    def element_id_from_int(self, val):
        """Construye un ElementId desde un entero, usando Int64 en 2024+."""
        from Autodesk.Revit.DB import ElementId
        from System import Int64
        try:
            return ElementId(Int64(int(val)))
        except Exception:
            try:
                return ElementId(int(val))
            except Exception:
                return ElementId.InvalidElementId

    def is_compatible(self, min_year=2024):
        """Verifica que la versión cumple el mínimo requerido."""
        return self.year >= min_year

    def get_revision_all_ids(self, doc):
        """
        Retorna los ElementId de todas las revisiones del proyecto en orden.

        Usa Revision.GetAllRevisionIds (API estable 2024+).
        """
        from Autodesk.Revit.DB import Revision, ElementId
        try:
            raw = Revision.GetAllRevisionIds(doc)
        except Exception:
            raw = None
        out = []
        if raw:
            for eid in raw:
                if eid is not None and eid != ElementId.InvalidElementId:
                    out.append(eid)
        return out
