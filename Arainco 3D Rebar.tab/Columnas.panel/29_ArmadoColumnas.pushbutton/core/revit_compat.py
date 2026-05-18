# -*- coding: utf-8 -*-
"""
Adapter de versión Revit.

Abstrae las diferencias de API entre Revit 2024/2025/2026 (UnitTypeId, etc.).
Portado y simplificado desde scripts/column_reinforcement/revit/versioning/adapters.py.
"""


class RevitVersionAdapter(object):
    """Wrapper ligero sobre la Application de Revit para conversión de unidades."""

    def __init__(self, app=None):
        self.app = app
        try:
            self.version = int(app.VersionNumber)
        except Exception:
            self.version = 2024

    # ------------------------------------------------------------------
    # Unidades
    # ------------------------------------------------------------------

    def _unit_mm(self):
        from Autodesk.Revit.DB import UnitTypeId
        return UnitTypeId.Millimeters

    def to_internal_mm(self, value_mm):
        """Convierte mm a unidades internas Revit (pies)."""
        from Autodesk.Revit.DB import UnitUtils
        return UnitUtils.ConvertToInternalUnits(float(value_mm), self._unit_mm())

    def from_internal_mm(self, value_internal):
        """Convierte unidades internas Revit (pies) a mm."""
        from Autodesk.Revit.DB import UnitUtils
        return UnitUtils.ConvertFromInternalUnits(float(value_internal), self._unit_mm())


class Revit2024Adapter(RevitVersionAdapter):
    pass


class Revit2025Adapter(Revit2024Adapter):
    pass


class Revit2026Adapter(Revit2025Adapter):
    pass


def create_version_adapter(app):
    """Factory: devuelve el adapter correspondiente a la versión instalada."""
    try:
        version = int(app.VersionNumber)
    except Exception:
        version = 2024
    if version >= 2026:
        return Revit2026Adapter(app)
    if version >= 2025:
        return Revit2025Adapter(app)
    return Revit2024Adapter(app)
