# -*- coding: utf-8 -*-
"""Adapters por versión Revit.

Revit 2024+ usa `UnitTypeId`; las clases 2025/2026 heredan hasta que exista
una diferencia API concreta que justificar.
"""


class RevitVersionAdapter(object):
    def __init__(self, app=None):
        self.app = app
        self.version_number = self._read_version(app)

    def _read_version(self, app):
        try:
            return int(app.VersionNumber)
        except Exception:
            return 2024

    def millimeters_unit_type_id(self):
        from Autodesk.Revit.DB import UnitTypeId

        return UnitTypeId.Millimeters

    def to_internal_mm(self, value_mm):
        from Autodesk.Revit.DB import UnitUtils

        return UnitUtils.ConvertToInternalUnits(
            float(value_mm),
            self.millimeters_unit_type_id(),
        )

    def from_internal_mm(self, value_internal):
        from Autodesk.Revit.DB import UnitUtils

        return UnitUtils.ConvertFromInternalUnits(
            float(value_internal),
            self.millimeters_unit_type_id(),
        )


class Revit2024Adapter(RevitVersionAdapter):
    pass


class Revit2025Adapter(Revit2024Adapter):
    pass


class Revit2026Adapter(Revit2025Adapter):
    pass


def create_version_adapter(app):
    try:
        version = int(app.VersionNumber)
    except Exception:
        version = 2024
    if version >= 2026:
        return Revit2026Adapter(app)
    if version >= 2025:
        return Revit2025Adapter(app)
    return Revit2024Adapter(app)
