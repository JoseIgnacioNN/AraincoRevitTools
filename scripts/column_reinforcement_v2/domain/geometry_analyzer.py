# -*- coding: utf-8 -*-
"""Extrae datos geométricos de elementos Revit para el wizard.

Depende de Revit API (importaciones diferidas para no romper tests unitarios).
"""

FEET_TO_MM = 304.8


class GeometryAnalyzer(object):
    """Extrae rango Z y ubicación en planta de columnas estructurales."""

    def get_z_range_mm(self, elem):
        """Devuelve (z_bottom_mm, z_top_mm) desde el BoundingBox del elemento."""
        bbox = elem.get_BoundingBox(None)
        return (
            bbox.Min.Z * FEET_TO_MM,
            bbox.Max.Z * FEET_TO_MM,
        )

    def get_location_point_mm(self, elem):
        """Devuelve (x_mm, y_mm) del punto de inserción de la columna."""
        try:
            loc = elem.Location
            pt = loc.Point
            return (pt.X * FEET_TO_MM, pt.Y * FEET_TO_MM)
        except Exception:
            return (0.0, 0.0)

    def are_vertically_continuous(self, elem_lower, elem_upper, tol_mm=50.0):
        """True si la parte superior de lower coincide (aprox.) con el fondo de upper."""
        _, z_top = self.get_z_range_mm(elem_lower)
        z_bot, _ = self.get_z_range_mm(elem_upper)
        return abs(z_top - z_bot) <= tol_mm
