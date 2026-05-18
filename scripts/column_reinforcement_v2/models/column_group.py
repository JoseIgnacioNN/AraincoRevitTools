# -*- coding: utf-8 -*-
"""Modelos de agrupación de columnas por sección geométrica."""


class SectionGeometry(object):
    """Dimensiones de sección transversal en milímetros."""

    SQUARE_TOL_MM = 5.0

    def __init__(self, width_mm, height_mm):
        self.width_mm = float(width_mm)
        self.height_mm = float(height_mm)

    @property
    def side_a_mm(self):
        """Lado menor (A) de la sección."""
        return min(self.width_mm, self.height_mm)

    @property
    def side_b_mm(self):
        """Lado mayor (B) de la sección."""
        return max(self.width_mm, self.height_mm)

    @property
    def is_square(self):
        return abs(self.width_mm - self.height_mm) <= self.SQUARE_TOL_MM

    def label(self):
        return u"{0:.0f} x {1:.0f}".format(self.width_mm, self.height_mm)

    def matches(self, other, tol_mm=10.0):
        """True si ambas secciones son equivalentes dentro de tolerancia."""
        return (
            abs(self.width_mm - other.width_mm) <= tol_mm
            and abs(self.height_mm - other.height_mm) <= tol_mm
        )

    def __repr__(self):
        return "SectionGeometry({0}x{1})".format(self.width_mm, self.height_mm)


class ColumnGroup(object):
    """Grupo de columnas apiladas verticalmente con la misma sección geométrica.

    Los IDs de elementos Revit se almacenan como enteros para evitar dependencia
    directa con la API en las capas de dominio.
    """

    # Cantidad mínima/máxima de barras por cara (validación básica)
    MIN_BARS_PER_SIDE = 2
    MAX_BARS_PER_SIDE = 20

    def __init__(self, group_id, section, element_ids, z_bottom_mm, z_top_mm):
        self.group_id = int(group_id)           # índice 1-based para UI
        self.section = section                   # SectionGeometry
        self.element_ids = list(element_ids)     # [int] ids Revit
        self.z_bottom_mm = float(z_bottom_mm)
        self.z_top_mm = float(z_top_mm)
        self.bars_side_a = 2
        self.bars_side_b = 2
        # Cotas Z de cada junta interna entre columnas apiladas (exc. extremos).
        # Poblado por ColumnGroupingService; permite al usuario elegir empalmes
        # incluso dentro de grupos de sección homogénea.
        self.column_z_joints = []               # [float] mm

    @property
    def column_count(self):
        return len(self.element_ids)

    @property
    def height_mm(self):
        return self.z_top_mm - self.z_bottom_mm

    @property
    def height_m(self):
        return self.height_mm / 1000.0

    def level_range_label(self):
        return u"N+{0:.2f} a N+{1:.2f}".format(
            self.z_bottom_mm / 1000.0,
            self.z_top_mm / 1000.0,
        )

    def column_range_label(self):
        return u"({0} a {1})".format(
            u"N+{0:.2f}".format(self.z_bottom_mm / 1000.0),
            u"N+{0:.2f}".format(self.z_top_mm / 1000.0),
        )

    def __repr__(self):
        return "ColumnGroup({0}, {1}, cols={2})".format(
            self.group_id, self.section, self.column_count
        )
