# -*- coding: utf-8 -*-
"""Gestiona y valida la distribución de barras por grupo de sección.

Pura: no importa Revit API.
"""

from column_reinforcement_v2.models.rebar_distribution import RebarDistribution


class DistributionService(object):
    """Crea, actualiza y valida RebarDistribution para cada ColumnGroup."""

    def create_defaults(self, column_groups):
        """Crea distribuciones por defecto para cada grupo."""
        distributions = []
        for group in column_groups:
            d = RebarDistribution(
                group_id=group.group_id,
                side_a_count=self._default_bars_for_dimension(group.section.side_a_mm),
                side_b_count=self._default_bars_for_dimension(group.section.side_b_mm),
            )
            distributions.append(d)
        return distributions

    @staticmethod
    def _default_bars_for_dimension(dim_mm):
        """Heurística: más barras para secciones más grandes."""
        if dim_mm >= 800:
            return 6
        if dim_mm >= 500:
            return 4
        return 2

    def increment_a(self, distribution):
        distribution.side_a_count = min(20, distribution.side_a_count + 1)

    def decrement_a(self, distribution):
        distribution.side_a_count = max(2, distribution.side_a_count - 1)

    def increment_b(self, distribution):
        distribution.side_b_count = min(20, distribution.side_b_count + 1)

    def decrement_b(self, distribution):
        distribution.side_b_count = max(2, distribution.side_b_count - 1)

    def get_for_group(self, distributions, group_id):
        for d in distributions:
            if d.group_id == group_id:
                return d
        return None
