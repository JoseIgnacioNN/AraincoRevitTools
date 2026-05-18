# -*- coding: utf-8 -*-
"""DTO de resultado del wizard; entrada para RebarCreationService."""


class WizardRequest(object):
    """Encapsula todas las decisiones tomadas en los 4 pasos del wizard.

    Es el único objeto que cruza la frontera UI → servicio de creación.
    """

    def __init__(self, column_groups, splice_segments, distributions, cover_mm=25.0):
        # Step 1 — grupos de columnas seleccionados
        self.column_groups = list(column_groups)          # [ColumnGroup]
        # Step 2 — segmentos de troceo generados
        self.splice_segments = list(splice_segments)      # [SpliceSegment]
        # Step 3+4 — distribución y diámetros (RebarDistribution por grupo)
        self.distributions = list(distributions)          # [RebarDistribution]
        self.cover_mm = float(cover_mm)

    @property
    def total_columns(self):
        return sum(g.column_count for g in self.column_groups)

    @property
    def segment_count(self):
        return len(self.splice_segments)

    @property
    def z_bottom_mm(self):
        if not self.column_groups:
            return 0.0
        return min(g.z_bottom_mm for g in self.column_groups)

    @property
    def z_top_mm(self):
        if not self.column_groups:
            return 0.0
        return max(g.z_top_mm for g in self.column_groups)

    @property
    def total_height_m(self):
        return (self.z_top_mm - self.z_bottom_mm) / 1000.0

    def distribution_for_group(self, group_id):
        for d in self.distributions:
            if d.group_id == group_id:
                return d
        return None


class WizardResult(object):
    def __init__(self, success=False, message=u"", bars_created=0):
        self.success = bool(success)
        self.message = message or u""
        self.bars_created = int(bars_created)
