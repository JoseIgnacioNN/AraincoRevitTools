# -*- coding: utf-8 -*-
"""Genera segmentos de troceo a partir de los puntos de corte seleccionados.

Reutiliza la geometría pura de column_reinforcement v1:
  geometry/segments.py → split_z_span_by_cut_values
"""

import os
import sys

# Asegurar acceso a la v1 para reutilizar geometría pura testeable
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_SCRIPTS_DIR = os.path.dirname(os.path.dirname(_THIS_DIR))
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

from column_reinforcement.geometry.segments import split_z_span_by_cut_values
from column_reinforcement_v2.models.splice_segment import SpliceSegment

# Tolerancia en mm para fusionar cortes cercanos (convertida a pies internamente
# solo si fuera necesario; aquí trabajamos en mm directo).
MERGE_TOL_MM = 50.0


class SpliceLogicService(object):
    """Genera [SpliceSegment] desde grupos de columna y puntos de corte elegidos.

    Los puntos de corte corresponden a las cotas z (en mm) de los límites de
    columna que el usuario marcó como empalmes en el Step 2.
    """

    def generate_segments(self, column_groups, cut_z_values_mm):
        """Devuelve lista de SpliceSegment ordenados de abajo a arriba.

        Args:
            column_groups: [ColumnGroup] — resultado del Step 1.
            cut_z_values_mm: [float] — cotas de corte en mm elegidas por usuario.
        """
        if not column_groups:
            return []

        z_bottom = min(g.z_bottom_mm for g in column_groups)
        z_top = max(g.z_top_mm for g in column_groups)

        # split_z_span_by_cut_values trabaja en unidades arbitrarias (aquí mm)
        raw_segments = split_z_span_by_cut_values(
            z_bottom,
            z_top,
            cut_z_values_mm,
            MERGE_TOL_MM,
        )

        result = []
        for idx, seg in enumerate(raw_segments):
            result.append(SpliceSegment(
                segment_id=idx + 1,
                z_start_mm=seg.z_start,
                z_end_mm=seg.z_end,
            ))
        return result

    def default_cut_points(self, column_groups):
        """Devuelve la lista de cotas z de los límites entre columnas (sin extremos).

        Son los puntos de corte por defecto: juntas entre columnas consecutivas.
        """
        if len(column_groups) <= 1:
            return []
        # Los límites entre grupos son los z_top del grupo inferior
        cuts = []
        for group in column_groups[:-1]:
            cuts.append(group.z_top_mm)
        return cuts
