# -*- coding: utf-8 -*-
"""Agrupa columnas estructurales seleccionadas en grupos por sección geométrica.

Depende de Revit API a través de GeometryAnalyzer y SectionClassifier.
"""

from column_reinforcement_v2.domain.geometry_analyzer import GeometryAnalyzer
from column_reinforcement_v2.domain.section_classifier import SectionClassifier
from column_reinforcement_v2.models.column_group import ColumnGroup

SECTION_TOL_MM = 10.0


class ColumnGroupingService(object):
    """Toma una lista de FamilyInstance (columnas) y devuelve [ColumnGroup].

    Algoritmo:
    1. Extraer z_bottom, z_top y sección de cada columna.
    2. Ordenar por z_bottom ascendente.
    3. Agrupar columnas consecutivas con sección igual (dentro de tolerancia).
    """

    def __init__(self):
        self._analyzer = GeometryAnalyzer()
        self._classifier = SectionClassifier()

    def group(self, elements):
        """Devuelve lista ordenada de ColumnGroup."""
        annotated = self._annotate(elements)
        annotated.sort(key=lambda x: x[1])  # sort by z_bottom
        return self._build_groups(annotated)

    def _annotate(self, elements):
        result = []
        for elem in elements:
            try:
                z_bot, z_top = self._analyzer.get_z_range_mm(elem)
                section = self._classifier.classify(elem)
                elem_id = elem.Id.IntegerValue
                result.append((elem_id, z_bot, z_top, section))
            except Exception:
                pass
        return result

    def _build_groups(self, annotated):
        if not annotated:
            return []

        groups = []
        group_id = 1

        current_section = annotated[0][3]
        current_items = [annotated[0]]   # [(elem_id, z_bot, z_top, section)]

        for item in annotated[1:]:
            _, _, _, section = item
            if section.matches(current_section, SECTION_TOL_MM):
                current_items.append(item)
            else:
                groups.append(self._make_group(group_id, current_section, current_items))
                group_id += 1
                current_section = section
                current_items = [item]

        groups.append(self._make_group(group_id, current_section, current_items))
        return groups

    @staticmethod
    def _make_group(group_id, section, items):
        """Construye un ColumnGroup desde los items del batch actual."""
        sorted_items = sorted(items, key=lambda x: x[1])   # por z_bot
        elem_ids  = [i[0] for i in sorted_items]
        z_bottom  = sorted_items[0][1]
        z_top     = sorted_items[-1][2]

        group = ColumnGroup(group_id, section, elem_ids, z_bottom, z_top)

        # Juntas entre columnas consecutivas: z_top de cada columna salvo la última.
        # Son los candidatos a empalme/troceo que el usuario puede activar en Step 2.
        group.column_z_joints = [i[2] for i in sorted_items[:-1]]
        return group
