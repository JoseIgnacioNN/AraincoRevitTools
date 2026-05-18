# -*- coding: utf-8 -*-
"""Adapter temporal hacia el motor legado.

La primera fase no reescribe los algoritmos de sólidos, fundaciones, patas ni
creación de líneas modelo. Este adapter centraliza la delegación para que el
resto de la arquitectura pueda crecer sin acoplarse al script monolítico.
"""

from column_reinforcement.models.requests import ColumnReinforcementResult


class LegacyColumnReinforcementService(object):
    """Ejecuta el `main()` legado dentro del contexto Revit actual."""

    def __init__(self, legacy_main):
        self.legacy_main = legacy_main

    def execute(self, context, request):
        if not request.use_legacy_engine:
            return ColumnReinforcementResult(
                False,
                "No hay motor alternativo configurado todavía.",
            )
        self.legacy_main()
        return ColumnReinforcementResult(True, "Ejecución completada.")
