# -*- coding: utf-8 -*-
"""Command Pattern para la ejecución de armado de columnas."""

from column_reinforcement.models.requests import ColumnReinforcementResult
from column_reinforcement.ui.main_window import show_singleton_dialog


class ColumnReinforcementCommand(object):
    """Orquesta la UI y delega la ejecución al motor configurado."""

    def __init__(self, execution_service):
        self.execution_service = execution_service

    def execute(self, context):
        request = show_singleton_dialog()
        if request is None:
            return ColumnReinforcementResult(False, "Ejecución cancelada.")
        return self.execution_service.execute(context, request)
