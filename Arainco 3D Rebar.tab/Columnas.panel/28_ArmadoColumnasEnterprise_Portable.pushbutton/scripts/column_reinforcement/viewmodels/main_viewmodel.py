# -*- coding: utf-8 -*-
"""ViewModel inicial para la herramienta enterprise de columnas."""

from column_reinforcement.models.requests import ColumnReinforcementRequest


class ColumnReinforcementViewModel(object):
    """Estado simple de la ventana WPF.

    Se mantiene deliberadamente pequeño: la primera fase preserva el motor legado y
    deja la UI preparada para mover parámetros reales sin mezclar lógica Revit.
    """

    def __init__(self):
        self.title = u"Arainco: Armado Columnas"
        self.subtitle = u"Flujo enterprise modular (pyRevit/WPF)"
        self.use_legacy_engine = True
        self.enable_split_planes = True
        self.enable_embedment = True

    def to_request(self):
        return ColumnReinforcementRequest(
            use_legacy_engine=self.use_legacy_engine,
            enable_split_planes=self.enable_split_planes,
            enable_embedment=self.enable_embedment,
            source="wpf",
        )
