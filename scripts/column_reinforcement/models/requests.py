# -*- coding: utf-8 -*-
"""DTOs de entrada/salida sin dependencia directa de Revit API."""


class ColumnReinforcementRequest(object):
    """Configuración de ejecución producida por la UI."""

    def __init__(
        self,
        use_legacy_engine=True,
        enable_split_planes=True,
        enable_embedment=True,
        source="wpf",
    ):
        self.use_legacy_engine = bool(use_legacy_engine)
        self.enable_split_planes = bool(enable_split_planes)
        self.enable_embedment = bool(enable_embedment)
        self.source = source or "wpf"


class ColumnReinforcementResult(object):
    """Resultado normalizado para futuras UI/reportes."""

    def __init__(self, success=False, message="", payload=None):
        self.success = bool(success)
        self.message = message or ""
        self.payload = payload or {}
