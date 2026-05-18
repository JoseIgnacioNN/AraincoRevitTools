# -*- coding: utf-8 -*-
"""Factories de strategies y servicios."""

from column_reinforcement.services.strategies import (
    LapExtensionPolicy,
    SplitPolicyA,
    SplitPolicyB,
)


class StrategyFactory(object):
    """Crea políticas por configuración/versionado."""

    def __init__(self, version_adapter=None):
        self.version_adapter = version_adapter

    def create_split_policies(self):
        return [SplitPolicyA(), SplitPolicyB()]

    def create_lap_extension_policy(self):
        return LapExtensionPolicy()
