# -*- coding: utf-8 -*-
"""Adapter de escritura de Structural Rebar."""


class RebarWriter(object):
    """Contrato de creación de barras para extraer el motor legado gradualmente."""

    def __init__(self, doc, create_from_curves):
        self.doc = doc
        self.create_from_curves = create_from_curves

    def create_from_curves_no_hooks(self, host, bar_type, curves, normal_vec):
        return self.create_from_curves(
            self.doc,
            host,
            bar_type,
            curves,
            normal_vec,
        )
