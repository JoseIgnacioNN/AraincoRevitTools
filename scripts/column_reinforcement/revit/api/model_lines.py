# -*- coding: utf-8 -*-
"""Adapter de escritura de líneas modelo."""


class ModelLineWriter(object):
    """Contrato inicial para separar creación Revit del servicio de armado."""

    def __init__(self, doc):
        self.doc = doc

    def create_model_curve(self, curve, sketch_plane):
        return self.doc.Create.NewModelCurve(curve, sketch_plane)
