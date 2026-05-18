# -*- coding: utf-8 -*-
"""Modelos geométricos simples, testeables fuera de Revit."""


class SegmentZ(object):
    """Segmento vertical en unidades internas Revit (pies) o unidades abstractas en tests."""

    def __init__(self, z_start, dz):
        self.z_start = float(z_start)
        self.dz = float(dz)

    @property
    def z_end(self):
        return self.z_start + self.dz

    def as_tuple(self):
        return (self.z_start, self.dz)

    def copy(self):
        return SegmentZ(self.z_start, self.dz)
