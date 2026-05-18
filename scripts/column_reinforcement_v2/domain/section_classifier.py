# -*- coding: utf-8 -*-
"""Clasifica la sección de una columna determinando lado A (menor) y B (mayor).

Pura: no importa Revit API.
"""

from column_reinforcement_v2.models.column_group import SectionGeometry


class SectionClassifier(object):
    """Extrae y clasifica las dimensiones de sección a partir de parámetros Revit."""

    FEET_TO_MM = 304.8
    PARAM_NAMES_WIDTH = ["b", "Width", "Ancho", "B", "width"]
    PARAM_NAMES_HEIGHT = ["h", "Depth", "Alto", "H", "depth", "height"]

    def classify(self, elem):
        """Devuelve SectionGeometry con width y height en mm.

        Estrategia:
        1. Parámetros en la instancia
        2. Parámetros en el tipo (Symbol)
        3. Bounding-box como fallback
        """
        w = self._read_param(elem, self.PARAM_NAMES_WIDTH)
        h = self._read_param(elem, self.PARAM_NAMES_HEIGHT)

        if w is not None and h is not None and w > 0 and h > 0:
            return SectionGeometry(w, h)

        # Fallback: bounding box en el plano horizontal
        return self._from_bbox(elem)

    def _read_param(self, elem, names):
        for name in names:
            val = self._try_param(elem, name)
            if val is not None:
                return val
            # Intentar en el tipo
            try:
                val = self._try_param(elem.Symbol, name)
                if val is not None:
                    return val
            except Exception:
                pass
        return None

    @staticmethod
    def _try_param(obj, name):
        try:
            p = obj.LookupParameter(name)
            if p is not None and p.HasValue:
                return p.AsDouble() * SectionClassifier.FEET_TO_MM
        except Exception:
            pass
        return None

    @staticmethod
    def _from_bbox(elem):
        try:
            bbox = elem.get_BoundingBox(None)
            dx = (bbox.Max.X - bbox.Min.X) * SectionClassifier.FEET_TO_MM
            dy = (bbox.Max.Y - bbox.Min.Y) * SectionClassifier.FEET_TO_MM
            # La altura Z no es la sección; usar X/Y
            w = max(dx, 1.0)
            h = max(dy, 1.0)
            return SectionGeometry(w, h)
        except Exception:
            return SectionGeometry(300.0, 300.0)  # fallback seguro
