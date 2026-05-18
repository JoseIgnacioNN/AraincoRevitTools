# -*- coding: utf-8 -*-
"""Segmento de barra generado por el troceo/traslapes."""

# Colores por segmento para la preview visual (índice 0-based)
SEGMENT_COLORS = [
    "#1B6CA8",  # azul
    "#17A589",  # verde-teal
    "#D4AC0D",  # dorado
    "#E67E22",  # naranja
    "#C0392B",  # rojo
    "#8E44AD",  # violeta
    "#2ECC71",  # verde
    "#E74C3C",  # rojo intenso
]


def segment_color(index):
    return SEGMENT_COLORS[index % len(SEGMENT_COLORS)]


class SpliceSegment(object):
    """Tramo vertical de barras entre dos puntos de corte consecutivos.

    Las cotas están en milímetros para facilitar su uso sin conversión en la UI.
    """

    def __init__(self, segment_id, z_start_mm, z_end_mm):
        self.segment_id = int(segment_id)   # 1-based para UI
        self.z_start_mm = float(z_start_mm)
        self.z_end_mm = float(z_end_mm)
        self.diameter_mm = 20               # diámetro asignado en Step 4

    @property
    def height_mm(self):
        return self.z_end_mm - self.z_start_mm

    @property
    def height_m(self):
        return self.height_mm / 1000.0

    @property
    def color(self):
        return segment_color(self.segment_id - 1)

    def lap_length_mm(self):
        """Longitud de traslape según regla 50Ø."""
        return 50.0 * self.diameter_mm

    def level_range_label(self):
        return u"N+{0:.2f} a N+{1:.2f}".format(
            self.z_start_mm / 1000.0,
            self.z_end_mm / 1000.0,
        )

    def diameter_label(self):
        return u"Ø {0}".format(self.diameter_mm)

    def __repr__(self):
        return "SpliceSegment({0}, {1:.0f}-{2:.0f}mm, Ø{3})".format(
            self.segment_id, self.z_start_mm, self.z_end_mm, self.diameter_mm
        )
