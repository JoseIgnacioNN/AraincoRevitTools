# -*- coding: utf-8 -*-
"""Distribución de barras longitudinales por grupo de sección."""


class RebarDistribution(object):
    """Cantidad de barras por cara para un grupo de sección dado.

    Convención:
      side_a = cara del lado menor  (A = min(width, height))
      side_b = cara del lado mayor  (B = max(width, height))

    El conteo incluye las barras de esquina.  Para una columna cuadrada
    con bars_side_a == bars_side_b == n, el total de barras es 4*(n-1).
    Para rectangular: 2*(n_a - 1) + 2*(n_b - 1) + 4 = 2*n_a + 2*n_b - 4
    ... pero siempre se obtiene sumando las 4 caras y restando esquinas.
    """

    def __init__(self, group_id, side_a_count=2, side_b_count=2):
        self.group_id = int(group_id)
        self._side_a = max(2, int(side_a_count))
        self._side_b = max(2, int(side_b_count))

    @property
    def side_a_count(self):
        return self._side_a

    @side_a_count.setter
    def side_a_count(self, value):
        self._side_a = max(2, int(value))

    @property
    def side_b_count(self):
        return self._side_b

    @side_b_count.setter
    def side_b_count(self, value):
        self._side_b = max(2, int(value))

    @property
    def total_bars(self):
        """Total de barras en la sección (esquinas contadas una vez)."""
        return 2 * self._side_a + 2 * self._side_b - 4

    def summary(self):
        return u"A={0}  B={1}  Total={2}".format(
            self._side_a, self._side_b, self.total_bars
        )

    def __repr__(self):
        return "RebarDistribution(group={0}, A={1}, B={2})".format(
            self.group_id, self._side_a, self._side_b
        )
