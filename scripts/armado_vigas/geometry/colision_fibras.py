# -*- coding: utf-8 -*-
"""
Colisión longitudinal de fibras — mismo criterio sup/inf y laterales.

Delega en :mod:`geometria_empotramiento_extremos` (sonda 50 mm, empotramiento,
refinamiento en **columna estructural** con pata L, extremo libre con pata L,
muro / obstáculo genérico con empotramiento).
"""

from __future__ import division

from armado_vigas.geometry.extremos import aplicar_extremos_a_linea_fusionada

__all__ = [u"aplicar_colision_extremos_fibra"]


def aplicar_colision_extremos_fibra(
    document,
    line,
    ids_seleccion,
    chain_elements,
    diam_mm,
    resolver_inicio=True,
    resolver_fin=True,
    end_mode_start=None,
    end_mode_end=None,
):
    """
    Aplica :func:`aplicar_extremos_a_linea_fusionada` (reglas Armado vigas sup/inf).

    ``end_mode_start`` / ``end_mode_end``: `auto`|`emp`|`pata_l` sobre extremos
    **0/1 de la curva** (no del canvas).

    Returns:
        ``(linea, meta_inicio, meta_fin)``
    """
    return aplicar_extremos_a_linea_fusionada(
        document,
        line,
        ids_seleccion,
        chain_elements,
        diam_mm,
        resolver_inicio=resolver_inicio,
        resolver_fin=resolver_fin,
        end_mode_start=end_mode_start,
        end_mode_end=end_mode_end,
    )
