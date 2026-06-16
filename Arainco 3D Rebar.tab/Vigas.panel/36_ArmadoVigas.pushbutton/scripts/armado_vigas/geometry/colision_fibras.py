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
):
    """
    Aplica :func:`aplicar_extremos_a_linea_fusionada` (reglas Armado vigas sup/inf).

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
    )
