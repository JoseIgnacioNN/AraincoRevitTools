# -*- coding: utf-8 -*-
"""Aplicación de extremos empotrado/gancho sobre líneas fusionadas."""

from geometria_empotramiento_extremos import (
    MODO_EMPOTRAMIENTO,
    MODO_GANCHO,
    aplicar_extremos_linea,
    element_ids_desde_elementos,
    resolver_extremo_linea,
)

__all__ = [
    "MODO_EMPOTRAMIENTO",
    "MODO_GANCHO",
    "aplicar_extremos_a_linea_fusionada",
    "element_ids_desde_elementos",
    "resolver_extremo_linea",
]


def aplicar_extremos_a_linea_fusionada(
    document,
    line,
    ids_seleccion,
    host_chain_elements,
    diam_nominal_mm,
    resolver_inicio=True,
    resolver_fin=True,
):
    """
    Wrapper de :func:`geometria_empotramiento_extremos.aplicar_extremos_linea`.

    ``host_chain_elements``: vigas de la cadena colineal (excluidas de colisión).
    """
    ids_excluir = element_ids_desde_elementos(host_chain_elements)
    return aplicar_extremos_linea(
        document,
        line,
        ids_seleccion,
        ids_excluir=ids_excluir,
        diam_nominal_mm=diam_nominal_mm,
        resolver_inicio=resolver_inicio,
        resolver_fin=resolver_fin,
    )
