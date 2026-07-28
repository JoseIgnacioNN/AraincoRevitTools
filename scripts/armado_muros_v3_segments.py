# -*- coding: utf-8 -*-
"""Segmentos de barra por troceo (misma semántica que ``build_cabezal_segments``)."""

from __future__ import print_function


def build_bar_segments(n_walls, empalme_indices):
    """
    Cada índice con empalme activo **inicia** un tramo.

    Ejemplo ``n=5``, empalmes ``[1, 3]`` →
    ``S0=[0]``, ``S1=[1,2] owner 1``, ``S2=[3,4] owner 3``.
    """
    try:
        n = max(0, int(n_walls or 0))
    except Exception:
        n = 0
    if n <= 0:
        return []
    E = sorted(int(i) for i in (empalme_indices or []))
    E = [i for i in E if 0 <= i < n]
    if not E:
        return [{u"id": 0, u"wall_indices": list(range(n)), u"owner_index": 0}]
    segs = []
    sid = 0
    if E[0] > 0:
        segs.append(
            {
                u"id": sid,
                u"wall_indices": list(range(0, E[0])),
                u"owner_index": 0,
            }
        )
        sid += 1
    for j in range(len(E)):
        if j + 1 < len(E):
            wall_indices = list(range(E[j], E[j + 1]))
        else:
            wall_indices = list(range(E[j], n))
        if not wall_indices:
            continue
        segs.append(
            {
                u"id": sid,
                u"wall_indices": wall_indices,
                u"owner_index": E[j],
            }
        )
        sid += 1
    return segs
