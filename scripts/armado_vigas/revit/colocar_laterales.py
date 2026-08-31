# -*- coding: utf-8 -*-
"""Colocación de barras laterales en caras del alma (Modelo B — ``armadura_vigas_capas``)."""

from __future__ import division

from armado_vigas.domain.laterales import (
    LATERALES_DIAM_DEFAULT,
    lateral_clear_mm_for_chain,
    session_n_laterales,
)
from armado_vigas.revit.laterales_colision import build_lateral_collision_ctx
from armado_vigas.revit.rebar_resources import resolve_bar_type_mm

try:
    from armadura_vigas_capas import (
        _build_collinear_chains_from_elements,
        _place_lateral_on_beam,
        _place_lateral_on_beam_aligned_chain,
    )
except Exception:
    _build_collinear_chains_from_elements = None
    _place_lateral_on_beam = None
    _place_lateral_on_beam_aligned_chain = None


def colocar_laterales(document, session, view=None):
    """
    Coloca laterales por cadena colineal (dos Rebar por cara ±ancho).

    ``view``: vista activa para forzar Unobscured OFF al finalizar (opcional).

    Returns:
        ``(n_barras, avisos, rebars_laterales, err)`` — ``rebars_laterales`` aparte de longitudinales.
    """
    if document is None or session is None:
        return 0, [], [], u"Sin documento o sesión."
    if not getattr(session, "lateralesEnabled", False):
        return 0, [], [], None
    if not session.framing_elements:
        return 0, [], [], u"No hay vigas en el lote."

    if (
        _build_collinear_chains_from_elements is None
        or _place_lateral_on_beam is None
        or _place_lateral_on_beam_aligned_chain is None
    ):
        return 0, [], [], u"No se cargó armadura_vigas_capas (laterales)."

    n_lat = session_n_laterales(session, 0)
    if n_lat < 1:
        return 0, [], [], None

    diam = getattr(session, "diamLaterales", None) or LATERALES_DIAM_DEFAULT
    bar_type = resolve_bar_type_mm(document, diam)
    if bar_type is None:
        return (
            0,
            [],
            [],
            u"No se encontró RebarBarType para laterales ø{0} mm.".format(diam),
        )

    chains = _build_collinear_chains_from_elements(document, session.framing_elements)
    if not chains:
        return 0, [], [], u"No hay vigas con eje válido para laterales."

    collected = []
    avisos = []
    total = 0
    beams_by_id = getattr(session, "domain_beams_by_element_id", None) or {}

    for chain in chains:
        clear_mm = lateral_clear_mm_for_chain(
            beams_by_id, chain, bar_diam_mm=diam,
        )
        collision_ctx = build_lateral_collision_ctx(session, chain, diam)
        # Recolecta avisos de estirón/empotramiento (mismas reglas SUP/INF).
        chain_avisos = []
        if collision_ctx is not None:
            collision_ctx[u"avisos"] = chain_avisos
        if len(chain) == 1:
            n, err = _place_lateral_on_beam(
                document,
                chain[0],
                bar_type,
                n_lat,
                None,
                collected_rebars=collected,
                lateral_clear_mm=clear_mm,
                collision_ctx=collision_ctx,
            )
        else:
            n, err = _place_lateral_on_beam_aligned_chain(
                document,
                chain,
                bar_type,
                n_lat,
                None,
                collected_rebars=collected,
                lateral_clear_mm=clear_mm,
                collision_ctx=collision_ctx,
            )
        total += int(n or 0)
        if chain_avisos:
            # Deduplicar mensajes de extremos (una vez por cadena).
            for msg in chain_avisos:
                if msg and msg not in avisos:
                    avisos.append(msg)
        if err:
            if n:
                avisos.append(err)
            else:
                try:
                    ids = u", ".join(unicode(e.Id.IntegerValue) for e in chain[:4])
                except Exception:
                    ids = u"?"
                avisos.append(u"[{0}] laterales: {1}".format(ids, err))

    # Al crear, Revit suele dejar Unobscured ON; quitarlo de inmediato en la vista.
    if collected and view is not None and document is not None:
        try:
            from bimtools_rebar_3d_visibility import ensure_rebar_obscured_in_view

            ensure_rebar_obscured_in_view(document, collected, view)
        except Exception:
            pass
        for rb in collected:
            if rb is None:
                continue
            try:
                el = document.GetElement(rb.Id) or rb
            except Exception:
                el = rb
            try:
                el.SetUnobscuredInView(view, False)
            except Exception:
                pass
            try:
                el.SetSolidInView(view, False)
            except Exception:
                pass

    return total, avisos, collected, None
