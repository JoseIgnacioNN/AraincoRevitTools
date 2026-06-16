# -*- coding: utf-8 -*-
"""
Colisiones de barras laterales — mismas reglas que longitudinales sup/inf.

``geometria_empotramiento_extremos`` vía :func:`aplicar_colision_extremos_fibra`:
sonda 50 mm, empotramiento en muro/obstáculo, refinamiento en **columna** (+ pata L),
extremo libre con pata L.

La colisión se resuelve sobre el **eje longitudinal** de la cadena (antes del
desplazamiento a la cara del alma); los metadatos alimentan la polilínea L al colocar.
"""

from __future__ import division

from armado_vigas.geometry.colision_fibras import aplicar_colision_extremos_fibra

try:
    from Autodesk.Revit.DB import Line
except Exception:
    Line = None


def build_lateral_collision_ctx(session, chain, diam_mm):
    """Contexto de colisión desde la sesión Armado vigas y la cadena colineal."""
    if session is None:
        return None
    return {
        u"ids_seleccion": list(getattr(session, "all_element_ids", None) or []),
        u"chain_elements": list(chain or []),
        u"beam_candidates": list(getattr(session, "framing_elements", None) or []),
        u"diam_mm": diam_mm,
        u"axis_line_prepared": False,
        u"axis_line": None,
        u"meta_inicio": None,
        u"meta_fin": None,
    }


def _axis_endpoints_sin_cara_lateral(
    p0,
    p1,
    axis,
    axis_off,
    axis_off_p0,
    axis_off_p1,
    host,
    cov_ft,
    bar_diam_ft,
    profile_elems,
):
    """Extremos de eje para colisión (sin offset a cara lateral del alma)."""
    trim_fn = None
    try:
        from armadura_vigas_capas import _axis_cover_trim_endpoints

        trim_fn = _axis_cover_trim_endpoints
    except Exception:
        pass

    if trim_fn is not None and host is not None:
        try:
            pa, pb = trim_fn(
                p0,
                p1,
                host,
                cov_ft,
                bar_diam_ft,
                lateral_offset_xyz=None,
                solid_profile_elems=profile_elems,
            )
            if pa is not None and pb is not None:
                return pa, pb
        except Exception:
            pass

    a0 = float(axis_off) if axis_off_p0 is None else float(axis_off_p0)
    a1 = float(axis_off) if axis_off_p1 is None else float(axis_off_p1)
    m_end = float(cov_ft) + 0.5 * max(float(bar_diam_ft), 1e-6)
    try:
        return p0 + axis * max(a0, m_end), p1 - axis * max(a1, m_end)
    except Exception:
        return None, None


def prepare_lateral_axis_collision(
    document,
    p0,
    p1,
    axis,
    collision_ctx,
    host,
    cov_ft,
    bar_diam_ft,
    profile_elems,
    axis_off,
    axis_off_p0=None,
    axis_off_p1=None,
):
    """
    Resuelve colisión una vez por cadena sobre el eje (igual que fibra sup/inf).

    Guarda ``axis_line``, ``meta_inicio`` y ``meta_fin`` en ``collision_ctx``.
    """
    if collision_ctx is None or Line is None:
        return None, None, None
    if collision_ctx.get(u"axis_line_prepared"):
        return (
            collision_ctx.get(u"axis_line"),
            collision_ctx.get(u"meta_inicio"),
            collision_ctx.get(u"meta_fin"),
        )

    pa, pb = _axis_endpoints_sin_cara_lateral(
        p0,
        p1,
        axis,
        axis_off,
        axis_off_p0,
        axis_off_p1,
        host,
        cov_ft,
        bar_diam_ft,
        profile_elems,
    )
    if pa is None or pb is None:
        collision_ctx[u"axis_line_prepared"] = True
        collision_ctx[u"axis_line"] = None
        return None, None, None

    try:
        ln_axis = Line.CreateBound(pa, pb)
    except Exception:
        collision_ctx[u"axis_line_prepared"] = True
        collision_ctx[u"axis_line"] = None
        return None, None, None

    line_out, meta_i, meta_f = apply_lateral_collision_rules(
        document, ln_axis, collision_ctx
    )
    collision_ctx[u"axis_line_prepared"] = True
    collision_ctx[u"axis_line"] = line_out
    collision_ctx[u"meta_inicio"] = meta_i
    collision_ctx[u"meta_fin"] = meta_f
    return line_out, meta_i, meta_f


def apply_lateral_collision_rules(document, line, collision_ctx):
    """
    Mismo recorte de extremos que fibras sup/inf (columna vs muro vía sonda).

    Returns:
        ``(line, meta_inicio, meta_fin)`` — ``line`` es ``None`` si inválida.
    """
    if line is None or not collision_ctx:
        return line, None, None
    if document is None:
        return line, None, None

    try:
        line_out, meta_i, meta_f = aplicar_colision_extremos_fibra(
            document,
            line,
            collision_ctx.get(u"ids_seleccion"),
            collision_ctx.get(u"chain_elements"),
            collision_ctx.get(u"diam_mm"),
        )
        if line_out is not None:
            return line_out, meta_i, meta_f
        return None, meta_i, meta_f
    except Exception:
        return line, None, None


def prioritize_lateral_hosts(document, hosts_try, line, collision_ctx):
    """Ordena hosts: primero el que contiene el punto medio (igual que longitudinales)."""
    if not hosts_try or line is None or not collision_ctx:
        return hosts_try
    try:
        from armado_vigas.revit.colocar_rebar import _pick_host_for_line
    except Exception:
        return hosts_try

    fallback = None
    for h in hosts_try:
        if h is not None:
            fallback = h
            break
    preferred = _pick_host_for_line(
        line,
        collision_ctx.get(u"beam_candidates"),
        fallback,
    )
    if preferred is None:
        return hosts_try

    out = [preferred]
    seen = set()
    try:
        seen.add(int(preferred.Id.IntegerValue))
    except Exception:
        pass
    for h in hosts_try:
        if h is None:
            continue
        try:
            hid = int(h.Id.IntegerValue)
        except Exception:
            hid = None
        if hid is not None and hid in seen:
            continue
        if hid is not None:
            seen.add(hid)
        out.append(h)
    return out
