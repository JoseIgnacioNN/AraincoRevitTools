# -*- coding: utf-8 -*-
"""
Colisiones de barras laterales — mismas reglas que longitudinales sup/inf.

Orden (idéntico al post-fusión / pre-troceo de ``longitudinales``):

1. Estirón por **columna** ``+(dim/2 − 25 mm)`` + pata L
2. Estirón por **viga no //** a la vista ``+(ancho/2 − 25 mm)`` + pata L
3. Estirón por **muro no //** ``+(espesor/2 − 25 mm)`` + pata L
4. **Muro //** a la vista → empotramiento según Ø (sin pata L)
5. Extremos libres restantes: sonda 50 mm (empotramiento / pata L / columna)

La colisión se resuelve sobre el **eje longitudinal** de la cadena (antes del
desplazamiento a la cara del alma); los metadatos alimentan la polilínea L al colocar.
"""

from __future__ import division

from armado_vigas.geometry.colision_fibras import aplicar_colision_extremos_fibra
from armado_vigas.geometry.longitudinales import (
    _apply_emp_after_parallel_wall,
    _apply_pata_l_after_beam_stretch,
    _apply_pre_troceo_wall_retract,
)

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
        u"stretch_meta": None,
        u"emp_meta": None,
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
    Misma resolución de extremos que longitudinales SUP/INF:

    - post-fusión: estirón col/viga/muro no// (+ pata L) y emp. muro //
    - luego sonda en extremos no marcados (columna / muro / libre)

    Returns:
        ``(line, meta_inicio, meta_fin)`` — ``line`` es ``None`` si inválida.
    """
    if line is None or not collision_ctx:
        return line, None, None
    if document is None:
        return line, None, None

    ids_sel = collision_ctx.get(u"ids_seleccion")
    chain = collision_ctx.get(u"chain_elements")
    try:
        diam_mm = float(collision_ctx.get(u"diam_mm") or 16)
    except Exception:
        diam_mm = 16.0
    if diam_mm <= 1e-9:
        diam_mm = 16.0

    avisos = collision_ctx.get(u"avisos")

    # 1–4) Estirón no// / emp. muro // (reutiliza pipeline de longitudinales).
    work = line
    stretch_meta = {u"start": None, u"end": None, u"applied": False}
    emp_meta = {u"start": None, u"end": None, u"applied": False}
    try:
        work, stretch_meta, emp_meta = _apply_pre_troceo_wall_retract(
            document,
            line,
            ids_sel,
            chain,
            avisos=avisos if isinstance(avisos, list) else None,
        )
    except Exception:
        work = line
        stretch_meta = {u"start": None, u"end": None, u"applied": False}
        emp_meta = {u"start": None, u"end": None, u"applied": False}

    if work is None:
        work = line

    stretch_s = bool(stretch_meta and stretch_meta.get(u"start"))
    stretch_e = bool(stretch_meta and stretch_meta.get(u"end"))
    emp_s = bool(emp_meta and emp_meta.get(u"start"))
    emp_e = bool(emp_meta and emp_meta.get(u"end"))
    # No anular estirón/pata L o emp. muro // con la sonda.
    res_i = not stretch_s and not emp_s
    res_f = not stretch_e and not emp_e

    # 5) Sonda / extremos libres en extremos aún abiertos.
    meta_i = None
    meta_f = None
    line_out = work
    try:
        line_out, meta_i, meta_f = aplicar_colision_extremos_fibra(
            document,
            work,
            ids_sel,
            chain,
            diam_mm,
            resolver_inicio=res_i,
            resolver_fin=res_f,
        )
    except Exception:
        line_out, meta_i, meta_f = work, None, None

    if line_out is None:
        line_out = work

    try:
        meta_i, meta_f = _apply_pata_l_after_beam_stretch(
            line_out, meta_i, meta_f, stretch_meta, diam_mm
        )
    except Exception:
        pass
    try:
        line_out, meta_i, meta_f = _apply_emp_after_parallel_wall(
            line_out, meta_i, meta_f, emp_meta, diam_mm
        )
    except Exception:
        pass

    collision_ctx[u"stretch_meta"] = stretch_meta
    collision_ctx[u"emp_meta"] = emp_meta
    return line_out, meta_i, meta_f


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
