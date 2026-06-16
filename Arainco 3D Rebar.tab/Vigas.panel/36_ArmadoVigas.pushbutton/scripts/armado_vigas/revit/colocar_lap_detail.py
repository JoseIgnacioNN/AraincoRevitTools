# -*- coding: utf-8 -*-
"""
Detail Items y cotas de traslape @ empalme en Armado Vigas (longitudinales).

Misma familia line-based y criterio de cota que ``geometria_viga_cara_superior_detalle``
y ``enfierrado_shaft_hashtag``; vínculo opcional en ``lap_detail_link_vigas_schema``.
"""

from __future__ import division

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import ElementId, FamilySymbol, XYZ

from armado_vigas.domain.layers import (
    beam_layer_diam_inf,
    beam_layer_diam_sup,
    beam_n_capas_inf,
    beam_n_capas_sup,
    ensure_beam_layers,
)
from armado_vigas.geometry.longitudinales import (
    merged_fiber_line,
    resolve_ref_beam_for_chain,
)
from armado_vigas.revit.rebar_resources import resolve_bar_type_mm

_MAX_WARNINGS = 8
_LAP_DIM_SCALE_REFERENCE = 50
_LAP_DIM_OFFSET_MM_AT_REF_SCALE = 450.0

try:
    from barras_bordes_losa_gancho_empotramiento import _find_fixed_lap_detail_symbol_id
except Exception:
    _find_fixed_lap_detail_symbol_id = None

try:
    from geometria_viga_cara_superior_detalle import (
        _OFFSET_SUPLES_SEGUNDA_CAPA_MM,
        _dedupe_sorted_cut_params,
        _parametros_corte_por_planos_empalme_location,
        _puntos_segmento_traslape_sobre_work,
        _traslapo_longitudinal_mm_desde_bar_type,
        vista_permite_detail_curve,
    )
except Exception:
    _OFFSET_SUPLES_SEGUNDA_CAPA_MM = 50.0
    _dedupe_sorted_cut_params = None
    _parametros_corte_por_planos_empalme_location = None
    _puntos_segmento_traslape_sobre_work = None
    _traslapo_longitudinal_mm_desde_bar_type = None

    def vista_permite_detail_curve(view):
        return view is not None

try:
    from enfierrado_shaft_hashtag import (
        _create_overlap_dimension_from_detail_refs,
        _get_named_left_right_refs_from_detail_instance,
        _place_line_based_detail_component,
        _view_accepts_overlap_dimension,
    )
except Exception:
    _create_overlap_dimension_from_detail_refs = None
    _get_named_left_right_refs_from_detail_instance = None
    _place_line_based_detail_component = None
    _view_accepts_overlap_dimension = None

try:
    from lap_detail_link_vigas_schema import set_lap_detail_vigas_rebar_link
except Exception:
    set_lap_detail_vigas_rebar_link = None


def _element_id_int(element):
    try:
        return int(element.Id.IntegerValue)
    except Exception:
        return None


def _empalme_elements_for_run(sorted_beams, run_indices, empalme_beam_ids):
    elems = []
    seen = set()
    for idx in run_indices or []:
        if idx < 0 or idx >= len(sorted_beams):
            continue
        beam = sorted_beams[idx]
        if beam.get("id") not in (empalme_beam_ids or set()):
            continue
        el = beam.get("element")
        if el is None:
            continue
        eid = _element_id_int(el)
        if eid is not None and eid in seen:
            continue
        if eid is not None:
            seen.add(eid)
        elems.append(el)
    return elems


def _lap_mm_for_layer(document, ref_beam, layer_num, es_cara_inferior, default_lap_mm):
    diam = (
        beam_layer_diam_inf(ref_beam, layer_num)
        if es_cara_inferior
        else beam_layer_diam_sup(ref_beam, layer_num)
    )
    lap_mm = float(default_lap_mm or 0.0)
    if document is not None and diam is not None:
        bt = resolve_bar_type_mm(document, diam)
        if bt is not None and _traslapo_longitudinal_mm_desde_bar_type is not None:
            try:
                lmc, _ = _traslapo_longitudinal_mm_desde_bar_type(bt)
                if lmc and float(lmc) > 0:
                    lap_mm = float(lmc)
            except Exception:
                pass
    return lap_mm


def _view_scale_denominator(view):
    if view is None:
        return _LAP_DIM_SCALE_REFERENCE
    try:
        s = int(view.Scale)
        if s > 0:
            return s
    except Exception:
        pass
    return _LAP_DIM_SCALE_REFERENCE


def _lap_dim_offset_mm(view):
    scale = _view_scale_denominator(view)
    ratio = float(scale) / float(_LAP_DIM_SCALE_REFERENCE)
    return float(_LAP_DIM_OFFSET_MM_AT_REF_SCALE) * ratio


def build_lap_jobs_for_empalme_run(
    document,
    chain,
    run_tramos,
    sorted_beams,
    run_indices,
    empalme_beam_ids,
    es_cara_inferior,
    rex_mm,
    rebar_bar_type,
    lap_mm,
    rebar_by_tramo_layer,
    domain_by_id=None,
):
    """
    Construye specs de empalme a partir de rebars ya colocados por tramo/capa.

    ``rebar_by_tramo_layer``: ``{(es_inf, layer_num, tramo_id): rebar}``.
    """
    if (
        not run_tramos
        or len(run_tramos) < 2
        or not lap_mm
        or float(lap_mm) <= 0
        or _parametros_corte_por_planos_empalme_location is None
        or _puntos_segmento_traslape_sobre_work is None
    ):
        return []

    merged, n_face = merged_fiber_line(
        document, chain, es_cara_inferior, rex_mm, rebar_bar_type
    )
    if merged is None or n_face is None:
        return []

    emp_elems = _empalme_elements_for_run(sorted_beams, run_indices, empalme_beam_ids)
    if not emp_elems:
        return []

    try:
        length = float(merged.Length)
    except Exception:
        return []

    cuts = _parametros_corte_por_planos_empalme_location(merged, emp_elems)
    if _dedupe_sorted_cut_params is not None:
        cuts = _dedupe_sorted_cut_params(cuts, length)
    if not cuts:
        return []

    tramos_ord = sorted(run_tramos, key=lambda t: int(t.get("id") or 0))
    if len(cuts) != len(tramos_ord) - 1:
        return []

    ref_beam = resolve_ref_beam_for_chain(
        chain, domain_by_id or {}, es_cara_inferior
    ) or {}
    ensure_beam_layers(ref_beam)
    n_capas = (
        beam_n_capas_inf(ref_beam)
        if es_cara_inferior
        else beam_n_capas_sup(ref_beam)
    )
    step_mm = float(_OFFSET_SUPLES_SEGUNDA_CAPA_MM)
    es_inf = bool(es_cara_inferior)
    jobs = []

    for j, cut_param in enumerate(cuts):
        tr_lo = tramos_ord[j]
        tr_hi = tramos_ord[j + 1]
        tid_lo = tr_lo.get("id")
        tid_hi = tr_hi.get("id")
        if tid_lo is None or tid_hi is None:
            continue
        for layer_num in range(1, n_capas + 1):
            ra = rebar_by_tramo_layer.get((es_inf, layer_num, tid_lo))
            rb = rebar_by_tramo_layer.get((es_inf, layer_num, tid_hi))
            if ra is None or rb is None:
                continue
            lap_layer = _lap_mm_for_layer(
                document, ref_beam, layer_num, es_cara_inferior, lap_mm
            )
            pa, pb = _puntos_segmento_traslape_sobre_work(
                merged, float(cut_param), lap_layer
            )
            if pa is None or pb is None:
                continue
            off_lap_mm = float(layer_num - 1) * step_mm
            if off_lap_mm > 1e-9 and n_face is not None:
                try:
                    nrm = n_face.Normalize()
                    d_cap = nrm.Multiply(-off_lap_mm / 304.8)
                    pa = pa + d_cap
                    pb = pb + d_cap
                except Exception:
                    pass
            jobs.append({
                u"ra": ra,
                u"rb": rb,
                u"pa": pa,
                u"pb": pb,
                u"n_face": n_face,
                u"layer_num": layer_num,
                u"es_cara_inferior": es_inf,
            })
    return jobs


def _create_lap_dimension(document, view, lap_inst, pa, pb, n_face):
    if (
        lap_inst is None
        or _get_named_left_right_refs_from_detail_instance is None
        or _create_overlap_dimension_from_detail_refs is None
        or _view_accepts_overlap_dimension is None
        or not _view_accepts_overlap_dimension(view)
    ):
        return None, None

    ref_l, ref_r, ref_err = _get_named_left_right_refs_from_detail_instance(lap_inst)
    if ref_l is None or ref_r is None:
        return None, ref_err

    axis_u = None
    try:
        dv = pb - pa
        if dv.GetLength() > 1e-9:
            axis_u = dv.Normalize()
    except Exception:
        axis_u = None

    inward_xy = None
    inward_3d = None
    if n_face is not None:
        try:
            inv = n_face.Negate()
            if inv.GetLength() > 1e-12:
                inward_3d = inv.Normalize()
            inward_xy = XYZ(float(inv.X), float(inv.Y), 0.0)
            if inward_xy.GetLength() > 1e-9:
                inward_xy = inward_xy.Normalize()
            else:
                inward_xy = None
        except Exception:
            inward_xy = None
            inward_3d = None

    ok_dim, msg_dim, dim_data = _create_overlap_dimension_from_detail_refs(
        document,
        view,
        ref_l,
        ref_r,
        pa,
        pb,
        axis_u,
        lateral_hint=None,
        line_offset_mm=_lap_dim_offset_mm(view),
        inward_dir_xy=inward_xy,
        inward_dir_3d=inward_3d,
        use_view_plane_dim_line=True,
        flip_dimension_side=False,
    )
    if not ok_dim:
        return None, msg_dim
    try:
        if dim_data and dim_data.get(u"dim_id"):
            return ElementId(int(dim_data[u"dim_id"])), None
    except Exception:
        pass
    return None, msg_dim


def colocar_marcadores_empalme_vigas(document, view, lap_jobs):
    """
    Coloca Detail Items line-based y cotas de traslape en la vista activa.

    Debe llamarse **dentro** de la transacción del caller (no abre transacción propia).

    Returns:
        ``dict`` con ``n_ok``, ``n_fail``, ``n_dims_ok``, ``n_dims_fail``, ``messages``.
    """
    result = {
        u"n_ok": 0,
        u"n_fail": 0,
        u"n_dims_ok": 0,
        u"n_dims_fail": 0,
        u"messages": [],
        u"elements_created": [],
    }
    if not lap_jobs:
        return result
    if (
        document is None
        or view is None
        or _place_line_based_detail_component is None
        or _find_fixed_lap_detail_symbol_id is None
    ):
        result[u"messages"].append(u"Empalme: módulo de detail/cota no disponible.")
        result[u"n_fail"] = len(lap_jobs)
        return result

    if not vista_permite_detail_curve(view):
        result[u"messages"].append(
            u"Empalme: la vista activa no admite detail components "
            u"(use planta, alzado o sección; no plantilla ni 3D).",
        )
        result[u"n_fail"] = len(lap_jobs)
        return result

    sid, sym_err = _find_fixed_lap_detail_symbol_id(document)
    if sid is None:
        if sym_err:
            result[u"messages"].append(sym_err)
        result[u"n_fail"] = len(lap_jobs)
        return result

    lap_sym = document.GetElement(sid)
    if lap_sym is None or not isinstance(lap_sym, FamilySymbol):
        result[u"messages"].append(u"Empalme: símbolo Detail Item no válido.")
        result[u"n_fail"] = len(lap_jobs)
        return result

    aviso_refs = None
    for spec in lap_jobs:
        pa = spec.get(u"pa")
        pb = spec.get(u"pb")
        ra = spec.get(u"ra")
        rb = spec.get(u"rb")
        n_face = spec.get(u"n_face")
        ok_d, err_d, lap_inst = _place_line_based_detail_component(
            document, view, lap_sym, pa, pb,
        )
        if not ok_d or lap_inst is None:
            result[u"n_fail"] += 1
            if err_d and len(result[u"messages"]) < _MAX_WARNINGS:
                result[u"messages"].append(err_d)
            continue

        result.setdefault(u"elements_created", [])
        if lap_inst is not None:
            result[u"elements_created"].append(lap_inst)

        dim_eid = None
        dim_eid, dim_err = _create_lap_dimension(
            document, view, lap_inst, pa, pb, n_face,
        )
        if dim_eid is not None:
            result[u"n_dims_ok"] += 1
            try:
                dim_el = document.GetElement(dim_eid)
                if dim_el is not None:
                    result[u"elements_created"].append(dim_el)
            except Exception:
                pass
        elif dim_err:
            result[u"n_dims_fail"] += 1
            if u"Left/Right" in (dim_err or u"") and aviso_refs is None:
                aviso_refs = dim_err
            elif len(result[u"messages"]) < _MAX_WARNINGS:
                cara = u"inf" if spec.get(u"es_cara_inferior") else u"sup"
                result[u"messages"].append(
                    u"Cota traslape ({0} capa {1}): {2}".format(
                        cara,
                        spec.get(u"layer_num") or u"?",
                        dim_err,
                    ),
                )

        if (
            ra is not None
            and rb is not None
            and set_lap_detail_vigas_rebar_link is not None
        ):
            try:
                set_lap_detail_vigas_rebar_link(
                    lap_inst, ra.Id, rb.Id, dim_eid,
                )
            except Exception:
                pass

        result[u"n_ok"] += 1

    if aviso_refs:
        result[u"messages"].append(aviso_refs)
    return result