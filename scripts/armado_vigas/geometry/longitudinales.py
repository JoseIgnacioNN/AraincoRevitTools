# -*- coding: utf-8 -*-
"""
Orquestación de guías longitudinales por cadena colineal.

Fusiona fibra sup/inf, aplica capas y extremos empotrado/gancho.
"""

from armado_vigas.domain.layers import (
    beam_layer_diam_inf,
    beam_layer_diam_sup,
    beam_n_capas_inf,
    beam_n_capas_sup,
    ensure_beam_layers,
)
from armado_vigas.domain.suple_inferior import (
    beam_suple_inf_enabled,
    beam_suple_layer_index,
    ensure_beam_suple_inferior,
    trim_line_central_portion,
)
from armado_vigas.domain.suple_superior import (
    SUPLE_END_PCT,
    beam_suple_sup_enabled,
    beam_suple_sup_layer_index,
    compute_suple_sup_segment_specs,
    ensure_beam_suple_superior,
    merged_suple_sup_trim_sides,
    suple_sup_resolver_at_view_side,
    trim_line_end_portion,
    trim_line_view_end_portion,
)
from armado_vigas.geometry.colision_fibras import aplicar_colision_extremos_fibra
from armado_vigas.geometry.extremos import aplicar_extremos_a_linea_fusionada

try:
    from armadura_vigas_capas import _build_collinear_chains_from_elements
except Exception:
    _build_collinear_chains_from_elements = None

try:
    from geometria_viga_cara_superior_detalle import (
        _OFFSET_SUPLES_SEGUNDA_CAPA_MM,
        _TRAZO_INFERIOR_USAR_LOCATION_UNIFICADA,
        _TRAZO_SUPERIOR_USAR_LOCATION_UNIFICADA,
        _curva_armadura_inferior_desde_location_unificada,
        _curva_armadura_inferior_en_fibra,
        _curva_armadura_superior_desde_location_unificada,
        _curva_armadura_superior_en_fibra,
        _dedupe_sorted_cut_params,
        _expand_merged_line_with_location_endpoints,
        _linea_desplazada_mm_reverso_normal_cara,
        _parametros_corte_por_planos_empalme_location,
        _split_line_by_distances_con_traslapos_empalme,
        _traslapo_longitudinal_mm_desde_bar_type,
        _unificar_lineas_colineales,
    )
except Exception:
    _OFFSET_SUPLES_SEGUNDA_CAPA_MM = 50.0
    _TRAZO_INFERIOR_USAR_LOCATION_UNIFICADA = False
    _TRAZO_SUPERIOR_USAR_LOCATION_UNIFICADA = False
    _curva_armadura_inferior_desde_location_unificada = None
    _curva_armadura_inferior_en_fibra = None
    _curva_armadura_superior_desde_location_unificada = None
    _curva_armadura_superior_en_fibra = None
    _dedupe_sorted_cut_params = None
    _expand_merged_line_with_location_endpoints = None
    _linea_desplazada_mm_reverso_normal_cara = None
    _parametros_corte_por_planos_empalme_location = None
    _split_line_by_distances_con_traslapos_empalme = None
    _traslapo_longitudinal_mm_desde_bar_type = None
    _unificar_lineas_colineales = None


def _append_suple_inferior_guide(
    guides,
    avisos,
    ref_beam,
    merged,
    n_face,
    chain,
    tramo_id=None,
):
    """Añade guía de suple inferior (capa n_inf+1, central 80 % de ``merged``) si está activo en ``ref_beam``."""
    if merged is None or n_face is None or ref_beam is None:
        return
    ensure_beam_suple_inferior(ref_beam)
    if not beam_suple_inf_enabled(ref_beam):
        return

    n_capas_inf = beam_n_capas_inf(ref_beam)
    layer_num = beam_suple_layer_index(ref_beam)
    step_mm = float(_OFFSET_SUPLES_SEGUNDA_CAPA_MM)
    off_mm = float(n_capas_inf) * step_mm

    seg = merged
    if off_mm > 1e-9 and _linea_desplazada_mm_reverso_normal_cara is not None:
        try:
            seg = _linea_desplazada_mm_reverso_normal_cara(merged, n_face, off_mm)
        except Exception:
            seg = None
    if seg is None:
        avisos.append(
            u"Suple inf.: sin geometría tras offset capa {0}.".format(layer_num)
        )
        return

    line_trim = trim_line_central_portion(seg)
    if line_trim is None:
        avisos.append(u"Suple inf.: longitud central 80 % inválida.")
        return

    diam = int(ref_beam.get("diamSupleInf") or 16)
    guide = {
        "line": line_trim,
        "meta_start": None,
        "meta_end": None,
        "layer": layer_num,
        "diam_mm": diam,
        "cara": u"inferior",
        "chain": chain,
        "n_face": n_face,
        "es_cara_inferior": True,
        "es_suple_inferior": True,
        "ref_beam": ref_beam,
    }
    if tramo_id is not None:
        guide["tramo_id"] = tramo_id
    guides.append(guide)


def _domain_beam_for_element(element, domain_by_id):
    eid = _element_id_int(element)
    if eid is None:
        return None
    return domain_by_id.get(eid)


def _append_suple_inferior_guides_per_beam(
    guides,
    avisos,
    document,
    beam_elements,
    domain_by_id,
    n_face,
    chain,
    rex_mm=0.0,
    rebar_bar_type=None,
    tramo_id=None,
):
    """Un suple por viga (fibra y toggle independientes)."""
    for elem in beam_elements or []:
        beam = _domain_beam_for_element(elem, domain_by_id)
        if beam is None:
            continue
        ensure_beam_suple_inferior(beam)
        if not beam_suple_inf_enabled(beam):
            continue
        merged_one, n_f = merged_fiber_line(
            document, [elem], True, rex_mm, rebar_bar_type
        )
        nf = n_f if n_f is not None else n_face
        if merged_one is None or nf is None:
            avisos.append(
                u"Suple inf. {0}: sin fibra en viga.".format(beam.get("id") or u"?")
            )
            continue
        _append_suple_inferior_guide(
            guides,
            avisos,
            beam,
            merged_one,
            nf,
            [elem],
            tramo_id=tramo_id,
        )


def resolve_ref_beam_for_chain(chain, domain_by_id, es_cara_inferior=False):
    """
    Viga de referencia para capas en una cadena colineal.

    Usa la que tenga más capas activas en la cara pedida (evita quedarse en 1 capa
    si ``chain[0]`` no es la viga editada en el configurador de tramo).
    """
    best = None
    best_n = -1
    for el in chain or []:
        eid = _element_id_int(el)
        if eid is None:
            continue
        beam = domain_by_id.get(eid)
        if beam is None:
            continue
        ensure_beam_layers(beam)
        n = beam_n_capas_inf(beam) if es_cara_inferior else beam_n_capas_sup(beam)
        if n > best_n:
            best_n = n
            best = beam
    return best


def build_collinear_chains(document, framing_elements):
    if _build_collinear_chains_from_elements is not None:
        try:
            chains = _build_collinear_chains_from_elements(document, framing_elements)
            if chains:
                return chains
        except Exception:
            pass
    return [[e] for e in framing_elements if e is not None]


def merged_fiber_line(document, chain, es_cara_inferior=False, rex_mm=0.0, rebar_bar_type=None):
    """
    Fibra fusionada sup/inf por cadena.

    Respeta :data:`_TRAZO_*_USAR_LOCATION_UNIFICADA` de geometria_viga (por defecto
    ``en_fibra`` por viga + unificación colineal), igual que la herramienta de detalle.
    """
    use_unified = (
        _TRAZO_INFERIOR_USAR_LOCATION_UNIFICADA
        if es_cara_inferior
        else _TRAZO_SUPERIOR_USAR_LOCATION_UNIFICADA
    )
    if use_unified:
        fn = (
            _curva_armadura_inferior_desde_location_unificada
            if es_cara_inferior
            else _curva_armadura_superior_desde_location_unificada
        )
        if fn is None:
            return None, None
        try:
            merged, n_face, _ = fn(document, chain, rex_mm, rebar_bar_type)
            return merged, n_face
        except Exception:
            return None, None

    fn_fibra = (
        _curva_armadura_inferior_en_fibra
        if es_cara_inferior
        else _curva_armadura_superior_en_fibra
    )
    if fn_fibra is None or _unificar_lineas_colineales is None:
        return None, None
    curvas_arm = []
    n_face = None
    for elem in chain or []:
        if elem is None:
            continue
        try:
            ln, n_f, _cara = fn_fibra(document, elem, rex_mm, rebar_bar_type)
        except Exception:
            ln, n_f = None, None
        if ln is None:
            continue
        curvas_arm.append(ln)
        if n_face is None and n_f is not None:
            n_face = n_f
    if not curvas_arm or n_face is None:
        return None, None
    try:
        merged = _unificar_lineas_colineales(curvas_arm, curvas_arm[0])
    except Exception:
        merged = None
    if merged is None:
        return None, None
    if _expand_merged_line_with_location_endpoints is not None:
        try:
            merged = _expand_merged_line_with_location_endpoints(merged, chain)
        except Exception:
            pass
    return merged, n_face


def build_longitudinal_guides_for_chain(
    document,
    chain,
    domain_beams_by_element_id,
    ids_seleccion,
    es_cara_inferior=False,
    rex_mm=0.0,
    rebar_bar_type=None,
    resolver_inicio=True,
    resolver_fin=True,
):
    """
    Returns list of dicts per capa activa en el tramo de referencia:
    ``line``, ``meta_start``, ``meta_end``, ``layer``, ``diam_mm``, ``cara``.
    """
    if not chain:
        return [], [u"Cadena vacía."]

    ref_beam = resolve_ref_beam_for_chain(
        chain, domain_beams_by_element_id, es_cara_inferior
    )
    if ref_beam is None:
        ref_beam = {"nCapasSup": 1, "nCapasInf": 1, "diamSup": 16, "diamInf": 16}
    ensure_beam_layers(ref_beam)
    n_capas = beam_n_capas_inf(ref_beam) if es_cara_inferior else beam_n_capas_sup(ref_beam)

    merged, n_face = merged_fiber_line(
        document, chain, es_cara_inferior, rex_mm, rebar_bar_type
    )
    cara_lbl = u"inferior" if es_cara_inferior else u"superior"
    if merged is None or n_face is None:
        return [], [u"Sin fibra fusionada (cara {0}).".format(cara_lbl)]

    avisos = []
    guides = []
    step_mm = float(_OFFSET_SUPLES_SEGUNDA_CAPA_MM)

    for layer_idx in range(n_capas):
        layer_num = layer_idx + 1
        off_mm = float(layer_idx) * step_mm
        seg = merged
        if off_mm > 1e-9 and _linea_desplazada_mm_reverso_normal_cara is not None:
            try:
                seg = _linea_desplazada_mm_reverso_normal_cara(merged, n_face, off_mm)
            except Exception:
                seg = None
        if seg is None:
            avisos.append(
                u"Capa {0} {1}: sin geometría tras offset {2:.0f} mm.".format(
                    layer_num, cara_lbl, off_mm
                )
            )
            continue

        diam = (
            beam_layer_diam_sup(ref_beam, layer_num)
            if not es_cara_inferior
            else beam_layer_diam_inf(ref_beam, layer_num)
        )

        line_out, meta_i, meta_f = aplicar_colision_extremos_fibra(
            document,
            seg,
            ids_seleccion,
            chain,
            diam,
            resolver_inicio=resolver_inicio,
            resolver_fin=resolver_fin,
        )
        if line_out is None:
            avisos.append(
                u"Capa {0} {1}: línea inválida tras extremos.".format(layer_num, cara_lbl)
            )
            continue

        guides.append({
            "line": line_out,
            "meta_start": meta_i,
            "meta_end": meta_f,
            "layer": layer_num,
            "diam_mm": diam,
            "cara": cara_lbl,
            "chain": chain,
            "n_face": n_face,
            "es_cara_inferior": es_cara_inferior,
        })

    if es_cara_inferior:
        _append_suple_inferior_guides_per_beam(
            guides,
            avisos,
            document,
            chain,
            domain_beams_by_element_id,
            n_face,
            chain,
            rex_mm=rex_mm,
            rebar_bar_type=rebar_bar_type,
        )

    return guides, avisos


def _chain_elements_for_indices(sorted_beams, indices):
    out = []
    for idx in indices or []:
        if idx < 0 or idx >= len(sorted_beams):
            continue
        el = sorted_beams[idx].get("element")
        if el is not None:
            out.append(el)
    return out


def _empalme_framing_for_run(sorted_beams, run_indices, empalme_beam_ids):
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


def _run_needs_empalme_troceo(run_tramos, empalme_beam_ids, split_empalme):
    if not split_empalme or not empalme_beam_ids or not run_tramos:
        return False
    if len(run_tramos) > 1:
        return True
    return any(t.get("fromEmpalme") for t in run_tramos)


def _resolve_traslape_for_face(
    document,
    ref_beam,
    es_cara_inferior=False,
    rebar_bar_type_hint=None,
):
    """
    ``RebarBarType`` y largo de traslape (mm) para la 1ª capa de la cara pedida.

    Returns:
        ``(rebar_bar_type, lap_mm, avisos, diam_mm)``
    """
    cara_tag = u"inf" if es_cara_inferior else u"sup"
    avisos = []
    diam_mm = None
    if ref_beam is not None:
        ensure_beam_layers(ref_beam)
        diam_mm = (
            beam_layer_diam_inf(ref_beam, 1)
            if es_cara_inferior
            else beam_layer_diam_sup(ref_beam, 1)
        )

    bar_type = None
    if document is not None and diam_mm is not None:
        try:
            from armado_vigas.revit.rebar_resources import resolve_bar_type_mm

            bar_type = resolve_bar_type_mm(document, diam_mm)
        except Exception:
            bar_type = None
    if bar_type is None:
        bar_type = rebar_bar_type_hint

    lap_mm = 0.0
    if bar_type is not None and _traslapo_longitudinal_mm_desde_bar_type is not None:
        try:
            lap_mm, lap_txt = _traslapo_longitudinal_mm_desde_bar_type(bar_type)
            if lap_txt:
                avisos.append(lap_txt)
        except Exception:
            lap_mm = 0.0
    elif diam_mm is not None and bar_type is None:
        avisos.append(
            u"Traslape {0}: sin RebarBarType para Ø{1} mm.".format(
                cara_tag, int(diam_mm)
            )
        )

    return bar_type, lap_mm, avisos, diam_mm


def _split_merged_line_at_empalmes(merged, emp_elems, lap_mm):
    """Trocea ``merged`` @ mitad de cada viga empalme, con traslape longitudinal."""
    if (
        merged is None
        or not emp_elems
        or _parametros_corte_por_planos_empalme_location is None
        or _dedupe_sorted_cut_params is None
        or _split_line_by_distances_con_traslapos_empalme is None
    ):
        return [merged] if merged is not None else []

    try:
        length = float(merged.Length)
    except Exception:
        return [merged]

    cuts = _parametros_corte_por_planos_empalme_location(merged, emp_elems)
    cuts = _dedupe_sorted_cut_params(cuts, length)
    if not cuts:
        return [merged]

    segments, _idxs = _split_line_by_distances_con_traslapos_empalme(
        merged, cuts, float(lap_mm or 0.0)
    )
    if not segments:
        return [merged]
    return segments


def build_longitudinal_guides_for_run(
    document,
    chain_elements,
    run_tramos,
    sorted_beams,
    domain_beams_by_element_id,
    empalme_beam_ids,
    ids_seleccion,
    es_cara_inferior=False,
    rex_mm=0.0,
    rebar_bar_type=None,
    split_empalme=True,
):
    """
    Guías longitudinales por tramo Tn de una corrida colineal.

    Si hay empalme @ mitad, trocea la fibra fusionada y aplica traslape entre tramos
    (largo según Ø de la 1ª capa de la cara pedida: sup o inf).

    Returns:
        ``(guides, avisos, lap_mm)`` — ``lap_mm`` es el traslape aplicado al troceo.
    """
    if not chain_elements:
        return [], [u"Cadena vacía."], 0.0
    if not run_tramos:
        return [], [u"Sin tramos Tn en la corrida."], 0.0

    run_indices = []
    for tramo in run_tramos:
        for idx in tramo.get("beamIndices") or []:
            if idx not in run_indices:
                run_indices.append(idx)

    if not _run_needs_empalme_troceo(run_tramos, empalme_beam_ids, split_empalme):
        guides, av = build_longitudinal_guides_for_chain(
            document,
            chain_elements,
            domain_beams_by_element_id,
            ids_seleccion,
            es_cara_inferior=es_cara_inferior,
            rex_mm=rex_mm,
            rebar_bar_type=rebar_bar_type,
        )
        return guides, av, 0.0

    cara_lbl = u"inferior" if es_cara_inferior else u"superior"
    ref_beam_lap = resolve_ref_beam_for_chain(
        chain_elements, domain_beams_by_element_id, es_cara_inferior
    )
    bar_type_face, lap_mm, lap_avisos, diam_mm = _resolve_traslape_for_face(
        document,
        ref_beam_lap,
        es_cara_inferior=es_cara_inferior,
        rebar_bar_type_hint=rebar_bar_type,
    )
    effective_bar_type = bar_type_face or rebar_bar_type

    merged, n_face = merged_fiber_line(
        document, chain_elements, es_cara_inferior, rex_mm, effective_bar_type
    )
    if merged is None or n_face is None:
        return [], [u"Sin fibra fusionada (cara {0}).".format(cara_lbl)], 0.0

    emp_elems = _empalme_framing_for_run(sorted_beams, run_indices, empalme_beam_ids)
    segments = _split_merged_line_at_empalmes(merged, emp_elems, lap_mm)
    avisos = list(lap_avisos or [])
    if lap_mm > 0:
        diam_lbl = int(diam_mm) if diam_mm is not None else u"?"
        avisos.append(
            u"Traslape {0} @ empalme (Ø{1}): ≈ {2:.0f} mm.".format(
                cara_lbl, diam_lbl, float(lap_mm)
            )
        )

    if len(segments) != len(run_tramos):
        avisos.append(
            u"Troceo empalme: {0} tramo(s) Tn ≠ {1} segmento(s); barra continua.".format(
                len(run_tramos), len(segments)
            )
        )
        guides, av = build_longitudinal_guides_for_chain(
            document,
            chain_elements,
            domain_beams_by_element_id,
            ids_seleccion,
            es_cara_inferior=es_cara_inferior,
            rex_mm=rex_mm,
            rebar_bar_type=effective_bar_type,
        )
        return guides, av, 0.0

    guides = []
    step_mm = float(_OFFSET_SUPLES_SEGUNDA_CAPA_MM)
    n_seg = len(segments)

    for seg_idx, (seg, tramo) in enumerate(zip(segments, run_tramos)):
        tramo_chain = _chain_elements_for_indices(
            sorted_beams, tramo.get("beamIndices") or []
        )
        ref_beam = resolve_ref_beam_for_chain(
            tramo_chain, domain_beams_by_element_id, es_cara_inferior
        )
        if ref_beam is None:
            ref_beam = {"nCapasSup": 1, "nCapasInf": 1, "diamSup": 16, "diamInf": 16}
        ensure_beam_layers(ref_beam)
        n_capas = (
            beam_n_capas_inf(ref_beam)
            if es_cara_inferior
            else beam_n_capas_sup(ref_beam)
        )

        resolver_inicio = seg_idx == 0
        resolver_fin = seg_idx == n_seg - 1

        for layer_idx in range(n_capas):
            layer_num = layer_idx + 1
            off_mm = float(layer_idx) * step_mm
            seg_layer = seg
            if off_mm > 1e-9 and _linea_desplazada_mm_reverso_normal_cara is not None:
                try:
                    seg_layer = _linea_desplazada_mm_reverso_normal_cara(
                        seg, n_face, off_mm
                    )
                except Exception:
                    seg_layer = None
            if seg_layer is None:
                avisos.append(
                    u"T{0} capa {1} {2}: sin geometría tras offset.".format(
                        tramo.get("id"), layer_num, cara_lbl
                    )
                )
                continue

            diam = (
                beam_layer_diam_sup(ref_beam, layer_num)
                if not es_cara_inferior
                else beam_layer_diam_inf(ref_beam, layer_num)
            )

            line_out, meta_i, meta_f = aplicar_extremos_a_linea_fusionada(
                document,
                seg_layer,
                ids_seleccion,
                tramo_chain or chain_elements,
                diam,
                resolver_inicio=resolver_inicio,
                resolver_fin=resolver_fin,
            )
            if line_out is None:
                avisos.append(
                    u"T{0} capa {1} {2}: línea inválida tras extremos.".format(
                        tramo.get("id"), layer_num, cara_lbl
                    )
                )
                continue

            guides.append({
                "line": line_out,
                "meta_start": meta_i,
                "meta_end": meta_f,
                "layer": layer_num,
                "diam_mm": diam,
                "cara": cara_lbl,
                "chain": tramo_chain or chain_elements,
                "n_face": n_face,
                "es_cara_inferior": es_cara_inferior,
                "tramo_id": tramo.get("id"),
                "ref_beam": ref_beam,
            })

        if es_cara_inferior:
            _append_suple_inferior_guides_per_beam(
                guides,
                avisos,
                document,
                tramo_chain or chain_elements,
                domain_beams_by_element_id,
                n_face,
                tramo_chain or chain_elements,
                rex_mm=rex_mm,
                rebar_bar_type=effective_bar_type,
                tramo_id=tramo.get("id"),
            )

    return guides, avisos, float(lap_mm or 0.0)


def _element_id_int(element):
    try:
        return int(element.Id.IntegerValue)
    except Exception:
        return None


def _suple_sup_fiber_line(document, elem, beam, rex_mm, rebar_bar_type):
    """Fibra superior desplazada a la capa de suple (n_capas_sup + 1)."""
    if elem is None or beam is None:
        return None, None
    ensure_beam_layers(beam)
    ensure_beam_suple_superior(beam)
    merged_one, n_face = merged_fiber_line(
        document, [elem], False, rex_mm, rebar_bar_type
    )
    if merged_one is None or n_face is None:
        return None, None
    n_capas = beam_n_capas_sup(beam)
    step_mm = float(_OFFSET_SUPLES_SEGUNDA_CAPA_MM)
    off_mm = float(n_capas) * step_mm
    seg = merged_one
    if off_mm > 1e-9 and _linea_desplazada_mm_reverso_normal_cara is not None:
        try:
            seg = _linea_desplazada_mm_reverso_normal_cara(merged_one, n_face, off_mm)
        except Exception:
            seg = None
    return seg, n_face


def _build_merged_suple_sup_line(
    document, beam_a, beam_b, elem_a, elem_b, rex_mm, rebar_bar_type
):
    """Fusiona 25 % en junta consecutiva (lado derecho A + izquierdo B en canvas)."""
    line_a, n_face = _suple_sup_fiber_line(
        document, elem_a, beam_a, rex_mm, rebar_bar_type
    )
    line_b, _ = _suple_sup_fiber_line(
        document, elem_b, beam_b, rex_mm, rebar_bar_type
    )
    if line_a is None or line_b is None or n_face is None:
        return None, None
    from_start_a, from_start_b = merged_suple_sup_trim_sides(beam_a, beam_b)
    ta = trim_line_end_portion(line_a, from_start=from_start_a, pct=SUPLE_END_PCT)
    tb = trim_line_end_portion(line_b, from_start=from_start_b, pct=SUPLE_END_PCT)
    if ta is None or tb is None:
        return None, None
    if _unificar_lineas_colineales is not None:
        try:
            merged = _unificar_lineas_colineales([ta, tb], ta)
            if merged is not None:
                return merged, n_face
        except Exception:
            pass
    try:
        from Autodesk.Revit.DB import Line

        p0 = ta.GetEndPoint(0)
        p1 = tb.GetEndPoint(1)
        if p0.DistanceTo(p1) < 1e-6:
            return None, None
        return Line.CreateBound(p0, p1), n_face
    except Exception:
        return None, None


def _apply_suple_sup_colision_extremos(
    document,
    line,
    ids_seleccion,
    chain,
    diam_mm,
    segment_type,
    avisos,
    beam_label,
    ref_beam=None,
):
    """
    Colisión + emp/gancho en extremos libres del tramo suple (no fusionado).

    ``start`` / ``end`` son lados del canvas; con ``axisReversed`` se invierte
    qué extremo de LocationCurve corresponde a cada lado.
    """
    resolver_inicio, resolver_fin = suple_sup_resolver_at_view_side(
        ref_beam, segment_type
    )
    line_out, meta_i, meta_f = aplicar_colision_extremos_fibra(
        document,
        line,
        ids_seleccion,
        chain,
        diam_mm,
        resolver_inicio=resolver_inicio,
        resolver_fin=resolver_fin,
    )
    if line_out is None:
        avisos.append(
            u"Suple sup. {0} ({1}): línea inválida tras colisión/extremos.".format(
                beam_label or u"?",
                segment_type or u"?",
            )
        )
        return None, None, None
    return line_out, meta_i, meta_f


def build_suple_superior_guides(
    document,
    sorted_beams,
    domain_beams_by_element_id,
    ids_seleccion,
    rex_mm=10.0,
    rebar_bar_type=None,
):
    """
    Guías de suple superior por extremo libre o tramo fusionado entre vigas consecutivas.

    Tramos ``start`` / ``end`` sin fusión: colisión de fibras + extremos emp/gancho
    (mismo postproceso que longitudinales sup.). Tramos ``merged``: geometría directa.

    Returns:
        ``(guides, avisos)``
    """
    beams = list(sorted_beams or [])
    specs = compute_suple_sup_segment_specs(beams)
    if not specs:
        return [], []

    guides = []
    avisos = []
    for spec in specs or []:
        typ = spec.get("type")
        idxs = spec.get("indices") or []
        line = None
        n_face = None
        ref_beam = None
        chain = []
        meta_i = None
        meta_f = None

        if typ == "merged" and len(idxs) >= 2:
            i, j = idxs[0], idxs[1]
            if i >= len(beams) or j >= len(beams):
                continue
            beam_a, beam_b = beams[i], beams[j]
            elem_a = beam_a.get("element")
            elem_b = beam_b.get("element")
            line, n_face = _build_merged_suple_sup_line(
                document,
                beam_a,
                beam_b,
                elem_a,
                elem_b,
                rex_mm,
                rebar_bar_type,
            )
            ref_beam = beam_a
            chain = [e for e in (elem_a, elem_b) if e is not None]
        elif typ == "start" and idxs:
            i = idxs[0]
            if i >= len(beams):
                continue
            ref_beam = beams[i]
            elem = ref_beam.get("element")
            line_f, n_face = _suple_sup_fiber_line(
                document, elem, ref_beam, rex_mm, rebar_bar_type
            )
            line = (
                trim_line_view_end_portion(line_f, ref_beam, u"start", pct=SUPLE_END_PCT)
                if line_f is not None
                else None
            )
            chain = [elem] if elem is not None else []
        elif typ == "end" and idxs:
            i = idxs[0]
            if i >= len(beams):
                continue
            ref_beam = beams[i]
            elem = ref_beam.get("element")
            line_f, n_face = _suple_sup_fiber_line(
                document, elem, ref_beam, rex_mm, rebar_bar_type
            )
            line = (
                trim_line_view_end_portion(line_f, ref_beam, u"end", pct=SUPLE_END_PCT)
                if line_f is not None
                else None
            )
            chain = [elem] if elem is not None else []
        else:
            continue

        if line is None or n_face is None or ref_beam is None:
            avisos.append(
                u"Suple sup. {0}: sin geometría válida.".format(
                    ref_beam.get("id") if ref_beam else u"?"
                )
            )
            continue

        diam = int(ref_beam.get("diamSupleSup") or 16)
        beam_label = ref_beam.get("id") or u"?"

        if typ in (u"start", u"end"):
            line, meta_i, meta_f = _apply_suple_sup_colision_extremos(
                document,
                line,
                ids_seleccion,
                chain,
                diam,
                typ,
                avisos,
                beam_label,
                ref_beam=ref_beam,
            )
            if line is None:
                continue

        guides.append({
            "line": line,
            "meta_start": meta_i,
            "meta_end": meta_f,
            "layer": beam_suple_sup_layer_index(ref_beam),
            "diam_mm": diam,
            "cara": u"superior",
            "chain": chain,
            "n_face": n_face,
            "es_cara_inferior": False,
            "es_suple_superior": True,
            "ref_beam": ref_beam,
            "suple_sup_segment": typ,
        })

    return guides, avisos
