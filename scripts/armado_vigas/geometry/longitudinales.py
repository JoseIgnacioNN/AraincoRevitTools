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
from armado_vigas.domain.constants import EMPALME_LAYER_ALT_LAP_K
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
from armado_vigas.geometry.extremos import (
    aplicar_empotramiento_extremos_marcados,
    aplicar_extremos_a_linea_fusionada,
    mark_pata_l_keep_geometry,
)
from armado_vigas.geometry.retract_muros_noparalelos import (
    aplicar_estiramiento_extremos_columnas,
    aplicar_estiramiento_extremos_vigas_noparalelas,
    aplicar_retracto_extremos_muros_noparalelos,
    detectar_extremos_muros_paralelos,
)

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

    # Evitar doble modelo si la misma viga se procesa en dos tramos (empalme).
    ref_key = _suple_inf_beam_key(ref_beam, chain)
    if ref_key is not None:
        for g in guides or []:
            if not g or not g.get(u"es_suple_inferior"):
                continue
            if _suple_inf_beam_key(g.get(u"ref_beam"), g.get(u"chain")) == ref_key:
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


def _suple_inf_beam_key(ref_beam, chain=None):
    """Clave estable de viga para deduplicar guías de suple INF."""
    if ref_beam is not None:
        bid = ref_beam.get(u"id")
        if bid:
            try:
                return u"id:{0}".format(unicode(bid))
            except Exception:
                return u"id:{0}".format(bid)
        el = ref_beam.get(u"element")
        eid = _element_id_int(el) if el is not None else None
        if eid is not None:
            return u"eid:{0}".format(eid)
    for el in chain or []:
        eid = _element_id_int(el)
        if eid is not None:
            return u"eid:{0}".format(eid)
    return None


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
    """Un suple por viga (fibra y toggle independientes). No duplicar la misma viga."""
    seen_keys = set()
    for elem in beam_elements or []:
        beam = _domain_beam_for_element(elem, domain_by_id)
        if beam is None:
            continue
        ensure_beam_suple_inferior(beam)
        if not beam_suple_inf_enabled(beam):
            continue
        key = _suple_inf_beam_key(beam, [elem])
        if key is not None and key in seen_keys:
            continue
        if key is not None:
            seen_keys.add(key)
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

    Preferencia: primera viga de la cadena con datos de capas (orden natural del
    tramo), con desempate por más capas en la cara pedida.
    """
    best = None
    best_n = -1
    first = None
    for el in chain or []:
        eid = _element_id_int(el)
        if eid is None:
            continue
        beam = domain_by_id.get(eid)
        if beam is None:
            continue
        ensure_beam_layers(beam)
        if first is None:
            first = beam
        n = beam_n_capas_inf(beam) if es_cara_inferior else beam_n_capas_sup(beam)
        if n > best_n:
            best_n = n
            best = beam
    return best or first


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


def _location_curve_midpoint(elem):
    """Punto medio del LocationCurve del framing (o ``None``)."""
    if elem is None:
        return None
    try:
        loc = elem.Location
        crv = loc.Curve if loc is not None else None
        if crv is None:
            return None
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
        return p0 + (p1 - p0).Multiply(0.5)
    except Exception:
        return None


def orient_line_run_left_to_right(line, sorted_beams, run_indices):
    """
    Orienta la línea fusionada para que el param 0 quede a la **izquierda**
    del run (índice menor / u menor) y el 1 a la derecha.

    Si la 1.ª viga de la cadena tiene ``axisReversed``, la unificación colineal
    puede dejar el trazo de derecha→izquierda y el troceo asigna T1 (cfg) a
    la geometría del extremo opuesto.
    """
    if line is None or not run_indices or not sorted_beams:
        return line
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
    except Exception:
        return line

    i0 = int(run_indices[0])
    i1 = int(run_indices[-1])
    if i0 < 0 or i1 < 0 or i0 >= len(sorted_beams) or i1 >= len(sorted_beams):
        return line
    b0 = sorted_beams[i0]
    b1 = sorted_beams[i1]

    mid0 = _location_curve_midpoint(b0.get(u"element"))
    mid1 = _location_curve_midpoint(b1.get(u"element"))
    if mid0 is None or mid1 is None:
        # Fallback: si una sola viga y axisReversed, invertir.
        if i0 == i1 and bool(b0.get(u"axisReversed")):
            try:
                from Autodesk.Revit.DB import Line as _Line

                return _Line.CreateBound(p1, p0)
            except Exception:
                return line
        return line

    try:
        ref = mid1 - mid0
        if float(ref.GetLength()) < 1e-9:
            if bool(b0.get(u"axisReversed")):
                from Autodesk.Revit.DB import Line as _Line

                return _Line.CreateBound(p1, p0)
            return line
        ax = ref.Normalize()
        t0 = float((p0 - mid0).DotProduct(ax))
        t1 = float((p1 - mid0).DotProduct(ax))
        if t0 <= t1 + 1e-9:
            return line
        from Autodesk.Revit.DB import Line as _Line

        flipped = _Line.CreateBound(p1, p0)
        return flipped
    except Exception:
        return line


def assign_segments_to_tramos_spatial(segments, run_tramos, sorted_beams):
    """
    Empareja segmentos de troceo con Tn por orden espacial (izquierda→derecha).

    Evita que `segs[i]↔tramos[i]` desfase si el troceo quedó mal ordenado.
    """
    segs = list(segments or [])
    tramos = list(run_tramos or [])
    if not segs or not tramos or len(segs) != len(tramos):
        return segs

    def _tramo_sort_key(t):
        idxs = t.get(u"beamIndices") or []
        if not idxs:
            return (10 ** 9, 0)
        i0 = int(idxs[0])
        if 0 <= i0 < len(sorted_beams or []):
            b = sorted_beams[i0]
            try:
                u = float(b.get(u"uStart") if b.get(u"uStart") is not None else b.get(u"u") or i0)
            except Exception:
                u = float(i0)
            return (u, i0, int(t.get(u"id") or 0))
        return (float(i0), i0, int(t.get(u"id") or 0))

    def _seg_sort_key(seg):
        try:
            pa = seg.GetEndPoint(0)
            pb = seg.GetEndPoint(1)
            mid = pa + (pb - pa).Multiply(0.5)
            # Proyectar sobre eje del run (primero→último Tn)
            ordered = sorted(tramos, key=_tramo_sort_key)
            b0 = sorted_beams[(ordered[0].get(u"beamIndices") or [0])[0]]
            b1 = sorted_beams[(ordered[-1].get(u"beamIndices") or [0])[-1]]
            m0 = _location_curve_midpoint(b0.get(u"element"))
            m1 = _location_curve_midpoint(b1.get(u"element"))
            if m0 is not None and m1 is not None:
                ax = (m1 - m0).Normalize()
                return float((mid - m0).DotProduct(ax))
            return float(mid.X) + float(mid.Y) + float(mid.Z)
        except Exception:
            return 0.0

    tramos_ord = sorted(tramos, key=_tramo_sort_key)
    segs_ord = sorted(segs, key=_seg_sort_key)
    # Devolver segmentos en el mismo orden que ``run_tramos`` de entrada
    # (puede diferir de tramos_ord si run no estaba sortado).
    id_to_seg = {}
    for t, s in zip(tramos_ord, segs_ord):
        id_to_seg[id(t)] = s
        try:
            id_to_seg[t.get(u"id")] = s
        except Exception:
            pass
    out = []
    for t in tramos:
        s = id_to_seg.get(id(t))
        if s is None:
            s = id_to_seg.get(t.get(u"id"))
        if s is None:
            return segs  # fallo: no reordenar a ciegas
        out.append(s)
    return out


def _merge_stretch_meta(into, other):
    """Combina metas de estirón (viga/muro) por extremo; ``applied`` si hay alguno."""
    if not other or not other.get(u"applied"):
        return into
    out = into or {u"start": None, u"end": None, u"applied": False}
    for key in (u"start", u"end"):
        if other.get(key) and not out.get(key):
            out[key] = other.get(key)
        elif other.get(key) and out.get(key):
            # Si ambos, conservar el de mayor stretch_mm.
            try:
                a = float((out[key] or {}).get(u"stretch_mm") or 0.0)
                b = float((other[key] or {}).get(u"stretch_mm") or 0.0)
            except Exception:
                a, b = 0.0, 0.0
            if b > a:
                out[key] = other.get(key)
    out[u"applied"] = bool(out.get(u"start") or out.get(u"end"))
    return out


def _apply_pre_troceo_wall_retract(document, merged, ids_seleccion, chain, avisos=None, view=None):
    """Post-fusión / pre-troceo: estirón no// (+ pata L) y detección muro //.

    Returns:
        ``(line, stretch_meta, emp_meta)`` —
        stretch → pata L; emp → empotramiento según Ø.
    """
    stretch_meta = {u"start": None, u"end": None, u"applied": False}
    emp_meta = {u"start": None, u"end": None, u"applied": False}
    if merged is None:
        return merged, stretch_meta, emp_meta
    work = merged
    if view is None and document is not None:
        try:
            view = document.ActiveView
        except Exception:
            view = None

    try:
        new_line_c, meta_c = aplicar_estiramiento_extremos_columnas(
            document,
            work,
            ids_seleccion,
            host_chain_elements=chain,
            view=view,
        )
    except Exception:
        new_line_c, meta_c = work, None
    if meta_c and meta_c.get(u"applied"):
        stretch_meta = _merge_stretch_meta(stretch_meta, meta_c)
        if avisos is not None:
            for key, label in ((u"start", u"inicio"), (u"end", u"fin")):
                m = meta_c.get(key)
                if not m:
                    continue
                dim = m.get(u"section_mm") or m.get(u"width_mm") or 0.0
                avisos.append(
                    u"Estirón columna ({0}): +{1:.0f} mm (½ sección {2:.0f} − 25 mm) + pata L.".format(
                        label,
                        float(m.get(u"stretch_mm") or 0.0),
                        float(dim),
                    )
                )
    work = new_line_c if new_line_c is not None else work

    try:
        new_line, meta = aplicar_estiramiento_extremos_vigas_noparalelas(
            document,
            work,
            ids_seleccion,
            host_chain_elements=chain,
            view=view,
        )
    except Exception:
        new_line, meta = work, None
    if meta and meta.get(u"applied"):
        stretch_meta = _merge_stretch_meta(stretch_meta, meta)
        if avisos is not None:
            for key, label in ((u"start", u"inicio"), (u"end", u"fin")):
                m = meta.get(key)
                if not m:
                    continue
                avisos.append(
                    u"Estirón viga no// vista ({0}): +{1:.0f} mm (½ ancho {2:.0f} − 25 mm) + pata L.".format(
                        label,
                        float(m.get(u"stretch_mm") or 0.0),
                        float(m.get(u"width_mm") or 0.0),
                    )
                )
    work = new_line if new_line is not None else work

    try:
        new_line2, meta2 = aplicar_retracto_extremos_muros_noparalelos(
            document,
            work,
            ids_seleccion,
            host_chain_elements=chain,
            view=view,
        )
    except Exception:
        new_line2, meta2 = work, None
    if meta2 and meta2.get(u"applied"):
        stretch_meta = _merge_stretch_meta(stretch_meta, meta2)
        if avisos is not None:
            for key, label in ((u"start", u"inicio"), (u"end", u"fin")):
                m = meta2.get(key)
                if not m:
                    continue
                w_mm = m.get(u"width_mm") or m.get(u"thickness_mm") or 0.0
                avisos.append(
                    u"Estirón muro no// vista ({0}): +{1:.0f} mm (½ ancho {2:.0f} − 25 mm) + pata L.".format(
                        label,
                        float(m.get(u"stretch_mm") or 0.0),
                        float(w_mm),
                    )
                )
    work = new_line2 if new_line2 is not None else work

    skip = []
    if stretch_meta.get(u"start"):
        skip.append(u"start")
    if stretch_meta.get(u"end"):
        skip.append(u"end")
    try:
        emp_meta = detectar_extremos_muros_paralelos(
            document,
            work,
            ids_seleccion,
            host_chain_elements=chain,
            view=view,
            skip_ends=skip,
        )
    except Exception:
        emp_meta = {u"start": None, u"end": None, u"applied": False}
    if avisos is not None and emp_meta and emp_meta.get(u"applied"):
        for key, label in ((u"start", u"inicio"), (u"end", u"fin")):
            if not emp_meta.get(key):
                continue
            avisos.append(
                u"Muro // vista ({0}): empotramiento según Ø.".format(label)
            )

    return work, stretch_meta, emp_meta


def _apply_pata_l_after_beam_stretch(line_out, meta_i, meta_f, stretch_meta, diam_mm):
    """Si hubo estirón por viga/muro/columna, fuerza pata L sin mover la geometría."""
    if not stretch_meta or not stretch_meta.get(u"applied"):
        return meta_i, meta_f
    p0 = p1 = None
    if line_out is not None:
        try:
            p0 = line_out.GetEndPoint(0)
            p1 = line_out.GetEndPoint(1)
        except Exception:
            p0 = p1 = None
    if stretch_meta.get(u"start"):
        meta_i = mark_pata_l_keep_geometry(meta_i, diam_mm, punto=p0)
    if stretch_meta.get(u"end"):
        meta_f = mark_pata_l_keep_geometry(meta_f, diam_mm, punto=p1)
    return meta_i, meta_f


def _apply_emp_after_parallel_wall(line_out, meta_i, meta_f, emp_meta, diam_mm):
    """Estira + marca empotramiento en extremos con muro // a la vista."""
    if not emp_meta or not emp_meta.get(u"applied"):
        return line_out, meta_i, meta_f
    try:
        new_line, m_i, m_f = aplicar_empotramiento_extremos_marcados(
            line_out, emp_meta, diam_mm
        )
    except Exception:
        return line_out, meta_i, meta_f
    if new_line is not None:
        line_out = new_line
    if m_i is not None:
        meta_i = m_i
    if m_f is not None:
        meta_f = m_f
    return line_out, meta_i, meta_f


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
    end_mode_start=None,
    end_mode_end=None,
):
    """
    Returns list of dicts per capa activa en el tramo de referencia:
    ``line``, ``meta_start``, ``meta_end``, ``layer``, ``diam_mm``, ``cara``.

    ``end_mode_start`` / ``end_mode_end``: modos sobre extremos **0/1 de la curva**.
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
    # Pre-troceo / post-fusión: estirón no// (+ pata L) y muro // (emp. según Ø).
    merged, stretch_meta, emp_meta = _apply_pre_troceo_wall_retract(
        document, merged, ids_seleccion, chain, avisos=avisos
    )

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

        # Extremos con estirón/pata L o emp. muro //: no sonda que anule.
        stretch_s = bool(stretch_meta and stretch_meta.get(u"start"))
        stretch_e = bool(stretch_meta and stretch_meta.get(u"end"))
        emp_s = bool(emp_meta and emp_meta.get(u"start"))
        emp_e = bool(emp_meta and emp_meta.get(u"end"))
        res_i = bool(resolver_inicio) and not stretch_s and not emp_s
        res_f = bool(resolver_fin) and not stretch_e and not emp_e
        em_s = end_mode_start if res_i else None
        em_e = end_mode_end if res_f else None

        line_out, meta_i, meta_f = aplicar_colision_extremos_fibra(
            document,
            seg,
            ids_seleccion,
            chain,
            diam,
            resolver_inicio=res_i,
            resolver_fin=res_f,
            end_mode_start=em_s,
            end_mode_end=em_e,
        )
        if line_out is None:
            avisos.append(
                u"Capa {0} {1}: línea inválida tras extremos.".format(layer_num, cara_lbl)
            )
            continue

        meta_i, meta_f = _apply_pata_l_after_beam_stretch(
            line_out, meta_i, meta_f, stretch_meta, diam
        )
        line_out, meta_i, meta_f = _apply_emp_after_parallel_wall(
            line_out, meta_i, meta_f, emp_meta, diam
        )

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
    try:
        from armado_vigas.domain.concrete_lengths import (
            lap_mm_for_diameter,
            session_concrete_grade,
        )

        grade = session_concrete_grade()
    except Exception:
        grade = None
    if diam_mm is not None:
        try:
            L = lap_mm_for_diameter(diam_mm, grade)
            if L is not None and float(L) > 0:
                lap_mm = float(L)
                avisos.append(
                    u"Traslape {0} {1} Ø{2} → {3:.0f} mm".format(
                        cara_tag, grade or u"", int(round(float(diam_mm))), lap_mm
                    )
                )
        except Exception:
            pass
    if lap_mm <= 0 and bar_type is not None and _traslapo_longitudinal_mm_desde_bar_type is not None:
        try:
            try:
                lap_mm, lap_txt = _traslapo_longitudinal_mm_desde_bar_type(
                    bar_type, concrete_grade=grade
                )
            except TypeError:
                lap_mm, lap_txt = _traslapo_longitudinal_mm_desde_bar_type(bar_type)
            if lap_txt:
                avisos.append(lap_txt)
        except Exception:
            lap_mm = 0.0
    elif diam_mm is not None and bar_type is None and lap_mm <= 0:
        avisos.append(
            u"Traslape {0}: sin RebarBarType para Ø{1} mm.".format(
                cara_tag, int(diam_mm)
            )
        )

    return bar_type, lap_mm, avisos, diam_mm


def empalme_cut_shift_ft_for_layer(layer_num, lap_mm, alt_k=None):
    """
    Desfase del corte de empalme a lo largo de la fibra (pies).

    Misma alternancia que el canvas:
      - capas **impares** (1, 3…): sin desfase (centro de viga)
      - capas **pares** (2, 4…): + ``(k/2)·lap_total`` con k=2 ⇒ un solape completo

    ``lap_mm`` es el largo **total** de traslape (tabla Ø).
    """
    try:
        li = max(1, int(layer_num))
    except Exception:
        li = 1
    if (li % 2) == 1:
        return 0.0
    try:
        lap = float(lap_mm or 0.0)
    except Exception:
        return 0.0
    if lap <= 1e-9:
        return 0.0
    try:
        k = float(alt_k if alt_k is not None else EMPALME_LAYER_ALT_LAP_K)
    except Exception:
        k = 2.0
    if k < 1e-9:
        return 0.0
    # k * (lap/2) en mm → pies; k=2 ⇒ lap total.
    return (0.5 * float(k) * lap) / 304.8


def _shift_empalme_cut_params(cuts, shift_ft, length):
    """Aplica desfase a parámetros de corte y re-dedupe dentro de ``(0, L)``."""
    if not cuts:
        return []
    try:
        sh = float(shift_ft or 0.0)
    except Exception:
        sh = 0.0
    try:
        L = float(length)
    except Exception:
        return list(cuts)
    if abs(sh) <= 1e-12:
        out = list(cuts)
    else:
        out = []
        for c in cuts:
            try:
                out.append(float(c) + sh)
            except Exception:
                continue
    if _dedupe_sorted_cut_params is not None:
        return _dedupe_sorted_cut_params(out, L)
    return out


def _shift_empalme_cuts_per_lap(cuts, laps_mm, layer_num, length):
    """Desfase de cada corte con el lap (mm) de **ese** nudo (capas pares)."""
    if not cuts:
        return []
    try:
        L = float(length)
    except Exception:
        return list(cuts)
    out = []
    for j, c in enumerate(cuts):
        try:
            if isinstance(laps_mm, (list, tuple)) and j < len(laps_mm):
                lap_j = float(laps_mm[j] or 0.0)
            else:
                lap_j = float(laps_mm or 0.0)
        except Exception:
            lap_j = 0.0
        sh = empalme_cut_shift_ft_for_layer(layer_num, lap_j)
        try:
            out.append(float(c) + float(sh or 0.0))
        except Exception:
            continue
    if _dedupe_sorted_cut_params is not None:
        return _dedupe_sorted_cut_params(out, L)
    return out


def _base_empalme_cut_params(merged, emp_elems):
    """Cortes base @ mitad de viga empalme (sin alternancia de capa)."""
    if (
        merged is None
        or not emp_elems
        or _parametros_corte_por_planos_empalme_location is None
    ):
        return [], 0.0
    try:
        length = float(merged.Length)
    except Exception:
        return [], 0.0
    try:
        cuts = _parametros_corte_por_planos_empalme_location(merged, emp_elems)
    except Exception:
        return [], length
    if _dedupe_sorted_cut_params is not None:
        cuts = _dedupe_sorted_cut_params(cuts, length)
    return list(cuts or []), length


def _split_merged_line_at_empalmes(merged, emp_elems, lap_mm, layer_num=1, base_cuts=None):
    """
    Trocea ``merged`` @ empalme + traslape, con **alternancia por capa**.

    ``lap_mm``: escalar o lista (un valor por nudo de empalme). Lista = traslape
    del mayor Ø de cada par de tramos consecutivos.

    ``layer_num`` 1-based: capas pares desplazan el corte un solape total
    (misma regla visual del canvas).
    """
    if (
        merged is None
        or _split_line_by_distances_con_traslapos_empalme is None
    ):
        return [merged] if merged is not None else []

    try:
        length = float(merged.Length)
    except Exception:
        return [merged]

    if base_cuts is None:
        cuts, length = _base_empalme_cut_params(merged, emp_elems)
    else:
        cuts = list(base_cuts or [])

    if not cuts:
        return [merged]

    if isinstance(lap_mm, (list, tuple)):
        cuts = _shift_empalme_cuts_per_lap(cuts, lap_mm, layer_num, length)
    else:
        shift_ft = empalme_cut_shift_ft_for_layer(layer_num, lap_mm)
        cuts = _shift_empalme_cut_params(cuts, shift_ft, length)
    if not cuts:
        return [merged]

    segments, _idxs = _split_line_by_distances_con_traslapos_empalme(
        merged, cuts, lap_mm
    )
    if not segments:
        return [merged]
    return segments


def _diam_mm_for_tramo_layer(
    session,
    tramo,
    layer_num,
    es_cara_inferior,
    sorted_beams,
    domain_by_id,
    fallback_beam=None,
):
    """Ø (mm) de capa en un Tn: config por tramo → viga de ref."""
    face = u"inf" if es_cara_inferior else u"sup"
    ref = fallback_beam
    try:
        tr_ch = _chain_elements_for_indices(
            sorted_beams, tramo.get("beamIndices") or []
        )
        rb = resolve_ref_beam_for_chain(tr_ch, domain_by_id, es_cara_inferior)
        if rb is not None:
            ref = rb
    except Exception:
        pass
    if ref is not None:
        ensure_beam_layers(ref)
    try:
        from armado_vigas.domain.tramo_armado import tramo_layer_diam

        d = tramo_layer_diam(
            session, face, tramo, layer_num, es_cara_inferior, ref
        )
        if d is not None:
            return float(d)
    except Exception:
        pass
    if ref is not None:
        try:
            if es_cara_inferior:
                return float(beam_layer_diam_inf(ref, layer_num) or 16)
            return float(beam_layer_diam_sup(ref, layer_num) or 16)
        except Exception:
            pass
    return 16.0


def _lap_mm_list_for_run_layer(
    document,
    run_tramos,
    layer_num,
    es_cara_inferior,
    sorted_beams,
    domain_by_id,
    default_lap_mm,
    session=None,
    n_cuts=None,
):
    """
    Traslape (mm) en cada nudo entre tramos consecutivos de la corrida.

    Si cambian de Ø, usa la tabla del **mayor** de los dos.
    """
    try:
        from armado_vigas.domain.concrete_lengths import (
            lap_mm_for_diameter,
            lap_mm_for_diameter_pair,
            session_concrete_grade,
        )

        grade = session_concrete_grade(session)
    except Exception:
        lap_mm_for_diameter = None
        lap_mm_for_diameter_pair = None
        grade = None

    tramos = list(run_tramos or [])
    try:
        tramos = sorted(
            tramos,
            key=lambda t: (
                min(t.get("beamIndices") or [10 ** 9]),
                int(t.get("id") or 0),
            ),
        )
    except Exception:
        pass
    n_j = max(0, len(tramos) - 1)
    if n_cuts is not None:
        try:
            n_j = min(n_j, int(n_cuts))
        except Exception:
            pass
    if n_j <= 0:
        return []

    diams = []
    for t in tramos:
        diams.append(
            _diam_mm_for_tramo_layer(
                session,
                t,
                layer_num,
                es_cara_inferior,
                sorted_beams,
                domain_by_id,
            )
        )

    try:
        default_lap = float(default_lap_mm or 0.0)
    except Exception:
        default_lap = 0.0

    out = []
    for j in range(n_j):
        da = diams[j]
        db = diams[j + 1] if j + 1 < len(diams) else da
        lap = None
        if lap_mm_for_diameter_pair is not None:
            try:
                lap = lap_mm_for_diameter_pair(da, db, grade)
            except Exception:
                lap = None
        if (lap is None or float(lap) <= 1e-9) and lap_mm_for_diameter is not None:
            try:
                dmax = max(float(da or 0), float(db or 0))
                lap = lap_mm_for_diameter(dmax, grade)
            except Exception:
                lap = None
        if lap is None or float(lap) <= 1e-9:
            # Fallback: lap del mayor Ø vía helper legacy por tipo.
            try:
                dmax = max(float(da or 0), float(db or 0))
            except Exception:
                dmax = 16.0
            lap = default_lap
            if document is not None and dmax > 0 and _traslapo_longitudinal_mm_desde_bar_type is not None:
                try:
                    from armado_vigas.revit.rebar_resources import resolve_bar_type_mm

                    bt = resolve_bar_type_mm(document, dmax)
                    if bt is not None:
                        try:
                            lmc, _ = _traslapo_longitudinal_mm_desde_bar_type(
                                bt, concrete_grade=grade
                            )
                        except TypeError:
                            lmc, _ = _traslapo_longitudinal_mm_desde_bar_type(bt)
                        if lmc and float(lmc) > 0:
                            lap = float(lmc)
                except Exception:
                    pass
        try:
            out.append(float(lap or 0.0))
        except Exception:
            out.append(0.0)
    return out


def _lap_mm_for_layer_num(document, ref_beam, layer_num, es_cara_inferior, default_lap_mm):
    """Traslape (mm) de la capa: Ø de capa → tabla dosificación; fallback ``default_lap_mm``."""
    try:
        lap = float(default_lap_mm or 0.0)
    except Exception:
        lap = 0.0
    if ref_beam is None:
        return lap
    ensure_beam_layers(ref_beam)
    diam = (
        beam_layer_diam_inf(ref_beam, layer_num)
        if es_cara_inferior
        else beam_layer_diam_sup(ref_beam, layer_num)
    )
    if diam is None:
        return lap
    try:
        from armado_vigas.domain.concrete_lengths import (
            lap_mm_for_diameter,
            session_concrete_grade,
        )

        L = lap_mm_for_diameter(diam, session_concrete_grade())
        if L is not None and float(L) > 0:
            return float(L)
    except Exception:
        pass
    if document is None:
        return lap
    if _traslapo_longitudinal_mm_desde_bar_type is None:
        return lap
    try:
        from armado_vigas.revit.rebar_resources import resolve_bar_type_mm
        from armado_vigas.domain.concrete_lengths import session_concrete_grade

        bt = resolve_bar_type_mm(document, diam)
        if bt is None:
            return lap
        try:
            lmc, _ = _traslapo_longitudinal_mm_desde_bar_type(
                bt, concrete_grade=session_concrete_grade()
            )
        except TypeError:
            lmc, _ = _traslapo_longitudinal_mm_desde_bar_type(bt)
        if lmc and float(lmc) > 0:
            return float(lmc)
    except Exception:
        pass
    return lap


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
    end_mode_start=None,
    end_mode_end=None,
    session=None,
):
    """
    Guías longitudinales por tramo Tn de una corrida colineal.

    Si hay empalme @ mitad, trocea la fibra fusionada y aplica traslape entre tramos
    (largo según el **mayor Ø** de tramos adyacentes en cada nudo). Capas pares
    desplazan el nudo de empalme un solape completo (alternancia canvas).

    ``end_mode_start`` / ``end_mode_end``: modos sobre extremos **0/1 de la curva**
    (ya mapeados desde izquierda/derecha de vista si aplica).

    Returns:
        ``(guides, avisos, lap_mm)`` — ``lap_mm`` es el traslape de 1.ª capa.
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
            end_mode_start=end_mode_start,
            end_mode_end=end_mode_end,
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

    # param 0 = izquierda del run (bandas / Tn en canvas); evita T1→geometría de T4.
    merged = orient_line_run_left_to_right(merged, sorted_beams, run_indices)

    avisos = list(lap_avisos or [])
    # Post-fusión / pre-troceo: estirón no// (+ pata L) y muro // (emp. según Ø).
    merged, stretch_meta, emp_meta = _apply_pre_troceo_wall_retract(
        document, merged, ids_seleccion, chain_elements, avisos=avisos
    )
    # Re-orientar tras estirones (por si invierten extremos).
    merged = orient_line_run_left_to_right(merged, sorted_beams, run_indices)

    emp_elems = _empalme_framing_for_run(sorted_beams, run_indices, empalme_beam_ids)
    base_cuts, _len_base = _base_empalme_cut_params(merged, emp_elems)

    # Capas máximas del run (cada tramo puede diferir un poco; cache de troceo por capa).
    n_capas_max = 1
    for tramo in run_tramos:
        tr_ch = _chain_elements_for_indices(
            sorted_beams, tramo.get("beamIndices") or []
        )
        rb = resolve_ref_beam_for_chain(
            tr_ch, domain_beams_by_element_id, es_cara_inferior
        ) or ref_beam_lap
        if rb is not None:
            ensure_beam_layers(rb)
            n_c = (
                beam_n_capas_inf(rb) if es_cara_inferior else beam_n_capas_sup(rb)
            )
            if n_c > n_capas_max:
                n_capas_max = int(n_c)

    # Alinear troceo: orden espacial de tramos (primer índice de viga).
    run_tramos = list(run_tramos)
    try:
        run_tramos = sorted(
            run_tramos,
            key=lambda t: (
                min(t.get("beamIndices") or [10 ** 9]),
                int(t.get("id") or 0),
            ),
        )
    except Exception:
        pass

    segments_by_layer = {}
    max_lap_reported = float(lap_mm or 0.0)
    n_base_cuts = len(base_cuts or [])
    for layer_num in range(1, max(1, int(n_capas_max)) + 1):
        # Por nudo: traslape del mayor Ø de tramos adyacentes (cambio de Ø entre Tn).
        laps_layer = _lap_mm_list_for_run_layer(
            document,
            run_tramos,
            layer_num,
            es_cara_inferior,
            sorted_beams,
            domain_beams_by_element_id,
            lap_mm,
            session=session,
            n_cuts=n_base_cuts,
        )
        if not laps_layer and n_base_cuts > 0:
            lap_uniform = _lap_mm_for_layer_num(
                document,
                ref_beam_lap,
                layer_num,
                es_cara_inferior,
                lap_mm,
            )
            laps_layer = [float(lap_uniform or 0.0)] * n_base_cuts
        try:
            max_l = max(float(x or 0.0) for x in laps_layer) if laps_layer else 0.0
        except Exception:
            max_l = float(lap_mm or 0.0)
        if max_l > max_lap_reported:
            max_lap_reported = max_l
        segs = _split_merged_line_at_empalmes(
            merged,
            emp_elems,
            laps_layer if laps_layer else lap_mm,
            layer_num=layer_num,
            base_cuts=base_cuts,
        )
        segs = assign_segments_to_tramos_spatial(segs, run_tramos, sorted_beams)
        segments_by_layer[layer_num] = segs
        if layer_num == 1 and max_l > 0:
            parts = []
            for j, Lcut in enumerate(laps_layer or []):
                try:
                    parts.append(u"nudo{0}={1:.0f}".format(j + 1, float(Lcut)))
                except Exception:
                    pass
            detail = u" · ".join(parts) if parts else u""
            avisos.append(
                u"Traslape {0} @ empalme (mayor Ø adyacente): max ≈ {1:.0f} mm{2}.".format(
                    cara_lbl,
                    float(max_l),
                    (u" · " + detail) if detail else u"",
                )
            )
        if layer_num % 2 == 0 and segs and len(segs) == len(run_tramos):
            shift_mm = empalme_cut_shift_ft_for_layer(layer_num, max_l) * 304.8
            if shift_mm > 1e-3:
                avisos.append(
                    u"Empalme capa {0}: desfase +{1:.0f} mm (alternancia).".format(
                        layer_num, float(shift_mm)
                    )
                )

    segs_l1 = segments_by_layer.get(1) or []
    if len(segs_l1) != len(run_tramos):
        avisos.append(
            u"Troceo empalme: {0} tramo(s) Tn ≠ {1} segmento(s); barra continua.".format(
                len(run_tramos), len(segs_l1)
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
            end_mode_start=end_mode_start,
            end_mode_end=end_mode_end,
        )
        return guides, av, 0.0

    guides = []
    step_mm = float(_OFFSET_SUPLES_SEGUNDA_CAPA_MM)
    n_seg = len(run_tramos)

    for seg_idx, tramo in enumerate(run_tramos):
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
        stretch_s = bool(
            resolver_inicio and stretch_meta and stretch_meta.get(u"start")
        )
        stretch_e = bool(
            resolver_fin and stretch_meta and stretch_meta.get(u"end")
        )
        emp_s = bool(resolver_inicio and emp_meta and emp_meta.get(u"start"))
        emp_e = bool(resolver_fin and emp_meta and emp_meta.get(u"end"))
        res_i = resolver_inicio and not stretch_s and not emp_s
        res_f = resolver_fin and not stretch_e and not emp_e
        em_s = end_mode_start if res_i else None
        em_e = end_mode_end if res_f else None

        for layer_idx in range(n_capas):
            layer_num = layer_idx + 1
            segs = segments_by_layer.get(layer_num) or segs_l1
            if len(segs) != n_seg:
                # Fallo de alternancia: usar troceo capa 1 sin perder el tramo.
                segs = segs_l1
            if seg_idx >= len(segs):
                continue
            seg = segs[seg_idx]

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
            try:
                from armado_vigas.domain.tramo_armado import tramo_layer_diam

                face_tag = u"inf" if es_cara_inferior else u"sup"
                d_tr = tramo_layer_diam(
                    session, face_tag, tramo, layer_num, es_cara_inferior, ref_beam
                )
                if d_tr is not None:
                    diam = int(d_tr)
            except Exception:
                pass

            line_out, meta_i, meta_f = aplicar_extremos_a_linea_fusionada(
                document,
                seg_layer,
                ids_seleccion,
                tramo_chain or chain_elements,
                diam,
                resolver_inicio=res_i,
                resolver_fin=res_f,
                end_mode_start=em_s,
                end_mode_end=em_e,
            )
            if line_out is None:
                avisos.append(
                    u"T{0} capa {1} {2}: línea inválida tras extremos.".format(
                        tramo.get("id"), layer_num, cara_lbl
                    )
                )
                continue

            seg_stretch = {
                u"start": stretch_meta.get(u"start") if stretch_s else None,
                u"end": stretch_meta.get(u"end") if stretch_e else None,
                u"applied": bool(stretch_s or stretch_e),
            }
            meta_i, meta_f = _apply_pata_l_after_beam_stretch(
                line_out, meta_i, meta_f, seg_stretch, diam
            )
            seg_emp = {
                u"start": emp_meta.get(u"start") if emp_s else None,
                u"end": emp_meta.get(u"end") if emp_e else None,
                u"applied": bool(emp_s or emp_e),
            }
            line_out, meta_i, meta_f = _apply_emp_after_parallel_wall(
                line_out, meta_i, meta_f, seg_emp, diam
            )

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
                "tramo": tramo,
                "ref_beam": ref_beam,
            })

    # Suple INF: una sola guía por viga del run (no por Tn).
    # Con empalme la viga de nudo está en dos tramos; el suple es 80 % de la viga
    # completa y no debe modelarse dos veces.
    if es_cara_inferior:
        _append_suple_inferior_guides_per_beam(
            guides,
            avisos,
            document,
            chain_elements,
            domain_beams_by_element_id,
            n_face,
            chain_elements,
            rex_mm=rex_mm,
            rebar_bar_type=effective_bar_type,
            tramo_id=None,
        )

    return guides, avisos, float(max_lap_reported or lap_mm or 0.0)


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
    """Fusiona L/3 en junta consecutiva (lado derecho A + izquierdo B en canvas)."""
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


def _suple_sup_rebuild_line(p0, p1, template=None):
    """Reconstruye ``Line`` entre extremos (helper local para suple SUP)."""
    if p0 is None or p1 is None:
        return template
    try:
        if p0.DistanceTo(p1) < 1e-9:
            return template
    except Exception:
        return template
    try:
        from Autodesk.Revit.DB import Line

        return Line.CreateBound(p0, p1)
    except Exception:
        return template


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
    Extremos de suple SUP — **misma lógica que el canvas de alzado**:

    - ``start`` / ``end``: solo el extremo de **apoyo** es libre.
    - ``merged``: ambos extremos son corte L/3 interior → sin estirón ni pata L.
    - Pata L **solo** si hubo estirón no// (columna / viga / muro) y no emp //.
    - Emp muro // → desarrollo sin pata L.
    - **Sin** sonda de extremos libres (esa vía mete pata L genérica que el
      alzado no dibuja en cortes L/3 ni en tramos fusionados).
    """
    if line is None:
        return None, None, None

    typ = segment_type or u""
    # Canvas: merged → edgeStart/edgeEnd = half (sin estirón ni pata).
    if typ == u"merged":
        return line, None, None

    resolver_inicio, resolver_fin = suple_sup_resolver_at_view_side(
        ref_beam, typ
    )
    if not resolver_inicio and not resolver_fin:
        return line, None, None

    try:
        p0_orig = line.GetEndPoint(0)
        p1_orig = line.GetEndPoint(1)
    except Exception:
        p0_orig = p1_orig = None

    work = line
    stretch_meta = {u"start": None, u"end": None, u"applied": False}
    emp_meta = {u"start": None, u"end": None, u"applied": False}
    try:
        work, stretch_meta, emp_meta = _apply_pre_troceo_wall_retract(
            document,
            line,
            ids_seleccion,
            chain,
            avisos=avisos if isinstance(avisos, list) else None,
        )
    except Exception:
        work = line
        stretch_meta = {u"start": None, u"end": None, u"applied": False}
        emp_meta = {u"start": None, u"end": None, u"applied": False}

    if work is None:
        work = line

    # Solo el extremo libre del tramo (apoyo) conserva estirón/emp.
    stretch_s = bool(
        resolver_inicio and stretch_meta and stretch_meta.get(u"start")
    )
    stretch_e = bool(
        resolver_fin and stretch_meta and stretch_meta.get(u"end")
    )
    emp_s = bool(resolver_inicio and emp_meta and emp_meta.get(u"start"))
    emp_e = bool(resolver_fin and emp_meta and emp_meta.get(u"end"))

    # Restaurar el corte L/3 (half) si pre-troceo movió el extremo interior.
    try:
        p0_w = work.GetEndPoint(0)
        p1_w = work.GetEndPoint(1)
    except Exception:
        p0_w = p1_w = None
    if p0_orig is not None and p1_orig is not None and p0_w is not None:
        if not resolver_inicio:
            p0_w = p0_orig
        if not resolver_fin:
            p1_w = p1_orig
        rebuilt = _suple_sup_rebuild_line(p0_w, p1_w, template=work)
        if rebuilt is not None:
            work = rebuilt

    line_out = work
    meta_i = None
    meta_f = None

    # Pata L = solo estirón no// (como ``pata = emp<=0.5 and ext>0.5`` en canvas).
    seg_stretch = {
        u"start": stretch_meta.get(u"start") if stretch_s else None,
        u"end": stretch_meta.get(u"end") if stretch_e else None,
        u"applied": bool(stretch_s or stretch_e),
    }
    try:
        meta_i, meta_f = _apply_pata_l_after_beam_stretch(
            line_out, meta_i, meta_f, seg_stretch, diam_mm
        )
    except Exception:
        pass

    seg_emp = {
        u"start": emp_meta.get(u"start") if emp_s else None,
        u"end": emp_meta.get(u"end") if emp_e else None,
        u"applied": bool(emp_s or emp_e),
    }
    try:
        line_out, meta_i, meta_f = _apply_emp_after_parallel_wall(
            line_out, meta_i, meta_f, seg_emp, diam_mm
        )
    except Exception:
        pass

    if line_out is None:
        avisos.append(
            u"Suple sup. {0} ({1}): línea inválida tras emp. muro //.".format(
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
    session=None,
):
    """
    Guías de suple superior por apoyo (L/3) o tramo fusionado.

    Extremos: misma regla que el canvas de alzado (pata L solo con estirón
    no//; muro // → emp.; cortes L/3 y tramos merged sin pata L).

    Returns:
        ``(guides, avisos)``
    """
    from armado_vigas.domain.suple_superior import (
        resolve_suple_sup_arm_for_spec,
        sync_beams_suple_from_apoyo_set,
    )

    beams = list(sorted_beams or [])
    if session is not None:
        try:
            sync_beams_suple_from_apoyo_set(session, beams)
        except Exception:
            pass
    specs = compute_suple_sup_segment_specs(beams, session=session)
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
        pct = float(spec.get("pct") or SUPLE_END_PCT)

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
                trim_line_view_end_portion(line_f, ref_beam, u"start", pct=pct)
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
                trim_line_view_end_portion(line_f, ref_beam, u"end", pct=pct)
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

        n_bars, diam = resolve_suple_sup_arm_for_spec(
            session, spec, fallback_beam=ref_beam
        )
        beam_label = ref_beam.get("id") or u"?"

        # start / end / merged: same end rules as barras SUP (inc. muro //).
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
            "n_bars": n_bars,
            "cara": u"superior",
            "chain": chain,
            "n_face": n_face,
            "es_cara_inferior": False,
            "es_suple_superior": True,
            "ref_beam": ref_beam,
            "suple_sup_segment": typ,
            "apoyo_id": spec.get("apoyo_id"),
        })

    return guides, avisos
