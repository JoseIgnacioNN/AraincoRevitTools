# -*- coding: utf-8 -*-
"""Colocación de Rebar longitudinales (sup/inf, capas, extremos emp/gancho)."""

from __future__ import division

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import Line, XYZ

from armado_vigas.domain.layers import (
    beam_layer_diam_inf,
    beam_layer_diam_sup,
    ensure_beam_layers,
    layer_bar_count,
)
from armado_vigas.domain.tramos import build_bar_runs, build_tramos, sort_beams
from armado_vigas.geometry.longitudinales import (
    build_longitudinal_guides_for_run,
    resolve_ref_beam_for_chain,
    _chain_elements_for_indices,
)
from armado_vigas.revit.pata_l_sketch import aplicar_patas_l_polilinea
from armado_vigas.revit.rebar_resources import resolve_bar_type_mm
from armado_vigas.revit.colocar_lap_detail import build_lap_jobs_for_empalme_run
from geometria_empotramiento_extremos import MODO_GANCHO

try:
    from rebar_fundacion_cara_inferior import (
        aplicar_layout_fixed_number_rebar,
        crear_rebar_desde_curva_linea_con_ganchos,
        _rebar_cantidad_posiciones,
    )
except Exception:
    aplicar_layout_fixed_number_rebar = None
    crear_rebar_desde_curva_linea_con_ganchos = None

    def _rebar_cantidad_posiciones(rebar):
        return 1

try:
    from geometria_viga_cara_superior_detalle import (
        _LAYOUT_ARRAY_SIDE_CLEARANCE_MM,
        _host_framing_para_segmento_rebar,
        _layout_max_spacing_array_length_ft,
        _media_diametro_nominal_rebar_mm,
        _norm_createfromcurves_desde_cara_y_tramo,
        obtener_cara_inferior_framing,
        obtener_cara_superior_framing,
    )
except Exception:
    _LAYOUT_ARRAY_SIDE_CLEARANCE_MM = 50.0
    _host_framing_para_segmento_rebar = None
    _layout_max_spacing_array_length_ft = None
    _media_diametro_nominal_rebar_mm = None
    _norm_createfromcurves_desde_cara_y_tramo = None
    obtener_cara_inferior_framing = None
    obtener_cara_superior_framing = None

try:
    from evaluacion_curva_puntos_obstaculos import _punto_en_volumen_solido
except Exception:
    _punto_en_volumen_solido = None

try:
    from geometria_colision_vigas import obtener_solidos_elemento
except Exception:
    obtener_solidos_elemento = None

try:
    from armadura_vigas_capas import _apply_fixed_number_layout
except Exception:
    _apply_fixed_number_layout = None

def _gancho_en_extremo(meta):
    # Sin meta = extremo interno (p. ej. traslapo @ empalme): barra recta, sin pata L.
    if not meta:
        return False
    if meta.get(u"modo") == MODO_GANCHO:
        return True
    if meta.get(u"pata_l"):
        try:
            return float(meta.get(u"hook_mm") or 0.0) > 0.1
        except Exception:
            return True
    return False


def _pick_host(chain):
    """Respaldo: primera viga válida de la cadena colineal."""
    for el in chain or []:
        if el is not None and el.IsValidObject:
            return el
    return None


def _elemento_contiene_punto(pt, elemento):
    """True si ``pt`` cae dentro del volumen del elemento (sólidos con volumen)."""
    if pt is None or elemento is None:
        return False
    if obtener_solidos_elemento is None or _punto_en_volumen_solido is None:
        return False
    try:
        solids = obtener_solidos_elemento(elemento)
    except Exception:
        return False
    for solid in solids or []:
        try:
            if _punto_en_volumen_solido(solid, pt):
                return True
        except Exception:
            continue
    return False


def _pick_host_for_line(line, beam_candidates, fallback=None):
    """
    Host de armadura por punto medio de la curva de la barra.

    Evalúa el punto contra las vigas del lote inicial (contención en sólido).
    Varias vigas contienen el punto → eje de location más cercano (orden de
    selección como desempate). Sin colisión sólida → bbox + distancia al eje
    (``_host_framing_para_segmento_rebar``) o ``fallback``.
    """
    pt_mid = _punto_medio_linea(line)
    if pt_mid is None:
        return fallback

    hits = []
    for el in beam_candidates or []:
        if el is None:
            continue
        try:
            if not el.IsValidObject:
                continue
        except Exception:
            pass
        if _elemento_contiene_punto(pt_mid, el):
            hits.append(el)

    if hits:
        if len(hits) == 1:
            return hits[0]
        if _host_framing_para_segmento_rebar is not None:
            return _host_framing_para_segmento_rebar(pt_mid, hits, hits[0])
        return hits[0]

    if _host_framing_para_segmento_rebar is not None:
        return _host_framing_para_segmento_rebar(pt_mid, beam_candidates, fallback)
    return fallback


def _tangente_unitaria_linea(line):
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        d = p1 - p0
        if d.GetLength() < 1e-12:
            return None
        return d.Normalize()
    except Exception:
        return None


def _punto_medio_linea(line):
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        return XYZ(
            (p0.X + p1.X) * 0.5,
            (p0.Y + p1.Y) * 0.5,
            (p0.Z + p1.Z) * 0.5,
        )
    except Exception:
        return None


def _transverse_dir(n_face, line):
    t = _tangente_unitaria_linea(line)
    if t is None or n_face is None:
        return None
    try:
        v = n_face.CrossProduct(t)
        if v.GetLength() < 1e-12:
            return None
        return v.Normalize()
    except Exception:
        return None


def _layout_v_dir(host, line, n_face):
    """
    Eje transversal de reparto común sup/inf (``width_dir`` en ``armadura_vigas_capas``).
    Siempre derivado de la cara superior para que n=2 y n=3 repartan igual en ambas capas.
    """
    if host is not None:
        sup_face = _cara_framing(host, False)
        if sup_face is not None:
            try:
                n_sup = sup_face.FaceNormal.Normalize()
                v_sup = _transverse_dir(n_sup, line)
                if v_sup is not None:
                    return v_sup
            except Exception:
                pass
    return _transverse_dir(n_face, line)


def _cara_framing(host, es_cara_inferior):
    if host is None:
        return None
    try:
        if es_cara_inferior and obtener_cara_inferior_framing is not None:
            return obtener_cara_inferior_framing(host)
        if not es_cara_inferior and obtener_cara_superior_framing is not None:
            return obtener_cara_superior_framing(host)
    except Exception:
        pass
    return None


def _face_span_along_dir_ft(planar_face, v_dir):
    """Ancho real de la cara (pies) medido a lo largo de ``v_dir``."""
    if planar_face is None or v_dir is None:
        return None
    min_s = None
    max_s = None
    ref = None
    try:
        edge_loops = planar_face.EdgeLoops
    except Exception:
        return None
    try:
        for eloop in edge_loops:
            for edge in eloop:
                try:
                    pts = edge.Tessellate()
                except Exception:
                    continue
                for pt in pts or []:
                    if ref is None:
                        ref = pt
                    try:
                        s = float((pt - ref).DotProduct(v_dir))
                    except Exception:
                        continue
                    if min_s is None or s < min_s:
                        min_s = s
                    if max_s is None or s > max_s:
                        max_s = s
    except Exception:
        return None
    if min_s is None or max_s is None or ref is None:
        return None
    width_ft = float(max_s - min_s)
    if width_ft < 1e-9:
        return None
    center_s = 0.5 * (float(min_s) + float(max_s))
    return width_ft, center_s, ref, float(min_s), float(max_s)


def _get_bar_position_transform(rebar, bar_index):
    try:
        return rebar.GetBarPositionTransform(int(bar_index))
    except Exception:
        pass
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is not None and hasattr(acc, "GetBarPositionTransform"):
            return acc.GetBarPositionTransform(int(bar_index))
    except Exception:
        pass
    return None


def _layout_margin_ft(bar_type, rex_mm):
    """Recubrimiento transversal mínimo (coherente con offsets geometria_viga)."""
    half_long_mm = 0.0
    if _media_diametro_nominal_rebar_mm is not None:
        half_long_mm = float(_media_diametro_nominal_rebar_mm(bar_type) or 0.0)
    return (25.0 + max(0.0, float(rex_mm or 0.0)) + half_long_mm) / 304.8


def _bars_inside_face_span(rebar, v_dir, ref, min_s, max_s, tol_ft=None):
    """
    True si las posiciones del conjunto caen dentro de ``[min_s, max_s]``.

    Si no se pueden leer transforms (API), devuelve ``None`` (= no validar).
    """
    if tol_ft is None:
        tol_ft = 3.0 / 304.8
    try:
        n = int(_rebar_cantidad_posiciones(rebar))
    except Exception:
        return None
    if n < 1 or v_dir is None or ref is None or min_s is None or max_s is None:
        return None
    lo = float(min_s) + float(tol_ft)
    hi = float(max_s) - float(tol_ft)
    if hi < lo:
        return None
    checked = 0
    for i in range(n):
        tr = _get_bar_position_transform(rebar, i)
        if tr is None:
            continue
        checked += 1
        try:
            s = float((tr.Origin - ref).DotProduct(v_dir))
        except Exception:
            return None
        if s < lo - 1e-4 or s > hi + 1e-4:
            return False
    if checked == 0:
        return None
    return True


def _apply_fixed_number_layout_qty(rebar, document, n_bars, span_ft):
    """Layout fijo aceptando la 1.ª combinación con cantidad correcta."""
    if rebar is None or n_bars <= 1:
        return False
    if _apply_fixed_number_layout is not None:
        if _apply_fixed_number_layout(rebar, n_bars, float(span_ft)):
            return True
    if aplicar_layout_fixed_number_rebar is not None and document is not None:
        ok, _err = aplicar_layout_fixed_number_rebar(
            rebar, document, int(n_bars), float(span_ft)
        )
        return bool(ok)
    return False


def _apply_fixed_number_layout_inward(rebar, document, n_bars, span_ft, v_dir, ref, min_s, max_s):
    """
    Layout *Fixed Number*: prioriza orientación con barras dentro de la cara;
    si no puede validarse, acepta layout con cantidad correcta (no borrar todo).
    """
    if rebar is None or n_bars <= 1:
        return False
    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception:
        acc = None
    if acc is None:
        return _apply_fixed_number_layout_qty(rebar, document, n_bars, span_ft)

    combos = (
        (True, True, True),
        (False, True, True),
        (True, False, False),
        (False, False, False),
    )

    def _regen():
        if document is not None:
            try:
                document.Regenerate()
            except Exception:
                pass

    def _qty_ok():
        try:
            return int(_rebar_cantidad_posiciones(rebar)) == int(n_bars)
        except Exception:
            return False

    def _try_passes(require_inside):
        qty_match = None
        for b_side, inc0, inc1 in combos:
            try:
                acc.SetLayoutAsFixedNumber(
                    int(n_bars), float(span_ft), b_side, inc0, inc1
                )
            except Exception:
                continue
            _regen()
            if not _qty_ok():
                continue
            if not require_inside:
                return True
            inside = _bars_inside_face_span(rebar, v_dir, ref, min_s, max_s)
            if inside is True:
                return True
            if inside is None and qty_match is None:
                qty_match = True
        if require_inside and qty_match:
            return True
        return False

    if _try_passes(True):
        return True
    try:
        acc.FlipRebarSet()
    except Exception:
        pass
    _regen()
    if _try_passes(True):
        return True
    return _apply_fixed_number_layout_qty(rebar, document, n_bars, span_ft)


def _span_ft_from_width_mm(width_mm, bar_type, rex_mm):
    try:
        w_mm = max(0.0, float(width_mm))
        cle = float(_LAYOUT_ARRAY_SIDE_CLEARANCE_MM)
        d_est = max(0.0, float(rex_mm or 0.0))
        d_long_mm = 0.0
        if _media_diametro_nominal_rebar_mm is not None:
            d_long_mm = 2.0 * float(_media_diametro_nominal_rebar_mm(bar_type) or 0.0)
        span_mm = max(0.0, w_mm - cle - 2.0 * d_est - d_long_mm)
        return span_mm / 304.8
    except Exception:
        return 0.0


def _shift_line_along_dir(line, v_dir, delta_ft):
    if line is None or v_dir is None:
        return line
    try:
        d = v_dir.Multiply(float(delta_ft))
        p0 = line.GetEndPoint(0) + d
        p1 = line.GetEndPoint(1) + d
        if p0.DistanceTo(p1) < 1e-9:
            return line
        return Line.CreateBound(p0, p1)
    except Exception:
        return line


def _width_center_at_guide(host, line, v_dir, es_inf):
    """Centro transversal de la cara en el punto medio longitudinal de la guía."""
    face = _cara_framing(host, es_inf)
    if face is None or line is None or v_dir is None:
        return None
    data = _face_span_along_dir_ft(face, v_dir)
    if data is None:
        return None
    _width_ft, center_s, ref, _min_s, _max_s = data
    mid = _punto_medio_linea(line)
    if mid is None:
        return None
    try:
        guide_s = float((mid - ref).DotProduct(v_dir))
        delta = float(center_s) - guide_s
        return mid + v_dir.Multiply(delta)
    except Exception:
        return None


def _line_seed_at_array_start(line, v_dir, array_len_ft, host, es_inf):
    """
    Semilla en ``v_ref = −half`` (``armadura_vigas_capas``): 1.ª barra en un alero;
    ``SetLayoutAsFixedNumber(n, array_len)`` reparte hasta el alero opuesto.
    Misma lógica para n=2, n=3, …; solo cambia la cantidad.
    """
    if line is None or v_dir is None:
        return line
    try:
        array_len = float(array_len_ft)
    except Exception:
        return line
    if array_len < 1e-12:
        return line
    half = array_len * 0.5
    center = _width_center_at_guide(host, line, v_dir, es_inf)
    mid = _punto_medio_linea(line)
    if center is None or mid is None:
        return _shift_line_along_dir(line, v_dir, -array_len)
    try:
        seed = center + v_dir.Multiply(-half)
        along = float((seed - mid).DotProduct(v_dir))
    except Exception:
        return _shift_line_along_dir(line, v_dir, -array_len)
    return _shift_line_along_dir(line, v_dir, along)


def _layout_bounds_relative_to_seed(span_ft, margin_ft):
    """Rango útil ``[0, span]`` relativo a la semilla (1.ª barra), con tolerancia."""
    tol = max(float(margin_ft or 0.0), 3.0 / 304.8)
    return -tol, float(span_ft) + tol


def _prepare_line_and_layout_span(document, host, line, n_face, bar_type, rex_mm, es_inf, n_bars, diam_mm=16):
    """
    Span transversal (``array_len = 2·half``) y semilla en ``−half`` respecto al centro
    del ancho — mismo criterio que ``armadura_vigas_capas`` / escenario n=2.
    """
    if line is None or n_bars <= 1:
        return line, 0.0, None, None, None, None, 0.0

    v_dir = _layout_v_dir(host, line, n_face)
    margin_ft = _layout_margin_ft(bar_type, rex_mm)
    array_len_ft = _layout_span_ft(document, host, bar_type, n_bars, diam_mm, rex_mm)

    if v_dir is None or array_len_ft < 1e-12:
        return line, float(array_len_ft or 0.0), None, None, None, None, margin_ft

    line = _line_seed_at_array_start(line, v_dir, array_len_ft, host, es_inf)
    ref = _punto_medio_linea(line)
    if ref is None:
        return line, float(array_len_ft), v_dir, None, None, None, margin_ft

    min_s, max_s = _layout_bounds_relative_to_seed(array_len_ft, margin_ft)
    return line, float(array_len_ft), v_dir, ref, min_s, max_s, margin_ft


def _norm_prioridad(n_face, line, v_dir=None):
    """Normal de layout = ``n_face × tangente`` (``armadura_vigas_capas.width_dir``)."""
    if v_dir is not None:
        try:
            return v_dir.Normalize()
        except Exception:
            pass
    if _norm_createfromcurves_desde_cara_y_tramo is None:
        return None
    return _norm_createfromcurves_desde_cara_y_tramo(n_face, line)


def _norms_for_create(n_face, line, v_dir=None, es_inf=False):
    """
    Normal(es) para ``CreateFromCurves*``. Misma ``width_dir`` sup/inf que el layout
    (``armadura_vigas_capas``); ``es_inf`` no invierte el eje de reparto.
    """
    primary = _norm_prioridad(n_face, line, v_dir)
    if primary is None and v_dir is not None:
        try:
            primary = v_dir.Normalize()
        except Exception:
            primary = None
    if primary is not None:
        try:
            neg = primary.Negate()
            return [primary, neg]
        except Exception:
            return [primary]
    return []


def _crear_rebar_recto(document, host, bar_type, line, es_inf, norm_list):
    """Barra longitudinal recta sin ``RebarHookType`` (patas L = polilínea posterior)."""
    base = dict(
        document=document,
        host=host,
        bar_type=bar_type,
        curve=line,
        normales_prioridad=norm_list or None,
        gancho_en_inicio=False,
        gancho_en_fin=False,
    )
    try:
        return crear_rebar_desde_curva_linea_con_ganchos(es_capa_inferior=False, **base)
    except TypeError:
        return crear_rebar_desde_curva_linea_con_ganchos(**base)


def _ensure_inferior_array_inward(rebar, document, es_inf, v_dir, ref, min_s, max_s):
    """Voltea el conjunto si las barras quedaron fuera del rango transversal."""
    if rebar is None:
        return
    inside = _bars_inside_face_span(rebar, v_dir, ref, min_s, max_s)
    if inside is not False:
        return
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is None:
            return
        acc.FlipRebarSet()
        if document is not None:
            try:
                document.Regenerate()
            except Exception:
                pass
    except Exception:
        pass
    if _bars_inside_face_span(rebar, v_dir, ref, min_s, max_s) is False:
        try:
            acc = rebar.GetShapeDrivenAccessor()
            acc.FlipRebarSet()
            if document is not None:
                document.Regenerate()
        except Exception:
            pass


def _layout_span_ft(document, host, bar_type, n_bars, diam_mm, rex_mm):
    span_ft = None
    if _layout_max_spacing_array_length_ft is not None:
        try:
            span_ft = _layout_max_spacing_array_length_ft(
                document,
                host,
                bar_type,
                diametro_estribo_mm=rex_mm,
            )
        except Exception:
            span_ft = None
    if n_bars > 1 and (span_ft is None or float(span_ft) < 1e-12):
        sep_mm = max(25.0, float(diam_mm or 16) + 25.0)
        span_ft = (sep_mm * float(n_bars - 1)) / 304.8
    return span_ft or 0.0


def place_guide_as_rebar_set(document, guide, host, n_bars, rex_mm=0.0):
    """
    Crea un conjunto Rebar desde una guía longitudinal ya resuelta (línea + metadatos extremos).

    Returns:
        ``(cantidad_barras_creadas, mensaje_error_o_None, rebar_o_None)``
    """
    if crear_rebar_desde_curva_linea_con_ganchos is None:
        return 0, u"Módulo rebar_fundacion_cara_inferior no disponible.", None
    if document is None or host is None or not guide:
        return 0, u"Host o guía inválidos.", None

    line = guide.get(u"line")
    if line is None or not isinstance(line, Line):
        return 0, u"Guía sin línea válida.", None

    diam_mm = guide.get(u"diam_mm") or 16
    bar_type = resolve_bar_type_mm(document, diam_mm)
    if bar_type is None:
        return 0, u"No hay RebarBarType ø{0} mm.".format(int(round(diam_mm))), None

    n_bars = max(1, int(n_bars or 1))
    meta_i = guide.get(u"meta_start")
    meta_f = guide.get(u"meta_end")
    gi = _gancho_en_extremo(meta_i)
    gf = _gancho_en_extremo(meta_f)
    es_inf = bool(guide.get(u"es_cara_inferior"))
    n_face = guide.get(u"n_face")

    line_use, span_ft, v_dir, face_ref, face_min_s, face_max_s, margin_ft = _prepare_line_and_layout_span(
        document,
        host,
        line,
        n_face,
        bar_type,
        rex_mm,
        es_inf,
        n_bars,
        diam_mm=diam_mm,
    )

    norm_list = _norms_for_create(n_face, line_use, v_dir, es_inf=es_inf)

    rb, err, _nv = _crear_rebar_recto(
        document,
        host,
        bar_type,
        line_use,
        es_inf,
        norm_list,
    )
    if rb is None:
        return 0, err or u"No se pudo crear Rebar.", None

    layout_ok = False
    err_lay = None
    if n_bars > 1:
        layout_ok = _apply_fixed_number_layout_qty(
            rb, document, n_bars, float(span_ft),
        )
        if not layout_ok:
            err_lay = u"SetLayoutAsFixedNumber falló."

        if not layout_ok:
            try:
                document.Delete(rb.Id)
            except Exception:
                pass
            return 0, err_lay or u"SetLayoutAsFixedNumber falló.", None
        try:
            qty_check = int(_rebar_cantidad_posiciones(rb))
        except Exception:
            qty_check = 0
        if qty_check != n_bars:
            try:
                document.Delete(rb.Id)
            except Exception:
                pass
            return 0, u"Layout fijo: cantidad {0} ≠ {1}.".format(qty_check, n_bars), None

        _ensure_inferior_array_inward(
            rb, document, es_inf, v_dir, face_ref, face_min_s, face_max_s
        )

    if gi or gf:
        rb_nuevo, err_pata = aplicar_patas_l_polilinea(
            document,
            rb,
            host,
            n_face,
            gi,
            gf,
            meta_inicio=meta_i,
            meta_fin=meta_f,
            diam_mm=diam_mm,
        )
        if err_pata:
            try:
                document.Delete(rb.Id)
            except Exception:
                pass
            return 0, err_pata, None
        rb = rb_nuevo

    try:
        qty = int(_rebar_cantidad_posiciones(rb))
    except Exception:
        qty = n_bars
    return max(1, qty), None, rb


def _tramos_for_bar_run(tramos, run_indices):
    idx_set = set(run_indices or [])
    out = []
    for tramo in tramos or []:
        idxs = tramo.get("beamIndices") or []
        if idxs and all(i in idx_set for i in idxs):
            out.append(tramo)
    return out


def colocar_armadura_longitudinal(document, session):
    """
    Coloca Rebar longitudinales sup/inf por tramo Tn (corrida colineal), capas y extremos.

    Respeta empalme @ mitad de viga: troceo + traslape entre tramos consecutivos.

    Returns:
        ``(n_barras_total, lista_avisos, lista_rebars_creados, rebars_by_side, lap_jobs)``

        ``rebars_by_side``: ``{"sup": [...], "inf": [...]}`` para etiquetado por cara.
        ``lap_jobs``: specs para detail/cota de traslape @ empalme.
    """
    if not session.framing_elements:
        return 0, [u"Sin vigas en el lote."], [], {u"sup": [], u"inf": []}, []

    beam_candidates = list(session.framing_elements or [])
    ids_sel = session.all_element_ids
    domain_by_id = session.domain_beams_by_element_id or {}
    sorted_beams = sort_beams(list(session.domain_beams or []))
    n_total = 0
    avisos = []
    rebars_creados = []
    rebars_sup = []
    rebars_inf = []
    lap_face_mm = {False: None, True: None}
    lap_jobs_all = []

    for es_inf in (False, True):
        tramos = (
            (session.tramos_inf if es_inf else session.tramos_sup)
            or build_tramos(
                sorted_beams,
                session.empalme_beam_ids_inf if es_inf else session.empalme_beam_ids_sup,
                session.split_empalme,
                es_cara_inferior=es_inf,
            )
        )
        empalme_ids = (
            session.empalme_beam_ids_inf if es_inf else session.empalme_beam_ids_sup
        )
        for run in build_bar_runs(sorted_beams, es_cara_inferior=es_inf):
            run_indices = list(run.get("indices") or [])
            chain = _chain_elements_for_indices(sorted_beams, run_indices)
            if not chain:
                avisos.append(u"Corrida sin vigas Revit.")
                continue

            run_tramos = _tramos_for_bar_run(tramos, run_indices)
            if not run_tramos:
                run_tramos = [{
                    "id": None,
                    "beamIndices": run_indices,
                    "fromEmpalme": False,
                }]

            fallback_host = _pick_host(chain)
            if fallback_host is None:
                avisos.append(u"Corrida sin host de armadura.")
                continue

            rex_mm = 10.0
            ref_any = resolve_ref_beam_for_chain(chain, domain_by_id, es_inf)
            if ref_any is None:
                ref_any = resolve_ref_beam_for_chain(chain, domain_by_id, not es_inf)
            if ref_any is not None:
                rex_mm = float(ref_any.get(u"estExtDiam") or 10)

            ref_beam = resolve_ref_beam_for_chain(chain, domain_by_id, es_inf) or ref_any or {}
            ensure_beam_layers(ref_beam)
            diam_face = (
                beam_layer_diam_inf(ref_beam, 1)
                if es_inf
                else beam_layer_diam_sup(ref_beam, 1)
            )
            guides, av, lap_mm = build_longitudinal_guides_for_run(
                document,
                chain,
                run_tramos,
                sorted_beams,
                domain_by_id,
                empalme_ids,
                ids_sel,
                es_cara_inferior=es_inf,
                rex_mm=rex_mm,
                rebar_bar_type=resolve_bar_type_mm(document, diam_face),
                split_empalme=session.split_empalme,
            )
            if lap_mm and lap_mm > 0:
                lap_val = float(lap_mm)
                prev = lap_face_mm.get(es_inf)
                if prev is None or lap_val > prev:
                    lap_face_mm[es_inf] = lap_val
            avisos.extend(av or [])
            run_rebar_map = {}
            for guide in guides or []:
                layer_num = int(guide.get(u"layer") or 1)
                guide_ref = guide.get(u"ref_beam") or ref_beam
                ensure_beam_layers(guide_ref)
                es_suple = bool(guide.get(u"es_suple_inferior"))
                if es_suple:
                    n_bars = max(1, int(guide_ref.get(u"nSupleInf") or 2))
                else:
                    n_bars = layer_bar_count(guide_ref, layer_num, es_inf)
                host = _pick_host_for_line(
                    guide.get(u"line"),
                    beam_candidates,
                    fallback_host,
                )
                if host is None:
                    avisos.append(
                        u"{0} capa {1}: sin host para la barra.".format(
                            guide.get(u"cara") or u"?",
                            layer_num,
                        )
                    )
                    continue
                qty, err, rb = place_guide_as_rebar_set(
                    document,
                    guide,
                    host,
                    n_bars,
                    rex_mm=rex_mm,
                )
                if qty > 0:
                    n_total += qty
                    if rb is not None:
                        rebars_creados.append(rb)
                        if es_inf:
                            rebars_inf.append(rb)
                        else:
                            rebars_sup.append(rb)
                        tramo_id = guide.get(u"tramo_id")
                        if (
                            tramo_id is not None
                            and not es_suple
                        ):
                            run_rebar_map[(es_inf, layer_num, tramo_id)] = rb
                elif err:
                    avisos.append(
                        u"{0} capa {1}: {2}".format(
                            guide.get(u"cara") or u"?",
                            layer_num,
                            err,
                        )
                    )

            if lap_mm and float(lap_mm) > 0 and len(run_tramos) > 1 and run_rebar_map:
                bar_type_face = resolve_bar_type_mm(document, diam_face)
                lap_jobs_all.extend(
                    build_lap_jobs_for_empalme_run(
                        document,
                        chain,
                        run_tramos,
                        sorted_beams,
                        run_indices,
                        empalme_ids,
                        es_inf,
                        rex_mm,
                        bar_type_face,
                        lap_mm,
                        run_rebar_map,
                        domain_by_id=domain_by_id,
                    )
                )

    lap_sup = lap_face_mm.get(False)
    lap_inf = lap_face_mm.get(True)
    if lap_sup and lap_inf:
        avisos.insert(
            0,
            u"Traslape sup: {0:.0f} mm · inf: {1:.0f} mm.".format(lap_sup, lap_inf),
        )
    elif lap_sup:
        avisos.insert(0, u"Traslape sup @ empalme: {0:.0f} mm.".format(lap_sup))
    elif lap_inf:
        avisos.insert(0, u"Traslape inf @ empalme: {0:.0f} mm.".format(lap_inf))

    return n_total, avisos, rebars_creados, {
        u"sup": rebars_sup,
        u"inf": rebars_inf,
    }, lap_jobs_all
