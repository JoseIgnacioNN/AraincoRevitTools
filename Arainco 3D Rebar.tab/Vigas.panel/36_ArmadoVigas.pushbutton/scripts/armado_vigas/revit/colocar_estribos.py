# -*- coding: utf-8 -*-
"""
Colocación de estribos Ext/Cent y confinamiento E (1ª capa) en vigas del lote.

Reutiliza ``geometria_estribos_viga`` para el estribo perimetral y bucles internos
(E-pares, trabas) según ``domain/confinement.py`` y ``estConfin`` por viga.
"""

from __future__ import division

import clr

clr.AddReference("RevitAPI")

from System.Collections.Generic import List
from Autodesk.Revit.DB import Curve, CurveLoop, Line, LocationCurve, UnitUtils, UnitTypeId, XYZ

from armado_vigas.domain.confinement import ensure_beam_confinement, find_confin_def
from armado_vigas.domain.layers import ensure_beam_layers, first_layer_bar_count
from armado_vigas.revit.rebar_resources import resolve_bar_type_mm

_COVER_MM = 25.0


def _geometria_estribos_viga_module():
    """
    ``geometria_estribos_viga`` con ``reload`` — pyRevit/IronPython cachea módulos
    entre ejecuciones y puede quedar una versión sin ``_crear_rebar_traba_multizonas``.
    """
    import sys

    mod_name = u"geometria_estribos_viga"
    if mod_name in sys.modules:
        try:
            return reload(sys.modules[mod_name])
        except Exception:
            pass
    import geometria_estribos_viga as gev

    return gev


try:
    from geometria_viga_cara_superior_detalle import (
        _LAYOUT_ARRAY_SIDE_CLEARANCE_MM,
        _layout_max_spacing_array_length_ft,
        obtener_cara_superior_framing,
    )
except Exception:
    _LAYOUT_ARRAY_SIDE_CLEARANCE_MM = 50.0
    _layout_max_spacing_array_length_ft = None
    obtener_cara_superior_framing = None


def _mm_to_ft(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _element_id_int(el):
    try:
        return el.Id.IntegerValue
    except Exception:
        return None


def _curve_list_rect(corners):
    """``List[Curve]`` cerrado desde vértices consecutivos (≥3)."""
    if not corners or len(corners) < 3:
        return None
    lst = List[Curve]()
    n = len(corners)
    for i in range(n):
        p0 = corners[i]
        p1 = corners[(i + 1) % n]
        if p0 is None or p1 is None:
            return None
        try:
            if p0.DistanceTo(p1) < 1e-9:
                continue
            lst.Add(Line.CreateBound(p0, p1))
        except Exception:
            return None
    if lst.Count < 2:
        return None
    return lst


def _project_scalar(pt, origin, axis):
    try:
        return float((pt - origin).DotProduct(axis))
    except Exception:
        return 0.0


def _tangente_linea(line):
    if line is None:
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        return (p1.Subtract(p0)).Normalize()
    except Exception:
        return None


def _transverse_dir_from_face(n_face, line):
    t = _tangente_linea(line)
    if t is None or n_face is None:
        return None
    try:
        v = n_face.CrossProduct(t)
        if v.GetLength() < 1e-12:
            return None
        return v.Normalize()
    except Exception:
        return None


def _beam_layout_axes(host, line_work):
    """
    ``width_dir`` / ``depth_dir`` alineados con cara superior (mismo criterio que
    ``colocar_rebar._layout_v_dir``) para que trabas unan las barras reales.
    """
    width_dir = depth_dir = None
    t = _tangente_linea(line_work)
    if host is not None and obtener_cara_superior_framing is not None and t is not None:
        try:
            sup = obtener_cara_superior_framing(host)
            if sup is not None:
                n_sup = sup.FaceNormal.Normalize()
                width_dir = _transverse_dir_from_face(n_sup, line_work)
                if width_dir is not None:
                    depth_dir = width_dir.CrossProduct(t).Normalize()
                    if depth_dir.Z < 0:
                        depth_dir = depth_dir.Negate()
                        width_dir = t.CrossProduct(depth_dir).Normalize()
        except Exception:
            pass
    if width_dir is None or depth_dir is None:
        try:
            from armadura_vigas_capas import _beam_frame
        except Exception:
            _beam_frame = None
        if _beam_frame is not None:
            frame = _beam_frame(line_work)
            if frame is not None:
                _, width_dir, depth_dir, _, _, _ = frame
    if width_dir is None or depth_dir is None:
        z_up = XYZ.BasisZ
        width_dir = t.CrossProduct(z_up)
        if width_dir.GetLength() < 1e-9:
            width_dir = t.CrossProduct(XYZ.BasisX)
        width_dir = width_dir.Normalize()
        depth_dir = width_dir.CrossProduct(t).Normalize()
        if depth_dir.Z < 0:
            depth_dir = depth_dir.Negate()
            width_dir = t.CrossProduct(depth_dir).Normalize()
    return width_dir, depth_dir


def _section_interior_pt(sup_pts, inf_pts, section_origin):
    """Centroide 1ª capa en sección (referencia ganchos 135° hacia el interior)."""
    pts = list(sup_pts or []) + list(inf_pts or [])
    if not pts:
        return section_origin
    acc = XYZ.Zero
    for p in pts:
        acc = acc.Add(p)
    try:
        return acc.Multiply(1.0 / float(len(pts)))
    except Exception:
        return pts[0]


def _project_xyz_on_face(face, pt):
    if face is None or pt is None:
        return None
    try:
        res = face.Project(pt)
        if res is None:
            return None
        return res.XYZPoint
    except Exception:
        return None


def _bar_center_from_face(face, probe_pt, off_ft):
    """Centro de barra en cara: proyección + offset hacia el interior del hormigón."""
    pt = _project_xyz_on_face(face, probe_pt)
    if pt is None:
        return None
    try:
        n = face.FaceNormal.Normalize()
        return pt.Subtract(n.Multiply(float(off_ft)))
    except Exception:
        return pt


def _first_layer_bar_positions(document, host, beam, axis, width_dir, depth_dir, origin, line_work):
    """
    Posiciones 3D sup/inf de la 1ª capa — reparto transversal como ``colocar_rebar`` y
    profundidad sobre caras superior/inferior (no desde el centro analítico de sección).
    """
    try:
        from armadura_vigas_capas import _read_width_depth_ft
        from armado_vigas.revit.colocar_rebar import (
            _cara_framing,
            _layout_span_ft,
            _layout_v_dir,
            _width_center_at_guide,
        )
    except Exception:
        return None, None, None

    loc = getattr(host, "Location", None)
    if not isinstance(loc, LocationCurve):
        return None, None, None
    curve = loc.Curve
    if curve is None:
        return None, None, None

    w_ft, d_ft = _read_width_depth_ft(document, host, curve)
    if w_ft is None or d_ft is None or float(w_ft) <= 0 or float(d_ft) <= 0:
        return None, None, None

    n_layer = max(1, int(first_layer_bar_count(beam)))
    n_sup = n_layer
    n_inf = n_layer
    diam_sup_mm = float(beam.get("diamSup") or 16)
    diam_inf_mm = float(beam.get("diamInf") or 16)
    rex_mm = float(beam.get("estExtDiam") or beam.get("estCentDiam") or 10)
    bar_type_sup = resolve_bar_type_mm(document, diam_sup_mm)

    lw = line_work or curve
    sup_face = _cara_framing(host, False)
    inf_face = _cara_framing(host, True)
    n_face = None
    if sup_face is not None:
        try:
            n_face = sup_face.FaceNormal
        except Exception:
            n_face = None
    v_dir = _layout_v_dir(host, lw, n_face)
    if v_dir is None:
        v_dir = width_dir

    array_len = _layout_span_ft(
        document, host, bar_type_sup, n_layer, diam_sup_mm, rex_mm,
    )
    if array_len is None or float(array_len) < 1e-12:
        if _layout_max_spacing_array_length_ft is not None:
            try:
                array_len = _layout_max_spacing_array_length_ft(
                    document,
                    host,
                    bar_type_sup,
                    diametro_estribo_mm=rex_mm,
                )
            except Exception:
                array_len = None
        if array_len is None or float(array_len) < 1e-12:
            w_mm = float(w_ft) * 304.8
            span_mm = max(
                0.0,
                w_mm
                - float(_LAYOUT_ARRAY_SIDE_CLEARANCE_MM)
                - 2.0 * rex_mm
                - diam_sup_mm,
            )
            array_len = _mm_to_ft(span_mm)

    try:
        seg = max(1.0, float(lw.Length) * 0.01)
        sect_line = Line.CreateBound(
            origin.Subtract(axis.Multiply(seg)),
            origin.Add(axis.Multiply(seg)),
        )
    except Exception:
        sect_line = lw

    along_w = 0.0
    wc = _width_center_at_guide(host, sect_line, v_dir, False)
    if wc is not None:
        try:
            along_w = float((wc - origin).DotProduct(v_dir))
        except Exception:
            along_w = 0.0

    cover_ft = _mm_to_ft(_COVER_MM)
    rex_ft = _mm_to_ft(rex_mm)
    off_sup = cover_ft + rex_ft + _mm_to_ft(0.5 * diam_sup_mm)
    off_inf = cover_ft + rex_ft + _mm_to_ft(0.5 * diam_inf_mm)
    d_half = 0.5 * float(d_ft)
    half_span = 0.5 * float(array_len)

    def _probe_at_width_index(i, n):
        if n <= 1:
            return origin.Add(v_dir.Multiply(along_w))
        seed = origin.Add(v_dir.Multiply(along_w - half_span))
        step = float(array_len) / float(n - 1)
        return seed.Add(v_dir.Multiply(i * step))

    def _row_pts(n, face, off_ft, es_inf):
        pts = []
        for i in range(n):
            probe = _probe_at_width_index(i, n)
            pt = _bar_center_from_face(face, probe, off_ft)
            if pt is None:
                sign = -1.0 if es_inf else 1.0
                base = origin.Add(v_dir.Multiply(
                    along_w + (0.0 if n <= 1 else (-half_span + i * float(array_len) / float(n - 1)))
                ))
                pt = base.Add(depth_dir.Multiply(sign * (d_half - off_ft)))
            pts.append(pt)
        return pts

    sup_pts = _row_pts(n_sup, sup_face, off_sup, False)
    inf_pts = _row_pts(n_inf, inf_face, off_inf, True)
    return sup_pts, inf_pts, v_dir


def _prepare_beam_stirrup_geometry(document, host, beam, avisos):
    """
    Eje recortado entre tapas, normal y posiciones de barra en plano de sección tipo.
    """
    try:
        gev = _geometria_estribos_viga_module()
        _AXIS_CASI_VERTICAL_TOL = gev._AXIS_CASI_VERTICAL_TOL
        _ESTRIBO_PLANO_DESDE_INICIO_MM = gev._ESTRIBO_PLANO_DESDE_INICIO_MM
        _MIN_EDGE_FT = gev._MIN_EDGE_FT
        _curva_location_framing = gev._curva_location_framing
        _line_bound_desde_location_curve = gev._line_bound_desde_location_curve
        _linea_entre_tapas_extremas_viga = gev._linea_entre_tapas_extremas_viga
        _solido_principal = gev._solido_principal
        _mm_to_internal = gev._mm_to_internal
    except Exception as ex:
        avisos.append(u"geometria_estribos_viga: {0}".format(ex))
        return None

    crv = _curva_location_framing(host)
    if crv is None:
        avisos.append(u"Viga {0}: sin LocationCurve.".format(_element_id_int(host)))
        return None

    line_full = _line_bound_desde_location_curve(crv)
    if line_full is None:
        avisos.append(u"Viga {0}: eje inválido.".format(_element_id_int(host)))
        return None

    try:
        p0f = line_full.GetEndPoint(0)
        p1f = line_full.GetEndPoint(1)
        t = (p1f.Subtract(p0f)).Normalize()
    except Exception:
        avisos.append(u"Viga {0}: dirección de eje inválida.".format(_element_id_int(host)))
        return None

    if t is None or abs(float(t.Z)) > _AXIS_CASI_VERTICAL_TOL:
        avisos.append(
            u"Viga {0}: eje casi vertical; estribos omitidos.".format(_element_id_int(host))
        )
        return None

    solid = _solido_principal(host)
    line_work = _linea_entre_tapas_extremas_viga(line_full, solid) if solid else None
    if line_work is None:
        line_work = line_full

    try:
        p0 = line_work.GetEndPoint(0)
        p1 = line_work.GetEndPoint(1)
        t = (p1.Subtract(p0)).Normalize()
        Lw = float(line_work.Length)
    except Exception:
        avisos.append(u"Viga {0}: tramo de eje inválido.".format(_element_id_int(host)))
        return None

    offset_ft = _mm_to_internal(_ESTRIBO_PLANO_DESDE_INICIO_MM)
    if Lw < offset_ft + _MIN_EDGE_FT:
        offset_ft = max(_MIN_EDGE_FT * 4.0, 0.5 * Lw - _MIN_EDGE_FT)

    try:
        pt_seccion = p0.Add(t.Multiply(float(offset_ft)))
    except Exception:
        avisos.append(u"Viga {0}: punto de sección inválido.".format(_element_id_int(host)))
        return None

    width_dir, depth_dir = _beam_layout_axes(host, line_work)

    d_ft = None
    try:
        from armadura_vigas_capas import _read_width_depth_ft

        _, d_ft = _read_width_depth_ft(document, host, crv)
    except Exception:
        pass
    rex_mm = float(beam.get("estExtDiam") or beam.get("estCentDiam") or 10)

    sup_pts, inf_pts, layout_v_dir = _first_layer_bar_positions(
        document, host, beam, t, width_dir, depth_dir, pt_seccion, line_work
    )
    if not sup_pts or not inf_pts:
        avisos.append(
            u"Viga {0}: no se calcularon posiciones de barra (1ª capa).".format(
                _element_id_int(host)
            )
        )
        return None

    return {
        "line_work": line_work,
        "axis": t,
        "width_dir": width_dir,
        "depth_dir": depth_dir,
        "v_dir": layout_v_dir or width_dir,
        "sup_pts": sup_pts,
        "inf_pts": inf_pts,
        "section_origin": pt_seccion,
        "d_ft": d_ft,
        "rex_mm": rex_mm,
    }


def _tie_stem_offset_dir(bar_pt, interior_pt, axis, depth_dir, v_dir, bar_index, n_bars):
    """
    Dirección transversal (ancho) para desplazar el eje de la traba respecto al longitudinal.

    Barras a la izquierda del centro de reparto → ``-v_dir``; a la derecha → ``+v_dir``.
    En la barra central (n impar) el vector barra→centroide es nulo: convención ``+v_dir``
    (exterior derecho), no hacia la barra vecina izquierda.
    """
    try:
        if n_bars > 1 and v_dir is not None:
            mid = 0.5 * float(n_bars - 1)
            bi = float(bar_index)
            if bi < mid - 1e-9:
                return v_dir.Negate()
            if bi > mid + 1e-9:
                return v_dir
            try:
                d = bar_pt.Subtract(interior_pt)
                if axis is not None:
                    d = d.Subtract(axis.Multiply(float(d.DotProduct(axis))))
                if depth_dir is not None:
                    d = d.Subtract(depth_dir.Multiply(float(d.DotProduct(depth_dir))))
                ln = float(d.GetLength())
                if ln > 1e-9:
                    return d.Multiply(1.0 / ln)
            except Exception:
                pass
            return v_dir
    except Exception:
        pass
    try:
        d = interior_pt.Subtract(bar_pt)
        d = d.Subtract(axis.Multiply(float(d.DotProduct(axis))))
        d = d.Subtract(depth_dir.Multiply(float(d.DotProduct(depth_dir))))
        ln = float(d.GetLength())
        if ln > 1e-9:
            return d.Multiply(-1.0 / ln)
    except Exception:
        pass
    try:
        return v_dir.Negate() if v_dir is not None else XYZ.BasisY
    except Exception:
        return XYZ.BasisY


def _tie_stem_endpoints(
    bar_sup,
    bar_inf,
    interior_pt,
    axis,
    depth_dir,
    v_dir,
    off_sup_ft,
    off_inf_ft,
    bar_index,
    n_bars,
    rex_mm=None,
    tie_diam_mm=None,
    diam_sup_mm=None,
    diam_inf_mm=None,
):
    """
    Tramo recto B entre patas horizontales del estribo perimetral (desde barras reales).

    - Extremo vertical: tangente exterior de la barra + Ø estribo nominal
      (``bar_r + rex``), cara exterior de la pata horizontal del E perimetral.
      Antes solo se llegaba al eje del estribo (``bar_r + rex/2``) → ~5 mm corto
      con estribo Ø10.
    - Desplazamiento lateral: ``bar_r + tie_r`` (tangente exterior barra + radio traba;
      mismo criterio que cabezal ``_cabezal_tie_offset_mm``) para que el gancho 135°
      abraze el longitudinal del índice.

    Returns:
        ``(curves, p_top, p_bot, hook_ref)``
    """
    rex_mm = float(rex_mm or 10)

    dir_lat = _tie_stem_offset_dir(
        bar_sup, interior_pt, axis, depth_dir, v_dir, bar_index, n_bars,
    )
    off_s = float(off_sup_ft)
    off_i = float(off_inf_ft)
    hook_ref = _tie_bar_reference_pt(bar_sup, bar_inf)

    up_sec = depth_dir
    try:
        dv = bar_sup.Subtract(bar_inf)
        if axis is not None:
            dv = dv.Subtract(axis.Multiply(float(dv.DotProduct(axis))))
        ln = float(dv.GetLength())
        if ln > 1e-9:
            up_sec = dv.Multiply(1.0 / ln)
    except Exception:
        pass

    # Tangente exterior barra → cara exterior pata horizontal del estribo (eje + r_estribo).
    gap_sup_ft = _mm_to_ft(0.5 * float(diam_sup_mm or 16) + rex_mm)
    gap_inf_ft = _mm_to_ft(0.5 * float(diam_inf_mm or 16) + rex_mm)

    try:
        p_top = bar_sup.Add(up_sec.Multiply(gap_sup_ft))
        p_bot = bar_inf.Subtract(up_sec.Multiply(gap_inf_ft))
    except Exception:
        p_top = bar_sup
        p_bot = bar_inf

    p_top = p_top.Add(dir_lat.Multiply(off_s))
    p_bot = p_bot.Add(dir_lat.Multiply(off_i))

    curves = List[Curve]()
    try:
        if p_top.DistanceTo(p_bot) > 1e-9:
            curves.Add(Line.CreateBound(p_top, p_bot))
    except Exception:
        pass

    return curves, p_top, p_bot, hook_ref


def _tie_bar_reference_pt(bar_sup, bar_inf):
    """Centro de la barra índice (referencia de ganchos hacia el longitudinal)."""
    try:
        return bar_sup.Add(bar_inf).Multiply(0.5)
    except Exception:
        return bar_sup


def _tie_lateral_offset_ft(long_diam_mm, tie_diam_mm):
    """Eje barra → eje traba: tangente exterior (bar_r) + radio traba (tie_r)."""
    try:
        return _mm_to_ft(0.5 * float(long_diam_mm) + 0.5 * float(tie_diam_mm))
    except Exception:
        return _mm_to_ft(12.0)


def _long_bar_face_offset_ft(cover_mm, rex_mm, long_diam_mm):
    """Cara hormigón → eje longitudinal (mismo criterio que ``_first_layer_bar_positions``)."""
    return _mm_to_ft(float(cover_mm) + float(rex_mm) + 0.5 * float(long_diam_mm))


def _stirrup_cover_center_ft(cover_mm, stirrup_diam_mm):
    """Cara hormigón → eje estribo (recubrimiento + radio; cara inferior / laterales)."""
    return _mm_to_ft(float(cover_mm) + 0.5 * float(stirrup_diam_mm))


def _stirrup_cover_face_ft(cover_mm):
    """Recubrimiento nominal cara hormigón → eje estribo (sin Ø/2)."""
    return _mm_to_ft(float(cover_mm))


def _stirrup_bar_clearance_ft(long_diam_mm, stirrup_diam_mm):
    """Tangente exterior longitudinal → eje estribo en lados interiores."""
    return _mm_to_ft(0.5 * float(long_diam_mm) + 0.5 * float(stirrup_diam_mm))


def _corners_e_pair(
    sup_pts,
    inf_pts,
    i0,
    i1,
    width_dir,
    depth_dir,
    cover_mm,
    rex_mm,
    diam_sup_mm,
    diam_inf_mm,
    stirrup_diam_mm,
    n_bars,
):
    if i0 >= len(sup_pts) or i1 >= len(sup_pts):
        return None
    if i0 >= len(inf_pts) or i1 >= len(inf_pts):
        return None
    pts = [sup_pts[i0], sup_pts[i1], inf_pts[i1], inf_pts[i0]]
    origin = pts[0]
    w_vals = [_project_scalar(p, origin, width_dir) for p in pts]
    d_vals = [_project_scalar(p, origin, depth_dir) for p in pts]
    w_min, w_max = min(w_vals), max(w_vals)
    d_min, d_max = min(d_vals), max(d_vals)

    cover_st_ft = _stirrup_cover_center_ft(cover_mm, stirrup_diam_mm)
    cover_top_ft = _stirrup_cover_face_ft(cover_mm)
    off_sup_ft = _long_bar_face_offset_ft(cover_mm, rex_mm, diam_sup_mm)
    off_inf_ft = _long_bar_face_offset_ft(cover_mm, rex_mm, diam_inf_mm)
    off_w_ft = off_sup_ft
    inner_w_ft = _stirrup_bar_clearance_ft(diam_sup_mm, stirrup_diam_mm)

    # Cara superior: recubrimiento nominal (25 mm); no sumar Ø_E/2 al eje (evita +4 mm con Ø8).
    # Cara inferior: eje a cover + Ø_E/2 (validado en Revit).
    d_top = d_max + off_sup_ft - cover_top_ft
    d_bot = d_min - off_inf_ft + cover_st_ft

    i_lo, i_hi = min(int(i0), int(i1)), max(int(i0), int(i1))
    w_left_bar = _project_scalar(sup_pts[i_lo], origin, width_dir)
    w_right_bar = _project_scalar(sup_pts[i_hi], origin, width_dir)

    if i_lo <= 0:
        w0 = w_left_bar - off_w_ft + cover_st_ft
        w0 = min(w0, w_min - inner_w_ft)
    else:
        w0 = w_min - inner_w_ft

    if i_hi >= int(n_bars) - 1:
        w1 = w_right_bar + off_w_ft - cover_st_ft
        w1 = max(w1, w_max + inner_w_ft)
    else:
        w1 = w_max + inner_w_ft

    if w1 - w0 < 1e-9 or d_top - d_bot < 1e-9:
        return None

    d0, d1 = d_bot, d_top

    return [
        origin + width_dir.Multiply(w0) + depth_dir.Multiply(d0),
        origin + width_dir.Multiply(w1) + depth_dir.Multiply(d0),
        origin + width_dir.Multiply(w1) + depth_dir.Multiply(d1),
        origin + width_dir.Multiply(w0) + depth_dir.Multiply(d1),
    ]


def _tie_interior_reference(sup_pts, inf_pts, section_origin):
    """Punto interior de la sección para orientar ganchos 135° de trabas."""
    if section_origin is not None:
        return section_origin
    pts = list(sup_pts or []) + list(inf_pts or [])
    if not pts:
        return None
    acc = XYZ.Zero
    for p in pts:
        acc = acc.Add(p)
    try:
        return acc.Multiply(1.0 / float(len(pts)))
    except Exception:
        return pts[0]


def _curve_loop_from_corners(corners):
    curves = _curve_list_rect(corners)
    if curves is None:
        return None
    try:
        loop = CurveLoop()
        for i in range(int(curves.Count)):
            loop.Append(curves[i])
        return loop
    except Exception:
        return None


def _append_conf_tag_jobs(conf_jobs, rebars_slice, host, beam, conf, zone_meta_slice=None, **fields):
    """Registra metadatos de etiquetado para rebars recién añadidos."""
    if conf_jobs is None or not rebars_slice:
        return
    try:
        from armado_vigas.domain.layers import first_layer_bar_count

        n_capas = int(first_layer_bar_count(beam))
    except Exception:
        n_capas = 0
    try:
        host_id = _element_id_int(host)
    except Exception:
        host_id = None
    for i, rb in enumerate(rebars_slice):
        if rb is None:
            continue
        job = {
            u"rebar_id": rb.Id,
            u"host_id": host_id,
            u"host": host,
            u"beam": beam,
            u"conf": conf,
            u"n_capas": n_capas,
        }
        if zone_meta_slice and i < len(zone_meta_slice):
            zm = zone_meta_slice[i] or {}
            if zm.get(u"zone_index") is not None:
                job[u"stirrup_zone_index"] = zm.get(u"zone_index")
            if zm.get(u"zone_kind"):
                job[u"stirrup_zone_kind"] = zm.get(u"zone_kind")
        job.update(fields)
        conf_jobs.append(job)


def _place_internal_loops(
    document,
    host,
    geo,
    beam,
    conf,
    bt_ext,
    bt_cent,
    bt_tie,
    sp_ext,
    sp_cent,
    avisos,
    rebars_out,
    view,
    conf_jobs=None,
):
    try:
        gev = _geometria_estribos_viga_module()
        _crear_rebar_estribos_multizonas = gev._crear_rebar_estribos_multizonas
        _crear_rebar_traba_multizonas = getattr(
            gev, u"_crear_rebar_traba_multizonas", None,
        )
        if _crear_rebar_traba_multizonas is None:
            avisos.append(
                u"Estribos internos: geometria_estribos_viga sin soporte de trabas "
                u"(recargue pyRevit)."
            )
            return 0
    except Exception as ex:
        avisos.append(u"Estribos internos: {0}".format(ex))
        return 0

    line_work = geo.get("line_work")
    axis = geo.get("axis")
    width_dir = geo.get("width_dir")
    depth_dir = geo.get("depth_dir")
    sup_pts = geo.get("sup_pts") or []
    inf_pts = geo.get("inf_pts") or []
    rex_mm = float(geo.get("rex_mm") or beam.get("estExtDiam") or beam.get("estCentDiam") or 10)
    stirrup_diam_mm = float(beam.get("estCentDiam") or beam.get("estExtDiam") or 8)
    diam_sup_mm = float(beam.get("diamSup") or 16)
    diam_inf_mm = float(beam.get("diamInf") or 16)
    n_bars = len(sup_pts)
    n_total = 0

    if bt_tie is None:
        avisos.append(
            u"Viga {0}: sin RebarBarType para traba.".format(
                beam.get("id") or _element_id_int(host)
            )
        )
        bt_tie = bt_ext or bt_cent

    def _emit_loop(corners, pair=None):
        loop_draw = _curve_loop_from_corners(corners)
        if loop_draw is None:
            return 0
        n_before = len(rebars_out)
        zone_meta = []
        n_created = _crear_rebar_estribos_multizonas(
            document,
            host,
            line_work,
            loop_draw,
            axis,
            bt_ext,
            bt_cent,
            sp_ext,
            sp_cent,
            avisos,
            rebars_creados=rebars_out,
            rebar_zone_meta_out=zone_meta,
            view=view,
        )
        _append_conf_tag_jobs(
            conf_jobs,
            rebars_out[n_before:],
            host,
            beam,
            conf,
            zone_meta_slice=zone_meta,
            job_kind=u"stirrup",
            stirrup_kind=u"pair",
            pair=list(pair) if pair is not None else None,
        )
        return n_created

    for pair in conf.get("pairs") or []:
        if not pair or len(pair) < 2:
            continue
        corners = _corners_e_pair(
            sup_pts,
            inf_pts,
            int(pair[0]),
            int(pair[1]),
            width_dir,
            depth_dir,
            _COVER_MM,
            rex_mm,
            diam_sup_mm,
            diam_inf_mm,
            stirrup_diam_mm,
            n_bars,
        )
        if corners is None:
            avisos.append(
                u"Viga {0}: par E({1}) inválido para N barras.".format(
                    beam.get("id") or _element_id_int(host),
                    u"–".join(str(i) for i in pair),
                )
            )
            continue
        n_total += _emit_loop(corners, pair=pair)

    interior_pt = _section_interior_pt(
        sup_pts, inf_pts, geo.get("section_origin"),
    )
    if interior_pt is None:
        avisos.append(
            u"Viga {0}: sin referencia interior para trabas.".format(
                beam.get("id") or _element_id_int(host)
            )
        )
    else:
        tie_diam_mm = float(beam.get("estCentDiam") or beam.get("estExtDiam") or 8)
        diam_sup_mm = float(beam.get("diamSup") or 16)
        diam_inf_mm = float(beam.get("diamInf") or 16)
        off_sup_ft = _tie_lateral_offset_ft(diam_sup_mm, tie_diam_mm)
        off_inf_ft = _tie_lateral_offset_ft(diam_inf_mm, tie_diam_mm)
        n_bars = len(sup_pts)
        v_dir = geo.get("v_dir") or width_dir
        for idx in conf.get("ties") or []:
            ii = int(idx)
            if ii >= len(sup_pts) or ii >= len(inf_pts):
                avisos.append(
                    u"Viga {0}: traba [{1}] fuera de rango.".format(
                        beam.get("id") or _element_id_int(host), idx
                    )
                )
                continue
            bar_sup = sup_pts[ii]
            bar_inf = inf_pts[ii]
            n_before = len(rebars_out)
            tie_curves, p_top, p_bot, hook_ref = _tie_stem_endpoints(
                bar_sup,
                bar_inf,
                interior_pt,
                axis,
                depth_dir,
                v_dir,
                off_sup_ft,
                off_inf_ft,
                ii,
                n_bars,
                rex_mm=geo.get("rex_mm"),
                tie_diam_mm=tie_diam_mm,
                diam_sup_mm=diam_sup_mm,
                diam_inf_mm=diam_inf_mm,
            )
            n_created = _crear_rebar_traba_multizonas(
                document,
                host,
                line_work,
                p_top,
                p_bot,
                axis,
                hook_ref,
                width_dir,
                depth_dir,
                bt_tie,
                bt_ext,
                bt_cent,
                sp_ext,
                sp_cent,
                avisos,
                rebars_out,
                view,
                curves_list=tie_curves,
            )
            if n_created <= 0:
                avisos.append(
                    u"Viga {0}: traba [{1}] no generada (revisar avisos / hook 135°).".format(
                        beam.get("id") or _element_id_int(host), idx
                    )
                )
            else:
                _append_conf_tag_jobs(
                    conf_jobs,
                    rebars_out[n_before:],
                    host,
                    beam,
                    conf,
                    job_kind=u"tie",
                    tie_index=ii,
                )
            n_total += n_created

    return n_total


def colocar_estribos_confinamiento(document, session, view=None):
    """
    Coloca estribos perimetrales (Ext/Cent) y confinamiento E por viga del lote.

    Returns:
        ``(n_posiciones_rebar, avisos, rebars_creados, conf_tag_jobs)``
    """
    if not session.framing_elements:
        return 0, [u"Sin vigas en el lote."], [], []

    try:
        gev = _geometria_estribos_viga_module()
        crear_model_lines_preview_estribo_viga = gev.crear_model_lines_preview_estribo_viga
    except Exception as ex:
        return 0, [u"No se cargó geometria_estribos_viga: {0}".format(ex)], [], []

    domain_by_id = session.domain_beams_by_element_id or {}
    n_total = 0
    avisos = []
    rebars = []
    conf_tag_jobs = []

    for el in session.framing_elements or []:
        if el is None:
            continue
        eid = _element_id_int(el)
        beam = domain_by_id.get(eid)
        if not beam:
            avisos.append(u"Viga Revit {0}: sin datos de dominio; estribos omitidos.".format(eid))
            continue

        ensure_beam_layers(beam)
        ensure_beam_confinement(beam)
        conf = find_confin_def(beam)
        bt_ext = resolve_bar_type_mm(document, beam.get("estExtDiam") or 10)
        bt_cent = resolve_bar_type_mm(document, beam.get("estCentDiam") or 8)
        tie_diam = beam.get("estCentDiam") or beam.get("estExtDiam") or 8
        bt_tie = resolve_bar_type_mm(document, tie_diam)
        sp_ext = int(beam.get("estExtSpacing") or 125)
        sp_cent = int(beam.get("estCentSpacing") or 200)
        reb_local = []

        pairs = conf.get("pairs") or []
        ties = conf.get("ties") or []
        # Estribos multizona Ext/Cent del vano: perimetral E o vano sin pares/trabas.
        colocar_estribos_zona = bool(conf.get("perimetral")) or (not pairs and not ties)
        if colocar_estribos_zona:
            n_reb_before = len(reb_local)
            zone_meta = []
            _, _, n_rb, av = crear_model_lines_preview_estribo_viga(
                document,
                el,
                cover_mm=_COVER_MM,
                crear_rebar_estribos=True,
                rebar_bar_type_ext=bt_ext,
                rebar_bar_type_cent=bt_cent,
                spacing_ext_mm=sp_ext,
                spacing_cent_mm=sp_cent,
                out_rebars_creados=reb_local,
                out_rebar_zone_meta=zone_meta,
                view=view,
            )
            n_total += int(n_rb or 0)
            avisos.extend(av or [])
            _append_conf_tag_jobs(
                conf_tag_jobs,
                reb_local[n_reb_before:],
                el,
                beam,
                conf,
                zone_meta_slice=zone_meta,
                job_kind=u"stirrup",
                stirrup_kind=u"zona_ext_cent",
            )

        if pairs or ties:
            geo = _prepare_beam_stirrup_geometry(document, el, beam, avisos)
            if geo is not None:
                n_total += _place_internal_loops(
                    document,
                    el,
                    geo,
                    beam,
                    conf,
                    bt_ext,
                    bt_cent,
                    bt_tie,
                    sp_ext,
                    sp_cent,
                    avisos,
                    reb_local,
                    view,
                    conf_jobs=conf_tag_jobs,
                )

        rebars.extend(reb_local)

    return n_total, avisos, rebars, conf_tag_jobs
