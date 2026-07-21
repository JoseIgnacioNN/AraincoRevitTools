# -*- coding: utf-8 -*-
"""
Cabezal en encuentro L — capas equitativas en espesor detectado, barras en espesor seleccionado.

S.I.C. «encuentros de muros»; mínimo 2 capas × 2 barras.
Confinamiento por defecto sin; UI Tipo 1/2/3 (enc_fiber*):
  Tipo 1 estribo perimetral a la fibra; Tipo 2 + traba ⊥ host; Tipo 3 + long.
"""

from __future__ import print_function

import math

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import ElementId, UnitUtils, UnitTypeId, Wall, XYZ

try:
    from armado_muros_lineales import location_curve_wall, obtener_espesor_muro_mm_approx
except Exception:
    location_curve_wall = None
    obtener_espesor_muro_mm_approx = None

CABEZAL_COVER_MM = 25.0
CABEZAL_EXTREMO_INICIO = u"inicio"
CABEZAL_EXTREMO_FIN = u"fin"
CABEZAL_MAX_CAPAS = 6
CABEZAL_MIN_CAPAS = 1
CABEZAL_MIN_BARRAS_POR_CAPA = 2
CABEZAL_MAX_BARRAS_POR_CAPA = 4
CABEZAL_CONFINEMENT_NONE = u"none"

_CAB_MOD = None


def _cab():
    """Import diferido de ``armado_muros_cabezal`` (evita ciclo con el módulo principal)."""
    global _CAB_MOD
    if _CAB_MOD is None:
        try:
            import armado_muros_cabezal as cab
            _CAB_MOD = cab
        except Exception:
            _CAB_MOD = False
    return _CAB_MOD if _CAB_MOD is not False else None

CABEZAL_ENC_TIPO_LIBRE = None
CABEZAL_ENC_TIPO_L = u"L"

ENC_L_MIN_CAPAS = 2
ENC_L_MIN_BARS = 2

SIC_ENCOUNTER_ROWS = (
    {u"key": u"le200", u"e_max": 200, u"total": 4, u"diam": 12},
    {u"key": u"le300", u"e_max": 300, u"total": 4, u"diam": 16},
    {u"key": u"le350", u"e_max": 350, u"total": 4, u"diam": 18},
    {u"key": u"lt600", u"e_max": 599, u"total": 9, u"diam": 22},
    {u"key": u"lt800", u"e_max": 799, u"total": 9, u"diam": 25},
    {u"key": u"ge800", u"e_max": 99999, u"total": 12, u"diam": 25},
)


def cabezal_extremo_es_encuentro_l(ex_cfg):
    if not ex_cfg or not isinstance(ex_cfg, dict):
        return False
    return ex_cfg.get(u"encuentro_tipo") == CABEZAL_ENC_TIPO_L


def rejilla_desde_total_sic(total):
    try:
        t = int(total)
    except Exception:
        t = 4
    if t <= 4:
        return ENC_L_MIN_CAPAS, ENC_L_MIN_BARS
    if t <= 9:
        return 3, 3
    return 4, 3


def sic_encuentro_lookup(espesor_mm):
    try:
        e = int(round(float(espesor_mm)))
    except Exception:
        e = 300
    for row in SIC_ENCOUNTER_ROWS:
        if e <= int(row[u"e_max"]):
            return row
    return SIC_ENCOUNTER_ROWS[-1]


def espesor_mm_wall(wall):
    if wall is None:
        return 200.0
    if obtener_espesor_muro_mm_approx is not None:
        try:
            return float(obtener_espesor_muro_mm_approx(wall) or 200.0)
        except Exception:
            pass
    try:
        return float(
            UnitUtils.ConvertFromInternalUnits(
                float(wall.Width), UnitTypeId.Millimeters,
            )
        )
    except Exception:
        return 200.0


def cabezal_encuentro_l_offset_trans_mm(cover_mm, conf_bar_type, bar_type=None):
    c = float(cover_mm if cover_mm is not None else CABEZAL_COVER_MM)
    cab = _cab()
    conf_d = 16.0
    if cab is not None and conf_bar_type is not None:
        try:
            conf_d = float(cab._bar_diameter_mm(conf_bar_type))
        except Exception:
            conf_d = 16.0
    return c + conf_d + 10.0


def cabezal_encuentro_l_layer_positions_mm(
    espesor_det_mm, n_capas, cover_mm=None, conf_bar_type=None,
):
    """Posiciones de capas (mm desde cara exterior del muro detectado)."""
    e = max(float(espesor_det_mm or 200.0), 50.0)
    try:
        n = max(1, int(n_capas))
    except Exception:
        n = ENC_L_MIN_CAPAS
    off = cabezal_encuentro_l_offset_trans_mm(cover_mm, conf_bar_type)
    usable = max(0.0, e - 2.0 * off)
    if n <= 1:
        return {
            u"xs": [off + usable * 0.5],
            u"pitch_equitativo_mm": 0.0,
            u"offset_mm": off,
            u"usable_mm": usable,
        }
    pitch = usable / float(n - 1)
    xs = [off + (usable * float(i)) / float(n - 1) for i in range(n)]
    return {
        u"xs": xs,
        u"pitch_equitativo_mm": pitch,
        u"offset_mm": off,
        u"usable_mm": usable,
    }


def cabezal_encuentro_l_bar_positions_mm(
    espesor_sel_mm, n_bars, cover_mm=None, conf_bar_type=None,
):
    """Posiciones de barras (mm desde cara exterior del muro seleccionado)."""
    e = max(float(espesor_sel_mm or 200.0), 50.0)
    try:
        n = max(1, int(n_bars))
    except Exception:
        n = ENC_L_MIN_BARS
    off = cabezal_encuentro_l_offset_trans_mm(cover_mm, conf_bar_type)
    distrib = max(0.0, e - 2.0 * off)
    if n <= 1:
        ys = [off + distrib * 0.5]
    else:
        ys = [off + (distrib * float(k)) / float(n - 1) for k in range(n)]
    return {u"ys": ys, u"distrib_mm": distrib, u"offset_mm": off}


def cabezal_encuentro_l_refresh_pitch_in_cfg(ex_cfg, host_wall, neighbor_wall):
    """Actualiza ``pitch_equitativo_mm`` y espesores en config de encuentro L."""
    if not cabezal_extremo_es_encuentro_l(ex_cfg):
        return
    e_det = espesor_mm_wall(neighbor_wall)
    e_sel = espesor_mm_wall(host_wall)
    ex_cfg[u"espesor_detectado_mm"] = e_det
    ex_cfg[u"espesor_seleccionado_mm"] = e_sel
    try:
        n_capas = int(ex_cfg.get(u"n_capas", ENC_L_MIN_CAPAS))
    except Exception:
        n_capas = ENC_L_MIN_CAPAS
    conf_bt = None
    lp = cabezal_encuentro_l_layer_positions_mm(
        e_det, n_capas, CABEZAL_COVER_MM, conf_bt,
    )
    ex_cfg[u"pitch_equitativo_mm"] = float(lp.get(u"pitch_equitativo_mm") or 0.0)


def _line_dir_xy(line):
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        d = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0)
        ln = float(d.GetLength())
        if ln < 1e-12:
            return None, None
        return p0, d.Multiply(1.0 / ln)
    except Exception:
        return None, None


def _intersect_lines_xy(o1, d1, o2, d2):
    if o1 is None or d1 is None or o2 is None or d2 is None:
        return None
    try:
        det = float(d1.X * d2.Y - d1.Y * d2.X)
        if abs(det) < 1e-14:
            return None
        dx = float(o2.X - o1.X)
        dy = float(o2.Y - o1.Y)
        t = (dx * d2.Y - dy * d2.X) / det
        return XYZ(
            float(o1.X + t * d1.X),
            float(o1.Y + t * d1.Y),
            float(o1.Z),
        )
    except Exception:
        return None


def cabezal_encuentro_l_p_join(doc, host_wall, neighbor_wall, extremo_host):
    """Intersección de ejes (LocationCurve) en planta; Z del extremo del host."""
    cab = _cab()
    if cab is None:
        return None
    geom_h = cab._wall_longitudinal_at_extremo(host_wall, extremo_host)
    if geom_h is None:
        return None
    z_ref = float(geom_h[u"pt_extremo"].Z)
    lc_h = location_curve_wall(host_wall) if location_curve_wall else None
    lc_n = location_curve_wall(neighbor_wall) if location_curve_wall else None
    if lc_h is None or lc_n is None:
        return geom_h[u"pt_extremo"]
    o1, d1 = _line_dir_xy(lc_h)
    o2, d2 = _line_dir_xy(lc_n)
    pt = _intersect_lines_xy(o1, d1, o2, d2)
    if pt is None:
        return geom_h[u"pt_extremo"]
    return XYZ(float(pt.X), float(pt.Y), z_ref)


def _vecino_extremo_mas_cercano(neighbor_wall, pt_ref):
    lc = location_curve_wall(neighbor_wall) if location_curve_wall else None
    if lc is None or pt_ref is None:
        return CABEZAL_EXTREMO_INICIO
    try:
        p0 = lc.GetEndPoint(0)
        p1 = lc.GetEndPoint(1)
        if p0.DistanceTo(pt_ref) <= p1.DistanceTo(pt_ref):
            return CABEZAL_EXTREMO_INICIO
        return CABEZAL_EXTREMO_FIN
    except Exception:
        return CABEZAL_EXTREMO_INICIO


def _dot_dirs_wall(wall_a, wall_b):
    lc_a = location_curve_wall(wall_a) if location_curve_wall else None
    lc_b = location_curve_wall(wall_b) if location_curve_wall else None
    if lc_a is None or lc_b is None:
        return 0.0
    try:
        da = lc_a.GetEndPoint(1).Subtract(lc_a.GetEndPoint(0))
        db = lc_b.GetEndPoint(1).Subtract(lc_b.GetEndPoint(0))
        la = float(da.GetLength())
        lb = float(db.GetLength())
        if la < 1e-12 or lb < 1e-12:
            return 0.0
        da = da.Normalize()
        db = db.Normalize()
        return abs(float(da.DotProduct(db)))
    except Exception:
        return 0.0


def clasificar_encuentro_en_extremo(doc, host, neighbor, extremo):
    """
    ``libre`` | ``L`` | ``otro`` (T u otro; cabezal extremo no usa T en mitad de tramo).
    """
    if doc is None or host is None or neighbor is None:
        return u"libre"
    try:
        import armado_muros_vecinos_extremos as vec_mod
        if not vec_mod.vecino_en_extremo_muro(doc, host, extremo, neighbor):
            return u"libre"
    except Exception:
        return u"libre"
    if _dot_dirs_wall(host, neighbor) > 0.92:
        return u"otro"
    return CABEZAL_ENC_TIPO_L


def cabezal_extremo_config_encuentro_l(
    doc,
    host_wall,
    neighbor_wall,
    extremo,
    fallback_bar_type_id=None,
    fallback_conf_bar_type_id=None,
    sic_basis=u"seleccionado",
):
    """
    Config de extremo para encuentro L (S.I.C. encuentros, 2×2 mínimo).

    Confinamiento inicia en ``none``; el usuario puede elegir Tipo 1/2/3 (E1–E3 zona).
    """
    e_sel = espesor_mm_wall(host_wall)
    e_det = espesor_mm_wall(neighbor_wall)
    if sic_basis == u"detectado":
        e_ref = e_det
    elif sic_basis == u"max":
        e_ref = max(e_det, e_sel)
    else:
        e_ref = e_sel
    row = sic_encuentro_lookup(e_ref)
    n_capas, n_bars = rejilla_desde_total_sic(row[u"total"])
    n_capas = max(ENC_L_MIN_CAPAS, min(CABEZAL_MAX_CAPAS, n_capas))
    n_bars = max(
        ENC_L_MIN_BARS,
        min(CABEZAL_MAX_BARRAS_POR_CAPA, n_bars),
    )
    diam_mm = float(row[u"diam"])
    cab = _cab()
    if cab is None:
        return None
    fb_bt = None
    if fallback_bar_type_id not in (None, ElementId.InvalidElementId):
        fb_bt = cab._element_to_bar_type(doc, fallback_bar_type_id)
    bt = cab._bar_type_for_catalog_diameter_mm(doc, diam_mm, fb_bt)
    bt_id = bt.Id if bt is not None else None
    ex_cfg = cab.default_cabezal_extremo_config()
    ex_cfg[u"encuentro_tipo"] = CABEZAL_ENC_TIPO_L
    try:
        try:
            from bimtools_element_id import wall_id_int as _wid_int
            ex_cfg[u"vecino_wall_id"] = _wid_int(neighbor_wall)
        except Exception:
            ex_cfg[u"vecino_wall_id"] = int(neighbor_wall.Id.IntegerValue)
    except Exception:
        ex_cfg[u"vecino_wall_id"] = None
    ex_cfg[u"espesor_detectado_mm"] = e_det
    ex_cfg[u"espesor_seleccionado_mm"] = e_sel
    ex_cfg[u"sic_encuentro_key"] = row[u"key"]
    ex_cfg[u"sic_encuentro_total"] = int(row[u"total"])
    ex_cfg[u"sic_encuentro_diam_mm"] = diam_mm
    ex_cfg[u"sic_basis"] = sic_basis
    ex_cfg[u"n_capas"] = n_capas
    layers = []
    for _ in range(n_capas):
        layers.append(cab.default_cabezal_layer_config(n_bars, bt_id))
    ex_cfg[u"layers"] = layers
    if bt_id is not None:
        ex_cfg[u"bar_type_id"] = bt_id
    if fallback_conf_bar_type_id not in (None, ElementId.InvalidElementId):
        ex_cfg[u"conf_bar_type_id"] = fallback_conf_bar_type_id
    ex_cfg[u"confinement"] = cab.normalize_cabezal_confinement(
        {u"type": CABEZAL_CONFINEMENT_NONE}, n_capas,
    )
    cabezal_encuentro_l_refresh_pitch_in_cfg(ex_cfg, host_wall, neighbor_wall)
    return ex_cfg


def cabezal_encuentro_l_capa_line_endpoints(
    doc,
    wall,
    extremo,
    layer_index,
    bar_type,
    conf_bar_type,
    neighbor_wall,
    n_capas=None,
    cover_mm=None,
):
    """
    ``p_lo`` / ``p_hi`` / ``distrib_ft`` para una capa en encuentro L.

    Capas a lo largo del espesor detectado; barras en espesor del muro seleccionado.
    """
    cab = _cab()
    if cab is None:
        return None, None, None, u"Módulo cabezal no disponible."
    geom_h = cab._wall_longitudinal_at_extremo(wall, extremo)
    if geom_h is None:
        return None, None, None, u"Sin LocationCurve u orientación válida."
    if neighbor_wall is None:
        return None, None, None, u"Encuentro L: sin muro vecino."

    try:
        li = int(layer_index)
    except Exception:
        li = 0
    try:
        nc = max(ENC_L_MIN_CAPAS, int(n_capas or ENC_L_MIN_CAPAS))
    except Exception:
        nc = ENC_L_MIN_CAPAS

    e_det_mm = espesor_mm_wall(neighbor_wall)
    lp = cabezal_encuentro_l_layer_positions_mm(
        e_det_mm, nc, cover_mm, conf_bar_type,
    )
    xs = lp.get(u"xs") or []
    slot_i = li
    if extremo == CABEZAL_EXTREMO_INICIO and nc > 1:
        slot_i = max(0, nc - 1 - li)
    if slot_i < 0 or slot_i >= len(xs):
        return None, None, None, u"Índice de capa fuera de rango (encuentro L)."

    layer_mm = float(xs[slot_i])
    p_join = cabezal_encuentro_l_p_join(doc, wall, neighbor_wall, extremo)
    if p_join is None:
        p_join = geom_h[u"pt_extremo"]

    ex_vec = _vecino_extremo_mas_cercano(neighbor_wall, p_join)
    frame_n = cab._wall_extremo_frame(neighbor_wall, ex_vec)
    if frame_n is None:
        return None, None, None, u"Marco del muro detectado no válido."
    n_det = frame_n[u"inward"]

    e_det_ft = cab._mm_to_internal(e_det_mm)
    try:
        p_ext_det = p_join.Add(n_det.Negate().Multiply(e_det_ft * 0.5))
        plan_pt = p_ext_det.Add(n_det.Multiply(cab._mm_to_internal(layer_mm)))
    except Exception as ex_pt:
        return None, None, None, u"Offset capa L: {0}".format(str(ex_pt))

    espesor_ft = float(geom_h[u"espesor_ft"])
    normal_muro = geom_h[u"normal_muro"]
    offset_trans_mm = cabezal_encuentro_l_offset_trans_mm(
        cover_mm, conf_bar_type, bar_type,
    )
    offset_trans_ft = cab._mm_to_internal(offset_trans_mm)
    desplazamiento_lateral = espesor_ft * 0.5 - offset_trans_ft
    try:
        inicio = plan_pt.Add(normal_muro.Multiply(desplazamiento_lateral))
    except Exception as ex_pt2:
        return None, None, None, u"Offset barras L: {0}".format(str(ex_pt2))

    z_bot, z_top = cab._wall_z_bounds_ft(wall)
    p_lo = XYZ(float(inicio.X), float(inicio.Y), z_bot)
    p_hi = XYZ(float(inicio.X), float(inicio.Y), z_top)
    distrib_ft = max(espesor_ft - 2.0 * offset_trans_ft, 0.0)
    return p_lo, p_hi, distrib_ft, None


def cabezal_seccion_preview_layout_encuentro_l(
    espesor_det_mm,
    espesor_sel_mm,
    layers,
    cover_mm=None,
    draw_w_px=None,
    draw_h_px=None,
    confinement_type=None,
    confinement_stirrup_diam_mm=None,
    preview_fill_zone=True,
    extremo=None,
):
    """
    Preview 2D encuentro L: eje X = detectado (capas), eje Y = seleccionado (barras).

    Capas fuera → dentro: ``layers[0]`` (1ªC.) en ``xs[0]`` (cara exterior del
    detectado en el croquis). No invertir por ``extremo`` aquí: esa inversión
    desfasaba el canvas respecto al modelo en Revit (creación intacta).

    Con ``preview_fill_zone=True`` (preview WPF en zona intersección), las posiciones en mm
    se proyectan al rect completo sin padding px adicional — solo el recubrimiento en mm.
    """
    # ``extremo`` se acepta por compatibilidad de llamada; el layout no lo usa.
    _ = extremo
    layers = list(layers or [])
    e_det = max(float(espesor_det_mm or 200.0), 50.0)
    e_sel = max(float(espesor_sel_mm or 200.0), 50.0)
    n_capas = max(1, len(layers))
    lp = cabezal_encuentro_l_layer_positions_mm(
        e_det, n_capas, cover_mm, None,
    )
    dots = []
    layer_bounds = []
    dw = max(40.0, float(draw_w_px or 80.0))
    dh = max(28.0, float(draw_h_px or 50.0))
    if preview_fill_zone:
        pad_x = 0.0
        pad_y = 0.0
        inner_w = dw
        inner_h = dh
    else:
        pad_x = 14.0
        pad_y = 10.0
        inner_w = max(1.0, dw - 2.0 * pad_x)
        inner_h = max(1.0, dh - 2.0 * pad_y)

    xs = lp.get(u"xs") or []
    for i, ly in enumerate(layers):
        try:
            nb = int(ly.get(u"n_bars", ENC_L_MIN_BARS))
        except Exception:
            nb = ENC_L_MIN_BARS
        nb = max(ENC_L_MIN_BARS, min(CABEZAL_MAX_BARRAS_POR_CAPA, nb))
        bp = cabezal_encuentro_l_bar_positions_mm(
            e_sel, nb, cover_mm, None,
        )
        # 1ªC. → xs[0] (exterior del croquis); sin flip por extremo.
        slot_i = i
        layer_mm = float(xs[slot_i]) if 0 <= slot_i < len(xs) else 0.0
        if preview_fill_zone:
            fx = (layer_mm / e_det) * dw
        else:
            fx = pad_x + (layer_mm / e_det) * inner_w
        fx = max(0.0, min(dw, fx))
        fxs = []
        fys = []
        for bi, y_mm in enumerate(bp.get(u"ys") or []):
            if preview_fill_zone:
                fy = (float(y_mm) / e_sel) * dh
            else:
                fy = pad_y + (float(y_mm) / e_sel) * inner_h
            fy = max(0.0, min(dh, fy))
            fys.append(fy)
            fxs.append(fx)
            dots.append({
                u"layer_index": i,
                u"bar_index": int(bi),
                u"fx": fx / dw,
                u"fy": fy / dh,
            })
        if fxs and fys:
            layer_bounds.append({
                u"layer_index": i,
                u"fx0": min(fxs) / dw,
                u"fx1": max(fxs) / dw,
                u"fy0": min(fys) / dh,
                u"fy1": max(fys) / dh,
            })

    stirrup_rect = None
    stirrup_segments = None
    tie_preview = None
    tie_previews = []
    cross_tie_previews = []
    stirrup_layer_indices = []
    tie_layer_indices = []
    cab = _cab()
    if (
        cab is not None
        and confinement_type
        and confinement_type != CABEZAL_CONFINEMENT_NONE
        and cab.cabezal_confinement_scenario_applies(n_capas)
    ):
        stirrup_layer_indices, tie_layer_indices = cab.cabezal_confinement_layout_spec(
            n_capas, confinement_type,
        )
        pitch_frac = 0.12
        try:
            pitch_mm = float(lp.get(u"pitch_equitativo_mm") or 0.0)
            if pitch_mm > 1e-6 and e_det > 1e-6:
                pitch_frac = max(0.04, min(0.45, pitch_mm / e_det))
        except Exception:
            pass
        bar_diam_mm = 16.0
        if layers:
            try:
                bar_diam_mm = float(layers[0].get(u"bar_diam_mm") or bar_diam_mm)
            except Exception:
                pass
        conf_diam = float(
            confinement_stirrup_diam_mm
            if confinement_stirrup_diam_mm is not None
            else 10.0
        )
        if (
            cab.cabezal_confinement_has_perimeter_stirrup(confinement_type)
            and stirrup_layer_indices
        ):
            # Estribo anclado a caras de zona (recubrimiento uniforme), no a AABB barras.
            fiber_idx = list(range(n_capas))
            try:
                c_mm = float(
                    cover_mm if cover_mm is not None else CABEZAL_COVER_MM,
                )
            except Exception:
                c_mm = float(CABEZAL_COVER_MM)
            fx0 = max(0.0, min(0.49, c_mm / e_det))
            fy0 = max(0.0, min(0.49, c_mm / e_sel))
            stirrup_rect = {
                u"fx0": fx0,
                u"fx1": 1.0 - fx0,
                u"fy0": fy0,
                u"fy1": 1.0 - fy0,
            }
            stirrup_segments = cab.cabezal_stirrup_preview_segments(stirrup_rect)
            stirrup_layer_indices = list(fiber_idx)
        # Encuentro Tipo 1 = solo estribo; Tipo 2/3 = + trabas ⊥ (tie_layer_indices).
        for li in tie_layer_indices:
            ly_t = layers[li] if len(layers) > li else {}
            try:
                bar_d = float(ly_t.get(u"bar_diam_mm") or bar_diam_mm)
            except Exception:
                bar_d = bar_diam_mm
            tp = cab.cabezal_tie_preview_geometry(
                dots,
                layer_index=li,
                pitch_frac=pitch_frac,
                bar_diam_mm=bar_d,
                tie_diam_mm=conf_diam,
            )
            if tp:
                tie_previews.append(tp)
        if tie_previews:
            tie_preview = tie_previews[0]
        if (
            cab.cabezal_confinement_is_perimeter_cross(confinement_type)
            or cab.cabezal_confinement_is_enc_fiber_cross(confinement_type)
        ):
            n_bars_cross = ENC_L_MIN_BARS
            if layers:
                try:
                    n_bars_cross = int(
                        layers[0].get(u"n_bars", ENC_L_MIN_BARS),
                    )
                except Exception:
                    n_bars_cross = ENC_L_MIN_BARS
            n_bars_cross = max(
                ENC_L_MIN_BARS,
                min(CABEZAL_MAX_BARRAS_POR_CAPA, n_bars_cross),
            )
            for bi in cab.cabezal_confinement_cross_tie_bar_indices(n_bars_cross):
                ctp = cab.cabezal_cross_tie_preview_geometry(
                    dots,
                    bar_index=bi,
                    pitch_frac=pitch_frac,
                    bar_diam_mm=bar_diam_mm,
                    tie_diam_mm=conf_diam,
                )
                if ctp:
                    cross_tie_previews.append(ctp)

    return {
        u"dots": dots,
        u"layer_bounds": layer_bounds,
        u"stirrups": [],
        u"ties": [],
        u"stirrup_rect": stirrup_rect,
        u"stirrup_segments": stirrup_segments,
        u"stirrup_layer_indices": stirrup_layer_indices,
        u"tie_layer_indices": tie_layer_indices,
        u"tie_preview": tie_preview,
        u"tie_previews": tie_previews,
        u"cross_tie_previews": cross_tie_previews,
        u"pitch_equitativo_mm": lp.get(u"pitch_equitativo_mm"),
        u"encuentro_l": True,
    }


CABEZAL_ENC_TIPO_T = u"T"


def cabezal_encuentro_stub_len_mm(esp_mm):
    esp = max(float(esp_mm or 200.0), 50.0)
    return max(26.0, min(48.0, esp * 0.42))


def cabezal_encuentro_plan_l_metrics_mm(e_det_mm, e_sel_mm):
    """
    Métricas planta L (mm): zona intersección eDet×eSel + patas stub.
    Compartido por encuentros L y T en preview.
    """
    e_det = max(float(e_det_mm or 200.0), 50.0)
    e_sel = max(float(e_sel_mm or 200.0), 50.0)
    zone_w = e_det
    zone_h = e_sel
    leg_det = zone_h + cabezal_encuentro_stub_len_mm(zone_h)
    leg_sel = zone_w + cabezal_encuentro_stub_len_mm(zone_w)
    return {
        u"e_det": e_det,
        u"e_sel": e_sel,
        u"zone_w": zone_w,
        u"zone_h": zone_h,
        u"leg_det": leg_det,
        u"leg_sel": leg_sel,
        u"draw_w": leg_sel,
        u"draw_h": leg_det,
    }


def cabezal_encuentro_plan_l_polygon_local_px(draw_w_px, draw_h_px, e_det_mm, e_sel_mm, mirror=False):
    """
    Polígono planta L en coords locales (0,0)=esq. sup. izq. del canvas lógico.
    Retorna (points, zone_rect) con zone_rect = {x,y,w,h} en px locales.
    """
    m = cabezal_encuentro_plan_l_metrics_mm(e_det_mm, e_sel_mm)
    zone_w = m[u"zone_w"]
    zone_h = m[u"zone_h"]
    leg_det = m[u"leg_det"]
    leg_sel = m[u"leg_sel"]
    draw_w_mm = m[u"draw_w"]
    draw_h_mm = m[u"draw_h"]
    dw = max(40.0, float(draw_w_px or 80.0))
    dh = max(28.0, float(draw_h_px or 50.0))
    scale = min(dw / draw_w_mm, dh / draw_h_mm)
    off_x = (dw - draw_w_mm * scale) * 0.5
    off_y = (dh - draw_h_mm * scale) * 0.5

    def _mm_x(x_mm):
        xm = (draw_w_mm - float(x_mm)) if mirror else float(x_mm)
        return off_x + xm * scale

    def _mm_y(y_mm):
        return dh - off_y - float(y_mm) * scale

    poly_mm = (
        (0.0, 0.0),
        (0.0, leg_det),
        (zone_w, leg_det),
        (zone_w, zone_h),
        (leg_sel, zone_h),
        (leg_sel, 0.0),
    )
    points = [(_mm_x(x), _mm_y(y)) for x, y in poly_mm]
    zx0 = min(_mm_x(0.0), _mm_x(zone_w))
    zx1 = max(_mm_x(0.0), _mm_x(zone_w))
    zy0 = min(_mm_y(zone_h), _mm_y(0.0))
    zy1 = max(_mm_y(zone_h), _mm_y(0.0))
    zone_rect = {
        u"x": zx0,
        u"y": zy0,
        u"w": max(1.0, zx1 - zx0),
        u"h": max(1.0, zy1 - zy0),
    }
    join_y = _mm_y(zone_h)
    join_x0 = _mm_x(0.0)
    join_x1 = _mm_x(zone_w)
    p_join = {
        u"x": _mm_x(zone_w * 0.5),
        u"y": _mm_y(zone_h * 0.5),
    }
    return {
        u"points": points,
        u"zone_rect": zone_rect,
        u"join_line": (join_x0, join_y, join_x1, join_y),
        u"p_join": p_join,
    }
