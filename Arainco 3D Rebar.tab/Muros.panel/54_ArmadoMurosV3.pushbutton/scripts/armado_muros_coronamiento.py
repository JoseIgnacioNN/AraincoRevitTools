# -*- coding: utf-8 -*-
"""
Coronamiento de muros — superior, inferior (fundación / pie), voladizo (reentrada stack).

Superior: host muro tope, recubrimiento 25 mm + Ø/2 desde cara superior.
Inferior fundación: host fundación, cara inferior + 50 mm + Ø/2, patas L hacia arriba.
Inferior pie: muro sin apilamiento ni fundación debajo, base + 25 mm + Ø/2, patas L arriba.
Voladizo superior: base del muro superior más largo + estirón, pata L hacia arriba.
Voladizo inferior: tope del muro inferior más largo bajo apilamiento más corto,
                  pata L hacia abajo en extremo libre.

Largo del tramo horizontal (U):
- **Punta** (sin vecino L en el extremo): de recub. en P0 a recub. en P1 (comportamiento histórico).
- **Encuentro L** (escenario adicional): anclaje en el nudo de ``LocationCurve`` del muro
  modelado; estiramiento ``e_host/2`` hacia la cara del extremo y después recub. + Ø/2 hacia
  el interior (como punta, pero partiendo del eje en encuentro). El otro extremo sigue punta.

Tras crear, se etiquetan con ``EST_A_STRUCTURAL REBAR TAG_WALL_HORIZONTAL``.
"""

from __future__ import print_function

import os
import sys

import clr

clr.AddReference("RevitAPI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Category,
    Curve,
    ElementId,
    Line,
    Transaction,
    UnitUtils,
    UnitTypeId,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarStyle,
)

try:
    from armado_muros_lineales import (
        location_curve_wall,
        obtener_espesor_muro_mm_approx,
        ordenar_muros_por_base_asc,
    )
except Exception:
    location_curve_wall = None
    obtener_espesor_muro_mm_approx = None
    ordenar_muros_por_base_asc = None

try:
    import armado_muros_cabezal as cabezal
except Exception:
    cabezal = None

try:
    import rebar_extender_l_ganchos_135_rps as l135
except Exception:
    l135 = None

try:
    import armado_muros_cabezal_tags as _cab_tags
except Exception:
    _cab_tags = None

try:
    from armado_muros_rebar_params import (
        activar_armadura_arainco,
        stamp_coronamiento_rebar,
    )
except Exception:
    activar_armadura_arainco = None
    stamp_coronamiento_rebar = None

CORONAMIENTO_COVER_SUPERIOR_MM = 25.0
CORONAMIENTO_COVER_INFERIOR_MM = 50.0
CORONAMIENTO_COVER_MM = CORONAMIENTO_COVER_SUPERIOR_MM
CORONAMIENTO_PATA_RESTA_MM = 50.0
CORONAMIENTO_PATA_MIN_MM = 100.0
CORONAMIENTO_TAG_EXTREMO_SUP = u"cor_sup"
CORONAMIENTO_TAG_EXTREMO_INF = u"cor_inf"
CORONAMIENTO_TAG_EXTREMO_PIE = u"cor_pie"
CORONAMIENTO_TAG_EXTREMO_VOL = u"cor_vol"
CORONAMIENTO_VOLADIZO_U_TOL_MM = 2.0
CORONAMIENTO_VOLADIZO_MIN_HORIZ_MM = 80.0
CORONAMIENTO_VOLADIZO_ROLE_SUP = u"sup"
CORONAMIENTO_VOLADIZO_ROLE_INF = u"inf"
CORONAMIENTO_EXTREMO_INICIO = u"inicio"
CORONAMIENTO_EXTREMO_FIN = u"fin"
CORONAMIENTO_MODO_PUNTA = u"punta"
CORONAMIENTO_MODO_ENCUENTRO_L = u"encuentro_l"
CORONAMIENTO_MIN_TRAMO_HORIZ_MM = 80.0


def _ensure_pushbutton_path():
    try:
        import bootstrap_paths
        return bootstrap_paths.pin_local_scripts_first()
    except Exception:
        d = os.path.dirname(os.path.abspath(__file__))
        if d and d not in sys.path:
            sys.path.insert(0, d)
        return d


_ensure_pushbutton_path()

try:
    from geometria_fundacion_cara_inferior import (
        evaluar_caras_paralelas_curva_mas_cercana,
        obtener_marco_coordenadas_cara_inferior,
        vector_reverso_cara_paralela_mas_cercana_a_barra,
    )
except Exception:
    evaluar_caras_paralelas_curva_mas_cercana = None
    obtener_marco_coordenadas_cara_inferior = None
    vector_reverso_cara_paralela_mas_cercana_a_barra = None

try:
    from rebar_fundacion_cara_inferior import (
        REBAR_SHAPE_NOMBRE_DEFECTO,
        aplicar_layout_fixed_number_rebar,
        crear_rebar_polilinea_recta_sin_ganchos,
        crear_rebar_polilinea_u_malla_inf_sup_curve_loop,
        crear_rebar_u_shape_desde_eje_rebar_shape_nombrado,
    )
except Exception:
    REBAR_SHAPE_NOMBRE_DEFECTO = u"03"
    aplicar_layout_fixed_number_rebar = None
    crear_rebar_polilinea_recta_sin_ganchos = None
    crear_rebar_polilinea_u_malla_inf_sup_curve_loop = None
    crear_rebar_u_shape_desde_eje_rebar_shape_nombrado = None


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _wall_z_bounds_ft(wall):
    if cabezal is not None:
        try:
            return cabezal._wall_z_bounds_ft(wall)
        except Exception:
            pass
    try:
        bb = wall.get_BoundingBox(None)
        if bb is not None:
            return float(bb.Min.Z), float(bb.Max.Z)
    except Exception:
        pass
    return 0.0, 3.0


def _element_z_bounds_ft(elem):
    try:
        bb = elem.get_BoundingBox(None)
        if bb is not None:
            return float(bb.Min.Z), float(bb.Max.Z)
    except Exception:
        pass
    return None, None


def _espesor_muro_mm(wall):
    if wall is None:
        return 200.0
    if obtener_espesor_muro_mm_approx is not None:
        try:
            th = obtener_espesor_muro_mm_approx(wall)
            if th is not None and float(th) > 0.1:
                return float(th)
        except Exception:
            pass
    if cabezal is not None:
        try:
            return float(cabezal._wall_thickness_mm_for_fusion(wall))
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


def coronamiento_tipico_por_espesor_mm(e_mm):
    """Armadura típica en coronamiento (S.I.C.) según espesor nominal."""
    try:
        e = int(round(float(e_mm)))
    except Exception:
        return None, None
    if e <= 200:
        return 2, 12
    if e <= 300:
        return 2, 16
    if e <= 400:
        return 3, 18
    if e <= 600:
        return 4, 22
    return 4, 22


def default_coronamiento_config(espesor_mm=None, wall=None):
    """Config UI/creación: tipico S.I.C. del espesor (tope) + flags de tipos."""
    e = espesor_mm
    if e is None and wall is not None:
        e = _espesor_muro_mm(wall)
    n_bars, diam_mm = coronamiento_tipico_por_espesor_mm(e if e is not None else 300)
    if n_bars is None:
        n_bars, diam_mm = 2, 16
    return {
        u"activo": True,
        u"n_bars": int(n_bars),
        u"diam_mm": int(diam_mm),
        u"crear_superior": False,
        u"crear_inferior": False,
        u"crear_voladizo": False,
    }


def normalize_coronamiento_config(cfg, espesor_mm=None, wall=None):
    """Fusiona overrides del usuario sobre tipico; no altera geometría ni pipeline."""
    out = default_coronamiento_config(espesor_mm=espesor_mm, wall=wall)
    if not cfg or not isinstance(cfg, dict):
        return out
    for key in (u"activo", u"crear_superior", u"crear_inferior", u"crear_voladizo"):
        if key in cfg:
            out[key] = bool(cfg.get(key))
    if cfg.get(u"n_bars") is not None:
        try:
            out[u"n_bars"] = max(2, min(4, int(cfg.get(u"n_bars"))))
        except Exception:
            pass
    if cfg.get(u"diam_mm") is not None:
        try:
            out[u"diam_mm"] = int(cfg.get(u"diam_mm"))
        except Exception:
            pass
    return out


def _resolve_n_bars_diam_mm(cfg, e_mm):
    """Nº/Ø desde cfg por muro; respaldo tipico S.I.C. por espesor."""
    cfg = cfg or {}
    n_bars = cfg.get(u"n_bars")
    diam_mm = cfg.get(u"diam_mm")
    if n_bars is None or diam_mm is None:
        tip_n, tip_d = coronamiento_tipico_por_espesor_mm(e_mm)
        if n_bars is None:
            n_bars = tip_n
        if diam_mm is None:
            diam_mm = tip_d
    if n_bars is None or diam_mm is None:
        return None, None
    return max(2, min(4, int(n_bars))), int(diam_mm)


def _normalize_coronamiento_por_muro_id(mapping):
    out = {}
    if not mapping:
        return out
    for k, v in list(mapping.items()):
        try:
            kid = int(k)
        except Exception:
            continue
        if isinstance(v, dict):
            out[kid] = dict(v)
    return out


def _cfg_for_wall(wall, coronamiento_por_muro_id, global_config):
    wid = _wall_id_int(wall)
    # Mapa por muro (V3): solo entradas explícitas; sin entrada → sin barras.
    if coronamiento_por_muro_id is not None:
        if wid is not None and wid in coronamiento_por_muro_id:
            return normalize_coronamiento_config(
                coronamiento_por_muro_id.get(wid), wall=wall,
            )
        out = default_coronamiento_config(wall=wall)
        out[u"activo"] = False
        return out
    if global_config:
        return normalize_coronamiento_config(global_config, wall=wall)
    return default_coronamiento_config(wall=wall)


def _n_bars_cfg_or_tipico(e_mm, res=None):
    if res is not None and res.get(u"cfg_n_bars") is not None:
        try:
            return max(2, min(4, int(res.get(u"cfg_n_bars"))))
        except Exception:
            pass
    n_bars, _ = coronamiento_tipico_por_espesor_mm(e_mm)
    return n_bars


def _diam_cfg_or_tipico(e_mm, res=None):
    if res is not None and res.get(u"cfg_diam_mm") is not None:
        try:
            return int(res.get(u"cfg_diam_mm"))
        except Exception:
            pass
    _, diam_mm = coronamiento_tipico_por_espesor_mm(e_mm)
    return diam_mm


def muro_tope_stack_global(walls_ordered):
    """Muro sin otro más arriba en la selección (tope global del stack)."""
    walls = [w for w in (walls_ordered or []) if w is not None]
    if not walls:
        return None
    if len(walls) == 1:
        return walls[0]
    best = None
    best_z = None
    tol = 1e-6
    for w in walls:
        try:
            _, z1 = _wall_z_bounds_ft(w)
        except Exception:
            continue
        if best is None or float(z1) > float(best_z) + tol:
            best = w
            best_z = float(z1)
    return best


def _stack_contact_z_tolerance_ft():
    if cabezal is not None:
        try:
            return float(cabezal._troceo_u_tolerance_ft())
        except Exception:
            pass
    return _mm_to_internal(10.0)


def _muro_tiene_apilamiento_inferior(wall, walls_ord, tol_z_ft=None):
    """
    True si otro muro de la selección apoya este muro por debajo (contacto en Z base).
    """
    if wall is None:
        return False
    z_bot, _ = _wall_z_bounds_ft(wall)
    if tol_z_ft is None:
        tol_z_ft = _stack_contact_z_tolerance_ft()
    try:
        wall_id = wall.Id.IntegerValue
    except Exception:
        wall_id = None
    for w_other in walls_ord or []:
        if w_other is None:
            continue
        if wall_id is not None:
            try:
                if int(w_other.Id.IntegerValue) == int(wall_id):
                    continue
            except Exception:
                pass
        _, z_top = _wall_z_bounds_ft(w_other)
        if abs(float(z_top) - float(z_bot)) <= float(tol_z_ft):
            return True
    return False


def _z_bar_inferior_pie_ft(wall, bar_type, fallback_diam_mm=None):
    """Pie del muro (sin apilamiento inferior) + recub. 25 mm + Ø/2."""
    z_bot, _ = _wall_z_bounds_ft(wall)
    axis_offset_ft = _mm_to_internal(_cover_axis_offset_mm(
        CORONAMIENTO_COVER_SUPERIOR_MM,
        bar_type,
        fallback_diam_mm=fallback_diam_mm,
    ))
    return float(z_bot) + axis_offset_ft


def _bar_type_for_diameter_mm(doc, diam_mm, fallback=None):
    if cabezal is not None:
        try:
            return cabezal._bar_type_for_diameter_mm(doc, diam_mm, fallback)
        except Exception:
            pass
    return fallback


def _bar_diameter_mm(bar_type, fallback_diam_mm=None):
    if bar_type is None:
        return float(fallback_diam_mm or 12.0)
    if cabezal is not None:
        try:
            return float(cabezal._bar_diameter_mm(bar_type))
        except Exception:
            pass
    if l135 is not None:
        try:
            return float(l135._nominal_diameter_mm(bar_type))
        except Exception:
            pass
    try:
        return float(
            UnitUtils.ConvertFromInternalUnits(
                float(bar_type.BarModelDiameter), UnitTypeId.Millimeters,
            )
        )
    except Exception:
        return float(fallback_diam_mm or 12.0)


def _cover_axis_offset_mm(cover_mm, bar_type, fallback_diam_mm=None):
    """Recubrimiento nominal a superficie de barra → offset del eje = c + Ø/2."""
    c = float(cover_mm if cover_mm is not None else CORONAMIENTO_COVER_SUPERIOR_MM)
    r = float(_bar_diameter_mm(bar_type, fallback_diam_mm)) * 0.5
    return c + r


def _perp_horizontal_xy(tu_xy):
    """Unitario horizontal ⟂ a ``tu_xy`` (proyección XY)."""
    if tu_xy is None:
        return XYZ.BasisX
    try:
        v = XYZ(float(tu_xy.X), float(tu_xy.Y), 0.0)
        if float(v.GetLength()) < 1e-12:
            return XYZ.BasisX
        u = v.Normalize()
        return XYZ(-float(u.Y), float(u.X), 0.0).Normalize()
    except Exception:
        return XYZ.BasisX


def _norm_propagacion_coronamiento_inf_fund(linea_horizontal, wall_normal):
    """
    Normal de ``CreateFromCurves*`` / propagación del conjunto en planta (⟂ al eje
    horizontal), alineada con el espesor del muro — mismo criterio que zapata de muro.
    """
    if linea_horizontal is None:
        return None, None
    try:
        tu = linea_horizontal.GetEndPoint(1).Subtract(linea_horizontal.GetEndPoint(0))
        tu_xy = XYZ(float(tu.X), float(tu.Y), 0.0)
        if float(tu_xy.GetLength()) < 1e-12:
            return None, None
        tu_xy = tu_xy.Normalize()
        n = _perp_horizontal_xy(tu_xy)
        if wall_normal is not None:
            try:
                w_xy = XYZ(float(wall_normal.X), float(wall_normal.Y), 0.0)
                if float(w_xy.GetLength()) > 1e-12:
                    w_xy = w_xy.Normalize()
                    if float(n.DotProduct(w_xy)) < 0.0:
                        n = n.Negate()
            except Exception:
                pass
        return n, n
    except Exception:
        return None, None


def _rebar_cantidad_posiciones(rebar):
    if rebar is None:
        return 0
    try:
        return int(rebar.Quantity)
    except Exception:
        try:
            return int(rebar.NumberOfBarPositions)
        except Exception:
            return 0


def _aplicar_layout_coronamiento_inf_fundacion(rebar, doc, n_bars, distrib_ft):
    """
    Fixed Number: primera barra en cara interior (+normal) y conjunto hacia la cara opuesta.
    Mismo criterio que coronamiento con host muro (``barsOnNormalSide=False`` primero).
    """
    if n_bars <= 1 or float(distrib_ft or 0.0) <= 1e-9:
        return True
    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception:
        return False
    target = int(n_bars)
    for b_side in (False, True):
        try:
            acc.SetLayoutAsFixedNumber(
                target, float(distrib_ft), b_side, True, True,
            )
            if doc is not None:
                try:
                    doc.Regenerate()
                except Exception:
                    pass
            if _rebar_cantidad_posiciones(rebar) == target:
                return True
        except Exception:
            continue
    return False


def _fundaciones_unidas_muro(doc, wall):
    if cabezal is not None:
        try:
            return cabezal._fundaciones_unidas_muro(doc, wall)
        except Exception:
            pass
    try:
        import armado_muros_verticales_embed_colision as _emb
        return _emb._fundaciones_estructurales_unidas_muro(doc, wall)
    except Exception:
        return []


def _fundacion_principal_muro(doc, wall):
    """Fundación estructural unida (host preferido: mayor volumen aprox. por bbox)."""
    funds = _fundaciones_unidas_muro(doc, wall) or []
    if not funds:
        return None
    best = None
    best_vol = None
    for fund in funds:
        if fund is None:
            continue
        z0, z1 = _element_z_bounds_ft(fund)
        if z0 is None or z1 is None:
            continue
        vol = abs(float(z1) - float(z0))
        if best is None or vol > best_vol:
            best = fund
            best_vol = vol
    return best or funds[0]


def _altura_fundacion_mm(foundation):
    if cabezal is not None:
        try:
            h = cabezal._altura_bbox_elemento_mm(foundation)
            if h is not None:
                return float(h)
        except Exception:
            pass
    z0, z1 = _element_z_bounds_ft(foundation)
    if z0 is None or z1 is None:
        return None
    try:
        return float(
            UnitUtils.ConvertFromInternalUnits(
                abs(float(z1) - float(z0)), UnitTypeId.Millimeters,
            )
        )
    except Exception:
        return None


def _nominal_diameter_bar_type_mm(bar_type, fallback_diam_mm=None):
    if bar_type is not None and l135 is not None:
        try:
            d_mm = float(l135._nominal_diameter_mm(bar_type))
            if d_mm > 0.1:
                return d_mm
        except Exception:
            pass
    if bar_type is not None and cabezal is not None:
        try:
            d_mm = cabezal._bar_diameter_mm(bar_type)
            if d_mm is not None and float(d_mm) > 0.1:
                return float(d_mm)
        except Exception:
            pass
    if fallback_diam_mm is not None:
        try:
            d_mm = float(fallback_diam_mm)
            if d_mm > 0.1:
                return d_mm
        except Exception:
            pass
    return None


def _largo_pata_l_eje_sketch_mm(bar_type, fallback_diam_mm=None):
    """
    Largo de eje para patas L sup./inf. (tabla BIMTools − Ø/2).

    Revit modela la pata desde el eje; sin compensación la geometría queda ~tabla + Ø/2.
    Reutiliza el helper de cabezal cuando está disponible.
    """
    d_mm = _nominal_diameter_bar_type_mm(bar_type, fallback_diam_mm)
    if d_mm is None:
        return None
    if cabezal is not None:
        try:
            leje = cabezal._pata_l_eje_sketch_mm_desde_diametro(d_mm)
            if leje is not None and float(leje) > 0.1:
                return float(leje)
        except Exception:
            pass
    try:
        from bimtools_rebar_hook_lengths import (
            hook_length_mm_from_nominal_diameter_mm,
            pata_eje_curve_loop_mm_desde_tabla_mm,
        )
        tabla_mm = hook_length_mm_from_nominal_diameter_mm(d_mm)
        if tabla_mm is not None and float(tabla_mm) > 0.1:
            leje = pata_eje_curve_loop_mm_desde_tabla_mm(tabla_mm, d_mm)
            if leje is not None and float(leje) > 0.1:
                return float(leje)
    except Exception:
        pass
    return None


def _largo_pata_l_mm(doc, host_geom, bar_type):
    host = host_geom
    if l135 is not None and doc is not None and host is not None:
        try:
            lp = float(
                l135.largo_pata_mm_desde_espesor_host(
                    doc, host, resta_mm=CORONAMIENTO_PATA_RESTA_MM,
                )
            )
            if lp > 0.1:
                try:
                    d_mm = float(l135._nominal_diameter_mm(bar_type))
                    lp = max(lp, 12.0 * d_mm, CORONAMIENTO_PATA_MIN_MM)
                except Exception:
                    lp = max(lp, CORONAMIENTO_PATA_MIN_MM)
                return lp
        except Exception:
            pass
    if host_geom is not None:
        try:
            e_mm = _espesor_muro_mm(host_geom)
        except Exception:
            e_mm = 200.0
    else:
        e_mm = 200.0
    return max(CORONAMIENTO_PATA_MIN_MM, float(e_mm) - CORONAMIENTO_PATA_RESTA_MM)


def _largo_pata_l_sup_inf_sketch_mm(doc, host_geom, bar_type, fallback_diam_mm=None):
    """Pata L coronamiento sup./inf.: tabla por Ø con compensación de eje; respaldo espesor."""
    leg_mm = _largo_pata_l_eje_sketch_mm(bar_type, fallback_diam_mm=fallback_diam_mm)
    if leg_mm is not None and float(leg_mm) > 0.1:
        return float(leg_mm)
    return _largo_pata_l_mm(doc, host_geom, bar_type)


def _wall_location_frame(wall):
    """Eje longitudinal, normal exterior y espesor desde LocationCurve del muro."""
    lc = location_curve_wall(wall) if location_curve_wall else None
    if lc is None:
        return None
    try:
        p0 = lc.GetEndPoint(0)
        p1 = lc.GetEndPoint(1)
        t_raw = p1.Subtract(p0)
        tl = float(t_raw.GetLength())
        if tl < 1e-12:
            return None
        t_hat = t_raw.Normalize()
    except Exception:
        return None
    try:
        normal = wall.Orientation.Normalize()
    except Exception:
        normal = None
    if normal is None or float(normal.GetLength()) < 1e-12:
        try:
            normal = t_hat.CrossProduct(XYZ.BasisZ).Normalize()
        except Exception:
            normal = XYZ.BasisY
    try:
        espesor_ft = float(wall.Width)
    except Exception:
        espesor_ft = _mm_to_internal(200.0)
    if espesor_ft < 1e-9:
        espesor_ft = _mm_to_internal(200.0)
    return {
        u"p0": p0,
        u"p1": p1,
        u"t_hat": t_hat,
        u"normal": normal,
        u"espesor_ft": espesor_ft,
        u"z_loc": 0.5 * (float(p0.Z) + float(p1.Z)),
    }


def _z_bar_superior_ft(wall, bar_type, cover_mm=None, fallback_diam_mm=None):
    axis_offset_ft = _mm_to_internal(_cover_axis_offset_mm(
        cover_mm if cover_mm is not None else CORONAMIENTO_COVER_SUPERIOR_MM,
        bar_type,
        fallback_diam_mm=fallback_diam_mm,
    ))
    _, z_top = _wall_z_bounds_ft(wall)
    return float(z_top) - axis_offset_ft


def _z_cara_inferior_fundacion_ft(foundation):
    """Cota Z de la cara inferior de la fundación (marco UVN o bbox)."""
    if foundation is None:
        return None
    if obtener_marco_coordenadas_cara_inferior is not None:
        try:
            marco = obtener_marco_coordenadas_cara_inferior(foundation)
            if marco is not None and marco[0] is not None:
                return float(marco[0].Z)
        except Exception:
            pass
    z_fund_bot, _ = _element_z_bounds_ft(foundation)
    return z_fund_bot


def _z_bar_inferior_fundacion_ft(wall, foundation, bar_type, fallback_diam_mm=None):
    """
    Cota del tramo horizontal inferior: cara **inferior** fundación + 50 mm + Ø/2.

    La planta (XY) y offsets longitudinales/transversales salen de LocationCurve
    del muro; solo se traslada verticalmente hasta el recubrimiento desde el
    fondo de la fundación.
    """
    frame = _wall_location_frame(wall)
    if frame is None:
        return None, u"Sin LocationCurve del muro."
    z_fund_bot = _z_cara_inferior_fundacion_ft(foundation)
    if z_fund_bot is None:
        return None, u"Sin geometría Z de fundación."
    axis_offset_ft = _mm_to_internal(_cover_axis_offset_mm(
        CORONAMIENTO_COVER_INFERIOR_MM,
        bar_type,
        fallback_diam_mm=fallback_diam_mm,
    ))
    z_bar = float(z_fund_bot) + axis_offset_ft
    return z_bar, None


def _element_id_int(elem_or_id):
    try:
        from bimtools_element_id import element_id_to_int
        if elem_or_id is not None and hasattr(elem_or_id, "Id"):
            v = element_id_to_int(elem_or_id.Id)
            if v is not None:
                return v
        return element_id_to_int(elem_or_id)
    except Exception:
        pass
    if elem_or_id is None:
        return None
    try:
        return int(elem_or_id.Value)
    except Exception:
        pass
    try:
        return int(elem_or_id.IntegerValue)
    except Exception:
        return None


def _rebar_host_id_int(rebar):
    if rebar is None:
        return None
    try:
        return _element_id_int(rebar.GetHostId())
    except Exception:
        return None


def _delete_rebar_safe(doc, rebar):
    if doc is None or rebar is None:
        return
    try:
        doc.Delete(rebar.Id)
    except Exception:
        pass


def _coronamiento_extremo_modo_y_vecino(doc, wall, extremo):
    """
    Modo por extremo: ``punta`` (sin encuentro L) o ``encuentro_l`` + muro vecino.
    """
    if doc is None or wall is None or extremo not in (
        CORONAMIENTO_EXTREMO_INICIO,
        CORONAMIENTO_EXTREMO_FIN,
    ):
        return CORONAMIENTO_MODO_PUNTA, None
    try:
        import armado_muros_vecinos_extremos as vec_ext
    except Exception:
        return CORONAMIENTO_MODO_PUNTA, None
    try:
        neighbor = vec_ext.vecino_principal_encuentro_l(doc, wall, extremo)
    except Exception:
        neighbor = None
    if neighbor is None:
        return CORONAMIENTO_MODO_PUNTA, None
    return CORONAMIENTO_MODO_ENCUENTRO_L, neighbor


def _coronamiento_p_join_encuentro_l(doc, wall, neighbor, extremo):
    if doc is None or wall is None or neighbor is None:
        return None
    try:
        import armado_muros_cabezal_encuentro_l as enc_l
        return enc_l.cabezal_encuentro_l_p_join(doc, wall, neighbor, extremo)
    except Exception:
        return None


def _coronamiento_espesor_vecino_mm(neighbor):
    if neighbor is None:
        return 200.0
    try:
        import armado_muros_cabezal_encuentro_l as enc_l
        return float(enc_l.espesor_mm_wall(neighbor))
    except Exception:
        return float(_espesor_muro_mm(neighbor))


def _coronamiento_into_wall_en_extremo(wall, extremo, t_hat_fallback):
    """Unitario hacia el interior del muro en ``extremo`` (misma convención que cabezal)."""
    if cabezal is not None:
        try:
            geom = cabezal._wall_longitudinal_at_extremo(wall, extremo)
            if geom is not None:
                v = geom.get(u"vector_longitudinal")
                if v is not None and float(v.GetLength()) > 1e-12:
                    return v.Normalize()
        except Exception:
            pass
    if extremo == CORONAMIENTO_EXTREMO_FIN:
        try:
            return t_hat_fallback.Negate().Normalize()
        except Exception:
            pass
    try:
        return t_hat_fallback.Normalize()
    except Exception:
        return XYZ.BasisX


def _coronamiento_pt_horizontal_en_extremo(
    doc,
    wall,
    extremo,
    z_bar,
    axis_offset_ft,
    t_hat,
    normal,
    lateral,
    p0z,
    p1z,
    modo,
    neighbor,
):
    """
    Punto del tramo horizontal en un extremo (cara exterior + recub. longitudinal).

    Punta: P0/P1 + recub. hacia el interior del muro.
    Encuentro L: ``P_join`` (Location Line) + ``e_host/2`` hacia la cara del extremo y
    después el mismo recub. longitudinal que punta, hacia el interior.
    """
    if modo == CORONAMIENTO_MODO_ENCUENTRO_L and neighbor is not None:
        p_join = _coronamiento_p_join_encuentro_l(doc, wall, neighbor, extremo)
        if p_join is not None:
            try:
                into_wall = _coronamiento_into_wall_en_extremo(wall, extremo, t_hat)
                e_host_mm = float(_espesor_muro_mm(wall))
                half_host_ft = _mm_to_internal(e_host_mm * 0.5)
                cover_in_ft = float(axis_offset_ft)
                pt_xy = XYZ(float(p_join.X), float(p_join.Y), float(z_bar))
                hacia_cara = into_wall.Negate()
                return (
                    pt_xy
                    + hacia_cara.Multiply(half_host_ft)
                    + into_wall.Multiply(cover_in_ft)
                    + normal.Multiply(lateral)
                )
            except Exception:
                pass
    if extremo == CORONAMIENTO_EXTREMO_INICIO:
        return p0z + t_hat.Multiply(axis_offset_ft) + normal.Multiply(lateral)
    return p1z - t_hat.Multiply(axis_offset_ft) + normal.Multiply(lateral)


def _coronamiento_extremos_resumen(doc, wall):
    """Texto inicio/fin → punta | encuentro_l (log / mensajes)."""
    parts = []
    for ex, lbl in (
        (CORONAMIENTO_EXTREMO_INICIO, u"ini"),
        (CORONAMIENTO_EXTREMO_FIN, u"fin"),
    ):
        modo, vec = _coronamiento_extremo_modo_y_vecino(doc, wall, ex)
        if modo == CORONAMIENTO_MODO_ENCUENTRO_L and vec is not None:
            try:
                e_h = int(round(_espesor_muro_mm(wall)))
            except Exception:
                e_h = u"?"
            try:
                e_v = int(round(_coronamiento_espesor_vecino_mm(vec)))
            except Exception:
                e_v = u"?"
            parts.append(
                u"{0}=enc.L e_mod={1} e_vec={2}".format(lbl, e_h, e_v),
            )
        else:
            parts.append(u"{0}={1}".format(lbl, modo or CORONAMIENTO_MODO_PUNTA))
    return u", ".join(parts)


def _coronamiento_chain_curves(
    doc,
    wall,
    z_bar_ft,
    leg_len_ft,
    bar_type,
    cover_mm=None,
    fallback_diam_mm=None,
    legs_up=False,
):
    """Polilínea U: patas verticales + tramo horizontal a ``z_bar_ft``."""
    frame = _wall_location_frame(wall)
    if frame is None:
        return None, None, None, u"Sin LocationCurve válida."
    axis_offset_mm = _cover_axis_offset_mm(
        cover_mm, bar_type, fallback_diam_mm=fallback_diam_mm,
    )
    axis_offset_ft = _mm_to_internal(axis_offset_mm)
    z_bar = float(z_bar_ft)

    p0 = frame[u"p0"]
    p1 = frame[u"p1"]
    t_hat = frame[u"t_hat"]
    normal = frame[u"normal"]
    espesor_ft = float(frame[u"espesor_ft"])
    dist_eje_cara = espesor_ft * 0.5
    offset_trans_ft = axis_offset_ft
    lateral = dist_eje_cara - offset_trans_ft

    modo_ini, vec_ini = _coronamiento_extremo_modo_y_vecino(
        doc, wall, CORONAMIENTO_EXTREMO_INICIO,
    )
    modo_fin, vec_fin = _coronamiento_extremo_modo_y_vecino(
        doc, wall, CORONAMIENTO_EXTREMO_FIN,
    )

    try:
        p0z = XYZ(float(p0.X), float(p0.Y), z_bar)
        p1z = XYZ(float(p1.X), float(p1.Y), z_bar)
        pt_start = _coronamiento_pt_horizontal_en_extremo(
            doc,
            wall,
            CORONAMIENTO_EXTREMO_INICIO,
            z_bar,
            axis_offset_ft,
            t_hat,
            normal,
            lateral,
            p0z,
            p1z,
            modo_ini,
            vec_ini,
        )
        pt_end = _coronamiento_pt_horizontal_en_extremo(
            doc,
            wall,
            CORONAMIENTO_EXTREMO_FIN,
            z_bar,
            axis_offset_ft,
            t_hat,
            normal,
            lateral,
            p0z,
            p1z,
            modo_fin,
            vec_fin,
        )
        horiz_len_ft = float(pt_start.DistanceTo(pt_end))
        min_ft = _mm_to_internal(CORONAMIENTO_MIN_TRAMO_HORIZ_MM)
        if horiz_len_ft < min_ft:
            try:
                horiz_mm = float(
                    UnitUtils.ConvertFromInternalUnits(
                        horiz_len_ft, UnitTypeId.Millimeters,
                    )
                )
            except Exception:
                horiz_mm = horiz_len_ft
            return None, None, None, (
                u"Tramo horizontal coronamiento demasiado corto ({0:.0f} mm); "
                u"extremos: {1}.".format(
                    horiz_mm,
                    _coronamiento_extremos_resumen(doc, wall),
                )
            )
    except Exception as ex_pt:
        return None, None, None, u"Geometría coronamiento: {0}".format(ex_pt)

    leg = max(float(leg_len_ft), 1e-6)
    try:
        if legs_up:
            up = XYZ.BasisZ
            pt_start_leg = pt_start + up.Multiply(leg)
            pt_end_leg = pt_end + up.Multiply(leg)
            c1 = Line.CreateBound(pt_start_leg, pt_start)
            c2 = Line.CreateBound(pt_start, pt_end)
            c3 = Line.CreateBound(pt_end, pt_end_leg)
        else:
            down = XYZ.BasisZ.Negate()
            pt_start_leg = pt_start + down.Multiply(leg)
            pt_end_leg = pt_end + down.Multiply(leg)
            c1 = Line.CreateBound(pt_start_leg, pt_start)
            c2 = Line.CreateBound(pt_start, pt_end)
            c3 = Line.CreateBound(pt_end, pt_end_leg)
    except Exception as ex_ln:
        return None, None, None, u"Line.CreateBound: {0}".format(ex_ln)

    distrib_ft = max(espesor_ft - 2.0 * offset_trans_ft, 0.0)
    return [c1, c2, c3], normal, distrib_ft, None


def _create_coronamiento_rebar(
    doc,
    geom_wall,
    host,
    n_bars,
    bar_type,
    z_bar_ft,
    cover_mm=None,
    fallback_diam_mm=None,
    legs_up=False,
    leg_host_geom=None,
):
    if doc is None or geom_wall is None or host is None or bar_type is None:
        return None, 0, u"Doc, muro, host o tipo de barra no válido."
    try:
        n_bars = int(n_bars)
    except Exception:
        n_bars = 2
    n_bars = max(2, min(4, n_bars))

    leg_mm = _largo_pata_l_sup_inf_sketch_mm(
        doc, leg_host_geom or geom_wall, bar_type, fallback_diam_mm=fallback_diam_mm,
    )
    leg_ft = _mm_to_internal(leg_mm)
    curves, normal, distrib_ft, err = _coronamiento_chain_curves(
        doc,
        geom_wall,
        z_bar_ft,
        leg_ft,
        bar_type,
        cover_mm=cover_mm,
        fallback_diam_mm=fallback_diam_mm,
        legs_up=legs_up,
    )
    if err:
        return None, 0, err

    try:
        curves_list = List[Curve]()
        for c in curves:
            curves_list.Add(c)
    except Exception as ex_cl:
        return None, 0, u"IList[Curve]: {0}".format(ex_cl)

    try:
        rebar = Rebar.CreateFromCurves(
            doc,
            RebarStyle.Standard,
            bar_type,
            None,
            None,
            host,
            normal,
            curves_list,
            RebarHookOrientation.Left,
            RebarHookOrientation.Left,
            True,
            True,
        )
    except Exception as ex_cf:
        try:
            return None, 0, unicode(ex_cf)
        except Exception:
            return None, 0, str(ex_cf)

    if rebar is None:
        return None, 0, u"CreateFromCurves devolvió None."

    if n_bars > 1 and float(distrib_ft or 0.0) > 1e-9:
        try:
            accessor = rebar.GetShapeDrivenAccessor()
            accessor.SetLayoutAsFixedNumber(
                int(n_bars), float(distrib_ft), False, True, True,
            )
        except Exception as ex_lay:
            try:
                return None, 0, u"SetLayoutAsFixedNumber: {0}".format(unicode(ex_lay))
            except Exception:
                return None, 0, u"SetLayoutAsFixedNumber: {0}".format(str(ex_lay))

    if stamp_coronamiento_rebar is not None:
        try:
            stamp_coronamiento_rebar(rebar)
        except Exception:
            pass
    elif activar_armadura_arainco is not None:
        try:
            activar_armadura_arainco(rebar)
        except Exception:
            pass

    return rebar, int(n_bars), None


def _voladizo_u_tolerance_ft():
    if cabezal is not None:
        try:
            return float(cabezal._troceo_u_tolerance_ft())
        except Exception:
            pass
    return _mm_to_internal(CORONAMIENTO_VOLADIZO_U_TOL_MM)


def _empotramiento_voladizo_mm(d_mm):
    if cabezal is not None:
        try:
            emb = cabezal._empotramiento_tabla_mm(float(d_mm))
            if emb is not None and float(emb) > 0.1:
                return float(emb)
        except Exception:
            pass
    try:
        return max(400.0, 40.0 * float(d_mm))
    except Exception:
        return 400.0


def _compute_stacked_layout(walls):
    try:
        from armado_muros_lineales import compute_stacked_wall_layout
        return compute_stacked_wall_layout(walls)
    except Exception:
        return None


def _xyz_on_wall_at_u(wall, item, u_val, z_bar, normal, lateral_ft):
    """Punto en el eje del muro a cota ``z_bar`` para coordenada ``u`` del stack layout."""
    frame = _wall_location_frame(wall)
    if frame is None or item is None:
        return None
    p0 = frame[u"p0"]
    p1 = frame[u"p1"]
    try:
        u_start = float(item.get(u"u_start", 0.0))
        u_end = float(item.get(u"u_end", u_start))
        u0 = float(item.get(u"u0", u_start))
        u1 = float(item.get(u"u1", u_end))
    except Exception:
        return None
    if abs(u0 - u_start) <= abs(u1 - u_start):
        u_at_p0 = u_start
        u_at_p1 = u_end
    else:
        u_at_p0 = u_end
        u_at_p1 = u_start
    span_u = float(u_at_p1) - float(u_at_p0)
    if abs(span_u) < 1e-12:
        frac = 0.5
    else:
        frac = (float(u_val) - float(u_at_p0)) / span_u
        frac = max(0.0, min(1.0, frac))
    try:
        p0z = XYZ(float(p0.X), float(p0.Y), float(z_bar))
        p1z = XYZ(float(p1.X), float(p1.Y), float(z_bar))
        delta = p1z.Subtract(p0z)
        pt = p0z + delta.Multiply(frac)
        if normal is not None and lateral_ft is not None:
            pt = pt + normal.Multiply(float(lateral_ft))
        return pt
    except Exception:
        return None


def _detect_voladizos_stack(walls_ord, stacked_layout):
    """
    Voladizos en reentrada entre muros adyacentes del stack.

    - sup: muro superior más largo → host superior, cota base + pata L arriba.
    - inf: muro inferior más largo (apilamiento encima más corto) → host inferior,
           cota tope + pata L abajo.
    """
    out = []
    walls = list(walls_ord or [])
    if len(walls) < 2 or stacked_layout is None:
        return out
    items = stacked_layout.get(u"items") or []
    tol_u = _voladizo_u_tolerance_ft()
    for i in range(1, len(walls)):
        if i >= len(items) or (i - 1) >= len(items):
            continue
        w_lo = walls[i - 1]
        w_hi = walls[i]
        it_lo = items[i - 1]
        it_hi = items[i]
        if w_lo is None or w_hi is None:
            continue
        try:
            u_end_lo = float(it_lo.get(u"u_end", 0.0))
            u_end_hi = float(it_hi.get(u"u_end", 0.0))
            u_start_lo = float(it_lo.get(u"u_start", 0.0))
            u_start_hi = float(it_hi.get(u"u_start", 0.0))
        except Exception:
            continue
        if u_end_hi > u_end_lo + tol_u:
            out.append({
                u"wall": w_hi,
                u"item": it_hi,
                u"role": CORONAMIENTO_VOLADIZO_ROLE_SUP,
                u"wall_lower_idx": i - 1,
                u"wall_idx": i,
                u"side": u"der",
                u"u_reent": u_end_lo,
                u"u_free": u_end_hi,
            })
        elif u_end_lo > u_end_hi + tol_u:
            out.append({
                u"wall": w_lo,
                u"item": it_lo,
                u"role": CORONAMIENTO_VOLADIZO_ROLE_INF,
                u"wall_lower_idx": i - 1,
                u"wall_idx": i,
                u"side": u"der",
                u"u_reent": u_end_hi,
                u"u_free": u_end_lo,
            })
        if u_start_hi < u_start_lo - tol_u:
            out.append({
                u"wall": w_hi,
                u"item": it_hi,
                u"role": CORONAMIENTO_VOLADIZO_ROLE_SUP,
                u"wall_lower_idx": i - 1,
                u"wall_idx": i,
                u"side": u"izq",
                u"u_reent": u_start_lo,
                u"u_free": u_start_hi,
            })
        elif u_start_lo < u_start_hi - tol_u:
            out.append({
                u"wall": w_lo,
                u"item": it_lo,
                u"role": CORONAMIENTO_VOLADIZO_ROLE_INF,
                u"wall_lower_idx": i - 1,
                u"wall_idx": i,
                u"side": u"izq",
                u"u_reent": u_start_hi,
                u"u_free": u_start_lo,
            })
    return out


def _z_bar_voladizo_ft(wall, bar_type, role, fallback_diam_mm=None):
    """Cota del tramo horizontal según rol de voladizo (base sup. o tope inf.)."""
    axis_offset_ft = _mm_to_internal(_cover_axis_offset_mm(
        CORONAMIENTO_COVER_SUPERIOR_MM,
        bar_type,
        fallback_diam_mm=fallback_diam_mm,
    ))
    z_bot, z_top = _wall_z_bounds_ft(wall)
    if role == CORONAMIENTO_VOLADIZO_ROLE_INF:
        return float(z_top) - axis_offset_ft
    return float(z_bot) + axis_offset_ft


def _intervalos_u_voladizo_barra(spec, bar_type, fallback_diam_mm):
    """Calcula u_embed (sin pata L) y u_free (con pata L) en coordenadas del layout."""
    item = spec.get(u"item")
    if item is None:
        item = spec.get(u"item_hi")
    if item is None:
        return None, None, u"Sin item de layout."
    side = spec.get(u"side")
    try:
        u_reent = float(spec.get(u"u_reent"))
        u_free_raw = float(spec.get(u"u_free"))
        u_start = float(item.get(u"u_start", u_free_raw))
        u_end = float(item.get(u"u_end", u_free_raw))
    except Exception:
        return None, None, u"Coordenadas U inválidas."
    long_off_ft = _mm_to_internal(_cover_axis_offset_mm(
        CORONAMIENTO_COVER_SUPERIOR_MM,
        bar_type,
        fallback_diam_mm=fallback_diam_mm,
    ))
    tol_u = _voladizo_u_tolerance_ft()
    long_off_u = float(long_off_ft) if long_off_ft else 0.0
    d_mm = _bar_diameter_mm(bar_type, fallback_diam_mm)
    emb_mm = _empotramiento_voladizo_mm(d_mm)
    emb_u = float(_mm_to_internal(emb_mm))

    if side == u"der":
        u_free = float(u_free_raw) - long_off_u
        u_vol_inner = float(u_reent) + long_off_u
        u_embed = float(u_reent) - emb_u
        u_embed = max(float(u_start) + long_off_u, u_embed)
        if u_free <= u_vol_inner + tol_u:
            return None, None, u"Voladizo derecho demasiado corto."
    elif side == u"izq":
        u_free = float(u_free_raw) + long_off_u
        u_vol_inner = float(u_reent) - long_off_u
        u_embed = float(u_reent) + emb_u
        u_embed = min(float(u_end) - long_off_u, u_embed)
        if u_free >= u_vol_inner - tol_u:
            return None, None, u"Voladizo izquierdo demasiado corto."
    else:
        return None, None, u"Lado de voladizo no válido."

    min_len_ft = _mm_to_internal(CORONAMIENTO_VOLADIZO_MIN_HORIZ_MM)
    if abs(float(u_free) - float(u_embed)) < max(min_len_ft, tol_u):
        return None, None, u"Tramo voladizo+empotramiento demasiado corto."
    return float(u_embed), float(u_free), None


def _coronamiento_voladizo_chain_curves(
    wall,
    item,
    u_embed,
    u_free,
    z_bar_ft,
    leg_len_ft,
    bar_type,
    leg_up=True,
    fallback_diam_mm=None,
):
    """
    Polilínea en L: tramo horizontal (empotramiento + voladizo) + pata L en extremo libre.
    ``leg_up``: True → pata hacia +Z (voladizo sup.); False → hacia −Z (voladizo inf.).
    """
    frame = _wall_location_frame(wall)
    if frame is None:
        return None, None, None, u"Sin LocationCurve válida."
    axis_offset_mm = _cover_axis_offset_mm(
        CORONAMIENTO_COVER_SUPERIOR_MM,
        bar_type,
        fallback_diam_mm=fallback_diam_mm,
    )
    axis_offset_ft = _mm_to_internal(axis_offset_mm)
    normal = frame[u"normal"]
    espesor_ft = float(frame[u"espesor_ft"])
    lateral = espesor_ft * 0.5 - axis_offset_ft
    z_bar = float(z_bar_ft)

    pt_embed = _xyz_on_wall_at_u(wall, item, u_embed, z_bar, normal, lateral)
    pt_free = _xyz_on_wall_at_u(wall, item, u_free, z_bar, normal, lateral)
    if pt_embed is None or pt_free is None:
        return None, None, None, u"No se ubicaron extremos del voladizo."

    leg = max(float(leg_len_ft), 1e-6)
    try:
        if float(pt_embed.DistanceTo(pt_free)) < 1e-6:
            return None, None, None, u"Tramo horizontal nulo."
        up = XYZ.BasisZ if leg_up else XYZ.BasisZ.Negate()
        pt_free_leg = pt_free + up.Multiply(leg)
        c_horiz = Line.CreateBound(pt_embed, pt_free)
        c_leg = Line.CreateBound(pt_free, pt_free_leg)
    except Exception as ex_ln:
        return None, None, None, u"Line.CreateBound voladizo: {0}".format(ex_ln)

    distrib_ft = max(espesor_ft - 2.0 * axis_offset_ft, 0.0)
    return [c_horiz, c_leg], normal, distrib_ft, None


def _create_coronamiento_voladizo_rebar(
    doc,
    wall,
    item,
    u_embed,
    u_free,
    n_bars,
    bar_type,
    z_bar_ft,
    leg_up=True,
    fallback_diam_mm=None,
):
    if doc is None or wall is None or bar_type is None:
        return None, 0, u"Doc, muro o tipo de barra no válido."
    try:
        n_bars = int(n_bars)
    except Exception:
        n_bars = 2
    n_bars = max(2, min(4, n_bars))

    leg_mm = _largo_pata_l_mm(doc, wall, bar_type)
    leg_ft = _mm_to_internal(leg_mm)
    curves, normal, distrib_ft, err = _coronamiento_voladizo_chain_curves(
        wall,
        item,
        u_embed,
        u_free,
        z_bar_ft,
        leg_ft,
        bar_type,
        leg_up=leg_up,
        fallback_diam_mm=fallback_diam_mm,
    )
    if err:
        return None, 0, err

    try:
        curves_list = List[Curve]()
        for c in curves:
            curves_list.Add(c)
    except Exception as ex_cl:
        return None, 0, u"IList[Curve]: {0}".format(ex_cl)

    try:
        rebar = Rebar.CreateFromCurves(
            doc,
            RebarStyle.Standard,
            bar_type,
            None,
            None,
            wall,
            normal,
            curves_list,
            RebarHookOrientation.Left,
            RebarHookOrientation.Left,
            True,
            True,
        )
    except Exception as ex_cf:
        try:
            return None, 0, unicode(ex_cf)
        except Exception:
            return None, 0, str(ex_cf)

    if rebar is None:
        return None, 0, u"CreateFromCurves voladizo devolvió None."

    if n_bars > 1 and float(distrib_ft or 0.0) > 1e-9:
        try:
            accessor = rebar.GetShapeDrivenAccessor()
            accessor.SetLayoutAsFixedNumber(
                int(n_bars), float(distrib_ft), False, True, True,
            )
        except Exception as ex_lay:
            try:
                return None, 0, u"SetLayoutAsFixedNumber voladizo: {0}".format(unicode(ex_lay))
            except Exception:
                return None, 0, u"SetLayoutAsFixedNumber voladizo: {0}".format(str(ex_lay))

    if stamp_coronamiento_rebar is not None:
        try:
            stamp_coronamiento_rebar(rebar)
        except Exception:
            pass
    elif activar_armadura_arainco is not None:
        try:
            activar_armadura_arainco(rebar)
        except Exception:
            pass

    return rebar, int(n_bars), None


def _crear_coronamiento_voladizo_muro(doc, spec, bar_type, diam_mm, res):
    wall = spec.get(u"wall")
    item = spec.get(u"item") or spec.get(u"item_hi")
    side = spec.get(u"side") or u"?"
    role = spec.get(u"role") or CORONAMIENTO_VOLADIZO_ROLE_SUP
    if wall is None or item is None:
        return
    u_embed, u_free, err_u = _intervalos_u_voladizo_barra(spec, bar_type, diam_mm)
    if err_u:
        res[u"n_voladizo_fail"] = int(res.get(u"n_voladizo_fail", 0)) + 1
        try:
            wid = wall.Id.IntegerValue
        except Exception:
            wid = u"?"
        if len(res.get(u"messages") or []) < 24:
            res.setdefault(u"messages", []).append(
                u"Voladizo {0} muro Id {1} ({2}): {3}".format(role, wid, side, err_u),
            )
        return

    e_mm = _espesor_muro_mm(wall)
    n_bars = _n_bars_cfg_or_tipico(e_mm, res)
    if n_bars is None:
        res[u"n_voladizo_fail"] = int(res.get(u"n_voladizo_fail", 0)) + 1
        return

    leg_up = role != CORONAMIENTO_VOLADIZO_ROLE_INF
    z_bar = _z_bar_voladizo_ft(
        wall, bar_type, role, fallback_diam_mm=diam_mm,
    )
    rb, n_layout, err = _create_coronamiento_voladizo_rebar(
        doc,
        wall,
        item,
        u_embed,
        u_free,
        n_bars,
        bar_type,
        z_bar,
        leg_up=leg_up,
        fallback_diam_mm=diam_mm,
    )
    try:
        wid = wall.Id.IntegerValue
    except Exception:
        wid = u"?"
    if rb is None:
        res[u"n_voladizo_fail"] = int(res.get(u"n_voladizo_fail", 0)) + 1
        res.setdefault(u"messages", []).append(
            u"Voladizo {0} muro Id {1} ({2}): {3}".format(
                role, wid, side, err or u"error",
            ),
        )
        return

    res[u"n_voladizo_created"] = int(res.get(u"n_voladizo_created", 0)) + 1
    res[u"n_voladizo_bars"] = int(res.get(u"n_voladizo_bars", 0)) + int(n_layout)
    tag_ext = u"{0}_{1}_{2}".format(CORONAMIENTO_TAG_EXTREMO_VOL, role, side)
    _registrar_coronamiento_rebar_tag(res, rb, wall, z_bar, tag_ext)
    emb_mm = _empotramiento_voladizo_mm(diam_mm)
    role_txt = u"sup." if role == CORONAMIENTO_VOLADIZO_ROLE_SUP else u"inf. (bajo apil.)"
    pata_txt = u"arriba" if leg_up else u"abajo"
    res.setdefault(u"messages", []).append(
        u"Voladizo {0} muro Id {1} ({2}): {3}Ø{4} mm, empot.≈{5:.0f} mm, pata L {6}.".format(
            role_txt, wid, side, int(n_bars), int(diam_mm), float(emb_mm), pata_txt,
        ),
    )


def _crear_coronamiento_voladizos_stack(
    doc, walls_ord, bar_type_fallback, res,
    coronamiento_por_muro_id=None, global_config=None,
):
    if len(walls_ord or []) < 2:
        return
    layout = _compute_stacked_layout(walls_ord)
    if layout is None:
        return
    specs = _detect_voladizos_stack(walls_ord, layout)
    use_per_wall = coronamiento_por_muro_id is not None
    for spec in specs:
        wall = spec.get(u"wall")
        if wall is None:
            continue
        cfg = _cfg_for_wall(wall, coronamiento_por_muro_id, global_config)
        if use_per_wall:
            if not bool(cfg.get(u"activo", False)):
                continue
            if not bool(cfg.get(u"crear_voladizo", False)):
                continue
        elif global_config is not None:
            if not bool(cfg.get(u"crear_voladizo", False)):
                continue
        e_mm = _espesor_muro_mm(wall)
        if use_per_wall or global_config is not None:
            n_b, d_mm = _resolve_n_bars_diam_mm(cfg, e_mm)
        else:
            n_b = _n_bars_cfg_or_tipico(e_mm, res)
            d_mm = _diam_cfg_or_tipico(e_mm, res)
        if n_b is None or d_mm is None:
            res[u"n_voladizo_fail"] = int(res.get(u"n_voladizo_fail", 0)) + 1
            continue
        bt = _bar_type_for_diameter_mm(doc, d_mm, bar_type_fallback)
        if bt is None:
            res[u"n_voladizo_fail"] = int(res.get(u"n_voladizo_fail", 0)) + 1
            continue
        _crear_coronamiento_voladizo_muro(doc, spec, bt, d_mm, res)


def _wall_id_int(wall):
    if wall is None:
        return None
    try:
        return int(wall.Id.IntegerValue)
    except Exception:
        return None


def _registrar_coronamiento_rebar_tag(res, rb, wall, z_bar_ft, extremo, layer_index=0):
    """Acumula ids y metadatos para etiquetado (misma API que cabezal long.)."""
    if res is None or rb is None or wall is None:
        return
    try:
        res.setdefault(u"rebars_coronamiento_ids", []).append(rb.Id)
        try:
            res.setdefault(u"rebars_coronamiento_id_ints", []).append(
                int(rb.Id.IntegerValue),
            )
        except Exception:
            pass
        res.setdefault(u"rebars_coronamiento_tag_meta", []).append({
            u"rebar_id": rb.Id,
            u"layer_index": int(layer_index),
            u"wid": _wall_id_int(wall),
            u"extremo": extremo,
            u"zs": float(z_bar_ft),
            u"span_seg": 0.01,
        })
    except Exception:
        pass


def _coronamiento_tag_view(uidoc):
    if uidoc is None:
        return None
    try:
        return uidoc.ActiveView
    except Exception:
        return None


def _resolve_rebar_element_ids(doc, id_list):
    """Normaliza ``ElementId`` o enteros a ``ElementId`` válidos."""
    out = []
    seen = set()
    for item in id_list or []:
        if item is None:
            continue
        eid = item if hasattr(item, "IntegerValue") else None
        if eid is None:
            try:
                eid = ElementId(int(item))
            except Exception:
                eid = None
        if eid is None:
            continue
        try:
            iv = int(eid.IntegerValue)
        except Exception:
            continue
        if iv in seen:
            continue
        seen.add(iv)
        out.append(eid)
    return out


def aplicar_etiquetado_coronamiento(doc, cor_res, uidoc=None, aplicar_visibilidad=True):
    """
    Etiqueta barras de coronamiento (sup./inf./pie/vol.).

    Si ``aplicar_visibilidad`` es True (default), aplica Unobscured en la vista
    activa (pie/zapata). En el flujo unificado se pasa False: la visibilidad la
    hace ``aplicar_unobscured_armado_muros_en_vista`` al final (evita txn duplicada).
    """
    if not cor_res:
        return cor_res
    tag_ids = _resolve_rebar_element_ids(
        doc, cor_res.get(u"rebars_coronamiento_ids") or [],
    )
    if not tag_ids:
        tag_ids = _resolve_rebar_element_ids(
            doc, cor_res.get(u"rebars_coronamiento_id_ints") or [],
        )
    if not tag_ids:
        n_sets = (
            int(cor_res.get(u"n_created", 0))
            + int(cor_res.get(u"n_inferior_created", 0))
            + int(cor_res.get(u"n_inferior_pie_created", 0))
            + int(cor_res.get(u"n_voladizo_created", 0))
        )
        if n_sets > 0 and len(cor_res.get(u"messages") or []) < 24:
            cor_res.setdefault(u"messages", []).append(
                u"Etiquetas coronamiento: sin ids de barra registrados.",
            )
        return cor_res
    view = _coronamiento_tag_view(uidoc)
    if view is None or _cab_tags is None:
        cor_res[u"n_cor_tags_fail"] = int(cor_res.get(u"n_cor_tags_fail", 0)) + len(tag_ids)
        if len(cor_res.get(u"messages") or []) < 24:
            cor_res.setdefault(u"messages", []).append(
                u"Etiquetas coronamiento: sin vista activa o módulo de tags.",
            )
        if aplicar_visibilidad:
            try:
                vis = aplicar_visibilidad_coronamiento_en_vista(doc, uidoc, cor_res)
                cor_res[u"n_cor_unobscured"] = int(vis.get(u"n_unobscured", 0) or 0)
            except Exception:
                pass
        return cor_res
    tag_meta = list(cor_res.get(u"rebars_coronamiento_tag_meta") or [])
    try:
        doc.Regenerate()
    except Exception:
        pass
    # Modo rápido: lotes acotados (mismo techo que cabezal). Animación: 1/txn.
    batch_tags = 1
    try:
        modo_rapido = True
        if cabezal is not None:
            modo_rapido = bool(getattr(cabezal, u"MODO_EJECUCION_RAPIDA", True))
        if modo_rapido:
            cap = 50
            try:
                if cabezal is not None:
                    cap = int(getattr(cabezal, u"CABEZAL_TAGS_POR_LOTE_RAPIDO", 50))
            except Exception:
                cap = 50
            batch_tags = min(len(tag_ids), max(1, cap))
        else:
            batch_tags = 1
    except Exception:
        batch_tags = min(len(tag_ids), 50)
    batch_tags = max(1, int(batch_tags))
    try:
        tag_res = _cab_tags.etiquetar_cabezal_longitudinales_en_vista_animado(
            doc,
            view,
            tag_ids,
            tag_meta=tag_meta,
            batch_size=batch_tags,
        )
    except Exception as ex_tag:
        cor_res[u"n_cor_tags_fail"] = len(tag_ids)
        cor_res.setdefault(u"messages", []).append(
            u"Etiquetas coronamiento: {0}.".format(ex_tag),
        )
        if aplicar_visibilidad:
            try:
                vis = aplicar_visibilidad_coronamiento_en_vista(doc, uidoc, cor_res)
                cor_res[u"n_cor_unobscured"] = int(vis.get(u"n_unobscured", 0) or 0)
            except Exception:
                pass
        return cor_res
    cor_res[u"n_cor_tags_created"] = int(tag_res.get(u"n_ok", 0))
    cor_res[u"n_cor_tags_fail"] = int(tag_res.get(u"n_fail", 0))
    for m in tag_res.get(u"messages") or []:
        if m and len(cor_res.get(u"messages") or []) < 24:
            cor_res.setdefault(u"messages", []).append(m)
    if not aplicar_visibilidad:
        return cor_res
    try:
        vis = aplicar_visibilidad_coronamiento_en_vista(doc, uidoc, cor_res)
        cor_res[u"n_cor_unobscured"] = int(vis.get(u"n_unobscured", 0) or 0)
        cor_res[u"n_cor_unhide"] = int(vis.get(u"n_unhide", 0) or 0)
        if (
            int(vis.get(u"n_unobscured", 0) or 0) > 0
            and len(cor_res.get(u"messages") or []) < 24
        ):
            cor_res.setdefault(u"messages", []).append(
                u"Coronamiento: visible/Unobscured en vista ({0} barra(s)).".format(
                    int(vis.get(u"n_unobscured", 0) or 0),
                ),
            )
    except Exception as ex_vis:
        if len(cor_res.get(u"messages") or []) < 24:
            cor_res.setdefault(u"messages", []).append(
                u"Coronamiento visibilidad: {0}.".format(ex_vis),
            )
    return cor_res


def aplicar_visibilidad_coronamiento_en_vista(doc, uidoc, cor_res):
    """
    En la vista activa: categoría Rebar visible, UnhideElements,
    ``SetUnobscuredInView(True)`` y ``SetSolidInView(True)`` para coronamiento
    (crítico en pie/zapata, ocultas por el sólido de fundación).
    """
    out = {u"n_unobscured": 0, u"n_unhide": 0}
    if doc is None or uidoc is None or not cor_res:
        return out
    view = _coronamiento_tag_view(uidoc)
    if view is None:
        return out
    try:
        if getattr(view, u"IsTemplate", False):
            return out
    except Exception:
        pass

    tag_ids = _resolve_rebar_element_ids(
        doc, cor_res.get(u"rebars_coronamiento_ids") or [],
    )
    if not tag_ids:
        tag_ids = _resolve_rebar_element_ids(
            doc, cor_res.get(u"rebars_coronamiento_id_ints") or [],
        )
    rebars = []
    id_list = List[ElementId]()
    for eid in tag_ids:
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        try:
            from Autodesk.Revit.DB.Structure import Rebar as _Rebar

            if not isinstance(el, _Rebar):
                continue
        except Exception:
            pass
        rebars.append(el)
        try:
            id_list.Add(el.Id)
        except Exception:
            pass
    if not rebars:
        return out

    try:
        from bimtools_rebar_3d_visibility import apply_rebar_unobscured_in_view
    except Exception:
        apply_rebar_unobscured_in_view = None

    t = Transaction(doc, u"Arainco: Visibilidad coronamiento en vista")
    try:
        t.Start()
        # Categoría Rebar / Structural Rebar visible en la vista.
        try:
            from Autodesk.Revit.DB import BuiltInCategory

            try:
                cat = Category.GetCategory(doc, BuiltInCategory.OST_Rebar)
            except Exception:
                cat = None
            if cat is not None:
                try:
                    view.SetCategoryHidden(cat.Id, False)
                except Exception:
                    pass
        except Exception:
            pass
        if id_list.Count > 0:
            try:
                view.UnhideElements(id_list)
                out[u"n_unhide"] = int(id_list.Count)
            except Exception:
                pass
        # Presentación completa por barra (sección/alzado).
        try:
            from Autodesk.Revit.DB.Structure import RebarPresentationMode

            for rb in rebars:
                try:
                    rb.SetPresentationMode(view, RebarPresentationMode.All)
                except Exception:
                    pass
        except Exception:
            pass
        if apply_rebar_unobscured_in_view is not None:
            n = apply_rebar_unobscured_in_view(doc, rebars, view)
            try:
                out[u"n_unobscured"] = int(n or 0)
            except Exception:
                out[u"n_unobscured"] = len(rebars)
        else:
            for rb in rebars:
                try:
                    rb.SetUnobscuredInView(view, True)
                    out[u"n_unobscured"] += 1
                except Exception:
                    pass
                try:
                    rb.SetSolidInView(view, True)
                except Exception:
                    pass
        t.Commit()
    except Exception:
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except Exception:
            pass
    return out


def _create_coronamiento_inferior_fundacion_rebar(
    doc,
    geom_wall,
    fund,
    n_bars,
    bar_type,
    z_bar_ft,
    fallback_diam_mm=None,
    leg_host_geom=None,
):
    """
    Coronamiento inferior con host obligatorio = fundación.

    Usa la API de ``rebar_fundacion_cara_inferior`` (marco cara inferior, normales
    del host) en lugar de ``CreateFromCurves`` genérico orientado al muro.
    """
    if doc is None or geom_wall is None or fund is None or bar_type is None:
        return None, 0, u"Doc, muro, fundación o tipo de barra no válido."
    try:
        n_bars = int(n_bars)
    except Exception:
        n_bars = 2
    n_bars = max(2, min(4, n_bars))

    leg_mm = _largo_pata_l_sup_inf_sketch_mm(
        doc, leg_host_geom or geom_wall, bar_type, fallback_diam_mm=fallback_diam_mm,
    )
    leg_ft = _mm_to_internal(leg_mm)
    curves, normal, distrib_ft, err = _coronamiento_chain_curves(
        doc,
        geom_wall,
        z_bar_ft,
        leg_ft,
        bar_type,
        cover_mm=CORONAMIENTO_COVER_INFERIOR_MM,
        fallback_diam_mm=fallback_diam_mm,
        legs_up=True,
    )
    if err:
        return None, 0, err
    if not curves or len(curves) != 3:
        return None, 0, u"Polilínea U inválida para coronamiento inferior."

    c1, c2, c3 = curves[0], curves[1], curves[2]

    norm_prop, _w_unit = _norm_propagacion_coronamiento_inf_fund(c2, normal)
    # Sin retroceso media luz: la cadena ya parte en recub. desde cara (+normal);
    # restar distrib/2 cancelaría ese offset y centraría la barra en el eje del muro.

    poli = (c1, c2, c3)

    marco_uvn = None
    if obtener_marco_coordenadas_cara_inferior is not None:
        try:
            marco_uvn = obtener_marco_coordenadas_cara_inferior(fund)
        except Exception:
            marco_uvn = None

    cara_pp = None
    if evaluar_caras_paralelas_curva_mas_cercana is not None:
        try:
            ev = evaluar_caras_paralelas_curva_mas_cercana(fund, c2)
            if ev is not None:
                cara_pp = ev.get(u"mejor")
        except Exception:
            cara_pp = None

    z_hook = XYZ.BasisZ
    if vector_reverso_cara_paralela_mas_cercana_a_barra is not None:
        try:
            zh = vector_reverso_cara_paralela_mas_cercana_a_barra(
                fund,
                c2,
                excluir_caras_tapas_horizontales=True,
            )
            if zh is not None:
                z_hook = zh
        except Exception:
            pass

    norm_pri = [norm_prop] if norm_prop is not None else None
    if norm_pri is None and normal is not None:
        norm_pri = [normal]

    rb = None
    err_msg = None
    norm_create = None
    fund_id = _element_id_int(fund)

    if crear_rebar_u_shape_desde_eje_rebar_shape_nombrado is not None:
        rb, err_msg, norm_create = crear_rebar_u_shape_desde_eje_rebar_shape_nombrado(
            doc,
            fund,
            bar_type,
            poli,
            shape_nombre=REBAR_SHAPE_NOMBRE_DEFECTO,
            marco_cara_uvn=marco_uvn,
            cara_paralela=cara_pp,
            eje_referencia_z_ganchos=z_hook,
            normales_prioridad=norm_pri,
        )
    if rb is None and crear_rebar_polilinea_u_malla_inf_sup_curve_loop is not None:
        rb, err_msg, norm_create = crear_rebar_polilinea_u_malla_inf_sup_curve_loop(
            doc,
            fund,
            bar_type,
            poli,
            c2,
            marco_cara_uvn=marco_uvn,
            cara_paralela=cara_pp,
            eje_referencia_z_ganchos=z_hook,
            normales_prioridad=norm_pri,
        )
    if rb is None and crear_rebar_polilinea_recta_sin_ganchos is not None:
        rb, err_msg, norm_create = crear_rebar_polilinea_recta_sin_ganchos(
            doc,
            fund,
            bar_type,
            poli,
            c2,
            marco_cara_uvn=marco_uvn,
            cara_paralela=cara_pp,
            eje_referencia_z_ganchos=z_hook,
            normales_prioridad=norm_pri,
        )

    if rb is None:
        return None, 0, err_msg or u"No se pudo crear coronamiento inf. en fundación."

    host_id = _rebar_host_id_int(rb)
    if fund_id is None or host_id != fund_id:
        _delete_rebar_safe(doc, rb)
        return None, 0, (
            u"Host incorrecto: esperado fund. Id {0}, obtenido {1}.".format(
                fund_id if fund_id is not None else u"?",
                host_id if host_id is not None else u"?",
            )
        )

    n_layout = 1
    if n_bars > 1 and float(distrib_ft or 0.0) > 1e-9:
        layout_ok = _aplicar_layout_coronamiento_inf_fundacion(
            rb, doc, int(n_bars), float(distrib_ft),
        )
        if not layout_ok:
            _delete_rebar_safe(doc, rb)
            return None, 0, u"No se pudo aplicar layout Fixed Number en fundación."
        try:
            n_layout = int(rb.Quantity)
        except Exception:
            try:
                n_layout = int(rb.NumberOfBarPositions)
            except Exception:
                n_layout = int(n_bars)

    if stamp_coronamiento_rebar is not None:
        try:
            stamp_coronamiento_rebar(rb)
        except Exception:
            pass
    elif activar_armadura_arainco is not None:
        try:
            activar_armadura_arainco(rb)
        except Exception:
            pass

    return rb, int(n_layout), None


def _crear_coronamiento_superior(doc, wall, bar_type, diam_mm, res, n_bars=None):
    if n_bars is None:
        n_bars = res.get(u"n_bars_spec")
    try:
        n_bars = max(2, min(4, int(n_bars)))
    except Exception:
        n_bars = int(res.get(u"n_bars_spec") or 2)
    z_bar = _z_bar_superior_ft(wall, bar_type, fallback_diam_mm=diam_mm)
    rb, n_layout, err = _create_coronamiento_rebar(
        doc,
        wall,
        wall,
        n_bars,
        bar_type,
        z_bar,
        cover_mm=CORONAMIENTO_COVER_SUPERIOR_MM,
        fallback_diam_mm=diam_mm,
        legs_up=False,
    )
    if rb is None:
        res[u"n_fail"] += 1
        res[u"messages"].append(err or u"Coronamiento superior no creado.")
        return
    res[u"n_created"] += 1
    res[u"n_bars"] += int(n_layout)
    _registrar_coronamiento_rebar_tag(
        res, rb, wall, z_bar, CORONAMIENTO_TAG_EXTREMO_SUP,
    )
    try:
        wid = wall.Id.IntegerValue
    except Exception:
        wid = res.get(u"host_wall_id") or u"?"
    res[u"messages"].append(
        u"Coronamiento sup. muro Id {0}: {1}Ø{2} mm ({3}).".format(
            wid,
            int(n_bars),
            int(diam_mm),
            _coronamiento_extremos_resumen(doc, wall),
        ),
    )


def _crear_coronamiento_inferior_pie_muro(doc, wall, bar_type, diam_mm, res, n_bars=None):
    """Coronamiento U en el pie del muro cuando no hay apilamiento ni fundación debajo."""
    z_bar = _z_bar_inferior_pie_ft(wall, bar_type, fallback_diam_mm=diam_mm)
    e_mm = _espesor_muro_mm(wall)
    if n_bars is None:
        n_bars = _n_bars_cfg_or_tipico(e_mm, res)
    if n_bars is None:
        res[u"n_inferior_pie_fail"] = int(res.get(u"n_inferior_pie_fail", 0)) + 1
        return

    rb, n_layout, err = _create_coronamiento_rebar(
        doc,
        wall,
        wall,
        n_bars,
        bar_type,
        z_bar,
        cover_mm=CORONAMIENTO_COVER_SUPERIOR_MM,
        fallback_diam_mm=diam_mm,
        legs_up=True,
    )
    try:
        wid = wall.Id.IntegerValue
    except Exception:
        wid = u"?"
    if rb is None:
        res[u"n_inferior_pie_fail"] = int(res.get(u"n_inferior_pie_fail", 0)) + 1
        res.setdefault(u"messages", []).append(
            u"Coronamiento pie muro Id {0}: {1}".format(wid, err or u"error"),
        )
        return

    res[u"n_inferior_pie_created"] = int(res.get(u"n_inferior_pie_created", 0)) + 1
    res[u"n_inferior_pie_bars"] = int(res.get(u"n_inferior_pie_bars", 0)) + int(n_layout)
    _registrar_coronamiento_rebar_tag(
        res, rb, wall, z_bar, CORONAMIENTO_TAG_EXTREMO_PIE,
    )
    res.setdefault(u"messages", []).append(
        u"Coronamiento pie muro Id {0}: {1}Ø{2} mm (sin apil./fund.; {3}).".format(
            wid, int(n_bars), int(diam_mm),
            _coronamiento_extremos_resumen(doc, wall),
        ),
    )


def _crear_coronamiento_inferior_muro(doc, wall, bar_type, diam_mm, res, n_bars=None):
    fund = _fundacion_principal_muro(doc, wall)
    if fund is None:
        return
    z_bar, err_z = _z_bar_inferior_fundacion_ft(wall, fund, bar_type, diam_mm)
    if err_z:
        res[u"n_inferior_fail"] += 1
        try:
            wid = wall.Id.IntegerValue
        except Exception:
            wid = u"?"
        res[u"messages"].append(
            u"Coronamiento inf. muro Id {0}: {1}".format(wid, err_z),
        )
        return
    e_mm = _espesor_muro_mm(wall)
    if n_bars is None:
        n_bars = _n_bars_cfg_or_tipico(e_mm, res)
    if n_bars is None:
        res[u"n_inferior_fail"] += 1
        return

    rb, n_layout, err = _create_coronamiento_inferior_fundacion_rebar(
        doc,
        wall,
        fund,
        n_bars,
        bar_type,
        z_bar,
        fallback_diam_mm=diam_mm,
        leg_host_geom=wall,
    )
    try:
        wid = wall.Id.IntegerValue
    except Exception:
        wid = u"?"
    try:
        fid = fund.Id.IntegerValue
    except Exception:
        fid = u"?"
    if rb is None:
        res[u"n_inferior_fail"] += 1
        res[u"messages"].append(
            u"Coronamiento inf. muro Id {0} (host fund. {1}): {2}".format(
                wid, fid, err or u"error",
            ),
        )
        return
    res[u"n_inferior_created"] += 1
    res[u"n_inferior_bars"] += int(n_layout)
    _registrar_coronamiento_rebar_tag(
        res, rb, wall, z_bar, CORONAMIENTO_TAG_EXTREMO_INF,
    )
    h_mm = _altura_fundacion_mm(fund)
    h_txt = u"{0:.0f}".format(float(h_mm)) if h_mm is not None else u"?"
    res[u"messages"].append(
        u"Coronamiento inf. muro Id {0}, host fund. Id {1}: {2}Ø{3} mm "
        u"(e={4} mm, H_fund≈{5} mm, recub. cara inf. 50+Ø/2; {6}).".format(
            wid, fid, int(n_bars), int(diam_mm), int(round(e_mm)), h_txt,
            _coronamiento_extremos_resumen(doc, wall),
        ),
    )


def aplicar_coronamiento_muros(
    doc, walls, bar_type_fallback=None, config=None, coronamiento_por_muro_id=None,
):
    """
    Coronamiento superior, inferior (fundación / pie), voladizo (reentrada).

    ``config`` (opcional): dict global legacy de ``normalize_coronamiento_config``.
    ``coronamiento_por_muro_id`` (opcional): mapa wid→cfg por muro (UI V3).
    """
    res = {
        u"n_created": 0,
        u"n_fail": 0,
        u"n_bars": 0,
        u"n_inferior_created": 0,
        u"n_inferior_fail": 0,
        u"n_inferior_bars": 0,
        u"n_inferior_pie_created": 0,
        u"n_inferior_pie_fail": 0,
        u"n_inferior_pie_bars": 0,
        u"n_voladizo_created": 0,
        u"n_voladizo_fail": 0,
        u"n_voladizo_bars": 0,
        u"n_cor_tags_created": 0,
        u"n_cor_tags_fail": 0,
        u"rebars_coronamiento_ids": [],
        u"rebars_coronamiento_id_ints": [],
        u"rebars_coronamiento_tag_meta": [],
        u"messages": [],
        u"host_wall_id": None,
        u"n_bars_spec": 0,
        u"diam_mm": None,
        u"espesor_mm": None,
        u"cfg_n_bars": None,
        u"cfg_diam_mm": None,
        u"skipped": False,
    }
    if doc is None:
        res[u"messages"].append(u"Sin documento Revit.")
        res[u"n_fail"] = 1
        return res

    if ordenar_muros_por_base_asc is not None:
        walls_ord = ordenar_muros_por_base_asc(walls)
    else:
        walls_ord = list(walls or [])

    wall_top = muro_tope_stack_global(walls_ord)
    if wall_top is None:
        res[u"messages"].append(u"Sin muros para coronamiento.")
        return res

    try:
        res[u"host_wall_id"] = wall_top.Id.IntegerValue
    except Exception:
        pass

    e_mm_top = _espesor_muro_mm(wall_top)
    res[u"espesor_mm"] = int(round(e_mm_top))
    cor_map = _normalize_coronamiento_por_muro_id(coronamiento_por_muro_id)
    # Mapa presente (aunque vacío o todos inactivos) = modo opt-in por muro.
    use_per_wall = coronamiento_por_muro_id is not None

    if use_per_wall:
        any_active = False
        for wall in walls_ord:
            if wall is None:
                continue
            cfg_w = _cfg_for_wall(wall, cor_map, config)
            if bool(cfg_w.get(u"activo", False)):
                any_active = True
                break
        if not any_active:
            res[u"skipped"] = True
            res[u"messages"].append(
                u"Coronamiento omitido (ningún muro activo en configuración).",
            )
            return res
    else:
        cfg = normalize_coronamiento_config(config, espesor_mm=e_mm_top, wall=wall_top)
        if not bool(cfg.get(u"activo", True)):
            res[u"skipped"] = True
            res[u"messages"].append(u"Coronamiento omitido (desactivado en configuración).")
            return res

        n_bars_top = int(cfg.get(u"n_bars") or 0) or None
        diam_mm_top = int(cfg.get(u"diam_mm") or 0) or None
        if n_bars_top is None or diam_mm_top is None:
            tip_n, tip_d = coronamiento_tipico_por_espesor_mm(e_mm_top)
            if n_bars_top is None:
                n_bars_top = tip_n
            if diam_mm_top is None:
                diam_mm_top = tip_d
        if n_bars_top is None or diam_mm_top is None:
            res[u"n_fail"] = 1
            res[u"messages"].append(
                u"Espesor e={0} mm sin regla típica de coronamiento.".format(
                    int(round(e_mm_top)),
                ),
            )
            return res

        n_bars_top = max(2, min(4, int(n_bars_top)))
        diam_mm_top = int(diam_mm_top)
        res[u"n_bars_spec"] = int(n_bars_top)
        res[u"diam_mm"] = int(diam_mm_top)
        res[u"cfg_n_bars"] = int(n_bars_top)
        res[u"cfg_diam_mm"] = int(diam_mm_top)

        bar_type_top = _bar_type_for_diameter_mm(doc, diam_mm_top, bar_type_fallback)
        if bar_type_top is None:
            res[u"n_fail"] = 1
            res[u"messages"].append(
                u"Sin RebarBarType Ø{0} mm para coronamiento.".format(int(diam_mm_top)),
            )
            return res

        crear_sup = bool(cfg.get(u"crear_superior", False))
        crear_inf = bool(cfg.get(u"crear_inferior", False))
        crear_vol = bool(cfg.get(u"crear_voladizo", False))

        txn = Transaction(doc, u"Arainco: Coronamiento muros")
        t_started = False
        try:
            try:
                from armado_muros_txn import attach_rebar_outside_host_swallower
                attach_rebar_outside_host_swallower(txn)
            except Exception:
                pass
            txn.Start()
            t_started = True
            if crear_sup:
                _crear_coronamiento_superior(
                    doc, wall_top, bar_type_top, diam_mm_top, res,
                    n_bars=n_bars_top,
                )

            if crear_inf:
                for wall in walls_ord:
                    if wall is None:
                        continue
                    e_mm = _espesor_muro_mm(wall)
                    d_mm = _diam_cfg_or_tipico(e_mm, res)
                    if d_mm is None:
                        continue
                    bt = _bar_type_for_diameter_mm(doc, d_mm, bar_type_fallback)
                    if bt is None:
                        res[u"n_inferior_fail"] += 1
                        continue
                    fund = _fundacion_principal_muro(doc, wall)
                    if fund is not None:
                        _crear_coronamiento_inferior_muro(doc, wall, bt, d_mm, res)
                    elif not _muro_tiene_apilamiento_inferior(wall, walls_ord):
                        _crear_coronamiento_inferior_pie_muro(doc, wall, bt, d_mm, res)

            if crear_vol:
                _crear_coronamiento_voladizos_stack(
                    doc, walls_ord, bar_type_fallback, res,
                    global_config=config,
                )

            if (
                res[u"n_created"] < 1
                and res[u"n_inferior_created"] < 1
                and res[u"n_inferior_pie_created"] < 1
                and res[u"n_voladizo_created"] < 1
            ):
                txn.RollBack()
                t_started = False
                if (
                    res[u"n_fail"] < 1
                    and res[u"n_inferior_fail"] < 1
                    and res[u"n_inferior_pie_fail"] < 1
                    and res[u"n_voladizo_fail"] < 1
                ):
                    if not (crear_sup or crear_inf or crear_vol):
                        res[u"skipped"] = True
                        res[u"messages"].append(
                            u"Coronamiento: ningún tipo seleccionado.",
                        )
                    else:
                        res[u"n_fail"] = 1
                        res[u"messages"].append(u"No se creó ningún coronamiento.")
                return res

            txn.Commit()
            t_started = False
        except Exception as ex:
            if t_started:
                try:
                    txn.RollBack()
                except Exception:
                    pass
            res[u"n_fail"] += 1
            res[u"messages"].append(u"Coronamiento: {0}".format(ex))
        return res

    txn = Transaction(doc, u"Arainco: Coronamiento muros")
    t_started = False
    crear_any_type = False
    try:
        try:
            from armado_muros_txn import attach_rebar_outside_host_swallower
            attach_rebar_outside_host_swallower(txn)
        except Exception:
            pass
        txn.Start()
        t_started = True
        for wall in walls_ord:
            if wall is None:
                continue
            cfg_w = _cfg_for_wall(wall, cor_map, config)
            if not bool(cfg_w.get(u"activo", False)):
                continue
            e_mm = _espesor_muro_mm(wall)
            n_bars, diam_mm = _resolve_n_bars_diam_mm(cfg_w, e_mm)
            if n_bars is None or diam_mm is None:
                res[u"n_fail"] += 1
                continue
            bt = _bar_type_for_diameter_mm(doc, diam_mm, bar_type_fallback)
            if bt is None:
                res[u"n_fail"] += 1
                continue
            if res.get(u"diam_mm") is None:
                res[u"diam_mm"] = int(diam_mm)
                res[u"cfg_diam_mm"] = int(diam_mm)
                res[u"cfg_n_bars"] = int(n_bars)
                res[u"n_bars_spec"] = int(n_bars)

            if bool(cfg_w.get(u"crear_superior", False)):
                crear_any_type = True
                _crear_coronamiento_superior(
                    doc, wall, bt, diam_mm, res, n_bars=n_bars,
                )

            if bool(cfg_w.get(u"crear_inferior", False)):
                crear_any_type = True
                fund = _fundacion_principal_muro(doc, wall)
                if fund is not None:
                    _crear_coronamiento_inferior_muro(
                        doc, wall, bt, diam_mm, res, n_bars=n_bars,
                    )
                elif not _muro_tiene_apilamiento_inferior(wall, walls_ord):
                    _crear_coronamiento_inferior_pie_muro(
                        doc, wall, bt, diam_mm, res, n_bars=n_bars,
                    )

        _crear_coronamiento_voladizos_stack(
            doc, walls_ord, bar_type_fallback, res,
            coronamiento_por_muro_id=cor_map,
            global_config=config,
        )
        if int(res.get(u"n_voladizo_created", 0)) > 0:
            crear_any_type = True

        if (
            res[u"n_created"] < 1
            and res[u"n_inferior_created"] < 1
            and res[u"n_inferior_pie_created"] < 1
            and res[u"n_voladizo_created"] < 1
        ):
            txn.RollBack()
            t_started = False
            if (
                res[u"n_fail"] < 1
                and res[u"n_inferior_fail"] < 1
                and res[u"n_inferior_pie_fail"] < 1
                and res[u"n_voladizo_fail"] < 1
            ):
                if not crear_any_type:
                    res[u"skipped"] = True
                    res[u"messages"].append(
                        u"Coronamiento: ningún tipo seleccionado en muros activos.",
                    )
                else:
                    res[u"n_fail"] = 1
                    res[u"messages"].append(u"No se creó ningún coronamiento.")
            return res

        txn.Commit()
        t_started = False
    except Exception as ex:
        if t_started:
            try:
                txn.RollBack()
            except Exception:
                pass
        res[u"n_fail"] += 1
        res[u"messages"].append(u"Coronamiento: {0}".format(ex))
    return res
