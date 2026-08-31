# -*- coding: utf-8 -*-
"""
Creación de armadura en fundación aislada (motor compartido).

Malla inferior / superior: dos conjuntos ``Rebar`` (luz menor / luz mayor) en U
vía ``CreateFromCurvesAndShape`` (forma «03») + ``SetLayoutAsMaximumSpacing``.
Soporta **diámetro y separación independientes por luz**.

Lateral: barras horizontales perimetrales + Maximum Spacing en altura.

Revit 2024–2026 · IronPython.
"""

from __future__ import print_function

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    StorageType,
    Transaction,
    Transform,
    View,
    ViewSection,
    XYZ,
)

from geometria_fundacion_cara_inferior import (
    RECUBRIMIENTO_EXTREMOS_MM,
    aplicar_recubrimiento_inferior_completo_mm,
    construir_polilinea_u_fundacion_desde_eje_horizontal,
    evaluar_caras_paralelas_curva_mas_cercana,
    extraer_curva_lado_mayor_cara_inferior,
    extraer_curva_lado_mayor_cara_superior,
    extraer_curva_lado_menor_cara_inferior,
    extraer_curva_lado_menor_cara_superior,
    largo_gancho_u_tabla_mm,
    linea_horizontal_cara_lateral_a_cota_z,
    lineas_horizontales_perimetro_inferior_exterior,
    longitud_array_lateral_altura_fundacion_menos_mm_ft,
    longitud_distribucion_perpendicular_barra_inferior_ft,
    longitud_pata_u_fundacion_inf_sup_ft,
    normal_saliente_horizontal_paramento_para_barra_horizontal,
    obtener_marco_coordenadas_cara_inferior,
    obtener_marco_coordenadas_cara_superior,
    offset_linea_adicional_hacia_interior_mm,
    offset_linea_eje_barra_desde_cara_inferior_mm,
    offset_linea_hacia_interior_desde_cara_inferior_mm,
    primera_cota_z_armadura_lateral_ft,
    rango_z_caras_laterales_o_bbox,
    vector_reverso_cara_paralela_mas_cercana_a_barra,
)
from rebar_fundacion_cara_inferior import (
    REBAR_SHAPE_NOMBRE_DEFECTO,
    aplicar_layout_maximum_spacing_rebar,
    crear_rebar_desde_curva_linea_con_ganchos,
    crear_rebar_polilinea_recta_sin_ganchos,
    crear_rebar_polilinea_u_malla_inf_sup_curve_loop,
    crear_rebar_u_shape_desde_eje_rebar_shape_nombrado,
)

try:
    from barras_bordes_losa_gancho_empotramiento import (
        _rebar_nominal_diameter_mm,
        element_id_to_int,
    )
except Exception:
    _rebar_nominal_diameter_mm = None

    def element_id_to_int(eid):
        if eid is None:
            return None
        try:
            return int(eid.IntegerValue)
        except Exception:
            pass
        try:
            return int(eid.Value)
        except Exception:
            return None

try:
    from conjunto_guid import (
        iniciar_armadura_conjunto_guid_ejecucion,
        stamp_armadura_arainco,
        stamp_armadura_conjunto_guid,
        stamp_armadura_malla,
    )
except Exception:
    iniciar_armadura_conjunto_guid_ejecucion = None
    stamp_armadura_conjunto_guid = None
    stamp_armadura_malla = None
    stamp_armadura_arainco = None


# ---------------------------------------------------------------------------
# Constantes (alineadas con enfierrado_fundacion_aislada)
# ---------------------------------------------------------------------------

_REC_PLANTA_MM = 100.0
_REC_HORIZONTAL_MM = 50.0
_REC_LATERAL_CARA_MM = 50.0
_OFFSET_PRIMERA_LATERAL_MM = 100.0
_DESCUENTO_ARRAY_LATERAL_MM = 200.0
_DESCUENTO_PATA_U_MM = 150.0

# Exportados para preview del sketch (misma fuente de verdad).
REC_PLANTA_MALLA_MM = _REC_PLANTA_MM
REC_HORIZONTAL_EJE_MM = _REC_HORIZONTAL_MM
REC_LATERAL_CARA_MM = _REC_LATERAL_CARA_MM
OFFSET_PRIMERA_LATERAL_MM = _OFFSET_PRIMERA_LATERAL_MM
DESCUENTO_ARRAY_LATERAL_MM = _DESCUENTO_ARRAY_LATERAL_MM
DESCUENTO_PATA_U_MM = _DESCUENTO_PATA_U_MM

_ARMA_UBICACION_PARAM = u"Armadura_Ubicacion"
_ARMA_UBICACION_INFERIOR = u"F"
_ARMA_UBICACION_SUPERIOR = u"F'"
_ARMA_UBICACION_LATERAL = u"L"

_TXN_NAME = u"Arainco: Armadura fundación aislada (sketch)"

_PARAM_SECTION_FILTER = u"Section Filter"
_DETALLE_MARGEN_MM = 400.0
_DETALLE_PROFUNDIDAD_MM = 800.0


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _d_mm(bar_type, fallback=8.0):
    if bar_type is not None and _rebar_nominal_diameter_mm is not None:
        try:
            v = _rebar_nominal_diameter_mm(bar_type)
            if v is not None:
                return float(v)
        except Exception:
            pass
    return float(fallback)


def _aplicar_ubicacion(rebar, valor):
    if rebar is None or valor is None:
        return
    try:
        p = rebar.LookupParameter(_ARMA_UBICACION_PARAM)
        if p is None or p.IsReadOnly:
            return
        p.Set(valor)
    except Exception:
        pass


def _aplicar_guid(rebar, conjunto_guid):
    if rebar is None or stamp_armadura_conjunto_guid is None:
        return
    try:
        stamp_armadura_conjunto_guid(rebar, conjunto_guid=conjunto_guid)
    except Exception:
        pass


def _aplicar_flags_malla_arainco(rebar):
    """``Armadura_Malla`` = Yes y ``Armadura_Arainco`` = Yes en cada Rebar creado."""
    if rebar is None:
        return
    if stamp_armadura_malla is not None:
        try:
            stamp_armadura_malla(rebar, yes=True)
        except Exception:
            pass
    if stamp_armadura_arainco is not None:
        try:
            stamp_armadura_arainco(rebar, yes=True)
        except Exception:
            pass


def _aplicar_params_rebar(rebar, ubicacion, conjunto_guid):
    _aplicar_ubicacion(rebar, ubicacion)
    _aplicar_guid(rebar, conjunto_guid)
    _aplicar_flags_malla_arainco(rebar)


def _curvas_mismo_segmento(c0, c1, tol_ft=None):
    if c0 is None or c1 is None:
        return False
    if tol_ft is None:
        tol_ft = 1.0 / 304.8
    try:
        a0, a1 = c0.GetEndPoint(0), c0.GetEndPoint(1)
        b0, b1 = c1.GetEndPoint(0), c1.GetEndPoint(1)
        return (
            a0.DistanceTo(b0) < tol_ft and a1.DistanceTo(b1) < tol_ft
        ) or (
            a0.DistanceTo(b1) < tol_ft and a1.DistanceTo(b0) < tol_ft
        )
    except Exception:
        return False


def _altura_nominal_fundacion_ft(elem):
    if elem is None:
        return None
    try:
        from Autodesk.Revit.DB import BuiltInParameter, UnitTypeId, UnitUtils
    except Exception:
        BuiltInParameter = None
        UnitUtils = None
        UnitTypeId = None
    tipo = None
    try:
        tipo = elem.Document.GetElement(elem.GetTypeId())
    except Exception:
        tipo = None
    try:
        name = None
        if tipo is not None and BuiltInParameter is not None:
            try:
                p = tipo.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if p is not None:
                    name = p.AsString()
            except Exception:
                name = None
        if not name and tipo is not None:
            try:
                name = tipo.Name
            except Exception:
                name = None
        if name:
            parts = _as_unicode(name).lower().replace(u"×", u"x").split(u"x")
            if len(parts) >= 3:
                tok = parts[-1].strip()
                num = u""
                for ch in tok:
                    if ch.isdigit() or ch in u".,":
                        num += ch
                    elif num:
                        break
                if num:
                    mm = float(num.replace(u",", u"."))
                    if mm > 10.0 and UnitUtils is not None:
                        return float(
                            UnitUtils.ConvertToInternalUnits(
                                mm, UnitTypeId.Millimeters
                            )
                        )
                    if mm > 10.0:
                        return mm / 304.8
    except Exception:
        pass
    for src in (elem, tipo):
        if src is None:
            continue
        for pname in (
            u"Height",
            u"Altura",
            u"Espesor",
            u"Thickness",
            u"h",
            u"Alto",
            u"Foundation Height",
            u"Depth",
        ):
            try:
                p = src.LookupParameter(pname)
                if p is not None and p.HasValue and p.StorageType == StorageType.Double:
                    val = float(p.AsDouble())
                    if val > 0.01:
                        return val
            except Exception:
                continue
    return None


def _leg_ft_pata_u_inferior(z0p, z1p, d_mm_bar, sup_on, elem=None):
    try:
        descuento_mm = float(_DESCUENTO_PATA_U_MM) + float(d_mm_bar) / 2.0
    except Exception:
        descuento_mm = float(_DESCUENTO_PATA_U_MM)
    h_param_ft = _altura_nominal_fundacion_ft(elem) if elem is not None else None
    if h_param_ft is not None:
        leg_max = longitud_pata_u_fundacion_inf_sup_ft(0.0, h_param_ft, descuento_mm)
    else:
        leg_max = longitud_pata_u_fundacion_inf_sup_ft(z0p, z1p, descuento_mm)
    if sup_on:
        return leg_max
    hook_mm = largo_gancho_u_tabla_mm(d_mm_bar)
    leg_ft = None
    if hook_mm is not None:
        try:
            from bimtools_rebar_hook_lengths import pata_eje_curve_loop_mm_desde_tabla_mm

            eje_mm = pata_eje_curve_loop_mm_desde_tabla_mm(hook_mm, d_mm_bar)
        except Exception:
            eje_mm = None
        if eje_mm is None:
            try:
                eje_mm = float(hook_mm) - float(d_mm_bar) / 2.0
            except Exception:
                eje_mm = float(hook_mm)
        try:
            from Autodesk.Revit.DB import UnitUtils, UnitTypeId

            leg_ft = UnitUtils.ConvertToInternalUnits(
                float(eje_mm), UnitTypeId.Millimeters
            )
        except Exception:
            leg_ft = float(eje_mm) / 304.8
    if leg_ft is not None and leg_max is not None:
        return min(leg_ft, leg_max)
    if leg_ft is not None:
        return leg_ft
    return leg_max


def _leg_ft_pata_u_superior(z0p, z1p, d_mm_bar, elem=None):
    try:
        descuento_mm = float(_DESCUENTO_PATA_U_MM) + float(d_mm_bar) / 2.0
    except Exception:
        descuento_mm = float(_DESCUENTO_PATA_U_MM)
    h_param_ft = _altura_nominal_fundacion_ft(elem)
    if h_param_ft is not None:
        return longitud_pata_u_fundacion_inf_sup_ft(0.0, h_param_ft, descuento_mm)
    return longitud_pata_u_fundacion_inf_sup_ft(z0p, z1p, descuento_mm)


def _dir_cfg(settings_group, dir_key):
    d = (settings_group or {}).get(dir_key) or {}
    bt = d.get(u"bar_type")
    try:
        sep = float(d.get(u"spacing_mm") or 150.0)
    except Exception:
        sep = 150.0
    dmm = d.get(u"diameter_mm")
    if dmm is None:
        dmm = _d_mm(bt, 8.0)
    else:
        try:
            dmm = float(dmm)
        except Exception:
            dmm = _d_mm(bt, 8.0)
    return bt, sep, dmm


def _extremos_mm(d_mm):
    return float(RECUBRIMIENTO_EXTREMOS_MM) + 0.5 * float(d_mm or 0.0)


def _array_len_ft(el, curva_tratada, perp_len_mm, rec_planta_mm, lado_etq):
    if perp_len_mm is not None:
        try:
            from Autodesk.Revit.DB import UnitUtils, UnitTypeId

            span_mm = max(float(perp_len_mm) - 2.0 * float(rec_planta_mm), 0.01)
            return float(
                UnitUtils.ConvertToInternalUnits(span_mm, UnitTypeId.Millimeters)
            )
        except Exception:
            pass
    return longitud_distribucion_perpendicular_barra_inferior_ft(
        el, curva_tratada, rec_planta_mm, lado_etq
    )


def _crear_rebar_u_o_fallback(
    doc,
    el,
    bt,
    curva_rebar,
    n_cara,
    leg_ft,
    d_mm_bar,
    marco_uvn,
    cara_pp,
    z_hook,
    allow_u,
):
    """Devuelve (rebar, err, norm, uso_fallback_ganchos)."""
    r, err_rb, norm_rb = None, None, None
    fallback = False
    poli_u = None
    if allow_u and leg_ft is not None and n_cara is not None:
        poli_u = construir_polilinea_u_fundacion_desde_eje_horizontal(
            curva_rebar,
            n_cara,
            leg_ft,
            d_mm_bar,
            acortar_eje_central_para_cota_revit=False,
        )
    if poli_u is not None:
        r, err_rb, norm_rb = crear_rebar_u_shape_desde_eje_rebar_shape_nombrado(
            doc,
            el,
            bt,
            poli_u,
            shape_nombre=REBAR_SHAPE_NOMBRE_DEFECTO,
            marco_cara_uvn=marco_uvn,
            cara_paralela=cara_pp,
            eje_referencia_z_ganchos=z_hook,
        )
        if r is None:
            r, err_rb, norm_rb = crear_rebar_polilinea_u_malla_inf_sup_curve_loop(
                doc,
                el,
                bt,
                poli_u,
                poli_u[1],
                marco_cara_uvn=marco_uvn,
                cara_paralela=cara_pp,
                eje_referencia_z_ganchos=z_hook,
            )
        if r is None:
            r, err_rb, norm_rb = crear_rebar_polilinea_recta_sin_ganchos(
                doc,
                el,
                bt,
                poli_u,
                poli_u[1],
                marco_cara_uvn=marco_uvn,
                cara_paralela=cara_pp,
                eje_referencia_z_ganchos=z_hook,
            )
    if r is None:
        r, err_rb, norm_rb = crear_rebar_desde_curva_linea_con_ganchos(
            doc,
            el,
            bt,
            curva_rebar,
            marco_cara_uvn=marco_uvn,
            cara_paralela=cara_pp,
            eje_referencia_z_ganchos=z_hook,
        )
        if r is not None:
            fallback = True
    return r, err_rb, norm_rb, fallback


def _edge_jobs_inferior(el):
    """Lista de (curva_bruta, lado_etq, perp_len_mm, marco)."""
    marco = obtener_marco_coordenadas_cara_inferior(el)
    res_menor = extraer_curva_lado_menor_cara_inferior(el)
    res_mayor = extraer_curva_lado_mayor_cara_inferior(
        el,
        excluir_curva=res_menor[0] if res_menor is not None else None,
    )
    if res_menor is not None and res_mayor is not None:
        if _curvas_mismo_segmento(res_menor[0], res_mayor[0]):
            res_mayor = None
    jobs = []
    if res_menor is not None:
        try:
            perp = (
                float(res_mayor[0].Length) * 304.8
                if res_mayor is not None
                else float(res_menor[0].Length) * 304.8
            )
        except Exception:
            perp = None
        jobs.append((res_menor[0], u"menor", perp, marco))
    if res_mayor is not None:
        try:
            perp = (
                float(res_menor[0].Length) * 304.8
                if res_menor is not None
                else float(res_mayor[0].Length) * 304.8
            )
        except Exception:
            perp = None
        jobs.append((res_mayor[0], u"mayor", perp, marco))
    return jobs


def _edge_jobs_superior(el):
    marco = obtener_marco_coordenadas_cara_superior(el)
    res_menor = extraer_curva_lado_menor_cara_superior(el)
    res_mayor = extraer_curva_lado_mayor_cara_superior(
        el,
        excluir_curva=res_menor[0] if res_menor is not None else None,
    )
    if res_menor is not None and res_mayor is not None:
        if _curvas_mismo_segmento(res_menor[0], res_mayor[0]):
            res_mayor = None
    jobs = []
    if res_menor is not None:
        try:
            perp = (
                float(res_mayor[0].Length) * 304.8
                if res_mayor is not None
                else float(res_menor[0].Length) * 304.8
            )
        except Exception:
            perp = None
        jobs.append((res_menor[0], u"menor", perp, marco))
    if res_mayor is not None:
        try:
            perp = (
                float(res_menor[0].Length) * 304.8
                if res_menor is not None
                else float(res_mayor[0].Length) * 304.8
            )
        except Exception:
            perp = None
        jobs.append((res_mayor[0], u"mayor", perp, marco))
    return jobs


def _crear_malla_capa(
    doc,
    el,
    settings_group,
    ubicacion,
    edge_jobs,
    z0p,
    z1p,
    leg_fn,
    allow_u,
    d_stack_mm,
    avisos,
    conjunto_guid,
):
    """
    Crea luz menor + luz mayor. ``d_stack_mm`` = ø de la 1ª capa (menor) para
    separar el eje de la 2ª (mayor).
    """
    n_ok = 0
    if not edge_jobs:
        avisos.append(
            u"Id {0}: no se extrajo borde menor/mayor para {1}.".format(
                element_id_to_int(el.Id), ubicacion
            )
        )
        return n_ok

    for curva, lado_etq, perp_len_mm, marco_uvn in edge_jobs:
        dir_key = u"luz_menor" if lado_etq == u"menor" else u"luz_mayor"
        bt, sep_mm, d_mm_bar = _dir_cfg(settings_group, dir_key)
        if bt is None:
            avisos.append(
                u"Id {0} ({1}): falta RebarBarType.".format(
                    element_id_to_int(el.Id), dir_key
                )
            )
            continue
        ext_mm = _extremos_mm(d_mm_bar)
        curva_tratada, _co = aplicar_recubrimiento_inferior_completo_mm(
            curva, el, _REC_PLANTA_MM, ext_mm
        )
        if curva_tratada is None:
            avisos.append(
                u"Id {0} ({1}): curva nula tras recubrimiento.".format(
                    element_id_to_int(el.Id), dir_key
                )
            )
            continue
        cara_pp = None
        try:
            ev_par = evaluar_caras_paralelas_curva_mas_cercana(el, curva_tratada)
            if ev_par and ev_par.get(u"mejor"):
                cara_pp = ev_par[u"mejor"]
        except Exception:
            pass
        n_cara = marco_uvn[3] if marco_uvn is not None and len(marco_uvn) > 3 else None
        curva_rebar = offset_linea_eje_barra_desde_cara_inferior_mm(
            curva_tratada,
            n_cara,
            _REC_HORIZONTAL_MM,
            d_mm_bar,
        )
        # 2ª capa: separación de ejes = ø de la capa inferior (luz menor).
        if lado_etq == u"mayor":
            stack = float(d_stack_mm) if d_stack_mm and d_stack_mm > 1e-9 else d_mm_bar
            if stack > 1e-9:
                curva_rebar = offset_linea_adicional_hacia_interior_mm(
                    curva_rebar, n_cara, stack
                )
        z_hook = vector_reverso_cara_paralela_mas_cercana_a_barra(el, curva_rebar)
        if ubicacion == _ARMA_UBICACION_SUPERIOR and z_hook is None:
            z_hook = XYZ(0.0, 0.0, -1.0)
        leg_ft = None
        if allow_u:
            leg_ft = leg_fn(z0p, z1p, d_mm_bar, el)
        r, err_rb, norm_rb, fallback = _crear_rebar_u_o_fallback(
            doc,
            el,
            bt,
            curva_rebar,
            n_cara,
            leg_ft,
            d_mm_bar,
            marco_uvn,
            cara_pp,
            z_hook,
            allow_u,
        )
        if r is None:
            avisos.append(
                u"Id {0} ({1}): {2}".format(
                    element_id_to_int(el.Id),
                    dir_key,
                    err_rb or u"CreateFromCurves falló.",
                )
            )
            continue
        if fallback:
            avisos.append(
                u"Id {0} ({1}): U rechazada; barra con ganchos en el eje.".format(
                    element_id_to_int(el.Id), dir_key
                )
            )
        _aplicar_params_rebar(r, ubicacion, conjunto_guid)
        array_len_ft = _array_len_ft(
            el, curva_tratada, perp_len_mm, _REC_PLANTA_MM, lado_etq
        )
        ok_lay, err_lay = aplicar_layout_maximum_spacing_rebar(
            r, doc, sep_mm, array_len_ft
        )
        if not ok_lay and err_lay:
            avisos.append(
                u"Id {0} ({1}): layout: {2}".format(
                    element_id_to_int(el.Id), dir_key, err_lay
                )
            )
        n_ok += 1
    return n_ok


def _crear_laterales(doc, el, settings, d_mm_inf, avisos, conjunto_guid):
    from Autodesk.Revit.DB import Line

    lat = settings.get(u"lateral") or {}
    bt = lat.get(u"bar_type")
    if bt is None:
        avisos.append(u"Lateral: falta RebarBarType.")
        return 0
    d_mm_lat = lat.get(u"diameter_mm")
    if d_mm_lat is None:
        d_mm_lat = _d_mm(bt, 8.0)
    else:
        d_mm_lat = float(d_mm_lat)
    try:
        sep_mm = float(lat.get(u"spacing_mm") or 200.0)
    except Exception:
        sep_mm = 200.0

    off_planta = float(_REC_LATERAL_CARA_MM) + float(d_mm_inf or 0.0) + 0.5 * d_mm_lat
    recorte_ext = float(_REC_LATERAL_CARA_MM) + float(d_mm_inf or 0.0)

    try:
        fz_min, fz_max = rango_z_caras_laterales_o_bbox(el)
    except Exception:
        fz_min, fz_max = None, None
    if fz_min is None or fz_max is None:
        avisos.append(
            u"Id {0}: sin rango Z para laterales.".format(element_id_to_int(el.Id))
        )
        return 0

    try:
        lh = lineas_horizontales_perimetro_inferior_exterior(el)
    except Exception:
        lh = None
    if lh is None or not lh[0]:
        avisos.append(
            u"Id {0}: sin perímetro inferior para laterales.".format(
                element_id_to_int(el.Id)
            )
        )
        return 0
    lineas_borde = lh[0]

    marco_inf = None
    try:
        marco_inf = obtener_marco_coordenadas_cara_inferior(el)
    except Exception:
        pass
    n_inferior = None
    if marco_inf is not None and len(marco_inf) > 3:
        try:
            ni = marco_inf[3]
            if ni is not None and float(ni.GetLength()) > 1e-12:
                n_inferior = ni.Normalize()
        except Exception:
            n_inferior = None

    array_len_ft = longitud_array_lateral_altura_fundacion_menos_mm_ft(
        fz_min, fz_max, _DESCUENTO_ARRAY_LATERAL_MM
    )
    n_ok = 0
    for line_borde in lineas_borde:
        curva_tratada, _co = aplicar_recubrimiento_inferior_completo_mm(
            line_borde, el, off_planta, recorte_ext
        )
        if curva_tratada is None:
            continue
        n_horiz = normal_saliente_horizontal_paramento_para_barra_horizontal(
            curva_tratada, el
        )
        if n_horiz is None:
            continue
        if n_inferior is not None:
            curva_rebar = offset_linea_hacia_interior_desde_cara_inferior_mm(
                curva_tratada,
                n_inferior,
                _OFFSET_PRIMERA_LATERAL_MM,
            )
        else:
            z_fb = primera_cota_z_armadura_lateral_ft(
                fz_min, fz_max, _REC_LATERAL_CARA_MM, d_mm_lat
            )
            curva_rebar = linea_horizontal_cara_lateral_a_cota_z(curva_tratada, z_fb)
        if curva_rebar is None:
            continue
        z_hook = vector_reverso_cara_paralela_mas_cercana_a_barra(
            el,
            curva_rebar,
            excluir_caras_tapas_horizontales=True,
        )
        if z_hook is None:
            try:
                zh = n_horiz.Negate()
                if zh is not None and float(zh.GetLength()) > 1e-12:
                    z_hook = zh.Normalize()
            except Exception:
                z_hook = None
        n_inf_rev = None
        if n_inferior is not None:
            try:
                nr = n_inferior.Negate()
                if nr is not None and float(nr.GetLength()) > 1e-12:
                    n_inf_rev = nr.Normalize()
            except Exception:
                n_inf_rev = None
        norm_pri = [n_inf_rev] if n_inf_rev is not None else None

        r, err_rb, _norm_rb = None, None, None
        hook_lat_mm = largo_gancho_u_tabla_mm(d_mm_lat)
        leg_ft_lat = None
        if hook_lat_mm is not None:
            try:
                d_round = float(int(round(float(d_mm_lat)))) if d_mm_lat else 0.0
                eje_lat_mm = max(float(hook_lat_mm) - 0.5 * d_round, 40.0)
            except Exception:
                eje_lat_mm = float(hook_lat_mm)
            try:
                from Autodesk.Revit.DB import UnitUtils, UnitTypeId

                leg_ft_lat = UnitUtils.ConvertToInternalUnits(
                    eje_lat_mm, UnitTypeId.Millimeters
                )
            except Exception:
                leg_ft_lat = eje_lat_mm / 304.8
        if leg_ft_lat is not None:
            curva_rebar_lat = curva_rebar
            try:
                from Autodesk.Revit.DB import UnitUtils, UnitTypeId

                half_d_ft = UnitUtils.ConvertToInternalUnits(
                    float(d_mm_lat) / 2.0, UnitTypeId.Millimeters
                )
            except Exception:
                half_d_ft = float(d_mm_lat) / 2.0 / 304.8
            try:
                p0 = curva_rebar.GetEndPoint(0)
                p1 = curva_rebar.GetEndPoint(1)
                tang = p1 - p0
                tlen = float(tang.GetLength())
                if tlen > 2.0 * half_d_ft + 1e-6:
                    tu = tang.Multiply(1.0 / tlen)
                    curva_rebar_lat = Line.CreateBound(
                        p0 + tu.Multiply(half_d_ft),
                        p1 - tu.Multiply(half_d_ft),
                    )
            except Exception:
                curva_rebar_lat = curva_rebar
            poli_u_lat = construir_polilinea_u_fundacion_desde_eje_horizontal(
                curva_rebar_lat,
                n_horiz,
                leg_ft_lat,
                d_mm_lat,
                acortar_eje_central_para_cota_revit=False,
            )
            if poli_u_lat is not None:
                r, err_rb, _norm_rb = crear_rebar_u_shape_desde_eje_rebar_shape_nombrado(
                    doc,
                    el,
                    bt,
                    poli_u_lat,
                    shape_nombre=REBAR_SHAPE_NOMBRE_DEFECTO,
                    marco_cara_uvn=marco_inf,
                    cara_paralela=None,
                    eje_referencia_z_ganchos=z_hook,
                )
                if r is None:
                    r, err_rb, _norm_rb = crear_rebar_polilinea_u_malla_inf_sup_curve_loop(
                        doc,
                        el,
                        bt,
                        poli_u_lat,
                        poli_u_lat[1],
                        marco_cara_uvn=marco_inf,
                        cara_paralela=None,
                        eje_referencia_z_ganchos=z_hook,
                    )
        if r is None:
            r, err_rb, _norm_rb = crear_rebar_polilinea_recta_sin_ganchos(
                doc,
                el,
                bt,
                [curva_rebar],
                curva_rebar,
                marco_cara_uvn=marco_inf,
                cara_paralela=None,
                eje_referencia_z_ganchos=z_hook,
                normales_prioridad=norm_pri,
            )
        if r is None:
            avisos.append(
                u"Id {0} (lateral): {1}".format(
                    element_id_to_int(el.Id),
                    err_rb or u"CreateFromCurves falló.",
                )
            )
            continue
        _aplicar_params_rebar(r, _ARMA_UBICACION_LATERAL, conjunto_guid)
        ok_lay, err_lay = aplicar_layout_maximum_spacing_rebar(
            r, doc, sep_mm, array_len_ft
        )
        if not ok_lay and err_lay:
            avisos.append(
                u"Id {0} (lateral): layout: {1}".format(
                    element_id_to_int(el.Id), err_lay
                )
            )
        n_ok += 1
    return n_ok


# ---------------------------------------------------------------------------
# Vistas Detail (Section Filter de la vista origen)
# ---------------------------------------------------------------------------


def _mm_to_ft(mm):
    try:
        return float(mm) / 304.8
    except Exception:
        return 0.0


def _mark_fundacion(elem):
    if elem is None:
        return u"?"
    for pname in (u"Mark", u"Marca", u"Numeracion", u"Número"):
        try:
            p = elem.LookupParameter(pname)
            if p is not None and p.HasValue:
                s = p.AsString()
                if s and _as_unicode(s).strip():
                    return _as_unicode(s).strip()
        except Exception:
            pass
    try:
        p = elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        if p is not None and p.HasValue:
            s = p.AsString()
            if s and _as_unicode(s).strip():
                return _as_unicode(s).strip()
    except Exception:
        pass
    try:
        return str(int(elem.Id.IntegerValue))
    except Exception:
        return u"?"


def _unique_view_name(doc, base_name):
    base = _as_unicode(base_name).strip() or u"DET. FUND"
    used = set()
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            if v is not None and (not v.IsTemplate) and v.Name:
                used.add(_as_unicode(v.Name).strip())
        except Exception:
            continue
    if base not in used:
        return base
    for i in range(2, 200):
        cand = u"{0} ({1})".format(base, i)
        if cand not in used:
            return cand
    return u"{0} (x)".format(base)


def _copy_section_filter_to_view(source_view, target_view):
    if source_view is None or target_view is None:
        return
    try:
        p_src = source_view.LookupParameter(_PARAM_SECTION_FILTER)
        p_dst = target_view.LookupParameter(_PARAM_SECTION_FILTER)
        if p_src is None or p_dst is None or p_dst.IsReadOnly:
            return
        if not p_src.HasValue:
            return
        st = p_src.StorageType
        if st == StorageType.String:
            p_dst.Set(p_src.AsString() or u"")
        elif st == StorageType.ElementId:
            p_dst.Set(p_src.AsElementId())
        else:
            try:
                vs = p_src.AsValueString()
                if vs:
                    p_dst.SetValueString(vs)
            except Exception:
                pass
    except Exception:
        pass


def _iter_rebars_by_ubicacion(doc, ubicacion_keep):
    """Yield Rebar whose Armadura_Ubicacion matches ``ubicacion_keep``."""
    if doc is None or not ubicacion_keep:
        return
    try:
        from Autodesk.Revit.DB.Structure import Rebar
    except Exception:
        return
    want = _as_unicode(ubicacion_keep).strip()
    for rb in FilteredElementCollector(doc).OfClass(Rebar).ToElements():
        if rb is None:
            continue
        val = u""
        try:
            p = rb.LookupParameter(_ARMA_UBICACION_PARAM)
            if p is not None and p.HasValue:
                val = p.AsString() or u""
        except Exception:
            val = u""
        if _as_unicode(val).strip() == want:
            yield rb


def _hide_rebars_not_ubicacion(view, doc, ubicacion_keep):
    if view is None or doc is None or not ubicacion_keep:
        return
    try:
        from Autodesk.Revit.DB.Structure import Rebar
        from System.Collections.Generic import List

        ids = List[ElementId]()
        want = _as_unicode(ubicacion_keep).strip()
        for rb in FilteredElementCollector(doc).OfClass(Rebar).ToElements():
            if rb is None:
                continue
            val = u""
            try:
                p = rb.LookupParameter(_ARMA_UBICACION_PARAM)
                if p is not None and p.HasValue:
                    val = p.AsString() or u""
            except Exception:
                val = u""
            if _as_unicode(val).strip() != want:
                ids.Add(rb.Id)
        if ids.Count > 0:
            view.HideElements(ids)
    except Exception:
        pass


def _apply_unobscured_rebars_en_vista(view, doc, ubicacion_keep=None):
    """
    Activa View Unobscured (+ sólido) en los Rebar visibles de ``view``.
    Si ``ubicacion_keep`` está definido, solo aplica a esa ubicación.
    """
    if view is None or doc is None:
        return 0
    try:
        if ubicacion_keep:
            rebars = list(_iter_rebars_by_ubicacion(doc, ubicacion_keep))
        else:
            from Autodesk.Revit.DB.Structure import Rebar

            rebars = list(FilteredElementCollector(doc).OfClass(Rebar).ToElements())
        if not rebars:
            return 0
        from bimtools_rebar_3d_visibility import apply_rebar_unobscured_in_view

        return int(apply_rebar_unobscured_in_view(doc, rebars, view) or 0)
    except Exception:
        # Fallback directo si el helper no está disponible.
        n = 0
        try:
            from Autodesk.Revit.DB.Structure import Rebar

            if ubicacion_keep:
                rebars = list(_iter_rebars_by_ubicacion(doc, ubicacion_keep))
            else:
                rebars = list(FilteredElementCollector(doc).OfClass(Rebar).ToElements())
            for rb in rebars:
                if rb is None:
                    continue
                try:
                    rb.SetUnobscuredInView(view, True)
                    n += 1
                except Exception:
                    pass
                try:
                    rb.SetSolidInView(view, True)
                except Exception:
                    pass
        except Exception:
            pass
        return n


def _bbox_detail_fundacion_mirando_abajo(elem, margin_mm=_DETALLE_MARGEN_MM):
    """
    BoundingBoxXYZ para ``ViewSection.CreateDetail``: corte horizontal mirando abajo
    (planta de la fundación).
    """
    from Autodesk.Revit.DB import BoundingBoxXYZ

    if elem is None:
        return None, u"Elemento nulo."
    try:
        bb = elem.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None:
        return None, u"Sin BoundingBox de la fundación."

    try:
        cx = 0.5 * (float(bb.Min.X) + float(bb.Max.X))
        cy = 0.5 * (float(bb.Min.Y) + float(bb.Max.Y))
        cz = 0.5 * (float(bb.Min.Z) + float(bb.Max.Z))
    except Exception:
        return None, u"BoundingBox inválido."

    # Triedro: mirada hacia abajo (mismo criterio que detalle extremo de muro).
    bz = XYZ(0.0, 0.0, -1.0)
    bx = XYZ(1.0, 0.0, 0.0)
    by = bz.CrossProduct(bx)
    try:
        by = by.Normalize()
    except Exception:
        by = XYZ(0.0, 1.0, 0.0)

    tr = Transform.Identity
    tr.Origin = XYZ(cx, cy, cz)
    tr.BasisX = bx
    tr.BasisY = by
    tr.BasisZ = bz

    corners = (
        XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
        XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
    )
    xs, ys, zs = [], [], []
    ox = tr.Origin
    for p in corners:
        d = p - ox
        xs.append(float(d.DotProduct(bx)))
        ys.append(float(d.DotProduct(by)))
        zs.append(float(d.DotProduct(bz)))
    if not xs:
        return None, u"No se pudieron proyectar vértices."

    m = _mm_to_ft(margin_mm)
    depth = _mm_to_ft(_DETALLE_PROFUNDIDAD_MM)
    box = BoundingBoxXYZ()
    box.Transform = tr
    box.Min = XYZ(min(xs) - m, min(ys) - m, min(zs) - 0.5 * depth)
    box.Max = XYZ(max(xs) + m, max(ys) + m, max(zs) + 0.5 * depth)
    return box, None


def _resolve_detail_vft_and_section_filter(doc, source_view):
    """
    Returns:
        (vft, section_filter_text, err)
    """
    try:
        from seccion_detalle_extremo_muro import (
            resolver_view_family_type_detail_desde_vista as _resolver_detail_vft,
        )

        return _resolver_detail_vft(doc, source_view)
    except Exception:
        try:
            from seccion_detalle_extremo_muro import (
                find_view_family_type_detail_by_name,
                leer_section_filter_texto,
            )

            sf, err = leer_section_filter_texto(doc, source_view)
            if sf is None:
                return None, None, err
            vft, err2 = find_view_family_type_detail_by_name(doc, sf)
            return vft, sf, err2
        except Exception as ex:
            return None, None, _as_unicode(ex)


def crear_vistas_detalle_mallas_fundacion(
    doc,
    foundation,
    source_view,
    inf_on,
    sup_on,
    avisos=None,
):
    """
    Crea Detail View(s) mirando la fundación en planta:

    - Solo malla inferior → 1 vista.
    - Inferior + superior → 1 vista por malla.
    - Solo superior → 1 vista (superior).

    El nombre incluye el ``Section Filter`` de ``source_view`` y el ViewFamilyType
    Detail se resuelve con ese mismo filtro (como detalle de extremo de muro).

    Debe llamarse **dentro** de una Transaction abierta.
    Returns: lista de vistas creadas.
    """
    if avisos is None:
        avisos = []
    created = []
    if doc is None or foundation is None:
        return created
    if not inf_on and not sup_on:
        return created
    if source_view is None:
        avisos.append(
            u"Vistas Detail: no hay vista origen para leer «Section Filter»."
        )
        return created

    vft, sf_text, err = _resolve_detail_vft_and_section_filter(doc, source_view)
    if vft is None:
        avisos.append(
            u"Vistas Detail no creadas: {0}".format(
                err or u"sin ViewFamilyType Detail para el Section Filter."
            )
        )
        return created
    if not sf_text:
        avisos.append(u"Vistas Detail: Section Filter vacío en la vista activa.")
        return created

    box, err_box = _bbox_detail_fundacion_mirando_abajo(foundation)
    if box is None:
        avisos.append(
            u"Vistas Detail no creadas: {0}".format(err_box or u"bbox inválido.")
        )
        return created

    jobs = []
    if inf_on and sup_on:
        jobs = [
            (_ARMA_UBICACION_INFERIOR, u"ARM. INFERIOR"),
            (_ARMA_UBICACION_SUPERIOR, u"ARM. SUPERIOR"),
        ]
    elif inf_on:
        jobs = [(_ARMA_UBICACION_INFERIOR, u"ARM. INFERIOR")]
    elif sup_on:
        jobs = [(_ARMA_UBICACION_SUPERIOR, u"ARM. SUPERIOR")]

    mark = _mark_fundacion(foundation)
    for ubicacion, tag in jobs:
        try:
            vs = ViewSection.CreateDetail(doc, vft.Id, box)
        except Exception as ex:
            avisos.append(
                u"CreateDetail ({0}) falló: {1}".format(tag, _as_unicode(ex))
            )
            continue
        if vs is None:
            avisos.append(u"CreateDetail ({0}) devolvió None.".format(tag))
            continue
        try:
            vs.CropBoxActive = True
        except Exception:
            pass
        try:
            vs.CropBoxVisible = False
        except Exception:
            pass
        try:
            vs.Scale = 25
        except Exception:
            pass

        base_name = u"{0}_DET. FUND. F{1}_{2}".format(
            _as_unicode(sf_text).strip(), mark, tag
        )
        try:
            vs.Name = _unique_view_name(doc, base_name)
        except Exception:
            pass

        _copy_section_filter_to_view(source_view, vs)
        _hide_rebars_not_ubicacion(vs, doc, ubicacion)
        _apply_unobscured_rebars_en_vista(vs, doc, ubicacion)

        # Visibilidad categoría Rebar
        try:
            from Autodesk.Revit.DB import BuiltInCategory

            cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rebar)
            if cat is not None:
                vs.SetCategoryHidden(cat.Id, False)
        except Exception:
            pass

        created.append(vs)
        try:
            avisos.append(u"Vista Detail: {0}".format(_as_unicode(vs.Name)))
        except Exception:
            avisos.append(u"Vista Detail creada ({0}).".format(tag))
    return created


def crear_armadura_fundacion_aislada(
    doc,
    foundation,
    settings,
    transaction_name=None,
    source_view=None,
    uidoc=None,
):
    """
    Coloca armadura según ``settings`` del sketch y genera Detail View(s)
    según mallas activas (Section Filter de ``source_view``).

    settings::
        inferior/superior: {enabled, luz_mayor: {bar_type, spacing_mm, diameter_mm}, luz_menor: {...}}
        lateral: {enabled, bar_type, spacing_mm, diameter_mm}

    Returns dict: ok, n_inferior, n_superior, n_lateral, avisos, message, conjunto_guid, views
    """
    result = {
        u"ok": False,
        u"n_inferior": 0,
        u"n_superior": 0,
        u"n_lateral": 0,
        u"avisos": [],
        u"message": u"",
        u"conjunto_guid": None,
        u"views": [],
    }
    if doc is None or foundation is None or not settings:
        result[u"message"] = u"Datos incompletos."
        return result

    inf = settings.get(u"inferior") or {}
    sup = settings.get(u"superior") or {}
    lat = settings.get(u"lateral") or {}
    inf_on = bool(inf.get(u"enabled"))
    sup_on = bool(sup.get(u"enabled"))
    lat_on = bool(lat.get(u"enabled"))
    if not (inf_on or sup_on or lat_on):
        result[u"message"] = u"Ningún grupo activo."
        return result

    try:
        from geometria_fundacion_cara_inferior import clear_face_cache

        clear_face_cache()
    except Exception:
        pass

    conjunto_guid = None
    if iniciar_armadura_conjunto_guid_ejecucion is not None:
        try:
            conjunto_guid = iniciar_armadura_conjunto_guid_ejecucion()
        except Exception:
            conjunto_guid = None
    result[u"conjunto_guid"] = conjunto_guid

    avisos = result[u"avisos"]
    txn_name = transaction_name or _TXN_NAME
    t = Transaction(doc, txn_name)
    t.Start()
    try:
        z0p, z1p = None, None
        try:
            z0p, z1p = rango_z_caras_laterales_o_bbox(foundation)
        except Exception:
            pass

        d_inf_stack = 0.0
        if inf_on:
            _bt_m, _sep_m, d_inf_stack = _dir_cfg(inf, u"luz_menor")
            jobs = _edge_jobs_inferior(foundation)
            result[u"n_inferior"] = _crear_malla_capa(
                doc,
                foundation,
                inf,
                _ARMA_UBICACION_INFERIOR,
                jobs,
                z0p,
                z1p,
                lambda z0, z1, d, el: _leg_ft_pata_u_inferior(
                    z0, z1, d, sup_on, elem=el
                ),
                True,
                d_inf_stack,
                avisos,
                conjunto_guid,
            )

        if sup_on:
            _bt_m, _sep_m, d_sup_stack = _dir_cfg(sup, u"luz_menor")
            jobs_s = _edge_jobs_superior(foundation)
            allow_u_sup = inf_on
            result[u"n_superior"] = _crear_malla_capa(
                doc,
                foundation,
                sup,
                _ARMA_UBICACION_SUPERIOR,
                jobs_s,
                z0p,
                z1p,
                lambda z0, z1, d, el: _leg_ft_pata_u_superior(z0, z1, d, elem=el),
                allow_u_sup,
                d_sup_stack,
                avisos,
                conjunto_guid,
            )

        if lat_on:
            d_inf_for_lat = d_inf_stack if inf_on else 0.0
            if not inf_on:
                d_inf_for_lat = 0.0
            result[u"n_lateral"] = _crear_laterales(
                doc,
                foundation,
                settings,
                d_inf_for_lat,
                avisos,
                conjunto_guid,
            )

        # Detail Views: solo cuando hay malla(s) inferior/superior.
        if inf_on or sup_on:
            try:
                views = crear_vistas_detalle_mallas_fundacion(
                    doc,
                    foundation,
                    source_view,
                    inf_on,
                    sup_on,
                    avisos=avisos,
                )
                result[u"views"] = list(views or [])
            except Exception as ex_v:
                avisos.append(
                    u"Vistas Detail no creadas: {0}".format(_as_unicode(ex_v))
                )

        t.Commit()
        result[u"ok"] = True
        parts = []
        if inf_on:
            parts.append(u"Inferior: {0} conjunto(s)".format(result[u"n_inferior"]))
        if sup_on:
            parts.append(u"Superior: {0} conjunto(s)".format(result[u"n_superior"]))
        if lat_on:
            parts.append(u"Lateral: {0} barra(s)".format(result[u"n_lateral"]))
        if result[u"views"]:
            parts.append(u"Vistas Detail: {0}".format(len(result[u"views"])))
        result[u"message"] = u" · ".join(parts) if parts else u"Sin barras."
        if avisos:
            result[u"message"] += u"\n\nAvisos:\n" + u"\n".join(avisos[:12])
            if len(avisos) > 12:
                result[u"message"] += u"\n… (+{0})".format(len(avisos) - 12)
        # Las vistas Detail se crean pero no se activan (se mantiene la vista activa).
    except Exception as ex:
        try:
            if t.HasStarted():
                t.RollBack()
        except Exception:
            pass
        result[u"ok"] = False
        result[u"message"] = u"Error al crear armadura:\n{0}".format(_as_unicode(ex))
    return result
