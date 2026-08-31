# -*- coding: utf-8 -*-
"""
Colocación Revit — Remate Mallas (pie fundación / superior coronamiento).

Modos:
- Pie + fundación: geometría muro, Z fondo fund., host WallFoundation.
- Superior: geometría muro, Z tope muro, host muro (coronamiento).

Revit 2024+ | IronPython / pyRevit
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    ElementId,
    JoinGeometryUtils,
    Transaction,
    TransactionGroup,
    UnitTypeId,
    UnitUtils,
    Wall,
    WallFoundation,
)

import armado_muros_coronamiento as cor

from armado_muros_rebar_params import (
    finalizar_armadura_conjunto_guid_ejecucion,
    finalizar_armadura_eje_ejecucion,
    iniciar_armadura_conjunto_guid_ejecucion,
    iniciar_armadura_eje_ejecucion,
)

from barras_retorno_malla_geom import (
    COVER_FUND_BOT_MM,
    COVER_WALL_BOT_MM,
    END_PATA,
    MAX_BARRA_COMERCIAL_MM,
    clamp_diam_mm,
    clamp_n_bars,
    cover_axis_offset_mm_for_layer,
    normalize_concrete_grade,
    normalize_end_condition,
    stagger_cuts_for_layer,
    to_dividir_lap_mode,
    traslape_mm_from_diam,
)

from barras_retorno_malla_place import (
    _active_view,
    _as_unicode,
    _bar_type,
    _build_rebar_groups_multicapas,
    _create_retorno_rebar,
    _cuts_ui_to_centerline_mm,
    _ensure_dividir_rebar_punto,
    _stamp_layer,
    _tag_retorno_rebars,
    _wall_foundation_for_wall,
    resolve_foundation,
    main_bar_length_for_wall_mm,
    wall_meta_for_ui,
)

from retorno_malla_fundacion_geom import (
    COVER_SUPERIOR_MM,
    MODE_INFERIOR_FUND,
    MODE_SUPERIOR,
    normalize_mode,
)


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _txn_names(mode):
    m = normalize_mode(mode)
    if m == MODE_SUPERIOR:
        return (
            u"Arainco: Remate Mallas superior",
            u"Arainco: Remate Mallas superior (capa)",
        )
    return (
        u"Arainco: Remate Mallas",
        u"Arainco: Remate Mallas (capa)",
    )


def _tag_extremo_for_mode(mode):
    if normalize_mode(mode) == MODE_SUPERIOR:
        return u"cor_sup"
    return u"cor_pie"


def _pata_leg_sign_for_mode(mode):
    return -1.0 if normalize_mode(mode) == MODE_SUPERIOR else 1.0


def _z_bar_layer_ft_for_mode(wall, fund, layers, layer_index, mode):
    m = normalize_mode(mode)
    if m == MODE_SUPERIOR:
        base_cover = COVER_SUPERIOR_MM
        offset_mm = cover_axis_offset_mm_for_layer(
            layers, layer_index, base_cover_mm=base_cover
        )
        try:
            _, z_top = cor._wall_z_bounds_ft(wall)
        except Exception:
            return None
        if z_top is None:
            return None
        return float(z_top) - float(_mm_to_internal(offset_mm))

    base_cover = COVER_FUND_BOT_MM if fund is not None else COVER_WALL_BOT_MM
    offset_mm = cover_axis_offset_mm_for_layer(
        layers, layer_index, base_cover_mm=base_cover
    )
    if fund is not None:
        try:
            z_bot = cor._z_cara_inferior_fundacion_ft(fund)
        except Exception:
            z_bot = None
        if z_bot is None:
            z0, _z1 = cor._element_z_bounds_ft(fund)
            z_bot = z0
        if z_bot is None:
            return None
        return float(z_bot) + float(_mm_to_internal(offset_mm))
    try:
        z_wall_bot, _ = cor._wall_z_bounds_ft(wall)
    except Exception:
        return None
    return float(z_wall_bot) + float(_mm_to_internal(offset_mm))


def _host_for_mode(wall, fund, mode):
    """
    Host estructural del Rebar.

    - Superior: muro.
    - Pie / fundación detectada: la fundación (WallFoundation preferida).
    """
    if normalize_mode(mode) == MODE_SUPERIOR:
        return wall
    return fund


def _resolve_fundacion_host(doc, wall):
    """
    Fundación para hospedar Rebar en modo pie.

    1) ``WallFoundation`` con ``WallId`` = muro
    2) Fundación estructural unida (join), priorizando ``WallFoundation``
    """
    wf = _wall_foundation_for_wall(doc, wall)
    if wf is not None:
        return wf
    fund = resolve_foundation(doc, wall)
    if fund is None:
        return None
    if isinstance(fund, WallFoundation):
        return fund
    # Si el join devolvió losa/otra fundación, intenta de nuevo WallFoundation
    wf2 = _wall_foundation_for_wall(doc, wall)
    return wf2 if wf2 is not None else fund


def _joined_ids_safe(doc, element):
    try:
        from bimtools_joined_geometry import get_joined_element_ids

        return list(get_joined_element_ids(doc, element) or [])
    except Exception:
        pass
    try:
        return list(JoinGeometryUtils.GetJoinedElements(doc, element) or [])
    except Exception:
        try:
            return list(JoinGeometryUtils.GetJoinedElements(doc, element.Id) or [])
        except Exception:
            return []


def _unjoin_host(doc, host, other_ids):
    for oid in other_ids or []:
        try:
            oth = doc.GetElement(oid)
        except Exception:
            oth = None
        if oth is None:
            continue
        try:
            JoinGeometryUtils.UnjoinGeometry(doc, host, oth)
        except Exception:
            pass


def _rejoin_host(doc, host, other_ids):
    for oid in other_ids or []:
        try:
            oth = doc.GetElement(oid)
        except Exception:
            oth = None
        if oth is None:
            continue
        try:
            if JoinGeometryUtils.AreElementsJoined(doc, host.Id, oid):
                continue
        except Exception:
            try:
                if JoinGeometryUtils.AreElementsJoined(doc, host, oth):
                    continue
            except Exception:
                pass
        try:
            JoinGeometryUtils.JoinGeometry(doc, host, oth)
        except Exception:
            pass


def _rebar_host_id(rebar):
    if rebar is None:
        return None
    try:
        return rebar.GetHostId()
    except Exception:
        return None


def _create_retorno_rebar_hosted(
    doc,
    wall,
    host,
    n_bars,
    bar_type,
    z_bar_ft,
    diam_mm,
    end_a,
    end_b,
    grade,
    pata_leg_sign=1.0,
):
    """
    Crea el Rebar hospedado en ``host``.

    Con ``WallFoundation``, desune geometría temporalmente (mismo criterio que
    enfierrado de fundación corrida) para que Revit no reasigne el host al muro.
    """
    joined_ids = []
    use_unjoin = host is not None and isinstance(host, WallFoundation)
    if use_unjoin:
        joined_ids = _joined_ids_safe(doc, host)
        if joined_ids:
            _unjoin_host(doc, host, joined_ids)
            try:
                doc.Regenerate()
            except Exception:
                pass
    try:
        rb, n_layout, err = _create_retorno_rebar(
            doc,
            wall,
            host,
            n_bars,
            bar_type,
            z_bar_ft,
            diam_mm,
            end_a,
            end_b,
            grade,
            pata_leg_sign=pata_leg_sign,
        )
        if rb is not None and host is not None:
            hid = _rebar_host_id(rb)
            try:
                want = host.Id
            except Exception:
                want = None
            if hid is not None and want is not None and hid != want:
                try:
                    doc.Delete(rb.Id)
                except Exception:
                    pass
                return (
                    None,
                    0,
                    u"El Rebar no quedó hospedado en la fundación detectada.",
                )
        return rb, n_layout, err
    finally:
        if use_unjoin and joined_ids:
            _rejoin_host(doc, host, joined_ids)


def place_retorno_malla_fundacion_wall(
    doc,
    uidoc,
    wall,
    layers,
    cuts_ref_mm=None,
    lap_mode_ui=None,
    concrete_grade=None,
    end_a=None,
    end_b=None,
    mode=None,
):
    """
    Coloca barras de retorno por capa según modo (pie/fundación o superior).

    Returns:
        dict: ok, messages, rebar_ids, n_layers, main_mm, exceeds_12m, has_fund, mode
    """
    op_mode = normalize_mode(mode)
    txn_group, txn_create = _txn_names(op_mode)
    result = {
        u"ok": False,
        u"messages": [],
        u"rebar_ids": [],
        u"n_layers": 0,
        u"main_mm": 0.0,
        u"exceeds_12m": False,
        u"has_fund": False,
        u"n_tags": 0,
        u"n_tags_fail": 0,
        u"mode": op_mode,
    }
    if doc is None or wall is None or not isinstance(wall, Wall):
        result[u"messages"].append(u"Muro no válido.")
        return result

    layers = list(layers or [{u"n_bars": 2, u"diam_mm": 10}])
    if not layers:
        result[u"messages"].append(u"Sin capas configuradas.")
        return result

    n_capas_cfg = len(layers)

    fund = _resolve_fundacion_host(doc, wall)
    if op_mode == MODE_INFERIOR_FUND and fund is None:
        try:
            wid = int(wall.Id.IntegerValue)
        except Exception:
            wid = u"?"
        result[u"messages"].append(
            u"Muro {0}: sin fundación corrida unida (WallFoundation).".format(wid)
        )
        return result

    host = _host_for_mode(wall, fund, op_mode)
    if host is None:
        result[u"messages"].append(u"Host no válido para el modo seleccionado.")
        return result
    if op_mode == MODE_INFERIOR_FUND and not isinstance(host, WallFoundation):
        # Preferir siempre WallFoundation cuando exista para el muro
        wf = _wall_foundation_for_wall(doc, wall)
        if wf is not None:
            host = wf
            fund = wf

    if fund is not None:
        result[u"has_fund"] = True

    grade = normalize_concrete_grade(concrete_grade)
    ea = normalize_end_condition(end_a if end_a is not None else END_PATA)
    eb = normalize_end_condition(end_b if end_b is not None else END_PATA)
    pata_sign = _pata_leg_sign_for_mode(op_mode)
    tag_ext = _tag_extremo_for_mode(op_mode)

    try:
        wall_id_int = int(wall.Id.IntegerValue)
    except Exception:
        wall_id_int = 0

    main_mm = main_bar_length_for_wall_mm(doc, wall)
    result[u"main_mm"] = float(main_mm)
    result[u"exceeds_12m"] = bool(main_mm > MAX_BARRA_COMERCIAL_MM)
    cuts_ref = list(cuts_ref_mm or [])
    lap_mode = to_dividir_lap_mode(lap_mode_ui)

    divide_fn = None
    if cuts_ref:
        ok_imp, err_imp, divide_fn = _ensure_dividir_rebar_punto()
        if not ok_imp or divide_fn is None:
            result[u"messages"].append(
                u"Traslape (56) no disponible: {0}".format(err_imp or u"import")
            )
            cuts_ref = []

    iniciar_armadura_conjunto_guid_ejecucion()
    tg = None
    tg_started = False
    view = _active_view(uidoc)
    tag_meta = []
    layer_segment_ids = []
    try:
        try:
            iniciar_armadura_eje_ejecucion(uidoc=uidoc, view=view)
        except Exception:
            pass

        tg = TransactionGroup(doc, txn_group)
        tg.Start()
        tg_started = True

        created = []
        for li, ly in enumerate(layers):
            n_bars = clamp_n_bars(ly.get(u"n_bars", 2))
            diam_mm = clamp_diam_mm(ly.get(u"diam_mm", 10))
            bt = _bar_type(doc, diam_mm)
            if bt is None:
                result[u"messages"].append(
                    u"Capa {0}: no hay RebarBarType Ø{1}.".format(li + 1, diam_mm)
                )
                continue
            z_bar = _z_bar_layer_ft_for_mode(wall, fund, layers, li, op_mode)
            if z_bar is None:
                result[u"messages"].append(
                    u"Capa {0}: no se pudo calcular elevación Z.".format(li + 1)
                )
                continue

            rb = None
            t = Transaction(doc, txn_create)
            t.Start()
            try:
                rb, n_layout, err = _create_retorno_rebar_hosted(
                    doc,
                    wall,
                    host,
                    n_bars,
                    bt,
                    z_bar,
                    diam_mm,
                    ea,
                    eb,
                    grade,
                    pata_leg_sign=pata_sign,
                )
                if rb is not None:
                    _stamp_layer(rb, li, n_capas_cfg)
                    t.Commit()
                else:
                    t.RollBack()
                    result[u"messages"].append(
                        u"Capa {0}: {1}".format(li + 1, err or u"error al crear")
                    )
                    continue
            except Exception as ex_c:
                try:
                    t.RollBack()
                except Exception:
                    pass
                result[u"messages"].append(
                    u"Capa {0}: {1}".format(li + 1, _as_unicode(ex_c))
                )
                continue

            lap_mm = traslape_mm_from_diam(diam_mm, grade)
            layer_cuts = stagger_cuts_for_layer(cuts_ref, li, main_mm, lap_mm)
            final_ids = []

            if layer_cuts and divide_fn is not None:
                cuts_cl = _cuts_ui_to_centerline_mm(
                    layer_cuts, diam_mm, ea, grade, bar_type=bt
                )
                try:
                    ok_div, msg_div, ids_new, _meta = divide_fn(
                        doc,
                        rb,
                        cuts_cl,
                        concrete_grade=grade,
                        view=view,
                        lap_mode=lap_mode,
                        place_lap_dims=True,
                        lap_dim_prefer_above=True,
                    )
                except TypeError:
                    try:
                        ok_div, msg_div, ids_new, _meta = divide_fn(
                            doc,
                            rb,
                            cuts_cl,
                            concrete_grade=grade,
                            view=view,
                            lap_mode=lap_mode,
                            place_lap_dims=True,
                        )
                    except TypeError:
                        try:
                            ok_div, msg_div, ids_new, _meta = divide_fn(
                                doc,
                                rb,
                                cuts_cl,
                                concrete_grade=grade,
                                view=view,
                                lap_mode=lap_mode,
                            )
                        except Exception as ex_d:
                            ok_div = False
                            msg_div = _as_unicode(ex_d)
                            ids_new = []
                    except Exception as ex_d:
                        ok_div = False
                        msg_div = _as_unicode(ex_d)
                        ids_new = []
                except Exception as ex_d:
                    ok_div = False
                    msg_div = _as_unicode(ex_d)
                    ids_new = []
                if ok_div and ids_new:
                    for iv in ids_new:
                        try:
                            final_ids.append(
                                int(iv.IntegerValue if hasattr(iv, "IntegerValue") else iv)
                            )
                        except Exception:
                            try:
                                final_ids.append(int(iv))
                            except Exception:
                                pass
                    for iv in final_ids:
                        try:
                            el = doc.GetElement(ElementId(int(iv)))
                            _stamp_layer(el, li, n_capas_cfg)
                        except Exception:
                            pass
                else:
                    result[u"messages"].append(
                        u"Capa {0}: empalme — {1}".format(
                            li + 1, msg_div or u"falló divide"
                        )
                    )
                    try:
                        final_ids.append(int(rb.Id.IntegerValue))
                    except Exception:
                        pass
            else:
                try:
                    final_ids.append(int(rb.Id.IntegerValue))
                except Exception:
                    pass

            for rid in final_ids:
                try:
                    tag_meta.append(
                        {
                            u"rebar_id": ElementId(int(rid)),
                            u"layer_index": int(li),
                            u"wid": wall_id_int,
                            u"extremo": tag_ext,
                            u"zs": float(z_bar),
                            u"span_seg": 0.01,
                        }
                    )
                except Exception:
                    pass

            created.extend(final_ids)
            layer_segment_ids.append(list(final_ids))
            result[u"n_layers"] += 1

        rebar_groups = _build_rebar_groups_multicapas(layer_segment_ids)
        host_lab = u"muro"
        if op_mode != MODE_SUPERIOR:
            if isinstance(host, WallFoundation):
                host_lab = u"WallFoundation"
            else:
                host_lab = u"fundación"
        result[u"rebar_ids"] = created
        if created:
            result[u"ok"] = True
            result[u"messages"].append(
                u"Creadas {0} barra(s)/set(s) en {1} capa(s) · host {2}.".format(
                    len(created), result[u"n_layers"], host_lab
                )
            )
            try:
                tag_res = _tag_retorno_rebars(
                    doc,
                    uidoc,
                    created,
                    tag_meta=tag_meta,
                    rebar_groups=rebar_groups,
                    n_capas=n_capas_cfg,
                )
                result[u"n_tags"] = int(tag_res.get(u"n_ok", 0) or 0)
                result[u"n_tags_fail"] = int(tag_res.get(u"n_fail", 0) or 0)
                if result[u"n_tags"] or result[u"n_tags_fail"]:
                    result[u"messages"].append(
                        u"Etiquetas: {0} ok, {1} fallos.".format(
                            result[u"n_tags"], result[u"n_tags_fail"]
                        )
                    )
                for m in tag_res.get(u"messages") or []:
                    if m and m not in result[u"messages"]:
                        result[u"messages"].append(m)
            except Exception as ex_tag:
                result[u"messages"].append(
                    u"Etiquetas: {0}".format(_as_unicode(ex_tag))
                )
        else:
            if not result[u"messages"]:
                result[u"messages"].append(u"No se creó ninguna barra.")

        if tg_started:
            if result[u"ok"]:
                tg.Assimilate()
            else:
                tg.RollBack()
            tg_started = False
    except Exception as ex:
        result[u"messages"].append(_as_unicode(ex))
        if tg_started:
            try:
                tg.RollBack()
            except Exception:
                pass
            tg_started = False
    finally:
        try:
            finalizar_armadura_eje_ejecucion()
        except Exception:
            pass
        try:
            finalizar_armadura_conjunto_guid_ejecucion()
        except Exception:
            pass

    return result


def place_retorno_malla_fundacion(
    doc,
    uidoc,
    walls,
    layers,
    cuts_ref_mm=None,
    lap_mode_ui=None,
    concrete_grade=None,
    end_a=None,
    end_b=None,
    mode=None,
):
    """Coloca en muros (misma configuración y modo)."""
    op_mode = normalize_mode(mode)
    wall_list = []
    if isinstance(walls, Wall):
        wall_list = [walls]
    else:
        for w in walls or []:
            if isinstance(w, Wall):
                wall_list.append(w)
    if not wall_list:
        return {
            u"ok": False,
            u"messages": [u"Sin muros."],
            u"rebar_ids": [],
            u"n_layers": 0,
            u"mode": op_mode,
        }

    agg = {
        u"ok": False,
        u"messages": [],
        u"rebar_ids": [],
        u"n_layers": 0,
        u"main_mm": 0.0,
        u"exceeds_12m": False,
        u"has_fund": False,
        u"n_walls": 0,
        u"n_skipped_no_fund": 0,
        u"n_tags": 0,
        u"n_tags_fail": 0,
        u"mode": op_mode,
    }
    for w in wall_list:
        fund = _resolve_fundacion_host(doc, w)
        if op_mode == MODE_INFERIOR_FUND and fund is None:
            agg[u"n_skipped_no_fund"] += 1
            try:
                wid = int(w.Id.IntegerValue)
            except Exception:
                wid = u"?"
            agg[u"messages"].append(
                u"Muro {0}: omitido (sin fundación corrida).".format(wid)
            )
            continue
        res = place_retorno_malla_fundacion_wall(
            doc,
            uidoc,
            w,
            layers,
            cuts_ref_mm=cuts_ref_mm,
            lap_mode_ui=lap_mode_ui,
            concrete_grade=concrete_grade,
            end_a=end_a,
            end_b=end_b,
            mode=op_mode,
        )
        agg[u"n_walls"] += 1
        agg[u"rebar_ids"].extend(res.get(u"rebar_ids") or [])
        agg[u"n_layers"] += int(res.get(u"n_layers") or 0)
        agg[u"n_tags"] += int(res.get(u"n_tags") or 0)
        agg[u"n_tags_fail"] += int(res.get(u"n_tags_fail") or 0)
        agg[u"main_mm"] = max(float(agg[u"main_mm"]), float(res.get(u"main_mm") or 0))
        if res.get(u"exceeds_12m"):
            agg[u"exceeds_12m"] = True
        if res.get(u"has_fund"):
            agg[u"has_fund"] = True
        for m in res.get(u"messages") or []:
            try:
                wid = int(w.Id.IntegerValue)
            except Exception:
                wid = u"?"
            agg[u"messages"].append(u"Muro {0}: {1}".format(wid, m))
        if res.get(u"ok"):
            agg[u"ok"] = True
    if (
        op_mode == MODE_INFERIOR_FUND
        and agg[u"n_walls"] == 0
        and agg[u"n_skipped_no_fund"] > 0
    ):
        agg[u"messages"].insert(
            0,
            u"Ningún muro seleccionado tiene fundación corrida unida.",
        )
    return agg
