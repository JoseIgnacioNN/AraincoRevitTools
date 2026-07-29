# -*- coding: utf-8 -*-
"""
Colocación Revit — Coronamiento muros (superior, capas, split con traslape).

Revit 2024+ | IronPython / pyRevit

Helpers internos (crear capa, stamp post-split, ``dividir_rebar_en_cortes``,
etiquetas/visibilidad) abren sus propias Transaction / TxnScope. El flujo
``place_coronamiento_wall`` las agrupa en un ``TransactionGroup`` con
``Assimilate`` para que Undo muestre un solo paso.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    ElementId,
    Transaction,
    TransactionGroup,
    UnitTypeId,
    UnitUtils,
    Wall,
)

try:
    from Autodesk.Revit.DB.Structure import Rebar
except Exception:
    Rebar = None

import armado_muros_coronamiento as cor
from armado_muros_lineales import location_curve_wall, obtener_espesor_muro_mm_approx
from armado_muros_rebar_params import (
    finalizar_armadura_conjunto_guid_ejecucion,
    finalizar_armadura_eje_ejecucion,
    iniciar_armadura_conjunto_guid_ejecucion,
    iniciar_armadura_eje_ejecucion,
    set_armadura_capa_desde_layer,
    stamp_coronamiento_rebar,
)
from coronamiento_muros_geom import (
    COVER_SUPERIOR_MM,
    clamp_diam_mm,
    clamp_n_bars,
    cover_axis_offset_mm_for_layer,
    estimate_bar_lengths_mm,
    stagger_cuts_for_layer,
    to_dividir_splice_mode,
    traslape_mm_from_diam,
)
from dividir_barra_traslape_punto import dividir_rebar_en_cortes

_TXN_GROUP = u"Arainco: Coronamiento muros"
_TXN_CREATE = u"Arainco: Coronamiento muros (capa)"
_TXN_STAMP = u"Arainco: Coronamiento muros (capa stamp)"
_TAG_EXTREMO = cor.CORONAMIENTO_TAG_EXTREMO_SUP


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def wall_largo_mm(wall):
    lc = location_curve_wall(wall) if location_curve_wall else None
    if lc is None:
        return 0.0
    try:
        return float(
            UnitUtils.ConvertFromInternalUnits(float(lc.Length), UnitTypeId.Millimeters)
        )
    except Exception:
        try:
            return float(lc.Length) * 304.8
        except Exception:
            return 0.0


def wall_espesor_mm(wall):
    try:
        e = obtener_espesor_muro_mm_approx(wall)
        if e is not None and float(e) > 1.0:
            return float(e)
    except Exception:
        pass
    try:
        return float(
            UnitUtils.ConvertFromInternalUnits(float(wall.Width), UnitTypeId.Millimeters)
        )
    except Exception:
        return 200.0


def wall_length_estimate(wall):
    return estimate_bar_lengths_mm(wall_largo_mm(wall), wall_espesor_mm(wall))


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _z_bar_layer_ft(wall, layers, layer_index):
    offset_mm = cover_axis_offset_mm_for_layer(layers, layer_index)
    try:
        _, z_top = cor._wall_z_bounds_ft(wall)
    except Exception:
        return None
    try:
        return float(z_top) - float(_mm_to_internal(offset_mm))
    except Exception:
        return float(z_top) - float(offset_mm) / 304.8


def _bar_type(doc, diam_mm, fallback=None):
    try:
        return cor._bar_type_for_diameter_mm(doc, diam_mm, fallback)
    except Exception:
        pass
    try:
        import armado_muros_cabezal as cabezal

        return cabezal._bar_type_for_diameter_mm(doc, diam_mm, fallback)
    except Exception:
        return fallback


def _stamp_layer(rebar, layer_index):
    if rebar is None:
        return
    try:
        stamp_coronamiento_rebar(rebar)
    except Exception:
        pass
    try:
        set_armadura_capa_desde_layer(rebar, int(layer_index))
    except Exception:
        pass


def _element_id_int(eid):
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid)
        except Exception:
            return None


def _stamp_ids(doc, id_list, layer_index):
    if not id_list:
        return
    t = Transaction(doc, _TXN_STAMP)
    t.Start()
    try:
        for iv in id_list:
            try:
                el = doc.GetElement(ElementId(int(iv)))
            except Exception:
                el = None
            if el is None:
                continue
            if Rebar is not None and not isinstance(el, Rebar):
                continue
            _stamp_layer(el, layer_index)
        t.Commit()
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass


def _register_tag_meta(cor_res, doc, id_ints, wall, z_bar_ft, layer_index):
    """Registra ids + meta para ``aplicar_etiquetado_coronamiento``."""
    if cor_res is None or not id_ints:
        return
    try:
        wid = int(wall.Id.IntegerValue)
    except Exception:
        wid = None
    for iv in id_ints:
        try:
            eid = ElementId(int(iv))
        except Exception:
            continue
        cor_res.setdefault(u"rebars_coronamiento_ids", []).append(eid)
        cor_res.setdefault(u"rebars_coronamiento_id_ints", []).append(int(iv))
        cor_res.setdefault(u"rebars_coronamiento_tag_meta", []).append(
            {
                u"rebar_id": eid,
                u"layer_index": int(layer_index),
                u"wid": wid,
                u"extremo": _TAG_EXTREMO,
                u"zs": float(z_bar_ft),
                u"span_seg": 0.01,
            }
        )


def place_coronamiento_wall(
    doc,
    uidoc,
    wall,
    layers,
    cuts_ref_mm=None,
    lap_mode_ui=None,
):
    """
    Crea coronamiento superior por capa; opcionalmente divide con traslape
    y etiqueta en la vista activa (``EST_A_STRUCTURAL REBAR TAG_WALL_HORIZONTAL``).

    Returns:
        dict: ok, messages, rebar_ids, n_layers, exceeds_12m, developed_mm, main_mm,
              n_tags, n_tags_fail
    """
    result = {
        u"ok": False,
        u"messages": [],
        u"rebar_ids": [],
        u"n_layers": 0,
        u"exceeds_12m": False,
        u"developed_mm": 0.0,
        u"main_mm": 0.0,
        u"n_tags": 0,
        u"n_tags_fail": 0,
    }
    if doc is None or wall is None or not isinstance(wall, Wall):
        result[u"messages"].append(u"Muro no válido.")
        return result

    layers = list(layers or [{u"n_bars": 2, u"diam_mm": 16}])
    if not layers:
        result[u"messages"].append(u"Sin capas configuradas.")
        return result

    est = wall_length_estimate(wall)
    result[u"developed_mm"] = float(est[u"developed_mm"])
    result[u"main_mm"] = float(est[u"main_mm"])
    result[u"exceeds_12m"] = bool(est[u"exceeds_12m"])
    main_mm = float(est[u"main_mm"])
    cuts_ref = list(cuts_ref_mm or [])
    splice_mode = to_dividir_splice_mode(lap_mode_ui)

    cor_res = {
        u"messages": [],
        u"rebars_coronamiento_ids": [],
        u"rebars_coronamiento_id_ints": [],
        u"rebars_coronamiento_tag_meta": [],
        u"n_created": 0,
    }

    iniciar_armadura_conjunto_guid_ejecucion()
    tg = None
    tg_started = False
    place_finished = False
    try:
        try:
            iniciar_armadura_eje_ejecucion(uidoc=uidoc)
        except Exception:
            pass

        tg = TransactionGroup(doc, _TXN_GROUP)
        tg.Start()
        tg_started = True

        created = []
        for li, ly in enumerate(layers):
            n_bars = clamp_n_bars(ly.get(u"n_bars", 2))
            diam_mm = clamp_diam_mm(ly.get(u"diam_mm", 16))
            bt = _bar_type(doc, diam_mm)
            if bt is None:
                result[u"messages"].append(
                    u"Capa {0}: no hay RebarBarType Ø{1}.".format(li + 1, diam_mm)
                )
                continue
            z_bar = _z_bar_layer_ft(wall, layers, li)
            if z_bar is None:
                result[u"messages"].append(
                    u"Capa {0}: no se pudo calcular elevación.".format(li + 1)
                )
                continue

            rb = None
            t = Transaction(doc, _TXN_CREATE)
            t.Start()
            try:
                rb, n_layout, err = cor._create_coronamiento_rebar(
                    doc,
                    wall,
                    wall,
                    n_bars,
                    bt,
                    z_bar,
                    cover_mm=COVER_SUPERIOR_MM,
                    fallback_diam_mm=diam_mm,
                    legs_up=False,
                )
                if rb is not None:
                    _stamp_layer(rb, li)
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

            lap_mm = traslape_mm_from_diam(diam_mm)
            layer_cuts = stagger_cuts_for_layer(cuts_ref, li, main_mm, lap_mm)
            final_ids = []

            if layer_cuts:
                ok_div, msg_div, ids_new = dividir_rebar_en_cortes(
                    doc,
                    rb,
                    layer_cuts,
                    lap_mm=lap_mm,
                    splice_mode=splice_mode,
                )
                if ok_div and ids_new:
                    for eid in ids_new:
                        iv = _element_id_int(eid)
                        if iv is not None:
                            final_ids.append(iv)
                    _stamp_ids(doc, final_ids, li)
                    result[u"messages"].append(
                        u"Capa {0}: {1}Ø{2} · {3} tramo(s).".format(
                            li + 1, n_bars, diam_mm, len(final_ids)
                        )
                    )
                else:
                    iv = _element_id_int(rb.Id)
                    if iv is not None:
                        final_ids.append(iv)
                    result[u"messages"].append(
                        u"Capa {0}: creada sin split ({1}).".format(
                            li + 1, msg_div or u"cortes no válidos"
                        )
                    )
            else:
                iv = _element_id_int(rb.Id)
                if iv is not None:
                    final_ids.append(iv)
                result[u"messages"].append(
                    u"Capa {0}: {1}Ø{2} · sin empalme.".format(li + 1, n_bars, diam_mm)
                )

            _register_tag_meta(cor_res, doc, final_ids, wall, z_bar, li)
            cor_res[u"n_created"] = int(cor_res.get(u"n_created", 0)) + 1
            created.extend(final_ids)
            result[u"n_layers"] += 1

        result[u"rebar_ids"] = created
        result[u"ok"] = len(created) > 0
        if result[u"exceeds_12m"]:
            result[u"messages"].insert(
                0,
                u"Aviso: L desarrollado ≈ {0:.0f} mm > 12 m comercial.".format(
                    result[u"developed_mm"]
                ),
            )

        # Etiquetas en vista activa (familia WALL_HORIZONTAL)
        if created:
            try:
                cor_res = cor.aplicar_etiquetado_coronamiento(
                    doc, cor_res, uidoc=uidoc, aplicar_visibilidad=True,
                )
                n_ok = int(cor_res.get(u"n_cor_tags_created", 0) or 0)
                n_fail = int(cor_res.get(u"n_cor_tags_fail", 0) or 0)
                result[u"n_tags"] = n_ok
                result[u"n_tags_fail"] = n_fail
                if n_ok or n_fail:
                    result[u"messages"].append(
                        u"Etiquetas: {0} ok, {1} fallo.".format(n_ok, n_fail)
                    )
                for m in cor_res.get(u"messages") or []:
                    if m and m not in result[u"messages"]:
                        # Evitar duplicar mensajes de creación; solo tags/visibilidad
                        if u"Etiqueta" in m or u"visible" in m or u"Unobscured" in m:
                            result[u"messages"].append(m)
            except Exception as ex_tag:
                result[u"messages"].append(
                    u"Etiquetas: {0}".format(_as_unicode(ex_tag))
                )

        place_finished = True
    except Exception as ex_place:
        result[u"messages"].append(
            u"Colocar: {0}".format(_as_unicode(ex_place))
        )
    finally:
        if tg_started and tg is not None:
            try:
                if place_finished:
                    tg.Assimilate()
                else:
                    tg.RollBack()
            except Exception:
                try:
                    if tg.HasStarted():
                        tg.RollBack()
                except Exception:
                    pass
        try:
            finalizar_armadura_eje_ejecucion()
        except Exception:
            pass
        try:
            finalizar_armadura_conjunto_guid_ejecucion()
        except Exception:
            pass

    return result
