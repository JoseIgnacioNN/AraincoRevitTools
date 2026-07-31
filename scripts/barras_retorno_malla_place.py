# -*- coding: utf-8 -*-
"""
Colocación Revit — Barras de retorno de malla.

Longitudinales en espesor del muro (≥2). Con fundación: Z = fondo fund. + 50 mm.
Sin fundación: Z = base muro + 50 mm. Host estructural = muro (misma estrategia
que coronamiento inferior). Extremos: Pata 90º / Empotramiento. Empalmes vía API 56.
Tras crear, etiquetas bajo las barras en ActiveView
(EST_A_STRUCTURAL REBAR TAG_WALL_HORIZONTAL; meta cor_pie → offset −Up).
Empalmes: Detail + cotas de traslape vía divide 56 (place_lap_dims).

Revit 2024+ | IronPython / pyRevit
"""

from __future__ import print_function

import os
import sys

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    Curve,
    ElementId,
    Line,
    Transaction,
    TransactionGroup,
    UnitTypeId,
    UnitUtils,
    Wall,
    XYZ,
)
from System.Collections.Generic import List

try:
    from Autodesk.Revit.DB.Structure import (
        Rebar,
        RebarHookOrientation,
        RebarStyle,
    )
except Exception:
    Rebar = None
    RebarHookOrientation = None
    RebarStyle = None

import armado_muros_coronamiento as cor
from armado_muros_lineales import location_curve_wall, obtener_espesor_muro_mm_approx
from armado_muros_rebar_params import (
    finalizar_armadura_conjunto_guid_ejecucion,
    finalizar_armadura_eje_ejecucion,
    iniciar_armadura_conjunto_guid_ejecucion,
    iniciar_armadura_eje_ejecucion,
    set_armadura_capa_desde_layer,
)

try:
    from armado_muros_rebar_params import activar_armadura_arainco
except Exception:
    activar_armadura_arainco = None

from barras_retorno_malla_geom import (
    COVER_FUND_BOT_MM,
    COVER_WALL_BOT_MM,
    COVER_WALL_MM,
    END_EMPOTRO,
    END_PATA,
    MARGIN_END_MM,
    MAX_BARRA_COMERCIAL_MM,
    clamp_diam_mm,
    clamp_n_bars,
    cover_axis_offset_mm_for_layer,
    empotramiento_mm_from_diam,
    long_bar_length_mm,
    normalize_concrete_grade,
    normalize_end_condition,
    pata_mm_from_diam,
    stagger_cuts_for_layer,
    to_dividir_lap_mode,
    traslape_mm_from_diam,
)

_TXN_GROUP = u"Arainco: Barras retorno malla"
_TXN_CREATE = u"Arainco: Barras retorno malla (capa)"
_DIALOG_TITLE = u"Arainco: Barras de retorno de malla"
_REBAR_TAG_FAMILY = u"EST_A_STRUCTURAL REBAR TAG_WALL_HORIZONTAL"

_DIVIDIR56_LOAD_ERROR = u""


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _ft_to_mm(ft):
    try:
        return float(
            UnitUtils.ConvertFromInternalUnits(float(ft), UnitTypeId.Millimeters)
        )
    except Exception:
        try:
            return float(ft) * 304.8
        except Exception:
            return 0.0


def _active_view(uidoc):
    if uidoc is None:
        return None
    try:
        return uidoc.ActiveView
    except Exception:
        return None


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


def wall_height_mm(wall):
    try:
        z0, z1 = cor._wall_z_bounds_ft(wall)
        return abs(_ft_to_mm(float(z1) - float(z0)))
    except Exception:
        return 2800.0


def resolve_foundation(doc, wall):
    """Fundación principal unida al muro, o None."""
    try:
        return cor._fundacion_principal_muro(doc, wall)
    except Exception:
        return None


def foundation_dims_mm(foundation, wall_thickness_mm=None):
    """width_mm, height_mm, offset_from_wall_mm (voladizo aprox. centrado en espesor muro)."""
    try:
        wt = float(wall_thickness_mm) if wall_thickness_mm is not None else 200.0
    except Exception:
        wt = 200.0
    if wt < 1.0:
        wt = 200.0
    if foundation is None:
        return None
    try:
        bb = foundation.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None:
        return {
            u"width_mm": 600.0,
            u"height_mm": 400.0,
            u"offset_from_wall_mm": max(0.0, (600.0 - wt) * 0.5),
        }
    try:
        dx = abs(float(bb.Max.X) - float(bb.Min.X))
        dy = abs(float(bb.Max.Y) - float(bb.Min.Y))
        dz = abs(float(bb.Max.Z) - float(bb.Min.Z))
        plan = max(dx, dy)
        fund_w = _ft_to_mm(plan)
        return {
            u"width_mm": fund_w,
            u"height_mm": _ft_to_mm(dz),
            u"offset_from_wall_mm": max(0.0, (fund_w - wt) * 0.5),
        }
    except Exception:
        return {
            u"width_mm": 600.0,
            u"height_mm": 400.0,
            u"offset_from_wall_mm": max(0.0, (600.0 - wt) * 0.5),
        }


def detect_joined_wall_relation(doc, wall):
    """
    Detecta muro unido arriba/abajo (solo para elevación). No bloquea ni muestra cards.

    Returns:
        dict|None: {id, relation: stacked_above|stacked_below}
    """
    if doc is None or wall is None:
        return None
    try:
        from wall_node_boolean_section_rps import (
            hay_muro_unido_cara_inferior,
            hay_muro_unido_cara_superior,
        )
    except Exception:
        try:
            from armado_muros_nodo_refuerzo_post import (
                hay_muro_unido_cara_inferior,
                hay_muro_unido_cara_superior,
            )
        except Exception:
            return None

    try:
        if hay_muro_unido_cara_superior(doc, wall):
            return {u"relation": u"stacked_above", u"id": u"↑"}
    except Exception:
        pass
    try:
        if hay_muro_unido_cara_inferior(doc, wall):
            return {u"relation": u"stacked_below", u"id": u"↓"}
    except Exception:
        pass
    return None


def wall_meta_for_ui(doc, wall):
    """Metadatos de muro + fundación opcional para preview."""
    thick = wall_espesor_mm(wall)
    fund = resolve_foundation(doc, wall) if doc is not None else None
    dims = (
        foundation_dims_mm(fund, wall_thickness_mm=thick)
        if fund is not None
        else None
    )
    joined = detect_joined_wall_relation(doc, wall)
    try:
        wid = int(wall.Id.IntegerValue)
    except Exception:
        wid = 0
    fund_info = None
    if fund is not None and dims is not None:
        try:
            fid = int(fund.Id.IntegerValue)
        except Exception:
            fid = 0
        fund_info = {
            u"id": fid,
            u"element": fund,
            u"width_mm": float(dims[u"width_mm"]),
            u"height_mm": float(dims[u"height_mm"]),
            u"offset_from_wall_mm": float(dims[u"offset_from_wall_mm"]),
        }
    return {
        u"wall": wall,
        u"id": wid,
        u"thickness_mm": float(thick),
        u"length_mm": wall_largo_mm(wall),
        u"height_mm": wall_height_mm(wall),
        u"foundation": fund_info,
        u"joined": joined,
        u"main_mm": long_bar_length_mm(wall_largo_mm(wall), MARGIN_END_MM),
    }


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


def _find_dividir_rebar_punto_scripts_dir():
    global _DIVIDIR56_LOAD_ERROR
    here = os.path.dirname(os.path.abspath(__file__))
    cursor = here
    for _ in range(24):
        candidate = os.path.join(
            cursor,
            u"BIMTools.tab",
            u"Armadura.panel",
            u"56_DividirRebarPuntoTraslape.pushbutton",
            u"scripts",
        )
        if os.path.isdir(candidate):
            return candidate
        # también relativo a extensión
        candidate2 = os.path.join(
            cursor,
            u"Armadura.panel",
            u"56_DividirRebarPuntoTraslape.pushbutton",
            u"scripts",
        )
        if os.path.isdir(candidate2):
            return candidate2
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    # walk from extension root via this file's parent chain
    ext = os.path.dirname(here)
    for root, dirs, _files in os.walk(ext):
        if u"56_DividirRebarPuntoTraslape.pushbutton" in dirs:
            sc = os.path.join(root, u"56_DividirRebarPuntoTraslape.pushbutton", u"scripts")
            if os.path.isdir(sc):
                return sc
        # limit depth
        if root.count(os.sep) - ext.count(os.sep) > 4:
            dirs[:] = []
    _DIVIDIR56_LOAD_ERROR = u"Scripts 56 no encontrados."
    return None


def _ensure_dividir_rebar_punto():
    global _DIVIDIR56_LOAD_ERROR
    scripts_dir = _find_dividir_rebar_punto_scripts_dir()
    if not scripts_dir:
        return False, _DIVIDIR56_LOAD_ERROR or u"Scripts 56 no encontrados.", None
    try:
        if scripts_dir in sys.path:
            sys.path.remove(scripts_dir)
        sys.path.insert(0, scripts_dir)
    except Exception:
        if scripts_dir not in sys.path:
            sys.path.insert(0, scripts_dir)
    try:
        from dividir_rebar_punto_core import divide_rebar_at_cuts

        return True, u"", divide_rebar_at_cuts
    except Exception as ex:
        _DIVIDIR56_LOAD_ERROR = _as_unicode(ex)
        return False, u"Import 56: {0}".format(_DIVIDIR56_LOAD_ERROR), None


def _xyz(x, y, z):
    return XYZ(float(x), float(y), float(z))


def _offset_point(p, direction, dist_ft):
    return _xyz(
        float(p.X) + float(direction.X) * float(dist_ft),
        float(p.Y) + float(direction.Y) * float(dist_ft),
        float(p.Z) + float(direction.Z) * float(dist_ft),
    )


def _z_bar_layer_ft(wall, fund, layers, layer_index):
    """
    Z del eje de la capa.
    Con fund: fondo fundación + cover_fund + Ø/2 + spacing.
    Sin fund: base muro + cover_wall_bot + Ø/2 + spacing.
    """
    base_cover = COVER_FUND_BOT_MM if fund is not None else COVER_WALL_BOT_MM
    offset_mm = cover_axis_offset_mm_for_layer(layers, layer_index, base_cover)
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


def _build_longitudinal_curves(
    wall,
    z_bar_ft,
    diam_mm,
    end_a,
    end_b,
    concrete_grade,
):
    """
    Polilínea longitudinal (1–3 tramos) en el plano del muro.

    Origen mm: LocationCurve P0 + MARGIN_END.
    Empotramiento: prolonga más allá del margen.
    Pata 90º: tramo vertical hacia +Z en el extremo.

    Returns:
        curves, normal, distrib_ft, main_mm, err
    """
    frame = cor._wall_location_frame(wall)
    if frame is None:
        return None, None, None, 0.0, u"Sin LocationCurve válida."

    p0 = frame[u"p0"]
    p1 = frame[u"p1"]
    t_hat = frame[u"t_hat"]
    normal = frame[u"normal"]
    espesor_ft = float(frame[u"espesor_ft"])
    wall_len_ft = float(p0.DistanceTo(p1))
    if wall_len_ft < 1e-9:
        return None, None, None, 0.0, u"Muro de longitud nula."

    margin_ft = float(_mm_to_internal(MARGIN_END_MM))
    cover_ft = float(_mm_to_internal(COVER_WALL_MM))
    # Misma convención que cabezal / coronamiento:
    # 1ª barra en cara +normal (wall.Orientation) a cover del eje;
    # SetLayoutAsFixedNumber(..., barsOnNormalSide=False) reparte hacia −normal
    # dentro del espesor del muro (no del voladizo de fundación).
    half = espesor_ft * 0.5
    lateral_ft = half - cover_ft
    if lateral_ft < 0:
        lateral_ft = 0.0

    z = float(z_bar_ft)
    base0 = _xyz(float(p0.X), float(p0.Y), z)
    base1 = _xyz(float(p1.X), float(p1.Y), z)
    try:
        n_hat = normal.Normalize()
    except Exception:
        n_hat = normal
    pt_a = _offset_point(base0, n_hat, lateral_ft)
    pt_b = _offset_point(base1, n_hat, lateral_ft)
    # holgura extremos
    pt_a = _offset_point(pt_a, t_hat, margin_ft)
    pt_b = _offset_point(pt_b, t_hat.Negate(), margin_ft)

    ea = normalize_end_condition(end_a)
    eb = normalize_end_condition(end_b)
    embed_mm = empotramiento_mm_from_diam(diam_mm, concrete_grade)
    embed_ft = float(_mm_to_internal(embed_mm))
    pata_mm = pata_mm_from_diam(diam_mm)
    pata_ft = float(_mm_to_internal(pata_mm))

    if ea == END_EMPOTRO:
        pt_a = _offset_point(pt_a, t_hat.Negate(), embed_ft)
    if eb == END_EMPOTRO:
        pt_b = _offset_point(pt_b, t_hat, embed_ft)

    main_mm = long_bar_length_mm(_ft_to_mm(wall_len_ft), MARGIN_END_MM)
    curves = []

    # Pata A: vertical desde pt_a_horiz; si pata, el horizontal empieza en pt_a
    start_h = pt_a
    end_h = pt_b
    up = XYZ.BasisZ

    if ea == END_PATA and pata_ft > 1e-6:
        p_leg_a = _offset_point(pt_a, up, pata_ft)
        try:
            curves.append(Line.CreateBound(p_leg_a, pt_a))
        except Exception:
            pass

    try:
        curves.append(Line.CreateBound(start_h, end_h))
    except Exception as ex:
        return None, None, None, main_mm, u"Tramo horizontal: {0}".format(_as_unicode(ex))

    if eb == END_PATA and pata_ft > 1e-6:
        p_leg_b = _offset_point(pt_b, up, pata_ft)
        try:
            curves.append(Line.CreateBound(pt_b, p_leg_b))
        except Exception:
            pass

    if not curves:
        return None, None, None, main_mm, u"Sin curvas."

    # Ancho de distribución = espesor − 2·cover (entre 1ª y última barra)
    usable = max(0.0, espesor_ft - 2.0 * cover_ft)
    distrib_ft = usable  # FixedNumber: 1ª en +normal·lateral → última en −normal·lateral

    return curves, n_hat, distrib_ft, main_mm, None


def _cuts_ui_to_centerline_mm(cuts_ui_mm, diam_mm, end_a, concrete_grade):
    """
    UI cuts (mm sobre vano claro entre holguras) → mm sobre centerline CreateFromCurves.

    Centerline: [pata A opcional] + horizontal (posible empotro A/B) + [pata B opcional].
    El vano claro empieza tras pata A (si hay) y tras el tramo de empotramiento A (si hay).
    """
    ea = normalize_end_condition(end_a)
    offset = 0.0
    if ea == END_PATA:
        offset += float(pata_mm_from_diam(diam_mm))
    elif ea == END_EMPOTRO:
        offset += float(empotramiento_mm_from_diam(diam_mm, concrete_grade))
    out = []
    for c in cuts_ui_mm or []:
        try:
            out.append(float(c) + offset)
        except Exception:
            continue
    return out


def _create_retorno_rebar(
    doc,
    wall,
    host,
    n_bars,
    bar_type,
    z_bar_ft,
    diam_mm,
    end_a,
    end_b,
    concrete_grade,
):
    """Crea un set longitudinal con layout Fixed Number en espesor."""
    if doc is None or wall is None or host is None or bar_type is None:
        return None, 0, u"Doc, muro, host o tipo no válido."
    if Rebar is None:
        return None, 0, u"API Rebar no disponible."

    n_bars = clamp_n_bars(n_bars)
    curves, normal, distrib_ft, _main, err = _build_longitudinal_curves(
        wall, z_bar_ft, diam_mm, end_a, end_b, concrete_grade
    )
    if err:
        return None, 0, err
    if not curves:
        return None, 0, u"Sin curvas."

    try:
        curves_list = List[Curve]()
        for c in curves:
            curves_list.Add(c)
    except Exception as ex_cl:
        return None, 0, u"IList[Curve]: {0}".format(_as_unicode(ex_cl))

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
        return None, 0, _as_unicode(ex_cf)

    if rebar is None:
        return None, 0, u"CreateFromCurves devolvió None."

    if n_bars > 1 and float(distrib_ft or 0.0) > 1e-9:
        try:
            accessor = rebar.GetShapeDrivenAccessor()
            accessor.SetLayoutAsFixedNumber(
                int(n_bars), float(distrib_ft), False, True, True
            )
        except Exception as ex_lay:
            return None, 0, u"SetLayoutAsFixedNumber: {0}".format(_as_unicode(ex_lay))

    if activar_armadura_arainco is not None:
        try:
            activar_armadura_arainco(rebar)
        except Exception:
            pass

    return rebar, int(n_bars), None


def _element_ids_from_ints(id_ints):
    out = []
    seen = set()
    for iv in id_ints or []:
        try:
            i = int(iv)
        except Exception:
            continue
        if i in seen:
            continue
        seen.add(i)
        try:
            out.append(ElementId(i))
        except Exception:
            pass
    return out


def _tag_retorno_rebars(doc, uidoc, rebar_id_ints, tag_meta=None):
    """
    Etiqueta sets en ActiveView (familia WALL_HORIZONTAL).

    Con ``tag_meta.extremo == cor_pie``, el helper de cabezal usa layout
    coronamiento: TagHeadPosition bajo las barras (−view.UpDirection).

    No lanza: si falta tipo/vista, deja barras y reporta mensaje.
    Returns:
        dict: n_ok, n_fail, messages
    """
    result = {u"n_ok": 0, u"n_fail": 0, u"messages": []}
    ids = _element_ids_from_ints(rebar_id_ints)
    if not ids:
        return result
    view = _active_view(uidoc)
    if view is None:
        result[u"n_fail"] = len(ids)
        result[u"messages"].append(
            u"Etiquetas: sin vista activa (UIDocument.ActiveView)."
        )
        return result
    try:
        from armado_muros_cabezal_tags import (
            CABEZAL_REBAR_TAG_FAMILY_NAME,
            etiquetar_cabezal_longitudinales_en_vista,
        )
    except Exception as ex_imp:
        result[u"n_fail"] = len(ids)
        result[u"messages"].append(
            u"Etiquetas: módulo no disponible ({0}).".format(_as_unicode(ex_imp))
        )
        return result

    fam = CABEZAL_REBAR_TAG_FAMILY_NAME or _REBAR_TAG_FAMILY
    try:
        tag_res = etiquetar_cabezal_longitudinales_en_vista(
            doc,
            view,
            ids,
            tag_meta=tag_meta,
            family_name=fam,
        )
    except Exception as ex_tag:
        result[u"n_fail"] = len(ids)
        result[u"messages"].append(
            u"Etiquetas: {0}".format(_as_unicode(ex_tag))
        )
        return result

    result[u"n_ok"] = int(tag_res.get(u"n_ok", 0) or 0)
    result[u"n_fail"] = int(tag_res.get(u"n_fail", 0) or 0)
    for m in tag_res.get(u"messages") or []:
        if m:
            result[u"messages"].append(m)
    if result[u"n_ok"] == 0 and result[u"n_fail"] > 0 and not result[u"messages"]:
        result[u"messages"].append(
            u"Etiquetas: no se pudo etiquetar (¿falta familia «{0}»?).".format(fam)
        )
    return result


def _stamp_layer(rebar, layer_index):
    if rebar is None:
        return
    try:
        set_armadura_capa_desde_layer(rebar, int(layer_index))
    except Exception:
        pass
    if activar_armadura_arainco is not None:
        try:
            activar_armadura_arainco(rebar)
        except Exception:
            pass


def place_barras_retorno_malla_wall(
    doc,
    uidoc,
    wall,
    layers,
    cuts_ref_mm=None,
    lap_mode_ui=None,
    concrete_grade=None,
    end_a=None,
    end_b=None,
):
    """
    Coloca barras de retorno por capa en un muro.

    Returns:
        dict: ok, messages, rebar_ids, n_layers, main_mm, exceeds_12m, has_fund
    """
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
    }
    if doc is None or wall is None or not isinstance(wall, Wall):
        result[u"messages"].append(u"Muro no válido.")
        return result

    layers = list(layers or [{u"n_bars": 2, u"diam_mm": 10}])
    if not layers:
        result[u"messages"].append(u"Sin capas configuradas.")
        return result

    grade = normalize_concrete_grade(concrete_grade)
    ea = normalize_end_condition(end_a if end_a is not None else END_EMPOTRO)
    eb = normalize_end_condition(end_b if end_b is not None else END_PATA)
    fund = resolve_foundation(doc, wall)
    # Host estructural = muro (geometría Z puede estar en fundación, como coronamiento inf.).
    host = wall
    result[u"has_fund"] = fund is not None
    try:
        wall_id_int = int(wall.Id.IntegerValue)
    except Exception:
        wall_id_int = 0

    main_mm = long_bar_length_mm(wall_largo_mm(wall), MARGIN_END_MM)
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
            diam_mm = clamp_diam_mm(ly.get(u"diam_mm", 10))
            bt = _bar_type(doc, diam_mm)
            if bt is None:
                result[u"messages"].append(
                    u"Capa {0}: no hay RebarBarType Ø{1}.".format(li + 1, diam_mm)
                )
                continue
            z_bar = _z_bar_layer_ft(wall, fund, layers, li)
            if z_bar is None:
                result[u"messages"].append(
                    u"Capa {0}: no se pudo calcular elevación Z.".format(li + 1)
                )
                continue

            rb = None
            t = Transaction(doc, _TXN_CREATE)
            t.Start()
            try:
                rb, n_layout, err = _create_retorno_rebar(
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

            lap_mm = traslape_mm_from_diam(diam_mm, grade)
            layer_cuts = stagger_cuts_for_layer(cuts_ref, li, main_mm, lap_mm)
            final_ids = []

            if layer_cuts and divide_fn is not None:
                cuts_cl = _cuts_ui_to_centerline_mm(
                    layer_cuts, diam_mm, ea, grade
                )
                # Cotas de traslape vía 56 (Detail + NewDimension); fallan en soft.
                # prefer_above: cotas hacia +Up; etiquetas van debajo (−Up, cor_pie).
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
                            final_ids.append(int(iv.IntegerValue if hasattr(iv, "IntegerValue") else iv))
                        except Exception:
                            try:
                                final_ids.append(int(iv))
                            except Exception:
                                pass
                    for iv in final_ids:
                        try:
                            el = doc.GetElement(ElementId(int(iv)))
                            _stamp_layer(el, li)
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
                    # cor_pie → layout coronamiento con offset −Up (etiqueta bajo la barra).
                    tag_meta.append(
                        {
                            u"rebar_id": ElementId(int(rid)),
                            u"layer_index": int(li),
                            u"wid": wall_id_int,
                            u"extremo": u"cor_pie",
                            u"zs": float(z_bar),
                            u"span_seg": 0.01,
                        }
                    )
                except Exception:
                    pass

            created.extend(final_ids)
            result[u"n_layers"] += 1

        result[u"rebar_ids"] = created
        if created:
            result[u"ok"] = True
            result[u"messages"].append(
                u"Creadas {0} barra(s)/set(s) en {1} capa(s).".format(
                    len(created), result[u"n_layers"]
                )
            )
            try:
                tag_res = _tag_retorno_rebars(doc, uidoc, created, tag_meta=tag_meta)
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


def place_barras_retorno_malla(
    doc,
    uidoc,
    walls,
    layers,
    cuts_ref_mm=None,
    lap_mode_ui=None,
    concrete_grade=None,
    end_a=None,
    end_b=None,
):
    """Coloca en uno o varios muros (misma configuración)."""
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
        u"n_tags": 0,
        u"n_tags_fail": 0,
    }
    for w in wall_list:
        res = place_barras_retorno_malla_wall(
            doc,
            uidoc,
            w,
            layers,
            cuts_ref_mm=cuts_ref_mm,
            lap_mode_ui=lap_mode_ui,
            concrete_grade=concrete_grade,
            end_a=end_a,
            end_b=end_b,
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
    return agg
