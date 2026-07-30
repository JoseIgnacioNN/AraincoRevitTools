# -*- coding: utf-8 -*-
"""
Colocación Revit — Coronamiento muros (superior, capas, split con traslape).

Revit 2024+ | IronPython / pyRevit

Traslapos: misma geometría/API que ``56_DividirRebarPuntoTraslape``
(``divide_rebar_at_cuts``). Los cortes de la UI son estaciones sobre el vano
horizontal desde el inicio U (LocationCurve P0); se convierten a distancia
sobre centerline (pata L + vano). El canvas puede espejar izq./der. según
la vista; el origen mm de corte no se invierte.

Helpers internos (crear capa, stamp post-split, divide, etiquetas/visibilidad)
abren sus propias Transaction / TxnScope. El flujo ``place_coronamiento_wall``
las agrupa en un ``TransactionGroup`` con ``Assimilate`` para un solo Undo.
"""

from __future__ import print_function

import os
import sys

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
    estimate_empotrado_bar_lengths_mm,
    stagger_cuts_for_layer,
    to_dividir_lap_mode,
    traslape_mm_from_diam,
)

_TXN_GROUP = u"Arainco: Coronamiento muros"
_TXN_CREATE = u"Arainco: Coronamiento muros (capa)"
_TXN_STAMP = u"Arainco: Coronamiento muros (capa stamp)"
_TAG_EXTREMO = cor.CORONAMIENTO_TAG_EXTREMO_SUP
_GEOM_U_LIBRE = u"u_libre"
_GEOM_EMPOTRADO = u"empotrado"

try:
    from armado_muros_lineales import ordenar_muros_por_base_asc
except Exception:
    ordenar_muros_por_base_asc = None

_DIVIDIR56_SCRIPTS = None
_DIVIDIR56_LOAD_ERROR = None


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _find_dividir_rebar_punto_scripts_dir():
    """Localiza ``56_DividirRebarPuntoTraslape.pushbutton/scripts`` bajo la extensión."""
    global _DIVIDIR56_SCRIPTS, _DIVIDIR56_LOAD_ERROR
    if _DIVIDIR56_SCRIPTS:
        return _DIVIDIR56_SCRIPTS
    here = os.path.dirname(os.path.abspath(__file__))
    ext_root = os.path.dirname(here)
    needle = u"56_DividirRebarPuntoTraslape.pushbutton"
    try:
        for dirpath, dirnames, _filenames in os.walk(ext_root):
            base = os.path.basename(dirpath)
            if base == needle or base.endswith(needle):
                cand = os.path.join(dirpath, u"scripts")
                core = os.path.join(cand, u"dividir_rebar_punto_core.py")
                if os.path.isfile(core):
                    _DIVIDIR56_SCRIPTS = cand
                    return cand
            for skip in (u".git", u"__pycache__", u"node_modules"):
                if skip in dirnames:
                    dirnames.remove(skip)
    except Exception as ex:
        _DIVIDIR56_LOAD_ERROR = _as_unicode(ex)
        return None
    _DIVIDIR56_LOAD_ERROR = u"No se encontró {0}/scripts".format(needle)
    return None


def _ensure_dividir_rebar_punto():
    """
    Antepone scripts de la herramienta 56 e importa ``divide_rebar_at_cuts``.

    Returns:
        (ok, err_msg, divide_fn_or_None)
    """
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


def _cuts_main_to_centerline_mm(rebar, cuts_main_mm):
    """
    Cortes UI (mm sobre el vano horizontal) → mm desde el inicio de la centerline.

    Misma convención que 56: distancia acumulada sobre la polilínea (pata + vano + pata).
    """
    cuts = []
    for c in cuts_main_mm or []:
        try:
            cuts.append(float(c))
        except Exception:
            continue
    if not cuts:
        return []
    try:
        from dividir_rebar_punto_core import (
            _centerline_curves,
            _curve_length,
            internal_to_mm,
        )
    except Exception:
        return cuts
    curves = _centerline_curves(rebar, 0, True, True)
    if not curves:
        curves = _centerline_curves(rebar, 0, True, False)
    if not curves:
        return cuts
    lengths = []
    for crv in curves:
        try:
            lengths.append(float(_curve_length(crv)))
        except Exception:
            lengths.append(0.0)
    if not lengths:
        return cuts
    i_main = 0
    best = -1.0
    for i, leng in enumerate(lengths):
        if leng > best:
            best = leng
            i_main = i
    offset_ft = 0.0
    for i in range(i_main):
        offset_ft += lengths[i]
    try:
        offset_mm = float(internal_to_mm(offset_ft))
    except Exception:
        offset_mm = offset_ft * 304.8
    return [offset_mm + c for c in cuts]


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


def wall_length_estimate(wall):
    return estimate_bar_lengths_mm(wall_largo_mm(wall), wall_espesor_mm(wall))


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


def walls_stacked_z_contact(wall_a, wall_b):
    """True si los dos muros contactan en Z (tope inferior ≈ base superior)."""
    if wall_a is None or wall_b is None:
        return False
    try:
        return bool(
            cor._muro_tiene_apilamiento_inferior(wall_a, [wall_a, wall_b])
            or cor._muro_tiene_apilamiento_inferior(wall_b, [wall_a, wall_b])
        )
    except Exception:
        return False


def _order_walls_base_asc(walls):
    if ordenar_muros_por_base_asc is not None:
        try:
            return list(ordenar_muros_por_base_asc(walls))
        except Exception:
            pass
    keyed = []
    for w in walls or []:
        if w is None:
            continue
        try:
            z_bot, _ = cor._wall_z_bounds_ft(w)
            keyed.append((float(z_bot), w))
        except Exception:
            keyed.append((0.0, w))
    keyed.sort(key=lambda t: t[0])
    return [w for _z, w in keyed]


def _overhang_mm_from_spec(spec):
    """Voladizo libre |u_free − u_reent| en mm (aprox.)."""
    if not spec:
        return 0.0
    try:
        u_reent = float(spec.get(u"u_reent"))
        u_free = float(spec.get(u"u_free"))
        return abs(_ft_to_mm(u_free - u_reent))
    except Exception:
        return 0.0


def _primary_inf_voladizo_spec(specs):
    """Mayor voladizo INF (host = inferior); None si no hay."""
    best = None
    best_oh = -1.0
    for sp in specs or []:
        if sp is None:
            continue
        if sp.get(u"role") != cor.CORONAMIENTO_VOLADIZO_ROLE_INF:
            continue
        oh = _overhang_mm_from_spec(sp)
        if oh > best_oh:
            best_oh = oh
            best = sp
    return best


def resolve_coronamiento_pick(walls):
    """
    Multi-pick → modo U libre / Empotrado.

    Returns:
        dict ok, geom_mode, host, upper, walls_ord, voladizo_specs,
        overhang_mm, embed_side, message
    """
    out = {
        u"ok": False,
        u"geom_mode": _GEOM_U_LIBRE,
        u"host": None,
        u"upper": None,
        u"walls_ord": [],
        u"voladizo_specs": [],
        u"overhang_mm": 0.0,
        u"embed_side": None,
        u"message": u"",
    }
    walls = [w for w in (walls or []) if isinstance(w, Wall)]
    if not walls:
        out[u"message"] = u"Sin muros."
        return out
    if len(walls) == 1:
        out[u"ok"] = True
        out[u"geom_mode"] = _GEOM_U_LIBRE
        out[u"host"] = walls[0]
        out[u"walls_ord"] = [walls[0]]
        return out
    if len(walls) > 2:
        out[u"message"] = u"Máximo 2 muros."
        return out

    walls_ord = _order_walls_base_asc(walls)
    if len(walls_ord) < 2:
        out[u"message"] = u"No se pudieron ordenar los muros."
        return out
    host = walls_ord[0]
    upper = walls_ord[1]
    if not walls_stacked_z_contact(host, upper):
        out[u"message"] = (
            u"Los muros seleccionados no están apilados (sin contacto en Z)."
        )
        return out

    layout = None
    try:
        layout = cor._compute_stacked_layout(walls_ord)
    except Exception:
        layout = None
    specs_all = []
    try:
        if layout is not None:
            specs_all = list(cor._detect_voladizos_stack(walls_ord, layout) or [])
    except Exception:
        specs_all = []
    specs_inf = [
        s
        for s in specs_all
        if s is not None and s.get(u"role") == cor.CORONAMIENTO_VOLADIZO_ROLE_INF
    ]
    primary = _primary_inf_voladizo_spec(specs_inf)
    overhang = _overhang_mm_from_spec(primary) if primary else 0.0
    side = primary.get(u"side") if primary else None

    out[u"ok"] = True
    out[u"geom_mode"] = _GEOM_EMPOTRADO
    out[u"host"] = host
    out[u"upper"] = upper
    out[u"walls_ord"] = walls_ord
    out[u"voladizo_specs"] = specs_inf
    out[u"overhang_mm"] = float(overhang)
    out[u"embed_side"] = side
    return out


def wall_length_estimate_empotrado(host, overhang_mm, diam_mm=16):
    """Estimación de largo Empotrado (host = muro Z baja)."""
    return estimate_empotrado_bar_lengths_mm(
        overhang_mm, wall_espesor_mm(host), diam_mm=diam_mm
    )


def _view_right_unit(view):
    """RightDirection unitario (rx, ry, rz) o None."""
    if view is None:
        return None
    try:
        rd = view.RightDirection
        rx = float(rd.X)
        ry = float(rd.Y)
        rz = float(rd.Z)
    except Exception:
        return None
    rl = (rx * rx + ry * ry + rz * rz) ** 0.5
    if rl < 1e-9:
        return None
    return (rx / rl, ry / rl, rz / rl)


def _view_up_unit(view):
    """UpDirection unitario; respaldo Z+."""
    if view is not None:
        try:
            ud = view.UpDirection
            ux = float(ud.X)
            uy = float(ud.Y)
            uz = float(ud.Z)
            ul = (ux * ux + uy * uy + uz * uz) ** 0.5
            if ul > 1e-9:
                return (ux / ul, uy / ul, uz / ul)
        except Exception:
            pass
    return (0.0, 0.0, 1.0)


def _dot3(pt, axis):
    return (
        float(pt.X) * float(axis[0])
        + float(pt.Y) * float(axis[1])
        + float(pt.Z) * float(axis[2])
    )


def wall_z_bounds_ft(wall):
    """(z_bot, z_top) en pies internos; bbox / helpers coronamiento."""
    if wall is None:
        return 0.0, 3.0
    try:
        z0, z1 = cor._wall_z_bounds_ft(wall)
        z0 = float(z0)
        z1 = float(z1)
        if z1 > z0 + 1e-9:
            return z0, z1
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter

        ph = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
        if ph is not None and ph.HasValue:
            hv = float(ph.AsDouble())
            if hv > 1e-6:
                bb = wall.get_BoundingBox(None)
                z0 = float(bb.Min.Z) if bb is not None else 0.0
                return z0, z0 + hv
    except Exception:
        pass
    try:
        bb = wall.get_BoundingBox(None)
        if bb is not None:
            return float(bb.Min.Z), float(bb.Max.Z)
    except Exception:
        pass
    return 0.0, 3.0


def wall_elev_prism_model(wall, view=None):
    """
    Prisma de elevación desde LocationCurve + altura real.

    Proyecta P0/P1 sobre ``view.RightDirection`` (u horizontal en vista).
    Vertical: cota Z del muro (bbox / WALL_USER_HEIGHT).

    Returns:
        dict | None: u0, u1, u_start, u_end, length_u, z_bot, z_top,
        height_ft, length_mm, height_mm
    """
    if wall is None:
        return None
    lc = location_curve_wall(wall) if location_curve_wall else None
    rd = _view_right_unit(view)
    z_bot, z_top = wall_z_bounds_ft(wall)
    height_ft = max(float(z_top) - float(z_bot), 1e-6)
    u0 = u1 = 0.0
    length_u = 1e-6
    if lc is not None:
        try:
            p0 = lc.GetEndPoint(0)
            p1 = lc.GetEndPoint(1)
            if rd is not None:
                u0 = _dot3(p0, rd)
                u1 = _dot3(p1, rd)
            else:
                # Sin vista: eje XY del propio muro.
                dx = float(p1.X) - float(p0.X)
                dy = float(p1.Y) - float(p0.Y)
                dl = (dx * dx + dy * dy) ** 0.5
                if dl > 1e-9:
                    u0 = 0.0
                    u1 = float(lc.Length)
                else:
                    u0 = 0.0
                    u1 = max(float(lc.Length), 1e-6)
            length_u = max(abs(u1 - u0), 1e-6)
        except Exception:
            try:
                length_u = max(float(lc.Length), 1e-6)
                u0, u1 = 0.0, length_u
            except Exception:
                pass
    u_start = min(u0, u1)
    u_end = max(u0, u1)
    return {
        u"u0": float(u0),
        u"u1": float(u1),
        u"u_start": float(u_start),
        u"u_end": float(u_end),
        u"length_u": float(max(u_end - u_start, length_u, 1e-6)),
        u"z_bot": float(z_bot),
        u"z_top": float(z_top),
        u"height_ft": float(height_ft),
        u"length_mm": float(_ft_to_mm(max(u_end - u_start, length_u))),
        u"height_mm": float(_ft_to_mm(height_ft)),
    }


def walls_elev_layout_model(walls, view=None):
    """
    Layout 1–2 muros en plano de elevación (Right + Z).

    Misma idea que ``compute_stacked_wall_layout(..., view_right_xy)``:
    posiciones relativas u_start/u_end y cotas Z reales. Escala común
    al mapear a canvas (no rectángulos genéricos a todo el ancho).

    Returns:
        dict | None: walls, items, global_u_min/max/span, global_z_min/max/span
    """
    walls = [w for w in (walls or []) if w is not None]
    if not walls:
        return None
    items = []
    u_mins = []
    u_maxs = []
    z_mins = []
    z_maxs = []
    rd = _view_right_unit(view)
    # Preferir stacked layout compartido (axis = Right XY) cuando hay vista.
    stacked = None
    try:
        from armado_muros_lineales import compute_stacked_wall_layout

        view_right_xy = None
        if rd is not None:
            rxy = (rd[0] * rd[0] + rd[1] * rd[1]) ** 0.5
            if rxy > 1e-9:
                view_right_xy = (rd[0] / rxy, rd[1] / rxy)
        stacked = compute_stacked_wall_layout(walls, view_right_xy=view_right_xy)
    except Exception:
        stacked = None
    st_items = (stacked or {}).get(u"items") or []

    for i, wall in enumerate(walls):
        prism = wall_elev_prism_model(wall, view)
        if prism is None:
            continue
        if i < len(st_items) and st_items[i] is not None:
            try:
                it = st_items[i]
                prism[u"u_start"] = float(it.get(u"u_start", prism[u"u_start"]))
                prism[u"u_end"] = float(it.get(u"u_end", prism[u"u_end"]))
                prism[u"u0"] = float(it.get(u"u0", prism[u"u0"]))
                prism[u"u1"] = float(it.get(u"u1", prism[u"u1"]))
                prism[u"length_u"] = float(
                    it.get(
                        u"length_u",
                        max(prism[u"u_end"] - prism[u"u_start"], 1e-6),
                    )
                )
                prism[u"length_mm"] = float(_ft_to_mm(prism[u"length_u"]))
            except Exception:
                pass
        prism[u"wall"] = wall
        items.append(prism)
        u_mins.append(float(prism[u"u_start"]))
        u_maxs.append(float(prism[u"u_end"]))
        z_mins.append(float(prism[u"z_bot"]))
        z_maxs.append(float(prism[u"z_top"]))
    if not items:
        return None
    g_u_min = min(u_mins)
    g_u_max = max(u_maxs)
    g_z_min = min(z_mins)
    g_z_max = max(z_maxs)
    return {
        u"walls": list(walls),
        u"items": items,
        u"global_u_min": float(g_u_min),
        u"global_u_max": float(g_u_max),
        u"global_span_u": float(max(g_u_max - g_u_min, 1e-6)),
        u"global_z_min": float(g_z_min),
        u"global_z_max": float(g_z_max),
        u"global_span_z": float(max(g_z_max - g_z_min, 1e-6)),
        u"right": rd,
        u"up": _view_up_unit(view),
    }


def wall_elev_canvas_flip_for_view(wall, view):
    """
    ¿Invertir el canvas de elevación para coincidir con la vista activa?

    Misma idea que ``compute_stacked_wall_layout(..., view_right_xy)`` /
    ``cabezal_extremos_en_lados_*``: proyectar extremos de LocationCurve
    sobre ``view.RightDirection``.

    Fórmula (cuando la vista es útil):
        rd = normalize(RightDirection)
        u0 = P0 · rd,  u1 = P1 · rd
        flip = (u0 > u1)   # P0 (inicio U / mm=0) queda a la derecha en pantalla

    Convención de cortes (no se invierte el origen mm):
        mm=0 → inicio del vano principal = extremo LocationCurve P0
               (inicio de la cadena U: pata → vano → pata).
        Solo se espeja el mapeo píxel↔mm del dibujo/clic.
        Las siluetas del muro NO se espejan: izq. canvas = menor proyección Right.

    Respaldo ``False`` (canvas izq = P0 = mm 0) si no hay vista/curva,
    |u1−u0| es despreciable frente al largo (muro de canto / eje ≈ ViewDirection),
    o falla la proyección.
    """
    if wall is None or view is None:
        return False
    lc = location_curve_wall(wall) if location_curve_wall else None
    if lc is None:
        return False
    rd = _view_right_unit(view)
    if rd is None:
        return False
    rx, ry, rz = rd
    try:
        p0 = lc.GetEndPoint(0)
        p1 = lc.GetEndPoint(1)
        u0 = _dot3(p0, (rx, ry, rz))
        u1 = _dot3(p1, (rx, ry, rz))
        length_ft = float(lc.Length)
    except Exception:
        return False
    du = u1 - u0
    # Proyección útil: al menos ~15 % del largo (evita muro casi de canto).
    min_span = max(1e-3, 0.15 * max(abs(length_ft), 1e-6))
    if abs(du) < min_span:
        return False
    return bool(u0 > u1)


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


def _register_tag_meta(cor_res, doc, final_ids, wall, z_bar_ft, layer_index, extremo=None):
    tag_ext = extremo if extremo else _TAG_EXTREMO
    for iv in final_ids or []:
        try:
            eid = ElementId(int(iv))
            el = doc.GetElement(eid) if doc is not None else None
        except Exception:
            continue
        if el is None:
            continue
        try:
            cor._registrar_coronamiento_rebar_tag(
                cor_res, el, wall, z_bar_ft, tag_ext, layer_index=layer_index
            )
        except Exception:
            cor_res.setdefault(u"rebars_coronamiento_ids", []).append(eid)
            cor_res.setdefault(u"rebars_coronamiento_id_ints", []).append(int(iv))
            cor_res.setdefault(u"rebars_coronamiento_tag_meta", []).append(
                {
                    u"rebar_id": eid,
                    u"layer_index": int(layer_index),
                    u"wid": int(wall.Id.IntegerValue) if wall is not None else None,
                    u"extremo": tag_ext,
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
    (API 56) y etiqueta en la vista activa (``EST_A_STRUCTURAL REBAR TAG_WALL_HORIZONTAL``).

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
    lap_mode = to_dividir_lap_mode(lap_mode_ui)

    divide_fn = None
    if cuts_ref:
        ok_imp, err_imp, divide_fn = _ensure_dividir_rebar_punto()
        if not ok_imp or divide_fn is None:
            result[u"messages"].append(
                u"Traslape (56) no disponible: {0}".format(err_imp or u"import")
            )
            cuts_ref = []

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
    view = _active_view(uidoc)
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
            layer_cuts_main = stagger_cuts_for_layer(cuts_ref, li, main_mm, lap_mm)
            final_ids = []

            if layer_cuts_main and divide_fn is not None:
                cuts_cl = _cuts_main_to_centerline_mm(rb, layer_cuts_main)
                # Cotas de traslape solo 1ª y 2ª capa; siempre sobre las barras.
                place_dims = int(li) in (0, 1)
                ok_div, msg_div, ids_new, _meta = divide_fn(
                    doc,
                    rb,
                    cuts_cl,
                    concrete_grade=None,
                    view=view,
                    lap_mode=lap_mode,
                    place_lap_dims=place_dims,
                    lap_dim_prefer_above=True,
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


def place_coronamiento_empotrado(
    doc,
    uidoc,
    host,
    layers,
    voladizo_specs,
    cuts_ref_mm=None,
    lap_mode_ui=None,
    overhang_mm=0.0,
):
    """
    Coronamiento Empotrado (V3 voladizo INF): barras en host (Z baja),
    empotro bajo el superior + pata L en voladizo libre.
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
        u"geom_mode": _GEOM_EMPOTRADO,
    }
    if doc is None or host is None or not isinstance(host, Wall):
        result[u"messages"].append(u"Muro host no válido.")
        return result

    specs = [
        s
        for s in (voladizo_specs or [])
        if s is not None and s.get(u"role") == cor.CORONAMIENTO_VOLADIZO_ROLE_INF
    ]
    if not specs:
        result[u"messages"].append(
            u"No hay voladizo (reentrada) en el muro inferior — "
            u"compruebe que el superior sea más corto."
        )
        return result

    layers = list(layers or [{u"n_bars": 2, u"diam_mm": 16}])
    if not layers:
        result[u"messages"].append(u"Sin capas configuradas.")
        return result

    diam0 = clamp_diam_mm(layers[0].get(u"diam_mm", 16))
    est = wall_length_estimate_empotrado(host, overhang_mm, diam_mm=diam0)
    result[u"developed_mm"] = float(est[u"developed_mm"])
    result[u"main_mm"] = float(est[u"main_mm"])
    result[u"exceeds_12m"] = bool(est[u"exceeds_12m"])
    main_mm = float(est[u"main_mm"])
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
    view = _active_view(uidoc)
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
            z_bar = _z_bar_layer_ft(host, layers, li)
            if z_bar is None:
                result[u"messages"].append(
                    u"Capa {0}: no se pudo calcular elevación.".format(li + 1)
                )
                continue

            lap_mm = traslape_mm_from_diam(diam_mm)
            est_li = wall_length_estimate_empotrado(
                host, overhang_mm, diam_mm=diam_mm
            )
            main_li = float(est_li[u"main_mm"])
            layer_cuts_main = stagger_cuts_for_layer(
                cuts_ref, li, main_li if main_li > 1.0 else main_mm, lap_mm
            )

            layer_created = 0
            for spec in specs:
                item = spec.get(u"item") or spec.get(u"item_hi")
                side = spec.get(u"side") or u"?"
                if item is None:
                    result[u"messages"].append(
                        u"Capa {0} voladizo ({1}): sin item de layout.".format(
                            li + 1, side
                        )
                    )
                    continue
                u_embed, u_free, err_u = cor._intervalos_u_voladizo_barra(
                    spec, bt, diam_mm
                )
                if err_u:
                    result[u"messages"].append(
                        u"Capa {0} voladizo ({1}): {2}".format(
                            li + 1, side, err_u
                        )
                    )
                    continue

                rb = None
                t = Transaction(doc, _TXN_CREATE)
                t.Start()
                try:
                    rb, n_layout, err = cor._create_coronamiento_voladizo_rebar(
                        doc,
                        host,
                        item,
                        u_embed,
                        u_free,
                        n_bars,
                        bt,
                        z_bar,
                        leg_up=False,
                        fallback_diam_mm=diam_mm,
                    )
                    if rb is not None:
                        _stamp_layer(rb, li)
                        t.Commit()
                    else:
                        t.RollBack()
                        result[u"messages"].append(
                            u"Capa {0} voladizo ({1}): {2}".format(
                                li + 1, side, err or u"error al crear"
                            )
                        )
                        continue
                except Exception as ex_c:
                    try:
                        t.RollBack()
                    except Exception:
                        pass
                    result[u"messages"].append(
                        u"Capa {0} voladizo ({1}): {2}".format(
                            li + 1, side, _as_unicode(ex_c)
                        )
                    )
                    continue

                final_ids = []
                if layer_cuts_main and divide_fn is not None:
                    cuts_cl = _cuts_main_to_centerline_mm(rb, layer_cuts_main)
                    place_dims = int(li) in (0, 1)
                    ok_div, msg_div, ids_new, _meta = divide_fn(
                        doc,
                        rb,
                        cuts_cl,
                        concrete_grade=None,
                        view=view,
                        lap_mode=lap_mode,
                        place_lap_dims=place_dims,
                        lap_dim_prefer_above=True,
                    )
                    if ok_div and ids_new:
                        for eid in ids_new:
                            iv = _element_id_int(eid)
                            if iv is not None:
                                final_ids.append(iv)
                        _stamp_ids(doc, final_ids, li)
                    else:
                        iv = _element_id_int(rb.Id)
                        if iv is not None:
                            final_ids.append(iv)
                        result[u"messages"].append(
                            u"Capa {0} ({1}): creada sin split ({2}).".format(
                                li + 1, side, msg_div or u"cortes no válidos"
                            )
                        )
                else:
                    iv = _element_id_int(rb.Id)
                    if iv is not None:
                        final_ids.append(iv)

                tag_ext = u"{0}_{1}_{2}".format(
                    cor.CORONAMIENTO_TAG_EXTREMO_VOL,
                    cor.CORONAMIENTO_VOLADIZO_ROLE_INF,
                    side,
                )
                _register_tag_meta(
                    cor_res, doc, final_ids, host, z_bar, li, extremo=tag_ext
                )
                created.extend(final_ids)
                layer_created += 1

            if layer_created > 0:
                cor_res[u"n_created"] = int(cor_res.get(u"n_created", 0)) + 1
                result[u"n_layers"] += 1
                result[u"messages"].append(
                    u"Capa {0}: {1}Ø{2} · {3} voladizo(s).".format(
                        li + 1, n_bars, diam_mm, layer_created
                    )
                )

        result[u"rebar_ids"] = created
        result[u"ok"] = len(created) > 0
        if result[u"exceeds_12m"]:
            result[u"messages"].insert(
                0,
                u"Aviso: L desarrollado ≈ {0:.0f} mm > 12 m comercial.".format(
                    result[u"developed_mm"]
                ),
            )

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
                        if u"Etiqueta" in m or u"visible" in m or u"Unobscured" in m:
                            result[u"messages"].append(m)
            except Exception as ex_tag:
                result[u"messages"].append(
                    u"Etiquetas: {0}".format(_as_unicode(ex_tag))
                )

        place_finished = True
    except Exception as ex_place:
        result[u"messages"].append(
            u"Colocar Empotrado: {0}".format(_as_unicode(ex_place))
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


def place_coronamiento(
    doc,
    uidoc,
    wall,
    layers,
    cuts_ref_mm=None,
    lap_mode_ui=None,
    geom_mode=None,
    voladizo_specs=None,
    overhang_mm=0.0,
):
    """Despacha U libre o Empotrado según ``geom_mode``."""
    mode = _as_unicode(geom_mode or _GEOM_U_LIBRE).strip().lower()
    if mode == _GEOM_EMPOTRADO:
        return place_coronamiento_empotrado(
            doc,
            uidoc,
            wall,
            layers,
            voladizo_specs,
            cuts_ref_mm=cuts_ref_mm,
            lap_mode_ui=lap_mode_ui,
            overhang_mm=overhang_mm,
        )
    return place_coronamiento_wall(
        doc,
        uidoc,
        wall,
        layers,
        cuts_ref_mm=cuts_ref_mm,
        lap_mode_ui=lap_mode_ui,
    )
