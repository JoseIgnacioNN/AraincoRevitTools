# -*- coding: utf-8 -*-
"""
Arainco: Capas Adicionales en Muro.

Flujo:
1. Pick rebar longitudinal con Armadura_Conjunto_GUID → analizar capas.
2. Inferir extremo (inicio/fin) desde la seed respecto al host.
3. Pick múltiple de muros (mismo filtro/vista que Armado Muros v3).
4. UI estilo Armado Muros v3: elevación (canvas) + rail derecho
   (N capas, Ø, barras, spacing, confinamiento).
   Empalmes Auto/Tramo/Cont. en el pie del canvas (mismo helper v3).
   Alternancia A/B siempre activa (paridad de capa GUID, como v3).
   Terminación cabeza (Z máx. por capa): Empotramiento L(Ø) o Pata L.
5. Crear vía pipeline cabezal (un extremo) con offset de índice de capa
   para continuar el stack GUID (paridad empalme incluida).
6. Mismo Armadura_Conjunto_GUID; Armadura_Capa consecutivo (máx+1…).

Transacciones (Revit 2025+ / pyRevit): ``TransactionGroup`` asimilado
(un solo Undo) + ``revit.Transaction`` en el post-proceso; excepción →
rollback limpio del grupo.

No modifica Armado Muros v3 ni Capas adicionales GUID.
"""

from __future__ import print_function

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System import AppDomain, EventHandler
from System.Windows import RoutedEventHandler, WindowState
from System.Windows.Markup import XamlReader

from Autodesk.Revit.DB import (
    ElementId,
    ElementTransformUtils,
    FilteredElementCollector,
    LocationCurve,
    Wall,
    XYZ,
)
from Autodesk.Revit.DB.Structure import MultiplanarOption, Rebar, RebarBarType
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog

from armado_muros_cabezal import (
    CABEZAL_CONFINEMENT_NONE,
    CABEZAL_EXTREMO_FIN,
    CABEZAL_EXTREMO_INICIO,
    CABEZAL_LAYER_PITCH_MM,
    CABEZAL_MAX_BARRAS_POR_CAPA,
    CABEZAL_MIN_BARRAS_POR_CAPA,
    TOP_TERMINATION_EMPOTRAMIENTO,
    TOP_TERMINATION_PATA_L,
    _mm_to_internal,
    _troceo_por_muro_from_extremo_cfg,
    _wall_longitudinal_at_extremo,
    aplicar_cabezales_muros,
    cabezal_aplicar_etiquetado_pendiente,
    cabezal_confinement_options,
    compute_auto_troceo_por_muro_flags,
    default_cabezal_layer_config,
    default_cabezal_muro_config,
    normalize_cabezal_confinement,
    sync_troceo_effective_for_extremo,
    wall_id_int,
)
from armado_muros_lineales import (
    _guard_vista_armado_muros,
    _pick_muros,
    cabezal_extremos_en_lados_stacked,
    compute_stacked_wall_layout,
    obtener_espesor_muro_mm_approx,
    ordenar_muros_por_base_asc,
)
from armado_muros_rebar_params import (
    finalizar_armadura_conjunto_guid_ejecucion,
    iniciar_armadura_conjunto_guid_ejecucion,
    iniciar_armadura_eje_ejecucion,
    set_armadura_capa_desde_layer,
    stamp_armadura_conjunto_guid,
)
from armado_muros_v3_segments import build_bar_segments
from armado_muros_v3_troceo import (
    TROCEO_AUTO,
    TROCEO_CONT,
    TROCEO_TRAMO,
    cycle_troceo_mode,
    effective_empalme,
    empalme_indices_from_modes,
    lowest_z_meta_index,
)
from bimtools_ui_tokens import WINDOW_CHROME_TITLE
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from capas_adicionales_guid import (
    _max_capa_index,
    _parse_capa_index,
    _pick_rebar_element,
    analyze_conjunto,
)
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    revit_main_hwnd,
)
from seleccionar_capa_armadura import get_armadura_capa

try:
    unicode
except NameError:
    unicode = str  # noqa: A001


_TOOL_TITLE = u"Arainco: Capas Adicionales en Muro"
_TXN_NAME = u"Arainco: Capas Adicionales en Muro"
_APPDOMAIN_WINDOW_KEY = u"BIMTools.CapasAdicionalesEnMuro.Window"
_APPDOMAIN_CTRL_KEY = u"BIMTools.CapasAdicionalesEnMuro.Controller"
_APPDOMAIN_CREATE_TARGET_KEY = u"BIMTools.CapasAdicionalesEnMuro.CreateTarget"
_APPDOMAIN_EMPALME_ALTERNANCIA = u"Arainco_Cabezal_Empalme_Alternancia"
_APPDOMAIN_TOP_TERMINATION = u"Arainco_Cabezal_TopTermination"
_APPDOMAIN_SKIP_TOP_EMBED_COLLISION = u"Arainco_Cabezal_SkipTopEmbedCollision"

_TERMINACION_CABEZA_OPTS = (
    (TOP_TERMINATION_EMPOTRAMIENTO, u"Empotramiento L(Ø)"),
    (TOP_TERMINATION_PATA_L, u"Pata L"),
)
_TERMINACION_CABEZA_DEFAULT = TOP_TERMINATION_EMPOTRAMIENTO


# ── utilidades ───────────────────────────────────────────────────────────────


def _pyrevit_revit():
    """Módulo ``pyrevit.revit`` (Transaction / TransactionGroup nativos)."""
    from pyrevit import revit as _rv

    return _rv


def _as_unicode(val):
    if val is None:
        return u""
    try:
        return unicode(val)
    except Exception:
        try:
            return str(val)
        except Exception:
            return u""


def _mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    hwnd = None
    try:
        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    try:
        from bimtools_instruction_dialog import show_message_dialog

        show_message_dialog(
            _TOOL_TITLE,
            instruction,
            content=content,
            ok_text=ok_text,
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        body = instruction
        if content:
            body = instruction + u"\n\n" + content
        TaskDialog.Show(_TOOL_TITLE, body)
    except Exception:
        pass


def _element_id_int(eid):
    if eid is None or eid == ElementId.InvalidElementId:
        return None
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid.Value)
        except Exception:
            return None


def _bbox_center(el):
    if el is None:
        return None
    try:
        bb = el.get_BoundingBox(None)
        if bb is None:
            return None
        return XYZ(
            (float(bb.Min.X) + float(bb.Max.X)) * 0.5,
            (float(bb.Min.Y) + float(bb.Max.Y)) * 0.5,
            (float(bb.Min.Z) + float(bb.Max.Z)) * 0.5,
        )
    except Exception:
        return None


def _host_wall_of(rebar, doc):
    if rebar is None or doc is None:
        return None
    try:
        hid = rebar.GetHostId()
        host = doc.GetElement(hid) if hid else None
        if isinstance(host, Wall):
            return host
    except Exception:
        pass
    return None


def _seed_point(analysis):
    """Punto representativo: centroide de la capa más exterior, o seed."""
    layers = (analysis or {}).get(u"layers") or []
    if layers:
        # Preferir capa índice mínimo (cara del extremo).
        try:
            outer = min(layers, key=lambda ly: int(ly.get(u"index", 999)))
        except Exception:
            outer = layers[0]
        c = outer.get(u"centroid")
        if c is not None:
            return c
        rbs = outer.get(u"rebars") or []
        if rbs:
            c = _bbox_center(rbs[0])
            if c is not None:
                return c
    seed = (analysis or {}).get(u"seed")
    return _bbox_center(seed)


def scope_analysis_to_seed_host(doc, analysis):
    """
    Acota capas del GUID al muro host de la seed.

    Evita mezclar copias del mismo GUID en otros extremos/hosts al calcular
    la última capa y el extremo.
    """
    if not analysis or not analysis.get(u"guid"):
        return analysis
    seed = analysis.get(u"seed")
    host = analysis.get(u"host")
    if not isinstance(host, Wall):
        host = _host_wall_of(seed, doc)
    host_id = wall_id_int(host) if isinstance(host, Wall) else None
    if host_id is None:
        out = dict(analysis)
        out[u"scoped_to_host"] = False
        out[u"scope_note"] = u"Sin host muro; se usan todas las capas del GUID."
        return out

    layers_out = []
    for ly in analysis.get(u"layers") or []:
        rbs_ok = []
        for rb in ly.get(u"rebars") or []:
            h = _host_wall_of(rb, doc)
            if isinstance(h, Wall) and wall_id_int(h) == host_id:
                rbs_ok.append(rb)
        if not rbs_ok:
            continue
        ly2 = dict(ly)
        ly2[u"rebars"] = rbs_ok
        # Recalcular centroide solo con rebars del host (evita copias GUID).
        try:
            xs, ys, zs, n = 0.0, 0.0, 0.0, 0
            for rb in rbs_ok:
                p = _bbox_center(rb)
                if p is None:
                    continue
                xs += float(p.X)
                ys += float(p.Y)
                zs += float(p.Z)
                n += 1
            if n > 0:
                ly2[u"centroid"] = XYZ(xs / n, ys / n, zs / n)
        except Exception:
            pass
        try:
            ly2[u"qty"] = max(
                1,
                max(
                    int(getattr(r, u"NumberOfBarPositions", 1) or 1)
                    for r in rbs_ok
                ),
            )
        except Exception:
            pass
        layers_out.append(ly2)

    out = dict(analysis)
    out[u"host"] = host
    if layers_out:
        out[u"layers"] = layers_out
        out[u"scoped_to_host"] = True
        out[u"scope_note"] = (
            u"Capas acotadas al host de la seed (Id {0}).".format(host_id)
        )
        out[u"ok"] = True
        out[u"error"] = None
    else:
        # Fallback: GUID global (copias pueden contaminar).
        out[u"scoped_to_host"] = False
        out[u"scope_note"] = (
            u"No hubo capas en el host de la seed; se usa el GUID completo."
        )
    return out


def ultima_capa_index_desde_guid(analysis):
    """
    Mayor índice 0-based de ``Armadura_Capa`` en las capas del análisis.

    Returns:
        int >= -1  (−1 = no hay capas → la primera nueva será 0 / 1ºC.)
    """
    return int(_max_capa_index((analysis or {}).get(u"layers") or []))


def base_offset_despues_ultima_capa(analysis):
    """Índice 0-based donde empieza la primera capa nueva (última+1)."""
    last = ultima_capa_index_desde_guid(analysis)
    return max(0, int(last) + 1)


def detect_extremo_from_seed(doc, analysis):
    """
    Infiera ``inicio`` / ``fin`` comparando el punto seed con los extremos
    de la LocationCurve del muro host (P0=inicio, P1=fin).
    """
    seed = (analysis or {}).get(u"seed")
    host = (analysis or {}).get(u"host")
    if not isinstance(host, Wall):
        host = _host_wall_of(seed, doc)
    if not isinstance(host, Wall):
        return CABEZAL_EXTREMO_INICIO, u"Sin host muro; se asume Inicio."

    pt = _seed_point(analysis)
    if pt is None:
        return CABEZAL_EXTREMO_INICIO, u"Sin posición de seed; se asume Inicio."

    try:
        lc = host.Location
        if not isinstance(lc, LocationCurve) or lc.Curve is None:
            return CABEZAL_EXTREMO_INICIO, u"Host sin LocationCurve; se asume Inicio."
        curve = lc.Curve
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        return CABEZAL_EXTREMO_INICIO, u"No se pudo leer extremos del muro; se asume Inicio."

    try:
        # Distancia en planta (XY) al P0 / P1.
        d0 = (float(pt.X) - float(p0.X)) ** 2 + (float(pt.Y) - float(p0.Y)) ** 2
        d1 = (float(pt.X) - float(p1.X)) ** 2 + (float(pt.Y) - float(p1.Y)) ** 2
        if d1 < d0:
            return CABEZAL_EXTREMO_FIN, u"Detectado extremo Final (cerca de P1)."
        return CABEZAL_EXTREMO_INICIO, u"Detectado extremo Inicio (cerca de P0)."
    except Exception:
        return CABEZAL_EXTREMO_INICIO, u"Error midiendo extremos; se asume Inicio."


def _extremo_label(extremo):
    if extremo == CABEZAL_EXTREMO_FIN:
        return u"Final"
    return u"Inicio"


def _wall_by_id(walls, wid):
    try:
        wid_i = int(wid)
    except Exception:
        return None
    for w in walls or []:
        if wall_id_int(w) == wid_i:
            return w
    return None




def ultima_capa_layer(analysis):
    """Dict de la capa con mayor ``Armadura_Capa``, o None."""
    layers = (analysis or {}).get(u"layers") or []
    if not layers:
        return None
    try:
        return max(layers, key=lambda ly: int(ly.get(u"index", -1)))
    except Exception:
        return layers[-1]


def _ft_to_mm(ft):
    try:
        from Autodesk.Revit.DB import UnitTypeId, UnitUtils

        return float(
            UnitUtils.ConvertFromInternalUnits(float(ft), UnitTypeId.Millimeters)
        )
    except Exception:
        return float(ft) * 304.8


def _xyz_xy(pt):
    if pt is None:
        return None
    try:
        return XYZ(float(pt.X), float(pt.Y), 0.0)
    except Exception:
        return None


def _normalize_xy(vec):
    if vec is None:
        return None
    try:
        x = float(vec.X)
        y = float(vec.Y)
        L = (x * x + y * y) ** 0.5
        if L < 1e-12:
            return None
        return XYZ(x / L, y / L, 0.0)
    except Exception:
        return None


def _stack_axis_from_guid_layers(analysis, wall, extremo):
    """
    Eje unitario XY del apilado real de capas del GUID (capa baja → última).

    Si solo hay una capa o falla, usa ``vector_longitudinal`` del extremo.
    """
    layers = list((analysis or {}).get(u"layers") or [])
    if len(layers) >= 2:
        try:
            layers_sorted = sorted(layers, key=lambda ly: int(ly.get(u"index", 0)))
        except Exception:
            layers_sorted = layers
        c0 = layers_sorted[0].get(u"centroid")
        c1 = layers_sorted[-1].get(u"centroid")
        if c0 is None:
            r0 = (layers_sorted[0].get(u"rebars") or [None])[0]
            c0 = _rebar_stem_point(r0) or _bbox_center(r0)
        if c1 is None:
            r1 = (layers_sorted[-1].get(u"rebars") or [None])[0]
            c1 = _rebar_stem_point(r1) or _bbox_center(r1)
        if c0 is not None and c1 is not None:
            try:
                axis = _normalize_xy(
                    XYZ(float(c1.X) - float(c0.X), float(c1.Y) - float(c0.Y), 0.0)
                )
                if axis is not None:
                    # Orientar hacia el interior (mismo sentido que vector_long).
                    try:
                        geom = _wall_longitudinal_at_extremo(wall, extremo)
                        vl = geom.get(u"vector_longitudinal") if geom else None
                        if vl is not None and float(axis.DotProduct(vl)) < 0.0:
                            axis = XYZ(-float(axis.X), -float(axis.Y), 0.0)
                    except Exception:
                        pass
                    return axis
            except Exception:
                pass
    try:
        geom = _wall_longitudinal_at_extremo(wall, extremo)
        if geom and geom.get(u"vector_longitudinal") is not None:
            return _normalize_xy(geom[u"vector_longitudinal"])
    except Exception:
        pass
    return None


def _rebar_stem_point(rebar):
    """
    Punto medio del tramo vertical más largo (eje de capa, sin ganchos).
    """
    if rebar is None or not isinstance(rebar, Rebar):
        return None
    best = None
    best_len = -1.0
    try:
        curves = rebar.GetCenterlineCurves(
            False, False, False,
            MultiplanarOption.IncludeAllMultiplanarCurves,
            0,
        )
    except Exception:
        curves = None
    for c in curves or []:
        if c is None:
            continue
        try:
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
            dx = float(p1.X) - float(p0.X)
            dy = float(p1.Y) - float(p0.Y)
            dz = float(p1.Z) - float(p0.Z)
            horiz = (dx * dx + dy * dy) ** 0.5
            # Tramo vertical: domina Z
            if abs(dz) < max(horiz * 2.0, 1e-3):
                continue
            length = float(c.Length)
            if length > best_len:
                best_len = length
                try:
                    best = c.Evaluate(0.5, True)
                except Exception:
                    best = XYZ(
                        0.5 * (float(p0.X) + float(p1.X)),
                        0.5 * (float(p0.Y) + float(p1.Y)),
                        0.5 * (float(p0.Z) + float(p1.Z)),
                    )
        except Exception:
            continue
    if best is not None:
        return best
    return _bbox_center(rebar)


def _depth_along_axis_mm(pt, origin_pt, axis_xy):
    """Proyección XY de ``pt`` sobre eje unitario desde ``origin_pt`` (mm)."""
    if pt is None or origin_pt is None or axis_xy is None:
        return None
    try:
        dx = float(pt.X) - float(origin_pt.X)
        dy = float(pt.Y) - float(origin_pt.Y)
        depth_ft = dx * float(axis_xy.X) + dy * float(axis_xy.Y)
        return _ft_to_mm(depth_ft)
    except Exception:
        return None


def _last_capa_origin_and_axis(analysis, wall, extremo):
    """
    Origen = punto de la última capa; eje = apilado GUID (o longitudinal muro).

    Returns:
        (origin_xyz, axis_xy) o (None, None)
    """
    axis = _stack_axis_from_guid_layers(analysis, wall, extremo)
    last = ultima_capa_layer(analysis)
    origin = None
    if last is not None:
        for rb in last.get(u"rebars") or []:
            origin = _rebar_stem_point(rb)
            if origin is not None:
                break
        if origin is None:
            origin = last.get(u"centroid")
    if origin is None:
        return None, axis
    return origin, axis


def measure_last_capa_depth_mm(analysis, wall, extremo):
    """
    Compat: profundidad de la última capa desde la cara del extremo.

    Preferir medición relativa en el post-proceso (origen = última capa).
    """
    origin, axis = _last_capa_origin_and_axis(analysis, wall, extremo)
    if origin is None:
        return None
    # Desde cara del extremo, para mensajes / fallback de create.
    try:
        geom = _wall_longitudinal_at_extremo(wall, extremo)
        pt_ext = geom.get(u"pt_extremo") if geom else None
        vl = _normalize_xy(geom.get(u"vector_longitudinal")) if geom else axis
        return _depth_along_axis_mm(origin, pt_ext, vl or axis)
    except Exception:
        return 0.0


def _with_placement_from_last_capa(
    base_offset, last_depth_mm, spacing_mm, fn,
):
    """
    Geometría local 0..n−1; stamp ``base_offset + k``.
    El hueco exacto lo fija el post-proceso sobre el eje real del GUID.
    """
    import armado_muros_cabezal as cab

    _ = last_depth_mm, spacing_mm
    try:
        offset = max(0, int(base_offset))
    except Exception:
        offset = 0

    orig_enum = cab.cabezal_enum_layer_index

    def _enum(layer_index, n_capas=None, extremo=None, ex_cfg=None):
        try:
            li = int(round(float(layer_index))) + offset
        except Exception:
            li = offset
        return orig_enum(li, n_capas=n_capas, extremo=extremo, ex_cfg=ex_cfg)

    cab.cabezal_enum_layer_index = _enum
    try:
        return fn()
    finally:
        cab.cabezal_enum_layer_index = orig_enum


def _move_rebar(doc, rebar_or_id, delta_xyz):
    """Trasladar rebar; si MoveElement falla, Copy+Delete."""
    if doc is None or delta_xyz is None:
        return None
    rid = rebar_or_id
    try:
        if hasattr(rebar_or_id, u"Id"):
            rid = rebar_or_id.Id
    except Exception:
        pass
    if rid is None:
        return None
    try:
        ElementTransformUtils.MoveElement(doc, rid, delta_xyz)
        return rid
    except Exception:
        pass
    try:
        copied = ElementTransformUtils.CopyElement(doc, rid, delta_xyz)
        new_id = None
        if copied is not None and len(list(copied)) > 0:
            new_id = list(copied)[0]
        doc.Delete(rid)
        return new_id
    except Exception:
        return None


def _postprocess_capas_despues_ultima(
    doc, walls, cab_res, base_offset, spacing_mm, guid,
    analysis=None, extremo=None, last_depth_mm=None,
):
    """
    1) ``Armadura_Capa`` = base_offset + k.
    2) Traslada el lote nuevo sobre el **eje real del apilado GUID** para que
       la distancia desde la última capa existente sea exactamente ``spacing_UI``.
       Varias pasadas hasta error < 1 mm.
    """
    meta_list = list((cab_res or {}).get(u"rebars_longitudinales_tag_meta") or [])
    if not meta_list or doc is None:
        return 0, 0, None

    try:
        spacing = float(spacing_mm)
    except Exception:
        spacing = float(CABEZAL_LAYER_PITCH_MM)
    if spacing < 1.0:
        spacing = float(CABEZAL_LAYER_PITCH_MM)
    try:
        offset = max(0, int(base_offset))
    except Exception:
        offset = 0

    n_stamp = 0
    n_move = 0
    gap_after = None
    # Copy+Delete puede cambiar ElementId: propagar a _seg_jobs_all (lap detail).
    global_id_map = {}
    _rv = _pyrevit_revit()
    # Excepción → rollback de esta txn (y del TransactionGroup padre si aplica).
    with _rv.Transaction(_TXN_NAME + u" — spacing desde última capa", doc):
        groups = {}
        for meta in meta_list:
            key = (meta.get(u"wid"), meta.get(u"extremo") or extremo)
            groups.setdefault(key, []).append(meta)

        for (wid, ex), metas in groups.items():
            wall = _wall_by_id(walls, wid)
            ex_use = ex or extremo or CABEZAL_EXTREMO_INICIO

            # Actualizar ids por si Copy+Delete cambia elementos
            id_map = {}

            for meta in metas:
                rid = meta.get(u"rebar_id")
                if rid is None:
                    continue
                try:
                    rb = doc.GetElement(rid)
                except Exception:
                    rb = None
                if rb is None:
                    continue
                stamped = _parse_capa_index(get_armadura_capa(rb))
                try:
                    meta_li = int(meta.get(u"layer_index", 0))
                except Exception:
                    meta_li = 0
                if stamped is None:
                    stamped = meta_li
                if stamped < offset:
                    local_i = stamped
                else:
                    local_i = stamped - offset
                target_li = offset + max(0, int(local_i))
                try:
                    if set_armadura_capa_desde_layer(rb, target_li):
                        n_stamp += 1
                except Exception:
                    pass
                try:
                    stamp_armadura_conjunto_guid(rb, conjunto_guid=guid)
                except Exception:
                    pass

            if wall is None or analysis is None:
                continue

            origin, axis = _last_capa_origin_and_axis(analysis, wall, ex_use)
            if origin is None or axis is None:
                continue

            # Identificar meta de la primera capa nueva (menor índice local)
            first_meta = None
            first_local = None
            for meta in metas:
                try:
                    li = int(meta.get(u"layer_index", 0))
                except Exception:
                    li = 0
                local = li - offset if li >= offset else li
                if first_local is None or local < first_local:
                    first_local = local
                    first_meta = meta
            if first_meta is None:
                continue

            def _rid(meta):
                rid = meta.get(u"rebar_id")
                return id_map.get(rid, rid)

            # Hasta 3 pasadas de corrección sobre el eje del GUID
            for _pass in range(3):
                try:
                    rb0 = doc.GetElement(_rid(first_meta))
                except Exception:
                    rb0 = None
                if rb0 is None:
                    break
                p0 = _rebar_stem_point(rb0)
                # Distancia desde la última capa existente a lo largo del stack
                gap = _depth_along_axis_mm(p0, origin, axis)
                if gap is None:
                    break
                delta_mm = float(spacing) - float(gap)
                gap_after = float(gap)
                if abs(delta_mm) < 1.0:
                    break
                try:
                    delta_xyz = XYZ(
                        float(axis.X) * _mm_to_internal(delta_mm),
                        float(axis.Y) * _mm_to_internal(delta_mm),
                        0.0,
                    )
                except Exception:
                    break
                for meta in metas:
                    rid = _rid(meta)
                    if rid is None:
                        continue
                    new_id = _move_rebar(doc, rid, delta_xyz)
                    if new_id is not None:
                        n_move += 1
                        if new_id != rid:
                            old_meta_id = meta.get(u"rebar_id")
                            id_map[old_meta_id] = new_id
                            meta[u"rebar_id"] = new_id
                            # Encadenar si el id ya venía de un Copy previo.
                            global_id_map[rid] = new_id
                            if old_meta_id is not None and old_meta_id != rid:
                                global_id_map[old_meta_id] = new_id
                            # Remapear claves previas que apuntaban a rid.
                            for k, v in list(global_id_map.items()):
                                if v == rid:
                                    global_id_map[k] = new_id
                try:
                    doc.Regenerate()
                except Exception:
                    pass
                try:
                    rb0 = doc.GetElement(_rid(first_meta))
                    p0 = _rebar_stem_point(rb0)
                    gap_after = _depth_along_axis_mm(p0, origin, axis)
                except Exception:
                    gap_after = float(spacing)

    _sync_seg_jobs_after_capas_move(doc, cab_res, global_id_map)
    return n_stamp, n_move, gap_after


def _sync_seg_jobs_after_capas_move(doc, cab_res, id_map=None):
    """
    Tras el spacing: actualizar ``rebar_id`` / ``bx``/``by`` en ``_seg_jobs_all``
    para que los lap detail de empalme coincidan con las barras ya movidas (como V3).
    """
    if not cab_res:
        return
    id_map = id_map or {}
    seg_jobs = cab_res.get(u"_seg_jobs_all") or []
    long_ids = cab_res.get(u"rebars_longitudinales_ids") or []

    if id_map:
        for sj in seg_jobs:
            rid = sj.get(u"rebar_id")
            if rid is None:
                continue
            mapped = id_map.get(rid)
            # Encadenar si hubo varios Copy
            seen = set()
            while mapped is not None and mapped not in seen:
                seen.add(mapped)
                sj[u"rebar_id"] = mapped
                nxt = id_map.get(mapped)
                if nxt is None or nxt == mapped:
                    break
                mapped = nxt
        for i, rid in enumerate(list(long_ids)):
            mapped = id_map.get(rid)
            seen = set()
            while mapped is not None and mapped not in seen:
                seen.add(mapped)
                long_ids[i] = mapped
                nxt = id_map.get(mapped)
                if nxt is None or nxt == mapped:
                    break
                mapped = nxt
        cab_res[u"rebars_longitudinales_ids"] = long_ids

    if doc is None:
        return
    for sj in seg_jobs:
        rid = sj.get(u"rebar_id")
        if rid is None:
            continue
        try:
            rb = doc.GetElement(rid)
        except Exception:
            rb = None
        if rb is None or not isinstance(rb, Rebar):
            continue
        pt = _rebar_stem_point(rb)
        if pt is None:
            continue
        try:
            sj[u"bx"] = float(pt.X)
            sj[u"by"] = float(pt.Y)
        except Exception:
            pass


def _collect_bar_types(doc):
    out = []
    if doc is None:
        return out
    try:
        # OfClass(RebarBarType) + WhereElementIsElementType: filtro nativo C#.
        out = list(
            FilteredElementCollector(doc)
            .OfClass(RebarBarType)
            .WhereElementIsElementType()
        )
    except Exception:
        out = []

    def _diam(bt):
        try:
            return float(bt.BarNominalDiameter)
        except Exception:
            return 0.0

    out.sort(key=_diam)
    return out


def _bar_type_label(bt):
    if bt is None:
        return u"(sin tipo)"
    try:
        d = int(round(float(bt.BarNominalDiameter) * 304.8))
    except Exception:
        d = None
    name = u""
    try:
        name = _as_unicode(bt.Name).strip()
    except Exception:
        pass
    if d:
        return u"Ø{0} · {1}".format(d, name) if name else u"Ø{0}".format(d)
    return name or u"RebarBarType"


def _default_spacing_mm(analysis):
    layers = (analysis or {}).get(u"layers") or []
    vals = []
    for ly in layers:
        v = ly.get(u"offset_from_prev_mm")
        if v is None:
            continue
        try:
            iv = int(v)
        except Exception:
            continue
        if iv > 0:
            vals.append(iv)
    if vals:
        return int(round(sum(vals) / float(len(vals))))
    return int(CABEZAL_LAYER_PITCH_MM)


def _default_n_bars_and_type(analysis, doc):
    layers = (analysis or {}).get(u"layers") or []
    if not layers:
        return CABEZAL_MIN_BARRAS_POR_CAPA, None
    try:
        last = max(layers, key=lambda ly: int(ly.get(u"index", -1)))
    except Exception:
        last = layers[-1]
    n_bars = CABEZAL_MIN_BARRAS_POR_CAPA
    try:
        n_bars = int(last.get(u"qty") or n_bars)
    except Exception:
        pass
    n_bars = max(
        CABEZAL_MIN_BARRAS_POR_CAPA,
        min(CABEZAL_MAX_BARRAS_POR_CAPA, n_bars),
    )
    bar_type_id = None
    rbs = last.get(u"rebars") or []
    if rbs and doc is not None:
        try:
            tid = rbs[0].GetTypeId()
            if tid and tid != ElementId.InvalidElementId:
                bar_type_id = tid
        except Exception:
            pass
    return n_bars, bar_type_id


def _wall_short_label(wall, index):
    wid = wall_id_int(wall) if wall is not None else u"?"
    name = u""
    try:
        name = _as_unicode(wall.Name).strip()
    except Exception:
        pass
    if name:
        return u"{0}. {1} (Id {2})".format(index + 1, name, wid)
    return u"{0}. Muro Id {1}".format(index + 1, wid)


def _default_troceo_modes(n_walls):
    """Igual que Armado Muros v3: todos en Auto (override None + geom)."""
    try:
        n = max(0, int(n_walls or 0))
    except Exception:
        n = 0
    return [TROCEO_AUTO] * n


def _auto_troceo_flags_v3(walls, stacked_layout, extremo):
    """Auto de creación v3: espesor / desfase U (no largo de fuste)."""
    try:
        return list(
            compute_auto_troceo_por_muro_flags(
                walls, stacked_layout, extremo,
            )
            or []
        )
    except Exception:
        return [False] * len(list(walls or []))


def _build_cabezal_por_muro(
    walls,
    extremo_activo,
    n_new_layers,
    n_bars,
    bar_type_id,
    spacing_mm,
    troceo_flags,
    confinement_type=None,
    conf_diam_mm=10.0,
    conf_spacing_mm=100.0,
    troceo_modes=None,
    stacked_layout=None,
):
    """
    Config cabezal: un extremo activo, troceo por muro (canvas), confinamiento.

    ``troceo_modes``: ``auto`` / ``tramo`` / ``cont`` alineado con ``walls``
    (misma semántica que Armado Muros v3). Si falta, se usa ``troceo_flags``.
    """
    cabezal_por_muro_id = {}
    n_new = max(1, int(n_new_layers))
    n_bars = max(
        CABEZAL_MIN_BARRAS_POR_CAPA,
        min(CABEZAL_MAX_BARRAS_POR_CAPA, int(n_bars)),
    )
    try:
        spacing = float(spacing_mm)
    except Exception:
        spacing = float(CABEZAL_LAYER_PITCH_MM)
    if spacing < 1.0:
        spacing = float(CABEZAL_LAYER_PITCH_MM)

    ctype = confinement_type or CABEZAL_CONFINEMENT_NONE
    try:
        c_diam = float(conf_diam_mm)
    except Exception:
        c_diam = 10.0
    try:
        c_sp = float(conf_spacing_mm)
    except Exception:
        c_sp = 100.0

    conf = normalize_cabezal_confinement(
        {
            u"type": ctype,
            u"stirrup_diam_mm": c_diam,
            u"stirrup_spacing_mm": c_sp,
        },
        n_capas=n_new,
        encuentro=False,
    )

    walls = list(walls or [])
    modes = list(troceo_modes or [])
    while len(modes) < len(walls):
        modes.append(TROCEO_AUTO)
    flags = list(troceo_flags or [])
    while len(flags) < len(walls):
        flags.append(False)

    for i, wall in enumerate(walls):
        if wall is None:
            continue
        wid = wall_id_int(wall)
        cfg = default_cabezal_muro_config()
        mode = modes[i] if i < len(modes) else TROCEO_AUTO
        flag_i = bool(flags[i]) if i < len(flags) else False
        # Override V3: Auto=None, Tramo=True, Cont.=False
        if mode == TROCEO_TRAMO:
            ov = True
            troceo_eff = True
        elif mode == TROCEO_CONT:
            ov = False
            troceo_eff = False
        else:
            ov = None
            troceo_eff = flag_i
        for ex in (CABEZAL_EXTREMO_INICIO, CABEZAL_EXTREMO_FIN):
            ex_cfg = cfg[ex]
            if ex != extremo_activo:
                ex_cfg[u"armado_activo"] = False
                continue
            ex_cfg[u"armado_activo"] = True
            ex_cfg[u"n_capas"] = n_new
            ex_cfg[u"layers"] = [
                default_cabezal_layer_config(n_bars=n_bars, bar_type_id=bar_type_id)
                for _ in range(n_new)
            ]
            ex_cfg[u"bar_type_id"] = bar_type_id
            ex_cfg[u"layer_spacing_mm"] = spacing
            ex_cfg[u"troceo_por_muro"] = bool(troceo_eff)
            ex_cfg[u"troceo_por_muro_override"] = ov
            # sync_troceo_effective rellena auto_geom / efectivo (como v3).
            ex_cfg[u"troceo_auto_geom"] = False
            ex_cfg[u"confinement"] = dict(conf)
            ex_cfg[u"segment_bar_type_ids"] = {}
            ex_cfg[u"post_encuentro_activo"] = False
        cabezal_por_muro_id[wid] = cfg

    # Resolver Auto con geometría (igual que v3).
    try:
        sync_troceo_effective_for_extremo(
            walls,
            cabezal_por_muro_id,
            stacked_layout,
            extremo_activo,
        )
    except Exception:
        pass

    # Tramo/Cont. del canvas prevalecen tras el sync.
    for i, wall in enumerate(walls):
        if wall is None:
            continue
        mode = modes[i] if i < len(modes) else TROCEO_AUTO
        if mode not in (TROCEO_TRAMO, TROCEO_CONT):
            continue
        try:
            wid = wall_id_int(wall)
            ex_cfg = (cabezal_por_muro_id.get(wid) or {}).get(extremo_activo) or {}
            if mode == TROCEO_TRAMO:
                ex_cfg[u"troceo_por_muro"] = True
                ex_cfg[u"troceo_por_muro_override"] = True
            else:
                ex_cfg[u"troceo_por_muro"] = False
                ex_cfg[u"troceo_por_muro_override"] = False
        except Exception:
            pass

    return cabezal_por_muro_id


def _ref_walls_troceo_from_cfg(walls, cabezal_por_muro_id, extremo):
    """Muros con empalme activo → ``ref_walls_troceo`` del pipeline cabezal."""
    out = []
    for wall in walls or []:
        if wall is None:
            continue
        try:
            wid = wall_id_int(wall)
            ex_cfg = (cabezal_por_muro_id.get(wid) or {}).get(extremo) or {}
            if _troceo_por_muro_from_extremo_cfg(ex_cfg):
                out.append(wall)
        except Exception:
            continue
    return out


def crear_capas_adicionales_en_muros(
    doc,
    uidoc,
    analysis,
    walls,
    extremo,
    n_new_layers,
    n_bars,
    bar_type_id,
    spacing_mm,
    troceo_flags,
    confinement_type=None,
    conf_diam_mm=10.0,
    conf_spacing_mm=100.0,
    troceo_modes=None,
    view_right_xy=None,
    terminacion_cabeza=None,
):
    """
    Crea capas longitudinales adicionales sobre ``walls`` (mismo GUID).
    Alternancia A/B siempre activa (paridad de capa GUID).
    ``terminacion_cabeza``: ``empotramiento`` | ``pata_l`` (barras de Z máxima).

    Returns:
        dict con ok, message, n_created, …
    """
    out = {
        u"ok": False,
        u"message": u"",
        u"n_created": 0,
        u"n_fail": 0,
        u"messages": [],
    }
    if doc is None:
        out[u"message"] = u"Documento no válido."
        return out
    if not analysis or not analysis.get(u"ok"):
        out[u"message"] = u"Análisis GUID no válido."
        return out
    gid = analysis.get(u"guid")
    if not gid:
        out[u"message"] = u"GUID vacío."
        return out
    walls = ordenar_muros_por_base_asc(list(walls or []))
    if not walls:
        out[u"message"] = u"No hay muros seleccionados."
        return out
    if extremo not in (CABEZAL_EXTREMO_INICIO, CABEZAL_EXTREMO_FIN):
        extremo = CABEZAL_EXTREMO_INICIO

    last_idx = ultima_capa_index_desde_guid(analysis)
    base_offset = base_offset_despues_ultima_capa(analysis)

    ref_wall = walls[0]
    last_depth_mm = measure_last_capa_depth_mm(analysis, ref_wall, extremo)
    if last_depth_mm is None and isinstance(analysis.get(u"host"), Wall):
        last_depth_mm = measure_last_capa_depth_mm(
            analysis, analysis.get(u"host"), extremo,
        )

    stacked = None
    try:
        stacked = compute_stacked_wall_layout(
            walls, view_right_xy=view_right_xy,
        )
    except Exception:
        stacked = None

    cabezal_por_muro_id = _build_cabezal_por_muro(
        walls,
        extremo,
        n_new_layers,
        n_bars,
        bar_type_id,
        spacing_mm,
        troceo_flags,
        confinement_type=confinement_type,
        conf_diam_mm=conf_diam_mm,
        conf_spacing_mm=conf_spacing_mm,
        troceo_modes=troceo_modes,
        stacked_layout=stacked,
    )

    # Conteo de empalmes efectivos (mismo criterio que v3 / fusion troceo_walls).
    n_emp = len(
        _ref_walls_troceo_from_cfg(walls, cabezal_por_muro_id, extremo)
    )

    # Alternancia A/B siempre activa (pares=A Z-base, impares=B +L), como v3.
    # Terminación cabeza: Empotramiento o Pata L (sin sonda de colisión).
    term = terminacion_cabeza or _TERMINACION_CABEZA_DEFAULT
    try:
        term_s = unicode(term).strip().lower()
    except Exception:
        term_s = _TERMINACION_CABEZA_DEFAULT
    if term_s not in (TOP_TERMINATION_EMPOTRAMIENTO, TOP_TERMINATION_PATA_L):
        term_s = _TERMINACION_CABEZA_DEFAULT
    try:
        AppDomain.CurrentDomain.SetData(
            _APPDOMAIN_EMPALME_ALTERNANCIA, None,
        )
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_TOP_TERMINATION, term_s)
        # Legacy: solo True si empotramiento (compat. helpers antiguos).
        AppDomain.CurrentDomain.SetData(
            _APPDOMAIN_SKIP_TOP_EMBED_COLLISION,
            True if term_s == TOP_TERMINATION_EMPOTRAMIENTO else None,
        )
    except Exception:
        pass

    iniciar_armadura_conjunto_guid_ejecucion(conjunto_guid=gid)
    cab_res = None
    n_stamp, n_move, gap_after = 0, 0, None
    fatal_ex = None
    try:
        try:
            iniciar_armadura_eje_ejecucion(uidoc=uidoc)
        except Exception:
            pass

        _rv = _pyrevit_revit()

        def _do_create():
            # Igual que Armado Muros v3: sin ref_walls_troceo legacy;
            # los cortes salen de troceo_por_muro → troceo_walls en fusión.
            # defer_etiquetado: lap detail + tags tras el post-proceso de spacing
            # (si se colocan antes, el Move desalinearía los Detail Items).
            # within_parent_transaction_group: el TransactionGroup de Capas
            # es dueño del Undo (Assimilate → un solo Deshacer).
            return aplicar_cabezales_muros(
                doc,
                walls,
                cabezal_por_muro_id,
                bar_type_fallback=bar_type_id,
                ref_walls_troceo=None,
                uidoc=uidoc,
                defer_etiquetado=True,
                within_parent_transaction_group=True,
            )

        # Grupo asimilado: todas las txn internas → un solo Undo.
        # Excepción en geometría / etiquetas → RollBack del grupo (doc limpio).
        with _rv.TransactionGroup(_TXN_NAME, doc, assimilate=True):
            cab_res = _with_placement_from_last_capa(
                base_offset, last_depth_mm, spacing_mm, _do_create,
            )
            n_stamp, n_move, gap_after = _postprocess_capas_despues_ultima(
                doc,
                walls,
                cab_res,
                base_offset,
                spacing_mm,
                gid,
                analysis=analysis,
                extremo=extremo,
                last_depth_mm=last_depth_mm,
            )
            # Lap detail + etiquetas como V3, con XY ya corregida tras el spacing.
            if cab_res and cab_res.get(u"_defer_etiquetado"):
                _sync_seg_jobs_after_capas_move(doc, cab_res, None)
                cabezal_aplicar_etiquetado_pendiente(doc, cab_res, uidoc=uidoc)
    except Exception as ex:
        fatal_ex = ex
        cab_res = None
        n_stamp, n_move, gap_after = 0, 0, None
    finally:
        finalizar_armadura_conjunto_guid_ejecucion()
        try:
            AppDomain.CurrentDomain.SetData(_APPDOMAIN_EMPALME_ALTERNANCIA, None)
        except Exception:
            pass
        try:
            AppDomain.CurrentDomain.SetData(_APPDOMAIN_TOP_TERMINATION, None)
        except Exception:
            pass
        try:
            AppDomain.CurrentDomain.SetData(
                _APPDOMAIN_SKIP_TOP_EMBED_COLLISION, None,
            )
        except Exception:
            pass

    if fatal_ex is not None:
        out[u"ok"] = False
        out[u"n_created"] = 0
        out[u"n_fail"] = 1
        out[u"messages"] = [_as_unicode(fatal_ex)]
        out[u"message"] = (
            u"Error al crear capas (cambios revertidos): {0}"
            .format(_as_unicode(fatal_ex))
        )
        return out

    n_created = int((cab_res or {}).get(u"n_created", 0) or 0)
    n_fail = int((cab_res or {}).get(u"n_fail", 0) or 0)
    n_conf = int((cab_res or {}).get(u"n_confinement_created", 0) or 0)
    n_troceo_seg = int((cab_res or {}).get(u"n_troceo_segments", 0) or 0)
    n_emp_ok = int((cab_res or {}).get(u"n_empalme_markers_ok", 0) or 0)
    n_emp_fail = int((cab_res or {}).get(u"n_empalme_markers_fail", 0) or 0)
    msgs = list((cab_res or {}).get(u"messages") or [])
    if n_emp:
        msgs.append(
            u"Empalmes canvas: {0} muro(s) de corte.".format(n_emp)
        )
    if n_troceo_seg:
        msgs.append(u"Segmentos troceo: {0}.".format(n_troceo_seg))
    elif n_emp:
        msgs.append(
            u"Aviso: había muros de empalme pero no se generaron segmentos "
            u"de troceo (revisa fusión colineal / espesores)."
        )
    if n_emp_ok or n_emp_fail:
        msgs.append(
            u"Lap detail empalme: {0} ok, {1} fallo.".format(n_emp_ok, n_emp_fail)
        )
    if gap_after is not None:
        try:
            msgs.append(
                u"Hueco última→1ª nueva: {0:.1f} mm (UI={1:g} mm)."
                .format(float(gap_after), float(spacing_mm))
            )
        except Exception:
            pass
    if n_stamp or n_move:
        msgs.append(
            u"Post: stamp={0}, move={1}."
            .format(n_stamp, n_move)
        )
    if n_conf:
        msgs.append(u"Confinamiento: {0} elemento(s).".format(n_conf))
    out[u"n_created"] = n_created
    out[u"n_fail"] = n_fail
    out[u"n_empalme_walls"] = n_emp
    out[u"n_troceo_segments"] = n_troceo_seg
    out[u"messages"] = msgs
    out[u"last_capa_index"] = last_idx
    out[u"base_offset"] = base_offset
    out[u"last_depth_mm"] = last_depth_mm
    out[u"gap_after_mm"] = gap_after
    if n_created > 0:
        out[u"ok"] = True
        last_disp = last_idx + 1 if last_idx >= 0 else 0
        gap_txt = u"?"
        try:
            if gap_after is not None:
                gap_txt = u"{0:.0f}".format(float(gap_after))
        except Exception:
            pass
        out[u"message"] = (
            u"Creadas {0} barra(s) · tras {1}ºC. → desde {2}ºC. · "
            u"hueco {3} mm (UI {4:g}) · empalmes {5} · A/B · cabeza {6} · "
            u"extremo {7}."
            .format(
                n_created,
                last_disp,
                base_offset + 1,
                gap_txt,
                float(spacing_mm),
                n_emp,
                (
                    u"Pata L"
                    if term_s == TOP_TERMINATION_PATA_L
                    else u"Empotramiento"
                ),
                _extremo_label(extremo),
            )
        )
    else:
        detail = u"; ".join([_as_unicode(m) for m in msgs[:4] if m])
        out[u"message"] = detail or u"No se crearon barras."
    return out


# ── singleton / ExternalEvent ────────────────────────────────────────────────


def _get_active_controller():
    try:
        return AppDomain.CurrentDomain.GetData(_APPDOMAIN_CTRL_KEY)
    except Exception:
        return None


def _set_active_controller(ctrl):
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_CTRL_KEY, ctrl)
    except Exception:
        pass


def _clear_active_controller(ctrl=None):
    cur = _get_active_controller()
    if ctrl is None or cur is ctrl:
        try:
            AppDomain.CurrentDomain.SetData(_APPDOMAIN_CTRL_KEY, None)
        except Exception:
            pass
        try:
            AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
        except Exception:
            pass


def _get_active_window():
    ctrl = _get_active_controller()
    if ctrl is not None:
        try:
            return ctrl._win
        except Exception:
            pass
    try:
        return AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
    except Exception:
        return None


class _CrearHandler(IExternalEventHandler):
    def GetName(self):
        return _TXN_NAME

    def Execute(self, uiapp):
        # Leer siempre desde AppDomain: tras reload de scripts el ExternalEvent
        # puede seguir apuntando a este handler, pero el global del módulo nuevo
        # no es el mismo que el del módulo viejo.
        ctrl = _get_create_target()
        if ctrl is None:
            try:
                print(u"[CapasAdicionalesEnMuro] ExternalEvent sin target.")
            except Exception:
                pass
            return
        try:
            ctrl._execute_create(uiapp)
        except Exception as ex:
            try:
                ctrl._set_status(_as_unicode(ex))
            except Exception:
                pass
            _mostrar_aviso(
                uiapp,
                u"Error al crear capas adicionales en muro.",
                content=_as_unicode(ex),
            )


_CREATE_TARGET = None
_CREATE_EVENT = None
_CREATE_HANDLER = None


def _get_create_target():
    global _CREATE_TARGET
    if _CREATE_TARGET is not None:
        return _CREATE_TARGET
    try:
        return AppDomain.CurrentDomain.GetData(_APPDOMAIN_CREATE_TARGET_KEY)
    except Exception:
        return None


def _ensure_create_event():
    """Crear ExternalEvent solo en contexto API (al abrir la herramienta)."""
    global _CREATE_EVENT, _CREATE_HANDLER
    if _CREATE_EVENT is not None:
        return _CREATE_EVENT
    _CREATE_HANDLER = _CrearHandler()
    _CREATE_EVENT = ExternalEvent.Create(_CREATE_HANDLER)
    return _CREATE_EVENT


def _set_create_target(ctrl):
    global _CREATE_TARGET
    _CREATE_TARGET = ctrl
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_CREATE_TARGET_KEY, ctrl)
    except Exception:
        pass


def _clear_create_target(ctrl=None):
    global _CREATE_TARGET
    cur = _get_create_target()
    if ctrl is None or cur is ctrl or _CREATE_TARGET is ctrl:
        _CREATE_TARGET = None
        try:
            AppDomain.CurrentDomain.SetData(_APPDOMAIN_CREATE_TARGET_KEY, None)
        except Exception:
            pass


# ── UI (shell elevación + rail, estilo Armado Muros v3) ───────────────────────


def _build_capas_xaml():
    return _XAML_CAPAS.replace(u"__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML)


_XAML_CAPAS = u"""<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="Arainco"
  Height="940" Width="1220"
  MinHeight="700" MinWidth="960"
  ResizeMode="CanResize"
  WindowStartupLocation="Manual"
  WindowState="Maximized"
  Background="#071018"
  FontFamily="Segoe UI"
  FontSize="12"
  ShowInTaskbar="False">
  <Window.Resources>
__BIMTOOLS_DARK_STYLES__
  </Window.Resources>
  <Border Background="#071018" BorderBrush="#21465C" BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <StackPanel Grid.Row="0" Margin="0,0,0,8">
        <TextBlock x:Name="TxtTitle" Text="Arainco: Capas Adicionales en Muro"
                   Foreground="#E8F4F8" FontSize="18" FontWeight="Bold"/>
        <TextBlock x:Name="TxtSubtitle" Margin="0,6,0,0" Foreground="#95B8CC" TextWrapping="Wrap"
                   Text="Elevación + rail · empalmes Auto/Tramo/Cont. en el pie · mismo GUID."/>
      </StackPanel>

      <TextBlock x:Name="TxtInfoMuros" Grid.Row="1" Foreground="#64748b" FontSize="10"
                 Margin="0,0,0,10" TextWrapping="Wrap"
                 Text="Empalmes como Armado Muros v3: pie Auto→Tramo→Cont. Auto = espesor/desfase U. Alternancia A/B siempre activa."/>

      <Grid Grid.Row="2">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="380"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0" Background="#0a1620" BorderBrush="#21465C" BorderThickness="1"
                CornerRadius="4,0,0,4" Padding="0">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Border Grid.Row="0" Background="#0a1620" BorderBrush="#21465C"
                    BorderThickness="0,0,0,1" Padding="10,8">
              <TextBlock x:Name="TxtElevHeader" Foreground="#95B8CC" FontSize="11"
                         FontWeight="SemiBold" Text="Sección / elevación"/>
            </Border>
            <ScrollViewer x:Name="ScrMuros" Grid.Row="1"
                          VerticalScrollBarVisibility="Auto"
                          HorizontalScrollBarVisibility="Disabled">
              <Border Background="#0a1620" BorderBrush="Transparent" BorderThickness="0"
                      Padding="8,4,8,12">
                <Grid x:Name="GrdListaMuros" Background="Transparent"
                      ClipToBounds="False" SnapsToDevicePixels="True"
                      HorizontalAlignment="Stretch"/>
              </Border>
            </ScrollViewer>
          </Grid>
        </Border>

        <Border Grid.Column="1" Background="#0a1620" BorderBrush="#21465C" BorderThickness="1"
                CornerRadius="0,4,4,0" Padding="8,8">
          <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
            <StackPanel x:Name="PnlSectionRail">

              <Border Background="#0a1620" BorderBrush="#5BC0DE"
                      BorderThickness="1.5" CornerRadius="4" Padding="10" Margin="0,0,0,10">
                <StackPanel>
                  <DockPanel Margin="0,0,0,6" LastChildFill="True">
                    <Border DockPanel.Dock="Right" Background="#0E1B32"
                            BorderBrush="#5BC0DE" BorderThickness="1" CornerRadius="3"
                            Padding="6,2" Margin="8,0,0,0" VerticalAlignment="Center">
                      <TextBlock x:Name="TxtPillExtremo" Text="INICIO" Foreground="#5BC0DE"
                                 FontSize="10" FontWeight="SemiBold"/>
                    </Border>
                    <TextBlock Text="Capas adicionales" Foreground="#E8F4F8"
                               FontSize="12" FontWeight="SemiBold"
                               VerticalAlignment="Center" TextWrapping="NoWrap"/>
                  </DockPanel>
                  <TextBlock x:Name="TxtGuid" Foreground="#5BC0DE" FontSize="10"
                             FontFamily="Consolas" TextWrapping="Wrap" Margin="0,0,0,6"
                             Text="—"/>
                  <TextBlock x:Name="TxtLayersInfo" Foreground="#64748b" FontSize="10"
                             TextWrapping="Wrap" Margin="0,0,0,10" Text="—"/>

                  <Grid Margin="0,0,0,8">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0">
                      <TextBlock Text="Capas nuevas" Foreground="#95B8CC" FontSize="10" Margin="0,0,0,2"/>
                      <ComboBox x:Name="CmbNCapas" Style="{StaticResource Combo}" Height="28"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2">
                      <TextBlock Text="Barras / capa" Foreground="#95B8CC" FontSize="10" Margin="0,0,0,2"/>
                      <ComboBox x:Name="CmbNBars" Style="{StaticResource Combo}" Height="28"/>
                    </StackPanel>
                  </Grid>

                  <TextBlock Text="Tipo de barra" Foreground="#95B8CC" FontSize="10" Margin="0,0,0,2"/>
                  <ComboBox x:Name="CmbBarType" Style="{StaticResource Combo}" Height="28" Margin="0,0,0,8"/>

                  <TextBlock Text="Separación capas (mm)" Foreground="#95B8CC" FontSize="10" Margin="0,0,0,2"/>
                  <TextBox x:Name="TxtSpacing" Style="{StaticResource BimToolsTextBoxDark}"
                           Height="28" Margin="0,0,0,10"/>

                  <TextBlock Text="Terminación cabeza (Z máx.)" Foreground="#95B8CC" FontSize="10" Margin="0,0,0,2"/>
                  <ComboBox x:Name="CmbTerminacionCabeza" Style="{StaticResource Combo}"
                            Height="28" Margin="0,0,0,4"/>
                  <TextBlock Foreground="#64748b" FontSize="10" TextWrapping="Wrap" Margin="0,0,0,10"
                             Text="Barras superiores de cada capa: Empotramiento L(Ø) o Pata L. Empalmes Auto/Tramo/Cont. · A/B siempre activa."/>

                  <TextBlock Text="Confinamiento (capas nuevas)" Foreground="#95B8CC" FontSize="11"
                             FontWeight="SemiBold" Margin="0,0,0,4"/>
                  <ComboBox x:Name="CmbConfinement" Style="{StaticResource Combo}" Height="28" Margin="0,0,0,6"/>
                  <Grid Margin="0,0,0,4">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0">
                      <TextBlock Text="Ø estribo/traba (mm)" Foreground="#95B8CC" FontSize="10" Margin="0,0,0,2"/>
                      <TextBox x:Name="TxtConfDiam" Style="{StaticResource BimToolsTextBoxDark}"
                               Height="28" Text="10"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2">
                      <TextBlock Text="Espaciamiento (mm)" Foreground="#95B8CC" FontSize="10" Margin="0,0,0,2"/>
                      <TextBox x:Name="TxtConfSpacing" Style="{StaticResource BimToolsTextBoxDark}"
                               Height="28" Text="100"/>
                    </StackPanel>
                  </Grid>
                  <TextBlock x:Name="TxtConfHint" Foreground="#64748b" FontSize="10" TextWrapping="Wrap"
                             Text="Tipo 1–3 requieren ≥2 capas nuevas."/>
                </StackPanel>
              </Border>

            </StackPanel>
          </ScrollViewer>
        </Border>
      </Grid>

      <Grid Grid.Row="3" Margin="0,14,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" VerticalAlignment="Center" Margin="0,0,12,0">
          <TextBlock x:Name="TxtStatus" Foreground="#64748b" FontSize="10" TextWrapping="Wrap"/>
          <TextBlock Foreground="#64748b" FontSize="10" TextWrapping="Wrap" Margin="0,4,0,0"
                     Text="Creación: longitudinales → confinamiento. Mismo GUID · orden base→cima."/>
        </StackPanel>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnRepickWalls" Content="Cambiar muros"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="120" Margin="0,0,10,0"/>
          <Button x:Name="BtnCreate" Content="Crear capas"
                  Style="{StaticResource BtnPrimary}" MinWidth="160"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>"""


class CapasAdicionalesEnMuroWindow(object):
    def __init__(self, revit, analysis, walls, extremo, extremo_note):
        self._uiapp = revit
        self._uidoc = getattr(revit, u"ActiveUIDocument", None)
        self._doc = self._uidoc.Document if self._uidoc is not None else None
        self._analysis = analysis
        self._walls = list(walls or [])
        self._extremo = extremo or CABEZAL_EXTREMO_INICIO
        self._extremo_note = extremo_note or u""
        self._win = None
        self._bar_types = []
        self._closed = False
        self._elev_canvas = None
        self._elev_viewport_size = (0.0, 0.0)
        self._elev_size_wired = False
        self._selected_wall = 0
        self._selected_segment = 0
        self._stacked_layout = None
        self._closing_for_create = False
        n = len(self._walls)
        self._troceo_modes = _default_troceo_modes(n)
        self._confinement_values = [CABEZAL_CONFINEMENT_NONE]

        self._win = XamlReader.Parse(_build_capas_xaml())
        try:
            self._win.Title = WINDOW_CHROME_TITLE
        except Exception:
            pass

        sub = self._win.FindName(u"TxtSubtitle")
        if sub is not None:
            sub.Text = (
                u"Capas nuevas · empalmes en elevación · confinamiento · "
                u"mismo GUID · extremo auto desde seed"
            )

        try:
            self._ensure_stacked_layout()
        except Exception:
            self._stacked_layout = None

        self._populate_static()
        self._create_event = _ensure_create_event()
        _set_create_target(self)
        self._wire_events()
        try:
            self._redraw_elevation()
        except Exception:
            pass

    # ── elevación ────────────────────────────────────────────────────────────

    def _view_right_xy(self):
        """RightDirection XY de la vista activa (mismo eje que Armado Muros v3)."""
        doc = self._doc
        if doc is None:
            return None
        try:
            rd = doc.ActiveView.RightDirection
            vr_x, vr_y = float(rd.X), float(rd.Y)
            if (vr_x * vr_x + vr_y * vr_y) > 1e-9:
                return (vr_x, vr_y)
        except Exception:
            pass
        return None

    def _ensure_stacked_layout(self, force=False):
        """
        Layout U fiel al modelo: proyecta LocationCurve sobre el RightDirection
        de la vista (igual que ``ArmadoMurosPreviewWindow`` / Machones).
        """
        if (not force) and getattr(self, u"_stacked_layout", None) is not None:
            return self._stacked_layout
        walls = list(self._walls or [])
        if not walls:
            self._stacked_layout = None
            return None
        view_right_xy = self._view_right_xy()
        try:
            self._stacked_layout = compute_stacked_wall_layout(
                walls,
                view_right_xy=view_right_xy,
            )
        except Exception:
            self._stacked_layout = None
        return self._stacked_layout

    def _extremos_lados(self):
        """(extremo_izq, extremo_der) según P0/P1 proyectados en el eje de vista."""
        walls = list(self._walls or [])
        stacked = self._ensure_stacked_layout()
        if not walls or stacked is None:
            return CABEZAL_EXTREMO_INICIO, CABEZAL_EXTREMO_FIN
        # Misma referencia que v3: muro cima del stack (último base→cima).
        ri = len(walls) - 1
        try:
            return cabezal_extremos_en_lados_stacked(walls[ri], ri, stacked)
        except Exception:
            return CABEZAL_EXTREMO_INICIO, CABEZAL_EXTREMO_FIN

    def _build_wall_meta(self):
        """Meta base→cima con u_start/length_u del stacked layout (v3)."""
        out = []
        walls = list(self._walls or [])
        stacked = self._ensure_stacked_layout()
        st_items = (stacked or {}).get(u"items") or []
        g_min = float((stacked or {}).get(u"global_min", 0.0) or 0.0)
        g_span = float((stacked or {}).get(u"global_span", 0.0) or 0.0)

        for i, wall in enumerate(walls):
            thick = 200.0
            try:
                thick = float(obtener_espesor_muro_mm_approx(wall) or 200.0)
            except Exception:
                pass
            height_mm = 3000.0
            z_mm = 0.0
            length_ft = 1.0
            u_start = g_min
            length_u = None
            try:
                bb = wall.get_BoundingBox(None)
                if bb is not None:
                    z_mm = float(bb.Min.Z) * 304.8
                    height_mm = max(1.0, (float(bb.Max.Z) - float(bb.Min.Z)) * 304.8)
            except Exception:
                pass
            try:
                loc = wall.Location
                crv = getattr(loc, u"Curve", None)
                if crv is not None:
                    length_ft = max(float(crv.Length), 1e-6)
            except Exception:
                length_ft = 1.0
            try:
                if 0 <= i < len(st_items):
                    item = st_items[i]
                    u_start = float(item.get(u"u_start", g_min))
                    try:
                        length_u = float(item.get(u"length_u") or 0.0)
                    except Exception:
                        length_u = None
            except Exception:
                pass
            if length_u is None or length_u <= 1e-9:
                length_u = float(length_ft)
            try:
                from Autodesk.Revit.DB import BuiltInParameter

                ph = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
                if ph is not None and ph.HasValue:
                    hv = float(ph.AsDouble()) * 304.8
                    if hv > 1.0:
                        height_mm = hv
            except Exception:
                pass
            type_name = u"Muro"
            try:
                wt = self._doc.GetElement(wall.GetTypeId()) if self._doc else None
                type_name = unicode(getattr(wt, u"Name", None) or u"Muro")
            except Exception:
                pass
            name = u"W{0}".format(i + 1)
            try:
                name = unicode(getattr(wall, u"Name", None) or name)
            except Exception:
                pass
            try:
                label_txt = u"{0:.3f}".format(float(z_mm) / 1000.0)
            except Exception:
                label_txt = u"—"
            out.append({
                u"elem": wall,
                u"height_mm": float(height_mm),
                u"thick_mm": float(thick),
                u"z_mm": float(z_mm),
                u"length_ft": float(length_ft),
                u"length_u": float(length_u),
                u"u_start": float(u_start),
                u"type_name": type_name,
                u"name": name,
                u"label": label_txt,
                u"level_z_m": float(z_mm) / 1000.0,
                u"_g_min": g_min,
                u"_span_u": g_span,
            })
        return out

    def _ensure_elev_host(self):
        from System.Windows.Controls import Canvas, Grid, ColumnDefinition, RowDefinition
        from System.Windows import (
            GridLength,
            GridUnitType,
            HorizontalAlignment,
            VerticalAlignment,
            FrameworkElement,
        )
        from System.Windows.Media import Brushes

        root = self._win.FindName(u"GrdListaMuros") if self._win else None
        if root is None:
            return None
        existing = getattr(self, u"_elev_canvas", None)
        if existing is not None:
            try:
                if existing.Parent is not root:
                    raise RuntimeError(u"elev canvas orphaned")
                _ = existing.Children
                self._wire_elev_size_changed()
                return existing
            except Exception:
                self._elev_canvas = None

        try:
            root.Children.Clear()
            root.RowDefinitions.Clear()
            root.ColumnDefinitions.Clear()
        except Exception:
            pass
        try:
            root.ClearValue(FrameworkElement.WidthProperty)
            root.HorizontalAlignment = HorizontalAlignment.Stretch
        except Exception:
            pass

        rd = RowDefinition()
        rd.Height = GridLength(1.0, GridUnitType.Star)
        root.RowDefinitions.Add(rd)
        cd = ColumnDefinition()
        cd.Width = GridLength(1.0, GridUnitType.Star)
        root.ColumnDefinitions.Add(cd)

        canv = Canvas()
        canv.Background = Brushes.Transparent
        canv.ClipToBounds = False
        canv.HorizontalAlignment = HorizontalAlignment.Left
        canv.VerticalAlignment = VerticalAlignment.Top
        try:
            Grid.SetRow(canv, 0)
            Grid.SetColumn(canv, 0)
        except Exception:
            pass
        root.Children.Add(canv)
        self._elev_canvas = canv
        self._wire_elev_size_changed()
        return canv

    def _wire_elev_size_changed(self):
        if getattr(self, u"_elev_size_wired", False):
            return
        if self._win is None:
            return
        scr = self._win.FindName(u"ScrMuros")
        if scr is None:
            return
        try:
            from System.Windows import SizeChangedEventHandler
        except Exception:
            return

        def _on_size(sender, args):
            try:
                nw = float(sender.ActualWidth or 0.0)
                nh = float(sender.ActualHeight or 0.0)
            except Exception:
                return
            if nw < 40.0 or nh < 40.0:
                return
            prev = getattr(self, u"_elev_viewport_size", (0.0, 0.0)) or (0.0, 0.0)
            if abs(nw - float(prev[0])) < 2.0 and abs(nh - float(prev[1])) < 2.0:
                return
            self._elev_viewport_size = (nw, nh)
            try:
                self._redraw_elevation()
            except Exception:
                pass

        try:
            scr.SizeChanged += SizeChangedEventHandler(_on_size)
            self._elev_size_wired = True
        except Exception:
            pass

    def _modes_for_extremo(self, extremo):
        """
        Capas: un solo juego de empalmes (extremo seed).
        Ambas bandas del canvas muestran el mismo estado para que los puntos
        de empalme se definan y lean en la elevación (como v3).
        """
        _ = extremo
        n = len(self._walls or [])
        modes = list(self._troceo_modes or [])
        while len(modes) < n:
            modes.append(TROCEO_AUTO)
        return modes[:n]

    def _segments_for_modes(self, modes, meta):
        n = len(meta or [])
        if n <= 0:
            return []
        autos = _auto_troceo_flags_v3(
            self._walls,
            self._ensure_stacked_layout(),
            self._extremo,
        )
        while len(autos) < n:
            autos.append(False)
        base_i = lowest_z_meta_index(meta)
        if base_i is None:
            base_i = 0
        emp = empalme_indices_from_modes(modes, autos, base_index=base_i)
        return build_bar_segments(n, emp)

    def _long_diam_mm(self):
        form = None
        try:
            form = self._read_form()
        except Exception:
            form = None
        bt_id = (form or {}).get(u"bar_type_id")
        if bt_id is not None and self._doc is not None:
            try:
                bt = self._doc.GetElement(bt_id)
                if bt is not None:
                    return float(bt.BarNominalDiameter) * 304.8
            except Exception:
                pass
        try:
            layers = (self._analysis or {}).get(u"layers") or []
            if layers:
                d = layers[-1].get(u"diameter_mm")
                if d is not None:
                    return float(d)
        except Exception:
            pass
        return 16.0

    def _conf_by_wall(self):
        ctype, diam, sp = self._read_confinement()
        entry = {u"diam_mm": float(diam), u"spacing_mm": float(sp)}
        if ctype == CABEZAL_CONFINEMENT_NONE:
            entry = {u"diam_mm": 10.0, u"spacing_mm": 200.0}
        return [dict(entry) for _ in (self._walls or [])]

    def _redraw_elevation(self):
        canv = self._ensure_elev_host()
        if canv is None:
            return
        try:
            import armado_muros_v3_elevation as elev
        except Exception:
            return

        meta = self._build_wall_meta()
        n = len(meta)
        if n <= 0:
            try:
                canv.Children.Clear()
            except Exception:
                pass
            return

        modes_ini = self._modes_for_extremo(CABEZAL_EXTREMO_INICIO)
        modes_fin = self._modes_for_extremo(CABEZAL_EXTREMO_FIN)
        segs_ini = self._segments_for_modes(modes_ini, meta)
        segs_fin = self._segments_for_modes(modes_fin, meta)
        active = self._extremo or CABEZAL_EXTREMO_INICIO
        modes = modes_fin if active == CABEZAL_EXTREMO_FIN else modes_ini
        segs = segs_fin if active == CABEZAL_EXTREMO_FIN else segs_ini
        sel_wall = int(getattr(self, u"_selected_wall", 0) or 0)
        if sel_wall < 0 or sel_wall >= n:
            sel_wall = 0
        sel_seg = int(getattr(self, u"_selected_segment", 0) or 0)

        vw = vh = 0.0
        try:
            scr = self._win.FindName(u"ScrMuros")
            if scr is not None:
                vw = float(scr.ActualWidth or 0.0)
                vh = float(scr.ActualHeight or 0.0)
                if vw < 40.0:
                    vw = float(scr.ViewportWidth or 0.0)
                if vh < 40.0:
                    vh = float(scr.ViewportHeight or 0.0)
        except Exception:
            vw = vh = 0.0
        prev = getattr(self, u"_elev_viewport_size", (0.0, 0.0)) or (0.0, 0.0)
        if vw >= 40.0 and vh >= 40.0:
            self._elev_viewport_size = (vw, vh)
        else:
            if float(prev[0]) >= 40.0 and float(prev[1]) >= 40.0:
                vw, vh = float(prev[0]), float(prev[1])

        def _noop_sel(sid):
            try:
                self._selected_segment = int(sid)
            except Exception:
                pass
            try:
                self._redraw_elevation()
            except Exception:
                pass

        def _on_select_wall(wi):
            """Clic en fuste: seleccionar + alternar empalme (Tramo↔Cont.) en canvas."""
            try:
                wi = int(wi)
            except Exception:
                return
            self._selected_wall = wi
            try:
                self._toggle_empalme_wall(wi)
            except Exception:
                try:
                    self._redraw_elevation()
                except Exception:
                    pass

        ex_left, ex_right = self._extremos_lados()

        elev.redraw_elevation(
            canv,
            meta,
            modes,
            segs,
            sel_seg,
            self._long_diam_mm(),
            self._on_cycle_wall,
            _noop_sel,
            selected_wall=sel_wall,
            selected_walls=[sel_wall],
            rail_focus=u"muro",
            conf_by_wall=self._conf_by_wall(),
            on_select_wall=_on_select_wall,
            cfg_by_segment=None,
            viewport_w=vw,
            viewport_h=vh,
            on_draw_extremo_marks=None,
            segments_inicio=segs_ini,
            selected_segment_inicio=sel_seg if active == CABEZAL_EXTREMO_INICIO else 0,
            on_select_segment_inicio=_noop_sel,
            segments_fin=segs_fin,
            selected_segment_fin=sel_seg if active == CABEZAL_EXTREMO_FIN else 0,
            on_select_segment_fin=_noop_sel,
            troceo_modes_inicio=modes_ini,
            troceo_modes_fin=modes_fin,
            on_cycle_wall_inicio=self._on_cycle_wall_inicio,
            on_cycle_wall_fin=self._on_cycle_wall_fin,
            active_extremo=active,
            extremo_left=ex_left,
            extremo_right=ex_right,
            show_tramo_bands=True,
            show_pie_controls=True,
        )

    def _base_wall_index(self):
        meta = self._build_wall_meta()
        base_i = lowest_z_meta_index(meta) if meta else 0
        if base_i is None:
            base_i = 0
        return int(base_i)

    def _toggle_empalme_wall(self, wi):
        """
        Define punto de empalme en canvas: clic fuste / pie.
        Base → sin empalme. Resto: Tramo ↔ Cont. (manual).
        """
        try:
            wi = int(wi)
        except Exception:
            return
        n = len(self._walls or [])
        if wi < 0 or wi >= n:
            return
        if wi == self._base_wall_index():
            self._set_status(u"Muro base: sin empalme.")
            try:
                self._redraw_elevation()
            except Exception:
                pass
            return
        modes = list(self._troceo_modes or [])
        while len(modes) < n:
            modes.append(TROCEO_AUTO)
        cur = modes[wi] or TROCEO_AUTO
        # Clic en fuste: forzar manual Tramo/Cont. (punto ON/OFF).
        if cur == TROCEO_TRAMO:
            modes[wi] = TROCEO_CONT
        else:
            modes[wi] = TROCEO_TRAMO
        self._troceo_modes = modes
        self._selected_wall = wi
        try:
            self._redraw_elevation()
        except Exception:
            pass
        self._set_status(
            u"Empalme muro {0}: {1} (canvas)".format(wi + 1, modes[wi])
        )

    def _on_cycle_wall(self, wi, extremo=None):
        """Pie Auto→Tramo→Cont. en elevación (define puntos de empalme)."""
        _ = extremo  # Capas: un solo stack de empalmes; cualquier pie edita.
        try:
            wi = int(wi)
        except Exception:
            return
        n = len(self._walls or [])
        if wi < 0 or wi >= n:
            return
        if wi == self._base_wall_index():
            return
        modes = list(self._troceo_modes or [])
        while len(modes) < n:
            modes.append(TROCEO_AUTO)
        modes[wi] = cycle_troceo_mode(modes[wi])
        self._troceo_modes = modes
        self._selected_wall = wi
        try:
            self._redraw_elevation()
        except Exception:
            pass
        self._set_status(
            u"Empalme muro {0}: {1}".format(wi + 1, modes[wi])
        )

    def _on_cycle_wall_inicio(self, wi):
        self._on_cycle_wall(wi, CABEZAL_EXTREMO_INICIO)

    def _on_cycle_wall_fin(self, wi):
        self._on_cycle_wall(wi, CABEZAL_EXTREMO_FIN)

    def _read_troceo_flags(self):
        n = len(self._walls or [])
        modes = list(self._troceo_modes or [])
        while len(modes) < n:
            modes.append(TROCEO_AUTO)
        autos = _auto_troceo_flags_v3(
            self._walls,
            self._ensure_stacked_layout(),
            self._extremo,
        )
        while len(autos) < n:
            autos.append(False)
        flags = []
        for i in range(n):
            if i == 0:
                flags.append(False)
                continue
            auto = bool(autos[i]) if i < len(autos) else False
            flags.append(bool(effective_empalme(modes[i], auto)))
        return flags

    # ── rail / formulario ────────────────────────────────────────────────────

    def _populate_static(self):
        analysis = self._analysis or {}
        guid_tb = self._win.FindName(u"TxtGuid")
        if guid_tb is not None:
            gid = _as_unicode(analysis.get(u"guid") or u"—")
            if len(gid) > 48:
                guid_tb.Text = gid[:24] + u"…" + gid[-12:]
            else:
                guid_tb.Text = gid

        pill = self._win.FindName(u"TxtPillExtremo")
        if pill is not None:
            if self._extremo == CABEZAL_EXTREMO_FIN:
                pill.Text = u"TERMINO"
                try:
                    from System.Windows.Media import SolidColorBrush, Color

                    pill.Foreground = SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80))
                except Exception:
                    pass
            else:
                pill.Text = u"INICIO"

        hdr = self._win.FindName(u"TxtElevHeader")
        if hdr is not None:
            hdr.Text = u"Sección / elevación · {0}".format(_extremo_label(self._extremo))

        info_muros = self._win.FindName(u"TxtInfoMuros")
        if info_muros is not None:
            info_muros.Text = (
                u"{0} muro(s) · extremo {1} · {2}\n"
                u"Empalmes en elevación: pie Auto→Tramo→Cont. o clic en fuste."
                .format(
                    len(self._walls),
                    _extremo_label(self._extremo),
                    self._extremo_note or u"",
                )
            )

        layers = analysis.get(u"layers") or []
        max_idx = ultima_capa_index_desde_guid(analysis)
        next_capa = base_offset_despues_ultima_capa(analysis) + 1
        info = self._win.FindName(u"TxtLayersInfo")
        if info is not None:
            parts = []
            for ly in layers:
                try:
                    parts.append(
                        u"{0}ºC. ({1}×Ø{2})".format(
                            int(ly.get(u"display") or (int(ly.get(u"index")) + 1)),
                            int(ly.get(u"qty") or 0),
                            ly.get(u"diameter_mm") or u"?",
                        )
                    )
                except Exception:
                    pass
            info.Text = (
                u"Última capa GUID: {0}ºC. · nuevas desde {1}ºC. "
                u"({2} en host)\n{3}"
                .format(
                    max_idx + 1 if max_idx >= 0 else 0,
                    next_capa,
                    len(layers),
                    u" · ".join(parts) if parts else u"(sin detalle)",
                )
            )

        cmb_n = self._win.FindName(u"CmbNCapas")
        if cmb_n is not None:
            for i in range(1, 7):
                cmb_n.Items.Add(unicode(i))
            cmb_n.SelectedIndex = 0
            try:
                cmb_n.SelectionChanged += self._on_n_capas_changed
            except Exception:
                pass

        cmb_b = self._win.FindName(u"CmbNBars")
        if cmb_b is not None:
            for i in range(
                CABEZAL_MIN_BARRAS_POR_CAPA,
                CABEZAL_MAX_BARRAS_POR_CAPA + 1,
            ):
                cmb_b.Items.Add(unicode(i))
            n_def, _bt = _default_n_bars_and_type(analysis, self._doc)
            try:
                cmb_b.SelectedIndex = int(n_def) - CABEZAL_MIN_BARRAS_POR_CAPA
            except Exception:
                cmb_b.SelectedIndex = 0

        self._bar_types = _collect_bar_types(self._doc)
        cmb_t = self._win.FindName(u"CmbBarType")
        n_def, bt_id = _default_n_bars_and_type(analysis, self._doc)
        sel_i = 0
        if cmb_t is not None:
            for i, bt in enumerate(self._bar_types):
                cmb_t.Items.Add(_bar_type_label(bt))
                try:
                    if bt_id is not None and bt.Id == bt_id:
                        sel_i = i
                except Exception:
                    pass
            if self._bar_types:
                cmb_t.SelectedIndex = sel_i
            try:
                cmb_t.SelectionChanged += self._on_form_changed_redraw
            except Exception:
                pass

        sp = self._win.FindName(u"TxtSpacing")
        if sp is not None:
            sp.Text = unicode(_default_spacing_mm(analysis))

        self._populate_terminacion_cabeza_combo()
        self._populate_confinement_combo()
        self._set_status(
            u"{0} muro(s) · listo. Empalmes en elevación (Auto/Tramo/Cont.)."
            .format(len(self._walls))
        )

    def _n_capas_ui(self):
        cmb_n = self._win.FindName(u"CmbNCapas")
        if cmb_n is None:
            return 1
        try:
            return int(cmb_n.SelectedItem)
        except Exception:
            try:
                return int(cmb_n.SelectedIndex) + 1
            except Exception:
                return 1

    def _populate_terminacion_cabeza_combo(self):
        cmb = self._win.FindName(u"CmbTerminacionCabeza")
        if cmb is None:
            return
        cmb.Items.Clear()
        self._terminacion_cabeza_values = []
        sel = 0
        for i, (val, lab) in enumerate(_TERMINACION_CABEZA_OPTS):
            cmb.Items.Add(lab)
            self._terminacion_cabeza_values.append(val)
            if val == _TERMINACION_CABEZA_DEFAULT:
                sel = i
        if cmb.Items.Count > 0:
            cmb.SelectedIndex = sel

    def _read_terminacion_cabeza(self):
        cmb = self._win.FindName(u"CmbTerminacionCabeza")
        vals = getattr(self, u"_terminacion_cabeza_values", None) or []
        if cmb is not None and vals:
            try:
                idx = int(cmb.SelectedIndex)
                if 0 <= idx < len(vals):
                    return vals[idx]
            except Exception:
                pass
        return _TERMINACION_CABEZA_DEFAULT

    def _populate_confinement_combo(self):
        cmb = self._win.FindName(u"CmbConfinement")
        if cmb is None:
            return
        n = self._n_capas_ui()
        prev = None
        try:
            prev = cmb.SelectedItem
        except Exception:
            pass
        cmb.Items.Clear()
        opts = cabezal_confinement_options(n, encuentro=False)
        labels = []
        for val, lab in opts:
            labels.append(lab)
            cmb.Items.Add(lab)
        self._confinement_values = [v for v, _l in opts]
        sel = 0
        if prev is not None:
            try:
                for i, lab in enumerate(labels):
                    if lab == prev:
                        sel = i
                        break
            except Exception:
                pass
        if cmb.Items.Count > 0:
            cmb.SelectedIndex = min(sel, cmb.Items.Count - 1)
        hint = self._win.FindName(u"TxtConfHint")
        if hint is not None:
            if n < 2:
                hint.Text = u"Con 1 capa nueva solo aplica «Sin confinamiento»."
            else:
                hint.Text = (
                    u"Tipo 1 trabas · Tipo 2 estribo+trabas · "
                    u"Tipo 3 + cruz. Índices relativos a las capas nuevas."
                )

    def _on_n_capas_changed(self, sender, args):
        try:
            self._populate_confinement_combo()
        except Exception:
            pass
        try:
            self._redraw_elevation()
        except Exception:
            pass

    def _on_form_changed_redraw(self, sender, args):
        try:
            self._redraw_elevation()
        except Exception:
            pass

    def _read_confinement(self):
        ctype = CABEZAL_CONFINEMENT_NONE
        cmb = self._win.FindName(u"CmbConfinement")
        vals = getattr(self, u"_confinement_values", None) or [CABEZAL_CONFINEMENT_NONE]
        if cmb is not None:
            try:
                idx = int(cmb.SelectedIndex)
                if 0 <= idx < len(vals):
                    ctype = vals[idx]
            except Exception:
                pass
        if self._n_capas_ui() < 2:
            ctype = CABEZAL_CONFINEMENT_NONE
        diam = 10.0
        sp = 100.0
        td = self._win.FindName(u"TxtConfDiam")
        if td is not None:
            try:
                diam = float(_as_unicode(td.Text).replace(u",", u".").strip())
            except Exception:
                pass
        ts = self._win.FindName(u"TxtConfSpacing")
        if ts is not None:
            try:
                sp = float(_as_unicode(ts.Text).replace(u",", u".").strip())
            except Exception:
                pass
        return ctype, diam, sp

    def _read_form(self):
        n_capas = self._n_capas_ui()

        n_bars = CABEZAL_MIN_BARRAS_POR_CAPA
        cmb_b = self._win.FindName(u"CmbNBars")
        if cmb_b is not None:
            try:
                n_bars = int(cmb_b.SelectedItem)
            except Exception:
                n_bars = CABEZAL_MIN_BARRAS_POR_CAPA + int(cmb_b.SelectedIndex)

        bar_type_id = None
        cmb_t = self._win.FindName(u"CmbBarType")
        if cmb_t is not None and self._bar_types:
            try:
                idx = int(cmb_t.SelectedIndex)
                if 0 <= idx < len(self._bar_types):
                    bar_type_id = self._bar_types[idx].Id
            except Exception:
                pass

        spacing = _default_spacing_mm(self._analysis)
        sp = self._win.FindName(u"TxtSpacing")
        if sp is not None:
            try:
                spacing = float(_as_unicode(sp.Text).replace(u",", u".").strip())
            except Exception:
                pass

        ctype, c_diam, c_sp = self._read_confinement()
        modes = list(self._troceo_modes or [])
        while len(modes) < len(self._walls or []):
            modes.append(TROCEO_AUTO)
        return {
            u"n_capas": n_capas,
            u"n_bars": n_bars,
            u"bar_type_id": bar_type_id,
            u"spacing": spacing,
            u"flags": self._read_troceo_flags(),
            u"troceo_modes": modes,
            u"confinement_type": ctype,
            u"conf_diam_mm": c_diam,
            u"conf_spacing_mm": c_sp,
            u"terminacion_cabeza": self._read_terminacion_cabeza(),
            u"view_right_xy": self._view_right_xy(),
        }

    def _wire_events(self):
        btn_c = self._win.FindName(u"BtnCreate")
        if btn_c is not None:
            btn_c.Click += RoutedEventHandler(self._on_create)
        else:
            try:
                print(u"[CapasAdicionalesEnMuro] BtnCreate no encontrado en XAML.")
            except Exception:
                pass
        btn_r = self._win.FindName(u"BtnRepickWalls")
        if btn_r is not None:
            btn_r.Click += RoutedEventHandler(self._on_repick_walls)
        try:
            self._win.Closed += EventHandler(self._on_closed)
        except Exception:
            pass

    def _set_status(self, text):
        st = self._win.FindName(u"TxtStatus") if self._win is not None else None
        if st is not None:
            st.Text = _as_unicode(text)

    def _on_closed(self, sender, args):
        self._closed = True
        try:
            AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
        except Exception:
            pass
        # Cierre por «Crear capas»: mantener target/controller hasta ExternalEvent.
        if getattr(self, u"_closing_for_create", False):
            return
        _clear_create_target(self)
        _clear_active_controller(self)

    def _on_repick_walls(self, sender, args):
        if self._uidoc is None:
            _mostrar_aviso(self._uiapp, u"No hay documento activo.")
            return
        try:
            self._win.Hide()
        except Exception:
            pass
        try:
            walls = _pick_muros(self._uidoc)
        except Exception as ex:
            walls = None
            _mostrar_aviso(self._uiapp, u"Error al seleccionar muros.", content=_as_unicode(ex))
        try:
            self._win.Show()
        except Exception:
            pass
        if walls is None:
            self._set_status(u"Selección de muros cancelada.")
            return
        if not walls:
            self._set_status(u"Sin muros válidos.")
            return
        self._walls = ordenar_muros_por_base_asc(list(walls))
        self._troceo_modes = _default_troceo_modes(len(self._walls))
        self._selected_wall = 0
        self._selected_segment = 0
        try:
            self._ensure_stacked_layout(force=True)
        except Exception:
            self._stacked_layout = None
        info_muros = self._win.FindName(u"TxtInfoMuros")
        if info_muros is not None:
            info_muros.Text = (
                u"{0} muro(s) · extremo {1} · {2}\n"
                u"Empalmes en elevación: pie Auto→Tramo→Cont. o clic en fuste."
                .format(
                    len(self._walls),
                    _extremo_label(self._extremo),
                    self._extremo_note or u"",
                )
            )
        try:
            self._redraw_elevation()
        except Exception:
            pass
        self._set_status(u"{0} muro(s) seleccionados.".format(len(self._walls)))

    def _on_create(self, sender, args):
        try:
            if not self._analysis or not self._analysis.get(u"ok"):
                _mostrar_aviso(self._uiapp, u"Análisis GUID no válido.")
                return
            if not self._walls:
                _mostrar_aviso(self._uiapp, u"Selecciona al menos un muro.")
                return
            form = self._read_form()
            if form.get(u"n_capas", 0) < 1:
                _mostrar_aviso(self._uiapp, u"Indica al menos 1 capa nueva.")
                return
            if form.get(u"bar_type_id") is None:
                _mostrar_aviso(self._uiapp, u"Selecciona un tipo de barra.")
                return
            if float(form.get(u"spacing") or 0) <= 0:
                _mostrar_aviso(self._uiapp, u"Separación de capas inválida.")
                return

            # Cerrar UI y crear vía ExternalEvent (como Armado Muros v3).
            self._pending = form
            self._closing_for_create = True
            _set_create_target(self)
            evt = self._create_event or _ensure_create_event()
            try:
                if self._win is not None:
                    self._win.Close()
            except Exception:
                pass
            try:
                evt.Raise()
            except Exception as ex:
                self._closing_for_create = False
                _clear_create_target(self)
                _clear_active_controller(self)
                _mostrar_aviso(
                    self._uiapp,
                    u"No se pudo iniciar la creación.",
                    content=_as_unicode(ex),
                )
        except Exception as ex:
            self._closing_for_create = False
            _mostrar_aviso(
                self._uiapp,
                u"Error al preparar la creación.",
                content=_as_unicode(ex),
            )

    def _execute_create(self, uiapp):
        pending = getattr(self, u"_pending", None) or {}
        uidoc = getattr(uiapp, u"ActiveUIDocument", None)
        doc = uidoc.Document if uidoc is not None else self._doc
        try:
            res = crear_capas_adicionales_en_muros(
                doc,
                uidoc,
                self._analysis,
                self._walls,
                self._extremo,
                pending.get(u"n_capas", 1),
                pending.get(u"n_bars", 2),
                pending.get(u"bar_type_id"),
                pending.get(u"spacing", CABEZAL_LAYER_PITCH_MM),
                pending.get(u"flags") or [],
                confinement_type=pending.get(u"confinement_type"),
                conf_diam_mm=pending.get(u"conf_diam_mm", 10.0),
                conf_spacing_mm=pending.get(u"conf_spacing_mm", 100.0),
                troceo_modes=pending.get(u"troceo_modes"),
                view_right_xy=pending.get(u"view_right_xy"),
                terminacion_cabeza=pending.get(
                    u"terminacion_cabeza", _TERMINACION_CABEZA_DEFAULT,
                ),
            )
            msg = res.get(u"message") or u""
            try:
                self._set_status(msg)
            except Exception:
                pass
            if res.get(u"ok"):
                extra_parts = []
                fails = res.get(u"n_fail") or 0
                if fails:
                    extra_parts.append(u"Avisos/fallos: {0}".format(fails))
                for m in (res.get(u"messages") or [])[:4]:
                    if m:
                        extra_parts.append(_as_unicode(m))
                _mostrar_aviso(
                    uiapp,
                    msg,
                    content=u"\n".join(extra_parts) or u"Revisa el modelo y el historial de deshacer.",
                )
            else:
                detail = u"\n".join(
                    [_as_unicode(m) for m in (res.get(u"messages") or [])[:6] if m]
                )
                _mostrar_aviso(uiapp, msg or u"No se crearon capas.", content=detail)
        finally:
            self._closing_for_create = False
            self._pending = None
            _clear_create_target(self)
            _clear_active_controller(self)

    def show(self):
        try:
            AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, self._win)
        except Exception:
            pass
        _set_active_controller(self)
        try:
            hwnd = revit_main_hwnd(self._uiapp)
            bind_center_wpf_on_revit_monitor(self._win, hwnd)
        except Exception:
            pass
        try:
            self._win.WindowState = WindowState.Maximized
        except Exception:
            pass
        self._win.Show()
        try:
            self._redraw_elevation()
        except Exception:
            pass

# ── entrada ──────────────────────────────────────────────────────────────────


def run(revit):
    """Pick rebar → pick muros → UI."""
    existing_ctrl = _get_active_controller()
    existing = None
    if existing_ctrl is not None:
        try:
            existing = existing_ctrl._win
        except Exception:
            existing = None
    if existing is None:
        existing = _get_active_window()
    if existing is not None:
        try:
            if existing.WindowState == WindowState.Minimized:
                existing.WindowState = WindowState.Normal
        except Exception:
            pass
        try:
            existing.Activate()
            existing.Focus()
        except Exception:
            pass
        _mostrar_aviso(revit, u"La herramienta ya esta en ejecucion.")
        return

    _clear_active_controller()
    _clear_create_target()
    # Registrar ExternalEvent mientras aún hay contexto API del comando.
    try:
        _ensure_create_event()
    except Exception as ex:
        _mostrar_aviso(
            revit,
            u"No se pudo inicializar el evento de creación.",
            content=_as_unicode(ex),
        )
        return

    uidoc = getattr(revit, u"ActiveUIDocument", None)
    if uidoc is None:
        _mostrar_aviso(revit, u"No hay documento activo.")
        return
    doc = uidoc.Document

    el = _pick_rebar_element(
        uidoc,
        u"Selecciona una rebar longitudinal con Armadura_Conjunto_GUID",
    )
    if el is None:
        _mostrar_aviso(revit, u"Operación cancelada.")
        return

    analysis = analyze_conjunto(doc, el)
    # inward solo lo exige Capas GUID (offset espesor); aquí usamos pipeline
    # cabezal. Si hay capas+GUID, continuar aunque falle la dirección inward.
    if not analysis.get(u"ok"):
        if analysis.get(u"guid") and (analysis.get(u"layers") or []):
            analysis = dict(analysis)
            analysis[u"ok"] = True
            analysis[u"error"] = None
        else:
            _mostrar_aviso(
                revit,
                analysis.get(u"error") or u"No se pudo analizar el conjunto GUID.",
            )
            return

    # Acotar al host de la seed: última capa / extremo sin mezclar copias GUID.
    analysis = scope_analysis_to_seed_host(doc, analysis)
    if not analysis.get(u"layers"):
        _mostrar_aviso(
            revit,
            u"No se detectaron capas longitudinales para este GUID/host.",
        )
        return

    extremo, extremo_note = detect_extremo_from_seed(doc, analysis)
    scope_note = analysis.get(u"scope_note") or u""
    if scope_note:
        extremo_note = (extremo_note + u" · " + scope_note).strip(u" ·")
    last_idx = ultima_capa_index_desde_guid(analysis)
    extremo_note = (
        u"{0} · última capa GUID: {1}ºC. → nuevas desde {2}ºC."
        .format(
            extremo_note,
            last_idx + 1 if last_idx >= 0 else 0,
            base_offset_despues_ultima_capa(analysis) + 1,
        )
    )

    if not _guard_vista_armado_muros(uidoc, revit):
        return

    walls = _pick_muros(uidoc)
    if walls is None:
        _mostrar_aviso(revit, u"Selección de muros cancelada.")
        return
    if not walls:
        _mostrar_aviso(revit, u"No se seleccionaron muros válidos.")
        return
    walls = ordenar_muros_por_base_asc(list(walls))

    win = CapasAdicionalesEnMuroWindow(
        revit,
        analysis=analysis,
        walls=walls,
        extremo=extremo,
        extremo_note=extremo_note,
    )
    win.show()
