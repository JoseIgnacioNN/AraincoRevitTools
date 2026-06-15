# -*- coding: utf-8 -*-
"""
Cabezal de muro — barras verticales en extremos (inicio / fin de LocationCurve).

Por extremo: 1–6 capas empurradas en profundidad (longitudinal); 2–4 barras por capa
repartidas en el espesor con ``SetLayoutAsFixedNumber`` (misma lógica que multicapa).

Cantidad (``n_bars``) por capa en cada muro; diámetro (``bar_type_id``) por capa
en la config de cada muro/extremo. Defaults UI: ``cabezal_longitudinal_sic_*``
(S.I.C. borde de muro por espesor; 1 capa, editable). Tras crear, etiqueta cada barra longitudinal
con Structural Rebar Tag (familia ``EST_A_STRUCTURAL REBAR TAG_HORIZONTAL``) en la vista activa.
Los empalmes activos definen segmentos verticales (lista de muros abajo→arriba).
Cada barra: altura del muro (``WALL_USER_HEIGHT_PARAM`` − cover).
"""

from __future__ import print_function

import math
import os
import sys

import clr

clr.AddReference("RevitAPI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BooleanOperationsType,
    BooleanOperationsUtils,
    BuiltInCategory,
    BuiltInParameter,
    Color,
    Curve,
    CurveLoop,
    ElementId,
    FilteredElementCollector,
    GeometryCreationUtilities,
    JoinGeometryUtils,
    Line,
    Options,
    OverrideGraphicSettings,
    Plane,
    SketchPlane,
    Transaction,
    TransactionGroup,
    UnitUtils,
    UnitTypeId,
    ViewDetailLevel,
    Wall,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarHookType,
    RebarShape,
    RebarStyle,
)

try:
    from armado_muros_lineales import (
        location_curve_wall,
        obtener_espesor_muro_mm_approx,
        _wall_bottom_z_sort_key_ft,
        MODO_EJECUCION_RAPIDA,
        _tamano_lote_ejecucion,
        _refrescar_vista_fin_flujo,
    )
except Exception:
    location_curve_wall = None
    obtener_espesor_muro_mm_approx = None
    _wall_bottom_z_sort_key_ft = None
    MODO_EJECUCION_RAPIDA = True
    _refrescar_vista_fin_flujo = None

    def _tamano_lote_ejecucion(n_items, lote_legacy):
        try:
            n = int(n_items)
        except Exception:
            n = 0
        if n < 1:
            return 1
        if MODO_EJECUCION_RAPIDA:
            return n
        try:
            return max(1, int(lote_legacy))
        except Exception:
            return 1

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
    from bimtools_runtime import (
        is_cpython3_runtime,
        is_legacy_revit2024_armado,
        pyrevit_progress_bar_enabled,
        use_transaction_group_armado_muros,
    )
except Exception:
    def is_cpython3_runtime():
        try:
            import sys
            return int(sys.version_info[0]) >= 3
        except Exception:
            return False

    def is_legacy_revit2024_armado(doc):
        return False

    def pyrevit_progress_bar_enabled(doc):
        return False

    def use_transaction_group_armado_muros(doc, within_parent_transaction_group=False):
        return False

try:
    from bimtools_element_id import (
        element_id_to_int,
        normalize_muro_id_dict,
        normalize_muro_id_key,
        wall_id_int,
    )
except Exception:
    def element_id_to_int(eid):
        if eid is None:
            return None
        try:
            return int(eid.Value)
        except Exception:
            pass
        try:
            return int(eid.IntegerValue)
        except Exception:
            return None

    def wall_id_int(wall):
        if wall is None:
            return None
        return element_id_to_int(wall.Id)

    def normalize_muro_id_key(key):
        if key is None:
            return None
        try:
            return int(key)
        except Exception:
            return element_id_to_int(key)

    def normalize_muro_id_dict(mapping):
        if not mapping:
            return mapping
        out = {}
        for k, v in mapping.items():
            nk = normalize_muro_id_key(k)
            if nk is not None:
                out[int(nk)] = v
        return out

try:
    from bimtools_rebar_hook_lengths import (
        hook_length_mm_from_nominal_diameter_mm,
        pata_eje_curve_loop_mm_desde_tabla_mm,
        traslape_mm_from_nominal_diameter_mm,
    )
except Exception:
    hook_length_mm_from_nominal_diameter_mm = None
    pata_eje_curve_loop_mm_desde_tabla_mm = None
    traslape_mm_from_nominal_diameter_mm = None

try:
    import rebar_extender_l_ganchos_135_rps as l135
except Exception:
    l135 = None

try:
    from arearein_verticales_empotramiento_rps import (
        _create_from_curves_no_hooks,
        _hook_orient_for_create,
    )
except Exception:
    _create_from_curves_no_hooks = None
    _hook_orient_for_create = None

try:
    import armado_muros_cabezal_tags as _cab_tags
except Exception:
    _cab_tags = None

try:
    from armado_muros_cabezal_empalme_markers import colocar_marcadores_empalme_cabezal
except Exception:
    colocar_marcadores_empalme_cabezal = None

try:
    import armado_muros_cabezal_encuentro_l as _cab_enc_l
except Exception:
    _cab_enc_l = None

try:
    from armado_muros_rebar_params import (
        activar_armadura_arainco,
        activar_armadura_arainco_por_ids,
        stamp_cabezal_longitudinal_rebar,
    )
except Exception:
    activar_armadura_arainco = None
    activar_armadura_arainco_por_ids = None
    stamp_cabezal_longitudinal_rebar = None

CABEZAL_MAX_CAPAS = 6
CABEZAL_MIN_CAPAS = 1
CABEZAL_MIN_BARRAS_POR_CAPA = 2
CABEZAL_MAX_BARRAS_POR_CAPA = 4
CABEZAL_COVER_MM = 25.0
# Largo máximo de barra de acero comercial (mm) — aviso UI en controlador de tramo.
CABEZAL_MAX_BARRA_COMERCIAL_MM = 12000.0
# Separación entre ejes de capas sucesivas (desde cara cabezal hacia interior).
CABEZAL_LAYER_PITCH_MM = 150.0
CABEZAL_LAYER_PITCH_MIN_MM = 28.0
CABEZAL_EXTREMO_INICIO = u"inicio"
CABEZAL_EXTREMO_FIN = u"fin"
CABEZAL_EXTREMOS = (CABEZAL_EXTREMO_INICIO, CABEZAL_EXTREMO_FIN)

CABEZAL_CONFINEMENT_NONE = u"none"
CABEZAL_CONFINEMENT_PERIMETER_0_1 = u"perimeter_0_1"
CABEZAL_CONFINEMENT_TIE_LAYER_1 = u"tie_layer_1"
CABEZAL_TIE_LAYER_INDEX = 1
# Escenarios con Tipo 1 / Tipo 2 definidos (otros n_capas: pendiente).
CABEZAL_CONFINEMENT_SCENARIO_CAPAS = (2, 3, 4, 5, 6)


def cabezal_confinement_scenario_applies(n_capas):
    """True si ``n_capas`` tiene escenario de confinamiento Tipo 1 / Tipo 2."""
    try:
        return int(n_capas) in CABEZAL_CONFINEMENT_SCENARIO_CAPAS
    except Exception:
        return False


def cabezal_confinement_layout_spec(n_capas, conf_type):
    """
    Índices de capa para estribo perimetral y trabas según escenario.

    2 capas — Tipo 1: trabas [1]; Tipo 2: estribo [0, 1].
    3 capas — Tipo 1: trabas [1, 2]; Tipo 2: estribo [0, 2] + traba [1].
    4 capas — Tipo 1: trabas [1, 2, 3]; Tipo 2: estribo [0, 3] + trabas [1, 2].
    5 capas — Tipo 1: trabas [1, 2, 3, 4]; Tipo 2: estribo [0, 4] + trabas [1, 2, 3].
    6 capas — Tipo 1: trabas [1, 2, 3, 4, 5]; Tipo 2: estribo [0, 5] + trabas [1, 2, 3, 4].
    """
    try:
        n = int(n_capas)
    except Exception:
        n = 0
    if not cabezal_confinement_scenario_applies(n):
        return [], []
    if n == 6:
        if cabezal_confinement_is_tie_layer_1(conf_type):
            return [], [1, 2, 3, 4, 5]
        if cabezal_confinement_is_perimeter(conf_type):
            return [0, 5], [1, 2, 3, 4]
        return [], []
    if cabezal_confinement_is_tie_layer_1(conf_type):
        return [], list(range(1, n))
    if cabezal_confinement_is_perimeter(conf_type):
        inner_ties = list(range(1, n - 1)) if n > 2 else []
        return [0, n - 1], inner_ties
    return [], []


def cabezal_confinement_is_perimeter(conf_type):
    try:
        return unicode(conf_type or u"").strip() == CABEZAL_CONFINEMENT_PERIMETER_0_1
    except Exception:
        return conf_type == CABEZAL_CONFINEMENT_PERIMETER_0_1


def cabezal_confinement_is_tie_layer_1(conf_type):
    try:
        return unicode(conf_type or u"").strip() == CABEZAL_CONFINEMENT_TIE_LAYER_1
    except Exception:
        return conf_type == CABEZAL_CONFINEMENT_TIE_LAYER_1


def cabezal_perimeter_stirrup_layer_indices(n_capas):
    """Índices de capa para estribo Tipo 2: siempre ``[0, n-1]``."""
    try:
        n = int(n_capas)
    except Exception:
        n = CABEZAL_MIN_CAPAS
    if not cabezal_confinement_scenario_applies(n):
        return []
    return [0, n - 1]


def cabezal_effective_n_capas(ex_cfg):
    """``n_capas`` activas del extremo (capas usadas al crear barras)."""
    if not ex_cfg or not isinstance(ex_cfg, dict):
        return CABEZAL_MIN_CAPAS
    layers = cabezal_active_layers(ex_cfg)
    if layers:
        return len(layers)
    try:
        return max(
            CABEZAL_MIN_CAPAS,
            min(CABEZAL_MAX_CAPAS, int(ex_cfg.get(u"n_capas", CABEZAL_MIN_CAPAS))),
        )
    except Exception:
        return CABEZAL_MIN_CAPAS


def cabezal_copy_extremo_config(ex_cfg):
    """Copia superficial de extremo (ElementId y listas/dicts mutables)."""
    if not ex_cfg or not isinstance(ex_cfg, dict):
        return default_cabezal_extremo_config()
    out = {}
    for k, v in ex_cfg.items():
        if k == u"layers":
            out[k] = [
                dict(ly) if isinstance(ly, dict) else ly
                for ly in (v or [])
            ]
        elif k == u"confinement" and isinstance(v, dict):
            out[k] = dict(v)
        elif k == u"segment_bar_type_ids" and isinstance(v, dict):
            out[k] = {
                sk: list(sv) if isinstance(sv, list) else sv
                for sk, sv in v.items()
            }
        else:
            out[k] = v
    return out


def cabezal_copy_muro_config(cfg):
    """Copia config cabezal de un muro (inicio/fin) para el handler de creación."""
    if not cfg or not isinstance(cfg, dict):
        return {}
    return {
        ex: cabezal_copy_extremo_config((cfg or {}).get(ex))
        for ex in CABEZAL_EXTREMOS
    }


def cabezal_stamp_confinement_type(ex_cfg, conf_type, doc=None, fallback_bar_type_id=None):
    """
    Fija ``confinement.type`` (p. ej. desde combo UI) y sincroniza Ø / índices.
    """
    if not ex_cfg or not isinstance(ex_cfg, dict):
        return ex_cfg
    _normalize_cabezal_extremo_layers(ex_cfg)
    n_capas = cabezal_effective_n_capas(ex_cfg)
    prev = ex_cfg.get(u"confinement") or {}
    merged = dict(prev) if isinstance(prev, dict) else {}
    if conf_type:
        merged[u"type"] = conf_type
    ex_cfg[u"confinement"] = normalize_cabezal_confinement(merged, n_capas)
    if cabezal_confinement_is_perimeter(ex_cfg[u"confinement"].get(u"type")):
        conf = ex_cfg[u"confinement"]
        conf[u"layer_indices"] = cabezal_perimeter_stirrup_layer_indices(n_capas)
        ex_cfg[u"confinement"] = conf
    cabezal_sync_confinement_from_extremo(
        ex_cfg,
        doc,
        cabezal_sync_fallback_bar_type_id(ex_cfg, fallback_bar_type_id),
    )
    return ex_cfg


def cabezal_sync_fallback_bar_type_id(ex_cfg, bar_type_fallback=None):
    """``RebarBarType`` id de cabezal (capa 0) para sincronizar confinamiento."""
    fb = None
    if ex_cfg and isinstance(ex_cfg, dict):
        layers = cabezal_active_layers(ex_cfg)
        if layers:
            try:
                bid = layers[0].get(u"bar_type_id")
            except Exception:
                bid = None
            if bid is not None and bid != ElementId.InvalidElementId:
                fb = bid
        if fb is None or fb == ElementId.InvalidElementId:
            fb = ex_cfg.get(u"bar_type_id")
    if fb is None or fb == ElementId.InvalidElementId:
        fb = bar_type_fallback
    return fb


def cabezal_resolve_bar_type_fallback(doc, cabezal_por_muro_id, walls=None):
    """
    Primer ``RebarBarType`` de cabezal (capas longitudinales), no malla horizontal.

    Mismo criterio que ``_CrearCabezalEjecutarHandler`` en la UI.
    """
    fallback_bt = None
    try:
        for _wid, cfg in (cabezal_por_muro_id or {}).items():
            if not cfg:
                continue
            for ex in CABEZAL_EXTREMOS:
                ex_cfg = (cfg or {}).get(ex) or {}
                _normalize_cabezal_extremo_layers(ex_cfg)
                if walls:
                    try:
                        segs = build_cabezal_segments(
                            len(walls),
                            _empalme_stack_indices(
                                walls, cabezal_por_muro_id, ex,
                            ),
                        )
                        _migrate_tramo_to_segment_bar_type_ids(
                            ex_cfg, segs, ex_cfg.get(u"bar_type_id"),
                        )
                    except Exception:
                        pass
                for ly in cabezal_active_layers(ex_cfg):
                    bid = ly.get(u"bar_type_id")
                    if bid and bid != ElementId.InvalidElementId:
                        return bid
                bid = ex_cfg.get(u"bar_type_id")
                if bid and bid != ElementId.InvalidElementId:
                    fallback_bt = bid
            if fallback_bt is not None:
                break
    except Exception:
        fallback_bt = None
    return fallback_bt


def cabezal_confinement_has_lote_z(conf_type):
    return (
        cabezal_confinement_is_perimeter(conf_type)
        or cabezal_confinement_is_tie_layer_1(conf_type)
    )


CABEZAL_STIRRUP_DIAM_MM = 10.0
CABEZAL_STIRRUP_SPACING_MM = 200.0
# Distancia del rebar set de confinamiento (estribos/trabas en Z); fija, no sigue @ malla horizontal.
CABEZAL_CONFINEMENT_REBAR_SET_SPACING_MM = 100.0
CABEZAL_CONFINEMENT_STIRRUP_SHAPE_NAME = u"10"
CABEZAL_CONFINEMENT_STIRRUP_PAD_MM = 10.0
# Estiramiento del array de estribos hacia −Z si hay fundación unida (Join Geometry).
CABEZAL_STIRRUP_FOUNDATION_DROP_MM = 300.0
# Diámetro por defecto (UI) de barras longitudinales del cabezal.
CABEZAL_DEFAULT_BAR_DIAM_MM = 12.0
# Ø comerciales para estribos/trabas de confinamiento (regla: >= d_long_capa0 / 3).
CABEZAL_CONFINEMENT_DIAMETERS_MM = (6, 8, 10, 12, 16, 18, 22, 25, 28, 32, 36)
# Lotes de transacción (1 = aparición progresiva de abajo hacia arriba).
CABEZAL_BARRAS_POR_LOTE_ANIMACION = 1
CABEZAL_CONFINAMIENTO_POR_LOTE_ANIMACION = 1
CABEZAL_TAGS_POR_LOTE_ANIMACION = 1
_CAB_PBAR_BASE_BARS = u"Arainco: Cabezal muros (barras)"
_CAB_PBAR_BASE_CONF = u"Arainco: Cabezal muros (confinamiento)"
_CAB_PBAR_BASE_TAGS = u"Arainco: Cabezal muros (etiquetas)"


def cabezal_confinement_diam_mm_for_long_layer0(d0_mm):
    """
    Ø de confinamiento (mm): menor valor del catálogo >= ``d0_mm / 3``
    (cumple o supera el mínimo normativo).
    """
    try:
        d0 = float(d0_mm)
    except Exception:
        d0 = float(CABEZAL_DEFAULT_BAR_DIAM_MM)
    if d0 < 1e-6:
        d0 = float(CABEZAL_DEFAULT_BAR_DIAM_MM)
    d_min = d0 / 3.0
    ok = [d for d in CABEZAL_CONFINEMENT_DIAMETERS_MM if d >= d_min - 1e-6]
    if not ok:
        return float(CABEZAL_CONFINEMENT_DIAMETERS_MM[-1])
    best = ok[0]
    best_key = (best - d_min, best)
    for d in ok[1:]:
        k = (d - d_min, d)
        if k < best_key:
            best, best_key = d, k
    return float(best)


def cabezal_longitudinal_layer0_diam_mm(doc, ex_cfg, fallback_bar_type_id=None):
    """Mayor Ø nominal (mm) de la capa longitudinal índice 0 del extremo."""
    if not ex_cfg or not isinstance(ex_cfg, dict):
        return float(CABEZAL_DEFAULT_BAR_DIAM_MM)
    n_tramos = 1
    try:
        if ex_cfg.get(u"troceo_por_muro"):
            seg = ex_cfg.get(u"segment_bar_type_ids") or {}
            if seg:
                n_tramos = max(
                    n_tramos,
                    max(len(v or []) for v in seg.values()),
                )
    except Exception:
        pass
    layers = _normalize_cabezal_extremo_layers(ex_cfg, n_tramos)
    ly0 = layers[0] if layers else {}
    diams = []
    bt0 = _resolver_bar_type_for_layer(
        doc, ex_cfg, ly0, fallback_bar_type_id, layer_index=0,
    )
    if bt0 is not None:
        diams.append(_bar_diameter_mm(bt0))
    for ti in range(max(1, int(n_tramos))):
        bt = _resolver_bar_type_for_layer_tramo(
            doc, ex_cfg, ly0, ti, fallback_bar_type_id,
        )
        if bt is not None:
            diams.append(_bar_diameter_mm(bt))
    if diams:
        return float(max(diams))
    if doc is not None and fallback_bar_type_id not in (None, ElementId.InvalidElementId):
        bt_fb = _element_to_bar_type(doc, fallback_bar_type_id)
        if bt_fb is not None:
            return float(_bar_diameter_mm(bt_fb))
    return float(CABEZAL_DEFAULT_BAR_DIAM_MM)


def _bar_type_for_catalog_diameter_mm(doc, diam_mm, fallback=None):
    """``RebarBarType`` cuyo nominal coincide con un Ø del catálogo de confinamiento."""
    if doc is None:
        return fallback
    try:
        target = float(diam_mm)
    except Exception:
        return fallback
    tol = 0.75
    best = None
    best_diff = None
    try:
        for bt in FilteredElementCollector(doc).OfClass(RebarBarType):
            d = _bar_diameter_mm(bt)
            if abs(d - target) <= tol:
                diff = abs(d - target)
                if best is None or diff < best_diff:
                    best = bt
                    best_diff = diff
                    if diff < 0.01:
                        break
    except Exception:
        pass
    if best is not None:
        return best
    return _bar_type_for_diameter_mm(doc, target, fallback)


def cabezal_sync_confinement_from_extremo(ex_cfg, doc=None, fallback_bar_type_id=None):
    """
    Ø confinamiento: catálogo >= d_long capa 0 / 3; @ rebar set fijo 100 mm.
    Actualiza ``confinement`` y ``conf_bar_type_id``.
    """
    if not ex_cfg or not isinstance(ex_cfg, dict):
        return ex_cfg
    try:
        n_capas = int(ex_cfg.get(u"n_capas", CABEZAL_MIN_CAPAS))
    except Exception:
        n_capas = CABEZAL_MIN_CAPAS
    d0 = cabezal_longitudinal_layer0_diam_mm(
        doc, ex_cfg, fallback_bar_type_id,
    )
    diam_conf = cabezal_confinement_diam_mm_for_long_layer0(d0)
    conf = normalize_cabezal_confinement(
        ex_cfg.get(u"confinement"), n_capas,
    )
    conf[u"stirrup_diam_mm"] = float(diam_conf)
    conf[u"stirrup_spacing_mm"] = float(CABEZAL_CONFINEMENT_REBAR_SET_SPACING_MM)
    ex_cfg[u"confinement"] = conf
    bt_conf = _bar_type_for_catalog_diameter_mm(
        doc, diam_conf, _element_to_bar_type(doc, fallback_bar_type_id),
    )
    if bt_conf is not None:
        try:
            ex_cfg[u"conf_bar_type_id"] = bt_conf.Id
        except Exception:
            pass
    return ex_cfg


def _stamp_armadura_arainco(rebar, layer_index=None):
    """
    Marca Rebar con ``Armadura_Arainco``.

    Si ``layer_index`` no es None, también rellena ``Armadura_Ubicacion`` (capa+1).
    """
    if rebar is None:
        return rebar
    if layer_index is not None and stamp_cabezal_longitudinal_rebar is not None:
        return stamp_cabezal_longitudinal_rebar(rebar, layer_index)
    if activar_armadura_arainco is not None:
        activar_armadura_arainco(rebar)
    return rebar


def default_cabezal_layer_config(n_bars=2, bar_type_id=None):
    return {
        u"n_bars": int(n_bars),
        u"bar_type_id": bar_type_id,
        u"tramo_bar_type_ids": [],
        u"tramo_n_bars": [],
    }


def _normalize_tramo_n_bars(ly, n_tramos, fallback_n_bars=None):
    """Lista ``tramo_n_bars`` de longitud ``n_tramos`` (T1=índice 0, base)."""
    if not ly or not isinstance(ly, dict):
        ly = {}
    n_tramos = max(1, int(n_tramos or 1))
    raw = ly.get(u"tramo_n_bars")
    if raw is None:
        raw = []
    try:
        fb = int(fallback_n_bars if fallback_n_bars is not None else ly.get(u"n_bars", CABEZAL_MIN_BARRAS_POR_CAPA))
    except Exception:
        fb = CABEZAL_MIN_BARRAS_POR_CAPA
    fb = max(CABEZAL_MIN_BARRAS_POR_CAPA, min(CABEZAL_MAX_BARRAS_POR_CAPA, fb))
    out = []
    for i in range(n_tramos):
        try:
            nb = int(raw[i]) if i < len(raw) else fb
        except Exception:
            nb = fb
        out.append(max(
            CABEZAL_MIN_BARRAS_POR_CAPA,
            min(CABEZAL_MAX_BARRAS_POR_CAPA, nb),
        ))
    return out


def count_cabezal_tramos_verticales(n_troceo_walls):
    """Número de tramos verticales UI: 1 + muros con empalme/troceo activo."""
    try:
        n = int(n_troceo_walls or 0)
    except Exception:
        n = 0
    return max(1, 1 + n)


def _normalize_tramo_bar_type_ids(ly, n_tramos, fallback_bar_type_id=None):
    """Lista ``tramo_bar_type_ids`` de longitud ``n_tramos`` (T1=índice 0, base)."""
    if not ly or not isinstance(ly, dict):
        ly = {}
    n_tramos = max(1, int(n_tramos or 1))
    raw = ly.get(u"tramo_bar_type_ids")
    try:
        from bimtools_clr_collections import as_python_list, list_get_or_last

        raw = as_python_list(raw)
    except Exception:
        if raw is None:
            raw = []
    out = []
    fb = ly.get(u"bar_type_id")
    if fb is None or fb == ElementId.InvalidElementId:
        fb = fallback_bar_type_id
    for i in range(n_tramos):
        try:
            from bimtools_clr_collections import list_get_or_last

            bid = list_get_or_last(raw, i, default=None)
        except Exception:
            bid = raw[i] if i < len(raw) else None
        if bid is None or bid == ElementId.InvalidElementId:
            bid = fb
        out.append(bid)
    return out


def default_cabezal_confinement_config(n_capas=None):
    try:
        n = int(n_capas if n_capas is not None else 2)
    except Exception:
        n = 2
    if cabezal_confinement_scenario_applies(n):
        stirrup_idx, tie_idx = cabezal_confinement_layout_spec(
            n, CABEZAL_CONFINEMENT_TIE_LAYER_1,
        )
        return {
            u"type": CABEZAL_CONFINEMENT_TIE_LAYER_1,
            u"layer_indices": list(stirrup_idx),
            u"tie_layer_indices": list(tie_idx),
            u"tie_layer_index": tie_idx[0] if tie_idx else CABEZAL_TIE_LAYER_INDEX,
            u"stirrup_diam_mm": CABEZAL_STIRRUP_DIAM_MM,
            u"stirrup_spacing_mm": CABEZAL_CONFINEMENT_REBAR_SET_SPACING_MM,
        }
    return {
        u"type": CABEZAL_CONFINEMENT_NONE,
        u"layer_indices": [],
        u"tie_layer_indices": [],
        u"stirrup_diam_mm": CABEZAL_STIRRUP_DIAM_MM,
        u"stirrup_spacing_mm": CABEZAL_CONFINEMENT_REBAR_SET_SPACING_MM,
    }


def default_cabezal_extremo_config():
    return {
        u"layers": [
            default_cabezal_layer_config(2, None),
            default_cabezal_layer_config(2, None),
        ],
        u"n_capas": 2,
        u"bar_type_id": None,
        u"conf_bar_type_id": None,
        u"layer_spacing_mm": CABEZAL_LAYER_PITCH_MM,
        u"troceo_por_muro": False,
        u"troceo_por_muro_override": None,
        u"troceo_auto_geom": False,
        u"armado_activo": True,
        u"segment_bar_type_ids": {},
        u"confinement": default_cabezal_confinement_config(),
    }


def cabezal_extremo_armado_activo(ex_cfg):
    """True si debe generarse armado completo (barras + confinamiento) en el extremo."""
    if not ex_cfg or not isinstance(ex_cfg, dict):
        return True
    try:
        return bool(ex_cfg.get(u"armado_activo", True))
    except Exception:
        return True


def cabezal_confinement_options(n_capas):
    """Opciones UI: (valor, etiqueta). Tipo 1/2 en escenarios de 2 a 6 capas."""
    opts = [(CABEZAL_CONFINEMENT_NONE, u"Sin confinamiento")]
    try:
        n = int(n_capas or 0)
    except Exception:
        n = 0
    if cabezal_confinement_scenario_applies(n):
        opts.append((
            CABEZAL_CONFINEMENT_PERIMETER_0_1,
            u"Tipo 2",
        ))
        opts.append((
            CABEZAL_CONFINEMENT_TIE_LAYER_1,
            u"Tipo 1",
        ))
    return opts


def normalize_cabezal_confinement(raw, n_capas=None):
    """Normaliza bloque ``confinement`` del extremo."""
    try:
        n = int(n_capas if n_capas is not None else CABEZAL_MAX_CAPAS)
    except Exception:
        n = CABEZAL_MAX_CAPAS
    base = default_cabezal_confinement_config(n)
    if not raw or not isinstance(raw, dict):
        raw = {}
    try:
        base[u"stirrup_diam_mm"] = float(
            raw.get(u"stirrup_diam_mm", CABEZAL_STIRRUP_DIAM_MM),
        )
    except Exception:
        base[u"stirrup_diam_mm"] = float(CABEZAL_STIRRUP_DIAM_MM)
    try:
        base[u"stirrup_spacing_mm"] = float(CABEZAL_CONFINEMENT_REBAR_SET_SPACING_MM)
    except Exception:
        pass
    ctype = raw.get(u"type")
    if ctype is None:
        ctype = (
            CABEZAL_CONFINEMENT_TIE_LAYER_1
            if cabezal_confinement_scenario_applies(n)
            else CABEZAL_CONFINEMENT_NONE
        )
    elif not ctype:
        ctype = CABEZAL_CONFINEMENT_NONE
    stirrup_idx, tie_idx = cabezal_confinement_layout_spec(n, ctype)
    if (
        ctype == CABEZAL_CONFINEMENT_PERIMETER_0_1
        and cabezal_confinement_scenario_applies(n)
    ):
        base[u"type"] = CABEZAL_CONFINEMENT_PERIMETER_0_1
        base[u"layer_indices"] = list(stirrup_idx)
        base[u"tie_layer_indices"] = list(tie_idx)
        if tie_idx:
            base[u"tie_layer_index"] = tie_idx[0]
        else:
            base.pop(u"tie_layer_index", None)
    elif (
        ctype in (CABEZAL_CONFINEMENT_TIE_LAYER_1, u"perimeter_0_1_tie_1")
        and cabezal_confinement_scenario_applies(n)
    ):
        base[u"type"] = CABEZAL_CONFINEMENT_TIE_LAYER_1
        base[u"layer_indices"] = list(stirrup_idx)
        base[u"tie_layer_indices"] = list(tie_idx)
        if tie_idx:
            base[u"tie_layer_index"] = tie_idx[0]
        else:
            base.pop(u"tie_layer_index", None)
    else:
        base[u"type"] = CABEZAL_CONFINEMENT_NONE
        base[u"layer_indices"] = []
        base[u"tie_layer_indices"] = []
    return base


def _parse_malla_spacing_mm(esp_txt):
    try:
        return float(unicode(esp_txt).strip().replace(u",", u"."))
    except Exception:
        try:
            return float(str(esp_txt).strip().replace(u",", u"."))
        except Exception:
            return float(CABEZAL_STIRRUP_SPACING_MM)


def cabezal_malla_horizontal_spec(params_dict, layer_active_dict, doc=None):
    """
    Barras horizontales de la malla (capa Major, mismo criterio que etiqueta H.).
    Retorna (diam_mm, spacing_mm, bar_type_id).
    """
    if not params_dict:
        params_dict = {}
    if not layer_active_dict:
        layer_active_dict = {}

    def _activa(key):
        return bool(layer_active_dict.get(key, True))

    def _par(key):
        if not _activa(key):
            return ElementId.InvalidElementId, None
        return params_dict.get(key, (ElementId.InvalidElementId, u"150"))

    pref = u"exterior"
    if not (_activa(u"exterior_major") or _activa(u"exterior_minor")):
        if _activa(u"interior_major") or _activa(u"interior_minor"):
            pref = u"interior"

    h_bar, h_esp = _par(u"{}_major".format(pref))
    if h_esp is None:
        _, h_esp_alt = _par(u"{}_minor".format(pref))
        if h_esp_alt is not None:
            h_esp = h_esp_alt
    if h_bar is None or h_bar == ElementId.InvalidElementId:
        h_bar, _ = _par(u"{}_minor".format(pref))

    spacing_mm = _parse_malla_spacing_mm(h_esp if h_esp is not None else u"150")
    diam_mm = float(CABEZAL_STIRRUP_DIAM_MM)
    if doc is not None and h_bar is not None and h_bar != ElementId.InvalidElementId:
        bt = _element_to_bar_type(doc, h_bar)
        if bt is not None:
            diam_mm = float(_bar_diameter_mm(bt))
    return diam_mm, spacing_mm, h_bar


def cabezal_sync_confinement_from_malla(ex_cfg, params_dict, layer_active_dict, doc=None):
    """
    Retrocompatibilidad de firma: Ø conf. por regla capa 0 (>= d0/3); @ 100 mm fijo.
    ``params_dict`` / ``layer_active_dict`` no alteran el Ø de confinamiento.
    """
    fb = cabezal_sync_fallback_bar_type_id(ex_cfg)
    return cabezal_sync_confinement_from_extremo(ex_cfg, doc, fb)


def cabezal_sync_muro_confinement_from_malla(cfg, params_dict, layer_active_dict, doc=None):
    """Sincroniza confinamiento (inicio/fin) según capa 0 longitudinal de cada extremo."""
    if not cfg or not isinstance(cfg, dict):
        return cfg
    for ex in CABEZAL_EXTREMOS:
        ex_cfg = cfg.setdefault(ex, default_cabezal_extremo_config())
        fb = cabezal_sync_fallback_bar_type_id(ex_cfg, cfg.get(u"bar_type_id"))
        cabezal_sync_confinement_from_extremo(ex_cfg, doc, fb)
    return cfg


def cabezal_confinement_tie_layer_indices(conf, n_capas=None):
    """Lista de índices de capa con traba (según escenario n_capas + tipo)."""
    if not conf or not isinstance(conf, dict):
        return []
    try:
        n = int(
            n_capas if n_capas is not None else conf.get(u"n_capas") or 2,
        )
    except Exception:
        n = 2
    ctype = conf.get(u"type")
    if cabezal_confinement_scenario_applies(n):
        _, tie_idx = cabezal_confinement_layout_spec(n, ctype)
        return list(tie_idx)
    _, tie_idx = cabezal_confinement_layout_spec(n, ctype)
    if tie_idx:
        return list(tie_idx)
    raw = conf.get(u"tie_layer_indices")
    if raw is not None:
        try:
            out = [int(i) for i in raw]
            if out:
                return out
        except Exception:
            pass
    try:
        li = conf.get(u"tie_layer_index")
        if li is not None:
            return [int(li)]
    except Exception:
        pass
    return []


def cabezal_build_confinement_jobs(
    wall, wid, extremo, ex_cfg, stack_idx, segment_ctx,
    line_jobs=None,
):
    """Un job por estribo y/o por cada traba (Tipo 2 @ 3 capas = estribo + traba)."""
    _normalize_cabezal_extremo_layers(ex_cfg)
    conf_type, stirrup_idx = cabezal_active_confinement(ex_cfg)
    if conf_type == CABEZAL_CONFINEMENT_NONE:
        return []
    n_capas = cabezal_effective_n_capas(ex_cfg)
    conf = normalize_cabezal_confinement(
        ex_cfg.get(u"confinement"), n_capas,
    )
    ex_cfg[u"confinement"] = conf
    tie_idx = cabezal_confinement_tie_layer_indices(conf, n_capas)
    base = {
        u"wall": wall,
        u"wid": wid,
        u"extremo": extremo,
        u"ex_cfg": ex_cfg,
        u"stack_idx": stack_idx,
        u"segment_ctx": segment_ctx,
        u"conf_type": conf_type,
        u"line_jobs": list(line_jobs or []),
    }
    jobs = []
    if stirrup_idx and cabezal_confinement_is_perimeter(conf_type):
        jobs.append(dict(base, job_kind=u"stirrup", layer_indices=list(stirrup_idx)))
    for li in tie_idx:
        jobs.append(dict(base, job_kind=u"tie", tie_layer_index=int(li)))
    return jobs


def cabezal_active_confinement(ex_cfg):
    """Retorna ``(type, layer_indices)`` activos para creación / preview."""
    if not ex_cfg or not isinstance(ex_cfg, dict):
        return CABEZAL_CONFINEMENT_NONE, []
    try:
        n_capas = int(ex_cfg.get(u"n_capas", CABEZAL_MIN_CAPAS))
    except Exception:
        n_capas = CABEZAL_MIN_CAPAS
    conf = normalize_cabezal_confinement(
        ex_cfg.get(u"confinement"), n_capas,
    )
    conf_type = conf.get(u"type")
    if cabezal_confinement_is_perimeter(conf_type):
        active_n = cabezal_effective_n_capas(ex_cfg)
        return conf_type, cabezal_perimeter_stirrup_layer_indices(active_n)
    return conf_type, list(conf.get(u"layer_indices") or [])


def build_cabezal_segments(n_walls, empalme_indices):
    """
    Segmentos post-empalme (muros abajo→arriba, índice 0 = menor Z).

    Cada índice con empalme activo **inicia** un tramo: el owner del ø es ese
    muro. El tramo incluye ese muro y los superiores hasta el siguiente empalme
    (sin incluir el muro del empalme superior).

    Ejemplo ``n=5``, empalmes ``[1, 3]`` →
    ``S0=[0]``, ``S1=[1,2] owner 1``, ``S2=[3,4] owner 3``.
    """
    try:
        n = max(0, int(n_walls or 0))
    except Exception:
        n = 0
    if n <= 0:
        return []
    E = sorted(int(i) for i in (empalme_indices or []))
    E = [i for i in E if 0 <= i < n]
    if not E:
        return [{u"id": 0, u"wall_indices": list(range(n)), u"owner_index": 0}]
    segs = []
    sid = 0
    if E[0] > 0:
        segs.append({
            u"id": sid,
            u"wall_indices": list(range(0, E[0])),
            u"owner_index": 0,
        })
        sid += 1
    for j in range(len(E)):
        if j + 1 < len(E):
            wall_indices = list(range(E[j], E[j + 1]))
        else:
            wall_indices = list(range(E[j], n))
        if not wall_indices:
            continue
        segs.append({
            u"id": sid,
            u"wall_indices": wall_indices,
            u"owner_index": E[j],
        })
        sid += 1
    return segs


def _empalme_stack_indices(walls, cabezal_por_muro_id, extremo):
    """Índices (stack abajo→arriba) de muros con empalme activo en ``extremo``."""
    out = []
    if not walls:
        return out
    for i, wall in enumerate(walls):
        if wall is None:
            continue
        try:
            wid = wall_id_int(wall)
        except Exception:
            continue
        cfg = (cabezal_por_muro_id or {}).get(wid) or {}
        ex_cfg = cfg.get(extremo) or {}
        if _troceo_por_muro_from_extremo_cfg(ex_cfg):
            out.append(i)
    return out


def _wall_stack_index(walls, wall):
    if wall is None or not walls:
        return 0
    try:
        target = wall_id_int(wall)
    except Exception:
        return 0
    for i, w in enumerate(walls):
        if w is None:
            continue
        try:
            if wall_id_int(w) == target:
                return i
        except Exception:
            pass
    return 0


def _stack_index_for_z(walls, z_ft, curve_tol=1e-6):
    """Índice de muro cuyo rango Z contiene ``z_ft`` (stack abajo→arriba)."""
    if not walls:
        return 0
    z = float(z_ft)
    tol = max(float(curve_tol or 1e-6), 1e-9)
    for i, wall in enumerate(walls):
        if wall is None:
            continue
        z0, z1 = _wall_z_bounds_ft(wall)
        if z0 - tol <= z <= z1 + tol:
            return i
    best_i = 0
    best_d = None
    for i, wall in enumerate(walls):
        if wall is None:
            continue
        z0, z1 = _wall_z_bounds_ft(wall)
        zm = 0.5 * (z0 + z1)
        d = abs(z - zm)
        if best_d is None or d < best_d:
            best_d = d
            best_i = i
    return best_i


def _segment_bar_type_ids_key(raw, seg_id):
    if not raw or not isinstance(raw, dict):
        return None
    if seg_id in raw:
        return seg_id
    try:
        sk = str(int(seg_id))
        if sk in raw:
            return sk
    except Exception:
        pass
    return None


def _normalize_segment_layer_bar_ids(
    ex_cfg, seg_id, n_layers, fallback_bar_type_id=None,
):
    """Lista de ``ElementId`` por capa para un segmento (longitud ``n_layers``)."""
    try:
        from bimtools_clr_collections import as_python_list, list_get_or_last
    except Exception:
        as_python_list = None
        list_get_or_last = None

    n_layers = max(1, int(n_layers or 1))
    raw_map = (ex_cfg or {}).get(u"segment_bar_type_ids") or {}
    key = _segment_bar_type_ids_key(raw_map, seg_id)
    raw = raw_map.get(key) if key is not None else None
    if as_python_list is not None:
        raw = as_python_list(raw)
    elif raw is None:
        raw = []
    layers = (ex_cfg or {}).get(u"layers") or []
    if as_python_list is not None:
        layers = as_python_list(layers)
    fb_ext = (ex_cfg or {}).get(u"bar_type_id")
    if fb_ext is None or fb_ext == ElementId.InvalidElementId:
        fb_ext = fallback_bar_type_id
    out = []
    for li in range(n_layers):
        if list_get_or_last is not None:
            bid = list_get_or_last(raw, li, default=None)
        else:
            bid = raw[li] if li < len(raw) else None
        if bid is None or bid == ElementId.InvalidElementId:
            ly = layers[li] if li < len(layers) else {}
            bid = (ly or {}).get(u"bar_type_id")
        if bid is None or bid == ElementId.InvalidElementId:
            bid = fb_ext
        out.append(bid)
    return out


def _migrate_tramo_to_segment_bar_type_ids(ex_cfg, segments, fallback_bar_type_id=None):
    """Copia ``tramo_bar_type_ids`` legacy a ``segment_bar_type_ids`` si hace falta."""
    if not ex_cfg or not isinstance(ex_cfg, dict) or not segments:
        return
    raw_map = ex_cfg.setdefault(u"segment_bar_type_ids", {})
    if not isinstance(raw_map, dict):
        raw_map = {}
        ex_cfg[u"segment_bar_type_ids"] = raw_map
    layers = ex_cfg.get(u"layers") or []
    n_layers = max(1, len(layers))
    for seg in segments:
        seg_id = seg[u"id"]
        if _segment_bar_type_ids_key(raw_map, seg_id) is not None:
            continue
        per_layer = []
        for li, ly in enumerate(layers):
            ly = ly or {}
            try:
                from bimtools_clr_collections import list_get_or_last

                tramo_ids = ly.get(u"tramo_bar_type_ids") or []
                bid = list_get_or_last(tramo_ids, int(seg_id), default=None)
            except Exception:
                tramo_ids = ly.get(u"tramo_bar_type_ids") or []
                bid = None
                ti = int(seg_id)
                if tramo_ids and len(tramo_ids) > 0:
                    if ti < len(tramo_ids):
                        bid = tramo_ids[ti]
                    else:
                        bid = tramo_ids[-1]
            if bid is None or bid == ElementId.InvalidElementId:
                bid = ly.get(u"bar_type_id")
            if bid is None or bid == ElementId.InvalidElementId:
                bid = ex_cfg.get(u"bar_type_id")
            if bid is None or bid == ElementId.InvalidElementId:
                bid = fallback_bar_type_id
            per_layer.append(bid)
        while len(per_layer) < n_layers:
            per_layer.append(fallback_bar_type_id)
        raw_map[seg_id] = per_layer


def resolve_segment_context_for_extremo(
    walls, cabezal_por_muro_id, extremo, n_layers, fallback_bar_type_id=None,
):
    """
    Contexto de segmentos para un extremo del stack completo.

    Retorna dict con ``segments``, ``wall_to_seg``, ``seg_bar_type_ids``.
    """
    n_layers = max(1, int(n_layers or 1))
    empalme_idxs = _empalme_stack_indices(walls, cabezal_por_muro_id, extremo)
    segments = build_cabezal_segments(len(walls or []), empalme_idxs)
    wall_to_seg = {}
    seg_bar_type_ids = {}
    for seg in segments:
        for wi in seg[u"wall_indices"]:
            wall_to_seg[wi] = seg[u"id"]
    for seg in segments:
        seg_id = seg[u"id"]
        owner_idx = int(seg[u"owner_index"])
        owner_wid = None
        if walls and 0 <= owner_idx < len(walls):
            try:
                owner_wid = wall_id_int(walls[owner_idx])
            except Exception:
                owner_wid = None
        owner_ex = {}
        if owner_wid is not None:
            owner_ex = (
                (cabezal_por_muro_id or {}).get(owner_wid) or {}
            ).get(extremo) or {}
        if not owner_ex:
            owner_ex = {}
        _migrate_tramo_to_segment_bar_type_ids(
            owner_ex, [seg], fallback_bar_type_id,
        )
        seg_bar_type_ids[seg_id] = _normalize_segment_layer_bar_ids(
            owner_ex, seg_id, n_layers, fallback_bar_type_id,
        )
    return {
        u"segments": segments,
        u"wall_to_seg": wall_to_seg,
        u"seg_bar_type_ids": seg_bar_type_ids,
        u"empalme_indices": empalme_idxs,
    }


def _bar_type_from_segment_context(doc, segment_ctx, stack_index, layer_index, fallback_id=None):
    if not segment_ctx:
        return _element_to_bar_type(doc, fallback_id)
    seg_id = (segment_ctx.get(u"wall_to_seg") or {}).get(int(stack_index), 0)
    bid = fallback_id
    try:
        seg_map = segment_ctx.get(u"seg_bar_type_ids") or {}
        raw = seg_map.get(seg_id)
        if raw is None:
            try:
                raw = seg_map.get(int(seg_id))
            except Exception:
                pass
        if raw is None:
            try:
                raw = seg_map.get(str(int(seg_id)))
            except Exception:
                pass
        from bimtools_clr_collections import list_get_or_last

        bid = list_get_or_last(
            _cabezal_as_python_list(raw),
            int(layer_index),
            default=fallback_id,
        )
    except Exception:
        try:
            bids = _cabezal_as_python_list(
                (segment_ctx.get(u"seg_bar_type_ids") or {}).get(seg_id),
            )
            li = int(layer_index)
            bid = bids[li] if li < len(bids) else (
                bids[-1] if len(bids) > 0 else fallback_id
            )
        except Exception:
            bid = fallback_id
    bt = _element_to_bar_type(doc, bid)
    if bt is not None:
        return bt
    return _element_to_bar_type(doc, fallback_id)


def _bar_type_for_cabezal_stack(
    doc, walls, cabezal_por_muro_id, segment_ctx, extremo,
    stack_index, layer_index, fallback_id=None,
):
    """Resuelve ø en un índice de stack (segmentos troceo o config del muro)."""
    try:
        si = int(stack_index)
    except Exception:
        si = 0
    if segment_ctx:
        bt = _bar_type_from_segment_context(
            doc, segment_ctx, si, layer_index, fallback_id,
        )
        if bt is not None:
            return bt
    if walls and cabezal_por_muro_id and extremo and 0 <= si < len(walls):
        try:
            wid = wall_id_int(walls[si])
            ex_cfg = ((cabezal_por_muro_id.get(wid) or {}).get(extremo) or {})
            layers = _cabezal_ex_cfg_layers(ex_cfg)
            li = int(layer_index)
            ly = layers[li] if li < len(layers) else {}
            return _resolver_bar_type_for_layer(
                doc, ex_cfg, ly, fallback_id,
            )
        except Exception:
            pass
    return _element_to_bar_type(doc, fallback_id)


def _n_bars_for_stack_index(cabezal_por_muro_id, walls, extremo, stack_index, layer_index):
    if not walls or stack_index < 0 or stack_index >= len(walls):
        return CABEZAL_MIN_BARRAS_POR_CAPA
    try:
        wid = wall_id_int(walls[stack_index])
    except Exception:
        return CABEZAL_MIN_BARRAS_POR_CAPA
    cfg = (cabezal_por_muro_id or {}).get(wid) or {}
    ex_cfg = cfg.get(extremo) or {}
    layers = _cabezal_ex_cfg_layers(ex_cfg)
    li = int(layer_index)
    ly = layers[li] if li < len(layers) else {}
    try:
        nb = int((ly or {}).get(u"n_bars", CABEZAL_MIN_BARRAS_POR_CAPA))
    except Exception:
        nb = CABEZAL_MIN_BARRAS_POR_CAPA
    return max(
        CABEZAL_MIN_BARRAS_POR_CAPA,
        min(CABEZAL_MAX_BARRAS_POR_CAPA, nb),
    )


def _normalize_cabezal_layer_dict(ly, fallback_bar_type_id=None, n_tramos=1):
    """Garantiza claves mínimas en una capa (``n_bars`` + ``bar_type_id`` legacy)."""
    if not ly or not isinstance(ly, dict):
        ly = {}
    out = dict(ly)
    try:
        nb = int(out.get(u"n_bars", CABEZAL_MIN_BARRAS_POR_CAPA))
    except Exception:
        nb = CABEZAL_MIN_BARRAS_POR_CAPA
    out[u"n_bars"] = max(
        CABEZAL_MIN_BARRAS_POR_CAPA,
        min(CABEZAL_MAX_BARRAS_POR_CAPA, nb),
    )
    bid = out.get(u"bar_type_id")
    if bid is None or bid == ElementId.InvalidElementId:
        bid = fallback_bar_type_id
    out[u"bar_type_id"] = bid
    return out


def _normalize_cabezal_extremo_layers(ex_cfg, n_tramos=1):
    """Normaliza ``layers`` in-place; retorna lista de capas."""
    if not ex_cfg or not isinstance(ex_cfg, dict):
        return []
    fb = ex_cfg.get(u"bar_type_id")
    raw = _cabezal_as_python_list(ex_cfg.get(u"layers"))
    layers = [
        _normalize_cabezal_layer_dict(ly, fb)
        for ly in raw
    ]
    try:
        n_capas = int(ex_cfg.get(u"n_capas", len(layers) or CABEZAL_MIN_CAPAS))
    except Exception:
        n_capas = len(layers) or CABEZAL_MIN_CAPAS
    if cabezal_extremo_es_encuentro_l(ex_cfg):
        n_capas = max(2, min(CABEZAL_MAX_CAPAS, n_capas))
    else:
        n_capas = max(
            CABEZAL_MIN_CAPAS,
            min(CABEZAL_MAX_CAPAS, n_capas),
        )
    if not layers:
        n_capas = max(n_capas, CABEZAL_MIN_CAPAS)
    while len(layers) < CABEZAL_MAX_CAPAS:
        prev = default_cabezal_layer_config(2, fb)
        if layers and len(layers) > 0:
            prev = layers[-1]
        try:
            prev_nb = int(prev.get(u"n_bars", CABEZAL_MIN_BARRAS_POR_CAPA))
        except Exception:
            prev_nb = CABEZAL_MIN_BARRAS_POR_CAPA
        prev_bid = prev.get(u"bar_type_id") or fb
        layers.append(default_cabezal_layer_config(prev_nb, prev_bid))
    layers = [
        _normalize_cabezal_layer_dict(ly, fb)
        for ly in layers[:CABEZAL_MAX_CAPAS]
    ]
    ex_cfg[u"layers"] = layers
    ex_cfg[u"n_capas"] = n_capas
    if u"segment_bar_type_ids" not in ex_cfg:
        ex_cfg[u"segment_bar_type_ids"] = {}
    ex_cfg[u"confinement"] = normalize_cabezal_confinement(
        ex_cfg.get(u"confinement"), n_capas,
    )
    if cabezal_extremo_es_encuentro_l(ex_cfg) and _cab_enc_l is not None:
        for ly in layers[:n_capas]:
            try:
                nb = int(ly.get(u"n_bars", CABEZAL_MIN_BARRAS_POR_CAPA))
            except Exception:
                nb = CABEZAL_MIN_BARRAS_POR_CAPA
            ly[u"n_bars"] = max(2, min(CABEZAL_MAX_BARRAS_POR_CAPA, nb))
    return layers


def cabezal_active_layers(ex_cfg):
    """Capas usadas al crear barras / preview (según ``n_capas``)."""
    if not ex_cfg or not isinstance(ex_cfg, dict):
        return []
    layers = _cabezal_ex_cfg_layers(ex_cfg)
    try:
        n_capas = int(ex_cfg.get(u"n_capas", len(layers) or CABEZAL_MIN_CAPAS))
    except Exception:
        n_capas = len(layers) or CABEZAL_MIN_CAPAS
    n_capas = max(
        CABEZAL_MIN_CAPAS,
        min(CABEZAL_MAX_CAPAS, n_capas),
    )
    if not layers:
        return []
    return list(layers[:n_capas])


def malla_n_capas_activas_raw(ex_cfg):
    """
    Capas de cabezal activas en un extremo (sin piso mínimo).

    Sin ``ex_cfg`` → 1 (default sin cabezal). Cabezal apagado → 0.
    """
    if ex_cfg is None:
        return 1
    if not cabezal_extremo_armado_activo(ex_cfg):
        return 0
    try:
        n_capas = int(ex_cfg.get(u"n_capas"))
    except Exception:
        n_capas = len(cabezal_active_layers(ex_cfg))
    if n_capas < 1:
        n_capas = len(cabezal_active_layers(ex_cfg)) or CABEZAL_MIN_CAPAS
    return min(
        max(CABEZAL_MIN_CAPAS, int(n_capas)),
        CABEZAL_MAX_CAPAS,
    )


def malla_n_remove_por_extremo(ex_cfg):
    """
    Barras de malla a excluir en un extremo = capas activas del cabezal.

    Capa cabezal índice ``k`` (0…5) → la ``k``-ésima barra de malla contigua a ese
    extremo. Sin cabezal → 1; cabezal apagado en el extremo → 0.
    """
    if ex_cfg is None:
        return 1
    if not cabezal_extremo_armado_activo(ex_cfg):
        return 0
    layers = cabezal_active_layers(ex_cfg)
    if layers:
        return min(len(layers), CABEZAL_MAX_CAPAS)
    return malla_n_capas_activas_raw(ex_cfg)


def _rebar_coincide_tipo_capas_malla(rebar, params_dict, layer_keys):
    """True si ``GetTypeId()`` coincide con algún tipo en ``layer_keys`` del panel."""
    if rebar is None or not params_dict or not layer_keys:
        return False
    try:
        from Autodesk.Revit.DB import ElementId
    except Exception:
        ElementId = None
    try:
        tid = rebar.GetTypeId()
    except Exception:
        return False
    for lk in layer_keys:
        try:
            bid = params_dict.get(lk, (None, u""))[0]
        except Exception:
            bid = None
        if (
            bid is not None
            and ElementId is not None
            and bid != ElementId.InvalidElementId
            and tid == bid
        ):
            return True
    return False


def rebar_coincide_tipo_capa_malla_vertical(rebar, params_dict, muro_contencion=False):
    """True si el ``Rebar`` pertenece a una capa vertical de malla (minor / major contención)."""
    if muro_contencion:
        keys = (u"exterior_major", u"interior_major")
    else:
        keys = (u"exterior_minor", u"interior_minor")
    return _rebar_coincide_tipo_capas_malla(rebar, params_dict, keys)


def rebar_coincide_tipo_capa_malla_horizontal(rebar, params_dict, muro_contencion=False):
    """True si el ``Rebar`` coincide con capa horizontal de malla (major / minor contención)."""
    if muro_contencion:
        keys = (u"exterior_minor", u"interior_minor")
    else:
        keys = (u"exterior_major", u"interior_major")
    return _rebar_coincide_tipo_capas_malla(rebar, params_dict, keys)


def aplicar_exclusion_horizontal_malla_ultima_barra(rebar, doc=None):
    u"""Remove last bar en cada set horizontal de malla (exterior/interior)."""
    if rebar is None:
        return False
    try:
        from armado_muros_nodo_shared import ajustar_inclusion_extremos_rebar_set_con_fallback
    except Exception:
        return False
    return bool(
        ajustar_inclusion_extremos_rebar_set_con_fallback(rebar, doc, True, False),
    )


def rebar_es_malla_vertical_por_tipo(rebar, params_dict, muro_contencion=False):
    """
    Tipo de capa vertical sin ambigüedad con la horizontal del mismo muro.

    Doble malla S.I.C. (mismo Ø en las 4 capas): un set horizontal comparte
    ``ElementId`` con minor/major; en ese caso retorna False y la orientación
    debe resolverse por geometría.
    """
    if not rebar_coincide_tipo_capa_malla_vertical(
        rebar, params_dict, muro_contencion,
    ):
        return False
    if rebar_coincide_tipo_capa_malla_horizontal(
        rebar, params_dict, muro_contencion,
    ):
        return False
    return True


def malla_indices_lineas_a_excluir(n_lines, ex_cfg_inicio, ex_cfg_fin):
    """
    Índices de líneas verticales de malla a excluir según cabezal ini/fin.

    Capa cabezal ``k`` (0…5) ↔ barra ``k`` contigua a ese extremo del muro.
    Si hay solapamiento ini+fin, se prioriza inicio.

    Rebar set malla vertical (post Remove System), alineado con ``LocationCurve``:
    índice ``0`` = **inicio** (P0); índice ``n-1`` = **fin / término** (P1).
    """
    try:
        n = int(n_lines)
    except Exception:
        n = 0
    if n < 1:
        return []
    ni = malla_n_remove_por_extremo(ex_cfg_inicio)
    nf = malla_n_remove_por_extremo(ex_cfg_fin)
    ni = max(0, min(int(ni), n))
    nf = max(0, min(int(nf), max(0, n - ni)))
    excl = set()
    for k in range(ni):
        excl.add(k)
    for k in range(nf):
        excl.add(n - 1 - k)
    return sorted(excl)


def cabezal_extremos_config_for_muro(cabezal_por_muro_id, wall_id):
    """``(ex_cfg_inicio, ex_cfg_fin)`` para correlación malla; ``None`` si no hay cabezal."""
    if not cabezal_por_muro_id:
        return None, None
    wid = normalize_muro_id_key(wall_id)
    if wid is None:
        return None, None
    cfg = cabezal_por_muro_id.get(wid)
    if not cfg:
        try:
            cfg = cabezal_por_muro_id.get(int(wid))
        except Exception:
            cfg = None
    if not cfg:
        cfg = {}
    if not cfg:
        return None, None
    ex_ini = cfg.get(CABEZAL_EXTREMO_INICIO)
    ex_fin = cfg.get(CABEZAL_EXTREMO_FIN)
    try:
        if ex_ini:
            _normalize_cabezal_extremo_layers(ex_ini)
        if ex_fin:
            _normalize_cabezal_extremo_layers(ex_fin)
    except Exception:
        pass
    return ex_ini, ex_fin


def aplicar_exclusion_verticales_malla_rebar(
    rebar, ex_cfg_inicio, ex_cfg_fin, doc=None, host=None,
):
    """
    Excluye barras verticales de malla según capas cabezal (post Remove System).

    Usa ``SetBarIncluded`` por índice (todas las posiciones 0…n_remove-1 por extremo).
    """
    if rebar is None:
        return False
    try:
        from armado_muros_nodo_shared import (
            _excluir_barras_por_indices,
            _rebar_bar_included,
            _rebar_cantidad_posiciones,
        )
    except Exception:
        return False
    if host is None and doc is not None:
        try:
            host = doc.GetElement(rebar.GetHostId())
        except Exception:
            host = None
    try:
        n = int(_rebar_cantidad_posiciones(rebar))
    except Exception:
        n = 0
    if n < 1:
        return False
    indices = malla_indices_lineas_a_excluir(n, ex_cfg_inicio, ex_cfg_fin)
    if not indices:
        return False
    ok = _excluir_barras_por_indices(rebar, indices, doc=doc)
    if doc is not None:
        try:
            doc.Regenerate()
        except Exception:
            pass
    if ok:
        return True
    try:
        pending = [i for i in indices if _rebar_bar_included(rebar, i)]
    except Exception:
        pending = list(indices)
    return len(pending) < len(indices)


def _troceo_por_muro_from_extremo_cfg(ex_cfg):
    """True si este extremo usa la Z base del muro como plano de troceo."""
    if not ex_cfg or not isinstance(ex_cfg, dict):
        return False
    try:
        return bool(ex_cfg.get(u"troceo_por_muro"))
    except Exception:
        return False


def troceo_override_from_extremo_cfg(ex_cfg):
    """None = seguir auto geom.; bool = ajuste puntual en el pie."""
    if not ex_cfg or not isinstance(ex_cfg, dict):
        return None
    if u"troceo_por_muro_override" not in ex_cfg:
        return None
    ov = ex_cfg.get(u"troceo_por_muro_override")
    if ov is None:
        return None
    return bool(ov)


def merge_troceo_effective(auto_val, ex_cfg):
    ov = troceo_override_from_extremo_cfg(ex_cfg)
    if ov is None:
        return bool(auto_val)
    return bool(ov)


def _troceo_u_tolerance_ft():
    try:
        return UnitUtils.ConvertToInternalUnits(2.0, UnitTypeId.Millimeters)
    except Exception:
        return 2.0 / 304.8


def _stack_row_index_for_ordered_index(n_walls, ordered_index):
    return max(0, int(n_walls) - 1 - int(ordered_index))


def _wall_extremo_u_on_stack(wall, ordered_index, walls_ordered, stacked_layout, extremo):
    if wall is None or not stacked_layout:
        return None
    walls = list(walls_ordered or [])
    n = len(walls)
    ri = _stack_row_index_for_ordered_index(n, ordered_index)
    items = stacked_layout.get(u"items") or []
    if not (0 <= ri < len(items)):
        return None
    item = items[ri]
    try:
        from armado_muros_lineales import cabezal_extremos_en_lados_stacked
        ex_izq, ex_der = cabezal_extremos_en_lados_stacked(wall, ri, stacked_layout)
    except Exception:
        ex_izq, ex_der = CABEZAL_EXTREMO_INICIO, CABEZAL_EXTREMO_FIN
    try:
        if extremo == ex_izq:
            return float(item.get(u"u_start", item.get(u"u0", 0.0)))
        if extremo == ex_der:
            return float(item.get(u"u_end", item.get(u"u1", 0.0)))
        if extremo == CABEZAL_EXTREMO_INICIO:
            return float(item.get(u"u0", item.get(u"u_start", 0.0)))
        return float(item.get(u"u1", item.get(u"u_end", 0.0)))
    except Exception:
        return None


def geometry_suggests_troceo_at_stack_index(
    walls_ordered,
    stacked_layout,
    extremo,
    stack_index,
    u_tol_ft=None,
):
    if stack_index < 1:
        return False
    walls = list(walls_ordered or [])
    if stack_index >= len(walls):
        return False
    if u_tol_ft is None:
        u_tol_ft = _troceo_u_tolerance_ft()
    w_prev = walls[stack_index - 1]
    w_cur = walls[stack_index]
    if _thickness_mm_key_for_fusion(
        _wall_thickness_mm_for_fusion(w_prev),
    ) != _thickness_mm_key_for_fusion(
        _wall_thickness_mm_for_fusion(w_cur),
    ):
        return True
    u_prev = _wall_extremo_u_on_stack(
        w_prev, stack_index - 1, walls, stacked_layout, extremo,
    )
    u_cur = _wall_extremo_u_on_stack(
        w_cur, stack_index, walls, stacked_layout, extremo,
    )
    if u_prev is None or u_cur is None:
        return False
    return abs(float(u_cur) - float(u_prev)) > float(u_tol_ft)


def compute_auto_troceo_por_muro_flags(walls_ordered, stacked_layout, extremo):
    n = len(walls_ordered or [])
    flags = [False] * n
    if n < 2:
        return flags
    for i in range(1, n):
        flags[i] = geometry_suggests_troceo_at_stack_index(
            walls_ordered, stacked_layout, extremo, i,
        )
    return flags


def sync_troceo_effective_for_extremo(
    walls_ordered,
    cabezal_por_muro_id,
    stacked_layout,
    extremo,
):
    """Recalcula troceo_por_muro efectivo (auto + override puntual) por muro."""
    flags = compute_auto_troceo_por_muro_flags(
        walls_ordered, stacked_layout, extremo,
    )
    for i, wall in enumerate(walls_ordered or []):
        if wall is None:
            continue
        try:
            wid = wall_id_int(wall)
        except Exception:
            continue
        cfg = (cabezal_por_muro_id or {}).setdefault(
            wid, default_cabezal_muro_config(),
        )
        ex_cfg = cfg.setdefault(extremo, default_cabezal_extremo_config())
        auto_i = bool(flags[i]) if i >= 1 else False
        ex_cfg[u"troceo_auto_geom"] = auto_i
        ex_cfg[u"troceo_por_muro"] = merge_troceo_effective(auto_i, ex_cfg)


def _count_troceo_walls_for_extremo(walls, cabezal_por_muro_id, extremo):
    """Cuenta muros con empalme activo en un extremo (para N tramos verticales)."""
    n = 0
    if not walls or not cabezal_por_muro_id:
        return n
    for wall in walls:
        if wall is None:
            continue
        try:
            wid = wall_id_int(wall)
        except Exception:
            continue
        cfg = cabezal_por_muro_id.get(wid) or {}
        ex_cfg = cfg.get(extremo) or {}
        if _troceo_por_muro_from_extremo_cfg(ex_cfg):
            n += 1
    return n


def default_cabezal_muro_config():
    return {
        CABEZAL_EXTREMO_INICIO: default_cabezal_extremo_config(),
        CABEZAL_EXTREMO_FIN: default_cabezal_extremo_config(),
    }


def cabezal_longitudinal_sic_por_espesor_mm(e_mm):
    """
    Armadura longitudinal típica en borde de muro (S.I.C.) — cantidad y Ø (mm).

    Valores explícitos de cálculo (150/200 → 2Ø12, 250 → 2Ø16, …) con rangos
    en espesores intermedios (fronteras 225, 275, 325, 500, 700 mm).
    Defaults de UI; el usuario puede modificarlos antes de crear.
    """
    try:
        e = int(round(float(e_mm)))
    except Exception:
        e = 200
    if e <= 224:
        return 2, 12
    if e <= 274:
        return 2, 16
    if e <= 324:
        return 3, 16
    if e <= 499:
        return 3, 18
    if e <= 699:
        return 4, 22
    return 4, 22


def cabezal_longitudinal_sic_espesor_mm_para_muro(wall):
    """Espesor nominal del muro (mm) para tabla S.I.C. de cabezal."""
    if wall is None:
        return 200.0
    if obtener_espesor_muro_mm_approx is not None:
        try:
            e = obtener_espesor_muro_mm_approx(wall)
            if e is not None:
                return float(e)
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


def cabezal_longitudinal_sic_defaults_para_muro(wall):
    """``(n_bars, diam_mm)`` según espesor del ``Wall`` (S.I.C. borde de muro)."""
    return cabezal_longitudinal_sic_por_espesor_mm(
        cabezal_longitudinal_sic_espesor_mm_para_muro(wall),
    )


def cabezal_extremo_config_con_sic_longitudinal(
    doc,
    wall,
    fallback_bar_type_id=None,
    fallback_conf_bar_type_id=None,
):
    """
    Configuración de extremo con 1 capa: ``n_bars`` y Ø longitudinal S.I.C. por espesor.
    """
    n_bars, diam_mm = cabezal_longitudinal_sic_defaults_para_muro(wall)
    n_bars = max(
        CABEZAL_MIN_BARRAS_POR_CAPA,
        min(CABEZAL_MAX_BARRAS_POR_CAPA, int(n_bars)),
    )
    fb_bt = None
    if fallback_bar_type_id not in (None, ElementId.InvalidElementId):
        fb_bt = _element_to_bar_type(doc, fallback_bar_type_id)
    bt = _bar_type_for_catalog_diameter_mm(doc, diam_mm, fb_bt)
    bt_id = None
    if bt is not None:
        try:
            bt_id = bt.Id
        except Exception:
            bt_id = None
    ex_cfg = default_cabezal_extremo_config()
    ex_cfg[u"n_capas"] = 1
    ex_cfg[u"layers"] = [default_cabezal_layer_config(n_bars, bt_id)]
    if bt_id is not None:
        ex_cfg[u"bar_type_id"] = bt_id
    if fallback_conf_bar_type_id not in (None, ElementId.InvalidElementId):
        ex_cfg[u"conf_bar_type_id"] = fallback_conf_bar_type_id
    cabezal_sync_confinement_from_extremo(
        ex_cfg,
        doc,
        fallback_bar_type_id or bt_id,
    )
    return ex_cfg


def filas_indices_por_n_barras(n_bars):
    n = int(n_bars)
    if n == 2:
        return (0, 3)
    if n == 3:
        return (0, 1, 3)
    if n == 4:
        return (0, 1, 2, 3)
    return ()


def validar_cabezal_config(cfg):
    """Retorna (ok, mensaje)."""
    if not cfg or not isinstance(cfg, dict):
        return False, u"Configuración de cabezal vacía."
    for ex in CABEZAL_EXTREMOS:
        ex_cfg = cfg.get(ex)
        if not ex_cfg:
            return False, u"Falta extremo «{0}».".format(ex)
        if not cabezal_extremo_armado_activo(ex_cfg):
            continue
        layers = ex_cfg.get(u"layers")
        if not layers:
            return False, u"Extremo «{0}»: sin capas.".format(ex)
        try:
            n_capas = int(ex_cfg.get(u"n_capas", len(layers)))
        except Exception:
            n_capas = len(layers)
        min_capas = 2 if cabezal_extremo_es_encuentro_l(ex_cfg) else CABEZAL_MIN_CAPAS
        if n_capas < min_capas or n_capas > CABEZAL_MAX_CAPAS:
            return False, u"Extremo «{0}»: n_capas={1} fuera de rango [{2},{3}].".format(
                ex, n_capas, min_capas, CABEZAL_MAX_CAPAS,
            )
        active = list(layers[:n_capas])
        if len(active) < n_capas:
            return False, u"Extremo «{0}»: capas insuficientes para n_capas={1}.".format(
                ex, n_capas,
            )
        for i, ly in enumerate(active):
            try:
                nb = int(ly.get(u"n_bars", 0))
            except Exception:
                nb = 0
            min_bars = 2 if cabezal_extremo_es_encuentro_l(ex_cfg) else CABEZAL_MIN_BARRAS_POR_CAPA
            if nb < min_bars or nb > CABEZAL_MAX_BARRAS_POR_CAPA:
                return False, u"Extremo «{0}», capa {1}: barras deben ser {2}–{3}.".format(
                    ex,
                    i + 1,
                    min_bars,
                    CABEZAL_MAX_BARRAS_POR_CAPA,
                )
    return True, u""


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _wall_z_bounds_ft(wall):
    try:
        bb = wall.get_BoundingBox(None)
        if bb is not None:
            z0 = float(bb.Min.Z)
            z1 = float(bb.Max.Z)
            if z1 > z0 + 1e-9:
                return z0, z1
    except Exception:
        pass
    try:
        lc = location_curve_wall(wall) if location_curve_wall else None
        if lc is not None:
            z0 = float(lc.GetEndPoint(0).Z)
            z1 = float(lc.GetEndPoint(1).Z)
            return min(z0, z1), max(z0, z1)
    except Exception:
        pass
    if _wall_bottom_z_sort_key_ft is not None:
        zb = float(_wall_bottom_z_sort_key_ft(wall))
        return zb, zb + 3.0
    return 0.0, 3.0


def altura_muro_mm_approx(wall):
    z0, z1 = _wall_z_bounds_ft(wall)
    try:
        return UnitUtils.ConvertFromInternalUnits(z1 - z0, UnitTypeId.Millimeters)
    except Exception:
        return (z1 - z0) * 304.8


def format_cabezal_mm_es(mm):
    """Entero mm con separador de miles (punto) para textos UI."""
    try:
        n = int(round(float(mm)))
    except Exception:
        return u"0"
    if n < 0:
        sign = u"-"
        n = abs(n)
    else:
        sign = u""
    s = str(n)
    parts = []
    while len(s) > 3:
        parts.insert(0, s[-3:])
        s = s[:-3]
    parts.insert(0, s)
    return sign + u".".join(parts)


def infer_tramo_bar_length_mm(walls, seg):
    """
    Largo vertical inferido (mm) de barra longitudinal en un tramo cabezal.

    Envelope Z de los muros del segmento (``wall_indices``). No incluye patas L
    ni solapes finos de empalme; suficiente para aviso preventivo en UI.
    """
    indices = list((seg or {}).get(u"wall_indices") or [])
    if not indices or not walls:
        return 0.0
    z_lo = None
    z_hi = None
    for wi in indices:
        try:
            wi = int(wi)
        except Exception:
            continue
        if not (0 <= wi < len(walls)):
            continue
        wall = walls[wi]
        if wall is None:
            continue
        z0, z1 = _wall_z_bounds_ft(wall)
        if z_lo is None or z0 < z_lo:
            z_lo = z0
        if z_hi is None or z1 > z_hi:
            z_hi = z1
    if z_lo is None or z_hi is None:
        return 0.0
    span_ft = max(0.0, float(z_hi) - float(z_lo))
    try:
        return float(
            UnitUtils.ConvertFromInternalUnits(span_ft, UnitTypeId.Millimeters),
        )
    except Exception:
        return span_ft * 304.8


def cabezal_tramo_bar_length_status(walls, seg):
    """``(exceeds_max_comercial, length_mm)`` para el tramo ``seg``."""
    L = infer_tramo_bar_length_mm(walls, seg)
    return L > float(CABEZAL_MAX_BARRA_COMERCIAL_MM), L


def cabezal_tramo_bar_length_warn_tooltip(length_mm):
    """Texto ToolTip del aviso compacto (variante B)."""
    max_mm = float(CABEZAL_MAX_BARRA_COMERCIAL_MM)
    try:
        L = float(length_mm or 0.0)
    except Exception:
        L = 0.0
    lines = [
        u"L inferido: {0} mm".format(format_cabezal_mm_es(L)),
        u"Máximo comercial: {0} mm".format(format_cabezal_mm_es(max_mm)),
    ]
    if L > max_mm:
        lines.append(
            u"Exceso: +{0} mm".format(format_cabezal_mm_es(L - max_mm)),
        )
        lines.append(
            u"Recomendación: subdivida el tramo activando empalme/troceo "
            u"en uno o más muros intermedios.",
        )
    else:
        lines.append(u"Dentro del límite.")
    return u"\n".join(lines)


def cabezal_tramo_bar_length_warn_compact_text(length_mm):
    """Texto visible del aviso compacto al pie del controlador."""
    L_txt = format_cabezal_mm_es(length_mm)
    return (
        u"Barra inferida {0} mm — subdivida el tramo con empalmes".format(L_txt)
    )


def _wall_extremo_frame(wall, extremo):
    """
    Marco en extremo del muro.

    ``station``: punto del extremo (LocationCurve).
    ``into_wall``: unitario hacia el interior del muro.
    ``inward``: unitario desde cara exterior hacia el interior (espesor).
    ``thickness_ft``: espesor nominal en pies.
    """
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

    if extremo == CABEZAL_EXTREMO_FIN:
        station = p1
        into_wall = t_hat.Negate()
    else:
        station = p0
        into_wall = t_hat

    inward = None
    try:
        orient = wall.Orientation
        if orient is not None and float(orient.GetLength()) > 1e-9:
            ext_hat = orient.Normalize()
            inward = ext_hat.Negate()
    except Exception:
        inward = None
    if inward is None:
        try:
            inward = into_wall.CrossProduct(XYZ.BasisZ)
            il = float(inward.GetLength())
            if il < 1e-12:
                inward = XYZ.BasisY
            else:
                inward = inward.Normalize()
        except Exception:
            inward = XYZ.BasisY

    th_mm = 200.0
    if obtener_espesor_muro_mm_approx is not None:
        try:
            th_mm = float(obtener_espesor_muro_mm_approx(wall) or 200.0)
        except Exception:
            th_mm = 200.0
    th_ft = _mm_to_internal(max(th_mm, 50.0))

    return {
        u"station": station,
        u"into_wall": into_wall,
        u"inward": inward,
        u"thickness_ft": th_ft,
    }


def _delete_wall_extremo_markers(doc, marker_ids):
    """Borra marcadores 3D huérfanos (legacy). Debe llamarse dentro de una Transaction."""
    if not marker_ids:
        return
    for eid in marker_ids:
        try:
            doc.Delete(eid)
        except Exception:
            pass


def _clamp_cover_frac(cover_mm, thickness_mm):
    e = max(float(thickness_mm or 200.0), 50.0)
    c = max(0.0, float(cover_mm or CABEZAL_COVER_MM))
    return min(0.38, max(0.04, c / e))


def _fila_y_normalizada_en_zona_util(row_index):
    """Posición vertical 0..1 en zona útil (filas 0..3)."""
    ri = int(row_index)
    if ri <= 0:
        return 0.0
    if ri >= 3:
        return 1.0
    return float(ri) / 3.0


def _cabezal_layer_pitch_mm(thickness_mm, n_capas, cover_mm=None):
    """Paso entre capas (mm) a lo largo del espesor; comprime si no caben."""
    e_mm = max(float(thickness_mm or 200.0), 50.0)
    c_mm = max(0.0, float(cover_mm if cover_mm is not None else CABEZAL_COVER_MM))
    usable_mm = max(40.0, e_mm - 2.0 * c_mm)
    n = max(1, int(n_capas))
    pitch = float(CABEZAL_LAYER_PITCH_MM)
    if n > 1:
        need = float(n - 1) * pitch
        if need > usable_mm * 0.88:
            pitch = max(
                float(CABEZAL_LAYER_PITCH_MIN_MM),
                (usable_mm * 0.88) / float(n - 1),
            )
    return pitch


def _cabezal_layer_center_frac_x(layer_index, thickness_mm, n_capas, cover_mm=None):
    """
    Fracción X (0 = cara cabezal) del eje de la capa ``layer_index``.
    Capa 1 en recubrimiento; capas 2+ hacia la derecha con paso fijo.
    """
    e_mm = max(float(thickness_mm or 200.0), 50.0)
    c_mm = float(cover_mm if cover_mm is not None else CABEZAL_COVER_MM)
    fx_cov = _clamp_cover_frac(c_mm, e_mm)
    pitch_frac = _cabezal_layer_pitch_mm(e_mm, n_capas, c_mm) / e_mm
    cx = fx_cov + float(layer_index) * pitch_frac
    right_lim = 1.0 - fx_cov - 0.012
    return min(cx, right_lim)


def cabezal_stirrup_preview_rect(dots, layer_indices, pad_frac=0.04):
    """Rectángulo envolvente (fracciones fx/fy) para preview WPF."""
    if not dots or not layer_indices:
        return None
    idx_set = set(int(i) for i in layer_indices)
    subset = [
        d for d in dots
        if int(d.get(u"layer_index", -1)) in idx_set
    ]
    if not subset:
        return None
    fxs = [float(d[u"fx"]) for d in subset]
    fys = [float(d[u"fy"]) for d in subset]
    pad = max(0.0, float(pad_frac))
    return {
        u"fx0": min(fxs) - pad,
        u"fx1": max(fxs) + pad,
        u"fy0": min(fys) - pad,
        u"fy1": max(fys) + pad,
    }


def cabezal_stirrup_preview_segments(stirrup_rect, hook_frac=0.22):
    """
    Segmentos del estribo en coordenadas normalizadas (fx, fy).

    Loop CW como en creación: BR→BL→TL→TR; ganchos 135° en BR (dos patas).
    """
    if not stirrup_rect:
        return None
    fx0 = float(stirrup_rect[u"fx0"])
    fx1 = float(stirrup_rect[u"fx1"])
    fy0 = float(stirrup_rect[u"fy0"])
    fy1 = float(stirrup_rect[u"fy1"])
    fx_lo, fx_hi = min(fx0, fx1), max(fx0, fx1)
    fy_lo, fy_hi = min(fy0, fy1), max(fy0, fy1)
    br = (fx_hi, fy_hi)
    bl = (fx_lo, fy_hi)
    tl = (fx_lo, fy_lo)
    tr = (fx_hi, fy_lo)
    loop = (
        (br, bl),
        (bl, tl),
        (tl, tr),
        (tr, br),
    )
    span = min(fx_hi - fx_lo, fy_hi - fy_lo)
    hook = max(0.015, float(hook_frac) * span)
    hooks = (
        (br, (br[0] - hook, br[1])),
        (br, (br[0], br[1] - hook)),
    )
    return list(loop) + list(hooks)


def cabezal_tie_preview_geometry(
    dots,
    layer_index=CABEZAL_TIE_LAYER_INDEX,
    hook_frac=0.12,
    inner_y0=None,
    inner_h=None,
    pitch_frac=None,
    bar_diam_mm=12.0,
    tie_diam_mm=None,
):
    """
    Traba capa [1] en coordenadas normalizadas (fx, fy).

    Pata en tangente interior (lado opuesto a capa [0]), empalmes horizontales por
    barra y ganchos 135° hacia las barras en cada extremo del espesor.
    """
    if not dots:
        return None
    try:
        li = int(layer_index)
    except Exception:
        li = CABEZAL_TIE_LAYER_INDEX
    subset = [
        d for d in dots
        if int(d.get(u"layer_index", -1)) == li
    ]
    if not subset:
        return None
    bar_fx = sum(float(d[u"fx"]) for d in subset) / float(len(subset))
    fys = [float(d[u"fy"]) for d in subset]
    iy0 = float(inner_y0 if inner_y0 is not None else min(fys) - 0.02)
    ih = float(inner_h if inner_h is not None else max(0.20, 1.0 - 2.0 * iy0))
    fy_top = iy0
    fy_bot = iy0 + ih

    offset_mm = _cabezal_tie_offset_mm(
        None, tie_diam_mm, bar_diam_mm_fallback=bar_diam_mm,
    )
    pf = float(pitch_frac) if pitch_frac is not None else 0.0
    if pf <= 1e-9:
        ref_li = max(0, li - 1)
        layer_ref = [
            d for d in dots
            if int(d.get(u"layer_index", -1)) == ref_li
        ]
        if layer_ref:
            pf = abs(bar_fx - float(layer_ref[0][u"fx"]))
    if pf > 1e-9:
        tie_fx = bar_fx + (offset_mm / float(CABEZAL_LAYER_PITCH_MM)) * pf
    else:
        tie_fx = bar_fx + offset_mm / max(float(CABEZAL_LAYER_PITCH_MM), 50.0)

    hook_len = max(0.015, float(hook_frac) * ih)
    cos45 = 0.707106781

    leg = ((tie_fx, fy_top), (tie_fx, fy_bot))
    # Ganchos hacia barras (interior, −fx desde tangente exterior de capa [1]).
    top_hook = ((tie_fx, fy_top), (tie_fx - hook_len * cos45, fy_top + hook_len * cos45))
    bottom_hook = (
        (tie_fx, fy_bot),
        (tie_fx - hook_len * cos45, fy_bot - hook_len * cos45),
    )
    bar_grips = []
    for d in subset:
        fy = float(d[u"fy"])
        bx = float(d[u"fx"])
        bar_grips.append(((tie_fx, fy), (bx, fy)))

    return {
        u"layer_index": li,
        u"tie_fx": tie_fx,
        u"bar_fx": bar_fx,
        u"fy_top": fy_top,
        u"fy_bot": fy_bot,
        u"leg": leg,
        u"top_hook": top_hook,
        u"bottom_hook": bottom_hook,
        u"bar_grips": bar_grips,
        u"segments": [leg, top_hook, bottom_hook] + bar_grips,
    }


def cabezal_seccion_preview_layout(thickness_mm, layers, cover_mm=None,
                                    row_inset_frac=0.10,
                                    row_inset_frac_x=None,
                                    row_inset_frac_y=None,
                                    pitch_frac_override=None,
                                    draw_w_px=None,
                                    layer_pitch_px=None,
                                    draw_h_px=None,
                                    bar_span_px=None,
                                    confinement_type=None,
                                    stirrup_pad_frac=0.04,
                                    confinement_stirrup_diam_mm=None):
    """
    Geometría del corte de cabezal para el canvas WPF.

    Convención (alineada con creación multicapa):
    - X (fx): profundidad longitudinal de la capa hacia el interior del muro.
    - Y (fy): reparto de barras en el espesor (SetLayoutAsFixedNumber).
    """
    layers = list(layers or [])
    n_capas = max(1, len(layers))
    e_mm = max(float(thickness_mm or 200.0), 50.0)
    c_mm = float(cover_mm if cover_mm is not None else CABEZAL_COVER_MM)
    _default = min(0.30, max(0.06, float(row_inset_frac)))
    fx_cov = min(0.40, max(0.02, float(row_inset_frac_x))) if row_inset_frac_x is not None else _default
    fy_cov = min(0.40, max(0.02, float(row_inset_frac_y))) if row_inset_frac_y is not None else _default
    inner_x0 = fx_cov
    inner_y0 = fy_cov
    inner_h = max(0.20, 1.0 - 2.0 * fy_cov)
    spacing_mm = float(CABEZAL_LAYER_PITCH_MM)
    conf_diam_mm = 16.0
    bar_diam_mm = 12.0
    offset_trans_nom = c_mm + conf_diam_mm + bar_diam_mm * 0.5

    ref_intervals = 5.0
    layer_zone = max(0.15, 1.0 - fx_cov - 0.012)
    if draw_w_px is not None and layer_pitch_px is not None:
        dw = max(1.0, float(draw_w_px))
        pitch_frac = max(1.0, float(layer_pitch_px)) / dw
    elif pitch_frac_override is not None and float(pitch_frac_override) > 0:
        pitch_frac = float(pitch_frac_override)
    elif n_capas <= 1:
        pitch_frac = layer_zone * 0.5
    else:
        # Paso fijo desde punta (~150 mm); comprimir solo si no caben N capas.
        pitch_frac = layer_zone / ref_intervals
        need_frac = float(n_capas - 1) * pitch_frac
        max_use = layer_zone * 0.88
        if need_frac > max_use:
            pitch_frac = max(
                (float(CABEZAL_LAYER_PITCH_MIN_MM) / e_mm) * (layer_zone / ref_intervals),
                max_use / float(max(n_capas - 1, 1)),
            )

    if offset_trans_nom > 0 and spacing_mm > 0:
        wall_end_frac = fx_cov + (e_mm - offset_trans_nom) / (ref_intervals * spacing_mm) * layer_zone
    else:
        wall_end_frac = 1.0
    wall_end_frac = max(0.0, wall_end_frac)

    def _bar_fy_values(nb):
        if nb <= 1:
            return [inner_y0 + inner_h * 0.5]
        if draw_h_px is not None and bar_span_px is not None:
            dh = max(1.0, float(draw_h_px))
            span = max(1.0, float(bar_span_px))
            inner_h_px = inner_h * dh
            start_px = inner_y0 * dh + max(0.0, (inner_h_px - span) * 0.5)
            step = span / float(nb - 1)
            return [(start_px + float(bi) * step) / dh for bi in range(nb)]
        return [
            inner_y0 + inner_h * (float(bi) / float(nb - 1))
            for bi in range(nb)
        ]

    dots = []
    layer_bounds = []
    for i, ly in enumerate(layers):
        try:
            nb = int(ly.get(u"n_bars", CABEZAL_MIN_BARRAS_POR_CAPA))
        except Exception:
            nb = CABEZAL_MIN_BARRAS_POR_CAPA
        nb = max(
            CABEZAL_MIN_BARRAS_POR_CAPA,
            min(CABEZAL_MAX_BARRAS_POR_CAPA, nb),
        )
        cx = fx_cov + float(i) * pitch_frac
        half_col = pitch_frac * 0.45
        x_left = cx - half_col
        x_right = cx + half_col
        layer_bounds.append((x_left, x_right, i + 1))
        fy_values = _bar_fy_values(nb)
        for bi, fy in enumerate(fy_values):
            dots.append({
                u"layer": i + 1,
                u"layer_index": i,
                u"fx": cx,
                u"fy": fy,
            })

    stirrup_rect = None
    stirrup_segments = None
    tie_preview = None
    tie_previews = []
    stirrup_layer_indices = []
    tie_layer_indices = []
    if confinement_type and cabezal_confinement_scenario_applies(n_capas):
        stirrup_layer_indices, tie_layer_indices = cabezal_confinement_layout_spec(
            n_capas, confinement_type,
        )
        if cabezal_confinement_is_perimeter(confinement_type) and stirrup_layer_indices:
            stirrup_rect = cabezal_stirrup_preview_rect(
                dots, stirrup_layer_indices, pad_frac=stirrup_pad_frac,
            )
            if stirrup_rect:
                stirrup_segments = cabezal_stirrup_preview_segments(stirrup_rect)
        for li in tie_layer_indices:
            ly_t = layers[li] if len(layers) > li else {}
            try:
                bar_d = float(ly_t.get(u"bar_diam_mm") or bar_diam_mm)
            except Exception:
                bar_d = bar_diam_mm
            tp = cabezal_tie_preview_geometry(
                dots,
                layer_index=li,
                inner_y0=inner_y0,
                inner_h=inner_h,
                pitch_frac=pitch_frac,
                bar_diam_mm=bar_d,
                tie_diam_mm=float(
                    confinement_stirrup_diam_mm or CABEZAL_STIRRUP_DIAM_MM,
                ),
            )
            if tp:
                tie_previews.append(tp)
        if tie_previews:
            tie_preview = tie_previews[0]

    return {
        u"thickness_mm": e_mm,
        u"cover_mm": c_mm,
        u"cover_frac_x": fx_cov,
        u"cover_frac_y": fy_cov,
        u"inner_x0": inner_x0,
        u"inner_y0": inner_y0,
        u"inner_w": max(0.15, 1.0 - 2.0 * fx_cov),
        u"inner_h": inner_h,
        u"layer_pitch_frac": pitch_frac,
        u"wall_end_frac": wall_end_frac,
        u"dots": dots,
        u"layer_bounds": layer_bounds,
        u"n_capas": n_capas,
        u"stirrup_rect": stirrup_rect,
        u"stirrup_segments": stirrup_segments,
        u"stirrup_layer_indices": stirrup_layer_indices,
        u"tie_layer_indices": tie_layer_indices,
        u"tie_preview": tie_preview,
        u"tie_previews": tie_previews,
    }


def _bar_diameter_mm(bar_type):
    if bar_type is None:
        return 12.0
    try:
        return float(
            UnitUtils.ConvertFromInternalUnits(
                float(bar_type.BarModelDiameter), UnitTypeId.Millimeters,
            )
        )
    except Exception:
        try:
            return float(bar_type.BarModelDiameter) * 304.8
        except Exception:
            return 12.0


def _cabezal_tie_offset_mm(bar_type, tie_diam_mm, bar_diam_mm_fallback=None):
    """
    Distancia eje barra longitudinal → eje traba en profundidad del muro.

    Tangente exterior de la barra + mitad del diámetro de la traba (sin pad de
    estribo perimetral).
    """
    if bar_type is not None:
        bar_r_mm = _bar_diameter_mm(bar_type) * 0.5
    else:
        bar_r_mm = float(bar_diam_mm_fallback or 12.0) * 0.5
    tie_r_mm = float(tie_diam_mm or CABEZAL_STIRRUP_DIAM_MM) * 0.5
    return bar_r_mm + tie_r_mm


def _cabezal_tie_offset_ft(bar_type, tie_diam_mm):
    return _mm_to_internal(_cabezal_tie_offset_mm(bar_type, tie_diam_mm))


def _wall_longitudinal_at_extremo(wall, extremo):
    """
    Punto del extremo, vector hacia el interior del muro y normal exterior.

    Alineado con el script multicapa de referencia (LocationCurve + Orientation).
    """
    lc = location_curve_wall(wall) if location_curve_wall else None
    if lc is None:
        return None
    try:
        p0 = lc.GetEndPoint(0)
        p1 = lc.GetEndPoint(1)
        if extremo == CABEZAL_EXTREMO_FIN:
            pt_ext = p1
            v_long = p0.Subtract(p1)
        else:
            pt_ext = p0
            v_long = p1.Subtract(p0)
        vl = float(v_long.GetLength())
        if vl < 1e-12:
            return None
        vector_long = v_long.Normalize()
    except Exception:
        return None

    try:
        normal_muro = wall.Orientation.Normalize()
    except Exception:
        normal_muro = None
    if normal_muro is None or float(normal_muro.GetLength()) < 1e-12:
        try:
            normal_muro = vector_long.CrossProduct(XYZ.BasisZ).Normalize()
        except Exception:
            normal_muro = XYZ.BasisY

    try:
        espesor_ft = float(wall.Width)
    except Exception:
        espesor_ft = _mm_to_internal(200.0)
    if espesor_ft < 1e-9:
        espesor_ft = _mm_to_internal(200.0)

    return {
        u"pt_extremo": pt_ext,
        u"vector_longitudinal": vector_long,
        u"normal_muro": normal_muro,
        u"espesor_ft": espesor_ft,
    }


def _cabezal_capa_offsets_mm(layer_index, bar_type, conf_bar_type, layer_spacing_mm, cover_mm):
    """
    Offsets transversal (espesor) y longitudinal (profundidad en el muro) por capa.

    Transversal constante por capa; longitudinal crece con el índice de capa.
    """
    cover_mm = float(cover_mm if cover_mm is not None else CABEZAL_COVER_MM)
    spacing_mm = float(layer_spacing_mm if layer_spacing_mm is not None else CABEZAL_LAYER_PITCH_MM)
    conf_diam_mm = _bar_diameter_mm(conf_bar_type)
    offset_trans_mm = cover_mm + conf_diam_mm + 10.0
    offset_long_mm = offset_trans_mm + float(layer_index) * spacing_mm
    return offset_trans_mm, offset_long_mm


def _cabezal_layer_bar_trans_coords_ft(trans_origin_ft, distrib_ft, n_bars):
    """
    Cotas transversales (eje ``normal_muro``) alineadas con
    ``SetLayoutAsFixedNumber(..., barsOnNormalSide=False, includeFirst=True, includeLast=True)``:
    primera barra en el eje maestro; reparto hacia ``-normal``.
    """
    try:
        n_bars = int(n_bars)
    except Exception:
        n_bars = CABEZAL_MIN_BARRAS_POR_CAPA
    origin = float(trans_origin_ft)
    span = float(distrib_ft or 0.0)
    if n_bars <= 1:
        return [origin]
    if span < 1e-12:
        return [origin]
    step = span / float(n_bars - 1)
    return [origin - float(k) * step for k in range(n_bars)]


def cabezal_extremo_es_encuentro_l(ex_cfg):
    """True si el extremo usa geometría de encuentro L."""
    if _cab_enc_l is not None:
        return _cab_enc_l.cabezal_extremo_es_encuentro_l(ex_cfg)
    if not ex_cfg or not isinstance(ex_cfg, dict):
        return False
    return ex_cfg.get(u"encuentro_tipo") == u"L"


def _neighbor_wall_from_ex_cfg(doc, ex_cfg):
    if doc is None or not ex_cfg:
        return None
    try:
        vid = ex_cfg.get(u"vecino_wall_id")
        if vid is None:
            return None
        el = doc.GetElement(ElementId(int(vid)))
        if isinstance(el, Wall):
            return el
    except Exception:
        pass
    return None


def _cabezal_capa_line_endpoints(
    wall, extremo, layer_index, bar_type, conf_bar_type,
    layer_spacing_mm=None, cover_mm=None,
    doc=None, ex_cfg=None,
):
    """
    Calcula ``p_lo`` / ``p_hi`` de la línea vertical de una capa (eje maestro).

    Una capa = una ``Line`` + ``SetLayoutAsFixedNumber`` para las barras en espesor.

    La extensión vertical (Z) se obtiene del BoundingBox del muro para respetar
    la geometría real (Base Offset, Top Offset, restricciones de nivel).
    """
    if cabezal_extremo_es_encuentro_l(ex_cfg) and _cab_enc_l is not None:
        neighbor = _neighbor_wall_from_ex_cfg(doc, ex_cfg)
        if neighbor is not None:
            try:
                n_capas = int(ex_cfg.get(u"n_capas") or 2)
            except Exception:
                n_capas = 2
            return _cab_enc_l.cabezal_encuentro_l_capa_line_endpoints(
                doc,
                wall,
                extremo,
                layer_index,
                bar_type,
                conf_bar_type,
                neighbor,
                n_capas=n_capas,
                cover_mm=cover_mm,
            )

    geom = _wall_longitudinal_at_extremo(wall, extremo)
    if geom is None:
        return None, None, None, u"Sin LocationCurve u orientación válida."

    pt_ext = geom[u"pt_extremo"]
    vector_long = geom[u"vector_longitudinal"]
    normal_muro = geom[u"normal_muro"]
    espesor_ft = float(geom[u"espesor_ft"])
    dist_eje_cara = espesor_ft * 0.5

    offset_trans_mm, offset_long_mm = _cabezal_capa_offsets_mm(
        layer_index, bar_type, conf_bar_type, layer_spacing_mm, cover_mm,
    )
    offset_trans_ft = _mm_to_internal(offset_trans_mm)
    offset_long_ft = _mm_to_internal(offset_long_mm)
    desplazamiento_lateral = dist_eje_cara - offset_trans_ft

    try:
        inicio = (
            pt_ext
            + vector_long.Multiply(offset_long_ft)
            + normal_muro.Multiply(desplazamiento_lateral)
        )
    except Exception as ex_pt:
        return None, None, None, u"Offset capa: {0}".format(str(ex_pt))

    z_bot, z_top = _wall_z_bounds_ft(wall)
    p_lo = XYZ(float(inicio.X), float(inicio.Y), z_bot)
    p_hi = XYZ(float(inicio.X), float(inicio.Y), z_top)
    distrib_ft = max(espesor_ft - 2.0 * offset_trans_ft, 0.0)
    return p_lo, p_hi, distrib_ft, None


def _create_cabezal_capa_rebar(
    doc, wall, extremo, layer_index, n_bars, bar_type, conf_bar_type,
    layer_spacing_mm=None, cover_mm=None, ex_cfg=None,
):
    """
    Crea un Rebar de cabezal para una capa (línea vertical + layout en espesor).

    Retorna ``(rebar, mensaje_error)``.
    """
    if doc is None or wall is None or bar_type is None:
        return None, u"Doc, muro o RebarBarType no válido."
    try:
        n_bars = int(n_bars)
    except Exception:
        n_bars = CABEZAL_MIN_BARRAS_POR_CAPA
    n_bars = max(
        CABEZAL_MIN_BARRAS_POR_CAPA,
        min(CABEZAL_MAX_BARRAS_POR_CAPA, n_bars),
    )

    p_lo, p_hi, distrib_ft, err_geom = _cabezal_capa_line_endpoints(
        wall, extremo, layer_index, bar_type, conf_bar_type,
        layer_spacing_mm=layer_spacing_mm, cover_mm=cover_mm,
        doc=doc, ex_cfg=ex_cfg,
    )
    if err_geom:
        return None, err_geom

    geom = _wall_longitudinal_at_extremo(wall, extremo)
    normal_muro = geom[u"normal_muro"] if geom else XYZ.BasisY

    try:
        if p_lo.DistanceTo(p_hi) < doc.Application.ShortCurveTolerance:
            return None, u"Eje más corto que ShortCurveTolerance."
    except Exception:
        pass

    try:
        vertical_line = Line.CreateBound(p_lo, p_hi)
    except Exception as ex_ln:
        return None, u"Line.CreateBound: {0}".format(str(ex_ln))

    try:
        curves_list = List[Curve]()
        curves_list.Add(vertical_line)
    except Exception as ex_cl:
        return None, u"IList[Curve]: {0}".format(str(ex_cl))

    try:
        rebar = Rebar.CreateFromCurves(
            doc,
            RebarStyle.Standard,
            bar_type,
            None,
            None,
            wall,
            normal_muro,
            curves_list,
            RebarHookOrientation.Left,
            RebarHookOrientation.Left,
            True,
            True,
        )
    except Exception as ex_cf:
        try:
            return None, unicode(ex_cf)
        except Exception:
            return None, str(ex_cf)

    if rebar is None:
        return None, u"CreateFromCurves devolvió None."

    if n_bars > 1 and distrib_ft is not None:
        try:
            tol = float(doc.Application.ShortCurveTolerance)
        except Exception:
            tol = 1e-6
        if float(distrib_ft) >= tol:
            try:
                accessor = rebar.GetShapeDrivenAccessor()
                accessor.SetLayoutAsFixedNumber(
                    int(n_bars), float(distrib_ft), False, True, True,
                )
            except Exception as ex_lay:
                try:
                    return None, u"SetLayoutAsFixedNumber: {0}".format(unicode(ex_lay))
                except Exception:
                    return None, u"SetLayoutAsFixedNumber: {0}".format(str(ex_lay))

    return _stamp_armadura_arainco(rebar, layer_index=layer_index), None


def _element_to_bar_type(doc, element_id):
    if doc is None or element_id is None or element_id == ElementId.InvalidElementId:
        return None
    try:
        bt = doc.GetElement(element_id)
        if isinstance(bt, RebarBarType):
            return bt
    except Exception:
        pass
    return None


def _resolver_bar_type(doc, ex_cfg, fallback_id=None):
    bid = None
    try:
        bid = (ex_cfg or {}).get(u"bar_type_id")
    except Exception:
        bid = None
    if bid is None or bid == ElementId.InvalidElementId:
        bid = fallback_id
    return _element_to_bar_type(doc, bid)


def _resolver_bar_type_for_layer(
    doc, ex_cfg, layer_dict, fallback_id=None,
    segment_ctx=None, stack_index=0, layer_index=0,
):
    """``RebarBarType`` de cabezal para una capa (config del muro)."""
    bid = None
    try:
        bid = (layer_dict or {}).get(u"bar_type_id")
    except Exception:
        bid = None
    if bid is None or bid == ElementId.InvalidElementId:
        bid = (ex_cfg or {}).get(u"bar_type_id")
    if bid is None or bid == ElementId.InvalidElementId:
        bid = fallback_id
    return _element_to_bar_type(doc, bid)


def _resolver_bar_type_for_layer_tramo(
    doc, ex_cfg, layer_dict, tramo_index, fallback_id=None,
):
    """``RebarBarType`` de cabezal para capa ``layer_dict`` y tramo ``tramo_index``."""
    bid = None
    try:
        from bimtools_clr_collections import list_get_or_last

        tramo_ids = (layer_dict or {}).get(u"tramo_bar_type_ids") or []
        bid = list_get_or_last(tramo_ids, int(tramo_index), default=None)
    except Exception:
        try:
            tramo_ids = (layer_dict or {}).get(u"tramo_bar_type_ids") or []
            ti = int(tramo_index)
            if tramo_ids and len(tramo_ids) > 0:
                if ti < len(tramo_ids):
                    bid = tramo_ids[ti]
                else:
                    bid = tramo_ids[-1]
        except Exception:
            bid = None
    if bid is None or bid == ElementId.InvalidElementId:
        bid = (layer_dict or {}).get(u"bar_type_id")
    if bid is None or bid == ElementId.InvalidElementId:
        bid = (ex_cfg or {}).get(u"bar_type_id")
    if bid is None or bid == ElementId.InvalidElementId:
        bid = fallback_id
    return _element_to_bar_type(doc, bid)


def _resolve_tramo_n_bars_for_layer(layer_dict, n_tramos, fallback_n_bars=None):
    """Lista de barras por tramo vertical (longitud ``n_tramos``)."""
    n_tramos = max(1, int(n_tramos or 1))
    try:
        fb = int(fallback_n_bars if fallback_n_bars is not None else CABEZAL_MIN_BARRAS_POR_CAPA)
    except Exception:
        fb = CABEZAL_MIN_BARRAS_POR_CAPA
    return _normalize_tramo_n_bars(layer_dict, n_tramos, fb)


def _resolve_tramo_bar_types_for_layer(
    doc, ex_cfg, layer_dict, n_tramos, fallback_id=None,
):
    """Lista de ``RebarBarType`` por tramo vertical (longitud ``n_tramos``)."""
    n_tramos = max(1, int(n_tramos or 1))
    out = []
    for ti in range(n_tramos):
        bt = _resolver_bar_type_for_layer_tramo(
            doc, ex_cfg, layer_dict, ti, fallback_id,
        )
        out.append(bt)
    return out


def _resolver_conf_bar_type(doc, ex_cfg, bar_type, fallback_id=None):
    """Diámetro de confinamiento para offset transversal (recub + Ø conf + Ø/2 barra)."""
    bid = None
    try:
        bid = (ex_cfg or {}).get(u"conf_bar_type_id")
    except Exception:
        bid = None
    conf = _element_to_bar_type(doc, bid)
    if conf is not None:
        return conf
    if bar_type is not None:
        return bar_type
    return _element_to_bar_type(doc, fallback_id)


def _wall_thickness_mm_for_fusion(wall):
    """Espesor nominal del muro (mm) para agrupar fusión colineal."""
    if wall is None:
        return 200.0
    if obtener_espesor_muro_mm_approx is not None:
        try:
            th = obtener_espesor_muro_mm_approx(wall)
            if th is not None and float(th) > 0.1:
                return float(th)
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


def _thickness_mm_key_for_fusion(thickness_mm):
    """Clave estable de espesor (mm enteros) para buckets de fusión."""
    try:
        return int(round(float(thickness_mm)))
    except Exception:
        return 200


_XY_KEY_DECIMALS = 9
_EMBED_PROBE_XY_MARGIN_MM = 1.0
_EMBED_PROBE_MIN_HALF_SIDE_MM = 2.0
_TOL_VOL_INTERSECCION_FT3 = 1.0e-9
NO_COLLISION_RETRACT_BASE_MM = 25.0
FOUNDATION_PROBE_BASE_MM = 100.0
FOUNDATION_STRETCH_RESTA_MM = 50.0
TROCEO_EMPALME_POLICY_BASE = u"base"


# ═══════════════════════════════════════════════════════════════════════════
# Geometría de sólidos y prismas de prueba
# ═══════════════════════════════════════════════════════════════════════════

def _geometry_options():
    opts = Options()
    opts.ComputeReferences = False
    opts.DetailLevel = ViewDetailLevel.Fine
    return opts


def _iter_solids_element(elem, opts):
    if elem is None:
        return
    try:
        geo_elem = elem.get_Geometry(opts)
    except Exception:
        return
    if geo_elem is None:
        return
    from Autodesk.Revit.DB import Solid, GeometryInstance
    try:
        from bimtools_clr_collections import iterate_net_collection, safe_solid_volume
    except Exception:
        iterate_net_collection = None
        safe_solid_volume = None
    if iterate_net_collection is not None:
        geom_items = iterate_net_collection(geo_elem)
    else:
        geom_items = []
        try:
            for g in geo_elem:
                geom_items.append(g)
        except Exception:
            try:
                n = int(geo_elem.Count)
            except Exception:
                n = 0
            for i in range(n):
                try:
                    geom_items.append(geo_elem[i])
                except Exception:
                    try:
                        geom_items.append(geo_elem.get_Item(i))
                    except Exception:
                        pass

    def _vol_ok(solid):
        if solid is None:
            return False
        if safe_solid_volume is not None:
            v = safe_solid_volume(solid)
            return v is not None and v > 1e-12
        try:
            return float(solid.Volume) > 1e-12
        except Exception:
            return False

    for g in geom_items:
        if g is None:
            continue
        if isinstance(g, Solid) and _vol_ok(g):
            yield g
        elif isinstance(g, GeometryInstance):
            try:
                inst_geom = g.GetInstanceGeometry()
                if inst_geom is None:
                    continue
                if iterate_net_collection is not None:
                    inst_items = iterate_net_collection(inst_geom)
                else:
                    inst_items = []
                    try:
                        for sg in inst_geom:
                            inst_items.append(sg)
                    except Exception:
                        pass
                for sg in inst_items:
                    if isinstance(sg, Solid) and _vol_ok(sg):
                        yield sg
            except Exception:
                pass


def _solidos_intersectan_volumen(solid_a, solid_b, tol_volumen=_TOL_VOL_INTERSECCION_FT3):
    if solid_a is None or solid_b is None:
        return False
    try:
        from bimtools_clr_collections import safe_solid_volume
    except Exception:
        safe_solid_volume = None
    if safe_solid_volume is not None:
        va = safe_solid_volume(solid_a)
        vb = safe_solid_volume(solid_b)
        if va is None or vb is None or va <= 1e-12 or vb <= 1e-12:
            return False
    else:
        try:
            if float(solid_a.Volume) <= 1e-12 or float(solid_b.Volume) <= 1e-12:
                return False
        except Exception:
            return False
    try:
        inter = BooleanOperationsUtils.ExecuteBooleanOperation(
            solid_a, solid_b, BooleanOperationsType.Intersect,
        )
    except Exception:
        return False
    if inter is None:
        return False
    try:
        if safe_solid_volume is not None:
            vi = safe_solid_volume(inter)
            if vi is None:
                return False
            return vi > float(tol_volumen)
        return float(inter.Volume) > float(tol_volumen)
    except Exception:
        return False


def _build_vertical_square_prism_solid_inline(px, py, z_start_ft, half_side_ft, height_ft):
    hw = abs(float(half_side_ft))
    hgt = abs(float(height_ft))
    if hgt <= 1e-12:
        return None
    hs = XYZ(float(px), float(py), float(z_start_ft))
    p1 = XYZ(hs.X - hw, hs.Y - hw, hs.Z)
    p2 = XYZ(hs.X + hw, hs.Y - hw, hs.Z)
    p3 = XYZ(hs.X + hw, hs.Y + hw, hs.Z)
    p4 = XYZ(hs.X - hw, hs.Y + hw, hs.Z)
    try:
        loop = CurveLoop.Create([
            Line.CreateBound(p1, p2),
            Line.CreateBound(p2, p3),
            Line.CreateBound(p3, p4),
            Line.CreateBound(p4, p1),
        ])
    except Exception:
        return None
    try:
        sol = GeometryCreationUtilities.CreateExtrusionGeometry(
            [loop], XYZ.BasisZ, hgt,
        )
    except Exception:
        return None
    if sol is None:
        return None
    try:
        if float(sol.Volume) < 1e-15:
            return None
    except Exception:
        return None
    return sol


def _build_vertical_square_prism_solid(px, py, z_start_ft, half_side_ft, height_ft):
    try:
        from bimtools_clr_collections import create_vertical_square_prism_solid

        sol = create_vertical_square_prism_solid(
            px, py, z_start_ft, half_side_ft, height_ft,
        )
        if sol is not None:
            return sol
    except Exception:
        pass
    return _build_vertical_square_prism_solid_inline(
        px, py, z_start_ft, half_side_ft, height_ft,
    )


def _build_vertical_prism_downward(px, py, z_top_ft, half_side_ft, height_ft):
    hgt = abs(float(height_ft))
    if hgt <= 1e-12:
        return None
    z_base = float(z_top_ft) - hgt
    return _build_vertical_square_prism_solid(
        float(px), float(py), z_base, abs(float(half_side_ft)), hgt,
    )


# ═══════════════════════════════════════════════════════════════════════════
# Colisión embed +Z / -Z con sólidos de muros
# ═══════════════════════════════════════════════════════════════════════════

def _embed_collides_wall_solids_upward(
    doc, px, py, z_top_ft, dz_embed_ft, bar_nominal_mm,
    wall_obstacles, host_wall_id, geom_opts,
):
    try:
        dz_e = abs(float(dz_embed_ft))
        if doc is None or dz_e <= 1e-12:
            return False
        half_w_mm = float(bar_nominal_mm) / 2.0 + _EMBED_PROBE_XY_MARGIN_MM
        half_w_mm = max(half_w_mm, _EMBED_PROBE_MIN_HALF_SIDE_MM)
        half_w_ft = _mm_to_internal(half_w_mm)
        probe = _build_vertical_square_prism_solid(
            float(px), float(py), float(z_top_ft), half_w_ft, dz_e,
        )
        if probe is None:
            return False
        try:
            host_id = element_id_to_int(host_wall_id)
        except Exception:
            host_id = None
        for wall in wall_obstacles or []:
            if wall is None:
                continue
            try:
                wid = wall_id_int(wall)
            except Exception:
                continue
            if host_id is not None and wid == host_id:
                continue
            for sd in _iter_solids_element(wall, geom_opts):
                if _solidos_intersectan_volumen(probe, sd):
                    return True
        return False
    except Exception:
        return False


def _embed_collides_wall_solids_downward(
    doc, px, py, z_bot_ft, dz_embed_ft, bar_nominal_mm,
    wall_obstacles, host_wall_id, geom_opts,
):
    try:
        dz_e = abs(float(dz_embed_ft))
        if doc is None or dz_e <= 1e-12:
            return False
        half_w_mm = float(bar_nominal_mm) / 2.0 + _EMBED_PROBE_XY_MARGIN_MM
        half_w_mm = max(half_w_mm, _EMBED_PROBE_MIN_HALF_SIDE_MM)
        half_w_ft = _mm_to_internal(half_w_mm)
        probe = _build_vertical_prism_downward(
            float(px), float(py), float(z_bot_ft), half_w_ft, dz_e,
        )
        if probe is None:
            return False
        try:
            host_id = element_id_to_int(host_wall_id)
        except Exception:
            host_id = None
        for wall in wall_obstacles or []:
            if wall is None:
                continue
            try:
                wid = wall_id_int(wall)
            except Exception:
                continue
            if host_id is not None and wid == host_id:
                continue
            for sd in _iter_solids_element(wall, geom_opts):
                if _solidos_intersectan_volumen(probe, sd):
                    return True
        return False
    except Exception:
        return False


# ═══════════════════════════════════════════════════════════════════════════
# Fundación unida al muro
# ═══════════════════════════════════════════════════════════════════════════

def _es_fundacion_estructural(element):
    if element is None:
        return False
    try:
        cat = element.Category
        if cat is None:
            return False
        return element_id_to_int(cat.Id) == int(BuiltInCategory.OST_StructuralFoundation)
    except Exception:
        return False


def _ids_elementos_unidos(doc, element):
    try:
        from bimtools_joined_geometry import get_joined_element_ids

        return list(get_joined_element_ids(doc, element) or [])
    except Exception:
        pass
    if doc is None or element is None:
        return []
    out = []
    raw = None
    try:
        raw = JoinGeometryUtils.GetJoinedElements(doc, element)
    except Exception:
        raw = None
    if raw is None:
        try:
            raw = JoinGeometryUtils.GetJoinedElements(doc, element.Id)
        except Exception:
            raw = None
    if raw is None:
        return out
    try:
        n = int(raw.Count)
    except Exception:
        n = 0
    for i in range(n):
        try:
            eid = raw[i]
            if eid is not None and eid != ElementId.InvalidElementId:
                out.append(eid)
        except Exception:
            pass
    return out


def _resolve_host_wall(bx, by, z_mid, walls, geom_opts, fallback_wall):
    """
    Genera un prisma pequeño en el punto medio de la barra y evalúa
    intersección con los sólidos de cada muro. El primero que intersecte
    se asigna como host. Fallback: BoundingBox Z, luego wall original.
    """
    probe = _build_vertical_square_prism_solid(
        float(bx), float(by), float(z_mid) - 0.01, 0.01, 0.02,
    )
    if probe is not None:
        for wall in walls:
            if wall is None:
                continue
            for sd in _iter_solids_element(wall, geom_opts):
                if _solidos_intersectan_volumen(probe, sd):
                    return wall
    for wall in walls:
        if wall is None:
            continue
        try:
            bb = wall.get_BoundingBox(None)
            if bb is None:
                continue
            if float(bb.Min.Z) <= z_mid <= float(bb.Max.Z):
                return wall
        except Exception:
            continue
    return fallback_wall


def _fundaciones_unidas_muro(doc, wall):
    if doc is None or wall is None:
        return []
    out = []
    for eid in _ids_elementos_unidos(doc, wall):
        el = doc.GetElement(eid)
        if _es_fundacion_estructural(el):
            out.append(el)
    return out


def _cabezal_stirrup_foundation_drop_ft(doc, wall):
    """
    Extensión fija del array de confinamiento hacia la base cuando el muro tiene
    fundación estructural unida (estribos perimetrales y trabas capa [1]).
    """
    if not _fundaciones_unidas_muro(doc, wall):
        return 0.0
    return _mm_to_internal(float(CABEZAL_STIRRUP_FOUNDATION_DROP_MM))


def _altura_bbox_elemento_mm(elem):
    if elem is None:
        return None
    try:
        bb = elem.get_BoundingBox(None)
        if bb is None:
            return None
        h_ft = float(bb.Max.Z - bb.Min.Z)
        if h_ft <= 1e-12:
            return None
        return float(UnitUtils.ConvertFromInternalUnits(h_ft, UnitTypeId.Millimeters))
    except Exception:
        return None


def _probe_colision_fundacion_desde_punto(
    xyz_ref, probe_mm, bar_nominal_mm, foundations, geom_opts,
):
    if xyz_ref is None or not foundations:
        return False, None
    dz_ft = _mm_to_internal(max(0.1, float(probe_mm)))
    half_w_mm = float(bar_nominal_mm) / 2.0 + _EMBED_PROBE_XY_MARGIN_MM
    half_w_mm = max(half_w_mm, _EMBED_PROBE_MIN_HALF_SIDE_MM)
    half_w_ft = _mm_to_internal(half_w_mm)
    probe = _build_vertical_prism_downward(
        float(xyz_ref.X), float(xyz_ref.Y), float(xyz_ref.Z), half_w_ft, dz_ft,
    )
    if probe is None:
        return False, None
    h_max = 0.0
    any_hit = False
    for fund in foundations:
        if fund is None:
            continue
        h_mm = _altura_bbox_elemento_mm(fund)
        for sd in _iter_solids_element(fund, geom_opts):
            if _solidos_intersectan_volumen(probe, sd):
                any_hit = True
                if h_mm is not None and float(h_mm) > h_max:
                    h_max = float(h_mm)
                break
    if not any_hit:
        return False, None
    return True, h_max if h_max > 0.1 else None


def _foundation_stretch_mm(bar_nominal_mm, h_fund_mm):
    base_mm = float(FOUNDATION_PROBE_BASE_MM)
    extra = float(h_fund_mm) - base_mm - float(FOUNDATION_STRETCH_RESTA_MM) - float(bar_nominal_mm) / 2.0
    return max(base_mm, base_mm + max(0.0, extra))


# ═══════════════════════════════════════════════════════════════════════════
# Empotramiento / traslape / pata por diámetro
# ═══════════════════════════════════════════════════════════════════════════

def _empotramiento_tabla_mm(d_mm, concrete_grade=None):
    if traslape_mm_from_nominal_diameter_mm is None:
        return None
    try:
        d = float(d_mm)
    except Exception:
        return None
    if d <= 0.0:
        return None
    try:
        L = traslape_mm_from_nominal_diameter_mm(d, concrete_grade)
    except Exception:
        return None
    if L is None or float(L) < 0.0:
        return None
    return max(0.0, float(L))


def _empotramiento_max_mm_at_z_joint(
    doc, walls, cabezal_por_muro_id, segment_ctx, extremo,
    layer_index, z_joint_ft, curve_tol, fallback_bar_type=None,
):
    """
    Mayor L(Ø) de tabla en una junta vertical: barras por encima y por debajo de ``z_joint_ft``.
    """
    embed_mm = 0.0
    tt = max(float(curve_tol or 1e-6), 1e-9)
    zj = float(z_joint_ft)
    for z_probe in (zj - tt, zj + tt):
        stack_j = _stack_index_for_z(walls, z_probe, tt)
        bt_j = _bar_type_for_cabezal_stack(
            doc, walls, cabezal_por_muro_id, segment_ctx, extremo,
            stack_j, layer_index, fallback_bar_type,
        )
        if bt_j is None:
            bt_j = fallback_bar_type
        if bt_j is None:
            continue
        emb = _empotramiento_tabla_mm(_bar_diameter_mm(bt_j)) or 0.0
        if emb > embed_mm:
            embed_mm = emb
    if embed_mm < 1e-6:
        return 860.0
    return embed_mm


def _retract_mm_sin_colision(d_mm):
    return float(NO_COLLISION_RETRACT_BASE_MM)


def _pata_l_mm_desde_diametro(d_mm, concrete_grade=None):
    if hook_length_mm_from_nominal_diameter_mm is None:
        return None
    try:
        return float(hook_length_mm_from_nominal_diameter_mm(d_mm, concrete_grade))
    except Exception:
        return None


def _pata_l_eje_sketch_mm_desde_diametro(d_mm, concrete_grade=None):
    """
    Longitud de eje para patas L de cabezal (CreateFromCurves).

    Revit modela la pata desde el eje de la barra; restamos Ø/2 al valor de tabla
    para que el largo modelado coincida con la tabla BIMTools.
    """
    tabla_mm = _pata_l_mm_desde_diametro(d_mm, concrete_grade)
    if tabla_mm is None:
        return None
    if pata_eje_curve_loop_mm_desde_tabla_mm is not None:
        try:
            Leje = pata_eje_curve_loop_mm_desde_tabla_mm(tabla_mm, d_mm)
            if Leje is not None:
                return float(Leje)
        except Exception:
            pass
    try:
        d = float(int(round(float(d_mm))))
    except Exception:
        d = 0.0
    if d > 1e-6:
        return max(40.0, float(tabla_mm) - 0.5 * d)
    return float(tabla_mm)


# ═══════════════════════════════════════════════════════════════════════════
# Planos de corte (troceo) desde muros de referencia
# ═══════════════════════════════════════════════════════════════════════════

def _wall_base_cut_plane(wall, tol):
    """Plano horizontal en la cota Z inferior del BoundingBox del muro."""
    z_bot, z_top = _wall_z_bounds_ft(wall)
    try:
        return Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0.0, 0.0, z_bot))
    except Exception:
        return None


def build_wall_a_cut_planes(ref_walls, tol):
    """Planos A: Z-base de cada muro de referencia."""
    planes = []
    policies = []
    if not ref_walls:
        return planes, policies
    used_z = []
    z_dup_ft = max(tol * 500.0, _mm_to_internal(2.0))
    for wall in ref_walls:
        pl = _wall_base_cut_plane(wall, tol)
        if pl is None:
            continue
        z_val = float(pl.Origin.Z)
        dup = False
        for uz in used_z:
            if abs(z_val - uz) < z_dup_ft:
                dup = True
                break
        if dup:
            continue
        used_z.append(z_val)
        planes.append(pl)
        policies.append(TROCEO_EMPALME_POLICY_BASE)
    return planes, policies


def _sort_ref_walls_by_z_base(ref_walls):
    """Muros de troceo ordenados por Z base ascendente (T1 abajo)."""
    if not ref_walls:
        return []
    keyed = []
    for wall in ref_walls:
        if wall is None:
            continue
        try:
            z_bot, _ = _wall_z_bounds_ft(wall)
            keyed.append((float(z_bot), wall))
        except Exception:
            keyed.append((0.0, wall))
    keyed.sort()
    return [w for _, w in keyed]


def build_wall_b_cut_planes(ref_walls, tol, embed_mm):
    """Planos B: mismos orígenes que A pero desplazados +L(Ø) en Z."""
    embeds = [embed_mm] * len(ref_walls or [])
    return build_wall_b_cut_planes_embeds(ref_walls, tol, embeds)


def build_wall_b_cut_planes_embeds(ref_walls, tol, embed_mm_list):
    """
    Planos B con L(Ø) por muro de referencia (orden Z base ascendente).
    ``embed_mm_list[i]`` corresponde al i-ésimo muro ordenado.
    """
    planes = []
    policies = []
    sorted_walls = _sort_ref_walls_by_z_base(ref_walls)
    if not sorted_walls:
        return planes, policies
    z_dup_ft = max(float(tol) * 500.0, _mm_to_internal(2.0))
    used_z = []
    for i, wall in enumerate(sorted_walls):
        pl = _wall_base_cut_plane(wall, tol)
        if pl is None:
            continue
        try:
            n_emb = 0
            if embed_mm_list is not None:
                try:
                    n_emb = int(len(embed_mm_list))
                except Exception:
                    try:
                        n_emb = int(embed_mm_list.Count)
                    except Exception:
                        n_emb = 0
            if n_emb < 1:
                embed_mm = 860.0
            elif i < n_emb:
                embed_mm = float(embed_mm_list[i])
            else:
                embed_mm = float(embed_mm_list[n_emb - 1])
        except Exception:
            embed_mm = 860.0
        off_ft = float(embed_mm or 860.0) / 304.8
        z_base = float(pl.Origin.Z) + off_ft
        dup = False
        for uz in used_z:
            if abs(z_base - uz) < z_dup_ft:
                dup = True
                break
        if dup:
            continue
        used_z.append(z_base)
        try:
            pl_b = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0.0, 0.0, z_base))
        except Exception:
            continue
        planes.append(pl_b)
        policies.append(TROCEO_EMPALME_POLICY_BASE)
    return planes, policies


def _z_cut_vertical_bar_plane(bx, by, z_lo, z_hi, plane, tol_ft):
    """Intersección Z de un hilo vertical en (bx,by) con un Plane."""
    if plane is None:
        return None
    try:
        n = plane.Normal
        ox = float(plane.Origin.X)
        oy = float(plane.Origin.Y)
        oz = float(plane.Origin.Z)
        nz = float(n.Z)
    except Exception:
        return None
    tt = abs(float(tol_ft))
    if tt < 1e-12:
        tt = 1e-12
    if abs(nz) < tt * 50.0:
        return None
    nx = float(n.X)
    ny = float(n.Y)
    z_int = oz - (nx * (float(bx) - ox) + ny * (float(by) - oy)) / nz
    if z_int <= float(z_lo) + tt or z_int >= float(z_hi) - tt:
        return None
    return z_int


def _split_z_span_by_planes(bx, by, z_lo, z_hi, planes, plane_policies, tol_ft):
    """
    Divide el intervalo ``[z_lo, z_hi]`` con los planos dados.
    Retorna ``(segments, joint_policies)`` donde segments = ``[(z_start, dz), ...]``.
    """
    z_lo = float(z_lo)
    z_hi = float(z_hi)
    tt = abs(float(tol_ft))
    if tt < 1e-12:
        tt = 1e-12
    merge_eps = max(tt * 4.0, tt)
    if z_hi <= z_lo + tt:
        return [], []
    if not planes:
        return [(z_lo, z_hi - z_lo)], []
    cuts = [(z_lo, None)]
    n_pl = min(len(planes), len(plane_policies or []))
    for i in range(n_pl):
        zi = _z_cut_vertical_bar_plane(bx, by, z_lo, z_hi, planes[i], tt)
        if zi is not None:
            cuts.append((float(zi), plane_policies[i]))
    cuts.append((z_hi, None))
    cuts.sort()
    merged_z = []
    merged_pol = []
    for zc, pc in cuts:
        if not merged_z or zc > merged_z[-1] + merge_eps:
            merged_z.append(zc)
            merged_pol.append(pc)
        elif pc is not None and merged_pol:
            merged_pol[-1] = pc
    if len(merged_z) < 2:
        return [(z_lo, z_hi - z_lo)], []
    segments = []
    joint_policies = []
    for i in range(len(merged_z) - 1):
        a = merged_z[i]
        b = merged_z[i + 1]
        dz = b - a
        if dz > tt * 4.0:
            segments.append((a, dz))
            jp = merged_pol[i + 1]
            if jp is None:
                jp = TROCEO_EMPALME_POLICY_BASE
            joint_policies.append(jp)
    if not segments:
        return [(z_lo, z_hi - z_lo)], []
    return segments, joint_policies


def _cabezal_as_python_list(collection):
    """Lista Python segura (IList .NET vacío puede ser truthy en CPython 3)."""
    try:
        from bimtools_clr_collections import as_python_list

        return as_python_list(collection)
    except Exception:
        pass
    if collection is None:
        return []
    if isinstance(collection, list):
        return collection
    try:
        if isinstance(collection, (set, tuple)):
            return list(collection)
    except Exception:
        pass
    try:
        return list(collection)
    except Exception:
        return []


def _cabezal_ex_cfg_layers(ex_cfg):
    return _cabezal_as_python_list((ex_cfg or {}).get(u"layers"))


def _cabezal_contrib_wids(fj):
    return _cabezal_as_python_list((fj or {}).get(u"contrib_wids"))


def _cabezal_rethrow_phase(phase, ex):
    raise Exception(u"{0}: {1}".format(phase, ex))


def _troceo_planificar_seg_jobs_from_fused_line(
    doc,
    walls,
    cabezal_por_muro_id,
    segment_ctx,
    fj,
    fj_idx,
    fund_stretch_ft_by_wid,
    legacy_ref_walls,
    curve_tol,
    geom_opts,
):
    """
    Troceo + empalme de una línea fusionada → lista de jobs de segmento.
    Errores incluyen fase (planos, split, empalme, segmentos).
    """
    bx = float(fj[u"p_lo"].X)
    by = float(fj[u"p_lo"].Y)
    z_lo = float(fj[u"p_lo"].Z)
    z_hi = float(fj[u"p_hi"].Z)
    if z_hi <= z_lo + curve_tol:
        return []

    extremo = fj.get(u"extremo") or CABEZAL_EXTREMO_INICIO
    layer_index = int(fj.get(u"layer_index", 0))

    try:
        d_mm = _bar_diameter_mm(fj[u"bar_type"])
    except Exception as ex:
        _cabezal_rethrow_phase(u"diametro barra", ex)

    ref_walls_job = _cabezal_as_python_list(fj.get(u"troceo_walls"))
    if legacy_ref_walls:
        seen_ids = set()
        merged = []
        for w in ref_walls_job + _cabezal_as_python_list(legacy_ref_walls):
            if w is None:
                continue
            try:
                iv = wall_id_int(w)
            except Exception:
                continue
            if iv in seen_ids:
                continue
            seen_ids.add(iv)
            merged.append(w)
        ref_walls_job = merged

    use_a = (int(layer_index) % 2 == 0)
    cut_planes = []
    cut_policies = []
    if ref_walls_job:
        try:
            if use_a:
                cut_planes, cut_policies = build_wall_a_cut_planes(
                    ref_walls_job, curve_tol,
                )
            else:
                sorted_ref = _sort_ref_walls_by_z_base(ref_walls_job)
                embeds_b = []
                for ri in range(len(sorted_ref)):
                    stack_ri = _wall_stack_index(walls, sorted_ref[ri])
                    bt_tr = _bar_type_for_cabezal_stack(
                        doc, walls, cabezal_por_muro_id, segment_ctx, extremo,
                        stack_ri, layer_index, fj[u"bar_type"],
                    )
                    if bt_tr is None:
                        bt_tr = fj[u"bar_type"]
                    d_tr = _bar_diameter_mm(bt_tr)
                    embeds_b.append(_empotramiento_tabla_mm(d_tr) or 860.0)
                cut_planes, cut_policies = build_wall_b_cut_planes_embeds(
                    ref_walls_job, curve_tol, embeds_b,
                )
        except Exception as ex:
            _cabezal_rethrow_phase(u"planos corte", ex)

    segments = [(z_lo, z_hi - z_lo)]
    if cut_planes:
        try:
            segs_try, _jp = _split_z_span_by_planes(
                bx, by, z_lo, z_hi, cut_planes, cut_policies, curve_tol,
            )
            if segs_try:
                segments = segs_try
        except Exception as ex:
            _cabezal_rethrow_phase(u"split Z", ex)

    n_seg = len(segments)
    seg_list = [[float(s[0]), float(s[1])] for s in segments]

    if n_seg > 1:
        try:
            for si in range(n_seg - 1):
                z_joint = seg_list[si][0] + seg_list[si][1]
                embed_joint = 0.0
                for z_probe in (z_joint - curve_tol, z_joint + curve_tol):
                    stack_j = _stack_index_for_z(walls, z_probe, curve_tol)
                    bt_j = _bar_type_for_cabezal_stack(
                        doc, walls, cabezal_por_muro_id, segment_ctx, extremo,
                        stack_j, layer_index, fj[u"bar_type"],
                    )
                    if bt_j is None:
                        bt_j = fj[u"bar_type"]
                    d_j = _bar_diameter_mm(bt_j)
                    emb = _empotramiento_tabla_mm(d_j) or 0.0
                    if emb > embed_joint:
                        embed_joint = emb
                dz_joint = _mm_to_internal(embed_joint) if embed_joint > 0.1 else 0.0
                if dz_joint > 1e-12:
                    seg_list[si][1] += float(dz_joint)
        except Exception as ex:
            _cabezal_rethrow_phase(u"empalme juntas", ex)

    has_fund = any(
        w in fund_stretch_ft_by_wid for w in _cabezal_contrib_wids(fj)
    )
    reverted_bot = False
    reverted_top = fj.get(u"_reverted_top", False)
    if not has_fund and seg_list:
        try:
            z_bot = seg_list[0][0]
            stack_bot = _stack_index_for_z(walls, z_bot + curve_tol, curve_tol)
            bt_bot = _bar_type_for_cabezal_stack(
                doc, walls, cabezal_por_muro_id, segment_ctx, extremo,
                stack_bot, layer_index, fj[u"bar_type"],
            )
            if bt_bot is None:
                bt_bot = fj[u"bar_type"]
            d_bot = _bar_diameter_mm(bt_bot)
            embed_bot = _empotramiento_tabla_mm(d_bot) or 0.0
            dz_embed_bot = _mm_to_internal(embed_bot) if embed_bot > 0.1 else 0.0
            retract_ft = _mm_to_internal(_retract_mm_sin_colision(d_bot))
            s0_z = seg_list[0][0]
            s0_dz = seg_list[0][1]
            if dz_embed_bot > 1e-12:
                collides = _embed_collides_wall_solids_downward(
                    doc, bx, by, s0_z, dz_embed_bot, d_bot,
                    walls, fj[u"wall"].Id, geom_opts,
                )
                if collides:
                    seg_list[0][0] = s0_z - float(dz_embed_bot)
                    seg_list[0][1] = s0_dz + float(dz_embed_bot)
                else:
                    seg_list[0][0] = s0_z + float(retract_ft)
                    seg_list[0][1] = max(s0_dz - float(retract_ft), curve_tol * 4.0)
                    reverted_bot = True
        except Exception as ex:
            _cabezal_rethrow_phase(u"embed base", ex)

    out_jobs = []
    try:
        for si, row in enumerate(seg_list):
            z_mid = float(row[0]) + float(row[1]) * 0.5
            stack_si = _stack_index_for_z(walls, z_mid, curve_tol)
            bt_seg = _bar_type_for_cabezal_stack(
                doc, walls, cabezal_por_muro_id, segment_ctx, extremo,
                stack_si, layer_index, fj[u"bar_type"],
            )
            if bt_seg is None:
                bt_seg = fj[u"bar_type"]
            n_bars_seg = _n_bars_for_stack_index(
                cabezal_por_muro_id, walls, extremo, stack_si, layer_index,
            )
            d_seg = _bar_diameter_mm(bt_seg)
            pata_eje_mm = _pata_l_eje_sketch_mm_desde_diametro(d_seg) or 0.0
            pata_ft_seg = _mm_to_internal(pata_eje_mm) if pata_eje_mm > 0.1 else 0.0
            want_bot = (si == 0 and (has_fund or reverted_bot) and pata_ft_seg > 1e-12)
            want_top = (si == n_seg - 1 and reverted_top and pata_ft_seg > 1e-12)
            out_jobs.append({
                u"bx": bx,
                u"by": by,
                u"zs": row[0],
                u"span_seg": row[1],
                u"wall": fj[u"wall"],
                u"wid": fj[u"wid"],
                u"bar_type": bt_seg,
                u"normal_muro": fj[u"normal_muro"],
                u"vec_long": fj[u"vec_long"],
                u"distrib_ft": fj[u"distrib_ft"],
                u"n_bars": int(n_bars_seg),
                u"contrib_wids": fj.get(u"contrib_wids"),
                u"want_bot_pata": want_bot,
                u"want_top_pata": want_top,
                u"pata_ft": pata_ft_seg,
                u"layer_index": layer_index,
                u"extremo": extremo,
                u"fusion_key": u"fj_{0}".format(fj_idx),
                u"seg_index": int(si),
                u"n_segments": int(n_seg),
            })
    except Exception as ex:
        _cabezal_rethrow_phase(u"segmentos", ex)

    return out_jobs, n_seg, ref_walls_job


# ═══════════════════════════════════════════════════════════════════════════
# Fusión colineal
# ═══════════════════════════════════════════════════════════════════════════

def _fuse_colinear_cabezal_lines(line_jobs):
    """
    Fusión colineal: agrupa por (XY, extremo, capa, espesor muro mm) y extiende
    el intervalo Z al rango combinado ``[min(Z), max(Z)]``.
    Solo fusiona líneas de muros con el mismo espesor nominal.
    Acumula muros con ``troceo_por_muro`` para planos de corte.
    """
    if not line_jobs:
        return []
    buckets = {}
    key_order = []
    for job in line_jobs:
        p_lo = job[u"p_lo"]
        p_hi = job[u"p_hi"]
        xf = round(float(p_lo.X), _XY_KEY_DECIMALS)
        yf = round(float(p_lo.Y), _XY_KEY_DECIMALS)
        zb = float(p_lo.Z)
        zt = float(p_hi.Z)
        extremo = job.get(u"extremo") or CABEZAL_EXTREMO_INICIO
        layer_index = int(job.get(u"layer_index", 0))
        th_key = _thickness_mm_key_for_fusion(
            job.get(u"thickness_mm", _wall_thickness_mm_for_fusion(job.get(u"wall"))),
        )
        k = (xf, yf, extremo, layer_index, th_key)
        if k not in buckets:
            troceo_walls = {}
            if job.get(u"troceo_por_muro"):
                troceo_walls[job[u"wid"]] = job[u"wall"]
            buckets[k] = {
                u"rx": float(p_lo.X), u"ry": float(p_lo.Y),
                u"z0": zb, u"z1": zt,
                u"distrib_ft": job[u"distrib_ft"],
                u"n_bars": job[u"n_bars"],
                u"bar_type": job[u"bar_type"],
                u"conf_type": job[u"conf_type"],
                u"normal_muro": job[u"normal_muro"],
                u"vec_long": job[u"vec_long"],
                u"wall": job[u"wall"],
                u"wid": job[u"wid"],
                u"layer_index": layer_index,
                u"extremo": extremo,
                u"contrib_wids": {job[u"wid"]},
                u"troceo_walls": troceo_walls,
            }
            key_order.append(k)
        else:
            bk = buckets[k]
            bk[u"z0"] = min(float(bk[u"z0"]), zb)
            bk[u"z1"] = max(float(bk[u"z1"]), zt)
            bk[u"contrib_wids"].add(job[u"wid"])
            if job.get(u"troceo_por_muro"):
                bk[u"troceo_walls"][job[u"wid"]] = job[u"wall"]
    fused = []
    for k in key_order:
        bk = buckets[k]
        span = float(bk[u"z1"]) - float(bk[u"z0"])
        if span < 1e-6:
            continue
        tw = bk.get(u"troceo_walls") or {}
        fused.append({
            u"p_lo": XYZ(float(bk[u"rx"]), float(bk[u"ry"]), float(bk[u"z0"])),
            u"p_hi": XYZ(float(bk[u"rx"]), float(bk[u"ry"]), float(bk[u"z1"])),
            u"distrib_ft": bk[u"distrib_ft"],
            u"n_bars": bk[u"n_bars"],
            u"bar_type": bk[u"bar_type"],
            u"conf_type": bk[u"conf_type"],
            u"normal_muro": bk[u"normal_muro"],
            u"vec_long": bk[u"vec_long"],
            u"wall": bk[u"wall"],
            u"wid": bk[u"wid"],
            u"layer_index": bk[u"layer_index"],
            u"extremo": bk.get(u"extremo"),
            u"contrib_wids": bk[u"contrib_wids"],
            u"troceo_walls": list(tw.values()),
        })
    return fused


# ═══════════════════════════════════════════════════════════════════════════
# Creación de rebar con patas opcionales
# ═══════════════════════════════════════════════════════════════════════════

def _pata_direction_for_wall(vec_long):
    """
    Dirección horizontal de la pata-L para cabezal: a lo largo del eje
    longitudinal del muro (hacia el interior del muro desde el extremo).
    ``vec_long`` ya apunta desde el extremo hacia el interior del muro.
    """
    lx = float(vec_long.X)
    ly = float(vec_long.Y)
    hl = math.hypot(lx, ly)
    if hl < 1e-9:
        return None, None
    return lx / hl, ly / hl


def _create_cabezal_rebar_with_optional_patas(
    doc, bx, by, zs, span_seg, host_wall, bar_type, normal_muro,
    vec_long, distrib_ft, n_bars, want_bot_pata, want_top_pata, pata_len_ft,
    layer_index=0,
):
    """
    Crea un Rebar de cabezal: tronco vertical con pata inferior y/o superior
    opcionales en la misma cadena de curvas.
    """
    bx = float(bx)
    by = float(by)
    zs = float(zs)
    span = float(span_seg)
    try:
        tol = abs(float(doc.Application.ShortCurveTolerance))
    except Exception:
        tol = 1e-6
    if span <= tol:
        return None, False, False

    z_top = zs + span
    foot = XYZ(bx, by, zs)
    head = XYZ(bx, by, z_top)

    ux, uy = _pata_direction_for_wall(vec_long)
    lf = float(pata_len_ft) if pata_len_ft else 0.0

    bot_leg = None
    top_leg = None
    if want_bot_pata and ux is not None and lf > tol:
        pa_b = XYZ(bx, by, zs)
        pb_b = XYZ(bx + ux * lf, by + uy * lf, zs)
        bot_leg = (pa_b, pb_b)
    if want_top_pata and ux is not None and lf > tol:
        pa_t = XYZ(bx, by, z_top)
        pb_t = XYZ(bx + ux * lf, by + uy * lf, z_top)
        top_leg = (pa_t, pb_t)

    curves = []
    if bot_leg and top_leg:
        pa_b, pb_b = bot_leg
        pa_t, pb_t = top_leg
        try:
            curves.append(Line.CreateBound(pb_b, pa_b))
            curves.append(Line.CreateBound(pa_b, head))
            curves.append(Line.CreateBound(pa_t, pb_t))
        except Exception:
            return None, False, False
    elif bot_leg:
        pa_b, pb_b = bot_leg
        try:
            curves.append(Line.CreateBound(head, pa_b))
            curves.append(Line.CreateBound(pa_b, pb_b))
        except Exception:
            bot_leg = None
            curves = []
    elif top_leg:
        pa_t, pb_t = top_leg
        try:
            curves.append(Line.CreateBound(foot, head))
            curves.append(Line.CreateBound(pa_t, pb_t))
        except Exception:
            top_leg = None
            curves = []

    if not curves:
        try:
            curves.append(Line.CreateBound(foot, head))
        except Exception:
            return None, False, False

    nvec = normal_muro

    rb = None
    try:
        cl = List[Curve]()
        for c in curves:
            cl.Add(c)
    except Exception:
        return None, False, False

    for use_ex, create_new in ((True, True), (False, True), (True, False), (False, False)):
        for nv in (nvec, nvec.Negate()):
            try:
                rb = Rebar.CreateFromCurves(
                    doc, RebarStyle.Standard, bar_type,
                    None, None, host_wall, nv, cl,
                    RebarHookOrientation.Left, RebarHookOrientation.Left,
                    use_ex, create_new,
                )
            except Exception:
                rb = None
            if rb is not None:
                break
        if rb is not None:
            break

    if rb is None:
        return None, False, False

    if n_bars > 1 and distrib_ft is not None and float(distrib_ft) >= tol:
        try:
            accessor = rb.GetShapeDrivenAccessor()
            accessor.SetLayoutAsFixedNumber(
                int(n_bars), float(distrib_ft), False, True, True,
            )
        except Exception:
            pass

    did_bot = bool(bot_leg is not None)
    did_top = bool(top_leg is not None)
    return _stamp_armadura_arainco(rb, layer_index=layer_index), did_bot, did_top


def _bar_type_for_diameter_mm(doc, diam_mm, fallback=None):
    """``RebarBarType`` más cercano al diámetro nominal (mm)."""
    if doc is None:
        return fallback
    try:
        target = float(diam_mm)
    except Exception:
        return fallback
    best = None
    best_diff = None
    try:
        for bt in FilteredElementCollector(doc).OfClass(RebarBarType):
            d = _bar_diameter_mm(bt)
            diff = abs(d - target)
            if best is None or diff < best_diff:
                best = bt
                best_diff = diff
                if diff < 0.01:
                    break
    except Exception:
        pass
    return best or fallback


def _cabezal_project_pt_extremo_axes_ft(geom, pt):
    """Cotas long/trans (ft) de ``pt`` en ejes del extremo del muro."""
    if geom is None or pt is None:
        return 0.0, 0.0
    try:
        delta = pt.Subtract(geom[u"pt_extremo"])
        long_c = float(delta.DotProduct(geom[u"vector_longitudinal"]))
        trans_c = float(delta.DotProduct(geom[u"normal_muro"]))
        return long_c, trans_c
    except Exception:
        return 0.0, 0.0


def _cabezal_layer_long_coord_ft(geom, p_lo, offset_long_ft, ex_cfg):
    """Coord. longitudinal (ft) alineada con barras longitudinales del extremo."""
    if p_lo is not None and geom is not None:
        long_c, _ = _cabezal_project_pt_extremo_axes_ft(geom, p_lo)
        return float(long_c)
    return float(offset_long_ft)


def _cabezal_stirrup_envelope_from_line_jobs_ft(
    wall, extremo, ex_cfg, doc, layer_indices, line_jobs,
    segment_ctx=None, stack_index=0, bar_type_fallback=None,
):
    """
    Envelope perimetral desde jobs de línea ya calculados (mismas ``p_lo`` que las barras).
    """
    geom = _wall_longitudinal_at_extremo(wall, extremo)
    if geom is None:
        return None, None, None, None, u"Sin geometría longitudinal del muro."

    conf = normalize_cabezal_confinement(
        ex_cfg.get(u"confinement"), ex_cfg.get(u"n_capas"),
    )
    pad_mm = (
        float(conf.get(u"stirrup_diam_mm") or CABEZAL_STIRRUP_DIAM_MM) * 0.5
        + CABEZAL_CONFINEMENT_STIRRUP_PAD_MM
    )
    pad_ft = _mm_to_internal(pad_mm)

    idx_set = set()
    for li in layer_indices or []:
        try:
            idx_set.add(int(li))
        except Exception:
            pass
    conf_norm = normalize_cabezal_confinement(
        ex_cfg.get(u"confinement"), ex_cfg.get(u"n_capas"),
    )
    if cabezal_confinement_is_perimeter(conf_norm.get(u"type")):
        idx_set.update(
            cabezal_perimeter_stirrup_layer_indices(
                cabezal_effective_n_capas(ex_cfg),
            ),
        )

    jobs_by_layer = {}
    for job in line_jobs or []:
        try:
            if int(job.get(u"layer_index", -1)) not in idx_set:
                continue
        except Exception:
            continue
        if job.get(u"extremo") != extremo:
            continue
        try:
            jw = job.get(u"wall")
            if jw is not None and wall is not None and jw.Id != wall.Id:
                continue
        except Exception:
            pass
        li = int(job.get(u"layer_index", 0))
        jobs_by_layer.setdefault(li, []).append(job)

    long_vals = []
    trans_vals = []
    contributed = set()
    for li in sorted(idx_set):
        layer_jobs = jobs_by_layer.get(li)
        if not layer_jobs:
            continue
        job = layer_jobs[0]
        p_lo = job.get(u"p_lo")
        if p_lo is None:
            continue
        try:
            n_bars = int(job.get(u"n_bars", CABEZAL_MIN_BARRAS_POR_CAPA))
        except Exception:
            n_bars = CABEZAL_MIN_BARRAS_POR_CAPA
        n_bars = max(
            CABEZAL_MIN_BARRAS_POR_CAPA,
            min(CABEZAL_MAX_BARRAS_POR_CAPA, n_bars),
        )
        bar_type = job.get(u"bar_type")
        if bar_type is None:
            layers = cabezal_active_layers(ex_cfg)
            ly = layers[li] if li < len(layers) else {}
            bar_type = _resolver_bar_type_for_layer(
                doc, ex_cfg, ly, bar_type_fallback,
                segment_ctx=segment_ctx,
                stack_index=stack_index,
                layer_index=li,
            )
        if bar_type is None:
            return None, None, None, None, u"Capa {0}: sin RebarBarType.".format(li + 1)
        distrib_ft = float(job.get(u"distrib_ft") or 0.0)
        if distrib_ft < 1e-12:
            _, _, distrib_ft, err_geom = _cabezal_capa_line_endpoints(
                wall, extremo, li, bar_type,
                job.get(u"conf_type") or bar_type,
                layer_spacing_mm=float(
                    ex_cfg.get(u"layer_spacing_mm") or CABEZAL_LAYER_PITCH_MM,
                ),
                doc=doc, ex_cfg=ex_cfg,
            )
            if err_geom:
                return None, None, None, None, err_geom
        long_c, trans_origin = _cabezal_project_pt_extremo_axes_ft(geom, p_lo)
        bar_r_ft = _mm_to_internal(_bar_diameter_mm(bar_type) * 0.5)
        trans_coords = _cabezal_layer_bar_trans_coords_ft(
            trans_origin, distrib_ft, n_bars,
        )
        long_vals.extend([
            float(long_c) - bar_r_ft - pad_ft,
            float(long_c) + bar_r_ft + pad_ft,
        ])
        for tc in trans_coords:
            trans_vals.extend([
                float(tc) - bar_r_ft - pad_ft,
                float(tc) + bar_r_ft + pad_ft,
            ])
        contributed.add(int(li))

    if not long_vals or not trans_vals:
        return None, None, None, None, u"Sin capas para confinamiento."

    if contributed != idx_set:
        return None, None, None, None, u"Capas incompletas en line_jobs."

    return (
        min(long_vals), max(long_vals),
        min(trans_vals), max(trans_vals),
        None,
    )


def _cabezal_stirrup_envelope_ft(
    wall, extremo, ex_cfg, doc, layer_indices,
    segment_ctx=None, stack_index=0, bar_type_fallback=None,
    line_jobs=None,
):
    """
    Envelope perimetral en ft (ejes ``v_long`` / ``normal_muro`` desde ``pt_extremo``).

    Misma lógica que ``31_ArmadoMurosCabezal``: offsets escalares por capa +
    reparto ``SetLayoutAsFixedNumber`` (``long_c = offset_long_ft``, sin proyección
    3D de ``p_lo``). ``line_jobs`` se ignora aquí (solo sirve para barras long.).
    """
    geom = _wall_longitudinal_at_extremo(wall, extremo)
    if geom is None:
        return None, None, None, None, u"Sin geometría longitudinal del muro."

    layer_spacing_mm = float(
        ex_cfg.get(u"layer_spacing_mm") or CABEZAL_LAYER_PITCH_MM,
    )
    conf = normalize_cabezal_confinement(
        ex_cfg.get(u"confinement"), ex_cfg.get(u"n_capas"),
    )
    pad_mm = (
        float(conf.get(u"stirrup_diam_mm") or CABEZAL_STIRRUP_DIAM_MM) * 0.5
        + CABEZAL_CONFINEMENT_STIRRUP_PAD_MM
    )
    pad_ft = _mm_to_internal(pad_mm)
    espesor_ft = float(geom[u"espesor_ft"])
    dist_eje_cara = espesor_ft * 0.5

    long_vals = []
    trans_vals = []
    layer_long_mm = {}
    layers = cabezal_active_layers(ex_cfg)

    jobs_by_layer = {}
    for job in line_jobs or []:
        try:
            jli = int(job.get(u"layer_index", -1))
        except Exception:
            continue
        if job.get(u"extremo") != extremo:
            continue
        try:
            jw = job.get(u"wall")
            if jw is not None and wall is not None and jw.Id != wall.Id:
                continue
        except Exception:
            pass
        jobs_by_layer.setdefault(jli, []).append(job)

    for li in layer_indices or []:
        try:
            li = int(li)
        except Exception:
            continue
        if li < 0 or li >= len(layers):
            continue
        ly = layers[li]
        try:
            n_bars = int(ly.get(u"n_bars", CABEZAL_MIN_BARRAS_POR_CAPA))
        except Exception:
            n_bars = CABEZAL_MIN_BARRAS_POR_CAPA
        n_bars = max(
            CABEZAL_MIN_BARRAS_POR_CAPA,
            min(CABEZAL_MAX_BARRAS_POR_CAPA, n_bars),
        )
        bar_type = _resolver_bar_type_for_layer(
            doc, ex_cfg, ly, bar_type_fallback,
            segment_ctx=segment_ctx,
            stack_index=stack_index,
            layer_index=li,
        )
        if bar_type is None:
            return None, None, None, None, u"Capa {0}: sin RebarBarType.".format(li + 1)
        conf_type = _resolver_conf_bar_type(
            doc, ex_cfg, bar_type, bar_type_fallback,
        )
        if conf_type is None:
            conf_type = bar_type
        _p_lo, _p_hi, distrib_ft, err_geom = _cabezal_capa_line_endpoints(
            wall, extremo, li, bar_type, conf_type,
            layer_spacing_mm=layer_spacing_mm,
        )
        if err_geom:
            return None, None, None, None, err_geom
        offset_trans_mm, offset_long_mm = _cabezal_capa_offsets_mm(
            li, bar_type, conf_type, layer_spacing_mm, None,
        )
        offset_trans_ft = _mm_to_internal(offset_trans_mm)
        offset_long_ft = _mm_to_internal(offset_long_mm)
        trans_origin = dist_eje_cara - offset_trans_ft
        long_c = float(offset_long_ft)
        long_src = u"offset"
        job0 = (jobs_by_layer.get(int(li)) or [None])[0]
        if job0 is not None and job0.get(u"p_lo") is not None:
            long_proj, _ = _cabezal_project_pt_extremo_axes_ft(
                geom, job0.get(u"p_lo"),
            )
            long_c = float(long_proj)
            long_src = u"p_lo"
        bar_r_ft = _mm_to_internal(_bar_diameter_mm(bar_type) * 0.5)
        trans_coords = _cabezal_layer_bar_trans_coords_ft(
            trans_origin, distrib_ft, n_bars,
        )
        long_vals.extend([
            long_c - bar_r_ft - pad_ft,
            long_c + bar_r_ft + pad_ft,
        ])
        layer_long_mm[int(li)] = {
            u"long_mm": round(float(long_c) * 304.8, 2),
            u"offset_long_mm": round(float(offset_long_ft) * 304.8, 2),
            u"src": long_src,
        }
        for tc in trans_coords:
            trans_vals.extend([
                float(tc) - bar_r_ft - pad_ft,
                float(tc) + bar_r_ft + pad_ft,
            ])

    if not long_vals or not trans_vals:
        return None, None, None, None, u"Sin capas para confinamiento."

    return (
        min(long_vals), max(long_vals),
        min(trans_vals), max(trans_vals),
        None,
    )


def _resolve_cabezal_stirrup_hook_135(doc):
    """``RebarHookType`` 135° para estribos de confinamiento (mismo criterio que columnas)."""
    if doc is None:
        return None, u"Documento no válido."
    if l135 is None:
        return None, u"Módulo rebar_extender_l_ganchos_135_rps no disponible."
    try:
        largo_mm = float(getattr(l135, u"HOOK_LENGTH_MM_135", 100.0))
        hid, err = l135._resolve_rebar_hook_135_id(doc, largo_mm)
    except Exception as ex:
        try:
            return None, u"Resolver gancho 135°: {0}".format(unicode(ex))
        except Exception:
            return None, u"Resolver gancho 135°: {0}".format(str(ex))
    if hid is None or hid == ElementId.InvalidElementId:
        return None, err or u"RebarHookType 135° no resuelto en el proyecto."
    try:
        ht = doc.GetElement(hid)
        if isinstance(ht, RebarHookType):
            return ht, None
    except Exception:
        pass
    return None, err or u"RebarHookType 135° no resuelto en el proyecto."


def _cabezal_hook_orient_inward(tangent, plane_normal, at_pt, interior_pt):
    """
    ``RebarHookOrientation`` (Left/Right) para que el gancho 135° apunte hacia
    ``interior_pt`` desde ``at_pt``. ``tangent`` = sentido del tramo en ese extremo.
    """
    try:
        ln = float(tangent.GetLength())
    except Exception:
        ln = 0.0
    if ln < 1e-12:
        return RebarHookOrientation.Left
    t = tangent.Multiply(1.0 / ln)
    try:
        pn_len = float(plane_normal.GetLength())
    except Exception:
        pn_len = 0.0
    if pn_len < 1e-12:
        return RebarHookOrientation.Left
    n = plane_normal.Multiply(1.0 / pn_len)
    to_axis = interior_pt.Subtract(at_pt)
    h = float(to_axis.DotProduct(n))
    to_plane = to_axis.Subtract(n.Multiply(h))
    tpl = float(to_plane.GetLength())
    if tpl < 1e-12:
        return RebarHookOrientation.Left
    d_in = to_plane.Multiply(1.0 / tpl)
    lat = n.CrossProduct(t)
    l = float(lat.GetLength())
    if l < 1e-12:
        return RebarHookOrientation.Left
    lat_u = lat.Multiply(1.0 / l)
    return (
        RebarHookOrientation.Right
        if float(lat_u.DotProduct(d_in)) < 0.0
        else RebarHookOrientation.Left
    )


def _flip_cabezal_hook_orient(orient):
    if orient == RebarHookOrientation.Left:
        return RebarHookOrientation.Right
    return RebarHookOrientation.Left


def _cabezal_confinement_hook_orientations(p_corner, p_next, p_prev, interior_pt, plane_normal):
    """
    Orientaciones en la esquina de cierre (path CW desde ``p_corner``).

    Revit invierte el sentido efectivo del gancho 135° respecto al cálculo geométrico
    en planta; se invierte Left↔Right en ambos extremos para que apunten al interior.
    """
    t_start = p_next.Subtract(p_corner)
    t_end = p_corner.Subtract(p_prev)
    o_start = _cabezal_hook_orient_inward(
        t_start.Negate(), plane_normal, p_corner, interior_pt,
    )
    o_end = _cabezal_hook_orient_inward(
        t_end.Negate(), plane_normal, p_corner, interior_pt,
    )
    return _flip_cabezal_hook_orient(o_start), _flip_cabezal_hook_orient(o_end)


def _cabezal_stirrup_bar_radius_ft(stirrup_type):
    try:
        return _mm_to_internal(_bar_diameter_mm(stirrup_type) * 0.5)
    except Exception:
        return _mm_to_internal(5.0)


def _cabezal_stirrup_inset_envelope_bounds(
    long_min, long_max, trans_min, trans_max, stirrup_r_ft,
):
    """
    Inset simétrico del rectángulo en ``stirrup_r`` (eje barra) para mantener
    paralelismo A/C sin trapecio. Compensa en parte el descuento por ganchos 135°.
    """
    try:
        r = float(stirrup_r_ft)
    except Exception:
        r = 0.0
    if r <= 1e-12:
        return long_min, long_max, trans_min, trans_max
    try:
        ln = float(long_max) - float(long_min)
        tn = float(trans_max) - float(trans_min)
    except Exception:
        return long_min, long_max, trans_min, trans_max
    if ln <= 2.0 * r or tn <= 2.0 * r:
        return long_min, long_max, trans_min, trans_max
    return (
        float(long_min) + r,
        float(long_max) - r,
        float(trans_min) + r,
        float(trans_max) - r,
    )


def _clear_cabezal_stirrup_hook_rotations(rebar):
    """Fuerza rotación de gancho 0° (orientación solo vía Left/Right en creación)."""
    if rebar is None:
        return
    for end_idx in (0, 1):
        try:
            fn = getattr(rebar, u"SetTerminationRotationAngle", None)
            if fn is not None:
                fn(int(end_idx), 0.0)
                continue
        except Exception:
            pass
        try:
            fn = getattr(rebar, u"SetHookRotationAngle", None)
            if fn is not None:
                fn(0.0, int(end_idx))
        except Exception:
            pass


def _cabezal_rebar_shape_name_key(name):
    try:
        t = unicode(name or u"")
    except Exception:
        t = u""
    try:
        t = t.replace(u"\u00A0", u" ")
    except Exception:
        pass
    return t.strip()


def _cabezal_rebar_shape_display_name(shape):
    if shape is None:
        return u""
    try:
        p = shape.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p is not None:
            s = p.AsString()
            if s:
                return _cabezal_rebar_shape_name_key(s)
    except Exception:
        pass
    try:
        p = shape.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if p is not None:
            s = p.AsString()
            if s:
                return _cabezal_rebar_shape_name_key(s)
    except Exception:
        pass
    return _cabezal_rebar_shape_name_key(getattr(shape, "Name", None))


def _find_cabezal_confinement_rebar_shape(doc, nombre):
    """``RebarShape`` por nombre visible (p. ej. «10»)."""
    if doc is None:
        return None
    key = _cabezal_rebar_shape_name_key(nombre)
    if not key:
        return None
    try:
        key_lower = key.lower()
    except Exception:
        key_lower = key
    key_digits = u"".join(ch for ch in key if ch in u"0123456789")
    match_lower = None
    match_digits = None
    try:
        shapes = FilteredElementCollector(doc).OfClass(RebarShape)
    except Exception:
        return None
    for sh in shapes:
        if sh is None:
            continue
        sn = _cabezal_rebar_shape_display_name(sh)
        if not sn:
            continue
        if sn == key:
            return sh
        try:
            sn_low = sn.lower()
        except Exception:
            sn_low = sn
        if sn_low == key_lower and match_lower is None:
            match_lower = sh
        dig = u"".join(ch for ch in sn if ch in u"0123456789")
        if dig == key and match_digits is None:
            match_digits = sh
        elif key_digits and dig == key_digits and match_digits is None:
            match_digits = sh
    return match_lower or match_digits


def _cabezal_stirrup_curve_variants(curves_list):
    variants = []
    if curves_list is not None:
        try:
            if int(curves_list.Count) >= 2:
                variants.append(curves_list)
        except Exception:
            variants.append(curves_list)
    try:
        if curves_list is None or int(curves_list.Count) < 2:
            return variants
        rev = List[Curve]()
        n = int(curves_list.Count)
        for i in range(n - 1, -1, -1):
            c = curves_list[i]
            try:
                rev.Add(c.CreateReversed())
            except Exception:
                rev.Add(c)
        if int(rev.Count) >= 2:
            variants.append(rev)
    except Exception:
        pass
    return variants


def _try_create_cabezal_stirrup_from_curves(
    doc, wall, bar_type, hook_type, curves_list, plane_normal,
    p_corner, p_next, p_prev, p_center,
):
    """
    Estribo perimetral con ``RebarShape`` «10» (``CreateFromCurvesAndShape``)
    y ganchos 135° hacia el interior del muro.
    """
    if hook_type is None:
        return None, u"RebarHookType 135° no válido."
    shape = _find_cabezal_confinement_rebar_shape(
        doc, CABEZAL_CONFINEMENT_STIRRUP_SHAPE_NAME,
    )
    if shape is None:
        return None, (
            u"No se encontró RebarShape «{0}» para estribo de confinamiento.".format(
                CABEZAL_CONFINEMENT_STIRRUP_SHAPE_NAME,
            )
        )
    last_err = None
    normals = []
    if plane_normal is not None:
        try:
            if float(plane_normal.GetLength()) > 1e-12:
                normals.append(plane_normal.Normalize())
                normals.append(plane_normal.Normalize().Negate())
        except Exception:
            pass
    if not normals:
        normals = [XYZ.BasisZ, XYZ.BasisZ.Negate()]
    invalid = ElementId.InvalidElementId
    curve_variants = _cabezal_stirrup_curve_variants(curves_list)
    for curves_clr in curve_variants:
        if curves_clr is None:
            continue
        try:
            if int(curves_clr.Count) < 2:
                continue
        except Exception:
            continue
        for nv in normals:
            o_start, o_end = _cabezal_confinement_hook_orientations(
                p_corner, p_next, p_prev, p_center, nv,
            )
            try:
                rb = Rebar.CreateFromCurvesAndShape(
                    doc,
                    shape,
                    bar_type,
                    hook_type,
                    hook_type,
                    wall,
                    nv,
                    curves_clr,
                    o_start,
                    o_end,
                    0.0,
                    0.0,
                    invalid,
                    invalid,
                )
                if rb is not None:
                    _clear_cabezal_stirrup_hook_rotations(rb)
                    return _stamp_armadura_arainco(rb), None
            except Exception as ex:
                try:
                    last_err = unicode(ex)
                except Exception:
                    last_err = str(ex)
            try:
                rb = Rebar.CreateFromCurvesAndShape(
                    doc,
                    shape,
                    bar_type,
                    hook_type,
                    hook_type,
                    wall,
                    nv,
                    curves_clr,
                    o_start,
                    o_end,
                )
                if rb is not None:
                    _clear_cabezal_stirrup_hook_rotations(rb)
                    return _stamp_armadura_arainco(rb), None
            except Exception as ex:
                try:
                    last_err = unicode(ex)
                except Exception:
                    last_err = str(ex)
    if last_err:
        return None, (
            u"CreateFromCurvesAndShape (RebarShape «{0}»): {1}".format(
                CABEZAL_CONFINEMENT_STIRRUP_SHAPE_NAME, last_err,
            )
        )
    return None, (
        u"CreateFromCurvesAndShape falló para RebarShape «{0}».".format(
            CABEZAL_CONFINEMENT_STIRRUP_SHAPE_NAME,
        )
    )


def _cabezal_tie_layer_geometry_ft(
    wall, extremo, ex_cfg, doc,
    segment_ctx=None, stack_index=0, bar_type_fallback=None,
    layer_index=None,
):
    """
    Geometría de traba en capa ``layer_index``: pata en tangente interior + empalmes.

    Retorna curvas 3D (``List[Curve]``), puntos extremos, centro interior y error.
    """
    try:
        li = int(
            layer_index if layer_index is not None else CABEZAL_TIE_LAYER_INDEX,
        )
    except Exception:
        li = CABEZAL_TIE_LAYER_INDEX
    geom = _wall_longitudinal_at_extremo(wall, extremo)
    if geom is None:
        return None, None, None, None, u"Sin geometría longitudinal del muro."

    pt_ext = geom[u"pt_extremo"]
    v_long = geom[u"vector_longitudinal"]
    n_muro = geom[u"normal_muro"]
    espesor_ft = float(geom[u"espesor_ft"])
    dist_eje_cara = espesor_ft * 0.5
    cover_ft = _mm_to_internal(CABEZAL_COVER_MM)

    layer_spacing_mm = float(
        ex_cfg.get(u"layer_spacing_mm") or CABEZAL_LAYER_PITCH_MM,
    )
    conf = normalize_cabezal_confinement(
        ex_cfg.get(u"confinement"), ex_cfg.get(u"n_capas"),
    )
    tie_diam_mm = float(conf.get(u"stirrup_diam_mm") or CABEZAL_STIRRUP_DIAM_MM)
    tie_r_ft = _mm_to_internal(tie_diam_mm * 0.5)

    layers = cabezal_active_layers(ex_cfg)
    if li < 0 or li >= len(layers):
        return None, None, None, None, u"Capa [1] no configurada."

    ly = layers[li]
    try:
        n_bars = int(ly.get(u"n_bars", CABEZAL_MIN_BARRAS_POR_CAPA))
    except Exception:
        n_bars = CABEZAL_MIN_BARRAS_POR_CAPA
    n_bars = max(
        CABEZAL_MIN_BARRAS_POR_CAPA,
        min(CABEZAL_MAX_BARRAS_POR_CAPA, n_bars),
    )
    bar_type = _resolver_bar_type_for_layer(
        doc, ex_cfg, ly, bar_type_fallback,
        segment_ctx=segment_ctx,
        stack_index=stack_index,
        layer_index=li,
    )
    if bar_type is None:
        return None, None, None, None, u"Capa [1]: sin RebarBarType."
    conf_type = _resolver_conf_bar_type(
        doc, ex_cfg, bar_type, bar_type_fallback,
    )
    if conf_type is None:
        conf_type = bar_type
    _p_lo, _p_hi, distrib_ft, err_geom = _cabezal_capa_line_endpoints(
        wall, extremo, li, bar_type, conf_type,
        layer_spacing_mm=layer_spacing_mm,
        doc=doc, ex_cfg=ex_cfg,
    )
    if err_geom:
        return None, None, None, None, err_geom
    offset_trans_mm, offset_long_mm = _cabezal_capa_offsets_mm(
        li, bar_type, conf_type, layer_spacing_mm, None,
    )
    offset_trans_ft = _mm_to_internal(offset_trans_mm)
    offset_long_ft = _mm_to_internal(offset_long_mm)
    trans_origin = dist_eje_cara - offset_trans_ft
    if cabezal_extremo_es_encuentro_l(ex_cfg):
        try:
            delta = _p_lo.Subtract(geom[u"pt_extremo"])
            into = geom[u"vector_longitudinal"]
            long_bar = abs(float(delta.DotProduct(into)))
        except Exception:
            long_bar = offset_long_ft
    else:
        long_bar = offset_long_ft
    bar_r_ft = _mm_to_internal(_bar_diameter_mm(bar_type) * 0.5)
    tie_offset_ft = _cabezal_tie_offset_ft(bar_type, tie_diam_mm)
    long_tie = float(long_bar) + tie_offset_ft
    trans_coords = _cabezal_layer_bar_trans_coords_ft(
        trans_origin, distrib_ft, n_bars,
    )
    if not trans_coords:
        return None, None, None, None, u"Capa [1]: sin cotas transversales."

    trans_bar_min = min(trans_coords)
    trans_bar_max = max(trans_coords)
    trans_cover_hi = float(dist_eje_cara) - cover_ft
    trans_cover_lo = float(-dist_eje_cara) + cover_ft
    trans_hi = max(trans_bar_max + bar_r_ft + tie_r_ft, trans_cover_hi)
    trans_lo = min(trans_bar_min - bar_r_ft - tie_r_ft, trans_cover_lo)
    if trans_hi < trans_lo:
        trans_hi, trans_lo = trans_lo, trans_hi
    if (trans_hi - trans_lo) < 1e-9:
        return None, None, None, None, u"Traba capa [1]: espesor nulo."

    z_bot, z_top = _wall_z_bounds_ft(wall)
    foundation_drop_ft = _cabezal_stirrup_foundation_drop_ft(doc, wall)
    z_plane = float(z_bot) - float(foundation_drop_ft)

    def _pt(long_c, trans_c):
        base = (
            pt_ext
            + v_long.Multiply(float(long_c))
            + n_muro.Multiply(float(trans_c))
        )
        return XYZ(float(base.X), float(base.Y), z_plane)

    p_top = _pt(long_tie, trans_hi)
    p_bot = _pt(long_tie, trans_lo)
    interior_pt = _pt(long_bar, 0.5 * (trans_bar_min + trans_bar_max))

    try:
        cl = List[Curve]()
        prev = p_top
        ordered = sorted(trans_coords, reverse=True)
        for tc in ordered:
            p_on = _pt(long_tie, float(tc))
            p_bar = _pt(long_bar, float(tc))
            try:
                tol = float(doc.Application.ShortCurveTolerance)
            except Exception:
                tol = 1e-6
            if prev.DistanceTo(p_on) > tol:
                cl.Add(Line.CreateBound(prev, p_on))
            if p_on.DistanceTo(p_bar) > tol:
                cl.Add(Line.CreateBound(p_on, p_bar))
            if p_bar.DistanceTo(p_on) > tol:
                cl.Add(Line.CreateBound(p_bar, p_on))
            prev = p_on
        if prev.DistanceTo(p_bot) > tol:
            cl.Add(Line.CreateBound(prev, p_bot))
        if cl.Count < 1:
            cl.Add(Line.CreateBound(p_top, p_bot))
    except Exception as ex_path:
        try:
            return None, None, None, None, u"Curvas traba: {0}".format(unicode(ex_path))
        except Exception:
            return None, None, None, None, u"Curvas traba: {0}".format(str(ex_path))

    return cl, p_top, p_bot, interior_pt, None


def _cabezal_tie_plane_normals(geom):
    """
    Normal al plano de la traba (array en +Z).

    Igual que estribos perimetrales: ``XYZ.BasisZ``. No usar ``v_long × n_muro``:
    con muros horizontales en Y puede dar -Z e invertir el array hacia abajo.
    """
    return [XYZ.BasisZ, XYZ.BasisZ.Negate()]


def _cabezal_curves_to_py_list(curves_list):
    out = []
    if curves_list is None:
        return out
    try:
        n = int(curves_list.Count)
    except Exception:
        n = 0
    for i in range(n):
        try:
            out.append(curves_list[i])
        except Exception:
            pass
    return out


def _cabezal_tie_hook_orientations(p_top, p_bot, interior_pt, plane_normal):
    """
    ``RebarHookOrientation`` en extremos de traba abierta (ext. 0 = p_top, ext. 1 = p_bot).

    Ganchos hacia ``interior_pt`` (barras capa [1]). Revit invierte el sentido en el
    extremo superior respecto al inferior; solo el superior lleva flip de compensación.
    """
    t_down = p_bot.Subtract(p_top)
    t_up = p_top.Subtract(p_bot)
    o_top = _flip_cabezal_hook_orient(
        _cabezal_hook_orient_inward(
            t_down.Negate(), plane_normal, p_top, interior_pt,
        ),
    )
    o_bot = _cabezal_hook_orient_inward(
        t_up.Negate(), plane_normal, p_bot, interior_pt,
    )
    return o_top, o_bot


def _try_create_cabezal_tie_from_curves(
    doc, wall, bar_type, hook_type, curves_list, plane_normals,
    p_start, p_end, interior_pt,
):
    """``CreateFromCurves`` traba con ganchos 135° — orientación geométrica primero."""
    if hook_type is None or curves_list is None:
        return None, u"RebarHookType 135° no válido."
    try:
        n_curves = int(curves_list.Count)
    except Exception:
        n_curves = 0
    if n_curves < 1:
        return None, u"Sin curvas para traba."

    hook_el = hook_type
    if not isinstance(hook_el, RebarHookType):
        try:
            hook_el = doc.GetElement(hook_type)
        except Exception:
            hook_el = None
    if hook_el is None:
        return None, u"RebarHookType 135° no resuelto."

    py_curves = _cabezal_curves_to_py_list(curves_list)
    if not py_curves:
        return None, u"Sin curvas para traba."

    last_err = None
    orient_pairs = (
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
    )
    styles = (RebarStyle.StirrupTie, RebarStyle.Standard)
    flag_pairs = ((True, True), (True, False), (False, True), (False, False))

    def _create_once(curves, nv, style, o0, o1, use_ex, create_new):
        try:
            rb = Rebar.CreateFromCurves(
                doc,
                style,
                bar_type,
                hook_el,
                hook_el,
                wall,
                nv,
                curves,
                o0,
                o1,
                use_ex,
                create_new,
            )
            if rb is not None:
                _clear_cabezal_stirrup_hook_rotations(rb)
                return _stamp_armadura_arainco(rb), None
        except Exception as ex:
            try:
                return None, unicode(ex)
            except Exception:
                return None, str(ex)
        return None, None

    for nv in plane_normals:
        o_top, o_bot = _cabezal_tie_hook_orientations(
            p_start, p_end, interior_pt, nv,
        )
        if l135 is not None:
            try:
                hid = hook_el.Id
            except Exception:
                hid = None
            if hid is not None and hid != ElementId.InvalidElementId:
                for style in styles:
                    try:
                        rb = l135._try_create_l_with_hook_types_both_ends(
                            doc, py_curves, wall, nv, bar_type, style,
                            o_top, o_bot, hid,
                        )
                        if rb is not None:
                            _clear_cabezal_stirrup_hook_rotations(rb)
                            return _stamp_armadura_arainco(rb), None
                    except Exception:
                        pass
        for style in styles:
            for use_ex, create_new in flag_pairs:
                rb, err = _create_once(
                    curves_list, nv, style, o_top, o_bot, use_ex, create_new,
                )
                if rb is not None:
                    return _stamp_armadura_arainco(rb), None
                if err:
                    last_err = err

    for nv in plane_normals:
        for style in styles:
            for o0, o1 in orient_pairs:
                for use_ex, create_new in flag_pairs:
                    rb, err = _create_once(
                        curves_list, nv, style, o0, o1, use_ex, create_new,
                    )
                    if rb is not None:
                        return _stamp_armadura_arainco(rb), None
                    if err:
                        last_err = err
    return None, last_err or u"CreateFromCurves traba: sin variante válida."


def _create_cabezal_confinement_tie(
    doc, wall, extremo, ex_cfg,
    segment_ctx=None, stack_index=0, bar_type_fallback=None,
    layer_index=None,
):
    """
    Traba en capa ``layer_index`` @ spacing a lo largo de Z (sin estribo perimetral).

    Con fundación unida: plano y array desde ``z_bot - 300 mm`` (igual que estribos).
    """
    conf_type, _ = cabezal_active_confinement(ex_cfg)
    try:
        n_capas = int(ex_cfg.get(u"n_capas", CABEZAL_MIN_CAPAS))
    except Exception:
        n_capas = CABEZAL_MIN_CAPAS
    tie_layers = cabezal_confinement_tie_layer_indices(
        normalize_cabezal_confinement(
            ex_cfg.get(u"confinement"), n_capas,
        ),
        n_capas,
    )
    if conf_type == CABEZAL_CONFINEMENT_NONE:
        return None, None
    if (
        not cabezal_confinement_is_tie_layer_1(conf_type)
        and not tie_layers
    ):
        return None, None
    try:
        li = int(
            layer_index if layer_index is not None else (
                tie_layers[0] if tie_layers else CABEZAL_TIE_LAYER_INDEX
            ),
        )
    except Exception:
        li = CABEZAL_TIE_LAYER_INDEX

    geom = _wall_longitudinal_at_extremo(wall, extremo)
    plane_normals = _cabezal_tie_plane_normals(geom)

    curves, p_top, p_bot, interior_pt, err = _cabezal_tie_layer_geometry_ft(
        wall, extremo, ex_cfg, doc,
        segment_ctx=segment_ctx,
        stack_index=stack_index,
        bar_type_fallback=bar_type_fallback,
        layer_index=li,
    )
    if err:
        return None, err

    z_bot, z_top = _wall_z_bounds_ft(wall)
    foundation_drop_ft = _cabezal_stirrup_foundation_drop_ft(doc, wall)
    z_array_bot = float(z_bot) - float(foundation_drop_ft)
    array_len = float(z_top) - z_array_bot
    if array_len < 1e-9:
        return None, u"Altura de muro nula para traba."

    conf = normalize_cabezal_confinement(
        ex_cfg.get(u"confinement"), ex_cfg.get(u"n_capas"),
    )
    tie_type = _bar_type_for_diameter_mm(
        doc, conf.get(u"stirrup_diam_mm"), bar_type_fallback,
    )
    if tie_type is None:
        return None, u"Sin RebarBarType ø{0} mm para traba.".format(
            conf.get(u"stirrup_diam_mm"),
        )

    hook_type, hook_err = _resolve_cabezal_stirrup_hook_135(doc)
    if hook_type is None:
        return None, hook_err or u"Sin RebarHookType 135° para traba."

    rb, create_err = _try_create_cabezal_tie_from_curves(
        doc, wall, tie_type, hook_type, curves, plane_normals,
        p_top, p_bot, interior_pt,
    )
    if rb is None:
        try:
            cl_simple = List[Curve]()
            cl_simple.Add(Line.CreateBound(p_top, p_bot))
            rb, create_err = _try_create_cabezal_tie_from_curves(
                doc, wall, tie_type, hook_type, cl_simple, plane_normals,
                p_top, p_bot, interior_pt,
            )
        except Exception as ex_fb:
            try:
                create_err = unicode(ex_fb)
            except Exception:
                create_err = str(ex_fb)
    if rb is None:
        try:
            msg = u"CreateFromCurves traba: {0}".format(unicode(create_err))
        except Exception:
            msg = u"CreateFromCurves traba: {0}".format(str(create_err))
        return None, msg

    spacing_ft = _mm_to_internal(float(CABEZAL_CONFINEMENT_REBAR_SET_SPACING_MM))
    try:
        accessor = rb.GetShapeDrivenAccessor()
        accessor.SetLayoutAsMaximumSpacing(
            float(spacing_ft), float(array_len), True, True, False,
        )
    except Exception as ex_lay:
        try:
            return _stamp_armadura_arainco(rb), u"SetLayoutAsMaximumSpacing traba: {0}".format(unicode(ex_lay))
        except Exception:
            return _stamp_armadura_arainco(rb), u"SetLayoutAsMaximumSpacing traba: {0}".format(str(ex_lay))

    return _stamp_armadura_arainco(rb), None


def _create_cabezal_confinement_stirrup(
    doc, wall, extremo, ex_cfg,
    segment_ctx=None, stack_index=0, bar_type_fallback=None,
    line_jobs=None,
):
    """
    Estribo perimetrico @ spacing fijo a lo largo de la altura del muro (Z),
    con ``RebarShape`` «10» (``CreateFromCurvesAndShape``) y ganchos 135°
    hacia el interior del muro.

    Si el muro tiene fundación estructural unida, el plano del loop y el array
    bajan ``CABEZAL_STIRRUP_FOUNDATION_DROP_MM`` (300 mm) y la longitud del
    array crece en la misma cantidad (criterio alineado con estribos de columnas).
    """
    conf_type, layer_indices = cabezal_active_confinement(ex_cfg)
    if conf_type != CABEZAL_CONFINEMENT_PERIMETER_0_1 or not layer_indices:
        return None, None

    long_min, long_max, trans_min, trans_max, err = _cabezal_stirrup_envelope_ft(
        wall, extremo, ex_cfg, doc, layer_indices,
        segment_ctx=segment_ctx,
        stack_index=stack_index,
        bar_type_fallback=bar_type_fallback,
        line_jobs=line_jobs,
    )
    if err:
        return None, err

    geom = _wall_longitudinal_at_extremo(wall, extremo)
    if geom is None:
        return None, u"Sin geometría longitudinal del muro."

    pt_ext = geom[u"pt_extremo"]
    v_long = geom[u"vector_longitudinal"]
    n_muro = geom[u"normal_muro"]
    z_bot, z_top = _wall_z_bounds_ft(wall)
    foundation_drop_ft = _cabezal_stirrup_foundation_drop_ft(doc, wall)
    z_stirrup_bot = float(z_bot) - float(foundation_drop_ft)
    array_len = float(z_top) - z_stirrup_bot
    if array_len < 1e-9:
        return None, u"Altura de muro nula para estribos."

    conf = normalize_cabezal_confinement(
        ex_cfg.get(u"confinement"), ex_cfg.get(u"n_capas"),
    )
    stirrup_type = _bar_type_for_diameter_mm(
        doc, conf.get(u"stirrup_diam_mm"), bar_type_fallback,
    )
    if stirrup_type is None:
        return None, u"Sin RebarBarType ø{0} mm para estribo.".format(
            conf.get(u"stirrup_diam_mm"),
        )

    stirrup_r_ft = _cabezal_stirrup_bar_radius_ft(stirrup_type)
    long_min, long_max, trans_min, trans_max = _cabezal_stirrup_inset_envelope_bounds(
        long_min, long_max, trans_min, trans_max, stirrup_r_ft,
    )

    def _corner(long_c, trans_c):
        base = (
            pt_ext
            + v_long.Multiply(float(long_c))
            + n_muro.Multiply(float(trans_c))
        )
        return XYZ(float(base.X), float(base.Y), float(z_stirrup_bot))

    p_br = _corner(long_max, trans_min)
    p_bl = _corner(long_min, trans_min)
    p_tl = _corner(long_min, trans_max)
    p_tr = _corner(long_max, trans_max)
    p_center = _corner(
        0.5 * (float(long_min) + float(long_max)),
        0.5 * (float(trans_min) + float(trans_max)),
    )

    try:
        cl = List[Curve]()
        # Rectángulo CW desde BR; cierre y ganchos en BR (columnas).
        path = [p_br, p_bl, p_tl, p_tr]
        for i in range(4):
            pa = path[i]
            pb = path[(i + 1) % 4]
            cl.Add(Line.CreateBound(pa, pb))
    except Exception as ex_ln:
        return None, u"Curvas estribo: {0}".format(ex_ln)

    hook_type, hook_err = _resolve_cabezal_stirrup_hook_135(doc)
    if hook_type is None:
        return None, hook_err or u"Sin RebarHookType 135° para estribo de confinamiento."

    plane_norm = XYZ.BasisZ
    spacing_ft = _mm_to_internal(float(CABEZAL_CONFINEMENT_REBAR_SET_SPACING_MM))
    rb, create_err = _try_create_cabezal_stirrup_from_curves(
        doc, wall, stirrup_type, hook_type, cl, plane_norm,
        p_br, p_bl, p_tr, p_center,
    )
    if rb is None:
        try:
            msg = u"CreateFromCurvesAndShape estribo (Shape 10): {0}".format(unicode(create_err))
        except Exception:
            msg = u"CreateFromCurvesAndShape estribo (Shape 10): {0}".format(str(create_err))
        return None, msg

    try:
        accessor = rb.GetShapeDrivenAccessor()
        # includeFirstBar=True, includeLastBar=False (Remove Last Bar en UI Revit).
        accessor.SetLayoutAsMaximumSpacing(
            float(spacing_ft), float(array_len), True, True, False,
        )
    except Exception as ex_lay:
        try:
            return _stamp_armadura_arainco(rb), u"SetLayoutAsMaximumSpacing: {0}".format(unicode(ex_lay))
        except Exception:
            return _stamp_armadura_arainco(rb), u"SetLayoutAsMaximumSpacing: {0}".format(str(ex_lay))

    return _stamp_armadura_arainco(rb), None


# ═══════════════════════════════════════════════════════════════════════════
# Pipeline principal
# ═══════════════════════════════════════════════════════════════════════════


def _cabezal_refrescar_vista_tras_lote(doc, uidoc, forzar=False):
    """Regenera/refresco tras lote; en modo rápido solo si ``forzar`` (cierre de flujo)."""
    if not forzar and MODO_EJECUCION_RAPIDA:
        return
    if doc is not None:
        try:
            doc.Regenerate()
        except Exception:
            pass
    if uidoc is None:
        return
    try:
        uidoc.RefreshActiveView()
    except Exception:
        pass
    try:
        uidoc.UpdateAllOpenViews()
    except Exception:
        pass


def _cabezal_refrescar_vista_fin_flujo(doc, uidoc):
    if _refrescar_vista_fin_flujo is not None:
        try:
            _refrescar_vista_fin_flujo(doc, uidoc)
            return
        except Exception:
            pass
    _cabezal_refrescar_vista_tras_lote(doc, uidoc, forzar=True)


def _cabezal_pbar_phase_title(base_title, total):
    try:
        t = max(int(total), 1)
    except Exception:
        t = 1
    return u"{} 0/{}".format(base_title, t)


def _cabezal_pbar_enabled(doc):
    try:
        return bool(pyrevit_progress_bar_enabled(doc))
    except Exception:
        return False


def _cabezal_pbar_start(title, count, doc=None):
    if not _cabezal_pbar_enabled(doc):
        return None
    if count is None or int(count) < 1:
        return None
    try:
        from pyrevit import forms as _pyrevit_forms

        pb = _pyrevit_forms.ProgressBar(title=title, cancellable=False)
        try:
            from System.Windows.Media import Color, SolidColorBrush

            pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(
                Color.FromRgb(91, 192, 222),
            )
        except Exception:
            pass
        return pb
    except Exception:
        return None


def _cabezal_pbar_step(pb, current_index, count, base_title):
    if pb is None:
        return
    c = int(count) if count else 0
    if c < 1:
        c = 1
    i = int(current_index) + 1
    try:
        if hasattr(pb, u"update_progress"):
            try:
                pb.update_progress(i, max_value=c)
            except TypeError:
                try:
                    pb.update_progress(i, max=c)
                except Exception:
                    pass
    except Exception:
        pass
    try:
        pb.title = u"{} {}/{}".format(base_title, i, c)
    except Exception:
        pass


def _cabezal_pbar_enter(pb):
    if pb is None:
        return False
    try:
        pb.__enter__()
        return True
    except Exception:
        return False


def _cabezal_pbar_exit(pb, pbar_open):
    if not pbar_open or pb is None:
        return
    try:
        pb.__exit__(None, None, None)
    except Exception:
        pass


def _cabezal_z_ft_to_m(z_ft):
    try:
        return float(
            UnitUtils.ConvertFromInternalUnits(float(z_ft), UnitTypeId.Meters),
        )
    except Exception:
        return float(z_ft) * 0.3048


def _cabezal_sort_seg_jobs(jobs, walls):
    """Orden abajo→arriba: muro, cota Z pie, capa, extremo."""

    def _key(sj):
        stack = _wall_stack_index(walls, sj.get(u"wall"))
        zs = float(sj.get(u"zs", 0.0))
        li = int(sj.get(u"layer_index", 0))
        ex = 0 if sj.get(u"extremo") == CABEZAL_EXTREMO_INICIO else 1
        return (stack, zs, li, ex)

    keyed = [(_key(sj), sj) for sj in (jobs or [])]
    keyed.sort()
    return [sj for _k, sj in keyed]


def _cabezal_sort_confinement_jobs(jobs, walls):
    """Orden abajo→arriba por cota base del muro, luego extremo."""

    def _key(cj):
        wall = cj.get(u"wall")
        try:
            z0, _z1 = _wall_z_bounds_ft(wall)
        except Exception:
            z0 = 0.0
        stack = int(cj.get(u"stack_idx", _wall_stack_index(walls, wall)))
        ex = 0 if cj.get(u"extremo") == CABEZAL_EXTREMO_INICIO else 1
        return (float(z0), stack, ex)

    keyed = [(_key(cj), cj) for cj in (jobs or [])]
    keyed.sort()
    return [cj for _k, cj in keyed]


def _cabezal_create_longitudinal_job(doc, sj, curve_tol, res, rebars_por):
    span = float(sj.get(u"span_seg", 0.0))
    if span < curve_tol:
        res[u"n_fail"] = int(res.get(u"n_fail", 0)) + 1
        return False
    rb, _did_bot, _did_top = _create_cabezal_rebar_with_optional_patas(
        doc,
        sj[u"bx"],
        sj[u"by"],
        sj[u"zs"],
        span,
        sj[u"wall"],
        sj[u"bar_type"],
        sj[u"normal_muro"],
        sj[u"vec_long"],
        sj[u"distrib_ft"],
        sj[u"n_bars"],
        sj[u"want_bot_pata"],
        sj[u"want_top_pata"],
        sj[u"pata_ft"],
        layer_index=int(sj.get(u"layer_index", 0) or 0),
    )
    if rb is not None:
        li = int(sj.get(u"layer_index", 0) or 0)
        _stamp_armadura_arainco(rb, layer_index=li)
        sj[u"rebar_id"] = rb.Id
        res[u"n_created"] = int(res.get(u"n_created", 0)) + 1
        res[u"n_bars_total"] = int(res.get(u"n_bars_total", 0)) + int(sj[u"n_bars"])
        try:
            long_ids = res.setdefault(u"rebars_longitudinales_ids", [])
            long_ids.append(rb.Id)
            tag_meta = res.setdefault(u"rebars_longitudinales_tag_meta", [])
            tag_meta.append({
                u"rebar_id": rb.Id,
                u"layer_index": int(sj.get(u"layer_index", 0)),
                u"wid": sj.get(u"wid"),
                u"extremo": sj.get(u"extremo"),
                u"zs": float(sj.get(u"zs", 0.0) or 0.0),
                u"span_seg": float(sj.get(u"span_seg", 0.0) or 0.0),
            })
        except Exception:
            pass
        for cwid in sj.get(u"contrib_wids") or []:
            try:
                rebars_por.setdefault(cwid, []).append(rb.Id)
            except Exception:
                pass
        return True
    res[u"n_fail"] = int(res.get(u"n_fail", 0)) + 1
    return False


def _cabezal_etiquetar_confinamiento_rebar(doc, view, res, cj, rb_conf):
    """Etiqueta una barra de confinamiento (individual o registro multihost)."""
    if view is None or _cab_tags is None or rb_conf is None:
        return False, None, None
    ex_cfg = cj.get(u"ex_cfg") or {}
    try:
        n_capas_conf = int(ex_cfg.get(u"n_capas", CABEZAL_MIN_CAPAS))
    except Exception:
        n_capas_conf = CABEZAL_MIN_CAPAS
    job_kind = cj.get(u"job_kind")
    if job_kind is None:
        job_kind = (
            u"tie"
            if cabezal_confinement_is_tie_layer_1(cj.get(u"conf_type"))
            else u"stirrup"
        )
    tag_mode = _cab_tags.confinement_tag_mode(
        cj.get(u"conf_type"), n_capas_conf, job_kind,
    )
    if tag_mode == u"conf_tag_multihost":
        try:
            _cab_tags.register_confinement_multihost_traba_pending(
                res, cj, rb_conf,
            )
            mh_order = res.setdefault(u"_conf_multihost_flush_order", [])
            mh_key = _cab_tags.confinement_multihost_group_key(cj)
            if mh_key not in mh_order:
                mh_order.append(mh_key)
        except Exception:
            pass
        return True, None, u"trabas multihost"
    if tag_mode != u"conf_tag":
        return False, None, None
    tag_map = res.get(u"_conf_tag_map")
    if tag_map is None:
        try:
            tag_map = _cab_tags.collect_confinement_tag_symbol_map(doc)
        except Exception:
            tag_map = {}
        res[u"_conf_tag_map"] = tag_map
    extra_mm = 0.0
    if job_kind == u"tie":
        extra_mm = _cab_tags.confinement_tag_extra_offset_mm_for_view(
            view,
            n_capas_conf,
            cj.get(u"conf_type"),
            job_kind,
            tie_layer_index=cj.get(u"tie_layer_index"),
        )
    anchor_override = None
    if (
        n_capas_conf == 3
        and job_kind == u"tie"
        and cabezal_confinement_is_perimeter(cj.get(u"conf_type"))
    ):
        anchor_override = _cab_tags.get_confinement_tipo2_stirrup_anchor(
            res, cj,
        )
    try:
        ok_tag, err_tag = _cab_tags.etiquetar_cabezal_estribo_confinamiento(
            doc,
            view,
            rb_conf,
            tag_map=tag_map,
            wall=cj.get(u"wall"),
            extremo=cj.get(u"extremo"),
            extra_offset_mm=extra_mm,
            n_capas=n_capas_conf,
            conf_type=cj.get(u"conf_type"),
            job_kind=job_kind,
            anchor_override=anchor_override,
        )
    except Exception as ex_tag:
        try:
            ok_tag, err_tag = False, unicode(ex_tag)
        except Exception:
            ok_tag, err_tag = False, str(ex_tag)
    if (
        ok_tag
        and n_capas_conf >= 3
        and job_kind == u"stirrup"
        and cabezal_confinement_is_perimeter(cj.get(u"conf_type"))
    ):
        try:
            anc = _cab_tags.compute_confinement_tag_anchor(
                doc,
                view,
                rb_conf,
                wall=cj.get(u"wall"),
                extremo=cj.get(u"extremo"),
            )
            if anc is not None:
                _cab_tags.register_confinement_tipo2_stirrup_anchor(
                    res, cj, anc,
                )
            head_e = _cab_tags.compute_confinement_stirrup_tag_head(
                doc,
                view,
                rb_conf,
                wall=cj.get(u"wall"),
                extremo=cj.get(u"extremo"),
                n_capas=n_capas_conf,
                conf_type=cj.get(u"conf_type"),
            )
            if head_e is not None:
                _cab_tags.register_confinement_tipo2_stirrup_head(
                    res, cj, head_e,
                )
        except Exception:
            pass
    kind_lbl = u"traba" if job_kind == u"tie" else u"estribo"
    return ok_tag, err_tag, kind_lbl


def _cabezal_create_confinement_job(
    doc, cj, res, rebars_por, bar_type_fallback=None, view=None,
    defer_tags=False,
):
    job_kind = cj.get(u"job_kind")
    if job_kind is None:
        job_kind = (
            u"tie"
            if cabezal_confinement_is_tie_layer_1(cj.get(u"conf_type"))
            else u"stirrup"
        )
    if job_kind == u"tie":
        rb_conf, err_conf = _create_cabezal_confinement_tie(
            doc,
            cj[u"wall"],
            cj[u"extremo"],
            cj[u"ex_cfg"],
            segment_ctx=cj.get(u"segment_ctx"),
            stack_index=cj.get(u"stack_idx", 0),
            bar_type_fallback=bar_type_fallback,
            layer_index=cj.get(u"tie_layer_index"),
        )
    else:
        rb_conf, err_conf = _create_cabezal_confinement_stirrup(
            doc,
            cj[u"wall"],
            cj[u"extremo"],
            cj[u"ex_cfg"],
            segment_ctx=cj.get(u"segment_ctx"),
            stack_index=cj.get(u"stack_idx", 0),
            bar_type_fallback=bar_type_fallback,
            line_jobs=cj.get(u"line_jobs"),
        )
    if rb_conf is not None:
        res[u"n_confinement_created"] = int(
            res.get(u"n_confinement_created", 0),
        ) + 1
        try:
            rebars_por.setdefault(cj[u"wid"], []).append(rb_conf.Id)
        except Exception:
            pass
        if err_conf:
            res[u"messages"].append(
                u"Muro {0} {1} confinamiento: {2}".format(
                    cj[u"wid"], cj[u"extremo"], err_conf,
                ),
            )
        if defer_tags:
            ex_cfg = cj.get(u"ex_cfg") or {}
            try:
                n_capas_conf = int(ex_cfg.get(u"n_capas", CABEZAL_MIN_CAPAS))
            except Exception:
                n_capas_conf = CABEZAL_MIN_CAPAS
            tag_mode = (
                _cab_tags.confinement_tag_mode(
                    cj.get(u"conf_type"), n_capas_conf, job_kind,
                )
                if _cab_tags is not None
                else None
            )
            if tag_mode == u"conf_tag_multihost":
                try:
                    _cab_tags.register_confinement_multihost_traba_pending(
                        res, cj, rb_conf,
                    )
                    mh_order = res.setdefault(u"_conf_multihost_flush_order", [])
                    mh_key = _cab_tags.confinement_multihost_group_key(cj)
                    if mh_key not in mh_order:
                        mh_order.append(mh_key)
                except Exception:
                    pass
            elif tag_mode == u"conf_tag":
                res.setdefault(u"_conf_inline_tag_pending", []).append({
                    u"cj": cj,
                    u"rebar_id": rb_conf.Id,
                })
        elif view is not None and _cab_tags is not None:
            ok_tag, err_tag, kind_lbl = _cabezal_etiquetar_confinamiento_rebar(
                doc, view, res, cj, rb_conf,
            )
            if kind_lbl and kind_lbl != u"trabas multihost":
                if ok_tag:
                    res[u"n_conf_tags_created"] = int(
                        res.get(u"n_conf_tags_created", 0),
                    ) + 1
                elif err_tag:
                    res[u"n_conf_tags_fail"] = int(
                        res.get(u"n_conf_tags_fail", 0),
                    ) + 1
                    if len(res.get(u"messages") or []) < 24:
                        res[u"messages"].append(
                            u"Muro {0} {1} etiqueta {2}: {3}".format(
                                cj[u"wid"],
                                cj[u"extremo"],
                                kind_lbl or u"confinamiento",
                                err_tag,
                            ),
                        )
        return True
    if err_conf:
        res[u"n_fail"] = int(res.get(u"n_fail", 0)) + 1
        kind = u"traba" if job_kind == u"tie" else u"confinamiento"
        cap_lbl = u""
        if job_kind == u"tie":
            try:
                cap_lbl = u" capa [{0}]".format(int(cj.get(u"tie_layer_index", 0)))
            except Exception:
                pass
        res[u"messages"].append(
            u"Muro {0} {1} {2}{3}: {4}".format(
                cj[u"wid"], cj[u"extremo"], kind, cap_lbl, err_conf,
            ),
        )
    else:
        res[u"n_fail"] = int(res.get(u"n_fail", 0)) + 1
    return False


def _cabezal_aplicar_etiquetado_confinamiento_animado(doc, view, res, uidoc):
    """Etiqueta confinamiento: individuales y multihost, lote a lote."""
    if _cab_tags is None or doc is None or view is None:
        return
    inline = list(res.pop(u"_conf_inline_tag_pending", None) or [])
    mh_pending = res.get(u"_conf_multihost_traba_pending") or {}
    if not inline and not mh_pending:
        return

    batch_tags = _tamano_lote_ejecucion(
        len(inline) + len(mh_pending), CABEZAL_TAGS_POR_LOTE_ANIMACION,
    )
    n_inline_lotes = int(math.ceil(float(len(inline)) / float(batch_tags))) if inline else 0
    mh_order = list(res.get(u"_conf_multihost_flush_order") or [])
    mh_keys = [k for k in mh_order if k in mh_pending]
    for k in mh_pending:
        if k not in mh_keys:
            mh_keys.append(k)
    n_mh_lotes = (
        int(math.ceil(float(len(mh_keys)) / float(batch_tags)))
        if mh_keys else 0
    )
    n_tag_lotes = n_inline_lotes + n_mh_lotes
    if n_tag_lotes <= 0:
        return

    pb_tags = _cabezal_pbar_start(
        _cabezal_pbar_phase_title(_CAB_PBAR_BASE_TAGS, n_tag_lotes),
        n_tag_lotes,
        doc=doc,
    )
    pbar_tags_open = _cabezal_pbar_enter(pb_tags)
    tag_lote_idx = [0]

    def _after_tag_batch():
        _cabezal_refrescar_vista_tras_lote(doc, uidoc)
        _cabezal_pbar_step(
            pb_tags, tag_lote_idx[0], n_tag_lotes, _CAB_PBAR_BASE_TAGS,
        )
        tag_lote_idx[0] += 1

    try:
        for i0 in range(0, len(inline), batch_tags):
            lote = inline[i0:i0 + batch_tags]
            i1 = min(i0 + batch_tags, len(inline))
            if len(lote) == 1:
                txn_name = u"Arainco: Cabezal muros — etiqueta confinamiento {0}/{1}".format(
                    i0 + 1, len(inline),
                )
            else:
                txn_name = (
                    u"Arainco: Cabezal muros — etiquetas confinamiento "
                    u"{0}–{1} de {2}".format(i0 + 1, i1, len(inline))
                )
            t = Transaction(doc, txn_name)
            t.Start()
            lote_ok = False
            try:
                for item in lote:
                    cj = item.get(u"cj")
                    rid = item.get(u"rebar_id")
                    rb = doc.GetElement(rid) if rid is not None else None
                    if cj is None or rb is None:
                        res[u"n_conf_tags_fail"] = int(
                            res.get(u"n_conf_tags_fail", 0),
                        ) + 1
                        continue
                    ok_tag, err_tag, kind_lbl = _cabezal_etiquetar_confinamiento_rebar(
                        doc, view, res, cj, rb,
                    )
                    if kind_lbl == u"trabas multihost":
                        continue
                    if ok_tag:
                        res[u"n_conf_tags_created"] = int(
                            res.get(u"n_conf_tags_created", 0),
                        ) + 1
                    else:
                        res[u"n_conf_tags_fail"] = int(
                            res.get(u"n_conf_tags_fail", 0),
                        ) + 1
                        if err_tag and len(res.get(u"messages") or []) < 24:
                            res[u"messages"].append(
                                u"Muro {0} {1} etiqueta {2}: {3}".format(
                                    cj.get(u"wid"),
                                    cj.get(u"extremo"),
                                    kind_lbl or u"confinamiento",
                                    err_tag,
                                ),
                            )
                t.Commit()
                lote_ok = True
            except Exception as ex_lote:
                try:
                    if t.HasStarted():
                        t.RollBack()
                except Exception:
                    pass
                res[u"messages"].append(
                    u"Etiquetado confinamiento lote {0}–{1}: {2}".format(
                        i0 + 1, i1, ex_lote,
                    ),
                )
            if lote_ok:
                _after_tag_batch()

        if mh_keys:
            _cab_tags.etiquetar_cabezal_confinamiento_multihost_animado(
                doc,
                view,
                res,
                batch_size=batch_tags,
                after_batch=_after_tag_batch,
            )
    except Exception as ex_tag:
        res[u"messages"].append(u"Etiquetado confinamiento animado: {0}".format(ex_tag))
    finally:
        _cabezal_pbar_exit(pb_tags, pbar_tags_open)
    if uidoc is not None and not MODO_EJECUCION_RAPIDA:
        _cabezal_refrescar_vista_fin_flujo(doc, uidoc)


def _cabezal_aplicar_etiquetado_longitudinal_animado(doc, view, res, uidoc):
    """Etiqueta longitudinales lote a lote (mismo orden que creación de barras)."""
    if _cab_tags is None or doc is None or view is None:
        return
    long_ids = list(res.get(u"rebars_longitudinales_ids") or [])
    tag_meta = list(res.get(u"rebars_longitudinales_tag_meta") or [])
    if not long_ids:
        return
    batch_tags = _tamano_lote_ejecucion(len(long_ids), CABEZAL_TAGS_POR_LOTE_ANIMACION)
    n_tag_lotes = int(math.ceil(float(len(long_ids)) / float(batch_tags)))
    pb_tags = _cabezal_pbar_start(
        _cabezal_pbar_phase_title(_CAB_PBAR_BASE_TAGS, n_tag_lotes),
        n_tag_lotes,
        doc=doc,
    )
    pbar_tags_open = _cabezal_pbar_enter(pb_tags)
    tag_lote_idx = [0]

    def _after_tag_batch():
        _cabezal_refrescar_vista_tras_lote(doc, uidoc)
        _cabezal_pbar_step(
            pb_tags, tag_lote_idx[0], n_tag_lotes, _CAB_PBAR_BASE_TAGS,
        )
        tag_lote_idx[0] += 1

    try:
        tag_res = _cab_tags.etiquetar_cabezal_longitudinales_en_vista_animado(
            doc,
            view,
            long_ids,
            tag_meta=tag_meta,
            batch_size=batch_tags,
            after_batch=_after_tag_batch,
        )
        res[u"n_tags_created"] = int(res.get(u"n_tags_created", 0)) + int(
            tag_res.get(u"n_ok", 0),
        )
        res[u"n_tags_fail"] = int(res.get(u"n_tags_fail", 0)) + int(
            tag_res.get(u"n_fail", 0),
        )
        for msg in (tag_res.get(u"messages") or [])[:8]:
            res[u"messages"].append(msg)
    except Exception as ex_tag:
        res[u"messages"].append(u"Etiquetado animado: {0}".format(ex_tag))
    finally:
        _cabezal_pbar_exit(pb_tags, pbar_tags_open)
    if uidoc is not None and not MODO_EJECUCION_RAPIDA:
        _cabezal_refrescar_vista_fin_flujo(doc, uidoc)


def _cabezal_aplicar_creacion_animada(
    doc,
    walls,
    seg_jobs_all,
    confinement_jobs,
    curve_tol,
    res,
    rebars_por,
    bar_type_fallback=None,
    uidoc=None,
    defer_etiquetado=False,
    within_parent_transaction_group=False,
):
    """
    1) Barras longitudinales lote a lote (abajo→arriba).
    2) Confinamiento lote a lote (sin etiquetas).
    3) Etiquetado animado longitudinales.
    4) Etiquetado animado confinamiento.

    Cada lote = transacción + commit + refresco de vista.
    """
    pb_bars = None
    pb_conf = None
    pbar_bars_open = False
    pbar_conf_open = False
    creacion_ok = False
    use_own_tg = use_transaction_group_armado_muros(
        doc, within_parent_transaction_group=within_parent_transaction_group,
    )
    tg = None
    tg_started = False

    try:
        seg_sorted = _cabezal_sort_seg_jobs(seg_jobs_all, walls)
        conf_sorted = _cabezal_sort_confinement_jobs(confinement_jobs, walls)
        batch_bars = _tamano_lote_ejecucion(
            len(seg_sorted), CABEZAL_BARRAS_POR_LOTE_ANIMACION,
        )
        batch_conf = _tamano_lote_ejecucion(
            len(conf_sorted), CABEZAL_CONFINAMIENTO_POR_LOTE_ANIMACION,
        )
        n_bar_lotes = (
            int(math.ceil(float(len(seg_sorted)) / float(batch_bars)))
            if seg_sorted else 0
        )
        n_conf_lotes = (
            int(math.ceil(float(len(conf_sorted)) / float(batch_conf)))
            if conf_sorted else 0
        )

        pb_bars = _cabezal_pbar_start(
            _cabezal_pbar_phase_title(_CAB_PBAR_BASE_BARS, n_bar_lotes),
            n_bar_lotes,
            doc=doc,
        )
        pb_conf = _cabezal_pbar_start(
            _cabezal_pbar_phase_title(_CAB_PBAR_BASE_CONF, n_conf_lotes),
            n_conf_lotes,
            doc=doc,
        )
        pbar_bars_open = _cabezal_pbar_enter(pb_bars)

        if use_own_tg:
            tg = TransactionGroup(doc, u"Arainco: Cabezal muros")
            tg.Start()
            tg_started = True
        bar_lote_idx = 0

        for i0 in range(0, len(seg_sorted), batch_bars):
            lote = seg_sorted[i0:i0 + batch_bars]
            i1 = min(i0 + batch_bars, len(seg_sorted))
            _cabezal_pbar_step(pb_bars, bar_lote_idx, n_bar_lotes, _CAB_PBAR_BASE_BARS)
            bar_lote_idx += 1

            if len(lote) == 1:
                sj0 = lote[0]
                txn_name = u"Arainco: Cabezal muros — barra capa {0} Z≈{1:.2f} m".format(
                    int(sj0.get(u"layer_index", 0)) + 1,
                    _cabezal_z_ft_to_m(sj0.get(u"zs", 0.0)),
                )
            else:
                txn_name = u"Arainco: Cabezal muros — barras {0}–{1} de {2}".format(
                    i0 + 1, i1, len(seg_sorted),
                )

            t = Transaction(doc, txn_name)
            t.Start()
            lote_ok = False
            try:
                for sj in lote:
                    _cabezal_create_longitudinal_job(
                        doc, sj, curve_tol, res, rebars_por,
                    )
                t.Commit()
                lote_ok = True
            except Exception as ex_lote:
                try:
                    if t.HasStarted():
                        t.RollBack()
                except Exception:
                    pass
                res[u"messages"].append(
                    u"Cabezal barras lote {0}–{1}: {2}".format(
                        i0 + 1, i1, str(ex_lote),
                    ),
                )
                res[u"n_fail"] = int(res.get(u"n_fail", 0)) + 1
            if lote_ok:
                if activar_armadura_arainco_por_ids is not None:
                    try:
                        lote_ids = []
                        for sj in lote:
                            rid = sj.get(u"rebar_id")
                            if rid is not None:
                                lote_ids.append(rid)
                        activar_armadura_arainco_por_ids(doc, lote_ids)
                    except Exception:
                        pass
                _cabezal_refrescar_vista_tras_lote(doc, uidoc)

        _cabezal_pbar_exit(pb_bars, pbar_bars_open)
        pbar_bars_open = False

        if conf_sorted:
            pbar_conf_open = _cabezal_pbar_enter(pb_conf)
            conf_lote_idx = 0
            for i0 in range(0, len(conf_sorted), batch_conf):
                lote = conf_sorted[i0:i0 + batch_conf]
                i1 = min(i0 + batch_conf, len(conf_sorted))
                _cabezal_pbar_step(
                    pb_conf, conf_lote_idx, n_conf_lotes, _CAB_PBAR_BASE_CONF,
                )
                conf_lote_idx += 1

                if len(lote) == 1:
                    cj0 = lote[0]
                    kind = (
                        u"traba"
                        if cj0.get(u"job_kind") == u"tie"
                        else u"estribo"
                    )
                    txn_name = u"Arainco: Cabezal muros — {0} muro {1} {2}".format(
                        kind,
                        cj0.get(u"wid"),
                        cj0.get(u"extremo") or u"",
                    )
                else:
                    txn_name = u"Arainco: Cabezal muros — confinamiento {0}–{1} de {2}".format(
                        i0 + 1, i1, len(conf_sorted),
                    )

                t = Transaction(doc, txn_name)
                t.Start()
                lote_ok = False
                try:
                    for cj in lote:
                        _cabezal_create_confinement_job(
                            doc, cj, res, rebars_por, bar_type_fallback,
                            defer_tags=True,
                        )
                    t.Commit()
                    lote_ok = True
                except Exception as ex_lote:
                    try:
                        if t.HasStarted():
                            t.RollBack()
                    except Exception:
                        pass
                    res[u"messages"].append(
                        u"Cabezal confinamiento lote {0}–{1}: {2}".format(
                            i0 + 1, i1, str(ex_lote),
                        ),
                    )
                    res[u"n_fail"] = int(res.get(u"n_fail", 0)) + 1
                if lote_ok:
                    _cabezal_refrescar_vista_tras_lote(doc, uidoc)

        if defer_etiquetado:
            res[u"_defer_etiquetado"] = True
            res[u"_seg_jobs_all"] = seg_jobs_all
        else:
            tag_view = None
            if uidoc is not None:
                try:
                    tag_view = uidoc.ActiveView
                except Exception:
                    tag_view = None
            if tag_view is not None:
                _cabezal_aplicar_etiquetado_longitudinal_animado(
                    doc, tag_view, res, uidoc,
                )
                _cabezal_aplicar_etiquetado_confinamiento_animado(
                    doc, tag_view, res, uidoc,
                )

        n_created = int(res.get(u"n_created", 0))
        n_conf = int(res.get(u"n_confinement_created", 0))
        if within_parent_transaction_group:
            creacion_ok = True
        elif n_created > 0 or n_conf > 0:
            creacion_ok = True
        else:
            res[u"n_skip"] = int(res.get(u"n_skip", 0)) + 1
    except Exception as ex:
        res[u"n_fail"] = int(res.get(u"n_fail", 0)) + 1
        res[u"messages"].append(u"Cabezal pipeline: {0}".format(str(ex)))
    finally:
        _cabezal_pbar_exit(pb_bars, pbar_bars_open)
        _cabezal_pbar_exit(pb_conf, pbar_conf_open)
        if use_own_tg and tg_started and tg is not None:
            try:
                if creacion_ok:
                    tg.Assimilate()
                else:
                    tg.RollBack()
            except Exception:
                try:
                    if tg.HasStarted():
                        tg.RollBack()
                except Exception:
                    pass

    if creacion_ok and uidoc is not None and not defer_etiquetado:
        _cabezal_refrescar_vista_fin_flujo(doc, uidoc)

    return creacion_ok


def _cabezal_tag_view(uidoc):
    if uidoc is None:
        return None
    try:
        return uidoc.ActiveView
    except Exception:
        return None


def cabezal_aplicar_etiquetado_longitudinal_pendiente(doc, res, uidoc=None):
    """Etiquetas de barras longitudinales (+ marcadores empalme) tras mallas en flujo unificado."""
    if not res or not res.get(u"_defer_etiquetado"):
        return res
    if res.get(u"_defer_long_tags_applied"):
        return res
    tag_view = _cabezal_tag_view(uidoc)
    if tag_view is not None:
        try:
            _cabezal_aplicar_etiquetado_longitudinal_animado(
                doc, tag_view, res, uidoc,
            )
        except Exception as ex_tag:
            res.setdefault(u"messages", []).append(
                u"Etiquetado longitudinales: {0}".format(ex_tag),
            )
    seg_jobs = res.get(u"_seg_jobs_all") or []
    if colocar_marcadores_empalme_cabezal is not None and tag_view is not None and seg_jobs:
        try:
            emp_res = colocar_marcadores_empalme_cabezal(
                doc, tag_view, seg_jobs,
            )
            res[u"n_empalme_markers_ok"] = int(emp_res.get(u"n_ok", 0))
            res[u"n_empalme_markers_fail"] = int(emp_res.get(u"n_fail", 0))
            res[u"n_empalme_dims_ok"] = int(emp_res.get(u"n_dims_ok", 0))
            res[u"n_empalme_dims_fail"] = int(emp_res.get(u"n_dims_fail", 0))
            for m in emp_res.get(u"messages") or []:
                if m:
                    res.setdefault(u"messages", []).append(m)
        except Exception as ex_emp:
            res.setdefault(u"messages", []).append(
                u"Marcadores empalme: {0}".format(ex_emp),
            )
    res[u"_defer_long_tags_applied"] = True
    return res


def cabezal_aplicar_etiquetado_confinamiento_pendiente(doc, res, uidoc=None):
    """Etiquetas de confinamiento (tras etiquetas longitudinales y mallas en flujo unificado)."""
    if not res or not res.get(u"_defer_etiquetado"):
        return res
    tag_view = _cabezal_tag_view(uidoc)
    if tag_view is not None:
        try:
            _cabezal_aplicar_etiquetado_confinamiento_animado(
                doc, tag_view, res, uidoc,
            )
        except Exception as ex_tag:
            res.setdefault(u"messages", []).append(
                u"Etiquetado confinamiento: {0}".format(ex_tag),
            )
    res[u"_defer_etiquetado"] = False
    return res


def cabezal_aplicar_etiquetado_pendiente(doc, res, uidoc=None):
    """Fase etiquetas cabezal completa (longitudinales + confinamiento)."""
    cabezal_aplicar_etiquetado_longitudinal_pendiente(doc, res, uidoc=uidoc)
    return cabezal_aplicar_etiquetado_confinamiento_pendiente(doc, res, uidoc=uidoc)


def aplicar_cabezales_muros(
    doc, walls, cabezal_por_muro_id,
    bar_type_fallback=None,
    ref_walls_troceo=None,
    uidoc=None,
    defer_etiquetado=False,
    within_parent_transaction_group=False,
):
    """
    Pipeline cabezal muros.
    Paso 1: generar líneas por muro/extremo/capa.
    Creación animada: primero todas las barras (longitudinales, luego confinamiento);
    después etiquetado animado de longitudinales y confinamiento, con transacción +
    refresco de vista por lote.

    ``uidoc``: opcional; refresca la vista activa tras cada lote de barras y etiquetas.
    """
    res = {
        u"n_created": 0,
        u"n_bars_total": 0,
        u"n_skip": 0,
        u"n_fail": 0,
        u"n_troceo_segments": 0,
        u"n_confinement_created": 0,
        u"n_conf_tags_created": 0,
        u"n_conf_tags_fail": 0,
        u"n_tags_created": 0,
        u"n_tags_fail": 0,
        u"n_empalme_markers_ok": 0,
        u"n_empalme_markers_fail": 0,
        u"n_empalme_dims_ok": 0,
        u"n_empalme_dims_fail": 0,
        u"messages": [],
        u"rebars_por_muro_id": {},
        u"rebars_longitudinales_ids": [],
        u"rebars_longitudinales_tag_meta": [],
    }
    if doc is None or not walls or not cabezal_por_muro_id:
        return res

    cabezal_por_muro_id = normalize_muro_id_dict(cabezal_por_muro_id)
    rebars_por = res[u"rebars_por_muro_id"]

    try:
        return _aplicar_cabezales_muros_pipeline(
            doc,
            walls,
            cabezal_por_muro_id,
            bar_type_fallback,
            ref_walls_troceo,
            uidoc,
            defer_etiquetado,
            within_parent_transaction_group,
            res,
            rebars_por,
        )
    except Exception as ex:
        res[u"n_fail"] = int(res.get(u"n_fail", 0)) + 1
        res[u"messages"].append(u"Pipeline cabezal: {0}".format(ex))
        return res


def _aplicar_cabezales_muros_pipeline(
    doc,
    walls,
    cabezal_por_muro_id,
    bar_type_fallback,
    ref_walls_troceo,
    uidoc,
    defer_etiquetado,
    within_parent_transaction_group,
    res,
    rebars_por,
):
    max_layers = CABEZAL_MIN_CAPAS
    for _wid, _cfg in (cabezal_por_muro_id or {}).items():
        for _ex in CABEZAL_EXTREMOS:
            _ex_cfg = (_cfg or {}).get(_ex) or {}
            max_layers = max(max_layers, len(_ex_cfg.get(u"layers") or []))

    segment_ctx_by_extremo = {}
    for extremo in CABEZAL_EXTREMOS:
        segment_ctx_by_extremo[extremo] = resolve_segment_context_for_extremo(
            walls, cabezal_por_muro_id, extremo, max_layers, bar_type_fallback,
        )

    # ── Paso 1: generar líneas por muro/extremo/capa ──────────────────────
    all_line_jobs = []
    try:
        for wall in walls:
            if wall is None or not isinstance(wall, Wall):
                continue
            wid = wall_id_int(wall)
            cfg = cabezal_por_muro_id.get(wid)
            if not cfg:
                continue
            ok_val, msg_val = validar_cabezal_config(cfg)
            if not ok_val:
                res[u"n_fail"] += 1
                res[u"messages"].append(u"Muro {0}: {1}".format(wid, msg_val))
                continue

            stack_idx = _wall_stack_index(walls, wall)

            for extremo in CABEZAL_EXTREMOS:
                ex_cfg = cfg.get(extremo) or {}
                if not cabezal_extremo_armado_activo(ex_cfg):
                    continue
                _normalize_cabezal_extremo_layers(ex_cfg)
                cabezal_sync_confinement_from_extremo(
                    ex_cfg,
                    doc,
                    cabezal_sync_fallback_bar_type_id(ex_cfg, bar_type_fallback),
                )
                layers = cabezal_active_layers(ex_cfg)
                segment_ctx = segment_ctx_by_extremo.get(extremo)
                if not layers:
                    res[u"n_fail"] += 1
                    res[u"messages"].append(
                        u"Muro {0} {1}: sin capas de cabezal.".format(wid, extremo),
                    )
                    continue
                layer_spacing_mm = float(
                    ex_cfg.get(u"layer_spacing_mm") or CABEZAL_LAYER_PITCH_MM,
                )
                geom = _wall_longitudinal_at_extremo(wall, extremo)
                normal_muro = geom[u"normal_muro"] if geom else XYZ.BasisY
                vec_long = geom[u"vector_longitudinal"] if geom else XYZ.BasisX

                for layer_index, ly in enumerate(layers):
                    try:
                        n_bars = int(ly.get(u"n_bars", CABEZAL_MIN_BARRAS_POR_CAPA))
                    except Exception:
                        n_bars = CABEZAL_MIN_BARRAS_POR_CAPA
                    n_bars = max(
                        CABEZAL_MIN_BARRAS_POR_CAPA,
                        min(CABEZAL_MAX_BARRAS_POR_CAPA, n_bars),
                    )
                    bar_type = _resolver_bar_type_for_layer(
                        doc, ex_cfg, ly, bar_type_fallback,
                        segment_ctx=segment_ctx,
                        stack_index=stack_idx,
                        layer_index=layer_index,
                    )
                    if bar_type is None:
                        res[u"n_fail"] += 1
                        res[u"messages"].append(
                            u"Muro {0} {1} capa {2}: sin RebarBarType.".format(
                                wid, extremo, layer_index + 1,
                            ),
                        )
                        continue
                    conf_type = _resolver_conf_bar_type(
                        doc, ex_cfg, bar_type, bar_type_fallback,
                    )
                    if conf_type is None:
                        conf_type = bar_type
                    p_lo, p_hi, distrib_ft, err_geom = _cabezal_capa_line_endpoints(
                        wall, extremo, layer_index, bar_type, conf_type,
                        layer_spacing_mm=layer_spacing_mm,
                        doc=doc, ex_cfg=ex_cfg,
                    )
                    if err_geom:
                        res[u"n_fail"] += 1
                        res[u"messages"].append(
                            u"Muro {0} {1} capa {2}: {3}".format(
                                wid, extremo, layer_index + 1, err_geom,
                            ),
                        )
                        continue
                    all_line_jobs.append({
                        u"p_lo": p_lo,
                        u"p_hi": p_hi,
                        u"distrib_ft": distrib_ft,
                        u"n_bars": n_bars,
                        u"bar_type": bar_type,
                        u"conf_type": conf_type,
                        u"normal_muro": normal_muro,
                        u"vec_long": vec_long,
                        u"wall": wall,
                        u"wid": wid,
                        u"layer_index": layer_index,
                        u"extremo": extremo,
                        u"thickness_mm": _wall_thickness_mm_for_fusion(wall),
                        u"troceo_por_muro": _troceo_por_muro_from_extremo_cfg(ex_cfg),
                    })
    except Exception as ex_lines:
        raise Exception(u"Líneas cabezal: {0}".format(ex_lines))

    if not all_line_jobs:
        return res

    # ── Paso 2: fundación — estirar p_lo por lote de muro ────────────────
    fund_stretch_ft_by_wid = {}
    try:
        for wall in walls:
            if wall is None or not isinstance(wall, Wall):
                continue
            wid = wall_id_int(wall)
            foundations = _fundaciones_unidas_muro(doc, wall)
            if not foundations:
                continue
            best_h_mm = 0.0
            for fund in foundations:
                h_mm = _altura_bbox_elemento_mm(fund)
                if h_mm is not None and h_mm > best_h_mm:
                    best_h_mm = h_mm
            stretch_mm = max(0.0, best_h_mm - float(FOUNDATION_STRETCH_RESTA_MM))
            if stretch_mm > 0.1:
                fund_stretch_ft_by_wid[wid] = _mm_to_internal(stretch_mm)
    except Exception as ex_fund:
        raise Exception(u"Fundación cabezal: {0}".format(ex_fund))

    geom_opts = _geometry_options()

    for job in all_line_jobs:
        dz_fund = fund_stretch_ft_by_wid.get(job[u"wid"], 0.0)
        if dz_fund > 1e-12:
            old_lo = job[u"p_lo"]
            job[u"p_lo"] = XYZ(float(old_lo.X), float(old_lo.Y),
                               float(old_lo.Z) - dz_fund)

    # ── Paso 3: fusión colineal ──────────────────────────────────────────
    try:
        fused_lines = _fuse_colinear_cabezal_lines(all_line_jobs)
    except Exception as ex_fuse:
        raise Exception(u"Fusión cabezal: {0}".format(ex_fuse))

    if not fused_lines:
        return res

    try:
        curve_tol = doc.Application.ShortCurveTolerance
    except Exception:
        curve_tol = 1e-6

    # ── Paso 4: colisión embed +Z (cabeza) ────────────────────────────────
    try:
        for fj in fused_lines:
            bx = float(fj[u"p_lo"].X)
            by = float(fj[u"p_lo"].Y)
            z_top = float(fj[u"p_hi"].Z)
            extremo = fj.get(u"extremo") or CABEZAL_EXTREMO_INICIO
            layer_index = int(fj.get(u"layer_index", 0))
            segment_ctx = segment_ctx_by_extremo.get(extremo)
            stack_idx_top = _stack_index_for_z(walls, z_top - curve_tol, curve_tol)
            bar_type_top = _bar_type_for_cabezal_stack(
                doc, walls, cabezal_por_muro_id, segment_ctx, extremo,
                stack_idx_top, layer_index, fj[u"bar_type"],
            )
            if bar_type_top is None:
                bar_type_top = fj[u"bar_type"]
            d_mm = _bar_diameter_mm(bar_type_top)
            host_wall = fj[u"wall"]

            embed_mm = _empotramiento_tabla_mm(d_mm) or 0.0
            dz_embed_ft = _mm_to_internal(embed_mm) if embed_mm > 0.1 else 0.0
            retract_mm = _retract_mm_sin_colision(d_mm)
            dz_retract_ft = _mm_to_internal(retract_mm) if retract_mm > 0.1 else 0.0

            if dz_embed_ft > 1e-12:
                collides = _embed_collides_wall_solids_upward(
                    doc, bx, by, z_top, dz_embed_ft, d_mm,
                    walls, host_wall.Id, geom_opts,
                )
                if collides:
                    old_hi = fj[u"p_hi"]
                    fj[u"p_hi"] = XYZ(float(old_hi.X), float(old_hi.Y),
                                       float(old_hi.Z) + dz_embed_ft)
                    fj[u"_reverted_top"] = False
                else:
                    old_hi = fj[u"p_hi"]
                    fj[u"p_hi"] = XYZ(float(old_hi.X), float(old_hi.Y),
                                       float(old_hi.Z) - dz_retract_ft)
                    fj[u"_reverted_top"] = True
            else:
                fj[u"_reverted_top"] = False
    except Exception as ex_embed:
        raise Exception(u"Embed cabezal: {0}".format(ex_embed))

    # ── Paso 5: troceo + paso 6: empalme ────────────────────────────────
    legacy_ref_walls = list(ref_walls_troceo or [])

    seg_jobs_all = []
    try:
        for fj_idx, fj in enumerate(fused_lines):
            try:
                new_jobs, n_seg, ref_walls_job = _troceo_planificar_seg_jobs_from_fused_line(
                    doc,
                    walls,
                    cabezal_por_muro_id,
                    segment_ctx_by_extremo.get(
                        fj.get(u"extremo") or CABEZAL_EXTREMO_INICIO,
                    ),
                    fj,
                    fj_idx,
                    fund_stretch_ft_by_wid,
                    legacy_ref_walls,
                    curve_tol,
                    geom_opts,
                )
                if n_seg > 1 and ref_walls_job:
                    res[u"n_troceo_segments"] = int(
                        res.get(u"n_troceo_segments", 0),
                    ) + int(n_seg)
                seg_jobs_all.extend(new_jobs)
            except Exception as ex_fj:
                try:
                    wid_fj = fj.get(u"wid")
                except Exception:
                    wid_fj = u"?"
                try:
                    capa_fj = int(fj.get(u"layer_index", 0)) + 1
                except Exception:
                    capa_fj = u"?"
                ex_fj_s = str(ex_fj)
                try:
                    extremo_fj = fj.get(u"extremo") or CABEZAL_EXTREMO_INICIO
                except Exception:
                    extremo_fj = u"?"
                res.setdefault(u"messages", []).append(
                    u"Troceo cabezal muro {0} capa {1} {2}: {3}".format(
                        wid_fj, capa_fj, extremo_fj, ex_fj_s,
                    ),
                )
    except Exception as ex_troceo:
        raise Exception(u"Troceo cabezal: {0}".format(ex_troceo))

    if not seg_jobs_all:
        return res

    confinement_jobs = []
    try:
        for wall in walls:
            if wall is None or not isinstance(wall, Wall):
                continue
            wid = wall_id_int(wall)
            cfg = cabezal_por_muro_id.get(wid)
            if not cfg:
                continue
            stack_idx = _wall_stack_index(walls, wall)
            for extremo in CABEZAL_EXTREMOS:
                ex_cfg = cfg.get(extremo) or {}
                if not cabezal_extremo_armado_activo(ex_cfg):
                    continue
                _normalize_cabezal_extremo_layers(ex_cfg)
                cabezal_sync_confinement_from_extremo(
                    ex_cfg,
                    doc,
                    cabezal_sync_fallback_bar_type_id(ex_cfg, bar_type_fallback),
                )
                ext_line_jobs = [
                    j for j in all_line_jobs
                    if j.get(u"wid") == wid and j.get(u"extremo") == extremo
                ]
                confinement_jobs.extend(
                    cabezal_build_confinement_jobs(
                        wall, wid, extremo, ex_cfg, stack_idx,
                        segment_ctx_by_extremo.get(extremo),
                        line_jobs=ext_line_jobs,
                    ),
                )
    except Exception as ex_conf:
        raise Exception(u"Jobs confinamiento: {0}".format(ex_conf))

    # ── Resolver host wall por punto medio ────────────────────────────────
    try:
        for sj in seg_jobs_all:
            z_mid = sj[u"zs"] + sj[u"span_seg"] * 0.5
            host = _resolve_host_wall(
                sj[u"bx"], sj[u"by"], z_mid,
                walls, geom_opts, sj[u"wall"],
            )
            sj[u"wall"] = host
    except Exception as ex_host:
        raise Exception(u"Host cabezal: {0}".format(ex_host))

    # ── Crear rebars (barras → confinamiento, lote a lote abajo→arriba) ───
    try:
        _cabezal_aplicar_creacion_animada(
            doc,
            walls,
            seg_jobs_all,
            confinement_jobs,
            curve_tol,
            res,
            rebars_por,
            bar_type_fallback=bar_type_fallback,
            uidoc=uidoc,
            defer_etiquetado=defer_etiquetado,
            within_parent_transaction_group=within_parent_transaction_group,
        )
    except Exception as ex_crea:
        res[u"n_fail"] = int(res.get(u"n_fail", 0)) + 1
        res[u"messages"].append(u"Creación cabezal: {0}".format(ex_crea))
        if defer_etiquetado:
            res[u"_seg_jobs_all"] = seg_jobs_all
        return res

    if defer_etiquetado:
        res[u"_seg_jobs_all"] = seg_jobs_all
        return res

    if colocar_marcadores_empalme_cabezal is not None:
        active_view = None
        if uidoc is not None:
            try:
                active_view = uidoc.ActiveView
            except Exception:
                active_view = None
        if active_view is not None:
            emp_res = colocar_marcadores_empalme_cabezal(
                doc, active_view, seg_jobs_all,
            )
            res[u"n_empalme_markers_ok"] = int(emp_res.get(u"n_ok", 0))
            res[u"n_empalme_markers_fail"] = int(emp_res.get(u"n_fail", 0))
            res[u"n_empalme_dims_ok"] = int(emp_res.get(u"n_dims_ok", 0))
            res[u"n_empalme_dims_fail"] = int(emp_res.get(u"n_dims_fail", 0))
            for m in emp_res.get(u"messages") or []:
                if m:
                    res[u"messages"].append(m)

    return res


def aplicar_post_proceso_cabezal_muros(doc, walls, rebars_por_muro_id):
    """
    Retrocompatibilidad: el post-proceso ahora es interno al pipeline.
    Esta función es un no-op que retorna contadores vacíos.
    """
    return {
        u"n_extended": 0,
        u"n_retracted": 0,
        u"n_pata_l": 0,
        u"n_pata_l_fund_pie": 0,
        u"n_fundacion_pie": 0,
        u"n_fundacion_retract": 0,
        u"n_skip": 0,
        u"n_fail": 0,
        u"messages": [],
    }
