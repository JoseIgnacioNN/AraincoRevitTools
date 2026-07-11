# -*- coding: utf-8 -*-
"""
Etiquetado Structural Rebar Tag para barras longitudinales de cabezal de muro.

Una ``IndependentTag`` por cada ``Rebar`` creado. Orientación horizontal, leader
activo. Columna alineada hacia fuera del muro; cada etiqueta conserva la cota del
tronco vertical (leaders horizontales) y se separa si hay riesgo de solape.
Distancias según escala de la vista activa (1/50, 1/75, 1/100, …).
"""

from __future__ import print_function

import clr
import math
import os
import re
import sys

clr.AddReference("RevitAPI")

from System.Collections.Generic import List

_d = os.path.dirname(os.path.abspath(__file__))
if _d not in sys.path:
    sys.path.insert(0, _d)
try:
    import bootstrap_paths
    bootstrap_paths.pin_local_scripts_first()
except Exception:
    pass

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilySymbol,
    FilteredElementCollector,
    IndependentTag,
    Options,
    Reference,
    StorageType,
    TagMode,
    TagOrientation,
    SubTransaction,
    Transaction,
    UnitTypeId,
    UnitUtils,
    View3D,
    Wall,
    XYZ,
)
from Autodesk.Revit.DB.Structure import Rebar

try:
    from armado_muros_lineales import location_curve_wall
except Exception:
    location_curve_wall = None

try:
    from enfierrado_shaft_hashtag import _rebar_reference_candidates_for_tag
except Exception:
    _rebar_reference_candidates_for_tag = None

try:
    from Autodesk.Revit.DB.Structure import MultiPlanarOption
except Exception:
    MultiPlanarOption = None

CABEZAL_REBAR_TAG_FAMILY_NAME = u"EST_A_STRUCTURAL REBAR TAG_WALL_HORIZONTAL"
CABEZAL_CONFINEMENT_TAG_FAMILY_NAME = u"EST_A_STRUCTURAL REBAR TAG_CONFINAMIENTO"
# Familia de etiqueta multihost (Add / Remove Host) para trabas (3+ capas Tipo 1 / Tipo 2).
CABEZAL_CONFINEMENT_MULTIHOST_TAG_FAMILY_NAME = (
    u"EST_A_STRUCTURAL REBAR TAG_CONFINAMIENTO_MULTI HOST"
)
# Desplazamiento genérico (2 capas, 4+ capas) fuera del muro @ 1/50.
CABEZAL_CONFINEMENT_TAG_OFFSET_MM_AT_REF_SCALE = 250.0
# Claves de espaciado @ 3 capas (mm modelo en vista de referencia 1/50).
CABEZAL_CONF_3C_TIPO1_MULTIHOST = u"tipo1_multihost_mm"
CABEZAL_CONF_3C_TIPO2_ESTRIBO = u"tipo2_estribo_mm"
CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA = u"tipo2_traba_extra_mm"
CABEZAL_CONF_3C_TIPO2_TRABA_LATERAL = u"tipo2_traba_lateral_mm"
# 3 capas — calibrado @ 1/50 (legibilidad en alzado; evita solape con armadura/entre etiquetas).
CABEZAL_CONFINEMENT_TAG_3CAPAS_SPACING_BY_SCALE = {
    50: {
        CABEZAL_CONF_3C_TIPO1_MULTIHOST: 340.0,
        CABEZAL_CONF_3C_TIPO2_ESTRIBO: 280.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA: 220.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_LATERAL: 0.0,
    },
    75: {
        CABEZAL_CONF_3C_TIPO1_MULTIHOST: 510.0,
        CABEZAL_CONF_3C_TIPO2_ESTRIBO: 420.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA: 330.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_LATERAL: 0.0,
    },
    100: {
        CABEZAL_CONF_3C_TIPO1_MULTIHOST: 680.0,
        CABEZAL_CONF_3C_TIPO2_ESTRIBO: 560.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA: 440.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_LATERAL: 0.0,
    },
}
# 5 capas @ 1/50 — Tipo 1: multihost [1,2,3,4]; Tipo 2: E. + multihost [1,2,3].
CABEZAL_CONFINEMENT_TAG_5CAPAS_SPACING_BY_SCALE = {
    50: {
        CABEZAL_CONF_3C_TIPO1_MULTIHOST: 160.0,
        CABEZAL_CONF_3C_TIPO2_ESTRIBO: 180.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA: 200.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_LATERAL: 0.0,
    },
    75: {
        CABEZAL_CONF_3C_TIPO1_MULTIHOST: 240.0,
        CABEZAL_CONF_3C_TIPO2_ESTRIBO: 270.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA: 300.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_LATERAL: 0.0,
    },
    100: {
        CABEZAL_CONF_3C_TIPO1_MULTIHOST: 320.0,
        CABEZAL_CONF_3C_TIPO2_ESTRIBO: 360.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA: 400.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_LATERAL: 0.0,
    },
}
# 4 capas @ 1/50 — Tipo 1: multihost [1,2,3]; Tipo 2: E. + multihost [1,2].
# Holgura desde cara exterior (tipo1 / estribo); tipo2_traba_extra = separación entre etiquetas.
CABEZAL_CONFINEMENT_TAG_4CAPAS_SPACING_BY_SCALE = {
    50: {
        CABEZAL_CONF_3C_TIPO1_MULTIHOST: 160.0,
        CABEZAL_CONF_3C_TIPO2_ESTRIBO: 180.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA: 200.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_LATERAL: 0.0,
    },
    75: {
        CABEZAL_CONF_3C_TIPO1_MULTIHOST: 240.0,
        CABEZAL_CONF_3C_TIPO2_ESTRIBO: 270.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA: 300.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_LATERAL: 0.0,
    },
    100: {
        CABEZAL_CONF_3C_TIPO1_MULTIHOST: 320.0,
        CABEZAL_CONF_3C_TIPO2_ESTRIBO: 360.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA: 400.0,
        CABEZAL_CONF_3C_TIPO2_TRABA_LATERAL: 0.0,
    },
}
# Referencia calibrada en escala 1/50 (View.Scale == 50).
CABEZAL_TAG_SCALE_REFERENCE = 50
CABEZAL_TAG_OFFSET_MM = 1000.0
CABEZAL_TAG_LAYER_STEP_MM = 800.0
# Escenarios explícitos: denominador de escala → (offset mm, paso vertical mm).
# Fórmula base: valor_ref × (escala / CABEZAL_TAG_SCALE_REFERENCE).
CABEZAL_TAG_SPACING_BY_SCALE = {
    50: (1000.0, 800.0),
    75: (1500.0, 1200.0),
    100: (2000.0, 1600.0),
}
# Coronamiento (barra horizontal en alzado): offset en vertical de vista, no en espesor.
CORONAMIENTO_TAG_SCALE_REFERENCE = 50
CORONAMIENTO_TAG_OFFSET_MM = 750.0
CORONAMIENTO_TAG_LAYER_STEP_MM = 500.0
CORONAMIENTO_TAG_SPACING_BY_SCALE = {
    50: (750.0, 500.0),
    75: (1125.0, 750.0),
    100: (1500.0, 1000.0),
}
CABEZAL_EXTREMO_INICIO = u"inicio"
CABEZAL_EXTREMO_FIN = u"fin"
CABEZAL_TAG_SHAPE_FALLBACK = u"01"
# Barras cabezal sin RebarShape: tramos de centerline → código de etiqueta ARAINCO.
CABEZAL_TAG_SHAPE_BY_SEGMENTS = {
    1: u"01",
    2: u"02",
    3: u"03",
    4: u"04",
}
_MAX_TAG_WARNINGS = 12


def _norm_key(s):
    if s is None:
        return u""
    try:
        t = unicode(s)
    except Exception:
        try:
            t = str(s)
        except Exception:
            t = u""
    return u" ".join(t.replace(u"\u00A0", u" ").split()).lower()


def _norm_alnum_key(s):
    try:
        t = unicode(s or u"")
    except Exception:
        t = str(s or u"")
    return re.sub(r"[^0-9a-z]+", u"", t.lower())


def _label_lookup_keys(raw):
    """Variantes de clave para cruzar RebarShape / tipos «01»…«13»."""
    keys = []
    seen = set()
    if raw is None:
        return keys

    def _add(k):
        if k is None:
            return
        try:
            ks = unicode(k)
        except Exception:
            ks = str(k)
        if not ks or ks in seen:
            return
        seen.add(ks)
        keys.append(ks)

    try:
        s = unicode(raw).strip()
    except Exception:
        s = str(raw or u"").strip()
    if not s:
        return keys

    _add(s)
    _add(_norm_key(s))
    _add(_norm_alnum_key(s))

    digits = re.sub(r"\D", u"", s)
    if digits:
        try:
            n = int(digits)
            _add(u"{0:02d}".format(n))
            _add(str(n))
            _add(_norm_key(u"{0:02d}".format(n)))
        except Exception:
            pass
    return keys


def _vista_permite_rebar_tags(view):
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        if isinstance(view, View3D):
            return False
    except Exception:
        pass
    return True


def _family_symbols_rebar_tag(document, family_name):
    if document is None or not family_name:
        return []
    tgt = _norm_key(family_name)
    out = []
    try:
        col = (
            FilteredElementCollector(document)
            .OfClass(FamilySymbol)
            .OfCategory(BuiltInCategory.OST_RebarTags)
        )
        for sym in col:
            if sym is None:
                continue
            fn = u""
            try:
                fn = sym.FamilyName
            except Exception:
                pass
            if not fn:
                try:
                    fam = sym.Family
                    if fam is not None:
                        fn = fam.Name
                except Exception:
                    fn = u""
            if _norm_key(fn) != tgt:
                continue
            out.append(sym)
    except Exception:
        return []
    return out


def _symbol_type_labels(sym):
    labels = []
    seen = set()
    if sym is None:
        return labels

    def _push(raw):
        for k in _label_lookup_keys(raw):
            if k not in seen:
                seen.add(k)
                labels.append(k)

    try:
        _push(getattr(sym, "Name", None))
    except Exception:
        pass
    for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = sym.get_Parameter(bip)
            if p is None or not p.HasValue or p.StorageType != StorageType.String:
                continue
            _push(p.AsString())
        except Exception:
            continue
    return labels


def _collect_tag_symbol_map(document, family_name):
    """Mapa clave de tipo (01, 1, …) → ``FamilySymbol``."""
    out = {}
    for sym in _family_symbols_rebar_tag(document, family_name):
        for lab in _symbol_type_labels(sym):
            if lab not in out:
                out[lab] = sym
    return out


def _lookup_tag_symbol(tag_map, label):
    if not tag_map:
        return None
    for k in _label_lookup_keys(label):
        sym = tag_map.get(k)
        if sym is not None:
            try:
                if not sym.IsActive:
                    sym.Activate()
            except Exception:
                pass
            return sym
    return None


def _get_rebar_shape_element(document, rebar):
    if document is None or rebar is None:
        return None
    sid = None
    try:
        sid = rebar.GetShapeId()
    except Exception:
        sid = None
    if sid is None or sid == ElementId.InvalidElementId:
        try:
            sid = rebar.RebarShapeId
        except Exception:
            sid = None
    if sid is None or sid == ElementId.InvalidElementId:
        return None
    try:
        return document.GetElement(sid)
    except Exception:
        return None


def _primary_rebar_shape_tag_key(document, rebar):
    shp = _get_rebar_shape_element(document, rebar)
    if shp is None:
        return None
    for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = shp.get_Parameter(bip)
            if p is None or not p.HasValue or p.StorageType != StorageType.String:
                continue
            s = _norm_key(p.AsString())
            if s:
                return s
        except Exception:
            continue
    try:
        s = _norm_key(getattr(shp, "Name", None))
        return s if s else None
    except Exception:
        return None


def _rebar_shape_name_candidates(document, rebar):
    out = []
    seen = set()
    if document is None or rebar is None:
        return out

    def push(raw):
        for k in _label_lookup_keys(raw):
            if k not in seen:
                seen.add(k)
                out.append(k)

    shp = _get_rebar_shape_element(document, rebar)
    if shp is not None:
        for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
            try:
                bip = getattr(BuiltInParameter, bip_name, None)
                if bip is None:
                    continue
                p = shp.get_Parameter(bip)
                if p is not None and p.HasValue and p.StorageType == StorageType.String:
                    push(p.AsString())
            except Exception:
                continue
        try:
            push(getattr(shp, "Name", None))
        except Exception:
            pass

    for bip_name in (u"REBAR_SHAPE",):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = rebar.get_Parameter(bip)
            if p is None or not p.HasValue:
                continue
            if p.StorageType == StorageType.String:
                push(p.AsString())
            elif p.StorageType == StorageType.ElementId:
                eid = p.AsElementId()
                if eid is not None and eid != ElementId.InvalidElementId:
                    el = document.GetElement(eid)
                    if el is not None:
                        push(getattr(el, "Name", None))
        except Exception:
            continue
    return out


def _centerline_segment_count(rebar):
    if rebar is None:
        return 0
    mpo = None
    if MultiPlanarOption is not None:
        try:
            mpo = MultiPlanarOption.IncludeAllMultiPlanarCurves
        except Exception:
            try:
                mpo = MultiPlanarOption.IncludeOnlyPlanarCurves
            except Exception:
                mpo = None
    attempts = []
    if mpo is not None:
        attempts.append((False, False, False, mpo, 0))
    attempts.append((False, False, False, 0))
    attempts.append((False, False, False))
    for args in attempts:
        try:
            curves = rebar.GetCenterlineCurves(*args)
            if curves is not None:
                n = len(list(curves))
                if n > 0:
                    return n
        except Exception:
            continue
    return 0


def _inferred_tag_shape_key(rebar):
    n_seg = _centerline_segment_count(rebar)
    if n_seg in CABEZAL_TAG_SHAPE_BY_SEGMENTS:
        return CABEZAL_TAG_SHAPE_BY_SEGMENTS[n_seg]
    if n_seg > 4:
        return CABEZAL_TAG_SHAPE_BY_SEGMENTS.get(4, CABEZAL_TAG_SHAPE_FALLBACK)
    return CABEZAL_TAG_SHAPE_FALLBACK


def _resolve_tag_symbol_for_rebar(document, family_name, tag_map, rebar):
    candidates = []
    primary = _primary_rebar_shape_tag_key(document, rebar)
    if primary:
        candidates.append(primary)
    candidates.extend(_rebar_shape_name_candidates(document, rebar))
    seen = set()
    ordered = []
    for c in candidates:
        if c and c not in seen:
            seen.add(c)
            ordered.append(c)

    for c in ordered:
        sym = _lookup_tag_symbol(tag_map, c)
        if sym is not None:
            return sym, c

    inferred = _inferred_tag_shape_key(rebar)
    sym = _lookup_tag_symbol(tag_map, inferred)
    if sym is not None:
        return sym, inferred + u" (geom)"

    sym = _lookup_tag_symbol(tag_map, CABEZAL_TAG_SHAPE_FALLBACK)
    if sym is not None:
        return sym, CABEZAL_TAG_SHAPE_FALLBACK + u" (fallback)"

    return None, None


def _mm_to_internal(mm):
    try:
        return float(UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters))
    except Exception:
        return float(mm) / 304.8


def _view_scale_denominator(view):
    """Denominador de escala de la vista (p. ej. 50 → 1/50)."""
    if view is None:
        return CABEZAL_TAG_SCALE_REFERENCE
    try:
        s = int(view.Scale)
        if s > 0:
            return s
    except Exception:
        pass
    try:
        bip = getattr(BuiltInParameter, u"VIEW_SCALE", None)
        if bip is not None:
            p = view.get_Parameter(bip)
            if p is not None and p.HasValue:
                s = int(p.AsInteger())
                if s > 0:
                    return s
    except Exception:
        pass
    return CABEZAL_TAG_SCALE_REFERENCE


def _cabezal_tag_spacing_mm_for_scale(scale_denom):
    """
    Distancias en mm modelo según escala de vista.

    Escenarios 1/50, 1/75 y 1/100 tabulados; otras escalas: proporcional a 1/50.
    """
    try:
        sd = int(scale_denom)
    except Exception:
        sd = CABEZAL_TAG_SCALE_REFERENCE
    if sd <= 0:
        sd = CABEZAL_TAG_SCALE_REFERENCE
    tabulated = CABEZAL_TAG_SPACING_BY_SCALE.get(sd)
    if tabulated is not None:
        return float(tabulated[0]), float(tabulated[1]), sd
    ratio = float(sd) / float(CABEZAL_TAG_SCALE_REFERENCE)
    offset_mm = CABEZAL_TAG_OFFSET_MM * ratio
    step_mm = CABEZAL_TAG_LAYER_STEP_MM * ratio
    return offset_mm, step_mm, sd


def _cabezal_tag_spacing_mm_for_view(view):
    return _cabezal_tag_spacing_mm_for_scale(_view_scale_denominator(view))


def _coronamiento_tag_spacing_mm_for_scale(scale_denom):
    try:
        sd = int(scale_denom)
    except Exception:
        sd = CORONAMIENTO_TAG_SCALE_REFERENCE
    if sd <= 0:
        sd = CORONAMIENTO_TAG_SCALE_REFERENCE
    tab = CORONAMIENTO_TAG_SPACING_BY_SCALE.get(sd)
    if tab is not None:
        return float(tab[0]), float(tab[1]), sd
    ratio = float(sd) / float(CORONAMIENTO_TAG_SCALE_REFERENCE)
    return (
        CORONAMIENTO_TAG_OFFSET_MM * ratio,
        CORONAMIENTO_TAG_LAYER_STEP_MM * ratio,
        sd,
    )


def _coronamiento_tag_spacing_mm_for_view(view):
    return _coronamiento_tag_spacing_mm_for_scale(_view_scale_denominator(view))


def _is_coronamiento_tag_meta(meta):
    if not meta:
        return False
    ex = meta.get(u"extremo")
    if ex is None:
        return False
    try:
        s = unicode(ex).strip().lower()
    except Exception:
        s = str(ex).strip().lower()
    return s.startswith(u"cor_")


def _coronamiento_tag_offset_sign(extremo):
    """
    Sentido del offset en ``view.UpDirection`` (+1 arriba en vista, −1 abajo).

    Criterio por **cota de la barra** en alzado (no por el nombre del rol):

    | Tipo | Cota | Offset |
    |------|------|--------|
    | cor_sup | tope muro | +Up |
    | cor_inf / cor_pie | base / fundación | −Up |
    | cor_vol_inf_* | tope muro inferior (reentrada) | +Up |
    | cor_vol_sup_* | base muro superior (voladizo) | −Up |
    """
    try:
        s = unicode(extremo or u"").strip().lower()
    except Exception:
        s = str(extremo or u"").strip().lower()
    if s in (u"cor_inf", u"cor_pie"):
        return -1
    if s.startswith(u"cor_vol_sup"):
        return -1
    if s.startswith(u"cor_vol_inf"):
        return 1
    return 1


def _longest_mostly_horizontal_centerline(rebar, view=None):
    """Tramo horizontal dominante (tronco de coronamiento U), no las patas L."""
    best = None
    best_len = -1.0
    vd = None
    if view is not None:
        try:
            vd = view.ViewDirection
        except Exception:
            vd = None
    for c in _rebar_centerline_curves_list(rebar):
        try:
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
            dz = abs(float(p1.Z) - float(p0.Z))
            ln = float(c.Length)
            if ln < 1e-9:
                continue
            if dz > ln * 0.35:
                continue
            if vd is not None:
                delta = p1 - p0
                d_proj = _project_to_view_plane(delta, vd)
                ln_view = float(d_proj.GetLength()) if d_proj is not None else ln
            else:
                ln_view = math.hypot(
                    float(p1.X) - float(p0.X),
                    float(p1.Y) - float(p0.Y),
                )
            if ln_view > best_len:
                best_len = ln_view
                best = c
        except Exception:
            continue
    return best


def _rebar_coronamiento_horizontal_anchor(rebar, meta, view):
    """Ancla en el centro del tramo horizontal (no bbox 3D ni tronco vertical)."""
    hc = _longest_mostly_horizontal_centerline(rebar, view=view)
    if hc is not None:
        try:
            return hc.Evaluate(0.5, True)
        except Exception:
            pass
    if meta is not None:
        try:
            zs = float(meta.get(u"zs"))
            bb = None
            try:
                bb = rebar.get_BoundingBox(view)
            except Exception:
                bb = None
            if bb is None:
                try:
                    bb = rebar.get_BoundingBox(None)
                except Exception:
                    bb = None
            if bb is not None:
                return XYZ(
                    (bb.Min.X + bb.Max.X) * 0.5,
                    (bb.Min.Y + bb.Max.Y) * 0.5,
                    zs,
                )
        except Exception:
            pass
    return _rebar_centerline_midpoint(rebar)


def _calcular_cabeza_tag_coronamiento(rebar, view, meta, offset_mm):
    anchor_raw = _rebar_coronamiento_horizontal_anchor(rebar, meta, view)
    if anchor_raw is None:
        anchor_raw = _punto_insercion_tag(rebar, view)
    anchor = _proyectar_punto_vista(anchor_raw, view)
    if anchor is None:
        return None, None
    sign = _coronamiento_tag_offset_sign(
        meta.get(u"extremo") if meta else None,
    )
    d = _unit_vector(view.UpDirection)
    if d is None:
        d = XYZ(0.0, 0.0, 1.0)
    if sign < 0:
        try:
            d = d.Negate()
        except Exception:
            d = XYZ(0.0, 0.0, -1.0)
    off_ft = _mm_to_internal(offset_mm)
    try:
        head = anchor + d.Multiply(off_ft)
    except Exception:
        head = XYZ(
            anchor.X + d.X * off_ft,
            anchor.Y + d.Y * off_ft,
            anchor.Z + d.Z * off_ft,
        )
    return anchor, head


def _precompute_coronamiento_tag_layout(document, view, tag_meta, offset_mm, layer_step_mm):
    """Cabezas desplazadas en vertical de vista (alzado), no en espesor del muro."""
    groups = {}
    for m in tag_meta or []:
        ak = _align_key_tag_meta(m)
        if ak is None:
            continue
        groups.setdefault(ak, []).append(m)

    tag_heads = {}
    offset_ft = _mm_to_internal(offset_mm)
    for _ak, items in groups.items():
        ordered = sorted(items, key=_tag_vertical_sort_key)
        group_heads = {}
        for m in ordered:
            rid_key = _rebar_id_int(m.get(u"rebar_id"))
            if rid_key is None:
                continue
            rb = document.GetElement(m.get(u"rebar_id"))
            if rb is None:
                continue
            anchor_raw = _rebar_coronamiento_horizontal_anchor(rb, m, view)
            anchor_v = _proyectar_punto_vista(anchor_raw, view)
            if anchor_v is None:
                continue
            sign = _coronamiento_tag_offset_sign(m.get(u"extremo"))
            d = _unit_vector(view.UpDirection)
            if d is None:
                d = XYZ(0.0, 0.0, 1.0)
            if sign < 0:
                try:
                    d = d.Negate()
                except Exception:
                    d = XYZ(0.0, 0.0, -1.0)
            try:
                head = anchor_v + d.Multiply(offset_ft)
            except Exception:
                continue
            group_heads[rid_key] = head
        group_heads = _spread_heads_vertical_no_overlap(
            group_heads, view, layer_step_mm,
        )
        tag_heads.update(group_heads)
    return tag_heads


def _confinement_ncapas_spacing_table(n_capas):
    try:
        n = int(n_capas)
    except Exception:
        n = 3
    if n >= 5:
        return CABEZAL_CONFINEMENT_TAG_5CAPAS_SPACING_BY_SCALE
    if n == 4:
        return CABEZAL_CONFINEMENT_TAG_4CAPAS_SPACING_BY_SCALE
    return CABEZAL_CONFINEMENT_TAG_3CAPAS_SPACING_BY_SCALE


def _confinement_ncapas_spacing_mm(view, n_capas, kind):
    """Mm modelo para espaciado de etiquetas @ 3–6 capas (tabla o proporcional a 1/50)."""
    try:
        sd = int(_view_scale_denominator(view))
    except Exception:
        sd = CABEZAL_TAG_SCALE_REFERENCE
    if sd <= 0:
        sd = CABEZAL_TAG_SCALE_REFERENCE
    tables = _confinement_ncapas_spacing_table(n_capas)
    tab = tables.get(sd)
    if tab is not None and kind in tab:
        return float(tab[kind])
    ref_tab = tables.get(CABEZAL_TAG_SCALE_REFERENCE)
    if ref_tab is None or kind not in ref_tab:
        return 0.0
    ratio = float(sd) / float(CABEZAL_TAG_SCALE_REFERENCE)
    return float(ref_tab[kind]) * ratio


def _confinement_3capas_spacing_mm(view, kind):
    """Compatibilidad: espaciado @ 3 capas."""
    return _confinement_ncapas_spacing_mm(view, 3, kind)


def _confinement_tag_offset_mm_for_view(view, n_capas=None, job_kind=None):
    """
    Offset mm modelo hacia fuera del muro.

    3 capas Tipo 2 estribo: tabla ``tipo2_estribo_mm``; resto: 250 mm @ 1/50 proporcional.
    """
    try:
        n = int(n_capas) if n_capas is not None else 0
    except Exception:
        n = 0
    if n in (3, 4, 5, 6) and job_kind == u"stirrup":
        return _confinement_ncapas_spacing_mm(
            view, n, CABEZAL_CONF_3C_TIPO2_ESTRIBO,
        )
    scale = _view_scale_denominator(view)
    ratio = float(scale) / float(CABEZAL_TAG_SCALE_REFERENCE)
    return float(CABEZAL_CONFINEMENT_TAG_OFFSET_MM_AT_REF_SCALE) * ratio


def confinement_tag_extra_offset_mm(n_capas, conf_type, job_kind, tie_layer_index=None):
    """
    Desplazamiento adicional (mm @ ref 1/50) para traba @ 3 capas Tipo 2.

    Suma a la columna del estribo (``tipo2_traba_extra_mm``); el desplazamiento lateral
    en vista se aplica aparte en ``_calcular_cabeza_tag_confinamiento_estribo``.
    """
    try:
        n = int(n_capas)
    except Exception:
        return 0.0
    if n != 3 or job_kind != u"tie":
        return 0.0
    ct = _norm_conf_type(conf_type)
    if ct != _CONF_TYPE_PERIMETER:
        return 0.0
    return float(
        CABEZAL_CONFINEMENT_TAG_3CAPAS_SPACING_BY_SCALE.get(
            CABEZAL_TAG_SCALE_REFERENCE, {},
        ).get(CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA, 220.0)
    )


def confinement_tag_extra_offset_mm_for_view(
    view, n_capas, conf_type, job_kind, tie_layer_index=None,
):
    """Incremento escalado a la vista (traba Tipo 2 @ 3 capas)."""
    ref = confinement_tag_extra_offset_mm(
        n_capas, conf_type, job_kind, tie_layer_index=tie_layer_index,
    )
    if ref <= 0.0:
        return 0.0
    return _confinement_ncapas_spacing_mm(
        view, int(n_capas), CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA,
    )


def _confinement_multihost_tipo2_tie(n_capas, conf_type, job_kind):
    """Trabas interiores Tipo 2 con multihost (4 capas: [1,2]; 5: [1,2,3]; 6: [1..4])."""
    try:
        n = int(n_capas)
    except Exception:
        return False
    return (
        n >= 4
        and job_kind == u"tie"
        and _norm_conf_type(conf_type) == _CONF_TYPE_PERIMETER
    )


def _confinement_multihost_tipo1_tie(n_capas, conf_type, job_kind):
    """Trabas Tipo 1 con multihost (3: [1,2]; 4: [1,2,3]; 5: [1..4]; 6: [1..5])."""
    try:
        n = int(n_capas)
    except Exception:
        return False
    return (
        n >= 3
        and job_kind == u"tie"
        and _norm_conf_type(conf_type) == _CONF_TYPE_TIE_LAYER_1
    )


def _confinement_tag_lateral_mm_for_view(view, n_capas, conf_type, job_kind):
    """Desplazamiento lateral (mm) de traba Tipo 2 @ 3 capas en ``UpDirection`` de la vista."""
    try:
        n = int(n_capas)
    except Exception:
        return 0.0
    if n != 3 or job_kind != u"tie":
        return 0.0
    if _norm_conf_type(conf_type) != _CONF_TYPE_PERIMETER:
        return 0.0
    return _confinement_ncapas_spacing_mm(
        view, n, CABEZAL_CONF_3C_TIPO2_TRABA_LATERAL,
    )


def _unit_vector(vec):
    if vec is None:
        return None
    try:
        ln = float(vec.GetLength())
        if ln < 1e-12:
            return None
        return vec.Normalize()
    except Exception:
        return None


def _rebar_centerline_curves_list(rebar):
    if rebar is None:
        return []
    mpo = None
    if MultiPlanarOption is not None:
        try:
            mpo = MultiPlanarOption.IncludeAllMultiPlanarCurves
        except Exception:
            try:
                mpo = MultiPlanarOption.IncludeOnlyPlanarCurves
            except Exception:
                mpo = None
    attempts = []
    if mpo is not None:
        attempts.append((False, False, False, mpo, 0))
    attempts.append((False, False, False, 0))
    attempts.append((False, False, False))
    for args in attempts:
        try:
            curves = rebar.GetCenterlineCurves(*args)
            if curves is not None:
                cl = list(curves)
                if cl:
                    return cl
        except Exception:
            continue
    return []


def _rebar_centerline_dominant_curve(rebar):
    curves = _rebar_centerline_curves_list(rebar)
    if not curves:
        return None
    best = None
    best_len = -1.0
    for c in curves:
        try:
            ln = float(c.Length)
            if ln > best_len:
                best_len = ln
                best = c
        except Exception:
            continue
    return best if best is not None else curves[0]


def _rebar_centerline_midpoint(rebar):
    c = _rebar_centerline_dominant_curve(rebar)
    if c is None:
        return None
    try:
        return c.Evaluate(0.5, True)
    except Exception:
        return None


def _project_to_view_plane(vec, view_dir):
    if vec is None:
        return None
    vd = _unit_vector(view_dir)
    if vd is None:
        return vec
    try:
        return vec - vd.Multiply(vec.DotProduct(vd))
    except Exception:
        return vec


def _norm_extremo(extremo):
    try:
        s = unicode(extremo or u"").strip().lower()
    except Exception:
        s = str(extremo or u"").strip().lower()
    if s in (u"fin", u"end", u"final", u"termino", u"término"):
        return CABEZAL_EXTREMO_FIN
    return CABEZAL_EXTREMO_INICIO


def _get_wall_from_meta(document, meta):
    if document is None or not meta:
        return None
    wid = meta.get(u"wid")
    if wid is None:
        return None
    try:
        eid = ElementId(int(wid))
    except Exception:
        return None
    try:
        el = document.GetElement(eid)
    except Exception:
        return None
    if el is None:
        return None
    try:
        if isinstance(el, Wall):
            return el
    except Exception:
        pass
    return el


def _wall_normal_outward_en_vista(wall, view):
    """
    Hacia la cara exterior del muro (normal al espesor), proyectado al plano de la vista.

    En alzado de cabezal el desplazamiento debe ser en el espesor, no a lo largo del eje
    del muro (evita etiquetas encima de las barras).
    """
    if wall is None or view is None:
        return None
    vd_u = None
    try:
        vd_u = _unit_vector(view.ViewDirection)
    except Exception:
        pass
    outward = None
    try:
        orient = wall.Orientation
        if orient is not None and float(orient.GetLength()) > 1e-9:
            outward = orient.Normalize()
    except Exception:
        pass
    if outward is None:
        return None
    if vd_u is not None:
        proj = _project_to_view_plane(outward, vd_u)
        u = _unit_vector(proj)
        if u is not None:
            return u
    u = _unit_vector(outward)
    if u is not None:
        return u
    for fb in (view.RightDirection, view.UpDirection):
        try:
            p = _project_to_view_plane(fb, vd_u) if vd_u is not None else fb
            u = _unit_vector(p)
            if u is not None:
                return u
        except Exception:
            continue
    return None


def _wall_station_point(wall, extremo):
    """Punto del extremo en la curva del muro (eje)."""
    lc = location_curve_wall(wall) if location_curve_wall else None
    if lc is None:
        return None
    try:
        if _norm_extremo(extremo) == CABEZAL_EXTREMO_FIN:
            return lc.GetEndPoint(1)
        return lc.GetEndPoint(0)
    except Exception:
        return None


def _wall_bbox_corners(bb):
    if bb is None:
        return []
    try:
        mn, mx = bb.Min, bb.Max
    except Exception:
        return []
    out = []
    for x in (float(mn.X), float(mx.X)):
        for y in (float(mn.Y), float(mx.Y)):
            for z in (float(mn.Z), float(mx.Z)):
                out.append(XYZ(x, y, z))
    return out


def _scalar_range_on_dir(points, dir_u):
    if not points or dir_u is None:
        return None, None
    vals = []
    for p in points:
        try:
            vals.append(float(p.DotProduct(dir_u)))
        except Exception:
            continue
    if not vals:
        return None, None
    return min(vals), max(vals)


def _confinement_thickness_outward_dir(wall, view, extremo, anchor):
    """
    Normal al espesor en el plano de la vista, hacia el lado de la armadura
    (fuera del interior del muro, no hacia el otro lado del espesor).
    """
    if wall is None or view is None:
        return None
    vd_u = None
    try:
        vd_u = _unit_vector(view.ViewDirection)
    except Exception:
        pass
    outward_raw = None
    try:
        orient = wall.Orientation
        if orient is not None and float(orient.GetLength()) > 1e-9:
            outward_raw = orient
    except Exception:
        pass
    if outward_raw is None:
        return None
    outward = outward_raw
    if vd_u is not None:
        outward = _project_to_view_plane(outward_raw, vd_u)
    outward = _unit_vector(outward)
    if outward is None:
        return None
    try:
        if float(outward.GetLength()) < 0.15:
            return None
    except Exception:
        pass
    if anchor is None:
        return outward
    station = _wall_station_point(wall, extremo)
    if station is None:
        return outward
    try:
        v = XYZ(
            float(anchor.X) - float(station.X),
            float(anchor.Y) - float(station.Y),
            float(anchor.Z) - float(station.Z),
        )
        if vd_u is not None:
            v = _project_to_view_plane(v, vd_u)
        if v is not None and float(outward.DotProduct(v)) < 0.0:
            outward = XYZ(
                -float(outward.X), -float(outward.Y), -float(outward.Z),
            )
    except Exception:
        pass
    return outward


def _confinement_exterior_column_scalar(wall, view, outward_u, anchor, clearance_ft):
    """
    Escalar sobre ``outward_u`` en la cara exterior del muro + holgura (pies).

    La cabecera queda fuera del contorno del muro en alzado de cabezal.
    """
    if wall is None or outward_u is None:
        return None
    bb = None
    try:
        if view is not None:
            bb = wall.get_BoundingBox(view)
    except Exception:
        bb = None
    if bb is None:
        try:
            bb = wall.get_BoundingBox(None)
        except Exception:
            bb = None
    lo, hi = _scalar_range_on_dir(_wall_bbox_corners(bb), outward_u)
    if lo is None or hi is None:
        return None
    try:
        clr = float(clearance_ft)
    except Exception:
        clr = 0.0
    if anchor is not None:
        try:
            s_a = float(anchor.DotProduct(outward_u))
        except Exception:
            s_a = 0.5 * (lo + hi)
        if s_a >= 0.5 * (lo + hi):
            return hi + clr
        return lo - clr
    return hi + clr


def _wall_outward_dir_at_extremo(wall, extremo, view):
    """
    Unitario en el plano de la vista apuntando desde el extremo hacia fuera del muro.
    Inicio → opuesto a p0→p1; fin → opuesto a p1→p0.
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

    if _norm_extremo(extremo) == CABEZAL_EXTREMO_FIN:
        into_wall = t_hat.Negate()
    else:
        into_wall = t_hat
    try:
        outward = into_wall.Negate()
    except Exception:
        outward = XYZ(-float(into_wall.X), -float(into_wall.Y), -float(into_wall.Z))

    vd = None
    try:
        vd = view.ViewDirection if view is not None else None
    except Exception:
        vd = None
    outward_proj = _project_to_view_plane(outward, vd)
    u = _unit_vector(outward_proj)
    if u is not None:
        return u
    return _unit_vector(outward)


def _offset_dir_for_tag_meta(document, view, meta, rebar=None):
    wall = _get_wall_from_meta(document, meta)
    extremo = meta.get(u"extremo") if meta else None
    if wall is not None and extremo:
        d = _wall_outward_dir_at_extremo(wall, extremo, view)
        if d is not None:
            return d
    if rebar is not None:
        return _offset_dir_tag_desde_rebar(rebar, view)
    try:
        rd = _unit_vector(view.RightDirection)
        if rd is not None:
            return rd
    except Exception:
        pass
    return XYZ(1.0, 0.0, 0.0)


def _offset_dir_tag_desde_rebar(rebar, view):
    """Dirección unitaria en el plano de la vista, perpendicular al rebar."""
    vd = None
    try:
        vd = view.ViewDirection
    except Exception:
        pass

    tan = None
    c = _rebar_centerline_dominant_curve(rebar)
    if c is not None:
        try:
            tr = c.ComputeDerivatives(0.5, True)
            if tr is not None:
                tan = tr.BasisX
        except Exception:
            tan = None
        if tan is None:
            try:
                tan = c.GetEndPoint(1) - c.GetEndPoint(0)
            except Exception:
                tan = None

    tan_proj = _project_to_view_plane(tan, vd)
    tan_u = _unit_vector(tan_proj)
    vd_u = _unit_vector(vd)
    if tan_u is not None and vd_u is not None:
        try:
            perp = vd_u.CrossProduct(tan_u)
            perp_u = _unit_vector(perp)
            if perp_u is not None:
                return perp_u
        except Exception:
            pass

    try:
        rd_u = _unit_vector(view.RightDirection)
        if rd_u is not None:
            return rd_u
    except Exception:
        pass
    return XYZ(1.0, 0.0, 0.0)


def _align_key_tag_meta(meta):
    """Una columna de etiquetas por muro y extremo (todos los tramos/capas)."""
    if not meta:
        return None
    try:
        wid = int(meta.get(u"wid"))
    except Exception:
        wid = meta.get(u"wid")
    extremo = meta.get(u"extremo")
    if extremo is None:
        extremo = u""
    try:
        extremo = unicode(extremo)
    except Exception:
        extremo = str(extremo)
    return (wid, extremo)


def _rebar_id_int(rebar_id):
    if rebar_id is None:
        return None
    try:
        return int(rebar_id.IntegerValue)
    except Exception:
        try:
            return int(rebar_id)
        except Exception:
            return None


def _tag_vertical_sort_key(meta):
    """
    Orden de apilado: tramo más alto primero; dentro del tramo capa 0 → N.
    ``zs`` + ``span_seg`` identifican el tramo vertical (troceo/empalme).
    """
    try:
        zs = float(meta.get(u"zs") or 0.0)
    except Exception:
        zs = 0.0
    try:
        span = float(meta.get(u"span_seg") or 0.0)
    except Exception:
        span = 0.0
    z_mid = zs + span * 0.5
    try:
        li = int(meta.get(u"layer_index") or 0)
    except Exception:
        li = 0
    return (-z_mid, li)


def _index_tag_meta_by_rebar_id(tag_meta):
    out = {}
    for m in tag_meta or []:
        if not m:
            continue
        rid = m.get(u"rebar_id")
        if rid is None:
            continue
        try:
            key = int(rid.IntegerValue)
        except Exception:
            try:
                key = int(rid)
            except Exception:
                continue
        out[key] = m
    return out


def _longest_mostly_vertical_centerline(rebar):
    best = None
    best_len = -1.0
    for c in _rebar_centerline_curves_list(rebar):
        try:
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
            dz = abs(float(p1.Z) - float(p0.Z))
            dxy = math.hypot(
                float(p1.X) - float(p0.X),
                float(p1.Y) - float(p0.Y),
            )
            if dz < 1e-9:
                continue
            if dxy > dz * 1.5:
                continue
            ln = float(c.Length)
            if ln > best_len:
                best_len = ln
                best = c
        except Exception:
            continue
    return best


def _rebar_cabezal_trunk_anchor(rebar, meta, view):
    """
    Punto medio del tronco vertical (ignora pata L). Prioriza ``zs`` + ``span_seg``
    del job cabezal para cotas correctas con patas inferiores.
    """
    if meta is not None:
        try:
            zs = float(meta.get(u"zs"))
            span = float(meta.get(u"span_seg") or 0.0)
            z_mid = zs + span * 0.5
            if rebar is not None:
                bb = None
                try:
                    bb = rebar.get_BoundingBox(view)
                except Exception:
                    bb = None
                if bb is None:
                    try:
                        bb = rebar.get_BoundingBox(None)
                    except Exception:
                        bb = None
                if bb is not None:
                    return XYZ(
                        (bb.Min.X + bb.Max.X) * 0.5,
                        (bb.Min.Y + bb.Max.Y) * 0.5,
                        z_mid,
                    )
        except Exception:
            pass
    vc = _longest_mostly_vertical_centerline(rebar)
    if vc is not None:
        try:
            return vc.Evaluate(0.5, True)
        except Exception:
            pass
    return _rebar_centerline_midpoint(rebar)


def _head_on_outward_column(anchor_v, outward_u, column_scalar):
    """Misma distancia al muro; conserva la cota del ancla → leader perpendicular."""
    if anchor_v is None or outward_u is None or column_scalar is None:
        return None
    try:
        along = float(anchor_v.DotProduct(outward_u))
    except Exception:
        along = 0.0
    try:
        perp = anchor_v - outward_u.Multiply(along)
    except Exception:
        perp = anchor_v
    try:
        return perp + outward_u.Multiply(float(column_scalar))
    except Exception:
        return None


def _move_along_view_down(head, view, distance_ft):
    d = _tag_stack_dir_in_view(view)
    try:
        dist = float(distance_ft)
    except Exception:
        dist = 0.0
    return XYZ(
        head.X + d.X * dist,
        head.Y + d.Y * dist,
        head.Z + d.Z * dist,
    )


def _spread_heads_vertical_no_overlap(heads_by_rid, view, min_step_mm):
    """Separa cabezas demasiado próximas en vertical (vista) sin mover la columna."""
    if not heads_by_rid:
        return heads_by_rid
    up = _unit_vector(view.UpDirection)
    if up is None:
        return dict(heads_by_rid)
    step_ft = _mm_to_internal(min_step_mm)
    entries = sorted(
        heads_by_rid.items(),
        key=lambda kv: float(kv[1].DotProduct(up)),
        reverse=True,
    )
    out = {}
    prev_up = None
    for rid, head in entries:
        h = head
        try:
            cu = float(h.DotProduct(up))
        except Exception:
            cu = 0.0
        if prev_up is not None and (prev_up - cu) < step_ft:
            target_cu = prev_up - step_ft
            delta = cu - target_cu
            h = _move_along_view_down(h, view, delta)
            try:
                cu = float(h.DotProduct(up))
            except Exception:
                cu = target_cu
        out[rid] = h
        prev_up = cu
    return out


def _tag_stack_dir_in_view(view):
    """Hacia abajo en la vista: capa 0 arriba, capas posteriores debajo."""
    ud = _unit_vector(view.UpDirection)
    if ud is not None:
        try:
            return ud.Negate()
        except Exception:
            return XYZ(-float(ud.X), -float(ud.Y), -float(ud.Z))
    return XYZ(0.0, -1.0, 0.0)


def _precompute_cabezal_tag_layout(document, view, tag_meta, offset_mm, layer_step_mm):
    """
    Cabezas por barra: ancla en tronco vertical, columna común hacia fuera del muro,
    cota vertical del ancla (leaders horizontales) y separación anti-solape.
    """
    groups = {}
    for m in tag_meta or []:
        ak = _align_key_tag_meta(m)
        if ak is None:
            continue
        groups.setdefault(ak, []).append(m)

    tag_heads = {}
    offset_ft = _mm_to_internal(offset_mm)
    for _ak, items in groups.items():
        ordered = sorted(items, key=_tag_vertical_sort_key)
        outward_u = None
        column_scalar = None
        pending = []

        for m in ordered:
            rid_key = _rebar_id_int(m.get(u"rebar_id"))
            if rid_key is None:
                continue
            rb = document.GetElement(m.get(u"rebar_id"))
            if rb is None:
                continue
            if outward_u is None:
                outward_u = _offset_dir_for_tag_meta(document, view, m, rebar=rb)
            anchor_raw = _rebar_cabezal_trunk_anchor(rb, m, view)
            anchor_v = _proyectar_punto_vista(anchor_raw, view)
            if anchor_v is None or outward_u is None:
                continue
            if column_scalar is None and int(m.get(u"layer_index", 0)) == 0:
                try:
                    column_scalar = float(anchor_v.DotProduct(outward_u)) + offset_ft
                except Exception:
                    column_scalar = None
            pending.append((rid_key, anchor_v))

        if column_scalar is None and pending:
            try:
                column_scalar = float(pending[0][1].DotProduct(outward_u)) + offset_ft
            except Exception:
                column_scalar = None

        group_heads = {}
        for rid_key, anchor_v in pending:
            head = _head_on_outward_column(anchor_v, outward_u, column_scalar)
            if head is not None:
                group_heads[rid_key] = head

        group_heads = _spread_heads_vertical_no_overlap(
            group_heads, view, layer_step_mm,
        )
        tag_heads.update(group_heads)
    return tag_heads


def _resolve_tag_head(document, view, rebar, meta, tag_heads, offset_mm):
    rid_key = _rebar_id_int(meta.get(u"rebar_id")) if meta else None
    if rid_key is not None and tag_heads and rid_key in tag_heads:
        return tag_heads[rid_key]
    if _is_coronamiento_tag_meta(meta):
        cor_off = float(offset_mm)
        try:
            if meta and meta.get(u"_ctx_cor_offset_mm") is not None:
                cor_off = float(meta.get(u"_ctx_cor_offset_mm"))
        except Exception:
            pass
        _, head = _calcular_cabeza_tag_coronamiento(rebar, view, meta, cor_off)
        return head
    offset_dir = _offset_dir_for_tag_meta(document, view, meta, rebar=rebar)
    anchor_raw = _rebar_cabezal_trunk_anchor(rebar, meta, view)
    if anchor_raw is None:
        anchor_raw = _punto_insercion_tag(rebar, view)
    anchor_v = _proyectar_punto_vista(anchor_raw, view)
    if anchor_v is None:
        return None
    off_ft = _mm_to_internal(offset_mm)
    if offset_dir is None:
        offset_dir = _offset_dir_tag_desde_rebar(rebar, view)
    try:
        col = float(anchor_v.DotProduct(offset_dir)) + off_ft
    except Exception:
        col = off_ft
    return _head_on_outward_column(anchor_v, offset_dir, col)


def _calcular_cabeza_tag(rebar, view, offset_mm, flip=False, offset_dir=None, meta=None):
    anchor_raw = _rebar_cabezal_trunk_anchor(rebar, meta, view)
    if anchor_raw is None:
        anchor_raw = _rebar_centerline_midpoint(rebar)
    if anchor_raw is None:
        anchor_raw = _punto_insercion_tag(rebar, view)
    anchor = _proyectar_punto_vista(anchor_raw, view)
    if anchor is None:
        return None, None

    d = offset_dir
    if d is None:
        d = _offset_dir_tag_desde_rebar(rebar, view)
    if flip:
        try:
            d = d.Negate()
        except Exception:
            d = XYZ(-float(d.X), -float(d.Y), -float(d.Z))

    off_ft = _mm_to_internal(offset_mm)
    head = XYZ(
        anchor.X + d.X * off_ft,
        anchor.Y + d.Y * off_ft,
        anchor.Z + d.Z * off_ft,
    )
    return anchor, head


def _aplicar_estilo_tag(tag, head):
    if tag is None or head is None:
        return
    try:
        tag.TagOrientation = TagOrientation.Horizontal
    except Exception:
        pass
    try:
        tag.HasLeader = True
    except Exception:
        pass
    try:
        tag.TagHeadPosition = head
    except Exception:
        pass


def _punto_insercion_tag(rebar, view):
    if rebar is None:
        return None
    try:
        bb = rebar.get_BoundingBox(view)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    try:
        bb0 = rebar.get_BoundingBox(None)
        if bb0 is not None:
            return (bb0.Min + bb0.Max) * 0.5
    except Exception:
        pass
    return None


def _proyectar_punto_vista(p, view):
    if p is None or view is None:
        return p
    try:
        vd = view.ViewDirection
        if vd is None or float(vd.GetLength()) < 1e-12:
            return p
        vd = vd.Normalize()
        vo = view.Origin
        if vo is None:
            return p
        d = float((p - vo).DotProduct(vd))
        return p - vd.Multiply(d)
    except Exception:
        return p


def _referencias_tag_rebar(document, rebar, view):
    refs = []
    seen = set()

    def _add_ref(r):
        if r is None:
            return
        try:
            k = r.ConvertToStableRepresentation(document)
        except Exception:
            try:
                k = unicode(r)
            except Exception:
                k = id(r)
        if k in seen:
            return
        seen.add(k)
        refs.append(r)

    try:
        subs = rebar.GetSubelements() if hasattr(rebar, "GetSubelements") else None
    except Exception:
        subs = None
    if subs:
        for sub in subs:
            if sub is None:
                continue
            try:
                if hasattr(sub, "GetReference"):
                    _add_ref(sub.GetReference())
            except Exception:
                continue

    try:
        npos = int(getattr(rebar, "NumberOfBarPositions", 0))
    except Exception:
        npos = 0
    if npos > 0:
        for idx in (0, int(npos / 2), max(0, npos - 1)):
            try:
                if hasattr(rebar, "GetReferenceToBarPosition"):
                    _add_ref(rebar.GetReferenceToBarPosition(idx))
                elif hasattr(rebar, "GetReferenceForBarPosition"):
                    _add_ref(rebar.GetReferenceForBarPosition(idx))
            except Exception:
                continue

    try:
        _add_ref(Reference(rebar))
    except Exception:
        pass

    try:
        opts = Options()
        opts.ComputeReferences = True
        opts.View = view
        opts.IncludeNonVisibleObjects = True
        geom = rebar.get_Geometry(opts)
        if geom is not None:
            for go in geom:
                if go is None:
                    continue
                try:
                    rgo = getattr(go, "Reference", None)
                    if rgo is not None:
                        _add_ref(rgo)
                except Exception:
                    pass
    except Exception:
        pass

    return refs


def _crear_tag_rebar(document, view, rebar, tag_symbol_id, head_pos=None):
    head = head_pos
    if head is None:
        offset_mm, _step, _sd = _cabezal_tag_spacing_mm_for_view(view)
        _anchor, head = _calcular_cabeza_tag(
            rebar,
            view,
            offset_mm,
            offset_dir=_offset_dir_tag_desde_rebar(rebar, view),
        )
    if head is None:
        return None, u"sin punto de inserción"
    refs = _referencias_tag_rebar(document, rebar, view)
    if not refs:
        return None, u"sin referencia API"
    orient = TagOrientation.Horizontal
    add_leader = True
    last_ex = None
    for ref in refs:
        try:
            tag = IndependentTag.Create(
                document,
                tag_symbol_id,
                view.Id,
                ref,
                add_leader,
                orient,
                head,
            )
            if tag is not None:
                _aplicar_estilo_tag(tag, head)
                return tag, None
        except Exception as ex:
            last_ex = ex
            tag = None
    for ref in refs:
        try:
            tag = IndependentTag.Create(
                document,
                view.Id,
                ref,
                add_leader,
                TagMode.TM_ADDBY_CATEGORY,
                orient,
                head,
            )
            if tag is not None:
                try:
                    tag.SetTypeId(tag_symbol_id)
                except Exception:
                    pass
                _aplicar_estilo_tag(tag, head)
                return tag, None
        except Exception as ex:
            last_ex = ex
    if last_ex is not None:
        try:
            return None, unicode(last_ex)
        except Exception:
            return None, str(last_ex)
    return None, u"no se pudo crear IndependentTag"


def _dedupe_rebar_ids(rebar_ids):
    out = []
    seen = set()
    for rid in rebar_ids or []:
        if rid is None:
            continue
        try:
            iv = int(rid.IntegerValue)
        except Exception:
            continue
        if iv in seen:
            continue
        seen.add(iv)
        out.append(rid)
    return out


def _prepare_cabezal_tag_context(document, view, tag_meta, family_name):
    """
    Mapa de símbolos, cabezas precomputadas y metadatos para etiquetar.
    Retorna ``(ctx, error_msg)``; ``ctx`` es None si falla la preparación.
    """
    if document is None:
        return None, u"Etiquetado: sin documento."
    if not _vista_permite_rebar_tags(view):
        return None, (
            u"Etiquetado: use planta, alzado o sección (no plantilla ni 3D)."
        )
    tag_map = _collect_tag_symbol_map(document, family_name)
    if not tag_map:
        return None, (
            u"Etiquetado: no hay tipos OST_RebarTags para familia «{0}».".format(
                family_name,
            ),
        )
    try:
        document.Regenerate()
    except Exception:
        pass
    meta_by_rid = _index_tag_meta_by_rebar_id(tag_meta)
    offset_mm, layer_step_mm, scale_denom = _cabezal_tag_spacing_mm_for_view(view)
    meta_list = list(tag_meta or [])
    cab_meta = [m for m in meta_list if not _is_coronamiento_tag_meta(m)]
    cor_meta = [m for m in meta_list if _is_coronamiento_tag_meta(m)]
    tag_heads = {}
    if cab_meta:
        tag_heads.update(
            _precompute_cabezal_tag_layout(
                document, view, cab_meta, offset_mm, layer_step_mm,
            ),
        )
    cor_offset_mm = CORONAMIENTO_TAG_OFFSET_MM
    cor_step_mm = CORONAMIENTO_TAG_LAYER_STEP_MM
    if cor_meta:
        cor_offset_mm, cor_step_mm, _ = _coronamiento_tag_spacing_mm_for_view(view)
        tag_heads.update(
            _precompute_coronamiento_tag_layout(
                document, view, cor_meta, cor_offset_mm, cor_step_mm,
            ),
        )
    return {
        u"tag_map": tag_map,
        u"meta_by_rid": meta_by_rid,
        u"tag_heads": tag_heads,
        u"offset_mm": offset_mm,
        u"layer_step_mm": layer_step_mm,
        u"scale_denom": scale_denom,
        u"cor_offset_mm": cor_offset_mm,
        u"cor_step_mm": cor_step_mm,
        u"family_name": family_name,
    }, None


def _etiquetar_cabezal_rebar_ids_lote(document, view, rid_list, ctx, result):
    """Etiqueta un subconjunto de rebars (sin transacción)."""
    if not rid_list or not ctx or result is None:
        return
    tag_map = ctx[u"tag_map"]
    meta_by_rid = ctx[u"meta_by_rid"]
    tag_heads = ctx[u"tag_heads"]
    offset_mm = ctx[u"offset_mm"]
    cor_offset_mm = ctx.get(u"cor_offset_mm", CORONAMIENTO_TAG_OFFSET_MM)
    family_name = ctx[u"family_name"]
    for rid in rid_list:
        rb = document.GetElement(rid)
        if rb is None or not isinstance(rb, Rebar):
            result[u"n_fail"] += 1
            continue
        try:
            rid_key = int(rid.IntegerValue)
        except Exception:
            rid_key = None
        meta = meta_by_rid.get(rid_key) if rid_key is not None else None
        if meta is not None and _is_coronamiento_tag_meta(meta):
            meta = dict(meta)
            meta[u"_ctx_cor_offset_mm"] = cor_offset_mm
        head_pos = _resolve_tag_head(
            document, view, rb, meta, tag_heads, offset_mm,
        )
        sym, shape_lbl = _resolve_tag_symbol_for_rebar(
            document, family_name, tag_map, rb,
        )
        if sym is None:
            result[u"n_fail"] += 1
            if len(result[u"messages"]) < _MAX_TAG_WARNINGS:
                n_seg = _centerline_segment_count(rb)
                try:
                    result[u"messages"].append(
                        u"Etiqueta rebar {0}: sin tipo (tramos={1}) en «{2}».".format(
                            rid.IntegerValue,
                            n_seg,
                            family_name,
                        ),
                    )
                except Exception:
                    pass
            continue
        tag, err = _crear_tag_rebar(
            document, view, rb, sym.Id, head_pos=head_pos,
        )
        if tag is not None:
            result[u"n_ok"] += 1
        else:
            result[u"n_fail"] += 1
            if len(result[u"messages"]) < _MAX_TAG_WARNINGS:
                try:
                    result[u"messages"].append(
                        u"Etiqueta rebar {0} ({1}): {2}.".format(
                            rid.IntegerValue,
                            shape_lbl or u"?",
                            err or u"fallo",
                        ),
                    )
                except Exception:
                    pass


def etiquetar_cabezal_longitudinales_en_vista_animado(
    document,
    view,
    rebar_ids,
    tag_meta=None,
    family_name=CABEZAL_REBAR_TAG_FAMILY_NAME,
    batch_size=1,
    after_batch=None,
):
    """
    Etiqueta en lotes (mismo orden que ``rebar_ids``) con commit + callback por lote.

    ``after_batch``: opcional, p. ej. refresco de vista para aparición progresiva.
    """
    result = {u"n_ok": 0, u"n_fail": 0, u"messages": []}
    ids = _dedupe_rebar_ids(rebar_ids)
    if not ids:
        return result
    ctx, err = _prepare_cabezal_tag_context(
        document, view, tag_meta, family_name,
    )
    if ctx is None:
        if err:
            result[u"messages"].append(err)
        result[u"n_fail"] = len(ids)
        return result

    batch_n = max(1, int(batch_size or 1))
    n_lotes = int(math.ceil(float(len(ids)) / float(batch_n)))
    for lote_idx in range(n_lotes):
        i0 = lote_idx * batch_n
        lote = ids[i0:i0 + batch_n]
        i1 = min(i0 + batch_n, len(ids))
        if len(lote) == 1:
            txn_name = u"Arainco: Cabezal muros — etiqueta barra {0}/{1}".format(
                i0 + 1, len(ids),
            )
        else:
            txn_name = u"Arainco: Cabezal muros — etiquetas {0}–{1} de {2}".format(
                i0 + 1, i1, len(ids),
            )
        t = Transaction(document, txn_name)
        t.Start()
        lote_ok = False
        try:
            _etiquetar_cabezal_rebar_ids_lote(document, view, lote, ctx, result)
            t.Commit()
            lote_ok = True
        except Exception as ex_lote:
            try:
                if t.HasStarted():
                    t.RollBack()
            except Exception:
                pass
            result[u"messages"].append(
                u"Etiquetado lote {0}–{1}: {2}".format(i0 + 1, i1, ex_lote),
            )
            result[u"n_fail"] += len(lote)
        if lote_ok and after_batch is not None:
            try:
                after_batch()
            except Exception:
                pass

    if int(result.get(u"n_ok", 0)) > 0:
        try:
            cor_off = float(ctx.get(u"cor_offset_mm", CORONAMIENTO_TAG_OFFSET_MM))
            cor_step = float(ctx.get(u"cor_step_mm", CORONAMIENTO_TAG_LAYER_STEP_MM))
            result[u"messages"].append(
                u"Etiquetado escala 1/{0}: cabezal offset {1:.0f} mm; "
                u"coronamiento offset {2:.0f} mm, paso {3:.0f} mm.".format(
                    int(ctx[u"scale_denom"]),
                    float(ctx[u"offset_mm"]),
                    cor_off,
                    cor_step,
                ),
            )
        except Exception:
            pass
    return result


def etiquetar_cabezal_longitudinales_en_vista(
    document,
    view,
    rebar_ids,
    tag_meta=None,
    family_name=CABEZAL_REBAR_TAG_FAMILY_NAME,
):
    """
    Etiqueta cada ``Rebar`` longitudinal con Structural Rebar Tag en ``view``.

    ``tag_meta``: ``rebar_id``, ``layer_index``, ``wid``, ``extremo``, ``zs``,
    ``span_seg`` — columna por extremo y apilado global con troceo/empalme.

    Retorna dict ``n_ok``, ``n_fail``, ``messages``.
    """
    result = {u"n_ok": 0, u"n_fail": 0, u"messages": []}
    ids = _dedupe_rebar_ids(rebar_ids)
    if not ids:
        return result

    ctx, err = _prepare_cabezal_tag_context(
        document, view, tag_meta, family_name,
    )
    if ctx is None:
        if err:
            result[u"messages"].append(err)
        result[u"n_fail"] = len(ids)
        return result

    t = Transaction(document, u"Arainco: Cabezal muros — etiquetas longitudinales")
    t.Start()
    try:
        _etiquetar_cabezal_rebar_ids_lote(document, view, ids, ctx, result)
        t.Commit()
        if int(result.get(u"n_ok", 0)) > 0:
            try:
                cor_off = float(ctx.get(u"cor_offset_mm", CORONAMIENTO_TAG_OFFSET_MM))
                cor_step = float(ctx.get(u"cor_step_mm", CORONAMIENTO_TAG_LAYER_STEP_MM))
                result[u"messages"].append(
                    u"Etiquetado escala 1/{0}: cabezal offset {1:.0f} mm; "
                    u"coronamiento offset {2:.0f} mm, paso {3:.0f} mm.".format(
                        int(ctx[u"scale_denom"]),
                        float(ctx[u"offset_mm"]),
                        cor_off,
                        cor_step,
                    ),
                )
            except Exception:
                pass
    except Exception as ex:
        try:
            if t.HasStarted():
                t.RollBack()
        except Exception:
            pass
        result[u"messages"].append(u"Etiquetado: {0}".format(ex))
        result[u"n_fail"] = len(ids)
    return result


def collect_confinement_tag_symbol_map(document):
    """Mapa clave de tipo (según ``RebarShape``) → ``FamilySymbol`` de etiqueta confinamiento."""
    return _collect_tag_symbol_map(document, CABEZAL_CONFINEMENT_TAG_FAMILY_NAME)


def collect_confinement_multihost_tag_symbol_map(document):
    """Mapa de tipos para familia multihost de confinamiento (3+ capas)."""
    return _collect_tag_symbol_map(
        document, CABEZAL_CONFINEMENT_MULTIHOST_TAG_FAMILY_NAME,
    )


_CONF_TYPE_PERIMETER = u"perimeter_0_1"
_CONF_TYPE_TIE_LAYER_1 = u"tie_layer_1"


def _norm_conf_type(conf_type):
    try:
        return unicode(conf_type or u"").strip()
    except Exception:
        return str(conf_type or u"").strip()


def confinement_tag_mode(conf_type, n_capas, job_kind):
    """
    Modo de etiqueta de confinamiento.

    **2 capas:** estribo (Tipo 2) o traba (Tipo 1) — ``TAG_CONFINAMIENTO``.
    **3 capas Tipo 1:** trabas [1,2] — multihost.
    **3 capas Tipo 2:** estribo + traba [1] — ``TAG_CONFINAMIENTO``.
    **4 capas Tipo 1:** trabas [1,2,3] — multihost.
    **4 capas Tipo 2:** estribo + trabas [1,2] — estribo + multihost.
    **5 capas Tipo 1:** trabas [1,2,3,4] — multihost.
    **5 capas Tipo 2:** estribo + trabas [1,2,3] — estribo + multihost.
    **6 capas Tipo 1:** trabas [1..5] — multihost.
    **6 capas Tipo 2:** estribo + trabas [1..4] — estribo + multihost.
    """
    ct = _norm_conf_type(conf_type)
    try:
        n = int(n_capas)
    except Exception:
        n = 0
    if job_kind == u"stirrup" and ct == _CONF_TYPE_PERIMETER:
        return u"conf_tag"
    if n == 2 and job_kind == u"tie" and ct == _CONF_TYPE_TIE_LAYER_1:
        return u"conf_tag"
    if _confinement_multihost_tipo1_tie(n, ct, job_kind):
        return u"conf_tag_multihost"
    if n == 3 and job_kind == u"tie" and ct == _CONF_TYPE_PERIMETER:
        return u"conf_tag"
    if _confinement_multihost_tipo2_tie(n, ct, job_kind):
        return u"conf_tag_multihost"
    return None


def _aplicar_estilo_tag_confinamiento_estribo(tag, head):
    """Vertical, sin leader, cabecera en ``head``."""
    if tag is None or head is None:
        return
    try:
        tag.TagOrientation = TagOrientation.Vertical
    except Exception:
        pass
    try:
        tag.HasLeader = False
    except Exception:
        pass
    try:
        tag.TagHeadPosition = head
    except Exception:
        pass


def compute_confinement_tag_anchor(
    document, view, rebar, wall=None, extremo=None, multihost=False, rebar_list=None,
):
    """Punto de anclaje compartido para alinear etiquetas Tipo 2 (estribo + traba)."""
    return _anchor_confinement_tag(
        document, view, rebar, rebar_list=rebar_list if multihost else None,
    )


def confinement_tipo2_group_key(cj):
    """Clave por muro/extremo/pila para anclaje compartido estribo–traba @ 3 capas Tipo 2."""
    return confinement_multihost_group_key(cj)


def register_confinement_tipo2_stirrup_anchor(res, cj, anchor_pt):
    if res is None or cj is None or anchor_pt is None:
        return
    key = confinement_tipo2_group_key(cj)
    res.setdefault(u"_conf_tipo2_stirrup_anchor", {})[key] = anchor_pt


def get_confinement_tipo2_stirrup_anchor(res, cj):
    if res is None or cj is None:
        return None
    key = confinement_tipo2_group_key(cj)
    return (res.get(u"_conf_tipo2_stirrup_anchor") or {}).get(key)


def register_confinement_tipo2_stirrup_head(res, cj, head_pt):
    """Cabecera real de +E. (Tipo 2) para alinear +nTr. multihost en la misma columna."""
    if res is None or cj is None or head_pt is None:
        return
    key = confinement_tipo2_group_key(cj)
    res.setdefault(u"_conf_tipo2_stirrup_head", {})[key] = head_pt


def get_confinement_tipo2_stirrup_head(res, cj):
    if res is None or cj is None:
        return None
    key = confinement_tipo2_group_key(cj)
    return (res.get(u"_conf_tipo2_stirrup_head") or {}).get(key)


def compute_confinement_stirrup_tag_head(
    document,
    view,
    rebar,
    wall=None,
    extremo=None,
    n_capas=None,
    conf_type=None,
):
    """Cabecera de etiqueta de estribo (misma lógica que al etiquetar al vuelo)."""
    return _calcular_cabeza_tag_confinamiento_estribo(
        document,
        view,
        rebar,
        wall,
        extremo,
        n_capas=n_capas,
        conf_type=conf_type,
        job_kind=u"stirrup",
    )


def _anchor_confinement_tag(document, view, rebar, rebar_list=None):
    """Punto de anclaje: centroide de varias trabas (multihost) o inserción de una barra."""
    candidatos = []
    if rebar_list:
        for rb in rebar_list:
            if isinstance(rb, Rebar):
                candidatos.append(rb)
    if not candidatos and isinstance(rebar, Rebar):
        candidatos = [rebar]
    pts = []
    for rb in candidatos:
        p_raw = _punto_insercion_tag(rb, view)
        p = _proyectar_punto_vista(p_raw, view)
        if p is not None:
            pts.append(p)
    if not pts:
        return None
    if len(pts) == 1:
        return pts[0]
    try:
        sx = sum(float(p.X) for p in pts)
        sy = sum(float(p.Y) for p in pts)
        sz = sum(float(p.Z) for p in pts)
        n = float(len(pts))
        return XYZ(sx / n, sy / n, sz / n)
    except Exception:
        return pts[0]


def _calcular_cabeza_tag_confinamiento_estribo(
    document,
    view,
    rebar,
    wall,
    extremo,
    extra_offset_mm=0.0,
    n_capas=None,
    conf_type=None,
    job_kind=None,
    multihost=False,
    rebar_list=None,
    anchor_override=None,
    stirrup_head_override=None,
):
    """
    Cabecera de etiqueta fuera del muro (desplazamiento según escala de vista).

    ``extra_offset_mm``: suma radial adicional (traba Tipo 2 @ 3 capas).
    ``multihost``: Tipo 1 → ``tipo1_multihost_mm``; Tipo 2 @ 4 capas → estribo + extra.
    ``anchor_override``: mismo anclaje que estribo (alineación Tipo 2).
    """
    if anchor_override is not None:
        anchor = anchor_override
    else:
        mh_list = rebar_list if multihost else None
        try:
            n_mh = int(n_capas or 0)
        except Exception:
            n_mh = 0
        if (
            multihost
            and mh_list
            and len(mh_list) > 1
            and n_mh >= 4
            and _norm_conf_type(conf_type) == _CONF_TYPE_TIE_LAYER_1
        ):
            # Tipo 1 @ 4+ capas: anclar en traba exterior (no centroide interior).
            mh_list = None
        anchor = _anchor_confinement_tag(
            document, view, rebar, rebar_list=mh_list,
        )
    if anchor is None:
        return None

    use_wall_face = wall is not None
    offset_dir = None
    if use_wall_face:
        offset_dir = _confinement_thickness_outward_dir(
            wall, view, extremo, anchor,
        )
    if offset_dir is None and use_wall_face and extremo:
        offset_dir = _wall_outward_dir_at_extremo(wall, extremo, view)
    if offset_dir is None and use_wall_face:
        offset_dir = _wall_normal_outward_en_vista(wall, view)
    if offset_dir is None and use_wall_face:
        try:
            bb = wall.get_BoundingBox(None)
            if bb is not None:
                cx = 0.5 * (float(bb.Min.X) + float(bb.Max.X))
                cy = 0.5 * (float(bb.Min.Y) + float(bb.Max.Y))
                inward = XYZ(cx - float(anchor.X), cy - float(anchor.Y), 0.0)
                if float(inward.GetLength()) > 1e-9:
                    offset_dir = inward.Normalize().Negate()
        except Exception:
            pass
    if offset_dir is None:
        offset_dir = _offset_dir_tag_desde_rebar(rebar, view)
    if offset_dir is None:
        try:
            offset_dir = _unit_vector(view.RightDirection)
        except Exception:
            offset_dir = XYZ(1.0, 0.0, 0.0)

    if multihost:
        try:
            n_mh = int(n_capas or 0)
        except Exception:
            n_mh = 3
        if (
            anchor_override is not None
            and stirrup_head_override is not None
        ):
            offset_mm = _confinement_ncapas_spacing_mm(
                view, n_mh, CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA,
            )
        elif anchor_override is not None:
            offset_mm = _confinement_tag_offset_mm_for_view(
                view, n_capas=n_mh, job_kind=u"stirrup",
            )
            offset_mm += _confinement_ncapas_spacing_mm(
                view, n_mh, CABEZAL_CONF_3C_TIPO2_TRABA_EXTRA,
            )
        else:
            offset_mm = _confinement_ncapas_spacing_mm(
                view, n_mh if n_mh > 0 else 3,
                CABEZAL_CONF_3C_TIPO1_MULTIHOST,
            )
    else:
        offset_mm = _confinement_tag_offset_mm_for_view(
            view, n_capas=n_capas, job_kind=job_kind,
        )
    offset_mm += float(extra_offset_mm or 0.0)
    off_ft = _mm_to_internal(offset_mm)

    head = None
    if use_wall_face and offset_dir is not None:
        try:
            n_mh = int(n_capas or 0)
        except Exception:
            n_mh = 0
        if (
            multihost
            and anchor_override is not None
            and n_mh >= 3
        ):
            estribo_mm = _confinement_tag_offset_mm_for_view(
                view, n_capas=n_mh, job_kind=u"stirrup",
            )
            col_estribo = _confinement_exterior_column_scalar(
                wall,
                view,
                offset_dir,
                anchor,
                _mm_to_internal(estribo_mm),
            )
            col = col_estribo + off_ft if col_estribo is not None else None
        else:
            col = _confinement_exterior_column_scalar(
                wall, view, offset_dir, anchor, off_ft,
            )
        if col is not None:
            head = _head_on_outward_column(anchor, offset_dir, col)
    if head is None:
        if (
            multihost
            and stirrup_head_override is not None
            and anchor_override is not None
        ):
            try:
                head = XYZ(
                    float(stirrup_head_override.X) + float(offset_dir.X) * off_ft,
                    float(stirrup_head_override.Y) + float(offset_dir.Y) * off_ft,
                    float(stirrup_head_override.Z) + float(offset_dir.Z) * off_ft,
                )
            except Exception:
                head = None
        else:
            try:
                col = float(anchor.DotProduct(offset_dir)) + off_ft
                head = _head_on_outward_column(anchor, offset_dir, col)
            except Exception:
                head = None
    if head is None:
        return None

    lat_mm = _confinement_tag_lateral_mm_for_view(
        view, n_capas, conf_type, job_kind,
    )
    if lat_mm > 0.0 and view is not None:
        try:
            v_up = view.UpDirection
            if v_up is not None and float(v_up.GetLength()) > 1e-12:
                v_up = v_up.Normalize()
                lat_ft = _mm_to_internal(lat_mm)
                head = XYZ(
                    float(head.X) + float(v_up.X) * lat_ft,
                    float(head.Y) + float(v_up.Y) * lat_ft,
                    float(head.Z) + float(v_up.Z) * lat_ft,
                )
        except Exception:
            pass
    return head


def _crear_tag_rebar_confinamiento(document, view, rebar, tag_symbol_id, head_pos):
    """``IndependentTag`` vertical, sin leader."""
    if head_pos is None:
        return None, u"sin punto de inserción"
    refs = _referencias_tag_rebar(document, rebar, view)
    if not refs:
        return None, u"sin referencia API"
    orient = TagOrientation.Vertical
    add_leader = False
    last_ex = None
    for ref in refs:
        try:
            tag = IndependentTag.Create(
                document,
                tag_symbol_id,
                view.Id,
                ref,
                add_leader,
                orient,
                head_pos,
            )
            if tag is not None:
                _aplicar_estilo_tag_confinamiento_estribo(tag, head_pos)
                try:
                    tag.TagHeadPosition = head_pos
                except Exception:
                    pass
                return tag, None
        except Exception as ex:
            last_ex = ex
    for ref in refs:
        try:
            tag = IndependentTag.Create(
                document,
                view.Id,
                ref,
                add_leader,
                TagMode.TM_ADDBY_CATEGORY,
                orient,
                head_pos,
            )
            if tag is not None:
                try:
                    tag.SetTypeId(tag_symbol_id)
                except Exception:
                    pass
                _aplicar_estilo_tag_confinamiento_estribo(tag, head_pos)
                try:
                    tag.TagHeadPosition = head_pos
                except Exception:
                    pass
                return tag, None
        except Exception as ex:
            last_ex = ex
    if last_ex is not None:
        try:
            return None, unicode(last_ex)
        except Exception:
            return None, str(last_ex)
    return None, u"no se pudo crear IndependentTag"


def etiquetar_cabezal_estribo_confinamiento(
    document,
    view,
    rebar,
    tag_map=None,
    wall=None,
    extremo=None,
    extra_offset_mm=0.0,
    n_capas=None,
    conf_type=None,
    job_kind=None,
    anchor_override=None,
):
    """
    Etiqueta de confinamiento (estribo / traba).

    Familia ``EST_A_STRUCTURAL REBAR TAG_CONFINAMIENTO``; tipo según ``RebarShape``.
    @ 3 capas Tipo 2: estribo 280 mm; traba +220 mm radial y mismo anclaje (alineada).
    """
    if document is None or view is None or rebar is None:
        return False, u"Etiqueta estribo: parámetros inválidos."
    if not isinstance(rebar, Rebar):
        return False, u"Etiqueta estribo: elemento no es Rebar."
    if not _vista_permite_rebar_tags(view):
        return False, (
            u"Etiqueta estribo: use planta, alzado o sección (no plantilla ni 3D)."
        )
    family = CABEZAL_CONFINEMENT_TAG_FAMILY_NAME
    if tag_map is None:
        tag_map = collect_confinement_tag_symbol_map(document)
    if not tag_map:
        return False, (
            u"Etiqueta estribo: no hay tipos OST_RebarTags para «{0}».".format(family)
        )
    sym, shape_lbl = _resolve_tag_symbol_for_rebar(
        document, family, tag_map, rebar,
    )
    if sym is None:
        shape_hint = shape_lbl or _primary_rebar_shape_tag_key(document, rebar) or u"?"
        return False, (
            u"Etiqueta estribo: sin tipo para RebarShape «{0}» en «{1}».".format(
                shape_hint, family,
            )
        )
    head = _calcular_cabeza_tag_confinamiento_estribo(
        document,
        view,
        rebar,
        wall,
        extremo,
        extra_offset_mm=extra_offset_mm,
        n_capas=n_capas,
        conf_type=conf_type,
        job_kind=job_kind,
        anchor_override=anchor_override,
    )
    tag, err = _crear_tag_rebar_confinamiento(
        document, view, rebar, sym.Id, head,
    )
    if tag is not None:
        return True, None
    return False, err or u"no se pudo crear IndependentTag"


def confinement_multihost_group_key(cj):
    """Una etiqueta multihost por muro y extremo (capas [1] y [2] del mismo cabezal)."""
    try:
        wid = int(cj.get(u"wid", 0) or 0)
    except Exception:
        wid = 0
    return (wid, _norm_extremo(cj.get(u"extremo")))


def register_confinement_multihost_traba_pending(res, cj, rebar):
    """Acumula trabas para una etiqueta multihost al cerrar el confinamiento."""
    if res is None or cj is None or rebar is None:
        return
    key = confinement_multihost_group_key(cj)
    pending = res.setdefault(u"_conf_multihost_traba_pending", {})
    grp = pending.get(key)
    if grp is None:
        grp = {
            u"items": [],
            u"wall": cj.get(u"wall"),
            u"extremo": cj.get(u"extremo"),
            u"conf_type": cj.get(u"conf_type"),
            u"n_capas": cj.get(u"ex_cfg", {}).get(u"n_capas"),
        }
        pending[key] = grp
    try:
        li = int(cj.get(u"tie_layer_index", 0) or 0)
    except Exception:
        li = 0
    grp.setdefault(u"items", []).append({
        u"rebar_id": rebar.Id,
        u"tie_layer_index": li,
    })


def _ordered_rebar_ids_from_multihost_grp(grp):
    items = list(grp.get(u"items") or [])
    if not items:
        return []
    try:
        items.sort(key=lambda x: int(x.get(u"tie_layer_index", 0)))
    except Exception:
        pass
    out = []
    seen = set()
    for it in items:
        rid = it.get(u"rebar_id")
        if rid is None:
            continue
        try:
            k = rid.IntegerValue
        except Exception:
            k = rid
        if k in seen:
            continue
        seen.add(k)
        out.append(rid)
    return out


def _refs_multihost_para_rebars(document, view, rebar_ids_extra):
    """Referencias para ``IndependentTag.AddReferences`` (patrón Borde Losa)."""
    if _rebar_reference_candidates_for_tag is None:
        return []
    refs = []
    seen = set()
    for rid in rebar_ids_extra or []:
        rb = document.GetElement(rid) if document is not None else None
        if not isinstance(rb, Rebar):
            continue
        cand = _rebar_reference_candidates_for_tag(document, view, rb)
        if not cand:
            continue
        ref0 = cand[0]
        key = None
        try:
            key = ref0.ConvertToStableRepresentation(document)
        except Exception:
            key = id(ref0)
        if key in seen:
            continue
        seen.add(key)
        refs.append(ref0)
    return refs


def _reposicionar_cabezera_tag_confinamiento(tag, head):
    if tag is None or head is None:
        return
    _aplicar_estilo_tag_confinamiento_estribo(tag, head)


def etiquetar_cabezal_trabas_multihost_confinamiento(
    document,
    view,
    group_rebar_ids_ordered,
    tag_map=None,
    wall=None,
    extremo=None,
    n_capas=None,
    conf_type=None,
    anchor_override=None,
    stirrup_head_override=None,
):
    """
    Una ``IndependentTag`` multihost para trabas (Tipo 1 / Tipo 2 @ 3+ capas).

    Familia ``EST_A_STRUCTURAL REBAR TAG_CONFINAMIENTO_MULTI HOST``.
    ``AddReferences`` (patrón Borde Losa). Sin fallback a etiquetas sueltas.
    """
    if document is None or view is None:
        return False, u"Etiqueta multihost: parámetros inválidos."
    if not _vista_permite_rebar_tags(view):
        return False, (
            u"Etiqueta multihost: use planta, alzado o sección (no plantilla ni 3D)."
        )
    ids = list(group_rebar_ids_ordered or [])
    if not ids:
        return False, u"Etiqueta multihost: grupo vacío."
    family = CABEZAL_CONFINEMENT_MULTIHOST_TAG_FAMILY_NAME
    if tag_map is None:
        tag_map = collect_confinement_multihost_tag_symbol_map(document)
    if not tag_map:
        return False, (
            u"Etiqueta multihost: no hay tipos OST_RebarTags para «{0}».".format(
                family,
            )
        )
    if _rebar_reference_candidates_for_tag is None:
        return False, u"Etiqueta multihost: módulo de referencias no disponible."

    primary_id = ids[-1]
    primary_rb = document.GetElement(primary_id)
    if not isinstance(primary_rb, Rebar):
        return False, u"Etiqueta multihost: barra principal no encontrada."

    sym, shape_lbl = _resolve_tag_symbol_for_rebar(
        document, family, tag_map, primary_rb,
    )
    if sym is None:
        shape_hint = shape_lbl or _primary_rebar_shape_tag_key(document, primary_rb) or u"?"
        return False, (
            u"Etiqueta multihost: sin tipo para RebarShape «{0}» en «{1}».".format(
                shape_hint, family,
            )
        )
    all_rebars = []
    for rid in ids:
        rb_i = document.GetElement(rid)
        if isinstance(rb_i, Rebar):
            all_rebars.append(rb_i)
    head = _calcular_cabeza_tag_confinamiento_estribo(
        document,
        view,
        primary_rb,
        wall,
        extremo,
        multihost=True,
        rebar_list=all_rebars,
        anchor_override=anchor_override,
        stirrup_head_override=stirrup_head_override,
        n_capas=n_capas,
        conf_type=conf_type,
    )
    if head is not None and anchor_override is not None:
        try:
            head = XYZ(
                float(head.X),
                float(head.Y),
                float(anchor_override.Z),
            )
        except Exception:
            pass
    tag, err = _crear_tag_rebar_confinamiento(
        document, view, primary_rb, sym.Id, head,
    )
    if tag is None:
        return False, err or u"no se pudo crear IndependentTag"

    if len(ids) < 2:
        _reposicionar_cabezera_tag_confinamiento(tag, head)
        return True, None

    add_fn = getattr(tag, "AddReferences", None)
    if add_fn is None:
        return False, (
            u"Revit no expone AddReferences en esta versión; "
            u"use etiqueta MULTI HOST manual."
        )

    extra_refs = _refs_multihost_para_rebars(document, view, ids[:-1])
    if not extra_refs:
        return False, u"sin referencias para hosts adicionales de traba"

    refs_add = List[Reference]()
    for ref in extra_refs:
        refs_add.Add(ref)

    try:
        document.Regenerate()
    except Exception:
        pass

    st = SubTransaction(document)
    try:
        st.Start()
    except Exception:
        return False, u"no se pudo abrir SubTransaction para multihost"
    try:
        add_fn(refs_add)
        st.Commit()
        _reposicionar_cabezera_tag_confinamiento(tag, head)
        return True, None
    except Exception as ex_mh:
        try:
            st.RollBack()
        except Exception:
            pass
        try:
            document.Delete(tag.Id)
        except Exception:
            pass
        try:
            msg = unicode(ex_mh)
        except Exception:
            msg = str(ex_mh)
        return False, u"Multihost trabas (AddReferences): {0}".format(msg)


def _flush_cabezal_confinement_multihost_one(document, view, res, _key, grp, tag_map):
    """Una etiqueta multihost (grupo muro/extremo). Retorna (ok, err)."""
    ordered_ids = _ordered_rebar_ids_from_multihost_grp(grp)
    if not ordered_ids:
        return False, u"grupo multihost vacío"
    try:
        n_capas_g = int(grp.get(u"n_capas", 3) or 3)
    except Exception:
        n_capas_g = 3
    conf_type_g = grp.get(u"conf_type")
    anchor_override = None
    stirrup_head_override = None
    cj_stub = {u"wid": _key[0], u"extremo": _key[1]}
    if _confinement_multihost_tipo2_tie(n_capas_g, conf_type_g, u"tie"):
        anchor_override = get_confinement_tipo2_stirrup_anchor(res, cj_stub)
        stirrup_head_override = get_confinement_tipo2_stirrup_head(res, cj_stub)
    return etiquetar_cabezal_trabas_multihost_confinamiento(
        document,
        view,
        ordered_ids,
        tag_map=tag_map,
        wall=grp.get(u"wall"),
        extremo=grp.get(u"extremo"),
        n_capas=n_capas_g,
        conf_type=conf_type_g,
        anchor_override=anchor_override,
        stirrup_head_override=stirrup_head_override,
    )


def _record_conf_multihost_tag_result(res, _key, ok, err):
    if ok:
        res[u"n_conf_tags_created"] = int(res.get(u"n_conf_tags_created", 0)) + 1
        return
    res[u"n_conf_tags_fail"] = int(res.get(u"n_conf_tags_fail", 0)) + 1
    msgs = res.get(u"messages")
    if msgs is not None and len(msgs) < 24 and err:
        try:
            wid = int(_key[0])
        except Exception:
            wid = _key[0]
        msgs.append(
            u"Muro {0} {1} etiqueta trabas multihost: {2}".format(
                wid, _key[1] or u"", err,
            ),
        )


def flush_cabezal_confinement_multihost_trabas(document, view, res):
    """
    Crea etiquetas multihost pendientes (3+ capas) y actualiza ``res``.

    Debe llamarse dentro de una ``Transaction`` abierta.
    """
    pending = res.pop(u"_conf_multihost_traba_pending", None) or {}
    res.pop(u"_conf_multihost_flush_order", None)
    if not pending or document is None or view is None:
        return
    tag_map = res.get(u"_conf_multihost_tag_map")
    if tag_map is None:
        try:
            tag_map = collect_confinement_multihost_tag_symbol_map(document)
        except Exception:
            tag_map = {}
        res[u"_conf_multihost_tag_map"] = tag_map
    for _key, grp in pending.items():
        ok, err = _flush_cabezal_confinement_multihost_one(
            document, view, res, _key, grp, tag_map,
        )
        _record_conf_multihost_tag_result(res, _key, ok, err)


def etiquetar_cabezal_confinamiento_multihost_animado(
    document,
    view,
    res,
    batch_size=1,
    after_batch=None,
):
    """
    Etiquetas multihost de confinamiento por grupo (muro/extremo), lote a lote.

    Usa ``_conf_multihost_traba_pending`` y orden ``_conf_multihost_flush_order``.
    """
    pending = res.get(u"_conf_multihost_traba_pending") or {}
    order = list(res.get(u"_conf_multihost_flush_order") or [])
    if not order:
        order = list(pending.keys())
    keys = [k for k in order if k in pending]
    for k in pending:
        if k not in keys:
            keys.append(k)
    if not keys or document is None or view is None:
        res.pop(u"_conf_multihost_traba_pending", None)
        res.pop(u"_conf_multihost_flush_order", None)
        return {u"n_ok": 0, u"n_fail": 0}

    tag_map = res.get(u"_conf_multihost_tag_map")
    if tag_map is None:
        try:
            tag_map = collect_confinement_multihost_tag_symbol_map(document)
        except Exception:
            tag_map = {}
        res[u"_conf_multihost_tag_map"] = tag_map

    batch_n = max(1, int(batch_size or 1))
    n_ok = 0
    n_fail = 0
    for lote_idx in range(0, len(keys), batch_n):
        lote_keys = keys[lote_idx:lote_idx + batch_n]
        if len(lote_keys) == 1:
            txn_name = u"Arainco: Cabezal muros — etiqueta multihost confinamiento"
        else:
            txn_name = (
                u"Arainco: Cabezal muros — etiquetas multihost confinamiento "
                u"{0}–{1}".format(lote_idx + 1, lote_idx + len(lote_keys))
            )
        t = Transaction(document, txn_name)
        t.Start()
        lote_ok = False
        try:
            for _key in lote_keys:
                grp = pending.pop(_key, None)
                if grp is None:
                    continue
                ok, err = _flush_cabezal_confinement_multihost_one(
                    document, view, res, _key, grp, tag_map,
                )
                _record_conf_multihost_tag_result(res, _key, ok, err)
                if ok:
                    n_ok += 1
                else:
                    n_fail += 1
            t.Commit()
            lote_ok = True
        except Exception as ex_lote:
            try:
                if t.HasStarted():
                    t.RollBack()
            except Exception:
                pass
            n_fail += len(lote_keys)
            msgs = res.get(u"messages")
            if msgs is not None and len(msgs) < 24:
                msgs.append(u"Multihost confinamiento lote: {0}".format(ex_lote))
        if lote_ok and after_batch is not None:
            try:
                after_batch()
            except Exception:
                pass

    res.pop(u"_conf_multihost_traba_pending", None)
    res.pop(u"_conf_multihost_flush_order", None)
    return {u"n_ok": n_ok, u"n_fail": n_fail}


