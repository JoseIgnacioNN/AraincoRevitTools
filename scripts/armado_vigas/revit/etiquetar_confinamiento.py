# -*- coding: utf-8 -*-
"""
Etiquetado de estribos/trabas de confinamiento (Armado vigas).

Familias de etiqueta (Armado vigas):

- ``EST_A_STRUCTURAL REBAR TAG_CONFINAMIENTO_VIGA`` — estribo o traba individual.
- ``EST_A_STRUCTURAL REBAR TAG_CONFINAMIENTO_VIGA_MULTI HOST`` — varias trabas en la misma viga
  o varios estribos del mismo tramo extremo/central en la misma viga.

El tipo dentro de cada familia se resuelve por ``RebarShape`` de la barra.

En sección/alzado (mismo criterio que **Armado columnas**):

- Cotas **lineales horizontales** (tipo «Linear - Confinamiento») **debajo** de la viga
  (offset 500 mm respecto al borde inferior visible del lote; 850 mm si hay cota de
  empalme inferior en la misma viga).
- Cotas de **empalme inferior**: a 500 mm bajo el borde inferior (más cerca del armado);
  la cota de confinamiento queda 350 mm más abajo (``LAP_CONFINEMENT_DIM_GAP_INF_MM``).
- Etiquetas **sin leader**, orientación **horizontal**, cabecera en el **centro del lote**
  sobre la misma línea Y que la cota.
"""

from __future__ import print_function

import os
import sys

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    BuiltInParameter,
    DimensionStyleType,
    DimensionType,
    ElementType,
    FilteredElementCollector,
    IndependentTag,
    Line,
    ReferenceArray,
    StorageType,
    TagMode,
    TagOrientation,
    ViewDrafting,
    ViewSection,
    ViewType,
    XYZ,
)
from Autodesk.Revit.DB.Structure import Rebar
from System.Collections.Generic import List

STIRRUP_TAG_OFFSET_BELOW_BEAM_MM = 500.0
LAP_CONFINEMENT_DIM_GAP_INF_MM = 350.0
_INFERIOR_LAP_DIM_HOST_IDS = set()


def reset_inferior_lap_dim_host_registry():
    """Limpia vigas con cota de empalme inferior (inicio de corrida Armado vigas)."""
    _INFERIOR_LAP_DIM_HOST_IDS.clear()


def register_inferior_lap_dim_host(host):
    """Marca la viga para apilar la cota de confinamiento bajo la de empalme inferior."""
    if host is None:
        return
    try:
        _INFERIOR_LAP_DIM_HOST_IDS.add(int(host.Id.IntegerValue))
    except Exception:
        pass


def _host_has_inferior_lap_dim(host):
    if host is None:
        return False
    try:
        return int(host.Id.IntegerValue) in _INFERIOR_LAP_DIM_HOST_IDS
    except Exception:
        return False


CONFINEMENT_DIM_MARKER_LENGTH_MM = 80.0
_FIXED_DIMSTYLE_NAME = u"Linear - Confinamiento"
_FIXED_DIMSTYLE_CACHE = {}

try:
    from confinement_dim_link_schema import set_confinement_dim_marker_link
except Exception:
    set_confinement_dim_marker_link = None

try:
    from confinement_dim_updater_dmu import ensure_confinement_dim_link_updater_registered
except Exception:
    ensure_confinement_dim_link_updater_registered = None

CONFINEMENT_TAG_FAMILY = u"EST_A_STRUCTURAL REBAR TAG_CONFINAMIENTO_VIGA"
CONFINEMENT_MULTIHOST_TAG_FAMILY = (
    u"EST_A_STRUCTURAL REBAR TAG_CONFINAMIENTO_VIGA_MULTI HOST"
)

_VIGA_EXTREMO_GROUP = u"viga"
_CONF_TYPE_PERIMETER = u"perimeter_0_1"
_CONF_TYPE_TIE_LAYER_1 = u"tie_layer_1"


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _is_section_or_elevation_view(view):
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        if isinstance(view, ViewDrafting):
            return False
    except Exception:
        pass
    try:
        if isinstance(view, ViewSection):
            return True
    except Exception:
        pass
    try:
        vt = view.ViewType
        return vt == ViewType.Section or vt == ViewType.Elevation
    except Exception:
        return False


def _dot_xyz(a, b):
    return (
        float(a.X) * float(b.X)
        + float(a.Y) * float(b.Y)
        + float(a.Z) * float(b.Z)
    )


def _view_local_xy(view, pt):
    o = view.Origin
    d = XYZ(
        float(pt.X) - float(o.X),
        float(pt.Y) - float(o.Y),
        float(pt.Z) - float(o.Z),
    )
    return _dot_xyz(d, view.RightDirection), _dot_xyz(d, view.UpDirection)


def _view_local_to_world(view, ref_world, x_loc, y_loc):
    x_ref, y_ref = _view_local_xy(view, ref_world)
    rd = view.RightDirection
    up = view.UpDirection
    return ref_world + rd.Multiply(float(x_loc - x_ref)) + up.Multiply(
        float(y_loc - y_ref)
    )


def _bbox_corners_world(bb):
    if bb is None:
        return []
    try:
        tr = bb.Transform
        mn, mx = bb.Min, bb.Max
        out = []
        for x in (float(mn.X), float(mx.X)):
            for y in (float(mn.Y), float(mx.Y)):
                for z in (float(mn.Z), float(mx.Z)):
                    out.append(tr.OfPoint(XYZ(x, y, z)))
        return out
    except Exception:
        try:
            return [bb.Min, bb.Max]
        except Exception:
            return []


def _element_bounding_box(elem, view):
    if elem is None:
        return None
    try:
        bb = elem.get_BoundingBox(view)
        if bb is not None:
            return bb
    except Exception:
        pass
    try:
        return elem.get_BoundingBox(None)
    except Exception:
        return None


def _min_view_local_up(view, elements):
    min_y = None
    for el in elements or []:
        for pt in _bbox_corners_world(_element_bounding_box(el, view)):
            _, y_loc = _view_local_xy(view, pt)
            if min_y is None or y_loc < min_y:
                min_y = float(y_loc)
    return min_y


def _confinement_annotation_y_dim(view, host, rebars=None):
    elements = []
    if host is not None:
        elements.append(host)
    for rb in rebars or []:
        if rb is not None:
            elements.append(rb)
    min_y = _min_view_local_up(view, elements)
    if min_y is None:
        min_y = 0.0
    offset_mm = float(STIRRUP_TAG_OFFSET_BELOW_BEAM_MM)
    if _host_has_inferior_lap_dim(host):
        offset_mm += float(LAP_CONFINEMENT_DIM_GAP_INF_MM)
    return min_y - _mm_to_ft(offset_mm)


def compute_inferior_lap_dim_offset_mm(view, host, lap_mid_world):
    """
    Offset en mm (modelo) para la cota de empalme en cara inferior.

    La línea de cota queda a ``STIRRUP_TAG_OFFSET_BELOW_BEAM_MM`` (500 mm) bajo
    el borde inferior visible del host — más cerca del armado. La cota de
    confinamiento en la misma viga se coloca ``LAP_CONFINEMENT_DIM_GAP_INF_MM``
    (350 mm) más abajo (850 mm total), medido en la dirección Up de la vista.
    """
    if view is None or host is None or lap_mid_world is None:
        return None
    min_y = _min_view_local_up(view, [host])
    if min_y is None:
        return None
    _, y_rebar = _view_local_xy(view, lap_mid_world)
    y_lap = float(min_y) - float(_mm_to_ft(STIRRUP_TAG_OFFSET_BELOW_BEAM_MM))
    delta_ft = float(y_rebar) - y_lap
    if delta_ft < _mm_to_ft(20.0):
        return None
    return delta_ft * 304.8


def _beam_reference_xyz(host, view):
    bb = _element_bounding_box(host, view)
    if bb is not None:
        try:
            return (bb.Min + bb.Max) * 0.5
        except Exception:
            pass
    try:
        from Autodesk.Revit.DB import LocationCurve

        loc = getattr(host, "Location", None)
        if isinstance(loc, LocationCurve) and loc.Curve is not None:
            return loc.Curve.Evaluate(0.5, True)
    except Exception:
        pass
    return XYZ(0.0, 0.0, 0.0)


def _rebar_center(rebar, view):
    try:
        bb = rebar.get_BoundingBox(view)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    try:
        bb = rebar.get_BoundingBox(None)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    return None


def _mean_view_local_x(view, ref_world, rebars):
    xs = []
    for rb in rebars or []:
        if not isinstance(rb, Rebar):
            continue
        p = _rebar_center(rb, view)
        if p is None:
            continue
        x_mid, _ = _view_local_xy(view, p)
        xs.append(float(x_mid))
    if not xs:
        return None
    return sum(xs) / float(len(xs))


def _rebar_array_extent_view_local_x(view, rebar):
    if rebar is None:
        return None, None
    try:
        n_pos = int(rebar.NumberOfBarPositions)
    except Exception:
        n_pos = 1
    if n_pos < 1:
        n_pos = 1
    origins = []
    for idx in (0, max(0, n_pos - 1)):
        try:
            tr = rebar.GetBarPositionTransform(int(idx))
            if tr is not None:
                origins.append(tr.Origin)
        except Exception:
            pass
    try:
        acc = rebar.GetShapeDrivenAccessor()
        alen = float(acc.ArrayLength)
    except Exception:
        alen = 0.0
    if len(origins) >= 2 and alen > 1e-9:
        o0, o1 = origins[0], origins[-1]
        try:
            delta = o1.Subtract(o0)
            ln = float(delta.GetLength())
            if ln > 1e-9:
                axis = delta.Multiply(1.0 / ln)
                o0 = o0.Add(axis.Multiply(-0.5 * alen))
                o1 = o0.Add(axis.Multiply(alen))
                origins = [o0, o1]
        except Exception:
            pass
    elif len(origins) == 1 and alen > 1e-9:
        o0 = origins[0]
        try:
            path = rebar.GetDistributionPath()
            if path is not None and int(path.Count) >= 1:
                p0 = path[0].GetEndPoint(0)
                p1 = path[int(path.Count) - 1].GetEndPoint(1)
                delta = p1.Subtract(p0)
                ln = float(delta.GetLength())
                if ln > 1e-9:
                    axis = delta.Multiply(1.0 / ln)
                    origins = [o0, o0.Add(axis.Multiply(alen))]
        except Exception:
            pass
    if not origins:
        return None, None
    xs = []
    for pt in origins:
        x_loc, _ = _view_local_xy(view, pt)
        xs.append(float(x_loc))
    if not xs:
        return None, None
    return min(xs), max(xs)


def _segment_view_local_x_extent(view, rebars):
    x_min = None
    x_max = None
    for rb in rebars or []:
        if not isinstance(rb, Rebar):
            continue
        bb = _element_bounding_box(rb, view)
        for pt in _bbox_corners_world(bb):
            x_loc, _ = _view_local_xy(view, pt)
            if x_min is None or x_loc < x_min:
                x_min = float(x_loc)
            if x_max is None or x_loc > x_max:
                x_max = float(x_loc)
    if x_min is not None and x_max is not None and (x_max - x_min) > _mm_to_ft(5.0):
        return x_min, x_max
    for rb in rebars or []:
        if not isinstance(rb, Rebar):
            continue
        x0, x1 = _rebar_array_extent_view_local_x(view, rb)
        if x0 is None or x1 is None:
            continue
        lo = min(float(x0), float(x1))
        hi = max(float(x0), float(x1))
        if x_min is None or lo < x_min:
            x_min = lo
        if x_max is None or hi > x_max:
            x_max = hi
    if x_min is None or x_max is None:
        return None, None
    return x_min, x_max


def _create_vertical_dim_marker_detailcurve(doc, view, center_world, length_mm):
    if doc is None or view is None or center_world is None:
        return None, None
    try:
        up = view.UpDirection
        if up is None or float(up.GetLength()) < 1e-12:
            return None, None
        up = up.Normalize()
        half = 0.5 * _mm_to_ft(float(length_mm))
        p0 = center_world.Add(up.Multiply(-half))
        p1 = center_world.Add(up.Multiply(half))
        ln = Line.CreateBound(p0, p1)
        dc = doc.Create.NewDetailCurve(view, ln)
        if dc is None:
            return None, None
        try:
            return dc, dc.GeometryCurve.Reference
        except Exception:
            return dc, None
    except Exception:
        return None, None


def _dimension_type_name_candidates(dt):
    out = []
    if dt is None:
        return out
    try:
        v = getattr(dt, "Name", None)
        if v:
            out.append(unicode(v).strip())
    except Exception:
        pass
    for bip in (BuiltInParameter.SYMBOL_NAME_PARAM, BuiltInParameter.ALL_MODEL_TYPE_NAME):
        try:
            p = dt.get_Parameter(bip)
            if p is not None and p.HasValue and p.StorageType == StorageType.String:
                s = unicode(p.AsString() or u"").strip()
                if s:
                    out.append(s)
        except Exception:
            pass
    seen = set()
    uniq = []
    for s in out:
        k = s.lower()
        if k and k not in seen:
            seen.add(k)
            uniq.append(s)
    return uniq


def _is_linear_dimension_type(dt):
    if dt is None:
        return False
    try:
        return dt.StyleType == DimensionStyleType.Linear
    except Exception:
        return True


def _collect_linear_dimension_types(doc):
    out = []
    seen = set()
    try:
        for dt in FilteredElementCollector(doc).OfClass(DimensionType):
            if dt is None or not _is_linear_dimension_type(dt):
                continue
            try:
                k = int(dt.Id.IntegerValue)
            except Exception:
                k = id(dt)
            if k in seen:
                continue
            seen.add(k)
            out.append(dt)
    except Exception:
        pass
    return out


def _get_fixed_dimension_type_id(doc):
    if doc is None:
        return None
    try:
        key = int(doc.GetHashCode())
    except Exception:
        key = id(doc)
    cached = _FIXED_DIMSTYLE_CACHE.get(key)
    if cached is not None:
        return cached
    target = _FIXED_DIMSTYLE_NAME.lower()
    for dt in _collect_linear_dimension_types(doc):
        for nm in _dimension_type_name_candidates(dt):
            if nm.lower() == target:
                _FIXED_DIMSTYLE_CACHE[key] = dt.Id
                return dt.Id
    _FIXED_DIMSTYLE_CACHE[key] = None
    return None


def _try_apply_fixed_dimension_type(doc, dim):
    dim_type_id = _get_fixed_dimension_type_id(doc)
    if dim is None or dim_type_id is None:
        return
    try:
        dim.DimensionTypeId = dim_type_id
    except Exception:
        try:
            dim.ChangeTypeId(dim_type_id)
        except Exception:
            pass


def _create_horizontal_confinement_dimension(doc, view, ref_world, y_dim, x_min, x_max):
    if doc is None or view is None or ref_world is None:
        return False, u"sin documento/vista"
    if x_min is None or x_max is None:
        return False, u"sin extensión horizontal del lote"
    try:
        x_lo = float(x_min)
        x_hi = float(x_max)
    except Exception:
        return False, u"extensión horizontal inválida"
    if abs(x_hi - x_lo) < _mm_to_ft(10.0):
        return False, u"lote demasiado corto para cota"
    if x_lo > x_hi:
        x_lo, x_hi = x_hi, x_lo

    pt_left = _view_local_to_world(view, ref_world, x_lo, float(y_dim))
    pt_right = _view_local_to_world(view, ref_world, x_hi, float(y_dim))
    dc_left, ref_left = _create_vertical_dim_marker_detailcurve(
        doc, view, pt_left, CONFINEMENT_DIM_MARKER_LENGTH_MM,
    )
    dc_right, ref_right = _create_vertical_dim_marker_detailcurve(
        doc, view, pt_right, CONFINEMENT_DIM_MARKER_LENGTH_MM,
    )
    if ref_left is None or ref_right is None:
        for dc in (dc_left, dc_right):
            if dc is not None:
                try:
                    doc.Delete(dc.Id)
                except Exception:
                    pass
        return False, u"sin referencias de marcadores para cota de confinamiento"

    try:
        dim_line = Line.CreateBound(pt_left, pt_right)
        ra = ReferenceArray()
        ra.Append(ref_left)
        ra.Append(ref_right)
        dim = doc.Create.NewDimension(view, dim_line, ra)
    except Exception as ex:
        dim = None
        err = ex
    else:
        err = None

    if dim is None:
        for dc in (dc_left, dc_right):
            if dc is not None:
                try:
                    doc.Delete(dc.Id)
                except Exception:
                    pass
        try:
            return False, unicode(err)
        except Exception:
            return False, str(err)

    _try_apply_fixed_dimension_type(doc, dim)
    if ensure_confinement_dim_link_updater_registered is not None:
        try:
            ensure_confinement_dim_link_updater_registered()
        except Exception:
            pass
    if set_confinement_dim_marker_link is not None:
        for dc in (dc_left, dc_right):
            if dc is not None:
                try:
                    set_confinement_dim_marker_link(dc, dim.Id, view.Id)
                except Exception:
                    pass
    return True, None


def _viga_conf_lot_key(job, multihost=False):
    try:
        hid = int(job.get(u"host_id") or 0)
    except Exception:
        hid = 0
    if job.get(u"job_kind") == u"stirrup":
        try:
            zi = int(job.get(u"stirrup_zone_index", 0) or 0)
        except Exception:
            zi = 0
        return (hid, u"stirrup_mh" if multihost else u"stirrup", zi)
    try:
        ti = int(job.get(u"tie_index") or 0)
    except Exception:
        ti = 0
    return (hid, u"tie_mh" if multihost else u"tie", ti)


def _register_viga_conf_lot(registry, lot_key, host, rebar):
    if lot_key is None or rebar is None:
        return
    entry = registry.get(lot_key)
    if entry is None:
        entry = {u"host": host, u"rebars": []}
        registry[lot_key] = entry
    if host is not None and entry.get(u"host") is None:
        entry[u"host"] = host
    rebars = entry.setdefault(u"rebars", [])
    try:
        rid = int(rebar.Id.IntegerValue)
    except Exception:
        rid = id(rebar)
    seen = entry.setdefault(u"_seen", set())
    if rid in seen:
        return
    seen.add(rid)
    rebars.append(rebar)


def _compute_viga_lot_layout(view, host, rebars):
    if view is None or not rebars:
        return None
    if not _is_section_or_elevation_view(view):
        return None
    ref_world = _beam_reference_xyz(host, view)
    y_dim = _confinement_annotation_y_dim(view, host, rebars)
    x_min, x_max = _segment_view_local_x_extent(view, rebars)
    if x_min is None or x_max is None:
        return None
    x_lot = _mean_view_local_x(view, ref_world, rebars)
    x_loc = float(x_lot) if x_lot is not None else 0.5 * (float(x_min) + float(x_max))
    return {
        u"ref_world": ref_world,
        u"y_dim": float(y_dim),
        u"x_min": float(x_min),
        u"x_max": float(x_max),
        u"x_loc": float(x_loc),
    }


def _head_from_viga_lot_layout(view, layout):
    if view is None or not layout:
        return None
    ref_world = layout.get(u"ref_world")
    if ref_world is None:
        return None
    return _view_local_to_world(
        view,
        ref_world,
        float(layout.get(u"x_loc", 0.0)),
        float(layout.get(u"y_dim", 0.0)),
    )


def _build_viga_conf_lot_layouts(view, lot_registry):
    layouts = {}
    for lot_key, entry in (lot_registry or {}).items():
        host = entry.get(u"host")
        rebars = entry.get(u"rebars") or []
        layout = _compute_viga_lot_layout(view, host, rebars)
        if layout is not None:
            layouts[lot_key] = layout
    return layouts


def _create_viga_confinement_dimensions(document, view, layouts, avisos):
    n_dims = 0
    for lot_key, layout in (layouts or {}).items():
        ok, err = _create_horizontal_confinement_dimension(
            document,
            view,
            layout.get(u"ref_world"),
            layout.get(u"y_dim"),
            layout.get(u"x_min"),
            layout.get(u"x_max"),
        )
        if ok:
            n_dims += 1
        elif err and avisos is not None and len(avisos) < 12:
            avisos.append(u"Cota confinamiento lote {0}: {1}".format(lot_key, err))
    return n_dims


def _load_cabezal_tags_module():
    if "armado_muros_cabezal_tags" in sys.modules:
        return sys.modules["armado_muros_cabezal_tags"]
    try:
        import armado_muros_cabezal_tags as mod
        return mod
    except Exception:
        pass
    here = os.path.dirname(os.path.abspath(__file__))
    scripts_dir = os.path.dirname(os.path.dirname(here))
    ext_root = os.path.dirname(scripts_dir)
    candidates = [
        os.path.join(
            ext_root,
            "BIMTools.tab",
            "Armadura.panel",
            "34_ArmadoMuros.pushbutton",
            "scripts",
        ),
    ]
    cursor = ext_root
    for _ in range(8):
        candidates.append(
            os.path.join(
                cursor,
                "BIMTools.tab",
                "Armadura.panel",
                "34_ArmadoMuros.pushbutton",
                "scripts",
            )
        )
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    seen = set()
    for d in candidates:
        if not d or d in seen:
            continue
        seen.add(d)
        mod_path = os.path.join(d, "armado_muros_cabezal_tags.py")
        if not os.path.isfile(mod_path):
            continue
        if d not in sys.path:
            sys.path.insert(0, d)
        try:
            import armado_muros_cabezal_tags as mod
            return mod
        except Exception:
            continue
    return None


def _collect_tag_symbol_map(document, family_name, tags_mod):
    """Mapa tipo (``RebarShape``) → ``FamilySymbol`` para ``family_name``."""
    if tags_mod is not None:
        try:
            return tags_mod._collect_tag_symbol_map(document, family_name)
        except Exception:
            pass
    try:
        from enfierrado_shaft_hashtag import _collect_rebar_tag_symbol_map

        return _collect_rebar_tag_symbol_map(document, family_name)
    except Exception:
        return {}


def _norm_conf(conf):
    return conf if isinstance(conf, dict) else {}


try:
    from enfierrado_shaft_hashtag import (
        _primary_rebar_shape_tag_key,
        _rebar_reference_candidates_for_tag,
        _rebar_shape_name_candidates,
    )
except Exception:
    _primary_rebar_shape_tag_key = None
    _rebar_reference_candidates_for_tag = None
    _rebar_shape_name_candidates = None


def _midpoint_centerline_curves(curves):
    if not curves:
        return None
    try:
        c = max(curves, key=lambda cv: float(cv.Length))
        return c.Evaluate(0.5, True)
    except Exception:
        return None


def _distribution_dir_from_host(host, near_pt):
    """Tangente del eje de viga en la estación más cercana a ``near_pt``."""
    curve = _host_location_curve(host)
    if curve is None:
        return None
    try:
        if near_pt is not None:
            proj = curve.Project(near_pt)
            if proj is not None:
                deriv = curve.ComputeDerivatives(proj.Parameter, True)
                t = deriv.BasisX
                if t is not None and float(t.GetLength()) > 1e-9:
                    return t.Normalize()
    except Exception:
        pass
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        d = p1.Subtract(p0)
        if float(d.GetLength()) > 1e-9:
            return d.Normalize()
    except Exception:
        pass
    return None


def _distribution_dir_rebar(rebar, host=None, near_pt=None):
    """Dirección del reparto del set (eje de ``ArrayLength``)."""
    if rebar is not None:
        try:
            path = rebar.GetDistributionPath()
            if path is not None and int(path.Count) > 0:
                p0 = path[0].GetEndPoint(0)
                p1 = path[int(path.Count) - 1].GetEndPoint(1)
                d = p1.Subtract(p0)
                if float(d.GetLength()) > 1e-9:
                    return d.Normalize()
        except Exception:
            pass
    if host is not None:
        ref = near_pt
        if ref is None and rebar is not None:
            try:
                bb = rebar.get_BoundingBox(None)
                if bb is not None:
                    ref = (bb.Min + bb.Max) * 0.5
            except Exception:
                pass
        return _distribution_dir_from_host(host, ref)
    return None


def _rebar_array_length_ft(rebar):
    if rebar is None:
        return 0.0
    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception:
        acc = None
    if acc is None:
        return 0.0
    try:
        return float(acc.ArrayLength)
    except Exception:
        try:
            return float(acc.GetArrayLength())
        except Exception:
            return 0.0


def _distribution_path_midpoint(rebar):
    """Centro del ``ArrayLength``: punto medio del tramo de reparto del set."""
    if rebar is None:
        return None
    try:
        path = rebar.GetDistributionPath()
        if path is None or int(path.Count) < 1:
            return None
        p0 = path[0].GetEndPoint(0)
        p1 = path[int(path.Count) - 1].GetEndPoint(1)
        try:
            if float(p0.DistanceTo(p1)) < 1e-9:
                return None
        except Exception:
            pass
        return XYZ(
            0.5 * (float(p0.X) + float(p1.X)),
            0.5 * (float(p0.Y) + float(p1.Y)),
            0.5 * (float(p0.Z) + float(p1.Z)),
        )
    except Exception:
        return None


def _arraylength_center_from_bar_transforms(rebar, host=None):
    """
    Centro del layout: origen barra 0 + mitad de ``ArrayLength`` en eje de reparto.

  En sección, ``GetBarPositionTransform`` suele devolver el mismo origen para todas
    las posiciones; entonces se usa el eje de la viga host.
    """
    if rebar is None:
        return None
    alen = _rebar_array_length_ft(rebar)
    if alen < 1e-12:
        return None
    try:
        tr0 = rebar.GetBarPositionTransform(0)
    except Exception:
        tr0 = None
    if tr0 is None:
        return None
    o0 = tr0.Origin
    n_pos = 1
    try:
        n_pos = int(rebar.NumberOfBarPositions)
    except Exception:
        n_pos = 1
    axis = None
    if n_pos > 1:
        try:
            trn = rebar.GetBarPositionTransform(n_pos - 1)
            if trn is not None:
                o1 = trn.Origin
                delta = o1.Subtract(o0)
                ln = float(delta.GetLength())
                if ln > 1e-9:
                    axis = delta.Multiply(1.0 / ln)
        except Exception:
            pass
    if axis is None:
        axis = _distribution_dir_rebar(rebar, host, o0)
    if axis is None:
        return None
    try:
        return o0.Add(axis.Multiply(0.5 * float(alen)))
    except Exception:
        return None


def _centro_arraylength_desde_bbox(rebar, host=None):
    """
    Centro del tramo de reparto proyectando el bbox 3D sobre el eje de distribución.

    Fiable en sección cuando ``GetCenterlineCurves`` no distingue posiciones del set.
    """
    if rebar is None:
        return None
    try:
        bb = rebar.get_BoundingBox(None)
        if bb is None:
            return None
        c_bb = (bb.Min + bb.Max) * 0.5
        axis = _distribution_dir_rebar(rebar, host, c_bb)
        if axis is None:
            return c_bb
        mn = bb.Min
        mx = bb.Max
        corners = (
            XYZ(mn.X, mn.Y, mn.Z),
            XYZ(mx.X, mn.Y, mn.Z),
            XYZ(mn.X, mx.Y, mn.Z),
            XYZ(mx.X, mx.Y, mn.Z),
            XYZ(mn.X, mn.Y, mx.Z),
            XYZ(mx.X, mn.Y, mx.Z),
            XYZ(mn.X, mx.Y, mx.Z),
            XYZ(mx.X, mx.Y, mx.Z),
        )
        scalars = [float((p - c_bb).DotProduct(axis)) for p in corners]
        s_mid = 0.5 * (min(scalars) + max(scalars))
        return c_bb.Add(axis.Multiply(s_mid))
    except Exception:
        return None


def _centro_por_extremos_posiciones(rebar):
    """Mitad entre centros 1ª/última posición; ``None`` si Revit devuelve el mismo punto."""
    if rebar is None:
        return None
    try:
        from Autodesk.Revit.DB.Structure import MultiplanarOption, Rebar

        if not isinstance(rebar, Rebar):
            return None
        n_pos = 1
        try:
            n_pos = int(rebar.NumberOfBarPositions)
        except Exception:
            n_pos = 1
        if n_pos <= 1:
            return None
        for mpo_name in (u"IncludeAllMultiplanarCurves", u"IncludeOnlyPlanarCurves"):
            try:
                mpo = getattr(MultiplanarOption, mpo_name)
            except Exception:
                mpo = None
            if mpo is None:
                continue
            try:
                cs0 = list(rebar.GetCenterlineCurves(False, False, False, mpo, 0))
                csn = list(
                    rebar.GetCenterlineCurves(False, False, False, mpo, n_pos - 1)
                )
                p0 = _midpoint_centerline_curves(cs0)
                pn = _midpoint_centerline_curves(csn)
                if p0 is None or pn is None:
                    continue
                try:
                    if float(p0.DistanceTo(pn)) < (25.0 / 304.8):
                        continue
                except Exception:
                    pass
                return XYZ(
                    0.5 * (float(p0.X) + float(pn.X)),
                    0.5 * (float(p0.Y) + float(pn.Y)),
                    0.5 * (float(p0.Z) + float(pn.Z)),
                )
            except Exception:
                continue
    except Exception:
        pass
    return None


def _resolve_array_midpoint(rebar, host=None):
    """Punto medio del ``ArrayLength`` del set."""
    if rebar is None:
        return None
    for fn in (
        lambda: _distribution_path_midpoint(rebar),
        lambda: _centro_arraylength_desde_bbox(rebar, host),
        lambda: _arraylength_center_from_bar_transforms(rebar, host),
        lambda: _centro_por_extremos_posiciones(rebar),
    ):
        try:
            pt = fn()
        except Exception:
            pt = None
        if pt is not None:
            return pt
    return None


def _punto_medio_arraylength_rebar(rebar, host=None):
    """Punto medio del tramo de reparto / ``ArrayLength`` (sin ajuste de sección)."""
    return _resolve_array_midpoint(rebar, host=host)


def _resolve_rebar_host(document, rebar, host=None):
    if host is not None:
        return host
    if document is None or rebar is None:
        return None
    try:
        hid = rebar.GetHostId()
        if hid is None or int(hid.IntegerValue) <= 0:
            return None
        return document.GetElement(hid)
    except Exception:
        return None


def _host_location_curve(host):
    if host is None:
        return None
    try:
        from Autodesk.Revit.DB import LocationCurve

        loc = getattr(host, "Location", None)
        if isinstance(loc, LocationCurve):
            return loc.Curve
    except Exception:
        pass
    return None


def _centro_seccion_viga_en_estacion(host, array_mid):
    """
    Conserva la estación del ``ArrayLength`` (``array_mid``) y centra solo el canto
    (mitad entre caras superior e inferior del host en esa estación).
    """
    if host is None or array_mid is None:
        return None
    curve = _host_location_curve(host)
    if curve is None:
        return None
    try:
        from armado_vigas.revit.colocar_rebar import _cara_framing
        from armado_vigas.revit.colocar_estribos import (
            _beam_layout_axes,
            _project_xyz_on_face,
        )
    except Exception:
        return None
    try:
        proj = curve.Project(array_mid)
        probe = proj.XYZPoint if proj is not None else array_mid
    except Exception:
        probe = array_mid
    try:
        _, depth_dir = _beam_layout_axes(host, curve)
    except Exception:
        depth_dir = None
    if depth_dir is None:
        return None
    sup = _cara_framing(host, False)
    inf = _cara_framing(host, True)
    if sup is None or inf is None:
        return None
    pt_top = _project_xyz_on_face(sup, probe)
    pt_bot = _project_xyz_on_face(inf, probe)
    if pt_top is None or pt_bot is None:
        return None
    try:
        d_top = float((pt_top - probe).DotProduct(depth_dir))
        d_bot = float((pt_bot - probe).DotProduct(depth_dir))
        d_mid = 0.5 * (d_top + d_bot)
        d_array = float((array_mid - probe).DotProduct(depth_dir))
        return array_mid.Add(depth_dir.Multiply(d_mid - d_array))
    except Exception:
        return XYZ(
            0.5 * (float(pt_top.X) + float(pt_bot.X)),
            0.5 * (float(pt_top.Y) + float(pt_bot.Y)),
            0.5 * (float(pt_top.Z) + float(pt_bot.Z)),
        )


def _centro_etiqueta_confinamiento_viga(document, view, rebar, host=None):
    """Cabecera en sección: centro del ``ArrayLength`` + centro del canto de viga."""
    if rebar is None:
        return None
    host = _resolve_rebar_host(document, rebar, host)
    array_mid = _resolve_array_midpoint(rebar, host=host)
    head = array_mid
    if array_mid is not None and host is not None:
        head_sec = _centro_seccion_viga_en_estacion(host, array_mid)
        if head_sec is not None:
            head = head_sec
    if head is None:
        try:
            bb = rebar.get_BoundingBox(view)
            if bb is None:
                bb = rebar.get_BoundingBox(None)
            if bb is not None:
                head = (bb.Min + bb.Max) * 0.5
        except Exception:
            pass
    return head


def _centro_array_confinamiento_en_vista(rebar, view, document=None, host=None):
    """Compatibilidad: delega en ``_centro_etiqueta_confinamiento_viga``."""
    return _centro_etiqueta_confinamiento_viga(document, view, rebar, host=host)


def _centro_multihost_rebars_en_vista(document, view, rebar_ids, host=None):
    ids = list(rebar_ids or [])
    if not ids:
        return None
    pts = []
    primary = None
    for rid in ids:
        rb = document.GetElement(rid) if document is not None else None
        if rb is None:
            continue
        if primary is None:
            primary = rb
        p = _punto_medio_arraylength_rebar(rb, host=host)
        if p is not None:
            pts.append(p)
    if not pts:
        return _centro_etiqueta_confinamiento_viga(
            document, view, primary, host=host,
        )
    if len(pts) == 1:
        array_mid = pts[0]
    else:
        try:
            sx = sum(float(p.X) for p in pts)
            sy = sum(float(p.Y) for p in pts)
            sz = sum(float(p.Z) for p in pts)
            n = float(len(pts))
            array_mid = XYZ(sx / n, sy / n, sz / n)
        except Exception:
            array_mid = pts[0]
    host = _resolve_rebar_host(document, primary, host)
    if host is not None:
        head = _centro_seccion_viga_en_estacion(host, array_mid)
        if head is not None:
            return head
    return array_mid


def _resolve_tag_symbol_viga(document, rebar, tag_map, tags_mod):
    if tags_mod is not None:
        try:
            sym, lbl = tags_mod._resolve_tag_symbol_for_rebar(
                document, CONFINEMENT_TAG_FAMILY, tag_map, rebar,
            )
            if sym is not None:
                return sym, lbl
        except Exception:
            pass
    if not tag_map or rebar is None:
        return None, None
    keys = []
    if _primary_rebar_shape_tag_key is not None:
        pk = _primary_rebar_shape_tag_key(document, rebar)
        if pk:
            keys.append(pk)
    if _rebar_shape_name_candidates is not None:
        keys.extend(_rebar_shape_name_candidates(document, rebar))
    seen = set()
    for k in keys:
        if not k or k in seen:
            continue
        seen.add(k)
        sym = tag_map.get(k)
        if sym is not None:
            return sym, k
    return None, None


def _referencias_tag_rebar_viga(document, rebar, view, tags_mod):
    if tags_mod is not None:
        try:
            refs = tags_mod._referencias_tag_rebar(document, rebar, view)
            if refs:
                return refs
        except Exception:
            pass
    if _rebar_reference_candidates_for_tag is not None:
        try:
            return _rebar_reference_candidates_for_tag(document, view, rebar) or []
        except Exception:
            pass
    return []


def _aplicar_estilo_tag_confinamiento_viga(tag, head):
    if tag is None or head is None:
        return
    try:
        tag.TagOrientation = TagOrientation.Horizontal
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


def _crear_tag_confinamiento_sin_leader(document, view, rebar, tag_type_id, head, tags_mod):
    if head is None:
        return None, u"sin punto de inserción"
    if tags_mod is not None:
        try:
            tag, err = tags_mod._crear_tag_rebar_confinamiento(
                document, view, rebar, tag_type_id, head,
            )
            if tag is not None:
                _aplicar_estilo_tag_confinamiento_viga(tag, head)
                return tag, None
            if err:
                return None, err
        except Exception as ex:
            try:
                return None, unicode(ex)
            except Exception:
                return None, str(ex)
    refs = _referencias_tag_rebar_viga(document, rebar, view, tags_mod)
    if not refs:
        return None, u"sin referencia API"
    last_ex = None
    for ref in refs:
        try:
            tag = IndependentTag.Create(
                document,
                tag_type_id,
                view.Id,
                ref,
                False,
                TagOrientation.Horizontal,
                head,
            )
            if tag is not None:
                _aplicar_estilo_tag_confinamiento_viga(tag, head)
                return tag, None
        except Exception as ex:
            last_ex = ex
    for ref in refs:
        try:
            tag = IndependentTag.Create(
                document,
                view.Id,
                ref,
                False,
                TagMode.TM_ADDBY_CATEGORY,
                TagOrientation.Horizontal,
                head,
            )
            if tag is not None:
                try:
                    tag.SetTypeId(tag_type_id)
                except Exception:
                    pass
                _aplicar_estilo_tag_confinamiento_viga(tag, head)
                return tag, None
        except Exception as ex:
            last_ex = ex
    if last_ex is not None:
        try:
            return None, unicode(last_ex)
        except Exception:
            return None, str(last_ex)
    return None, u"no se pudo crear IndependentTag"


def _etiquetar_rebar_confinamiento_viga(
    document, view, rebar, tag_map, tags_mod=None, host=None, head_override=None,
):
    if rebar is None or document is None or view is None:
        return False, u"Rebar/vista inválidos"
    sym, shape_lbl = _resolve_tag_symbol_viga(document, rebar, tag_map, tags_mod)
    if sym is None:
        return False, u"sin tipo para RebarShape «{0}»".format(shape_lbl or u"?")
    head = head_override
    if head is None:
        head = _centro_etiqueta_confinamiento_viga(document, view, rebar, host=host)
    if head is None:
        return False, u"sin centro de array para etiqueta"
    tag, err = _crear_tag_confinamiento_sin_leader(
        document, view, rebar, sym.Id, head, tags_mod,
    )
    if tag is not None:
        return True, None
    return False, err or u"etiqueta no creada"


def _etiquetar_multihost_confinamiento_viga(
    document, view, rebar_ids, tag_map, tags_mod, host=None, head_override=None,
):
    from Autodesk.Revit.DB import Reference

    ids = list(rebar_ids or [])
    if not ids:
        return False, u"grupo multihost vacío"
    primary = document.GetElement(ids[-1])
    if primary is None:
        return False, u"barra principal no encontrada"
    sym, shape_lbl = _resolve_tag_symbol_viga(document, primary, tag_map, tags_mod)
    if sym is None:
        return False, u"sin tipo multihost para shape «{0}»".format(shape_lbl or u"?")
    head = head_override
    if head is None:
        head = _centro_multihost_rebars_en_vista(document, view, ids, host=host)
        if head is None:
            head = _centro_etiqueta_confinamiento_viga(document, view, primary, host=host)
    if head is None:
        return False, u"sin centro multihost para etiqueta"
    tag, err = _crear_tag_confinamiento_sin_leader(
        document, view, primary, sym.Id, head, tags_mod,
    )
    if tag is None:
        return False, err or u"multihost no creado"
    if len(ids) < 2:
        return True, None
    add_fn = getattr(tag, u"AddReferences", None)
    if add_fn is None:
        return False, u"Revit sin AddReferences para multihost"
    extra_refs = []
    if tags_mod is not None and hasattr(tags_mod, u"_refs_multihost_para_rebars"):
        try:
            extra_refs = tags_mod._refs_multihost_para_rebars(document, view, ids[:-1])
        except Exception:
            extra_refs = []
    if not extra_refs and _rebar_reference_candidates_for_tag is not None:
        seen = set()
        for rid in ids[:-1]:
            rb = document.GetElement(rid)
            for ref in _rebar_reference_candidates_for_tag(document, view, rb) or []:
                try:
                    key = ref.ConvertToStableRepresentation(document)
                except Exception:
                    key = id(ref)
                if key in seen:
                    continue
                seen.add(key)
                extra_refs.append(ref)
    if not extra_refs:
        return False, u"sin referencias adicionales multihost"
    refs_add = List[Reference]()
    for ref in extra_refs:
        refs_add.Add(ref)
    try:
        add_fn(refs_add)
    except Exception as ex:
        try:
            document.Delete(tag.Id)
        except Exception:
            pass
        try:
            return False, unicode(ex)
        except Exception:
            return False, str(ex)
    _aplicar_estilo_tag_confinamiento_viga(tag, head)
    return True, None


def viga_conf_type(conf):
    """Equivalente muros Tipo 1 / Tipo 2 para offsets y multihost."""
    c = _norm_conf(conf)
    if c.get("perimetral"):
        return _CONF_TYPE_PERIMETER
    ties = c.get("ties") or []
    pairs = c.get("pairs") or []
    if len(ties) >= 2 and not pairs:
        return _CONF_TYPE_PERIMETER
    if len(ties) >= 2:
        return _CONF_TYPE_PERIMETER
    return _CONF_TYPE_TIE_LAYER_1


def viga_confinement_tag_mode(conf, job_kind, stirrup_group_count=1):
    """
    Modo de etiqueta por barra de confinamiento en vigas.

    Returns:
        ``conf_tag`` | ``conf_tag_multihost`` | ``None``
    """
    c = _norm_conf(conf)
    ties = c.get("ties") or []
    if job_kind == u"stirrup":
        try:
            n_grp = int(stirrup_group_count or 1)
        except Exception:
            n_grp = 1
        if n_grp >= 2:
            return u"conf_tag_multihost"
        return u"conf_tag"
    if job_kind == u"tie":
        if len(ties) >= 2:
            return u"conf_tag_multihost"
        return u"conf_tag"
    return None


def _job_sort_key(job):
    kind = job.get(u"job_kind") or u""
    sk = job.get(u"stirrup_kind") or u""
    if kind == u"stirrup":
        if sk in (u"zona_ext_cent", u"perimetral"):
            return (0, 0, 0)
        if sk == u"pair":
            return (0, 1, int((job.get(u"pair") or [0, 0])[0]))
        return (0, 2, 0)
    try:
        ti = int(job.get(u"tie_index") or 0)
    except Exception:
        ti = 0
    return (1, ti, 0)


def _job_to_cj(job):
    conf = _norm_conf(job.get(u"conf"))
    sk = job.get(u"stirrup_kind") or u""
    pair = job.get(u"pair") or []
    if sk == u"pair" and pair:
        try:
            stirrup_sort = (0, 1, int(pair[0]))
        except Exception:
            stirrup_sort = (0, 1, 0)
    elif sk in (u"zona_ext_cent", u"perimetral"):
        stirrup_sort = (0, 0, 0)
    else:
        stirrup_sort = (0, 2, 0)
    try:
        stirrup_zone_index = int(job.get(u"stirrup_zone_index", 0) or 0)
    except Exception:
        stirrup_zone_index = 0
    return {
        u"wid": job.get(u"host_id"),
        u"extremo": _VIGA_EXTREMO_GROUP,
        u"wall": None,
        u"conf_type": viga_conf_type(conf),
        u"job_kind": job.get(u"job_kind"),
        u"tie_layer_index": job.get(u"tie_index"),
        u"stirrup_zone_index": stirrup_zone_index,
        u"stirrup_zone_kind": job.get(u"stirrup_zone_kind"),
        u"stirrup_sort_key": stirrup_sort,
        u"ex_cfg": {u"n_capas": int(job.get(u"n_capas") or 0)},
    }


def _append_conf_tag_message(avisos, job, err):
    if not err or avisos is None:
        return
    if len(avisos) >= 12:
        return
    kind = u"traba" if job.get(u"job_kind") == u"tie" else u"estribo"
    try:
        hid = int(job.get(u"host_id") or 0)
    except Exception:
        hid = job.get(u"host_id")
    avisos.append(u"Viga {0} etiqueta {1}: {2}".format(hid, kind, err))


def _tag_single_confinement(document, view, res, job, tags_mod, tag_map, head_override=None):
    """Etiqueta estribo/traba: sin leader, debajo de la viga si hay layout de lote."""
    rb = document.GetElement(job.get(u"rebar_id"))
    if rb is None:
        return False, u"Rebar no encontrado"
    return _etiquetar_rebar_confinamiento_viga(
        document,
        view,
        rb,
        tag_map,
        tags_mod=tags_mod,
        host=job.get(u"host"),
        head_override=head_override,
    )


def _ordered_rebar_ids_from_multihost_grp(grp):
    items = list((grp or {}).get(u"items") or [])
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
            key = rid.IntegerValue
        except Exception:
            key = rid
        if key in seen:
            continue
        seen.add(key)
        out.append(rid)
    return out


def _stirrup_group_key(job):
    """Agrupa estribos del mismo tramo longitudinal (extremo/central) en la misma viga."""
    try:
        hid = int(job.get(u"host_id") or 0)
    except Exception:
        hid = 0
    try:
        zi = int(job.get(u"stirrup_zone_index", 0) or 0)
    except Exception:
        zi = 0
    return (hid, zi)


def _stirrup_multihost_group_key(cj):
    try:
        wid = int(cj.get(u"wid", 0) or 0)
    except Exception:
        wid = 0
    try:
        zi = int(cj.get(u"stirrup_zone_index", 0) or 0)
    except Exception:
        zi = 0
    return (wid, zi)


def _count_stirrup_groups(jobs):
    counts = {}
    for job in jobs or []:
        if job.get(u"job_kind") != u"stirrup":
            continue
        key = _stirrup_group_key(job)
        counts[key] = counts.get(key, 0) + 1
    return counts


def _ordered_rebar_ids_from_stirrup_grp(grp):
    items = list((grp or {}).get(u"items") or [])
    if not items:
        return []
    try:
        items.sort(key=lambda x: tuple(x.get(u"stirrup_sort_key") or (0, 0, 0)))
    except Exception:
        pass
    out = []
    seen = set()
    for it in items:
        rid = it.get(u"rebar_id")
        if rid is None:
            continue
        try:
            key = rid.IntegerValue
        except Exception:
            key = rid
        if key in seen:
            continue
        seen.add(key)
        out.append(rid)
    return out


def _multihost_group_key(cj):
    try:
        wid = int(cj.get(u"wid", 0) or 0)
    except Exception:
        wid = 0
    return (wid, cj.get(u"extremo") or _VIGA_EXTREMO_GROUP)


def _register_multihost_stirrup_pending(res, cj, rebar):
    if res is None or cj is None or rebar is None:
        return
    key = _stirrup_multihost_group_key(cj)
    pending = res.setdefault(u"_conf_multihost_stirrup_pending", {})
    grp = pending.get(key)
    if grp is None:
        grp = {
            u"items": [],
            u"wall": cj.get(u"wall"),
            u"stirrup_zone_index": cj.get(u"stirrup_zone_index"),
            u"stirrup_zone_kind": cj.get(u"stirrup_zone_kind"),
            u"conf_type": cj.get(u"conf_type"),
            u"n_capas": cj.get(u"ex_cfg", {}).get(u"n_capas"),
        }
        pending[key] = grp
    grp.setdefault(u"items", []).append({
        u"rebar_id": rebar.Id,
        u"stirrup_sort_key": cj.get(u"stirrup_sort_key"),
    })


def _register_multihost_traba_pending(res, cj, rebar, tags_mod):
    if tags_mod is not None:
        tags_mod.register_confinement_multihost_traba_pending(res, cj, rebar)
        return
    if res is None or cj is None or rebar is None:
        return
    key = _multihost_group_key(cj)
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


def _flush_multihost_stirrup_viga(document, view, res, tag_map, tags_mod, lot_layouts=None):
    pending = res.pop(u"_conf_multihost_stirrup_pending", None) or {}
    n_ok = 0
    n_fail = 0
    mh_map = tag_map
    if tags_mod is not None:
        try:
            mh_map = _collect_tag_symbol_map(
                document, CONFINEMENT_MULTIHOST_TAG_FAMILY, tags_mod,
            ) or tag_map
        except Exception:
            mh_map = tag_map
    for key, grp in pending.items():
        ids = _ordered_rebar_ids_from_stirrup_grp(grp)
        host = None
        try:
            from Autodesk.Revit.DB import ElementId

            wid = int(key[0] or 0)
            if wid > 0:
                host = document.GetElement(ElementId(wid))
        except Exception:
            host = None
        lot_key = (int(key[0] or 0), u"stirrup_mh", int(key[1] or 0))
        head = _head_from_viga_lot_layout(view, (lot_layouts or {}).get(lot_key))
        ok, err = _etiquetar_multihost_confinamiento_viga(
            document, view, ids, mh_map, tags_mod, host=host, head_override=head,
        )
        if ok:
            n_ok += 1
        else:
            n_fail += 1
            if err:
                res.setdefault(u"messages", []).append(err)
    res[u"n_conf_stirrup_tags_created"] = n_ok
    res[u"n_conf_stirrup_tags_fail"] = n_fail


def _flush_multihost_confinamiento_viga(document, view, res, tag_map, tags_mod, lot_layouts=None):
    pending = res.pop(u"_conf_multihost_traba_pending", None) or {}
    res.pop(u"_conf_multihost_flush_order", None)
    n_ok = 0
    n_fail = 0
    mh_map = tag_map
    if tags_mod is not None:
        try:
            mh_map = _collect_tag_symbol_map(
                document, CONFINEMENT_MULTIHOST_TAG_FAMILY, tags_mod,
            ) or tag_map
        except Exception:
            mh_map = tag_map
    for key, grp in pending.items():
        ids = _ordered_rebar_ids_from_multihost_grp(grp)
        host = None
        try:
            from Autodesk.Revit.DB import ElementId

            wid = int(key[0] or 0)
            if wid > 0:
                host = document.GetElement(ElementId(wid))
        except Exception:
            host = None
        lot_key = (int(key[0] or 0), u"tie_mh", 0)
        head = _head_from_viga_lot_layout(view, (lot_layouts or {}).get(lot_key))
        ok, err = _etiquetar_multihost_confinamiento_viga(
            document, view, ids, mh_map, tags_mod, host=host, head_override=head,
        )
        if ok:
            n_ok += 1
        else:
            n_fail += 1
            if err:
                res.setdefault(u"messages", []).append(err)
    res[u"n_conf_tags_created"] = n_ok
    res[u"n_conf_tags_fail"] = n_fail


def etiquetar_confinamiento_en_vista(
    document,
    view,
    conf_jobs,
    use_transaction=False,
):
    """
    Etiqueta estribos/trabas de confinamiento en ``view``.

    ``use_transaction=False`` cuando ya hay ``Arainco: Armado vigas`` abierta.

    Returns:
        ``(n_etiquetas, avisos, err)``
    """
    jobs = [j for j in (conf_jobs or []) if j and j.get(u"rebar_id") is not None]
    if not jobs:
        return 0, [], None
    if document is None or view is None:
        return 0, [], u"Sin documento o vista activa para etiquetar confinamiento."

    tags_mod = _load_cabezal_tags_module()

    res = {}
    avisos = []
    n_ok = 0
    n_fail = 0

    tag_map = _collect_tag_symbol_map(document, CONFINEMENT_TAG_FAMILY, tags_mod)
    if not tag_map:
        return (
            0,
            [],
            u"Sin tipos OST_RebarTags para «{0}».".format(CONFINEMENT_TAG_FAMILY),
        )

    if tags_mod is not None:
        res[u"_conf_multihost_tag_map"] = _collect_tag_symbol_map(
            document, CONFINEMENT_MULTIHOST_TAG_FAMILY, tags_mod,
        )

    ordered = sorted(jobs, key=_job_sort_key)
    stirrup_group_counts = _count_stirrup_groups(jobs)
    lot_registry = {}
    single_tag_jobs = []

    for job in ordered:
        conf = _norm_conf(job.get(u"conf"))
        job_kind = job.get(u"job_kind")
        stirrup_grp_n = stirrup_group_counts.get(_stirrup_group_key(job), 1)
        mode = viga_confinement_tag_mode(conf, job_kind, stirrup_group_count=stirrup_grp_n)
        cj = _job_to_cj(job)
        rb = document.GetElement(job.get(u"rebar_id"))
        if rb is None:
            n_fail += 1
            _append_conf_tag_message(avisos, job, u"Rebar no encontrado")
            continue
        host = job.get(u"host")

        if mode == u"conf_tag_multihost":
            try:
                if job_kind == u"stirrup":
                    _register_multihost_stirrup_pending(res, cj, rb)
                    lot_key = (
                        int(job.get(u"host_id") or 0),
                        u"stirrup_mh",
                        int(job.get(u"stirrup_zone_index", 0) or 0),
                    )
                else:
                    _register_multihost_traba_pending(res, cj, rb, tags_mod)
                    lot_key = (int(job.get(u"host_id") or 0), u"tie_mh", 0)
                    mh_order = res.setdefault(u"_conf_multihost_flush_order", [])
                    mh_key = (
                        tags_mod.confinement_multihost_group_key(cj)
                        if tags_mod is not None
                        else _multihost_group_key(cj)
                    )
                    if mh_key not in mh_order:
                        mh_order.append(mh_key)
                _register_viga_conf_lot(lot_registry, lot_key, host, rb)
            except Exception as ex_mh:
                n_fail += 1
                try:
                    err = unicode(ex_mh)
                except Exception:
                    err = str(ex_mh)
                _append_conf_tag_message(avisos, job, err)
            continue

        if mode != u"conf_tag":
            continue

        _register_viga_conf_lot(
            lot_registry, _viga_conf_lot_key(job, multihost=False), host, rb,
        )
        single_tag_jobs.append(job)

    lot_layouts = _build_viga_conf_lot_layouts(view, lot_registry)
    n_dims = _create_viga_confinement_dimensions(document, view, lot_layouts, avisos)
    if n_dims > 0:
        avisos.append(u"Cotas confinamiento: {0}.".format(n_dims))

    for job in single_tag_jobs:
        lot_key = _viga_conf_lot_key(job, multihost=False)
        head = _head_from_viga_lot_layout(view, lot_layouts.get(lot_key))
        ok_tag, err_tag = _tag_single_confinement(
            document, view, res, job, tags_mod, tag_map, head_override=head,
        )
        if ok_tag:
            n_ok += 1
        else:
            n_fail += 1
            _append_conf_tag_message(avisos, job, err_tag)

    pending_traba = res.get(u"_conf_multihost_traba_pending") or {}
    if pending_traba:
        try:
            _flush_multihost_confinamiento_viga(
                document, view, res, tag_map, tags_mod, lot_layouts=lot_layouts,
            )
            n_ok += int(res.get(u"n_conf_tags_created", 0))
            n_fail += int(res.get(u"n_conf_tags_fail", 0))
            for msg in res.get(u"messages") or []:
                if len(avisos) < 12:
                    avisos.append(msg)
        except Exception as ex_flush:
            avisos.append(u"Multihost confinamiento vigas: {0}".format(
                unicode(ex_flush) if ex_flush else u"?"
            ))

    pending_stirrup = res.get(u"_conf_multihost_stirrup_pending") or {}
    if pending_stirrup:
        try:
            res[u"messages"] = []
            _flush_multihost_stirrup_viga(
                document, view, res, tag_map, tags_mod, lot_layouts=lot_layouts,
            )
            n_ok += int(res.get(u"n_conf_stirrup_tags_created", 0))
            n_fail += int(res.get(u"n_conf_stirrup_tags_fail", 0))
            for msg in res.get(u"messages") or []:
                if len(avisos) < 12:
                    avisos.append(msg)
        except Exception as ex_flush:
            avisos.append(u"Multihost estribos vigas: {0}".format(
                unicode(ex_flush) if ex_flush else u"?"
            ))

    if n_ok <= 0 and n_fail <= 0:
        return 0, avisos, None
    return n_ok, avisos, None
