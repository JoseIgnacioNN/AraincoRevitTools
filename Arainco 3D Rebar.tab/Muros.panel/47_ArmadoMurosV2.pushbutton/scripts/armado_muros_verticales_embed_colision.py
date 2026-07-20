# -*- coding: utf-8 -*-
"""
Empotramiento en cabeza de barras verticales (cara exterior e interior).

Criterio cabeza (V2 — apilamiento en Z, sin sonda de sólidos tras estirar):
1. Muros de la selección ordenados / evaluados por contacto en Z + solape en planta.
2. Si el host tiene **muro apilado encima** en la selección → estirar L(Ø) de tabla
   (empotramiento).
3. Si **no** hay muro apilado encima → retraída ``25 mm + Ø/2`` + pata L.
4. Largo pata L = espesor muro − 50 mm − Ø horiz. ext. − Ø horiz. int.
   Orden de tramos: **segmento 0** = eje vertical (pie→cabeza); **segmento 1** = pata L.

Pie:
5. Si el muro host tiene **fundación estructural** unida → pie vs fundación
   (prisma 100 mm; colisión → estira + pata L; sin colisión → ``25+Ø/2`` + pata L int.).
6. Si **no** hay fundación → si hay **muro apilado debajo** no mutar; si no,
   retraer ``25+Ø/2`` + pata L en pie.

Orden por barra: **cabeza** → **pie**.
"""

from __future__ import print_function

import os
import sys

import clr

clr.AddReference("RevitAPI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BooleanOperationsType,
    BooleanOperationsUtils,
    BuiltInCategory,
    Curve,
    CurveLoop,
    ElementId,
    GeometryCreationUtilities,
    JoinGeometryUtils,
    Line,
    Options,
    Transaction,
    UnitUtils,
    UnitTypeId,
    ViewDetailLevel,
    Wall,
    XYZ,
    GeometryInstance,
    Solid,
)
from Autodesk.Revit.DB.Structure import (
    MultiplanarOption,
    Rebar,
    RebarBarType,
    RebarStyle,
)

# ── Rutas locales del pushbutton ────────────────────────────────────────────
_EMBED_PROBE_XY_MARGIN_MM = 1.0
_EMBED_PROBE_MIN_HALF_SIDE_MM = 2.0
_TOL_VOL_INTERSECCION_FT3 = 1.0e-9
# None → tabla base BIMTools (860 mm @ ø12). G25/G35/G45 usan tablas de proyecto.
CONCRETE_GRADE = None
NO_COLLISION_RETRACT_BASE_MM = 25.0
# Recubrimiento 25 mm cara ext. + 25 mm cara int. (descontado del espesor en pata L cabeza).
PATA_L_RECUBRIMIENTO_CARAS_MM = 50.0
FOUNDATION_PROBE_BASE_MM = 100.0
PIE_MURO_SIN_FUND_EVAL_STRETCH_MM = 100.0
FOUNDATION_STRETCH_RESTA_MM = 50.0
MURO_VERT_MIN_ABS_TZ = 0.45


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
    from armado_muros_nodo_shared import (
        ajustar_inclusion_extremos_rebar_set_con_fallback,
        desactivar_extremos_rebar_set,
    )
except Exception:
    ajustar_inclusion_extremos_rebar_set_con_fallback = None
    desactivar_extremos_rebar_set = None

try:
    from arearein_verticales_empotramiento_rps import (
        _copy_layout_rebar_shape_driven,
        _create_from_curves_no_hooks,
        _hook_orient_for_create,
        _nominal_diameter_mm_from_rebar,
        _rebar_eje_p_start_p_end,
        _rebar_es_vertical_por_criterio,
        _rebar_normal,
        _tangent_at_end_of_curve,
        _tangent_start_curve,
        _extender_rebar_por_eje_mm,
    )
except Exception:
    _copy_layout_rebar_shape_driven = None
    _create_from_curves_no_hooks = None
    _hook_orient_for_create = None
    _nominal_diameter_mm_from_rebar = None
    _rebar_eje_p_start_p_end = None
    _rebar_es_vertical_por_criterio = None
    _rebar_normal = None
    _tangent_at_end_of_curve = None
    _tangent_start_curve = None
    _extender_rebar_por_eje_mm = None

try:
    import rebar_extender_l_ganchos_135_rps as l135
except Exception:
    l135 = None

try:
    from arearein_exterior_h_l135_rps import _rebar_solo_cara_exterior
except Exception:
    _rebar_solo_cara_exterior = None

try:
    from arearein_interior_h_l135_rps import _rebar_solo_cara_interior
except Exception:
    _rebar_solo_cara_interior = None


def _reload_embed_vendor_modules():
    """Re-pinea ``scripts/`` local y recarga vendors si el import inicial falló."""
    global _copy_layout_rebar_shape_driven
    global _create_from_curves_no_hooks
    global _hook_orient_for_create
    global _nominal_diameter_mm_from_rebar
    global _rebar_eje_p_start_p_end
    global _rebar_es_vertical_por_criterio
    global _rebar_normal
    global _tangent_at_end_of_curve
    global _tangent_start_curve
    global _extender_rebar_por_eje_mm
    global traslape_mm_from_nominal_diameter_mm
    global hook_length_mm_from_nominal_diameter_mm
    global pata_eje_curve_loop_mm_desde_tabla_mm
    global l135
    global _rebar_solo_cara_exterior
    global _rebar_solo_cara_interior

    if (
        _extender_rebar_por_eje_mm is not None
        and traslape_mm_from_nominal_diameter_mm is not None
        and (
            _rebar_solo_cara_exterior is not None
            or _rebar_solo_cara_interior is not None
        )
    ):
        return

    try:
        import bootstrap_paths

        bootstrap_paths.pin_local_scripts_first()
    except Exception:
        _ensure_pushbutton_path()

    for _mod_name in (
        u"arearein_verticales_empotramiento_rps",
        u"arearein_exterior_h_l135_rps",
        u"arearein_interior_h_l135_rps",
        u"rebar_extender_l_ganchos_135_rps",
        u"bimtools_rebar_hook_lengths",
    ):
        try:
            sys.modules.pop(_mod_name, None)
        except Exception:
            pass

    try:
        from bimtools_rebar_hook_lengths import (
            hook_length_mm_from_nominal_diameter_mm,
            pata_eje_curve_loop_mm_desde_tabla_mm,
            traslape_mm_from_nominal_diameter_mm,
        )
    except Exception:
        pass

    try:
        from arearein_verticales_empotramiento_rps import (
            _copy_layout_rebar_shape_driven,
            _create_from_curves_no_hooks,
            _hook_orient_for_create,
            _nominal_diameter_mm_from_rebar,
            _rebar_eje_p_start_p_end,
            _rebar_es_vertical_por_criterio,
            _rebar_normal,
            _tangent_at_end_of_curve,
            _tangent_start_curve,
            _extender_rebar_por_eje_mm,
        )
    except Exception:
        pass

    try:
        import rebar_extender_l_ganchos_135_rps as _l135

        l135 = _l135
    except Exception:
        pass

    try:
        from arearein_exterior_h_l135_rps import _rebar_solo_cara_exterior
    except Exception:
        pass

    try:
        from arearein_interior_h_l135_rps import _rebar_solo_cara_interior
    except Exception:
        pass


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _element_id_int(eid):
    if eid is None:
        return None
    try:
        v = getattr(eid, "Value", None)
        if v is not None:
            return int(v)
    except Exception:
        pass
    try:
        return int(eid.IntegerValue)
    except Exception:
        pass
    try:
        return int(eid)
    except Exception:
        return None


def _geometry_options():
    opts = Options()
    try:
        opts.ComputeReferences = False
        opts.IncludeNonVisibleObjects = False
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    return opts


def _solids_list_element(elem, opts):
    """Lista de sólidos de un elemento (para caché de colisión)."""
    out = []
    for sd in _iter_solids_element(elem, opts):
        out.append(sd)
    return out


def _build_solids_cache(elements, geom_opts):
    """``{ element_id_int: [Solid, ...] }`` — una sola lectura de geometría por elemento."""
    cache = {}
    for el in elements or []:
        if el is None:
            continue
        eid = _element_id_int(getattr(el, "Id", None))
        if eid is None:
            continue
        try:
            cache[int(eid)] = _solids_list_element(el, geom_opts)
        except Exception:
            cache[int(eid)] = []
    return cache


def _bbox_z_range(elem):
    """``(z_min, z_max)`` del bounding box o ``None``."""
    if elem is None:
        return None
    try:
        bb = elem.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None:
        return None
    try:
        return float(bb.Min.Z), float(bb.Max.Z)
    except Exception:
        return None


def _build_bbox_z_cache(elements):
    cache = {}
    for el in elements or []:
        if el is None:
            continue
        eid = _element_id_int(getattr(el, "Id", None))
        if eid is None:
            continue
        zr = _bbox_z_range(el)
        if zr is not None:
            cache[int(eid)] = zr
    return cache


def _bbox_xy_overlap(bb_a, bb_b, tol_xy):
    if bb_a is None or bb_b is None:
        return False
    try:
        return not (
            float(bb_a.Max.X) < float(bb_b.Min.X) - tol_xy
            or float(bb_b.Max.X) < float(bb_a.Min.X) - tol_xy
            or float(bb_a.Max.Y) < float(bb_b.Min.Y) - tol_xy
            or float(bb_b.Max.Y) < float(bb_a.Min.Y) - tol_xy
        )
    except Exception:
        return False


def _stack_z_tolerance_ft():
    try:
        return float(_mm_to_internal(40.0))
    except Exception:
        return 0.12


def _muro_apilado_sobre_host(host, other, tol_z=None):
    """
    True si ``other`` apoya sobre la cara superior de ``host``
    (contacto Z + solape en planta). Criterio alineado a wall_node / coronamiento.
    """
    if host is None or other is None:
        return False
    if not isinstance(host, Wall) or not isinstance(other, Wall):
        return False
    try:
        if other.Id == host.Id:
            return False
    except Exception:
        pass
    try:
        hbb = host.get_BoundingBox(None)
        obb = other.get_BoundingBox(None)
    except Exception:
        return False
    if hbb is None or obb is None:
        return False
    if tol_z is None:
        tol_z = _stack_z_tolerance_ft()
    try:
        d_face = float(_mm_to_internal(4.0))
        band = max(float(tol_z), 1e-4)
        tol_xy = max(float(tol_z), float(_mm_to_internal(50.0)))
    except Exception:
        d_face = 0.01
        band = 0.12
        tol_xy = 0.15
    if not _bbox_xy_overlap(hbb, obb, tol_xy):
        return False
    try:
        if abs(float(obb.Min.Z) - float(hbb.Max.Z)) > d_face + band:
            return False
        if float(obb.Max.Z) <= float(hbb.Max.Z) + 1e-5:
            return False
    except Exception:
        return False
    return True


def _muro_apilado_bajo_host(host, other, tol_z=None):
    """True si ``other`` contacta la cara inferior de ``host`` (apilado debajo)."""
    if host is None or other is None:
        return False
    if not isinstance(host, Wall) or not isinstance(other, Wall):
        return False
    try:
        if other.Id == host.Id:
            return False
    except Exception:
        pass
    try:
        hbb = host.get_BoundingBox(None)
        obb = other.get_BoundingBox(None)
    except Exception:
        return False
    if hbb is None or obb is None:
        return False
    if tol_z is None:
        tol_z = _stack_z_tolerance_ft()
    try:
        d_face = float(_mm_to_internal(4.0))
        band = max(float(tol_z), 1e-4)
        tol_xy = max(float(tol_z), float(_mm_to_internal(50.0)))
    except Exception:
        d_face = 0.01
        band = 0.12
        tol_xy = 0.15
    if not _bbox_xy_overlap(hbb, obb, tol_xy):
        return False
    try:
        if abs(float(obb.Max.Z) - float(hbb.Min.Z)) > d_face + band:
            return False
        if float(obb.Min.Z) >= float(hbb.Min.Z) - 1e-5:
            return False
    except Exception:
        return False
    return True


def _build_apilamiento_maps(walls):
    """
    Precalcula por host_id si hay muro apilado encima / debajo en la selección.

    :returns: ``(ids_con_sobre, ids_con_bajo)`` — sets de ``int`` wall id.
    """
    sobre = set()
    bajo = set()
    walls_list = [w for w in (walls or []) if w is not None and isinstance(w, Wall)]
    tol = _stack_z_tolerance_ft()
    n = len(walls_list)
    for i in range(n):
        host = walls_list[i]
        hid = _element_id_int(getattr(host, "Id", None))
        if hid is None:
            continue
        for j in range(n):
            if i == j:
                continue
            other = walls_list[j]
            if hid not in sobre and _muro_apilado_sobre_host(host, other, tol):
                sobre.add(int(hid))
            if hid not in bajo and _muro_apilado_bajo_host(host, other, tol):
                bajo.add(int(hid))
            if hid in sobre and hid in bajo:
                break
    return sobre, bajo


def _host_tiene_apilado_sobre(host, ids_con_sobre):
    hid = _element_id_int(getattr(host, "Id", None))
    if hid is None:
        return False
    try:
        return int(hid) in (ids_con_sobre or set())
    except Exception:
        return False


def _host_tiene_apilado_bajo(host, ids_con_bajo):
    hid = _element_id_int(getattr(host, "Id", None))
    if hid is None:
        return False
    try:
        return int(hid) in (ids_con_bajo or set())
    except Exception:
        return False


def _iter_obstacle_solids(
    wall_obstacles,
    host_wall_id,
    geom_opts,
    solids_cache=None,
    bbox_z_cache=None,
    z_lo=None,
    z_hi=None,
):
    """
    Itera sólidos de obstáculos omitiendo el host.
    Con caché no llama ``get_Geometry``; con ``bbox_z_cache`` descarta muros
    cuyo bbox no solapa ``[z_lo, z_hi]``.
    """
    host_id = _element_id_int(host_wall_id)
    for wall in wall_obstacles or []:
        if wall is None:
            continue
        wid = _element_id_int(getattr(wall, "Id", None))
        if host_id is not None and wid is not None and wid == host_id:
            continue
        if (
            bbox_z_cache is not None
            and wid is not None
            and z_lo is not None
            and z_hi is not None
        ):
            zr = bbox_z_cache.get(int(wid))
            if zr is not None:
                zmin, zmax = zr
                if zmax < float(z_lo) - 1e-9:
                    continue
                if zmin > float(z_hi) + 1e-9:
                    continue
        solids = None
        if solids_cache is not None and wid is not None:
            solids = solids_cache.get(int(wid))
        if solids is None:
            solids = _solids_list_element(wall, geom_opts)
        for sd in solids or []:
            yield sd


def _iter_solids_element(elem, opts):
    if elem is None:
        return
    try:
        ge = elem.get_Geometry(opts)
    except Exception:
        return
    if ge is None:
        return
    try:
        from bimtools_clr_collections import iterate_net_collection, safe_solid_volume
    except Exception:
        iterate_net_collection = None
        safe_solid_volume = None

    if iterate_net_collection is not None:
        geom_items = iterate_net_collection(ge)
    else:
        geom_items = []
        try:
            for obj in ge:
                geom_items.append(obj)
        except Exception:
            try:
                n = int(ge.Count)
            except Exception:
                n = 0
            for i in range(n):
                try:
                    geom_items.append(ge[i])
                except Exception:
                    try:
                        geom_items.append(ge.get_Item(i))
                    except Exception:
                        pass

    def _vol_ok(solid):
        if solid is None:
            return False
        if safe_solid_volume is not None:
            v = safe_solid_volume(solid)
            return v is not None and abs(v) > 1e-12
        try:
            return float(solid.Volume) > 1e-12
        except Exception:
            return False

    def _yield_from_geom_collection(collection):
        if collection is None:
            return
        if iterate_net_collection is not None:
            sub_items = iterate_net_collection(collection)
        else:
            sub_items = []
            try:
                for sg in collection:
                    sub_items.append(sg)
            except Exception:
                pass
        for sg in sub_items:
            if isinstance(sg, Solid) and _vol_ok(sg):
                yield sg

    for obj in geom_items:
        if obj is None:
            continue
        if isinstance(obj, Solid):
            if _vol_ok(obj):
                yield obj
        elif isinstance(obj, GeometryInstance):
            try:
                inst_geom = obj.GetInstanceGeometry()
                for sg in _yield_from_geom_collection(inst_geom):
                    yield sg
            except Exception:
                pass
            try:
                xf = obj.Transform
            except Exception:
                xf = None
            if xf is not None:
                try:
                    sym_geom = obj.GetSymbolGeometry()
                    for sg in _yield_from_geom_collection(sym_geom):
                        try:
                            sg_t = sg.CreateTransformed(xf)
                        except Exception:
                            sg_t = sg
                        if isinstance(sg_t, Solid) and _vol_ok(sg_t):
                            yield sg_t
                except Exception:
                    pass


def _solidos_intersectan_volumen(solid_a, solid_b, tol_volumen=_TOL_VOL_INTERSECCION_FT3):
    if solid_a is None or solid_b is None:
        return False
    try:
        from bimtools_clr_collections import safe_solid_volume
    except Exception:
        safe_solid_volume = None

    def _vol(sol):
        if sol is None:
            return None
        if safe_solid_volume is not None:
            return safe_solid_volume(sol)
        try:
            return float(sol.Volume)
        except Exception:
            return None

    va = _vol(solid_a)
    vb = _vol(solid_b)
    if va is None or vb is None or va <= 1e-12 or vb <= 1e-12:
        return False
    try:
        inter = BooleanOperationsUtils.ExecuteBooleanOperation(
            solid_a, solid_b, BooleanOperationsType.Intersect,
        )
    except Exception:
        return False
    if inter is None:
        return False
    vi = _vol(inter)
    if vi is None:
        return False
    return vi > float(tol_volumen)


def _build_vertical_square_prism_solid_ip_fallback(
    px, py, z_start_ft, half_side_ft, height_ft,
):
    """Respaldo IList / listas Python si el helper CLR no devuelve sólido."""
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
        lines = [
            Line.CreateBound(p1, p2),
            Line.CreateBound(p2, p3),
            Line.CreateBound(p3, p4),
            Line.CreateBound(p4, p1),
        ]
    except Exception:
        return None
    sol = None
    try:
        cl = List[Curve]()
        for ln in lines:
            cl.Add(ln)
        loop = CurveLoop.Create(cl)
        if loop is not None:
            loops = List[CurveLoop]()
            loops.Add(loop)
            sol = GeometryCreationUtilities.CreateExtrusionGeometry(
                loops, XYZ.BasisZ, hgt,
            )
    except Exception:
        sol = None
    if sol is None:
        try:
            loop = CurveLoop.Create(lines)
            if loop is not None:
                sol = GeometryCreationUtilities.CreateExtrusionGeometry(
                    [loop], XYZ.BasisZ, hgt,
                )
        except Exception:
            sol = None
    if sol is None:
        return None
    try:
        if float(sol.Volume) < 1e-15:
            return None
    except Exception:
        pass
    return sol


def _build_vertical_square_prism_solid(px, py, z_start_ft, half_side_ft, height_ft):
    probe = None
    try:
        from bimtools_clr_collections import create_vertical_square_prism_solid

        probe = create_vertical_square_prism_solid(
            px, py, z_start_ft, half_side_ft, height_ft,
        )
    except Exception:
        probe = None
    if probe is not None:
        return probe
    return _build_vertical_square_prism_solid_ip_fallback(
        px, py, z_start_ft, half_side_ft, height_ft,
    )


def _punto_cabeza_vertical(p0, p1):
    if p0 is None or p1 is None:
        return None
    try:
        eps = _mm_to_internal(0.5)
    except Exception:
        eps = 1.0e-4
    if float(p1.Z) >= float(p0.Z) - float(eps):
        return p1
    return p0


def _punto_pie_vertical(p0, p1):
    """Extremo de menor Z del eje (pie en barras verticales)."""
    if p0 is None or p1 is None:
        return None
    try:
        eps = _mm_to_internal(0.5)
    except Exception:
        eps = 1.0e-4
    if float(p1.Z) >= float(p0.Z) - float(eps):
        return p0
    return p1


def _empotramiento_tabla_mm(d_mm, concrete_grade=None):
    """Valor de tabla traslape/empotramiento (mm) según Ø nominal y grado."""
    if traslape_mm_from_nominal_diameter_mm is None:
        return None
    try:
        d = float(d_mm)
    except Exception:
        return None
    if d <= 0.0 or d != d:
        return None
    g = CONCRETE_GRADE if concrete_grade is None else concrete_grade
    try:
        L = traslape_mm_from_nominal_diameter_mm(d, g)
    except Exception:
        return None
    if L is None or L != L or float(L) < 0.0:
        return None
    return max(0.0, float(L))


def _empotramiento_cabeza_mm_desde_diametro(d_mm, concrete_grade=None):
    """
    Largo de estiramiento en cabeza (mm): valor de tabla traslape/empotramiento por Ø
    (sin suma adicional; p. ej. ø12 → 860 mm con tabla base BIMTools).
    """
    return _empotramiento_tabla_mm(d_mm, concrete_grade)


def _empotramiento_cabeza_mm_desde_rebar(rebar, doc, concrete_grade=None):
    if _nominal_diameter_mm_from_rebar is None:
        return None
    d_mm = _nominal_diameter_mm_from_rebar(rebar, doc)
    if d_mm is None or d_mm <= 0.0:
        return None
    return _empotramiento_cabeza_mm_desde_diametro(d_mm, concrete_grade)


def _extender_vertical_cabeza_tabla_empotramiento(doc, rebar, pos_idx=0, concrete_grade=None):
    """Estira la cabeza (mayor Z) con L = tabla empotramientos por Ø del RebarBarType."""
    if _extender_rebar_por_eje_mm is None or _rebar_eje_p_start_p_end is None:
        return False, u"Extensión por tabla no disponible.", None
    L = _empotramiento_cabeza_mm_desde_rebar(rebar, doc, concrete_grade)
    if L is None:
        return False, u"Ø nominal o tabla no resuelta.", None
    if L < 0.1:
        return True, u"L=0, sin extender.", rebar
    p0, p1 = _rebar_eje_p_start_p_end(rebar, int(pos_idx))
    if p0 is None or p1 is None:
        return False, u"Sin curvas de boceto.", None
    try:
        eps = _mm_to_internal(0.5)
    except Exception:
        eps = 1.0e-4
    if float(p1.Z) >= float(p0.Z) - float(eps):
        return _extender_rebar_por_eje_mm(doc, rebar, 0.0, L, int(pos_idx))
    return _extender_rebar_por_eje_mm(doc, rebar, L, 0.0, int(pos_idx))


def _extender_vertical_pie_mm(doc, rebar, mm_pie, pos_idx=0):
    """Estira el pie (menor Z) del trazado vertical ``mm_pie`` mm."""
    if _extender_rebar_por_eje_mm is None or _rebar_eje_p_start_p_end is None:
        return False, u"Extensión en pie no disponible.", None
    m = max(0.0, float(mm_pie))
    if m < 0.1:
        return True, u"mm pie = 0, sin extender.", rebar
    p0, p1 = _rebar_eje_p_start_p_end(rebar, int(pos_idx))
    if p0 is None or p1 is None:
        return False, u"Sin curvas de boceto.", None
    try:
        eps = _mm_to_internal(0.5)
    except Exception:
        eps = 1.0e-4
    if float(p1.Z) >= float(p0.Z) - float(eps):
        return _extender_rebar_por_eje_mm(doc, rebar, m, 0.0, int(pos_idx))
    return _extender_rebar_por_eje_mm(doc, rebar, 0.0, m, int(pos_idx))


def _es_fundacion_estructural(element):
    if element is None:
        return False
    try:
        cat = element.Category
        if cat is None:
            return False
        cid = _element_id_int(cat.Id)
        if cid is None:
            return False
        return int(cid) == int(BuiltInCategory.OST_StructuralFoundation)
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
    for getter in (
        lambda: JoinGeometryUtils.GetJoinedElements(doc, element),
        lambda: JoinGeometryUtils.GetJoinedElements(doc, element.Id),
    ):
        try:
            raw = getter()
        except Exception:
            continue
        if raw is None:
            continue
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
                try:
                    eid = raw.get_Item(i)
                    if eid is not None and eid != ElementId.InvalidElementId:
                        out.append(eid)
                except Exception:
                    pass
        if out:
            return out
    return out


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


def _build_vertical_prism_downward(px, py, z_top_ft, half_side_ft, height_ft):
    """Prisma cuadrado extruido hacia abajo (−Z) desde ``z_top_ft``."""
    hgt = abs(float(height_ft))
    if hgt <= 1e-12:
        return None
    z_base = float(z_top_ft) - hgt
    return _build_vertical_square_prism_solid(
        float(px), float(py), z_base, abs(float(half_side_ft)), hgt,
    )


def _fundaciones_estructurales_unidas_muro(doc, wall):
    """Lista de elementos ``OST_StructuralFoundation`` unidos al muro."""
    if doc is None or wall is None:
        return []
    out = []
    for eid in _ids_elementos_unidos(doc, wall):
        el = doc.GetElement(eid)
        if _es_fundacion_estructural(el):
            out.append(el)
    return out


def _probe_colision_fundacion_desde_punto(
    xyz_ref,
    probe_mm,
    bar_nominal_mm,
    foundations,
    geom_opts,
    solids_cache=None,
):
    """
    Prisma de ensayo hacia abajo desde ``xyz_ref``.

    :returns: ``(colisiona, altura_fund_mm_máx)`` entre fundaciones intersectadas.
    """
    if xyz_ref is None or not foundations:
        return False, None
    dz_ft = _mm_to_internal(max(0.1, float(probe_mm)))
    half_w_mm = float(bar_nominal_mm) / 2.0 + float(_EMBED_PROBE_XY_MARGIN_MM)
    half_w_mm = max(half_w_mm, float(_EMBED_PROBE_MIN_HALF_SIDE_MM))
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
        hit_fund = False
        solids = None
        fid = _element_id_int(getattr(fund, "Id", None))
        if solids_cache is not None and fid is not None:
            solids = solids_cache.get(int(fid))
        if solids is None:
            solids = _solids_list_element(fund, geom_opts)
        for sd in solids or []:
            if _solidos_intersectan_volumen(probe, sd):
                hit_fund = True
                break
        if hit_fund:
            any_hit = True
            if h_mm is not None and float(h_mm) > h_max:
                h_max = float(h_mm)
    if not any_hit:
        return False, None
    return True, h_max if h_max > 0.1 else None


def _estiramiento_fundacion_pie_por_colision(
    doc,
    rebar,
    host,
    foundations,
    geom_opts,
    solids_cache=None,
):
    """
    Estiramiento en pie (solo barras verticales) por colisión en el pie.

    :returns: ``(mm, colisiona_fundacion)`` — ``mm`` positivo estira pie; si no hay
        colisión, ``mm = 25 + Ø/2`` para retraer el pie.
    """
    if rebar is None or host is None or not foundations:
        return None
    if _rebar_es_vertical_por_criterio is None:
        return None
    try:
        if not _rebar_es_vertical_por_criterio(rebar, host, 0):
            return None
    except Exception:
        return None

    d_mm = _nominal_diameter_mm_from_rebar(rebar, doc)
    if d_mm is None or float(d_mm) <= 0.0:
        return None

    p0, p1 = _rebar_eje_p_start_p_end(rebar, 0)
    xyz_ref = _punto_pie_vertical(p0, p1)
    if xyz_ref is None:
        return None

    base_mm = float(FOUNDATION_PROBE_BASE_MM)
    collides, h_fund_mm = _probe_colision_fundacion_desde_punto(
        xyz_ref, base_mm, d_mm, foundations, geom_opts,
        solids_cache=solids_cache,
    )
    if collides and h_fund_mm is not None:
        extra = float(h_fund_mm) - base_mm - float(FOUNDATION_STRETCH_RESTA_MM) - float(d_mm) / 2.0
        stretch = max(base_mm, base_mm + max(0.0, extra))
        return stretch, True
    return _retract_mm_sin_colision(d_mm), False


def _pata_l_mm_desde_diametro(d_mm, concrete_grade=None):
    """Largo pata L (mm) según tabla BIMTools de patas/ganchos por Ø."""
    if hook_length_mm_from_nominal_diameter_mm is None:
        return None
    g = CONCRETE_GRADE if concrete_grade is None else concrete_grade
    try:
        return float(hook_length_mm_from_nominal_diameter_mm(d_mm, g))
    except Exception:
        return None


def _pata_l_eje_sketch_mm_desde_tabla(tabla_mm, d_mm):
    """
    Longitud del tramo de eje para boceto (CreateFromCurves).

    Revit modela la pata desde el eje de la barra; si el eje mide lo mismo que la
    tabla, la geometría resultante queda ~tabla + Ø/2. Restamos medio diámetro
    nominal para que el largo modelado coincida con el valor de tabla.
    """
    if tabla_mm is None:
        return None
    try:
        Ltab = float(tabla_mm)
    except Exception:
        return None
    if Ltab < 0.1:
        return None
    if pata_eje_curve_loop_mm_desde_tabla_mm is not None:
        try:
            Leje = pata_eje_curve_loop_mm_desde_tabla_mm(Ltab, d_mm)
            if Leje is not None:
                return float(Leje)
        except Exception:
            pass
    try:
        d = float(int(round(float(d_mm))))
    except Exception:
        d = 0.0
    if d > 1e-6:
        return max(40.0, Ltab - 0.5 * d)
    return Ltab


def _retract_mm_sin_colision(d_mm):
    """Estiramiento negativo en cabeza cuando no hay colisión: 25 mm + Ø/2."""
    return float(NO_COLLISION_RETRACT_BASE_MM) + float(d_mm) / 2.0


def _capas_horizontales_muro_keys(muro_contencion=False):
    u"""Capas con barras horizontales: major en muro tradicional, minor en contención."""
    if muro_contencion:
        return (u"exterior_minor", u"interior_minor")
    return (u"exterior_major", u"interior_major")


def _diametro_nominal_bar_type_mm(doc, bar_type_id):
    if doc is None or bar_type_id is None:
        return 0.0
    try:
        if bar_type_id == ElementId.InvalidElementId:
            return 0.0
    except Exception:
        pass
    try:
        bt = doc.GetElement(bar_type_id)
    except Exception:
        return 0.0
    if not isinstance(bt, RebarBarType):
        return 0.0
    try:
        d = bt.BarModelDiameter
        return float(UnitUtils.ConvertFromInternalUnits(d, UnitTypeId.Millimeters))
    except Exception:
        return 0.0


def largo_pata_l_vertical_cabeza_mm(
    doc,
    host,
    params_dict=None,
    layer_active_dict=None,
    muro_contencion=False,
):
    """
    Pata L en cabeza (sin colisión): espesor − 50 mm − Ø horiz. ext. − Ø horiz. int.
    ``params_dict`` / ``layer_active_dict``: configuración de malla del panel (por muro).
    """
    if doc is None or host is None:
        return None
    th_mm = None
    if l135 is not None:
        try:
            th_mm = l135._obtener_espesor_host_mm(doc, host)
        except Exception:
            th_mm = None
    if th_mm is None:
        if l135 is not None:
            try:
                return float(l135.largo_pata_mm_desde_espesor_host(doc, host))
            except Exception:
                pass
        return None

    d_ext = d_int = 0.0
    if params_dict is not None and layer_active_dict is not None:
        k_ext, k_int = _capas_horizontales_muro_keys(muro_contencion)
        if layer_active_dict.get(k_ext, True):
            tup = params_dict.get(k_ext) or (None, u"")
            try:
                bid = tup[0]
            except Exception:
                bid = None
            d_ext = _diametro_nominal_bar_type_mm(doc, bid)
        if layer_active_dict.get(k_int, True):
            tup = params_dict.get(k_int) or (None, u"")
            try:
                bid = tup[0]
            except Exception:
                bid = None
            d_int = _diametro_nominal_bar_type_mm(doc, bid)

    largo = (
        float(th_mm)
        - float(PATA_L_RECUBRIMIENTO_CARAS_MM)
        - float(d_ext)
        - float(d_int)
    )
    min_largo = 10.0
    if l135 is not None:
        try:
            min_largo = float(getattr(l135, u"PATA_LARGO_MIN_MM", 10.0))
        except Exception:
            pass
    return max(min_largo, largo)


def _rebar_es_vertical_interior(rebar, host_wall):
    if rebar is None or host_wall is None:
        return False
    if _rebar_es_vertical_por_criterio is None or _rebar_solo_cara_interior is None:
        return False
    if not isinstance(rebar, Rebar):
        return False
    if not _rebar_es_vertical_por_criterio(rebar, host_wall, 0):
        return False
    if _rebar_solo_cara_exterior is not None and _rebar_solo_cara_exterior(rebar, host_wall):
        return False
    return bool(_rebar_solo_cara_interior(rebar, host_wall))


def _aplicar_pata_l_pie_fundacion(doc, rebar, host, res, solo_interior=False):
    """Pata L en pie tras evaluación de fundación (tabla patas BIMTools por Ø)."""
    if rebar is None or host is None:
        return rebar
    if solo_interior:
        if not _rebar_es_vertical_interior(rebar, host):
            return rebar
    elif not _rebar_es_vertical_cara_ext_o_int(rebar, host):
        return rebar

    d_mm = _nominal_diameter_mm_from_rebar(rebar, doc)
    if d_mm is None:
        return rebar
    largo_tabla = _pata_l_mm_desde_diametro(d_mm)
    largo_l = _pata_l_eje_sketch_mm_desde_tabla(largo_tabla, d_mm)
    if largo_l is None or float(largo_l) < 0.1:
        return rebar

    es_exterior = _rebar_es_vertical_exterior(rebar, host)
    invertir = bool(getattr(l135, u"INVERTIR_DIRECCION_PATA", False))
    if not es_exterior:
        invertir = not invertir

    if l135 is not None:
        pata_en_final = l135.pata_en_extremo_final_para_pie_por_elevacion(rebar, 0)
    else:
        pata_en_final = False

    ok_l, msg_l, rb_l = _agregar_pata_l_extremo_sketch(
        doc,
        rebar,
        host,
        float(largo_l),
        pata_en_final,
        invertir,
        u"Arainco: Armado muros lineales — pata L pie fundación vertical",
        0,
    )
    if ok_l:
        res[u"n_pata_l_fund_pie"] += 1
        return rb_l if rb_l is not None else rebar

    res[u"n_fail"] += 1
    rid = _element_id_int(getattr(rebar, "Id", None))
    res[u"messages"].append(
        u"Rebar {0} (pata L fund. pie): {1}".format(rid, msg_l or u"error"),
    )
    return rebar


def _aplicar_pata_l_pie_muro_sin_fundacion(
    doc,
    rebar,
    host,
    res,
    params_dict=None,
    layer_active_dict=None,
    muro_contencion=False,
):
    """Pata L en pie (sin fundación, sin colisión vs muros): ext. e int."""
    if rebar is None or host is None:
        return rebar
    if not _rebar_es_vertical_cara_ext_o_int(rebar, host):
        return rebar

    largo_l = largo_pata_l_vertical_cabeza_mm(
        doc, host, params_dict, layer_active_dict, muro_contencion,
    )
    if largo_l is None or float(largo_l) < 0.1:
        return rebar

    es_exterior = _rebar_es_vertical_exterior(rebar, host)
    invertir = bool(getattr(l135, u"INVERTIR_DIRECCION_PATA", False))
    if not es_exterior:
        invertir = not invertir

    if l135 is not None:
        pata_en_final = l135.pata_en_extremo_final_para_pie_por_elevacion(rebar, 0)
    else:
        pata_en_final = False

    ok_l, msg_l, rb_l = _agregar_pata_l_extremo_sketch(
        doc,
        rebar,
        host,
        float(largo_l),
        pata_en_final,
        invertir,
        u"Arainco: Armado muros lineales — pata L pie muro sin fundación",
        0,
    )
    if ok_l:
        res[u"n_pie_muro_pata_l"] += 1
        if es_exterior:
            res[u"n_pie_muro_pata_l_ext"] = int(res.get(u"n_pie_muro_pata_l_ext", 0)) + 1
        elif _rebar_es_vertical_interior(rebar, host):
            res[u"n_pie_muro_pata_l_int"] = int(res.get(u"n_pie_muro_pata_l_int", 0)) + 1
        return rb_l if rb_l is not None else rebar

    res[u"n_fail"] += 1
    rid = _element_id_int(getattr(rebar, "Id", None))
    res[u"messages"].append(
        u"Rebar {0} (pata L pie sin fund.): {1}".format(rid, msg_l or u"error"),
    )
    return rebar


def _procesar_rebar_vertical_pie_colision_muro_sin_fundacion(
    doc,
    rebar,
    host,
    walls,
    geom_opts,
    res,
    params_dict=None,
    layer_active_dict=None,
    muro_contencion=False,
    ids_con_apilado_bajo=None,
    solids_cache=None,
    bbox_z_cache=None,
):
    """
    Pie sin fundación: si hay muro apilado debajo → no mutar; si no →
    retraer ``25+Ø/2`` + pata L.
    """
    if not _rebar_es_vertical_cara_ext_o_int(rebar, host):
        res[u"n_skip"] += 1
        return rebar

    d_mm = _nominal_diameter_mm_from_rebar(rebar, doc)
    if d_mm is None:
        res[u"n_fail"] += 1
        return rebar

    if _host_tiene_apilado_bajo(host, ids_con_apilado_bajo):
        res[u"n_pie_muro_colision_revert"] += 1
        return rebar

    retract_mm = _retract_mm_sin_colision(d_mm)
    ok, msg, rb_out = _acortar_vertical_pie_mm(doc, rebar, retract_mm, 0)
    if not ok:
        res[u"n_fail"] += 1
        rid = _element_id_int(getattr(rebar, "Id", None))
        res[u"messages"].append(
            u"Rebar {0} (pie sin muro apilado): {1}".format(rid, msg or u"error"),
        )
        return rebar

    res[u"n_pie_muro_retract"] += 1
    rb_work = rb_out if rb_out is not None else rebar
    return _aplicar_pata_l_pie_muro_sin_fundacion(
        doc,
        rb_work,
        host,
        res,
        params_dict,
        layer_active_dict,
        muro_contencion,
    )


def _aplicar_estiramiento_fundacion_pie(
    doc, rebar, host, foundations, geom_opts, res, solids_cache=None,
):
    """
    Evalúa colisión en pie (100 mm) en verticales con fundación unida.
    Colisión → estira pie (+ pata L ext/int); sin colisión → retrae pie 25 mm + Ø/2
    (+ pata L solo interior).
    """
    if not foundations or rebar is None or host is None:
        return rebar
    if _rebar_es_vertical_por_criterio is None:
        return rebar
    try:
        if not _rebar_es_vertical_por_criterio(rebar, host, 0):
            return rebar
    except Exception:
        return rebar

    eval_res = _estiramiento_fundacion_pie_por_colision(
        doc, rebar, host, foundations, geom_opts,
        solids_cache=solids_cache,
    )
    if eval_res is None:
        return rebar
    mm_accion, collided = eval_res
    if mm_accion is None or float(mm_accion) < 0.1:
        return rebar

    if collided:
        ok, msg, rb_out = _extender_vertical_pie_mm(doc, rebar, mm_accion, 0)
        if not ok:
            res[u"n_fail"] += 1
            rid = _element_id_int(getattr(rebar, "Id", None))
            res[u"messages"].append(
                u"Rebar {0} (fundación pie vertical): {1}".format(rid, msg or u"error al estirar"),
            )
            return rebar
        res[u"n_fundacion_pie"] += 1
        rb_work = rb_out if rb_out is not None else rebar
    else:
        ok, msg, rb_out = _acortar_vertical_pie_mm(doc, rebar, mm_accion, 0)
        if not ok:
            res[u"n_fail"] += 1
            rid = _element_id_int(getattr(rebar, "Id", None))
            res[u"messages"].append(
                u"Rebar {0} (fundación pie sin colisión): {1}".format(
                    rid, msg or u"error al retraer",
                ),
            )
            return rebar
        res[u"n_fundacion_retract"] += 1
        rb_work = rb_out if rb_out is not None else rebar
        return _aplicar_pata_l_pie_fundacion(doc, rb_work, host, res, solo_interior=True)

    return _aplicar_pata_l_pie_fundacion(doc, rb_work, host, res)


def _acortar_vertical_pie_mm(doc, rebar, mm_retiro, pos_idx=0):
    """
    Acorta el eje en el extremo de menor Z (pie) ``mm_retiro`` mm hacia la cabeza.
    """
    if doc is None or rebar is None:
        return False, u"Doc o rebar no válido.", None
    m = max(0.0, float(mm_retiro))
    if m < 0.1:
        return True, u"Retiro 0 mm, sin cambio.", rebar
    if (
        _tangent_start_curve is None
        or _create_from_curves_no_hooks is None
        or _copy_layout_rebar_shape_driven is None
    ):
        return False, u"Módulo de acortamiento no disponible.", None

    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        crvs = rebar.GetCenterlineCurves(False, False, False, mpo, int(pos_idx))
    except Exception as ex:
        return False, u"GetCenterlineCurves: {0!s}".format(ex), None
    if crvs is None or int(crvs.Count) < 1:
        return False, u"Sin curvas de eje (pos. {0}).".format(pos_idx), None

    chain = [crvs[i] for i in range(crvs.Count)]
    c_first = chain[0]
    c_last = chain[-1]
    t0 = _tangent_start_curve(c_first)
    t1 = _tangent_at_end_of_curve(c_last)
    if t0 is None or t1 is None:
        return False, u"Tangente nula (geometría de eje).", None

    p0s = c_first.GetEndPoint(0)
    p0e = c_first.GetEndPoint(1)
    p1s = c_last.GetEndPoint(0)
    p1e = c_last.GetEndPoint(1)
    d_internal = _mm_to_internal(m)

    try:
        eps = _mm_to_internal(0.5)
    except Exception:
        eps = 1.0e-4
    cabeza_en_fin = float(p1e.Z) >= float(p0s.Z) - float(eps)

    if cabeza_en_fin:
        new_p0 = p0s + t0.Multiply(d_internal)
        new_p1e = p1e
    else:
        new_p0 = p0s
        new_p1e = p1e - t1.Multiply(d_internal)

    if int(crvs.Count) == 1:
        try:
            if new_p0.DistanceTo(new_p1e) < 1e-6:
                return False, u"Retiro deja barra nula.", None
        except Exception:
            pass
        c_new = Line.CreateBound(new_p0, new_p1e)
        new_chain = [c_new]
    else:
        c_first_new = Line.CreateBound(new_p0, p0e)
        c_last_new = Line.CreateBound(p1s, new_p1e)
        new_chain = [c_first_new] + chain[1:-1] + [c_last_new]
        for i in range(len(new_chain) - 1):
            e_prev = new_chain[i].GetEndPoint(1)
            s_next = new_chain[i + 1].GetEndPoint(0)
            gap = (e_prev - s_next).GetLength()
            if gap > 0.01:
                return (
                    False,
                    u"Polilínea no consecutiva (gap {0:,.4f} ft).".format(gap),
                    None,
                )

    host = doc.GetElement(rebar.GetHostId())
    if host is None:
        return False, u"Host inválido.", None
    bar_type = doc.GetElement(rebar.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return False, u"RebarBarType no resuelto.", None
    try:
        style = rebar.Style
    except Exception:
        style = RebarStyle.Standard
    norm = _rebar_normal(rebar)
    o0 = _hook_orient_for_create(rebar, 0)
    o1 = _hook_orient_for_create(rebar, 1)

    from armado_muros_txn import TxnScope

    scope = TxnScope(
        doc, u"Arainco: Armado muros lineales — retraer pie vertical fundación",
    )
    try:
        new_rb = _create_from_curves_no_hooks(
            doc, new_chain, host, norm, bar_type, style, o0, o1,
        )
        if new_rb is None:
            scope.rollback()
            return False, u"CreateFromCurves devolvió None.", None
        ok_lay, err_lay, new_rb = _copy_layout_rebar_y_excluir_extremos(doc, rebar, new_rb)
        if not ok_lay:
            scope.rollback()
            return False, u"Layout: {0}".format(err_lay or u"?"), None
        try:
            doc.Delete(rebar.Id)
        except Exception as ex2:
            scope.rollback()
            return False, u"Delete rebar: {0!s}".format(ex2), None
        try:
            from armado_muros_rebar_params import stamp_malla_vertical_rebar
            stamp_malla_vertical_rebar(new_rb)
        except Exception:
            pass
        scope.commit()
    except Exception as ex:
        scope.rollback()
        return False, u"{0!s}".format(ex), None
    return (
        True,
        u"Retiro pie {0} mm; nuevo id {1}.".format(
            int(round(m)), _element_id_int(getattr(new_rb, "Id", None)),
        ),
        new_rb,
    )


def _acortar_vertical_cabeza_mm(doc, rebar, mm_retiro, pos_idx=0):
    """
    Acorta el eje en el extremo de mayor Z (cabeza) ``mm_retiro`` mm hacia el pie.
    Sustituye el Rebar conservando layout (misma lógica que extender, sentido inverso).
    """
    if doc is None or rebar is None:
        return False, u"Doc o rebar no válido.", None
    m = max(0.0, float(mm_retiro))
    if m < 0.1:
        return True, u"Retiro 0 mm, sin cambio.", rebar
    if (
        _tangent_start_curve is None
        or _create_from_curves_no_hooks is None
        or _copy_layout_rebar_shape_driven is None
    ):
        return False, u"Módulo de acortamiento no disponible.", None

    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        crvs = rebar.GetCenterlineCurves(False, False, False, mpo, int(pos_idx))
    except Exception as ex:
        return False, u"GetCenterlineCurves: {0!s}".format(ex), None
    if crvs is None or int(crvs.Count) < 1:
        return False, u"Sin curvas de eje (pos. {0}).".format(pos_idx), None

    chain = [crvs[i] for i in range(crvs.Count)]
    c_first = chain[0]
    c_last = chain[-1]
    t0 = _tangent_start_curve(c_first)
    t1 = _tangent_at_end_of_curve(c_last)
    if t0 is None or t1 is None:
        return False, u"Tangente nula (geometría de eje).", None

    p0s = c_first.GetEndPoint(0)
    p0e = c_first.GetEndPoint(1)
    p1s = c_last.GetEndPoint(0)
    p1e = c_last.GetEndPoint(1)
    d_internal = _mm_to_internal(m)

    try:
        eps = _mm_to_internal(0.5)
    except Exception:
        eps = 1.0e-4
    cabeza_en_fin = float(p1e.Z) >= float(p0s.Z) - float(eps)

    if cabeza_en_fin:
        new_p0 = p0s
        new_p1e = p1e - t1.Multiply(d_internal)
    else:
        new_p0 = p0s + t0.Multiply(d_internal)
        new_p1e = p1e

    if int(crvs.Count) == 1:
        try:
            if new_p0.DistanceTo(new_p1e) < 1e-6:
                return False, u"Retiro deja barra nula.", None
        except Exception:
            pass
        c_new = Line.CreateBound(new_p0, new_p1e)
        new_chain = [c_new]
    else:
        c_first_new = Line.CreateBound(new_p0, p0e)
        c_last_new = Line.CreateBound(p1s, new_p1e)
        new_chain = [c_first_new] + chain[1:-1] + [c_last_new]
        for i in range(len(new_chain) - 1):
            e_prev = new_chain[i].GetEndPoint(1)
            s_next = new_chain[i + 1].GetEndPoint(0)
            gap = (e_prev - s_next).GetLength()
            if gap > 0.01:
                return (
                    False,
                    u"Polilínea no consecutiva (gap {0:,.4f} ft).".format(gap),
                    None,
                )

    host = doc.GetElement(rebar.GetHostId())
    if host is None:
        return False, u"Host inválido.", None
    bar_type = doc.GetElement(rebar.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return False, u"RebarBarType no resuelto.", None
    try:
        style = rebar.Style
    except Exception:
        style = RebarStyle.Standard
    norm = _rebar_normal(rebar)
    o0 = _hook_orient_for_create(rebar, 0)
    o1 = _hook_orient_for_create(rebar, 1)

    from armado_muros_txn import TxnScope

    scope = TxnScope(
        doc, u"Arainco: Armado muros lineales — retraer cabeza vertical",
    )
    try:
        new_rb = _create_from_curves_no_hooks(
            doc, new_chain, host, norm, bar_type, style, o0, o1,
        )
        if new_rb is None:
            scope.rollback()
            return False, u"CreateFromCurves devolvió None.", None
        ok_lay, err_lay, new_rb = _copy_layout_rebar_y_excluir_extremos(doc, rebar, new_rb)
        if not ok_lay:
            scope.rollback()
            return False, u"Layout: {0}".format(err_lay or u"?"), None
        try:
            doc.Delete(rebar.Id)
        except Exception as ex2:
            scope.rollback()
            return False, u"Delete rebar: {0!s}".format(ex2), None
        try:
            from armado_muros_rebar_params import stamp_malla_vertical_rebar
            stamp_malla_vertical_rebar(new_rb)
        except Exception:
            pass
        scope.commit()
    except Exception as ex:
        scope.rollback()
        return False, u"{0!s}".format(ex), None
    return (
        True,
        u"Retiro cabeza {0} mm; nuevo id {1}.".format(
            int(round(m)), _element_id_int(getattr(new_rb, "Id", None)),
        ),
        new_rb,
    )


def embed_stretch_collides_any_wall_solids(
    doc,
    xyz_top,
    dz_embed_ft,
    bar_nominal_mm,
    wall_obstacles,
    host_wall_id,
    geom_opts,
    solids_cache=None,
    bbox_z_cache=None,
):
    """
    True si el prisma de ensayo (+Z desde ``xyz_top.Z``) intersecta algún sólido
    de los muros en ``wall_obstacles``, omitiendo el host.

    ``solids_cache`` / ``bbox_z_cache``: opcionales; evitan ``get_Geometry`` repetido
    y permiten filtrar por bbox Z (mismos resultados si la caché está fresca).
    """
    dz_e = abs(float(dz_embed_ft))
    if doc is None or xyz_top is None or dz_e <= 1e-12:
        return False
    z0 = float(xyz_top.Z)
    half_w_mm = float(bar_nominal_mm) / 2.0 + float(_EMBED_PROBE_XY_MARGIN_MM)
    half_w_mm = max(half_w_mm, float(_EMBED_PROBE_MIN_HALF_SIDE_MM))
    half_w_ft = _mm_to_internal(half_w_mm)
    probe = _build_vertical_square_prism_solid(
        float(xyz_top.X), float(xyz_top.Y), z0, half_w_ft, dz_e,
    )
    if probe is None:
        return False
    for sd in _iter_obstacle_solids(
        wall_obstacles,
        host_wall_id,
        geom_opts,
        solids_cache=solids_cache,
        bbox_z_cache=bbox_z_cache,
        z_lo=z0,
        z_hi=z0 + dz_e,
    ):
        if _solidos_intersectan_volumen(probe, sd):
            return True
    return False


def embed_stretch_collides_wall_solids_downward(
    doc,
    xyz_pie,
    dz_down_ft,
    bar_nominal_mm,
    wall_obstacles,
    host_wall_id,
    geom_opts,
    solids_cache=None,
    bbox_z_cache=None,
):
    """
    True si el prisma de ensayo (−Z desde ``xyz_pie.Z``) intersecta algún sólido
    de los muros en ``wall_obstacles``, omitiendo el host.
    """
    dz_e = abs(float(dz_down_ft))
    if doc is None or xyz_pie is None or dz_e <= 1e-12:
        return False
    half_w_mm = float(bar_nominal_mm) / 2.0 + float(_EMBED_PROBE_XY_MARGIN_MM)
    half_w_mm = max(half_w_mm, float(_EMBED_PROBE_MIN_HALF_SIDE_MM))
    half_w_ft = _mm_to_internal(half_w_mm)
    z0 = float(xyz_pie.Z)
    probe = _build_vertical_prism_downward(
        float(xyz_pie.X), float(xyz_pie.Y), z0, half_w_ft, dz_e,
    )
    if probe is None:
        return False
    for sd in _iter_obstacle_solids(
        wall_obstacles,
        host_wall_id,
        geom_opts,
        solids_cache=solids_cache,
        bbox_z_cache=bbox_z_cache,
        z_lo=z0 - dz_e,
        z_hi=z0,
    ):
        if _solidos_intersectan_volumen(probe, sd):
            return True
    return False


def _excluir_extremos_rebar_set(doc, rebar, host=None):
    """
    Excluye extremos según orientación (horizontal: última; vertical: 1.ª y última).
    """
    if doc is None or rebar is None or ajustar_inclusion_extremos_rebar_set_con_fallback is None:
        return rebar
    if host is None:
        try:
            host = doc.GetElement(rebar.GetHostId())
        except Exception:
            host = None
    try:
        if (
            _rebar_es_vertical_por_criterio is not None
            and host is not None
            and _rebar_es_vertical_por_criterio(rebar, host, 0)
        ):
            ajustar_inclusion_extremos_rebar_set_con_fallback(rebar, doc, False, False)
        else:
            ajustar_inclusion_extremos_rebar_set_con_fallback(rebar, doc, True, False)
    except Exception:
        pass
    try:
        rb = doc.GetElement(rebar.Id)
        if rb is not None and isinstance(rb, Rebar):
            return rb
    except Exception:
        pass
    return rebar


def _excluir_ambos_extremos_rebar_set(doc, rebar):
    """Excluye 1.ª y últ. posición (sets verticales recreados en post-proceso)."""
    if doc is None or rebar is None or desactivar_extremos_rebar_set is None:
        return rebar
    try:
        desactivar_extremos_rebar_set(rebar, doc)
    except Exception:
        pass
    try:
        rb = doc.GetElement(rebar.Id)
        if rb is not None and isinstance(rb, Rebar):
            return rb
    except Exception:
        pass
    return rebar


def _copy_layout_rebar_y_excluir_extremos(doc, src, dst):
    """Copia layout shape-driven (exclusión por cabezal en post-proceso lineales)."""
    if _copy_layout_rebar_shape_driven is None:
        return False, u"Copy layout no disponible.", dst
    ok_lay, err_lay = _copy_layout_rebar_shape_driven(src, dst)
    if not ok_lay:
        return False, err_lay or u"?", dst
    return True, u"", dst


def _rebar_es_vertical_exterior(rebar, host_wall):
    if rebar is None or host_wall is None:
        return False
    if _rebar_es_vertical_por_criterio is None or _rebar_solo_cara_exterior is None:
        return False
    if not isinstance(rebar, Rebar):
        return False
    if not _rebar_es_vertical_por_criterio(rebar, host_wall, 0):
        return False
    return bool(_rebar_solo_cara_exterior(rebar, host_wall))


def _punto_pie_y_cabeza_en_cadena(chain):
    """Extremos de menor y mayor cota Z en la polilínea del eje."""
    p_pie = p_cab = None
    z_pie = z_cab = None
    for c in chain:
        if c is None:
            continue
        for i in (0, 1):
            try:
                p = c.GetEndPoint(i)
                z = float(p.Z)
            except Exception:
                continue
            if p_pie is None or z < z_pie:
                z_pie, p_pie = z, p
            if p_cab is None or z > z_cab:
                z_cab, p_cab = z, p
    return p_pie, p_cab


def _ordenar_cadena_desde_hasta(chain, p_start, p_end, tol_ft=None):
    """Ordena tramos conectados desde ``p_start`` hasta ``p_end``."""
    if not chain or p_start is None or p_end is None:
        return None
    if tol_ft is None:
        tol_ft = _mm_to_internal(0.5)
    remaining = list(chain)
    out = []
    cur = p_start

    def _near(a, b):
        try:
            return a.DistanceTo(b) < tol_ft
        except Exception:
            return False

    while remaining:
        found = False
        for i, c in enumerate(remaining):
            try:
                p0 = c.GetEndPoint(0)
                p1 = c.GetEndPoint(1)
            except Exception:
                continue
            if _near(p0, cur):
                out.append(c)
                cur = p1
                remaining.pop(i)
                found = True
                break
            if _near(p1, cur):
                out.append(Line.CreateBound(p1, p0))
                cur = p0
                remaining.pop(i)
                found = True
                break
        if not found:
            return None
    if not _near(cur, p_end):
        return None
    return out


def _construir_cadena_l_cabeza_segmento_0_principal(chain, norm, largo_p_mm, invertir):
    """
    Polilínea en L con pata en cabeza (mayor Z): segmento 0 = tramo principal pie→cabeza,
    segmento 1 = pata L (no al revés).
    """
    if not chain:
        return None, u"Cadena vacía."
    p_pie, p_cab = _punto_pie_y_cabeza_en_cadena(chain)
    if p_pie is None or p_cab is None:
        return None, u"No se resolvieron pie/cabeza por cota Z."
    le = _mm_to_internal(float(largo_p_mm))
    try:
        t_vec = p_cab.Subtract(p_pie)
    except Exception:
        return None, u"Vector pie→cabeza inválido."
    if t_vec.GetLength() < 1e-9:
        return None, u"Tramo vertical nulo."
    t_vec = t_vec.Normalize()
    b_vec = l135._perp_in_plane(norm, t_vec)
    if b_vec is None:
        return None, u"Perpendicular in-plane nula."
    if invertir:
        b_vec = b_vec.Negate()
    # Con eje pie→cabeza el «−b» histórico apunta hacia afuera; «+b» entra al muro.
    p_tip = p_cab + b_vec.Multiply(le)
    leg = Line.CreateBound(p_cab, p_tip)
    if len(chain) == 1:
        try:
            main_line = Line.CreateBound(p_pie, p_cab)
        except Exception as ex:
            return None, u"Eje principal: {0!s}".format(ex)
        return [main_line, leg], None
    ordered = _ordenar_cadena_desde_hasta(chain, p_pie, p_cab)
    if ordered is None:
        return None, u"Polilínea pie→cabeza no consecutiva."
    return ordered + [leg], None


def _agregar_pata_l_extremo_sketch(
    doc,
    rebar,
    host,
    largo_p_mm,
    pata_en_final,
    invertir,
    txn_name,
    pos_idx=0,
    cabeza_segmento_0_principal=False,
):
    """
    Añade pata L en un extremo del boceto (edit sketch: polilínea L + CreateFromCurves).
    """
    if l135 is None or doc is None or rebar is None or host is None:
        return False, u"Módulo pata L no disponible.", None
    if _create_from_curves_no_hooks is None or _copy_layout_rebar_shape_driven is None:
        return False, u"Helpers de boceto no disponibles.", None
    if largo_p_mm is None or float(largo_p_mm) < 0.1:
        return False, u"Largo pata L inválido.", None

    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        crvs = rebar.GetCenterlineCurves(False, False, False, mpo, int(pos_idx))
    except Exception as ex:
        return False, u"GetCenterlineCurves: {0!s}".format(ex), None
    if crvs is None or int(crvs.Count) < 1:
        return False, u"Sin curvas de eje.", None

    chain = [crvs[i] for i in range(crvs.Count)]
    le = _mm_to_internal(float(largo_p_mm))
    norm = l135._rebar_normal(rebar)
    if bool(getattr(l135, u"INVERTIR_NORMAL_REBAR", False)):
        norm = norm.Negate()

    if cabeza_segmento_0_principal:
        new_chain, err_chain = _construir_cadena_l_cabeza_segmento_0_principal(
            chain, norm, float(largo_p_mm), invertir,
        )
        if new_chain is None:
            return False, err_chain or u"Cadena L cabeza inválida.", None
    elif not pata_en_final:
        c0 = chain[0]
        t_vec = l135._tangent_start_first_curve(c0)
        if t_vec is None:
            return False, u"Tangente nula en inicio.", None
        b_vec = l135._perp_in_plane(norm, t_vec)
        if b_vec is None:
            return False, u"Perpendicular in-plane nula.", None
        if invertir:
            b_vec = b_vec.Negate()
        p0 = c0.GetEndPoint(0)
        p_leg = p0 - b_vec.Multiply(le)
        leg = Line.CreateBound(p_leg, p0)
        new_chain = [leg] + chain
    else:
        c_last = chain[-1]
        t_vec = l135._tangent_start_first_curve(c_last)
        if t_vec is None:
            return False, u"Tangente nula en extremo.", None
        b_vec = l135._perp_in_plane(norm, t_vec)
        if b_vec is None:
            return False, u"Perpendicular in-plane nula.", None
        if invertir:
            b_vec = b_vec.Negate()
        p_end = c_last.GetEndPoint(1)
        p_tip = p_end - b_vec.Multiply(le)
        leg = Line.CreateBound(p_end, p_tip)
        new_chain = chain + [leg]

    bar_type = doc.GetElement(rebar.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return False, u"RebarBarType no resuelto.", None
    try:
        style = rebar.Style
    except Exception:
        style = RebarStyle.Standard
    o0 = _hook_orient_for_create(rebar, 0)
    o1 = _hook_orient_for_create(rebar, 1)
    orig_id = rebar.Id

    from armado_muros_txn import TxnScope

    scope = TxnScope(doc, txn_name)
    try:
        new_rb = None
        fn_shape = getattr(l135, u"_try_create_l_from_rebar_shape_2seg", None)
        if callable(fn_shape):
            try:
                new_rb = fn_shape(doc, new_chain, host, norm, bar_type, style, o0, o1)
            except Exception:
                new_rb = None
        if new_rb is None:
            new_rb = _create_from_curves_no_hooks(
                doc, new_chain, host, norm, bar_type, style, o0, o1,
            )
        if new_rb is None:
            scope.rollback()
            return False, u"CreateFromCurves devolvió None.", None
        ok_lay, err_lay, new_rb = _copy_layout_rebar_y_excluir_extremos(doc, rebar, new_rb)
        if not ok_lay:
            scope.rollback()
            return False, u"Layout: {0}".format(err_lay or u"?"), None
        try:
            doc.Delete(orig_id)
        except Exception as ex2:
            scope.rollback()
            return False, u"Delete rebar: {0!s}".format(ex2), None
        try:
            from armado_muros_rebar_params import stamp_malla_vertical_rebar
            stamp_malla_vertical_rebar(new_rb)
        except Exception:
            pass
        scope.commit()
    except Exception as ex:
        scope.rollback()
        return False, u"{0!s}".format(ex), None

    return (
        True,
        u"Pata L {0} mm; nuevo id {1}.".format(
            int(round(float(largo_p_mm))), _element_id_int(getattr(new_rb, "Id", None)),
        ),
        new_rb,
    )


def _agregar_pata_l_cabeza_vertical_sketch(
    doc,
    rebar,
    host,
    pos_idx=0,
    params_dict=None,
    layer_active_dict=None,
    muro_contencion=False,
):
    """
    Pata L en cabeza (vertical ext/int sin colisión en cabeza).
    Largo = espesor − 50 mm − Ø horiz. ext. − Ø horiz. int.; sentido según cara.
    """
    if l135 is None:
        return False, u"Módulo pata L no disponible.", None
    largo_p = largo_pata_l_vertical_cabeza_mm(
        doc, host, params_dict, layer_active_dict, muro_contencion,
    )
    if largo_p is None or float(largo_p) < 0.1:
        return False, (
            u"No se pudo calcular largo pata L "
            u"(espesor − 50 − Ø horiz. ext. − Ø horiz. int.)."
        ), None
    es_exterior = _rebar_es_vertical_exterior(rebar, host)
    invertir = bool(getattr(l135, u"INVERTIR_DIRECCION_PATA", False))
    if not es_exterior:
        invertir = not invertir
    cara_lbl = u"exterior" if es_exterior else u"interior"
    txn = u"Arainco: Armado muros lineales — pata L cabeza vertical {0}".format(
        cara_lbl,
    )
    return _agregar_pata_l_extremo_sketch(
        doc,
        rebar,
        host,
        float(largo_p),
        False,
        invertir,
        txn,
        int(pos_idx),
        cabeza_segmento_0_principal=True,
    )


def _agregar_pata_l_cabeza_exterior_sketch(doc, rebar, host, pos_idx=0):
    """Alias retrocompatible → :func:`_agregar_pata_l_cabeza_vertical_sketch`."""
    return _agregar_pata_l_cabeza_vertical_sketch(doc, rebar, host, pos_idx)


def _rebar_es_vertical_cara_ext_o_int(rebar, host_wall):
    """
    Vertical en muro y cara exterior o interior (exterior tiene prioridad, como
    ``arearein_ambas_caras_h_l135_rps``).
    """
    if rebar is None or host_wall is None:
        return False
    if _rebar_es_vertical_por_criterio is None:
        return False
    if not isinstance(rebar, Rebar):
        return False
    if not _rebar_es_vertical_por_criterio(rebar, host_wall, 0):
        return False
    if _rebar_solo_cara_exterior is not None and _rebar_solo_cara_exterior(rebar, host_wall):
        return True
    if _rebar_solo_cara_interior is not None and _rebar_solo_cara_interior(rebar, host_wall):
        return True
    return False


def _procesar_rebar_vertical_cabeza_colision(
    doc,
    rebar,
    host,
    walls,
    geom_opts,
    concrete_grade,
    res,
    params_dict=None,
    layer_active_dict=None,
    muro_contencion=False,
    ids_con_apilado_sobre=None,
    solids_cache=None,
    bbox_z_cache=None,
):
    """
    Cabeza: empotrar solo si hay muro apilado encima en la selección;
    si no, retraer ``25+Ø/2`` + pata L.
    """
    if not _rebar_es_vertical_cara_ext_o_int(rebar, host):
        res[u"n_skip"] += 1
        return rebar

    es_exterior = _rebar_es_vertical_exterior(rebar, host)

    d_mm = _nominal_diameter_mm_from_rebar(rebar, doc)
    if d_mm is None:
        res[u"n_fail"] += 1
        return rebar

    L_eval_mm = _empotramiento_cabeza_mm_desde_diametro(d_mm, concrete_grade)
    if L_eval_mm is None or L_eval_mm < 0.1:
        res[u"n_skip"] += 1
        return rebar

    if _host_tiene_apilado_sobre(host, ids_con_apilado_sobre):
        ok_eval, msg_eval, rebar_eval = _extender_vertical_cabeza_tabla_empotramiento(
            doc, rebar, 0, concrete_grade,
        )
        if not ok_eval:
            res[u"n_fail"] += 1
            rid = _element_id_int(getattr(rebar, "Id", None))
            res[u"messages"].append(
                u"Rebar {0} (empotramiento por apilado): {1}".format(
                    rid, msg_eval or u"error al estirar",
                ),
            )
            return rebar
        res[u"n_extended"] += 1
        return rebar_eval if rebar_eval is not None else rebar

    retract_mm = _retract_mm_sin_colision(d_mm)
    ok, msg, rebar_final = _acortar_vertical_cabeza_mm(
        doc, rebar, retract_mm, 0,
    )
    if ok:
        res[u"n_retracted"] += 1
        rebar_out = rebar_final if rebar_final is not None else rebar
        if rebar_out is not None:
            ok_l, msg_l, rb_l = _agregar_pata_l_cabeza_vertical_sketch(
                doc,
                rebar_out,
                host,
                0,
                params_dict,
                layer_active_dict,
                muro_contencion,
            )
            if ok_l:
                res[u"n_pata_l"] += 1
                if rb_l is not None:
                    rebar_out = rb_l
            else:
                res[u"n_fail"] += 1
                rid = _element_id_int(getattr(rebar_out, "Id", None))
                cara_lbl = u"exterior" if es_exterior else u"interior"
                res[u"messages"].append(
                    u"Rebar {0} (pata L cabeza {1}): {2}".format(
                        rid, cara_lbl, msg_l or u"error",
                    ),
                )
        return rebar_out

    res[u"n_fail"] += 1
    rid = _element_id_int(getattr(rebar, "Id", None))
    res[u"messages"].append(
        u"Rebar {0} (sin apilado en cabeza): {1}".format(
            rid, msg or u"error al retraer",
        ),
    )
    return rebar


def aplicar_empotramiento_verticales_cara_por_colision(
    doc,
    walls,
    rebars_por_muro_id,
    concrete_grade=None,
    evaluar_colision_cabeza=True,
    params_por_muro_id=None,
    muro_contencion=False,
    cabezal_por_muro_id=None,
):
    """
    Post-proceso de verticales ext/int: fundación unida (pie) y, si aplica, empotramiento
    en cabeza según **muro apilado encima** en la selección (orden Z / contacto bbox).

    ``rebars_por_muro_id``: ``{ wall_id_int: [ElementId, ...], ... }``
    ``evaluar_colision_cabeza``: si ``False`` (herramienta Cabezal muros), no ejecuta
    estiramiento/retraída por apilamiento en cabeza.

    Retorna dict con contadores y mensajes.
    """
    _reload_embed_vendor_modules()
    res = {
        u"n_extended": 0,
        u"n_retracted": 0,
        u"n_pata_l": 0,
        u"n_pata_l_fund_pie": 0,
        u"n_fundacion_pie": 0,
        u"n_fundacion_retract": 0,
        u"n_pie_muro_colision_revert": 0,
        u"n_pie_muro_retract": 0,
        u"n_pie_muro_pata_l": 0,
        u"n_pie_muro_pata_l_ext": 0,
        u"n_pie_muro_pata_l_int": 0,
        u"n_skip": 0,
        u"n_fail": 0,
        u"messages": [],
    }
    if doc is None or not walls or not rebars_por_muro_id:
        return res
    if (
        traslape_mm_from_nominal_diameter_mm is None
        or _extender_rebar_por_eje_mm is None
        or _rebar_eje_p_start_p_end is None
    ):
        res[u"messages"].append(
            u"Empotramiento vertical ext/int: módulos locales no disponibles.",
        )
        return res
    if _rebar_solo_cara_exterior is None and _rebar_solo_cara_interior is None:
        res[u"messages"].append(
            u"Empotramiento vertical ext/int: filtros de cara no disponibles.",
        )
        return res

    wall_by_id = {}
    for w in walls:
        try:
            _wi = _element_id_int(w.Id)
            if _wi is not None:
                wall_by_id[int(_wi)] = w
        except Exception:
            pass

    geom_opts = _geometry_options()
    g = CONCRETE_GRADE if concrete_grade is None else concrete_grade

    def _run_vertical_lote():
        # Apilamiento en Z (bbox) una vez por lote — sin sonda de sólidos por barra.
        ids_sobre, ids_bajo = _build_apilamiento_maps(walls)
        fund_solids = {}
        # Fundación sigue usando sólidos; cachear tras un Regenerate si hace falta.
        need_fund_geom = False
        for wid0 in rebars_por_muro_id:
            host0 = wall_by_id.get(int(wid0))
            if host0 is not None and _fundaciones_estructurales_unidas_muro(doc, host0):
                need_fund_geom = True
                break
        if need_fund_geom:
            try:
                doc.Regenerate()
            except Exception:
                pass

        for wid, eid_list in rebars_por_muro_id.items():
            host = wall_by_id.get(int(wid))
            if host is None:
                continue
            params_dict = None
            layer_active_dict = None
            if params_por_muro_id:
                try:
                    tup = params_por_muro_id.get(int(wid))
                    if tup is not None:
                        params_dict, layer_active_dict = tup
                except Exception:
                    pass
            foundations = _fundaciones_estructurales_unidas_muro(doc, host)
            if foundations:
                for fund in foundations:
                    fid = _element_id_int(getattr(fund, "Id", None))
                    if fid is None or int(fid) in fund_solids:
                        continue
                    fund_solids[int(fid)] = _solids_list_element(fund, geom_opts)
            for idx, eid in enumerate(eid_list or []):
                rebar = doc.GetElement(eid)
                if rebar is None or not isinstance(rebar, Rebar):
                    res[u"n_skip"] += 1
                    continue
                rebar_work = rebar
                if evaluar_colision_cabeza:
                    rebar_work = _procesar_rebar_vertical_cabeza_colision(
                        doc,
                        rebar_work,
                        host,
                        walls,
                        geom_opts,
                        g,
                        res,
                        params_dict,
                        layer_active_dict,
                        muro_contencion,
                        ids_con_apilado_sobre=ids_sobre,
                    )
                if foundations:
                    rebar_work = _aplicar_estiramiento_fundacion_pie(
                        doc, rebar_work, host, foundations, geom_opts, res,
                        solids_cache=fund_solids,
                    )
                else:
                    rebar_work = _procesar_rebar_vertical_pie_colision_muro_sin_fundacion(
                        doc,
                        rebar_work,
                        host,
                        walls,
                        geom_opts,
                        res,
                        params_dict,
                        layer_active_dict,
                        muro_contencion,
                        ids_con_apilado_bajo=ids_bajo,
                    )
                try:
                    eid_list[idx] = rebar_work.Id
                except Exception:
                    pass
                if cabezal_por_muro_id and rebar_work is not None:
                    try:
                        import armado_muros_cabezal as _cab_malla
                        import armado_muros_lineales as _lin_malla

                        ex_ini, ex_fin = _cab_malla.cabezal_extremos_config_for_muro(
                            cabezal_por_muro_id, int(wid),
                        )
                        if _lin_malla._rebar_es_malla_vertical_para_exclusion(
                            rebar_work,
                            host,
                            params_dict,
                            muro_contencion,
                        ):
                            _cab_malla.aplicar_exclusion_verticales_malla_rebar(
                                rebar_work,
                                ex_ini,
                                ex_fin,
                                doc=doc,
                                host=host,
                                regenerate=False,
                            )
                    except Exception:
                        pass
                try:
                    import armado_muros_lineales as _lin_malla
                    if _lin_malla._rebar_es_malla_vertical_para_stamp(
                        rebar_work,
                        host,
                        params_dict,
                        muro_contencion,
                    ):
                        from armado_muros_rebar_params import stamp_malla_vertical_rebar
                        stamp_malla_vertical_rebar(rebar_work)
                except Exception:
                    pass

    try:
        from armado_muros_txn import run_in_transaction

        run_in_transaction(
            doc,
            u"Arainco: Armado muros lineales — post verticales lote",
            _run_vertical_lote,
        )
    except Exception:
        _run_vertical_lote()

    return res


def aplicar_empotramiento_verticales_interior_por_colision(
    doc,
    walls,
    rebars_por_muro_id,
    concrete_grade=None,
    evaluar_colision_cabeza=True,
    params_por_muro_id=None,
    muro_contencion=False,
):
    """Alias retrocompatible; procesa exterior e interior."""
    return aplicar_empotramiento_verticales_cara_por_colision(
        doc,
        walls,
        rebars_por_muro_id,
        concrete_grade,
        evaluar_colision_cabeza=evaluar_colision_cabeza,
        params_por_muro_id=params_por_muro_id,
        muro_contencion=muro_contencion,
    )
