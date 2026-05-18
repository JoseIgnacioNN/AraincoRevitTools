# -*- coding: utf-8 -*-
"""
Motor de troceo: planos de corte, detección de colisiones de empotramiento
y cálculo de tramos para cada línea de fierro fusionada.

REGLA DE CAPA:
- No crea elementos Revit ni abre transacciones.
- Puede leer geometría de sólidos (operaciones booleanas de Revit) pero
  NO escribe en el modelo.
- No importa módulos de ui/ ni creators/.

Portado desde column_reinforcement_layout_rps.py.
"""
from __future__ import print_function

import clr
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BooleanOperationsType,
    BooleanOperationsUtils,
    BuiltInCategory,
    GeometryInstance,
    JoinGeometryUtils,
    Options,
    Plane,
    Solid,
    UnitTypeId,
    UnitUtils,
    ViewDetailLevel,
    XYZ,
)

from core.geometry import (
    _element_id_iv,
    _iter_solids_revit_element,
    _iter_vertices_from_solid,
    _solid_aggregate_vertex_ranges_ft,
    _geometry_options_structure_solids,
    _FOUNDATION_JOIN_FACE_Z_TOLERANCE_MM,
    _FOUNDATION_JOIN_OVERLAP_XY_MM,
    _FOUNDATION_STRETCH_DEDUCTION_MM,
    _CAT_STRUCT_FOUNDATION_IV,
    _EMBED_PROBE_XY_MARGIN_MM,
    _EMBED_PROBE_MIN_HALF_SIDE_MM,
    _TOL_VOL_INTERSECCION_EMBED_FT3,
    _REVOKE_EMBED_EXTRA_SHRINK_MM,
    get_column_dimensions,
)
from core.tables import hook_length_mm_from_nominal_diameter_mm


# ---------------------------------------------------------------------------
# Planos de corte A (troceo) — un plano horizontal cada joint entre pilares
# ---------------------------------------------------------------------------

def build_column_cut_planes_from_elements(columns_ordered, dims_cache):
    """
    Para cada transición entre dos columnas apiladas calcula un plano horizontal
    (origen = Z techo del tramo inferior) que actúa como plano de troceo tipo A.

    ``dims_cache`` es ``{element_id_iv: (width, depth, height, center_xyz, vs, vl)}``.

    Devuelve lista de ``Plane`` (puede estar vacía si solo hay una columna).
    """
    if not columns_ordered or len(columns_ordered) < 2:
        return []

    planes = []
    for i in range(len(columns_ordered) - 1):
        col_low  = columns_ordered[i]
        col_high = columns_ordered[i + 1]
        iv_low   = _element_id_iv(col_low)
        iv_high  = _element_id_iv(col_high)
        if iv_low < 0 or iv_high < 0:
            continue
        dims_low  = dims_cache.get(iv_low)
        dims_high = dims_cache.get(iv_high)
        if dims_low is None or dims_high is None:
            continue
        try:
            z_top_low   = float(dims_low[3].Z)  + float(dims_low[2])
            z_bot_high  = float(dims_high[3].Z)
        except Exception:
            continue
        # El plano se coloca en la media del gap (generalmente 0) o en z_top_low
        z_cut = 0.5 * (z_top_low + z_bot_high)
        try:
            pl = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0.0, 0.0, z_cut))
            planes.append(pl)
        except Exception:
            pass
    return planes


# ---------------------------------------------------------------------------
# Colisión de empotramiento (extremo +Z / −Z del tramo fusionado)
# ---------------------------------------------------------------------------

def _probe_solid_xyz(pt, half_side, probe_dz):
    """
    Devuelve un ``Solid`` cúbico de prueba para detección de colisión
    de empotramiento; ``None`` si no se puede crear.
    """
    from Autodesk.Revit.DB import (
        CurveLoop,
        Extrusion,
        Line,
        SolidOptions,
        GeometryCreationUtilities,
    )
    hs = abs(float(half_side))
    if hs < 1e-12:
        return None
    try:
        loop = CurveLoop()
        pts = [
            XYZ(pt.X - hs, pt.Y - hs, pt.Z),
            XYZ(pt.X + hs, pt.Y - hs, pt.Z),
            XYZ(pt.X + hs, pt.Y + hs, pt.Z),
            XYZ(pt.X - hs, pt.Y + hs, pt.Z),
        ]
        for j in range(4):
            loop.Append(
                Line.CreateBound(pts[j], pts[(j + 1) % 4])
            )
        loops = [loop]
        direction = XYZ(0.0, 0.0, 1.0) if probe_dz > 0 else XYZ(0.0, 0.0, -1.0)
        distance  = abs(float(probe_dz))
        solid = GeometryCreationUtilities.CreateExtrusionGeometry(
            loops, direction, distance
        )
        return solid
    except Exception:
        return None


def _embed_probe_intersects_column_solids(probe, column_solids):
    """True si la intersección booleana del probe con algún sólido de columna es > 0."""
    if probe is None:
        return False
    for solid in column_solids:
        if solid is None:
            continue
        try:
            inter = BooleanOperationsUtils.ExecuteBooleanOperation(
                probe, solid, BooleanOperationsType.Intersect
            )
            if inter is not None and float(inter.Volume) > _TOL_VOL_INTERSECCION_EMBED_FT3:
                return True
        except Exception:
            pass
    return False


def embed_stretch_collides_any_column_solids(
    pt_base,
    stretch_z_ft,
    nominal_diam_mm,
    column_solids,
):
    """
    Verifica si el tramo de empotramiento (desde pt_base ± stretch_z_ft)
    colisiona con sólidos de otras columnas. Devuelve True/False.
    """
    try:
        margin_ft = UnitUtils.ConvertToInternalUnits(
            float(_EMBED_PROBE_XY_MARGIN_MM), UnitTypeId.Millimeters
        )
        half_d_ft = UnitUtils.ConvertToInternalUnits(
            float(nominal_diam_mm) / 2.0, UnitTypeId.Millimeters
        )
        min_half_ft = UnitUtils.ConvertToInternalUnits(
            float(_EMBED_PROBE_MIN_HALF_SIDE_MM), UnitTypeId.Millimeters
        )
    except Exception:
        margin_ft  = float(_EMBED_PROBE_XY_MARGIN_MM) / 304.8
        half_d_ft  = float(nominal_diam_mm) / 2.0 / 304.8
        min_half_ft = float(_EMBED_PROBE_MIN_HALF_SIDE_MM) / 304.8

    half_side = max(float(half_d_ft) + float(margin_ft), float(min_half_ft))
    probe = _probe_solid_xyz(pt_base, half_side, float(stretch_z_ft))
    return _embed_probe_intersects_column_solids(probe, column_solids)


# ---------------------------------------------------------------------------
# Fundaciones unidas bajo la columna
# ---------------------------------------------------------------------------

def _vertex_ranges_xy_overlap_padded(rng_a, rng_b, pad_ft):
    if rng_a is None or rng_b is None:
        return False
    p = abs(float(pad_ft))
    try:
        if rng_a[0] is None or rng_b[0] is None:
            return False
        min_ax, max_ax = float(rng_a[0]), float(rng_a[1])
        min_ay, max_ay = float(rng_a[2]), float(rng_a[3])
        min_bx, max_bx = float(rng_b[0]), float(rng_b[1])
        min_by, max_by = float(rng_b[2]), float(rng_b[3])
        if max_ax + p < min_bx - p: return False
        if max_bx + p < min_ax - p: return False
        if max_ay + p < min_by - p: return False
        if max_by + p < min_ay - p: return False
    except Exception:
        return False
    return True


def _coerce_icollection_join_to_element_ids(raw):
    if raw is None:
        return []
    out = []
    try:
        for jid in raw:
            if jid is not None:
                out.append(jid)
        if out:
            return out
    except Exception:
        pass
    try:
        n = int(raw.Count)
    except Exception:
        n = 0
    for i in range(n):
        jid = None
        try:
            jid = raw[i]
        except Exception:
            try:
                jid = raw.get_Item(i)
            except Exception:
                jid = None
        if jid is not None:
            out.append(jid)
    return out


def _joined_elem_ids_revit(doc, host):
    if doc is None or host is None:
        return []
    raw = None
    for getter in (
        lambda: JoinGeometryUtils.GetJoinedElements(doc, host),
        lambda: JoinGeometryUtils.GetJoinedElements(doc, host.Id),
    ):
        try:
            raw = getter()
        except Exception:
            raw = None
        if raw is not None:
            break
    return _coerce_icollection_join_to_element_ids(raw)


def _elem_is_structural_foundation(elem):
    if elem is None:
        return False
    try:
        cat = elem.Category
        if cat is None:
            return False
        return int(cat.Id.IntegerValue) == _CAT_STRUCT_FOUNDATION_IV
    except Exception:
        return False


def column_bottom_joined_foundation_stretch_down_mm(doc, column):
    """
    Retorna mm a alargar hacia −Z (``altura_geom − 50``). 0 si no aplica.

    Busca fundaciones estructurales unidas cuya cara superior esté alineada
    con la base del pilar y tenga solape XY suficiente.
    """
    if doc is None or column is None:
        return 0.0
    rng_c = _solid_aggregate_vertex_ranges_ft(column)
    if rng_c is None:
        return 0.0
    try:
        z_col_min = float(rng_c[4])
    except Exception:
        return 0.0
    try:
        tol_z_ft  = UnitUtils.ConvertToInternalUnits(
            float(_FOUNDATION_JOIN_FACE_Z_TOLERANCE_MM), UnitTypeId.Millimeters
        )
        pad_xy_ft = UnitUtils.ConvertToInternalUnits(
            float(_FOUNDATION_JOIN_OVERLAP_XY_MM), UnitTypeId.Millimeters
        )
    except Exception:
        tol_z_ft  = float(_FOUNDATION_JOIN_FACE_Z_TOLERANCE_MM) / 304.8
        pad_xy_ft = float(_FOUNDATION_JOIN_OVERLAP_XY_MM) / 304.8

    for jid in _joined_elem_ids_revit(doc, column):
        try:
            jelem = doc.GetElement(jid)
        except Exception:
            continue
        if not _elem_is_structural_foundation(jelem):
            continue
        rng_f = _solid_aggregate_vertex_ranges_ft(jelem)
        if rng_f is None:
            continue
        try:
            z_fond_max = float(rng_f[5])
        except Exception:
            continue
        if abs(z_fond_max - z_col_min) > abs(float(tol_z_ft)):
            continue
        if not _vertex_ranges_xy_overlap_padded(
            (rng_c[0], rng_c[1], rng_c[2], rng_c[3]),
            (rng_f[0], rng_f[1], rng_f[2], rng_f[3]),
            pad_xy_ft,
        ):
            continue
        try:
            h_fond_ft = float(rng_f[5]) - float(rng_f[4])
            h_fond_mm = UnitUtils.ConvertFromInternalUnits(h_fond_ft, UnitTypeId.Millimeters)
        except Exception:
            try:
                h_fond_mm = (float(rng_f[5]) - float(rng_f[4])) * 304.8
            except Exception:
                continue
        stretch_mm = h_fond_mm - float(_FOUNDATION_STRETCH_DEDUCTION_MM)
        if stretch_mm > 0.0:
            return float(stretch_mm)
    return 0.0


# ---------------------------------------------------------------------------
# Cálculo de empotramiento por extremo fusionado
# ---------------------------------------------------------------------------

def calc_embed_and_revoke(
    doc,
    pt_bar,
    z_fused_start_ft,
    z_fused_end_ft,
    nominal_diam_mm,
    concrete_grade,
    column_solids_list,
    column,
    short_curve_tolerance_ft,
):
    """
    Calcula:
    - ``embed_top_ft``: estiramiento +Z por empotramiento (0 si colisiona).
    - ``embed_bot_ft``: estiramiento −Z (fundación o empotramiento).
    - ``revoke_bot_ft``: recorte −Z si no hay colisión que justifique el estiramiento.

    Devuelve ``(embed_top_ft, embed_bot_ft, revoke_bot_ft)``.
    """
    from core.tables import (
        hook_length_mm_from_nominal_diameter_mm,
        traslape_mm_from_nominal_diameter_mm,
    )
    tol = abs(float(short_curve_tolerance_ft))

    try:
        hook_mm = hook_length_mm_from_nominal_diameter_mm(nominal_diam_mm, concrete_grade)
        hook_ft = UnitUtils.ConvertToInternalUnits(float(hook_mm), UnitTypeId.Millimeters)
    except Exception:
        hook_ft = 0.0

    # Empotramiento superior (+Z)
    try:
        pt_top = XYZ(float(pt_bar.X), float(pt_bar.Y), float(z_fused_end_ft))
        if embed_stretch_collides_any_column_solids(
            pt_top, hook_ft, nominal_diam_mm, column_solids_list
        ):
            embed_top_ft = 0.0
        else:
            embed_top_ft = hook_ft
    except Exception:
        embed_top_ft = 0.0

    # Empotramiento inferior (−Z): primero fundaciones unidas
    try:
        found_mm = column_bottom_joined_foundation_stretch_down_mm(doc, column)
    except Exception:
        found_mm = 0.0

    if found_mm > 0.0:
        try:
            embed_bot_ft = UnitUtils.ConvertToInternalUnits(found_mm, UnitTypeId.Millimeters)
        except Exception:
            embed_bot_ft = found_mm / 304.8
        revoke_bot_ft = 0.0
    else:
        # Sin fundación: empotramiento por tabla o recorte
        try:
            pt_bot = XYZ(float(pt_bar.X), float(pt_bar.Y), float(z_fused_start_ft))
            if embed_stretch_collides_any_column_solids(
                pt_bot, -hook_ft, nominal_diam_mm, column_solids_list
            ):
                embed_bot_ft = hook_ft
            else:
                embed_bot_ft  = 0.0
                try:
                    extra_mm = (
                        float(_REVOKE_EMBED_EXTRA_SHRINK_MM)
                        + float(nominal_diam_mm) / 2.0
                    )
                    revoke_bot_ft = UnitUtils.ConvertToInternalUnits(extra_mm, UnitTypeId.Millimeters)
                except Exception:
                    revoke_bot_ft = 0.0
        except Exception:
            embed_bot_ft  = 0.0
            revoke_bot_ft = 0.0

    return embed_top_ft, embed_bot_ft, revoke_bot_ft
